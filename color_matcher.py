"""
DyeForge - Game color matching & auto-adjustment tool

Built for Blue Protocol: Star Resonance (BPSR) costume dyeing.
Captures colors from any window, compares with target, and automatically
inputs HEX codes into the game's color picker with iterative correction.

Features:
  - Window capture & color pick (target + real-time monitoring)
  - Similarity score with HSV-based visual hints
  - Auto-adjust via HEX code input with adaptive damping correction loop
  - Game color range awareness (RGB 0x1A-0xDA, HSV saturation limits)
"""

import ctypes
import ctypes.wintypes
import colorsys
import logging
import math
import os
import statistics
import sys
import threading
import time
import tkinter as tk
from tkinter import ttk
from PIL import Image, ImageTk

# Log to file (works even with --windowed EXE where stdout is closed)
logging.basicConfig(
    filename=os.path.join(os.path.expanduser("~"), "dyeforge.log"),
    level=logging.DEBUG,
    format="%(asctime)s %(message)s",
    datefmt="%H:%M:%S",
)

try:
    import numpy as np
    HAS_NUMPY = True
except ImportError:
    HAS_NUMPY = False


def _resource_path(rel):
    """Locate a bundled/adjacent data file.

    Resolution order:
      1. PyInstaller --onefile extract dir (sys._MEIPASS) if frozen
      2. Directory of the exe (for files dropped next to DyeForge.exe)
      3. Directory of this .py (dev environment)
    Returns the first existing path, or None if none found.
    """
    candidates = []
    if getattr(sys, "frozen", False):
        if hasattr(sys, "_MEIPASS"):
            candidates.append(os.path.join(sys._MEIPASS, rel))
        candidates.append(os.path.join(os.path.dirname(sys.executable), rel))
    candidates.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), rel))
    for p in candidates:
        if os.path.exists(p):
            return p
    return None

# ── Win32 API helpers ──────────────────────────────────────────────

user32 = ctypes.windll.user32
gdi32 = ctypes.windll.gdi32
dwmapi = ctypes.windll.dwmapi

SRCCOPY = 0x00CC0020
CAPTUREBLT = 0x40000000
PW_RENDERFULLCONTENT = 0x00000002
PW_CLIENTONLY = 0x00000001
DIB_RGB_COLORS = 0
BI_RGB = 0
DWMWA_EXTENDED_FRAME_BOUNDS = 9
SW_HIDE = 0
SW_SHOW = 5
SW_MINIMIZE = 6
SW_RESTORE = 9

# Mouse input
INPUT_MOUSE = 0
MOUSEEVENTF_LEFTDOWN = 0x0002
MOUSEEVENTF_LEFTUP = 0x0004

# Keyboard input
INPUT_KEYBOARD = 1
KEYEVENTF_KEYUP = 0x0002
KEYEVENTF_UNICODE = 0x0004
VK_CONTROL = 0x11
VK_RETURN = 0x0D
VK_A = 0x41
VK_BACK = 0x08


class MOUSEINPUT(ctypes.Structure):
    _fields_ = [
        ("dx", ctypes.wintypes.LONG),
        ("dy", ctypes.wintypes.LONG),
        ("mouseData", ctypes.wintypes.DWORD),
        ("dwFlags", ctypes.wintypes.DWORD),
        ("time", ctypes.wintypes.DWORD),
        ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong)),
    ]


class KEYBDINPUT(ctypes.Structure):
    _fields_ = [
        ("wVk", ctypes.wintypes.WORD),
        ("wScan", ctypes.wintypes.WORD),
        ("dwFlags", ctypes.wintypes.DWORD),
        ("time", ctypes.wintypes.DWORD),
        ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong)),
    ]


class INPUT(ctypes.Structure):
    class _UNION(ctypes.Union):
        _fields_ = [("mi", MOUSEINPUT), ("ki", KEYBDINPUT)]

    _fields_ = [
        ("type", ctypes.wintypes.DWORD),
        ("union", _UNION),
    ]


class BITMAPINFOHEADER(ctypes.Structure):
    _fields_ = [
        ("biSize", ctypes.wintypes.DWORD),
        ("biWidth", ctypes.wintypes.LONG),
        ("biHeight", ctypes.wintypes.LONG),
        ("biPlanes", ctypes.wintypes.WORD),
        ("biBitCount", ctypes.wintypes.WORD),
        ("biCompression", ctypes.wintypes.DWORD),
        ("biSizeImage", ctypes.wintypes.DWORD),
        ("biXPelsPerMeter", ctypes.wintypes.LONG),
        ("biYPelsPerMeter", ctypes.wintypes.LONG),
        ("biClrUsed", ctypes.wintypes.DWORD),
        ("biClrImportant", ctypes.wintypes.DWORD),
    ]


class BITMAPINFO(ctypes.Structure):
    _fields_ = [
        ("bmiHeader", BITMAPINFOHEADER),
        ("bmiColors", ctypes.wintypes.DWORD * 3),
    ]


def click_at(screen_x, screen_y, delay=0.05):
    """Move mouse to (screen_x, screen_y) and left-click.

    Uses multiple methods for compatibility with games:
    1. SetCursorPos to move the mouse
    2. SendInput for the click (works with most apps)
    3. Also calls mouse_event as fallback (works with some games that ignore SendInput)
    """
    sx, sy = int(screen_x), int(screen_y)
    user32.SetCursorPos(sx, sy)
    time.sleep(delay)

    # Method 1: SendInput
    inp_down = INPUT()
    inp_down.type = INPUT_MOUSE
    inp_down.union.mi.dwFlags = MOUSEEVENTF_LEFTDOWN
    user32.SendInput(1, ctypes.byref(inp_down), ctypes.sizeof(INPUT))

    time.sleep(0.03)

    inp_up = INPUT()
    inp_up.type = INPUT_MOUSE
    inp_up.union.mi.dwFlags = MOUSEEVENTF_LEFTUP
    user32.SendInput(1, ctypes.byref(inp_up), ctypes.sizeof(INPUT))

    time.sleep(0.03)

    # Method 2: mouse_event fallback (some games only accept this)
    user32.mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    time.sleep(0.03)
    user32.mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)


def _send_key(vk, up=False):
    """Send a single virtual key press or release."""
    inp = INPUT()
    inp.type = INPUT_KEYBOARD
    inp.union.ki.wVk = vk
    inp.union.ki.dwFlags = KEYEVENTF_KEYUP if up else 0
    user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(INPUT))


def _send_unicode_char(char):
    """Send a single Unicode character via SendInput."""
    down = INPUT()
    down.type = INPUT_KEYBOARD
    down.union.ki.wVk = 0
    down.union.ki.wScan = ord(char)
    down.union.ki.dwFlags = KEYEVENTF_UNICODE
    user32.SendInput(1, ctypes.byref(down), ctypes.sizeof(INPUT))
    time.sleep(0.01)
    up = INPUT()
    up.type = INPUT_KEYBOARD
    up.union.ki.wVk = 0
    up.union.ki.wScan = ord(char)
    up.union.ki.dwFlags = KEYEVENTF_UNICODE | KEYEVENTF_KEYUP
    user32.SendInput(1, ctypes.byref(up), ctypes.sizeof(INPUT))


def type_hex_into_field(screen_x, screen_y, hex_code):
    """Click a HEX input field, clear it, and type a 6-char hex code.

    hex_code: e.g. "FF8040" (without #)
    """
    # Click the input field
    click_at(screen_x, screen_y, delay=0.1)
    time.sleep(0.1)

    # Triple-click to select all text in the field
    for _ in range(2):
        click_at(screen_x, screen_y, delay=0.02)
    time.sleep(0.05)

    # Ctrl+A as backup select-all
    _send_key(VK_CONTROL)
    time.sleep(0.02)
    _send_key(VK_A)
    time.sleep(0.02)
    _send_key(VK_A, up=True)
    _send_key(VK_CONTROL, up=True)
    time.sleep(0.05)

    # Type the hex code character by character
    for ch in hex_code.upper():
        _send_unicode_char(ch)
        time.sleep(0.02)

    time.sleep(0.05)

    # Press Enter to confirm
    _send_key(VK_RETURN)
    time.sleep(0.02)
    _send_key(VK_RETURN, up=True)


def get_window_list():
    """Return list of visible windows as [{"id": hwnd, "title": str}, ...]."""
    results = []

    def _cb(hwnd, _):
        if not user32.IsWindowVisible(hwnd):
            return True
        length = user32.GetWindowTextLengthW(hwnd)
        if length == 0:
            return True
        buf = ctypes.create_unicode_buffer(length + 1)
        user32.GetWindowTextW(hwnd, buf, length + 1)
        title = buf.value
        if title:
            results.append({"id": hwnd, "title": title})
        return True

    WNDENUMPROC = ctypes.WINFUNCTYPE(
        ctypes.wintypes.BOOL, ctypes.wintypes.HWND, ctypes.wintypes.LPARAM
    )
    user32.EnumWindows(WNDENUMPROC(_cb), 0)
    return results


def find_game_window():
    """Find BPSR game window by searching known titles. Returns hwnd or None."""
    for w in get_window_list():
        for name in GAME_WINDOW_NAMES:
            if name in w["title"]:
                return w["id"]
    return None


def capture_window_printwindow(hwnd):
    """Capture using PrintWindow (works even if window is occluded)."""
    rect = ctypes.wintypes.RECT()
    user32.GetClientRect(hwnd, ctypes.byref(rect))
    w = rect.right - rect.left
    h = rect.bottom - rect.top
    if w <= 0 or h <= 0:
        user32.GetWindowRect(hwnd, ctypes.byref(rect))
        w = rect.right - rect.left
        h = rect.bottom - rect.top
        if w <= 0 or h <= 0:
            return None

    hwnd_dc = user32.GetWindowDC(hwnd)
    mem_dc = gdi32.CreateCompatibleDC(hwnd_dc)
    bitmap = gdi32.CreateCompatibleBitmap(hwnd_dc, w, h)
    old = gdi32.SelectObject(mem_dc, bitmap)

    result = user32.PrintWindow(hwnd, mem_dc, PW_RENDERFULLCONTENT)

    bmi = BITMAPINFO()
    bmi.bmiHeader.biSize = ctypes.sizeof(BITMAPINFOHEADER)
    bmi.bmiHeader.biWidth = w
    bmi.bmiHeader.biHeight = -h
    bmi.bmiHeader.biPlanes = 1
    bmi.bmiHeader.biBitCount = 32
    bmi.bmiHeader.biCompression = BI_RGB

    buf = ctypes.create_string_buffer(w * h * 4)
    gdi32.GetDIBits(mem_dc, bitmap, 0, h, buf, ctypes.byref(bmi), DIB_RGB_COLORS)

    gdi32.SelectObject(mem_dc, old)
    gdi32.DeleteObject(bitmap)
    gdi32.DeleteDC(mem_dc)
    user32.ReleaseDC(hwnd, hwnd_dc)

    if result == 0:
        return None

    img = Image.frombuffer("RGBX", (w, h), buf, "raw", "BGRX", 0, 1).convert("RGB")

    if HAS_NUMPY:
        arr = np.array(img)
        if arr.max() < 5:
            return None
    else:
        extrema = img.getextrema()
        if all(ch[1] < 5 for ch in extrema):
            return None

    return img


def capture_screen_region(hwnd):
    """Capture from the screen (BitBlt). Requires the window to be visible."""
    rect = ctypes.wintypes.RECT()
    res = dwmapi.DwmGetWindowAttribute(
        hwnd, DWMWA_EXTENDED_FRAME_BOUNDS, ctypes.byref(rect), ctypes.sizeof(rect)
    )
    if res != 0:
        user32.GetWindowRect(hwnd, ctypes.byref(rect))

    w = rect.right - rect.left
    h = rect.bottom - rect.top
    if w <= 0 or h <= 0:
        return None

    screen_dc = user32.GetDC(0)
    mem_dc = gdi32.CreateCompatibleDC(screen_dc)
    bitmap = gdi32.CreateCompatibleBitmap(screen_dc, w, h)
    old = gdi32.SelectObject(mem_dc, bitmap)
    gdi32.BitBlt(mem_dc, 0, 0, w, h, screen_dc, rect.left, rect.top,
                 SRCCOPY | CAPTUREBLT)

    bmi = BITMAPINFO()
    bmi.bmiHeader.biSize = ctypes.sizeof(BITMAPINFOHEADER)
    bmi.bmiHeader.biWidth = w
    bmi.bmiHeader.biHeight = -h
    bmi.bmiHeader.biPlanes = 1
    bmi.bmiHeader.biBitCount = 32
    bmi.bmiHeader.biCompression = BI_RGB

    buf = ctypes.create_string_buffer(w * h * 4)
    gdi32.GetDIBits(mem_dc, bitmap, 0, h, buf, ctypes.byref(bmi), DIB_RGB_COLORS)

    gdi32.SelectObject(mem_dc, old)
    gdi32.DeleteObject(bitmap)
    gdi32.DeleteDC(mem_dc)
    user32.ReleaseDC(0, screen_dc)

    img = Image.frombuffer("RGBX", (w, h), buf, "raw", "BGRX", 0, 1).convert("RGB")
    return img


def capture_window(hwnd, hide_callback=None, show_callback=None):
    """Capture a window. Tries PrintWindow first, falls back to BitBlt."""
    img = capture_window_printwindow(hwnd)
    if img is not None:
        return img

    if hide_callback:
        hide_callback()
        time.sleep(0.15)
    user32.SetForegroundWindow(hwnd)
    time.sleep(0.4)
    img = capture_screen_region(hwnd)
    if show_callback:
        show_callback()
    return img


# ── Color utilities (matching the original JS logic) ──────────────

def rgb_to_hsv(r, g, b):
    """Convert RGB (0-255) to HSV: h in [0,360], s/v in [0,100]."""
    r2, g2, b2 = r / 255, g / 255, b / 255
    h, s, v = colorsys.rgb_to_hsv(r2, g2, b2)
    return {"h": h * 360, "s": s * 100, "v": v * 100}


def similarity_score(target, current):
    """Compute color similarity (0-100%) using Euclidean distance in RGB space."""
    if target is None or current is None:
        return 0.0
    dr = target[0] - current[0]
    dg = target[1] - current[1]
    db = target[2] - current[2]
    dist = math.sqrt(dr * dr + dg * dg + db * db)
    max_dist = math.sqrt(255**2 * 3)
    return (1 - dist / max_dist) * 100


def _srgb_to_linear(c):
    """sRGB channel (0-1) to linear RGB."""
    if c <= 0.04045:
        return c / 12.92
    return ((c + 0.055) / 1.055) ** 2.4


def rgb_to_lab(r, g, b):
    """sRGB (0-255) → CIE Lab (D65). Returns (L, a, b)."""
    # sRGB → linear RGB
    rl = _srgb_to_linear(r / 255)
    gl = _srgb_to_linear(g / 255)
    bl = _srgb_to_linear(b / 255)
    # Linear RGB → XYZ (D65)
    x = rl * 0.4124564 + gl * 0.3575761 + bl * 0.1804375
    y = rl * 0.2126729 + gl * 0.7151522 + bl * 0.0721750
    z = rl * 0.0193339 + gl * 0.1191920 + bl * 0.9503041
    # Normalize to D65 white
    xn, yn, zn = x / 0.95047, y / 1.00000, z / 1.08883

    def f(t):
        delta = 6 / 29
        if t > delta ** 3:
            return t ** (1 / 3)
        return t / (3 * delta * delta) + 4 / 29

    fx, fy, fz = f(xn), f(yn), f(zn)
    L = 116 * fy - 16
    a = 500 * (fx - fy)
    b_lab = 200 * (fy - fz)
    return (L, a, b_lab)


def lab_to_rgb_approx(lab):
    """CIE Lab (D65) → sRGB (0-255 int). Approximate inverse of rgb_to_lab.
    Used by the LUT refinement loop to convert Lab-space residuals back
    to RGB targets for find_best_input.
    """
    L, a, b_lab = lab
    fy = (L + 16) / 116
    fx = a / 500 + fy
    fz = fy - b_lab / 200

    def finv(t):
        delta = 6 / 29
        if t > delta:
            return t ** 3
        return 3 * delta * delta * (t - 4 / 29)

    xn, yn, zn = finv(fx) * 0.95047, finv(fy) * 1.00000, finv(fz) * 1.08883
    rl = 3.2404542 * xn - 1.5371385 * yn - 0.4985314 * zn
    gl = -0.9692660 * xn + 1.8760108 * yn + 0.0415560 * zn
    bl = 0.0556434 * xn - 0.2040259 * yn + 1.0572252 * zn

    def gamma_encode(c):
        c = max(0.0, c)
        if c <= 0.0031308:
            return 12.92 * c
        return 1.055 * (c ** (1 / 2.4)) - 0.055

    R = max(0, min(255, round(gamma_encode(rl) * 255)))
    G = max(0, min(255, round(gamma_encode(gl) * 255)))
    B = max(0, min(255, round(gamma_encode(bl) * 255)))
    return (R, G, B)


def delta_e(rgb1, rgb2):
    """CIE76 Delta E — perceptual color distance.
    < 1: indistinguishable, < 2: trained eye only, < 10: similar, > 20: different.
    """
    if rgb1 is None or rgb2 is None:
        return 999.0
    L1, a1, b1 = rgb_to_lab(*rgb1)
    L2, a2, b2 = rgb_to_lab(*rgb2)
    return math.sqrt((L1 - L2) ** 2 + (a1 - a2) ** 2 + (b1 - b2) ** 2)


def delta_e_category(de):
    """Human-readable category for Delta E value."""
    if de < 1:
        return "同一"
    if de < 2:
        return "ほぼ同一"
    if de < 5:
        return "近い"
    if de < 10:
        return "やや違う"
    if de < 20:
        return "違う"
    return "全く違う"


def generate_hint(target, current):
    """Generate Japanese hint text (matching xm() in the original JS)."""
    if target is None or current is None:
        return "色を取得するとここにリアルタイムのアドバイスが表示されます"
    tr, tg, tb = target
    cr, cg, cb = current
    if tr == cr and tg == cg and tb == cb:
        return "完全に一致しました！(100%)"

    t_hsv = rgb_to_hsv(tr, tg, tb)
    c_hsv = rgb_to_hsv(cr, cg, cb)

    hints = []
    dh = t_hsv["h"] - c_hsv["h"]
    if abs(dh) > 0.1:
        if (0 < dh <= 180) or dh < -180:
            hints.append("スライダーを上へ")
        else:
            hints.append("スライダーを下へ")

    ds = t_hsv["s"] - c_hsv["s"]
    dv = t_hsv["v"] - c_hsv["v"]
    lr = "右" if ds > 0.1 else ("左" if ds < -0.1 else "")
    ud = "上" if dv > 0.1 else ("下" if dv < -0.1 else "")
    if lr or ud:
        hints.append(f"ポインタを{ud}{lr}へ")

    if not hints:
        return "あとごく僅かです..."
    return "、".join(hints) + "動かしてください。"


# ── Colours / theme ───────────────────────────────────────────────

BG = "#09090b"          # zinc-950
CARD_BG = "#18181b"     # zinc-900
CARD_BORDER = "#27272a" # zinc-800
TEXT_PRIMARY = "#f4f4f5" # zinc-100
TEXT_DIM = "#71717a"     # zinc-500
INDIGO = "#4f46e5"       # indigo-600
INDIGO_LIGHT = "#6366f1" # indigo-500
INDIGO_DIM = "#4338ca"   # indigo-700
HOVER_BG = "#3f3f46"    # zinc-700
GREEN = "#22c55e"        # green-500

# Supported game window titles (BPSR: Global, 繁體中文, 日本語, 韓国語)
GAME_WINDOW_NAMES = ['Blue Protocol', '星痕共鳴', 'ブループロトコル', '블루 프로토콜']


# ── Main Application ──────────────────────────────────────────────

class ColorMatcherApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("DyeForge")
        self.root.configure(bg=BG)
        self.root.geometry("850x660")
        self.root.minsize(750, 600)

        # State
        self.target_color = None
        self.current_color = None
        self.pick_mode = None
        self.captured_image = None
        self.hover_color = None
        self.window_list = []
        self.selected_hwnd = find_game_window()  # auto-detect BPSR
        self.cursor_img_pos = None

        # Polling
        self._watch_img_pos = None
        self._polling = False
        self._poll_after_id = None

        # Always-on-top
        self._topmost = False

        # Calibration (HEX input field position only)
        self.calibration = None
        self._calib_overlay = None
        self._cancel_adjust = False

        # Auto-adjust settings
        self._skip_click = tk.BooleanVar(value=True)
        self._eye_mode = tk.BooleanVar(value=False)
        self._threshold = tk.DoubleVar(value=1.0)  # Delta E threshold (lower = stricter)
        self._max_retries = tk.IntVar(value=30)
        self._use_lut = tk.BooleanVar(value=True)  # use 3D LUT for initial input (see lut_lookup.py)

        # 3D LUT (optional — loaded at init, None if lut_3d.json missing)
        self._lut = None
        self._lut_path = _resource_path("lut_3d.json")
        self._load_lut()

        self._load_calibration()
        self._build_ui()

    # ── Persistence & toggles ────────────────────────────────────

    def _load_calibration(self):
        self.calibration = None

    def _save_calibration(self):
        pass  # no file persistence — calibrate each session

    def _load_lut(self):
        """Load the 3D LUT from lut_3d.json if available.
        Silent failure — self._lut stays None and the auto-adjust falls back
        to the legacy iterative path.
        """
        if self._lut_path is None:
            logging.info("LUT未検出: lut_3d.json not found — fallback to iterative-only")
            return
        try:
            import lut_lookup  # lazy to avoid circular import at module scope
            self._lut = lut_lookup.load_lut(self._lut_path)
            logging.info(f"LUT読込完了: {len(self._lut)} points from {self._lut_path}")
        except Exception as e:
            logging.info(f"LUT読込失敗: {e} — fallback to iterative-only")
            self._lut = None

    def _toggle_topmost(self):
        self._topmost = not self._topmost
        self.root.attributes("-topmost", self._topmost)
        if self._topmost:
            self._topmost_btn.configure(bg=INDIGO, fg="white")
        else:
            self._topmost_btn.configure(bg=CARD_BORDER, fg=TEXT_DIM)

    # ── UI construction ───────────────────────────────────────────

    def _build_ui(self):
        # Header
        header = tk.Frame(self.root, bg=BG)
        header.pack(fill="x", padx=24, pady=(16, 8))

        tk.Label(header, text="DyeForge", fg=TEXT_PRIMARY, bg=BG,
                 font=("Segoe UI", 20, "bold")).pack(side="left")

        self._topmost_btn = tk.Button(
            header, text="pin", bg=CARD_BORDER, fg=TEXT_DIM,
            font=("Segoe UI", 9), relief="flat", cursor="hand2",
            activebackground=HOVER_BG, command=self._toggle_topmost
        )
        self._topmost_btn.pack(side="right")

        # Body: 2 columns
        body = tk.Frame(self.root, bg=BG)
        body.pack(fill="both", expand=True, padx=24, pady=(0, 16))

        # Left column (color panels)
        left = tk.Frame(body, bg=BG, width=280)
        left.pack(side="left", fill="y", padx=(0, 12))
        left.pack_propagate(False)

        self._build_color_panel(left, "Target", "target", is_target=True)
        self._build_color_panel(left, "Current", "watch", is_target=False)

        # Right column
        right = tk.Frame(body, bg=CARD_BG, highlightbackground=CARD_BORDER,
                         highlightthickness=1)
        right.pack(side="left", fill="both", expand=True)

        self._build_score_section(right)
        self._build_hsv_section(right)

        tk.Frame(right, bg=CARD_BORDER, height=1).pack(fill="x", padx=20, pady=(4, 0))

        self._build_auto_section(right)

    def _build_color_panel(self, parent, label, mode, is_target):
        frame = tk.Frame(parent, bg=CARD_BG, highlightbackground=CARD_BORDER,
                         highlightthickness=1)
        frame.pack(fill="both", expand=True, pady=(0, 8))

        # Swatch
        swatch = tk.Canvas(frame, bg="#18181b", highlightthickness=0, height=60)
        swatch.pack(fill="x", padx=12, pady=(12, 4))

        # Label
        tk.Label(frame, text=label.upper(), fg=TEXT_DIM, bg=CARD_BG,
                 font=("Segoe UI", 8, "bold")).pack()

        # RGB + HEX values
        value_label = tk.Label(frame, text="---", fg=TEXT_PRIMARY, bg=CARD_BG,
                               font=("Consolas", 10))
        value_label.pack()
        hex_label = tk.Label(frame, text="", fg=TEXT_DIM, bg=CARD_BG,
                             font=("Consolas", 9))
        hex_label.pack()

        # Button
        btn_bg = "#f4f4f5" if is_target else INDIGO
        btn_fg = "#18181b" if is_target else "white"
        btn_text = "色を取得" if is_target else "監視設定"
        btn = tk.Button(frame, text=btn_text, bg=btn_bg, fg=btn_fg,
                        font=("Segoe UI", 9, "bold"), relief="flat", cursor="hand2",
                        activebackground="#e4e4e7" if is_target else INDIGO_LIGHT,
                        command=lambda: self._open_window_picker(mode))
        btn.pack(fill="x", padx=12, pady=(4, 12), ipady=4)

        if is_target:
            self._target_swatch = swatch
            self._target_label = value_label
            self._target_hex = hex_label
        else:
            self._current_swatch = swatch
            self._current_label = value_label
            self._current_hex = hex_label

    def _build_score_section(self, parent):
        frame = tk.Frame(parent, bg=CARD_BG)
        frame.pack(fill="x", padx=20, pady=(16, 0))

        tk.Label(frame, text="SCORE", fg=TEXT_DIM, bg=CARD_BG,
                 font=("Segoe UI", 8, "bold")).pack()

        score_row = tk.Frame(frame, bg=CARD_BG)
        score_row.pack()

        self._score_label = tk.Label(score_row, text="0.0", fg=INDIGO_LIGHT, bg=CARD_BG,
                                     font=("Segoe UI", 56, "bold"))
        self._score_label.pack(side="left")
        tk.Label(score_row, text="%", fg=INDIGO_DIM, bg=CARD_BG,
                 font=("Segoe UI", 22, "bold")).pack(side="left", anchor="s", pady=(0, 10))

        # Progress bar
        bar_bg = tk.Frame(frame, bg="#27272a", height=6)
        bar_bg.pack(fill="x", padx=30, pady=(4, 0))
        bar_bg.pack_propagate(False)
        self._bar_fill = tk.Frame(bar_bg, bg=INDIGO_LIGHT, height=6)
        self._bar_fill.place(relx=0, rely=0, relheight=1, relwidth=0)

        # Delta E perceptual distance
        self._de_label = tk.Label(
            frame, text="", fg=TEXT_DIM, bg=CARD_BG,
            font=("Segoe UI", 10, "bold")
        )
        self._de_label.pack(pady=(4, 0))

        # Hint
        self._hint_label = tk.Label(
            frame, text="色を取得するとアドバイスが表示されます",
            fg=TEXT_PRIMARY, bg=CARD_BG, font=("Segoe UI", 10),
            wraplength=400, justify="center"
        )
        self._hint_label.pack(fill="x", pady=(8, 0))

    def _build_hsv_section(self, parent):
        """HSV color comparison: SV square + hue bar with target/current markers."""
        frame = tk.Frame(parent, bg=CARD_BG)
        frame.pack(fill="x", padx=20, pady=(6, 0))

        self._hsv_canvas = tk.Canvas(frame, bg=CARD_BG, height=130,
                                     highlightthickness=0)
        self._hsv_canvas.pack(fill="x")

    def _draw_hsv_comparison(self):
        canvas = self._hsv_canvas
        canvas.delete("all")

        if not self.target_color or not self.current_color:
            if self.target_color:
                # Show target only
                t_hsv = rgb_to_hsv(*self.target_color)
                self._draw_hsv_with_markers(canvas, t_hsv, None)
            return

        t_hsv = rgb_to_hsv(*self.target_color)
        c_hsv = rgb_to_hsv(*self.current_color)
        self._draw_hsv_with_markers(canvas, t_hsv, c_hsv)

    def _draw_hsv_with_markers(self, canvas, t_hsv, c_hsv):
        sq = 120  # SV square size
        sl_w = 16  # hue slider width
        gap = 8
        total_w = canvas.winfo_width()
        start_x = max(10, (total_w - sq - sl_w - gap) // 2)
        y0 = 5

        # Draw SV square
        h_deg = t_hsv["h"]
        base_r, base_g, base_b = [
            int(c * 255) for c in colorsys.hsv_to_rgb(h_deg / 360, 1, 1)
        ]

        # Game-reachable zone boundaries
        s_max_frac = ColorMatcherApp.GAME_S_MAX_PCT / 100
        v_min_frac = ColorMatcherApp.GAME_V_MIN_PCT / 100
        v_max_frac = ColorMatcherApp.GAME_V_MAX_PCT / 100

        # Semi-transparent filter blend: keeps original color visible
        # unreachable = original * 0.55 + dark_red_tint * 0.45
        TINT = (30, 10, 10)  # dark reddish tint
        TINT_ALPHA = 0.45

        if HAS_NUMPY:
            xs = np.linspace(0, 1, sq).reshape(1, sq)
            ys = np.linspace(1, 0, sq).reshape(sq, 1)
            r_ch = ((1 - xs) + xs * base_r / 255) * ys * 255
            g_ch = ((1 - xs) + xs * base_g / 255) * ys * 255
            b_ch = ((1 - xs) + xs * base_b / 255) * ys * 255
            arr = np.stack([r_ch, g_ch, b_ch], axis=-1).clip(0, 255).astype(np.uint8)

            s_mask = xs > s_max_frac
            v_low_mask = ys < v_min_frac
            v_high_mask = ys > v_max_frac
            unreachable_2d = s_mask | v_low_mask | v_high_mask
            unreachable_3d = np.broadcast_to(unreachable_2d[..., None], arr.shape)
            # Blend: arr*0.55 + tint*0.45
            tint_arr = np.array(TINT, dtype=np.float32)
            blended = (arr.astype(np.float32) * (1 - TINT_ALPHA) + tint_arr * TINT_ALPHA)
            arr = np.where(unreachable_3d, blended.astype(np.uint8), arr)
            sv_img = Image.fromarray(arr, "RGB")
        else:
            sv_img = Image.new("RGB", (sq, sq))
            for py in range(sq):
                for px in range(sq):
                    s = px / sq
                    v = 1 - py / sq
                    r = int(((1 - s) + s * base_r / 255) * v * 255)
                    g = int(((1 - s) + s * base_g / 255) * v * 255)
                    b = int(((1 - s) + s * base_b / 255) * v * 255)
                    if s > s_max_frac or v < v_min_frac or v > v_max_frac:
                        r = int(r * (1 - TINT_ALPHA) + TINT[0] * TINT_ALPHA)
                        g = int(g * (1 - TINT_ALPHA) + TINT[1] * TINT_ALPHA)
                        b = int(b * (1 - TINT_ALPHA) + TINT[2] * TINT_ALPHA)
                    sv_img.putpixel((px, py), (r, g, b))

        self._sv_photo = ImageTk.PhotoImage(sv_img)
        canvas.create_image(start_x, y0, anchor="nw", image=self._sv_photo)

        # Boundary lines (red dashed)
        # Vertical: S limit
        boundary_x = start_x + s_max_frac * sq
        canvas.create_line(boundary_x, y0, boundary_x, y0 + sq,
                           fill="#ef4444", dash=(3, 2), width=2)
        # Horizontal: V min (bottom)
        v_min_y = y0 + (1 - v_min_frac) * sq
        canvas.create_line(start_x, v_min_y, start_x + sq, v_min_y,
                           fill="#ef4444", dash=(3, 2), width=2)
        # Horizontal: V max (top)
        v_max_y = y0 + (1 - v_max_frac) * sq
        canvas.create_line(start_x, v_max_y, start_x + sq, v_max_y,
                           fill="#ef4444", dash=(3, 2), width=2)

        # "範囲外" label in the unreachable corner
        canvas.create_text(start_x + sq - 10, y0 + 10,
                           text="範囲外", fill="#ef4444",
                           font=("Segoe UI", 7, "bold"), anchor="ne")

        # Target marker (white circle with dot)
        tx = start_x + t_hsv["s"] / 100 * sq
        ty = y0 + (1 - t_hsv["v"] / 100) * sq
        canvas.create_oval(tx - 7, ty - 7, tx + 7, ty + 7, outline="white", width=2)
        canvas.create_oval(tx - 2, ty - 2, tx + 2, ty + 2, fill="white", outline="")

        # If target is in unreachable zone, show an arrow to the achievable target
        if t_hsv["s"] > ColorMatcherApp.GAME_S_MAX_PCT:
            ax = start_x + ColorMatcherApp.GAME_S_MAX_PCT / 100 * sq
            ay = ty
            canvas.create_oval(ax - 5, ay - 5, ax + 5, ay + 5,
                               outline="#fbbf24", fill="", width=2, dash=(2, 2))
            canvas.create_line(tx, ty, ax, ay, fill="#fbbf24", width=1, arrow="last")

        # Current marker (indigo)
        if c_hsv:
            cx = start_x + c_hsv["s"] / 100 * sq
            cy = y0 + (1 - c_hsv["v"] / 100) * sq
            canvas.create_oval(cx - 5, cy - 5, cx + 5, cy + 5, fill=INDIGO_LIGHT,
                               outline="white", width=1)

        # Hue bar (reversed: top=red→purple→blue→green→yellow→red=bottom)
        sl_x = start_x + sq + gap
        if HAS_NUMPY:
            hs = np.linspace(0, 1, sq)
            rows = [[int(c * 255) for c in colorsys.hsv_to_rgb((1 - h) % 1.0, 1, 1)] for h in hs]
            col = np.array(rows, dtype=np.uint8).reshape(sq, 1, 3)
            hue_arr = np.broadcast_to(col, (sq, sl_w, 3)).copy()
            hue_img = Image.fromarray(hue_arr, "RGB")
        else:
            hue_img = Image.new("RGB", (sl_w, sq))
            for py in range(sq):
                hh = (1 - py / sq) % 1.0
                r, g, b = [int(c * 255) for c in colorsys.hsv_to_rgb(hh, 1, 1)]
                for px in range(sl_w):
                    hue_img.putpixel((px, py), (r, g, b))

        self._hue_photo = ImageTk.PhotoImage(hue_img)
        canvas.create_image(sl_x, y0, anchor="nw", image=self._hue_photo)

        # Target hue marker (white rect)
        t_hue_y = y0 + ((360 - t_hsv["h"]) % 360) / 360 * sq
        canvas.create_rectangle(sl_x - 2, t_hue_y - 4, sl_x + sl_w + 2, t_hue_y + 4,
                                outline="white", width=2)

        # Current hue marker (indigo)
        if c_hsv:
            c_hue_y = y0 + ((360 - c_hsv["h"]) % 360) / 360 * sq
            canvas.create_rectangle(sl_x - 2, c_hue_y - 1, sl_x + sl_w + 2, c_hue_y + 1,
                                    fill=INDIGO_LIGHT, outline="white", width=1)

        # Labels
        canvas.create_text(start_x + sq // 2, y0 + sq + 10,
                           text="S→", fill=TEXT_DIM, font=("Segoe UI", 7))
        canvas.create_text(start_x - 8, y0 + sq // 2,
                           text="V↑", fill=TEXT_DIM, font=("Segoe UI", 7))
        canvas.create_text(sl_x + sl_w // 2, y0 + sq + 10,
                           text="H", fill=TEXT_DIM, font=("Segoe UI", 7))

    def _build_auto_section(self, parent):
        frame = tk.Frame(parent, bg=CARD_BG)
        frame.pack(fill="x", padx=20, pady=(8, 12))

        # Title row
        title_row = tk.Frame(frame, bg=CARD_BG)
        title_row.pack(fill="x")

        tk.Label(title_row, text="AUTO ADJUST", fg=INDIGO_LIGHT, bg=CARD_BG,
                 font=("Segoe UI", 9, "bold")).pack(side="left")

        # Calibration button + status (right side of title)
        self._calib_status = tk.Label(
            title_row,
            text="  Ready" if self.calibration else "  未設定",
            fg=GREEN if self.calibration else TEXT_DIM,
            bg=CARD_BG, font=("Segoe UI", 8)
        )
        self._calib_status.pack(side="right")

        self._calib_btn = tk.Button(
            title_row, text="calibrate", bg=CARD_BORDER, fg=TEXT_PRIMARY,
            font=("Segoe UI", 8), relief="flat", cursor="hand2",
            activebackground=HOVER_BG, command=self._start_calibration
        )
        self._calib_btn.pack(side="right", padx=(0, 4))

        # Settings row
        settings = tk.Frame(frame, bg=CARD_BG)
        settings.pack(fill="x", pady=(6, 0))

        tk.Checkbutton(
            settings, text="HEXのみ", variable=self._skip_click,
            bg=CARD_BG, fg=TEXT_PRIMARY, selectcolor=CARD_BORDER,
            activebackground=CARD_BG, activeforeground=TEXT_PRIMARY,
            font=("Segoe UI", 8)
        ).pack(side="left")

        tk.Checkbutton(
            settings, text="👁 瞳", variable=self._eye_mode,
            bg=CARD_BG, fg=TEXT_PRIMARY, selectcolor=CARD_BORDER,
            activebackground=CARD_BG, activeforeground=TEXT_PRIMARY,
            font=("Segoe UI", 8)
        ).pack(side="left", padx=(6, 0))

        tk.Checkbutton(
            settings, text="LUT", variable=self._use_lut,
            bg=CARD_BG, fg=TEXT_PRIMARY, selectcolor=CARD_BORDER,
            activebackground=CARD_BG, activeforeground=TEXT_PRIMARY,
            font=("Segoe UI", 8),
            state="normal" if self._lut is not None else "disabled",
            disabledforeground=TEXT_DIM,
        ).pack(side="left", padx=(6, 0))

        tk.Spinbox(settings, from_=1, to=20, increment=1, width=3,
                   textvariable=self._max_retries, font=("Segoe UI", 8),
                   bg=CARD_BORDER, fg=TEXT_PRIMARY, buttonbackground=CARD_BORDER
                   ).pack(side="right")
        tk.Label(settings, text="回数", fg=TEXT_DIM, bg=CARD_BG,
                 font=("Segoe UI", 8)).pack(side="right", padx=(4, 2))

        tk.Spinbox(settings, from_=0.5, to=20, increment=0.5, width=5,
                   textvariable=self._threshold, font=("Segoe UI", 8),
                   bg=CARD_BORDER, fg=TEXT_PRIMARY, buttonbackground=CARD_BORDER
                   ).pack(side="right")
        tk.Label(settings, text="ΔE", fg=TEXT_DIM, bg=CARD_BG,
                 font=("Segoe UI", 8)).pack(side="right", padx=(4, 2))

        # Execute button + status
        self._auto_btn = tk.Button(
            frame, text="▶  自動調整", bg=INDIGO, fg="white",
            font=("Segoe UI", 10, "bold"), relief="flat", cursor="hand2",
            activebackground=INDIGO_LIGHT, disabledforeground=TEXT_DIM,
            command=self._on_auto_adjust_click
        )
        self._auto_btn.pack(fill="x", pady=(8, 4), ipady=4)

        self._auto_status = tk.Label(
            frame, text="", fg=TEXT_DIM, bg=CARD_BG,
            font=("Consolas", 8), wraplength=350, justify="left"
        )
        self._auto_status.pack(anchor="w")

        self._update_auto_btn_state()

    def _update_auto_btn_state(self):
        has_target = self.target_color is not None
        has_calib = self.calibration is not None

        if has_target and has_calib:
            self._auto_btn.configure(state="normal", bg=INDIGO, fg="white")
        else:
            self._auto_btn.configure(state="disabled", bg=CARD_BORDER, fg=TEXT_DIM)
            parts = []
            if not has_target:
                parts.append("Target未設定")
            if not has_calib:
                parts.append("要calibrate")
            self._auto_status.configure(text=" / ".join(parts), fg=TEXT_DIM)

    # ── Calibration (1 click: HEX input field) ───────────────────

    def _start_calibration(self):
        # Auto-focus BPSR game window before showing overlay
        game_hwnd = self.selected_hwnd if (self.selected_hwnd and user32.IsWindow(self.selected_hwnd)) else find_game_window()
        if game_hwnd:
            self.selected_hwnd = game_hwnd
            # Restore if minimized
            if user32.IsIconic(game_hwnd):
                user32.ShowWindow(game_hwnd, SW_RESTORE)
            user32.SetForegroundWindow(game_hwnd)
            time.sleep(0.3)

        self.root.withdraw()
        self.root.update()
        time.sleep(0.2)

        sw = user32.GetSystemMetrics(0)
        sh = user32.GetSystemMetrics(1)
        screen_dc = user32.GetDC(0)
        mem_dc = gdi32.CreateCompatibleDC(screen_dc)
        bitmap = gdi32.CreateCompatibleBitmap(screen_dc, sw, sh)
        old = gdi32.SelectObject(mem_dc, bitmap)
        gdi32.BitBlt(mem_dc, 0, 0, sw, sh, screen_dc, 0, 0, SRCCOPY)
        bmi = BITMAPINFO()
        bmi.bmiHeader.biSize = ctypes.sizeof(BITMAPINFOHEADER)
        bmi.bmiHeader.biWidth = sw
        bmi.bmiHeader.biHeight = -sh
        bmi.bmiHeader.biPlanes = 1
        bmi.bmiHeader.biBitCount = 32
        bmi.bmiHeader.biCompression = BI_RGB
        buf = ctypes.create_string_buffer(sw * sh * 4)
        gdi32.GetDIBits(mem_dc, bitmap, 0, sh, buf, ctypes.byref(bmi), DIB_RGB_COLORS)
        gdi32.SelectObject(mem_dc, old)
        gdi32.DeleteObject(bitmap)
        gdi32.DeleteDC(mem_dc)
        user32.ReleaseDC(0, screen_dc)
        screen_img = Image.frombuffer("RGBX", (sw, sh), buf, "raw", "BGRX", 0, 1).convert("RGB")

        if HAS_NUMPY:
            arr = np.array(screen_img)
            arr = (arr * 0.6).astype(np.uint8)
            screen_img = Image.fromarray(arr)
        else:
            from PIL import ImageEnhance
            screen_img = ImageEnhance.Brightness(screen_img).enhance(0.6)

        self._calib_overlay = tk.Toplevel()
        self._calib_overlay.overrideredirect(True)
        self._calib_overlay.attributes("-topmost", True)
        self._calib_overlay.geometry(f"{sw}x{sh}+0+0")

        canvas = tk.Canvas(self._calib_overlay, width=sw, height=sh,
                           bg="black", highlightthickness=0, cursor="crosshair")
        canvas.pack(fill="both", expand=True)

        self._calib_bg_photo = ImageTk.PhotoImage(screen_img)
        canvas.create_image(0, 0, anchor="nw", image=self._calib_bg_photo)

        band_y = int(sh * 0.70)
        canvas.create_rectangle(0, band_y - 5, sw, band_y + 75,
                                fill="black", stipple="gray50", outline="")
        canvas.create_text(
            sw // 2, band_y + 20,
            text="ゲームのHEXコード入力欄をクリック",
            fill="white", font=("Segoe UI", 22, "bold"),
        )
        canvas.create_text(
            sw // 2, band_y + 55,
            text="Escでキャンセル",
            fill="#aaaaaa", font=("Segoe UI", 12),
        )

        self._calib_overlay.focus_force()
        self._calib_overlay.lift()

        def on_click(event):
            self.calibration = {
                "hex_input_pos": [event.x_root, event.y_root],
            }
            self._save_calibration()
            self._finish_calibration("Ready")

        canvas.bind("<Button-1>", on_click)
        self._calib_overlay.bind("<Escape>", self._cancel_calibration)

    def _cancel_calibration(self, event=None):
        self._finish_calibration("Ready" if self.calibration else "未設定")

    def _finish_calibration(self, status_text):
        if self._calib_overlay:
            self._calib_overlay.destroy()
            self._calib_overlay = None
        self.root.deiconify()
        self.root.update()
        is_ok = self.calibration is not None
        self._calib_status.configure(
            text=f"  {status_text}",
            fg=GREEN if is_ok else TEXT_DIM
        )
        self._update_auto_btn_state()

    # ── Legacy HSV → screen coordinate mapping (for click fallback) ──

    def _hsv_to_hue_screen_pos(self, hue_deg):
        """Map hue to screen pos. Used only when skip_click=False with legacy calibration."""
        if "hue_rect" not in self.calibration:
            return (0, 0)
        x1, y1, x2, y2 = self.calibration["hue_rect"]
        center_x = (x1 + x2) / 2
        lut = self.calibration.get("hue_to_y")
        if lut:
            return (center_x, lut[int(round(hue_deg)) % 361])
        mapped = (360 - hue_deg) % 360
        return (center_x, y1 + (mapped / 360) * (y2 - y1))

    def _sv_to_screen_pos(self, sat_pct, val_pct):
        """Map S/V to screen pos. Used only when skip_click=False with legacy calibration."""
        if "sv_rect" not in self.calibration:
            return (0, 0)
        x1, y1, x2, y2 = self.calibration["sv_rect"]
        return (x1 + (sat_pct / 100) * (x2 - x1),
                y1 + (1 - val_pct / 100) * (y2 - y1))

    # ── Auto-adjust execution ────────────────────────────────────

    def _on_auto_adjust_click(self):
        if not self.target_color or not self.calibration:
            return
        if not ctypes.windll.shell32.IsUserAnAdmin():
            self._auto_status.configure(
                text="管理者権限が必要です。右クリック→管理者として実行", fg="#ef4444")
            return
        self._cancel_adjust = False
        self._auto_btn.configure(state="disabled", bg=CARD_BORDER)
        self._auto_status.configure(text="調整中...", fg=INDIGO_LIGHT)
        threading.Thread(target=self._do_auto_adjust, daemon=True).start()

    def _sample_game_color(self, use_eye=False):
        """Capture the game window and read the color at the watch position.

        Default: 2 samples × 150ms, returns average (fast, for real-time display).
        use_eye=True: 6 samples × 150ms, returns per-channel median.
          Blinks (~0.3s = up to 2 samples) are rejected as outliers.
          Only enabled when the user has toggled _eye_mode ON.
        """
        if not self._watch_img_pos or not self.selected_hwnd:
            return None
        ix, iy = self._watch_img_pos

        eye_mode = use_eye and self._eye_mode.get()
        n_samples = 6 if eye_mode else 2

        samples = []
        for _ in range(n_samples):
            img = capture_window_printwindow(self.selected_hwnd)
            if img is None:
                img = capture_screen_region(self.selected_hwnd)
            if img and 0 <= ix < img.size[0] and 0 <= iy < img.size[1]:
                samples.append(img.getpixel((ix, iy))[:3])
            time.sleep(0.15)
        if not samples:
            return None

        if eye_mode and len(samples) >= 3:
            # Median per channel rejects blink outliers
            r = round(statistics.median(s[0] for s in samples))
            g = round(statistics.median(s[1] for s in samples))
            b = round(statistics.median(s[2] for s in samples))
        else:
            r = round(sum(s[0] for s in samples) / len(samples))
            g = round(sum(s[1] for s in samples) / len(samples))
            b = round(sum(s[2] for s in samples) / len(samples))
        return (r, g, b)

    # Game's accepted RGB range
    GAME_RGB_MIN = 0x1A  # 26
    GAME_RGB_MAX = 0xDA  # 218
    # Game's HSV saturation ceiling (flat across H/V, determined by test4)
    GAME_S_MAX_PCT = 78.0
    # Game's effective output V range (due to gamma transform)
    # Input 0x1A → output ~23%, Input 0xDA → output ~91%
    GAME_V_MIN_PCT = 23.0
    GAME_V_MAX_PCT = 91.0
    # HSV saturation limit: S_max = 0.25 * V + 55 (with safety margin)
    # From testing:
    #   V=85.5% → S_max≈79%, V=62.7% → S_max≈71%, V=50.2% → S_max≈70%

    @staticmethod
    def _clamp_rgb(r, g, b):
        """Clamp RGB to game's allowed range, respecting HSV saturation limits.

        The game rejects high-saturation colors. The saturation limit depends
        on value (brightness): S_max = 0.25 * V + 55 (percent).
        When exceeded, saturation is reduced while preserving hue and value.
        """
        ri, gi, bi = int(r), int(g), int(b)
        MIN = ColorMatcherApp.GAME_RGB_MIN
        MAX = ColorMatcherApp.GAME_RGB_MAX

        # 1. Absolute range clamp
        ri = max(MIN, min(MAX, ri))
        gi = max(MIN, min(MAX, gi))
        bi = max(MIN, min(MAX, bi))

        # 2. HSV saturation limit
        h, s, v = colorsys.rgb_to_hsv(ri / 255, gi / 255, bi / 255)
        s_pct = s * 100
        v_pct = v * 100
        # Based on exhaustive H/V testing (test4): game has a flat S ceiling
        # at ~80% regardless of hue or value. Use 78 for small safety margin.
        s_max = ColorMatcherApp.GAME_S_MAX_PCT

        if s_pct > s_max and v_pct > 0:
            logging.info(f"彩度制限: S={s_pct:.1f}% > S_max={s_max:.1f}% → S={s_max:.1f}%に制限")
            new_s = s_max / 100
            nr, ng, nb = colorsys.hsv_to_rgb(h, new_s, v)
            ri = max(MIN, min(MAX, int(nr * 255)))
            gi = max(MIN, min(MAX, int(ng * 255)))
            bi = max(MIN, min(MAX, int(nb * 255)))

        return (ri, gi, bi)

    @staticmethod
    def _rgb_to_hex(r, g, b):
        """Clamp RGB to game's allowed range and return 6-char hex."""
        cr, cg, cb = ColorMatcherApp._clamp_rgb(r, g, b)
        return f"{cr:02X}{cg:02X}{cb:02X}"

    def _do_auto_adjust(self):
        def _status(text, color=INDIGO_LIGHT):
            self.root.after(0, lambda: self._auto_status.configure(text=text, fg=color))

        max_retries = self._max_retries.get()
        threshold = self._threshold.get()
        skip_click = self._skip_click.get()

        try:
            tr, tg, tb = self.target_color

            hex_pos = self.calibration.get("hex_input_pos")
            has_hex = hex_pos is not None

            _status("ウィンドウを非表示中...")
            self.root.after(0, self.root.withdraw)
            time.sleep(0.3)

            if self.selected_hwnd:
                user32.SetForegroundWindow(self.selected_hwnd)
                time.sleep(0.5)

            # === Phase 1: Initial set via click (skippable) ===
            if not skip_click or not has_hex:
                target_hsv = rgb_to_hsv(tr, tg, tb)
                hx, hy = self._hsv_to_hue_screen_pos(target_hsv["h"])
                _status("[1] 色相クリック...")
                click_at(hx, hy, delay=0.1)
                time.sleep(0.3)

                sx, sy = self._sv_to_screen_pos(target_hsv["s"], target_hsv["v"])
                _status("[1] SVクリック...")
                click_at(sx, sy, delay=0.1)
                time.sleep(0.3)

            # === Phase 2: Initial HEX input (LUT path or legacy clamp-send) ===
            use_lut = self._use_lut.get() and self._lut is not None and has_hex

            # Default initial state (overwritten by whichever path runs below)
            cr, cg, cb = self._clamp_rgb(tr, tg, tb)
            input_r, input_g, input_b = float(cr), float(cg), float(cb)
            last_hex_sent = ""

            # LUT-path tracking (None/empty if legacy path taken)
            lut_predicted_de = None
            lut_refinements = []  # list of (tag, hex_sent, measured, de)
            lut_best_de = 999.0
            lut_best_hex = ""
            lut_best_result = None
            lut_converged = False  # True if B1/B2 already hit threshold

            if use_lut:
                # --- A1: LUT trilinear inverse lookup ---
                # NOTE: find_best_input takes ~300-600ms and is not cancellable
                # (single non-responsive window per refinement round).
                _status("LUT探索中...")
                try:
                    import lut_lookup  # lazy import to avoid circular dep at module scope
                    best_in, pred_out, pred_de = lut_lookup.find_best_input(
                        self.target_color, self._lut
                    )
                except Exception as e:
                    logging.info(f"LUT探索失敗: {e}")
                    best_in = None
                    pred_out = None
                    pred_de = None

                if best_in is None:
                    logging.info("LUT探索結果なし → fallback to clamp initial")
                    use_lut = False
                else:
                    lut_predicted_de = pred_de
                    cr, cg, cb = best_in
                    hex_code = self._rgb_to_hex(cr, cg, cb)
                    input_r, input_g, input_b = float(cr), float(cg), float(cb)
                    last_hex_sent = hex_code

                    if pred_de >= 15:
                        _status(f"到達困難色 予測ΔE={pred_de:.1f} #{hex_code}", "#eab308")
                    else:
                        _status(f"[A1] LUT入力 #{hex_code} 予測ΔE={pred_de:.1f}")

                    logging.info(
                        f"LUT A1: best_in={best_in} pred_out={pred_out} 予測ΔE={pred_de:.2f}"
                    )

                    type_hex_into_field(hex_pos[0], hex_pos[1], hex_code)
                    time.sleep(0.8)

                    # --- B1/B2: Lab-residual refinement (max 2 rounds) ---
                    if self._watch_img_pos:
                        target_lab = rgb_to_lab(*self.target_color)
                        effective_target_lab = target_lab
                        for b_round in (1, 2):
                            if self._cancel_adjust:
                                break
                            _status(f"[B{b_round}] 実測中...")
                            measured = self._sample_game_color(use_eye=True)
                            if measured is None:
                                logging.info(f"LUT B{b_round}: キャプチャ失敗")
                                break
                            de_now = delta_e(self.target_color, measured)
                            lut_refinements.append(
                                (f"B{b_round}", last_hex_sent, measured, de_now)
                            )
                            logging.info(
                                f"LUT B{b_round}: sent=#{last_hex_sent} "
                                f"measured={measured} ΔE={de_now:.2f}"
                            )

                            if de_now < lut_best_de:
                                lut_best_de = de_now
                                lut_best_hex = last_hex_sent
                                lut_best_result = measured

                            if de_now <= threshold:
                                _status(
                                    f"[B{b_round}] 完了 ΔE={de_now:.2f} "
                                    f"({delta_e_category(de_now)})",
                                    GREEN,
                                )
                                lut_converged = True
                                break

                            # Safety: runaway measurement → bail to iterative loop
                            if de_now > 30:
                                logging.info(
                                    f"LUT B{b_round}: 実測ΔE>30 — 反復ループにハンドオフ"
                                )
                                break

                            # Compute Lab residual, adjust virtual target, re-search
                            measured_lab = rgb_to_lab(*measured)
                            effective_target_lab = (
                                2 * effective_target_lab[0] - measured_lab[0],
                                2 * effective_target_lab[1] - measured_lab[1],
                                2 * effective_target_lab[2] - measured_lab[2],
                            )
                            eff_rgb = lab_to_rgb_approx(effective_target_lab)
                            _status(f"[B{b_round}] 残差補正探索...")
                            try:
                                new_best_in, _, _ = lut_lookup.find_best_input(
                                    eff_rgb, self._lut
                                )
                            except Exception as e:
                                logging.info(f"LUT B{b_round} 探索失敗: {e}")
                                break
                            if new_best_in is None:
                                break
                            if new_best_in == (int(cr), int(cg), int(cb)):
                                logging.info(
                                    f"LUT B{b_round}: 同一入力 — 反復ループにハンドオフ"
                                )
                                break
                            cr, cg, cb = new_best_in
                            hex_code = self._rgb_to_hex(cr, cg, cb)
                            input_r, input_g, input_b = float(cr), float(cg), float(cb)
                            last_hex_sent = hex_code
                            type_hex_into_field(hex_pos[0], hex_pos[1], hex_code)
                            time.sleep(0.8)

            if not use_lut:
                # --- Legacy path: send clamped target HEX as initial guess ---
                if has_hex:
                    before = self._sample_game_color(use_eye=True) if self._watch_img_pos else None
                    hex_code = self._rgb_to_hex(tr, tg, tb)
                    _status(f"[2] HEX入力: #{hex_code}")
                    type_hex_into_field(hex_pos[0], hex_pos[1], hex_code)
                    time.sleep(0.8)
                    after = self._sample_game_color(use_eye=True) if self._watch_img_pos else None
                    logging.info(f"  HEX入力テスト: #{hex_code}")
                    logging.info(f"    入力前: {before}")
                    logging.info(f"    入力後: {after}")
                    if before and after and before == after:
                        logging.info(f"    ⚠ 色が変わっていない！HEX入力が効いていない可能性")
                    last_hex_sent = hex_code

                cr, cg, cb = self._clamp_rgb(tr, tg, tb)
                input_r, input_g, input_b = float(cr), float(cg), float(cb)

            # === Phase 3: Verify + correct loop ===
            # Always correct towards the ORIGINAL target.
            # _rgb_to_hex handles clamping at output time only.
            # The game's gamma/color transform means input!=output, so we let
            # the correction loop discover the right input values empirically.

            if self._watch_img_pos:
                # Inherit best from LUT B1/B2 if they ran; otherwise fresh
                if use_lut and lut_refinements:
                    best_de = lut_best_de
                    best_score = (similarity_score(self.target_color, lut_best_result)
                                  if lut_best_result else 0.0)
                    best_result = lut_best_result
                    best_hex = lut_best_hex
                else:
                    best_de = 999.0
                    best_score = 0.0
                    best_result = None
                    best_hex = ""
                prev_result = None
                stall_count = 0
                same_hex_count = 0
                decline_count = 0
                recent_results = []
                fine_tune_mode = False  # activated after first oscillation, uses smaller steps
                fine_tune_oscillations = 0

                def get_damping(attempt, max_err, declining, fine):
                    if fine:
                        # Fine-tune mode: very small steps to squeeze last bit of accuracy
                        return max(0.15, 0.25 - fine_tune_oscillations * 0.05)
                    base = min(0.95, 0.5 + attempt * 0.05)
                    err_factor = min(1.0, max(0.6, 1.0 - max_err / 400))
                    d = base * err_factor
                    if declining:
                        d = min(d, 0.5)
                    return d

                logging.info(f"\n{'='*80}")
                logging.info(f"自動調整開始")
                logging.info(f"  Target: ({tr}, {tg}, {tb}) = #{tr:02X}{tg:02X}{tb:02X}")
                logging.info(f"  初期HEX(クランプ済): #{cr:02X}{cg:02X}{cb:02X}")
                logging.info(f"  判定: Delta E (CIE76) 知覚距離")
                logging.info(f"  設定: ΔE閾値={threshold} 最大={max_retries}回")
                if lut_predicted_de is not None:
                    logging.info(
                        f"  LUT: 予測ΔE={lut_predicted_de:.2f}  "
                        f"refinements={len(lut_refinements)}"
                    )
                    for tag, h, m, de_r in lut_refinements:
                        logging.info(f"    {tag}: #{h} → {m} ΔE={de_r:.2f}")
                if lut_converged:
                    logging.info(f"  LUT で閾値到達 → 反復ループスキップ (best ΔE={best_de:.2f})")
                logging.info(f"{'='*80}")
                logging.info(f"{'回':>3} | {'送信HEX':>8} | {'実測RGB':>15} | {'誤差':>15} | {'input':>20} | {'damp':>5} | {'RGB%':>5} | {'ΔE':>5} | {'best':>5} | {'状態'}")
                logging.info(f"{'-'*3}-+-{'-'*8}-+-{'-'*15}-+-{'-'*15}-+-{'-'*20}-+-{'-'*5}-+-{'-'*5}-+-{'-'*5}-+-{'-'*5}-+-{'-'*10}")

                # Skip iterative loop if LUT already converged
                effective_max_retries = 0 if lut_converged else max_retries
                for attempt in range(effective_max_retries):
                    if self._cancel_adjust:
                        logging.info("ユーザーによるキャンセル")
                        _status("キャンセル", TEXT_DIM)
                        break
                    result = self._sample_game_color(use_eye=True)
                    if result is None:
                        logging.info(f"{attempt+1:3} | {'':>8} | {'キャプチャ失敗':>15} |")
                        _status(f"[補正{attempt+1}] キャプチャ失敗", "#ef4444")
                        break

                    score_orig = similarity_score(self.target_color, result)
                    de = delta_e(self.target_color, result)
                    _status(f"[補正{attempt+1}/{max_retries}] ΔE={de:.1f} ({delta_e_category(de)}) {score_orig:.1f}%")

                    # Track best by Delta E (perceptual)
                    if de < best_de:
                        best_de = de
                        best_score = score_orig
                        best_result = result
                        best_hex = last_hex_sent
                        stall_count = 0
                        decline_count = 0
                    else:
                        decline_count += 1

                    # Error vs target in RGB space (correction still happens in RGB)
                    rr, rg, rb = result
                    err_r = tr - rr
                    err_g = tg - rg
                    err_b = tb - rb
                    max_err = max(abs(err_r), abs(err_g), abs(err_b), 1)
                    declining = decline_count >= 2
                    damping = get_damping(attempt, max_err, declining, fine_tune_mode)

                    # Detect oscillation: current result matches a past result (within 2 units)
                    oscillating = False
                    for past in recent_results:
                        if abs(rr-past[0]) <= 2 and abs(rg-past[1]) <= 2 and abs(rb-past[2]) <= 2:
                            oscillating = True
                            break
                    recent_results.append(result)
                    if len(recent_results) > 3:
                        recent_results.pop(0)

                    # Stop conditions (based on Delta E, the perceptual metric)
                    stop = ""
                    if best_de <= threshold:
                        stop = f"閾値到達(ΔE={best_de:.1f})"
                    elif decline_count >= 5 and fine_tune_mode:
                        stop = f"微調整でも改善停滞→best復帰(ΔE={best_de:.1f})"
                    elif oscillating and attempt >= 3:
                        if not fine_tune_mode:
                            # First oscillation: enter fine-tune mode instead of stopping
                            fine_tune_mode = True
                            fine_tune_oscillations = 0
                            recent_results.clear()
                            decline_count = 0
                            logging.info(f"  → 振動検出、微調整モード開始 (damping 0.25→)")
                        else:
                            fine_tune_oscillations += 1
                            if fine_tune_oscillations >= 2:
                                stop = f"微調整完了(ΔE={best_de:.1f} {delta_e_category(best_de)})"
                            else:
                                # Oscillating in fine-tune, reduce damping further
                                recent_results.clear()
                                logging.info(f"  → 微調整中再振動、dampingさらに低下")
                    elif prev_result and abs(rr-prev_result[0]) <= 1 and \
                         abs(rg-prev_result[1]) <= 1 and abs(rb-prev_result[2]) <= 1:
                        stall_count += 1
                        if stall_count >= 2 and best_de > 30 and attempt <= 2:
                            stop = "入力未反映(要キャリブ再設定)"
                        elif stall_count >= 3:
                            stop = f"ゲーム上限到達(ΔE={best_de:.1f} {delta_e_category(best_de)})"
                    else:
                        stall_count = 0

                    logging.info(f"{attempt+1:3} | #{last_hex_sent:>6} | ({rr:3d},{rg:3d},{rb:3d}) | ({err_r:+4d},{err_g:+4d},{err_b:+4d}) | ({input_r:6.1f},{input_g:6.1f},{input_b:6.1f}) | {damping:.2f} | {score_orig:4.1f}% | {de:5.1f} | {best_de:5.1f} | {stop}")

                    if stop:
                        if "未反映" in stop:
                            _status(f"入力が効いていません。キャリブレーションを再設定してください", "#ef4444")
                        elif "閾値" in stop:
                            _status(f"完了 ΔE={best_de:.1f} ({delta_e_category(best_de)}) RGB {best_score:.1f}%", GREEN)
                        else:
                            color = GREEN if best_de < 10 else "#eab308" if best_de < 20 else "#ef4444"
                            _status(f"完了 ΔE={best_de:.1f} ({delta_e_category(best_de)}) RGB {best_score:.1f}%", color)
                        break

                    prev_result = result

                    # Apply correction with damping, clamp to prevent divergence
                    input_r = max(-50, min(300, input_r + err_r * damping))
                    input_g = max(-50, min(300, input_g + err_g * damping))
                    input_b = max(-50, min(300, input_b + err_b * damping))

                    if has_hex:
                        corrected_hex = self._rgb_to_hex(input_r, input_g, input_b)

                        # Note: input_r/g/b intentionally NOT clamped to game range.
                        # Only the HEX output is clamped. This lets input drift freely
                        # below the saturation floor, so channels stuck at HEX limits
                        # can still explore via the other channels' correction.

                        if corrected_hex == last_hex_sent:
                            same_hex_count += 1
                            if same_hex_count >= 3:
                                logging.info(f"  → HEX同一 #{corrected_hex} ×3回 — 収束")
                                _status(f"収束 ΔE={best_de:.1f} RGB {best_score:.1f}%", GREEN)
                                break
                            input_r += 1 if err_r > 0 else -1 if err_r < 0 else 0
                            input_g += 1 if err_g > 0 else -1 if err_g < 0 else 0
                            input_b += 1 if err_b > 0 else -1 if err_b < 0 else 0
                            corrected_hex = self._rgb_to_hex(input_r, input_g, input_b)
                            if corrected_hex == last_hex_sent:
                                logging.info(f"  → HEX同一(nudge後も) #{corrected_hex} — 収束")
                                _status(f"収束 ΔE={best_de:.1f} RGB {best_score:.1f}%", GREEN)
                                break
                        else:
                            same_hex_count = 0

                        type_hex_into_field(hex_pos[0], hex_pos[1], corrected_hex)
                        last_hex_sent = corrected_hex
                        time.sleep(0.8)
                    else:
                        corr_hsv = rgb_to_hsv(input_r, input_g, input_b)
                        chx, chy = self._hsv_to_hue_screen_pos(corr_hsv["h"])
                        click_at(chx, chy, delay=0.1)
                        time.sleep(0.2)
                        csx, csy = self._sv_to_screen_pos(corr_hsv["s"], corr_hsv["v"])
                        click_at(csx, csy, delay=0.1)
                        time.sleep(0.4)

                logging.info(f"{'-'*80}")
                logging.info(f"  Best: ΔE={best_de:.1f} RGB={best_score:.1f}% {best_result} (HEX: #{best_hex})")

                # Restore best HEX if we overshot
                if has_hex and best_hex and best_hex != last_hex_sent and best_result is not None:
                    logging.info(f"  → 最良値に復帰: #{best_hex}")
                    type_hex_into_field(hex_pos[0], hex_pos[1], best_hex)
                    time.sleep(0.8)

                # Final check
                final = self._sample_game_color(use_eye=True)
                use = final if final else best_result
                if use:
                    score = similarity_score(self.target_color, use)
                    de_final = delta_e(self.target_color, use)
                    self.current_color = use
                    self.root.after(0, self._update_display)
                    color = GREEN if de_final < 5 else "#eab308" if de_final < 15 else "#ef4444"
                    _status(f"調整完了 ΔE={de_final:.1f} ({delta_e_category(de_final)}) RGB {score:.1f}%", color)
                else:
                    _status("調整完了 (検証なし)", GREEN)
            else:
                _status(
                    f"調整完了 (検証なし) #{self._rgb_to_hex(tr, tg, tb)}",
                    GREEN
                )

            self.root.after(0, self.root.deiconify)

        except Exception as e:
            self.root.after(0, self.root.deiconify)
            _status(f"エラー: {e}", "#ef4444")
        finally:
            self.root.after(0, lambda: self._auto_btn.configure(
                state="normal", bg=INDIGO
            ))

    # ── Window picker overlay ────────────────────────────────────

    def _open_window_picker(self, mode):
        self.pick_mode = mode

        # "watch" mode: skip picker, go straight to BPSR capture
        if mode == "watch":
            hwnd = self.selected_hwnd
            if not hwnd or not user32.IsWindow(hwnd):
                hwnd = find_game_window()
            if hwnd:
                self.selected_hwnd = hwnd
                self._capture_and_pick(hwnd)
                return

        # "target" mode or game not found: show manual window picker
        self.window_list = get_window_list()
        my_title = self.root.title()
        self.window_list = [
            w for w in self.window_list
            if w["title"] != my_title and w["title"] != "対象を選択"
        ]
        def _sort_key(w):
            for i, name in enumerate(GAME_WINDOW_NAMES):
                if name in w["title"]:
                    return (0, i)
            return (1, w["title"])
        self.window_list.sort(key=_sort_key)

        self._picker_win = tk.Toplevel(self.root)
        self._picker_win.title("対象を選択")
        self._picker_win.configure(bg=BG)
        self._picker_win.geometry("700x500")
        self._picker_win.transient(self.root)
        self._picker_win.grab_set()

        header = tk.Frame(self._picker_win, bg=BG)
        header.pack(fill="x", padx=24, pady=(16, 8))
        tk.Label(header, text="🖥  対象を選択", fg=TEXT_PRIMARY, bg=BG,
                 font=("Segoe UI", 18, "bold")).pack(side="left")
        tk.Button(header, text="✕", fg=TEXT_DIM, bg=BG, font=("Segoe UI", 14),
                  relief="flat", command=self._picker_win.destroy,
                  activebackground="#27272a").pack(side="right")

        container = tk.Frame(self._picker_win, bg=BG)
        container.pack(fill="both", expand=True, padx=24, pady=(0, 16))

        canvas = tk.Canvas(container, bg=BG, highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        scroll_frame = tk.Frame(canvas, bg=BG)

        scroll_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        cw_id = canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        def _on_canvas_configure(event):
            canvas.itemconfig(cw_id, width=event.width)
        canvas.bind("<Configure>", _on_canvas_configure)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        for winfo in self.window_list:
            btn = tk.Button(
                scroll_frame, text=winfo["title"], bg=CARD_BG, fg=TEXT_PRIMARY,
                font=("Segoe UI", 11), anchor="w", relief="flat", padx=16, pady=10,
                activebackground=HOVER_BG, activeforeground=TEXT_PRIMARY,
                cursor="hand2",
                command=lambda w=winfo: self._on_window_selected(w)
            )
            btn.pack(fill="x", pady=2)
            btn.bind("<Enter>", lambda e, b=btn: b.configure(bg=HOVER_BG))
            btn.bind("<Leave>", lambda e, b=btn: b.configure(bg=CARD_BG))

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        self._picker_win.bind_all("<MouseWheel>", _on_mousewheel)
        self._picker_win.bind("<Destroy>",
                              lambda e: self._picker_win.unbind_all("<MouseWheel>"))

    def _capture_and_pick(self, hwnd):
        """Capture a window and open the image picker directly."""
        def _hide():
            self.root.withdraw()
            self.root.update()
        def _show():
            self.root.deiconify()
            self.root.update()

        img = capture_window(hwnd, hide_callback=_hide, show_callback=_show)
        if img is None:
            return
        self.captured_image = img
        # Only update selected_hwnd for "watch" mode (keep BPSR for monitoring)
        if self.pick_mode == "watch":
            self.selected_hwnd = hwnd
        self._open_image_picker(img)

    def _on_window_selected(self, winfo):
        self._picker_win.destroy()
        self._capture_and_pick(winfo["id"])

    # ── Image picker ─────────────────────────────────────────────

    def _open_image_picker(self, img):
        self._img_win = tk.Toplevel(self.root)
        self._img_win.title("色を選択 (クリックで色を取得 / Escで戻る)")
        self._img_win.configure(bg="black", cursor="crosshair")
        self._img_win.attributes("-fullscreen", True)
        self._img_win.attributes("-topmost", True)
        self._img_win.focus_force()
        self._img_win.lift()

        self._pick_img = img
        self._pick_photo = None
        self.hover_color = None
        self._img_configured = False

        canvas = tk.Canvas(self._img_win, bg="black", highlightthickness=0)
        canvas.pack(fill="both", expand=True)
        self._pick_canvas = canvas

        self._img_offset_x = 0
        self._img_offset_y = 0
        self._img_display_w = 1
        self._img_display_h = 1
        self._img_scale = 1.0

        def _on_configure(event):
            if self._img_configured:
                return
            self._img_configured = True

            cw, ch = event.width, event.height
            if cw < 10 or ch < 10:
                return
            iw, ih = img.size
            scale = min(cw / iw, ch / ih)
            new_w, new_h = int(iw * scale), int(ih * scale)
            resized = img.resize((new_w, new_h), Image.LANCZOS)
            self._pick_photo = ImageTk.PhotoImage(resized)
            self._img_offset_x = (cw - new_w) // 2
            self._img_offset_y = (ch - new_h) // 2
            self._img_display_w = new_w
            self._img_display_h = new_h
            self._img_scale = scale
            canvas.create_image(self._img_offset_x, self._img_offset_y,
                                anchor="nw", image=self._pick_photo, tags="img")

            canvas.create_text(cw // 2, 30,
                               text="クリックで色を取得  |  Escで戻る",
                               fill="#aaaaaa", font=("Segoe UI", 12),
                               tags="help")

        canvas.bind("<Configure>", _on_configure)

        def _on_motion(event):
            ix = event.x - self._img_offset_x
            iy = event.y - self._img_offset_y
            if 0 <= ix < self._img_display_w and 0 <= iy < self._img_display_h:
                ox = min(int(ix / self._img_scale), img.size[0] - 1)
                oy = min(int(iy / self._img_scale), img.size[1] - 1)
                r, g, b = img.getpixel((ox, oy))[:3]
                self.hover_color = (r, g, b)
                self.cursor_img_pos = (ox, oy)

                canvas.delete("mag")
                mx = event.x + 25
                my = event.y - 45
                if mx + 80 > canvas.winfo_width():
                    mx = event.x - 105
                if my < 10:
                    my = event.y + 25

                color_hex = f"#{r:02x}{g:02x}{b:02x}"
                canvas.create_rectangle(mx, my, mx + 80, my + 40,
                                        fill=color_hex, outline="white", width=3,
                                        tags="mag")
                canvas.create_rectangle(mx, my + 40, mx + 80, my + 58,
                                        fill="black", outline="white", width=1,
                                        tags="mag")
                canvas.create_text(mx + 40, my + 49,
                                   text=f"R:{r} G:{g} B:{b}",
                                   fill="white", font=("Consolas", 8),
                                   tags="mag")
                canvas.delete("cross")
                cx, cy = event.x, event.y
                canvas.create_line(cx - 15, cy, cx - 5, cy, fill="white",
                                   width=1, tags="cross")
                canvas.create_line(cx + 5, cy, cx + 15, cy, fill="white",
                                   width=1, tags="cross")
                canvas.create_line(cx, cy - 15, cx, cy - 5, fill="white",
                                   width=1, tags="cross")
                canvas.create_line(cx, cy + 5, cx, cy + 15, fill="white",
                                   width=1, tags="cross")
            else:
                self.hover_color = None
                canvas.delete("mag")
                canvas.delete("cross")

        canvas.bind("<Motion>", _on_motion)

        def _on_click(event):
            if self.hover_color is None:
                return
            color = self.hover_color
            if self.pick_mode == "target":
                self.target_color = color
                self._img_win.destroy()
                self._update_display()
            elif self.pick_mode == "watch":
                self.current_color = color
                # Record image coordinates for real-time polling
                if self.cursor_img_pos and self.selected_hwnd:
                    self._watch_img_pos = (
                        self.cursor_img_pos[0],
                        self.cursor_img_pos[1],
                    )
                self._img_win.destroy()
                self._update_display()
                # Start polling
                self._start_polling()

        canvas.bind("<Button-1>", _on_click)
        self._img_win.bind("<Escape>", lambda e: self._img_win.destroy())

    # ── Real-time polling ───────────────────────────────────────

    def _start_polling(self):
        """Start polling the watch pixel every 400ms."""
        self._polling = True
        self._poll_tick()

    def _stop_polling(self):
        self._polling = False
        if self._poll_after_id:
            self.root.after_cancel(self._poll_after_id)
            self._poll_after_id = None

    def _poll_tick(self):
        if not self._polling or not self._watch_img_pos or not self.selected_hwnd:
            return
        if not user32.IsWindow(self.selected_hwnd):
            self._polling = False
            self._auto_status.configure(text="ゲームウィンドウが見つかりません", fg="#ef4444")
            return
        ix, iy = self._watch_img_pos
        try:
            img = capture_window_printwindow(self.selected_hwnd)
            method = "PrintWindow"
            if img is None:
                img = capture_screen_region(self.selected_hwnd)
                method = "BitBlt(画面)"

            if img and 0 <= ix < img.size[0] and 0 <= iy < img.size[1]:
                r, g, b = img.getpixel((ix, iy))[:3]
                self.current_color = (r, g, b)
                self._update_display()

                # Show capture method on first tick (then only on change)
                if not hasattr(self, '_last_poll_method') or self._last_poll_method != method:
                    self._last_poll_method = method
                    hint = ""
                    if method == "BitBlt(画面)":
                        hint = " (ウィンドウを重ねないで)"
                    self._auto_status.configure(
                        text=f"監視中: {method}{hint}", fg=TEXT_DIM
                    )
            elif img:
                self._auto_status.configure(
                    text=f"座標が範囲外 ({ix},{iy}) 画像={img.size}", fg="#ef4444"
                )
        except Exception as e:
            self._auto_status.configure(text=f"監視エラー: {e}", fg="#ef4444")
        self._poll_after_id = self.root.after(400, self._poll_tick)

    # ── Display update ───────────────────────────────────────────

    def _update_display(self):
        if self.target_color:
            r, g, b = self.target_color
            self._target_swatch.configure(bg=f"#{r:02x}{g:02x}{b:02x}")
            self._target_label.configure(text=f"{r}, {g}, {b}")
            hex_code = f"#{r:02X}{g:02X}{b:02X}"
            # Show clamped (achievable) HEX if different from target
            cr, cg, cb = ColorMatcherApp._clamp_rgb(r, g, b)
            if (cr, cg, cb) != (r, g, b):
                hex_code = f"{hex_code} → #{cr:02X}{cg:02X}{cb:02X}"
            self._target_hex.configure(text=hex_code)
        else:
            self._target_swatch.configure(bg="#18181b")
            self._target_label.configure(text="---")
            self._target_hex.configure(text="")

        if self.current_color:
            r, g, b = self.current_color
            self._current_swatch.configure(bg=f"#{r:02x}{g:02x}{b:02x}")
            self._current_label.configure(text=f"{r}, {g}, {b}")
            self._current_hex.configure(text=f"#{r:02X}{g:02X}{b:02X}")
        else:
            self._current_swatch.configure(bg="#18181b")
            self._current_label.configure(text="---")
            self._current_hex.configure(text="")

        score = similarity_score(self.target_color, self.current_color)
        self._score_label.configure(text=f"{score:.1f}")
        self._bar_fill.place(relx=0, rely=0, relheight=1, relwidth=score / 100)

        # Delta E (perceptual distance)
        if self.target_color and self.current_color:
            de = delta_e(self.target_color, self.current_color)
            cat = delta_e_category(de)
            de_color = GREEN if de < 5 else INDIGO_LIGHT if de < 10 else "#eab308" if de < 20 else "#ef4444"
            self._de_label.configure(text=f"ΔE {de:.1f}  ({cat})", fg=de_color)
        else:
            self._de_label.configure(text="")

        hint = generate_hint(self.target_color, self.current_color)
        self._hint_label.configure(text=f"「{hint}」")

        self._update_auto_btn_state()
        self._draw_hsv_comparison()


# ── Entry point ───────────────────────────────────────────────────

def run_as_admin():
    """Re-launch this script with admin privileges if not already elevated."""
    import sys
    if ctypes.windll.shell32.IsUserAnAdmin():
        return True  # already admin
    # Re-run with ShellExecuteW "runas"
    params = " ".join(f'"{a}"' for a in sys.argv)
    ret = ctypes.windll.shell32.ShellExecuteW(
        None, "runas", sys.executable, params, None, 1
    )
    # ShellExecuteW returns >32 on success
    return ret > 32


def main():
    ctypes.windll.shcore.SetProcessDpiAwareness(2)

    # Auto-elevate to admin (needed for SendInput to games running as admin)
    if not ctypes.windll.shell32.IsUserAnAdmin():
        if run_as_admin():
            return  # new elevated process started, exit this one
        # User declined UAC — run anyway but warn later

    root = tk.Tk()
    app = ColorMatcherApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
