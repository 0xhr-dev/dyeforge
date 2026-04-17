"""
DyeForge - 3D LUT 逆引きモジュール

lut_3d.json (校正データ) を使って、target color から最良の入力 HEX を求める。

主要関数:
  load_lut(path) → LUT 辞書
  trilinear_forward(in_rgb, lut) → 予測出力 (Lab tuple)
  find_best_input(target_rgb, lut, fine=True) → (best_input, predicted_output, predicted_de)
  is_reachable_input(in_rgb) → bool (color_matcher.py の _clamp_rgb 通過判定)
"""

import json
import math

from color_matcher import rgb_to_lab, delta_e


# ── 設定 (lut_3d_calibration.py と一致させる) ──

GRID_LEVELS = [26, 53, 80, 107, 134, 161, 188, 218]
GAME_RGB_MIN = 26
GAME_RGB_MAX = 218
GAME_S_MAX_PCT = 78.0


# ── LUT 読み込み ──

def load_lut(path):
    """JSON LUT を読み込んで dict に変換。
    Returns: {(in_R, in_G, in_B): {"out": (R, G, B), "out_lab": (L, a, b)}}
    Lab を事前計算してキャッシュしておく (探索ループ高速化のため)。
    """
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    lut = {}
    for v in data["results"].values():
        if v.get("out") is None:
            continue
        in_rgb = tuple(v["in"])
        out_rgb = tuple(v["out"])
        out_lab = rgb_to_lab(*out_rgb)
        lut[in_rgb] = {"out": out_rgb, "out_lab": out_lab}

    return lut


# ── 入力の到達可能性判定 ──

def is_reachable_input(in_rgb):
    """color_matcher.py._clamp_rgb を modify なしで通過するか。
    (per-channel [26, 218] かつ S <= 78%)
    """
    r, g, b = in_rgb
    if not (GAME_RGB_MIN <= r <= GAME_RGB_MAX and
            GAME_RGB_MIN <= g <= GAME_RGB_MAX and
            GAME_RGB_MIN <= b <= GAME_RGB_MAX):
        return False
    mx, mn = max(r, g, b), min(r, g, b)
    if mx == 0:
        return False
    s = (mx - mn) / mx * 100
    return s <= GAME_S_MAX_PCT


# ── Trilinear 補間 (forward map) ──

def _bracket(v):
    """v を含む GRID_LEVELS の区間 (lo, hi, t) を返す。t は [0, 1]。"""
    if v <= GRID_LEVELS[0]:
        return GRID_LEVELS[0], GRID_LEVELS[0], 0.0
    if v >= GRID_LEVELS[-1]:
        return GRID_LEVELS[-1], GRID_LEVELS[-1], 0.0
    for i in range(len(GRID_LEVELS) - 1):
        if GRID_LEVELS[i] <= v <= GRID_LEVELS[i + 1]:
            lo, hi = GRID_LEVELS[i], GRID_LEVELS[i + 1]
            t = (v - lo) / (hi - lo)
            return lo, hi, t
    return GRID_LEVELS[-1], GRID_LEVELS[-1], 0.0


def trilinear_forward(in_rgb, lut):
    """LUT を使って任意の input に対する予測出力を計算 (Lab で返す)。
    8 corner のうち欠損 (S>78% で未測定) があれば、利用可能な corner で
    重みを正規化する (degraded interpolation)。
    Returns: (L, a, b) or None (全 corner 欠損時)
    """
    r, g, b = in_rgb
    r_lo, r_hi, tr = _bracket(r)
    g_lo, g_hi, tg = _bracket(g)
    b_lo, b_hi, tb = _bracket(b)

    corners = [
        ((r_lo, g_lo, b_lo), (1 - tr) * (1 - tg) * (1 - tb)),
        ((r_hi, g_lo, b_lo), tr * (1 - tg) * (1 - tb)),
        ((r_lo, g_hi, b_lo), (1 - tr) * tg * (1 - tb)),
        ((r_hi, g_hi, b_lo), tr * tg * (1 - tb)),
        ((r_lo, g_lo, b_hi), (1 - tr) * (1 - tg) * tb),
        ((r_hi, g_lo, b_hi), tr * (1 - tg) * tb),
        ((r_lo, g_hi, b_hi), (1 - tr) * tg * tb),
        ((r_hi, g_hi, b_hi), tr * tg * tb),
    ]

    total_w = 0.0
    L_sum = a_sum = b_sum = 0.0
    for corner, w in corners:
        if corner in lut and w > 0:
            lab = lut[corner]["out_lab"]
            L_sum += w * lab[0]
            a_sum += w * lab[1]
            b_sum += w * lab[2]
            total_w += w

    if total_w < 1e-9:
        return None

    return (L_sum / total_w, a_sum / total_w, b_sum / total_w)


# ── 逆引き (target → 最良入力) ──

def _lab_dist(lab1, lab2):
    """Lab Euclidean = ΔE_CIE76。"""
    return math.sqrt((lab1[0] - lab2[0]) ** 2 +
                     (lab1[1] - lab2[1]) ** 2 +
                     (lab1[2] - lab2[2]) ** 2)


def find_best_input(target_rgb, lut, coarse_step=8, fine_step=1, fine_radius=10):
    """target に最も近い出力を返す入力を探す (2 段階探索)。

    Phase 1: coarse_step ごとに全空間を粗く探索
    Phase 2: phase 1 の最良点の周辺 ±fine_radius を fine_step で詳細探索

    Returns: (best_input_tuple, predicted_output_rgb, predicted_de)
    """
    target_lab = rgb_to_lab(*target_rgb)

    best_de = float("inf")
    best_in = None
    best_lab = None

    # Phase 1: 粗探索
    for r in range(GAME_RGB_MIN, GAME_RGB_MAX + 1, coarse_step):
        for g in range(GAME_RGB_MIN, GAME_RGB_MAX + 1, coarse_step):
            for b in range(GAME_RGB_MIN, GAME_RGB_MAX + 1, coarse_step):
                if not is_reachable_input((r, g, b)):
                    continue
                pred_lab = trilinear_forward((r, g, b), lut)
                if pred_lab is None:
                    continue
                de = _lab_dist(pred_lab, target_lab)
                if de < best_de:
                    best_de = de
                    best_in = (r, g, b)
                    best_lab = pred_lab

    if best_in is None:
        return None, None, float("inf")

    # Phase 2: 詳細探索
    cr, cg, cb = best_in
    r_min = max(GAME_RGB_MIN, cr - fine_radius)
    r_max = min(GAME_RGB_MAX, cr + fine_radius)
    g_min = max(GAME_RGB_MIN, cg - fine_radius)
    g_max = min(GAME_RGB_MAX, cg + fine_radius)
    b_min = max(GAME_RGB_MIN, cb - fine_radius)
    b_max = min(GAME_RGB_MAX, cb + fine_radius)

    for r in range(r_min, r_max + 1, fine_step):
        for g in range(g_min, g_max + 1, fine_step):
            for b in range(b_min, b_max + 1, fine_step):
                if not is_reachable_input((r, g, b)):
                    continue
                pred_lab = trilinear_forward((r, g, b), lut)
                if pred_lab is None:
                    continue
                de = _lab_dist(pred_lab, target_lab)
                if de < best_de:
                    best_de = de
                    best_in = (r, g, b)
                    best_lab = pred_lab

    # 予測出力を RGB で再計算 (Lab → 表示用)
    predicted_out = _trilinear_forward_rgb(best_in, lut)
    return best_in, predicted_out, best_de


def _trilinear_forward_rgb(in_rgb, lut):
    """Trilinear 補間で RGB 直接版 (確認・表示用)。"""
    r, g, b = in_rgb
    r_lo, r_hi, tr = _bracket(r)
    g_lo, g_hi, tg = _bracket(g)
    b_lo, b_hi, tb = _bracket(b)

    corners = [
        ((r_lo, g_lo, b_lo), (1 - tr) * (1 - tg) * (1 - tb)),
        ((r_hi, g_lo, b_lo), tr * (1 - tg) * (1 - tb)),
        ((r_lo, g_hi, b_lo), (1 - tr) * tg * (1 - tb)),
        ((r_hi, g_hi, b_lo), tr * tg * (1 - tb)),
        ((r_lo, g_lo, b_hi), (1 - tr) * (1 - tg) * tb),
        ((r_hi, g_lo, b_hi), tr * (1 - tg) * tb),
        ((r_lo, g_hi, b_hi), (1 - tr) * tg * tb),
        ((r_hi, g_hi, b_hi), tr * tg * tb),
    ]

    total_w = 0.0
    R_sum = G_sum = B_sum = 0.0
    for corner, w in corners:
        if corner in lut and w > 0:
            o = lut[corner]["out"]
            R_sum += w * o[0]
            G_sum += w * o[1]
            B_sum += w * o[2]
            total_w += w

    if total_w < 1e-9:
        return None
    return (round(R_sum / total_w),
            round(G_sum / total_w),
            round(B_sum / total_w))
