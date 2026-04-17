"""
DyeForge - 3D LUT 校正スクリプト

8×8×8 グリッドのうち、color_matcher.py の _clamp_rgb を通過する組み合わせ
(per-channel [0x1A, 0xDA] かつ S <= 78%) のみを測定して 3D LUT を作成する。
512 候補 → 実測 380 点。チャネル間相互作用を含む完全なゲーム応答テーブル。

特徴:
  - 中断 (Ctrl+C, クラッシュ等) しても次回起動で続きから再開
  - 20 点ごとに自動保存
  - 開始時に既存ファイルを検出して再開を提案

手順:
  1. ゲームの染色画面を開く
  2. 管理者権限で実行: python lut_3d_calibration.py
  3. HEX 入力欄をクリック → 染色プレビューをクリック
  4. 妥当性自動判定 → OK で続行
  5. 約 18-22 分待つ (中断 OK)
  6. lut_3d.json (LUT 本体) と lut_3d.csv (確認用) が保存される

注意:
  - ゲームウィンドウを動かさない
  - サンプル位置の材質を変えない
  - PC スリープも避ける (HEX 入力が伝わらなくなる)
  - マウスは触らないこと (script がクリック操作する)
"""

import ctypes
import ctypes.wintypes
import csv
import json
import os
import sys
import time

from color_matcher import (
    click_at, type_hex_into_field,
    user32, gdi32,
    delta_e,
)


# ── 設定 ────────────────────────────────────────────────────────

# 8 段階 in [GAME_RGB_MIN, GAME_RGB_MAX]。等間隔。
GRID_LEVELS = [26, 53, 80, 107, 134, 161, 188, 218]
# color_matcher.py._clamp_rgb と同じ S 上限 (これを超える組み合わせはアプリが送れない)
GAME_S_MAX_PCT = 78.0

SAMPLES_PER_POINT = 5
SAMPLE_INTERVAL = 0.18
SETTLE_DELAY = 0.7
SAVE_INTERVAL = 20      # この点数ごとに途中保存
RECHECK_INTERVAL = 100  # この点数ごとに位置妥当性を再確認

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_JSON = os.path.join(SCRIPT_DIR, "lut_3d.json")
OUTPUT_CSV = os.path.join(SCRIPT_DIR, "lut_3d.csv")


# ── 共通プリミティブ ──────────────────────────────────────────

def capture_screen_pixel(x, y):
    screen_dc = user32.GetDC(0)
    color_ref = gdi32.GetPixel(screen_dc, int(x), int(y))
    user32.ReleaseDC(0, screen_dc)
    if color_ref == -1:
        return None
    return (color_ref & 0xFF, (color_ref >> 8) & 0xFF, (color_ref >> 16) & 0xFF)


def wait_for_click(prompt):
    VK_LBUTTON = 0x01
    print(prompt)
    while user32.GetAsyncKeyState(VK_LBUTTON) & 0x8000:
        time.sleep(0.01)
    while not (user32.GetAsyncKeyState(VK_LBUTTON) & 0x8000):
        time.sleep(0.01)
    pt = ctypes.wintypes.POINT()
    user32.GetCursorPos(ctypes.byref(pt))
    while user32.GetAsyncKeyState(VK_LBUTTON) & 0x8000:
        time.sleep(0.01)
    time.sleep(0.1)
    return (pt.x, pt.y)


def measure(hex_pos, sample_pos, in_rgb):
    """HEX を入力し、安定後に N サンプルの中央値を返す。"""
    r, g, b = in_rgb
    hex_code = f"{r:02X}{g:02X}{b:02X}"
    type_hex_into_field(hex_pos[0], hex_pos[1], hex_code)
    time.sleep(SETTLE_DELAY)

    samples = []
    for _ in range(SAMPLES_PER_POINT):
        s = capture_screen_pixel(*sample_pos)
        if s is not None:
            samples.append(s)
        time.sleep(SAMPLE_INTERVAL)

    if not samples:
        return None, []

    rs = sorted(s[0] for s in samples)
    gs = sorted(s[1] for s in samples)
    bs = sorted(s[2] for s in samples)
    mid = len(samples) // 2
    return (rs[mid], gs[mid], bs[mid]), samples


def hex_key(rgb):
    return f"{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"


def grid_points():
    """8×8×8 = 512 候補のうち、S<=78% の到達可能な組み合わせのみ (380 点)。
    color_matcher.py の _clamp_rgb は S>78% を 78% に丸めるので、
    raw のまま送っても LUT 上で別エントリとして意味がない。
    """
    points = []
    for r in GRID_LEVELS:
        for g in GRID_LEVELS:
            for b in GRID_LEVELS:
                mx, mn = max(r, g, b), min(r, g, b)
                s_pct = ((mx - mn) / mx * 100) if mx > 0 else 0
                if s_pct <= GAME_S_MAX_PCT:
                    points.append((r, g, b))
    return points


# ── 妥当性判定 (gamma_calibration.py と共通ロジック) ────────────

def validate_position(hex_pos, sample_pos):
    """2 色送って妥当性を自動判定。Returns (ok, message, samples)。"""
    out1, _ = measure(hex_pos, sample_pos, (128, 128, 128))
    if out1 is None:
        return False, "サンプル取得失敗 (#808080)", []
    out2, _ = measure(hex_pos, sample_pos, (64, 26, 26))
    if out2 is None:
        return False, "サンプル取得失敗 (#401A1A)", [out1]

    de = delta_e(out1, out2)
    spread = max(out1) - min(out1)
    issues = []
    if spread > 30:
        issues.append(f"#808080 が灰色でない (max-min={spread})")
    if max(out1) < 80:
        issues.append(f"#808080 が暗すぎ ({out1})")
    if min(out1) > 240:
        issues.append(f"#808080 が明るすぎ ({out1})")
    if de < 5:
        issues.append(f"反応性なし (ΔE={de:.1f})")

    msg = (f"#808080 → ({out1[0]},{out1[1]},{out1[2]}), "
           f"#401A1A → ({out2[0]},{out2[1]},{out2[2]}), ΔE={de:.2f}")
    return (len(issues) == 0), msg + (" | " + " / ".join(issues) if issues else ""), [out1, out2]


# ── LUT データ管理 (resume 対応) ──────────────────────────────

def load_existing():
    if not os.path.exists(OUTPUT_JSON):
        return None
    try:
        with open(OUTPUT_JSON, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        print(f"  ⚠ 既存ファイル読み込み失敗: {e}")
        return None


def save_progress(data):
    """atomic write: 一時ファイルに書いてから rename。"""
    tmp = OUTPUT_JSON + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp, OUTPUT_JSON)


def export_csv(data):
    with open(OUTPUT_CSV, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["hex_in", "in_R", "in_G", "in_B", "out_R", "out_G", "out_B"])
        for k in sorted(data["results"].keys()):
            v = data["results"][k]
            ir, ig, ib = v["in"]
            o = v.get("out") or [None, None, None]
            w.writerow([k, ir, ig, ib, o[0], o[1], o[2]])


# ── メイン ────────────────────────────────────────────────────

def main():
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        pass

    try:
        is_admin = bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        is_admin = False
    if not is_admin:
        print("⚠ 管理者権限ではありません。Enter で続行 / Ctrl+C で中止")
        try:
            input()
        except KeyboardInterrupt:
            return

    print("=" * 70)
    print("  DyeForge 3D LUT 校正スクリプト")
    print(f"  グリッド {len(GRID_LEVELS)}^3 = {len(GRID_LEVELS)**3} 点")
    print("=" * 70)

    grid = grid_points()
    n_total = len(grid)

    # ── 既存ファイル検出 (resume) ──
    existing = load_existing()
    use_existing = False
    if existing and existing.get("results"):
        n_done = len(existing["results"])
        print(f"\n既存ファイルが見つかりました: {n_done}/{n_total} 点完了")
        try:
            ans = input("続きから再開しますか？ [Y/n]: ").strip().lower()
        except KeyboardInterrupt:
            return
        if ans != "n":
            use_existing = True

    if use_existing:
        hex_pos = tuple(existing["hex_pos"])
        sample_pos = tuple(existing["sample_pos"])
        results = existing["results"]
        print(f"  HEX 位置 (前回): {hex_pos}")
        print(f"  サンプル位置 (前回): {sample_pos}")
        # 再開時に位置がまだ妥当か確認
        print("\n位置妥当性を再チェック中...")
        ok, msg, _ = validate_position(hex_pos, sample_pos)
        print(f"  {msg}")
        if not ok:
            print("  ⚠ 異常検知。前回と環境が変わっている可能性があります。")
            try:
                ans = input("  それでも続行? [y/N]: ").strip().lower()
            except KeyboardInterrupt:
                return
            if ans != "y":
                return
    else:
        hex_pos = wait_for_click("\n→ ゲームの HEX 入力欄をクリックしてください")
        print(f"  HEX 位置: {hex_pos}")
        sample_pos = wait_for_click("\n→ 染色プレビューの単色部分をクリックしてください")
        print(f"  サンプル位置: {sample_pos}\n")
        results = {}

        print("位置確認中 (#808080 と #401A1A を送って判定)...")
        ok, msg, _ = validate_position(hex_pos, sample_pos)
        print(f"  {msg}")
        if not ok:
            try:
                ans = input("\n  それでも続行? [y/N]: ").strip().lower()
            except KeyboardInterrupt:
                return
            if ans != "y":
                return
        else:
            print("  ✓ 妥当な位置です")

    # ── 残り点数を計算 ──
    remaining = [p for p in grid if hex_key(p) not in results]
    n_remaining = len(remaining)
    n_done_initial = n_total - n_remaining

    if n_remaining == 0:
        print("\n全 512 点が既に測定済みです。")
        print("CSV を再生成するなら lut_3d.json を削除してから再実行してください。")
        export_csv({"results": results})
        return

    est_per_point = SETTLE_DELAY + SAMPLES_PER_POINT * SAMPLE_INTERVAL + 1.0
    est_time = n_remaining * est_per_point
    print(f"\n残り {n_remaining} 点を測定 (推定 {est_time/60:.1f} 分)")
    print(f"中断は Ctrl+C (途中保存あり、次回再開可能)\n")

    data = {
        "version": 1,
        "hex_pos": list(hex_pos),
        "sample_pos": list(sample_pos),
        "grid_levels": GRID_LEVELS,
        "samples_per_point": SAMPLES_PER_POINT,
        "settle_delay": SETTLE_DELAY,
        "results": results,
    }

    start_time = time.time()
    last_save_count = n_done_initial
    last_recheck_count = n_done_initial
    failed = 0

    try:
        for i, in_rgb in enumerate(remaining, 1):
            global_idx = n_done_initial + i
            out, samples = measure(hex_pos, sample_pos, in_rgb)
            key = hex_key(in_rgb)

            if out:
                results[key] = {
                    "in": list(in_rgb),
                    "out": list(out),
                    "samples": [list(s) for s in samples],
                }
                out_str = f"({out[0]:3d},{out[1]:3d},{out[2]:3d})"
            else:
                out_str = "FAILED"
                failed += 1

            # ETA 計算
            elapsed = time.time() - start_time
            avg = elapsed / i
            eta = avg * (n_remaining - i)
            print(f"  [{global_idx:3d}/{n_total}] in ({in_rgb[0]:3d},{in_rgb[1]:3d},{in_rgb[2]:3d}) "
                  f"→ {out_str}  ETA {eta/60:.1f}min")

            # 途中保存
            if global_idx - last_save_count >= SAVE_INTERVAL:
                save_progress(data)
                last_save_count = global_idx

            # 定期妥当性チェック (位置がずれてないか)
            if global_idx - last_recheck_count >= RECHECK_INTERVAL and global_idx < n_total:
                print("  ─ 位置妥当性を再チェック中... ", end="", flush=True)
                check_out, _ = measure(hex_pos, sample_pos, (128, 128, 128))
                if check_out:
                    spread = max(check_out) - min(check_out)
                    if spread > 30 or max(check_out) < 80:
                        print(f"⚠ 異常 #808080→{check_out}")
                        try:
                            ans = input("  続行? [Y/n]: ").strip().lower()
                            if ans == "n":
                                break
                        except KeyboardInterrupt:
                            break
                    else:
                        print(f"OK #808080→{check_out}")
                else:
                    print("× サンプル失敗")
                last_recheck_count = global_idx

    except KeyboardInterrupt:
        print("\n中断されました。これまでの結果を保存します...")

    save_progress(data)
    export_csv(data)

    n_done = len(results)
    print(f"\n{'=' * 70}")
    print(f"  保存完了: {n_done}/{n_total} 点 ({n_done/n_total*100:.1f}%)")
    if failed:
        print(f"  ⚠ FAILED: {failed} 点")
    print(f"  {OUTPUT_JSON}")
    print(f"  {OUTPUT_CSV}")
    print(f"{'=' * 70}")
    if n_done < n_total:
        print(f"\n再実行すれば残り {n_total - n_done} 点から再開できます。")


if __name__ == "__main__":
    main()
