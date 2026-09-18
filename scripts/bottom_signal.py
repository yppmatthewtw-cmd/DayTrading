"""
深紅動能耗盡 → 第一見底訊號 (Bottom Signal from Dark-Red Exhaustion)
=====================================================================

核心前提：深紅柱 = h < 0 且 h 仍在變大（跌勢加速），
         它本身永遠不會是底部。要捉的是「深紅這一段耗盡的那一刻」。

訊號階梯（由早到遲，早的必然不準，這是無法迴避的取捨）：

    B1 深紅減速   |Δh| 由峰值收窄 ≥ decel_min，且仍是深紅
                  ← 唯一比「第一根淺紅」更早、而又可即時知道的訊號
    B2 首根淺紅   深紅 → 淺紅（跌勢由加速轉為減速）
    B3 淺紅確認   連續 k 根淺紅
    B4 柱體翻正   h > 0（DIF 升穿 DEA）
    B5 背馳確認   底背馳 + 價格升穿前高

關於「深紅動能峰值」的重要澄清：
    峰值只能事後知道。實盤中你察覺峰值已過的那一根，就是第一根淺紅。
    所以「計深紅峰值來捉底」和「等第一根淺紅」在即時交易中是同一時點。
    真正能再早一步的，是深紅柱的「增量」Δh —— 它有自己的峰值，
    且在柱體仍是深紅時就會開始收窄。這就是 B1。

要設的參數（全部可由自己的標的量出來，不必猜）：
    fast/slow/signal  MACD 週期，決定一段深紅通常有多長
    min_dark_run  N   深紅至少連續幾根，才視為一段「夠份量」的跌勢
    confirm_bars  k   需要幾根淺紅才確認
    decel_min         B1 的 |Δh| 收窄門檻
"""

from __future__ import annotations

import argparse
import numpy as np
import pandas as pd

from macd_histogram_momentum import (
    macd, classify_bars, decay_lambda, split_waves,
    DARK_PINK, LIGHT_PINK, PRESETS,
)

DEFAULT_MIN_DARK_RUN = None   # None = 自動取「該標的深紅段長度的中位數」
DEFAULT_CONFIRM_BARS = 2
DEFAULT_DECEL_MIN    = 0.40


# ==========================================================================
# 1. 深紅段的長度分佈 —— 「N 該設多少」的答案就在這裡，不必猜
# ==========================================================================
def dark_run_stats(colour: pd.Series, dark=DARK_PINK) -> dict:
    """量出這個標的 / 這個週期的深紅連續根數分佈。

    N 設在中位數附近，等於只對「長度高於一半」的跌段出手。
    設得比中位數高 → 訊號少而準而遲；低 → 多而吵而早。
    這是唯一有根據的設法：用標的自己的節奏，而不是抄別人的數字。
    """
    is_dark = (colour == dark)
    grp = (is_dark != is_dark.shift(1)).cumsum()
    runs = [int(len(g)) for k, g in is_dark.groupby(grp) if g.iloc[0]]
    if not runs:
        return dict(n=0)
    a = np.array(runs)
    return dict(n=len(a), 平均=float(a.mean()), 中位數=float(np.median(a)),
                P25=float(np.percentile(a, 25)), P75=float(np.percentile(a, 75)),
                P90=float(np.percentile(a, 90)), 最長=int(a.max()),
                分佈=np.bincount(a).tolist())


# ==========================================================================
# 2. 深紅減速度 —— B1 的量化基礎
# ==========================================================================
def dark_deceleration(hist: pd.Series, colour: pd.Series) -> pd.Series:
    """深紅段內，|Δh| 相對「本段至今 |Δh| 峰值」的收窄比例，∈[0,1]。

    h  的峰值告訴你「跌得最深的一刻」——事後才知道。
    Δh 的峰值告訴你「跌勢加速得最猛的一刻」——柱體仍是深紅時就能算。
    這就是 B1 能比 B2 早的全部理由。用 cummax，不偷看未來。
    """
    d = hist.diff()
    is_dark = (colour == DARK_PINK)
    seg = (is_dark != is_dark.shift(1)).cumsum()
    mag = d.abs().where(is_dark)
    peak = mag.groupby(seg).cummax()
    decel = (peak - mag) / peak.replace(0, np.nan)
    return decel.clip(0, 1)


# ==========================================================================
# 3. 五級底部訊號
# ==========================================================================
def bottom_signals(close: pd.Series, fast=12, slow=26, signal=9, hist_mult=2.0,
                   min_dark_run=DEFAULT_MIN_DARK_RUN,
                   confirm_bars=DEFAULT_CONFIRM_BARS,
                   decel_min=DEFAULT_DECEL_MIN) -> tuple[pd.DataFrame, dict]:
    dif, dea, hist = macd(close, fast, slow, signal, hist_mult)
    colour = classify_bars(hist)
    stats = dark_run_stats(colour)

    if min_dark_run is None:
        min_dark_run = max(2, int(round(stats.get("中位數", 3))))

    is_dark = (colour == DARK_PINK)
    is_light = (colour == LIGHT_PINK)

    # 目前這一段深紅已經走了幾根
    dark_seg = (is_dark != is_dark.shift(1)).cumsum()
    dark_len = is_dark.groupby(dark_seg).cumcount().add(1).where(is_dark, 0)

    # 上一段深紅的長度（淺紅出現時，用來判斷「之前那段跌勢夠不夠份量」）
    prev_dark_len = dark_len.replace(0, np.nan).ffill().fillna(0)

    light_seg = (is_light != is_light.shift(1)).cumsum()
    light_len = is_light.groupby(light_seg).cumcount().add(1).where(is_light, 0)

    decel = dark_deceleration(hist, colour)
    qualified = prev_dark_len >= min_dark_run   # 跌段夠長，才值得找底

    df = pd.DataFrame(dict(
        close=close, hist=hist, colour=colour,
        深紅根數=dark_len.astype(int), 淺紅根數=light_len.astype(int),
        深紅減速度=decel, lam=decay_lambda(hist),
    ))
    # 每一級訊號每段只響一次。用 >= 會在整段淺紅上逐根重複觸發，
    # 觸發次數被灌水近十倍，與其他級別根本無法比較。
    b1_raw = is_dark & (dark_len >= min_dark_run) & (decel.fillna(0) >= decel_min)
    df["B1_深紅減速"] = b1_raw & ~b1_raw.shift(1, fill_value=False)
    df["B2_首根淺紅"] = is_light & (light_len == 1) & qualified
    df["B3_淺紅確認"] = is_light & (light_len == confirm_bars) & qualified
    df["B4_柱體翻正"] = (hist > 0) & (hist.shift(1) <= 0) & qualified
    prev_high = close.shift(1).rolling(20).max()
    above = close > prev_high
    df["B5_破結構"]  = above & ~above.shift(1, fill_value=False) & qualified

    params = dict(fast=fast, slow=slow, signal=signal, min_dark_run=int(min_dark_run),
                  confirm_bars=confirm_bars, decel_min=decel_min, **stats)
    return df, params


# ==========================================================================
# 4. 評估：每一級訊號距真正底部多少根、命中率多少
# ==========================================================================
def find_swing_lows(close: pd.Series, window=10) -> pd.Series:
    """真正的底：前後 window 根之內的最低點。這是事後標準答案，只用於評估。"""
    roll_min = close.rolling(window * 2 + 1, center=True).min()
    return close == roll_min


def evaluate(df: pd.DataFrame, close: pd.Series, fwd_bars=20, window=10) -> pd.DataFrame:
    lows = find_swing_lows(close, window)
    low_pos = np.where(lows.fillna(False).to_numpy())[0]
    fwd = close.shift(-fwd_bars) / close - 1
    base = float(fwd.mean())

    out = []
    for col in [c for c in df.columns if c.startswith("B")]:
        m = df[col].fillna(False).to_numpy()
        idx = np.where(m)[0]
        if len(idx) == 0:
            out.append(dict(訊號=col, 觸發次數=0)); continue
        if len(low_pos):
            # 每個訊號到最近真底的距離（負=早於底，正=遲於底）
            dist = np.array([idx_i - low_pos[np.abs(low_pos - idx_i).argmin()]
                             for idx_i in idx])
        else:
            dist = np.array([np.nan])
        f = fwd.to_numpy()[idx]
        f = f[~np.isnan(f)]
        out.append(dict(
            訊號=col, 觸發次數=int(len(idx)), 觸發率=len(idx) / len(df),
            距真底_中位數根=float(np.median(dist)),
            距真底_絕對值中位數=float(np.median(np.abs(dist))),
            早於底比例=float((dist < 0).mean()),
            其後報酬=float(f.mean()) if len(f) else np.nan,
            上漲命中率=float((f > 0).mean()) if len(f) else np.nan,
        ))
    d = pd.DataFrame(out)
    d.attrs["基準報酬"] = base
    return d


# ==========================================================================
# 5. 參數掃描 —— 直接看見「早 vs 準」的取捨曲線
# ==========================================================================
def sweep(close: pd.Series, preset="daily", n_range=range(2, 9),
          decel_range=(0.2, 0.3, 0.4, 0.5, 0.6), fwd_bars=20, window=10):
    prm = PRESETS[preset]
    lows = find_swing_lows(close, window)
    low_pos = np.where(lows.fillna(False).to_numpy())[0]
    fwd = close.shift(-fwd_bars) / close - 1
    rows = []
    for N in n_range:
        for dm in decel_range:
            df, _ = bottom_signals(close, min_dark_run=N, decel_min=dm, **prm)
            for col in ("B1_深紅減速", "B2_首根淺紅", "B3_淺紅確認"):
                idx = np.where(df[col].fillna(False).to_numpy())[0]
                if len(idx) == 0 or not len(low_pos):
                    continue
                dist = np.array([i - low_pos[np.abs(low_pos - i).argmin()] for i in idx])
                f = fwd.to_numpy()[idx]; f = f[~np.isnan(f)]
                rows.append(dict(N=N, decel_min=dm, 訊號=col, 次數=len(idx),
                                 距底中位數=float(np.median(dist)),
                                 誤差中位數=float(np.median(np.abs(dist))),
                                 其後報酬=float(f.mean()) if len(f) else np.nan,
                                 命中率=float((f > 0).mean()) if len(f) else np.nan))
    return pd.DataFrame(rows)


# ==========================================================================
# Demo 資料：多個底部，供參數校準
# ==========================================================================
def multi_cycle(n=600, seed=11) -> pd.Series:
    """造出多個高低點的序列，讓底部訊號有足夠樣本可評估。"""
    rng = np.random.default_rng(seed)
    t = np.arange(n)
    trend = 100 + 0.02 * t
    cyc = 9 * np.sin(2 * np.pi * t / 84) + 4.5 * np.sin(2 * np.pi * t / 31 + 1.1)
    noise = np.cumsum(rng.normal(0, 0.30, n))
    p = trend + cyc + noise
    idx = pd.date_range("2024-01-01", periods=n, freq="B")
    return pd.Series(p, index=idx, name="close")


def calibrate(close: pd.Series, preset="daily", window=10, fwd_bars=20) -> pd.DataFrame:
    """量出這個標的的專屬參數，不必猜。

    掃不同 MACD 週期，回報每組的提前量與誤差。選哪一組，取決於
    你要的是「早」還是「準」—— 因為在這裡兩者是同一個數字。
    """
    lows = np.where(find_swing_lows(close, window).fillna(False).to_numpy())[0]
    grids = {
        "daily": [(5,13,5),(8,17,9),(12,26,9),(19,39,9),(24,52,18)],
        "5min":  [(3,8,3),(5,13,5),(8,17,9),(12,26,9)],
        "15min": [(5,13,5),(8,17,9),(12,26,9),(19,39,9)],
    }
    rows = []
    for f, sl, g in grids.get(preset, grids["daily"]):
        df, prm = bottom_signals(close, fast=f, slow=sl, signal=g)
        for col in ("B1_深紅減速", "B2_首根淺紅", "B3_淺紅確認"):
            idx = np.where(df[col].fillna(False).to_numpy())[0]
            if len(idx) == 0 or not len(lows):
                continue
            d = np.array([i - lows[np.abs(lows - i).argmin()] for i in idx])
            fw = (close.shift(-fwd_bars)/close - 1).to_numpy()[idx]
            fw = fw[~np.isnan(fw)]
            rows.append(dict(MACD=f"{f}/{sl}/{g}", 訊號=col,
                             深紅段中位長=prm["中位數"], 建議N=prm["min_dark_run"],
                             觸發次數=len(idx), 距底中位數=float(np.median(d)),
                             誤差中位數=float(np.median(np.abs(d))),
                             其後報酬=float(fw.mean()) if len(fw) else np.nan,
                             命中率=float((fw > 0).mean()) if len(fw) else np.nan))
    return pd.DataFrame(rows)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--csv"); ap.add_argument("--preset", choices=list(PRESETS), default="daily")
    ap.add_argument("--min-dark-run", type=int); ap.add_argument("--confirm-bars", type=int, default=2)
    ap.add_argument("--decel-min", type=float, default=DEFAULT_DECEL_MIN)
    ap.add_argument("--sweep", action="store_true")
    ap.add_argument("--calibrate", action="store_true",
                    help="掃 MACD 週期，量出這個標的的專屬提前量與誤差")
    a = ap.parse_args()

    if a.csv:
        raw = pd.read_csv(a.csv, index_col=0, parse_dates=True)
        close = raw["close" if "close" in raw else "Close"].astype(float)
    else:
        close = multi_cycle()
        print(f"[demo] 合成 {len(close)} 根，多週期序列\n")

    prm = PRESETS[a.preset]
    df, params = bottom_signals(close, min_dark_run=a.min_dark_run,
                                confirm_bars=a.confirm_bars, decel_min=a.decel_min, **prm)
    pd.set_option("display.width", 220, "display.max_columns", 40)

    print("=== 深紅段長度分佈（N 該設多少，答案在這裡）===")
    for k in ("n", "平均", "中位數", "P25", "P75", "P90", "最長"):
        print(f"  {k:6s} {params[k]}")
    print(f"  → 自動採用 min_dark_run = {params['min_dark_run']}\n")

    ev = evaluate(df, close)
    print("=== 五級訊號評估（距真底：負=早於底，正=遲於底）===")
    print(ev.round(3).to_string(index=False))
    print(f"  全樣本 20 根基準報酬 {ev.attrs['基準報酬']:+.2%}")

    if a.calibrate:
        print("\n=== 參數校準：選哪一組，就是選你要多早 ===")
        print(calibrate(close, a.preset).round(3).to_string(index=False))

    if a.sweep:
        print("\n=== 參數掃描 ===")
        sw = sweep(close, a.preset)
        print(sw.round(3).to_string(index=False))


if __name__ == "__main__":
    main()
