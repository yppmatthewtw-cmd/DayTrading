"""
MACD 柱狀圖動能量化 (MACD Histogram Momentum Quantification)
============================================================

把「深綠 / 淺綠 / 淺紅 / 深紅」四色柱由視覺訊號轉為可計算的數值指標，
用於量化「動能衰竭 → 轉勢」的過程。

四色定義 (h = 當根柱值, h_prev = 前一根)
    深綠 DARK_GREEN : h > 0 且 h >= h_prev   上升動能「加速」
    淺綠 LIGHT_GREEN: h > 0 且 h <  h_prev   上升動能「衰減」  ← 頂部預警
    淺紅 LIGHT_PINK : h < 0 且 h >= h_prev   下跌動能「衰減」  ← 底部預警
    深紅 DARK_PINK  : h < 0 且 h <  h_prev   下跌動能「加速」

核心觀念：柱體 h = DIF - DEA，本身已是價格的一階動能差；
         顏色變化 = h 的一階差分 = 價格的二階導數。
         所以「深綠轉淺綠」比「柱體翻負」早，「柱體翻負」又比「均線交叉」早。
         愈早的訊號，假訊號率愈高 —— 故必須配合面積與背馳條件過濾。

用法
    python3 scripts/macd_histogram_momentum.py --demo
    python3 scripts/macd_histogram_momentum.py --csv bars.csv --fast 5 --slow 13 --signal 5
"""

from __future__ import annotations

import argparse
import numpy as np
import pandas as pd

# --------------------------------------------------------------------------
# 參數組合：日線用經典值，日內必須縮短，否則訊號慢到沒有交易價值
# --------------------------------------------------------------------------
PRESETS = {
    "daily":  dict(fast=12, slow=26, signal=9),   # 日線 / 週線
    "1h":     dict(fast=12, slow=26, signal=9),   # 小時線仍可用經典值
    "15min":  dict(fast=8,  slow=17, signal=9),   # 日內中速
    "5min":   dict(fast=5,  slow=13, signal=5),   # 日內快速
    "1min":   dict(fast=5,  slow=13, signal=5),   # 極短線 (雜訊極高)
}

DARK_GREEN, LIGHT_GREEN, LIGHT_PINK, DARK_PINK = "深綠", "淺綠", "淺紅", "深紅"

# 判定門檻 —— 這些數值是「起點」而非真理，必須用自己的標的回測校準
LAMBDA_WARN     = 0.50   # 峰值衰減率超過此值 = 動能過半耗盡
RUN_CONFIRM     = 3      # 連續幾根同色才算確認 (過濾單根雜訊)
DVG_SIGNIFICANT = 0.30   # 背馳強度門檻：面積較前波萎縮 30% 以上
MIN_WAVE_BARS   = 3      # 少於此根數的波視為雜訊，不參與背馳比較
MIN_AREA_RATIO  = 0.15   # 面積不足「同向波中位數」此比例者視為雜訊波
LATCH_BARS      = 10     # 證據鎖存期：條件觸發後仍計分的根數 (轉勢是過程，非單根)


# ==========================================================================
# 1. MACD 本體
# ==========================================================================
def macd(close: pd.Series, fast=12, slow=26, signal=9, hist_mult=2.0):
    """回傳 (DIF, DEA, HIST)。hist_mult=2 是中港台常見寫法，歐美多用 1。"""
    ema_f = close.ewm(span=fast, adjust=False).mean()
    ema_s = close.ewm(span=slow, adjust=False).mean()
    dif = ema_f - ema_s
    dea = dif.ewm(span=signal, adjust=False).mean()
    hist = (dif - dea) * hist_mult
    return dif, dea, hist


# ==========================================================================
# 2. 四色分類
# ==========================================================================
def classify_bars(hist: pd.Series) -> pd.Series:
    h, hp = hist, hist.shift(1)
    colour = pd.Series(index=hist.index, dtype=object)
    colour[(h > 0) & (h >= hp)] = DARK_GREEN
    colour[(h > 0) & (h <  hp)] = LIGHT_GREEN
    colour[(h <= 0) & (h <  hp)] = DARK_PINK
    colour[(h <= 0) & (h >= hp)] = LIGHT_PINK
    return colour


# ==========================================================================
# 3. 動能波段切分 —— 一個「波」= 柱體同號的連續區間
# ==========================================================================
def split_waves(hist: pd.Series, close: pd.Series, colour: pd.Series) -> pd.DataFrame:
    """把柱狀圖切成綠波 / 紅波，逐波計算動能量化指標。"""
    sign = np.sign(hist.fillna(0.0))
    wave_id = (sign != sign.shift(1)).cumsum()

    rows = []
    for wid, idx in hist.groupby(wave_id).groups.items():
        seg_h = hist.loc[idx].dropna()
        if len(seg_h) == 0 or seg_h.abs().sum() == 0:
            continue
        seg_p = close.loc[seg_h.index]
        seg_c = colour.loc[seg_h.index]
        is_green = seg_h.iloc[0] > 0

        area = float(seg_h.abs().sum())            # 動能總量 (面積)
        peak = float(seg_h.abs().max())            # 動能峰值
        peak_pos = int(seg_h.abs().to_numpy().argmax()) + 1   # 第幾根見峰
        n = len(seg_h)

        rows.append(dict(
            wave=int(wid),
            方向="綠波 (多)" if is_green else "紅波 (空)",
            起="" if len(seg_h) == 0 else str(seg_h.index[0]),
            訖=str(seg_h.index[-1]),
            柱數=n,
            面積M=area,
            面積M_pct=area / float(seg_p.mean()) * 100 if seg_p.mean() else np.nan,
            峰值H=peak,
            峰值位置=peak_pos,
            密度D=area / n,                          # 每根柱平均動能 = 急緩程度
            深色根數=int((seg_c == (DARK_GREEN if is_green else DARK_PINK)).sum()),
            淺色根數=int((seg_c == (LIGHT_GREEN if is_green else LIGHT_PINK)).sum()),
            價格起=float(seg_p.iloc[0]),
            價格訖=float(seg_p.iloc[-1]),
            價格極值=float(seg_p.max() if is_green else seg_p.min()),
        ))
    waves = pd.DataFrame(rows)
    if waves.empty:
        return waves
    # ---- 級別過濾：背馳只在「同級別」的波之間才有意義 ----------------
    # 一兩根柱的反彈不是一個波，拿它去和主升浪比面積會得出 99% 的假背馳。
    waves["級別"] = "雜訊波"
    for d in waves["方向"].unique():
        m = waves["方向"] == d
        big = waves.loc[m & (waves["柱數"] >= MIN_WAVE_BARS), "面積M"]
        thresh = big.median() * MIN_AREA_RATIO if len(big) else 0.0
        waves.loc[m & (waves["柱數"] >= MIN_WAVE_BARS)
                    & (waves["面積M"] >= thresh), "級別"] = "主波"
    return waves


# ==========================================================================
# 3b. 逐色動能總量 —— 「深綠有多少上升動能、深紅有多少下跌動能」的直接答案
# ==========================================================================
def colour_energy(hist: pd.Series, colour: pd.Series, close: pd.Series) -> pd.DataFrame:
    """把四色柱各自的動能加總。單位是價格單位，故同時給出對價格的百分比，
    使不同價位、不同標的之間可以互相比較。

    動能純度 = 深色動能 / (深色 + 淺色)
        > 0.65  趨勢健康，推動力持續
        0.45-0.65 正常
        < 0.45  多數時間在衰減 —— 是蠕動而非趨勢，不宜追
    """
    p = float(close.mean())
    rows = []
    for c in (DARK_GREEN, LIGHT_GREEN, LIGHT_PINK, DARK_PINK):
        m = colour == c
        e = float(hist[m].abs().sum())
        rows.append(dict(顏色=c, 根數=int(m.sum()), 動能總量=e,
                         動能_佔價格pct=e / p * 100 if p else np.nan,
                         平均每根=e / max(int(m.sum()), 1),
                         最大單根=float(hist[m].abs().max()) if m.any() else 0.0))
    df = pd.DataFrame(rows)
    dg, lg = df.loc[0, "動能總量"], df.loc[1, "動能總量"]
    lp, dp = df.loc[2, "動能總量"], df.loc[3, "動能總量"]
    purity = {
        "多方動能純度": dg / (dg + lg) if (dg + lg) else np.nan,
        "空方動能純度": dp / (dp + lp) if (dp + lp) else np.nan,
        "多空動能比": (dg + lg) / (dp + lp) if (dp + lp) else np.nan,
    }
    return df, purity


# ==========================================================================
# 4. 衰減率 λ —— 當前柱相對「本波至今峰值」已衰減多少
# ==========================================================================
def decay_lambda(hist: pd.Series) -> pd.Series:
    """λ = (峰值 - 現值) / 峰值 ∈ [0, 1]。λ=0 剛創動能新高；λ→1 動能耗盡。

    注意用 cummax 而非全波 max —— 實盤中未來的峰值是不可知的，
    只能用「至今為止的峰值」，否則就是未來函數 (look-ahead bias)。
    """
    sign = np.sign(hist.fillna(0.0))
    wave_id = (sign != sign.shift(1)).cumsum()
    absh = hist.abs()
    running_peak = absh.groupby(wave_id).cummax()
    lam = (running_peak - absh) / running_peak.replace(0, np.nan)
    return lam.clip(0, 1)


# ==========================================================================
# 5. 背馳量化 —— 轉勢訊號中唯一有統計意義的一項
# ==========================================================================
def divergence(waves: pd.DataFrame) -> pd.DataFrame:
    """比較同向的相鄰兩波：價格創新高/新低，動能面積卻萎縮 = 背馳。

    頂背馳強度 = 1 - A2/A1   (且 P2 > P1)
    底背馳強度 = 1 - A2/A1   (且 P2 < P1)
    """
    out = []
    if waves.empty:
        return pd.DataFrame(out)
    main = waves[waves["級別"] == "主波"]
    for direction in main["方向"].unique():
        sub = main[main["方向"] == direction].reset_index(drop=True)
        bullish = direction.startswith("綠")
        for i in range(1, len(sub)):
            p1, p2 = sub.loc[i - 1, "價格極值"], sub.loc[i, "價格極值"]
            a1, a2 = sub.loc[i - 1, "面積M"], sub.loc[i, "面積M"]
            h1, h2 = sub.loc[i - 1, "峰值H"], sub.loc[i, "峰值H"]
            price_new_extreme = (p2 > p1) if bullish else (p2 < p1)
            strength = 1 - (a2 / a1) if a1 else np.nan
            out.append(dict(
                前波=int(sub.loc[i - 1, "wave"]), 本波=int(sub.loc[i, "wave"]),
                方向=direction,
                價格創新極值="是" if price_new_extreme else "否",
                價格變動_pct=(p2 - p1) / p1 * 100 if p1 else np.nan,
                面積比A2_A1=a2 / a1 if a1 else np.nan,
                峰值比H2_H1=h2 / h1 if h1 else np.nan,
                背馳強度=strength if price_new_extreme else np.nan,
                判定=("顯著背馳" if (price_new_extreme and strength is not np.nan
                                    and strength > DVG_SIGNIFICANT)
                      else "輕微背馳" if (price_new_extreme and strength and strength > 0)
                      else "動能同步" if price_new_extreme else "未創新極值"),
            ))
    return pd.DataFrame(out)


# ==========================================================================
# 6. 轉勢綜合分 TRS (0-100)
# ==========================================================================
def _latch(flag, bars=LATCH_BARS):
    """證據鎖存。轉勢是一個過程而非一根柱：背馳發生在柱體翻負「之前」，
    破結構發生在「之後」，若只看當根，這些證據永不同時成立 —— 分數上限
    會被結構性地壓在 75 分，80 分那一級在數學上不可能出現。
    鎖存 N 根，讓證據沿著轉勢過程累積。"""
    return flag.fillna(False).astype(bool).rolling(bars, min_periods=1).max().astype(bool)



TRS_WEIGHTS = {
    "顏色由深轉淺":   20,
    "衰減率λ>0.5":    15,
    "連續淺色≥3根":   15,
    "背馳強度>0.3":   25,
    "柱體已過零軸":   15,
    "價格破前低/前高": 10,
}


def turn_score(df: pd.DataFrame, waves: pd.DataFrame, dvg: pd.DataFrame,
               top: bool = True) -> pd.DataFrame:
    """逐根計算頂部(top=True)或底部轉勢分數。"""
    colour = df["colour"]
    lam = df["lambda"]
    hist = df["hist"]
    close = df["close"]

    light = LIGHT_GREEN if top else LIGHT_PINK
    dark  = DARK_GREEN  if top else DARK_PINK

    # 連續淺色根數
    is_light = (colour == light)
    streak = is_light * (is_light.groupby((~is_light).cumsum()).cumcount() + 1)

    # 該波的背馳強度 (整波共用一個值)
    sign = np.sign(hist.fillna(0.0))
    wave_id = (sign != sign.shift(1)).cumsum()
    dvg_map = {}
    if not dvg.empty:
        want = "綠波 (多)" if top else "紅波 (空)"
        for _, r in dvg[dvg["方向"] == want].iterrows():
            dvg_map[r["本波"]] = r["背馳強度"]
    wave_dvg = wave_id.map(dvg_map).astype(float)

    # 價格結構：破前 20 根的低 (頂部) / 高 (底部)
    if top:
        broke = close < close.shift(1).rolling(20).min()
    else:
        broke = close > close.shift(1).rolling(20).max()

    s = pd.DataFrame(index=df.index)
    zero_cross = (hist <= 0) if top else (hist >= 0)
    s["c_顏色由深轉淺"] = _latch(is_light)                        * TRS_WEIGHTS["顏色由深轉淺"]
    s["c_衰減率"]   = _latch(lam.fillna(0) > LAMBDA_WARN)         * TRS_WEIGHTS["衰減率λ>0.5"]
    s["c_連續淺色"] = _latch(streak >= RUN_CONFIRM)               * TRS_WEIGHTS["連續淺色≥3根"]
    s["c_背馳"]     = _latch(wave_dvg.fillna(0) > DVG_SIGNIFICANT) * TRS_WEIGHTS["背馳強度>0.3"]
    s["c_過零軸"]   = zero_cross.fillna(False).astype(int)         * TRS_WEIGHTS["柱體已過零軸"]
    s["c_破結構"]   = _latch(broke)                               * TRS_WEIGHTS["價格破前低/前高"]

    s["TRS"] = s.sum(axis=1)

    # ---- 情境閘 ------------------------------------------------------
    # 「頂部轉勢」只有在確實處於升勢時才有意義。跌勢中每次深紅轉淺紅都會
    # 拿到分數，那不是轉勢訊號，是反彈。用中期均線定義情境，避免誤讀。
    # 用「近 N 根內曾處於該趨勢」而非「當下」，否則轉勢一旦完成、價格跌穿
    # 均線，分數立刻歸零 —— 恰好掐掉最高分的確認階段，永遠拿不到 80 分。
    ma = close.rolling(50, min_periods=20).mean()
    raw_ctx = (close > ma) if top else (close < ma)
    in_context = raw_ctx.rolling(15, min_periods=1).max().astype(bool)
    s["情境"] = np.where(in_context.fillna(False), "有效", "情境不符")
    s.loc[s["情境"] == "情境不符", "TRS"] = np.nan
    s["分級"] = pd.cut(s["TRS"], [-1, 35, 60, 80, 101],
                       labels=["趨勢延續", "動能衰竭·減倉", "轉勢確認中·離場", "已轉勢·可反手"])
    s["分級"] = s["分級"].cat.add_categories(["情境不符"]).fillna("情境不符")
    return s


# ==========================================================================
# 7. 一次過跑完
# ==========================================================================
def analyse(close: pd.Series, fast=12, slow=26, signal=9, hist_mult=2.0):
    dif, dea, hist = macd(close, fast, slow, signal, hist_mult)
    colour = classify_bars(hist)
    lam = decay_lambda(hist)
    df = pd.DataFrame(dict(close=close, dif=dif, dea=dea, hist=hist,
                           colour=colour, **{"lambda": lam}))
    waves = split_waves(hist, close, colour)
    energy, purity = colour_energy(hist, colour, close)
    dvg = divergence(waves)
    top_s = turn_score(df, waves, dvg, top=True)
    bot_s = turn_score(df, waves, dvg, top=False)
    df["TRS_頂"] = top_s["TRS"]
    df["TRS_頂_分級"] = top_s["分級"]
    df["TRS_底"] = bot_s["TRS"]
    df["TRS_底_分級"] = bot_s["分級"]
    return df, waves, dvg, energy, purity


# ==========================================================================
# Demo：合成資料驗證 —— 造一段「先強升、後背馳、再轉跌」的價格
# ==========================================================================
def demo_series(n=260, seed=7) -> pd.Series:
    rng = np.random.default_rng(seed)
    t = np.arange(n)
    # 第一波強升 → 回檔 → 第二波「價格更高但斜率變緩」(製造背馳) → 轉跌
    p = np.concatenate([
        100 + 0.55 * t[:70],                                  # 強升
        138.5 - 0.30 * np.arange(30),                         # 回檔
        129.5 + 0.22 * np.arange(80),                         # 弱升 (價創新高, 動能收縮)
        147.1 - 0.45 * np.arange(80),                         # 轉跌
    ])[:n]
    p = p + rng.normal(0, 0.45, len(p))
    idx = pd.date_range("2026-01-02", periods=len(p), freq="B")
    return pd.Series(p, index=idx, name="close")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--csv", help="含 close 欄的 CSV (第一欄為時間索引)")
    ap.add_argument("--preset", choices=list(PRESETS), default="daily")
    ap.add_argument("--fast", type=int); ap.add_argument("--slow", type=int)
    ap.add_argument("--signal", type=int)
    ap.add_argument("--hist-mult", type=float, default=2.0)
    ap.add_argument("--demo", action="store_true")
    ap.add_argument("--out", help="輸出 CSV")
    a = ap.parse_args()

    prm = dict(PRESETS[a.preset])
    for k in ("fast", "slow", "signal"):
        if getattr(a, k) is not None:
            prm[k] = getattr(a, k)

    if a.demo or not a.csv:
        close = demo_series()
        print(f"[demo] 合成 {len(close)} 根日線，參數 {prm}\n")
    else:
        raw = pd.read_csv(a.csv, index_col=0, parse_dates=True)
        col = "close" if "close" in raw else "Close"
        close = raw[col].astype(float)

    df, waves, dvg, energy, purity = analyse(close, hist_mult=a.hist_mult, **prm)

    pd.set_option("display.width", 200)
    print("=== 逐色動能總量 ===")
    print(energy.round(3).to_string(index=False))
    print(f"  多方動能純度 {purity['多方動能純度']:.1%}   "
          f"空方動能純度 {purity['空方動能純度']:.1%}   "
          f"多空動能比 {purity['多空動能比']:.2f}\n")

    pd.set_option("display.width", 200, "display.max_columns", 50)
    print("=== 動能波段 (每一波的量化指標) ===")
    print(waves[["wave", "方向", "級別", "柱數", "面積M", "面積M_pct", "峰值H",
                 "峰值位置", "密度D", "深色根數", "淺色根數", "價格極值"]]
          .round(3).to_string(index=False))
    print(f"  (主波 {int((waves['級別']=='主波').sum())} 個，"
          f"雜訊波 {int((waves['級別']=='雜訊波').sum())} 個已排除出背馳比較)")
    print("\n=== 背馳檢測 ===")
    print(dvg.round(3).to_string(index=False) if not dvg.empty else "(無)")

    fired = df[df["TRS_頂"] >= 60]
    print(f"\n=== 頂部轉勢分 TRS >= 60 的首次觸發 ===")
    if len(fired):
        f = fired.iloc[0]
        print(f"  時間 {fired.index[0].date()}  收市 {f['close']:.2f}  "
              f"TRS={f['TRS_頂']:.0f}  ({f['TRS_頂_分級']})")
        peak_i = df["close"].idxmax()
        print(f"  價格最高點 {peak_i.date()} @ {df['close'].max():.2f} "
              f"→ 訊號較高點 {'早' if fired.index[0] < peak_i else '遲'} "
              f"{abs((fired.index[0] - peak_i).days)} 個曆日")
    else:
        print("  (未觸發)")

    print("\n=== 最後 12 根明細 ===")
    print(df.tail(12)[["close", "hist", "colour", "lambda", "TRS_頂", "TRS_頂_分級"]]
          .round(3).to_string())

    if a.out:
        df.to_csv(a.out)
        print(f"\n已輸出 {a.out}")


if __name__ == "__main__":
    main()
