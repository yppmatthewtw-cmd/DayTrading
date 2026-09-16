# -*- coding: utf-8 -*-
"""
高交額 20 大 → 篩選「波幅小 + 20 分鐘均線平穩向上」的個股。

本環境的行情 API 被網絡政策封鎖，故此腳本設計為在有行情權限的環境執行。

用法:
    pip install yfinance pandas numpy
    python3 screen_20min_ma.py --date 2026-09-08
    python3 screen_20min_ma.py --date 2026-09-08 --universe AAPL,MSFT,NVDA,...

注意: yfinance 的 1 分鐘資料通常只保留最近 30 天。若要回溯更早的日期，
請改接 Polygon / Databento / IBKR 等有歷史 intraday 的資料源，
只需替換 fetch_intraday() 與 fetch_daily() 兩個函式即可，其餘邏輯不變。
"""
import argparse
import sys

import numpy as np
import pandas as pd

# 預設 universe: 常態美元成交額最大的美股 (可用 --universe 覆寫)
DEFAULT_UNIVERSE = [
    "NVDA", "TSLA", "AAPL", "MSFT", "AMZN", "META", "GOOGL", "AVGO", "AMD", "MU",
    "INTC", "ORCL", "PLTR", "QCOM", "NFLX", "CRWV", "IREN", "SOFI", "PFE", "BE",
    "SMCI", "COIN", "MARA", "BAC", "F", "T", "XOM", "JPM", "WMT", "LLY",
]

# 篩選門檻
RISING_RATIO_MIN = 0.70   # 20 分鐘均線的一階差分為正的比例
R2_MIN           = 0.60   # 線性迴歸 R^2: 越高代表路徑越接近直線 (平穩)
SLOPE_MIN        = 0.0    # 標準化斜率須為正 (向上)
MIN_MA_POINTS    = 10     # 均線至少要有這麼多個點，判斷才有意義


def fetch_daily(tickers, date):
    """回傳當日的 close / volume / high / low，用於計算成交額與日內波幅。"""
    import yfinance as yf
    start = pd.Timestamp(date)
    end = start + pd.Timedelta(days=1)
    df = yf.download(tickers, start=start, end=end, interval="1d",
                     group_by="ticker", auto_adjust=False, progress=False)
    out = {}
    for t in tickers:
        try:
            sub = df[t] if isinstance(df.columns, pd.MultiIndex) else df
            row = sub.dropna().iloc[-1]
            out[t] = dict(close=float(row["Close"]), volume=float(row["Volume"]),
                          high=float(row["High"]), low=float(row["Low"]),
                          open=float(row["Open"]))
        except Exception:
            continue
    return out


def fetch_intraday(ticker, date):
    """回傳當日常規交易時段 (09:30-16:00 ET) 的 1 分鐘 K 線。"""
    import yfinance as yf
    start = pd.Timestamp(date)
    end = start + pd.Timedelta(days=1)
    df = yf.download(ticker, start=start, end=end, interval="1m",
                     auto_adjust=False, progress=False)
    if df.empty:
        return df
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = df.columns.get_level_values(0)
    idx = df.index
    if idx.tz is None:
        idx = idx.tz_localize("America/New_York")
    else:
        idx = idx.tz_convert("America/New_York")
    df.index = idx
    return df.between_time("09:30", "16:00")


def ma_quality(intraday, bar="20min", ma_window=20):
    """
    將 1 分鐘 K 線重採樣為指定週期的 K 線，計算移動平均，
    再量化該均線的「向上程度」與「平穩程度」。
    """
    if intraday is None or intraday.empty:
        return None

    bars = intraday["Close"].resample(bar).last().dropna()
    if len(bars) < 5:
        return None

    # 一個交易日只有 ~19.5 根 20 分鐘 K，故均線窗口自動收斂到可用長度
    win = min(ma_window, max(3, len(bars) // 2))
    ma = bars.rolling(win).mean().dropna()
    if len(ma) < 4:
        return None

    diffs = ma.diff().dropna()
    rising_ratio = float((diffs > 0).mean())

    # 對均線做線性迴歸: 標準化斜率 (以每根 K 的百分比表示) 與 R^2
    x = np.arange(len(ma), dtype=float)
    y = ma.to_numpy(dtype=float)
    slope, intercept = np.polyfit(x, y, 1)
    fitted = slope * x + intercept
    ss_res = float(((y - fitted) ** 2).sum())
    ss_tot = float(((y - y.mean()) ** 2).sum())
    r2 = 1.0 - ss_res / ss_tot if ss_tot > 0 else 0.0
    slope_pct = float(slope / y.mean()) if y.mean() else 0.0

    # 波幅: 20 分鐘收益率的標準差
    rets = bars.pct_change().dropna()
    vol_20m = float(rets.std()) if len(rets) > 1 else np.nan

    return dict(bars=len(bars), ma_points=len(ma), rising_ratio=rising_ratio,
                slope_pct_per_bar=slope_pct, r2=float(r2), vol_bar=vol_20m)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--date", required=True, help="交易日 YYYY-MM-DD")
    ap.add_argument("--universe", default=None, help="逗號分隔的代號清單")
    ap.add_argument("--top", type=int, default=20, help="取成交額前 N 名")
    ap.add_argument("--bar", default="20min",
                    help="重採樣週期，例如 10min / 20min / 30min（預設 20min）")
    ap.add_argument("--out", default=None, help="輸出 CSV 路徑")
    args = ap.parse_args()

    universe = ([t.strip().upper() for t in args.universe.split(",")]
                if args.universe else DEFAULT_UNIVERSE)

    print(f"[1/3] 取 {args.date} 日線資料，計算成交額 …", file=sys.stderr)
    daily = fetch_daily(universe, args.date)
    if not daily:
        sys.exit("取不到日線資料 — 請檢查日期是否為交易日，以及行情來源是否可用。")

    turnover = sorted(
        ((t, d["close"] * d["volume"], d) for t, d in daily.items()),
        key=lambda x: x[1], reverse=True,
    )[: args.top]

    print(f"[2/3] 取前 {len(turnover)} 名的 1 分鐘 K 線並計算 {args.bar} 均線 …", file=sys.stderr)
    recs = []
    for rank, (t, tv, d) in enumerate(turnover, 1):
        q = ma_quality(fetch_intraday(t, args.date), bar=args.bar)
        day_range = (d["high"] - d["low"]) / d["close"] if d["close"] else np.nan
        rec = dict(rank=rank, ticker=t, close=d["close"],
                   chg_pct=(d["close"] / d["open"] - 1) if d["open"] else np.nan,
                   turnover_usd=tv, day_range_pct=day_range)
        rec.update(q or dict(rising_ratio=np.nan, slope_pct_per_bar=np.nan,
                             r2=np.nan, vol_bar=np.nan))
        recs.append(rec)

    df = pd.DataFrame(recs)

    # 篩選: 平穩向上
    df["passes"] = (
        (df["rising_ratio"] >= RISING_RATIO_MIN)
        & (df["slope_pct_per_bar"] > SLOPE_MIN)
        & (df["r2"] >= R2_MIN)
    )
    # 在通過者之中，以波幅由小至大排序
    df = df.sort_values(["passes", "day_range_pct"], ascending=[False, True])

    print(f"[3/3] 完成。\n", file=sys.stderr)
    cols = ["rank", "ticker", "close", "chg_pct", "turnover_usd", "day_range_pct",
            "vol_bar", "rising_ratio", "slope_pct_per_bar", "r2", "passes"]
    with pd.option_context("display.width", 200, "display.max_columns", 50):
        print(df[cols].to_string(index=False,
              float_format=lambda v: f"{v:,.4f}"))

    need = pd.Timedelta(args.bar) * MIN_MA_POINTS
    print(f"\n(判定門檻: rising_ratio>={RISING_RATIO_MIN}, slope>0, R2>={R2_MIN}。"
          f"{args.bar} 週期至少需要 {need} 的盤中資料才有意義。)", file=sys.stderr)

    hits = df[df["passes"]]
    print("\n=== 符合『波幅小 + 20 分鐘均線平穩向上』 ===")
    if hits.empty:
        print("無。當日沒有個股同時滿足三個門檻 "
              f"(rising_ratio>={RISING_RATIO_MIN}, slope>0, R2>={R2_MIN})。")
    else:
        for _, r in hits.iterrows():
            print(f"  {r['ticker']:<6} 日內波幅 {r['day_range_pct']:.2%}  "
                  f"均線上升比例 {r['rising_ratio']:.0%}  R²={r['r2']:.2f}  "
                  f"斜率 {r['slope_pct_per_bar']:+.3%}/根")

    if args.out:
        df.to_csv(args.out, index=False)
        print(f"\n已寫出: {args.out}")


if __name__ == "__main__":
    main()
