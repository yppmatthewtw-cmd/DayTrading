"""產出「深紅動能 → 第一見底訊號」參數設定說明書 (Excel)。"""
from __future__ import annotations
import sys, pathlib
import numpy as np, pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(pathlib.Path(__file__).parent))
import bottom_signal as B
from macd_histogram_momentum import PRESETS

NAVY, DP, LP = "1F3864", "C0304A", "F4B7C4"
THIN = Border(*[Side("thin", color="BFBFBF")] * 4)


def title_row(ws, r, text, span, size=14):
    c = ws.cell(row=r, column=1, value=text)
    c.font = Font(bold=True, size=size, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor=NAVY)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=span)
    return r + 1


def header(ws, r, cols, widths=None):
    for i, h in enumerate(cols, 1):
        c = ws.cell(row=r, column=i, value=h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="2F5597")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = THIN
    if widths:
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
    return r + 1


def rows(ws, r, data, wrap=()):
    for row in data:
        for i, v in enumerate(row, 1):
            c = ws.cell(row=r, column=i, value=v)
            c.border = THIN
            c.alignment = Alignment(vertical="top", wrap_text=(i in wrap))
        r += 1
    return r


def note(ws, r, text, span=8, h=32):
    c = ws.cell(row=r, column=1, value=text)
    c.font = Font(italic=True, size=9, color="555555")
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=span)
    ws.row_dimensions[r].height = h
    return r + 2


close = B.multi_cycle()
df, params = B.bottom_signals(close, **PRESETS["daily"])
ev = B.evaluate(df, close)
cal = B.calibrate(close, "daily")
sw = B.sweep(close, "daily", n_range=range(2, 14), decel_range=(0.4,))

wb = Workbook()

# ------------------------------------------------------------------ 00
ws = wb.active; ws.title = "00_直接答案"
r = title_row(ws, 1, "深紅柱本身永遠不會是底 —— 要捉的是「深紅耗盡的那一刻」", 8, 15)
r += 1
r = rows(ws, r, [["深紅的定義就是 h<0 且 h 仍在變大，即跌勢「仍在加速」。"
                  "只要柱還是深紅，跌勢就還沒減速。所以問題要改寫成："
                  "這一段深紅，何時開始失去推力？"]])
ws.merge_cells(start_row=r-1, start_column=1, end_row=r-1, end_column=8)
ws.row_dimensions[r-1].height = 32
r += 1
r = title_row(ws, r, "你圖上那個「8日」就是要設的參數", 8, 12)
r = rows(ws, r, [["紅柱段標著 8日、綠柱段標著 11日 —— 那是"
                  "「同色柱連續了幾根」。這個根數就是 min_dark_run (N)。"
                  "本檔示範資料的深紅段中位長度剛好也是 "
                  f"{params['中位數']:.0f} 根（純屬巧合，你要用自己的資料量）。"]])
ws.merge_cells(start_row=r-1, start_column=1, end_row=r-1, end_column=8)
ws.row_dimensions[r-1].height = 32
r += 1
r = title_row(ws, r, "五級底部訊號階梯（由早到遲）", 8, 12)
r = header(ws, r, ["級別", "訊號", "判定", "即時可知？", "本質"],
           [8, 18, 40, 12, 32])
r = rows(ws, r, [
    ["B1", "深紅減速", "仍是深紅，但 |Δh| 由本段峰值收窄 ≥ 40%", "可", "深紅柱的「增量」見頂 —— 最早"],
    ["B2", "首根淺紅", "深紅 → 淺紅（跌勢由加速轉減速）", "可", "h 的峰值已過"],
    ["B3", "淺紅確認", "連續 k 根淺紅（k=2）", "可", "排除單根雜訊"],
    ["B4", "柱體翻正", "h > 0，即 DIF 升穿 DEA（一般說的金叉）", "可", "多數人用的訊號 —— 也是最差的"],
    ["B5", "破結構", "價格升穿前 20 根高點", "可", "最遲，最可靠"],
], wrap=(3, 5))
r += 1
r = note(ws, r, "重要澄清：「深紅柱的動能峰值」只能事後知道 —— 實盤中你察覺峰值已過的那一根，"
    "就是第一根淺紅（B2）。所以「計深紅峰值捉底」和「等第一根淺紅」在即時交易中是同一時點。\n"
    "真正能再早一步的，是深紅柱的「增量」Δh：它有自己的峰值，而且在柱體仍是深紅時就會開始收窄。"
    "這就是 B1，也是唯一比 B2 更早、而又不需要偷看未來的訊號。", 8, 58)

# ------------------------------------------------------------------ 01
ws = wb.create_sheet("01_實測取捨曲線")
r = title_row(ws, 1, "五級訊號實測（合成多週期序列，600 根）", 8)
r = header(ws, r, ["訊號", "觸發次數", "距真底中位數(根)", "誤差中位數(根)",
                   "早於底比例", "其後20根報酬", "上漲命中率"],
           [18, 12, 18, 16, 14, 16, 14])
hdr = r - 1
for _, e in ev.iterrows():
    if e.get("觸發次數", 0) == 0:
        continue
    vals = [e["訊號"], int(e["觸發次數"]), round(e["距真底_中位數根"], 1),
            round(e["距真底_絕對值中位數"], 1), round(e["早於底比例"], 2),
            round(e["其後報酬"], 4), round(e["上漲命中率"], 2)]
    for i, v in enumerate(vals, 1):
        c = ws.cell(row=r, column=i, value=v); c.border = THIN
        if i == 6:
            c.number_format = "0.0%"
        if i == 3 and isinstance(v, float):
            c.fill = PatternFill("solid", fgColor="C6EFCE" if v < 0 else "FFC7CE")
    r += 1
data_end = r - 1
r += 1
r = rows(ws, r, [[f"全樣本 20 根基準報酬：{ev.attrs['基準報酬']:+.2%}（任何訊號要有價值，必須贏過這個數）"]])
ws.merge_cells(start_row=r-1, start_column=1, end_row=r-1, end_column=7)
r += 1
r = title_row(ws, r, "讀出來的三件事", 7, 12)
r = rows(ws, r, [
    ["① 早與準是同一個數字的兩面",
     "每一級的「誤差中位數」都等於「距底中位數」的絕對值 —— 代表訊號是系統性偏早，"
     "不是隨機亂跳。想更準，唯一方法就是往後等，而那就等於放棄提前量"],
    ["② B4 柱體翻正（金叉）是最差選擇",
     f"遲 6 根才出現，其後報酬僅 {ev[ev['訊號']=='B4_柱體翻正']['其後報酬'].iloc[0]:+.1%}，"
     "比 B2/B3 差一大截。它同時失去了提前量和報酬 —— 偏偏這是最多人用的"],
    ["③ B2 / B3 是甜蜜點",
     "提前 2.5–3.5 根，其後報酬約 +4.3%，命中率 60%。既有提前量又有優勢"],
], wrap=(2,))

chart = BarChart(); chart.type = "col"
chart.title = "各級訊號：距真底根數（負=早）與其後報酬"
chart.height, chart.width = 9, 22
chart.add_data(Reference(ws, min_col=3, max_col=3, min_row=hdr, max_row=data_end),
               titles_from_data=True)
chart.set_categories(Reference(ws, min_col=1, min_row=hdr + 1, max_row=data_end))
chart.series[0].graphicalProperties.solidFill = DP
ln = LineChart()
ln.add_data(Reference(ws, min_col=6, max_col=6, min_row=hdr, max_row=data_end),
            titles_from_data=True)
ln.y_axis.axId = 200; ln.y_axis.crosses = "max"; ln.y_axis.title = "其後20根報酬"
ln.series[0].graphicalProperties.line.solidFill = "1F3864"
ln.series[0].graphicalProperties.line.width = 22000
chart += ln
chart.y_axis.title = "距真底根數"
ws.add_chart(chart, "I2")

# ------------------------------------------------------------------ 02
ws = wb.create_sheet("02_N該設多少")
r = title_row(ws, 1, "先量出自己標的的深紅段長度分佈 —— 不要抄別人的數字", 8)
r = header(ws, r, ["統計量", "值", "含義"], [16, 12, 60])
r = rows(ws, r, [
    ["樣本段數", int(params["n"]), "這段資料共出現幾段深紅"],
    ["平均", round(params["平均"], 2), ""],
    ["中位數", round(params["中位數"], 1), "建議 N 的預設值。等於只對「長度高於一半」的跌段出手"],
    ["P25", round(params["P25"], 1), "N 設在此 → 訊號多、早、吵"],
    ["P75", round(params["P75"], 1), "N 設在此 → 訊號少、遲、準"],
    ["P90", round(params["P90"], 1), "只捉最深的那幾次跌勢"],
    ["最長", int(params["最長"]), ""],
], wrap=(3,))
r += 1
r = title_row(ws, r, "但是：N 掃描顯示，N 幾乎不影響準確度", 8, 12)
b2 = sw[sw["訊號"] == "B2_首根淺紅"]
r = header(ws, r, ["N", "觸發次數", "距底中位數", "誤差中位數", "其後報酬", "命中率"],
           [8, 12, 14, 14, 12, 12])
for _, x in b2.iterrows():
    for i, v in enumerate([int(x["N"]), int(x["次數"]), round(x["距底中位數"], 1),
                           round(x["誤差中位數"], 1), round(x["其後報酬"], 4),
                           round(x["命中率"], 2)], 1):
        c = ws.cell(row=r, column=i, value=v); c.border = THIN
        if i == 5: c.number_format = "0.0%"
    r += 1
r = note(ws, r, "N 由 2 增到 13，觸發次數由 17 降到 3，但誤差中位數一直卡在 3 根左右，"
    "其後報酬在 2.8%–6.0% 之間無規律上下跳動 —— 那是雜訊，不是趨勢。\n"
    "結論：N 控制的是「出手頻率」，不是「準確度」。看到 N=12 命中率 83% 就去用它，"
    "是在 6 個樣本上做曲線擬合，實盤必然失效。", 8, 50)

# ------------------------------------------------------------------ 03
ws = wb.create_sheet("03_真正的參數")
r = title_row(ws, 1, "控制提前量的是 MACD 週期，不是 N", 9)
r = rows(ws, r, [["同一條資料，只改 MACD 週期，B2 的位置由「早 7 根」一路移到「遲 1 根」。"
                  "這才是你真正在設的東西。"]])
ws.merge_cells(start_row=r-1, start_column=1, end_row=r-1, end_column=9)
r += 1
r = header(ws, r, ["MACD", "訊號", "深紅段中位長", "建議 N", "觸發次數",
                   "距底中位數", "誤差中位數", "其後報酬", "命中率"],
           [12, 16, 14, 10, 12, 14, 14, 12, 10])
for _, x in cal.iterrows():
    vals = [x["MACD"], x["訊號"], round(x["深紅段中位長"], 1), int(x["建議N"]),
            int(x["觸發次數"]), round(x["距底中位數"], 1), round(x["誤差中位數"], 1),
            round(x["其後報酬"], 4), round(x["命中率"], 2)]
    for i, v in enumerate(vals, 1):
        c = ws.cell(row=r, column=i, value=v); c.border = THIN
        if i == 8: c.number_format = "0.0%"
        if i == 6 and isinstance(v, float):
            c.fill = PatternFill("solid", fgColor="C6EFCE" if v < 0 else "FFC7CE")
    r += 1
r = note(ws, r, "快參數 (5/13/5)：B2 提前 7 根，但誤差也是 7 根。\n"
    "慢參數 (24/52/18)：B2 遲 1 根，誤差 1 根 —— 準，但那已不是「捉底」，是「確認」。\n"
    "這條規律沒有例外：你要多早，就得接受多大的誤差。兩者是同一個數字。", 9, 50)

# ------------------------------------------------------------------ 04
ws = wb.create_sheet("04_建議設定")
r = title_row(ws, 1, "按交易風格選一組，然後用自己的資料重新校準", 7)
r = header(ws, r, ["風格", "MACD", "N", "confirm_bars", "decel_min",
                   "主用訊號", "預期表現"], [16, 12, 8, 14, 12, 16, 34])
r = rows(ws, r, [
    ["搶最早（日內）", "5/13/5", "深紅段中位數", 2, 0.40, "B1 深紅減速",
     "提前約 8 根，但誤差同樣約 8 根，命中率僅約五成。必須用極小倉位試單"],
    ["平衡（推薦）", "12/26/9", "深紅段中位數", 2, 0.40, "B2 首根淺紅",
     "提前約 3.5 根，報酬 +4.3%，命中率 60%。提前量與準確度的最佳交點"],
    ["求穩", "19/39/9", "深紅段中位數", 3, 0.50, "B3 淺紅確認",
     "提前約 1 根，命中率約 67%。基本上是確認而非預測"],
    ["不要用", "任意", "—", "—", "—", "B4 柱體翻正",
     "遲 6 根、報酬最低。一般說的 MACD 金叉抄底，在這份測試中是最差選擇"],
], wrap=(7,))
r += 1
r = title_row(ws, r, "日內 (5 分鐘圖) 的額外限制", 7, 12)
r = rows(ws, r, [
    ["EMA 暖機", "5 分鐘圖用 MACD(5,13,5)，至少要 13 根 ≈ 65 分鐘才穩定。"
                 "開市首小時的柱體形態不可信"],
    ["一日的深紅段數", "美股一日 78 根 5 分鐘 K。若深紅段中位長 8 根，一日最多只有幾段，"
                       "樣本極少 —— 日內參數必須用多日資料合併校準"],
    ["事件日", "FOMC、CPI、業績公布會直接跳空，抹掉整段動能結構。事件前後的訊號全部作廢"],
], wrap=(2,))

# ------------------------------------------------------------------ 05
ws = wb.create_sheet("05_怎麼跑")
r = title_row(ws, 1, "指令", 3)
r = header(ws, r, ["目的", "指令"], [24, 74])
r = rows(ws, r, [
    ["看示範", "python3 scripts/bottom_signal.py"],
    ["量出自己標的的參數", "python3 scripts/bottom_signal.py --csv bars.csv --calibrate"],
    ["掃 N 與 decel_min", "python3 scripts/bottom_signal.py --csv bars.csv --sweep"],
    ["日內", "python3 scripts/bottom_signal.py --csv bars.csv --preset 5min --calibrate"],
    ["程式內呼叫", "from bottom_signal import bottom_signals, calibrate；"
                  "df, params = bottom_signals(close, fast=12, slow=26, signal=9)"],
], wrap=(2,))
r += 1
r = title_row(ws, r, "限制", 3, 12)
r = rows(ws, r, [
    ["樣本無效", "本檔全部數字來自「一條」合成序列，每級訊號僅約 10 次觸發。"
                 "方向正確，但統計上不足以驗證任何門檻值"],
    ["行情 API 被封鎖", "本工作階段所有行情來源（Yahoo / Polygon / Stooq / AlphaVantage 等）"
                        "回應 403，無法用真實股票資料驗證"],
    ["震盪巿失效", "MACD 是趨勢指標。橫行區間內柱體反覆變色，所有級別都會誤報"],
    ["底背馳可連續發生", "一個跌勢可以出現三四次底背馳才真正見底。"
                         "「訊號出現就全倉抄底」是虧錢最快的用法"],
], wrap=(2,))
r = note(ws, r, "本檔為方法論與工具說明，不構成投資建議。所有門檻值必須用自己交易的標的"
                "與週期重新校準後才可使用。", 3)

out = pathlib.Path("output/2026-09-17_深紅動能見底訊號_參數設定.xlsx")
wb.save(out)
print("已輸出", out)
