"""產出 MACD 柱狀圖動能量化說明書 (Excel)。"""
from __future__ import annotations
import sys, pathlib
import numpy as np, pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(pathlib.Path(__file__).parent))
import macd_histogram_momentum as M

NAVY, GREY, DG, LG, LP, DP = "1F3864", "D9E1F2", "1E7B34", "A9D5A9", "F4B7C4", "C0304A"
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


def rows(ws, r, data, wrap_cols=()):
    for row in data:
        for i, v in enumerate(row, 1):
            c = ws.cell(row=r, column=i, value=v)
            c.border = THIN
            c.alignment = Alignment(vertical="top", wrap_text=(i in wrap_cols))
        r += 1
    return r


def note(ws, r, text, span=8):
    c = ws.cell(row=r, column=1, value=text)
    c.font = Font(italic=True, size=9, color="555555")
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=span)
    ws.row_dimensions[r].height = 30
    return r + 2


# ===================== 計算 demo 供圖表與驗證 =====================
close = M.demo_series()
df, waves, dvg, energy, purity = M.analyse(close)
fwd = df["close"].shift(-20) / df["close"] - 1

wb = Workbook()

# ---------------------------------------------------------------- 00
ws = wb.active; ws.title = "00_核心答案"
r = title_row(ws, 1, "MACD 柱狀圖動能量化 — 可以計算，而且分四個層級", 8, 15)
r += 1
r = rows(ws, r, [["結論：可以。柱體 h = DIF − DEA 本身就是數值；"
                  "「深綠/深粉」只是 h 的一階差分符號，把它還原成數字即可。"]])
ws.merge_cells(start_row=r-1, start_column=1, end_row=r-1, end_column=8)
r += 1
r = header(ws, r, ["層級", "指標", "公式", "回答什麼問題", "訊號早晚", "假訊號率"],
           [6, 20, 42, 34, 12, 10])
r = rows(ws, r, [
    ["①", "柱值 h", "h = (DIF − DEA) × 2", "此刻動能有多強（正=多方）", "—", "—"],
    ["②", "顏色 (四色)", "h 的正負 × h 一階差分的正負", "動能在加速還是衰減", "最早", "很高"],
    ["③", "面積 M / 峰值 H / 密度 D", "M = Σ|h| 於同號區間；H = max|h|；D = M / 根數",
     "整波推動力有多大、多急", "中", "中"],
    ["④", "背馳強度 DVG", "1 − A₂/A₁ （且價格創新極值）",
     "價格新高但動能萎縮多少 ← 最有意義", "中", "較低"],
    ["⑤", "轉勢綜合分 TRS", "六項證據加權，0–100", "轉勢走到哪一步", "綜合", "見 05 頁"],
], wrap_cols=(3, 4))
r += 1
r = title_row(ws, r, "直接回答：深綠的上升動能數值 / 深粉的下跌動能數值", 8, 12)
r = header(ws, r, ["顏色", "含義", "根數", "動能總量 Σ|h|", "佔價格 %", "平均每根", "最大單根"],
           [10, 24, 8, 16, 12, 12, 12])
meaning = {"深綠": "上升動能加速（h>0 且變大）", "淺綠": "上升動能衰減（h>0 但變小）",
           "淺紅": "下跌動能衰減（h<0 但收窄）", "深紅": "下跌動能加速（h<0 且變大）"}
fills = {"深綠": DG, "淺綠": LG, "淺紅": LP, "深紅": DP}
for _, e in energy.iterrows():
    ws.cell(row=r, column=1, value=e["顏色"]).fill = PatternFill("solid", fgColor=fills[e["顏色"]])
    ws.cell(row=r, column=1).font = Font(bold=True,
        color="FFFFFF" if e["顏色"] in ("深綠", "深紅") else "000000")
    for i, v in enumerate([meaning[e["顏色"]], int(e["根數"]), round(e["動能總量"], 3),
                           round(e["動能_佔價格pct"], 2), round(e["平均每根"], 4),
                           round(e["最大單根"], 4)], 2):
        ws.cell(row=r, column=i, value=v).border = THIN
    ws.cell(row=r, column=1).border = THIN
    r += 1
r += 1
r = rows(ws, r, [
    ["動能純度（多方）", f"{purity['多方動能純度']:.1%}", "深綠 ÷ (深綠+淺綠)。>65% 趨勢健康；<45% 是蠕動不是趨勢"],
    ["動能純度（空方）", f"{purity['空方動能純度']:.1%}", "深紅 ÷ (深紅+淺紅)"],
    ["多空動能比", f"{purity['多空動能比']:.2f}", ">1 多方主導；<1 空方主導"],
], wrap_cols=(3,))
r = note(ws, r, "以上數值取自本檔內建的合成驗證資料（見 05 頁），用作示範計算，"
                "不是任何真實標的的讀數。換入自己的資料後數字會全部改變。")

# ---------------------------------------------------------------- 01
ws = wb.create_sheet("01_四色定義與公式")
r = title_row(ws, 1, "四色柱的嚴格定義", 6)
r = header(ws, r, ["顏色", "條件", "數學意義", "對應價格", "交易含義"],
           [10, 26, 30, 26, 34])
r = rows(ws, r, [
    ["深綠", "h > 0 且 h ≥ h₋₁", "h 正、Δh 正：動能二階為正", "升勢且加速", "持有／順勢加碼，最安全的一段"],
    ["淺綠", "h > 0 且 h < h₋₁", "h 正、Δh 負：動能開始衰減", "仍在升，但變慢", "頂部第一預警。不等於轉勢，可能只是整理"],
    ["淺紅", "h < 0 且 h ≥ h₋₁", "h 負、Δh 正：跌勢收窄", "仍在跌，但變慢", "底部第一預警。抄底最早的線索，也最容易被套"],
    ["深紅", "h < 0 且 h < h₋₁", "h 負、Δh 負：下跌加速", "跌勢且加速", "空方主導，不接飛刀"],
], wrap_cols=(3, 4, 5))
r += 1
r = title_row(ws, r, "為什麼顏色比均線交叉早", 6, 12)
r = rows(ws, r, [
    ["價格 P", "0 階", "—"],
    ["DIF = EMA₁₂ − EMA₂₆", "≈ 1 階（速度）", "價格的變化率"],
    ["h = DIF − DEA", "≈ 2 階（加速度）", "DIF 相對自身均線的偏離"],
    ["顏色 = sign(Δh)", "≈ 3 階（加加速度）", "最早，也最吵"],
])
r = note(ws, r, "訊號序列固定為：深綠→淺綠（最早，假訊號最多）→ 柱體翻負（DIF 跌穿 DEA）"
                "→ DIF 跌穿零軸（最遲，最可靠）。愈早的訊號必須用愈嚴的過濾條件，"
                "否則在震盪巿會被反覆打臉。")
r = title_row(ws, r, "MACD 參數（日內必須縮短，經典 12/26/9 在 5 分鐘圖上慢到沒有交易價值）", 6, 12)
r = header(ws, r, ["週期", "fast", "slow", "signal", "說明"], [14, 8, 8, 10, 46])
r = rows(ws, r, [
    ["日線 / 週線", 12, 26, 9, "經典值，適合波段"],
    ["1 小時", 12, 26, 9, "仍可用經典值"],
    ["15 分鐘", 8, 17, 9, "日內中速"],
    ["5 分鐘", 5, 13, 5, "日內主力參數"],
    ["1 分鐘", 5, 13, 5, "雜訊極高，顏色訊號幾乎不可單獨使用"],
], wrap_cols=(5,))

# ---------------------------------------------------------------- 02
ws = wb.create_sheet("02_動能量化指標")
r = title_row(ws, 1, "把一整波的動能變成一個數字", 6)
r = header(ws, r, ["指標", "符號", "公式", "解讀", "門檻建議"], [18, 8, 34, 40, 24])
r = rows(ws, r, [
    ["動能面積", "M", "M = Σ|hₜ|，t 遍歷同號連續區間",
     "整波推動力的總量。可跨波比較，是背馳的基礎", "無絕對門檻，只比相對值"],
    ["標準化面積", "M%", "M% = M ÷ 平均價 × 100",
     "去掉價位影響，讓不同標的可比", "同上"],
    ["動能峰值", "H", "H = max|hₜ|", "這一波最猛的一刻有多猛", "同上"],
    ["峰值位置", "—", "峰值出現在第幾根", "前段見峰=衝高回落；後段見峰=推升有力", "峰值在前 1/3 = 續航力弱"],
    ["動能密度", "D", "D = M ÷ 根數", "急升急跌 vs 慢牛慢熊", "D 高=情緒盤，回撤大"],
    ["衰減率", "λ", "λ = (至今峰值 − 現值) ÷ 至今峰值",
     "本波動能已耗掉幾成。λ∈[0,1]", "λ<0.3 仍強；0.3–0.6 警戒；>0.6 基本耗盡"],
    ["動能純度", "—", "深色動能 ÷ (深色+淺色)", "趨勢是推動還是蠕動", ">0.65 健康；<0.45 不宜追"],
], wrap_cols=(3, 4, 5))
r = note(ws, r, "λ 一定要用「至今為止的峰值」(cummax) 而非整波峰值。用整波峰值等於偷看未來，"
                "回測會漂亮、實盤會失效 —— 這是動能指標最常見的未來函數陷阱。")
r += 1
r = title_row(ws, r, "示範：合成資料切出的動能波（雜訊波已標示，不參與背馳比較）", 12, 12)
hdr = r
r = header(ws, r, ["波", "方向", "級別", "柱數", "面積 M", "M%", "峰值 H", "峰值位置",
                   "密度 D", "深色根數", "淺色根數", "價格極值"],
           [6, 12, 8, 8, 12, 10, 10, 10, 10, 10, 10, 12])
for _, w_ in waves.iterrows():
    vals = [int(w_["wave"]), w_["方向"], w_["級別"], int(w_["柱數"]), round(w_["面積M"], 3),
            round(w_["面積M_pct"], 2), round(w_["峰值H"], 3), int(w_["峰值位置"]),
            round(w_["密度D"], 3), int(w_["深色根數"]), int(w_["淺色根數"]),
            round(w_["價格極值"], 2)]
    for i, v in enumerate(vals, 1):
        c = ws.cell(row=r, column=i, value=v); c.border = THIN
        if w_["級別"] == "雜訊波":
            c.font = Font(color="999999", italic=True)
    r += 1
r = note(ws, r, "級別過濾：柱數 < 3 或面積不足同向波中位數 15% 者列為雜訊波。"
                "沒有這一步，一兩根柱的反彈會被拿去和主升浪比面積，得出 99% 的假背馳。"
                "背馳只有在同級別的波之間才成立。", 12)

# ---------------------------------------------------------------- 03
ws = wb.create_sheet("03_背馳量化")
r = title_row(ws, 1, "背馳是柱狀圖唯一有統計意義的轉勢訊號", 9)
r = rows(ws, r, [["定義：價格創新高（或新低），但同向動能面積萎縮。"
                  "前提是兩波必須同級別。"]])
ws.merge_cells(start_row=r-1, start_column=1, end_row=r-1, end_column=9)
r += 1
r = header(ws, r, ["項目", "公式 / 判定"], [22, 62])
r = rows(ws, r, [
    ["頂背馳條件", "P₂ > P₁（價格新高） 且 A₂ < A₁（綠波面積萎縮）"],
    ["底背馳條件", "P₂ < P₁（價格新低） 且 A₂ < A₁（紅波面積萎縮）"],
    ["背馳強度", "DVG = 1 − A₂/A₁ ∈ (0, 1)"],
    ["判定", "DVG > 0.30 顯著背馳；0 < DVG ≤ 0.30 輕微；A₂ ≥ A₁ 動能同步（不是背馳）"],
    ["輔助", "峰值比 H₂/H₁ 同步看。面積與峰值同時萎縮，訊號強度最高"],
], wrap_cols=(2,))
r += 1
r = title_row(ws, r, "示範：合成資料的背馳檢測（僅主波）", 9, 12)
r = header(ws, r, ["前波", "本波", "方向", "價格創新極值", "價格變動%",
                   "面積比 A₂/A₁", "峰值比 H₂/H₁", "背馳強度", "判定"],
           [8, 8, 12, 14, 12, 14, 14, 12, 14])
for _, d_ in dvg.iterrows():
    vals = [d_["前波"], d_["本波"], d_["方向"], d_["價格創新極值"],
            round(d_["價格變動_pct"], 2), round(d_["面積比A2_A1"], 3),
            round(d_["峰值比H2_H1"], 3),
            None if pd.isna(d_["背馳強度"]) else round(d_["背馳強度"], 3), d_["判定"]]
    for i, v in enumerate(vals, 1):
        c = ws.cell(row=r, column=i, value=v); c.border = THIN
        if i == 9 and v == "顯著背馳":
            c.fill = PatternFill("solid", fgColor="FFD966"); c.font = Font(bold=True)
    r += 1
r = note(ws, r, "留意 wave 2 → wave 8：價格 +6.29%，但動能面積只剩 61.8%，背馳強度 0.382 "
                "—— 這正是合成資料中刻意造出的頂背馳。演算法把它抓了出來。", 9)

# ---------------------------------------------------------------- 04
ws = wb.create_sheet("04_轉勢評分TRS")
r = title_row(ws, 1, "轉勢綜合分 TRS（0–100）", 5)
r = header(ws, r, ["證據", "權重", "判定", "在轉勢過程中出現的次序"], [26, 10, 40, 26])
r = rows(ws, r, [
    ["顏色由深轉淺", 20, "出現淺色柱（深綠→淺綠 / 深紅→淺紅）", "① 最早"],
    ["衰減率 λ > 0.5", 15, "本波動能已耗掉一半以上", "②"],
    ["連續淺色 ≥ 3 根", 15, "排除單根雜訊", "③"],
    ["背馳強度 > 0.3", 25, "同級別波的面積萎縮逾三成（權重最高）", "④"],
    ["柱體已過零軸", 15, "h 翻負（頂）／翻正（底）", "⑤"],
    ["價格破前低／前高", 10, "跌穿前 20 根低點（頂）／升穿高點（底）", "⑥ 最遲"],
    ["合計", 100, "", ""],
], wrap_cols=(3,))
r += 1
r = header(ws, r, ["TRS", "分級", "行動"], [14, 22, 50])
for lo, hi, g, act, colr in [
    (0, 35, "趨勢延續", "持有。淺色柱在健康升勢中本來就會不斷出現", "C6EFCE"),
    (35, 60, "動能衰竭·減倉", "減倉、收緊止蝕。不反手", "FFEB9C"),
    (60, 80, "轉勢確認中·離場", "清倉。此時多半已過高點", "FFC7A0"),
    (80, 100, "已轉勢·可反手", "可考慮反向，但確認度換來的是滯後", "FFC7CE")]:
    for i, v in enumerate([f"{lo}–{hi}", g, act], 1):
        c = ws.cell(row=r, column=i, value=v); c.border = THIN
        c.fill = PatternFill("solid", fgColor=colr)
    r += 1
r += 1
r = note(ws, r, "兩個必要的工程細節，缺一個分數就會失真：\n"
    "① 證據鎖存（10 根）：背馳發生在柱體翻負「之前」，破結構在「之後」，"
    "若只看當根，兩者永不同時成立，分數上限被結構性壓在 75 分，80 分那一級數學上不可能出現。\n"
    "② 情境閘：「頂部轉勢」只在升勢情境中計分（近 15 根內曾站上 MA50），"
    "否則跌勢中每次深紅轉淺紅都會拿分 —— 那是反彈，不是轉勢。", 5)
ws.row_dimensions[r-2].height = 75

# ---------------------------------------------------------------- 05
ws = wb.create_sheet("05_實測驗證")
r = title_row(ws, 1, "合成資料驗證 —— 連同它的失敗一起列出", 7)
peak_i = df["close"].idxmax()
r = header(ws, r, ["檢驗項目", "結果", "判讀"], [30, 30, 48])
res = []
for th in (60, 80):
    f = df[df["TRS_頂"] >= th]
    if len(f):
        res.append([f"TRS ≥ {th} 首次觸發", f"{f.index[0].date()} @ {f['close'].iloc[0]:.2f}",
                    f"距價格最高點 {(f.index[0]-peak_i).days:+d} 曆日"])
res += [
    ["價格最高點", f"{peak_i.date()} @ {df['close'].max():.2f}", "合成資料的真實頂部"],
    ["最高點當日 TRS", f"{df.loc[peak_i,'TRS_頂']:.0f}（{df.loc[peak_i,'TRS_頂_分級']}）",
     "頂部當日已進入離場級別"],
    ["次一根 TRS", f"{df['TRS_頂'].iloc[df.index.get_loc(peak_i)+1]:.0f}（已轉勢·可反手）",
     "轉勢確認滯後一根 —— 這是可接受的滯後"],
]
for th in (60, 80):
    m = df["TRS_頂"] >= th
    res.append([f"TRS ≥ {th} 觸發頻率", f"{int(m.sum())} / {len(df)} 根（{m.sum()/len(df):.0%}）",
                f"其後 20 根平均報酬 {fwd[m].mean():+.2%}，"
                f"全樣本基準 {fwd.mean():+.2%}，下跌命中率 {(fwd[m]<0).mean():.0%}"])
r = rows(ws, r, res, wrap_cols=(3,))
r += 1
r = title_row(ws, r, "誠實的結論", 7, 12)
r = rows(ws, r, [
    ["① 指標本身是可靠的算術", "面積、峰值、密度、衰減率、背馳強度都是確定性計算，沒有擬合成分"],
    ["② TRS ≥ 80 方向正確但樣本無效", "本檔只在「一條」合成序列上測過：其後 20 根報酬 −1.08% vs 基準 +0.99%，"
     "看似有 2.1 個百分點優勢，但 n=6，統計上毫無意義"],
    ["③ TRS ≥ 60 幾乎沒有優勢", "觸發率高達 30%，其後報酬 +0.48% vs 基準 +0.99% —— 只能當狀態顯示，不能當進出訊號"],
    ["④ 背馳訊號天生偏早", "TRS 首次過 60 比真正頂部早 165 個曆日。背馳是「風險在上升」，不是「今天見頂」"],
    ["⑤ 必須用自己的標的回測", "門檻 0.30 / 0.50 / 3 根 / 10 根鎖存全部是起點，不是真理"],
], wrap_cols=(2,))
r = note(ws, r, "本頁所有數字來自 scripts/macd_histogram_momentum.py 的 --demo 合成序列。"
                "本工作階段的行情 API 全部被網絡政策封鎖（403），因此無法用真實股票資料驗證。"
                "換入真實資料的方法見 06 頁。", 7)

# ---------------------------------------------------------------- 06 圖
ws = wb.create_sheet("06_四色柱示意圖")
r = title_row(ws, 1, "四色柱狀圖（合成資料，頂部轉勢前後 70 根）", 8)
seg = df.iloc[150:220]
hdr_r = r + 1
ws.cell(row=hdr_r, column=1, value="日期")
for i, c in enumerate(["深綠", "淺綠", "淺紅", "深紅", "收市價"], 2):
    ws.cell(row=hdr_r, column=i, value=c).font = Font(bold=True)
rr = hdr_r + 1
for idx, row_ in seg.iterrows():
    ws.cell(row=rr, column=1, value=idx.strftime("%m-%d"))
    for i, c in enumerate(["深綠", "淺綠", "淺紅", "深紅"], 2):
        ws.cell(row=rr, column=i, value=float(row_["hist"]) if row_["colour"] == c else None)
    ws.cell(row=rr, column=6, value=round(float(row_["close"]), 2))
    rr += 1

bar = BarChart(); bar.type = "col"; bar.grouping = "stacked"; bar.overlap = 100
bar.title = "MACD 柱狀圖四色分解"; bar.height = 9; bar.width = 30
data = Reference(ws, min_col=2, max_col=5, min_row=hdr_r, max_row=rr - 1)
cats = Reference(ws, min_col=1, min_row=hdr_r + 1, max_row=rr - 1)
bar.add_data(data, titles_from_data=True); bar.set_categories(cats)
for s_, colr in zip(bar.series, [DG, LG, LP, DP]):
    s_.graphicalProperties.solidFill = colr
    s_.graphicalProperties.line.solidFill = colr
bar.y_axis.title = "柱值 h"
ln = LineChart()
ln.add_data(Reference(ws, min_col=6, min_row=hdr_r, max_row=rr - 1), titles_from_data=True)
ln.y_axis.axId = 200; ln.y_axis.title = "收市價"; ln.y_axis.crosses = "max"
ln.series[0].graphicalProperties.line.solidFill = "1F3864"
ln.series[0].graphicalProperties.line.width = 20000
bar += ln
ws.add_chart(bar, "H2")
for col in range(1, 7):
    ws.column_dimensions[get_column_letter(col)].width = 11
ws.column_dimensions["A"].width = 9

# ---------------------------------------------------------------- 07
ws = wb.create_sheet("07_實戰應用與陷阱")
r = title_row(ws, 1, "日內交易怎麼用", 4)
r = header(ws, r, ["情境", "做法"], [26, 68])
r = rows(ws, r, [
    ["參數", "5 分鐘圖用 MACD(5,13,5)；15 分鐘用 (8,17,9)。用 12/26/9 跑 5 分鐘圖，"
             "訊號出來時當日行情已經走完"],
    ["開市首 30 分鐘", "不要用。EMA 需要暖機，5 分鐘圖至少要 26 根 ≈ 130 分鐘才穩定，"
             "即約 11:40 ET。這與 10 分鐘均線需要 100 分鐘才成形是同一個道理"],
    ["主要用法", "不是拿來找入場點，是拿來「退出」。持倉中出現 淺綠 + λ>0.5 + 連續 3 根 → 收緊止蝕"],
    ["配合價格結構", "圖中的「升破前頂 / 跌破前底」正是 TRS 的第 ⑥ 項。單看柱體不夠，"
             "必須有價格結構確認，否則假訊號率無法接受"],
    ["事件日", "FOMC、業績、CPI 公布前後，柱體形態完全失效 —— 事件會直接跳空抹掉既有動能結構"],
], wrap_cols=(2,))
r += 1
r = title_row(ws, r, "五個必須知道的陷阱", 4, 12)
r = rows(ws, r, [
    ["① 顏色不是訊號", "深綠轉淺綠只表示「上升變慢」。健康升勢中這會發生幾十次，單獨用它必然被反覆打臉"],
    ["② 背馳可以連續發生", "頂背馳之後再創新高、再背馳，可以重複三四次。「背馳了就沽」是虧錢最快的用法"],
    ["③ 級別必須一致", "一兩根柱的反彈不是一個波。不做級別過濾，背馳強度會出現 0.99 這種假數字"],
    ["④ 不要用整波峰值算 λ", "那是未來函數。必須用 cummax（至今為止的峰值）"],
    ["⑤ 震盪巿全部失效", "MACD 是趨勢指標。橫行區間內柱體反覆變色，TRS 會不斷誤報"],
], wrap_cols=(2,))
r = note(ws, r, "免責：本檔為方法論與計算工具說明，不構成投資建議。"
                "所有門檻值必須用自己交易的標的與週期回測校準後才可使用。", 4)

# ---------------------------------------------------------------- 08
ws = wb.create_sheet("08_如何換入真實資料")
r = title_row(ws, 1, "把合成資料換成真實行情", 3)
r = header(ws, r, ["步驟", "做法"], [10, 84])
r = rows(ws, r, [
    ["1", "準備一個 CSV：第一欄為時間索引，並含 close（或 Close）欄"],
    ["2", "執行：python3 scripts/macd_histogram_momentum.py --csv bars.csv --preset 5min"],
    ["3", "或在程式內：from macd_histogram_momentum import analyse；"
          "df, waves, dvg, energy, purity = analyse(close_series, fast=5, slow=13, signal=5)"],
    ["4", "輸出明細：加 --out result.csv"],
    ["5", "若已有行情來源，可直接沿用 scripts/screen_20min_ma.py 內的 fetch_intraday()，"
          "它已封裝好取價邏輯，換掉資料源即可"],
], wrap_cols=(2,))
r = note(ws, r, "本工作階段的行情 API（Yahoo / Polygon / Stooq / AlphaVantage 等）"
                "全部被網絡出口政策封鎖，回應 403，故無法在此直接接上真實資料。", 3)

out = pathlib.Path("output/2026-09-17_MACD柱狀圖動能量化_轉勢指標.xlsx")
out.parent.mkdir(exist_ok=True)
wb.save(out)
print("已輸出", out)
