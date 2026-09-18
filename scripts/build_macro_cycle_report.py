# -*- coding: utf-8 -*-
"""
四大力量牛熊週期 + 日內交易框架  (Data as of 2026-09-09)
產出：Excel (含原生圖表) + Word 報告。不產生 PDF / 圖片。
"""
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, BarChart, Reference, Series
from openpyxl.chart.label import DataLabelList

OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "output")
os.makedirs(OUT, exist_ok=True)
ASOF = "2026-09-09"

# ---------- styles ----------
NAVY   = "1F3864"; SLATE = "2F5597"; LIGHT = "D9E2F3"
RED    = "FFC7CE"; AMBER = "FFEB9C"; GREEN = "C6EFCE"
REDF   = "9C0006"; AMBERF= "9C6500"; GREENF= "006100"
THIN   = Side(style="thin", color="BFBFBF")
BOX    = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

def title_row(ws, r, text, span, size=14):
    c = ws.cell(row=r, column=1, value=text)
    c.font = Font(bold=True, size=size, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor=NAVY)
    c.alignment = Alignment(vertical="center")
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=span)
    ws.row_dimensions[r].height = 24

def note_row(ws, r, text, span):
    c = ws.cell(row=r, column=1, value=text)
    c.font = Font(italic=True, size=9, color="595959")
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=span)

def header(ws, r, headers):
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=r, column=i, value=h)
        c.font = Font(bold=True, size=10, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=SLATE)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BOX
    ws.row_dimensions[r].height = 30

def body(ws, r0, rows, wraps=(), zebra=True):
    for j, row in enumerate(rows):
        for i, v in enumerate(row, start=1):
            c = ws.cell(row=r0 + j, column=i, value=v)
            c.font = Font(size=10)
            c.border = BOX
            c.alignment = Alignment(vertical="top", wrap_text=(i in wraps))
            if zebra and j % 2 == 1:
                c.fill = PatternFill("solid", fgColor="F2F5FB")
    return r0 + len(rows)

def widths(ws, spec):
    for col, w in spec.items():
        ws.column_dimensions[col].width = w

def rag(ws, row, col, level):
    fills = {"R": (RED, REDF), "A": (AMBER, AMBERF), "G": (GREEN, GREENF)}
    bg, fg = fills[level]
    c = ws.cell(row=row, column=col)
    c.fill = PatternFill("solid", fgColor=bg)
    c.font = Font(size=10, bold=True, color=fg)
    c.alignment = Alignment(horizontal="center", vertical="center")

wb = Workbook()

# =====================================================================
# 00 現況定位 Dashboard
# =====================================================================
ws = wb.active; ws.title = "00_現況定位"
widths(ws, {"A":18,"B":26,"C":10,"D":30,"E":34,"F":34,"G":16})
title_row(ws, 1, f"四大力量牛熊週期定位儀表板    資料截至 {ASOF}", 7, 15)
note_row(ws, 2, "階段進度 = 該力量在本輪週期中已走完的百分比估算（0% = 週期起點，100% = 週期完結/反轉）。", 7)

header(ws, 4, ["力量", "本輪週期", "階段進度", "目前所處階段", "上一個轉折點（觸發事件）",
               "下一個轉折點（需要看到什麼）", "對股市方向"])
dash = [
 ["① 宏觀流動性\n（息口）", "2024-09 減息 → 2026 轉鷹", 0.90,
  "寬鬆週期已結束，處於\n『政策轉向再收緊』的門檻上",
  "2025-12 最後一次減息至 3.50–3.75%；\n2026-05 Warsh 接任主席宣示 regime change",
  "9/16 FOMC 加息或點陣圖上移 = 正式進入收緊；\n反之若油價回落、頭條CPI<3% 則延續觀望",
  "偏空（估值壓縮）"],
 ["② 政治週期\n（總統）", "Trump 第二任期第 2 年\n（期中選舉年）", 0.75,
  "四年週期最弱的一段：\n第2年 Q3–Q4『築底期』",
  "2026-01 起支持度自 50%+ 跌至高 30%；\n施政受阻、分裂政府預期升溫",
  "11/3 期中選舉揭盅 = 不確定性出清點；\n1950 年以來期中選舉後 12 個月標普 100% 上漲",
  "短空長多（拐點將近）"],
 ["③ 技術革命\n（產業）", "AI 資本開支超級週期\n2023 →", 0.65,
  "資本開支高潮期（加速段末端），\n尚未進入狂熱崩解",
  "2026 五大雲廠 capex 指引升至 7,750–8,000 億美元\n（2025: 4,100億 / 2024: 2,380億）",
  "首次出現 capex 指引下修、折舊年限爭議、\n債務融資 capex 利差擴大 = 見頂訊號",
  "偏多（唯一支撐力量）"],
 ["④ 黑天鵝\n（地緣/戰爭）", "2026 伊朗戰爭 /\n荷姆茲海峽危機", 0.55,
  "由『衝擊期』轉入『持久消耗期』，\n尚未見到解決催化劑",
  "2026-02-28 美以空襲伊朗；03-19 海峽關閉；\n07-14 恢復海上封鎖；9月油輪戰持續",
  "停火/換屆談判、伊朗石油恢復出口，\n或需求破壞使 Brent 回落至 80 以下",
  "偏空（透過油價傳導）"],
]
r = body(ws, 5, dash, wraps=(1,2,4,5,6,7))
for i in range(5, r):
    ws.cell(row=i, column=3).number_format = "0%"
    ws.cell(row=i, column=3).alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(row=i, column=3).font = Font(size=12, bold=True)
    ws.row_dimensions[i].height = 78

r += 1
title_row(ws, r, "核心結論：四力合成判斷", 7, 13); r += 1
concl = [
 ["主線邏輯", "荷姆茲海峽供應中斷 → Brent 按年 +46.7% → 頭條CPI 3.4% 高於核心 2.5%（差距 0.9pt 全來自能源）"
             " → 鷹派新主席 Warsh 被迫考慮加息 → 10 年期 4.77% → 對一個由 AI 成長定價的 7,681 點標普造成估值壓縮。"
             "四大力量中有三個同時指向逆風，唯一支撐是 AI 資本開支。"],
 ["最大不對稱", "VIX 只有 15.3，但油價按年 +46.7%、加息機率 60%、10年期 4.77%。"
              "『低波動 + 高尾部風險』= 保護成本極低而風險未被定價，這是目前市場最錯價的地方。"],
 ["歷史對標", "1973–74（禁運 + 收緊 + 水門/期中年）、1990（科威特 + 衰退 + 期中年）、"
             "2022（烏克蘭 + 最快加息 + 期中年）。三次的共通結構與現在高度相似，且三次的底部都出現在期中選舉年的 Q3–Q4。"],
 ["時間窗", "9/11 CPI 與 9/16 FOMC 是本輪最密集的風險窗口；11/3 期中選舉是不確定性出清點。"
           "歷史型態指向『10 月前後築底、選後至翌年 Q2 為最強上升段』。"],
 ["定位建議", "戰術上防守（現金/能源/黃金/短天期），戰略上準備在 Q4 的恐慌中建倉。"
             "不要在 7,600 之上追多，也不要在 VIX 低於 16 時裸露空頭 gamma。"],
]
header(ws, r, ["項目", "說明"]); r += 1
for k, v in concl:
    ws.cell(row=r, column=1, value=k).font = Font(bold=True, size=10)
    ws.cell(row=r, column=1).border = BOX
    ws.cell(row=r, column=1).alignment = Alignment(vertical="top", wrap_text=True)
    c = ws.cell(row=r, column=2, value=v)
    c.font = Font(size=10); c.border = BOX
    c.alignment = Alignment(vertical="top", wrap_text=True)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
    ws.row_dimensions[r].height = 46
    r += 1

# hidden helper for the position chart
hr = r + 2
ws.cell(row=hr, column=1, value="力量")
ws.cell(row=hr, column=2, value="階段進度")
for i, d in enumerate(dash, start=1):
    ws.cell(row=hr + i, column=1, value=d[0].replace("\n", ""))
    ws.cell(row=hr + i, column=2, value=round(d[2] * 100))
ch = BarChart(); ch.type = "bar"; ch.style = 10
ch.title = "四大力量：本輪週期已走完的進度 (%)"
ch.y_axis.title = "%"; ch.height = 8; ch.width = 20
data = Reference(ws, min_col=2, min_row=hr, max_row=hr + 4)
cats = Reference(ws, min_col=1, min_row=hr + 1, max_row=hr + 4)
ch.add_data(data, titles_from_data=True); ch.set_categories(cats)
ch.dLbls = DataLabelList(); ch.dLbls.showVal = True
ch.legend = None
ws.add_chart(ch, f"A{r+2}")
for rr in range(hr, hr + 5):
    ws.row_dimensions[rr].hidden = True

# =====================================================================
# 01 宏觀流動性週期
# =====================================================================
ws = wb.create_sheet("01_宏觀流動性週期")
widths(ws, {"A":16,"B":10,"C":14,"D":13,"E":36,"F":14,"G":11,"H":36})
title_row(ws, 1, "① 宏觀流動性（息口）牛熊週期  1970–2026", 8)
note_row(ws, 2, "峰值利率 = 該收緊週期的聯邦基金利率高點；谷值利率 = 該寬鬆週期的低點。標普報酬為該區間近似總幅度。", 8)
header(ws, 4, ["週期", "起", "迄", "政策方向", "轉折觸發事件", "峰/谷利率(%)", "標普期間", "股市階段特徵"])
liq = [
 ["收緊", 1971, 1974, "Burns 收緊", "1973 阿拉伯石油禁運 → 停滯性通脹", 13.0, "-48%", "熊市：通脹+衰退雙殺，估值與盈利同跌"],
 ["寬鬆", 1975, 1976, "急速減息", "1974-10 標普見底、通脹回落", 4.75, "+66%", "牛市第一段：估值修復"],
 ["收緊", 1977, 1981, "Volcker 震撼療法", "1979 伊朗革命第二次油震 → 1979-08 Volcker 上任", 20.0, "-27%", "熊市：利率史上最高，但1980後段開始領先築底"],
 ["寬鬆", 1982, 1986, "解除通脹後大寬鬆", "1982-08 Volcker 轉向 → 世紀大牛市起點", 5.88, "+229%", "牛市主升段：估值大擴張"],
 ["收緊", 1987, 1989, "防過熱收緊", "1987-10-19 黑色星期一（單日 -22.6%）", 9.75, "-33%(單次)", "急跌但無衰退：流動性快速救援"],
 ["寬鬆", 1989, 1992, "S&L 危機 + 波灣戰", "1990-08 伊拉克入侵科威特 → 衰退", 3.0, "-20%", "熊市短促：1990-10 見底"],
 ["收緊", 1994, 1995, "預防式加息", "1994 債市大屠殺（債券史上最差年）", 6.0, "+1%", "股市橫行、債市重挫"],
 ["寬鬆", 1995, 1998, "軟著陸 + LTCM 救援", "1998-09 LTCM 倒閉 → 三次保險式減息", 4.75, "+117%", "牛市：軟著陸的教科書案例"],
 ["收緊", 1999, 2000, "打壓科網泡沫", "2000-03 納指見頂 → 科網爆破", 6.5, "-49%", "熊市：估值崩潰為主因"],
 ["寬鬆", 2001, 2003, "911 + 通縮防禦", "2003-03 伊拉克戰爭開打當日見底", 1.0, "+34%", "牛市重啟"],
 ["收緊", 2004, 2006, "循序漸進加息", "2006 樓價見頂 → 次按裂縫", 5.25, "+35%", "後期牛市：信貸驅動"],
 ["寬鬆", 2007, 2008, "GFC + QE1", "2008-09 雷曼倒閉", 0.25, "-57%", "史詩級熊市：信貸+去槓桿"],
 ["寬鬆", 2009, 2015, "ZIRP + QE2/3", "2013 縮減恐慌（Taper Tantrum）", 0.25, "+205%", "十年大牛市：流動性定價一切"],
 ["收緊", 2015, 2018, "正常化 + 縮表", "2018-Q4 縮表過度 → 單季 -19.8%", 2.50, "+22%", "牛市末段，Q4 急跌"],
 ["寬鬆", 2019, 2021, "保險式減息 + COVID 無限QE", "2020-03 COVID 崩盤 → 無限量QE", 0.25, "+114%", "V 型 + 泡沫化牛市"],
 ["收緊", 2022, 2023, "40年來最快加息", "2022 通脹 9.1% → 一年加 525bp", 5.50, "-25%", "熊市：估值壓縮，2022-10 見底"],
 ["寬鬆", 2024, 2025, "校準式減息", "2025-12 最後一次減息至 3.50–3.75% 後暫停", 3.625, "+50%", "牛市：AI 主導的窄幅上升"],
 ["轉鷹?", 2026, 0, "Warsh regime change", "2026-05 Warsh 上任；油價推高頭條CPI 至 3.4%；市場定價 60% 加息機率", 3.625, "區間", "★ 現在位置：寬鬆已完、收緊未確認的『真空期』"],
]
r = body(ws, 5, liq, wraps=(5,8))
for i in range(5, r):
    ws.row_dimensions[i].height = 30
    if ws.cell(row=i, column=2).value == 2026:
        for cc in range(1, 9):
            ws.cell(row=i, column=cc).fill = PatternFill("solid", fgColor=AMBER)
            ws.cell(row=i, column=cc).font = Font(size=10, bold=True, color=AMBERF)
    ws.cell(row=i, column=3).number_format = "0"
# chart data (fed funds peak/trough path)
hr = r + 2
ws.cell(row=hr, column=1, value="年"); ws.cell(row=hr, column=2, value="聯邦基金利率 (%)")
for i, row in enumerate(liq, start=1):
    ws.cell(row=hr + i, column=1, value=row[1])
    ws.cell(row=hr + i, column=2, value=row[5])
ch = LineChart(); ch.title = "① 聯邦基金利率週期高低點 1971–2026（流動性牛熊的骨架）"
ch.y_axis.title = "利率 %"; ch.x_axis.title = "週期起始年"
ch.height = 9; ch.width = 26; ch.style = 12
d = Reference(ws, min_col=2, min_row=hr, max_row=hr + len(liq))
c = Reference(ws, min_col=1, min_row=hr + 1, max_row=hr + len(liq))
ch.add_data(d, titles_from_data=True); ch.set_categories(c)
ws.add_chart(ch, f"A{r+2}")
for rr in range(hr, hr + len(liq) + 1):
    ws.row_dimensions[rr].hidden = True
rr = r + 21
title_row(ws, rr, "階段判讀", 8, 12)
ws.cell(row=rr+1, column=1, value=(
 "規律：過去 50 年，每一次『油價衝擊 + 央行被迫由鴿轉鷹』的組合都導致 12–24 個月內的熊市或深度修正"
 "（1973、1979、1990、2022）。唯一例外是 1998 年——當時油價是下跌的。\n"
 "現在：聯邦基金停在 3.50–3.75% 已 9 個月，寬鬆週期實質結束。9/16 FOMC 的點陣圖是本輪流動性週期方向的裁決點。\n"
 "領先指標優先序：① Brent 是否站穩 100 → ② 頭條CPI 與核心CPI 的差距是否擴大 → ③ 加息機率定價 → ④ 10年期是否破 5.0%。"))
ws.cell(row=rr+1, column=1).alignment = Alignment(vertical="top", wrap_text=True)
ws.merge_cells(start_row=rr+1, start_column=1, end_row=rr+1, end_column=8)
ws.row_dimensions[rr+1].height = 74

# =====================================================================
# 02 政治週期
# =====================================================================
ws = wb.create_sheet("02_政治週期")
widths(ws, {"A":12,"B":26,"C":16,"D":14,"E":16,"F":18,"G":40})
title_row(ws, 1, "② 政治週期（總統四年週期）：期中選舉年的牛熊型態  1962–2026", 7)
note_row(ws, 2, "期中選舉年是四年週期中最弱的一年，但其低點往往是下一輪主升段的起點。最大回撤為該年度內的高至低幅度（近似值）。", 7)
header(ws, 4, ["期中年", "當年主導事件", "年內最大回撤", "低點月份", "低點後12個月", "總統", "週期意義"])
mid = [
 [1962, "古巴導彈危機 + 甘迺迪打壓鋼價", -0.27, "1962-06", 0.33, "甘迺迪", "危機解除即為底部"],
 [1966, "越戰升級 + 信貸緊縮 (credit crunch)", -0.22, "1966-10", 0.33, "詹森", "緊縮見頂＝股市見底"],
 [1970, "越戰 + Penn Central 破產", -0.26, "1970-05", 0.44, "尼克遜", "流動性危機出清"],
 [1974, "石油禁運 + 水門事件 + 通脹", -0.37, "1974-10", 0.38, "尼克遜/福特", "★ 最接近今日的結構對標"],
 [1978, "通脹重臨 + 美元危機", -0.14, "1978-03", 0.13, "卡特", "弱勢反彈"],
 [1982, "Volcker 高息 + 衰退", -0.19, "1982-08", 0.58, "列根", "政策轉向＝世紀大牛起點"],
 [1986, "油價崩跌 + 減稅法案", -0.09, "1986-09", 0.35, "列根", "牛市延續"],
 [1990, "伊拉克入侵科威特 + 油價翻倍 + 衰退", -0.20, "1990-10", 0.29, "老布殊", "★ 開戰日即為底部"],
 [1994, "聯儲局倍增利率 + 債市大屠殺", -0.09, "1994-12", 0.34, "克林頓", "加息結束＝底部"],
 [1998, "亞洲金融風暴 + LTCM + 俄羅斯違約", -0.19, "1998-10", 0.24, "克林頓", "保險式減息救市"],
 [2002, "科網爆破尾聲 + 企業會計醜聞", -0.34, "2002-10", 0.34, "小布殊", "泡沫出清完成"],
 [2006, "樓市見頂前的最後一年", -0.08, "2006-06", 0.18, "小布殊", "假象平靜"],
 [2010, "歐債危機 + 閃電崩盤 + QE2", -0.16, "2010-07", 0.26, "奧巴馬", "QE2 宣示即底部"],
 [2014, "油價崩跌 + QE 結束", -0.07, "2014-10", 0.05, "奧巴馬", "淺回調"],
 [2018, "貿易戰 + 縮表過度", -0.20, "2018-12", 0.29, "特朗普", "聯儲局轉鴿即底部"],
 [2022, "烏克蘭戰爭 + 通脹9.1% + 最快加息", -0.25, "2022-10", 0.22, "拜登", "★ 通脹見頂＝股市見底"],
 [2026, "伊朗戰爭 + 荷姆茲封鎖 + 油價 +47% + 聯儲局轉鷹", None, "待定", None, "特朗普", "★ 現在位置：低點尚未確認"],
]
r = body(ws, 5, mid, wraps=(2,7))
for i in range(5, r):
    ws.cell(row=i, column=3).number_format = "0%"
    ws.cell(row=i, column=5).number_format = "0%"
    ws.row_dimensions[i].height = 22
for cc in range(1, 8):
    ws.cell(row=r-1, column=cc).fill = PatternFill("solid", fgColor=AMBER)
    ws.cell(row=r-1, column=cc).font = Font(size=10, bold=True, color=AMBERF)

stats_r = r + 1
title_row(ws, stats_r, "統計事實（1962–2022，16 個期中選舉年）", 7, 12)
sr = stats_r + 1
facts = [
 ["期中年平均最大回撤", "-19.5%", "四年週期中最深的一年"],
 ["低點後 12 個月平均漲幅", "+29.7%", "四年週期中最強的一段"],
 ["低點出現在 Q3–Q4 的比例", "11 / 16 = 69%", "9–12 月是築底的高機率窗口"],
 ["1950 年以來期中選舉後 12 個月上漲比例", "100%", "無一例外（樣本 19 次）"],
 ["四年週期各年平均報酬（標普，1950–）", "Y1 +6.7% / Y2 +3.3% / Y3 +16.8% / Y4 +6.6%", "第3年（2027）為最強"],
]
header(ws, sr, ["統計項目", "數值", "含意"]); sr += 1
for a, b, c_ in facts:
    ws.cell(row=sr, column=1, value=a).border = BOX
    ws.cell(row=sr, column=1).font = Font(size=10, bold=True)
    ws.cell(row=sr, column=2, value=b).border = BOX
    ws.cell(row=sr, column=2).font = Font(size=10, bold=True, color=REDF)
    ws.cell(row=sr, column=3, value=c_).border = BOX
    ws.merge_cells(start_row=sr, start_column=3, end_row=sr, end_column=7)
    sr += 1

ch = BarChart(); ch.type = "col"; ch.style = 11
ch.title = "② 期中選舉年：年內最大回撤 vs 低點後 12 個月漲幅"
ch.y_axis.title = "幅度"; ch.x_axis.title = "期中選舉年"
ch.height = 10; ch.width = 28
d = Reference(ws, min_col=3, max_col=3, min_row=4, max_row=r-2)
d2 = Reference(ws, min_col=5, max_col=5, min_row=4, max_row=r-2)
c = Reference(ws, min_col=1, min_row=5, max_row=r-2)
ch.add_data(d, titles_from_data=True); ch.add_data(d2, titles_from_data=True)
ch.set_categories(c)
ws.add_chart(ch, f"A{sr+2}")

jr = sr + 22
title_row(ws, jr, "階段判讀", 7, 12)
ws.cell(row=jr+1, column=1, value=(
 "現在位置：特朗普第二任期第 2 年、期中選舉年的 9 月——正是歷史上回撤最深、但也最接近轉折的位置。\n"
 "支持度由就任時 50%+ 跌至高 30%，眾議院預期易手、參議院 50-50，分裂政府機率高。\n"
 "分裂政府對市場的歷史含意：重大立法停擺（中性偏正面，政策不確定性下降），但債限與撥款攻防（10–12 月）會製造事件性波動。\n"
 "交易含意：11/3 之前偏防守、之後偏進攻。歷史上『期中選舉日』本身就是一個高勝率的做多起點。"))
ws.cell(row=jr+1, column=1).alignment = Alignment(vertical="top", wrap_text=True)
ws.merge_cells(start_row=jr+1, start_column=1, end_row=jr+1, end_column=7)
ws.row_dimensions[jr+1].height = 74

# =====================================================================
# 03 技術革命週期
# =====================================================================
ws = wb.create_sheet("03_技術革命週期")
widths(ws, {"A":18,"B":12,"C":12,"D":12,"E":30,"F":14,"G":40})
title_row(ws, 1, "③ 技術革命（產業）週期：萌芽 → 加速 → 狂熱 → 崩潰 → 實用化", 7)
note_row(ws, 2, "每一次技術革命的股市週期都遵循同一結構；差別只在時間長度與槓桿程度。崩潰幅度為該革命核心指數的高至低跌幅。", 7)
header(ws, 4, ["技術革命", "萌芽", "加速", "狂熱見頂", "見頂的觸發事件", "核心指數崩幅", "實用化/重建期特徵"])
tech = [
 ["鐵路（英國）", 1830, 1843, 1846, "1846 國會批出過量路線 + 資金鏈斷裂", "-65%", "1850s 鐵路實際運量爆發，但股東血本無歸"],
 ["電力/汽車/無線電", 1900, 1922, 1929, "1929-09 聯儲局收緊保證金 + 信貸緊縮", "-89%", "1930s–50s 技術普及但資本市場停滯 25 年"],
 ["電子/Nifty Fifty", 1955, 1965, 1972, "1973 石油禁運 + 估值 50 倍崩解", "-45%", "1970s 半導體實際落地，估值重設"],
 ["個人電腦", 1975, 1982, 1987, "1987-10 程式交易 + 估值過高", "-33%", "1990s PC 真正普及，微軟/英特爾主導"],
 ["互聯網/電訊", 1994, 1997, 2000, "2000-03 聯儲局加至 6.5% + 供應商融資爆煲", "-78%(納指)", "2003–07 寬頻普及，Google/Amazon 勝出"],
 ["移動/雲端", 2007, 2013, 2021, "2021-11 通脹爆發 + 零息結束", "-36%(納指)", "2022– SaaS 由成長轉現金流估值"],
 ["人工智能", 2023, 2025, None, "★ 尚未出現（見下方見頂訊號清單）", "—", "—"],
]
r = body(ws, 5, tech, wraps=(5,7))
for i in range(5, r):
    ws.row_dimensions[i].height = 26
for cc in range(1, 8):
    ws.cell(row=r-1, column=cc).fill = PatternFill("solid", fgColor=AMBER)
    ws.cell(row=r-1, column=cc).font = Font(size=10, bold=True, color=AMBERF)

cr = r + 1
title_row(ws, cr, "本輪 AI 週期的硬數據（資本開支是這輪革命的心跳）", 7, 12)
cr += 1
header(ws, cr, ["年份", "五大雲廠 AI 資本開支 (十億美元)", "按年增幅", "階段標記", "", "", ""]); cr += 1
capex = [[2023, 150, None, "萌芽：ChatGPT 引爆需求"],
         [2024, 238, 0.587, "加速：GPU 供不應求"],
         [2025, 410, 0.723, "加速：資料中心大規模動工"],
         [2026, 788, 0.922, "★ 高潮：增速仍在上升 = 未見頂"],
         [2027, 1000, 0.269, "預估：增速首次大幅放緩 ← 關鍵觀察點"]]
cr0 = cr
for row in capex:
    for i, v in enumerate(row, start=1):
        c = ws.cell(row=cr, column=i, value=v); c.border = BOX; c.font = Font(size=10)
    ws.cell(row=cr, column=3).number_format = "0.0%"
    cr += 1
ch = BarChart(); ch.type = "col"; ch.style = 12
ch.title = "③ AI 資本開支超級週期：增速尚未見頂 = 泡沫仍在膨脹階段"
ch.y_axis.title = "十億美元"; ch.height = 9; ch.width = 24
d = Reference(ws, min_col=2, min_row=cr0-1, max_row=cr-1)
c = Reference(ws, min_col=1, min_row=cr0, max_row=cr-1)
ch.add_data(d, titles_from_data=True); ch.set_categories(c)
ch.dLbls = DataLabelList(); ch.dLbls.showVal = True
ws.add_chart(ch, f"A{cr+2}")

sr = cr + 21
title_row(ws, sr, "AI 週期見頂訊號清單（逐項打勾，越多越接近 2000-03）", 7, 12)
sr += 1
header(ws, sr, ["訊號", "歷史對照", "現況", "", "", "", ""]); sr += 1
sig = [
 ["① 首家超大型雲廠下修 capex 指引", "2000-Q4 電訊商集體削減資本開支", "未出現"],
 ["② 資本開支由現金轉為債務融資", "1999 供應商融資（Lucent/Nortel）", "已部分出現：AI 資料中心發債激增"],
 ["③ 折舊年限被拉長以美化盈利", "2000 電訊業拉長折舊", "已出現爭議：GPU 折舊 3 年 vs 6 年"],
 ["④ 龍頭股市值佔指數比重見頂回落", "2000-03 思科成為全球最大市值後見頂", "未出現：NVDA 4.3 萬億美元仍在高位"],
 ["⑤ AI 相關信用利差擴闊", "1999-Q4 高收益電訊債利差先行擴闊", "需密切監察（本輪最靈敏的領先指標）"],
 ["⑥ 終端需求貨幣化不及預期", "2000 網絡流量增速不及預測", "混合訊號：應用增長強但 ROIC 未證實"],
]
for a, b, c_ in sig:
    ws.cell(row=sr, column=1, value=a).border = BOX; ws.cell(row=sr, column=1).font = Font(size=10, bold=True)
    ws.cell(row=sr, column=1).alignment = Alignment(wrap_text=True, vertical="top")
    ws.cell(row=sr, column=2, value=b).border = BOX; ws.cell(row=sr, column=2).font = Font(size=10)
    ws.cell(row=sr, column=2).alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=sr, start_column=2, end_row=sr, end_column=4)
    ws.cell(row=sr, column=5, value=c_).border = BOX; ws.cell(row=sr, column=5).font = Font(size=10)
    ws.cell(row=sr, column=5).alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=sr, start_column=5, end_row=sr, end_column=7)
    ws.row_dimensions[sr].height = 26
    sr += 1
sr += 1
title_row(ws, sr, "階段判讀", 7, 12)
ws.cell(row=sr+1, column=1, value=(
 "現在位置：對標互聯網週期的 1998–1999，而非 2000 年 3 月。理由：資本開支的『增速』仍在上升（2026 +92% 按年），"
 "而歷史上見頂永遠發生在增速轉負之後，不是絕對值高的時候。\n"
 "但風險在於：1999 年的聯儲局也是在加息（4.75%→6.5%），最終正是利率壓垮了估值。若 Warsh 真的加息，"
 "AI 這根唯一的支撐柱會同時面對『折現率上升』與『融資成本上升』雙重壓力。\n"
 "最靈敏的領先指標：AI 資料中心相關高收益債利差。它在 1999 年比納指提早約 4 個月轉向。"))
ws.cell(row=sr+1, column=1).alignment = Alignment(vertical="top", wrap_text=True)
ws.merge_cells(start_row=sr+1, start_column=1, end_row=sr+1, end_column=7)
ws.row_dimensions[sr+1].height = 76

# =====================================================================
# 04 黑天鵝地緣週期
# =====================================================================
ws = wb.create_sheet("04_黑天鵝地緣週期")
widths(ws, {"A":24,"B":12,"C":18,"D":14,"E":14,"F":14,"G":42})
title_row(ws, 1, "④ 黑天鵝（地緣事件 / 戰爭）週期：油震與戰爭的股市型態  1973–2026", 7)
note_row(ws, 2, "關鍵規律：地緣事件的股市底部，通常出現在『軍事行動確定性最高』的那一刻，而非事件解決之時。市場怕的是不確定性，不是戰爭本身。", 7)
header(ws, 4, ["事件", "起始", "油價變動", "標普最大跌幅", "見底時點", "見底至事件起始", "週期完結的訊號"])
geo = [
 ["1973 阿拉伯石油禁運", "1973-10", "$3 → $12 (+300%)", -0.48, "1974-10", "12 個月", "禁運解除 + 通脹見頂 + 聯儲局停止加息"],
 ["1979 伊朗革命", "1979-01", "$14 → $39 (+179%)", -0.17, "1980-03", "14 個月", "供應恢復 + Volcker 確立可信度"],
 ["1980 兩伊戰爭", "1980-09", "$39 → $35 (需求破壞)", -0.27, "1982-08", "23 個月", "需求破壞使油價自行崩跌"],
 ["1990 伊拉克入侵科威特", "1990-08", "$17 → $41 (+141%)", -0.20, "1990-10", "2 個月", "★ 空戰開打當日(1991-01-17)油價單日 -33%，股市起飛"],
 ["2001 九一一事件", "2001-09", "$28 → $22 (需求崩)", -0.12, "2001-09-21", "10 日", "市場重開後 19 個交易日收復失地"],
 ["2003 伊拉克戰爭", "2003-03", "$37 → $25", -0.14, "2003-03-11", "開戰前 9 日", "★ 開戰前見底：不確定性出清"],
 ["2022 俄烏戰爭", "2022-02", "$76 → $139 (+83%)", -0.25, "2022-10", "8 個月", "通脹見頂 + 歐洲避過能源斷供"],
 ["2026 伊朗戰爭 / 荷姆茲", "2026-02-28", "$66 → $97 (+47% 按年)", None, "待定", "已 6.3 個月", "★ 停火/伊朗石油復出，或需求破壞使 Brent < $80"],
]
r = body(ws, 5, geo, wraps=(1,7))
for i in range(5, r):
    ws.cell(row=i, column=4).number_format = "0%"
    ws.row_dimensions[i].height = 26
for cc in range(1, 8):
    ws.cell(row=r-1, column=cc).fill = PatternFill("solid", fgColor=AMBER)
    ws.cell(row=r-1, column=cc).font = Font(size=10, bold=True, color=AMBERF)

tr = r + 1
title_row(ws, tr, "2026 伊朗戰爭 / 荷姆茲危機時序", 7, 12); tr += 1
header(ws, tr, ["日期", "事件", "市場反應", "", "", "", ""]); tr += 1
tl = [
 ["2026-02-28", "美以突襲伊朗軍事與政府目標，最高領袖哈梅內伊被刺殺", "油價跳升，風險資產急挫"],
 ["2026-03-19", "伊朗封鎖荷姆茲海峽；美國展開空中戰役以重開海峽", "海峽油輪流量近乎歸零（全球 25% 海運原油）"],
 ["2026-03-27", "伊朗革命衛隊宣佈禁止往來美/以及其盟友港口的船隻", "油輪運費與保費暴漲"],
 ["2026-07-14", "美國重啟對伊朗的海上封鎖", "7 月中攻擊強度創 4 月以來新高"],
 ["2026-08 下旬", "胡塞武裝襲擊沙特能源設施，削減其產能", "Brent 單月 +11%"],
 ["2026-09 初", "美伊互擊油輪；伊朗向美軍艦發射彈道導彈", "Brent 9/3 見 $99.38，9/8 收 $97.41"],
 ["2026-09-09", "★ 現在：油輪戰持續，無停火跡象", "S&P 7,681 於 7,600–7,800 區間；VIX 僅 15.3"],
]
for a, b, c_ in tl:
    ws.cell(row=tr, column=1, value=a).border = BOX; ws.cell(row=tr, column=1).font = Font(size=10, bold=True)
    ws.cell(row=tr, column=2, value=b).border = BOX; ws.cell(row=tr, column=2).font = Font(size=10)
    ws.cell(row=tr, column=2).alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=tr, start_column=2, end_row=tr, end_column=4)
    ws.cell(row=tr, column=5, value=c_).border = BOX; ws.cell(row=tr, column=5).font = Font(size=10)
    ws.cell(row=tr, column=5).alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=tr, start_column=5, end_row=tr, end_column=7)
    ws.row_dimensions[tr].height = 24
    tr += 1

ch = BarChart(); ch.type = "col"; ch.style = 13
ch.title = "④ 地緣黑天鵝：標普最大跌幅（1973–2022）"
ch.y_axis.title = "最大跌幅"; ch.height = 9; ch.width = 26
d = Reference(ws, min_col=4, min_row=4, max_row=r-2)
c = Reference(ws, min_col=1, min_row=5, max_row=r-2)
ch.add_data(d, titles_from_data=True); ch.set_categories(c)
ws.add_chart(ch, f"A{tr+2}")

jr = tr + 20
title_row(ws, jr, "階段判讀", 7, 12)
ws.cell(row=jr+1, column=1, value=(
 "現在位置：事件已持續 6.3 個月，處於『持久消耗期』而非『初始衝擊期』。歷史上這一段最難交易——"
 "沒有新的壞消息，但也沒有解決方案，油價在高位橫行並持續侵蝕企業毛利與消費者購買力。\n"
 "兩條出路（決定 Q4 方向）：\n"
 "  (A) 供應面解決：停火、伊朗石油復出 → Brent 快速回落至 $70–80 → 頭條CPI 回落 → 聯儲局不需加息 → 股市大幅反彈（1991 型態）。\n"
 "  (B) 需求破壞：油價維持 $100+ 直至經濟衰退 → 盈利下修 → 股市再跌一段後才見底（1974/1980 型態）。\n"
 "分辨的關鍵指標：Brent 期貨曲線。深度逆價差(backwardation)＝實體短缺仍在；轉為正價差(contango)＝需求已被破壞，衰退在路上。"))
ws.cell(row=jr+1, column=1).alignment = Alignment(vertical="top", wrap_text=True)
ws.merge_cells(start_row=jr+1, start_column=1, end_row=jr+1, end_column=7)
ws.row_dimensions[jr+1].height = 108

# =====================================================================
# 05 領先指標儀表板
# =====================================================================
ws = wb.create_sheet("05_領先指標儀表板")
widths(ws, {"A":26,"B":16,"C":22,"D":8,"E":46,"F":14})
title_row(ws, 1, f"領先指標儀表板（LEADING INDICATORS）  資料截至 {ASOF}", 6)
note_row(ws, 2, "訊號：R = 看空/警示，A = 中性偏警惕，G = 看多/支撐。排序＝對後市影響力由高至低。", 6)
header(ws, 4, ["指標", "現值", "觸發閾值", "訊號", "解讀（為何是領先指標）", "更新頻率"])
li = [
 ["Brent 原油", "$97.41 (9/8)", "> $100 = 紅；< $80 = 綠", "R",
  "整條傳導鏈的源頭：油價 → 頭條CPI → 聯儲局立場 → 折現率 → 股票估值。目前是四大力量的實際驅動軸。", "即時"],
 ["Brent 按年變動", "+46.7%", "> +40% 歷來 12 個月內伴隨衰退", "R",
  "1973/1979/1990/2008/2022 五次油價按年 >40% 中，四次於 12 個月內出現衰退。這是目前最硬的警號。", "即時"],
 ["頭條CPI vs 核心CPI 差距", "3.4% vs 2.5% = +0.9pt", "> +1.0pt = 紅", "R",
  "差距全部來自能源。差距擴大＝油價正在滲入通脹預期，這正是逼使央行轉鷹的機制。8/CPI 於 9/11 公佈。", "月 (9/11)"],
 ["9/16 FOMC 加息機率", "約 60%", "> 50% = 紅", "R",
  "寬鬆週期是否正式結束的裁決。若點陣圖上移，2024–25 的估值擴張前提失效。", "即時"],
 ["10 年期美債孳息", "4.77%", "> 5.00% = 紅", "R",
  "破 5% 是估值壓縮的機械性觸發點；2023-10 與 2025 兩次觸及 5% 均引發 8%+ 修正。", "即時"],
 ["VIX", "15.30", "< 16 且油價 >$95 = 極度不對稱", "R",
  "低波動不等於低風險。目前尾部保護的成本處於歷史低位，而尾部風險處於高位——最值得利用的錯價。", "即時"],
 ["標普 500", "7,681", "支撐 7,600 / 阻力 7,800", "A",
  "區間收窄中。7,600 是本輪上升趨勢的最後防線，跌破則量化基金與 CTA 的機械式減倉會放大跌幅。", "即時"],
 ["黃金", "$4,480", "續創新高 = 貨幣/地緣避險需求", "A",
  "黃金與實質利率同步上升＝市場在對沖『央行失去控制』的尾部情境，而非單純的通脹交易。", "即時"],
 ["Brent 期貨曲線形態", "逆價差 (backwardation)", "轉 contango = 需求破壞確認", "A",
  "分辨『供應短缺型高油價』（可持續、通脹型）與『衰退前夕型』（需求崩、通縮型）的唯一可靠指標。", "日"],
 ["特朗普支持度", "約 38%", "< 40% = 分裂政府機率上升", "A",
  "決定 11/3 期中選舉結果與 Q4 債限/撥款的攻防強度；不確定性在 11/3 出清。", "週"],
 ["五大雲廠 AI 資本開支", "2026: $7,750–8,000 億", "指引下修 = 轉紅", "G",
  "目前四大力量中唯一的支撐。增速仍在上升（按年 +92%），未見任何見頂訊號。", "季"],
 ["AI 相關高收益債利差", "需監察", "較納指提早約 4 個月轉向", "A",
  "1999 年最靈敏的領先指標。AI 資本開支正由現金轉向債務融資，令這一指標的重要性在本輪大幅上升。", "日"],
]
r = body(ws, 5, li, wraps=(1,3,5))
for i in range(5, r):
    rag(ws, i, 4, li[i-5][3])
    ws.row_dimensions[i].height = 42
    ws.cell(row=i, column=2).font = Font(size=11, bold=True)
    ws.cell(row=i, column=2).alignment = Alignment(horizontal="center", vertical="center")

r += 1
title_row(ws, r, "傳導鏈（由左至右，每一環都是下一環的領先指標）", 6, 12)
ws.cell(row=r+1, column=1, value=(
 "荷姆茲海峽中斷  →  Brent +47% 按年  →  頭條CPI 3.4% > 核心 2.5%  →  通脹預期上移  →  "
 "Warsh 由觀望轉加息  →  10 年期破 5%  →  折現率上升  →  AI 成長股估值壓縮  →  標普跌破 7,600\n\n"
 "要打斷這條鏈，只需要打斷第一環：荷姆茲通航恢復。因此對日內交易者而言，"
 "『中東停火 / 復航 / OPEC+ 增產』的頭條，是目前市場上最高 beta 的單一催化劑——"
 "它會同時觸發：能源股急跌、成長股急升、債息急跌、VIX 崩塌。"))
ws.cell(row=r+1, column=1).alignment = Alignment(vertical="top", wrap_text=True)
ws.merge_cells(start_row=r+1, start_column=1, end_row=r+1, end_column=6)
ws.row_dimensions[r+1].height = 90

# =====================================================================
# 06 日內交易劇本
# =====================================================================
ws = wb.create_sheet("06_日內交易劇本")
widths(ws, {"A":18,"B":24,"C":26,"D":26,"E":40,"F":14})
title_row(ws, 1, f"日內交易劇本 (DAY TRADING PLAYBOOK)   {ASOF}", 6)
note_row(ws, 2, "本頁為框架與情境準備，非投資建議。所有價位為當時市場資料的近似值，執行前務必以即時報價覆核。", 6)

r = 4
title_row(ws, r, "A. 未來 8 週事件日曆（波動集中點）", 6, 12); r += 1
header(ws, r, ["日期", "事件", "為何重要", "偏向", "劇本", ""]); r += 1
cal = [
 ["9/11 (五)", "8 月 CPI", "頭條 vs 核心的差距是否擴大＝油價滲透的證據", "高波動雙向",
  "8:30 前不持倉。頭條 >3.6% → 做多波動、空成長；<3.2% → 空油多科技。"],
 ["9/16 (三)", "FOMC + 點陣圖 + 記者會", "本輪流動性週期方向的裁決點；市場定價 60% 加息", "最高風險日",
  "2:00 決議、2:30 記者會。經典型態：決議後首個方向常被 3:00 前反轉，勿追第一根。"],
 ["9/30 (三)", "季結 + 財政年度結束", "再平衡流量 + 政府撥款死線", "流量驅動",
  "月底/季底最後 30 分鐘的 MOC 失衡是可交易的。留意撥款僵局頭條。"],
 ["10/2 (五)", "9 月非農就業", "『油價侵蝕需求』是否開始出現在勞動力市場", "高波動",
  "此時就業轉弱＝壞消息就是壞消息（因為聯儲局被通脹綁住，不會來救）。"],
 ["10 月中", "Q3 業績期開始（金融股先行）", "油價對毛利的實際侵蝕首次入賬", "個股分化",
  "留意指引而非當季數字。運輸/零售/消費品的成本轉嫁能力是關鍵。"],
 ["10/27–28", "FOMC（無記者會月份）", "若 9 月未動，10 月為第二個窗口", "中",
  "無 SEP 的會議通常波動較低，除非 9 月留下懸念。"],
 ["11/3 (二)", "★ 期中選舉", "政治不確定性的出清點", "★ 結構性轉折",
  "歷史上 1950 年以來期中選舉後 12 個月標普 100% 上漲。選舉夜的恐慌是高勝率的做多位。"],
 ["全期 · 每週三", "EIA 原油庫存 10:30 ET", "荷姆茲中斷對實體庫存的影響", "油氣板塊日內主軸",
  "周二 16:30 API 先行；兩者背離時，10:30 的反向動能最強。"],
 ["全期 · 隨時", "中東頭條（停火/襲擊/OPEC+）", "目前市場最高 beta 的單一催化劑", "★ Gap 風險",
  "襲擊多發生在亞洲時段 → 原油期貨隔夜跳空 → 美股開盤前已定調。隔夜留倉需計算此風險。"],
]
for row in cal:
    ws.cell(row=r, column=1, value=row[0]).font = Font(size=10, bold=True)
    for i, v in enumerate(row, start=1):
        c = ws.cell(row=r, column=i, value=v); c.border = BOX
        c.alignment = Alignment(vertical="top", wrap_text=True)
        if i > 1: c.font = Font(size=10)
    ws.cell(row=r, column=6).border = BOX
    ws.row_dimensions[r].height = 40
    r += 1

r += 1
title_row(ws, r, "B. 關鍵價位地圖（標普 500）", 6, 12); r += 1
header(ws, r, ["層級", "價位", "性質", "被觸發後的含意", "對應交易", ""]); r += 1
lv = [
 ["阻力 2", "8,000", "整數關口 / 心理", "中東緩和 + CPI 降溫的組合才可能達到", "獲利了結區"],
 ["阻力 1", "7,800", "區間上緣", "突破需要新的正面催化劑（停火或鴿派 FOMC）", "區間上緣做空，除非放量突破"],
 ["現價", "7,681", "區間中軸", "無方向，低勝率區", "不交易區間中軸"],
 ["支撐 1", "7,600", "★ 區間下緣 / 趨勢最後防線", "跌破 = CTA 與波動率控制基金機械式減倉啟動", "跌破且回抽不上 → 做空至 7,400"],
 ["支撐 2", "7,400", "約 -3.7%，前次整固平台", "一般修正的第一目標", "第一目標，部分回補"],
 ["支撐 3", "7,300", "約 -5%，典型的 5% 回調", "常規回調的完成位", "第二目標"],
 ["支撐 4", "6,900–7,000", "約 -10%，期中選舉年的中位回撤", "對應期中年平均 -19.5% 回撤的一半", "戰略建倉的第一區"],
 ["支撐 5", "6,150–6,300", "約 -19%，期中年平均最大回撤", "對應 1974/1990/2022 型態的完整出清", "★ 戰略建倉的核心區"],
]
for row in lv:
    for i, v in enumerate(row, start=1):
        c = ws.cell(row=r, column=i, value=v); c.border = BOX; c.font = Font(size=10)
        c.alignment = Alignment(vertical="top", wrap_text=True)
    ws.cell(row=r, column=6).border = BOX
    if "7,600" in str(row[1]) or "6,150" in str(row[1]):
        for cc in range(1, 7):
            ws.cell(row=r, column=cc).fill = PatternFill("solid", fgColor=AMBER)
    ws.row_dimensions[r].height = 26
    r += 1

r += 1
title_row(ws, r, "C. 交易時段地圖（美東時間）與各時段的主導變數", 6, 12); r += 1
header(ws, r, ["時段", "美東時間", "主導變數", "典型行為", "日內策略", ""]); r += 1
sess = [
 ["亞洲盤", "20:00–03:00", "中東頭條、油輪襲擊、荷姆茲航運資料", "原油期貨跳空的主要形成時段", "隔夜倉必須以 CL 的隔夜區間計算風險；不留裸空 gamma"],
 ["倫敦開盤", "03:00–05:00", "歐洲能源與天然氣定價、Brent 現貨", "當日油價方向多在此定調", "以 Brent 的倫敦開盤方向作為當日 XLE / 航運股的偏向"],
 ["盤前", "08:00–09:30", "8:30 經濟數據（CPI/NFP/PPI）", "數據日的最大單根波動常在 08:30–08:35", "數據前平倉；數據後等第一個 5 分鐘 K 完成再進場"],
 ["開盤 30 分鐘", "09:30–10:00", "隔夜 gap 的消化、開盤區間建立", "開盤區間突破 (ORB) 的最佳樣本", "ORB 策略；但 FOMC/CPI 日的 ORB 假突破率顯著上升"],
 ["10:00 / 10:30", "10:00–10:30", "10:00 次級數據；周三 10:30 EIA", "油氣板塊當日主要動能點", "EIA 日：10:30 前不持有能源股方向倉"],
 ["午盤", "11:30–14:00", "流動性最低，趨勢最弱", "假突破高發時段", "縮小倉位或不交易；除非 FOMC 日"],
 ["FOMC 2:00", "14:00–15:00", "決議 + 點陣圖 + 記者會", "首個方向常於 3:00 前被反轉", "★ 不追第一根；等 14:45 後的第二方向"],
 ["尾盤", "15:00–16:00", "MOC 失衡、指數基金再平衡、Gamma", "月底/季底的失衡可交易", "15:50 公佈 MOC 失衡；配合期權到期日的 pin 效應"],
]
for row in sess:
    for i, v in enumerate(row, start=1):
        c = ws.cell(row=r, column=i, value=v); c.border = BOX; c.font = Font(size=10)
        c.alignment = Alignment(vertical="top", wrap_text=True)
    ws.cell(row=r, column=6).border = BOX
    ws.row_dimensions[r].height = 32
    r += 1

r += 1
title_row(ws, r, "D. 四力框架下的板塊與工具偏向", 6, 12); r += 1
header(ws, r, ["方向", "標的 / 板塊", "理由（對應哪一股力量）", "風險（什麼會令它失效）", "", ""]); r += 1
sec = [
 ["偏多", "能源 (XLE)、油服、油輪航運", "④ 黑天鵝：荷姆茲供應中斷 + 運費與保費暴漲", "停火或伊朗復出 → 單日可跌 8–10%"],
 ["偏多", "黃金 / 貴金屬", "① + ④：央行可信度風險 + 地緣避險，已創新高", "若 Warsh 加息且成功壓通脹 → 實質利率上升壓黃金"],
 ["偏多", "金融、醫療、工業、原材料", "① 高利率環境下的相對受惠者（券商共識偏好板塊）", "若轉為衰退敘事，金融的信貸成本會反噬"],
 ["偏多", "波動率 (VIX 買權 / 尾部保護)", "★ 最不對稱：VIX 15.3 對上 60% 加息機率與 $97 油價", "時間價值損耗；需嚴格控制在總資產 1–2%"],
 ["偏空", "消費非必需品、房地產", "① + ④：油價侵蝕可支配收入 + 高利率壓地產", "油價急跌則此交易立即失效"],
 ["偏空", "長天期美債（若 10 年期未破 5%）", "① 加息風險 + ④ 通脹的能源推力", "若轉衰退敘事，長債會由通脹資產變避險資產"],
 ["中性/持有", "AI 半導體與雲廠龍頭", "③ 唯一仍在擴張的力量，但同時對折現率最敏感", "利率與 capex 指引，任一轉向即需重估"],
]
for row in sec:
    for i, v in enumerate(row, start=1):
        c = ws.cell(row=r, column=i, value=v); c.border = BOX; c.font = Font(size=10)
        c.alignment = Alignment(vertical="top", wrap_text=True)
    lvl = {"偏多": GREEN, "偏空": RED}.get(row[0], AMBER)
    ws.cell(row=r, column=1).fill = PatternFill("solid", fgColor=lvl)
    ws.cell(row=r, column=1).font = Font(size=10, bold=True)
    ws.cell(row=r, column=5).border = BOX; ws.cell(row=r, column=6).border = BOX
    ws.row_dimensions[r].height = 30
    r += 1

r += 1
title_row(ws, r, "E. 日內風控規則（在目前這個特定的市場結構下）", 6, 12); r += 1
rules = [
 "1. 隔夜 gap 風險已結構性上升。中東襲擊多在亞洲時段發生，美股開盤前方向已定。日內倉位原則上不過夜；必須過夜則將部位縮至平日的 1/3。",
 "2. VIX 15.3 是陷阱而非安全訊號。低波動使停損距離看似很近，但一則頭條就能製造 3 個標準差的跳空。以 ATR 而非固定點數設停損。",
 "3. 9/11 與 9/16 兩天不做隔夜倉、不做方向性裸賣期權。這兩天的隱含波動溢價是合理的，不要當作免費的錢去賣。",
 "4. 7,600 是機械式的分水嶺。之上：區間交易（上緣空、下緣多）。跌破且回抽不上：轉為趨勢跟隨做空，不再逆勢接刀。",
 "5. 油價與股市的相關性目前是負的（通脹體制）。一旦轉為正相關（油跌股也跌），代表市場敘事已由『通脹』切換為『衰退』——這是整套劇本需要重寫的訊號。",
 "6. 不要在期中選舉年的 Q4 做結構性的長期空頭。歷史型態（16 次中 11 次於 Q3–Q4 見底，選後 12 個月 100% 上漲）與短線的空頭訊號方向相反；短空可以，長空是逆歷史機率。",
 "7. 單一頭條的最大殺傷力來源＝『中東停火』。持有能源多頭 + 成長股空頭的組合，在停火頭條下會兩邊同時虧損。務必分開計算這個關聯風險，不要當作對沖。",
]
for t in rules:
    c = ws.cell(row=r, column=1, value=t)
    c.font = Font(size=10); c.border = BOX
    c.alignment = Alignment(vertical="top", wrap_text=True)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    ws.row_dimensions[r].height = 32
    r += 1

r += 1
c = ws.cell(row=r, column=1, value=(
 "免責聲明：本文件為宏觀週期研究與交易框架整理，非投資建議、非買賣要約。所有數據為公開資料於 2026-09-09 的近似值，"
 "歷史規律不保證未來重演。日內交易涉及高風險並可能損失全部本金，執行前請以即時報價覆核所有價位並自行評估風險承受能力。"))
c.font = Font(size=9, italic=True, color="808080")
c.alignment = Alignment(vertical="top", wrap_text=True)
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
ws.row_dimensions[r].height = 40

# =====================================================================
# 07 資料來源
# =====================================================================
ws = wb.create_sheet("07_資料來源")
widths(ws, {"A":32,"B":70,"C":16})
title_row(ws, 1, "資料來源與查核日期", 3)
header(ws, 4, ["項目", "來源", "查核日"])
src = [
 ["聯邦基金利率 / FOMC", "Federal Reserve — FOMC statement 2026-07-29；H.15 Selected Interest Rates 2026-09-04", ASOF],
 ["Fed 主席立場", "CNBC / NPR / PBS — Warsh 於 Jackson Hole 2026-08-28 講話；federalreserve.gov 2026-07-14 國會證詞", ASOF],
 ["CPI", "BLS — Consumer Price Index Summary 2026 M07（7 月頭條 3.4%，核心 2.5%）；8 月數據 2026-09-11 公佈", ASOF],
 ["標普 500 / 納指水平", "Trading Economics；CNBC 市場報導 2026-09-01 至 2026-09-04", ASOF],
 ["油價", "Trading Economics（Brent 2026-09-08 = $97.41）；Fortune 每日油價 2026-09-01 至 09-04；Forbes Advisor", ASOF],
 ["10 年期孳息 / VIX / 黃金", "Trading Economics / FRED / Investing.com", ASOF],
 ["伊朗戰爭時序", "Britannica《2026 Iran war》；Al Jazeera 2026-09-06；ABC News 現場報導", ASOF],
 ["AI 資本開支", "MUFG Americas；AL Capital Advisory；CoBank Knowledge Exchange；ValueAdd VC（2026 五大雲廠 $775–800B）", ASOF],
 ["期中選舉與支持度", "Morgan Stanley；Charles Schwab 2026-08-20；Brookings；Capital Group", ASOF],
 ["地緣風險框架", "WEF Global Risks Report 2026；BlackRock Geopolitical Risk Dashboard；Wellington；Stimson Center", ASOF],
 ["歷史週期數據", "作者依公開歷史資料整理（聯邦基金利率、標普 500 回撤、油價衝擊、技術革命週期）", ASOF],
]
r = body(ws, 5, src, wraps=(1,2))
for i in range(5, r):
    ws.row_dimensions[i].height = 30

xlsx_path = os.path.join(OUT, f"{ASOF}_四大力量牛熊週期_日內交易框架.xlsx")
wb.save(xlsx_path)
print("XLSX ->", xlsx_path)
