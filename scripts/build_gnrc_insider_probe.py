"""GNRC 2026-09-16/17 Amazon 交易公布 — 事前資訊洩漏法證檢查 (Excel)。"""
from __future__ import annotations
import pathlib
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

NAVY = "1F3864"
GREEN, AMBER, RED, GREY = "C6EFCE", "FFEB9C", "FFC7CE", "D9D9D9"
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


def rows(ws, r, data, wrap=(), fills=None):
    for j, row in enumerate(data):
        for i, v in enumerate(row, 1):
            c = ws.cell(row=r, column=i, value=v)
            c.border = THIN
            c.alignment = Alignment(vertical="top", wrap_text=(i in wrap))
            if fills and fills[j] and i == 1:
                c.fill = PatternFill("solid", fgColor=fills[j])
        r += 1
    return r


def note(ws, r, text, span=8, h=32):
    c = ws.cell(row=r, column=1, value=text)
    c.font = Font(italic=True, size=9, color="555555")
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=span)
    ws.row_dimensions[r].height = h
    return r + 2


wb = Workbook()

# ------------------------------------------------------------------ 00
ws = wb.active; ws.title = "00_結論"
r = title_row(ws, 1, "GNRC 2026-09-16 Amazon 協議 — 公開紀錄中沒有資訊洩漏的跡象", 8, 15)
r += 1
r = rows(ws, r, [["六項法證標記逐項檢查，沒有一項指向事前洩漏。"
                  "最關鍵的一項是：協議簽署當日（9/16）正常交易時段收 $175.07，"
                  "對比前收 $175.11 幾乎完全持平 —— 事前搶跑必然在收巿前留下量價足跡，這裡沒有。"]])
ws.merge_cells(start_row=r-1, start_column=1, end_row=r-1, end_column=8)
ws.row_dimensions[r-1].height = 40
r += 1
ws.cell(row=r, column=1, value="標的").font = Font(bold=True)
ws.cell(row=r, column=2, value="GNRC")
ws.cell(row=r, column=3, value="Generac Holdings（NYSE）")
ws.cell(row=r, column=4, value="交易對手")
ws.cell(row=r, column=5, value="AMZN")
r += 2
r = header(ws, r, ["#", "法證標記", "查到什麼", "判讀"], [5, 24, 52, 16])
data = [
    ["1", "公布前價格漂移", "9/16 正常時段收 $175.07，前收 $175.11 —— 持平。"
          "公布在收巿後，股價其後於盤後最多 +45%", "無跡象"],
    ["2", "內部人 Form 4 方向", "公布前數週的申報全部是「賣出」，非買入。"
          "CEO Jagdfeld 9/1 賣 5,000 股 @ $182.21；"
          "Home 總裁 Taffe 9/1 賣 3,932 股 @ $181.22、9/8 再賣數筆 —— 均為 10b5-1 預定計劃",
     "反向，無跡象"],
    ["3", "異常期權活動", "傳媒引述的期權流是 9/17 上午 10:57、股價已在 $204 時的數據，"
          "且集中在翌日（9/18）到期的價外 call —— 那是公布後追價，不是事前部署", "時序不成立"],
    ["4", "事前傳聞 / 媒體洩漏", "搜尋不到 9/16 之前任何關於 Amazon 為交易對手的報道或傳聞", "無跡象"],
    ["5", "資訊其實早已公開", "Q2 業績電話會議已公開披露：季內簽下兩份 hyperscale 供應協議，"
          "第一份約 $700M（2027 交付）；第二份因產品條款未定案而未計入 $1.6B backlog，"
          "並明言「至少與第一份同樣大」、涵蓋 2027 與 2028 —— 交易的存在不是秘密",
     "不需要內幕"],
    ["6", "披露速度", "9/16 簽署，9/16 收巿後公布、9/17 遞交 8-K。"
          "8-K 法定期限為 4 個工作日，此處近乎即時 —— 洩漏窗口被壓到最小", "無跡象"],
]
fills = [GREEN, GREEN, GREEN, GREEN, GREEN, GREEN]
r = rows(ws, r, data, wrap=(3,), fills=fills)
for rr in range(r - len(data), r):
    ws.row_dimensions[rr].height = 46
r += 1
r = title_row(ws, r, "真正的「意外」是什麼", 8, 12)
r = rows(ws, r, [
    ["已公開（Q2 就知道）", "第二個 hyperscale 客戶存在、合約已簽、規模至少 $700M、2027–2028 交付"],
    ["9/16 才揭曉", "交易對手是 Amazon；初期規模 $2.4B（即約為 $700M 參照值的 3.4 倍）；"
                    "上限 $8B；附帶認股權證 169 萬股、行使價 $200.93（當時股價約 $175，溢價約 15%）"],
], wrap=(2,))
r = note(ws, r, "換句話說：市場不是不知道有這件事，而是把規模估少了 3 倍以上。"
                "這是「公開但被低估」的資訊，屬於研究優勢，不是內幕。", 8)

# ------------------------------------------------------------------ 01
ws = wb.create_sheet("01_時間線")
r = title_row(ws, 1, "事件時間線", 4)
r = header(ws, r, ["日期", "事件", "股價 / 動作", "來源性質"], [14, 46, 24, 16])
r = rows(ws, r, [
    ["2026-06-24", "與第二個 hyperscale 客戶簽約（當時未具名）", "—", "公司披露"],
    ["Q2 業績", "公開說明第二份協議「至少與第一份（$700M）同樣大」，"
                "因條款未定案未計入 $1.6B backlog", "—", "公司披露"],
    ["2026-09-01", "CEO Jagdfeld 賣 5,000 股（10b5-1）", "@ $182.21", "Form 4"],
    ["2026-09-01", "Home 總裁 Taffe 賣 3,932 股", "@ $181.22", "Form 4"],
    ["2026-09-08", "Taffe 再賣數筆並行使選擇權", "@ $188–190", "Form 4"],
    ["2026-09-16", "Wells Fargo 指 Generac 大致豁免於新行政命令，"
                   "重申 Overweight、目標價 $280", "盤中利好", "賣方研究"],
    ["2026-09-16", "與 Amazon 正式簽署長期供應協議", "正常時段收 $175.07（持平）", "公司"],
    ["2026-09-16 盤後", "消息公布", "盤後最多 +45%", "公司 / 通訊社"],
    ["2026-09-17", "遞交 8-K；股價開 $229.50，盤中見約 $231.89，收 $207.23",
     "當日 +27% 至 +33.7%，上巿以來最大單日升幅", "SEC / 巿場"],
], wrap=(2,))
r = note(ws, r, "注意 9/16 的兩件事同日發生：Wells Fargo 的監管風險解讀（公開研究）"
    "與 Amazon 協議簽署（當時未公開）。前者可以解釋 9/16 盤中的任何波動，"
    "而當日最終仍收平 —— 這反而進一步削弱「有人事前知情並買入」的推論。", 4, 44)

# ------------------------------------------------------------------ 02
ws = wb.create_sheet("02_查不到的部分")
r = title_row(ws, 1, "以下項目本次無法查證 —— 結論的邊界在這裡", 3)
r = header(ws, r, ["未能查證項目", "為何重要", "受阻原因"], [30, 42, 26])
r = rows(ws, r, [
    ["9/8–9/16 成交量 vs 20 日均量", "價格持平但成交量放大，仍可能是吸貨。"
                                     "價格無漂移只排除了「推高價格的買盤」，未排除「不推價的吸納」",
     "所有行情 API 被網絡政策封鎖（403）"],
    ["9/16 之前的期權未平倉量變化", "真正的事前部署會出現在 10 月／11 月到期的 call，"
                                   "而非 9/18 到期的當週 call", "同上"],
    ["Form 4 原始文件", "本檔的內部人交易資料來自搜尋摘要，並非逐份閱讀原文。"
                        "另外 Form 4 有 2 個工作日申報期 —— 9/15–16 的交易未必已入索引",
     "sec.gov 被網絡政策封鎖"],
    ["8-K 的精確受理時間戳", "可界定洩漏窗口的確切長度", "sec.gov 被封鎖"],
    ["13F / 13D-G 機構持倉變化", "季度申報，滯後太久，對單一事件無診斷價值", "本質限制"],
    ["實際下單者身分", "只有 SEC / FINRA 能看到成交歸屬。公開數據永遠只能看到「異常」，"
                       "看不到「是誰」", "本質限制"],
], wrap=(2, 3))
r = note(ws, r, "誠實的結論邊界：以上說的是「在我能查到的公開紀錄中沒有跡象」，"
    "不等於「確定沒有發生」。成交量與期權未平倉量這兩項若日後能取得，"
    "是最有機會改變結論的兩個變數。", 3, 40)

# ------------------------------------------------------------------ 03
ws = wb.create_sheet("03_可重用檢查表")
r = title_row(ws, 1, "下次遇到爆升，按這六項查 —— 附判讀門檻", 5)
r = header(ws, r, ["#", "檢查項", "怎麼量", "可疑門檻", "備註"],
           [5, 22, 34, 24, 28])
r = rows(ws, r, [
    ["1", "公布前價格漂移", "公布前 1–10 個交易日的累積異常報酬（扣除同業與大盤）",
     "CAR > +5% 且無對應公開消息", "最直接的標記"],
    ["2", "公布前成交量", "公布前 1–10 日成交量 ÷ 20 日均量",
     "連續數日 > 2 倍且價格未動", "價平量增最可疑"],
    ["3", "期權未平倉量", "看「到期日在公布日之後一個月以上」的價外 call OI 增幅",
     "OI 單週增逾 3 倍", "當週到期的 call 多半是事後追價"],
    ["4", "內部人 Form 4", "公布前 90 日的買賣方向與金額；是否 10b5-1",
     "非 10b5-1 的集中買入", "賣出反而是反證"],
    ["5", "資訊是否已公開", "翻上一季業績電話會議與 10-Q 有無預告",
     "完全無預告才算突發", "最常被忽略的一項"],
    ["6", "披露速度", "簽署日到 8-K 遞交日的間隔（法定 4 個工作日）",
     "拖近上限且期間有異動", "近乎即時 = 窗口極小"],
], wrap=(3, 4, 5))
r += 1
r = note(ws, r, "兩個必須守住的判讀紀律：\n"
    "① 時序先於一切。任何「異常期權活動」報道都要先確認時間戳在公布之前 —— "
    "本案傳媒引述的期權流全部發生在公布後，卻很容易被誤讀成事前訊號。\n"
    "② 異常 ≠ 證據。公開數據只能顯示統計異常，無法識別下單者。"
    "依 Form 4 公開申報的交易屬合法披露行為，與內幕交易是兩回事，不應混為一談。", 5, 62)

out = pathlib.Path("output/2026-09-18_GNRC_Amazon公布前資訊洩漏法證檢查.xlsx")
wb.save(out)
print("已輸出", out)
