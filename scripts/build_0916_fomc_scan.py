# -*- coding: utf-8 -*-
"""2026-09-16 FOMC 決議日盤中掃描（美東 09:53，開市 23 分鐘）"""
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, Reference

OUT=os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),"output")
os.makedirs(OUT,exist_ok=True)
STAMP="2026-09-16 09:53 ET（港時 21:53）— 開市 23 分鐘 | FOMC 決議日"

NAVY="1F3864"; SLATE="2F5597"
RED="FFC7CE"; AMBER="FFEB9C"; GREEN="C6EFCE"
REDF="9C0006"; AMBERF="9C6500"; GREENF="006100"
THIN=Side(style="thin",color="BFBFBF"); BOX=Border(left=THIN,right=THIN,top=THIN,bottom=THIN)

def title_row(ws,r,t,span,size=14):
    c=ws.cell(row=r,column=1,value=t); c.font=Font(bold=True,size=size,color="FFFFFF")
    c.fill=PatternFill("solid",fgColor=NAVY); c.alignment=Alignment(vertical="center",wrap_text=True)
    ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=span); ws.row_dimensions[r].height=26
def note_row(ws,r,t,span,h=36):
    c=ws.cell(row=r,column=1,value=t); c.font=Font(italic=True,size=9,color="595959")
    c.alignment=Alignment(wrap_text=True,vertical="top")
    ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=span); ws.row_dimensions[r].height=h
def header(ws,r,hs,h=30):
    for i,x in enumerate(hs,1):
        c=ws.cell(row=r,column=i,value=x); c.font=Font(bold=True,size=10,color="FFFFFF")
        c.fill=PatternFill("solid",fgColor=SLATE); c.border=BOX
        c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
    ws.row_dimensions[r].height=h
def widths(ws,spec):
    for k,v in spec.items(): ws.column_dimensions[k].width=v
def put(ws,r,row,wrap=(),height=26,fill=None):
    for i,v in enumerate(row,1):
        c=ws.cell(row=r,column=i,value=v); c.border=BOX; c.font=Font(size=10)
        c.alignment=Alignment(vertical="top",wrap_text=(i in wrap))
        if fill: c.fill=PatternFill("solid",fgColor=fill)
    ws.row_dimensions[r].height=height
def block(ws,r,text,span,fill=None,size=10.5,h=100):
    c=ws.cell(row=r,column=1,value=text); c.font=Font(size=size); c.border=BOX
    c.alignment=Alignment(vertical="top",wrap_text=True)
    if fill: c.fill=PatternFill("solid",fgColor=fill)
    ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=span)
    ws.row_dimensions[r].height=h

wb=Workbook()

# ================= 00 結論 =================
ws=wb.active; ws.title="00_掃描結論"
widths(ws,{"A":22,"B":14,"C":14,"D":52})
title_row(ws,1,f"盤中掃描：波幅小 + 10 分鐘均線平穩向上　|　{STAMP}",4)

title_row(ws,3,"★ 結論一：現在這個時點，這個掃描做不出可靠答案（三個獨立原因）",4,12)
header(ws,4,["原因","數值","嚴重程度","說明"])
reasons=[
 ["① 樣本不足","2 根 K 線","致命",
  "09:30 開市至 09:53 只有 2 根完成的 10 分鐘 K 線（09:30-09:40、09:40-09:50）。\n"
  "10 期的 10 分鐘均線需要 100 分鐘資料 → 最早 11:10 ET 才成形；\n"
  "即使只用 5 期也需要 50 分鐘 → 10:20 ET。\n"
  "用 2 個點去判斷「平穩向上」，任何股票都會呈現一條直線 —— 這是統計上的無意義。"],
 ["② 週期由 20 分鐘縮短至 10 分鐘","噪音約 ×1.4","高",
  "同一段走勢改用一半的取樣週期，每根 K 的隨機波動約放大 √2 倍。\n"
  "而開市首 30 分鐘本來就是全日波動最高、訊噪比最低的時段（開盤區間通常要 15–30 分鐘才確立）。\n"
  "兩者疊加 = 現在是全日最不適合用 10 分鐘均線做判斷的時刻。"],
 ["③ 今日 14:00 是 FOMC 決議","92.9% 加息機率","致命",
  "市場定價 92.9% 加息 25bp 至 3.75–4.00% —— 這是 2023 年以來首次加息。\n"
  "14:00 決議 + 點陣圖，14:30 Warsh 記者會。\n"
  "09:30–14:00 之間形成的任何均線趨勢，都是薄量的會前漂移，14:00 會被整段抹掉。\n"
  "★ 「FOMC 上午平穩上行的均線」正正是最典型的陷阱形態。"],
]
r=5
for row in reasons:
    put(ws,r,row,wrap=(1,4),height=max(56,16*row[3].count("\n")+42))
    ws.cell(row=r,column=3).fill=PatternFill("solid",fgColor=RED if row[2]=="致命" else AMBER)
    ws.cell(row=r,column=3).font=Font(size=10,bold=True,color=REDF if row[2]=="致命" else AMBERF)
    ws.cell(row=r,column=3).alignment=Alignment(horizontal="center",vertical="center")
    ws.cell(row=r,column=2).alignment=Alignment(horizontal="center",vertical="center")
    ws.cell(row=r,column=2).font=Font(size=11,bold=True)
    r+=1

r+=1
title_row(ws,r,"★ 結論二：即使撇開時點問題，今日常見的候選也幾乎全部失效",4,12); r+=1
header(ws,r,["類別","代表","狀態","為何今日不符合"]); r+=1
cand=[
 ["半導體","NVDA / AVGO / AMD","✘ 失效",
  "昨日費城半導體指數 -5.9%（NVDA -3.36%、AVGO 與 AMD 均 -4% 以上）。\n"
  "今日 NVDA 再跌逾 2% —— 中國網信辦要求大型企業終止 RTX Pro 6000D 訂單。趨勢向下，非向上。"],
 ["超大型科技 / 雲廠","MSFT / GOOGL / AMZN / ORCL","✘ 失效",
  "近日被借貸成本上升打擊（GOOGL -1.3%、MSFT -1.6%、AMZN -2%、ORCL -3.1%）。\n"
  "ORCL 另有新一輪裁員消息，跌至約 $141 的兩週低位。AI 雲廠對利率最敏感，而今日正是加息日。"],
 ["黃金","GOLD","✘ 失效","已跌穿 $4,300（9/9 時為約 $4,480）。方向不符。"],
 ["小型股","RUT","✘ 失效","今早 -0.76%，與標普 +0.25% 背離 —— 這正是加息定價的典型型態。方向不符。"],
 ["能源綜合型","XOM / CVX","△ 唯一仍成立",
  "Brent 已升至 $106.47、WTI $101.88 —— 四大力量框架中唯一仍在持續推進的資金流。\n"
  "★ 但今日仍不建議：① 油價已極度延伸（9/8 時 $97 → 現在 $106）；\n"
  "② 加息對石油需求是負面，14:00 之後可能反轉；③ 中東襲擊升級的頭條風險是雙向的。"],
]
for row in cand:
    put(ws,r,row,wrap=(2,4),height=max(40,16*row[3].count("\n")+36))
    f=GREEN if "唯一" in row[2] else RED
    ws.cell(row=r,column=3).fill=PatternFill("solid",fgColor=AMBER if "唯一" in row[2] else RED)
    ws.cell(row=r,column=3).font=Font(size=10,bold=True,color=AMBERF if "唯一" in row[2] else REDF)
    ws.cell(row=r,column=3).alignment=Alignment(horizontal="center",vertical="center")
    r+=1

r+=1
title_row(ws,r,"★ 結論三：標普 500 正好坐在 7,600 這條關鍵支撐上",4,12); r+=1
block(ws,r,(
 "標普 500 現報約 7,601。這正是我在 9/9 框架中標示的「本輪上升趨勢最後防線」—— 當時距離 0.72%，現在是 0.0%。\n\n"
 "這條線的意義：跌破後，CTA 與波動率控制基金的機械式減倉會啟動，令跌勢自我放大。\n"
 "而今日 14:00 的 FOMC 決議，正是決定它守不守得住的那個催化劑。\n\n"
 "★ 交易含意：現在不是在選股的時候，是在等一個二元結果。指數坐在支撐上 + 4 小時後有決議 "
 "= 任何方向性倉位都是在賭 FOMC，而不是在做趨勢跟隨。"),4,fill=AMBER,size=11,h=118); r+=2

title_row(ws,r,"那應該怎麼做？",4,12); r+=1
block(ws,r,(
 "1. 若一定要今日交易 → 等 14:30 記者會之後。FOMC 日真正可交易的趨勢通常在 15:00 之後才成形，\n"
 "   而且那才是有成交量支撐的走勢。上午的「平穩上行」是假的。\n\n"
 "2. 若要現在就掃描 → 最早 11:10 ET（10 期 10 分鐘均線成形）再跑，而且要接受它只有 3 小時壽命。\n\n"
 "3. 更實際的做法 → 今日不做趨勢跟隨，改為在 13:55 前完全平倉，14:30 後重新評估。\n"
 "   這是 2023 年以來第一次加息，點陣圖的訊息量遠大於加息本身，市場需要時間消化。\n\n"
 "4. 若堅持要一個名字 → XOM（能源綜合型，油價 $106 的唯一持續性資金流，波幅低於 E&P），\n"
 "   但倉位要當作「FOMC 前的短線」而非「趨勢倉」，13:55 前必須離場。"),4,size=11,h=150)

# ================= 01 市場狀態 =================
ws=wb.create_sheet("01_今日市場狀態")
widths(ws,{"A":22,"B":16,"C":62})
title_row(ws,1,f"今日市場狀態（已核實）　|　{STAMP}",3)
header(ws,3,["項目","數值","說明"])
mkt=[
 ["S&P 500","7,601 (+0.25%)","★ 正好坐在 7,600 關鍵支撐上"],
 ["道指","+0.15%",""],
 ["納指","+0.44%","表面領先，但 NVDA 跌逾 2%，升幅來自非半導體板塊"],
 ["羅素2000","-0.76%","★ 與大型股背離 = 加息定價的典型型態"],
 ["FOMC 加息機率","92.9%","CME FedWatch。加 25bp 至 3.75–4.00%，2023 年以來首次"],
 ["Brent 原油","$106.47 (+2.07%)","中東襲擊升級。9/8 時為 $97，8 天內 +9.8%"],
 ["WTI 原油","$101.88 (+1.9%)",""],
 ["10 年期美債孳息","19 年高位（2007 年以來）","★ 已突破我框架中 5.0% 的估值壓縮觸發點"],
 ["黃金","跌穿 $4,300","9/9 時為約 $4,480。實質利率上升壓黃金"],
 ["NVDA","跌逾 2%","中國網信辦要求終止 RTX Pro 6000D 訂單"],
 ["費城半導體指數（昨日）","-5.9%","NVDA -3.36%、AVGO 與 AMD 均 -4% 以上"],
 ["ORCL","約 $141","兩週低位，新一輪裁員"],
]
r=4
for row in mkt:
    put(ws,r,row,wrap=(3,),height=22)
    ws.cell(row=r,column=2).font=Font(size=10,bold=True)
    if row[0] in ("S&P 500","羅素2000","10 年期美債孳息","FOMC 加息機率"):
        for cc in range(1,4): ws.cell(row=r,column=cc).fill=PatternFill("solid",fgColor=AMBER)
    r+=1

r+=1
title_row(ws,r,"四大力量框架更新（對比 9/9）",3,12); r+=1
header(ws,r,["力量","9/9 → 9/16","變化"]); r+=1
upd=[
 ["① 宏觀流動性","90% → 100%","★ 本輪寬鬆週期正式結束。92.9% 機率今日加息 = 我在 9/9 標示的「裁決點」正在以加息落實。"
  "框架由「寬鬆已完、收緊未確認的真空期」進入「確認再收緊」。"],
 ["② 政治週期","75% → 78%","期中選舉 11/3，距今 48 天。無實質變化。"],
 ["③ 技術革命","65% → 68%","★ 首次出現裂縫：半導體單日 -5.9%，觸發點是 AI 安全疑慮 + 中國禁令，"
  "而非我原先預期的 capex 指引下修。雲廠同時被借貸成本打擊 —— 這正是我 9/9 警告的「折現率上升 + 融資成本上升」雙重壓力。"],
 ["④ 黑天鵝","55% → 62%","Brent 由 $97 升至 $106（8 天 +9.8%），中東襲擊升級。仍在持久消耗期，未見解決催化劑。"],
]
for row in upd:
    put(ws,r,row,wrap=(3,),height=44)
    ws.cell(row=r,column=2).font=Font(size=10,bold=True)
    ws.cell(row=r,column=2).alignment=Alignment(horizontal="center",vertical="center")
    r+=1
r+=1
block(ws,r,(
 "★ 傳導鏈已經走完全程：荷姆茲中斷 → Brent $106 → 通脹預期上移 → Warsh 加息（今日）→ "
 "10 年期 19 年高位 → 折現率上升 → AI 雲廠與半導體同步下跌 → 標普跌至 7,600。\n"
 "我在 9/9 描述的這條鏈，7 天內每一環都已實現。現在的問題不再是「會不會發生」，而是「7,600 守不守得住」。"),3,fill=AMBER,size=11,h=70)

# ================= 02 FOMC 日時間表 =================
ws=wb.create_sheet("02_FOMC日時間表")
widths(ws,{"A":14,"B":26,"C":60})
title_row(ws,1,"FOMC 決議日：剩餘時段時間表（美東時間）",3)
note_row(ws,2,"FOMC 日有極其固定的日內結構。認識它比選股更重要。",3,20)
header(ws,4,["時間","事件","行為特徵與應對"])
tl=[
 ["09:53","★ 現在","開市 23 分鐘。只有 2 根 10 分鐘 K 線，均線未成形。開盤區間尚未確立。"],
 ["10:00–11:00","開盤區間建立","FOMC 日的開盤區間比平日窄。此時的方向多為隔夜期貨的延續，非當日真實方向。"],
 ["11:10","10 期 10 分鐘均線成形","★ 最早可以做這個掃描的時點。但仍要接受它只有約 3 小時壽命。"],
 ["11:30–13:30","會前真空期","全日成交量最低。價格傾向窄幅橫行或緩慢漂移。\n★ 這段的「平穩上行均線」是薄量產物，不代表買盤，是 FOMC 日最大的陷阱。"],
 ["13:55","★ 最後離場時點","若不打算賭 FOMC，應在此前完全平倉。"],
 ["14:00","決議 + 點陣圖","★ 92.9% 已定價加息，所以加息本身不是新聞 —— 點陣圖（2027 年的路徑）才是。\n"
  "首個 1–2 分鐘的方向經常是錯的（演算法先讀標題）。"],
 ["14:30","Warsh 記者會","★ 真正的波動來源。他上任後首次加息投票，措辭定調後市。\n"
  "經典型態：14:00 的方向在 14:30–15:00 之間被反轉。不要追 14:00 的第一根。"],
 ["15:00–16:00","真實趨勢形成","FOMC 日唯一有成交量支撐的趨勢段。若要做趨勢跟隨，這才是時候。"],
 ["盤後 / 隔夜","中東頭條","Brent $106，襲擊多在亞洲時段發生 → 隔夜跳空風險仍然最高。"],
]
r=5
for row in tl:
    put(ws,r,row,wrap=(3,),height=max(22,16*row[2].count("\n")+24))
    ws.cell(row=r,column=1).font=Font(size=10,bold=True)
    if row[0] in ("09:53","13:55","14:30","11:10"):
        for cc in range(1,4): ws.cell(row=r,column=cc).fill=PatternFill("solid",fgColor=AMBER)
    r+=1

r+=1
title_row(ws,r,"標普 500 關鍵價位（FOMC 後的兩條路）",3,12); r+=1
header(ws,r,["情境","價位","含意"]); r+=1
for a,b,c_ in [
 ["守住 7,600","7,600 → 7,800","鴿派加息（加息但點陣圖顯示到頂）→ 反彈回區間上緣。這是多頭僅存的劇本。"],
 ["★ 跌破 7,600","7,400","CTA 與波動率控制基金機械式減倉啟動。第一目標，約 -2.6%。"],
 ["續跌","7,300","約 -4%，典型回調完成位。"],
 ["深度修正","6,900–7,000","約 -9%，戰略建倉第一區。"],
 ["期中年型態","6,150–6,300","約 -19%，對應期中選舉年平均最大回撤。★ 戰略建倉核心區。"],
]:
    put(ws,r,[a,b,c_],wrap=(3,),height=22)
    ws.cell(row=r,column=1).font=Font(size=10,bold=True)
    if "7,600" in a: 
        for cc in range(1,4): ws.cell(row=r,column=cc).fill=PatternFill("solid",fgColor=RED)
    r+=1

# ================= 03 限制與來源 =================
ws=wb.create_sheet("03_限制與來源")
widths(ws,{"A":24,"B":76})
title_row(ws,1,"資料限制與來源",2)
items=[
 ["★ 資料限制","本執行環境的網絡政策仍然封鎖全部行情 API（Yahoo Finance、Stooq、Polygon、Alpha Vantage、"
  "TwelveData、marketdata.app、Nasdaq、TradingView 全部回 403），本次已再度逐一測試確認。\n"
  "因此我拿不到逐分鐘 K 線，無法計算 10 分鐘均線。\n"
  "本檔案中：指數與個股變動、油價、加息機率、各項事件 —— 均為已核實的公開資料；"
  "波幅與趨勢判斷為推斷。但請注意：本次的核心結論（時點不足、FOMC 日結構、候選失效）"
  "並不依賴逐分鐘資料，而是依賴時間算術與已核實的板塊狀態，因此結論本身是穩固的。"],
 ["如何取得實測結果","scripts/screen_20min_ma.py 支援自訂重採樣週期。改用 10 分鐘：\n"
  "  在 ma_quality() 中把 bar=\"20min\" 改為 \"10min\"，或直接以參數傳入。\n"
  "  python3 scripts/screen_20min_ma.py --date 2026-09-16 --top 20 --out scan.csv\n"
  "建議 11:10 ET 之後才執行，否則樣本不足。"],
 ["為何 10 分鐘比 20 分鐘難用","取樣週期減半，每根 K 的隨機波動約放大 √2 倍，而趨勢訊號不變 → 訊噪比下降。\n"
  "在 FOMC 日的會前真空期（成交量全日最低），這個問題會被進一步放大："
  "薄量下的小額買盤就能造出一條漂亮的上升均線，但它不代表真實買盤。"],
 ["免責聲明","本文件為市場資料整理與交易框架研究，非投資建議、非買賣要約。"
  "日內交易涉及高風險並可能損失全部本金。所有價位與判斷請以即時報價覆核。"],
]
r=3
for a,b in items:
    ws.cell(row=r,column=1,value=a).font=Font(size=10,bold=True); ws.cell(row=r,column=1).border=BOX
    ws.cell(row=r,column=1).alignment=Alignment(vertical="top",wrap_text=True)
    c=ws.cell(row=r,column=2,value=b); c.font=Font(size=10); c.border=BOX
    c.alignment=Alignment(vertical="top",wrap_text=True)
    ws.row_dimensions[r].height=max(50,16*b.count("\n")+54)
    r+=1

r+=1
header(ws,r,["項目","來源"]); r+=1
for a,b in [
 ["今日指數變動","Yahoo Finance 9/16 live；TheStreet《Stock Market Today (Sept. 16, 2026)》；CNBC 9/15"],
 ["S&P 7,601","Trading Economics（US500，2026-09-16）"],
 ["FOMC 92.9% 加息機率","CME FedWatch，經 Yahoo Finance / CNBC 9/16 報導"],
 ["Warsh 與 Jackson Hole 轉向","CNBC 2026-08-28《September Fed decision now a coin flip》；CNBC 08-31"],
 ["NVDA 中國禁令","Financial Times，經 CNBC 9/16 報導（網信辦要求終止 RTX Pro 6000D 訂單）"],
 ["半導體昨日跌幅","TradingKey《AI Chip Stocks Plunge》；FX Leaders 9/15"],
 ["雲廠跌幅 / ORCL","Trading Economics 市場評述；TradingKey"],
 ["Brent $106.47 / WTI $101.88","TradingKey；CNBC 9/14《oil surges, 10-year yield hits 19-year high》"],
 ["10 年期 19 年高位","Yahoo Finance 9/15 live；CNBC 9/14"],
 ["黃金跌穿 $4,300","TradingKey 市場回顧"],
]:
    put(ws,r,[a,b],wrap=(1,2),height=22)
    ws.cell(row=r,column=1).font=Font(size=10,bold=True)
    r+=1

p=os.path.join(OUT,"2026-09-16_FOMC日盤中掃描_低波幅10分鐘均線.xlsx")
wb.save(p); print("XLSX ->",p)
