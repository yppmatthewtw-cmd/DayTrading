# -*- coding: utf-8 -*-
"""2026-09-08 美股高交額 20 大 + 低波幅/20分鐘均線平穩向上 篩選"""
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList

OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "output")
os.makedirs(OUT, exist_ok=True)
SESSION = "2026-09-08"

NAVY="1F3864"; SLATE="2F5597"
RED="FFC7CE"; AMBER="FFEB9C"; GREEN="C6EFCE"
REDF="9C0006"; AMBERF="9C6500"; GREENF="006100"
THIN=Side(style="thin", color="BFBFBF"); BOX=Border(left=THIN,right=THIN,top=THIN,bottom=THIN)

def title_row(ws,r,t,span,size=14):
    c=ws.cell(row=r,column=1,value=t); c.font=Font(bold=True,size=size,color="FFFFFF")
    c.fill=PatternFill("solid",fgColor=NAVY); c.alignment=Alignment(vertical="center")
    ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=span); ws.row_dimensions[r].height=24
def note_row(ws,r,t,span):
    c=ws.cell(row=r,column=1,value=t); c.font=Font(italic=True,size=9,color="595959")
    c.alignment=Alignment(wrap_text=True,vertical="top")
    ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=span)
def header(ws,r,hs):
    for i,h in enumerate(hs,1):
        c=ws.cell(row=r,column=i,value=h); c.font=Font(bold=True,size=10,color="FFFFFF")
        c.fill=PatternFill("solid",fgColor=SLATE); c.border=BOX
        c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
    ws.row_dimensions[r].height=34
def widths(ws,spec):
    for k,v in spec.items(): ws.column_dimensions[k].width=v

wb=Workbook()

# ---------------- Sheet 1: 高交額 20 大 ----------------
ws=wb.active; ws.title="01_高交額20大"
widths(ws,{"A":6,"B":9,"C":20,"D":12,"E":11,"F":13,"G":11,"H":9,"I":34})
title_row(ws,1,f"美股 {SESSION} (二) 高交額 20 大  ——  收市：S&P 500 7,673.52 (-0.58%)、道指 52,786.07 (-1.18%)、納指 26,421.41 (-0.32%)",9)
note_row(ws,2,("資料狀態欄：【核實】= 已由公開新聞來源核對；【推斷】= 由當日板塊/新聞背景推導，非逐筆行情。"
 "本環境的行情 API 被網絡政策封鎖，無法取得逐分鐘成交資料，故『成交額排名』為依常態美元成交額與當日最活躍名單交叉推導的近似次序，非交易所官方排行。"),9)
header(ws,4,["#","代號","名稱","收市價","日變動%","日內波幅估算","波幅級別","資料狀態","當日主導事件 / 交易性質"])

rows=[
 [1,"NVDA","NVIDIA",230.36,0.0084,0.018,"低","核實","全市場美元成交額最大；晶片股走強(SOXX +1.7%)，逆大市收升"],
 [2,"TSLA","Tesla",354.08,-0.0592,0.065,"極高","核實","單日 -5.92%，當日大型股最大跌幅之一；全日趨勢向下"],
 [3,"AMD","AMD",None,0.0590,0.062,"極高","核實","+5.90%，AI 晶片動能；升幅大＝波幅大"],
 [4,"INTC","Intel",82.57,0.0470,0.055,"極高","核實","盤前已跳升 4.7%；消息驅動、波動劇烈"],
 [5,"AVGO","Broadcom",None,0.0298,0.032,"中","核實","★ 隨半導體板塊穩步走高，無單一突發消息＝典型的穩健趨勢日"],
 [6,"AAPL","Apple",271.06,None,0.014,"低","核實(價)","超大型股，日內相對波幅全場最小；當日無重大消息"],
 [7,"MSFT","Microsoft",424.60,None,0.015,"低","核實(價)","超大型股，交投極厚；當日無重大消息"],
 [8,"MU","Micron",496.72,None,0.045,"高","核實(價)","記憶體高 beta，日內擺幅一向偏大"],
 [9,"ORCL","Oracle",None,0.0236,0.035,"中高","核實","+2.36%；★ 但業績期臨近，引伸波幅偏高，隔夜風險大"],
 [10,"AMZN","Amazon",None,-0.0060,0.020,"低","核實","-0.60%，全日輕微向下"],
 [11,"QCOM","Qualcomm",None,0.0317,0.048,"高","核實","與 AWS 合作消息，盤中一度 +7%；消息跳空＝波幅大且不平穩"],
 [12,"META","Meta",None,None,0.022,"低中","推斷","當日無重大個股消息；隨大市偏軟"],
 [13,"GOOGL","Alphabet",344.40,None,0.020,"低中","核實(價)","隨大市偏軟"],
 [14,"PLTR","Palantir",143.09,None,0.055,"高","核實(價)","高估值高 beta，日內振幅一向大"],
 [15,"NFLX","Netflix",None,None,0.030,"中","推斷","期權最活躍名單之一"],
 [16,"CRWV","CoreWeave",None,None,0.080,"極高","推斷","Neocloud，日內波幅極大"],
 [17,"IREN","IREN",44.68,None,0.075,"極高","核實(價)","5 日內急升約 23%；波幅極高"],
 [18,"SOFI","SoFi",None,-0.0120,0.029,"中","核實","日內 17.96–18.48，區間 2.9%；收跌 1.2%"],
 [19,"PFE","Pfizer",None,None,0.025,"中","推斷","醫療板塊當日輪動，Amgen -10.15% 拖累情緒"],
 [20,"BE","Bloom Energy",None,0.0600,0.070,"極高","核實","+6%，9/21 納入標普 500；★ 指數納入的買盤最『平穩』，但升幅太大＝波幅高"],
]
r=5
for row in rows:
    for i,v in enumerate(row,1):
        c=ws.cell(row=r,column=i,value=v); c.border=BOX; c.font=Font(size=10)
        c.alignment=Alignment(vertical="top",wrap_text=(i==9))
    ws.cell(row=r,column=5).number_format="+0.00%;-0.00%"
    ws.cell(row=r,column=6).number_format="0.0%"
    lvl={"低":GREEN,"低中":GREEN,"中":AMBER,"中高":AMBER,"高":RED,"極高":RED}[row[6]]
    fg={"低":GREENF,"低中":GREENF,"中":AMBERF,"中高":AMBERF,"高":REDF,"極高":REDF}[row[6]]
    ws.cell(row=r,column=7).fill=PatternFill("solid",fgColor=lvl)
    ws.cell(row=r,column=7).font=Font(size=10,bold=True,color=fg)
    ws.cell(row=r,column=7).alignment=Alignment(horizontal="center",vertical="center")
    ws.row_dimensions[r].height=30
    r+=1

ch=BarChart(); ch.type="col"; ch.style=11
ch.title="2026-09-08 高交額 20 大：日內波幅估算（越低越符合『波幅小』的要求）"
ch.y_axis.title="日內波幅 (高-低)/收市"; ch.height=10; ch.width=30
d=Reference(ws,min_col=6,min_row=4,max_row=r-1)
c=Reference(ws,min_col=2,min_row=5,max_row=r-1)
ch.add_data(d,titles_from_data=True); ch.set_categories(c); ch.legend=None
ws.add_chart(ch,f"A{r+2}")

# ---------------- Sheet 2: 篩選結論 ----------------
ws=wb.create_sheet("02_篩選結論")
widths(ws,{"A":8,"B":10,"C":11,"D":11,"E":13,"F":50})
title_row(ws,1,"篩選：波幅較小 + 20 分鐘均線大致平穩向上",6)
note_row(ws,2,("★ 重要前提：當日 S&P 500『收在全日最低』——午後反彈熄火，尾盤因伊朗襲擊美軍報道再跌一浪。"
 "因此絕大部分跟隨大市的股票，其 20 分鐘均線在下午都是向下的。"
 "能做到『全日 20 分鐘均線平穩向上』的，必然是有自身催化劑、且能抵抗尾盤沽壓的個股。"),6)
header(ws,4,["排名","代號","日變動%","波幅估算","綜合評分","理由 / 需要注意的地方"])
sel=[
 [1,"NVDA",0.0084,0.018,"★★★★☆",
  "最佳綜合。全市場最厚的交投＝日內波幅相對最小、跳動最平滑；在大市收於全日低位的環境下仍收升 0.84%，"
  "代表全日買盤持續而非單點爆發。這是最符合『波幅小 + 均線平穩向上』的候選。\n"
  "注意：+0.84% 的斜率很平緩，20 分鐘均線會是『微微向上』而非明顯上升，日內獲利空間有限。"],
 [2,"AVGO",0.0298,0.032,"★★★★☆",
  "最佳的『趨勢平穩度』。升 2.98% 但沒有單一突發新聞——是跟隨 SOXX(+1.7%) 的板塊性穩步走高，"
  "這種升法通常是全日級距式上移，20 分鐘均線的上升角度最一致。\n"
  "注意：3.2% 的日內波幅屬中等，不算『低波幅』，但相對於 2.98% 的升幅，其『波幅/升幅比』是全場最優。"],
 [3,"AAPL",None,0.014,"★★★☆☆",
  "日內相對波幅全場最小（超大型股+極厚流動性）。若追求的是『均線平穩、少假訊號』，AAPL 的 20 分鐘線"
  "是最乾淨的。\n注意：當日無催化劑，方向大致跟隨大市（偏軟），『向上』的條件未必成立——只滿足『波幅小』。"],
 [4,"MSFT",None,0.015,"★★★☆☆",
  "與 AAPL 同理：波幅低、均線平滑。但同樣缺乏當日向上的催化劑。適合作為『低波幅基準』對照組。"],
 [5,"ORCL",0.0236,0.035,"★★☆☆☆",
  "升 2.36% 且走勢偏穩，但★業績期臨近令引伸波幅偏高，隔夜跳空風險大。日內可以，不宜留倉。"],
]
r=5
for row in sel:
    for i,v in enumerate(row,1):
        c=ws.cell(row=r,column=i,value=v); c.border=BOX; c.font=Font(size=10)
        c.alignment=Alignment(vertical="top",wrap_text=(i==6))
    ws.cell(row=r,column=3).number_format="+0.00%;-0.00%"
    ws.cell(row=r,column=4).number_format="0.0%"
    ws.cell(row=r,column=5).alignment=Alignment(horizontal="center",vertical="center")
    ws.cell(row=r,column=5).font=Font(size=11,bold=True,color=GREENF)
    for cc in range(1,7): ws.cell(row=r,column=cc).fill=PatternFill("solid",fgColor="EAF3EA")
    ws.row_dimensions[r].height=64
    r+=1

r+=1
title_row(ws,r,"明確排除（不符合『波幅小』）",6,12); r+=1
header(ws,r,["代號","日變動%","波幅估算","","",  "排除理由"]); r+=1
exc=[
 ["TSLA",-0.0592,0.065,"單日 -5.92%，方向與要求相反，且波幅極大"],
 ["BE",0.0600,0.070,"雖然指數納入的買盤最『平穩』，但 +6% 的升幅＝波幅過大"],
 ["AMD",0.0590,0.062,"+5.90%，動能股，日內回撤幅度大"],
 ["INTC",0.0470,0.055,"盤前跳空 +4.7%，消息驅動，均線呈階梯狀而非平穩"],
 ["QCOM",0.0317,0.048,"AWS 合作消息，盤中曾 +7% 後回吐——典型的『沖高回落』，均線不平穩"],
 ["IREN",None,0.075,"5 日 +23%，波幅極高"],
 ["CRWV",None,0.080,"Neocloud，日內波幅全場最高"],
 ["MU",None,0.045,"記憶體高 beta"],
 ["PLTR",None,0.055,"高估值高 beta"],
 ["SOFI",-0.0120,0.029,"收跌 1.2%，方向不符"],
]
for row in exc:
    ws.cell(row=r,column=1,value=row[0]).border=BOX; ws.cell(row=r,column=1).font=Font(size=10,bold=True)
    ws.cell(row=r,column=2,value=row[1]).border=BOX; ws.cell(row=r,column=2).number_format="+0.00%;-0.00%"
    ws.cell(row=r,column=3,value=row[2]).border=BOX; ws.cell(row=r,column=3).number_format="0.0%"
    ws.cell(row=r,column=4,value="").border=BOX; ws.cell(row=r,column=5,value="").border=BOX
    c=ws.cell(row=r,column=6,value=row[3]); c.border=BOX; c.font=Font(size=10)
    c.alignment=Alignment(vertical="top",wrap_text=True)
    for cc in range(1,7): ws.cell(row=r,column=cc).fill=PatternFill("solid",fgColor="FDECEC")
    ws.row_dimensions[r].height=24
    r+=1

r+=1
title_row(ws,r,"一句話結論",6,13); r+=1
c=ws.cell(row=r,column=1,value=(
 "若嚴格按『波幅小 + 20 分鐘均線平穩向上』兩個條件同時滿足 → NVDA 最接近（波幅最小、逆市收升，但斜率很平）；\n"
 "若可以接受中等波幅換取更清晰的上升斜率 → AVGO 最理想（板塊性穩步推升，無消息跳空，均線角度最一致）；\n"
 "若只要『均線最平滑、最少假訊號』作為交易載體 → AAPL / MSFT，但當日並不向上。\n\n"
 "★ 但必須提醒：當日大市收於全日最低，尾盤因伊朗襲擊美軍的報道再跌一浪。在這種『尾盤崩』的日子做『順勢做多低波幅股』，"
 "最大的風險不是選錯股，而是選對股卻在 15:30 之後被大市的最後一浪拖累。這一天真正正確的做法是縮短持倉時間，不留尾盤。"))
c.font=Font(size=11); c.alignment=Alignment(vertical="top",wrap_text=True)
ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=6); ws.row_dimensions[r].height=120

# ---------------- Sheet 3: 方法論與限制 ----------------
ws=wb.create_sheet("03_方法論與限制")
widths(ws,{"A":24,"B":80})
title_row(ws,1,"方法論、資料限制與可重現的計算方式",2)
items=[
 ["本次的資料限制",
  "此執行環境的網絡政策封鎖了所有行情 API（Yahoo Finance、Stooq、TradingView、Nasdaq、stockanalysis 等全部回 403）。"
  "因此我無法取得 2026-09-08 的逐分鐘成交資料，也就無法真正計算 20 分鐘移動平均線的斜率。\n"
  "本檔案中：收市價與日變動% 為已核實的公開資料；『日內波幅估算』與『20 分鐘均線是否平穩向上』為依當日板塊表現、"
  "新聞驅動性質與各股常態波動特性所作的推斷，並已在資料狀態欄逐項標示。請勿當作實測數據使用。"],
 ["嚴謹的做法",
  "同目錄下的 scripts/screen_20min_ma.py 實作了完整的篩選邏輯。在有行情權限的環境（本機或有 API key 的伺服器）執行，"
  "即可得出實測結果，不需再依賴推斷。"],
 ["篩選邏輯（腳本實作）",
  "1) 取當日全市場成交額 (close × volume) 前 20 名；\n"
  "2) 對每檔取 1 分鐘 K 線，重採樣為 20 分鐘 K 線，計算 20 分鐘收盤價的移動平均；\n"
  "3) 平穩度 = 20 分鐘均線的一階差分中『為正』的比例（rising_ratio），越接近 1 越平穩向上；\n"
  "4) 上升斜率 = 對 20 分鐘均線做線性迴歸，取標準化斜率與 R²（R² 越高＝越『平穩』而非鋸齒）；\n"
  "5) 波幅 = 日內 (High−Low)/Close，以及 20 分鐘收益率的標準差（年化前的原始值）；\n"
  "6) 綜合排序 = 先以波幅由小至大排序，再要求 rising_ratio ≥ 0.7 且 迴歸斜率 > 0 且 R² ≥ 0.6。"],
 ["為何用 R² 而非只看斜率",
  "『平穩向上』有兩個獨立條件：向上（斜率為正）與平穩（路徑接近直線）。只看斜率會把『沖高回落後再拉起』"
  "這種鋸齒型走勢誤判為合格——例如當日的 QCOM。加入 R² 門檻可以有效濾走消息跳空型的個股。"],
 ["當日的關鍵背景（影響所有判讀）",
  "S&P 500 收於全日最低點；午後反彈熄火，尾盤因『伊朗向美軍發動襲擊』的報道再跌一浪。"
  "Brent 盤中最高逼近 $99，WTI 企穩 $90 之上；另有新的加拿大關稅消息與債息上升。"
  "在這種環境下，『20 分鐘均線全日平穩向上』的個股數量本身就會非常稀少。"],
 ["免責聲明",
  "本文件為市場資料整理與交易框架研究，非投資建議、非買賣要約。日內交易涉及高風險並可能損失全部本金。"
  "所有價位與判斷請以即時報價與自身研究覆核。"],
]
r=4
for a,b in items:
    ws.cell(row=r,column=1,value=a).font=Font(size=10,bold=True)
    ws.cell(row=r,column=1).border=BOX
    ws.cell(row=r,column=1).alignment=Alignment(vertical="top",wrap_text=True)
    c=ws.cell(row=r,column=2,value=b); c.font=Font(size=10); c.border=BOX
    c.alignment=Alignment(vertical="top",wrap_text=True)
    ws.row_dimensions[r].height=max(56, 15*b.count("\n")+56)
    r+=1

# ---------------- Sheet 4: 來源 ----------------
ws=wb.create_sheet("04_資料來源")
widths(ws,{"A":34,"B":76})
title_row(ws,1,"資料來源",2)
header(ws,4,["項目","來源"])
src=[
 ["指數收市 (9/8)","CNBC《Stock market news for Sept. 8, 2026》；TheStreet；econcurrents"],
 ["盤中走勢與尾盤","CNBC / TheStreet：『S&P 500 收於全日低位』、『午後反彈熄火』、『尾盤因伊朗襲擊美軍報道再跌』"],
 ["最活躍名單 (20 檔)","Market Rebellion《Mid-session IV Report September 8, 2026》最活躍期權名單"],
 ["NVDA / TSLA 收市","Motley Fool / CNBC 個股報導 (9/8)"],
 ["AMD / AVGO / ORCL / QCOM / AMZN","Motley Fool《Stock Market Today, Sept. 8, 2026》"],
 ["INTC / SOXX / Broadcom 盤前","CNBC 盤前報導 (9/8)"],
 ["AAPL/MSFT/GOOGL/INTC/MU/PLTR 價格","Edward Jones stock table（截至 9/8 的價格表）"],
 ["SOFI 日內高低","MarketBeat 個股快訊 2026-09-08（高 18.48 / 低 17.96）"],
 ["BE (Bloom Energy)","CNBC：9/21 納入標普 500，當日升逾 6%"],
 ["IREN","公開報導：最近收 $44.68，5 日約 +23%"],
 ["Amgen / Salesforce / Home Depot","當日最大跌幅股報導 (9/8)"],
 ["油價背景","Trading Economics（Brent 9/8 = $97.41）；CNBC 盤中報導（Brent 盤中近 $99）"],
]
r=5
for a,b in src:
    ws.cell(row=r,column=1,value=a).border=BOX; ws.cell(row=r,column=1).font=Font(size=10,bold=True)
    ws.cell(row=r,column=1).alignment=Alignment(vertical="top",wrap_text=True)
    c=ws.cell(row=r,column=2,value=b); c.border=BOX; c.font=Font(size=10)
    c.alignment=Alignment(vertical="top",wrap_text=True)
    ws.row_dimensions[r].height=26
    r+=1

p=os.path.join(OUT,f"{SESSION}_高交額20大_低波幅20分鐘均線篩選.xlsx")
wb.save(p); print("XLSX ->",p)
