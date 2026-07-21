# -*- coding: utf-8 -*-
"""
R1 版生成器：把《規則庫（83 條）》全部規則逐條套用到 DCFC 成功日 / 失敗日
對照案例上，生成含「逐條檢核表＋成敗歸因」的 ross_toolkit_R1.html。

輸入: ross_toolkit_offline.html（自包含離線版，圖表已內嵌）
輸出: ross_toolkit_R1.html

編號說明: 本檢核一律使用頁內《規則庫》表格的 # 欄編號（1–83）並附規則名稱。
舊版 case-study 圖表與對照表使用 CSV 行號（= 表格編號 + 1），故編號相差 1。

執行: python3 build_r1.py
"""
import io, os, re

SRC = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ross_toolkit_offline.html")
DST = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ross_toolkit_R1.html")

P, F, W, N = "P", "F", "W", "N"   # 遵守 / 違反 / 邊緣 / 當日未觸發
BADGE = {P: ("v-pass", "✅ 遵守"), F: ("v-fail", "❌ 違反"),
         W: ("v-part", "⚠️ 邊緣"), N: ("v-na", "➖ 未觸發")}

# ─────────────────────────── 逐條檢核資料 ───────────────────────────
# (編號, 規則名稱, 門檻摘要, 成功日判定, 成功日證據, 失敗日判定, 失敗日證據)

SEC_A = [  # 選股與市場環境 —— 證明「成敗不在選股」
 (7,  "五大選股標準完整框架", "5/5 全達標＝最高確信；3/5 以下不入場",
  P, "5/5：價格甜區＋低 Float＋漲幅榜首＋量 >5x＋白宮級消息",
  W, "表面仍 4–5/5，但「消息」一項已是第 2 天舊聞——評分虛高（見 #50/#76）"),
 (4,  "Float 低＝供需失衡＝拋物線動能", "Float<10M＝低供應",
  P, "deSPAC 低流通股本，開盤即爆量拉升",
  P, "同一隻股票——連環 halt up 正是供需失衡的證明"),
 (6,  "大 Float（>50M）避免", "厚 Float 不做動能策略",
  P, "非大 Float 股", P, "同左"),
 (16, "高日掃描器觸發＋Float 確認", "掃描器捕捉 → 確認 Float＋消息",
  P, "兩天皆為掃描器榜首股", P, "同左——選股流程沒有錯"),
 (35, "Trade-Ideas 掃描篩選流程", "缺口>4% → Float → 消息確認",
  P, "缺口＋消息＋Float 全確認", P, "同左"),
 (38, "消息催化劑質量分級", "FDA>臨床>合作>PP>普通；政策/白宮級屬最高",
  P, "白宮點名＝政策級（高於 FDA 級），且為第 1 天",
  F, "級別雖高但已是第 2 天媒體轉載——把「高級別」誤當「新消息」"),
 (51, "大公司/大人物名字在標題", "名人效應 → 情緒爆發 → 量能激增",
  P, "Biden＋白宮＝本規則的極端形態，情緒最大化",
  N, "名人效應已於第 1 天釋放完畢"),
 (48, "主題股跟進", "識別當前最強主題",
  P, "EV 充電＋白宮政策＝當時最強主題",
  W, "主題判斷仍對，但已進入第 2 天（見 #50）"),
 (54, "$2–$10 股價甜區", "<$2 點差差；>$10 波動率下降",
  P, "$6.6–8.5 正中甜區",
  W, "$10.5–11.9 已在甜區上緣之外，百分比空間縮小"),
 (55, "<$2 Penny 股避免", "點差佔比過高", P, "達標", P, "達標"),
 (72, "明顯股票法則", "只交易當天最明顯的 1–2 隻",
  P, "全市場最明顯的股票",
  P, "仍是最明顯的股票——兩天選股同分，成敗不在選股"),
 (32, "Hot / Cold 市場識別", "Hot 進攻、Cold 防守",
  P, "事件日動能充足，進攻合理",
  F, "Ross 自述當時屬冷市，卻用熱市打法重倉進攻"),
]

SEC_B = [  # 買入規則
 (9,  "Gap &amp; Go 核心策略", "缺口>4%＋消息＋開盤量>5x＝觸發",
  P, "教科書執行：破盤前高 $7.00 觸發，觸發位客觀、可預掛條件單",
  N, "當日未使用此結構——連環 halt 中的追價不是 Gap&amp;Go 計劃"),
 (1,  "VWAP 動態支撐（Long Bias）", "價>VWAP 做多；跌破＋放量出場",
  P, "入場貼近 VWAP，多頭結構完整",
  F, "$10.50 與 $11.32 均價皆在離 VWAP >10% 的延伸段"),
 (2,  "VWAP 偏離 >5%＝異常", "偏離過大＝均值回歸壓力",
  P, "入場點偏離小，回歸壓力低",
  F, "偏離 >10% 仍追入——在回歸壓力最大處接棒"),
 (3,  "VWAP 觸底＋halt 第一波入場", "守住 VWAP 後的第一個 halt resumption",
  N, "當日走 Gap&amp;Go，未用此結構",
  F, "買的不是「守住 VWAP 後的第一個 halt」，而是連環 halt 後段的高位延伸"),
 (10, "Day 2 Play 延續策略", "前提：今日『新消息』＋突破前日高",
  N, "本身就是第 1 天",
  F, "兩個前提皆缺：無新消息、未破前日高——不構成合格 Day 2 Play"),
 (12, "Bull Flag 第一次回踩入場", "縮量回踩 → 首根新高 K 線",
  P, "入場②＝首次縮量回踩後首根新高 K 線",
  N, "未使用此結構"),
 (13, "第二次回踩勝率遞減", "突破次數越多勝率越低",
  N, "未貪第三波",
  F, "$11.50 加回已是行情第 N 波，勝率遞減被無視"),
 (14, "Mini Pullback 精確識別", "量縮＋未破前波 20%＋MACD 正",
  P, "回踩縮量、守住前波、MACD 仍正",
  F, "把崩回 VWAP 的深度下殺當成 Pullback 抄底"),
 (15, "Dump 識別避免接刀", "放量上影＋MACD 轉負＋回踩>30%＝Dump",
  N, "Dump 出現前已出場",
  F, "$11.90 見頂崩回 VWAP＝Dump 訊號齊備，仍向下接刀"),
 (22, "LULD Halt 交易", "resumption 順向做，『第一個下破前低』止損",
  N, "未涉及",
  F, "方向雖對（halt up 後續拉），但下破前低時不止損反而攤平"),
 (23, "MACD 綠燈/紅燈系統", "MACD<0 不做多，無例外",
  P, "動能段 MACD 正時進場",
  F, "攤平買入發生在 MACD 轉負的紅燈區"),
 (24, "MACD Bull Trap 避免", "紅燈時任何突破/低吸都不買",
  N, "未涉及",
  F, "紅燈區 $10.30/$10.40 低吸＝教科書級接刀"),
 (25, "均線多頭排列", "5>20>50>200 排列做多",
  P, "開盤強勢段均線多頭排列",
  F, "加倉時價格已跌破短期 EMA、貼近 VWAP——結構已壞"),
 (29, "9:30–9:45 黃金 15 分鐘", "最大流動性窗口",
  P, "兩進兩出全部在開盤 25 分鐘內完成",
  W, "入場雖也在早盤，但錯不在時段而在位置——本規則救不了他"),
 (52, "大公司關聯消息交易", "快進快出；目標＝前高/整數關口",
  P, "快進快出，整數關口前出清",
  F, "$11.50 追入時距 $12 整數關口僅 4%——把「目標位」當「入場位」"),
 (61, "Educated Intuition 說不", "setup 不合格就不交易",
  P, "只打合格 setup，全日僅兩筆",
  F, "第 2 天＋延伸段＋冷市＝三重不合格，仍硬入場"),
 (66, "Jack-knife 反轉避免", "急漲急崩刀型走勢，Choppy 市常見",
  N, "未遇到",
  F, "$11.90 急漲即崩＝典型 jack-knife，冷市特徵被無視"),
 (67, "假突破識別", "突破量 <3x＝假突破不追",
  P, "突破量 >5x 才追",
  W, "$11.50「新高 K 線」旋即見頂回崩——事後證實為假突破（量能記載不可考）"),
 (71, "Parabolic 識別與倉位分配", "拋物線＝最強動能：分批跟隨＋止損",
  N, "未遇到拋物線段",
  F, "拋物線判斷沒錯，但規則要求分批＋止損——他一鍵 9,000 股、無止損"),
 (75, "Back-side 後段禁區", "從高點回落/跌破均線後不追多",
  P, "只在前側（創新高階段）交易",
  F, "攤平發生在跌破短期均線、崩回 VWAP 的後側結構——後段做多勝率 <40%"),
 (76, "消息已 Price-in 識別", "次日無新催化＝追入風險高",
  P, "第 1 天消息未 price-in，動能最純",
  F, "第 2 天舊消息已 price-in——本日一切錯誤的總根源"),
 (50, "主題第 2–3 天風險", "主題股第 2–3 天後動能衰減",
  P, "只打第 1 天",
  F, "核心錯誤：追已 price-in 的第 2 天行情"),
]

SEC_C = [  # 賣出規則
 (59, "恐懼導致過早止盈", "預設目標位對抗恐懼",
  P, "延伸棒目標位預設：第一延伸棒賣半倉、第二延伸棒出清",
  N, "當日問題方向相反（拿不住的是虧損不是利潤）"),
 (60, "恐懼導致過晚止損", "止損拖延＝最大錯誤",
  N, "止損未被觸發",
  P, "唯一救贖：先以攤平拖延，但最終 $10.85 反彈認賠——避開 -$3/股 × 9,000＝-$27,000"),
 (30, "9:30–10:05 動能衰減管理", "10:05 後動能急降；11:30 後停止",
  P, "10:00 前完成全部交易並收工",
  F, "動能已衰減（高點回落）仍在場內加碼"),
 (44, "盈虧比必須 ≥1", "平均虧損>平均盈利＝勝率再高也虧",
  P, "突破位入場止損近、目標延伸棒——盈虧比 >1.5",
  F, "均價 $11.32、上方到 $12 僅 6%、下方無止損——盈虧比嚴重倒掛"),
 (47, "峰值回吐 50% 停止", "從當日峰值回吐一半＝立即收工",
  P, "賺夠收工，利潤零回吐",
  F, "峰值約 +$5,500 → 回吐 100% 轉紅仍未停手"),
 (78, "給回一半獲利／移動止損", "大贏後設移動止損保護浮盈",
  P, "峰值利潤全額落袋",
  F, "浮盈無移動止損保護，+$4,500 僥倖利潤原路吐回"),
 (46, "日最大虧損強制停止", "觸及 Max Loss 無例外停止",
  N, "未觸及",
  P, "認賠後停手，未演變成情緒化連環虧損"),
 (45, "大倉位滑價風險", "倉位須匹配流動性；大倉限價出場",
  P, "計劃內倉位，從容分批出場",
  F, "9,000 股高位倉恐慌出場，平均損失 $0.47/股"),
]

SEC_D = [  # 倉位與風控
 (36, "觀察清單 SOP／預設訂單", "開盤前預設入場、止損、倉位",
  P, "預掛條件單，入場＝計劃被觸發",
  F, "FOMO 無計劃入場；熱鍵誤按 9,000 股＝無預設倉位上限"),
 (68, "5 項清單缺項降倉", "缺任 1 項 → 降倉或觀望",
  P, "5/5 滿分，正常倉位進攻",
  F, "形態、位置兩項不達標，按規則應降倉觀望——實際反而放最大倉位"),
 (41, "大贏日次日縮倉 50%", "大贏後只打 5/5 setup 且減半倉",
  N, "（本身是大贏日）",
  F, "+$12k 大贏之後的交易日，應縮倉——實際重倉 9,000 股"),
 (42, "Stair-Step vs 攤平", "虧損後同倉位重試；攤平＝典型錯誤",
  N, "未觸發虧損",
  F, "$10.30/$10.40 向下攤平＝本規則點名的頭號錯誤"),
 (40, "Revenge Trading 識別", "加碼求回本＝危險信號",
  N, "未觸發",
  F, "攤平加碼求回本＝變相 revenge trading"),
 (62, "質量 > 數量", "少而精，學會不交易",
  P, "全日僅 2 筆入場、2 筆出場",
  F, "追高 → 僥倖賣 → 加回 → 攤平 → 認賠，交易鏈完全失控"),
 (70, "市場熱度動態倉位", "Hot +50%／Cold -50%",
  P, "熱門事件日正常倉位進攻",
  F, "冷市本應 -50% 倉位，實際反而放大到 9,000 股"),
 (74, "連贏漸進增倉", "每 10 次成功 +10–20%，不跳級",
  N, "未涉及",
  F, "由常規倉位一步跳到 9,000 股（雖屬誤操作，跳級後果照樣發生）"),
 (79, "執行平台／熱鍵風險", "執行基礎設施是風控的一部分",
  N, "未涉及",
  F, "熱鍵誤操作買入 9,000 股——倉位失控的直接技術原因"),
 (33, "冷市正確應對", "冷市縮倉等待，勿加頻率求彌補",
  N, "非冷市議題",
  F, "Ross 自認根因：「冷市裡踩不住煞車」——所有錯誤的心理土壤"),
]

NA_RULES = [  # 兩日皆不適用（覆蓋滿 83 條）
 (5,  "Float<1M 極端波動機會", "DCFC 非 sub-1M float"),
 (8,  "Biotech/FDA 行業偏好", "EV 充電行業（消息級別已計入 #38）"),
 (11, "下午突破 Afternoon Breakout", "兩日交易均在早盤"),
 (17, "HTB 空頭擠壓", "借券費率無記載"),
 (18, "ETB 二波策略", "借券狀態無記載"),
 (19, "Sub-1M Float 超級擠壓", "非 sub-1M float"),
 (20, "Reverse Split 降 Float 效應", "無 reverse split 背景"),
 (21, "T1 Halt 消息暫停", "當日均為 LULD 熔斷，非 T1"),
 (26, "EMA vs SMA 工具選擇", "工具設定類規則，不涉當日判定"),
 (27, "Level 2 關鍵賣壓", "盤口資料無記載"),
 (28, "Ascending Wedge 突破", "形態未出現"),
 (31, "7AM 整點突破窗口", "盤前交易無記載"),
 (34, "節假日前後冷市標準", "非節假日窗口"),
 (37, "1500% 漲幅股識別", "非該量級事件"),
 (39, "月初動能設置", "非月初前三天窗口"),
 (43, "連紅後的心理管理", "無連續紅日背景"),
 (49, "代碼熱詞效應", "DCFC 代碼無熱詞"),
 (53, "國防/石油主題（2025）", "2022 年事件"),
 (56, "$2,000 小帳戶挑戰", "主帳戶交易"),
 (57, "$25,000 PDT 門檻", "帳戶遠超門檻"),
 (58, "PDT 規則改革（2025）", "2022 年事件"),
 (63, "熱市倉位管理（77 筆）", "當時屬冷市（見 #32/#33）"),
 (64, "Daily Breakout Setup", "日線形態非本案觸發"),
 (65, "10AM 後突發 Squeeze", "主要行情在開盤段"),
 (69, "月度 Catalyst Profile 回顧", "月度統計類規則"),
 (73, "連紅後縮倉反彈", "無連紅背景"),
 (77, "中國上市股風險", "Tritium 為澳洲公司"),
 (80, "宏觀事件指數大波動", "個股政策消息，非指數級宏觀日"),
 (81, "加密催化劑衰退（2025）", "非加密主題"),
 (82, "強勢月份識別與計劃", "月度節奏類規則"),
 (83, "SVB 銀行危機影響", "2023 年事件"),
]

ALL_SECS = [
 ("A", "選股與市場環境", "兩天幾乎同分 —— 用數據證明「成敗不在選股」", SEC_A),
 ("B", "買入規則逐條檢核", "入場位置、消息天數、結構確認 —— 失敗日的重災區", SEC_B),
 ("C", "賣出規則逐條檢核", "止盈預設化、止損不拖延、利潤保護", SEC_C),
 ("D", "倉位與風控規則逐條檢核", "計劃、倉位上限、攤平禁令、市場溫度", SEC_D),
]


def tally():
    cnt = {"win": {P: 0, F: 0, W: 0, N: 0}, "red": {P: 0, F: 0, W: 0, N: 0}}
    for _, _, _, rows in ALL_SECS:
        for (_n, _name, _g, wv, _wn, rv, _rn) in rows:
            cnt["win"][wv] += 1
            cnt["red"][rv] += 1
    return cnt


def badge(v):
    cls, txt = BADGE[v]
    return '<span class="v %s">%s</span>' % (cls, txt)


def build_tables():
    out = io.StringIO()
    for key, title, sub, rows in ALL_SECS:
        out.write('<h3 class="r1h">%s · %s</h3>\n' % (key, title))
        out.write('<div class="hint">%s</div>\n' % sub)
        out.write('<div style="overflow-x:auto"><table class="r1tbl">\n')
        out.write('<thead><tr><th style="width:34px">#</th><th style="width:20%">規則（門檻）</th>'
                  '<th style="width:38%">✅ 成功日（+$12k）</th>'
                  '<th style="width:38%">⛔ 失敗日（紅日）</th></tr></thead><tbody>\n')
        for (num, name, gist, wv, wn, rv, rn) in rows:
            out.write('<tr data-w="%s" data-r="%s"><td class="mono">%d</td>'
                      '<td><b>%s</b><br><small>%s</small></td>'
                      '<td>%s<div class="ev">%s</div></td>'
                      '<td>%s<div class="ev">%s</div></td></tr>\n'
                      % (wv, rv, num, name, gist, badge(wv), wn, badge(rv), rn))
        out.write('</tbody></table></div>\n')
    return out.getvalue()


def build_na():
    items = "".join('<tr><td class="mono">%d</td><td>%s</td><td>%s</td></tr>' % (n, name, why)
                    for n, name, why in NA_RULES)
    return ('<details class="r1na"><summary><b>➖ 其餘 %d 條規則兩日皆不適用（點開看原因）'
            '</b>——帳戶制度、其他年代主題、資料不可考等，逐條列明以覆蓋全部 83 條</summary>'
            '<div style="overflow-x:auto"><table class="r1tbl"><thead><tr><th>#</th><th>規則</th>'
            '<th>不適用原因</th></tr></thead><tbody>%s</tbody></table></div></details>'
            % (len(NA_RULES), items))


def build_section():
    t = tally()
    win_p, red_p = t["win"][P], t["red"][P]
    red_f, red_w = t["red"][F], t["red"][W]
    assessed = sum(len(r) for _, _, _, r in ALL_SECS)

    s = io.StringIO()
    s.write('<section id="r1" class="layer">\n')
    s.write('<h2>🔬 R1 · 全 83 條規則逐條檢核：同一隻股票，為什麼一天 +$12k、一天轉紅</h2>\n')
    s.write('<div class="hint">方法：把下方《規則庫（83 條）》<b>每一條</b>規則逐一對照成功日'
            '（2022/2/8，消息第 1 天，+$12k）與失敗日（白宮消息第 2 天，紅日）的實際決策，'
            '判定 ✅遵守 / ❌違反 / ⚠️邊緣 / ➖當日未觸發。共 %d 條可判定、%d 條兩日皆不適用（見文末摺疊清單）。'
            '規則編號以《規則庫》表格 # 欄為準並附名稱；舊版圖表註記使用 CSV 行號（= 本編號 + 1），請以名稱對照。'
            '</div>\n' % (assessed, len(NA_RULES)))

    # 記分板
    s.write('<div class="r1stats">\n')
    s.write('<div class="r1card"><div class="r1lab">✅ 成功日（+$12k）</div>'
            '<div class="r1num good">%d 遵守 · 0 違反</div>'
            '<small>其餘 %d 條當日未被觸發——因為根本沒讓自己陷入需要它們的處境</small></div>\n'
            % (win_p, t["win"][N]))
    s.write('<div class="r1card"><div class="r1lab">⛔ 失敗日（紅日）</div>'
            '<div class="r1num warn">%d 違反 + %d 邊緣 · 僅 %d 遵守</div>'
            '<small>遵守的規則裡有兩條（#60 認賠、#46 停手）正是把紅日擋在 -$27k 之外的救贖</small></div>\n'
            % (red_f, red_w, red_p))
    s.write('<div class="r1card"><div class="r1lab">同一個交易者、同一隻股票</div>'
            '<div class="r1num">36 : 0 → 8 : 35</div>'
            '<small>選股層面（A 組）兩天幾乎同分——差距 100% 出在執行面（B/C/D 組）</small></div>\n')
    s.write('</div>\n')

    # 過濾按鈕
    s.write('<div class="r1filter">顯示：'
            '<button class="fbtn on" onclick="r1Filter(this,\'all\')">全部</button>'
            '<button class="fbtn" onclick="r1Filter(this,\'F\')">只看失敗日違反（%d）</button>'
            '<button class="fbtn" onclick="r1Filter(this,\'W\')">只看邊緣 ⚠️（%d）</button>'
            '<button class="fbtn" onclick="r1Filter(this,\'P\')">只看失敗日仍遵守（%d）</button></div>\n'
            % (red_f, red_w, red_p))

    s.write(build_tables())
    s.write(build_na())

    # 為什麼成功
    s.write("""
<h3 class="r1h">✅ 為什麼成功日成功 —— 五個機制</h3>
<ol class="r1list">
 <li><b>只打消息第 1 天</b>（#50 主題天數／#76 price-in）：買盤尚未被消耗，動能最純——這是當天所有優勢的源頭。</li>
 <li><b>觸發位客觀</b>（#9 Gap&amp;Go 破盤前高 $7.00、#12 首次回踩首根新高 K 線）：入場是「計劃被觸發」而非「情緒被觸發」，止損自然貼近入場位。</li>
 <li><b>三重確認過濾假訊號</b>（#67 量 &gt;5x、#23 MACD 綠燈、#1 貼近 VWAP）：同一時刻要求量、動能、位置三者同時成立。</li>
 <li><b>出場預設化</b>（#59 延伸棒目標分批止盈、#52 整數關口前不戀戰）：同時擋住「恐懼提早止盈」與「貪婪不出場」兩個方向的情緒。</li>
 <li><b>停利即收工</b>（#47 峰值回吐規則／#78 移動止損／#30 時間衰減）：全日兩進兩出、開盤 25 分鐘內完成，利潤零回吐——當日權益曲線只上不下。</li>
</ol>
<div class="out">成功日的本質：<b>36 條遵守、0 條違反</b>。不是「做對了某一件事」，而是讓每個決策點都落在規則框架內，於是 16 條風控規則（攤平禁令、認賠、Max&nbsp;Loss…）整天都沒有被觸發的機會。</div>
""")

    # 為什麼失敗
    s.write("""
<h3 class="r1h">⛔ 為什麼失敗日失敗 —— 四個連環錯誤 + 一個根因 + 一個救贖</h3>
<ol class="r1list r1chain">
 <li><b>錯誤① 入場位置錯</b>（違反 #1/#2/#75/#52）：$10.50 FOMO 追高，離 VWAP &gt;10%；隨後 $11.50 再追（均價 $11.32），距 $12 整數關口只剩 4% 空間——把「賣出目標位」當成了「買入位」。</li>
 <li><b>錯誤② 消息天數錯</b>（違反 #76/#50/#38/#10）：追的是白宮消息第 2 天的媒體轉載。Day 2 Play 的兩個前提（新消息、突破前日高）一個都不存在。</li>
 <li><b>錯誤③ 倉位失控</b>（違反 #36/#79/#41/#74/#70/#45）：熱鍵誤按 9,000 股、大贏日次日不縮倉、冷市反而放大倉位、倉位與流動性不匹配（出場滑價 $0.47/股）。</li>
 <li><b>錯誤④ 虧損處理錯</b>（違反 #42/#24/#15/#23）：$11.90 見頂崩回 VWAP 是 Dump 不是回踩，他卻在 MACD 紅燈區 $10.30/$10.40 向下攤平——規則庫點名的頭號錯誤。</li>
 <li><b>根因</b>（違反 #33/#32/#61）：Ross 自己的結論——「冷市裡踩不住煞車」。當天先有 +$1,000 又僥倖 +$4,500，峰值 +$5,500 全數回吐轉紅（違反 #47/#78）。$10.95 halt 前那筆僥倖獲利最危險：它獎勵了壞行為。</li>
 <li class="r1save"><b>救贖</b>（遵守 #60/#46）：反彈 $10.85 全部認賠出場。若拖到 $8.30 = -$3/股 × 9,000 股 = <b>-$27,000</b>。當日只做對兩條規則，但這兩條規則價值 $27k——<i>“You got to cut the losses. You got to learn from them. You got to move on.”</i></li>
</ol>
""")

    # 結論
    s.write("""
<h3 class="r1h">🎯 R1 結論</h3>
<div class="out"><b>1 · 成敗不在選股。</b>A 組 12 條選股規則兩天幾乎同分（同一隻最明顯的股票 #72）。差距全部出在 B/C/D 組：入場位置、消息天數、倉位控制、虧損處理。
<b>2 · 買入規則決定你賺多少，賣出／風控規則決定你能活多久。</b>失敗日違反 35 條仍只是「紅日」而非「災難日」，唯一原因是守住了 #60 認賠——賣出紀律的價值是不對稱的（-$0.47/股 vs -$3/股）。
<b>3 · 35 條違反中約 30 條可被機械化防呆擋掉。</b>層二檢查表的否決條件（消息第 2 天、離 VWAP &gt;10%、後側結構）可擋錯誤①②④；層三計算器的倉位上限與預設止損可擋錯誤③。情緒無法根除，但可以被流程隔離。
<b>4 · 最危險的一筆是賺錢的那筆。</b>$10.95 halt 前的 +$4,500 僥倖獲利，直接餵大了 $11.50 的加碼——檢核違規時，要連「賺錢的違規」一起算。</div>
""")

    s.write('</section>\n')

    # 過濾 JS（不用 f-string，避免大括號衝突）
    s.write("""<script>
function r1Filter(btn, mode){
  document.querySelectorAll('.r1filter .fbtn').forEach(function(b){b.classList.remove('on')});
  btn.classList.add('on');
  document.querySelectorAll('.r1tbl tbody tr[data-r]').forEach(function(tr){
    tr.style.display = (mode==='all' || tr.dataset.r===mode) ? '' : 'none';
  });
}
</script>
""")
    return s.getvalue()


R1_CSS = """
  /* ── R1 檢核 ── */
  .r1h{font-size:14.5px; margin:22px 0 2px}
  .r1stats{display:grid; gap:12px; grid-template-columns:repeat(auto-fit,minmax(220px,1fr)); margin:14px 0 4px}
  .r1card{border:1px solid var(--line); border-radius:10px; padding:12px 14px; background:var(--bg)}
  .r1lab{font-size:12px; color:var(--ink2); margin-bottom:4px}
  .r1num{font-size:20px; font-weight:800}
  .r1card small{display:block; color:var(--ink2); font-size:11.5px; margin-top:5px; line-height:1.45}
  .r1filter{margin:10px 0 2px; font-size:12.5px; color:var(--ink2)}
  .fbtn{border:1px solid var(--line); background:var(--card); color:var(--ink); font-size:12px;
        padding:4px 10px; border-radius:99px; margin:0 4px 4px 0; cursor:pointer}
  .fbtn.on{background:var(--blue); border-color:var(--blue); color:#fff}
  .v{display:inline-block; font-size:11px; font-weight:700; padding:1px 8px; border-radius:99px; white-space:nowrap}
  .v-pass{background:rgba(22,163,74,.14); color:var(--green)}
  .v-fail{background:rgba(220,38,38,.14); color:var(--red)}
  .v-part{background:rgba(217,119,6,.16); color:var(--amber)}
  .v-na{background:rgba(100,116,139,.15); color:var(--ink2)}
  .r1tbl td{vertical-align:top}
  .r1tbl .ev{font-size:12.5px; color:var(--ink2); margin-top:3px; line-height:1.45}
  .r1tbl small{color:var(--ink2); font-size:11.5px}
  .r1list{margin:10px 0 0 20px; font-size:13.5px}
  .r1list li{margin-bottom:8px; line-height:1.55}
  .r1chain li{border-left:3px solid var(--red); padding-left:10px; list-style-position:inside; margin-left:-20px}
  .r1chain li.r1save{border-left-color:var(--green)}
  .r1na{margin-top:16px}
  .r1na summary{cursor:pointer; font-size:13px}
"""


def main():
    html = open(SRC, encoding="utf-8").read()

    html = html.replace(
        "<title>Ross Momentum 交易工具包（離線完整版）</title>",
        "<title>Ross Momentum 交易工具包 R1（離線完整版・含 83 條規則成敗檢核）</title>", 1)

    i = html.find("</style>")
    assert i > 0, "style block not found"
    html = html[:i] + R1_CSS + html[i:]

    nav_anchor = '<a href="#rules">'
    assert nav_anchor in html, "nav anchor not found"
    html = html.replace(nav_anchor, '<a href="#r1">🔬 R1 規則檢核</a>\n  ' + nav_anchor, 1)

    sec_anchor = '<section id="rules"'
    assert sec_anchor in html, "rules section not found"
    html = html.replace(sec_anchor, build_section() + "\n" + sec_anchor, 1)

    with open(DST, "w", encoding="utf-8") as f:
        f.write(html)
    print("written:", DST, len(html), "bytes")


if __name__ == "__main__":
    main()
