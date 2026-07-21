# -*- coding: utf-8 -*-
"""
R2 版生成器：在 R1（全 83 條規則逐條檢核）之上，新增「☑️ 決策檢查表對決」章節——
把兩個例子的比較做成視覺化 checklist：

  · 左右對決記分卡（成功日 vs 失敗日，總分色條）
  · 兩天的簡化價格路徑 SVG（入場/出場/錯誤點直接標在圖上，並排對照）
  · 決策時間軸（每一步一個色點：綠=正確、紅=違規、橙=僥倖）
  · 五階段決策檢查表（盤前選股→入場確認→倉位風控→虧損處理→出場收工），
    每個檢查點左右兩格大色塊 ✓/✗ 直觀對照，並附階段計分

輸入: ross_toolkit_R1.html（含 R1 檢核章節）
輸出: ross_toolkit_R2.html

執行: python3 build_r2.py
"""
import io, os

SRC = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ross_toolkit_R1.html")
DST = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ross_toolkit_R2.html")

OK, BAD, PT, NA = "ok", "bad", "part", "na"
GLYPH = {OK: "✓", BAD: "✗", PT: "!", NA: "–"}

# ─────────────────── 檢查表（依交易生命週期分五階段） ───────────────────
# (檢查點問題, 規則編號, 成功日判定, 成功日短註, 失敗日判定, 失敗日短註)

STAGES = [
 ("①", "盤前 · 選股與環境", "兩天在這裡幾乎同分——成敗不是選股決定的", [
  ("這是今天全市場最明顯的股票嗎？", "#72",
   OK, "DCFC 漲幅／成交量榜首", OK, "同樣是榜首——選股沒有錯"),
  ("低 Float ＋ $2–10 甜區 ＋ 相對量 >5x？", "#4 #54",
   OK, "$6.6–8.5 正中甜區", PT, "Float／量仍達標，但價已升破 $10 甜區上緣"),
  ("有頂級催化劑（白宮／政策級）？", "#38 #51",
   OK, "白宮點名＝最高級，且是第 1 天", PT, "級別同樣高——但已是第 2 天舊聞"),
  ("五大選股標準 5/5？", "#7",
   OK, "5/5 滿分", PT, "表面 4–5/5，「消息」一項虛高"),
 ]),
 ("②", "入場前 · 確認", "失敗日的重災區：8 個檢查點 7 個亮紅燈", [
  ("消息是第 1 天（未 price-in）？", "#76 #50",
   OK, "第 1 天，動能最純", BAD, "第 2 天媒體轉載，已 price-in"),
  ("市場溫度允許進攻？", "#32 #33",
   OK, "事件日動能充足", BAD, "冷市——Ross 自認的根因"),
  ("有預設計劃與客觀觸發位？", "#36 #9",
   OK, "預掛條件單：破盤前高 $7.00", BAD, "FOMO 追價，無計劃"),
  ("價格貼近 VWAP（偏離 <5%）？", "#1 #2",
   OK, "貼近 VWAP 入場", BAD, "$10.50／$11.32 離 VWAP >10%"),
  ("突破帶量（>5x）？", "#67",
   OK, "量 >5x 才追", PT, "$11.50 突破旋即回崩＝假突破"),
  ("MACD 綠燈？", "#23 #24",
   OK, "綠燈進場", BAD, "攤平發生在紅燈區"),
  ("仍在前側（未跌破均線／VWAP）？", "#75 #15",
   OK, "只在創新高階段交易", BAD, "攤平時已崩回 VWAP＝後側"),
  ("距整數關口有足夠空間？", "#52",
   OK, "整數關口前出清", BAD, "$11.50 入場距 $12 僅 4%"),
 ]),
 ("③", "倉位 · 風控", "一個熱鍵，抵消了所有選股功力", [
  ("倉位在計劃上限內？", "#36 #68",
   OK, "計劃內分批", BAD, "熱鍵誤按 9,000 股"),
  ("大贏日次日有縮倉？", "#41",
   NA, "（本身是大贏日）", BAD, "+$12k 之後反而重倉"),
  ("冷市倉位減半？", "#70 #33",
   NA, "非冷市議題", BAD, "冷市反而放最大倉"),
  ("倉位匹配流動性？", "#45",
   OK, "出場零滑價壓力", BAD, "9,000 股出場滑價 $0.47／股"),
 ]),
 ("④", "持倉 · 虧損處理", "成功日整天沒讓自己走進這一關", [
  ("絕不向下攤平？", "#42 #40",
   NA, "未觸發虧損", BAD, "$10.30／$10.40 攤平——頭號錯誤"),
  ("跌破前低即止損、不拖延？", "#22 #60",
   NA, "止損未被觸發", PT, "先以攤平拖延，最終才認賠"),
  ("Dump 訊號不接刀？", "#15 #24",
   NA, "Dump 前已離場", BAD, "訊號齊備仍接刀"),
 ]),
 ("⑤", "出場 · 收工紀律", "成功日的利潤保險絲；失敗日全部熔斷，只剩最後一根", [
  ("預設目標位分批止盈？", "#59",
   OK, "延伸棒賣半倉 → 出清", BAD, "無目標位，僥倖賣在 halt 前"),
  ("峰值回吐 50% 即收工？", "#47",
   OK, "利潤零回吐", BAD, "+$5,500 回吐 100% 轉紅"),
  ("移動止損保護浮盈？", "#78",
   OK, "全額落袋", BAD, "浮盈裸奔"),
  ("賺夠收工不回場？", "#30 #62",
   OK, "25 分鐘完事收工", BAD, "連環加碼直到轉紅"),
  ("守不住也要認賠離場？", "#60 #46",
   NA, "未需要", OK, "救贖：$10.85 認賠，避開 −$27,000"),
 ]),
]

# ─────────────────── 決策時間軸 ───────────────────
FLOW_WIN = [
 ("盤前", "掃描鎖定 DCFC", OK), ("09:28", "預掛條件單", OK),
 ("09:31", "破 $7.00 入場①", OK), ("", "延伸棒賣半倉", OK),
 ("", "縮量回踩加倉②", OK), ("", "第二延伸棒出清 $8.42", OK),
 ("10:00 前", "收工 +$12,000", OK),
]
FLOW_RED = [
 ("開盤", "連環 halt·手上已 +$1,000", OK), ("", "FOMO $10.50 誤按 9,000 股", BAD),
 ("", "$10.95 halt 前僥倖 +$4,500", PT), ("", "$11.50 頂部加回（均價 $11.32）", BAD),
 ("", "$11.90 見頂崩回 VWAP", BAD), ("", "$10.30／$10.40 向下攤平", BAD),
 ("", "$10.85 全部認賠", OK), ("收盤", "當日轉紅", BAD),
]

# ─────────────────── 簡化價格路徑（與 case-study 圖同源的錨點） ───────────────────
ANCH_WIN = [(0, 6.60), (6, 6.80), (12, 6.95), (14, 7.00), (15, 6.85), (17, 6.95),
            (18, 7.08), (21, 7.45), (24, 7.80), (27, 7.58), (30, 7.60), (31, 7.88),
            (34, 8.15), (38, 8.42), (41, 8.30), (45, 8.55), (50, 8.35), (58, 8.10),
            (66, 8.25), (75, 8.15)]
ANCH_RED = [(0, 8.20), (2, 8.42), (7, 9.15), (10, 9.90), (11, 10.30), (12, 10.50),
            (14, 10.95), (19, 11.50), (21, 11.55), (23, 11.90), (26, 11.30),
            (29, 10.80), (33, 10.35), (36, 10.45), (38, 10.85), (42, 10.40),
            (48, 9.90), (55, 9.55), (63, 8.90), (70, 8.30), (75, 8.45)]

W, H, PAD_L, PAD_R, PAD_T, PAD_B = 560, 252, 34, 12, 30, 18


def scaler(anchors, ymin, ymax):
    xmax = anchors[-1][0]
    def sx(x): return PAD_L + x / xmax * (W - PAD_L - PAD_R)
    def sy(y): return PAD_T + (ymax - y) / (ymax - ymin) * (H - PAD_T - PAD_B)
    return sx, sy


def svg_path(anchors, sx, sy, upto=None, frm=None):
    pts = [(x, y) for x, y in anchors
           if (upto is None or x <= upto) and (frm is None or x >= frm)]
    return " ".join("%s%.1f,%.1f" % ("M" if i == 0 else "L", sx(x), sy(y))
                    for i, (x, y) in enumerate(pts))


def marker(sx, sy, x, y, kind, label, above, dx=0):
    color = {OK: "var(--green)", BAD: "var(--red)", PT: "var(--amber)"}[kind]
    cx, cy = sx(x), sy(y)
    ty = cy - 14 if above else cy + 21
    return ('<circle cx="%.1f" cy="%.1f" r="7" style="fill:%s"/>' % (cx, cy, color) +
            '<text x="%.1f" y="%.1f" text-anchor="middle" style="fill:#fff;font:700 9px sans-serif">%s</text>'
            % (cx, cy + 3, GLYPH[kind]) +
            '<text x="%.1f" y="%.1f" text-anchor="middle" style="fill:var(--ink);font:600 9.5px sans-serif">%s</text>'
            % (min(max(cx + dx, 40), W - 46), ty, label))


def build_svg_win():
    sx, sy = scaler(ANCH_WIN, 6.3, 9.0)
    s = ['<svg viewBox="0 0 %d %d" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="成功日簡化價格路徑">' % (W, H)]
    s.append('<text x="%d" y="16" style="fill:var(--green);font:700 12px sans-serif">✅ 成功日 · 消息第 1 天 · Gap&amp;Go</text>' % PAD_L)
    y7 = sy(7.00)
    s.append('<line x1="%d" y1="%.1f" x2="%d" y2="%.1f" style="stroke:var(--ink2);stroke-width:1;stroke-dasharray:4 3"/>' % (PAD_L, y7, W - PAD_R, y7))
    s.append('<text x="%d" y="%.1f" style="fill:var(--ink2);font:9.5px sans-serif">盤前高 $7.00</text>' % (PAD_L + 2, y7 - 4))
    s.append('<path d="%s" style="fill:none;stroke:var(--blue);stroke-width:2.2;stroke-linejoin:round"/>' % svg_path(ANCH_WIN, sx, sy))
    s.append(marker(sx, sy, 18, 7.08, OK, "入場① 破盤前高", False))
    s.append(marker(sx, sy, 24, 7.80, OK, "延伸棒賣半倉", True))
    s.append(marker(sx, sy, 31, 7.88, OK, "回踩加倉②", False))
    s.append(marker(sx, sy, 38, 8.42, OK, "出清 $8.42", True))
    s.append('<text x="%.1f" y="%.1f" text-anchor="middle" style="fill:var(--green);font:700 11px sans-serif">之後收工 · 利潤零回吐 +$12k</text>' % (sx(58), sy(8.55) - 24))
    s.append('</svg>')
    return "".join(s)


def build_svg_red():
    sx, sy = scaler(ANCH_RED, 7.9, 12.6)
    s = ['<svg viewBox="0 0 %d %d" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="失敗日簡化價格路徑">' % (W, H)]
    s.append('<text x="%d" y="16" style="fill:var(--red);font:700 12px sans-serif">⛔ 失敗日 · 消息第 2 天 · FOMO 追高</text>' % PAD_L)
    for a, b in ((3, 7), (15, 19)):
        s.append('<rect x="%.1f" y="%d" width="%.1f" height="%d" style="fill:var(--ink2);opacity:.16"/>'
                 % (sx(a), PAD_T, sx(b) - sx(a), H - PAD_T - PAD_B))
    s.append('<text x="%.1f" y="%d" text-anchor="middle" style="fill:var(--ink2);font:8.5px sans-serif">halt</text>' % (sx(5), H - PAD_B + 10))
    s.append('<text x="%.1f" y="%d" text-anchor="middle" style="fill:var(--ink2);font:8.5px sans-serif">halt</text>' % (sx(17), H - PAD_B + 10))
    y12 = sy(12.00)
    s.append('<line x1="%d" y1="%.1f" x2="%d" y2="%.1f" style="stroke:var(--ink2);stroke-width:1;stroke-dasharray:4 3"/>' % (PAD_L, y12, W - PAD_R, y12))
    s.append('<text x="%d" y="%.1f" style="fill:var(--ink2);font:9.5px sans-serif">整數關口 $12——他在距此 4%% 處加倉</text>' % (PAD_L + 2, y12 - 4))
    s.append('<path d="%s" style="fill:none;stroke:var(--blue);stroke-width:2.2;stroke-linejoin:round"/>' % svg_path(ANCH_RED, sx, sy, upto=38))
    s.append('<path d="%s" style="fill:none;stroke:var(--red);stroke-width:1.6;stroke-dasharray:3 4;opacity:.75"/>' % svg_path(ANCH_RED, sx, sy, frm=38))
    s.append(marker(sx, sy, 12, 10.50, BAD, "FOMO $10.50 ×9,000 股", False))
    s.append(marker(sx, sy, 14, 10.95, PT, "僥倖 +$4,500", True, dx=-58))
    s.append(marker(sx, sy, 19, 11.50, BAD, "加回 $11.50", False, dx=44))
    s.append('<text x="%.1f" y="%.1f" text-anchor="middle" style="fill:var(--ink);font:600 9.5px sans-serif">高點 $11.90</text>' % (sx(23) + 46, sy(11.90) - 8))
    s.append(marker(sx, sy, 33, 10.35, BAD, "攤平 $10.30/40", False))
    s.append(marker(sx, sy, 38, 10.85, OK, "認賠 $10.85", True))
    s.append('<text x="%.1f" y="%.1f" text-anchor="end" style="fill:var(--red);font:700 10px sans-serif">若不認賠 → $8.30＝−$27,000</text>' % (sx(74), sy(8.30) - 8))
    s.append('</svg>')
    return "".join(s)


# ─────────────────── 組件 ───────────────────

def cell(kind, note):
    return ('<div class="ckcell ck-%s"><span class="cbx cb-%s">%s</span>'
            '<span class="cknote">%s</span></div>' % (kind, kind, GLYPH[kind], note))


def stage_score(rows, idx):
    c = {OK: 0, BAD: 0, PT: 0, NA: 0}
    for r in rows:
        c[r[2 + idx * 2]] += 1
    n = len(rows) - c[NA]
    return c, n


def chip(c, n, day):
    if day == "win":
        if n == 0:
            return '<span class="ckchip" style="background:rgba(100,116,139,.12);color:var(--ink2)">– 未觸發</span>'
        return '<span class="ckchip chip-win">✓ %d/%d</span>' % (c[OK], n)
    txt = "✓ %d · ! %d · ✗ %d" % (c[OK], c[PT], c[BAD])
    return '<span class="ckchip chip-red">%s</span>' % txt


def build_checklist():
    out = io.StringIO()
    for mark, title, sub, rows in STAGES:
        cw, nw = stage_score(rows, 0)
        cr, nr = stage_score(rows, 1)
        out.write('<div class="ckgroup">\n')
        out.write('<div class="ckhead"><div class="ckt"><b>%s %s</b>'
                  '<small>%s</small></div>%s%s</div>\n'
                  % (mark, title, sub, chip(cw, nw, "win"), chip(cr, nr, "red")))
        out.write('<div class="ckcols"><div></div><div>✅ 成功日 +$12k</div><div>⛔ 失敗日 紅日</div></div>\n')
        for (q, refs, wv, wn, rv, rn) in rows:
            out.write('<div class="ckrow"><div class="ckq">%s<small>%s</small></div>%s%s</div>\n'
                      % (q, refs, cell(wv, wn), cell(rv, rn)))
        out.write('</div>\n')
    return out.getvalue()


def totals():
    tw = {OK: 0, BAD: 0, PT: 0, NA: 0}
    tr = {OK: 0, BAD: 0, PT: 0, NA: 0}
    n = 0
    for _, _, _, rows in STAGES:
        for r in rows:
            n += 1
            tw[r[2]] += 1
            tr[r[4]] += 1
    return n, tw, tr


def build_flow(flow, cls):
    steps = "".join(
        '<div class="fstep"><span class="fdot fd-%s">%s</span>'
        '<span class="ftime">%s</span><span class="flab">%s</span></div>'
        % (k, GLYPH[k], t, lab) for t, lab, k in flow)
    return '<div class="vsflow %s">%s</div>' % (cls, steps)


def build_section():
    n, tw, tr = totals()
    nw = n - tw[NA]          # 成功日適用檢查點
    nr = n - tr[NA]          # 失敗日適用檢查點
    pw = round(tw[OK] / nw * 100)
    r_ok = round(tr[OK] / nr * 100)
    r_pt = round(tr[PT] / nr * 100)

    s = io.StringIO()
    s.write('<section id="vs" class="layer">\n')
    s.write('<h2>☑️ R2 · 決策檢查表對決：%d 個檢查點，一眼看懂為什麼一天贏一天輸</h2>\n' % n)
    s.write('<div class="hint">把 R1 的 83 條規則檢核濃縮成 %d 個「入場前可以問自己」的檢查點，'
            '按交易生命週期分五關。左格＝成功日、右格＝失敗日：綠＝✓ 通過、紅＝✗ 違規、橙＝! 僥倖/邊緣、灰＝– 當日未觸發。'
            '每一格的規則出處與完整證據見下方 R1 檢核章節。</div>\n' % n)

    # 對決記分卡
    s.write('<div class="vshead">\n')
    s.write('<div class="vscard vswin"><div class="vsday">✅ 成功日</div>'
            '<div class="vspl">+$12,000</div><div class="vsmeta">2022/2/8 · 消息第 1 天 · 兩進兩出 25 分鐘</div>'
            '<div class="vsbar"><span class="vb-ok" style="width:%d%%"></span></div>'
            '<div class="vsscore">%d／%d 個適用檢查點全部通過</div></div>\n'
            % (pw, tw[OK], nw))
    s.write('<div class="vsvs">VS</div>\n')
    s.write('<div class="vscard vsred"><div class="vsday">⛔ 失敗日</div>'
            '<div class="vspl">轉紅（曾 +$5,500）</div><div class="vsmeta">消息第 2 天 · 追高→加倉→攤平→認賠</div>'
            '<div class="vsbar"><span class="vb-ok" style="width:%d%%"></span>'
            '<span class="vb-pt" style="width:%d%%"></span></div>'
            '<div class="vsscore">%d ✓ · %d ! · <b>%d ✗</b>（%d 個適用檢查點）</div></div>\n'
            % (r_ok, r_pt, tr[OK], tr[PT], tr[BAD], nr))
    s.write('</div>\n')

    # 並排簡化路徑圖
    s.write('<div class="vscharts"><div class="vschart">%s</div><div class="vschart">%s</div></div>\n'
            % (build_svg_win(), build_svg_red()))
    s.write('<div class="hint" style="margin-top:4px">簡化路徑示意（錨點與 case-study 重構圖同源；'
            '完整 K 線＋量能＋VWAP 圖見上方成敗案例章節）。失敗日虛線段＝認賠出場後的續跌——若不出場即 −$27k。</div>\n')

    # 決策時間軸
    s.write('<h3 class="r1h">🕒 決策時間軸：每一步的顏色</h3>\n')
    s.write(build_flow(FLOW_WIN, "flow-win"))
    s.write(build_flow(FLOW_RED, "flow-red"))

    # 五關檢查表
    s.write('<h3 class="r1h">☑️ 五關決策檢查表（左：成功日 · 右：失敗日）</h3>\n')
    s.write(build_checklist())

    # 收尾
    s.write('<div class="out"><b>怎麼用這張表：</b>下次入場前把 ①② 關的 12 個檢查點當作硬性關卡——'
            '成功日 12 格全綠、失敗日在第 ② 關就已經 7 格紅。第 ③④⑤ 關的紅格全部是第 ② 關放行後的連鎖反應；'
            '唯一的例外是最後一格「認賠離場」——它是失敗日僅有的綠格，也是紅日沒有變成 −$27,000 災難日的全部原因。</div>\n')
    s.write('</section>\n')
    return s.getvalue()


R2_CSS = """
  /* ── R2 檢查表對決 ── */
  .vshead{display:grid; grid-template-columns:1fr auto 1fr; gap:12px; align-items:stretch; margin:14px 0}
  .vscard{border-radius:12px; padding:14px 16px; border:1.5px solid}
  .vswin{background:rgba(22,163,74,.08); border-color:rgba(22,163,74,.45)}
  .vsred{background:rgba(220,38,38,.07); border-color:rgba(220,38,38,.45)}
  .vsday{font-size:13px; font-weight:700}
  .vswin .vsday{color:var(--green)} .vsred .vsday{color:var(--red)}
  .vspl{font-size:24px; font-weight:800; margin:2px 0}
  .vswin .vspl{color:var(--green)} .vsred .vspl{color:var(--red)}
  .vsmeta{font-size:11.5px; color:var(--ink2)}
  .vsbar{display:flex; height:10px; border-radius:99px; overflow:hidden; background:rgba(220,38,38,.25); margin:10px 0 6px}
  .vswin .vsbar{background:rgba(100,116,139,.2)}
  .vb-ok{background:var(--green)} .vb-pt{background:var(--amber)}
  .vsscore{font-size:12px; color:var(--ink2)}
  .vsvs{align-self:center; font-size:16px; font-weight:800; color:var(--ink2)}
  .vscharts{display:grid; grid-template-columns:repeat(auto-fit,minmax(300px,1fr)); gap:12px; margin-top:6px}
  .vschart{border:1px solid var(--line); border-radius:10px; background:var(--bg); padding:6px}
  .vschart svg{width:100%; height:auto; display:block}
  .vsflow{display:flex; gap:0; overflow-x:auto; padding:10px 2px 6px; position:relative}
  .fstep{flex:1 0 110px; min-width:110px; position:relative; text-align:center; padding-top:4px}
  .fstep:not(:last-child):after{content:""; position:absolute; top:15px; left:calc(50% + 14px);
     width:calc(100% - 28px); height:2px; background:var(--line)}
  .fdot{display:inline-flex; align-items:center; justify-content:center; width:24px; height:24px;
     border-radius:50%; color:#fff; font-weight:800; font-size:12px; position:relative; z-index:1}
  .fd-ok{background:var(--green)} .fd-bad{background:var(--red)} .fd-part{background:var(--amber)}
  .ftime{display:block; font-size:10px; color:var(--ink2); margin-top:3px; min-height:13px}
  .flab{display:block; font-size:11px; line-height:1.35; margin-top:1px; padding:0 4px}
  .flow-win{border-left:3px solid var(--green); margin-bottom:4px}
  .flow-red{border-left:3px solid var(--red)}
  .ckgroup{border:1px solid var(--line); border-radius:12px; margin-top:14px; overflow:hidden}
  .ckhead{display:flex; align-items:center; gap:8px; padding:10px 14px; background:var(--bg)}
  .ckt{flex:1} .ckt b{font-size:13.5px}
  .ckt small{display:block; color:var(--ink2); font-size:11.5px}
  .ckchip{font-size:11.5px; font-weight:700; padding:3px 10px; border-radius:99px; white-space:nowrap}
  .chip-win{background:rgba(22,163,74,.14); color:var(--green)}
  .chip-red{background:rgba(220,38,38,.12); color:var(--red)}
  .ckcols{display:grid; grid-template-columns:1.15fr 1fr 1fr; gap:1px; font-size:11px; color:var(--ink2);
     padding:6px 14px 0; font-weight:600}
  .ckrow{display:grid; grid-template-columns:1.15fr 1fr 1fr; gap:8px; padding:7px 14px; border-top:1px solid var(--line); align-items:stretch}
  .ckq{font-size:13px; font-weight:600; align-self:center}
  .ckq small{display:block; color:var(--ink2); font-weight:400; font-size:11px}
  .ckcell{display:flex; gap:8px; align-items:flex-start; border-radius:8px; padding:7px 9px}
  .ck-ok{background:rgba(22,163,74,.10)} .ck-bad{background:rgba(220,38,38,.09)}
  .ck-part{background:rgba(217,119,6,.12)} .ck-na{background:rgba(100,116,139,.08)}
  .cbx{flex:none; display:inline-flex; align-items:center; justify-content:center; width:20px; height:20px;
     border-radius:6px; color:#fff; font-weight:800; font-size:12.5px; margin-top:1px}
  .cb-ok{background:var(--green)} .cb-bad{background:var(--red)}
  .cb-part{background:var(--amber)} .cb-na{background:#94a3b8}
  .cknote{font-size:12px; line-height:1.4; color:var(--ink)}
  @media (max-width:760px){
    .vshead{grid-template-columns:1fr} .vsvs{display:none}
    .ckcols{display:none}
    .ckrow{grid-template-columns:1fr 1fr}
    .ckq{grid-column:1/-1}
  }
"""


def main():
    html = open(SRC, encoding="utf-8").read()

    html = html.replace(
        "<title>Ross Momentum 交易工具包 R1（離線完整版・含 83 條規則成敗檢核）</title>",
        "<title>Ross Momentum 交易工具包 R2（離線完整版・檢查表對決＋83 條規則檢核）</title>", 1)

    i = html.find("</style>")
    assert i > 0, "style block not found"
    html = html[:i] + R2_CSS + html[i:]

    nav_anchor = '<a href="#r1">'
    assert nav_anchor in html, "r1 nav anchor not found"
    html = html.replace(nav_anchor, '<a href="#vs">☑️ 檢查表對決</a>\n  ' + nav_anchor, 1)

    sec_anchor = '<section id="r1"'
    assert sec_anchor in html, "r1 section not found"
    html = html.replace(sec_anchor, build_section() + "\n" + sec_anchor, 1)

    with open(DST, "w", encoding="utf-8") as f:
        f.write(html)
    print("written:", DST, len(html), "chars")


if __name__ == "__main__":
    main()
