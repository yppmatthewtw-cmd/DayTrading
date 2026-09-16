# DayTrading — 宏觀週期分析與日內交易框架

以「四大力量」框架定位當前市場所處的牛熊週期階段，並輸出可操作的日內交易劇本。

## 四大力量

| 力量 | 內容 |
|---|---|
| ① 宏觀流動性 | 息口、聯儲局政策週期 |
| ② 政治週期 | 總統四年週期、期中選舉 |
| ③ 技術革命 | 產業資本開支超級週期 |
| ④ 黑天鵝 | 地緣事件、戰爭、油價衝擊 |

## 產出檔案 (`output/`)

| 檔案 | 內容 |
|---|---|
| `2026-09-09_四大力量牛熊週期_日內交易框架.xlsx` | 8 個工作表 + 5 個原生圖表：現況定位、四大力量各自的歷史週期表與轉折觸發事件、領先指標儀表板、日內交易劇本、資料來源 |
| `2026-09-09_四大力量牛熊週期_日內交易報告.docx` | 同一分析的敘事版報告（10 個表格） |
| `2026-09-08_高交額20大_低波幅20分鐘均線篩選.xlsx` | 2026-09-08 美股高交額 20 大，篩選「波幅小 + 20 分鐘均線平穩向上」的個股，含方法論與資料限制說明 |
| `2026-09-09_盤中掃描_低波幅20分鐘均線.xlsx` | 2026-09-09 盤中掃描（美東 11:12）：篩選結論、板塊資金流、當日背景與剩餘時段時間表 |
| `2026-09-16_FOMC日盤中掃描_低波幅10分鐘均線.xlsx` | 2026-09-16 FOMC 決議日掃描（美東 09:53）：為何此時點掃描不可靠、候選失效分析、四大力量框架更新、FOMC 日時間表 |

依偏好設定，只產生 Excel / Word 原始檔，不產生 PDF 或圖片。

## 腳本 (`scripts/`)

```bash
pip install openpyxl python-docx
python3 scripts/build_macro_cycle_report.py     # → 四大力量週期 Excel
python3 scripts/build_word_report.py            # → Word 報告
python3 scripts/build_0908_turnover_screen.py   # → 9/8 高交額篩選 Excel
python3 scripts/build_0909_intraday_scan.py     # → 9/9 盤中掃描 Excel
python3 scripts/build_0916_fomc_scan.py         # → 9/16 FOMC 日掃描 Excel
python3 scripts/add_tradingview_links.py        # ← 最後一步：為所有 Excel 的 ticker 加 TradingView 連結
```

> ⚠️ `add_tradingview_links.py` **必須最後執行**。上面的 build 腳本會整份覆寫 xlsx，重跑任何一個都會沖走連結。

### `scripts/add_tradingview_links.py` — TradingView 超連結

為 `output/*.xlsx` 內的 ticker 加上 TradingView 圖表連結（`https://www.tradingview.com/chart/?symbol=<交易所>%3A<代號>`）。

| 情況 | 處理方式 |
|---|---|
| 短標籤格、只含 1 個 ticker（如 `XOM`、`SMH 半導體`、`能源 (XLE)`） | 直接在原格加連結，保留原有粗體/底色，只轉藍色加底線 |
| 敘述段落（>30 字元） | 不動 — 避免把整段文字變成連結 |
| 單格內含多個 ticker（如 `TSLA / MU / PLTR / CRWV / IREN`） | Excel 不支援單一儲存格內的部分文字連結，改由索引表覆蓋 |
| 指標名稱（`Brent 原油`、`10 年期美債孳息`、`標普 500`、`黃金`、`道指`、`納指`、`羅素2000`） | 完全相符時連結至對應的 TVC / SP / DJ 代號 |

每個檔案另加一張 **`TV_連結`** 索引表，列出該檔案出現過的所有代號、名稱、可點擊連結與出現位置。

腳本可重複執行（已有連結的格會略過，索引表每次重建）。

### `scripts/screen_20min_ma.py` — 20 分鐘均線篩選器

實測版的篩選邏輯。需要行情權限的環境（本 repo 的 CI/容器環境行情 API 被封鎖）。

```bash
pip install yfinance pandas numpy
python3 scripts/screen_20min_ma.py --date 2026-09-08 --top 20 --out screen.csv
python3 scripts/screen_20min_ma.py --date 2026-09-16 --bar 10min --top 20   # 改用 10 分鐘週期
```

`--bar` 可指定重採樣週期（`10min` / `20min` / `30min`）。週期減半，每根 K 的隨機波動約放大 √2 倍，訊噪比下降；盤中資料不足以形成均線時，該股回傳「資料不足」而非勉強給出斜率。

篩選邏輯：

1. 取當日全市場成交額（`close × volume`）前 N 名
2. 取 1 分鐘 K 線 → 重採樣為 20 分鐘 K 線 → 計算移動平均
3. **平穩度** = 均線一階差分為正的比例（`rising_ratio ≥ 0.70`）
4. **上升度** = 對均線做線性迴歸，要求標準化斜率 `> 0` 且 `R² ≥ 0.60`
5. **波幅** = 日內 `(High−Low)/Close`，以及 20 分鐘收益率標準差
6. 排序 = 三個門檻全部通過者，再以波幅由小至大排列

加入 `R²` 門檻的原因：「平穩向上」是兩個獨立條件。只看斜率會把「沖高回落後再拉起」的鋸齒走勢誤判為合格。

## 免責聲明

本 repo 為市場資料整理與交易框架研究，非投資建議、非買賣要約。歷史規律不保證未來重演。日內交易涉及高風險並可能損失全部本金。
