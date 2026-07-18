# 主系統（國軍軍購案分析系統）的維護主檔與驗證腳本

產生主系統成品（專案根目錄，內嵌 SheetJS 0.18.5）的模板與獨立驗證腳本。**現行版為 `國軍軍購案分析系統v1150718.html`（約 943KB）**。

## 檔案

| 檔案 | 說明 |
|---|---|
| `template_v1150718.html` | **現行維護主檔**。含 `<!--SHEETJS_PLACEHOLDER-->`，未內嵌 SheetJS（約 82KB，可直接編輯）。自 v1150717 template 複製後加風險分布氣泡圖與排名表摺疊 |
| `template_v1150717.html` | 前一版維護主檔（評分模組第 1~3 步），保留 |
| `expected_score.js` | 風險**評分**的預期答案腳本（每案 exec/fin/total）。**與 HTML 獨立實作** |
| `expected_chart.js` | 風險分布圖**幾何**的預期答案腳本（每案泡泡座標/半徑、分區、前 10 名、泡泡數）。**與 HTML 獨立實作**，重用評分核心 |

## 由來：主系統原本沒有 template

v1150715（含）以前的主系統只有內嵌成品，**沒有**未內嵌的維護主檔。1150717 做風險評分時，從 `國軍軍購案分析系統v1150715.html` 把 SheetJS 區塊抽回成 placeholder，才得到這份可編輯 template。抽取要點：

- v1150715 的 SheetJS 是**單一乾淨區塊**：第一個 `<script>` ~ 第一個 `</script>`（該檔第 227~253 行），內含整份 `xlsx.full.min.js`；第二個 script 才是 app code。
- 抽出的 `xlsx.full.min.js`（約 639KB）另存，區塊換成 `<script><!--SHEETJS_PLACEHOLDER--></script>`。
- 驗證：template 位元組 ＝ 原檔 − SheetJS ＋ placeholder；重新內嵌後與原檔逐位元組相同（1150717 已驗證往返一致）。

## 要改主系統就改現行 template，不要改根目錄那個 943KB 的檔

改完 `template_v1150718.html` 後，用 PowerShell 把 SheetJS 內嵌回去（**不可用 CDN**，目標環境無網路）：

```powershell
$sp  = "<scratchpad 或任一放 xlsx.full.min.js 的路徑>"   # 可從任一已建置 HTML 抽出
$tpl = "主系統腳本\template_v1150718.html"
$out = "國軍軍購案分析系統v1150718.html"
$t = [System.IO.File]::ReadAllText($tpl, [System.Text.Encoding]::UTF8)
$j = [System.IO.File]::ReadAllText((Join-Path $sp "xlsx.full.min.js"), [System.Text.Encoding]::UTF8)
$t = $t.Replace("<!--SHEETJS_PLACEHOLDER-->", $j)
[System.IO.File]::WriteAllText($out, $t, (New-Object System.Text.UTF8Encoding($false)))
```

`.Replace()` 是刻意的——避免把 639KB 的 JS 讀進編輯器。

## 驗證

```bash
node expected_score.js     # 評分：每案 exec/fin/total
node expected_chart.js     # 圖幾何：每案泡泡座標/半徑、分區、前 10 名
# 兩者皆需同目錄有 xlsx.full.min.js（可從已建置 HTML 抽出）
```

以 `執行中軍購案個案管制表(演示用).xlsx`（451 筆）算出預期答案，再與瀏覽器實跑逐案比對（起靜態伺服器 → 頁面 `fetch` 演示檔 → `runAnalysis()` → `runScoring()` → 讀 `caseScores`／圖上 `circle[data-ri]` 的 cx/cy/r 比對）。

**兩腳本都必須與 HTML 獨立實作**——共用程式碼就驗不出東西。

### 評分驗收基準（演示檔 451 筆，1150717 已通過）

- 有風險分數 **172** 案／未命中（總分 0）**279** 案
- 全部風險總分加總 **1466.09**（exec 加總 764.86、fin 加總 701.21）
- 軍投 Top 1＝TWMYUY（執 22.06／財 5.44／總 27.50）；作維 Top 1＝TWBKQK（執 35.43／財 0／總 35.43）
- 20 控制點計數與 v1150715 完全相同（回歸）：1-3+1-4=30、1-7+1-8=32、2-3+2-4=19、2-7+2-8=54、2-10=21…
- 明細頁欄序（第二階段＋評分並存）＝ base 18 ＋ 檢討態樣/說明（18,19）＋ 本控制點分/風險總分（20,21），共 22 欄
- 軍投 1-10 明細頁「本控制點分」顯示「併入1-3/1-4」（不獨立計分）

### 圖幾何驗收基準（演示檔 451 筆，1150718 逐案 0 誤差通過）

- 軍投 **63** 泡泡（分區 綠44/黃15/橘4/紅0）＋0 分註記 **63**；作維 **109** 泡泡（綠79/黃25/橘4/紅1）＋0 分註記 **216**；合計 63+63+109+216=451
- 座標／半徑：TWMYUY (267.6, 378.3) r22.4、TWBKQK (394.6, 430.0) r16.4（貼 X 軸、被 clipPath 裁下半顆）
- 版面幾何：繪圖區正方形 PL58/PR438/PT50/PB430；`chartX/chartY/bubbleR/zoneOf` 須與 HTML 完全一致
- PNG 匯出 940×1000（2 倍）、SVG 可開；發價值 ≤0/缺值 → 最小半徑 4px 不 NaN

## 評分機制的來源

錨點與三項結構決議、執行/財務分類，見 CLAUDE.md「風險評分」節與 `控制點篩選結果為所有軍購案評分的標準.docx`。
**改動評分前務必先讀那一節**——成對合併、數值錨點法、軍投 1-10 併入 1-3/1-4 不計分，這三件都是推導出來的，不是偏好。
