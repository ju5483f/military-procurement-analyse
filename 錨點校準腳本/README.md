# 錨點校準工具的原始檔與驗證腳本

產生 `錨點校準工具v1150717.html`（專案根目錄）的模板與獨立驗證腳本。

## 檔案

| 檔案 | 說明 |
|---|---|
| `template_v1150717.html` | **維護主檔**。含 `<!--SHEETJS_PLACEHOLDER-->`，未內嵌 SheetJS（57KB，可直接編輯） |
| `expected_calib.js` | 預期答案腳本。**與 HTML 獨立實作**，用於驗證 HTML 算得對不對 |

## 要改工具就改模板，不要改根目錄那個 939KB 的檔

改完 `template_v1150717.html` 後，用 PowerShell 把 SheetJS 內嵌回去（**不可用 CDN**，目標環境無網路）：

```powershell
$sp  = "<scratchpad 路徑>"          # 需有 xlsx.full.min.js（可從 v1150715.html 抽出）
$tpl = "錨點校準腳本\template_v1150717.html"
$out = "錨點校準工具v1150717.html"
$t = [System.IO.File]::ReadAllText($tpl, [System.Text.Encoding]::UTF8)
$j = [System.IO.File]::ReadAllText((Join-Path $sp "xlsx.full.min.js"), [System.Text.Encoding]::UTF8)
$t = $t.Replace("<!--SHEETJS_PLACEHOLDER-->", $j)
[System.IO.File]::WriteAllText($out, $t, (New-Object System.Text.UTF8Encoding($false)))
```

`.Replace()` 是刻意的——避免把 882KB 的 JS 讀進編輯器。

## 驗證

```bash
node expected_calib.js     # 需同目錄有 xlsx.full.min.js
```

以 `執行中軍購案個案管制表(演示用).xlsx`（451 筆）算出 A／B／C 的預期數字，再與瀏覽器實跑的結果逐格比對。基準值見 CLAUDE.md 的「錨點校準工具 v1150717」節。

**此腳本必須與 HTML 獨立實作**——共用程式碼就驗不出東西。它確實抓過錯：早期版本把待平衡支付值軸算進「金額上錨」的去重貼頂計數，造成恆定 +1 偏移（誤得「10億→2案」，正解為 1 案）。徵兆是上錨拉到 500 億時貼頂數仍不歸零。

## 設計要點

- **16 個評分維度**（成對控制點已合併），來源與理由見 CLAUDE.md「風險評分（設計中）」
- `SCORE_DIMS` 是單一資料結構，驅動全部運算與呈現——改維度只改這個陣列
- 命中判定用控制點門檻（`P`，固定），計分用錨點（`A`，可調）；**兩者互不影響**，改錨點不會改變命中群體
- 比率軸線性換算、金額軸取對數（金額跨 3 個數量級，線性會讓中小案全擠在 1 分附近）
- `recommend()` 的三個不可誤改性質（只從階梯取值／不依賴目前參數／不自動套用）與 `A.jHi` 的百分數陷阱，見 CLAUDE.md「錨點校準工具 v1150717」節
