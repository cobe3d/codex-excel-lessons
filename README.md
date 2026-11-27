# 用 Codex 做 Excel

教學生用 Codex 產生 Excel 檔的課程教材：一頁學生講義、11 個分級專案，以及 5 個可以對照的範例 Excel 檔。

**線上講義：** https://cobe3d.github.io/codex-excel-lessons/

## 內容

| 路徑 | 說明 |
|---|---|
| [`index.html`](index.html) | 學生講義網頁：Codex 的運作方式、提示詞寫法、11 個專案（附可複製的提示詞）、驗收清單、常見問題 |
| [`範例檔/`](範例檔) | 5 個範例 Excel 檔，每個檔案的第一個工作表「說明」都寫有學習重點、提示詞和驗收方式 |
| [`tools/`](tools) | 產生範例檔的 Python 程式，以及用 Excel 重新計算、檢查公式錯誤的 PowerShell 腳本 |

## 範例檔

| 檔案 | 難度 | 示範的功能 |
|---|---|---|
| 專案1_班級成績計算表 | 入門 | SUM、RANK、等第判斷、不及格整列變紅、分數只能填 0–100、長條圖 |
| 專案2_個人記帳本 | 入門 | 類別下拉選單、SUMIF 統計、預算超支提醒、圓餅圖 |
| 專案5_圖書借閱紀錄 | 中級 | 應還日期、逾期天數與罰款、狀態變色、設定值獨立成一格 |
| 專案8_學生資料查詢表 | 進階 | 選學號後用 INDEX＋MATCH 帶出資料、班排名、比較長條圖 |
| 專案11_自動批改測驗卷 | 進階 | 作答後即時顯示 ✔／✘、計分和評語 |

淺黃色底的儲存格是可以修改的輸入格，其他欄位都是公式。範例中的學生姓名和成績都是隨機產生的假資料。

圖書借閱紀錄的「今天日期」固定為 2026/11/25，方便課堂示範時出現各種借閱狀態；實際使用時把該格改成 `=TODAY()` 即可。

## 重新產生範例檔

需要 Python 3 和 `openpyxl`：

```powershell
pip install openpyxl
python tools/build_examples.py 範例檔
```

openpyxl 寫入的公式沒有計算結果，用 Excel 開啟一次就會計算。若電腦裝有 Excel，可以用下面的腳本一次重新計算全部檔案並檢查公式錯誤：

```powershell
powershell -ExecutionPolicy Bypass -File tools/recalc.ps1 -Dir 範例檔
```
