# 土石方聯單 Word 自動產生器

上傳「聯單資料列表」CSV，自動產生三種 Word 文件：

| 檔案 | 產生規則 |
|---|---|
| 每日出場紀錄 | CSV 裡每出現一個「出場日期」就產生一份 |
| 統計月報表 | CSV 裡每出現一個「年-月」就產生一份，1~31 日自動對應 |
| 運送時間一覽表 | 與每日出場紀錄相同，一天一份，逐筆列出進出場時間 |

## 欄位對應邏輯

- 車牌號碼／車號：`出場車頭車號/出場車斗車號`（或進場對應欄位）合併顯示
- 運送數量：優先取「實際出土量」，缺值則用「載運土方量」
- 4 項檢查勾選：CSV「狀態」＝已完成 → 自動打勾，其餘留空
- 施工廠商：固定寫入「力勤工程實業有限公司」（如需更改，修改 `docgen.py` 內的 `CONTRACTOR_NAME`）
- 工程名稱：預設抓 CSV 的「工程名稱」欄，也可在網頁側邊欄手動輸入覆蓋
- 月報表「期初累計」：因為月報表是累計性質的表格，若你上傳的 CSV 不是從月初第一天開始，請在側邊欄手動輸入「期初累計立方公尺」，程式才能算出正確的累計數字
- 每日出場紀錄的「累計」欄位，會依同一個月份、按日期先後順序自動往下累加（不需要每天都重新輸入期初累計，只要當月第一次上傳資料時填一次即可；之後同一個月的資料建議一次補齊或依序上傳、並手動延續累計基準）

## 本機測試

```bash
pip install -r requirements.txt
streamlit run app.py
```

瀏覽器開啟 http://localhost:8501，上傳 `sample_data.csv` 測試即可看到效果。

## 部署到 GitHub + Streamlit Community Cloud

1. 在 GitHub 建立新的 repository（例如 `soil-transport-docgen`），把這個資料夾所有檔案 push 上去：

   ```bash
   git init
   git add .
   git commit -m "init: 土石方聯單 word 自動產生器"
   git branch -M main
   git remote add origin https://github.com/<你的帳號>/soil-transport-docgen.git
   git push -u origin main
   ```

2. 前往 [share.streamlit.io](https://share.streamlit.io)，用 GitHub 帳號登入
3. 點選「New app」，選擇剛剛的 repository、branch（main）、主檔案填 `app.py`
4. 按 Deploy，等待幾分鐘後就會產生一個網址（例如 `https://xxx.streamlit.app`），之後就能直接在瀏覽器上傳 CSV 使用，不需要再打開電腦執行程式

## 之後如果 Word 樣板格式有變動

只要把新的 `.docx` 檔案取代 `templates/` 資料夾內同名檔案，重新 commit + push 到 GitHub，Streamlit Cloud 會自動重新部署，不需要改程式碼（除非表格欄位順序或列數配置也跟著變動，那就需要調整 `docgen.py` 裡對應的欄位索引）。

> 注意：範本必須是 `.docx`（不是舊版 `.doc`）。如果之後拿到的新樣板是 `.doc`，需要先用 Word 另存新檔為 `.docx` 格式再放進 `templates/` 資料夾。

## 檔案結構

```
.
├── app.py              # Streamlit 網頁介面
├── docgen.py           # 核心邏輯：CSV 解析 + Word 樣板填值
├── requirements.txt    # Python 套件需求
├── sample_data.csv     # 範例聯單資料（可用來測試）
└── templates/
    ├── 每日出場紀錄.docx
    ├── 統計月報表.docx
    └── 運送時間一覽表.docx
```
