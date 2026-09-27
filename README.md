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
- **累計記錄自動接續**：如果沒有設定 Google 試算表（見下方教學），每次產生完文件後，網頁會提供一個「累計記錄.json」下載，下次上傳新資料時把這個 json 檔案一起上傳到側邊欄，就能自動接續計算累計數字。如果已經設定好 Google 試算表串接，這一步完全自動化，不需要手動上傳/下載任何檔案。
- 月報表的「憑證起迄序號」會自動找出共同前綴，只顯示流水號後4碼（例如 `EYG10099EYG21699_B2-3_0012 ~ 0041`），不同標段（B1、B2、B5...）前綴不同時會各自正確顯示。
- **排除異常聯單**：上傳 CSV 後，網頁會列出所有聯單讓你勾選要排除的（例如退車、拒收）。排除後：
  - 每日出場紀錄：該列仍正常顯示（車輛確實有出場），但不計入當日/累計總量
  - 運送時間一覽表：只保留出場日期/時間，入場時間與運載數量留空，土資場名稱欄改填你輸入的排除原因
  - 統計月報表：扣除該筆數量，並在備註欄寫上「聯單號碼 + 原因」

## Google 試算表串接（選用，設定好之後累計記錄全自動）

設定完成後，程式會自動把「每月累計立方公尺」寫進你指定試算表的「累計總表」分頁，把「排除的異常聯單」寫進「異常退車記錄」分頁，不用再手動上傳/下載 JSON 記錄檔。

### 步驟一：建立 Google 服務帳號（Service Account）

1. 前往 [Google Cloud Console](https://console.cloud.google.com/)，建立一個新專案（或使用現有專案）
2. 左側選單「API 和服務」→「已啟用的 API 和服務」→ 分別搜尋並啟用：
   - **Google Sheets API**
   - **Google Drive API**
3. 左側選單「IAM 與管理」→「服務帳戶」→「建立服務帳戶」，名稱隨意（例如 `docgen-bot`），角色可以不用選，直接完成
4. 建好之後點進這個服務帳戶 →「金鑰」分頁 →「新增金鑰」→「建立新的金鑰」→ 選 **JSON** → 下載金鑰檔案（存好，之後要用）

### 步驟二：把試算表分享給服務帳號

1. 打開這把 JSON 金鑰檔案，找到 `"client_email"` 那一行，複製那個 email（長得像 `xxx@xxx.iam.gserviceaccount.com`）
2. 打開你現有的 Google 試算表，右上角「共用」，把這個 email 加進去，權限選「編輯者」

### 步驟三：把金鑰內容貼到 Streamlit Cloud 的 Secrets

1. 進入你的 App 頁面，右下角「⋮」或「Manage app」→「Settings」→「Secrets」
2. 貼入以下格式（把 JSON 金鑰檔案裡對應的值填進去，`private_key` 要保留 `\n` 換行符號，整段用三個引號包起來）：

   ```toml
   GOOGLE_SHEET_ID = "你的試算表ID"

   [gcp_service_account]
   type = "service_account"
   project_id = "你的project_id"
   private_key_id = "你的private_key_id"
   private_key = """-----BEGIN PRIVATE KEY-----
   ...金鑰內容...
   -----END PRIVATE KEY-----
   """
   client_email = "xxx@xxx.iam.gserviceaccount.com"
   client_id = "你的client_id"
   auth_uri = "https://accounts.google.com/o/oauth2/auth"
   token_uri = "https://oauth2.googleapis.com/token"
   auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
   client_x509_cert_url = "你的client_x509_cert_url"
   ```

   `GOOGLE_SHEET_ID` 是你試算表網址中間那一段，例如網址是
   `https://docs.google.com/spreadsheets/d/1AbCDeFGhijklmnoPQRstuVWxyz/edit`，
   ID 就是 `1AbCDeFGhijklmnoPQRstuVWxyz`。

3. 儲存後，Streamlit Cloud 會自動重啟 App。重啟後側邊欄會顯示「✅ 已連接 Google 試算表」，代表設定成功。

> 如果沒有設定這些 Secrets，程式會自動退回原本手動上傳/下載 JSON 記錄檔的方式，不影響正常使用。

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
├── sheets.py           # Google 試算表串接（選用）
├── requirements.txt    # Python 套件需求
├── sample_data.csv     # 範例聯單資料（可用來測試）
└── templates/
    ├── daily_record_template.docx
    ├── monthly_report_template.docx
    └── time_overview_template.docx
```
