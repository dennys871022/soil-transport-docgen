# -*- coding: utf-8 -*-
"""
sheets.py
---------
Google 試算表串接：把每月累計數字、異常排除紀錄寫進使用者現有的試算表（新增分頁）。

需要在 Streamlit 的 Secrets 設定：
    GOOGLE_SHEET_ID = "你的試算表ID（網址中 /d/ 與 /edit 之間那一段）"

    [gcp_service_account]
    type = "service_account"
    project_id = "..."
    private_key_id = "..."
    private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
    client_email = "xxx@xxx.iam.gserviceaccount.com"
    client_id = "..."
    auth_uri = "https://accounts.google.com/o/oauth2/auth"
    token_uri = "https://oauth2.googleapis.com/token"
    auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
    client_x509_cert_url = "..."

並且要把該服務帳號的 email（client_email）加到你的 Google 試算表「共用」名單，給予「編輯者」權限。
"""

from datetime import datetime

CUMULATIVE_SHEET_NAME = "累計總表"
CUMULATIVE_HEADER = ["年月", "累計立方公尺", "最後更新時間"]
# 注意：「累計立方公尺」欄位是「累計到該月月底為止的總量」（跨月持續往上加，不會每月歸零），
# 不是那個月自己單月的運送量。若懷疑舊資料存錯（例如看起來像是每月各自獨立的量），
# 使用 docgen.fix_legacy_cumulative() 做一次性修正。

EXCLUDED_SHEET_NAME = "異常退車記錄"
EXCLUDED_HEADER = ["日期", "聯單序號", "車號", "原因", "數量(m3)", "記錄時間"]

PROCESSED_SHEET_NAME = "已處理聯單"
PROCESSED_HEADER = ["聯單序號", "記錄時間"]

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


class SheetsNotConfigured(Exception):
    """尚未設定 Google 試算表連線（缺少 secrets）。"""


def is_configured(st) -> bool:
    """檢查 Streamlit secrets 裡有沒有設定好 Google 試算表連線資訊。"""
    try:
        return "gcp_service_account" in st.secrets and "GOOGLE_SHEET_ID" in st.secrets
    except Exception:  # noqa: BLE001
        return False


def get_spreadsheet(st):
    """依 st.secrets 裡的服務帳號憑證連線，回傳目標試算表物件。"""
    if not is_configured(st):
        raise SheetsNotConfigured("尚未在 Streamlit Secrets 設定 gcp_service_account 或 GOOGLE_SHEET_ID。")

    import gspread
    from google.oauth2.service_account import Credentials

    info = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    client = gspread.authorize(creds)
    sheet_id = st.secrets["GOOGLE_SHEET_ID"]
    return client.open_by_key(sheet_id)


def _get_or_create_worksheet(spreadsheet, title, header):
    import gspread
    try:
        ws = spreadsheet.worksheet(title)
    except gspread.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title=title, rows=200, cols=max(6, len(header)))
        ws.append_row(header)
    return ws


def load_cumulative_from_sheet(spreadsheet) -> dict:
    """從「累計總表」分頁讀取目前每個月的累計立方公尺，回傳 {"YYYY-MM": 數字}。"""
    ws = _get_or_create_worksheet(spreadsheet, CUMULATIVE_SHEET_NAME, CUMULATIVE_HEADER)
    records = ws.get_all_records()
    result = {}
    for r in records:
        ym = str(r.get("年月", "")).strip()
        if not ym:
            continue
        try:
            result[ym] = float(r.get("累計立方公尺", 0) or 0)
        except (TypeError, ValueError):
            continue
    return result


def save_cumulative_to_sheet(spreadsheet, state: dict):
    """把更新過的累計狀態寫回「累計總表」分頁（該月份已存在就更新，不存在就新增一列）。"""
    ws = _get_or_create_worksheet(spreadsheet, CUMULATIVE_SHEET_NAME, CUMULATIVE_HEADER)
    records = ws.get_all_records()
    existing_row_of = {}
    for idx, r in enumerate(records):
        ym = str(r.get("年月", "")).strip()
        if ym:
            existing_row_of[ym] = idx + 2  # +2: 第1列是標題，get_all_records 從第2列開始且是0-index

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    updates = []
    appends = []
    for ym, total in state.items():
        if ym in existing_row_of:
            row_no = existing_row_of[ym]
            updates.append({"range": f"B{row_no}:C{row_no}", "values": [[total, now_str]]})
        else:
            appends.append([ym, total, now_str])

    if updates:
        ws.batch_update(updates)
    if appends:
        ws.append_rows(appends)


def log_excluded_tickets(spreadsheet, excluded_log: list):
    """把這次處理到的異常排除聯單，加進「異常退車記錄」分頁（會自動跳過已經記錄過的聯單號碼，避免重複）。"""
    if not excluded_log:
        return 0
    ws = _get_or_create_worksheet(spreadsheet, EXCLUDED_SHEET_NAME, EXCLUDED_HEADER)
    existing_records = ws.get_all_records()
    already_logged = {str(r.get("聯單序號", "")).strip() for r in existing_records}

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    rows_to_add = []
    for item in excluded_log:
        ticket = str(item.get("聯單序號", "")).strip()
        if ticket in already_logged:
            continue
        rows_to_add.append([
            item.get("日期", ""),
            ticket,
            item.get("車號", ""),
            item.get("原因", ""),
            item.get("數量", ""),
            now_str,
        ])
    if rows_to_add:
        ws.append_rows(rows_to_add)
    return len(rows_to_add)


def load_processed_tickets(spreadsheet) -> set:
    """從「已處理聯單」分頁讀取所有先前已經處理過（已計入累計）的聯單序號，回傳 set。"""
    ws = _get_or_create_worksheet(spreadsheet, PROCESSED_SHEET_NAME, PROCESSED_HEADER)
    col = ws.col_values(1)[1:]  # 跳過標題列
    return {v.strip() for v in col if v.strip()}


def save_processed_tickets(spreadsheet, all_tickets: set, already_known: set = None):
    """把這批新出現的聯單序號（尚未記錄過的）加進「已處理聯單」分頁，避免重複寫入。"""
    already_known = already_known or set()
    new_tickets = sorted(set(all_tickets) - already_known)
    if not new_tickets:
        return 0
    ws = _get_or_create_worksheet(spreadsheet, PROCESSED_SHEET_NAME, PROCESSED_HEADER)
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ws.append_rows([[t, now_str] for t in new_tickets])
    return len(new_tickets)
