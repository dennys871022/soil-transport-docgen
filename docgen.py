# -*- coding: utf-8 -*-
"""
docgen.py
---------
核心邏輯：讀取聯單 CSV -> 依日期/月份分組 -> 填入三種 Word 樣板

三種輸出文件：
1. 每日出場紀錄 (每個日期一份)
2. 統計月報表   (每個年-月一份)
3. 運送時間一覽表 (每個日期一份)
"""

import copy
import io
from datetime import datetime
from calendar import monthrange

import pandas as pd
from docx import Document

# ---------------------------------------------------------------------------
# 常數設定
# ---------------------------------------------------------------------------

CONTRACTOR_NAME = "力勤工程實業有限公司"

import os

_BASE_DIR = os.path.dirname(os.path.abspath(__file__))

_TEMPLATE_FILENAMES = {
    "daily": "daily_record_template.docx",
    "monthly": "monthly_report_template.docx",
    "time": "time_overview_template.docx",
}


def _resolve_template_path(filename: str) -> str:
    """依序在 templates/ 子資料夾、專案根目錄尋找樣板檔案，回傳第一個存在的路徑。
    若都找不到，回傳 templates/ 底下的路徑（讓後續錯誤訊息可以明確指出預期位置）。"""
    candidates = [
        os.path.join(_BASE_DIR, "templates", filename),
        os.path.join(_BASE_DIR, filename),
    ]
    for path in candidates:
        if os.path.isfile(path):
            return path
    return candidates[0]


TEMPLATE_DAILY = _resolve_template_path(_TEMPLATE_FILENAMES["daily"])
TEMPLATE_MONTHLY = _resolve_template_path(_TEMPLATE_FILENAMES["monthly"])
TEMPLATE_TIME = _resolve_template_path(_TEMPLATE_FILENAMES["time"])


def _check_templates_exist():
    missing = [p for p in [TEMPLATE_DAILY, TEMPLATE_MONTHLY, TEMPLATE_TIME] if not os.path.isfile(p)]
    if missing:
        templates_dir = os.path.join(_BASE_DIR, "templates")
        existing_in_templates = os.listdir(templates_dir) if os.path.isdir(templates_dir) else "（無 templates 資料夾）"
        existing_in_root = [f for f in os.listdir(_BASE_DIR) if f.endswith(".docx")]
        raise FileNotFoundError(
            "找不到樣板檔案：\n" + "\n".join(missing) +
            f"\n\ntemplates/ 資料夾內容：{existing_in_templates}"
            f"\n專案根目錄下的 .docx 檔案：{existing_in_root}\n"
            "請確認三個樣板檔案（daily_record_template.docx、monthly_report_template.docx、"
            "time_overview_template.docx）已經正確上傳到 GitHub repo（放在根目錄或 templates/ 子資料夾皆可）。"
        )

CSV_ENCODING_CANDIDATES = ["utf-8-sig", "utf-8", "cp950", "big5"]


# ---------------------------------------------------------------------------
# CSV 讀取與整理
# ---------------------------------------------------------------------------

def read_csv_any_encoding(file_obj) -> pd.DataFrame:
    """嘗試多種編碼讀取 CSV（file_obj 可以是路徑或檔案物件）。"""
    last_err = None
    for enc in CSV_ENCODING_CANDIDATES:
        try:
            if hasattr(file_obj, "seek"):
                file_obj.seek(0)
            df = pd.read_csv(file_obj, encoding=enc, dtype=str)
            return df
        except Exception as e:  # noqa: BLE001
            last_err = e
    raise ValueError(f"無法讀取 CSV，請確認編碼是否為 UTF-8 / Big5。錯誤：{last_err}")


def prepare_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """整理欄位型別、補上輔助欄位（日期、時間、車號組合、加總數量等）。"""
    df = df.copy()

    required_cols = [
        "聯單序號", "工程名稱", "土資場名稱", "載運土方量", "狀態",
        "出場車頭車號", "出場車斗車號", "出場日期",
        "進場車頭車號", "進場車斗車號", "進場日期",
        "司機姓名", "實際出土量", "實際進土量",
    ]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise ValueError(f"CSV 缺少必要欄位：{', '.join(missing)}")

    # 出場/進場日期時間解析
    df["出場日期時間"] = pd.to_datetime(df["出場日期"], errors="coerce")
    df["進場日期時間"] = pd.to_datetime(df["進場日期"], errors="coerce")

    # 用出場日期當作這一筆記錄所屬的「日期」
    df["日期"] = df["出場日期時間"].dt.date
    df["年月"] = df["出場日期時間"].dt.to_period("M")

    # 數量：優先用實際出土量，缺值則用載運土方量
    def pick_qty(row):
        for col in ["實際出土量", "載運土方量"]:
            v = row.get(col)
            if pd.notna(v) and str(v).strip() != "":
                try:
                    return float(v)
                except ValueError:
                    continue
        return 0.0

    df["數量"] = df.apply(pick_qty, axis=1)

    # 車號組合：車頭/車斗
    def combine_plate(head, tail):
        head = "" if pd.isna(head) else str(head).strip()
        tail = "" if pd.isna(tail) else str(tail).strip()
        if head and tail:
            return f"{head}/{tail}"
        return head or tail

    df["出場車號"] = df.apply(lambda r: combine_plate(r["出場車頭車號"], r["出場車斗車號"]), axis=1)
    df["進場車號"] = df.apply(lambda r: combine_plate(r["進場車頭車號"], r["進場車斗車號"]), axis=1)

    # 依出場時間排序，車次才會照時間先後編號
    df = df.sort_values("出場日期時間").reset_index(drop=True)

    return df


def apply_exclusions(df: pd.DataFrame, exclude_map: dict = None) -> pd.DataFrame:
    """依聯單序號標記排除（異常退車等不列入清運數量的紀錄）。
    exclude_map: {原始完整聯單序號: 排除原因文字}
    加上兩個欄位：排除 (bool)、排除原因 (str)。
    """
    df = df.copy()
    exclude_map = exclude_map or {}
    df["排除"] = df["聯單序號"].astype(str).isin(exclude_map.keys())
    df["排除原因"] = df["聯單序號"].astype(str).map(exclude_map).fillna("")
    return df


def apply_duplicate_flags(df: pd.DataFrame, known_tickets: set = None) -> pd.DataFrame:
    """依「先前已處理過的聯單序號」標記這批資料裡的重複項目。
    加上一個欄位：重複 (bool)。重複的聯單不會被重複計入累計，但文件內容仍正常顯示
    （因為那是真實發生過的紀錄，只是這次上傳的資料跟之前重疊了）。
    """
    df = df.copy()
    known_tickets = known_tickets or set()
    df["重複"] = df["聯單序號"].astype(str).isin(known_tickets)
    return df


def check_duplicates(df: pd.DataFrame, known_tickets: set = None) -> dict:
    """上傳 CSV 後、產生文件前先呼叫，檢查有沒有跟之前已處理過的資料重疊。
    回傳 {"tickets": [...重複的聯單序號(縮短顯示)...], "dates": [...重複的日期...],
          "duplicate_qty": 重複部分原本會被重複計算的立方公尺數}
    """
    known_tickets = known_tickets or set()
    df = df.copy()
    df["_dup"] = df["聯單序號"].astype(str).isin(known_tickets)
    dup_df = df[df["_dup"]]
    if dup_df.empty:
        return {"tickets": [], "dates": [], "duplicate_qty": 0.0}
    dup_tickets = sorted(shorten_ticket(t) for t in dup_df["聯單序號"].astype(str).unique())
    dup_dates = sorted(str(d) for d in dup_df["日期"].dropna().unique()) if "日期" in dup_df.columns else []
    dup_qty = dup_df["數量"].sum() if "數量" in dup_df.columns else 0.0
    return {"tickets": dup_tickets, "dates": dup_dates, "duplicate_qty": float(dup_qty)}


def shorten_ticket(ticket: str) -> str:
    """把單一聯單序號的流水號部分縮短成4位數。
    例如 'EYG10099EYG21699_B2-3_00000001' -> 'EYG10099EYG21699_B2-3_0001'
    若最後一段不是純數字，就原樣返回不做縮短。
    """
    ticket = str(ticket)
    if "_" not in ticket:
        return ticket
    prefix, suf = ticket.rsplit("_", 1)
    if suf.isdigit():
        return f"{prefix}_{int(suf):04d}"
    return ticket


def format_ticket_range(tickets):
    """把一組聯單序號簡化成『共同前綴_起始 ~ 結束』的格式，且流水號只保留4位數。
    例如 ['EYG10099EYG21699_B2-3_00000012', 'EYG10099EYG21699_B2-3_00000041']
    會變成 'EYG10099EYG21699_B2-3_0012 ~ 0041'。
    若只有一張，直接回傳縮短後的該序號；若前綴不同（理論上同一天不會發生），
    退回顯示完整兩個序號（各自縮短）。
    """
    if not tickets:
        return ""
    if len(tickets) == 1:
        return shorten_ticket(tickets[0])
    first, last = tickets[0], tickets[-1]
    if "_" in first and "_" in last:
        prefix1, suf1 = first.rsplit("_", 1)
        prefix2, suf2 = last.rsplit("_", 1)
        if prefix1 == prefix2 and suf1.isdigit() and suf2.isdigit():
            return f"{prefix1}_{int(suf1):04d} ~ {int(suf2):04d}"
    return f"{shorten_ticket(first)} ~ {shorten_ticket(last)}"


# ---------------------------------------------------------------------------
# 累計狀態（跨月/跨次上傳的期初累計記憶）
# ---------------------------------------------------------------------------

def load_cumulative_state(file_obj) -> dict:
    """讀取先前下載的累計記錄 JSON 檔。
    回傳 {"cumulative": {"YYYY-MM": 累計立方公尺,...}, "processed_tickets": [...已處理過的聯單序號...]}。
    相容舊版格式（舊版檔案內容是單純的 {"YYYY-MM": 數字} 沒有 processed_tickets）。
    """
    import json
    if file_obj is None:
        return {"cumulative": {}, "processed_tickets": []}
    try:
        if hasattr(file_obj, "seek"):
            file_obj.seek(0)
        raw = file_obj.read()
        if isinstance(raw, bytes):
            raw = raw.decode("utf-8-sig")
        data = json.loads(raw)
        if "cumulative" in data or "processed_tickets" in data:
            cumulative = {str(k): float(v) for k, v in (data.get("cumulative") or {}).items()}
            processed = list(data.get("processed_tickets") or [])
        else:
            # 舊版格式：整份檔案就是 {"YYYY-MM": 數字}
            cumulative = {str(k): float(v) for k, v in data.items()}
            processed = []
        return {"cumulative": cumulative, "processed_tickets": processed}
    except Exception as e:  # noqa: BLE001
        raise ValueError(f"累計記錄檔格式錯誤，請確認是上次程式產生的 JSON 檔：{e}")


def dump_cumulative_state(cumulative: dict, processed_tickets=None) -> str:
    import json
    payload = {
        "cumulative": cumulative,
        "processed_tickets": sorted(set(processed_tickets or [])),
    }
    return json.dumps(payload, ensure_ascii=False, indent=2)


def summarize_dates(df: pd.DataFrame):
    """回傳資料中出現的所有日期（datetime.date），由小到大排序。"""
    return sorted(d for d in df["日期"].dropna().unique())


def summarize_months(df: pd.DataFrame):
    """回傳資料中出現的所有年月（pandas Period），由小到大排序。"""
    return sorted(m for m in df["年月"].dropna().unique())


# ---------------------------------------------------------------------------
# Word 表格輔助函式
# ---------------------------------------------------------------------------

def _clone_row(table, src_row_index: int, insert_before_index: int):
    """複製 table 中 src_row_index 那一列的 XML，插入到 insert_before_index 之前。
    回傳新插入的 row 物件。"""
    src_tr = table.rows[src_row_index]._tr
    new_tr = copy.deepcopy(src_tr)
    ref_tr = table.rows[insert_before_index]._tr
    ref_tr.addprevious(new_tr)
    # 重新取得 row 物件（python-docx 會依 XML 順序重建 rows 集合）
    return table.rows[insert_before_index - 1]


def _append_cloned_row(table, src_row_index: int):
    """複製 table 中 src_row_index 那一列，附加到表格最後面。回傳新列物件。"""
    src_tr = table.rows[src_row_index]._tr
    new_tr = copy.deepcopy(src_tr)
    table._tbl.append(new_tr)
    return table.rows[len(table.rows) - 1]


def _clear_row_text(row):
    for cell in row.cells:
        for p in cell.paragraphs:
            for run in p.runs:
                run.text = ""
            if not p.runs and p.text:
                p.text = ""


def _set_cell_text(cell, text):
    """保留原本第一個 run 的格式，把文字換成指定內容；清掉多餘 run。"""
    if not cell.paragraphs:
        cell.text = text
        return
    p = cell.paragraphs[0]
    if p.runs:
        p.runs[0].text = text
        for extra in p.runs[1:]:
            extra.text = ""
    else:
        p.text = text
    # 其餘段落清空
    for extra_p in cell.paragraphs[1:]:
        for run in extra_p.runs:
            run.text = ""


def _replace_in_paragraphs(doc: Document, replacements: dict):
    """在整份文件的段落文字中做關鍵字取代（保留原有格式，直接操作 run）。"""
    for p in doc.paragraphs:
        for key, val in replacements.items():
            if key in p.text:
                _replace_in_paragraph(p, key, val)


def _replace_in_paragraph(paragraph, key, val):
    """處理文字可能被拆成多個 run 的情況：把整段文字合併後重寫回第一個 run。"""
    full_text = "".join(run.text for run in paragraph.runs)
    if key not in full_text:
        return
    new_text = full_text.replace(key, val)
    if paragraph.runs:
        paragraph.runs[0].text = new_text
        for run in paragraph.runs[1:]:
            run.text = ""
    else:
        paragraph.text = new_text


# ---------------------------------------------------------------------------
# 1. 每日出場紀錄
# ---------------------------------------------------------------------------

def generate_daily_record(day_df: pd.DataFrame, date, engineering_name: str,
                           cumulative_before: float) -> Document:
    doc = Document(TEMPLATE_DAILY)

    engineering_name = engineering_name or (day_df["工程名稱"].iloc[0] if len(day_df) else "")

    for p in doc.paragraphs:
        if p.text.startswith("工程名稱："):
            _replace_in_paragraph(p, p.text, f"工程名稱：{engineering_name}")
        if p.text.startswith("施工廠商："):
            _replace_in_paragraph(p, p.text, f"施工廠商：{CONTRACTOR_NAME}")
        if p.text.startswith("出埸日期："):
            _replace_in_paragraph(
                p, p.text,
                f"出埸日期：{date.year - 1911:>3}年{date.month:>2}月{date.day:>2}日"
            )

    table = doc.tables[0]
    header_rows = 2
    total_row_index = len(table.rows) - 1  # 目前總計列的位置
    data_row_count = total_row_index - header_rows  # 目前模板內的資料列數（預設 5）

    n = len(day_df)
    # 資料列數不夠 -> 在總計列之前插入新列（複製最後一列資料列的格式）
    while data_row_count < n:
        _clone_row(table, total_row_index - 1, total_row_index)
        total_row_index += 1
        data_row_count += 1

    daily_total = 0.0
    cumulative_delta = 0.0
    for i in range(data_row_count):
        row = table.rows[header_rows + i]
        if i < n:
            r = day_df.iloc[i]
            qty = r["數量"]
            if not r.get("排除", False):
                daily_total += qty
                if not r.get("重複", False):
                    cumulative_delta += qty
            _set_cell_text(row.cells[0], str(i + 1))
            _set_cell_text(row.cells[1], shorten_ticket(r["聯單序號"]))
            _set_cell_text(row.cells[2], str(r["出場車號"]))
            _set_cell_text(row.cells[3], f"{qty:g}")
            # 4項檢查是出場當下的車輛/駕駛檢查，只要有出場紀錄就代表已通過檢查，
            # 跟後續是否異常退車（清運數量認定）無關，一律打勾
            checked = "✓"
            _set_cell_text(row.cells[4], checked)
            _set_cell_text(row.cells[5], checked)
            _set_cell_text(row.cells[6], checked)
            _set_cell_text(row.cells[7], checked)
            out_time = r["出場日期時間"]
            _set_cell_text(row.cells[8], out_time.strftime("%H:%M") if pd.notna(out_time) else "")
            _set_cell_text(row.cells[9], str(r.get("司機姓名", "")))
        else:
            # 多出來的空白列：保留車次編號，其餘留空
            _set_cell_text(row.cells[0], str(i + 1))
            for c in range(1, 10):
                _set_cell_text(row.cells[c], "")

    cumulative_after = cumulative_before + cumulative_delta
    total_row = table.rows[total_row_index]
    total_text = (
        f"有價土石方運送數量    ：{daily_total:g}   立方公尺\n"
        f"累計有價土石方運送數量：{cumulative_after:g}   立方公尺"
    )
    for c in range(3, 10):
        _set_cell_text(total_row.cells[c], total_text)

    return doc, daily_total, cumulative_after


# ---------------------------------------------------------------------------
# 2. 統計月報表
# ---------------------------------------------------------------------------

def generate_monthly_report(month_df: pd.DataFrame, year_month, engineering_name: str,
                             cumulative_before: float) -> Document:
    doc = Document(TEMPLATE_MONTHLY)

    engineering_name = engineering_name or (month_df["工程名稱"].iloc[0] if len(month_df) else "")
    roc_year = year_month.year - 1911

    for p in doc.paragraphs:
        if "年" in p.text and "月" in p.text and p.text.strip().startswith("（"):
            _replace_in_paragraph(p, p.text, f"（  {roc_year}  年  {year_month.month}  月）")
        if p.text.startswith("工程名稱："):
            _replace_in_paragraph(p, "工程名稱：", f"工程名稱：{engineering_name}")
        if p.text.startswith("施工廠商："):
            _replace_in_paragraph(p, p.text, f"施工廠商：{CONTRACTOR_NAME}")

    table = doc.tables[0]
    days_in_month = monthrange(year_month.year, year_month.month)[1]

    grouped = {d: g for d, g in month_df.groupby(month_df["日期"])}

    month_total = 0.0
    month_cumulative_delta = 0.0
    for day in range(1, days_in_month + 1):
        row = table.rows[day]  # row 1 對應 1 日 ... row 31 對應 31 日
        the_date = pd.Timestamp(year=year_month.year, month=year_month.month, day=day).date()
        g = grouped.get(the_date)
        if g is not None and len(g):
            tickets = sorted(g["聯單序號"].astype(str).tolist())
            ticket_range = format_ticket_range(tickets)
            count = len(g)
            included = g[~g["排除"]]
            excluded = g[g["排除"]]
            qty_sum = included["數量"].sum()
            month_total += qty_sum
            month_cumulative_delta += included[~included["重複"]]["數量"].sum()
            _set_cell_text(row.cells[1], ticket_range)
            _set_cell_text(row.cells[2], str(count))
            _set_cell_text(row.cells[3], f"{qty_sum:g}")
            if len(excluded):
                notes = []
                for _, er in excluded.iterrows():
                    notes.append(
                        f"{shorten_ticket(er['聯單序號'])} 異常不列入計量"
                        f"（{er['排除原因'] or '原因未填'}）"
                    )
                _set_cell_text(row.cells[4], "；".join(notes))
            else:
                _set_cell_text(row.cells[4], "")
        # 若當天沒有資料，保留原樣（空白）

    cumulative_after = cumulative_before + month_cumulative_delta
    total_row = table.rows[len(table.rows) - 1]  # 範本固定 33 列(含 1~31 日)，總計永遠在最後一列
    total_text = (
        f"有價土石方運送數量    ：{month_total:g}   立方公尺\n"
        f"累計有價土石方運送數量：{cumulative_after:g}   立方公尺"
    )
    for c in range(1, 5):
        _set_cell_text(total_row.cells[c], total_text)

    return doc, month_total, cumulative_after


# ---------------------------------------------------------------------------
# 3. 運送時間一覽表
# ---------------------------------------------------------------------------

def generate_time_overview(day_df: pd.DataFrame, date, engineering_name: str) -> Document:
    doc = Document(TEMPLATE_TIME)

    engineering_name = engineering_name or (day_df["工程名稱"].iloc[0] if len(day_df) else "")
    for p in doc.paragraphs:
        if p.text.startswith("工程名稱："):
            _replace_in_paragraph(p, "工程名稱：", f"工程名稱：{engineering_name}")

    table = doc.tables[0]
    header_rows = 3  # row0 大標題, row1 群組標題, row2 欄位標題
    data_start = header_rows  # index 3 起為資料列（範本第一列是範例資料）
    total_rows = len(table.rows)
    data_row_count = total_rows - header_rows

    n = len(day_df)
    while data_row_count < n:
        _append_cloned_row(table, total_rows - 1)
        total_rows += 1
        data_row_count += 1

    for i in range(data_row_count):
        row = table.rows[data_start + i]
        if i < n:
            r = day_df.iloc[i]
            out_dt = r["出場日期時間"]
            in_dt = r["進場日期時間"]
            is_excluded = bool(r.get("排除", False))
            _set_cell_text(row.cells[0], str(engineering_name))
            _set_cell_text(row.cells[1], str(r["出場車號"]))
            _set_cell_text(row.cells[2], out_dt.strftime("%Y.%m.%d") if pd.notna(out_dt) else "")
            _set_cell_text(row.cells[3], out_dt.strftime("%H:%M") if pd.notna(out_dt) else "")
            if is_excluded:
                # 異常退車：沒有實際入場，入場日期/時間與載運數量留空，
                # 土資場名稱欄位改填使用者輸入的排除原因
                _set_cell_text(row.cells[4], "")
                _set_cell_text(row.cells[5], "")
                _set_cell_text(row.cells[6], "")
                reason = r.get("排除原因", "") or "異常退車"
                _set_cell_text(row.cells[7], f"異常退車：{reason}")
            else:
                _set_cell_text(row.cells[4], in_dt.strftime("%Y.%m.%d") if pd.notna(in_dt) else "")
                _set_cell_text(row.cells[5], in_dt.strftime("%H:%M") if pd.notna(in_dt) else "")
                _set_cell_text(row.cells[6], f"{r['數量']:g}")
                _set_cell_text(row.cells[7], str(r.get("土資場名稱", "")))
        else:
            for c in range(len(row.cells)):
                _set_cell_text(row.cells[c], "")

    return doc


# ---------------------------------------------------------------------------
# 高階流程：一次處理整份 CSV，回傳 {檔名: Document}
# ---------------------------------------------------------------------------

def fix_legacy_cumulative(state: dict) -> dict:
    """一次性修正工具：把舊版「每月各自累計歸零」的錯誤資料，轉換成正確的「跨月持續累加」格式。

    舊版錯誤：state 裡每個月的數字其實是「那個月自己的運送量」（例如 6月1464、7月7200、8月1596）。
    正確格式：state 裡每個月的數字應該是「累計到那個月月底為止的總量」
             （6月1464、7月1464+7200=8664、8月8664+1596=10260）。

    做法：把所有月份依時間先後排序，逐月往上累加，產生新的正確版本。
    只應該手動觸發執行「一次」，修正完之後往後的資料就會是正確格式，不能重複套用
    （重複套用會把已經正確的累計值再錯誤地疊加一次）。
    """
    if not state:
        return {}
    ordered_keys = sorted(state.keys())
    fixed = {}
    running = 0.0
    for k in ordered_keys:
        running += float(state[k])
        fixed[k] = running
    return fixed


def build_all_documents(df: pd.DataFrame, engineering_name_override: str = "",
                         monthly_cumulative_start: float = 0.0,
                         cumulative_state: dict = None,
                         exclude_map: dict = None,
                         known_tickets: set = None):
    """回傳 (outputs, summary, updated_cumulative_state, excluded_log, updated_known_tickets, duplicate_info)。

    cumulative_state: {"YYYY-MM": 累計到該月月底為止的立方公尺總量}。
      這是「整個工程從開始到現在」的累計，跨月會持續往上加、不會每月重新歸零
      （例如6月底累計1464，7月運了7200，7月底累計就是1464+7200=8664，不是7200）。
      取用時一律用目前所有月份中「最大值」當作最新的累計基準，之後往下疊加這批資料的新量。
      若完全沒有記錄，就用 monthly_cumulative_start 當作起點。

    exclude_map: {完整聯單序號: 排除原因}，這些聯單會：
      - 每日出場紀錄：正常顯示整列，但不計入當日/累計總量
      - 運送時間一覽表：只留出場日期/時間，入場與數量留空，土資場名稱改填排除原因
      - 統計月報表：不計入當日數量加總，並在備註欄註明原因與聯單號碼

    known_tickets: 先前已經處理過、已計入累計的聯單序號集合（防止重複計算）。
      這批資料裡如果有聯單序號出現在 known_tickets 裡，文件內容依然正常顯示
      （因為是真實發生過的紀錄），但不會被重複加進累計數字。

    excluded_log: list[dict]，這次的異常排除紀錄，供寫入 Google 試算表存查。
    updated_known_tickets: 這批處理完之後，合併進所有聯單序號的新集合，之後要記得存起來。
    duplicate_info: {"tickets":[...], "dates":[...], "duplicate_qty": 數字}，這批資料裡有多少
      是先前已經處理過的重複資料（供介面顯示提醒用）。
    """
    _check_templates_exist()
    df = prepare_dataframe(df)
    known_tickets = set(known_tickets or set())
    duplicate_info = check_duplicates(df, known_tickets)

    df = apply_exclusions(df, exclude_map)
    df = apply_duplicate_flags(df, known_tickets)
    outputs = {}
    summary = []
    cumulative_state = dict(cumulative_state or {})

    excluded_log = []
    for _, r in df[df["排除"]].iterrows():
        excluded_log.append({
            "日期": r["日期"].strftime("%Y-%m-%d") if pd.notna(r["日期"]) else "",
            "聯單序號": shorten_ticket(r["聯單序號"]),
            "車號": r.get("出場車號", ""),
            "原因": r.get("排除原因", ""),
            "數量": r.get("數量", 0.0),
        })

    # 整個工程「目前為止」的累計基準：所有已知月份中的最大值（因為是持續往上加的總量，
    # 最大值必然是最新的進度；若完全沒有記錄，就用使用者填的期初累計）
    running_total = max(cumulative_state.values()) if cumulative_state else monthly_cumulative_start

    # --- 每日出場紀錄 + 運送時間一覽表：逐日產生（日期已由 summarize_dates 由小到大排序）---
    dates = summarize_dates(df)
    cumulative_after_month = {}  # 記錄這批資料處理完，每個月「最後」的累計值，供月報表使用

    for d in dates:
        day_df = df[df["日期"] == d].reset_index(drop=True)
        eng_name = engineering_name_override or day_df["工程名稱"].iloc[0]
        period = pd.Period(year=d.year, month=d.month, freq="M")

        doc_daily, daily_total, cum_after = generate_daily_record(
            day_df, d, eng_name, running_total
        )
        running_total = cum_after
        cumulative_after_month[period] = running_total

        fname_daily = f"每日出場紀錄_{d.strftime('%Y%m%d')}.docx"
        outputs[fname_daily] = doc_daily

        doc_time = generate_time_overview(day_df, d, eng_name)
        fname_time = f"運送時間一覽表_{d.strftime('%Y%m%d')}.docx"
        outputs[fname_time] = doc_time

        summary.append({
            "日期": d.strftime("%Y-%m-%d"),
            "筆數": len(day_df),
            "當日數量(m3)": daily_total,
            "累計(m3)": cum_after,
        })
        cumulative_state[str(period)] = running_total

    # --- 統計月報表：逐月產生 ---
    months = summarize_months(df)
    for m in months:
        month_df = df[df["年月"] == m].reset_index(drop=True)
        eng_name = engineering_name_override or month_df["工程名稱"].iloc[0]
        month_cumulative_delta = month_df[(~month_df["排除"]) & (~month_df["重複"])]["數量"].sum()
        cum_end_of_month = cumulative_after_month.get(m, running_total)
        cum_before_month = cum_end_of_month - month_cumulative_delta

        doc_monthly, month_total, cum_after = generate_monthly_report(
            month_df, m, eng_name, cum_before_month
        )
        fname_monthly = f"統計月報表_{m.strftime('%Y%m')}.docx"
        outputs[fname_monthly] = doc_monthly

    updated_known_tickets = known_tickets | set(df["聯單序號"].astype(str).tolist())

    return outputs, summary, cumulative_state, excluded_log, updated_known_tickets, duplicate_info


def doc_to_bytes(doc: Document) -> bytes:
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()
