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

TEMPLATE_DAILY = os.path.join(_BASE_DIR, "templates", "daily_record_template.docx")
TEMPLATE_MONTHLY = os.path.join(_BASE_DIR, "templates", "monthly_report_template.docx")
TEMPLATE_TIME = os.path.join(_BASE_DIR, "templates", "time_overview_template.docx")


def _check_templates_exist():
    missing = [p for p in [TEMPLATE_DAILY, TEMPLATE_MONTHLY, TEMPLATE_TIME] if not os.path.isfile(p)]
    if missing:
        templates_dir = os.path.join(_BASE_DIR, "templates")
        existing = os.listdir(templates_dir) if os.path.isdir(templates_dir) else []
        raise FileNotFoundError(
            "找不到樣板檔案：\n" + "\n".join(missing) +
            f"\n\ntemplates/ 資料夾目前實際內容：{existing}\n"
            "請確認 templates 資料夾內的三個 .docx 檔案已經正確 commit/push 到 GitHub，"
            "且檔名與程式碼裡的完全一致（英文檔名，無需擔心中文編碼問題）。"
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
    for i in range(data_row_count):
        row = table.rows[header_rows + i]
        if i < n:
            r = day_df.iloc[i]
            qty = r["數量"]
            daily_total += qty
            _set_cell_text(row.cells[0], str(i + 1))
            _set_cell_text(row.cells[1], str(r["聯單序號"]))
            _set_cell_text(row.cells[2], str(r["出場車號"]))
            _set_cell_text(row.cells[3], f"{qty:g}")
            checked = "✓" if str(r.get("狀態", "")).strip() == "已完成" else ""
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

    cumulative_after = cumulative_before + daily_total
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
    for day in range(1, days_in_month + 1):
        row = table.rows[day]  # row 1 對應 1 日 ... row 31 對應 31 日
        the_date = pd.Timestamp(year=year_month.year, month=year_month.month, day=day).date()
        g = grouped.get(the_date)
        if g is not None and len(g):
            tickets = sorted(g["聯單序號"].astype(str).tolist())
            ticket_range = tickets[0] if len(tickets) == 1 else f"{tickets[0]} ~ {tickets[-1]}"
            count = len(g)
            qty_sum = g["數量"].sum()
            month_total += qty_sum
            _set_cell_text(row.cells[1], ticket_range)
            _set_cell_text(row.cells[2], str(count))
            _set_cell_text(row.cells[3], f"{qty_sum:g}")
            _set_cell_text(row.cells[4], "")
        # 若當天沒有資料，保留原樣（空白）

    cumulative_after = cumulative_before + month_total
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
            _set_cell_text(row.cells[0], str(engineering_name))
            _set_cell_text(row.cells[1], str(r["出場車號"]))
            _set_cell_text(row.cells[2], out_dt.strftime("%Y.%m.%d") if pd.notna(out_dt) else "")
            _set_cell_text(row.cells[3], out_dt.strftime("%H:%M") if pd.notna(out_dt) else "")
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

def build_all_documents(df: pd.DataFrame, engineering_name_override: str = "",
                         monthly_cumulative_start: float = 0.0):
    """回傳 dict: {輸出檔名: docx.Document}，以及一份處理摘要 list[dict]。"""
    _check_templates_exist()
    df = prepare_dataframe(df)
    outputs = {}
    summary = []

    # --- 每日出場紀錄 + 運送時間一覽表：逐日產生 ---
    dates = summarize_dates(df)
    daily_cumulative = {}  # 每月累計，key=年月
    running_cum_by_month = {}

    for d in dates:
        day_df = df[df["日期"] == d].reset_index(drop=True)
        eng_name = engineering_name_override or day_df["工程名稱"].iloc[0]
        period = pd.Period(year=d.year, month=d.month, freq="M")
        cum_before = running_cum_by_month.get(period, monthly_cumulative_start)

        doc_daily, daily_total, cum_after = generate_daily_record(
            day_df, d, eng_name, cum_before
        )
        running_cum_by_month[period] = cum_after

        fname_daily = f"每日出場紀錄_{d.strftime('%Y%m%d')}.docx"
        outputs[fname_daily] = doc_daily

        doc_time = generate_time_overview(day_df, d, eng_name)
        fname_time = f"運送時間一覽表_{d.strftime('%Y%m%d')}.docx"
        outputs[fname_time] = doc_time

        summary.append({
            "日期": d.strftime("%Y-%m-%d"),
            "筆數": len(day_df),
            "當日數量(m3)": daily_total,
            "當月累計(m3)": cum_after,
        })

    # --- 統計月報表：逐月產生（用同一組累計，銜接每日出場紀錄的累計）---
    months = summarize_months(df)
    for m in months:
        month_df = df[df["年月"] == m].reset_index(drop=True)
        eng_name = engineering_name_override or month_df["工程名稱"].iloc[0]
        doc_monthly, month_total, cum_after = generate_monthly_report(
            month_df, m, eng_name, monthly_cumulative_start
        )
        fname_monthly = f"統計月報表_{m.strftime('%Y%m')}.docx"
        outputs[fname_monthly] = doc_monthly

    return outputs, summary


def doc_to_bytes(doc: Document) -> bytes:
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()
