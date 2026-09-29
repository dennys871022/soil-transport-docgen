# -*- coding: utf-8 -*-
"""
Streamlit 主程式：上傳聯單 CSV，自動產生三種 Word 報表並打包下載。
支援：排除異常聯單（不列入清運數量）、重複資料防呆提醒、Google 試算表自動記錄。
"""

import io
import zipfile

import pandas as pd
import streamlit as st

import sheets
from docgen import (
    CONTRACTOR_NAME,
    build_all_documents,
    check_duplicates,
    doc_to_bytes,
    dump_cumulative_state,
    fix_legacy_cumulative,
    load_cumulative_state,
    prepare_dataframe,
    read_csv_any_encoding,
    shorten_ticket,
    summarize_dates,
    summarize_months,
    validate_cumulative_monotonic,
    _check_templates_exist,
)

st.set_page_config(page_title="土石方聯單 Word 自動產生器", page_icon="🚛", layout="wide")

try:
    _check_templates_exist()
except FileNotFoundError as e:
    st.error(str(e))
    st.stop()

st.title("🚛 有價土石方聯單 → Word 報表自動產生器")
st.caption("上傳聯單 CSV，自動產生「每日出場紀錄」「統計月報表」「運送時間一覽表」三種 Word 文件")

sheets_enabled = sheets.is_configured(st)
spreadsheet = None
if sheets_enabled:
    try:
        spreadsheet = sheets.get_spreadsheet(st)
    except Exception as e:  # noqa: BLE001
        sheets_enabled = False
        st.warning(f"Google 試算表連線設定看起來有問題，暫時改用手動記錄檔模式。錯誤訊息：{e}")

with st.sidebar:
    st.header("⚙️ 設定")
    engineering_name_override = st.text_input(
        "工程名稱（留空則自動採用 CSV 內的工程名稱）", value=""
    )
    st.text_input("施工廠商", value=CONTRACTOR_NAME, disabled=True)

    st.markdown("---")
    st.subheader("📌 累計記錄 / 防重複計算")

    cumulative_state = {}
    known_tickets = set()

    if sheets_enabled:
        st.success("✅ 已連接 Google 試算表，累計數字與已處理聯單會自動讀取與寫回")
        try:
            cumulative_state = sheets.load_cumulative_from_sheet(spreadsheet)
            known_tickets = sheets.load_processed_tickets(spreadsheet)
            st.caption(f"目前已記錄 {len(known_tickets)} 張聯單、{len(cumulative_state)} 個月份的累計")
            if cumulative_state:
                st.json(cumulative_state, expanded=False)
        except Exception as e:  # noqa: BLE001
            st.error(f"讀取 Google 試算表資料失敗：{e}")
    else:
        st.caption("尚未設定 Google 試算表連線，使用手動上傳/下載記錄檔的方式（見 README 可改為自動連線）。")
        state_file = st.file_uploader(
            "上傳「上次的記錄檔」(.json)，程式會自動接續累計並防止重複計算",
            type=["json"], key="state_file",
        )
        if state_file is not None:
            try:
                loaded = load_cumulative_state(state_file)
                cumulative_state = loaded["cumulative"]
                known_tickets = set(loaded["processed_tickets"])
                st.success(f"已讀取記錄：{len(cumulative_state)} 個月份的累計、{len(known_tickets)} 張已處理聯單")
                if cumulative_state:
                    st.json(cumulative_state, expanded=False)
            except ValueError as e:
                st.error(str(e))
                st.stop()

    monthly_cumulative_start = st.number_input(
        "若某個月份沒有累計記錄可以沿用，這個月要從多少開始累計？（新工程第一次使用時填這裡）",
        min_value=0.0, value=0.0, step=1.0,
    )

    if cumulative_state:
        with st.expander("🔧 累計數字好像不對？點此一次性修正"):
            st.caption(
                "如果你發現「累計」欄位看起來像是每個月各自獨立的量（例如6月1464、7月7200、8月1596，"
                "而不是持續往上加），可以用這個工具修正成正確的跨月累計。"
                "**這個修正只能執行一次**，資料正確之後請不要重複套用，否則會疊加錯誤。"
            )
            st.write("目前的累計資料：")
            st.json(cumulative_state, expanded=False)
            fixed_preview = fix_legacy_cumulative(cumulative_state)
            st.write("修正後預覽（假設目前數字是各月各自的量，依時間先後累加）：")
            st.json(fixed_preview, expanded=False)
            if st.button("✅ 確認套用這個修正"):
                cumulative_state = fixed_preview
                if sheets_enabled:
                    try:
                        sheets.save_cumulative_to_sheet(spreadsheet, cumulative_state)
                        st.success("已修正並寫回 Google 試算表，重新整理頁面即可看到正確數字。")
                    except Exception as e:  # noqa: BLE001
                        st.error(f"寫回試算表失敗：{e}")
                else:
                    st.success("已在本次畫面套用修正。請記得等一下產生文件後，下載新的記錄檔保存這個修正結果。")
                st.session_state["_cumulative_state_override"] = cumulative_state

    if "_cumulative_state_override" in st.session_state:
        cumulative_state = st.session_state["_cumulative_state_override"]

    st.markdown("---")
    st.markdown(
        "**產生規則**\n"
        "- CSV 中每一個「出場日期」都會各自產生一份「每日出場紀錄」與「運送時間一覽表」\n"
        "- CSV 中每一個「年-月」都會產生一份「統計月報表」\n"
        "- 檢查項目 4 欄：一律打勾（出場當下已通過車輛/駕駛檢查）\n"
        "- 運送數量：優先採用「實際出土量」，缺值則用「載運土方量」\n"
        "- 月報表的憑證序號會自動省略共同前綴，只顯示後4碼「起始 ~ 結束」\n"
        "- 重複上傳過的聯單序號不會被重複計入累計（但文件內容仍正常顯示）\n"
    )

uploaded_file = st.file_uploader("上傳聯單資料 CSV 檔", type=["csv"])

if uploaded_file is None:
    st.info("請上傳 CSV 檔案以開始。CSV 需包含：聯單序號、工程名稱、土資場名稱、載運土方量、狀態、出場/進場車頭車號、出場/進場車斗車號、出場/進場日期、司機姓名、實際出土量、實際進土量 等欄位。")
    st.stop()

try:
    raw_df = read_csv_any_encoding(uploaded_file)
except Exception as e:  # noqa: BLE001
    st.error(f"讀取 CSV 失敗：{e}")
    st.stop()

try:
    df = prepare_dataframe(raw_df)
except Exception as e:  # noqa: BLE001
    st.error(f"CSV 欄位格式有誤：{e}")
    st.stop()

dates = summarize_dates(df)
months = summarize_months(df)

col1, col2, col3 = st.columns(3)
col1.metric("資料筆數", len(df))
col2.metric("涵蓋天數", len(dates))
col3.metric("涵蓋月份數", len(months))

st.subheader("📋 資料預覽")
st.dataframe(
    df[["聯單序號", "工程名稱", "出場車號", "進場車號", "數量", "狀態", "司機姓名", "出場日期時間", "進場日期時間"]],
    use_container_width=True,
    height=280,
)

# ---------------------------------------------------------------------------
# 防呆：重複資料檢查（及時提醒，避免重複計算）
# ---------------------------------------------------------------------------
dup_info = check_duplicates(df, known_tickets)
force_recount = False
if dup_info["tickets"]:
    st.error(
        f"⚠️ 偵測到 {len(dup_info['tickets'])} 張聯單先前已經處理過（重複上傳），"
        f"共 {dup_info['duplicate_qty']:g} 立方公尺。這批重複的資料**不會**被重複計入累計數字，"
        "但文件內容仍會正常顯示（真實發生過的紀錄）。"
    )
    with st.expander("查看重複的聯單序號 / 日期"):
        st.write("重複日期：", "、".join(dup_info["dates"]) or "無")
        st.write("重複聯單序號：")
        st.code("\n".join(dup_info["tickets"]))
    force_recount = st.checkbox(
        "我確認這些不是重複資料（例如系統誤判），這次要強制正常計入累計 —— 請謹慎勾選",
        value=False,
    )

effective_known_tickets = set() if force_recount else known_tickets

# ---------------------------------------------------------------------------
# 排除異常聯單（例如退車、拒收）
# ---------------------------------------------------------------------------
st.subheader("🚫 排除異常聯單（退車 / 拒收，不列入清運數量）")
st.caption(
    "選擇這次要排除的聯單。每日出場紀錄仍會正常顯示這一列（因為車輛確實有出場）；"
    "但運送時間一覽表只會留出場時間，入場時間與數量留空，並在土資場名稱欄改填你輸入的原因；"
    "統計月報表會扣掉這台的數量，並在備註欄註明原因與聯單號碼；所有總計/累計也都不會計入。"
)

ticket_options = df["聯單序號"].astype(str).tolist()
ticket_labels = {}
for t, (_, row) in zip(ticket_options, df.iterrows()):
    out_time = row['出場日期時間'].strftime('%H:%M') if pd.notna(row['出場日期時間']) else ''
    ticket_labels[t] = f"{shorten_ticket(t)}｜{row['出場車號']}｜{row['日期']} {out_time}"

selected_tickets = st.multiselect(
    "選擇要排除的聯單",
    options=ticket_options,
    format_func=lambda t: ticket_labels.get(t, t),
)

exclude_map = {}
if selected_tickets:
    st.markdown("**請填寫每一筆的排除原因：**")
    for t in selected_tickets:
        reason = st.text_input(
            f"排除原因 — {ticket_labels.get(t, t)}",
            key=f"reason_{t}",
            placeholder="例如：液壓故障，土資場拒收",
        )
        exclude_map[t] = reason.strip()

st.subheader("📅 將產生的檔案")
st.write(f"每日出場紀錄 × {len(dates)}、運送時間一覽表 × {len(dates)}、統計月報表 × {len(months)}")
st.write("日期：", "、".join(d.strftime("%Y-%m-%d") for d in dates))
st.write("月份：", "、".join(m.strftime("%Y-%m") for m in months))
if exclude_map:
    st.write(f"本次排除 {len(exclude_map)} 筆聯單，不計入清運數量。")

if st.button("🚀 產生 Word 文件", type="primary"):
    with st.spinner("處理中..."):
        try:
            outputs, summary, updated_state, excluded_log, updated_known_tickets, batch_dup_info = build_all_documents(
                raw_df,
                engineering_name_override=engineering_name_override.strip(),
                monthly_cumulative_start=monthly_cumulative_start,
                cumulative_state=cumulative_state,
                exclude_map=exclude_map,
                known_tickets=effective_known_tickets,
            )
        except Exception as e:  # noqa: BLE001
            st.error(f"產生文件時發生錯誤：{e}")
            st.stop()

    st.success(f"完成！共產生 {len(outputs)} 份 Word 文件。")

    if summary:
        st.subheader("📊 每日彙總")
        st.dataframe(pd.DataFrame(summary), use_container_width=True)

    if excluded_log:
        st.subheader("🚫 本次排除的異常聯單")
        st.dataframe(pd.DataFrame(excluded_log), use_container_width=True)

    if batch_dup_info["tickets"] and not force_recount:
        st.info(f"本次已自動避開 {len(batch_dup_info['tickets'])} 張重複聯單，未重複計入累計。")

    st.subheader("📌 累計記錄")
    st.json(updated_state, expanded=False)

    problems = validate_cumulative_monotonic(updated_state)
    if problems:
        st.error(
            "⚠️ 偵測到累計數字異常（不應該發生，請截圖回報）：\n\n" + "\n".join(f"- {p}" for p in problems) +
            "\n\n為了安全起見，這次**不會**自動寫回 Google 試算表或視為可信的累計記錄，"
            "請先確認資料正確性。文件仍然可以下載。"
        )
    elif sheets_enabled:
        try:
            sheets.save_cumulative_to_sheet(spreadsheet, updated_state)
            logged_count = sheets.log_excluded_tickets(spreadsheet, excluded_log)
            newly_tracked = sheets.save_processed_tickets(spreadsheet, updated_known_tickets, known_tickets)
            msg = f"已自動寫回 Google 試算表「{sheets.CUMULATIVE_SHEET_NAME}」分頁"
            if logged_count:
                msg += f"，新增 {logged_count} 筆異常退車紀錄"
            if newly_tracked:
                msg += f"，新增 {newly_tracked} 張聯單到已處理清單"
            st.success(msg + "。")
        except Exception as e:  # noqa: BLE001
            st.error(f"寫入 Google 試算表失敗，請確認試算表已分享給服務帳號並給予編輯權限。錯誤訊息：{e}")
    else:
        st.download_button(
            "⬇️ 下載本次記錄檔 (JSON，含累計與已處理聯單清單，下次上傳請記得帶上)",
            data=dump_cumulative_state(updated_state, updated_known_tickets),
            file_name="累計記錄.json",
            mime="application/json",
        )

    # 打包成 zip 供一次下載
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for fname, doc in outputs.items():
            zf.writestr(fname, doc_to_bytes(doc))
        if not sheets_enabled:
            zf.writestr("累計記錄.json", dump_cumulative_state(updated_state, updated_known_tickets))
    zip_buffer.seek(0)

    st.download_button(
        "⬇️ 下載全部檔案 (ZIP)",
        data=zip_buffer,
        file_name="土石方報表輸出.zip",
        mime="application/zip",
        type="primary",
    )

    st.subheader("或單獨下載每份 Word 檔案")
    for fname, doc in outputs.items():
        st.download_button(
            f"⬇️ {fname}",
            data=doc_to_bytes(doc),
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key=fname,
        )
