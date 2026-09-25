# -*- coding: utf-8 -*-
"""
Streamlit 主程式：上傳聯單 CSV，自動產生三種 Word 報表並打包下載。
"""

import io
import zipfile

import pandas as pd
import streamlit as st

from docgen import (
    CONTRACTOR_NAME,
    build_all_documents,
    doc_to_bytes,
    prepare_dataframe,
    read_csv_any_encoding,
    summarize_dates,
    summarize_months,
)

st.set_page_config(page_title="土石方聯單 Word 自動產生器", page_icon="🚛", layout="wide")

st.title("🚛 有價土石方聯單 → Word 報表自動產生器")
st.caption("上傳聯單 CSV，自動產生「每日出場紀錄」「統計月報表」「運送時間一覽表」三種 Word 文件")

with st.sidebar:
    st.header("⚙️ 設定")
    engineering_name_override = st.text_input(
        "工程名稱（留空則自動採用 CSV 內的工程名稱）", value=""
    )
    st.text_input("施工廠商", value=CONTRACTOR_NAME, disabled=True)
    monthly_cumulative_start = st.number_input(
        "期初累計立方公尺（若這份 CSV 不是從月初開始，請填入之前已累計的數量）",
        min_value=0.0, value=0.0, step=1.0,
    )
    st.markdown("---")
    st.markdown(
        "**產生規則**\n"
        "- CSV 中每一個「出場日期」都會各自產生一份「每日出場紀錄」與「運送時間一覽表」\n"
        "- CSV 中每一個「年-月」都會產生一份「統計月報表」\n"
        "- 檢查項目 4 欄：狀態＝已完成 → 自動打勾\n"
        "- 運送數量：優先採用「實際出土量」，缺值則用「載運土方量」"
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

st.subheader("📅 將產生的檔案")
st.write(f"每日出場紀錄 × {len(dates)}、運送時間一覽表 × {len(dates)}、統計月報表 × {len(months)}")
st.write("日期：", "、".join(d.strftime("%Y-%m-%d") for d in dates))
st.write("月份：", "、".join(m.strftime("%Y-%m") for m in months))

if st.button("🚀 產生 Word 文件", type="primary"):
    with st.spinner("處理中..."):
        try:
            outputs, summary = build_all_documents(
                raw_df,
                engineering_name_override=engineering_name_override.strip(),
                monthly_cumulative_start=monthly_cumulative_start,
            )
        except Exception as e:  # noqa: BLE001
            st.error(f"產生文件時發生錯誤：{e}")
            st.stop()

    st.success(f"完成！共產生 {len(outputs)} 份 Word 文件。")

    if summary:
        st.subheader("📊 每日彙總")
        st.dataframe(pd.DataFrame(summary), use_container_width=True)

    # 打包成 zip 供一次下載
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for fname, doc in outputs.items():
            zf.writestr(fname, doc_to_bytes(doc))
    zip_buffer.seek(0)

    st.download_button(
        "⬇️ 下載全部檔案 (ZIP)",
        data=zip_buffer,
        file_name="土石方報表輸出.zip",
        mime="application/zip",
        type="primary",
    )

    st.subheader("或單獨下載每份檔案")
    for fname, doc in outputs.items():
        st.download_button(
            f"⬇️ {fname}",
            data=doc_to_bytes(doc),
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key=fname,
        )
