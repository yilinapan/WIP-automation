# -*- coding: utf-8 -*-
"""
WIP 轉貼紙 Streamlit 應用
上傳 WIP Excel → 轉換 → 下載貼紙格式 Excel
"""

import streamlit as st
from datetime import date

from processor import process_uploaded_excel, parse_output_filename_from_input, verify_output

st.set_page_config(page_title="WIP 轉貼紙", page_icon="📋", layout="wide")

# 讓 multiselect 標籤不被截斷
st.markdown("""
<style>
    .stMultiSelect [data-baseweb="tag"] {
        max-width: none !important;
        white-space: nowrap !important;
    }
</style>
""", unsafe_allow_html=True)

st.title("📋 WIP 轉貼紙")
st.caption("上傳 WIP 格式的 Excel 檔案，轉換為貼紙格式後下載")

if "output_bytes" not in st.session_state:
    st.session_state.output_bytes = None
if "output_df" not in st.session_state:
    st.session_state.output_df = None
if "output_filename" not in st.session_state:
    st.session_state.output_filename = ""
if "verify_results" not in st.session_state:
    st.session_state.verify_results = None

uploaded = st.file_uploader("上傳 Excel 檔案", type=["xlsx", "xls"], help="檔名通常包含「WIP」")

if not uploaded:
    st.session_state.output_bytes = None
    st.session_state.output_df = None
    st.session_state.output_filename = ""
    st.session_state.verify_results = None

if uploaded:
    xl_sheets = None
    sheet_info = []
    try:
        import pandas as pd
        from io import BytesIO
        from openpyxl import load_workbook

        file_bytes = uploaded.getvalue()
        xl = pd.ExcelFile(BytesIO(file_bytes))
        xl_sheets = xl.sheet_names

        # 偵測隱藏工作表
        wb_check = load_workbook(BytesIO(file_bytes), read_only=True)
        for ws in wb_check.worksheets:
            state = ws.sheet_state  # 'visible', 'hidden', 'veryHidden'
            sheet_info.append((ws.title, state))
        wb_check.close()
    except Exception as e:
        st.error(f"無法讀取檔案：{e}")

    if xl_sheets:
        visible = [name for name, state in sheet_info if state == "visible"]
        hidden = [name for name, state in sheet_info if state != "visible"]

        st.subheader("選擇要處理的工作表")
        caption = f"偵測到 {len(xl_sheets)} 個工作表"
        if hidden:
            caption += f"（其中 {len(hidden)} 個為隱藏工作表：{', '.join(hidden)}）"
        st.caption(caption)

        selected = st.multiselect(
            "若未選擇則預設處理全部可見工作表",
            options=xl_sheets,
            default=visible if visible else xl_sheets,
        )

        default_name = parse_output_filename_from_input(uploaded.name)
        st.subheader("輸出檔名")
        output_name = st.text_input(
            "可自訂輸出檔名（不含副檔名會自動加 .xlsx）",
            value=default_name,
            placeholder=default_name,
        )
        if output_name and not output_name.endswith(".xlsx"):
            output_name = output_name + ".xlsx"

        with st.expander("📷 貼紙圖片", expanded=False):
            st.caption("轉出的 Excel 會在資料下方加入吊牌貼紙與單件袋貼紙範例。預設使用專案內的圖片，也可上傳 .png 或 .jpg（JPEG 會自動轉為 PNG）。")
            col_ht, col_bag = st.columns(2)
            with col_ht:
                hangtag_upload = st.file_uploader(
                    "吊牌貼紙（選填）",
                    type=["png", "jpg", "jpeg"],
                    key="hangtag",
                    help="未上傳則使用專案內的 吊牌貼紙.png 或 吊牌貼紙.jpg",
                )
            with col_bag:
                bag_upload = st.file_uploader(
                    "單件袋貼紙（選填）",
                    type=["png", "jpg", "jpeg"],
                    key="bag",
                    help="未上傳則使用專案內的 單件袋貼紙.png 或 單件袋貼紙.jpg",
                )

        hangtag_bytes = hangtag_upload.getvalue() if hangtag_upload else None
        bag_bytes = bag_upload.getvalue() if bag_upload else None

        if st.button("🔄 轉換成貼紙", type="primary"):
            sheets_to_use = selected if selected else (visible if visible else xl_sheets)
            with st.spinner("資料處理中，請稍後"):
                try:
                    excel_bytes, df = process_uploaded_excel(
                        uploaded.getvalue(),
                        sheet_names=sheets_to_use,
                        hangtag_image=hangtag_bytes,
                        bag_image=bag_bytes,
                    )
                    st.session_state.output_bytes = excel_bytes
                    st.session_state.output_df = df
                    st.session_state.output_filename = output_name or default_name
                    st.session_state.verify_results = verify_output(
                        df, uploaded.getvalue(), sheet_names=sheets_to_use,
                    )
                except Exception as e:
                    st.error(f"處理失敗：{e}")
                    st.exception(e)
                    st.session_state.output_bytes = None
                    st.session_state.output_df = None
                    st.session_state.verify_results = None

if st.session_state.output_bytes and st.session_state.output_df is not None:
    st.success("✅ 處理完成，請下載檔案")
    fn = st.session_state.output_filename
    st.download_button(
        "📥 下載檔案",
        data=st.session_state.output_bytes,
        file_name=fn,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    # 驗證結果
    if st.session_state.verify_results:
        st.subheader("驗證結果")
        vr = st.session_state.verify_results
        passes = [r for r in vr if r["type"] == "pass"]
        warns = [r for r in vr if r["type"] == "warn"]
        errors = [r for r in vr if r["type"] == "error"]

        if not errors and not warns:
            st.success(f"全部通過（{len(passes)} 項檢查皆正常）")
        else:
            if errors:
                st.error(f"發現 {len(errors)} 個錯誤")
            if warns:
                st.warning(f"發現 {len(warns)} 個警告")

        categories = ["結構", "格式", "遺漏款式", "多出款式", "數量比對", "欄位比對", "數量", "比對"]
        seen_cats = []
        for cat in categories:
            items = [r for r in vr if r["category"] == cat]
            if items:
                seen_cats.append(cat)
        for cat in vr:
            if cat["category"] not in seen_cats:
                seen_cats.append(cat["category"])

        for cat in seen_cats:
            items = [r for r in vr if r["category"] == cat]
            if not items:
                continue
            cat_passes = [r for r in items if r["type"] == "pass"]
            cat_issues = [r for r in items if r["type"] != "pass"]
            if cat_passes and not cat_issues:
                st.markdown(f"✅ **{cat}**：{cat_passes[0]['message']}")
            else:
                with st.expander(
                    f"{'❌' if any(r['type'] == 'error' for r in cat_issues) else '⚠️'} "
                    f"**{cat}**（{len(cat_issues)} 個問題）",
                    expanded=True,
                ):
                    for r in cat_issues:
                        icon = "❌" if r["type"] == "error" else "⚠️"
                        st.markdown(f"{icon} {r['message']}")

    st.subheader("預覽")
    st.dataframe(st.session_state.output_df, use_container_width=True, height=400)
