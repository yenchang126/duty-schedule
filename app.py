"""
勤務分配表自動填入系統
Streamlit 網頁介面
"""

import streamlit as st
import re
from processor import get_available_dates, process_files

# 頁面設定
st.set_page_config(
    page_title="勤務分配表產生器",
    page_icon="📋",
    layout="centered"
)

# 標題
st.title("📋 勤務分配表產生器")
st.markdown("---")

# 說明
st.markdown("""
### 使用說明
1. 上傳**勤務表**（如：115.1月.勤1修1----勤務表.xls）
2. 上傳**空白分配表模板**（如：[20260120] 屏二分隊勤務分配表.xlsx）
3. 系統會**自動識別日期**，或手動選擇
4. 點擊「產生分配表」按鈕
5. 下載產生好的檔案
""")

st.markdown("---")


def extract_date_from_filename(filename: str) -> str:
    """
    從檔名中提取日期
    例如: "[20260120] 屏二分隊勤務分配表.xlsx" -> "0120"
    """
    # 嘗試匹配 [YYYYMMDD] 格式
    match = re.search(r'\[(\d{4})(\d{2})(\d{2})\]', filename)
    if match:
        month = match.group(2)
        day = match.group(3)
        return f"{month}{day}"

    # 嘗試匹配 YYYYMMDD 格式（無括號）
    match = re.search(r'(\d{4})(\d{2})(\d{2})', filename)
    if match:
        month = match.group(2)
        day = match.group(3)
        return f"{month}{day}"

    return None


# 檔案上傳區
col1, col2 = st.columns(2)

with col1:
    st.subheader("① 上傳勤務表")
    duty_file = st.file_uploader(
        "選擇勤務表檔案 (.xls)",
        type=['xls', 'xlsx'],
        key="duty"
    )

with col2:
    st.subheader("② 上傳空白分配表")
    template_file = st.file_uploader(
        "選擇分配表模板 (.xlsx)",
        type=['xlsx'],
        key="template"
    )

st.markdown("---")

# 日期處理
selected_date = None
available_dates = []

if duty_file is not None:
    try:
        available_dates = get_available_dates(duty_file)
        duty_file.seek(0)
    except Exception as e:
        st.error(f"讀取勤務表時發生錯誤：{str(e)}")

# 自動識別日期或手動選擇
if duty_file is not None and template_file is not None and available_dates:
    st.subheader("③ 確認日期")

    # 嘗試從檔名提取日期
    detected_date = extract_date_from_filename(template_file.name)

    if detected_date and detected_date in available_dates:
        # 自動識別成功
        month = detected_date[:2]
        day = detected_date[2:]
        st.success(f"✅ 已從檔名自動識別日期：**{month}月{day}日**")
        selected_date = detected_date

        # 提供手動修改的選項
        if st.checkbox("手動選擇其他日期"):
            date_options = {f"{d[:2]}月{d[2:]}日": d for d in available_dates}
            selected_display = st.selectbox(
                "選擇日期",
                options=list(date_options.keys()),
                index=list(date_options.values()).index(detected_date)
            )
            selected_date = date_options[selected_display]
    else:
        # 無法自動識別，顯示手動選擇
        if detected_date:
            st.warning(f"⚠️ 從檔名識別到日期 {detected_date[:2]}月{detected_date[2:]}日，但勤務表中沒有此日期")
        else:
            st.info("無法從檔名自動識別日期，請手動選擇")

        date_options = {f"{d[:2]}月{d[2:]}日": d for d in available_dates}
        selected_display = st.selectbox(
            "請選擇要產生分配表的日期",
            options=list(date_options.keys())
        )
        selected_date = date_options[selected_display]

elif duty_file is not None and not available_dates:
    st.error("無法從勤務表中找到有效的日期工作表")

st.markdown("---")

# 產生按鈕
st.subheader("④ 產生分配表")

if duty_file is not None and template_file is not None and selected_date is not None:
    if st.button("🚀 產生分配表", type="primary", use_container_width=True):
        try:
            with st.spinner("處理中..."):
                # 重設檔案指標
                duty_file.seek(0)
                template_file.seek(0)

                # 處理檔案
                result_file, filename = process_files(duty_file, template_file, selected_date)

            # 成功訊息
            st.success("✅ 分配表產生完成！")

            # 下載按鈕
            st.download_button(
                label="📥 下載分配表",
                data=result_file,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )

        except Exception as e:
            st.error(f"處理時發生錯誤：{str(e)}")
            st.exception(e)
else:
    st.info("請先完成上方步驟 ①②")

# 頁尾
st.markdown("---")
st.markdown(
    "<div style='text-align: center; color: gray;'>屏東第二分隊勤務分配表自動產生系統</div>",
    unsafe_allow_html=True
)
