"""
勤務分配表自動填入系統
Streamlit 網頁介面
"""

import streamlit as st
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
2. 上傳**空白分配表模板**（如：屏二分隊勤務分配表.xlsx）
3. 選擇要產生的**日期**
4. 點擊「產生分配表」按鈕
5. 下載產生好的檔案
""")

st.markdown("---")

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

# 日期選擇（只有在上傳勤務表後才顯示）
if duty_file is not None:
    st.subheader("③ 選擇日期")

    try:
        available_dates = get_available_dates(duty_file)
        # 重設檔案指標
        duty_file.seek(0)

        if available_dates:
            # 將日期格式化顯示
            date_options = {f"{d[:2]}月{d[2:]}日": d for d in available_dates}
            selected_display = st.selectbox(
                "請選擇要產生分配表的日期",
                options=list(date_options.keys())
            )
            selected_date = date_options[selected_display]
        else:
            st.error("無法從勤務表中找到有效的日期工作表")
            selected_date = None

    except Exception as e:
        st.error(f"讀取勤務表時發生錯誤：{str(e)}")
        selected_date = None
else:
    selected_date = None

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
    st.info("請先完成上方步驟 ①②③")

# 頁尾
st.markdown("---")
st.markdown(
    "<div style='text-align: center; color: gray;'>屏東第二分隊勤務分配表自動產生系統</div>",
    unsafe_allow_html=True
)
