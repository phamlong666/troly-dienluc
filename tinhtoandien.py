
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import io

st.set_page_config(page_title="Trợ lý AI Điện lực Định Hóa", layout="wide")
st.title(":zap: Trợ lý AI - Phân tích tổn thất điện năng")
st.markdown("### Phục vụ Điện lực Định Hóa - EVNNPC")

uploaded_files = st.file_uploader(
    "Tải lên các file Excel tổn thất (nhiều tháng)", 
    type=["xlsx", "xls"], 
    accept_multiple_files=True
)

if uploaded_files:
    all_data = []
    for file in uploaded_files:
        try:
            df = pd.read_excel(file, skiprows=6)
            df["Tên file"] = file.name
            all_data.append(df)
        except Exception as e:
            st.error(f"Không thể đọc file {file.name}: {str(e)}")

    if not all_data:
        st.warning("Không có file hợp lệ để xử lý.")
        st.stop()

    df_all = pd.concat(all_data, ignore_index=True)

    # Hiển thị các cột nếu không tìm được cột tổn thất
    tonthat_col = None
    for col in df_all.columns:
        if "tổn" in str(col).lower():
            tonthat_col = col
            break

    if not tonthat_col:
        st.error("❌ Không tìm thấy cột tổn thất trong dữ liệu.")
        with st.expander("📋 Danh sách cột hiện có trong dữ liệu:"):
            st.write(df_all.columns.tolist())
        st.stop()

    df_all[tonthat_col] = pd.to_numeric(df_all[tonthat_col], errors="coerce")

    name_cols = [col for col in df_all.columns if "trạm" in str(col).lower()]
    ten_col = name_cols[0] if name_cols else df_all.columns[0]

    thang_cols = [col for col in df_all.columns if "tháng" in str(col).lower() or "file" in str(col).lower()]
    thang_col = thang_cols[0] if thang_cols else "Tên file"

    st.sidebar.markdown("## 🔎 Bộ lọc dữ liệu")

    if ten_col not in df_all.columns or df_all[ten_col].dropna().empty:
        st.warning("Không có dữ liệu tên trạm hợp lệ.")
        st.stop()

    if thang_col not in df_all.columns or df_all[thang_col].dropna().empty:
        st.warning("Không có dữ liệu Tháng/Năm hợp lệ.")
        st.stop()

    selected_tba = st.sidebar.selectbox("Chọn trạm biến áp", sorted(df_all[ten_col].dropna().unique()))
    selected_month = st.sidebar.selectbox("Chọn Tháng/Năm", sorted(df_all[thang_col].dropna().unique()))

    df_filtered = df_all[
        (df_all[ten_col] == selected_tba) & 
        (df_all[thang_col] == selected_month)
    ]

    if df_filtered.empty:
        st.warning("Không có dữ liệu phù hợp với bộ lọc.")
        st.stop()

    st.markdown("### 📊 Phân tích tổn thất")
    df_filtered["Nguyên nhân"] = st.text_input("Nhập nguyên nhân tổn thất (ghi chung cho dòng này):", "")
    st.dataframe(df_filtered[[ten_col, thang_col, tonthat_col, "Nguyên nhân"]])

    avg = df_filtered[tonthat_col].mean()
    st.metric("Tổn thất trung bình (%)", f"{avg:.2f}")

    fig, ax = plt.subplots()
    df_filtered[tonthat_col].dropna().plot(kind="bar", ax=ax, color="orange")
    ax.set_title("Biểu đồ tổn thất")
    ax.set_ylabel("% tổn thất")
    st.pyplot(fig)

    def create_excel():
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_filtered.to_excel(writer, index=False, sheet_name="TonThat")
        return output.getvalue()

    st.download_button(
        "📥 Xuất báo cáo Excel",
        data=create_excel(),
        file_name=f"BaoCao_TonThat_{selected_tba}_{selected_month}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
