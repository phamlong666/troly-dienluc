import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import numpy as np

st.set_page_config(page_title="Báo cáo tổn thất TBA", layout="wide")
st.title("📥 Tải dữ liệu đầu vào - Báo cáo tổn thất")

st.markdown("### 🔍 Chọn loại dữ liệu tổn thất để tải lên:")

# --- Khởi tạo Session State cho dữ liệu tải lên ---
if 'df_tba_thang' not in st.session_state:
    st.session_state.df_tba_thang = None
if 'df_tba_luyke' not in st.session_state:
    st.session_state.df_tba_luyke = None
if 'df_tba_ck' not in st.session_state:
    st.session_state.df_tba_ck = None
if 'df_ha_thang' not in st.session_state:
    st.session_state.df_ha_thang = None
if 'df_ha_luyke' not in st.session_state: # Added
    st.session_state.df_ha_luyke = None
if 'df_ha_ck' not in st.session_state: # Added
    st.session_state.df_ha_ck = None
if 'df_trung_thang_tt' not in st.session_state:
    st.session_state.df_trung_thang_tt = None
if 'df_trung_luyke_tt' not in st.session_state: # Added
    st.session_state.df_trung_luyke_tt = None
if 'df_trung_ck_tt' not in st.session_state: # Added
    st.session_state.df_trung_ck_tt = None
if 'df_trung_thang_dy' not in st.session_state:
    st.session_state.df_trung_thang_dy = None
if 'df_trung_luyke_dy' not in st.session_state: # Added
    st.session_state.df_trung_luyke_dy = None
if 'df_trung_ck_dy' not in st.session_state: # Added
    st.session_state.df_trung_ck_dy = None
if 'df_dv_thang' not in st.session_state:
    st.session_state.df_dv_thang = None
if 'df_dv_luyke' not in st.session_state: # Added
    st.session_state.df_dv_luyke = None
if 'df_dv_ck' not in st.session_state: # Added
    st.session_state.df_dv_ck = None


# --- Nút "Làm mới" ---
if st.button("🔄 Làm mới dữ liệu"):
    st.session_state.df_tba_thang = None
    st.session_state.df_tba_luyke = None
    st.session_state.df_tba_ck = None
    st.session_state.df_ha_thang = None
    st.session_state.df_ha_luyke = None
    st.session_state.df_ha_ck = None
    st.session_state.df_trung_thang_tt = None
    st.session_state.df_trung_luyke_tt = None
    st.session_state.df_trung_ck_tt = None
    st.session_state.df_trung_thang_dy = None
    st.session_state.df_trung_luyke_dy = None
    st.session_state.df_trung_ck_dy = None
    st.session_state.df_dv_thang = None
    st.session_state.df_dv_luyke = None
    st.session_state.df_dv_ck = None
    st.experimental_rerun()


# Hàm phân loại tổn thất theo ngưỡng (di chuyển lên đầu để dễ tái sử dụng)
def phan_loai_nghiem(x):
    try:
        x = float(str(x).replace(",", "."))
    except (ValueError, AttributeError):
        return "Không rõ"
    if x < 2:
        return "<2%"
    elif 2 <= x < 3:
        return ">=2 và <3%"
    elif 3 <= x < 4:
        return ">=3 và <4%"
    elif 4 <= x < 5:
        return ">=4 và <5%"
    elif 5 <= x < 7:
        return ">=5 và <7%"
    else:
        return ">=7%"

# Hàm xử lý DataFrame và trả về số lượng TBA theo ngưỡng
def process_tba_data(df):
    if df is None:
        return None, None
    df_temp = pd.DataFrame()
    # Ensure column 14 (index 13) exists before accessing
    # Based on the problem description, data for mapping is from "Bảng Kết quả ánh xạ dữ liệu"
    # and the column index for "Tỷ lệ tổn thất" seems to be consistently 14.
    # However, if 'Bảng Kết quả ánh xạ dữ liệu' is already the mapped result,
    # then column indices might be different or column names might be available.
    # For now, let's assume it still follows the original column index logic for "Tỷ lệ tổn thất".
    # If the new sheet has column headers, it's better to use column names directly.
    # Example: df.get('Tỷ lệ tổn thất', pd.NA) or df.get('Tỷ lệ tổn thất', df.iloc[:, 14])
    # For robust handling, let's try to infer if it's the raw sheet or the mapped sheet.
    # If 'Tỷ lệ tổn thất' column is directly available, use it. Otherwise, fall back to iloc.

    if 'Tỷ lệ tổn thất' in df.columns:
        df_temp["Tỷ lệ tổn thất"] = df['Tỷ lệ tổn thất'].map(lambda x: f"{x:.2f}".replace(".", ",") if pd.notna(x) else "")
    elif df.shape[1] > 14: # Check for the 15th column (index 14)
        df_temp["Tỷ lệ tổn thất"] = df.iloc[:, 14].map(lambda x: f"{x:.2f}".replace(".", ",") if pd.notna(x) else "")
    else:
        st.warning("File Excel không có cột 'Tỷ lệ tổn thất' hoặc không đủ cột để tính toán. Vui lòng kiểm tra định dạng file và sheet.")
        return None, None

    df_temp["Ngưỡng"] = df_temp["Tỷ lệ tổn thất"].apply(phan_loai_nghiem)
    tong_so = len(df_temp)
    tong_theo_nguong = df_temp["Ngưỡng"].value_counts().reindex(["<2%", ">=2 và <3%", ">=3 và <4%", ">=4 và <5%", ">=5 và <7%", ">=7%"], fill_value=0)
    return tong_so, tong_theo_nguong


# Tạo các tiện ích con theo phân nhóm
with st.expander("🔌 Tổn thất các TBA công cộng"):
    temp_upload_tba_thang = st.file_uploader("📅 Tải dữ liệu TBA công cộng - Theo tháng", type=["xlsx"], key="tba_thang")
    if temp_upload_tba_thang:
        try:
            # IMPORTANT: Replace "Bảng Kết quả ánh xạ dữ liệu" with the EXACT sheet name you confirmed
            st.session_state.df_tba_thang = pd.read_excel(temp_upload_tba_thang, sheet_name="Bảng Kết quả ánh xạ dữ liệu", skiprows=6) # <<< CHANGE THIS LINE
            st.success("✅ Đã tải dữ liệu tổn thất TBA công cộng theo tháng!")
        except ValueError as e:
            st.error(f"Lỗi khi đọc sheet: {e}. Vui lòng kiểm tra tên sheet trong file Excel.")
            st.session_state.df_tba_thang = None
        except Exception as e:
            st.error(f"Đã xảy ra lỗi không mong muốn khi đọc file: {e}")
            st.session_state.df_tba_thang = None


    temp_upload_tba_luyke = st.file_uploader("📊 Tải dữ liệu TBA công cộng - Lũy kế", type=["xlsx"], key="tba_luyke")
    if temp_upload_tba_luyke:
        try:
            # IMPORTANT: Replace "Bảng Kết quả ánh xạ dữ liệu" with the EXACT sheet name you confirmed
            st.session_state.df_tba_luyke = pd.read_excel(temp_upload_tba_luyke, sheet_name="Bảng Kết quả ánh xạ dữ liệu", skiprows=6) # <<< CHANGE THIS LINE
            st.success("✅ Đã tải dữ liệu tổn thất TBA công cộng - Lũy kế!")
        except ValueError as e:
            st.error(f"Lỗi khi đọc sheet: {e}. Vui lòng kiểm tra tên sheet.")
            st.session_state.df_tba_luyke = None
        except Exception as e:
            st.error(f"Đã xảy ra lỗi không mong muốn khi đọc file Lũy kế: {e}")
            st.session_state.df_tba_luyke = None

    temp_upload_tba_ck = st.file_uploader("📈 Tải dữ liệu TBA công cộng - Cùng kỳ", type=["xlsx"], key="tba_ck")
    if temp_upload_tba_ck:
        try:
            # IMPORTANT: Replace "Bảng Kết quả ánh xạ dữ liệu" with the EXACT sheet name you confirmed
            st.session_state.df_tba_ck = pd.read_excel(temp_upload_tba_ck, sheet_name="Bảng Kết quả ánh xạ dữ liệu", skiprows=6) # <<< CHANGE THIS LINE
            st.success("✅ Đã tải dữ liệu tổn thất TBA công cộng - Cùng kỳ!")
        except ValueError as e:
            st.error(f"Lỗi khi đọc sheet: {e}. Vui lòng kiểm tra tên sheet.")
            st.session_state.df_tba_ck = None
        except Exception as e:
            st.error(f"Đã xảy ra lỗi không mong muốn khi đọc file Cùng kỳ: {e}")
            st.session_state.df_tba_ck = None


# --- Xử lý và hiển thị dữ liệu tổng hợp nếu có ít nhất một file được tải lên ---
if st.session_state.df_tba_thang is not None or \
   st.session_state.df_tba_luyke is not None or \
   st.session_state.df_tba_ck is not None:

    st.markdown("### 📊 Kết quả ánh xạ dữ liệu:")

    # Xử lý dữ liệu từng loại và chuẩn bị cho biểu đồ
    tong_so_thang, tong_theo_nguong_thang = process_tba_data(st.session_state.df_tba_thang)
    tong_so_luyke, tong_theo_nguong_luyke = process_tba_data(st.session_state.df_tba_luyke)
    tong_so_ck, tong_theo_nguong_ck = process_tba_data(st.session_state.df_tba_ck)

    col1, col2 = st.columns([2,2])

    with col1:
        st.markdown("#### 📊 Số lượng TBA theo ngưỡng tổn thất")
        fig_bar = go.Figure()
        colors = ['steelblue', 'darkorange', 'forestgreen', 'goldenrod', 'teal', 'red'] # Màu sắc cho từng ngưỡng

        # Thêm các thanh cho "Theo tháng"
        if tong_theo_nguong_thang is not None:
            fig_bar.add_trace(go.Bar(
                name='Theo tháng',
                x=tong_theo_nguong_thang.index,
                y=tong_theo_nguong_thang.values,
                text=tong_theo_nguong_thang.values,
                textposition='outside',
                textfont=dict(color='black')
            ))

        # Thêm các thanh cho "Lũy kế"
        if tong_theo_nguong_luyke is not None:
            fig_bar.add_trace(go.Bar(
                name='Lũy kế',
                x=tong_theo_nguong_luyke.index,
                y=tong_theo_nguong_luyke.values,
                text=tong_theo_nguong_luyke.values,
                textposition='outside',
                textfont=dict(color='black')
            ))

        # Thêm các thanh cho "Cùng kỳ"
        if tong_theo_nguong_ck is not None:
            fig_bar.add_trace(go.Bar(
                name='Cùng kỳ',
                x=tong_theo_nguong_ck.index,
                y=tong_theo_nguong_ck.values,
                text=tong_theo_nguong_ck.values,
                textposition='outside',
                textfont=dict(color='black')
            ))

        fig_bar.update_layout(
            barmode='group',
            height=400,
            xaxis_title='Ngưỡng tổn thất',
            yaxis_title='Số lượng TBA',
            margin=dict(l=20, r=20, t=40, b=40),
            legend_title_text='Loại dữ liệu'
        )
        st.plotly_chart(fig_bar, use_container_width=True)

    with col2:
        st.markdown("#### 🧩 Tỷ trọng TBA theo ngưỡng tổn thất")

        if tong_theo_nguong_thang is not None:
            st.markdown(f"##### Theo tháng (Tổng số: {tong_so_thang})")
            fig_pie_thang = go.Figure(data=[
                go.Pie(
                    labels=tong_theo_nguong_thang.index,
                    values=tong_theo_nguong_thang.values,
                    hole=0.5,
                    marker=dict(colors=colors),
                    textinfo='percent+label',
                    name='Theo tháng'
                )
            ])
            fig_pie_thang.update_layout(height=300, margin=dict(l=20, r=20, t=40, b=40), showlegend=False)
            st.plotly_chart(fig_pie_thang, use_container_width=True)

        if tong_theo_nguong_luyke is not None:
            st.markdown(f"##### Lũy kế (Tổng số: {tong_so_luyke})")
            fig_pie_luyke = go.Figure(data=[
                go.Pie(
                    labels=tong_theo_nguong_luyke.index,
                    values=tong_theo_nguong_luyke.values,
                    hole=0.5,
                    marker=dict(colors=colors),
                    textinfo='percent+label',
                    name='Lũy kế'
                )
            ])
            fig_pie_luyke.update_layout(height=300, margin=dict(l=20, r=20, t=40, b=40), showlegend=False)
            st.plotly_chart(fig_pie_luyke, use_container_width=True)

        if tong_theo_nguong_ck is not None:
            st.markdown(f"##### Cùng kỳ (Tổng số: {tong_so_ck})")
            fig_pie_ck = go.Figure(data=[
                go.Pie(
                    labels=tong_theo_nguong_ck.index,
                    values=tong_theo_nguong_ck.values,
                    hole=0.5,
                    marker=dict(colors=colors),
                    textinfo='percent+label',
                    name='Cùng kỳ'
                )
            ])
            fig_pie_ck.update_layout(height=300, margin=dict(l=20, r=20, t=40, b=40), showlegend=False)
            st.plotly_chart(fig_pie_ck, use_container_width=True)


    # Hiển thị DataFrame ánh xạ cho file "Theo tháng" nếu nó tồn tại
    if st.session_state.df_tba_thang is not None:
        st.markdown("##### Dữ liệu TBA công cộng - Theo tháng:")
        df_test = st.session_state.df_tba_thang
        df_result = pd.DataFrame()
        # Ensure column indices are correct based on the 'Bảng Kết quả ánh xạ dữ liệu' sheet
        # Assuming the structure is consistent:
        # Col 2 -> Tên TBA (index 2)
        # Col 3 -> Công suất (index 3)
        # Col 6 -> Điện nhận (index 6)
        # Col 7 -> Điện tổn thất (index 7)
        # Col 13 -> Điện tổn thất (index 13) (assuming this is where the numerical value comes from)
        # Col 14 -> Tỷ lệ tổn thất (index 14)
        # Col 15 -> Kế hoạch (index 15)
        # Col 16 -> So sánh (index 16)
        # You might need to adjust these indices if the 'Bảng Kết quả ánh xạ dữ liệu' sheet has a different layout.

        # Let's add a robust check for column existence if possible
        required_cols_iloc = {
            "Tên TBA": 2,
            "Công suất": 3,
            "Điện nhận": 6,
            "Điện tổn thất (cho phép tính TP)": 7, # Assuming this is for Thương phẩm calculation
            "Điện tổn thất (cho mapping)": 13, # Assuming this is the raw loss value
            "Tỷ lệ tổn thất": 14,
            "Kế hoạch": 15,
            "So sánh": 16
        }

        # Check if df_test has enough columns before assigning
        if df_test.shape[1] > max(required_cols_iloc.values()):
            df_result["STT"] = range(1, len(df_test) + 1)
            df_result["Tên TBA"] = df_test.iloc[:, required_cols_iloc["Tên TBA"]]
            df_result["Công suất"] = df_test.iloc[:, required_cols_iloc["Công suất"]]
            df_result["Điện nhận"] = df_test.iloc[:, required_cols_iloc["Điện nhận"]]
            df_result["Thương phẩm"] = df_test.iloc[:, required_cols_iloc["Điện nhận"]] - df_test.iloc[:, required_cols_iloc["Điện tổn thất (cho phép tính TP)"]]
            df_result["Điện tổn thất"] = df_test.iloc[:, required_cols_iloc["Điện tổn thất (cho mapping)"]].round(0).astype("Int64")
            df_result["Tỷ lệ tổn thất"] = df_test.iloc[:, required_cols_iloc["Tỷ lệ tổn thất"]].map(lambda x: f"{x:.2f}".replace(".", ",") if pd.notna(x) else "")
            df_result["Kế hoạch"] = df_test.iloc[:, required_cols_iloc["Kế hoạch"]].map(lambda x: f"{x:.2f}".replace(".", ",") if pd.notna(x) else "")
            df_result["So sánh"] = df_test.iloc[:, required_cols_iloc["So sánh"]].map(lambda x: f"{x:.2f}".replace(".", ",") if pd.notna(x) else "")
            st.dataframe(df_result)
        else:
            st.warning("Dữ liệu TBA công cộng - Theo tháng: Không đủ cột để ánh xạ. Vui lòng kiểm tra cấu trúc sheet 'TÊN_SHEET_CHÍNH_XÁC_CỦA_BẠN'.")


with st.expander("⚡ Tổn thất hạ thế"):
    upload_ha_thang = st.file_uploader("📅 Tải dữ liệu hạ áp - Theo tháng", type=["xlsx"], key="ha_thang")
    if upload_ha_thang:
        try:
            # IMPORTANT: Replace "Bảng Kết quả ánh xạ dữ liệu" with the EXACT sheet name you confirmed
            st.session_state.df_ha_thang = pd.read_excel(upload_ha_thang, sheet_name="Bảng Kết quả ánh xạ dữ liệu", skiprows=6) # <<< CHANGE THIS LINE
            st.success("✅ Đã tải dữ liệu tổn thất hạ áp - Theo tháng!")
        except ValueError as e:
            st.error(f"Lỗi khi đọc sheet hạ áp theo tháng: {e}. Vui lòng kiểm tra tên sheet trong file Excel.")
            st.session_state.df_ha_thang = None
        except Exception as e:
            st.error(f"Đã xảy ra lỗi không mong muốn khi đọc file hạ áp theo tháng: {e}")
            st.session_state.df_ha_thang = None

    upload_ha_luyke = st.file_uploader("📊 Tải dữ liệu hạ áp - Lũy kế", type=["xlsx"], key="ha_luyke")
    if upload_ha_luyke:
        try:
            # IMPORTANT: Replace "Bảng Kết quả ánh xạ dữ liệu" with the EXACT sheet name you confirmed
            st.session_state.df_ha_luyke = pd.read_excel(upload_ha_luyke, sheet_name="Bảng Kết quả ánh xạ dữ liệu", skiprows=6) # <<< CHANGE THIS LINE
            st.success("✅ Đã tải dữ liệu tổn thất hạ áp - Lũy kế!")
        except ValueError as e:
            st.error(f"Lỗi khi đọc sheet hạ áp lũy kế: {e}. Vui lòng kiểm tra tên sheet trong file Excel.")
            st.session_state.df_ha_luyke = None
        except Exception as e:
            st.error(f"Đã xảy ra lỗi không mong muốn khi đọc file hạ áp lũy kế: {e}")
            st.session_state.df_ha_luyke = None

    upload_ha_ck = st.file_uploader("📈 Tải dữ liệu hạ áp - Cùng kỳ", type=["xlsx"], key="ha_ck")
    if upload_ha_ck:
        try:
            # IMPORTANT: Replace "Bảng Kết quả ánh xạ dữ liệu" with the EXACT sheet name you confirmed
            st.session_state.df_ha_ck = pd.read_excel(upload_ha_ck, sheet_name="Bảng Kết quả ánh xạ dữ liệu", skiprows=6) # <<< CHANGE THIS LINE
            st.success("✅ Đã tải dữ liệu tổn thất hạ áp - Cùng kỳ!")
        except ValueError as e:
            st.error(f"Lỗi khi đọc sheet hạ áp cùng kỳ: {e}. Vui lòng kiểm tra tên sheet trong file Excel.")
            st.session_state.df_ha_ck = None
        except Exception as e:
            st.error(f"Đã xảy ra lỗi không mong muốn khi đọc file hạ áp cùng kỳ: {e}")
            st.session_state.df_ha_ck = None


with st.expander("⚡ Tổn thất trung thế"):
    upload_trung_thang_tt = st.file_uploader("📅 Tải dữ liệu Trung áp - Theo tháng", type=["xlsx"], key="trung_thang_tt")
    if upload_trung_thang_tt:
        try:
            # IMPORTANT: Replace "Bảng Kết quả ánh xạ dữ liệu" with the EXACT sheet name you confirmed
            st.session_state.df_trung_thang_tt = pd.read_excel(upload_trung_thang_tt, sheet_name="Bảng Kết quả ánh xạ dữ liệu", skiprows=6) # <<< CHANGE THIS LINE
            st.success("✅ Đã tải dữ liệu tổn thất Trung áp (Trung thế) - Theo tháng!")
        except ValueError as e:
            st.error(f"Lỗi khi đọc sheet trung áp (TT) theo tháng: {e}. Vui lòng kiểm tra tên sheet trong file Excel.")
            st.session_state.df_trung_thang_tt = None
        except Exception as e:
            st.error(f"Đã xảy ra lỗi không mong muốn khi đọc file trung áp (TT) theo tháng: {e}")
            st.session_state.df_trung_thang_tt = None

    upload_trung_luyke_tt = st.file_uploader("📊 Tải dữ liệu Trung áp - Lũy kế", type=["xlsx"], key="trung_luyke_tt")
    if upload_trung_luyke_tt:
        try:
            # IMPORTANT: Replace "Bảng Kết quả ánh xạ dữ liệu" with the EXACT sheet name you confirmed
            st.session_state.df_trung_luyke_tt = pd.read_excel(upload_trung_luyke_tt, sheet_name="Bảng Kết quả ánh xạ dữ liệu", skiprows=6) # <<< CHANGE THIS LINE
            st.success("✅ Đã tải dữ liệu tổn thất Trung áp (Trung thế) - Lũy kế!")
        except ValueError as e:
            st.error(f"Lỗi khi đọc sheet trung áp (TT) lũy kế: {e}. Vui lòng kiểm tra tên sheet trong file Excel.")
            st.session_state.df_trung_luyke_tt = None
        except Exception as e:
            st.error(f"Đã xảy ra lỗi không mong muốn khi đọc file trung áp (TT) lũy kế: {e}")
            st.session_state.df_trung_luyke_tt = None

    upload_trung_ck_tt = st.file_uploader("📈 Tải dữ liệu Trung áp - Cùng kỳ", type=["xlsx"], key="trung_ck_tt")
    if upload_trung_ck_tt:
        try:
            # IMPORTANT: Replace "Bảng Kết quả ánh xạ dữ liệu" with the EXACT sheet name you confirmed
            st.session_state.df_trung_ck_tt = pd.read_excel(upload_trung_ck_tt, sheet_name="Bảng Kết quả ánh xạ dữ liệu", skiprows=6) # <<< CHANGE THIS LINE
            st.success("✅ Đã tải dữ liệu tổn thất Trung áp (Trung thế) - Cùng kỳ!")
        except ValueError as e:
            st.error(f"Lỗi khi đọc sheet trung áp (TT) cùng kỳ: {e}. Vui lòng kiểm tra tên sheet trong file Excel.")
            st.session_state.df_trung_ck_tt = None
        except Exception as e:
            st.error(f"Đã xảy ra lỗi không mong muốn khi đọc file trung áp (TT) cùng kỳ: {e}")
            st.session_state.df_trung_ck_tt = None


with st.expander("⚡ Tổn thất các đường dây trung thế"):
    upload_trung_thang_dy = st.file_uploader("📅 Tải dữ liệu Trung áp - Theo tháng", type=["xlsx"], key="trung_thang_dy")
    if upload_trung_thang_dy:
        try:
            # IMPORTANT: Replace "Bảng Kết quả ánh xạ dữ liệu" with the EXACT sheet name you confirmed
            st.session_state.df_trung_thang_dy = pd.read_excel(upload_trung_thang_dy, sheet_name="Bảng Kết quả ánh xạ dữ liệu", skiprows=6) # <<< CHANGE THIS LINE
            st.success("✅ Đã tải dữ liệu tổn thất Đường dây Trung thế - Theo tháng!")
        except ValueError as e:
            st.error(f"Lỗi khi đọc sheet đường dây trung thế theo tháng: {e}. Vui lòng kiểm tra tên sheet trong file Excel.")
            st.session_state.df_trung_thang_dy = None
        except Exception as e:
            st.error(f"Đã xảy ra lỗi không mong muốn khi đọc file đường dây trung thế theo tháng: {e}")
            st.session_state.df_trung_thang_dy = None

    upload_trung_luyke_dy = st.file_uploader("📊 Tải dữ liệu Trung áp - Lũy kế", type=["xlsx"], key="trung_luyke_dy")
    if upload_trung_luyke_dy:
        try:
            # IMPORTANT: Replace "Bảng Kết quả ánh xạ dữ liệu" with the EXACT sheet name you confirmed
            st.session_state.df_trung_luyke_dy = pd.read_excel(upload_trung_luyke_dy, sheet_name="Bảng Kết quả ánh xạ dữ liệu", skiprows=6) # <<< CHANGE THIS LINE
            st.success("✅ Đã tải dữ liệu tổn thất Đường dây Trung thế - Lũy kế!")
        except ValueError as e:
            st.error(f"Lỗi khi đọc sheet đường dây trung thế lũy kế: {e}. Vui lòng kiểm tra tên sheet trong file Excel.")
            st.session_state.df_trung_luyke_dy = None
        except Exception as e:
            st.error(f"Đã xảy ra lỗi không mong muốn khi đọc file đường dây trung thế lũy kế: {e}")
            st.session_state.df_trung_luyke_dy = None

    upload_trung_ck_dy = st.file_uploader("📈 Tải dữ liệu Trung áp - Cùng kỳ", type=["xlsx"], key="trung_ck_dy")
    if upload_trung_ck_dy:
        try:
            # IMPORTANT: Replace "Bảng Kết quả ánh xạ dữ liệu" with the EXACT sheet name you confirmed
            st.session_state.df_trung_ck_dy = pd.read_excel(upload_trung_ck_dy, sheet_name="Bảng Kết quả ánh xạ dữ liệu", skiprows=6) # <<< CHANGE THIS LINE
            st.success("✅ Đã tải dữ liệu tổn thất Đường dây Trung thế - Cùng kỳ!")
        except ValueError as e:
            st.error(f"Lỗi khi đọc sheet đường dây trung thế cùng kỳ: {e}. Vui lòng kiểm tra tên sheet trong file Excel.")
            st.session_state.df_trung_ck_dy = None
        except Exception as e:
            st.error(f"Đã xảy ra lỗi không mong muốn khi đọc file đường dây trung thế cùng kỳ: {e}")
            st.session_state.df_trung_ck_dy = None


with st.expander("🏢 Tổn thất toàn đơn vị"):
    upload_dv_thang = st.file_uploader("📅 Tải dữ liệu Đơn vị - Theo tháng", type=["xlsx"], key="dv_thang")
    if upload_dv_thang:
        try:
            # IMPORTANT: Replace "Bảng Kết quả ánh xạ dữ liệu" with the EXACT sheet name you confirmed
            st.session_state.df_dv_thang = pd.read_excel(upload_dv_thang, sheet_name="Bảng Kết quả ánh xạ dữ liệu", skiprows=6) # <<< CHANGE THIS LINE
            st.success("✅ Đã tải dữ liệu tổn thất Toàn đơn vị - Theo tháng!")
        except ValueError as e:
            st.error(f"Lỗi khi đọc sheet đơn vị theo tháng: {e}. Vui lòng kiểm tra tên sheet trong file Excel.")
            st.session_state.df_dv_thang = None
        except Exception as e:
            st.error(f"Đã xảy ra lỗi không mong muốn khi đọc file đơn vị theo tháng: {e}")
            st.session_state.df_dv_thang = None

    upload_dv_luyke = st.file_uploader("📊 Tải dữ liệu Đơn vị - Lũy kế", type=["xlsx"], key="dv_luyke")
    if upload_dv_luyke:
        try:
            # IMPORTANT: Replace "Bảng Kết quả ánh xạ dữ liệu" with the EXACT sheet name you confirmed
            st.session_state.df_dv_luyke = pd.read_excel(upload_dv_luyke, sheet_name="Bảng Kết quả ánh xạ dữ liệu", skiprows=6) # <<< CHANGE THIS LINE
            st.success("✅ Đã tải dữ liệu tổn thất Toàn đơn vị - Lũy kế!")
        except ValueError as e:
            st.error(f"Lỗi khi đọc sheet đơn vị lũy kế: {e}. Vui lòng kiểm tra tên sheet trong file Excel.")
            st.session_state.df_dv_luyke = None
        except Exception as e:
            st.error(f"Đã xảy ra lỗi không mong muốn khi đọc file đơn vị lũy kế: {e}")
            st.session_state.df_dv_luyke = None

    upload_dv_ck = st.file_uploader("📈 Tải dữ liệu Đơn vị - Cùng kỳ", type=["xlsx"], key="dv_ck")
    if upload_dv_ck:
        try:
            # IMPORTANT: Replace "Bảng Kết quả ánh xạ dữ liệu" with the EXACT sheet name you confirmed
            st.session_state.df_dv_ck = pd.read_excel(upload_dv_ck, sheet_name="Bảng Kết quả ánh xạ dữ liệu", skiprows=6) # <<< CHANGE THIS LINE
            st.success("✅ Đã tải dữ liệu tổn thất Toàn đơn vị - Cùng kỳ!")
        except ValueError as e:
            st.error(f"Lỗi khi đọc sheet đơn vị cùng kỳ: {e}. Vui lòng kiểm tra tên sheet trong file Excel.")
            st.session_state.df_dv_ck = None
        except Exception as e:
            st.error(f"Đã xảy ra lỗi không mong muốn khi đọc file đơn vị cùng kỳ: {e}")
            st.session_state.df_dv_ck = None