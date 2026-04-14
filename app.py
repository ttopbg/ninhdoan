import streamlit as st
import pandas as pd
import os
from io import BytesIO

# ── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Thống kê Cơ quan theo Tỉnh",
    page_icon="📊",
    layout="centered",
)

# ── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Be+Vietnam+Pro:wght@300;400;600;700&display=swap');

html, body, [class*="css"] {
    font-family: 'Be Vietnam Pro', sans-serif;
}

/* Main background */
.stApp {
    background: linear-gradient(135deg, #0f2027 0%, #203a43 50%, #2c5364 100%);
    min-height: 100vh;
}

/* Title */
h1 {
    color: #e0f7fa !important;
    font-weight: 700 !important;
    letter-spacing: -0.5px;
    text-align: center;
    padding-bottom: 4px;
}

/* Subtitle */
.subtitle {
    text-align: center;
    color: #80cbc4;
    font-size: 0.95rem;
    margin-bottom: 2rem;
    margin-top: -0.5rem;
}

/* Card wrapper */
.card {
    background: rgba(255, 255, 255, 0.06);
    border: 1px solid rgba(255,255,255,0.12);
    border-radius: 16px;
    padding: 2rem 2.5rem;
    backdrop-filter: blur(12px);
    box-shadow: 0 8px 32px rgba(0,0,0,0.3);
    margin-bottom: 1.5rem;
}

/* File uploader tweak */
[data-testid="stFileUploader"] {
    border: 2px dashed #4dd0e1 !important;
    border-radius: 12px !important;
    padding: 1rem !important;
    background: rgba(77, 208, 225, 0.05) !important;
}

/* Button */
.stButton > button {
    background: linear-gradient(90deg, #00b4d8, #0077b6) !important;
    color: white !important;
    border: none !important;
    border-radius: 10px !important;
    font-weight: 600 !important;
    font-size: 1rem !important;
    padding: 0.65rem 2.5rem !important;
    letter-spacing: 0.3px;
    transition: opacity 0.2s;
    width: 100%;
}
.stButton > button:hover {
    opacity: 0.88 !important;
}

/* Download button */
[data-testid="stDownloadButton"] > button {
    background: linear-gradient(90deg, #43e97b, #38f9d7) !important;
    color: #0d1b2a !important;
    border: none !important;
    border-radius: 10px !important;
    font-weight: 700 !important;
    font-size: 1rem !important;
    padding: 0.65rem 2.5rem !important;
    width: 100%;
}

/* Metrics */
[data-testid="metric-container"] {
    background: rgba(255,255,255,0.07);
    border: 1px solid rgba(255,255,255,0.1);
    border-radius: 12px;
    padding: 1rem;
}
[data-testid="stMetricValue"] {
    color: #4dd0e1 !important;
    font-weight: 700 !important;
}
[data-testid="stMetricLabel"] {
    color: #b2dfdb !important;
}

/* Dataframe */
[data-testid="stDataFrame"] {
    border-radius: 10px;
    overflow: hidden;
}

/* Success / info / warning */
.stSuccess, .stInfo {
    border-radius: 10px !important;
}

/* Divider colour */
hr { border-color: rgba(255,255,255,0.1) !important; }

/* Section headers */
h3 { color: #e0f7fa !important; }
p, li { color: #cfd8dc !important; }
</style>
""", unsafe_allow_html=True)

# ── Header ────────────────────────────────────────────────────────────────────
st.title("📊 Thống kê Cơ quan theo Tỉnh")
st.markdown('<p class="subtitle">Tải lên file Excel → nhận ngay báo cáo tổng hợp số cơ quan của mỗi tỉnh</p>', unsafe_allow_html=True)

# ── Upload card ───────────────────────────────────────────────────────────────
st.markdown('<div class="card">', unsafe_allow_html=True)
st.markdown("### 📁 Chọn file Excel đầu vào")
uploaded_file = st.file_uploader(
    label="Kéo thả hoặc Browse file (.xlsx / .xls)",
    type=["xlsx", "xls"],
    label_visibility="collapsed",
)

col_info1, col_info2 = st.columns(2)
with col_info1:
    st.markdown("""
    **Yêu cầu file:**
    - Cột A: Tỉnh
    - Cột B: Cơ quan
    """)
with col_info2:
    st.markdown("""
    **Xử lý tự động:**
    - Bỏ qua dòng Tỉnh hoặc Cơ quan trống
    - Tổng hợp tại cột D & E
    """)
st.markdown('</div>', unsafe_allow_html=True)

# ── Process ───────────────────────────────────────────────────────────────────
if uploaded_file is not None:
    original_name = os.path.splitext(uploaded_file.name)[0]
    output_filename = f"{original_name}_output.xlsx"

    st.markdown('<div class="card">', unsafe_allow_html=True)
    start_btn = st.button("▶ Bắt đầu xử lý", use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

    if start_btn:
        with st.spinner("⏳ Đang đọc và xử lý dữ liệu…"):
            try:
                # Read with chunking hint (chunksize not available for xlsx but read_excel is fast enough)
                df = pd.read_excel(uploaded_file, header=0, dtype=str)

                # Normalise: take only first 2 columns regardless of header name
                df = df.iloc[:, :2]
                df.columns = ["Tinh", "Co_quan"]

                total_rows = len(df)

                # Drop rows where Tinh OR Co_quan is empty/NaN
                df_clean = df.dropna(subset=["Tinh", "Co_quan"])
                df_clean = df_clean[
                    (df_clean["Tinh"].str.strip() != "") &
                    (df_clean["Co_quan"].str.strip() != "")
                ].copy()

                skipped = total_rows - len(df_clean)

                # ── Summary: count unique Co_quan per Tinh ──────────────────
                summary = (
                    df_clean.groupby("Tinh", sort=True)["Co_quan"]
                    .nunique()
                    .reset_index()
                )
                summary.columns = ["Tỉnh", "Số cơ quan"]
                summary = summary.sort_values("Tỉnh").reset_index(drop=True)

                # ── Build output Excel ───────────────────────────────────────
                output = BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    # Write original cleaned data to cols A, B (index=False)
                    df_clean.rename(columns={"Tinh": "Tỉnh", "Co_quan": "Cơ quan"}).to_excel(
                        writer, sheet_name="Sheet1", index=False, startrow=0, startcol=0
                    )

                    ws = writer.sheets["Sheet1"]

                    # Write summary header to col D (index 3) and E (index 4)
                    ws.cell(row=1, column=4, value="Tỉnh")
                    ws.cell(row=1, column=5, value="Số cơ quan")

                    for i, row in summary.iterrows():
                        ws.cell(row=i + 2, column=4, value=row["Tỉnh"])
                        ws.cell(row=i + 2, column=5, value=row["Số cơ quan"])

                    # Auto-fit column widths
                    from openpyxl.utils import get_column_letter
                    for col_idx in range(1, 6):
                        max_len = 0
                        col_letter = get_column_letter(col_idx)
                        for cell in ws[col_letter]:
                            if cell.value:
                                max_len = max(max_len, len(str(cell.value)))
                        ws.column_dimensions[col_letter].width = min(max_len + 4, 40)

                output.seek(0)

                # ── Metrics ──────────────────────────────────────────────────
                st.markdown("---")
                st.markdown("### ✅ Kết quả xử lý")
                m1, m2, m3, m4 = st.columns(4)
                m1.metric("Tổng dòng", f"{total_rows:,}")
                m2.metric("Dòng hợp lệ", f"{len(df_clean):,}")
                m3.metric("Dòng bỏ qua", f"{skipped:,}")
                m4.metric("Số tỉnh", f"{len(summary):,}")

                # ── Preview summary ──────────────────────────────────────────
                st.markdown("### 📋 Bảng tổng hợp")
                st.dataframe(
                    summary.style.background_gradient(
                        subset=["Số cơ quan"], cmap="Blues"
                    ),
                    use_container_width=True,
                    height=min(400, (len(summary) + 1) * 35 + 3),
                )

                # ── Download ─────────────────────────────────────────────────
                st.markdown("### 💾 Tải file kết quả")
                st.download_button(
                    label=f"⬇ Tải xuống  {output_filename}",
                    data=output,
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

                st.success(f"Đã xử lý thành công! File output: **{output_filename}**")

            except Exception as e:
                st.error(f"❌ Lỗi khi xử lý file: {e}")
else:
    st.info("👆 Vui lòng tải lên file Excel để bắt đầu.")

# ── Footer ────────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown(
    '<p style="text-align:center; color:#546e7a; font-size:0.8rem;">Thống kê Cơ quan theo Tỉnh · Streamlit App</p>',
    unsafe_allow_html=True,
)
