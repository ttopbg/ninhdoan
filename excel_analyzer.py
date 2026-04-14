import streamlit as st
import pandas as pd
import os
from io import BytesIO

# ─── PAGE CONFIG ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Thống kê của Ninh Đoàng",
    page_icon="🧋",
    layout="centered",
)

# ─── CUSTOM CSS ─────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Be+Vietnam+Pro:wght@300;400;500;600;700&family=Playfair+Display:wght@600&display=swap');

/* Reset & base */
html, body, [class*="css"] {
    font-family: 'Be Vietnam Pro', sans-serif;
}

/* Background */
.stApp {
    background: linear-gradient(145deg, #f0f4f8 0%, #e8f0fe 50%, #f5f0ff 100%);
    min-height: 100vh;
}

/* Hide default streamlit elements */
#MainMenu, footer, header {visibility: hidden;}

/* ── Hero header ── */
.hero {
    text-align: center;
    padding: 2.5rem 1rem 1.5rem;
}
.hero h1 {
    font-family: 'Playfair Display', serif;
    font-size: 2.4rem;
    color: #2d3a5e;
    margin-bottom: 0.3rem;
    letter-spacing: -0.5px;
}
.hero p {
    color: #6b7a99;
    font-size: 1rem;
    font-weight: 300;
    margin: 0;
}

/* ── Card ── */
.card {
    background: rgba(255,255,255,0.80);
    backdrop-filter: blur(12px);
    border: 1px solid rgba(255,255,255,0.9);
    border-radius: 20px;
    padding: 2rem 2rem 1.6rem;
    box-shadow: 0 8px 32px rgba(60,80,160,0.08), 0 1.5px 4px rgba(60,80,160,0.04);
    margin-bottom: 1.4rem;
}

/* ── Section label ── */
.section-label {
    font-size: 0.72rem;
    font-weight: 700;
    letter-spacing: 1.8px;
    text-transform: uppercase;
    color: #8a96b8;
    margin-bottom: 0.6rem;
}

/* ── File uploader overrides ── */
[data-testid="stFileUploader"] {
    border: 2px dashed #b8c4e8 !important;
    border-radius: 14px !important;
    background: rgba(232,240,254,0.35) !important;
    transition: border-color 0.2s;
}
[data-testid="stFileUploader"]:hover {
    border-color: #6b88e6 !important;
}
[data-testid="stFileUploader"] label {
    color: #4a5a8a !important;
    font-weight: 500 !important;
}

/* ── Button ── */
.stButton > button {
    background: linear-gradient(135deg, #4f72e3 0%, #7b5cde 100%) !important;
    color: white !important;
    border: none !important;
    border-radius: 12px !important;
    padding: 0.75rem 2.2rem !important;
    font-family: 'Be Vietnam Pro', sans-serif !important;
    font-size: 1rem !important;
    font-weight: 600 !important;
    letter-spacing: 0.3px !important;
    box-shadow: 0 4px 18px rgba(79,114,227,0.35) !important;
    transition: all 0.2s ease !important;
    width: 100% !important;
}
.stButton > button:hover {
    transform: translateY(-2px) !important;
    box-shadow: 0 8px 28px rgba(79,114,227,0.45) !important;
}
.stButton > button:active {
    transform: translateY(0px) !important;
}

/* ── Download button ── */
[data-testid="stDownloadButton"] > button {
    background: linear-gradient(135deg, #2ecc8f 0%, #1aa879 100%) !important;
    color: white !important;
    border: none !important;
    border-radius: 12px !important;
    padding: 0.75rem 2rem !important;
    font-family: 'Be Vietnam Pro', sans-serif !important;
    font-weight: 600 !important;
    font-size: 1rem !important;
    box-shadow: 0 4px 18px rgba(46,204,143,0.30) !important;
    width: 100% !important;
}
[data-testid="stDownloadButton"] > button:hover {
    transform: translateY(-2px) !important;
    box-shadow: 0 8px 24px rgba(46,204,143,0.40) !important;
}

/* ── Metric cards ── */
.metrics-row {
    display: flex;
    gap: 1rem;
    margin: 1.2rem 0 0.6rem;
}
.metric-box {
    flex: 1;
    background: linear-gradient(135deg, #eef2ff 0%, #f5f0ff 100%);
    border: 1px solid #dde4fc;
    border-radius: 14px;
    padding: 1rem 1.2rem;
    text-align: center;
}
.metric-box .val {
    font-size: 1.9rem;
    font-weight: 700;
    color: #3d56c8;
    line-height: 1.1;
}
.metric-box .lbl {
    font-size: 0.78rem;
    color: #8a96b8;
    font-weight: 500;
    margin-top: 0.2rem;
    letter-spacing: 0.4px;
}

/* ── Status messages ── */
.msg-info {
    background: #eef2ff;
    border-left: 4px solid #6b88e6;
    border-radius: 8px;
    padding: 0.8rem 1rem;
    color: #3d56c8;
    font-size: 0.9rem;
    margin: 0.8rem 0;
}
.msg-success {
    background: #eafaf4;
    border-left: 4px solid #2ecc8f;
    border-radius: 8px;
    padding: 0.8rem 1rem;
    color: #1a8a60;
    font-size: 0.9rem;
    margin: 0.8rem 0;
}
.msg-warn {
    background: #fff8ec;
    border-left: 4px solid #f0a500;
    border-radius: 8px;
    padding: 0.8rem 1rem;
    color: #9a6700;
    font-size: 0.9rem;
    margin: 0.8rem 0;
}

/* ── Table ── */
[data-testid="stDataFrame"] {
    border-radius: 12px !important;
    overflow: hidden !important;
    border: 1px solid #e2e8f8 !important;
}

/* ── Divider ── */
hr {
    border: none;
    border-top: 1px solid #e4e9f5;
    margin: 1.2rem 0;
}
</style>
""", unsafe_allow_html=True)


# ─── HERO ────────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="hero">
    <h1>🙇 Thống kê thuế cơ sở và DN</h1>
    <p>Tổng hợp dữ liệu thuế cơ sở & Doanh nghiệp từ file Excel</p>
</div>
""", unsafe_allow_html=True)


# ─── UPLOAD CARD ─────────────────────────────────────────────────────────────────
# st.markdown('<div class="card">', unsafe_allow_html=True)
st.markdown('<div class="section-label">📁 Tải lên file Excel</div>', unsafe_allow_html=True)

uploaded_file = st.file_uploader(
    label="Chọn file Excel (.xlsx hoặc .xls)",
    type=["xlsx", "xls"],
    help="File cần có 3 cột: Tỉnh (A), Cơ sở (B), Doanh nghiệp (C)",
    label_visibility="collapsed",
)

if uploaded_file:
    st.markdown(f"""
    <div class="msg-info">
        ✅ Đã tải lên: <strong>{uploaded_file.name}</strong>
        &nbsp;|&nbsp; Kích thước: <strong>{uploaded_file.size / 1024:.1f} KB</strong>
    </div>
    """, unsafe_allow_html=True)

# st.markdown('</div>', unsafe_allow_html=True)


# ─── PROCESS CARD ────────────────────────────────────────────────────────────────
st.markdown('<div class="card">', unsafe_allow_html=True)
st.markdown('<div class="section-label">⚙️ Xử lý dữ liệu</div>', unsafe_allow_html=True)

run_btn = st.button("🚀 Bắt đầu thống kê", disabled=(uploaded_file is None))

if run_btn and uploaded_file:
    with st.spinner("Đang xử lý dữ liệu, vui lòng chờ..."):
        try:
            # ── Read file ──
            file_bytes = uploaded_file.read()
            df_raw = pd.read_excel(BytesIO(file_bytes), header=0, dtype=str)

            # ── Normalise columns A, B, C ──
            df_raw.columns = [str(c).strip() for c in df_raw.columns]

            # Grab first 3 columns regardless of header name
            col_tinh = df_raw.columns[0]
            col_co_so = df_raw.columns[1]
            col_dn = df_raw.columns[2]

            total_rows = len(df_raw)

            # ── Drop rows where any of the 3 key columns is blank / NaN ──
            df = df_raw[[col_tinh, col_co_so, col_dn]].copy()
            df = df.replace(r'^\s*$', pd.NA, regex=True)
            df = df.dropna(subset=[col_tinh, col_co_so, col_dn])
            dropped = total_rows - len(df)

            # ── Convert numeric columns ──
            df[col_co_so] = pd.to_numeric(df[col_co_so], errors='coerce')
            df[col_dn]    = pd.to_numeric(df[col_dn],    errors='coerce')
            df = df.dropna(subset=[col_co_so, col_dn])

            # ── Build summary ──
            summary = (
                df.groupby(col_tinh, sort=True)
                  .agg(
                      so_co_so=(col_co_so, 'sum'),
                      so_dn=(col_dn, 'sum'),
                  )
                  .reset_index()
            )
            summary.columns = ['Tỉnh', 'Số Cơ Sở', 'Số Doanh Nghiệp']
            summary['Số Cơ Sở']       = summary['Số Cơ Sở'].astype(int)
            summary['Số Doanh Nghiệp'] = summary['Số Doanh Nghiệp'].astype(int)

            # ── Build output Excel ──
            # Original data keeps all original columns
            df_out = df_raw.copy()

            # Ensure columns E, F, G exist (index 4, 5, 6)
            while len(df_out.columns) < 4:
                df_out[f'_col{len(df_out.columns)}'] = pd.NA

            # Add blank column D if not present
            if len(df_out.columns) == 3:
                df_out.insert(3, '_blank', pd.NA)

            # Add E, F, G columns with headers
            df_out['Tỉnh (tổng hợp)']  = pd.NA
            df_out['Số Cơ Sở']         = pd.NA
            df_out['Số Doanh Nghiệp']  = pd.NA

            e_col = df_out.columns[-3]
            f_col = df_out.columns[-2]
            g_col = df_out.columns[-1]

            # Fill summary values starting from row 0 (header already set by column names)
            for i, row in summary.iterrows():
                if i < len(df_out):
                    df_out.at[i, e_col] = row['Tỉnh']
                    df_out.at[i, f_col] = row['Số Cơ Sở']
                    df_out.at[i, g_col] = row['Số Doanh Nghiệp']

            # ── Save to BytesIO ──
            output_buffer = BytesIO()
            with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
                df_out.to_excel(writer, index=False, sheet_name='Data')

                # Style the output
                workbook  = writer.book
                worksheet = writer.sheets['Data']

                from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
                from openpyxl.utils import get_column_letter

                # Header row style
                header_fill = PatternFill(start_color="4F72E3", end_color="4F72E3", fill_type="solid")
                summary_fill = PatternFill(start_color="7B5CDE", end_color="7B5CDE", fill_type="solid")
                header_font = Font(color="FFFFFF", bold=True, name='Calibri', size=11)
                center_align = Alignment(horizontal='center', vertical='center')
                thin_border = Border(
                    left=Side(style='thin', color='DDEAFF'),
                    right=Side(style='thin', color='DDEAFF'),
                    bottom=Side(style='thin', color='DDEAFF'),
                )

                num_cols = len(df_out.columns)
                for col_idx in range(1, num_cols + 1):
                    cell = worksheet.cell(row=1, column=col_idx)
                    if col_idx <= num_cols - 3:
                        cell.fill = header_fill
                    else:
                        cell.fill = summary_fill
                    cell.font = header_font
                    cell.alignment = center_align

                # Data rows alternate fill
                alt_fill = PatternFill(start_color="F0F4FF", end_color="F0F4FF", fill_type="solid")
                summary_data_fill = PatternFill(start_color="F5F0FF", end_color="F5F0FF", fill_type="solid")

                for row_idx in range(2, len(df_out) + 2):
                    for col_idx in range(1, num_cols + 1):
                        cell = worksheet.cell(row=row_idx, column=col_idx)
                        cell.border = thin_border
                        cell.alignment = Alignment(vertical='center')
                        if row_idx % 2 == 0:
                            if col_idx > num_cols - 3:
                                cell.fill = summary_data_fill
                            else:
                                cell.fill = alt_fill

                # Auto-fit columns
                for col_idx in range(1, num_cols + 1):
                    col_letter = get_column_letter(col_idx)
                    max_len = 0
                    for row_idx in range(1, min(len(df_out) + 2, 500)):
                        val = worksheet.cell(row=row_idx, column=col_idx).value
                        if val:
                            max_len = max(max_len, len(str(val)))
                    worksheet.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 40)

                # Freeze header
                worksheet.freeze_panes = 'A2'

            output_buffer.seek(0)

            # ── Output filename ──
            base_name = os.path.splitext(uploaded_file.name)[0]
            output_filename = f"{base_name}_output.xlsx"

            # ── Store in session state ──
            st.session_state['output_data']     = output_buffer.getvalue()
            st.session_state['output_filename'] = output_filename
            st.session_state['summary']         = summary
            st.session_state['total_rows']      = total_rows
            st.session_state['dropped']         = dropped
            st.session_state['valid_rows']      = len(df)
            st.session_state['num_provinces']   = len(summary)

        except Exception as e:
            st.markdown(f'<div class="msg-warn">❌ Lỗi xử lý file: <strong>{e}</strong></div>', unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)


# ─── RESULTS ─────────────────────────────────────────────────────────────────────
if 'output_data' in st.session_state:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="section-label">📈 Kết quả thống kê</div>', unsafe_allow_html=True)

    # Metrics
    st.markdown(f"""
    <div class="metrics-row">
        <div class="metric-box">
            <div class="val">{st.session_state['total_rows']:,}</div>
            <div class="lbl">Tổng dòng gốc</div>
        </div>
        <div class="metric-box">
            <div class="val">{st.session_state['dropped']:,}</div>
            <div class="lbl">Dòng bỏ qua (trống)</div>
        </div>
        <div class="metric-box">
            <div class="val">{st.session_state['valid_rows']:,}</div>
            <div class="lbl">Dòng hợp lệ</div>
        </div>
        <div class="metric-box">
            <div class="val">{st.session_state['num_provinces']}</div>
            <div class="lbl">Số tỉnh duy nhất</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown('<hr>', unsafe_allow_html=True)

    # Summary table
    st.markdown('<div class="section-label">🗂 Bảng tổng hợp theo Tỉnh</div>', unsafe_allow_html=True)
    st.dataframe(
        st.session_state['summary'].style.format({'Số Cơ Sở': '{:,}', 'Số Doanh Nghiệp': '{:,}'}),
        use_container_width=True,
        hide_index=True,
    )

    st.markdown('<hr>', unsafe_allow_html=True)

    # Download
    st.markdown('<div class="section-label">💾 Tải file kết quả</div>', unsafe_allow_html=True)
    st.markdown(f"""
    <div class="msg-success">
        ✅ File đã sẵn sàng: <strong>{st.session_state['output_filename']}</strong>
    </div>
    """, unsafe_allow_html=True)

    st.download_button(
        label="⬇️  Tải xuống file Excel",
        data=st.session_state['output_data'],
        file_name=st.session_state['output_filename'],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.markdown('</div>', unsafe_allow_html=True)


# ─── FOOTER ──────────────────────────────────────────────────────────────────────
st.markdown("""
<div style="text-align:center; color:#b0b8d0; font-size:0.78rem; padding: 1.5rem 0 0.5rem; font-weight:300;">
    Thống Kê Tỉnh · Được xây dựng với Streamlit &amp; Pandas
</div>
""", unsafe_allow_html=True)
