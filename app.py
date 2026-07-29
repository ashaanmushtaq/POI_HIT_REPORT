import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io
import re
from datetime import datetime

st.set_page_config(
    page_title="POI Hit Report Processor | Dev by Ashaan",
    page_icon="📍",
    layout="wide"
)

# ------------------------------------------------------------------------------
# AESTHETIC UI STYLING
# ------------------------------------------------------------------------------
st.markdown("""
    <style>
    .header-card {
        background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%);
        border-radius: 16px;
        padding: 24px;
        box-shadow: 0 10px 25px rgba(0,0,0,0.2);
        margin-bottom: 25px;
        color: white;
        border: 1px solid #334155;
    }
    .header-title {
        font-family: 'Segoe UI', sans-serif;
        font-size: 26px;
        font-weight: 700;
        text-align: center;
        letter-spacing: 0.5px;
        color: #F8FAFC;
    }
    .header-sub {
        font-family: 'Segoe UI', sans-serif;
        text-align: center;
        color: #94A3B8;
        font-size: 14px;
        margin-top: 4px;
    }
    .header-sub b {
        color: #38BDF8;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="header-card">
    <div class="header-title">📍 POI HIT & CONTAINER AUDIT DASHBOARD</div>
    <div class="header-sub">Developed by: <b>Muhammad Ashaan</b> | Bellanix Tech</div>
</div>
""", unsafe_allow_html=True)

# ------------------------------------------------------------------------------
# FILE UPLOADER
# ------------------------------------------------------------------------------
uploaded_file = st.file_uploader("📥 Upload POI Hit Report (.xlsx)", type=["xlsx"])

if uploaded_file:
    try:
        df_raw = pd.read_excel(uploaded_file, sheet_name=0, header=None)
        
        # Detect Metadata & Header Row
        meta_title = str(df_raw.iloc[0, 0]).strip() if pd.notna(df_raw.iloc[0, 0]) else "POI Hit Report"
        generated_at = str(df_raw.iloc[1, 1]).strip() if pd.notna(df_raw.iloc[1, 1]) else datetime.now().strftime("%Y-%m-%d")
        time_period = str(df_raw.iloc[2, 1]).strip() if pd.notna(df_raw.iloc[2, 1]) else "N/A"

        # Find Data Header Row
        header_row_idx = -1
        for r in range(min(10, len(df_raw))):
            row_str = [str(x).strip().upper() for x in df_raw.iloc[r].tolist()]
            if 'SER#' in row_str and 'CONT#' in row_str and 'VEHICLE' in row_str:
                header_row_idx = r
                break

        if header_row_idx == -1:
            st.error("❌ Invalid POI Hit Report format. 'Ser#', 'Cont#', and 'Vehicle' columns were not found.")
        else:
            headers = [str(x).strip() for x in df_raw.iloc[header_row_idx].tolist()]
            df_data = df_raw.iloc[header_row_idx + 1:].copy()
            df_data.columns = headers

            # Separate Picked vs Not Picked
            df_data['Vehicle_Clean'] = df_data['Vehicle'].astype(str).str.strip()
            
            not_picked_df = df_data[df_data['Vehicle_Clean'].str.upper() == 'NOT PICKED!'].copy()
            picked_df = df_data[df_data['Vehicle_Clean'].str.upper() != 'NOT PICKED!'].copy()

            total_locations = df_data['Cont#'].nunique()
            not_picked_unique = not_picked_df['Cont#'].nunique()
            picked_unique = picked_df['Cont#'].nunique()
            pick_percentage = (picked_unique / total_locations * 100) if total_locations > 0 else 0

            # ------------------------------------------------------------------
            # DASHBOARD METRICS
            # ------------------------------------------------------------------
            st.markdown("### 📊 Executive Overview")
            
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Total POI Locations", total_locations)
            m2.metric("Picked Containers", picked_unique, delta=f"{pick_percentage:.1f}% Picked", delta_color="normal")
            m3.metric("Not Picked Containers", not_picked_unique, delta=f"-{100-pick_percentage:.1f}% Missed", delta_color="inverse")
            m4.metric("Total Hit Entries", len(df_data))

            st.info(f"📅 **Report Period:** {time_period} | 🕒 **Generated:** {generated_at}")

            # ------------------------------------------------------------------
            # NOT PICKED VEHICLES / CONTAINERS LIST
            # ------------------------------------------------------------------
            st.markdown("---")
            st.markdown(f"### ⚠️ Not Picked Containers List ({not_picked_unique} Locations)")
            
            # Remove empty columns (#, Time, Proximity, PITB StopPointId) for clean display
            empty_cols_to_remove = ['#', 'Time', 'Proximity', 'PITB StopPointId', 'Vehicle_Clean']
            clean_not_picked_cols = [c for c in not_picked_df.columns if c not in empty_cols_to_remove]
            
            display_not_picked = not_picked_df[clean_not_picked_cols].drop_duplicates(subset=['Cont#']).reset_index(drop=True)
            display_not_picked.index = display_not_picked.index + 1
            
            st.dataframe(
                display_not_picked[['Ser#', 'Cont#', 'Vehicle', 'Latitude', 'Longitude']],
                use_container_width=True,
                height=320
            )

            # ------------------------------------------------------------------
            # EXCEL REPORT GENERATOR (OUTSTANDING STYLING)
            # ------------------------------------------------------------------
            wb = openpyxl.Workbook()
            
            # Styling definitions
            font_header = Font(name='Segoe UI', size=11, bold=True, color='FFFFFF')
            font_title = Font(name='Segoe UI', size=15, bold=True, color='FFFFFF')
            font_sub = Font(name='Segoe UI', size=10, italic=True, color='E2E8F0')
            font_body = Font(name='Segoe UI', size=10, color='1E293B')
            font_alert = Font(name='Segoe UI', size=10, bold=True, color='991B1B')
            
            fill_header = PatternFill(start_color='1E293B', end_color='1E293B', fill_type='solid')
            fill_title = PatternFill(start_color='0F172A', end_color='0F172A', fill_type='solid')
            fill_sub = PatternFill(start_color='334155', end_color='334155', fill_type='solid')
            fill_alt = PatternFill(start_color='F8FAFC', end_color='F8FAFC', fill_type='solid')
            fill_not_picked = PatternFill(start_color='FEE2E2', end_color='FEE2E2', fill_type='solid')
            
            thin_border = Border(
                left=Side(style='thin', color='CBD5E0'),
                right=Side(style='thin', color='CBD5E0'),
                top=Side(style='thin', color='CBD5E0'),
                bottom=Side(style='thin', color='CBD5E0')
            )
            center_align = Alignment(horizontal='center', vertical='center')
            left_align = Alignment(horizontal='left', vertical='center')

            # --------------------------------------------------
            # SHEET 1: NOT PICKED CONTAINERS (EXCLUDING EMPTY COLS)
            # --------------------------------------------------
            ws_not_picked = wb.active
            ws_not_picked.title = "Not Picked Containers"
            
            np_headers = [c for c in headers if c not in ['#', 'Time', 'Proximity', 'PITB StopPointId']]
            
            # Title
            ws_not_picked.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(np_headers))
            cell = ws_not_picked.cell(row=1, column=1, value=f"NOT PICKED CONTAINERS REPORT ({not_picked_unique} LOCATIONS)")
            cell.font = font_title
            cell.fill = fill_title
            cell.alignment = center_align
            ws_not_picked.row_dimensions[1].height = 36

            # Subtitle
            ws_not_picked.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(np_headers))
            cell = ws_not_picked.cell(row=2, column=1, value=f"Period: {time_period} | Generated: {generated_at}")
            cell.font = font_sub
            cell.fill = fill_sub
            cell.alignment = center_align
            ws_not_picked.row_dimensions[2].height = 22

            # Table Header
            ws_not_picked.row_dimensions[4].height = 28
            for col_i, h_text in enumerate(np_headers, 1):
                c = ws_not_picked.cell(row=4, column=col_i, value=h_text)
                c.font = font_header
                c.fill = PatternFill(start_color='991B1B', end_color='991B1B', fill_type='solid') # Red header for attention
                c.alignment = center_align
                c.border = thin_border

            # Data Rows
            row_idx = 5
            for _, r in display_not_picked[np_headers].iterrows():
                ws_not_picked.row_dimensions[row_idx].height = 20
                for col_i, val in enumerate(r, 1):
                    c = ws_not_picked.cell(row=row_idx, column=col_i, value=val)
                    c.alignment = left_align if col_i == 2 else center_align
                    c.border = thin_border
                    c.font = font_alert if col_i == 3 else font_body
                    c.fill = fill_not_picked if col_i == 3 else (fill_alt if row_idx % 2 == 0 else PatternFill(fill_type=None))
                row_idx += 1

            for col in ws_not_picked.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = get_column_letter(col[0].column)
                ws_not_picked.column_dimensions[col_letter].width = max(max_len + 4, 15)

            # --------------------------------------------------
            # SHEET 2: ALL / PICKED HITS (FULL COLUMNS)
            # --------------------------------------------------
            ws_hits = wb.create_sheet(title="Picked Hits")
            
            # Title
            ws_hits.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
            cell = ws_hits.cell(row=1, column=1, value="SUCCESSFUL PICKED POI HITS")
            cell.font = font_title
            cell.fill = fill_title
            cell.alignment = center_align
            ws_hits.row_dimensions[1].height = 36

            # Subtitle
            ws_hits.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(headers))
            cell = ws_hits.cell(row=2, column=1, value=f"Period: {time_period} | Generated: {generated_at}")
            cell.font = font_sub
            cell.fill = fill_sub
            cell.alignment = center_align
            ws_hits.row_dimensions[2].height = 22

            # Table Header
            ws_hits.row_dimensions[4].height = 28
            for col_i, h_text in enumerate(headers, 1):
                c = ws_hits.cell(row=4, column=col_i, value=h_text)
                c.font = font_header
                c.fill = fill_header
                c.alignment = center_align
                c.border = thin_border

            # Data Rows
            row_idx = 5
            for _, r in picked_df[headers].iterrows():
                ws_hits.row_dimensions[row_idx].height = 20
                for col_i, val in enumerate(r, 1):
                    c = ws_hits.cell(row=row_idx, column=col_i, value=val if pd.notna(val) else "")
                    c.alignment = left_align if col_i == 2 else center_align
                    c.border = thin_border
                    c.font = font_body
                    if row_idx % 2 == 0:
                        c.fill = fill_alt
                row_idx += 1

            for col in ws_hits.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = get_column_letter(col[0].column)
                ws_hits.column_dimensions[col_letter].width = max(max_len + 4, 14)

            # --------------------------------------------------
            # DOWNLOAD EXCEL
            # --------------------------------------------------
            excel_buffer = io.BytesIO()
            wb.save(excel_buffer)
            excel_buffer.seek(0)

            st.markdown("---")
            st.download_button(
                label="📥 Download Executive POI Hit Report (.xlsx)",
                data=excel_buffer,
                file_name=f"POI_Hit_Executive_Report_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"❌ Error processing POI Hit Report: {str(e)}")
