import streamlit as st
import pandas as pd
import numpy as np
import io

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Page config - Standard Excel Wide Layout
st.set_page_config(page_title="PO vs PR Summary", page_icon="📊", layout="wide")

# Excel Grid Styling + High-Contrast KPI Cards
st.markdown("""
    <style>
        /* Hide default Streamlit headers and footers */
        #MainMenu {visibility: hidden;}
        header {visibility: hidden;}
        footer {visibility: hidden;}
        [data-testid="stHeader"] {display: none;}
        
        /* White Main Page Background */
        .main {
            background-color: #ffffff !important;
        }

        /* High Contrast Light Cards for Top Metrics */
        [data-testid="stMetric"] {
            background-color: #f8f9fa !important;
            border: 1px solid #dee2e6 !important;
            padding: 12px 16px !important;
            border-radius: 8px !important;
            box-shadow: 0 1px 3px rgba(0,0,0,0.05) !important;
        }

        [data-testid="stMetricValue"] {
            font-size: 22px !important;
            font-weight: 700 !important;
            color: #111111 !important;
        }
        
        [data-testid="stMetricLabel"] {
            font-size: 13px !important;
            font-weight: 600 !important;
            color: #495057 !important;
        }

        .stTable {
            background-color: #ffffff !important;
            font-family: "Segoe UI", Arial, sans-serif !important;
            font-size: 13px !important;
            color: #000000 !important;
        }
        
        .stTable table {
            border-collapse: collapse !important;
            width: 100% !important;
            border: 1px solid #d9d9d9 !important;
        }
        
        /* Excel Table Headers */
        .stTable th {
            background-color: #e6e6e6 !important;
            color: #000000 !important;
            font-weight: bold !important;
            text-align: center !important;
            border: 1px solid #bfbfbf !important;
            padding: 6px 10px !important;
            white-space: nowrap !important;
        }
        
        /* Excel Data Cells & Gridlines */
        .stTable td {
            border: 1px solid #d9d9d9 !important;
            padding: 6px 10px !important;
            color: #000000 !important;
            white-space: nowrap !important;
        }
        
        /* Zebra Striping */
        .stTable tr:nth-child(even) {
            background-color: #f9f9f9 !important;
        }
    </style>
""", unsafe_allow_html=True)

st.title("📊 PO vs PR Excel View")

# Control Panel: Toggle panel visibility & Select Lookback Window
control_col1, control_col2 = st.columns([1, 2])
with control_col1:
    show_uploaders = st.toggle("🎚 Show Upload Panel", value=True)
with control_col2:
    selected_window = st.radio(
        "⏱ Select PR Lookback Window:",
        options=["Last 15 Hours", "Last 24 Hours"],
        horizontal=True,
        index=0
    )

# Determine hours numerical value
hours_window = 15 if selected_window == "Last 15 Hours" else 24

if show_uploaders:
    col1, col2 = st.columns(2)
    with col1:
        po_file = st.file_uploader("Upload PO Monthly Data (Excel)", type=['xlsx', 'xls'])
    with col2:
        pr_file = st.file_uploader("Upload PR Monthly Data (Excel)", type=['xlsx', 'xls'])
        
    if po_file and pr_file:
        try:
            with st.spinner("Processing monthly data..."):
                # Load raw datasets
                df_po = pd.read_excel(po_file)
                df_pr = pd.read_excel(pr_file)

                # Strip whitespace from headers
                df_po.columns = df_po.columns.str.strip()
                df_pr.columns = df_pr.columns.str.strip()

                # Dynamic Column Identification
                po_num_col = 'Purchase Order Number' if 'Purchase Order Number' in df_po.columns else ('PO Number' if 'PO Number' in df_po.columns else 'Purchase Order ID')
                po_qty_col = 'QuantityOrdered' if 'QuantityOrdered' in df_po.columns else ('TotalQuantityOrdered' if 'TotalQuantityOrdered' in df_po.columns else 'Quantity')

                pr_no_col = 'Receive Number' if 'Receive Number' in df_pr.columns else ('PR Number' if 'PR Number' in df_pr.columns else 'Purchase Receive ID')
                qty_pr_col = 'Quantity Received' if 'Quantity Received' in df_pr.columns else 'Quantity'
                
                # Prioritize CreatedTime for precision timestamp
                if 'CreatedTime' in df_pr.columns:
                    time_col = 'CreatedTime'
                elif 'Receive Date' in df_pr.columns:
                    time_col = 'Receive Date'
                else:
                    time_col = 'PR Date'

                # Clean PR Data
                df_pr_clean = df_pr.dropna(subset=['PO Number']).copy()
                df_pr_clean['Clean_PR_Qty'] = pd.to_numeric(df_pr_clean[qty_pr_col], errors='coerce').fillna(0)
                
                # Parse Created Timestamp
                df_pr_clean['DT'] = pd.to_datetime(df_pr_clean[time_col], errors='coerce')
                
                # Dynamic Lookback Calculation relative to the LATEST created timestamp in the PR file
                latest_pr_time = df_pr_clean['DT'].max()
                
                if pd.notna(latest_pr_time):
                    cutoff_time = latest_pr_time - pd.Timedelta(hours=hours_window)
                    # Today's PRs: Created within the selected lookback window (15h or 24h)
                    df_today = df_pr_clean[df_pr_clean['DT'] >= cutoff_time]
                    # PRMTD: Created before the selected lookback window
                    df_prior = df_pr_clean[df_pr_clean['DT'] < cutoff_time]
                else:
                    df_today = df_pr_clean
                    df_prior = pd.DataFrame(columns=df_pr_clean.columns)

                # Group 1: Today's PR Data (Within lookback window)
                today_summary = df_today.groupby('PO Number').agg(
                    PR_no=(pr_no_col, lambda x: " & ".join(sorted(x.dropna().astype(str).unique()))),
                    PR_Qty=('Clean_PR_Qty', 'sum'),
                    PR_Date=('DT', 'max'),
                    Vendor_PR=('Vendor Name', 'first')
                ).reset_index()

                # Group 2: PRMTD Data (Created prior to selected hours for the same PO)
                prior_summary = df_prior.groupby('PO Number').agg(
                    PRMTD=('Clean_PR_Qty', 'sum')
                ).reset_index()

                # Group 3: PO Monthly Data
                df_po_clean = df_po.dropna(subset=[po_num_col]).copy()
                df_po_clean['Clean_PO_Qty'] = pd.to_numeric(df_po_clean[po_qty_col], errors='coerce').fillna(0)
                
                po_summary = df_po_clean.groupby(po_num_col).agg(
                    PO_Qty=('Clean_PO_Qty', 'sum'),
                    Vendor_PO=('Vendor Name', 'first')
                ).reset_index()

                # Merge Datasets
                merged = pd.merge(today_summary, po_summary, left_on='PO Number', right_on=po_num_col, how='left')
                merged = pd.merge(merged, prior_summary, on='PO Number', how='left')

                # Sort chronologically by date
                merged = merged.sort_values(by='PR_Date', ascending=True).reset_index(drop=True)

                # Format Display Columns
                merged['Date'] = merged['PR_Date'].dt.strftime('%d-%m-%y').fillna('')
                merged['Vendor Name'] = merged['Vendor_PR'].fillna(merged['Vendor_PO']).fillna('')
                
                merged['PO Qty'] = pd.to_numeric(merged['PO_Qty'], errors='coerce').fillna(0).astype(int)
                merged['PR Qty'] = pd.to_numeric(merged['PR_Qty'], errors='coerce').fillna(0).astype(int)
                merged['PRMTD'] = pd.to_numeric(merged['PRMTD'], errors='coerce').fillna(0).astype(int)
                
                # Total Receipts = Today PR Qty + PRMTD (Prior)
                merged['Total Received'] = merged['PR Qty'] + merged['PRMTD']
                
                # Pending / Unfulfilled Qty Calculation
                merged['Pending Qty'] = (merged['PO Qty'] - merged['Total Received']).apply(lambda x: int(x) if x > 0 else 0)
                
                # Re-assign Sl.no in strictly sorted order
                merged['Sl.no'] = range(1, len(merged) + 1)

                # Overall KPI Metrics
                total_pos_cnt = len(merged)
                total_po = int(merged['PO Qty'].sum())
                total_pr = int(merged['PR Qty'].sum())
                total_prmtd = int(merged['PRMTD'].sum())
                total_pending = int(merged['Pending Qty'].sum())
                total_fr_val = ((total_pr + total_prmtd) / total_po * 100) if total_po > 0 else 0

                # Column Formatting & Headers
                expected_headers = ["Sl.no", "Date", "PO Number", "PR_no", "Vendor Name", "PO Qty", "PR Qty", "PRMTD", "Pending Qty"]
                final_df = merged[expected_headers].rename(columns={'PO Number': 'PO No', 'PR_no': 'PR No', 'PR Qty': 'Today PR Qty', 'Pending Qty': 'Pending Qty'})
                
                # Fill Rate %
                fr_numeric = np.where(final_df['PO Qty'] > 0, ((final_df['Today PR Qty'] + final_df['PRMTD']) / final_df['PO Qty']) * 100, 0)
                final_df['PO FR %'] = np.round(fr_numeric).astype(int).astype(str) + '%'

                # Summary Row at Bottom
                total_row = pd.DataFrame([{
                    "Sl.no": "",
                    "Date": "",
                    "PO No": "",
                    "PR No": "",
                    "Vendor Name": "Total",
                    "PO Qty": total_po,
                    "Today PR Qty": total_pr,
                    "PRMTD": total_prmtd,
                    "Pending Qty": total_pending,
                    "PO FR %": f"{int(round(total_fr_val))}%"
                }])

                display_df = pd.concat([final_df, total_row], ignore_index=True)

                st.session_state["processed_df"] = display_df
                st.session_state["kpi_metrics"] = (total_pos_cnt, total_po, total_pr, total_prmtd, total_pending, total_fr_val)
                st.session_state["anchor_time"] = latest_pr_time
                st.session_state["window_hours"] = hours_window

        except Exception as e:
            st.error(f"Error processing files: {e}")

# Render UI
if "processed_df" in st.session_state:
    display_df = st.session_state["processed_df"]
    total_pos_cnt, total_po, total_pr, total_prmtd, total_pending, total_fr_val = st.session_state["kpi_metrics"]
    anchor_time = st.session_state.get("anchor_time", None)
    active_window = st.session_state.get("window_hours", 15)

    # Top Summary Metrics Header
    st.markdown("### 🎯 Total Summary")
    if anchor_time:
        st.markdown(f"🕒 **Last PR Created Time:** `{anchor_time.strftime('%d-%b-%Y %H:%M:%S')}` | **{active_window}-Hour Window:** `{ (anchor_time - pd.Timedelta(hours=active_window)).strftime('%d-%b-%Y %H:%M:%S') }` onwards")

    kpi1, kpi2, kpi3, kpi4, kpi5, kpi6 = st.columns(6)
    kpi1.metric("Total POs", f"{total_pos_cnt:,}")
    kpi2.metric("PO Qty", f"{total_po:,}")
    kpi3.metric(f"Today PR Qty ({active_window}h)", f"{total_pr:,}")
    kpi4.metric("PRMTD Qty", f"{total_prmtd:,}")
    kpi5.metric("Pending Qty", f"{total_pending:,}")
    kpi6.metric("Fill Rate", f"{total_fr_val:.1f}%")

    st.markdown("---")

    # Styling function for UI Display Table
    def highlight_excel_cells(df):
        styles = pd.DataFrame('', index=df.index, columns=df.columns)
        for idx, row in df.iterrows():
            if str(row['Vendor Name']) == 'Total':
                styles.loc[idx, :] = 'background-color: #f4b084; font-weight: bold; color: black; border-top: 2px solid black; border-bottom: 2px double black'
            else:
                try:
                    if float(row['Pending Qty']) > 0:
                        styles.loc[idx, 'Pending Qty'] = 'background-color: #fff3cd; color: #856404; font-weight: bold;'
                except (ValueError, TypeError):
                    pass
                    
        return styles

    styled_df = display_df.style.apply(highlight_excel_cells, axis=None)

    # Render Streamlit Table
    st.table(styled_df)

    # --- FULLY STYLED OPENPYXL EXCEL GENERATOR ---
    def generate_excel_file(df):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "PO vs PR Summary"
        ws.views.sheetView[0].showGridLines = True

        # Color definitions
        header_fill = PatternFill(start_color="E6E6E6", end_color="E6E6E6", fill_type="solid")
        zebra_fill = PatternFill(start_color="F9F9F9", end_color="F9F9F9", fill_type="solid")
        white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
        yellow_pending_fill = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")
        orange_total_fill = PatternFill(start_color="F4B084", end_color="F4B084", fill_type="solid")

        # Borders
        thin_side = Side(border_style="thin", color="D9D9D9")
        thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
        
        total_top_side = Side(border_style="thin", color="000000")
        total_bottom_side = Side(border_style="double", color="000000")
        total_border = Border(left=thin_side, right=thin_side, top=total_top_side, bottom=total_bottom_side)

        # Fonts
        font_header = Font(name="Segoe UI", size=11, bold=True, color="000000")
        font_regular = Font(name="Segoe UI", size=11, bold=False, color="000000")
        font_pending = Font(name="Segoe UI", size=11, bold=True, color="856404")
        font_total = Font(name="Segoe UI", size=11, bold=True, color="000000")

        # Alignments
        align_center = Alignment(horizontal="center", vertical="center")
        align_left = Alignment(horizontal="left", vertical="center")
        align_right = Alignment(horizontal="right", vertical="center")

        headers = list(df.columns)
        ws.append(headers)

        # Style Header Row
        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = font_header
            cell.alignment = align_center
            cell.border = Border(left=Side(border_style="thin", color="BFBFBF"),
                                 right=Side(border_style="thin", color="BFBFBF"),
                                 top=Side(border_style="thin", color="BFBFBF"),
                                 bottom=Side(border_style="thin", color="BFBFBF"))

        pending_col_idx = headers.index('Pending Qty') + 1 if 'Pending Qty' in headers else None

        # Append Data Rows & Apply Styles
        for r_idx, row_data in enumerate(df.values, start=2):
            ws.append(list(row_data))
            is_total_row = (str(row_data[headers.index('Vendor Name')]) == 'Total') if 'Vendor Name' in headers else False

            for col_idx in range(1, len(headers) + 1):
                cell = ws.cell(row=r_idx, column=col_idx)
                val = cell.value

                if is_total_row:
                    cell.fill = orange_total_fill
                    cell.font = font_total
                    cell.border = total_border
                else:
                    cell.fill = zebra_fill if (r_idx % 2 == 0) else white_fill
                    cell.font = font_regular
                    cell.border = thin_border

                    # Apply Yellow Highlight to Pending Qty > 0
                    if col_idx == pending_col_idx:
                        try:
                            if float(val) > 0:
                                cell.fill = yellow_pending_fill
                                cell.font = font_pending
                        except (ValueError, TypeError):
                            pass

                # Alignments
                if headers[col_idx - 1] in ['Sl.no', 'Date', 'PO No', 'PR No', 'PO FR %']:
                    cell.alignment = align_center
                elif headers[col_idx - 1] in ['Vendor Name']:
                    cell.alignment = align_left
                else:
                    cell.alignment = align_right

        # Adjust Column Widths Dynamically
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

        output = io.BytesIO()
        wb.save(output)
        return output.getvalue()

    st.markdown("---")
    excel_bytes = generate_excel_file(display_df)
    st.download_button(
        label="📥 Download Summary as Excel (.xlsx)",
        data=excel_bytes,
        file_name=f"PO_PR_Summary_{active_window}h.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
