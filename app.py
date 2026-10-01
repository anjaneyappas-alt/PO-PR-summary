import streamlit as st
import pandas as pd
import numpy as np
import io

# Page config - Standard Excel Wide Layout
st.set_page_config(page_title="PO vs PR Summary", page_icon="📊", layout="wide")

# Excel Grid Styling
st.markdown("""
    <style>
        /* Hide default Streamlit headers and footers */
        #MainMenu {visibility: hidden;}
        header {visibility: hidden;}
        footer {visibility: hidden;}
        [data-testid="stHeader"] {display: none;}
        
        /* White Background Container */
        .main {
            background-color: #ffffff !important;
        }

        /* KPI Top Metric Cards */
        [data-testid="stMetricValue"] {
            font-size: 20px !important;
            font-weight: bold !important;
            color: #111111 !important;
        }
        
        [data-testid="stMetricLabel"] {
            font-size: 13px !important;
            color: #555555 !important;
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

# Toggle panel to hide uploaders
show_uploaders = st.toggle("🎚️️ Show Upload Panel", value=True)

if show_uploaders:
    uploaded_file = st.file_uploader("Upload Excel File (with 'po', 'pr', and 'PRMTD' sheets)", type=['xlsx', 'xls'])
        
    if uploaded_file:
        try:
            with st.spinner("Processing sheets safely..."):
                xls = pd.ExcelFile(uploaded_file)
                
                # 1. Handle whitespace/case issues in sheet names dynamically
                sheet_map = {str(name).strip().lower(): name for name in xls.sheet_names}

                po_sheet = sheet_map.get('po', xls.sheet_names[1] if len(xls.sheet_names) > 1 else xls.sheet_names[0])
                pr_sheet = sheet_map.get('pr', xls.sheet_names[2] if len(xls.sheet_names) > 2 else xls.sheet_names[0])
                prmtd_sheet = sheet_map.get('prmtd', xls.sheet_names[3] if len(xls.sheet_names) > 3 else None)

                # Load raw data
                df_po = pd.read_excel(uploaded_file, sheet_name=po_sheet)
                df_pr = pd.read_excel(uploaded_file, sheet_name=pr_sheet)
                df_prmtd = pd.read_excel(uploaded_file, sheet_name=prmtd_sheet) if prmtd_sheet else pd.DataFrame()

                # Clean column names
                df_po.columns = df_po.columns.str.strip()
                df_pr.columns = df_pr.columns.str.strip()
                if not df_prmtd.empty:
                    df_prmtd.columns = df_prmtd.columns.str.strip()

                # Detect columns
                pr_no_col = 'Receive Number' if 'Receive Number' in df_pr.columns else 'PR Number'
                date_col = 'Receive Date' if 'Receive Date' in df_pr.columns else 'PR Date'
                qty_pr_col = 'Quantity Received' if 'Quantity Received' in df_pr.columns else 'Quantity'

                # Clean PR Data & drop invalid PO Numbers
                df_pr_clean = df_pr.dropna(subset=['PO Number']).copy()
                df_pr_clean['Clean_PR_Qty'] = pd.to_numeric(df_pr_clean[qty_pr_col], errors='coerce').fillna(0)
                df_pr_clean['DT'] = pd.to_datetime(df_pr_clean[date_col], errors='coerce')

                # Today's PR Summary
                today_summary = df_pr_clean.groupby('PO Number').agg(
                    PR_no=(pr_no_col, lambda x: " & ".join(sorted(x.dropna().astype(str).unique()))),
                    PR_Qty=('Clean_PR_Qty', 'sum'),
                    PR_Date=('DT', 'max'),
                    Vendor_PR=('Vendor Name', 'first')
                ).reset_index()

                # PRMTD (Month to date) Summary
                if not df_prmtd.empty and 'PO Number' in df_prmtd.columns:
                    df_prmtd_clean = df_prmtd.dropna(subset=['PO Number']).copy()
                    prmtd_qty_col = 'Quantity Received' if 'Quantity Received' in df_prmtd_clean.columns else 'Quantity'
                    df_prmtd_clean['Clean_PRMTD_Qty'] = pd.to_numeric(df_prmtd_clean[prmtd_qty_col], errors='coerce').fillna(0)
                    
                    prior_summary = df_prmtd_clean.groupby('PO Number').agg(
                        PRMTD=('Clean_PRMTD_Qty', 'sum')
                    ).reset_index()
                else:
                    prior_summary = pd.DataFrame(columns=['PO Number', 'PRMTD'])

                # PO Data processing
                po_num_col = 'Purchase Order Number' if 'Purchase Order Number' in df_po.columns else 'PO Number'
                po_qty_col = 'QuantityOrdered' if 'QuantityOrdered' in df_po.columns else 'Quantity'
                
                df_po_clean = df_po.dropna(subset=[po_num_col]).copy()
                df_po_clean['Clean_PO_Qty'] = pd.to_numeric(df_po_clean[po_qty_col], errors='coerce').fillna(0)
                
                po_summary = df_po_clean.groupby(po_num_col).agg(
                    PO_Qty=('Clean_PO_Qty', 'sum'),
                    Vendor_PO=('Vendor Name', 'first')
                ).reset_index()

                # Merge Datasets
                merged = pd.merge(today_summary, po_summary, left_on='PO Number', right_on=po_num_col, how='left')
                merged = pd.merge(merged, prior_summary, on='PO Number', how='left')

                # Clean & Format text/dates
                merged['Date'] = merged['PR_Date'].dt.strftime('%d-%m-%y').fillna('')
                merged['Vendor Name'] = merged['Vendor_PR'].fillna(merged['Vendor_PO']).fillna('')
                
                # Coerce quantities to safe integer types
                merged['PO Qty'] = pd.to_numeric(merged['PO_Qty'], errors='coerce').fillna(0).astype(int)
                merged['PR Qty'] = pd.to_numeric(merged['PR_Qty'], errors='coerce').fillna(0).astype(int)
                merged['PRMTD'] = pd.to_numeric(merged['PRMTD'], errors='coerce').fillna(0).astype(int)
                
                # Excess & Short calculations based on PRMTD
                merged['Diff'] = merged['PRMTD'] - merged['PO Qty']
                merged['Excess'] = merged['Diff'].apply(lambda x: int(x) if x > 0 else 0)
                merged['Short'] = merged['Diff'].apply(lambda x: int(x) if x < 0 else 0)
                merged['Sl.no'] = range(1, len(merged) + 1)

                # Overall Totals
                total_pos_cnt = len(merged)
                total_po = int(merged['PO Qty'].sum())
                total_pr = int(merged['PR Qty'].sum())
                total_prmtd = int(merged['PRMTD'].sum())
                total_excess = int(merged['Excess'].sum())
                total_short = int(merged['Short'].sum())
                total_fr_val = (total_prmtd / total_po * 100) if total_po > 0 else 0

                # Formatted headers
                expected_headers = ["Sl.no", "Date", "PO Number", "PR_no", "Vendor Name", "PO Qty", "PRMTD", "PR Qty", "Excess", "Short"]
                final_df = merged[expected_headers].rename(columns={'PO Number': 'PO No', 'PR_no': 'PR No', 'PRMTD': 'PRMTD (MTD)'})
                
                fr_numeric = np.where(final_df['PO Qty'] > 0, (final_df['PRMTD (MTD)'] / final_df['PO Qty']) * 100, 0)
                final_df['PO FR %'] = np.round(fr_numeric).astype(int).astype(str) + '%'

                # Total summary row at bottom
                total_row = pd.DataFrame([{
                    "Sl.no": "",
                    "Date": "",
                    "PO No": "",
                    "PR No": "",
                    "Vendor Name": "Total",
                    "PO Qty": total_po,
                    "PRMTD (MTD)": total_prmtd,
                    "PR Qty": total_pr,
                    "Excess": total_excess,
                    "Short": total_short,
                    "PO FR %": f"{int(round(total_fr_val))}%"
                }])

                display_df = pd.concat([final_df, total_row], ignore_index=True)

                st.session_state["processed_df"] = display_df
                st.session_state["kpi_metrics"] = (total_pos_cnt, total_po, total_pr, total_prmtd, total_excess, total_short, total_fr_val)

        except Exception as e:
            st.error(f"Error processing file: {e}")

# Render UI
if "processed_df" in st.session_state:
    display_df = st.session_state["processed_df"]
    total_pos_cnt, total_po, total_pr, total_prmtd, total_excess, total_short, total_fr_val = st.session_state["kpi_metrics"]

    # Top Metric Header Cards
    st.markdown("### 🎯 Total Summary")
    kpi1, kpi2, kpi3, kpi4, kpi5, kpi6, kpi7 = st.columns(7)
    kpi1.metric("Total POs", f"{total_pos_cnt:,}")
    kpi2.metric("PO Qty", f"{total_po:,}")
    kpi3.metric("PRMTD (MTD)", f"{total_prmtd:,}")
    kpi4.metric("Today PR Qty", f"{total_pr:,}")
    kpi5.metric("Excess Qty", f"{total_excess:,}")
    kpi6.metric("Short Qty", f"{total_short:,}")
    kpi7.metric("Fill Rate", f"{total_fr_val:.1f}%")

    st.markdown("---")

    # Safe Styling function ignoring empty strings and invalid values
    def highlight_excel_cells(df):
        styles = pd.DataFrame('', index=df.index, columns=df.columns)
        for idx, row in df.iterrows():
            if str(row['Vendor Name']) == 'Total':
                styles.loc[idx, :] = 'background-color: #f4b084; font-weight: bold; color: black; border-top: 2px solid black; border-bottom: 2px double black'
            else:
                try:
                    if float(row['Excess']) > 0:
                        styles.loc[idx, 'Excess'] = 'background-color: #d4edda; color: #155724; font-weight: bold;'
                except (ValueError, TypeError):
                    pass

                try:
                    if float(row['Short']) < 0:
                        styles.loc[idx, 'Short'] = 'background-color: #f8d7da; color: #721c24; font-weight: bold;'
                except (ValueError, TypeError):
                    pass
                    
        return styles

    styled_df = display_df.style.apply(highlight_excel_cells, axis=None)

    # Display table
    st.table(styled_df)

    # Excel Download
    def generate_excel_file(df):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='PO vs PR Summary')
        return output.getvalue()

    st.markdown("---")
    excel_bytes = generate_excel_file(display_df)
    st.download_button(
        label="📥 Download Summary as Excel (.xlsx)",
        data=excel_bytes,
        file_name="PO_PR_Summary.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
