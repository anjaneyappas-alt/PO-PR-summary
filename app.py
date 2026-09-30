import streamlit as st
import pandas as pd

# Standard page config with wide layout
st.set_page_config(page_title="PO vs PR Summary", page_icon="📊", layout="wide")

# Excel Grid Styling (Light Theme, Monospace Font, Borders, Compact Padding)
st.markdown("""
    <style>
        /* Hide default Streamlit headers, footers, and menu bars */
        #MainMenu {visibility: hidden;}
        header {visibility: hidden;}
        footer {visibility: hidden;}
        [data-testid="stHeader"] {display: none;}
        
        /* Force Excel Worksheet Aesthetics */
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
        }
        
        /* Excel Data Cells & Gridlines */
        .stTable td {
            border: 1px solid #d9d9d9 !important;
            padding: 4px 8px !important;
            color: #000000 !important;
        }
        
        /* Zebra Striping (Light Excel rows) */
        .stTable tr:nth-child(even) {
            background-color: #f9f9f9 !important;
        }
    </style>
""", unsafe_allow_html=True)

st.title("📊 PO vs PR Excel View")

# Toggle panel to hide uploaders when screenshotting
show_uploaders = st.toggle("🎚️ Show Upload Panel", value=True)

if show_uploaders:
    col1, col2 = st.columns(2)
    with col1:
        po_file = st.file_uploader("Upload PO Data (Excel)", type=['xlsx', 'xls'])
    with col2:
        pr_file = st.file_uploader("Upload PR Data (Excel)", type=['xlsx', 'xls'])
        
    if po_file and pr_file:
        try:
            with st.spinner("Processing files..."):
                # Load raw data
                df_po = pd.read_excel(po_file)
                df_pr = pd.read_excel(pr_file)

                # Clean column names
                df_po.columns = df_po.columns.str.strip()
                df_pr.columns = df_pr.columns.str.strip()

                # Process PR Data (Concatenates multiple PRs using ' & ')
                df_pr_clean = df_pr.dropna(subset=['PO Number']).copy()
                pr_no_col = 'Receive Number' if 'Receive Number' in df_pr_clean.columns else 'PR Number'
                
                pr_summary = df_pr_clean.groupby('PO Number').agg(
                    PR_no=(pr_no_col, lambda x: " & ".join(sorted(x.dropna().astype(str).unique()))),
                    PR_Qty=('Quantity Received', 'sum'),
                    PR_Date=('Receive Date', 'first'),
                    Vendor_PR=('Vendor Name', 'first')
                ).reset_index()

                # Process PO Data
                df_po_clean = df_po.dropna(subset=['Purchase Order Number']).copy()
                po_summary = df_po_clean.groupby('Purchase Order Number').agg(
                    PO_Qty=('QuantityOrdered', 'sum'),
                    Vendor_PO=('Vendor Name', 'first')
                ).reset_index()

                # Merge Data
                merged = pd.merge(pr_summary, po_summary, left_on='PO Number', right_on='Purchase Order Number', how='left')
                
                # Sort chronologically from earliest date to latest date
                merged['Raw_Date'] = pd.to_datetime(merged['PR_Date'], errors='coerce')
                merged = merged.sort_values(by='Raw_Date', ascending=True).reset_index(drop=True)
                merged['Date'] = merged['Raw_Date'].dt.strftime('%d-%m-%y')
                
                # Clean values
                merged['Vendor Name'] = merged['Vendor_PR'].fillna(merged['Vendor_PO'])
                merged['PO Qty'] = merged['PO_Qty'].fillna(0).astype(int)
                merged['PR Qty'] = merged['PR_Qty'].fillna(0).astype(int)
                merged['Excess / Short'] = merged['PR Qty'] - merged['PO Qty']
                merged['Sl.no'] = range(1, len(merged) + 1)

                # Overall Calculations
                total_po = merged['PO Qty'].sum()
                total_pr = merged['PR Qty'].sum()
                total_diff = merged['Excess / Short'].sum()
                total_fr = (total_pr / total_po * 100) if total_po > 0 else 0

                # Column Ordering
                expected_headers = ["Sl.no", "Date", "PO Number", "PR_no", "Vendor Name", "PO Qty", "PR Qty", "Excess / Short"]
                final_df = merged[expected_headers].rename(columns={'PO Number': 'PO No', 'PR_no': 'PR No'})

                # Format Fill Rate %
                final_df['PO FR %'] = ((final_df['PR Qty'] / final_df['PO Qty']).fillna(0) * 100).round(0).astype(int).astype(str) + '%'

                # Excel Total Row at bottom
                total_row = pd.DataFrame([{
                    "Sl.no": "",
                    "Date": "",
                    "PO No": "",
                    "PR No": "",
                    "Vendor Name": "Total",
                    "PO Qty": total_po,
                    "PR Qty": total_pr,
                    "Excess / Short": total_diff,
                    "PO FR %": f"{int(round(total_fr))}%"
                }])

                # Append Total Row
                final_df = pd.concat([final_df, total_row], ignore_index=True)

                # Cache results in session state
                st.session_state["processed_df"] = final_df

        except Exception as e:
            st.error(f"Error processing files: {e}")

# Display Excel Worksheet directly
if "processed_df" in st.session_state:
    final_df = st.session_state["processed_df"]

    # Excel-style Highlight for Total Row
    def highlight_total_row(row):
        if row['Vendor Name'] == 'Total':
            return ['background-color: #f4b084; font-weight: bold; color: black; border-top: 2px solid black; border-bottom: 2px double black'] * len(row)
        return [''] * len(row)

    styled_df = final_df.style.apply(highlight_total_row, axis=1)

    # Render as native HTML Excel Table
    st.table(styled_df)
