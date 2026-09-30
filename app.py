import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import io

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
            white-space: nowrap !important;
        }
        
        /* Excel Data Cells & Gridlines */
        .stTable td {
            border: 1px solid #d9d9d9 !important;
            padding: 4px 8px !important;
            color: #000000 !important;
            white-space: nowrap !important;
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
                
                # Calculate separate Excess and Short columns
                merged['Diff'] = merged['PR Qty'] - merged['PO Qty']
                merged['Excess'] = merged['Diff'].apply(lambda x: x if x > 0 else 0)
                merged['Short'] = merged['Diff'].apply(lambda x: abs(x) if x < 0 else 0)
                
                merged['Sl.no'] = range(1, len(merged) + 1)

                # Overall Calculations
                total_po = merged['PO Qty'].sum()
                total_pr = merged['PR Qty'].sum()
                total_excess = merged['Excess'].sum()
                total_short = merged['Short'].sum()
                total_fr = (total_pr / total_po * 100) if total_po > 0 else 0

                # Column Ordering
                expected_headers = ["Sl.no", "Date", "PO Number", "PR_no", "Vendor Name", "PO Qty", "PR Qty", "Excess", "Short"]
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
                    "Excess": total_excess,
                    "Short": total_short,
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

    # Excel-style Highlight for Total Row and Conditional Coloring for Excess / Short
    def apply_custom_styles(row):
        styles = [''] * len(row)
        if row['Vendor Name'] == 'Total':
            return ['background-color: #f4b084; font-weight: bold; color: black; border-top: 2px solid black; border-bottom: 2px double black'] * len(row)
        
        # Excess column highlight (Light Green)
        if row['Excess'] > 0:
            excess_idx = final_df.columns.get_loc('Excess')
            styles[excess_idx] = 'background-color: #d4edda; color: #155724; font-weight: bold;'
            
        # Short column highlight (Light Red)
        if row['Short'] > 0:
            short_idx = final_df.columns.get_loc('Short')
            styles[short_idx] = 'background-color: #f8d7da; color: #721c24; font-weight: bold;'
            
        return styles

    styled_df = final_df.style.apply(apply_custom_styles, axis=1)

    # Render as native HTML Excel Table
    st.table(styled_df)

    # --- FUNCTION TO GENERATE HIGH-RES PNG WITH WIDER COLUMNS AND CONDITIONAL COLORING ---
    def generate_summary_image(df):
        # Increased figure width to 18 to ensure full visibility for Vendor Name
        fig, ax = plt.subplots(figsize=(18, len(df) * 0.35 + 1.5))
        ax.axis('off')
        
        # Draw Matplotlib table matching Excel layout
        table = ax.table(
            cellText=df.values,
            colLabels=df.columns,
            cellLoc='center',
            loc='center'
        )
        table.auto_set_font_size(False)
        table.set_fontsize(9)
        table.scale(1.2, 1.4)

        # Explicitly set column widths to prevent truncation of Vendor Name
        col_widths = {
            0: 0.04,  # Sl.no
            1: 0.08,  # Date
            2: 0.10,  # PO No
            3: 0.14,  # PR No
            4: 0.24,  # Vendor Name (Wide enough for long vendor names)
            5: 0.08,  # PO Qty
            6: 0.08,  # PR Qty
            7: 0.08,  # Excess
            8: 0.08,  # Short
            9: 0.08   # PO FR %
        }
        for (row, col), cell in table.get_celld().items():
            cell.set_width(col_widths.get(col, 0.1))

        excess_col_idx = df.columns.get_loc('Excess')
        short_col_idx = df.columns.get_loc('Short')

        # Color headers, conditional rows, and Total row in image
        for (row, col), cell in table.get_celld().items():
            if row == 0:
                cell.set_facecolor('#e6e6e6')
                cell.set_text_props(color='black', weight='bold')
            elif row == len(df):
                cell.set_facecolor('#f4b084')
                cell.set_text_props(color='black', weight='bold')
            else:
                # Excess Highlight (Light Green)
                if col == excess_col_idx and df.iloc[row - 1]['Excess'] > 0:
                    cell.set_facecolor('#d4edda')
                    cell.set_text_props(color='#155724', weight='bold')
                # Short Highlight (Light Red)
                elif col == short_col_idx and df.iloc[row - 1]['Short'] > 0:
                    cell.set_facecolor('#f8d7da')
                    cell.set_text_props(color='#721c24', weight='bold')

        img_buf = io.BytesIO()
        plt.savefig(img_buf, format='png', bbox_inches='tight', dpi=200)
        plt.close(fig)
        return img_buf.getvalue()

    # Image Download Button at the bottom
    st.markdown("---")
    img_bytes = generate_summary_image(final_df)
    st.download_button(
        label="📸 Download Summary as PNG Image",
        data=img_bytes,
        file_name="PO_PR_Summary.png",
        mime="image/png"
    )
