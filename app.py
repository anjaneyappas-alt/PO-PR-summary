import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import io

# Page config - Full Wide Layout
st.set_page_config(page_title="PO vs PR Dashboard", page_icon="📊", layout="wide")

# Custom Modern Clean Dashboard CSS
st.markdown("""
    <style>
        /* Hide default Streamlit headers, footers, and menu bars */
        #MainMenu {visibility: hidden;}
        header {visibility: hidden;}
        footer {visibility: hidden;}
        [data-testid="stHeader"] {display: none;}
        
        /* Dashboard Container Styling */
        .main {
            background-color: #0e1117;
        }

        /* Modern Styled HTML Table */
        .clean-table-container {
            width: 100%;
            overflow-x: auto;
            margin-top: 15px;
            border-radius: 10px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.15);
        }

        .clean-table {
            width: 100%;
            border-collapse: separate;
            border-spacing: 0;
            background-color: #1e222d;
            color: #e0e0e0;
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
            font-size: 14px;
            border-radius: 10px;
            overflow: hidden;
        }

        .clean-table th {
            background-color: #2b303c;
            color: #ffffff;
            font-weight: 600;
            padding: 14px 16px;
            text-align: center;
            border-bottom: 2px solid #3b4252;
            white-space: nowrap;
        }

        .clean-table td {
            padding: 12px 16px;
            text-align: center;
            border-bottom: 1px solid #2e3440;
            white-space: nowrap;
        }

        .clean-table tr:hover {
            background-color: #262c38;
        }

        /* Column Specific Text Alignments */
        .clean-table td.vendor-col {
            text-align: left !important;
            font-weight: 500;
            white-space: normal !important;
            min-width: 220px;
        }

        .clean-table td.pr-col {
            white-space: normal !important;
            min-width: 160px;
        }

        /* Status Colors */
        .badge-excess {
            background-color: rgba(40, 167, 69, 0.2);
            color: #2ecc71;
            font-weight: bold;
            padding: 4px 8px;
            border-radius: 4px;
        }

        .badge-short {
            background-color: rgba(220, 53, 69, 0.2);
            color: #e74c3c;
            font-weight: bold;
            padding: 4px 8px;
            border-radius: 4px;
        }

        .total-row td {
            background-color: #f4b084 !important;
            color: #000000 !important;
            font-weight: bold !important;
            font-size: 15px;
            border-top: 2px solid #e08e53;
        }
    </style>
""", unsafe_allow_html=True)

st.title("📊 PO vs PR Dashboard")

# Toggle switch to show/hide upload inputs
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

                # Process PR Data
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
                
                # Sort chronologically from earliest to latest date
                merged['Raw_Date'] = pd.to_datetime(merged['PR_Date'], errors='coerce')
                merged = merged.sort_values(by='Raw_Date', ascending=True).reset_index(drop=True)
                merged['Date'] = merged['Raw_Date'].dt.strftime('%d-%m-%y')
                
                # Clean values
                merged['Vendor Name'] = merged['Vendor_PR'].fillna(merged['Vendor_PO'])
                merged['PO Qty'] = merged['PO_Qty'].fillna(0).astype(int)
                merged['PR Qty'] = merged['PR_Qty'].fillna(0).astype(int)
                
                # Excess & Short
                merged['Diff'] = merged['PR Qty'] - merged['PO Qty']
                merged['Excess'] = merged['Diff'].apply(lambda x: x if x > 0 else 0)
                merged['Short'] = merged['Diff'].apply(lambda x: abs(x) if x < 0 else 0)
                merged['Sl.no'] = range(1, len(merged) + 1)

                # Totals
                total_po = merged['PO Qty'].sum()
                total_pr = merged['PR Qty'].sum()
                total_excess = merged['Excess'].sum()
                total_short = merged['Short'].sum()
                total_fr = (total_pr / total_po * 100) if total_po > 0 else 0

                # Formatted headers
                expected_headers = ["Sl.no", "Date", "PO Number", "PR_no", "Vendor Name", "PO Qty", "PR Qty", "Excess", "Short"]
                final_df = merged[expected_headers].rename(columns={'PO Number': 'PO No', 'PR_no': 'PR No'})
                final_df['PO FR %'] = ((final_df['PR Qty'] / final_df['PO Qty']).fillna(0) * 100).round(0).astype(int).astype(str) + '%'

                # Create Total row
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

                # Cache in Session State
                st.session_state["processed_df"] = final_df
                st.session_state["metrics"] = (len(merged), total_po, total_pr, total_excess, total_short, total_fr)

        except Exception as e:
            st.error(f"Error processing files: {e}")

# Render Clean UI Dashboard
if "processed_df" in st.session_state:
    final_df = st.session_state["processed_df"]
    total_pos, total_po, total_pr, total_excess, total_short, total_fr = st.session_state["metrics"]

    # Top KPI Cards
    st.markdown("### 🎯 Key Metrics")
    kpi1, kpi2, kpi3, kpi4, kpi5, kpi6 = st.columns(6)
    kpi1.metric("Total POs", f"{total_pos:,}")
    kpi2.metric("PO Qty", f"{total_po:,}")
    kpi3.metric("PR Qty", f"{total_pr:,}")
    kpi4.metric("Excess Qty", f"{total_excess:,}")
    kpi5.metric("Short Qty", f"-{total_short:,}")
    kpi6.metric("Fill Rate", f"{total_fr:.1f}%")

    st.markdown("---")

    # Render HTML Clean Table
    html_table = "<div class='clean-table-container'><table class='clean-table'><thead><tr>"
    for col in final_df.columns:
        html_table += f"<th>{col}</th>"
    html_table += "</tr></thead><tbody>"

    for _, row in final_df.iterrows():
        is_total = row['Vendor Name'] == 'Total'
        tr_class = "class='total-row'" if is_total else ""
        html_table += f"<tr {tr_class}>"
        
        for col in final_df.columns:
            val = row[col]
            td_class = ""
            
            if col == "Vendor Name":
                td_class = "class='vendor-col'"
            elif col == "PR No":
                td_class = "class='pr-col'"
            
            # Badges for Excess / Short
            if not is_total and col == "Excess" and isinstance(val, (int, float)) and val > 0:
                val_str = f"<span class='badge-excess'>+{val}</span>"
            elif not is_total and col == "Short" and isinstance(val, (int, float)) and val > 0:
                val_str = f"<span class='badge-short'>-{val}</span>"
            else:
                val_str = str(val)
                
            html_table += f"<td {td_class}>{val_str}</td>"
        html_table += "</tr>"

    html_table += "</tbody></table></div>"
    st.markdown(html_table, unsafe_allow_html=True)

    # --- ULTRA HD FULL DASHBOARD CAPTURE (METRICS + TABLE) ---
    def generate_full_dashboard_hd(df, metrics_tuple):
        total_pos, total_po, total_pr, total_excess, total_short, total_fr = metrics_tuple
        
        # Calculate dynamic figure height based on rows
        fig_height = 3.5 + len(df) * 0.45
        fig = plt.figure(figsize=(20, fig_height), facecolor='#0e1117')
        
        # Grid Layout: Top 15% for KPIs, Bottom 85% for Table
        gs = fig.add_gridspec(2, 1, height_ratios=[1.2, len(df) * 0.45])
        
        # 1. Render Key Metrics Section
        ax_kpi = fig.add_subplot(gs[0])
        ax_kpi.set_facecolor('#0e1117')
        ax_kpi.axis('off')
        
        ax_kpi.text(0.01, 0.8, "🎯 Key Metrics", color='white', fontsize=16, fontweight='bold')
        
        kpis = [
            ("Total POs", f"{total_pos:,}"),
            ("PO Qty", f"{total_po:,}"),
            ("PR Qty", f"{total_pr:,}"),
            ("Excess Qty", f"{total_excess:,}"),
            ("Short Qty", f"-{total_short:,}"),
            ("Fill Rate", f"{total_fr:.1f}%")
        ]
        
        col_width = 1.0 / len(kpis)
        for i, (label, val) in enumerate(kpis):
            x_pos = i * col_width + col_width / 2
            # Metric Box
            rect = plt.Rectangle((i * col_width + 0.01, 0.1), col_width - 0.02, 0.55, 
                                 facecolor='#1e222d', edgecolor='#2e3440', cornercolor='none',
                                 transform=ax_kpi.transAxes, zorder=2)
            ax_kpi.add_patch(rect)
            ax_kpi.text(x_pos, 0.48, label, color='#a0a0a0', fontsize=10, ha='center', va='center')
            ax_kpi.text(x_pos, 0.25, val, color='white', fontsize=14, fontweight='bold', ha='center', va='center')

        # 2. Render Full Data Table
        ax_table = fig.add_subplot(gs[1])
        ax_table.set_facecolor('#0e1117')
        ax_table.axis('off')
        
        table = ax_table.table(
            cellText=df.values,
            colLabels=df.columns,
            cellLoc='center',
            loc='upper center'
        )
        table.auto_set_font_size(False)
        table.set_fontsize(10)
        table.scale(1.2, 1.6)

        # Full visibility column widths
        col_widths = {
            0: 0.04,  # Sl.no
            1: 0.07,  # Date
            2: 0.09,  # PO No
            3: 0.13,  # PR No
            4: 0.27,  # Vendor Name
            5: 0.07,  # PO Qty
            6: 0.07,  # PR Qty
            7: 0.07,  # Excess
            8: 0.07,  # Short
            9: 0.07   # PO FR %
        }
        for (row, col), cell in table.get_celld().items():
            cell.set_width(col_widths.get(col, 0.1))

        excess_col_idx = df.columns.get_loc('Excess')
        short_col_idx = df.columns.get_loc('Short')

        # Lossless Color Styling matching Dashboard Theme
        for (row, col), cell in table.get_celld().items():
            if row == 0:
                cell.set_facecolor('#2b303c')
                cell.set_text_props(color='white', weight='bold')
            elif row == len(df):
                cell.set_facecolor('#f4b084')
                cell.set_text_props(color='black', weight='bold')
            else:
                cell.set_facecolor('#1e222d')
                cell.set_text_props(color='#e0e0e0')
                # Excess Highlight Badge
                if col == excess_col_idx and df.iloc[row - 1]['Excess'] > 0:
                    cell.set_facecolor('#1e3a29')
                    cell.set_text_props(color='#2ecc71', weight='bold')
                # Short Highlight Badge
                elif col == short_col_idx and df.iloc[row - 1]['Short'] > 0:
                    cell.set_facecolor('#3a1e22')
                    cell.set_text_props(color='#e74c3c', weight='bold')

        img_buf = io.BytesIO()
        plt.tight_layout()
        # High-DPI Lossless PNG Output (300 DPI)
        plt.savefig(img_buf, format='png', bbox_inches='tight', dpi=300, facecolor='#0e1117')
        plt.close(fig)
        return img_buf.getvalue()

    # HD Capture Download Button
    st.markdown("---")
    hd_img_bytes = generate_full_dashboard_hd(final_df, (total_pos, total_po, total_pr, total_excess, total_short, total_fr))
    st.download_button(
        label="📸 Capture & Download Full Dashboard (HD PNG)",
        data=hd_img_bytes,
        file_name="PO_PR_Full_Dashboard_HD.png",
        mime="image/png"
    )
