import streamlit as st
import pandas as pd

# 1. Page Configuration
st.set_page_config(page_title="PO vs PR Dashboard", page_icon="📊", layout="wide")

# 2. Header Section
st.title("📊 Automated PO vs PR Dashboard")
st.markdown("Upload your daily extracts below to generate a real-time reconciliation dashboard.")
st.markdown("---")

# 3. File Uploaders
col1, col2 = st.columns(2)
with col1:
    po_file = st.file_uploader("📥 1. Upload PO Data (Excel)", type=['xlsx', 'xls'])
with col2:
    pr_file = st.file_uploader("📥 2. Upload PR Data (Excel)", type=['xlsx', 'xls'])

if po_file and pr_file:
    try:
        with st.spinner("Crunching the data..."):
            # Load raw data
            df_po = pd.read_excel(po_file)
            df_pr = pd.read_excel(pr_file)

            # Clean column names
            df_po.columns = df_po.columns.str.strip()
            df_pr.columns = df_pr.columns.str.strip()

            # Process PR Data
            df_pr_clean = df_pr.dropna(subset=['PO Number']).copy()
            pr_summary = df_pr_clean.groupby('PO Number').agg(
                PR_nos=('Receive Number', lambda x: " & ".join(sorted(x.dropna().astype(str).unique()))),
                PR_Qty=('Quantity Received', 'sum'),
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
            
            # Fill Missing Values & Calculate Quantities
            merged['Vendor Name'] = merged['Vendor_PR'].fillna(merged['Vendor_PO'])
            merged['PO Qty'] = merged['PO_Qty'].fillna(0).astype(int)
            merged['PR Qty'] = merged['PR_Qty'].fillna(0).astype(int)
            
            # Calculate Excess or Short
            merged['Excess / Short'] = merged['PR Qty'] - merged['PO Qty']
            
            # Calculate Fill Rate for the progress bar
            merged['PO FR %'] = (merged['PR Qty'] / merged['PO Qty']).fillna(0) * 100
            merged['Sl.no'] = range(1, len(merged) + 1)

            # --- TOP KPI METRICS SECTION ---
            st.markdown("### 🎯 Key Performance Indicators")
            
            kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
            
            total_pos = len(merged)
            total_po_qty = merged['PO Qty'].sum()
            total_pr_qty = merged['PR Qty'].sum()
            net_variance = total_pr_qty - total_po_qty
            overall_fill_rate = (total_pr_qty / total_po_qty * 100) if total_po_qty > 0 else 0

            kpi1.metric("Total POs", f"{total_pos:,}")
            kpi2.metric("Total PO Qty", f"{total_po_qty:,}")
            kpi3.metric("Total PR Qty", f"{total_pr_qty:,}")
            kpi4.metric("Excess / Short", f"{net_variance:,}")
            kpi5.metric("Overall Fill Rate", f"{overall_fill_rate:.1f}%")
            
            st.markdown("---")
            
            # --- DETAILED TABLE SECTION ---
            st.markdown("### 📋 Detailed Summary Tracker")
            
            expected_headers = [
                "Sl.no", "PO Number", "PR_nos", "Vendor Name", 
                "PO Qty", "PR Qty", "Excess / Short", "PO FR %"
            ]
            final_df = merged[expected_headers].rename(columns={'PO Number': 'PO No', 'PR_nos': 'PR no'})

            # Display updated interactive table
            st.dataframe(
                final_df,
                hide_index=True,
                use_container_width=True,
                column_config={
                    "PO Qty": st.column_config.NumberColumn(format="%d"),
                    "PR Qty": st.column_config.NumberColumn(format="%d"),
                    "Excess / Short": st.column_config.NumberColumn(
                        "Excess / Short", 
                        help="Negative number means shortage. Positive means excess.",
                        format="%d"
                    ),
                    "PO FR %": st.column_config.ProgressColumn(
                        "Fill Rate %",
                        help="Percentage of PO quantity fulfilled",
                        format="%.0f%%",
                        min_value=0,
                        max_value=100,
                    )
                }
            )

    except Exception as e:
        st.error(f"An error occurred while processing the files. Error details: {e}")
