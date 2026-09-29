import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="PO & PR Tracker", layout="wide")

st.title("📊 Automated PO vs PR Summary Generator")
st.write("Upload your daily raw PO and PR Excel files below to instantly generate your summary tracker.")

# File uploaders
col1, col2 = st.columns(2)
with col1:
    po_file = st.file_uploader("1. Upload PO Data (Excel)", type=['xlsx', 'xls'])
with col2:
    pr_file = st.file_uploader("2. Upload PR Data (Excel)", type=['xlsx', 'xls'])

if po_file and pr_file:
    try:
        with st.spinner("Processing data..."):
            # Load raw data
            df_po = pd.read_excel(po_file)
            df_pr = pd.read_excel(pr_file)

            # Clean column names (strip trailing spaces just in case)
            df_po.columns = df_po.columns.str.strip()
            df_pr.columns = df_pr.columns.str.strip()

            # 1. Process PR Data
            df_pr_clean = df_pr.dropna(subset=['PO Number']).copy()
            pr_summary = df_pr_clean.groupby('PO Number').agg(
                PR_nos=('Receive Number', lambda x: " & ".join(sorted(x.dropna().astype(str).unique()))),
                PR_Qty=('Quantity Received', 'sum'),
                PR_Date=('Receive Date', 'first'),
                Vendor_PR=('Vendor Name', 'first')
            ).reset_index()

            # 2. Process PO Data
            df_po_clean = df_po.dropna(subset=['Purchase Order Number']).copy()
            po_summary = df_po_clean.groupby('Purchase Order Number').agg(
                PO_Qty=('QuantityOrdered', 'sum'),
                PO_Date=('Purchase Order Date', 'first'),
                Vendor_PO=('Vendor Name', 'first')
            ).reset_index()

            # 3. Merge PR and PO Data
            merged = pd.merge(pr_summary, po_summary, left_on='PO Number', right_on='Purchase Order Number', how='left')

            # 4. Calculate Ageing and Formats
            # Strip timestamps from dates
            merged['PO Issued Date'] = pd.to_datetime(merged['PO_Date'], errors='coerce').dt.strftime('%d-%m-%Y')
            merged['PR Done Date'] = pd.to_datetime(merged['PR_Date'], errors='coerce').dt.strftime('%d-%m-%Y')
            
            # Calculate Ageing in days
            merged['Ageing (Days)'] = (pd.to_datetime(merged['PR_Date'], errors='coerce') - pd.to_datetime(merged['PO_Date'], errors='coerce')).dt.days
            
            # Merge Vendors and calculate quantities
            merged['Vendor Name'] = merged['Vendor_PR'].fillna(merged['Vendor_PO'])
            merged['PO Qty'] = merged['PO_Qty'].fillna(0).astype(int)
            merged['PR Qty'] = merged['PR_Qty'].fillna(0).astype(int)
            
            # Calculate Fill Rate %
            merged['PO FR %'] = (merged['PR Qty'] / merged['PO Qty']).fillna(0)
            
            # Add Serial Number
            merged['Sl.no'] = range(1, len(merged) + 1)

            # 5. Format Final Output Table
            expected_headers = [
                "Sl.no", "PO Issued Date", "PR Done Date", "Ageing (Days)", 
                "PO Number", "PR_nos", "Vendor Name", "PO Qty", "PR Qty", "PO FR %"
            ]
            final_df = merged[expected_headers].rename(columns={'PO Number': 'PO No', 'PR_nos': 'PR no'})

            # Format percentage for the web display
            display_df = final_df.copy()
            display_df['PO FR %'] = (display_df['PO FR %'] * 100).round(0).astype(int).astype(str) + '%'

            st.success("✅ Summary Generated Successfully!")
            st.dataframe(display_df, use_container_width=True)

            # 6. Export to Excel
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                final_df.to_excel(writer, index=False, sheet_name='Summary Tracker')
                # Optional: Format the percentage column in the exported Excel
                workbook = writer.book
                worksheet = writer.sheets['Summary Tracker']
                for cell in worksheet['J'][1:]:  # Column J is PO FR %
                    cell.number_format = '0%'

            st.download_button(
                label="📥 Download Summary Excel File",
                data=buffer.getvalue(),
                file_name="Daily_PO_PR_Summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"An error occurred while processing the files. Please ensure the uploaded files are the correct PO and PR exports. Error details: {e}")
