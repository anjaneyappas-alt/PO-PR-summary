import streamlit as st
import pandas as pd

st.set_page_config(page_title="PO vs PR Summary", layout="wide")

st.title("PO vs PR Summary")

# File uploaders
col1, col2 = st.columns(2)
with col1:
    po_file = st.file_uploader("Upload PO Data (Excel)", type=['xlsx', 'xls'])
with col2:
    pr_file = st.file_uploader("Upload PR Data (Excel)", type=['xlsx', 'xls'])

if po_file and pr_file:
    try:
        with st.spinner("Processing..."):
            # Load raw data
            df_po = pd.read_excel(po_file)
            df_pr = pd.read_excel(pr_file)

            # Clean column names
            df_po.columns = df_po.columns.str.strip()
            df_pr.columns = df_pr.columns.str.strip()

            # Process PR Data
            df_pr_clean = df_pr.dropna(subset=['PO Number']).copy()
            pr_summary = df_pr_clean.groupby('PO Number').agg(
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
            
            # Format Data
            merged['Date'] = pd.to_datetime(merged['PR_Date'], errors='coerce').dt.strftime('%d-%m-%y')
            merged['Vendor Name'] = merged['Vendor_PR'].fillna(merged['Vendor_PO'])
            merged['PO Qty'] = merged['PO_Qty'].fillna(0).astype(int)
            merged['PR Qty'] = merged['PR_Qty'].fillna(0).astype(int)
            
            # Add Serial Number
            merged['Sl.no'] = range(1, len(merged) + 1)
            
            # Reorder and rename columns to match your image exactly
            expected_headers = ["Sl.no", "Date", "PO Number", "Vendor Name", "PO Qty", "PR Qty"]
            final_df = merged[expected_headers].rename(columns={'PO Number': 'PO No'})

            # Calculate overall totals for the bottom row
            total_po = final_df['PO Qty'].sum()
            total_pr = final_df['PR Qty'].sum()
            total_fr = (total_pr / total_po * 100) if total_po > 0 else 0

            # Calculate individual Fill Rates and format as %
            final_df['PO FR %'] = ((final_df['PR Qty'] / final_df['PO Qty']).fillna(0) * 100).round(0).astype(int).astype(str) + '%'

            # Create the Total Row
            total_row = pd.DataFrame([{
                "Sl.no": "",
                "Date": "",
                "PO No": "",
                "Vendor Name": "Total",
                "PO Qty": total_po,
                "PR Qty": total_pr,
                "PO FR %": f"{int(round(total_fr))}%"
            }])

            # Append the Total Row to the bottom of the table
            final_df = pd.concat([final_df, total_row], ignore_index=True)

            # Highlight the Total row in a distinct color so it matches your image
            def highlight_total_row(row):
                if row['Vendor Name'] == 'Total':
                    return ['background-color: #f4b084; font-weight: bold; color: black'] * len(row)
                return [''] * len(row)

            # Apply the style and display
            styled_df = final_df.style.apply(highlight_total_row, axis=1)
            
            # Display it on the website
            st.dataframe(styled_df, hide_index=True, use_container_width=True)

    except Exception as e:
        st.error(f"Error: {e}")
