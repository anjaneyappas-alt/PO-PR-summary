import streamlit as st
import pandas as pd

# Page config
st.set_page_config(page_title="PO vs PR Summary", page_icon="📊", layout="wide")

# CSS to hide top menu, header decorations, and footer for clean screenshots
st.markdown("""
    <style>
        #MainMenu {visibility: hidden;}
        header {visibility: hidden;}
        footer {visibility: hidden;}
        [data-testid="stHeader"] {display: none;}
    </style>
""", unsafe_allow_html=True)

# --- 1. PASSWORD PROTECTION (PIN: 1006) ---
def check_password():
    if "authenticated" not in st.session_state:
        st.session_state["authenticated"] = False

    if not st.session_state["authenticated"]:
        st.title("🔒 Password Protected")
        password = st.text_input("Enter Passcode to Access Dashboard:", type="password")
        if st.button("Login"):
            if password == "1006":
                st.session_state["authenticated"] = True
                st.rerun()
            else:
                st.error("❌ Incorrect Passcode")
        return False
    return True

if check_password():
    st.title("📊 PO vs PR Summary")

    # --- 2. SLIDE / TOGGLE PANEL TO SHOW OR HIDE UPLOADERS ---
    show_uploaders = st.toggle("🎚️ Show Upload Panel", value=True)

    if show_uploaders:
        col1, col2 = st.columns(2)
        with col1:
            po_file = st.file_uploader("Upload PO Data (Excel)", type=['xlsx', 'xls'], key="po_up")
        with col2:
            pr_file = st.file_uploader("Upload PR Data (Excel)", type=['xlsx', 'xls'], key="pr_up")
    else:
        # Retain uploaded files in session state when toggled off
        po_file = st.session_state.get("po_up")
        pr_file = st.session_state.get("pr_up")

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
                
                # Sort chronologically from earliest to latest
                merged['Raw_Date'] = pd.to_datetime(merged['PR_Date'], errors='coerce')
                merged = merged.sort_values(by='Raw_Date', ascending=True).reset_index(drop=True)
                merged['Date'] = merged['Raw_Date'].dt.strftime('%d-%m-%y')
                
                # Clean values
                merged['Vendor Name'] = merged['Vendor_PR'].fillna(merged['Vendor_PO'])
                merged['PO Qty'] = merged['PO_Qty'].fillna(0).astype(int)
                merged['PR Qty'] = merged['PR_Qty'].fillna(0).astype(int)
                merged['Excess / Short'] = merged['PR Qty'] - merged['PO Qty']
                merged['Sl.no'] = range(1, len(merged) + 1)

                # Metrics
                total_po = merged['PO Qty'].sum()
                total_pr = merged['PR Qty'].sum()
                total_diff = merged['Excess / Short'].sum()
                total_fr = (total_pr / total_po * 100) if total_po > 0 else 0

                # Display Key Metrics
                st.markdown("### 🎯 Key Metrics")
                kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
                kpi1.metric("Total POs", f"{len(merged):,}")
                kpi2.metric("PO Qty", f"{total_po:,}")
                kpi3.metric("PR Qty", f"{total_pr:,}")
                kpi4.metric("Excess/Short", f"{total_diff:,}")
                kpi5.metric("Fill Rate", f"{total_fr:.1f}%")

                st.markdown("---")

                # Final Table Formatting
                expected_headers = ["Sl.no", "Date", "PO Number", "Vendor Name", "PO Qty", "PR Qty", "Excess / Short"]
                final_df = merged[expected_headers].rename(columns={'PO Number': 'PO No'})

                # Format Fill Rate %
                final_df['PO FR %'] = ((final_df['PR Qty'] / final_df['PO Qty']).fillna(0) * 100).round(0).astype(int).astype(str) + '%'

                # Total Row
                total_row = pd.DataFrame([{
                    "Sl.no": "",
                    "Date": "",
                    "PO No": "",
                    "Vendor Name": "Total",
                    "PO Qty": total_po,
                    "PR Qty": total_pr,
                    "Excess / Short": total_diff,
                    "PO FR %": f"{int(round(total_fr))}%"
                }])

                # Append Total Row
                final_df = pd.concat([final_df, total_row], ignore_index=True)

                # Style Total Row (Orange Highlight)
                def highlight_total_row(row):
                    if row['Vendor Name'] == 'Total':
                        return ['background-color: #f4b084; font-weight: bold; color: black'] * len(row)
                    return [''] * len(row)

                styled_df = final_df.style.apply(highlight_total_row, axis=1)

                # Display Table
                st.table(styled_df)

        except Exception as e:
            st.error(f"Error processing files: {e}")
