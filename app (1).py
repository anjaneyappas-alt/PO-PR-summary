import streamlit as st
import pandas as pd
import numpy as np
import io
import json
import requests
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor, as_completed

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import altair as alt


# Page Config
st.set_page_config(page_title="PO vs PR · DC Inward", page_icon="📦", layout="wide",
                   initial_sidebar_state="expanded")

# ==============================================================
# THEME — light, same family as the DC01 Bin Finder app:
# grey-blue canvas, white cards, blue accent, Fira Sans + Fira Code.
# ==============================================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Fira+Code:wght@400;500;600;700&family=Fira+Sans:wght@400;500;600;700&display=swap');

:root {
    --bg: #EEF2F7;
    --side: #F8FAFC;
    --card: #FFFFFF;
    --fg: #0F172A;
    --muted: #475569;
    --faint: #64748B;
    --border: #E2E8F0;
    --border-2: #CBD5E1;
    --blue: #1D4ED8;
    --blue-soft: #DBEAFE;
    --green: #15803D; --green-soft: #DCFCE7;
    --amber: #B45309; --amber-soft: #FEF3C7;
    --red: #B91C1C;   --red-soft: #FEE2E2;
    --radius: 14px;
}

/* Keep Streamlit's header (it holds the sidebar open/close arrow), just make it blend in */
[data-testid="stHeader"] { background: transparent !important; }
[data-testid="stDecoration"], [data-testid="stStatusWidget"], .stDeployButton,
[data-testid="stAppDeployButton"], #MainMenu, footer { display: none !important; }

html, body, .stApp, .stMarkdown, button, input, label, p, li { font-family: "Fira Sans", system-ui, "Segoe UI", sans-serif; }
.stApp { background: var(--bg) !important; color: var(--fg); }
.block-container { padding-top: 1.4rem !important; max-width: 1500px; }

/* ---------- Sidebar ---------- */
[data-testid="stSidebar"] { background: var(--side) !important; border-right: 1px solid var(--border); }
[data-testid="stSidebar"] .block-container, [data-testid="stSidebarContent"] { padding-top: 0.6rem; }
.brand { display: flex; align-items: center; gap: 12px; margin: 0 0 22px 4px; }
.brand-mark { width: 38px; height: 38px; border-radius: 10px; display: grid; place-items: center;
              background: var(--blue); color: #FFFFFF; }
.brand-name { font-weight: 700; font-size: 17px; color: var(--fg); line-height: 1.2; }
.brand-sub  { font-size: 13px; color: var(--muted); }

/* Sidebar radio -> nav list (works with old and new Streamlit radio markup) */
[data-testid="stSidebar"] [data-testid="stRadio"], [data-testid="stSidebar"] [data-testid="stRadioGroup"],
[data-testid="stSidebar"] [role="radiogroup"], [data-testid="stSidebar"] [role="radiogroup"] > * { width: 100% !important; max-width: 100% !important; }
[data-testid="stSidebar"] [role="radiogroup"] { display: flex; flex-direction: column; gap: 4px; }
[data-testid="stSidebar"] [role="radiogroup"] label,
[data-testid="stSidebar"] [data-testid="stRadioOption"] {
    width: 100% !important; margin: 0 !important; padding: 11px 14px !important; border-radius: 10px; cursor: pointer;
    transition: background-color 150ms ease, color 150ms ease;
}
/* hide the radio circle */
[data-testid="stSidebar"] [data-testid="stRadioOption"] > div > div:not([data-testid="stMarkdownContainer"]),
[data-testid="stSidebar"] [role="radiogroup"] label > div:first-child:not(:has([data-testid="stMarkdownContainer"])) { display: none !important; }
[data-testid="stSidebar"] [role="radiogroup"] label p {
    font-size: 16px !important; font-weight: 500 !important; color: var(--muted) !important; margin: 0;
}
[data-testid="stSidebar"] [role="radiogroup"] label:hover { background: #EEF2F7; }
[data-testid="stSidebar"] [role="radiogroup"] label[data-selected="true"],
[data-testid="stSidebar"] [role="radiogroup"] label:has(input:checked) { background: var(--blue-soft); }
[data-testid="stSidebar"] [role="radiogroup"] label[data-selected="true"] p,
[data-testid="stSidebar"] [role="radiogroup"] label:has(input:checked) p { color: var(--blue) !important; font-weight: 600 !important; }

.side-foot { font-size: 13px; color: var(--muted); line-height: 1.55; border-top: 1px solid var(--border);
             padding-top: 14px; margin-top: 28px; }
.side-foot code { font-family: "Fira Code", monospace; background: none; color: var(--fg); padding: 0; font-size: 12px; }

/* ---------- Top bar ---------- */
.stTextInput input {
    background: var(--card) !important; color: var(--fg) !important; min-height: 48px;
    border-radius: 12px !important; font-size: 16px;
}
div[data-testid="stTextInput"] div[data-baseweb="input"] { border-radius: 12px !important; border-color: var(--border-2) !important; }
.st-key-topsearch input {
    padding-left: 44px !important;
    background: var(--card) url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='18' height='18' viewBox='0 0 24 24' fill='none' stroke='%2364748B' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'%3E%3Ccircle cx='11' cy='11' r='8'/%3E%3Cpath d='m21 21-4.3-4.3'/%3E%3C/svg%3E") no-repeat 16px center !important;
}
/* Dark selectors, like the warehouse picker (new + old Streamlit markup) */
.st-key-topsource [role="group"], .st-key-topwindow [role="group"],
.st-key-topsource [data-baseweb="select"] > div, .st-key-topwindow [data-baseweb="select"] > div {
    background: #0F172A !important; border-color: #0F172A !important; border-radius: 12px !important; min-height: 48px;
}
.st-key-topsource input, .st-key-topwindow input,
.st-key-topsource [data-baseweb="select"] *, .st-key-topwindow [data-baseweb="select"] *,
.st-key-topsource button, .st-key-topwindow button, .st-key-topsource svg, .st-key-topwindow svg {
    color: #FFFFFF !important; fill: #FFFFFF !important; -webkit-text-fill-color: #FFFFFF; font-weight: 600; cursor: pointer;
}
[data-testid="stTextInputRootElement"] { border-radius: 12px !important; border-color: var(--border-2) !important; background: var(--card) !important; }

/* ---------- Page heading ---------- */
.page-h1 { font-size: 40px; font-weight: 700; color: var(--fg); letter-spacing: -0.02em; margin: 14px 0 4px; line-height: 1.15; }
.page-sub { font-size: 17px; color: var(--muted); margin-bottom: 22px; }
.page-sub code { font-family: "Fira Code", monospace; background: var(--card); border: 1px solid var(--border);
                 color: var(--fg); border-radius: 6px; padding: 2px 6px; font-size: 14px; }

/* ---------- Cards / KPIs ---------- */
.bento { display: grid; gap: 14px; margin-bottom: 22px; grid-template-columns: 1.5fr repeat(3, 1fr); }
.card { background: var(--card); border: 1px solid var(--border); border-radius: var(--radius); padding: 18px 20px; }
.k-top { display: flex; align-items: center; justify-content: space-between; }
.k-label { font-size: 14px; color: var(--muted); font-weight: 500; }
.k-icon { width: 34px; height: 34px; border-radius: 10px; display: grid; place-items: center; background: #F1F5F9; color: var(--faint); }
.k-value { font-family: "Fira Code", monospace; font-variant-numeric: tabular-nums; font-size: 28px; font-weight: 600; color: var(--fg); margin-top: 10px; line-height: 1.1; }
.k-sub { font-size: 13px; color: var(--faint); margin-top: 6px; }
.card.warn .k-icon { background: var(--amber-soft); color: var(--amber); } .card.warn .k-value { color: var(--amber); }
.card.danger .k-icon { background: var(--red-soft); color: var(--red); }  .card.danger .k-value { color: var(--red); }
.card.blue .k-icon { background: var(--blue-soft); color: var(--blue); }
.card.hero-kpi { grid-row: span 2; display: flex; flex-direction: column; justify-content: space-between; }
.hero-kpi .k-value { font-size: 54px; font-weight: 700; margin-top: 14px; }
.hero-kpi .k-value small { font-size: 22px; color: var(--faint); font-weight: 500; }
.meter-track { height: 10px; border-radius: 999px; background: #E2E8F0; overflow: hidden; margin-top: 18px; }
.meter-fill  { height: 100%; border-radius: 999px; background: var(--blue); }
.meter-legend { display: flex; justify-content: space-between; font-size: 12px; color: var(--faint); margin-top: 8px; font-family: "Fira Code", monospace; }
@media (max-width: 1100px) { .bento { grid-template-columns: repeat(2, 1fr); } .card.hero-kpi { grid-row: auto; grid-column: span 2; } }
@media (max-width: 560px)  { .bento { grid-template-columns: 1fr; } .card.hero-kpi { grid-column: auto; } .page-h1 { font-size: 30px; } }

/* Section card header (like "Rack A1 / 13 of 18 bins ...") */
.sec-title { font-size: 22px; font-weight: 700; color: var(--fg); margin: 0; }
.sec-sub { font-size: 15px; color: var(--muted); margin: 2px 0 14px; }

/* Bordered containers -> white cards */
[data-testid="stVerticalBlockBorderWrapper"], [class*="st-key-card"] { background: var(--card); border-color: var(--border) !important; border-radius: var(--radius) !important; }

.chips { display: flex; flex-wrap: wrap; gap: 8px; margin: 4px 0 14px; }
.chip { font-size: 13px; color: var(--fg); background: #F8FAFC; border: 1px solid var(--border); border-radius: 999px; padding: 5px 12px; }
.chip b { color: var(--muted); font-weight: 500; margin-right: 6px; }
.chip span { font-family: "Fira Code", monospace; font-weight: 600; }
.chip.amber { background: var(--amber-soft); border-color: #FDE68A; } .chip.amber span { color: var(--amber); }
.chip.red { background: var(--red-soft); border-color: #FECACA; } .chip.red span { color: var(--red); }
.chip.green { background: var(--green-soft); border-color: #BBF7D0; } .chip.green span { color: var(--green); }

.empty { text-align: center; padding: 48px 20px; border: 1px dashed var(--border-2); border-radius: var(--radius); color: var(--muted); background: var(--card); }
.empty .t { color: var(--fg); font-weight: 600; font-size: 17px; margin-bottom: 4px; }

[data-testid="stDataFrame"] { border: 1px solid var(--border); border-radius: 12px; overflow: hidden; }

/* ---------- Pills / segmented control ---------- */
[data-testid="stButtonGroup"] button { border-radius: 999px !important; min-height: 40px; font-weight: 600; }

/* ---------- Buttons ---------- */
.stButton > button, .stDownloadButton > button {
    min-height: 48px; border-radius: 12px !important; font-weight: 600 !important; cursor: pointer;
    transition: background-color 150ms ease, border-color 150ms ease, color 150ms ease;
}
.stButton > button[kind="secondary"], .stDownloadButton > button {
    background: var(--card) !important; color: var(--fg) !important; border: 1px solid var(--border-2) !important;
}
.stButton > button[kind="secondary"]:hover, .stDownloadButton > button:hover { border-color: var(--blue) !important; color: var(--blue) !important; }
.stButton > button[kind="primary"] { background: var(--blue) !important; border-color: var(--blue) !important; color: #FFFFFF !important; }
.stButton > button[kind="primary"]:hover { background: #1E40AF !important; border-color: #1E40AF !important; }
button:focus-visible, input:focus-visible, [role="radio"]:focus-visible { outline: 2px solid var(--blue) !important; outline-offset: 2px !important; }

@media (prefers-reduced-motion: reduce) { * { transition: none !important; animation: none !important; } }
</style>
""", unsafe_allow_html=True)

# Lucide icons (stroke SVG), decorative -> aria-hidden
_ICON_PATHS = {
    "package": '<path d="m7.5 4.27 9 5.15"/><path d="M21 8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16Z"/><path d="m3.3 7 8.7 5 8.7-5"/><path d="M12 22V12"/>',
    "file": '<path d="M15 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7Z"/><path d="M14 2v4a2 2 0 0 0 2 2h4"/><path d="M16 13H8"/><path d="M16 17H8"/>',
    "truck": '<path d="M14 18V6a2 2 0 0 0-2-2H4a2 2 0 0 0-2 2v11a1 1 0 0 0 1 1h2"/><path d="M15 18H9"/><path d="M19 18h2a1 1 0 0 0 1-1v-3.65a1 1 0 0 0-.22-.624l-3.48-4.35A1 1 0 0 0 17.52 8H14"/><circle cx="17" cy="18" r="2"/><circle cx="7" cy="18" r="2"/>',
    "calendar": '<rect width="18" height="18" x="3" y="4" rx="2"/><path d="M16 2v4"/><path d="M8 2v4"/><path d="M3 10h18"/>',
    "clock": '<circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/>',
    "alert": '<path d="m21.73 18-8-14a2 2 0 0 0-3.48 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.73-3"/><path d="M12 9v4"/><path d="M12 17h.01"/>',
    "gauge": '<path d="m12 14 4-4"/><path d="M3.34 19a10 10 0 1 1 17.32 0"/>',
}

def icon(name, size=18):
    return (f'<svg aria-hidden="true" width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" '
            f'stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">'
            f'{_ICON_PATHS[name]}</svg>')



def find_first_existing_col(df, candidates, default=None):
    for col in candidates:
        if col in df.columns:
            return col
    return default

def detect_damage_reason(pr_number, notes=""):
    combined = f"{pr_number} {notes}".lower()
    if "short" in combined or "shortage" in combined:
        return "Shortage"
    return "Damage"

def classify_pr_source(attachment_names, notes="", receive_number="", bill_number=""):
    combined_meta = f"{notes} {receive_number} {bill_number}".lower()
    if attachment_names:
        for name in attachment_names:
            name_lower = str(name).lower()
            if "bill" in name_lower:
                return "HHD"
            if name_lower.startswith("202") or "whatsapp" in name_lower:
                return "Zoho"
    if "hhd" in combined_meta:
        return "HHD"
    return "Zoho"



# --- ZOHO API HELPERS ---
def get_zoho_access_token():
    client_id = st.secrets["zoho"]["client_id"]
    client_secret = st.secrets["zoho"]["client_secret"]
    refresh_token = st.secrets["zoho"]["refresh_token"]
    accounts_url = st.secrets["zoho"].get("accounts_url", "https://accounts.zoho.in")

    url = f"{accounts_url}/oauth/v2/token"
    payload = {
        "refresh_token": refresh_token,
        "client_id": client_id,
        "client_secret": client_secret,
        "grant_type": "refresh_token"
    }
    headers = {"Content-Type": "application/x-www-form-urlencoded"}
    response = requests.post(url, data=payload, headers=headers)
    res_data = response.json()
    if "access_token" in res_data:
        return res_data["access_token"]
    else:
        raise Exception(f"Failed to refresh Zoho Token: {res_data}")

def fetch_po_by_number(po_number, headers, domain, org_id):
    try:
        url = f"https://www.zohoapis.{domain}/inventory/v1/purchaseorders"
        params = {"organization_id": org_id, "purchaseorder_number": po_number}
        res = requests.get(url, headers=headers, params=params).json()
        po_list = res.get("purchaseorders", [])
        if po_list:
            po_id = po_list[0].get("purchaseorder_id")
            single_url = f"https://www.zohoapis.{domain}/inventory/v1/purchaseorders/{po_id}"
            single_res = requests.get(single_url, headers=headers, params={"organization_id": org_id}).json()
            if "purchaseorder" in single_res:
                po_detail = single_res["purchaseorder"]
                line_items = po_detail.get("line_items", [])
                total_qty = sum(float(item.get("quantity", 0)) for item in line_items)
                po_detail["calculated_po_qty"] = total_qty
                return po_detail
            else:
                po = po_list[0]
                po["calculated_po_qty"] = float(po.get("quantity", po.get("quantity_ordered", 0)))
                return po
    except Exception:
        pass
    return None

def fetch_single_pr_details(pr_id, headers, domain, org_id):
    try:
        url = f"https://www.zohoapis.{domain}/inventory/v1/purchasereceives/{pr_id}"
        res = requests.get(url, headers=headers, params={"organization_id": org_id}).json()
        if "purchasereceive" in res:
            pr_data = res["purchasereceive"]
            attachments = []

            for doc in pr_data.get("documents", []):
                name = doc.get("file_name", "") if isinstance(doc, dict) else str(doc)
                if name:
                    attachments.append(name)

            for att in pr_data.get("attachments", []):
                name = att.get("file_name", "") if isinstance(att, dict) else str(att)
                if name:
                    attachments.append(name)

            if not attachments:
                try:
                    att_url = f"https://www.zohoapis.{domain}/inventory/v1/purchasereceives/{pr_id}/attachments"
                    att_res = requests.get(att_url, headers=headers, params={"organization_id": org_id}, timeout=8).json()
                    for item in att_res.get("attachments", []):
                        f_name = item.get("file_name", "")
                        if f_name:
                            attachments.append(f_name)
                except Exception:
                    pass

            pr_data["attachment_names"] = attachments
            return pr_data
    except Exception:
        pass
    return None

def fetch_zoho_data_live():
    access_token = get_zoho_access_token()
    org_id = st.secrets["zoho"]["organization_id"]
    domain = st.secrets["zoho"].get("domain", "zoho.in")
    headers = {"Authorization": f"Zoho-oauthtoken {access_token}"}
    date_30_days_ago = (datetime.now() - timedelta(days=30)).strftime("%Y-%m-%d")

    pr_url = f"https://www.zohoapis.{domain}/inventory/v1/purchasereceives"
    pr_params = {
        "organization_id": org_id,
        "date_after": date_30_days_ago,
        "sort_column": "created_time",
        "sort_order": "D"
    }
    pr_res = requests.get(pr_url, headers=headers, params=pr_params).json()
    if "purchasereceives" not in pr_res:
        raise Exception(f"Zoho Purchase Receives API error: {pr_res}")

    pr_list = pr_res.get("purchasereceives", [])
    df_pr = pd.DataFrame(pr_list)

    po_ref_col = find_first_existing_col(df_pr, ['purchaseorder_number', 'PO Number', 'po_number', 'purchaseorder_no'])
    unique_po_numbers = list(set(df_pr[po_ref_col].dropna().unique())) if (not df_pr.empty and po_ref_col in df_pr.columns) else []

    detailed_pos = []
    if unique_po_numbers:
        with ThreadPoolExecutor(max_workers=10) as executor:
            futures = [
                executor.submit(fetch_po_by_number, po_num, headers, domain, org_id)
                for po_num in unique_po_numbers
            ]
            for future in as_completed(futures):
                result = future.result()
                if result:
                    detailed_pos.append(result)

    df_po = pd.DataFrame(detailed_pos) if detailed_pos else pd.DataFrame()
    return df_po, df_pr


# ==============================================================
# TOP BAR — controls live in the page itself, so the app works
# even when the sidebar is collapsed (e.g. on a phone).
# ==============================================================
df_po_raw, df_pr_raw = None, None

tb1, tb2, tb3, tb4 = st.columns([3.2, 1.3, 1.2, 1.1], vertical_alignment="center")
with tb1:
    query = st.text_input("Search", key="topsearch", placeholder="Search a PO, PR, vendor or item",
                          label_visibility="collapsed")
with tb2:
    data_source = st.selectbox("Source", ["Live Zoho API Sync", "Upload Files"], key="topsource",
                               label_visibility="collapsed",
                               format_func=lambda s: "Live · Zoho" if s.startswith("Live") else "Upload files")
with tb3:
    selected_window = st.selectbox("Window", ["Last 15 Hours", "Last 24 Hours"], key="topwindow",
                                   label_visibility="collapsed")
    hours_window = 15 if selected_window == "Last 15 Hours" else 24
with tb4:
    refresh = st.button(":material/refresh: Refresh data", width="stretch",
                        disabled=(data_source == "Upload Files"))

if data_source == "Upload Files":
    u1, u2 = st.columns(2)
    with u1:
        po_file = st.file_uploader("PO monthly data (Excel)", type=['xlsx', 'xls'])
    with u2:
        pr_file = st.file_uploader("PR monthly data (Excel)", type=['xlsx', 'xls'])
    if po_file and pr_file:
        df_po_raw = pd.read_excel(po_file)
        df_pr_raw = pd.read_excel(pr_file)
else:
    if refresh:
        try:
            with st.spinner("Fetching PO & PR line items from Zoho..."):
                df_po_raw, df_pr_raw = fetch_zoho_data_live()
                st.session_state["raw_po"] = df_po_raw
                st.session_state["raw_pr"] = df_pr_raw
                st.session_state["fetched_at"] = datetime.now()
        except Exception as e:
            st.error(f"Zoho API error: {e}", icon=":material/error:")

    if "raw_po" in st.session_state and "raw_pr" in st.session_state:
        df_po_raw = st.session_state["raw_po"]
        df_pr_raw = st.session_state["raw_pr"]



# Main Processing
if df_po_raw is not None and df_pr_raw is not None:
    try:
        with st.spinner("Reconciling and checking damage notes / attachments..."):
            df_po = df_po_raw.copy()
            df_pr = df_pr_raw.copy()

            df_po.columns = df_po.columns.str.strip()
            df_pr.columns = df_pr.columns.str.strip()

            po_num_col = find_first_existing_col(df_po, ['purchaseorder_number', 'Purchase Order Number', 'PO Number', 'purchaseorder_no'])
            po_qty_col = find_first_existing_col(df_po, ['calculated_po_qty', 'QuantityOrdered', 'Quantity', 'quantity', 'quantity_ordered'])

            pr_no_col = find_first_existing_col(df_pr, ['purchasereceive_number', 'receive_number', 'Receive Number', 'PR Number', 'purchasereceive_no'])
            qty_pr_col = find_first_existing_col(df_pr, ['quantity', 'quantity_received', 'Quantity Received', 'total_quantity', 'Quantity'])
            notes_col = find_first_existing_col(df_pr, ['Notes', 'notes', 'Remarks', 'remarks'])
            bill_col = find_first_existing_col(df_pr, ['Bill Number', 'bill_number', 'Bill No', 'bill_no'])

            po_ref_in_pr = find_first_existing_col(df_pr, ['purchaseorder_number', 'PO Number', 'po_number', 'purchaseorder_no'])
            time_col = find_first_existing_col(df_pr, ['created_time', 'CreatedTime', 'date', 'Receive Date'])

            df_pr_clean = df_pr.dropna(subset=[po_ref_in_pr]).copy() if po_ref_in_pr else df_pr.copy()

            if qty_pr_col in df_pr_clean.columns:
                df_pr_clean['Clean_PR_Qty'] = pd.to_numeric(df_pr_clean[qty_pr_col], errors='coerce').fillna(0)
            else:
                df_pr_clean['Clean_PR_Qty'] = 0

            df_pr_clean['DT'] = pd.to_datetime(df_pr_clean[time_col], errors='coerce') if time_col else pd.NaT
            latest_pr_time = df_pr_clean['DT'].max() if 'DT' in df_pr_clean.columns else pd.NaT

            if pd.notna(latest_pr_time):
                cutoff_time = latest_pr_time - pd.Timedelta(hours=hours_window)
                df_today = df_pr_clean[df_pr_clean['DT'] >= cutoff_time]
                df_prior = df_pr_clean[df_pr_clean['DT'] < cutoff_time]
            else:
                df_today = df_pr_clean
                df_prior = pd.DataFrame(columns=df_pr_clean.columns)

            vendor_col_pr = find_first_existing_col(df_pr_clean, ['vendor_name', 'Vendor Name'])
            vendor_col_po = find_first_existing_col(df_po, ['vendor_name', 'Vendor Name'])

            today_summary = df_today.groupby(po_ref_in_pr).agg(
                PR_no=(pr_no_col, lambda x: " & ".join(sorted(x.dropna().astype(str).unique()))),
                PR_Qty=('Clean_PR_Qty', 'sum'),
                PR_Date=('DT', 'max'),
                Vendor_PR=(vendor_col_pr, 'first') if (vendor_col_pr and vendor_col_pr in df_today.columns) else (pr_no_col, 'first')
            ).reset_index()

            prior_summary = df_prior.groupby(po_ref_in_pr).agg(
                PRMTD=('Clean_PR_Qty', 'sum')
            ).reset_index()

            df_po_clean = df_po.dropna(subset=[po_num_col]).copy() if (po_num_col and not df_po.empty) else df_po.copy()

            if po_qty_col and po_qty_col in df_po_clean.columns:
                df_po_clean['Clean_PO_Qty'] = pd.to_numeric(df_po_clean[po_qty_col], errors='coerce').fillna(0)
            else:
                df_po_clean['Clean_PO_Qty'] = 0

            po_summary = df_po_clean.groupby(po_num_col).agg(
                PO_Qty=('Clean_PO_Qty', 'sum'),
                Vendor_PO=(vendor_col_po, 'first') if (vendor_col_po and vendor_col_po in df_po_clean.columns) else (po_num_col, 'first')
            ).reset_index() if not df_po_clean.empty else pd.DataFrame(columns=[po_ref_in_pr, 'PO_Qty', 'Vendor_PO'])

            merged = pd.merge(today_summary, po_summary, left_on=po_ref_in_pr, right_on=po_num_col if not po_summary.empty else po_ref_in_pr, how='left')
            merged = pd.merge(merged, prior_summary, on=po_ref_in_pr, how='left')

            if 'PR_Date' in merged.columns and pd.api.types.is_datetime64_any_dtype(merged['PR_Date']):
                merged = merged.sort_values(by='PR_Date', ascending=True).reset_index(drop=True)
                merged['Date'] = merged['PR_Date'].dt.strftime('%d-%m-%y').fillna('')
            else:
                merged['Date'] = ''

            merged['Vendor Name'] = merged['Vendor_PR'].fillna(merged.get('Vendor_PO', '')).fillna('')
            merged['PO Qty'] = pd.to_numeric(merged['PO_Qty'], errors='coerce').fillna(0).astype(int)
            merged['PR Qty'] = pd.to_numeric(merged['PR_Qty'], errors='coerce').fillna(0).astype(int)
            merged['PRMTD'] = pd.to_numeric(merged['PRMTD'], errors='coerce').fillna(0).astype(int)

            merged['Total Received'] = merged['PR Qty'] + merged['PRMTD']
            merged['Pending Qty'] = (merged['PO Qty'] - merged['Total Received']).apply(lambda x: int(x) if x > 0 else 0)
            merged['Sl.no'] = range(1, len(merged) + 1)

            total_pos_cnt = len(merged)
            total_po = int(merged['PO Qty'].sum())
            total_pr = int(merged['PR Qty'].sum())
            total_prmtd = int(merged['PRMTD'].sum())
            total_pending = int(merged['Pending Qty'].sum())
            total_fr_val = ((total_pr + total_prmtd) / total_po * 100) if total_po > 0 else 0

            expected_headers = ["Sl.no", "Date", po_ref_in_pr, "PR_no", "Vendor Name", "PO Qty", "PR Qty", "PRMTD", "Pending Qty"]
            final_df = merged[expected_headers].rename(columns={po_ref_in_pr: 'PO No', 'PR_no': 'PR No', 'PR Qty': 'Today PR Qty'})

            fr_numeric = np.where(final_df['PO Qty'] > 0, ((final_df['Today PR Qty'] + final_df['PRMTD']) / final_df['PO Qty']) * 100, 0)
            final_df['PO FR %'] = np.round(fr_numeric).astype(int).astype(str) + '%'

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

            # ----------------------------------------------------
            # DAMAGE & SHORT PR INSPECTION
            # ----------------------------------------------------
            detailed_damage_items = []

            if not df_today.empty:
                if data_source == "Live Zoho API Sync":
                    access_token = get_zoho_access_token()
                    org_id = st.secrets["zoho"]["organization_id"]
                    domain = st.secrets["zoho"].get("domain", "zoho.in")
                    headers = {"Authorization": f"Zoho-oauthtoken {access_token}"}

                    pr_id_col = find_first_existing_col(df_today, ['purchasereceive_id', 'Receive ID'])
                    unique_today_pr_ids = df_today[pr_id_col].dropna().unique().tolist() if pr_id_col else []

                    with ThreadPoolExecutor(max_workers=10) as executor:
                        futures = [
                            executor.submit(fetch_single_pr_details, pid, headers, domain, org_id)
                            for pid in unique_today_pr_ids
                        ]
                        for future in as_completed(futures):
                            pr_detail = future.result()
                            if pr_detail:
                                pr_raw_number = pr_detail.get("purchasereceive_number") or pr_detail.get("receive_number") or ""
                                notes_text = str(pr_detail.get("notes", "") or "")
                                bill_text = str(pr_detail.get("bill_number", "") or "")

                                combined_check = f"{pr_raw_number} {notes_text}".lower()
                                if any(kw in combined_check for kw in ["damage", "damaged", "expiry", "expired", "short", "shortage"]):
                                    r_date = pr_detail.get("date", "")
                                    if isinstance(r_date, pd.Timestamp):
                                        formatted_date = r_date.strftime("%d/%m/%Y")
                                    elif isinstance(r_date, str) and r_date:
                                        try:
                                            formatted_date = pd.to_datetime(r_date).strftime("%d/%m/%Y")
                                        except:
                                            formatted_date = r_date
                                    else:
                                        formatted_date = ""

                                    v_name = pr_detail.get("vendor_name", "")
                                    po_num = pr_detail.get("purchaseorder_number", "")
                                    item_reason = detect_damage_reason(pr_raw_number, notes_text)
                                    att_list = pr_detail.get("attachment_names", [])
                                    pr_done_on = classify_pr_source(att_list, notes_text, pr_raw_number, bill_text)

                                    line_items = pr_detail.get("line_items", [])
                                    if line_items:
                                        for item in line_items:
                                            qty_val = item.get("quantity", item.get("quantity_received", 0))
                                            try:
                                                qty_clean = int(round(float(qty_val)))
                                            except:
                                                qty_clean = qty_val

                                            detailed_damage_items.append({
                                                "Date": formatted_date,
                                                "vendor name": v_name,
                                                "PO NO": po_num,
                                                "Item": item.get("name", item.get("item_name", "")),
                                                "Qty": qty_clean,
                                                "reason": item_reason,
                                                "PR done on": pr_done_on
                                            })
                                    else:
                                        detailed_damage_items.append({
                                            "Date": formatted_date,
                                            "vendor name": v_name,
                                            "PO NO": po_num,
                                            "Item": "General Item",
                                            "Qty": 1,
                                            "reason": item_reason,
                                            "PR done on": pr_done_on
                                        })

                elif data_source == "Upload Files":
                    mask_pr = df_today[pr_no_col].astype(str).str.contains("damage|expiry|expired|short", case=False, na=False) if pr_no_col in df_today.columns else pd.Series(False, index=df_today.index)
                    mask_n = df_today[notes_col].astype(str).str.contains("damage|damaged|expiry|expired|short|shortage", case=False, na=False) if (notes_col and notes_col in df_today.columns) else pd.Series(False, index=df_today.index)
                    dmg_subset = df_today[mask_pr | mask_n]

                    item_col = find_first_existing_col(dmg_subset, ['Item Name', 'item_name', 'Item', 'name'])
                    att_col = find_first_existing_col(dmg_subset, ['attachment', 'Attachment', 'File Name', 'Documents', 'documents'])

                    for _, row in dmg_subset.iterrows():
                        r_date = row.get(time_col, "")
                        if isinstance(r_date, pd.Timestamp):
                            formatted_date = r_date.strftime("%d/%m/%Y")
                        elif isinstance(r_date, str) and r_date:
                            try:
                                formatted_date = pd.to_datetime(r_date).strftime("%d/%m/%Y")
                            except:
                                formatted_date = r_date
                        else:
                            formatted_date = ""

                        try:
                            qty_clean = int(round(float(row.get("Clean_PR_Qty", 0))))
                        except:
                            qty_clean = row.get("Clean_PR_Qty", 0)

                        row_notes = str(row.get(notes_col, "")) if notes_col else ""
                        row_bill = str(row.get(bill_col, "")) if bill_col else ""
                        item_reason = detect_damage_reason(row.get(pr_no_col, ""), row_notes)
                        att_val = [str(row[att_col])] if (att_col and pd.notna(row.get(att_col))) else []
                        pr_done_on = classify_pr_source(att_val, row_notes, str(row.get(pr_no_col, "")), row_bill)

                        detailed_damage_items.append({
                            "Date": formatted_date,
                            "vendor name": row.get(vendor_col_pr, ""),
                            "PO NO": row.get(po_ref_in_pr, ""),
                            "Item": row.get(item_col, "") if item_col else "",
                            "Qty": qty_clean,
                            "reason": item_reason,
                            "PR done on": pr_done_on
                        })

            st.session_state["raw_damage"] = pd.DataFrame(detailed_damage_items)

    except Exception as e:
        st.error(f"Error processing data: {e}", icon=":material/error:")



# ==============================================================
# EXCEL EXPORTS (unchanged format, matches the Google Sheet)
# ==============================================================

def generate_excel_file(df):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "PO vs PR Summary"
    ws.views.sheetView[0].showGridLines = True

    header_fill = PatternFill(start_color="E6E6E6", end_color="E6E6E6", fill_type="solid")
    zebra_fill = PatternFill(start_color="F9F9F9", end_color="F9F9F9", fill_type="solid")
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    yellow_pending_fill = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")
    orange_total_fill = PatternFill(start_color="F4B084", end_color="F4B084", fill_type="solid")

    thin_side = Side(border_style="thin", color="D9D9D9")
    thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    total_top_side = Side(border_style="thin", color="000000")
    total_bottom_side = Side(border_style="double", color="000000")
    total_border = Border(left=thin_side, right=thin_side, top=total_top_side, bottom=total_bottom_side)

    font_header = Font(name="Segoe UI", size=11, bold=True, color="000000")
    font_regular = Font(name="Segoe UI", size=11, bold=False, color="000000")
    font_pending = Font(name="Segoe UI", size=11, bold=True, color="856404")
    font_total = Font(name="Segoe UI", size=11, bold=True, color="000000")

    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    headers = list(df.columns)
    ws.append(headers)

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

                if col_idx == pending_col_idx:
                    try:
                        if float(val) > 0:
                            cell.fill = yellow_pending_fill
                            cell.font = font_pending
                    except (ValueError, TypeError):
                        pass

            if headers[col_idx - 1] in ['Sl.no', 'Date', 'PO No', 'PR No', 'PO FR %']:
                cell.alignment = align_center
            elif headers[col_idx - 1] in ['Vendor Name']:
                cell.alignment = align_left
            else:
                cell.alignment = align_right

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


def generate_damage_excel(df_dmg):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Damage Report"
        ws.views.sheetView[0].showGridLines = True

        header_fill = PatternFill(start_color="F2DC9B", end_color="F2DC9B", fill_type="solid")
        white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

        border_side = Side(border_style="thin", color="000000")
        full_border = Border(left=border_side, right=border_side, top=border_side, bottom=border_side)

        font_header = Font(name="Calibri", size=11, bold=True, color="000000")
        font_regular = Font(name="Calibri", size=11, bold=False, color="000000")

        align_center = Alignment(horizontal="center", vertical="center")
        align_left = Alignment(horizontal="left", vertical="center")
        align_right = Alignment(horizontal="right", vertical="center")

        cols = ["Date", "vendor name", "PO NO", "Item", "Qty", "reason", "PR done on"]
        ws.append(cols)
        for col_idx in range(1, len(cols) + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = font_header
            cell.border = full_border
            cell.alignment = align_center

        for r_idx, row in enumerate(df_dmg[cols].values, start=2):
            ws.append(list(row))
            for col_idx in range(1, len(cols) + 1):
                cell = ws.cell(row=r_idx, column=col_idx)
                cell.fill = white_fill
                cell.font = font_regular
                cell.border = full_border

                col_name = cols[col_idx - 1]
                if col_name in ["Date", "PO NO", "PR done on"]:
                    cell.alignment = align_center
                elif col_name == "Qty":
                    cell.alignment = align_right
                else:
                    cell.alignment = align_left

        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 4, 14)

        out = io.BytesIO()
        wb.save(out)
        return out.getvalue()



# ==============================================================
# CHART HELPER (light)
# ==============================================================
CHART_TEXT = "#334155"
CHART_GRID = "#E2E8F0"

def hbar(df, label_col, value_col, color, value_fmt=",.0f", ascending=False):
    """Horizontal bars with value labels."""
    height = max(150, 32 * len(df))
    base = alt.Chart(df).encode(
        y=alt.Y(f"{label_col}:N", sort="x" if ascending else "-x", title=None,
                axis=alt.Axis(labelColor=CHART_TEXT, labelFontSize=13, labelLimit=280, ticks=False, domain=False,
                              labelFont="Fira Sans")),
        x=alt.X(f"{value_col}:Q", title=None,
                axis=alt.Axis(labelColor="#64748B", gridColor=CHART_GRID, domain=False, ticks=False, format=value_fmt,
                              labelFont="Fira Code")),
        tooltip=[alt.Tooltip(f"{label_col}:N"), alt.Tooltip(f"{value_col}:Q", format=value_fmt)],
    )
    bars = base.mark_bar(color=color, cornerRadiusEnd=5, height=18)
    labels = base.mark_text(align="left", dx=6, color="#0F172A", font="Fira Code", fontSize=12).encode(
        text=alt.Text(f"{value_col}:Q", format=value_fmt))
    return (bars + labels).properties(height=height).configure_view(strokeWidth=0).configure(background="transparent")


def matches(df, cols, q):
    if not q:
        return df
    q = q.strip().lower()
    return df[df[cols].astype(str).apply(lambda c: c.str.lower().str.contains(q, regex=False)).any(axis=1)]


# ==============================================================
# STATE
# ==============================================================
has_data = "processed_df" in st.session_state
df_damage = st.session_state.get("raw_damage", pd.DataFrame()) if has_data else pd.DataFrame()
dmg_records = len(df_damage)
anchor_time = st.session_state.get("anchor_time") if has_data else None
active_window = st.session_state.get("window_hours", hours_window)
fetched_at = st.session_state.get("fetched_at")

# ==============================================================
# SIDEBAR — brand, navigation, sync info
# ==============================================================
NAV = {
    "Overview": ":material/space_dashboard: Overview",
    "Reconciliation": ":material/table_rows: Reconciliation",
    "Damage": f":material/warning: Damage & short{f'  ·  {dmg_records}' if dmg_records else ''}",
    "Insights": ":material/bar_chart: Insights",
}
with st.sidebar:
    st.markdown(f"""
        <div class="brand">
            <div class="brand-mark">{icon("package", 20)}</div>
            <div><div class="brand-name">DC01 Inward</div><div class="brand-sub">PO vs PR control room</div></div>
        </div>
    """, unsafe_allow_html=True)
    page = st.radio("Navigate", list(NAV.keys()), format_func=lambda k: NAV[k], key="nav",
                    label_visibility="collapsed")

    foot = []
    if data_source == "Upload Files":
        foot.append("Source: uploaded Excel files.")
    else:
        foot.append("Live from Zoho Inventory.")
    if fetched_at:
        foot.append(f"Synced <code>{fetched_at.strftime('%d %b, %H:%M')}</code>.")
    if anchor_time is not None and pd.notna(anchor_time):
        foot.append(f"Newest PR <code>{anchor_time.strftime('%d %b, %H:%M')}</code>, {active_window}h window.")
    if has_data and "raw_pr" in st.session_state:
        foot.append(f"{len(st.session_state['raw_pr']):,} PR records loaded.")
    st.markdown(f'<div class="side-foot">{" ".join(foot)}</div>', unsafe_allow_html=True)

# ==============================================================
# PAGES
# ==============================================================
TITLES = {
    "Overview": ("Overview", "Today's receipts against open purchase orders."),
    "Reconciliation": ("Reconciliation", "Every PO with a PR in the window, with what is still pending."),
    "Damage": ("Damage & short", "Item-wise damage and shortage from flagged PRs, ready to push to the sheet."),
    "Insights": ("Insights", "Where the pending quantity sits and which vendors fill slowest."),
}
h1, sub = TITLES[page]
if has_data and anchor_time is not None and pd.notna(anchor_time):
    sub += (f' Window: <code>{(anchor_time - pd.Timedelta(hours=active_window)).strftime("%d %b %H:%M")}</code>'
            f' to <code>{anchor_time.strftime("%d %b %H:%M")}</code>.')
st.markdown(f'<div class="page-h1">{h1}</div><div class="page-sub">{sub}</div>', unsafe_allow_html=True)

if not has_data:
    st.markdown("""
    <div class="empty">
        <div class="t">No data loaded</div>
        Click <b>Refresh data</b> at the top right, or choose <b>Upload files</b> and add the PO and PR exports.
    </div>""", unsafe_allow_html=True)
    st.stop()

display_df = st.session_state["processed_df"]
total_pos_cnt, total_po, total_pr, total_prmtd, total_pending, total_fr_val = st.session_state["kpi_metrics"]
total_received = total_pr + total_prmtd

recon_df = display_df[display_df["Vendor Name"] != "Total"].copy()
recon_df["Fill %"] = pd.to_numeric(recon_df["PO FR %"].astype(str).str.rstrip("%"), errors="coerce").fillna(0)
for c in ["PO Qty", "Today PR Qty", "PRMTD", "Pending Qty"]:
    recon_df[c] = pd.to_numeric(recon_df[c], errors="coerce").fillna(0).astype(int)

target_cols = ["Date", "vendor name", "PO NO", "Item", "Qty", "reason", "PR done on"]
if not df_damage.empty:
    for c in target_cols:
        if c not in df_damage.columns:
            df_damage[c] = ""
    df_damage_display = df_damage[target_cols]
else:
    df_damage_display = pd.DataFrame(columns=target_cols)
dmg_qty_num = pd.to_numeric(df_damage_display["Qty"], errors="coerce").fillna(0)
dmg_reason = df_damage_display["reason"].astype(str).str.lower()
dmg_units = int(dmg_qty_num.sum())


def top_pending_df(n=10):
    return (recon_df[recon_df["Pending Qty"] > 0].sort_values("Pending Qty", ascending=False).head(n)
            .assign(PO=lambda d: d["PO No"].astype(str) + " · " + d["Vendor Name"].astype(str).str.slice(0, 24)))


# ---------------- Overview ----------------
if page == "Overview":
    fr_bar = max(0.0, min(100.0, float(total_fr_val)))
    st.markdown(f"""
    <div class="bento">
        <div class="card hero-kpi">
            <div>
                <div class="k-top"><span class="k-label">Fill rate</span><span class="k-icon">{icon("gauge")}</span></div>
                <div class="k-value">{total_fr_val:.1f}<small>%</small></div>
                <div class="k-sub">Received {total_received:,} of {total_po:,} ordered across {total_pos_cnt:,} POs</div>
            </div>
            <div role="img" aria-label="Fill rate {total_fr_val:.1f} percent">
                <div class="meter-track"><div class="meter-fill" style="width:{fr_bar:.1f}%"></div></div>
                <div class="meter-legend"><span>0%</span><span>50%</span><span>100%</span></div>
            </div>
        </div>
        <div class="card"><div class="k-top"><span class="k-label">POs touched</span><span class="k-icon">{icon("file")}</span></div>
            <div class="k-value">{total_pos_cnt:,}</div><div class="k-sub">with a PR in this window</div></div>
        <div class="card"><div class="k-top"><span class="k-label">PO qty</span><span class="k-icon">{icon("package")}</span></div>
            <div class="k-value">{total_po:,}</div><div class="k-sub">ordered</div></div>
        <div class="card blue"><div class="k-top"><span class="k-label">Today PR qty</span><span class="k-icon">{icon("truck")}</span></div>
            <div class="k-value">{total_pr:,}</div><div class="k-sub">last {active_window} hours</div></div>
        <div class="card"><div class="k-top"><span class="k-label">PRMTD qty</span><span class="k-icon">{icon("calendar")}</span></div>
            <div class="k-value">{total_prmtd:,}</div><div class="k-sub">received earlier</div></div>
        <div class="card warn"><div class="k-top"><span class="k-label">Pending qty</span><span class="k-icon">{icon("clock")}</span></div>
            <div class="k-value">{total_pending:,}</div><div class="k-sub">still to receive</div></div>
        <div class="card danger"><div class="k-top"><span class="k-label">Damage / short</span><span class="k-icon">{icon("alert")}</span></div>
            <div class="k-value">{dmg_records:,}</div><div class="k-sub">{dmg_units:,} units flagged</div></div>
    </div>
    """, unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        with st.container(border=True, key="card_1"):
            st.markdown('<div class="sec-title">Most pending</div><div class="sec-sub">Top 5 POs by quantity still to receive</div>',
                        unsafe_allow_html=True)
            tp = top_pending_df(5)
            if tp.empty:
                st.markdown("Every PO in this window is fully received.")
            else:
                st.altair_chart(hbar(tp[["PO", "Pending Qty"]], "PO", "Pending Qty", "#D97706"), width="stretch")
    with c2:
        with st.container(border=True, key="card_2"):
            st.markdown('<div class="sec-title">Damage by vendor</div><div class="sec-sub">Units flagged as damage or short</div>',
                        unsafe_allow_html=True)
            if df_damage_display.empty:
                st.markdown("Nothing flagged in this window.")
            else:
                bv = (df_damage_display.assign(Units=dmg_qty_num).groupby("vendor name", as_index=False)["Units"].sum()
                      .sort_values("Units", ascending=False).head(5).rename(columns={"vendor name": "Vendor"}))
                st.altair_chart(hbar(bv, "Vendor", "Units", "#DC2626"), width="stretch")

# ---------------- Reconciliation ----------------
elif page == "Reconciliation":
    with st.container(border=True, key="card_3"):
        h_l, h_r = st.columns([3, 2], vertical_alignment="center")
        with h_l:
            st.markdown(f'<div class="sec-title">PO-wise detail</div>'
                        f'<div class="sec-sub">{int((recon_df["Pending Qty"] > 0).sum())} of {len(recon_df)} POs still have quantity pending</div>',
                        unsafe_allow_html=True)
        with h_r:
            scope = st.segmented_control("Show", ["All POs", "Pending only"], default="All POs",
                                         key="recon_scope", label_visibility="collapsed")

        view = matches(recon_df, ["PO No", "PR No", "Vendor Name"], query)
        if scope == "Pending only":
            view = view[view["Pending Qty"] > 0]

        cols = ["Sl.no", "Date", "PO No", "PR No", "Vendor Name", "PO Qty", "Today PR Qty", "PRMTD", "Pending Qty", "Fill %"]
        styler = view[cols].style
        styler = (styler.map if hasattr(styler, "map") else styler.applymap)(
            lambda v: "color: #B45309; font-weight: 600;" if isinstance(v, (int, np.integer)) and v > 0 else "color: #64748B;",
            subset=["Pending Qty"])
        st.dataframe(
            styler, hide_index=True, width="stretch", height=min(600, 38 * (len(view) + 1) + 4),
            column_config={
                "Sl.no": st.column_config.TextColumn("#", width="small"),
                "Date": st.column_config.TextColumn("Date", width="small"),
                "PO No": st.column_config.TextColumn("PO no"),
                "PR No": st.column_config.TextColumn("PR no"),
                "Vendor Name": st.column_config.TextColumn("Vendor", width="medium"),
                "PO Qty": st.column_config.NumberColumn("PO qty", format="%d"),
                "Today PR Qty": st.column_config.NumberColumn("Today PR", format="%d"),
                "PRMTD": st.column_config.NumberColumn("PRMTD", format="%d"),
                "Pending Qty": st.column_config.NumberColumn("Pending", format="%d"),
                "Fill %": st.column_config.ProgressColumn("Fill rate", min_value=0, max_value=100, format="%d%%"),
            })

        st.markdown(f"""
        <div class="chips">
            <span class="chip"><b>Showing</b><span>{len(view):,} / {len(recon_df):,}</span></span>
            <span class="chip"><b>PO</b><span>{total_po:,}</span></span>
            <span class="chip"><b>Today PR</b><span>{total_pr:,}</span></span>
            <span class="chip"><b>PRMTD</b><span>{total_prmtd:,}</span></span>
            <span class="chip amber"><b>Pending</b><span>{total_pending:,}</span></span>
            <span class="chip green"><b>Fill rate</b><span>{int(round(total_fr_val))}%</span></span>
        </div>""", unsafe_allow_html=True)
        st.download_button(":material/download: Download summary (.xlsx)", data=generate_excel_file(display_df),
                           file_name=f"PO_PR_Summary_{active_window}h.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ---------------- Damage & short ----------------
elif page == "Damage":
    if df_damage_display.empty:
        st.markdown(f'<div class="empty"><div class="t">No damage or short PRs</div>'
                    f'Nothing was flagged in the last {active_window} hours.</div>', unsafe_allow_html=True)
    else:
        with st.container(border=True, key="card_4"):
            st.markdown(f"""
            <div class="sec-title">Flagged items</div>
            <div class="sec-sub">{len(df_damage_display)} item(s) from {df_damage_display['vendor name'].nunique()} vendor(s)</div>
            <div class="chips">
                <span class="chip red"><b>Damage units</b><span>{int(dmg_qty_num[dmg_reason == 'damage'].sum()):,}</span></span>
                <span class="chip amber"><b>Short units</b><span>{int(dmg_qty_num[dmg_reason == 'shortage'].sum()):,}</span></span>
            </div>""", unsafe_allow_html=True)

            dview = matches(df_damage_display, ["vendor name", "PO NO", "Item"], query)
            dstyler = dview.style
            dstyler = (dstyler.map if hasattr(dstyler, "map") else dstyler.applymap)(
                lambda v: "color: #B91C1C; font-weight: 600;" if str(v).lower() == "damage"
                else ("color: #B45309; font-weight: 600;" if str(v).lower() == "shortage" else ""),
                subset=["reason"])
            st.dataframe(
                dstyler, hide_index=True, width="stretch", height=min(520, 38 * (len(dview) + 1) + 4),
                column_config={
                    "vendor name": st.column_config.TextColumn("Vendor"),
                    "PO NO": st.column_config.TextColumn("PO no"),
                    "Item": st.column_config.TextColumn("Item", width="large"),
                    "Qty": st.column_config.NumberColumn("Qty", format="%d"),
                    "reason": st.column_config.TextColumn("Reason"),
                    "PR done on": st.column_config.TextColumn("PR done on"),
                })

        with st.container(border=True, key="card_5"):
            st.markdown('<div class="sec-title">Send to Google Sheet</div>'
                        '<div class="sec-sub">Rows already in the sheet are skipped, so pushing twice is safe.</div>',
                        unsafe_allow_html=True)
            a1, a2, a3 = st.columns([1.4, 1, 1], vertical_alignment="bottom")
            with a1:
                sheet_tab_name = st.text_input("Sheet tab", value="Damage_Short")
            with a2:
                push = st.button(":material/cloud_upload: Push to Google Sheet", type="primary", width="stretch")
            with a3:
                st.download_button(":material/download: Damage report (.xlsx)",
                                   data=generate_damage_excel(df_damage_display),
                                   file_name=f"Damage_Report_{active_window}h.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   width="stretch")

            if push:
                try:
                    webhook_url = st.secrets.get("google", {}).get("webhook_url", "")
                    if not webhook_url:
                        st.error("Add `webhook_url` under `[google]` in Streamlit Secrets.", icon=":material/error:")
                    else:
                        with st.spinner(f"Pushing to '{sheet_tab_name}'..."):
                            payload = {"sheet_name": sheet_tab_name,
                                       "rows": df_damage_display.to_dict(orient="records")}
                            res = requests.post(webhook_url, data=json.dumps(payload),
                                                headers={"Content-Type": "text/plain;charset=utf-8"}, timeout=20)
                            try:
                                result_json = res.json()
                                if result_json.get("status") == "success":
                                    st.success(f"Added {result_json.get('appended', 0)} row(s), skipped "
                                               f"{result_json.get('skipped', 0)} duplicate(s) in '{sheet_tab_name}'.",
                                               icon=":material/check_circle:")
                                else:
                                    st.error(f"Apps Script error: {result_json.get('message')}", icon=":material/error:")
                            except Exception:
                                st.error("Webhook returned HTML instead of JSON. Deploy the Apps Script as a Web App "
                                         "with 'Who has access: Anyone'.", icon=":material/error:")
                except Exception as e:
                    st.error(f"Sync error: {e}", icon=":material/error:")

# ---------------- Insights ----------------
elif page == "Insights":
    c1, c2 = st.columns(2)
    with c1:
        with st.container(border=True, key="card_6"):
            st.markdown('<div class="sec-title">Top 10 POs by pending qty</div>'
                        '<div class="sec-sub">Chase these vendors first</div>', unsafe_allow_html=True)
            tp = top_pending_df(10)
            if tp.empty:
                st.markdown("Every PO in this window is fully received.")
            else:
                st.altair_chart(hbar(tp[["PO", "Pending Qty"]], "PO", "Pending Qty", "#D97706"), width="stretch")
    with c2:
        with st.container(border=True, key="card_7"):
            st.markdown('<div class="sec-title">Lowest fill rate by vendor</div>'
                        '<div class="sec-sub">Received ÷ ordered, in %</div>', unsafe_allow_html=True)
            vend = (recon_df.groupby("Vendor Name", as_index=False)
                    .agg(PO=("PO Qty", "sum"), Rec=("Today PR Qty", "sum"), Prior=("PRMTD", "sum")))
            vend = vend[vend["PO"] > 0]
            if vend.empty:
                st.markdown("Fill rate needs PO quantities.")
            else:
                vend["Fill rate"] = ((vend["Rec"] + vend["Prior"]) / vend["PO"] * 100).clip(upper=100).round(0)
                low = vend.sort_values("Fill rate").head(10).rename(columns={"Vendor Name": "Vendor"})
                low["Vendor"] = low["Vendor"].astype(str).str.slice(0, 30)
                st.altair_chart(hbar(low[["Vendor", "Fill rate"]], "Vendor", "Fill rate", "#2563EB",
                                     value_fmt=".0f", ascending=True), width="stretch")
