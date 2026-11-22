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
st.set_page_config(page_title="PO vs PR Control Room", page_icon="📦", layout="wide",
                   initial_sidebar_state="expanded")

# ==============================================================
# THEME — dark "control room": navy glass cards, status green,
# Fira Sans for text, Fira Code for numbers. Tokens live in :root.
# ==============================================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Fira+Code:wght@400;500;600;700&family=Fira+Sans:wght@400;500;600;700&display=swap');

:root {
    --bg: #0B1120;
    --card: rgba(27, 35, 54, 0.72);
    --card-solid: #1B2336;
    --border: rgba(148, 163, 184, 0.18);
    --fg: #F8FAFC;
    --muted: #94A3B8;
    --accent: #22C55E;
    --accent-2: #4ADE80;
    --amber: #FBBF24;
    --red: #F87171;
    --blue: #60A5FA;
    --radius: 16px;
}

#MainMenu, footer {visibility: hidden;}
[data-testid="stHeader"] {background: transparent !important;}
[data-testid="stToolbar"] {display: none;}

html, body, .stApp, .stMarkdown, button, input, label, p, span, div {
    font-family: "Fira Sans", system-ui, -apple-system, "Segoe UI", sans-serif;
}
.stApp {
    background:
        radial-gradient(1100px 520px at 8% -12%, rgba(34,197,94,0.14), transparent 60%),
        radial-gradient(900px 480px at 100% -5%, rgba(96,165,250,0.12), transparent 60%),
        var(--bg) !important;
    color: var(--fg);
}
.block-container { padding-top: 1.6rem !important; max-width: 1480px; }

/* ---------- Sidebar ---------- */
[data-testid="stSidebar"] {
    background: rgba(15, 23, 42, 0.92) !important;
    border-right: 1px solid var(--border);
    backdrop-filter: blur(14px);
}
.brand { display: flex; align-items: center; gap: 12px; margin: 4px 0 22px; }
.brand-mark {
    width: 40px; height: 40px; border-radius: 12px; display: grid; place-items: center;
    background: linear-gradient(135deg, var(--accent), #15803D); color: #052E16;
}
.brand-name { font-weight: 700; font-size: 16px; color: var(--fg); line-height: 1.2; }
.brand-sub  { font-size: 12px; color: var(--muted); }
.side-label {
    font-size: 11px; font-weight: 600; letter-spacing: .08em; text-transform: uppercase;
    color: var(--muted); margin: 18px 0 6px;
}
.side-note {
    font-size: 12px; color: var(--muted); border-top: 1px solid var(--border);
    padding-top: 14px; margin-top: 22px; line-height: 1.5;
}

/* ---------- Hero ---------- */
.hero {
    display: flex; justify-content: space-between; align-items: flex-end; gap: 16px; flex-wrap: wrap;
    margin-bottom: 22px;
}
.eyebrow { font-family: "Fira Code", monospace; font-size: 12px; letter-spacing: .12em; color: var(--accent-2); }
.hero h1 { margin: 6px 0 4px; font-size: 32px; font-weight: 700; color: var(--fg); letter-spacing: -0.02em; line-height: 1.15; }
.hero p  { margin: 0; color: var(--muted); font-size: 15px; }
.status {
    display: inline-flex; align-items: center; gap: 10px;
    background: var(--card); border: 1px solid var(--border); border-radius: 999px;
    padding: 8px 14px; font-size: 13px; color: var(--fg); backdrop-filter: blur(12px);
}
.status code { font-family: "Fira Code", monospace; background: none; color: var(--fg); padding: 0; }
.dot { width: 9px; height: 9px; border-radius: 50%; background: var(--accent); box-shadow: 0 0 0 4px rgba(34,197,94,.18); }
.dot.off { background: #64748B; box-shadow: 0 0 0 4px rgba(100,116,139,.18); }

/* ---------- Bento KPI grid ---------- */
.bento {
    display: grid; gap: 14px; margin-bottom: 26px;
    grid-template-columns: 1.5fr repeat(3, 1fr);
    grid-template-rows: auto auto;
}
.card {
    background: var(--card); border: 1px solid var(--border); border-radius: var(--radius);
    padding: 18px 20px; backdrop-filter: blur(14px); -webkit-backdrop-filter: blur(14px);
    transition: border-color 180ms ease, transform 180ms ease;
}
.card:hover { border-color: rgba(148,163,184,0.38); }
.card .k-top { display: flex; align-items: center; justify-content: space-between; }
.card .k-label { font-size: 13px; color: var(--muted); font-weight: 500; }
.card .k-icon { width: 34px; height: 34px; border-radius: 10px; display: grid; place-items: center; background: rgba(148,163,184,.10); color: var(--muted); }
.card .k-value { font-family: "Fira Code", monospace; font-variant-numeric: tabular-nums; font-size: 28px; font-weight: 600; color: var(--fg); margin-top: 10px; line-height: 1.1; }
.card .k-sub { font-size: 12px; color: var(--muted); margin-top: 6px; }
.card.warn .k-icon { background: rgba(251,191,36,.12); color: var(--amber); }
.card.warn .k-value { color: var(--amber); }
.card.danger .k-icon { background: rgba(248,113,113,.12); color: var(--red); }
.card.danger .k-value { color: var(--red); }
.card.blue .k-icon { background: rgba(96,165,250,.12); color: var(--blue); }

.card.hero-kpi {
    grid-row: span 2; display: flex; flex-direction: column; justify-content: space-between;
    background:
        radial-gradient(420px 220px at 100% 0%, rgba(34,197,94,.20), transparent 70%),
        var(--card);
}
.hero-kpi .k-value { font-size: 56px; font-weight: 700; color: var(--fg); margin-top: 14px; }
.hero-kpi .k-value small { font-size: 24px; color: var(--muted); font-weight: 500; }
.meter { margin-top: 18px; }
.meter-track { height: 10px; border-radius: 999px; background: rgba(148,163,184,.16); overflow: hidden; }
.meter-fill  { height: 100%; border-radius: 999px; background: linear-gradient(90deg, #16A34A, var(--accent-2)); }
.meter-legend { display: flex; justify-content: space-between; font-size: 12px; color: var(--muted); margin-top: 8px; font-family: "Fira Code", monospace; }

@media (max-width: 1100px) {
    .bento { grid-template-columns: repeat(2, 1fr); }
    .card.hero-kpi { grid-row: auto; grid-column: span 2; }
}
@media (max-width: 560px) {
    .bento { grid-template-columns: 1fr; }
    .card.hero-kpi { grid-column: auto; }
    .hero h1 { font-size: 26px; }
}

/* ---------- Tabs ---------- */
.stTabs [data-baseweb="tab-list"] {
    gap: 6px; background: var(--card); border: 1px solid var(--border);
    border-radius: 12px; padding: 5px; width: fit-content; max-width: 100%; overflow-x: auto;
}
.stTabs [data-baseweb="tab"] {
    height: 40px; padding: 0 16px; border-radius: 9px; color: var(--muted);
    background: transparent; transition: background-color 150ms ease, color 150ms ease;
}
.stTabs [data-baseweb="tab"]:hover { color: var(--fg); }
.stTabs [aria-selected="true"] { background: rgba(34,197,94,.14) !important; color: var(--accent-2) !important; }
.stTabs [data-baseweb="tab-highlight"], .stTabs [data-baseweb="tab-border"] { display: none; }
.stTabs [data-baseweb="tab-panel"] { padding-top: 18px; }

/* ---------- Panels, chips, tables ---------- */
.panel-title { font-size: 16px; font-weight: 600; color: var(--fg); margin: 4px 0 2px; }
.panel-sub { font-size: 13px; color: var(--muted); margin-bottom: 12px; }
.chips { display: flex; flex-wrap: wrap; gap: 8px; margin: 12px 0 16px; }
.chip {
    font-family: "Fira Code", monospace; font-size: 13px; color: var(--fg);
    background: var(--card); border: 1px solid var(--border); border-radius: 999px; padding: 6px 12px;
}
.chip b { color: var(--muted); font-weight: 500; font-family: "Fira Sans", sans-serif; margin-right: 6px; }
.chip.amber { border-color: rgba(251,191,36,.35); }
.chip.red { border-color: rgba(248,113,113,.35); }
.empty {
    text-align: center; padding: 48px 20px; border: 1px dashed var(--border); border-radius: var(--radius);
    color: var(--muted); background: rgba(27,35,54,.4);
}
.empty .t { color: var(--fg); font-weight: 600; font-size: 16px; margin-bottom: 4px; }

[data-testid="stDataFrame"] { border: 1px solid var(--border); border-radius: 12px; overflow: hidden; }
[data-testid="stVegaLiteChart"] { background: var(--card); border: 1px solid var(--border); border-radius: var(--radius); padding: 12px; }

/* ---------- Buttons & inputs ---------- */
.stButton > button, .stDownloadButton > button {
    min-height: 44px; border-radius: 10px !important; font-weight: 600 !important; cursor: pointer;
    transition: background-color 150ms ease, border-color 150ms ease, color 150ms ease;
}
.stButton > button[kind="primary"], .stButton > button[data-testid="stBaseButton-primary"],
.stButton > button[data-testid="baseButton-primary"] {
    background: var(--accent) !important; border: 1px solid var(--accent) !important; color: #052E16 !important;
}
.stButton > button[kind="primary"]:hover, .stButton > button[data-testid="stBaseButton-primary"]:hover,
.stButton > button[data-testid="baseButton-primary"]:hover {
    background: var(--accent-2) !important; border-color: var(--accent-2) !important;
}
.stButton > button[kind="secondary"], .stDownloadButton > button {
    background: rgba(148,163,184,.08) !important; color: var(--fg) !important; border: 1px solid var(--border) !important;
}
.stButton > button[kind="secondary"]:hover, .stDownloadButton > button:hover {
    border-color: rgba(148,163,184,.5) !important; background: rgba(148,163,184,.14) !important;
}
.stTextInput input { min-height: 44px; border-radius: 10px !important; font-family: "Fira Code", monospace; }
button:focus-visible, input:focus-visible, [role="radio"]:focus-visible, [role="tab"]:focus-visible {
    outline: 2px solid var(--accent-2) !important; outline-offset: 2px !important;
}

@media (prefers-reduced-motion: reduce) {
    * { transition: none !important; animation: none !important; }
}
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
# SIDEBAR — all controls live here
# ==============================================================
df_po_raw, df_pr_raw = None, None

with st.sidebar:
    st.markdown(f"""
        <div class="brand">
            <div class="brand-mark">{icon("package", 22)}</div>
            <div><div class="brand-name">DC Inward</div><div class="brand-sub">PO vs PR control room</div></div>
        </div>
    """, unsafe_allow_html=True)

    st.markdown('<div class="side-label">Source</div>', unsafe_allow_html=True)
    data_source = st.radio("Data mode", options=["Upload Files", "Live Zoho API Sync"], index=1,
                           label_visibility="collapsed")

    st.markdown('<div class="side-label">PR lookback window</div>', unsafe_allow_html=True)
    selected_window = st.radio("PR lookback window", options=["Last 15 Hours", "Last 24 Hours"],
                               index=0, horizontal=True, label_visibility="collapsed")
    hours_window = 15 if selected_window == "Last 15 Hours" else 24

    st.markdown('<div class="side-label">Data</div>', unsafe_allow_html=True)
    if data_source == "Upload Files":
        po_file = st.file_uploader("PO monthly data (Excel)", type=['xlsx', 'xls'])
        pr_file = st.file_uploader("PR monthly data (Excel)", type=['xlsx', 'xls'])
        if po_file and pr_file:
            df_po_raw = pd.read_excel(po_file)
            df_pr_raw = pd.read_excel(pr_file)
    else:
        if st.button(":material/sync: Fetch live data", type="primary", width="stretch"):
            try:
                with st.spinner("Fetching PO & PR line items from Zoho..."):
                    df_po_raw, df_pr_raw = fetch_zoho_data_live()
                    st.session_state["raw_po"] = df_po_raw
                    st.session_state["raw_pr"] = df_pr_raw
                    st.session_state["fetched_at"] = datetime.now()
                    st.success("Live data fetched.", icon=":material/check_circle:")
            except Exception as e:
                st.error(f"Zoho API error: {e}", icon=":material/error:")

        if "raw_po" in st.session_state and "raw_pr" in st.session_state:
            df_po_raw = st.session_state["raw_po"]
            df_pr_raw = st.session_state["raw_pr"]

    st.markdown('<div class="side-label">Google Sheet</div>', unsafe_allow_html=True)
    sheet_tab_name = st.text_input("Damage tab name", value="Damage_Short")

    fetched_at = st.session_state.get("fetched_at")
    st.markdown(
        f'<div class="side-note">Damage rows push to the <b>{sheet_tab_name}</b> tab. '
        f'Rows already in the sheet are skipped.'
        + (f'<br>Last fetch: <code>{fetched_at.strftime("%d-%b %H:%M")}</code>' if fetched_at else '')
        + '</div>',
        unsafe_allow_html=True
    )



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
# MAIN LAYOUT
# ==============================================================
CHART_TEXT = "#CBD5E1"
CHART_GRID = "rgba(148,163,184,0.15)"

def hbar(df, label_col, value_col, color, title, value_fmt=",.0f", ascending=False):
    """Horizontal bar chart with value labels, styled for the dark theme."""
    height = max(160, 34 * len(df))
    base = alt.Chart(df).encode(
        y=alt.Y(f"{label_col}:N", sort="x" if ascending else "-x", title=None,
                axis=alt.Axis(labelColor=CHART_TEXT, labelFontSize=12, labelLimit=260, ticks=False, domain=False)),
        x=alt.X(f"{value_col}:Q", title=None,
                axis=alt.Axis(labelColor=CHART_TEXT, gridColor=CHART_GRID, domain=False, ticks=False, format=value_fmt)),
        tooltip=[alt.Tooltip(f"{label_col}:N"), alt.Tooltip(f"{value_col}:Q", format=value_fmt)],
    )
    bars = base.mark_bar(color=color, cornerRadiusEnd=5, height=18)
    labels = base.mark_text(align="left", dx=6, color="#F8FAFC", font="Fira Code", fontSize=12).encode(
        text=alt.Text(f"{value_col}:Q", format=value_fmt)
    )
    return (bars + labels).properties(
        height=height, title=alt.TitleParams(title, color="#F8FAFC", fontSize=14, anchor="start", font="Fira Sans")
    ).configure_view(strokeWidth=0).configure(background="transparent")


has_data = "processed_df" in st.session_state
anchor_time = st.session_state.get("anchor_time", None) if has_data else None
active_window = st.session_state.get("window_hours", hours_window)

if has_data and anchor_time is not None and pd.notna(anchor_time):
    status_html = (f'<span class="dot"></span>Last PR <code>{anchor_time.strftime("%d-%b %H:%M")}</code>'
                   f' · {active_window}h window')
else:
    status_html = '<span class="dot off"></span>Not synced yet'

st.markdown(f"""
<div class="hero">
    <div>
        <div class="eyebrow">DC INWARD · ZOHO INVENTORY</div>
        <h1>PO vs PR Control Room</h1>
        <p>Today's receipts against open purchase orders, plus damage &amp; short tracking.</p>
    </div>
    <div class="status">{status_html}</div>
</div>
""", unsafe_allow_html=True)

if not has_data:
    st.markdown("""
    <div class="empty">
        <div class="t">No data loaded</div>
        Use <b>Fetch live data</b> in the sidebar, or switch to <b>Upload Files</b> and add the PO and PR exports.
    </div>
    """, unsafe_allow_html=True)
    st.stop()

display_df = st.session_state["processed_df"]
total_pos_cnt, total_po, total_pr, total_prmtd, total_pending, total_fr_val = st.session_state["kpi_metrics"]
df_damage = st.session_state.get("raw_damage", pd.DataFrame())
dmg_records = len(df_damage)
dmg_qty = int(pd.to_numeric(df_damage.get("Qty", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()) if dmg_records else 0
total_received = total_pr + total_prmtd
fr_bar = max(0.0, min(100.0, float(total_fr_val)))

# ---------- Bento KPIs ----------
st.markdown(f"""
<div class="bento">
    <div class="card hero-kpi">
        <div>
            <div class="k-top"><span class="k-label">Fill rate</span><span class="k-icon">{icon("gauge")}</span></div>
            <div class="k-value">{total_fr_val:.1f}<small>%</small></div>
            <div class="k-sub">Received {total_received:,} of {total_po:,} ordered across {total_pos_cnt:,} POs</div>
        </div>
        <div class="meter" role="img" aria-label="Fill rate {total_fr_val:.1f} percent">
            <div class="meter-track"><div class="meter-fill" style="width:{fr_bar:.1f}%"></div></div>
            <div class="meter-legend"><span>0%</span><span>50%</span><span>100%</span></div>
        </div>
    </div>
    <div class="card">
        <div class="k-top"><span class="k-label">POs touched</span><span class="k-icon">{icon("file")}</span></div>
        <div class="k-value">{total_pos_cnt:,}</div><div class="k-sub">with a PR in this window</div>
    </div>
    <div class="card">
        <div class="k-top"><span class="k-label">PO qty</span><span class="k-icon">{icon("package")}</span></div>
        <div class="k-value">{total_po:,}</div><div class="k-sub">ordered</div>
    </div>
    <div class="card blue">
        <div class="k-top"><span class="k-label">Today PR qty</span><span class="k-icon">{icon("truck")}</span></div>
        <div class="k-value">{total_pr:,}</div><div class="k-sub">last {active_window} hours</div>
    </div>
    <div class="card">
        <div class="k-top"><span class="k-label">PRMTD qty</span><span class="k-icon">{icon("calendar")}</span></div>
        <div class="k-value">{total_prmtd:,}</div><div class="k-sub">received earlier</div>
    </div>
    <div class="card warn">
        <div class="k-top"><span class="k-label">Pending qty</span><span class="k-icon">{icon("clock")}</span></div>
        <div class="k-value">{total_pending:,}</div><div class="k-sub">still to receive</div>
    </div>
    <div class="card danger">
        <div class="k-top"><span class="k-label">Damage / short</span><span class="k-icon">{icon("alert")}</span></div>
        <div class="k-value">{dmg_records:,}</div><div class="k-sub">{dmg_qty:,} units flagged</div>
    </div>
</div>
""", unsafe_allow_html=True)

# Table data without the Total row (totals live in the KPI cards and footer chips)
recon_df = display_df[display_df["Vendor Name"] != "Total"].copy()
recon_df["Fill %"] = pd.to_numeric(recon_df["PO FR %"].astype(str).str.rstrip("%"), errors="coerce").fillna(0)
for c in ["PO Qty", "Today PR Qty", "PRMTD", "Pending Qty"]:
    recon_df[c] = pd.to_numeric(recon_df[c], errors="coerce").fillna(0).astype(int)

tab_recon, tab_damage, tab_insights = st.tabs([
    ":material/table_chart: Reconciliation",
    f":material/report: Damage & short ({dmg_records})",
    ":material/insights: Insights",
])

# ---------- Tab 1: Reconciliation ----------
with tab_recon:
    f1, f2 = st.columns([3, 1])
    with f1:
        query = st.text_input("Search PO, PR or vendor", placeholder="e.g. HC/26/09191 or Greenply",
                              label_visibility="collapsed")
    with f2:
        only_pending = st.toggle("Only pending", value=False)

    view = recon_df
    if query:
        q = query.strip().lower()
        view = view[view[["PO No", "PR No", "Vendor Name"]].astype(str)
                    .apply(lambda col: col.str.lower().str.contains(q, regex=False)).any(axis=1)]
    if only_pending:
        view = view[view["Pending Qty"] > 0]

    cols = ["Sl.no", "Date", "PO No", "PR No", "Vendor Name", "PO Qty", "Today PR Qty", "PRMTD", "Pending Qty", "Fill %"]
    styler = view[cols].style
    styler = (styler.map if hasattr(styler, "map") else styler.applymap)(
        lambda v: "color: #FBBF24; font-weight: 600;" if isinstance(v, (int, np.integer)) and v > 0 else "color: #94A3B8;",
        subset=["Pending Qty"]
    )
    st.dataframe(
        styler,
        hide_index=True,
        width="stretch",
        height=min(560, 38 * (len(view) + 1) + 4),
        column_config={
            "Sl.no": st.column_config.TextColumn("#", width="small"),
            "Date": st.column_config.TextColumn("Date", width="small"),
            "Vendor Name": st.column_config.TextColumn("Vendor", width="medium"),
            "PO Qty": st.column_config.NumberColumn("PO qty", format="%d"),
            "Today PR Qty": st.column_config.NumberColumn("Today PR", format="%d"),
            "PRMTD": st.column_config.NumberColumn("PRMTD", format="%d"),
            "Pending Qty": st.column_config.NumberColumn("Pending", format="%d"),
            "Fill %": st.column_config.ProgressColumn("Fill rate", min_value=0, max_value=100, format="%d%%"),
        },
    )

    st.markdown(f"""
    <div class="chips">
        <span class="chip"><b>Showing</b>{len(view):,} of {len(recon_df):,} POs</span>
        <span class="chip"><b>PO</b>{total_po:,}</span>
        <span class="chip"><b>Today PR</b>{total_pr:,}</span>
        <span class="chip"><b>PRMTD</b>{total_prmtd:,}</span>
        <span class="chip amber"><b>Pending</b>{total_pending:,}</span>
        <span class="chip"><b>Fill rate</b>{int(round(total_fr_val))}%</span>
    </div>
    """, unsafe_allow_html=True)

    st.download_button(
        label=":material/download: Download summary (.xlsx)",
        data=generate_excel_file(display_df),
        file_name=f"PO_PR_Summary_{active_window}h.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# ---------- Tab 2: Damage & short ----------
with tab_damage:
    if df_damage.empty:
        st.markdown(f"""
        <div class="empty">
            <div class="t">No damage or short PRs</div>
            Nothing was flagged in the last {active_window} hours.
        </div>
        """, unsafe_allow_html=True)
    else:
        target_cols = ["Date", "vendor name", "PO NO", "Item", "Qty", "reason", "PR done on"]
        for c in target_cols:
            if c not in df_damage.columns:
                df_damage[c] = ""
        df_damage_display = df_damage[target_cols]

        qty_num = pd.to_numeric(df_damage_display["Qty"], errors="coerce").fillna(0)
        reason_l = df_damage_display["reason"].astype(str).str.lower()
        damage_units = int(qty_num[reason_l == "damage"].sum())
        short_units = int(qty_num[reason_l == "shortage"].sum())
        vendors = df_damage_display["vendor name"].nunique()

        st.markdown(f"""
        <div class="chips">
            <span class="chip"><b>Records</b>{len(df_damage_display):,}</span>
            <span class="chip"><b>Vendors</b>{vendors:,}</span>
            <span class="chip red"><b>Damage units</b>{damage_units:,}</span>
            <span class="chip amber"><b>Short units</b>{short_units:,}</span>
        </div>
        """, unsafe_allow_html=True)

        by_vendor = (df_damage_display.assign(Units=qty_num)
                     .groupby("vendor name", as_index=False)["Units"].sum()
                     .sort_values("Units", ascending=False).head(8)
                     .rename(columns={"vendor name": "Vendor"}))
        st.altair_chart(hbar(by_vendor, "Vendor", "Units", "#F87171", "Units flagged by vendor"),
                        width="stretch")

        dstyler = df_damage_display.style
        dstyler = (dstyler.map if hasattr(dstyler, "map") else dstyler.applymap)(
            lambda v: "color: #F87171; font-weight: 600;" if str(v).lower() == "damage"
            else ("color: #FBBF24; font-weight: 600;" if str(v).lower() == "shortage" else ""),
            subset=["reason"]
        )
        st.dataframe(
            dstyler, hide_index=True, width="stretch",
            height=min(480, 38 * (len(df_damage_display) + 1) + 4),
            column_config={
                "vendor name": st.column_config.TextColumn("Vendor"),
                "PO NO": st.column_config.TextColumn("PO no"),
                "Item": st.column_config.TextColumn("Item", width="large"),
                "Qty": st.column_config.NumberColumn("Qty", format="%d"),
                "reason": st.column_config.TextColumn("Reason"),
                "PR done on": st.column_config.TextColumn("PR done on"),
            },
        )

        a1, a2, _ = st.columns([1, 1, 2])
        with a1:
            st.download_button(
                label=":material/download: Damage report (.xlsx)",
                data=generate_damage_excel(df_damage_display),
                file_name=f"Damage_Report_{active_window}h.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width="stretch",
            )
        with a2:
            push = st.button(":material/cloud_upload: Push to Google Sheet", type="primary", width="stretch")

        if push:
            try:
                webhook_url = st.secrets.get("google", {}).get("webhook_url", "")
                if not webhook_url:
                    st.error("Add `webhook_url` under `[google]` in Streamlit Secrets.", icon=":material/error:")
                else:
                    with st.spinner(f"Pushing to '{sheet_tab_name}'..."):
                        payload = {
                            "sheet_name": sheet_tab_name,
                            "rows": df_damage_display.to_dict(orient="records")
                        }
                        res = requests.post(
                            webhook_url,
                            data=json.dumps(payload),
                            headers={"Content-Type": "text/plain;charset=utf-8"},
                            timeout=20
                        )
                        try:
                            result_json = res.json()
                            if result_json.get("status") == "success":
                                st.success(
                                    f"Added {result_json.get('appended', 0)} row(s), skipped "
                                    f"{result_json.get('skipped', 0)} duplicate(s) in '{sheet_tab_name}'.",
                                    icon=":material/check_circle:"
                                )
                            else:
                                st.error(f"Apps Script error: {result_json.get('message')}", icon=":material/error:")
                        except Exception:
                            st.error("Webhook returned HTML instead of JSON. Deploy the Apps Script as a Web App with 'Who has access: Anyone'.",
                                     icon=":material/error:")
            except Exception as e:
                st.error(f"Sync error: {e}", icon=":material/error:")

# ---------- Tab 3: Insights ----------
with tab_insights:
    c1, c2 = st.columns(2)
    with c1:
        top_pending = (recon_df[recon_df["Pending Qty"] > 0]
                       .sort_values("Pending Qty", ascending=False).head(10)
                       .assign(PO=lambda d: d["PO No"].astype(str) + " · " + d["Vendor Name"].astype(str).str.slice(0, 22)))
        if top_pending.empty:
            st.markdown('<div class="empty"><div class="t">Nothing pending</div>Every PO in this window is fully received.</div>',
                        unsafe_allow_html=True)
        else:
            st.altair_chart(hbar(top_pending[["PO", "Pending Qty"]], "PO", "Pending Qty", "#FBBF24",
                                 "Top 10 POs by pending qty"), width="stretch")
    with c2:
        vend = (recon_df.groupby("Vendor Name", as_index=False)
                .agg(PO=("PO Qty", "sum"), Rec=("Today PR Qty", "sum"), Prior=("PRMTD", "sum")))
        vend = vend[vend["PO"] > 0]
        if vend.empty:
            st.markdown('<div class="empty"><div class="t">No PO quantities</div>Fill rate needs PO data.</div>',
                        unsafe_allow_html=True)
        else:
            vend["Fill rate"] = ((vend["Rec"] + vend["Prior"]) / vend["PO"] * 100).clip(upper=100).round(0)
            low = vend.sort_values("Fill rate").head(10).rename(columns={"Vendor Name": "Vendor"})
            low["Vendor"] = low["Vendor"].astype(str).str.slice(0, 30)
            st.altair_chart(hbar(low[["Vendor", "Fill rate"]], "Vendor", "Fill rate", "#60A5FA",
                                 "Lowest fill rate by vendor (%)", value_fmt=".0f", ascending=True), width="stretch")
