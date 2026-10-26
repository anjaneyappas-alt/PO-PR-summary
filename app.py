import streamlit as st
import pandas as pd
import numpy as np
import io
import requests
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor, as_completed

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Page Config
st.set_page_config(page_title="PO vs PR Summary", page_icon="📊", layout="wide")

# Excel Grid Styling + High-Contrast KPI Cards
st.markdown("""
    <style>
        #MainMenu {visibility: hidden;}
        header {visibility: hidden;}
        footer {visibility: hidden;}
        [data-testid="stHeader"] {display: none;}
        
        .main {
            background-color: #ffffff !important;
        }

        [data-testid="stMetric"] {
            background-color: #f8f9fa !important;
            border: 1px solid #dee2e6 !important;
            padding: 12px 16px !important;
            border-radius: 8px !important;
            box-shadow: 0 1px 3px rgba(0,0,0,0.05) !important;
        }

        [data-testid="stMetricValue"] {
            font-size: 22px !important;
            font-weight: 700 !important;
            color: #111111 !important;
        }
        
        [data-testid="stMetricLabel"] {
            font-size: 13px !important;
            font-weight: 600 !important;
            color: #495057 !important;
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
        
        .stTable th {
            background-color: #e6e6e6 !important;
            color: #000000 !important;
            font-weight: bold !important;
            text-align: center !important;
            border: 1px solid #bfbfbf !important;
            padding: 6px 10px !important;
            white-space: nowrap !important;
        }
        
        .stTable td {
            border: 1px solid #d9d9d9 !important;
            padding: 6px 10px !important;
            color: #000000 !important;
            white-space: nowrap !important;
        }
        
        .stTable tr:nth-child(even) {
            background-color: #f9f9f9 !important;
        }
    </style>
""", unsafe_allow_html=True)

st.title("📊 PO vs PR Excel View")

# Helper for robust column lookup
def find_first_existing_col(df, candidates, default=None):
    for col in candidates:
        if col in df.columns:
            return col
    return default

# Control Panel
control_col1, control_col2 = st.columns([1, 2])
with control_col1:
    data_source = st.radio("📡 Data Mode:", options=["Upload Files", "Live Zoho API Sync"], horizontal=True)
with control_col2:
    selected_window = st.radio(
        "⏱ Select PR Lookback Window:",
        options=["Last 15 Hours", "Last 24 Hours"],
        horizontal=True,
        index=0
    )

hours_window = 15 if selected_window == "Last 15 Hours" else 24


# --- ZOHO API HELPER FUNCTIONS ---
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
    headers = {
        "Content-Type": "application/x-www-form-urlencoded"
    }
    response = requests.post(url, data=payload, headers=headers)
    res_data = response.json()
    if "access_token" in res_data:
        return res_data["access_token"]
    else:
        raise Exception(f"Failed to refresh Zoho Token: {res_data}")

def fetch_po_by_number(po_number, headers, domain, org_id):
    """Directly fetch a PO by its exact number and calculate line-item quantities."""
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
    """Fetch individual PR to retrieve item-level line items (SKU, Item Name, etc.)."""
    try:
        url = f"https://www.zohoapis.{domain}/inventory/v1/purchasereceives/{pr_id}"
        res = requests.get(url, headers=headers, params={"organization_id": org_id}).json()
        if "purchasereceive" in res:
            return res["purchasereceive"]
    except Exception:
        pass
    return None

def fetch_zoho_data_live():
    access_token = get_zoho_access_token()
    org_id = st.secrets["zoho"]["organization_id"]
    domain = st.secrets["zoho"].get("domain", "zoho.in")

    headers = {"Authorization": f"Zoho-oauthtoken {access_token}"}
    date_30_days_ago = (datetime.now() - timedelta(days=30)).strftime("%Y-%m-%d")

    # 1. Fetch PRs from last 30 days
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

    # Check for Damage PRs and pull line-item details
    damage_prs = [
        pr for pr in pr_list 
        if "damage" in str(pr.get("purchasereceive_number", "")).lower()
    ]

    detailed_damage_items = []
    if damage_prs:
        with ThreadPoolExecutor(max_
