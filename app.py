import streamlit as st
import pandas as pd
import numpy as np
import io
import requests
from datetime import datetime, timedelta

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
