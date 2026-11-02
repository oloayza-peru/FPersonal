import pandas as pd
import streamlit as st
import sqlite3
import io
from datetime import datetime

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(
    page_title="SAP Fiori - Control de Ingresos y Egresos",
    page_icon="📊",
    layout="wide"
)

# --- ENLACE A GOOGLE DRIVE ---
DRIVE_FOLDER_URL = "https://drive.google.com/drive/folders/1Mqp3Lf-U2PC5BXPCr949R4y3VbSBVk_B?usp=drive_link"
DRIVE_FOLDER_NAME = "01 Personal"

# --- ESTILOS CSS PERSONALIZADOS (SAP FIORI / HORIZON THEME) ---
SAP_STYLE = """
<style>
    @import url('https://fonts.googleapis.com/css2?family=Open+Sans:wght@400;600;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Open Sans', '72', 'Segoe UI', Arial, sans-serif !important;
    }

    .stApp {
        background-color: #F5F6F7 !important;
    }

    /* SAP Fiori ShellBar Superior */
    .sap-shell-bar {
        background: linear-gradient(90deg, #1D2D3E 0%, #354A5F 100%);
        color: #FFFFFF;
        padding: 12px 24px;
        margin-bottom: 24px;
        border-radius: 6px;
        box-shadow: 0 2px 6px rgba(0,0,0,0.2);
        display: flex;
        align-items: center;
        justify-content: space-between;
    }
    .sap-shell-title {
        font-size: 1.25rem;
        font-weight: 700;
        color: #FFFFFF;
        letter-spacing: 0.5px;
    }
    .sap-shell-badge {
        background-color: #F0AB00;
        color: #1D2D3E;
        font-weight: 800;
        font-size: 0.75rem;
        padding: 3px 8px;
        border-radius: 3px;
        text-transform: uppercase;
    }

    .sap-drive-link {
        color: #FFFFFF !important;
        background-color: rgba(255, 255, 255, 0.15);
        padding: 4px 10px;
        border-radius: 4px;
        font-size: 0.85rem;
        text-decoration: none !important;
        font-weight: 600;
        transition: background-color 0.2s;
    }
    .sap-drive-link:hover {
        background-color: rgba(255, 255, 255, 0.3);
    }

    /* Tabs Estilo SAP Fiori */
    .stTabs [data-baseweb="tab-list"] {
        background-color: #FFFFFF;
        border-bottom: 2px solid #D9D9D9;
        border-radius: 6px 6px 0 0;
        gap: 8px;
        padding: 6px 12px 0 12px;
    }

    .stTabs [data-baseweb="tab"] {
        font-weight: 600 !important;
        color: #556B
