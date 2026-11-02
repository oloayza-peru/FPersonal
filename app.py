import pandas as pd
import streamlit as st
import sqlite3
from datetime import datetime

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(
    page_title="SAP Fiori - Control de Ingresos y Egresos",
    page_icon="📊",
    layout="wide"
)

# --- ESTILOS CSS PERSONALIZADOS (ESTILO SAP FIORI / S/4HANA) ---
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
        color: #556B82 !important;
        border-radius: 4px 4px 0 0 !important;
        padding: 10px 24px !important;
    }

    .stTabs [aria-selected="true"] {
        background-color: #EBEFF5 !important;
        color: #0070F2 !important;
        border-bottom: 3px solid #0070F2 !important;
    }

    /* Formularios estilo SAP Cards */
    [data-testid="stForm"], div.element-container > div.stExpander {
        background-color: #FFFFFF !important;
        border: 1px solid #D9D9D9 !important;
        border-radius: 6px !important;
        box-shadow: 0 1px 4px rgba(0,0,0,0.06) !important;
    }

    /* Tarjetas KPI Analíticas SAP */
    .sap-kpi-card {
        background-color: #FFFFFF;
        border: 1px solid #D9D9D9;
        border-left: 6px solid #0070F2;
        border-radius: 6px;
        padding: 18px 20px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    .sap-kpi-title {
        font-size: 0.8rem;
        color: #6A6D70;
        font-weight: 700;
        text-transform: uppercase;
        margin-bottom: 6px;
    }
    .sap-kpi-value {
        font-size: 1.7rem;
        font-weight: 700;
        color: #1D2D3E;
    }
    .sap-kpi-positive { border-left-color: #107C41 !important; }
    .sap-kpi-negative { border-left-color: #BB0000 !important; }

    /* Botones primarios SAP */
    .stButton > button, div[data-testid="stFormSubmitButton"] > button {
        background-color: #0070F2 !important;
        color: #FFFFFF !important;
        font-weight: 600 !important;
        border: 1px solid #0070F2 !important;
        border-radius: 4px !important;
        padding: 8px 20px !important;
        transition: background-color 0.2s ease, border-color 0.2s ease;
    }

    .stButton > button:hover, div[data-testid="stFormSubmitButton"] > button:hover {
        background
