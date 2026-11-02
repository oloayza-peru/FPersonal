import pandas as pd
import streamlit as st
import sqlite3
import io
from datetime import datetime
from github import Github

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(
    page_title="SAP Fiori - Control de Ingresos y Egresos",
    page_icon="📊",
    layout="wide"
)

# --- ENLACES A GOOGLE DRIVE Y GOOGLE SHEETS ---
DRIVE_FOLDER_URL = "https://drive.google.com/drive/folders/1Mqp3Lf-U2PC5BXPCr949R4y3VbSBVk_B?usp=drive_link"
DRIVE_FOLDER_NAME = "Carpeta Personal"
SHEETS_REGISTROS_URL = "https://docs.google.com/spreadsheets/d/10nh75p-uzJAsXjVJxnMcicEul7pvw4TiZB9tgHExlvA/edit?usp=sharing"

# --- ESTILOS CSS PERSONALIZADOS (SAP FIORI / HORIZON THEME) ---
def cargar_estilos_sap():
    css = """
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Open+Sans:wght@400;600;700&display=swap');
        
        html, body, [class*="css"] {
            font-family: 'Open Sans', '72', 'Segoe UI', Arial, sans-serif !important;
        }

        .stApp {
            background-color: #F5F6F7 !important;
        }

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
        .sap-drive-link {
            color: #FFFFFF !important;
            background-color: rgba(255, 255, 255, 0.15);
            padding: 4px 10px;
            border-radius: 4px;
            font-size: 0.85rem;
            text-decoration: none !important;
            font-weight: 600;
            transition: background-color 0.2s;
            margin-left: 8px;
        }
        .sap-drive-link:hover {
            background-color: rgba(255, 255, 255, 0.3);
        }

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

        [data-testid="stForm"], div.element-container > div.stExpander, [data-testid="stMetric"] {
            background-color: #FFFFFF !important;
            border: 1px solid #D9D9D9 !important;
            border-radius: 6px !important;
            padding: 12px 16px !important;
            box-shadow: 0 1px 4px rgba(0,0,0,0.06) !important;
        }

        .stButton > button, div[data-testid="stFormSubmitButton"] > button, .stDownloadButton > button {
            background-color: #0070F2 !important;
            color: #FFFFFF !important;
            font-weight: 600 !important;
            border: 1px solid #0070F2 !important;
            border-radius: 4px !important;
            padding: 8px 20px !important;
            transition: background-color 0.2s ease, border-color 0.2s ease;
        }

        .stButton > button:hover, div[data-testid="stFormSubmitButton"] > button:hover, .stDownloadButton > button:hover {
            background-color: #0056B3 !important;
            border-color: #0056B3 !important;
        }

        [data-testid="stDataFrame"] {
            background-color: #FFFFFF;
            border: 1px solid #D9D9D9;
            border-radius: 6px;
        }
    </style>
    """
    st.markdown(css, unsafe_allow_html=True)

cargar_estilos_sap()

# --- FUNCION PARA AUTO-COMMIT EN GITHUB ---
def sincronizar_con_github(nombre_archivo="respaldo_finanzas.csv"):
    if "GITHUB_TOKEN" in st.secrets and "GITHUB_REPO" in st.secrets:
        try:
            token = st.secrets["GITHUB_TOKEN"]
            repo_name = st.secrets["GITHUB_REPO"]
            g = Github(token)
            repo = g.get_repo(repo_name)
            
            with open(nombre_archivo, "r", encoding="utf-8") as f:
                contenido = f.read()
            
            commit_message = f"Auto-backup: Sincronización desde App ({datetime.now().strftime('%Y-%m-%d %H:%M:%S')})"
            
            try:
                contents = repo.get_contents(nombre_archivo)
                repo.update_file(
                    path=contents.path,
                    message=commit_message,
                    content=contenido,
                    sha=contents.sha
                )
                st.toast("☁️ Copia de seguridad guardada automáticamente en GitHub", icon="✅")
            except Exception:
                repo.create_file(
                    path=nombre_archivo,
                    message=commit_message,
                    content=contenido
                )
                st.toast("☁️ Archivo de respaldo creado en GitHub", icon="✅")
        except Exception as e:
            st.error(f"Error al sincronizar con GitHub: {e}")

# --- INICIALIZACIÓN DE BASE DE DATOS LOCAL ---
def init_db():
    conn = sqlite3.connect("finanzas.db")
    c = conn.cursor()
    c.execute("CREATE TABLE IF NOT EXISTS ingresos (id INTEGER PRIMARY KEY AUTOINCREMENT, anio INTEGER, mes TEXT, tipo TEXT, descripcion TEXT, monto REAL, moneda TEXT)")
    c.execute("CREATE TABLE IF NOT EXISTS egresos (id INTEGER PRIMARY KEY AUTOINCREMENT, codigo TEXT, anio INTEGER, mes TEXT, tipo TEXT, descripcion TEXT, monto REAL, moneda TEXT)")
    conn.commit()
    conn.close()

init_db()

# --- CONSTANTES Y FECHA ACTUAL DEL SISTEMA ---
MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", 
         "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
ANIOS = list(range(2026, 2031))
MONEDAS = ["PEN", "USD"]
TIPOS_INGRESO = ["Planilla", "Préstamo", "Extra", "Otro trabajo"]
TIPOS_EGRESO = ["Casa", "Servicios", "Préstamo", "Educación", "Otros"]

NOW = datetime.now()
MES_ACTUAL_INDEX = NOW.month - 1
ANIO_ACTUAL = NOW.year
ANIO_ACTUAL_INDEX = ANIOS.index(ANIO_ACTUAL) if ANIO_ACTUAL in ANIOS else 0

EGRESOS_FIJOS_BASE = [
    ("Casa", "Alquiler"),
    ("Casa", "Agua"),
    ("Casa", "Internet"),
    ("Casa", "Alimentos"),
    ("Casa", "Lonchera"),
    ("Casa", "Otros Casa"),
    ("Servicios", "Celular Entel"),
    ("Servicios", "BBVA")
]

if "cant_egresos_adicionales" not in st.session_state:
    st.session_state.cant_egresos_adicionales = 0

if "edit_ingreso_id" not in st.session_state:
    st.session_state.edit_ingreso_id = None

if "edit_egreso_codigo" not in st.session_state:
    st.session_state.edit_egreso_codigo = None

def agregar_fila_adicional():
    st.session_state.cant_egresos_adicionales += 1

def obtener_siguiente_codigo():
    conn = sqlite3.connect("finanzas.db")
    df = pd.read_sql_query("SELECT codigo FROM egresos", conn)
    conn.close()
    if df.empty or df['codigo'].dropna().empty:
        return "E001"
    else: