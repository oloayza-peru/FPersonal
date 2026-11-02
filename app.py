import pandas as pd
import streamlit as st
import openpyxl
import io
from datetime import datetime
import os

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(
    page_title="SAP Fiori - Control de Ingresos y Egresos",
    page_icon="📊",
    layout="wide"
)

# --- ENLACE AL ARCHIVO EXCEL ONLINE EN ONEDRIVE ---
ONEDRIVE_URL = "https://1drv.ms/x/c/f2550978c0f7f8d5/IQCbpvC_igOlS4veC_-Pl1fPAdhDS8b5yrSeblI8uyQAIog?e=b82J1A"
EXCEL_FILE = "Control_Financiero.xlsx"

# --- INICIALIZACIÓN DEL ARCHIVO EXCEL SI NO EXISTE ---
def init_excel():
    if not os.path.exists(EXCEL_FILE):
        wb = openpyxl.Workbook()
        ws_ing = wb.active
        ws_ing.title = "Ingresos"
        ws_ing.append(["Doc #", "Año", "Mes", "Tipo", "Descripción", "Monto", "Moneda"])
        
        ws_egr = wb.create_sheet(title="Egresos")
        ws_egr.append(["Doc #", "Código", "Año", "Mes", "Tipo", "Descripción", "Monto", "Moneda"])
        wb.save(EXCEL_FILE)

init_excel()

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
    st.html(css)

cargar_estilos_sap()

# --- CONSTANTES Y FECHA ACTUAL ---
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
    try:
        df_egr = pd.read_excel(EXCEL_FILE, sheet_name="Egresos")
        if df_egr.empty or 'Código' not in df_egr.columns or df_egr['Código'].dropna().empty:
            return "E001"
        else:
            codigos = df_egr['Código'].astype(str).str.replace('E', '').astype(int)
            siguiente = codigos.max() + 1
            return f"E{siguiente:03d}"
    except Exception:
        return "E001"

# --- CABECERA SAP FIORI CON ENLACE A ONEDRIVE ---
shell_bar_html = f"""
<div class="sap-shell-bar">
    <div style="display: flex; align-items: center; gap: 10px; flex-wrap: wrap;">
        <span class="sap-shell-title">SAP S/4HANA — Financial Control & Treasury Management</span>
        <a href="{ONEDRIVE_URL}" target="_blank" class="sap-drive-link">☁️ Abrir Control_Financiero.xlsx (OneDrive)</a>
    </div>
    <div class="sap-shell-badge">SAP Horizon</div>
</div>
"""
st.html(shell_bar_html)

# Botón de descarga general de Excel consolidado
col_down_1, col_down_2 = st.columns([3, 1])
with col_down_2:
    if os.path.exists(EXCEL_FILE):
        with open(EXCEL_FILE, "rb") as f:
            excel_bytes = f.read()
        st.download_button(
            label="📥 Descargar Excel Actualizado",
            data=excel_bytes,
            file_name="Control_Financiero.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

tabs = st.tabs(["💵 Gestor de Ingresos", "💸 Gestor de Egresos", "📊 Dashboard Analítico"])

# ==========================================
# PESTAÑA 1: INGRESOS
# ==========================================
with tabs[0]:
    st.subheader("Entradas de Dinero (Ingresos)")
    
    with st.expander("➕ Formulario de Entrada de Ingreso", expanded=True):
        with st.form(key="form_ingresos", clear_on_submit=True):
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                anio_ing = st.selectbox("Año Fiscal", ANIOS, index=ANIO_ACTUAL_INDEX, key="ing_anio")
                tipo_ing = st.selectbox("Clase de Ingreso", TIPOS_INGRESO, key="ing_tipo")
            with col2:
                mes_ing = st.selectbox("Período / Mes", MESES, index=MES_ACTUAL_INDEX, key="ing_mes")
                moneda_ing = st.selectbox("Moneda Base", MONEDAS, key="ing_moneda")
            with col3:
                desc_ing = st.text_input("Descripción del Asiento", key="ing_desc")
            with col4:
                monto_ing = st.number_input("Monto", min_value=0.0, format="%.2f", key="ing_monto")
            
            st.markdown("---")
            btn_guardar_ing = st.form_submit_button("Guardar Ingreso")
            
            if btn_guardar_ing:
                if desc_ing and monto_ing > 0:
                    df_ing = pd.read_excel(EXCEL_FILE, sheet_name="Ingresos")
                    nuevo_id = int(df_ing['Doc #'].max() + 1) if not df_ing.empty and 'Doc #' in df_ing.columns and not df_ing['Doc #'].dropna().empty else 1
                    
                    nuevo_reg = pd.DataFrame([{
                        'Doc #': nuevo_id,
                        'Año': anio_ing,
                        'Mes': mes_ing,
                        'Tipo': tipo_ing,
                        'Descripción': desc_ing,
                        'Monto': monto_ing,
                        'Moneda': moneda_ing
                    }])
                    
                    df_ing = pd.concat([df_ing, nuevo_reg], ignore_index=True)
                    
                    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                        df_ing.to_excel(writer, sheet_name='Ingresos', index=False)
                        pd.read_excel(EXCEL_FILE, sheet_name='Egresos').to_excel(writer, sheet_name='Egresos', index=False)
                    
                    st.success("Ingreso registrado y actualizado automáticamente en el Excel con éxito.")
                    st.rerun()
                else:
                    st.warning("Verifique que la descripción y el monto sean válidos.")

    st.markdown("#### Registros Contabilizados")
    try:
        df_ingresos = pd.read_excel(EXCEL_FILE, sheet_name="Ingresos")
    except Exception:
        df_ingresos = pd.DataFrame(columns=['Doc #', 'Año', 'Mes', 'Tipo', 'Descripción', 'Monto', 'Moneda'])

    if not df_ingresos.empty:
        st.dataframe(df_ingresos, use_container_width=True)

        st.markdown("##### ✏️ Modificación / Eliminación de Registro de Ingreso")
        col_ed1, col_ed2, col_ed3 = st.columns([2, 1, 1])
        doc_ids = df_ingresos['Doc #'].tolist() if 'Doc #' in df_ingresos.columns else []
        id_ing_sel = col_ed1.selectbox("Seleccionar Documento (Doc #)", doc_ids, key="select_ing_edit")
        
        if col_ed2.button("✏️ Cargar para Editar", key="btn_load_ing"):
            st.session_state.edit_ingreso_id = id_ing_sel

        if col_ed3.button("🗑️ Eliminar Registro", key="btn_del_ing"):
            df_ingresos = df_ingresos[df_ingresos['Doc #'] != id_ing_sel]
            with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df_ingresos.to_excel(writer, sheet_name='Ingresos', index=False)
                pd.read_excel(EXCEL_FILE, sheet_name='Egresos').to_excel(writer, sheet_name='Egresos', index=False)
            if st.session_state.edit_ingreso_id == id_ing_sel:
                st.session_state.edit_ingreso_id = None
            st.success(f"Ingreso Doc #{id_ing_sel} eliminado y actualizado en el Excel.")
            st.rerun()

        if st.session_state.edit_ingreso_id:
            row_ing = df_ingresos[df_ingresos['Doc #'] == st.session_state.edit_ingreso_id].iloc[0]
            with st.expander(f"📝 Editando Ingreso Doc #{st.session_state.edit_ingreso_id}", expanded=True):
                with st.form(key="form_edit_ingreso"):
                    c1, c2, c3, c4 = st.columns(4)
                    anio_e = c1.selectbox("Año Fiscal", ANIOS, index=ANIOS.index(int(row_ing['Año'])) if int(row_ing['Año']) in ANIOS else 0)
                    mes_e = c2.selectbox("Período / Mes", MESES, index=MESES.index(row_ing['Mes']) if row_ing['Mes'] in MESES else 0)
                    tipo_e = c3.selectbox("Clase de Ingreso", TIPOS_INGRESO, index=TIPOS_INGRESO.index(row_ing['Tipo']) if row_ing['Tipo'] in TIPOS_INGRESO else 0)
                    moneda_e = c4.selectbox("Moneda Base", MONEDAS, index=MONEDAS.index(row_ing['Moneda']) if row_ing['Moneda'] in MONEDAS else 0)

                    c5, c6 = st.columns([3, 1])
                    desc_e = c5.text_input("Descripción", value=row_ing['Descripción'])
                    monto_e = c6.number_input("Monto", value=float(row_ing['Monto']), min_value=0.0, format="%.2f")

                    btn_actualizar = st.form_submit_button("💾 Actualizar Registro")
                    
                    if btn_actualizar:
                        df_ingresos.loc[df_ingresos['Doc #'] == st.session_state.edit_ingreso_id, ['Año', 'Mes', 'Tipo', 'Descripción', 'Monto', 'Moneda']] = [anio_e, mes_e, tipo_e, desc_e, monto_e, moneda_e]
                        with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                            df_ingresos.to_excel(writer, sheet_name='Ingresos', index=False)
                            pd.read_excel(EXCEL_FILE, sheet_name='Egresos').to_excel(writer, sheet_name='Egresos', index=False)
                        st.session_state.edit_ingreso_id = None
                        st.success("Registro actualizado y sincronizado en el Excel.")
                        st.rerun()

        col_sum1, col_sum2 = st.columns(2)
        monto_pen = df_ingresos[df_ingresos['Moneda'] == 'PEN']['Monto'].sum() if 'Moneda' in df_ingresos.columns else 0
        monto_usd = df_ingresos[df_ingresos['Moneda'] == 'USD']['Monto'].sum() if 'Moneda' in df_ingresos.columns else 0
        
        col_sum1.metric("Acumulado Ingresos (PEN)", f"S/ {monto_pen:,.2f}")
        col_sum2.metric("Acumulado Ingresos (USD)", f"$ {monto_usd:,.2f}")
    else:
        st.info("Sin registros de ingresos en el archivo.")

# ==========================================
# PESTAÑA 2: EGRESOS
# ==========================================
with tabs[1]:
    st.subheader("Registro de Documentos de Egreso (Orden de Gastos)")
    
    siguiente_codigo = obtener_siguiente_codigo()
    
    with st.expander("➕ Registro de Partidas de Egreso (Fijas y Adicionales)", expanded=True):
        col_h1, col_h2, col_h3 = st.columns(3)
        anio_egr = col_h1.selectbox("Ejercicio (Año)", ANIOS, index=ANIO_ACTUAL_INDEX, key="egr_anio")
        mes_egr = col_h2.selectbox("Período Mensual", MESES, index=MES_ACTUAL_INDEX, key="egr_mes")
        codigo_egr = col_h3.text_input("N° de Registro Correlativo", value=siguiente_codigo, disabled=True)

        st.markdown("---")
        st.button("➕ Añadir Registro Adicional (No Fijo)", on_click=agregar_fila_adicional)
        
        with st.form(key="form_egresos"):
            gastos_input = []
            
            st.markdown("**Partidas Fijas de Gastos**")
            for: i, (cat, desc) in enumerate
