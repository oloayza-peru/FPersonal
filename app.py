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

        /* Botones primarios SAP y Descargas */
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

# --- INICIALIZACIÓN DE BASE DE DATOS LOCAL ---
def init_db():
    conn = sqlite3.connect("finanzas.db")
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS ingresos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            anio INTEGER,
            mes TEXT,
            tipo TEXT,
            descripcion TEXT,
            monto REAL,
            moneda TEXT
        )
    ''')
    c.execute('''
        CREATE TABLE IF NOT EXISTS egresos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            codigo TEXT,
            anio INTEGER,
            mes TEXT,
            tipo TEXT,
            descripcion TEXT,
            monto REAL,
            moneda TEXT
        )
    ''')
    conn.commit()
    conn.close()

init_db()

# --- CONSTANTES Y LISTAS ---
MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", 
         "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
ANIOS = list(range(2026, 2031))
MONEDAS = ["PEN", "USD"]
TIPOS_INGRESO = ["Planilla", "Préstamo", "Extra", "Otro trabajo"]
TIPOS_EGRESO = ["Casa", "Servicios", "Préstamo", "Educación", "Otros"]

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

def agregar_fila_adicional():
    st.session_state.cant_egresos_adicionales += 1

def obtener_siguiente_codigo():
    conn = sqlite3.connect("finanzas.db")
    df = pd.read_sql_query("SELECT codigo FROM egresos", conn)
    conn.close()
    if df.empty or df['codigo'].dropna().empty:
        return "E001"
    else:
        codigos = df['codigo'].str.replace('E', '').astype(int)
        siguiente = codigos.max() + 1
        return f"E{siguiente:03d}"

def exportar_dataframe_a_excel(df: pd.DataFrame, sheet_name: str = "Registros") -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()

def exportar_respaldo_csv():
    conn = sqlite3.connect("finanzas.db")
    df_ing = pd.read_sql_query("SELECT 'INGRESO' AS Tipo_Registro, anio, mes, tipo, descripcion, monto, moneda FROM ingresos", conn)
    df_egr = pd.read_sql_query("SELECT 'EGRESO' AS Tipo_Registro, anio, mes, tipo, descripcion, monto, moneda FROM egresos", conn)
    conn.close()
    df_total = pd.concat([df_ing, df_egr], ignore_index=True)
    df_total.to_csv("respaldo_finanzas.csv", index=False)

def render_kpi(titulo: str, valor: str, estilo_clase: str = "sap-kpi-positive"):
    html_str = f'<div class="sap-kpi-card {estilo_clase}"><div class="sap-kpi-title">{titulo}</div><div class="sap-kpi-value">{valor}</div></div>'
    st.markdown(html_str, unsafe_allow_html=True)

# --- CABECERA SAP FIORI ---
st.markdown(
    f'''
    <div class="sap-shell-bar">
        <div style="display: flex; align-items: center; gap: 15px;">
            <span class="sap-shell-title">SAP S/4HANA — Financial Control & Treasury Management</span>
            <a href="{DRIVE_FOLDER_URL}" target="_blank" class="sap-drive-link">📁 Carpeta Drive ({DRIVE_FOLDER_NAME})</a>
        </div>
        <div class="sap-shell-badge">SAP Horizon</div>
    </div>
    ''',
    unsafe_allow_html=True
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
                anio_ing = st.selectbox("Año Fiscal", ANIOS, key="ing_anio")
                tipo_ing = st.selectbox("Clase de Ingreso", TIPOS_INGRESO, key="ing_tipo")
            with col2:
                mes_ing = st.selectbox("Período / Mes", MESES, key="ing_mes")
                moneda_ing = st.selectbox("Moneda Base", MONEDAS, key="ing_moneda")
            with col3:
                desc_ing = st.text_input("Descripción del Asiento", key="ing_desc")
            with col4:
                monto_ing = st.number_input("Monto", min_value=0.0, format="%.2f", key="ing_monto")
            
            st.markdown("---")
            btn_guardar_ing = st.form_submit_button("Guardar")
            
            if btn_guardar_ing:
                if desc_ing and monto_ing > 0:
                    conn = sqlite3.connect("finanzas.db")
                    c = conn.cursor()
                    c.execute(
                        "INSERT INTO ingresos (anio, mes, tipo, descripcion, monto, moneda) VALUES (?, ?, ?, ?, ?, ?)",
                        (anio_ing, mes_ing, tipo_ing, desc_ing, monto_ing, moneda_ing)
                    )
                    conn.commit()
                    conn.close()
                    exportar_respaldo_csv()
                    st.success("Ingreso guardado exitosamente.")
                    st.rerun()
                else:
                    st.warning("Verifique que la descripción y el monto sean válidos.")

    st.markdown("#### Registros Contabilizados")
    conn = sqlite3.connect("finanzas.db")
    df_ingresos = pd.read_sql_query("SELECT id AS 'Doc #', anio AS Año, mes AS Mes, tipo AS Tipo, descripcion AS Descripción, monto AS Monto, moneda AS Moneda FROM ingresos", conn)
    conn.close()

    if not df_ingresos.empty:
        excel_bytes_ing = exportar_dataframe_a_excel(df_ingresos, sheet_name="Ingresos")
        st.download_button(
            label="📥 Descargar Ingresos (Excel)",
            data=excel_bytes_ing,
            file_name=f"Ingresos_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.dataframe(df_ingresos, use_container_width=True)
        
        col_sum1, col_sum2 = st.columns(2)
        monto_pen = df_ingresos[df_ingresos['Moneda'] == 'PEN']['Monto'].sum()
        monto_usd = df_ingresos[df_ingresos['Moneda'] == 'USD']['Monto'].sum()
        
        with col_sum1:
            render_kpi("Acumulado Ingresos (PEN)", f"S/ {monto_pen:,.2f}", "sap-kpi-positive")

        with col_sum2:
            render_kpi("Acumulado Ingresos (USD)", f"$ {monto_usd:,.2f}", "sap-kpi-positive")
    else:
        st.info("Sin registros de ingresos en la base de datos.")

# ==========================================
# PESTAÑA 2: EGRESOS
# ==========================================
with tabs[1]:
    st.subheader("Registro de Documentos de Egreso (Orden de Gastos)")
    
    siguiente_codigo = obtener_siguiente_codigo()
    
    with st.expander("➕ Registro de Partidas de Egreso (Fijas y Adicionales)", expanded=True):
        col_h1, col_h2, col_h3 = st.columns(3)
        anio_egr = col_h1.selectbox("Ejercicio (Año)", ANIOS, key="egr_anio")
        mes_egr = col_h2.selectbox("Período Mensual", MESES, key="egr_mes")
        codigo_egr = col_h3.text_input("N° de Registro Correlativo", value=siguiente_codigo, disabled=True)

        st.markdown("---")
        st.button("➕ Añadir Registro Adicional (No Fijo)", on_click=agregar_fila_adicional)
        
        with st.form(key="form_egresos"):
            gastos_input = []
            
            st.markdown("**Partidas Fijas de Gastos**")
            for i, (cat, desc) in enumerate(EGRESOS_FIJOS_BASE):
                c1, c2, c3, c4, c5 = st.columns([2, 3, 2, 2, 1])
                c1.text_input(f"Categoría #{i+1}", value=cat, key=f"cat_{i}", disabled=True)
                c2.text_input(f"Concepto #{i+1}", value=desc, key=f"desc_{i}")
                monto = c3.number_input(f"Importe #{i+1}", min_value=0.0, format="%.2f", key=f"monto_{i}")
                moneda = c4.selectbox(f"Moneda #{i+1}", MONEDAS, key=f"moneda_{i}")
                incluir = c5.checkbox("Activo", value=True, key=f"inc_{i}")
                gastos_input.append((cat, desc, monto, moneda, incluir))

            if st.session_state.cant_egresos_adicionales > 0:
                st.markdown("**Partidas Adicionales (No Fijas)**")
                for j in range(st.session_state.cant_egresos_adicionales):
                    idx = len(EGRESOS_FIJOS_BASE) + j
                    c1, c2, c3, c4, c5 = st.columns([2, 3, 2, 2, 1])
                    cat_val = c1.selectbox(f"Categoría #{idx+1}", TIPOS_EGRESO, key=f"cat_{idx}")
                    desc_val = c2.text_input(f"Concepto #{idx+1}", key=f"desc_{idx}", placeholder="Ej. Reparación, Compra Extra")
                    monto_val = c3.number_input(f"Importe #{idx+1}", min_value=0.0, format="%.2f", key=f"monto_{idx}")
                    moneda_val = c4.selectbox(f"Moneda #{idx+1}", MONEDAS, key=f"moneda_{idx}")
                    inc_val = c5.checkbox("Activo", value=True, key=f"inc_{idx}")
                    gastos_input.append((cat_val, desc_val, monto_val, moneda_val, inc_val))

            st.markdown("---")
            btn_registrar_egresos = st.form_submit_button("Guardar")
            
            if btn_registrar_egresos:
                conn = sqlite3.connect("finanzas.db")
                c = conn.cursor()
                guardados = 0
                for cat, desc, monto, moneda, incluir in gastos_input:
                    if incluir and desc and monto > 0:
                        c.execute(
                            "INSERT INTO egresos (codigo, anio, mes, tipo, descripcion, monto, moneda) VALUES (?, ?, ?, ?, ?, ?, ?)",
                            (codigo_egr, anio_egr, mes_egr, cat, desc, monto, moneda)
                        )
                        guardados += 1
                conn.commit()
                conn.close()
                if guardados > 0:
                    st.session_state.cant_egresos_adicionales = 0
                    exportar_respaldo_csv()
                    st.success(f"Documento {codigo_egr} guardado exitosamente con {guardados} partidas.")
                    st.rerun()
                else:
                    st.warning("Ingrese al menos una partida con descripción e importe mayor a cero.")

    st.markdown("#### Documentos de Egresos Registrados")
    conn = sqlite3.connect("finanzas.db")
    df_egresos = pd.read_sql_query("SELECT codigo AS 'Código', anio AS Año, mes AS Mes, tipo AS Tipo, descripcion AS Descripción, monto AS Monto, moneda AS Moneda FROM egresos", conn)
    conn.close()

    if not df_egresos.empty:
        excel_bytes_egr = exportar_dataframe_a_excel(df_egresos, sheet_name="Egresos")
        st.download_button(
            label="📥 Descargar Egresos (Excel)",
            data=excel_bytes_egr,
            file_name=f"Egresos_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        codigos_unicos = df_egresos['Código'].unique()
        for cod in codigos_unicos:
            sub_df = df_egresos[df_egresos['Código'] == cod]
            anio_val = sub_df['Año'].iloc[0]
            mes_val = sub_df['Mes'].iloc[0]
            sum_pen = sub_df[sub_df['Moneda'] == 'PEN']['Monto'].sum()
            sum_usd = sub_df[sub_df['Moneda'] == 'USD']['Monto'].sum()

            with st.expander(f"Documento Lote: {cod} | Período: {mes_val} {anio_val} | Total PEN: S/ {sum_pen:,.2f} | Total USD: $ {sum_usd:,.2f}"):
                st.dataframe(sub_df[['Tipo', 'Descripción', 'Monto', 'Moneda']], use_container_width=True)
    else:
        st.info("No hay documentos de egresos registrados.")

# ==========================================
# PESTAÑA 3: DASHBOARD DE CONTROL Y SALDOS
# ==========================================
with tabs[2]:
    st.subheader("Análisis KPI y Balance de Cierre")
    
    col_f1, col_f2 = st.columns(2)
    anio_sel = col_f1.selectbox("Filtrar Ejercicio", ANIOS, index=0, key="dash_anio")
    mes_sel = col_f2.selectbox("Filtrar Período", MESES, index=0, key="dash_mes")

    conn = sqlite3.connect("finanzas.db")
    df_ing_m = pd.read_sql_query("SELECT 'INGRESO' AS Tipo, anio AS Año, mes AS Mes, tipo AS Categoria, descripcion AS Descripción, monto AS Monto, moneda AS Moneda FROM ingresos WHERE anio=? AND mes=?", conn, params=(anio_sel, mes_sel))
    df_egr_m = pd.read_sql_query("SELECT 'EGRESO' AS Tipo, anio AS Año, mes AS Mes, tipo AS Categoria, descripcion AS Descripción, monto AS Monto, moneda AS Moneda FROM egresos WHERE anio=? AND mes=?", conn, params=(anio_sel, mes_sel))
    conn.close()

    df_consolidado = pd.concat([df_ing_m, df_egr_m], ignore_index=True)

    if not df_consolidado.empty:
        excel_bytes_dash = exportar_dataframe_a_excel(df_consolidado, sheet_name=f"Balance_{mes_sel}_{anio_sel}")
        st.download_button(
            label="📥 Descargar Balance Mensual (Excel)",
            data=excel_bytes_dash,
            file_name=f"Balance_{mes_sel}_{anio_sel}_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    ing_pen = df_ing_m[df_ing_m['Moneda'] == 'PEN']['Monto'].sum()
    egr_pen = df_egr_m[df_egr_m['Moneda'] == 'PEN']['Monto'].sum()
    balance_pen = ing_pen - egr_pen

    st.markdown(f"##### Reporte de Balance de Período ({mes_sel} {anio_sel})")
    
    kpi1, kpi2, kpi3 = st.columns(3)
    
    with kpi1:
        render_kpi("Ingresos Totales (PEN)", f"S/ {ing_pen:,.2f}", "sap-kpi-positive")
        
    with kpi2:
        render_kpi("Egresos Totales (PEN)", f"S/ {egr_pen:,.2f}", "sap-kpi-negative")
        
    with kpi3:
        card_class = "sap-kpi-positive" if balance_pen >= 0 else "sap-kpi-negative"
        render_kpi("Flujo Neto / Superávit (PEN)", f"S/ {balance_