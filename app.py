import pandas as pd
import streamlit as st
import sqlite3
from datetime import datetime

# --- CONFIGURACIÓN DE PÁGINA STYLED FOR SAP FIORI ---
st.set_page_config(
    page_title="SAP Fiori - Control de Finanzas Personales",
    page_icon="💼",
    layout="wide"
)

# --- ESTILOS CSS PERSONALIZADOS (ESTILO SAP FIORI HORIZON / QUARTZ THEME) ---
SAP_CSS = """
<style>
    @import url('https://fonts.googleapis.com/css2?family=72:wght@400;600;700&family=Inter:wght@400;500;600;700&display=swap');

    html, body, [class*="css"] {
        font-family: '72', 'Inter', 'Segoe UI', -apple-system, sans-serif;
        background-color: #F5F7FA !important;
        color: #1D2D3E;
    }

    header[data-testid="stHeader"] {
        background-color: transparent !important;
    }
    .main .block-container {
        padding-top: 1rem !important;
        padding-bottom: 2rem !important;
        max-width: 96% !important;
    }

    /* BARRA DE ENCABEZADO ESTILO SAP SHELLBAR */
    .sap-shellbar {
        background: linear-gradient(90deg, #1C2D42 0%, #354A5F 100%);
        color: #FFFFFF;
        padding: 12px 24px;
        border-radius: 6px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        box-shadow: 0 4px 10px rgba(0, 0, 0, 0.15);
        margin-bottom: 20px;
    }
    .sap-shellbar .logo-title {
        display: flex;
        align-items: center;
        gap: 12px;
    }
    .sap-logo {
        background-color: #0070F2;
        color: #FFFFFF;
        font-weight: 800;
        font-size: 14px;
        padding: 4px 8px;
        border-radius: 3px;
        letter-spacing: 1px;
    }
    .sap-title {
        font-size: 18px;
        font-weight: 600;
        color: #FFFFFF;
        margin: 0;
    }
    .sap-env-badge {
        background-color: rgba(255, 255, 255, 0.15);
        border: 1px solid rgba(255, 255, 255, 0.3);
        color: #E2E8F0;
        font-size: 12px;
        padding: 3px 10px;
        border-radius: 12px;
    }

    /* TILES / CARDS ESTILO SAP FIORI LAUNCHPAD */
    .sap-tile {
        background: #FFFFFF;
        border: 1px solid #D9E1E8;
        border-radius: 8px;
        padding: 18px 20px;
        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.04);
        transition: transform 0.2s, box-shadow 0.2s;
        margin-bottom: 15px;
    }
    .sap-tile:hover {
        box-shadow: 0 4px 12px rgba(0, 112, 242, 0.12);
        border-color: #0070F2;
    }
    .sap-tile-header {
        font-size: 13px;
        font-weight: 600;
        color: #556B82;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        margin-bottom: 8px;
    }
    .sap-tile-value {
        font-size: 26px;
        font-weight: 700;
        color: #1C2D42;
        margin-bottom: 6px;
    }
    .sap-tile-footer {
        font-size: 12px;
        color: #6A7D93;
        display: flex;
        align-items: center;
        gap: 6px;
    }
    .sap-tile-bar-green { border-bottom: 4px solid #107E3E; }
    .sap-tile-bar-red { border-bottom: 4px solid #BB0000; }
    .sap-tile-bar-blue { border-bottom: 4px solid #0070F2; }

    /* BOTONES ESTILO SAP FIORI */
    .stButton > button, div[data-testid="stForm"] button {
        background-color: #0070F2 !important;
        color: #FFFFFF !important;
        border: none !important;
        border-radius: 4px !important;
        font-weight: 600 !important;
        font-size: 14px !important;
        padding: 8px 20px !important;
        box-shadow: 0 2px 4px rgba(0, 112, 242, 0.2) !important;
        transition: all 0.2s !important;
    }
    .stButton > button:hover, div[data-testid="stForm"] button:hover {
        background-color: #0056B3 !important;
        box-shadow: 0 4px 8px rgba(0, 112, 242, 0.3) !important;
    }

    /* PESTAÑAS ESTILO SAP FIORI ICON TAB BAR */
    div[data-baseweb="tab-list"] {
        background-color: #FFFFFF;
        padding: 4px 12px;
        border-radius: 6px;
        border: 1px solid #D9E1E8;
        gap: 10px;
    }
    button[data-baseweb="tab"] {
        color: #556B82 !important;
        font-weight: 600 !important;
        font-size: 14px !important;
        border-bottom: 3px solid transparent !important;
        padding: 10px 16px !important;
    }
    button[aria-selected="true"] {
        color: #0070F2 !important;
        border-bottom: 3px solid #0070F2 !important;
        background-color: #F0F6FF !important;
        border-radius: 4px 4px 0 0;
    }

    /* BADGES DE ESTADO ESTILO SAP */
    .sap-badge {
        display: inline-block;
        padding: 2px 8px;
        border-radius: 4px;
        font-size: 12px;
        font-weight: 600;
    }
    .sap-badge-success { background-color: #E5F6ED; color: #107E3E; border: 1px solid #B7E4C7; }
    .sap-badge-warning { background-color: #FFF4E5; color: #E9730C; border: 1px solid #FFE0B2; }
    .sap-badge-danger { background-color: #FDE8E8; color: #BB0000; border: 1px solid #F8B4B4; }
    .sap-badge-info { background-color: #EBF3FE; color: #0070F2; border: 1px solid #B8D5FA; }

    /* CONTENEDORES DE FORMULARIO */
    div[data-testid="stForm"] {
        background-color: #FFFFFF;
        border: 1px solid #D9E1E8;
        border-radius: 8px;
        padding: 20px;
        box-shadow: 0 2px 6px rgba(0,0,0,0.02);
    }
    
    /* TABLAS DE DATOS */
    div[data-testid="stDataFrame"] {
        border: 1px solid #D9E1E8;
        border-radius: 6px;
        overflow: hidden;
    }
</style>
"""
st.markdown(SAP_CSS, unsafe_allow_html=True)

# --- BASE DE DATOS LOCAL ---
def init_db():
    conn = sqlite3.connect("finanzas_sap.db")
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS ingresos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            anio INTEGER,
            mes TEXT,
            tipo TEXT,
            descripcion TEXT,
            monto REAL,
            moneda TEXT
        )
    """)
    c.execute("""
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
    """)
    conn.commit()
    conn.close()

init_db()

# --- CONSTANTES ---
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

def obtener_siguiente_codigo():
    conn = sqlite3.connect("finanzas_sap.db")
    df = pd.read_sql_query("SELECT codigo FROM egresos", conn)
    conn.close()
    if df.empty or df['codigo'].dropna().empty:
        return "E001"
    else:
        codigos = df['codigo'].str.replace('E', '').astype(int)
        siguiente = codigos.max() + 1
        return f"E{siguiente:03d}"

# --- BARRA SUPERIOR SHELLBAR SAP FIORI ---
st.markdown("""
<div class="sap-shellbar">
    <div class="logo-title">
        <span class="sap-logo">SAP</span>
        <span class="sap-title">Fiori Launchpad — Personal Financial Management & Budget Control</span>
    </div>
    <div class="sap-env-badge">S/4HANA Edition v2026.1</div>
</div>
""", unsafe_allow_html=True)

# --- PESTAÑAS TIPO SAP ICON TAB BAR ---
tabs = st.tabs(["💰 Control de Ingresos", "💸 Control de Egresos", "📊 Dashboard & Analítica Fiori"])

# ==========================================
# PESTAÑA 1: INGRESOS
# ==========================================
with tabs[0]:
    st.markdown("### 📥 Entrada de Documentos de Ingreso")
    st.markdown("Registre nuevos flujos de caja entrantes para consolidación mensual.")
    
    with st.form(key="form_ingresos_sap"):
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            anio_ing = st.selectbox("Año Fiscal", ANIOS, key="ing_anio")
            tipo_ing = st.selectbox("Clase de Ingreso", TIPOS_INGRESO, key="ing_tipo")
        with col2:
            mes_ing = st.selectbox("Período (Mes)", MESES, key="ing_mes")
            moneda_ing = st.selectbox("Moneda de Registro", MONEDAS, key="ing_moneda")
        with col3:
            desc_ing = st.text_input("Texto de Cabecera / Descripción", key="ing_desc")
        with col4:
            monto_ing = st.number_input("Monto Imputable", min_value=0.0, format="%.2f", key="ing_monto")
        
        btn_guardar = st.form_submit_button("💾 Contabilizar Ingreso")
        if btn_guardar:
            if desc_ing and monto_ing > 0:
                conn = sqlite3.connect("finanzas_sap.db")
                c = conn.cursor()
                c.execute(
                    "INSERT INTO ingresos (anio, mes, tipo, descripcion, monto, moneda) VALUES (?, ?, ?, ?, ?, ?)",
                    (anio_ing, mes_ing, tipo_ing, desc_ing, monto_ing, moneda_ing)
                )
                conn.commit()
                conn.close()
                st.markdown('<span class="sap-badge sap-badge-success">✔ Registro Contabilizado Correctamente</span>', unsafe_allow_html=True)
                st.rerun()
            else:
                st.markdown('<span class="sap-badge sap-badge-warning">⚠ Complete la descripción y asegúrese de registrar un monto mayor a cero.</span>', unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("#### 📑 Partidas Individuales de Ingresos Registradas")
    conn = sqlite3.connect("finanzas_sap.db")
    df_ingresos = pd.read_sql_query("SELECT id AS 'Doc ID', anio AS 'Año Fiscal', mes AS Período, tipo AS Clase, descripcion AS Descripción, monto AS Monto, moneda AS Moneda FROM ingresos", conn)
    conn.close()

    if not df_ingresos.empty:
        st.dataframe(df_ingresos, use_container_width=True)
        
        col_t1, col_t2 = st.columns(2)
        sum_pen = df_ingresos[df_ingresos['Moneda'] == 'PEN']['Monto'].sum()
        sum_usd = df_ingresos[df_ingresos['Moneda'] == 'USD']['Monto'].sum()
        
        with col_t1:
            st.markdown(f"""
            <div class="sap-tile sap-tile-bar-green">
                <div class="sap-tile-header">Acumulado Ingresos PEN</div>
                <div class="sap-tile-value">S/ {sum_pen:,.2f}</div>
                <div class="sap-tile-footer"><span class="sap-badge sap-badge-success">Moneda Local</span></div>
            </div>
            """, unsafe_allow_html=True)
        with col_t2:
            st.markdown(f"""
            <div class="sap-tile sap-tile-bar-blue">
                <div class="sap-tile-header">Acumulado Ingresos USD</div>
                <div class="sap-tile-value">$ {sum_usd:,.2f}</div>
                <div class="sap-tile-footer"><span class="sap-badge sap-badge-info">Moneda Extranjera</span></div>
            </div>
            """, unsafe_allow_html=True)
    else:
        st.info("No se encontraron partidas individuales de ingreso.")

# ==========================================
# PESTAÑA 2: EGRESOS
# ==========================================
with tabs[1]:
    st.markdown("### 📤 Registro de Documentos de Egreso (Orden de Gastos)")
    
    siguiente_codigo = obtener_siguiente_codigo()
    
    col_e1, col_e2, col_e3 = st.columns(3)
    anio_egr = col_e1.selectbox("Ejercicio (Año)", ANIOS, key="egr_anio")
    mes_egr = col_e2.selectbox("Período Mensual", MESES, key="egr_mes")
    codigo_egr = col_e3.text_input("Nº de Registro Correlativo", value=siguiente_codigo, disabled=True)

    st.markdown("##### 📋 Matriz de Partidas de Egreso (Fijas y Adicionales)")
    
    with st.form(key="form_egresos_sap"):
        gastos_input = []
        for i, (cat, desc) in enumerate(EGRESOS_FIJOS_BASE):
            c1, c2, c3, c4, c5 = st.columns([2, 3, 2, 2, 1])
            c1.text_input(f"Categoría #{i+1}", value=cat, key=f"cat_{i}", disabled=True)
            c2.text_input(f"Concepto #{i+1}", value=desc, key=f"desc_{i}")
            monto = c3.number_input(f"Importe", min_value=0.0, format="%.2f", key=f"monto_{i}")
            moneda = c4.selectbox(f"Moneda", MONEDAS, key=f"moneda_{i}")
            incluir = c5.checkbox("Activo", value=True, key=f"inc_{i}")
            gastos_input.append((cat, desc, monto, moneda, incluir))

        btn_registrar_egresos = st.form_submit_button("💾 Liquidar y Contabilizar Lote")
        
        if btn_registrar_egresos:
            conn = sqlite3.connect("finanzas_sap.db")
            c = conn.cursor()
            guardados = 0
            for cat, desc, monto, moneda, incluir in gastos_input:
                if incluir and monto > 0:
                    c.execute(
                        "INSERT INTO egresos (codigo, anio, mes, tipo, descripcion, monto, moneda) VALUES (?, ?, ?, ?, ?, ?, ?)",
                        (codigo_egr, anio_egr, mes_egr, cat, desc, monto, moneda)
                    )
                    guardados += 1
            conn.commit()
            conn.close()
            if guardados > 0:
                st.success(f"Lote {codigo_egr} contabilizado con {guardados} posiciones.")
                st.rerun()
            else:
                st.warning("No se ingresaron importes superiores a cero.")

    st.markdown("---")
    st.markdown("#### 📁 Documentos de Egreso Procesados")
    conn = sqlite3.connect("finanzas_sap.db")
    df_egresos = pd.read_sql_query("SELECT codigo, anio, mes, tipo, descripcion, monto, moneda FROM egresos", conn)
    conn.close()

    if not df_egresos.empty:
        codigos_unicos = df_egresos['codigo'].unique()
        for cod in codigos_unicos:
            sub_df = df_egresos[df_egresos['codigo'] == cod]
            anio_val = sub_df['anio'].iloc[0]
            mes_val = sub_df['mes'].iloc[0]
            sum_pen = sub_df[sub_df['moneda'] == 'PEN']['monto'].sum()
            sum_usd = sub_df[sub_df['moneda'] == 'USD']['monto'].sum()

            with st.expander(f"📄 Documento SAP: {cod} | Ejercicio: {mes_val} {anio_val} | Total PEN: S/ {sum_pen:,.2f} | Total USD: $ {sum_usd:,.2f}"):
                st.dataframe(sub_df[['tipo', 'descripcion', 'monto', 'moneda']], use_container_width=True)
    else:
        st.info("No existen egresos contabilizados en el sistema.")

# ==========================================
# PESTAÑA 3: DASHBOARD FIORI (ANÁLISIS DE BALANCE)
# ==========================================
with tabs[2]:
    st.markdown("### 📊 SAP Executive Financial Dashboard (Fiori Analytics)")
    
    col_d1, col_d2 = st.columns(2)
    anio_sel = col_d1.selectbox("Seleccionar Ejercicio", ANIOS, index=0, key="dash_anio")
    mes_sel = col_d2.selectbox("Seleccionar Período", MESES, index=0, key="dash_mes")

    conn = sqlite3.connect("finanzas_sap.db")
    df_ing_m = pd.read_sql_query("SELECT monto, moneda FROM ingresos WHERE anio=? AND mes=?", conn, params=(anio_sel, mes_sel))
    df_egr_m = pd.read_sql_query("SELECT monto, moneda FROM egresos WHERE anio=? AND mes=?", conn, params=(anio_sel, mes_sel))
    conn.close()

    ing_pen = df_ing_m[df_ing_m['moneda'] == 'PEN']['monto'].sum()
    egr_pen = df_egr_m[df_egr_m['moneda'] == 'PEN']['monto'].sum()
    balance_pen = ing_pen - egr_pen

    st.markdown(f"##### KPI Cockpit — {mes_sel} {anio_sel}")
    
    c_kpi1, c_kpi2, c_kpi3 = st.columns(3)
    
    with c_kpi1:
        st.markdown(f"""
        <div class="sap-tile sap-tile-bar-green">
            <div class="sap-tile-header">Total Entradas (Ingresos)</div>
            <div class="sap-tile-value">S/ {ing_pen:,.2f}</div>
            <div class="sap-tile-footer"><span class="sap-badge sap-badge-success">✔ Flujo Positivo</span></div>
        </div>
        """, unsafe_allow_html=True)
        
    with c_kpi2:
        st.markdown(f"""
        <div class="sap-tile sap-tile-bar-red">
            <div class="sap-tile-header">Total Salidas (Egresos)</div>
            <div class="sap-tile-value">S/ {egr_pen:,.2f}</div>
            <div class="sap-tile-footer"><span class="sap-badge sap-badge-danger">🔻 Consumo de Recurso</span></div>
        </div>
        """, unsafe_allow_html=True)

    with c_kpi3:
        status_badge = '<span class="sap-badge sap-badge-success">✔ Superávit</span>' if balance_pen >= 0 else '<span class="sap-badge sap-badge-danger">⚠ Déficit Fiscal</span>'
        bar_color = "sap-tile-bar-blue" if balance_pen >= 0 else "sap-tile-bar-red"
        st.markdown(f"""
        <div class="sap-tile {bar_color}">
            <div class="sap-tile-header">Resultado Neto (Balance)</div>
            <div class="sap-tile-value">S/ {balance_pen:,.2f}</div>
            <div class="sap-tile-footer">{status_badge}</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("##### 🎯 Control Presupuestario y Regla de Gestión 50/30/20")
    if ing_pen > 0:
        tasa_ahorro = (balance_pen / ing_pen) * 100
        st.write(f"**Tasa de Cobertura / Ahorro:** {tasa_ahorro:.1f}%")
        st.progress(max(0.0, min(1.0, balance_pen / ing_pen)))
        
        c_r1, c_r2, c_r3 = st.columns(3)
        c_r1.info(f"**50% Necesidades Fijas:** S/ {ing_pen * 0.5:,.2f}")
        c_r2.warning(f"**30% Gastos Variables:** S/ {ing_pen * 0.3:,.2f}")
        c_r3.success(f"**20% Ahorro Mínimo Sugerido:** S/ {ing_pen * 0.2:,.2f}")
    else:
        st.info("Sin registros de ingresos en el período seleccionado para calcular métricas de ahorro.")