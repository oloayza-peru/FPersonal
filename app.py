
def generar_excel_ingresos_egresos() -> bytes:
    conn = sqlite3.connect("finanzas.db")
    df_ing = pd.read_sql_query("SELECT id AS 'Doc #', anio AS Año, mes AS Mes, tipo AS Tipo, descripcion AS Descripción, monto AS Monto, moneda AS Moneda FROM ingresos", conn)
    df_egr = pd.read_sql_query("SELECT id AS 'Doc #', codigo AS Código, anio AS Año, mes AS Mes, tipo AS Tipo, descripcion AS Descripción, monto AS Monto, moneda AS Moneda FROM egresos", conn)
    conn.close()
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_ing.to_excel(writer, sheet_name='Ingresos', index=False)
        df_egr.to_excel(writer, sheet_name='Egresos', index=False)
    
    return output.getvalue()


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

# --- ENLACES A GOOGLE DRIVE Y GOOGLE SHEETS ---
DRIVE_FOLDER_URL = "https://drive.google.com/drive/folders/1Mqp3Lf-U2PC5BXPCr949R4y3VbSBVk_B?usp=drive_link"
DRIVE_FOLDER_NAME = "Carpeta Personal"
SHEETS_REGISTROS_URL = "https://docs.google.com/spreadsheets/d/10nh75p-uzJAsXjVJxnMcicEul7pvw4TiZB9tgHExlvA/edit?usp=sharing"

# --- CREDENCIALES DE GITHUB NUEVAS ---

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

# --- FUNCION PARA GUARDAR EN GITHUB ---
def sincronizar_con_github(nombre_archivo="respaldo_finanzas.csv"):
    pass

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
        codigos = df['codigo'].str.replace('E', '').astype(int)
        siguiente = codigos.max() + 1
        return f"E{siguiente:03d}"

def exportar_dataframe_a_csv(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False, sep=';', encoding='utf-8-sig').encode('utf-8-sig')

def exportar_respaldo_csv(sincronizar=False):
    conn = sqlite3.connect("finanzas.db")
    df_ing = pd.read_sql_query("SELECT 'INGRESO' AS Tipo_Registro, anio, mes, tipo, descripcion, monto, moneda FROM ingresos", conn)
    df_egr = pd.read_sql_query("SELECT 'EGRESO' AS Tipo_Registro, anio, mes, tipo, descripcion, monto, moneda FROM egresos", conn)
    conn.close()
    df_total = pd.concat([df_ing, df_egr], ignore_index=True)
    df_total.to_csv("respaldo_finanzas.csv", index=False)

# --- CABECERA SAP FIORI ---
shell_bar_html = """
<div class="sap-shell-bar">
    <div style="display: flex; align-items: center; gap: 10px; flex-wrap: wrap;">
        <span class="sap-shell-title">SAP S/4HANA — Financial Control & Treasury Management</span>
    </div>
    <div class="sap-shell-badge">SAP Horizon</div>
</div>
"""
st.html(shell_bar_html)

# Botón de descarga general de Excel consolidado en la parte superior o lateral
col_down_1, col_down_2 = st.columns([3, 1])
with col_down_2:
    excel_data = generar_excel_ingresos_egresos()
    st.download_button(
        label="📥 Descargar Todo (.xlsx)",
        data=excel_data,
        file_name=f"Control_Financiero_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
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
                    conn = sqlite3.connect("finanzas.db")
                    c = conn.cursor()
                    c.execute(
                        "INSERT INTO ingresos (anio, mes, tipo, descripcion, monto, moneda) VALUES (?, ?, ?, ?, ?, ?)",
                        (anio_ing, mes_ing, tipo_ing, desc_ing, monto_ing, moneda_ing)
                    )
                    conn.commit()
                    conn.close()
                    exportar_respaldo_csv(sincronizar=True)
                    st.success("Ingreso guardado exitosamente.")
                    st.rerun()
                else:
                    st.warning("Verifique que la descripción y el monto sean válidos.")

    st.markdown("#### Registros Contabilizados")
    conn = sqlite3.connect("finanzas.db")
    df_ingresos = pd.read_sql_query("SELECT id AS 'Doc #', anio AS Año, mes AS Mes, tipo AS Tipo, descripcion AS Descripción, monto AS Monto, moneda AS Moneda FROM ingresos", conn)
    conn.close()

    if not df_ingresos.empty:
        col_btn1, col_btn2 = st.columns([1, 2])
        with col_btn1:
            st.download_button(
                label="📥 Descargar Ingresos (.csv)",
                data=exportar_dataframe_a_csv(df_ingresos),
                file_name=f"Ingresos_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv",
                use_container_width=True
            )


        st.dataframe(df_ingresos, use_container_width=True)

        st.markdown("##### ✏️ Modificación / Eliminación de Registro de Ingreso")
        col_ed1, col_ed2, col_ed3 = st.columns([2, 1, 1])
        id_ing_sel = col_ed1.selectbox("Seleccionar Documento (Doc #)", df_ingresos['Doc #'].tolist(), key="select_ing_edit")
        
        if col_ed2.button("✏️ Cargar para Editar", key="btn_load_ing"):
            st.session_state.edit_ingreso_id = id_ing_sel

        if col_ed3.button("🗑️ Eliminar Registro", key="btn_del_ing"):
            conn = sqlite3.connect("finanzas.db")
            c = conn.cursor()
            c.execute("DELETE FROM ingresos WHERE id=?", (id_ing_sel,))
            conn.commit()
            conn.close()
            if st.session_state.edit_ingreso_id == id_ing_sel:
                st.session_state.edit_ingreso_id = None
            exportar_respaldo_csv(sincronizar=True)
            st.success(f"Ingreso Doc #{id_ing_sel} eliminado exitosamente.")
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
                        conn = sqlite3.connect("finanzas.db")
                        c = conn.cursor()
                        c.execute("""
                            UPDATE ingresos 
                            SET anio=?, mes=?, tipo=?, descripcion=?, monto=?, moneda=?
                            WHERE id=?
                        """, (anio_e, mes_e, tipo_e, desc_e, monto_e, moneda_e, st.session_state.edit_ingreso_id))
                        conn.commit()
                        conn.close()
                        st.session_state.edit_ingreso_id = None
                        exportar_respaldo_csv(sincronizar=True)
                        st.success("Registro actualizado y guardado localmente exitosamente.")
                        st.rerun()

        col_sum1, col_sum2 = st.columns(2)
        monto_pen = df_ingresos[df_ingresos['Moneda'] == 'PEN']['Monto'].sum()
        monto_usd = df_ingresos[df_ingresos['Moneda'] == 'USD']['Monto'].sum()
        
        col_sum1.metric("Acumulado Ingresos (PEN)", f"S/ {monto_pen:,.2f}")
        col_sum2.metric("Acumulado Ingresos (USD)", f"$ {monto_usd:,.2f}")
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
        anio_egr = col_h1.selectbox("Ejercicio (Año)", ANIOS, index=ANIO_ACTUAL_INDEX, key="egr_anio")
        mes_egr = col_h2.selectbox("Período Mensual", MESES, index=MES_ACTUAL_INDEX, key="egr_mes")
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
            btn_registrar_egresos = st.form_submit_button("Guardar Lote de Egresos")
            
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
                    exportar_respaldo_csv(sincronizar=True)
                    st.success(f"Documento {codigo_egr} guardado y guardado localmente con {guardados} partidas.")
                    st.rerun()
                else:
                    st.warning("Ingrese al menos una partida con descripción e importe mayor a cero.")

    st.markdown("#### Documentos de Egresos Registrados")
    conn = sqlite3.connect("finanzas.db")
    df_egresos = pd.read_sql_query("SELECT id, codigo AS 'Código', anio AS Año, mes AS Mes, tipo AS Tipo, descripcion AS Descripción, monto AS Monto, moneda AS Moneda FROM egresos", conn)
    conn.close()

    if not df_egresos.empty:
        col_btn1, col_btn2 = st.columns([1, 2])
        with col_btn1:
            st.download_button(
                label="📥 Descargar Egresos (.csv)",
                data=exportar_dataframe_a_csv(df_egresos[['Código', 'Año', 'Mes', 'Tipo', 'Descripción', 'Monto', 'Moneda']]),
                file_name=f"Egresos_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv",
                use_container_width=True
            )


        st.markdown("##### ✏️ Modificación / Eliminación de Lote de Egresos")
        codigos_unicos = df_egresos['Código'].unique().tolist()
        col_ed_eg1, col_ed_eg2, col_ed_eg3 = st.columns([2, 1, 1])
        cod_egr_sel = col_ed_eg1.selectbox("Seleccionar Documento Lote (Código)", codigos_unicos, key="select_egr_edit")
        
        if col_ed_eg2.button("✏️ Cargar para Editar", key="btn_load_egr"):
            st.session_state.edit_egreso_codigo = cod_egr_sel

        if col_ed_eg3.button("🗑️ Eliminar Lote Completo", key="btn_del_egr"):
            conn = sqlite3.connect("finanzas.db")
            c = conn.cursor()
            c.execute("DELETE FROM egresos WHERE codigo=?", (cod_egr_sel,))
            conn.commit()
            conn.close()
            if st.session_state.edit_egreso_codigo == cod_egr_sel:
                st.session_state.edit_egreso_codigo = None
            exportar_respaldo_csv(sincronizar=True)
            st.success(f"Documento Lote {cod_egr_sel} eliminado exitosamente.")
            st.rerun()

        if st.session_state.edit_egreso_codigo:
            sub_egr = df_egresos[df_egresos['Código'] == st.session_state.edit_egreso_codigo]
            with st.expander(f"📝 Editando Partidas del Documento Lote: {st.session_state.edit_egreso_codigo}", expanded=True):
                with st.form(key="form_edit_egreso_lote"):
                    c_head1, c_head2 = st.columns(2)
                    anio_lote_val = int(sub_egr['Año'].iloc[0])
                    mes_lote_val = sub_egr['Mes'].iloc[0]
                    
                    anio_eg_edit = c_head1.selectbox("Ejercicio (Año)", ANIOS, index=ANIOS.index(anio_lote_val) if anio_lote_val in ANIOS else 0, key="edit_lote_anio")
                    mes_eg_edit = c_head2.selectbox("Período Mensual", MESES, index=MESES.index(mes_lote_val) if mes_lote_val in MESES else 0, key="edit_lote_mes")
                    st.markdown("---")
                    
                    items_editados = []
                    for idx_p, row_p in sub_egr.iterrows():
                        p_id = row_p['id']
                        c1, c2, c3, c4 = st.columns([2, 3, 2, 2])
                        tipo_p = c1.selectbox(f"Categoría #{p_id}", TIPOS_EGRESO, index=TIPOS_EGRESO.index(row_p['Tipo']) if row_p['Tipo'] in TIPOS_EGRESO else 0, key=f"edit_p_tipo_{p_id}")
                        desc_p = c2.text_input(f"Descripción #{p_id}", value=row_p['Descripción'], key=f"edit_p_desc_{p_id}")
                        monto_p = c3.number_input(f"Monto #{p_id}", value=float(row_p['Monto']), min_value=0.0, format="%.2f", key=f"edit_p_monto_{p_id}")
                        moneda_p = c4.selectbox(f"Moneda #{p_id}", MONEDAS, index=MONEDAS.index(row_p['Moneda']) if row_p['Moneda'] in MONEDAS else 0, key=f"edit_p_moneda_{p_id}")
                        items_editados.append((p_id, tipo_p, desc_p, monto_p, moneda_p))

                    btn_update_egr = st.form_submit_button("💾 Actualizar Lote Completo")
                    if btn_update_egr:
                        conn = sqlite3.connect("finanzas.db")
                        c = conn.cursor()
                        for p_id, tipo_p, desc_p, monto_p, moneda_p in items_editados:
                            c.execute("""
                                UPDATE egresos
                                SET anio=?, mes=?, tipo=?, descripcion=?, monto=?, moneda=?
                                WHERE id=?
                            """, (anio_eg_edit, mes_eg_edit, tipo_p, desc_p, monto_p, moneda_p, p_id))
                        conn.commit()
                        conn.close()
                        st.session_state.edit_egreso_codigo = None
                        exportar_respaldo_csv(sincronizar=True)
                        st.success(f"Documento Lote {cod_egr_sel} actualizado y guardado localmente.")
                        st.rerun()

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
    anio_sel = col_f1.selectbox("Filtrar Ejercicio", ANIOS, index=ANIO_ACTUAL_INDEX, key="dash_anio")
    mes_sel = col_f2.selectbox("Filtrar Período", MESES, index=MES_ACTUAL_INDEX, key="dash_mes")

    conn = sqlite3.connect("finanzas.db")
    df_ing_m = pd.read_sql_query("SELECT 'INGRESO' AS Tipo, anio AS Año, mes AS Mes, tipo AS Categoria, descripcion AS Descripción, monto AS Monto, moneda AS Moneda FROM ingresos WHERE anio=? AND mes=?", conn, params=(anio_sel, mes_sel))
    df_egr_m = pd.read_sql_query("SELECT 'EGRESO' AS Tipo, anio AS Año, mes AS Mes, tipo AS Categoria, descripcion AS Descripción, monto AS Monto, moneda AS Moneda FROM egresos WHERE anio=? AND mes=?", conn, params=(anio_sel, mes_sel))
    conn.close()

    df_consolidado = pd.concat([df_ing_m, df_egr_m], ignore_index=True)

    if not df_consolidado.empty:
        col_btn1, col_btn2 = st.columns([1, 2])
        with col_btn1:
            st.download_button(
                label="📥 Descargar Balance Mensual (.csv)",
                data=exportar_dataframe_a_csv(df_consolidado),
                file_name=f"Balance_{mes_sel}_{anio_sel}_{datetime.now().strftime('%Y%m%d')}.csv",
                mime="text/csv",
                use_container_width=True
            )


    ing_pen = df_ing_m[df_ing_m['Moneda'] == 'PEN']['Monto'].sum()
    egr_pen = df_egr_m[df_egr_m['Moneda'] == 'PEN']['Monto'].sum()
    balance_pen = ing_pen - egr_pen

    st.markdown(f"##### Reporte de Balance de Período ({mes_sel} {anio_sel})")
    
    kpi1, kpi2, kpi3 = st.columns(3)
    
    kpi1.metric("Ingresos Totales (PEN)", f"S/ {ing_pen:,.2f}")
    kpi2.metric("Egresos Totales (PEN)", f"S/ {egr_pen:,.2f}")
    kpi3.metric("Flujo Neto / Superávit (PEN)", f"S/ {balance_pen:,.2f}")

    st.markdown("---")
    if ing_pen > 0:
        tasa_ahorro = (balance_pen / ing_pen) * 100
        st.progress(max(0.0, min(1.0, balance_pen / ing_pen)))
        st.caption(f"Tasa de Ahorro Neto alcanzada: **{tasa_ahorro:.1f}%** (Umbral SAP recomendado: ≥ 20.0%)")
