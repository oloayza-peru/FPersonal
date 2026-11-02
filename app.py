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
                df_egr = pd.read_excel(EXCEL_FILE, sheet_name="Egresos")
                nuevo_id_egr = int(df_egr['Doc #'].max() + 1) if not df_egr.empty and 'Doc #' in df_egr.columns and not df_egr['Doc #'].dropna().empty else 1
                
                guardados = 0
                nuevas_filas = []
                for cat, desc, monto, moneda, incluir in gastos_input:
                    if incluir and desc and monto > 0:
                        nuevas_filas.append({
                            'Doc #': nuevo_id_egr,
                            'Código': codigo_egr,
                            'Año': anio_egr,
                            'Mes': mes_egr,
                            'Tipo': cat,
                            'Descripción': desc,
                            'Monto': monto,
                            'Moneda': moneda
                        })
                        nuevo_id_egr += 1
                        guardados += 1
                
                if guardados > 0:
                    df_nuevos = pd.DataFrame(nuevas_filas)
                    df_egr = pd.concat([df_egr, df_nuevos], ignore_index=True)
                    
                    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                        pd.read_excel(EXCEL_FILE, sheet_name='Ingresos').to_excel(writer, sheet_name='Ingresos', index=False)
                        df_egr.to_excel(writer, sheet_name='Egresos', index=False)
                    
                    st.session_state.cant_egresos_adicionales = 0
                    st.success(f"Documento Lote {codigo_egr} guardado y actualizado en el Excel con {guardados} partidas.")
                    st.rerun()
                else:
                    st.warning("Ingrese al menos una partida con descripción e importe mayor a cero.")

    st.markdown("#### Documentos de Egresos Registrados")
    try:
        df_egresos = pd.read_excel(EXCEL_FILE, sheet_name="Egresos")
    except Exception:
        df_egresos = pd.DataFrame(columns=['Doc #', 'Código', 'Año', 'Mes', 'Tipo', 'Descripción', 'Monto', 'Moneda'])

    if not df_egresos.empty and 'Código' in df_egresos.columns:
        st.markdown("##### ✏️ Modificación / Eliminación de Lote de Egresos")
        codigos_unicos = df_egresos['Código'].dropna().unique().tolist()
        col_ed_eg1, col_ed_eg2, col_ed_eg3 = st.columns([2, 1, 1])
        cod_egr_sel = col_ed_eg1.selectbox("Seleccionar Documento Lote (Código)", codigos_unicos, key="select_egr_edit")
        
        if col_ed_eg2.button("✏️ Cargar para Editar", key="btn_load_egr"):
            st.session_state.edit_egreso_codigo = cod_egr_sel

        if col_ed_eg3.button("🗑️ Eliminar Lote Completo", key="btn_del_egr"):
            df_egresos = df_egresos[df_egresos['Código'] != cod_egr_sel]
            with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                pd.read_excel(EXCEL_FILE, sheet_name='Ingresos').to_excel(writer, sheet_name='Ingresos', index=False)
                df_egresos.to_excel(writer, sheet_name='Egresos', index=False)
            if st.session_state.edit_egreso_codigo == cod_egr_sel:
                st.session_state.edit_egreso_codigo = None
            st.success(f"Documento Lote {cod_egr_sel} eliminado y actualizado en el Excel.")
            st.rerun()

        if st.session_state.edit_egreso_codigo:
            sub_egr = df_egresos[df_egresos['Código'] == st.session_state.edit_egreso_codigo]
            with st.expander(f"📝 Editando Partidas del Documento Lote: {st.session_state.edit_egreso_codigo}", expanded=True):
                with st.form(key="form_edit_egreso_lote"):
                    c_head1, c_head2 = st.columns(2)
                    anio_lote_val = int(sub_egr['Año'].iloc[0]) if not sub_egr.empty else ANIO_ACTUAL
                    mes_lote_val = sub_egr['Mes'].iloc[0] if not sub_egr.empty else MESES[0]
                    
                    anio_eg_edit = c_head1.selectbox("Ejercicio (Año)", ANIOS, index=ANIOS.index(anio_lote_val) if anio_lote_val in ANIOS else 0, key="edit_lote_anio")
                    mes_eg_edit = c_head2.selectbox("Período Mensual", MESES, index=MESES.index(mes_lote_val) if mes_lote_val in MESES else 0, key="edit_lote_mes")
                    st.markdown("---")
                    
                    items_editados = []
                    for idx_p, row_p in sub_egr.iterrows():
                        p_id = row_p['Doc #']
                        c1, c2, c3, c4 = st.columns([2, 3, 2, 2])
                        tipo_p = c1.selectbox(f"Categoría #{idx_p}", TIPOS_EGRESO, index=TIPOS_EGRESO.index(row_p['Tipo']) if row_p['Tipo'] in TIPOS_EGRESO else 0, key=f"edit_p_tipo_{idx_p}")
                        desc_p = c2.text_input(f"Descripción #{idx_p}", value=row_p['Descripción'], key=f"edit_p_desc_{idx_p}")
                        monto_p = c3.number_input(f"Monto #{idx_p}", value=float(row_p['Monto']), min_value=0.0, format="%.2f", key=f"edit_p_monto_{idx_p}")
                        moneda_p = c4.selectbox(f"Moneda #{idx_p}", MONEDAS, index=MONEDAS.index(row_p['Moneda']) if row_p['Moneda'] in MONEDAS else 0, key=f"edit_p_moneda_{idx_p}")
                        items_editados.append((idx_p, tipo_p, desc_p, monto_p, moneda_p))

                    btn_update_egr = st.form_submit_button("💾 Actualizar Lote Completo")
                    if btn_update_egr:
                        for idx_p, tipo_p, desc_p, monto_p, moneda_p in items_editados:
                            df_egresos.loc[df_egresos.index == idx_p, ['Año', 'Mes', 'Tipo', 'Descripción', 'Monto', 'Moneda']] = [anio_eg_edit, mes_eg_edit, tipo_p, desc_p, monto_p, moneda_p]
                        
                        with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                            pd.read_excel(EXCEL_FILE, sheet_name='Ingresos').to_excel(writer, sheet_name='Ingresos', index=False)
                            df_egresos.to_excel(writer, sheet_name='Egresos', index=False)
                        st.session_state.edit_egreso_codigo = None
                        st.success(f"Documento Lote {cod_egr_sel} actualizado y sincronizado en el Excel.")
                        st.rerun()

        st.markdown("##### 🔍 Consulta de Partidas por Lista Desplegable de Lotes")
        opciones_lotes = {}
        for cod in codigos_unicos:
            sub_df = df_egresos[df_egresos['Código'] == cod]
            anio_val = sub_df['Año'].iloc[0] if not sub_df.empty else ""
            mes_val = sub_df['Mes'].iloc[0] if not sub_df.empty else ""
            sum_pen = sub_df[sub_df['Moneda'] == 'PEN']['Monto'].sum() if 'Moneda' in sub_df.columns else 0
            sum_usd = sub_df[sub_df['Moneda'] == 'USD']['Monto'].sum() if 'Moneda' in sub_df.columns else 0
            label = f"Documento Lote: {cod} | Período: {mes_val} {anio_val} | Total PEN: S/ {sum_pen:,.2f} | Total USD: $ {sum_usd:,.2f}"
            opciones_lotes[label] = cod

        selected_label = st.selectbox("Seleccionar Documento Lote para visualizar partidas", list(opciones_lotes.keys()), key="select_ver_lote")
        
        if selected_label:
            cod_seleccionado = opciones_lotes[selected_label]
            sub_df_sel = df_egresos[df_egresos['Código'] == cod_seleccionado]
            
            sum_p = sub_df_sel[sub_df_sel['Moneda'] == 'PEN']['Monto'].sum() if 'Moneda' in sub_df_sel.columns else 0
            sum_u = sub_df_sel[sub_df_sel['Moneda'] == 'USD']['Monto'].sum() if 'Moneda' in sub_df_sel.columns else 0
            cant_partidas = len(sub_df_sel)
            
            st.info(f"📋 **Resumen General - Lote {cod_seleccionado}**: {cant_partidas} partidas registradas | Total PEN: S/ {sum_p:,.2f} | Total USD: $ {sum_u:,.2f}")
            
            with st.expander(f"🔍 Ver vista de tabla completa del lote {cod_seleccionado} (Detalle de partidas)"):
                st.dataframe(sub_df_sel[['Tipo', 'Descripción', 'Monto', 'Moneda']], use_container_width=True)
    else:
        st.info("No hay documentos de egresos registrados.")

# ==========================================
# PESTAÑA 3: DASHBOARD ANALÍTICO Y ASESORÍA FINANCIERA
# ==========================================
with tabs[2]:
    st.subheader("Dashboard Analítico Avanzado y Plan de Salud Financiera")
    
    st.markdown("##### ⚙️ Filtros de Comparativa Multi-Período")
    col_f1, col_f2 = st.columns(2)
    anios_sel = col_f1.multiselect("Seleccionar Ejercicio(s)", ANIOS, default=[ANIO_ACTUAL], key="dash_anios")
    meses_sel = col_f2.multiselect("Seleccionar Mes(es)", MESES, default=MESES[:NOW.month], key="dash_meses")

    if anios_sel and meses_sel:
        try:
            df_ing_all = pd.read_excel(EXCEL_FILE, sheet_name="Ingresos")
            df_egr_all = pd.read_excel(EXCEL_FILE, sheet_name="Egresos")
        except Exception:
            df_ing_all = pd.DataFrame(columns=['Año', 'Mes', 'Tipo', 'Descripción', 'Monto', 'Moneda'])
            df_egr_all = pd.DataFrame(columns=['Año', 'Mes', 'Tipo', 'Descripción', 'Monto', 'Moneda'])

        df_ing_f = df_ing_all[(df_ing_all['Año'].isin(anios_sel)) & (df_ing_all['Mes'].isin(meses_sel))] if not df_ing_all.empty else pd.DataFrame()
        df_egr_f = df_egr_all[(df_egr_all['Año'].isin(anios_sel)) & (df_egr_all['Mes'].isin(meses_sel))] if not df_egr_all.empty else pd.DataFrame()

        ing_total_pen = df_ing_f[df_ing_f['Moneda'] == 'PEN']['Monto'].sum() if not df_ing_f.empty and 'Moneda' in df_ing_f.columns else 0
        egr_total_pen = df_egr_f[df_egr_f['Moneda'] == 'PEN']['Monto'].sum() if not df_egr_f.empty and 'Moneda' in df_egr_f.columns else 0
        superavit_pen = ing_total_pen - egr_total_pen
        tasa_ahorro_global = (superavit_pen / ing_total_pen * 100) if ing_total_pen > 0 else 0.0

        st.markdown("---")
        st.markdown("#### 📊 Resumen Estadístico Consolidado (Períodos Seleccionados)")
        
        kpi1, kpi2, kpi3, kpi4 = st.columns(4)
        kpi1.metric("Ingresos Totales (PEN)", f"S/ {ing_total_pen:,.2f}")
        kpi2.metric("Egresos Totales (PEN)", f"S/ {egr_total_pen:,.2f}")
        kpi3.metric("Superávit / Flujo Neto", f"S/ {superavit_pen:,.2f}", delta=f"{tasa_ahorro_global:.1f}% Ahorro")
        kpi4.metric("Tasa de Ahorro Neto", f"{tasa_ahorro_global:.1f}%", delta="Meta: ≥ 20%")

        st.markdown("---")
        st.markdown("#### 📈 Gráficas Comparativas por Mes y Año")
        
        if not df_ing_f.empty or not df_egr_f.empty:
            df_ing_grp = df_ing_f[df_ing_f['Moneda'] == 'PEN'].groupby('Mes')['Monto'].sum().reindex(MESES).fillna(0) if not df_ing_f.empty else pd.Series(0, index=MESES)
            df_egr_grp = df_egr_f[df_egr_f['Moneda'] == 'PEN'].groupby('Mes')['Monto'].sum().reindex(MESES).fillna(0) if not df_egr_f.empty else pd.Series(0, index=MESES)
            
            df_chart = pd.DataFrame({
                'Ingresos (PEN)': df_ing_grp,
                'Egresos (PEN)': df_egr_grp
            })
            
            df_chart_filtered = df_chart.loc[meses_sel]
            st.bar_chart(df_chart_filtered)
        else:
            st.info("No hay suficientes transacciones en PEN para los períodos seleccionados para mostrar gráficos.")

        st.markdown("---")
        st.markdown("#### 💡 Asesoría y Recomendaciones de Especialista en Finanzas Personales")
        
        recomendaciones = []
        if tasa_ahorro_global < 10:
            recomendaciones.append("🚨 **Alerta de Tasa de Ahorro Crítica:** Su tasa de ahorro actual está por debajo del 10%. Como especialista en finanzas personales, le sugiero aplicar la regla **50/30/20** (50% necesidades básicas, 30% gastos personales y 20% ahorro e inversión intocable).")
        elif 10 <= tasa_ahorro_global < 20:
            recomendaciones.append("⚠️ **Salud Financiera en Desarrollo:** Su tasa de ahorro es positiva pero moderada. Para alcanzar la estabilidad económica a largo plazo, intente recortar un 10% en categorías de gastos discrecionales (entretenimiento, compras extras).")
        else:
            recomendaciones.append("🌟 **Excelente Desempeño Financiero:** ¡Felicitaciones! Posee una sólida disciplina de ahorro superior al umbral recomendado del 20%. Es momento de evaluar instrumentos de inversión de renta fija o fondos mutuos para blindar su capital frente a la inflación.")

        if egr_total_pen > ing_total_pen:
            recomendaciones.append("📉 **Déficit Financiero Detectado:** En los períodos seleccionados sus egresos superan a sus ingresos. Es prioritario realizar una auditoría de sus gastos fijos en la pestaña *Egresos* y renegociar pasivos o deudas con tasas elevadas.")
        else:
            recomendaciones.append("🛡️ **Fondo de Emergencia:** Asegúrese de mantener un fondo de emergencia equivalente a 3 o 6 meses de sus egresos totales guardado en una cuenta líquida separada.")

        for rec in recomendaciones:
            st.info(rec)

        st.markdown("---")
        st.markdown("##### 📋 Detalle Analítico de Transacciones Filtradas")
        df_completo_filtrado = pd.concat([df_ing_f, df_egr_f], ignore_index=True) if not df_ing_f.empty or not df_egr_f.empty else pd.DataFrame()
        st.dataframe(df_completo_filtrado, use_container_width=True)
    else:
        st.warning("Por favor, seleccione al menos un año y un mes en los filtros superiores para generar el análisis financiero.")
