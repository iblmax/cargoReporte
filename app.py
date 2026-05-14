import streamlit as st
import pandas as pd
import estilos 
from io import BytesIO
from datetime import datetime

# Configuración de página
st.set_page_config(page_title="Sistema CARGO PESCA PRO", layout="wide")

# Inicialización de estados
if "df_editada" not in st.session_state: st.session_state.df_editada = None
if "columnas_confirmadas" not in st.session_state: st.session_state.columnas_confirmadas = False
if "modo" not in st.session_state: st.session_state.modo = None
if "registro_a_editar" not in st.session_state: st.session_state.registro_a_editar = None

estilos.aplicar_estilos()
estilos.mostrar_cabecera()

# --- FUNCIÓN DE EXCEL PROFESIONAL ---
def descargar_excel_profesional(df, titulo_reporte):
    output = BytesIO()
    df_clean = df.fillna("")
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_clean.to_excel(writer, index=False, sheet_name='Reporte', startrow=3)
        workbook = writer.book
        worksheet = writer.sheets['Reporte']

        # Estilos: Azul con blanco y bordes
        fmt_titulo = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center'})
        fmt_header = workbook.add_format({'bold': True, 'bg_color': '#0070C0', 'font_color': 'white', 'border': 1, 'align': 'center'})
        fmt_celda = workbook.add_format({'border': 1, 'align': 'center'})

        worksheet.merge_range(1, 0, 1, len(df.columns)-1, titulo_reporte.upper(), fmt_titulo)

        for col_num, value in enumerate(df.columns.values):
            worksheet.write(3, col_num, value, fmt_header)
            worksheet.set_column(col_num, col_num, 18)

        # ✅ Autofiltro en cabecera
        worksheet.autofilter(3, 0, 3, len(df.columns)-1)
        
        for row in range(len(df_clean)):
            for col in range(len(df_clean.columns)):
                worksheet.write(row + 4, col, df_clean.iloc[row, col], fmt_celda)
                
    return output.getvalue()

# --- 1. SELECCIÓN DE TÍTULOS CON VISTA PREVIA ---
archivo = st.file_uploader("Subir archivo de CARGO PESCA", type=["xlsx", "xls"])

if archivo and not st.session_state.columnas_confirmadas:
    st.subheader("🛠️ Paso 1: Configurar Estructura")
    df_ref = pd.read_excel(archivo, header=None, nrows=15)
    
    c1, c2 = st.columns([1, 2])
    with c1:
        orientacion = st.radio("📑 Títulos en:", ["Horizontal (Fila)", "Vertical (Columna)"])
        idx_titulo = st.number_input("Número de Índice:", 0, 14, 0)
        
        if st.button("✅ Confirmar Estructura", type="primary"):
            if orientacion == "Horizontal (Fila)":
                df_res = pd.read_excel(archivo, skiprows=idx_titulo)
            else:
                df_raw = pd.read_excel(archivo, header=None)
                df_res = df_raw.iloc[:, idx_titulo:].T
                df_res.columns = df_res.iloc[0]
                df_res = df_res.drop(df_res.index[0]).reset_index(drop=True)

            df_res.columns = [str(c).strip().upper() for c in df_res.columns]
            c_f = next((c for c in df_res.columns if 'FECHA' in c), None)
            if c_f:
                df_res.rename(columns={c_f: 'FECHA'}, inplace=True)
                df_res['FECHA'] = pd.to_datetime(df_res['FECHA'], errors='coerce')
            
            st.session_state.df_editada = df_res
            st.session_state.columnas_confirmadas = True
            st.rerun()

    with c2:
        def highlight(x):
            color = 'background-color: #d1e7dd'
            ds = pd.DataFrame('', index=x.index, columns=x.columns)
            if orientacion == "Horizontal (Fila)":
                if idx_titulo in x.index: ds.loc[idx_titulo, :] = color
            else:
                if idx_titulo in x.columns: ds.loc[:, idx_titulo] = color
            return ds
        st.write("Vista Previa (Marcado en verde):")
        st.dataframe(df_ref.style.apply(highlight, axis=None), use_container_width=True)

# --- INTERFAZ DE GESTIÓN ---
if st.session_state.columnas_confirmadas:
    df = st.session_state.df_editada

    # 2. SELECCIÓN DE COLUMNAS
    st.subheader("🔍 Gestión y Edición")
    cols_v = st.multiselect("Seleccionar columnas de trabajo:", df.columns.tolist(), default=df.columns.tolist()[:7])

    df_disp = df.copy()
    if 'FECHA' in df_disp.columns:
        df_disp['FECHA_V'] = df_disp['FECHA'].dt.strftime('%d/%m/%Y')
    
    df_disp.insert(0, "SELEC", False)
    cols_finales = ["SELEC"] + ([c for c in cols_v if c != 'FECHA'] + (['FECHA_V'] if 'FECHA' in df.columns else []))
    editor = st.data_editor(df_disp[cols_finales], hide_index=True, use_container_width=True)
    
    indices = editor[editor["SELEC"] == True].index.tolist()

    bg1, bg2, bg3 = st.columns(3)
    if bg1.button("🔧 Editar Registro", disabled=not indices):
        st.session_state.modo = "editar"; st.session_state.registro_a_editar = indices[0]
    if bg2.button("🗑️ Eliminar Registro", type="primary", disabled=not indices):
        st.session_state.modo = "confirmar_borrado"; st.session_state.indices_borrar = indices
    if bg3.button("➕ Agregar Registro"):
        st.session_state.modo = "nuevo"

    # FORMULARIO CON COLUMNAS SELECCIONADAS
    if st.session_state.modo in ["editar", "nuevo"]:
        with st.expander("📝 Formulario de Registro", expanded=True):
            with st.form("form_gestion"):
                nuevos = {}
                f_cols = st.columns(3)
                for i, col in enumerate(cols_v):
                    val = df.loc[st.session_state.registro_a_editar, col] if st.session_state.modo == "editar" else ""
                    nuevos[col] = f_cols[i % 3].text_input(col, value=str(val))
                
                c_f1, c_f2 = st.columns(2)
                if c_f1.form_submit_button("💾 Guardar Cambios"):
                    fila_base = df.loc[st.session_state.registro_a_editar].to_dict() if st.session_state.modo == "editar" else {c: "" for c in df.columns}
                    fila_base.update(nuevos)
                    
                    if st.session_state.modo == "editar":
                        st.session_state.df_editada.loc[st.session_state.registro_a_editar] = pd.Series(fila_base)
                    else:
                        st.session_state.df_editada = pd.concat([df, pd.DataFrame([fila_base])], ignore_index=True)
                    
                    st.session_state.modo = None; st.rerun()
                
                if c_f2.form_submit_button("🧹 Limpiar y Cerrar"):
                    st.session_state.modo = None; st.rerun()

    if st.session_state.modo == "confirmar_borrado":
        if st.button("⚠️ Confirmar Eliminación"):
            st.session_state.df_editada = df.drop(st.session_state.indices_borrar).reset_index(drop=True)
            st.session_state.modo = None; st.rerun()
        if st.button("❌ Cancelar"):
            st.session_state.modo = None; st.rerun()

    # --- REPORTES Y ESTADÍSTICAS ---
    st.write("---")
    st.subheader("📊 Paso 3: Reportes de Totales")
    t1, t2 = st.tabs(["📄 Exportación de Datos", "📉 Cuadro Estadístico de Totales"])

    with t1:
        tipo_exp = st.radio("Alcance:", ["Completo", "Filtrado por columna"], horizontal=True)
        df_e = df.copy()
        if tipo_exp == "Filtrado por columna":
            c_e = st.multiselect("Columnas de exportación:", df.columns.tolist(), default=df.columns.tolist())
            df_e = df_e[c_e]
        
        if 'FECHA' in df_e.columns: df_e['FECHA'] = df_e['FECHA'].dt.strftime('%d/%m/%Y')
        st.dataframe(df_e, use_container_width=True)
        
        btn_e = descargar_excel_profesional(df_e, "Reporte de Operaciones")
        st.download_button("🚀 Generar Excel de Datos", btn_e, "Reporte_Cargo.xlsx")

    with t2:
        st.markdown("<h3 style='text-align: center;'>CUADRO ESTADÍSTICO DE OPERACIONES</h3>", unsafe_allow_html=True)
        op_calc = st.radio("Cálculo:", ["Contar Registros", "Sumar Cantidades"], horizontal=True)
        c_stats = st.multiselect("Columnas para totalizar:", [c for c in df.columns if c != 'FECHA'])
        
        if c_stats and 'FECHA' in df.columns:
            res = df.copy()
            if "Sumar" in op_calc:
                for c in c_stats: res[c] = pd.to_numeric(res[c].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
            
            resumen = res.groupby(res['FECHA'].dt.date)[c_stats].agg('sum' if "Sumar" in op_calc else 'count').reset_index()
            
            # Fila de Total General
            tot_vals = {col: resumen[col].sum() for col in c_stats}
            tot_vals['FECHA'] = "TOTAL GENERAL"
            res_final = pd.concat([resumen, pd.DataFrame([tot_vals])], ignore_index=True)
            res_final['FECHA'] = res_final['FECHA'].apply(lambda x: x.strftime('%d/%m/%Y') if hasattr(x, 'strftime') else x)
            
            _, col_mid, _ = st.columns([1, 4, 1])
            with col_mid:
                st.table(res_final)
            
            btn_s = descargar_excel_profesional(res_final, "Cuadro Estadístico de Totales")
            st.download_button("🚀 Exportar Totales Profesionales", btn_s, "Estadisticas_Cargo.xlsx")