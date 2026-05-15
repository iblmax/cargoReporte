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
if "indices_editar" not in st.session_state: st.session_state.indices_editar = []
if "historial" not in st.session_state: st.session_state.historial = []
if "archivo_actual" not in st.session_state: st.session_state.archivo_actual = None
if "numero_reporte" not in st.session_state: st.session_state.numero_reporte = None

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

        fmt_titulo = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center'})
        fmt_header = workbook.add_format({'bold': True, 'bg_color': '#0070C0', 'font_color': 'white', 'border': 1, 'align': 'center'})
        fmt_celda = workbook.add_format({'border': 1, 'align': 'center'})

        worksheet.merge_range(1, 0, 1, len(df.columns)-1, titulo_reporte.upper(), fmt_titulo)

        for col_num, value in enumerate(df.columns.values):
            worksheet.write(3, col_num, value, fmt_header)
            worksheet.set_column(col_num, col_num, 18)

        worksheet.autofilter(3, 0, 3, len(df.columns)-1)
        
        for row in range(len(df_clean)):
            for col in range(len(df_clean.columns)):
                worksheet.write(row + 4, col, df_clean.iloc[row, col], fmt_celda)
                
    return output.getvalue()

# --- SUBIDA DE ARCHIVO Y NÚMERO DE REPORTE ---
archivo = st.file_uploader("Subir archivo de CARGO PESCA", type=["xlsx", "xls"], key="file_uploader_principal")

if archivo:
    if st.session_state.archivo_actual != archivo.name:
        st.session_state.archivo_actual = archivo.name
        st.session_state.numero_reporte = f"REP-{datetime.now().strftime('%d%m%Y')}-{datetime.now().strftime('%H%M%S')}"
        st.session_state.columnas_confirmadas = False
        st.session_state.df_editada = None

    st.success(f"📑 Número de Reporte generado: {st.session_state.numero_reporte}")
    if st.button("🔖 Mostrar Número de Reporte", key="btn_mostrar_reporte"):
        st.info(f"Este archivo corresponde al reporte: **{st.session_state.numero_reporte}**")


# --- 1. CONFIGURAR ESTRUCTURA ---
if archivo and not st.session_state.columnas_confirmadas:
    try:
        st.subheader("🛠️ Paso 1: Configurar Estructura")
        df_ref = pd.read_excel(archivo, header=None, nrows=15)
        
        c1, c2 = st.columns([1, 2])
        with c1:
            orientacion = st.radio("📑 Títulos en:", ["Horizontal (Fila)", "Vertical (Columna)"], key="radio_orientacion")
            idx_titulo = st.number_input("Número de Índice:", 0, 14, 0, key="num_indice")
            
            if st.button("✅ Confirmar Estructura", type="primary", key="btn_confirmar_estructura"):
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
    except Exception as e:
        st.error(f"❌ Error al procesar archivo: {e}")
# --- 2. GESTIÓN Y EDICIÓN ---
if st.session_state.columnas_confirmadas:
    df = st.session_state.df_editada
    st.subheader("🔍 Gestión y Edición")

    # Filtro por fecha
    if 'FECHA' in df.columns:
        fecha_min, fecha_max = df['FECHA'].min(), df['FECHA'].max()
        rango = st.date_input("📅 Rango de fechas:", [fecha_min, fecha_max], key="date_rango")
        if len(rango) == 2:
            df = df[(df['FECHA'] >= pd.to_datetime(rango[0])) & (df['FECHA'] <= pd.to_datetime(rango[1]))]

    # Barra de búsqueda con autocompletado
    query = st.text_input("🔎 Buscar por nombre o número de guía:", key="txt_busqueda")
    sugerencias = []
    if query:
        for col in ["NO. GUIA", "CLIENTE", "CHOFER"]:
            if col in df.columns:
                coincidencias = df[df[col].astype(str).str.contains(query, case=False)][col].unique().tolist()
                sugerencias.extend(coincidencias)
        sugerencias = list(set(sugerencias))
    if sugerencias:
        seleccion = st.selectbox("Coincidencias encontradas:", sugerencias, key="selectbox_busqueda")
        df = df[df.apply(lambda row: seleccion in row.values, axis=1)]

    # Selección de columnas
    cols_v = st.multiselect("Seleccionar columnas de trabajo:", df.columns.tolist(), default=df.columns.tolist()[:7], key="multiselect_columnas")

    df_disp = df.copy()
    if 'FECHA' in df_disp.columns:
        df_disp['FECHA_V'] = df_disp['FECHA'].dt.strftime('%d/%m/%Y')
    
    df_disp.insert(0, "SELEC", False)
    cols_finales = ["SELEC"] + ([c for c in cols_v if c != 'FECHA'] + (['FECHA_V'] if 'FECHA' in df.columns else []))
    editor = st.data_editor(df_disp[cols_finales], hide_index=True, use_container_width=True, key="editor_tabla")
    
    indices = editor[editor["SELEC"] == True].index.tolist()

    bg1, bg2, bg3 = st.columns(3)
    if bg1.button("🔧 Editar Registro", disabled=not indices, key="btn_editar"):
        st.session_state.modo = "editar"; st.session_state.indices_editar = indices
    if bg2.button("🗑️ Eliminar Registro", type="primary", disabled=not indices, key="btn_eliminar"):
        st.session_state.modo = "confirmar_borrado"; st.session_state.indices_borrar = indices
    if bg3.button("➕ Agregar Registro", key="btn_agregar"):
        st.session_state.modo = "nuevo"
# --- FORMULARIO DE EDICIÓN / NUEVO ---
if st.session_state.modo in ["editar", "nuevo"]:
    with st.expander("📝 Formulario de Registro", expanded=True):
        with st.form("form_gestion", clear_on_submit=False):
            nuevos = {}
            f_cols = st.columns(3)

            # --- NO. GUIA ---
            if st.session_state.modo == "editar" and len(st.session_state.indices_editar) > 1 and "NO. GUIA" in df.columns:
                guias = df.loc[st.session_state.indices_editar, "NO. GUIA"].tolist()
                guia_sel = f_cols[0].selectbox("NO. GUIA", guias, key="selectbox_guia")

                # Guardar el índice del registro seleccionado
                idx_sel = df[df["NO. GUIA"] == guia_sel].index[0]
                st.session_state.registro_a_editar = idx_sel
                nuevos["NO. GUIA"] = guia_sel
            else:
                idx_sel = st.session_state.indices_editar[0] if st.session_state.modo == "editar" and st.session_state.indices_editar else None
                st.session_state.registro_a_editar = idx_sel
                val_guia = df.loc[idx_sel, "NO. GUIA"] if st.session_state.modo == "editar" and idx_sel is not None else ""
                nuevos["NO. GUIA"] = f_cols[0].text_input("NO. GUIA", value=str(val_guia), key="textinput_guia")

            # --- OTRAS COLUMNAS ---
            for i, col in enumerate([c for c in cols_v if c != "NO. GUIA"]):
                val = df.loc[st.session_state.registro_a_editar, col] if st.session_state.modo == "editar" and st.session_state.registro_a_editar is not None else ""
                if col == "FECHA" and val != "":
                    try:
                        val = pd.to_datetime(val).strftime("%d/%m/%Y")
                    except:
                        val = str(val)
                nuevos[col] = f_cols[(i+1) % 3].text_input(col, value=str(val), key=f"textinput_{col}")

            # --- BOTONES ---
            c_f1, c_f2 = st.columns(2)
            if c_f1.form_submit_button("💾 Guardar Cambios", key="btn_guardar"):
                fila_base = df.loc[st.session_state.registro_a_editar].to_dict() if st.session_state.modo == "editar" and st.session_state.registro_a_editar is not None else {c: "" for c in df.columns}
                fila_base.update(nuevos)

                # Convertir FECHA al guardar
                if "FECHA" in fila_base and fila_base["FECHA"] != "":
                    try:
                        fila_base["FECHA"] = pd.to_datetime(fila_base["FECHA"], format="%d/%m/%Y", errors="coerce")
                    except:
                        pass

                if st.session_state.modo == "editar":
                    st.session_state.df_editada.loc[st.session_state.registro_a_editar] = pd.Series(fila_base)
                    st.session_state.historial.append(f"Editado registro {st.session_state.registro_a_editar}")
                else:
                    st.session_state.df_editada = pd.concat([df, pd.DataFrame([fila_base])], ignore_index=True)
                    st.session_state.historial.append("Agregado nuevo registro")

                st.session_state.modo = None
                st.rerun()

            if c_f2.form_submit_button("🧹 Limpiar y Cerrar", key="btn_limpiar"):
                st.session_state.modo = None
                st.rerun()

    # --- ELIMINACIÓN ---
    if st.session_state.modo == "confirmar_borrado":
        if st.button("⚠️ Confirmar Eliminación", key="btn_confirmar_borrado"):
            st.session_state.df_editada = df.drop(st.session_state.indices_borrar).reset_index(drop=True)
            st.session_state.historial.append(f"Eliminados registros {st.session_state.indices_borrar}")
            st.session_state.modo = None
            st.rerun()
        if st.button("❌ Cancelar", key="btn_cancelar_borrado"):
            st.session_state.modo = None
            st.rerun()

# --- REPORTES Y ESTADÍSTICAS ---
if st.session_state.columnas_confirmadas:
    st.write("---")
    st.subheader("📊 Paso 3: Reportes de Totales")
    t1, t2, t3 = st.tabs(["📄 Exportación de Datos", "📉 Cuadro Estadístico de Totales", "📈 Gráficos Interactivos"])

    with t1:
        tipo_exp = st.radio("Alcance:", ["Completo", "Filtrado por columna"], horizontal=True, key="radio_export")
        df_e = st.session_state.df_editada.copy()
        if tipo_exp == "Filtrado por columna":
            c_e = st.multiselect("Columnas de exportación:", st.session_state.df_editada.columns.tolist(), default=st.session_state.df_editada.columns.tolist(), key="multiselect_export")
            df_e = df_e[c_e]
        
        if 'FECHA' in df_e.columns: 
            df_e['FECHA'] = df_e['FECHA'].dt.strftime('%d/%m/%Y')
        st.dataframe(df_e, use_container_width=True)
        
        btn_e = descargar_excel_profesional(df_e, "Reporte de Operaciones")
        st.download_button("🚀 Generar Excel de Datos", btn_e, "Reporte_Cargo.xlsx", key="btn_excel_datos")

    with t2:
        st.markdown("<h3 style='text-align: center;'>CUADRO ESTADÍSTICO DE OPERACIONES</h3>", unsafe_allow_html=True)
        op_calc = st.radio("Cálculo:", ["Contar Registros", "Sumar Cantidades"], horizontal=True, key="radio_calc")
        c_stats = st.multiselect("Columnas para totalizar:", [c for c in st.session_state.df_editada.columns if c != 'FECHA'], key="multiselect_stats")
        
        if c_stats and 'FECHA' in st.session_state.df_editada.columns:
            res = st.session_state.df_editada.copy()
            if "Sumar" in op_calc:
                for c in c_stats: 
                    res[c] = pd.to_numeric(res[c].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
            
            resumen = res.groupby(res['FECHA'].dt.date)[c_stats].agg('sum' if "Sumar" in op_calc else 'count').reset_index()
            
            tot_vals = {col: resumen[col].sum() for col in c_stats}
            tot_vals['FECHA'] = "TOTAL GENERAL"
            res_final = pd.concat([resumen, pd.DataFrame([tot_vals])], ignore_index=True)
            res_final['FECHA'] = res_final['FECHA'].apply(lambda x: x.strftime('%d/%m/%Y') if hasattr(x, 'strftime') else x)
            
            _, col_mid, _ = st.columns([1, 4, 1])
            with col_mid:
                st.table(res_final)
            
            btn_s = descargar_excel_profesional(res_final, "Cuadro Estadístico de Totales")
            st.download_button("🚀 Exportar Totales Profesionales", btn_s, "Estadisticas_Cargo.xlsx", key="btn_excel_totales")

    with t3:
        st.markdown("<h3 style='text-align: center;'>📈 Gráficos Interactivos</h3>", unsafe_allow_html=True)
        if 'FECHA' in st.session_state.df_editada.columns:
            c_stats_chart = st.multiselect("Columnas para gráficos:", [c for c in st.session_state.df_editada.columns if c != 'FECHA'], key="multiselect_chart")
            if c_stats_chart:
                res = st.session_state.df_editada.copy()
                resumen = res.groupby(res['FECHA'].dt.date)[c_stats_chart].agg('sum').reset_index()
                chart_data = resumen.set_index("FECHA")
                st.bar_chart(chart_data)
                st.line_chart(chart_data)

    # --- HISTORIAL DE CAMBIOS ---
    st.write("---")
    st.subheader("📜 Historial de Cambios")
    if st.session_state.historial:
        for h in st.session_state.historial:
            st.write(f"- {h}")
    else:
        st.info("No se han registrado cambios aún.")
