import streamlit as st
import pandas as pd
import estilos 
import procesador 
from io import BytesIO

# Configuración de la página
st.set_page_config(page_title="Sistema CARGO PESCA PRO", layout="wide")

# 1. Estética e Inicialización
estilos.aplicar_estilos()
estilos.mostrar_cabecera()

if "df_editada" not in st.session_state:
    st.session_state.df_editada = None
if "columnas_confirmadas" not in st.session_state:
    st.session_state.columnas_confirmadas = False
if "modo" not in st.session_state:
    st.session_state.modo = None

st.write("---")

# 2. Carga de Archivo (Soporte XLS y XLSX)
archivo = st.file_uploader("Subir archivo de CARGO PESCA (.xlsx, .xls)", type=["xlsx", "xls"])

if archivo:
    if not st.session_state.columnas_confirmadas:
        st.subheader("🛠️ Paso 1: Configurar Estructura de Datos")
        
        # Lectura inicial para referencia
        df_ref = pd.read_excel(archivo, header=None, nrows=15)
        
        col_c1, col_c2 = st.columns([1, 3])
        with col_c1:
            orientacion = st.radio("¿Dónde están los títulos?", ["En una Fila (Horizontal)", "En una Columna (Vertical)"])
            idx_titulo = st.number_input("Número de Fila/Columna de títulos:", 0, 14, 0)
            
            if st.button("✅ Confirmar y Procesar", type="primary"):
                if orientacion == "En una Fila (Horizontal)":
                    df_cargado = pd.read_excel(archivo, skiprows=idx_titulo)
                else:
                    # Lógica para títulos en columnas (transposición)
                    df_raw = pd.read_excel(archivo, header=None)
                    df_trans = df_raw.iloc[:, idx_titulo:].T
                    df_trans.columns = df_trans.iloc[0]
                    df_cargado = df_trans.drop(df_trans.index[0]).reset_index(drop=True)

                # Estandarización de nombres
                df_cargado.columns = [str(c).replace('.', '').strip().upper() for c in df_cargado.columns]
                
                # Normalización de Fecha
                c_fecha = next((c for c in df_cargado.columns if 'FECHA' in c), None)
                if c_fecha:
                    df_cargado.rename(columns={c_fecha: 'FECHA'}, inplace=True)
                    df_cargado['FECHA'] = pd.to_datetime(df_cargado['FECHA'], errors='coerce')
                
                st.session_state.df_editada = df_cargado.reset_index(drop=True)
                st.session_state.columnas_confirmadas = True
                st.rerun()
        
        with col_c2:
            st.write("Vista previa del archivo original:")
            st.dataframe(df_ref, use_container_width=True)

# 3. Interfaz de Gestión y Reportes
if st.session_state.columnas_confirmadas and st.session_state.df_editada is not None:
    df = st.session_state.df_editada

    # --- SECCIÓN: GESTIÓN Y EDICIÓN ---
    st.subheader("🔍 Paso 2: Gestión y Filtro por Fecha")
    
    col_g1, col_g2, col_g3 = st.columns([1, 1, 1])
    with col_g1:
        # Filtro específico de fecha solicitado
        fechas_disponibles = sorted(df['FECHA'].dt.date.dropna().unique())
        filtro_fecha = st.selectbox("📅 Filtrar por día específico:", ["Ver Todo"] + fechas_disponibles)
    with col_g2:
        busqueda_gen = st.text_input("🔎 Búsqueda general:")
    with col_g3:
        cols_visibles = st.multiselect("Columnas en tabla:", df.columns.tolist(), default=df.columns.tolist()[:6])

    # Aplicación de Filtros
    df_f = df.copy()
    if filtro_fecha != "Ver Todo":
        df_f = df_f[df_f['FECHA'].dt.date == filtro_fecha]
    if busqueda_gen:
        mask = df_f.apply(lambda x: x.astype(str).str.contains(busqueda_gen, case=False, na=False).any(), axis=1)
        df_f = df_f[mask]

    # Tabla Maestra con Selector
    df_sel = df_f[cols_visibles].copy()
    if 'FECHA' in df_sel.columns:
        df_sel['FECHA'] = df_sel['FECHA'].dt.strftime('%d/%m/%Y')
    df_sel.insert(0, "SELECCIONAR", False)
    
    editor = st.data_editor(df_sel, hide_index=False, use_container_width=True, key="main_editor")
    marcados = editor[editor["SELECCIONAR"] == True].index.tolist()

    # Botones de Gestión (Editar / Eliminar / Agregar)
    cg1, cg2, cg3 = st.columns(3)
    if cg1.button("🔧 Editar Marcados", disabled=not marcados):
        st.session_state.modo = "editar"
        st.session_state.indices_form = marcados
    if cg2.button("🗑️ Eliminar Marcados", type="primary", disabled=not marcados):
        st.session_state.df_editada = st.session_state.df_editada.drop(marcados).reset_index(drop=True)
        st.rerun()
    if cg3.button("➕ Nuevo Registro"):
        st.session_state.modo = "agregar"

    # Formularios (Edición/Agregar) se mantienen igual que la versión anterior...
    # [Insertar lógica de formularios aquí]

    # --- SECCIÓN: REPORTES MEJORADOS ---
    st.write("---")
    st.subheader("📊 Paso 3: Reportes de Totales Personalizados")
    
    tab_datos, tab_totales = st.tabs(["📄 Reporte de Datos", "📉 Reporte Estadístico por Columnas"])

    with tab_totales:
        st.info("Seleccione las columnas que desea incluir en el cálculo de totales.")
        if 'FECHA' in df.columns:
            # Selección de columnas para el reporte
            cols_reporte = st.multiselect(
                "Seleccionar columnas para totalizar:",
                options=[c for c in df.columns if c != 'FECHA'],
                default=df.columns.tolist()[1:5]
            )
            
            col_f1, col_f2 = st.columns(2)
            f_i = col_f1.date_input("Fecha Inicio", df['FECHA'].min().date())
            f_f = col_f2.date_input("Fecha Fin", df['FECHA'].max().date())
            
            df_r = df[(df['FECHA'].dt.date >= f_i) & (df['FECHA'].dt.date <= f_f)].copy()
            
            if not df_r.empty and cols_reporte:
                df_r['FECHA_TXT'] = df_r['FECHA'].dt.strftime('%d/%m/%Y')
                # Solo procesar las columnas seleccionadas por el usuario
                logica_agg = {c: ('sum' if pd.api.types.is_numeric_dtype(df_r[c]) else 'count') 
                             for c in cols_reporte}
                
                resumen = df_r.groupby('FECHA_TXT').agg(logica_agg).reset_index()
                resumen.rename(columns={'FECHA_TXT': 'FECHA'}, inplace=True)
                
                # Fila de Totales
                totales_fila = {"FECHA": "TOTAL GENERAL"}
                for c in cols_reporte: totales_fila[c] = resumen[c].sum()
                df_final_totales = pd.concat([resumen, pd.DataFrame([totales_fila])], ignore_index=True)
                
                # Visualización con celdas
                st.table(df_final_totales)

                if st.button("🚀 Exportar Cuadro de Totales"):
                    out = BytesIO()
                    with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
                        df_final_totales.to_excel(writer, index=False, sheet_name='Reporte')
                        workbook = writer.book
                        worksheet = writer.sheets['Reporte']
                        fmt = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
                        for i, col in enumerate(df_final_totales.columns):
                            worksheet.set_column(i, i, 18, fmt)
                    st.download_button("⬇️ Descargar Totales", out.getvalue(), "totales_personalizados.xlsx")