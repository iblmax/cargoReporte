import streamlit as st
import pandas as pd
import estilos 
import procesador 
from io import BytesIO
from datetime import datetime

# Configuración de la página
st.set_page_config(page_title="Sistema CARGO PESCA PRO", layout="wide")

# 1. Estética e Inicialización
estilos.aplicar_estilos()
estilos.mostrar_cabecera()

if "df_editada" not in st.session_state:
    st.session_state.df_editada = None
if "columnas_confirmadas" not in st.session_state:
    st.session_state.columnas_confirmadas = False

st.write("---")

# 2. Carga de Archivo (Soporte XLS y XLSX)
archivo = st.file_uploader("Subir archivo de CARGO PESCA (.xlsx, .xls)", type=["xlsx", "xls"])

if archivo:
    if not st.session_state.columnas_confirmadas:
        st.subheader("🛠️ Paso 1: Configurar Estructura de Datos")
        df_ref = pd.read_excel(archivo, header=None, nrows=20)
        
        col_c1, col_c2 = st.columns([1, 3])
        with col_c1:
            orientacion = st.radio("📑 Orientación de títulos:", ["Horizontal (En Fila)", "Vertical (En Columna)"])
            idx_titulo = st.number_input("Índice de ubicación:", 0, 19, 0)
            
            if st.button("✅ Confirmar Estructura", type="primary"):
                if orientacion == "Horizontal (En Fila)":
                    df_cargado = pd.read_excel(archivo, skiprows=idx_titulo)
                else:
                    df_raw = pd.read_excel(archivo, header=None)
                    df_trans = df_raw.iloc[:, idx_titulo:].T
                    df_trans.columns = df_trans.iloc[0]
                    df_cargado = df_trans.drop(df_trans.index[0]).reset_index(drop=True)

                df_cargado.columns = [str(c).replace('.', '').strip().upper() for c in df_cargado.columns]
                c_fecha = next((c for c in df_cargado.columns if 'FECHA' in c), None)
                if c_fecha:
                    df_cargado.rename(columns={c_fecha: 'FECHA'}, inplace=True)
                    df_cargado['FECHA'] = pd.to_datetime(df_cargado['FECHA'], errors='coerce')
                
                st.session_state.df_editada = df_cargado.reset_index(drop=True)
                st.session_state.columnas_confirmadas = True
                st.rerun()
        
        with col_c2:
            def highlight_selection(x):
                color = 'background-color: #d1e7dd'
                df_style = pd.DataFrame('', index=x.index, columns=x.columns)
                if orientacion == "Horizontal (En Fila)":
                    if idx_titulo in x.index: df_style.loc[idx_titulo, :] = color
                else:
                    if idx_titulo in x.columns: df_style.loc[:, idx_titulo] = color
                return df_style
            st.dataframe(df_ref.style.apply(highlight_selection, axis=None), use_container_width=True)

# 3. Interfaz de Gestión y Reportes
if st.session_state.columnas_confirmadas and st.session_state.df_editada is not None:
    df = st.session_state.df_editada

    # --- SECCIÓN: GESTIÓN CON CALENDARIO ---
    st.subheader("🔍 Paso 2: Gestión y Edición")
    col_g1, col_g2, col_g3 = st.columns([1, 1, 1])
    with col_g1:
        # Selector de fecha estilo calendario visual
        f_dia = st.date_input("📅 Seleccionar Día para Gestionar:", 
                             value=df['FECHA'].min().date() if not df.empty else datetime.now())
    with col_g2:
        busqueda_gen = st.text_input("🔎 Búsqueda rápida:")
    with col_g3:
        cols_visibles = st.multiselect("Columnas en tabla:", df.columns.tolist(), default=df.columns.tolist()[:6])

    df_f = df[df['FECHA'].dt.date == f_dia].copy()
    if busqueda_gen:
        mask = df_f.apply(lambda x: x.astype(str).str.contains(busqueda_gen, case=False, na=False).any(), axis=1)
        df_f = df_f[mask]

    df_sel = df_f[cols_visibles].copy()
    if 'FECHA' in df_sel.columns:
        df_sel['FECHA'] = df_sel['FECHA'].dt.strftime('%d/%m/%Y')
    df_sel.insert(0, "SELECCIONAR", False)
    
    editor = st.data_editor(df_sel, hide_index=False, use_container_width=True, key="main_editor")
    marcados = editor[editor["SELECCIONAR"] == True].index.tolist()

    # --- SECCIÓN: REPORTES INTEGRADOS ---
    st.write("---")
    st.subheader("📊 Paso 3: Reportes de Totales")
    
    tab_datos, tab_totales = st.tabs(["📄 Exportación de Datos", "📉 Cuadro Estadístico de Totales"])

    with tab_datos:
        st.write("### Configurar Reporte de Exportación")
        m_reporte = st.radio("Alcance:", ["Completo", "Filtrado por fecha actual"], horizontal=True)
        
        # MEJORA: Selección de columnas para la exportación
        cols_exportar = st.multiselect("Seleccionar columnas a exportar:", 
                                      options=df.columns.tolist(), 
                                      default=df.columns.tolist())
        
        df_export = df.copy() if m_reporte == "Completo" else df_f.copy()
        df_export = df_export[cols_exportar]
        
        if 'FECHA' in df_export.columns:
            df_export['FECHA'] = df_export['FECHA'].dt.strftime('%d/%m/%Y')
            
        st.dataframe(df_export, use_container_width=True)
        
        if st.button("🚀 Generar Excel de Datos"):
            buf = procesador.generar_excel_formateado(df_export)
            st.download_button("⬇️ Descargar Reporte", buf, "reporte_datos.xlsx")

    with tab_totales:
        st.write("### Cuadro Estadístico de Totales Personalizado")
        
        # MEJORA: Opción de Tipo de Cálculo (Sumar 1 o Sumar Valores)
        tipo_calculo = st.radio("Tipo de operación para el reporte:", 
                               ["Contar Registros (Suma 1 por dato)", "Sumar Cantidades (Suma los números de la columna)"],
                               horizontal=True)
        
        opciones_cols = [c for c in df.columns if c != 'FECHA']
        cols_stats = st.multiselect("Seleccionar columnas para totalizar:", 
                                   options=opciones_cols, 
                                   default=opciones_cols[:4] if len(opciones_cols) >= 4 else opciones_cols)
        
        c_f1, c_f2 = st.columns(2)
        f_i = c_f1.date_input("📅 Fecha Inicio", df['FECHA'].min().date())
        f_f = c_f2.date_input("📅 Fecha Hasta", df['FECHA'].max().date())
        
        df_r = df[(df['FECHA'].dt.date >= f_i) & (df['FECHA'].dt.date <= f_f)].copy()
        
        if not df_r.empty and cols_stats:
            df_r['FECHA_TXT'] = df_r['FECHA'].dt.strftime('%d/%m/%Y')
            
            # Aplicación de lógica de cálculo seleccionada
            if "Contar" in tipo_calculo:
                agg_map = {c: 'count' for c in cols_stats}
            else:
                # Asegurar que las columnas sean numéricas para sumar valores
                for c in cols_stats:
                    df_r[c] = pd.to_numeric(df_r[c], errors='coerce').fillna(0)
                agg_map = {c: 'sum' for c in cols_stats}
                
            resumen = df_r.groupby('FECHA_TXT').agg(agg_map).reset_index()
            resumen.rename(columns={'FECHA_TXT': 'FECHA'}, inplace=True)
            
            # Fila de Total General con bordes y celdas ajustadas
            tot_row = {"FECHA": "TOTAL GENERAL"}
            for c in cols_stats: tot_row[c] = resumen[c].sum()
            df_final_totales = pd.concat([resumen, pd.DataFrame([tot_row])], ignore_index=True)
            
            st.table(df_final_totales)

            if st.button("🚀 Exportar Cuadro de Totales Profesional"):
                out = BytesIO()
                with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
                    df_final_totales.to_excel(writer, index=False, sheet_name='Totales')
                    workbook, worksheet = writer.book, writer.sheets['Totales']
                    # Formato con bordes y alineación central
                    fmt = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
                    for i, col in enumerate(df_final_totales.columns):
                        worksheet.set_column(i, i, 18, fmt)
                st.download_button("⬇️ Descargar Cuadro de Totales", out.getvalue(), "totales_pesca.xlsx")