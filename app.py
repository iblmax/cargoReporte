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

# 2. Carga y Configuración Inicial
archivo = st.file_uploader("Subir archivo de CARGO PESCA (.xlsx, .xls)", type=["xlsx", "xls"])

if archivo and not st.session_state.columnas_confirmadas:
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
    
    st.data_editor(df_sel, hide_index=False, use_container_width=True, key="main_editor")

    # --- SECCIÓN: REPORTES ---
    st.write("---")
    st.subheader("📊 Paso 3: Reportes de Totales")
    
    tab_datos, tab_totales = st.tabs(["📄 Exportación de Datos", "📉 Cuadro Estadístico de Totales"])

    with tab_datos:
        st.write("### Configurar Reporte de Exportación")
        m_reporte = st.radio("Alcance:", ["Completo", "Filtrado por fecha actual"], horizontal=True)
        cols_exportar = st.multiselect("Columnas a exportar:", options=df.columns.tolist(), default=df.columns.tolist())
        
        df_export = df.copy() if m_reporte == "Completo" else df_f.copy()
        df_export = df_export[cols_exportar]
        if 'FECHA' in df_export.columns:
            df_export['FECHA'] = df_export['FECHA'].dt.strftime('%d/%m/%Y')
        st.dataframe(df_export, use_container_width=True)

    with tab_totales:
        st.write("### Cuadro Estadístico de Totales Personalizado")
        tipo_calculo = st.radio("Tipo de operación para el reporte:", 
                               ["Contar Registros (Suma 1 por dato)", "Sumar Cantidades (Suma los números de la columna)"],
                               horizontal=True)
        
        opciones_cols = [c for c in df.columns if c != 'FECHA']
        cols_stats = st.multiselect("Seleccionar columnas para totalizar:", options=opciones_cols)
        
        c_f1, c_f2 = st.columns(2)
        f_i = c_f1.date_input("📅 Fecha Inicio", df['FECHA'].min().date() if not df.empty else datetime.now())
        f_f = c_f2.date_input("📅 Fecha Hasta", df['FECHA'].max().date() if not df.empty else datetime.now())
        
        df_r = df[(df['FECHA'].dt.date >= f_i) & (df['FECHA'].dt.date <= f_f)].copy()
        
        if not df_r.empty and cols_stats:
            df_r['FECHA_TXT'] = df_r['FECHA'].dt.strftime('%d/%m/%Y')
            
            # --- CORRECCIÓN CRÍTICA DE SUMA ---
            if "Sumar Cantidades" in tipo_calculo:
                for col in cols_stats:
                    # Elimina comas y convierte a número para que la suma no dé 0.0000
                    df_r[col] = pd.to_numeric(df_r[col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
                agg_map = {c: 'sum' for c in cols_stats}
            else:
                agg_map = {c: 'count' for c in cols_stats}
                
            resumen = df_r.groupby('FECHA_TXT').agg(agg_map).reset_index()
            resumen.rename(columns={'FECHA_TXT': 'FECHA'}, inplace=True)
            
            # Fila de Totales
            tot_row = {"FECHA": "TOTAL GENERAL"}
            for c in cols_stats: tot_row[c] = resumen[c].sum()
            df_final_totales = pd.concat([resumen, pd.DataFrame([tot_row])], ignore_index=True)
            
            st.table(df_final_totales)