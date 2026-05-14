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

# Inicialización de estados
if "df_editada" not in st.session_state:
    st.session_state.df_editada = None
if "columnas_confirmadas" not in st.session_state:
    st.session_state.columnas_confirmadas = False
if "modo" not in st.session_state:
    st.session_state.modo = None
if "registro_a_editar" not in st.session_state:
    st.session_state.registro_a_editar = None

st.write("---")

# 2. CARGA Y CONFIGURACIÓN INICIAL (ELECCIÓN DE FILA/COLUMNA)
archivo = st.file_uploader("Subir archivo de CARGO PESCA (.xlsx, .xls)", type=["xlsx", "xls"])

if archivo and not st.session_state.columnas_confirmadas:
    st.subheader("🛠️ Paso 1: Configurar Estructura de Datos")
    # Vista previa para elegir títulos
    df_ref = pd.read_excel(archivo, header=None, nrows=20)
    
    col_c1, col_c2 = st.columns([1, 3])
    with col_c1:
        # Aquí eliges si quieres Fila o Columna
        orientacion = st.radio("📑 Títulos en:", ["Horizontal (Fila)", "Vertical (Columna)"])
        idx_titulo = st.number_input("Índice de ubicación:", 0, 19, 0)
        
        if st.button("✅ Confirmar Estructura", type="primary"):
            if orientacion == "Horizontal (Fila)":
                df_cargado = pd.read_excel(archivo, skiprows=idx_titulo)
            else:
                df_raw = pd.read_excel(archivo, header=None)
                df_trans = df_raw.iloc[:, idx_titulo:].T
                df_trans.columns = df_trans.iloc[0]
                df_cargado = df_trans.drop(df_trans.index[0]).reset_index(drop=True)

            # Limpiar nombres de columnas
            df_cargado.columns = [str(c).replace('.', '').strip().upper() for c in df_cargado.columns]
            
            # Estandarizar Fecha
            c_fecha = next((c for c in df_cargado.columns if 'FECHA' in c), None)
            if c_fecha:
                df_cargado.rename(columns={c_fecha: 'FECHA'}, inplace=True)
                df_cargado['FECHA'] = pd.to_datetime(df_cargado['FECHA'], errors='coerce')
            
            st.session_state.df_editada = df_cargado.reset_index(drop=True)
            st.session_state.columnas_confirmadas = True
            st.rerun()

    with col_c2:
        # Marcación en verde de la selección
        def highlight_selection(x):
            color = 'background-color: #d1e7dd'
            df_style = pd.DataFrame('', index=x.index, columns=x.columns)
            if orientacion == "Horizontal (Fila)":
                if idx_titulo in x.index: df_style.loc[idx_titulo, :] = color
            else:
                if idx_titulo in x.columns: df_style.loc[:, idx_titulo] = color
            return df_style
        st.dataframe(df_ref.style.apply(highlight_selection, axis=None), use_container_width=True)

# 3. INTERFAZ DE GESTIÓN (EDICIÓN Y BOTONES)
if st.session_state.columnas_confirmadas and st.session_state.df_editada is not None:
    df = st.session_state.df_editada

    st.subheader("🔍 Paso 2: Gestión y Edición")
    
    # Calendario visual para gestión
    f_dia = st.date_input("📅 Seleccionar Día para Gestionar:", 
                         value=df['FECHA'].min().date() if not df.empty else datetime.now())
    
    # Tabla con selector
    df_f = df[df['FECHA'].dt.date == f_dia].copy()
    df_f.insert(0, "SELECCIONAR", False)
    
    editor = st.data_editor(df_f, hide_index=False, use_container_width=True, key="main_editor")
    marcados = editor[editor["SELECCIONAR"] == True].index.tolist()

    # Botones de Acción
    cg1, cg2, cg3 = st.columns(3)
    if cg1.button("🔧 Editar Registro", disabled=not marcados):
        st.session_state.modo = "editar"
        st.session_state.registro_a_editar = marcados[0]
    
    if cg2.button("🗑️ Eliminar Selección", type="primary", disabled=not marcados):
        st.session_state.modo = "confirmar_borrado"
        st.session_state.indices_borrar = marcados

    if cg3.button("➕ Nuevo Registro"):
        st.session_state.modo = "nuevo"

    # FORMULARIOS PARA EDITAR O AGREGAR
    if st.session_state.modo in ["editar", "nuevo"]:
        with st.form("form_datos"):
            st.write(f"### {'Editando' if st.session_state.modo == 'editar' else 'Nuevo'} Registro")
            nuevos_datos = {}
            cols = st.columns(3)
            for i, col_name in enumerate(df.columns):
                val_previo = df.loc[st.session_state.registro_a_editar, col_name] if st.session_state.modo == "editar" else ""
                nuevos_datos[col_name] = cols[i % 3].text_input(col_name, value=str(val_previo))
            
            if st.form_submit_button("💾 Guardar Datos"):
                fila_nueva = pd.DataFrame([nuevos_datos])
                fila_nueva['FECHA'] = pd.to_datetime(fila_nueva['FECHA'])
                if st.session_state.modo == "editar":
                    st.session_state.df_editada.loc[st.session_state.registro_a_editar] = fila_nueva.iloc[0]
                else:
                    st.session_state.df_editada = pd.concat([df, fila_nueva], ignore_index=True)
                st.session_state.modo = None
                st.rerun()

    if st.session_state.modo == "confirmar_borrado":
        st.warning(f"⚠️ ¿Confirmar eliminación de {len(st.session_state.indices_borrar)} filas?")
        if st.button("✔️ Confirmar"):
            st.session_state.df_editada = df.drop(st.session_state.indices_borrar).reset_index(drop=True)
            st.session_state.modo = None
            st.rerun()

    # 4. REPORTES PROFESIONALES (TÍTULOS Y DESCARGAS)
    st.write("---")
    st.subheader("📊 Paso 3: Reportes de Totales")
    tab1, tab2 = st.tabs(["📄 Exportación de Datos", "📉 Cuadro Estadístico"])

    with tab1:
        st.markdown("### REPORTE DETALLADO DE OPERACIONES - CARGO PESCA S.A.")
        cols_exp = st.multiselect("Seleccionar Columnas:", df.columns.tolist(), default=df.columns.tolist())
        df_exp = df[cols_exp].copy()
        st.dataframe(df_exp, use_container_width=True) # Adaptable a cualquier tamaño
        
        if st.button("🚀 Generar Excel"):
            buf = BytesIO()
            with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
                df_exp.to_excel(writer, index=False, sheet_name='Datos')
            st.download_button("⬇️ Descargar Reporte", buf.getvalue(), "reporte_pesca.xlsx")

    with tab2:
        st.markdown("### CUADRO ESTADÍSTICO DE TOTALES")
        tipo = st.radio("Operación:", ["Contar Registros", "Sumar Cantidades"], horizontal=True)
        cols_st = st.multiselect("Columnas a calcular:", [c for c in df.columns if c != 'FECHA'])
        
        if cols_st:
            res = df.copy()
            if "Sumar" in tipo:
                for c in cols_st:
                    res[c] = pd.to_numeric(res[c].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
            
            # Agrupación y Fila de Total General
            resumen = res.groupby(res['FECHA'].dt.date)[cols_st].agg('sum' if "Sumar" in tipo else 'count').reset_index()
            tot_row = {"FECHA": "TOTAL GENERAL"}
            for c in cols_st: tot_row[c] = resumen[c].sum()
            resumen_final = pd.concat([resumen, pd.DataFrame([tot_row])], ignore_index=True)
            
            st.table(resumen_final) # Vista limpia de totales