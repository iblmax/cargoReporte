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

# Inicialización de estados para formularios
if "df_editada" not in st.session_state:
    st.session_state.df_editada = None
if "columnas_confirmadas" not in st.session_state:
    st.session_state.columnas_confirmadas = False
if "modo" not in st.session_state:
    st.session_state.modo = None
if "registro_a_editar" not in st.session_state:
    st.session_state.registro_a_editar = None

st.write("---")

# 2. Carga de Archivo
archivo = st.file_uploader("Subir archivo de CARGO PESCA", type=["xlsx", "xls"])

if archivo and not st.session_state.columnas_confirmadas:
    # Lógica de configuración inicial (se mantiene igual)
    st.session_state.df_editada = pd.read_excel(archivo) # Simplificado para el ejemplo
    st.session_state.columnas_confirmadas = True
    st.rerun()

# 3. Interfaz Principal
if st.session_state.columnas_confirmadas:
    df = st.session_state.df_editada

    # --- SECCIÓN GESTIÓN ---
    st.subheader("🔍 Gestión de Datos")
    f_dia = st.date_input("📅 Filtrar por Día:", value=datetime.now())
    
    # Tabla de Selección
    df_temp = df.copy()
    df_temp.insert(0, "SELECCIONAR", False)
    editor = st.data_editor(df_temp, use_container_width=True, key="gestor_tabla")
    
    indices_sel = editor[editor["SELECCIONAR"] == True].index.tolist()

    # Botones de Acción
    c1, c2, c3 = st.columns(3)
    if c1.button("🔧 Editar Selección", disabled=not indices_sel):
        st.session_state.modo = "editar"
        st.session_state.registro_a_editar = indices_sel[0]
    
    if c2.button("🗑️ Eliminar Selección", type="primary", disabled=not indices_sel):
        st.session_state.modo = "confirmar_borrado"
        st.session_state.indices_borrar = indices_sel

    if c3.button("➕ Nuevo Registro"):
        st.session_state.modo = "nuevo"

    # --- FORMULARIOS DINÁMICOS (EDITAR / NUEVO) ---
    if st.session_state.modo in ["editar", "nuevo"]:
        with st.expander(f"📝 {'Editar Registro' if st.session_state.modo == 'editar' else 'Nuevo Registro'}", expanded=True):
            with st.form("form_registro"):
                valores = {}
                cols_form = st.columns(3)
                for i, col_name in enumerate(df.columns):
                    default_val = df.loc[st.session_state.registro_a_editar, col_name] if st.session_state.modo == "editar" else ""
                    valores[col_name] = cols_form[i % 3].text_input(col_name, value=str(default_val))
                
                if st.form_submit_button("💾 Guardar Cambios"):
                    nuevo_df = pd.DataFrame([valores])
                    if st.session_state.modo == "editar":
                        st.session_state.df_editada.loc[st.session_state.registro_a_editar] = nuevo_df.iloc[0]
                    else:
                        st.session_state.df_editada = pd.concat([df, nuevo_df], ignore_index=True)
                    st.session_state.modo = None
                    st.rerun()

    # --- CONFIRMACIÓN DE ELIMINACIÓN ---
    if st.session_state.modo == "confirmar_borrado":
        st.warning(f"⚠️ ¿Está seguro que desea eliminar {len(st.session_state.indices_borrar)} registros?")
        if st.button("✔️ Sí, Confirmar Eliminación"):
            st.session_state.df_editada = df.drop(st.session_state.indices_borrar).reset_index(drop=True)
            st.session_state.modo = None
            st.rerun()
        if st.button("❌ Cancelar"):
            st.session_state.modo = None
            st.rerun()

    # --- SECCIÓN REPORTES MEJORADOS ---
    st.write("---")
    st.subheader("📊 Reportes Profesionales")
    t1, t2 = st.tabs(["📄 Exportación Detallada", "📉 Estadística de Totales"])

    with t1:
        st.markdown("### REPORTE DE GESTIÓN LOGÍSTICA - CARGO PESCA S.A.")
        cols_rep = st.multiselect("Columnas a incluir:", df.columns.tolist(), default=df.columns.tolist())
        df_rep = df[cols_rep]
        st.dataframe(df_rep, use_container_width=True) # Se adapta a cualquier tamaño
        
        if st.button("🚀 Exportar Excel Detallado"):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_rep.to_excel(writer, index=False, sheet_name='Reporte')
                # Título en Excel
                writer.sheets['Reporte'].write('A1', 'REPORTE DETALLADO DE OPERACIONES')
            st.download_button("⬇️ Descargar Excel", output.getvalue(), "Reporte_Detallado.xlsx")

    with t2:
        st.markdown("### CUADRO ESTADÍSTICO DE OPERACIONES")
        # Lógica de totales dinámica
        tipo = st.radio("Cálculo:", ["Contar Registros", "Sumar Cantidades"], horizontal=True)
        col_sumar = st.multiselect("Columnas a calcular:", [c for c in df.columns if c != 'FECHA'])
        
        if col_sumar:
            # Limpieza y suma dinámica
            resumen = df.copy()
            for c in col_sumar:
                resumen[c] = pd.to_numeric(resumen[c].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
            
            final = resumen.groupby(resumen['FECHA'].dt.date)[col_sumar].agg('sum' if "Sumar" in tipo else 'count')
            st.table(final) # Tabla limpia para totales