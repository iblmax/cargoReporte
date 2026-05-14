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

estilos.aplicar_estilos()
estilos.mostrar_cabecera()

# --- SOLUCIÓN AL TYPEERROR: MANEJO DE NAN Y NONE ---
def descargar_excel_estilizado(df, titulo_reporte):
    output = BytesIO()
    # Llenar NaNs con vacío para evitar el TypeError en xlsxwriter
    df_clean = df.fillna("") 
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_clean.to_excel(writer, index=False, sheet_name='Reporte', startrow=3)
        workbook = writer.book
        worksheet = writer.sheets['Reporte']

        # Formatos
        fmt_titulo = workbook.add_format({'bold': True, 'font_size': 16, 'font_color': '#1f4e78', 'align': 'center'})
        fmt_header = workbook.add_format({'bold': True, 'bg_color': '#2f75b5', 'font_color': 'white', 'border': 1, 'align': 'center'})
        fmt_celda = workbook.add_format({'border': 1, 'align': 'center'})

        # Escribir Título arriba de la tabla
        worksheet.merge_range(1, 0, 1, len(df.columns)-1, titulo_reporte.upper(), fmt_titulo)

        # Aplicar formato a cabeceras
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(3, col_num, value, fmt_header)
        
        # Escribir datos con formato de celda (Bordes)
        for row in range(len(df_clean)):
            for col in range(len(df_clean.columns)):
                val = df_clean.iloc[row, col]
                # Escribir valor asegurando que no sea None
                worksheet.write(row + 4, col, val, fmt_celda)
                
    return output.getvalue()

# 1. CARGA Y CONFIGURACIÓN (RESTAURADA VISTA PREVIA Y MARCADO VERDE)
archivo = st.file_uploader("Subir archivo de CARGO PESCA", type=["xlsx", "xls"])

if archivo and not st.session_state.columnas_confirmadas:
    st.subheader("🛠️ Paso 1: Configurar Estructura de Datos")
    df_ref = pd.read_excel(archivo, header=None, nrows=15)
    
    col_c1, col_c2 = st.columns([1, 2])
    with col_c1:
        orientacion = st.radio("📑 Títulos en:", ["Fila (Horizontal)", "Columna (Vertical)"])
        idx_titulo = st.number_input("Número de Índice:", 0, 14, 0)
        
        if st.button("✅ Confirmar Estructura", type="primary"):
            if orientacion == "Fila (Horizontal)":
                df_final = pd.read_excel(archivo, skiprows=idx_titulo)
            else:
                df_raw = pd.read_excel(archivo, header=None)
                df_final = df_raw.iloc[:, idx_titulo:].T
                df_final.columns = df_final.iloc[0]
                df_final = df_final.drop(df_final.index[0]).reset_index(drop=True)

            # Limpieza y Fecha sin Hora
            df_final.columns = [str(c).strip().upper() for c in df_final.columns]
            if 'FECHA' in df_final.columns:
                df_final['FECHA'] = pd.to_datetime(df_final['FECHA'], errors='coerce')
            
            st.session_state.df_editada = df_final
            st.session_state.columnas_confirmadas = True
            st.rerun()

    with col_c2:
        # LÓGICA DE MARCADO EN VERDE
        def highlight_selection(x):
            color = 'background-color: #d1e7dd' # Verde claro
            df_style = pd.DataFrame('', index=x.index, columns=x.columns)
            if orientacion == "Fila (Horizontal)":
                if idx_titulo in x.index: df_style.loc[idx_titulo, :] = color
            else:
                if idx_titulo in x.columns: df_style.loc[:, idx_titulo] = color
            return df_style
        
        st.write("Vista previa (La marca verde indica los títulos):")
        st.dataframe(df_ref.style.apply(highlight_selection, axis=None), use_container_width=True)

# 2. GESTIÓN Y EDICIÓN
if st.session_state.columnas_confirmadas:
    df = st.session_state.df_editada
    st.subheader("🔍 Gestión y Edición")

    # Selección de columnas para la tabla
    cols_visibles = st.multiselect("Columnas visibles:", df.columns.tolist(), default=df.columns.tolist()[:6])
    
    # Formatear fecha DD/MM/AAAA para la vista de tabla
    df_display = df.copy()
    if 'FECHA' in df_display.columns:
        df_display['FECHA'] = df_display['FECHA'].dt.strftime('%d/%m/%Y')

    df_display.insert(0, "SELECCIONAR", False)
    editor = st.data_editor(df_display[["SELECCIONAR"] + cols_visibles], hide_index=True, use_container_width=True)
    
    # --- REPORTES CON DESCARGA ESTILIZADA ---
    st.write("---")
    t1, t2 = st.tabs(["📄 Exportación", "📉 Estadísticas"])

    with t1:
        st.markdown("### REPORTE DETALLADO")
        st.dataframe(df_display[cols_visibles], use_container_width=True)
        
        # Título arriba de la tabla en el Excel generado
        excel_data = descargar_excel_estilizado(df_display[cols_visibles], "Reporte de Operaciones Cargo Pesca")
        st.download_button("⬇️ Descargar Reporte Estilizado", excel_data, "Reporte_Profesional.xlsx")

    with t2:
        st.markdown("### CUADRO ESTADÍSTICO")
        # Botón de descarga para el cuadro de totales
        col_sum = st.multiselect("Columnas a Sumar:", [c for c in df.columns if c != 'FECHA'])
        if col_sum:
            resumen = df.groupby(df['FECHA'].dt.date)[col_sum].sum().reset_index()
            resumen['FECHA'] = pd.to_datetime(resumen['FECHA']).dt.strftime('%d/%m/%Y')
            st.table(resumen)
            
            stat_data = descargar_excel_estilizado(resumen, "Cuadro Estadístico de Totales")
            st.download_button("⬇️ Descargar Estadísticas", stat_data, "Totales_Profesional.xlsx")