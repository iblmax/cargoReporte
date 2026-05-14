import streamlit as st
import pandas as pd
import estilos 
import procesador 
from io import BytesIO

# Configuración de la página
st.set_page_config(page_title="Sistema CARGO PESCA", layout="wide")

# 1. Estética e Inicialización
estilos.aplicar_estilos()
estilos.mostrar_cabecera()

if "df_editada" not in st.session_state:
    st.session_state.df_editada = None
if "columnas_confirmadas" not in st.session_state:
    st.session_state.columnas_confirmadas = False
if "modo" not in st.session_state:
    st.session_state.modo = None
if "indices_temp" not in st.session_state:
    st.session_state.indices_temp = []

st.write("---")

# 2. Carga de Archivo y Configuración de Títulos
archivo = st.file_uploader("Seleccione el archivo de Excel de CARGO PESCA", type=["xlsx"])

if archivo:
    if not st.session_state.columnas_confirmadas:
        st.subheader("🛠️ Paso 1: Configurar Estructura")
        df_ref = pd.read_excel(archivo, header=None, nrows=15)
        fila_idx = st.number_input("Fila de títulos (donde están los encabezados):", 0, 14, 0)
        
        if st.button("✅ Confirmar Estructura", type="primary"):
            df_cargado = pd.read_excel(archivo, skiprows=fila_idx)
            df_cargado.columns = [str(c).replace('.', '').strip().upper() for c in df_cargado.columns]
            
            c_fecha = next((c for c in df_cargado.columns if 'FECHA' in c), None)
            if c_fecha:
                df_cargado.rename(columns={c_fecha: 'FECHA'}, inplace=True)
                df_cargado['FECHA'] = pd.to_datetime(df_cargado['FECHA'], errors='coerce')
            
            st.session_state.df_editada = df_cargado.reset_index(drop=True)
            st.session_state.columnas_confirmadas = True
            st.rerun()
        
        st.dataframe(df_ref.style.apply(lambda x: ['background-color: #d1e7dd' if x.name == fila_idx else '' for i in x], axis=1))

# 3. Interfaz de Gestión y Reportes (Solo tras confirmar títulos)
if st.session_state.columnas_confirmadas and st.session_state.df_editada is not None:
    df = st.session_state.df_editada

    with st.sidebar:
        if st.button("📁 Cargar otro archivo"):
            st.session_state.df_editada = None
            st.session_state.columnas_confirmadas = False
            st.rerun()

    # --- SECCIÓN: GESTIÓN Y EDICIÓN ---
    st.subheader("🔍 Paso 2: Gestión y Edición de Registros")
    
    col_v1, col_v2 = st.columns([2, 1])
    with col_v1:
        columnas_visibles = st.multiselect(
            "Columnas visibles en tabla:", options=df.columns.tolist(), default=df.columns.tolist()[:7]
        )
    with col_v2:
        busqueda = st.text_input("🔎 Filtrar por cualquier dato:")

    df_vista = df.copy()
    if busqueda:
        mask = df_vista.apply(lambda x: x.astype(str).str.contains(busqueda, case=False, na=False).any(), axis=1)
        df_vista = df_vista[mask]

    # TABLA MAESTRA CON SELECTOR
    df_con_check = df_vista[columnas_visibles].copy()
    if 'FECHA' in df_con_check.columns:
        df_con_check['FECHA'] = df_con_check['FECHA'].dt.strftime('%d/%m/%Y')
    
    df_con_check.insert(0, "SELECCIONAR", False)
    
    editor_res = st.data_editor(
        df_con_check,
        hide_index=False,
        use_container_width=True,
        key="tabla_maestra"
    )

    indices_marcados = editor_res[editor_res["SELECCIONAR"] == True].index.tolist()

    # BOTONES DE ACCIÓN
    c1, c2, c3 = st.columns(3)
    if c1.button("🔧 Editar Marcados", disabled=not indices_marcados):
        st.session_state.modo = "editar"
        st.session_state.indices_temp = indices_marcados
    
    if c2.button("🗑️ Eliminar Marcados", type="primary", disabled=not indices_marcados):
        st.session_state.df_editada = st.session_state.df_editada.drop(indices_marcados).reset_index(drop=True)
        st.success("Registros eliminados correctamente")
        st.rerun()
        
    if c3.button("➕ Agregar Nuevo Registro"):
        st.session_state.modo = "agregar"

    # FORMULARIOS DE GESTIÓN
    if st.session_state.modo == "editar":
        with st.form("form_edicion"):
            st.write(f"### 🔧 Editando {len(st.session_state.indices_temp)} filas")
            for idx in st.session_state.indices_temp:
                st.markdown(f"**Fila ID: {idx}**")
                cols_form = st.columns(len(columnas_visibles))
                for i, col_name in enumerate(columnas_visibles):
                    val_actual = df.at[idx, col_name]
                    # Formatear fecha para el input si es necesario
                    if col_name == 'FECHA' and pd.notnull(val_actual):
                        val_actual = val_actual.strftime('%Y-%m-%d')
                    
                    nuevo_val = cols_form[i].text_input(f"{col_name}", value=str(val_actual), key=f"ed_{idx}_{col_name}")
                    
                    # Guardar el cambio (con conversión a fecha si es la columna FECHA)
                    if col_name == 'FECHA':
                        st.session_state.df_editada.at[idx, col_name] = pd.to_datetime(nuevo_val, errors='coerce')
                    else:
                        st.session_state.df_editada.at[idx, col_name] = nuevo_val
            
            if st.form_submit_button("💾 Guardar Cambios"):
                st.session_state.modo = None
                st.success("Cambios aplicados")
                st.rerun()

    if st.session_state.modo == "agregar":
        with st.form("form_agregar"):
            st.write("### ➕ Nuevo Registro")
            nueva_fila = {}
            cols_add = st.columns(3)
            for i, col_name in enumerate(df.columns):
                nueva_fila[col_name] = cols_add[i % 3].text_input(col_name)
            
            if st.form_submit_button("✅ Registrar"):
                # Convertir la fecha del input manual
                if 'FECHA' in nueva_fila:
                    nueva_fila['FECHA'] = pd.to_datetime(nueva_fila['FECHA'], errors='coerce')
                
                df_nuevo = pd.DataFrame([nueva_fila])
                st.session_state.df_editada = pd.concat([st.session_state.df_editada, df_nuevo], ignore_index=True)
                st.session_state.modo = None
                st.rerun()

    # --- SECCIÓN: REPORTES ---
    st.write("---")
    st.subheader("📊 Paso 3: Generar Reportes")
    
    t_data, t_stats = st.tabs(["📄 Reporte de Datos", "📉 Reporte Estadístico (Totales)"])

    with t_data:
        m_reporte = st.radio("Alcance del reporte:", ["Completo", "Filtrado por columna"], horizontal=True)
        df_final_data = df.copy()
        
        if m_reporte == "Filtrado por columna":
            c_filtro = st.selectbox("Columna:", [c for c in df.columns if c != 'FECHA'])
            v_filtro = st.selectbox("Valor único:", sorted(df[c_filtro].dropna().unique().astype(str)))
            df_final_data = df[df[c_filtro].astype(str) == v_filtro].copy()
        
        if 'FECHA' in df_final_data.columns:
            df_final_data['FECHA'] = df_final_data['FECHA'].dt.strftime('%d/%m/%Y')

        st.dataframe(df_final_data, use_container_width=True)

        if st.button("🚀 Exportar Excel de Datos"):
            buf = procesador.generar_excel_formateado(df_final_data)
            st.download_button("⬇️ Descargar Archivo", buf, "reporte_datos.xlsx")

    with t_stats:
        st.write("Resumen de totales por fecha.")
        if 'FECHA' in df.columns:
            col_f1, col_f2 = st.columns(2)
            f_ini = col_f1.date_input("Desde", df['FECHA'].min().date())
            f_fin = col_f2.date_input("Hasta", df['FECHA'].max().date())
            
            df_r = df[(df['FECHA'].dt.date >= f_ini) & (df['FECHA'].dt.date <= f_fin)].copy()
            
            if not df_r.empty:
                df_r['FECHA_STR'] = df_r['FECHA'].dt.strftime('%d/%m/%Y')
                agg_map = {c: ('sum' if pd.api.types.is_numeric_dtype(df_r[c]) else 'count') 
                          for c in df_r.columns if c not in ['FECHA', 'FECHA_STR']}
                
                df_resumen = df_r.groupby('FECHA_STR').agg(agg_map).reset_index()
                df_resumen.rename(columns={'FECHA_STR': 'FECHA'}, inplace=True)
                
                f_total = {"FECHA": "TOTAL GENERAL"}
                for c in agg_map: f_total[c] = df_resumen[c].sum()
                df_completo = pd.concat([df_resumen, pd.DataFrame([f_total])], ignore_index=True)
                
                st.table(df_completo)

                if st.button("🚀 Exportar Cuadro Estadístico"):
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        df_completo.to_excel(writer, index=False, sheet_name='Totales')
                        workbook = writer.book
                        worksheet = writer.sheets['Totales']
                        formato_celda = workbook.add_format({'border': 1, 'align': 'center'})
                        for idx, col in enumerate(df_completo.columns):
                            worksheet.set_column(idx, idx, 15, formato_celda)
                    st.download_button("⬇️ Descargar Cuadro", output.getvalue(), "totales_fechas.xlsx")