import pandas as pd
import streamlit as st
import io

def cargar_datos(archivo):
    try:
        # header=2 indica que los títulos están en la FILA 3 de Excel
        df = pd.read_excel(archivo, header=2)
        
        # Eliminamos filas que estén totalmente vacías
        df = df.dropna(how='all').reset_index(drop=True)
        
        # Limpiamos los nombres de las columnas
        df.columns = [str(col).strip() for col in df.columns]
        
        return df
    except Exception as e:
        st.error(f"Error al procesar el Excel: {e}")
        return None

def procesar_resumen(df, col_fecha, col_chofer, categorias):
    # Limpiar fechas
    df[col_fecha] = pd.to_datetime(df[col_fecha], errors='coerce')
    df = df.dropna(subset=[col_fecha])
    
    # Crear copia para no alterar el original
    resumen_df = df.copy()
    
    # Contar choferes (Columna B usualmente)
    resumen_df['TOTAL CHOFER'] = resumen_df[col_chofer].apply(lambda x: 1 if pd.notnull(x) and str(x).strip() != "" else 0)
    
    # Contar categorías (1 si hay dato, 0 si no)
    for cat in categorias:
        resumen_df[cat] = resumen_df[cat].apply(lambda x: 1 if pd.notnull(x) and str(x).strip() != "" else 0)
    
    # Agrupar por fecha
    columnas_a_sumar = ['TOTAL CHOFER'] + categorias
    resultado = resumen_df.groupby(col_fecha)[columnas_a_sumar].sum().reset_index()
    
    return resultado

def generar_excel_formateado(df_final):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df_final.to_excel(writer, index=False, sheet_name='Reporte_Cargo_Pesca')
        
        workbook = writer.book
        worksheet = writer.sheets['Reporte_Cargo_Pesca']
        
        # Formato para el encabezado verde
        header_fmt = workbook.add_format({
            'bold': True, 'fg_color': '#1E8449', 'font_color': 'white', 'border': 1
        })

        # Ajuste automático de columnas (optimizado)
        for i, col in enumerate(df_final.columns):
            # 1. Obtenemos el largo del título
            largo_titulo = len(str(col))
            
            # 2. Obtenemos el largo del dato más largo en la columna (optimizado)
            largo_max_datos = df_final[col].fillna('').astype(str).str.len().max() if not df_final[col].empty else 0
            
            # 3. Elegimos el mayor y sumamos un pequeño margen
            ancho_final = max(largo_titulo, largo_max_datos) + 2
            
            # 4. Limitamos el ancho (máximo 60 caracteres) para que no se deforme
            ancho_final = min(ancho_final, 60)
            
            # Aplicamos al Excel
            worksheet.set_column(i, i, ancho_final)
            worksheet.write(0, i, col, header_fmt)

    buffer.seek(0)
    return buffer