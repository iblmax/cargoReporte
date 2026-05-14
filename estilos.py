import streamlit as st

def aplicar_estilos():
    st.set_page_config(
        page_title="CARGO PESCA S.A. - Sistema de Reportes",
        page_icon="🚢",
        layout="wide"
    )

    st.markdown("""
        <style>
        .main {
            background-color: #f4f7f6;
        }
        .stTitle {
            color: #d63384; /* Color magenta del logo */
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            text-align: center;
            font-weight: bold;
        }
        .upload-card {
            background-color: #ffffff;
            padding: 20px;
            border-radius: 10px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.05);
            border-top: 5px solid #007bbd; /* Azul del logo */
        }
        </style>
    """, unsafe_allow_html=True)

def mostrar_cabecera():
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        # Intenta cargar el logo si existe, si no, muestra el título
        try:
            st.image("logo_empresa.png", width=200) 
        except:
            pass
        st.markdown("<h1 class='stTitle'>CARGO PESCA S.A.</h1>", unsafe_allow_html=True)
        st.markdown("<p style='text-align: center;'>Gestión Inteligente de Logística y Transporte</p>", unsafe_allow_html=True) 