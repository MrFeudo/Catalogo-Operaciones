import streamlit as st
import pandas as pd

# Configuración de la página ancha para ordenador
st.set_page_config(page_title="Buscador Técnico de Operaciones", layout="wide")

# 1. SISTEMA DE SEGURIDAD (Contraseña)
def check_password():
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False

    if not st.session_state.authenticated:
        st.title("🔐 Acceso Red de Dealers")
        password = st.text_input("Introduce la contraseña de acceso:", type="password")
        if st.button("Entrar"):
            if password == "DealersOJ2026": # Puedes cambiar esta contraseña por la que quieras
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("❌ Contraseña incorrecta")
        return False
    return True

if check_password():
    # 2. CARGA DEL EXCEL (Lee la pestaña exacta de las horas de trabajo)
    @st.cache_data
    def load_data():
        # Aquí le decimos que lea el archivo .xlsx y la pestaña correcta
        df = pd.read_excel("DMS_Active_Spare_Parts.xlsx", sheet_name="new_srv_workhours")
        
        # Filtramos y renombramos las columnas para tus dealers
        columnas_utiles = {
            'new_productmodel_idname': 'Modelo',
            'new_stationname': 'Pieza / Componente',
            'new_code': 'Código Pieza',
            'new_name': 'Operación Técnica',
            'new_standardhour': 'Tiempo Estándar (UT/Horas)',
            'new_remark': 'Notas / Exclusiones'
        }
        df = df[list(columnas_utiles.keys())].rename(columns=columnas_utiles)
        return df

    try:
        data = load_data()
        
        st.title("🚗 Buscador de Operaciones y Tiempos de Mano de Obra")
        st.write("Consulta piezas, modelos y tiempos asignados directamente desde el DMS.")
        st.markdown("---")

        # 3. INTERFAZ DE BÚSQUEDA TRIPLE
        col1, col2, col3 = st.columns(3)
        
        with col1:
            modelos_disponibles = ["Todos"] + list(data['Modelo'].dropna().unique())
            modelo_seleccionado = st.selectbox("1. Filtrar por Modelo:", modelos_disponibles)
            
        with col2:
            buscar_pieza = st.text_input("2. Buscar por nombre o código de pieza:", "").strip()
            
        with col3:
            buscar_operacion = st.text_input("3. Buscar por tipo de operación (ej: Remove):", "").strip()

        # 4. LÓGICA DEL FILTRADO
        df_filtrado = data.copy()
        
        if modelo_seleccionado != "Todos":
            df_filtrado = df_filtrado[df_filtrado['Modelo'] == modelo_seleccionado]
            
        if buscar_pieza:
            df_filtrado = df_filtrado[
                df_filtrado['Pieza / Componente'].str.contains(buscar_pieza, case=False, na=False) |
                df_filtrado['Código Pieza'].astype(str).str.contains(buscar_pieza, case=False, na=False)
            ]
            
        if buscar_operacion:
            df_filtrado = df_filtrado[df_filtrado['Operación Técnica'].str.contains(buscar_operacion, case=False, na=False)]

        # 5. MOSTRAR RESULTADOS
        st.markdown(f"### 📋 Resultados encontrados: {len(df_filtrado)} operaciones")
        
        if not df_filtrado.empty:
            st.dataframe(df_filtrado, use_container_width=True, hide_index=True)
        else:
            st.warning("⚠️ No se encontraron operaciones con los criterios seleccionados.")
            
    except Exception as e:
        st.error(f"Error al cargar la base de datos: {e}")
