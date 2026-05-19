import streamlit as st
import pandas as pd

st.set_page_config(page_title="Buscador Técnico de Operaciones", layout="wide")

# 1. SEGURIDAD
def check_password():
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False
    if not st.session_state.authenticated:
        st.title("🔐 Acceso Red de Dealers")
        password = st.text_input("Introduce la contraseña de acceso:", type="password")
        if st.button("Entrar"):
            if password == "DealersOJ2026":
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("❌ Contraseña incorrecta")
        return False
    return True

if check_password():
# 2. CARGA DEL EXCEL
    @st.cache_data
    def load_data():
        df = pd.read_excel("DMS_Active_Spare_Parts.xlsx", sheet_name="new_srv_workhours")
        
        # EXCEPCIÓN: Filtramos para dejar las que tienen nombre de pieza O las que son operaciones universales (999999)
        # Nota: Convertimos new_stationname a string para que compare bien aunque el Excel lo lea como número o texto
        condicion_nombre = df['new_product_idname'].dropna().astype(str).str.strip() != ""
        condicion_universal = df['new_stationname'].astype(str).str.strip() == "999999"
        
        # Aplicamos el filtro combinando ambas condiciones con el símbolo "|" (que significa "Ó")
        df = df[condicion_nombre | condicion_universal]
        
        # MAPEO CON TU NUEVA COLUMNA DE NOMBRES INCLUIDA
        df = df.rename(columns={
            'new_productmodel_idname': 'Modelo',
            'new_product_idname': 'Nombre de la Pieza',
            'new_code': 'Código de Referencia',
            'new_name': 'Operación Técnica',
            'new_standardhour': 'Tiempo Estándar (UT/Horas)',
            'new_remark': 'Notas / Exclusiones'
        })
        
        # Definimos el orden en el que se verán las columnas en la web
        columnas_finales = [
            'Modelo', 
            'Nombre de la Pieza', 
            'Código de Referencia', 
            'Operación Técnica', 
            'Tiempo Estándar (UT/Horas)', 
            'Notas / Exclusiones'
        ]
        
        # Por si acaso alguna fila no tiene datos, rellenamos los vacíos con texto limpio
        df = df.fillna("")
        
        # Filtramos para mostrar solo las columnas que existan realmente en el Excel
        columnas_presentes = [col for col in columnas_finales if col in df.columns]
        return df[columnas_presentes]

    try:
        data = load_data()
        
        st.title("🚗 Buscador de Operaciones y Tiempos de Mano de Obra")
        st.write("Consulta piezas, modelos y tiempos asignados directamente desde el DMS.")
        st.markdown("---")

        # 3. INTERFAZ DE BÚSQUEDA TRIPLE (Visualmente más cómodo)
        col1, col2, col3 = st.columns([1, 1.5, 1.5])
        
        with col1:
            modelos_disponibles = ["Todos"] + list(data['Modelo'].dropna().unique())
            modelo_seleccionado = st.selectbox("1. Filtrar por Modelo:", modelos_disponibles)
            
        with col2:
            # Buscador por texto en el Nombre o Código de la pieza
            buscar_pieza = st.text_input("2. Buscar por Nombre o Código de pieza:", "").strip()
            
        with col3:
            # Buscador por texto en la Operación Técnica
            buscar_operacion = st.text_input("3. Buscar por tipo de operación (ej: Remove, Paint...):", "").strip()

        # 4. LÓGICA DE FILTRADO
        df_filtrado = data.copy()
        
        if modelo_seleccionado != "Todos":
            df_filtrado = df_filtrado[df_filtrado['Modelo'] == modelo_seleccionado]
            
        if buscar_pieza:
            # Busca a la vez tanto en la columna de nombres comerciales como en la de códigos
            df_filtrado = df_filtrado[
                df_filtrado['Nombre de la Pieza'].astype(str).str.contains(buscar_pieza, case=False, na=False) |
                df_filtrado['Código de Referencia'].astype(str).str.contains(buscar_pieza, case=False, na=False)
            ]
            
        if buscar_operacion:
            df_filtrado = df_filtrado[df_filtrado['Operación Técnica'].astype(str).str.contains(buscar_operacion, case=False, na=False)]

        # 5. RESULTADOS
        st.markdown(f"### 📋 Resultados encontrados: {len(df_filtrado)} operaciones")
        
        if not df_filtrado.empty:
            st.dataframe(df_filtrado, use_container_width=True, hide_index=True)
        else:
            st.warning("⚠️ No se encontraron operaciones con los criterios seleccionados.")
            
    except Exception as e:
        st.error(f"Error al procesar la base de datos: {e}")
