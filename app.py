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
    @st.cache_data(ttl=600)
    def load_data():
        df = pd.read_excel("DMS_Active_Spare_Parts.xlsx", sheet_name="new_srv_workhours")
        
        # 1. Aseguramos que no haya fallos de formato pasando las columnas a texto limpio
        df['new_product_idname'] = df['new_product_idname'].astype(str).str.strip()
        df['new_stationname'] = df['new_stationname'].astype(str).str.strip()
        
        # 2. Definimos las condiciones exactas:
        # Tiene nombre real (ni vacío, ni la palabra "nan")
        tiene_nombre = (df['new_product_idname'] != "") & (df['new_product_idname'] != "nan")
        # O es la excepción universal exacta
        es_universal = df['new_stationname'] == "999999"
        
        # 3. Filtramos: Solo entra si tiene nombre válido O si es la estación universal
        df = df[tiene_nombre | es_universal].copy()
        
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
        
        # Rellenamos los vacíos reales con texto vacío para que no aparezcan "nan" visualmente
        df = df.fillna("")
        df = df.replace("nan", "")
        
        # Filtramos para mostrar solo las columnas que existan realmente en el Excel
        columnas_presentes = [col for col in columnas_finales if col in df.columns]
        
        # Reseteamos el índice para que la tabla quede perfecta
        return df[columnas_presentes].reset_index(drop=True)

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

       # 4. LÓGICA DE FILTRADO (Corregida y limpia)
        df_filtrado = data.copy()
        
        if modelo_seleccionado != "Todos":
            df_filtrado = df_filtrado[df_filtrado['Modelo'] == modelo_seleccionado]
            
        if org_seleccionada != "Todas":
            df_filtrado = df_filtrado[df_filtrado['Organización'] == org_seleccionada]
            
        if est_seleccionado != "Todos":
            df_filtrado = df_filtrado[df_filtrado['Estado'] == est_seleccionado] # <-- Corregido aquí

        if buscar_pieza:
            df_filtrado = df_filtrado[
                df_filtrado['Nombre de la Pieza'].astype(str).str.contains(buscar_pieza, case=False, na=False) |
                df_filtrado['Código de Referencia'].astype(str).str.contains(buscar_pieza, case=False, na=False)
            ]
            
        if buscar_operacion:
            df_filtrado = df_filtrado[df_filtrado['Operación Técnica'].astype(str).str.contains(buscar_operacion, case=False, na=False)]
            
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
