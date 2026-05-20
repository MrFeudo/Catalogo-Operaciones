import streamlit as st
import pandas as pd

st.set_page_config(page_title="Buscador Técnico OMODA & JAECOO", layout="wide")

# ==========================================
# 1. BARRA LATERAL: LOGO + MENÚ DE NAVEGACIÓN
# ==========================================
try:
    st.sidebar.image("logo_empresa.png", use_container_width=True)
except Exception:
    st.sidebar.write("🏢 **OMODA & JAECOO**")

st.sidebar.markdown("---")

st.sidebar.markdown("### 🗺️ Menú de Navegación")
opcion_menu = st.sidebar.radio(
    "Selecciona una herramienta:",
    ["📋 Tiempos de Taller", "💰 Precios de Recambios"]
)

# 2. SEGURIDAD (Global para toda la app)
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
    
    # =========================================================================
    # PANTALLA 1: TIEMPOS DE TALLER (Tu versión de seguridad exacta)
    # =========================================================================
    if opcion_menu == "📋 Tiempos de Taller":
        
        @st.cache_data(ttl=600)
        def load_data():
            df = pd.read_excel("DMS_Active_Spare_Parts.xlsx", sheet_name="new_srv_workhours")
            
            df['new_product_idname'] = df['new_product_idname'].astype(str).str.strip()
            df['new_stationname'] = df['new_stationname'].astype(str).str.strip()
            
            tiene_nombre = (df['new_product_idname'] != "") & (df['new_product_idname'] != "nan")
            es_universal = df['new_stationname'] == "999999"
            df = df[tiene_nombre | es_universal].copy()
            
            df = df.rename(columns={
                'new_productmodel_idname': 'Modelo',
                'new_product_idname': 'Nombre de la Pieza',
                'new_code': 'Código de Referencia',
                'new_name': 'Operación Técnica',
                'new_standardhour': 'Tiempo Estándar (UT/Horas)',
                'new_remark': 'Notas / Exclusiones',
                'Organization': 'Organización',
                'statuscodename': 'Estado'
            })
            
            columnas_finales = [
                'Modelo', 'Nombre de la Pieza', 'Código de Referencia', 
                'Operación Técnica', 'Tiempo Estándar (UT/Horas)', 'Notas / Exclusiones',
                'Organización', 'Estado'
            ]
            
            df = df.fillna("")
            df = df.replace("nan", "")
            
            columnas_presentes = [col for col in columnas_finales if col in df.columns]
            return df[columnas_presentes].reset_index(drop=True)

        try:
            data = load_data()
            
            col_logo, col_titulo = st.columns([1.5, 5])
            with col_logo:
                try: st.image("logo_empresa.png", width=220)
                except Exception: pass
                    
            with col_titulo:
                st.title("Catálogo Operaciones de mano de obra")
                st.write("Consulta piezas, modelos y tiempos activos en DMS.")
                
            st.markdown("---")

            # Interfaz de Filtros Fila 1
            col1, col2, col3 = st.columns([1, 1.5, 1.5])
            with col1:
                modelos_disponibles = ["Todos"] + list(data['Modelo'].dropna().unique())
                modelo_seleccionado = st.selectbox("1. Filtrar por Modelo:", modelos_disponibles)
            with col2:
                buscar_pieza = st.text_input("2. Buscar por Nombre o Código de pieza:", "").strip()
            with col3:
                buscar_operacion = st.text_input("3. Buscar por tipo de operación (ej: Remove, Paint...):", "").strip()

            # Interfaz de Filtros Fila 2
            col_org, col_est = st.columns(2)
            with col_org:
                org_disponibles = ["Todas"] + list(data['Organización'].dropna().unique())
                org_seleccionada = st.selectbox("4. Filtrar por Organización:", org_disponibles)
            with col_est:
                est_disponibles = ["Todos"] + list(data['Estado'].dropna().unique())
                est_seleccionado = st.selectbox("5. Filtrar por Estado (Activo/Inactivo):", est_disponibles)

            # Lógica de Filtrado
            df_filtrado = data.copy()
            if modelo_seleccionado != "Todos":
                df_filtrado = df_filtrado[df_filtrado['Modelo'] == modelo_seleccionado]
            if org_seleccionada != "Todas":
                df_filtrado = df_filtrado[df_filtrado['Organización'] == org_seleccionada]
            if est_seleccionado != "Todos":
                df_filtrado = df_filtrado[df_filtrado['Estado'] == est_seleccionado]

            if buscar_pieza:
                df_filtrado = df_filtrado[
                    df_filtrado['Nombre de la Pieza'].astype(str).str.contains(buscar_pieza, case=False, na=False) |
                    df_filtrado['Código de Referencia'].astype(str).str.contains(buscar_pieza, case=False, na=False)
                ]
            if buscar_operacion:
                df_filtrado = df_filtrado[df_filtrado['Operación Técnica'].astype(str).str.contains(buscar_operacion, case=False, na=False)]

            st.markdown(f"### 📋 Resultados encontrados: {len(df_filtrado)} operaciones")
            if not df_filtrado.empty:
                st.dataframe(df_filtrado, use_container_width=True, hide_index=True)
            else:
                st.warning("⚠️ No se encontraron operaciones con los criterios seleccionados.")
                
        except Exception as e:
            st.error(f"Error al procesar la base de datos de tiempos: {e}")

    # =========================================================================
    # PANTALLA 2: PRECIOS DE RECAMBIOS (Filtrado Posicional Absoluto Ineludible)
    # =========================================================================
    elif opcion_menu == "💰 Precios de Recambios":
        
        @st.cache_data(ttl=600)
        def load_data_prices():
            # La caché solo lee el almacenamiento físico para ir rápido
            return pd.read_excel("DMS_Active_Spare_Parts.xlsx", sheet_name="Parts price")

        try:
            # 1. Recuperamos los datos brutos del Excel
            df_original = load_data_prices().copy()
            
            # 2. ESCANEO Y FILTRADO HORIZONTAL (Ignora el nombre técnico de las cabeceras):
            # Busca la cadena "Spain OJ" celda por celda en todo el documento. Si está, se queda la fila.
            mascara_spain = df_original.astype(str).apply(lambda x: x.str.strip().str.contains("Spain OJ", case=False, na=False)).any(axis=1)
            df_solo_spain = df_original[mascara_spain].copy()
            
            # 3. Renombramos las columnas sobre la estructura de España limpia
            df_solo_spain = df_solo_spain.rename(columns={
                'new_partscode': 'Código de Recambio',
                'new_product_idname': 'Descripción de la Pieza',
                'new_price': 'Precio Venta',
                'transactioncurrencyidname': 'Moneda',
                'new_pricetypename': 'Tipo de Tarifa',
                'new_businessunit_idname': 'Mercado / Organización',
                'statecodename': 'Estado'
            })
            
            # 4. Establecemos el formato definitivo de salida
            columnas_finales_precios = [
                'Código de Recambio', 'Descripción de la Pieza', 
                'Precio Venta', 'Moneda', 'Tipo de Tarifa', 
                'Mercado / Organización', 'Estado'
            ]
            columnas_visibles = [col for col in columnas_finales_precios if col in df_solo_spain.columns]
            df_preparado = df_solo_spain[columnas_visibles].reset_index(drop=True)
            
            # Rellenamos celdas vacías para el renderizado
            df_preparado = df_preparado.fillna("")
            df_preparado = df_preparado.replace("nan", "")
            
            # --- INTERFAZ GRÁFICA DE PRECIOS ---
            col_logo, col_titulo = st.columns([1.5, 5])
            with col_logo:
                try: st.image("logo_empresa.png", width=220)
                except Exception: pass
                    
            with col_titulo:
                st.title("Maestro de Tarifas y Precios de Recambios")
                st.write("Consulta oficializada de precios y tarifas de la red Spain OJ.")
                
            st.markdown("---")
            
            # Buscador en texto superior
            buscar_recambio = st.text_input("🔍 Introduce el Código de recambio o la Descripción de la pieza:", "").strip()
            
            # 5. Aplicamos la caja de búsqueda sobre el set exclusivo de España
            if buscar_recambio:
                df_tabla_final = df_preparado[
                    df_preparado['Código de Recambio'].astype(str
