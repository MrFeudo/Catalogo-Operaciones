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

# Menú interactivo lateral
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
    # PANTALLA 2: SECCIÓN DE PRECIOS (Optimizada para España OJ)
    # =========================================================================
    elif opcion_menu == "💰 Precios de Recambios":
        
        @st.cache_data(ttl=600)
        def load_data_prices():
            df = pd.read_excel("DMS_Active_Spare_Parts.xlsx", sheet_name="Parts price")
            
            # 1. FILTRO SEGURO: Solo cargamos el mercado de España OJ
            df['new_businessunit_idname'] = df['new_businessunit_idname'].astype(str).str.strip()
            df = df[df['new_businessunit_idname'] == "Spain OJ"].copy()
            
            # 2. Traducimos y mapeamos las columnas de valor real de tu hoja
            df = df.rename(columns={
                'new_partscode': 'Código de Recambio',
                'new_product_idname': 'Descripción de la Pieza',
                'new_price': 'Precio Venta',
                'transactioncurrencyidname': 'Moneda',
                'new_pricetypename': 'Tipo de Tarifa',
                'statecodename': 'Estado'
            })
            
            columnas_precios = [
                'Código de Recambio', 'Descripción de la Pieza', 
                'Precio Venta', 'Moneda', 'Tipo de Tarifa', 'Estado'
            ]
            
            df = df.fillna("")
            df = df.replace("nan", "")
            columnas_presentes = [col for col in columnas_precios if col in df.columns]
            return df[columnas_presentes].reset_index(drop=True)
        
        try:
            prices_data = load_data_prices()
            
            col_logo, col_titulo = st.columns([1.5, 5])
            with col_logo:
                try: st.image("logo_empresa.png", width=220)
                except Exception: pass
                    
            with col_titulo:
                st.title("Maestro de Tarifas y Precios de Recambios")
                st.write("Consulta oficial de PVPs y tarifas de distribución para la red Spain OJ.")
                
            st.markdown("---")
            
            # Buscador en pantalla de precios
            c1, c2 = st.columns([2, 1])
            with c1:
                buscar_recambio = st.text_input("🔍 Buscar por Código de recambio o Descripción de pieza:", "").strip()
            with c2:
                # Extrae los tipos de tarifa reales que aplican en España (ej: Wholesale Price, Sales Guide...)
                tarifas_disponibles = ["Todas"] + list(prices_data['Tipo de Tarifa'].dropna().unique())
                tarifa_seleccionada = st.selectbox("Filtrar por Tipo de Tarifa:", tarifas_disponibles)
            
            # Lógica de filtrado en cascada
            df_precios = prices_data.copy()
            if buscar_recambio:
                df_precios = df_precios[
                    df_precios['Código de Recambio'].astype(str).str.contains(buscar_recambio, case=False) |
                    df_precios['Descripción de la Pieza'].astype(str).str.contains(buscar_recambio, case=False)
                ]
            if tarifa_seleccionada != "Todas":
                df_precios = df_precios[df_precios['Tipo de Tarifa'] == tarifa_seleccionada]
                
            st.markdown(f"### 📦 {len(df_precios)} referencias localizadas para España")
            if not df_precios.empty:
                st.dataframe(df_precios, use_container_width=True, hide_index=True)
            else:
                st.warning("⚠️ No se encontraron recambios en España con los criterios introducidos.")
                
        except Exception as e:
            st.error(f"Error al procesar el maestro de precios: {e}")
