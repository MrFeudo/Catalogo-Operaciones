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

# ==========================================
# 2. SEGURIDAD (Acceso único global)
# ==========================================
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
    # PANTALLA 1: TIEMPOS DE TALLER
    # =========================================================================
    if opcion_menu == "📋 Tiempos de Taller":
        
        @st.cache_data
        def load_data_tiempos_v3():
            df = pd.read_excel("DMS_Active_Spare_Parts.xlsx", sheet_name="new_srv_workhours")
            
            # Limpieza de espacios en los nombres de las columnas
            df.columns = df.columns.astype(str).str.strip()
            
            # Mapeamos usando "organization" que es el nombre real en esta pestaña
            df = df.rename(columns={
                'new_productmodel_idname': 'Modelo',
                'new_product_idname': 'Nombre de la Pieza',
                'new_code': 'Código de Referencia',
                'new_name': 'Operación Técnica',
                'new_standardhour': 'Tiempo Estándar (UT/Horas)',
                'new_remark': 'Notas / Exclusiones',
                'Organization': 'Mercado / Organización',  # <-- Corregido aquí
                'statecodename': 'Estado'
            })
            
            columnas_finales = [
                'Modelo', 'Nombre de la Pieza', 'Código de Referencia', 
                'Operación Técnica', 'Tiempo Estándar (UT/Horas)', 'Notas / Exclusiones',
                'Mercado / Organización', 'Estado'
            ]
            
            df = df.fillna("")
            df = df.replace("nan", "")
            
            columnas_presentes = [col for col in columnas_finales if col in df.columns]
            return df[columnas_presentes].reset_index(drop=True)

        try:
            data = load_data_tiempos_v3()
            
            st.title("🚗 Catálogo Operaciones de mano de obra")
            st.write("Consulta piezas, modelos y tiempos asignados directamente desde el DMS.")
            st.markdown("---")

            # --- FILA 1 DE FILTROS ---
            col1, col2, col3 = st.columns([1, 1.5, 1.5])
            with col1:
                modelos_disponibles = ["Todos"] + list(data['Modelo'].dropna().unique())
                modelo_seleccionado = st.selectbox("1. Filtrar por Modelo:", modelos_disponibles)
            with col2:
                buscar_pieza = st.text_input("2. Buscar por Nombre o Código de pieza:", "").strip()
            with col3:
                buscar_operacion = st.text_input("3. Buscar por tipo de operación (ej: Remove, Paint...):", "").strip()

            # --- FILA 2 DE FILTROS ---
            col_m, col_e = st.columns([2, 2])
            
            with col_m:
                if 'Mercado / Organización' in data.columns:
                    mercados_disponibles = ["Todos"] + [str(m).strip() for m in data['Mercado / Organización'].unique() if str(m).strip() != ""]
                    
                    # Preselección inteligente para Spain OJ
                    indice_defecto = 0
                    for idx, m in enumerate(mercados_disponibles):
                        if "spain" in m.lower() or "oj spain" in m.lower():
                            indice_defecto = idx
                            break
                    
                    mercado_seleccionado = st.selectbox("Filtrar por Mercado / Organización (Taller):", mercados_disponibles, index=indice_defecto)
                else:
                    mercado_seleccionado = "Todos"
                    
            with col_e:
                if 'Estado' in data.columns:
                    estados_disponibles = ["Todos"] + [str(e).strip() for e in data['Estado'].unique() if str(e).strip() != ""]
                    indice_est_defecto = estados_disponibles.index("Active") if "Active" in estados_disponibles else 0
                    estado_seleccionado = st.selectbox("Filtrar por Estado de Operación (Taller):", estados_disponibles, index=indice_est_defecto)
                else:
                    estado_seleccionado = "Todos"

            # --- LÓGICA DE FILTRADO ---
            df_filtrado = data.copy()
            
            if modelo_seleccionado != "Todos":
                df_filtrado = df_filtrado[df_filtrado['Modelo'] == modelo_seleccionado]
                
            if mercado_seleccionado != "Todos" and 'Mercado / Organización' in df_filtrado.columns:
                df_filtrado = df_filtrado[df_filtrado['Mercado / Organización'].astype(str).str.strip() == mercado_seleccionado]
                
            if estado_seleccionado != "Todos" and 'Estado' in df_filtrado.columns:
                df_filtrado = df_filtrado[df_filtrado['Estado'].astype(str).str.strip() == estado_seleccionado]

            if buscar_pieza:
                df_filtrado = df_filtrado[
                    df_filtrado['Nombre de la Pieza'].astype(str).str.contains(buscar_pieza, case=False, na=False) |
                    df_filtrado['Código de Referencia'].astype(str).str.contains(buscar_pieza, case=False, na=False)
                ]
                
            if buscar_operacion:
                df_filtrado = df_filtrado[df_filtrado['Operación Técnica'].astype(str).str.contains(buscar_operacion, case=False, na=False)]

            # --- TABLA DE TIEMPOS ---
            st.markdown(f"### 📋 Resultados encontrados: {len(df_filtrado)} operaciones")
            if not df_filtrado.empty:
                st.dataframe(df_filtrado, use_container_width=True, hide_index=True)
            else:
                st.warning("⚠️ No se encontraron operaciones con los criterios seleccionados.")
                
        except Exception as e:
            st.error(f"Error al procesar la base de datos de tiempos: {e}")

    # =========================================================================
    # PANTALLA 2: PRECIOS DE RECAMBIOS
    # =========================================================================
    elif opcion_menu == "💰 Precios de Recambios":
        
        @st.cache_data
        def load_prices_nueva_version():
            df = pd.read_excel("DMS_Active_Spare_Parts.xlsx", sheet_name="Parts price")
            df.columns = df.columns.astype(str).str.strip()
            
            # Mapeamos usando "new_businessunit_idname" que corresponde a esta hoja
            df = df.rename(columns={
                'new_partscode': 'Código de Recambio',
                'new_product_idname': 'Descripción de la Pieza',
                'new_price': 'Precio Venta',
                'transactioncurrencyidname': 'Moneda',
                'new_pricetypename': 'Tipo de Tarifa',
                'new_businessunit_idname': 'Mercado / Organización',
                'statecodename': 'Estado'
            })
            
            columnas_finales_precios = [
                'Código de Recambio', 'Descripción de la Pieza', 
                'Precio Venta', 'Moneda', 'Tipo de Tarifa', 
                'Mercado / Organización', 'Estado'
            ]
            
            df = df.fillna("")
            df = df.replace("nan", "")
            
            columnas_visibles = [col for col in columnas_finales_precios if col in df.columns]
            return df[columnas_visibles].reset_index(drop=True)

        try:
            prices_data = load_prices_nueva_version()
            
            st.title("💰 Maestro de Tarifas y Precios de Recambios")
            st.write("Consulta oficializada de precios y tarifas de distribución vigentes.")
            st.markdown("---")
            
            # --- FILTROS DE PRECIOS ---
            col_busc, col_org_p, col_tar = st.columns([2, 1, 1])
            
            with col_busc:
                buscar_recambio = st.text_input("🔍 Buscar por Código de recambio o Descripción de pieza:", "").strip()
                
            with col_org_p:
                mercados_disponibles = ["Todos"] + [str(m).strip() for m in prices_data['Mercado / Organización'].unique() if str(m).strip() != ""]
                
                indice_defecto = 0
                for idx, m in enumerate(mercados_disponibles):
                    if "spain" in m.lower() or "oj spain" in m.lower():
                        indice_defecto = idx
                        break
                        
                mercado_seleccionado = st.selectbox("Filtrar por Mercado / Organización:", mercados_disponibles, index=indice_defecto)
                
            with col_tar:
                tarifas_disponibles = ["Todas"] + [str(t).strip() for t in prices_data['Tipo de Tarifa'].unique() if str(t).strip() != ""]
                tarifa_seleccionada = st.selectbox("Filtrar por Tipo de Tarifa:", tarifas_disponibles)

            # --- LÓGICA DE FILTRADO ---
            df_final_precios = prices_data.copy()
            
            if mercado_seleccionado != "Todos":
                df_final_precios = df_final_precios[df_final_precios['Mercado / Organización'].astype(str).str.strip()]
