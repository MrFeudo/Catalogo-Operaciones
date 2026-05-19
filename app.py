import streamlit as st
import pandas as pd

st.set_page_config(page_title="Buscador Técnico de Operaciones", layout="wide")

# ==========================================
# 1. LOGO EN LA BARRA LATERAL (SIEMPRE VISIBLE)
# ==========================================
try:
    st.sidebar.image("logo_empresa.png", use_container_width=True)
except Exception:
    # Si aún no se ha subido el logo a GitHub, muestra un texto limpio transitorio
    st.sidebar.write("🏢 **OMODA & JAECOO**")

st.sidebar.markdown("---")

# 2. SEGURIDAD
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
    # 3. CARGA DEL EXCEL
    @st.cache_data(ttl=600)
    def load_data():
        df = pd.read_excel("DMS_Active_Spare_Parts.xlsx", sheet_name="new_srv_workhours")
        
        # Pasamos a texto limpio las columnas de control
        df['new_product_idname'] = df['new_product_idname'].astype(str).str.strip()
        df['new_stationname'] = df['new_stationname'].astype(str).str.strip()
        
        # Filtro estricto + excepción universal (999999)
        tiene_nombre = (df['new_product_idname'] != "") & (df['new_product_idname'] != "nan")
        es_universal = df['new_stationname'] == "999999"
        df = df[tiene_nombre | es_universal].copy()
        
        # MAPEO CON LOS NOMBRES REALES DE TU EXCEL
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
        
        # Columnas que se mostrarán en la tabla final
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
        
        # ==========================================
        # 4. CABECERA PRINCIPAL CON LOGO Y TÍTULO
        # ==========================================
        col_logo, col_titulo = st.columns([1.5, 5])
        
        with col_logo:
            try:
                # Ajustamos el tamaño para que mantenga buena proporción horizontal
                st.image("logo_empresa.png", width=220)
            except Exception:
                pass
                
        with col_titulo:
            st.title("Catálogo Operaciones de mano de obra")
            st.write("Consulta piezas, modelos y tiempos activos en DMS.")
            
        st.markdown("---")

        # 5. INTERFAZ DE FILTROS (Fila 1)
        col1, col2, col3 = st.columns([1, 1.5, 1.5])
        
        with col1:
            modelos_disponibles = ["Todos"] + list(data['Modelo'].dropna().unique())
            modelo_seleccionado = st.selectbox("1. Filtrar por Modelo:", modelos_disponibles)
            
        with col2:
            buscar_pieza = st.text_input("2. Buscar por Nombre o Código de pieza:", "").strip()
            
        with col3:
            buscar_operacion = st.text_input("3. Buscar por tipo de operación (ej: Remove, Paint...):", "").strip()

        # INTERFAZ DE FILTROS (Fila 2)
        col_org, col_est = st.columns(2)
        
        with col_org:
            org_disponibles = ["Todas"] + list(data['Organización'].dropna().unique())
            org_seleccionada = st.selectbox("4. Filtrar por Organización:", org_disponibles)
            
        with col_est:
            est_disponibles = ["Todos"] + list(data['Estado'].dropna().unique())
            est_seleccionado = st.selectbox("5. Filtrar por Estado (Activo/Inactivo):", est_disponibles)

        # 6. LÓGICA DE FILTRADO (Cascada independiente limpia)
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

        # 7. RESULTADOS
        st.markdown(f"### 📋 Resultados encontrados: {len(df_filtrado)} operaciones")
        
        if not df_filtrado.empty:
            st.dataframe(df_filtrado, use_container_width=True, hide_index=True)
        else:
            st.warning("⚠️ No se encontraron operaciones con los criterios seleccionados.")
            
    except Exception as e:
        st.error(f"Error al procesar la base de datos: {e}")
