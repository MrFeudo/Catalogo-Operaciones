import streamlit as st
import pandas as pd
import datetime
import io
import unicodedata
from google import genai
from google.genai import types
from google.oauth2 import service_account
from PIL import Image

def normalizar_texto(texto):
    texto = str(texto)
    return ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn').lower()

st.set_page_config(page_title="Buscador Técnico OMODA & JAECOO", layout="wide")

# =========================================================================
# 1. INICIALIZACIÓN ABSOLUTA DEL SESSION STATE
# =========================================================================
if "lista_solicitudes" not in st.session_state:
    st.session_state.lista_solicitudes = []

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if "idioma" not in st.session_state:
    st.session_state.idioma = "Español"

# URL Directa (Raw) al archivo Excel en GitHub
URL_GITHUB_EXCEL = "https://github.com/MrFeudo/Catalogo-Operaciones/raw/main/DMS_Active_Spare_Parts.xlsb"

# =========================================================================
# 2. DICCIONARIO DE TRADUCCIÓN (Internacionalización - i18n para TFM)
# =========================================================================
IDIOMAS = {
    "Español": {
        "menu_titulo": "### 🗺️ Menú de Navegación",
        "menu_radio": "Selecciona una herramienta:",
        "menu_taller": "📋 Tiempos de Taller",
        "menu_precios": "💰 Precios de Recambios",
        "menu_solicitar": "📝 Solicitar Operación",
        "pass_titulo": "🔐 Acceso Red de Dealers",
        "pass_input": "Introduce la contraseña de acceso:",
        "pass_boton": "Entrar",
        "pass_error": "❌ Contraseña incorrecta",
        "taller_titulo": "🚗 Catálogo Operaciones de mano de obra",
        "taller_sub": "Consulta piezas, modelos y tiempos asignados directamente desde el DMS.",
        "f_modelo": "1. Filtrar por Modelo:",
        "f_pieza": "2. Buscar por Nombre o Código de pieza:",
        "f_operacion": "3. Buscar por tipo de operación (ej: Remove, Paint...):",
        "f_mercado_taller": "Filtrar por Mercado / Organización (Taller):",
        "f_estado_taller": "Filtrar por Estado de Operación (Taller):",
        "res_taller": "### 📋 Resultados encontrados: {} operaciones",
        "warn_taller": "⚠️ No se encontraron operaciones con los criterios seleccionados.",
        "err_taller": "Error al procesar la base de datos de tiempos: {}",
        "precios_titulo": "💰 Maestro de Tarifas y Precios de Recambios",
        "precios_sub": "Consulta oficializada de precios y tarifas de distribución vigentes.",
        "f_buscar_recambio": "🔍 Buscar por Código de recambio o Descripción de pieza:",
        "f_mercado_precios": "Filtrar por Mercado / Organización:",
        "f_tarifa": "Filtrar por Tipo de Tarifa:",
        "res_precios": "### 📦 {} referencias de recambios localizadas",
        "warn_precios": "⚠️ No se encontraron recambios con los criterios seleccionados.",
        "err_precios": "Error al procesar el maestro de precios: {}",
        "todos": "Todos",
        "todas": "Todas",
        "filtro_modelo": "Filtrar por Modelo:",
        "solicitar_titulo": "📝 Solicitud de Operaciones Adicionales de Mano de Obra",
        "solicitar_sub": "Utilice este formulario para solicitar el alta de nuevas operaciones o precios en el maestro de HQ.",
        "form_sub": "Datos de la Solicitud (Campos obligatorios *)",
        "form_marca": "Marca del vehículo *",
        "form_modelo": "INTRODUCIR MODELO *",
        "form_vin": "INTRODUCIR VIN (Bastidor) *",
        "form_vin_holder": "17 caracteres",
        "form_dealer": "DEALER (Concesionario) *",
        "form_hq_code": "CÓDIGO DE PRODUCTO (Asignado por HQ)",
        "form_ref": "REFERENCIA DE PIEZA (Opcional)",
        "form_ref_holder": "Ej. 7365747465AA",
        "form_op": "OPERACIÓN QUE SE SOLICITA AÑADIR *",
        "form_op_holder": "Describa detalladamente la operation técnica o falta de precio que requiere el taller...",
        "form_btn": "Enviar Solicitud a Central",
        "err_campos": "❌ Por favor, rellene todos los campos obligatorios (*).",
        "err_vin_corto": "❌ El VIN introducido es demasiado corto. Revíselo.",
        "success_sheet": "✅ ¡Solicitud registrada con éxito! Los datos se han volcado a la plantilla de Central.",
        "warn_contingencia": "⚠️ Formulario correcto, guardado en modo de contingencia local."
    },
    "English": {
        "menu_titulo": "### 🗺️ Navigation Menu",
        "menu_radio": "Select a tool:",
        "menu_taller": "📋 Workshop Times",
        "menu_precios": "💰 Spare Parts Prices",
        "menu_solicitar": "📝 Request Operation",
        "pass_titulo": "🔐 Dealer Network Access",
        "pass_input": "Enter access password:",
        "pass_boton": "Login",
        "pass_error": "❌ Incorrect password",
        "taller_titulo": "🚗 Labor Operations Catalog",
        "taller_sub": "Consult parts, models, and assigned times directly from the DMS.",
        "f_modelo": "1. Filter by Model:",
        "f_pieza": "2. Search by Part Name or Code:",
        "f_operacion": "3. Search by operation type (e.g., Remove, Paint...):",
        "f_mercado_taller": "Filter by Market / Organization (Workshop):",
        "f_estado_taller": "Filter by Operation Status (Workshop):",
        "res_taller": "### 📋 Results found: {} operations",
        "warn_taller": "⚠️ No operations found matching the selected criteria.",
        "err_taller": "Error processing workshop times database: {}",
        "precios_titulo": "💰 Master Rate & Spare Parts Prices",
        "precios_sub": "Official consultation of current prices and distribution rates.",
        "f_buscar_recambio": "🔍 Search by Part Code or Description:",
        "f_mercado_precios": "Filter by Market / Organization:",
        "f_tarifa": "Filter by Rate Type:",
        "res_precios": "### 📦 {} spare parts references located",
        "warn_precios": "⚠️ No spare parts found matching the selected criteria.",
        "err_precios": "Error processing master price list: {}",
        "todos": "All",
        "todas": "All",
        "filtro_modelo": "Filter by Model:",
        "solicitar_titulo": "📝 Request for Additional Labor Operations",
        "solicitar_sub": "Use this form to request new operations or prices to be added to HQ master list.",
        "form_sub": "Request Details (* Required fields)",
        "form_marca": "Vehicle Brand *",
        "form_modelo": "ENTER MODEL *",
        "form_vin": "ENTER VIN (Chassis) *",
        "form_vin_holder": "17 characters",
        "form_dealer": "DEALER *",
        "form_hq_code": "PRODUCT CODE (Assigned by HQ)",
        "form_ref": "PART REFERENCE (Optional)",
        "form_ref_holder": "e.g., 7365747465AA",
        "form_op": "OPERATION REQUESTED TO BE ADDED *",
        "form_op_holder": "Describe in detail the technical operation or missing price required by the workshop...",
        "form_btn": "Send Request to HQ",
        "err_campos": "❌ Please fill in all required fields (*).",
        "err_vin_corto": "❌ The entered VIN is too short. Please check it.",
        "success_sheet": "✅ Request successfully registered! Data transferred to the HQ template.",
        "warn_contingencia": "⚠️ Form valid, saved in local contingency mode."
    },
    "Chinese (中文)": {
        "menu_titulo": "### 🗺️ 导航菜单",
        "menu_radio": "选择工具:",
        "menu_taller": "📋 车间工时",
        "menu_precios": "💰 零配件价格",
        "menu_solicitar": "📝 请求操作",
        "pass_titulo": "🔐 经销商网络访问",
        "pass_input": "输入访问密码:",
        "pass_boton": "登录",
        "pass_error": "❌ 密码错误",
        "taller_titulo": "🚗 工时操作目录",
        "taller_sub": "直接从 DMS 查询零件、车型和分配的时间。",
        "f_modelo": "1. 按车型筛选:",
        "f_pieza": "2. 按零件名称或代码搜索:",
        "f_operacion": "3. 按操作类型搜索 (例如: Remove, Paint...):",
        "f_mercado_taller": "按市场 / 组织筛选 (车间):",
        "f_estado_taller": "按操作状态筛选 (车间):",
        "res_taller": "### 📋 找到的结果: {} 个操作",
        "warn_taller": "⚠️ 未找到符合选择条件的工时操作。",
        "err_taller": "处理车间工时数据库时出错: {}",
        "precios_titulo": "💰 零售价与零配件价格总表",
        "precios_sub": "官方查询现行价格及分销费率。",
        "f_buscar_recambio": "🔍 按零件代码或描述搜索:",
        "f_mercado_precios": "按市场 / 组织筛选:",
        "f_tarifa": "按费率类型筛选:",
        "res_precios": "### 📦 已定位 {} 个零配件参考",
        "warn_precios": "⚠️ 未找到符合选择条件的零配件。",
        "err_precios": "处理价格总表时出错: {}",
        "todos": "全部",
        "todas": "全部",
        "filtro_modelo": "按车型过滤:",
        "solicitar_titulo": "📝 申请新增工时操作",
        "solicitar_sub": "使用此表单申请在总部(HQ)主数据中添加新工时操作 or 价格。",
        "form_sub": "申请信息 (* 为必填项)",
        "form_marca": "车辆品牌 *",
        "form_modelo": "输入车型 *",
        "form_vin": "输入 VIN (车架号) *",
        "form_vin_holder": "17位字符",
        "form_dealer": "经销商 *",
        "form_hq_code": "产品代码 (由总部分配)",
        "form_ref": "零件编号 (选填)",
        "form_ref_holder": "例如: 7365747465AA",
        "form_op": "申请添加的操作内容 *",
        "form_op_holder": "请详细描述车间所需的工时操作 or 缺失的价格...",
        "form_btn": "发送申请至总部",
        "err_campos": "❌ 请填写所有必填项 (*)。",
        "err_vin_corto": "❌ 输入的 VIN 太短，请检查。",
        "success_sheet": "✅ 申请登记成功！数据已同步至总部模板。",
        "warn_contingencia": "⚠️ 表单正确，已保存至本地应急模式。"
    }
}

# ==========================================
# 3. BARRA LATERAL: LOGO + SELECCIÓN IDIOMA + MENÚ
# ==========================================
try:
    st.sidebar.image("logo_empresa.png", use_container_width=True)
except Exception:
    st.sidebar.write("🏢 **OMODA & JAECOO**")

st.sidebar.markdown("---")

# CORRECCIÓN DE SEGURIDAD: Añadimos 'key' explícita para evitar duplicación de ID en recargas
idioma_seleccionado = st.sidebar.selectbox(
    "🌐 Language / Idioma / 语言:",
    ["Español", "English", "Chinese (中文)"],
    index=["Español", "English", "Chinese (中文)"].index(st.session_state.idioma),
    key="selector_idioma_global"
)
st.session_state.idioma = idioma_seleccionado
txt = IDIOMAS[st.session_state.idioma]

st.sidebar.markdown("---")
st.sidebar.sidebar_markdown_target = st.sidebar.markdown(txt["menu_titulo"])

opcion_menu = st.sidebar.radio(
    txt["menu_radio"],
    [txt["menu_taller"], txt["menu_solicitar"], "🧠 Consultorio Técnico IA"],
    key="menu_navegacion_app"
)

def buscador_inteligente_excel(consulta_usuario, df_contexto):
    try:
        if "GEMINI_API_KEY" not in st.secrets:
            return "⚠️ **Error**: No se ha encontrado la clave API en st.secrets."
            
        client = genai.Client(api_key=st.secrets["GEMINI_API_KEY"])

        # 🎯 1. MAPEO SEMÁNTICO DE RAÍCES
        mapa_raices = {
            # --- 🛠️ ACCIONES, VERBOS Y REGLAS DE TRABAJO ---
            "cambiar": "remove and reinstall|replace|remove|reinstall",
            "cambio": "remove and reinstall|replace|remove|reinstall",
            "sustituir": "remove and reinstall|replace|remove|reinstall",
            "sustitucion": "remove and reinstall|replace|remove|reinstall",
            "reemplazar": "remove and reinstall|replace|remove|reinstall",
            "reemplazo": "remove and reinstall|replace|remove|reinstall",
            "desmontar": "remove", 
            "montar": "reinstall",
            "comprobar": "check|inspection|test|diagnostic|measurement",
            "verificar": "check|inspection|test|diagnostic",
            "revisar": "check|inspection|test|diagnostic",
            "diagnosis": "check|inspection|test|diagnostic|troubleshooting",
            "diagnostico": "check|inspection|test|diagnostic",
            "actualizar": "refresh|update|software|flash",
            "programar": "refresh|update|software|flash|coding|program",
            "codificar": "refresh|update|software|flash|coding|program",
            "reprogramar": "refresh|update|software|flash|coding|program",
            "ajustar": "adjust|adjustment|alignment|calibrate|calibration",
            "alinear": "alignment|adjust",
            "calibrar": "calibrate|calibration",
            "limpiar": "clean|cleaning|wash",
            "pulir": "polishing|polish", "pulido": "polishing|polish",
            
            # --- 🔌 CENTRALITAS, MÓDULOS Y ELECTRÓNICA DE CONTROL ---
            "ecu": "ecu|engine control unit|engine control module",
            "ems": "ems|engine management system",
            "mcu": "mcu|motor control unit|motor control module",
            "vcu": "vcu|vehicle control unit|vehicle control module",
            "hcu": "hcu|hybrid control unit|hybrid module",
            "tcu": "tcu|transmission control unit|transmission module",
            "hpu": "hpu|hybrid powertrain unit",
            "tecu": "tecu|traction electric control unit",
            "bcm": "bcm|body control module|bdm", 
            "bdm": "bcm|body control module|bdm",
            "icm": "icm|instrument cluster module", 
            "clm": "climate|clm|hvac", 
            "ihu": "ihu|infotainment head unit|display",
            "cgw": "cgw|central gateway|gateway", 
            "pas": "pas|passenger area system",
            "tbox": "t-box|tbox|telematics|information communication module",
            "centralita": "control unit|control module|ecu|bcm|mcu|vcu|tcu|hcu",
            "centralitas": "control unit|control module",
            "modulo": "control module|module", 
            "modulos": "control module|module",

            # --- ⚡ BATERÍAS, ALTA TENSIÓN Y SISTEMA ELÉCTRICO ---
            "bateria": "battery|storage battery|bms|tecu", 
            "vateria": "battery|storage battery", # Errata
            "baterias": "battery|storage battery",
            "traccion": "traction|traction battery", 
            "alta tension": "high voltage",
            "bms": "bms|battery management system",
            "cdu": "cdu|conversion & distribution unit|conversion distribution",
            "alternador": "alternator|generator",
            "arranque": "starter|starter motor", "motor arranque": "starter",
            "cable": "wiring|harness|wire|cable", 
            "cableado": "wiring|harness|wire|cable", 
            "instalacion": "wiring|harness",
            "mazo": "wiring|harness",
            "fusible": "fuse", "fusibles": "fuse|box", "caja fusibles": "fuse block|fuse box",
            "bujia": "spark plug", "bujias": "spark plug",
            "bobina": "ignition coil", "bobinas": "ignition coil",

            # --- 🛡️ SEGURIDAD, ASISTENCIAS ADAS Y SENSORES ---
            "fcm": "fcm|front camera module|forward camera|front view camera", 
            "camara": "camera|fcm|avm|rear view", "camaras": "camera",
            "frm": "frm|front radar module|forward radar", 
            "radar": "radar|frm|bsd", "radares": "radar",
            "avm": "avm|around view monitor|360 camera",
            "bsd": "bsd|blind spot detection|blind spot monitoring", 
            "punto ciego": "blind spot|bsd",
            "tpms": "tpms|tire pressure monitor", 
            "pdc": "pdc|park distance control|parking sensor", 
            "aparcamiento": "park|pdc|parking", 
            "sensor": "sensor|probe|detector", "sensores": "sensor",
            "sonda": "sensor|oxygen sensor|lambda", "lambda": "oxygen sensor|lambda",
            "airbag": "airbag|air bag|abm|srs|curtain airbag", 
            "airbags": "airbag|air bag",
            "abm": "abm|air bag module",
            "srs": "srs|supplemental restraint system",
            "cinturon": "seatbelt|seat belt|belt", "cinturones": "seatbelt|seat belt",
            "hebilla": "buckle", "hebillas": "buckle",
            "pretensor": "pretensioner",

            # --- ⚙️ MOTOR, ADMISIÓN, ESCAPE Y REFRIGERACIÓN ---
            "motor": "engine assy|motor|engine", "motores": "engine",
            "culata": "cylinder head", "piston": "piston", "biela": "connecting rod",
            "cigüeñal": "crankshaft", "arbol": "camshaft", "levas": "camshaft",
            "valvula": "valve|solenoid valve", "valvulas": "valve",
            "turbo": "turbocharger|turbo", "turbocompresor": "turbocharger",
            "intercooler": "intercooler|charge air cooler",
            "colector": "manifold", "admision": "intake", "escape": "exhaust",
            "catalizador": "catalytic converter|catalyst", "fap": "dpf|particulate filter",
            "silenciador": "muffler|exhaust silencer",
            "radiador": "radiator", "intercambiador": "heat exchanger",
            "ventilador": "fan|cooling fan", "electroventilador": "cooling fan",
            "bomba": "pump|water pump|oil pump|fuel pump", "bombas": "pump",
            "bomba agua": "water pump", "bomba aceite": "oil pump", "bomba combustible": "fuel pump",
            "termostato": "thermostat",
            "canister": "canister|evap canister", "vapores": "canister|evap|solenoid|pipe",

            # --- 🚗 TRANSMISIÓN, CAJA DE CAMBIOS Y EMBRAGUE ---
            "dct": "dct|dual clutch transmission|double clutch",
            "cambio": "transmission|gearbox|dct|gearshift", 
            "caja": "transmission|gearbox", "caja cambios": "transmission|gearbox",
            "embrague": "clutch", "bimasa": "dual mass flywheel|flywheel",
            "volante": "flywheel", # Nota: volante motor es flywheel, volante direccion es steering wheel
            "volante motor": "flywheel",
            "palier": "drive shaft|axle shaft|half shaft", "palieres": "drive shaft",
            "transmision": "transmission|propeller shaft|drive shaft",
            "diferencial": "differential", "reductora": "reducer",
            "selector": "selector|shifter|gearshift lever",

            # --- 🥾 CHASIS, SUSPENSIÓN Y FRENOS ---
            "esp": "esp|electronic stability program",
            "eps": "eps|electric power steering",
            "epb": "epb|electrical park brake|parking brake",
            "ipb": "ipb|integrated power brake|power brake",
            "abs": "abs|anti-lock brake system",
            "freno": "brake|ipb|epb|abs", "frenos": "brake|abs",
            "pastilla": "pads|brake pads", "pastillas": "pads|brake pads",
            "disco": "disc|brake disc", "discos": "disc|brake disc",
            "pinza": "caliper|brake caliper", "pinzas": "caliper",
            "latiguillo": "brake hose", "servo": "brake booster|power brake",
            "amortiguador": "shock absorber|strut|damper", "amortiguadores": "shock absorber|strut",
            "muelle": "spring|coil spring", "muelles": "spring",
            "ballesta": "leaf spring",
            "barra": "bar|stabilizer bar|rod", "estabilizadora": "stabilizer",
            "trapecio": "control arm|suspension arm|wishbone", "brazo": "control arm|suspension arm",
            "mangueta": "knuckle|steering knuckle",
            "buge": "hub|wheel hub", "cojinete": "bearing", "rodamiento": "bearing",
            "direccion": "steering|eps", "cremallera": "steering gear|steering rack",

            # --- 📦 CARROCERÍA, INTERIOR, EXTERIOR Y COLISIÓN ---
            "capo": "hood|engine hood",
            "paragolpes": "bumper", "defensa": "bumper", "parachoques": "bumper",
            "faro": "headlamp|headlight", "faros": "headlamp", 
            "piloto": "lamp|rear lamp|tail lamp", "pilotos": "lamp|tail lamp",
            "antiniebla": "fog lamp|foglight",
            "intermitente": "turn signal|indicator",
            "espejo": "mirror|rearview mirror", "retrovisor": "mirror|rearview mirror",
            "aleta": "fender|wing", "aletas": "fender",
            "puerta": "door", "puertas": "door", "porton": "tailgate|back door|rear door",
            "techo": "sunroof|roof|panoramic roof", "solar": "sunroof",
            "cristal": "glass|window", "luna": "windshield|windscreen|glass", "parabrisas": "windshield|windscreen",
            "elevalunas": "window regulator|window lifter",
            "cerradura": "door lock|lock assy", "cierre": "lock|latch",
            "manilla": "handle|door handle", "maneta": "handle",
            "asiento": "seat", "asientos": "seat",
            "salpicadero": "dashboard|instrument panel",
            "moldura": "molding|trim", "molduras": "molding|trim",
            "limpiaparabrisas": "wiper|wiper blade", "limpiaparabrisas": "wiper", "motor limpia": "wiper motor",

            # --- 🩻 SUPORTACIÓN, ELEMENTOS DE UNIÓN Y MENUDENCIA ---
            "soporte": "bracket|support|mount|holder", "soportes": "bracket|support|mount",
            "cuna": "subframe|cradle|bracket|salver|tray", 
            "bandeja": "tray|salver",
            "tapa": "cover|cap|lid", "cubierta": "cover|protector", "protector": "protector|shield|guard",
            "varilla": "rod|stay", "tirante": "rod|link|stay",
            "placa": "plate", "panel": "panel",
            "grapa": "clip|retainer", "tornillo": "bolt|screw", "tuerca": "nut",
            "abrazadera": "clamp|clip",

            # --- 🧪 FLUIDOS, JUNTAS Y FILTROS ---
            "filtro": "filter", "filtros": "filter",
            "filtro aceite": "oil filter", "filtro aire": "air filter", "filtro habitaculo": "cabin filter|pollen filter",
            "junta": "gasket|seal", "juntas": "gasket|seal", "reten": "oil seal|seal",
            "aceite": "oil|lubricant", "liquido": "fluid|liquid", "refrigerante": "coolant",
            "anticongelante": "coolant",
            "tubo": "pipe|tube|hose", "manguito": "hose", "conducto": "pipe|line|duct",

            # --- 📍 UBICACIONES, ORIENTACIÓN Y LADOS ---
            "delantero": "fr", "delantera": "fr", "frontal": "fr", "alante": "fr",
            "trasero": "rr", "trasera": "rr", "posterior": "rr", "atras": "rr",
            "izquierdo": "lh", "izquierda": "lh", "izq": "lh", "izda": "lh",
            "derecho": "rh", "derecha": "rh", "der": "rh", "drcha": "rh",
            "superior": "upper", "inferior": "lower", "interno": "inner", "externo": "outer",
            "central": "central|middle", "lateral": "side"
        }

        # --- DICCIONARIO DE EXPANSIÓN DE MODELOS (MANTENIDO) ---
        abreviaturas_modelos = {
            "j5": "jaecoo 5", "jaecoo5": "jaecoo 5", "j-5": "jaecoo 5",
            "j7": "jaecoo 7", "jaecoo7": "jaecoo 7", "j-7": "jaecoo 7",
            "j8": "jaecoo 8", "jaecoo8": "jaecoo 8",
            "o5": "omoda 5", "omoda5": "omoda 5", "o-5": "omoda 5",
            "hibrido": "hev", "electrico": "bev", "gasolina": "ice"
        }
        # Limpieza inicial de texto sin acentos
        consulta_limpia = consulta_usuario.lower().strip()
        for orig, dest in [("í", "i"), ("ó", "o"), ("á", "a"), ("é", "e"), ("ú", "u"), ("ñ", "n")]:
            consulta_limpia = consulta_limpia.replace(orig, dest)

        # Expandimos los modelos
        for abrev, mod_real in abreviaturas_modelos.items():
            if abrev in consulta_limpia.split() or abrev in consulta_limpia:
                consulta_limpia = consulta_limpia.replace(abrev, mod_real)

        # Lista de palabras sueltas escritas en la consulta
        lista_palabras_usuario = consulta_limpia.split()

        # Construimos las traducciones al inglés
        palabras_regex = []
        for esp, eng in mapa_raices.items():
            if esp in consulta_limpia:
                palabras_regex.extend(eng.split('|'))

        for p in lista_palabras_usuario:
            if len(p) > 2 and p not in ["quiero", "para", "con", "del", "una", "uno", "el", "la", "los", "las", "este", "un", "de"]:
                if not (p.isdigit() and len(p) == 1):
                    palabras_regex.append(p)

        palabras_regex = list(set(palabras_regex))

        # 🔍 2. MOTOR DE FILTRADO CORREGIDO POR INTERSECCIÓN DE PALABRAS COMPLETAS
        try:
            terminos_manuales = ["manual", "adicional", "extra", "tiempo mas", "añadir horas", "universal", "marron", "baremo no"]
            if any(tm in lista_palabras_usuario for tm in terminos_manuales):
                df_recortado = df_contexto[df_contexto['Operación Técnica'].astype(str).str.lower().str.contains("universal", na=False)].head(20)
            else:
                df_base = df_contexto.copy()
                
                # Forzamos minúsculas en las columnas
                for col in ['Modelo', 'Nombre de la Pieza', 'Operación Técnica']:
                    df_base[col] = df_base[col].astype(str).str.lower().str.strip()

                # Criba por marca
                if "omoda" in consulta_limpia:
                    df_base = df_base[df_base['Modelo'].str.contains("omoda", na=False)]
                elif "jaecoo" in consulta_limpia:
                    df_base = df_base[df_base['Modelo'].str.contains("jaecoo", na=False)]

                # Intersección obligatoria del componente base
                componentes_encontrados = []
                for esp, eng in mapa_raices.items():
                    if esp in consulta_limpia and esp not in ["cambiar", "sustituir", "cambio", "sustitucion", "reemplazar", "desmontar", "montar"]:
                        componentes_encontrados.extend(eng.split('|'))
                
                if componentes_encontrados:
                    regex_comp = '|'.join(set(componentes_encontrados))
                    df_base = df_base[df_base['Nombre de la Pieza'].str.contains(regex_comp, na=False) | 
                                      df_base['Operación Técnica'].str.contains(regex_comp, na=False)]

                # 🔴 ALGORITMO DE EXCLUSIÓN MEJORADO (Comprobación por palabra exacta independiente)
                filtros_secundarios = {
                    "wiring|harness|wire": ["cable", "cableado", "instalacion", "mazo"],
                    "sensor": ["sensor", "sonda"],
                    "bracket|salver|tray|support|pressure|rod|plate": ["soporte", "cuna", "bandeja", "tapa", "cubierta", "varilla", "placa"]
                }
                
                for eng_purgar, esp_palabras in filtros_secundarios.items():
                    # REGLA DE ORO: Buscamos si la palabra exacta está en la lista de palabras (no dentro de otra palabra)
                    usuario_pide_secundario = any(w in lista_palabras_usuario for w in esp_palabras)
                    
                    if not usuario_pide_secundario:
                        condicion_purgar = df_base['Nombre de la Pieza'].str.contains(eng_purgar, na=False) | \
                                           df_base['Operación Técnica'].str.contains(eng_purgar, na=False)
                        df_base = df_base[~condicion_purgar]

                # Puntuación final
                df_base['score'] = 0
                if palabras_regex:
                    regex_puntos = '|'.join(palabras_regex)
                    df_base['score'] += df_base['Modelo'].str.contains(regex_puntos, na=False).astype(int) * 5
                    df_base['score'] += df_base['Nombre de la Pieza'].str.contains(regex_puntos, na=False).astype(int) * 10
                    df_base['score'] += df_base['Operación Técnica'].str.contains(regex_puntos, na=False).astype(int) * 10
                    
                    df_recortado = df_base.sort_values(by='score', ascending=False).head(100)
                else:
                    df_recortado = df_base.head(40)

                # Red de seguridad si se vacía
                if df_recortado.empty:
                    df_recortado = df_contexto[df_contexto['Modelo'].astype(str).str.lower().str.contains("omoda", na=False)].head(60)

                df_base.drop(columns=['score'], errors='ignore')

            resumen_excel = df_recortado[['Modelo', 'Nombre de la Pieza', 'Código de Referencia', 'Operación Técnica']].to_string(index=False)
        except Exception as e:
            return f"❌ Error interno al procesar el filtro de relevancia: {str(e)}"

        # 🧠 3. PROMPT MAESTRO FLEXIBLE
        prompt_sistema = (
            "Eres el Buscador Inteligente Avanzado del catálogo oficial de OMODA & JAECOO España.\n\n"
            "MISION DE ANÁLISIS ABIERTO:\n"
            "- El usuario es personal de taller y te va a pedir piezas mezclando motorizaciones o de forma genérica.\n"
            "- Tu objetivo es mostrar TODAS las operaciones válidas que encuentres en el extracto inferior relacionadas con el componente solicitado.\n\n"
            "GUÍA DE TRADUCCIÓN RÁPIDA:\n"
            "- 'FR' = Front (Delantero) | 'RR' = Rear (Trasero)\n"
            "- 'LH' = Left Hand (Izquierdo) | 'RH' = Right Hand (Derecho)\n"
            "- 'Remove and reinstall' / 'Replace' = Cambiar, sustituir, reinstalar, desmontar y montar.\n\n"
            "REGLAS DE SALIDA:\n"
            "1. Devuelve los resultados organizados en una lista Markdown limpia y estructurada.\n"
            "2. Si el componente solicitado no tiene ninguna relación con lo que hay en el extracto inferior, saca el mensaje oficial de derivación al formulario.\n"
            "3. Prohibido inventar códigos de referencia.\n\n"
            f"--- EXTRACTO DE AMPLIO ESPECTRO DEL CATÁLOGO --- \n{resumen_excel}"
        )

        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=[f"Consulta del operario de taller: '{consulta_usuario}'"],
            config=types.GenerateContentConfig(
                system_instruction=prompt_sistema,
                temperature=0.1
            )
        )
        
        if "tokens_totales_input" not in st.session_state:
            st.session_state.tokens_totales_input = 0
            st.session_state.tokens_totales_output = 0
            st.session_state.dinero_total_gastado = 0.0
            st.session_state.ultima_consulta_info = "Ninguna consulta."

        if response.text and response.usage_metadata:
            t_input = response.usage_metadata.prompt_token_count
            t_output = response.usage_metadata.candidates_token_count
            coste = ((t_input * 0.075) / 1_000_000) + ((t_output * 0.30) / 1_000_000)
            
            st.session_state.tokens_totales_input += t_input
            st.session_state.tokens_totales_output += t_output
            st.session_state.dinero_total_gastado += coste
            st.session_state.ultima_consulta_info = f"Última: In: {t_input} | Out: {t_output} (+{coste:.5f}$)"
            
        return response.text if response.text else "❌ No se encontraron coincidencias."
    except Exception as e:
        return f"❌ Error en el motor de la IA de Gemini: {str(e)}"

       # =====================================================================
        # 🧠 3. PROMPT MAESTRO CON CÁLCULO AUTOMÁTICO DE TIEMPOS (UTs y Minutos)
        # =====================================================================
        prompt_sistema = (
            "Eres el Buscador Inteligente Avanzado del catálogo oficial de OMODA & JAECOO España.\n\n"
            "MISION DE ANÁLISIS ABIERTO:\n"
            "- El usuario es personal de taller. Busca operaciones mostrando resultados en una lista Markdown limpia.\n\n"
            "🔴 REGLA MAESTRA DE CÁLCULO Y CONVERSIÓN DE TIEMPOS (OBLIGATORIA):\n"
            "En la columna 'Tiempo Estándar (UT/Horas)' el catálogo te da el valor base en UTs (Unidades de Tiempo).\n"
            "Para cada operación técnica válida que decidas mostrar, estás OBLIGADO a desglosar el tiempo aplicando estrictamente esta fórmula:\n"
            "  - Valor base del catálogo = El número indicado en la tabla (ej: 20, 30, 40...).\n"
            "  - Equivalencia estándar: 100 UTs = 1 Hora = 60 Minutos.\n"
            "  - Tiempo en Horas = Valor base / 100.\n"
            "  - Tiempo en Minutos = (Valor base / 100) * 60.\n\n"
            "FORMATO DE SALIDA DE TIEMPOS (SÍGUELO ESTRICTAMENTE):\n"
            "Al listar cada operación, píntalo con este formato exacto en negrita:\n"
            "- **Tiempo:** X UTs (~Horas: Y hr | ~Minutos: Z min)\n"
            "  *(Ejemplo real si el valor es 20: **Tiempo:** 20 UTs (~Horas: 0.20 hr | ~Minutos: 12 min))*\n"
            "  *(Ejemplo real si el valor es 30: **Tiempo:** 30 UTs (~Horas: 0.30 hr | ~Minutos: 18 min))*\n"
            "  *(Ejemplo real si el valor es 110: **Tiempo:** 110 UTs (~Horas: 1.10 hr | ~Minutos: 66 min))*\n\n"
            "GUÍA DE TRADUCCIÓN RÁPIDA:\n"
            "- 'FR' = Front (Delantero) | 'RR' = Rear (Trasero)\n"
            "- 'LH' = Left Hand (Izquierdo) | 'RH' = Right Hand (Derecho)\n"
            "- 'Remove and reinstall' / 'Replace' = Cambiar, sustituir, reinstalar, desmontar y montar.\n\n"
            "REGLAS DE SALIDA:\n"
            "1. Devuelve los resultados organizados en una lista Markdown limpia, clara y estructurada.\n"
            "2. Si el componente solicitado no tiene ninguna relación con el extracto, saca el mensaje oficial de derivación al formulario.\n"
            "3. Prohibido inventar códigos de referencia o variar los números de las UTs.\n\n"
            f"--- EXTRACTO DEL CATÁLOGO CON TIEMPOS --- \n{resumen_excel}"
        )

        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=[f"Consulta del operario de taller: '{consulta_usuario}'"],
            config=types.GenerateContentConfig(
                system_instruction=prompt_sistema,
                temperature=0.1
            )
        )
        
        if "tokens_totales_input" not in st.session_state:
            st.session_state.tokens_totales_input = 0
            st.session_state.tokens_totales_output = 0
            st.session_state.dinero_total_gastado = 0.0
            st.session_state.ultima_consulta_info = "Ninguna consulta."

        if response.text and response.usage_metadata:
            t_input = response.usage_metadata.prompt_token_count
            t_output = response.usage_metadata.candidates_token_count
            coste = ((t_input * 0.075) / 1_000_000) + ((t_output * 0.30) / 1_000_000)
            
            st.session_state.tokens_totales_input += t_input
            st.session_state.tokens_totales_output += t_output
            st.session_state.dinero_total_gastado += coste
            st.session_state.ultima_consulta_info = f"Última: In: {t_input} | Out: {t_output} (+{coste:.5f}$)"
            
        return response.text if response.text else "❌ No se encontraron coincidencias."
    except Exception as e:
        return f"❌ Error en el motor de la IA de Gemini: {str(e)}"
            
        # =====================================================================
        # 🧠 3. PROMPT MAESTRO ULTRA-FLEXIBLE PARA GEMINI (DAR TODO)
        # =====================================================================
        prompt_sistema = (
            "Eres el Buscador Inteligente Avanzado del catálogo oficial de OMODA & JAECOO España.\n\n"
            "MISION DE ANÁLISIS ABIERTO:\n"
            "- El usuario es personal de taller y va a mil por hora. Te va a pedir piezas mezclando motorizaciones o de forma genérica.\n"
            "- Tu objetivo es mostrar TODAS las operaciones válidas que encuentres en el extracto inferior relacionadas con el componente solicitado.\n\n"
            "🔴 REGLA DE APERTURA TOTAL POR MOTORIZACIONES:\n"
            "- No restrinjas los resultados. Si el usuario te pide la batería de un modelo (ej: Omoda 5) y en el extracto te aparecen las operaciones "
            "tanto de la versión térmica, de la híbrida (HEV) como de la eléctrica (BEV), MUESTRA TODAS LAS VARIANTES desglosadas limpiamente.\n"
            "- Deja que el operario de taller vea las diferentes opciones en la lista para que él elija la que corresponde al coche que tiene físicamente en el elevador.\n\n"
            "GUÍA DE TRADUCCIÓN RÁPIDA:\n"
            "1. POSICIONES:\n"
            "   - 'FR' = Front (Delantero) | 'RR' = Rear (Trasero)\n"
            "   - 'LH' = Left Hand (Izquierdo) | 'RH' = Right Hand (Derecho)\n"
            "2. ACCIONES:\n"
            "   - 'Remove and reinstall' / 'Replace' = Cambiar, sustituir, reemplazo, desmontar y montar, reinstalar.\n"
            "   - 'Lubrication' = Engrasar, lubricar, mantenimiento.\n"
            "   - 'Polishing' / 'Polish' = Pulir, pulido, abrillantar.\n\n"
            "REGLAS DE SALIDA:\n"
            "1. Devuelve los resultados organizados en una lista Markdown limpia, clara y fácil de leer en la tablet del taller.\n"
            "2. Si el componente solicitado no tiene ninguna relación con lo que hay en el extracto inferior, saca el mensaje oficial de derivación al formulario.\n"
            "3. Queda totalmente prohibido inventar códigos de referencia.\n\n"
            f"--- EXTRACTO DE AMPLIO ESPECTRO DEL CATÁLOGO --- \n{resumen_excel}"
        )

        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=[f"Consulta del operario de taller: '{consulta_usuario}'"],
            config=types.GenerateContentConfig(
                system_instruction=prompt_sistema,
                temperature=0.1
            )
        )
        
        # Sincronización de tokens del sidebar
        if "tokens_totales_input" not in st.session_state:
            st.session_state.tokens_totales_input = 0
            st.session_state.tokens_totales_output = 0
            st.session_state.dinero_total_gastado = 0.0
            st.session_state.ultima_consulta_info = "Ninguna consulta."

        if response.text and response.usage_metadata:
            t_input = response.usage_metadata.prompt_token_count
            t_output = response.usage_metadata.candidates_token_count
            coste = ((t_input * 0.075) / 1_000_000) + ((t_output * 0.30) / 1_000_000)
            
            st.session_state.tokens_totales_input += t_input
            st.session_state.tokens_totales_output += t_output
            st.session_state.dinero_total_gastado += coste
            st.session_state.ultima_consulta_info = f"Última: In: {t_input} | Out: {t_output} (+{coste:.5f}$)"
            
        return response.text if response.text else "❌ No se encontraron coincidencias."
    except Exception as e:
        return f"❌ Error en el motor de la IA de Gemini: {str(e)}"
        # 🧠 3. EL PROMPT MAESTRO (Instrucciones ultra-flexibles y resolución semántica para Gemini)
        prompt_sistema = (
            "Eres el Buscador Inteligente Avanzado del catálogo oficial de OMODA & JAECOO España.\n\n"
            "COMPORTAMIENTO ANTE CONSULTAS DE TALLER:\n"
            "- Ten en cuenta que los usuarios son mecánicos, asesores o jefes de taller que escriben rápido. "
            "Pueden cometer errores, omitir palabras o usar términos en español ('cinturón delantero derecho') "
            "mientras que el catálogo de operaciones inferior está redactado en inglés.\n\n"
            "DICCIONARIO DE TRADUCCIÓN Y MAPEO DE CAMPOS (TRADUCE MENTALMENTE):\n"
            "1. POSICIONES DE PIEZAS:\n"
            "   - 'FR' = Front = Delantero / Delantera / Frontal\n"
            "   - 'RR' = Rear = Trasero / Trasera / Posterior\n"
            "   - 'LH' = Left Hand = Izquierdo / Izquierda / Lado Conductor (en España)\n"
            "   - 'RH' = Right Hand = Derecho / Derecha / Lado Acompañante\n"
            "2. ACCIONES RECURRENTES:\n"
            "   - 'Remove and reinstall' = Desmontar y montar, sustituir, cambiar, reemplazo, reinstalar, sustitución.\n"
            "   - 'Replace' = Cambiar, sustituir, reemplazar de forma directa.\n"
            "   - 'Polishing' / 'Polish' = Pulir, pulido, abrillantar, quitar arañazo superficial.\n"
            "   - 'Lubrication' = Engrasar, lubricar, limpieza y engrase, mantenimiento preventivo.\n"
            "3. COMPATIBILIDAD ENTRE MOTORIZACIONES:\n"
            "   - Si el usuario busca una operación para un modelo eléctrico (BEV) o híbrido (HEV) pero la fila exacta de ese modelo "
            "     tiene el código en blanco o no aparece, busca en el listado la misma operación en el modelo base térmico o gasolina. "
            "     Si la chapa o el componente es idéntico, muéstrale ese código e infórmale de que es una operación compartida.\n"
            "   - Si pide una pieza imposible para ese motor (ej: tubo de vapores/cánister en un eléctrico), facilítale el código de la versión híbrida o gasolina "
            "     y acláraselo de forma muy directa y breve.\n\n"
            "⚠️ REGLA CRÍTICA ESPECIAL (TIEMPOS ADICIONALES):\n"
            "Si detectas que piden meter horas a mano, tiempos adicionales o manuales, muéstrales la operación 'Universal Work Item' "
            "e indícales la nota literal de Central sobre la excepción del filtro 'Spain OJ'.\n\n"
            "REGLAS DE SALIDA:\n"
            "1. Devuelve los resultados válidos en una lista Markdown limpia y estructurada.\n"
            "2. Si la operación solicitada no guarda ninguna relación semántica con lo que hay en el extracto inferior, responde textualmente:\n"
            "   '❌ No se ha encontrado esta operación en el catálogo oficial de la marca. Por favor, dirígete a la pestaña **📝 Solicitar Operación** en el menú lateral izquierdo para rellenar el formulario de solicitud y que Central pueda darla de alta.'\n"
            "3. Prohibido inventarse códigos o nombres bajo ningún concepto.\n\n"
            f"--- EXTRACTO RELEVANTE DEL CATÁLOGO --- \n{resumen_excel}"
        )

        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=[f"Consulta del operario de taller: '{consulta_usuario}'"],
            config=types.GenerateContentConfig(
                system_instruction=prompt_sistema,
                temperature=0.1
            )
        )
        
        # Métricas de consumo del sidebar
        if "tokens_totales_input" not in st.session_state:
            st.session_state.tokens_totales_input = 0
            st.session_state.tokens_totales_output = 0
            st.session_state.dinero_total_gastado = 0.0
            st.session_state.ultima_consulta_info = "Ninguna consulta."

        if response.text and response.usage_metadata:
            t_input = response.usage_metadata.prompt_token_count
            t_output = response.usage_metadata.candidates_token_count
            coste = ((t_input * 0.075) / 1_000_000) + ((t_output * 0.30) / 1_000_000)
            
            st.session_state.tokens_totales_input += t_input
            st.session_state.tokens_totales_output += t_output
            st.session_state.dinero_total_gastado += coste
            st.session_state.ultima_consulta_info = f"Última: In: {t_input} | Out: {t_output} (+{coste:.5f}$)"
            
        return response.text if response.text else "❌ No se encontraron coincidencias."
    except Exception as e:
        return f"❌ Error en el motor de la IA de Gemini: {str(e)}"
        
# =========================================================================
# FUNCIÓN DEL CONSULTORIO DE IA (BLINDADA CONTRA ERRORES DE SESSION STATE)
# =========================================================================
def consultar_ia_garantias(descripcion_averia, archivo_imagen=None):
    """
    Procesa la consulta técnica aplicando un criterio estricto de Central.
    Fuerza respuestas ultra-estructuradas, con frases cortas, veredictos
    inmediatos y respaldo explícito en los artículos de la política.
    """
    try:
        if "GEMINI_API_KEY" not in st.secrets:
            return "⚠️ **Error de Configuración**: No se ha encontrado la clave 'GEMINI_API_KEY' en st.secrets."
            
        client = genai.Client(api_key=st.secrets["GEMINI_API_KEY"])

        try:
            with open("Politica_conocimiento.txt", "r", encoding="utf-8") as f:
                politica_texto = f.read()
        except FileNotFoundError:
            politica_texto = "Política oficial no disponible localmente. Exigir cumplimiento normativo general."

        # PROMPT DE SISTEMA: Establece el tono imperativo, directo y la obligación de citar la política
        prompt_sistema = (
            "Eres un Ingeniero de Garantías Senior para OMODA & JAECOO España. Your task is to issue definitive rulings.\n\n"
            "REGLAS CRÍTICAS DE ESTILO Y FORMATO:\n"
            "1. Frases cortas, cortantes y directas. Evita la paja y los párrafos densos. Usa Markdown exhaustivo (negritas y listas).\n"
            "2. Prohibido incluir cualquier tipo de saludo, introducción o transición. Empieza DIRECTAMENTE con el bloque del Dictamen Preliminar.\n"
            "3. Cada argumento técnico o decisión técnica DEBE citar obligatoriamente la Sección, Artículo o Punto exacto de la política adjunta.\n\n"
            f"--- POLÍTICA DE CONOCIMIENTO OFICIAL ---\n{politica_texto}"
        )

        contenidos = []
        
        # Procesamiento y compresión efímera de imágenes (máximo 2)
        if archivo_imagen is not None:
            lista_archivos = archivo_imagen if isinstance(archivo_imagen, list) else [archivo_imagen]
            for archivo in lista_archivos[:2]:
                imagen_pil = Image.open(io.BytesIO(archivo.read() if hasattr(archivo, 'read') else archivo))
                imagen_pil.thumbnail((1024, 1024))
                contenidos.append(imagen_pil)
            
       # PROMPT DE USUARIO: Criterio laxo para actualizaciones de software integrado
        prompt_usuario = (
            f"Caso reportado por el taller:\n'{descripcion_averia}'\n\n"
            "Genera el dictamen técnico estructurado. No incluyas introducciones. "
            "Usa frases muy cortas. Sigue estrictamente este orden y pautas:\n\n"
            "**📢 VEREDICTO INMEDIATO Y DICTAMEN DE COBERTURA**\n"
            "- Indica en la primera línea si el caso se **ACEPTA**, se **RECHAZA** o si requiere **PRE-AUTORIZACIÓN** (Cita Sección y Artículo).\n"
            "- Argumenta la decisión basándote en la política (aplica la regla de que daños ocultos bajo guarnecidos/consolas en transporte o que lleguen así a puerto SÍ se cubren).\n\n"
            "**1. EVALUACIÓN Y CATEGORÍA TÉCNICA**\n"
            "- **Componente afectado**: Identifícalo en negrita.\n"
            "- **Criticidad**: Evalúala (🔴 Crítico / 🟡 Medio / 🟢 Bajo).\n"
            "- **Naturaleza**: Tipifica el fallo (mecánico, eléctrico, estético, software) con frases de una sola línea.\n\n"
            "**2. ANÁLISIS DE LA EVIDENCIA VISUAL (FOTOS)**\n"
            "- Si NO hay imágenes: Muestra una tabla Markdown detallando las fotos o capturas exactas que debe subir el taller. "
            "⚠️ REGLA DE LAXITUD PARA SOFTWARE: Si el caso es una actualización de software, NO exijas vídeo. Especifica en la tabla que es suficiente con aportar dos fotos: una de la versión de software anterior y otra de la nueva versión ya instalada en el vehículo.\n"
            "- Si SÍ hay imágenes: Describe en puntos breves los hallazgos técnicos (marcas, versiones de pantalla, etc.).\n\n"
            "**3. ACCIÓN REQUERIDA Y PROTOCOLO DE TRABAJO**\n"
            "- Lista numerada (1, 2, 3...) muy escueta con las instrucciones técnicas exactas que debe ejecutar el operario para resolver o certificar la incidencia."
        )
        contenidos.append(prompt_usuario)

        # Ejecución con temperatura baja (0.3) para garantizar precisión y evitar divagaciones
        response = client.models.generate_content(
            model='gemini-3.1-flash-lite',
            contents=contenidos,
            config=types.GenerateContentConfig(
                system_instruction=prompt_sistema,
                temperature=0.3
            )
        )
        
        # 🛡️ INTERCEPCIÓN Y SEGURO DE INICIALIZACIÓN (Evita el fallo que te ha saltado)
        if "tokens_totales_input" not in st.session_state:
            st.session_state.tokens_totales_input = 0
        if "tokens_totales_output" not in st.session_state:
            st.session_state.tokens_totales_output = 0
        if "dinero_total_gastado" not in st.session_state:
            st.session_state.dinero_total_gastado = 0.0
        if "ultima_consulta_info" not in st.session_state:
            st.session_state.ultima_consulta_info = "Ninguna consulta en esta sesión."

        # Inyección segura de métricas tras verificar que existen
        if response.text and response.usage_metadata:
            t_input = response.usage_metadata.prompt_token_count
            t_output = response.usage_metadata.candidates_token_count
            coste_estimado = ((t_input * 0.075) / 1_000_000) + ((t_output * 0.30) / 1_000_000)
            
            st.session_state.tokens_totales_input += t_input
            st.session_state.tokens_totales_output += t_output
            st.session_state.dinero_total_gastado += coste_estimado
            st.session_state.ultima_consulta_info = f"Última: In: {t_input} | Out: {t_output} (+{coste_estimado:.5f}$)"
            
        return response.text if response.text else "⚠️ La IA procesó la solicitud pero no devolvió contenido."

    except Exception as e:
        return f"❌ **Error en la API de Gemini**:\n```text\n{str(e)}\n```"
# ==========================================
# 4. SISTEMA DE SEGURIDAD CONTRASEÑA
# ==========================================
def check_password():
    if not st.session_state.authenticated:
        st.title(txt["pass_titulo"])
        # Añadimos key única al password input también para blindar el acceso contra duplicados
        password = st.text_input(txt["pass_input"], type="password", key="pass_input_unico")
        if st.button(txt["pass_boton"], key="pass_btn_unico"):
            if password == "DealersOJ2026":
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error(txt["error_pass"] if "error_pass" in txt else txt.get("pass_error", "❌ Contraseña incorrecta"))
        return False
    return True

if check_password():
    
  # =========================================================================
    # PANTALLA 1: TIEMPOS DE TALLER (CON ESTADO PERSISTENTE DE IA)
    # =========================================================================
    if opcion_menu == txt["menu_taller"]:
        
        @st.cache_data
        def load_data_tiempos_v3():
            df = pd.read_excel(URL_GITHUB_EXCEL, sheet_name="new_srv_workhours")
            df.columns = df.columns.astype(str).str.strip()
            
            df = df.rename(columns={
                'new_productmodel_idname': 'Modelo',
                'new_product_idname': 'Nombre de la Pieza',
                'new_code': 'Código de Referencia',
                'new_name': 'Operación Técnica',
                'new_standardhour': 'Tiempo Estándar (UT/Horas)',
                'new_remark': 'Notas / Exclusiones',
                'Organization': 'Mercado / Organización',
                'statecodename': 'Estado'
            })

            @st.cache_data
        def load_data_tiempos_v3():
            df = pd.read_excel(URL_GITHUB_EXCEL, sheet_name="new_srv_workhours")
            df.columns = df.columns.astype(str).str.strip()
            
            df = df.rename(columns={
                'new_productmodel_idname': 'Modelo',
                'new_product_idname': 'Nombre de la Pieza',
                'new_code': 'Código de Referencia',
                'new_name': 'Operación Técnica',
                'new_standardhour': 'Tiempo Estándar (UT/Horas)',
                'new_remark': 'Notas / Exclusiones',
                'Organization': 'Mercado / Organización',
                'statecodename': 'Estado'
            })
            
            # 🧹 LIMPIEZA DEL ERROR BINARIO (0x2a / asteriscos basura):
            if 'Mercado / Organización' in df.columns:
                df['Mercado / Organización'] = df['Mercado / Organización'].astype(str)
                # Reemplaza cualquier '0x2a' u otro código hexadecimal por 'Todos'
                df['Mercado / Organización'] = df['Mercado / Organización'].replace(
                    to_replace=r'^0x.*$', value='Todos', regex=True
                )
            
            columnas_finales = [
                'Modelo', 'Nombre de la Pieza', 'Código de Referencia', 
                'Operación Técnica', 'Tiempo Estándar (UT/Horas)', 'Notas / Exclusiones',
                'Mercado / Organización', 'Estado'
            ]
            
            df = df.fillna("")
            df = df.replace(["nan", "None"], "")
            
            columnas_presentes = [col for col in columnas_finales if col in df.columns]
            return df[columnas_presentes].reset_index(drop=True)
            
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
            
            st.title(txt["taller_titulo"])
            st.write(txt["taller_sub"])
            st.markdown("---")

            # 🛡️ INICIALIZADOR DEL ESTADO DE BÚSQUEDA (Evita que el resultado desaparezca)
            if "resultado_ia_excel" not in st.session_state:
                st.session_state.resultado_ia_excel = None

            # =================================================================
            # 🤖 SECCIÓN A: ASISTENTE IA DE BÚSQUEDA BILINGÜE
            # =================================================================
            st.subheader("🤖 Buscar operación")
            st.write("Escribe tu consulta en español. La IA traducirá los términos mecánicos y buscará en las columnas en inglés.")
            
            consulta_rapida = st.text_input(
                "¿Qué operación, pieza o modelo necesitas localizar?",
                placeholder="Ejemplo: cambiar pastillas de freno delanteras del omoda 5 / desmontar paragolpes jaecoo 7...",
                key="campo_consulta_ia_excel"
            )

            st.warning("""
            ⚠️ **RECORDATORIO** Antes de tramitar cualquier reclamación, verifique obligatoriamente que **la pieza a reclamar coincide con el pedido exacto realizado a Recambios** para esta reparación. 
            """)

            if st.button("Buscar operación", type="secondary", width='stretch'):
                if not consulta_rapida.strip():
                    st.warning("⚠️ Introduce una descripción o término para realizar la búsqueda.")
                else:
                    with st.spinner("🔍 Traduciendo y escaneando el catálogo de operaciones..."):
                        # Guardamos el resultado en la sesión persistente
                        st.session_state.resultado_ia_excel = buscador_inteligente_excel(consulta_rapida, data)
                        st.rerun()

            # RENDERIZADO ESTÁTICO DEL RESULTADO: Si hay algo guardado, se pinta sí o sí
            if st.session_state.resultado_ia_excel:
                st.markdown("#### ⚙️ Resultado de la Consulta:")
                if "❌ No se ha encontrado" in st.session_state.resultado_ia_excel:
                    st.error(st.session_state.resultado_ia_excel)
                else:
                    st.info(st.session_state.resultado_ia_excel)
                
                # Botón auxiliar opcional para limpiar la pantalla de la IA
                if st.button("🗑️ Limpiar búsqueda de la IA", key="btn_limpiar_ia"):
                    st.session_state.resultado_ia_excel = None
                    st.rerun()

            st.markdown("---")

            # =================================================================
            # 📊 SECCIÓN B: FILTROS Y TABLA TRADICIONAL (BÚSQUEDA MANUAL)
            # =================================================================
            st.subheader("📊 Catálogo Completo (Filtros Manuales)")
            
            col1, col2, col3 = st.columns([1, 1.5, 1.5])
            with col1:
                # 1. Extraemos todos los modelos reales en texto limpio
                modelos_raw = [str(m).strip() for m in data['Modelo'].dropna().unique()]
    
                # 2. Lista blanca oficial para limpiar la paja del Excel
                modelos_oficiales = [
                       "JAECOO 7 PHEV", "JAECOO 8 PHEV", 
                       "OMODA 5", "OMODA 5 EV", "OMODA 5 HEV", 
                       "OMODA 7 PHEV", "OMODA 9 PHEV"
                ]
    
                # 3. Filtramos: Nos quedamos con los oficiales + cualquier variante de LEPAS que venga en el Excel
                
                modelos_filtrados = [
                    m for m in modelos_raw 
                    if any(marca in m.upper() for marca in ["OMODA", "JAECOO", "LEPAS L8 PHEV"])
                ]

    
                # 4. Creamos la lista final con el botón "Todos" de tu diccionario de idiomas
                modelos_disponibles = [txt["todos"]] + sorted(list(set(modelos_filtrados)))
    
                # 5. Tu selectbox actual
                modelo_seleccionado = st.selectbox(txt["f_modelo"], modelos_disponibles)
            with col2:
                buscar_pieza = st.text_input(txt["f_pieza"], "").strip()
            with col3:
                buscar_operacion = st.text_input(txt["f_operacion"], "").strip()

            col_m, col_e = st.columns([2, 2])
            
            with col_m:
                if 'Mercado / Organización' in data.columns:
                    mercados_disponibles = [txt["todos"]] + [str(m).strip() for m in data['Mercado / Organización'].unique() if str(m).strip() != ""]
                    
                    indice_defecto = 0
                    for idx, m in enumerate(mercados_disponibles):
                        if "spain" in m.lower() or "oj spain" in m.lower():
                            indice_defecto = idx
                            break
                    
                    market_label = txt["f_mercado_taller"]
                    mercado_seleccionado = st.selectbox(market_label, mercados_disponibles, index=indice_defecto)
                else:
                    mercado_seleccionado = txt["todos"]
                    
            with col_e:
                if 'Estado' in data.columns:
                    estados_disponibles = [txt["todos"]] + [str(e).strip() for e in data['Estado'].unique() if str(e).strip() != ""]
                    indice_est_defecto = estados_disponibles.index("Active") if "Active" in estados_disponibles else 0
                    estado_seleccionado = st.selectbox(txt["f_estado_taller"], estados_disponibles, index=indice_est_defecto)
                else:
                    estado_seleccionado = txt["todos"]

            df_filtrado = data.copy()
            
            if modelo_seleccionado != txt["todos"]:
                df_filtrado = df_filtrado[df_filtrado['Modelo'] == modelo_seleccionado]
                
            if mercado_seleccionado != txt["todos"] and 'Mercado / Organización' in df_filtrado.columns:
                df_filtrado = df_filtrado[df_filtrado['Mercado / Organización'].astype(str).str.strip() == mercado_seleccionado]
                
            if estado_seleccionado != txt["todos"] and 'Estado' in df_filtrado.columns:
                df_filtrado = df_filtrado[df_filtrado['Estado'].astype(str).str.strip() == estado_seleccionado]

            if buscar_pieza:
                df_filtrado = df_filtrado[
                    df_filtrado['Nombre de la Pieza'].astype(str).str.contains(buscar_pieza, case=False, na=False) |
                    df_filtrado['Código de Referencia'].astype(str).str.contains(buscar_pieza, case=False, na=False)
                ]
                
            if buscar_operacion:
                df_filtrado = df_filtrado[df_filtrado['Operación Técnica'].astype(str).str.contains(buscar_operacion, case=False, na=False)]

            st.markdown(txt["res_taller"].format(len(df_filtrado)))
            if not df_filtrado.empty:
                st.dataframe(df_filtrado, width='stretch', hide_index=True)
            else:
                st.warning(txt["warn_taller"])
                
        except Exception as e:
            st.error(txt["err_taller"].format(e))

    # =========================================================================
    # PANTALLA 3: SOLICITUD DE OPERACIONES ADICIONALES (CONEXIÓN GOOGLE SHEETS)
    # =========================================================================
    elif opcion_menu == txt["menu_solicitar"]:
        st.title(txt["solicitar_titulo"])
        st.write(txt["solicitar_sub"])
        st.markdown("---")

        LISTA_DEALERS = sorted([
            "ACAI MOTOR MÁLAGA", "ALIFAVISA BILBAO", "ALIMOTOR ELCHE", "ANFERPA SEGOVIA", 
            "AUTO YALDE LOGROÑO", "AUTOCAM MOTOR VILAFRANCA", "AUTOCYL PALENCIA", "AUTOCYL VALLADOLID", 
            "AUTOTERMINAL", "AUTOVIDAL PALMA DE MALLORCA", "AXIS MOTORS", "BLENDIO LAREDO", 
            "BLENDIO LUGO", "BLENDIO OURENSE", "BLENDIO SANTANDER", "BLENDIO TORRELAVEGA", 
            "BORJAMOTOR ALICANTE", "CERVERA AVILA", "CERVERA SALAMANCA", "CHINARES GUADALAJARA", 
            "DUMOSA BENAVENTE", "ESLAUTO LEON", "GRUP BASOLS IGUALADA", "GRUPO JULIAN BURGOS", 
            "GRUPO NIETO MÁLAGA", "GRUPO NIETO MARBELLA", "HIMASA SEDAVÍ", "JEMOYA SORIA", 
            "LASACAR MIRANDA DE EBRO", "LASACAR VITORIA", "M TECNIK ALCALÁ DE HENARES", 
            "M TECNIK BARCELONA MAQUINIST", "M TECNIK CASTELLÓN", "M TECNIK GERONA", 
            "M TECNIK MATARÓ", "M TECNIK VINAROZ", "MARTIN LIZAGA", "MARTIN LIZAGA TERUEL", 
            "MAS AUTO LEGANÉS", "MAVEN BADAJOZ", "MAVEN CÁCERES", "MOLL MOTOR DENIA", 
            "MOLL MOTOR GANDIA", "MONECAR AUTOMOCION", "MONECAR CUENCA", "MOTOR NACIENTE", 
            "MOVINSUR GRANADA", "MOVINSUR JAÉN", "MOVINSUR MOTRIL", "MY CARS CÓRDOBA", 
            "NOATUM", "NOVACAR BCN SANT BOI", "PALAUSA ZAMORA", "PRUNA CAR GO GRANOLLERS", 
            "PROCHERY ALBACETE", "PROCHERY CARTAGENA", "PROCHERY MURCIA", "RAFAEL AFONSO AGUIMES", 
            "RAFAEL AFONSO LANZAROTE", "RAFAEL AFONSO LAS PALMAS", "RAFAEL AFONSO TENERIFE", 
            "RESNOVA MOTOR CORUÑA", "RESNOVA MOTOR GIJÓN", "RESNOVA MOTOR NARÓN", 
            "RESNOVA MOTOR OVIEDO", "RESNOVA MOTOR SANTIAGO", "RESNOVA MOTOR VIGO", 
            "SEGRE MOTORS LERIDA", "SERTECAUTO PONFERRADA", "SYRSA ALGECIRAS", 
            "SYRSA ALMERIA", "SYRSA EJIDO", "SYRSA HUELVA", "SYRSA SEVILLA", 
            "TALAUTO CAZALEGAS", "TALAUTO TOLEDO", "TALLERES CHINARES", "TECNOTARRACO TARRAGONA", 
            "TERRY MOBILITY JERÉZ", "TRADECAR GAMBOA ALCORCÓN", "TRADECAR GAMBOA MADRID", 
            "TRADECAR GAMBOA MAJADAHON", "TRADECAR GAMBOA RIVAS", "TUMASA HUESCA", 
            "TUMASA MONZÓN", "UNIONE ALCAZAR DE SAN JUAN", "UNIONE CIUDAD REAL", 
            "VALLESCAR SABADELL", "VALLESCAR TERRASSA", "VIAN AUTOMOBILE VILLALBA", 
            "ZEN MOTOR OLABERRIA", "ZEN MOTOR PAMPLONA", "ZEN MOTOR SAN SEBASTIÁN", 
            "ZEN MOTOR ZARAGOZA"
        ])
        
        MAPEO_MODELOS = {
            "OMODA 5 (Gasolina)": "T19C", "OMODA 5 HEV (Híbrido)": "T19C HEV", "OMODA 5 EV (Eléctrico)": "T19C EV",
            "OMODA 7 PHEV": "T1GC PHEV", "OMODA 9 PHEV": "T22 PHEV", "JAECOO 5 (Gasolina)": "T13J",
            "JAECOO 5 HEV": "T13J HEV", "JAECOO 5 BEV": "T13J BEV", "JAECOO 7 (Gasolina)": "T1EJ",
            "JAECOO 7 HEV": "T1EJ HEV", "JAECOO 7 PHEV": "T1EJ PHEV", "JAECOO 8 PHEV": "T26", "LEPAS L8 PHEV": "T1G PHEV"
        }
        
        st.subheader(txt["form_sub"])
        col1, col2 = st.columns(2)
        with col1:
            marca = st.selectbox(txt["form_marca"], ["OMODA", "JAECOO", "LEPAS"])
            modelos_filtrados = [mod for mod in MAPEO_MODELOS.keys() if mod.upper().startswith(marca.upper())]
            modelo_comercial = st.selectbox(txt["form_modelo"], modelos_filtrados)
        with col2:
            dealer = st.selectbox(txt["form_dealer"], LISTA_DEALERS)
            codigo_producto_auto = MAPEO_MODELOS[modelo_comercial]
            st.text_input(txt["form_hq_code"], value=codigo_producto_auto, disabled=True)
        
        # Formulario estructurado
        with st.form("hq_operation_form", clear_on_submit=True):
            c1, c2 = st.columns(2)
            with c1:
                vin = st.text_input(txt["form_vin"], max_chars=17, placeholder=txt["form_vin_holder"]).strip().upper()
            with c2:
                referencia = st.text_input(txt["form_ref"], placeholder=txt["form_ref_holder"]).strip().upper()
            
            operacion_solicitada = st.text_area(txt["form_op"], placeholder=txt["form_op_holder"]).strip()
            boton_enviar = st.form_submit_button(txt["form_btn"])
            
            if boton_enviar:
                if not vin or not operacion_solicitada:
                    st.error(txt["err_campos"])
                elif len(vin) != 17:
                    st.error("❌ Error: El número de bastidor (VIN) debe tener exactamente 17 caracteres.")
                else:
                    ahora = datetime.datetime.now()
                    columnas_orden = [
                        "SN", "Submitted on", "Respondents", "Fecha del día", 
                        "Marca del vehículo", "INTRODUCIR MODELO", "INTRODUCIR VIN", 
                        "Mercado", "CÓDIGO DE PRODUCTO", "REFERENCIA DE PIEZA", 
                        "OPERACIÓN QUE SE SOLICITA AÑADIR", "DEALER"
                    ]
                    
                    subida_exitosa = False
                    
                    # 1. Intentar conectar y leer el estado actual de la nube
                    try:
                        from streamlit_gsheets import GSheetsConnection
                        conn = st.connection("gsheets", type=GSheetsConnection)
                        
                        if "connections" in st.secrets and "gsheets" in st.secrets["connections"]:
                            spreadsheet_url = st.secrets["connections"]["gsheets"]["spreadsheet"]
                        elif "gsheets" in st.secrets and "spreadsheet" in st.secrets["gsheets"]:
                            spreadsheet_url = st.secrets["gsheets"]["spreadsheet"]
                        else:
                            spreadsheet_url = st.secrets.get("spreadsheet", "")

                        df_cloud = conn.read(spreadsheet=spreadsheet_url)
                        
                        if df_cloud.empty or len(df_cloud.columns) < 2:
                            df_cloud = pd.DataFrame(columns=columnas_orden)
                        else:
                            df_cloud = df_cloud.dropna(how='all').loc[:, ~df_cloud.columns.str.contains('^Unnamed')]
                        
                        nuevo_sn = len(df_cloud) + 1
                    except Exception:
                        nuevo_sn = len(st.session_state.lista_solicitudes) + 1
                        df_cloud = pd.DataFrame(columns=columnas_orden)
                        spreadsheet_url = ""
                    
                    # 2. Estructurar la nueva fila
                    nueva_solicitud = {
                        "SN": int(nuevo_sn),
                        "Submitted on": str(ahora.strftime("%Y-%m-%d %H:%M:%S")),
                        "Respondents": str(f"Dealer App ({dealer})"),
                        "Fecha del día": str(ahora.strftime("%Y-%m-%d")),
                        "Marca del vehículo": str(marca),
                        "INTRODUCIR MODELO": str(modelo_comercial),
                        "INTRODUCIR VIN": str(vin),
                        "Mercado": "Spain OJ",
                        "CÓDIGO DE PRODUCTO": str(codigo_producto_auto),
                        "REFERENCIA DE PIEZA": str(referencia) if referencia else "NaN",
                        "OPERACIÓN QUE SE SOLICITA AÑADIR": str(operacion_solicitada),
                        "DEALER": str(dealer)
                    }
                    
                    # 3. Intentar subir los datos a Google Sheets
                    try:
                        df_nuevo = pd.DataFrame([nueva_solicitud])
                        df_nuevo = df_nuevo.reindex(columns=columnas_orden)
                        df_cloud = df_cloud.reindex(columns=columnas_orden)
                        
                        df_actualizado = pd.concat([df_cloud, df_nuevo], ignore_index=True)
                        
                        if spreadsheet_url:
                            conn.update(
                                spreadsheet=spreadsheet_url,
                                data=df_actualizado
                            )
                            st.session_state.lista_solicitudes.append(nueva_solicitud)
                            subida_exitosa = True
                        else:
                            raise ValueError("No se encontró la URL del archivo de Sheets en st.secrets.")
                            
                    except Exception as e:
                        st.error(f"❌ Error de conexión con Google Sheets: {e}")
                        st.info("💡 Por seguridad, hemos guardado esta línea en la tabla inferior (Caché Local).")
                        st.session_state.lista_solicitudes.append(nueva_solicitud)

                        if subida_exitosa:
                            st.success("✅ **Operación registrada con éxito.** La solicitud ha sido transmitida al departamento de Garantías de Central para su validación.")
                            import time
                            time.sleep(1.5)
                            st.rerun()
                        
# =========================================================================
    # PANTALLA 4: CONSULTORIO IA DE GARANTÍAS (VERSION DEFINITIVA)
    # =========================================================================
    elif opcion_menu == "🧠 Consultorio Técnico IA":
        st.title("🤖 Consultor Técnico de Garantías (Inteligencia Artificial)")
        st.write("Analiza de forma preliminar si una avería está cubierta según el manual de políticas oficial e identifica los pasos técnicos a seguir.")
        st.markdown("---")

        # Inicializamos la variable en la sesión para que el informe no se borre al refrescar
        if "resultado_consultorio" not in st.session_state:
            st.session_state.resultado_consultorio = None

        st.subheader("📝 Detalles de la Consulta")
        descripcion_averia = st.text_area(
            "Descripción de la avería o síntomas del vehículo:",
            placeholder="Ejemplo: Cliente reporta ruido metálico al girar el volante a la izquierda en OMODA 5...",
            height=150
        )
        
        # Cargador múltiple blindado a un máximo de 2 fotos y 20MB por archivo
        archivos_imagenes = st.file_uploader(
            "📸 Adjuntar evidencias o fotos de la avería (Máximo 2 imágenes - 20MB máx por archivo):", 
            type=["jpg", "jpeg", "png"],
            accept_multiple_files=True,
            key="cargador_imagenes_taller"
        )
        
        archivos_validos = []
        peso_correcto = True
        
        if archivos_imagenes:
            if len(archivos_imagenes) > 2:
                st.error("❌ **Error**: El sistema solo acepta un máximo de 2 imágenes por consulta.")
                peso_correcto = False
            else:
                for archivo in archivos_imagenes:
                    if archivo.size > 20 * 1024 * 1024:
                        st.error(f"❌ **El archivo '{archivo.name}' supera el límite permitido de 20 MB.**")
                        peso_correcto = False
                if peso_correcto:
                    archivos_validos = archivos_imagenes
        
        # Botón de envío
        if st.button("🔍 Enviar Consulta a la IA", type="primary", use_container_width=True):
            if not descripcion_averia.strip():
                st.error("⚠️ Por favor, introduce una descripción de la avería antes de realizar la consulta.")
            elif archivos_imagenes and len(archivos_imagenes) > 2:
                st.error("❌ Corrige la cantidad de imágenes antes de continuar.")
            elif archivos_imagenes and not peso_correcto:
                st.error("❌ Una o más imágenes superan los 20 MB.")
            else:
                with st.spinner("🧠 Analizando la documentación oficial y generando el informe técnico..."):
                    parametro_imagenes = archivos_validos if archivos_validos else None
                    # Guardamos el resultado en el estado de la sesión
                    st.session_state.resultado_consultorio = consultar_ia_garantias(descripcion_averia, parametro_imagenes)

        # RENDERIZADO PERSISTENTE (Fuera del botón para que no desaparezca nada)
        if st.session_state.resultado_consultorio:
            st.markdown("### 📋 Informe de Diagnóstico Generado")
            st.markdown(st.session_state.resultado_consultorio)
            st.success("✅ Análisis preliminar finalizado.")
            
            st.markdown("<br>", unsafe_allow_html=True)
            
            # 💛 EL DISCLAIMER EN AMARILLO FIJO Y RESALTADO ABAJO DEL TODO
            st.warning("""
            #### ⚠️ NOTA OBLIGATORIA DE CENTRAL
            Este informe constituye una **valoración preliminar e informativa** basada exclusivamente en los síntomas y evidencias gráficas aportadas por el taller. 
            
            Para validar definitivamente el diagnóstico técnico, proceder con la autorización de la reparación bajo garantía o reportar de forma oficial un fallo de fabricación de origen, **es obligatorio abrir un canal oficial en la plataforma aportando el bastidor (VIN) completo**:
            
            * 🛠️ **¿Dudas sobre el diagnóstico técnico o el proceso de reparación?** Abra un **Ticket de Asistencia Técnica** en el sistema o envíe un correo detallado a: [soportetecnico@omodaes.com](mailto:soportetecnico@omodaes.com)
            * 📝 **¿Consultas sobre los plazos de cobertura de garantía o tramitación?** Contacte de forma directa con el departamento administrativo en: [garantias@omodaes.com](mailto:garantias@omodaes.com)
            """)
