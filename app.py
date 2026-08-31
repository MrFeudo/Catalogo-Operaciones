# -*- coding: utf-8 -*-

import csv
import io
import json
import re
import time
import unicodedata
import datetime
from pathlib import Path

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from PIL import Image

try:
    from google import genai
    from google.genai import types
except Exception:
    genai = None
    types = None


# =========================================================================
# CONFIGURACIÓN GENERAL
# =========================================================================
st.set_page_config(page_title="Buscador Técnico OMODA & JAECOO", layout="wide")

URL_GITHUB_EXCEL = "https://github.com/MrFeudo/Catalogo-Operaciones/raw/main/DMS_Active_Spare_Parts.xlsb"

BASE_DIR = Path(__file__).resolve().parent
STATS_FILE = BASE_DIR / "usage_stats.json"
LOG_FILE = BASE_DIR / "warranty_comments_log.csv"
DETAIL_LOG_FILE = BASE_DIR / "warranty_comments_log_detail.csv"

LOG_FIELDNAMES = [
    "log_id", "timestamp", "date", "time", "claim_number",
    "reason_ids", "reason_labels", "reason_categories",
    "base_comment", "final_comment", "was_edited"
]

DETAIL_LOG_FIELDNAMES = [
    "log_id", "timestamp", "date", "time", "claim_number",
    "reason_id", "reason_label", "reason_category",
    "base_comment", "final_comment", "was_edited"
]


# =========================================================================
# INICIALIZACIÓN SESSION STATE
# =========================================================================
DEFAULT_SESSION_VALUES = {
    "lista_solicitudes": [],
    "authenticated": False,
    "idioma": "Español",
    "resultado_ia_excel": None,
    "resultado_consultorio": None,
    "selected_keys": [],
    "previous_selected_keys": [],
    "claim_val": "",
    "final_comment_area": "",
    "last_saved_comment": "",
    "pending_clear_comment_area": False,
    "pending_clear_selection": False,
    "confirm_missing_claim": False,
    "tokens_totales_input": 0,
    "tokens_totales_output": 0,
    "dinero_total_gastado": 0.0,
    "ultima_consulta_info": "Ninguna consulta.",
}

for key, value in DEFAULT_SESSION_VALUES.items():
    if key not in st.session_state:
        st.session_state[key] = value


# =========================================================================
# DICCIONARIO DE TRADUCCIÓN
# =========================================================================
IDIOMAS = {
    "Español": {
        "menu_titulo": "### 🗺️ Menú de Navegación",
        "menu_radio": "Selecciona una herramienta:",
        "menu_taller": "📋 Tiempos de Taller",
        "menu_generador": "📊 Estadísticas de garantías devueltas",
        "menu_solicitar": "📝 Solicitar Operación",
        "menu_consultorio": "🧠 Consultorio Técnico IA",
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
        "todos": "Todos",
        "solicitar_titulo": "📝 Solicitud de Operaciones Adicionales de Mano de Obra",
        "solicitar_sub": "Utilice este formulario para solicitar el alta de nuevas operaciones o precios en el maestro de HQ.",
        "form_sub": "Datos de la Solicitud (Campos obligatorios *)",
        "form_marca": "Marca del vehículo *",
        "form_modelo": "INTRODUCIR MODELO *",
        "form_vin": "INTRODUCIR VIN (Bastidor) *",
        "form_vin_holder": "17 caracteres",
        "form_hq_code": "CÓDIGO DE PRODUCTO (Asignado por HQ)",
        "form_ref": "REFERENCIA DE PIEZA (Opcional)",
        "form_ref_holder": "Ej. 7365747465AA",
        "form_op": "OPERACIÓN QUE SE SOLICITA AÑADIR *",
        "form_op_holder": "Describa detalladamente la operación técnica o falta de precio que requiere el taller...",
        "form_btn": "Enviar Solicitud a Central",
        "err_campos": "❌ Por favor, rellene todos los campos obligatorios (*).",
    },
    "English": {
        "menu_titulo": "### 🗺️ Navigation Menu",
        "menu_radio": "Select a tool:",
        "menu_taller": "📋 Workshop Times",
        "menu_generador": "📊 Returned Warranty Statistics",
        "menu_solicitar": "📝 Request Operation",
        "menu_consultorio": "🧠 Technical AI Consultant",
        "pass_titulo": "🔐 Dealer Network Access",
        "pass_input": "Enter access password:",
        "pass_boton": "Login",
        "pass_error": "❌ Incorrect password",
        "taller_titulo": "🚗 Labor Operations Catalog",
        "taller_sub": "Consult parts, models and assigned times directly from the DMS.",
        "f_modelo": "1. Filter by Model:",
        "f_pieza": "2. Search by Part Name or Code:",
        "f_operacion": "3. Search by operation type:",
        "f_mercado_taller": "Filter by Market / Organization:",
        "f_estado_taller": "Filter by Operation Status:",
        "res_taller": "### 📋 Results found: {} operations",
        "warn_taller": "⚠️ No operations found matching the selected criteria.",
        "err_taller": "Error processing workshop times database: {}",
        "todos": "All",
        "solicitar_titulo": "📝 Request for Additional Labor Operations",
        "solicitar_sub": "Use this form to request new operations or prices to be added to HQ master list.",
        "form_sub": "Request Details (* Required fields)",
        "form_marca": "Vehicle Brand *",
        "form_modelo": "ENTER MODEL *",
        "form_vin": "ENTER VIN (Chassis) *",
        "form_vin_holder": "17 characters",
        "form_hq_code": "PRODUCT CODE (Assigned by HQ)",
        "form_ref": "PART REFERENCE (Optional)",
        "form_ref_holder": "e.g., 7365747465AA",
        "form_op": "OPERATION REQUESTED TO BE ADDED *",
        "form_op_holder": "Describe in detail the technical operation or missing price required by the workshop...",
        "form_btn": "Send Request to HQ",
        "err_campos": "❌ Please fill in all required fields (*).",
    },
    "Chinese (中文)": {
        "menu_titulo": "### 🗺️ 导航菜单",
        "menu_radio": "选择工具:",
        "menu_taller": "📋 车间工时",
        "menu_generador": "📊 退回保修统计",
        "menu_solicitar": "📝 请求操作",
        "menu_consultorio": "🧠 技术 AI 咨询",
        "pass_titulo": "🔐 经销商网络访问",
        "pass_input": "输入访问密码:",
        "pass_boton": "登录",
        "pass_error": "❌ 密码错误",
        "taller_titulo": "🚗 工时操作目录",
        "taller_sub": "直接从 DMS 查询零件、车型和分配的时间。",
        "f_modelo": "1. 按车型筛选:",
        "f_pieza": "2. 按零件名称或代码搜索:",
        "f_operacion": "3. 按操作类型搜索:",
        "f_mercado_taller": "按市场 / 组织筛选:",
        "f_estado_taller": "按操作状态筛选:",
        "res_taller": "### 📋 找到的结果: {} 个操作",
        "warn_taller": "⚠️ 未找到符合选择条件的工时操作。",
        "err_taller": "处理车间工时数据库时出错: {}",
        "todos": "全部",
        "solicitar_titulo": "📝 申请新增工时操作",
        "solicitar_sub": "使用此表单申请在总部主数据中添加新工时操作或价格。",
        "form_sub": "申请信息 (* 为必填项)",
        "form_marca": "车辆品牌 *",
        "form_modelo": "输入车型 *",
        "form_vin": "输入 VIN (车架号) *",
        "form_vin_holder": "17位字符",
        "form_hq_code": "产品代码 (由总部分配)",
        "form_ref": "零件编号 (选填)",
        "form_ref_holder": "例如: 7365747465AA",
        "form_op": "申请添加的操作内容 *",
        "form_op_holder": "请详细描述车间所需的工时操作或缺失的价格...",
        "form_btn": "发送申请至总部",
        "err_campos": "❌ 请填写所有必填项 (*)。",
    }
}


# =========================================================================
# GENERADOR DE COMENTARIOS - DATOS
# =========================================================================
COMMENTS = {
    "1": {"category": "Costes, mano de obra y piezas", "label": "Mano de obra adicional no justificada", "text": "No se justifica la mano de obra adicional. Adjuntar fichajes, desglosar y justificar el tiempo extra o ajustar la mano de obra adicional a 0."},
    "2": {"category": "Tipo de reclamación / Cobertura", "label": "Tipo de garantía incorrecto", "text": "Cambiar el tipo de garantía a PDI."},
    "3": {"category": "Costes, mano de obra y piezas", "label": "Tiempo adicional no aceptado", "text": "No se acepta el tiempo de reclamación adicional en esta operación."},
    "4": {"category": "Costes, mano de obra y piezas", "label": "Referencia incorrecta", "text": "La referencia reclamada es incorrecta."},
    "5": {"category": "Evidencias / documentación", "label": "Adjuntar evidencias", "text": "Adjuntar evidencias."},
    "6": {"category": "Evidencias / documentación", "label": "Evidencia de diagnóstico y reparación", "text": "Adjuntar evidencia del proceso de diagnóstico y reparación."},
    "7": {"category": "Evidencias / documentación", "label": "Evidencia pieza sustituida y nueva", "text": "Adjuntar evidencias de la pieza sustituida y la nueva."},
    "8": {"category": "Evidencias / documentación", "label": "Ticket técnico no adjuntado / sin resumen en la reclamación", "text": "Adjuntar ticket técnico en el apartado correspondiente. Siempre que haya ticket, se deben resumir también las indicaciones recibidas y las pruebas realizadas para que la reclamación y la solución final adoptada puedan entenderse correctamente."},
    "9": {"category": "Operaciones frecuentes", "label": "Elegir operación de actualización / refresh", "text": "Elegir la operación de actualización correspondiente (refresh o software update)."},
    "10": {"category": "Operaciones frecuentes", "label": "Elegir operación de pulido / polish", "text": "Elegir la operación de pulido (polish)."},
    "11": {"category": "Operaciones frecuentes", "label": "Elegir operación de pintado / paint", "text": "Elegir la operación de pintado correspondiente (paint)."},
    "12": {"category": "Tipo de reclamación / Cobertura", "label": "Reclamar como actualización técnica", "text": "Reclamar como actualización técnica."},
    "13": {"category": "Costes, mano de obra y piezas", "label": "Coste auxiliar en operaciones subcontratadas", "text": "Pon el coste auxiliar en el apartado de costes de operaciones subcontratadas."},
    "14": {"category": "Costes, mano de obra y piezas", "label": "Cantidad de pieza a 0", "text": "Poner cantidad de pieza a 0."},
    "15": {"category": "Costes, mano de obra y piezas", "label": "Operación externa: coste íntegro + presupuesto", "text": "Si es una operación externa, poner el coste íntegro en el apartado de costes de operaciones subcontratadas y adjuntar presupuesto."},
    "16": {"category": "Información y campos", "label": "Información en campos incorrectos", "text": "Respetar la función de cada campo. El diagnóstico y la solución deben indicarse en sus apartados correspondientes. Los comentarios adicionales deben usarse solo para aclaraciones sobre la evidencia o respuestas a preguntas directas, y la descripción de otros costes solo para desglosar los costes adicionales reclamados."},
    "17": {"category": "Información y campos", "label": "Información insuficiente en descripción/diagnóstico", "text": "Información insuficiente. Los campos de descripción y/o diagnóstico no contienen una explicación suficientemente detallada para entender la reclamación. Ampliar la información indicando claramente el síntoma, el diagnóstico realizado, la causa de la avería y la solución aplicada. Si lo hubiera, se debe adjuntar también el ticket técnico."},
    "18": {"category": "Tipo de reclamación / Cobertura", "label": "No cubierto por garantía", "text": "Esto no parece corresponder a un defecto de producto cubierto por garantía."}
}

CATEGORY_ORDER = [
    "Evidencias / documentación",
    "Operaciones frecuentes",
    "Tipo de reclamación / Cobertura",
    "Costes, mano de obra y piezas",
    "Información y campos"
]

TOP_RED_LIMIT = 3
TOP_AMBER_LIMIT = 5

CATEGORY_COLOR_MAP = {
    "Evidencias / documentación": "#ff7f0e",
    "Operaciones frecuentes": "#2ca02c",
    "Tipo de reclamación / Cobertura": "#9467bd",
    "Costes, mano de obra y piezas": "#1f77b4",
    "Información y campos": "#d62728",
    "Comentario manual": "#7f7f7f",
}


# =========================================================================
# UTILIDADES
# =========================================================================
def normalizar_texto(texto):
    texto = str(texto)
    return ''.join(
        c for c in unicodedata.normalize('NFD', texto)
        if unicodedata.category(c) != 'Mn'
    ).lower()


def normalize_text(text):
    return normalizar_texto(text)


def ensure_token_state():
    for key, value in {
        "tokens_totales_input": 0,
        "tokens_totales_output": 0,
        "dinero_total_gastado": 0.0,
        "ultima_consulta_info": "Ninguna consulta.",
    }.items():
        if key not in st.session_state:
            st.session_state[key] = value


def register_gemini_usage(response):
    ensure_token_state()
    if getattr(response, "text", None) and getattr(response, "usage_metadata", None):
        t_input = response.usage_metadata.prompt_token_count
        t_output = response.usage_metadata.candidates_token_count
        coste = ((t_input * 0.075) / 1_000_000) + ((t_output * 0.30) / 1_000_000)
        st.session_state.tokens_totales_input += t_input
        st.session_state.tokens_totales_output += t_output
        st.session_state.dinero_total_gastado += coste
        st.session_state.ultima_consulta_info = f"Última: In: {t_input} | Out: {t_output} (+{coste:.5f}$)"


# =========================================================================
# GENERADOR DE COMENTARIOS - LÓGICA
# =========================================================================
def load_usage_stats():
    if not STATS_FILE.exists():
        return {}
    try:
        with open(STATS_FILE, "r", encoding="utf-8") as file:
            return json.load(file)
    except Exception:
        return {}


def save_usage_stats(stats):
    with open(STATS_FILE, "w", encoding="utf-8") as file:
        json.dump(stats, file, ensure_ascii=False, indent=4)


def update_usage_stats(selected_keys):
    stats = load_usage_stats()
    for key in selected_keys:
        stats[key] = stats.get(key, 0) + 1
    save_usage_stats(stats)


def generate_log_id(now):
    return now.strftime("%Y%m%d_%H%M%S_%f")


def was_comment_edited(base_comment, final_comment):
    return " ".join(str(base_comment).split()) != " ".join(str(final_comment).split())


def get_categories_for_keys(selected_keys):
    categories = []
    for category in CATEGORY_ORDER:
        for key in selected_keys:
            if key in COMMENTS and COMMENTS[key]["category"] == category and category not in categories:
                categories.append(category)
    return categories


def migrate_csv_if_needed(path, fieldnames):
    if not path.exists() or path.stat().st_size == 0:
        return
    try:
        with open(path, "r", newline="", encoding="utf-8-sig") as file:
            reader = csv.DictReader(file)
            old_fields = reader.fieldnames or []
            rows = list(reader)
    except Exception:
        return

    if old_fields == fieldnames:
        return

    migrated = []
    for row in rows:
        migrated.append({field: row.get(field, "") for field in fieldnames})

    with open(path, "w", newline="", encoding="utf-8-sig") as file:
        writer = csv.DictWriter(file, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(migrated)


def log_generated_comment(selected_keys, final_comment, base_comment, claim_number):
    now = datetime.datetime.now()
    log_id = generate_log_id(now)
    claim_str = str(claim_number or "").strip().upper() or "NO INFORMADO"
    edited = "YES" if was_comment_edited(base_comment, final_comment) else "NO"

    selected_labels = [COMMENTS[k]["label"] for k in selected_keys if k in COMMENTS]
    categories = get_categories_for_keys(selected_keys)
    cat_text = " | ".join(categories) if categories else "Comentario manual"

    migrate_csv_if_needed(LOG_FILE, LOG_FIELDNAMES)
    migrate_csv_if_needed(DETAIL_LOG_FILE, DETAIL_LOG_FIELDNAMES)

    summary_row = {
        "log_id": log_id,
        "timestamp": now.isoformat(timespec="seconds"),
        "date": now.strftime("%Y-%m-%d"),
        "time": now.strftime("%H:%M:%S"),
        "claim_number": claim_str,
        "reason_ids": ", ".join(selected_keys) if selected_keys else "MANUAL",
        "reason_labels": " | ".join(selected_labels) if selected_labels else "Comentario manual",
        "reason_categories": cat_text,
        "base_comment": base_comment,
        "final_comment": final_comment,
        "was_edited": edited,
    }

    file_exists = LOG_FILE.exists() and LOG_FILE.stat().st_size > 0
    with open(LOG_FILE, "a", newline="", encoding="utf-8-sig") as file:
        writer = csv.DictWriter(file, fieldnames=LOG_FIELDNAMES)
        if not file_exists:
            writer.writeheader()
        writer.writerow(summary_row)

    detail_exists = DETAIL_LOG_FILE.exists() and DETAIL_LOG_FILE.stat().st_size > 0
    with open(DETAIL_LOG_FILE, "a", newline="", encoding="utf-8-sig") as file:
        writer = csv.DictWriter(file, fieldnames=DETAIL_LOG_FIELDNAMES)
        if not detail_exists:
            writer.writeheader()

        if selected_keys:
            for key in selected_keys:
                if key in COMMENTS:
                    writer.writerow({
                        "log_id": log_id,
                        "timestamp": now.isoformat(timespec="seconds"),
                        "date": now.strftime("%Y-%m-%d"),
                        "time": now.strftime("%H:%M:%S"),
                        "claim_number": claim_str,
                        "reason_id": key,
                        "reason_label": COMMENTS[key]["label"],
                        "reason_category": COMMENTS[key]["category"],
                        "base_comment": base_comment,
                        "final_comment": final_comment,
                        "was_edited": edited,
                    })
        else:
            writer.writerow({
                "log_id": log_id,
                "timestamp": now.isoformat(timespec="seconds"),
                "date": now.strftime("%Y-%m-%d"),
                "time": now.strftime("%H:%M:%S"),
                "claim_number": claim_str,
                "reason_id": "MANUAL",
                "reason_label": "Comentario manual",
                "reason_category": "Comentario manual",
                "base_comment": base_comment,
                "final_comment": final_comment,
                "was_edited": edited,
            })


def get_usage_rank_map(usage_stats):
    used_counts = sorted(
        {usage_stats.get(k, 0) for k in COMMENTS if usage_stats.get(k, 0) > 0},
        reverse=True
    )
    count_rank_map = {count: idx + 1 for idx, count in enumerate(used_counts)}
    return {
        k: count_rank_map[usage_stats.get(k, 0)]
        for k in COMMENTS
        if usage_stats.get(k, 0) > 0
    }


def get_stats_dataframe(usage_stats):
    rank_map = get_usage_rank_map(usage_stats)
    total_uses = sum(v for v in usage_stats.values() if isinstance(v, int))
    rows = []
    for key in COMMENTS:
        uses = usage_stats.get(key, 0)
        if uses <= 0:
            continue
        rows.append({
            "ID": key,
            "TOP": f"TOP {rank_map.get(key, '-')}",
            "Usos": uses,
            "%": (uses / total_uses * 100) if total_uses else 0,
            "Categoría": COMMENTS[key]["category"],
            "Motivo": COMMENTS[key]["label"],
        })
    return pd.DataFrame(rows).sort_values(by=["Usos", "ID"], ascending=[False, True]) if rows else pd.DataFrame()


def get_category_stats_dataframe(usage_stats):
    rows = []
    total = sum(v for v in usage_stats.values() if isinstance(v, int))
    for category in CATEGORY_ORDER:
        uses = sum(usage_stats.get(k, 0) for k, item in COMMENTS.items() if item["category"] == category)
        if uses > 0:
            rows.append({
                "Categoría": category,
                "Usos": uses,
                "%": (uses / total * 100) if total else 0,
            })
    return pd.DataFrame(rows).sort_values(by="Usos", ascending=False) if rows else pd.DataFrame()


# =========================================================================
# GEMINI - BUSCADOR DE OPERACIONES
# =========================================================================
def build_semantic_map():
    return {
        "cambiar": "remove and reinstall|replace|remove|reinstall",
        "cambio": "remove and reinstall|replace|remove|reinstall",
        "sustituir": "remove and reinstall|replace|remove|reinstall",
        "sustitucion": "remove and reinstall|replace|remove|reinstall",
        "reemplazar": "remove and reinstall|replace|remove|reinstall",
        "desmontar": "remove",
        "montar": "reinstall",
        "comprobar": "check|inspection|test|diagnostic|measurement",
        "verificar": "check|inspection|test|diagnostic",
        "diagnostico": "check|inspection|test|diagnostic",
        "actualizar": "refresh|update|software|flash",
        "programar": "refresh|update|software|flash|coding|program",
        "calibrar": "calibrate|calibration",
        "pulir": "polishing|polish",
        "pulido": "polishing|polish",
        "bateria": "battery|storage battery|bms|tecu",
        "centralita": "control unit|control module|ecu|bcm|mcu|vcu|tcu|hcu",
        "modulo": "control module|module",
        "camara": "camera|fcm|avm|rear view",
        "radar": "radar|frm|bsd",
        "sensor": "sensor|probe|detector",
        "airbag": "airbag|air bag|abm|srs",
        "cinturon": "seatbelt|seat belt|belt",
        "motor": "engine assy|motor|engine",
        "turbo": "turbocharger|turbo",
        "radiador": "radiator",
        "bomba": "pump|water pump|oil pump|fuel pump",
        "dct": "dct|dual clutch transmission",
        "cambio": "transmission|gearbox|dct|gearshift|remove and reinstall|replace",
        "caja": "transmission|gearbox",
        "embrague": "clutch",
        "palier": "drive shaft|axle shaft|half shaft",
        "freno": "brake|ipb|epb|abs",
        "pastilla": "pads|brake pads",
        "disco": "disc|brake disc",
        "amortiguador": "shock absorber|strut|damper",
        "trapecio": "control arm|suspension arm|wishbone",
        "direccion": "steering|eps",
        "paragolpes": "bumper",
        "faro": "headlamp|headlight",
        "retrovisor": "mirror|rearview mirror",
        "puerta": "door",
        "porton": "tailgate|back door|rear door",
        "techo": "sunroof|roof|panoramic roof",
        "cristal": "glass|window",
        "asiento": "seat",
        "soporte": "bracket|support|mount|holder",
        "cuna": "subframe|cradle|bracket|salver|tray",
        "tapa": "cover|cap|lid",
        "filtro": "filter",
        "aceite": "oil|lubricant",
        "refrigerante": "coolant",
        "tubo": "pipe|tube|hose",
        "manguito": "hose",
        "delantero": "fr",
        "delantera": "fr",
        "trasero": "rr",
        "trasera": "rr",
        "izquierdo": "lh",
        "izquierda": "lh",
        "derecho": "rh",
        "derecha": "rh",
    }


def filter_catalog_for_ai(consulta_usuario, df_contexto):
    consulta_limpia = normalizar_texto(consulta_usuario.strip())
    mapa_raices = build_semantic_map()

    abreviaturas_modelos = {
        "j5": "jaecoo 5", "jaecoo5": "jaecoo 5", "j-5": "jaecoo 5",
        "j7": "jaecoo 7", "jaecoo7": "jaecoo 7", "j-7": "jaecoo 7",
        "j8": "jaecoo 8", "jaecoo8": "jaecoo 8",
        "o5": "omoda 5", "omoda5": "omoda 5", "o-5": "omoda 5",
        "hibrido": "hev", "electrico": "bev", "gasolina": "ice",
    }

    for abrev, mod_real in abreviaturas_modelos.items():
        if abrev in consulta_limpia.split() or abrev in consulta_limpia:
            consulta_limpia = consulta_limpia.replace(abrev, mod_real)

    lista_palabras_usuario = consulta_limpia.split()

    palabras_regex = []
    for esp, eng in mapa_raices.items():
        if esp in consulta_limpia:
            palabras_regex.extend(eng.split("|"))

    for palabra in lista_palabras_usuario:
        if len(palabra) > 2 and palabra not in ["quiero", "para", "con", "del", "una", "uno", "los", "las", "este", "de"]:
            palabras_regex.append(palabra)

    palabras_regex = list(set(palabras_regex))

    df_base = df_contexto.copy()
    for col in ["Modelo", "Nombre de la Pieza", "Operación Técnica"]:
        if col in df_base.columns:
            df_base[col] = df_base[col].astype(str).str.lower().str.strip()

    if "omoda" in consulta_limpia and "Modelo" in df_base.columns:
        df_base = df_base[df_base["Modelo"].str.contains("omoda", na=False)]
    elif "jaecoo" in consulta_limpia and "Modelo" in df_base.columns:
        df_base = df_base[df_base["Modelo"].str.contains("jaecoo", na=False)]

    componentes_encontrados = []
    for esp, eng in mapa_raices.items():
        if esp in consulta_limpia and esp not in ["cambiar", "sustituir", "cambio", "sustitucion", "reemplazar", "desmontar", "montar"]:
            componentes_encontrados.extend(eng.split("|"))

    if componentes_encontrados:
        regex_comp = "|".join(set(componentes_encontrados))
        mask = pd.Series(False, index=df_base.index)
        for col in ["Nombre de la Pieza", "Operación Técnica"]:
            if col in df_base.columns:
                mask = mask | df_base[col].str.contains(regex_comp, na=False)
        df_base = df_base[mask]

    terminos_manuales = ["manual", "adicional", "extra", "tiempo mas", "añadir horas", "universal", "baremo"]
    if any(term in consulta_limpia for term in terminos_manuales) and "Operación Técnica" in df_contexto.columns:
        df_base = df_contexto[
            df_contexto["Operación Técnica"].astype(str).str.lower().str.contains("universal", na=False)
        ]

    if palabras_regex and not df_base.empty:
        regex_puntos = "|".join(palabras_regex)
        df_base["score"] = 0
        for col, weight in [("Modelo", 5), ("Nombre de la Pieza", 10), ("Operación Técnica", 10)]:
            if col in df_base.columns:
                df_base["score"] += df_base[col].astype(str).str.contains(regex_puntos, na=False).astype(int) * weight
        df_base = df_base.sort_values(by="score", ascending=False).drop(columns=["score"], errors="ignore").head(100)
    else:
        df_base = df_base.head(60)

    if df_base.empty:
        df_base = df_contexto.head(60)

    wanted_cols = [
        "Modelo", "Nombre de la Pieza", "Código de Referencia",
        "Operación Técnica", "Tiempo Estándar (UT/Horas)", "Notas / Exclusiones"
    ]
    present_cols = [c for c in wanted_cols if c in df_base.columns]
    return df_base[present_cols].head(100)


def buscador_inteligente_excel(consulta_usuario, df_contexto):
    try:
        if genai is None or types is None:
            return "⚠️ **Error**: No se ha podido importar el SDK de Gemini."
        if "GEMINI_API_KEY" not in st.secrets:
            return "⚠️ **Error**: No se ha encontrado la clave API en st.secrets."

        client = genai.Client(api_key=st.secrets["GEMINI_API_KEY"])
        df_recortado = filter_catalog_for_ai(consulta_usuario, df_contexto)
        resumen_excel = df_recortado.to_string(index=False)

        prompt_sistema = (
            "Eres el Buscador Inteligente Avanzado del catálogo oficial de OMODA & JAECOO España.\n\n"
            "MISIÓN:\n"
            "- El usuario es personal de taller. Puede escribir rápido, con términos en español, errores o abreviaturas.\n"
            "- El catálogo está en inglés. Debes traducir mentalmente componentes, acciones y posiciones.\n"
            "- Muestra TODAS las operaciones válidas que encuentres en el extracto relacionadas con el componente solicitado.\n\n"
            "REGLA DE TIEMPOS:\n"
            "- La columna 'Tiempo Estándar (UT/Horas)' contiene UTs.\n"
            "- 100 UTs = 1 hora = 60 minutos.\n"
            "- Si muestras una operación con tiempo, incluye: **Tiempo:** X UTs (~Horas: Y hr | ~Minutos: Z min).\n\n"
            "GUÍA RÁPIDA:\n"
            "- FR = Front / delantero. RR = Rear / trasero. LH = izquierdo. RH = derecho.\n"
            "- Remove and reinstall / Replace = cambiar, sustituir, desmontar y montar.\n"
            "- Polishing / Polish = pulido. Refresh / update / software = actualización.\n"
            "- Si piden tiempos adicionales/manuales, busca y muestra Universal Work Item si está en el extracto.\n\n"
            "REGLAS DE SALIDA:\n"
            "1. Devuelve una lista Markdown limpia y fácil de leer.\n"
            "2. Prohibido inventar códigos, piezas o tiempos.\n"
            "3. Si no hay relación semántica suficiente, responde exactamente:\n"
            "'❌ No se ha encontrado esta operación en el catálogo oficial de la marca. Por favor, dirígete a la pestaña **📝 Solicitar Operación** en el menú lateral izquierdo para rellenar el formulario de solicitud y que Central pueda darla de alta.'\n\n"
            f"--- EXTRACTO DEL CATÁLOGO ---\n{resumen_excel}"
        )

        response = client.models.generate_content(
            model="gemini-2.5-flash",
            contents=[f"Consulta del operario de taller: '{consulta_usuario}'"],
            config=types.GenerateContentConfig(
                system_instruction=prompt_sistema,
                temperature=0.1
            )
        )

        register_gemini_usage(response)
        return response.text if response.text else "❌ No se encontraron coincidencias."

    except Exception as exc:
        return f"❌ Error en el motor de la IA de Gemini: {str(exc)}"


# =========================================================================
# CONSULTORIO IA GARANTÍAS
# =========================================================================
def consultar_ia_garantias(descripcion_averia, archivo_imagen=None):
    try:
        if genai is None or types is None:
            return "⚠️ **Error**: No se ha podido importar el SDK de Gemini."
        if "GEMINI_API_KEY" not in st.secrets:
            return "⚠️ **Error de Configuración**: No se ha encontrado la clave 'GEMINI_API_KEY' en st.secrets."

        client = genai.Client(api_key=st.secrets["GEMINI_API_KEY"])

        try:
            with open(BASE_DIR / "Politica_conocimiento.txt", "r", encoding="utf-8") as file:
                politica_texto = file.read()
        except FileNotFoundError:
            politica_texto = "Política oficial no disponible localmente. Exigir cumplimiento normativo general."

        prompt_sistema = (
            "Eres un Ingeniero de Garantías Senior para OMODA & JAECOO España. Your task is to issue definitive rulings.\n\n"
            "REGLAS CRÍTICAS DE ESTILO Y FORMATO:\n"
            "1. Frases cortas, directas y sin paja. Usa Markdown, negritas y listas.\n"
            "2. Prohibido incluir saludos, introducciones o transiciones. Empieza directamente con el dictamen.\n"
            "3. Cada argumento técnico o decisión debe citar la Sección, Artículo o Punto exacto de la política adjunta cuando esté disponible.\n\n"
            f"--- POLÍTICA DE CONOCIMIENTO OFICIAL ---\n{politica_texto}"
        )

        contenidos = []

        if archivo_imagen is not None:
            lista_archivos = archivo_imagen if isinstance(archivo_imagen, list) else [archivo_imagen]
            for archivo in lista_archivos[:2]:
                raw = archivo.read() if hasattr(archivo, "read") else archivo
                imagen_pil = Image.open(io.BytesIO(raw))
                imagen_pil.thumbnail((1024, 1024))
                contenidos.append(imagen_pil)

        prompt_usuario = (
            f"Caso reportado por el taller:\n'{descripcion_averia}'\n\n"
            "Genera el dictamen técnico estructurado. No incluyas introducciones. "
            "Usa frases muy cortas. Sigue estrictamente este orden:\n\n"
            "**📢 VEREDICTO INMEDIATO Y DICTAMEN DE COBERTURA**\n"
            "- Indica si el caso se **ACEPTA**, se **RECHAZA** o requiere **PRE-AUTORIZACIÓN**.\n"
            "- Argumenta la decisión según política.\n\n"
            "**1. EVALUACIÓN Y CATEGORÍA TÉCNICA**\n"
            "- **Componente afectado**.\n"
            "- **Criticidad**: 🔴 Crítico / 🟡 Medio / 🟢 Bajo.\n"
            "- **Naturaleza**: mecánico, eléctrico, estético, software, etc.\n\n"
            "**2. ANÁLISIS DE LA EVIDENCIA VISUAL (FOTOS)**\n"
            "- Si no hay imágenes, indica exactamente qué fotos o capturas debe subir el taller.\n"
            "- Para software, no exijas vídeo: bastan fotos de versión anterior y nueva versión instalada.\n\n"
            "**3. ACCIÓN REQUERIDA Y PROTOCOLO DE TRABAJO**\n"
            "- Lista numerada muy escueta con instrucciones técnicas."
        )
        contenidos.append(prompt_usuario)

        response = client.models.generate_content(
            model="gemini-2.5-flash",
            contents=contenidos,
            config=types.GenerateContentConfig(
                system_instruction=prompt_sistema,
                temperature=0.3
            )
        )

        register_gemini_usage(response)
        return response.text if response.text else "⚠️ La IA procesó la solicitud pero no devolvió contenido."

    except Exception as exc:
        return f"❌ **Error en la API de Gemini**:\n```text\n{str(exc)}\n```"


# =========================================================================
# AUTENTICACIÓN Y SIDEBAR
# =========================================================================
def render_sidebar_and_get_option():
    try:
        st.sidebar.image("logo_empresa.png", use_container_width=True)
    except Exception:
        st.sidebar.write("🏢 **OMODA & JAECOO**")

    st.sidebar.markdown("---")

    idioma_seleccionado = st.sidebar.selectbox(
        "🌐 Language / Idioma / 语言:",
        ["Español", "English", "Chinese (中文)"],
        index=["Español", "English", "Chinese (中文)"].index(st.session_state.idioma),
        key="selector_idioma_global"
    )
    st.session_state.idioma = idioma_seleccionado
    txt_local = IDIOMAS[st.session_state.idioma]

    st.sidebar.markdown("---")
    st.sidebar.markdown(txt_local["menu_titulo"])

    opciones = [
        txt_local["menu_taller"],
        txt_local["menu_generador"],
        txt_local["menu_solicitar"],
        txt_local["menu_consultorio"],
    ]

    opcion = st.sidebar.radio(
        txt_local["menu_radio"],
        opciones,
        key="menu_navegacion_app"
    )

    st.sidebar.markdown("---")
    st.sidebar.caption(st.session_state.ultima_consulta_info)
    st.sidebar.caption(
        f"Tokens In: {st.session_state.tokens_totales_input} · "
        f"Out: {st.session_state.tokens_totales_output} · "
        f"Coste: {st.session_state.dinero_total_gastado:.5f}$"
    )

    return txt_local, opcion


def check_password(txt_local):
    if not st.session_state.authenticated:
        st.title(txt_local["pass_titulo"])
        password = st.text_input(txt_local["pass_input"], type="password", key="pass_input_unico")
        if st.button(txt_local["pass_boton"], key="pass_btn_unico"):
            if password == "DealersOJ2026":
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error(txt_local["pass_error"])
        return False
    return True


# =========================================================================
# PANTALLA 1 - TIEMPOS DE TALLER
# =========================================================================
@st.cache_data
def load_data_tiempos_v3():
    df = pd.read_excel(URL_GITHUB_EXCEL, sheet_name="new_srv_workhours")
    df.columns = df.columns.astype(str).str.strip()

    mapeo_columnas = {
        "new_productmodel_idname": "Modelo",
        "new_product_idname": "Nombre de la Pieza",
        "new_code": "Código de Referencia",
        "new_name": "Operación Técnica",
        "new_standardhour": "Tiempo Estándar (UT/Horas)",
        "new_remark": "Notas / Exclusiones",
        "Organization": "Mercado / Organización",
        "statecodename": "Estado",
    }

    cols_existentes = [col for col in mapeo_columnas if col in df.columns]
    df_limpio = df[cols_existentes].copy().rename(columns=mapeo_columnas)
    df_limpio = df_limpio.replace(to_replace=r"^0x.*$", value="", regex=True)
    df_limpio = df_limpio.fillna("")
    df_limpio = df_limpio.replace(["nan", "None", "NaN"], "")

    columnas_finales = [
        "Modelo", "Nombre de la Pieza", "Código de Referencia",
        "Operación Técnica", "Tiempo Estándar (UT/Horas)", "Notas / Exclusiones",
        "Mercado / Organización", "Estado",
    ]
    columnas_presentes = [col for col in columnas_finales if col in df_limpio.columns]
    return df_limpio[columnas_presentes].reset_index(drop=True)


def render_tiempos_taller(txt_local):
    try:
        data = load_data_tiempos_v3()

        st.title(txt_local["taller_titulo"])
        st.write(txt_local["taller_sub"])
        st.markdown("---")

        st.subheader("🤖 Buscar operación")
        st.write("Escribe tu consulta en español. La IA traducirá términos mecánicos y buscará en columnas en inglés.")

        consulta_rapida = st.text_input(
            "¿Qué operación, pieza o modelo necesitas localizar?",
            placeholder="Ejemplo: cambiar pastillas de freno delanteras del omoda 5 / desmontar paragolpes jaecoo 7...",
            key="campo_consulta_ia_excel"
        )

        st.warning(
            "⚠️ **RECORDATORIO** Antes de tramitar cualquier reclamación, verifique obligatoriamente "
            "que **la pieza a reclamar coincide con el pedido exacto realizado a Recambios** para esta reparación."
        )

        if st.button("Buscar operación", type="secondary", use_container_width=True):
            if not consulta_rapida.strip():
                st.warning("⚠️ Introduce una descripción o término para realizar la búsqueda.")
            else:
                with st.spinner("🔍 Traduciendo y escaneando el catálogo de operaciones..."):
                    st.session_state.resultado_ia_excel = buscador_inteligente_excel(consulta_rapida, data)
                    st.rerun()

        if st.session_state.resultado_ia_excel:
            st.markdown("#### ⚙️ Resultado de la Consulta:")
            if "❌ No se ha encontrado" in st.session_state.resultado_ia_excel:
                st.error(st.session_state.resultado_ia_excel)
            else:
                st.info(st.session_state.resultado_ia_excel)

            if st.button("🗑️ Limpiar búsqueda de la IA", key="btn_limpiar_ia"):
                st.session_state.resultado_ia_excel = None
                st.rerun()

        st.markdown("---")
        st.subheader("📊 Catálogo Completo (Filtros Manuales)")

        col1, col2, col3 = st.columns([1, 1.5, 1.5])

        with col1:
            modelos_raw = [str(m).strip() for m in data["Modelo"].dropna().unique()] if "Modelo" in data.columns else []
            modelos_filtrados = [
                m for m in modelos_raw
                if any(marca in m.upper() for marca in ["OMODA", "JAECOO", "LEPAS"])
            ]
            modelos_disponibles = [txt_local["todos"]] + sorted(list(set(modelos_filtrados)))
            modelo_seleccionado = st.selectbox(txt_local["f_modelo"], modelos_disponibles)

        with col2:
            buscar_pieza = st.text_input(txt_local["f_pieza"], "").strip()

        with col3:
            buscar_operacion = st.text_input(txt_local["f_operacion"], "").strip()

        col_m, col_e = st.columns([2, 2])

        with col_m:
            if "Mercado / Organización" in data.columns:
                mercados_disponibles = [txt_local["todos"]] + [
                    str(m).strip()
                    for m in data["Mercado / Organización"].unique()
                    if str(m).strip() != ""
                ]
                indice_defecto = 0
                for idx, mercado in enumerate(mercados_disponibles):
                    if "spain" in mercado.lower() or "oj spain" in mercado.lower():
                        indice_defecto = idx
                        break
                mercado_seleccionado = st.selectbox(txt_local["f_mercado_taller"], mercados_disponibles, index=indice_defecto)
            else:
                mercado_seleccionado = txt_local["todos"]

        with col_e:
            if "Estado" in data.columns:
                estados_disponibles = [txt_local["todos"]] + [
                    str(e).strip()
                    for e in data["Estado"].unique()
                    if str(e).strip() != ""
                ]
                indice_est_defecto = estados_disponibles.index("Active") if "Active" in estados_disponibles else 0
                estado_seleccionado = st.selectbox(txt_local["f_estado_taller"], estados_disponibles, index=indice_est_defecto)
            else:
                estado_seleccionado = txt_local["todos"]

        df_filtrado = data.copy()

        if modelo_seleccionado != txt_local["todos"] and "Modelo" in df_filtrado.columns:
            df_filtrado = df_filtrado[df_filtrado["Modelo"] == modelo_seleccionado]

        if mercado_seleccionado != txt_local["todos"] and "Mercado / Organización" in df_filtrado.columns:
            df_filtrado = df_filtrado[df_filtrado["Mercado / Organización"].astype(str).str.strip() == mercado_seleccionado]

        if estado_seleccionado != txt_local["todos"] and "Estado" in df_filtrado.columns:
            df_filtrado = df_filtrado[df_filtrado["Estado"].astype(str).str.strip() == estado_seleccionado]

        if buscar_pieza and {"Nombre de la Pieza", "Código de Referencia"}.issubset(df_filtrado.columns):
            df_filtrado = df_filtrado[
                df_filtrado["Nombre de la Pieza"].astype(str).str.contains(buscar_pieza, case=False, na=False) |
                df_filtrado["Código de Referencia"].astype(str).str.contains(buscar_pieza, case=False, na=False)
            ]

        if buscar_operacion and "Operación Técnica" in df_filtrado.columns:
            df_filtrado = df_filtrado[df_filtrado["Operación Técnica"].astype(str).str.contains(buscar_operacion, case=False, na=False)]

        st.markdown(txt_local["res_taller"].format(len(df_filtrado)))
        if not df_filtrado.empty:
            st.dataframe(df_filtrado, use_container_width=True, hide_index=True)
        else:
            st.warning(txt_local["warn_taller"])

    except Exception as exc:
        st.error(txt_local["err_taller"].format(exc))




def copy_to_system_clipboard(text):
    """
    Copia automática desactivada en Streamlit.

    En una app web, Python corre en el servidor y no debe intentar acceder al
    portapapeles del usuario. La copia real se hace con el botón de navegador
    basado en navigator.clipboard dentro de render_browser_copy_button().
    """
    return False, (
        "La copia automática desde Python está desactivada. "
        "Usa el botón de navegador 'Copiar comentario al portapapeles'."
    )

def render_browser_copy_button(text, button_text="📋 Copiar al portapapeles", key="copy_button"):
    """
    Botón real de copia en navegador para Streamlit.

    Usa navigator.clipboard.writeText. Funciona especialmente bien en localhost
    y en conexiones HTTPS. No guarda logs: solo copia el texto visible.
    """
    safe_text = json.dumps(text or "")
    safe_button_text = json.dumps(button_text)
    safe_key = re.sub(r"[^a-zA-Z0-9_]", "_", str(key))

    components.html(
        f"""
        <div style="display:flex; align-items:center; gap:10px; margin: 2px 0 8px 0;">
            <button id="btn_{safe_key}" type="button" style="
                background-color:#f0f2f6;
                border:1px solid #d0d3da;
                border-radius:6px;
                padding:8px 14px;
                cursor:pointer;
                font-family:Arial, sans-serif;
                font-size:14px;
            "></button>
            <span id="status_{safe_key}" style="font-family:Arial, sans-serif; font-size:13px; color:#2e7d32;"></span>
        </div>
        <script>
            const textToCopy_{safe_key} = {safe_text};
            const btn_{safe_key} = document.getElementById("btn_{safe_key}");
            const status_{safe_key} = document.getElementById("status_{safe_key}");
            btn_{safe_key}.innerText = {safe_button_text};
            btn_{safe_key}.onclick = async function() {{
                if (!textToCopy_{safe_key}) {{
                    status_{safe_key}.style.color = "#b00020";
                    status_{safe_key}.innerText = "No hay comentario para copiar.";
                    return;
                }}
                try {{
                    await navigator.clipboard.writeText(textToCopy_{safe_key});
                    status_{safe_key}.style.color = "#2e7d32";
                    status_{safe_key}.innerText = "Copiado. Pega con Ctrl + V.";
                }} catch (err) {{
                    status_{safe_key}.style.color = "#b00020";
                    status_{safe_key}.innerText = "No se pudo copiar. Selecciona el texto manualmente.";
                }}
            }};
        </script>
        """,
        height=52,
    )



def render_keyboard_shortcuts():
    """
    Atajos de teclado para la pantalla de estadísticas de garantías devueltas.

    Importante: en Streamlit Cloud/Python no se puede copiar al portapapeles
    del usuario desde el servidor. Por eso los atajos copian el texto desde el
    navegador con navigator.clipboard durante el propio evento de teclado y,
    después, hacen click en el botón de guardado correspondiente.
    """
    components.html(
        """
        <script>
        (function () {
            const parentWindow = window.parent;
            const parentDocument = parentWindow.document;

            // Reinstalamos el listener en cada rerun para mantener la última versión.
            if (parentWindow.__ojWarrantyKeyboardShortcutsHandler) {
                parentDocument.removeEventListener(
                    'keydown',
                    parentWindow.__ojWarrantyKeyboardShortcutsHandler,
                    true
                );
            }

            function showShortcutStatus(message, isError=false) {
                let box = parentDocument.getElementById('oj_shortcut_status_box');
                if (!box) {
                    box = parentDocument.createElement('div');
                    box.id = 'oj_shortcut_status_box';
                    box.style.position = 'fixed';
                    box.style.right = '22px';
                    box.style.bottom = '22px';
                    box.style.zIndex = '999999';
                    box.style.padding = '10px 14px';
                    box.style.borderRadius = '8px';
                    box.style.fontFamily = 'Arial, sans-serif';
                    box.style.fontSize = '13px';
                    box.style.boxShadow = '0 2px 10px rgba(0,0,0,0.18)';
                    parentDocument.body.appendChild(box);
                }
                box.style.background = isError ? '#fdecea' : '#e8f5e9';
                box.style.color = isError ? '#b00020' : '#1b5e20';
                box.style.border = isError ? '1px solid #f5c2c7' : '1px solid #b7dfb9';
                box.innerText = message;
                window.clearTimeout(parentWindow.__ojShortcutStatusTimeout);
                parentWindow.__ojShortcutStatusTimeout = window.setTimeout(function () {
                    if (box && box.parentNode) {
                        box.parentNode.removeChild(box);
                    }
                }, 2600);
            }

            function buttonByText(text) {
                const buttons = Array.from(parentDocument.querySelectorAll('button'));
                return buttons.find(function (button) {
                    return (button.innerText || '').includes(text);
                });
            }

            function clickButton(text) {
                const button = buttonByText(text);
                if (button) {
                    button.click();
                    return true;
                }
                showShortcutStatus('No encuentro el botón: ' + text, true);
                return false;
            }

            function focusInputByPlaceholder(partialPlaceholder) {
                const inputs = Array.from(parentDocument.querySelectorAll('input, textarea'));
                const input = inputs.find(function (element) {
                    return ((element.placeholder || '').includes(partialPlaceholder));
                });
                if (input) {
                    input.focus();
                    if (typeof input.select === 'function') {
                        input.select();
                    }
                    return true;
                }
                return false;
            }

            function getCommentTextarea() {
                const textareas = Array.from(parentDocument.querySelectorAll('textarea'));

                // En esta pantalla el comentario editable suele ser el textarea con más contenido.
                // Así evitamos depender de clases internas de Streamlit, que cambian mucho.
                const candidates = textareas
                    .filter(function (textarea) {
                        const value = textarea.value || '';
                        const placeholder = textarea.placeholder || '';
                        const aria = textarea.getAttribute('aria-label') || '';
                        return (
                            value.trim().length > 0 ||
                            placeholder.includes('Revisa') ||
                            aria.includes('Revisa') ||
                            aria.includes('modifica')
                        );
                    })
                    .sort(function (a, b) {
                        return (b.value || '').length - (a.value || '').length;
                    });

                return candidates.length ? candidates[0] : null;
            }

            async function copyCurrentCommentFromBrowser() {
                const textarea = getCommentTextarea();
                const text = textarea ? (textarea.value || '').trim() : '';

                if (!text) {
                    showShortcutStatus('No hay comentario para copiar.', true);
                    return false;
                }

                // 1) Intento moderno. Usamos el navigator de la ventana padre porque
                // Streamlit mete components.html dentro de un iframe y, según navegador
                // o despliegue, el iframe puede no tener permiso de clipboard.
                try {
                    if (parentWindow.navigator && parentWindow.navigator.clipboard) {
                        await parentWindow.navigator.clipboard.writeText(text);
                        showShortcutStatus('Comentario copiado. Pega con Ctrl + V.');
                        return true;
                    }
                } catch (err) {
                    // Pasamos al fallback clásico.
                }

                // 2) Fallback clásico: seleccionar temporalmente el textarea real de
                // Streamlit y ejecutar copy desde el documento padre. Suele funcionar
                // mejor con atajos de teclado porque el evento viene de una acción del usuario.
                try {
                    textarea.focus();
                    textarea.select();
                    textarea.setSelectionRange(0, textarea.value.length);

                    const copied = parentDocument.execCommand('copy');

                    // Quitamos la selección para no dejar la pantalla rara.
                    if (parentWindow.getSelection) {
                        const selection = parentWindow.getSelection();
                        if (selection && selection.removeAllRanges) {
                            selection.removeAllRanges();
                        }
                    }

                    if (copied) {
                        showShortcutStatus('Comentario copiado. Pega con Ctrl + V.');
                        return true;
                    }
                } catch (err) {
                    // Pasamos al fallback invisible.
                }

                // 3) Último fallback: crear un textarea invisible en el documento padre,
                // copiarlo y borrarlo.
                try {
                    const helper = parentDocument.createElement('textarea');
                    helper.value = text;
                    helper.setAttribute('readonly', '');
                    helper.style.position = 'fixed';
                    helper.style.left = '-9999px';
                    helper.style.top = '-9999px';
                    parentDocument.body.appendChild(helper);
                    helper.focus();
                    helper.select();
                    const copied = parentDocument.execCommand('copy');
                    parentDocument.body.removeChild(helper);

                    if (copied) {
                        showShortcutStatus('Comentario copiado. Pega con Ctrl + V.');
                        return true;
                    }
                } catch (err) {
                    // Nada más que probar.
                }

                showShortcutStatus('No se pudo copiar automáticamente. Usa el botón o selecciona el texto.', true);
                return false;
            }

            async function copyThenClick(buttonText) {
                const copied = await copyCurrentCommentFromBrowser();
                if (!copied) {
                    return;
                }
                window.setTimeout(function () {
                    clickButton(buttonText);
                }, 80);
            }

            parentWindow.__ojWarrantyKeyboardShortcutsHandler = function (event) {
                const key = (event.key || '').toLowerCase();
                const target = event.target;
                const tagName = target && target.tagName ? target.tagName.toLowerCase() : '';
                const isTypingField = tagName === 'input' || tagName === 'textarea' || (target && target.isContentEditable);
                const ctrlOrCmd = event.ctrlKey || event.metaKey;

                // Ctrl/Cmd + Shift + Enter: copiar desde navegador, guardar y limpiar.
                // Funciona también dentro del comentario editable.
                if (ctrlOrCmd && event.shiftKey && event.key === 'Enter') {
                    event.preventDefault();
                    copyThenClick('Guardar y limpiar');
                    return;
                }

                // Ctrl/Cmd + Enter: copiar desde navegador y guardar.
                // Funciona también dentro del comentario editable.
                if (ctrlOrCmd && !event.shiftKey && event.key === 'Enter') {
                    event.preventDefault();
                    copyThenClick('Guardar');
                    return;
                }

                // Enter normal: copiar, guardar y limpiar solo si NO estás escribiendo en un campo.
                if (!ctrlOrCmd && !event.shiftKey && !event.altKey && event.key === 'Enter' && !isTypingField) {
                    event.preventDefault();
                    copyThenClick('Guardar y limpiar');
                    return;
                }

                // Alt + N: foco rápido en número de reclamación/garantía.
                if (event.altKey && key === 'n') {
                    event.preventDefault();
                    focusInputByPlaceholder('CO202608310001');
                    return;
                }

                // Alt + B: foco rápido en buscador.
                if (event.altKey && key === 'b') {
                    event.preventDefault();
                    focusInputByPlaceholder('Filtrar por ID');
                    return;
                }

                // Alt + L: limpiar selección.
                if (event.altKey && key === 'l') {
                    event.preventDefault();
                    clickButton('Limpiar selección');
                    return;
                }
            };

            parentDocument.addEventListener(
                'keydown',
                parentWindow.__ojWarrantyKeyboardShortcutsHandler,
                true
            );
        })();
        </script>
        """,
        height=0,
    )

# =========================================================================
# PANTALLA 2 - GENERADOR DE COMENTARIOS
# =========================================================================
def render_generador_comentarios():
    render_keyboard_shortcuts()

    st.title("📊 Estadísticas de garantías devueltas")
    st.caption("Registra los motivos de garantías devueltas, revisa el comentario, cópialo al portapapeles y guárdalo en los CSV.")
    st.caption(
        "Atajos: Enter = copiar + guardar y limpiar cuando no estás escribiendo · "
        "Ctrl + Enter = copiar + guardar · Ctrl + Shift + Enter = copiar + guardar y limpiar · "
        "Alt + N = nº reclamación · Alt + B = buscador · Alt + L = limpiar selección"
    )

    # Streamlit mantiene estado propio para cada checkbox y para el text_area.
    # Las limpiezas deben aplicarse ANTES de pintar esos widgets, si no,
    # la selección visual puede quedarse marcada aunque borremos selected_keys.
    if st.session_state.get("pending_clear_selection", False):
        st.session_state.selected_keys = []
        st.session_state.previous_selected_keys = []
        for comment_key in COMMENTS:
            st.session_state[f"chk_{comment_key}"] = False
        st.session_state.pending_clear_selection = False

    if st.session_state.get("pending_clear_comment_area", False):
        st.session_state.final_comment_area = ""
        st.session_state.pending_clear_comment_area = False

    if st.session_state.last_saved_comment:
        st.success("✅ Último comentario guardado en CSVs. Cópialo con el botón de navegador de abajo.")
        render_browser_copy_button(
            st.session_state.last_saved_comment,
            button_text="📋 Copiar último comentario guardado",
            key="copy_last_saved_comment"
        )
        with st.expander("Ver último comentario guardado", expanded=False):
            st.code(st.session_state.last_saved_comment, language=None)
        if st.button("Ocultar último comentario guardado", key="hide_last_saved_comment"):
            st.session_state.last_saved_comment = ""
            st.rerun()

    usage_stats = load_usage_stats()
    rank_map = get_usage_rank_map(usage_stats)

    col_c1, col_c2, col_c3 = st.columns([1.5, 2, 1.2])
    with col_c1:
        st.session_state.claim_val = st.text_input(
            "Nº reclamación / garantía:",
            value=st.session_state.claim_val,
            placeholder="Ej: CO202608310001",
            help="Rellénalo antes de guardar si quieres poder cruzar después el registro con el DMS."
        ).strip().upper()

    with col_c2:
        search_query = st.text_input(
            "Buscar:",
            placeholder="Filtrar por ID, motivo, categoría o texto..."
        ).strip()

    with col_c3:
        highlight_top = st.checkbox("Resaltar TOP en lista", value=False)
        warn_missing_claim = st.checkbox("Avisar si falta claim", value=True)

    # Disclaimer fijo: la claim es opcional, pero es la única clave sólida para cruzar con DMS.
    if not st.session_state.claim_val:
        st.warning(
            "⚠️ **Antes de guardar:** revisa si quieres informar el **número de reclamación/garantía**. "
            "Si lo dejas vacío, el registro se guardará como `NO INFORMADO` y luego no podrás cruzarlo de forma fiable con el DMS."
        )
    else:
        st.success(f"✅ Reclamación informada para el registro: `{st.session_state.claim_val}`")

    st.markdown("---")

    query_norm = normalize_text(search_query)
    grid_cols = st.columns(3)

    for idx, category in enumerate(CATEGORY_ORDER):
        target_col = grid_cols[idx % 3]
        cat_items = []
        for key, item in COMMENTS.items():
            if item["category"] == category:
                searchable = normalize_text(f"{key} {item['category']} {item['label']} {item['text']}")
                if not query_norm or query_norm in searchable:
                    cat_items.append((key, item))

        if cat_items:
            with target_col:
                st.markdown(f"#### {category}")
                for key, item in cat_items:
                    uses = usage_stats.get(key, 0)
                    rank = rank_map.get(key)

                    usage_text = f"{uses} usos"
                    if highlight_top and rank is not None:
                        if rank <= TOP_RED_LIMIT:
                            usage_text += f" 🔥 TOP {rank}"
                        elif rank <= TOP_AMBER_LIMIT:
                            usage_text += f" ⚠️ TOP {rank}"

                    label_text = f"**{item['label']}**  \n:gray[({usage_text})]"
                    is_checked = key in st.session_state.selected_keys

                    checked = st.checkbox(label_text, value=is_checked, key=f"chk_{key}")
                    if checked and key not in st.session_state.selected_keys:
                        st.session_state.selected_keys.append(key)
                    elif not checked and key in st.session_state.selected_keys:
                        st.session_state.selected_keys.remove(key)

    st.markdown("---")

    ordered_keys = [key for key in COMMENTS if key in st.session_state.selected_keys]
    base_comment = " ".join(COMMENTS[key]["text"] for key in ordered_keys).strip()

    # Streamlit mantiene el valor de un text_area con key aunque cambie el parámetro value.
    # Por eso sincronizamos manualmente el comentario editable cuando cambia la selección
    # de motivos, replicando el comportamiento de la app Tkinter original.
    previous_ordered_keys = [key for key in COMMENTS if key in st.session_state.previous_selected_keys]
    selection_changed = ordered_keys != previous_ordered_keys

    if selection_changed:
        st.session_state.final_comment_area = base_comment
        st.session_state.previous_selected_keys = ordered_keys.copy()

    st.subheader("Comentario generado editable:")
    final_comment = st.text_area(
        "Revisa o modifica el texto antes de copiar:",
        height=110,
        key="final_comment_area"
    ).strip()

    render_browser_copy_button(
        final_comment,
        button_text="📋 Copiar comentario al portapapeles",
        key="copy_current_comment"
    )

    def procesar_copia():
        if not final_comment.strip():
            st.warning("⚠️ Sin comentario: No hay ningún comentario para copiar.")
            return False

        if warn_missing_claim and not st.session_state.claim_val and not st.session_state.confirm_missing_claim:
            st.session_state.confirm_missing_claim = True
            st.error(
                "🚨 **No has informado el número de reclamación/garantía.** "
                "Si guardas así, se registrará como `NO INFORMADO` y no servirá para cruce claim a claim con DMS. "
                "Rellena el campo de arriba o vuelve a pulsar el botón de guardar para confirmar que quieres guardarlo sin claim."
            )
            return False

        log_generated_comment(ordered_keys, final_comment, base_comment, st.session_state.claim_val)

        if ordered_keys:
            update_usage_stats(ordered_keys)

        st.session_state.confirm_missing_claim = False
        return True

    col_b1, col_b2, col_b3, col_b4 = st.columns([1.35, 1.45, 1.2, 1.6])

    with col_b1:
        if st.button("💾 Guardar", type="primary", use_container_width=True):
            if procesar_copia():
                st.session_state.last_saved_comment = final_comment
                st.code(final_comment, language=None)
                st.success("✅ Comentario registrado en CSVs. Cópialo con el botón de navegador.")
                render_browser_copy_button(
                    final_comment,
                    button_text="📋 Copiar ahora",
                    key="copy_after_save"
                )

    with col_b2:
        if st.button("🧹 Guardar y limpiar", use_container_width=True):
            if procesar_copia():
                st.session_state.last_saved_comment = final_comment
                st.session_state.claim_val = ""
                st.session_state.pending_clear_selection = True
                st.session_state.pending_clear_comment_area = True
                st.rerun()

    with col_b3:
        if st.button("🗑️ Limpiar selección", use_container_width=True):
            st.session_state.confirm_missing_claim = False
            st.session_state.pending_clear_selection = True
            st.session_state.pending_clear_comment_area = True
            st.rerun()

    with col_b4:
        if st.button("🧽 Limpiar claim", use_container_width=True):
            st.session_state.claim_val = ""
            st.rerun()

    st.markdown("---")
    with st.expander("📊 Ver Estadísticas y Análisis de Motivos", expanded=False):
        df_stats = get_stats_dataframe(usage_stats)
        df_cat = get_category_stats_dataframe(usage_stats)

        if df_stats.empty:
            st.info("Aún no hay usos registrados. Cuando copies comentarios, aparecerán aquí.")
            return

        total_uses = int(df_stats["Usos"].sum())
        st.write(
            f"**Total de usos registrados:** {total_uses} · "
            f"**Motivos usados:** {len(df_stats)} · "
            f"**Categorías usadas:** {len(df_cat)} · "
            "No se incluyen motivos con 0 usos. El ranking respeta empates."
        )

        tab_tbl, tab_bars, tab_pie = st.tabs(["Tabla de Ranking", "Barras por Motivo", "Pie por Categoría"])

        with tab_tbl:
            display_df = df_stats.copy()
            display_df["%"] = display_df["%"].map(lambda x: f"{x:.1f}%")
            st.dataframe(
                display_df[["TOP", "Usos", "%", "Categoría", "Motivo"]],
                use_container_width=True,
                hide_index=True
            )

        color_domain = CATEGORY_ORDER + ["Comentario manual"]
        color_range = [CATEGORY_COLOR_MAP.get(category, "#7f7f7f") for category in color_domain]

        with tab_bars:
            try:
                import altair as alt

                bar_df = df_stats.copy()
                bar_df["Porcentaje"] = bar_df["%"].map(lambda value: f"{value:.1f}%")

                bar_chart = (
                    alt.Chart(bar_df)
                    .mark_bar()
                    .encode(
                        x=alt.X("Usos:Q", title="Usos"),
                        y=alt.Y(
                            "Motivo:N",
                            title="Motivo",
                            sort=alt.EncodingSortField(field="Usos", order="descending"),
                        ),
                        color=alt.Color(
                            "Categoría:N",
                            scale=alt.Scale(domain=color_domain, range=color_range),
                            legend=alt.Legend(title="Categoría"),
                        ),
                        tooltip=[
                            alt.Tooltip("TOP:N", title="TOP"),
                            alt.Tooltip("Motivo:N", title="Motivo"),
                            alt.Tooltip("Categoría:N", title="Categoría"),
                            alt.Tooltip("Usos:Q", title="Usos"),
                            alt.Tooltip("Porcentaje:N", title="%"),
                        ],
                    )
                    .properties(height=max(320, len(bar_df) * 28))
                )

                st.altair_chart(bar_chart, use_container_width=True)
            except Exception as exc:
                st.warning(f"No se pudo mostrar el gráfico de barras. Muestro la tabla en su lugar. Detalle: {exc}")
                st.dataframe(df_stats[["Motivo", "Categoría", "Usos"]], use_container_width=True, hide_index=True)

        with tab_pie:
            if df_cat.empty:
                st.info("Aún no hay categorías con usos registrados.")
            else:
                try:
                    import altair as alt

                    pie_df = df_cat.copy()
                    pie_df["Porcentaje"] = pie_df["%"].map(lambda value: f"{value:.1f}%")

                    pie_chart = (
                        alt.Chart(pie_df)
                        .mark_arc(innerRadius=45)
                        .encode(
                            theta=alt.Theta("Usos:Q", title="Usos"),
                            color=alt.Color(
                                "Categoría:N",
                                scale=alt.Scale(domain=color_domain, range=color_range),
                                legend=alt.Legend(title="Categoría"),
                            ),
                            tooltip=[
                                alt.Tooltip("Categoría:N", title="Categoría"),
                                alt.Tooltip("Usos:Q", title="Usos"),
                                alt.Tooltip("Porcentaje:N", title="%"),
                            ],
                        )
                        .properties(height=430)
                    )

                    st.altair_chart(pie_chart, use_container_width=True)
                    st.caption("Los colores del pie y de las barras usan la misma leyenda por categoría.")
                except Exception as exc:
                    st.warning(f"No se pudo mostrar el pie chart. Muestro la tabla en su lugar. Detalle: {exc}")
                    display_cat = df_cat.copy()
                    display_cat["%"] = display_cat["%"].map(lambda value: f"{value:.1f}%")
                    st.dataframe(display_cat, use_container_width=True, hide_index=True)


# =========================================================================
# PANTALLA 3 - SOLICITAR OPERACIÓN
# =========================================================================
def render_solicitar_operacion(txt_local):
    st.title(txt_local["solicitar_titulo"])
    st.write(txt_local["solicitar_sub"])
    st.markdown("---")

    MAPEO_MODELOS = {
        "OMODA 5 (Gasolina)": "T19C",
        "OMODA 5 HEV (Híbrido)": "T19C HEV",
        "OMODA 5 EV (Eléctrico)": "T19C EV",
        "OMODA 7 PHEV": "T1GC PHEV",
        "OMODA 9 PHEV": "T22 PHEV",
        "JAECOO 5 (Gasolina)": "T13J",
        "JAECOO 5 HEV": "T13J HEV",
        "JAECOO 5 BEV": "T13J BEV",
        "JAECOO 7 (Gasolina)": "T1EJ",
        "JAECOO 7 HEV": "T1EJ HEV",
        "JAECOO 7 PHEV": "T1EJ PHEV",
        "JAECOO 8 PHEV": "T26 PHEV",
        "LEPAS L8 PHEV": "T1G PHEV",
    }

    st.subheader(txt_local["form_sub"])

    col1, col2 = st.columns(2)
    with col1:
        marca = st.selectbox(txt_local["form_marca"], ["OMODA", "JAECOO", "LEPAS"])
        modelos_filtrados = [mod for mod in MAPEO_MODELOS if mod.upper().startswith(marca.upper())]
        modelo_comercial = st.selectbox(txt_local["form_modelo"], modelos_filtrados)
    with col2:
        codigo_producto_auto = MAPEO_MODELOS[modelo_comercial]
        st.text_input(txt_local["form_hq_code"], value=codigo_producto_auto, disabled=True)

    with st.form("hq_operation_form", clear_on_submit=True):
        numero_garantia = st.text_input(
            "Nº de Garantía:",
            placeholder="Ej: CO202607290001",
            help="Formato requerido: COYYYYMMDDXXXX"
        ).strip().upper()

        c1, c2 = st.columns(2)
        with c1:
            vin = st.text_input(txt_local["form_vin"], max_chars=17, placeholder=txt_local["form_vin_holder"]).strip().upper()
        with c2:
            referencia = st.text_input(txt_local["form_ref"], placeholder=txt_local["form_ref_holder"]).strip().upper()

        operacion_solicitada = st.text_area(
            txt_local["form_op"],
            placeholder=txt_local["form_op_holder"]
        ).strip()

        boton_enviar = st.form_submit_button(txt_local["form_btn"])

        if boton_enviar:
            patron_garantia = r"^CO\d{8}[A-Z0-9]{4}$"

            if not numero_garantia or not vin or not operacion_solicitada:
                st.error(txt_local["err_campos"])
            elif not re.match(patron_garantia, numero_garantia):
                st.error("❌ **Error en el Número de Garantía:** Debe cumplir el patrón **COYYYYMMDDXXXX**.")
            elif len(vin) != 17:
                st.error("❌ **Error en el VIN:** El número de bastidor debe tener exactamente 17 caracteres.")
            else:
                ahora = datetime.datetime.now()
                columnas_orden = [
                    "SN", "Submitted on", "Respondents", "Fecha del día",
                    "Marca del vehículo", "INTRODUCIR MODELO", "INTRODUCIR VIN",
                    "Mercado", "CÓDIGO DE PRODUCTO", "REFERENCIA DE PIEZA",
                    "OPERACIÓN QUE SE SOLICITA AÑADIR", "DEALER"
                ]

                nueva_solicitud = {
                    "SN": len(st.session_state.lista_solicitudes) + 1,
                    "Submitted on": ahora.strftime("%Y-%m-%d %H:%M:%S"),
                    "Respondents": f"Garantía: {numero_garantia}",
                    "Fecha del día": ahora.strftime("%Y-%m-%d"),
                    "Marca del vehículo": marca,
                    "INTRODUCIR MODELO": modelo_comercial,
                    "INTRODUCIR VIN": vin,
                    "Mercado": "Spain OJ",
                    "CÓDIGO DE PRODUCTO": codigo_producto_auto,
                    "REFERENCIA DE PIEZA": referencia if referencia else "NaN",
                    "OPERACIÓN QUE SE SOLICITA AÑADIR": operacion_solicitada,
                    "DEALER": numero_garantia,
                }

                subida_exitosa = False

                try:
                    from streamlit_gsheets import GSheetsConnection
                    conn = st.connection("gsheets", type=GSheetsConnection)

                    if "connections" in st.secrets and "gsheets" in st.secrets["connections"]:
                        spreadsheet_url = st.secrets["connections"]["gsheets"]["spreadsheet"]
                    elif "gsheets" in st.secrets and "spreadsheet" in st.secrets["gsheets"]:
                        spreadsheet_url = st.secrets["gsheets"]["spreadsheet"]
                    else:
                        spreadsheet_url = st.secrets.get("spreadsheet", "")

                    df_cloud = conn.read(spreadsheet=spreadsheet_url) if spreadsheet_url else pd.DataFrame(columns=columnas_orden)
                    if df_cloud.empty or len(df_cloud.columns) < 2:
                        df_cloud = pd.DataFrame(columns=columnas_orden)
                    else:
                        df_cloud = df_cloud.dropna(how="all").loc[:, ~df_cloud.columns.str.contains("^Unnamed")]

                    nueva_solicitud["SN"] = len(df_cloud) + 1
                    df_nuevo = pd.DataFrame([nueva_solicitud]).reindex(columns=columnas_orden)
                    df_cloud = df_cloud.reindex(columns=columnas_orden)
                    df_actualizado = pd.concat([df_cloud, df_nuevo], ignore_index=True)

                    if spreadsheet_url:
                        conn.update(spreadsheet=spreadsheet_url, data=df_actualizado)
                        subida_exitosa = True
                    else:
                        raise ValueError("No se encontró la URL del archivo de Sheets en st.secrets.")

                except Exception as exc:
                    st.error(f"❌ Error de conexión con Google Sheets: {exc}")
                    st.info("💡 Por seguridad, hemos guardado esta línea en la caché local.")

                st.session_state.lista_solicitudes.append(nueva_solicitud)

                if subida_exitosa:
                    st.success("✅ **Operación registrada con éxito.** La solicitud ha sido transmitida a Central.")
                    time.sleep(1.5)
                    st.rerun()

    if st.session_state.lista_solicitudes:
        st.markdown("---")
        st.subheader("📌 Solicitudes registradas en esta sesión")
        st.dataframe(pd.DataFrame(st.session_state.lista_solicitudes), use_container_width=True, hide_index=True)


# =========================================================================
# PANTALLA 4 - CONSULTORIO IA
# =========================================================================
def render_consultorio_ia():
    st.title("🤖 Consultor Técnico de Garantías (Inteligencia Artificial)")
    st.write("Analiza de forma preliminar si una avería está cubierta según el manual de políticas oficial e identifica los pasos técnicos a seguir.")
    st.markdown("---")

    st.subheader("📝 Detalles de la Consulta")
    descripcion_averia = st.text_area(
        "Descripción de la avería o síntomas del vehículo:",
        placeholder="Ejemplo: Cliente reporta ruido metálico al girar el volante a la izquierda en OMODA 5...",
        height=150
    )

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
                st.session_state.resultado_consultorio = consultar_ia_garantias(descripcion_averia, parametro_imagenes)

    if st.session_state.resultado_consultorio:
        st.markdown("### 📋 Informe de Diagnóstico Generado")
        st.markdown(st.session_state.resultado_consultorio)
        st.success("✅ Análisis preliminar finalizado.")
        st.markdown("<br>", unsafe_allow_html=True)

        st.warning("""
#### ⚠️ NOTA OBLIGATORIA DE CENTRAL
Este informe constituye una **valoración preliminar e informativa** basada exclusivamente en los síntomas y evidencias gráficas aportadas por el taller.

Para validar definitivamente el diagnóstico técnico, proceder con la autorización de la reparación bajo garantía o reportar de forma oficial un fallo de fabricación de origen, **es obligatorio abrir un canal oficial en la plataforma aportando el bastidor (VIN) completo**:

* 🛠️ **¿Dudas sobre el diagnóstico técnico o el proceso de reparación?** Abra un **Ticket de Asistencia Técnica** o escriba a [soportetecnico@omodaes.com](mailto:soportetecnico@omodaes.com)
* 📝 **¿Consultas sobre plazos de cobertura o tramitación?** Contacte con [garantias@omodaes.com](mailto:garantias@omodaes.com)
""")


# =========================================================================
# MAIN
# =========================================================================
txt, opcion_menu = render_sidebar_and_get_option()

if check_password(txt):
    if opcion_menu == txt["menu_taller"]:
        render_tiempos_taller(txt)
    elif opcion_menu == txt["menu_generador"]:
        render_generador_comentarios()
    elif opcion_menu == txt["menu_solicitar"]:
        render_solicitar_operacion(txt)
    elif opcion_menu == txt["menu_consultorio"]:
        render_consultorio_ia()

