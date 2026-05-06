import os
import re
import json
import ssl
import sqlite3
import traceback
import inspect
from datetime import datetime

import pandas as pd
import etap.api

import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


# =============================================================================
# CONFIGURACION GENERAL
# =============================================================================

BASE_ADDRESS = "https://localhost:60000"

GET_ONLINE_DATA = False
ONLINE_CONFIG_ONLY = False
TIMEOUT_SECS = 60

# Optimización de desempeño:
# True: extracción rápida de datos del modelo usando getallelementdata() y getelementnames().
# False: permite consultas profundas getelementprop() elemento por elemento, pero puede ser muy lento.
FAST_MODEL_EXTRACTION = True

# Enriquecimiento selectivo:
# True: consulta únicamente propiedades críticas por elemento (pocas columnas).
# Es mucho más rápido que el modo profundo completo y evita que las tablas queden solo con ID.
ENRICH_MODEL_CRITICAL_PROPERTIES = False

# Máximo de elementos por categoría a enriquecer con propiedades críticas.
MAX_ENRICH_ELEMENTS_PER_CATEGORY = 80

# Genera archivos CSV de diagnóstico con columnas crudas de ETAP API.
EXPORT_MODEL_DEBUG_FILES = False

# Versión rápida: el capítulo 4 muestra inventario resumido de elementos solo con ID.
MODEL_ELEMENTS_SUMMARY_ONLY = True

# Límite de filas en tablas grandes para reducir tiempo de generación del DOCX.
MAX_MAIN_VOLTAGE_ROWS = 35
MAX_ANNEX_VOLTAGE_ROWS = 80
MAX_MODEL_TABLE_ROWS = 35

DEFAULT_MIN_VOLTAGE_PERCENT = 95.0
DEFAULT_MAX_VOLTAGE_PERCENT = 105.0
DEFAULT_MAX_LOADING_PERCENT = 100.0

MIN_VOLTAGE_PERCENT = DEFAULT_MIN_VOLTAGE_PERCENT
MAX_VOLTAGE_PERCENT = DEFAULT_MAX_VOLTAGE_PERCENT
MAX_LOADING_PERCENT = DEFAULT_MAX_LOADING_PERCENT

FALLBACK_REVISIONS = ["Base"]
FALLBACK_CONFIGURATIONS = ["Normal"]
FALLBACK_PRESENTATIONS = ["Study View"]
DEFAULT_PRESENTATION = "Study View"
FALLBACK_STUDY_CASES = ["LF Report", "LF 100A", "LF"]
COMMON_LF_STUDY_CASE_CANDIDATES = ["LF Report", "LF 100A", "LF"]
COMMON_PRESENTATION_CANDIDATES = ["Study View", "OLV1", "One-Line View", "One-Line Diagram", "Default"]
FALLBACK_OUTPUT_REPORTS = ["Untitled"]

# Propiedades base del documento Word, definidas a partir del archivo de referencia
# suministrado por el usuario. El script NO usa plantilla: crea un .docx nuevo.
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

DOC_FONT_NAME = "Segoe UI"
DOC_FONT_SIZE = 10
DOC_TABLE_FONT_SIZE = 8
DOC_TITLE_SIZE = 14
DOC_SUBTITLE_SIZE = 12
DOC_HEADING1_SIZE = 12
DOC_HEADING2_SIZE = 10

DOC_HEADER_FILL = "D9EAF7"
DOC_BLUE = "000000"
DOC_GRAY = "F2F2F2"
DOC_BORDER = "BFBFBF"



# =============================================================================
# UTILIDADES DE INTERFAZ GRAFICA
# =============================================================================

def center_window(root, width, height):
    """
    Centra la ventana en pantalla y evita que los botones queden fuera del área visible.
    """
    root.update_idletasks()

    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()

    # Mantener margen para barra de tareas / escalado de Windows.
    width = min(width, max(500, screen_width - 80))
    height = min(height, max(350, screen_height - 120))

    x = int((screen_width - width) / 2)
    y = int((screen_height - height) / 2)

    root.geometry(str(width) + "x" + str(height) + "+" + str(x) + "+" + str(y))
    root.minsize(min(width, 520), min(height, 360))


def make_scrollable_body(root):
    """
    Crea un cuerpo desplazable y deja una zona inferior fija para botones.
    Esto evita que los botones Aceptar/Cancelar queden recortados.
    """
    outer = ttk.Frame(root)
    outer.pack(fill="both", expand=True, padx=10, pady=(0, 5))

    canvas = tk.Canvas(outer, highlightthickness=0)
    scrollbar = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
    body = ttk.Frame(canvas)

    body.bind(
        "<Configure>",
        lambda event: canvas.configure(scrollregion=canvas.bbox("all"))
    )

    canvas_window = canvas.create_window((0, 0), window=body, anchor="nw")

    def resize_body(event):
        canvas.itemconfig(canvas_window, width=event.width)

    canvas.bind("<Configure>", resize_body)
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    return body


def make_fixed_button_bar(root):
    """
    Crea una barra inferior fija para botones.
    """
    separator = ttk.Separator(root, orient="horizontal")
    separator.pack(side="bottom", fill="x")

    button_frame = ttk.Frame(root)
    button_frame.pack(side="bottom", fill="x", padx=12, pady=10)

    return button_frame

# =============================================================================
# UTILIDADES GENERALES
# =============================================================================

def unique_keep_order(values):
    result = []
    seen = set()

    for value in values:
        if value is None:
            continue

        text = str(value).strip()

        if text == "":
            continue

        if len(text) > 250:
            continue

        key = text.lower()

        if key not in seen:
            result.append(text)
            seen.add(key)

    return result


def normalize_text(value):
    text = str(value).lower().strip()
    replacements = {
        "á": "a", "é": "e", "í": "i", "ó": "o", "ú": "u", "ñ": "n",
        "-": "", "_": "", " ": "", ".": "", ",": "", "%": "percent"
    }

    for old, new in replacements.items():
        text = text.replace(old, new)

    return text


def flatten_strings(value):
    result = []

    if value is None:
        return result

    if isinstance(value, bytes):
        try:
            value = value.decode("utf-8", errors="ignore")
        except Exception:
            value = str(value)

    if isinstance(value, str):
        text = value.strip()

        if text == "":
            return result

        try:
            parsed = json.loads(text)
            return flatten_strings(parsed)
        except Exception:
            pass

        if "<" in text and ">" in text:
            tags = re.findall(r">([^<>]+)<", text)
            for item in tags:
                item = item.strip()
                if item:
                    result.append(item)
            if result:
                return result

        if "," in text and len(text) < 1000:
            for part in text.split(","):
                part = part.strip()
                if part:
                    result.append(part)
            return result

        result.append(text)
        return result

    if isinstance(value, dict):
        preferred_keys = [
            "Name", "name", "DisplayName", "displayName",
            "ID", "id", "Value", "value", "Title", "title",
            "Text", "text", "Description", "description",
            "StudyCase", "studyCase", "OutputReport", "outputReport",
            "Presentation", "presentation", "Configuration", "configuration",
            "Revision", "revision"
        ]

        for key in preferred_keys:
            if key in value:
                result.extend(flatten_strings(value[key]))

        for key, val in value.items():
            if key in preferred_keys:
                continue
            if isinstance(val, (list, tuple, dict)):
                result.extend(flatten_strings(val))

        return result

    if isinstance(value, (list, tuple, set)):
        for item in value:
            result.extend(flatten_strings(item))
        return result

    text = str(value).strip()
    if text:
        result.append(text)

    return result


def looks_like_option(text):
    if text is None:
        return False

    text = str(text).strip()

    if text == "":
        return False

    if len(text) > 120:
        return False

    rejected = [
        "http://", "https://", "{", "}", "[", "]", "<", ">",
        "system.", "microsoft.", "traceback", "exception",
        "success", "true", "false", "null"
    ]

    low = text.lower()

    for item in rejected:
        if item in low:
            return False

    return True


def filter_options(values):
    return unique_keep_order([v for v in values if looks_like_option(v)])


def merge_options(*sources):
    values = []
    for source in sources:
        values.extend(source)
    return filter_options(values)


def parse_float(value, default_value):
    try:
        if value is None:
            return default_value

        text = str(value).replace(",", ".").replace("%", "").strip()

        if text == "":
            return default_value

        return float(text)
    except Exception:
        return default_value


def to_number(value):
    try:
        if value is None:
            return None

        if isinstance(value, str):
            value = value.replace(",", "")
            value = value.replace("%", "")
            value = value.strip()

        return float(value)
    except Exception:
        return None


def clean_filename(text):
    text = str(text).strip()
    if text == "":
        text = "Informe_Flujo_de_Potencia"

    for ch in ['\\', '/', ':', '*', '?', '"', '<', '>', '|']:
        text = text.replace(ch, "_")

    return text


# =============================================================================
# CONEXION ETAP
# =============================================================================

def connect_to_etap():
    print("Conectando con ETAP DataHub...")
    e = etap.api.connect(BASE_ADDRESS)

    print("Verificando conexion...")
    ping_result = e.application.ping()
    print("Ping:", str(ping_result))

    return e


# =============================================================================
# DESCUBRIMIENTO DE OPCIONES ETAP
# =============================================================================

def safe_call_no_arg(obj, method_name):
    try:
        if obj is None:
            return []

        if not hasattr(obj, method_name):
            return []

        member = getattr(obj, method_name)

        if not callable(member):
            return flatten_strings(member)

        try:
            sig = inspect.signature(member)
            required = [
                p for p in sig.parameters.values()
                if p.default is inspect._empty
                and p.kind in (p.POSITIONAL_ONLY, p.POSITIONAL_OR_KEYWORD, p.KEYWORD_ONLY)
            ]
            if len(required) > 0:
                return []
        except Exception:
            pass

        response = member()
        return flatten_strings(response)

    except Exception:
        return []


def collect_all_zero_arg_methods(obj, section_name):
    rows = []

    if obj is None:
        return rows

    try:
        names = dir(obj)
    except Exception:
        return rows

    allowed_tokens = [
        "get", "list", "all", "names", "name", "case", "study",
        "report", "presentation", "configuration", "config", "revision",
        "scenario", "output", "load", "lf"
    ]

    blocked_tokens = [
        "run", "set", "delete", "remove", "create", "update", "open", "close",
        "save", "start", "stop", "connect", "disconnect", "shutdown", "write",
        "post", "put", "patch"
    ]

    for method_name in names:
        if method_name.startswith("_"):
            continue

        low = method_name.lower()

        if any(token in low for token in blocked_tokens):
            continue

        if not any(token in low for token in allowed_tokens):
            continue

        values = safe_call_no_arg(obj, method_name)

        if values:
            rows.append({
                "section": section_name,
                "method": method_name,
                "values": filter_options(values)
            })

    return rows


def classify_discovered_values(discovery_rows):
    options = {
        "revisions": [],
        "configurations": [],
        "presentations": [],
        "study_cases": [],
        "output_reports": []
    }

    for row in discovery_rows:
        method_norm = normalize_text(row["method"])
        values = row["values"]

        if not values:
            continue

        if "revision" in method_norm:
            options["revisions"].extend(values)

        if "configuration" in method_norm or "config" in method_norm:
            options["configurations"].extend(values)

        if "presentation" in method_norm:
            options["presentations"].extend(values)

        if "study" in method_norm or "case" in method_norm or "lf" in method_norm or "loadflow" in method_norm:
            options["study_cases"].extend(values)

        if "report" in method_norm or "output" in method_norm:
            options["output_reports"].extend(values)

    for key in options:
        options[key] = filter_options(options[key])

    return options


def discover_from_python_api(e):
    sections = []

    candidate_section_names = [
        "application", "projectData", "projectdata", "project",
        "scenario", "scenarios", "studies", "study", "presentation",
        "presentations", "configuration", "configurations"
    ]

    for section_name in candidate_section_names:
        try:
            section = getattr(e, section_name)
            sections.append((section_name, section))
        except Exception:
            pass

    discovery_rows = []

    for section_name, section in sections:
        discovery_rows.extend(collect_all_zero_arg_methods(section, section_name))

    options = classify_discovered_values(discovery_rows)

    return options, discovery_rows


def safe_urlopen_json(url):
    try:
        import urllib.request

        context = ssl._create_unverified_context()
        with urllib.request.urlopen(url, context=context, timeout=5) as response:
            data = response.read().decode("utf-8", errors="ignore")
            return json.loads(data)
    except Exception:
        return None


def safe_get_text(url):
    try:
        import urllib.request

        context = ssl._create_unverified_context()
        with urllib.request.urlopen(url, context=context, timeout=8) as response:
            return response.read().decode("utf-8", errors="ignore")
    except Exception:
        return ""


def discover_from_swagger():
    swagger_candidates = [
        BASE_ADDRESS + "/swagger/v1/swagger.json",
        BASE_ADDRESS + "/swagger.json",
        BASE_ADDRESS + "/openapi.json"
    ]

    swagger = None

    for url in swagger_candidates:
        swagger = safe_urlopen_json(url)
        if swagger is not None:
            break

    options = {
        "revisions": [],
        "configurations": [],
        "presentations": [],
        "study_cases": [],
        "output_reports": []
    }

    if swagger is None:
        return options

    paths = swagger.get("paths", {})

    keyword_map = {
        "revisions": ["revision"],
        "configurations": ["configuration", "config"],
        "presentations": ["presentation"],
        "study_cases": ["studycase", "study-case", "loadflowstudy", "lfstudy", "study"],
        "output_reports": ["outputreport", "output-report", "report"]
    }

    for path, operations in paths.items():
        if "get" not in operations:
            continue

        op = operations["get"]
        params = op.get("parameters", [])

        required_params = [p for p in params if p.get("required", False)]

        if required_params:
            continue

        path_norm = normalize_text(path + " " + op.get("summary", "") + " " + op.get("operationId", ""))

        for key, keywords in keyword_map.items():
            matched = False

            for kw in keywords:
                if normalize_text(kw) in path_norm:
                    matched = True
                    break

            if not matched:
                continue

            url = BASE_ADDRESS + path
            text_response = safe_get_text(url)
            values = filter_options(flatten_strings(text_response))

            options[key].extend(values)

    for key in options:
        options[key] = filter_options(options[key])

    return options


def scan_existing_lf_reports(search_folder):
    reports = []

    if not search_folder or not os.path.exists(search_folder):
        return reports

    for root, dirs, files in os.walk(search_folder):
        for file_name in files:
            if file_name.lower().endswith(".lf1s"):
                reports.append(os.path.splitext(file_name)[0])

    return filter_options(reports)



def is_load_flow_related_name(name):
    """
    Filtro estricto para nombres de casos/reportes de flujo de carga.
    Evita abreviaturas globales de ETAP como ILS, MS, SC, RA, TDLF, etc.
    """
    if name is None:
        return False

    text = str(name).strip()
    if text == "":
        return False

    low = text.lower()

    # Rechazar módulos/abreviaturas globales observadas en ETAP.
    global_or_other_modules = {
        "edit", "ils", "loadalloc", "ms", "train", "ra", "sc", "sm",
        "szm", "agc", "so", "ssm", "tdlf", "load_flow",
        "short_circuit", "arc_flash", "motor_acceleration",
        "harmonic", "transient_stability", "star"
    }

    if low in global_or_other_modules:
        return False

    # Rechazar estudios que claramente no son flujo de carga.
    reject_tokens = [
        "short", "circuit", "arc", "flash", "harmonic", "motor",
        "acceleration", "transient", "stability", "star", "ground",
        "reliability", "protection", "coordination", "battery", "emt",
        "switching", "optimal", "opf"
    ]

    for token in reject_tokens:
        if token in low:
            return False

    # Aceptar solo nombres explícitos de flujo de carga o reportes LF.
    if low.startswith("lf"):
        return True

    if "load flow" in low or "loadflow" in low or "flujo" in low:
        return True

    # Reportes de salida específicos pueden llamarse REPCASO1, REPCASO2, etc.
    if low.startswith("rep"):
        return True

    return False


def filter_load_flow_options(values):
    result = []
    for value in values:
        text = str(value).strip()
        if text and is_load_flow_related_name(text):
            result.append(text)
    return filter_options(result)


def scan_lf_reports_with_paths(search_folder):
    """
    Escanea únicamente la carpeta seleccionada por el usuario para reportes .LF1S
    y guarda sus rutas reales en REPORT_NAME_TO_PATH.
    """
    reports = []

    try:
        REPORT_NAME_TO_PATH.clear()
    except Exception:
        pass

    if not search_folder or not os.path.exists(search_folder):
        return reports

    for root, dirs, files in os.walk(search_folder):
        for file_name in files:
            if file_name.lower().endswith(".lf1s"):
                report_name = os.path.splitext(file_name)[0]
                full_path = os.path.join(root, file_name)
                reports.append(report_name)

                try:
                    # Si hay duplicados, conservar el más reciente.
                    if report_name not in REPORT_NAME_TO_PATH:
                        REPORT_NAME_TO_PATH[report_name] = full_path
                    else:
                        if os.path.getmtime(full_path) > os.path.getmtime(REPORT_NAME_TO_PATH[report_name]):
                            REPORT_NAME_TO_PATH[report_name] = full_path
                except Exception:
                    REPORT_NAME_TO_PATH[report_name] = full_path

    return filter_options(reports)


def clean_project_study_case_name(value):
    """
    Limpia candidatos a nombre de caso de estudio encontrados en archivos .LF1S.
    """
    if value is None:
        return ""

    text = str(value).strip()

    if text == "":
        return ""

    if len(text) > 80:
        return ""

    low = text.lower()

    reject_exact = {
        "edit", "load_flow", "short_circuit", "arc_flash", "harmonic",
        "motor_acceleration", "transient_stability", "star", "true", "false",
        "none", "null", "nan"
    }

    if low in reject_exact:
        return ""

    reject_tokens = [
        "short_circuit", "arc_flash", "harmonic", "motor_acceleration",
        "transient_stability", "schema", "sqlite", "table", "index"
    ]

    if any(token in low for token in reject_tokens):
        return ""

    # Aceptar cualquier nombre razonable extraído de metadata del archivo del proyecto.
    return text


def scan_lf_study_cases_from_report_db(report_path):
    """
    Intenta leer un archivo .LF1S como SQLite y extraer nombres de casos de estudio.
    La estructura interna puede cambiar por versión, por eso se busca en tablas y
    columnas con nombres relacionados con Study/Case/Scenario.
    """
    candidates = []

    if not report_path or not os.path.exists(report_path):
        return candidates

    try:
        conn = sqlite3.connect(report_path)
        cur = conn.cursor()

        cur.execute("SELECT name FROM sqlite_master WHERE type='table';")
        tables = [row[0] for row in cur.fetchall()]

        # 1) Buscar primero columnas con nombres claros.
        for table in tables:
            try:
                cur.execute('PRAGMA table_info("' + table + '");')
                cols = [row[1] for row in cur.fetchall()]

                interesting_cols = []
                for col in cols:
                    c = str(col).lower()
                    if (
                        "study" in c
                        or "case" in c
                        or "scenario" in c
                        or "studycase" in c
                    ):
                        interesting_cols.append(col)

                for col in interesting_cols:
                    try:
                        sql = 'SELECT DISTINCT "' + col + '" FROM "' + table + '" LIMIT 50;'
                        cur.execute(sql)
                        for row in cur.fetchall():
                            if row and row[0] is not None:
                                name = clean_project_study_case_name(row[0])
                                if name:
                                    candidates.append(name)
                    except Exception:
                        pass
            except Exception:
                pass

        # 2) Si no encontró, buscar cadenas de texto en tablas pequeñas de metadata.
        if not candidates:
            metadata_table_tokens = ["info", "meta", "study", "case", "setting", "option", "summary"]
            for table in tables:
                if not any(tok in str(table).lower() for tok in metadata_table_tokens):
                    continue

                try:
                    cur.execute('SELECT * FROM "' + table + '" LIMIT 100;')
                    rows = cur.fetchall()
                    for row in rows:
                        for item in row:
                            name = clean_project_study_case_name(item)
                            if name and is_load_flow_related_name(name):
                                candidates.append(name)
                except Exception:
                    pass

        conn.close()
    except Exception:
        pass

    return filter_options(candidates)


def scan_project_lf_study_cases(search_folder):
    """
    Escanea únicamente los .LF1S de la carpeta del proyecto para identificar casos
    de estudio propios del proyecto activo.
    """
    cases = []

    if not search_folder or not os.path.exists(search_folder):
        return cases

    for root, dirs, files in os.walk(search_folder):
        for file_name in files:
            if file_name.lower().endswith(".lf1s"):
                report_path = os.path.join(root, file_name)
                cases.extend(scan_lf_study_cases_from_report_db(report_path))

    cases = filter_options(cases)

    # Filtrar módulos globales por seguridad.
    cases = [c for c in cases if is_load_flow_related_name(c)]

    return filter_options(cases)



def safe_get_study_case(e, study_case_name):
    """
    Valida un caso de estudio usando ProjectData.getstudycase(study_case_name).

    Si ETAP devuelve un JSON/dict con contenido válido, se considera que el caso
    existe en el proyecto activo.
    """
    try:
        pd_obj = get_projectdata_object(e)
        if pd_obj is None:
            return None

        if not hasattr(pd_obj, "getstudycase"):
            return None

        response = pd_obj.getstudycase(study_case_name)

        if response is None:
            return None

        if isinstance(response, dict):
            data = response
        else:
            text = str(response).strip()
            if text == "":
                return None

            # Rechazar errores comunes.
            low = text.lower()
            if "error" in low or "exception" in low or "not found" in low or "invalid" in low:
                return None

            try:
                data = json.loads(text)
            except Exception:
                # Si devuelve texto no vacío sin error, conservarlo como señal débil.
                data = {"raw": text}

        if not data:
            return None

        return data

    except Exception:
        return None


def is_load_flow_study_case_data(data):
    """
    Determina si la respuesta de getstudycase corresponde a flujo de carga.
    En ETAP, el campo module puede indicar LOAD_FLOW, UNBALANCED_LOAD_FLOW, etc.
    analysisType también puede ser 0 para load flow en algunos ejemplos.
    """
    if data is None:
        return False

    if not isinstance(data, dict):
        return True

    module = str(data.get("module", "")).lower()
    study_id = str(data.get("id", "")).lower()

    load_flow_tokens = [
        "load_flow",
        "loadflow",
        "unbalanced_load_flow",
        "tdlf",
        "load"
    ]

    if any(token in module for token in load_flow_tokens):
        return True

    # Si el ID o nombre luce como LF, aceptarlo.
    if study_id.startswith("lf") or "load" in study_id:
        return True

    # Fallback: si analysisType existe y es 0, el ejemplo de ETAP muestra 0 para load flow.
    try:
        analysis_type = int(data.get("analysisType"))
        if analysis_type == 0:
            return True
    except Exception:
        pass

    return False


def get_project_load_flow_study_cases(e, search_folder=None):
    """
    Obtiene casos de estudio de flujo de carga del proyecto activo validando
    candidatos con ProjectData.getstudycase().

    Fuentes de candidatos:
    - Nombres típicos de ETAP.
    - Nombres inferidos desde reportes .LF1S de la carpeta del proyecto.
    - Nombres basados en reportes existentes, como REP-NORMAL -> LF NORMAL,
      como sugerencia validable.
    """
    candidates = []

    # Candidatos típicos.
    candidates.extend(COMMON_LF_STUDY_CASE_CANDIDATES)
    candidates.extend(FALLBACK_STUDY_CASES)

    # Candidatos desde metadata de .LF1S.
    try:
        candidates.extend(scan_project_lf_study_cases(search_folder))
    except Exception:
        pass

    # Candidatos derivados de nombres de reportes LF1S.
    try:
        report_names = scan_lf_reports_with_paths(search_folder)
        for rep in report_names:
            rep_text = str(rep).strip()
            if not rep_text:
                continue

            candidates.append(rep_text)

            low = rep_text.lower()
            # REP-NORMAL -> NORMAL / LF NORMAL / LF-NORMAL
            if low.startswith("rep"):
                suffix = rep_text[3:].strip("-_ ")
                if suffix:
                    candidates.append(suffix)
                    candidates.append("LF " + suffix)
                    candidates.append("LF-" + suffix)
                    candidates.append("LF_" + suffix)

            # CASO1 -> LF CASO1
            if "caso" in low:
                candidates.append("LF " + rep_text)
                candidates.append("LF-" + rep_text)

    except Exception:
        pass

    candidates = filter_options(candidates)

    valid_cases = []

    for candidate in candidates:
        data = safe_get_study_case(e, candidate)
        if data is None:
            continue

        if is_load_flow_study_case_data(data):
            valid_cases.append(candidate)

    valid_cases = filter_options(valid_cases)

    if valid_cases:
        return valid_cases

    # Último recurso editable, pero ahora no se asume que necesariamente funcione.
    return FALLBACK_STUDY_CASES



def discover_project_options(e, search_folder):
    """
    Busca opciones disponibles del proyecto sin mezclar casos/reportes globales.

    Casos de estudio:
    - Se validan con ProjectData.getstudycase(nombre), siguiendo el ejemplo oficial
      de ETAP. Solo se aceptan candidatos que devuelvan información válida del
      proyecto activo.
    - No se usan listas globales de ETAP API.

    Reportes de salida:
    - Se toman únicamente de archivos .LF1S dentro de la carpeta del proyecto.
    """
    print("")
    print("Buscando opciones disponibles del proyecto...")

    api_options, discovery_rows = discover_from_python_api(e)

    revisions = merge_options(
        api_options.get("revisions", []),
        FALLBACK_REVISIONS
    )

    configurations = merge_options(
        api_options.get("configurations", []),
        FALLBACK_CONFIGURATIONS
    )

    presentations = merge_options(api_options.get("presentations", []), FALLBACK_PRESENTATIONS, COMMON_PRESENTATION_CANDIDATES)

    # Casos propios del proyecto validados con getstudycase().
    study_cases = get_project_load_flow_study_cases(e, search_folder)

    # Reportes: solamente .LF1S encontrados en carpeta de proyecto.
    existing_reports = scan_lf_reports_with_paths(search_folder)

    if existing_reports:
        output_reports = existing_reports
    else:
        output_reports = FALLBACK_OUTPUT_REPORTS

    options = {
        "revisions": revisions,
        "configurations": configurations,
        "presentations": presentations,
        "study_cases": study_cases,
        "output_reports": output_reports
    }

    print("Opciones encontradas / editables:")
    print(" - Revisiones: " + ", ".join(options["revisions"]))
    print(" - Configuraciones: " + ", ".join(options["configurations"]))
    print(" - Presentaciones: " + ", ".join(options["presentations"]))
    print(" - Casos de estudio LF del proyecto: " + ", ".join(options["study_cases"]))
    print(" - Reportes de salida LF del proyecto: " + ", ".join(options["output_reports"]))

    return options, discovery_rows


# =============================================================================
# VENTANAS
# =============================================================================

def ask_project_information():
    data = {}
    cancelled = {}

    root = tk.Tk()
    root.title("Información del Proyecto")
    center_window(root, 900, 760)
    root.resizable(True, True)

    ttk.Label(
        root,
        text="Información del Proyecto",
        font=("Segoe UI", 14, "bold")
    ).pack(pady=(14, 6))

    ttk.Label(
        root,
        text="Esta información se utilizará para la portada, encabezado y cuerpo del informe Word.",
        font=("Segoe UI", 9)
    ).pack(pady=(0, 8))

    button_frame = make_fixed_button_bar(root)
    body = make_scrollable_body(root)

    frame = ttk.Frame(body)
    frame.pack(padx=20, pady=12, fill="both", expand=True)

    fields = [
        ("Título del informe", "titulo_informe", "INFORME TÉCNICO"),
        ("Subtítulo", "subtitulo", "ESTUDIO DE FLUJO DE POTENCIA"),
        ("Nombre del proyecto", "nombre_proyecto", "Proyecto ETAP"),
        ("Cliente", "cliente", ""),
        ("Ubicación del proyecto", "ubicacion", ""),
        ("Preparado por", "preparado_por", ""),
        ("Cargo / Especialidad", "cargo", "Consultor en Ingeniería Eléctrica"),
        ("Matrícula profesional", "matricula", ""),
        ("Fecha de emisión", "fecha_emision", datetime.now().strftime("%d/%m/%Y")),
        ("Revisión", "revision", "Rev. 0"),
        ("Estado del documento", "estado", "Emitido Para Revisión por el Cliente"),
        ("Software utilizado", "software", "ETAP"),
        ("Método de solución", "metodo", "Newton-Raphson")
    ]

    widgets = {}

    for row_idx, (label, key, default) in enumerate(fields):
        ttk.Label(frame, text=label + ":", font=("Segoe UI", 9, "bold")).grid(
            row=row_idx, column=0, padx=8, pady=6, sticky="w"
        )

        entry = ttk.Entry(frame, width=68)
        entry.grid(row=row_idx, column=1, padx=8, pady=6, sticky="ew")
        entry.insert(0, default)

        widgets[key] = entry

    frame.columnconfigure(1, weight=1)

    ttk.Label(
        frame,
        text="Descripción breve del sistema / alcance:",
        font=("Segoe UI", 9, "bold")
    ).grid(row=len(fields), column=0, padx=8, pady=6, sticky="nw")

    txt_alcance = tk.Text(frame, height=7, width=68, font=("Segoe UI", 9), wrap="word")
    txt_alcance.grid(row=len(fields), column=1, padx=8, pady=6, sticky="ew")
    txt_alcance.insert(
        "1.0",
        "El presente informe documenta el estudio de flujo de potencia desarrollado para evaluar el comportamiento del sistema eléctrico en estado estable, incluyendo perfiles de tensión, cargabilidad de equipos y comparación de escenarios operativos."
    )

    def accept():
        for label, key, default in fields:
            data[key] = widgets[key].get().strip()

        data["alcance"] = txt_alcance.get("1.0", "end").strip()

        if data["titulo_informe"] == "":
            messagebox.showerror("Campo requerido", "El título del informe es obligatorio.")
            return

        if data["nombre_proyecto"] == "":
            messagebox.showerror("Campo requerido", "El nombre del proyecto es obligatorio.")
            return

        root.destroy()

    def cancel():
        cancelled["value"] = True
        root.destroy()

    ttk.Button(button_frame, text="Aceptar", command=accept, width=18).pack(side="right", padx=8)
    ttk.Button(button_frame, text="Cancelar", command=cancel, width=18).pack(side="right", padx=8)

    root.bind("<Return>", lambda event: accept())
    root.bind("<Escape>", lambda event: cancel())

    root.mainloop()

    if cancelled.get("value", False):
        raise Exception("Operación cancelada por el usuario.")

    return data


def ask_acceptance_criteria():
    global MIN_VOLTAGE_PERCENT
    global MAX_VOLTAGE_PERCENT
    global MAX_LOADING_PERCENT

    selected = {}

    root = tk.Tk()
    root.title("Criterios de aceptación")
    center_window(root, 620, 360)
    root.resizable(True, True)

    ttk.Label(
        root,
        text="Criterios de aceptación del estudio",
        font=("Segoe UI", 13, "bold")
    ).pack(pady=16)

    ttk.Label(
        root,
        text="Ingrese los límites en porcentaje. Ejemplo: 95, 105 y 100.",
        font=("Segoe UI", 9)
    ).pack(pady=3)

    frame = ttk.Frame(root)
    frame.pack(padx=20, pady=16, fill="x")

    ttk.Label(frame, text="Tensión mínima aceptable (%):", font=("Segoe UI", 9)).grid(row=0, column=0, padx=8, pady=8, sticky="w")
    min_var = tk.StringVar(value=str(DEFAULT_MIN_VOLTAGE_PERCENT))
    ttk.Entry(frame, textvariable=min_var, width=18).grid(row=0, column=1, padx=8, pady=8)

    ttk.Label(frame, text="Tensión máxima aceptable (%):", font=("Segoe UI", 9)).grid(row=1, column=0, padx=8, pady=8, sticky="w")
    max_var = tk.StringVar(value=str(DEFAULT_MAX_VOLTAGE_PERCENT))
    ttk.Entry(frame, textvariable=max_var, width=18).grid(row=1, column=1, padx=8, pady=8)

    ttk.Label(frame, text="Cargabilidad máxima aceptable (%):", font=("Segoe UI", 9)).grid(row=2, column=0, padx=8, pady=8, sticky="w")
    loading_var = tk.StringVar(value=str(DEFAULT_MAX_LOADING_PERCENT))
    ttk.Entry(frame, textvariable=loading_var, width=18).grid(row=2, column=1, padx=8, pady=8)

    def accept():
        min_v = parse_float(min_var.get(), DEFAULT_MIN_VOLTAGE_PERCENT)
        max_v = parse_float(max_var.get(), DEFAULT_MAX_VOLTAGE_PERCENT)
        max_loading = parse_float(loading_var.get(), DEFAULT_MAX_LOADING_PERCENT)

        if min_v <= 0 or max_v <= 0 or max_loading <= 0:
            messagebox.showerror("Valores inválidos", "Todos los valores deben ser mayores que cero.")
            return

        if min_v >= max_v:
            messagebox.showerror("Valores inválidos", "La tensión mínima debe ser menor que la tensión máxima.")
            return

        selected["min_v"] = min_v
        selected["max_v"] = max_v
        selected["max_loading"] = max_loading
        root.destroy()

    def cancel():
        selected["cancelled"] = True
        root.destroy()

    button_frame = ttk.Frame(root)
    button_frame.pack(pady=12)

    ttk.Button(button_frame, text="Aceptar criterios", command=accept, width=20).grid(row=0, column=0, padx=8)
    ttk.Button(button_frame, text="Cancelar", command=cancel, width=16).grid(row=0, column=1, padx=8)

    root.bind("<Return>", lambda event: accept())
    root.bind("<Escape>", lambda event: cancel())

    root.mainloop()

    if selected.get("cancelled", False):
        raise Exception("Operación cancelada por el usuario.")

    MIN_VOLTAGE_PERCENT = selected["min_v"]
    MAX_VOLTAGE_PERCENT = selected["max_v"]
    MAX_LOADING_PERCENT = selected["max_loading"]

    print("")
    print("Criterios configurados:")
    print(" - Tensión mínima aceptable: " + str(MIN_VOLTAGE_PERCENT) + " %")
    print(" - Tensión máxima aceptable: " + str(MAX_VOLTAGE_PERCENT) + " %")
    print(" - Cargabilidad máxima aceptable: " + str(MAX_LOADING_PERCENT) + " %")


def ask_search_folder():
    root = tk.Tk()
    root.withdraw()

    messagebox.showinfo(
        "Carpeta del proyecto",
        "Seleccione la carpeta del proyecto ETAP activo. "
        "Esta carpeta se usará para listar únicamente los reportes .LF1S de este proyecto."
    )

    folder = filedialog.askdirectory(
        title="Seleccione la carpeta del proyecto ETAP activo"
    )

    root.destroy()

    if folder is None:
        folder = ""

    return folder


def edit_options_window(options):
    edited = {}

    root = tk.Tk()
    root.title("Opciones disponibles para flujo de carga ETAP")
    center_window(root, 980, 720)
    root.resizable(True, True)

    ttk.Label(
        root,
        text="Revise o complete las opciones disponibles",
        font=("Segoe UI", 13, "bold")
    ).pack(pady=12)

    ttk.Label(
        root,
        text=(
            "Si alguna lista no muestra todas las opciones del proyecto, agréguelas aquí "
            "una por línea. Luego aparecerán en los desplegables."
        ),
        font=("Segoe UI", 9)
    ).pack(pady=2)

    container = ttk.Frame(root)
    container.pack(fill="both", expand=True, padx=14, pady=10)

    fields = [
        ("Revisiones", "revisions"),
        ("Configuraciones", "configurations"),
        ("Presentaciones", "presentations"),
        ("Casos de estudio de flujo de carga", "study_cases"),
        ("Reportes de salida de flujo de carga", "output_reports")
    ]

    text_widgets = {}

    for idx, (label, key) in enumerate(fields):
        frame = ttk.LabelFrame(container, text=label)
        frame.grid(row=idx // 2, column=idx % 2, padx=8, pady=8, sticky="nsew")

        txt = tk.Text(frame, height=8, width=45, font=("Segoe UI", 9))
        txt.pack(fill="both", expand=True, padx=6, pady=6)

        values = options.get(key, [])
        txt.insert("1.0", "\n".join(values))

        text_widgets[key] = txt

    container.columnconfigure(0, weight=1)
    container.columnconfigure(1, weight=1)
    container.rowconfigure(0, weight=1)
    container.rowconfigure(1, weight=1)
    container.rowconfigure(2, weight=1)

    def accept():
        for key, txt in text_widgets.items():
            text = txt.get("1.0", "end").strip()
            values = [line.strip() for line in text.splitlines() if line.strip()]
            edited[key] = filter_options(values)

        root.destroy()

    def cancel():
        edited["cancelled"] = True
        root.destroy()

    button_frame = ttk.Frame(root)
    button_frame.pack(pady=10)

    ttk.Button(button_frame, text="Aceptar opciones", command=accept, width=20).grid(row=0, column=0, padx=8)
    ttk.Button(button_frame, text="Cancelar", command=cancel, width=16).grid(row=0, column=1, padx=8)

    root.mainloop()

    if edited.get("cancelled", False):
        raise Exception("Operación cancelada por el usuario.")

    for key in ["revisions", "configurations", "presentations", "study_cases", "output_reports"]:
        if key not in edited or not edited[key]:
            edited[key] = options.get(key, [])

    return edited


def ask_number_of_scenarios():
    selected = {}

    root = tk.Tk()
    root.title("Cantidad de escenarios")
    center_window(root, 500, 260)
    root.resizable(True, True)

    ttk.Label(
        root,
        text="¿Cuántos escenarios desea analizar?",
        font=("Segoe UI", 12, "bold")
    ).pack(pady=18)

    ttk.Label(
        root,
        text="Seleccione una cantidad entre 1 y 10.",
        font=("Segoe UI", 9)
    ).pack(pady=4)

    value_var = tk.IntVar(value=2)

    spin = ttk.Spinbox(
        root,
        from_=1,
        to=10,
        textvariable=value_var,
        width=8,
        justify="center"
    )
    spin.pack(pady=10)

    def accept():
        try:
            value = int(value_var.get())
        except Exception:
            messagebox.showerror("Valor inválido", "Ingrese un número válido.")
            return

        if value < 1 or value > 10:
            messagebox.showerror("Valor inválido", "Seleccione una cantidad entre 1 y 10.")
            return

        selected["count"] = value
        root.destroy()

    def cancel():
        selected["cancelled"] = True
        root.destroy()

    button_frame = ttk.Frame(root)
    button_frame.pack(pady=15)

    ttk.Button(button_frame, text="Continuar", command=accept, width=18).grid(row=0, column=0, padx=8)
    ttk.Button(button_frame, text="Cancelar", command=cancel, width=16).grid(row=0, column=1, padx=8)

    root.mainloop()

    if selected.get("cancelled", False):
        raise Exception("Operación cancelada por el usuario.")

    if "count" not in selected:
        raise Exception("No se seleccionó la cantidad de escenarios.")

    return selected["count"]


def make_combo(parent, values, default_value="", width=42):
    cb = ttk.Combobox(parent, values=values, state="normal", width=width)

    if default_value:
        cb.set(default_value)
    elif values:
        cb.set(values[0])

    return cb


def choose_single_scenario(index, options):
    selected = {}

    root = tk.Tk()
    root.title("Escenario " + str(index + 1))
    center_window(root, 780, 560)
    root.resizable(True, True)

    ttk.Label(
        root,
        text="Configuración del escenario " + str(index + 1),
        font=("Segoe UI", 13, "bold")
    ).pack(pady=15)

    frame = ttk.Frame(root)
    frame.pack(padx=20, pady=10, fill="both", expand=True)

    suggested_names = [
        "LF-Report",
        "LF-Scn-100A",
        "Escenario-3",
        "Escenario-4"
    ]

    suggested_reports = [
        "LF-Report",
        "LF-Scn-100A",
        "LF-Report-3",
        "LF-Report-4"
    ]

    fields = [
        ("Nombre visible", "name", suggested_names),
        ("Revisión", "revisionName", options.get("revisions", [])),
        ("Configuración", "configName", options.get("configurations", [])),
        ("Caso de estudio", "studyCase", options.get("study_cases", [])),
        ("Reporte de salida", "outputReport", options.get("output_reports", [])),
    ]

    widgets = {}

    for row_idx, (label, key, values) in enumerate(fields):
        ttk.Label(frame, text=label, font=("Segoe UI", 9, "bold")).grid(
            row=row_idx, column=0, padx=10, pady=10, sticky="w"
        )

        if key == "name":
            default_value = suggested_names[index] if index < len(suggested_names) else "Escenario-" + str(index + 1)
            all_values = merge_options(values, options.get("output_reports", []), options.get("study_cases", []))
        elif key == "outputReport":
            default_value = suggested_reports[index] if index < len(suggested_reports) else "LF-Report-" + str(index + 1)
            all_values = merge_options(values, suggested_reports)
        elif key == "studyCase":
            all_values = values
            if "LF Report" in values:
                default_value = "LF Report"
            elif values:
                default_value = values[0]
            else:
                default_value = ""
        else:
            all_values = values
            default_value = values[0] if values else ""

        cb = make_combo(frame, all_values, default_value=default_value)
        cb.grid(row=row_idx, column=1, padx=10, pady=10, sticky="w")

        widgets[key] = cb

    note = ttk.Label(
        root,
        text=(
            "Los desplegables son editables: si una opción existe en ETAP pero no aparece, escríbala exactamente. "
            "Esto evita quedar limitado a las opciones detectadas automáticamente."
        ),
        font=("Segoe UI", 8),
        foreground="#404040",
        wraplength=620
    )
    note.pack(pady=5)

    def accept():
        scenario = {
            "name": widgets["name"].get().strip(),
            "revisionName": widgets["revisionName"].get().strip(),
            "configName": widgets["configName"].get().strip(),
            # No se solicita por escenario, pero se prueban las presentaciones disponibles del proyecto.
            "presentation": (options.get("presentations", [DEFAULT_PRESENTATION])[0] if options.get("presentations", []) else DEFAULT_PRESENTATION),
            "presentationCandidates": options.get("presentations", [DEFAULT_PRESENTATION]),
            "studyCase": widgets["studyCase"].get().strip(),
            "outputReport": widgets["outputReport"].get().strip(),
            "whatIfCommands": {"Commands": []}
        }

        for key, value in scenario.items():
            if key == "whatIfCommands":
                continue
            if str(value).strip() == "":
                messagebox.showerror("Campo requerido", "Todos los campos deben estar completos.")
                return

        selected["scenario"] = scenario
        root.destroy()

    def cancel():
        selected["cancelled"] = True
        root.destroy()

    button_frame = ttk.Frame(root)
    button_frame.pack(pady=12)

    ttk.Button(button_frame, text="Aceptar escenario", command=accept, width=20).grid(row=0, column=0, padx=8)
    ttk.Button(button_frame, text="Cancelar", command=cancel, width=16).grid(row=0, column=1, padx=8)

    root.mainloop()

    if selected.get("cancelled", False):
        raise Exception("Operación cancelada por el usuario.")

    if "scenario" not in selected:
        raise Exception("No se configuró el escenario.")

    return selected["scenario"]


def select_scenarios_workflow(options):
    number_of_scenarios = ask_number_of_scenarios()

    scenarios = []

    for index in range(number_of_scenarios):
        scenario = choose_single_scenario(index, options)
        scenarios.append(scenario)

    return scenarios


# =============================================================================
# EJECUCION LOAD FLOW
# =============================================================================

def parse_report_path(response):
    if response is None:
        return ""

    try:
        response_dict = json.loads(response)
        report_path = response_dict.get("ReportPath", "")
        if report_path:
            return report_path
    except Exception:
        pass

    text = str(response).strip()

    try:
        match = re.search(r'"?ReportPath"?\s*:\s*"?([^"}]+)"?', text)
        if match:
            return match.group(1).strip()
    except Exception:
        pass

    try:
        match = re.search(r'([A-Za-z]:\\[^"<>\']+\.LF1S)', text, flags=re.IGNORECASE)
        if match:
            return match.group(1).strip()
    except Exception:
        pass

    return ""


def find_recent_output_report(output_report_name, search_roots):
    """
    Busca el reporte .LF1S por nombre. Primero usa la caché de la carpeta del
    proyecto y luego escanea las rutas de búsqueda.
    """
    candidates = []
    target_file = output_report_name + ".LF1S"

    try:
        if output_report_name in REPORT_NAME_TO_PATH and os.path.exists(REPORT_NAME_TO_PATH[output_report_name]):
            candidates.append(REPORT_NAME_TO_PATH[output_report_name])
    except Exception:
        pass

    for search_root in search_roots:
        if not search_root or not os.path.exists(search_root):
            continue

        for root, dirs, files in os.walk(search_root):
            for file_name in files:
                if file_name.lower() == target_file.lower():
                    full_path = os.path.join(root, file_name)
                    candidates.append(full_path)

    if not candidates:
        return ""

    candidates = list(dict.fromkeys(candidates))
    candidates.sort(key=os.path.getmtime, reverse=True)
    return candidates[0]


def run_load_flow(e, scenario, search_roots):
    print("")
    print("Ejecutando escenario:", scenario["name"])
    print("Revisión:", scenario["revisionName"])
    print("Configuración:", scenario["configName"])
    print("Presentación inicial:", scenario["presentation"])
    print("Caso de estudio ETAP:", scenario["studyCase"])
    print("Reporte de salida:", scenario["outputReport"])

    existing_report_path = find_recent_output_report(scenario["outputReport"], search_roots)

    # Candidatos de caso de estudio: primero el seleccionado.
    raw_study_cases = []
    if scenario.get("studyCase", "").strip():
        raw_study_cases.append(scenario["studyCase"].strip())

    for candidate in COMMON_LF_STUDY_CASE_CANDIDATES:
        if candidate not in raw_study_cases:
            raw_study_cases.append(candidate)

    for candidate in FALLBACK_STUDY_CASES:
        if candidate not in raw_study_cases:
            raw_study_cases.append(candidate)

    # Validar candidatos con getstudycase, siguiendo el ejemplo de ETAP.
    valid_study_cases = []
    for candidate in raw_study_cases:
        data = safe_get_study_case(e, candidate)
        if data is not None and is_load_flow_study_case_data(data):
            valid_study_cases.append(candidate)

    if not valid_study_cases:
        valid_study_cases = raw_study_cases

    study_case_candidates = filter_options(valid_study_cases)

    # Candidatos de presentación: primero los obtenidos del proyecto/ventana.
    presentation_candidates = []
    for p in scenario.get("presentationCandidates", []):
        if p and p not in presentation_candidates:
            presentation_candidates.append(p)

    if scenario.get("presentation", "") and scenario["presentation"] not in presentation_candidates:
        presentation_candidates.insert(0, scenario["presentation"])

    for p in COMMON_PRESENTATION_CANDIDATES:
        if p not in presentation_candidates:
            presentation_candidates.append(p)

    presentation_candidates = filter_options(presentation_candidates)

    last_response = ""
    last_error = ""
    tried_pairs = []

    for study_case in study_case_candidates:
        for presentation in presentation_candidates:
            try:
                tried_pairs.append(study_case + " / " + presentation)
                print("Intentando ejecutar flujo de carga con caso:", study_case, "| presentación:", presentation)

                response = e.studies.runLF(
                    scenario["revisionName"],
                    scenario["configName"],
                    study_case,
                    presentation,
                    scenario["outputReport"],
                    GET_ONLINE_DATA,
                    ONLINE_CONFIG_ONLY,
                    TIMEOUT_SECS,
                    scenario["whatIfCommands"]
                )

                last_response = str(response)

                # Si ETAP devuelve ErrorMessage, no intentar parsear como ruta válida.
                try:
                    parsed_resp = json.loads(last_response)
                    if isinstance(parsed_resp, dict) and parsed_resp.get("ErrorMessage"):
                        print("ETAP respondió error:", parsed_resp.get("ErrorMessage"))
                        continue
                except Exception:
                    pass

                report_path = parse_report_path(response)

                print("Respuesta ETAP:", str(response))
                print("Archivo de resultados devuelto:", report_path)

                if report_path and os.path.exists(report_path):
                    scenario["studyCase"] = study_case
                    scenario["presentation"] = presentation
                    return report_path

                fallback_path = find_recent_output_report(scenario["outputReport"], search_roots)
                if fallback_path and os.path.exists(fallback_path):
                    print("ETAP no devolvió ReportPath, pero se encontró el reporte:", fallback_path)
                    scenario["studyCase"] = study_case
                    scenario["presentation"] = presentation
                    return fallback_path

            except Exception as ex:
                last_error = str(ex)
                print("No fue posible ejecutar con", study_case, "/", presentation, ":", last_error)

    if existing_report_path and os.path.exists(existing_report_path):
        print("")
        print("Advertencia: ETAP no devolvió ReportPath al ejecutar el estudio.")
        print("Se usará el reporte existente encontrado en la carpeta del proyecto:")
        print(existing_report_path)
        return existing_report_path

    raise Exception(
        "ETAP no devolvió la ruta del reporte para el escenario: "
        + scenario["name"]
        + "\nCaso de estudio seleccionado: "
        + scenario["studyCase"]
        + "\nPresentación inicial: "
        + scenario["presentation"]
        + "\nCombinaciones probadas: "
        + "; ".join(tried_pairs)
        + "\nReporte de salida usado: "
        + scenario["outputReport"]
        + "\nÚltima respuesta ETAP: "
        + str(last_response)
        + "\nÚltimo error: "
        + str(last_error)
        + "\n\nRevise que el caso de estudio, presentación y reporte de salida existan exactamente en el proyecto activo."
    )


# =============================================================================
# SQLITE / LFR
# =============================================================================

def connect_sqlite(db_file):
    return sqlite3.connect(db_file)


def list_tables(conn):
    query = """
    SELECT name
    FROM sqlite_master
    WHERE type='table'
    ORDER BY name;
    """

    df = pd.read_sql_query(query, conn)

    if df.empty:
        return []

    return df["name"].tolist()


def read_table(conn, table_name):
    query = 'SELECT * FROM "' + table_name + '"'
    return pd.read_sql_query(query, conn)


def get_lfr_table(report_path):
    conn = connect_sqlite(report_path)
    tables = list_tables(conn)

    if "LFR" not in tables:
        conn.close()

        raise Exception(
            "No se encontró la tabla LFR en el archivo:\n"
            + report_path
            + "\n\nTablas encontradas: "
            + ", ".join(tables)
        )

    lfr = read_table(conn, "LFR")
    conn.close()

    return lfr


# =============================================================================
# ANALISIS TECNICO
# =============================================================================

def classify_voltage_percent(value):
    number = to_number(value)

    if number is None:
        return "Sin dato"

    if number < MIN_VOLTAGE_PERCENT:
        return "Baja tensión"

    if number > MAX_VOLTAGE_PERCENT:
        return "Sobretensión"

    return "Aceptable"


def classify_loading_percent(value):
    number = to_number(value)

    if number is None:
        return "Sin dato"

    if number > MAX_LOADING_PERCENT:
        return "Sobrecargado"

    return "Aceptable"


def build_bus_voltage_table(scenario_name, lfr):
    required_columns = ["IDFrom", "VoltMag", "VoltAng", "kV", "TYPE"]

    for col in required_columns:
        if col not in lfr.columns:
            raise Exception("La columna requerida no existe en LFR: " + col)

    df = lfr.copy()

    df = df[
        (pd.to_numeric(df["TYPE"], errors="coerce") != 0)
        & (pd.to_numeric(df["kV"], errors="coerce") != 0)
    ].copy()

    result = df[["IDFrom", "kV", "VoltMag", "VoltAng"]].drop_duplicates().copy()

    result = result.rename(columns={
        "IDFrom": "Barra",
        "kV": "kV nominal",
        "VoltMag": "Tensión (%)",
        "VoltAng": "Ángulo (deg)"
    })

    result.insert(0, "Escenario", scenario_name)

    result["Tensión (%)"] = pd.to_numeric(result["Tensión (%)"], errors="coerce")
    result["Ángulo (deg)"] = pd.to_numeric(result["Ángulo (deg)"], errors="coerce")
    result["Tensión (pu)"] = result["Tensión (%)"] / 100.0
    result["Estado"] = result["Tensión (%)"].apply(classify_voltage_percent)

    result = result.sort_values(["Escenario", "Tensión (%)"], ascending=[True, True])

    return result


def build_equipment_table(scenario_name, lfr):
    df = lfr.copy()

    result = pd.DataFrame()
    result["Escenario"] = [scenario_name] * len(df)

    if "IDFrom" in df.columns:
        result["Desde"] = df["IDFrom"]
    else:
        result["Desde"] = ""

    if "IDTo" in df.columns:
        result["Hasta"] = df["IDTo"]
    else:
        result["Hasta"] = ""

    if "TYPE" in df.columns:
        result["TYPE"] = df["TYPE"]
    else:
        result["TYPE"] = ""

    if "kV" in df.columns:
        result["kV"] = df["kV"]

    possible_cols = [
        "MW", "Mvar", "MVA", "Amp", "AmpFrom", "AmpTo",
        "PF", "Loading", "LoadingPercent", "PercentLoading",
        "VoltMag", "VoltAng"
    ]

    for col in possible_cols:
        if col in df.columns:
            result[col] = df[col]

    loading_col = None

    for col in result.columns:
        col_norm = str(col).lower().replace("_", "").replace(" ", "")

        if "loading" in col_norm:
            loading_col = col
            break

    if loading_col is not None:
        result["Cargabilidad (%)"] = pd.to_numeric(result[loading_col], errors="coerce")
        result["Estado Cargabilidad"] = result["Cargabilidad (%)"].apply(classify_loading_percent)
    else:
        result["Cargabilidad (%)"] = ""
        result["Estado Cargabilidad"] = "Informativo"

    return result


def create_scenario_diagnosis(scenario_name, bus_count, min_v, max_v, low_v, high_v, overloaded):
    if bus_count == 0:
        return "No se detectaron barras para el escenario " + scenario_name + "."

    messages = []

    if low_v == 0 and high_v == 0:
        messages.append("Los perfiles de tensión se encuentran dentro del criterio configurado.")

    if low_v > 0:
        messages.append(
            ("Se identificó 1 barra con baja tensión. Revisar caídas de tensión, taps y compensación reactiva." if low_v == 1 else "Se identificaron " + str(low_v) + " barras con baja tensión. Revisar caídas de tensión, taps y compensación reactiva.")
        )

    if high_v > 0:
        messages.append(
            ("Se identificó 1 barra con sobretensión. Revisar regulación de tensión y taps." if high_v == 1 else "Se identificaron " + str(high_v) + " barras con sobretensión. Revisar regulación de tensión y taps.")
        )

    if overloaded > 0:
        messages.append(
            ("Se identificó 1 equipo sobrecargado. Revisar capacidad nominal y redistribución de carga." if overloaded == 1 else "Se identificaron " + str(overloaded) + " equipos sobrecargados. Revisar capacidad nominal y redistribución de carga.")
        )

    if len(messages) == 0:
        messages.append("No se identificaron desviaciones relevantes.")

    return " ".join(messages)


def create_scenario_summary(scenarios, bus_df, equipment_df):
    rows = []

    for scenario in scenarios:
        scenario_name = scenario["name"]

        buses = bus_df[bus_df["Escenario"] == scenario_name].copy()
        equipment = equipment_df[equipment_df["Escenario"] == scenario_name].copy()

        if not buses.empty:
            min_v = buses["Tensión (%)"].min()
            max_v = buses["Tensión (%)"].max()
            avg_v = buses["Tensión (%)"].mean()

            low_v = len(buses[buses["Estado"] == "Baja tensión"])
            high_v = len(buses[buses["Estado"] == "Sobretensión"])
            ok_v = len(buses[buses["Estado"] == "Aceptable"])
        else:
            min_v = ""
            max_v = ""
            avg_v = ""
            low_v = 0
            high_v = 0
            ok_v = 0

        if not equipment.empty and "Estado Cargabilidad" in equipment.columns:
            overloaded = len(equipment[equipment["Estado Cargabilidad"] == "Sobrecargado"])
        else:
            overloaded = 0

        rows.append({
            "Escenario": scenario_name,
            "Revisión": scenario["revisionName"],
            "Configuración": scenario["configName"],
            "Presentación": scenario["presentation"],
            "Caso de estudio ETAP": scenario["studyCase"],
            "Reporte de salida": scenario["outputReport"],
            "Barras evaluadas": len(buses),
            "Tensión mínima (%)": min_v,
            "Tensión máxima (%)": max_v,
            "Tensión promedio (%)": avg_v,
            "Barras con baja tensión": low_v,
            "Barras con sobretensión": high_v,
            "Barras aceptables": ok_v,
            "Equipos evaluados": len(equipment),
            "Equipos sobrecargados": overloaded,
            "Diagnóstico": create_scenario_diagnosis(
                scenario_name,
                len(buses),
                min_v,
                max_v,
                low_v,
                high_v,
                overloaded
            )
        })

    result = pd.DataFrame(rows)

    integer_cols = [
        "Barras evaluadas",
        "Barras con baja tensión",
        "Barras con sobretensión",
        "Barras aceptables",
        "Equipos evaluados",
        "Equipos sobrecargados"
    ]

    for col in integer_cols:
        if col in result.columns:
            result[col] = pd.to_numeric(result[col], errors="coerce").fillna(0).astype(int)

    return result


def create_comparison_summary(summary_df):
    if len(summary_df) < 2:
        return pd.DataFrame([["No hay suficientes escenarios con resultados para comparar."]], columns=["Resultado"])

    # Ignorar escenarios sin barras evaluadas, porque no tienen métricas válidas.
    valid = summary_df.copy()
    valid["Barras evaluadas num"] = pd.to_numeric(valid["Barras evaluadas"], errors="coerce").fillna(0)
    valid = valid[valid["Barras evaluadas num"] > 0].copy()

    if len(valid) < 2:
        return pd.DataFrame([["No hay suficientes escenarios con resultados válidos para comparar."]], columns=["Resultado"])

    rows = []
    base = valid.iloc[0]

    for index in range(1, len(valid)):
        comp = valid.iloc[index]

        def diff(col):
            try:
                return float(comp[col]) - float(base[col])
            except Exception:
                return ""

        comparison_items = [
            ("Tensión mínima (%)", "Diferencia de tensión mínima entre escenarios."),
            ("Tensión máxima (%)", "Diferencia de tensión máxima entre escenarios."),
            ("Tensión promedio (%)", "Diferencia de tensión promedio entre escenarios."),
            ("Barras con baja tensión", "Incremento o reducción de barras por debajo del límite."),
            ("Barras con sobretensión", "Incremento o reducción de barras por encima del límite."),
            ("Equipos sobrecargados", "Incremento o reducción de equipos sobrecargados.")
        ]

        for variable, interpretation in comparison_items:
            rows.append({
                "Variable": variable,
                "Escenario base": base["Escenario"],
                "Valor base": base[variable],
                "Escenario comparado": comp["Escenario"],
                "Valor comparado": comp[variable],
                "Diferencia": diff(variable),
                "Interpretación": interpretation
            })

    return pd.DataFrame(rows)


def create_bus_delta(scenarios, bus_df):
    if len(scenarios) < 2:
        return pd.DataFrame([["Se requieren dos escenarios con resultados para comparar barras."]], columns=["Resultado"])

    if bus_df is None or bus_df.empty:
        return pd.DataFrame([["No hay datos suficientes para comparar barras."]], columns=["Resultado"])

    valid_names = []
    for scenario in scenarios:
        name = scenario["name"]
        if not bus_df[bus_df["Escenario"] == name].empty:
            valid_names.append(name)

    if len(valid_names) < 2:
        return pd.DataFrame([["No hay suficientes escenarios con barras calculadas para comparar."]], columns=["Resultado"])

    base_name = valid_names[0]
    base = bus_df[bus_df["Escenario"] == base_name].copy()

    final_frames = []

    for comp_name in valid_names[1:]:
        comp = bus_df[bus_df["Escenario"] == comp_name].copy()

        base_small = base[["Barra", "Tensión (%)", "Ángulo (deg)", "Estado"]].copy()
        comp_small = comp[["Barra", "Tensión (%)", "Ángulo (deg)", "Estado"]].copy()

        base_small = base_small.rename(columns={
            "Tensión (%)": "Tensión " + base_name + " (%)",
            "Ángulo (deg)": "Ángulo " + base_name + " (deg)",
            "Estado": "Estado " + base_name
        })

        comp_small = comp_small.rename(columns={
            "Tensión (%)": "Tensión " + comp_name + " (%)",
            "Ángulo (deg)": "Ángulo " + comp_name + " (deg)",
            "Estado": "Estado " + comp_name
        })

        merged = pd.merge(base_small, comp_small, on="Barra", how="inner")

        if merged.empty:
            continue

        merged["Escenario base"] = base_name
        merged["Escenario comparado"] = comp_name

        merged["Delta tensión (%)"] = (
            pd.to_numeric(merged["Tensión " + comp_name + " (%)"], errors="coerce")
            - pd.to_numeric(merged["Tensión " + base_name + " (%)"], errors="coerce")
        )

        merged["Delta ángulo (deg)"] = (
            pd.to_numeric(merged["Ángulo " + comp_name + " (deg)"], errors="coerce")
            - pd.to_numeric(merged["Ángulo " + base_name + " (deg)"], errors="coerce")
        )

        def interpret(row):
            delta = to_number(row["Delta tensión (%)"])

            if delta is None:
                return "Sin interpretación."

            if delta < 0:
                return "La tensión disminuye en " + comp_name + " respecto a " + base_name + "."

            if delta > 0:
                return "La tensión aumenta en " + comp_name + " respecto a " + base_name + "."

            return "No se observa cambio de tensión."

        merged["Interpretación"] = merged.apply(interpret, axis=1)
        merged["Delta abs"] = merged["Delta tensión (%)"].abs()

        final_frames.append(merged)

    if not final_frames:
        return pd.DataFrame([["No se encontraron barras coincidentes entre escenarios con resultados válidos."]], columns=["Resultado"])

    result = pd.concat(final_frames, ignore_index=True, sort=False)
    result = result.sort_values("Delta abs", ascending=False)

    return result


def create_findings(bus_df, equipment_df):
    rows = []

    deviated_buses = bus_df[bus_df["Estado"].isin(["Baja tensión", "Sobretensión"])].copy()

    for _, row in deviated_buses.iterrows():
        status = row["Estado"]
        scenario = row["Escenario"]
        bus = row["Barra"]
        voltage = row["Tensión (%)"]

        if status == "Baja tensión":
            recommendation = (
                "Revisar caída de tensión, posición de taps, cargabilidad aguas abajo "
                "y necesidad de compensación reactiva."
            )
        else:
            recommendation = (
                "Revisar ajuste de taps, regulación de tensión y aportes de potencia reactiva."
            )

        rows.append({
            "Escenario": scenario,
            "Tipo": status,
            "Elemento": bus,
            "Valor": voltage,
            "Unidad": "%",
            "Hallazgo": "El elemento presenta " + status.lower() + ".",
            "Recomendación": recommendation
        })

    if "Estado Cargabilidad" in equipment_df.columns:
        overloaded = equipment_df[equipment_df["Estado Cargabilidad"] == "Sobrecargado"].copy()

        for _, row in overloaded.iterrows():
            rows.append({
                "Escenario": row["Escenario"],
                "Tipo": "Sobrecarga",
                "Elemento": str(row.get("Desde", "")) + " - " + str(row.get("Hasta", "")),
                "Valor": row.get("Cargabilidad (%)", ""),
                "Unidad": "%",
                "Hallazgo": "El equipo presenta sobrecarga.",
                "Recomendación": (
                    "Verificar capacidad nominal, corriente operativa, redistribución de carga "
                    "o necesidad de refuerzo."
                )
            })

    if len(rows) == 0:
        return pd.DataFrame([["No se identificaron desviaciones con los criterios configurados."]], columns=["Resultado"])

    return pd.DataFrame(rows)


def create_conclusions(summary_df):
    conclusions = []

    if summary_df.empty:
        return ["No se encontraron datos suficientes para emitir conclusiones sobre el estudio de flujo de carga."]

    total_low = int(summary_df["Barras con baja tensión"].sum())
    total_high = int(summary_df["Barras con sobretensión"].sum())
    total_overloaded = int(summary_df["Equipos sobrecargados"].sum())

    if total_low == 0 and total_high == 0:
        conclusions.append("Los perfiles de tensión de las barras evaluadas se encuentran dentro de los límites configurados para los escenarios analizados.")
    else:
        if total_low > 0:
            conclusions.append("Se identificaron barras con tensión inferior al límite mínimo configurado, por lo que se recomienda revisar caídas de tensión, taps de transformadores y condiciones de carga.")
        if total_high > 0:
            conclusions.append("Se identificaron barras con tensión superior al límite máximo configurado, por lo que se recomienda revisar regulación de tensión, taps y aportes de potencia reactiva.")

    if total_overloaded == 0:
        conclusions.append("No se identificaron equipos sobrecargados con base en la información de cargabilidad disponible en los resultados procesados.")
    else:
        conclusions.append("Se identificaron equipos con cargabilidad superior al criterio configurado; se recomienda revisar capacidad nominal, condiciones operativas y posibles refuerzos.")

    conclusions.append("Los resultados deben interpretarse considerando la calidad del modelo eléctrico, los datos de entrada, las condiciones operativas seleccionadas y las premisas utilizadas para cada escenario.")

    return conclusions



# =============================================================================
# EXTRACCION Y RESUMEN DE ELEMENTOS DEL MODELO
# =============================================================================

def json_to_dataframe(value):
    if value is None:
        return pd.DataFrame()
    if isinstance(value, str):
        text = value.strip()
        if text == "":
            return pd.DataFrame()
        try:
            value = json.loads(text)
        except Exception:
            return pd.DataFrame({"Valor": [text]})
    if isinstance(value, dict):
        for key in ["Data", "data", "Items", "items", "Result", "result", "Rows", "rows", "Elements", "elements"]:
            if key in value:
                return json_to_dataframe(value[key])
        return pd.DataFrame([value])
    if isinstance(value, list):
        if len(value) == 0:
            return pd.DataFrame()
        if all(isinstance(item, dict) for item in value):
            return pd.DataFrame(value)
        return pd.DataFrame({"Valor": value})
    return pd.DataFrame({"Valor": [str(value)]})


def safe_projectdata_call(e, method_name, args=None):
    try:
        if args is None:
            args = []
        pd_obj = getattr(e, "projectdata", None)
        if pd_obj is None:
            pd_obj = getattr(e, "projectData", None)
        if pd_obj is None or not hasattr(pd_obj, method_name):
            return pd.DataFrame()
        method = getattr(pd_obj, method_name)
        if not callable(method):
            return json_to_dataframe(method)
        response = method(*args)
        return json_to_dataframe(response)
    except Exception:
        return pd.DataFrame()


def get_projectdata_object(e):
    """
    Devuelve el objeto ProjectData de ETAP considerando diferencias de mayúsculas
    entre versiones de la API.
    """
    pd_obj = getattr(e, "projectdata", None)
    if pd_obj is None:
        pd_obj = getattr(e, "projectData", None)
    return pd_obj


# Caché de consultas ProjectData para evitar llamadas repetidas al DataHub.
_PROJECTDATA_CACHE = {
    "types": None,
    "names": {},
    "all_data": {},
    "props": {},
}

# Caché de reportes .LF1S encontrados en la carpeta del proyecto.
REPORT_NAME_TO_PATH = {}




def safe_get_element_names(e, element_type):
    """
    Consulta e.projectdata.getelementnames(element_type) con caché.
    """
    cache_key = str(element_type)

    if cache_key in _PROJECTDATA_CACHE["names"]:
        return _PROJECTDATA_CACHE["names"][cache_key]

    try:
        pd_obj = get_projectdata_object(e)
        if pd_obj is None or not hasattr(pd_obj, "getelementnames"):
            _PROJECTDATA_CACHE["names"][cache_key] = []
            return []

        response = pd_obj.getelementnames(element_type)
        values = filter_options(flatten_strings(response))
        _PROJECTDATA_CACHE["names"][cache_key] = values
        return values
    except Exception:
        _PROJECTDATA_CACHE["names"][cache_key] = []
        return []


def element_names_dataframe(e, element_types, category_label):
    """
    Intenta extraer nombres de elementos del modelo a partir de varios códigos
    posibles de tipo de elemento. La disponibilidad de códigos puede variar entre
    versiones de ETAP y librerías de API.
    """
    rows = []
    used_types = []

    for element_type in element_types:
        names = safe_get_element_names(e, element_type)

        if names:
            used_types.append(element_type)

        for name in names:
            rows.append({
                "ID": name,
                "Tipo consultado": element_type,
                "Categoría": category_label,
                "Fuente": "ProjectData.getelementnames"
            })

    df = pd.DataFrame(rows)

    if df.empty:
        return df, ""

    df = df.drop_duplicates(subset=["ID", "Categoría"]).reset_index(drop=True)

    return df, "ProjectData.getelementnames(" + ", ".join(used_types) + ")"


def enrich_element_names_with_basic_data(e, names_df):
    """
    Conserva compatibilidad con versiones anteriores. La extracción principal
    se hace ahora con build_element_table_from_projectdata().
    """
    if names_df is None or names_df.empty:
        return pd.DataFrame()
    return names_df.copy()


def safe_get_element_types(e):
    """
    Usa ProjectData.getelementtypes() con caché.
    """
    if _PROJECTDATA_CACHE["types"] is not None:
        return _PROJECTDATA_CACHE["types"]

    try:
        pd_obj = get_projectdata_object(e)
        if pd_obj is None or not hasattr(pd_obj, "getelementtypes"):
            _PROJECTDATA_CACHE["types"] = []
            return []

        response = pd_obj.getelementtypes()
        values = filter_options(flatten_strings(response))
        _PROJECTDATA_CACHE["types"] = values
        return values
    except Exception:
        _PROJECTDATA_CACHE["types"] = []
        return []


def safe_get_property_names_xml(e, element_type):
    """
    Obtiene nombres de propiedades mediante getelementpropertynamesxml(element_type) con caché.

    Esta versión es más flexible porque ETAP puede devolver XML con:
    - <string>Field</string>
    - <Name>Field</Name>
    - atributos name / fieldName / propertyName
    - tablas XML con columnas FieldName / PropertyName / Name
    """
    cache_key = str(element_type)

    if cache_key in _PROJECTDATA_CACHE["props"]:
        return _PROJECTDATA_CACHE["props"][cache_key]

    try:
        pd_obj = get_projectdata_object(e)
        if pd_obj is None or not hasattr(pd_obj, "getelementpropertynamesxml"):
            _PROJECTDATA_CACHE["props"][cache_key] = []
            return []

        response = pd_obj.getelementpropertynamesxml(element_type)
        text = str(response).strip()

        names = []

        # 1) Intento JSON.
        try:
            parsed = json.loads(text)
            df = json_to_dataframe(parsed)
            if df is not None and not df.empty:
                for col in df.columns:
                    if normalize_prop_name(col) in ["name", "fieldname", "propertyname", "id"]:
                        names.extend(df[col].dropna().astype(str).tolist())
        except Exception:
            pass

        # 2) Intento XML estructurado.
        if not names:
            try:
                import xml.etree.ElementTree as ET
                root = ET.fromstring(text)

                for node in root.iter():
                    tag = xml_local_name(node.tag)
                    tag_norm = normalize_prop_name(tag)
                    node_text = (node.text or "").strip()

                    # Caso: <string>MW</string> o <Name>MW</Name>
                    if tag_norm in ["string", "name", "fieldname", "propertyname", "columnname", "field", "property"]:
                        if node_text:
                            names.append(node_text)

                    # Caso: atributo Name="MW", FieldName="MW", PropertyName="MW"
                    for ak, av in node.attrib.items():
                        ak_norm = normalize_prop_name(ak)
                        if ak_norm in ["name", "fieldname", "propertyname", "columnname", "field", "property"]:
                            if str(av).strip():
                                names.append(str(av).strip())

                    # Caso: una fila con hijos FieldName/PropertyName
                    row = {}
                    for child in list(node):
                        child_tag = normalize_prop_name(xml_local_name(child.tag))
                        child_text = (child.text or "").strip()
                        if child_text:
                            row[child_tag] = child_text
                    for key in ["fieldname", "propertyname", "name", "columnname"]:
                        if key in row and row[key]:
                            names.append(row[key])
            except Exception:
                pass

        # 3) Regex de respaldo.
        if not names:
            patterns = [
                r"<Name>(.*?)</Name>",
                r"<FieldName>(.*?)</FieldName>",
                r"<PropertyName>(.*?)</PropertyName>",
                r"<string[^>]*>(.*?)</string>",
                r'name="([^"]+)"',
                r'Name="([^"]+)"',
                r'fieldName="([^"]+)"',
                r'FieldName="([^"]+)"',
                r'propertyName="([^"]+)"',
                r'PropertyName="([^"]+)"',
                r'"Name"\s*:\s*"([^"]+)"',
                r'"FieldName"\s*:\s*"([^"]+)"',
                r'"PropertyName"\s*:\s*"([^"]+)"',
            ]

            for pattern in patterns:
                for match in re.findall(pattern, text, flags=re.IGNORECASE):
                    clean = str(match).strip()
                    if clean:
                        names.append(clean)

        # 4) Último respaldo: contenido entre etiquetas, filtrado.
        if not names:
            for item in re.findall(r">([^<>]+)<", text):
                item = item.strip()
                if item and len(item) < 80:
                    names.append(item)

        # Filtrar ruido común.
        cleaned = []
        bad_tokens = [
            "schema", "documentelement", "diffgram", "dataset", "table",
            "newdataset", "xml", "xmlns", "true", "false"
        ]

        for name in names:
            name = str(name).strip()
            if not name:
                continue
            if len(name) > 80:
                continue
            low = name.lower()
            if any(tok == low for tok in bad_tokens):
                continue
            if "<" in name or ">" in name or "http://" in low or "https://" in low:
                continue
            cleaned.append(name)

        values = filter_options(cleaned)
        _PROJECTDATA_CACHE["props"][cache_key] = values
        return values
    except Exception:
        _PROJECTDATA_CACHE["props"][cache_key] = []
        return []


def safe_get_element_prop(e, element_type, element_name, field_name):
    """
    Consulta ProjectData.getelementprop(ElementType, ElementName, FieldName).
    Devuelve vacío si ETAP responde 'Invalid field name' u otro valor no útil.
    """
    try:
        pd_obj = get_projectdata_object(e)
        if pd_obj is None:
            return ""

        if not hasattr(pd_obj, "getelementprop"):
            return ""

        response = pd_obj.getelementprop(element_type, element_name, field_name)

        value = response

        # Algunas versiones devuelven JSON {"Value":"..."}.
        if isinstance(response, str):
            text = response.strip()
            try:
                parsed = json.loads(text)
                if isinstance(parsed, dict):
                    for key in ["Value", "value", "Result", "result"]:
                        if key in parsed:
                            value = parsed[key]
                            break
                    else:
                        value = text
                else:
                    value = text
            except Exception:
                value = text

        elif isinstance(response, dict):
            for key in ["Value", "value", "Result", "result"]:
                if key in response:
                    value = response[key]
                    break

        return clean_etap_value(value)
    except Exception:
        return ""


def normalize_prop_name(name):
    return str(name).lower().replace(" ", "").replace("_", "").replace("-", "").replace(".", "")


def is_invalid_etap_value(value):
    """
    Identifica respuestas no útiles de ETAP API para evitar que aparezcan en el informe:
    Invalid field name, nan, null, None, listas vacías, etc.
    """
    if value is None:
        return True

    text = str(value).strip()

    if text == "":
        return True

    low = text.lower()

    invalid_tokens = [
        "invalid field name",
        "invalid property",
        "invalid field",
        "field not found",
        "property not found",
        "not supported",
        "exception",
        "traceback",
        "error",
        "none",
        "null",
        "nan",
        "[]",
        "{}"
    ]

    for token in invalid_tokens:
        if token in low:
            return True

    return False


def clean_etap_value(value):
    """
    Limpia valores antes de llevarlos a tablas del informe.
    """
    if is_invalid_etap_value(value):
        return ""

    text = str(value).strip()

    # Normalizar booleanos para presentación.
    if text.lower() == "true":
        return "Sí"

    if text.lower() == "false":
        return "No"

    return value



def safe_get_all_element_data(e, element_type):
    """
    Intenta leer información completa de elementos usando ProjectData.getallelementdata.
    Usa caché para no repetir consultas al DataHub.
    """
    cache_key = str(element_type)

    if cache_key in _PROJECTDATA_CACHE["all_data"]:
        return _PROJECTDATA_CACHE["all_data"][cache_key].copy()

    try:
        pd_obj = get_projectdata_object(e)
        if pd_obj is None or not hasattr(pd_obj, "getallelementdata"):
            _PROJECTDATA_CACHE["all_data"][cache_key] = pd.DataFrame()
            return pd.DataFrame()

        response = pd_obj.getallelementdata(element_type)
        if response is None:
            _PROJECTDATA_CACHE["all_data"][cache_key] = pd.DataFrame()
            return pd.DataFrame()

        text = str(response).strip()
        if text == "":
            _PROJECTDATA_CACHE["all_data"][cache_key] = pd.DataFrame()
            return pd.DataFrame()

        df = parse_etap_xml_table(text, element_type)
        _PROJECTDATA_CACHE["all_data"][cache_key] = df.copy()
        return df
    except Exception:
        _PROJECTDATA_CACHE["all_data"][cache_key] = pd.DataFrame()
        return pd.DataFrame()


def xml_local_name(tag):
    text = str(tag)
    if "}" in text:
        text = text.split("}", 1)[1]
    return text


def parse_etap_xml_table(xml_text, element_type=""):
    """
    Parser flexible para XML de ETAP ProjectData.getallelementdata.

    Corrige el caso en que getallelementdata devuelve más información, pero el
    parser anterior solo extraía ID. Soporta:
    - nodos tipo fila con columnas como hijos
    - propiedades repetidas con atributos Name/Value
    - pares <Name> / <Value>
    - atributos directos del elemento
    """
    text = str(xml_text).strip()

    if text == "":
        return pd.DataFrame()

    # Si realmente viene JSON, procesarlo como JSON.
    try:
        if text[0] in ["{", "["]:
            return json_to_dataframe(json.loads(text))
    except Exception:
        pass

    try:
        import xml.etree.ElementTree as ET
        root = ET.fromstring(text)
    except Exception:
        return pd.DataFrame()

    def clean_key(k):
        return str(xml_local_name(k)).strip()

    def add_value(row, key, value):
        key = clean_key(key)
        value = "" if value is None else str(value).strip()

        if key == "" or value == "":
            return

        if normalize_prop_name(key) in ["id", "name", "elementname", "elementid"]:
            key = "ID"

        # Evitar sobreescrituras con campos vacíos; si ya existe, crear sufijo.
        if key in row and str(row[key]).strip() != value:
            i = 2
            new_key = key + "_" + str(i)
            while new_key in row:
                i += 1
                new_key = key + "_" + str(i)
            row[new_key] = value
        else:
            row[key] = value

    def flatten_node(node):
        row = {}

        # Atributos del nodo.
        for key, value in node.attrib.items():
            add_value(row, key, value)

        children = list(node)

        # Nodos propiedad: <Property Name="kV" Value="13.8" />
        prop_name = None
        prop_value = None
        for key, value in node.attrib.items():
            nk = normalize_prop_name(key)
            if nk in ["name", "fieldname", "propertyname", "columnname", "field"]:
                prop_name = value
            if nk in ["value", "val", "data", "text"]:
                prop_value = value
        if prop_name and prop_value:
            add_value(row, prop_name, prop_value)

        # Nodos con hijos directos.
        child_texts = {}
        for child in children:
            tag = clean_key(child.tag)
            tag_norm = normalize_prop_name(tag)
            ctext = (child.text or "").strip()

            # Caso <FieldName>MW</FieldName><Value>7500</Value>
            if ctext:
                child_texts[tag_norm] = ctext

            # Caso hijo simple <MW>7500</MW>
            if ctext and len(list(child)) == 0:
                add_value(row, tag, ctext)

            # Atributos de hijos, especialmente Property Name/Value.
            child_prop_name = None
            child_prop_value = None
            for ak, av in child.attrib.items():
                ank = normalize_prop_name(ak)
                if ank in ["name", "fieldname", "propertyname", "columnname", "field"]:
                    child_prop_name = av
                elif ank in ["value", "val", "data", "text"]:
                    child_prop_value = av
                else:
                    # Guardar atributos útiles con prefijo del tag.
                    add_value(row, tag + "_" + clean_key(ak), av)

            if child_prop_name and child_prop_value:
                add_value(row, child_prop_name, child_prop_value)

            # Caso hijo con subhijos: <Property><Name>MW</Name><Value>7500</Value></Property>
            sub_map = {}
            for sub in list(child):
                sub_tag = normalize_prop_name(clean_key(sub.tag))
                sub_text = (sub.text or "").strip()
                if sub_text:
                    sub_map[sub_tag] = sub_text

                for sak, sav in sub.attrib.items():
                    sank = normalize_prop_name(sak)
                    if sank in ["name", "fieldname", "propertyname", "columnname"]:
                        sub_map["name"] = sav
                    elif sank in ["value", "val", "data", "text"]:
                        sub_map["value"] = sav

            possible_name = (
                sub_map.get("name") or sub_map.get("fieldname") or
                sub_map.get("propertyname") or sub_map.get("columnname")
            )
            possible_value = (
                sub_map.get("value") or sub_map.get("val") or
                sub_map.get("data") or sub_map.get("text")
            )
            if possible_name and possible_value:
                add_value(row, possible_name, possible_value)

            # Si no hay par propiedad/valor, aplanar subhijos con prefijo.
            if not (possible_name and possible_value):
                for sk, sv in sub_map.items():
                    if sk not in ["name", "fieldname", "propertyname", "columnname", "value", "val", "data", "text"]:
                        add_value(row, tag + "_" + sk, sv)

        # Caso en el propio nodo: <Name>MW</Name><Value>7500</Value>
        possible_name = child_texts.get("name") or child_texts.get("fieldname") or child_texts.get("propertyname") or child_texts.get("columnname")
        possible_value = child_texts.get("value") or child_texts.get("val") or child_texts.get("data") or child_texts.get("text")
        if possible_name and possible_value:
            add_value(row, possible_name, possible_value)

        return row

    rows = []

    for node in root.iter():
        tag = clean_key(node.tag)
        tag_norm = normalize_prop_name(tag)

        # Saltar nodos de esquema/diffgram.
        if tag_norm in ["schema", "element", "complextype", "sequence", "diffgram"]:
            continue

        row = flatten_node(node)
        if not row:
            continue

        keys_norm = [normalize_prop_name(k) for k in row.keys()]
        has_id = any(k in keys_norm for k in ["id", "name", "elementname", "elementid"])
        has_many_fields = len([k for k, v in row.items() if not is_invalid_etap_value(v)]) >= 2
        type_match = normalize_prop_name(element_type) in tag_norm if element_type else False

        if has_id or (type_match and has_many_fields):
            if "ID" not in row:
                # Usar nombre del tag solo si parece un elemento, no una tabla genérica.
                if tag_norm not in ["table", "row", "data", "property"]:
                    row["ID"] = tag
            row["Tipo ETAP"] = element_type
            rows.append(row)

    if not rows:
        return pd.DataFrame()

    df = pd.DataFrame(rows)
    df = remove_empty_invalid_columns(df)

    # Evitar filas que solo contienen Tipo ETAP.
    if "Tipo ETAP" in df.columns:
        useful_cols = [c for c in df.columns if c != "Tipo ETAP"]
        if useful_cols:
            df = df[df[useful_cols].apply(lambda r: any(not is_invalid_etap_value(v) for v in r), axis=1)]

    # Deduplicar conservando filas con más datos.
    if "ID" in df.columns:
        df["_data_count"] = df.apply(lambda r: sum(1 for v in r.values if not is_invalid_etap_value(v)), axis=1)
        df = df.sort_values("_data_count", ascending=False)
        df = df.drop_duplicates(subset=["ID"]).drop(columns=["_data_count"]).reset_index(drop=True)

    return df



def get_numeric_value_by_alias(row, aliases, allow_zero=True, min_value=None, max_value=None):
    """
    Compatibilidad: evita NameError si alguna función antigua la llama.
    """
    try:
        for alias in aliases:
            if alias in row:
                value = numeric_value(row.get(alias))
                if value is None:
                    continue
                if not allow_zero and value == 0:
                    continue
                if min_value is not None and value < min_value:
                    continue
                if max_value is not None and value > max_value:
                    continue
                return value
    except Exception:
        pass
    return ""

def choose_first_numeric(row, aliases, allow_zero=False, min_value=None, max_value=None):
    return get_numeric_value_by_alias(row, aliases, allow_zero=allow_zero, min_value=min_value, max_value=max_value)


def blank_if_same_as_voltage(power_value, kv_value):
    """
    Evita una incoherencia observada en el informe: algunas propiedades devueltas
    por ETAP tenían el mismo valor que la tensión nominal y estaban entrando como kVA.
    """
    p = numeric_value(power_value)
    kv = numeric_value(kv_value)
    if p is None:
        return ""
    if kv is not None and abs(p - kv) < 1e-6:
        return ""
    if p < 0:
        return ""
    return round(p, 3)


def standardize_transformer_table(df):
    if df is None or df.empty:
        return pd.DataFrame()

    rows = []
    for _, row in df.iterrows():
        primary_kv = choose_first_numeric(row, [
            "Tensión Primaria", "kV primario", "PrimarykV", "PrimaryKV", "PriKV",
            "RatedPriKV", "RatedPrimaryKV", "HVkV", "HVKV", "kV1", "KV1", "NominalkV1"
        ], allow_zero=False, min_value=0.05, max_value=1000)

        secondary_kv = choose_first_numeric(row, [
            "Tensión Secundaria", "kV secundario", "SecondarykV", "SecondaryKV", "SecKV",
            "RatedSecKV", "RatedSecondaryKV", "LVkV", "LVKV", "kV2", "KV2", "NominalkV2"
        ], allow_zero=False, min_value=0.05, max_value=1000)

        kva = get_power_kva_by_alias(row, kva_aliases=[
            "Potencia kVA", "kVA", "KVA", "RatedkVA", "RatedKVA", "RatingkVA", "PowerkVA",
            "Rating", "Rating1", "Class1kVA"
        ], mva_aliases=[
            "MVA", "RatedMVA", "RatingMVA", "PowerMVA", "BaseMVA",
            "Class1MVA", "Class2MVA", "Class3MVA"
        ], allow_zero=False)

        out = {
            "ID": get_text_value_by_alias(row, ["ID", "Name", "ElementName", "Element ID"]),
            "Potencia kVA": kva,
            "Tensión Primaria": primary_kv,
            "Tensión Secundaria": secondary_kv,
        }

        # Si ETAP no expone potencia/tensiones, conservar barras de conexión como respaldo informativo.
        prim_bus = get_text_value_by_alias(row, ["PrimaryBus", "Primario", "PriBus", "FromBus", "Desde"])
        sec_bus = get_text_value_by_alias(row, ["SecondaryBus", "Secundario", "SecBus", "ToBus", "Hasta"])
        if prim_bus:
            out["Barra primaria"] = prim_bus
        if sec_bus:
            out["Barra secundaria"] = sec_bus

        rows.append(out)

    result = pd.DataFrame(rows)
    return remove_empty_invalid_columns(result)


def standardize_generator_table(df):
    if df is None or df.empty:
        return pd.DataFrame()

    rows = []
    for _, row in df.iterrows():
        kv = choose_first_numeric(row, ["kV", "KV", "RatedkV", "RatedKV", "NominalkV"], allow_zero=False, min_value=0.05, max_value=1000)
        kw = get_power_kw_by_alias(row, allow_zero=False)
        kva = get_power_kva_by_alias(row, allow_zero=False)
        kva = blank_if_same_as_voltage(kva, kv)

        rows.append({
            "ID": get_text_value_by_alias(row, ["ID", "Name", "ElementName", "Element ID"]),
            "Barra": get_text_value_by_alias(row, ["Barra", "Bus", "TerminalBus", "ConnectedBus", "ConnectionBus"]),
            "kW": kw,
            "kVA": kva,
            "kV": kv,
            "FP": choose_first_numeric(row, ["FP", "PF", "PowerFactor"], allow_zero=False, min_value=0, max_value=100),
            "Modo": get_text_value_by_alias(row, ["Modo", "Mode", "ControlMode", "OperationMode"]),
        })

    return remove_empty_invalid_columns(pd.DataFrame(rows))


def standardize_line_cable_table(df):
    if df is None or df.empty:
        return pd.DataFrame()

    rows = []
    for _, row in df.iterrows():
        kv = choose_first_numeric(row, ["kV", "KV", "RatedkV", "RatedKV", "NominalkV"], allow_zero=False, min_value=0.05, max_value=1000)
        amp = choose_first_numeric(row, ["Ampacidad", "Ampacity", "AmpacityA", "Amp", "A", "RatingAmp", "CurrentRating", "ContinuousAmp"], allow_zero=False, min_value=0)
        length = choose_first_numeric(row, ["Longitud", "Length", "Len", "CableLength", "LineLength"], allow_zero=True, min_value=0)

        kva = ""
        if kv != "" and amp != "":
            kva = calculate_kva_from_kv_amp(kv, amp)
            if numeric_value(kva) is not None and numeric_value(kva) <= 0:
                kva = ""

        rows.append({
            "ID": get_text_value_by_alias(row, ["ID", "Name", "ElementName", "Element ID"]),
            "Barra Inicial": get_text_value_by_alias(row, ["Barra Inicial", "Desde", "FromBus", "Bus1", "From", "SourceBus", "SendingBus"]),
            "Barra final": get_text_value_by_alias(row, ["Barra final", "Hasta", "ToBus", "Bus2", "To", "DestinationBus", "ReceivingBus"]),
            "Longitud": length,
            "kV": kv,
            "Ampacidad": amp,
            "kVA": kva,
        })

    return remove_empty_invalid_columns(pd.DataFrame(rows))


def standardize_load_table(df):
    if df is None or df.empty:
        return pd.DataFrame()

    rows = []
    for _, row in df.iterrows():
        kv = choose_first_numeric(row, ["kV", "KV", "RatedkV", "RatedKV", "NominalkV"], allow_zero=False, min_value=0.05, max_value=1000)
        kw = get_power_kw_by_alias(row, allow_zero=True)
        kva = get_power_kva_by_alias(row, allow_zero=False)
        kva = blank_if_same_as_voltage(kva, kv)

        rows.append({
            "ID": get_text_value_by_alias(row, ["ID", "Name", "ElementName", "Element ID"]),
            "Barra": get_text_value_by_alias(row, ["Barra", "Bus", "TerminalBus", "ConnectedBus", "ConnectionBus"]),
            "kW": kw,
            "kVA": kva,
            "kV": kv,
            "FP": choose_first_numeric(row, ["FP", "PF", "PowerFactor"], allow_zero=False, min_value=0, max_value=100),
            "HP": choose_first_numeric(row, ["HP", "HorsePower"], allow_zero=False, min_value=0),
        })

    return remove_empty_invalid_columns(pd.DataFrame(rows))


def standardize_compensation_table(df):
    if df is None or df.empty:
        return pd.DataFrame()

    rows = []
    for _, row in df.iterrows():
        kv = choose_first_numeric(row, ["kV", "KV", "RatedkV", "RatedKV", "NominalkV"], allow_zero=False, min_value=0.05, max_value=1000)
        kvar = choose_first_numeric(row, ["kvar", "KVAR", "Ratedkvar", "RatedKVAR"], allow_zero=False)
        mvar = choose_first_numeric(row, ["Mvar", "MVAR", "RatedMvar", "RatedMVAR"], allow_zero=False)
        kva = get_power_kva_by_alias(row, allow_zero=False)
        kva = blank_if_same_as_voltage(kva, kv)

        rows.append({
            "ID": get_text_value_by_alias(row, ["ID", "Name", "ElementName", "Element ID"]),
            "Barra": get_text_value_by_alias(row, ["Barra", "Bus", "TerminalBus", "ConnectedBus", "ConnectionBus"]),
            "kV": kv,
            "kvar": kvar,
            "Mvar": mvar,
            "kVA": kva,
        })

    return remove_empty_invalid_columns(pd.DataFrame(rows))

def remove_empty_invalid_columns(df):
    """
    Elimina columnas vacías o sin valor técnico real.
    """
    if df is None or df.empty:
        return pd.DataFrame()

    clean = df.copy()

    for col in list(clean.columns):
        clean[col] = clean[col].apply(clean_etap_value)

    drop_cols = []

    for col in clean.columns:
        values = [str(v).strip() for v in clean[col].tolist()]
        useful_values = [v for v in values if not is_invalid_etap_value(v)]
        if len(useful_values) == 0:
            drop_cols.append(col)

    if drop_cols:
        clean = clean.drop(columns=drop_cols)

    return clean


def validate_property_names_for_type(e, element_type, element_names, prop_candidates, max_elements=3):
    """
    Valida propiedades consultándolas en algunos elementos. Solo conserva
    propiedades que devuelvan al menos un valor útil.
    """
    valid_props = []

    if not element_names:
        return valid_props

    sample_names = element_names[:max_elements]

    for prop in prop_candidates:
        has_valid_value = False

        for element_name in sample_names:
            value = safe_get_element_prop(e, element_type, element_name, prop)

            if not is_invalid_etap_value(value):
                has_valid_value = True
                break

        if has_valid_value:
            valid_props.append(prop)

    return valid_props


def normalize_report_columns(df):
    """
    Renombra columnas comunes para que el informe se vea más natural en español.
    """
    if df is None or df.empty:
        return pd.DataFrame()

    rename_map = {
        "ID": "ID",
        "Tipo ETAP": "Tipo ETAP",
        "InService": "En servicio",
        "NominalkV": "kV nominal",
        "NominalKV": "kV nominal",
        "KV": "kV",
        "kV": "kV",
        "Bus": "Barra",
        "TerminalBus": "Barra",
        "FromBus": "Desde",
        "ToBus": "Hasta",
        "PrimaryBus": "Primario",
        "SecondaryBus": "Secundario",
        "PriBus": "Primario",
        "SecBus": "Secundario",
        "PrimarykV": "kV primario",
        "SecondarykV": "kV secundario",
        "kVA": "kVA",
        "MVA": "MVA",
        "MW": "MW",
        "KW": "kW",
        "kW": "kW",
        "Mvar": "Mvar",
        "kvar": "kvar",
        "PF": "FP",
        "PowerFactor": "FP",
        "ControlMode": "Modo de control",
        "Mode": "Modo",
        "Length": "Longitud",
        "Len": "Longitud",
        "Ampacity": "Ampacidad",
        "Amp": "A",
        "HP": "HP",
        "Z": "Z",
        "ZPercent": "Z (%)",
        "PercentZ": "Z (%)",
    }

    cols = {}
    for col in df.columns:
        cols[col] = rename_map.get(col, col)

    clean = df.rename(columns=cols)

    # Evitar columnas duplicadas tras renombrar.
    clean = clean.loc[:, ~clean.columns.duplicated()]

    return clean


def canonicalize_model_columns(df):
    """
    Normaliza columnas equivalentes para que las tablas 4-2 a 4-6 no queden vacías
    cuando ETAP devuelve nombres de propiedades distintos entre versiones.
    """
    if df is None or df.empty:
        return pd.DataFrame()

    clean = df.copy()

    def find_col(possible_names):
        normalized = {str(c).lower().replace(" ", "").replace("_", "").replace("-", ""): c for c in clean.columns}
        for name in possible_names:
            key = str(name).lower().replace(" ", "").replace("_", "").replace("-", "")
            if key in normalized:
                return normalized[key]
        return None

    mappings = {
        "ID": ["ID", "Name", "ElementName", "Nombre"],
        "Barra": ["Barra", "Bus", "TerminalBus", "BusName", "FromBus", "Bus1"],
        "Barra Inicial": ["Barra Inicial", "Desde", "FromBus", "Bus1", "SendingBus", "From"],
        "Barra final": ["Barra final", "Hasta", "ToBus", "Bus2", "ReceivingBus", "To"],
        "Potencia kVA": ["Potencia kVA", "kVA", "RatedkVA", "RatingkVA", "NominalkVA", "PowerkVA", "SizekVA", "MVA", "RatedMVA", "RatingMVA"],
        "Tensión Primaria": ["Tensión Primaria", "kV primario", "PrimarykV", "Pri kV", "PriKV", "HVkV", "RatedkV1"],
        "Tensión Secundaria": ["Tensión Secundaria", "kV secundario", "SecondarykV", "Sec kV", "SecKV", "LVkV", "RatedkV2"],
        "kW": ["kW", "KW", "MW", "RatedkW", "RatingkW"],
        "kVA": ["kVA", "MVA", "RatedkVA", "RatingkVA", "NominalkVA", "PowerkVA"],
        "kV": ["kV", "kV nominal", "NominalkV", "NominalKV", "RatedkV"],
        "FP": ["FP", "PF", "PowerFactor"],
        "Modo": ["Modo", "Mode", "ControlMode"],
        "Longitud": ["Longitud", "Length", "Len", "CableLength"],
        "Ampacidad": ["Ampacidad", "Ampacity", "Amp", "Rating", "CurrentRating"],
        "HP": ["HP", "hp", "RatedHP"]
    }

    for target, possible in mappings.items():
        if target in clean.columns:
            continue

        col = find_col(possible)
        if col is not None:
            clean[target] = clean[col]

    # Conversión de unidades si solo existe MW/MVA y se requiere kW/kVA.
    if "kW" in clean.columns:
        source = find_col(["MW"])
        if source is not None and str(source) != "kW":
            try:
                vals = pd.to_numeric(clean[source], errors="coerce")
                if vals.notna().any():
                    clean["kW"] = vals * 1000.0
            except Exception:
                pass

    if "kVA" in clean.columns:
        source = find_col(["MVA", "RatedMVA", "RatingMVA"])
        if source is not None and str(source) != "kVA":
            try:
                vals = pd.to_numeric(clean[source], errors="coerce")
                if vals.notna().any():
                    clean["kVA"] = vals * 1000.0
            except Exception:
                pass

    if "Potencia kVA" in clean.columns:
        source = find_col(["MVA", "RatedMVA", "RatingMVA"])
        if source is not None and str(source) != "Potencia kVA":
            try:
                vals = pd.to_numeric(clean[source], errors="coerce")
                if vals.notna().any():
                    clean["Potencia kVA"] = vals * 1000.0
            except Exception:
                pass

    return clean


def critical_properties_for_category(category):
    """
    Propiedades mínimas por categoría para evitar consultas masivas lentas y, a la vez,
    evitar tablas del informe con solo ID. Incluye aliases frecuentes de ETAP API.
    """
    if category == "transformadores":
        return [
            "PrimaryBus", "SecondaryBus", "FromBus", "ToBus", "PriBus", "SecBus",
            "PrimarykV", "SecondarykV", "PriKV", "SecKV", "KV1", "KV2", "kV1", "kV2",
            "RatedPriKV", "RatedSecKV", "RatedPrimaryKV", "RatedSecondaryKV",
            "kVA", "KVA", "MVA", "RatedkVA", "RatedKVA", "RatingkVA", "RatedMVA", "RatingMVA",
            "Class1MVA", "Class2MVA", "Class3MVA", "Rating", "Rating1", "Rating2",
            "Z", "ZPercent", "PercentZ", "X/R", "XR"
        ]

    if category == "generadores":
        return [
            "Bus", "TerminalBus", "ConnectedBus",
            "kW", "KW", "MW", "Rating", "RatedkW", "RatedMW",
            "kVA", "KVA", "MVA", "RatedkVA", "RatedMVA",
            "kV", "KV", "RatedkV", "RatedKV", "NominalkV",
            "PF", "PowerFactor", "PowerFactorPercent",
            "Mode", "ControlMode", "OperationMode"
        ]

    if category == "cargas":
        return [
            "Bus", "TerminalBus", "ConnectedBus",
            "kW", "KW", "MW", "Rating", "RatedkW",
            "kVA", "KVA", "MVA", "RatedkVA", "RatedMVA",
            "kvar", "KVAR", "Mvar", "MVAR",
            "kV", "KV", "RatedkV", "RatedKV", "NominalkV",
            "PF", "PowerFactor", "HP", "HorsePower"
        ]

    if category == "lineas_cables":
        return [
            "FromBus", "ToBus", "Bus1", "Bus2", "SourceBus", "DestinationBus",
            "Length", "Len", "CableLength", "LineLength",
            "kV", "KV", "RatedkV", "RatedKV", "NominalkV",
            "Ampacity", "AmpacityA", "Amp", "RatingAmp", "CurrentRating", "ContinuousAmp",
            "kVA", "KVA", "MVA"
        ]

    if category == "compensacion":
        return [
            "Bus", "TerminalBus", "ConnectedBus",
            "kV", "KV", "RatedkV", "RatedKV", "NominalkV",
            "kvar", "KVAR", "Mvar", "MVAR", "kVA", "KVA", "MVA",
            "Bank", "Mode", "ControlMode"
        ]

    if category == "barras":
        return ["NominalkV", "NominalKV", "kV", "KV", "RatedkV", "Area", "Zone"]

    return []


def enrich_table_with_critical_props(e, df, category, max_elements=None):
    """
    Enriquece selectivamente una tabla usando:
    - propiedades críticas por categoría
    - propiedades reales detectadas por getelementpropertynamesxml
    - validación por valor útil

    Se ejecuta aun si existe getallelementdata, porque algunos XML solo entregan ID.
    """
    if df is None or df.empty:
        return pd.DataFrame()

    if max_elements is None:
        max_elements = MAX_ENRICH_ELEMENTS_PER_CATEGORY

    if not ENRICH_MODEL_CRITICAL_PROPERTIES:
        return canonicalize_model_columns(df)

    if "ID" not in df.columns:
        return canonicalize_model_columns(df)

    # Si no existe Tipo ETAP, intentar inferirlo de la tabla completa.
    if "Tipo ETAP" not in df.columns:
        return canonicalize_model_columns(df)

    rows = []
    df_limited = df.head(max_elements).copy()

    for element_type in unique_keep_order(df_limited["Tipo ETAP"].dropna().astype(str).tolist()):
        subset = df_limited[df_limited["Tipo ETAP"].astype(str) == element_type].copy()
        names = subset["ID"].dropna().astype(str).tolist()

        if not names:
            continue

        available_props = safe_get_property_names_xml(e, element_type)
        prop_candidates = unique_keep_order(
            critical_properties_for_category(category) +
            select_relevant_property_names(element_type, available_props, category) +
            available_props
        )

        # Limitar candidatos para no volver al modo excesivamente lento.
        prop_candidates = prop_candidates[:60]

        # Validar solo con algunos elementos.
        valid_props = validate_property_names_for_type(
            e,
            element_type,
            names,
            prop_candidates,
            max_elements=3
        )

        # Si no se validó nada, al menos intentar las propiedades críticas sin validación previa.
        if not valid_props:
            valid_props = critical_properties_for_category(category)[:25]

        for _, row in subset.iterrows():
            element_id = str(row.get("ID", "")).strip()
            out = row.to_dict()

            if element_id:
                for prop in valid_props:
                    if prop in out and not is_invalid_etap_value(out.get(prop)):
                        continue

                    value = safe_get_element_prop(e, element_type, element_id, prop)
                    if not is_invalid_etap_value(value):
                        out[prop] = value

            rows.append(out)

    # Conservar filas restantes si excede el límite.
    if len(df) > max_elements:
        for _, row in df.iloc[max_elements:].iterrows():
            rows.append(row.to_dict())

    if not rows:
        return canonicalize_model_columns(df)

    result = pd.DataFrame(rows)
    result = remove_empty_invalid_columns(result)
    result = normalize_report_columns(result)
    result = canonicalize_model_columns(result)

    return result


def select_relevant_property_names(element_type, available_props, category):
    """
    Selecciona propiedades relevantes combinando:
    - Propiedades reportadas por getelementpropertynamesxml.
    - Candidatos típicos de ETAP API.
    La validación posterior evita que entren campos inválidos.
    """
    base_candidates = [
        "ID", "Name", "ElementName", "InService", "NominalkV", "NominalKV", "kV", "KV", "RatedkV", "RatedKV",
        "Bus", "TerminalBus", "ConnectedBus", "ConnectionBus",
        "FromBus", "ToBus", "Bus1", "Bus2", "From", "To",
        "PrimaryBus", "SecondaryBus", "PriBus", "SecBus",
        "PrimarykV", "PrimaryKV", "SecondarykV", "SecondaryKV", "PriKV", "SecKV", "RatedPriKV", "RatedSecKV",
        "HVkV", "LVkV", "kV1", "kV2", "KV1", "KV2", "NominalkV1", "NominalkV2",
        "kVA", "KVA", "RatedkVA", "RatedKVA", "RatingkVA", "PowerkVA",
        "MVA", "RatedMVA", "RatingMVA", "PowerMVA", "BaseMVA",
        "MW", "kW", "KW", "RatedMW", "RatedkW", "PowerMW", "PowerkW",
        "Mvar", "MVAR", "kvar", "KVAR", "HP",
        "PF", "PowerFactor",
        "Length", "Len", "CableLength", "LineLength",
        "Ampacity", "AmpacityA", "Amp", "RatingAmp", "CurrentRating", "ContinuousAmp",
        "Mode", "ControlMode", "OperationMode",
        "Z", "ZPercent", "PercentZ", "AnsiPosZ", "AnsiZeroZ", "X/R", "XR"
    ]

    category_candidates = {
        "barras": ["ID", "Name", "NominalkV", "NominalKV", "kV", "RatedkV", "InService", "Type", "Area", "Zone"],
        "transformadores": ["ID", "Name", "PrimarykV", "SecondarykV", "PriKV", "SecKV", "RatedPriKV", "RatedSecKV", "kV1", "kV2", "kVA", "RatedkVA", "MVA", "RatedMVA", "PrimaryBus", "SecondaryBus", "PriBus", "SecBus", "FromBus", "ToBus"],
        "generadores": ["ID", "Name", "Bus", "TerminalBus", "kW", "MW", "kVA", "MVA", "RatedkVA", "RatedMVA", "kV", "PF", "PowerFactor", "Mode", "ControlMode"],
        "cargas": ["ID", "Name", "Bus", "TerminalBus", "kW", "MW", "kVA", "MVA", "kvar", "Mvar", "kV", "PF", "PowerFactor", "HP"],
        "lineas_cables": ["ID", "Name", "FromBus", "ToBus", "Bus1", "Bus2", "Length", "CableLength", "kV", "Ampacity", "AmpacityA", "Amp", "RatingAmp"],
        "compensacion": ["ID", "Name", "Bus", "TerminalBus", "kV", "kvar", "Mvar", "kVA", "MVA", "Bank", "Mode"],
    }

    return filter_options(category_candidates.get(category, []) + base_candidates + list(available_props or []))[:80]


def build_element_table_from_projectdata(e, category, element_types):
    """
    Construye tabla de elementos del modelo de forma optimizada.

    Modo rápido por defecto:
    1. Usa getallelementdata(ElementType), que entrega la información en bloque.
    2. Si no hay tabla, usa getelementnames(ElementType) solo para listar elementos.
    3. No consulta getelementprop() por cada propiedad/elemento, porque esa era la
       principal causa de demora.

    Para diagnóstico profundo puede cambiar FAST_MODEL_EXTRACTION = True.
    """
    source_parts = []
    frames = []

    available_types = safe_get_element_types(e)
    available_norm = {normalize_prop_name(t): t for t in available_types}

    types_to_try = []
    for t in element_types:
        nt = normalize_prop_name(t)
        if available_norm:
            if nt in available_norm:
                types_to_try.append(available_norm[nt])
        else:
            types_to_try.append(t)

    if not types_to_try:
        types_to_try = element_types

    for element_type in unique_keep_order(types_to_try):
        # 1) Lectura en bloque: mucho más rápida que getelementprop elemento por elemento.
        df_xml = safe_get_all_element_data(e, element_type)
        if df_xml is not None and not df_xml.empty:
            frames.append(df_xml)
            source_parts.append(element_type + ":getallelementdata")
            continue

        # 2) Respaldo rápido: solo nombres.
        names = safe_get_element_names(e, element_type)
        if names:
            frames.append(pd.DataFrame({
                "ID": names,
                "Tipo ETAP": [element_type] * len(names)
            }))
            source_parts.append(element_type + ":getelementnames")

        # 3) Respaldo profundo opcional. Desactivado por defecto por desempeño.
        if (not FAST_MODEL_EXTRACTION) and names:
            available_props = safe_get_property_names_xml(e, element_type)
            prop_candidates = select_relevant_property_names(element_type, available_props, category)
            valid_props = validate_property_names_for_type(e, element_type, names, prop_candidates[:20], max_elements=2)

            rows = []
            for element_name in names[:80]:
                row = {"ID": element_name, "Tipo ETAP": element_type}
                for prop in valid_props[:12]:
                    value = safe_get_element_prop(e, element_type, element_name, prop)
                    if not is_invalid_etap_value(value):
                        row[prop] = value
                rows.append(row)

            if rows:
                frames.append(pd.DataFrame(rows))
                source_parts.append(element_type + ":getelementprop")

    if not frames:
        return pd.DataFrame(), "No disponible"

    df = pd.concat(frames, ignore_index=True, sort=False)
    df = remove_empty_invalid_columns(df)
    df = normalize_report_columns(df)

    id_col = None
    for candidate in ["ID", "Name", "Nombre", "ElementName"]:
        if candidate in df.columns:
            id_col = candidate
            break

    if id_col:
        df["_data_count"] = df.apply(lambda r: sum(1 for v in r.values if not is_invalid_etap_value(v)), axis=1)
        df = df.sort_values("_data_count", ascending=False)
        df = df.drop_duplicates(subset=[id_col]).drop(columns=["_data_count"]).reset_index(drop=True)

    source = "ProjectData optimizado: " + ", ".join(unique_keep_order(source_parts))
    return df, source


def first_non_empty_dataframe(e, candidates):
    for method_name, args in candidates:
        df = safe_projectdata_call(e, method_name, args)
        if df is not None and not df.empty:
            return df, method_name
    return pd.DataFrame(), ""


def compact_dataframe(df, preferred_columns=None, max_rows=30):
    if df is None or df.empty:
        return pd.DataFrame()

    clean = remove_empty_invalid_columns(df)
    clean = normalize_report_columns(clean)

    if clean.empty:
        return pd.DataFrame()

    clean = clean.dropna(axis=1, how="all")

    if preferred_columns:
        cols = []
        normalized = {str(c).lower().replace(" ", "").replace("_", ""): c for c in clean.columns}

        for p in preferred_columns:
            key = p.lower().replace(" ", "").replace("_", "")
            if key in normalized:
                cols.append(normalized[key])

        if cols:
            extra_cols = [c for c in clean.columns if c not in cols]
            # Máximo 2 columnas adicionales para no dañar la legibilidad.
            clean = clean[cols + extra_cols[:2]]

    # Eliminar columnas duplicadas.
    clean = clean.loc[:, ~clean.columns.duplicated()]

    if len(clean) > max_rows:
        clean = clean.head(max_rows)

    return clean


def infer_model_elements_from_lfr(bus_all, equipment_all):
    buses = pd.DataFrame()
    branches = pd.DataFrame()
    if bus_all is not None and not bus_all.empty:
        buses = bus_all[["Barra", "kV nominal"]].drop_duplicates().copy()
        buses = buses.rename(columns={"Barra": "ID", "kV nominal": "Un [kV]"})
    if equipment_all is not None and not equipment_all.empty:
        cols = [c for c in ["Desde", "Hasta", "TYPE", "kV", "MW", "Mvar", "MVA", "Amp", "Cargabilidad (%)"] if c in equipment_all.columns]
        if cols:
            branches = equipment_all[cols].drop_duplicates().copy()
            branches = branches.rename(columns={"TYPE": "Tipo ETAP"})
    return buses, branches


def extract_branch_impedance_table(e, branches_df, max_rows=25):
    """
    Usa e.projectdata.getbranchimpedance solo cuando existan nombres reales de
    ramas/líneas/cables y buses From/To. No usa barras como nombre de rama.
    """
    if branches_df is None or branches_df.empty:
        return pd.DataFrame()

    pd_obj = get_projectdata_object(e)
    if pd_obj is None or not hasattr(pd_obj, "getbranchimpedance"):
        return pd.DataFrame()

    rows = []
    checked = 0

    # Identificar columnas posibles.
    id_cols = [c for c in ["ID", "Name", "ElementName"] if c in branches_df.columns]
    from_cols = [c for c in ["FromBus", "Desde", "Bus1", "From"] if c in branches_df.columns]
    to_cols = [c for c in ["ToBus", "Hasta", "Bus2", "To"] if c in branches_df.columns]

    if not id_cols or not from_cols or not to_cols:
        return pd.DataFrame()

    id_col = id_cols[0]
    from_col = from_cols[0]
    to_col = to_cols[0]

    for _, row in branches_df.iterrows():
        if checked >= max_rows:
            break

        branch_name = str(row.get(id_col, "")).strip()
        ref_bus = str(row.get(from_col, "")).strip()
        remote_bus = str(row.get(to_col, "")).strip()

        if branch_name == "" or ref_bus == "" or remote_bus == "":
            continue

        if branch_name.lower() == "nan" or ref_bus.lower() == "nan" or remote_bus.lower() == "nan":
            continue

        try:
            response = pd_obj.getbranchimpedance(branch_name, ref_bus, remote_bus, "False")
            parsed = json.loads(response) if isinstance(response, str) else response

            if isinstance(parsed, dict):
                # Si todos los valores numéricos vienen en cero, no se descarta,
                # pero queda identificado con nombre real de rama, no con barras.
                parsed["ID / Rama"] = branch_name
                parsed["Desde"] = ref_bus
                parsed["Hasta"] = remote_bus
                rows.append(parsed)
                checked += 1
        except Exception:
            continue

    if not rows:
        return pd.DataFrame()

    df = pd.DataFrame(rows)
    key_cols = ["ID / Rama", "Desde", "Hasta"]
    cols = key_cols + [c for c in df.columns if c not in key_cols]
    return df[cols]



def summarize_elements_by_id(df):
    """
    Devuelve una tabla resumida solo con ID.

    Esta función se usa en la versión rápida del informe para evitar consultas
    lentas de propiedades elemento por elemento. Acepta tablas provenientes de
    getelementnames(), getbusnames(), getallelementdata() o datos inferidos desde LFR.
    """
    if df is None or df.empty:
        return pd.DataFrame(columns=["ID"])

    clean = df.copy()

    id_col = None
    for candidate in ["ID", "Name", "ElementName", "Element ID", "Nombre", "Valor", "Bus", "Barra", "Desde", "FromBus"]:
        if candidate in clean.columns:
            id_col = candidate
            break

    if id_col is None:
        id_col = clean.columns[0]

    result = clean[[id_col]].copy()
    result.columns = ["ID"]

    result = result.dropna()
    result["ID"] = result["ID"].astype(str).str.strip()

    result = result[result["ID"] != ""]
    result = result[result["ID"].str.lower() != "nan"]
    result = result[result["ID"].str.lower() != "none"]
    result = result[result["ID"].str.lower() != "null"]

    result = result.drop_duplicates().reset_index(drop=True)

    return result

def summarize_elements_by_id_and_type(df, element_type_label):
    """
    Devuelve inventario resumido ID + Tipo de elemento.
    """
    base = summarize_elements_by_id(df)
    if base.empty:
        return pd.DataFrame(columns=["ID", "Tipo de elemento"])

    base["Tipo de elemento"] = element_type_label
    return base[["ID", "Tipo de elemento"]]


def get_fast_all_element_data_for_types(e, element_types):
    """
    Intenta usar getallelementdata() para obtener datos en bloque, sin consultas por elemento.
    Si no devuelve datos útiles, retorna DataFrame vacío.
    """
    frames = []

    for element_type in element_types:
        try:
            df = safe_get_all_element_data(e, element_type)
            if df is not None and not df.empty:
                df = remove_empty_invalid_columns(df)
                df = normalize_report_columns(df)
                if not df.empty:
                    frames.append(df)
        except Exception:
            pass

    if not frames:
        return pd.DataFrame()

    out = pd.concat(frames, ignore_index=True, sort=False)
    out = out.loc[:, ~out.columns.duplicated()]
    return out


def get_ids_by_element_types(e, element_types, element_type_label):
    """
    Obtiene IDs con getelementnames() en modo rápido.
    """
    rows = []

    available_types = safe_get_element_types(e)
    available_norm = {normalize_prop_name(t): t for t in available_types}

    types_to_try = []
    for t in element_types:
        nt = normalize_prop_name(t)
        if available_norm:
            if nt in available_norm:
                types_to_try.append(available_norm[nt])
        else:
            types_to_try.append(t)

    if not types_to_try:
        types_to_try = element_types

    used_types = []

    for element_type in unique_keep_order(types_to_try):
        names = safe_get_element_names(e, element_type)
        if names:
            used_types.append(element_type)
            for name in names:
                rows.append({
                    "ID": name,
                    "Tipo de elemento": element_type_label
                })

    if not rows:
        return pd.DataFrame(columns=["ID", "Tipo de elemento"]), "No disponible"

    df = pd.DataFrame(rows)
    df["ID"] = df["ID"].astype(str).str.strip()
    df = df[df["ID"] != ""]
    df = df.drop_duplicates(subset=["ID"]).reset_index(drop=True)

    source = "ProjectData.getelementnames: " + ", ".join(used_types)
    return df, source


def build_fast_bus_inventory(e, bus_all):
    """
    Barras: ID + kV nominal usando principalmente LFR.
    """
    # Fuente principal: LFR, porque ya tiene IDFrom y kV.
    if bus_all is not None and not bus_all.empty:
        df = bus_all[["Barra", "kV nominal"]].drop_duplicates().copy()
        df = df.rename(columns={"Barra": "ID"})
        df["ID"] = df["ID"].astype(str).str.strip()
        df = df[df["ID"] != ""]
        df = df.drop_duplicates(subset=["ID"]).reset_index(drop=True)
        if not df.empty:
            return df, "LFR: IDFrom/kV"

    # Respaldo: getbusnames().
    try:
        pd_obj = get_projectdata_object(e)
        if pd_obj is not None and hasattr(pd_obj, "getbusnames"):
            response = pd_obj.getbusnames()
            names = filter_options(flatten_strings(response))
            if names:
                return pd.DataFrame({"ID": names}), "ProjectData.getbusnames"
    except Exception:
        pass

    # Último respaldo: getelementnames(BUS).
    df, src = get_ids_by_element_types(e, ["BUS"], "Barra")
    return df[["ID"]] if not df.empty else df, src


def build_fast_branch_inventory(e, equipment_all):
    """
    Líneas/cables/ramas: ID + Barra Inicial + Barra final + kV desde LFR.
    No intenta extraer propiedades lentas.
    """
    rows = []

    if equipment_all is not None and not equipment_all.empty:
        temp = equipment_all.copy()

        # Usar solo filas con desde/hasta válidos.
        for _, row in temp.iterrows():
            desde = str(row.get("Desde", "")).strip()
            hasta = str(row.get("Hasta", "")).strip()

            if desde == "" or hasta == "" or desde.lower() == "nan" or hasta.lower() == "nan":
                continue

            if desde == hasta:
                continue

            kv = row.get("kV", "")
            try:
                kv_value = float(kv)
                if kv_value == 0:
                    kv = ""
            except Exception:
                pass

            rows.append({
                "ID": "Rama " + desde + " - " + hasta,
                "Barra Inicial": desde,
                "Barra final": hasta,
                "kV": kv
            })

    if rows:
        df = pd.DataFrame(rows)
        df = df.drop_duplicates(subset=["Barra Inicial", "Barra final", "kV"]).reset_index(drop=True)
        return df, "LFR: IDFrom/IDTo/kV"

    # Respaldo con IDs de CABLE/LINE si no se puede inferir conexión.
    df, src = get_ids_by_element_types(e, ["CABLE", "LINE", "TLINE", "BRANCH"], "Línea/Cable")
    return df, src


def merge_bulk_data_with_ids(ids_df, bulk_df, preferred_cols, element_type_label):
    """
    Usa getallelementdata() solo si entregó columnas útiles en bloque.
    Si no, conserva ID + Tipo de elemento.
    """
    if ids_df is None or ids_df.empty:
        ids_df = pd.DataFrame(columns=["ID", "Tipo de elemento"])

    if bulk_df is None or bulk_df.empty:
        return ids_df

    bulk = remove_empty_invalid_columns(bulk_df)
    bulk = normalize_report_columns(bulk)
    bulk = canonicalize_model_columns(bulk)

    if bulk.empty:
        return ids_df

    # Encontrar columna ID en datos bulk.
    id_col = None
    for candidate in ["ID", "Name", "ElementName", "Nombre"]:
        if candidate in bulk.columns:
            id_col = candidate
            break

    if id_col is None:
        return ids_df

    bulk = bulk.rename(columns={id_col: "ID"})
    bulk["ID"] = bulk["ID"].astype(str).str.strip()
    bulk = bulk[bulk["ID"] != ""]
    bulk = bulk.drop_duplicates(subset=["ID"]).reset_index(drop=True)

    # Si el bulk solo tiene ID o Tipo, no aporta.
    useful_cols = [c for c in bulk.columns if c not in ["ID", "Tipo ETAP", "Tipo de elemento"]]
    if len(useful_cols) == 0:
        return ids_df

    merged = pd.merge(ids_df, bulk, on="ID", how="left")
    merged["Tipo de elemento"] = element_type_label

    # Mantener columnas de interés si existen; si no, ID + tipo.
    cols = ["ID", "Tipo de elemento"]
    for col in preferred_cols:
        if col in merged.columns and col not in cols:
            # Solo conservar si tiene al menos un valor real.
            vals = merged[col].dropna().astype(str).str.strip()
            vals = vals[(vals != "") & (vals.str.lower() != "nan") & (vals.str.lower() != "none")]
            if len(vals) > 0:
                cols.append(col)

    # Si no hubo columnas preferidas, conservar hasta 2 útiles del bulk.
    if len(cols) == 2:
        for col in useful_cols[:2]:
            if col in merged.columns and col not in cols:
                cols.append(col)

    return merged[cols].drop_duplicates().reset_index(drop=True)


def build_fast_element_inventory(e, element_types, element_type_label, preferred_cols):
    """
    Inventario rápido para transformadores, generadores, cargas y compensación:
    - IDs con getelementnames()
    - Datos adicionales solo si getallelementdata() los entrega en bloque
    - Sin getelementprop() por elemento
    """
    ids_df, src_names = get_ids_by_element_types(e, element_types, element_type_label)
    bulk_df = get_fast_all_element_data_for_types(e, element_types)

    result = merge_bulk_data_with_ids(ids_df, bulk_df, preferred_cols, element_type_label)

    src = src_names
    if bulk_df is not None and not bulk_df.empty:
        src += " + getallelementdata"

    return result, src



def scale_power_columns_to_report_units(df):
    """
    Corrige escala de columnas kW/kVA cuando ETAP entrega valores en W/VA.
    Regla práctica: si los valores son muy altos para kW/kVA, se dividen entre 1000.
    Esto corrige casos como 7,500,000 -> 7,500 kW y 8,823,530 -> 8,823.53 kVA.
    """
    if df is None or df.empty:
        return pd.DataFrame()

    result = df.copy()

    for col in ["kW", "kVA"]:
        if col not in result.columns:
            continue

        values = pd.to_numeric(result[col], errors="coerce")
        valid = values.dropna()

        if valid.empty:
            continue

        # Si la mayoría de valores supera 10,000, probablemente está en W/VA.
        high_ratio = (valid.abs() > 10000).sum() / len(valid)

        if high_ratio >= 0.5:
            result[col] = values / 1000.0

    return result


def model_summary_dataframe_fast(model):
    """
    Resumen simple del modelo sin columna fuente.
    Se omiten líneas/cables y compensación reactiva porque no aportan información
    relevante en el informe generado con el modo rápido.
    """
    labels = {
        "barras": "Barras",
        "transformadores": "Transformadores",
        "generadores": "Generadores",
        "cargas": "Cargas"
    }

    rows = []
    for key, label in labels.items():
        df = model.get(key, pd.DataFrame())
        rows.append({
            "Tipo de elemento": label,
            "Cantidad identificada": 0 if df is None or df.empty else len(df)
        })

    result = pd.DataFrame(rows)
    if "Cantidad identificada" in result.columns:
        result["Cantidad identificada"] = pd.to_numeric(result["Cantidad identificada"], errors="coerce").fillna(0).astype(int)
    return result



def extract_model_information(e, bus_all, equipment_all):
    """
    Inventario rápido enriquecido sin consultas lentas elemento por elemento.

    - Barras: ID + kV nominal desde LFR.
    - Líneas/cables/ramas: ID inferido + barra inicial + barra final + kV desde LFR.
    - Transformadores, generadores, cargas y compensación: ID + tipo de elemento.
      Si getallelementdata() entrega datos útiles en bloque, se anexan automáticamente.
    """
    model = {}
    sources = {}

    print("Construyendo inventario rápido de barras desde LFR...")
    model["barras"], sources["barras"] = build_fast_bus_inventory(e, bus_all)

    print("Construyendo inventario rápido de transformadores...")
    model["transformadores"], sources["transformadores"] = build_fast_element_inventory(
        e,
        ["XFORM2W", "XFORM3W", "T2", "T3", "XFMR", "TRANSFORMER"],
        "Transformador",
        ["Potencia kVA", "Tensión Primaria", "Tensión Secundaria", "kVA", "MVA", "kV primario", "kV secundario"]
    )

    print("Construyendo inventario rápido de generadores...")
    model["generadores"], sources["generadores"] = build_fast_element_inventory(
        e,
        ["GEN", "SYNCGEN", "SYNGEN", "GENERATOR"],
        "Generador",
        ["kW", "kVA", "kV", "FP", "Modo", "MW", "MVA"]
    )

    print("Construyendo inventario rápido de líneas/cables desde LFR...")
    model["lineas_cables"], sources["lineas_cables"] = build_fast_branch_inventory(e, equipment_all)

    print("Construyendo inventario rápido de cargas...")
    model["cargas"], sources["cargas"] = build_fast_element_inventory(
        e,
        ["LUMPEDLOAD", "LOAD", "STATICLOAD", "MOTOR", "INDMOTOR"],
        "Carga",
        ["kW", "kVA", "kV", "FP", "HP", "MW", "MVA"]
    )

    print("Construyendo inventario rápido de compensación reactiva...")
    model["compensacion"], sources["compensacion"] = build_fast_element_inventory(
        e,
        ["CAPACITOR", "CAP", "REACTOR", "SHUNT"],
        "Compensación reactiva",
        ["kV", "kvar", "Mvar", "kVA", "MVA"]
    )

    # Normalización final ligera.
    model["barras"] = compact_dataframe(model["barras"], ["ID", "kV nominal", "kV"], MAX_MODEL_TABLE_ROWS)
    model["transformadores"] = compact_dataframe(model["transformadores"], ["ID"], MAX_MODEL_TABLE_ROWS)
    model["generadores"] = compact_dataframe(model["generadores"], ["ID", "Tipo de elemento", "kW", "kVA", "kV", "FP", "Modo"], MAX_MODEL_TABLE_ROWS)
    model["generadores"] = scale_power_columns_to_report_units(model["generadores"])
    model["lineas_cables"] = compact_dataframe(model["lineas_cables"], ["ID", "Barra Inicial", "Barra final", "kV", "Tipo de elemento"], MAX_MODEL_TABLE_ROWS)
    model["cargas"] = compact_dataframe(model["cargas"], ["ID", "Tipo de elemento", "kW", "kVA", "kV", "FP", "HP"], MAX_MODEL_TABLE_ROWS)
    model["cargas"] = scale_power_columns_to_report_units(model["cargas"])
    model["compensacion"] = compact_dataframe(model["compensacion"], ["ID", "Tipo de elemento", "kV", "kvar", "Mvar", "kVA"], MAX_MODEL_TABLE_ROWS)

    return model, sources


def model_summary_dataframe(model, sources):
    rows = []
    labels = {
        "barras": "Barras",
        "transformadores": "Transformadores",
        "generadores": "Generadores",
        "cargas": "Cargas",
        "lineas_cables": "Líneas y cables",
        "compensacion": "Compensación reactiva"
    }
    for key, label in labels.items():
        df = model.get(key, pd.DataFrame())
        rows.append({
            "Tipo de elemento": label,
            "Cantidad identificada": 0 if df is None or df.empty else len(df)
        })
    return pd.DataFrame(rows)

def definitions_dataframe():
    return pd.DataFrame([
        {"Término": "Un", "Definición": "Tensión nominal del equipo o barra del sistema eléctrico."},
        {"Término": "Flujo de potencia / flujo de carga", "Definición": "Estudio en estado estable que calcula tensiones, ángulos, corrientes, potencia activa, potencia reactiva y cargabilidad en el sistema."},
        {"Término": "Barra", "Definición": "Nodo eléctrico del modelo donde se conectan elementos como transformadores, líneas, generadores o cargas."},
        {"Término": "Ampacidad", "Definición": "Corriente máxima que un conductor puede transportar continuamente sin exceder su temperatura admisible."},
        {"Término": "PF", "Definición": "Factor de potencia, relación entre potencia activa y potencia aparente."},
        {"Término": "FLA", "Definición": "Corriente a plena carga del equipo."},
        {"Término": "Z [%]", "Definición": "Impedancia porcentual de un transformador u otro equipo, referida a una base eléctrica."},
        {"Término": "RPM", "Definición": "Revoluciones por minuto de una máquina rotativa."},
        {"Término": "EFF [%]", "Definición": "Eficiencia del equipo expresada en porcentaje."},
        {"Término": "Swing / Slack", "Definición": "Barra o fuente de referencia que absorbe el desbalance de potencia del sistema y fija referencia de tensión y ángulo."},
        {"Término": "Voltage Control (PV)", "Definición": "Modo de control en el cual el generador mantiene tensión objetivo ajustando potencia reactiva dentro de sus límites."},
        {"Término": "Mvar Control", "Definición": "Modo de operación con potencia activa y reactiva fijadas, sin regulación automática de tensión."},
        {"Término": "PF Control", "Definición": "Modo de operación en el cual el generador ajusta su excitación para mantener un factor de potencia especificado."},
        {"Término": "Cargabilidad", "Definición": "Relación porcentual entre la carga operativa de un equipo y su capacidad nominal."},
        {"Término": "Regulación de tensión", "Definición": "Variación de la tensión respecto al valor nominal o referencia adoptada para el análisis."},
    ])


def normative_references_dataframe():
    return pd.DataFrame([
        {"Norma / guía": "IEEE Std 3002.2-2018", "Aplicación en el informe": "Referencia principal para estudios de flujo de carga en sistemas industriales y comerciales; soporta alcance, datos requeridos, escenarios, análisis de tensión, cargabilidad y validación de resultados."},
        {"Norma / guía": "IEEE Std 399", "Aplicación en el informe": "Referencia histórica de análisis de sistemas eléctricos industriales y comerciales, incluyendo estudios de flujo de carga."}
    ])

# =============================================================================
# DOCX FORMATO
# =============================================================================

def set_cell_shading(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_cell_text(cell, text, bold=False, italic=False):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(str(text))
    run.bold = bold
    run.italic = italic
    run.font.name = DOC_FONT_NAME
    run.font.size = Pt(DOC_TABLE_FONT_SIZE)
    return p


def set_table_borders(table):
    tbl = table._tbl
    tbl_pr = tbl.tblPr
    borders = tbl_pr.first_child_found_in("w:tblBorders")
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)

    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tag = "w:" + edge
        element = borders.find(qn(tag))
        if element is None:
            element = OxmlElement(tag)
            borders.append(element)
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), "4")
        element.set(qn("w:space"), "0")
        element.set(qn("w:color"), "BFBFBF")


def keep_table_on_page(table):
    for row in table.rows:
        tr_pr = row._tr.get_or_add_trPr()
        cant_split = OxmlElement("w:cantSplit")
        tr_pr.append(cant_split)


def set_repeat_table_header(row):
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tbl_header.set(qn("w:val"), "true")
    tr_pr.append(tbl_header)


def clear_document_body_keep_sections(doc):
    body = doc._body._element

    for child in list(body):
        if child.tag == qn("w:sectPr"):
            continue
        body.remove(child)


def set_default_styles(doc):
    """
    Aplica propiedades de fuente y estilos base similares al informe de referencia:
    Arial, cuerpo técnico en tamaño 10 pt, encabezados jerárquicos y tablas compactas.
    """
    try:
        styles = doc.styles

        for style_name in ["Normal", "Body Text"]:
            if style_name in styles:
                style = styles[style_name]
                style.font.name = DOC_FONT_NAME
                style.font.size = Pt(DOC_FONT_SIZE)
                style._element.rPr.rFonts.set(qn("w:eastAsia"), DOC_FONT_NAME)

        if "Heading 1" in styles:
            style = styles["Heading 1"]
            style.font.name = DOC_FONT_NAME
            style.font.size = Pt(DOC_HEADING1_SIZE)
            style.font.bold = True
            style.font.color.rgb = RGBColor(0, 0, 0)
            style._element.rPr.rFonts.set(qn("w:eastAsia"), DOC_FONT_NAME)

        if "Heading 2" in styles:
            style = styles["Heading 2"]
            style.font.name = DOC_FONT_NAME
            style.font.size = Pt(DOC_HEADING2_SIZE)
            style.font.bold = True
            style.font.color.rgb = RGBColor(0, 0, 0)
            style._element.rPr.rFonts.set(qn("w:eastAsia"), DOC_FONT_NAME)

        if "Heading 3" in styles:
            style = styles["Heading 3"]
            style.font.name = DOC_FONT_NAME
            style.font.size = Pt(DOC_HEADING2_SIZE)
            style.font.bold = True
            style._element.rPr.rFonts.set(qn("w:eastAsia"), DOC_FONT_NAME)
    except Exception:
        pass


def configure_new_document_layout(doc):
    """
    Configura márgenes, encabezado/pie, tamaño de papel y fuente base.
    Estas propiedades reemplazan el uso de una plantilla externa.
    """
    for section in doc.sections:
        section.top_margin = Inches(0.65)
        section.bottom_margin = Inches(0.65)
        section.left_margin = Inches(0.70)
        section.right_margin = Inches(0.70)
        section.header_distance = Inches(0.25)
        section.footer_distance = Inches(0.25)

    set_default_styles(doc)


def add_paragraph(doc, text="", style=None, bold=False, italic=False, align=None):
    if style is None:
        p = doc.add_paragraph()
    else:
        try:
            p = doc.add_paragraph(style=style)
        except Exception:
            p = doc.add_paragraph()

    if text:
        r = p.add_run(str(text))
        r.bold = bold
        r.italic = italic
        r.font.name = DOC_FONT_NAME
        r.font.size = Pt(DOC_FONT_SIZE)

    if align is not None:
        p.alignment = align

    return p


def add_heading_1(doc, text):
    """
    Agrega un capítulo principal iniciando en nueva página.
    """
    try:
        if len(doc.paragraphs) > 0:
            doc.add_page_break()
    except Exception:
        pass

    try:
        p = doc.add_paragraph(style="Heading 1")
    except Exception:
        p = doc.add_paragraph()

    r = p.add_run(text.upper())
    r.bold = True
    r.font.name = DOC_FONT_NAME
    r.font.size = Pt(DOC_HEADING1_SIZE)
    r.font.color.rgb = RGBColor(0, 0, 0)
    return p


def add_heading_2(doc, text):
    try:
        p = doc.add_paragraph(style="Heading 2")
    except Exception:
        p = doc.add_paragraph()
    r = p.add_run(text)
    r.bold = True
    r.font.name = DOC_FONT_NAME
    r.font.size = Pt(DOC_HEADING2_SIZE)
    r.font.color.rgb = RGBColor(0, 0, 0)
    return p


def add_bullet(doc, text):
    try:
        p = doc.add_paragraph(style="List Bullet")
    except Exception:
        p = doc.add_paragraph()
        p.style = doc.styles["Normal"]
    p.add_run(str(text))
    return p


def format_report_value(value, column_name=None):
    """
    Formatea valores del informe según el tipo de dato y la columna.

    Reglas:
    - Conteos/cantidades: sin decimales.
    - Versiones/revisiones: texto, sin conversión numérica.
    - Magnitudes técnicas: dos decimales.
    - Identificadores, escenarios, rutas y textos: se conservan como texto.
    """
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass

    if isinstance(value, bool):
        return "Sí" if value else "No"

    col = "" if column_name is None else str(column_name).strip()
    col_norm = (
        col.lower()
        .replace("á", "a")
        .replace("é", "e")
        .replace("í", "i")
        .replace("ó", "o")
        .replace("ú", "u")
        .replace("ñ", "n")
    )

    text = str(value).strip()

    if text == "":
        return ""

    # Columnas que siempre deben tratarse como texto.
    text_columns_tokens = [
        "id", "barra", "escenario", "revision", "version", "observaciones",
        "realizo", "aprobo", "cliente", "ubicacion", "descripcion",
        "criterio", "unidad", "estado", "tipo", "elemento", "hallazgo",
        "recomendacion", "diagnostico", "archivo", "ruta", "metodo",
        "software", "concepto", "termino", "definicion", "norma", "guia",
        "caso de estudio", "reporte de salida", "configuracion"
    ]

    for token in text_columns_tokens:
        if token in col_norm:
            # Excepción: columna "Tipo de elemento" también es texto.
            return text

    # Columnas de conteo/cantidad que deben ir como entero.
    integer_columns_tokens = [
        "cantidad identificada",
        "barras evaluadas",
        "barras con baja tension",
        "barras con sobretension",
        "barras aceptables",
        "equipos evaluados",
        "equipos sobrecargados",
        "numero de barras",
        "numero",
        "conteo"
    ]

    # Convertir texto numérico si aplica.
    num = None
    try:
        normalized = text.replace(",", "")
        # Acepta enteros y decimales simples.
        if re.fullmatch(r"[-+]?\d+(\.\d+)?", normalized):
            num = float(normalized)
    except Exception:
        num = None

    # Números puros de Python.
    if num is None and isinstance(value, (int, float)):
        try:
            num = float(value)
        except Exception:
            num = None

    if num is None:
        return text

    for token in integer_columns_tokens:
        if token in col_norm:
            return str(int(round(num)))

    # Si el valor es numérico pero la columna parece identificador, conservar texto.
    if col_norm in ["id"]:
        return text

    # Magnitudes técnicas: dos decimales.
    return f"{num:.2f}"


def add_dataframe_table(doc, df, title=None, max_rows=None, preferred_cols=None):
    if title:
        p = add_paragraph(doc)
        r = p.add_run(title)
        r.bold = True
        r.italic = True
        r.font.color.rgb = RGBColor(0, 0, 0)

    if df is None or df.empty:
        add_paragraph(doc, "Sin datos disponibles.")
        return None

    display_df = df.copy()

    if preferred_cols is not None:
        cols = [c for c in preferred_cols if c in display_df.columns]
        if cols:
            display_df = display_df[cols]

    if max_rows is not None and len(display_df) > max_rows:
        display_df = display_df.head(max_rows).copy()

    # Estandarizar valores considerando el contexto de la columna.
    for col in display_df.columns:
        display_df[col] = display_df[col].apply(lambda value: format_report_value(value, col))

    table = doc.add_table(rows=1, cols=len(display_df.columns))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Encabezados
    hdr_cells = table.rows[0].cells
    for i, col_name in enumerate(display_df.columns):
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(str(col_name))
        run.bold = True
        run.font.name = DOC_FONT_NAME
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0, 0, 0)
        hdr_cells[i].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    # Filas
    for _, row in display_df.iterrows():
        cells = table.add_row().cells
        for i, value in enumerate(row):
            p = cells[i].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run("" if value is None else str(value))
            run.font.name = DOC_FONT_NAME
            run.font.size = Pt(9)
            run.font.color.rgb = RGBColor(0, 0, 0)
            cells[i].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    try:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for r in p.runs:
                        r.font.name = DOC_FONT_NAME
                        r.font.size = Pt(9)
                        r.font.color.rgb = RGBColor(0, 0, 0)
    except Exception:
        pass

    try:
        set_table_borders(table)
        keep_table_on_page(table)
    except Exception:
        pass

    doc.add_paragraph()

    return table


def force_all_text_black(doc):
    """
    Fuerza todo el texto del documento a color negro.
    """
    try:
        for paragraph in doc.paragraphs:
            for run in paragraph.runs:
                run.font.color.rgb = RGBColor(0, 0, 0)

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            run.font.color.rgb = RGBColor(0, 0, 0)

        for section in doc.sections:
            for container in [section.header, section.footer]:
                for paragraph in container.paragraphs:
                    for run in paragraph.runs:
                        run.font.color.rgb = RGBColor(0, 0, 0)
                for table in container.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            for paragraph in cell.paragraphs:
                                for run in paragraph.runs:
                                    run.font.color.rgb = RGBColor(0, 0, 0)
    except Exception:
        pass




def add_image_placeholder(doc, title, instruction_text, height_lines=6):
    """
    Inserta un título y un recuadro para pegar imágenes manualmente,
    por ejemplo el diagrama unifilar general o unifilares con resultados.
    """
    try:
        add_heading_2(doc, title)
    except Exception:
        add_paragraph(doc, title, bold=True)

    add_paragraph(doc, instruction_text)

    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"

    cell = table.rows[0].cells[0]
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    placeholder_text = (
        "\n" * height_lines
        + "Espacio reservado para pegar imagen / diagrama"
        + "\n" * height_lines
    )

    try:
        set_cell_text(cell, placeholder_text)
    except Exception:
        cell.text = placeholder_text

    try:
        set_table_borders(table)
        keep_table_on_page(table)
    except Exception:
        pass

    doc.add_paragraph()

    return table


def create_voltage_chart_image(bus_df, output_folder, filename, title, ascending=True, top_n=12):
    """
    Crea un gráfico tipo lollipop horizontal para que el nombre de la barra
    se lea mejor que en un gráfico de barras convencional.
    """
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt

        if bus_df is None or bus_df.empty:
            return ""

        df = bus_df.copy()
        df["Tensión (%)"] = pd.to_numeric(df["Tensión (%)"], errors="coerce")
        df = df.dropna(subset=["Tensión (%)"])

        if df.empty:
            return ""

        df = df.sort_values("Tensión (%)", ascending=ascending).head(top_n).copy()
        df["Etiqueta"] = df["Barra"].astype(str) + " (" + df["Escenario"].astype(str) + ")"

        # Orden visual: el valor más alto o más bajo se ve arriba.
        df = df.sort_values("Tensión (%)", ascending=ascending).reset_index(drop=True)
        df = df.iloc[::-1].reset_index(drop=True)

        y = list(range(len(df)))

        fig_height = max(4.8, 0.48 * len(df) + 1.8)
        fig, ax = plt.subplots(figsize=(10.5, fig_height))

        x_min = max(0, float(df["Tensión (%)"].min()) - 2.0)
        x_max = float(df["Tensión (%)"].max()) + 2.5

        # Lollipop chart
        ax.hlines(y=y, xmin=x_min, xmax=df["Tensión (%)"], linewidth=2.0)
        ax.plot(df["Tensión (%)"], y, "o", markersize=8)

        ax.set_yticks(y)
        ax.set_yticklabels(df["Etiqueta"], fontsize=10)
        ax.set_xlabel("Tensión (%)")
        ax.set_title(title)
        ax.set_xlim(x_min, x_max)
        ax.grid(axis="x", linestyle="--", linewidth=0.5, alpha=0.6)

        for yi, value in zip(y, df["Tensión (%)"]):
            ax.text(float(value) + 0.15, yi, f"{value:.2f}", va="center", fontsize=9)

        plt.tight_layout()
        plt.subplots_adjust(left=0.36)

        os.makedirs(output_folder, exist_ok=True)
        path = os.path.join(output_folder, filename)
        fig.savefig(path, dpi=180, bbox_inches="tight")
        plt.close(fig)

        return path
    except Exception as ex:
        print("No se pudo generar gráfico de tensión:", str(ex))
        return ""


def add_voltage_chart_to_doc(doc, image_path, fallback_text):
    """
    Inserta un gráfico si existe; si no, deja un mensaje para no detener el informe.
    """
    if image_path and os.path.exists(image_path):
        try:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run()
            run.add_picture(image_path, width=Inches(6.7))
            doc.add_paragraph()
            return
        except Exception:
            pass

    add_paragraph(doc, fallback_text)


# =============================================================================
# GENERACION DOCX
# =============================================================================


def update_headers_footers(doc, project_info):
    """
    Crea encabezado y pie de página con estilo técnico sobrio.
    Función requerida por generate_word_report().
    """
    try:
        for section in doc.sections:
            header = section.header
            footer = section.footer

            if not header.paragraphs:
                header.add_paragraph()

            p = header.paragraphs[0]
            p.text = ""
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

            r = p.add_run(
                project_info.get("nombre_proyecto", "")
                + "    |    Informe Técnico - Estudio de Flujo de Potencia    |    "
                + project_info.get("revision", "")
            )
            r.font.name = DOC_FONT_NAME
            r.font.size = Pt(8)
            r.font.color.rgb = RGBColor(0, 0, 0)
            r.bold = True

            if not footer.paragraphs:
                footer.add_paragraph()

            p = footer.paragraphs[0]
            p.text = ""
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

            r = p.add_run(
                project_info.get("estado", "")
                + "    |    "
                + project_info.get("fecha_emision", "")
            )
            r.font.name = DOC_FONT_NAME
            r.font.size = Pt(8)
            r.font.color.rgb = RGBColor(0, 0, 0)

    except Exception:
        pass



def add_page_break(doc):
    doc.add_page_break()


def generate_word_report(e, project_info, scenarios, report_paths, lfr_tables):
    bus_frames = []
    equipment_frames = []

    for scenario_name, report_path in report_paths.items():
        lfr = lfr_tables[scenario_name]

        bus_df = build_bus_voltage_table(scenario_name, lfr)
        equipment_df = build_equipment_table(scenario_name, lfr)

        bus_frames.append(bus_df)
        equipment_frames.append(equipment_df)

    bus_all = pd.concat(bus_frames, ignore_index=True, sort=False)
    equipment_all = pd.concat(equipment_frames, ignore_index=True, sort=False)

    summary_df = create_scenario_summary(scenarios, bus_all, equipment_all)
    comparison_df = create_comparison_summary(summary_df)
    bus_delta_df = create_bus_delta(scenarios, bus_all)
    findings_df = create_findings(bus_all, equipment_all)

    print("Extrayendo información resumida del modelo en modo rápido...")
    model_info, model_sources = extract_model_information(e, bus_all, equipment_all)
    model_summary_df = model_summary_dataframe_fast(model_info)
    print("Información del modelo procesada. Si alguna tabla muestra solo ID, revise los nombres de propiedades disponibles en ETAP API o cambie ENRICH_MODEL_CRITICAL_PROPERTIES a True.")

    first_report_path = next(iter(report_paths.values()))
    output_folder = os.path.dirname(first_report_path)

    output_file = os.path.join(
        output_folder,
        clean_filename("Informe_Flujo_de_Potencia_" + project_info.get("nombre_proyecto", "Proyecto")) + ".docx"
    )

    doc = Document()
    configure_new_document_layout(doc)
    update_headers_footers(doc, project_info)

    # -------------------------------------------------------------------------
    # PORTADA
    # -------------------------------------------------------------------------
    p = add_paragraph(doc, project_info.get("titulo_informe", "INFORME TÉCNICO"), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    for run in p.runs:
        run.font.size = Pt(DOC_TITLE_SIZE)
        run.font.color.rgb = RGBColor(0, 0, 0)

    p = add_paragraph(doc, project_info.get("subtitulo", "ESTUDIO DE FLUJO DE POTENCIA"), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    for run in p.runs:
        run.font.size = Pt(DOC_SUBTITLE_SIZE)
        run.font.color.rgb = RGBColor(0, 0, 0)

    p = add_paragraph(doc, project_info.get("nombre_proyecto", ""), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    for run in p.runs:
        run.font.size = Pt(DOC_SUBTITLE_SIZE)

    doc.add_paragraph()
    doc.add_paragraph()

    portada_df = pd.DataFrame([
        {"Concepto": "Cliente", "Descripción": project_info.get("cliente", "")},
        {"Concepto": "Ubicación del Proyecto", "Descripción": project_info.get("ubicacion", "")},
        {"Concepto": "Preparado por", "Descripción": project_info.get("preparado_por", "")},
        {"Concepto": "Cargo / Especialidad", "Descripción": project_info.get("cargo", "")},
        {"Concepto": "Matrícula profesional", "Descripción": project_info.get("matricula", "")},
        {"Concepto": "Fecha de emisión", "Descripción": project_info.get("fecha_emision", "")},
        {"Concepto": "Revisión", "Descripción": project_info.get("revision", "")},
        {"Concepto": "Estado del documento", "Descripción": project_info.get("estado", "")},
    ])
    add_dataframe_table(doc, portada_df)

    add_paragraph(doc, "Contenido:", bold=True)
    add_bullet(doc, "Flujo de Potencia")

    doc.add_paragraph()

    add_paragraph(doc, "Control de Cambios", bold=True)

    control_df = pd.DataFrame([
        {
            "Versión": str(project_info.get("revision", "Rev. 0").replace("Rev.", "").strip()),
            "Observaciones": "Versión inicial",
            "Realizó": project_info.get("preparado_por", ""),
            "Aprobó": project_info.get("cliente", "")
        }
    ])

    add_dataframe_table(doc, control_df)

    # -------------------------------------------------------------------------
    # INTRODUCCION GENERAL
    # -------------------------------------------------------------------------
    add_heading_1(doc, "1. Introducción General")
    add_heading_2(doc, "1.1. Objeto del informe")
    add_paragraph(
        doc,
        "El presente informe técnico tiene por objeto documentar el desarrollo y los resultados del estudio de flujo de potencia realizado para el proyecto "
        + project_info.get("nombre_proyecto", "")
        + ", ubicado en "
        + project_info.get("ubicacion", "")
        + "."
    )
    add_paragraph(
        doc,
        "El documento ha sido preparado con el propósito de evaluar condiciones relevantes del sistema eléctrico en estado estable, suministrando criterios técnicos, resultados de análisis y recomendaciones orientadas a apoyar decisiones de diseño, validación y operación."
    )

    add_heading_2(doc, "1.2. Alcance del estudio")
    add_paragraph(doc, project_info.get("alcance", ""))
    add_paragraph(
        doc,
        "De manera específica, el estudio comprende la evaluación de perfiles de tensión, cargabilidad de equipos, comportamiento operativo por escenario y comparación de resultados entre las condiciones analizadas."
    )

    add_heading_2(doc, "1.3. Exclusiones y limitaciones")
    add_paragraph(
        doc,
        "El presente informe no contempla actividades de levantamiento de información en campo, pruebas eléctricas, puesta en servicio, supervisión de construcción ni validación física de instalaciones."
    )
    add_paragraph(
        doc,
        "La validez de los resultados aquí presentados depende directamente de la calidad, consistencia y suficiencia de la información suministrada y del modelo eléctrico utilizado. Cualquier cambio posterior en la configuración del sistema, parámetros eléctricos o filosofía de operación podrá requerir la revisión o actualización de los resultados."
    )

    add_heading_2(doc, "1.4. Software y herramientas utilizadas")
    add_paragraph(
        doc,
        "Para el desarrollo del estudio se empleó "
        + project_info.get("software", "ETAP")
        + ", a partir del modelo eléctrico disponible para el proyecto y de los escenarios operativos definidos para esta evaluación."
    )

    add_heading_2(doc, "1.5. Normas y guías técnicas aplicables")
    add_paragraph(
        doc,
        "El estudio de flujo de potencia se desarrolla tomando como referencia buenas prácticas de ingeniería reconocidas para sistemas eléctricos industriales y comerciales. Para flujo de carga, la referencia principal corresponde a IEEE Std 3002.2."
    )
    add_dataframe_table(
        doc,
        normative_references_dataframe(),
        title="Tabla 1-1. Referencias normativas y guías técnicas aplicables"
    )

    # -------------------------------------------------------------------------
    # DEFINICIONES
    # -------------------------------------------------------------------------
    add_heading_1(doc, "2. Definiciones")
    add_paragraph(
        doc,
        "Para una mejor comprensión de la terminología técnica utilizada en el informe, se incluye a continuación una recopilación de los términos principales empleados en el estudio de flujo de potencia."
    )
    add_dataframe_table(
        doc,
        definitions_dataframe(),
        title="Tabla 2-1. Definiciones principales"
    )

    # -------------------------------------------------------------------------
    # RESUMEN EJECUTIVO
    # -------------------------------------------------------------------------
    add_heading_1(doc, "3. Resumen Ejecutivo")
    add_heading_2(doc, "3.1. Resumen de resultados")
    add_paragraph(
        doc,
        "El estudio de flujo de potencia permitió evaluar el comportamiento del sistema eléctrico bajo los escenarios seleccionados, identificando los niveles de tensión en barras, la cargabilidad de los equipos y las condiciones operativas más relevantes."
    )

    add_dataframe_table(
        doc,
        summary_df,
        title="Tabla 3-1. Resumen general de escenarios evaluados",
        preferred_cols=[
            "Escenario",
            "Barras evaluadas",
            "Tensión mínima (%)",
            "Tensión máxima (%)",
            "Tensión promedio (%)",
            "Barras con baja tensión",
            "Barras con sobretensión",
            "Equipos sobrecargados"
        ]
    )

    add_heading_2(doc, "3.2. Hallazgos principales")
    if "Resultado" in findings_df.columns:
        add_paragraph(doc, str(findings_df.iloc[0, 0]))
    else:
        add_dataframe_table(
            doc,
            findings_df,
            title="Tabla 3-2. Hallazgos relevantes",
            max_rows=12,
            preferred_cols=["Escenario", "Tipo", "Elemento", "Valor", "Unidad", "Hallazgo", "Recomendación"]
        )

    # -------------------------------------------------------------------------
    # DESCRIPCION DEL SISTEMA
    # -------------------------------------------------------------------------
    add_heading_1(doc, "4. Descripción del Sistema Eléctrico")
    add_heading_2(doc, "4.1. Descripción general del sistema eléctrico analizado")
    add_paragraph(
        doc,
        "El sistema eléctrico objeto del presente informe corresponde al conjunto de elementos representados en el modelo eléctrico utilizado para el estudio de flujo de potencia. La evaluación considera las fuentes de alimentación, barras, transformadores, alimentadores, cargas y equipos asociados incluidos en los escenarios seleccionados."
    )

    add_heading_2(doc, "4.2. Resumen de elementos identificados en el modelo")
    add_paragraph(
        doc,
        "A partir de la información disponible en ETAP, y de forma complementaria desde los resultados del flujo de carga, se presenta un inventario resumido de los principales elementos eléctricos identificados en el modelo: barras, transformadores, generadores y cargas."
    )
    add_dataframe_table(doc, model_summary_df, title="Tabla 4-1. Resumen de elementos del modelo", preferred_cols=["Tipo de elemento", "Cantidad identificada"])

    add_heading_2(doc, "4.3. Inventario resumido de transformadores")
    if model_info.get("transformadores", pd.DataFrame()).empty:
        add_paragraph(doc, "No se identificaron transformadores mediante las consultas rápidas disponibles. Se recomienda complementar esta sección con el reporte de elementos de ETAP si aplica.")
    else:
        add_dataframe_table(doc, model_info["transformadores"], title="Tabla 4-2. Inventario de transformadores", max_rows=MAX_MODEL_TABLE_ROWS, preferred_cols=["ID"])

    add_heading_2(doc, "4.4. Inventario resumido de generadores")
    if model_info.get("generadores", pd.DataFrame()).empty:
        add_paragraph(doc, "No se identificaron generadores mediante las consultas rápidas disponibles. Se recomienda complementar esta sección con el reporte de elementos de ETAP si aplica.")
    else:
        add_dataframe_table(doc, model_info["generadores"], title="Tabla 4-3. Inventario de generadores", max_rows=MAX_MODEL_TABLE_ROWS, preferred_cols=["ID", "Tipo de elemento", "kW", "kVA", "kV", "FP", "Modo"])

    add_heading_2(doc, "4.5. Inventario resumido de cargas")
    if model_info.get("cargas", pd.DataFrame()).empty:
        add_paragraph(doc, "No se identificaron cargas mediante las consultas rápidas disponibles. Se recomienda complementar esta sección con el reporte de elementos de ETAP si aplica.")
    else:
        add_dataframe_table(doc, model_info["cargas"], title="Tabla 4-4. Inventario de cargas", max_rows=MAX_MODEL_TABLE_ROWS, preferred_cols=["ID", "Tipo de elemento", "kW", "kVA", "kV", "FP", "HP"])

    add_heading_2(doc, "4.6. Diagrama unifilar del modelo")
    add_image_placeholder(
        doc,
        "Espacio para diagrama unifilar general",
        "Pegue en el siguiente recuadro una imagen del diagrama unifilar general del modelo utilizado en ETAP.",
        height_lines=6
    )

    add_heading_2(doc, "4.7. Escenarios de operación considerados")
    scenarios_df = pd.DataFrame([
        {
            "Escenario": s["name"],
            "Revisión": s["revisionName"],
            "Configuración": s["configName"],
            "Caso de estudio": s["studyCase"],
            "Reporte de salida": s["outputReport"]
        }
        for s in scenarios
    ])

    add_dataframe_table(doc, scenarios_df, title="Tabla 4-5. Escenarios de operación evaluados")

    # -------------------------------------------------------------------------
    # DATOS DE ENTRADA Y BASES
    # -------------------------------------------------------------------------
    add_heading_1(doc, "5. Datos de Entrada y Bases de Diseño")
    add_heading_2(doc, "5.1. Datos de entrada")
    add_paragraph(
        doc,
        "Para el desarrollo del estudio de flujo de potencia se utilizaron como datos de entrada el modelo eléctrico del sistema en ETAP, la configuración de los escenarios seleccionados, los parámetros eléctricos de los equipos y las condiciones operativas representadas en el modelo."
    )
    add_paragraph(
        doc,
        "En línea con las recomendaciones de IEEE Std 3002.2 para estudios de flujo de carga, el informe documenta objetivos del estudio, datos de entrada, criterios de aceptación, escenarios de operación, resultados de tensión, cargabilidad, hallazgos, limitaciones y recomendaciones. La precisión del análisis depende de la calidad de los datos del modelo, la representación de las cargas, los modos de operación de fuentes y generadores, y la validación de las condiciones operativas simuladas."
    )

    add_heading_2(doc, "5.2. Criterios de aceptación")
    criteria_df = pd.DataFrame([
        {"Criterio": "Tensión mínima aceptable", "Valor": MIN_VOLTAGE_PERCENT, "Unidad": "%", "Descripción": "Barras por debajo de este valor se clasifican como baja tensión."},
        {"Criterio": "Tensión máxima aceptable", "Valor": MAX_VOLTAGE_PERCENT, "Unidad": "%", "Descripción": "Barras por encima de este valor se clasifican como sobretensión."},
        {"Criterio": "Cargabilidad máxima aceptable", "Valor": MAX_LOADING_PERCENT, "Unidad": "%", "Descripción": "Equipos por encima de este valor se clasifican como sobrecargados, si existe información de cargabilidad."},
        {"Criterio": "Método de solución", "Valor": project_info.get("metodo", "Newton-Raphson"), "Unidad": "", "Descripción": "Método numérico utilizado para resolver el flujo de potencia."},
        {"Criterio": "Software utilizado", "Valor": project_info.get("software", "ETAP"), "Unidad": "", "Descripción": "Herramienta utilizada para ejecutar el estudio."}
    ])

    add_dataframe_table(doc, criteria_df, title="Tabla 5-1. Criterios de aceptación y bases de cálculo")

    add_heading_2(doc, "5.3. Archivos de resultados utilizados")
    db_df = pd.DataFrame([
        {
            "Escenario": name,
            "Archivo de resultados": path
        }
        for name, path in report_paths.items()
    ])
    add_dataframe_table(doc, db_df, title="Tabla 5-2. Archivos de resultados procesados", max_rows=20)

    # -------------------------------------------------------------------------
    # ESTUDIO DE FLUJO DE CARGA
    # -------------------------------------------------------------------------
    add_heading_1(doc, "6. Estudio de Flujo de Carga")
    add_heading_2(doc, "6.1. Objetivo")
    add_paragraph(
        doc,
        "El presente capítulo tiene por objeto documentar el desarrollo del estudio de flujo de carga, con el fin de evaluar el comportamiento del sistema eléctrico en estado estable bajo los escenarios operativos definidos."
    )
    add_paragraph(
        doc,
        "El estudio permite determinar la distribución de potencia activa y reactiva en la red, los niveles de tensión en las barras, la cargabilidad de los equipos principales y las condiciones generales de operación del sistema dentro del alcance analizado."
    )

    add_heading_2(doc, "6.2. Método de solución")
    add_paragraph(
        doc,
        "El estudio fue desarrollado utilizando el método numérico "
        + project_info.get("metodo", "Newton-Raphson")
        + ", por tratarse de una metodología robusta, precisa y ampliamente utilizada para la solución de sistemas eléctricos de potencia en régimen permanente."
    )
    add_paragraph(
        doc,
        "El problema de flujo de carga corresponde a un sistema de ecuaciones no lineales en el cual, a partir de la topología de la red, las impedancias de los elementos, la generación disponible y la demanda conectada, se busca determinar el estado operativo del sistema bajo condiciones estables de operación."
    )

    add_heading_2(doc, "6.3. Resultados por escenario")
    add_dataframe_table(
        doc,
        summary_df,
        title="Tabla 6-1. Resumen de resultados de flujo de carga por escenario",
        preferred_cols=[
            "Escenario",
            "Barras evaluadas",
            "Tensión mínima (%)",
            "Tensión máxima (%)",
            "Tensión promedio (%)",
            "Barras con baja tensión",
            "Barras con sobretensión",
            "Equipos sobrecargados",
            "Diagnóstico"
        ]
    )

    add_heading_2(doc, "6.4. Perfil de tensión por barra")
    add_dataframe_table(
        doc,
        bus_all.sort_values(["Escenario", "Tensión (%)"], ascending=[True, True]),
        title="Tabla 6-2. Perfil de tensión por barra",
        max_rows=MAX_MAIN_VOLTAGE_ROWS,
        preferred_cols=["Escenario", "Barra", "kV nominal", "Tensión (%)", "Tensión (pu)", "Ángulo (deg)", "Estado"]
    )

    high_voltage_chart = create_voltage_chart_image(
        bus_all,
        output_folder,
        "grafico_barras_tensiones_altas.png",
        "Barras con mayor tensión calculada",
        ascending=False,
        top_n=12
    )
    add_voltage_chart_to_doc(
        doc,
        high_voltage_chart,
        "No fue posible generar el gráfico de barras con mayor tensión."
    )

    add_heading_2(doc, "6.5. Barras con menor tensión")
    critical_low = bus_all.sort_values("Tensión (%)", ascending=True).head(20)
    add_dataframe_table(
        doc,
        critical_low,
        title="Tabla 6-3. Barras con menor tensión calculada",
        preferred_cols=["Escenario", "Barra", "kV nominal", "Tensión (%)", "Tensión (pu)", "Estado"]
    )

    low_voltage_chart = create_voltage_chart_image(
        bus_all,
        output_folder,
        "grafico_barras_tensiones_bajas.png",
        "Barras con menor tensión calculada",
        ascending=True,
        top_n=12
    )
    add_voltage_chart_to_doc(
        doc,
        low_voltage_chart,
        "No fue posible generar el gráfico de barras con menor tensión."
    )

    add_heading_2(doc, "6.6. Análisis comparativo entre escenarios")
    add_paragraph(
        doc,
        "El análisis comparativo permite identificar variaciones de tensión entre el escenario base y los demás escenarios evaluados, facilitando la identificación de condiciones operativas que incrementan o reducen los márgenes eléctricos del sistema."
    )
    add_dataframe_table(
        doc,
        comparison_df,
        title="Tabla 6-4. Comparativo general entre escenarios",
        max_rows=60
    )

    add_dataframe_table(
        doc,
        bus_delta_df,
        title="Tabla 6-5. Barras con mayor variación de tensión entre escenarios",
        max_rows=20,
        preferred_cols=["Barra", "Escenario base", "Escenario comparado", "Delta tensión (%)", "Delta ángulo (deg)", "Interpretación"]
    )

    add_heading_2(doc, "6.7. Análisis de cargabilidad")
    add_paragraph(
        doc,
        "La evaluación de cargabilidad permite determinar si los transformadores, conductores, alimentadores y demás elementos del sistema operan dentro de márgenes admisibles bajo los escenarios considerados."
    )
    equipment_report = equipment_all.copy()
    if "TYPE" in equipment_report.columns:
        equipment_report = equipment_report[~equipment_report["TYPE"].astype(str).isin(["0", "0.0"])]
    if "Estado Cargabilidad" in equipment_report.columns:
        overloaded_report = equipment_report[equipment_report["Estado Cargabilidad"] == "Sobrecargado"].copy()
    else:
        overloaded_report = pd.DataFrame()

    if overloaded_report.empty:
        add_paragraph(
            doc,
            "No se identificaron equipos sobrecargados con la información de cargabilidad disponible en los resultados LFR. Por esta razón, no se incluye una tabla extensa de ramas informativas sin porcentaje de cargabilidad."
        )
    else:
        add_dataframe_table(
            doc,
            overloaded_report,
            title="Tabla 6-6. Equipos sobrecargados identificados desde resultados LFR",
            max_rows=40,
            preferred_cols=["Escenario", "Desde", "Hasta", "TYPE", "kV", "Cargabilidad (%)", "Estado Cargabilidad"]
        )

    add_heading_2(doc, "6.8. Hallazgos relevantes")
    if "Resultado" in findings_df.columns:
        add_paragraph(doc, str(findings_df.iloc[0, 0]))
    else:
        add_dataframe_table(
            doc,
            findings_df,
            title="Tabla 6-7. Hallazgos y recomendaciones",
            max_rows=30,
            preferred_cols=["Escenario", "Tipo", "Elemento", "Valor", "Unidad", "Hallazgo", "Recomendación"]
        )

    add_heading_2(doc, "6.9. Conclusiones del estudio de flujo de carga")
    for item in create_conclusions(summary_df):
        add_bullet(doc, item)

    # -------------------------------------------------------------------------
    # ANEXOS
    # -------------------------------------------------------------------------
    add_heading_1(doc, "Anexo A. Datos completos procesados")
    add_paragraph(
        doc,
        "El presente anexo incluye una muestra de los datos completos procesados desde la tabla LFR de los archivos de resultados generados por ETAP. Para trazabilidad detallada, se recomienda conservar los archivos .LF1S junto con este informe."
    )

    add_heading_2(doc, "Anexo A.1. Diagramas unifilares con resultados de flujo de potencia")
    for scenario in scenarios:
        add_image_placeholder(
            doc,
            "Unifilar con resultados - " + scenario["name"],
            "Pegue en el siguiente recuadro la captura del diagrama unifilar de ETAP correspondiente al escenario " + scenario["name"] + ", incluyendo los resultados visibles de flujo de potencia.",
            height_lines=5
        )

    add_dataframe_table(
        doc,
        bus_all,
        title="Tabla A-1. Perfil de tensión completo",
        max_rows=MAX_ANNEX_VOLTAGE_ROWS,
        preferred_cols=["Escenario", "Barra", "kV nominal", "Tensión (%)", "Tensión (pu)", "Ángulo (deg)", "Estado"]
    )

    print("Guardando archivo Word...")
    force_all_text_black(doc)
    doc.save(output_file)
    print("Archivo Word guardado.")

    return output_file


# =============================================================================
# EJECUCION PRINCIPAL
# =============================================================================

def main():
    print("")
    print("============================================================")
    print("INFORME TÉCNICO DE FLUJO DE POTENCIA - ETAP")
    print("============================================================")

    print("El informe Word se generará desde cero con propiedades de documento basadas en el archivo de referencia.")
    e = connect_to_etap()

    project_info = ask_project_information()
    ask_acceptance_criteria()
    search_folder = ask_search_folder()

    options, discovery_rows = discover_project_options(e, search_folder)
    options = edit_options_window(options)
    scenarios = select_scenarios_workflow(options)

    print("Proyecto:", project_info.get("nombre_proyecto", ""))
    print("Escenarios seleccionados:", " vs ".join([s["name"] for s in scenarios]))
    print("")

    search_roots = [
        os.getcwd(),
        os.path.dirname(os.getcwd())
    ]

    if search_folder and os.path.exists(search_folder):
        search_roots.append(search_folder)

    report_paths = {}
    lfr_tables = {}

    for scenario in scenarios:
        report_path = run_load_flow(e, scenario, search_roots)

        report_folder = os.path.dirname(report_path)
        if report_folder not in search_roots:
            search_roots.append(report_folder)

        scenario_name = scenario["name"]
        report_paths[scenario_name] = report_path

        print("")
        print("Leyendo tabla LFR para:", scenario_name)

        lfr = get_lfr_table(report_path)
        lfr_tables[scenario_name] = lfr

        print("Filas LFR:", len(lfr))
        print("Columnas LFR:", ", ".join([str(c) for c in lfr.columns]))

    output = generate_word_report(e, project_info, scenarios, report_paths, lfr_tables)

    print("")
    print("============================================================")
    print("INFORME WORD GENERADO CORRECTAMENTE")
    print("============================================================")
    print(output)
    print("")


if __name__ == "__main__":
    try:
        main()
    except Exception as ex:
        print("")
        print("ERROR:")
        print(str(ex))
        print("")
        traceback.print_exc()
        print("")
        input("Presione Enter para cerrar...")
