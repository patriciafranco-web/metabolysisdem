from datetime import date, datetime, timedelta
import io
import csv
import os
import re
import json
import math
import requests
from flask import Flask, render_template, request, redirect, url_for, Response, send_file, session, send_from_directory, flash
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import text
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from functools import wraps
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# Importación opcional para leer archivos de Word (.docx)
try:
    import docx
except ImportError:
    docx = None

# Importación opcional para leer archivos PDF (.pdf)
try:
    from pypdf import PdfReader as _PdfReader
except ImportError:
    try:
        from PyPDF2 import PdfReader as _PdfReader
    except ImportError:
        _PdfReader = None

app = Flask(__name__)

# --- CONFIGURACIÓN DE DISCO SEGURO PARA RENDER ---
DATA_DIR = '/var/data' if os.path.exists('/var/data') else '.'

app.config['SQLALCHEMY_DATABASE_URI'] = f"sqlite:///{os.path.join(DATA_DIR, 'metabolisys_pro.db')}"
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.secret_key = 'metabolisys_secreto_seguro_2026'

UPLOAD_FOLDER = os.path.join(DATA_DIR, 'uploads')
ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'webp', 'docx', 'pdf'}
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)
# -------------------------------------------------

PLAN_LIMITS_IA = {
    'basico': 250,
    'profesional': 600,
    'clinica': 1500,
    'admin': 999999
}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def format_text(text_val):
    if not text_val:
        return ""
    lines = text_val.split('\n')
    formatted_lines = [" ".join([w.capitalize() for w in line.split()]) for line in lines if line.strip()]
    return "\n".join(formatted_lines)

# --- FUNCIONES BLINDADAS PARA NÚMEROS ---
def parse_float(value):
    if value is None or value == "":
        return None
    try:
        return float(str(value).replace(',', '.').strip())
    except:
        return None

def parse_int(value):
    if value is None or value == "":
        return None
    try:
        return int(float(str(value).replace(',', '.').strip()))
    except:
        return None
# ----------------------------------------

db = SQLAlchemy(app)

class User(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(50), unique=True, nullable=False)
    full_name = db.Column(db.String(100), nullable=True)
    password_hash = db.Column(db.Text, nullable=False)
    expiry_date = db.Column(db.String(20), nullable=True)
    plan_name = db.Column(db.String(30), nullable=True, default='basico')
    monthly_ai_limit = db.Column(db.Integer, nullable=True, default=250)

class AiUsageLog(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
    year_month = db.Column(db.String(7), nullable=False)  # Ej: "2026-11"
    extractions_count = db.Column(db.Integer, nullable=False, default=0)
    monthly_limit = db.Column(db.Integer, nullable=False, default=250)
    updated_at = db.Column(db.String(30), nullable=True)

def get_current_user():
    uid = session.get('user_id')
    return User.query.get(uid) if uid else None

def _limite_ia_usuario(user):
    if not user:
        return 250
    if user.username == 'admin':
        return PLAN_LIMITS_IA['admin']
    if user.monthly_ai_limit and user.monthly_ai_limit > 0:
        return user.monthly_ai_limit
    plan_key = (user.plan_name or 'basico').lower().strip()
    return PLAN_LIMITS_IA.get(plan_key, 250)

def obtener_o_crear_uso_ia(user):
    ym = date.today().strftime('%Y-%m')
    limite = _limite_ia_usuario(user)
    uso = AiUsageLog.query.filter_by(user_id=user.id, year_month=ym).first()
    if not uso:
        uso = AiUsageLog(
            user_id=user.id,
            year_month=ym,
            extractions_count=0,
            monthly_limit=limite,
            updated_at=datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        )
        db.session.add(uso)
        db.session.commit()
    elif uso.monthly_limit != limite:
        uso.monthly_limit = limite
        db.session.commit()
    return uso

def puede_usar_ia(user, unidades_necesarias=1):
    if not user:
        return False, None
    if user.username == 'admin':
        return True, obtener_o_crear_uso_ia(user)
    uso = obtener_o_crear_uso_ia(user)
    restante = (uso.monthly_limit or 250) - (uso.extractions_count or 0)
    return restante >= unidades_necesarias, uso

def registrar_uso_ia(user, unidades=1):
    if not user or unidades <= 0:
        return
    uso = obtener_o_crear_uso_ia(user)
    uso.extractions_count = (uso.extractions_count or 0) + int(unidades)
    uso.updated_at = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    db.session.commit()
    print(f"USO IA [{uso.year_month}] usuario={user.username}: {uso.extractions_count}/{uso.monthly_limit}", flush=True)

# --- EXTRACTOR CLÍNICO CON IA (CLAUDE SONNET 5.5 + FALLBACK / ANTHROPIC) ---
MODELOS_CLAUDE = [
    "claude-sonnet-5-5",
    "claude-sonnet-5.5",
    "claude-sonnet-4-6",
    "claude-haiku-4-5-20251001"
]

def _limpiar_y_cargar_json(raw_text):
    texto = (raw_text or "").strip()
    if not texto:
        return None
    if texto.startswith("```"):
        texto = texto.replace("```json", "").replace("```", "").strip()
    inicio = texto.find("{")
    fin = texto.rfind("}")
    if inicio != -1 and fin != -1 and fin > inicio:
        texto = texto[inicio:fin + 1]
    try:
        return json.loads(texto)
    except Exception:
        arreglado = texto.replace("\r\n", "\\n").replace("\n", "\\n").replace("\r", "\\n")
        if arreglado.count('"') % 2 != 0:
            arreglado += '"'
        if arreglado.count("{") > arreglado.count("}"):
            arreglado += "}" * (arreglado.count("{") - arreglado.count("}"))
        if arreglado.count("[") > arreglado.count("]"):
            arreglado += "]" * (arreglado.count("[") - arreglado.count("]"))
        try:
            return json.loads(arreglado)
        except Exception:
            return None

def extraer_datos_con_ia(texto_consulta, contar_uso=True):
    api_key = os.environ.get('ANTHROPIC_API_KEY', '').strip()
    if not api_key:
        session['last_ia_error'] = "Falta la variable ANTHROPIC_API_KEY en Render."
        print("IA ERROR: No se encontró ANTHROPIC_API_KEY en Render.", flush=True)
        return {}
    if not texto_consulta or not texto_consulta.strip():
        return {}

    prompt = f"""
Eres un asistente médico experto en extracción fiel de datos clínicos en español (Paraguay/Latinoamérica).
Lee la siguiente nota de consulta médica tal como está escrita y extrae ÚNICAMENTE los valores que aparezcan explícitamente en el texto.
NO inventes, NO completes por deducción ni asumas valores que no estén escritos.
Respeta el texto original y las abreviaturas tal como aparecen en la nota.

Reglas de interpretación estricta:
1. Para valores numéricos devuelve solo el número con punto decimal (ej. 76.25), sin unidades. Si hay dos pesos mencionados (ej. uno de evaluación previa y otro "hoy"), toma el peso actual/del día en `weight` y deja el detalle completo en `notes`.
2. Fechas cortas como `DD.MM.AA` o `DD/MM/AA` (por ejemplo `27.03.26` -> `2026-03-27`, `29.04.26` -> `2026-04-29`): si hay fecha explícita de laboratorio distinta a la consulta, ponla en `lab_date` en formato `YYYY-MM-DD`; si no, `null`.
3. `Inbody` / `InBody` / `Eval nutr Inbody`: NO es un estudio de imagen ni va a `otros_estudios`. Es composición corporal / balanza:
   - Si menciona `Inbody` o `InBody`, pon `"InBody"` en `marca_balanza`.
   - El peso de la consulta/día va a `weight` (si hay un peso de evaluación previa como `Eval nutr Inbody 26.03.26: peso 97.4 kg` y otro `Peso hoy 27.03.26: 98.3`, toma el peso del día/consulta en `weight` y conserva todo el detalle en `notes`).
   - `MM` o masa muscular -> `muscle_mass_kg`
   - `Grasa ...%` o `%grasa` -> `body_fat_pct`
   - `GV` o grasa visceral -> `visceral_fat`
4. Abreviaturas comunes en consultas/evaluaciones:
   - `PA` o presión arterial -> `pa`
   - `FC` -> `fc`
   - `Hb` o hemoglobina -> `hemoglobin` (no lo mandes a `otros_laboratorios` si es un valor numérico)
   - `Fe` o hierro sérico / ferremia -> `iron`
   - `Ferritina` -> `ferritin`
   - `GOT` / `AST` -> `got`; `GPT` / `ALT` -> `gpt` (extrae solo los números, sin duplicar en `otros_laboratorios`)
   - `Tirze` / tirzepatida / fármacos y dosis -> `medications`
   - Plan, indicaciones, dieta, pedidos -> `indicaciones_tratamiento`
   - Próximo control / retorno (ej. `Control en 1 mes`) -> `proximo_control`
   - Plaquetas: si vienen en miles completos (ej. 250000 o 250.000), conviértelo a 10^3/uL (ej. 250); si ya está en cientos (ej. 250), déjalo igual.
   - `otros_estudios`: úsalo ÚNICAMENTE para estudios complementarios como ecografías, Doppler, tomografías, resonancias, holter, MAPA, etc., NUNCA para datos de InBody (peso, MM, %grasa, GV).

Devuelve EXCLUSIVAMENTE un objeto JSON válido (sin texto adicional ni bloques markdown) con estas claves exactas:
{{
  "lab_date": "string YYYY-MM-DD o null",
  "marca_balanza": "string o null",
  "pa": "string o null",
  "fc": "int o null",
  "height": "float en cm o null",
  "weight": "float en kg o null",
  "waist": "float en cm o null",
  "body_fat_pct": "float o null",
  "body_fat_kg": "float o null",
  "visceral_fat": "float o null",
  "muscle_mass_kg": "float o null",
  "basal_metabolism": "float o null",
  "hemoglobin": "float o null",
  "glucose": "float o null",
  "insulin": "float o null",
  "hba1c": "float o null",
  "triglycerides": "float o null",
  "cholesterol_tot": "float o null",
  "hdl": "float o null",
  "ldl": "float o null",
  "vldl": "float o null",
  "uric_acid": "float o null",
  "got": "float o null",
  "gpt": "float o null",
  "fa": "float o null",
  "bili_tot": "float o null",
  "bili_dir": "float o null",
  "bili_ind": "float o null",
  "urea": "float o null",
  "creatinine": "float o null",
  "tsh": "float o null",
  "t4": "float o null",
  "t3": "float o null",
  "anti_tpo": "float o null",
  "vit_d": "float o null",
  "vit_b12": "float o null",
  "iron": "float o null",
  "ferritin": "float o null",
  "calcium": "float o null",
  "magnesium": "float o null",
  "phosphorus": "float o null",
  "chlorine": "float o null",
  "platelets": "float o null",
  "otros_laboratorios": "string o null",
  "ecografia_abdominal": "string o null",
  "ecografia_tiroidea": "string o null",
  "otros_estudios": "string o null",
  "medications": "string o null",
  "indicaciones_tratamiento": "string o null",
  "proximo_control": "string o null"
}}

NOTA MÉDICA:
\"\"\"{texto_consulta}\"\"\"
"""
    url = "https://api" + ".anthropic" + ".com/v1/messages"
    headers = {
        "x-api-key": api_key,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json"
    }

    ultimo_error = ""
    for modelo in MODELOS_CLAUDE:
        payload = {
            "model": modelo,
            "max_tokens": 1400,
            "messages": [
                {"role": "user", "content": prompt}
            ]
        }
        try:
            resp = requests.post(url, headers=headers, json=payload, timeout=20)
            if resp.status_code == 200:
                session.pop('last_ia_error', None)
                print(f"IA OK usando modelo: {modelo}", flush=True)
                content_blocks = resp.json().get("content", [])
                raw_text = "".join(
                    block.get("text", "") for block in content_blocks if block.get("type") == "text"
                ).strip()
                parsed = _limpiar_y_cargar_json(raw_text)
                if isinstance(parsed, dict):
                    if contar_uso:
                        registrar_uso_ia(get_current_user(), 1)
                    return parsed
            else:
                ultimo_error = f"HTTP {resp.status_code} ({modelo}): {resp.text[:180]}"
                print(f"IA ERROR {ultimo_error}", flush=True)
        except Exception as e:
            ultimo_error = f"Excepción ({modelo}): {e}"
            print(f"IA EXCEPTION {ultimo_error}", flush=True)

    session['last_ia_error'] = ultimo_error
    return {}


def _normalizar_fecha_corta(match):
    d, m, y = match.group(1), match.group(2), match.group(3)
    if len(y) == 2:
        y = "20" + y
    try:
        dt = datetime(int(y), int(m), int(d))
        return dt.strftime("%Y-%m-%d")
    except Exception:
        return f"{y}-{m.zfill(2)}-{d.zfill(2)}"


def _separar_bloques_por_fecha(texto, fecha_por_defecto):
    patron_fecha = re.compile(r'\b(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})\b')
    lineas = [ln.strip() for ln in (texto or "").splitlines() if ln.strip()]
    if not lineas:
        return []

    bloques = []
    fecha_actual = None
    lineas_actuales = []

    for ln in lineas:
        es_encabezado_fecha = bool(re.match(r'^(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})\b', ln))
        m = patron_fecha.search(ln)
        if es_encabezado_fecha and m:
            if lineas_actuales:
                bloques.append((fecha_actual or fecha_por_defecto, "\n".join(lineas_actuales)))
                lineas_actuales = []
            fecha_actual = _normalizar_fecha_corta(m)
            lineas_actuales.append(ln)
        else:
            if fecha_actual is None and m and not ln.lower().startswith(("eval", "lab")):
                fecha_actual = _normalizar_fecha_corta(m)
            lineas_actuales.append(ln)

    if lineas_actuales:
        bloques.append((fecha_actual or fecha_por_defecto, "\n".join(lineas_actuales)))

    return bloques


def _separar_labs_por_fecha(texto, fecha_por_defecto):
    lineas = [ln.strip() for ln in (texto or "").splitlines() if ln.strip()]
    if not lineas:
        return []

    patron_lab_fecha = re.compile(
        r'^\s*lab(?:oratorio)?s?\s*[:\-]?\s*(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})\b',
        re.IGNORECASE
    )
    bloques_lab = []
    actual_fecha = None
    actual_lineas = []

    for ln in lineas:
        m = patron_lab_fecha.match(ln)
        if m:
            if actual_fecha is not None:
                bloques_lab.append((actual_fecha, "\n".join(actual_lineas).strip()))
            actual_fecha = _normalizar_fecha_corta(m)
            actual_lineas = [ln]
        elif actual_fecha is not None:
            if re.match(r'^(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})\b', ln) or ln.lower().startswith("plan"):
                bloques_lab.append((actual_fecha, "\n".join(actual_lineas).strip()))
                actual_fecha = None
                actual_lineas = []
            else:
                actual_lineas.append(ln)

    if actual_fecha is not None and actual_lineas:
        bloques_lab.append((actual_fecha, "\n".join(actual_lineas).strip()))

    fechas_unicas = {f for f, _ in bloques_lab}
    if len(fechas_unicas) > 1:
        return bloques_lab
    return []


def extraer_historial_por_fecha_con_ia(texto_historial, fecha_por_defecto):
    api_key = os.environ.get('ANTHROPIC_API_KEY', '').strip()
    if not api_key:
        session['last_ia_error'] = "Falta la variable ANTHROPIC_API_KEY en Render."
        print("IA ERROR: No se encontró ANTHROPIC_API_KEY en Render.", flush=True)
        return []
    texto = (texto_historial or "").strip()
    if not texto:
        return []

    bloques = _separar_bloques_por_fecha(texto, fecha_por_defecto)
    if not bloques:
        bloques = [(fecha_por_defecto, texto)]

    consultas_resultado = []
    procesados_ia = 0
    for fecha_bloque, texto_bloque in bloques[:6]:
        datos_ia = extraer_datos_con_ia(texto_bloque, contar_uso=False)
        item = dict(datos_ia) if isinstance(datos_ia, dict) else {}
        if datos_ia:
            procesados_ia += 1
        item["date"] = fecha_bloque
        item["notes"] = texto_bloque
        consultas_resultado.append(item)

    if procesados_ia > 0:
        registrar_uso_ia(get_current_user(), procesados_ia)

    return consultas_resultado
# -----------------------------------------

class Patient(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=True)
    full_name = db.Column(db.String(100), nullable=True)
    ci = db.Column(db.String(20), unique=True, nullable=True)
    birth_date = db.Column(db.String(20))
    age = db.Column(db.Integer, nullable=True, default=0)
    gender = db.Column(db.String(10), nullable=True, default='Mujer')
    phone = db.Column(db.String(20))
    sanatorio = db.Column(db.String(100), nullable=True, default='Consultorio Principal')
    procedencia = db.Column(db.String(100))
    profession = db.Column(db.String(100))
    pathology_type = db.Column(db.String(100))
    evolution_years = db.Column(db.Integer)
    
    seguro_medico = db.Column(db.String(100))
    fecha_inicio_tratamiento = db.Column(db.String(50))

    antecedentes_familiares = db.Column(db.Text)
    antecedentes_personales = db.Column(db.Text)
    comorbidities = db.Column(db.String(200))
    tratamiento_base = db.Column(db.Text)
    historial_cirugias = db.Column(db.Text)
    alergia_farmacos = db.Column(db.Text)
    cant_hijos = db.Column(db.String(50))
    menarca = db.Column(db.String(50))
    menstruacion_regular = db.Column(db.String(50))
    anticonceptivos = db.Column(db.String(50))
    checklist_glp1_notas = db.Column(db.Text)
    calidad_sueno = db.Column(db.String(100))
    actividad_fisica = db.Column(db.String(100))
    fuma_alcohol = db.Column(db.String(100))
    veces_come = db.Column(db.String(100))
    consume_gaseosa_jugo = db.Column(db.String(100))

    records = db.relationship('MedicalRecord', backref='patient', lazy=True, cascade="all, delete-orphan")

class MedicalRecord(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    patient_id = db.Column(db.Integer, db.ForeignKey('patient.id'), nullable=False)
    date = db.Column(db.String(20), nullable=False)
    lab_date = db.Column(db.String(20), nullable=True)
    
    marca_balanza = db.Column(db.String(50))
    pa = db.Column(db.String(20))
    fc = db.Column(db.Integer)
    
    height = db.Column(db.Float)
    weight = db.Column(db.Float)
    body_fat_kg = db.Column(db.Float)
    body_fat_pct = db.Column(db.Float)
    visceral_fat = db.Column(db.Float)
    muscle_mass_kg = db.Column(db.Float)
    muscle_mass_pct = db.Column(db.Float)
    basal_metabolism = db.Column(db.Float)
    waist = db.Column(db.Float)
    
    masa_grasa_kg = db.Column(db.Float)
    porcentaje_musculo = db.Column(db.Float)
    porcentaje_musculo_esqueletico = db.Column(db.Float)
    masa_muscular_esqueletica = db.Column(db.Float)
    porcentaje_oseo = db.Column(db.Float)
    masa_osea = db.Column(db.Float)
    porcentaje_proteina = db.Column(db.Float)
    masa_proteica = db.Column(db.Float)
    porcentaje_agua_corporal = db.Column(db.Float)
    masa_agua_corporal = db.Column(db.Float)
    peso_sin_grasa = db.Column(db.Float)
    grasa_subcutanea_pct = db.Column(db.Float)
    edad_metabolica = db.Column(db.Integer)
    
    weight_lost = db.Column(db.Float, default=0.0)
    fat_lost = db.Column(db.Float, default=0.0)
    muscle_delta = db.Column(db.Float, default=0.0)
    waist_lost = db.Column(db.Float, default=0.0)
    
    hemoglobin = db.Column(db.Float)
    glucose = db.Column(db.Float)
    insulin = db.Column(db.Float)
    hba1c = db.Column(db.Float)
    triglycerides = db.Column(db.Float)
    cholesterol_tot = db.Column(db.Float)
    hdl = db.Column(db.Float)
    ldl = db.Column(db.Float)
    vldl = db.Column(db.Float)
    uric_acid = db.Column(db.Float)
    got = db.Column(db.Float)
    gpt = db.Column(db.Float)
    fa = db.Column(db.Float)
    bili_tot = db.Column(db.Float)
    bili_dir = db.Column(db.Float)
    bili_ind = db.Column(db.Float)
    urea = db.Column(db.Float)
    creatinine = db.Column(db.Float)
    tsh = db.Column(db.Float)
    t4 = db.Column(db.Float)
    t3 = db.Column(db.Float)
    anti_tpo = db.Column(db.Float)
    vit_d = db.Column(db.Float)
    vit_b12 = db.Column(db.Float)
    iron = db.Column(db.Float)
    ferritin = db.Column(db.Float)
    calcium = db.Column(db.Float)
    magnesium = db.Column(db.Float)
    phosphorus = db.Column(db.Float)
    chlorine = db.Column(db.Float)
    platelets = db.Column(db.Float)
    
    otros_laboratorios = db.Column(db.Text)

    ecografia_abdominal = db.Column(db.Text)
    ecografia_tiroidea = db.Column(db.Text)
    otros_estudios = db.Column(db.Text)
    
    indicaciones_tratamiento = db.Column(db.Text)
    medications = db.Column(db.Text)
    proximo_control = db.Column(db.String(30))
    notes = db.Column(db.Text)
    conducta_medica = db.Column(db.Text)

    @property
    def bmi(self):
        if self.weight and self.height and self.height > 0:
            h_m = self.height / 100.0 if self.height > 3 else self.height
            return round(self.weight / (h_m ** 2), 2)
        return 0

    @property
    def ica_value(self):
        if self.waist and self.height and self.height > 0:
            h_cm = self.height if self.height > 3 else self.height * 100
            return round(self.waist / h_cm, 2)
        return None

    @property
    def ica_status(self):
        val = self.ica_value
        if not val: return ("Sin datos", "secondary")
        if val < 0.5: return ("Saludable (Riesgo Bajo)", "success")
        elif 0.5 <= val <= 0.6: return ("Sobrepeso Abdominal (Moderado)", "warning")
        else: return ("Obesidad Central (Alto Riesgo)", "danger")

    @property
    def cardio_risk_estimate(self):
        score = 0
        if self.cholesterol_tot and self.cholesterol_tot > 240: score += 2
        elif self.cholesterol_tot and self.cholesterol_tot > 200: score += 1
        if self.pa:
            try:
                sys_bp = int(self.pa.split('/')[0])
                if sys_bp >= 140: score += 2
                elif sys_bp >= 130: score += 1
            except: pass
        if score >= 3: return ("Riesgo Cardiovascular Alto", "danger")
        elif score == 2: return ("Riesgo Cardiovascular Moderado", "warning")
        else: return ("Riesgo Cardiovascular Bajo / Estable", "success")
        
    @property
    def homair(self):
        if self.insulin and self.glucose and self.insulin > 0 and self.glucose > 0:
            return round((self.insulin * self.glucose) / 405, 2)
        return None

    @property
    def tg_hdl(self):
        if self.triglycerides is not None and self.hdl is not None and self.hdl > 0:
            return round(self.triglycerides / self.hdl, 2)
        return None

    @property
    def fib4(self):
        try:
            if self.patient.age and self.got and self.gpt and self.platelets and self.platelets > 0 and self.gpt > 0:
                resultado = (self.patient.age * self.got) / (self.platelets * math.sqrt(self.gpt))
                return round(resultado, 2)
        except:
            pass
        return None

class PatientPhoto(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    patient_id = db.Column(db.Integer, db.ForeignKey('patient.id'), nullable=False)
    record_id = db.Column(db.Integer, db.ForeignKey('medical_record.id'), nullable=True)
    filename = db.Column(db.String(200), nullable=False)
    category = db.Column(db.String(50), nullable=False)
    date = db.Column(db.String(20), nullable=False)
    notes = db.Column(db.Text)
    record = db.relationship('MedicalRecord', backref='photos', lazy=True)

Patient.photos = db.relationship('PatientPhoto', backref='patient', lazy=True, cascade="all, delete-orphan")

def recalculate_deltas(patient_id):
    patient = Patient.query.get(patient_id)
    if not patient: return
    
    valid_records = sorted([r for r in patient.records if r.weight is not None and r.weight > 0], key=lambda x: x.date)
    if not valid_records: return
    
    w_init = next((r.weight for r in valid_records if r.weight), 0)
    f_init = next((r.body_fat_kg for r in valid_records if r.body_fat_kg), 0)
    m_init = next((r.muscle_mass_kg for r in valid_records if r.muscle_mass_kg), 0)
    wst_init = next((r.waist for r in valid_records if r.waist), 0)
    
    for r in valid_records:
        r.weight_lost = round(w_init - (r.weight or 0), 2)
        r.fat_lost = round(f_init - (r.body_fat_kg or 0), 2)
        r.muscle_delta = round((r.muscle_mass_kg or 0) - m_init, 2)
        r.waist_lost = round(wst_init - (r.waist or 0), 2)
        
    db.session.commit()

with app.app_context():
    db.create_all()

    for alter_sql in [
        'ALTER TABLE medical_record ADD COLUMN platelets FLOAT',
        'ALTER TABLE medical_record ADD COLUMN hemoglobin FLOAT',
        'ALTER TABLE medical_record ADD COLUMN iron FLOAT',
        'ALTER TABLE medical_record ADD COLUMN ferritin FLOAT',
        'ALTER TABLE patient ADD COLUMN user_id INTEGER',
        'ALTER TABLE user ADD COLUMN full_name VARCHAR(100)',
        "ALTER TABLE user ADD COLUMN plan_name VARCHAR(30) DEFAULT 'basico'",
        'ALTER TABLE user ADD COLUMN monthly_ai_limit INTEGER DEFAULT 250'
    ]:
        try:
            db.session.execute(text(alter_sql))
            db.session.commit()
        except Exception:
            db.session.rollback()

    admin_user = User.query.filter_by(username='admin').first()
    if not admin_user:
        hashed_pw = generate_password_hash('metabosys2026', method='pbkdf2:sha256')
        admin_user = User(username='admin', full_name='Administrador', password_hash=hashed_pw, expiry_date=None, plan_name='admin', monthly_ai_limit=999999)
        db.session.add(admin_user)
        db.session.commit()
    else:
        if not admin_user.full_name:
            admin_user.full_name = 'Administrador'
        admin_user.plan_name = 'admin'
        admin_user.monthly_ai_limit = 999999
        db.session.commit()

    # Cliente Plan Básico activado con 30 días de vencimiento (250 extracciones IA/mes)
    PLAN_BASICO_TAG_FILE = os.path.join(DATA_DIR, '.plan_basico_30d_v1')
    exp_30d = (date.today() + timedelta(days=30)).strftime('%Y-%m-%d')

    dra_muller = User.query.filter_by(username='alejandra.muller').first()
    if not dra_muller:
        hashed_pw = generate_password_hash('dra12345', method='pbkdf2:sha256')
        dra_muller = User(
            username='alejandra.muller',
            full_name='Dra. Alejandra Müller',
            password_hash=hashed_pw,
            expiry_date=exp_30d,
            plan_name='basico',
            monthly_ai_limit=250
        )
        db.session.add(dra_muller)
        db.session.commit()
        with open(PLAN_BASICO_TAG_FILE, 'w') as f:
            f.write('applied')
    else:
        if not dra_muller.full_name:
            dra_muller.full_name = 'Dra. Alejandra Müller'
        dra_muller.plan_name = 'basico'
        dra_muller.monthly_ai_limit = 250
        if not os.path.exists(PLAN_BASICO_TAG_FILE):
            dra_muller.expiry_date = exp_30d
            with open(PLAN_BASICO_TAG_FILE, 'w') as f:
                f.write('applied')
        db.session.commit()

    if dra_muller:
        Patient.query.filter(Patient.user_id.is_(None)).update({Patient.user_id: dra_muller.id}, synchronize_session=False)
        db.session.commit()

    dra_camila = User.query.filter_by(username='camila.chaparro').first()
    if not dra_camila:
        hashed_pw_camila = generate_password_hash('camila12345', method='pbkdf2:sha256')
        exp_camila = (date.today() + timedelta(days=7)).strftime('%Y-%m-%d')
        dra_camila = User(
            username='camila.chaparro',
            full_name='Dra. Camila Chaparro',
            password_hash=hashed_pw_camila,
            expiry_date=exp_camila,
            plan_name='basico',
            monthly_ai_limit=250
        )
        db.session.add(dra_camila)
        db.session.commit()
    elif not dra_camila.full_name:
        dra_camila.full_name = 'Dra. Camila Chaparro'
        db.session.commit()

    dra_tamires = User.query.filter_by(username='tamires.soster').first()
    if not dra_tamires:
        hashed_pw_tamires = generate_password_hash('tamires12345', method='pbkdf2:sha256')
        exp_tamires = (date.today() + timedelta(days=7)).strftime('%Y-%m-%d')
        dra_tamires = User(
            username='tamires.soster',
            full_name='Dra. Tamires Soster',
            password_hash=hashed_pw_tamires,
            expiry_date=exp_tamires,
            plan_name='basico',
            monthly_ai_limit=250
        )
        db.session.add(dra_tamires)
        db.session.commit()
    elif not dra_tamires.full_name:
        dra_tamires.full_name = 'Dra. Tamires Soster'
        db.session.commit()

def get_patient_for_current_user(patient_id):
    patient = Patient.query.get_or_404(patient_id)
    current_user = get_current_user()
    if not current_user:
        return None
    if current_user.username == 'admin':
        return patient
    if patient.user_id != current_user.id:
        return None
    return patient

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        
        user = User.query.get(session.get('user_id'))
        if user and user.expiry_date:
            try:
                exp_date = datetime.strptime(user.expiry_date, '%Y-%m-%d').date()
                if date.today() > exp_date:
                    session.clear()
                    return render_template('login.html', error='Su periodo ha vencido. Contacte a MEDYSIS PY para renovar.')
            except:
                pass

        return f(*args, **kwargs)
    return decorated_function

@app.route('/login', methods=['GET', 'POST'])
def login():
    error = None
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '').strip()
        user = User.query.filter_by(username=username).first()
        
        if user and check_password_hash(user.password_hash, password):
            if user.expiry_date:
                try:
                    exp_date = datetime.strptime(user.expiry_date, '%Y-%m-%d').date()
                    if date.today() > exp_date:
                        return render_template('login.html', error='Su periodo ha vencido. Contacte a MEDYSIS PY para renovar.')
                except:
                    pass

            session['user_id'] = user.id
            session['username'] = username
            
            if user.full_name:
                session['user_name'] = user.full_name
            elif username == 'alejandra.muller':
                session['user_name'] = 'Dra. Alejandra Müller'
            elif username == 'camila.chaparro':
                session['user_name'] = 'Dra. Camila Chaparro'
            elif username == 'tamires.soster':
                session['user_name'] = 'Dra. Tamires Soster'
            else:
                session['user_name'] = 'Administrador'
                
            return redirect(url_for('index'))
        else:
            error = 'Usuario o contraseña incorrectos.'
    return render_template('login.html', error=error)

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

@app.route('/')
@login_required
def index():
    current_user = get_current_user()
    search_query = request.args.get('q', '').strip()
    sanatorio_filter = request.args.get('sanatorio', '').strip()

    query = Patient.query
    if current_user and current_user.username != 'admin':
        query = query.filter_by(user_id=current_user.id)

    if sanatorio_filter:
        query = query.filter_by(sanatorio=sanatorio_filter)
    if search_query:
        query = query.filter((Patient.full_name.ilike(f'%{search_query}%')) | (Patient.ci.ilike(f'%{search_query}%')))

    patients = query.all()

    base_san = Patient.query
    if current_user and current_user.username != 'admin':
        base_san = base_san.filter_by(user_id=current_user.id)
    sanatorios = [s[0] for s in base_san.with_entities(Patient.sanatorio).distinct().all() if s[0]]

    uso_ia = obtener_o_crear_uso_ia(current_user) if current_user else None

    return render_template(
        'index.html',
        patients=patients,
        sanatorios=sanatorios,
        current_sanatorio=sanatorio_filter,
        search_query=search_query,
        uso_ia=uso_ia,
        current_user=current_user
    )

@app.route('/patient/new', methods=['GET', 'POST'])
@login_required
def new_patient():
    current_user = get_current_user()
    if request.method == 'POST':
        ci_input = request.form.get('ci', '').strip()
        if not ci_input:
            ci_input = f"SC-{datetime.now().strftime('%Y%m%d%H%M%S')}"
        elif Patient.query.filter_by(ci=ci_input).first():
            return f"<script>alert('❌ ERROR: Ya existe un paciente registrado con la Cédula {ci_input}. No se puede duplicar.'); window.history.back();</script>"

        nombre_input = format_text(request.form.get('full_name', '')) or 'Paciente Sin Nombre'
        sanatorio_input = format_text(request.form.get('sanatorio', '')) or 'Consultorio Principal'
        genero_input = request.form.get('gender') or 'Mujer'

        patient = Patient(
            user_id=current_user.id if current_user else None,
            full_name=nombre_input,
            ci=ci_input,
            birth_date=request.form.get('birth_date'),
            age=parse_int(request.form.get('age')),
            gender=genero_input,
            phone=request.form.get('phone'),
            sanatorio=sanatorio_input,
            procedencia=format_text(request.form.get('procedencia')),
            profession=format_text(request.form.get('profession')),
            pathology_type=format_text(request.form.get('pathology_type')),
            evolution_years=parse_int(request.form.get('evolution_years')),
            seguro_medico=format_text(request.form.get('seguro_medico')),
            fecha_inicio_tratamiento=format_text(request.form.get('fecha_inicio_tratamiento')),
            antecedentes_familiares=format_text(request.form.get('antecedentes_familiares')),
            antecedentes_personales=format_text(request.form.get('antecedentes_personales')),
            comorbidities=format_text(request.form.get('comorbidities')),
            tratamiento_base=format_text(request.form.get('tratamiento_base')),
            historial_cirugias=format_text(request.form.get('historial_cirugias')),
            alergia_farmacos=format_text(request.form.get('alergia_farmacos')),
            cant_hijos=request.form.get('cant_hijos'),
            menarca=format_text(request.form.get('menarca')),
            menstruacion_regular=request.form.get('menstruacion_regular'),
            anticonceptivos=format_text(request.form.get('anticonceptivos')),
            checklist_glp1_notas=format_text(request.form.get('checklist_glp1_notas')),
            calidad_sueno=format_text(request.form.get('calidad_sueno')),
            actividad_fisica=format_text(request.form.get('actividad_fisica')),
            fuma_alcohol=format_text(request.form.get('fuma_alcohol')),
            veces_come=format_text(request.form.get('veces_come')),
            consume_gaseosa_jugo=format_text(request.form.get('consume_gaseosa_jugo'))
        )
        db.session.add(patient)
        db.session.commit()
        return redirect(url_for('patient_detail', id=patient.id))
    return render_template('new_patient.html')

@app.route('/patient/edit/<int:id>', methods=['GET', 'POST'])
@login_required
def edit_patient(id):
    patient = get_patient_for_current_user(id)
    if not patient:
        flash("No tiene permisos para acceder a este paciente.", "warning")
        return redirect(url_for('index'))
    if request.method == 'POST':
        ci_input = request.form.get('ci', '').strip()
        existing = Patient.query.filter_by(ci=ci_input).first()
        if existing and existing.id != patient.id:
            return f"<script>alert('❌ ERROR: La Cédula {ci_input} ya le pertenece a otro paciente en el sistema.'); window.history.back();</script>"
        
        patient.full_name = format_text(request.form.get('full_name', patient.full_name))
        patient.ci = ci_input or patient.ci
        patient.birth_date = request.form.get('birth_date') or patient.birth_date
        patient.age = parse_int(request.form.get('age')) or patient.age
        patient.gender = request.form.get('gender') or patient.gender
        patient.phone = request.form.get('phone') or patient.phone
        patient.sanatorio = format_text(request.form.get('sanatorio', patient.sanatorio))
        patient.procedencia = format_text(request.form.get('procedencia')) or patient.procedencia
        patient.profession = format_text(request.form.get('profession')) or patient.profession
        patient.pathology_type = format_text(request.form.get('pathology_type')) or patient.pathology_type
        patient.evolution_years = parse_int(request.form.get('evolution_years')) or patient.evolution_years
        patient.seguro_medico = format_text(request.form.get('seguro_medico')) or patient.seguro_medico
        patient.fecha_inicio_tratamiento = format_text(request.form.get('fecha_inicio_tratamiento')) or patient.fecha_inicio_tratamiento
        
        patient.antecedentes_familiares = format_text(request.form.get('antecedentes_familiares')) or patient.antecedentes_familiares
        patient.antecedentes_personales = format_text(request.form.get('antecedentes_personales')) or patient.antecedentes_personales
        patient.comorbidities = format_text(request.form.get('comorbidities')) or patient.comorbidities
        patient.tratamiento_base = format_text(request.form.get('tratamiento_base')) or patient.tratamiento_base
        patient.historial_cirugias = format_text(request.form.get('historial_cirugias')) or patient.historial_cirugias
        patient.alergia_farmacos = format_text(request.form.get('alergia_farmacos')) or patient.alergia_farmacos
        
        patient.cant_hijos = request.form.get('cant_hijos') or patient.cant_hijos
        patient.menarca = format_text(request.form.get('menarca')) or patient.menarca
        patient.menstruacion_regular = request.form.get('menstruacion_regular') or patient.menstruacion_regular
        patient.anticonceptivos = format_text(request.form.get('anticonceptivos')) or patient.anticonceptivos
        patient.checklist_glp1_notas = format_text(request.form.get('checklist_glp1_notas')) or patient.checklist_glp1_notas
        patient.calidad_sueno = format_text(request.form.get('calidad_sueno')) or patient.calidad_sueno
        patient.actividad_fisica = format_text(request.form.get('actividad_fisica')) or patient.actividad_fisica
        patient.fuma_alcohol = format_text(request.form.get('fuma_alcohol')) or patient.fuma_alcohol
        patient.veces_come = format_text(request.form.get('veces_come')) or patient.veces_come
        patient.consume_gaseosa_jugo = format_text(request.form.get('consume_gaseosa_jugo')) or patient.consume_gaseosa_jugo

        db.session.commit()
        return redirect(url_for('patient_detail', id=patient.id))
    return render_template('edit_patient.html', patient=patient)

@app.route('/patient/delete/<int:id>', methods=['POST'])
@login_required
def delete_patient(id):
    patient = get_patient_for_current_user(id)
    if not patient:
        flash("No tiene permisos para eliminar este paciente.", "warning")
        return redirect(url_for('index'))
    db.session.delete(patient)
    db.session.commit()
    return redirect(url_for('index'))

@app.route('/patient/<int:id>', methods=['GET', 'POST'])
@login_required
def patient_detail(id):
    patient = get_patient_for_current_user(id)
    if not patient:
        flash("No tiene permisos para acceder a este paciente.", "warning")
        return redirect(url_for('index'))
    current_user = get_current_user()
    all_records = sorted(patient.records, key=lambda x: (x.date or '', x.id), reverse=True)
    today_str = date.today().strftime('%Y-%m-%d')

    if request.method == 'POST':
        fecha = request.form.get('date') or today_str
        nota_libre = request.form.get('nota_libre', '').strip()

        if nota_libre:
            bloques_labs = _separar_labs_por_fecha(nota_libre, fecha)
            unidades_req = len(bloques_labs) if bloques_labs else 1

            permitido, uso_ia = puede_usar_ia(current_user, unidades_req)
            if not permitido:
                record = MedicalRecord(
                    patient_id=patient.id,
                    date=fecha,
                    height=next((r.height for r in all_records if r.height), None),
                    notes=nota_libre
                )
                db.session.add(record)
                db.session.commit()
                flash(
                    f"⚠️ Consulta guardada en notas, pero alcanzó su límite mensual de IA ({uso_ia.extractions_count}/{uso_ia.monthly_limit} extracciones).",
                    "warning"
                )
                return redirect(url_for('patient_detail', id=patient.id))

            altura_base = next((r.height for r in all_records if r.height), None)

            if bloques_labs:
                creados = 0
                alguna_ia_ok = False
                ultimo_err = ""
                for fecha_lab_item, texto_lab_item in bloques_labs:
                    datos = extraer_datos_con_ia(texto_lab_item, contar_uso=True)
                    if datos:
                        alguna_ia_ok = True
                    else:
                        ultimo_err = session.pop('last_ia_error', '') or ultimo_err

                    altura_item = parse_float(datos.get('height')) or altura_base
                    if altura_item:
                        altura_base = altura_item

                    record = MedicalRecord(
                        patient_id=patient.id,
                        date=fecha_lab_item,
                        lab_date=datos.get('lab_date') or fecha_lab_item,
                        marca_balanza=datos.get('marca_balanza'),
                        pa=datos.get('pa'),
                        fc=parse_int(datos.get('fc')),
                        height=altura_item,
                        weight=parse_float(datos.get('weight')),
                        waist=parse_float(datos.get('waist')),
                        body_fat_pct=parse_float(datos.get('body_fat_pct')),
                        body_fat_kg=parse_float(datos.get('body_fat_kg')),
                        visceral_fat=parse_float(datos.get('visceral_fat')),
                        muscle_mass_kg=parse_float(datos.get('muscle_mass_kg')),
                        basal_metabolism=parse_float(datos.get('basal_metabolism')),
                        hemoglobin=parse_float(datos.get('hemoglobin')),
                        glucose=parse_float(datos.get('glucose')),
                        insulin=parse_float(datos.get('insulin')),
                        hba1c=parse_float(datos.get('hba1c')),
                        triglycerides=parse_float(datos.get('triglycerides')),
                        cholesterol_tot=parse_float(datos.get('cholesterol_tot')),
                        hdl=parse_float(datos.get('hdl')),
                        ldl=parse_float(datos.get('ldl')),
                        vldl=parse_float(datos.get('vldl')),
                        uric_acid=parse_float(datos.get('uric_acid')),
                        got=parse_float(datos.get('got')),
                        gpt=parse_float(datos.get('gpt')),
                        fa=parse_float(datos.get('fa')),
                        bili_tot=parse_float(datos.get('bili_tot')),
                        bili_dir=parse_float(datos.get('bili_dir')),
                        bili_ind=parse_float(datos.get('bili_ind')),
                        urea=parse_float(datos.get('urea')),
                        creatinine=parse_float(datos.get('creatinine')),
                        tsh=parse_float(datos.get('tsh')),
                        t4=parse_float(datos.get('t4')),
                        t3=parse_float(datos.get('t3')),
                        anti_tpo=parse_float(datos.get('anti_tpo')),
                        vit_d=parse_float(datos.get('vit_d')),
                        vit_b12=parse_float(datos.get('vit_b12')),
                        iron=parse_float(datos.get('iron')),
                        ferritin=parse_float(datos.get('ferritin')),
                        calcium=parse_float(datos.get('calcium')),
                        magnesium=parse_float(datos.get('magnesium')),
                        phosphorus=parse_float(datos.get('phosphorus')),
                        chlorine=parse_float(datos.get('chlorine')),
                        platelets=parse_float(datos.get('platelets')),
                        otros_laboratorios=datos.get('otros_laboratorios'),
                        ecografia_abdominal=datos.get('ecografia_abdominal'),
                        ecografia_tiroidea=datos.get('ecografia_tiroidea'),
                        otros_estudios=datos.get('otros_estudios'),
                        medications=datos.get('medications'),
                        indicaciones_tratamiento=datos.get('indicaciones_tratamiento'),
                        proximo_control=datos.get('proximo_control'),
                        notes=texto_lab_item
                    )
                    db.session.add(record)
                    creados += 1

                db.session.commit()
                recalculate_deltas(patient.id)

                uso_actual = obtener_o_crear_uso_ia(current_user) if current_user else None
                cupo_txt = f" (Cupo IA mes: {uso_actual.extractions_count}/{uso_actual.monthly_limit})" if uso_actual and current_user.username != 'admin' else ""
                if alguna_ia_ok:
                    flash(f"✅ Se separaron y clasificaron {creados} laboratorio(s) por fecha en la planilla{cupo_txt}.", "success")
                elif ultimo_err:
                    flash(f"⚠️ Registros guardados, pero la IA no pudo clasificarlos ({ultimo_err}).", "warning")
                else:
                    flash("✅ Registros guardados por fecha.", "info")
            else:
                datos = extraer_datos_con_ia(nota_libre, contar_uso=True)

                altura_final = parse_float(datos.get('height')) or altura_base

                record = MedicalRecord(
                    patient_id=patient.id,
                    date=fecha,
                    lab_date=datos.get('lab_date'),
                    marca_balanza=datos.get('marca_balanza'),
                    pa=datos.get('pa'),
                    fc=parse_int(datos.get('fc')),
                    height=altura_final,
                    weight=parse_float(datos.get('weight')),
                    waist=parse_float(datos.get('waist')),
                    body_fat_pct=parse_float(datos.get('body_fat_pct')),
                    body_fat_kg=parse_float(datos.get('body_fat_kg')),
                    visceral_fat=parse_float(datos.get('visceral_fat')),
                    muscle_mass_kg=parse_float(datos.get('muscle_mass_kg')),
                    basal_metabolism=parse_float(datos.get('basal_metabolism')),
                    hemoglobin=parse_float(datos.get('hemoglobin')),
                    glucose=parse_float(datos.get('glucose')),
                    insulin=parse_float(datos.get('insulin')),
                    hba1c=parse_float(datos.get('hba1c')),
                    triglycerides=parse_float(datos.get('triglycerides')),
                    cholesterol_tot=parse_float(datos.get('cholesterol_tot')),
                    hdl=parse_float(datos.get('hdl')),
                    ldl=parse_float(datos.get('ldl')),
                    vldl=parse_float(datos.get('vldl')),
                    uric_acid=parse_float(datos.get('uric_acid')),
                    got=parse_float(datos.get('got')),
                    gpt=parse_float(datos.get('gpt')),
                    fa=parse_float(datos.get('fa')),
                    bili_tot=parse_float(datos.get('bili_tot')),
                    bili_dir=parse_float(datos.get('bili_dir')),
                    bili_ind=parse_float(datos.get('bili_ind')),
                    urea=parse_float(datos.get('urea')),
                    creatinine=parse_float(datos.get('creatinine')),
                    tsh=parse_float(datos.get('tsh')),
                    t4=parse_float(datos.get('t4')),
                    t3=parse_float(datos.get('t3')),
                    anti_tpo=parse_float(datos.get('anti_tpo')),
                    vit_d=parse_float(datos.get('vit_d')),
                    vit_b12=parse_float(datos.get('vit_b12')),
                    iron=parse_float(datos.get('iron')),
                    ferritin=parse_float(datos.get('ferritin')),
                    calcium=parse_float(datos.get('calcium')),
                    magnesium=parse_float(datos.get('magnesium')),
                    phosphorus=parse_float(datos.get('phosphorus')),
                    chlorine=parse_float(datos.get('chlorine')),
                    platelets=parse_float(datos.get('platelets')),
                    otros_laboratorios=datos.get('otros_laboratorios'),
                    ecografia_abdominal=datos.get('ecografia_abdominal'),
                    ecografia_tiroidea=datos.get('ecografia_tiroidea'),
                    otros_estudios=datos.get('otros_estudios'),
                    medications=datos.get('medications'),
                    indicaciones_tratamiento=datos.get('indicaciones_tratamiento'),
                    proximo_control=datos.get('proximo_control'),
                    notes=nota_libre
                )
                db.session.add(record)
                db.session.commit()
                recalculate_deltas(patient.id)

                err_ia = session.pop('last_ia_error', '')
                uso_actual = obtener_o_crear_uso_ia(current_user) if current_user else None
                cupo_txt = f" (Cupo IA mes: {uso_actual.extractions_count}/{uso_actual.monthly_limit})" if uso_actual and current_user.username != 'admin' else ""
                if datos:
                    flash(f"✅ Consulta guardada y datos clasificados automáticamente con IA en la planilla{cupo_txt}.", "success")
                elif err_ia:
                    flash(f"⚠️ Consulta guardada, pero la IA no pudo clasificar los datos ({err_ia}).", "warning")
                else:
                    flash("✅ Consulta guardada (sin clasificación IA: verifique ANTHROPIC_API_KEY si desea extracción automática).", "info")

        return redirect(url_for('patient_detail', id=patient.id))

    last_record = all_records[0] if all_records else None
    valid_records = sorted([r for r in patient.records if r.weight is not None and r.weight > 0], key=lambda x: x.date)
    dates = [r.date for r in valid_records]
    hba1c_data = [r.hba1c for r in valid_records if r.hba1c]
    weight_data = [r.weight for r in valid_records if r.weight]
    uso_ia = obtener_o_crear_uso_ia(current_user) if current_user else None

    return render_template(
        'patient_detail.html',
        patient=patient,
        all_records=all_records,
        last_record=last_record,
        valid_records=valid_records,
        dates=dates,
        hba1c_data=hba1c_data,
        weight_data=weight_data,
        today_str=today_str,
        uso_ia=uso_ia
    )

# --- RUTA PARA SUBIR Y PREVISUALIZAR ARCHIVO WORD (.docx) O PDF (.pdf) ---
@app.route('/patient/import_docx/<int:patient_id>', methods=['GET', 'POST'])
@login_required
def import_docx(patient_id):
    patient = get_patient_for_current_user(patient_id)
    if not patient:
        flash("No tiene permisos para acceder a este paciente.", "warning")
        return redirect(url_for('index'))
    current_user = get_current_user()

    session_key_text = f"docx_preview_text_{patient.id}"
    session_key_file = f"docx_preview_file_{patient.id}"

    def extraer_texto_docx(doc_obj):
        partes = []
        for p in doc_obj.paragraphs:
            if p.text and p.text.strip():
                partes.append(p.text.strip())
        for table in doc_obj.tables:
            for row in table.rows:
                celdas = [c.text.strip() for c in row.cells if c.text and c.text.strip()]
                celdas_limpias = []
                for c in celdas:
                    if not celdas_limpias or celdas_limpias[-1] != c:
                        celdas_limpias.append(c)
                if celdas_limpias:
                    partes.append(" | ".join(celdas_limpias))
        return "\n".join(partes)

    def extraer_texto_pdf(pdf_bytes):
        if _PdfReader is None:
            raise RuntimeError("Falta instalar pypdf o PyPDF2 en requirements.txt para leer archivos .pdf.")
        reader = _PdfReader(io.BytesIO(pdf_bytes))
        paginas = []
        for page in reader.pages:
            txt = page.extract_text() or ""
            if txt.strip():
                paginas.append(txt.strip())
        return "\n".join(paginas)

    if request.method == 'POST':
        action = request.form.get('action')

        if action == 'upload':
            file = request.files.get('word_file')
            if not file or not file.filename:
                flash("Seleccione un archivo .docx o .pdf antes de continuar.", "warning")
                return redirect(url_for('import_docx', patient_id=patient.id))

            fname_lower = file.filename.lower()
            es_docx = fname_lower.endswith('.docx')
            es_pdf = fname_lower.endswith('.pdf')

            if not es_docx and not es_pdf:
                flash("Solo se admiten archivos con extensión .docx o .pdf.", "warning")
                return redirect(url_for('import_docx', patient_id=patient.id))

            if es_docx and docx is None:
                flash("Falta instalar python-docx en el servidor (agréguelo a requirements.txt).", "danger")
                return redirect(url_for('import_docx', patient_id=patient.id))

            try:
                raw_bytes = file.read()
                if not raw_bytes:
                    flash("El archivo seleccionado está vacío o no pudo leerse desde el equipo/nube.", "warning")
                    return redirect(url_for('import_docx', patient_id=patient.id))

                if es_docx:
                    doc = docx.Document(io.BytesIO(raw_bytes))
                    extracted_text = extraer_texto_docx(doc)
                else:
                    extracted_text = extraer_texto_pdf(raw_bytes)

                if not extracted_text.strip():
                    flash("El documento se abrió, pero no contiene texto legible.", "warning")
                    return redirect(url_for('import_docx', patient_id=patient.id))

                filename = secure_filename(f"historial_{patient.id}_{file.filename}")
                try:
                    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                    with open(filepath, 'wb') as f_out:
                        f_out.write(raw_bytes)
                except Exception as save_err:
                    print(f"AVISO guardando copia archivo: {save_err}", flush=True)

                session[session_key_text] = extracted_text
                session[session_key_file] = filename
                return redirect(url_for('import_docx', patient_id=patient.id))
            except Exception as e:
                print(f"ERROR LEYENDO ARCHIVO: {e}", flush=True)
                flash(f"No se pudo leer el archivo: {e}", "danger")
                return redirect(url_for('import_docx', patient_id=patient.id))

        elif action == 'confirm':
            final_text = request.form.get('final_text', '').strip()
            session.pop(session_key_text, None)
            session.pop(session_key_file, None)

            if final_text:
                today_str = date.today().strftime('%Y-%m-%d')
                existentes = sorted(patient.records, key=lambda x: (x.date or '', x.id), reverse=True)
                altura_base = next((r.height for r in existentes if r.height), None)

                permitido, uso_ia = puede_usar_ia(current_user, 1)
                if not permitido:
                    record = MedicalRecord(
                        patient_id=patient.id,
                        date=today_str,
                        height=altura_base,
                        weight=None,
                        notes=final_text,
                        indicaciones_tratamiento="Historial importado sin IA por límite mensual alcanzado."
                    )
                    db.session.add(record)
                    db.session.commit()
                    flash(
                        f"⚠️ Historial guardado en notas, pero alcanzó su límite mensual de IA ({uso_ia.extractions_count}/{uso_ia.monthly_limit} extracciones).",
                        "warning"
                    )
                    return redirect(url_for('patient_detail', id=patient.id))

                consultas = extraer_historial_por_fecha_con_ia(final_text, today_str)

                if consultas:
                    consultas_ordenadas = sorted(consultas, key=lambda x: str(x.get('date') or today_str))
                    creados = 0
                    for item in consultas_ordenadas:
                        fecha_item = str(item.get('date') or today_str).strip() or today_str
                        altura_item = parse_float(item.get('height')) or altura_base
                        if altura_item:
                            altura_base = altura_item

                        record = MedicalRecord(
                            patient_id=patient.id,
                            date=fecha_item,
                            lab_date=item.get('lab_date'),
                            marca_balanza=item.get('marca_balanza'),
                            pa=item.get('pa'),
                            fc=parse_int(item.get('fc')),
                            height=altura_item,
                            weight=parse_float(item.get('weight')),
                            waist=parse_float(item.get('waist')),
                            body_fat_pct=parse_float(item.get('body_fat_pct')),
                            body_fat_kg=parse_float(item.get('body_fat_kg')),
                            visceral_fat=parse_float(item.get('visceral_fat')),
                            muscle_mass_kg=parse_float(item.get('muscle_mass_kg')),
                            basal_metabolism=parse_float(item.get('basal_metabolism')),
                            hemoglobin=parse_float(item.get('hemoglobin')),
                            glucose=parse_float(item.get('glucose')),
                            insulin=parse_float(item.get('insulin')),
                            hba1c=parse_float(item.get('hba1c')),
                            triglycerides=parse_float(item.get('triglycerides')),
                            cholesterol_tot=parse_float(item.get('cholesterol_tot')),
                            hdl=parse_float(item.get('hdl')),
                            ldl=parse_float(item.get('ldl')),
                            vldl=parse_float(item.get('vldl')),
                            uric_acid=parse_float(item.get('uric_acid')),
                            got=parse_float(item.get('got')),
                            gpt=parse_float(item.get('gpt')),
                            fa=parse_float(item.get('fa')),
                            bili_tot=parse_float(item.get('bili_tot')),
                            bili_dir=parse_float(item.get('bili_dir')),
                            bili_ind=parse_float(item.get('bili_ind')),
                            urea=parse_float(item.get('urea')),
                            creatinine=parse_float(item.get('creatinine')),
                            tsh=parse_float(item.get('tsh')),
                            t4=parse_float(item.get('t4')),
                            t3=parse_float(item.get('t3')),
                            anti_tpo=parse_float(item.get('anti_tpo')),
                            vit_d=parse_float(item.get('vit_d')),
                            vit_b12=parse_float(item.get('vit_b12')),
                            iron=parse_float(item.get('iron')),
                            ferritin=parse_float(item.get('ferritin')),
                            calcium=parse_float(item.get('calcium')),
                            magnesium=parse_float(item.get('magnesium')),
                            phosphorus=parse_float(item.get('phosphorus')),
                            chlorine=parse_float(item.get('chlorine')),
                            platelets=parse_float(item.get('platelets')),
                            otros_laboratorios=item.get('otros_laboratorios'),
                            ecografia_abdominal=item.get('ecografia_abdominal'),
                            ecografia_tiroidea=item.get('ecografia_tiroidea'),
                            otros_estudios=item.get('otros_estudios'),
                            medications=item.get('medications'),
                            indicaciones_tratamiento=item.get('indicaciones_tratamiento'),
                            proximo_control=item.get('proximo_control'),
                            notes=item.get('notes') or final_text
                        )
                        db.session.add(record)
                        creados += 1

                    db.session.commit()
                    recalculate_deltas(patient.id)
                    uso_actual = obtener_o_crear_uso_ia(current_user) if current_user else None
                    cupo_txt = f" (Cupo IA mes: {uso_actual.extractions_count}/{uso_actual.monthly_limit})" if uso_actual and current_user.username != 'admin' else ""
                    flash(f"✅ Historial procesado: se cargaron {creados} consulta(s) por fecha en la planilla{cupo_txt}.", "success")
                else:
                    record = MedicalRecord(
                        patient_id=patient.id,
                        date=today_str,
                        height=altura_base,
                        weight=None,
                        notes=final_text,
                        indicaciones_tratamiento="Ver historial importado en notas."
                    )
                    db.session.add(record)
                    db.session.commit()
                    err_ia = session.pop('last_ia_error', '')
                    if err_ia:
                        flash(f"⚠️ Se guardó el texto, pero la IA no pudo clasificar los datos ({err_ia}).", "warning")
                    else:
                        flash("✅ ¡Historial importado exitosamente!", "success")

                return redirect(url_for('patient_detail', id=patient.id))

    preview_text = session.get(session_key_text, '')
    filename_saved = session.get(session_key_file, '')
    uso_ia = obtener_o_crear_uso_ia(current_user) if current_user else None
    return render_template('import_docx.html', patient=patient, preview_text=preview_text, filename_saved=filename_saved, uso_ia=uso_ia)

@app.route('/record/new/<int:patient_id>', methods=['GET', 'POST'])
@login_required
def new_record(patient_id):
    patient = get_patient_for_current_user(patient_id)
    if not patient:
        flash("No tiene permisos para acceder a este paciente.", "warning")
        return redirect(url_for('index'))
    if request.method == 'POST':
        record = MedicalRecord(
            patient_id=patient.id,
            date=request.form['date'],
            lab_date=request.form.get('lab_date'),
            marca_balanza=request.form.get('marca_balanza'),
            pa=request.form.get('pa'),
            fc=parse_int(request.form.get('fc')),
            height=parse_float(request.form.get('height')),
            weight=parse_float(request.form.get('weight')),
            body_fat_kg=parse_float(request.form.get('body_fat_kg') or request.form.get('masa_grasa_kg')),
            body_fat_pct=parse_float(request.form.get('body_fat_pct') or request.form.get('porcentaje_grasa_corporal')),
            visceral_fat=parse_float(request.form.get('visceral_fat') or request.form.get('grasa_visceral')),
            muscle_mass_kg=parse_float(request.form.get('muscle_mass_kg') or request.form.get('masa_muscular_esqueletica')),
            muscle_mass_pct=parse_float(request.form.get('muscle_mass_pct') or request.form.get('porcentaje_musculo_esqueletico')),
            basal_metabolism=parse_float(request.form.get('basal_metabolism') or request.form.get('tasa_metabolica_basal')),
            waist=parse_float(request.form.get('waist')),
            masa_grasa_kg=parse_float(request.form.get('masa_grasa_kg')),
            porcentaje_musculo=parse_float(request.form.get('porcentaje_musculo')),
            porcentaje_musculo_esqueletico=parse_float(request.form.get('porcentaje_musculo_esqueletico')),
            masa_muscular_esqueletica=parse_float(request.form.get('masa_muscular_esqueletica')),
            porcentaje_oseo=parse_float(request.form.get('porcentaje_oseo')),
            masa_osea=parse_float(request.form.get('masa_osea')),
            porcentaje_proteina=parse_float(request.form.get('porcentaje_proteina')),
            masa_proteica=parse_float(request.form.get('masa_proteica')),
            porcentaje_agua_corporal=parse_float(request.form.get('porcentaje_agua_corporal')),
            masa_agua_corporal=parse_float(request.form.get('masa_agua_corporal')),
            peso_sin_grasa=parse_float(request.form.get('peso_sin_grasa')),
            grasa_subcutanea_pct=parse_float(request.form.get('grasa_subcutanea_pct')),
            edad_metabolica=parse_int(request.form.get('edad_metabolica')),
            hemoglobin=parse_float(request.form.get('hemoglobin')),
            glucose=parse_float(request.form.get('glucose')),
            insulin=parse_float(request.form.get('insulin')),
            hba1c=parse_float(request.form.get('hba1c')),
            triglycerides=parse_float(request.form.get('triglycerides')),
            cholesterol_tot=parse_float(request.form.get('cholesterol_tot')),
            hdl=parse_float(request.form.get('hdl')),
            ldl=parse_float(request.form.get('ldl')),
            vldl=parse_float(request.form.get('vldl')),
            uric_acid=parse_float(request.form.get('uric_acid')),
            got=parse_float(request.form.get('got')),
            gpt=parse_float(request.form.get('gpt')),
            fa=parse_float(request.form.get('fa')),
            bili_tot=parse_float(request.form.get('bili_tot')),
            bili_dir=parse_float(request.form.get('bili_dir')),
            bili_ind=parse_float(request.form.get('bili_ind')),
            urea=parse_float(request.form.get('urea')),
            creatinine=parse_float(request.form.get('creatinine')),
            tsh=parse_float(request.form.get('tsh')),
            t4=parse_float(request.form.get('t4')),
            t3=parse_float(request.form.get('t3')),
            anti_tpo=parse_float(request.form.get('anti_tpo')),
            vit_d=parse_float(request.form.get('vit_d')),
            vit_b12=parse_float(request.form.get('vit_b12')),
            iron=parse_float(request.form.get('iron')),
            ferritin=parse_float(request.form.get('ferritin')),
            calcium=parse_float(request.form.get('calcium')),
            magnesium=parse_float(request.form.get('magnesium')),
            phosphorus=parse_float(request.form.get('phosphorus')),
            chlorine=parse_float(request.form.get('chlorine')),
            platelets=parse_float(request.form.get('platelets')),
            otros_laboratorios=format_text(request.form.get('otros_laboratorios')),
            ecografia_abdominal=request.form.get('ecografia_abdominal'),
            ecografia_tiroidea=request.form.get('ecografia_tiroidea'),
            otros_estudios=request.form.get('otros_estudios'),
            indicaciones_tratamiento=format_text(request.form.get('indicaciones_tratamiento')),
            medications=format_text(request.form.get('medications')),
            proximo_control=request.form.get('proximo_control'),
            notes=format_text(request.form.get('notes')),
            conducta_medica=format_text(request.form.get('conducta_medica'))
        )
        db.session.add(record)
        db.session.commit()
        
        recalculate_deltas(patient.id)

        if 'photo' in request.files:
            file = request.files['photo']
            if file and allowed_file(file.filename):
                filename = secure_filename(f"p_{patient.id}_r_{record.id}_{file.filename}")
                file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
                photo = PatientPhoto(patient_id=patient.id, record_id=record.id, filename=filename, category=request.form.get('photo_category', 'Evolución'), date=request.form['date'], notes=request.form.get('photo_notes', ''))
                db.session.add(photo)
                db.session.commit()

        return redirect(url_for('patient_detail', id=patient.id))
    return render_template('new_record.html', patient=patient)

@app.route('/record/edit/<int:id>', methods=['GET', 'POST'])
@login_required
def edit_record(id):
    record = MedicalRecord.query.get_or_404(id)
    patient = get_patient_for_current_user(record.patient_id)
    if not patient:
        flash("No tiene permisos para editar este registro.", "warning")
        return redirect(url_for('index'))
    if request.method == 'POST':
        record.date = request.form['date']
        record.lab_date = request.form.get('lab_date')
        record.marca_balanza = request.form.get('marca_balanza')
        record.pa = request.form.get('pa')
        record.fc = parse_int(request.form.get('fc'))
        record.height = parse_float(request.form.get('height'))
        record.weight = parse_float(request.form.get('weight'))
        record.body_fat_kg = parse_float(request.form.get('body_fat_kg') or request.form.get('masa_grasa_kg'))
        record.body_fat_pct = parse_float(request.form.get('body_fat_pct') or request.form.get('porcentaje_grasa_corporal'))
        record.visceral_fat = parse_float(request.form.get('visceral_fat') or request.form.get('grasa_visceral'))
        record.muscle_mass_kg = parse_float(request.form.get('muscle_mass_kg') or request.form.get('masa_muscular_esqueletica'))
        record.muscle_mass_pct = parse_float(request.form.get('muscle_mass_pct') or request.form.get('porcentaje_musculo_esqueletico'))
        record.basal_metabolism = parse_float(request.form.get('basal_metabolism') or request.form.get('tasa_metabolica_basal'))
        record.waist = parse_float(request.form.get('waist'))
        record.masa_grasa_kg = parse_float(request.form.get('masa_grasa_kg'))
        record.porcentaje_musculo = parse_float(request.form.get('porcentaje_musculo'))
        record.porcentaje_musculo_esqueletico = parse_float(request.form.get('porcentaje_musculo_esqueletico'))
        record.masa_muscular_esqueletica = parse_float(request.form.get('masa_muscular_esqueletica'))
        record.porcentaje_oseo = parse_float(request.form.get('porcentaje_oseo'))
        record.masa_osea = parse_float(request.form.get('masa_osea'))
        record.porcentaje_proteina = parse_float(request.form.get('porcentaje_proteina'))
        record.masa_proteica = parse_float(request.form.get('masa_proteica'))
        record.porcentaje_agua_corporal = parse_float(request.form.get('porcentaje_agua_corporal'))
        record.masa_agua_corporal = parse_float(request.form.get('masa_agua_corporal'))
        record.peso_sin_grasa = parse_float(request.form.get('peso_sin_grasa'))
        record.grasa_subcutanea_pct = parse_float(request.form.get('grasa_subcutanea_pct'))
        record.edad_metabolica = parse_int(request.form.get('edad_metabolica'))
        record.hemoglobin = parse_float(request.form.get('hemoglobin'))
        record.glucose = parse_float(request.form.get('glucose'))
        record.insulin = parse_float(request.form.get('insulin'))
        record.hba1c = parse_float(request.form.get('hba1c'))
        record.triglycerides = parse_float(request.form.get('triglycerides'))
        record.cholesterol_tot = parse_float(request.form.get('cholesterol_tot'))
        record.hdl = parse_float(request.form.get('hdl'))
        record.ldl = parse_float(request.form.get('ldl'))
        record.vldl = parse_float(request.form.get('vldl'))
        record.uric_acid = parse_float(request.form.get('uric_acid'))
        record.got = parse_float(request.form.get('got'))
        record.gpt = parse_float(request.form.get('gpt'))
        record.fa = parse_float(request.form.get('fa'))
        record.bili_tot = parse_float(request.form.get('bili_tot'))
        record.bili_dir = parse_float(request.form.get('bili_dir'))
        record.bili_ind = parse_float(request.form.get('bili_ind'))
        record.urea = parse_float(request.form.get('urea'))
        record.creatinine = parse_float(request.form.get('creatinine'))
        record.tsh = parse_float(request.form.get('tsh'))
        record.t4 = parse_float(request.form.get('t4'))
        record.t3 = parse_float(request.form.get('t3'))
        record.anti_tpo = parse_float(request.form.get('anti_tpo'))
        record.vit_d = parse_float(request.form.get('vit_d'))
        record.vit_b12 = parse_float(request.form.get('vit_b12'))
        record.iron = parse_float(request.form.get('iron'))
        record.ferritin = parse_float(request.form.get('ferritin'))
        record.calcium = parse_float(request.form.get('calcium'))
        record.magnesium = parse_float(request.form.get('magnesium'))
        record.phosphorus = parse_float(request.form.get('phosphorus'))
        record.chlorine = parse_float(request.form.get('chlorine'))
        record.platelets = parse_float(request.form.get('platelets'))
        record.otros_laboratorios = format_text(request.form.get('otros_laboratorios'))
        record.ecografia_abdominal = request.form.get('ecografia_abdominal')
        record.ecografia_tiroidea = request.form.get('ecografia_tiroidea')
        record.otros_estudios = request.form.get('otros_estudios')
        record.indicaciones_tratamiento = format_text(request.form.get('indicaciones_tratamiento'))
        record.medications = format_text(request.form.get('medications'))
        record.proximo_control = request.form.get('proximo_control')
        record.notes = request.form.get('notes')
        record.conducta_medica = format_text(request.form.get('conducta_medica'))
        
        db.session.commit()
        recalculate_deltas(patient.id)
        
        return redirect(url_for('patient_detail', id=patient.id))
    return render_template('edit_record.html', patient=patient, record=record)

@app.route('/record/delete/<int:id>', methods=['POST'])
@login_required
def delete_record(id):
    record = MedicalRecord.query.get_or_404(id)
    patient = get_patient_for_current_user(record.patient_id)
    if not patient:
        flash("No tiene permisos para eliminar este registro.", "warning")
        return redirect(url_for('index'))
    patient_id = record.patient_id
    db.session.delete(record)
    db.session.commit()
    recalculate_deltas(patient_id)
    return redirect(url_for('patient_detail', id=patient_id))

@app.route('/treatment/edit/<int:patient_id>', methods=['GET', 'POST'])
@login_required
def edit_treatment(patient_id):
    patient = get_patient_for_current_user(patient_id)
    if not patient:
        flash("No tiene permisos para acceder a este paciente.", "warning")
        return redirect(url_for('index'))
    last_record = MedicalRecord.query.filter_by(patient_id=patient.id).order_by(MedicalRecord.id.desc()).first()
    
    if request.method == 'POST':
        today_str = date.today().strftime('%Y-%m-%d')
        if not last_record:
            last_record = MedicalRecord(patient_id=patient.id, date=today_str, weight=None, height=None)
            db.session.add(last_record)
            db.session.commit()
            
        last_record.indicaciones_tratamiento = format_text(request.form.get('indicaciones_tratamiento'))
        last_record.medications = format_text(request.form.get('medications'))
        last_record.proximo_control = request.form.get('proximo_control')
        last_record.conducta_medica = format_text(request.form.get('conducta_medica'))
        db.session.commit()

        if 'photo' in request.files:
            file = request.files['photo']
            if file and allowed_file(file.filename):
                filename = secure_filename(f"p_{patient.id}_r_{last_record.id}_{file.filename}")
                file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
                
                photo = PatientPhoto(
                    patient_id=patient.id,
                    record_id=last_record.id,
                    filename=filename,
                    category=request.form.get('photo_category', 'Evolución'),
                    date=last_record.date,
                    notes=request.form.get('photo_notes', '')
                )
                db.session.add(photo)
                db.session.commit()

        return redirect(url_for('patient_detail', id=patient.id))
        
    return render_template('edit_treatment.html', patient=patient, record=last_record)

@app.route('/export/excel')
@login_required
def export_excel():
    current_user = get_current_user()
    output = io.StringIO()
    writer = csv.writer(output)
    headers_csv = [
        'ID', 'Nombre', 'CI', 'Fecha Nacimiento', 'Edad', 'Género', 'Teléfono', 'Sanatorio', 'Procedencia', 'Profesión', 'Patología', 
        'Seguro Médico', 'Inicio Tratamiento', 'Años Evolución', 'Antecedentes Familiares', 'Antecedentes Personales', 'Comorbilidades', 'Tratamiento Base',
        'Fecha Control', 'Fecha Lab', 'Balanza', 'PA', 'FC', 'Altura (cm)', 'Peso (kg)', 'Peso Perdido (kg)', 'IMC', 'Grasa Corp (kg)', 'Grasa Perdida (kg)', 'Grasa Corp (%)', 'Grasa Visceral', 
        'Músculo (kg)', 'Músculo Delta (kg)', 'Músculo (%)', 'Metabolismo Basal', 'Masa Grasa (kg)', '% Músculo', '% Músculo Esquelético', 'Masa Músculo Esquelético (kg)', '% Óseo', 'Masa Ósea (kg)', '% Proteína', 'Masa Proteica (kg)', '% Agua Corporal', 'Masa Agua Corporal (kg)', 'Peso Sin Grasa (kg)', 'Grasa Subcutánea (%)', 'Edad Metabólica', 'Cintura (cm)', 'Cintura Perdida (cm)', 'Hemoglobina', 'Glicemia', 'Insulina', 'HbA1c (%)', 
        'Triglicéridos', 'Col. Tot', 'HDL', 'LDL', 'VLDL', 'Tg/HDL', 'Ácido Úrico', 'GOT', 'GPT', 'FA', 'Bili Tot', 'Bili Dir', 
        'Bili Ind', 'Urea', 'Creatinina', 'TSH', 'T4', 'T3', 'Anti-TPO', 'Vit D', 'Vit B12', 'Hierro', 'Ferritina', 'Calcio', 'Magnesio', 
        'Fósforo', 'Cloro', 'Plaquetas', 'Otros Laboratorios', 'Eco Abdominal', 'Eco Tiroidea', 'Otros Estudios', 'Evolución / Notas', 'Indicaciones / Tratamiento', 'Medicación', 
        'Próximo Control', 'HOMA-IR', 'FIB-4', 'Cintura/Altura', 'Conducta Médica'
    ]
    writer.writerow(headers_csv)
    if current_user and current_user.username == 'admin':
        patients = Patient.query.all()
    elif current_user:
        patients = Patient.query.filter_by(user_id=current_user.id).all()
    else:
        patients = []
    for p in patients:
        valid_recs = sorted([r for r in p.records if r.weight is not None and r.weight > 0], key=lambda x: x.date)
        if valid_recs:
            for r in valid_recs:
                writer.writerow([
                    p.id, p.full_name, p.ci, p.birth_date, p.age, p.gender, p.phone, p.sanatorio, p.procedencia, p.profession, p.pathology_type,
                    p.seguro_medico, p.fecha_inicio_tratamiento, p.evolution_years, p.antecedentes_familiares, p.antecedentes_personales, p.comorbidities, p.tratamiento_base,
                    r.date, r.lab_date, r.marca_balanza, r.pa, r.fc, r.height, r.weight, r.weight_lost, r.bmi, r.body_fat_kg, r.fat_lost, r.body_fat_pct, r.visceral_fat, r.muscle_mass_kg, r.muscle_delta, r.muscle_mass_pct,
                    r.basal_metabolism, r.masa_grasa_kg, r.porcentaje_musculo, r.porcentaje_musculo_esqueletico, r.masa_muscular_esqueletica, r.porcentaje_oseo, r.masa_osea, r.porcentaje_proteina, r.masa_proteica, r.porcentaje_agua_corporal, r.masa_agua_corporal, r.peso_sin_grasa, r.grasa_subcutanea_pct, r.edad_metabolica, r.waist, r.waist_lost, r.hemoglobin, r.glucose, r.insulin, r.hba1c, r.triglycerides, r.cholesterol_tot, r.hdl, r.ldl,
                    r.vldl, r.tg_hdl, r.uric_acid, r.got, r.gpt, r.fa, r.bili_tot, r.bili_dir, r.bili_ind, r.urea, r.creatinine,
                    r.tsh, r.t4, r.t3, r.anti_tpo, r.vit_d, r.vit_b12, r.iron, r.ferritin, r.calcium, r.magnesium, r.phosphorus, r.chlorine, r.platelets,
                    r.otros_laboratorios, r.ecografia_abdominal, r.ecografia_tiroidea, r.otros_estudios, r.notes, r.indicaciones_tratamiento, r.medications,
                    r.proximo_control, r.homair, r.fib4, r.ica_value, r.conducta_medica
                ])
        else:
            empty_row = [
                p.id, p.full_name, p.ci, p.birth_date, p.age, p.gender, p.phone, p.sanatorio, p.procedencia, p.profession, p.pathology_type,
                p.seguro_medico, p.fecha_inicio_tratamiento, p.evolution_years, p.antecedentes_familiares, p.antecedentes_personales, p.comorbidities, p.tratamiento_base
            ] + [''] * (len(headers_csv) - 18)
            writer.writerow(empty_row)
    output.seek(0)
    return Response(output.getvalue(), mimetype="text/csv", headers={"Content-Disposition": "attachment;filename=matriz_metabolica_pacientes.csv"})

@app.route('/report/pdf/<int:patient_id>')
@login_required
def generate_pdf(patient_id):
    patient = get_patient_for_current_user(patient_id)
    if not patient:
        flash("No tiene permisos para generar el informe de este paciente.", "warning")
        return redirect(url_for('index'))
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
    story = []
    
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('DocTitle', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=14, alignment=1, spaceAfter=15, textColor=colors.HexColor('#003366'))
    section_style = ParagraphStyle('SectionHeading', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10, alignment=1, spaceAfter=8, spaceBefore=10, textColor=colors.HexColor('#003366'))
    body_style = ParagraphStyle('DocBody', parent=styles['Normal'], fontName='Helvetica', fontSize=9, leading=13, textColor=colors.HexColor('#003366'))
    contact_style = ParagraphStyle('ContactInfo', parent=styles['Normal'], fontName='Helvetica', fontSize=8, alignment=1, textColor=colors.HexColor('#555555'), spaceAfter=15)

    nombre_profesional = session.get('user_name') or 'Dra. Alejandra Müller'
    encabezado_sistema = f"<b>Plataforma Metabolisys</b> | {nombre_profesional}"
    
    logo_path = os.path.join(app.config['UPLOAD_FOLDER'].replace('/var/data/uploads', 'static'), 'logo.png')
    if not os.path.exists(logo_path): logo_path = os.path.join('static', 'logo.png')
        
    if os.path.exists(logo_path):
        story.append(RLImage(logo_path, width=200, height=70))
        story.append(Spacer(1, 5))
        story.append(Paragraph(encabezado_sistema, contact_style))
        story.append(Spacer(1, 10))
    else:
        story.append(Paragraph(encabezado_sistema, contact_style))
        story.append(Spacer(1, 15))
    
    valid_records = sorted([r for r in patient.records if r.weight is not None and r.weight > 0], key=lambda x: x.date)
    last_r = valid_records[-1] if valid_records else None
    report_date = last_r.date if last_r else "N/D"
    
    story.append(Paragraph(f"{patient.full_name.upper()} &nbsp;&nbsp;&nbsp;&nbsp; {report_date}", title_style))
    story.append(Spacer(1, 5))
    
    if valid_records:
        table_data = [['FECHA', 'BAL.', 'PA / FC', 'ALT.\n(cm)', 'PESO\n(kg)', 'GRASA\n(kg)', 'GRASA\n(%)', 'VISC.', 'MÚSC.\n(kg)', 'MÚSC.\n(%)', 'METAB.\n(kcal)', 'IMC', 'CA\n(cm)']]
        for r in valid_records:
            pa_fc = f"{r.pa or '-'} / {r.fc or '-'}"
            table_data.append([str(r.date), str(r.marca_balanza or '-'), pa_fc, str(r.height or '-'), str(r.weight or '-'), str(r.body_fat_kg or '-'), str(r.body_fat_pct or '-'), str(r.visceral_fat or '-'), str(r.muscle_mass_kg or '-'), str(r.muscle_mass_pct or '-'), str(r.basal_metabolism or '-'), str(r.bmi or '-'), str(r.waist or '-')])
        
        t = Table(table_data, colWidths=[55, 45, 50, 35, 35, 38, 35, 35, 38, 35, 45, 30, 30])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#e8f4f4')), ('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'), ('FONTSIZE', (0,0), (-1,0), 6.5), ('FONTSIZE', (0,1), (-1,-1), 6.5),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#cccccc')), ('BOTTOMPADDING', (0,0), (-1,-1), 4), ('TOPPADDING', (0,0), (-1,-1), 4),
            ('TEXTCOLOR', (0,0), (-1,-1), colors.HexColor('#003366'))
        ]))
        story.append(t)
        story.append(Spacer(1, 10))
        
        story.append(Paragraph("RESUMEN EVOLUTIVO Y COMPARATIVO", section_style))
        
        r_init = valid_records[0]
        r_prev = valid_records[-2] if len(valid_records) > 1 else r_init
        r_curr = valid_records[-1]
        
        def safe_val(rec, attr):
            val = getattr(rec, attr, None)
            return float(val) if val is not None else 0.0

        params = [
            ("Peso", "weight", "kg", 1),
            ("Grasa", "body_fat_pct", "%", 1),
            ("Grasa (kg)", "body_fat_kg", "kg", 1),
            ("Músculo (kg)", "muscle_mass_kg", "kg", 1),
            ("Cintura", "waist", "cm", 1),
            ("Metabolismo", "basal_metabolism", "kcal", 0)
        ]
        
        summary_data = [['PARÁMETRO', 'INICIAL', 'ANTERIOR', 'ACTUAL', 'DIF. ANTERIOR', 'DIF. TOTAL']]
        
        for name, attr, unit, dec in params:
            v_init = safe_val(r_init, attr)
            v_prev = safe_val(r_prev, attr)
            v_curr = safe_val(r_curr, attr)
            
            diff_prev = round(v_curr - v_prev, dec) if v_prev else 0
            diff_total = round(v_curr - v_init, dec) if v_init else 0
            
            f_init = f"{v_init:.{dec}f} {unit}" if v_init else "-"
            f_prev = f"{v_prev:.{dec}f} {unit}" if v_prev else "-"
            f_curr = f"{v_curr:.{dec}f} {unit}" if v_curr else "-"
            
            def format_col_diff(val):
                if val == 0: return f"0.0 {unit}"
                str_val = f"+{val}" if val > 0 else f"{val}"
                return f"{str_val} {unit}"

            summary_data.append([name, f_init, f_prev, f_curr, format_col_diff(diff_prev), format_col_diff(diff_total)])
            
        t_summary = Table(summary_data, colWidths=[90, 75, 75, 75, 85, 85])
        t_summary.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#003366')),
            ('TEXTCOLOR', (0,0), (-1,0), colors.white),
            ('TEXTCOLOR', (0,1), (-1,-1), colors.HexColor('#003366')),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,0), 8),
            ('FONTSIZE', (0,1), (-1,-1), 8),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#cccccc')),
            ('BOTTOMPADDING', (0,0), (-1,-1), 3),
            ('TOPPADDING', (0,0), (-1,-1), 3),
        ]))
        story.append(t_summary)
        story.append(Spacer(1, 10))
        
        story.append(Paragraph("<b>INDICACIONES Y PLAN TERAPÉUTICO</b>", section_style))
        story.append(Spacer(1, 5))
        
        meds, inds, plan_text = last_r.medications or "", last_r.indicaciones_tratamiento or "", ""
        if meds:
            for line in meds.split('\n'):
                if line.strip(): plan_text += f"• &nbsp; <b>{line.upper()}</b><br/>"
        if inds:
            for line in inds.split('\n'):
                if line.strip(): plan_text += f"• &nbsp; {line.upper()}<br/>"
        if last_r.proximo_control: plan_text += f"<br/>• &nbsp; <b>CONTROL EN {last_r.proximo_control.upper()}.</b>"
        if not plan_text: plan_text = "Sin indicaciones registradas para este control."
            
        story.append(Paragraph(plan_text, body_style))
    else:
        story.append(Paragraph("No hay controles antropométricos registrados todavía para este paciente.", body_style))
        
    doc.build(story)
    buffer.seek(0)
    return send_file(buffer, as_attachment=True, download_name=f"Informe_{patient.full_name}.pdf", mimetype='application/pdf')

@app.route('/uploads/<filename>')
def uploaded_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
