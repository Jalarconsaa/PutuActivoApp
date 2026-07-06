"""
PUTÚ ACTIVO — Sistema de Gestión Integral v5.0
Base de datos SQLite unificada · Exportación/Importación Excel
"""
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, date, timedelta
import warnings, os, hashlib, io, base64, time, sqlite3, re, math, urllib.request
import qrcode
try:
    from reportlab.lib.pagesizes import A4, landscape, letter
    from reportlab.lib.units import cm
    from reportlab.lib import colors as rl_colors
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.enums import TA_CENTER
    from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
        Paragraph, Spacer, Image as RLImage, KeepTogether)
    from reportlab.pdfbase.pdfmetrics import stringWidth
    REPORTLAB_OK = True
except ImportError:
    REPORTLAB_OK = False
try:
    import cv2
    import numpy as np
    QR_SCAN_DISPONIBLE = True
except ImportError:
    QR_SCAN_DISPONIBLE = False

warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────────────────────────────────────
# RUTAS
# ─────────────────────────────────────────────────────────────────────────────
BASE_DIR  = os.path.dirname(os.path.abspath(__file__))
DB_PATH   = os.path.join(BASE_DIR, "putu_activo.db")
FOTOS_DIR = os.path.join(BASE_DIR, "fotos_clientes")
RUTS_DIR  = os.path.join(BASE_DIR, "rutinas_pdf")
LOGO_PATH = next((os.path.join(BASE_DIR,f) for f in os.listdir(BASE_DIR)
    if f.lower().startswith("logo") and f.lower().endswith((".png",".jpg",".jpeg",".svg"))), None)
os.makedirs(FOTOS_DIR, exist_ok=True)
os.makedirs(RUTS_DIR,  exist_ok=True)

# ─────────────────────────────────────────────────────────────────────────────
# COLORES
# ─────────────────────────────────────────────────────────────────────────────
VERDE="#6DBE45"; VERDE_DK="#4A8A2A"; NEGRO="#0D0D0D"
GRIS="#141414";  GRIS2="#1E1E1E";   GRIS3="#2E2E2E"; GRIS4="#3E3E3E"
BLANCO="#F2F2F2"; GRIS_T="#CCCCCC"
ROJO="#E85050";  NARANJA="#E8A838"; AZUL="#3A9BD5"

# ─────────────────────────────────────────────────────────────────────────────
# USUARIOS
# ─────────────────────────────────────────────────────────────────────────────
def _h(pw): return hashlib.sha256(pw.encode()).hexdigest()
USUARIOS = {
    "admin":      {"hash":_h(""),"rol":"Administrador","nombre":"Administrador",
                   "permisos":["dashboard","clientes","nuevo","pagos","asistencia","clases",
                               "egresos","reportes","db","eval_editar","cliente_eliminar",
                               "asist_eliminar","config_usuarios"]},
    "entrenador": {"hash":_h("gym123"),  "rol":"Entrenador",   "nombre":"Entrenador",
                   "permisos":["dashboard","clientes","asistencia","eval_editar","reportes"]},
    "asistente":  {"hash":_h("asist1"),  "rol":"Asistente",    "nombre":"Asistente",
                   "permisos":["dashboard","clientes","nuevo","pagos","asistencia","clases","reportes"]},
}
SESSION_H = 5

def tiene_permiso(permiso):
    """Verifica si el usuario actual tiene un permiso específico."""
    u = st.session_state.get("usuario","")
    return permiso in USUARIOS.get(u,{}).get("permisos",[])

PERIODOS = ["Diario","Quincenal","Mensual","Bimensual","Trimestral","Semestral","Anual"]
_PERIODOS_ALIAS = {"Menstrual":"Mensual","menstrual":"Mensual"}
PERIODO_DIAS = {"Diario":1,"Quincenal":15,"Mensual":30,"Bimensual":60,
                "Trimestral":90,"Semestral":180,"Anual":365}
PLANES    = ["PM","AM","Estudiante","Adulto Mayor","Pase diario"]
FRECUENCIAS = ["Pase diario","2 x Sem.","3 x Sem.","Full Sem."]
HORARIOS  = ["08:30 a 11:30","11:30 a 13:30","15:30 a 18:00","18:00 a 20:00","20:00 a 22:00"]
OBJETIVOS = ["Perder peso","Ganar masa muscular","Aumentar la energía",
             "Mejorar la flexibilidad","Mejorar la resistencia","Aumentar la fuerza","Reducir el estrés"]
NIVELES   = ["Principiante","Intermedio","Avanzado"]
MESES_ESP = ["enero","febrero","marzo","abril","mayo","junio",
             "julio","agosto","septiembre","octubre","noviembre","diciembre"]

# ─────────────────────────────────────────────────────────────────────────────
# DB
# ─────────────────────────────────────────────────────────────────────────────
def get_conn():
    return sqlite3.connect(DB_PATH, check_same_thread=False)

def init_db():
    c = get_conn(); cur = c.cursor()
    cur.executescript("""
    CREATE TABLE IF NOT EXISTS clientes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nombre TEXT, rut TEXT UNIQUE,
        fecha_nacimiento TEXT, edad INTEGER, mes_cumpleanos TEXT,
        sexo TEXT, direccion TEXT, celular TEXT, email TEXT,
        contacto_emergencia TEXT, celular_emergencia TEXT, parentesco TEXT,
        fecha_inscripcion TEXT, periodo_vencimiento TEXT,
        tipo_plan TEXT, frecuencia TEXT, horario TEXT,
        lunes TEXT, martes TEXT, miercoles TEXT, jueves TEXT, viernes TEXT, sabado TEXT,
        valor_plan REAL, fecha_vencimiento TEXT, fecha_renovacion TEXT,
        estado TEXT, rutina TEXT, nivel TEXT,
        enfermedad TEXT, restricciones TEXT, objetivo TEXT, talla TEXT,
        foto_path TEXT, rutina_pdf_path TEXT,
        mensaje_cumpleanos TEXT, mensaje_vencimiento TEXT, mensaje_renovacion TEXT,
        creado TEXT, modificado TEXT
    );
    CREATE TABLE IF NOT EXISTS pagos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        rut TEXT, nombre TEXT, fecha TEXT, monto REAL,
        concepto TEXT, tipo_plan TEXT, frecuencia TEXT,
        medio_pago TEXT, observacion TEXT, usuario TEXT
    );
    CREATE TABLE IF NOT EXISTS asistencia (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        rut TEXT, nombre TEXT, fecha TEXT, hora TEXT, usuario TEXT
    );
    CREATE TABLE IF NOT EXISTS evaluaciones (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        rut TEXT, nombre TEXT, fecha TEXT,
        estatura REAL, peso REAL, imc REAL, grasa_pct REAL,
        masa_musc REAL, agua_pct REAL, grasa_visceral REAL,
        metabolismo REAL, proteina_pct REAL,
        brazos REAL, abdomen REAL, cadera REAL,
        observaciones TEXT, usuario TEXT
    );
    CREATE TABLE IF NOT EXISTS clases (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        tipo TEXT, titulo TEXT, fecha TEXT, hora TEXT,
        participante TEXT, monto REAL, observacion TEXT, usuario TEXT
    );
    CREATE TABLE IF NOT EXISTS egresos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        fecha TEXT, categoria TEXT, monto REAL, descripcion TEXT, usuario TEXT
    );
    CREATE TABLE IF NOT EXISTS productos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        fecha TEXT, producto TEXT, monto REAL, usuario TEXT
    );
    CREATE TABLE IF NOT EXISTS ejercicios (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nombre TEXT NOT NULL,
        url_imagen TEXT, descripcion TEXT, ejecucion TEXT,
        musculo_primario TEXT, musculo_secundario TEXT, url_body TEXT,
        casa INTEGER NOT NULL DEFAULT 0,
        gimnasio INTEGER NOT NULL DEFAULT 0,
        estiramiento INTEGER NOT NULL DEFAULT 0,
        rehabilitacion INTEGER NOT NULL DEFAULT 0,
        video TEXT, nivel TEXT
    );
    CREATE TABLE IF NOT EXISTS rutinas (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        cliente_rut TEXT,
        nombre TEXT, activa INTEGER NOT NULL DEFAULT 1,
        fecha_vencimiento TEXT,
        fecha_creacion TEXT DEFAULT CURRENT_TIMESTAMP
    );
    CREATE TABLE IF NOT EXISTS rutina_ejercicios (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        rutina_id INTEGER NOT NULL,
        ejercicio_id INTEGER NOT NULL,
        dia_semana TEXT NOT NULL,
        orden INTEGER NOT NULL DEFAULT 0,
        metodo TEXT DEFAULT 'Normal',
        series TEXT, repeticiones TEXT, peso TEXT,
        tempo_descanso TEXT, notas TEXT,
        FOREIGN KEY (rutina_id) REFERENCES rutinas(id) ON DELETE CASCADE,
        FOREIGN KEY (ejercicio_id) REFERENCES ejercicios(id) ON DELETE CASCADE
    );
    CREATE TABLE IF NOT EXISTS alimentos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nombre TEXT NOT NULL,
        cantidad_ref TEXT, unidad TEXT, peso_neto_g REAL DEFAULT 100,
        kcal REAL DEFAULT 0, proteina_g REAL DEFAULT 0,
        lipidos_g REAL DEFAULT 0, hdc_g REAL DEFAULT 0, fibra_g REAL DEFAULT 0,
        grupo TEXT, ig REAL, sodio_mg REAL, calcio_mg REAL
    );
    CREATE UNIQUE INDEX IF NOT EXISTS idx_alimentos_nombre ON alimentos(nombre);
    CREATE TABLE IF NOT EXISTS usuarios_sistema (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        usuario TEXT UNIQUE NOT NULL,
        nombre TEXT NOT NULL,
        password_hash TEXT NOT NULL,
        rol TEXT NOT NULL DEFAULT 'entrenador',
        activo INTEGER DEFAULT 1,
        creado TEXT DEFAULT CURRENT_TIMESTAMP
    );
    CREATE TABLE IF NOT EXISTS activity_log (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        timestamp TEXT DEFAULT CURRENT_TIMESTAMP,
        usuario TEXT,
        rol TEXT,
        accion TEXT,
        detalle TEXT,
        rut_afectado TEXT
    );
    CREATE TABLE IF NOT EXISTS usuarios_clientes (
        rut TEXT PRIMARY KEY,
        email TEXT UNIQUE NOT NULL,
        password_hash TEXT NOT NULL,
        activo INTEGER DEFAULT 1,
        ultimo_acceso TEXT,
        creado TEXT DEFAULT CURRENT_TIMESTAMP
    );
    CREATE TABLE IF NOT EXISTS planes_nutri (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        cliente_rut TEXT NOT NULL,
        nombre TEXT NOT NULL,
        profesional TEXT,
        objetivo TEXT,
        kcal_objetivo REAL,
        activo INTEGER DEFAULT 1,
        notas TEXT,
        fecha_creacion TEXT DEFAULT CURRENT_TIMESTAMP
    );
    CREATE TABLE IF NOT EXISTS comidas_plan (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        plan_id INTEGER NOT NULL,
        dia_semana TEXT NOT NULL,
        tipo_comida TEXT NOT NULL,
        orden INTEGER DEFAULT 0,
        FOREIGN KEY (plan_id) REFERENCES planes_nutri(id) ON DELETE CASCADE
    );
    CREATE TABLE IF NOT EXISTS comida_alimentos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        comida_id INTEGER NOT NULL,
        alimento_id INTEGER NOT NULL,
        cantidad_g REAL DEFAULT 100,
        kcal_calc REAL, prot_calc REAL, lip_calc REAL, hdc_calc REAL, fibra_calc REAL,
        notas TEXT,
        FOREIGN KEY (comida_id) REFERENCES comidas_plan(id) ON DELETE CASCADE,
        FOREIGN KEY (alimento_id) REFERENCES alimentos(id) ON DELETE CASCADE
    );
    """)
    c.commit()
    # Migración: agregar columnas nuevas si la DB ya existía sin ellas
    migraciones = [
        "ALTER TABLE clientes ADD COLUMN rutina_pdf_path TEXT",
        "ALTER TABLE clientes ADD COLUMN foto_path TEXT",
        "ALTER TABLE clientes ADD COLUMN periodo_vencimiento TEXT",
        "ALTER TABLE clientes ADD COLUMN fecha_renovacion TEXT",
        "ALTER TABLE clientes ADD COLUMN lunes TEXT",
        "ALTER TABLE clientes ADD COLUMN martes TEXT",
        "ALTER TABLE clientes ADD COLUMN miercoles TEXT",
        "ALTER TABLE clientes ADD COLUMN jueves TEXT",
        "ALTER TABLE clientes ADD COLUMN viernes TEXT",
        "ALTER TABLE clientes ADD COLUMN sabado TEXT",
        "ALTER TABLE clientes ADD COLUMN mensaje_cumpleanos TEXT",
        "ALTER TABLE clientes ADD COLUMN mensaje_vencimiento TEXT",
        "ALTER TABLE clientes ADD COLUMN mensaje_renovacion TEXT",
        "ALTER TABLE clientes ADD COLUMN mes_cumpleanos TEXT",
        "ALTER TABLE clientes ADD COLUMN talla TEXT",
        "ALTER TABLE clientes ADD COLUMN parentesco TEXT",
        "ALTER TABLE clientes ADD COLUMN celular_emergencia TEXT",
        "ALTER TABLE clientes ADD COLUMN contacto_emergencia TEXT",
        "ALTER TABLE clientes ADD COLUMN nivel TEXT",
        "ALTER TABLE clientes ADD COLUMN objetivo TEXT",
        "ALTER TABLE clientes ADD COLUMN restricciones TEXT",
        "ALTER TABLE clientes ADD COLUMN enfermedad TEXT",
        "ALTER TABLE clientes ADD COLUMN rutina TEXT",
        "ALTER TABLE clientes ADD COLUMN creado TEXT",
        "ALTER TABLE clientes ADD COLUMN modificado TEXT",
        "ALTER TABLE clases ADD COLUMN titulo TEXT",
        "ALTER TABLE asistencia ADD COLUMN tipo TEXT DEFAULT 'ingreso'",
        "ALTER TABLE asistencia ADD COLUMN hora_salida TEXT",
    ]
    for m in migraciones:
        try: c.execute(m)
        except: pass  # columna ya existe
    c.commit(); c.close()

init_db()

# Seed ejercicios si la tabla está vacía
def _seed_ejercicios():
    """Carga ejercicios desde CSV solo si no existen por nombre (evita duplicados)."""
    try:
        _seed_path = os.path.join(BASE_DIR,"ejercicios_seed.csv")
        if not os.path.exists(_seed_path): return
        _conn_s = get_conn()
        # Agregar índice único por nombre si no existe (previene duplicados futuros)
        try:
            _conn_s.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_ejercicios_nombre ON ejercicios(nombre)")
            _conn_s.commit()
        except: pass
        import csv
        with open(_seed_path,encoding="utf-8") as _f:
            _reader = csv.DictReader(_f)
            _rows = list(_reader)
        _bool_cols = ["casa","gimnasio","estiramiento","rehabilitacion"]
        for _r in _rows:
            for _bc in _bool_cols:
                _r[_bc] = 1 if str(_r.get(_bc,"")).strip().lower() in ["1","true","si","sí","yes"] else 0
            for _nc in ["url_imagen","descripcion","ejecucion","musculo_primario",
                        "musculo_secundario","url_body","video","nivel"]:
                if _nc not in _r: _r[_nc] = None
        # INSERT OR IGNORE: no falla si el ejercicio ya existe
        _conn_s.executemany("""INSERT OR IGNORE INTO ejercicios
            (nombre,url_imagen,descripcion,ejecucion,musculo_primario,
             musculo_secundario,url_body,casa,gimnasio,estiramiento,rehabilitacion,video,nivel)
            VALUES (:nombre,:url_imagen,:descripcion,:ejecucion,:musculo_primario,
            :musculo_secundario,:url_body,:casa,:gimnasio,:estiramiento,:rehabilitacion,:video,:nivel)""",
            _rows)
        _conn_s.commit(); _conn_s.close()
    except Exception as _e: pass

def _dedup_ejercicios():
    """Elimina duplicados en ejercicios (mismo nombre), deja solo el de menor id."""
    try:
        _c=get_conn()
        _c.execute("""DELETE FROM ejercicios WHERE id NOT IN (
            SELECT MIN(id) FROM ejercicios GROUP BY nombre)""")
        _c.commit(); _c.close()
    except: pass
_seed_ejercicios()
_dedup_ejercicios()

# Seed alimentos desde CSV si la tabla está vacía
def _seed_alimentos():
    try:
        _seed_path = os.path.join(BASE_DIR,"alimentos_seed.csv")
        if not os.path.exists(_seed_path): return
        _conn_a = get_conn()
        if _conn_a.execute("SELECT COUNT(*) FROM alimentos").fetchone()[0] > 0:
            _conn_a.close(); return
        import csv
        with open(_seed_path,encoding="utf-8") as _f:
            _rows = list(csv.DictReader(_f))
        for _r in _rows:
            for _nc in ['kcal','proteina_g','lipidos_g','hdc_g','fibra_g','peso_neto_g']:
                try: _r[_nc]=float(_r[_nc])
                except: _r[_nc]=0.0
        _conn_a.executemany("""INSERT OR IGNORE INTO alimentos
            (nombre,cantidad_ref,unidad,peso_neto_g,kcal,proteina_g,lipidos_g,hdc_g,fibra_g,grupo,ig,sodio_mg,calcio_mg)
            VALUES (:nombre,:cantidad_ref,:unidad,:peso_neto_g,:kcal,:proteina_g,:lipidos_g,:hdc_g,:fibra_g,:grupo,:ig,:sodio_mg,:calcio_mg)""",
            _rows)
        _conn_a.commit(); _conn_a.close()
    except Exception as _ea: pass
_seed_alimentos()

# ── Seed usuarios_sistema desde dict USUARIOS si tabla vacía ─────────────
try:
    _cn_us=get_conn()
    if _cn_us.execute("SELECT COUNT(*) FROM usuarios_sistema").fetchone()[0]==0:
        for _uk,_uv in USUARIOS.items():
            _cn_us.execute("INSERT OR IGNORE INTO usuarios_sistema (usuario,nombre,password_hash,rol,activo) VALUES (?,?,?,?,1)",
                (_uk,_uv["nombre"],_uv["hash"],_uv["rol"]))
        _cn_us.commit()
    _cn_us.close()
except: pass

# ── Backup automático cada 24h (últimos 5) ───────────────────────────────
try:
    import shutil, glob
    _bk_dir = os.path.join(BASE_DIR,"backups")
    os.makedirs(_bk_dir, exist_ok=True)
    _bk_last_f = os.path.join(_bk_dir,"last_backup.txt")
    _bk_do = True
    if os.path.exists(_bk_last_f):
        with open(_bk_last_f) as _f: _bk_last=_f.read().strip()
        try:
            _bk_dt=datetime.fromisoformat(_bk_last)
            if (datetime.now()-_bk_dt).total_seconds() < 86400: _bk_do=False
        except: pass
    if _bk_do and os.path.exists(DB_PATH):
        _bk_name=f"putu_activo_{datetime.now().strftime('%Y-%m-%d_%H-%M')}.db"
        shutil.copy2(DB_PATH, os.path.join(_bk_dir,_bk_name))
        with open(_bk_last_f,"w") as _f: _f.write(datetime.now().isoformat())
        _bk_files=sorted(glob.glob(os.path.join(_bk_dir,"putu_activo_*.db")))
        for _bk_old in _bk_files[:-5]: os.remove(_bk_old)
except: pass

# ── Migrar días de semana a Día 1..6 por rutina (consecutivo según orden de aparición) ──
try:
    _cm=get_conn()
    # Orden natural de días de la semana para determinar secuencia
    _ORDEN_DIAS=["Lunes","Martes","Miércoles","Jueves","Viernes","Sábado","Domingo",
                 "Día 1","Día 2","Día 3","Día 4","Día 5","Día 6"]
    # Obtener todas las rutinas que tienen al menos un día con nombre antiguo
    _rutinas=_cm.execute("SELECT DISTINCT rutina_id FROM rutina_ejercicios").fetchall()
    for (_rid,) in _rutinas:
        _dias_usados=_cm.execute(
            "SELECT DISTINCT dia_semana FROM rutina_ejercicios WHERE rutina_id=? ORDER BY dia_semana",
            (_rid,)).fetchall()
        _dias_usados=[d[0] for d in _dias_usados]
        # Si ya son todos "Día N" consecutivos sin huecos, no hacer nada
        _esperados=[f"Día {i+1}" for i in range(len(_dias_usados))]
        if _dias_usados==_esperados:
            continue
        # Ordenar por orden natural (semana primero, luego Día N)
        def _orden_dia(d):
            try: return _ORDEN_DIAS.index(d)
            except: return 99
        _dias_ord=sorted(_dias_usados,key=_orden_dia)
        # Crear mapeo: día_original → "Día N" consecutivo
        _mapa={_dias_ord[i]:f"Día {i+1}" for i in range(len(_dias_ord))}
        # Aplicar en orden inverso para evitar colisiones (ej: "Día 1"→"Día 1" no pisa)
        for _old,_new in _mapa.items():
            if _old!=_new:
                # Usar valor temporal para evitar colisiones
                _cm.execute("UPDATE rutina_ejercicios SET dia_semana=? WHERE rutina_id=? AND dia_semana=?",
                    (f"__tmp_{_new}__",_rid,_old))
        for _new in _mapa.values():
            _cm.execute("UPDATE rutina_ejercicios SET dia_semana=? WHERE rutina_id=? AND dia_semana=?",
                (_new,_rid,f"__tmp_{_new}__"))
    _cm.commit(); _cm.close()
except: pass

@st.cache_data(ttl=8, show_spinner=False)
def db_query(sql, params=()):
    c = get_conn()
    df = pd.read_sql_query(sql, c, params=params)
    c.close(); return df

def db_exec(sql, params=()):
    c = get_conn(); c.execute(sql, params); c.commit(); c.close()
    db_query.clear()  # Invalidar caché de queries tras escritura
    try: get_clientes.clear()
    except: pass

# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def fmt_fecha(v):
    """Convierte cualquier fecha a dd/mm/yyyy."""
    if not v or str(v) in ["nan","None","NaT","",None]: return ""
    try:
        s = str(v)[:10]
        d = date.fromisoformat(s)
        return d.strftime("%d/%m/%Y")
    except: return str(v)[:10]

def fmt_cel(v):
    if v is None or (isinstance(v,float) and pd.isna(v)): return ""
    try:
        s = str(v).strip().replace(" ","").replace("-","").replace("+","")
        if "." in s:
            try: s = str(int(float(s)))
            except: pass
        if s and not s.startswith("56") and len(s)<=9:
            s = "56"+s.lstrip("0")
        return s
    except: return str(v).strip()

def wa_url(cel, msg):
    num = fmt_cel(cel)
    if not num or len(num)<8: return "#"
    enc = str(msg).replace(" ","%20").replace("\n","%0A").replace("*","%2A").replace('"','%22')
    return f"https://wa.me/{num}?text={enc}"

def generar_qr_b64(texto):
    qr = qrcode.QRCode(version=1,box_size=8,border=2)
    qr.add_data(texto); qr.make(fit=True)
    img = qr.make_image(fill_color="#6DBE45",back_color="#1E1E1E")
    buf = io.BytesIO(); img.save(buf,format="PNG")
    return base64.b64encode(buf.getvalue()).decode()

def decodificar_qr(imagen_bytes):
    """Decodifica un QR desde bytes de imagen. Retorna el RUT extraído o None.
    Espera formato PUTU|RUT|NOMBRE generado por generar_qr_b64."""
    if not QR_SCAN_DISPONIBLE:
        return None
    try:
        arr = np.frombuffer(imagen_bytes, dtype=np.uint8)
        img = cv2.imdecode(arr, cv2.IMREAD_COLOR)
        if img is None: return None
        detector = cv2.QRCodeDetector()
        data, _, _ = detector.detectAndDecode(img)
        if not data: return None
        partes = data.split("|")
        if len(partes) >= 2 and partes[0] == "PUTU":
            return partes[1].strip().upper()
        return None
    except Exception:
        return None

def _vence_pronto(x, dias=30):
    try:
        h=date.today()
        if isinstance(x,str): x=date.fromisoformat(x[:10])
        if isinstance(x,datetime): x=x.date()
        return isinstance(x,date) and 0<=(x-h).days<=dias
    except: return False

def dias_para_vencer(x):
    try:
        h=date.today()
        if isinstance(x,str): x=date.fromisoformat(x[:10])
        if isinstance(x,datetime): x=x.date()
        return (x-h).days
    except: return None

def calcular_vencimiento(fecha_inicio, periodo):
    """Calcula fecha de vencimiento según periodo."""
    try:
        if isinstance(fecha_inicio,str): fecha_inicio=date.fromisoformat(fecha_inicio[:10])
        dias = PERIODO_DIAS.get(periodo,30)
        return str(fecha_inicio + timedelta(days=dias))
    except: return ""

def mes_de_nacimiento(fecha_nac):
    """Extrae mes en español desde fecha."""
    try:
        if isinstance(fecha_nac,str): fecha_nac=date.fromisoformat(fecha_nac[:10])
        if isinstance(fecha_nac,datetime): fecha_nac=fecha_nac.date()
        return MESES_ESP[fecha_nac.month-1]
    except: return ""

def msg_cumpleanos(nombre):
    return f"Hola *{nombre}*. %0AMuchas felicidades! 🥳 El equipo de Putú Activo te envía un gran abrazo en tu día especial."

def msg_vencimiento(nombre, fecha_venc):
    fv = str(fecha_venc)[:10] if fecha_venc else "pronto"
    return f"Hola *{nombre}*, tu plan en Putú Activo vence el *{fv}*. ¡Renueva y sigue alcanzando tus objetivos! 💪"

def msg_renovacion(nombre, fecha_venc):
    fv = str(fecha_venc)[:10] if fecha_venc else ""
    return f"Hola *{nombre}*, tu plan en Putú Activo se renovó hasta el *{fv}*. ¡Sigue entrenando con nosotros! 🏋️"

def log_action(accion, detalle="", rut_afectado=""):
    """Registra actividad en el log."""
    try:
        import streamlit as _st2
        _usr=_st2.session_state.get("usuario","sistema")
        _rol=_st2.session_state.get("rol","")
        _cn=get_conn()
        _cn.execute("INSERT INTO activity_log (timestamp,usuario,rol,accion,detalle,rut_afectado) VALUES (?,?,?,?,?,?)",
            (datetime.now().isoformat(),_usr,_rol,accion,str(detalle)[:500],str(rut_afectado)))
        _cn.commit(); _cn.close()
    except: pass

def sv(row, col, d=""):
    v = row.get(col,d) if isinstance(row,dict) else getattr(row,col,d)
    s = d if str(v) in ["nan","None","NaT","<NA>",""] else str(v)
    # Sanitizar errores de locale Windows en español Chile
    if col in ("periodo_vencimiento","concepto") and s=="Menstrual": s="Mensual"
    return s

# ─────────────────────────────────────────────────────────────────────────────
# IMPORTAR / EXPORTAR
# ─────────────────────────────────────────────────────────────────────────────
def importar_excel_full(uploaded_file):
    """Importa todas las tablas desde Excel unificado."""
    xl = pd.ExcelFile(uploaded_file, engine="openpyxl")
    n_total = 0
    if "clientes" in xl.sheet_names:
        df = xl.parse("clientes")
        df = df[df["nombre"].notna() & (df["nombre"].astype(str).str.strip()!="")]
        for _,r in df.iterrows():
            try:
                db_exec("""INSERT OR REPLACE INTO clientes
                    (nombre,rut,fecha_nacimiento,edad,mes_cumpleanos,sexo,direccion,celular,email,
                     contacto_emergencia,celular_emergencia,parentesco,fecha_inscripcion,
                     periodo_vencimiento,tipo_plan,frecuencia,horario,
                     lunes,martes,miercoles,jueves,viernes,sabado,
                     valor_plan,fecha_vencimiento,fecha_renovacion,estado,rutina,nivel,
                     enfermedad,restricciones,objetivo,talla,foto_path,rutina_pdf_path,
                     mensaje_cumpleanos,mensaje_vencimiento,mensaje_renovacion,creado,modificado)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                    tuple(str(r.get(c,"")) if pd.notna(r.get(c,"")) else "" for c in [
                        "nombre","rut","fecha_nacimiento","edad","mes_cumpleanos","sexo","direccion","celular","email",
                        "contacto_emergencia","celular_emergencia","parentesco","fecha_inscripcion",
                        "periodo_vencimiento","tipo_plan","frecuencia","horario",
                        "lunes","martes","miercoles","jueves","viernes","sabado",
                        "valor_plan","fecha_vencimiento","fecha_renovacion","estado","rutina","nivel",
                        "enfermedad","restricciones","objetivo","talla","foto_path","rutina_pdf_path",
                        "mensaje_cumpleanos","mensaje_vencimiento","mensaje_renovacion","creado","modificado"]))
                n_total+=1
            except: pass
    for tabla in ["pagos","asistencia","evaluaciones","clases","egresos","productos"]:
        if tabla in xl.sheet_names:
            df2 = xl.parse(tabla)
            for _,r in df2.iterrows():
                try:
                    cols = [c for c in df2.columns if c!="id"]
                    vals = [str(r.get(c,"")) if pd.notna(r.get(c,"")) else "" for c in cols]
                    ph   = ",".join(["?"]*len(cols))
                    db_exec(f"INSERT INTO {tabla} ({','.join(cols)}) VALUES ({ph})", vals)
                    n_total+=1
                except: pass
    return n_total

def importar_excel_clientes_legacy(uploaded_file):
    """Importa desde datosgym.xlsx original."""
    xl = pd.read_excel(uploaded_file, sheet_name="BBDD Clientes", header=0, engine="openpyxl")
    xl = xl[xl["Nombre"].notna() & (xl["Nombre"].astype(str).str.strip()!="")]
    n=0
    for _,r in xl.iterrows():
        nombre  = str(r.get("Nombre","")).strip()
        rut     = str(r.get("Rut","")).strip()
        if not nombre or not rut: continue
        fn_raw  = r.get("Fecha Nacimiento","")
        fn_str  = str(fn_raw)[:10] if pd.notna(fn_raw) else ""
        mes_c   = mes_de_nacimiento(fn_str) if fn_str else str(r.get("Mes cumpleaños",""))
        edad    = int(r["Edad"]) if pd.notna(r.get("Edad")) else 0
        cel1    = fmt_cel(r.get("Celular",""))
        cel2    = fmt_cel(r.get("Celular.1",""))
        fi      = str(r.get("Fecha de Inscripción",""))[:10] if pd.notna(r.get("Fecha de Inscripción")) else ""
        fv      = str(r.get("Fecha VENCIMIENTO",""))[:10] if pd.notna(r.get("Fecha VENCIMIENTO")) else ""
        fr_     = str(r.get("Fecha RENOVACIÓN",""))[:10] if pd.notna(r.get("Fecha RENOVACIÓN")) else ""
        periodo = str(r.get("Periodo_Vencimiento","Mensual")) if pd.notna(r.get("Periodo_Vencimiento")) else "Mensual"
        tp      = str(r.get("Tipo plan","")).strip()
        mc      = msg_cumpleanos(nombre)
        mv      = msg_vencimiento(nombre, fv)
        mr      = msg_renovacion(nombre, fv)
        try:
            db_exec("""INSERT OR IGNORE INTO clientes
                (nombre,rut,fecha_nacimiento,edad,mes_cumpleanos,sexo,direccion,celular,email,
                 contacto_emergencia,celular_emergencia,parentesco,fecha_inscripcion,
                 periodo_vencimiento,tipo_plan,frecuencia,horario,
                 lunes,martes,miercoles,jueves,viernes,sabado,
                 valor_plan,fecha_vencimiento,fecha_renovacion,estado,rutina,nivel,
                 enfermedad,restricciones,objetivo,talla,foto_path,rutina_pdf_path,
                 mensaje_cumpleanos,mensaje_vencimiento,mensaje_renovacion,creado,modificado)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (nombre,rut,fn_str,edad,mes_c,
                 str(r.get("Sexo","")),str(r.get("Dirección Particular","")),cel1,str(r.get("E-mail","")),
                 str(r.get("Contacto Emergencia","")),cel2,str(r.get("Parentesco","")),
                 fi,periodo,tp,str(r.get("Frecuencia semanal","")),str(r.get("Horario estimado","")),
                 str(r.get("Día 1","")),str(r.get("Día 2","")),str(r.get("Día 3","")),
                 str(r.get("Día 4","")),str(r.get("Día 5","")),str(r.get("Día 6","")),
                 float(r["Valor $ pagado"]) if pd.notna(r.get("Valor $ pagado")) else 0,
                 fv,fr_,str(r.get("Estado / pago","")),
                 str(r.get("Rutina","")),str(r.get("Nivel","")),
                 str(r.get("Enfermedad:","")),str(r.get("Restricciones","")),
                 str(r.get("Objetivo","")),str(r.get("Talla","")),
                 "","",mc,mv,mr,
                 datetime.now().isoformat(),datetime.now().isoformat()))
            n+=1
        except: pass
    return n

def exportar_todo_excel():
    """Exporta TODA la base de datos en un único Excel con múltiples hojas."""
    buf = io.BytesIO()
    tablas = ["clientes","pagos","asistencia","evaluaciones","clases","egresos","productos"]
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for t in tablas:
            df = db_query(f"SELECT * FROM {t}")
            df.to_excel(w, index=False, sheet_name=t)
    return buf.getvalue()

# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Putú Activo",page_icon="🏋️",
                   layout="wide",initial_sidebar_state="auto")

# ─────────────────────────────────────────────────────────────────────────────
# CSS
# ─────────────────────────────────────────────────────────────────────────────
st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700;900&display=swap');
html,body,[class*="css"]{{font-family:'Inter',sans-serif;background:{NEGRO};color:{BLANCO};font-size:17px;}}
.stApp{{background:{NEGRO};}}
  .block-container{{padding-top:0.2rem !important;padding-bottom:0rem !important;max-width:100% !important;}}
  /* Reducir espacios internos para aprovechar pantalla */
  div[data-testid="stVerticalBlock"] > div {{gap:0.25rem !important;}}
  div[data-testid="stHorizontalBlock"] {{gap:0.4rem !important;}}
  .element-container {{margin-bottom:0.1rem !important;}}
  /* Tabs pegados sin espacio extra */
  .stTabs [data-baseweb="tab-panel"] {{padding-top:0.4rem !important;padding-bottom:0.2rem !important;}}
  div[data-testid="stTabsContent"] {{padding-top:0.3rem !important;}}
  /* Reducir espacio entre sección y tabs */
  hr {{margin:0.3rem 0 !important;}}
  /* Métricas más compactas */
  div[data-testid="metric-container"] {{padding:10px 14px !important;}}
  div[data-testid="metric-container"] div[data-testid="metric-value"] {{font-size:1.8rem !important;}}
  header[data-testid="stHeader"]{{height:0 !important;min-height:0 !important;display:none !important;}}
  div[data-testid="stToolbar"]{{display:none !important;}}
  div[data-testid="stDecoration"]{{display:none !important;}}
  #MainMenu{{visibility:hidden;}}
  footer{{visibility:hidden;}}
  /* Sidebar compacto desde arriba */
  section[data-testid="stSidebar"] > div:first-child{{padding-top:0 !important;padding-bottom:0 !important;}}
  section[data-testid="stSidebar"] .block-container{{padding-top:0 !important;}}
  /* Reducir espacio sobre tabs */
  .stTabs{{margin-top:-0.5rem;}}
  /* Reducir espacio entre elementos del formulario */
  div[data-testid="stVerticalBlock"] > div{{gap:0.3rem !important;}}
  /* Section headers más compactos */
  /* Botones eliminar / quitar / borrar → Rojo */
  button[data-testid="baseButton-secondary"]:has(div:contains("🗑️")) {{
    background:{ROJO} !important; color:white !important; border:2px solid {ROJO} !important;
  }}
  button[data-testid="baseButton-secondary"]:has(div:contains("Quitar")) {{
    background:{ROJO} !important; color:white !important; border:2px solid {ROJO} !important;
  }}
  button[data-testid="baseButton-secondary"]:has(div:contains("Eliminar")) {{
    background:{ROJO} !important; color:white !important; border:2px solid {ROJO} !important;
  }}
  button[data-testid="baseButton-secondary"]:has(div:contains("Borrar")) {{
    background:{ROJO} !important; color:white !important; border:2px solid {ROJO} !important;
  }}
  button[data-testid="baseButton-secondary"]:has(div:contains("Sí, eliminar")) {{
    background:{ROJO} !important; color:white !important; border:2px solid {ROJO} !important;
  }}
  button[data-testid="baseButton-secondary"]:has(div:contains("eliminar")) {{
    background:{ROJO} !important; color:white !important; border:2px solid {ROJO} !important;
  }}
  /* Botones Ver detalle → Amarillo */
  button[data-testid="baseButton-secondary"]:has(div:contains("Ver detalle")) {{
    background:#F5C518 !important; color:#000000 !important; border:2px solid #F5C518 !important;
    font-weight:700 !important;
  }}
  /* Fallback para browsers sin :has() */
  .del-btn-red > button {{
    background:{ROJO} !important; color:white !important; border:2px solid {ROJO} !important;
  }}
  /* Sidebar — eliminar padding superior para subir contenido al tope */
  section[data-testid="stSidebar"] > div:first-child {{
    padding-top: 0 !important;
    margin-top: 0 !important;
  }}
  section[data-testid="stSidebar"] .block-container {{
    padding-top: 0 !important;
  }}

  /* Botones Salir e Inicio del sidebar — más pequeños */
  div[data-testid="stSidebar"] button[data-testid="baseButton-secondary"] {{
    font-size:.72rem !important;
    padding:4px 8px !important;
    min-height:28px !important;
  }}
  @media (max-width:768px){{
    .block-container{{padding-left:0.5rem !important;padding-right:0.5rem !important;}}
    div[data-testid="metric-container"]{{padding:10px 12px;}}
    div[data-testid="metric-container"] div[data-testid="metric-value"]{{font-size:2rem;}}
      .stTabs [data-baseweb="tab"]{{padding:7px 10px;font-size:.88rem;}}
    .rut-display{{font-size:2rem !important;padding:12px !important;}}
    .stButton>button{{padding:12px 16px;font-size:1.05rem;}}
  }}
  /* Viewport meta vía CSS para permitir zoom en móvil */
  @-ms-viewport{{width:device-width;}}
p,span,label,li{{color:{BLANCO};}}
section[data-testid="stSidebar"]{{background:{GRIS} !important;border-right:1px solid {GRIS3};}}
section[data-testid="stSidebar"] *{{color:{BLANCO} !important;font-size:1.15rem !important;}}
section[data-testid="stSidebar"] .stRadio label{{font-size:.95rem !important;padding:3px 4px;font-weight:600;line-height:1.2;}}
div[data-testid="metric-container"]{{background:{GRIS2};border:1px solid {GRIS3};border-radius:14px;padding:18px 22px;border-left:5px solid {VERDE};}}
div[data-testid="metric-container"] label{{color:{GRIS_T} !important;font-size:.85rem;text-transform:uppercase;letter-spacing:.07em;}}
div[data-testid="metric-container"] div[data-testid="metric-value"]{{color:{VERDE} !important;font-size:2.2rem;font-weight:900;}}
.stButton>button{{background:{VERDE};color:{NEGRO};font-weight:700;border:none;border-radius:9px;padding:10px 24px;font-size:1rem;}}
.stButton>button:hover{{background:{VERDE_DK};color:{BLANCO};}}
/* Botón collapse sidebar — solo color, sin tocar display/visibility/position */
button[data-testid="baseButton-headerNoPadding"] {{
    background:transparent !important;
    box-shadow:none !important;
}}
button[data-testid="baseButton-headerNoPadding"] svg {{
    fill:{VERDE} !important;
    color:{VERDE} !important;
}}
div[data-testid="stFormSubmitButton"]>button{{background:{VERDE};color:{NEGRO};font-weight:700;border:none;border-radius:9px;padding:12px 32px;font-size:1rem;}}
div[data-testid="stFormSubmitButton"]>button:hover{{background:{VERDE_DK};color:{BLANCO};}}
.stTextInput input,.stSelectbox select,.stNumberInput input,.stDateInput input,.stTextArea textarea{{background:{GRIS2} !important;color:{BLANCO} !important;border:1px solid {GRIS3} !important;border-radius:9px !important;font-size:1.05rem !important;}}
.stTextInput label,.stSelectbox label,.stNumberInput label,.stDateInput label,.stTextArea label{{color:{GRIS_T} !important;font-size:1rem;font-weight:600;}}
.stTabs [data-baseweb="tab-list"]{{background:{GRIS2};border-radius:11px;gap:4px;padding:5px;}}
.stTabs [data-baseweb="tab"]{{background:transparent;color:{GRIS_T};border-radius:9px;padding:10px 20px;font-size:.97rem;font-weight:500;}}
.stTabs [aria-selected="true"]{{background:{VERDE} !important;color:{NEGRO} !important;font-weight:700;}}
.stDataFrame{{border-radius:11px;overflow:hidden;}}
.stDataFrame thead th{{background:{VERDE} !important;color:{NEGRO} !important;font-weight:700;}}
.stDataFrame tbody td{{background:{GRIS2} !important;color:{BLANCO} !important;}}
.section-header{{font-size:1.5rem;font-weight:800;color:{VERDE};margin-bottom:12px;padding-bottom:7px;border-bottom:2px solid {GRIS3};}}
.card{{background:{GRIS2};border:1px solid {GRIS3};border-radius:13px;padding:20px;margin-bottom:12px;}}
.alert-box{{background:#ff444422;border-left:5px solid #ff4444;border-radius:9px;padding:13px 17px;color:#ffaaaa;margin:7px 0;}}
.success-box{{background:{VERDE}22;border-left:5px solid {VERDE};border-radius:9px;padding:13px 17px;color:{VERDE};margin:7px 0;}}
.info-box{{background:{AZUL}22;border-left:5px solid {AZUL};border-radius:9px;padding:13px 17px;color:#7dc8f0;margin:7px 0;}}
.warn-box{{background:{NARANJA}22;border-left:5px solid {NARANJA};border-radius:9px;padding:13px 17px;color:{NARANJA};margin:7px 0;}}
div[data-testid="stExpander"]{{background:{GRIS2};border:1px solid {GRIS3};border-radius:11px;}}
.stDownloadButton>button{{background:{GRIS2};color:{VERDE};border:1px solid {VERDE};border-radius:9px;}}
/* Teclado QR */
.numpad-btn{{font-size:2rem;font-weight:900;background:{GRIS3};color:{BLANCO};border:2px solid {GRIS4};border-radius:14px;padding:18px;text-align:center;cursor:pointer;margin:4px;transition:.15s;}}
.numpad-btn:hover{{background:{VERDE};color:{NEGRO};}}
.rut-display{{font-size:3.5rem;font-weight:900;color:{VERDE};text-align:center;letter-spacing:.15em;background:{GRIS2};border-radius:14px;padding:24px;margin-bottom:16px;border:2px solid {GRIS3};min-height:100px;}}
/* Pantalla aeropuerto */
.airport-row{{background:{GRIS2};border-bottom:1px solid {GRIS3};padding:10px 20px;display:flex;justify-content:space-between;align-items:center;animation:fadein .5s;}}
@keyframes fadein{{from{{opacity:0}}to{{opacity:1}}}}
@media print{{section[data-testid="stSidebar"],div[data-testid="stToolbar"],.stButton,.stDownloadButton,.no-print{{display:none !important;}}body{{background:white !important;color:#111 !important;}}}}
</style>
""", unsafe_allow_html=True)

PL = dict(paper_bgcolor="rgba(0,0,0,0)",plot_bgcolor="rgba(0,0,0,0)",
    font=dict(family="Inter",color=BLANCO,size=13),title_font=dict(color=BLANCO,size=15),
    legend=dict(bgcolor="rgba(0,0,0,0)",font=dict(color=BLANCO,size=12)),
    colorway=[VERDE,AZUL,NARANJA,ROJO,"#A855F7","#06B6D4"],
    margin=dict(l=16,r=16,t=42,b=16),
    xaxis=dict(gridcolor=GRIS3,tickfont=dict(color=GRIS_T,size=12)),
    yaxis=dict(gridcolor=GRIS3,tickfont=dict(color=GRIS_T,size=12)))

# ─────────────────────────────────────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────────────────────────────────────
for k,v in [("logueado",False),("login_time",None),("usuario",""),
             ("rol",""),("nombre_u",""),("ver_rut",None),
             ("pago_rut",None),("rut_buffer",""),
             ("modo",""),  # "cliente" | "admin"
             ("rut_cliente",""),  # RUT ingresado en modo cliente
             ]:
    if k not in st.session_state: st.session_state[k]=v

def sesion_valida():
    if not st.session_state.get("logueado"): return False
    if not st.session_state.get("login_time"): return False
    elapsed = time.time() - st.session_state.login_time
    if elapsed >= SESSION_H * 3600:
        return False
    # Sliding window: renovar timer si el usuario sigue activo
    if elapsed > 60:  # renovar solo después de 1 min para no escribir en cada frame
        st.session_state.login_time = time.time()
    return True

# ─────────────────────────────────────────────────────────────────────────────
# LOGIN
# ─────────────────────────────────────────────────────────────────────────────
# ── Pantalla de selección de modo si no hay sesión ni modo elegido ──────────
if not sesion_valida() and not st.session_state.modo:
    st.session_state.logueado = False
    # Logo pequeño centrado
    if LOGO_PATH:
        st.markdown(f'<div style="text-align:center;margin-bottom:8px"><img src="data:image/png;base64,{__import__("base64").b64encode(open(LOGO_PATH,"rb").read()).decode()}" style="width:180px;max-width:60vw"></div>',unsafe_allow_html=True)
    else:
        st.markdown(f'<div style="text-align:center;font-size:2rem;font-weight:900;color:{VERDE}">🏋️ PUTÚ ACTIVO</div>',unsafe_allow_html=True)
    st.markdown(f'<div style="text-align:center;color:{GRIS_T};font-size:.95rem;margin:4px 0 16px;letter-spacing:.08em;">CENTRO DE ENTRENAMIENTO</div>',unsafe_allow_html=True)
    _,bc1,bc2,bc3,_ = st.columns([1,1.4,1.4,1.4,1])
    with bc1:
        st.markdown(f'<style>div[data-testid="stButton"]:has(button[data-testid="btn_cliente"]) button{{background:{GRIS2}!important;border:3px solid {VERDE}!important;border-radius:16px!important;padding:28px 16px!important;height:auto!important;white-space:normal!important;color:{VERDE}!important;font-size:1.1rem!important;font-weight:900!important;line-height:2!important;letter-spacing:.05em}}</style>',unsafe_allow_html=True)
        if st.button("🏋️\n\nASISTENCIA\n\nMarcar entrada · Salida", key="btn_cliente", use_container_width=True):
            st.session_state.modo="cliente"; st.rerun()
    with bc2:
        st.markdown(f'<style>div[data-testid="stButton"]:has(button[data-testid="btn_micuenta"]) button{{background:{GRIS2}!important;border:3px solid {AZUL}!important;border-radius:16px!important;padding:28px 16px!important;height:auto!important;white-space:normal!important;color:{AZUL}!important;font-size:1.1rem!important;font-weight:900!important;line-height:2!important;letter-spacing:.05em}}</style>',unsafe_allow_html=True)
        if st.button("👤\n\nMI CUENTA\n\nRutina · Evaluación · Pagos", key="btn_micuenta", use_container_width=True):
            st.session_state.modo="micuenta"; st.rerun()
    with bc3:
        st.markdown(f'<style>div[data-testid="stButton"]:has(button[data-testid="btn_admin"]) button{{background:{GRIS2}!important;border:3px solid #888!important;border-radius:16px!important;padding:28px 16px!important;height:auto!important;white-space:normal!important;color:{BLANCO}!important;font-size:1.1rem!important;font-weight:900!important;line-height:2!important;letter-spacing:.05em}}</style>',unsafe_allow_html=True)
        if st.button("🔐\n\nADMIN\n\nGestión completa", key="btn_admin", use_container_width=True):
            st.session_state.modo="admin"; st.rerun()
    st.stop()

# ── Login ADMIN ──────────────────────────────────────────────────────────────
# Solo mostrar login si modo es admin Y no hay sesión válida Y no viene de un rerun de navegación
if st.session_state.modo == "admin" and not sesion_valida() and not st.session_state.get("logueado"):
    _,cc,_ = st.columns([1.2,0.8,1.2])
    with cc:
        if LOGO_PATH:
            import base64 as _b64a
            _logo_b64a=_b64a.b64encode(open(LOGO_PATH,"rb").read()).decode()
            st.markdown(f'<div style="text-align:center;margin:0 auto 6px auto"><img src="data:image/png;base64,{_logo_b64a}" style="width:120px;max-width:55vw;display:block;margin:0 auto"></div>',unsafe_allow_html=True)
        st.markdown(f'<div style="text-align:center;padding:4px 0;color:{GRIS_T};font-size:.88rem;letter-spacing:.1em">ADMINISTRACIÓN</div>',unsafe_allow_html=True)
        with st.form("login_admin"):
            u=st.text_input("👤  Usuario",placeholder="admin / entrenador / asistente")
            p=st.text_input("🔒  Contraseña",type="password",placeholder="••••••••")
            ok=st.form_submit_button("Ingresar →",use_container_width=True)
        if ok:
            uu=u.strip().lower()
            # Buscar en BD primero, luego fallback a dict
            _us_row=get_conn().execute("SELECT nombre,rol,password_hash FROM usuarios_sistema WHERE usuario=? AND activo=1",(uu,)).fetchone()
            if _us_row and _us_row[2]==_h(p):
                st.session_state.update({"logueado":True,"login_time":time.time(),
                    "usuario":uu,"rol":_us_row[1],"nombre_u":_us_row[0],"ver_rut":None,"modo":"admin"})
                log_action("LOGIN",f"Acceso exitoso de {_us_row[0]}"); st.rerun()
            elif uu in USUARIOS and USUARIOS[uu]["hash"]==_h(p):
                st.session_state.update({"logueado":True,"login_time":time.time(),
                    "usuario":uu,"rol":USUARIOS[uu]["rol"],"nombre_u":USUARIOS[uu]["nombre"],
                    "ver_rut":None,"modo":"admin"})
                log_action("LOGIN",f"Acceso exitoso (fallback) de {USUARIOS[uu]['nombre']}"); st.rerun()
            else:
                st.markdown('<div class="alert-box">⚠️ Usuario o contraseña incorrectos.</div>',unsafe_allow_html=True)
        st.markdown(f'<div style="background:{GRIS2};border-radius:10px;padding:12px 16px;margin-top:12px;font-size:.82rem;color:#666;">admin/putu2025 · entrenador/gym123 · asistente/asist1</div>',unsafe_allow_html=True)
        if st.button("← Volver al inicio", key="volver_login"):
            st.session_state.modo=""; st.rerun()
    st.stop()

# ── Login MI CUENTA (cliente con email+password) ────────────────────────────
if st.session_state.modo == "micuenta" and not st.session_state.get("cliente_logueado"):
    _,_cc2,_ = st.columns([1.2,0.8,1.2])
    with _cc2:
        if LOGO_PATH:
            import base64 as _b64c
            _logo_b64c=_b64c.b64encode(open(LOGO_PATH,"rb").read()).decode()
            st.markdown(f'<div style="text-align:center;margin:0 auto 6px auto"><img src="data:image/png;base64,{_logo_b64c}" style="width:120px;max-width:55vw;display:block;margin:0 auto"></div>',unsafe_allow_html=True)
        st.markdown(f'<div style="text-align:center;padding:4px 0;color:{GRIS_T};font-size:.88rem;letter-spacing:.1em">MI CUENTA</div>',unsafe_allow_html=True)
        with st.form("login_cliente"):
            _email_c=st.text_input("✉️ Email",placeholder="tu@email.com")
            _pass_c=st.text_input("🔒 Contraseña",type="password",placeholder="••••••••")
            _ok_c=st.form_submit_button("Ingresar →",use_container_width=True)
        if _ok_c:
            _uc=get_conn().execute(
                "SELECT rut,activo FROM usuarios_clientes WHERE email=? AND password_hash=?",
                (_email_c.strip().lower(),_h(_pass_c))).fetchone()
            if _uc and _uc[1]==1:
                get_conn().execute("UPDATE usuarios_clientes SET ultimo_acceso=? WHERE rut=?",
                    (datetime.now().isoformat(),_uc[0]))
                get_conn().commit()
                st.session_state.update({"cliente_logueado":True,"cliente_rut":_uc[0],"modo":"micuenta"})
                st.rerun()
            else:
                st.markdown('<div class="alert-box">⚠️ Email o contraseña incorrectos, o cuenta inactiva.</div>',unsafe_allow_html=True)
        if st.button("← Volver al inicio",key="volver_micuenta"):
            st.session_state.modo=""; st.rerun()
    st.stop()

# ── Vista MI CUENTA — ficha personal del cliente ─────────────────────────────
if st.session_state.modo == "micuenta" and st.session_state.get("cliente_logueado"):
    _rut_mc=st.session_state.get("cliente_rut","")
    _df_mc=db_query("SELECT * FROM clientes WHERE rut=?",(_rut_mc,))
    if _df_mc.empty:
        st.error("No se encontró tu ficha. Contacta al administrador.")
        if st.button("← Salir"): st.session_state.update({"cliente_logueado":False,"modo":""}); st.rerun()
        st.stop()
    _r_mc=_df_mc.iloc[0].to_dict()
    # Encabezado
    _mc1,_mc2=st.columns([5,1])
    _mc1.markdown(f'<div style="border-left:4px solid {AZUL};padding-left:12px;"><span style="font-size:1.3rem;font-weight:900;color:{AZUL}">{sv(_r_mc,"nombre")}</span> <span style="color:{GRIS_T};font-size:.85rem">· {_rut_mc}</span></div>',unsafe_allow_html=True)
    if _mc2.button("🚪 Salir",key="mc_salir"):
        st.session_state.update({"cliente_logueado":False,"cliente_rut":"","modo":""}); st.rerun()
    st.markdown("---")
    # Tabs solo lectura
    _tmc1,_tmc2,_tmc3,_tmc4=st.tabs(["💪 Mi Rutina","📏 Evaluaciones","💳 Pagos","🥗 Nutrición"])

    with _tmc1:
        _rut_mc_r=db_query("SELECT * FROM rutinas WHERE cliente_rut=? AND activa=1 ORDER BY id DESC LIMIT 1",(_rut_mc,))
        if _rut_mc_r.empty:
            st.markdown('<div class="info-box">Sin rutina activa. Consulta a tu entrenador.</div>',unsafe_allow_html=True)
        else:
            _rutmc=_rut_mc_r.iloc[0].to_dict()
            st.markdown(f"<b style='color:{AZUL}'>{_rutmc['nombre']}</b>",unsafe_allow_html=True)
            _ejs_mc=db_query("""SELECT re.*,e.nombre,e.url_imagen,e.musculo_primario
                FROM rutina_ejercicios re JOIN ejercicios e ON e.id=re.ejercicio_id
                WHERE re.rutina_id=? ORDER BY re.dia_semana,re.orden""",(int(_rutmc["id"]),))
            if not _ejs_mc.empty:
                _dias_mc=[d for d in ["Día 1","Día 2","Día 3","Día 4","Día 5","Día 6"] if d in _ejs_mc["dia_semana"].values]
                _hcols_mc=st.columns(len(_dias_mc)) if _dias_mc else []
                for _ic_mc,_cc_mc in enumerate(_hcols_mc):
                    _cc_mc.markdown(f'<div style="background:{AZUL};color:#fff;text-align:center;font-weight:700;font-size:.78rem;padding:5px 2px;border-radius:6px 6px 0 0">{_dias_mc[_ic_mc]}</div>',unsafe_allow_html=True)
                _bcols_mc=st.columns(len(_dias_mc)) if _dias_mc else []
                for _ic2_mc,(_cb_mc,_dfl_mc) in enumerate(zip(_bcols_mc,_dias_mc)):
                    _ddf_mc=_ejs_mc[_ejs_mc["dia_semana"]==_dfl_mc]
                    with _cb_mc:
                        for _ii_mc,(_,_ef_mc) in enumerate(_ddf_mc.iterrows(),1):
                            _efd_mc=_ef_mc.to_dict()
                            with st.container(border=True):
                                _iurl_mc=str(_efd_mc.get("url_imagen","")).strip()
                                if _iurl_mc and _iurl_mc!="nan":
                                    try: st.image(_iurl_mc,use_container_width=True)
                                    except: pass
                                st.markdown(f'<div style="font-size:.72rem;font-weight:700;color:{BLANCO}">{_ii_mc}. {_efd_mc.get("nombre","")}</div>',unsafe_allow_html=True)
                                st.markdown(f'<div style="font-size:.68rem;color:{AZUL}">{sv(_efd_mc,"series","—")}×{sv(_efd_mc,"repeticiones","—")}</div>',unsafe_allow_html=True)

    with _tmc2:
        _ev_mc=db_query("SELECT * FROM evaluaciones WHERE rut=? ORDER BY fecha DESC",(_rut_mc,))
        if _ev_mc.empty:
            st.markdown('<div class="info-box">Sin evaluaciones registradas.</div>',unsafe_allow_html=True)
        else:
            _ult_mc=_ev_mc.iloc[0]
            _em1,_em2,_em3,_em4=st.columns(4)
            def _emv_mc(k,u=""): v=_ult_mc.get(k); return f"{float(v):.1f}{u}" if v and str(v) not in ["nan","None","0.0"] else "—"
            _em1.metric("⚖️ Peso",_emv_mc("peso"," kg")); _em2.metric("📊 IMC",_emv_mc("imc"))
            _em3.metric("🔴 Grasa",_emv_mc("grasa_pct"," %")); _em4.metric("💪 Musc.",_emv_mc("masa_musc"," %"))
            _ev_mc_d=_ev_mc.copy(); _ev_mc_d["fecha"]=_ev_mc_d["fecha"].apply(fmt_fecha)
            st.dataframe(_ev_mc_d[["fecha","peso","imc","grasa_pct","masa_musc","agua_pct"]].rename(columns={"fecha":"Fecha","peso":"Peso kg","imc":"IMC","grasa_pct":"Grasa %","masa_musc":"Musc. %","agua_pct":"Agua %"}),use_container_width=True,hide_index=True)

    with _tmc3:
        _pg_mc=db_query("SELECT fecha,monto,concepto,tipo_plan,medio_pago FROM pagos WHERE rut=? ORDER BY fecha DESC",(_rut_mc,))
        if _pg_mc.empty:
            st.markdown('<div class="info-box">Sin pagos registrados.</div>',unsafe_allow_html=True)
        else:
            _pg_mc["fecha"]=_pg_mc["fecha"].apply(fmt_fecha)
            _pg_mc["monto"]=_pg_mc["monto"].apply(lambda x:f"${int(x):,}")
            st.dataframe(_pg_mc.rename(columns={"fecha":"Fecha","monto":"Monto","concepto":"Período","tipo_plan":"Plan","medio_pago":"Medio"}),use_container_width=True,hide_index=True)

    with _tmc4:
        _pn_mc=db_query("SELECT * FROM planes_nutri WHERE cliente_rut=? AND activo=1 ORDER BY id DESC LIMIT 1",(_rut_mc,))
        if _pn_mc.empty:
            st.markdown('<div class="info-box">Sin plan nutricional activo.</div>',unsafe_allow_html=True)
        else:
            _pnr_mc=_pn_mc.iloc[0].to_dict()
            st.markdown(f'<div style="background:{GRIS2};border-radius:8px;padding:10px 14px;"><b style="color:{AZUL}">{_pnr_mc["nombre"]}</b> · 🎯 {_pnr_mc.get("objetivo","—")} · 🔥 {int(_pnr_mc.get("kcal_objetivo") or 0):,} kcal/día</div>',unsafe_allow_html=True)
            st.markdown(f"<span style='color:{GRIS_T};font-size:.85rem'>Profesional: {_pnr_mc.get('profesional','—')}</span>",unsafe_allow_html=True)
    st.stop()

hoy = date.today()

# ── Auto-inactivar clientes vencidos ────────────────────────────────────────
try:
    _conn_auto_inac = get_conn()
    _conn_auto_inac.execute("""
        UPDATE clientes SET estado='Inactivo', modificado=?
        WHERE estado='Activo'
          AND fecha_vencimiento IS NOT NULL
          AND fecha_vencimiento != ''
          AND fecha_vencimiento < ?
    """, (datetime.now().isoformat(), str(hoy)))
    _conn_auto_inac.commit()
    _conn_auto_inac.close()
except: pass

# ════════════════════════════════════════════════════════════════════════════
# MODO CLIENTE — sin login, sin sidebar completo
# ════════════════════════════════════════════════════════════════════════════
if st.session_state.modo == "cliente":

    # Auto-salida +3h
    try:
        _lim=(datetime.now()-timedelta(hours=3)).strftime("%H:%M")
        _ca=get_conn()
        _ca.execute("UPDATE asistencia SET hora_salida='Auto' WHERE fecha=? AND tipo='ingreso' AND hora_salida IS NULL AND hora<=?",(str(hoy),_lim))
        _ca.commit(); _ca.close()
    except: pass

    # Sidebar mínimo para modo cliente
    with st.sidebar:
        if LOGO_PATH:
            st.markdown(f'''<div style="text-align:center;padding:0;margin:-12px -8px 6px -8px;">
              <img src="data:image/png;base64,{__import__("base64").b64encode(open(LOGO_PATH,"rb").read()).decode()}"
                   style="width:100%;max-width:220px;display:block;margin:0 auto;"/>
            </div>''',unsafe_allow_html=True)
        else:
            st.markdown(f'<div style="font-size:1.4rem;font-weight:900;color:{VERDE};text-align:center">🏋️ PUTÚ ACTIVO</div>',unsafe_allow_html=True)
        st.markdown("---")
        # Mostrar próximas clases en sidebar (sin link)
        clases_prox=db_query("SELECT titulo,fecha,hora FROM clases WHERE fecha>=? AND participante LIKE '%[CLASE]%' ORDER BY fecha,hora LIMIT 5",(str(hoy),))
        if not clases_prox.empty:
            st.markdown(f'<div style="background:{GRIS2};border-radius:8px;padding:8px 10px;margin-bottom:6px;">',unsafe_allow_html=True)
            st.markdown(f'<div style="color:{VERDE};font-weight:700;font-size:.85rem;margin-bottom:4px">📅 Próximas clases</div>',unsafe_allow_html=True)
            for _,cl in clases_prox.iterrows():
                st.markdown(f'<div style="font-size:.78rem;color:{GRIS_T};padding:2px 0"><b style="color:{BLANCO}">{cl["titulo"]}</b> · {fmt_fecha(cl["fecha"])} {str(cl["hora"])[:5]}</div>',unsafe_allow_html=True)
            st.markdown('</div>',unsafe_allow_html=True)
        modo_cli = st.radio("",["✅  Asistencia","🔍  Mi ficha"],
            index=["✅  Asistencia","🔍  Mi ficha"].index(
                st.session_state.pop("_modo_cli_override","✅  Asistencia")
                if st.session_state.get("_modo_cli_override") else "✅  Asistencia"))
        st.markdown("---")
        if st.button("← Inicio"):
            st.session_state.modo=""; st.session_state.rut_cliente=""; st.rerun()

    # ── ASISTENCIA (modo cliente) ──────────────────────────────────────────
    if modo_cli == "✅  Asistencia":
        if "rut_buf" not in st.session_state: st.session_state.rut_buf=""
        if "asist_ok" not in st.session_state: st.session_state.asist_ok=None
        if "mostrar_scan_qr" not in st.session_state: st.session_state.mostrar_scan_qr=False

        def procesar_rut_cliente(rut_raw, tipo_mov):
            rut_u=rut_raw.strip().upper()
            if not rut_u: return
            cli=db_query("SELECT * FROM clientes WHERE UPPER(rut)=?",(rut_u,))
            if cli.empty:
                st.session_state.asist_ok={"ok":False,"nombre":rut_u,"msg":"RUT no encontrado","tipo":tipo_mov}; return
            c=cli.iloc[0]
            if str(c.get("estado","")).capitalize()!="Activo":
                st.session_state.asist_ok={"ok":False,"nombre":c["nombre"],"rut":c["rut"],"msg":"Usuario bloqueado, consultar en administración","tipo":tipo_mov}; return
            ahora=datetime.now(); hora_str=ahora.strftime("%H:%M")
            ya=db_query("SELECT id FROM asistencia WHERE UPPER(rut)=? AND fecha=? AND tipo='ingreso' AND (hora_salida IS NULL OR hora_salida='')",(rut_u,str(hoy)))
            if tipo_mov=="ingreso":
                if not ya.empty:
                    st.session_state.asist_ok={"ok":False,"nombre":c["nombre"],"rut":c["rut"],"msg":"Ya tiene ingreso activo hoy","tipo":"ingreso"}; return
                db_exec("INSERT INTO asistencia (rut,nombre,fecha,hora,tipo,usuario) VALUES (?,?,?,?,?,?)",(c["rut"],c["nombre"],str(hoy),hora_str,"ingreso","kiosko"))
                st.session_state.asist_ok={"ok":True,"nombre":c["nombre"],"rut":c["rut"],"plan":c["tipo_plan"],"hora":hora_str,"tipo":"ingreso","emoji":"🏋️"}
            else:
                if ya.empty:
                    st.session_state.asist_ok={"ok":False,"nombre":c["nombre"],"rut":c["rut"],"msg":"No tiene ingreso activo hoy","tipo":"salida"}; return
                _cn=get_conn(); _cn.execute("UPDATE asistencia SET hora_salida=? WHERE id=?",(hora_str,int(ya.iloc[0]["id"]))); _cn.commit(); _cn.close()
                st.session_state.asist_ok={"ok":True,"nombre":c["nombre"],"rut":c["rut"],"plan":c["tipo_plan"],"hora":hora_str,"tipo":"salida","emoji":"👋"}

        st.markdown(f'<div class="section-header" style="font-size:1.1rem;margin-bottom:6px">✅ Asistencia</div>',unsafe_allow_html=True)
        col_tec,col_disp=st.columns([1,1])
        with col_tec:
            # ── Display en vivo del RUT siendo tecleado ──
            rut_vivo = st.session_state.rut_buf or "_ _ _ _ _ _ _"
            st.markdown(f'''<div style="font-size:2.2rem;font-weight:900;color:{VERDE};text-align:center;
                background:{GRIS2};border:3px solid {VERDE};border-radius:14px;
                padding:14px;margin-bottom:8px;letter-spacing:.15em;">{rut_vivo}</div>''',unsafe_allow_html=True)

            # ── Una sola barra RUT — teclado físico Y numérico táctil escriben aquí ──
            with st.form("f_cli_rut",clear_on_submit=True):
                rut_unico=st.text_input("🪪 RUT (teclado físico o táctil):",value="",
                    placeholder="12345678-9 ó 9876543-K",key="cli_rut_unico",
                    label_visibility="visible")
                ff1,ff2=st.columns(2)
                ok_ing=ff1.form_submit_button("✅ INGRESO",use_container_width=True)
                ok_sal=ff2.form_submit_button("🚪 SALIDA",use_container_width=True)
            if ok_ing and rut_unico.strip():
                procesar_rut_cliente(rut_unico,"ingreso"); st.session_state.rut_buf=""; st.rerun()
            elif ok_sal and rut_unico.strip():
                procesar_rut_cliente(rut_unico,"salida"); st.session_state.rut_buf=""; st.rerun()
            elif ok_ing or ok_sal:
                # Si el form se envió vacío, usar lo acumulado por el teclado táctil
                tipo_buf="ingreso" if ok_ing else "salida"
                if st.session_state.rut_buf.strip():
                    procesar_rut_cliente(st.session_state.rut_buf,tipo_buf)
                    st.session_state.rut_buf=""; st.rerun()

            # Teclado numérico táctil — añade al buffer (display crece en vivo)
            # Teclado grande — CSS extra para botones más grandes en kiosko
            st.markdown(f"""<style>
            div[data-testid="stHorizontalBlock"] button[kind="secondary"] {{
                font-size:2rem !important; font-weight:900 !important;
                padding:6px 4px !important; min-height:52px !important; line-height:1 !important;
            }}
            </style>""",unsafe_allow_html=True)
            nums=[["1","2","3"],["4","5","6"],["7","8","9"],["-","0","⌫"]]
            for fila in nums:
                cols=st.columns(3)
                for i,d in enumerate(fila):
                    if cols[i].button(d,key=f"cli_np_{d}",use_container_width=True):
                        if d=="⌫": st.session_state.rut_buf=st.session_state.rut_buf[:-1]
                        else: st.session_state.rut_buf+=d
                        st.session_state.asist_ok=None; st.rerun()
            ck1,ck2=st.columns(2)
            if ck1.button("K",key="cli_K",use_container_width=True):
                st.session_state.rut_buf+="K"; st.session_state.asist_ok=None; st.rerun()
            if ck2.button("🗑 Borrar",key="cli_bor",use_container_width=True):
                st.session_state.rut_buf=""; st.session_state.asist_ok=None; st.rerun()

            # ── Escanear QR con cámara ──
            if QR_SCAN_DISPONIBLE:
                if st.button("📷 Escanear mi QR", key="btn_scan_qr", use_container_width=True):
                    st.session_state.mostrar_scan_qr = not st.session_state.mostrar_scan_qr
                if st.session_state.mostrar_scan_qr:
                    foto_qr = st.camera_input("Apunta tu QR a la cámara", key="cam_qr_cliente", label_visibility="collapsed")
                    if foto_qr is not None:
                        rut_detectado = decodificar_qr(foto_qr.getvalue())
                        if rut_detectado:
                            st.markdown(f'<div class="success-box">✅ QR detectado: {rut_detectado}</div>',unsafe_allow_html=True)
                            cqr1,cqr2=st.columns(2)
                            if cqr1.button("✅ Marcar INGRESO",key="qr_ingreso",use_container_width=True):
                                procesar_rut_cliente(rut_detectado,"ingreso")
                                st.session_state.mostrar_scan_qr=False; st.rerun()
                            if cqr2.button("🚪 Marcar SALIDA",key="qr_salida",use_container_width=True):
                                procesar_rut_cliente(rut_detectado,"salida")
                                st.session_state.mostrar_scan_qr=False; st.rerun()
                        else:
                            st.markdown('<div class="alert-box">❌ No se detectó un QR válido. Intenta de nuevo.</div>',unsafe_allow_html=True)

        with col_disp:
            asist_hoy=db_query("SELECT * FROM asistencia WHERE fecha=? ORDER BY hora DESC",(str(hoy),))
            st.markdown(f'<div style="background:{GRIS2};border:1px solid {GRIS3};border-radius:12px;padding:10px;">',unsafe_allow_html=True)
            if not asist_hoy.empty and "tipo" in asist_hoy.columns:
                en_sala=asist_hoy[(asist_hoy["tipo"]=="ingreso")&(asist_hoy["hora_salida"].isna()|asist_hoy["hora_salida"].eq(""))]
            else: en_sala=asist_hoy
            st.markdown(f'<div style="color:{VERDE};font-weight:700;font-size:.95rem;margin-bottom:6px">🏟️ EN SALA — {len(en_sala)} persona(s)</div>',unsafe_allow_html=True)
            if not en_sala.empty:
                for _si,a in en_sala.iterrows():
                    _cn2,_ch2,_cs2,_cf2=st.columns([2.2,.9,1,1])
                    _cn2.markdown(f"<span style='font-weight:700;font-size:.95rem'>{a['nombre']}</span>",unsafe_allow_html=True)
                    _ch2.markdown(f"<span style='color:{VERDE};font-size:.85rem'>🏋️ {a['hora']}</span>",unsafe_allow_html=True)
                    if _cs2.button("🚪",key=f"cli_sal_s_{a.get('id',_si)}",use_container_width=True,help="Marcar salida"):
                        _hs=datetime.now().strftime("%H:%M")
                        _cs3=get_conn(); _cs3.execute("UPDATE asistencia SET hora_salida=? WHERE id=?",(_hs,int(a["id"]))); _cs3.commit(); _cs3.close()
                        st.rerun()
                    if _cf2.button("👤",key=f"cli_ficha_s_{a.get('id',_si)}",use_container_width=True,help="Ver ficha"):
                        _rut_sala=str(a.get("rut",""))
                        st.session_state.rut_cliente=_rut_sala.upper()
                        st.session_state["_modo_cli_override"]="🔍  Mi ficha"; st.rerun()
            st.markdown("</div>",unsafe_allow_html=True)

            # ── Mensaje de bienvenida/salida — debajo de "En sala" ──
            if st.session_state.asist_ok:
                ao=st.session_state.asist_ok
                col_r=VERDE if (ao["ok"] and ao.get("tipo")=="ingreso") else AZUL if (ao["ok"] and ao.get("tipo")=="salida") else ROJO
                emoji=ao.get("emoji","✅") if ao["ok"] else "❌"
                msg_r="✅ ¡Ingreso registrado! 💪" if (ao.get("tipo")=="ingreso" and ao["ok"]) else ("🚪 ¡Salida registrada! 👋" if ao["ok"] else ao.get("msg","Error"))
                st.markdown(f'''<div style="background:{col_r}22;border:2px solid {col_r};border-radius:12px;
                    padding:12px;text-align:center;margin-top:8px;">
                  <div style="font-size:1.6rem">{emoji}</div>
                  <div style="font-size:1.2rem;font-weight:900;color:{col_r};margin:4px 0">{ao["nombre"]}</div>
                  <div style="color:{GRIS_T};font-size:.85rem">{ao.get("plan","")} · {ao.get("hora","")}</div>
                  <div style="color:{col_r};font-weight:700;margin-top:4px">{msg_r}</div>
                </div>''',unsafe_allow_html=True)
                # Ver ficha directo después de marcar
                if ao["ok"]:
                    _rut_marcado=ao.get("rut","")
                    if not _rut_marcado:
                        _cli_buscado=db_query("SELECT rut FROM clientes WHERE nombre=?",(ao["nombre"],))
                        if not _cli_buscado.empty: _rut_marcado=_cli_buscado.iloc[0]["rut"]
                    if _rut_marcado and st.button("👤 Ver mi ficha",key="cli_ver_ficha_post",use_container_width=True):
                        st.session_state.rut_cliente=_rut_marcado.upper()
                        st.session_state["_modo_cli_override"]="🔍  Mi ficha"
                        st.rerun()

    # ── MI FICHA (modo cliente) ───────────────────────────────────────────
    elif modo_cli == "🔍  Mi ficha":
        st.markdown(f'<div class="section-header">🔍 Consulta de Ficha</div>',unsafe_allow_html=True)
        if st.button("← Volver a Asistencia",key="cli_ficha_volver"):
            st.session_state["_modo_cli_override"]="✅  Asistencia"
            st.session_state.rut_cliente=""; st.rerun()
        with st.form("f_consulta_rut",clear_on_submit=False):
            busq_c=st.text_input("Ingresa tu nombre o RUT:",
                placeholder="Ej: Juan Pérez  ó  12345678-9",key="rut_consulta_cli")
            ok_c=st.form_submit_button("🔍 Buscar",use_container_width=True)
        if ok_c and busq_c.strip():
            st.session_state.rut_cliente=busq_c.strip()
        if st.session_state.rut_cliente:
            _bq=st.session_state.rut_cliente.upper()
            rows=db_query("SELECT * FROM clientes WHERE UPPER(rut)=? OR UPPER(nombre) LIKE ?",
                (_bq, f"%{_bq}%"))
            if len(rows)>1:
                _opts=[f"{r['nombre']} — {r['rut']}" for _,r in rows.iterrows()]
                _sel=st.selectbox("Selecciona:",_opts,key="cli_sel_ficha")
                rows=rows.iloc[[_opts.index(_sel)]]
            if rows.empty:
                st.markdown('<div class="alert-box">❌ RUT no encontrado.</div>',unsafe_allow_html=True)
            else:
                rc=rows.iloc[0].to_dict()
                es_a=str(rc.get("estado","")).capitalize()=="Activo"
                col_e2=VERDE if es_a else ROJO
                try:
                    fn_rc=date.fromisoformat(str(rc.get("fecha_nacimiento",""))[:10])
                    edad_rc=int((hoy-fn_rc).days/365.25)
                except: edad_rc=rc.get("edad","?")
                dias_ent=" ".join(filter(None,[str(rc.get("lunes","")),str(rc.get("martes","")),
                    str(rc.get("miercoles","")),str(rc.get("jueves","")),
                    str(rc.get("viernes","")),str(rc.get("sabado",""))]))
                st.markdown(f'''<div class="card" style="border-left:5px solid {col_e2};">
                  <div style="font-size:1.6rem;font-weight:900;color:{VERDE}">{rc.get("nombre","")}</div>
                  <div style="color:{GRIS_T};font-size:.9rem;margin-top:4px">
                    RUT: <b>{rc.get("rut","")}</b> · {rc.get("sexo","")} · {edad_rc} años
                  </div>
                  <div style="display:flex;gap:8px;flex-wrap:wrap;margin-top:8px;font-size:.88rem;color:{GRIS_T};">
                    <span>💳 <b>{rc.get("tipo_plan","")}</b></span>
                    <span>📅 Vence: <b>{fmt_fecha(rc.get("fecha_vencimiento",""))}</b></span>
                    <span>⏰ {rc.get("horario","")}</span>
                    <span style="color:{col_e2};font-weight:700">{rc.get("estado","")}</span>
                  </div>
                  <div style="color:{GRIS_T};font-size:.85rem;margin-top:4px">
                    🗓 Días: <b>{dias_ent}</b> · 🎯 {rc.get("objetivo","")} · 📊 {rc.get("nivel","")}
                  </div>
                </div>''',unsafe_allow_html=True)
                # Mostrar rutina vigente desde BD
                _rut_ki=db_query("SELECT * FROM rutinas WHERE cliente_rut=? AND activa=1 ORDER BY id DESC LIMIT 1",(rc["rut"],))
                if not _rut_ki.empty:
                    _rutki=_rut_ki.iloc[0].to_dict()
                    st.markdown(f"<b style='color:{VERDE};font-size:1rem'>💪 {_rutki['nombre']}</b>",unsafe_allow_html=True)
                    _ejs_ki=db_query("""SELECT re.*,e.nombre,e.url_imagen,e.musculo_primario
                        FROM rutina_ejercicios re JOIN ejercicios e ON e.id=re.ejercicio_id
                        WHERE re.rutina_id=? ORDER BY re.dia_semana,re.orden""",(int(_rutki["id"]),))
                    if not _ejs_ki.empty:
                        _dias_ki=[d for d in ["Día 1","Día 2","Día 3","Día 4","Día 5","Día 6"] if d in _ejs_ki["dia_semana"].values]
                        _es_fem_ki=str(rc.get("sexo","")).lower()=="femenino"
                        _col_dia_ki="#E91E8C" if _es_fem_ki else "#3A9BD5"
                        _hcols_ki=st.columns(len(_dias_ki))
                        for _ic_ki,_hc_ki in enumerate(_hcols_ki):
                            _hc_ki.markdown(f'<div style="background:{_col_dia_ki};color:#fff;text-align:center;font-weight:700;font-size:.78rem;padding:5px 2px;border-radius:6px 6px 0 0">{_dias_ki[_ic_ki]}</div>',unsafe_allow_html=True)
                        _bcols_ki=st.columns(len(_dias_ki))
                        for _ic2_ki,(_cb_ki,_dfl_ki) in enumerate(zip(_bcols_ki,_dias_ki)):
                            _ddf_ki=_ejs_ki[_ejs_ki["dia_semana"]==_dfl_ki]
                            with _cb_ki:
                                for _ii_ki,(_,_ef_ki) in enumerate(_ddf_ki.iterrows(),1):
                                    _efd_ki=_ef_ki.to_dict()
                                    with st.container(border=True):
                                        _iurl_ki=str(_efd_ki.get("url_imagen","")).strip()
                                        if _iurl_ki and _iurl_ki!="nan":
                                            try: st.image(_iurl_ki,use_container_width=True)
                                            except: st.markdown('<div style="font-size:1.5rem;text-align:center">🏋️</div>',unsafe_allow_html=True)
                                        else: st.markdown('<div style="font-size:1.5rem;text-align:center">🏋️</div>',unsafe_allow_html=True)
                                        st.markdown(f'<div style="font-size:.72rem;font-weight:700;color:#fff;text-align:center;line-height:1.2">{_ii_ki}. {str(_efd_ki.get("nombre",""))[:22]}</div>',unsafe_allow_html=True)
                                        _sr_ki=sv(_efd_ki,"series","—"); _rp_ki=sv(_efd_ki,"repeticiones","—")
                                        st.markdown(f'<div style="font-size:.68rem;color:{_col_dia_ki};text-align:center">{_sr_ki}×{_rp_ki}</div>',unsafe_allow_html=True)
                    else:
                        st.markdown('<div class="info-box">Esta rutina aún no tiene ejercicios.</div>',unsafe_allow_html=True)
                else:
                    st.markdown('<div class="info-box">Sin rutina activa. Consulta con tu entrenador.</div>',unsafe_allow_html=True)
                # Últimas asistencias
                ult_as=db_query("SELECT fecha,hora,tipo FROM asistencia WHERE UPPER(rut)=? ORDER BY fecha DESC,hora DESC LIMIT 10",(st.session_state.rut_cliente,))
                if not ult_as.empty:
                    st.markdown(f'<div style="font-weight:700;color:{VERDE};margin:12px 0 6px">✅ Últimas asistencias</div>',unsafe_allow_html=True)
                    st.dataframe(ult_as.rename(columns={"fecha":"Fecha","hora":"Hora","tipo":"Tipo"}),use_container_width=True,height=220)

    st.stop()

# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    # Logo centrado, pegado arriba, sin padding extra
    if LOGO_PATH:
        st.markdown(f'''<div style="text-align:center;padding:0;margin:-16px -8px 2px -8px;">
          <img src="data:image/png;base64,{__import__("base64").b64encode(open(LOGO_PATH,"rb").read()).decode()}"
               style="width:90%;max-width:180px;display:block;margin:0 auto;"/>
        </div>''',unsafe_allow_html=True)
    else:
        st.markdown(f'<div style="font-size:1.3rem;font-weight:900;color:{VERDE};padding:2px 4px;">🏋️ PUTÚ ACTIVO</div>',unsafe_allow_html=True)
    st.markdown(f'<div style="background:{GRIS3};border-radius:5px;padding:2px 7px;margin:2px 0;font-size:.68rem;">👤 <b style="color:{VERDE}">{st.session_state.nombre_u}</b> <span style="color:#555">· {st.session_state.rol}</span></div>',unsafe_allow_html=True)
    st.markdown('<hr style="margin:3px 0;border-color:#2E2E2E">',unsafe_allow_html=True)
    _opciones=[
        "🏠 Dashboard","👥 Clientes",
        "💳 Pagos y Renovaciones",
        "✅ Asistencia","🏃 Clases & Talleres",
        "🛍 Venta Productos","💪 Ejercicios","📋 Rutinas","📊 Reportes","⚙️ Base de Datos"]
    # Filtrar opciones según permisos del usuario PRIMERO
    _permisos_pagina = {
        "🏠 Dashboard":              "dashboard",
        "👥 Clientes":               "clientes",
        "💳 Pagos y Renovaciones":   "pagos",
        "✅ Asistencia":             "asistencia",
        "🏃 Clases & Talleres":      "clases",
        "🛍 Venta Productos":        "reportes",
        "💪 Ejercicios":             "clientes",
        "📋 Rutinas":                "clientes",
        "📊 Reportes":               "reportes",
        "⚙️ Base de Datos":          "db",
    }
    _opciones=[o for o in _opciones if tiene_permiso(_permisos_pagina.get(o,"dashboard"))]
    # Calcular índice DESPUÉS de filtrar
    _idx_default=0
    if st.session_state.get("_goto") and st.session_state["_goto"] in _opciones:
        _idx_default=_opciones.index(st.session_state["_goto"])
        st.session_state["_goto"]=None
    pagina=st.radio("",_opciones,index=min(_idx_default,len(_opciones)-1))
    st.markdown("---")
    mins=max(0,int((SESSION_H*3600-(time.time()-st.session_state.login_time))/60))
    st.markdown(f'<div style="color:#555;font-size:.65rem;margin-bottom:4px">Sesión: ~{mins} min</div>',unsafe_allow_html=True)
    _sb1,_sb2=st.columns(2)
    if _sb1.button("🚪 Salir",key="btn_cerrar_sesion",use_container_width=True):
        st.session_state.logueado=False; st.session_state.login_time=None
        st.session_state.modo=""; st.rerun()
    if _sb2.button("🏠 Inicio",key="btn_inicio",use_container_width=True):
        st.session_state.logueado=False; st.session_state.login_time=None
        st.session_state.modo=""; st.rerun()

# ─────────────────────────────────────────────────────────────────────────────
# DATOS GLOBALES
# ─────────────────────────────────────────────────────────────────────────────
@st.cache_data(ttl=30)
def get_clientes():
    df=db_query("SELECT * FROM clientes ORDER BY nombre")
    df["N°"]=range(1,len(df)+1)
    df["es_pase"]=df["tipo_plan"].astype(str).str.upper().str.strip()=="PASE DIARIO"
    # Calcular edad, día y mes desde fecha_nacimiento
    def calc_edad(fn):
        try: return int((date.today()-date.fromisoformat(str(fn)[:10])).days/365.25)
        except: return None
    def get_dia_nac(fn):
        try: return date.fromisoformat(str(fn)[:10]).day
        except: return None
    def get_mes_nac(fn):
        try: return date.fromisoformat(str(fn)[:10]).month
        except: return None
    df["edad"]     = df["fecha_nacimiento"].apply(calc_edad)
    df["dia_nac"]  = df["fecha_nacimiento"].apply(get_dia_nac)
    df["mes_nac"]  = df["fecha_nacimiento"].apply(get_mes_nac)
    return df

df_all  = get_clientes()
df_cli  = df_all[~df_all["es_pase"]].copy()
df_act  = df_cli[df_cli["estado"].astype(str).str.capitalize()=="Activo"].copy()
df_inac = df_cli[df_cli["estado"].astype(str).str.capitalize()!="Activo"].copy()

# ════════════════════════════════════════════════════════════════════════════
# 🏠 DASHBOARD
# ════════════════════════════════════════════════════════════════════════════
if pagina=="🏠 Dashboard":
    st.markdown('<div class="section-header">🏠 Dashboard General</div>',unsafe_allow_html=True)
    tot_a=len(df_act); tot_i=len(df_inac); tot=len(df_cli)
    masc=int((df_act["sexo"].str.lower()=="masculino").sum())
    fem=int((df_act["sexo"].str.lower()=="femenino").sum())
    vencen=int(df_act["fecha_vencimiento"].apply(_vence_pronto).sum())
    # Ticket promedio — solo planes Quincenal y Mensual
    _planes_ticket=["Quincenal","Mensual","PM","AM"]
    _df_tick=df_act[df_act["tipo_plan"].str.upper().str.strip().isin([p.upper() for p in _planes_ticket])]
    vp_vals=pd.to_numeric(_df_tick["valor_plan"],errors="coerce").dropna()
    ticket_prom=int(vp_vals.mean()) if len(vp_vals)>0 else 0
    # Tasa de abandono (inactivos / total *100)
    fm_hoy=hoy.strftime("%Y-%m")
    # Fila 1: métricas clave (sin tasa abandono ni ingresos/egresos)
    mc1,mc2,mc3,mc4=st.columns(4)
    mc1.metric("✅ Activos",tot_a); mc2.metric("🔴 Inactivos",tot_i)
    mc3.metric("💰 Ticket prom.",f"${ticket_prom:,}")
    mc4.metric("⚠️ Vencen pronto",vencen)
    st.markdown("")
    # Fila gráficos
    ca,cb=st.columns(2)
    with ca:
        sc=df_act["sexo"].str.capitalize().value_counts().reset_index(); sc.columns=["Sexo","N"]
        fig_d=go.Figure(go.Pie(labels=sc["Sexo"],values=sc["N"],hole=.55,
            marker_colors=[VERDE,AZUL,NARANJA],textfont=dict(color=BLANCO,size=12),
            hovertemplate="%{label}: <b>%{value}</b> (%{percent})<extra></extra>"))
        fig_d.add_annotation(text=f"<b>{tot_a}</b><br>activos",x=.5,y=.5,font=dict(size=18,color=VERDE),showarrow=False)
        fig_d.update_layout(title="Activos por sexo",paper_bgcolor="rgba(0,0,0,0)",plot_bgcolor="rgba(0,0,0,0)",
            font=dict(family="Inter",color=BLANCO,size=12),
            legend=dict(bgcolor="rgba(0,0,0,0)",orientation="h",x=.5,xanchor="center",y=-.08),
            margin=dict(l=16,r=16,t=36,b=36),height=260)
        st.plotly_chart(fig_d,use_container_width=True)
    with cb:
        # Gráfico finanzas últimos 6 meses
        _meses=[]; _ing=[]; _egr=[]
        # Finanzas 6 meses — una sola query en vez de 12
        _fin_df=db_query("SELECT strftime('%Y-%m',fecha) as mes,SUM(monto) as tot FROM pagos WHERE fecha>=date('now','-6 months') GROUP BY mes")
        _egr_df=db_query("SELECT strftime('%Y-%m',fecha) as mes,SUM(monto) as tot FROM egresos WHERE fecha>=date('now','-6 months') GROUP BY mes")
        for _mi in range(5,-1,-1):
            _md=date(hoy.year,hoy.month,1)
            _mn=(_md.month-1-_mi)%12+1; _yn=_md.year if _md.month-_mi>0 else _md.year-1
            _fmi=f"{_yn}-{_mn:02d}"
            _meses.append(_fmi)
            _ing_row=_fin_df[_fin_df["mes"]==_fmi]["tot"].values
            _egr_row=_egr_df[_egr_df["mes"]==_fmi]["tot"].values
            _ing.append(int(_ing_row[0]) if len(_ing_row)>0 else 0)
            _egr.append(int(_egr_row[0]) if len(_egr_row)>0 else 0)
        fig_fin=go.Figure()
        fig_fin.add_trace(go.Bar(name="Ingresos",x=_meses,y=_ing,marker_color=VERDE))
        fig_fin.add_trace(go.Bar(name="Egresos",x=_meses,y=_egr,marker_color=ROJO))
        fig_fin.update_layout(title="Finanzas últimos 6 meses",paper_bgcolor="rgba(0,0,0,0)",plot_bgcolor="rgba(0,0,0,0)",font=dict(family="Inter",color=BLANCO,size=12),title_font=dict(color=BLANCO,size=14),margin=dict(l=16,r=16,t=36,b=40),xaxis=dict(gridcolor=GRIS3),yaxis=dict(gridcolor=GRIS3),barmode="group",height=260,legend=dict(bgcolor="rgba(0,0,0,0)",orientation="h",x=.5,xanchor="center",y=-.18))
        st.plotly_chart(fig_fin,use_container_width=True)
    # Mini-reporte rutinas
    _rut_stats=db_query("SELECT SUM(CASE WHEN activa=1 THEN 1 ELSE 0 END) as activas,COUNT(*) as total FROM rutinas WHERE cliente_rut IS NOT NULL").iloc[0]
    _tot_rut=int(_rut_stats["activas"] or 0); _sin_rut=max(0,len(df_act)-_tot_rut)
    _cr1,_cr2,_cr3=st.columns(3)
    _cr1.metric("📋 Con rutina activa",_tot_rut)
    _cr2.metric("⚠️ Sin rutina",_sin_rut)
    _cr3.metric("📚 Total rutinas guardadas",int(_rut_stats["total"] or 0))
    st.markdown("")
    # Fila inferior: vencimientos + cumpleaños
    cv,cc=st.columns(2)
    with cv:
        st.markdown(f"<b style='color:{VERDE}'>⚠️ Próximos vencimientos</b>",unsafe_allow_html=True)
        dv=df_act[df_act["fecha_vencimiento"].notna()].copy()
        dv["días"]=dv["fecha_vencimiento"].apply(dias_para_vencer)
        dv=dv[dv["días"].notna()&(dv["días"]>=0)].sort_values("días").head(5)
        if dv.empty:
            st.markdown(f'<div class="info-box">Sin vencimientos próximos.</div>',unsafe_allow_html=True)
        for _,r in dv.iterrows():
            d=int(r["días"]); c=ROJO if d<=5 else NARANJA if d<=15 else GRIS_T
            mv2=str(r.get("mensaje_vencimiento","")) or msg_vencimiento(r["nombre"],r["fecha_vencimiento"])
            uv2=wa_url(r["celular"],mv2)
            st.markdown(f'''<div style="background:{GRIS2};border-left:4px solid {c};border-radius:8px;
                padding:7px 12px;margin:3px 0;display:flex;justify-content:space-between;align-items:center;">
              <div style="font-size:.88rem"><b>{r["nombre"]}</b>
              <span style="color:{GRIS_T}"> · {fmt_fecha(r["fecha_vencimiento"])}</span></div>
              <div style="display:flex;gap:6px;align-items:center;">
                <span style="color:{c};font-weight:900;font-size:.95rem">{d}d</span>
                <a href="{uv2}" target="_blank" style="background:#25D366;color:white;padding:3px 9px;border-radius:6px;text-decoration:none;font-size:.8rem;">WA</a>
              </div></div>''',unsafe_allow_html=True)
    with cc:
        st.markdown(f"<b style='color:{VERDE}'>🎂 Próximos cumpleaños</b>",unsafe_allow_html=True)
        def _dias_cumple(fn):
            try:
                fd=date.fromisoformat(str(fn)[:10])
                prox=date(hoy.year,fd.month,fd.day)
                if prox<hoy: prox=date(hoy.year+1,fd.month,fd.day)
                return (prox-hoy).days
            except: return 999
        df_cb2=df_act[df_act["fecha_nacimiento"].notna()].copy()
        df_cb2["_dias_c"]=df_cb2["fecha_nacimiento"].apply(_dias_cumple)
        df_cb2=df_cb2.sort_values("_dias_c").head(5)
        if df_cb2.empty:
            st.markdown(f'<div class="info-box">Sin datos de cumpleaños.</div>',unsafe_allow_html=True)
        for _,r in df_cb2.iterrows():
            d_c=int(r["_dias_c"])
            mc2_txt=str(r.get("mensaje_cumpleanos","")) or msg_cumpleanos(r["nombre"])
            uc2=wa_url(r["celular"],mc2_txt)
            badge="🎂 HOY" if d_c==0 else f"en {d_c}d"
            col_b=VERDE if d_c==0 else AZUL if d_c<=7 else GRIS_T
            try: fn_str=date.fromisoformat(fmt_fecha(r["fecha_nacimiento"])).strftime("%d/%m")
            except: fn_str=""
            st.markdown(f'''<div style="background:{GRIS2};border-left:4px solid {col_b};border-radius:8px;
                padding:7px 12px;margin:3px 0;display:flex;justify-content:space-between;align-items:center;">
              <div style="font-size:.88rem"><b>{r["nombre"]}</b>
              <span style="color:{GRIS_T}"> · {fn_str}</span></div>
              <div style="display:flex;gap:6px;align-items:center;">
                <span style="color:{col_b};font-weight:900;font-size:.9rem">{badge}</span>
                <a href="{uc2}" target="_blank" style="background:#25D366;color:white;padding:3px 9px;border-radius:6px;text-decoration:none;font-size:.8rem;">WA</a>
              </div></div>''',unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════════════════════
# 👥 CLIENTES — lista + ficha
# ════════════════════════════════════════════════════════════════════════════
elif pagina=="👥 Clientes":

    def ficha_cliente(rut):
        rows = db_query("SELECT * FROM clientes WHERE rut=?", (rut,))
        if rows.empty: st.session_state.ver_rut=None; st.rerun()
        r   = rows.iloc[0].to_dict()
        es_act  = sv(r,"estado").capitalize() == "Activo"
        es_fem  = sv(r,"sexo").lower() == "femenino"
        col_borde = "#E91E8C" if es_fem and es_act else VERDE if es_act else ROJO
        col_e   = VERDE if es_act else ROJO  # badge estado siempre verde/rojo
        cel     = sv(r,"celular")
        mc      = sv(r,"mensaje_cumpleanos") or msg_cumpleanos(sv(r,"nombre"))
        mv      = sv(r,"mensaje_vencimiento") or msg_vencimiento(sv(r,"nombre"),sv(r,"fecha_vencimiento"))
        uc      = wa_url(cel, mc)
        uv      = wa_url(cel, mv)

        # ── Botón volver ──
        _tabs_names=["📋 Datos","💪 Rutina","🥗 Nutrición","📏 Evaluación","💳 Pagos","📲 QR","📄 Documentos"]
        if st.button("← Volver", key="vt"): st.session_state.ver_rut=None; st.rerun()
        col_foto, col_info, col_acc = st.columns([1,3,2])
        with col_foto:
                fp = sv(r,"foto_path")
                if fp and os.path.exists(fp):
                    st.image(fp, width=110)
                    fup = st.file_uploader("📷",type=["jpg","jpeg","png"],key=f"fu_{rut}",
                        label_visibility="collapsed",help="Cambiar foto")
                else:
                    _qr_fp=generar_qr_b64(f"PUTU|{rut}|{sv(r,'nombre')}")
                    st.markdown(f'<img src="data:image/png;base64,{_qr_fp}" style="width:110px;height:110px;border-radius:8px;border:1px solid {GRIS3}">',unsafe_allow_html=True)
                    st.markdown(f'<div style="font-size:.58rem;color:{GRIS_T};text-align:center;margin-top:2px">QR asistencia</div>',unsafe_allow_html=True)
                    fup = st.file_uploader("📷 Subir foto",type=["jpg","jpeg","png"],key=f"fu_{rut}",
                        label_visibility="collapsed",help="Subir foto del cliente")
                if fup:
                    fp2 = os.path.join(FOTOS_DIR,f"{rut.replace('-','_')}.{fup.name.split('.')[-1]}")
                    with open(fp2,"wb") as f2: f2.write(fup.getbuffer())
                    conn2=get_conn(); conn2.execute("UPDATE clientes SET foto_path=? WHERE rut=?",(fp2,rut)); conn2.commit(); conn2.close()
                    st.cache_data.clear(); st.rerun()
        with col_info:
                st.markdown(f'''<div style="border-left:4px solid {col_borde};padding-left:12px;">
                  <div style="font-size:1.4rem;font-weight:900;color:{VERDE};line-height:1.1">{sv(r,"nombre")}</div>
                  <div style="color:{GRIS_T};font-size:.88rem;margin-top:3px">
                    <b>{sv(r,"rut")}</b> · {sv(r,"sexo")} · {sv(r,"edad")} años · 🎂 {sv(r,"mes_cumpleanos")}
                  </div>
                  <div style="color:{GRIS_T};font-size:.85rem;margin-top:2px">📍 {sv(r,"direccion")}</div>
                  <div style="color:{GRIS_T};font-size:.85rem">📱 {fmt_cel(cel)} · ✉️ {sv(r,"email")}</div>
                  <div style="margin-top:6px;display:flex;gap:8px;flex-wrap:wrap;align-items:center;">
                    <span style="background:{col_e}22;color:{col_e};border:1px solid {col_e};
                      border-radius:20px;padding:2px 12px;font-size:.82rem;font-weight:700">{sv(r,"estado")}</span>
                    <span style="color:{GRIS_T};font-size:.82rem">💳 {sv(r,"tipo_plan")} · {sv(r,"frecuencia")} · {sv(r,"horario")} · 🔄 {sv(r,"periodo_vencimiento")}</span>
                  </div>
                  <div style="color:{GRIS_T};font-size:.82rem;margin-top:2px">
                    📅 {("🔴 VENCIDO" if dias_para_vencer(sv(r,"fecha_vencimiento")) is not None and dias_para_vencer(sv(r,"fecha_vencimiento"))<0 else "Vence")}: <b style="color:{"#E85050" if dias_para_vencer(sv(r,"fecha_vencimiento")) is not None and dias_para_vencer(sv(r,"fecha_vencimiento"))<0 else VERDE}">{fmt_fecha(sv(r,"fecha_vencimiento"))}</b> · 💵 <b>${int(float(r.get("valor_plan") or 0)):,}</b>
                  </div>
                </div>''', unsafe_allow_html=True)
        with col_acc:
                a1,a2 = st.columns(2)
                if a1.button("💳 Pagar", key="go_pago", use_container_width=True):
                    st.session_state.pago_rut=rut; st.session_state._goto="💳 Pagos y Renovaciones"
                    st.session_state.ver_rut=None; st.rerun()
                est_actual=sv(r,"estado").capitalize()
                lbl_blk="🔓 Activar" if est_actual=="Inactivo" else "🔒 Bloquear"
                if a2.button(lbl_blk, key="go_blk", use_container_width=True):
                    nv="Activo" if est_actual=="Inactivo" else "Inactivo"
                    conn2=get_conn(); conn2.execute("UPDATE clientes SET estado=?,modificado=? WHERE rut=?",(nv,datetime.now().isoformat(),rut)); conn2.commit(); conn2.close()
                    st.cache_data.clear(); st.rerun()
                if a1.button("🎂 WA Cumple", key="go_wac", use_container_width=True):
                    st.markdown(f'<a href="{uc}" target="_blank">Abrir WhatsApp</a>',unsafe_allow_html=True)
                if a2.button("⏰ WA Vence", key="go_wav", use_container_width=True):
                    st.markdown(f'<a href="{uv}" target="_blank" style="color:{VERDE};font-weight:700;">📲 Enviar aviso vencimiento</a>',unsafe_allow_html=True)
                if a1.button("🖨️ Imprimir", key="go_print", use_container_width=True):
                    st.session_state["show_print_"+rut]=True
                if st.session_state.get("show_print_"+rut):
                    _dev_p=db_query("SELECT * FROM evaluaciones WHERE rut=? ORDER BY fecha DESC LIMIT 1",(rut,))
                    _dpg_p=db_query("SELECT * FROM pagos WHERE rut=? ORDER BY fecha DESC LIMIT 5",(rut,))
                    _ev_str=""
                    if not _dev_p.empty:
                        _ep=_dev_p.iloc[0]
                        _ev_str=f"<tr><td>Peso</td><td>{_ep.get('peso','')} kg</td><td>IMC</td><td>{_ep.get('imc','')}</td></tr><tr><td>Grasa %</td><td>{_ep.get('grasa_pct','')}</td><td>Masa Musc.</td><td>{_ep.get('masa_musc','')}</td></tr>"
                    _pg_str="".join([f"<tr><td>{fmt_fecha(str(pg.get('fecha','')))}</td><td>${int(float(pg.get('monto',0))):,}</td><td>{pg.get('concepto','')}</td><td>{pg.get('medio_pago','')}</td></tr>" for _,pg in _dpg_p.iterrows()]) if not _dpg_p.empty else "<tr><td colspan='4'>Sin pagos</td></tr>"
                    _qr_b64_p=generar_qr_b64(f"PUTU|{rut}|{sv(r,'nombre')}")
                    _html_print=f"""<!DOCTYPE html><html><head><meta charset='utf-8'><title>Ficha {sv(r,'nombre')}</title>
                    <style>@page{{size:letter;margin:1.5cm}}body{{font-family:Arial;font-size:10pt;color:#111}}
                    h2{{color:#6DBE45;border-bottom:2px solid #6DBE45;padding-bottom:4px}}h3{{color:#333;margin-top:14px}}
                    table{{width:100%;border-collapse:collapse;margin:6px 0}}th{{background:#6DBE45;color:black;padding:5px 8px;text-align:left;font-size:9pt}}
                    td{{padding:4px 8px;border-bottom:1px solid #eee;font-size:9pt}}.qr{{text-align:center;margin-top:10px}}</style></head><body>
                    <h2>PUTÚ ACTIVO — Ficha de Cliente</h2>
                    <h3>📋 Datos Personales</h3>
                    <table><tr><th>Nombre</th><td><b>{sv(r,'nombre')}</b></td><th>RUT</th><td>{rut}</td></tr>
                    <tr><th>Plan</th><td>{sv(r,'tipo_plan')}</td><th>Vencimiento</th><td>{fmt_fecha(sv(r,'fecha_vencimiento'))}</td></tr></table>
                    <h3>📏 Última Evaluación</h3>
                    <table><tr><th>Campo</th><th>Valor</th><th>Campo</th><th>Valor</th></tr>{_ev_str if _ev_str else "<tr><td colspan='4'>Sin evaluaciones</td></tr>"}</table>
                    <h3>💳 Últimos Pagos</h3>
                    <table><tr><th>Fecha</th><th>Monto</th><th>Concepto</th><th>Medio</th></tr>{_pg_str}</table>
                    <div class='qr'><h3>📲 QR</h3><img src='data:image/png;base64,{_qr_b64_p}' width='120'></div>
                    </body></html>"""
                    # Generar PDF con ReportLab — diseño mejorado
                    if REPORTLAB_OK:
                        import urllib.request as _ur_fi, re as _re_fi
                        from PIL import Image as _PIL_fi
                        _pb_fi=io.BytesIO()
                        _pd_fi=SimpleDocTemplate(_pb_fi,pagesize=A4,
                            leftMargin=1.5*cm,rightMargin=1.5*cm,
                            topMargin=1.2*cm,bottomMargin=1.2*cm)
                        _AW=A4[0]-3*cm; _st_fi=[]

                        # Color según sexo
                        _es_fem_fi=sv(r,"sexo").lower()=="femenino"
                        _COL=rl_colors.HexColor("#E91E8C") if _es_fem_fi else rl_colors.HexColor("#3A9BD5")
                        _COL_LT=rl_colors.HexColor("#F5F5F5")  # gris claro neutro

                        # Estilos
                        _sT=lambda n,**k: ParagraphStyle(n,**k)
                        _stnom=_sT("fn",fontName="Helvetica-Bold",fontSize=16,textColor=_COL,leading=20)
                        _stsub=_sT("fs",fontName="Helvetica",fontSize=9,textColor=rl_colors.HexColor("#555"),leading=13)
                        _stsec=_sT("fsc",fontName="Helvetica-Bold",fontSize=10,textColor=rl_colors.HexColor("#6DBE45"),leading=14)
                        _stdat=_sT("fd",fontName="Helvetica",fontSize=9,textColor=rl_colors.HexColor("#222"),leading=12)
                        _stfoot=_sT("ff",fontName="Helvetica-Oblique",fontSize=7.5,textColor=rl_colors.HexColor("#888"),alignment=TA_CENTER)
                        _stlogo=_sT("fl",fontName="Helvetica-Bold",fontSize=13,textColor=rl_colors.HexColor("#6DBE45"))
                        _stgen=_sT("fg",fontName="Helvetica-Bold",fontSize=9,textColor=rl_colors.white,alignment=TA_CENTER)

                        # ── ENCABEZADO: Logo + título + QR ──────────────────
                        _logo_fi=None
                        try:
                            # Preferir logo con fondo negro si existe
                            _lp_fi_neg=os.path.join(BASE_DIR,"LOGO_PUTÚ_ACTIVO_Fondo_negro.jpg")
                            _lp_fi=_lp_fi_neg if os.path.exists(_lp_fi_neg) else os.path.join(BASE_DIR,"logo.png")
                            if os.path.exists(_lp_fi):
                                _im_fi=_PIL_fi.open(_lp_fi); _iw_fi,_ih_fi=_im_fi.size
                                _ratio_fi=1.6*cm/_ih_fi
                                _logo_fi=RLImage(_lp_fi,width=_iw_fi*_ratio_fi,height=1.6*cm)
                        except: pass

                        _qr_fi_b64=generar_qr_b64(f"PUTU|{rut}|{sv(r,'nombre')}")
                        _qr_fi_img=None
                        try:
                            _qr_fi_bytes=__import__("base64").b64decode(_qr_fi_b64)
                            _qr_fi_img=RLImage(io.BytesIO(_qr_fi_bytes),width=1.8*cm,height=1.8*cm)
                        except: pass

                        _hdr_left=[_logo_fi or Paragraph("PUTÚ ACTIVO",_stlogo),
                            Spacer(1,0.1*cm),
                            Paragraph("PUTÚ ACTIVO — Ficha de Cliente",_stlogo),
                            Paragraph(f"Generado: {fmt_fecha(str(hoy))}",_stsub)]
                        _hdr_tbl=Table([[_hdr_left,_qr_fi_img or ""]],
                            colWidths=[_AW-2.2*cm,2.2*cm])
                        _hdr_tbl.setStyle(TableStyle([
                            ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
                            ("ALIGN",(1,0),(1,0),"RIGHT"),
                            ("LEFTPADDING",(0,0),(-1,-1),0),
                            ("RIGHTPADDING",(0,0),(-1,-1),0),
                            ("TOPPADDING",(0,0),(-1,-1),0),
                            ("BOTTOMPADDING",(0,0),(-1,-1),0),
                        ]))
                        _st_fi.append(_hdr_tbl)
                        # Barra de color
                        _bar_tbl=Table([[""]],colWidths=[_AW])
                        _bar_tbl.setStyle(TableStyle([
                            ("BACKGROUND",(0,0),(-1,-1),rl_colors.HexColor("#6DBE45")),
                            ("TOPPADDING",(0,0),(-1,-1),3),
                            ("BOTTOMPADDING",(0,0),(-1,-1),3),
                        ]))
                        _st_fi.append(Spacer(1,0.2*cm))
                        _st_fi.append(_bar_tbl)
                        _st_fi.append(Spacer(1,0.3*cm))

                        # ── SECCIÓN PERFIL: Foto + Datos personales + Membresía ──
                        # Foto o avatar
                        _foto_fi=None
                        _fp_fi=sv(r,"foto_path")
                        try:
                            if _fp_fi and os.path.exists(_fp_fi):
                                _im2=_PIL_fi.open(_fp_fi); _iw2,_ih2=_im2.size
                                _r2=min(3.5*cm/_iw2,3.5*cm/_ih2)
                                _foto_fi=RLImage(_fp_fi,width=_iw2*_r2,height=_ih2*_r2)
                        except: pass

                        if not _foto_fi:
                            # Avatar con inicial
                            _ini=sv(r,"nombre")[:1].upper() or "?"
                            _av_buf=io.BytesIO()
                            try:
                                from PIL import Image as _PILAV, ImageDraw as _IDAD, ImageFont as _IFT
                                _av=_PILAV.new("RGB",(140,140),color=(30,30,30))
                                _avd=_IDAD.Draw(_av)
                                _avd.ellipse([5,5,135,135],fill=(_COL.red*255,_COL.green*255,_COL.blue*255) if hasattr(_COL,"red") else (58,155,213))
                                _avd.text((70,70),_ini,fill="white",anchor="mm")
                                _av.save(_av_buf,"PNG")
                                _av_buf.seek(0)
                                _foto_fi=RLImage(_av_buf,width=3.5*cm,height=3.5*cm)
                            except: pass

                        # Datos personales col derecha
                        _dat_pers=[
                            Paragraph(f"<b>{sv(r,'nombre')}</b>",_stnom),
                            Spacer(1,0.1*cm),
                            Paragraph(f"RUT: <b>{rut}</b> &nbsp;·&nbsp; {sv(r,'sexo')} &nbsp;·&nbsp; {sv(r,'edad')} años &nbsp;·&nbsp; 🎂 {sv(r,'mes_cumpleanos')}",_stdat),
                            Paragraph(f"📍 {sv(r,'direccion')}",_stdat),
                            Paragraph(f"📱 {fmt_cel(sv(r,'celular'))} &nbsp;·&nbsp; ✉️ {sv(r,'email')}",_stdat),
                        ]
                        # Línea divisora membresía
                        _memb_bar=Table([[""]],colWidths=[_AW-4.2*cm])
                        _memb_bar.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),rl_colors.HexColor("#DDDDDD")),("TOPPADDING",(0,0),(-1,-1),1),("BOTTOMPADDING",(0,0),(-1,-1),1)]))
                        _dat_pers.append(Spacer(1,0.15*cm))
                        _dat_pers.append(_memb_bar)
                        _dat_pers.append(Spacer(1,0.1*cm))
                        _dat_pers.append(Paragraph(f"💳 <b>{sv(r,'tipo_plan')}</b> &nbsp;·&nbsp; {sv(r,'frecuencia')} &nbsp;·&nbsp; {sv(r,'horario')}",_stdat))
                        _venc_color="#E85050" if (dias_para_vencer(sv(r,'fecha_vencimiento')) or 1)<0 else "#222"
                        _dat_pers.append(Paragraph(f"Estado: <b>{sv(r,'estado')}</b> &nbsp;·&nbsp; Vence: <b>{fmt_fecha(sv(r,'fecha_vencimiento'))}</b> &nbsp;·&nbsp; ${int(float(sv(r,'valor_plan') or 0)):,}",_stdat))
                        _dat_pers.append(Paragraph(f"🎯 {sv(r,'objetivo')} &nbsp;·&nbsp; 📊 {sv(r,'nivel')}",_stdat))

                        _perfil_tbl=Table([[_foto_fi or "",_dat_pers]],
                            colWidths=[3.8*cm,_AW-3.8*cm])
                        _perfil_tbl.setStyle(TableStyle([
                            ("VALIGN",(0,0),(-1,-1),"TOP"),
                            ("ALIGN",(0,0),(0,0),"CENTER"),
                            ("LEFTPADDING",(0,0),(-1,-1),0),
                            ("RIGHTPADDING",(0,0),(-1,-1),0),
                            ("TOPPADDING",(0,0),(-1,-1),0),
                            ("BOTTOMPADDING",(0,0),(-1,-1),0),
                            ("LINEBEFORE",(1,0),(1,0),2,rl_colors.HexColor("#6DBE45")),
                            ("LEFTPADDING",(1,0),(1,0),10),
                        ]))
                        _st_fi.append(_perfil_tbl)
                        _st_fi.append(Spacer(1,0.35*cm))

                        # ── EVALUACIÓN ───────────────────────────────────────
                        if not _dev_p.empty:
                            _ep=_dev_p.iloc[0]
                            _st_fi.append(Paragraph("📏 ÚLTIMA EVALUACIÓN",_stsec))
                            _st_fi.append(Spacer(1,0.1*cm))
                            _ev_hdr=[Paragraph(h,_stgen) for h in ["Peso","IMC","Grasa %","Musc. %","Agua %","Fecha"]]
                            _ev_vals=[Paragraph(str(x),_stdat) for x in [
                                f"{_ep.get('peso','')} kg",str(_ep.get('imc','')),
                                f"{_ep.get('grasa_pct','')}%",f"{_ep.get('masa_musc','')}%",
                                f"{_ep.get('agua_pct','')}%",fmt_fecha(str(_ep.get('fecha','')[:10]))]]
                            _tev=Table([_ev_hdr,_ev_vals],colWidths=[_AW/6]*6)
                            _tev.setStyle(TableStyle([
                                ("BACKGROUND",(0,0),(-1,0),rl_colors.HexColor("#1A1A1A")),
                                ("TOPPADDING",(0,0),(-1,-1),5),("BOTTOMPADDING",(0,0),(-1,-1),5),
                                ("ROWBACKGROUNDS",(0,1),(-1,-1),[rl_colors.HexColor("#F5F5F5")]),
                                ("GRID",(0,0),(-1,-1),0.3,rl_colors.HexColor("#DDD")),
                                ("ALIGN",(0,0),(-1,-1),"CENTER"),
                                ("LEFTPADDING",(0,0),(-1,-1),4),
                            ]))
                            _st_fi.append(_tev)
                            _st_fi.append(Spacer(1,0.3*cm))

                            # Gráfico evolución de peso — barras horizontales
                            _ev_hist=db_query("SELECT fecha,peso FROM evaluaciones WHERE rut=? AND peso IS NOT NULL ORDER BY fecha DESC LIMIT 6",(rut,))
                            if not _ev_hist.empty and len(_ev_hist)>1:
                                _st_fi.append(Paragraph("📈 EVOLUCIÓN DE PESO",_stsec))
                                _st_fi.append(Spacer(1,0.1*cm))
                                _ev_rev=_ev_hist.iloc[::-1].reset_index(drop=True)
                                _w_max=float(_ev_rev["peso"].max() or 1)
                                _w_min=float(_ev_rev["peso"].min() or 0)
                                _bar_aw=_AW-3.5*cm
                                _chart_rows=[]
                                for _,_evr in _ev_rev.iterrows():
                                    _wv=float(_evr.get("peso",0) or 0)
                                    _blen=max(0.3*cm,(_wv-_w_min+1)/(_w_max-_w_min+1)*_bar_aw)
                                    _bar=Table([[""]],colWidths=[_blen])
                                    _bar.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),rl_colors.HexColor("#6DBE45")),("TOPPADDING",(0,0),(-1,-1),5),("BOTTOMPADDING",(0,0),(-1,-1),5)]))
                                    _fecha_lbl=str(_evr.get("fecha",""))[:7]
                                    _chart_rows.append([
                                        Paragraph(_fecha_lbl,_stdat),
                                        _bar,
                                        Paragraph(f"<b>{_wv} kg</b>",_stdat)
                                    ])
                                _tchart=Table(_chart_rows,colWidths=[2.2*cm,_bar_aw,1.2*cm])
                                _tchart.setStyle(TableStyle([
                                    ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
                                    ("LEFTPADDING",(0,0),(-1,-1),2),
                                    ("RIGHTPADDING",(0,0),(-1,-1),2),
                                    ("TOPPADDING",(0,0),(-1,-1),3),
                                    ("BOTTOMPADDING",(0,0),(-1,-1),3),
                                    ("ROWBACKGROUNDS",(0,0),(-1,-1),[rl_colors.white,rl_colors.HexColor("#F5F5F5")]),
                                ]))
                                _st_fi.append(_tchart)
                                _st_fi.append(Spacer(1,0.3*cm))

                        # ── PAGOS ─────────────────────────────────────────────
                        if not _dpg_p.empty:
                            _st_fi.append(Paragraph("💳 ÚLTIMOS PAGOS",_stsec))
                            _st_fi.append(Spacer(1,0.1*cm))
                            _pg_hdr=[[Paragraph(h,_stgen) for h in ["Fecha","Monto","Período","Medio"]]]
                            _pg_rows=[[Paragraph(fmt_fecha(str(pg.get('fecha',''))),_stdat),
                                Paragraph(f"${int(float(pg.get('monto',0))):,}",_stdat),
                                Paragraph(str(pg.get('concepto','')),_stdat),
                                Paragraph(str(pg.get('medio_pago','')),_stdat)] for _,pg in _dpg_p.iterrows()]
                            _tpg=Table(_pg_hdr+_pg_rows,colWidths=[_AW*0.2,_AW*0.2,_AW*0.35,_AW*0.25])
                            _tpg.setStyle(TableStyle([
                                ("BACKGROUND",(0,0),(-1,0),rl_colors.HexColor("#1A1A1A")),
                                ("TOPPADDING",(0,0),(-1,-1),5),("BOTTOMPADDING",(0,0),(-1,-1),5),
                                ("ROWBACKGROUNDS",(0,1),(-1,-1),[rl_colors.white,rl_colors.HexColor("#F5F5F5")]),
                                ("GRID",(0,0),(-1,-1),0.3,rl_colors.HexColor("#DDD")),
                                ("LEFTPADDING",(0,0),(-1,-1),6),
                            ]))
                            _st_fi.append(_tpg)
                            _st_fi.append(Spacer(1,0.4*cm))

                        _st_fi.append(Paragraph(f"Putú Activo — Centro de Entrenamiento · {hoy.strftime('%d/%m/%Y')}",_stfoot))
                        _pd_fi.build(_st_fi)
                        _pdf_fi=_pb_fi.getvalue()
                        st.download_button("⬇️ Descargar ficha PDF",_pdf_fi,f"ficha_{rut}.pdf","application/pdf",key=f"dl_print_{rut}",use_container_width=True)
                    st.session_state.pop("show_print_"+rut,None)
        st.markdown("---")
        # ── TABS DE FICHA ──
        st.markdown('<style>.stTabs{margin-top:-0.5rem}[data-testid="stFileUploaderDropzoneInstructions"]{display:none!important}.stFileUploaderDropzone small{display:none!important}section[data-testid="stFileUploaderDropzone"] small{display:none!important}.stTabs [data-baseweb="tab-list"]{gap:2px}.stTabs [data-baseweb="tab"]{padding:6px 10px;font-size:.82rem}</style>',unsafe_allow_html=True)
        tab_datos,tab_rut,tab_nutr,tab_eval,tab_pagos,tab_qr,tab_doc=st.tabs(_tabs_names)


        # ── TAB DATOS ──
        with tab_datos:
            # Vista rápida de datos actuales — 3 columnas
            import random, string
            def _gen_pass(n=8):
                chars=string.ascii_letters+string.digits
                return ''.join(random.choices(chars,k=n))

            _d1,_d2,_d3=st.columns(3)
            _d1.markdown(f"""<div style='background:{GRIS2};border-radius:9px;padding:12px 16px;font-size:.88rem;height:100%'>
                <b style='color:{VERDE}'>Datos personales</b><br>
                <b>Nombre:</b> {sv(r,"nombre")}<br>
                <b>RUT:</b> {rut}<br>
                <b>Sexo:</b> {sv(r,"sexo")} &nbsp;·&nbsp; <b>Edad:</b> {sv(r,"edad")} años<br>
                <b>Celular:</b> {fmt_cel(sv(r,"celular"))}<br>
                <b>Email:</b> {sv(r,"email")}<br>
                <b>Dirección:</b> {sv(r,"direccion")}
            </div>""",unsafe_allow_html=True)
            _d2.markdown(f"""<div style='background:{GRIS2};border-radius:9px;padding:12px 16px;font-size:.88rem;height:100%'>
                <b style='color:{VERDE}'>Membresía</b><br>
                <b>Plan:</b> {sv(r,"tipo_plan")} · {sv(r,"frecuencia")}<br>
                <b>Valor:</b> ${int(float(sv(r,"valor_plan") or 0)):,}<br>
                <b>Estado:</b> <span style='color:{col_e};font-weight:700'>{sv(r,"estado")}</span><br>
                <b>Inscripción:</b> {fmt_fecha(sv(r,"fecha_inscripcion"))}<br>
                <b>Vencimiento:</b> {fmt_fecha(sv(r,"fecha_vencimiento"))}<br>
                <b>Objetivo:</b> {sv(r,"objetivo")} · <b>Nivel:</b> {sv(r,"nivel")}
            </div>""",unsafe_allow_html=True)
            with _d3:
                _uc_row=get_conn().execute("SELECT email,activo FROM usuarios_clientes WHERE rut=?",(rut,)).fetchone()
                _email_cli=sv(r,"email")
                st.markdown(f"<div style='background:{GRIS2};border-radius:9px;padding:12px 16px;font-size:.88rem'>",unsafe_allow_html=True)
                st.markdown(f"<b style='color:{AZUL}'>🔑 Acceso Mi Cuenta</b>",unsafe_allow_html=True)
                if _uc_row:
                    _uc_estado="✅ Activo" if _uc_row[1] else "🔴 Inactivo"
                    st.markdown(f"<span style='font-size:.82rem'>✉️ {_uc_row[0]}<br>{_uc_estado}</span>",unsafe_allow_html=True)
                    _ucc1,_ucc2=st.columns(2)
                    if _ucc1.button("🔄 Reset clave",key=f"uc_reset_{rut}",use_container_width=True):
                        st.session_state[f"uc_show_{rut}"]=True
                    if _ucc2.button("🚫 Desactivar" if _uc_row[1] else "✅ Activar",key=f"uc_toggle_{rut}",use_container_width=True):
                        _cnuc=get_conn()
                        _cnuc.execute("UPDATE usuarios_clientes SET activo=? WHERE rut=?",(0 if _uc_row[1] else 1,rut))
                        _cnuc.commit(); _cnuc.close(); db_query.clear(); st.rerun()
                else:
                    st.markdown(f"<span style='font-size:.82rem;color:{GRIS_T}'>Sin acceso creado</span>",unsafe_allow_html=True)
                    if st.button("🔑 Crear acceso",key=f"uc_crear_{rut}",use_container_width=True):
                        # Generar clave automática y guardar
                        _auto_pass=_gen_pass()
                        _cnuc3=get_conn()
                        _cnuc3.execute("""INSERT INTO usuarios_clientes (rut,email,password_hash,activo)
                            VALUES (?,?,?,1) ON CONFLICT(rut) DO UPDATE SET email=?,password_hash=?,activo=1""",
                            (rut,_email_cli.strip().lower(),_h(_auto_pass),_email_cli.strip().lower(),_h(_auto_pass)))
                        _cnuc3.commit(); _cnuc3.close(); db_query.clear()
                        st.session_state[f"uc_pass_gen_{rut}"]=_auto_pass
                        st.rerun()
                # Mostrar clave generada si existe
                if st.session_state.get(f"uc_pass_gen_{rut}"):
                    _pg=st.session_state[f"uc_pass_gen_{rut}"]
                    st.markdown(f"""<div style='background:#1A2F1A;border:1px solid {VERDE};border-radius:7px;padding:8px;margin-top:6px;font-size:.82rem'>
                        ✅ Acceso creado<br>
                        ✉️ <b>{_email_cli}</b><br>
                        🔑 Clave: <b style='color:{VERDE};font-size:1rem'>{_pg}</b><br>
                        <span style='color:{GRIS_T};font-size:.72rem'>Comunícala al cliente</span>
                    </div>""",unsafe_allow_html=True)
                    if st.button("✓ Entendido",key=f"uc_ok_{rut}",use_container_width=True):
                        st.session_state.pop(f"uc_pass_gen_{rut}",None); st.rerun()
                # Reset clave — también automático
                if st.session_state.get(f"uc_show_{rut}"):
                    _new_pass=_gen_pass()
                    _cnuc4=get_conn()
                    _cnuc4.execute("UPDATE usuarios_clientes SET password_hash=? WHERE rut=?",(_h(_new_pass),rut))
                    _cnuc4.commit(); _cnuc4.close(); db_query.clear()
                    st.session_state.pop(f"uc_show_{rut}",None)
                    st.session_state[f"uc_pass_gen_{rut}"]=_new_pass
                    st.rerun()
                st.markdown("</div>",unsafe_allow_html=True)
            st.markdown("")
            with st.expander("✏️ Editar datos del cliente",expanded=False):
             with st.form("edit_ficha"):
                st.markdown("**Datos personales**")
                a1,a2=st.columns(2)
                nm=a1.text_input("Nombre *",value=sv(r,"nombre"))
                fn_raw=sv(r,"fecha_nacimiento")
                try: fn_d=date.fromisoformat(fn_raw[:10]) if fn_raw else date(2000,1,1)
                except: fn_d=date(2000,1,1)
                fn=a1.date_input("Fecha nacimiento",value=fn_d,min_value=date(1920,1,1),format="DD/MM/YYYY")
                sx_v=sv(r,"sexo"); SEXOS=["Femenino","Masculino","Otro"]
                sx=a2.selectbox("Sexo",SEXOS,index=SEXOS.index(sx_v) if sx_v in SEXOS else 0)
                dr=a2.text_input("Dirección",value=sv(r,"direccion"))
                ce3=a1.text_input("Celular",value=sv(r,"celular"))
                em2=a2.text_input("E-mail",value=sv(r,"email"))
                tl_v=sv(r,"talla"); TALLAS=["XS","S","M","L","XL","XXL"]
                tl=a1.selectbox("Talla",TALLAS,index=TALLAS.index(tl_v) if tl_v in TALLAS else 2)
                st.markdown("**Contacto emergencia**")
                e1,e2,e3=st.columns(3)
                cen=e1.text_input("Nombre contacto",value=sv(r,"contacto_emergencia"))
                cec=e2.text_input("Celular contacto",value=sv(r,"celular_emergencia"))
                cep=e3.text_input("Parentesco",value=sv(r,"parentesco"))
                st.markdown("**Membresía**")
                m1,m2,m3=st.columns(3)
                fi_raw=sv(r,"fecha_inscripcion")
                try: fi_d=date.fromisoformat(fi_raw[:10]) if fi_raw else hoy
                except: fi_d=hoy
                fi=m1.date_input("Fecha inscripción",value=fi_d,format="DD/MM/YYYY")
                tp_v=sv(r,"tipo_plan"); tp=m2.selectbox("Plan",PLANES,index=PLANES.index(tp_v) if tp_v in PLANES else 0)
                fr_v=sv(r,"frecuencia"); fr=m3.selectbox("Frecuencia",FRECUENCIAS,index=FRECUENCIAS.index(fr_v) if fr_v in FRECUENCIAS else 0)
                ho_v=sv(r,"horario"); ho=m1.selectbox("Horario",HORARIOS,index=HORARIOS.index(ho_v) if ho_v in HORARIOS else 0)
                vp=m2.number_input("Valor $ plan",0,10000000,int(float(sv(r,"valor_plan") or 0)),500)
                pe_v=_PERIODOS_ALIAS.get(sv(r,"periodo_vencimiento","Mensual"),sv(r,"periodo_vencimiento","Mensual"))
                pe=m3.selectbox("Período",PERIODOS,index=PERIODOS.index(pe_v) if pe_v in PERIODOS else 3)
                es_v=sv(r,"estado"); ESTADOS_CLI=["Activo","Inactivo"]
                es3=m1.selectbox("Estado",ESTADOS_CLI,index=ESTADOS_CLI.index(es_v) if es_v in ESTADOS_CLI else 0)
                fv_calc=calcular_vencimiento(fi,pe)
                st.markdown(f'<div class="info-box">📅 Vencimiento calculado: <b>{fmt_fecha(fv_calc)}</b></div>',unsafe_allow_html=True)
                st.markdown("**Salud & objetivo**")
                s1,s2,s3=st.columns(3)
                enf=s1.text_input("Condición médica",value=sv(r,"enfermedad"))
                rst=s2.text_input("Restricciones",value=sv(r,"restricciones"))
                obj_v=sv(r,"objetivo"); obj=s3.selectbox("Objetivo",OBJETIVOS,index=OBJETIVOS.index(obj_v) if obj_v in OBJETIVOS else 0)
                niv_v=sv(r,"nivel").capitalize(); niv=s1.selectbox("Nivel",NIVELES,index=NIVELES.index(niv_v) if niv_v in NIVELES else 0)
                ok_ed=st.form_submit_button("💾 Guardar todos los cambios",use_container_width=True)
             if ok_ed:
                edad_n=int((hoy-fn).days/365.25); mes_n=mes_de_nacimiento(fn)
                mc_n=msg_cumpleanos(nm); mv_n=msg_vencimiento(nm,fv_calc); mr_n=msg_renovacion(nm,fv_calc)
                conn2=get_conn()
                conn2.execute("""UPDATE clientes SET nombre=?,fecha_nacimiento=?,edad=?,mes_cumpleanos=?,sexo=?,
                    direccion=?,celular=?,email=?,contacto_emergencia=?,celular_emergencia=?,parentesco=?,
                    tipo_plan=?,frecuencia=?,horario=?,valor_plan=?,periodo_vencimiento=?,
                    fecha_inscripcion=?,fecha_vencimiento=?,estado=?,talla=?,
                    enfermedad=?,restricciones=?,objetivo=?,nivel=?,
                    mensaje_cumpleanos=?,mensaje_vencimiento=?,mensaje_renovacion=?,modificado=? WHERE rut=?""",
                    (nm,str(fn),edad_n,mes_n,sx,dr,fmt_cel(ce3),em2,cen,fmt_cel(cec),cep,
                     tp,fr,ho,vp,pe,str(fi),fv_calc,es3,tl,
                     enf,rst,obj,niv,mc_n,mv_n,mr_n,datetime.now().isoformat(),rut))
                conn2.commit(); conn2.close()
                st.cache_data.clear()
                st.success("✅ Datos actualizados.")
                st.rerun()

        # ── TAB RUTINA ──
        with tab_rut:
            _rut_fic=db_query("SELECT * FROM rutinas WHERE cliente_rut=? AND activa=1 ORDER BY id DESC LIMIT 1",(rut,))
            if _rut_fic.empty:
                st.markdown('<div class="info-box">Sin rutina activa. Asigna una desde 📋 Rutinas o desde las rutinas guardadas.</div>',unsafe_allow_html=True)
                _todas_rf=db_query("""SELECT r.id,r.nombre,(SELECT COUNT(*) FROM rutina_ejercicios re WHERE re.rutina_id=r.id) as n
                    FROM rutinas r ORDER BY r.fecha_creacion DESC""")
                if not _todas_rf.empty:
                    st.markdown(f"<b style='color:{VERDE}'>Asignar rutina existente:</b>",unsafe_allow_html=True)
                    _opts_rf=[f"{row['nombre']} ({row['n']} ej.)" for _,row in _todas_rf.iterrows()]
                    _sel_rf=st.selectbox("Selecciona rutina:",_opts_rf,key="fic_rut_sel")
                    if st.button("✅ Asignar esta rutina",key="fic_rut_asig",type="primary",use_container_width=True):
                        _idx_rf=_opts_rf.index(_sel_rf)
                        _rid_rf=int(_todas_rf.iloc[_idx_rf]["id"])
                        _caf=get_conn()
                        _caf.execute("UPDATE rutinas SET activa=0 WHERE cliente_rut=?",(rut,))
                        _nrf=_caf.execute("INSERT INTO rutinas (cliente_rut,nombre,activa) VALUES (?,?,1)",
                            (rut,_todas_rf.iloc[_idx_rf]["nombre"])).lastrowid
                        _caf.commit()
                        _ejs_rf=_caf.execute("SELECT * FROM rutina_ejercicios WHERE rutina_id=? ORDER BY dia_semana,orden",(_rid_rf,)).fetchall()
                        for _ep_rf in _ejs_rf:
                            _caf.execute("INSERT INTO rutina_ejercicios (rutina_id,ejercicio_id,dia_semana,orden,metodo,series,repeticiones,peso,tempo_descanso,notas) VALUES (?,?,?,?,?,?,?,?,?,?)",
                                (_nrf,_ep_rf[2],_ep_rf[3],_ep_rf[4],_ep_rf[5],_ep_rf[6],_ep_rf[7],_ep_rf[8],_ep_rf[9],_ep_rf[10]))
                        _caf.commit(); _caf.close(); db_query.clear(); st.rerun()
            else:
                _rutfic=_rut_fic.iloc[0].to_dict()
                _rr1,_rr2=st.columns([4,1])
                _rr1.markdown(f"<b style='color:{VERDE}'>{_rutfic['nombre']}</b>",unsafe_allow_html=True)
                if _rutfic.get("fecha_vencimiento"): _rr1.caption(f"⏳ Vence {fmt_fecha(sv(_rutfic,'fecha_vencimiento'))}")
                if _rr2.button("🗑️ Quitar",key="fic_rut_del",use_container_width=True):
                    _cdrf=get_conn(); _cdrf.execute("UPDATE rutinas SET activa=0 WHERE cliente_rut=?",(rut,)); _cdrf.commit(); _cdrf.close(); db_query.clear(); st.rerun()
                _rb1,_rb2,_rb3=st.columns(3)
                if _rb1.button("✏️ Editar rutina",key="fic_rut_edit",use_container_width=True):
                    st.session_state["_rut_id_activo"]=int(_rutfic["id"])
                    st.session_state._goto="📋 Rutinas"; st.session_state.ver_rut=None; st.rerun()
                if _rb2.button("📄 Ver Rutina PDF",key="fic_ver_rut_pdf",use_container_width=True):
                    st.session_state["fic_show_pdf_"+rut]=not st.session_state.get("fic_show_pdf_"+rut,False)
                    st.rerun()
                # Botón WhatsApp — envía mensaje con resumen de la rutina
                _cel_rut=fmt_cel(sv(r,"celular"))
                _dias_rut_wa=[]
                _ejs_wa=db_query("""SELECT re.dia_semana,e.nombre,re.series,re.repeticiones
                    FROM rutina_ejercicios re JOIN ejercicios e ON e.id=re.ejercicio_id
                    WHERE re.rutina_id=? ORDER BY re.dia_semana,re.orden""",(int(_rutfic["id"]),))
                if not _ejs_wa.empty:
                    for _dwa in ["Día 1","Día 2","Día 3","Día 4","Día 5","Día 6"]:
                        _ddf_wa=_ejs_wa[_ejs_wa["dia_semana"]==_dwa]
                        if _ddf_wa.empty: continue
                        _lineas_wa=[f"*{_dwa}*"]
                        for _ii_wa,(_,_ew) in enumerate(_ddf_wa.iterrows(),1):
                            _lineas_wa.append(f"{_ii_wa}. {_ew['nombre']} — {_ew.get('series','—')}×{_ew.get('repeticiones','—')}")
                        _dias_rut_wa.append("%0A".join(_lineas_wa))
                _msg_wa_rut=(
                    f"Hola *{sv(r,'nombre')}* 💪%0A"
                    f"Aquí tienes tu rutina: *{_rutfic['nombre']}*%0A"
                    f"━━━━━━━━━━━━━━━%0A"
                    +"%0A%0A".join(_dias_rut_wa)+
                    f"%0A━━━━━━━━━━━━━━━%0A"
                    f"¡A entrenar! Putú Activo 🏋️"
                )
                _url_wa_rut=wa_url(_cel_rut,_msg_wa_rut)
                _rb3.markdown(f'<a href="{_url_wa_rut}" target="_blank" style="display:block;background:#25D366;color:white;font-weight:700;padding:9px 8px;border-radius:9px;text-decoration:none;text-align:center;font-size:.85rem;margin-top:2px;">📲 Enviar por WhatsApp</a>',unsafe_allow_html=True)
                # ── Generar PDF con calendario semanal + imagen ──
                if st.session_state.get("fic_show_pdf_"+rut,False):
                    _ejs_vpdf=db_query("""SELECT re.*,e.nombre,e.url_imagen,e.musculo_primario,e.video,e.ejecucion
                        FROM rutina_ejercicios re JOIN ejercicios e ON e.id=re.ejercicio_id
                        WHERE re.rutina_id=? ORDER BY re.dia_semana,re.orden""",(int(_rutfic["id"]),))
                    if not _ejs_vpdf.empty and REPORTLAB_OK:
                        import urllib.request as _ur_fp, re as _re_fp
                        from PIL import Image as _PILfp

                        # ── Color según sexo ─────────────────────────────────
                        _es_fem_pdf = sv(r,"sexo").lower()=="femenino"
                        _COLOR_PDF  = rl_colors.HexColor("#E91E8C") if _es_fem_pdf else rl_colors.HexColor("#6DBE45")
                        _COLOR_TXT  = rl_colors.HexColor("#C2185B") if _es_fem_pdf else rl_colors.HexColor("#3A7A1E")

                        # ── Estilos ──────────────────────────────────────────
                        def _Sfp(n,**k): return ParagraphStyle(n,**k)
                        _sttit =_Sfp("ftit",fontName="Helvetica-Bold",  fontSize=22,textColor=_COLOR_PDF,   spaceAfter=2, leading=26,alignment=TA_CENTER)
                        _stsubt=_Sfp("fsbt",fontName="Helvetica-Bold",  fontSize=13,textColor=rl_colors.HexColor("#1A1A1A"),spaceAfter=1,leading=18,alignment=TA_CENTER)
                        _stsub =_Sfp("fsub",fontName="Helvetica",        fontSize=9.5,textColor=rl_colors.white,spaceAfter=2,leading=15)
                        _stdia =_Sfp("fdia",fontName="Helvetica-Bold",   fontSize=14,textColor=rl_colors.white,alignment=TA_CENTER)
                        _stnom =_Sfp("fnom",fontName="Helvetica-Bold",   fontSize=10,textColor=rl_colors.black,leading=14)
                        _stser =_Sfp("fser",fontName="Helvetica-Bold",   fontSize=9.5,textColor=_COLOR_PDF,  leading=14)
                        _stdet =_Sfp("fdet",fontName="Helvetica",        fontSize=8,  textColor=rl_colors.HexColor("#6B7280"),leading=12)
                        _stlnk =_Sfp("flnk",fontName="Helvetica-Bold",  fontSize=8,  textColor=rl_colors.HexColor("#1D4ED8"),leading=11)
                        _stnum =_Sfp("fnum",fontName="Helvetica-Bold",   fontSize=16, textColor=_COLOR_PDF,  alignment=TA_CENTER)
                        _stejec=_Sfp("fejec",fontName="Helvetica-Oblique",fontSize=8, textColor=rl_colors.HexColor("#333333"),leading=12)
                        _stfot =_Sfp("ffot",fontName="Helvetica-Oblique",fontSize=7.5,textColor=rl_colors.HexColor("#6B7280"),alignment=TA_CENTER)

                        # ── Helpers ──────────────────────────────────────────
                        def _fetch_img_fp(url,w,h):
                            if not url or str(url).strip() in ["","nan"]: return None
                            try:
                                _req=_ur_fp.Request(str(url).strip(),headers={"User-Agent":"Mozilla/5.0"})
                                _dat=_ur_fp.urlopen(_req,timeout=4).read()
                                _im=_PILfp.open(io.BytesIO(_dat)); _iw,_ih=_im.size
                                _r=min(w/_iw,h/_ih)
                                return RLImage(io.BytesIO(_dat),width=_iw*_r,height=_ih*_r)
                            except: return None

                        def _yt_fp(u):
                            if not u or str(u).strip() in ["","nan"]: return None
                            _m=_re_fp.search(r"(?:embed/|youtu\.be/|watch\?v=)([A-Za-z0-9_-]{6,})",str(u))
                            return f"https://youtu.be/{_m.group(1)}" if _m else None

                        def _logo_fp(h=2.0*cm):
                            try:
                                _lp=os.path.join(BASE_DIR,"logo.png")
                                if not os.path.exists(_lp): return None
                                _im=_PILfp.open(_lp); _iw,_ih=_im.size; _ratio=h/_ih
                                return RLImage(_lp,width=_iw*_ratio,height=h)
                            except: return None

                        def _hdr_fp(cli_nom,rut_nom,rut_id):
                            """Encabezado unificado — A4 portrait."""
                            _logo=_logo_fp(2.2*cm)
                            _entren=st.session_state.get("nombre_u","—")
                            _strt=ParagraphStyle("hrt",fontName="Helvetica-Bold",fontSize=20,
                                textColor=_COLOR_PDF,leading=24,alignment=TA_CENTER)
                            _strs=ParagraphStyle("hrs",fontName="Helvetica-Bold",fontSize=11,
                                textColor=rl_colors.white,leading=15,alignment=TA_CENTER)
                            _strd=ParagraphStyle("hrd",fontName="Helvetica",fontSize=9,
                                textColor=rl_colors.HexColor("#AAAAAA"),leading=13)
                            _hdr_inner=[
                                Paragraph("Rutina de Entrenamiento",_strt),
                                Paragraph(rut_nom,_strs),
                                Spacer(1,0.15*cm),
                                Paragraph(
                                    f"<b>Cliente:</b> {cli_nom} &nbsp;·&nbsp; <b>RUT:</b> {rut_id} &nbsp;·&nbsp; <b>Entrenador:</b> {_entren} &nbsp;·&nbsp; <b>Generado:</b> {fmt_fecha(str(hoy))}",
                                    _strd),
                            ]
                            _logo_cell=_logo if _logo else Paragraph("PA",_strd)
                            _ht=Table([[_logo_cell,_hdr_inner]],
                                colWidths=[2.8*cm,_PW_fp-2.8*cm])
                            _ht.setStyle(TableStyle([
                                ("BACKGROUND",(0,0),(-1,-1),rl_colors.HexColor("#1A1A1A")),
                                ("TOPPADDING",(0,0),(-1,-1),10),("BOTTOMPADDING",(0,0),(-1,-1),10),
                                ("LEFTPADDING",(0,0),(0,0),8),("RIGHTPADDING",(0,0),(0,0),8),
                                ("LEFTPADDING",(1,0),(1,0),10),
                                ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
                                ("LINEBELOW",(0,0),(-1,-1),3,_COLOR_PDF),
                            ]))
                            return _ht

                        # ── Documento apaisado ────────────────────────────────
                        _pb_fp=io.BytesIO()
                        _pdoc_fp=SimpleDocTemplate(_pb_fp,pagesize=A4,
                            leftMargin=1.5*cm,rightMargin=1.5*cm,topMargin=1.2*cm,bottomMargin=1*cm)
                        _PW_fp=A4[0]-_pdoc_fp.leftMargin-_pdoc_fp.rightMargin
                        _PH_fp=A4[1]-_pdoc_fp.topMargin-_pdoc_fp.bottomMargin
                        # 2 columnas por día
                        _COL_fp=(_PW_fp-0.4*cm)/2
                        _story_fp=[]

                        DIAS_FP=["Día 1","Día 2","Día 3","Día 4","Día 5","Día 6"]
                        _dias_con=[d for d in DIAS_FP if d in _ejs_vpdf["dia_semana"].values]

                        for _pg_i,_dcon in enumerate(_dias_con):
                            if _pg_i>0:
                                from reportlab.platypus import PageBreak as _PB
                                _story_fp.append(_PB())
                            # Header con título grande
                            _story_fp.append(_hdr_fp(sv(r,"nombre"),_rutfic["nombre"],rut))
                            _story_fp.append(Spacer(1,0.25*cm))
                            # Banner del día
                            _td=Table([[Paragraph(_dcon.upper(),_stdia)]],colWidths=[_PW_fp])
                            _td.setStyle(TableStyle([
                                ("BACKGROUND",(0,0),(-1,-1),_COLOR_PDF),
                                ("TOPPADDING",(0,0),(-1,-1),7),("BOTTOMPADDING",(0,0),(-1,-1),7),
                                ("LEFTPADDING",(0,0),(-1,-1),10),
                            ]))
                            _story_fp.append(_td)
                            _story_fp.append(Spacer(1,0.2*cm))

                            # Ejercicios del día — lista vertical de tarjetas compactas
                            _ddf_fp=_ejs_vpdf[_ejs_vpdf["dia_semana"]==_dcon]
                            _ej_list=_ddf_fp.to_dict("records")

                            # Calcular altura disponible por ejercicio
                            _hdr_h=2.8*cm   # encabezado + día banner
                            _foot_h=0.8*cm  # pie
                            _avail=_PH_fp-_hdr_h-_foot_h
                            _n_ej=max(len(_ej_list),1)
                            _avail_portrait=_PH_fp-2.5*cm  # descontar header+banner+footer
                            _card_h=min((_avail_portrait/_n_ej)-0.15*cm, 6.0*cm)
                            _img_h=min(_card_h-1.0*cm, 4.5*cm)
                            _img_w=_img_h*1.3

                            # Ejecución: ajustar chars según espacio
                            _ejec_chars=max(80, int(_img_h/cm * 55))

                            # Ancho de cada columna (2 cols)
                            _npc2=(_n_ej+1)//2

                            def _card_fp(idx,row,cw):
                                _img=_fetch_img_fp(row.get("url_imagen"),_img_w,_img_h)
                                _vid=_yt_fp(row.get("video",""))
                                _ic=_img if _img else Paragraph("🏋️",_stnum)
                                # Izquierda: número + imagen
                                _left=Table([[Paragraph(str(idx),_stnum)],[_ic]],
                                    colWidths=[_img_w+0.3*cm])
                                _left.setStyle(TableStyle([
                                    ("ALIGN",(0,0),(-1,-1),"CENTER"),
                                    ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
                                    ("BACKGROUND",(0,0),(-1,0),rl_colors.HexColor("#1A1A1A")),
                                    ("BACKGROUND",(0,1),(-1,-1),rl_colors.HexColor("#EFEFEF")),
                                    ("TOPPADDING",(0,0),(-1,-1),3),
                                    ("BOTTOMPADDING",(0,0),(-1,-1),3),
                                ]))
                                # Derecha: detalle
                                _rw=cw-(_img_w+0.3*cm)-0.2*cm
                                _txts=[Paragraph(f"<b>{row['nombre']}</b>",_stnom)]
                                _sr=f"Series: <b>{row.get('series','—')}</b>  Reps: <b>{row.get('repeticiones','—')}</b>"
                                if row.get("peso") and str(row.get("peso","")).strip() not in ["","—","nan"]:
                                    _sr+=f"  Carga: <b>{row['peso']}</b>"
                                _txts.append(Paragraph(_sr,_stser))
                                if row.get("musculo_primario") and str(row.get("musculo_primario","")) not in ["","nan"]:
                                    _txts.append(Paragraph(f"💪 {row['musculo_primario']}",_stdet))
                                _ejec=str(row.get("ejecucion","")).strip()
                                if _ejec and _ejec not in ["","nan"]:
                                    _txts.append(Paragraph(f"{_ejec[:_ejec_chars]}{'…' if len(_ejec)>_ejec_chars else ''}",_stejec))
                                if _vid:
                                    _txts.append(Paragraph(f'<link href="{_vid}">▶ Ver video</link>',_stlnk))
                                _right=Table([[_txts]],colWidths=[_rw])
                                _right.setStyle(TableStyle([
                                    ("VALIGN",(0,0),(-1,-1),"TOP"),
                                    ("TOPPADDING",(0,0),(-1,-1),5),
                                    ("BOTTOMPADDING",(0,0),(-1,-1),5),
                                    ("LEFTPADDING",(0,0),(-1,-1),6),
                                    ("RIGHTPADDING",(0,0),(-1,-1),4),
                                ]))
                                _c=Table([[_left,_right]],
                                    colWidths=[_img_w+0.3*cm, _rw])
                                _c.setStyle(TableStyle([
                                    ("LINEABOVE",(0,0),(-1,0),2.5,_COLOR_PDF),
                                    ("BOX",(0,0),(-1,-1),0.5,rl_colors.HexColor("#DDDDDD")),
                                    ("VALIGN",(0,0),(-1,-1),"TOP"),
                                    ("LEFTPADDING",(0,0),(-1,-1),0),
                                    ("RIGHTPADDING",(0,0),(-1,-1),0),
                                    ("TOPPADDING",(0,0),(-1,-1),0),
                                    ("BOTTOMPADDING",(0,0),(-1,-1),0),
                                ]))
                                return _c

                            # 2 columnas de ejercicios
                            _cols_fp=[]
                            for _ci in range(2):
                                _chunk=_ej_list[_ci*_npc2:(_ci+1)*_npc2]
                                _inner=[]
                                for _ki,_ej in enumerate(_chunk):
                                    _inner.append([_card_fp(_ci*_npc2+_ki+1,_ej,_COL_fp)])
                                    _inner.append([Spacer(1,0.12*cm)])
                                if _inner:
                                    _ct=Table(_inner,colWidths=[_COL_fp],
                                        splitByRow=True)
                                    _ct.setStyle(TableStyle([
                                        ("LEFTPADDING",(0,0),(-1,-1),0),
                                        ("RIGHTPADDING",(0,0),(-1,-1),0),
                                        ("TOPPADDING",(0,0),(-1,-1),0),
                                        ("BOTTOMPADDING",(0,0),(-1,-1),0),
                                    ]))
                                    _cols_fp.append(_ct)
                                else:
                                    _cols_fp.append(Spacer(1,0.1*cm))
                            _grid=Table([_cols_fp],colWidths=[_COL_fp]*2,
                                splitByRow=True)
                            _grid.setStyle(TableStyle([
                                ("VALIGN",(0,0),(-1,-1),"TOP"),
                                ("LEFTPADDING",(0,0),(-1,-1),3),
                                ("RIGHTPADDING",(0,0),(-1,-1),3),
                                ("TOPPADDING",(0,0),(-1,-1),0),
                                ("BOTTOMPADDING",(0,0),(-1,-1),0),
                                ("LINEBEFORE",(1,0),(1,-1),0.5,
                                    rl_colors.HexColor("#DDDDDD")),
                            ]))
                            _story_fp.append(_grid)
                            _story_fp.append(Spacer(1,0.2*cm))
                            from reportlab.platypus import HRFlowable as _HR
                            _story_fp.append(_HR(width=_PW_fp,thickness=0.5,color=_COLOR_PDF,spaceAfter=3))
                            _story_fp.append(Paragraph("Putú Activo — Centro de Entrenamiento · 2026",_stfot))

                        _pdoc_fp.build(_story_fp)
                        _pdf_fp=_pb_fp.getvalue()
                        _pdf_b64_fp=base64.b64encode(_pdf_fp).decode()
                        _fdl1,_fdl2,_fdl3=st.columns(3)
                        _fdl1.download_button("⬇️ Descargar PDF rutina",_pdf_fp,
                            f"Rutina_{sv(r,'nombre').replace(' ','_')}.pdf","application/pdf",
                            key=f"dl_rut_fic_{rut}",use_container_width=True)
                        _fdl2.markdown(f'<a href="data:application/pdf;base64,{_pdf_b64_fp}" target="_blank" style="display:block;background:{GRIS2};color:{VERDE};font-weight:700;border:1px solid {VERDE};padding:9px 16px;border-radius:9px;text-decoration:none;text-align:center;font-size:.92rem;">🔗 Abrir en ventana</a>',unsafe_allow_html=True)
                        if _fdl3.button("✕ Cerrar PDF",key=f"close_pdf_{rut}",use_container_width=True):
                            st.session_state["fic_show_pdf_"+rut]=False; st.rerun()
                        st.markdown(f'<object data="data:application/pdf;base64,{_pdf_b64_fp}" type="application/pdf" width="100%" height="600px" style="border:1px solid {GRIS3};border-radius:10px;margin-top:8px;"><p><a href="data:application/pdf;base64,{_pdf_b64_fp}" download>Descargar</a></p></object>',unsafe_allow_html=True)
                    elif not REPORTLAB_OK:
                        st.warning("ReportLab no disponible para generar PDF.")
                    else:
                        st.markdown('<div class="info-box">Esta rutina no tiene ejercicios aún.</div>',unsafe_allow_html=True)
                # Calendario semanal con imagen y nombre
                _ejs_fic=db_query("""SELECT re.*,e.nombre,e.url_imagen,e.musculo_primario,e.ejecucion,e.video
                    FROM rutina_ejercicios re JOIN ejercicios e ON e.id=re.ejercicio_id
                    WHERE re.rutina_id=? ORDER BY re.dia_semana,re.orden""",(int(_rutfic["id"]),))
                TODOS_DIAS=["Día 1","Día 2","Día 3","Día 4","Día 5","Día 6"]
                if not _ejs_fic.empty:
                    _dias_activos=[d for d in TODOS_DIAS if d in _ejs_fic["dia_semana"].values]
                    if not _dias_activos:
                        st.markdown('<div class="info-box">Sin ejercicios.</div>',unsafe_allow_html=True)
                    else:
                        # ── Vista detalle de ejercicio a pantalla completa ──
                        _fic_ej_sel=st.session_state.get(f"fic_ej_sel_{rut}")
                        if _fic_ej_sel:
                            _fic_ej_df=db_query("SELECT * FROM ejercicios WHERE id=?",(_fic_ej_sel,))
                            if not _fic_ej_df.empty:
                                _fej=_fic_ej_df.iloc[0].to_dict()
                                if st.button("← Volver a la rutina",key=f"fic_ej_volver_{rut}",type="primary"):
                                    st.session_state.pop(f"fic_ej_sel_{rut}",None); st.rerun()
                                st.markdown(f"<h3 style='color:{VERDE};margin:4px 0'>{_fej['nombre']}</h3>",unsafe_allow_html=True)
                                _fd1,_fd2=st.columns([1,1.6])
                                with _fd1:
                                    if str(_fej.get("url_imagen","")).strip() not in ["","nan"]:
                                        try: st.image(_fej["url_imagen"],use_container_width=True)
                                        except: st.markdown('<div style="font-size:3rem;text-align:center">🏋️</div>',unsafe_allow_html=True)
                                    else: st.markdown('<div style="font-size:3rem;text-align:center">🏋️</div>',unsafe_allow_html=True)
                                    st.markdown(f"""<div style='background:{GRIS2};border-radius:9px;padding:12px 14px;margin-top:10px'>
                                        <b style='color:{VERDE}'>💪 Músculos</b><br>
                                        <span style='font-size:.88rem'><b>Principal:</b> {sv(_fej,"musculo_primario","—")}</span><br>
                                        <span style='font-size:.88rem'><b>Secundario:</b> {sv(_fej,"musculo_secundario","—")}</span>
                                    </div>""",unsafe_allow_html=True)
                                with _fd2:
                                    # Series/reps del ejercicio en esta rutina
                                    _fej_re=_ejs_fic[_ejs_fic["ejercicio_id"]==_fic_ej_sel] if "ejercicio_id" in _ejs_fic.columns else pd.DataFrame()
                                    if not _fej_re.empty:
                                        _fej_red=_fej_re.iloc[0].to_dict()
                                        st.markdown(f"""<div style='background:{GRIS2};border-radius:9px;padding:12px 14px;margin-bottom:10px'>
                                            <b style='color:{VERDE}'>📊 Prescripción</b><br>
                                            <span style='font-size:.9rem'>Series: <b>{sv(_fej_red,"series","—")}</b> &nbsp; Reps: <b>{sv(_fej_red,"repeticiones","—")}</b></span><br>
                                            <span style='font-size:.88rem'>Carga: <b>{sv(_fej_red,"peso","—")}</b> &nbsp; Descanso: <b>{sv(_fej_red,"tempo_descanso","—")}</b></span>
                                        </div>""",unsafe_allow_html=True)
                                    if str(_fej.get("ejecucion","")).strip() not in ["","nan"]:
                                        st.markdown(f"""<div style='background:{GRIS2};border-radius:9px;padding:12px 14px;margin-bottom:10px'>
                                            <b style='color:{VERDE}'>📋 Ejecución</b><br>
                                            <span style='font-size:.87rem;line-height:1.6'>{str(_fej["ejecucion"]).replace(chr(10),"<br>")}</span>
                                        </div>""",unsafe_allow_html=True)
                                    if str(_fej.get("video","")).strip() not in ["","nan"]:
                                        import re as _re_fic
                                        _vm_fic=_re_fic.search(r"(?:embed/|youtu\.be/|watch\?v=)([A-Za-z0-9_-]{6,})",str(_fej["video"]))
                                        if _vm_fic:
                                            st.markdown(f"<b style='color:{VERDE}'>🎥 Video</b>",unsafe_allow_html=True)
                                            st.video(f"https://www.youtube.com/watch?v={_vm_fic.group(1)}")
                            else:
                                st.session_state.pop(f"fic_ej_sel_{rut}",None); st.rerun()
                        else:
                            # ── Calendario: días en columnas, ejercicios debajo ──
                            _n_dias=len(_dias_activos)
                            _col_dia_rut="#E91E8C" if es_fem else "#3A9BD5"
                            _cal_h=st.columns(_n_dias)
                            for _ic,_cc in enumerate(_cal_h):
                                _cc.markdown(f'<div style="background:{_col_dia_rut};color:#fff;text-align:center;font-weight:700;font-size:.8rem;padding:6px 2px;border-radius:6px 6px 0 0">{_dias_activos[_ic]}</div>',unsafe_allow_html=True)
                            _cal_b=st.columns(_n_dias)
                            for _ic2,(_cb,_dfl_fic) in enumerate(zip(_cal_b,_dias_activos)):
                                _dfic2=_ejs_fic[_ejs_fic["dia_semana"]==_dfl_fic]
                                with _cb:
                                    if _dfic2.empty:
                                        st.markdown(f'<div style="background:{GRIS2};border:1px solid #2E2E2E;border-top:none;border-radius:0 0 6px 6px;padding:8px 4px;min-height:80px;text-align:center;color:{GRIS_T};font-size:.7rem">Descanso</div>',unsafe_allow_html=True)
                                    else:
                                        st.markdown(f'<div style="background:{GRIS2};border:1px solid #2E2E2E;border-top:none;border-radius:0 0 6px 6px;padding:4px;">',unsafe_allow_html=True)
                                        for _iifc,(_,_effc) in enumerate(_dfic2.iterrows(),1):
                                            _effd=_effc.to_dict()
                                            _nn2=str(_effd.get("nombre",""))
                                            _sr3=sv(_effd,"series","—"); _rp3=sv(_effd,"repeticiones","—")
                                            _ps3=sv(_effd,"peso","")
                                            _img_url=str(_effd.get("url_imagen","")).strip()
                                            with st.container(border=True):
                                                if _img_url and _img_url!="nan":
                                                    try: st.image(_img_url,use_container_width=True)
                                                    except: st.markdown(f'<div style="height:80px;background:{GRIS3};border-radius:6px;display:flex;align-items:center;justify-content:center;font-size:2rem">🏋️</div>',unsafe_allow_html=True)
                                                else:
                                                    st.markdown(f'<div style="height:80px;background:{GRIS3};border-radius:6px;display:flex;align-items:center;justify-content:center;font-size:2rem">🏋️</div>',unsafe_allow_html=True)
                                                st.markdown(f'<div style="font-size:.75rem;font-weight:700;color:{BLANCO};line-height:1.2;margin:3px 0">{_iifc}. {_nn2}</div>',unsafe_allow_html=True)
                                                st.markdown(f'<div style="font-size:.72rem;color:{VERDE}">{_sr3}×{_rp3}{" · "+_ps3 if _ps3 and _ps3 not in ["","—"] else ""}</div>',unsafe_allow_html=True)
                                                if st.button("Ver detalle",key=f"fic_ej_{rut}_{_ic2}_{_iifc}",use_container_width=True):
                                                    st.session_state[f"fic_ej_sel_{rut}"]=int(_effd.get("ejercicio_id",_effd.get("id",0))); st.rerun()
                                        st.markdown('</div>',unsafe_allow_html=True)
                else:
                    st.markdown('<div class="info-box">Esta rutina aún no tiene ejercicios. Edítala desde 📋 Rutinas.</div>',unsafe_allow_html=True)

        # ── TAB NUTRICIÓN ──
        with tab_nutr:
            DIAS_N=["Lunes","Martes","Miércoles","Jueves","Viernes","Sábado","Domingo"]
            DIAS_N_S=["Lun","Mar","Mié","Jue","Vie","Sáb","Dom"]
            COMIDAS_N=["Desayuno","Colación AM","Almuerzo","Colación PM","Cena","Extra"]
            COMIDAS_ICON={"Desayuno":"🌅","Colación AM":"🍎","Almuerzo":"🍽️","Colación PM":"🥑","Cena":"🌙","Extra":"➕"}

            # Helper: calcular macros proporcional a cantidad_g
            def _calc_macros(alim, cantidad_g):
                base = float(alim.get("peso_neto_g") or 100) or 100
                f = cantidad_g / base
                return {
                    "kcal":  round(float(alim.get("kcal",0))*f, 1),
                    "prot":  round(float(alim.get("proteina_g",0))*f, 1),
                    "lip":   round(float(alim.get("lipidos_g",0))*f, 1),
                    "hdc":   round(float(alim.get("hdc_g",0))*f, 1),
                    "fibra": round(float(alim.get("fibra_g",0))*f, 1),
                }

            # ── Planes del cliente ──────────────────────────────────────
            _planes = db_query("SELECT * FROM planes_nutri WHERE cliente_rut=? ORDER BY activo DESC,fecha_creacion DESC",(rut,))

            # Selector de plan activo
            _plan_id = st.session_state.get(f"nutr_plan_{rut}")
            if not _planes.empty:
                _opts_p = [f"{'✅' if r2['activo'] else '📦'} {r2['nombre']}" for _,r2 in _planes.iterrows()]
                _idx_p = 0
                if _plan_id and _plan_id in _planes["id"].values:
                    _idx_p = _planes[_planes["id"]==_plan_id].index[0] if len(_planes[_planes["id"]==_plan_id])>0 else 0
                    _idx_p = list(_planes["id"]).index(_plan_id)
                _sel_p = st.selectbox("Plan nutricional activo",_opts_p,index=_idx_p,key=f"nutr_sel_{rut}")
                _plan_row = _planes.iloc[_opts_p.index(_sel_p)].to_dict()
                _plan_id = int(_plan_row["id"])
                st.session_state[f"nutr_plan_{rut}"] = _plan_id
            else:
                _plan_row = None; _plan_id = None
                st.markdown(f'<div class="info-box">Sin planes nutricionales. Crea uno abajo.</div>',unsafe_allow_html=True)

            # Crear nuevo plan
            with st.expander("➕ Crear nuevo plan nutricional",expanded=_plan_id is None):
                with st.form(f"form_plan_{rut}",clear_on_submit=True):
                    _pn1,_pn2=st.columns(2)
                    _np_nom=_pn1.text_input("Nombre del plan *",placeholder="Ej: Pérdida de grasa · Volumen")
                    _np_prof=_pn2.text_input("Profesional que diseña",value=st.session_state.get("nombre_u",""))
                    _pn3,_pn4=st.columns(2)
                    _np_obj=_pn3.selectbox("Objetivo",["Pérdida de grasa","Mantención","Ganancia muscular","Rendimiento","Otro"])
                    _np_kcal=_pn4.number_input("Kcal objetivo/día",0,10000,2000,50)
                    _np_notas=st.text_area("Notas / indicaciones",height=60)
                    _ok_plan=st.form_submit_button("💾 Crear plan",type="primary",use_container_width=True)
                if _ok_plan and _np_nom.strip():
                    _cn_plan=get_conn()
                    _new_plan=_cn_plan.execute(
                        "INSERT INTO planes_nutri (cliente_rut,nombre,profesional,objetivo,kcal_objetivo,notas,activo) VALUES (?,?,?,?,?,?,1)",
                        (rut,_np_nom.strip(),_np_prof.strip(),_np_obj,_np_kcal,_np_notas)).lastrowid
                    _cn_plan.commit(); _cn_plan.close()
                    st.session_state[f"nutr_plan_{rut}"]=int(_new_plan)
                    db_query.clear(); st.rerun()

            if _plan_id is not None:
                st.markdown(f"""<div style='background:{GRIS2};border-radius:8px;padding:8px 14px;margin:4px 0;font-size:.85rem'>
                    <b style='color:{VERDE}'>{_plan_row["nombre"]}</b> &nbsp;·&nbsp;
                    👨‍⚕️ <b>{_plan_row.get("profesional") or "—"}</b> &nbsp;·&nbsp;
                    🎯 {_plan_row.get("objetivo") or "—"} &nbsp;·&nbsp;
                    🔥 <b>{int(_plan_row.get("kcal_objetivo") or 0):,} kcal/día objetivo</b>
                </div>""",unsafe_allow_html=True)

                st.divider()

                # ── MODO: Diseñar | Ver calendario | Informe PDF ───────────
                _modo_n=st.radio("",["📅 Calendario semanal","✏️ Agregar comidas","📄 Informe PDF"],
                    horizontal=True,key=f"modo_n_{rut}")

                # ── Cargar todos los datos del plan ─────────────────────────
                _comidas_df=db_query("""SELECT cp.*,
                    GROUP_CONCAT(a.nombre||'|'||ca.cantidad_g||'|'||ca.kcal_calc||'|'||ca.prot_calc||'|'||ca.lip_calc||'|'||ca.hdc_calc||'|'||ca.fibra_calc||'|'||ca.id, ';;') as items
                    FROM comidas_plan cp
                    LEFT JOIN comida_alimentos ca ON ca.comida_id=cp.id
                    LEFT JOIN alimentos a ON a.id=ca.alimento_id
                    WHERE cp.plan_id=? GROUP BY cp.id ORDER BY cp.dia_semana,cp.orden""",(_plan_id,))

                def _parse_items(items_str):
                    """Devuelve lista de dicts con datos de alimentos de una comida."""
                    if not items_str or str(items_str)=="nan": return []
                    result=[]
                    for it in str(items_str).split(";;"):
                        p=it.split("|")
                        if len(p)>=8:
                            result.append({"nombre":p[0],"cantidad_g":float(p[1] or 0),
                                "kcal":float(p[2] or 0),"prot":float(p[3] or 0),
                                "lip":float(p[4] or 0),"hdc":float(p[5] or 0),
                                "fibra":float(p[6] or 0),"ca_id":int(p[7] or 0)})
                    return result

                def _totales_dia(dia):
                    """Suma macros de todas las comidas de un día."""
                    tot={"kcal":0,"prot":0,"lip":0,"hdc":0,"fibra":0}
                    if _comidas_df.empty: return tot
                    _dc=_comidas_df[_comidas_df["dia_semana"]==dia]
                    for _,_cm in _dc.iterrows():
                        for _it in _parse_items(_cm.get("items")):
                            for k in tot: tot[k]=round(tot[k]+_it.get(k,0),1)
                    return tot

                # ══════════════════════════════════════════════════════════
                # MODO 1 — CALENDARIO SEMANAL
                # ══════════════════════════════════════════════════════════
                if _modo_n=="📅 Calendario semanal":
                    # Cabecera 7 columnas
                    _cal_hn=st.columns(7)
                    for _ic,_cc in enumerate(_cal_hn):
                        _tot_d=_totales_dia(DIAS_N[_ic])
                        _cc.markdown(f'<div style="background:{VERDE};color:#000;text-align:center;font-weight:700;font-size:.78rem;padding:5px 2px;border-radius:6px 6px 0 0">{DIAS_N_S[_ic]}<br><span style="font-size:.65rem;font-weight:500">{int(_tot_d["kcal"])} kcal</span></div>',unsafe_allow_html=True)

                    _cal_bn=st.columns(7)
                    for _ic2,(_cb2,_dfl2) in enumerate(zip(_cal_bn,DIAS_N)):
                        with _cb2:
                            _dc2=_comidas_df[_comidas_df["dia_semana"]==_dfl2] if not _comidas_df.empty else pd.DataFrame()
                            _hcal=f'<div style="background:#1A1A1A;border:1px solid #2E2E2E;border-top:none;border-radius:0 0 6px 6px;padding:4px;min-height:180px;">'
                            if _dc2.empty:
                                _hcal+=f'<div style="color:{GRIS_T};font-size:.6rem;text-align:center;padding:8px 0">Sin comidas</div>'
                            for _,_cm2 in _dc2.iterrows():
                                _tipo=_cm2.get("tipo_comida","")
                                _icon=COMIDAS_ICON.get(_tipo,"🍴")
                                _items2=_parse_items(_cm2.get("items"))
                                _tot_cm=sum(i["kcal"] for i in _items2)
                                _hcal+=f'<div style="margin:2px 0;padding:2px 3px;background:#252525;border-radius:3px;border-left:2px solid {VERDE};">'
                                _hcal+=f'<div style="font-size:.6rem;font-weight:700;color:{VERDE}">{_icon} {_tipo}</div>'
                                for _it2 in _items2[:3]:
                                    _nn3=str(_it2["nombre"])[:16]
                                    _hcal+=f'<div style="font-size:.55rem;color:{BLANCO};line-height:1.2">{_nn3}</div>'
                                if len(_items2)>3: _hcal+=f'<div style="font-size:.52rem;color:{GRIS_T}">+{len(_items2)-3} más</div>'
                                if _tot_cm>0: _hcal+=f'<div style="font-size:.55rem;color:{NARANJA}">{int(_tot_cm)} kcal</div>'
                                _hcal+='</div>'
                            _hcal+='</div>'
                            st.markdown(_hcal,unsafe_allow_html=True)

                    # Resumen semanal bajo el calendario
                    st.markdown(f"<b style='color:{VERDE}'>Resumen semanal</b>",unsafe_allow_html=True)
                    _totales_sem={"kcal":0,"prot":0,"lip":0,"hdc":0,"fibra":0}
                    _dias_con_data=0
                    for _ds in DIAS_N:
                        _td=_totales_dia(_ds)
                        if _td["kcal"]>0: _dias_con_data+=1
                        for k in _totales_sem: _totales_sem[k]=round(_totales_sem[k]+_td[k],1)
                    _prom=lambda v: round(v/_dias_con_data,1) if _dias_con_data>0 else 0
                    _rs1,_rs2,_rs3,_rs4,_rs5=st.columns(5)
                    _rs1.metric("🔥 Kcal prom/día",f"{_prom(_totales_sem['kcal'])} kcal")
                    _rs2.metric("🥩 Prot prom/día",f"{_prom(_totales_sem['prot'])} g")
                    _rs3.metric("🌾 HDC prom/día",f"{_prom(_totales_sem['hdc'])} g")
                    _rs4.metric("🥑 Grasas prom/día",f"{_prom(_totales_sem['lip'])} g")
                    _rs5.metric("🌿 Fibra prom/día",f"{_prom(_totales_sem['fibra'])} g")

                # ══════════════════════════════════════════════════════════
                # MODO 2 — AGREGAR COMIDAS
                # ══════════════════════════════════════════════════════════
                elif _modo_n=="✏️ Agregar comidas":
                    _an1,_an2=st.columns(2)
                    _dia_sel_n=_an1.selectbox("Día",DIAS_N,key=f"dia_n_{rut}")
                    _com_sel_n=_an2.selectbox("Tipo de comida",COMIDAS_N,key=f"com_n_{rut}")

                    # Comidas existentes del día seleccionado
                    _dc_sel=_comidas_df[(_comidas_df["dia_semana"]==_dia_sel_n)&(_comidas_df["tipo_comida"]==_com_sel_n)] if not _comidas_df.empty else pd.DataFrame()

                    if not _dc_sel.empty:
                        _com_id=int(_dc_sel.iloc[0]["id"])
                        _items_sel=_parse_items(_dc_sel.iloc[0].get("items"))
                        # Mostrar alimentos en esta comida
                        if _items_sel:
                            st.markdown(f"<b style='color:{VERDE}'>{COMIDAS_ICON.get(_com_sel_n,'')} {_com_sel_n} — {_dia_sel_n}</b>",unsafe_allow_html=True)
                            _tot_c={"kcal":0,"prot":0,"lip":0,"hdc":0,"fibra":0}
                            for _it3 in _items_sel:
                                _tot_c["kcal"]+=_it3["kcal"]; _tot_c["prot"]+=_it3["prot"]
                                _tot_c["lip"]+=_it3["lip"]; _tot_c["hdc"]+=_it3["hdc"]; _tot_c["fibra"]+=_it3["fibra"]
                                _nic1,_nic2,_nic3,_nic4,_nic5,_nic6=st.columns([2.5,.8,.8,.8,.8,.5])
                                _nic1.markdown(f"<span style='font-size:.85rem'>{_it3['nombre']}</span>",unsafe_allow_html=True)
                                _nic2.markdown(f"<span style='font-size:.82rem;color:{GRIS_T}'>{_it3['cantidad_g']}g</span>",unsafe_allow_html=True)
                                _nic3.markdown(f"<span style='font-size:.8rem;color:{NARANJA}'>{_it3['kcal']} kcal</span>",unsafe_allow_html=True)
                                _nic4.markdown(f"<span style='font-size:.8rem;color:{VERDE}'>{_it3['prot']}g P</span>",unsafe_allow_html=True)
                                _nic5.markdown(f"<span style='font-size:.8rem;color:{AZUL}'>{_it3['hdc']}g C</span>",unsafe_allow_html=True)
                                if _nic6.button("🗑️",key=f"del_ca_{_it3['ca_id']}",use_container_width=True):
                                    _cad=get_conn(); _cad.execute("DELETE FROM comida_alimentos WHERE id=?",(_it3["ca_id"],)); _cad.commit(); _cad.close(); db_query.clear(); st.rerun()
                            # Totales comida
                            st.markdown(f'<div style="background:{GRIS2};border-radius:6px;padding:5px 12px;font-size:.8rem"><b>Total {_com_sel_n}:</b> 🔥{round(_tot_c["kcal"],1)} kcal · 🥩{round(_tot_c["prot"],1)}g P · 🌾{round(_tot_c["hdc"],1)}g C · 🥑{round(_tot_c["lip"],1)}g G</div>',unsafe_allow_html=True)
                    else:
                        st.caption(f"Sin alimentos en {_com_sel_n} del {_dia_sel_n}.")

                    # Buscador de alimentos
                    st.markdown(f"<b style='color:{VERDE}'>➕ Buscar y agregar alimento</b>",unsafe_allow_html=True)
                    _nb1,_nb2,_nb3=st.columns([3,1.5,1])
                    _btxt=_nb1.text_input("🔍 Nombre del alimento",placeholder="pollo, arroz, manzana...",key=f"nutr_bej_{rut}",label_visibility="collapsed")
                    _bgrupo=_nb2.selectbox("Grupo",["Todos"]+sorted(get_conn().execute("SELECT DISTINCT grupo FROM alimentos ORDER BY grupo").fetchall(),key=lambda x:x[0]),
                        format_func=lambda x:"Todos" if x=="Todos" else (x[0] if isinstance(x,tuple) else x),key=f"nutr_grp_{rut}",label_visibility="collapsed")
                    _bcant=_nb3.number_input("Porción (g)",10,2000,100,10,key=f"nutr_cant_{rut}",label_visibility="collapsed")

                    _bgrupo_val=None if _bgrupo=="Todos" else (_bgrupo[0] if isinstance(_bgrupo,tuple) else _bgrupo)
                    if _btxt.strip() or _bgrupo_val:
                        _aw=[]; _ap=[]
                        if _btxt.strip(): _aw.append("nombre LIKE ?"); _ap.append(f"%{_btxt}%")
                        if _bgrupo_val: _aw.append("grupo=?"); _ap.append(_bgrupo_val)
                        _alim_res=db_query("SELECT * FROM alimentos WHERE "+" AND ".join(_aw)+" ORDER BY nombre LIMIT 18",tuple(_ap))
                    else:
                        _alim_res=db_query("SELECT * FROM alimentos ORDER BY nombre LIMIT 18")

                    if not _alim_res.empty:
                        for _ach in range(0,len(_alim_res),3):
                            _acols=st.columns(3)
                            for _aci,(_,_afr) in enumerate(_alim_res.iloc[_ach:_ach+3].iterrows()):
                                _afd=_afr.to_dict()
                                with _acols[_aci]:
                                    with st.container(border=True):
                                        st.markdown(f"<span style='font-size:.82rem;font-weight:700'>{_afd['nombre']}</span>",unsafe_allow_html=True)
                                        st.caption(f"{_afd.get('grupo','')} · {_afd.get('peso_neto_g',100)}g base")
                                        _m=_calc_macros(_afd,_bcant)
                                        st.markdown(f"<span style='font-size:.75rem;color:{NARANJA}'>🔥{_m['kcal']} kcal</span> <span style='font-size:.73rem;color:{VERDE}'>P:{_m['prot']}g</span> <span style='font-size:.73rem;color:{AZUL}'>C:{_m['hdc']}g</span> <span style='font-size:.73rem;color:{GRIS_T}'>G:{_m['lip']}g</span>",unsafe_allow_html=True)
                                        if st.button("➕ Agregar",key=f"add_n_{rut}_{_dia_sel_n}_{_com_sel_n}_{_afd['id']}",use_container_width=True):
                                            _cna=get_conn()
                                            # Crear comida si no existe
                                            if _dc_sel.empty:
                                                _max_ord=_cna.execute("SELECT COALESCE(MAX(orden),0) FROM comidas_plan WHERE plan_id=? AND dia_semana=?",(_plan_id,_dia_sel_n)).fetchone()[0]
                                                _com_id=_cna.execute("INSERT INTO comidas_plan (plan_id,dia_semana,tipo_comida,orden) VALUES (?,?,?,?)",
                                                    (_plan_id,_dia_sel_n,_com_sel_n,_max_ord+1)).lastrowid
                                                _cna.commit()
                                            _m2=_calc_macros(_afd,_bcant)
                                            _cna.execute("INSERT INTO comida_alimentos (comida_id,alimento_id,cantidad_g,kcal_calc,prot_calc,lip_calc,hdc_calc,fibra_calc) VALUES (?,?,?,?,?,?,?,?)",
                                                (_com_id,int(_afd["id"]),_bcant,_m2["kcal"],_m2["prot"],_m2["lip"],_m2["hdc"],_m2["fibra"]))
                                            _cna.commit(); _cna.close(); db_query.clear(); st.rerun()

                    # Resumen del día
                    st.divider()
                    _td_n=_totales_dia(_dia_sel_n)
                    _kcal_obj=float(_plan_row.get("kcal_objetivo") or 0)
                    st.markdown(f"<b style='color:{VERDE}'>Resumen {_dia_sel_n}</b>",unsafe_allow_html=True)
                    _rm1,_rm2,_rm3,_rm4=st.columns(4)
                    _rm1.metric("🔥 Kcal",f"{_td_n['kcal']}",f"obj {int(_kcal_obj)}" if _kcal_obj else None)
                    _rm2.metric("🥩 Proteínas",f"{_td_n['prot']} g")
                    _rm3.metric("🌾 Carbohid.",f"{_td_n['hdc']} g")
                    _rm4.metric("🥑 Grasas",f"{_td_n['lip']} g")

                # ══════════════════════════════════════════════════════════
                # MODO 3 — INFORME PDF
                # ══════════════════════════════════════════════════════════
                elif _modo_n=="📄 Informe PDF":
                    if REPORTLAB_OK:
                        if st.button("📄 Generar informe nutricional",type="primary",use_container_width=True,key=f"gen_nutr_pdf_{rut}"):
                            _pbuf_n=io.BytesIO()
                            _pdoc_n=SimpleDocTemplate(_pbuf_n,pagesize=landscape(A4),
                                leftMargin=1.2*cm,rightMargin=1.2*cm,topMargin=1*cm,bottomMargin=1*cm)
                            _aw_n=landscape(A4)[0]-_pdoc_n.leftMargin-_pdoc_n.rightMargin
                            _story_n=[]
                            _stt_n=ParagraphStyle("tn",fontName="Helvetica-Bold",fontSize=13,textColor=rl_colors.HexColor("#6DBE45"))
                            _sts_n=ParagraphStyle("sn",fontName="Helvetica",fontSize=8,textColor=rl_colors.HexColor("#6B7280"))
                            _stdia_n=ParagraphStyle("dn",fontName="Helvetica-Bold",fontSize=9,textColor=rl_colors.white,alignment=TA_CENTER)
                            _stcom_n=ParagraphStyle("cn",fontName="Helvetica-Bold",fontSize=7.5,textColor=rl_colors.HexColor("#6DBE45"),leading=10)
                            _stali_n=ParagraphStyle("an",fontName="Helvetica",fontSize=7,textColor=rl_colors.black,leading=9)
                            _sttot_n=ParagraphStyle("ttn",fontName="Helvetica-Bold",fontSize=7,textColor=rl_colors.HexColor("#333"),leading=9)
                            # Encabezado
                            _story_n.append(Paragraph(f"PUTÚ ACTIVO — Plan Nutricional: {_plan_row['nombre']}",_stt_n))
                            _story_n.append(Paragraph(
                                f"Cliente: {sv(r,'nombre')} · RUT: {rut} · Profesional: {_plan_row.get('profesional') or '—'} · "
                                f"Objetivo: {_plan_row.get('objetivo') or '—'} · Meta: {int(_plan_row.get('kcal_objetivo') or 0):,} kcal/día · "
                                f"Generado: {fmt_fecha(str(hoy))}",_sts_n))
                            _story_n.append(Spacer(1,0.3*cm))
                            # Calendario semanal — tabla 7 columnas
                            DIAS_PDF_N=["Lunes","Martes","Miércoles","Jueves","Viernes","Sábado","Domingo"]
                            _ac_n=_aw_n/7
                            _hdr_n=[Paragraph(d,_stdia_n) for d in DIAS_PDF_N]
                            _body_n=[]
                            for _dpn in DIAS_PDF_N:
                                _dcn=_comidas_df[_comidas_df["dia_semana"]==_dpn] if not _comidas_df.empty else pd.DataFrame()
                                _cel=[]
                                if _dcn.empty:
                                    _cel.append(Paragraph("Sin plan",_sts_n))
                                else:
                                    for _,_cmn in _dcn.iterrows():
                                        _icon_c=COMIDAS_ICON.get(_cmn["tipo_comida"],"")
                                        _cel.append(Paragraph(f"{_icon_c} {_cmn['tipo_comida']}",_stcom_n))
                                        _itn=_parse_items(_cmn.get("items"))
                                        for _it_n in _itn:
                                            _cel.append(Paragraph(f"• {_it_n['nombre'][:22]} {_it_n['cantidad_g']}g",_stali_n))
                                        _tot_cm_n={"kcal":sum(i["kcal"] for i in _itn),"prot":sum(i["prot"] for i in _itn)}
                                        if _tot_cm_n["kcal"]>0:
                                            _cel.append(Paragraph(f"{int(_tot_cm_n['kcal'])} kcal · P:{_tot_cm_n['prot']}g",_sttot_n))
                                    _td_n2=_totales_dia(_dpn)
                                    _cel.append(Paragraph(f"TOTAL: {int(_td_n2['kcal'])} kcal",
                                        ParagraphStyle("tt2",fontName="Helvetica-Bold",fontSize=7.5,textColor=rl_colors.HexColor("#6DBE45"))))
                                _body_n.append(_cel)
                            _tbl_n=Table([_hdr_n,[_body_n]],colWidths=[_ac_n]*7)
                            _tbl_n.setStyle(TableStyle([
                                ("BACKGROUND",(0,0),(-1,0),rl_colors.HexColor("#6DBE45")),
                                ("TOPPADDING",(0,0),(-1,0),5),("BOTTOMPADDING",(0,0),(-1,0),5),
                                ("VALIGN",(0,1),(-1,1),"TOP"),("TOPPADDING",(0,1),(-1,1),6),
                                ("LEFTPADDING",(0,0),(-1,-1),4),("RIGHTPADDING",(0,0),(-1,-1),4),
                                ("BOX",(0,0),(-1,-1),0.6,rl_colors.HexColor("#CCCCCC")),
                                ("INNERGRID",(0,0),(-1,-1),0.3,rl_colors.HexColor("#DDDDDD")),
                                ("ROWBACKGROUNDS",(0,1),(-1,-1),[rl_colors.HexColor("#FAFAFA"),rl_colors.HexColor("#F3F3F3")]),
                            ]))
                            _story_n.append(KeepTogether(_tbl_n))
                            _story_n.append(Spacer(1,0.3*cm))
                            # Resumen semanal
                            _tot_sem_n={"kcal":0,"prot":0,"lip":0,"hdc":0,"fibra":0}
                            _dias_n_c=0
                            for _ds2 in DIAS_PDF_N:
                                _td2=_totales_dia(_ds2)
                                if _td2["kcal"]>0: _dias_n_c+=1
                                for k in _tot_sem_n: _tot_sem_n[k]=round(_tot_sem_n[k]+_td2[k],1)
                            _prom2=lambda v: round(v/_dias_n_c,1) if _dias_n_c>0 else 0
                            _resumen_tbl=Table([[
                                Paragraph("Promedio diario semanal",_stcom_n),
                                Paragraph(f"🔥 {_prom2(_tot_sem_n['kcal'])} kcal",_stali_n),
                                Paragraph(f"🥩 {_prom2(_tot_sem_n['prot'])}g proteínas",_stali_n),
                                Paragraph(f"🌾 {_prom2(_tot_sem_n['hdc'])}g carbohidratos",_stali_n),
                                Paragraph(f"🥑 {_prom2(_tot_sem_n['lip'])}g grasas",_stali_n),
                                Paragraph(f"🌿 {_prom2(_tot_sem_n['fibra'])}g fibra",_stali_n),
                            ]],colWidths=[_aw_n/6]*6)
                            _resumen_tbl.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),rl_colors.HexColor("#F0F0F0")),
                                ("TOPPADDING",(0,0),(-1,-1),5),("BOTTOMPADDING",(0,0),(-1,-1),5),
                                ("LEFTPADDING",(0,0),(-1,-1),6),("BOX",(0,0),(-1,-1),0.4,rl_colors.HexColor("#CCCCCC"))]))
                            _story_n.append(_resumen_tbl)
                            _story_n.append(Spacer(1,0.2*cm))
                            _story_n.append(Paragraph(f"Notas: {_plan_row.get('notas') or '—'} · Putú Activo — Centro de Entrenamiento",_sts_n))
                            _pdoc_n.build(_story_n)
                            _pdf_bytes_n=_pbuf_n.getvalue()
                            _pdf_b64_n=base64.b64encode(_pdf_bytes_n).decode()
                            _dn1,_dn2=st.columns(2)
                            _dn1.download_button("⬇️ Descargar PDF",_pdf_bytes_n,
                                f"Nutricion_{sv(r,'nombre').replace(' ','_')}.pdf","application/pdf",
                                key=f"dl_nutr_{rut}",use_container_width=True)
                            _dn2.markdown(f'<a href="data:application/pdf;base64,{_pdf_b64_n}" target="_blank" style="display:block;background:{GRIS2};color:{VERDE};font-weight:700;border:1px solid {VERDE};padding:9px 16px;border-radius:9px;text-decoration:none;text-align:center;font-size:.92rem;margin-top:2px;">🔗 Abrir en ventana</a>',unsafe_allow_html=True)
                            st.markdown(f'<object data="data:application/pdf;base64,{_pdf_b64_n}" type="application/pdf" width="100%" height="600px" style="border:1px solid {GRIS3};border-radius:10px;margin-top:8px;"><p><a href="data:application/pdf;base64,{_pdf_b64_n}" download>Descargar</a></p></object>',unsafe_allow_html=True)
                    else:
                        st.warning("ReportLab no está disponible para generar PDF.")

        # ── TAB EVALUACIÓN ──
        with tab_eval:
            dev=db_query("SELECT * FROM evaluaciones WHERE rut=? ORDER BY fecha",(rut,))
            teh,ten=st.tabs(["📊 Historial","➕ Nueva"])
            with teh:
                if not dev.empty:
                    ult=dev.iloc[-1]; em=st.columns(5)
                    def emv(k,u=""): v=ult.get(k); return f"{float(v):.1f}{u}" if v and str(v) not in ["nan","None","0.0"] else "—"
                    em[0].metric("Peso",emv("peso"," kg")); em[1].metric("IMC",emv("imc"))
                    em[2].metric("Grasa %",emv("grasa_pct"," %")); em[3].metric("Masa Musc.",emv("masa_musc"," %")); em[4].metric("Agua %",emv("agua_pct"," %"))
                    # Gráfico de progreso
                    if len(dev)>1:
                        dev_num=dev.copy()
                        dev_num["fecha"]=dev_num["fecha"].apply(fmt_fecha)
                        _metricas_plot=[c for c in ["peso","imc","grasa_pct","masa_musc","agua_pct"] if c in dev_num.columns and dev_num[c].notna().sum()>1]
                        if _metricas_plot:
                            _met_sel=st.selectbox("Ver progreso de:",_metricas_plot,
                                format_func=lambda x:{"peso":"Peso (kg)","imc":"IMC","grasa_pct":"Grasa %","masa_musc":"Masa Muscular %","agua_pct":"Agua %"}.get(x,x),
                                key="ev_met_sel")
                            _fig_ev=go.Figure()
                            _fig_ev.add_trace(go.Scatter(
                                x=dev_num["fecha"], y=pd.to_numeric(dev_num[_met_sel],errors="coerce"),
                                mode="lines+markers+text",
                                text=pd.to_numeric(dev_num[_met_sel],errors="coerce").round(1).astype(str),
                                textposition="top center",
                                line=dict(color=VERDE,width=2),
                                marker=dict(size=8,color=VERDE)))
                            _fig_ev.update_layout(**PL,height=240,
                                title=f"Progreso: {_met_sel}",
                                xaxis_title="Fecha",yaxis_title="Valor")
                            st.plotly_chart(_fig_ev,use_container_width=True)
                    dev_d=dev.copy(); dev_d["fecha"]=dev_d["fecha"].apply(fmt_fecha)
                    cols_ev=[c for c in ["fecha","peso","imc","grasa_pct","masa_musc","agua_pct","brazos","abdomen","cadera","muslos","observacion"] if c in dev_d.columns]
                    dev_d.insert(0,"N°",range(1,len(dev_d)+1))
                    st.dataframe(dev_d[["N°"]+cols_ev],use_container_width=True,height=200)
                    # WA
                    def _emv_wa(k,u=""): v=ult.get(k); return f"{float(v):.1f}{u}" if v and str(v) not in ["nan","None","0.0",""] else "—"
                    _msg_ev=(f"📋 *Evaluación Putú Activo*%0A👤 *{sv(r,'nombre')}*%0A📅 Fecha: {fmt_fecha(str(ult.get('fecha','')))}%0A"
                        f"⚖️ Peso: *{_emv_wa('peso',' kg')}*%0A📊 IMC: *{_emv_wa('imc')}*%0A"
                        f"🔴 Grasa: *{_emv_wa('grasa_pct',' %')}*%0A💪 Masa Musc.: *{_emv_wa('masa_musc',' %')}*%0A"
                        f"💪 ¡Sigue entrenando! Putú Activo.")
                    st.markdown(f'<a href="{wa_url(sv(r,"celular"),_msg_ev)}" target="_blank" style="display:inline-block;background:#25D366;color:white;font-weight:700;padding:7px 18px;border-radius:8px;text-decoration:none;font-size:.92rem;margin-top:6px;">📲 Compartir evaluación por WhatsApp</a>',unsafe_allow_html=True)
                    for _vi2,_vr2 in dev.reset_index(drop=True).iterrows():
                        if st.button("🗑️",key=f"del_ev_{_vr2.get('id',_vi2)}",help="Eliminar"):
                            db_exec("DELETE FROM evaluaciones WHERE id=?",(int(_vr2["id"]),)); st.cache_data.clear(); st.rerun()
                else:
                    st.markdown('<div class="info-box">Sin evaluaciones.</div>',unsafe_allow_html=True)
            with ten:
                with st.form("fev"):
                    ev1,ev2=st.columns(2)
                    fev=ev1.date_input("Fecha",value=hoy,format="DD/MM/YYYY")
                    peso_v=ev2.number_input("Peso (kg)",0.0,300.0,0.0,0.1)
                    ev3,ev4,ev5=st.columns(3)
                    talla_ev=ev3.number_input("Talla (cm)",0.0,250.0,0.0,0.5)
                    imc_v=ev4.number_input("IMC",0.0,100.0,0.0,0.1)
                    gr_v=ev5.number_input("Grasa %",0.0,100.0,0.0,0.1)
                    ev6,ev7,ev8=st.columns(3)
                    mm_v=ev6.number_input("Masa Musc. %",0.0,100.0,0.0,0.1)
                    ag_v=ev7.number_input("Agua %",0.0,100.0,0.0,0.1)
                    bz_v=ev8.number_input("Brazos (cm)",0.0,100.0,0.0,0.1)
                    ev9,ev10,ev11=st.columns(3)
                    ab_v=ev9.number_input("Abdomen (cm)",0.0,200.0,0.0,0.1)
                    ca_v=ev10.number_input("Cadera (cm)",0.0,200.0,0.0,0.1)
                    mu_v=ev11.number_input("Muslos (cm)",0.0,200.0,0.0,0.1)
                    obs_ev=st.text_area("Observaciones",height=60)
                    ok_ev=st.form_submit_button("💾 Registrar evaluación",type="primary",use_container_width=True)
                if ok_ev:
                    db_exec("INSERT INTO evaluaciones (rut,fecha,peso,talla,imc,grasa_pct,masa_musc,agua_pct,brazos,abdomen,cadera,muslos,observacion) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)",
                        (rut,str(fev),peso_v or None,talla_ev or None,imc_v or None,gr_v or None,mm_v or None,ag_v or None,bz_v or None,ab_v or None,ca_v or None,mu_v or None,obs_ev or None))
                    st.success("✅ Evaluación registrada."); st.rerun()

        # ── TAB PAGOS ──
        with tab_pagos:
            dp_h=db_query("SELECT p.*,c.nombre FROM pagos p LEFT JOIN clientes c ON c.rut=p.rut WHERE p.rut=? ORDER BY p.fecha DESC",(rut,))
            if st.session_state.get("pago_ok_msg"):
                st.markdown(f'<div class="success-box">{st.session_state.pop("pago_ok_msg")}</div>',unsafe_allow_html=True)
            if not dp_h.empty:
                dp_h["Monto"]=dp_h["monto"].apply(lambda x:f"${int(x):,}")
                dp_h["fecha"]=dp_h["fecha"].apply(fmt_fecha)
                _hdr_pg=st.columns([1.5,1.5,1.5,1.5,1.5,.5])
                for _htxt,_hcol in zip(["Fecha","Monto","Período","Plan","Medio",""],_hdr_pg):
                    _hcol.markdown(f"<span style='color:{VERDE};font-weight:700;font-size:.82rem'>{_htxt}</span>",unsafe_allow_html=True)
                for _pi,_pr in dp_h.reset_index(drop=True).iterrows():
                    _pc=st.columns([1.5,1.5,1.5,1.5,1.5,.5])
                    _pc[0].markdown(f"<span style='font-size:.82rem'>{_pr['fecha']}</span>",unsafe_allow_html=True)
                    _pc[1].markdown(f"<span style='font-size:.82rem;color:{VERDE}'>{_pr['Monto']}</span>",unsafe_allow_html=True)
                    _pc[2].markdown(f"<span style='font-size:.82rem;color:{GRIS_T}'>{_pr.get('concepto','')}</span>",unsafe_allow_html=True)
                    _pc[3].markdown(f"<span style='font-size:.82rem;color:{GRIS_T}'>{_pr.get('tipo_plan','')}</span>",unsafe_allow_html=True)
                    _pc[4].markdown(f"<span style='font-size:.82rem;color:{GRIS_T}'>{_pr.get('medio_pago','')}</span>",unsafe_allow_html=True)
                    if _pc[5].button("🗑️",key=f"del_pag_f_{_pr.get('id',_pi)}",help="Eliminar"):
                        db_exec("DELETE FROM pagos WHERE id=?",(int(dp_h.iloc[_pi]["id"]),)); st.cache_data.clear(); st.rerun()
            else:
                st.markdown('<div class="info-box">Sin pagos registrados.</div>',unsafe_allow_html=True)

        # ── TAB QR ──
        with tab_qr:
            qr_data=f"PUTU|{rut}|{sv(r,'nombre')}"
            qr_b64=generar_qr_b64(qr_data)
            st.markdown(f'<div style="text-align:center"><img src="data:image/png;base64,{qr_b64}" width="220"><br><span style="color:{GRIS_T};font-size:.85rem">{rut}</span></div>',unsafe_allow_html=True)
            st.download_button("⬇️ Descargar QR",base64.b64decode(qr_b64),f"qr_{rut}.png","image/png",use_container_width=True)

        # ── TAB DOCUMENTOS ──
        with tab_doc:
            _sub_tab1,_sub_tab2=st.tabs(["📄 Contrato","📋 Derecho a Saber"])
            with _sub_tab1:
                fhs2=hoy.strftime("%d/%m/%Y")
                _td2=lambda t,v: f'<td style="padding:5px 8px;border:1px solid #ccc;color:#111;font-size:.85rem">{v}</td>'
                _th2=lambda t: f'<td style="padding:5px 8px;border:1px solid #ddd;background:#f5f5f5;color:#111;font-weight:700;font-size:.85rem">{t}</td>'
                st.markdown(f"""<div style="background:white;color:#111;padding:22px;border-radius:10px;font-family:Arial;font-size:.87rem;line-height:1.7;border:1px solid #ddd;max-height:500px;overflow-y:auto">
                  <div style="text-align:center;border-bottom:3px solid #6DBE45;padding-bottom:8px;margin-bottom:12px"><b style="font-size:1.1rem;color:#111">PUTÚ ACTIVO — CONTRATO DE PRESTACIÓN DE SERVICIOS</b></div>
                  <p style="color:#111">En Putú, a <b>{fhs2}</b>, entre <b>Putú Activo</b> y <b>{sv(r,"nombre")}</b>, RUT <b>{rut}</b>.</p>
                  <table style="width:100%;border-collapse:collapse;margin:8px 0">
                    <tr>{_th2("Plan")}{_td2("p",sv(r,"tipo_plan"))}{_th2("Frecuencia")}{_td2("f",sv(r,"frecuencia"))}</tr>
                    <tr>{_th2("Valor")}{_td2("v","${:,}".format(int(float(sv(r,"valor_plan") or 0))))}{_th2("Horario")}{_td2("h",sv(r,"horario"))}</tr>
                    <tr>{_th2("Inicio")}{_td2("i",fmt_fecha(sv(r,"fecha_inscripcion")))}{_th2("Vencimiento")}{_td2("vc",fmt_fecha(sv(r,"fecha_vencimiento")))}</tr>
                  </table>
                  <p style="color:#111"><b>1.</b> El Centro proveerá servicios de entrenamiento según plan contratado.</p>
                  <p style="color:#111"><b>2.</b> El cliente asistirá en horarios pactados y respetará normas de convivencia e higiene.</p>
                  <p style="color:#111"><b>3. Salud:</b> Declara estar apto físicamente. Condición médica: <b>{sv(r,"enfermedad") or "Sin antecedentes"}</b>. Restricciones: <b>{sv(r,"restricciones") or "Ninguna"}</b>.</p>
                  <p style="color:#111"><b>4.</b> Vigencia hasta vencimiento. Se renueva al registrar pago. Datos tratados según Ley N°19.628.</p>
                  <div style="display:flex;justify-content:space-between;margin-top:20px">
                    <div style="border-top:1px solid #333;width:44%;text-align:center;padding-top:4px;color:#111">Firma Cliente<br><b>{sv(r,"nombre")}</b></div>
                    <div style="border-top:1px solid #333;width:44%;text-align:center;padding-top:4px;color:#111">Firma Representante<br><b>Putú Activo</b></div>
                  </div></div>""",unsafe_allow_html=True)
                html_c=f"""<!DOCTYPE html><html><head><meta charset='utf-8'><style>body{{font-family:Arial;font-size:11pt;color:#111;line-height:1.7;padding:20px}}h2{{text-align:center;border-bottom:3px solid #6DBE45;padding-bottom:6px}}table{{width:100%;border-collapse:collapse;margin:8px 0}}td,th{{padding:6px 10px;border:1px solid #bbb;color:#111}}th{{background:#f0f0f0;font-weight:700}}</style></head><body>
<h2>PUTÚ ACTIVO — CONTRATO DE PRESTACIÓN DE SERVICIOS</h2>
<p>En Putú, a <b>{fhs2}</b>, entre <b>Putú Activo</b> y <b>{sv(r,"nombre")}</b>, RUT <b>{rut}</b>.</p>
<table><tr><th>Plan</th><td>{sv(r,"tipo_plan")}</td><th>Frecuencia</th><td>{sv(r,"frecuencia")}</td></tr>
<tr><th>Valor</th><td>${int(float(sv(r,"valor_plan") or 0)):,}</td><th>Horario</th><td>{sv(r,"horario")}</td></tr>
<tr><th>Inicio</th><td>{fmt_fecha(sv(r,"fecha_inscripcion"))}</td><th>Vencimiento</th><td>{fmt_fecha(sv(r,"fecha_vencimiento"))}</td></tr></table>
<p><b>1.</b> Servicios de entrenamiento según plan.</p><p><b>2.</b> Asistir en horarios pactados, respetar normas.</p>
<p><b>3. Salud:</b> {sv(r,"enfermedad") or "Sin antecedentes"}. Restricciones: {sv(r,"restricciones") or "Ninguna"}.</p>
<p><b>4.</b> Vigencia hasta vencimiento. Datos según Ley N°19.628.</p>
<br><br><table style="border:none"><tr>
<td style="border:none;border-top:1px solid #333;width:45%;text-align:center;padding-top:6px">Firma Cliente<br>{sv(r,"nombre")}</td>
<td style="border:none;width:10%"></td>
<td style="border:none;border-top:1px solid #333;width:45%;text-align:center;padding-top:6px">Putú Activo — {fhs2}</td>
</tr></table></body></html>"""
                _dc1c,_dc2c=st.columns(2)
                # Contrato PDF
                if REPORTLAB_OK:
                    _pb_ct=io.BytesIO()
                    _pd_ct=SimpleDocTemplate(_pb_ct,pagesize=A4,leftMargin=1.5*cm,rightMargin=1.5*cm,topMargin=1.5*cm,bottomMargin=1.5*cm)
                    _aw_ct=A4[0]-3*cm; _st_ct=[]
                    _sHc=ParagraphStyle("cH",fontName="Helvetica-Bold",fontSize=14,textColor=rl_colors.HexColor("#6DBE45"),spaceAfter=4,alignment=TA_CENTER)
                    _sSc=ParagraphStyle("cS",fontName="Helvetica-Bold",fontSize=10,textColor=rl_colors.black,spaceBefore=8,spaceAfter=3)
                    _sNc=ParagraphStyle("cN",fontName="Helvetica",fontSize=9,textColor=rl_colors.black,leading=13)
                    _st_ct.append(Paragraph("CENTRO DE ENTRENAMIENTO - PUTÚ ACTIVO",_sHc))
                    _st_ct.append(Paragraph("CONTRATO DE PRESTACIÓN DE SERVICIOS",ParagraphStyle("cS2",fontName="Helvetica-Bold",fontSize=11,textColor=rl_colors.black,alignment=TA_CENTER,spaceAfter=8)))
                    _st_ct.append(Paragraph(f"En Putú, a <b>{fhs2}</b>, entre <b>Putú Activo</b> y <b>{sv(r,'nombre')}</b>, RUT <b>{rut}</b>.",_sNc))
                    _ct_data=[["Plan",sv(r,"tipo_plan"),"Frecuencia",sv(r,"frecuencia")],["Valor",f"${int(float(sv(r,'valor_plan') or 0)):,}","Horario",sv(r,"horario")],["Inicio",fmt_fecha(sv(r,"fecha_inscripcion")),"Vencimiento",fmt_fecha(sv(r,"fecha_vencimiento"))]]
                    _tct=Table(_ct_data,colWidths=[_aw_ct*0.18,_aw_ct*0.32,_aw_ct*0.18,_aw_ct*0.32])
                    _tct.setStyle(TableStyle([("FONTNAME",(0,0),(-1,-1),"Helvetica"),("FONTSIZE",(0,0),(-1,-1),9),("ROWBACKGROUNDS",(0,0),(-1,-1),[rl_colors.HexColor("#F0F0F0"),rl_colors.white]),("GRID",(0,0),(-1,-1),0.3,rl_colors.HexColor("#CCC")),("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4),("LEFTPADDING",(0,0),(-1,-1),6)]))
                    _st_ct.append(Spacer(1,0.2*cm)); _st_ct.append(_tct); _st_ct.append(Spacer(1,0.3*cm))
                    for _cp in [f"<b>1.</b> El Centro proveerá servicios de entrenamiento según plan contratado.",f"<b>2.</b> El cliente asistirá en horarios pactados y respetará normas de convivencia.",f"<b>3. Salud:</b> Condición médica: <b>{sv(r,'enfermedad') or 'Sin antecedentes'}</b>. Restricciones: <b>{sv(r,'restricciones') or 'Ninguna'}</b>.",f"<b>4.</b> Vigencia hasta vencimiento. Datos tratados según Ley N°19.628."]:
                        _st_ct.append(Paragraph(_cp,_sNc))
                    _st_ct.append(Spacer(1,1*cm))
                    _firm=[["Firma Cliente","","Firma Representante"],[sv(r,"nombre"),"",f"Putú Activo — {fhs2}"]]
                    _tf2=Table(_firm,colWidths=[_aw_ct*0.45,_aw_ct*0.1,_aw_ct*0.45])
                    _tf2.setStyle(TableStyle([("FONTNAME",(0,0),(-1,-1),"Helvetica"),("FONTSIZE",(0,0),(-1,-1),9),("LINEABOVE",(0,0),(0,0),0.5,rl_colors.black),("LINEABOVE",(2,0),(2,0),0.5,rl_colors.black),("TOPPADDING",(0,0),(-1,-1),6),("LEFTPADDING",(0,0),(-1,-1),0)]))
                    _st_ct.append(_tf2)
                    _pd_ct.build(_st_ct)
                    _dc1c.download_button("⬇️ Descargar Contrato PDF",_pb_ct.getvalue(),"contrato.pdf","application/pdf",use_container_width=True)
                if _dc2c.button("🖨️ Imprimir Contrato",key="prt_cont",use_container_width=True):
                    st.markdown("<script>setTimeout(function(){ window.print(); }, 300);</script>",unsafe_allow_html=True)
            with _sub_tab2:
                fhsd=hoy.strftime("%d/%m/%Y")
                html_d=f"""<div style="background:white;color:#111;padding:26px;border-radius:10px;font-family:Arial;line-height:1.65;font-size:.85rem;border:1px solid #ddd;max-height:500px;overflow-y:auto">
  <div style="text-align:center;border-bottom:3px solid #6DBE45;padding-bottom:8px;margin-bottom:14px">
    <b style="font-size:1.1rem;color:#111">PUTÚ ACTIVO — DERECHO A SABER</b><br>
    <small style="color:#555">DS N°40/1969 y DS N°594/1999</small>
  </div>
  <table style="width:100%;border-collapse:collapse;margin:8px 0;font-size:.85rem;color:#111">
    <tr><td style="border:1px solid #ddd;padding:5px 8px;width:35%"><b>Nombre</b></td><td style="border:1px solid #ddd;padding:5px 8px">{sv(r,"nombre")}</td></tr>
    <tr><td style="border:1px solid #ddd;padding:5px 8px"><b>RUT</b></td><td style="border:1px solid #ddd;padding:5px 8px">{rut}</td></tr>
    <tr><td style="border:1px solid #ddd;padding:5px 8px"><b>Celular</b></td><td style="border:1px solid #ddd;padding:5px 8px">{fmt_cel(cel)}</td></tr>
    <tr><td style="border:1px solid #ddd;padding:5px 8px"><b>Contacto emergencia</b></td><td style="border:1px solid #ddd;padding:5px 8px">{sv(r,"contacto_emergencia")} · {fmt_cel(sv(r,"celular_emergencia"))}</td></tr>
    <tr><td style="border:1px solid #ddd;padding:5px 8px"><b>Condición médica</b></td><td style="border:1px solid #ddd;padding:5px 8px">{sv(r,"enfermedad")} — {sv(r,"restricciones")}</td></tr>
  </table>
  <p style="color:#111"><b>1. Riesgos:</b> ☑ Caídas ☑ Golpes con máquinas ☑ Lesiones musculares ☑ Fatiga cardiovascular ☑ Riesgo eléctrico ☑ Riesgo acústico</p>
  <p style="color:#111"><b>2. Obligaciones:</b> ☑ Informar condición de salud ☑ Indumentaria adecuada ☑ Seguir instrucciones ☑ Higiene ☑ Orden del equipamiento</p>
  <p style="color:#111"><b>3. Normas:</b> ☑ Prohibido fumar/alcohol/drogas ☑ No niños en sala ☑ No alimentos en sala ☑ Toalla obligatoria ☑ No lanzar pesas ☑ Respetar aforo</p>
  <p style="color:#111"><b>4. Declaración:</b> Declaro haber recibido información sobre riesgos, obligaciones y normas de Putú Activo.</p>
  <div style="display:flex;justify-content:space-between;margin-top:30px">
    <div style="border-top:1px solid #333;width:44%;text-align:center;padding-top:4px;color:#111">Firma encargado — Putú Activo</div>
    <div style="border-top:1px solid #333;width:44%;text-align:center;padding-top:4px;color:#111">{sv(r,"nombre")}<br><small>Firma cliente</small></div>
  </div>
  <p style="color:#111;margin-top:14px">Fecha: <b>{fhsd}</b></p>
</div>"""
                st.markdown(html_d,unsafe_allow_html=True)
                if st.button("🖨️ Imprimir Derecho a Saber",key="prt_der"):
                    st.markdown("<script>setTimeout(function(){ window.print(); }, 300);</script>",unsafe_allow_html=True)

        if st.button("← Volver a la lista",key="vb"): st.session_state.ver_rut=None; st.rerun()

    # ── ROUTER ──
    if st.session_state.ver_rut:
        ficha_cliente(st.session_state.ver_rut)
    else:
        st.markdown('<div class="section-header">👥 Clientes</div>',unsafe_allow_html=True)
        tab_lista, tab_nuevo = st.tabs(["📋 Lista de clientes","➕ Nuevo Cliente"])

        with tab_lista:
            f1,f2,f3=st.columns(3)
            bq=f1.text_input("🔍 Nombre o RUT")
            md=f2.selectbox("Mostrar",["Activos","Todos","Inactivos (consulta)"])
            tf=f3.selectbox("Plan",["Todos"]+sorted(df_cli["tipo_plan"].dropna().unique().tolist()))
            dv=df_act if md=="Activos" else (df_inac if "Inactivos" in md else df_cli)
            if bq.strip(): dv=dv[dv["nombre"].str.contains(bq,case=False,na=False)|dv["rut"].str.contains(bq,case=False,na=False)]
            if tf!="Todos": dv=dv[dv["tipo_plan"]==tf]
            dv=dv.reset_index(drop=True); dv["N°"]=range(1,len(dv)+1)
            st.caption(f"**{len(dv)}** registros · Pincha 👁 para abrir la ficha")
            hd=st.columns([.4,2.2,1.2,.5,.8,1,1,1.1,.4,.4])
            for h,c in zip(["N°","Nombre","RUT","Edad","Sexo","Plan","Estado","Vencimiento","",""],hd):
                c.markdown(f"<span style='color:{VERDE};font-weight:700;font-size:.85rem'>{h}</span>",unsafe_allow_html=True)
            for i,row in dv.iterrows():
                rc=st.columns([.4,2.2,1.2,.5,.8,1,1,1.1,.4,.4])
                ce=VERDE if str(row.get("estado",""))=="Activo" else ROJO
                rc[0].markdown(f"<span style='color:{GRIS_T};font-size:.9rem'>{row['N°']}</span>",unsafe_allow_html=True)
                _es_fem_lista = str(row.get("sexo","")).lower() == "femenino"
                _col_sexo_lista = "#E91E8C" if _es_fem_lista else AZUL
                rc[1].markdown(f"<span style='font-weight:700;color:{BLANCO}'>{row.get('nombre','')}</span>",unsafe_allow_html=True)
                rc[2].markdown(f"<span style='color:{GRIS_T};font-size:.85rem'>{row.get('rut','')}</span>",unsafe_allow_html=True)
                rc[3].markdown(f"<span style='font-size:.88rem'>{int(row['edad']) if pd.notna(row.get('edad')) else ''}</span>",unsafe_allow_html=True)
                rc[4].markdown(f"<span style='color:{_col_sexo_lista};font-size:.85rem;font-weight:600'>{str(row.get('sexo',''))[:3]}</span>",unsafe_allow_html=True)
                rc[5].markdown(f"<span style='color:{GRIS_T};font-size:.85rem'>{row.get('tipo_plan','')}</span>",unsafe_allow_html=True)
                rc[6].markdown(f"<span style='color:{ce};font-weight:700;font-size:.85rem'>{row.get('estado','')}</span>",unsafe_allow_html=True)
                _fv_row=row.get("fecha_vencimiento","")
                _dv_row=dias_para_vencer(_fv_row)
                _venc_txt=fmt_fecha(_fv_row)
                _venc_col=ROJO if (_dv_row is not None and _dv_row<0) else GRIS_T
                _venc_lbl=f"🔴 VENCIDO {_venc_txt}" if (_dv_row is not None and _dv_row<0) else _venc_txt
                rc[7].markdown(f"<span style='color:{_venc_col};font-size:.82rem;font-weight:{"700" if _dv_row is not None and _dv_row<0 else "400"}'>{_venc_lbl}</span>",unsafe_allow_html=True)
                if rc[8].button("👁",key=f"v_{row['rut']}_{i}"):
                    st.session_state.ver_rut=row["rut"]; st.rerun()
                if rc[9].button("🗑️",key=f"del_{row['rut']}_{i}",help="Eliminar cliente"):
                    st.session_state[f"confirm_del_{row['rut']}"]=True; st.rerun()
                if st.session_state.get(f"confirm_del_{row['rut']}"):
                    st.warning(f"¿Eliminar a **{row['nombre']}**? Esta acción no se puede deshacer.")
                    cd1,cd2=st.columns(2)
                    if cd1.button("✅ Sí, eliminar",key=f"yes_{row['rut']}"):
                        db_exec("DELETE FROM clientes WHERE rut=?",(row["rut"],))
                        st.cache_data.clear(); st.session_state.pop(f"confirm_del_{row['rut']}",None); st.rerun()
                    if cd2.button("❌ Cancelar",key=f"no_{row['rut']}"):
                        st.session_state.pop(f"confirm_del_{row['rut']}",None); st.rerun()
        st.markdown("---")
        xls=exportar_todo_excel()
        st.download_button("⬇️ Exportar base completa Excel",xls,"putu_activo_completo.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        # ════ NUEVO CLIENTE (tab dentro de Clientes) ════
        with tab_nuevo:
            with st.form("fnc",clear_on_submit=True):
                st.markdown("**Datos personales**")
                n1,n2=st.columns(2)
                nm=n1.text_input("Nombre *"); rut_n=n2.text_input("RUT * (12345678-9)")
                fn=n1.date_input("Fecha nacimiento",value=None,min_value=date(1920,1,1),format="DD/MM/YYYY")
                sx=n2.selectbox("Sexo",["(seleccionar)","Femenino","Masculino","Otro"])
                dr=n1.text_input("Dirección"); ce=n2.text_input("Celular (56912345678)")
                em=n1.text_input("E-mail"); tl=n2.selectbox("Talla",["XS","S","M","L","XL","XXL"])
                st.markdown("**Contacto emergencia**")
                e1,e2,e3=st.columns(3)
                cen=e1.text_input("Nombre contacto"); cec=e2.text_input("Celular contacto"); cep=e3.text_input("Parentesco")
                st.markdown("**Membresía**")
                m1,m2,m3=st.columns(3)
                fi=m1.date_input("Fecha inscripción / pago",value=hoy,format="DD/MM/YYYY")
                tp=m2.selectbox("Plan",PLANES); fr=m3.selectbox("Frecuencia",FRECUENCIAS)
                ho=m1.selectbox("Horario",HORARIOS); vp=m2.number_input("Valor $ plan",0,10000000,0,500)
                pe=m3.selectbox("Período",PERIODOS,index=2)
                es=m1.selectbox("Estado",["Activo","Inactivo"])
                # Vencimiento calculado desde fecha de pago
                fv_auto=calcular_vencimiento(fi,pe)
                nv1,nv2,nv3=st.columns(3)
                nv1.markdown(f'<div style="background:{GRIS3};border-radius:8px;padding:10px;text-align:center;"><div style="color:#888;font-size:.78rem">VENCIMIENTO CALCULADO</div><div style="color:{VERDE};font-weight:900;font-size:1.1rem">{fmt_fecha(fv_auto)}</div></div>',unsafe_allow_html=True)
                fv_manual_nc=nv2.date_input("✏️ Vencimiento manual (opcional)",value=date.fromisoformat(fv_auto),format="DD/MM/YYYY",key="nc_fvenc")
                obs_venc_nc=nv3.text_input("Motivo ajuste manual",placeholder="Ej: cortesía",key="nc_obs_v")
                # Usar manual si fue modificado o tiene observación
                fv_final_nc=str(fv_manual_nc) if (fv_manual_nc!=date.fromisoformat(fv_auto) or obs_venc_nc.strip()) else fv_auto
                st.markdown("**Días de entrenamiento**")
                dc1,dc2,dc3,dc4,dc5,dc6=st.columns(6)
                d_lu=dc1.checkbox("Día 1"); d_ma=dc2.checkbox("Día 2"); d_mi=dc3.checkbox("Día 3")
                d_ju=dc4.checkbox("Día 4"); d_vi=dc5.checkbox("Día 5"); d_sa=dc6.checkbox("Día 6")
                st.markdown("**Salud & objetivo**")
                s1,s2,s3=st.columns(3)
                enf=s1.text_input("Condición médica"); rst=s2.text_input("Restricciones")
                obj=s3.selectbox("Objetivo",OBJETIVOS); niv=s1.selectbox("Nivel",NIVELES)
                ok=st.form_submit_button("💾 Registrar cliente")
            if ok:
                if not nm.strip() or not rut_n.strip():
                    st.markdown('<div class="alert-box">⚠️ Nombre y RUT son obligatorios.</div>',unsafe_allow_html=True)
                elif not fn:
                    st.markdown('<div class="alert-box">⚠️ Selecciona la fecha de nacimiento.</div>',unsafe_allow_html=True)
                elif sx=="(seleccionar)":
                    st.markdown('<div class="alert-box">⚠️ Selecciona el sexo.</div>',unsafe_allow_html=True)
                else:
                    edad_n=int((hoy-fn).days/365.25)
                    mes_n=mes_de_nacimiento(fn)
                    mc_n=msg_cumpleanos(nm); mv_n=msg_vencimiento(nm,fv_final_nc); mr_n=msg_renovacion(nm,fv_final_nc)
                    dias_lu="Lu" if d_lu else ""; dias_ma="Ma" if d_ma else ""
                    dias_mi="Mi" if d_mi else ""; dias_ju="Ju" if d_ju else ""
                    dias_vi="Vi" if d_vi else ""; dias_sa="Sá" if d_sa else ""
                    try:
                        db_exec("""INSERT INTO clientes (nombre,rut,fecha_nacimiento,edad,mes_cumpleanos,sexo,
                            direccion,celular,email,contacto_emergencia,celular_emergencia,parentesco,
                            fecha_inscripcion,periodo_vencimiento,tipo_plan,frecuencia,horario,valor_plan,
                            fecha_vencimiento,estado,enfermedad,restricciones,objetivo,nivel,talla,
                            lunes,martes,miercoles,jueves,viernes,sabado,
                            mensaje_cumpleanos,mensaje_vencimiento,mensaje_renovacion,creado,modificado)
                            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                            (nm.strip(),rut_n.strip(),str(fn),edad_n,mes_n,sx,dr,fmt_cel(ce),em,
                             cen,fmt_cel(cec),cep,str(fi),pe,tp,fr,ho,vp,fv_final_nc,es,enf,rst,obj,niv,tl,
                             dias_lu,dias_ma,dias_mi,dias_ju,dias_vi,dias_sa,
                             mc_n,mv_n,mr_n,datetime.now().isoformat(),datetime.now().isoformat()))
                        st.cache_data.clear()
                        # Mensaje de bienvenida WA
                        _cel_nc=fmt_cel(ce)
                        _msg_wa_nc=f"Hola *{nm}* 👋 ¡Bienvenido/a a Putú Activo! 🏋️%0ATu inscripción fue registrada exitosamente.%0A📅 Plan: *{tp}* · Vence: *{fmt_fecha(fv_final_nc)}*%0A¡Nos vemos entrenando! 💪"
                        _url_wa_nc=wa_url(_cel_nc,_msg_wa_nc)
                        st.session_state["_nc_wa_url"]=_url_wa_nc
                        st.session_state["_nc_wa_nom"]=nm
                        # Ir directo a Pagos con el cliente pre-cargado + valores de la inscripción
                        st.session_state.pago_rut=rut_n.strip()
                        st.session_state["_nc_plan"]=tp
                        st.session_state["_nc_fr"]=fr
                        st.session_state["_nc_pe"]=pe
                        st.session_state["_nc_monto"]=vp
                        st.session_state["_nc_msg"]=f"✅ Cliente **{nm}** registrado. Vencimiento: {fmt_fecha(fv_final_nc)} · Ahora registra el primer pago:"
                        st.session_state._goto="💳 Pagos y Renovaciones"
                        st.rerun()
                    except Exception as ex: st.error(f"Error: {ex}")


# ════════════════════════════════════════════════════════════════════════════
# ✅ ASISTENCIA — Teclado numérico grande + Display aeropuerto
# ════════════════════════════════════════════════════════════════════════════
elif pagina=="💳 Pagos y Renovaciones":
    st.markdown('<div class="section-header">💳 Pagos y Renovaciones</div>',unsafe_allow_html=True)
    # Botón volver
    if st.button("← Volver",key="pr_volver"): st.session_state._goto="🏠 Dashboard"; st.rerun()
    # Mensaje si viene de Nuevo Cliente
    if st.session_state.get("_nc_msg"):
        _nc_wa=st.session_state.pop("_nc_wa_url","#")
        _nc_nom=st.session_state.pop("_nc_wa_nom","")
        _nc_txt=st.session_state.pop("_nc_msg","")
        st.markdown(f'''<div class="success-box" style="font-size:1rem;">{_nc_txt}
          &nbsp;&nbsp;<a href="{_nc_wa}" target="_blank"
          style="background:#25D366;color:white;padding:7px 18px;border-radius:8px;
          text-decoration:none;font-weight:700;font-size:.95rem;">
          📲 Enviar WhatsApp de Bienvenida a {_nc_nom}</a></div>''',unsafe_allow_html=True)
    tr,th=st.tabs(["➕ Registrar","📜 Historial cliente"])
    with tr:
        if st.session_state.get("pago_ok_msg_pr"):
            _rpr=st.session_state.get("pago_ok_rut_pr",""); _wpr=st.session_state.get("pago_ok_wa_pr","#")
            st.markdown(f'''<div class="success-box" style="font-size:1.05rem;">{st.session_state["pago_ok_msg_pr"]}
              &nbsp;&nbsp;<a href="{_wpr}" target="_blank" style="color:#25D366;font-weight:700;">📲 WhatsApp</a></div>''',unsafe_allow_html=True)
            c_vpr1,c_vpr2=st.columns(2)
            if c_vpr1.button("← Ir a ficha del cliente",key="pr_ok_volver",use_container_width=True):
                st.session_state.ver_rut=_rpr; st.session_state._goto="👥 Clientes"
                for k in ["pago_ok_msg_pr","pago_ok_wa_pr","pago_ok_rut_pr"]: st.session_state.pop(k,None)
                st.rerun()
            if c_vpr2.button("✕ Cerrar",key="pr_ok_close",use_container_width=True):
                for k in ["pago_ok_msg_pr","pago_ok_wa_pr","pago_ok_rut_pr"]: st.session_state.pop(k,None)
                st.rerun()
        # Auto-búsqueda si viene de ficha con pago_rut
        _pago_rut_pre = st.session_state.get("pago_rut","") or ""
        bq_pr=st.text_input("Buscar cliente",value=_pago_rut_pre,key="pr_busq")
        # Usar el valor del campo O el pre-cargado desde ficha
        _buscar_con = bq_pr.strip() or _pago_rut_pre.strip()
        cs_pr=None
        if _buscar_con:
            res_pr=db_query("SELECT * FROM clientes WHERE nombre LIKE ? OR rut LIKE ? ORDER BY nombre",(f"%{_buscar_con}%",f"%{_buscar_con}%"))
            res_pr=res_pr[res_pr["tipo_plan"].str.upper().str.strip()!="PASE DIARIO"] if not res_pr.empty else res_pr
            if not res_pr.empty:
                opts_pr=[f"{r['nombre']} — {r['rut']}" for _,r in res_pr.iterrows()]
                sel_pr=st.selectbox("Cliente",opts_pr,key="pr_sel")
                cs_pr=res_pr.iloc[opts_pr.index(sel_pr)]
                # Limpiar pago_rut una vez cargado para no interferir
                if _pago_rut_pre: st.session_state.pago_rut=None
        if cs_pr is not None:
            st.markdown(f'<div style="background:{GRIS2};border-radius:10px;padding:10px 14px;margin:6px 0;"><b style="color:{VERDE}">{cs_pr["nombre"]}</b> · {cs_pr["rut"]} · Plan: <b>{cs_pr["tipo_plan"]}</b> · {cs_pr["frecuencia"]} · Vence: <b>{fmt_fecha(cs_pr["fecha_vencimiento"])}</b></div>',unsafe_allow_html=True)
            # Pre-cargar valores desde nuevo cliente si vienen de ahí
            _pr_plan_def = st.session_state.pop("_nc_plan", cs_pr["tipo_plan"])
            _pr_fr_def   = st.session_state.pop("_nc_fr",   cs_pr["frecuencia"])
            _pr_pe_def   = st.session_state.pop("_nc_pe",   str(cs_pr.get("periodo_vencimiento","Mensual")))
            _pr_monto_def= int(st.session_state.pop("_nc_monto", int(float(cs_pr.get("valor_plan") or 0))))
            with st.form("fp_pr"):
                p1,p2=st.columns(2)
                fp_pr=p1.date_input("Fecha pago",value=hoy,format="DD/MM/YYYY")
                mo_pr=p2.number_input("Monto $",0,10000000,_pr_monto_def,500)
                p3,p4=st.columns(2)
                tp_pr=p3.selectbox("Plan",PLANES,index=PLANES.index(_pr_plan_def) if _pr_plan_def in PLANES else 0)
                fr_pr=p4.selectbox("Frecuencia",FRECUENCIAS,index=FRECUENCIAS.index(_pr_fr_def) if _pr_fr_def in FRECUENCIAS else 0)
                p5,p6=st.columns(2)
                _per_opts=["Diario","Semanal","Quincenal","Mensual","Bimensual","Trimestral","Semestral","Anual"]
                conc_pr=p5.selectbox("Período",_per_opts,index=_per_opts.index(_pr_pe_def) if _pr_pe_def in _per_opts else 3)
                med_pr=p6.selectbox("Medio de pago",["Efectivo","Transferencia","Débito","Crédito","Otro"])
                p7pr,p8pr=st.columns(2)
                desc_pr=p7pr.number_input("% Descuento",0,100,0,1)
                obs_pr=p8pr.text_input("Observación")
                mo_pr_final=int(mo_pr*(1-desc_pr/100)) if desc_pr>0 else mo_pr
                if desc_pr>0: p7pr.markdown(f'<div style="color:{VERDE};font-size:.85rem">Con descuento: <b>${mo_pr_final:,}</b></div>',unsafe_allow_html=True)
                ok_pr=st.form_submit_button("💾 Registrar pago",use_container_width=True)
            if ok_pr:
                db_exec("INSERT INTO pagos (rut,nombre,fecha,monto,concepto,tipo_plan,frecuencia,medio_pago,observacion,usuario) VALUES (?,?,?,?,?,?,?,?,?,?)",
                    (str(cs_pr["rut"]),str(cs_pr["nombre"]),str(fp_pr),mo_pr_final,conc_pr,tp_pr,fr_pr,med_pr,obs_pr,st.session_state.nombre_u))
                nv_pr=calcular_vencimiento(fp_pr,conc_pr)
                c2=get_conn()
                c2.execute("""UPDATE clientes SET
                    fecha_vencimiento=?,fecha_renovacion=?,
                    tipo_plan=?,frecuencia=?,valor_plan=?,
                    periodo_vencimiento=?,modificado=? WHERE rut=?""",
                    (nv_pr,str(fp_pr),tp_pr,fr_pr,mo_pr_final,conc_pr,
                     datetime.now().isoformat(),str(cs_pr["rut"])))
                c2.commit(); c2.close()
                st.cache_data.clear()
                st.cache_data.clear()
                mok_pr=f"Hola *{cs_pr['nombre']}*, tu pago de ${mo_pr:,} fue recibido en Putú Activo el {fp_pr}. Nuevo vencimiento: {nv_pr}. ¡Gracias! 💪"
                uok_pr=wa_url(cs_pr["celular"],mok_pr)
                st.session_state["pago_ok_msg_pr"]=f"✅ Pago **${mo_pr:,}** registrado. Nuevo vencimiento: **{nv_pr}**"
                st.session_state["pago_ok_wa_pr"]=uok_pr
                st.session_state["pago_ok_rut_pr"]=str(cs_pr["rut"])
                st.session_state.pago_rut=None; st.rerun()
        if st.session_state.get("pago_rut") and not bq_pr.strip():
            st.session_state.pago_rut=None
    with th:
        bh_pr=st.text_input("Buscar",key="pr_hist")
        if bh_pr.strip():
            dp_h=db_query("SELECT * FROM pagos WHERE nombre LIKE ? OR rut LIKE ? ORDER BY fecha DESC",(f"%{bh_pr}%",f"%{bh_pr}%"))
            if not dp_h.empty:
                st.metric("Total pagado",f"${int(dp_h['monto'].sum()):,}")
                dp_h["Monto"]=dp_h["monto"].apply(lambda x:f"${int(x):,}")
                dp_h["fecha"]=dp_h["fecha"].apply(fmt_fecha)
                # Tabla con botón eliminar inline por fila
                for _pi,_pr in dp_h.reset_index(drop=True).iterrows():
                    _pc1,_pc2,_pc3,_pc4,_pc5,_pc6,_pc7=st.columns([1.5,2.5,1.5,1.5,1.5,1.5,.6])
                    _pc1.markdown(f"<span style='font-size:.82rem'>{_pr['fecha']}</span>",unsafe_allow_html=True)
                    _pc2.markdown(f"<span style='font-size:.82rem;font-weight:600'>{_pr.get('nombre','')}</span>",unsafe_allow_html=True)
                    _pc3.markdown(f"<span style='font-size:.82rem;color:{VERDE}'>{_pr['Monto']}</span>",unsafe_allow_html=True)
                    _pc4.markdown(f"<span style='font-size:.78rem;color:{GRIS_T}'>{_pr.get('concepto','')}</span>",unsafe_allow_html=True)
                    _pc5.markdown(f"<span style='font-size:.78rem;color:{GRIS_T}'>{_pr.get('tipo_plan','')}</span>",unsafe_allow_html=True)
                    _pc6.markdown(f"<span style='font-size:.78rem;color:{GRIS_T}'>{_pr.get('medio_pago','')}</span>",unsafe_allow_html=True)
                    if _pc7.button("🗑️",key=f"del_pg_{_pr.get('id',_pi)}",help="Eliminar"):
                        db_exec("DELETE FROM pagos WHERE id=?",(int(_pr["id"]),))
                        st.cache_data.clear(); st.rerun()
            else: st.markdown('<div class="info-box">Sin pagos.</div>',unsafe_allow_html=True)

elif pagina=="✅ Asistencia":
    try:
        limite=(datetime.now()-timedelta(hours=3)).strftime("%H:%M")
        _ca2=get_conn()
        _ca2.execute("UPDATE asistencia SET hora_salida='Auto' WHERE fecha=? AND tipo='ingreso' AND hora_salida IS NULL AND hora<=?",(str(hoy),limite))
        _ca2.commit(); _ca2.close()
    except: pass

    st.markdown('<div class="section-header">✅ Control de Asistencia</div>',unsafe_allow_html=True)
    if st.button("← Volver",key="asist_volver"): st.session_state._goto="🏠 Dashboard"; st.rerun()
    tm,=st.tabs(["📌 Marcar asistencia"])

    with tm:
        if "rut_buf" not in st.session_state: st.session_state.rut_buf=""
        if "asist_ok" not in st.session_state: st.session_state.asist_ok=None
        if "asist_tipo" not in st.session_state: st.session_state.asist_tipo="ingreso"

        st.markdown(f"""<style>
        .rut-big-input input[type="text"] {{
            font-size:2.4rem !important; font-weight:900 !important;
            text-align:center !important; letter-spacing:.2em !important;
            height:80px !important; color:{VERDE} !important;
            background:{GRIS2} !important; border:3px solid {VERDE} !important;
            border-radius:14px !important;
        }}</style>""", unsafe_allow_html=True)

        def procesar_rut_asist(rut_raw, tipo_mov):
            """Registra ingreso o salida. Si ya ingresó hoy, registra salida."""
            rut_u = rut_raw.strip().upper()
            if not rut_u: return
            cli = db_query("SELECT * FROM clientes WHERE UPPER(rut)=?", (rut_u,))
            if cli.empty:
                st.session_state.asist_ok={"ok":False,"nombre":rut_u,"msg":"RUT no encontrado","tipo":tipo_mov}
                return
            c=cli.iloc[0]; es_a=str(c.get("estado","")).capitalize()=="Activo"
            if not es_a:
                st.session_state.asist_ok={"ok":False,"nombre":c["nombre"],"msg":"Usuario bloqueado, consultar en administración","tipo":tipo_mov}
                return
            ahora=datetime.now(); hora_str=ahora.strftime("%H:%M")
            # Verificar si ya tiene ingreso hoy sin salida
            ya_adentro=db_query("SELECT id FROM asistencia WHERE UPPER(rut)=? AND fecha=? AND tipo='ingreso' AND (hora_salida IS NULL OR hora_salida='')",(rut_u,str(hoy)))
            if tipo_mov=="ingreso":
                if not ya_adentro.empty:
                    st.session_state.asist_ok={"ok":False,"nombre":c["nombre"],"msg":"Ya tiene ingreso activo hoy","tipo":"ingreso"}
                    return
                db_exec("INSERT INTO asistencia (rut,nombre,fecha,hora,tipo,usuario) VALUES (?,?,?,?,?,?)",
                    (c["rut"],c["nombre"],str(hoy),hora_str,"ingreso","kiosko"))
                st.session_state.asist_ok={"ok":True,"nombre":c["nombre"],"plan":c["tipo_plan"],"hora":hora_str,"tipo":"ingreso","emoji":"🏋️"}
            else:  # salida
                if ya_adentro.empty:
                    st.session_state.asist_ok={"ok":False,"nombre":c["nombre"],"msg":"No tiene ingreso activo hoy","tipo":"salida"}
                    return
                reg_id=int(ya_adentro.iloc[0]["id"])
                conn2=get_conn(); conn2.execute("UPDATE asistencia SET hora_salida=? WHERE id=?",(hora_str,reg_id)); conn2.commit(); conn2.close()
                st.session_state.asist_ok={"ok":True,"nombre":c["nombre"],"plan":c["tipo_plan"],"hora":hora_str,"tipo":"salida","emoji":"👋"}

        col_pad, col_disp = st.columns([1,1])
        with col_pad:
            # ── Campo teclado físico ARRIBA (siempre visible) ──
            with st.form("form_asist_fisico", clear_on_submit=True):
                rut_fis=st.text_input("RUT (teclado físico / escáner):",
                    placeholder="12345678-9  ó  9876543-K",
                    key="rut_fisico",label_visibility="visible")
                ff1,ff2=st.columns(2)
                ok_ing_f=ff1.form_submit_button("✅ INGRESO", use_container_width=True)
                ok_sal_f=ff2.form_submit_button("🚪 SALIDA",  use_container_width=True)
            if ok_ing_f and rut_fis.strip():
                procesar_rut_asist(rut_fis,"ingreso"); st.rerun()
            if ok_sal_f and rut_fis.strip():
                procesar_rut_asist(rut_fis,"salida"); st.rerun()

            st.markdown("---")

            # ── Teclado numérico en pantalla ──
            rut_display=st.session_state.rut_buf or "_ _ _ _ _ _"
            st.markdown(f'''<div style="font-size:2.6rem;font-weight:900;color:{VERDE};text-align:center;
                background:{GRIS2};border:3px solid {VERDE};border-radius:14px;
                padding:16px;margin:8px 0;letter-spacing:.2em;">{rut_display}</div>''',
                unsafe_allow_html=True)

            nums=[["1","2","3"],["4","5","6"],["7","8","9"],["-","0","⌫"]]
            for fila in nums:
                cols=st.columns(3)
                for i,d in enumerate(fila):
                    if cols[i].button(d,key=f"np_{d}",use_container_width=True):
                        if d=="⌫": st.session_state.rut_buf=st.session_state.rut_buf[:-1]
                        else: st.session_state.rut_buf+=d
                        st.session_state.asist_ok=None; st.rerun()

            ck1,ck2,ck3,ck4=st.columns(4)
            if ck1.button("K",key="np_K",use_container_width=True):
                st.session_state.rut_buf+="K"; st.session_state.asist_ok=None; st.rerun()
            if ck2.button("✅ INGRESO",key="btn_ing",use_container_width=True):
                procesar_rut_asist(st.session_state.rut_buf,"ingreso")
                st.session_state.rut_buf=""; st.rerun()
            if ck3.button("🚪 SALIDA",key="btn_sal",use_container_width=True):
                procesar_rut_asist(st.session_state.rut_buf,"salida")
                st.session_state.rut_buf=""; st.rerun()
            if ck4.button("🗑",key="btn_bor",use_container_width=True):
                st.session_state.rut_buf=""; st.session_state.asist_ok=None; st.rerun()

            # ── Resultado ──
            if st.session_state.asist_ok:
                ao=st.session_state.asist_ok
                if ao["ok"]:
                    tipo_lbl="✅ Ingreso correcto" if ao["tipo"]=="ingreso" else "🚪 Salida registrada"
                    col_r=VERDE if ao["tipo"]=="ingreso" else AZUL
                    st.markdown(f'''<div style="background:{col_r}22;border:2px solid {col_r};border-radius:14px;
                        padding:18px;text-align:center;margin-top:10px;">
                      <div style="font-size:1.8rem">{ao["emoji"]}</div>
                      <div style="font-size:1.4rem;font-weight:900;color:{col_r};margin:6px 0">{ao["nombre"]}</div>
                      <div style="color:{GRIS_T};font-size:.95rem">{ao["plan"]} · {ao["hora"]}</div>
                      <div style="color:{col_r};font-weight:700;margin-top:6px">{tipo_lbl} registrado</div>
                    </div>''',unsafe_allow_html=True)
                else:
                    st.markdown(f'''<div style="background:{ROJO}22;border:2px solid {ROJO};border-radius:14px;
                        padding:18px;text-align:center;margin-top:10px;">
                      <div style="font-size:1.8rem">❌</div>
                      <div style="font-size:1.2rem;font-weight:900;color:{ROJO};margin:6px 0">{ao["msg"]}</div>
                      <div style="color:{GRIS_T}">{ao["nombre"]}</div>
                    </div>''',unsafe_allow_html=True)

        with col_disp:
            # Display aeropuerto: clientes en sala hoy
            asist_hoy=db_query("SELECT * FROM asistencia WHERE fecha=? ORDER BY hora DESC",(str(hoy),))
            clases_hoy=db_query("SELECT * FROM clases WHERE fecha>=? ORDER BY fecha,hora LIMIT 10",(str(hoy),))
            if st.session_state.get("_panel","sala")=="sala":
                st.markdown(f"""<div style="background:{GRIS2};border:1px solid {GRIS3};border-radius:13px;padding:16px;">
                  <div style="color:{VERDE};font-weight:700;font-size:1.1rem;margin-bottom:10px">
                    🏟️ EN SALA HOY — {len(asist_hoy)} personas</div>""",unsafe_allow_html=True)
                if not asist_hoy.empty:
                    if "tipo" in asist_hoy.columns:
                        en_sala=asist_hoy[(asist_hoy["tipo"]=="ingreso")&(asist_hoy["hora_salida"].isna()|asist_hoy["hora_salida"].eq(""))]
                    else:
                        en_sala=asist_hoy
                    st.markdown(f"<b style='color:{VERDE}'>{len(en_sala)} persona(s) en sala</b>",unsafe_allow_html=True)
                    for idx_s,a in en_sala.iterrows():
                        col_n,col_h,col_fic,col_sal,col_del=st.columns([2.3,0.9,0.9,1.0,.5])
                        col_n.markdown(f"<span style='font-weight:700;font-size:1rem'>{a['nombre']}</span>",unsafe_allow_html=True)
                        col_h.markdown(f"<span style='color:{VERDE};font-size:.9rem'>🏋️ {a['hora']}</span>",unsafe_allow_html=True)
                        # Botón SALIDA directo — sin necesidad de ingresar RUT de nuevo
                        if col_fic.button("👤 Ficha",key=f"adm_fic_{a.get('id',idx_s)}",use_container_width=True):
                            st.session_state.ver_rut=str(a.get("rut",""))
                            st.session_state._goto="👥 Clientes"; st.rerun()
                        if col_sal.button("🚪 Salida",key=f"sal_{a.get('id',idx_s)}",use_container_width=True):
                            ahora_s=datetime.now().strftime("%H:%M")
                            conn_sal=get_conn()
                            conn_sal.execute("UPDATE asistencia SET hora_salida=? WHERE id=?",(ahora_s,int(a["id"])))
                            conn_sal.commit(); conn_sal.close()
                            db_query.clear()
                            st.session_state.asist_ok={"ok":True,"nombre":a["nombre"],"plan":"","hora":ahora_s,"tipo":"salida","emoji":"👋"}
                            st.rerun()
                        if st.session_state.get("rol")=="Administrador":
                            if col_del.button("🗑️",key=f"del_a_{a.get('id',idx_s)}",help="Eliminar"):
                                db_exec("DELETE FROM asistencia WHERE id=?",(int(a["id"]),))
                                st.rerun()
                else:
                    st.markdown(f'<div style="color:{GRIS_T};text-align:center;padding:20px">Aún no hay asistencias hoy</div>',unsafe_allow_html=True)
                st.markdown("</div>",unsafe_allow_html=True)
                if st.button("Ver próximas clases →",key="sw_clases"):
                    st.session_state._panel="clases"; st.rerun()
            else:
                st.markdown(f"""<div style="background:{GRIS2};border:1px solid {GRIS3};border-radius:13px;padding:16px;">
                  <div style="color:{VERDE};font-weight:700;font-size:1.1rem;margin-bottom:10px">
                    📅 PRÓXIMAS CLASES & TALLERES</div>""",unsafe_allow_html=True)
                if not clases_hoy.empty:
                    for _,cl in clases_hoy.iterrows():
                        titulo=str(cl.get("titulo","")) or str(cl.get("tipo","Clase"))
                        st.markdown(f"""<div style="background:{GRIS3};border-radius:8px;padding:10px 14px;margin:5px 0;display:flex;justify-content:space-between;align-items:center;">
                          <div><b style="color:{VERDE}">{titulo}</b><br><span style="color:{GRIS_T};font-size:.85rem">{cl['tipo']} · {cl['hora']}</span></div>
                          <div style="text-align:right;color:{GRIS_T};font-size:.85rem">{cl['fecha']}</div></div>""",unsafe_allow_html=True)
                else:
                    st.markdown(f'<div style="color:{GRIS_T};text-align:center;padding:20px">Sin clases próximas agendadas</div>',unsafe_allow_html=True)
                st.markdown("</div>",unsafe_allow_html=True)
                if st.button("← Ver en sala",key="sw_sala"):
                    st.session_state._panel="sala"; st.rerun()



# ════════════════════════════════════════════════════════════════════════════
# 🏃 CLASES & TALLERES
# ════════════════════════════════════════════════════════════════════════════
elif pagina=="🏃 Clases & Talleres":
    st.markdown('<div class="section-header">🏃 Clases & Talleres</div>',unsafe_allow_html=True)
    if st.button("← Volver",key="cls_volver"): st.session_state._goto="🏠 Dashboard"; st.rerun()
    tcc,trc,thc=st.tabs(["➕ Crear clase/taller","👤 Inscribir participante","📋 Historial"])
    with tcc:
        st.markdown("**Crear nueva clase o taller**")
        with st.form("fcl_crear"):
            cr1,cr2=st.columns(2)
            cr_nom=cr1.text_input("Nombre de la clase/taller *",placeholder="Ej: Funcional Adultos, Yoga, Zumba")
            _inst_opts=["Juan Alarcón Saavedra","Vanessa Aguilera Ferrada","Diego Cancino González","Otro..."]
            _inst_sel=cr1.selectbox("Instructor *",_inst_opts,key="cr_inst_sel")
            if _inst_sel=="Otro...":
                cr_inst=cr2.text_input("Escribe el nombre del instructor",key="cr_inst_txt",placeholder="Nombre completo")
            else:
                cr_inst=_inst_sel
                cr2.markdown(f'<div style="padding:8px 4px;font-size:.9rem;color:#6DBE45;font-weight:600">✓ {cr_inst}</div>',unsafe_allow_html=True)
            cr2b,cr3b,cr4b=st.columns(3)
            cr_fc=cr2b.date_input("Fecha",value=hoy,format="DD/MM/YYYY")
            cr_ho=cr3b.time_input("Horario")
            cr_cupos=cr4b.number_input("Cupos máximos",1,200,10,1)
            cr_tipo=cr1.selectbox("Tipo",["Adultos","Niños","Taller especial","Pase diario","Otro"])
            cr_obs=cr2.text_input("Observaciones")
            ok_cr=st.form_submit_button("💾 Crear clase",use_container_width=True)
        if ok_cr and cr_nom.strip():
            db_exec("INSERT INTO clases (tipo,titulo,fecha,hora,participante,monto,observacion,usuario) VALUES (?,?,?,?,?,?,?,?)",
                (cr_tipo,cr_nom,str(cr_fc),str(cr_ho),
                 f"[CLASE] Instructor:{cr_inst} Cupos:{cr_cupos}",0,cr_obs,st.session_state.nombre_u))
            st.markdown(f'<div class="success-box">✅ Clase <b>{cr_nom}</b> creada para el {fmt_fecha(str(cr_fc))} a las {str(cr_ho)[:5]}. Cupos: {cr_cupos}</div>',unsafe_allow_html=True)
    with trc:
        st.markdown("**Inscribir participante en una clase existente**")
        # Buscar clases creadas (tienen [CLASE] en participante)
        clases_creadas=db_query("SELECT * FROM clases WHERE participante LIKE '%[CLASE]%' ORDER BY fecha DESC,hora")
        if clases_creadas.empty:
            st.markdown('<div class="info-box">No hay clases creadas aún. Crea una en la pestaña anterior.</div>',unsafe_allow_html=True)
        else:
            opts_cls=["(seleccionar)"]+[f"{fmt_fecha(r['fecha'])} {str(r['hora'])[:5]} — {r['titulo']}" for _,r in clases_creadas.iterrows()]
            sel_cls=st.selectbox("Seleccionar clase",opts_cls,key="cls_sel_i")
            # Autocompletar fecha/hora/tipo desde la clase seleccionada
            fc2_def=hoy; hc2_def=None; ti2_def=""; tp2_def="Adultos"; cupos_max=10
            if sel_cls!="(seleccionar)":
                idx_c=opts_cls.index(sel_cls)-1
                cls_row=clases_creadas.iloc[idx_c]
                try: fc2_def=date.fromisoformat(str(cls_row["fecha"])[:10])
                except: pass
                ti2_def=str(cls_row.get("titulo",""))
                tp2_def=str(cls_row.get("tipo","Adultos"))
                # Extraer cupos del campo participante
                m_cupos=re.search(r"Cupos:(\d+)",str(cls_row.get("participante","")))
                if m_cupos: cupos_max=int(m_cupos.group(1))
                # Contar ya inscritos
                ya_inscritos=db_query("SELECT COUNT(*) as n FROM clases WHERE titulo=? AND fecha=? AND participante NOT LIKE '%[CLASE]%'",(ti2_def,str(cls_row["fecha"]))).iloc[0]["n"]
                cupos_disp=cupos_max-ya_inscritos
                color_c=VERDE if cupos_disp>5 else NARANJA if cupos_disp>0 else ROJO
                st.markdown(f'<div style="background:{color_c}22;border:1px solid {color_c};border-radius:8px;padding:8px 14px;margin:4px 0;font-size:.92rem;">📋 <b>{ti2_def}</b> · {fmt_fecha(str(cls_row["fecha"]))} · Instructor: {str(cls_row.get("participante","")).replace("[CLASE]","").split("Instructor:")[-1].split("Cupos:")[0].strip()} · <b style="color:{color_c}">Cupos disponibles: {cupos_disp}/{cupos_max}</b></div>',unsafe_allow_html=True)
            # Buscar cliente
            pa_busq=st.text_input("Buscar cliente (nombre o RUT)",placeholder="Escribe para buscar...")
            pa_final_pre=""
            if pa_busq.strip():
                res_pa=db_query("SELECT nombre,rut FROM clientes WHERE nombre LIKE ? OR rut LIKE ? ORDER BY nombre",(f"%{pa_busq}%",f"%{pa_busq}%"))
                if not res_pa.empty:
                    opts_pa=[f"{r['nombre']} — {r['rut']}" for _,r in res_pa.iterrows()]
                    sel_pa=st.selectbox("Cliente encontrado:",opts_pa,key="cls_pa_sel")
                    pa_final_pre=sel_pa.split(" — ")[0]
            with st.form("fcl2",clear_on_submit=True):
                c1,c2=st.columns(2)
                tp2=c1.selectbox("Tipo",["Adultos","Niños","Pase diario","Taller especial","Otro"],
                    index=["Adultos","Niños","Pase diario","Taller especial","Otro"].index(tp2_def) if tp2_def in ["Adultos","Niños","Pase diario","Taller especial","Otro"] else 0)
                ti2=c2.text_input("Título",value=ti2_def)
                c3,c4=st.columns(2)
                fc2=c3.date_input("Fecha",value=fc2_def,format="DD/MM/YYYY")
                # Hora viene de la clase seleccionada; si no hay selección, permite editar
                try:
                    _hc2_parts=str(cls_row["hora"])[:5].split(":") if sel_cls!="(seleccionar)" else ["08","00"]
                    _hc2_def=datetime.strptime(":".join(_hc2_parts[:2]),"%H:%M").time()
                except: _hc2_def=datetime.strptime("08:00","%H:%M").time()
                hc2=c4.time_input("Hora",value=_hc2_def)
                p1,p2,p3=st.columns(3)
                pa_manual=p1.text_input("Participante",value=pa_final_pre,placeholder="Nombre o escribe manual")
                mo2=p2.number_input("Monto $",min_value=0,max_value=10000000,value=None,step=500,placeholder="Ingrese monto")
                ob2=p3.text_input("Obs.")
                ok2=st.form_submit_button("💾 Inscribir",use_container_width=True)
            if ok2:
                pa_reg=(pa_manual or pa_busq).strip()
                if pa_reg:
                    db_exec("INSERT INTO clases (tipo,titulo,fecha,hora,participante,monto,observacion,usuario) VALUES (?,?,?,?,?,?,?,?)",
                        (tp2,ti2,str(fc2),str(hc2),pa_reg,mo2,ob2,st.session_state.nombre_u))
                    # Contar cupos tras inscripción
                    ya2=db_query("SELECT COUNT(*) as n FROM clases WHERE titulo=? AND fecha=? AND participante NOT LIKE '%[CLASE]%'",(ti2,str(fc2))).iloc[0]["n"]
                    st.markdown(f'<div class="success-box">✅ <b>{pa_reg}</b> inscrito/a en <b>{ti2}</b>. Inscritos: {ya2}/{cupos_max}</div>',unsafe_allow_html=True)
                else:
                    st.markdown('<div class="alert-box">⚠️ Ingresa el nombre del participante.</div>',unsafe_allow_html=True)
    with thc:
        dc=db_query("SELECT * FROM clases ORDER BY fecha DESC,hora DESC")
        if not dc.empty:
            hf1,hf2=st.columns(2)
            tf=hf1.selectbox("Filtrar tipo",["Todos","Adultos","Niños","Pase diario","Taller especial","Otro"])
            titulo_f=hf2.text_input("Buscar por título",placeholder="Ej: Yoga, Zumba...")
            hf3,hf4=st.columns(2)
            # Filtro por clase grabada
            _cls_grabadas=db_query("SELECT DISTINCT titulo FROM clases WHERE participante LIKE '%[CLASE]%' ORDER BY titulo")
            _cls_opts=["Todas las clases"]+(_cls_grabadas["titulo"].tolist() if not _cls_grabadas.empty else [])
            _cls_sel=hf3.selectbox("Clase grabada",_cls_opts,key="cls_hist_sel")
            _fecha_cls=hf4.date_input("Filtrar por fecha",value=None,key="cls_hist_fecha")
            solo_inscritos=st.checkbox("Solo inscritos (ocultar clases creadas)",value=True,key="cls_solo_ins")
            dc_f=dc.copy()
            if solo_inscritos: dc_f=dc_f[~dc_f["participante"].str.contains(r"\[CLASE\]",na=False,regex=True)]
            if tf!="Todos": dc_f=dc_f[dc_f["tipo"]==tf]
            if titulo_f.strip(): dc_f=dc_f[dc_f["titulo"].str.contains(titulo_f,case=False,na=False)]
            if _cls_sel!="Todas las clases": dc_f=dc_f[dc_f["titulo"]==_cls_sel]
            if _fecha_cls: dc_f=dc_f[dc_f["fecha"]==str(_fecha_cls)]
            c1,c2=st.columns(2)
            dc_f["monto"]=pd.to_numeric(dc_f["monto"],errors="coerce").fillna(0)
            c1.metric("Total inscritos",len(dc_f))
            # Cabecera tabla
            _hc=st.columns([1.4,1.1,1.8,1.8,1.2,.7])
            for _ht,_hcol in zip(["Fecha","Hora","Clase","Participante","Estado",""],_hc):
                _hcol.markdown(f"<span style='color:{VERDE};font-weight:700;font-size:.8rem'>{_ht}</span>",unsafe_allow_html=True)
            # Tabla inline con botón eliminar + estado pagado/pendiente
            for _ci2,_cr2 in dc_f.reset_index(drop=True).iterrows():
                if "[CLASE]" in str(_cr2.get("participante","")): continue
                _cca,_ccb,_ccc,_ccd,_cce,_ccf=st.columns([1.4,1.1,1.8,1.8,1.2,.7])
                _cca.markdown(f"<span style='font-size:.82rem'>{fmt_fecha(_cr2['fecha'])}</span>",unsafe_allow_html=True)
                _ccb.markdown(f"<span style='font-size:.82rem;color:{GRIS_T}'>{str(_cr2.get('hora',''))[:5]}</span>",unsafe_allow_html=True)
                _ccc.markdown(f"<span style='font-size:.82rem;font-weight:600'>{_cr2.get('titulo','')}</span>",unsafe_allow_html=True)
                _ccd.markdown(f"<span style='font-size:.82rem'>{_cr2.get('participante','')}</span>",unsafe_allow_html=True)
                # Estado pagado/pendiente según monto
                _monto_cls=float(_cr2.get("monto",0) or 0)
                _est_pago="✅ Pagado" if _monto_cls>0 else "⏳ Pendiente"
                _col_pago=VERDE if _monto_cls>0 else NARANJA
                _cce.markdown(f"<span style='font-size:.78rem;color:{_col_pago};font-weight:700'>{_est_pago}</span>",unsafe_allow_html=True)
                if tiene_permiso("clases") and _ccf.button("🗑️",key=f"del_cls_{_cr2.get('id',_ci2)}"):
                    db_exec("DELETE FROM clases WHERE id=?",(int(_cr2["id"]),))
                    st.cache_data.clear(); st.rerun()
            # Exportar listado como HTML imprimible (tamaño carta)
            if st.button("🖨️ Exportar listado PDF carta",key="cls_pdf",use_container_width=True):
                filas="".join([f"<tr><td>{fmt_fecha(r.get('fecha',''))}</td><td>{r.get('hora','')}</td><td>{r.get('tipo','')}</td><td>{r.get('titulo','')}</td><td>{r.get('participante','')}</td><td style='text-align:right'>${int(r.get('monto',0)):,}</td></tr>" for _,r in dc_f.iterrows()])
                total_f=int(dc_f["monto"].sum())
                titulo_rep=f"Clases & Talleres — {tf}" + (f" / {titulo_f}" if titulo_f.strip() else "")
                html_cls=f"""<!DOCTYPE html><html><head><meta charset='utf-8'>
                <style>
                  @page{{size:letter;margin:2cm}}
                  body{{font-family:Arial;font-size:10pt;color:#111}}
                  h2{{color:#6DBE45;border-bottom:2px solid #6DBE45;padding-bottom:6px}}
                  table{{width:100%;border-collapse:collapse;margin-top:12px}}
                  th{{background:#6DBE45;color:black;padding:7px;text-align:left}}
                  td{{padding:6px;border-bottom:1px solid #ddd}}
                  tr:nth-child(even){{background:#f9f9f9}}
                  .total{{font-weight:bold;background:#eee}}
                  .footer{{margin-top:20px;font-size:.85em;color:#888;text-align:center}}
                </style></head><body>
                <h2>PUTÚ ACTIVO — {titulo_rep}</h2>
                <p>Generado: {hoy.strftime('%d/%m/%Y')} &nbsp;|&nbsp; Total registros: {len(dc_f)}</p>
                <table>
                  <tr><th>Fecha</th><th>Hora</th><th>Tipo</th><th>Título</th><th>Participante</th><th>Monto</th></tr>
                  {filas}
                  <tr class='total'><td colspan='5'><b>TOTAL INGRESOS</b></td><td style='text-align:right'><b>${total_f:,}</b></td></tr>
                </table>
                <div class='footer'>Putú Activo — Centro de Entrenamiento · Constitución</div>
                </body></html>"""
                if REPORTLAB_OK:
                    _pb_cl=io.BytesIO()
                    _pd_cl=SimpleDocTemplate(_pb_cl,pagesize=A4,leftMargin=1.5*cm,rightMargin=1.5*cm,topMargin=1.5*cm,bottomMargin=1.5*cm)
                    _aw_cl=A4[0]-3*cm; _st_cl=[]
                    _st_cl.append(Paragraph(f"PUTÚ ACTIVO — {titulo_rep}",ParagraphStyle("clH",fontName="Helvetica-Bold",fontSize=13,textColor=rl_colors.HexColor("#6DBE45"),spaceAfter=4)))
                    _st_cl.append(Paragraph(f"Generado: {hoy.strftime('%d/%m/%Y')} | Registros: {len(dc_f)}",ParagraphStyle("clS",fontName="Helvetica",fontSize=9,textColor=rl_colors.HexColor("#888"),spaceAfter=8)))
                    _cl_hdr=[["Fecha","Hora","Tipo","Título","Participante","Monto"]]
                    _cl_rows=[[fmt_fecha(str(ro.get("fecha",""))),str(ro.get("hora","")),str(ro.get("tipo","")),str(ro.get("titulo",""))[:30],str(ro.get("nombre",""))[:25],f"${int(float(ro.get('monto',0))):,}"] for _,ro in dc_f.iterrows()]
                    _cl_rows.append(["","","","","TOTAL",f"${total_f:,}"])
                    _tcl=Table(_cl_hdr+_cl_rows,colWidths=[_aw_cl*0.13,_aw_cl*0.1,_aw_cl*0.12,_aw_cl*0.25,_aw_cl*0.25,_aw_cl*0.15])
                    _tcl.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),rl_colors.HexColor("#6DBE45")),("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),("FONTSIZE",(0,0),(-1,-1),8),("ROWBACKGROUNDS",(0,1),(-1,-2),[rl_colors.white,rl_colors.HexColor("#F5F5F5")]),("BACKGROUND",(0,-1),(-1,-1),rl_colors.HexColor("#1A1A1A")),("TEXTCOLOR",(0,-1),(-1,-1),rl_colors.white),("FONTNAME",(0,-1),(-1,-1),"Helvetica-Bold"),("GRID",(0,0),(-1,-1),0.3,rl_colors.HexColor("#DDD")),("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4),("LEFTPADDING",(0,0),(-1,-1),5)]))
                    _st_cl.append(_tcl)
                    _pd_cl.build(_st_cl)
                    st.download_button("⬇️ Descargar PDF clases",_pb_cl.getvalue(),f"clases_{hoy}.pdf","application/pdf",use_container_width=True,key="cls_dl_pdf")
        else: st.markdown('<div class="info-box">Sin registros.</div>',unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════════════════════
# 💰 INGRESOS & EGRESOS
# ════════════════════════════════════════════════════════════════════════════
elif pagina=="🛍 Venta Productos":
    st.markdown('<div class="section-header">🛍 Venta de Productos</div>',unsafe_allow_html=True)
    if st.button("← Volver",key="vp_volver"): st.session_state._goto="🏠 Dashboard"; st.rerun()
    tv1,tv2=st.tabs(["➕ Registrar venta","📋 Historial"])
    with tv1:
        with st.form("fvp_main"):
            vp1,vp2,vp3=st.columns(3)
            vfp=vp1.date_input("Fecha",value=hoy,key="vpm_f",format="DD/MM/YYYY")
            # Lista de productos con opción de escribir
            _prod_opts=["Agua c/gas","Agua S/gas","Powerade","Redbull","Score","Barra protein","Alfajor","Chocolate","Otro..."]
            _prod_sel=vp2.selectbox("Producto",_prod_opts,key="vpm_psel")
            if _prod_sel=="Otro...":
                vpn=vp3.text_input("Escribe el producto",key="vpm_p",placeholder="Nombre del producto")
            else:
                vpn=_prod_sel
                vp3.markdown(f'<div style="padding:8px 4px;font-size:.9rem;color:#6DBE45;font-weight:600">✓ {vpn}</div>',unsafe_allow_html=True)
            vp1b,vp2b,vp3b=st.columns(3)
            vmp=vp1b.number_input("Monto $",min_value=0,max_value=10000000,value=None,step=500,key="vpm_m",placeholder="Ingrese monto")
            _med_vp=vp2b.selectbox("Medio de pago",["Efectivo","Transferencia","Débito","Crédito","Otro"],key="vpm_med")
            vpu=vp3b.text_input("Cliente (opcional)",key="vpm_c")
            ok_vp=st.form_submit_button("💾 Registrar venta")
        _monto_vp=vmp if vmp is not None else 0
        if ok_vp and vpn.strip():
            _prod_final=vpn.strip()
            db_exec("INSERT INTO productos (fecha,producto,monto,usuario) VALUES (?,?,?,?)",
                (str(vfp),f"{_prod_final} [{_med_vp}]",int(_monto_vp),st.session_state.nombre_u))
            st.markdown(f'<div class="success-box">✅ Venta de <b>{vpn}</b> por ${vmp:,} registrada.</div>',unsafe_allow_html=True)
    with tv2:
        hv1,hv2=st.columns(2)
        mes_vp=hv1.selectbox("Mes",range(1,13),index=hoy.month-1,
            format_func=lambda x:MESES_ESP[x-1].capitalize(),key="vpm_mes")
        anio_vp=hv2.number_input("Año",2024,2030,hoy.year,key="vpm_anio")
        fm_vp=f"{anio_vp}-{mes_vp:02d}"
        dvp=db_query("SELECT * FROM productos WHERE fecha LIKE ? ORDER BY fecha DESC",(f"{fm_vp}%",))
        if not dvp.empty:
            st.metric(f"Total ventas {fm_vp}",f"${int(dvp['monto'].sum()):,}")
            dvp_disp=dvp.copy()
            dvp_disp["fecha"]=dvp_disp["fecha"].apply(fmt_fecha)
            dvp_disp["Monto"]=dvp_disp["monto"].apply(lambda x:f"${int(x):,}")
            dvp_disp=dvp_disp.reset_index(drop=True)
            # Tabla inline con botón eliminar en cada fila
            _hdr=st.columns([.4,1.5,2.5,1.2,1.5,.5])
            for _h,_c in zip(["N°","Fecha","Producto","Monto","Vendedor",""],_hdr):
                _c.markdown(f"<span style='color:{VERDE};font-weight:700;font-size:.82rem'>{_h}</span>",unsafe_allow_html=True)
            for _vi,_vr in dvp_disp.iterrows():
                _rc=st.columns([.4,1.5,2.5,1.2,1.5,.5])
                _rc[0].markdown(f"<span style='font-size:.82rem;color:{GRIS_T}'>{_vi+1}</span>",unsafe_allow_html=True)
                _rc[1].markdown(f"<span style='font-size:.82rem'>{_vr['fecha']}</span>",unsafe_allow_html=True)
                _rc[2].markdown(f"<span style='font-size:.82rem;font-weight:600'>{_vr['producto']}</span>",unsafe_allow_html=True)
                _rc[3].markdown(f"<span style='font-size:.82rem;color:{VERDE}'>{_vr['Monto']}</span>",unsafe_allow_html=True)
                _rc[4].markdown(f"<span style='font-size:.82rem;color:{GRIS_T}'>{_vr.get('usuario','')}</span>",unsafe_allow_html=True)
                if tiene_permiso("reportes") and _rc[5].button("🗑️",key=f"del_vp_{_vr.get('id',_vi)}",help="Eliminar"):
                    db_exec("DELETE FROM productos WHERE id=?",(int(dvp.iloc[_vi]["id"]),))
                    st.cache_data.clear(); st.rerun()
            # Exportar PDF resumen
            pdf_rows="".join([f"<tr><td>{fmt_fecha(r['fecha'])}</td><td>{r['producto']}</td><td style='text-align:right'>${int(r['monto']):,}</td><td>{r['usuario']}</td></tr>" for _,r in dvp.iterrows()])
            html_vp=f"""<html><body style='font-family:Arial;font-size:11pt;padding:20px'>
            <h2 style='color:#6DBE45'>PUTÚ ACTIVO — Ventas de Productos {fm_vp}</h2>
            <table border='1' cellpadding='6' style='border-collapse:collapse;width:100%'>
            <tr style='background:#6DBE45;color:black'><th>Fecha</th><th>Producto</th><th>Monto</th><th>Vendedor</th></tr>
            {pdf_rows}
            <tr style='background:#eee'><td colspan='2'><b>TOTAL</b></td><td style='text-align:right'><b>${int(dvp['monto'].sum()):,}</b></td><td></td></tr>
            </table></body></html>"""
            if REPORTLAB_OK:
                _pb_vp=io.BytesIO()
                _pd_vp=SimpleDocTemplate(_pb_vp,pagesize=A4,leftMargin=1.5*cm,rightMargin=1.5*cm,topMargin=1.5*cm,bottomMargin=1.5*cm)
                _aw_vp=A4[0]-3*cm; _st_vp=[]
                _st_vp.append(Paragraph(f"PUTÚ ACTIVO — Ventas de Productos {fm_vp}",ParagraphStyle("vH",fontName="Helvetica-Bold",fontSize=13,textColor=rl_colors.HexColor("#6DBE45"),spaceAfter=6)))
                _vp_hdr=[["Fecha","Producto","Monto","Vendedor"]]
                _vp_rows=[[fmt_fecha(str(r2.get("fecha",""))),str(r2.get("producto",""))[:30],f"${int(float(r2.get('monto',0))):,}",str(r2.get("usuario",""))] for _,r2 in dvp.iterrows()]
                _vp_rows.append(["","TOTAL",f"${int(dvp['monto'].sum()):,}",""])
                _tvp=Table(_vp_hdr+_vp_rows,colWidths=[_aw_vp*0.2,_aw_vp*0.4,_aw_vp*0.2,_aw_vp*0.2])
                _tvp.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),rl_colors.HexColor("#6DBE45")),("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),("FONTSIZE",(0,0),(-1,-1),9),("ROWBACKGROUNDS",(0,1),(-1,-2),[rl_colors.white,rl_colors.HexColor("#F5F5F5")]),("BACKGROUND",(0,-1),(-1,-1),rl_colors.HexColor("#1A1A1A")),("TEXTCOLOR",(0,-1),(-1,-1),rl_colors.white),("FONTNAME",(0,-1),(-1,-1),"Helvetica-Bold"),("GRID",(0,0),(-1,-1),0.3,rl_colors.HexColor("#DDD")),("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4),("LEFTPADDING",(0,0),(-1,-1),6)]))
                _st_vp.append(_tvp)
                _pd_vp.build(_st_vp)
                st.download_button("⬇️ Descargar PDF ventas",_pb_vp.getvalue(),f"ventas_{fm_vp}.pdf","application/pdf",use_container_width=True)
        else:
            st.markdown(f'<div class="info-box">Sin ventas en {fm_vp}.</div>',unsafe_allow_html=True)

elif pagina=="📊 Reportes":
    st.markdown('<div class="section-header">📊 Reportes</div>',unsafe_allow_html=True)
    if st.button("← Volver",key="rep_volver"): st.session_state._goto="🏠 Dashboard"; st.rerun()
    ta,ti,tv,tc2,tfc,tast,tven,tie,tseg=st.tabs(["✅ Activos","🔴 Inactivos","⚠️ Vencimientos","🎂 Cumpleaños","💰 Flujo de caja","📊 Asistencias","🛍 Ventas","💵 Ingresos & Egresos","👥 Seguimiento"])
    with ta:
        st.markdown(f"**{len(df_act)} clientes activos**")
        c1,c2=st.columns(2)
        with c1:
            od=df_act["objetivo"].dropna().value_counts().reset_index(); od.columns=["Objetivo","N"]
            fig=px.pie(od,names="Objetivo",values="N",hole=.4,title="Objetivos",color_discrete_sequence=[VERDE,AZUL,NARANJA,ROJO,"#A855F7","#06B6D4","#F59E0B"])
            fig.update_layout(**PL); st.plotly_chart(fig,use_container_width=True)
        with c2:
            # Rango etario por sexo
            _df_e=df_act.copy()
            _df_e["Rango"]=pd.cut(pd.to_numeric(_df_e["edad"],errors="coerce"),
                bins=[0,17,24,34,44,54,120],labels=["<18","18-24","25-34","35-44","45-54","55+"])
            _df_e["Sexo"]=_df_e["sexo"].str.capitalize()
            _et=_df_e.groupby(["Rango","Sexo"],observed=True).size().reset_index(name="N")
            fig2=px.bar(_et,x="Rango",y="N",color="Sexo",barmode="group",
                title="Activos por rango etario y sexo",
                color_discrete_map={"Masculino":AZUL,"Femenino":"#E91E8C","Otro":NARANJA},
                text="N")
            fig2.update_traces(textposition="outside",textfont_color=BLANCO,textfont_size=11)
            fig2.update_layout(**PL,height=300); st.plotly_chart(fig2,use_container_width=True)
        da2=df_act[["N°","nombre","rut","tipo_plan","frecuencia","horario","fecha_vencimiento","celular","objetivo","nivel"]].copy()
        da2["N°"]=range(1,len(da2)+1)
        da2["fecha_vencimiento"]=da2["fecha_vencimiento"].apply(fmt_fecha)
        st.dataframe(da2,use_container_width=True)
        _exa,_exb=st.columns(2)
        _csv_a=da2.to_csv(index=False).encode("utf-8-sig")
        _exa.download_button("⬇️ CSV",_csv_a,"activos.csv","text/csv",use_container_width=True,key="rep_act_csv")
        _filas_a="".join([f"<tr><td>{r['N°']}</td><td>{r['nombre']}</td><td>{r['rut']}</td><td>{r['tipo_plan']}</td><td>{r['fecha_vencimiento']}</td><td>{r['celular']}</td></tr>" for _,r in da2.iterrows()])
        _html_a=f"""<!DOCTYPE html><html><head><meta charset='utf-8'><style>@page{{size:letter;margin:2cm}}body{{font-family:Arial;font-size:9pt}}h2{{color:#6DBE45}}table{{width:100%;border-collapse:collapse}}th{{background:#6DBE45;color:black;padding:5px}}td{{padding:4px;border-bottom:1px solid #ddd}}tr:nth-child(even){{background:#f9f9f9}}</style></head><body><h2>PUTÚ ACTIVO — Clientes Activos</h2><p>Fecha: {hoy.strftime('%d/%m/%Y')} | Total: {len(da2)}</p><table><tr><th>#</th><th>Nombre</th><th>RUT</th><th>Plan</th><th>Vencimiento</th><th>Celular</th></tr>{_filas_a}</table></body></html>"""
        if REPORTLAB_OK:
            _pb_ac=io.BytesIO()
            _pd_ac=SimpleDocTemplate(_pb_ac,pagesize=A4,leftMargin=1.5*cm,rightMargin=1.5*cm,topMargin=1.5*cm,bottomMargin=1.5*cm)
            _aw_ac=A4[0]-3*cm; _st_ac=[]
            _st_ac.append(Paragraph(f"PUTÚ ACTIVO — Clientes Activos",ParagraphStyle("aH",fontName="Helvetica-Bold",fontSize=13,textColor=rl_colors.HexColor("#6DBE45"),spaceAfter=4)))
            _st_ac.append(Paragraph(f"Fecha: {hoy.strftime('%d/%m/%Y')} | Total: {len(da2)}",ParagraphStyle("aS",fontName="Helvetica",fontSize=9,textColor=rl_colors.HexColor("#888"),spaceAfter=8)))
            _ac_hdr=[["#","Nombre","RUT","Plan","Vencimiento","Celular"]]
            _ac_rows=[[str(r3["N°"]),str(r3["nombre"]),str(r3["rut"]),str(r3["tipo_plan"]),str(r3["fecha_vencimiento"]),str(r3["celular"])] for _,r3 in da2.iterrows()]
            _tac=Table(_ac_hdr+_ac_rows,colWidths=[_aw_ac*0.06,_aw_ac*0.28,_aw_ac*0.16,_aw_ac*0.14,_aw_ac*0.18,_aw_ac*0.18])
            _tac.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),rl_colors.HexColor("#6DBE45")),("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),("FONTSIZE",(0,0),(-1,-1),8),("ROWBACKGROUNDS",(0,1),(-1,-1),[rl_colors.white,rl_colors.HexColor("#F5F5F5")]),("GRID",(0,0),(-1,-1),0.3,rl_colors.HexColor("#DDD")),("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4),("LEFTPADDING",(0,0),(-1,-1),5)]))
            _st_ac.append(_tac)
            _pd_ac.build(_st_ac)
            _exb.download_button("⬇️ PDF Clientes Activos",_pb_ac.getvalue(),"activos.pdf","application/pdf",use_container_width=True,key="rep_act_pdf")
    with ti:
        st.markdown(f'<div class="info-box">📋 {len(df_inac)} inactivos — solo consulta.</div>',unsafe_allow_html=True)
        bi=st.text_input("Buscar",key="ib2")
        di=df_inac.copy()
        if bi.strip(): di=di[di["nombre"].str.contains(bi,case=False,na=False)|di["rut"].str.contains(bi,case=False,na=False)]
        di2=di[["nombre","rut","tipo_plan","fecha_vencimiento","celular","estado"]].copy(); di2.insert(0,"N°",range(1,len(di2)+1)); di2["fecha_vencimiento"]=di2["fecha_vencimiento"].apply(fmt_fecha)
        st.dataframe(di2.reset_index(drop=True),use_container_width=True,height=400)
    with tv:
        # Filtros
        vf1,vf2,vf3 = st.columns(3)
        mes_venc = vf1.selectbox("Filtrar por mes",
            ["Todos los meses"] + [MESES_ESP[i].capitalize() for i in range(12)],
            index=0, key="venc_mes")
        anio_venc = vf2.number_input("Año",2024,2030,hoy.year,key="venc_anio")
        rango_dias = vf3.selectbox("Urgencia",
            ["Todos","🚨 Hasta 5 días","⚠️ Hasta 15 días","📅 Hasta 30 días","Todos los activos"],
            index=0, key="venc_rango")

        dv2 = df_act[df_act["fecha_vencimiento"].notna()].copy()
        dv2["días"] = dv2["fecha_vencimiento"].apply(dias_para_vencer)
        dv2 = dv2[dv2["días"].notna()].copy()

        # Filtro por mes/año de vencimiento
        if mes_venc != "Todos los meses":
            mes_idx = [m.capitalize() for m in MESES_ESP].index(mes_venc)+1
            def _mes_venc(fv):
                try: return date.fromisoformat(str(fv)[:10]).month==mes_idx and date.fromisoformat(str(fv)[:10]).year==anio_venc
                except: return False
            dv2 = dv2[dv2["fecha_vencimiento"].apply(_mes_venc)]
        else:
            dv2 = dv2[dv2["días"] >= 0]  # solo futuros por defecto

        # Filtro por urgencia
        if "5 días" in rango_dias:   dv2=dv2[dv2["días"]<=5]
        elif "15 días" in rango_dias: dv2=dv2[dv2["días"]<=15]
        elif "30 días" in rango_dias: dv2=dv2[dv2["días"]<=30]

        # Ordenar de menor a mayor días
        dv2 = dv2.sort_values("días")

        # Métricas rápidas
        mv1,mv2_,mv3_ = st.columns(3)
        mv1.metric("Total mostrados", len(dv2))
        mv2_.metric("🚨 Urgentes (≤5d)", int((dv2["días"]<=5).sum()))
        mv3_.metric("⚠️ Próximos (≤15d)", int((dv2["días"]<=15).sum()))

        if dv2.empty:
            st.markdown('<div class="info-box">Sin vencimientos para el filtro seleccionado.</div>',unsafe_allow_html=True)
        else:
            for _,r in dv2.iterrows():
                d=int(r["días"])
                c=ROJO if d<=5 else NARANJA if d<=15 else VERDE if d<=30 else GRIS_T
                badge="🚨 URGENTE" if d<=5 else "⚠️ Pronto" if d<=15 else "📅 Este mes" if d<=30 else "📆"
                mv2_txt=str(r.get("mensaje_vencimiento","")) or msg_vencimiento(r["nombre"],r["fecha_vencimiento"])
                uv2=wa_url(r["celular"],mv2_txt)
                st.markdown(f'''<div style="background:{GRIS2};border:1px solid {GRIS3};border-left:5px solid {c};
                    border-radius:11px;padding:10px 16px;margin:4px 0;
                    display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px;">
                  <div>
                    <span style="color:{c};font-weight:700;font-size:.9rem">{badge} </span>
                    <b>{r["nombre"]}</b>
                    <span style="color:{GRIS_T};font-size:.85rem"> · {r["tipo_plan"]} · vence {r["fecha_vencimiento"]}</span>
                    <div style="color:{GRIS_T};font-size:.8rem">📱 {fmt_cel(r["celular"])}</div>
                  </div>
                  <div style="display:flex;align-items:center;gap:10px;">
                    <span style="color:{c};font-weight:900;font-size:1.3rem">{d}d</span>
                    <a href="{uv2}" target="_blank"
                       style="background:#25D366;color:white;padding:6px 14px;border-radius:8px;
                              text-decoration:none;font-size:.88rem;font-weight:700;">📲 WA</a>
                  </div>
                </div>''',unsafe_allow_html=True)
    with tc2:
        st.markdown(f'<div style="font-size:1.1rem;font-weight:700;color:{VERDE};margin-bottom:10px;">🎂 Cumpleaños por mes</div>',unsafe_allow_html=True)
        ms_sel=st.selectbox("Seleccionar mes",MESES_ESP,index=hoy.month-1,key="cumple_mes_top")
        mes_num_sel=MESES_ESP.index(ms_sel.lower())+1 if ms_sel.lower() in MESES_ESP else hoy.month

        # Filtrar desde fecha_nacimiento directamente
        def _get_mes_fn(fn):
            try: return date.fromisoformat(str(fn)[:10]).month
            except: return 0
        def _get_dia_fn(fn):
            try: return date.fromisoformat(str(fn)[:10]).day
            except: return 99

        if "mes_nac" in df_act.columns:
            df_cm=df_act[df_act["mes_nac"]==mes_num_sel].copy()
        else:
            df_cm=df_act[df_act["fecha_nacimiento"].apply(_get_mes_fn)==mes_num_sel].copy()

        if df_cm.empty:
            st.markdown(f'<div class="info-box">Sin cumpleaños activos en {ms_sel.capitalize()}.</div>',unsafe_allow_html=True)
        else:
            df_cm["_dia"]=df_cm["fecha_nacimiento"].apply(_get_dia_fn)
            df_cm=df_cm.sort_values("_dia")
            st.markdown(f"<b style='color:{VERDE}'>{ms_sel.capitalize()} — {len(df_cm)} cliente(s)</b>",unsafe_allow_html=True)
            st.markdown("")
            for _,r in df_cm.iterrows():
                dia_num=int(r["_dia"]) if r["_dia"]!=99 else "?"
                es_hoy=(r["_dia"]==hoy.day and mes_num_sel==hoy.month)
                borde=f"3px solid {VERDE}" if es_hoy else f"1px solid {GRIS3}"
                badge='<span style="background:#6DBE45;color:#0D0D0D;border-radius:20px;padding:2px 8px;font-size:.78rem;font-weight:700;margin-left:6px">🎂 HOY</span>' if es_hoy else ""
                mc2=str(r.get("mensaje_cumpleanos","")) or msg_cumpleanos(r["nombre"])
                uc2=wa_url(r["celular"],mc2)
                try:
                    fn_d=date.fromisoformat(fmt_fecha(r["fecha_nacimiento"]))
                    edad_c=int((hoy-fn_d).days/365.25); edad_str=f"{edad_c} años"
                except: edad_str=""
                st.markdown(f'''<div style="background:#1E1E1E;border:{borde};border-radius:10px;
                    padding:10px 14px;margin:4px 0;
                    display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:6px;">
                  <div>
                    <span style="color:#6DBE45;font-weight:900;font-size:1.05rem;min-width:50px;display:inline-block;">Día {dia_num}</span>
                    <b style="font-size:.98rem">{r["nombre"]}</b>{badge}
                    <span style="color:#CCCCCC;font-size:.82rem;margin-left:8px">{edad_str}</span>
                    <div style="color:#CCCCCC;font-size:.82rem;margin-top:2px">
                      📱 {fmt_cel(r["celular"])} · 💳 {r.get("tipo_plan","")} · {fmt_fecha(r.get("fecha_nacimiento",""))}
                    </div>
                  </div>
                  <a href="{uc2}" target="_blank"
                     style="background:#25D366;color:white;padding:7px 16px;border-radius:8px;
                            text-decoration:none;font-size:.88rem;font-weight:700;white-space:nowrap;">
                    🎂 WhatsApp
                  </a>
                </div>''',unsafe_allow_html=True)
    with tfc:
        st.markdown("**Flujo de caja mensual resumido**")
        n_m=st.slider("Últimos meses",2,12,6)
        ay2=st.number_input("Año base",2024,2030,hoy.year)
        filas=[]
        for i in range(n_m-1,-1,-1):
            mi=(hoy.month-1-i)%12+1; ai=ay2 if (hoy.month-i)>0 else ay2-1; fm_i=f"{ai}-{mi:02d}"
            pm=int(db_query("SELECT COALESCE(SUM(monto),0) as s FROM pagos WHERE fecha LIKE ?",(f"{fm_i}%",)).iloc[0]["s"])
            cm2=int(db_query("SELECT COALESCE(SUM(monto),0) as s FROM clases WHERE fecha LIKE ?",(f"{fm_i}%",)).iloc[0]["s"])
            prm=int(db_query("SELECT COALESCE(SUM(monto),0) as s FROM productos WHERE fecha LIKE ?",(f"{fm_i}%",)).iloc[0]["s"])
            em2=int(db_query("SELECT COALESCE(SUM(monto),0) as s FROM egresos WHERE fecha LIKE ?",(f"{fm_i}%",)).iloc[0]["s"])
            ing=pm+cm2+prm; filas.append({"Mes":fm_i,"Ingresos":ing,"Egresos":em2,"Saldo":ing-em2})
        dfc=pd.DataFrame(filas)
        fig_fc=go.Figure()
        fig_fc.add_trace(go.Bar(name="Ingresos",x=dfc["Mes"],y=dfc["Ingresos"],marker_color=VERDE))
        fig_fc.add_trace(go.Bar(name="Egresos",x=dfc["Mes"],y=dfc["Egresos"],marker_color=ROJO))
        fig_fc.add_trace(go.Scatter(name="Saldo",x=dfc["Mes"],y=dfc["Saldo"],mode="lines+markers",line=dict(color=NARANJA,width=2.5),marker_size=8))
        fig_fc.update_layout(title="Flujo de caja mensual",**PL,barmode="group"); st.plotly_chart(fig_fc,use_container_width=True)
        dfc2=dfc.copy()
        for col in ["Ingresos","Egresos","Saldo"]: dfc2[col]=dfc2[col].apply(lambda x:f"${x:,}")
        st.dataframe(dfc2,use_container_width=True,hide_index=True)

    with tast:
        st.markdown(f"<b style='color:{VERDE}'>Filtrar asistencias</b>",unsafe_allow_html=True)
        _af1,_af2,_af3=st.columns(3)
        _a_modo=_af1.radio("Ver por",["Día","Mes","Cliente"],horizontal=True,key="ast_modo")
        if _a_modo=="Día":
            fd_r=_af2.date_input("Día",value=hoy,key="rep_dia",format="DD/MM/YYYY")
            dad_r=db_query("SELECT * FROM asistencia WHERE fecha=? ORDER BY hora",(str(fd_r),))
            if not dad_r.empty:
                _en_r=dad_r[(dad_r["tipo"]=="ingreso")&(dad_r["hora_salida"].isna()|dad_r["hora_salida"].eq(""))] if "tipo" in dad_r.columns else dad_r
                c1r,c2r=st.columns(2); c1r.metric("Total registros",len(dad_r)); c2r.metric("En sala",len(_en_r))
                cols_r=[c for c in ["hora","tipo","nombre","hora_salida","rut"] if c in dad_r.columns]
                _dad_d=dad_r[cols_r].copy(); _dad_d.insert(0,"N°",range(1,len(_dad_d)+1))
                st.dataframe(_dad_d.rename(columns={"hora":"Ingreso","tipo":"Tipo","hora_salida":"Salida","nombre":"Nombre","rut":"RUT"}),use_container_width=True)
            else: st.markdown(f'<div class="info-box">Sin asistencias el {fmt_fecha(str(fd_r))}.</div>',unsafe_allow_html=True)
        elif _a_modo=="Mes":
            cm_r,cy_r=_af2.columns(2) if False else (_af2,_af3)
            ms_r=cm_r.selectbox("Mes",range(1,13),index=hoy.month-1,format_func=lambda x:MESES_ESP[x-1].capitalize(),key="rm")
            ay_r=cy_r.number_input("Año",2024,2030,hoy.year,key="ry")
            fm_r=f"{ay_r}-{ms_r:02d}"
            dam_r=db_query("SELECT * FROM asistencia WHERE fecha LIKE ? ORDER BY fecha,hora",(f"{fm_r}%",))
            if not dam_r.empty:
                st.metric(f"Asistencias en {MESES_ESP[ms_r-1].capitalize()} {ay_r}",len(dam_r))
                resumen=dam_r.groupby("fecha").size().reset_index(name="N"); resumen.columns=["Fecha","Asistencias"]
                fig_am=px.bar(resumen,x="Fecha",y="Asistencias",color_discrete_sequence=[VERDE],title=f"Asistencia diaria {fm_r}")
                fig_am.update_layout(**PL,height=220); st.plotly_chart(fig_am,use_container_width=True)
                dam_r_disp=dam_r.copy(); dam_r_disp["fecha"]=dam_r_disp["fecha"].apply(fmt_fecha)
                dam_r_disp.insert(0,"N°",range(1,len(dam_r_disp)+1))
                st.dataframe(dam_r_disp[["N°","fecha","hora","nombre","rut"]].rename(columns={"fecha":"Fecha","hora":"Hora","nombre":"Nombre","rut":"RUT"}),use_container_width=True,height=260)
            else: st.markdown(f'<div class="info-box">Sin asistencias en {fm_r}.</div>',unsafe_allow_html=True)
        else:
            bac_r=_af2.text_input("Buscar cliente",key="rc")
            if bac_r.strip():
                da_r=db_query("SELECT * FROM asistencia WHERE nombre LIKE ? OR rut LIKE ? ORDER BY fecha DESC",(f"%{bac_r}%",f"%{bac_r}%"))
                if not da_r.empty:
                    st.metric("Total asistencias",len(da_r))
                    da_r_d=da_r.copy(); da_r_d["fecha"]=da_r_d["fecha"].apply(fmt_fecha)
                    da_r_d.insert(0,"N°",range(1,len(da_r_d)+1))
                    st.dataframe(da_r_d[["N°","fecha","hora","nombre","rut"]].rename(columns={"fecha":"Fecha","hora":"Hora","nombre":"Nombre","rut":"RUT"}),use_container_width=True,height=300)
                else: st.markdown('<div class="info-box">Sin asistencias.</div>',unsafe_allow_html=True)

    with tven:
        st.markdown(f"<b style='color:{VERDE}'>Ventas de Productos</b>",unsafe_allow_html=True)
        _tv1,_tv2=st.columns(2)
        _tv_mes=_tv1.selectbox("Mes",range(1,13),index=hoy.month-1,format_func=lambda x:MESES_ESP[x-1].capitalize(),key="tv_mes")
        _tv_anio=_tv2.number_input("Año",2024,2030,hoy.year,key="tv_anio")
        _fm_tv=f"{_tv_anio}-{_tv_mes:02d}"
        _dvp_r=db_query("SELECT * FROM productos WHERE fecha LIKE ? ORDER BY fecha DESC",(_fm_tv+"%",))
        if not _dvp_r.empty:
            st.metric(f"Total ventas {MESES_ESP[_tv_mes-1].capitalize()} {_tv_anio}",f"${int(_dvp_r['monto'].sum()):,}")
            _dvp_r["fecha"]=_dvp_r["fecha"].apply(fmt_fecha)
            _dvp_r["Monto"]=_dvp_r["monto"].apply(lambda x:f"${int(x):,}")
            _dvp_r.insert(0,"N°",range(1,len(_dvp_r)+1))
            st.dataframe(_dvp_r[["N°","fecha","producto","Monto","usuario"]].rename(columns={"fecha":"Fecha","producto":"Producto","usuario":"Vendedor"}),use_container_width=True)
            # Resumen por producto
            _res_prod=_dvp_r.groupby("producto")["monto"].agg(["sum","count"]).reset_index()
            _res_prod.columns=["Producto","Total $","Cantidad"]
            _res_prod=_res_prod.sort_values("Total $",ascending=False)
            _res_prod["Total $"]=_res_prod["Total $"].apply(lambda x:f"${int(x):,}")
            st.markdown(f"<b style='color:{VERDE}'>Resumen por producto</b>",unsafe_allow_html=True)
            st.dataframe(_res_prod.reset_index(drop=True),use_container_width=True)
        else: st.markdown(f'<div class="info-box">Sin ventas en {MESES_ESP[_tv_mes-1].capitalize()} {_tv_anio}.</div>',unsafe_allow_html=True)

    with tie:
        _teg,_tres=st.tabs(["➕ Registrar egreso","📊 Resumen mensual"])
        with _teg:
            with st.form("feg_r"):
                _e1,_e2,_e3=st.columns(3)
                _fe=_e1.date_input("Fecha",value=hoy,key="r_eg_f",format="DD/MM/YYYY")
                _ca=_e2.selectbox("Categoría",["Arriendo","Servicios básicos","Equipamiento","Sueldos","Limpieza","Marketing","Otros"],key="r_eg_c")
                _mo=_e3.number_input("Monto $",min_value=0,max_value=100000000,value=None,step=1000,key="r_eg_m",placeholder="Ingrese monto")
                _de=_e1.text_input("Descripción",key="r_eg_d")
                _ok=st.form_submit_button("💾 Registrar egreso")
            if _ok:
                db_exec("INSERT INTO egresos (fecha,categoria,monto,descripcion,usuario) VALUES (?,?,?,?,?)",
                    (str(_fe),_ca,_mo,_de,st.session_state.nombre_u))
                st.markdown('<div class="success-box">✅ Egreso registrado.</div>',unsafe_allow_html=True)
            st.markdown(f'<div class="info-box">Para registrar ventas de productos usa el menú <b>🛍 Venta Productos</b>.</div>',unsafe_allow_html=True)
        with _tres:
            _cm2,_cy2=st.columns(2)
            _ms2=_cm2.selectbox("Mes",range(1,13),index=hoy.month-1,format_func=lambda x:MESES_ESP[x-1].capitalize(),key="r_res_m")
            _ay2=_cy2.number_input("Año",2024,2030,hoy.year,key="r_res_y")
            _fm2=f"{_ay2}-{_ms2:02d}"
            _dp2=db_query("SELECT monto FROM pagos WHERE fecha LIKE ? AND concepto!='Clase funcional'",(_fm2+"%",))
            _dc2=db_query("SELECT monto FROM clases WHERE fecha LIKE ?",(_fm2+"%",))
            _dpr2=db_query("SELECT monto FROM productos WHERE fecha LIKE ?",(_fm2+"%",))
            _deg2=db_query("SELECT * FROM egresos WHERE fecha LIKE ? ORDER BY fecha",(_fm2+"%",))
            _pm2=int(_dp2["monto"].sum()) if not _dp2.empty else 0
            _cm3=int(_dc2["monto"].sum()) if not _dc2.empty else 0
            _pr2=int(_dpr2["monto"].sum()) if not _dpr2.empty else 0
            _eg2=int(_deg2["monto"].sum()) if not _deg2.empty else 0
            _it2=_pm2+_cm3+_pr2; _sl2=_it2-_eg2
            _r1,_r2,_r3,_r4=st.columns(4)
            _r1.metric("Mensualidades",f"${_pm2:,}"); _r2.metric("Clases",f"${_cm3:,}")
            _r3.metric("Productos",f"${_pr2:,}"); _r4.metric("Total ingresos",f"${_it2:,}")
            _r5,_r6,_,_=st.columns(4)
            _r5.metric("Egresos",f"${_eg2:,}"); _r6.metric("Saldo",f"${_sl2:,}")
            # Ingresos del mes desde pagos
            _dp3=db_query("SELECT * FROM pagos WHERE fecha LIKE ? ORDER BY fecha",(_fm2+"%",))
            if not _dp3.empty:
                st.markdown(f"**Pagos del mes {_fm2}**")
                _dp3["Monto"]=_dp3["monto"].apply(lambda x:f"${int(x):,}")
                st.dataframe(_dp3[["fecha","nombre","Monto","concepto","tipo_plan","medio_pago"]].rename(columns={"fecha":"Fecha","nombre":"Cliente","concepto":"Período","tipo_plan":"Plan","medio_pago":"Medio"}),use_container_width=True,height=220)

# ════════════════════════════════════════════════════════════════════════════
# ⚙️ BASE DE DATOS
# ════════════════════════════════════════════════════════════════════════════

    with tseg:
        st.markdown(f"<b style='color:{VERDE}'>Resumen de seguimiento por cliente activo</b>",unsafe_allow_html=True)
        # Consultas
        _rut_ids=set(db_query("SELECT DISTINCT cliente_rut FROM rutinas WHERE activa=1 AND cliente_rut IS NOT NULL")["cliente_rut"].tolist()) if not db_query("SELECT DISTINCT cliente_rut FROM rutinas WHERE activa=1 AND cliente_rut IS NOT NULL").empty else set()
        _eval_ids=set(db_query("SELECT DISTINCT rut FROM evaluaciones")["rut"].tolist()) if not db_query("SELECT DISTINCT rut FROM evaluaciones").empty else set()
        _nutr_ids=set(db_query("SELECT DISTINCT cliente_rut FROM planes_nutri WHERE activo=1")["cliente_rut"].tolist()) if not db_query("SELECT DISTINCT cliente_rut FROM planes_nutri WHERE activo=1").empty else set()
        # Métricas resumen
        _sm1,_sm2,_sm3,_sm4=st.columns(4)
        _sm1.metric("👥 Total activos",len(df_act))
        _sm2.metric("💪 Con rutina",len([r for r in df_act["rut"] if r in _rut_ids]))
        _sm3.metric("📏 Con evaluación",len([r for r in df_act["rut"] if r in _eval_ids]))
        _sm4.metric("🥗 Con nutrición",len([r for r in df_act["rut"] if r in _nutr_ids]))
        st.divider()
        # Filtros
        _sf1,_sf2,_sf3=st.columns(3)
        _sf_rut=_sf1.selectbox("Rutina",["Todos","Con rutina","Sin rutina"],key="seg_rut")
        _sf_eval=_sf2.selectbox("Evaluación",["Todos","Con evaluación","Sin evaluación"],key="seg_eval")
        _sf_nutr=_sf3.selectbox("Nutrición",["Todos","Con nutrición","Sin nutrición"],key="seg_nutr")
        # Tabla
        _seg_rows=[]
        for _,_sc in df_act.iterrows():
            _sr=_sc["rut"]
            _tiene_rut=_sr in _rut_ids
            _tiene_eval=_sr in _eval_ids
            _tiene_nutr=_sr in _nutr_ids
            if _sf_rut=="Con rutina" and not _tiene_rut: continue
            if _sf_rut=="Sin rutina" and _tiene_rut: continue
            if _sf_eval=="Con evaluación" and not _tiene_eval: continue
            if _sf_eval=="Sin evaluación" and _tiene_eval: continue
            if _sf_nutr=="Con nutrición" and not _tiene_nutr: continue
            if _sf_nutr=="Sin nutrición" and _tiene_nutr: continue
            _seg_rows.append({
                "Nombre":_sc["nombre"],
                "RUT":_sr,
                "Plan":_sc.get("tipo_plan",""),
                "💪 Rutina":"✅" if _tiene_rut else "❌",
                "📏 Evaluación":"✅" if _tiene_eval else "❌",
                "🥗 Nutrición":"✅" if _tiene_nutr else "❌",
            })
        if _seg_rows:
            import pandas as _pd_seg
            _seg_df=_pd_seg.DataFrame(_seg_rows)
            st.dataframe(_seg_df,use_container_width=True,hide_index=True,height=min(40*len(_seg_rows)+40,600))
            st.caption(f"{len(_seg_rows)} clientes mostrados")
        else:
            st.markdown('<div class="info-box">Sin resultados con los filtros aplicados.</div>',unsafe_allow_html=True)

elif pagina=="💪 Ejercicios":
    NIVELES_R=["Principiante","Intermedio","Avanzado"]
    COLOR_NIVEL_R={"Principiante":"#22C55E","Intermedio":"#EAB308","Avanzado":"#3B82F6"}

    def _s(v,d=""):
        import math
        if v is None: return d
        if isinstance(v,float) and math.isnan(v): return d
        return str(v)
    def _has_v(v):
        import math
        if v is None: return False
        if isinstance(v,float) and math.isnan(v): return False
        return str(v).strip()!=""
    def _badge_niv(n):
        c=COLOR_NIVEL_R.get(n,GRIS_T)
        return f'<span style="background:{c}22;color:{c};padding:1px 8px;border-radius:10px;font-size:.72rem;font-weight:600;border:1px solid {c}55">● {n}</span>'
    def _ej_get_musculos():
        return [r[0] for r in get_conn().execute("SELECT DISTINCT musculo_primario FROM ejercicios WHERE musculo_primario IS NOT NULL ORDER BY 1").fetchall()]

    # ── Estado: si hay ejercicio seleccionado → pantalla completa de detalle ──
    _sel_id = st.session_state.get("ej_sel_id")

    if _sel_id:
        # ── VISTA DETALLE — pantalla completa ─────────────────────────────
        _sel_df = db_query("SELECT * FROM ejercicios WHERE id=?",(_sel_id,))
        if _sel_df.empty:
            st.session_state.pop("ej_sel_id",None); st.rerun()
        _ejdet = _sel_df.iloc[0].to_dict()

        # Botón volver
        if st.button("← Volver a la biblioteca",key="ej_volver",type="primary"):
            st.session_state.pop("ej_sel_id",None); st.rerun()

        st.markdown(f"<h2 style='color:{VERDE};margin:0 0 4px 0'>{_ejdet['nombre']}</h2>",unsafe_allow_html=True)

        _tipos2=[]
        if _ejdet.get("gimnasio"): _tipos2.append("🏋️ Gym")
        if _ejdet.get("casa"): _tipos2.append("🏠 Casa")
        if _ejdet.get("estiramiento"): _tipos2.append("🤸 Estiramiento")
        if _ejdet.get("rehabilitacion"): _tipos2.append("🩺 Rehabilitación")

        _badges=""
        if _has_v(_ejdet.get("nivel")): _badges+=_badge_niv(_ejdet["nivel"])+" "
        for _t2 in _tipos2: _badges+=f'<span style="background:{GRIS2};color:{BLANCO};padding:1px 8px;border-radius:10px;font-size:.72rem;border:1px solid {GRIS3}">{_t2}</span> '
        if _badges: st.markdown(_badges,unsafe_allow_html=True)
        st.markdown("")

        _dc1,_dc2=st.columns([1,1.6])
        with _dc1:
            if _has_v(_ejdet.get("url_imagen")):
                try: st.image(_ejdet["url_imagen"],use_container_width=True)
                except: st.markdown(f'<div style="font-size:4rem;text-align:center">🏋️</div>',unsafe_allow_html=True)
            else:
                st.markdown(f'<div style="font-size:4rem;text-align:center;padding:40px 0">🏋️</div>',unsafe_allow_html=True)
            # Músculos
            st.markdown(f"""<div style="background:{GRIS2};border-radius:10px;padding:14px 16px;margin-top:12px">
                <b style="color:{VERDE}">💪 Músculos</b><br>
                <span style="font-size:.9rem"><b>Principal:</b> {_s(_ejdet.get("musculo_primario"),"—")}</span><br>
                <span style="font-size:.9rem"><b>Secundario:</b> {_s(_ejdet.get("musculo_secundario"),"—")}</span>
            </div>""",unsafe_allow_html=True)

        with _dc2:
            # Descripción
            if _has_v(_ejdet.get("descripcion")):
                st.markdown(f"""<div style="background:{GRIS2};border-radius:10px;padding:14px 16px;margin-bottom:12px">
                    <b style="color:{VERDE}">📝 Descripción</b><br>
                    <span style="font-size:.88rem;color:{BLANCO}">{_s(_ejdet.get("descripcion"))}</span>
                </div>""",unsafe_allow_html=True)
            # Ejecución
            if _has_v(_ejdet.get("ejecucion")):
                st.markdown(f"""<div style="background:{GRIS2};border-radius:10px;padding:14px 16px;margin-bottom:12px">
                    <b style="color:{VERDE}">📋 Ejecución</b><br>
                    <span style="font-size:.88rem;color:{BLANCO};line-height:1.6">{_s(_ejdet.get("ejecucion")).replace(chr(10),"<br>")}</span>
                </div>""",unsafe_allow_html=True)
            # Video
            if _has_v(_ejdet.get("video")):
                import re as _re2
                _vm=_re2.search(r"(?:embed/|youtu\.be/|watch\?v=)([A-Za-z0-9_-]{6,})",str(_ejdet["video"]))
                if _vm:
                    st.markdown(f"<b style='color:{VERDE}'>🎥 Video</b>",unsafe_allow_html=True)
                    st.video(f"https://www.youtube.com/watch?v={_vm.group(1)}")

        st.divider()
        # Editar ejercicio
        with st.expander("✏️ Editar este ejercicio",expanded=False):
            with st.form(f"ej_ed_{_sel_id}"):
                _en=st.text_input("Nombre",value=_s(_ejdet.get("nombre")),key=f"ej_nom_{_sel_id}")
                _ec1,_ec2=st.columns(2)
                _emp=_ec1.text_input("Músculo primario",value=_s(_ejdet.get("musculo_primario")),key=f"ej_mp_{_sel_id}")
                _ems=_ec2.text_input("Músculo secundario",value=_s(_ejdet.get("musculo_secundario")),key=f"ej_ms_{_sel_id}")
                _env=st.selectbox("Nivel",["Sin definir"]+NIVELES_R,
                    index=(NIVELES_R.index(_ejdet["nivel"])+1) if _ejdet.get("nivel") in NIVELES_R else 0,
                    key=f"ej_niv_ed_{_sel_id}")
                _edesc=st.text_area("Descripción",value=_s(_ejdet.get("descripcion")),height=70,key=f"ej_desc_{_sel_id}")
                _eejec=st.text_area("Ejecución",value=_s(_ejdet.get("ejecucion")),height=90,key=f"ej_ejec_{_sel_id}")
                _eimg=st.text_input("URL imagen",value=_s(_ejdet.get("url_imagen")),key=f"ej_img_{_sel_id}")
                _evid=st.text_input("URL video YouTube",value=_s(_ejdet.get("video")),key=f"ej_vid_{_sel_id}")
                _eb1,_eb2,_eb3,_eb4=st.columns(4)
                _ecasa=_eb1.checkbox("Casa",value=bool(_ejdet.get("casa")),key=f"ej_casa_{_sel_id}")
                _egym=_eb2.checkbox("Gym",value=bool(_ejdet.get("gimnasio")),key=f"ej_gym_{_sel_id}")
                _eest=_eb3.checkbox("Estiram.",value=bool(_ejdet.get("estiramiento")),key=f"ej_est_{_sel_id}")
                _erehab=_eb4.checkbox("Rehab.",value=bool(_ejdet.get("rehabilitacion")),key=f"ej_reh_{_sel_id}")
                _esave,_edel=st.columns([3,1])
                _ok_ej_ed=_esave.form_submit_button("💾 Guardar cambios",use_container_width=True,type="primary")
                _ok_ej_del=_edel.form_submit_button("🗑️ Eliminar",use_container_width=True)
            if _ok_ej_ed and _en.strip():
                _cnej=get_conn()
                _cnej.execute("""UPDATE ejercicios SET nombre=?,url_imagen=?,descripcion=?,ejecucion=?,
                    musculo_primario=?,musculo_secundario=?,casa=?,gimnasio=?,estiramiento=?,rehabilitacion=?,
                    video=?,nivel=? WHERE id=?""",
                    (_en.strip(),_eimg or None,_edesc or None,_eejec or None,_emp or None,_ems or None,
                     int(_ecasa),int(_egym),int(_eest),int(_erehab),_evid or None,
                     _env if _env!="Sin definir" else None,_sel_id))
                _cnej.commit(); _cnej.close()
                st.session_state.pop("ej_sel_id",None); db_query.clear(); st.rerun()
            if _ok_ej_del:
                _cdej=get_conn(); _cdej.execute("DELETE FROM ejercicios WHERE id=?",(_sel_id,)); _cdej.commit(); _cdej.close()
                st.session_state.pop("ej_sel_id",None); db_query.clear(); st.rerun()

    else:
        # ── VISTA BIBLIOTECA — grid + buscador ────────────────────────────
        st.markdown('<div class="section-header">💪 Biblioteca de Ejercicios</div>',unsafe_allow_html=True)
        tab_bej,tab_aej,tab_iej=st.tabs(["🔍 Buscar","➕ Agregar ejercicio","📥 Importar CSV"])

        with tab_bej:
            _bf1,_bf2,_bf3=st.columns([3,1.5,1.5])
            ej_txt=_bf1.text_input("🔍 Buscar por nombre o músculo",placeholder="sentadilla, curl, pecho...",key="ej_txt",label_visibility="collapsed")
            ej_musc=_bf2.selectbox("Músculo",["Todos"]+_ej_get_musculos(),key="ej_musc",label_visibility="collapsed")
            ej_niv=_bf3.selectbox("Nivel",["Todos"]+NIVELES_R,key="ej_niv",label_visibility="collapsed")
            _bb1,_bb2,_bb3,_bb4=st.columns(4)
            ej_gym=_bb1.checkbox("🏋️ Gym",key="ej_gym"); ej_casa=_bb2.checkbox("🏠 Casa",key="ej_casa")
            ej_estir=_bb3.checkbox("🤸 Estiram.",key="ej_estir"); ej_rehab=_bb4.checkbox("🩺 Rehab.",key="ej_rehab")

            _ej_where=[]; _ej_params=[]
            if ej_txt.strip(): _ej_where.append("(nombre LIKE ? OR musculo_primario LIKE ?)"); _ej_params.extend([f"%{ej_txt}%",f"%{ej_txt}%"])
            if ej_musc!="Todos": _ej_where.append("(musculo_primario=? OR musculo_secundario=?)"); _ej_params.extend([ej_musc,ej_musc])
            if ej_niv!="Todos": _ej_where.append("nivel=?"); _ej_params.append(ej_niv)
            if ej_gym: _ej_where.append("gimnasio=1")
            if ej_casa: _ej_where.append("casa=1")
            if ej_estir: _ej_where.append("estiramiento=1")
            if ej_rehab: _ej_where.append("rehabilitacion=1")
            _ej_q="SELECT * FROM ejercicios"+(" WHERE "+" AND ".join(_ej_where) if _ej_where else "")+" ORDER BY nombre COLLATE NOCASE LIMIT 120"
            _ej_res=db_query(_ej_q,tuple(_ej_params))
            st.caption(f"{get_conn().execute('SELECT COUNT(*) FROM ejercicios').fetchone()[0]} ejercicios en la base de datos — {len(_ej_res)} resultado(s) — haz click en un ejercicio para ver el detalle")

            if not _ej_res.empty:
                for _chunk in range(0,len(_ej_res),6):
                    _row_cols=st.columns(6)
                    for _ci2,(_,_ej) in enumerate(_ej_res.iloc[_chunk:_chunk+6].iterrows()):
                        _ejd=_ej.to_dict()
                        with _row_cols[_ci2]:
                            if _has_v(_ejd.get("url_imagen")):
                                try: st.image(_ejd["url_imagen"],use_container_width=True)
                                except: st.markdown(f'<div style="height:120px;background:{GRIS3};border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:2rem">🏋️</div>',unsafe_allow_html=True)
                            else:
                                st.markdown(f'<div style="height:120px;background:{GRIS3};border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:2rem">🏋️</div>',unsafe_allow_html=True)
                            st.markdown(f'<div style="font-size:.75rem;font-weight:700;color:{BLANCO};line-height:1.2;margin:3px 0;text-align:center">{_ejd["nombre"][:28]}</div>',unsafe_allow_html=True)
                            if _has_v(_ejd.get("nivel")): st.markdown(f'<div style="text-align:center">{_badge_niv(_ejd["nivel"])}</div>',unsafe_allow_html=True)
                            if st.button("Ver detalle",key=f"ej_sel_{_ejd['id']}",use_container_width=True):
                                st.session_state["ej_sel_id"]=_ejd["id"]; st.rerun()
            else:
                st.markdown('<div class="info-box">Sin resultados — prueba otro término.</div>',unsafe_allow_html=True)

        with tab_aej:
            with st.form("form_add_ej",clear_on_submit=True):
                _an=st.text_input("Nombre *")
                _ac1,_ac2=st.columns(2)
                _amp=_ac1.text_input("Músculo primario"); _ams=_ac2.text_input("Músculo secundario")
                _aniv=st.selectbox("Nivel",["Sin definir"]+NIVELES_R,key="add_ej_niv")
                _adesc=st.text_area("Descripción",height=70); _aejec=st.text_area("Ejecución",height=70)
                _aimg=st.text_input("URL imagen"); _avid=st.text_input("URL video YouTube")
                _ab1,_ab2,_ab3,_ab4=st.columns(4)
                _acasa=_ab1.checkbox("Casa"); _agym=_ab2.checkbox("Gym")
                _aest=_ab3.checkbox("Estiram."); _arehab=_ab4.checkbox("Rehab.")
                _ok_add_ej=st.form_submit_button("➕ Agregar ejercicio",type="primary")
            if _ok_add_ej:
                if not _an.strip(): st.error("El nombre es obligatorio.")
                else:
                    _cnaj=get_conn()
                    _cnaj.execute("""INSERT INTO ejercicios (nombre,url_imagen,descripcion,ejecucion,musculo_primario,
                        musculo_secundario,casa,gimnasio,estiramiento,rehabilitacion,video,nivel)
                        VALUES (?,?,?,?,?,?,?,?,?,?,?,?)""",
                        (_an.strip(),_aimg or None,_adesc or None,_aejec or None,_amp or None,_ams or None,
                         int(_acasa),int(_agym),int(_aest),int(_arehab),_avid or None,
                         _aniv if _aniv!="Sin definir" else None))
                    _cnaj.commit(); _cnaj.close()
                    db_query.clear(); st.success(f"'{_an}' agregado ✅")

        with tab_iej:
            st.caption("CSV con columnas: nombre, musculo_primario, musculo_secundario, nivel, descripcion, ejecucion, url_imagen, video, casa, gimnasio, estiramiento, rehabilitacion")
            _arch_ej=st.file_uploader("CSV o Excel",type=["csv","xlsx"],key="ej_import_file")
            if _arch_ej:
                try:
                    _dfi=pd.read_csv(_arch_ej) if _arch_ej.name.endswith(".csv") else pd.read_excel(_arch_ej)
                    _dfi.columns=[c.strip().lower() for c in _dfi.columns]
                    if "nombre" not in _dfi.columns: st.error("Necesita columna 'nombre'.")
                    else:
                        _bcols2=["casa","gimnasio","estiramiento","rehabilitacion"]
                        for _bc2 in _bcols2:
                            if _bc2 not in _dfi.columns: _dfi[_bc2]=0
                            else: _dfi[_bc2]=_dfi[_bc2].astype(str).str.lower().isin(["si","sí","true","1","yes"]).astype(int)
                        st.dataframe(_dfi.head(5),use_container_width=True)
                        if st.button(f"Importar {len(_dfi)} ejercicios",type="primary",key="ej_import_btn"):
                            _cnij=get_conn()
                            for _rij in _dfi.to_dict("records"):
                                for _nc in ["url_imagen","descripcion","ejecucion","musculo_primario","musculo_secundario","url_body","video","nivel"]:
                                    if _nc not in _rij: _rij[_nc]=None
                                _cnij.execute("""INSERT OR IGNORE INTO ejercicios (nombre,url_imagen,descripcion,ejecucion,
                                    musculo_primario,musculo_secundario,casa,gimnasio,estiramiento,rehabilitacion,video,nivel)
                                    VALUES (:nombre,:url_imagen,:descripcion,:ejecucion,:musculo_primario,:musculo_secundario,
                                    :casa,:gimnasio,:estiramiento,:rehabilitacion,:video,:nivel)""",_rij)
                            _cnij.commit(); _cnij.close()
                            db_query.clear(); st.success(f"{len(_dfi)} ejercicios importados ✅"); st.rerun()
                except Exception as _eij: st.error(f"Error: {_eij}")

elif pagina=="📋 Rutinas":
    st.markdown('<div class="section-header">📋 Rutinas de Entrenamiento</div>',unsafe_allow_html=True)
    if st.button("← Volver",key="rut_volver"): st.session_state._goto="🏠 Dashboard"; st.rerun()

    TODOS_DIAS_R=["Día 1","Día 2","Día 3","Día 4","Día 5","Día 6"]
    METODOS_R=["Normal","Bi-serie","Superserie","Dropset","Piramidal","Isométrico","Al fallo","Excéntrico","Circuito","Otro"]

    def _rv(v,d=""):
        import math
        if v is None: return d
        if isinstance(v,float) and math.isnan(v): return d
        return str(v)
    def _hv(v):
        import math
        if v is None: return False
        if isinstance(v,float) and math.isnan(v): return False
        return str(v).strip()!=""

    # ── helper: render editor ──────────────────────────────────────────────
    def _render_editor(rut_id, label=""):
        _ejs=db_query("""SELECT re.*,e.nombre,e.url_imagen,e.musculo_primario
            FROM rutina_ejercicios re JOIN ejercicios e ON e.id=re.ejercicio_id
            WHERE re.rutina_id=? ORDER BY re.dia_semana,re.orden""",(rut_id,))

        # Días activos (solo los que tienen ejercicios, en orden)
        _dias_con=[d for d in TODOS_DIAS_R if not _ejs.empty and d in _ejs["dia_semana"].values]
        _n_dias=len(_dias_con)

        # Obtener n_dias guardado en session para este rut_id (para agregar días)
        _n_key=f"n_dias_{label}_{rut_id}"
        if _n_key not in st.session_state:
            st.session_state[_n_key]=max(_n_dias,1)
        _n_sel=st.session_state[_n_key]

        # Selector de cantidad de días
        _nd1,_nd2=st.columns([3,1])
        _nd1.markdown(f"<span style='color:{VERDE};font-weight:700'>Días de entrenamiento en esta rutina</span>",unsafe_allow_html=True)
        _n_new=_nd2.number_input("Días",1,6,_n_sel,1,key=f"n_dias_inp_{label}_{rut_id}",label_visibility="collapsed")
        if _n_new!=_n_sel:
            st.session_state[_n_key]=_n_new; st.rerun()

        # Días disponibles según cantidad seleccionada
        _dias_disp=TODOS_DIAS_R[:_n_sel]

        # Calendario resumen — solo días activos
        _dias_show=[d for d in _dias_disp if not _ejs.empty and d in _ejs["dia_semana"].values]
        if _dias_show:
            _hcols=st.columns(len(_dias_show))
            for _ic,_hc in enumerate(_hcols):
                _hc.markdown(f'<div style="background:{VERDE};color:#000;text-align:center;font-weight:700;font-size:.78rem;padding:4px 2px;border-radius:5px 5px 0 0">{_dias_show[_ic]}</div>',unsafe_allow_html=True)
            _dcols=st.columns(len(_dias_show))
            for _dc,_dfl in zip(_dcols,_dias_show):
                _ddf=_ejs[_ejs["dia_semana"]==_dfl] if not _ejs.empty else pd.DataFrame()
                with _dc:
                    _h=f'<div style="background:{GRIS2};border:1px solid #2E2E2E;border-top:none;border-radius:0 0 5px 5px;padding:3px;min-height:80px;">'
                    if _ddf.empty:
                        _h+=f'<div style="color:{GRIS_T};font-size:.62rem;text-align:center;padding:6px 0">—</div>'
                    for _ii2,(_,_ej2) in enumerate(_ddf.iterrows(),1):
                        _ejd2=_ej2.to_dict()
                        _nn2=str(_ejd2.get("nombre",""))[:14]
                        _sr2=_rv(_ejd2.get("series"),"—"); _rp2=_rv(_ejd2.get("repeticiones"),"—")
                        _h+=f'<div style="margin:1px 0;padding:2px 3px;background:#252525;border-radius:3px;border-left:2px solid {VERDE};"><div style="font-size:.62rem;font-weight:700;color:{BLANCO};line-height:1.2">{_ii2}. {_nn2}</div><div style="font-size:.58rem;color:{VERDE}">{_sr2}×{_rp2}</div></div>'
                    _h+="</div>"
                    st.markdown(_h,unsafe_allow_html=True)
        else:
            st.markdown(f'<div class="info-box">Aún no hay ejercicios. Agrégalos en los tabs de abajo.</div>',unsafe_allow_html=True)

        st.markdown(f"<b style='color:{VERDE};font-size:.93rem'>✏️ Editar por día</b>",unsafe_allow_html=True)

        # Tabs solo para los días disponibles
        _dtabs=st.tabs(_dias_disp)
        for _tab_d,_dfl2 in zip(_dtabs,_dias_disp):
            with _tab_d:
                _ddf2=_ejs[_ejs["dia_semana"]==_dfl2] if not _ejs.empty else pd.DataFrame()
                if not _ddf2.empty:
                    for _idx_d,(_,_re) in enumerate(_ddf2.iterrows()):
                        _red=_re.to_dict(); _rid_e=int(_red["id"])
                        with st.container(border=True):
                            _ci,_cn,_cb1,_cb2,_cb3,_cb4=st.columns([1,3.5,.55,.55,.55,.55])
                            if _hv(_red.get("url_imagen")):
                                try: _ci.image(_red["url_imagen"],width=52)
                                except: _ci.markdown("🏋️")
                            else: _ci.markdown("<div style='font-size:1.3rem'>🏋️</div>",unsafe_allow_html=True)
                            _cn.markdown(f"**{_idx_d+1}. {_red['nombre']}**")
                            _cn.caption(f"{_rv(_red.get('metodo'),'Normal')} · {_rv(_red.get('series'),'—')}×{_rv(_red.get('repeticiones'),'—')} · {_rv(_red.get('peso'),'—')}")
                            if _cb1.button("⬆️",key=f"up_{label}_{_rid_e}",use_container_width=True):
                                _hm=get_conn().execute("SELECT id,orden FROM rutina_ejercicios WHERE rutina_id=? AND dia_semana=? ORDER BY orden,id",(rut_id,_dfl2)).fetchall()
                                _ids=[r[0] for r in _hm]; _pos=_ids.index(_rid_e)
                                if _pos>0:
                                    _cu=get_conn(); _cu.execute("UPDATE rutina_ejercicios SET orden=? WHERE id=?",(_hm[_pos-1][1],_rid_e)); _cu.execute("UPDATE rutina_ejercicios SET orden=? WHERE id=?",(_hm[_pos][1],_ids[_pos-1])); _cu.commit(); _cu.close(); db_query.clear(); st.rerun()
                            if _cb2.button("⬇️",key=f"dn_{label}_{_rid_e}",use_container_width=True):
                                _hm2=get_conn().execute("SELECT id,orden FROM rutina_ejercicios WHERE rutina_id=? AND dia_semana=? ORDER BY orden,id",(rut_id,_dfl2)).fetchall()
                                _ids2=[r[0] for r in _hm2]; _pos2=_ids2.index(_rid_e)
                                if _pos2<len(_ids2)-1:
                                    _cd=get_conn(); _cd.execute("UPDATE rutina_ejercicios SET orden=? WHERE id=?",(_hm2[_pos2+1][1],_rid_e)); _cd.execute("UPDATE rutina_ejercicios SET orden=? WHERE id=?",(_hm2[_pos2][1],_ids2[_pos2+1])); _cd.commit(); _cd.close(); db_query.clear(); st.rerun()
                            if _cb3.button("✏️",key=f"ed_{label}_{_rid_e}",use_container_width=True):
                                _k=f"edit_{label}_{_rid_e}"; st.session_state[_k]=not st.session_state.get(_k,False)
                            if _cb4.button("🗑️",key=f"del_{label}_{_rid_e}",use_container_width=True):
                                _cdel=get_conn(); _cdel.execute("DELETE FROM rutina_ejercicios WHERE id=?",(_rid_e,)); _cdel.commit(); _cdel.close(); db_query.clear(); st.rerun()
                            if st.session_state.get(f"edit_{label}_{_rid_e}",False):
                                with st.form(f"form_ed_{label}_{_rid_e}"):
                                    _f_met=st.selectbox("Método",METODOS_R,index=METODOS_R.index(_rv(_red.get("metodo"),"Normal")) if _rv(_red.get("metodo")) in METODOS_R else 0)
                                    _fe1,_fe2,_fe3,_fe4=st.columns(4)
                                    _f_ser=_fe1.text_input("Series",value=_rv(_red.get("series"),"3"))
                                    _f_rep=_fe2.text_input("Reps",value=_rv(_red.get("repeticiones"),"10-12"))
                                    _f_pes=_fe3.text_input("Carga",value=_rv(_red.get("peso"),""))
                                    _f_tmp=_fe4.text_input("Descanso",value=_rv(_red.get("tempo_descanso"),"60s"))
                                    _f_not=st.text_input("Notas",value=_rv(_red.get("notas"),""))
                                    _fs,_fc=st.columns(2)
                                    _ok_ed=_fs.form_submit_button("💾 Guardar",type="primary",use_container_width=True)
                                    _ca_ed=_fc.form_submit_button("✕ Cancelar",use_container_width=True)
                                if _ok_ed:
                                    _cfe=get_conn(); _cfe.execute("UPDATE rutina_ejercicios SET metodo=?,series=?,repeticiones=?,peso=?,tempo_descanso=?,notas=? WHERE id=?",(_f_met,_f_ser,_f_rep,_f_pes,_f_tmp,_f_not,_rid_e)); _cfe.commit(); _cfe.close()
                                    st.session_state.pop(f"edit_{label}_{_rid_e}",None); db_query.clear(); st.rerun()
                                if _ca_ed: st.session_state.pop(f"edit_{label}_{_rid_e}",None); st.rerun()
                else:
                    st.caption("Sin ejercicios — usa el buscador de abajo.")

                # Buscador por día
                st.markdown(f"<span style='color:{VERDE};font-weight:700'>➕ Agregar a {_dfl2}</span>",unsafe_allow_html=True)
                _ba1,_ba2,_ba3=st.columns([3,1.5,1.5])
                _bej_d=_ba1.text_input("🔍",placeholder="nombre o músculo...",key=f"bej_{label}_{_dfl2}",label_visibility="collapsed")
                _met_d=_ba2.selectbox("Método",METODOS_R,key=f"met_{label}_{_dfl2}",label_visibility="collapsed")
                _sxr=_ba3.text_input("Series×Reps",value="3×10-12",key=f"sxr_{label}_{_dfl2}",label_visibility="collapsed")
                _ej_d=db_query("SELECT * FROM ejercicios WHERE nombre LIKE ? OR musculo_primario LIKE ? ORDER BY nombre LIMIT 18",(f"%{_bej_d}%",f"%{_bej_d}%")) if _bej_d.strip() else db_query("SELECT * FROM ejercicios ORDER BY nombre LIMIT 18")
                if not _ej_d.empty:
                    for _ch in range(0,len(_ej_d),6):
                        _ecols=st.columns(6)
                        for _ei,(_,_efr) in enumerate(_ej_d.iloc[_ch:_ch+6].iterrows()):
                            _efd=_efr.to_dict()
                            with _ecols[_ei]:
                                with st.container(border=True):
                                    if _hv(_efd.get("url_imagen")):
                                        try: st.image(_efd["url_imagen"],use_container_width=True)
                                        except: pass
                                    st.markdown(f"<div style='font-size:.72rem;font-weight:700;color:{BLANCO};text-align:center'>{_efd['nombre'][:20]}</div>",unsafe_allow_html=True)
                                    if st.button("➕",key=f"add_{label}_{_dfl2}_{_efd['id']}",use_container_width=True,help=_efd['nombre']):
                                        try:
                                            _parts=_sxr.replace("×","x").replace("X","x").split("x")
                                            _ser_v=_parts[0].strip(); _rep_v=_parts[1].strip() if len(_parts)>1 else "10-12"
                                        except: _ser_v="3"; _rep_v="10-12"
                                        _mo=get_conn().execute("SELECT COALESCE(MAX(orden),-1) FROM rutina_ejercicios WHERE rutina_id=? AND dia_semana=?",(rut_id,_dfl2)).fetchone()[0]
                                        _ca=get_conn(); _ca.execute("INSERT INTO rutina_ejercicios (rutina_id,ejercicio_id,dia_semana,orden,metodo,series,repeticiones,peso,tempo_descanso,notas) VALUES (?,?,?,?,?,?,?,?,?,?)",(rut_id,int(_efd["id"]),_dfl2,_mo+1,_met_d,_ser_v,_rep_v,"","60s","")); _ca.commit(); _ca.close(); db_query.clear(); st.rerun()

    # ── 3 TABS PRINCIPALES ─────────────────────────────────────────────────
    tab_crear,tab_editar,tab_guardadas=st.tabs(["➕ Crear rutina","✏️ Editar rutina","📚 Rutinas guardadas"])

    # ══════════════════════════════════════════════════════════════════
    # TAB 1 — CREAR RUTINA
    # ══════════════════════════════════════════════════════════════════
    with tab_crear:
        _p1,_p2=st.columns([1,3])
        _tipo_r=_p1.radio("Tipo",["👤 Cliente","📋 Plantilla"],horizontal=False,key="rut_tipo")

        if _tipo_r=="👤 Cliente":
            _cli_rut=st.selectbox("Cliente",["(seleccionar)"]+df_act["rut"].tolist(),
                format_func=lambda x:"(seleccionar)" if x=="(seleccionar)" else f"{df_act[df_act['rut']==x]['nombre'].iloc[0]} — {x}",
                key="rut_cli_sel")
            _cli_nom="" if _cli_rut=="(seleccionar)" else df_act[df_act["rut"]==_cli_rut]["nombre"].iloc[0]
            if st.session_state.get("_rut_last_cli")!=_cli_rut:
                st.session_state["_rut_last_cli"]=_cli_rut; st.session_state.pop("_rut_id_activo",None)
        else:
            _cli_rut=None; _cli_nom="Plantilla"
            if st.session_state.get("_rut_last_cli")!="__PLAN__":
                st.session_state["_rut_last_cli"]="__PLAN__"; st.session_state.pop("_rut_id_activo",None)

        _rut_id=st.session_state.get("_rut_id_activo")
        if not _rut_id and _cli_rut and _cli_rut!="(seleccionar)":
            _r_ex=db_query("SELECT id FROM rutinas WHERE cliente_rut=? AND activa=1 ORDER BY id DESC LIMIT 1",(_cli_rut,))
            if not _r_ex.empty:
                _rut_id=int(_r_ex.iloc[0]["id"]); st.session_state["_rut_id_activo"]=_rut_id

        _nb1,_nb2,_nb3=st.columns([3,1,1])
        _rut_nom=_nb1.text_input("Nombre de la rutina",placeholder="Ej: Fuerza Julio · Full Body Principiante",key="rut_nom_inp")
        _crear_ok=_nb2.button("💾 Crear",type="primary",use_container_width=True,key="rut_btn_crear")
        _nueva_ok=_nb3.button("🔄 En blanco",use_container_width=True,key="rut_btn_nueva")

        if _crear_ok:
            if not _rut_nom.strip(): st.warning("Escribe un nombre.")
            elif _tipo_r=="👤 Cliente" and _cli_rut in (None,"(seleccionar)"): st.warning("Selecciona un cliente.")
            else:
                _cn_r=get_conn()
                _rut_id_exist=st.session_state.get("_rut_id_activo")
                if _rut_id_exist:
                    # Ya existe — solo actualizar nombre, NO duplicar
                    _cn_r.execute("UPDATE rutinas SET nombre=? WHERE id=?",(_rut_nom.strip(),_rut_id_exist))
                    _cn_r.commit(); _cn_r.close()
                    db_query.clear(); st.success(f"✅ Nombre actualizado a '{_rut_nom}'."); st.rerun()
                else:
                    if _cli_rut: _cn_r.execute("UPDATE rutinas SET activa=0 WHERE cliente_rut=?",(_cli_rut,))
                    _new_id=_cn_r.execute("INSERT INTO rutinas (cliente_rut,nombre,activa) VALUES (?,?,1)",(_cli_rut,_rut_nom.strip())).lastrowid
                    _cn_r.commit(); _cn_r.close()
                    st.session_state["_rut_id_activo"]=int(_new_id); db_query.clear(); st.rerun()

        if _nueva_ok and _cli_rut not in (None,"(seleccionar)"):
            _cn_nb=get_conn()
            if _cli_rut: _cn_nb.execute("UPDATE rutinas SET activa=0 WHERE cliente_rut=?",(_cli_rut,))
            _new_id2=_cn_nb.execute("INSERT INTO rutinas (cliente_rut,nombre,activa) VALUES (?,?,1)",(_cli_rut,"Nueva rutina")).lastrowid
            _cn_nb.commit(); _cn_nb.close()
            st.session_state["_rut_id_activo"]=int(_new_id2); db_query.clear(); st.rerun()

        if not _rut_id:
            st.markdown(f'<div class="info-box">{"Selecciona un cliente o crea una rutina nueva." if _tipo_r=="👤 Cliente" else "Escribe un nombre y presiona Crear."}</div>',unsafe_allow_html=True)
        else:
            _ri=db_query("SELECT nombre FROM rutinas WHERE id=?",(_rut_id,))
            if not _ri.empty:
                _rb1,_rb2=st.columns([5,1])
                _rb1.markdown(f'<div style="background:{GRIS2};border-radius:8px;padding:7px 14px"><span style="color:{VERDE};font-weight:700">✅ Editando: {_rv(_ri.iloc[0]["nombre"])}</span></div>',unsafe_allow_html=True)
                if _rb2.button("✕ Cerrar",key="rut_cerrar",use_container_width=True):
                    st.session_state.pop("_rut_id_activo",None); st.rerun()
            st.divider()
            _render_editor(_rut_id,label="cr")

    # ══════════════════════════════════════════════════════════════════
    # TAB 2 — EDITAR RUTINA
    # ══════════════════════════════════════════════════════════════════
    with tab_editar:
        _rut_ed_id=st.session_state.get("_rut_ed_id")
        if _rut_ed_id:
            _ri_ed=db_query("SELECT * FROM rutinas WHERE id=?",(_rut_ed_id,))
            if not _ri_ed.empty:
                _re_d=_ri_ed.iloc[0].to_dict()
                _nom_c_ed=df_act[df_act["rut"]==_re_d.get("cliente_rut")]["nombre"].iloc[0] if _re_d.get("cliente_rut") and _re_d.get("cliente_rut") in df_act["rut"].values else "Plantilla"
                _red_c1,_red_c2=st.columns([5,1])
                _red_c1.markdown(f'<div style="background:{GRIS2};border-radius:8px;padding:8px 14px"><span style="color:{VERDE};font-weight:700">✏️ {_rv(_re_d["nombre"])}</span> <span style="color:{GRIS_T};font-size:.82rem">· {_nom_c_ed}</span></div>',unsafe_allow_html=True)
                if _red_c2.button("← Volver",key="ed_cerrar",use_container_width=True):
                    st.session_state.pop("_rut_ed_id",None); st.rerun()
                _render_editor(_rut_ed_id,label="ed")
            else:
                st.session_state.pop("_rut_ed_id",None); st.rerun()
        else:
            st.markdown(f"<b style='color:{VERDE}'>Selecciona una rutina para editar</b>",unsafe_allow_html=True)
            _ef1,_ef2=st.columns(2)
            _etxt=_ef1.text_input("🔍",placeholder="nombre o cliente...",key="ed_busq",label_visibility="collapsed")
            _etipo=_ef2.radio("Tipo",["Todas","Clientes","Plantillas"],horizontal=True,key="ed_tipo")
            _todas_ed=db_query("""SELECT r.id,r.nombre,r.activa,r.cliente_rut,
                (SELECT COUNT(*) FROM rutina_ejercicios re WHERE re.rutina_id=r.id) as n_ej
                FROM rutinas r ORDER BY r.fecha_creacion DESC""")
            if not _todas_ed.empty:
                if _etipo=="Clientes": _todas_ed=_todas_ed[_todas_ed["cliente_rut"].notna()&(_todas_ed["cliente_rut"]!="")]
                elif _etipo=="Plantillas": _todas_ed=_todas_ed[_todas_ed["cliente_rut"].isna()|(_todas_ed["cliente_rut"]=="")]
                if _etxt.strip():
                    _todas_ed=_todas_ed[_todas_ed["nombre"].str.contains(_etxt,case=False,na=False)|
                        _todas_ed["cliente_rut"].fillna("").apply(lambda x:_etxt.lower() in (df_act[df_act["rut"]==x]["nombre"].iloc[0].lower() if x in df_act["rut"].values else ""))]
            if not _todas_ed.empty:
                for _,_gr in _todas_ed.iterrows():
                    _grd=_gr.to_dict()
                    _nom_c=df_act[df_act["rut"]==_grd.get("cliente_rut")]["nombre"].iloc[0] if _grd.get("cliente_rut") and _grd.get("cliente_rut") in df_act["rut"].values else "📋 Plantilla"
                    _act_badge="✅" if _grd.get("activa") else "📦"
                    # Detectar cuántos días tiene
                    _dias_rutina=db_query("SELECT DISTINCT dia_semana FROM rutina_ejercicios WHERE rutina_id=? ORDER BY dia_semana",(int(_grd["id"]),))
                    _n_dias_tag=f"{len(_dias_rutina)} día(s)" if not _dias_rutina.empty else "sin días"
                    _ec1,_ec2=st.columns([4,1])
                    _ec1.markdown(f'<div style="background:{GRIS2};border-radius:8px;padding:7px 12px;"><b>{_grd["nombre"]}</b> {_act_badge} <span style="color:{GRIS_T};font-size:.82rem">· {_nom_c} · {_grd["n_ej"]} ej. · {_n_dias_tag}</span></div>',unsafe_allow_html=True)
                    if _ec2.button("✏️ Editar",key=f"sel_ed_{_grd['id']}",use_container_width=True):
                        st.session_state["_rut_ed_id"]=int(_grd["id"]); st.rerun()
            else:
                st.markdown('<div class="info-box">Sin rutinas guardadas.</div>',unsafe_allow_html=True)

    # ══════════════════════════════════════════════════════════════════
    # TAB 3 — RUTINAS GUARDADAS
    # ══════════════════════════════════════════════════════════════════
    with tab_guardadas:
        _g1,_g2=st.columns(2)
        _fil_tipo=_g1.radio("Mostrar",["Todas","Solo clientes","Solo plantillas"],horizontal=True,key="rut_fil_tipo")
        _fil_txt=_g2.text_input("Buscar",placeholder="nombre o cliente...",key="rut_fil_txt",label_visibility="collapsed")
        _todas_g=db_query("""SELECT r.id,r.nombre,r.activa,r.cliente_rut,r.fecha_creacion,
            (SELECT COUNT(*) FROM rutina_ejercicios re WHERE re.rutina_id=r.id) as n_ej
            FROM rutinas r ORDER BY r.fecha_creacion DESC""")
        if _todas_g.empty:
            st.markdown('<div class="info-box">Sin rutinas guardadas.</div>',unsafe_allow_html=True)
        else:
            if _fil_tipo=="Solo clientes": _todas_g=_todas_g[_todas_g["cliente_rut"].notna()&(_todas_g["cliente_rut"]!="")]
            elif _fil_tipo=="Solo plantillas": _todas_g=_todas_g[_todas_g["cliente_rut"].isna()|(_todas_g["cliente_rut"]=="")]
            if _fil_txt.strip():
                _todas_g=_todas_g[_todas_g["nombre"].str.contains(_fil_txt,case=False,na=False)|
                    _todas_g["cliente_rut"].fillna("").apply(lambda x:_fil_txt.lower() in (df_act[df_act["rut"]==x]["nombre"].iloc[0].lower() if x in df_act["rut"].values else ""))]
            for _,_gr in _todas_g.iterrows():
                _grd=_gr.to_dict()
                _nom_c=df_act[df_act["rut"]==_grd.get("cliente_rut")]["nombre"].iloc[0] if _grd.get("cliente_rut") and _grd.get("cliente_rut") in df_act["rut"].values else "📋 Plantilla"
                _act_b="✅" if _grd.get("activa") else "📦"
                _dias_g=db_query("SELECT DISTINCT dia_semana FROM rutina_ejercicios WHERE rutina_id=? ORDER BY dia_semana",(int(_grd["id"]),))
                _dias_g_tag=", ".join(_dias_g["dia_semana"].tolist()) if not _dias_g.empty else "sin días"
                _gc1,_gc3,_gc4,_gc5=st.columns([3,1,1,1])
                _gc1.markdown(f'<div style="background:{GRIS2};border-radius:8px;padding:7px 12px;"><b>{_grd["nombre"]}</b> {_act_b}<br><span style="color:{GRIS_T};font-size:.8rem">👤 {_nom_c} · {_grd["n_ej"]} ej. · {_dias_g_tag} · {fmt_fecha(_rv(_grd.get("fecha_creacion")))}</span></div>',unsafe_allow_html=True)
                if _gc3.button("📋",key=f"g_dup_{_grd['id']}",use_container_width=True,help="Duplicar"):
                    _dup=get_conn()
                    _new_dup=_dup.execute("INSERT INTO rutinas (cliente_rut,nombre,activa) VALUES (NULL,?,0)",(f"{_grd['nombre']} (copia)",)).lastrowid
                    _dup.commit()
                    _ejs_dup=_dup.execute("SELECT * FROM rutina_ejercicios WHERE rutina_id=? ORDER BY dia_semana,orden",(int(_grd["id"]),)).fetchall()
                    for _ed in _ejs_dup:
                        _dup.execute("INSERT INTO rutina_ejercicios (rutina_id,ejercicio_id,dia_semana,orden,metodo,series,repeticiones,peso,tempo_descanso,notas) VALUES (?,?,?,?,?,?,?,?,?,?)",
                            (_new_dup,_ed[2],_ed[3],_ed[4],_ed[5],_ed[6],_ed[7],_ed[8],_ed[9],_ed[10]))
                    _dup.commit(); _dup.close(); db_query.clear(); st.success("✅ Duplicada"); st.rerun()
                if _gc4.button("👤",key=f"g_asig_{_grd['id']}",use_container_width=True,help="Asignar a cliente"):
                    st.session_state["asig_id_g"]=int(_grd["id"])
                if st.session_state.get("asig_id_g")==int(_grd["id"]):
                    _dest_g=st.selectbox("Asignar a:",["(seleccionar)"]+df_act["rut"].tolist(),
                        format_func=lambda x:"(seleccionar)" if x=="(seleccionar)" else f"{df_act[df_act['rut']==x]['nombre'].iloc[0]} — {x}",
                        key=f"dest_g_{_grd['id']}")
                    _dg1,_dg2=st.columns(2)
                    if _dg1.button("✅ Confirmar",key=f"ok_asig_g_{_grd['id']}") and _dest_g!="(seleccionar)":
                        _cag=get_conn(); _cag.execute("UPDATE rutinas SET activa=0 WHERE cliente_rut=?",(_dest_g,))
                        _nag=_cag.execute("INSERT INTO rutinas (cliente_rut,nombre,activa) VALUES (?,?,1)",(_dest_g,_grd["nombre"])).lastrowid
                        _cag.commit()
                        _ejs_ag=_cag.execute("SELECT * FROM rutina_ejercicios WHERE rutina_id=? ORDER BY dia_semana,orden",(int(_grd["id"]),)).fetchall()
                        for _eag in _ejs_ag:
                            _cag.execute("INSERT INTO rutina_ejercicios (rutina_id,ejercicio_id,dia_semana,orden,metodo,series,repeticiones,peso,tempo_descanso,notas) VALUES (?,?,?,?,?,?,?,?,?,?)",
                                (_nag,_eag[2],_eag[3],_eag[4],_eag[5],_eag[6],_eag[7],_eag[8],_eag[9],_eag[10]))
                        _cag.commit(); _cag.close(); db_query.clear()
                        st.session_state.pop("asig_id_g",None); st.success("✅ Asignada"); st.rerun()
                    if _dg2.button("✕",key=f"cancel_asig_g_{_grd['id']}"):
                        st.session_state.pop("asig_id_g",None); st.rerun()
                if _gc5.button("🗑️",key=f"g_del_{_grd['id']}",use_container_width=True,help="Eliminar"):
                    st.session_state[f"g_del_c_{_grd['id']}"]=True
                if st.session_state.get(f"g_del_c_{_grd['id']}"):
                    _gd1,_gd2=st.columns(2)
                    if _gd1.button("✅ Sí, eliminar",key=f"g_del_yes_{_grd['id']}"):
                        _cdg=get_conn(); _cdg.execute("DELETE FROM rutina_ejercicios WHERE rutina_id=?",(int(_grd["id"]),)); _cdg.execute("DELETE FROM rutinas WHERE id=?",(int(_grd["id"]),)); _cdg.commit(); _cdg.close()
                        st.session_state.pop(f"g_del_c_{_grd['id']}",None); db_query.clear(); st.rerun()
                    if _gd2.button("❌ Cancelar",key=f"g_del_no_{_grd['id']}"):
                        st.session_state.pop(f"g_del_c_{_grd['id']}",None); st.rerun()

elif pagina=="⚙️ Base de Datos":
    st.markdown('<div class="section-header">⚙️ Base de Datos Unificada</div>',unsafe_allow_html=True)
    if st.button("← Volver",key="db_volver"): st.session_state._goto="🏠 Dashboard"; st.rerun()
    _es_admin_bd=st.session_state.get("rol","").lower() in ("admin","administrador")
    if _es_admin_bd:
        tim,tex,tad,tusr,tlog=st.tabs(["📥 Importar","📤 Exportar","🗄️ Administración","👥 Usuarios","📋 Log"])
    else:
        tim,tex,tad=st.tabs(["📥 Importar","📤 Exportar","🗄️ Administración"])
    with tim:
        st.markdown("""<div class="info-box">
          <b>Dos modos de importación:</b><br>
          1. <b>Excel unificado</b> (exportado desde este sistema) — restaura todas las tablas.<br>
          2. <b>datosgym.xlsx original</b> — importa solo la hoja "BBDD Clientes". Los pases diarios se excluyen.
        </div>""",unsafe_allow_html=True)
        modo=st.radio("Modo",["Excel unificado (sistema)","datosgym.xlsx original"])
        up=st.file_uploader("Subir archivo Excel",type=["xlsx","xls"])
        if up and st.button("📥 Importar"):
            try:
                if "unificado" in modo: n=importar_excel_full(up)
                else: n=importar_excel_clientes_legacy(up)
                st.cache_data.clear()
                st.markdown(f'<div class="success-box">✅ Importados {n} registros.</div>',unsafe_allow_html=True)
            except Exception as ex: st.error(f"Error: {ex}")
    with tex:
        st.markdown("Exporta **toda la base de datos** en un único archivo Excel con múltiples hojas.")
        xls=exportar_todo_excel()
        st.download_button("📥 Descargar base completa Excel",xls,"putu_activo_completo.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
        st.markdown(f'<div class="info-box">El archivo incluye: clientes, pagos, asistencia, evaluaciones, clases, egresos, productos.</div>',unsafe_allow_html=True)
    with tad:
        if st.session_state.rol=="Administrador":
            tots={t:int(db_query(f"SELECT COUNT(*) as n FROM {t}").iloc[0]["n"]) for t in ["clientes","pagos","asistencia","evaluaciones","clases","egresos","productos"]}
            c=st.columns(4)
            for i,(t,n) in enumerate(tots.items()): c[i%4].metric(t.capitalize(),n)
            st.markdown(f"**DB:** `{DB_PATH}`")
        else: st.markdown('<div class="info-box">Solo Administradores.</div>',unsafe_allow_html=True)

    # ── TAB USUARIOS (solo admin) ─────────────────────────────────────────
    if _es_admin_bd:
        with tusr:
            st.markdown(f"<b style='color:{VERDE}'>👥 Gestión de usuarios del sistema</b>",unsafe_allow_html=True)
            _TODOS_ACCESOS=["🏠 Dashboard","👥 Clientes","💳 Pagos y Renovaciones","✅ Asistencia",
                "🏃 Clases & Talleres","🛍 Venta Productos","💪 Ejercicios","📋 Rutinas",
                "📊 Reportes","⚙️ Base de Datos"]
            _ACCESOS_DEF={
                "admin":         _TODOS_ACCESOS[:],
                "Administrador": _TODOS_ACCESOS[:],
                "entrenador":    ["🏠 Dashboard","👥 Clientes","✅ Asistencia","💪 Ejercicios","📋 Rutinas","📊 Reportes"],
                "asistente":     ["🏠 Dashboard","👥 Clientes","💳 Pagos y Renovaciones","✅ Asistencia","🏃 Clases & Talleres","🛍 Venta Productos","📊 Reportes"],
            }
            _us_df=db_query("SELECT id,usuario,nombre,rol,activo FROM usuarios_sistema ORDER BY id")

            if not _us_df.empty:
                # Construir dataframe editable
                import pandas as _pd_usr
                _edit_rows=[]
                for _,_ur in _us_df.iterrows():
                    _urd=_ur.to_dict()
                    _acc_rol=_ACCESOS_DEF.get(_urd['rol'],_TODOS_ACCESOS)
                    _row={"id":int(_urd['id']),"Usuario":_urd['usuario'],"Nombre":_urd['nombre'],
                          "Rol":_urd['rol'],"Activo":bool(_urd['activo'])}
                    for _acc in _TODOS_ACCESOS:
                        _row[_acc.split(" ",1)[-1][:12]]=bool(_acc in _acc_rol)
                    _edit_rows.append(_row)
                _edit_df=_pd_usr.DataFrame(_edit_rows)

                # Columnas configuración para data_editor
                _col_cfg={"id":st.column_config.NumberColumn("ID",disabled=True,width="small"),
                    "Usuario":st.column_config.TextColumn("Usuario",disabled=True,width="small"),
                    "Nombre":st.column_config.TextColumn("Nombre",width="medium"),
                    "Rol":st.column_config.SelectboxColumn("Rol",options=["admin","entrenador","asistente"],width="small"),
                    "Activo":st.column_config.CheckboxColumn("Activo",width="small")}
                for _acc in _TODOS_ACCESOS:
                    _k=_acc.split(" ",1)[-1][:12]
                    _col_cfg[_k]=st.column_config.CheckboxColumn(_k,width="small")

                _edited=st.data_editor(_edit_df,column_config=_col_cfg,
                    use_container_width=True,hide_index=True,
                    num_rows="fixed",key="usr_table_editor")

                # Botón guardar cambios
                if st.button("💾 Guardar cambios en tabla",key="usr_save",type="primary"):
                    _cu=get_conn()
                    for _,_er in _edited.iterrows():
                        _eid=int(_er["id"])
                        _enom=_er["Nombre"]; _erol=_er["Rol"]; _eact=int(_er["Activo"])
                        _cu.execute("UPDATE usuarios_sistema SET nombre=?,rol=?,activo=? WHERE id=?",
                            (_enom,_erol,_eact,_eid))
                    _cu.commit(); _cu.close()
                    log_action("USUARIOS_EDIT","Tabla de usuarios actualizada")
                    db_query.clear(); st.success("✅ Cambios guardados"); st.rerun()

                # Editar clave individualmente
                with st.expander("🔑 Cambiar contraseña de un usuario"):
                    _pw1,_pw2,_pw3=st.columns(3)
                    _pw_usr=_pw1.selectbox("Usuario",_us_df["usuario"].tolist(),key="pw_sel")
                    _pw_new=_pw2.text_input("Nueva contraseña",type="password",key="pw_new")
                    if _pw3.button("💾 Cambiar clave",key="pw_save",use_container_width=True):
                        if len(_pw_new)>=4:
                            _cpu=get_conn(); _cpu.execute("UPDATE usuarios_sistema SET password_hash=? WHERE usuario=?",(_h(_pw_new),_pw_usr)); _cpu.commit(); _cpu.close()
                            log_action("CLAVE_CAMBIO",f"Clave cambiada para {_pw_usr}")
                            st.success(f"✅ Clave de {_pw_usr} actualizada")
                        else: st.warning("Mínimo 4 caracteres")

                # Eliminar usuario
                with st.expander("🗑️ Eliminar usuario"):
                    _del_opts=[u for u in _us_df["usuario"].tolist() if u!="admin"]
                    if _del_opts:
                        _del_u=st.selectbox("Usuario a eliminar",_del_opts,key="del_usr_sel")
                        if st.button("🗑️ Eliminar",key="del_usr_btn"):
                            _cdu=get_conn(); _cdu.execute("DELETE FROM usuarios_sistema WHERE usuario=?",(_del_u,)); _cdu.commit(); _cdu.close()
                            log_action("USUARIO_ELIMINAR",f"Usuario {_del_u} eliminado")
                            db_query.clear(); st.rerun()
                    else:
                        st.caption("No hay usuarios eliminables.")

            st.divider()
            # ── Crear nuevo usuario ──────────────────────────────────────────
            st.markdown(f"<b style='color:{VERDE}'>➕ Crear nuevo usuario</b>",unsafe_allow_html=True)
            with st.form("form_new_usr",clear_on_submit=True):
                _nu1,_nu2=st.columns(2)
                _nu_usr=_nu1.text_input("Usuario (login)",placeholder="juan.perez")
                _nu_nom=_nu2.text_input("Nombre completo",placeholder="Juan Pérez")
                _nu3,_nu4=st.columns(2)
                _ROLES=["admin","entrenador","asistente"]
                _nu_rol=_nu3.selectbox("Rol",_ROLES)
                _nu_pas=_nu4.text_input("Contraseña",type="password")
                _ok_nu=st.form_submit_button("➕ Crear usuario",type="primary",use_container_width=True)
            if _ok_nu and _nu_usr.strip() and _nu_nom.strip() and len(_nu_pas)>=4:
                try:
                    _cnu=get_conn()
                    _cnu.execute("INSERT INTO usuarios_sistema (usuario,nombre,password_hash,rol,activo) VALUES (?,?,?,?,1)",
                        (_nu_usr.strip().lower(),_nu_nom.strip(),_h(_nu_pas),_nu_rol))
                    _cnu.commit(); _cnu.close()
                    log_action("USUARIO_CREAR",f"Nuevo usuario {_nu_usr} rol={_nu_rol}")
                    st.success(f"✅ Usuario '{_nu_usr}' creado"); st.rerun()
                except Exception as _eu: st.error(f"Error: {_eu}")

        with tlog:
            st.markdown(f"<b style='color:{VERDE}'>📋 Log de actividad</b>",unsafe_allow_html=True)
            _lg1,_lg2,_lg3=st.columns(3)
            _lg_usr=_lg1.selectbox("Usuario",["Todos"]+[r[0] for r in get_conn().execute("SELECT DISTINCT usuario FROM activity_log ORDER BY usuario").fetchall()],key="lg_usr")
            _lg_acc=_lg2.selectbox("Acción",["Todas"]+[r[0] for r in get_conn().execute("SELECT DISTINCT accion FROM activity_log ORDER BY accion").fetchall()],key="lg_acc")
            _lg_lim=_lg3.selectbox("Mostrar",[50,100,250,500],key="lg_lim")
            _lg_w=[]; _lg_p=[]
            if _lg_usr!="Todos": _lg_w.append("usuario=?"); _lg_p.append(_lg_usr)
            if _lg_acc!="Todas": _lg_w.append("accion=?"); _lg_p.append(_lg_acc)
            _lg_q="SELECT timestamp,usuario,rol,accion,detalle,rut_afectado FROM activity_log"+(" WHERE "+" AND ".join(_lg_w) if _lg_w else "")+" ORDER BY id DESC LIMIT ?"
            _lg_p.append(_lg_lim)
            _lg_df=db_query(_lg_q,tuple(_lg_p))
            if not _lg_df.empty:
                _lg_df["timestamp"]=_lg_df["timestamp"].apply(lambda x:x[:19] if x else "")
                st.dataframe(_lg_df.rename(columns={"timestamp":"Fecha/Hora","usuario":"Usuario","rol":"Rol","accion":"Acción","detalle":"Detalle","rut_afectado":"RUT"}),
                    use_container_width=True,hide_index=True,height=500)
                st.caption(f"{len(_lg_df)} registros mostrados")
                import io as _io_lg
                _xls_lg=_io_lg.BytesIO()
                _lg_df.to_excel(_xls_lg,index=False)
                st.download_button("⬇️ Exportar log Excel",_xls_lg.getvalue(),"activity_log.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
            else:
                st.markdown('<div class="info-box">Sin registros de actividad aún.</div>',unsafe_allow_html=True)
