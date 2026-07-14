import os
import socket
import csv
import io
import uuid
import re
import pandas as pd
from datetime import datetime, timedelta
from functools import wraps
from PIL import Image

from flask import Flask, render_template, request, redirect, url_for, flash, session, Response, send_file
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from flask_sqlalchemy import SQLAlchemy
from jinja2 import DictLoader
from sqlalchemy import event, or_, and_
from sqlalchemy.engine import Engine

# ==============================================================================
# CONFIGURACIÓN DE LA APLICACIÓN Y OPTIMIZACIONES MULTIUSUARIO
# ==============================================================================
app = Flask(__name__)
app.config['SECRET_KEY'] = os.environ.get('FLASK_SECRET_KEY', 'clave_super_secreta_mvc_2026')

# Seguridad de Sesión
app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(minutes=60)
app.config['SESSION_COOKIE_HTTPONLY'] = True
app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'

# Límite de tamaño de subida de archivos (16 MB máximo)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

# Base de datos con ruta absoluta segura
DB_FILENAME = 'registro_afectados_definitivo.db' 
DB_PATH = os.path.join(app.root_path, DB_FILENAME)
app.config['SQLALCHEMY_DATABASE_URI'] = f'sqlite:///{DB_PATH}'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

# Optimización para concurrencia multiusuario
app.config['SQLALCHEMY_ENGINE_OPTIONS'] = {
    'connect_args': {'timeout': 15}
}

app.config['UPLOAD_FOLDER'] = os.path.join(app.root_path, 'static', 'uploads')
ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'webp', 'xlsx', 'xls', 'csv'}

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

db = SQLAlchemy(app)

@event.listens_for(Engine, "connect")
def set_sqlite_pragma(dbapi_connection, connection_record):
    cursor = dbapi_connection.cursor()
    cursor.execute("PRAGMA journal_mode=WAL")
    cursor.execute("PRAGMA synchronous=NORMAL")
    cursor.execute("PRAGMA cache_size=-10000")
    cursor.execute("PRAGMA foreign_keys=ON")
    cursor.close()

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def comprimir_y_guardar_foto(file, folder):
    if not file: return None
    filename = secure_filename(file.filename)
    unique_name = f"{uuid.uuid4().hex}_{filename}"
    filepath = os.path.join(folder, unique_name)
    try:
        img = Image.open(file)
        img.thumbnail((800, 800))
        if img.mode != 'RGB':
            img = img.convert('RGB')
        img.save(filepath, format='JPEG', quality=85)
        return unique_name
    except:
        return None

# ==============================================================================
# SISTEMAS DE CACHÉ Y SEGURIDAD
# ==============================================================================
CACHE_DICT = {}

def clear_stats_cache():
    CACHE_DICT.pop('stats', None)

@app.before_request
def ensure_csrf():
    if 'csrf_token' not in session:
        session['csrf_token'] = uuid.uuid4().hex

    if request.method == "POST":
        if request.endpoint == 'logout': 
            return
            
        token = session.get('csrf_token')
        form_token = request.form.get('csrf_token')
        
        if not token or token != form_token:
            flash('Error de seguridad (Sesión expirada o intento no autorizado).', 'error')
            return redirect(request.referrer or url_for('login'))

@app.context_processor
def inject_csrf():
    return dict(csrf_token=session.get('csrf_token', ''))

# ==============================================================================
# MOTOR LECTOR UNIVERSAL DE EXCEL
# ==============================================================================
def leer_excel_dinamico(file):
    df_dict = {}
    try:
        if file.filename.endswith('.csv'):
            file.seek(0)
            df_raw = pd.read_csv(file, header=None, dtype=str)
            header_idx = 0
            for i, r in df_raw.head(20).iterrows():
                rt = ' '.join(str(v).upper() for v in r.values)
                if any(x in rt for x in ['CEDULA', 'CÉDULA', 'C.I', 'NOMBRE', 'TRABAJADOR', 'PARENTESCO']):
                    header_idx = i
                    break
            file.seek(0)
            df = pd.read_csv(file, header=header_idx, dtype=str)
            df.columns = df.columns.astype(str).str.strip().str.upper()
            df_dict['DATOS_CSV'] = df
        else:
            file.seek(0)
            xls = pd.ExcelFile(file)
            for sheet in xls.sheet_names:
                df_raw = pd.read_excel(xls, sheet_name=sheet, header=None, dtype=str)
                if df_raw.empty: continue
                header_idx = 0
                for i, r in df_raw.head(20).iterrows():
                    rt = ' '.join(str(v).upper() for v in r.values)
                    if any(x in rt for x in ['CEDULA', 'CÉDULA', 'C.I', 'NOMBRE', 'TRABAJADOR', 'PARENTESCO']):
                        header_idx = i
                        break
                df = pd.read_excel(xls, sheet_name=sheet, header=header_idx, dtype=str)
                df.columns = df.columns.astype(str).str.strip().str.upper()
                df_dict[sheet] = df
    except Exception as e:
        print(f"Error procesando el archivo dinámico: {e}")
    return df_dict

def parse_age(age_val):
    s = str(age_val).strip().upper()
    if s in ['NAN', 'NONE', '']: return -1 
    if 'MES' in s: return 0 
    m = re.search(r'\d+', s)
    return int(m.group()) if m else -1

# ==============================================================================
# MODELOS DE BASE DE DATOS
# ==============================================================================
class Usuario(db.Model):
    username = db.Column(db.String(50), primary_key=True)
    password_hash = db.Column(db.String(200), nullable=False)
    rol = db.Column(db.String(20), nullable=False)
    acceso_modulos = db.Column(db.String(50), default='TOTAL')

class Parroquia(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    nombre = db.Column(db.String(100), nullable=False, unique=True)

class Configuracion(db.Model):
    clave = db.Column(db.String(50), primary_key=True)
    valor = db.Column(db.String(255), default='')

class JefeFamilia(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    nombre = db.Column(db.String(150), nullable=False, index=True) 
    ci = db.Column(db.String(20), index=True) 
    edad = db.Column(db.Integer)
    genero = db.Column(db.String(1))
    telefono = db.Column(db.String(20))
    situacion = db.Column(db.String(50), nullable=False)
    parroquia = db.Column(db.String(100), nullable=False)
    requerimiento = db.Column(db.Text, default='NINGUNO')
    observacion = db.Column(db.Text, default='NINGUNA') 
    lugar_remision = db.Column(db.String(200), default='') 
    discapacidad = db.Column(db.String(100), default='NINGUNA')
    patologia = db.Column(db.String(100), default='NINGUNA')
    foto = db.Column(db.String(200))
    foto_vivienda = db.Column(db.String(200))
    es_embarazada = db.Column(db.Integer, default=0) 
    fecha_registro = db.Column(db.String(50), nullable=False)
    usuario_registra = db.Column(db.String(50), nullable=False)
    
    cargas = db.relationship('NucleoFamiliar', backref='jefe', lazy=True, cascade="all, delete-orphan")

class NucleoFamiliar(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    jefe_id = db.Column(db.Integer, db.ForeignKey('jefe_familia.id'), nullable=False, index=True)
    nombre = db.Column(db.String(150), nullable=False, index=True) 
    ci = db.Column(db.String(20), default='')
    parentesco = db.Column(db.String(50), nullable=False)
    edad = db.Column(db.Integer)
    genero = db.Column(db.String(1))
    discapacidad = db.Column(db.String(100), default='NINGUNA')
    patologia = db.Column(db.String(100), default='NINGUNA')
    es_embarazada = db.Column(db.Integer, default=0)

class HistorialCambio(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    usuario = db.Column(db.String(50), nullable=False)
    accion = db.Column(db.String(100), nullable=False)
    fecha = db.Column(db.DateTime, default=datetime.now)
    detalle = db.Column(db.Text)

# ==============================================================================
# SEGURIDAD Y AUDITORÍA
# ==============================================================================
def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'username' not in session:
            flash('Debes iniciar sesión para acceder.', 'error')
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function

def role_required(*roles):
    def decorator(f):
        @wraps(f)
        def decorated_function(*args, **kwargs):
            if 'username' not in session: return redirect(url_for('login'))
            if session.get('rol') not in roles:
                flash(f'Acceso denegado. No tienes permisos para esta área.', 'error')
                return redirect(url_for('index'))
            return f(*args, **kwargs)
        return decorated_function
    return decorator

def registrar_auditoria(usuario, accion, detalle):
    auditoria = HistorialCambio(usuario=usuario, accion=accion, detalle=detalle)
    db.session.add(auditoria)

def inferir_genero(nombre, parentesco, genero_previo):
    gen = str(genero_previo).strip().upper()
    if gen in ['M', 'F']: return gen
    
    par = str(parentesco).strip().upper()
    femeninos = ['HIJA', 'ESPOSA', 'MADRE', 'ABUELA', 'TIA', 'SOBRINA', 'NIETA', 'HERMANA', 'JEFA', 'MAMA', 'MUJER']
    masculinos = ['HIJO', 'ESPOSO', 'PADRE', 'ABUELO', 'TIO', 'SOBRINO', 'NIETO', 'HERMANO', 'JEFE', 'PAPA', 'HOMBRE']
    
    if any(f in par for f in femeninos): return 'F'
    if any(m in par for m in masculinos): return 'M'
    
    if nombre:
        primer_nombre = nombre.split()[0]
        if primer_nombre.endswith('A') or primer_nombre.endswith('YS'): return 'F'
    return 'M' 

# ==============================================================================
# VISTAS MVC (Plantillas Jinja2 in-memory)
# ==============================================================================
TEMPLATES = {
    'base.html': """
<!DOCTYPE html>
<html lang="es" class="h-full">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
    <title>SGC - Registro de Afectados</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet">
    <script src="https://unpkg.com/lucide@latest"></script>
    <style>
        body { font-family: 'Plus Jakarta Sans', sans-serif; background-color: #f8fafc; -webkit-tap-highlight-color: transparent; }
        .accordion-content { max-height: 0; overflow: hidden; transition: max-height 0.3s ease-out; }
        .med-tag { display: inline-block; font-size: 10px; font-weight: 700; padding: 3px 8px; border-radius: 6px; margin-right: 4px; margin-top: 4px; }
        .med-disc { background-color: #f0fdf4; color: #15803d; border: 1px solid #bbf7d0; text-transform: uppercase; }
        .med-pat { background-color: #fef3c7; color: #d97706; border: 1px solid #fcd34d; text-transform: uppercase; }
        .badge { padding: 5px 12px; border-radius: 6px; font-size: 10px; font-weight: 800; text-transform: uppercase; letter-spacing: 0.5px; white-space: nowrap; }
        .badge-danger { background-color: #fee2e2; color: #991b1b; border: 1px solid #f87171;} 
        .badge-warning { background-color: #ffedd5; color: #c2410c; border: 1px solid #fb923c;}
        .badge-success { background-color: #dcfce7; color: #166534; border: 1px solid #4ade80;}
        .badge-neutral { background-color: #f1f5f9; color: #475569; border: 1px solid #cbd5e1;}
        .hide-scrollbar::-webkit-scrollbar { display: none; }
        .hide-scrollbar { -ms-overflow-style: none; scrollbar-width: none; }
        @media print { nav, footer, .no-print { display: none !important; } body { background: white !important; } .accordion-content { max-height: none !important; display: block !important; } }
    </style>
</head>
<body class="text-slate-800 min-h-screen flex flex-col antialiased">
    {% if session.get('username') %}
    <nav class="bg-[#0f172a] text-white shadow-lg border-b border-slate-800 no-print sticky top-0 z-40">
        <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
            <div class="flex flex-col md:flex-row md:items-center justify-between py-3 md:h-16">
                <div class="flex items-center space-x-3 mb-3 md:mb-0">
                    <div class="bg-blue-600 p-2 rounded-lg shadow-inner"><i data-lucide="shield-alert" class="w-5 h-5"></i></div>
                    <div>
                        <span class="font-extrabold text-base md:text-lg block leading-none tracking-tight">SGC - AFECTADOS</span>
                        <span class="text-[9px] md:text-[10px] text-blue-300 font-bold uppercase tracking-wider block mt-1">Control Demográfico | ROL: {{ session.get('rol') }}</span>
                    </div>
                </div>
                <div class="flex space-x-1 items-center overflow-x-auto hide-scrollbar pb-1 md:pb-0 w-full md:w-auto">
                    {% set acceso = session.get('acceso_modulos', 'TOTAL') %}
                    
                    {% if acceso in ['TOTAL', 'INICIO'] %}
                    <a href="{{ url_for('index') }}" class="px-3 py-2 rounded-lg text-xs md:text-sm font-semibold hover:bg-slate-800 flex items-center gap-1.5 transition-colors whitespace-nowrap"><i data-lucide="home" class="w-4 h-4"></i> Inicio</a>
                    {% endif %}
                    
                    {% if session.get('rol') in ['ADMIN', 'REGISTRAR'] %}
                    <a href="{{ url_for('registrar') }}" class="px-3 py-2 rounded-lg text-xs md:text-sm font-semibold bg-blue-600 hover:bg-blue-700 flex items-center gap-1.5 transition-colors whitespace-nowrap"><i data-lucide="user-plus" class="w-4 h-4"></i> Registrar</a>
                    {% endif %}
                    
                    {% if session.get('rol') == 'ADMIN' %}
                    <a href="{{ url_for('config_tv') }}" class="px-3 py-2 rounded-lg text-xs md:text-sm font-semibold hover:bg-slate-800 flex items-center gap-1.5 text-indigo-400 transition-colors whitespace-nowrap"><i data-lucide="tv" class="w-4 h-4"></i> TV</a>
                    <a href="{{ url_for('cargar_datos') }}" class="px-3 py-2 rounded-lg text-xs md:text-sm font-semibold hover:bg-slate-800 flex items-center gap-1.5 text-emerald-400 transition-colors whitespace-nowrap"><i data-lucide="database" class="w-4 h-4"></i> BD</a>
                    <a href="{{ url_for('usuarios') }}" class="px-3 py-2 rounded-lg text-xs md:text-sm font-semibold hover:bg-slate-800 flex items-center gap-1.5 transition-colors whitespace-nowrap"><i data-lucide="users-round" class="w-4 h-4"></i> Usuarios</a>
                    <a href="{{ url_for('parroquias') }}" class="px-3 py-2 rounded-lg text-xs md:text-sm font-semibold hover:bg-slate-800 flex items-center gap-1.5 transition-colors whitespace-nowrap"><i data-lucide="map-pin" class="w-4 h-4"></i> Zonas</a>
                    <a href="{{ url_for('ver_historial') }}" class="px-3 py-2 rounded-lg text-xs md:text-sm font-semibold hover:bg-slate-800 flex items-center gap-1.5 text-amber-400 transition-colors whitespace-nowrap"><i data-lucide="history" class="w-4 h-4"></i> Logs</a>
                    {% endif %}
                    <div class="h-6 w-[1px] bg-slate-700 mx-2 flex-shrink-0"></div>
                    <a href="{{ url_for('logout') }}" class="bg-rose-500/10 text-rose-400 hover:bg-rose-500 hover:text-white px-3 py-2 rounded-lg text-xs font-bold transition-all flex items-center gap-1 whitespace-nowrap"><i data-lucide="log-out" class="w-3 h-3"></i> Salir</a>
                </div>
            </div>
        </div>
    </nav>
    {% endif %}

    <main class="flex-grow max-w-7xl w-full mx-auto px-2 md:px-6 lg:px-8 py-4 md:py-8">
        {% with messages = get_flashed_messages(with_categories=true) %}
          {% if messages %}
            <div class="mb-4 md:mb-6 space-y-2 no-print">
            {% for category, message in messages %}
              <div class="p-3 md:p-4 rounded-xl font-semibold text-xs md:text-sm flex items-center gap-3 {% if category == 'error' %}bg-rose-50 text-rose-800 border border-rose-200 shadow-sm{% else %}bg-emerald-50 text-emerald-800 border border-emerald-200 shadow-sm{% endif %}">
                <i data-lucide="{% if category == 'error' %}alert-circle{% else %}check-circle{% endif %}" class="w-5 h-5 flex-shrink-0"></i>
                {{ message }}
              </div>
            {% endfor %}
            </div>
          {% endif %}
        {% endwith %}

        {% block content %}{% endblock %}
    </main>
    <script>lucide.createIcons();</script>
</body>
</html>
    """,
    'login.html': """
{% extends 'base.html' %}
{% block content %}
<div class="max-w-md mx-auto mt-8 md:mt-16 bg-white rounded-2xl shadow-xl border border-slate-200 overflow-hidden">
    <div class="px-6 py-8 md:py-10 bg-[#0f172a] text-white text-center">
        <div class="bg-blue-600 text-white w-14 h-14 md:w-16 md:h-16 rounded-2xl mx-auto flex items-center justify-center mb-4 shadow-lg"><i data-lucide="shield-alert" class="w-7 h-7 md:w-8 md:h-8"></i></div>
        <h2 class="text-xl md:text-2xl font-extrabold tracking-tight">Registro de Afectados</h2>
        <p class="text-[10px] md:text-xs text-slate-400 mt-2 font-semibold uppercase tracking-wider">Módulo de Acceso Seguro</p>
    </div>
    <form method="POST" class="p-6 md:p-8 space-y-4 md:space-y-5">
        <input type="hidden" name="csrf_token" value="{{ csrf_token }}"/>
        <div><label class="block text-xs font-bold text-slate-500 uppercase tracking-wide mb-2">Usuario</label><input type="text" name="username" required class="w-full px-4 py-3 border border-slate-300 rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 bg-slate-50"></div>
        <div><label class="block text-xs font-bold text-slate-500 uppercase tracking-wide mb-2">Contraseña</label><input type="password" name="password" required class="w-full px-4 py-3 border border-slate-300 rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 bg-slate-50"></div>
        <button type="submit" class="w-full bg-blue-600 hover:bg-blue-700 text-white font-bold py-3.5 rounded-xl text-sm transition-all shadow-md">Ingresar al Sistema</button>
    </form>
</div>
{% endblock %}
    """,
    'inicio.html': """
{% extends 'base.html' %}
{% block content %}
<div class="flex flex-col lg:flex-row justify-between items-start lg:items-center mb-6 md:mb-8 gap-4 no-print">
    <div>
        <h1 class="text-xl md:text-2xl font-extrabold text-[#0f172a] tracking-tight">Consolidado de Inicio</h1>
        <p class="text-xs md:text-sm text-slate-500 font-medium mt-1">Panel Automatizado de Gestión Demográfica</p>
    </div>
    <div class="flex flex-wrap items-center gap-2 w-full lg:w-auto">
        <a href="{{ url_for('dashboard') }}" target="_blank" class="bg-indigo-600 hover:bg-indigo-700 text-white font-bold py-2.5 px-4 rounded-xl flex items-center gap-2 shadow-sm text-xs md:text-sm transition-all w-full sm:w-auto justify-center"><i data-lucide="monitor" class="w-4 h-4"></i> Pantalla Fija (TV)</a>
        <a href="{{ url_for('exportar_csv') }}" class="bg-slate-800 hover:bg-slate-900 text-white font-bold py-2.5 px-4 rounded-xl flex items-center gap-2 shadow-sm text-xs md:text-sm transition-all w-full sm:w-auto justify-center"><i data-lucide="download-cloud" class="w-4 h-4"></i> Exportar CSV</a>
        <button onclick="printStatsBox()" class="bg-blue-600 hover:bg-blue-700 text-white font-bold py-2.5 px-4 rounded-xl flex items-center gap-2 shadow-sm text-xs md:text-sm transition-all w-full sm:w-auto justify-center hidden md:flex"><i data-lucide="file-bar-chart" class="w-4 h-4"></i> Resumen</button>
    </div>
</div>

<div id="printStatsArea">
    <div class="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-6 gap-3 md:gap-4 mb-6 md:mb-8">
        <div class="bg-white p-3 md:p-4 rounded-2xl shadow-sm border border-slate-200 border-l-4 border-l-blue-800"><h3 class="text-[9px] md:text-[10px] font-bold text-slate-500 uppercase tracking-wider">Jefes Familia</h3><div class="text-2xl md:text-3xl font-black text-[#0f172a] mt-1">{{ stats.familias }}</div></div>
        <div class="bg-white p-3 md:p-4 rounded-2xl shadow-sm border border-slate-200 border-l-4 border-l-blue-500"><h3 class="text-[9px] md:text-[10px] font-bold text-slate-500 uppercase tracking-wider">Población Total</h3><div class="text-2xl md:text-3xl font-black text-blue-600 mt-1">{{ stats.poblacion }}</div></div>
        <div class="bg-white p-3 md:p-4 rounded-2xl shadow-sm border border-slate-200 border-l-4 border-l-rose-500"><h3 class="text-[9px] md:text-[10px] font-bold text-slate-500 uppercase tracking-wider">Casos Críticos</h3><div class="text-2xl md:text-3xl font-black text-rose-600 mt-1">{{ stats.criticos }}</div></div>
        <div class="bg-white p-3 md:p-4 rounded-2xl shadow-sm border border-slate-200 border-l-4 border-l-pink-500"><h3 class="text-[9px] md:text-[10px] font-bold text-slate-500 uppercase tracking-wider">Embarazadas</h3><div class="text-2xl md:text-3xl font-black text-pink-500 mt-1">{{ stats.embarazadas }}</div></div>
        <div class="bg-white p-3 md:p-4 rounded-2xl shadow-sm border border-slate-200 border-l-4 border-l-teal-500"><h3 class="text-[9px] md:text-[10px] font-bold text-slate-500 uppercase tracking-wider">Pers. Condición</h3><div class="text-2xl md:text-3xl font-black text-teal-600 mt-1">{{ stats.discapacitados }}</div></div>
        <div class="bg-white p-3 md:p-4 rounded-2xl shadow-sm border border-slate-200 border-l-4 border-l-purple-500"><h3 class="text-[9px] md:text-[10px] font-bold text-slate-500 uppercase tracking-wider">Remitidos a Obras</h3><div class="text-2xl md:text-3xl font-black text-purple-600 mt-1">{{ stats.remitidos }}</div></div>
    </div>

    <div class="bg-white rounded-2xl shadow-sm border border-slate-200 p-4 md:p-5 mb-6 md:mb-8">
        <h3 class="text-[11px] md:text-xs font-bold text-slate-800 uppercase tracking-wider mb-3 md:mb-4 flex items-center gap-2"><i data-lucide="pie-chart" class="w-4 h-4 text-purple-500"></i> Desglose Poblacional</h3>
        <div class="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-6 gap-3 md:gap-4">
            <div class="bg-slate-50 p-3 md:p-4 rounded-xl border border-slate-100 flex flex-col justify-between sm:col-span-2 md:col-span-2">
                <div>
                    <h4 class="text-[9px] md:text-[10px] font-bold text-slate-500 uppercase">Total Infantes (0-12)</h4>
                    <div class="text-xl md:text-2xl font-black text-slate-800 mt-1">{{ stats.ninos }}</div>
                </div>
                <div class="grid grid-cols-2 gap-2 mt-3 border-t border-slate-200 pt-2">
                    <div class="border-r border-slate-200 pr-2">
                        <div class="text-blue-600 font-bold text-[10px] md:text-xs mb-1 flex items-center gap-1"><i data-lucide="user" class="w-3 h-3"></i> Varones: {{ stats.ninos_m }}</div>
                        <div class="flex flex-col gap-0.5 text-[9px] text-slate-500 font-medium">
                            <span class="flex justify-between"><span>0-3 años:</span> <span class="font-bold">{{ stats.ninos_m_0_3 }}</span></span>
                            <span class="flex justify-between"><span>4-6 años:</span> <span class="font-bold">{{ stats.ninos_m_4_6 }}</span></span>
                            <span class="flex justify-between"><span>7-12 años:</span> <span class="font-bold">{{ stats.ninos_m_7_12 }}</span></span>
                        </div>
                    </div>
                    <div class="pl-1">
                        <div class="text-pink-600 font-bold text-[10px] md:text-xs mb-1 flex items-center gap-1"><i data-lucide="user" class="w-3 h-3"></i> Niñas: {{ stats.ninos_f }}</div>
                        <div class="flex flex-col gap-0.5 text-[9px] text-slate-500 font-medium">
                            <span class="flex justify-between"><span>0-3 años:</span> <span class="font-bold">{{ stats.ninos_f_0_3 }}</span></span>
                            <span class="flex justify-between"><span>4-6 años:</span> <span class="font-bold">{{ stats.ninos_f_4_6 }}</span></span>
                            <span class="flex justify-between"><span>7-12 años:</span> <span class="font-bold">{{ stats.ninos_f_7_12 }}</span></span>
                        </div>
                    </div>
                </div>
            </div>
            <div class="bg-slate-50 p-3 md:p-4 rounded-xl border border-slate-100 flex flex-col justify-between">
                <div>
                    <h4 class="text-[9px] md:text-[10px] font-bold text-slate-500 uppercase">Adolescentes (13-17)</h4>
                    <div class="text-xl md:text-2xl font-black text-slate-800 mt-1">{{ stats.adol }}</div>
                </div>
                <div class="text-[9px] md:text-[10px] text-slate-600 mt-2 border-t border-slate-200 pt-1.5">
                    <div class="flex justify-between mb-1">
                        <span class="text-blue-600 font-bold">{{ stats.adol_m }} Masc.</span>
                        <span class="text-pink-600 font-bold">{{ stats.adol_f }} Fem.</span>
                    </div>
                </div>
            </div>
            <div class="bg-slate-50 p-3 md:p-4 rounded-xl border border-slate-100 flex flex-col justify-between">
                <div>
                    <h4 class="text-[9px] md:text-[10px] font-bold text-slate-500 uppercase">Adultos (18-54)</h4>
                    <div class="text-xl md:text-2xl font-black text-slate-800 mt-1">{{ stats.adultos }}</div>
                </div>
                <div class="text-[9px] md:text-[10px] text-slate-600 mt-2 border-t border-slate-200 pt-1.5">
                    <div class="flex justify-between mb-1">
                        <span class="text-blue-600 font-bold">{{ stats.adultos_m }} Masc.</span>
                        <span class="text-pink-600 font-bold">{{ stats.adultos_f }} Fem.</span>
                    </div>
                </div>
            </div>
            <div class="bg-slate-50 p-3 md:p-4 rounded-xl border border-slate-100 flex flex-col justify-between"><h4 class="text-[9px] md:text-[10px] font-bold text-slate-500 uppercase">Muj. Mayores (≥55)</h4><div class="text-xl md:text-2xl font-black text-slate-800 mt-1">{{ stats.muj_may }}</div></div>
            <div class="bg-slate-50 p-3 md:p-4 rounded-xl border border-slate-100 flex flex-col justify-between"><h4 class="text-[9px] md:text-[10px] font-bold text-slate-500 uppercase">Hom. Mayores (≥60)</h4><div class="text-xl md:text-2xl font-black text-slate-800 mt-1">{{ stats.hom_may }}</div></div>
        </div>
    </div>
</div>

<!-- FILTRO INTELIGENTE INTEGRADO -->
<form method="GET" action="{{ url_for('index') }}" class="bg-white border border-slate-200 p-4 rounded-2xl mb-6 no-print shadow-sm">
    <div class="flex items-center gap-2 mb-3 border-b border-slate-100 pb-2">
        <i data-lucide="filter" class="w-5 h-5 text-blue-600"></i>
        <h3 class="font-extrabold text-xs md:text-sm text-slate-800 uppercase tracking-wider">Filtro Inteligente de Datos</h3>
    </div>
    <div class="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-5 gap-3">
        <div class="md:col-span-2 relative">
            <div class="absolute inset-y-0 left-0 pl-3 flex items-center pointer-events-none"><i data-lucide="search" class="w-4 h-4 text-slate-400"></i></div>
            <input type="text" name="q" value="{{ request.args.get('q', '') }}" placeholder="Buscar por Nombre o Cédula..." class="w-full pl-10 pr-3 py-2.5 border border-slate-300 rounded-xl text-xs md:text-sm focus:ring-2 focus:ring-blue-500 outline-none transition-shadow bg-slate-50">
        </div>
        <div>
            <select name="grupo" class="w-full px-3 py-2.5 border border-slate-300 rounded-xl text-xs md:text-sm focus:ring-2 focus:ring-blue-500 outline-none font-bold text-slate-600 bg-slate-50 cursor-pointer">
                <option value="">Todos los Grupos Pob.</option>
                <option value="ninos" {% if request.args.get('grupo') == 'ninos' %}selected{% endif %}>Niños (0-12)</option>
                <option value="ninas" {% if request.args.get('grupo') == 'ninas' %}selected{% endif %}>Niñas (0-12)</option>
                <option value="adol" {% if request.args.get('grupo') == 'adol' %}selected{% endif %}>Adolescentes (13-17)</option>
                <option value="mayores" {% if request.args.get('grupo') == 'mayores' %}selected{% endif %}>Adultos Mayores</option>
                <option value="embarazadas" {% if request.args.get('grupo') == 'embarazadas' %}selected{% endif %}>Embarazadas</option>
                <option value="discapacidad" {% if request.args.get('grupo') == 'discapacidad' %}selected{% endif %}>Con Condición/Discap.</option>
            </select>
        </div>
        <div>
            <select name="parroquia" class="w-full px-3 py-2.5 border border-slate-300 rounded-xl text-xs md:text-sm focus:ring-2 focus:ring-blue-500 outline-none font-bold text-slate-600 bg-slate-50 cursor-pointer">
                <option value="">Todas las Zonas</option>
                {% for p in parroquias %}
                <option value="{{ p.nombre }}" {% if request.args.get('parroquia') == p.nombre %}selected{% endif %}>{{ p.nombre }}</option>
                {% endfor %}
            </select>
        </div>
        <div class="flex gap-2">
            <button type="submit" class="flex-1 bg-blue-600 hover:bg-blue-700 text-white font-bold py-2.5 rounded-xl shadow-md text-xs md:text-sm transition-colors flex items-center justify-center gap-1.5"><i data-lucide="check-circle-2" class="w-4 h-4"></i> Aplicar</button>
            <a href="{{ url_for('index') }}" class="flex-1 bg-white hover:bg-slate-100 text-slate-600 font-bold py-2.5 border border-slate-300 rounded-xl text-xs md:text-sm transition-colors text-center flex items-center justify-center gap-1.5" title="Limpiar"><i data-lucide="rotate-ccw" class="w-4 h-4"></i> Limpiar</a>
        </div>
    </div>
</form>

<!-- PAGINACIÓN TOP -->
{% if pagination and pagination.pages > 1 %}
<div class="flex justify-between items-center mt-2 mb-4 no-print gap-2 text-xs">
    <span class="text-slate-500 font-bold">Página {{ pagination.page }} de {{ pagination.pages }}</span>
    <div class="flex gap-1">
        {% if pagination.has_prev %}
            <a href="{{ url_for('index', page=pagination.prev_num, q=request.args.get('q', ''), parroquia=request.args.get('parroquia', ''), grupo=request.args.get('grupo', '')) }}" class="px-3 py-1.5 bg-white border border-slate-300 rounded-lg text-slate-600 hover:bg-slate-50 font-bold">&laquo; Ant</a>
        {% endif %}
        {% if pagination.has_next %}
            <a href="{{ url_for('index', page=pagination.next_num, q=request.args.get('q', ''), parroquia=request.args.get('parroquia', ''), grupo=request.args.get('grupo', '')) }}" class="px-3 py-1.5 bg-white border border-slate-300 rounded-lg text-slate-600 hover:bg-slate-50 font-bold">Sig &raquo;</a>
        {% endif %}
    </div>
</div>
{% endif %}

<div class="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden mb-6 no-print">
    {% for parroquia, familias in agrupacion.items() %}
    <div class="border-b border-slate-200 last:border-0 accordion-container">
        <button onclick="toggleAccordion('{{ loop.index }}')" class="w-full px-4 py-3 md:px-6 md:py-4 flex items-center justify-between bg-slate-50/80 hover:bg-slate-100 transition-colors focus:outline-none">
            <div class="text-left"><span class="font-extrabold text-[#0f172a] text-xs md:text-sm flex items-center gap-2"><i data-lucide="map-pin" class="w-4 h-4 text-blue-600 flex-shrink-0"></i> {{ parroquia }}</span></div>
            <div class="flex items-center gap-2 md:gap-3"><span class="bg-white border border-slate-200 text-slate-700 text-[9px] md:text-[10px] font-bold px-2 py-1 md:px-3 rounded-full shadow-sm">{{ familias|length }} Núcleos en esta pág.</span><i id="accordion-icon-{{ loop.index }}" data-lucide="chevron-down" class="w-4 h-4 text-slate-400 transition-transform no-print transform"></i></div>
        </button>
        <div id="accordion-content-{{ loop.index }}" class="accordion-content bg-white">
            <div class="overflow-x-auto p-0">
                <table class="w-full text-left border-collapse min-w-[800px] md:min-w-[900px]">
                    <thead>
                        <tr class="bg-white border-b border-slate-200 text-[9px] md:text-[10px] uppercase tracking-widest text-slate-400">
                            <th class="p-3 md:p-4 font-bold w-10 md:w-12 text-center">N°</th>
                            <th class="p-3 md:p-4 font-bold w-1/4">Jefe de Familia</th>
                            <th class="p-3 md:p-4 font-bold">C.I. / Datos</th>
                            <th class="p-3 md:p-4 font-bold w-48">Situación / Riesgo</th>
                            <th class="p-3 md:p-4 font-bold w-2/5">Núcleo Familiar</th>
                            <th class="p-3 md:p-4 font-bold text-center no-print">Acción</th>
                        </tr>
                    </thead>
                    <tbody class="divide-y divide-slate-100 text-xs md:text-sm">
                        {% for jefe in familias %}
                        <tr class="hover:bg-slate-50/50 data-row group transition-colors">
                            <td class="p-3 md:p-4 text-center font-bold text-slate-400 text-[10px] md:text-xs">{{ loop.index }}</td>
                            <td class="p-3 md:p-4 align-top">
                                <div class="flex items-start gap-2 md:gap-3">
                                    {% if jefe.foto %}
                                    <img src="{{ url_for('static', filename='uploads/' + jefe.foto) }}" onclick="verFotoModal('{{ url_for('static', filename='uploads/' + jefe.foto) }}')" class="w-8 h-8 md:w-10 md:h-10 rounded-full object-cover border border-slate-200 shadow-sm mt-1 cursor-pointer hover:opacity-80 transition-opacity flex-shrink-0">
                                    {% else %}
                                    <div class="w-8 h-8 md:w-10 md:h-10 rounded-full bg-slate-100 flex items-center justify-center text-slate-400 border border-slate-200 mt-1 flex-shrink-0"><i data-lucide="user" class="w-4 h-4 md:w-5 md:h-5"></i></div>
                                    {% endif %}
                                    <div>
                                        <div class="font-bold text-slate-800 uppercase text-[11px] md:text-[13px] leading-tight">{{ jefe.nombre }}</div>
                                        <div class="mt-1 md:mt-1.5 flex flex-wrap gap-1">
                                            {% if jefe.es_embarazada %}<span class="med-tag bg-pink-100 text-pink-700 border border-pink-200">🤰 EMB</span>{% endif %}
                                            {% if jefe.discapacidad and jefe.discapacidad|upper not in ['NINGUNA', 'NINGUNO', 'NO', ''] %}<span class="med-tag med-disc" title="Persona con Condición"><i data-lucide="accessibility" class="w-3 h-3 inline"></i> {{ jefe.discapacidad }}</span>{% endif %}
                                            {% if jefe.patologia and jefe.patologia|upper not in ['NINGUNA', 'NINGUNO', 'NO', ''] %}<span class="med-tag med-pat">🩺 {{ jefe.patologia }}</span>{% endif %}
                                        </div>
                                    </div>
                                </div>
                            </td>
                            <td class="p-3 md:p-4 align-top font-medium text-slate-600">
                                <div class="font-mono text-[10px] md:text-xs font-bold text-blue-900">{{ jefe.ci or 'S/C' }}</div>
                                <div class="text-[9px] md:text-[11px] mt-1 text-slate-500 font-semibold">{% if jefe.edad == -1 %}N/D{% else %}{{ jefe.edad }}{% endif %} años &bull; {{ jefe.genero }}</div>
                                <div class="text-[9px] md:text-[11px] mt-0.5 text-slate-500 flex items-center gap-1"><i data-lucide="phone" class="w-3 h-3"></i> {{ jefe.telefono or 'S/T' }}</div>
                            </td>
                            
                            <td class="p-3 md:p-4 align-top">
                                {% set sit = (jefe.situacion or '').upper() %}
                                <div class="mb-2 md:mb-3">
                                    {% if 'INHABITABLE' in sit or 'SIN CASA' in sit or 'ALTO RIESGO' in sit or 'DERRUMBE' in sit %}
                                        <span class="badge badge-danger">{{ jefe.situacion }}</span>
                                    {% elif 'RIESGO MODERADO' in sit or 'REFUGIO' in sit or 'CANCHA' in sit %}
                                        <span class="badge badge-warning">{{ jefe.situacion }}</span>
                                    {% elif 'ESTABLE' in sit %}
                                        <span class="badge badge-success">{{ jefe.situacion }}</span>
                                    {% else %}
                                        <span class="badge badge-neutral">{{ jefe.situacion }}</span>
                                    {% endif %}
                                </div>
                                
                                {% if jefe.foto_vivienda %}
                                <div class="mb-2">
                                    <button type="button" onclick="verFotoModal('{{ url_for('static', filename='uploads/' + jefe.foto_vivienda) }}')" class="w-full text-[9px] font-bold text-rose-700 bg-rose-50 border border-rose-200 hover:bg-rose-100 px-2 py-1.5 rounded shadow-sm transition-colors text-center uppercase flex justify-center items-center gap-1.5">
                                        <i data-lucide="home" class="w-3 h-3"></i> VER INSPECCIÓN
                                    </button>
                                </div>
                                {% endif %}

                                {% if jefe.requerimiento and jefe.requerimiento|upper != 'NINGUNO' and jefe.requerimiento != '' %}
                                <div class="mb-2 p-1.5 md:p-2 bg-blue-50 border border-blue-100 rounded-lg text-[9px] md:text-[10px] text-blue-800 shadow-sm leading-tight uppercase font-bold">
                                    <strong class="block text-[8px] text-blue-600 mb-0.5"><i data-lucide="package" class="inline w-2.5 h-2.5"></i> REQ:</strong>
                                    {{ jefe.requerimiento }}
                                </div>
                                {% endif %}

                                {% if jefe.lugar_remision and jefe.lugar_remision != '' %}
                                <div class="mb-2 p-1.5 md:p-2 bg-purple-50 border border-purple-200 rounded-lg text-[9px] md:text-[10px] text-purple-800 shadow-sm leading-tight uppercase font-bold">
                                    <strong class="block text-[8px] text-purple-600 mb-0.5"><i data-lucide="send" class="inline w-2.5 h-2.5"></i> REMITIDO A:</strong>
                                    {{ jefe.lugar_remision }}
                                </div>
                                {% endif %}

                                {% if jefe.observacion and jefe.observacion|upper != 'NINGUNA' and jefe.observacion|upper != 'NINGUNO' and jefe.observacion != '' %}
                                <div class="mb-2 p-1.5 md:p-2 bg-amber-50 border border-amber-100 rounded-lg text-[9px] md:text-[10px] text-amber-800 shadow-sm leading-tight uppercase font-bold">
                                    <strong class="block text-[8px] text-amber-600 mb-0.5"><i data-lucide="clipboard-list" class="inline w-2.5 h-2.5"></i> OBS:</strong>
                                    {{ jefe.observacion }}
                                </div>
                                {% endif %}

                                {% if session.get('rol') in ['ADMIN', 'REGISTRAR'] %}
                                <form method="POST" action="{{ url_for('actualizar_situacion', id=jefe.id) }}" class="no-print mt-2">
                                    <input type="hidden" name="csrf_token" value="{{ csrf_token }}"/>
                                    <select name="nueva_situacion" onchange="this.form.submit()" class="text-[9px] md:text-[10px] w-full p-1.5 border border-slate-300 rounded-lg bg-white font-bold text-slate-600 cursor-pointer focus:ring-2 focus:ring-blue-500 outline-none shadow-sm transition-shadow">
                                        <option disabled selected>Estado</option>
                                        <option value="Estable">Estable</option>
                                        <option value="Refugio / Cancha">Refugio / Cancha</option>
                                        <option value="Riesgo Moderado">Riesgo Moderado</option>
                                        <option value="Alto Riesgo">Alto Riesgo</option>
                                        <option value="Sin Casa">Sin Casa</option>
                                        <option value="Derrumbe">Derrumbe</option>
                                        <option value="Inhabitable">Inhabitable</option>
                                        <option value="Evaluación Médica">Evaluación Médica</option>
                                    </select>
                                </form>
                                {% endif %}
                            </td>
                            
                            <td class="p-3 md:p-4 align-top">
                                {% if jefe.cargas %}
                                <div class="bg-slate-50 border border-slate-200 p-2 md:p-3 rounded-xl text-[10px] md:text-xs max-h-40 overflow-y-auto">
                                    <div class="font-extrabold text-slate-400 mb-1.5 uppercase tracking-wider text-[8px] md:text-[9px] flex items-center gap-1"><i data-lucide="users" class="w-3 h-3"></i> Cargas ({{ jefe.cargas|length }})</div>
                                    <div class="space-y-1.5 md:space-y-2">
                                        {% for m in jefe.cargas %}
                                        <div class="border-b border-slate-200/50 pb-1.5 last:border-0 last:pb-0">
                                            <div class="flex items-start md:items-center flex-col md:flex-row gap-0 md:gap-1 leading-tight">
                                                <span class="font-bold text-slate-700 uppercase text-[10px] md:text-[11px]">{{ m.nombre }}</span> 
                                                <span class="text-slate-500 text-[9px] md:text-[10px] uppercase">({{ m.parentesco }}, {% if m.edad == -1 %}N/D{% else %}{{ m.edad }}a{% endif %})</span>
                                                <!-- BOTON CARNET PARA CARGAS -->
                                                <button type="button" onclick="pedirNumeroFamiliaCarnet('carga', {{ m.id }})" class="ml-1 text-blue-500 hover:text-blue-700 transition-colors" title="Imprimir Carnet Trabajador">
                                                    <i data-lucide="id-card" class="w-3 h-3"></i>
                                                </button>
                                            </div>
                                            <div class="font-mono text-[9px] text-slate-400 font-bold mb-0.5 mt-0.5">C.I. {{ m.ci or 'S/C' }}</div>
                                            <div class="mt-0.5 md:mt-1 flex flex-wrap gap-1">
                                                {% if m.es_embarazada %}<span class="med-tag bg-pink-100 text-pink-700 border-pink-200">🤰 EMB</span>{% endif %}
                                                {% if m.discapacidad and m.discapacidad|upper not in ['NINGUNA', 'NINGUNO', 'NO', ''] %}<span class="med-tag med-disc" title="Persona con Condición"><i data-lucide="accessibility" class="w-3 h-3 inline"></i> {{ m.discapacidad }}</span>{% endif %}
                                                {% if m.patologia and m.patologia|upper not in ['NINGUNA', 'NINGUNO', 'NO', ''] %}<span class="med-tag med-pat">{{ m.patologia }}</span>{% endif %}
                                            </div>
                                        </div>
                                        {% endfor %}
                                    </div>
                                </div>
                                {% else %}<div class="text-[9px] md:text-[11px] italic text-slate-400 px-2 py-1 bg-slate-50 rounded border border-slate-100 inline-block">Sin cargas registradas</div>{% endif %}
                            </td>
                            <td class="p-3 md:p-4 text-center no-print align-middle">
                                <div class="flex items-center justify-center gap-1.5 md:gap-2 flex-wrap w-16 md:w-20">
                                    <button type="button" onclick="pedirNumeroFamilia({{ jefe.id }})" class="text-slate-500 hover:text-slate-700 p-1.5 md:p-2 rounded-lg hover:bg-slate-50 border border-transparent hover:border-slate-200 transition-all shadow-sm bg-white" title="Imprimir Ficha Completa">
                                        <i data-lucide="printer" class="w-3.5 h-3.5 md:w-4 md:h-4"></i>
                                    </button>
                                    <!-- BOTON CARNET PARA JEFE -->
                                    <button type="button" onclick="pedirNumeroFamiliaCarnet('jefe', {{ jefe.id }})" class="text-indigo-500 hover:text-indigo-700 p-1.5 md:p-2 rounded-lg hover:bg-indigo-50 border border-transparent hover:border-indigo-200 transition-all shadow-sm bg-white" title="Imprimir Carnet Pase Laboral">
                                        <i data-lucide="id-card" class="w-3.5 h-3.5 md:w-4 md:h-4"></i>
                                    </button>

                                    {% if session.get('rol') == 'ADMIN' %}
                                    <a href="{{ url_for('editar_jefe', id=jefe.id) }}" class="text-blue-500 hover:text-blue-700 p-1.5 md:p-2 rounded-lg hover:bg-blue-50 border border-transparent hover:border-blue-100 transition-all shadow-sm bg-white" title="Editar">
                                        <i data-lucide="pencil" class="w-3.5 h-3.5 md:w-4 md:h-4"></i>
                                    </a>
                                    <form method="POST" action="{{ url_for('eliminar_jefe', id=jefe.id) }}" onsubmit="return confirm('¿Eliminar familia completa? Esta acción se auditará.');" class="inline-block">
                                        <input type="hidden" name="csrf_token" value="{{ csrf_token }}"/>
                                        <button type="submit" class="text-rose-500 hover:text-rose-700 p-1.5 md:p-2 rounded-lg hover:bg-rose-50 border border-transparent hover:border-rose-100 transition-all shadow-sm bg-white" title="Eliminar">
                                            <i data-lucide="trash-2" class="w-3.5 h-3.5 md:w-4 md:h-4"></i>
                                        </button>
                                    </form>
                                    {% endif %}
                                </div>
                            </td>
                        </tr>
                        {% endfor %}
                    </tbody>
                </table>
            </div>
        </div>
    </div>
    {% endfor %}
    {% if not agrupacion %}
    <div class="p-12 md:p-16 text-center flex flex-col items-center text-slate-400 bg-white rounded-2xl border border-slate-200">
        <div class="w-12 h-12 md:w-16 md:h-16 bg-slate-50 rounded-full flex items-center justify-center mb-3 md:mb-4"><i data-lucide="folder-open" class="w-6 h-6 md:w-8 md:h-8 text-slate-300"></i></div>
        <span class="font-bold text-slate-500 text-sm">No hay registros almacenados.</span>
        <p class="text-[10px] md:text-xs mt-2">Ve a la sección de BD para cargar tu Excel.</p>
    </div>
    {% endif %}
</div>

<!-- PAGINACIÓN BOTTOM -->
{% if pagination and pagination.pages > 1 %}
<div class="flex justify-center mt-6 mb-12 no-print gap-1 md:gap-2">
    {% if pagination.has_prev %}
        <a href="{{ url_for('index', page=pagination.prev_num, q=request.args.get('q', ''), parroquia=request.args.get('parroquia', ''), grupo=request.args.get('grupo', '')) }}" class="px-3 py-2 bg-white border border-slate-300 rounded-lg text-slate-600 hover:bg-slate-50 text-xs md:text-sm font-bold">&laquo;</a>
    {% endif %}

    {% for page_num in pagination.iter_pages(left_edge=1, right_edge=1, left_current=2, right_current=2) %}
        {% if page_num %}
            {% if pagination.page == page_num %}
                <span class="px-3.5 py-2 bg-blue-600 text-white rounded-lg font-bold shadow-sm text-xs md:text-sm">{{ page_num }}</span>
            {% else %}
                <a href="{{ url_for('index', page=page_num, q=request.args.get('q', ''), parroquia=request.args.get('parroquia', ''), grupo=request.args.get('grupo', '')) }}" class="px-3.5 py-2 bg-white border border-slate-300 rounded-lg text-slate-600 hover:bg-slate-50 text-xs md:text-sm font-semibold">{{ page_num }}</a>
            {% endif %}
        {% else %}
            <span class="px-3 py-2 text-slate-400">...</span>
        {% endif %}
    {% endfor %}

    {% if pagination.has_next %}
        <a href="{{ url_for('index', page=pagination.next_num, q=request.args.get('q', ''), parroquia=request.args.get('parroquia', ''), grupo=request.args.get('grupo', '')) }}" class="px-3 py-2 bg-white border border-slate-300 rounded-lg text-slate-600 hover:bg-slate-50 text-xs md:text-sm font-bold">&raquo;</a>
    {% endif %}
</div>
{% endif %}

<!-- Modal para ver fotografías ampliada -->
<div id="fotoModal" class="fixed inset-0 bg-slate-900/80 backdrop-blur-sm z-50 hidden flex items-center justify-center p-4 transition-opacity" onclick="cerrarFotoModal()">
    <div class="bg-white p-2 rounded-xl shadow-2xl max-w-2xl w-full" onclick="event.stopPropagation()">
        <div class="flex justify-between items-center mb-3 px-3 pt-2">
            <h3 class="font-bold text-slate-800 flex items-center gap-2"><i data-lucide="camera" class="w-5 h-5 text-blue-600"></i> Soporte Fotográfico</h3>
            <button onclick="cerrarFotoModal()" class="text-slate-400 hover:text-rose-500 transition-colors"><i data-lucide="x" class="w-6 h-6"></i></button>
        </div>
        <div class="bg-slate-100 rounded-lg p-2 flex justify-center items-center overflow-hidden">
            <img id="modalImgSrc" src="" alt="Cargando imagen..." class="w-full h-auto max-h-[70vh] object-contain rounded-md shadow-inner border border-slate-200">
        </div>
    </div>
</div>

<script>
    function toggleAccordion(id) {
        const content = document.getElementById('accordion-content-' + id);
        const icon = document.getElementById('accordion-icon-' + id);
        if (content.style.maxHeight && content.style.maxHeight !== '0px') {
            content.style.maxHeight = '0px'; icon.style.transform = 'rotate(0deg)';
        } else {
            content.style.maxHeight = content.scrollHeight + 'px'; icon.style.transform = 'rotate(180deg)';
        }
    }
    
    document.addEventListener('DOMContentLoaded', function() {
        if(document.getElementById('accordion-content-1')){
            toggleAccordion('1');
        }
    });

    function verFotoModal(url) {
        document.getElementById('modalImgSrc').src = url;
        document.getElementById('fotoModal').classList.remove('hidden');
    }
    function cerrarFotoModal() {
        document.getElementById('fotoModal').classList.add('hidden');
        document.getElementById('modalImgSrc').src = '';
    }

    function pedirNumeroFamilia(id) {
        let nro = prompt("Ingrese el N° de Familia para imprimir en la ficha (Dejar en blanco para omitir):");
        let url = "/imprimir_ficha/" + id;
        if (nro !== null && nro.trim() !== "") {
            url += "?nro=" + encodeURIComponent(nro.trim());
        }
        window.open(url, '_blank');
    }
    
    function pedirNumeroFamiliaCarnet(tipo, id) {
        let nro = prompt("Ingrese el N° de Familia para el Carnet del trabajador (Dejar en blanco para usar el ID de sistema):");
        let url = "/imprimir_carnet/" + tipo + "/" + id;
        if (nro !== null && nro.trim() !== "") {
            url += "?nro=" + encodeURIComponent(nro.trim());
        }
        window.open(url, '_blank');
    }

    function printStatsBox() {
        const box = document.getElementById('printStatsArea').cloneNode(true);
        const win = window.open('', '_blank');
        win.document.write(`
            <html>
                <head>
                    <title>Resumen Estadístico y Demográfico</title>
                    <script src="https://cdn.tailwindcss.com"><\\/script>
                    <script src="https://unpkg.com/lucide@latest"><\\/script>
                    <style>
                        body { background-color: #f8fafc; padding: 2rem; }
                        .no-print { display: none !important; }
                    </style>
                </head>
                <body>
                    <div class="max-w-5xl mx-auto bg-white border border-slate-200 p-8 rounded-3xl shadow-xl">
                        <div class="text-center border-b border-slate-200 pb-6 mb-6">
                            <h1 class="text-2xl font-black text-[#0f172a] uppercase tracking-tight">Reporte Consolidado de Totales</h1>
                            <p class="text-xs text-slate-500 mt-1 font-bold">Censo de Control Social y Distribución Poblacional</p>
                        </div>
                        ${box.outerHTML}
                    </div>
                    <script>
                        lucide.createIcons();
                        window.onload = function() {
                            setTimeout(function() {
                                window.print();
                                setTimeout(function() { window.close(); }, 500);
                            }, 800);
                        }
                    <\\/script>
                </body>
            </html>
        `);
        win.document.close();
    }
</script>
{% endblock %}
    """,
    'config_tv.html': """
{% extends 'base.html' %}
{% block content %}
<div class="max-w-2xl mx-auto bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
    <div class="px-4 py-4 md:px-6 md:py-5 bg-indigo-50 border-b border-indigo-100 flex items-center gap-2">
        <i data-lucide="tv" class="text-indigo-600 w-5 h-5"></i>
        <h2 class="text-base md:text-lg font-extrabold text-indigo-900">Ajustes Manuales del Dashboard TV</h2>
    </div>
    <div class="p-4 md:p-8">
        <div class="mb-6 bg-slate-50 p-4 rounded-xl border border-slate-200 text-xs text-slate-600 shadow-sm">
            <p class="flex items-start gap-2">
                <i data-lucide="info" class="w-4 h-4 text-blue-500 mt-0.5 flex-shrink-0"></i> 
                <span><strong>Nota:</strong> Solo aplica para Familias Egresadas. Si dejas un campo en blanco, el Dashboard TV calculará el valor automáticamente basándose en los registros reales de la Base de Datos. Si escribes un número, ese número forzará su aparición en la pantalla del TV <strong>sin afectar las estadísticas, exportaciones o reportes reales</strong> del sistema.</span>
            </p>
        </div>
        <form method="POST" class="space-y-5">
            <input type="hidden" name="csrf_token" value="{{ csrf_token }}"/>

            <div class="grid grid-cols-1 md:grid-cols-2 gap-5">
                <div>
                    <label class="block text-[11px] md:text-xs font-bold text-slate-600 uppercase tracking-wide mb-1.5">Familias Egresadas</label>
                    <input type="number" name="tv_familias_egresadas" value="{{ v_f_egr }}" placeholder="Auto (Dato: {{ stats.remitidos }})" class="w-full px-4 py-3 border border-slate-300 rounded-xl bg-white focus:ring-2 focus:ring-indigo-500 text-sm outline-none transition-shadow">
                </div>

                <div>
                    <label class="block text-[11px] md:text-xs font-bold text-slate-600 uppercase tracking-wide mb-1.5">Población Egresada</label>
                    <input type="number" name="tv_poblacion_egresada" value="{{ v_p_egr }}" placeholder="Auto (Dato: {{ stats.poblacion_remitida }})" class="w-full px-4 py-3 border border-slate-300 rounded-xl bg-white focus:ring-2 focus:ring-indigo-500 text-sm outline-none transition-shadow">
                </div>
            </div>

            <div class="pt-6 border-t border-slate-100 flex flex-col sm:flex-row justify-end gap-3">
                <a href="{{ url_for('index') }}" class="px-5 py-3 rounded-xl border text-slate-600 font-bold text-sm hover:bg-slate-50 text-center transition-colors">Volver</a>
                <button type="submit" class="px-5 py-3 rounded-xl bg-indigo-600 text-white font-bold text-sm shadow-md hover:bg-indigo-700 flex justify-center items-center gap-2 transition-colors"><i data-lucide="save" class="w-4 h-4"></i> Guardar Ajustes TV</button>
            </div>
        </form>
    </div>
</div>
{% endblock %}
    """,
    'dashboard.html': """
<!DOCTYPE html>
<html lang="es" class="h-full">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta http-equiv="refresh" content="300">
    <title>Dashboard SGC - Pantalla Fija</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800;900&display=swap" rel="stylesheet">
    <script src="https://unpkg.com/lucide@latest"></script>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.2.0"></script>
    <style>
        body { font-family: 'Plus Jakarta Sans', sans-serif; background-color: #020617; color: #f8fafc; }
        .glow { box-shadow: 0 0 20px rgba(59, 130, 246, 0.3); }
        .glow-red { box-shadow: 0 0 20px rgba(239, 68, 68, 0.2); }
    </style>
</head>
<body class="min-h-screen flex flex-col antialiased p-6 lg:p-12">
    <!-- Header Dashboard -->
    <div class="flex items-center justify-between border-b border-slate-800 pb-6 mb-8">
        <div class="flex items-center gap-4">
            <div class="bg-blue-600 p-4 rounded-2xl shadow-lg glow"><i data-lucide="shield-alert" class="w-10 h-10 text-white"></i></div>
            <div>
                <h1 class="text-4xl lg:text-5xl font-black tracking-tight text-white">Guayana Esequiba</h1>
                <p class="text-sm lg:text-base text-slate-400 font-bold uppercase tracking-widest mt-1">Fundación Niño Simón</p>
            </div>
        </div>
        <div class="text-right">
            <div class="text-3xl font-black text-slate-200" id="clock">--:--:--</div>
            <div class="text-xs text-blue-400 font-bold uppercase tracking-wider mt-1">Última Actualización: Automática</div>
        </div>
    </div>

    <!-- Indicadores Principales Gigantes -->
    <div class="grid grid-cols-1 md:grid-cols-3 gap-6 mb-8">
        <div class="bg-slate-900 border border-slate-800 p-6 rounded-3xl glow flex flex-col justify-center items-center text-center">
            <h3 class="text-sm lg:text-base text-slate-400 font-bold uppercase tracking-widest mb-2 flex items-center gap-2"><i data-lucide="users" class="w-5 h-5 text-blue-500"></i> Familias en el Campamento</h3>
            <div class="text-6xl lg:text-7xl font-black text-white">{{ stats.familias }}</div>
        </div>
        <div class="bg-blue-900/30 border border-blue-800/50 p-6 rounded-3xl glow flex flex-col justify-center items-center text-center">
            <h3 class="text-sm lg:text-base text-blue-300 font-bold uppercase tracking-widest mb-2 flex items-center gap-2"><i data-lucide="globe-2" class="w-5 h-5"></i> Población Total General</h3>
            <div class="text-6xl lg:text-7xl font-black text-blue-400">{{ stats.poblacion }}</div>
        </div>
        <div class="bg-rose-900/20 border border-rose-800/50 p-6 rounded-3xl glow-red flex flex-col justify-center items-center text-center">
            <h3 class="text-sm lg:text-base text-rose-300 font-bold uppercase tracking-widest mb-2 flex items-center gap-2"><i data-lucide="alert-triangle" class="w-5 h-5"></i> Casos Críticos / Refugio</h3>
            <div class="text-6xl lg:text-7xl font-black text-rose-500">{{ stats.criticos }}</div>
        </div>
    </div>

    <!-- DESGLOSE POBLACIONAL DETALLADO (ESTILO TV) -->
    <div class="bg-slate-900 border border-slate-800 p-6 lg:p-8 rounded-3xl mb-8 shadow-xl">
        <div class="flex items-center justify-between mb-6">
            <h3 class="text-sm lg:text-base font-black text-slate-400 uppercase tracking-widest flex items-center gap-2">
                <i data-lucide="pie-chart" class="w-5 h-5 text-purple-400"></i> Desglose Poblacional Detallado (Datos Reales)
            </h3>
        </div>
        
        <div class="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-6 gap-4 lg:gap-6">
            <!-- Box Infantes -->
            <div class="bg-slate-800/50 p-5 rounded-2xl border border-slate-700/50 flex flex-col justify-between sm:col-span-2">
                <div class="flex items-start justify-between">
                    <div>
                        <h4 class="text-[11px] lg:text-xs font-bold text-slate-400 uppercase tracking-widest">Total Infantes (0-12)</h4>
                        <div class="text-3xl lg:text-4xl font-black text-white mt-1">{{ stats.ninos }}</div>
                    </div>
                    <i data-lucide="baby" class="w-6 h-6 text-emerald-400"></i>
                </div>
                <div class="grid grid-cols-2 gap-3 mt-4 border-t border-slate-700/50 pt-4">
                    <!-- Varones -->
                    <div class="border-r border-slate-700/50 pr-3">
                        <div class="text-blue-400 font-bold text-[10px] lg:text-xs mb-2 flex items-center gap-1"><i data-lucide="user" class="w-3 h-3"></i> Varones: {{ stats.ninos_m }}</div>
                        <div class="flex flex-col gap-1.5 text-[10px] lg:text-[11px] text-slate-400 font-medium">
                            <span class="flex justify-between"><span>0-3 años:</span> <span class="font-bold text-slate-200">{{ stats.ninos_m_0_3 }}</span></span>
                            <span class="flex justify-between"><span>4-6 años:</span> <span class="font-bold text-slate-200">{{ stats.ninos_m_4_6 }}</span></span>
                            <span class="flex justify-between"><span>7-12 años:</span> <span class="font-bold text-slate-200">{{ stats.ninos_m_7_12 }}</span></span>
                        </div>
                    </div>
                    <!-- Niñas -->
                    <div class="pl-2">
                        <div class="text-pink-400 font-bold text-[10px] lg:text-xs mb-2 flex items-center gap-1"><i data-lucide="user" class="w-3 h-3"></i> Niñas: {{ stats.ninos_f }}</div>
                        <div class="flex flex-col gap-1.5 text-[10px] lg:text-[11px] text-slate-400 font-medium">
                            <span class="flex justify-between"><span>0-3 años:</span> <span class="font-bold text-slate-200">{{ stats.ninos_f_0_3 }}</span></span>
                            <span class="flex justify-between"><span>4-6 años:</span> <span class="font-bold text-slate-200">{{ stats.ninos_f_4_6 }}</span></span>
                            <span class="flex justify-between"><span>7-12 años:</span> <span class="font-bold text-slate-200">{{ stats.ninos_f_7_12 }}</span></span>
                        </div>
                    </div>
                </div>
            </div>
            
            <!-- Box Adolescentes -->
            <div class="bg-slate-800/50 p-5 rounded-2xl border border-slate-700/50 flex flex-col justify-between">
                <div>
                    <h4 class="text-[11px] lg:text-xs font-bold text-slate-400 uppercase tracking-widest text-center">Adolesc. (13-17)</h4>
                    <div class="text-3xl lg:text-4xl font-black text-white mt-2 text-center">{{ stats.adol }}</div>
                </div>
                <div class="text-[10px] lg:text-xs text-slate-400 mt-4 border-t border-slate-700/50 pt-3">
                    <div class="flex justify-between mb-1">
                        <span class="text-blue-400 font-bold">{{ stats.adol_m }} Masc.</span>
                        <span class="text-pink-400 font-bold">{{ stats.adol_f }} Fem.</span>
                    </div>
                </div>
            </div>

            <!-- Box Adultos -->
            <div class="bg-slate-800/50 p-5 rounded-2xl border border-slate-700/50 flex flex-col justify-between">
                <div>
                    <h4 class="text-[11px] lg:text-xs font-bold text-slate-400 uppercase tracking-widest text-center">Adultos (18-54)</h4>
                    <div class="text-3xl lg:text-4xl font-black text-white mt-2 text-center">{{ stats.adultos }}</div>
                </div>
                <div class="text-[10px] lg:text-xs text-slate-400 mt-4 border-t border-slate-700/50 pt-3">
                    <div class="flex justify-between mb-1">
                        <span class="text-blue-400 font-bold">{{ stats.adultos_m }} Masc.</span>
                        <span class="text-pink-400 font-bold">{{ stats.adultos_f }} Fem.</span>
                    </div>
                </div>
            </div>

            <!-- Box Muj Mayores -->
            <div class="bg-slate-800/50 p-5 rounded-2xl border border-slate-700/50 flex flex-col justify-between items-center text-center">
                <h4 class="text-[11px] lg:text-xs font-bold text-slate-400 uppercase tracking-widest">Muj. Mayores (≥55)</h4>
                <div class="text-3xl lg:text-4xl font-black text-pink-300 mt-2">{{ stats.muj_may }}</div>
                <div class="mt-auto pt-3 border-t border-slate-700/50 w-full"><i data-lucide="user" class="w-5 h-5 text-pink-400 mx-auto"></i></div>
            </div>

            <!-- Box Hom Mayores -->
            <div class="bg-slate-800/50 p-5 rounded-2xl border border-slate-700/50 flex flex-col justify-between items-center text-center">
                <h4 class="text-[11px] lg:text-xs font-bold text-slate-400 uppercase tracking-widest">Hom. Mayores (≥60)</h4>
                <div class="text-3xl lg:text-4xl font-black text-blue-300 mt-2">{{ stats.hom_may }}</div>
                <div class="mt-auto pt-3 border-t border-slate-700/50 w-full"><i data-lucide="user" class="w-5 h-5 text-blue-400 mx-auto"></i></div>
            </div>
        </div>
    </div>

    <!-- Indicadores de Salud y Remisiones -->
    <div class="grid grid-cols-1 md:grid-cols-3 gap-6 mb-8">
        <div class="bg-pink-900/20 border border-pink-800/40 p-6 rounded-3xl flex flex-col items-center justify-center">
            <i data-lucide="heart-pulse" class="w-8 h-8 text-pink-400 mb-3"></i>
            <h4 class="text-sm lg:text-base text-pink-300 font-bold uppercase tracking-widest mb-1 text-center">Embarazadas</h4>
            <div class="text-5xl font-black text-pink-400">{{ stats.embarazadas }}</div>
        </div>
        
        <div class="bg-teal-900/20 border border-teal-800/40 p-6 rounded-3xl flex flex-col items-center justify-center">
            <i data-lucide="accessibility" class="w-8 h-8 text-teal-400 mb-3"></i>
            <h4 class="text-sm lg:text-base text-teal-300 font-bold uppercase tracking-widest mb-1 text-center">Pers. con Condición</h4>
            <div class="text-5xl font-black text-teal-400">{{ stats.discapacitados }}</div>
        </div>

        <div class="bg-purple-900/20 border border-purple-800/40 p-6 rounded-3xl flex flex-col items-center justify-center">
            <i data-lucide="send" class="w-8 h-8 text-purple-400 mb-3"></i>
            <h4 class="text-sm lg:text-base text-purple-300 font-bold uppercase tracking-widest mb-1 text-center">Familias Egresadas</h4>
            <div class="flex items-center gap-6 mt-2">
                <div class="text-center">
                    <div class="text-4xl lg:text-5xl font-black text-purple-400">{{ manual_f_egr }}</div>
                    <div class="text-[10px] lg:text-xs text-purple-300 font-bold uppercase mt-1 tracking-wider">Familias</div>
                </div>
                <div class="h-10 w-px bg-purple-700/50"></div>
                <div class="text-center">
                    <div class="text-4xl lg:text-5xl font-black text-purple-400">{{ manual_p_egr }}</div>
                    <div class="text-[10px] lg:text-xs text-purple-300 font-bold uppercase mt-1 tracking-wider">Personas</div>
                </div>
            </div>
        </div>
    </div>

    <!-- Gráficos -->
    <div class="grid grid-cols-1 lg:grid-cols-2 gap-6 flex-grow pb-8">
        <div class="bg-slate-900/80 border border-slate-800 p-6 rounded-3xl flex flex-col h-full min-h-[300px]">
            <h3 class="text-sm text-slate-400 font-bold uppercase tracking-widest mb-4 flex items-center gap-2"><i data-lucide="bar-chart-2" class="w-4 h-4 text-indigo-400"></i> Distribución por Rangos de Edad</h3>
            <div class="relative flex-grow w-full">
                <canvas id="ageChart"></canvas>
            </div>
        </div>
        <div class="bg-slate-900/80 border border-slate-800 p-6 rounded-3xl flex flex-col h-full min-h-[300px]">
            <h3 class="text-sm text-slate-400 font-bold uppercase tracking-widest mb-4 flex items-center gap-2"><i data-lucide="map" class="w-4 h-4 text-teal-400"></i> Densidad Poblacional por Parroquia / Sector</h3>
            <div class="relative flex-grow w-full">
                <canvas id="parishChart"></canvas>
            </div>
        </div>
    </div>
    
    <button onclick="window.close()" class="absolute bottom-6 right-6 p-3 bg-slate-800/50 hover:bg-rose-500 rounded-full text-slate-400 hover:text-white transition-colors border border-slate-700 focus:outline-none z-50">
        <i data-lucide="x" class="w-6 h-6"></i>
    </button>

    <script>
        lucide.createIcons();
        function updateTime() {
            const now = new Date();
            document.getElementById('clock').innerText = now.toLocaleTimeString('es-ES', {hour: '2-digit', minute:'2-digit', second:'2-digit'});
        }
        setInterval(updateTime, 1000);
        updateTime();

        // Configuración Global Chart.js
        Chart.defaults.color = '#94a3b8';
        Chart.defaults.font.family = "'Plus Jakarta Sans', sans-serif";

        // Registro global de plugin DataLabels 
        Chart.register(ChartDataLabels);

        // Alto Contraste para DataLabels (Número en el medio de la barra)
        const customDataLabelsConfig = {
            color: '#ffffff',
            anchor: 'center',
            align: 'center',
            font: {
                weight: '900',
                size: 20,
                family: "'Plus Jakarta Sans', sans-serif"
            },
            textStrokeColor: 'rgba(2, 6, 23, 0.95)', // Borde oscuro denso
            textStrokeWidth: 5,
            formatter: function(value) {
                return value > 0 ? value : ''; 
            }
        };

        // Gráfico de Edades
        const ageCtx = document.getElementById('ageChart').getContext('2d');
        new Chart(ageCtx, {
            type: 'bar',
            data: {
                labels: {{ age_labels | tojson | safe }},
                datasets: [{
                    label: 'Habitantes',
                    data: {{ age_data | tojson | safe }},
                    backgroundColor: ['#3b82f6', '#06b6d4', '#6366f1', '#a855f7'],
                    borderRadius: 8,
                    borderSkipped: false
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { display: false },
                    datalabels: customDataLabelsConfig
                },
                scales: {
                    y: { 
                        beginAtZero: true, 
                        grid: { color: '#1e293b' },
                        ticks: { font: { size: 11 } }
                    },
                    x: { 
                        grid: { display: false },
                        ticks: { font: { size: 12, weight: 'bold' } }
                    }
                }
            }
        });

        // Gráfico de Parroquias
        const parishCtx = document.getElementById('parishChart').getContext('2d');
        new Chart(parishCtx, {
            type: 'bar',
            data: {
                labels: {{ parish_labels | tojson | safe }},
                datasets: [{
                    label: 'Habitantes por Sector',
                    data: {{ parish_data | tojson | safe }},
                    backgroundColor: '#10b981',
                    borderRadius: 8,
                    borderSkipped: false
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { display: false },
                    datalabels: customDataLabelsConfig
                },
                scales: {
                    y: { 
                        beginAtZero: true, 
                        grid: { color: '#1e293b' },
                        ticks: { font: { size: 11 } }
                    },
                    x: { 
                        grid: { display: false },
                        ticks: { font: { size: 11, weight: 'bold' } }
                    }
                }
            }
        });
    </script>
</body>
</html>
    """,
    'imprimir_ficha.html': """
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>Ficha Individual - {{ jefe.nombre }}</title>
    <script src="https://cdn.tailwindcss.com"></script>
</head>
<body class="p-8 bg-white text-black font-sans relative" onload="window.print(); setTimeout(()=>window.close(), 500);">
    
    {% if request.args.get('nro') %}
    <div class="absolute top-10 right-10 text-black border-4 border-black px-4 py-2 rounded-lg bg-white" style="font-size: 26px; font-weight: 900; z-index: 10;">
        N° FAMILIA: {{ request.args.get('nro') }}
    </div>
    {% endif %}

    <div class="max-w-3xl mx-auto border-2 border-black p-8 rounded-2xl mt-4">
        <div class="text-center border-b-2 border-black pb-4 mb-6">
            <h1 class="text-2xl font-black uppercase text-black">Ficha Individual de Control</h1>
            <p class="text-sm font-bold text-black">Sistema de Gestión Demográfica y Control Social</p>
        </div>
        
        <div class="grid grid-cols-2 gap-4 mb-6">
            <div>
                <p class="text-xs text-black uppercase font-bold">Jefe de Familia / Representante</p>
                <p class="text-lg font-black uppercase text-blue-700">{{ jefe.nombre }}</p>
            </div>
            <div>
                <p class="text-xs text-black uppercase font-bold">Cédula de Identidad</p>
                <p class="text-lg font-black text-black">{{ jefe.ci or 'S/C' }}</p>
            </div>
            <div>
                <p class="text-xs text-black uppercase font-bold">Edad y Género</p>
                <p class="text-base font-bold text-black">{% if jefe.edad == -1 %}N/D{% else %}{{ jefe.edad }}{% endif %} años - {{ jefe.genero }}</p>
            </div>
            <div>
                <p class="text-xs text-black uppercase font-bold">Teléfono de Contacto</p>
                <p class="text-base font-bold text-black">{{ jefe.telefono or 'No registrado' }}</p>
            </div>
        </div>
        
        <div class="bg-white p-4 rounded-xl mb-6 border-2 border-black">
            <div class="grid grid-cols-2 gap-4">
                <div>
                    <p class="text-xs text-black uppercase font-bold">Ubicación / Parroquia</p>
                    <p class="text-base font-black text-black">{{ jefe.parroquia }}</p>
                </div>
                <div>
                    <p class="text-xs text-black uppercase font-bold">Situación de Riesgo</p>
                    <p class="text-base font-black text-black uppercase">{{ jefe.situacion }}</p>
                </div>
                <div class="col-span-2">
                    <p class="text-xs text-black uppercase font-bold">Requerimientos o Entregas</p>
                    <p class="text-sm font-bold text-black uppercase">{{ jefe.requerimiento }}</p>
                </div>
                
                {% if jefe.lugar_remision and jefe.lugar_remision != '' %}
                <div class="col-span-2">
                    <p class="text-xs text-black uppercase font-bold">Remitido Hacia</p>
                    <p class="text-sm font-bold text-black uppercase">{{ jefe.lugar_remision }}</p>
                </div>
                {% endif %}
                
                <div class="col-span-2">
                    <p class="text-xs text-black uppercase font-bold">Observaciones Generales</p>
                    <p class="text-sm font-bold text-black uppercase">{{ jefe.observacion }}</p>
                </div>
                <div class="col-span-2">
                    <p class="text-xs text-black uppercase font-bold">Condiciones Médicas (Representante)</p>
                    <p class="text-sm font-bold text-black uppercase">Persona con Condición: {{ jefe.discapacidad }} | Patología: {{ jefe.patologia }} {% if jefe.es_embarazada %}| <strong class="text-black">EMBARAZADA</strong>{% endif %}</p>
                </div>
            </div>
        </div>

        <div>
            <h3 class="text-sm font-black text-black uppercase border-b-2 border-black pb-2 mb-4">Núcleo Familiar ({{ jefe.cargas|length }} Miembros Adicionales)</h3>
            {% if jefe.cargas %}
            <table class="w-full text-left border-collapse border-2 border-black">
                <thead>
                    <tr class="bg-white text-[10px] uppercase text-black border-b-2 border-black">
                        <th class="p-2 border border-black font-bold">Nombre Completo</th>
                        <th class="p-2 border border-black font-bold text-center">C.I.</th>
                        <th class="p-2 border border-black font-bold text-center">Edad/Gen</th>
                        <th class="p-2 border border-black font-bold">Parentesco</th>
                        <th class="p-2 border border-black font-bold">Alertas Médicas</th>
                    </tr>
                </thead>
                <tbody class="text-xs">
                    {% for m in jefe.cargas %}
                    <tr>
                        <td class="p-2 border border-black font-bold uppercase text-black">{{ m.nombre }}</td>
                        <td class="p-2 border border-black text-center font-mono text-black">{{ m.ci or 'S/C' }}</td>
                        <td class="p-2 border border-black text-center text-black">{% if m.edad == -1 %}N/D{% else %}{{ m.edad }}a{% endif %} - {{ m.genero }}</td>
                        <td class="p-2 border border-black uppercase font-bold text-black">{{ m.parentesco }}</td>
                        <td class="p-2 border border-black text-[10px] uppercase text-black">
                            {% if m.discapacidad and m.discapacidad|upper not in ['NINGUNA', 'NO', 'NINGUNO', ''] %}<strong>CONDICIÓN:</strong> {{ m.discapacidad }}<br>{% endif %}
                            {% if m.patologia and m.patologia|upper not in ['NINGUNA', 'NO', 'NINGUNO', ''] %}<strong>PATOL:</strong> {{ m.patologia }}<br>{% endif %}
                            {% if m.es_embarazada %}<strong>EMBARAZADA</strong>{% endif %}
                        </td>
                    </tr>
                    {% endfor %}
                </tbody>
            </table>
            {% else %}
            <p class="text-sm italic font-bold text-black">No hay cargas familiares vinculadas a este registro.</p>
            {% endif %}
        </div>
        
        <div class="mt-12 text-center text-[10px] text-black uppercase font-bold tracking-wider">
            <p>Documento generado por el Sistema de Gestión Demográfica</p>
            <p>Fecha de Impresión: <script>document.write(new Date().toLocaleDateString())</script></p>
        </div>
    </div>
</body>
</html>
    """,
    'carnet.html': """
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>Carnet Trabajador - {{ persona.nombre }}</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <style>
        /* Tamaño exacto estándar internacional de Carnet CR80 */
        @page { size: 54mm 86mm; margin: 0; }
        body { background-color: #f1f5f9; display: flex; justify-content: center; align-items: center; min-height: 100vh; margin: 0; font-family: 'Arial', sans-serif; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
        
        /* Contenedor principal del carnet */
        .carnet { 
            width: 54mm; 
            height: 86mm; 
            background-color: white; 
            position: relative; 
            box-shadow: 0 10px 25px rgba(0,0,0,0.15); 
            border-radius: 4px; 
            overflow: hidden; 
            display: flex; 
            flex-direction: column; 
        }
        
        /* Patrón de puntos tipo fondo corporativo */
        .bg-dots { 
            position: absolute; 
            inset: 0; 
            background-image: radial-gradient(#e2e8f0 1px, transparent 1px); 
            background-size: 8px 8px; 
            opacity: 0.8; 
            z-index: 0; 
        }
        
        /* Bloques rojos superior e inferior */
        .top-red { background-color: #ef4444; height: 16mm; width: 100%; position: absolute; top: 0; z-index: 1; }
        .bottom-red { background-color: #ef4444; height: 12mm; width: 100%; position: absolute; bottom: 0; z-index: 1; }
        
        /* Cabecera azul corporativo */
        .blue-card { 
            background-color: #1d4ed8; 
            margin: 4mm 4mm 0 4mm; 
            height: 16mm; 
            border-radius: 8px; 
            position: relative; 
            z-index: 2; 
            display: flex; 
            align-items: center; 
            justify-content: center; 
            text-align: center; 
            padding: 4px;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
        }
        
        /* Área de contenido del carnet */
        .content-wrapper { 
            flex: 1; 
            position: relative; 
            z-index: 2; 
            display: flex; 
            flex-direction: column; 
            align-items: center; 
            padding-top: 4mm; 
        }
        
        /* Marco de la fotografía */
        .photo-frame { 
            width: 25mm; 
            height: 31mm; 
            border: 3px solid #1d4ed8; 
            border-radius: 8px; 
            overflow: hidden; 
            background: white; 
            display: flex; 
            align-items: center; 
            justify-content: center; 
            margin-bottom: 3mm;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
        }
        .photo-frame img { width: 100%; height: 100%; object-fit: cover; }
        
        .btn-print { position: fixed; bottom: 20px; right: 20px; background: #1d4ed8; color: white; padding: 10px 20px; border-radius: 8px; cursor: pointer; z-index: 100; font-weight: bold; }
        
        @media print {
            body { background: none; align-items: flex-start; justify-content: flex-start; }
            .carnet { box-shadow: none; border: 1px dashed #cbd5e1; border-radius: 0; margin: 0; }
            .btn-print { display: none; }
        }
    </style>
</head>
<body>
    <div class="carnet">
        <div class="bg-dots"></div>
        <div class="top-red"></div>
        <div class="bottom-red"></div>
        
        <!-- CINTILLO AZUL SUPERIOR -->
        <div class="blue-card">
            <h1 class="text-white font-black uppercase leading-tight tracking-wide" style="font-size: 10px;">CAMPAMENTO GUAYANA<br>ESEQUIBA</h1>
        </div>

        <div class="content-wrapper">
            <!-- FOTOGRAFÍA -->
            <div class="photo-frame">
                {% if tipo == 'jefe' and persona.foto %}
                    <img src="{{ url_for('static', filename='uploads/' + persona.foto) }}" alt="Foto">
                {% else %}
                    <!-- Placeholder tipo silueta para los que no tienen foto en sistema -->
                    <svg class="w-12 h-12 text-slate-300" fill="currentColor" viewBox="0 0 24 24"><path d="M24 20.993V24H0v-2.996A14.977 14.977 0 0112.004 15c4.904 0 9.26 2.354 11.996 5.993zM16.002 8.999a4 4 0 11-8 0 4 4 0 018 0z" /></svg>
                {% endif %}
            </div>
            
            <!-- DATOS PERSONALES -->
            <h2 class="font-black text-[#0f172a] uppercase text-center w-full px-2 leading-tight" style="font-size: 13px;">{{ persona.nombre }}</h2>
            <p class="font-bold text-[#1e293b] mt-1 tracking-wider" style="font-size: 9px;">C.I. {{ persona.ci or 'NO REGISTRADA' }}</p>
            
            <!-- VINCULACIÓN AL NÚCLEO FAMILIAR -->
            <div class="mt-2 w-full text-center">
                <p class="font-black text-[#0f172a] tracking-widest" style="font-size: 9px;">FAMILIA N° {% if nro_familia %}{{ nro_familia }}{% else %}{{ jefe.id|string|zfill(4) }}{% endif %}</p>
                {% if tipo == 'carga' %}
                <p class="font-bold text-slate-600 mt-0.5 tracking-wider" style="font-size: 6px;">REP: {{ jefe.nombre }}</p>
                {% endif %}
            </div>
            
            <!-- ROL DEL TRABAJADOR DENTRO DE LA FAMILIA -->
            <p class="font-bold text-blue-700 uppercase tracking-widest mt-1" style="font-size: 6px;">{{ rol }}</p>
        </div>
        
        <!-- CINTILLO ROJO INFERIOR -->
        <div class="absolute bottom-1.5 w-full text-center z-10 flex flex-col items-center justify-center">
            <p class="text-white font-bold uppercase tracking-wider" style="font-size: 6.5px; line-height: 1.3;">AUTORIZACIÓN ACCESO LABORAL<br>PASE INTRANSFERIBLE</p>
        </div>
    </div>
    <button class="btn-print shadow-xl" onclick="window.print();">🖨️ Imprimir Carnet</button>
</body>
</html>
    """,
    'editar.html': """
{% extends 'base.html' %}
{% block content %}
<div class="max-w-5xl mx-auto bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
    <div class="px-4 py-4 md:px-6 md:py-5 bg-blue-50/50 border-b border-blue-100 flex flex-col md:flex-row md:items-center justify-between gap-3">
        <h2 class="text-base md:text-lg font-extrabold text-[#0f172a] flex items-center gap-2"><i data-lucide="pencil" class="text-blue-600 w-4 h-4 md:w-5 md:h-5"></i> Editar Familiar: {{ jefe.nombre }}</h2>
        <span class="bg-blue-100 text-blue-800 text-[9px] md:text-[10px] font-bold px-3 py-1 rounded-full uppercase tracking-wider w-fit">Modo Edición Admin</span>
    </div>
    
    <form method="POST" enctype="multipart/form-data" class="p-4 md:p-8">
        <input type="hidden" name="csrf_token" value="{{ csrf_token }}"/>
        <div class="space-y-4 md:space-y-6">
            <h3 class="text-xs md:text-sm font-bold text-slate-400 uppercase tracking-wider border-b border-slate-100 pb-2">Datos Principales (Jefe de Familia)</h3>
            <div class="grid grid-cols-1 md:grid-cols-3 gap-4 md:gap-5">
                <div class="md:col-span-2"><label class="block text-[10px] md:text-xs font-bold text-slate-600 mb-1.5 uppercase">Nombre Completo *</label><input type="text" name="nombre" value="{{ jefe.nombre }}" required class="w-full px-3.5 py-2.5 border border-slate-300 rounded-xl text-xs md:text-sm uppercase bg-white focus:ring-2 focus:ring-blue-500 outline-none transition-all"></div>
                <div><label class="block text-[10px] md:text-xs font-bold text-slate-600 mb-1.5 uppercase">Cédula de Identidad</label><input type="text" name="ci" value="{{ jefe.ci }}" class="w-full px-3.5 py-2.5 border border-slate-300 rounded-xl text-xs md:text-sm bg-white focus:ring-2 focus:ring-blue-500 outline-none transition-all font-mono font-bold text-blue-900 uppercase"></div>
                
                <div><label class="block text-[10px] md:text-xs font-bold text-slate-600 mb-1.5 uppercase">Edad *</label><input type="number" name="edad" value="{{ jefe.edad }}" required min="0" class="w-full px-3.5 py-2.5 border border-slate-300 rounded-xl text-xs md:text-sm bg-white focus:ring-2 focus:ring-blue-500 outline-none transition-all"></div>
                <div><label class="block text-[10px] md:text-xs font-bold text-slate-600 mb-1.5 uppercase">Género *</label>
                    <select name="genero" required class="w-full px-3.5 py-2.5 border border-slate-300 rounded-xl text-xs md:text-sm bg-white focus:ring-2 focus:ring-blue-500 outline-none transition-all">
                        <option value="M" {% if jefe.genero == 'M' %}selected{% endif %}>Masculino</option>
                        <option value="F" {% if jefe.genero == 'F' %}selected{% endif %}>Femenino</option>
                    </select>
                </div>
                <div><label class="block text-[10px] md:text-xs font-bold text-slate-600 mb-1.5 uppercase">Teléfono de Contacto</label><input type="text" name="telefono" value="{{ jefe.telefono }}" class="w-full px-3.5 py-2.5 border border-slate-300 rounded-xl text-xs md:text-sm bg-white focus:ring-2 focus:ring-blue-500 outline-none transition-all"></div>
                
                <div class="md:col-span-1">
                    <label class="block text-[10px] md:text-xs font-bold text-slate-600 mb-1.5 uppercase text-blue-600"><i data-lucide="image" class="inline w-3 h-3"></i> Actualizar Foto (Opcional)</label>
                    <input type="file" name="foto" accept="image/png, image/jpeg, image/webp" class="w-full text-xs text-slate-500 file:mr-2 file:py-1.5 file:px-3 file:rounded-lg file:border-0 file:text-[10px] md:file:text-xs file:font-bold file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100 cursor-pointer">
                </div>
                
                <div class="md:col-span-1"><label class="block text-[10px] md:text-xs font-bold text-slate-600 mb-1.5 uppercase">Situación / Riesgo *</label>
                    <select name="situacion" required class="w-full px-3.5 py-2.5 border border-slate-300 rounded-xl text-xs md:text-sm font-bold bg-white focus:ring-2 focus:ring-blue-500 outline-none transition-all">
                        <option value="Estable" {% if jefe.situacion == 'Estable' %}selected{% endif %}>Estable</option>
                        <option value="Refugio / Cancha" {% if 'Refugio' in jefe.situacion or 'Cancha' in jefe.situacion %}selected{% endif %}>Refugio / Cancha</option>
                        <option value="Riesgo Moderado" {% if jefe.situacion == 'Riesgo Moderado' %}selected{% endif %}>Riesgo Moderado</option>
                        <option value="Alto Riesgo" {% if jefe.situacion == 'Alto Riesgo' %}selected{% endif %}>Alto Riesgo</option>
                        <option value="Sin Casa" {% if jefe.situacion == 'Sin Casa' %}selected{% endif %}>Sin Casa</option>
                        <option value="Derrumbe" {% if jefe.situacion == 'Derrumbe' %}selected{% endif %}>Derrumbe</option>
                        <option value="Inhabitable" {% if jefe.situacion == 'Inhabitable' %}selected{% endif %}>Inhabitable</option>
                        <option value="Evaluación Médica" {% if jefe.situacion == 'Evaluación Médica' %}selected{% endif %}>Evaluación Médica</option>
                    </select>
                </div>

                <div class="md:col-span-1">
                    <label class="block text-[10px] md:text-xs font-bold text-slate-600 mb-1.5 uppercase text-rose-600"><i data-lucide="home" class="inline w-3 h-3"></i> Foto Vivienda/Etiqueta</label>
                    <input type="file" name="foto_vivienda" accept="image/png, image/jpeg, image/webp" class="w-full text-xs text-slate-500 file:mr-2 file:py-1.5 file:px-3 file:rounded-lg file:border-0 file:text-[10px] md:file:text-xs file:font-bold file:bg-rose-50 file:text-rose-700 hover:file:bg-rose-100 cursor-pointer">
                </div>
                
                <div class="md:col-span-3"><label class="block text-[10px] md:text-xs font-bold text-slate-600 mb-1.5 uppercase">Ubicación (Parroquia) *</label>
                    <select name="parroquia" required class="w-full px-3.5 py-2.5 border border-slate-300 rounded-xl text-xs md:text-sm font-bold bg-white text-blue-900 focus:ring-2 focus:ring-blue-500 outline-none transition-all">
                        <option value="{{ jefe.parroquia }}" selected>{{ jefe.parroquia }} (Actual)</option>
                        {% for p in parroquias %}{% if p.nombre != jefe.parroquia %}<option value="{{ p.nombre }}">{{ p.nombre }}</option>{% endif %}{% endfor %}
                    </select>
                </div>
                
                <div class="md:col-span-3"><label class="block text-[10px] md:text-xs font-bold text-slate-600 mb-1.5 uppercase">Requerimiento / Necesidad</label><textarea name="requerimiento" rows="2" class="w-full px-3.5 py-2.5 border border-slate-300 rounded-xl text-xs md:text-sm uppercase bg-white focus:ring-2 focus:ring-blue-500 outline-none transition-all">{{ jefe.requerimiento }}</textarea></div>
                <div class="md:col-span-3"><label class="block text-[10px] md:text-xs font-bold text-slate-600 mb-1.5 uppercase">Observación General</label><textarea name="observacion" rows="2" class="w-full px-3.5 py-2.5 border border-slate-300 rounded-xl text-xs md:text-sm uppercase bg-white focus:ring-2 focus:ring-blue-500 outline-none transition-all">{{ jefe.observacion }}</textarea></div>

                <!-- BLOQUE DE REMISIÓN EXPLÍCITO -->
                <div class="md:col-span-3 bg-purple-50 border border-purple-200 p-4 md:p-5 rounded-xl shadow-sm">
                    <div class="flex flex-col sm:flex-row gap-4 items-start sm:items-center">
                        <div class="flex-shrink-0">
                            <label class="flex items-center gap-2 cursor-pointer bg-white px-3 py-2 rounded-lg border border-purple-300 shadow-sm">
                                <input type="checkbox" id="es_remitido" name="es_remitido" value="1" onchange="toggleRemision()" class="w-5 h-5 text-purple-600 rounded border-purple-300 focus:ring-purple-500" {% if jefe.lugar_remision and jefe.lugar_remision != '' %}checked{% endif %}>
                                <span class="text-xs md:text-sm font-bold text-purple-800 uppercase"><i data-lucide="send" class="w-4 h-4 inline"></i> ¿El Caso Fue Remitido?</span>
                            </label>
                        </div>
                        <div class="flex-grow w-full">
                            <label class="block text-[10px] md:text-xs font-bold text-purple-700 mb-1 uppercase">¿Hacia dónde / A qué institución?</label>
                            <input type="text" id="lugar_remision" name="lugar_remision" value="{{ jefe.lugar_remision }}" placeholder="Ej: OBRAS PÚBLICAS, MISIÓN VIVIENDA, HOSPITAL UNIVERSITARIO..." class="w-full px-3.5 py-2 border border-purple-300 rounded-lg text-xs md:text-sm uppercase bg-white focus:ring-2 focus:ring-purple-500 outline-none transition-all disabled:opacity-50 disabled:bg-purple-100" {% if not jefe.lugar_remision %}disabled{% endif %}>
                        </div>
                    </div>
                </div>

                <div class="md:col-span-3 grid grid-cols-1 md:grid-cols-2 gap-4 md:gap-5 p-4 md:p-5 bg-slate-50 border border-slate-200 rounded-xl">
                    <div><label class="block text-[10px] md:text-xs font-bold text-teal-700 mb-1.5 uppercase flex items-center gap-1"><i data-lucide="accessibility" class="w-3 h-3"></i> Persona con Condición</label><input type="text" name="discapacidad" value="{{ jefe.discapacidad }}" class="w-full px-3.5 py-2.5 border border-slate-300 rounded-lg text-xs md:text-sm uppercase bg-white focus:ring-2 focus:ring-teal-400 outline-none transition-all"></div>
                    <div><label class="block text-[10px] md:text-xs font-bold text-amber-600 mb-1.5 uppercase flex items-center gap-1"><i data-lucide="stethoscope" class="w-3 h-3"></i> Patología Médica</label><input type="text" name="patologia" value="{{ jefe.patologia }}" class="w-full px-3.5 py-2.5 border border-slate-300 rounded-lg text-xs md:text-sm uppercase bg-white focus:ring-2 focus:ring-amber-400 outline-none transition-all"></div>
                    
                    <div class="md:col-span-2 pt-2 border-t border-slate-200 mt-2">
                        <label class="flex items-center gap-3 cursor-pointer bg-white p-3 rounded-xl border border-slate-200 shadow-sm hover:bg-slate-50 transition-colors w-fit">
                            <input type="checkbox" name="es_embarazada" value="1" {% if jefe.es_embarazada %}checked{% endif %} class="w-5 h-5 text-pink-600 rounded border-slate-300 focus:ring-pink-500">
                            <span class="text-xs md:text-sm font-bold text-slate-700 uppercase">Marcar a Jefe(a) como Mujer Embarazada 🤰</span>
                        </label>
                    </div>
                </div>
            </div>
        </div>
        
        <div class="mt-6 md:mt-8 border-t border-slate-200 pt-4 md:pt-6 overflow-x-hidden">
            <h3 class="text-xs md:text-sm font-bold text-slate-400 uppercase tracking-wider mb-4">Carga Familiar Actual</h3>
            {% if jefe.cargas %}
                <div class="space-y-4">
                {% for m in jefe.cargas %}
                    <div class="p-3 md:p-4 bg-slate-50 border border-slate-200 rounded-xl relative shadow-sm">
                        <input type="hidden" name="carga_id[]" value="{{ m.id }}">
                        
                        <div class="absolute top-2 right-2 md:top-3 md:right-3 z-10">
                            <button type="button" onclick="eliminarCargaExistente({{ m.id }})" class="text-rose-500 hover:text-white p-1 hover:bg-rose-500 rounded-lg border border-rose-100 bg-white transition-colors shadow-sm" title="Eliminar Familiar del Sistema">
                                <i data-lucide="trash-2" class="w-4 h-4"></i>
                            </button>
                        </div>
                        
                        <div class="grid grid-cols-1 md:grid-cols-5 gap-3 mt-4 md:mt-0">
                            <!-- AHORA INCLUIMOS EL CAMPO DE CEDULA PARA LAS CARGAS -->
                            <div class="md:col-span-1"><label class="block text-[9px] md:text-[10px] font-bold text-slate-500 uppercase mb-1">Nombre</label><input type="text" name="carga_nombre[]" value="{{ m.nombre }}" class="w-full px-3 py-2 border border-slate-300 rounded-lg text-xs uppercase bg-white"></div>
                            <div class="md:col-span-1"><label class="block text-[9px] md:text-[10px] font-bold text-slate-500 uppercase mb-1">C.I.</label><input type="text" name="carga_ci[]" value="{{ m.ci }}" class="w-full px-3 py-2 border border-slate-300 rounded-lg text-xs uppercase bg-white font-mono font-bold"></div>
                            
                            <div><label class="block text-[9px] md:text-[10px] font-bold text-slate-500 uppercase mb-1">Parentesco</label><input type="text" name="carga_parentesco[]" value="{{ m.parentesco }}" class="w-full px-3 py-2 border border-slate-300 rounded-lg text-xs uppercase bg-white"></div>
                            <div class="grid grid-cols-2 gap-2 md:block">
                                <div><label class="block text-[9px] md:text-[10px] font-bold text-slate-500 uppercase mb-1">Edad</label><input type="number" name="carga_edad[]" value="{{ m.edad }}" class="w-full px-3 py-2 border border-slate-300 rounded-lg text-xs bg-white"></div>
                                <div><label class="block text-[9px] md:text-[10px] font-bold text-slate-500 uppercase mb-1">Género</label>
                                    <select name="carga_genero[]" class="w-full px-3 py-2 border border-slate-300 rounded-lg text-xs bg-white">
                                        <option value="M" {% if m.genero == 'M' %}selected{% endif %}>Masculino</option>
                                        <option value="F" {% if m.genero == 'F' %}selected{% endif %}>Femenino</option>
                                    </select>
                                </div>
                            </div>
                            <div class="md:col-span-2"><label class="block text-[9px] md:text-[10px] font-bold text-slate-500 uppercase mb-1">Persona con Condición</label><input type="text" name="carga_discapacidad[]" value="{{ m.discapacidad }}" class="w-full px-3 py-2 border border-slate-300 rounded-lg text-xs uppercase bg-white"></div>
                            <div class="md:col-span-2"><label class="block text-[9px] md:text-[10px] font-bold text-slate-500 uppercase mb-1">Patología</label><input type="text" name="carga_patologia[]" value="{{ m.patologia }}" class="w-full px-3 py-2 border border-slate-300 rounded-lg text-xs uppercase bg-white"></div>
                            <div class="flex items-end pb-0.5">
                                <label class="flex items-center gap-2 cursor-pointer bg-white p-2 rounded-lg border border-slate-200 shadow-sm w-full h-[34px]">
                                    <input type="checkbox" name="carga_embarazo_{{ m.id }}" value="1" {% if m.es_embarazada %}checked{% endif %} class="w-4 h-4 text-pink-600 rounded">
                                    <span class="text-[9px] md:text-[10px] font-bold text-pink-700 uppercase">¿Embarazada?</span>
                                </label>
                            </div>
                        </div>
                    </div>
                {% endfor %}
                </div>
            {% else %}
                <p class="text-xs md:text-sm text-slate-500 italic bg-slate-50 p-4 rounded-xl border border-slate-200">No hay cargas familiares registradas para este jefe.</p>
            {% endif %}
            
            <div class="mt-6 flex flex-col sm:flex-row justify-between sm:items-center gap-3 border-t border-slate-100 pt-4">
                 <h4 class="text-[10px] md:text-xs font-bold text-slate-500 uppercase">Añadir Nuevas Cargas al Núcleo</h4>
                 <button type="button" onclick="addFamilyMember()" class="bg-emerald-50 text-emerald-700 px-3 py-2 md:py-1.5 rounded-lg text-xs font-bold border border-emerald-200 shadow-sm hover:bg-emerald-100 transition-colors text-center w-full sm:w-auto">Añadir Nuevo Miembro +</button>
            </div>
            <div id="miembros-container" class="space-y-4 mt-4"></div>
        </div>

        <div class="border-t border-slate-200 mt-6 pt-6 flex flex-col-reverse sm:flex-row justify-end gap-3">
            <a href="{{ url_for('index') }}" class="px-6 py-3 rounded-xl border border-slate-300 text-slate-700 font-bold text-sm hover:bg-slate-50 transition-colors text-center w-full sm:w-auto">Cancelar y Volver</a>
            <button type="submit" class="px-6 py-3 rounded-xl bg-blue-600 hover:bg-blue-700 text-white font-bold text-sm shadow-md flex items-center justify-center gap-2 transition-colors w-full sm:w-auto"><i data-lucide="save" class="w-4 h-4"></i> Guardar Cambios</button>
        </div>
    </form>
</div>

<!-- Form oculto para usar CSRF en eliminación vía Fetch -->
<form id="deleteCargaForm" style="display:none;">
    <input type="hidden" name="csrf_token" id="delete_csrf_token" value="{{ csrf_token }}"/>
</form>

<script>
    function toggleRemision() {
        const checkbox = document.getElementById('es_remitido');
        const input = document.getElementById('lugar_remision');
        if (checkbox.checked) {
            input.disabled = false;
            input.required = true;
            input.focus();
        } else {
            input.disabled = true;
            input.required = false;
            input.value = '';
        }
    }

    let miembroCount = 0;
    function addFamilyMember() {
        miembroCount++;
        const div = document.createElement('div');
        div.className = "bg-white border border-emerald-200 shadow-sm rounded-xl p-3 md:p-4 relative";
        div.id = `miembro-${miembroCount}`;
        div.innerHTML = `
            <button type="button" onclick="document.getElementById('miembro-${miembroCount}').remove();" class="absolute top-2 right-2 md:top-3 md:right-3 text-rose-500 hover:text-white p-1 hover:bg-rose-500 rounded-lg border border-rose-100 transition-colors"><i data-lucide="x" class="w-4 h-4"></i></button>
            <div class="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-5 gap-3 md:gap-4 pr-6 md:pr-8">
                <div class="md:col-span-1"><label class="block text-[9px] md:text-[10px] font-bold text-slate-500 mb-1 uppercase">Nombre *</label><input type="text" name="m_nombre[]" required class="w-full px-3 py-2 border border-slate-200 rounded-lg text-xs uppercase bg-slate-50 focus:ring-1 focus:ring-emerald-500 outline-none"></div>
                <div class="md:col-span-1"><label class="block text-[9px] md:text-[10px] font-bold text-slate-500 mb-1 uppercase">C.I.</label><input type="text" name="m_ci[]" class="w-full px-3 py-2 border border-slate-200 rounded-lg text-xs uppercase bg-slate-50 font-mono font-bold focus:ring-1 focus:ring-emerald-500 outline-none"></div>
                <div><label class="block text-[9px] md:text-[10px] font-bold text-slate-500 mb-1 uppercase">Parentesco *</label><input type="text" name="m_parentesco[]" required class="w-full px-3 py-2 border border-slate-200 rounded-lg text-xs uppercase bg-slate-50 focus:ring-1 focus:ring-emerald-500 outline-none"></div>
                <div class="grid grid-cols-2 gap-2 md:block">
                    <div><label class="block text-[9px] md:text-[10px] font-bold text-slate-500 mb-1 uppercase">Edad *</label><input type="number" name="m_edad[]" required min="0" class="w-full px-3 py-2 border border-slate-200 rounded-lg text-xs bg-slate-50 focus:ring-1 focus:ring-emerald-500 outline-none"></div>
                    <div><label class="block text-[9px] md:text-[10px] font-bold text-slate-500 mb-1 uppercase">Género *</label><select name="m_genero[]" required class="w-full px-3 py-2 border border-slate-200 rounded-lg text-xs"><option value="M">Masculino</option><option value="F">Femenino</option></select></div>
                </div>
                <div class="md:col-span-2 flex flex-col md:flex-row gap-3"><div class="flex-1"><label class="block text-[9px] md:text-[10px] font-bold text-slate-500 mb-1 uppercase">Persona con Condición</label><input type="text" name="m_discapacidad[]" class="w-full px-3 py-2 border border-slate-200 rounded-lg text-xs uppercase focus:ring-1 focus:ring-emerald-500 outline-none"></div><div class="flex-1"><label class="block text-[9px] md:text-[10px] font-bold text-slate-500 mb-1 uppercase">Patología</label><input type="text" name="m_patologia[]" class="w-full px-3 py-2 border border-slate-200 rounded-lg text-xs uppercase focus:ring-1 focus:ring-emerald-500 outline-none"></div></div>
                <div class="flex items-center justify-center bg-pink-50 border border-pink-200 rounded-lg p-2"><label class="flex items-center gap-2 cursor-pointer"><input type="checkbox" name="m_embarazo[]" value="${miembroCount-1}" class="w-4 h-4 text-pink-600 rounded"><span class="text-[9px] md:text-[10px] font-bold text-pink-800 uppercase">¿Embarazada?</span></label></div>
            </div>`;
        document.getElementById('miembros-container').appendChild(div); lucide.createIcons();
    }
    
    function eliminarCargaExistente(id) {
        if(confirm('¿Seguro que deseas eliminar definitivamente a este familiar del sistema?')) {
            const csrfToken = document.getElementById('delete_csrf_token').value;
            const formData = new FormData();
            formData.append('csrf_token', csrfToken);
            
            fetch('/eliminar_carga/' + id, { method: 'POST', body: formData })
            .then(response => {
                if(response.ok) {
                    window.location.reload();
                } else {
                    alert('Error al intentar eliminar.');
                }
            });
        }
    }
</script>
{% endblock %}
    """,
    'registrar.html': """
{% extends 'base.html' %}
{% block content %}
<div class="max-w-4xl mx-auto bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
    <div class="px-4 py-4 md:px-6 md:py-5 bg-slate-50 border-b border-slate-100 flex flex-col sm:flex-row justify-between sm:items-center gap-2">
        <h2 class="text-base md:text-lg font-extrabold text-[#0f172a] flex items-center gap-2"><i data-lucide="users" class="text-blue-600 w-5 h-5"></i> Registro de Familia Afectada</h2>
        <span class="text-[10px] md:text-xs text-slate-500"><i data-lucide="info" class="inline w-3 h-3 md:w-3 md:h-3 text-blue-500"></i> Las fotos se comprimen automáticamente</span>
    </div>
    <form method="POST" enctype="multipart/form-data" class="p-4 md:p-8 space-y-6 md:space-y-8">
        <input type="hidden" name="csrf_token" value="{{ csrf_token }}"/>
        <div>
            <div class="flex flex-col sm:flex-row justify-between sm:items-center gap-3 mb-4 border-b border-slate-100 pb-2">
                <h3 class="text-xs md:text-sm font-bold text-slate-400 uppercase tracking-wider">1. Datos del Jefe(a) de Familia</h3>
                <label class="flex items-center justify-center gap-2 cursor-pointer bg-pink-50 px-3 py-1.5 rounded-lg border border-pink-200"><input type="checkbox" name="embarazada_jefe" value="1" class="w-4 h-4 text-pink-600 rounded"><span class="text-[10px] md:text-xs font-bold text-pink-800 uppercase">¿Jefa Embarazada?</span></label>
            </div>
            <div class="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-4 md:gap-5">
                <div class="md:col-span-2"><label class="block text-[10px] md:text-xs font-bold text-slate-600 mb-1.5 uppercase">Nombre *</label><input type="text" name="nombre" required class="w-full px-3.5 py-2.5 border rounded-xl text-xs md:text-sm uppercase bg-slate-50 focus:ring-2 focus:ring-blue-500 outline-none"></div>
                <div><label class="block text-[10px] md:text-xs font-bold text-slate-600 mb-1.5 uppercase">C.I.</label><input type="text" name="ci" class="w-full px-3.5 py-2.5 border rounded-xl text-xs md:text-sm uppercase bg-slate-50 font-mono font-bold focus:ring-2 focus:ring-blue-500 outline-none"></div>
                <div><label class="block text-[10px] md:text-xs font-bold text-slate-600 mb-1.5 uppercase">Edad *</label><input type="number" name="edad" required min="0" class="w-full px-3.5 py-2.5 border rounded-xl text-xs md:text-sm bg-slate-50 focus:ring-2 focus:ring-blue-500 outline-none"></div>
                <div><label class="block text-[10px] md:text-xs font-bold text-slate-600 mb-1.5 uppercase">Género *</label><select name="genero" required class="w-full px-3.5 py-2.5 border rounded-xl text-xs md:text-sm bg-white"><option value="M">Masculino</option><option value="F">Femenino</option></select></div>
                <div><label class="block text-[10px] md:text-xs font-bold text-slate-600 mb-1.5 uppercase">Teléfono</label><input type="text" name="telefono" class="w-full px-3.5 py-2.5 border rounded-xl text-xs md:text-sm bg-slate-50 focus:ring-2 focus:ring-blue-500 outline-none"></div>
                
                <div class="md:col-span-1"><label class="block text-[10px] md:text-xs font-bold text-slate-600 mb-1.5 uppercase text-blue-600"><i data-lucide="image" class="inline w-3 h-3"></i> Foto (Opcional)</label><input type="file" name="foto" accept="image/png, image/jpeg, image/webp" class="w-full text-xs text-slate-500 file:mr-2 file:py-1.5 file:px-3 file:rounded-lg file:border-0 file:text-[10px] md:file:text-xs file:font-bold file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100 cursor-pointer"></div>
                <div class="md:col-span-1"><label class="block text-[10px] md:text-xs font-bold text-slate-600 mb-1.5 uppercase">Situación Actual *</label>
                    <select name="situacion" required class="w-full px-3.5 py-2.5 border rounded-xl text-xs md:text-sm font-bold bg-white">
                        <option value="Estable">Estable</option>
                        <option value="Refugio / Cancha">Refugio / Cancha</option>
                        <option value="Riesgo Moderado">Riesgo Moderado</option>
                        <option value="Alto Riesgo">Alto Riesgo</option>
                        <option value="Sin Casa">Sin Casa</option>
                        <option value="Derrumbe">Derrumbe</option>
                        <option value="Inhabitable">Inhabitable</option>
                        <option value="Evaluación Médica">Evaluación Médica</option>
                    </select>
                </div>

                <div class="md:col-span-1">
                    <label class="block text-[10px] md:text-xs font-bold text-slate-600 mb-1.5 uppercase text-rose-600"><i data-lucide="home" class="inline w-3 h-3"></i> Foto Vivienda/Etiqueta</label>
                    <input type="file" name="foto_vivienda" accept="image/png, image/jpeg, image/webp" class="w-full text-xs text-slate-500 file:mr-2 file:py-1.5 file:px-3 file:rounded-lg file:border-0 file:text-[10px] md:file:text-xs file:font-bold file:bg-rose-50 file:text-rose-700 hover:file:bg-rose-100 cursor-pointer">
                </div>

                <div class="md:col-span-3"><label class="block text-[10px] md:text-xs font-bold text-slate-600 mb-1.5 uppercase">Ubicación (Parroquia / Estado) *</label><select name="parroquia" required class="w-full px-3.5 py-2.5 border rounded-xl text-xs md:text-sm font-bold bg-white text-blue-900">{% for p in parroquias %}<option value="{{ p.nombre }}">{{ p.nombre }}</option>{% endfor %}</select></div>
                
                <div class="md:col-span-3"><label class="block text-[10px] md:text-xs font-bold text-slate-600 mb-1.5 uppercase">Requerimiento / Necesidad</label><textarea name="requerimiento" rows="2" class="w-full px-3.5 py-2.5 border rounded-xl text-xs md:text-sm uppercase bg-slate-50 focus:ring-2 focus:ring-blue-500 outline-none"></textarea></div>
                <div class="md:col-span-3"><label class="block text-[10px] md:text-xs font-bold text-slate-600 mb-1.5 uppercase">Observación General</label><textarea name="observacion" rows="2" class="w-full px-3.5 py-2.5 border rounded-xl text-xs md:text-sm uppercase bg-slate-50 focus:ring-2 focus:ring-blue-500 outline-none"></textarea></div>

                <!-- BLOQUE DE REMISIÓN EXPLÍCITO -->
                <div class="md:col-span-3 bg-purple-50 border border-purple-200 p-4 md:p-5 rounded-xl shadow-sm">
                    <div class="flex flex-col sm:flex-row gap-4 items-start sm:items-center">
                        <div class="flex-shrink-0">
                            <label class="flex items-center gap-2 cursor-pointer bg-white px-3 py-2 rounded-lg border border-purple-300 shadow-sm">
                                <input type="checkbox" id="es_remitido" name="es_remitido" value="1" onchange="toggleRemision()" class="w-5 h-5 text-purple-600 rounded border-purple-300 focus:ring-purple-500">
                                <span class="text-xs md:text-sm font-bold text-purple-800 uppercase"><i data-lucide="send" class="w-4 h-4 inline"></i> ¿El Caso Fue Remitido?</span>
                            </label>
                        </div>
                        <div class="flex-grow w-full">
                            <label class="block text-[10px] md:text-xs font-bold text-purple-700 mb-1 uppercase">¿Hacia dónde / A qué institución?</label>
                            <input type="text" id="lugar_remision" name="lugar_remision" placeholder="Ej: OBRAS PÚBLICAS, MISIÓN VIVIENDA, HOSPITAL UNIVERSITARIO..." class="w-full px-3.5 py-2 border border-purple-300 rounded-lg text-xs md:text-sm uppercase bg-white focus:ring-2 focus:ring-purple-500 outline-none transition-all disabled:opacity-50 disabled:bg-purple-100" disabled>
                        </div>
                    </div>
                </div>

                <div class="md:col-span-3 grid grid-cols-1 md:grid-cols-2 gap-4 md:gap-5 p-4 bg-slate-50 border rounded-xl">
                    <div><label class="block text-[9px] md:text-[10px] font-bold text-teal-700 mb-1.5 uppercase"><i data-lucide="accessibility" class="w-3 h-3 inline"></i> Persona con Condición</label><input type="text" name="discapacidad" class="w-full px-3.5 py-2 border rounded-lg text-xs md:text-sm uppercase focus:ring-2 focus:ring-teal-500 outline-none"></div>
                    <div><label class="block text-[9px] md:text-[10px] font-bold text-amber-600 mb-1.5 uppercase">Patología</label><input type="text" name="patologia" class="w-full px-3.5 py-2 border rounded-lg text-xs md:text-sm uppercase focus:ring-2 focus:ring-amber-500 outline-none"></div>
                </div>
            </div>
        </div>

        <div class="overflow-x-hidden">
            <div class="flex flex-col sm:flex-row justify-between sm:items-center gap-3 mb-4 border-b border-slate-100 pb-2">
                <h3 class="text-xs md:text-sm font-bold text-slate-400 uppercase tracking-wider">2. Núcleo Familiar (Cargas)</h3>
                <button type="button" onclick="addFamilyMember()" class="bg-emerald-50 text-emerald-700 px-3 py-2 md:py-1.5 rounded-lg text-xs font-bold border border-emerald-200 shadow-sm hover:bg-emerald-100 transition-colors w-full sm:w-auto text-center">Añadir Miembro +</button>
            </div>
            <div id="miembros-container" class="space-y-4"></div>
        </div>
        <div class="border-t border-slate-200 pt-6 flex flex-col-reverse sm:flex-row justify-end gap-3">
            <a href="{{ url_for('index') }}" class="px-5 py-3 md:py-2.5 rounded-xl border text-slate-700 font-semibold text-sm hover:bg-slate-50 transition-colors w-full sm:w-auto text-center">Cancelar</a>
            <button type="submit" class="px-5 py-3 md:py-2.5 rounded-xl bg-blue-600 text-white font-bold text-sm shadow-md flex items-center justify-center gap-2 hover:bg-blue-700 transition-colors w-full sm:w-auto"><i data-lucide="save" class="w-4 h-4"></i> Guardar Registro</button>
        </div>
    </form>
</div>
<script>
    function toggleRemision() {
        const checkbox = document.getElementById('es_remitido');
        const input = document.getElementById('lugar_remision');
        if (checkbox.checked) {
            input.disabled = false;
            input.required = true;
            input.focus();
        } else {
            input.disabled = true;
            input.required = false;
            input.value = '';
        }
    }

    let miembroCount = 0;
    function addFamilyMember() {
        miembroCount++;
        const div = document.createElement('div');
        div.className = "bg-white border border-emerald-200 shadow-sm rounded-xl p-3 md:p-4 relative";
        div.id = `miembro-${miembroCount}`;
        div.innerHTML = `
            <button type="button" onclick="document.getElementById('miembro-${miembroCount}').remove();" class="absolute top-2 right-2 md:top-3 md:right-3 text-rose-500 hover:text-white p-1 hover:bg-rose-500 rounded-lg border border-rose-100 transition-colors"><i data-lucide="x" class="w-4 h-4"></i></button>
            <div class="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-5 gap-3 md:gap-4 pr-6 md:pr-8">
                <div class="md:col-span-1"><label class="block text-[9px] md:text-[10px] font-bold text-slate-500 mb-1 uppercase">Nombre *</label><input type="text" name="m_nombre[]" required class="w-full px-3 py-2 border border-slate-200 rounded-lg text-xs uppercase bg-slate-50 focus:ring-1 focus:ring-emerald-500 outline-none"></div>
                <div class="md:col-span-1"><label class="block text-[9px] md:text-[10px] font-bold text-slate-500 mb-1 uppercase">C.I.</label><input type="text" name="m_ci[]" class="w-full px-3 py-2 border border-slate-200 rounded-lg text-xs uppercase bg-slate-50 font-mono font-bold focus:ring-1 focus:ring-emerald-500 outline-none"></div>
                <div><label class="block text-[9px] md:text-[10px] font-bold text-slate-500 mb-1 uppercase">Parentesco *</label><input type="text" name="m_parentesco[]" required class="w-full px-3 py-2 border border-slate-200 rounded-lg text-xs uppercase bg-slate-50 focus:ring-1 focus:ring-emerald-500 outline-none"></div>
                <div class="grid grid-cols-2 gap-2 md:block">
                    <div><label class="block text-[9px] md:text-[10px] font-bold text-slate-500 mb-1 uppercase">Edad *</label><input type="number" name="m_edad[]" required min="0" class="w-full px-3 py-2 border border-slate-200 rounded-lg text-xs bg-slate-50 focus:ring-1 focus:ring-emerald-500 outline-none"></div>
                    <div><label class="block text-[9px] md:text-[10px] font-bold text-slate-500 mb-1 uppercase">Género *</label><select name="m_genero[]" required class="w-full px-3 py-2 border border-slate-200 rounded-lg text-xs"><option value="M">Masculino</option><option value="F">Femenino</option></select></div>
                </div>
                <div class="md:col-span-2 flex flex-col md:flex-row gap-3"><div class="flex-1"><label class="block text-[9px] md:text-[10px] font-bold text-slate-500 mb-1 uppercase">Persona con Condición</label><input type="text" name="m_discapacidad[]" class="w-full px-3 py-2 border border-slate-200 rounded-lg text-xs uppercase focus:ring-1 focus:ring-emerald-500 outline-none"></div><div class="flex-1"><label class="block text-[9px] md:text-[10px] font-bold text-slate-500 mb-1 uppercase">Patología</label><input type="text" name="m_patologia[]" class="w-full px-3 py-2 border border-slate-200 rounded-lg text-xs uppercase focus:ring-1 focus:ring-emerald-500 outline-none"></div></div>
                <div class="flex items-center justify-center bg-pink-50 border border-pink-200 rounded-lg p-2"><label class="flex items-center gap-2 cursor-pointer"><input type="checkbox" name="m_embarazo[]" value="${miembroCount-1}" class="w-4 h-4 text-pink-600 rounded"><span class="text-[9px] md:text-[10px] font-bold text-pink-800 uppercase">¿Embarazada?</span></label></div>
            </div>`;
        document.getElementById('miembros-container').appendChild(div); lucide.createIcons();
    }
</script>
{% endblock %}
    """,
    'cargar.html': """
{% extends 'base.html' %}
{% block content %}
<div class="max-w-3xl mx-auto space-y-4 md:space-y-6">
    <!-- SECCIÓN DE CARGA MASIVA -->
    <div class="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
        <div class="px-4 py-4 md:px-6 md:py-5 bg-emerald-50 border-b border-emerald-100 flex items-center gap-2"><i data-lucide="upload-cloud" class="text-emerald-600 w-5 h-5 md:w-6 md:h-6"></i><h2 class="text-base md:text-lg font-extrabold text-emerald-900">Carga Inteligente de Afectados</h2></div>
        <div class="p-4 md:p-8">
            <div class="mb-4 md:mb-6 bg-blue-50 p-3 md:p-4 rounded-xl border border-blue-100 text-xs md:text-sm text-blue-800 shadow-sm">
                <p class="font-bold mb-2 flex items-center gap-2"><i data-lucide="info" class="w-4 h-4"></i> Optimizaciones Activas del Motor de Lectura:</p>
                <ol class="list-decimal list-inside space-y-1 ml-1 text-[10px] md:text-xs font-medium">
                    <li><strong>Lectura Jerárquica Refinada:</strong> Detecta el número de familia para identificar al representante y sus cargas correspondientes.</li>
                    <li><strong>Carga Explícita de Nombres:</strong> Lee los nombres, edades y parentescos reales directamente de las filas.</li>
                    <li><strong>Detección de Marcas Cruzadas:</strong> Detecta automáticamente las "X", los "1" o "SÍ" en las columnas de Booleanos.</li>
                    <li><strong>Universal para Varias Hojas:</strong> Soporta múltiples hojas Excel de forma automática.</li>
                    <li class="text-blue-900 mt-1"><strong>Actualización Inteligente (No Destructiva):</strong> Si una persona ya existe, el sistema evalúa campo por campo. Si habías completado la información manualmente (ej. agregaste un teléfono o patología), se respetará tu trabajo. Solo se llenarán los campos vacíos o por defecto.</li>
                </ol>
            </div>
            <form method="POST" action="{{ url_for('cargar_datos') }}" enctype="multipart/form-data" class="space-y-4">
                <input type="hidden" name="csrf_token" value="{{ csrf_token }}"/>
                <div class="border-2 border-dashed border-slate-300 rounded-xl p-6 md:p-8 text-center bg-slate-50 hover:bg-slate-100 transition-colors">
                    <i data-lucide="file-spreadsheet" class="w-10 h-10 md:w-12 md:h-12 text-slate-400 mx-auto mb-3"></i>
                    <label class="block text-xs md:text-sm font-bold text-slate-700 mb-3">Selecciona tu archivo (.xlsx, .xls, .csv)</label>
                    <input type="file" name="archivo" accept=".xlsx, .xls, .csv" required class="w-full max-w-xs mx-auto text-xs text-slate-500 file:mr-4 file:py-2 file:px-4 file:rounded-full file:border-0 file:text-[10px] md:file:text-sm file:font-bold file:bg-emerald-100 file:text-emerald-700 hover:file:bg-emerald-200 cursor-pointer block transition-colors">
                </div>
                <button type="submit" class="w-full bg-emerald-600 hover:bg-emerald-700 text-white font-bold py-3.5 rounded-xl text-xs md:text-sm transition-all shadow-md flex items-center justify-center gap-2"><i data-lucide="upload" class="w-4 h-4"></i> Procesar y Cruzar Datos ⚡</button>
            </form>
        </div>
    </div>

    <!-- SECCIÓN DE RESPALDO DE BASE DE DATOS -->
    <div class="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
        <div class="px-4 py-4 md:px-6 md:py-5 bg-amber-50 border-b border-amber-100 flex items-center justify-between">
            <div class="flex items-center gap-2">
                <i data-lucide="database-backup" class="text-amber-600 w-5 h-5 md:w-6 md:h-6"></i>
                <h2 class="text-base md:text-lg font-extrabold text-amber-900">Respaldo Físico</h2>
            </div>
        </div>
        <div class="p-4 md:p-8 text-center">
            <p class="text-xs md:text-sm text-slate-600 mb-5 font-medium">Descarga una copia exacta y cifrada del archivo maestro actual (.db) para resguardo o auditoría técnica.</p>
            <a href="{{ url_for('respaldar_bd') }}" class="w-full md:w-auto inline-flex bg-amber-600 hover:bg-amber-700 text-white font-bold py-3.5 md:py-3 px-8 rounded-xl text-xs md:text-sm transition-all shadow-md items-center justify-center gap-2">
                <i data-lucide="download" class="w-4 h-4"></i> Descargar Backup (.db)
            </a>
        </div>
    </div>
</div>
{% endblock %}
    """,
    'historial.html': """
{% extends 'base.html' %}
{% block content %}
<div class="max-w-5xl mx-auto bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
    <div class="px-4 md:px-6 py-4 md:py-5 bg-slate-50 border-b border-slate-100 flex items-center gap-2">
        <i data-lucide="history" class="text-amber-600 w-5 h-5"></i>
        <h2 class="text-base md:text-lg font-extrabold text-[#0f172a]">Auditoría e Historial de Cambios</h2>
    </div>
    <div class="p-0 overflow-x-auto w-full hide-scrollbar">
        <table class="w-full text-left min-w-[600px]">
            <thead class="bg-slate-50 border-b border-slate-200 text-[10px] md:text-[11px] text-slate-500 uppercase tracking-wider">
                <tr>
                    <th class="px-4 md:px-6 py-3 md:py-4 font-bold">Fecha / Hora</th>
                    <th class="px-4 md:px-6 py-3 md:py-4 font-bold">Usuario</th>
                    <th class="px-4 md:px-6 py-3 md:py-4 font-bold">Acción</th>
                    <th class="px-4 md:px-6 py-3 md:py-4 font-bold">Detalle</th>
                </tr>
            </thead>
            <tbody class="divide-y divide-slate-100 text-xs md:text-sm">
                {% for h in historial %}
                <tr class="hover:bg-slate-50/50 transition-colors">
                    <td class="px-4 md:px-6 py-3 md:py-4 text-slate-500 font-mono text-[10px] md:text-xs font-semibold">{{ h.fecha.strftime('%d/%m/%Y %H:%M') }}</td>
                    <td class="px-4 md:px-6 py-3 md:py-4 font-bold text-slate-700 flex items-center gap-1.5"><i data-lucide="user" class="w-3 h-3 text-slate-400"></i> {{ h.usuario }}</td>
                    <td class="px-4 md:px-6 py-3 md:py-4">
                        <span class="px-2 py-1.5 rounded text-[9px] md:text-[10px] font-bold tracking-wider uppercase
                        {% if h.accion == 'CARGA MASIVA' or 'IMPORTACIÓN' in h.accion or 'EFICIENTE' in h.accion %}bg-emerald-100 text-emerald-700
                        {% elif h.accion == 'ELIMINACIÓN' %}bg-rose-100 text-rose-700
                        {% elif h.accion == 'NUEVO REGISTRO' %}bg-blue-100 text-blue-700
                        {% elif h.accion == 'RESPALDO BD' %}bg-purple-100 text-purple-700
                        {% else %}bg-amber-100 text-amber-700{% endif %}">
                            {{ h.accion }}
                        </span>
                    </td>
                    <td class="px-4 md:px-6 py-3 md:py-4 text-slate-600 text-[11px] md:text-[13px] font-medium">{{ h.detalle }}</td>
                </tr>
                {% endfor %}
            </tbody>
        </table>
        {% if not historial %}
        <div class="p-12 text-center flex flex-col items-center text-slate-400">
            <i data-lucide="inbox" class="w-10 h-10 md:w-12 md:h-12 mb-3 text-slate-300"></i>
            <span class="font-bold text-slate-500 text-xs md:text-sm">No hay registros en la auditoría.</span>
        </div>
        {% endif %}
    </div>
</div>
{% endblock %}
    """,
    'usuarios.html': """
{% extends 'base.html' %}
{% block content %}
<div class="max-w-5xl mx-auto bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
    <div class="px-4 py-4 md:px-6 md:py-5 bg-slate-50 border-b border-slate-100 flex items-center gap-2"><i data-lucide="users-round" class="text-blue-600 w-5 h-5"></i><h2 class="text-base md:text-lg font-extrabold text-[#0f172a]">Administración de Usuarios y Permisos</h2></div>
    <div class="p-4 md:p-8 overflow-x-hidden">
        {% if session.get('username') == 'admin' %}
        <form method="POST" class="bg-slate-50 p-4 md:p-5 rounded-xl border border-slate-200 mb-6 md:mb-8 shadow-sm">
            <input type="hidden" name="csrf_token" value="{{ csrf_token }}"/>
            <h3 class="text-[10px] md:text-xs font-bold text-slate-500 uppercase tracking-wider mb-3 md:mb-4 border-b border-slate-200 pb-2">Añadir Nuevo Usuario</h3>
            <div class="grid grid-cols-1 md:grid-cols-4 gap-3 md:gap-4 items-end">
                <div><label class="block text-[9px] md:text-[10px] font-bold text-slate-500 mb-1 uppercase">Username</label><input type="text" name="username" required class="w-full px-3 py-2 border border-slate-300 rounded-lg text-xs md:text-sm bg-white focus:ring-2 focus:ring-blue-500 outline-none"></div>
                <div><label class="block text-[9px] md:text-[10px] font-bold text-slate-500 mb-1 uppercase">Contraseña</label><input type="password" name="password" required class="w-full px-3 py-2 border border-slate-300 rounded-lg text-xs md:text-sm bg-white focus:ring-2 focus:ring-blue-500 outline-none"></div>
                <div>
                    <label class="block text-[9px] md:text-[10px] font-bold text-slate-500 mb-1 uppercase">Rol</label>
                    <select name="rol" required class="w-full px-3 py-2 border border-slate-300 rounded-lg text-xs md:text-sm bg-white font-semibold focus:ring-2 focus:ring-blue-500 outline-none"><option value="REGISTRAR">REGISTRAR (Carga)</option><option value="CONSULTAR">CONSULTAR (Lectura)</option><option value="ADMIN">ADMIN (Total)</option></select>
                </div>
                <button type="submit" class="bg-blue-600 hover:bg-blue-700 text-white font-bold py-2.5 px-4 rounded-lg text-xs md:text-sm shadow-md transition-colors w-full mt-2 md:mt-0">Crear Usuario</button>
            </div>
        </form>
        {% else %}
        <div class="mb-6 md:mb-8 p-4 bg-blue-50 border border-blue-100 rounded-xl text-blue-800 text-xs md:text-sm font-bold flex items-center gap-2">
            <i data-lucide="info" class="w-4 h-4 flex-shrink-0"></i> Solo el superadministrador principal (admin) puede crear o eliminar cuentas de usuario.
        </div>
        {% endif %}
        <div class="overflow-x-auto hide-scrollbar">
            <table class="w-full text-left border border-slate-200 rounded-xl overflow-hidden min-w-[600px]">
                <thead class="bg-slate-100 border-b border-slate-200 text-[10px] md:text-xs text-slate-500 uppercase tracking-wider"><tr><th class="px-4 md:px-5 py-3 font-bold">Usuario</th><th class="px-4 md:px-5 py-3 font-bold">Rol Asignado</th><th class="px-4 md:px-5 py-3 font-bold">Módulos</th>{% if session.get('username') == 'admin' %}<th class="px-4 md:px-5 py-3 text-right font-bold">Acción</th>{% endif %}</tr></thead>
                <tbody class="divide-y divide-slate-100 text-xs md:text-sm">
                    {% for u in usuarios %}
                    <tr class="hover:bg-slate-50 transition-colors">
                        <td class="px-4 md:px-5 py-3 md:py-4 font-bold text-slate-700 flex items-center gap-2"><i data-lucide="user" class="w-3.5 h-3.5 md:w-4 md:h-4 text-slate-400"></i> {{ u.username }}</td>
                        <td class="px-4 md:px-5 py-3 md:py-4"><span class="px-2 md:px-3 py-1.5 rounded-full text-[9px] md:text-[10px] font-bold tracking-wider uppercase {% if u.rol == 'ADMIN' %}bg-purple-100 text-purple-700{% elif u.rol == 'REGISTRAR' %}bg-blue-100 text-blue-700{% else %}bg-slate-200 text-slate-700{% endif %}">{{ u.rol }}</span></td>
                        <td class="px-4 md:px-5 py-3 md:py-4 text-slate-600 font-semibold text-[11px] md:text-xs"><i data-lucide="layout-grid" class="inline w-3 h-3 mr-1 text-slate-400"></i> {{ u.acceso_modulos }}</td>
                        {% if session.get('username') == 'admin' %}
                        <td class="px-4 md:px-5 py-3 md:py-4 text-right">
                            {% if u.username != 'admin' %}<form method="POST" action="{{ url_for('eliminar_usuario', username=u.username) }}" onsubmit="return confirm('¿Eliminar usuario?');"><input type="hidden" name="csrf_token" value="{{ csrf_token }}"/><button type="submit" class="text-rose-500 hover:text-white px-3 py-1.5 bg-white hover:bg-rose-500 rounded-lg border border-rose-200 text-[10px] md:text-xs font-bold transition-colors shadow-sm">Eliminar</button></form>{% else %}<span class="text-[10px] md:text-xs text-slate-400 italic font-medium bg-slate-100 px-2 py-1 rounded">Fijo</span>{% endif %}
                        </td>
                        {% endif %}
                    </tr>
                    {% endfor %}
                </tbody>
            </table>
        </div>
    </div>
</div>
{% endblock %}
    """,
    'parroquias.html': """
{% extends 'base.html' %}
{% block content %}
<div class="max-w-2xl mx-auto bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
    <div class="px-4 md:px-6 py-4 md:py-5 bg-slate-50 border-b border-slate-100"><h2 class="text-base md:text-lg font-extrabold text-[#0f172a] flex items-center gap-2"><i data-lucide="map-pin" class="text-blue-600 w-5 h-5"></i> Gestión de Ubicaciones</h2></div>
    <div class="p-4 md:p-6 border-b border-slate-100">
        <form method="POST" class="flex flex-col sm:flex-row gap-3 items-end">
            <input type="hidden" name="csrf_token" value="{{ csrf_token }}"/>
            <div class="flex-1 w-full"><label class="block text-[10px] md:text-xs font-bold text-slate-500 uppercase mb-1.5">Añadir Nueva Ubicación</label><input type="text" name="nombre" placeholder="Ej. ESTADO LA GUAIRA" required class="w-full px-3.5 py-2.5 border border-slate-300 rounded-xl text-xs md:text-sm uppercase bg-white focus:ring-2 focus:ring-blue-500 outline-none"></div>
            <button type="submit" class="w-full sm:w-auto bg-blue-600 hover:bg-blue-700 text-white font-bold py-2.5 px-6 rounded-xl text-xs md:text-sm shadow-md transition-colors">Guardar</button>
        </form>
    </div>
    <div class="overflow-x-auto w-full hide-scrollbar">
        <table class="w-full text-left min-w-[300px]">
            <thead class="bg-slate-50 border-b border-slate-200 text-[10px] md:text-[11px] text-slate-500 uppercase tracking-wider"><tr><th class="px-4 md:px-6 py-3 font-bold">Nombre</th><th class="px-4 md:px-6 py-3 text-right font-bold">Acción</th></tr></thead>
            <tbody class="divide-y divide-slate-100 text-xs md:text-sm">
                {% for p in parroquias %}
                <tr class="hover:bg-slate-50 transition-colors"><td class="px-4 md:px-6 py-3 md:py-4 font-bold text-slate-700">{{ p.nombre }}</td>
                    <td class="px-4 md:px-6 py-3 md:py-4 text-right">
                        <form method="POST" action="{{ url_for('eliminar_parroquia', id=p.id) }}"><input type="hidden" name="csrf_token" value="{{ csrf_token }}"/><button type="submit" class="text-rose-500 hover:text-white font-bold text-[10px] md:text-xs px-3 py-1.5 bg-white rounded-lg hover:bg-rose-500 border border-rose-200 shadow-sm transition-colors">Eliminar</button></form>
                    </td>
                </tr>
                {% endfor %}
            </tbody>
        </table>
    </div>
</div>
{% endblock %}
    """
}

app.jinja_loader = DictLoader(TEMPLATES)

# ==============================================================================
# INICIALIZACIÓN DE DATOS Y AUTO-MIGRACIÓN
# ==============================================================================
def inicializar_bd():
    with app.app_context():
        db.create_all()
        
        # Auto-Migración (Permite agregar columnas a SQLite sin perder datos)
        try:
            with db.engine.connect() as conn:
                # Revisar y Agregar a JefeFamilia
                for col_name, col_type in [
                    ("requerimiento", "TEXT DEFAULT 'NINGUNO'"), 
                    ("discapacidad", "VARCHAR(100) DEFAULT 'NINGUNA'"), 
                    ("patologia", "VARCHAR(100) DEFAULT 'NINGUNA'"), 
                    ("foto", "VARCHAR(200)"), 
                    ("foto_vivienda", "VARCHAR(200)"), 
                    ("es_embarazada", "INTEGER DEFAULT 0"),
                    ("observacion", "TEXT DEFAULT 'NINGUNA'"),
                    ("lugar_remision", "VARCHAR(200) DEFAULT ''")
                ]:
                    try: conn.execute(db.text(f"SELECT {col_name} FROM jefe_familia LIMIT 1"))
                    except: conn.execute(db.text(f"ALTER TABLE jefe_familia ADD COLUMN {col_name} {col_type}"))
                        
                # Revisar y Agregar a NucleoFamiliar
                for col_name, col_type in [("discapacidad", "VARCHAR(100) DEFAULT 'NINGUNA'"), ("patologia", "VARCHAR(100) DEFAULT 'NINGUNA'"), ("es_embarazada", "INTEGER DEFAULT 0"), ("ci", "VARCHAR(20) DEFAULT ''")]:
                    try: conn.execute(db.text(f"SELECT {col_name} FROM nucleo_familiar LIMIT 1"))
                    except: conn.execute(db.text(f"ALTER TABLE nucleo_familiar ADD COLUMN {col_name} {col_type}"))
                
                # Revisar y Agregar Accesos a Usuario
                try: conn.execute(db.text("SELECT acceso_modulos FROM usuario LIMIT 1"))
                except: conn.execute(db.text("ALTER TABLE usuario ADD COLUMN acceso_modulos VARCHAR(50) DEFAULT 'TOTAL'"))

                conn.commit()
        except Exception as e:
            print(f"Nota de migración: {e}")
            
        # Configuraciones por defecto TV
        try:
            for c_clave in ['tv_familias_egresadas', 'tv_poblacion_egresada']:
                if not Configuracion.query.filter_by(clave=c_clave).first():
                    db.session.add(Configuracion(clave=c_clave, valor=''))
            db.session.commit()
        except Exception as e:
            print(f"Error inicializando configuracion: {e}")

        # Usuarios por defecto
        if not Usuario.query.filter_by(username='admin').first():
            db.session.add(Usuario(username='admin', password_hash=generate_password_hash('admin123'), rol='ADMIN', acceso_modulos='TOTAL'))
        
        # Parroquias por defecto
        for p_nombre in ['SAN BERNARDINO', 'EL RECREO', 'ALTAGRACIA', 'SUCRE', 'CATEDRAL', 'ESTADO LA GUAIRA', 'ESTADO MIRANDA']:
            if not Parroquia.query.filter_by(nombre=p_nombre).first():
                db.session.add(Parroquia(nombre=p_nombre))
        db.session.commit()

# ==============================================================================
# CONTROLADORES / RUTAS
# ==============================================================================
@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form.get('username', '').strip().lower()
        password = request.form.get('password', '')
        user = Usuario.query.filter_by(username=username).first()
        if user and check_password_hash(user.password_hash, password):
            session.permanent = True
            session['username'] = user.username
            session['rol'] = user.rol
            session['acceso_modulos'] = user.acceso_modulos or 'TOTAL'
            
            return redirect(url_for('index'))
        flash('Credenciales incorrectas.', 'error')
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

# Uso de Caché de Reportes
def calcular_estadisticas():
    now = datetime.now()
    if 'stats' in CACHE_DICT and CACHE_DICT.get('stats_expire', now) > now:
        return CACHE_DICT['stats']['data'], CACHE_DICT['stats']['parroquias']

    jefes = JefeFamilia.query.all()
    stats = {
        'familias': len(jefes), 'poblacion': 0, 'criticos': 0, 'embarazadas': 0,
        'ninos': 0, 'ninos_m': 0, 'ninos_f': 0, 
        'ninos_m_0_3': 0, 'ninos_m_4_6': 0, 'ninos_m_7_12': 0,
        'ninos_f_0_3': 0, 'ninos_f_4_6': 0, 'ninos_f_7_12': 0,
        'adol': 0, 'adol_m': 0, 'adol_f': 0,
        'adultos': 0, 'adultos_m': 0, 'adultos_f': 0,
        'muj_may': 0, 'hom_may': 0, 'discapacitados': 0,
        'remitidos': 0, 'poblacion_remitida': 0
    }
    parroquia_counts = {}
    
    def procesar_demografia(edad, genero):
        if edad is None or edad < 0:
            stats['adultos'] += 1
            if genero == 'F': stats['adultos_f'] += 1
            else: stats['adultos_m'] += 1
            return

        if 0 <= edad <= 12:
            stats['ninos'] += 1
            if genero == 'F': 
                stats['ninos_f'] += 1
                if 0 <= edad <= 3: stats['ninos_f_0_3'] += 1
                elif 4 <= edad <= 6: stats['ninos_f_4_6'] += 1
                else: stats['ninos_f_7_12'] += 1
            else: 
                stats['ninos_m'] += 1
                if 0 <= edad <= 3: stats['ninos_m_0_3'] += 1
                elif 4 <= edad <= 6: stats['ninos_m_4_6'] += 1
                else: stats['ninos_m_7_12'] += 1

        elif 13 <= edad <= 17:
            stats['adol'] += 1
            if genero == 'F': stats['adol_f'] += 1
            else: stats['adol_m'] += 1

        else:
            if genero == 'F' and edad >= 55:
                stats['muj_may'] += 1
            elif genero == 'M' and edad >= 60:
                stats['hom_may'] += 1
            else:
                stats['adultos'] += 1
                if genero == 'F': stats['adultos_f'] += 1
                else: stats['adultos_m'] += 1

    for jefe in jefes:
        poblacion_nucleo = 1 + len(jefe.cargas)
        stats['poblacion'] += poblacion_nucleo
        parroquia_counts[jefe.parroquia] = parroquia_counts.get(jefe.parroquia, 0) + poblacion_nucleo

        sit = (jefe.situacion or "").upper()
        if any(r in sit for r in ['INHABITABLE', 'SIN CASA', 'ALTO RIESGO', 'DERRUMBE']):
            stats['criticos'] += 1
        
        if jefe.es_embarazada: stats['embarazadas'] += 1
        if jefe.discapacidad and jefe.discapacidad.upper() not in ['NINGUNA', 'NINGUNO', 'NO', '']: stats['discapacitados'] += 1

        if jefe.lugar_remision and jefe.lugar_remision.strip() != '':
            stats['remitidos'] += 1
            stats['poblacion_remitida'] += poblacion_nucleo

        procesar_demografia(jefe.edad, jefe.genero)
        
        for m in jefe.cargas:
            if m.es_embarazada: stats['embarazadas'] += 1
            if m.discapacidad and m.discapacidad.upper() not in ['NINGUNA', 'NINGUNO', 'NO', '']: stats['discapacitados'] += 1
            procesar_demografia(m.edad, m.genero)
            
    CACHE_DICT['stats'] = {'data': stats, 'parroquias': parroquia_counts}
    CACHE_DICT['stats_expire'] = now + timedelta(hours=1)

    return stats, parroquia_counts

@app.route('/')
@login_required
def index():
    # PAGINACIÓN Y BÚSQUEDA BACKEND CON FILTRO INTELIGENTE
    q = request.args.get('q', '').strip()
    parroquia_filtro = request.args.get('parroquia', '')
    grupo_filtro = request.args.get('grupo', '')
    page = request.args.get('page', 1, type=int)
    
    query = JefeFamilia.query
    
    # Búsqueda por texto (Nombre o CI)
    if q:
        query = query.filter(or_(
            JefeFamilia.nombre.ilike(f'%{q}%'),
            JefeFamilia.ci.ilike(f'%{q}%')
        ))
        
    # Filtro por Parroquia/Sector
    if parroquia_filtro:
        query = query.filter(JefeFamilia.parroquia == parroquia_filtro)
        
    # Filtro Inteligente de Grupos Poblacionales
    if grupo_filtro:
        cond_jefe = None
        cond_carga = None
        
        if grupo_filtro == 'ninos':
            cond_jefe = JefeFamilia.edad <= 12
            cond_carga = NucleoFamiliar.edad <= 12
        elif grupo_filtro == 'ninas':
            cond_jefe = and_(JefeFamilia.edad <= 12, JefeFamilia.genero == 'F')
            cond_carga = and_(NucleoFamiliar.edad <= 12, NucleoFamiliar.genero == 'F')
        elif grupo_filtro == 'adol':
            cond_jefe = and_(JefeFamilia.edad >= 13, JefeFamilia.edad <= 17)
            cond_carga = and_(NucleoFamiliar.edad >= 13, NucleoFamiliar.edad <= 17)
        elif grupo_filtro == 'mayores':
            cond_jefe = or_(
                and_(JefeFamilia.genero == 'M', JefeFamilia.edad >= 60),
                and_(JefeFamilia.genero == 'F', JefeFamilia.edad >= 55)
            )
            cond_carga = or_(
                and_(NucleoFamiliar.genero == 'M', NucleoFamiliar.edad >= 60),
                and_(NucleoFamiliar.genero == 'F', NucleoFamiliar.edad >= 55)
            )
        elif grupo_filtro == 'embarazadas':
            cond_jefe = JefeFamilia.es_embarazada == 1
            cond_carga = NucleoFamiliar.es_embarazada == 1
        elif grupo_filtro == 'discapacidad':
            cond_jefe = and_(JefeFamilia.discapacidad.isnot(None), JefeFamilia.discapacidad != '', JefeFamilia.discapacidad != 'NINGUNA')
            cond_carga = and_(NucleoFamiliar.discapacidad.isnot(None), NucleoFamiliar.discapacidad != '', NucleoFamiliar.discapacidad != 'NINGUNA')
            
        if cond_jefe is not None and cond_carga is not None:
            # Trae al Jefe si él mismo cumple la condición O si al menos una de sus cargas la cumple
            query = query.filter(or_(
                cond_jefe,
                JefeFamilia.cargas.any(cond_carga)
            ))
            
    jefes_paginados = query.order_by(JefeFamilia.fecha_registro.desc()).paginate(page=page, per_page=50, error_out=False)
    
    # Agrupamos solo los registros de la PÁGINA ACTUAL
    agrupacion = {}
    for jefe in jefes_paginados.items:
        agrupacion.setdefault(jefe.parroquia, []).append(jefe)

    stats, _ = calcular_estadisticas()
    
    parroquias_lista = Parroquia.query.order_by(Parroquia.nombre).all()
    
    return render_template('inicio.html', agrupacion=agrupacion, stats=stats, pagination=jefes_paginados, parroquias=parroquias_lista)

@app.route('/dashboard')
@login_required
def dashboard():
    stats, parroquia_counts = calcular_estadisticas()
    
    # Evaluar Ajustes TV Manuales
    c_f_egr = Configuracion.query.filter_by(clave='tv_familias_egresadas').first()
    c_p_egr = Configuracion.query.filter_by(clave='tv_poblacion_egresada').first()

    manual_f_egr = c_f_egr.valor if c_f_egr and c_f_egr.valor.strip() != '' else stats['remitidos']
    manual_p_egr = c_p_egr.valor if c_p_egr and c_p_egr.valor.strip() != '' else stats['poblacion_remitida']

    age_labels = ['Niños (0-12)', 'Adolescentes (13-17)', 'Adultos (18-54)', 'Mayores (+55)']
    age_data = [stats['ninos'], stats['adol'], stats['adultos'], stats['muj_may'] + stats['hom_may']]
    
    sorted_parishes = sorted(parroquia_counts.items(), key=lambda item: item[1], reverse=True)
    parish_labels = [p[0] for p in sorted_parishes]
    parish_data = [p[1] for p in sorted_parishes]

    return render_template('dashboard.html', 
                           stats=stats, 
                           manual_f_egr=manual_f_egr,
                           manual_p_egr=manual_p_egr,
                           age_labels=age_labels, age_data=age_data,
                           parish_labels=parish_labels, parish_data=parish_data)

@app.route('/config_tv', methods=['GET', 'POST'])
@role_required('ADMIN')
def config_tv():
    if request.method == 'POST':
        f_egr = request.form.get('tv_familias_egresadas', '').strip()
        p_egr = request.form.get('tv_poblacion_egresada', '').strip()

        def update_conf(k, v):
            c = Configuracion.query.filter_by(clave=k).first()
            if c:
                c.valor = v
            else:
                db.session.add(Configuracion(clave=k, valor=v))

        update_conf('tv_familias_egresadas', f_egr)
        update_conf('tv_poblacion_egresada', p_egr)

        db.session.commit()
        flash('Configuración de valores en pantalla TV guardada.', 'success')
        return redirect(url_for('config_tv'))

    c_f_egr = Configuracion.query.filter_by(clave='tv_familias_egresadas').first()
    c_p_egr = Configuracion.query.filter_by(clave='tv_poblacion_egresada').first()
    
    stats, _ = calcular_estadisticas()

    return render_template('config_tv.html',
        v_f_egr=c_f_egr.valor if c_f_egr else '',
        v_p_egr=c_p_egr.valor if c_p_egr else '',
        stats=stats
    )

@app.route('/imprimir_ficha/<int:id>')
@login_required
def imprimir_ficha(id):
    jefe = JefeFamilia.query.get_or_404(id)
    return render_template('imprimir_ficha.html', jefe=jefe)

# NUEVA RUTA PARA CARNET DE TRABAJADORES
@app.route('/imprimir_carnet/<tipo>/<int:id>')
@login_required
def imprimir_carnet(tipo, id):
    nro_familia = request.args.get('nro', '')
    
    if tipo == 'jefe':
        persona = JefeFamilia.query.get_or_404(id)
        jefe = persona
        rol = "REPRESENTANTE / JEFE DE FAMILIA"
    else:
        persona = NucleoFamiliar.query.get_or_404(id)
        jefe = JefeFamilia.query.get(persona.jefe_id)
        rol = "MIEMBRO DE NÚCLEO FAMILIAR"
        
    return render_template('carnet.html', persona=persona, jefe=jefe, rol=rol, nro_familia=nro_familia, tipo=tipo)

@app.route('/editar_jefe/<int:id>', methods=['GET', 'POST'])
@role_required('ADMIN')
def editar_jefe(id):
    jefe = JefeFamilia.query.get_or_404(id)
    if request.method == 'POST':
        if 'foto' in request.files and request.files['foto'].filename != '':
            foto_name = comprimir_y_guardar_foto(request.files['foto'], app.config['UPLOAD_FOLDER'])
            if foto_name:
                jefe.foto = foto_name
        
        if 'foto_vivienda' in request.files and request.files['foto_vivienda'].filename != '':
            foto_vivienda_name = comprimir_y_guardar_foto(request.files['foto_vivienda'], app.config['UPLOAD_FOLDER'])
            if foto_vivienda_name:
                jefe.foto_vivienda = foto_vivienda_name
        
        jefe.nombre = request.form.get('nombre', jefe.nombre).strip().upper()
        jefe.ci = request.form.get('ci', jefe.ci).strip().upper()
        edad_str = str(request.form.get('edad', '0')).strip()
        jefe.edad = int(edad_str) if edad_str.lstrip('-').isdigit() else jefe.edad
        jefe.genero = request.form.get('genero', jefe.genero).upper()
        jefe.telefono = request.form.get('telefono', jefe.telefono).strip()
        jefe.situacion = request.form.get('situacion', jefe.situacion).strip()
        jefe.parroquia = request.form.get('parroquia', jefe.parroquia).strip()
        
        jefe.requerimiento = request.form.get('requerimiento', jefe.requerimiento).strip().upper() or 'NINGUNO'
        jefe.observacion = request.form.get('observacion', jefe.observacion).strip().upper() or 'NINGUNA'
        jefe.discapacidad = request.form.get('discapacidad', jefe.discapacidad).strip().upper() or 'NINGUNA'
        jefe.patologia = request.form.get('patologia', jefe.patologia).strip().upper() or 'NINGUNA'
        
        jefe.es_embarazada = 1 if request.form.get('es_embarazada') else 0
        
        es_remitido = request.form.get('es_remitido')
        jefe.lugar_remision = request.form.get('lugar_remision', '').strip().upper() if es_remitido else ''
        
        carga_ids = request.form.getlist('carga_id[]')
        carga_nombres = request.form.getlist('carga_nombre[]')
        carga_ci = request.form.getlist('carga_ci[]')
        carga_parentescos = request.form.getlist('carga_parentesco[]')
        carga_edades = request.form.getlist('carga_edad[]')
        carga_generos = request.form.getlist('carga_genero[]')
        carga_discapacidades = request.form.getlist('carga_discapacidad[]')
        carga_patologias = request.form.getlist('carga_patologia[]')
        
        for i, cid_str in enumerate(carga_ids):
            if i < len(carga_nombres):
                cid = int(cid_str)
                carga = NucleoFamiliar.query.get(cid)
                if carga and carga.jefe_id == jefe.id:
                    carga.nombre = carga_nombres[i].strip().upper()
                    carga.ci = carga_ci[i].strip().upper() if i < len(carga_ci) else ''
                    carga.parentesco = carga_parentescos[i].strip().upper()
                    c_edad_str = str(carga_edades[i]).strip()
                    carga.edad = int(c_edad_str) if c_edad_str.lstrip('-').isdigit() else carga.edad
                    carga.genero = carga_generos[i].upper()
                    carga.discapacidad = carga_discapacidades[i].strip().upper() or 'NINGUNA'
                    carga.patologia = carga_patologias[i].strip().upper() or 'NINGUNA'
                    carga.es_embarazada = 1 if request.form.get(f'carga_embarazo_{carga.id}') else 0
        
        n_nombres = request.form.getlist('m_nombre[]')
        n_ci = request.form.getlist('m_ci[]')
        n_parentescos = request.form.getlist('m_parentesco[]')
        n_edades = request.form.getlist('m_edad[]')
        n_generos = request.form.getlist('m_genero[]')
        n_discapacidades = request.form.getlist('m_discapacidad[]')
        n_patologias = request.form.getlist('m_patologia[]')
        n_embarazos = request.form.getlist('m_embarazo[]')
        
        for i in range(len(n_nombres)):
            if n_nombres[i].strip():
                n_edad_str = str(n_edades[i]).strip()
                nueva_carga = NucleoFamiliar(
                    jefe_id=jefe.id, nombre=n_nombres[i].strip().upper(), parentesco=n_parentescos[i].strip().upper(),
                    ci=n_ci[i].strip().upper() if i < len(n_ci) else '',
                    edad=int(n_edad_str) if n_edad_str.lstrip('-').isdigit() else 0, genero=n_generos[i].upper(),
                    discapacidad=n_discapacidades[i].strip().upper() or 'NINGUNA', patologia=n_patologias[i].strip().upper() or 'NINGUNA',
                    es_embarazada=1 if str(i) in n_embarazos else 0
                )
                db.session.add(nueva_carga)
        
        db.session.commit()
        clear_stats_cache()
        registrar_auditoria(session['username'], 'EDICIÓN', f'Se editó la familia de {jefe.nombre}')
        flash('Registro actualizado correctamente.', 'success')
        return redirect(url_for('index'))
        
    parroquias = Parroquia.query.order_by(Parroquia.nombre).all()
    return render_template('editar.html', jefe=jefe, parroquias=parroquias)

@app.route('/eliminar_carga/<int:id>', methods=['POST'])
@role_required('ADMIN')
def eliminar_carga(id):
    carga = NucleoFamiliar.query.get_or_404(id)
    nombre_carga = carga.nombre
    db.session.delete(carga)
    db.session.commit()
    clear_stats_cache()
    registrar_auditoria(session['username'], 'ELIMINACIÓN', f'Se eliminó carga familiar específica: {nombre_carga}')
    return '', 204

@app.route('/eliminar_jefe/<int:id>', methods=['POST'])
@role_required('ADMIN')
def eliminar_jefe(id):
    jefe = JefeFamilia.query.get_or_404(id)
    nombre = jefe.nombre
    db.session.delete(jefe)
    db.session.commit()
    clear_stats_cache()
    registrar_auditoria(session['username'], 'ELIMINACIÓN', f'Se eliminó familia: {nombre}')
    flash('Registro eliminado exitosamente.', 'success')
    return redirect(url_for('index'))

@app.route('/actualizar_situacion/<int:id>', methods=['POST'])
@role_required('ADMIN', 'REGISTRAR')
def actualizar_situacion(id):
    jefe = JefeFamilia.query.get_or_404(id)
    nueva = request.form.get('nueva_situacion')
    if nueva:
        jefe.situacion = nueva
        db.session.commit()
        clear_stats_cache()
        flash('Estatus actualizado.', 'success')
    return redirect(url_for('index', page=request.args.get('page', 1)))

@app.route('/registrar', methods=['GET', 'POST'])
@role_required('ADMIN', 'REGISTRAR')
def registrar():
    if request.method == 'POST':
        foto_name = None
        if 'foto' in request.files:
            foto_name = comprimir_y_guardar_foto(request.files['foto'], app.config['UPLOAD_FOLDER'])
        
        foto_vivienda_name = None
        if 'foto_vivienda' in request.files:
            foto_vivienda_name = comprimir_y_guardar_foto(request.files['foto_vivienda'], app.config['UPLOAD_FOLDER'])
            
        edad_jefe_str = str(request.form.get('edad', '0')).strip()
        edad_jefe = int(edad_jefe_str) if edad_jefe_str.lstrip('-').isdigit() else 0
        
        es_remitido = request.form.get('es_remitido')
        lugar_remision = request.form.get('lugar_remision', '').strip().upper() if es_remitido else ''
            
        jefe = JefeFamilia(
            nombre=request.form.get('nombre').strip().upper(), ci=request.form.get('ci').strip(),
            edad=edad_jefe, genero=request.form.get('genero').upper(),
            telefono=request.form.get('telefono').strip(), situacion=request.form.get('situacion'),
            parroquia=request.form.get('parroquia'), 
            
            requerimiento=request.form.get('requerimiento', '').strip().upper() or 'NINGUNO',
            observacion=request.form.get('observacion', '').strip().upper() or 'NINGUNA',
            lugar_remision=lugar_remision,
            discapacidad=request.form.get('discapacidad', '').strip().upper() or 'NINGUNA', 
            patologia=request.form.get('patologia', '').strip().upper() or 'NINGUNA',
            
            es_embarazada=1 if request.form.get('embarazada_jefe') else 0, 
            foto=foto_name,
            foto_vivienda=foto_vivienda_name,
            fecha_registro=datetime.now().strftime('%Y-%m-%d %H:%M:%S'), usuario_registra=session['username']
        )
        db.session.add(jefe)
        db.session.flush()
        
        nombres = request.form.getlist('m_nombre[]')
        ci_cargas = request.form.getlist('m_ci[]')
        parentescos = request.form.getlist('m_parentesco[]')
        edades = request.form.getlist('m_edad[]')
        generos = request.form.getlist('m_genero[]')
        discapacidades = request.form.getlist('m_discapacidad[]')
        patologias = request.form.getlist('m_patologia[]')
        embarazos = request.form.getlist('m_embarazo[]')
        
        for i in range(len(nombres)):
            if nombres[i].strip():
                n_edad_str = str(edades[i]).strip() if i < len(edades) else "0"
                carga_edad = int(n_edad_str) if n_edad_str.lstrip('-').isdigit() else 0
                
                carga = NucleoFamiliar(
                    jefe_id=jefe.id, nombre=nombres[i].strip().upper(), parentesco=parentescos[i].strip().upper(),
                    ci=ci_cargas[i].strip().upper() if i < len(ci_cargas) else '',
                    edad=carga_edad, genero=generos[i].upper(),
                    discapacidad=discapacidades[i].strip().upper() or 'NINGUNA', patologia=patologias[i].strip().upper() or 'NINGUNA',
                    es_embarazada=1 if str(i) in embarazos else 0
                )
                db.session.add(carga)
                
        db.session.commit()
        clear_stats_cache()
        registrar_auditoria(session['username'], 'NUEVO REGISTRO', f'Jefe: {jefe.nombre}')
        flash('Registro guardado correctamente.', 'success')
        return redirect(url_for('index'))
        
    parroquias = Parroquia.query.order_by(Parroquia.nombre).all()
    return render_template('registrar.html', parroquias=parroquias)

# ==============================================================================
# MOTOR DE CARGA OPTIMIZADO - ACTUALIZACIÓN INTELIGENTE
# ==============================================================================
def is_empty_or_default(v):
    if v is None: return True
    s = str(v).strip().upper()
    return s in ['', 'NINGUNA', 'NINGUNO', 'S/T', 'S/C', 'POR CLASIFICAR', 'REFUGIO / CANCHA', 'NAN', 'NONE']

@app.route('/cargar_datos', methods=['GET', 'POST'])
@role_required('ADMIN')
def cargar_datos():
    if request.method == 'POST':
        if 'archivo' not in request.files:
            flash('No se seleccionó ningún archivo.', 'error')
            return redirect(request.url)
        
        file = request.files['archivo']
        if file.filename == '':
            flash('Nombre de archivo vacío.', 'error')
            return redirect(request.url)
            
        if file and allowed_file(file.filename):
            try:
                file.filename = secure_filename(file.filename) 
                
                df_dict = leer_excel_dinamico(file)
                if not df_dict:
                    flash('No se pudo leer el archivo. Asegúrate que tenga los encabezados básicos.', 'error')
                    return redirect(url_for('cargar_datos'))

                jefes_agregados, jefes_actualizados, cargas_agregadas, cargas_actualizadas = 0, 0, 0, 0
                
                def val(row, possibilities):
                    for p in possibilities:
                        for col in row.index:
                            if p in str(col).strip().upper():
                                v = str(row[col]).strip()
                                if v.upper() not in ['NAN', 'NAT', 'NONE', '']:
                                    return v
                    return ""

                consolidated_map = {}
                
                for sheet_name, df_sheet in df_dict.items():
                    current_jefe_key = None
                    
                    for idx, row in df_sheet.iterrows():
                        nombre_raw = val(row, ["NOMBRES Y APELLIDOS", "NOMBRE Y APELLIDO", "NOMBRE"]).upper().strip()
                        if not nombre_raw or nombre_raw in ['NAN', 'NONE']:
                            continue 
                            
                        ci = val(row, ["C.I", "CI", "CEDULA", "C.I."]).replace('.0','').replace('.','').strip()
                        if ci in ['NAN', 'NONE', 'S/C', '']: ci = ""

                        nro = val(row, ["FAMILIA N°", "FAMILIA N", "N°", "NRO.", "NRO"])
                        edad = val(row, ["EDAD"])
                        parentesco = val(row, ["PARENTESCO"]).upper()
                        
                        is_head = False
                        if "JEFE" in parentesco or "JEFA" in parentesco:
                            is_head = True
                        elif nro and str(nro).upper() not in ['NAN', 'NONE', '']:
                            is_head = True
                        elif current_jefe_key is None:
                            is_head = True
                            
                        nombre_limpio = nombre_raw.replace("JEFE DE FAMILIA", "").replace("JEFA DE FAMILIA", "").replace("-", "").replace(":", "").strip()

                        discap_raw = val(row, ["CON CONDICIÓN", "CON CONDICION", "DISCAPACIDAD", "POSEE DISCAPACIDAD"])
                        patol_raw = val(row, ["PATOLOGIA", "PATOLOGÍA"])
                        preg_raw = val(row, ["MUJER EMBARAZADA", "EMBARAZADA"])
                        req_raw = val(row, ["REQUERIMIENTO", "REQUERIMIENTO "])
                        
                        discap = "Sí" if str(discap_raw).upper() in ['1', '1.0', 'X', 'SI', 'SÍ'] else (discap_raw if discap_raw else "")
                        patol = patol_raw if patol_raw else ""
                        preg = 1 if str(preg_raw).upper() in ['1', '1.0', 'X', 'SI', 'SÍ'] else 0

                        if is_head:
                            telf = val(row, ["TELEFONO", "TÉLEFONO", "TELF"])
                            parr = val(row, ["PARROQUIA", "PARROQUIA "]).upper()
                            sit = val(row, ["SITUACIÓN ACTUAL", "SITUACION ACTUAL", "SITUACION"]).title()
                            obs = val(row, ["OBSERVACIÓN", "OBSERVACION", "OBSERVACIONES"])
                            lugar_rem = val(row, ["REMITIDO A", "INSTITUCION", "OBSERVACION DE LOS QUE FUERON REMITIDOS"])

                            current_jefe = {
                                "nombre": nombre_limpio,
                                "ci": ci,
                                "edad": edad,
                                "telf": telf,
                                "parroquia": parr if parr else "POR CLASIFICAR",
                                "situacion": sit if sit else "Refugio / Cancha",
                                "observacion": obs,
                                "lugar_remision": lugar_rem,
                                "discapacidad": discap,
                                "patologia": patol,
                                "es_embarazada": preg,
                                "requerimiento": req_raw if req_raw else "NINGUNO",
                                "nucleoFamiliar": [] 
                            }
                            current_jefe_key = ci if ci else f"{nombre_limpio}_{idx}"
                            consolidated_map[current_jefe_key] = current_jefe

                        else:
                            if current_jefe_key is not None and current_jefe_key in consolidated_map:
                                consolidated_map[current_jefe_key]["nucleoFamiliar"].append({
                                    "nombre": nombre_limpio,
                                    "ci": ci,
                                    "edad": edad,
                                    "parentesco": parentesco if parentesco else "FAMILIAR",
                                    "discapacidad": discap,
                                    "embarazada": preg,
                                    "patologia": patol,
                                    "requerimiento": req_raw if req_raw else "NINGUNO"
                                })

                for c_data in consolidated_map.values():
                    if c_data["parroquia"] != 'POR CLASIFICAR':
                        if not Parroquia.query.filter_by(nombre=c_data["parroquia"]).first():
                            db.session.add(Parroquia(nombre=c_data["parroquia"]))
                            db.session.flush()

                    edad_val = parse_age(c_data["edad"])
                    genero_val = inferir_genero(c_data["nombre"], "JEFE", "")

                    jefe_db = JefeFamilia.query.filter_by(ci=c_data["ci"]).first() if c_data["ci"] else JefeFamilia.query.filter_by(nombre=c_data["nombre"]).first()

                    obs_final = str(c_data.get("observacion", "")).strip().upper()
                    
                    if jefe_db:
                        if is_empty_or_default(jefe_db.situacion) and not is_empty_or_default(c_data["situacion"]):
                            jefe_db.situacion = c_data["situacion"]
                            
                        if is_empty_or_default(jefe_db.telefono) and not is_empty_or_default(c_data["telf"]): 
                            jefe_db.telefono = c_data["telf"]
                            
                        if is_empty_or_default(jefe_db.parroquia) and c_data["parroquia"] != 'POR CLASIFICAR': 
                            jefe_db.parroquia = c_data["parroquia"]
                            
                        if is_empty_or_default(jefe_db.requerimiento) and str(c_data["requerimiento"]).upper() != "NINGUNO": 
                            jefe_db.requerimiento = str(c_data["requerimiento"]).upper()
                            
                        if is_empty_or_default(jefe_db.observacion) and obs_final: 
                            jefe_db.observacion = obs_final
                            
                        if is_empty_or_default(jefe_db.lugar_remision) and c_data["lugar_remision"]: 
                            jefe_db.lugar_remision = c_data["lugar_remision"].upper()
                            
                        if is_empty_or_default(jefe_db.discapacidad) and c_data["discapacidad"]: 
                            jefe_db.discapacidad = str(c_data["discapacidad"]).upper()
                            
                        if is_empty_or_default(jefe_db.patologia) and c_data["patologia"]: 
                            jefe_db.patologia = str(c_data["patologia"]).upper()
                            
                        if not jefe_db.es_embarazada and c_data["es_embarazada"]: 
                            jefe_db.es_embarazada = 1
                            
                        jefes_actualizados += 1
                        jefe_id = jefe_db.id
                    else:
                        nuevo_jefe = JefeFamilia(
                            nombre=c_data["nombre"], ci=c_data["ci"] or 'S/C', edad=edad_val, genero=genero_val,
                            telefono=c_data["telf"], situacion=c_data["situacion"], parroquia=c_data["parroquia"],
                            requerimiento=str(c_data["requerimiento"]).upper() if c_data["requerimiento"] else "NINGUNO", 
                            observacion=obs_final if obs_final else "NINGUNA",
                            lugar_remision=c_data.get("lugar_remision", "").upper(),
                            discapacidad=str(c_data["discapacidad"]).upper() if c_data["discapacidad"] else "NINGUNA",
                            patologia=str(c_data["patologia"]).upper() if c_data["patologia"] else "NINGUNA", 
                            es_embarazada=c_data["es_embarazada"],
                            fecha_registro=datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                            usuario_registra=session['username']
                        )
                        db.session.add(nuevo_jefe)
                        db.session.flush()
                        jefes_agregados += 1
                        jefe_id = nuevo_jefe.id

                    for m in c_data["nucleoFamiliar"]:
                        m_edad = parse_age(m["edad"])
                        m_gen = inferir_genero(m["nombre"], m["parentesco"], "")
                        
                        carga_db = NucleoFamiliar.query.filter_by(jefe_id=jefe_id, nombre=m["nombre"]).first()
                        
                        if carga_db:
                            if m_edad >= 0 and carga_db.edad < 0:
                                carga_db.edad = m_edad
                            if is_empty_or_default(carga_db.discapacidad) and m["discapacidad"]:
                                carga_db.discapacidad = str(m["discapacidad"]).upper()
                            if is_empty_or_default(carga_db.patologia) and m["patologia"]:
                                carga_db.patologia = str(m["patologia"]).upper()
                            if not carga_db.es_embarazada and m["embarazada"]:
                                carga_db.es_embarazada = 1
                            cargas_actualizadas += 1
                        else:
                            nueva_carga = NucleoFamiliar(
                                jefe_id=jefe_id, nombre=m["nombre"], parentesco=m["parentesco"].upper() or 'FAMILIAR',
                                ci=m["ci"] or '',
                                edad=m_edad, genero=m_gen, 
                                discapacidad=str(m["discapacidad"]).upper() if m["discapacidad"] else "NINGUNA", 
                                patologia=str(m["patologia"]).upper() if m["patologia"] else "NINGUNA", 
                                es_embarazada=m["embarazada"]
                            )
                            db.session.add(nueva_carga)
                            cargas_agregadas += 1

                db.session.commit()
                clear_stats_cache()
                registrar_auditoria(session['username'], "CARGA MASIVA EFICIENTE", f"Nuevos Jefes: {jefes_agregados} | Actualizados: {jefes_actualizados}")
                flash(f'Lectura Inteligente Completada de manera Segura. Jefes Actualizados o Nuevos: {jefes_agregados+jefes_actualizados}. Cargas analizadas: {cargas_agregadas + cargas_actualizadas}.', 'success')

            except Exception as e:
                db.session.rollback()
                flash(f'Error procesando el formato del archivo: {str(e)}', 'error')

            return redirect(url_for('index'))
            
    return render_template('cargar.html')


@app.route('/exportar_csv')
@login_required
def exportar_csv():
    jefes = JefeFamilia.query.all()
    si = io.StringIO()
    cw = csv.writer(si)
    cw.writerow(['NOMBRE_JEFE', 'CI_JEFE', 'EDAD', 'GENERO', 'TELEFONO', 'SITUACION', 'PARROQUIA', 'REQUERIMIENTO', 'REMITIDO_A', 'OBSERVACION', 'PERSONA_CON_CONDICION', 'PATOLOGIA', 'ES_EMBARAZADA', 'CANTIDAD_CARGAS'])
    for j in jefes:
        cw.writerow([j.nombre, j.ci, j.edad, j.genero, j.telefono, j.situacion, j.parroquia, j.requerimiento, j.lugar_remision, j.observacion, j.discapacidad, j.patologia, 'SI' if j.es_embarazada else 'NO', len(j.cargas)])
    return Response(si.getvalue(), mimetype='text/csv', headers={"Content-Disposition": "attachment;filename=afectados_export.csv"})

@app.route('/ver_historial')
@role_required('ADMIN')
def ver_historial():
    h = HistorialCambio.query.order_by(HistorialCambio.fecha.desc()).limit(150).all()
    return render_template('historial.html', historial=h)

@app.route('/usuarios', methods=['GET', 'POST'])
@role_required('ADMIN')
def usuarios():
    if request.method == 'POST':
        if session.get('username') != 'admin':
            flash('Acción denegada: Solo el superadministrador puede crear usuarios.', 'error')
            return redirect(url_for('usuarios'))
            
        u = request.form.get('username').strip().lower()
        p = request.form.get('password')
        r = request.form.get('rol')
        acceso = request.form.get('acceso_modulos', 'TOTAL')
        if not Usuario.query.filter_by(username=u).first():
            db.session.add(Usuario(username=u, password_hash=generate_password_hash(p), rol=r, acceso_modulos=acceso))
            db.session.commit()
            flash('Usuario creado con configuración de accesos y rol asignada.', 'success')
        else: flash('El usuario ya existe.', 'error')
        return redirect(url_for('usuarios'))
    return render_template('usuarios.html', usuarios=Usuario.query.all())

@app.route('/eliminar_usuario/<username>', methods=['POST'])
@role_required('ADMIN')
def eliminar_usuario(username):
    if session.get('username') != 'admin':
        flash('Acción denegada: Solo el superadministrador puede eliminar usuarios.', 'error')
        return redirect(url_for('usuarios'))
        
    if username == 'admin': return redirect(url_for('usuarios'))
    u = Usuario.query.filter_by(username=username).first()
    if u: db.session.delete(u); db.session.commit(); flash('Usuario eliminado.', 'success')
    return redirect(url_for('usuarios'))

@app.route('/parroquias', methods=['GET', 'POST'])
@role_required('ADMIN')
def parroquias():
    if request.method == 'POST':
        n = request.form.get('nombre').strip().upper()
        if not Parroquia.query.filter_by(nombre=n).first():
            db.session.add(Parroquia(nombre=n))
            db.session.commit()
            flash('Ubicación añadida.', 'success')
        else: flash('La ubicación ya existe.', 'error')
        return redirect(url_for('parroquias'))
    return render_template('parroquias.html', parroquias=Parroquia.query.order_by(Parroquia.nombre).all())

@app.route('/eliminar_parroquia/<int:id>', methods=['POST'])
@role_required('ADMIN')
def eliminar_parroquia(id):
    p = Parroquia.query.get_or_404(id)
    db.session.delete(p)
    db.session.commit()
    flash('Ubicación eliminada.', 'success')
    return redirect(url_for('parroquias'))

@app.route('/respaldar_bd')
@role_required('ADMIN')
def respaldar_bd():
    registrar_auditoria(session['username'], 'RESPALDO BD', 'Se descargó backup SQLite.')
    db.session.commit()
    return send_file(DB_PATH, as_attachment=True, download_name=f'backup_bd_{datetime.now().strftime("%Y%m%d_%H%M")}.db')

if __name__ == '__main__':
    inicializar_bd()
    app.run(debug=True, host='0.0.0.0', port=5000)