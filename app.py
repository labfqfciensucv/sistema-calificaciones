from flask import Flask, render_template, request, jsonify, redirect, url_for, session
import json
import os
import openpyxl
from functools import wraps
from datetime import datetime

app = Flask(__name__)
app.secret_key = 'tu_clave_secreta_aqui_cambiala'

# Datos de usuarios
USERS = {
    'admin': {
        'password': 'admin123',
        'role': 'admin',
        'name': 'Administrador'
    },
    '28233037': {
        'password': 'agustin2026',
        'role': 'student',
        'name': 'AGUSTIN ANDRADE'
    },
    '27111264': {
        'password': 'mauricio2026',
        'role': 'student',
        'name': 'MAURICIO LEAL'
    },
    '21073829': {
        'password': 'maria2026',
        'role': 'student',
        'name': 'MARIA DANIELA MEDINA'
    },
    '29529906': {
        'password': 'luis2026',
        'role': 'student',
        'name': 'LUIS RIOS'
    },
    '29501138': {
        'password': 'daniela2026',
        'role': 'student',
        'name': 'DANIELA RONDON'
    }
}

# Decorador para login requerido
def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function

# Decorador para admin
def admin_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session or session.get('role') != 'admin':
            return jsonify({'error': 'Acceso denegado'}), 403
        return f(*args, **kwargs)
    return decorated_function

# Función helper para obtener valor seguro de celda
def get_safe_value(cell):
    """Obtiene el valor de una celda manejando diferentes tipos"""
    valor = cell.value
    if valor is None:
        return 0
    if isinstance(valor, str):
        if valor.startswith('='):
            return 0
        try:
            return float(valor)
        except:
            return 0
    try:
        return float(valor)
    except:
        return 0

# ============ FUNCIONES DE CÁLCULO MEJORADAS ============

def calcular_promedio_semanas(notas, semanas_totales=10):
    """
    Calcula el promedio de todas las semanas que tienen notas (ignora ceros)
    
    Args:
        notas: Lista de notas (10 semanas)
        semanas_totales: Número total de semanas (por defecto 10)
    
    Returns:
        float: Promedio de las semanas con notas
    """
    if not notas:
        return 0
    
    # Tomar solo las semanas que tienen nota (valor > 0)
    notas_validas = [n for n in notas[:semanas_totales] if n > 0]
    
    if not notas_validas:
        return 0
    
    return sum(notas_validas) / len(notas_validas)

def contar_semanas_evaluadas(notas, semanas_totales=10):
    """Cuenta cuántas semanas tienen nota (valor > 0)"""
    if not notas:
        return 0
    return len([n for n in notas[:semanas_totales] if n > 0])

def calcular_nota_cuaderno(notas):
    """Calcula la nota del cuaderno (promedio de semanas con notas)"""
    return calcular_promedio_semanas(notas, 10)

def calcular_nota_gestion(notas):
    """Calcula la nota de gestión (promedio de semanas con notas)"""
    return calcular_promedio_semanas(notas, 10)

def calcular_nota_quiz(notas):
    """Calcula la nota de quiz (promedio de semanas con notas)"""
    return calcular_promedio_semanas(notas, 10)

def calcular_nota_informe(informe_notas):
    """Calcula la nota del informe (suma de todos los componentes)"""
    if not informe_notas:
        return 0
    return sum(informe_notas)

def calcular_nota_bloqueA(cuaderno, gestion, quiz, informe, ponderaciones):
    """
    Calcula la nota del Bloque A usando ponderaciones
    
    ponderaciones: dict con las ponderaciones de cada componente
    - cuaderno_pct: porcentaje del cuaderno (ej: 0.20 = 20%)
    - gestion_pct: porcentaje de gestión (ej: 0.15 = 15%)
    - quiz_pct: porcentaje de quiz (ej: 0.25 = 25%)
    - informe_pct: porcentaje de informe (ej: 0.40 = 40%)
    """
    nota_cuaderno = calcular_nota_cuaderno(cuaderno)
    nota_gestion = calcular_nota_gestion(gestion)
    nota_quiz = calcular_nota_quiz(quiz)
    nota_informe = calcular_nota_informe(informe)
    
    # Aplicar ponderaciones (todas en base 20)
    nota_ponderada = (
        nota_cuaderno * ponderaciones.get('cuaderno_pct', 0.20) +
        nota_gestion * ponderaciones.get('gestion_pct', 0.15) +
        nota_quiz * ponderaciones.get('quiz_pct', 0.25) +
        nota_informe * ponderaciones.get('informe_pct', 0.40)
    )
    
    return round(nota_ponderada, 2)

def calcular_nota_final(bloqueA, bloqueB, ponderacion_bloqueA=0.85, ponderacion_bloqueB=0.15):
    """
    Calcula la nota final a partir de Bloque A y Bloque B
    
    Ambos bloques deben estar en base 20
    """
    nota_final = (bloqueA * ponderacion_bloqueA) + (bloqueB * ponderacion_bloqueB)
    return round(nota_final, 2)

# ============ FIN FUNCIONES DE CÁLCULO ============

# Cargar datos desde el archivo Excel
def cargar_datos_excel():
    try:
        wb = openpyxl.load_workbook('Calificaciones I-2026.xlsx', data_only=True)
        datos = {}
        
        estudiantes = ['AGUSTIN ANDRADE', 'MAURICIO LEAL', 'MARIA DANIELA MEDINA', 'LUIS RIOS', 'DANIELA RONDON']
        
        # --- Cargar datos de Cuaderno (10 semanas) ---
        ws_cuaderno = wb['Cuaderno']
        datos['cuaderno'] = {}
        for i, estudiante in enumerate(estudiantes, start=16):
            notas = []
            for j in range(3, 13):  # Columnas C a L (10 semanas)
                valor = get_safe_value(ws_cuaderno.cell(row=i, column=j))
                notas.append(valor)
            datos['cuaderno'][estudiante] = notas
        
        # --- Cargar datos de Gestión del laboratorio (10 semanas) ---
        ws_gestion = wb['Gestión del laboratorio']
        datos['gestion'] = {}
        for i, estudiante in enumerate(estudiantes, start=4):
            notas = []
            for j in range(3, 13):  # Columnas C a L (10 semanas)
                valor = get_safe_value(ws_gestion.cell(row=i, column=j))
                notas.append(valor)
            datos['gestion'][estudiante] = notas
        
        # --- Cargar datos de Quiz (10 semanas) ---
        ws_quiz = wb['Quiz']
        datos['quiz'] = {}
        for i, estudiante in enumerate(estudiantes, start=5):
            notas = []
            for j in range(3, 13):  # Columnas C a L (10 semanas)
                valor = get_safe_value(ws_quiz.cell(row=i, column=j))
                notas.append(valor)
            datos['quiz'][estudiante] = notas
        
        # --- Cargar datos de Informes (5 informes A-E) ---
        ws_informes = wb['Notas informes']
        datos['informes'] = {}
        for i, estudiante in enumerate(estudiantes, start=4):
            notas = []
            for j in range(3, 8):  # Columnas C a G (A-E)
                valor = get_safe_value(ws_informes.cell(row=i, column=j))
                notas.append(valor)
            datos['informes'][estudiante] = notas
        
        # --- Cargar ponderaciones desde la hoja Ponderaciones ---
        ws_pond = wb['Ponderaciones']
        ponderaciones = {
            'cuaderno_pct': get_safe_value(ws_pond.cell(row=7, column=3)) / 100,
            'gestion_pct': get_safe_value(ws_pond.cell(row=8, column=3)) / 100,
            'quiz_pct': get_safe_value(ws_pond.cell(row=9, column=3)) / 100,
            'informe_pct': get_safe_value(ws_pond.cell(row=10, column=3)) / 100,
            'bloqueA_pct': get_safe_value(ws_pond.cell(row=3, column=3)) / 100,
            'bloqueB_pct': get_safe_value(ws_pond.cell(row=4, column=3)) / 100
        }
        datos['ponderaciones'] = ponderaciones
        
        # --- Calcular Bloque A para cada estudiante ---
        datos['bloqueA'] = {}
        datos['semanas_evaluadas'] = {}
        for estudiante in estudiantes:
            cuaderno = datos['cuaderno'][estudiante]
            gestion = datos['gestion'][estudiante]
            quiz = datos['quiz'][estudiante]
            informe = datos['informes'][estudiante]
            
            # Contar semanas evaluadas (para mostrar en el dashboard)
            semanas_cuaderno = contar_semanas_evaluadas(cuaderno, 10)
            semanas_gestion = contar_semanas_evaluadas(gestion, 10)
            semanas_quiz = contar_semanas_evaluadas(quiz, 10)
            
            datos['semanas_evaluadas'][estudiante] = {
                'cuaderno': semanas_cuaderno,
                'gestion': semanas_gestion,
                'quiz': semanas_quiz
            }
            
            bloqueA = calcular_nota_bloqueA(
                cuaderno, gestion, quiz, informe,
                {
                    'cuaderno_pct': ponderaciones['cuaderno_pct'],
                    'gestion_pct': ponderaciones['gestion_pct'],
                    'quiz_pct': ponderaciones['quiz_pct'],
                    'informe_pct': ponderaciones['informe_pct']
                }
            )
            datos['bloqueA'][estudiante] = bloqueA
        
        # --- Cargar Bloque B (Seminario) - ya está en base 20 ---
        ws_bloqueB = wb['Seminario - Bloque B']
        datos['bloqueB_raw'] = {}
        for i, estudiante in enumerate(estudiantes, start=3):
            valor = get_safe_value(ws_bloqueB.cell(row=i, column=3))
            datos['bloqueB_raw'][estudiante] = valor  # Ya está en base 20
        
        # --- Calcular Nota Final para cada estudiante ---
        datos['final'] = {}
        for estudiante in estudiantes:
            bloqueA = datos['bloqueA'][estudiante]
            bloqueB = datos['bloqueB_raw'].get(estudiante, 0)
            
            nota_final = calcular_nota_final(
                bloqueA, 
                bloqueB,
                ponderaciones['bloqueA_pct'],
                ponderaciones['bloqueB_pct']
            )
            datos['final'][estudiante] = nota_final
        
        # Imprimir resumen para verificar
        print("\n" + "="*70)
        print("📊 DATOS CARGADOS Y CALCULADOS (Base 20 - 10 Semanas)")
        print("="*70)
        print(f"\nPonderaciones:")
        print(f"  Cuaderno: {ponderaciones['cuaderno_pct']*100:.0f}%")
        print(f"  Gestión: {ponderaciones['gestion_pct']*100:.0f}%")
        print(f"  Quiz: {ponderaciones['quiz_pct']*100:.0f}%")
        print(f"  Informe: {ponderaciones['informe_pct']*100:.0f}%")
        print(f"  Bloque A: {ponderaciones['bloqueA_pct']*100:.0f}%")
        print(f"  Bloque B: {ponderaciones['bloqueB_pct']*100:.0f}%")
        
        print("\n--- NOTAS CALCULADAS (10 semanas) ---")
        for estudiante in estudiantes:
            cuaderno_avg = calcular_nota_cuaderno(datos['cuaderno'][estudiante])
            gestion_avg = calcular_nota_gestion(datos['gestion'][estudiante])
            quiz_avg = calcular_nota_quiz(datos['quiz'][estudiante])
            informe_sum = calcular_nota_informe(datos['informes'][estudiante])
            
            semanas_info = datos['semanas_evaluadas'][estudiante]
            
            print(f"\n{estudiante}:")
            print(f"  Cuaderno (prom): {cuaderno_avg:.2f} ({semanas_info['cuaderno']} semanas)")
            print(f"  Gestión (prom): {gestion_avg:.2f} ({semanas_info['gestion']} semanas)")
            print(f"  Quiz (prom): {quiz_avg:.2f} ({semanas_info['quiz']} semanas)")
            print(f"  Informe (sum): {informe_sum:.2f}")
            print(f"  Bloque A: {datos['bloqueA'][estudiante]:.2f}")
            print(f"  Bloque B: {datos['bloqueB_raw'][estudiante]:.2f}")
            print(f"  NOTA FINAL: {datos['final'][estudiante]:.2f}")
        print("="*70)
        
        return datos
    
    except FileNotFoundError:
        print("El archivo Excel no se encuentra. Usando datos de ejemplo.")
        return generar_datos_ejemplo()
    except Exception as e:
        print(f"Error al cargar Excel: {e}")
        import traceback
        traceback.print_exc()
        return generar_datos_ejemplo()

def generar_datos_ejemplo():
    estudiantes = ['AGUSTIN ANDRADE', 'MAURICIO LEAL', 'MARIA DANIELA MEDINA', 'LUIS RIOS', 'DANIELA RONDON']
    datos = {
        'cuaderno': {},
        'gestion': {},
        'quiz': {},
        'informes': {},
        'bloqueA': {},
        'bloqueB_raw': {},
        'final': {},
        'semanas_evaluadas': {},
        'ponderaciones': {
            'cuaderno_pct': 0.20,
            'gestion_pct': 0.15,
            'quiz_pct': 0.25,
            'informe_pct': 0.40,
            'bloqueA_pct': 0.85,
            'bloqueB_pct': 0.15
        }
    }
    
    for estudiante in estudiantes:
        # 10 semanas de notas (solo primeras 4-5 con valores)
        datos['cuaderno'][estudiante] = [15.5, 18.0, 19.0, 17.0, 0, 0, 0, 0, 0, 0]
        datos['gestion'][estudiante] = [18.0, 19.0, 17.0, 18.0, 0, 0, 0, 0, 0, 0]
        datos['quiz'][estudiante] = [17.0, 18.0, 16.0, 17.5, 0, 0, 0, 0, 0, 0]
        datos['informes'][estudiante] = [15.0, 0, 0, 0, 0]
        
        datos['semanas_evaluadas'][estudiante] = {
            'cuaderno': 4,
            'gestion': 4,
            'quiz': 4
        }
        
        bloqueA = calcular_nota_bloqueA(
            datos['cuaderno'][estudiante],
            datos['gestion'][estudiante],
            datos['quiz'][estudiante],
            datos['informes'][estudiante],
            {
                'cuaderno_pct': 0.20,
                'gestion_pct': 0.15,
                'quiz_pct': 0.25,
                'informe_pct': 0.40
            }
        )
        datos['bloqueA'][estudiante] = bloqueA
        datos['bloqueB_raw'][estudiante] = 0
        datos['final'][estudiante] = bloqueA * 0.85
    
    return datos

# ============ FILTROS PARA FORMATO DECIMAL CON COMA ============

@app.template_filter('comma')
def format_with_comma(value):
    """Formatea un número con coma decimal SIEMPRE (incluso números enteros)"""
    if value is None:
        return "0,00"
    try:
        num = float(value)
        formatted = f"{num:.2f}".replace('.', ',')
        return formatted
    except (ValueError, TypeError):
        return "0,00"

@app.template_filter('comma_one')
def format_with_comma_one(value):
    """Formatea un número con coma decimal y 1 decimal SIEMPRE"""
    if value is None:
        return "0,0"
    try:
        num = float(value)
        formatted = f"{num:.1f}".replace('.', ',')
        return formatted
    except (ValueError, TypeError):
        return "0,0"

@app.template_filter('comma_zero')
def format_with_comma_zero(value):
    """Formatea un número sin decimales (número entero)"""
    if value is None:
        return "0"
    try:
        num = float(value)
        formatted = f"{num:.0f}".replace('.', ',')
        return formatted
    except (ValueError, TypeError):
        return "0"

@app.template_filter('comma_auto')
def format_with_comma_auto(value):
    """
    Formatea un número con coma decimal:
    - Si es entero: muestra sin decimales (ej: 15)
    - Si tiene decimales: muestra con 2 decimales (ej: 15,50)
    """
    if value is None:
        return "0"
    try:
        num = float(value)
        if num == int(num):
            return f"{int(num)}"
        else:
            formatted = f"{num:.2f}".replace('.', ',')
            return formatted
    except (ValueError, TypeError):
        return "0"

# Registrar filtros en Jinja2
app.jinja_env.filters['comma'] = format_with_comma
app.jinja_env.filters['comma_one'] = format_with_comma_one
app.jinja_env.filters['comma_zero'] = format_with_comma_zero
app.jinja_env.filters['comma_auto'] = format_with_comma_auto

# ============ FIN FILTROS ============

# Cargar datos al iniciar
DATOS = cargar_datos_excel()

# Guardar datos en JSON para acceso rápido
def guardar_datos_json():
    os.makedirs('data', exist_ok=True)
    with open('data/calificaciones.json', 'w', encoding='utf-8') as f:
        json.dump(DATOS, f, ensure_ascii=False, indent=2)

guardar_datos_json()

# Rutas

@app.route('/')
def index():
    if 'user_id' in session:
        if session.get('role') == 'admin':
            return redirect(url_for('admin'))
        return redirect(url_for('dashboard'))
    return redirect(url_for('login'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        
        if username in USERS and USERS[username]['password'] == password:
            session['user_id'] = username
            session['user_name'] = USERS[username]['name']
            session['role'] = USERS[username]['role']
            
            if session['role'] == 'admin':
                return redirect(url_for('admin'))
            return redirect(url_for('dashboard'))
        
        return render_template('login.html', error='Usuario o contraseña incorrectos')
    
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

@app.route('/dashboard')
@login_required
def dashboard():
    if session.get('role') == 'admin':
        return redirect(url_for('admin'))
    
    estudiante = session.get('user_name')
    cedula = session.get('user_id')
    
    # Obtener datos del estudiante
    datos_estudiante = {
        'nombre': estudiante,
        'cedula': cedula,
        'cuaderno': DATOS['cuaderno'].get(estudiante, [0]*10),
        'gestion': DATOS['gestion'].get(estudiante, [0]*10),
        'quiz': DATOS['quiz'].get(estudiante, [0]*10),
        'informes': DATOS['informes'].get(estudiante, [0]*5),
        'bloqueA': DATOS['bloqueA'].get(estudiante, 0),
        'bloqueB': DATOS['bloqueB_raw'].get(estudiante, 0),
        'final': DATOS['final'].get(estudiante, 0),
        'semanas': DATOS['semanas_evaluadas'].get(estudiante, {'cuaderno': 0, 'gestion': 0, 'quiz': 0})
    }
    
    # Agregar también los promedios para mostrar en el dashboard
    datos_estudiante['cuaderno_prom'] = calcular_nota_cuaderno(datos_estudiante['cuaderno'])
    datos_estudiante['gestion_prom'] = calcular_nota_gestion(datos_estudiante['gestion'])
    datos_estudiante['quiz_prom'] = calcular_nota_quiz(datos_estudiante['quiz'])
    datos_estudiante['informe_sum'] = calcular_nota_informe(datos_estudiante['informes'])
    
    return render_template('dashboard.html', estudiante=datos_estudiante)

@app.route('/admin')
@login_required
@admin_required
def admin():
    return render_template('admin.html', estudiantes=DATOS, enumerate=enumerate, range=range)

@app.route('/api/datos')
@login_required
def api_datos():
    if session.get('role') == 'admin':
        return jsonify(DATOS)
    
    estudiante = session.get('user_name')
    return jsonify({
        'nombre': estudiante,
        'datos': {
            'cuaderno': DATOS['cuaderno'].get(estudiante, []),
            'gestion': DATOS['gestion'].get(estudiante, []),
            'quiz': DATOS['quiz'].get(estudiante, []),
            'informes': DATOS['informes'].get(estudiante, []),
            'bloqueA': DATOS['bloqueA'].get(estudiante, 0),
            'bloqueB': DATOS['bloqueB_raw'].get(estudiante, 0),
            'final': DATOS['final'].get(estudiante, 0)
        }
    })

@app.route('/api/actualizar', methods=['POST'])
@login_required
@admin_required
def api_actualizar():
    data = request.json
    estudiante = data.get('estudiante')
    campo = data.get('campo')
    valor = data.get('valor')
    indice = data.get('indice', None)
    
    if not estudiante or not campo or valor is None:
        return jsonify({'error': 'Datos incompletos'}), 400
    
    try:
        valor = float(valor)
        if valor < 0 or valor > 20:
            return jsonify({'error': 'La nota debe estar entre 0 y 20'}), 400
            
        # Actualizar el campo específico
        if campo == 'bloqueA':
            DATOS['bloqueA'][estudiante] = valor
            # Recalcular nota final
            bloqueB = DATOS['bloqueB_raw'].get(estudiante, 0)
            DATOS['final'][estudiante] = calcular_nota_final(
                valor, 
                bloqueB,
                DATOS['ponderaciones']['bloqueA_pct'],
                DATOS['ponderaciones']['bloqueB_pct']
            )
        elif campo == 'bloqueB':
            DATOS['bloqueB_raw'][estudiante] = valor
            # Recalcular nota final
            bloqueA = DATOS['bloqueA'].get(estudiante, 0)
            DATOS['final'][estudiante] = calcular_nota_final(
                bloqueA, 
                valor,
                DATOS['ponderaciones']['bloqueA_pct'],
                DATOS['ponderaciones']['bloqueB_pct']
            )
        elif campo == 'final':
            # Si se modifica la final directamente, solo guardar el valor
            DATOS['final'][estudiante] = valor
        elif campo in ['cuaderno', 'gestion', 'quiz']:
            if indice is not None:
                DATOS[campo][estudiante][indice] = valor
            else:
                DATOS[campo][estudiante] = [float(v) for v in valor]
            
            # Actualizar conteo de semanas evaluadas
            semanas = contar_semanas_evaluadas(DATOS[campo][estudiante], 10)
            if campo == 'cuaderno':
                DATOS['semanas_evaluadas'][estudiante]['cuaderno'] = semanas
            elif campo == 'gestion':
                DATOS['semanas_evaluadas'][estudiante]['gestion'] = semanas
            elif campo == 'quiz':
                DATOS['semanas_evaluadas'][estudiante]['quiz'] = semanas
            
            # Recalcular Bloque A y Nota Final
            cuaderno = DATOS['cuaderno'][estudiante]
            gestion = DATOS['gestion'][estudiante]
            quiz = DATOS['quiz'][estudiante]
            informe = DATOS['informes'][estudiante]
            
            nuevo_bloqueA = calcular_nota_bloqueA(
                cuaderno, gestion, quiz, informe,
                {
                    'cuaderno_pct': DATOS['ponderaciones']['cuaderno_pct'],
                    'gestion_pct': DATOS['ponderaciones']['gestion_pct'],
                    'quiz_pct': DATOS['ponderaciones']['quiz_pct'],
                    'informe_pct': DATOS['ponderaciones']['informe_pct']
                }
            )
            DATOS['bloqueA'][estudiante] = nuevo_bloqueA
            
            # Recalcular nota final
            bloqueB = DATOS['bloqueB_raw'].get(estudiante, 0)
            DATOS['final'][estudiante] = calcular_nota_final(
                nuevo_bloqueA, 
                bloqueB,
                DATOS['ponderaciones']['bloqueA_pct'],
                DATOS['ponderaciones']['bloqueB_pct']
            )
        elif campo == 'informes':
            if indice is not None:
                DATOS['informes'][estudiante][indice] = valor
            else:
                DATOS['informes'][estudiante] = [float(v) for v in valor]
            
            # Recalcular Bloque A y Nota Final
            cuaderno = DATOS['cuaderno'][estudiante]
            gestion = DATOS['gestion'][estudiante]
            quiz = DATOS['quiz'][estudiante]
            informe = DATOS['informes'][estudiante]
            
            nuevo_bloqueA = calcular_nota_bloqueA(
                cuaderno, gestion, quiz, informe,
                {
                    'cuaderno_pct': DATOS['ponderaciones']['cuaderno_pct'],
                    'gestion_pct': DATOS['ponderaciones']['gestion_pct'],
                    'quiz_pct': DATOS['ponderaciones']['quiz_pct'],
                    'informe_pct': DATOS['ponderaciones']['informe_pct']
                }
            )
            DATOS['bloqueA'][estudiante] = nuevo_bloqueA
            
            # Recalcular nota final
            bloqueB = DATOS['bloqueB_raw'].get(estudiante, 0)
            DATOS['final'][estudiante] = calcular_nota_final(
                nuevo_bloqueA, 
                bloqueB,
                DATOS['ponderaciones']['bloqueA_pct'],
                DATOS['ponderaciones']['bloqueB_pct']
            )
        else:
            return jsonify({'error': 'Campo no válido'}), 400
        
        guardar_datos_json()
        return jsonify({'success': True, 'nuevos_datos': {
            'bloqueA': DATOS['bloqueA'][estudiante],
            'final': DATOS['final'][estudiante]
        }})
    except Exception as e:
        return jsonify({'error': str(e)}), 400

@app.route('/api/estudiantes')
@login_required
@admin_required
def api_estudiantes():
    return jsonify(list(DATOS['final'].keys()))

@app.route('/api/recalcular')
@login_required
@admin_required
def api_recalcular():
    """Recalcula todas las notas a partir de los datos actuales"""
    try:
        ponderaciones = DATOS['ponderaciones']
        estudiantes = list(DATOS['cuaderno'].keys())
        
        for estudiante in estudiantes:
            cuaderno = DATOS['cuaderno'][estudiante]
            gestion = DATOS['gestion'][estudiante]
            quiz = DATOS['quiz'][estudiante]
            informe = DATOS['informes'][estudiante]
            
            # Actualizar conteo de semanas
            DATOS['semanas_evaluadas'][estudiante] = {
                'cuaderno': contar_semanas_evaluadas(cuaderno, 10),
                'gestion': contar_semanas_evaluadas(gestion, 10),
                'quiz': contar_semanas_evaluadas(quiz, 10)
            }
            
            nuevo_bloqueA = calcular_nota_bloqueA(
                cuaderno, gestion, quiz, informe,
                {
                    'cuaderno_pct': ponderaciones['cuaderno_pct'],
                    'gestion_pct': ponderaciones['gestion_pct'],
                    'quiz_pct': ponderaciones['quiz_pct'],
                    'informe_pct': ponderaciones['informe_pct']
                }
            )
            DATOS['bloqueA'][estudiante] = nuevo_bloqueA
            
            bloqueB = DATOS['bloqueB_raw'].get(estudiante, 0)
            DATOS['final'][estudiante] = calcular_nota_final(
                nuevo_bloqueA, 
                bloqueB,
                ponderaciones['bloqueA_pct'],
                ponderaciones['bloqueB_pct']
            )
        
        guardar_datos_json()
        return jsonify({'success': True, 'mensaje': 'Notas recalculadas correctamente'})
    except Exception as e:
        return jsonify({'error': str(e)}), 400

@app.route('/api/estadisticas')
@login_required
@admin_required
def api_estadisticas():
    """Retorna estadísticas del curso"""
    notas = DATOS['final'].values()
    
    estadisticas = {
        'promedio': round(sum(notas) / len(notas), 2) if notas else 0,
        'maxima': max(notas) if notas else 0,
        'minima': min(notas) if notas else 0,
        'aprobados': len([n for n in notas if n >= 10]),  # Base 20
        'reprobados': len([n for n in notas if n < 10]),
        'total': len(notas)
    }
    
    return jsonify(estadisticas)

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
