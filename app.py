import io
import os
import re
import calendar
from datetime import datetime
from zoneinfo import ZoneInfo

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from flask import Flask, render_template, request, send_file
from dotenv import load_dotenv

load_dotenv()

SANTIAGO_TZ = ZoneInfo('America/Santiago')

def now_santiago():
    return datetime.now(SANTIAGO_TZ).replace(tzinfo=None)

app = Flask(__name__)
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'dev-hotel-key')
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

# ── Cabeceras ──────────────────────────────────────────────────────────────

HOTEL_HEADERS = [
    'HABITACION', 'MODULO', 'RUT', 'NOMBRE', 'EMPRESA',
    'N CONTRATO', 'GERENCIA', 'SISTEMA TURNO',
    'CALENDARIO ASIGNADO EN SALTO', 'CO MEL', 'GENERO', 'NOMBRE DE TURNO'
]

HOTEL_WIDTHS = [14, 12, 14, 28, 22, 14, 22, 16, 30, 10, 10, 22]

# ── Estilos ────────────────────────────────────────────────────────────────

ARAMARK_RED = 'B42318'

def thin():
    s = Side(style='thin')
    return Border(left=s, right=s, top=s, bottom=s)

def style_header(cell, fill_hex=ARAMARK_RED):
    cell.font = Font(bold=True, color='FFFFFF', size=10)
    cell.fill = PatternFill(start_color=fill_hex, end_color=fill_hex, fill_type='solid')
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    cell.border = thin()

def style_cell(cell):
    cell.alignment = Alignment(horizontal='center', vertical='center')
    cell.border = thin()

# ── Generar Plantilla 1: Hoteleria ─────────────────────────────────────────

def generate_hoteleria():
    wb = Workbook()
    ws = wb.active
    ws.title = 'Base Hoteleria'
    ws.row_dimensions[1].height = 36

    for col, h in enumerate(HOTEL_HEADERS, 1):
        cell = ws.cell(row=1, column=col, value=h)
        style_header(cell)

    for i, w in enumerate(HOTEL_WIDTHS, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = 'A1:' + get_column_letter(len(HOTEL_HEADERS)) + '1'

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out

# ── Generar Plantilla 2: Censo Ontracking ─────────────────────────────────

def generate_censo(mes, anio):
    days = calendar.monthrange(int(anio), int(mes))[1]
    wb = Workbook()
    ws = wb.active
    ws.title = 'Censo Ontracking'
    ws.row_dimensions[1].height = 30

    headers = ['Modulo', 'Habitacion'] + [str(d) for d in range(1, days + 1)]
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        style_header(cell, '8B0000')

    ws.column_dimensions['A'].width = 16
    ws.column_dimensions['B'].width = 16
    for d in range(1, days + 1):
        ws.column_dimensions[get_column_letter(d + 2)].width = 5

    ws.freeze_panes = 'C2'
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out

# ── Leer Excel ─────────────────────────────────────────────────────────────

def read_rows(file_obj):
    wb = openpyxl.load_workbook(file_obj, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    return rows

# ── Normalización de valores ───────────────────────────────────────────────

def normalize(val):
    """Limpia y convierte a string. Convierte 101.0 -> '101'."""
    if val is None:
        return ''
    s = str(val).strip()
    # Excel a veces guarda enteros como float: "101.0" -> "101"
    if re.match(r'^\d+\.0+$', s):
        s = str(int(float(s)))
    return s

def normalize_mod(val):
    """
    Normaliza nombres de modulo a forma canonica 'M<N>' (sin ceros iniciales).
    Ejemplos:
      'M1', 'M01', 'M 1', 'Modulo 1', 'MOD 1', 'MODULO01', '1' -> 'M1'
      'M42' -> 'M42'
    """
    s = normalize(val).strip().upper()
    if not s:
        return ''
    # Reemplaza prefijos textuales por 'M'
    s = re.sub(r'^(MODULO|MOD)\s*', 'M', s)
    # 'M 01', 'M-01', 'M_01' -> quita separador y ceros iniciales del numero
    m = re.match(r'^M[\s\-_]*0*(\d+)(.*)$', s)
    if m:
        return 'M' + str(int(m.group(1))) + m.group(2).strip()
    # Solo numero: '1', '01' -> 'M1'
    m = re.match(r'^0*(\d+)$', s)
    if m:
        return 'M' + str(int(m.group(1)))
    # Si ya tiene forma razonable, devuelve tal cual limpio
    return s

def normalize_hab(val):
    """
    Normaliza codigos de habitacion eliminando ceros iniciales en la parte numerica.
    Ejemplos:
      '101', '0101', '101.0' -> '101'
      '4201A', '04201A'      -> '4201A'
      '46001'                -> '46001'  (5 digitos con sentido propio)
    """
    s = normalize(val).strip().upper()
    if not s:
        return ''
    # Numero puro: elimina ceros iniciales
    if re.match(r'^\d+$', s):
        return str(int(s))
    # Numero + letra(s) al final (ej. '4201A', '04201B')
    m = re.match(r'^0*(\d+)([A-Z]+)$', s)
    if m:
        return str(int(m.group(1))) + m.group(2)
    return s

def make_key(mod_val, hab_val):
    return (normalize_mod(mod_val), normalize_hab(hab_val))

# ── Procesar y cruzar datos ────────────────────────────────────────────────

def process(hotel_file, censo_file):
    # ── Leer hoteleria ──────────────────────────────────────────────────────
    hotel_rows = read_rows(hotel_file)
    if not hotel_rows:
        raise ValueError('Plantilla hoteleria vacia')

    hotel_headers = [normalize(c).upper() for c in hotel_rows[0]]

    def find_col(fragments):
        for i, h in enumerate(hotel_headers):
            if any(f in h for f in fragments):
                return i
        return None

    idx_hab = find_col(['HABITACI', 'HABITACION'])
    idx_mod = find_col(['MODULO', 'MÓDULO'])
    idx_nom = find_col(['NOMBRE'])
    idx_rut = find_col(['RUT'])
    idx_emp = find_col(['EMPRESA'])
    idx_ger = find_col(['GERENCIA'])
    idx_tur = find_col(['SISTEMA', 'TURNO'])
    idx_gen = find_col(['GENERO', 'GÉNERO'])
    idx_con = find_col(['CONTRATO'])
    idx_cal = find_col(['CALENDARIO', 'SALTO'])
    idx_mel = find_col(['MEL'])

    if idx_hab is None or idx_mod is None:
        raise ValueError(
            f'Plantilla hoteleria: faltan columnas HABITACION/MODULO. '
            f'Columnas detectadas: {hotel_headers}'
        )

    hotel_data = {}
    hotel_rows_loaded = 0
    for row in hotel_rows[1:]:
        if not any(row):
            continue
        raw_hab = normalize(row[idx_hab] if idx_hab < len(row) else '')
        raw_mod = normalize(row[idx_mod] if idx_mod < len(row) else '')
        if not raw_hab or not raw_mod:
            continue
        key = make_key(raw_mod, raw_hab)

        def get(idx):
            return normalize(row[idx] if idx is not None and idx < len(row) else '')

        persona = {
            'nombre':    get(idx_nom),
            'rut':       get(idx_rut),
            'empresa':   get(idx_emp),
            'gerencia':  get(idx_ger),
            'turno':     get(idx_tur),
            'genero':    get(idx_gen),
            'contrato':  get(idx_con),
            'calendario':get(idx_cal),
            'co_mel':    get(idx_mel),
            '_raw_mod':  raw_mod,
            '_raw_hab':  raw_hab,
        }
        hotel_data.setdefault(key, []).append(persona)
        hotel_rows_loaded += 1

    # ── Leer censo ─────────────────────────────────────────────────────────
    censo_rows = read_rows(censo_file)
    if not censo_rows:
        raise ValueError('Plantilla censo vacia')

    censo_headers = [normalize(c) for c in censo_rows[0]]
    day_indices = [(i, h) for i, h in enumerate(censo_headers) if h.isdigit()]
    total_dias = len(day_indices)

    if total_dias == 0:
        raise ValueError('Plantilla censo: no se encontraron columnas de dias (1, 2, 3...)')

    mod_col2 = 0
    hab_col2 = 1
    for i, h in enumerate(censo_headers):
        hu = h.upper()
        if 'MODULO' in hu or 'MÓDULO' in hu:
            mod_col2 = i
        elif 'HABITACI' in hu or 'HABITACION' in hu:
            hab_col2 = i

    rooms = []
    matched_count = 0
    for row in censo_rows[1:]:
        if not any(row):
            continue
        raw_mod = normalize(row[mod_col2] if mod_col2 < len(row) else '')
        raw_hab = normalize(row[hab_col2] if hab_col2 < len(row) else '')
        if not raw_mod or not raw_hab:
            continue

        key = make_key(raw_mod, raw_hab)

        daily = []
        for i, _ in day_indices:
            v = row[i] if i < len(row) else None
            try:
                daily.append(int(float(str(v))) if v not in (None, '', 'None') else 0)
            except (ValueError, TypeError):
                daily.append(0)

        dias_ocupados   = sum(1 for v in daily if v > 0)
        personas_dia    = sum(daily)
        camas_estimadas = max(daily) if daily else 0
        asignados = hotel_data.get(key, [])
        if asignados:
            matched_count += 1
        empresas = list({p['empresa'] for p in asignados if p['empresa']})

        rooms.append({
            'modulo':           raw_mod,
            'habitacion':       raw_hab,
            'key':              key,
            'daily':            daily,
            'dias_ocupados':    dias_ocupados,
            'dias_vacios':      total_dias - dias_ocupados,
            'personas_dia':     personas_dia,
            'camas_estimadas':  camas_estimadas,
            'pct_ocupacion':    round(dias_ocupados / total_dias * 100, 1) if total_dias else 0,
            'asignados':        asignados,
            'n_asignados':      len(asignados),
            'empresas':         empresas,
        })

    if not rooms:
        raise ValueError('No se encontraron habitaciones en el censo')

    # ── Estadisticas de matching ───────────────────────────────────────────
    # Muestra hasta 5 claves de cada archivo para diagnostico
    hotel_sample = sorted(hotel_data.keys())[:5]
    censo_sample = [make_key(
        normalize(row[mod_col2] if mod_col2 < len(row) else ''),
        normalize(row[hab_col2] if hab_col2 < len(row) else '')
    ) for row in censo_rows[1:] if any(row)][:5]

    match_info = {
        'hotel_rows':    hotel_rows_loaded,
        'censo_rooms':   len(rooms),
        'matched':       matched_count,
        'hotel_sample':  hotel_sample,
        'censo_sample':  censo_sample,
    }

    # ── Estadisticas por modulo ────────────────────────────────────────────
    modulos = {}
    for r in rooms:
        m = r['modulo']
        if m not in modulos:
            modulos[m] = {
                'modulo': m, 'habitaciones': 0, 'camas': 0,
                'personas_asignadas': 0, 'dias_ocupados': 0,
                'dias_totales': 0, 'personas_dia_total': 0,
            }
        modulos[m]['habitaciones']       += 1
        modulos[m]['camas']              += r['camas_estimadas']
        modulos[m]['personas_asignadas'] += r['n_asignados']
        modulos[m]['dias_ocupados']      += r['dias_ocupados']
        modulos[m]['dias_totales']       += total_dias
        modulos[m]['personas_dia_total'] += r['personas_dia']

    for m in modulos.values():
        m['pct_ocupacion'] = round(
            m['dias_ocupados'] / m['dias_totales'] * 100, 1
        ) if m['dias_totales'] else 0

    module_list = sorted(modulos.values(), key=lambda x: x['modulo'])

    # ── Estadisticas por empresa ───────────────────────────────────────────
    empresas = {}
    def get_emp(name):
        if name not in empresas:
            empresas[name] = {
                'empresa': name, 'habitaciones': 0, 'personas': 0,
                'dias_ocupados': 0, 'dias_totales': 0, 'dias_perdidos': 0,
            }
        return empresas[name]

    for r in rooms:
        for p in r['asignados']:
            get_emp(p['empresa'] or 'Sin empresa')['personas'] += 1

        tag_emps = r['empresas'] if r['empresas'] else (['Sin asignar'] if not r['asignados'] else [])
        for emp in tag_emps:
            e = get_emp(emp)
            e['habitaciones']  += 1
            e['dias_ocupados'] += r['dias_ocupados']
            e['dias_totales']  += total_dias
            e['dias_perdidos'] += r['dias_vacios']

    for e in empresas.values():
        e['pct_ocupacion'] = round(
            e['dias_ocupados'] / e['dias_totales'] * 100, 1
        ) if e['dias_totales'] else 0

    empresa_list = sorted(empresas.values(), key=lambda x: x['dias_perdidos'], reverse=True)

    total_hab  = len(rooms)
    total_asig = sum(r['n_asignados'] for r in rooms)
    total_oc   = sum(r['dias_ocupados'] for r in rooms)
    total_pct  = round(total_oc / (total_hab * total_dias) * 100, 1) if total_hab * total_dias else 0
    total_camas = sum(r['camas_estimadas'] for r in rooms)
    company_names = sorted({p['empresa'] for r in rooms for p in r['asignados'] if p['empresa']})

    return {
        'total_habitaciones': total_hab,
        'total_asignadas':    total_asig,
        'total_camas':        total_camas,
        'total_dias':         total_dias,
        'pct_ocupacion':      total_pct,
        'modulos':            module_list,
        'empresas':           empresa_list,
        'rooms':              sorted(rooms, key=lambda x: (x['modulo'], x['habitacion'])),
        'company_names':      company_names,
        'match_info':         match_info,
    }

# ── Rutas ──────────────────────────────────────────────────────────────────

MESES = ['Enero','Febrero','Marzo','Abril','Mayo','Junio',
         'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']

@app.route('/')
def index():
    now = now_santiago()
    return render_template('index.html',
        mes=now.month, anio=now.year, meses=MESES,
        anios=list(range(now.year - 1, now.year + 2)))

@app.route('/download/hoteleria')
def download_hoteleria():
    out = generate_hoteleria()
    return send_file(out,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True, download_name='plantilla_hoteleria.xlsx')

@app.route('/download/censo')
def download_censo():
    now = now_santiago()
    mes  = request.args.get('mes',  now.month,  type=int)
    anio = request.args.get('anio', now.year,   type=int)
    out = generate_censo(mes, anio)
    return send_file(out,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=f'plantilla_censo_{anio}_{str(mes).zfill(2)}.xlsx')

@app.route('/procesar', methods=['POST'])
def procesar():
    hotel_file = request.files.get('hoteleria')
    censo_file = request.files.get('censo')
    now = now_santiago()

    if not hotel_file or not censo_file:
        return render_template('index.html',
            error='Debes subir ambas plantillas.',
            mes=now.month, anio=now.year, meses=MESES,
            anios=list(range(now.year - 1, now.year + 2)))
    try:
        results = process(hotel_file, censo_file)
        return render_template('resultados.html', r=results)
    except Exception as e:
        return render_template('index.html',
            error=f'Error al procesar: {str(e)}',
            mes=now.month, anio=now.year, meses=MESES,
            anios=list(range(now.year - 1, now.year + 2)))

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5001)
