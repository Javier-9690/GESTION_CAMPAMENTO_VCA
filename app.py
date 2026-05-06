import io
import os
import calendar
from datetime import datetime
from zoneinfo import ZoneInfo

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from flask import Flask, render_template, request, send_file, redirect, url_for
from dotenv import load_dotenv

load_dotenv()

SANTIAGO_TZ = ZoneInfo('America/Santiago')

def now_santiago():
    return datetime.now(SANTIAGO_TZ).replace(tzinfo=None)

app = Flask(__name__)
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'dev-hotel-key')
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

# ── Cabeceras de plantillas ────────────────────────────────────────────────

HOTEL_HEADERS = [
    'HABITACIÓN', 'MÓDULO', 'RUT', 'NOMBRE', 'EMPRESA',
    'N°CONTRATO', 'GERENCIA', 'SISTEMA TURNO', 'CALENDARIO',
    'ASIGNADO EN SALTO', 'CO MEL', 'GÉNERO', 'NOMBRE DE TURNO'
]

HOTEL_WIDTHS = [14, 12, 14, 28, 22, 14, 22, 16, 14, 18, 10, 10, 22]

# ── Estilos ────────────────────────────────────────────────────────────────

def thin_border():
    s = Side(style='thin')
    return Border(left=s, right=s, top=s, bottom=s)

def style_header(cell, fill_hex):
    cell.font = Font(bold=True, color='FFFFFF', size=10)
    cell.fill = PatternFill(start_color=fill_hex, end_color=fill_hex, fill_type='solid')
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    cell.border = thin_border()

def style_cell(cell, center=True):
    cell.alignment = Alignment(horizontal='center' if center else 'left', vertical='center')
    cell.border = thin_border()

# ── Generar Plantilla 1: Hotelería ─────────────────────────────────────────

def generate_hoteleria():
    wb = Workbook()
    ws = wb.active
    ws.title = 'Base Hotelería'
    ws.row_dimensions[1].height = 36

    for col, h in enumerate(HOTEL_HEADERS, 1):
        cell = ws.cell(row=1, column=col, value=h)
        style_header(cell, '1A5276')

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
        style_header(cell, '117A65')

    ws.column_dimensions['A'].width = 16
    ws.column_dimensions['B'].width = 16
    for d in range(1, days + 1):
        ws.column_dimensions[get_column_letter(d + 2)].width = 5

    ws.freeze_panes = 'C2'
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out

# ── Procesar archivos ──────────────────────────────────────────────────────

def read_excel_rows(file_obj):
    wb = openpyxl.load_workbook(file_obj, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    return rows

def normalize(val):
    if val is None:
        return ''
    return str(val).strip()

def process(hotel_file, censo_file):
    # ── Leer hotelería ──────────────────────────────────────────────────────
    hotel_rows = read_excel_rows(hotel_file)
    if not hotel_rows:
        raise ValueError('Plantilla hotelería vacía')

    hotel_headers = [normalize(c).upper() for c in hotel_rows[0]]

    def col(name_fragments):
        for i, h in enumerate(hotel_headers):
            if any(f in h for f in name_fragments):
                return i
        return None

    idx_hab  = col(['HABITACI'])
    idx_mod  = col(['MÓDULO', 'MODULO'])
    idx_nom  = col(['NOMBRE'])
    idx_rut  = col(['RUT'])
    idx_emp  = col(['EMPRESA'])
    idx_ger  = col(['GERENCIA'])
    idx_tur  = col(['SISTEMA', 'TURNO'])
    idx_gen  = col(['GÉNERO', 'GENERO'])

    if idx_hab is None or idx_mod is None:
        raise ValueError('Plantilla hotelería: faltan columnas HABITACIÓN / MÓDULO')

    hotel_data = {}  # key=(modulo, habitacion) -> list of personas
    for row in hotel_rows[1:]:
        if not any(row):
            continue
        hab = normalize(row[idx_hab] if idx_hab < len(row) else '')
        mod = normalize(row[idx_mod] if idx_mod < len(row) else '')
        if not hab or not mod:
            continue
        key = (mod.upper(), hab.upper())
        persona = {
            'nombre':  normalize(row[idx_nom] if idx_nom is not None and idx_nom < len(row) else ''),
            'rut':     normalize(row[idx_rut] if idx_rut is not None and idx_rut < len(row) else ''),
            'empresa': normalize(row[idx_emp] if idx_emp is not None and idx_emp < len(row) else ''),
            'gerencia':normalize(row[idx_ger] if idx_ger is not None and idx_ger < len(row) else ''),
            'turno':   normalize(row[idx_tur] if idx_tur is not None and idx_tur < len(row) else ''),
            'genero':  normalize(row[idx_gen] if idx_gen is not None and idx_gen < len(row) else ''),
        }
        hotel_data.setdefault(key, []).append(persona)

    # ── Leer censo ─────────────────────────────────────────────────────────
    censo_rows = read_excel_rows(censo_file)
    if not censo_rows:
        raise ValueError('Plantilla censo vacía')

    censo_headers = [normalize(c) for c in censo_rows[0]]
    day_indices = [(i, h) for i, h in enumerate(censo_headers) if h.isdigit()]
    total_dias = len(day_indices)

    if total_dias == 0:
        raise ValueError('Plantilla censo: no se encontraron columnas de días (1, 2, 3...)')

    mod_col2 = 0
    hab_col2 = 1
    for i, h in enumerate(censo_headers):
        hu = h.upper()
        if 'MODULO' in hu or 'MÓDULO' in hu:
            mod_col2 = i
        elif 'HABITACI' in hu or 'HABITACION' in hu:
            hab_col2 = i

    # Procesar habitaciones del censo
    rooms = []
    for row in censo_rows[1:]:
        if not any(row):
            continue
        mod = normalize(row[mod_col2] if mod_col2 < len(row) else '')
        hab = normalize(row[hab_col2] if hab_col2 < len(row) else '')
        if not mod or not hab:
            continue

        daily = []
        for i, _ in day_indices:
            v = row[i] if i < len(row) else None
            try:
                daily.append(int(float(str(v))) if v not in (None, '', 'None') else 0)
            except (ValueError, TypeError):
                daily.append(0)

        dias_ocupados = sum(1 for v in daily if v > 0)
        personas_dia  = sum(daily)
        camas_estimadas = max(daily) if daily else 0

        key = (mod.upper(), hab.upper())
        asignados = hotel_data.get(key, [])
        empresas  = list({p['empresa'] for p in asignados if p['empresa']})

        rooms.append({
            'modulo':           mod,
            'habitacion':       hab,
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

    # ── Estadísticas por módulo ────────────────────────────────────────────
    modulos = {}
    for r in rooms:
        m = r['modulo']
        if m not in modulos:
            modulos[m] = {
                'modulo': m,
                'habitaciones': 0,
                'camas': 0,
                'personas_asignadas': 0,
                'dias_ocupados': 0,
                'dias_totales': 0,
                'personas_dia_total': 0,
            }
        modulos[m]['habitaciones']        += 1
        modulos[m]['camas']               += r['camas_estimadas']
        modulos[m]['personas_asignadas']  += r['n_asignados']
        modulos[m]['dias_ocupados']       += r['dias_ocupados']
        modulos[m]['dias_totales']        += total_dias
        modulos[m]['personas_dia_total']  += r['personas_dia']

    for m in modulos.values():
        m['pct_ocupacion'] = round(m['dias_ocupados'] / m['dias_totales'] * 100, 1) if m['dias_totales'] else 0

    module_list = sorted(modulos.values(), key=lambda x: x['modulo'])

    # ── Estadísticas por empresa ───────────────────────────────────────────
    empresas = {}
    for r in rooms:
        for p in r['asignados']:
            emp = p['empresa'] or 'Sin empresa'
            if emp not in empresas:
                empresas[emp] = {
                    'empresa': emp,
                    'habitaciones': 0,
                    'personas': 0,
                    'dias_ocupados': 0,
                    'dias_totales': 0,
                    'dias_perdidos': 0,
                }
            empresas[emp]['personas'] += 1

        if r['empresas']:
            for emp in r['empresas']:
                if emp not in empresas:
                    empresas[emp] = {
                        'empresa': emp,
                        'habitaciones': 0,
                        'personas': 0,
                        'dias_ocupados': 0,
                        'dias_totales': 0,
                        'dias_perdidos': 0,
                    }
                empresas[emp]['habitaciones'] += 1
                empresas[emp]['dias_ocupados'] += r['dias_ocupados']
                empresas[emp]['dias_totales']  += total_dias
                empresas[emp]['dias_perdidos'] += r['dias_vacios']
        elif not r['asignados']:
            emp = 'Sin asignar'
            if emp not in empresas:
                empresas[emp] = {
                    'empresa': emp,
                    'habitaciones': 0,
                    'personas': 0,
                    'dias_ocupados': 0,
                    'dias_totales': 0,
                    'dias_perdidos': 0,
                }
            empresas[emp]['habitaciones'] += 1
            empresas[emp]['dias_ocupados'] += r['dias_ocupados']
            empresas[emp]['dias_totales']  += total_dias
            empresas[emp]['dias_perdidos'] += r['dias_vacios']

    for e in empresas.values():
        e['pct_ocupacion'] = round(e['dias_ocupados'] / e['dias_totales'] * 100, 1) if e['dias_totales'] else 0

    empresa_list = sorted(empresas.values(), key=lambda x: x['dias_perdidos'], reverse=True)

    # ── Totales globales ───────────────────────────────────────────────────
    total_hab    = len(rooms)
    total_asig   = sum(r['n_asignados'] for r in rooms)
    total_oc     = sum(r['dias_ocupados'] for r in rooms)
    total_pct    = round(total_oc / (total_hab * total_dias) * 100, 1) if total_hab * total_dias else 0
    total_camas  = sum(r['camas_estimadas'] for r in rooms)

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
    }

# ── Rutas ──────────────────────────────────────────────────────────────────

@app.route('/')
def index():
    now = now_santiago()
    meses = [
        'Enero','Febrero','Marzo','Abril','Mayo','Junio',
        'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'
    ]
    return render_template('index.html',
        mes=now.month, anio=now.year, meses=meses,
        anios=list(range(now.year - 1, now.year + 2)))

@app.route('/download/hoteleria')
def download_hoteleria():
    out = generate_hoteleria()
    return send_file(out,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='plantilla_hoteleria.xlsx')

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
    meses = [
        'Enero','Febrero','Marzo','Abril','Mayo','Junio',
        'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'
    ]

    if not hotel_file or not censo_file:
        return render_template('index.html',
            error='Debes subir ambas plantillas.',
            mes=now.month, anio=now.year, meses=meses,
            anios=list(range(now.year - 1, now.year + 2)))

    try:
        results = process(hotel_file, censo_file)
        return render_template('resultados.html', r=results)
    except Exception as e:
        return render_template('index.html',
            error=f'Error al procesar archivos: {str(e)}',
            mes=now.month, anio=now.year, meses=meses,
            anios=list(range(now.year - 1, now.year + 2)))

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5001)
