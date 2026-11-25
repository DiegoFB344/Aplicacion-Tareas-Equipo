import os
import io
import json
import uuid
import base64
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from flask import Flask, render_template, request, jsonify, session, redirect, url_for, send_file
from werkzeug.middleware.proxy_fix import ProxyFix
import gspread
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import Flow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Inches
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CLIENT_SECRET_PATH = os.path.join(BASE_DIR, 'client_secret_web.json')

# CONFIGURACIÓN: Usa variables de entorno en la nube o reemplaza los valores por defecto
SPREADSHEET_ID = os.environ.get('SPREADSHEET_ID', '1j8yHiDL6fHpwCZTNXVEvQOr-R9oPlXsyueH2ZvmwqHM')
TZ = ZoneInfo('America/Santiago')

app = Flask(
    __name__,
    template_folder=os.path.join(BASE_DIR, 'templates'),
    static_folder=os.path.join(BASE_DIR, 'static')
)
app.secret_key = os.environ.get('FLASK_SECRET_KEY', 'cambia-esta-clave-secreta-por-una-larga-2026')
app.wsgi_app = ProxyFix(app.wsgi_app, x_proto=1, x_host=1)

# Cookies seguras para HTTPS (Render / PythonAnywhere)
is_local = os.environ.get('FLASK_ENV') == 'development'
app.config.update(
    SESSION_COOKIE_SECURE=not is_local,
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE='Lax',
    PERMANENT_SESSION_LIFETIME=timedelta(days=30)
)

SHEETS = {
    'P': 'Proyectos', 'T': 'Tareas', 'U': 'Publicaciones',
    'R': 'Responsables', 'L': 'Historial', 'C': 'Configuracion'
}
ST = ['Pendiente', 'En proceso', 'En revisión', 'Finalizada']
SP = ['Reel', 'Feed', 'Historia', 'Haka Ara TV', 'Presentación', 'Impresión', 'Web', 'Audiovisual', 'Comunicado', 'Story']

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/calendar',
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/userinfo.email',
    'openid'
]

def create_oauth_flow(redirect_uri, state=None):
    """Soporta tanto archivo client_secret_web.json como variable de entorno GOOGLE_CLIENT_SECRET_JSON."""
    env_secret = os.environ.get('GOOGLE_CLIENT_SECRET_JSON')
    if env_secret:
        client_config = json.loads(env_secret)
        return Flow.from_client_config(client_config, scopes=SCOPES, state=state, redirect_uri=redirect_uri)
    if not os.path.exists(CLIENT_SECRET_PATH):
        raise FileNotFoundError("Falta client_secret_web.json o la variable GOOGLE_CLIENT_SECRET_JSON")
    return Flow.from_client_secrets_file(CLIENT_SECRET_PATH, scopes=SCOPES, state=state, redirect_uri=redirect_uri)

def creds_to_dict(creds):
    return {
        'token': creds.token,
        'refresh_token': creds.refresh_token,
        'token_uri': creds.token_uri,
        'client_id': creds.client_id,
        'client_secret': creds.client_secret,
        'scopes': creds.scopes
    }

def get_session_credentials():
    info = session.get('credentials')
    if not info:
        return None
    creds = Credentials(**info)
    if not creds.valid:
        if creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
                session['credentials'] = creds_to_dict(creds)
            except Exception:
                session.pop('credentials', None)
                return None
        else:
            session.pop('credentials', None)
            return None
    return creds

def get_google_services():
    creds = get_session_credentials()
    if not creds:
        raise PermissionError("AUTH_REQUIRED")
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(SPREADSHEET_ID)
    cal_service = build('calendar', 'v3', credentials=creds, cache_discovery=False)
    drive_service = build('drive', 'v3', credentials=creds, cache_discovery=False)
    oauth_service = build('oauth2', 'v2', credentials=creds, cache_discovery=False)
    user_info = oauth_service.userinfo().get().execute()
    user_email = user_info.get('email', 'Usuario')
    return sh, cal_service, drive_service, user_email

def all_rows(sh, sheet_name):
    ws = sh.worksheet(sheet_name)
    values = ws.get_all_values()
    if len(values) <= 1:
        return []
    return [{'v': row, 'row': i + 2} for i, row in enumerate(values[1:])]

def normalize_date(v):
    s = str(v or '').strip()[:10]
    if not s:
        return ''
    for fmt in ('%Y-%m-%d', '%d/%m/%Y', '%d-%m-%Y', '%Y/%m/%d'):
        try:
            return datetime.strptime(s, fmt).strftime('%Y-%m-%d')
        except ValueError:
            continue
    return s

def fmt_date(v):
    return normalize_date(v)

def valid_date(v):
    d = normalize_date(v)
    datetime.strptime(d, '%Y-%m-%d')
    return d

def next_day_str(date_str):
    dt = datetime.strptime(date_str, '%Y-%m-%d') + timedelta(days=1)
    return dt.strftime('%Y-%m-%d')

def state_progress(st):
    return [0, 33, 67, 100][max(0, ST.index(st))] if st in ST else 0

def get_shared_calendar_id(sh):
    for r in all_rows(sh, SHEETS['C']):
        if r['v'] and r['v'][0] == 'Calendar ID':
            cal_id = str(r['v'][1] if len(r['v']) > 1 else '').strip()
            if cal_id and cal_id != 'primary':
                return cal_id
    raise ValueError("Configura primero el 'Calendar ID' en la hoja Configuracion.")

def log_action(sh, action, item_id, detail, user_email='App Web/Móvil'):
    now = datetime.now(TZ).strftime('%Y-%m-%d %H:%M:%S')
    sh.worksheet(SHEETS['L']).append_row([now, action, str(item_id), str(detail), user_email])

def upsert_calendar_event(cal_service, cal_id, event_id, title, description, date_str):
    body = {
        'summary': title,
        'description': description,
        'start': {'date': date_str},
        'end': {'date': next_day_str(date_str)}
    }
    if event_id:
        try:
            updated = cal_service.events().update(calendarId=cal_id, eventId=event_id, body=body).execute()
            return updated['id']
        except Exception:
            pass
    created = cal_service.events().insert(calendarId=cal_id, body=body).execute()
    return created['id']

def build_agenda_payload(sh):
    ps = []
    for x in all_rows(sh, SHEETS['P']):
        v = x['v']
        if not v or not v[0]:
            continue
        archived = str(v[6]).upper() == 'TRUE' if len(v) > 6 else False
        ps.append({
            'id': str(v[0]),
            'name': str(v[1]) if len(v) > 1 else '',
            'description': str(v[2]) if len(v) > 2 else '',
            'priority': str(v[3]) if len(v) > 3 and v[3] else 'Media',
            'archived': archived,
            'archivedAt': fmt_date(v[7]) if len(v) > 7 else ''
        })

    ts = []
    for x in all_rows(sh, SHEETS['T']):
        v = x['v']
        if not v or not v[0]:
            continue
        ts.append({
            'id': str(v[0]),
            'projectId': str(v[1]) if len(v) > 1 else '',
            'title': str(v[2]) if len(v) > 2 else '',
            'support': str(v[3]) if len(v) > 3 else '',
            'text': str(v[4]) if len(v) > 4 else '',
            'assignee': str(v[5]) if len(v) > 5 else '',
            'email': str(v[6]) if len(v) > 6 else '',
            'start': fmt_date(v[7]) if len(v) > 7 else '',
            'due': fmt_date(v[8]) if len(v) > 8 else '',
            'status': str(v[9]) if len(v) > 9 and v[9] else 'Pendiente',
            'progress': int(v[10]) if len(v) > 10 and str(v[10]).isdigit() else 0,
            'calendarEventId': str(v[11]) if len(v) > 11 else ''
        })

    pubs = []
    for x in all_rows(sh, SHEETS['U']):
        v = x['v']
        if not v or not v[0]:
            continue
        pubs.append({
            'id': str(v[0]),
            'taskId': str(v[1]) if len(v) > 1 else '',
            'date': fmt_date(v[2]) if len(v) > 2 else '',
            'published': str(v[3]).upper() == 'TRUE' if len(v) > 3 else False
        })

    people = []
    for x in all_rows(sh, SHEETS['R']):
        v = x['v']
        if len(v) >= 2 and v[0] and v[1]:
            people.append({'name': str(v[0]), 'email': str(v[1])})

    return {'projects': ps, 'tasks': ts, 'publications': pubs, 'people': people, 'supports': SP, 'states': ST}

@app.route('/static/icon-512.png')
def pwa_icon():
    from PIL import Image, ImageDraw
    icon_file = os.path.join(BASE_DIR, 'static', '512.png')
    if os.path.exists(icon_file):
        return send_file(icon_file, mimetype='image/png')

    ico_path = os.path.join(BASE_DIR, 'logo.ico')
    if os.path.exists(ico_path):
        img = Image.open(ico_path).convert('RGBA').resize((512, 512))
    else:
        img = Image.new('RGBA', (512, 512), (54, 92, 207, 255))
        draw = ImageDraw.Draw(img)
        draw.rounded_rectangle([96, 96, 416, 416], radius=64, fill=(255, 255, 255, 255))
        draw.rectangle([150, 175, 362, 205], fill=(54, 92, 207, 255))
        draw.rectangle([150, 240, 362, 270], fill=(54, 92, 207, 255))
        draw.rectangle([150, 305, 290, 335], fill=(34, 160, 107, 255))

    buf = io.BytesIO()
    img.save(buf, format='PNG')
    buf.seek(0)
    return send_file(buf, mimetype='image/png')
# --- RUTAS PWA ---

@app.route('/sw.js')
def service_worker():
    return app.send_static_file('sw.js'), 200, {'Content-Type': 'application/javascript'}

@app.route('/')
def index():
    return render_template('index.html')

# --- RUTAS OAUTH WEB ---

@app.route('/api/authStatus', methods=['GET'])
def auth_status():
    creds = get_session_credentials()
    if not creds:
        return jsonify({'authenticated': False})
    try:
        oauth_service = build('oauth2', 'v2', credentials=creds, cache_discovery=False)
        info = oauth_service.userinfo().get().execute()
        return jsonify({'authenticated': True, 'email': info.get('email', '')})
    except Exception:
        return jsonify({'authenticated': False})

@app.route('/auth/login', methods=['GET'])
def auth_login():
    scheme = 'http' if is_local else 'https'
    redirect_uri = url_for('auth_callback', _external=True, _scheme=scheme)
    flow = create_oauth_flow(redirect_uri=redirect_uri)
    authorization_url, state = flow.authorization_url(
        access_type='offline',
        include_granted_scopes='true',
        prompt='select_account'
    )
    session['oauth_state'] = state
    return redirect(authorization_url)

@app.route('/oauth2callback', methods=['GET'])
def auth_callback():
    state = session.get('oauth_state')
    scheme = 'http' if is_local else 'https'
    redirect_uri = url_for('auth_callback', _external=True, _scheme=scheme)
    flow = create_oauth_flow(redirect_uri=redirect_uri, state=state)

    auth_response = request.url if is_local else request.url.replace('http://', 'https://')
    flow.fetch_token(authorization_response=auth_response)

    session.permanent = True
    session['credentials'] = creds_to_dict(flow.credentials)
    return redirect('/')

@app.route('/api/logout', methods=['POST'])
def logout():
    session.clear()
    return jsonify({'authenticated': False})

# --- EXPORTACIÓN DE GRÁFICOS (DESCARGA DIRECTA NAVEGADOR / APK) ---

@app.route('/api/exportCharts', methods=['POST'])
def export_charts():
    try:
        data = request.get_json(silent=True) or {}
        images = data.get('images', [])
        fmt = str(data.get('format', 'pdf')).lower()
        if not images:
            raise ValueError('No se recibieron gráficos para exportar')

        sh, _, _, user_email = get_google_services()
        now = datetime.now(TZ).strftime('%d-%m-%Y_%H-%M')
        now_readable = datetime.now(TZ).strftime('%d-%m-%Y %H:%M')
        buf = io.BytesIO()

        if fmt == 'pdf':
            filename = f"Estadisticas_Proyectos_{now}.pdf"
            mimetype = 'application/pdf'
            c = canvas.Canvas(buf, pagesize=letter)
            width, height = letter

            c.setFont("Helvetica-Bold", 18)
            c.setFillColorRGB(0.09, 0.19, 0.33)
            c.drawString(45, height - 50, "REPORTE ESTADÍSTICO DE PROYECTOS")

            c.setFont("Helvetica", 10)
            c.setFillColorRGB(0.38, 0.44, 0.55)
            c.drawString(45, height - 68, f"Generado: {now_readable}  |  Por: {user_email}")
            c.line(45, height - 78, width - 45, height - 78)

            y_pos = height - 105
            for idx, img in enumerate(images):
                if idx > 0 and idx % 2 == 0:
                    c.showPage()
                    y_pos = height - 60

                c.setFont("Helvetica-Bold", 12)
                c.setFillColorRGB(0.09, 0.19, 0.33)
                c.drawString(45, y_pos, img.get('title', 'Gráfico'))

                _, encoded = img['dataUrl'].split(',', 1)
                img_bytes = io.BytesIO(base64.b64decode(encoded))
                reader = ImageReader(img_bytes)
                iw, ih = reader.getSize()

                draw_w = 460
                draw_h = draw_w * (ih / float(iw))
                if draw_h > 265:
                    draw_h = 265
                    draw_w = draw_h * (iw / float(ih))

                c.drawImage(reader, 65, y_pos - draw_h - 12, width=draw_w, height=draw_h, mask='auto')
                y_pos -= (draw_h + 48)

            c.save()
        else:
            filename = f"Estadisticas_Proyectos_{now}.docx"
            mimetype = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            doc = Document()
            doc.add_heading('REPORTE ESTADÍSTICO DE PROYECTOS', 0)
            doc.add_paragraph(f"Generado: {now_readable} | Por: {user_email}")

            for img in images:
                doc.add_heading(img.get('title', 'Gráfico'), level=2)
                _, encoded = img['dataUrl'].split(',', 1)
                img_bytes = io.BytesIO(base64.b64decode(encoded))
                doc.add_picture(img_bytes, width=Inches(5.8))

            doc.save(buf)

        buf.seek(0)
        log_action(sh, f'Exportar gráficos {fmt.upper()}', '', filename, user_email)
        return send_file(buf, as_attachment=True, download_name=filename, mimetype=mimetype)
    except Exception as e:
        return jsonify({'error': str(e)}), 400

@app.route('/api/downloadArchivedReport', methods=['GET'])
def download_archived_report():
    try:
        sh, _, _, user_email = get_google_services()
        projects = [x for x in all_rows(sh, SHEETS['P']) if len(x['v']) > 6 and str(x['v'][6]).upper() == 'TRUE']
        if not projects:
            return jsonify({'error': 'No hay proyectos archivados'}), 400

        tasks = all_rows(sh, SHEETS['T'])
        pubs = all_rows(sh, SHEETS['U'])
        now = datetime.now(TZ).strftime('%d-%m-%Y_%H-%M')
        now_readable = datetime.now(TZ).strftime('%d-%m-%Y %H:%M')

        archived_ids = {str(p['v'][0]) for p in projects}
        archived_tasks = [t for t in tasks if str(t['v'][1]) in archived_ids]

        doc = Document()
        doc.add_heading('INFORME DE PROYECTOS FINALIZADOS', 0)
        doc.add_paragraph(f"Generado: {now_readable} | Proyectos archivados: {len(projects)} | Total tareas: {len(archived_tasks)} | Por: {user_email}")

        if archived_tasks:
            fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(10, 4.2))
            fig.patch.set_facecolor('white')

            resp_counts = {}
            for t in archived_tasks:
                name = str(t['v'][5] or 'Sin asignar').split(' ')[0]
                resp_counts[name] = resp_counts.get(name, 0) + 1
            ax1.bar(resp_counts.keys(), resp_counts.values(), color='#365ccf')
            ax1.set_title('Tareas finalizadas por responsable', fontsize=11, fontweight='bold', color='#183153')
            ax1.yaxis.get_major_locator().set_params(integer=True)

            supp_counts = {}
            for t in archived_tasks:
                sp = str(t['v'][3] or 'Otro')
                supp_counts[sp] = supp_counts.get(sp, 0) + 1
            ax2.barh(list(supp_counts.keys()), list(supp_counts.values()), color='#22a06b')
            ax2.set_title('Producción por tipo de soporte', fontsize=11, fontweight='bold', color='#183153')
            ax2.xaxis.get_major_locator().set_params(integer=True)

            plt.tight_layout()
            img_buffer = io.BytesIO()
            plt.savefig(img_buffer, format='png', dpi=150)
            plt.close(fig)
            img_buffer.seek(0)

            doc.add_heading('Resumen Estadístico', level=1)
            doc.add_picture(img_buffer, width=Inches(6.2))

        doc.add_heading('Detalle de Proyectos', level=1)
        for p in projects:
            doc.add_heading(str(p['v'][1]), level=2)
            doc.add_paragraph(f"Prioridad: {p['v'][3]} | Fecha de archivo: {fmt_date(p['v'][7])}")
            if p['v'][2]:
                doc.add_paragraph(f"Descripción: {p['v'][2]}")

            own = [t for t in archived_tasks if str(t['v'][1]) == str(p['v'][0])]
            for t in own:
                p_task = doc.add_paragraph(style='List Bullet')
                p_task.add_run(f"{t['v'][2]} ").bold = True
                p_task.add_run(f"| {t['v'][3]} | Responsable: {t['v'][5]} ({t['v'][6]})")
                doc.add_paragraph(f"Inicio: {fmt_date(t['v'][7])} | Entrega: {fmt_date(t['v'][8])} | Estado: {t['v'][9]}")
                if t['v'][4]:
                    doc.add_paragraph(f"Texto: {t['v'][4]}")
                tp = [u for u in pubs if str(u['v'][1]) == str(t['v'][0])]
                if tp:
                    pub_txt = '; '.join(f"{fmt_date(u['v'][2])} ({'Publicado' if str(u['v'][3]).upper()=='TRUE' else 'No publicado'})" for u in tp)
                    doc.add_paragraph(f"Publicaciones: {pub_txt}")

        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        filename = f"Informe_Proyectos_Finalizados_{now}.docx"
        log_action(sh, 'Generar informe Word', '', filename, user_email)
        return send_file(buf, as_attachment=True, download_name=filename, mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    except Exception as e:
        return jsonify({'error': str(e)}), 400

# --- RUTA GENÉRICA DE OPERACIONES ---

@app.route('/api/<action>', methods=['POST'])
def api_handler(action):
    try:
        data = request.get_json(silent=True) or {}
        args = data.get('args', [])
        sh, cal_service, drive_service, user_email = get_google_services()

        if action == 'getAgenda':
            return jsonify(build_agenda_payload(sh))

        elif action == 'probarCalendarV5':
            cal_id = get_shared_calendar_id(sh)
            cal = cal_service.calendars().get(calendarId=cal_id).execute()
            return jsonify({'message': f"Calendario compartido conectado: {cal.get('summary')} | ID: {cal.get('id')}"})

        elif action == 'saveProject':
            payload = args[0]
            name = str(payload.get('name', '')).strip()
            if not name:
                raise ValueError('Indica el nombre del proyecto')
            priority = str(payload.get('priority', 'Media'))
            ws = sh.worksheet(SHEETS['P'])
            rows = all_rows(sh, SHEETS['P'])
            old = next((x for x in rows if str(x['v'][0]) == str(payload.get('id'))), None) if payload.get('id') else None
            pid = str(old['v'][0]) if old else str(uuid.uuid4())
            now = datetime.now(TZ).strftime('%Y-%m-%d %H:%M')
            created = old['v'][4] if old and len(old['v']) > 4 else now
            archived = (str(old['v'][6]).upper() == 'TRUE') if old and len(old['v']) > 6 else False
            archived_at = old['v'][7] if old and len(old['v']) > 7 else ''
            row_data = [pid, name, str(payload.get('description', '')), priority, created, now, archived, archived_at]
            if old:
                ws.update(f"A{old['row']}:H{old['row']}", [row_data])
            else:
                ws.append_row(row_data, value_input_option='USER_ENTERED')
            log_action(sh, 'Guardar proyecto', pid, name, user_email)

        elif action == 'saveTask':
            payload = args[0]
            p = next((x for x in all_rows(sh, SHEETS['P']) if str(x['v'][0]) == str(payload.get('projectId'))), None)
            if not p:
                raise ValueError('Selecciona un proyecto existente')
            if len(p['v']) > 6 and str(p['v'][6]).upper() == 'TRUE':
                raise ValueError('Restaura el proyecto archivado antes de editar sus tareas')

            title = str(payload.get('title', '')).strip()
            support = str(payload.get('support', ''))
            if not title or support not in SP:
                raise ValueError('Completa el título y soporte válido')

            people = [x['v'] for x in all_rows(sh, SHEETS['R']) if len(x['v']) >= 2]
            who = next((x for x in people if x[1] == payload.get('email')), None)
            if not who:
                raise ValueError('Selecciona un responsable válido')

            due = valid_date(payload.get('due'))
            start = valid_date(payload.get('start')) if payload.get('start') else ''
            if start and start > due:
                raise ValueError('Inicio posterior al vencimiento')

            ws = sh.worksheet(SHEETS['T'])
            old = next((x for x in all_rows(sh, SHEETS['T']) if str(x['v'][0]) == str(payload.get('id'))), None) if payload.get('id') else None
            tid = str(old['v'][0]) if old else str(uuid.uuid4())
            status = str(old['v'][9]) if old and len(old['v']) > 9 else 'Pendiente'
            cal_id = get_shared_calendar_id(sh)
            old_ev_id = str(old['v'][11]) if old and len(old['v']) > 12 and str(old['v'][12]) == cal_id else ''

            desc = f"Proyecto: {p['v'][1]}\nTarea: {title}\nSoporte: {support}\nResponsable: {who[0]} ({who[1]})\nEstado: {status}\nTexto:\n{payload.get('text', '')}\nID tarea: {tid}"
            ev_title = f"[Proyecto] {p['v'][1]} | {title} ({support})"
            ev_id = upsert_calendar_event(cal_service, cal_id, old_ev_id, ev_title, desc, due)

            now = datetime.now(TZ).strftime('%Y-%m-%d %H:%M')
            row_data = [tid, str(p['v'][0]), title, support, str(payload.get('text', '')), who[0], who[1], start, due, status, state_progress(status), ev_id, cal_id, now]
            if old:
                ws.update(f"A{old['row']}:N{old['row']}", [row_data])
            else:
                ws.append_row(row_data, value_input_option='USER_ENTERED')
            log_action(sh, 'Guardar tarea', tid, title, user_email)

        elif action == 'shiftStatus':
            tid, delta = str(args[0]), int(args[1])
            ws = sh.worksheet(SHEETS['T'])
            old = next((x for x in all_rows(sh, SHEETS['T']) if str(x['v'][0]) == tid), None)
            if not old:
                raise ValueError('Tarea inexistente')
            current = str(old['v'][9])
            next_st = ST[max(0, min(3, ST.index(current) + delta))]
            if next_st != current:
                now = datetime.now(TZ).strftime('%Y-%m-%d %H:%M')
                ws.update(f"J{old['row']}:K{old['row']}", [[next_st, state_progress(next_st)]])
                ws.update(f"N{old['row']}", [[now]])
                p = next((x for x in all_rows(sh, SHEETS['P']) if str(x['v'][0]) == str(old['v'][1])), None)
                cal_id = get_shared_calendar_id(sh)
                old_ev_id = str(old['v'][11]) if len(old['v']) > 12 and str(old['v'][12]) == cal_id else ''
                desc = f"Proyecto: {p['v'][1]}\nTarea: {old['v'][2]}\nSoporte: {old['v'][3]}\nResponsable: {old['v'][5]} ({old['v'][6]})\nEstado: {next_st}\nTexto:\n{old['v'][4]}\nID tarea: {tid}"
                upsert_calendar_event(cal_service, cal_id, old_ev_id, f"[Proyecto] {p['v'][1]} | {old['v'][2]} ({old['v'][3]})", desc, valid_date(old['v'][8]))
                log_action(sh, 'Cambiar estado', tid, f"{current} → {next_st}", user_email)

        elif action == 'savePublication':
            payload = args[0]
            t = next((x for x in all_rows(sh, SHEETS['T']) if str(x['v'][0]) == str(payload.get('taskId'))), None)
            if not t:
                raise ValueError('Selecciona una tarea existente')
            p = next((x for x in all_rows(sh, SHEETS['P']) if str(x['v'][0]) == str(t['v'][1])), None)
            date_str = valid_date(payload.get('date'))
            ws = sh.worksheet(SHEETS['U'])
            rows = all_rows(sh, SHEETS['U'])
            old = next((x for x in rows if str(x['v'][0]) == str(payload.get('id'))), None) if payload.get('id') else None

            if any(str(x['v'][1]) == str(t['v'][0]) and fmt_date(x['v'][2]) == date_str and (not old or x['row'] != old['row']) for x in rows):
                raise ValueError('Esta tarea ya tiene una publicación en esa fecha')

            pub_id = str(old['v'][0]) if old else str(uuid.uuid4())
            published = (str(old['v'][3]).upper() == 'TRUE') if old and len(old['v']) > 3 else False
            cal_id = get_shared_calendar_id(sh)
            old_ev_id = str(old['v'][5]) if old and len(old['v']) > 6 and str(old['v'][6]) == cal_id else ''
            desc = f"Publicación programada\nProyecto: {p['v'][1]}\nTarea: {t['v'][2]}\nSoporte: {t['v'][3]}\nResponsable: {t['v'][5]} ({t['v'][6]})\nPublicado: {'Sí' if published else 'No'}"
            ev_id = upsert_calendar_event(cal_service, cal_id, old_ev_id, f"[Publicación] {p['v'][1]} | {t['v'][3]} | {t['v'][2]}", desc, date_str)

            alert_sent = old['v'][4] if old and len(old['v']) > 4 and fmt_date(old['v'][2]) == date_str else ''
            row_data = [pub_id, str(t['v'][0]), date_str, published, alert_sent, ev_id, cal_id]
            if old:
                ws.update(f"A{old['row']}:G{old['row']}", [row_data])
            else:
                ws.append_row(row_data, value_input_option='USER_ENTERED')
            log_action(sh, 'Programar publicación', pub_id, date_str, user_email)

        elif action == 'togglePublished':
            pub_id, val = str(args[0]), bool(args[1])
            ws = sh.worksheet(SHEETS['U'])
            pub = next((x for x in all_rows(sh, SHEETS['U']) if str(x['v'][0]) == pub_id), None)
            if not pub:
                raise ValueError('Publicación inexistente')
            ws.update(f"D{pub['row']}", [[val]])
            t = next((x for x in all_rows(sh, SHEETS['T']) if str(x['v'][0]) == str(pub['v'][1])), None)
            p = next((x for x in all_rows(sh, SHEETS['P']) if str(x['v'][0]) == str(t['v'][1])), None)
            cal_id = get_shared_calendar_id(sh)
            old_ev_id = str(pub['v'][5]) if len(pub['v']) > 6 and str(pub['v'][6]) == cal_id else ''
            desc = f"Publicación programada\nProyecto: {p['v'][1]}\nTarea: {t['v'][2]}\nSoporte: {t['v'][3]}\nResponsable: {t['v'][5]} ({t['v'][6]})\nPublicado: {'Sí' if val else 'No'}"
            upsert_calendar_event(cal_service, cal_id, old_ev_id, f"[Publicación] {p['v'][1]} | {t['v'][3]} | {t['v'][2]}", desc, valid_date(pub['v'][2]))
            log_action(sh, 'Publicado', pub_id, 'Sí' if val else 'No', user_email)

        elif action == 'deletePublication':
            pub_id = str(args[0])
            ws = sh.worksheet(SHEETS['U'])
            pub = next((x for x in all_rows(sh, SHEETS['U']) if str(x['v'][0]) == pub_id), None)
            if not pub:
                raise ValueError('Publicación inexistente')
            cal_id = get_shared_calendar_id(sh)
            if len(pub['v']) > 6 and pub['v'][5] and str(pub['v'][6]) == cal_id:
                try:
                    cal_service.events().delete(calendarId=cal_id, eventId=str(pub['v'][5])).execute()
                except Exception:
                    pass
            ws.delete_rows(pub['row'])
            log_action(sh, 'Eliminar publicación', pub_id, '', user_email)

        elif action == 'archivarProyectoV7':
            pid = str(args[0])
            ws = sh.worksheet(SHEETS['P'])
            p = next((x for x in all_rows(sh, SHEETS['P']) if str(x['v'][0]) == pid), None)
            if not p:
                raise ValueError('Proyecto no encontrado')
            tasks = [x for x in all_rows(sh, SHEETS['T']) if str(x['v'][1]) == pid]
            if not tasks or not all(str(x['v'][9]) == 'Finalizada' for x in tasks):
                raise ValueError('Todas las tareas deben estar en estado Finalizada')
            ws.update(f"G{p['row']}:H{p['row']}", [[True, datetime.now(TZ).strftime('%Y-%m-%d')]])
            log_action(sh, 'Archivar proyecto', pid, str(p['v'][1]), user_email)

        elif action == 'restaurarProyectoV7':
            pid = str(args[0])
            ws = sh.worksheet(SHEETS['P'])
            p = next((x for x in all_rows(sh, SHEETS['P']) if str(x['v'][0]) == pid), None)
            if not p:
                raise ValueError('Proyecto no encontrado')
            ws.update(f"G{p['row']}:H{p['row']}", [[False, '']])
            log_action(sh, 'Restaurar proyecto', pid, str(p['v'][1]), user_email)

        elif action == 'generarInformeFinalizadosV7':
            projects = [x for x in all_rows(sh, SHEETS['P']) if len(x['v']) > 6 and str(x['v'][6]).upper() == 'TRUE']
            if not projects:
                raise ValueError('Todavía no hay proyectos archivados para informar')
            return jsonify({'downloadUrl': '/api/downloadArchivedReport', 'count': len(projects)})

        return jsonify(build_agenda_payload(sh))

    except PermissionError as pe:
        return jsonify({'error': str(pe)}), 401
    except Exception as e:
        return jsonify({'error': str(e)}), 400

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)