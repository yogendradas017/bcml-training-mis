import os
import sqlite3
import io
from datetime import date, datetime
from functools import wraps
from flask import (Flask, render_template, request, redirect, url_for,
                   session, jsonify, flash, send_file, g)
from werkzeug.security import generate_password_hash, check_password_hash
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'bcml-tms-2627-xK9pQ')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH  = os.path.join(BASE_DIR, 'data', 'training.db')

PLANTS = [
    {'id': 1,  'name': 'Balrampur',  'unit_code': 'BCM'},
    {'id': 2,  'name': 'Babhnan',    'unit_code': 'BBN'},
    {'id': 3,  'name': 'Rauzagaon',  'unit_code': 'RCM'},
    {'id': 4,  'name': 'Maizapur',   'unit_code': 'MZP'},
    {'id': 5,  'name': 'Mankapur',   'unit_code': 'MCM'},
    {'id': 6,  'name': 'Gularia',    'unit_code': 'GCM'},
    {'id': 7,  'name': 'Tulsipur',   'unit_code': 'TCM'},
    {'id': 8,  'name': 'Kumbhi',     'unit_code': 'KCM'},
    {'id': 9,  'name': 'Haidergarh', 'unit_code': 'HCM'},
    {'id': 10, 'name': 'Akbarpur',   'unit_code': 'ACM'},
]
PLANT_MAP = {p['id']: p for p in PLANTS}

PROG_TYPES   = ['Behavioural/Leadership', 'Cane', 'Commercial', 'EHS/HR', 'IT', 'Technical']
MODES        = ['Classroom', 'OJT', 'SOP', 'Online']
LEVELS       = ['General', 'Specialized']
AUDIENCES    = ['Blue Collared', 'White Collared', 'Common']
STATUSES     = ['To Be Planned', 'Conducted', 'Re-Scheduled', 'Cancelled']
INT_EXT      = ['Internal', 'External', 'Online']
MONTHS_FY    = ['April','May','June','July','August','September',
                'October','November','December','January','February','March']
CAL_NEW      = ['Calendar Program', 'New Program']
GENDERS      = ['Male', 'Female', 'Others']
TYPE_ABBREV  = {
    'Behavioural/Leadership': 'BEH', 'Cane': 'CAN', 'Commercial': 'COM',
    'EHS/HR': 'EHS', 'IT': 'ITC', 'Technical': 'TEC'
}

# ─── DB ──────────────────────────────────────────────────────────────────────

def get_db():
    if 'db' not in g:
        g.db = sqlite3.connect(DB_PATH)
        g.db.row_factory = sqlite3.Row
        g.db.execute("PRAGMA foreign_keys = ON")
    return g.db

@app.teardown_appcontext
def close_db(e=None):
    db = g.pop('db', None)
    if db:
        db.close()

def init_db():
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    db = sqlite3.connect(DB_PATH)
    with open(os.path.join(BASE_DIR, 'schema.sql')) as f:
        db.executescript(f.read())
    # Seed plants
    for p in PLANTS:
        db.execute('INSERT OR IGNORE INTO plants(id,name,unit_code) VALUES(?,?,?)',
                   (p['id'], p['name'], p['unit_code']))
    # Default users
    users = [('central', 'bcml@1234', 'central', None),
             ('admin',   'admin@bcml', 'admin',   None)]
    for p in PLANTS:
        users.append((p['name'].lower(), 'bcml@1234', 'spoc', p['id']))
    for u in users:
        db.execute('INSERT OR IGNORE INTO users(username,password,role,plant_id) VALUES(?,?,?,?)',
                   (u[0], generate_password_hash(u[1]), u[2], u[3]))
    db.commit()
    db.close()

# ─── AUTH ─────────────────────────────────────────────────────────────────────

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated

def spoc_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        if session.get('role') not in ('spoc', 'admin'):
            flash('Access denied.', 'danger')
            return redirect(url_for('central_dashboard'))
        if session.get('role') == 'admin' and not session.get('plant_id'):
            flash('Please select a plant first to access SPOC functions.', 'warning')
            return redirect(url_for('central_dashboard'))
        return f(*args, **kwargs)
    return decorated

def central_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        if session.get('role') not in ('central', 'admin'):
            flash('Access denied.', 'danger')
            return redirect(url_for('spoc_dashboard'))
        return f(*args, **kwargs)
    return decorated

@app.route('/')
def index():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    if session.get('role') in ('central', 'admin'):
        return redirect(url_for('central_dashboard'))
    return redirect(url_for('spoc_dashboard'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form['username'].strip().lower()
        password = request.form['password']
        db = get_db()
        user = db.execute('SELECT * FROM users WHERE username=?', (username,)).fetchone()
        if user and check_password_hash(user['password'], password):
            session.clear()
            session['user_id']  = user['id']
            session['username'] = user['username']
            session['role']     = user['role']
            session['plant_id'] = user['plant_id']
            if user['plant_id']:
                session['plant_name'] = PLANT_MAP[user['plant_id']]['name']
                session['unit_code']  = PLANT_MAP[user['plant_id']]['unit_code']
            if user['role'] in ('central', 'admin'):
                return redirect(url_for('central_dashboard'))
            return redirect(url_for('spoc_dashboard'))
        flash('Invalid username or password.', 'danger')
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

# ─── ADMIN PLANT SELECTOR ─────────────────────────────────────────────────────

@app.route('/admin/plant/<int:plant_id>')
def admin_select_plant(plant_id):
    if session.get('role') != 'admin':
        return redirect(url_for('index'))
    plant = PLANT_MAP.get(plant_id)
    if not plant:
        flash('Plant not found.', 'danger')
        return redirect(url_for('central_dashboard'))
    session['plant_id']   = plant['id']
    session['plant_name'] = plant['name']
    session['unit_code']  = plant['unit_code']
    flash(f"Now viewing as SPOC for {plant['name']}. Use 'Switch Plant' in sidebar to go back.", 'info')
    return redirect(url_for('spoc_dashboard'))

@app.route('/admin/clear-plant')
def admin_clear_plant():
    if session.get('role') != 'admin':
        return redirect(url_for('index'))
    session.pop('plant_id',   None)
    session.pop('plant_name', None)
    session.pop('unit_code',  None)
    return redirect(url_for('central_dashboard'))

# ─── SPOC DASHBOARD ───────────────────────────────────────────────────────────

@app.route('/dashboard')
@spoc_required
def spoc_dashboard():
    plant_id = session['plant_id']
    db = get_db()
    stats = {
        'total_emp':    db.execute('SELECT COUNT(*) FROM employees WHERE plant_id=? AND is_active=1', (plant_id,)).fetchone()[0],
        'blue_collar':  db.execute("SELECT COUNT(*) FROM employees WHERE plant_id=? AND is_active=1 AND collar='Blue Collared'", (plant_id,)).fetchone()[0],
        'white_collar': db.execute("SELECT COUNT(*) FROM employees WHERE plant_id=? AND is_active=1 AND collar='White Collared'", (plant_id,)).fetchone()[0],
        'tni_count':    db.execute('SELECT COUNT(*) FROM tni WHERE plant_id=?', (plant_id,)).fetchone()[0],
        'sessions':     db.execute('SELECT COUNT(*) FROM calendar WHERE plant_id=?', (plant_id,)).fetchone()[0],
        'conducted':    db.execute("SELECT COUNT(*) FROM calendar WHERE plant_id=? AND status='Conducted'", (plant_id,)).fetchone()[0],
        'trainings':    db.execute('SELECT COUNT(*) FROM emp_training WHERE plant_id=?', (plant_id,)).fetchone()[0],
        'manhours':     db.execute('SELECT COALESCE(SUM(hrs),0) FROM emp_training WHERE plant_id=?', (plant_id,)).fetchone()[0],
    }
    return render_template('dashboard.html', stats=stats)

# ─── EMPLOYEE MASTER ──────────────────────────────────────────────────────────

@app.route('/employees')
@spoc_required
def employees():
    plant_id = session['plant_id']
    db = get_db()
    show_exited = request.args.get('show_exited', '0') == '1'
    if show_exited:
        emps = db.execute('SELECT * FROM employees WHERE plant_id=? ORDER BY name', (plant_id,)).fetchall()
    else:
        emps = db.execute('SELECT * FROM employees WHERE plant_id=? AND is_active=1 ORDER BY name', (plant_id,)).fetchall()
    return render_template('employees.html', employees=emps, show_exited=show_exited,
                           genders=GENDERS)

@app.route('/employees/add', methods=['POST'])
@spoc_required
def add_employee():
    plant_id = session['plant_id']
    f = request.form
    db = get_db()
    collar = normalise_collar(f.get('collar',''))
    try:
        db.execute('''INSERT INTO employees
            (plant_id,emp_code,name,designation,grade,collar,department,section,
             category,gender,physically_handicapped,remarks)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?)''',
            (plant_id, f['emp_code'].strip(), f['name'].strip(),
             f.get('designation',''), f.get('grade',''), collar,
             f.get('department',''), f.get('section',''), f.get('category',''),
             f.get('gender',''), f.get('physically_handicapped','No'),
             f.get('remarks','')))
        db.commit()
        flash(f"Employee {f['name'].strip()} added successfully.", 'success')
    except sqlite3.IntegrityError:
        flash(f"Employee code {f['emp_code'].strip()} already exists.", 'danger')
    return redirect(url_for('employees'))

@app.route('/employees/<int:emp_id>/exit', methods=['POST'])
@spoc_required
def exit_employee(emp_id):
    db = get_db()
    exit_date   = request.form.get('exit_date', str(date.today()))
    exit_reason = request.form.get('exit_reason', '')
    db.execute('UPDATE employees SET is_active=0, exit_date=?, exit_reason=? WHERE id=? AND plant_id=?',
               (exit_date, exit_reason, emp_id, session['plant_id']))
    db.commit()
    flash('Employee marked as exited.', 'warning')
    return redirect(url_for('employees'))

@app.route('/employees/<int:emp_id>/reactivate', methods=['POST'])
@spoc_required
def reactivate_employee(emp_id):
    db = get_db()
    db.execute('UPDATE employees SET is_active=1, exit_date=NULL, exit_reason=NULL WHERE id=? AND plant_id=?',
               (emp_id, session['plant_id']))
    db.commit()
    flash('Employee reactivated.', 'success')
    return redirect(url_for('employees') + '?show_exited=1')

# ─── TNI TRACKING ─────────────────────────────────────────────────────────────

@app.route('/tni')
@spoc_required
def tni():
    plant_id = session['plant_id']
    db = get_db()
    records = db.execute('''
        SELECT t.*, e.name, e.designation, e.grade, e.collar, e.department, e.section
        FROM tni t
        LEFT JOIN employees e ON e.emp_code=t.emp_code AND e.plant_id=t.plant_id
        WHERE t.plant_id=?
        ORDER BY t.id DESC
    ''', (plant_id,)).fetchall()

    # Mark completed where a matching 2A record exists
    completed_map = {}
    for r in records:
        exists = db.execute('''SELECT 1 FROM emp_training
            WHERE plant_id=? AND emp_code=? AND programme_name=?''',
            (plant_id, r['emp_code'], r['programme_name'])).fetchone()
        completed_map[r['id']] = 'Yes' if exists else 'No'

    emps = db.execute('SELECT emp_code, name FROM employees WHERE plant_id=? AND is_active=1 ORDER BY name', (plant_id,)).fetchall()
    programmes = _get_programme_names(plant_id, db)
    return render_template('tni.html', records=records, completed_map=completed_map,
                           employees=emps, programmes=programmes,
                           prog_types=PROG_TYPES, modes=MODES, months=MONTHS_FY)

@app.route('/tni/add', methods=['POST'])
@spoc_required
def add_tni():
    plant_id = session['plant_id']
    f = request.form
    db = get_db()
    db.execute('''INSERT INTO tni(plant_id,emp_code,programme_name,prog_type,mode,target_month,planned_hours)
                  VALUES(?,?,?,?,?,?,?)''',
               (plant_id, f['emp_code'], f['programme_name'].strip(),
                f.get('prog_type',''), f.get('mode',''),
                f.get('target_month',''), float(f.get('planned_hours') or 0)))
    db.commit()
    flash('TNI entry added.', 'success')
    return redirect(url_for('tni'))

@app.route('/tni/<int:tni_id>/delete', methods=['POST'])
@spoc_required
def delete_tni(tni_id):
    db = get_db()
    db.execute('DELETE FROM tni WHERE id=? AND plant_id=?', (tni_id, session['plant_id']))
    db.commit()
    flash('TNI entry deleted.', 'warning')
    return redirect(url_for('tni'))

@app.route('/tni/template')
@spoc_required
def tni_template():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'TNI_Bulk_Upload'
    headers = ['Employee Code', 'Programme Name', 'Type of Programme',
               'Mode', 'Target Month', 'Planned Hours']
    hdr_fill = PatternFill('solid', start_color='1F4E79')
    hdr_font = Font(bold=True, color='FFFFFF')
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=i, value=h)
        cell.fill = hdr_fill
        cell.font = hdr_font
        ws.column_dimensions[get_column_letter(i)].width = 22
    # Sample rows
    samples = [
        ['21700011', 'Fire Safety Training', 'EHS/HR', 'Classroom', 'June', 4],
        ['21101568', 'Leadership Development', 'Behavioural/Leadership', 'Classroom', 'July', 8],
    ]
    for r, row in enumerate(samples, 2):
        for c, val in enumerate(row, 1):
            ws.cell(row=r, column=c, value=val)
    ws['A4'] = 'VALID Programme Types:'
    ws['B4'] = 'Behavioural/Leadership | Cane | Commercial | EHS/HR | IT | Technical'
    ws['A5'] = 'VALID Modes:'
    ws['B5'] = 'Classroom | OJT | SOP | Online'
    ws['A6'] = 'VALID Months:'
    ws['B6'] = 'April | May | June | July | August | September | October | November | December | January | February | March'
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return send_file(out, download_name='TNI_Bulk_Upload_Template.xlsx', as_attachment=True,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@app.route('/tni/bulk', methods=['POST'])
@spoc_required
def tni_bulk_upload():
    plant_id = session['plant_id']
    f = request.files.get('file')
    if not f or f.filename == '':
        flash('No file selected.', 'danger')
        return redirect(url_for('tni'))
    try:
        df = _read_upload_file(f)
    except Exception as e:
        flash(f'Could not read file: {e}', 'danger')
        return redirect(url_for('tni'))

    db       = get_db()
    inserted = 0
    errors   = []
    for i, row in df.iterrows():
        emp_code  = _clean(row, ['employee code', 'emp code', 'empcode', 'employee_code'])
        prog_name = _clean(row, ['programme name', 'program name', 'programme_name', 'training name'])
        prog_type = _clean(row, ['type of programme', 'type', 'prog type', 'programme type'])
        mode      = _clean(row, ['mode'])
        month     = _clean(row, ['target month', 'month'])
        hours     = _safe_float(_clean(row, ['planned hours', 'hours', 'hrs'])) or 0

        if not emp_code or not prog_name:
            errors.append(f'Row {i+2}: Employee Code and Programme Name are required.')
            continue
        emp = db.execute('SELECT 1 FROM employees WHERE emp_code=? AND plant_id=? AND is_active=1',
                         (emp_code, plant_id)).fetchone()
        if not emp:
            errors.append(f'Row {i+2}: Employee {emp_code} not found in your plant.')
            continue
        db.execute('INSERT INTO tni(plant_id,emp_code,programme_name,prog_type,mode,target_month,planned_hours) VALUES(?,?,?,?,?,?,?)',
                   (plant_id, emp_code, prog_name, prog_type, mode, month, hours))
        inserted += 1
    db.commit()
    if inserted:
        flash(f'Bulk upload complete: {inserted} TNI entries added successfully.', 'success')
    if errors:
        for err in errors:
            flash(err, 'upload_error')
        flash(f'{len(errors)} row(s) had errors — see details below.', 'warning')
    return redirect(url_for('tni'))

# ─── TRAINING CALENDAR ────────────────────────────────────────────────────────

@app.route('/calendar')
@spoc_required
def training_calendar():
    plant_id = session['plant_id']
    db = get_db()

    # Auto-update statuses from 2C
    _sync_calendar_from_2c(plant_id, db)

    sessions = db.execute('SELECT * FROM calendar WHERE plant_id=? ORDER BY id DESC', (plant_id,)).fetchall()
    # TNI programme demand counts
    demand_map = {}
    for row in db.execute('SELECT programme_name, COUNT(*) as cnt FROM tni WHERE plant_id=? GROUP BY programme_name', (plant_id,)):
        demand_map[row['programme_name']] = row['cnt']

    return render_template('calendar.html', sessions=sessions, demand_map=demand_map,
                           prog_types=PROG_TYPES, modes=MODES, levels=LEVELS,
                           audiences=AUDIENCES, months=MONTHS_FY, statuses=STATUSES)

@app.route('/calendar/add', methods=['POST'])
@spoc_required
def add_calendar():
    plant_id = session['plant_id']
    f = request.form
    db = get_db()
    prog_name = f['programme_name'].strip()
    prog_type = f.get('prog_type', '')

    prog_code    = _get_or_create_prog_code(plant_id, prog_name, prog_type, db)
    session_code = _new_session_code(plant_id, prog_code, db)

    db.execute('''INSERT INTO calendar
        (plant_id,prog_code,session_code,source,programme_name,prog_type,
         planned_month,plan_start,plan_end,time_from,time_to,duration_hrs,
         level,mode,target_audience,planned_pax,trainer_vendor,status)
        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
        (plant_id, prog_code, session_code,
         f.get('source','TNI'), prog_name, prog_type,
         f.get('planned_month',''), f.get('plan_start',''), f.get('plan_end',''),
         f.get('time_from',''), f.get('time_to',''),
         float(f.get('duration_hrs') or 0),
         f.get('level',''), f.get('mode',''), f.get('target_audience',''),
         int(f.get('planned_pax') or 0), f.get('trainer_vendor',''),
         'To Be Planned'))
    db.commit()
    flash(f'Session {session_code} added to calendar.', 'success')
    return redirect(url_for('training_calendar'))

@app.route('/calendar/<int:cal_id>/delete', methods=['POST'])
@spoc_required
def delete_calendar(cal_id):
    db = get_db()
    db.execute('DELETE FROM calendar WHERE id=? AND plant_id=?', (cal_id, session['plant_id']))
    db.commit()
    flash('Calendar entry deleted.', 'warning')
    return redirect(url_for('training_calendar'))

# ─── EMPLOYEE TRAINING (2A) ───────────────────────────────────────────────────

@app.route('/training')
@spoc_required
def emp_training():
    plant_id = session['plant_id']
    db = get_db()
    records = db.execute('''
        SELECT t.*, e.name as emp_name, e.designation, e.grade, e.collar,
               e.department, e.section
        FROM emp_training t
        LEFT JOIN employees e ON e.emp_code=t.emp_code AND e.plant_id=t.plant_id
        WHERE t.plant_id=?
        ORDER BY t.id DESC
    ''', (plant_id,)).fetchall()

    emps = db.execute('SELECT emp_code, name FROM employees WHERE plant_id=? AND is_active=1 ORDER BY name', (plant_id,)).fetchall()
    sessions_list = db.execute("SELECT session_code, programme_name FROM calendar WHERE plant_id=? ORDER BY session_code", (plant_id,)).fetchall()
    return render_template('training_2a.html', records=records, employees=emps,
                           sessions=sessions_list, months=MONTHS_FY)

@app.route('/training/add', methods=['POST'])
@spoc_required
def add_emp_training():
    plant_id = session['plant_id']
    f = request.form
    db = get_db()
    emp_code     = f['emp_code']
    session_code = f.get('session_code', '').strip()
    start_date   = f.get('start_date', '')
    end_date     = f.get('end_date', '')

    # Auto-fill from calendar
    prog_name = f.get('programme_name', '').strip()
    prog_type = level = mode = cal_new = ''
    if session_code:
        cal = db.execute('SELECT * FROM calendar WHERE session_code=? AND plant_id=?',
                         (session_code, plant_id)).fetchone()
        if cal:
            prog_name = cal['programme_name']
            prog_type = cal['prog_type']
            level     = cal['level']
            mode      = cal['mode']
            cal_new   = 'Calendar Program'
            if not start_date:
                start_date = cal['plan_start'] or ''
            if not end_date:
                end_date = cal['plan_end'] or ''

    month = _date_to_month(start_date)

    db.execute('''INSERT INTO emp_training
        (plant_id,emp_code,session_code,programme_name,start_date,end_date,
         hrs,prog_type,level,mode,cal_new,pre_rating,post_rating,venue,month)
        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
        (plant_id, emp_code, session_code, prog_name,
         start_date, end_date, float(f.get('hrs') or 0),
         prog_type, level, mode, cal_new,
         _safe_float(f.get('pre_rating')), _safe_float(f.get('post_rating')),
         f.get('venue',''), month))
    db.commit()
    flash('Training record added.', 'success')
    return redirect(url_for('emp_training'))

@app.route('/training/<int:rec_id>/delete', methods=['POST'])
@spoc_required
def delete_emp_training(rec_id):
    db = get_db()
    db.execute('DELETE FROM emp_training WHERE id=? AND plant_id=?', (rec_id, session['plant_id']))
    db.commit()
    flash('Training record deleted.', 'warning')
    return redirect(url_for('emp_training'))

@app.route('/training/template')
@spoc_required
def training_template():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '2A_Bulk_Upload'
    headers = ['Employee Code', 'Session Code (optional)', 'Programme Name',
               'Start Date (YYYY-MM-DD)', 'End Date (YYYY-MM-DD)',
               'Hours', 'Venue', 'Pre-Session Rating (1-5)', 'Post-Session Rating (1-5)']
    hdr_fill = PatternFill('solid', start_color='1F4E79')
    hdr_font = Font(bold=True, color='FFFFFF')
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=i, value=h)
        cell.fill = hdr_fill
        cell.font = hdr_font
        ws.column_dimensions[get_column_letter(i)].width = 26
    samples = [
        ['21700011', 'BCM/EHS/001/B01', 'Fire Safety Training', '2026-06-10', '2026-06-10', 4, 'Training Hall', 3.5, 4.2],
        ['21101568', '', 'MS Office Basics', '2026-07-05', '2026-07-06', 8, 'Computer Lab', '', 4.0],
    ]
    for r, row in enumerate(samples, 2):
        for c, val in enumerate(row, 1):
            ws.cell(row=r, column=c, value=val)
    ws['A5'] = 'NOTE: Session Code is optional. If provided, Programme Name/Type/Mode auto-fill from Calendar.'
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return send_file(out, download_name='2A_Training_Bulk_Upload_Template.xlsx', as_attachment=True,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@app.route('/training/bulk', methods=['POST'])
@spoc_required
def training_bulk_upload():
    plant_id = session['plant_id']
    f = request.files.get('file')
    if not f or f.filename == '':
        flash('No file selected.', 'danger')
        return redirect(url_for('emp_training'))
    try:
        df = _read_upload_file(f)
    except Exception as e:
        flash(f'Could not read file: {e}', 'danger')
        return redirect(url_for('emp_training'))

    db       = get_db()
    inserted = 0
    errors   = []
    for i, row in df.iterrows():
        emp_code     = _clean(row, ['employee code', 'emp code', 'empcode'])
        session_code = _clean(row, ['session code', 'session code (optional)', 'sessioncode'])
        prog_name    = _clean(row, ['programme name', 'program name', 'training name'])
        start_date   = _clean(row, ['start date', 'start date (yyyy-mm-dd)', 'startdate', 'date'])
        end_date     = _clean(row, ['end date', 'end date (yyyy-mm-dd)', 'enddate'])
        hrs          = _safe_float(_clean(row, ['hours', 'hrs', 'duration'])) or 0
        venue        = _clean(row, ['venue'])
        pre_r        = _safe_float(_clean(row, ['pre-session rating', 'pre rating', 'pre_rating']))
        post_r       = _safe_float(_clean(row, ['post-session rating', 'post rating', 'post_rating']))

        if not emp_code:
            errors.append(f'Row {i+2}: Employee Code is required.')
            continue
        emp = db.execute('SELECT 1 FROM employees WHERE emp_code=? AND plant_id=? AND is_active=1',
                         (emp_code, plant_id)).fetchone()
        if not emp:
            errors.append(f'Row {i+2}: Employee {emp_code} not found.')
            continue

        prog_type = level = mode = cal_new = ''
        if session_code:
            cal = db.execute('SELECT * FROM calendar WHERE session_code=? AND plant_id=?',
                             (session_code, plant_id)).fetchone()
            if cal:
                prog_name = prog_name or cal['programme_name']
                prog_type = cal['prog_type']
                level     = cal['level']
                mode      = cal['mode']
                cal_new   = 'Calendar Program'
                start_date = start_date or cal['plan_start'] or ''
                end_date   = end_date or cal['plan_end'] or ''

        if not prog_name:
            errors.append(f'Row {i+2}: Programme Name required (no session code matched).')
            continue

        month = _date_to_month(start_date)
        db.execute('''INSERT INTO emp_training
            (plant_id,emp_code,session_code,programme_name,start_date,end_date,
             hrs,prog_type,level,mode,cal_new,pre_rating,post_rating,venue,month)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
            (plant_id, emp_code, session_code, prog_name,
             start_date, end_date, hrs, prog_type, level, mode, cal_new,
             pre_r, post_r, venue, month))
        inserted += 1
    db.commit()
    if inserted:
        flash(f'Bulk upload complete: {inserted} training records added successfully.', 'success')
    if errors:
        for err in errors:
            flash(err, 'upload_error')
        flash(f'{len(errors)} row(s) had errors — see details below.', 'warning')
    return redirect(url_for('emp_training'))

# ─── PROGRAMME DETAILS (2C) ───────────────────────────────────────────────────

@app.route('/programme')
@spoc_required
def programme_details():
    plant_id = session['plant_id']
    db = get_db()
    records = db.execute('''
        SELECT p.*,
               (SELECT COUNT(*) FROM emp_training t WHERE t.session_code=p.session_code AND t.plant_id=p.plant_id) as participants,
               (SELECT COALESCE(SUM(t.hrs),0) FROM emp_training t WHERE t.session_code=p.session_code AND t.plant_id=p.plant_id) as man_hours
        FROM programme_details p
        WHERE p.plant_id=?
        ORDER BY p.id DESC
    ''', (plant_id,)).fetchall()

    cal_sessions = db.execute(
        "SELECT session_code, programme_name FROM calendar WHERE plant_id=? ORDER BY session_code",
        (plant_id,)).fetchall()
    return render_template('programme_2c.html', records=records,
                           cal_sessions=cal_sessions,
                           int_ext=INT_EXT, audiences=AUDIENCES, months=MONTHS_FY)

@app.route('/programme/add', methods=['POST'])
@spoc_required
def add_programme_details():
    plant_id = session['plant_id']
    f = request.form
    db = get_db()
    session_code = f['session_code'].strip()

    if db.execute('SELECT 1 FROM programme_details WHERE session_code=? AND plant_id=?',
                  (session_code, plant_id)).fetchone():
        flash(f'Session {session_code} already recorded. Edit the existing entry.', 'warning')
        return redirect(url_for('programme_details'))

    cal = db.execute('SELECT * FROM calendar WHERE session_code=? AND plant_id=?',
                     (session_code, plant_id)).fetchone()
    prog_name = cal['programme_name'] if cal else f.get('programme_name','')
    prog_type = cal['prog_type'] if cal else ''
    level     = cal['level'] if cal else ''
    cal_new   = 'Calendar Program' if cal else 'New Program'
    mode      = cal['mode'] if cal else ''
    audience  = cal['target_audience'] if cal else ''

    db.execute('''INSERT INTO programme_details
        (plant_id,session_code,programme_name,prog_type,level,cal_new,mode,
         start_date,end_date,audience,hours_actual,faculty_name,int_ext,cost,
         venue,course_feedback,faculty_feedback,trainer_fb_participants,trainer_fb_facilities)
        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
        (plant_id, session_code, prog_name, prog_type, level, cal_new, mode,
         f.get('start_date',''), f.get('end_date',''), audience,
         float(f.get('hours_actual') or 0), f.get('faculty_name',''),
         f.get('int_ext',''), float(f.get('cost') or 0),
         f.get('venue',''),
         _safe_float(f.get('course_feedback')),
         _safe_float(f.get('faculty_feedback')),
         _safe_float(f.get('trainer_fb_participants')),
         _safe_float(f.get('trainer_fb_facilities'))))
    db.commit()

    # Update calendar status
    act_date = f.get('start_date','')
    db.execute("UPDATE calendar SET status='Conducted', actual_date=? WHERE session_code=? AND plant_id=?",
               (act_date, session_code, plant_id))
    db.commit()
    flash(f'Programme {session_code} details saved.', 'success')
    return redirect(url_for('programme_details'))

@app.route('/programme/<int:rec_id>/delete', methods=['POST'])
@spoc_required
def delete_programme(rec_id):
    db = get_db()
    rec = db.execute('SELECT session_code FROM programme_details WHERE id=? AND plant_id=?',
                     (rec_id, session['plant_id'])).fetchone()
    if rec:
        db.execute('DELETE FROM programme_details WHERE id=? AND plant_id=?', (rec_id, session['plant_id']))
        db.execute("UPDATE calendar SET status='To Be Planned', actual_date=NULL WHERE session_code=? AND plant_id=?",
                   (rec['session_code'], session['plant_id']))
        db.commit()
    flash('Programme record deleted.', 'warning')
    return redirect(url_for('programme_details'))

# ─── MONTHLY SUMMARY ──────────────────────────────────────────────────────────

@app.route('/summary')
@spoc_required
def monthly_summary():
    plant_id     = session['plant_id']
    sel_month    = request.args.get('month', '')
    db           = get_db()
    summary_rows = _calc_summary(plant_id, sel_month, db)
    totals       = _calc_totals(summary_rows)
    compliance   = _calc_compliance(plant_id, db)
    return render_template('summary.html', summary_rows=summary_rows,
                           totals=totals, compliance=compliance,
                           months=MONTHS_FY, selected_month=sel_month,
                           prog_types=PROG_TYPES)

# ─── CENTRAL DASHBOARD ────────────────────────────────────────────────────────

@app.route('/central')
@central_required
def central_dashboard():
    db = get_db()
    plant_summaries = []
    for p in PLANTS:
        pid = p['id']
        bc  = db.execute("SELECT COUNT(*) FROM employees WHERE plant_id=? AND is_active=1 AND collar='Blue Collared'", (pid,)).fetchone()[0]
        wc  = db.execute("SELECT COUNT(*) FROM employees WHERE plant_id=? AND is_active=1 AND collar='White Collared'", (pid,)).fetchone()[0]
        sessions_cnt  = db.execute("SELECT COUNT(*) FROM calendar WHERE plant_id=?", (pid,)).fetchone()[0]
        conducted_cnt = db.execute("SELECT COUNT(*) FROM calendar WHERE plant_id=? AND status='Conducted'", (pid,)).fetchone()[0]
        manhours = db.execute("SELECT COALESCE(SUM(hrs),0) FROM emp_training WHERE plant_id=?", (pid,)).fetchone()[0]
        bc_hrs   = db.execute("SELECT COALESCE(SUM(t.hrs),0) FROM emp_training t JOIN employees e ON e.emp_code=t.emp_code AND e.plant_id=t.plant_id WHERE t.plant_id=? AND e.collar='Blue Collared'", (pid,)).fetchone()[0]
        wc_hrs   = db.execute("SELECT COALESCE(SUM(t.hrs),0) FROM emp_training t JOIN employees e ON e.emp_code=t.emp_code AND e.plant_id=t.plant_id WHERE t.plant_id=? AND e.collar='White Collared'", (pid,)).fetchone()[0]
        bc_mandate = bc * 12
        wc_mandate = wc * 24
        bc_pct = round((bc_hrs / bc_mandate * 100), 1) if bc_mandate else 0
        wc_pct = round((wc_hrs / wc_mandate * 100), 1) if wc_mandate else 0
        plant_summaries.append({**p,
            'blue_collar': bc, 'white_collar': wc, 'total_emp': bc + wc,
            'sessions': sessions_cnt, 'conducted': conducted_cnt,
            'manhours': round(manhours, 1),
            'bc_pct': bc_pct, 'wc_pct': wc_pct
        })
    grand = {
        'total_emp': sum(p['total_emp'] for p in plant_summaries),
        'manhours':  round(sum(p['manhours'] for p in plant_summaries), 1),
        'sessions':  sum(p['sessions'] for p in plant_summaries),
        'conducted': sum(p['conducted'] for p in plant_summaries),
    }
    return render_template('central.html', plants=plant_summaries, grand=grand)

@app.route('/central/plant/<int:plant_id>')
@central_required
def central_plant_view(plant_id):
    if plant_id not in PLANT_MAP:
        flash('Plant not found.', 'danger')
        return redirect(url_for('central_dashboard'))
    plant     = PLANT_MAP[plant_id]
    db        = get_db()
    sel_month = request.args.get('month', '')
    summary_rows = _calc_summary(plant_id, sel_month, db)
    totals       = _calc_totals(summary_rows)
    compliance   = _calc_compliance(plant_id, db)
    return render_template('central_plant.html', plant=plant,
                           summary_rows=summary_rows, totals=totals,
                           compliance=compliance, months=MONTHS_FY,
                           selected_month=sel_month)

# ─── EXPORT TO EXCEL ──────────────────────────────────────────────────────────

@app.route('/export/<int:plant_id>')
@login_required
def export_excel(plant_id):
    if session.get('role') == 'spoc' and session.get('plant_id') != plant_id:
        flash('Access denied.', 'danger')
        return redirect(url_for('spoc_dashboard'))
    plant = PLANT_MAP.get(plant_id)
    if not plant:
        flash('Plant not found.', 'danger')
        return redirect(url_for('index'))

    db   = get_db()
    wb   = openpyxl.Workbook()
    fy   = '2026-27'
    hdr  = Font(bold=True, color='FFFFFF')
    hdr_fill = PatternFill('solid', start_color='1F4E79')
    sub_fill = PatternFill('solid', start_color='2E75B6')
    thin = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin'))

    def style_header(ws, row, cols):
        for c in range(1, cols + 1):
            cell = ws.cell(row=row, column=c)
            cell.font = hdr
            cell.fill = hdr_fill
            cell.border = thin
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    def set_title(ws, title):
        ws['A1'] = 'BALRAMPUR CHINI MILLS LIMITED'
        ws['A1'].font = Font(bold=True, size=13)
        ws['A2'] = title
        ws['A2'].font = Font(bold=True, size=11)

    # Sheet 1 – Employee Master
    ws = wb.active
    ws.title = 'EMP_MASTER'
    set_title(ws, f"{plant['name'].upper()}  —  EMPLOYEE MASTER  |  FY {fy}")
    headers = ['Sr.','Emp Code','Name','Designation','Grade','Collar','Department','Section','Category','Gender','PH','Exit Date','Exit Reason','Remarks']
    for i, h in enumerate(headers, 1):
        ws.cell(row=4, column=i, value=h)
    style_header(ws, 4, len(headers))
    emps = db.execute('SELECT * FROM employees WHERE plant_id=? ORDER BY name', (plant_id,)).fetchall()
    for r, emp in enumerate(emps, 5):
        row_data = [r-4, emp['emp_code'], emp['name'], emp['designation'], emp['grade'],
                    emp['collar'], emp['department'], emp['section'], emp['category'],
                    emp['gender'], emp['physically_handicapped'],
                    emp['exit_date'] or '', emp['exit_reason'] or '', emp['remarks'] or '']
        for c, val in enumerate(row_data, 1):
            cell = ws.cell(row=r, column=c, value=val)
            cell.border = thin
    for col in ws.columns:
        ws.column_dimensions[get_column_letter(col[0].column)].width = 18

    # Sheet 2 – TNI
    ws2 = wb.create_sheet('TNI_Tracking')
    set_title(ws2, f"{plant['name'].upper()}  —  TNI TRACKING  |  FY {fy}")
    headers2 = ['Sr.','Emp Code','Name','Designation','Grade','Collar','Dept','Section','Programme Name','Type','Mode','Target Month','Planned Hrs','Completed?']
    for i, h in enumerate(headers2, 1):
        ws2.cell(row=4, column=i, value=h)
    style_header(ws2, 4, len(headers2))
    tni_recs = db.execute('''SELECT t.*,e.name,e.designation,e.grade,e.collar,e.department,e.section
        FROM tni t LEFT JOIN employees e ON e.emp_code=t.emp_code AND e.plant_id=t.plant_id
        WHERE t.plant_id=?''', (plant_id,)).fetchall()
    for r, rec in enumerate(tni_recs, 5):
        done = db.execute('SELECT 1 FROM emp_training WHERE plant_id=? AND emp_code=? AND programme_name=?',
                          (plant_id, rec['emp_code'], rec['programme_name'])).fetchone()
        row_data = [r-4, rec['emp_code'], rec['name'] or '', rec['designation'] or '',
                    rec['grade'] or '', rec['collar'] or '', rec['department'] or '',
                    rec['section'] or '', rec['programme_name'], rec['prog_type'] or '',
                    rec['mode'] or '', rec['target_month'] or '', rec['planned_hours'],
                    'Yes' if done else 'No']
        for c, val in enumerate(row_data, 1):
            cell = ws2.cell(row=r, column=c, value=val)
            cell.border = thin

    # Sheet 3 – Calendar
    ws3 = wb.create_sheet('Cal_Plan_vs_Actual')
    set_title(ws3, f"{plant['name'].upper()}  —  TRAINING CALENDAR  |  FY {fy}")
    headers3 = ['S/N','PROG CODE','SESSION CODE','Source','Programme Name','Type','Planned Month',
                'Plan Start','Plan End','Time From','Time To','Duration(Hrs)','Level','Mode',
                'Target Audience','Planned Pax','Trainer/Vendor','STATUS','Actual Date','Actual Pax']
    for i, h in enumerate(headers3, 1):
        ws3.cell(row=4, column=i, value=h)
    style_header(ws3, 4, len(headers3))
    cal_recs = db.execute('SELECT * FROM calendar WHERE plant_id=? ORDER BY id', (plant_id,)).fetchall()
    for r, rec in enumerate(cal_recs, 5):
        pd2c = db.execute('SELECT start_date, COUNT(*) as pax FROM programme_details WHERE session_code=? AND plant_id=?',
                          (rec['session_code'], plant_id)).fetchone()
        act_pax = db.execute('SELECT COUNT(*) FROM emp_training WHERE session_code=? AND plant_id=?',
                             (rec['session_code'], plant_id)).fetchone()[0]
        row_data = [r-4, rec['prog_code'], rec['session_code'], rec['source'],
                    rec['programme_name'], rec['prog_type'], rec['planned_month'],
                    rec['plan_start'], rec['plan_end'], rec['time_from'], rec['time_to'],
                    rec['duration_hrs'], rec['level'], rec['mode'], rec['target_audience'],
                    rec['planned_pax'], rec['trainer_vendor'], rec['status'],
                    pd2c['start_date'] if pd2c and pd2c['start_date'] else '',
                    act_pax]
        for c, val in enumerate(row_data, 1):
            cell = ws3.cell(row=r, column=c, value=val)
            cell.border = thin

    # Sheet 4 – 2A Employee Training
    ws4 = wb.create_sheet('2A_Emp_Training')
    set_title(ws4, f"{plant['name'].upper()}  —  2A: EMPLOYEE TRAINING DETAILS  |  FY {fy}")
    headers4 = ['Sr.','Emp Code','Name','Designation','Grade','Collar','Dept','Section',
                'Start Date','End Date','Hrs','Programme Name','Type','Level','Mode',
                'Cal/New','Pre Rating','Post Rating','Venue','Month']
    for i, h in enumerate(headers4, 1):
        ws4.cell(row=4, column=i, value=h)
    style_header(ws4, 4, len(headers4))
    trng = db.execute('''SELECT t.*,e.name as emp_name,e.designation,e.grade,e.collar,e.department,e.section
        FROM emp_training t LEFT JOIN employees e ON e.emp_code=t.emp_code AND e.plant_id=t.plant_id
        WHERE t.plant_id=? ORDER BY t.id''', (plant_id,)).fetchall()
    for r, rec in enumerate(trng, 5):
        row_data = [r-4, rec['emp_code'], rec['emp_name'] or '', rec['designation'] or '',
                    rec['grade'] or '', rec['collar'] or '', rec['department'] or '',
                    rec['section'] or '', rec['start_date'], rec['end_date'],
                    rec['hrs'], rec['programme_name'], rec['prog_type'] or '',
                    rec['level'] or '', rec['mode'] or '', rec['cal_new'] or '',
                    rec['pre_rating'], rec['post_rating'], rec['venue'] or '', rec['month'] or '']
        for c, val in enumerate(row_data, 1):
            cell = ws4.cell(row=r, column=c, value=val)
            cell.border = thin

    # Sheet 5 – 2C Programme Details
    ws5 = wb.create_sheet('2C_Programme')
    set_title(ws5, f"{plant['name'].upper()}  —  2C: PROGRAMME-WISE DETAILS  |  FY {fy}")
    headers5 = ['Sr.','Session Code','Programme Name','Type','Level','Cal/New','Mode',
                'Start Date','End Date','Audience','Hours Actual','Faculty Name','Int/Ext',
                'Cost (Rs.)','Venue','Course FB','Faculty FB','Trainer FB-Participants',
                'Trainer FB-Facilities','Participants','Man-Hours']
    for i, h in enumerate(headers5, 1):
        ws5.cell(row=4, column=i, value=h)
    style_header(ws5, 4, len(headers5))
    prog_recs = db.execute('SELECT * FROM programme_details WHERE plant_id=? ORDER BY id', (plant_id,)).fetchall()
    for r, rec in enumerate(prog_recs, 5):
        participants = db.execute('SELECT COUNT(*) FROM emp_training WHERE session_code=? AND plant_id=?',
                                  (rec['session_code'], plant_id)).fetchone()[0]
        man_hours = db.execute('SELECT COALESCE(SUM(hrs),0) FROM emp_training WHERE session_code=? AND plant_id=?',
                               (rec['session_code'], plant_id)).fetchone()[0]
        row_data = [r-4, rec['session_code'], rec['programme_name'], rec['prog_type'],
                    rec['level'], rec['cal_new'], rec['mode'],
                    rec['start_date'], rec['end_date'], rec['audience'],
                    rec['hours_actual'], rec['faculty_name'], rec['int_ext'],
                    rec['cost'], rec['venue'], rec['course_feedback'],
                    rec['faculty_feedback'], rec['trainer_fb_participants'],
                    rec['trainer_fb_facilities'], participants, round(man_hours, 1)]
        for c, val in enumerate(row_data, 1):
            cell = ws5.cell(row=r, column=c, value=val)
            cell.border = thin

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    filename = f"BCML_{plant['unit_code']}_Training_MIS_{fy.replace('-','')}.xlsx"
    return send_file(output, download_name=filename,
                     as_attachment=True,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

# ─── API (AJAX auto-fill) ─────────────────────────────────────────────────────

@app.route('/api/employee/<emp_code>')
@login_required
def api_employee(emp_code):
    plant_id = session.get('plant_id')
    if not plant_id:
        return jsonify({})
    db = get_db()
    emp = db.execute('SELECT * FROM employees WHERE emp_code=? AND plant_id=? AND is_active=1',
                     (emp_code.strip(), plant_id)).fetchone()
    if not emp:
        return jsonify({})
    return jsonify({
        'name': emp['name'], 'designation': emp['designation'],
        'grade': emp['grade'], 'collar': emp['collar'],
        'department': emp['department'], 'section': emp['section'],
        'category': emp['category'], 'gender': emp['gender']
    })

@app.route('/api/session/<session_code>')
@login_required
def api_session(session_code):
    plant_id = session.get('plant_id')
    if not plant_id:
        return jsonify({})
    db = get_db()
    cal = db.execute('SELECT * FROM calendar WHERE session_code=? AND plant_id=?',
                     (session_code.strip(), plant_id)).fetchone()
    if not cal:
        return jsonify({})
    return jsonify({
        'programme_name': cal['programme_name'], 'prog_type': cal['prog_type'],
        'level': cal['level'], 'mode': cal['mode'],
        'plan_start': cal['plan_start'], 'plan_end': cal['plan_end'],
        'duration_hrs': cal['duration_hrs'], 'target_audience': cal['target_audience']
    })

@app.route('/api/employees_list')
@login_required
def api_employees_list():
    plant_id = session.get('plant_id')
    if not plant_id:
        return jsonify([])
    db = get_db()
    emps = db.execute('SELECT emp_code, name FROM employees WHERE plant_id=? AND is_active=1 ORDER BY name',
                      (plant_id,)).fetchall()
    return jsonify([{'code': e['emp_code'], 'name': e['name']} for e in emps])

# ─── HELPERS ─────────────────────────────────────────────────────────────────

def _read_upload_file(file_storage):
    import pandas as pd
    fname = file_storage.filename.lower()
    if fname.endswith('.csv'):
        return pd.read_csv(file_storage, dtype=str).fillna('')
    else:
        return pd.read_excel(file_storage, dtype=str).fillna('')

def _clean(row, keys):
    for k in keys:
        for col in row.index:
            if str(col).strip().lower() == k:
                val = str(row[col]).strip()
                return '' if val.lower() in ('nan', 'none', '') else val
    return ''

def normalise_collar(val):
    v = str(val).strip().upper()
    if any(x in v for x in ['WHITE', 'WC', 'W C']):
        return 'White Collared'
    if any(x in v for x in ['BLUE', 'BC', 'B C']):
        return 'Blue Collared'
    return val.strip()

def _safe_float(val):
    try:
        return float(val) if val and str(val).strip() != '' else None
    except (ValueError, TypeError):
        return None

def _date_to_month(date_str):
    if not date_str:
        return ''
    try:
        d = datetime.strptime(str(date_str)[:10], '%Y-%m-%d')
        return d.strftime('%B')
    except Exception:
        return ''

def _get_programme_names(plant_id, db):
    rows = db.execute('SELECT DISTINCT programme_name FROM calendar WHERE plant_id=? ORDER BY programme_name', (plant_id,)).fetchall()
    return [r['programme_name'] for r in rows]

def _get_or_create_prog_code(plant_id, prog_name, prog_type, db):
    existing = db.execute('SELECT prog_code FROM calendar WHERE plant_id=? AND programme_name=? LIMIT 1',
                          (plant_id, prog_name)).fetchone()
    if existing:
        return existing['prog_code']
    unit_code = PLANT_MAP[plant_id]['unit_code']
    abbrev    = TYPE_ABBREV.get(prog_type, 'GEN')
    count     = db.execute("SELECT COUNT(DISTINCT prog_code) FROM calendar WHERE plant_id=? AND prog_code LIKE ?",
                           (plant_id, f'{unit_code}/{abbrev}/%')).fetchone()[0]
    return f'{unit_code}/{abbrev}/{count+1:03d}'

def _new_session_code(plant_id, prog_code, db):
    count = db.execute('SELECT COUNT(*) FROM calendar WHERE plant_id=? AND prog_code=?',
                       (plant_id, prog_code)).fetchone()[0]
    return f'{prog_code}/B{count+1:02d}'

def _sync_calendar_from_2c(plant_id, db):
    db.execute('''UPDATE calendar SET status='Conducted'
        WHERE plant_id=? AND session_code IN
        (SELECT session_code FROM programme_details WHERE plant_id=?)
        AND status='To Be Planned' ''', (plant_id, plant_id))
    db.commit()

def _calc_summary(plant_id, month_filter, db):
    rows = []
    for pt in PROG_TYPES:
        clause = "AND t.month=?" if month_filter else ""
        params_bc = [plant_id, pt, 'Blue Collared'] + ([month_filter] if month_filter else [])
        params_wc = [plant_id, pt, 'White Collared'] + ([month_filter] if month_filter else [])
        params_all = [plant_id, pt] + ([month_filter] if month_filter else [])

        bc_progs = db.execute(f'''SELECT COUNT(DISTINCT t.session_code) FROM emp_training t
            JOIN employees e ON e.emp_code=t.emp_code AND e.plant_id=t.plant_id
            WHERE t.plant_id=? AND t.prog_type=? AND e.collar=? {clause}''', params_bc).fetchone()[0]
        wc_progs = db.execute(f'''SELECT COUNT(DISTINCT t.session_code) FROM emp_training t
            JOIN employees e ON e.emp_code=t.emp_code AND e.plant_id=t.plant_id
            WHERE t.plant_id=? AND t.prog_type=? AND e.collar=? {clause}''', params_wc).fetchone()[0]
        bc_seats = db.execute(f'''SELECT COUNT(*) FROM emp_training t
            JOIN employees e ON e.emp_code=t.emp_code AND e.plant_id=t.plant_id
            WHERE t.plant_id=? AND t.prog_type=? AND e.collar=? {clause}''', params_bc).fetchone()[0]
        wc_seats = db.execute(f'''SELECT COUNT(*) FROM emp_training t
            JOIN employees e ON e.emp_code=t.emp_code AND e.plant_id=t.plant_id
            WHERE t.plant_id=? AND t.prog_type=? AND e.collar=? {clause}''', params_wc).fetchone()[0]
        bc_hrs = db.execute(f'''SELECT COALESCE(SUM(t.hrs),0) FROM emp_training t
            JOIN employees e ON e.emp_code=t.emp_code AND e.plant_id=t.plant_id
            WHERE t.plant_id=? AND t.prog_type=? AND e.collar=? {clause}''', params_bc).fetchone()[0]
        wc_hrs = db.execute(f'''SELECT COALESCE(SUM(t.hrs),0) FROM emp_training t
            JOIN employees e ON e.emp_code=t.emp_code AND e.plant_id=t.plant_id
            WHERE t.plant_id=? AND t.prog_type=? AND e.collar=? {clause}''', params_wc).fetchone()[0]
        int_prog = db.execute(f'''SELECT COUNT(DISTINCT p.session_code) FROM programme_details p
            WHERE p.plant_id=? AND p.prog_type=? AND p.int_ext='Internal' {clause.replace('t.month','p.created_at')}''',
            [plant_id, pt] + ([month_filter] if month_filter else [])).fetchone()[0]
        ext_prog = db.execute(f'''SELECT COUNT(DISTINCT p.session_code) FROM programme_details p
            WHERE p.plant_id=? AND p.prog_type=? AND p.int_ext='External' {clause.replace('t.month','p.created_at')}''',
            [plant_id, pt] + ([month_filter] if month_filter else [])).fetchone()[0]
        rows.append({
            'prog_type': pt,
            'bc_progs': bc_progs, 'wc_progs': wc_progs,
            'int_prog': int_prog, 'ext_prog': ext_prog,
            'total_prog': bc_progs + wc_progs,
            'bc_seats': bc_seats, 'wc_seats': wc_seats,
            'total_seats': bc_seats + wc_seats,
            'bc_hrs': round(bc_hrs, 1), 'wc_hrs': round(wc_hrs, 1),
            'total_hrs': round(bc_hrs + wc_hrs, 1)
        })
    return rows

def _calc_totals(rows):
    t = {k: 0 for k in rows[0]} if rows else {}
    t['prog_type'] = 'TOTAL'
    for r in rows:
        for k, v in r.items():
            if k != 'prog_type':
                t[k] = round(t.get(k, 0) + (v or 0), 1)
    return t

def _calc_compliance(plant_id, db):
    bc = db.execute("SELECT COUNT(*) FROM employees WHERE plant_id=? AND is_active=1 AND collar='Blue Collared'", (plant_id,)).fetchone()[0]
    wc = db.execute("SELECT COUNT(*) FROM employees WHERE plant_id=? AND is_active=1 AND collar='White Collared'", (plant_id,)).fetchone()[0]
    bc_act = db.execute('''SELECT COALESCE(SUM(t.hrs),0) FROM emp_training t
        JOIN employees e ON e.emp_code=t.emp_code AND e.plant_id=t.plant_id
        WHERE t.plant_id=? AND e.collar='Blue Collared' ''', (plant_id,)).fetchone()[0]
    wc_act = db.execute('''SELECT COALESCE(SUM(t.hrs),0) FROM emp_training t
        JOIN employees e ON e.emp_code=t.emp_code AND e.plant_id=t.plant_id
        WHERE t.plant_id=? AND e.collar='White Collared' ''', (plant_id,)).fetchone()[0]
    bc_mandate = bc * 12
    wc_mandate = wc * 24
    return {
        'bc_emp': bc, 'wc_emp': wc,
        'bc_mandate': bc_mandate, 'wc_mandate': wc_mandate,
        'bc_actual': round(bc_act, 1), 'wc_actual': round(wc_act, 1),
        'bc_pct': round(bc_act / bc_mandate * 100, 1) if bc_mandate else 0,
        'wc_pct': round(wc_act / wc_mandate * 100, 1) if wc_mandate else 0,
        'total_pct': round((bc_act + wc_act) / (bc_mandate + wc_mandate) * 100, 1) if (bc_mandate + wc_mandate) else 0
    }

# ─── ENTRY POINT ─────────────────────────────────────────────────────────────

if __name__ == '__main__':
    init_db()
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
