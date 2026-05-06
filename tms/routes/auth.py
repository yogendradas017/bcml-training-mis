from datetime import date
from flask import render_template, request, redirect, url_for, session, flash
from werkzeug.security import check_password_hash

from tms.constants import PLANT_MAP
from tms.db import get_db
from tms.decorators import spoc_required


def _register(app):

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

    @app.route('/admin/tni-archives')
    def admin_tni_archives():
        if session.get('role') != 'admin':
            return redirect(url_for('index'))
        db = get_db()
        archives = db.execute('''
            SELECT a.archive_token, a.archived_at, a.plant_id,
                   p.name AS plant_name, a.fy_year, COUNT(*) AS row_count
            FROM tni_archive a
            JOIN plants p ON p.id = a.plant_id
            GROUP BY a.archive_token
            ORDER BY a.archived_at DESC
        ''').fetchall()
        return render_template('admin_tni_archives.html', archives=archives)

    @app.route('/admin/tni-archives/restore', methods=['POST'])
    def admin_tni_restore():
        if session.get('role') != 'admin':
            return redirect(url_for('index'))
        token = request.form.get('token', '').strip()
        if not token:
            flash('No archive token provided.', 'danger')
            return redirect(url_for('admin_tni_archives'))
        db = get_db()
        meta = db.execute(
            'SELECT plant_id, fy_year FROM tni_archive WHERE archive_token=? LIMIT 1', (token,)
        ).fetchone()
        if not meta:
            flash('Archive not found.', 'danger')
            return redirect(url_for('admin_tni_archives'))
        plant_id = meta['plant_id']
        fy_year  = meta['fy_year']
        db.execute('DELETE FROM tni WHERE plant_id=? AND fy_year=?', (plant_id, fy_year))
        db.execute('''
            INSERT OR IGNORE INTO tni
                (plant_id, emp_code, programme_name, prog_type, mode,
                 target_month, planned_hours, source, fy_year)
            SELECT plant_id, emp_code, programme_name, prog_type, mode,
                   target_month, planned_hours, source, fy_year
            FROM tni_archive WHERE archive_token=?
        ''', (token,))
        restored = db.execute(
            'SELECT COUNT(*) FROM tni WHERE plant_id=? AND fy_year=?', (plant_id, fy_year)
        ).fetchone()[0]
        from tms.helpers import _sync_master_from_tni
        _sync_master_from_tni(plant_id, db)
        db.commit()
        plant_name = db.execute('SELECT name FROM plants WHERE id=?', (plant_id,)).fetchone()['name']
        flash(f'Restored {restored} TNI rows for {plant_name} ({fy_year}). Programme master rebuilt.', 'success')
        return redirect(url_for('admin_tni_archives'))

    @app.route('/dashboard')
    @spoc_required
    def spoc_dashboard():
        plant_id = session['plant_id']
        db = get_db()
        # Distinct central-hosted sessions attended by this plant's employees
        # (counted in addition to plant's own calendar sessions, since they're
        # also real training events that consumed plant manhours).
        central_attended = db.execute(
            "SELECT COUNT(DISTINCT session_code) FROM emp_training "
            "WHERE plant_id=? AND host_plant_id=99 AND session_code IS NOT NULL AND session_code!=''",
            (plant_id,)).fetchone()[0]
        stats = {
            'total_emp':    db.execute('SELECT COUNT(*) FROM employees WHERE plant_id=? AND is_active=1', (plant_id,)).fetchone()[0],
            'blue_collar':  db.execute("SELECT COUNT(*) FROM employees WHERE plant_id=? AND is_active=1 AND collar='Blue Collared'", (plant_id,)).fetchone()[0],
            'white_collar': db.execute("SELECT COUNT(*) FROM employees WHERE plant_id=? AND is_active=1 AND collar='White Collared'", (plant_id,)).fetchone()[0],
            'tni_count':    db.execute('SELECT COUNT(DISTINCT emp_code || "|" || programme_name) FROM tni WHERE plant_id=?', (plant_id,)).fetchone()[0],
            'sessions':     db.execute('SELECT COUNT(*) FROM calendar WHERE plant_id=?', (plant_id,)).fetchone()[0] + central_attended,
            'conducted':    db.execute("SELECT COUNT(*) FROM calendar WHERE plant_id=? AND status='Conducted'", (plant_id,)).fetchone()[0] + central_attended,
            'trainings':    db.execute('SELECT COUNT(*) FROM emp_training WHERE plant_id=?', (plant_id,)).fetchone()[0],
            'manhours':     db.execute('SELECT COALESCE(SUM(hrs),0) FROM emp_training WHERE plant_id=?', (plant_id,)).fetchone()[0],
        }
        return render_template('dashboard.html', stats=stats)
