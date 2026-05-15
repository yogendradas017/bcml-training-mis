import io
import os
import shutil
from datetime import date, datetime, timedelta
from flask import render_template, request, redirect, url_for, session, flash, send_file
from werkzeug.security import check_password_hash, generate_password_hash
import pyotp
import qrcode
import base64

from tms.constants import PLANT_MAP, DB_PATH
from tms.db import get_db
from tms.decorators import spoc_required, login_required, admin_required
from tms.helpers import _current_fy
from tms.audit import log_action


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
            username = request.form.get('username', '').strip().lower()
            password = request.form.get('password', '')
            if not username or not password:
                flash('Username and password are required.', 'danger')
                return render_template('login.html')
            db = get_db()
            user = db.execute('SELECT * FROM users WHERE username=?', (username,)).fetchone()

            # Account lockout check
            if user and user['locked_until']:
                try:
                    if datetime.now() < datetime.fromisoformat(user['locked_until']):
                        flash('Account locked. Try again in 15 minutes.', 'danger')
                        log_action('LOGIN_FAIL', f'locked:{username}', username=username)
                        return render_template('login.html')
                    else:
                        db.execute('UPDATE users SET failed_attempts=0, locked_until=NULL WHERE id=?', (user['id'],))
                        db.commit()
                except (ValueError, TypeError):
                    db.execute('UPDATE users SET failed_attempts=0, locked_until=NULL WHERE id=?', (user['id'],))
                    db.commit()

            if user and check_password_hash(user['password'], password):
                db.execute('UPDATE users SET failed_attempts=0, locked_until=NULL WHERE id=?', (user['id'],))
                db.commit()
                # If 2FA enabled → hold in pending state, redirect to TOTP verify
                if user['totp_enabled'] and user['totp_secret']:
                    session.clear()
                    session['2fa_uid']  = user['id']
                    session['2fa_next'] = 'change_password' if user['must_change_password'] else (
                        'central_dashboard' if user['role'] in ('central', 'admin') else 'spoc_dashboard'
                    )
                    return redirect(url_for('login_2fa'))
                session.clear()
                session['user_id']  = user['id']
                session['username'] = user['username']
                session['role']     = user['role']
                session['plant_id'] = user['plant_id']
                if user['plant_id']:
                    session['plant_name'] = PLANT_MAP[user['plant_id']]['name']
                    session['unit_code']  = PLANT_MAP[user['plant_id']]['unit_code']
                log_action('LOGIN_OK', f"role:{user['role']}", user_id=user['id'],
                           username=user['username'], plant_id=user['plant_id'])
                if user['must_change_password']:
                    flash('Please set a new password before continuing.', 'warning')
                    return redirect(url_for('change_password'))
                if user['role'] in ('central', 'admin'):
                    return redirect(url_for('central_dashboard'))
                return redirect(url_for('spoc_dashboard'))
            else:
                if user:
                    new_count = (user['failed_attempts'] or 0) + 1
                    locked_until = None
                    if new_count >= 5:
                        locked_until = (datetime.now() + timedelta(minutes=15)).isoformat()
                        flash('Too many failed attempts. Account locked for 15 minutes.', 'danger')
                        log_action('ACCOUNT_LOCKED', f'user:{username}', username=username)
                    else:
                        remaining = 5 - new_count
                        flash(f'Invalid username or password. {remaining} attempt(s) before lockout.', 'danger')
                    db.execute('UPDATE users SET failed_attempts=?, locked_until=? WHERE id=?',
                               (new_count, locked_until, user['id']))
                    db.commit()
                else:
                    flash('Invalid username or password.', 'danger')
                log_action('LOGIN_FAIL', f'username:{username}', username=username)
        return render_template('login.html')

    @app.route('/login/2fa', methods=['GET', 'POST'])
    def login_2fa():
        uid = session.get('2fa_uid')
        if not uid:
            return redirect(url_for('login'))
        if request.method == 'POST':
            code = request.form.get('totp_code', '').strip().replace(' ', '')
            db = get_db()
            user = db.execute('SELECT * FROM users WHERE id=?', (uid,)).fetchone()
            if not user or not user['totp_secret']:
                session.clear()
                flash('Session error. Please log in again.', 'danger')
                return redirect(url_for('login'))
            totp = pyotp.TOTP(user['totp_secret'])
            if totp.verify(code, valid_window=1):
                next_ep = session.pop('2fa_next', 'spoc_dashboard')
                session.clear()
                session['user_id']  = user['id']
                session['username'] = user['username']
                session['role']     = user['role']
                session['plant_id'] = user['plant_id']
                if user['plant_id'] and user['plant_id'] in PLANT_MAP:
                    session['plant_name'] = PLANT_MAP[user['plant_id']]['name']
                    session['unit_code']  = PLANT_MAP[user['plant_id']]['unit_code']
                log_action('LOGIN_OK', f"2fa:role:{user['role']}", user_id=user['id'],
                           username=user['username'], plant_id=user['plant_id'])
                if user['must_change_password']:
                    flash('Please set a new password before continuing.', 'warning')
                    return redirect(url_for('change_password'))
                return redirect(url_for(next_ep))
            else:
                log_action('LOGIN_FAIL', f'2fa_wrong_code:{user["username"]}', username=user['username'])
                flash('Invalid or expired code. Try again.', 'danger')
        return render_template('login_2fa.html')

    @app.route('/logout')
    def logout():
        log_action('LOGOUT')
        session.clear()
        return redirect(url_for('login'))

    @app.route('/admin/users')
    @admin_required
    def admin_users():
        db = get_db()
        rows = db.execute('''
            SELECT u.id, u.username, u.role, u.must_change_password,
                   u.failed_attempts, u.locked_until,
                   p.name AS plant_name
            FROM users u
            LEFT JOIN plants p ON p.id = u.plant_id
            ORDER BY u.role, u.username
        ''').fetchall()
        return render_template('admin_users.html', users=rows)

    @app.route('/admin/audit-log')
    @admin_required
    def admin_audit_log():
        db = get_db()
        q       = request.args.get('q', '').strip()
        action  = request.args.get('action', '').strip()
        filters = ['1=1']
        params  = []
        if q:
            filters.append('(username LIKE ? OR detail LIKE ? OR ip_address LIKE ?)')
            params += [f'%{q}%', f'%{q}%', f'%{q}%']
        if action:
            filters.append('action=?')
            params.append(action)
        where = ' AND '.join(filters)
        logs = db.execute(
            f'SELECT * FROM audit_log WHERE {where} ORDER BY ts DESC LIMIT 500',
            params
        ).fetchall()
        actions = [r[0] for r in db.execute(
            'SELECT DISTINCT action FROM audit_log ORDER BY action'
        ).fetchall()]
        return render_template('admin_audit_log.html', logs=logs, actions=actions,
                               q=q, sel_action=action)

    @app.route('/admin/2fa/setup/<int:user_id>')
    @admin_required
    def admin_2fa_setup(user_id):
        db = get_db()
        user = db.execute('SELECT * FROM users WHERE id=?', (user_id,)).fetchone()
        if not user:
            flash('User not found.', 'danger')
            return redirect(url_for('admin_users'))
        secret = user['totp_secret'] or pyotp.random_base32()
        if not user['totp_secret']:
            db.execute('UPDATE users SET totp_secret=? WHERE id=?', (secret, user_id))
            db.commit()
        uri = pyotp.TOTP(secret).provisioning_uri(
            name=user['username'],
            issuer_name='BCML TMS'
        )
        img = qrcode.make(uri)
        buf = io.BytesIO()
        img.save(buf, format='PNG')
        qr_b64 = base64.b64encode(buf.getvalue()).decode()
        return render_template('admin_2fa_setup.html', user=user, secret=secret, qr_b64=qr_b64)

    @app.route('/admin/2fa/enable/<int:user_id>', methods=['POST'])
    @admin_required
    def admin_2fa_enable(user_id):
        code = request.form.get('totp_code', '').strip()
        db = get_db()
        user = db.execute('SELECT * FROM users WHERE id=?', (user_id,)).fetchone()
        if not user or not user['totp_secret']:
            flash('Setup 2FA first.', 'danger')
            return redirect(url_for('admin_users'))
        if pyotp.TOTP(user['totp_secret']).verify(code, valid_window=1):
            db.execute('UPDATE users SET totp_enabled=1 WHERE id=?', (user_id,))
            db.commit()
            log_action('RECORD_EDIT', f'2fa_enabled:user:{user["username"]}')
            flash(f"2FA enabled for '{user['username']}'.", 'success')
        else:
            flash('Invalid code — scan QR again and retry.', 'danger')
            return redirect(url_for('admin_2fa_setup', user_id=user_id))
        return redirect(url_for('admin_users'))

    @app.route('/admin/2fa/disable/<int:user_id>', methods=['POST'])
    @admin_required
    def admin_2fa_disable(user_id):
        db = get_db()
        user = db.execute('SELECT username FROM users WHERE id=?', (user_id,)).fetchone()
        if not user:
            flash('User not found.', 'danger')
            return redirect(url_for('admin_users'))
        db.execute('UPDATE users SET totp_enabled=0, totp_secret=NULL WHERE id=?', (user_id,))
        db.commit()
        log_action('RECORD_EDIT', f'2fa_disabled:user:{user["username"]}')
        flash(f"2FA disabled for '{user['username']}'.", 'warning')
        return redirect(url_for('admin_users'))

    @app.route('/admin/plant/<int:plant_id>')
    @admin_required
    def admin_select_plant(plant_id):
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
    @admin_required
    def admin_clear_plant():
        session.pop('plant_id',   None)
        session.pop('plant_name', None)
        session.pop('unit_code',  None)
        return redirect(url_for('central_dashboard'))

    @app.route('/admin/tni-archives')
    @admin_required
    def admin_tni_archives():
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
    @admin_required
    def admin_tni_restore():
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
        log_action('RECORD_ADD', f'tni_restore:{plant_name}:{fy_year}:{restored}rows')
        flash(f'Restored {restored} TNI rows for {plant_name} ({fy_year}). Programme master rebuilt.', 'success')
        return redirect(url_for('admin_tni_archives'))

    @app.route('/admin/backup/download')
    @admin_required
    def admin_backup_download():
        if not os.path.exists(DB_PATH):
            flash('Database file not found.', 'danger')
            return redirect(url_for('admin_users'))
        stamp = datetime.now().strftime('%Y-%m-%d_%H%M')
        download_name = f'training_{stamp}.db'
        log_action('RECORD_ADD', f'backup_download:{download_name}')
        return send_file(DB_PATH, as_attachment=True, download_name=download_name)

    @app.route('/admin/backup/restore', methods=['GET', 'POST'])
    @admin_required
    def admin_backup_restore():
        if request.method == 'POST':
            f = request.files.get('backup_file')
            if not f or not f.filename:
                flash('No file selected.', 'danger')
                return redirect(url_for('admin_backup_restore'))
            if not f.filename.endswith('.db'):
                flash('Invalid file type. Upload a .db file.', 'danger')
                return redirect(url_for('admin_backup_restore'))
            header = f.read(16)
            if not header.startswith(b'SQLite format 3'):
                flash('File is not a valid SQLite database.', 'danger')
                return redirect(url_for('admin_backup_restore'))
            f.seek(0)
            # Save current DB as emergency backup before overwriting
            stamp = datetime.now().strftime('%Y-%m-%d_%H%M%S')
            pre_backup = DB_PATH + f'.pre_restore_{stamp}'
            if os.path.exists(DB_PATH):
                shutil.copy2(DB_PATH, pre_backup)
            # Close all DB connections from g
            from flask import g
            db = g.pop('db', None)
            if db:
                db.close()
            # Write uploaded file atomically
            tmp = DB_PATH + '.restore_tmp'
            f.save(tmp)
            os.replace(tmp, DB_PATH)
            # Remove WAL/SHM leftovers from old DB
            for ext in ('.db-wal', '.db-shm'):
                leftover = DB_PATH.replace('.db', ext) if DB_PATH.endswith('.db') else DB_PATH + ext
                if os.path.exists(leftover):
                    os.remove(leftover)
            log_action('RECORD_ADD', f'backup_restore:{f.filename}')
            flash('Database restored successfully. All data is back.', 'success')
            return redirect(url_for('central_dashboard'))
        return render_template('admin_backup_restore.html')

    @app.route('/change-password', methods=['GET', 'POST'])
    @login_required
    def change_password():
        db = get_db()
        if request.method == 'POST':
            current = request.form.get('current_password', '')
            new_pw  = request.form.get('new_password', '').strip()
            confirm = request.form.get('confirm_password', '').strip()
            user = db.execute('SELECT * FROM users WHERE id=?', (session['user_id'],)).fetchone()
            if not check_password_hash(user['password'], current):
                flash('Current password is incorrect.', 'danger')
                return redirect(url_for('change_password'))
            if len(new_pw) < 8:
                flash('New password must be at least 8 characters.', 'danger')
                return redirect(url_for('change_password'))
            if not any(c.isdigit() for c in new_pw):
                flash('New password must contain at least one number.', 'danger')
                return redirect(url_for('change_password'))
            if new_pw != confirm:
                flash('Passwords do not match.', 'danger')
                return redirect(url_for('change_password'))
            db.execute('UPDATE users SET password=?, must_change_password=0 WHERE id=?',
                       (generate_password_hash(new_pw), session['user_id']))
            db.commit()
            log_action('PWD_CHANGE', 'self')
            flash('Password changed successfully.', 'success')
            role = session.get('role')
            return redirect(url_for('central_dashboard') if role in ('central', 'admin') else url_for('spoc_dashboard'))
        return render_template('change_password.html')

    @app.route('/admin/reset-password/<int:user_id>', methods=['POST'])
    @login_required
    def admin_reset_password(user_id):
        if session.get('role') != 'admin':
            flash('Access denied.', 'danger')
            return redirect(url_for('index'))
        db = get_db()
        user = db.execute('SELECT username FROM users WHERE id=?', (user_id,)).fetchone()
        if not user:
            flash('User not found.', 'danger')
            return redirect(url_for('central_dashboard'))
        db.execute('UPDATE users SET password=?, must_change_password=1, failed_attempts=0, locked_until=NULL WHERE id=?',
                   (generate_password_hash('bcml@1234'), user_id))
        db.commit()
        log_action('PWD_RESET', f"reset_user:{user['username']}")
        flash(f"Password for '{user['username']}' reset to default. User must change on next login.", 'success')
        return redirect(url_for('admin_users'))

    @app.route('/dashboard')
    @spoc_required
    def spoc_dashboard():
        plant_id = session['plant_id']
        db = get_db()
        fy_start, fy_end = _current_fy()
        central_attended = db.execute(
            "SELECT COUNT(DISTINCT session_code) FROM emp_training "
            "WHERE plant_id=? AND host_plant_id=99 AND session_code IS NOT NULL AND session_code!=''"
            " AND start_date BETWEEN ? AND ?",
            (plant_id, fy_start, fy_end)).fetchone()[0]
        from tms.helpers import _fy_label
        fy = _fy_label()
        stats = {
            'total_emp':    db.execute('SELECT COUNT(*) FROM employees WHERE plant_id=? AND is_active=1', (plant_id,)).fetchone()[0],
            'blue_collar':  db.execute("SELECT COUNT(*) FROM employees WHERE plant_id=? AND is_active=1 AND collar='Blue Collared'", (plant_id,)).fetchone()[0],
            'white_collar': db.execute("SELECT COUNT(*) FROM employees WHERE plant_id=? AND is_active=1 AND collar='White Collared'", (plant_id,)).fetchone()[0],
            'tni_count':    db.execute('SELECT COUNT(DISTINCT emp_code || "|" || programme_name) FROM tni WHERE plant_id=? AND fy_year=?', (plant_id, fy)).fetchone()[0],
            'sessions':     db.execute("SELECT COUNT(*) FROM calendar WHERE plant_id=? AND plan_start BETWEEN ? AND ?", (plant_id, fy_start, fy_end)).fetchone()[0],
            'conducted':    db.execute("SELECT COUNT(*) FROM calendar WHERE plant_id=? AND status='Conducted' AND plan_start BETWEEN ? AND ?", (plant_id, fy_start, fy_end)).fetchone()[0],
            'central_sessions': central_attended,
            'trainings':    db.execute('SELECT COUNT(*) FROM emp_training WHERE plant_id=? AND start_date BETWEEN ? AND ?', (plant_id, fy_start, fy_end)).fetchone()[0],
            'manhours':     db.execute('SELECT COALESCE(SUM(hrs),0) FROM emp_training WHERE plant_id=? AND start_date BETWEEN ? AND ?', (plant_id, fy_start, fy_end)).fetchone()[0],
        }
        tni_by_type = db.execute(
            'SELECT prog_type, COUNT(DISTINCT emp_code || "|" || programme_name) as cnt'
            ' FROM tni WHERE plant_id=? AND fy_year=? AND prog_type IS NOT NULL AND prog_type!=""'
            ' GROUP BY prog_type ORDER BY cnt DESC',
            (plant_id, fy)).fetchall()
        return render_template('dashboard.html', stats=stats, tni_by_type=tni_by_type)
