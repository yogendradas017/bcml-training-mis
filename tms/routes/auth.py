import io
import os
import shutil
from datetime import date, datetime, timedelta
from flask import render_template, request, redirect, url_for, session, flash, send_file, jsonify
from werkzeug.security import check_password_hash, generate_password_hash
import pyotp
import qrcode
import base64

from tms.constants import PLANT_MAP, DB_PATH
from tms.db import get_db
from tms.decorators import spoc_required, login_required, admin_required
from tms.helpers import _current_fy, _now_ist, _calc_compliance, _fy_label
from tms.audit import log_action


# Top-10k worst passwords subset — block obvious bad picks
_WEAK_PASSWORDS = {
    'password', 'password1', 'password123', 'qwerty', 'qwerty123', '12345678',
    'abc123', 'iloveyou', 'admin', 'admin123', 'letmein', 'welcome',
    'monkey', 'dragon', 'sunshine', 'football', 'baseball',
    'bcml@1234', 'bcml1234', 'bcml@123', 'admin@bcml', 'changeme', 'password@1',
}


def _validate_password_strength(pw, username=''):
    """Return error string if password is weak, else None.
    Policy: 10+ chars, mixed case, digit, special, not a weak pick, not = username."""
    if not pw or len(pw) < 10:
        return 'Password must be at least 10 characters long.'
    if len(pw) > 128:
        return 'Password is too long (max 128 characters).'
    if not any(c.isupper() for c in pw):
        return 'Password must contain at least one uppercase letter (A–Z).'
    if not any(c.islower() for c in pw):
        return 'Password must contain at least one lowercase letter (a–z).'
    if not any(c.isdigit() for c in pw):
        return 'Password must contain at least one digit (0–9).'
    if not any(not c.isalnum() for c in pw):
        return 'Password must contain at least one special character (e.g. @ # $ ! % &).'
    if pw.lower() in _WEAK_PASSWORDS:
        return 'This password is too common. Pick something unique.'
    if username and username.lower() in pw.lower():
        return 'Password must not contain your username.'
    return None


# ── QC analytics chart data (live, plant-scoped) ──────────────────────────────
# Coverage convention everywhere below: denominator = TNI-nominated rows ONLY
# (never all employees), BC/WC kept SEPARATE (never averaged), scoped by plant_id.

def _qc_pareto(db, plant_id, fy_start, fy_end):
    """Pareto — uncovered headcount per programme (nominated − trained), desc."""
    fy = _fy_label()
    rows = db.execute('''
        SELECT programme_name,
               COUNT(*) - SUM(CASE WHEN trained THEN 1 ELSE 0 END) AS uncovered
        FROM (
            SELECT t.programme_name, t.emp_code,
                   EXISTS(
                       SELECT 1 FROM emp_training et
                       JOIN programme_details pd ON pd.session_code = et.session_code
                       WHERE et.emp_code = t.emp_code
                         AND pd.programme_name = t.programme_name
                         AND pd.plant_id = t.plant_id
                   ) AS trained
            FROM tni t
            JOIN employees e ON e.emp_code = t.emp_code AND e.plant_id = t.plant_id
            WHERE t.plant_id = ? AND t.fy_year = ? AND e.is_active = 1
        )
        GROUP BY programme_name
        HAVING uncovered > 0
        ORDER BY uncovered DESC, programme_name ASC
        LIMIT 8
    ''', (plant_id, fy)).fetchall()
    return {
        'paretoLabels': [r['programme_name'] for r in rows],
        'paretoData':   [int(r['uncovered']) for r in rows],
    }


def _qc_histogram(db, plant_id, fy_start, fy_end):
    """Histogram — employees per total-FY-training-hours bucket, split BC vs WC.
    Fixed 8 buckets, upper edges [2,4,6,8,12,16,24, inf]. Never-trained → '0-2'.
    Also returns each collar's average hours/employee and its man-hour target,
    so the chart can overlay average + target reference lines."""
    from tms.config import get_config
    EDGES = [2, 4, 6, 8, 12, 16, 24, float('inf')]
    rows = db.execute('''
        SELECT e.collar AS collar, COALESCE(SUM(t.hrs), 0) AS hrs
        FROM employees e
        LEFT JOIN emp_training t
          ON t.emp_code = e.emp_code
         AND t.plant_id = e.plant_id
         AND t.start_date BETWEEN ? AND ?
        WHERE e.plant_id = ?
          AND e.is_active = 1
          AND e.collar IN ('Blue Collared', 'White Collared')
        GROUP BY e.emp_code
    ''', (fy_start, fy_end, plant_id)).fetchall()
    bc, wc = [0] * 8, [0] * 8
    bc_sum = bc_n = wc_sum = wc_n = 0
    for r in rows:
        h = r['hrs'] or 0
        for i, edge in enumerate(EDGES):
            if h < edge:
                idx = i
                break
        else:
            idx = 7
        if r['collar'] == 'Blue Collared':
            bc[idx] += 1; bc_sum += h; bc_n += 1
        else:
            wc[idx] += 1; wc_sum += h; wc_n += 1
    return {
        'histBc': bc, 'histWc': wc,
        'histAvgBc': round(bc_sum / bc_n, 1) if bc_n else 0,
        'histAvgWc': round(wc_sum / wc_n, 1) if wc_n else 0,
        'histTargetBc': get_config('mh_target_bc', 12, plant_id=plant_id),
        'histTargetWc': get_config('mh_target_wc', 24, plant_id=plant_id),
    }


def _qc_dept_compliance(db, plant_id, fy_start, fy_end):
    """Per-department man-hour target compliance: how many active employees met
    their collar target (BC vs WC, from config) vs fell below. Each employee is
    judged against their own collar's target; results grouped by department."""
    from tms.config import get_config
    bc_t = get_config('mh_target_bc', 12, plant_id=plant_id)
    wc_t = get_config('mh_target_wc', 24, plant_id=plant_id)
    rows = db.execute('''
        SELECT COALESCE(NULLIF(e.department, ''), 'Unassigned') AS dept,
               e.collar AS collar, COALESCE(SUM(t.hrs), 0) AS hrs
        FROM employees e
        LEFT JOIN emp_training t
          ON t.emp_code = e.emp_code AND t.plant_id = e.plant_id
         AND t.start_date BETWEEN ? AND ?
        WHERE e.plant_id = ? AND e.is_active = 1
          AND e.collar IN ('Blue Collared', 'White Collared')
        GROUP BY e.emp_code
    ''', (fy_start, fy_end, plant_id)).fetchall()
    agg = {}  # dept -> {met, below}
    for r in rows:
        target = bc_t if r['collar'] == 'Blue Collared' else wc_t
        d = agg.setdefault(r['dept'], {'met': 0, 'below': 0})
        d['met' if (r['hrs'] or 0) >= target else 'below'] += 1
    out = []
    for dept, c in agg.items():
        total = c['met'] + c['below']
        out.append({'dept': dept, 'met': c['met'], 'below': c['below'],
                    'total': total, 'pct': round(c['met'] / total * 100) if total else 0})
    out.sort(key=lambda x: x['pct'], reverse=True)
    return {'deptCompliance': out}


def _qc_cumulative(db, plant_id, fy_start, fy_end):
    """Cumulative run — BC vs WC cumulative TNI-coverage % across 12 FY months.
    Separate denominators per collar (never averaged); monotonic non-decreasing."""
    fy = _fy_label()
    start_year = int(fy_start[:4])
    month_ends = [
        f'{start_year}-04-30', f'{start_year}-05-31', f'{start_year}-06-30',
        f'{start_year}-07-31', f'{start_year}-08-31', f'{start_year}-09-30',
        f'{start_year}-10-31', f'{start_year}-11-30', f'{start_year}-12-31',
        f'{start_year + 1}-01-31', f'{start_year + 1}-02-28', f'{start_year + 1}-03-31',
    ]
    # Two flat, index-driven queries + a Python match — avoids the tni×sessions×
    # attendance join explosion that made the SQL-join versions take 35–66s.
    # 1) trained facts: earliest in-FY training date per (emp, programme).
    trained = {}  # (emp_code, programme_name) -> earliest start_date
    for r in db.execute('''
            SELECT et.emp_code AS emp, pd.programme_name AS prog, MIN(et.start_date) AS first_train
            FROM emp_training et
            JOIN programme_details pd ON pd.session_code = et.session_code AND pd.plant_id = et.plant_id
            WHERE et.plant_id = ? AND et.start_date BETWEEN ? AND ?
            GROUP BY et.emp_code, pd.programme_name
        ''', (plant_id, fy_start, fy_end)).fetchall():
        trained[(r['emp'], r['prog'])] = r['first_train']
    # 2) nominations (the coverage universe) with collar.
    denom = {'Blue Collared': 0, 'White Collared': 0}
    firsts = {'Blue Collared': [], 'White Collared': []}
    for r in db.execute('''
            SELECT t.emp_code AS emp, t.programme_name AS prog, e.collar AS collar
            FROM tni t
            JOIN employees e ON e.emp_code = t.emp_code AND e.plant_id = t.plant_id
            WHERE t.plant_id = ? AND t.fy_year = ? AND e.is_active = 1
              AND e.collar IN ('Blue Collared', 'White Collared')
        ''', (plant_id, fy)).fetchall():
        collar = r['collar']
        denom[collar] += 1
        firsts[collar].append(trained.get((r['emp'], r['prog'])))

    def cum_series(collar):
        d = denom[collar]
        if d == 0:
            return [0] * 12
        return [round(100.0 * sum(1 for ft in firsts[collar] if ft and ft <= me) / d, 1)
                for me in month_ends]

    return {
        'bcCumulative': cum_series('Blue Collared'),
        'wcCumulative': cum_series('White Collared'),
    }


def _qc_hclass(pct):
    """Heatmap cell CSS class by coverage %: h0<25, h25<50, h50<75, h75<90, h90<100, h100."""
    if pct >= 100: return 'h100'
    if pct >= 90:  return 'h90'
    if pct >= 75:  return 'h75'
    if pct >= 50:  return 'h50'
    if pct >= 25:  return 'h25'
    return 'h0'


def _qc_heatmap(db, plant_id, fy_start, fy_end):
    """Heat map — prog_type × collar TNI-coverage % matrix in display column order.
    Coverage = trained / nominated per cell, over TNI rows ONLY, never averaged."""
    DISPLAY_COLS = [
        ('Technical',  'Technical'),
        ('EHS/HR',     'EHS/HR'),
        ('Behav.',     'Behavioural/Leadership'),
        ('Cane',       'Cane'),
        ('Commercial', 'Commercial'),
        ('IT',         'IT'),
    ]
    ROW_COLLARS = ['Blue Collared', 'White Collared']
    rows = db.execute('''
        SELECT prog_type, collar,
               ROUND(100.0 * SUM(CASE WHEN trained THEN 1 ELSE 0 END)
                     / NULLIF(COUNT(*), 0), 1) AS cov_pct
        FROM (
            SELECT t.prog_type, e.collar, t.emp_code,
                   EXISTS(
                       SELECT 1 FROM emp_training et
                       JOIN programme_details pd ON pd.session_code = et.session_code
                       WHERE et.emp_code = t.emp_code
                         AND pd.programme_name = t.programme_name
                         AND pd.plant_id = t.plant_id
                   ) AS trained
            FROM tni t
            JOIN employees e ON e.emp_code = t.emp_code AND e.plant_id = t.plant_id
            WHERE t.plant_id = ? AND e.is_active = 1
        )
        GROUP BY prog_type, collar
    ''', (plant_id,)).fetchall()
    matrix = {}
    for r in rows:
        matrix.setdefault(r['collar'] or '', {})[r['prog_type'] or ''] = r['cov_pct']
    heatmap = []
    for collar in ROW_COLLARS:
        cells = []
        for _label, ptype in DISPLAY_COLS:
            pct = matrix.get(collar, {}).get(ptype)
            pct = int(round(pct)) if pct is not None else 0
            cells.append({'pct': pct, 'cls': _qc_hclass(pct)})
        heatmap.append({'collar': collar, 'cells': cells})
    return {'heatmap': heatmap}


def _register(app):

    @app.route('/_dashboard-mockup')
    def _dashboard_mockup():
        """Temp: serve the mockup standalone (used inside compare iframe)."""
        return render_template('_dashboard_mockup.html')

    @app.route('/_dashboard-compare')
    def _dashboard_compare():
        """Temp: side-by-side current /dashboard vs proposed mockup."""
        return render_template('_dashboard_compare.html')

    @app.route('/_dashboard-styles')
    def _dashboard_styles():
        """Temp: Vercel vs Power BI style comparison for design direction."""
        return render_template('_dashboard_styles.html')

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
                    if _now_ist() < datetime.fromisoformat(user['locked_until']):
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
                session['user_id']      = user['id']
                session['username']     = user['username']
                session['role']         = user['role']
                session['plant_id']     = user['plant_id']
                session['totp_enabled'] = bool(user['totp_enabled'])
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
                        locked_until = (_now_ist() + timedelta(minutes=15)).isoformat()
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
                session['user_id']      = user['id']
                session['username']     = user['username']
                session['role']         = user['role']
                session['plant_id']     = user['plant_id']
                session['totp_enabled'] = True
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
                   u.totp_enabled, u.totp_secret,
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

    @app.route('/admin/audit-log/verify', methods=['POST'])
    @admin_required
    def admin_audit_log_verify():
        """Recompute the full audit-log hash chain and flash result.
        Surfaces any tampered rows by id. Cheap enough to run interactively;
        also wired to the nightly cron for unattended verification."""
        from tms.audit import verify_chain
        db = get_db()
        broken = verify_chain(db)
        if broken:
            flash(f'AUDIT CHAIN BROKEN — {len(broken)} row(s) tampered: '
                  + ', '.join(str(i) for i in broken[:20])
                  + ('…' if len(broken) > 20 else ''), 'danger')
        else:
            count = db.execute('SELECT COUNT(*) FROM audit_log').fetchone()[0]
            flash(f'Audit chain intact — verified {count} rows.', 'success')
        log_action('AUDIT_VERIFY', f'broken={len(broken)}')
        return redirect(url_for('admin_audit_log'))

    @app.route('/2fa/setup', methods=['GET', 'POST'])
    @login_required
    def self_2fa_setup():
        """Self-service 2FA enrollment for the logged-in user.
        Mandatory for central/admin (enforced by decorators).
        """
        db = get_db()
        uid = session['user_id']
        user = db.execute('SELECT * FROM users WHERE id=?', (uid,)).fetchone()
        if not user:
            session.clear()
            return redirect(url_for('login'))
        if request.method == 'POST':
            code = request.form.get('totp_code', '').strip().replace(' ', '')
            if not user['totp_secret']:
                flash('Setup error — reload page and try again.', 'danger')
                return redirect(url_for('self_2fa_setup'))
            if pyotp.TOTP(user['totp_secret']).verify(code, valid_window=1):
                db.execute('UPDATE users SET totp_enabled=1 WHERE id=?', (uid,))
                db.commit()
                session['totp_enabled'] = True
                log_action('RECORD_EDIT', '2fa_self_enabled')
                flash('Two-factor authentication enabled successfully.', 'success')
                role = session.get('role')
                return redirect(url_for('central_dashboard') if role in ('central', 'admin') else url_for('spoc_dashboard'))
            flash('Invalid code — scan QR again and retry.', 'danger')
        secret = user['totp_secret'] or pyotp.random_base32()
        if not user['totp_secret']:
            db.execute('UPDATE users SET totp_secret=? WHERE id=?', (secret, uid))
            db.commit()
        uri = pyotp.TOTP(secret).provisioning_uri(name=user['username'], issuer_name='BCML TMS')
        img = qrcode.make(uri)
        buf = io.BytesIO()
        img.save(buf, format='PNG')
        qr_b64 = base64.b64encode(buf.getvalue()).decode()
        return render_template('admin_2fa_setup.html', user=user, secret=secret, qr_b64=qr_b64,
                               self_service=True)

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
            if not user:
                session.clear()
                flash('Session expired. Please log in again.', 'danger')
                return redirect(url_for('login'))
            if not check_password_hash(user['password'], current):
                flash('Current password is incorrect.', 'danger')
                return redirect(url_for('change_password'))
            policy_err = _validate_password_strength(new_pw, user['username'])
            if policy_err:
                flash(policy_err, 'danger')
                return redirect(url_for('change_password'))
            if new_pw != confirm:
                flash('Passwords do not match.', 'danger')
                return redirect(url_for('change_password'))
            # Block re-using current password
            if check_password_hash(user['password'], new_pw):
                flash('New password must be different from your current password.', 'danger')
                return redirect(url_for('change_password'))
            db.execute('UPDATE users SET password=?, must_change_password=0 WHERE id=?',
                       (generate_password_hash(new_pw), session['user_id']))
            db.commit()
            log_action('PWD_CHANGE', 'self')
            flash('Password changed successfully.', 'success')
            role = session.get('role')
            return redirect(url_for('central_dashboard') if role in ('central', 'admin') else url_for('spoc_dashboard'))
        return render_template('change_password.html')

    @app.route('/admin/users/<int:user_id>/set-role', methods=['POST'])
    @login_required
    def admin_set_role(user_id):
        if session.get('role') != 'admin':
            flash('Access denied.', 'danger')
            return redirect(url_for('index'))
        if user_id == session.get('user_id'):
            flash('Cannot change your own role — ask another admin.', 'danger')
            return redirect(url_for('admin_users'))
        new_role = (request.form.get('role') or '').strip()
        if new_role not in ('spoc', 'central', 'admin'):
            flash('Invalid role.', 'danger')
            return redirect(url_for('admin_users'))
        db = get_db()
        user = db.execute('SELECT username FROM users WHERE id=?', (user_id,)).fetchone()
        if not user:
            flash('User not found.', 'danger')
            return redirect(url_for('admin_users'))
        db.execute('UPDATE users SET role=? WHERE id=?', (new_role, user_id))
        db.commit()
        log_action('ROLE_CHANGE', f"user:{user['username']}:to={new_role}")
        flash(f"Role for '{user['username']}' set to {new_role}.", 'success')
        return redirect(url_for('admin_users'))

    @app.route('/admin/reset-password/<int:user_id>', methods=['POST'])
    @login_required
    def admin_reset_password(user_id):
        if session.get('role') != 'admin':
            flash('Access denied.', 'danger')
            return redirect(url_for('index'))
        # Block self-reset: prevents session-hijack DoS where stolen admin session
        # could lock out the legitimate admin. Another admin must reset.
        if user_id == session.get('user_id'):
            flash('Cannot reset your own password from here. Use Change Password (with current) '
                  'or ask another admin to reset on your behalf.', 'danger')
            return redirect(url_for('admin_users'))
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
        compliance = _calc_compliance(plant_id, db)
        target_hrs = compliance['bc_mandate'] + compliance['wc_mandate']
        # Avg Hrs per Employee — ties to manhour mandate (BC 12/yr, WC 24/yr).
        avg_hrs = round(stats['manhours'] / stats['total_emp'], 1) if stats['total_emp'] else 0
        # QC analytics charts are fetched async via /api/dashboard-qc so the
        # dashboard page renders immediately instead of blocking on their queries.
        return render_template(
            'dashboard.html',
            stats=stats,
            compliance=compliance,
            target_hrs=target_hrs,
            avg_hrs=avg_hrs,
            bc_pct=compliance['bc_pct'],
            wc_pct=compliance['wc_pct'],
            headline_pct=compliance['headline_pct'],
            headline_rag=compliance['headline_rag'],
            worst_cells=compliance['worst_cells'],
        )

    @app.route('/api/dashboard-qc')
    @spoc_required
    def api_dashboard_qc():
        """QC analytics datasets (Pareto, Histogram, Heat map, Cumulative) as JSON,
        fetched async by the dashboard so the page is not blocked on these queries."""
        plant_id = session['plant_id']
        db = get_db()
        fy_start, fy_end = _current_fy()
        out = {}
        out.update(_qc_pareto(db, plant_id, fy_start, fy_end))
        out.update(_qc_histogram(db, plant_id, fy_start, fy_end))
        out.update(_qc_dept_compliance(db, plant_id, fy_start, fy_end))
        out.update(_qc_cumulative(db, plant_id, fy_start, fy_end))
        out.update(_qc_heatmap(db, plant_id, fy_start, fy_end))
        return jsonify(out)

    @app.route('/admin/seed-demo', methods=['GET', 'POST'])
    @admin_required
    def admin_seed_demo():
        if request.method == 'POST':
            import seed_synthetic as _s
            db = get_db()
            db.execute("PRAGMA foreign_keys = OFF")
            db.execute("PRAGMA journal_mode = WAL")
            db.execute("DELETE FROM emp_training")
            db.execute("DELETE FROM programme_details")
            db.execute("DELETE FROM calendar")
            db.execute("DELETE FROM tni WHERE plant_id != 1")
            db.execute("DELETE FROM programme_master WHERE plant_id != 1")
            db.commit()
            cal, et, pd_ = _s.seed(db)
            log_action('BULK_UPLOAD', f'seed_demo cal={cal} et={et} pd={pd_}')
            flash(f'Demo data seeded — {cal} sessions, {et} attendance rows, {pd_} programme details.', 'success')
            return redirect(url_for('central_dashboard'))
        return render_template('admin_seed_demo.html')
