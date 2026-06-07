import io
import hmac
import secrets
import datetime
from urllib.parse import urlparse

import qrcode
from flask import (abort, flash, jsonify, redirect, render_template,
                   request, send_file, session, url_for)

from tms.db import get_db
from tms.decorators import spoc_required, central_required, spoc_or_central_required
from tms.helpers import _date_to_month, _recompute_session_actuals, _now_ist, _today_ist
from tms.constants import CENTRAL_PLANT_ID
from tms.audit import log_action


# ── helpers ──────────────────────────────────────────────────────────────────

def _now_iso():
    # IST wall-clock; on Render (UTC) datetime.now() naive would drift 5.5h.
    return _now_ist().isoformat(timespec='seconds')


def _avg(vals):
    v = [x for x in vals if x is not None]
    return round(sum(v) / len(v), 2) if v else None


def _validate_token(token, db, check_expiry=True):
    row = db.execute('''
        SELECT q.*, c.programme_name, c.prog_type, c.level, c.mode,
               c.plan_start, c.plan_end, c.duration_hrs,
               c.time_from, c.time_to, c.target_audience, c.session_pin,
               c.is_central,
               p.name AS plant_name
        FROM session_qr q
        JOIN calendar c ON c.session_code=q.session_code AND c.plant_id=q.plant_id
        JOIN plants p ON p.id=q.plant_id
        WHERE q.token=? AND q.is_active=1
    ''', (token,)).fetchone()
    if not row:
        abort(404)
    if check_expiry and row['expires_at'] and _now_iso() > row['expires_at']:
        abort(410)
    return row


def _recompute_feedback_aggregates(plant_id, session_code, db):
    """Fold QR feedback averages into the session's 2C row IF one exists.

    COALESCE-only: fills ONLY the feedback columns the SPOC left NULL, so it
    never clobbers manually-entered 2C feedback. If no programme_details (2C)
    row exists yet, this is a deliberate no-op — the responses live safely in
    feedback_response (the feedback report reads them directly) and are folded
    in when the SPOC saves 2C (add_programme_details calls this after insert).

    We intentionally do NOT create a stub programme_details row here. The old
    stub (audit Tier 2) caused a CRITICAL chain break: it tripped the
    'already recorded' guard so the real 2C save returned early (its merge
    branch became dead code), and _sync_calendar_from_2c auto-promoted the
    stub to 'Conducted' — fabricating a phantom conducted programme with 0
    hours, no faculty/cost, no verification, and no effectiveness seeding.
    """
    row = db.execute('''
        SELECT
            AVG(NULLIF(q_obj_explained,0))        AS q1,
            AVG(NULLIF(q_well_structured,0))      AS q2,
            AVG(NULLIF(q_content_appropriate,0))  AS q3,
            AVG(NULLIF(q_presentation_quality,0)) AS q4,
            AVG(NULLIF(q_time_reasonable,0))      AS q5,
            AVG(NULLIF(q_inputs_appropriate,0))   AS q6,
            AVG(NULLIF(q_communication_clear,0))  AS q7,
            AVG(NULLIF(q_queries_responded,0))    AS q8,
            AVG(NULLIF(q_well_involved,0))        AS q9,
            COUNT(*) AS total_responses
        FROM feedback_response
        WHERE plant_id=? AND session_code=?
    ''', (plant_id, session_code)).fetchone()

    prog_avg    = _avg([row['q1'], row['q2'], row['q3'], row['q4'], row['q5']])
    trainer_avg = _avg([row['q6'], row['q7'], row['q8'], row['q9']])

    db.execute('''UPDATE programme_details SET
                   course_feedback         = COALESCE(course_feedback, ?),
                   faculty_feedback        = COALESCE(faculty_feedback, ?),
                   trainer_fb_participants = COALESCE(trainer_fb_participants, ?),
                   trainer_fb_facilities   = COALESCE(trainer_fb_facilities, ?)
                  WHERE plant_id=? AND session_code=?''',
               (prog_avg, trainer_avg, row['q8'], row['q9'], plant_id, session_code))


def _is_cross_origin_post():
    """For the CSRF-exempt public QR POSTs: block forged cross-site submissions.
    Flask-WTF CSRF is exempted on these routes (phone scans carry no session), so
    we fall back to an Origin/Referer same-host check. If the browser sent an
    Origin/Referer and its host differs from ours, reject; absent header (some
    mobile scanners strip Referer on same-origin form posts) is allowed."""
    if request.method != 'POST':
        return False
    origin = request.headers.get('Origin') or request.headers.get('Referer') or ''
    if not origin:
        return False
    try:
        return urlparse(origin).netloc != request.host
    except Exception:
        return True


def _make_qr_png(url):
    img = qrcode.make(url)
    buf = io.BytesIO()
    img.save(buf, format='PNG')
    buf.seek(0)
    return buf


# ── register ─────────────────────────────────────────────────────────────────

def _cal_home():
    """Redirect back to whichever calendar the user came from."""
    if session.get('role') in ('central', 'admin'):
        return redirect(url_for('central_calendar'))
    return redirect(url_for('training_calendar'))


def _register(app):

    # ── SPOC: generate QR for a calendar session ─────────────────────────────

    @app.route('/calendar/<int:cal_id>/qr/generate', methods=['POST'])
    @spoc_or_central_required
    def qr_generate(cal_id):
        role = session.get('role')
        db = get_db()
        if role in ('central', 'admin'):
            cal = db.execute('SELECT * FROM calendar WHERE id=?', (cal_id,)).fetchone()
        else:
            cal = db.execute('SELECT * FROM calendar WHERE id=? AND plant_id=?',
                             (cal_id, session['plant_id'])).fetchone()
        if not cal:
            flash('Session not found.', 'danger')
            return _cal_home()
        plant_id = cal['plant_id']

        stage = request.form.get('stage', 'attendance')
        if stage not in ('attendance', 'feedback'):
            stage = 'attendance'

        # Lock: QR generate only allowed while status='To Be Planned'.
        # Once 2C exists (Awaiting Verification / Conducted) or session is
        # Lapsed/Cancelled/Re-Scheduled, generating a new QR would let new
        # scans mutate verified records.
        if cal['status'] != 'To Be Planned':
            flash(f'Cannot generate QR — session is "{cal["status"]}". QR is locked once 2C is recorded.', 'danger')
            return _cal_home()

        # GAP 16: block QR generation if calendar has no plan_end
        if not cal['plan_end'] or not cal['plan_start']:
            flash('Cannot generate QR — session has no planned start/end date. Set dates in calendar first.', 'danger')
            return _cal_home()

        try:
            plan_end    = datetime.date.fromisoformat(cal['plan_end'])
            expiry_date = plan_end + datetime.timedelta(days=1)
            if expiry_date < _today_ist():
                expiry_date = _today_ist() + datetime.timedelta(days=30)
            expires = expiry_date.isoformat() + 'T23:59:59'
        except Exception:
            flash('Cannot generate QR — calendar plan_end date is invalid.', 'danger')
            return _cal_home()

        existing = db.execute(
            'SELECT id FROM session_qr WHERE plant_id=? AND session_code=? AND stage=?',
            (plant_id, cal['session_code'], stage)
        ).fetchone()

        if existing:
            new_token = secrets.token_urlsafe(16)
            # IST timestamp written explicitly — schema has no SQL DEFAULT so Render
            # (UTC host) and local (IST host) agree.
            now_ist_iso = _now_ist().isoformat(timespec='seconds')
            db.execute('''UPDATE session_qr SET token=?, is_active=1, expires_at=?,
                          created_at=?, created_by=?
                          WHERE id=?''',
                       (new_token, expires, now_ist_iso, session.get('user_id'), existing['id']))
            db.commit()
            flash(f'QR regenerated for {stage}.', 'success')
        else:
            new_token = secrets.token_urlsafe(16)
            # IST timestamp written explicitly — schema has no SQL DEFAULT.
            now_ist_iso = _now_ist().isoformat(timespec='seconds')
            db.execute('''INSERT INTO session_qr
                (plant_id, session_code, token, stage, created_at, expires_at, created_by)
                VALUES(?,?,?,?,?,?,?)''',
                (plant_id, cal['session_code'], new_token, stage, now_ist_iso, expires,
                 session.get('user_id')))
            db.commit()
            flash(f'QR generated for {stage}.', 'success')

        return redirect(url_for('qr_poster', token=new_token))

    # ── SPOC: QR image (PNG stream) ───────────────────────────────────────────

    @app.route('/qr/<token>/image.png')
    @spoc_or_central_required
    def qr_image(token):
        db = get_db()
        row = db.execute('SELECT 1 FROM session_qr WHERE token=?', (token,)).fetchone()
        if not row:
            abort(404)
        url = request.host_url + f'q/{token}'
        buf = _make_qr_png(url)
        return send_file(buf, mimetype='image/png')

    # ── SPOC: printable poster ────────────────────────────────────────────────

    @app.route('/qr/<token>/poster')
    @spoc_or_central_required
    def qr_poster(token):
        db = get_db()
        qr = _validate_token(token, db, check_expiry=False)
        qr_url = request.host_url + f'q/{token}'
        return render_template('qr_poster.html', qr=qr, token=token, qr_url=qr_url)

    # ── SPOC: revoke QR ───────────────────────────────────────────────────────

    @app.route('/qr/<int:qr_id>/revoke', methods=['POST'])
    @spoc_or_central_required
    def qr_revoke(qr_id):
        role = session.get('role')
        db = get_db()
        # Lock: only revoke QR for sessions still in 'To Be Planned'.
        # Post-2C, the QR is already de-facto frozen by qr_attend/qr_feedback
        # guards; allowing UPDATE here would create misleading audit churn.
        q = db.execute(
            'SELECT q.id, q.plant_id, c.status FROM session_qr q '
            'JOIN calendar c ON c.session_code=q.session_code AND c.plant_id=q.plant_id '
            'WHERE q.id=?', (qr_id,)).fetchone()
        if not q:
            flash('QR not found.', 'danger')
            return _cal_home()
        if role not in ('central', 'admin') and q['plant_id'] != session.get('plant_id'):
            flash('Not your plant.', 'danger')
            return _cal_home()
        if q['status'] != 'To Be Planned':
            flash(f'Cannot revoke — session is "{q["status"]}". QR is locked.', 'danger')
            return _cal_home()
        db.execute('UPDATE session_qr SET is_active=0 WHERE id=?', (qr_id,))
        db.commit()
        flash('QR revoked — old QR no longer accepts scans.', 'warning')
        return _cal_home()

    # ── SPOC: set/clear session PIN ──────────────────────────────────────────

    @app.route('/calendar/<int:cal_id>/set-pin', methods=['POST'])
    @spoc_or_central_required
    def qr_set_pin(cal_id):
        role = session.get('role')
        plant_id = session['plant_id']
        db = get_db()
        # Lock: PIN mutates calendar row. Block once session moves past planning.
        cal = db.execute('SELECT status, plant_id FROM calendar WHERE id=?',
                         (cal_id,)).fetchone()
        if not cal:
            flash('Session not found.', 'danger')
            return _cal_home()
        if role not in ('central', 'admin') and cal['plant_id'] != plant_id:
            flash('Not your plant.', 'danger')
            return _cal_home()
        if cal['status'] != 'To Be Planned':
            flash(f'Cannot set PIN — session is "{cal["status"]}". PIN is locked.', 'danger')
            return redirect(url_for('qr_live', cal_id=cal_id))
        pin = request.form.get('pin', '').strip()
        if pin and (len(pin) != 4 or not pin.isdigit()):
            flash('PIN must be exactly 4 digits.', 'danger')
            return redirect(url_for('qr_live', cal_id=cal_id))
        if role in ('central', 'admin'):
            db.execute('UPDATE calendar SET session_pin=? WHERE id=?', (pin or None, cal_id))
        else:
            db.execute('UPDATE calendar SET session_pin=? WHERE id=? AND plant_id=?',
                       (pin or None, cal_id, plant_id))
        db.commit()
        if pin:
            flash(f'Session PIN set to {pin}. Announce it to participants.', 'success')
        else:
            flash('Session PIN cleared — attendance open without PIN.', 'warning')
        return redirect(url_for('qr_live', cal_id=cal_id))

    # ── SPOC: live attendance monitor ─────────────────────────────────────────

    @app.route('/calendar/<int:cal_id>/live')
    @spoc_or_central_required
    def qr_live(cal_id):
        role = session.get('role')
        db = get_db()
        if role in ('central', 'admin'):
            cal = db.execute('SELECT * FROM calendar WHERE id=?', (cal_id,)).fetchone()
        else:
            cal = db.execute('SELECT * FROM calendar WHERE id=? AND plant_id=?',
                             (cal_id, session['plant_id'])).fetchone()
        if not cal:
            flash('Session not found.', 'danger')
            return _cal_home()
        plant_id = cal['plant_id']

        qr_rows = db.execute(
            'SELECT * FROM session_qr WHERE plant_id=? AND session_code=? ORDER BY stage',
            (plant_id, cal['session_code'])
        ).fetchall()

        is_central = (plant_id == CENTRAL_PLANT_ID)
        if is_central:
            attendees = db.execute('''
                SELECT t.emp_code, t.created_at,
                       COALESCE(e.name, cm.name) AS name,
                       COALESCE(e.designation, cm.designation) AS designation,
                       COALESCE(e.department, cm.department) AS department,
                       e.collar,
                       p.unit_code AS unit_code
                FROM emp_training t
                LEFT JOIN employees e ON e.emp_code=t.emp_code AND e.plant_id=t.plant_id
                LEFT JOIN corp_members cm ON cm.emp_code=t.emp_code AND t.plant_id=99
                LEFT JOIN plants p ON p.id=t.plant_id
                WHERE t.session_code=? AND (t.host_plant_id=99 OR t.plant_id=99)
                ORDER BY t.created_at DESC
            ''', (cal['session_code'],)).fetchall()
        else:
            attendees = db.execute('''
                SELECT t.emp_code, t.created_at, e.name, e.designation, e.department,
                       e.collar, NULL AS unit_code
                FROM emp_training t
                LEFT JOIN employees e ON e.emp_code=t.emp_code AND e.plant_id=t.plant_id
                WHERE t.plant_id=? AND t.session_code=?
                ORDER BY t.created_at DESC
            ''', (plant_id, cal['session_code'])).fetchall()

        fb_count = db.execute(
            'SELECT COUNT(*) FROM feedback_response WHERE plant_id=? AND session_code=?',
            (plant_id, cal['session_code'])
        ).fetchone()[0]

        return render_template('qr_live.html', cal=cal, qr_rows=qr_rows,
                               attendees=attendees, fb_count=fb_count,
                               is_central=is_central)

    # ── SPOC: feedback reports index ─────────────────────────────────────────

    @app.route('/feedback-reports')
    @spoc_or_central_required
    def feedback_reports_index():
        role = session.get('role')
        if role == 'central':
            plant_id = CENTRAL_PLANT_ID
        elif role == 'admin':
            # If admin has switched to a specific plant, show that plant's feedback
            plant_id = session.get('plant_id') or CENTRAL_PLANT_ID
        else:
            plant_id = session['plant_id']
        db = get_db()
        rows = db.execute('''
            SELECT c.id AS cal_id, c.session_code, c.programme_name,
                   c.plan_start, c.plan_end, c.status,
                   COUNT(f.id) AS fb_count,
                   AVG(CASE WHEN f.q_obj_explained>0 THEN f.q_obj_explained END +
                       CASE WHEN f.q_well_structured>0 THEN f.q_well_structured END +
                       CASE WHEN f.q_content_appropriate>0 THEN f.q_content_appropriate END +
                       CASE WHEN f.q_presentation_quality>0 THEN f.q_presentation_quality END +
                       CASE WHEN f.q_time_reasonable>0 THEN f.q_time_reasonable END +
                       CASE WHEN f.q_inputs_appropriate>0 THEN f.q_inputs_appropriate END +
                       CASE WHEN f.q_communication_clear>0 THEN f.q_communication_clear END +
                       CASE WHEN f.q_queries_responded>0 THEN f.q_queries_responded END +
                       CASE WHEN f.q_well_involved>0 THEN f.q_well_involved END) / 9.0 AS avg_score
            FROM calendar c
            JOIN feedback_response f ON f.session_code=c.session_code AND f.plant_id=c.plant_id
            WHERE c.plant_id=?
            GROUP BY c.id
            ORDER BY c.plan_start DESC
        ''', (plant_id,)).fetchall()
        return render_template('feedback_reports_index.html', rows=rows)

    # ── SPOC: feedback analysis report ───────────────────────────────────────

    @app.route('/calendar/<int:cal_id>/feedback-report')
    @spoc_or_central_required
    def qr_feedback_report(cal_id):
        role = session.get('role')
        db = get_db()
        if role in ('central', 'admin'):
            cal = db.execute('SELECT * FROM calendar WHERE id=?', (cal_id,)).fetchone()
        else:
            cal = db.execute('SELECT * FROM calendar WHERE id=? AND plant_id=?',
                             (cal_id, session['plant_id'])).fetchone()
        if not cal:
            flash('Session not found.', 'danger')
            return _cal_home()
        plant_id = cal['plant_id']

        rows = db.execute('''
            SELECT r.*,
                   COALESCE(e.name, cm.name) AS name,
                   COALESCE(e.designation, cm.designation) AS designation,
                   COALESCE(e.department, cm.department) AS department
            FROM feedback_response r
            LEFT JOIN employees e ON e.emp_code=r.emp_code AND e.plant_id=r.plant_id
            LEFT JOIN corp_members cm ON cm.emp_code=r.emp_code AND r.plant_id=99
            WHERE r.plant_id=? AND r.session_code=?
            ORDER BY r.submitted_at
        ''', (plant_id, cal['session_code'])).fetchall()

        q_fields = [
            ('q_obj_explained',       'Objectives clearly explained'),
            ('q_well_structured',     'Programme well structured'),
            ('q_content_appropriate', 'Content appropriate for group'),
            ('q_presentation_quality','Quality of presentation was high'),
            ('q_time_reasonable',     'Time allocation was reasonable'),
            ('q_inputs_appropriate',  'Faculty inputs were appropriate'),
            ('q_communication_clear', 'Faculty communication was clear'),
            ('q_queries_responded',   'Queries responded by faculty'),
            ('q_well_involved',       'Participants well involved by faculty'),
        ]

        def _analyse(field, rows):
            vals = [r[field] for r in rows if r[field] and 1 <= r[field] <= 4]
            if not vals:
                return {'sd':0,'d':0,'a':0,'sa':0,'avg':None,'pct':None,'n':0}
            return {
                'sd':  vals.count(1),
                'd':   vals.count(2),
                'a':   vals.count(3),
                'sa':  vals.count(4),
                'avg': round(sum(vals)/len(vals), 2),
                'pct': round(sum(vals)/len(vals)/4*100, 1),
                'n':   len(vals),
            }

        q_stats = [(label, _analyse(field, rows)) for field, label in q_fields]

        def _subtotal(stats_slice):
            avgs = [s['avg'] for _, s in stats_slice if s['avg'] is not None]
            if not avgs:
                return None, None
            avg = round(sum(avgs)/len(avgs), 2)
            return avg, round(avg/4*100, 1)

        prog_avg,    prog_pct    = _subtotal(q_stats[:5])
        trainer_avg, trainer_pct = _subtotal(q_stats[5:])
        overall_avg, overall_pct = _subtotal(q_stats)

        learnings   = [r['key_learnings']  for r in rows if r['key_learnings']  and r['key_learnings'].strip()]
        suggestions = [r['suggestions']     for r in rows if r['suggestions']    and r['suggestions'].strip()]

        return render_template('feedback_report.html',
                               cal=cal, rows=rows,
                               q_stats=q_stats, q_fields=q_fields,
                               prog_avg=prog_avg, prog_pct=prog_pct,
                               trainer_avg=trainer_avg, trainer_pct=trainer_pct,
                               overall_avg=overall_avg, overall_pct=overall_pct,
                               learnings=learnings, suggestions=suggestions)

    # ── SPOC: live JSON poll ──────────────────────────────────────────────────

    @app.route('/api/qr/<int:cal_id>/live.json')
    @spoc_or_central_required
    def qr_live_json(cal_id):
        role = session.get('role')
        db = get_db()
        if role in ('central', 'admin'):
            cal = db.execute('SELECT session_code, planned_pax, plant_id FROM calendar WHERE id=?',
                             (cal_id,)).fetchone()
        else:
            cal = db.execute('SELECT session_code, planned_pax, plant_id FROM calendar WHERE id=? AND plant_id=?',
                             (cal_id, session['plant_id'])).fetchone()
        if not cal:
            return jsonify({'error': 'not found'}), 404
        plant_id = cal['plant_id']

        if plant_id == CENTRAL_PLANT_ID:
            rows = db.execute('''
                SELECT t.emp_code, t.created_at,
                       COALESCE(e.name, cm.name) AS name,
                       COALESCE(e.designation, cm.designation) AS designation,
                       COALESCE(e.department, cm.department) AS department,
                       p.unit_code AS unit_code
                FROM emp_training t
                LEFT JOIN employees e ON e.emp_code=t.emp_code AND e.plant_id=t.plant_id
                LEFT JOIN corp_members cm ON cm.emp_code=t.emp_code AND t.plant_id=99
                LEFT JOIN plants p ON p.id=t.plant_id
                WHERE t.session_code=? AND (t.host_plant_id=99 OR t.plant_id=99)
                ORDER BY t.created_at DESC
            ''', (cal['session_code'],)).fetchall()
        else:
            rows = db.execute('''
                SELECT t.emp_code, t.created_at, e.name, e.designation, e.department,
                       NULL AS unit_code
                FROM emp_training t
                LEFT JOIN employees e ON e.emp_code=t.emp_code AND e.plant_id=t.plant_id
                WHERE t.plant_id=? AND t.session_code=?
                ORDER BY t.created_at DESC
            ''', (plant_id, cal['session_code'])).fetchall()

        fb_count = db.execute(
            'SELECT COUNT(*) FROM feedback_response WHERE plant_id=? AND session_code=?',
            (plant_id, cal['session_code'])
        ).fetchone()[0]

        return jsonify({
            'count': len(rows),
            'planned_pax': cal['planned_pax'] or 0,
            'fb_count': fb_count,
            'rows': [{'emp_code': r['emp_code'], 'name': r['name'] or '',
                      'designation': r['designation'] or '',
                      'department': r['department'] or '',
                      'unit_code': r['unit_code'] or '',
                      'scanned_at': r['created_at'] or ''} for r in rows]
        })

    # ── SPOC: emp search for live view ────────────────────────────────────────

    @app.route('/api/emp-search')
    @spoc_required
    def spoc_emp_search():
        plant_id = session['plant_id']
        q = request.args.get('q', '').strip()
        if len(q) < 2:
            return jsonify([])
        db = get_db()
        rows = db.execute('''
            SELECT emp_code, name, designation, department
            FROM employees
            WHERE plant_id=? AND is_active=1
              AND (LOWER(name) LIKE LOWER(?) OR emp_code LIKE ?)
            ORDER BY name LIMIT 20
        ''', (plant_id, f'%{q}%', f'%{q}%')).fetchall()
        return jsonify([dict(r) for r in rows])

    # ── PUBLIC: landing ───────────────────────────────────────────────────────

    @app.route('/q/<token>')
    def qr_landing(token):
        db = get_db()
        try:
            qr = _validate_token(token, db)
        except Exception:
            return render_template('qr_error.html',
                                   msg='This QR code is invalid or has expired.'), 410
        lang = request.args.get('lang', 'en')
        if lang not in ('en', 'hi'):
            lang = 'en'
        if qr['stage'] == 'attendance':
            return render_template('qr_attendance.html', qr=qr, token=token,
                                   has_pin=bool(qr['session_pin']), error=None, lang=lang)
        return render_template('qr_feedback.html', qr=qr, token=token, error=None, lang=lang)

    # ── PUBLIC: thanks (PRG target — GET only) ────────────────────────────────

    @app.route('/q/<token>/thanks')
    def qr_thanks(token):
        db = get_db()
        try:
            qr = _validate_token(token, db, check_expiry=False)
        except Exception:
            return render_template('qr_error.html',
                                   msg='This QR code is invalid or has expired.'), 410
        msg      = request.args.get('msg', 'attendance_ok')
        emp_name = request.args.get('emp_name') or None
        return render_template('qr_thanks.html', qr=qr, msg=msg, emp_name=emp_name)

    # ── PUBLIC: employee name search (uses token for plant resolution) ─────────

    @app.route('/q/<token>/emp-search')
    def qr_emp_search(token):
        db = get_db()
        row = db.execute('SELECT plant_id, expires_at FROM session_qr WHERE token=? AND is_active=1',
                         (token,)).fetchone()
        if not row:
            return jsonify([])
        # Honour token expiry — an expired QR must not keep leaking the directory.
        if row['expires_at'] and _now_iso() > row['expires_at']:
            return jsonify([])
        q = request.args.get('q', '').strip()
        if len(q) < 2:
            return jsonify([])

        if row['plant_id'] == CENTRAL_PLANT_ID:
            # Cross-plant: search all plants + corp members
            plant_rows = db.execute('''
                SELECT e.emp_code, e.name, e.designation, e.department,
                       e.plant_id, p.name AS plant_name, p.unit_code
                FROM employees e
                JOIN plants p ON p.id=e.plant_id
                WHERE e.is_active=1
                  AND (LOWER(e.name) LIKE LOWER(?) OR e.emp_code LIKE ?)
                ORDER BY e.name LIMIT 20
            ''', (f'%{q}%', f'%{q}%')).fetchall()
            corp_rows = db.execute('''
                SELECT emp_code, name, designation, department,
                       99 AS plant_id, 'Corporate' AS plant_name, 'CEN' AS unit_code
                FROM corp_members
                WHERE is_active=1
                  AND (LOWER(name) LIKE LOWER(?) OR emp_code LIKE ?)
                ORDER BY name LIMIT 10
            ''', (f'%{q}%', f'%{q}%')).fetchall()
            results = [dict(r) for r in plant_rows] + [dict(r) for r in corp_rows]
            return jsonify(results[:25])
        else:
            rows = db.execute('''
                SELECT emp_code, name, designation, department,
                       plant_id, '' AS plant_name, '' AS unit_code
                FROM employees
                WHERE plant_id=? AND is_active=1
                  AND (LOWER(name) LIKE LOWER(?) OR emp_code LIKE ?)
                ORDER BY name LIMIT 15
            ''', (row['plant_id'], f'%{q}%', f'%{q}%')).fetchall()
            return jsonify([dict(r) for r in rows])

    # ── PUBLIC: submit attendance ─────────────────────────────────────────────

    @app.route('/q/<token>/attend', methods=['GET', 'POST'])
    def qr_attend(token):
        if request.method == 'GET':
            lang = request.args.get('lang')
            return redirect(url_for('qr_landing', token=token, lang=lang) if lang in ('en', 'hi')
                            else url_for('qr_landing', token=token), 302)
        db = get_db()
        try:
            qr = _validate_token(token, db)
        except Exception:
            return render_template('qr_error.html',
                                   msg='This QR code is invalid or has expired.'), 410

        lang = request.args.get('lang', request.form.get('lang', 'en'))
        if lang not in ('en', 'hi'):
            lang = 'en'

        def _att_err(error):
            return render_template('qr_attendance.html', qr=qr, token=token,
                                   has_pin=bool(qr['session_pin']), error=error, lang=lang)

        if _is_cross_origin_post():
            return _att_err('Could not verify the request origin. Please rescan the QR and try again.')

        if qr['session_pin']:
            entered_pin = request.form.get('session_pin', '').strip()
            # Constant-time compare to avoid leaking the PIN via response timing.
            if not hmac.compare_digest(entered_pin, str(qr['session_pin'])):
                return _att_err('Incorrect session code. Ask your trainer for the 4-digit code.')

        # GAP 6 (time gate): block scans before session start date
        if qr['plan_start']:
            today_iso = _today_ist().isoformat()
            if today_iso < qr['plan_start']:
                return _att_err(f'Session has not started. Attendance opens on {qr["plan_start"]}.')

        # Lock: attendance scan only while status='To Be Planned'.
        # Awaiting Verification / Conducted = 2C saved, central reviewing/done;
        # Lapsed / Cancelled / Re-Scheduled = session not happening as planned.
        # Allowing scans in any of these would mutate verified records.
        cal_status = db.execute(
            'SELECT status FROM calendar WHERE session_code=? AND plant_id=?',
            (qr['session_code'], qr['plant_id'])).fetchone()
        if cal_status and cal_status['status'] != 'To Be Planned':
            return _att_err(f'Attendance closed — session is "{cal_status["status"]}".')

        emp_code = request.form.get('emp_code', '').strip().upper()
        if not emp_code:
            return _att_err('Please enter or select your Employee Code.')

        is_central_session = (qr['plant_id'] == CENTRAL_PLANT_ID)
        # Write created_at explicitly in IST with a time component. The schema
        # default is date('now') = UTC, date-only — which both drifts by date
        # near IST midnight on Render AND gives the live monitor no scan time.
        scan_ts = _now_iso()

        if is_central_session:
            # Determine attendee's home plant from form (set by JS suggestion picker)
            try:
                attendee_plant_id = int(request.form.get('attendee_plant_id', '0') or 0)
            except (ValueError, TypeError):
                attendee_plant_id = 0

            emp = None
            if attendee_plant_id and attendee_plant_id != CENTRAL_PLANT_ID:
                emp = db.execute(
                    'SELECT name, collar, designation, department, plant_id FROM employees '
                    'WHERE plant_id=? AND emp_code=? AND is_active=1',
                    (attendee_plant_id, emp_code)
                ).fetchone()

            if not emp:
                # Try corp members
                corp = db.execute(
                    'SELECT name, designation, department FROM corp_members '
                    'WHERE emp_code=? AND is_active=1', (emp_code,)
                ).fetchone()
                if corp:
                    emp_plant = CENTRAL_PLANT_ID
                    emp_name  = corp['name']
                    host_pid  = CENTRAL_PLANT_ID
                else:
                    # Fallback: search all plants
                    found = db.execute(
                        'SELECT name, collar, designation, department, plant_id FROM employees '
                        'WHERE emp_code=? AND is_active=1 LIMIT 1', (emp_code,)
                    ).fetchone()
                    if not found:
                        return _att_err(f'Employee code "{emp_code}" not found.')
                    emp_plant = found['plant_id']
                    emp_name  = found['name']
                    host_pid  = CENTRAL_PLANT_ID
            else:
                emp_plant = emp['plant_id']
                emp_name  = emp['name']
                host_pid  = CENTRAL_PLANT_ID

            month = _date_to_month(qr['plan_start'] or '')
            # GAP 7: cal_new must be 'Calendar Program' (session exists in calendar)
            db.execute('''INSERT OR IGNORE INTO emp_training
                (plant_id, emp_code, session_code, programme_name, start_date, end_date,
                 hrs, prog_type, level, mode, cal_new, venue, month, host_plant_id, created_at)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
                (emp_plant, emp_code, qr['session_code'], qr['programme_name'],
                 qr['plan_start'] or '', qr['plan_end'] or '',
                 qr['duration_hrs'] or 0, qr['prog_type'] or '',
                 qr['level'] or '', qr['mode'] or '', 'Calendar Program',
                 '', month, host_pid, scan_ts))
            # Auto-update programme_master for attendee's plant + central
            for pid in ({emp_plant, CENTRAL_PLANT_ID}):
                db.execute('''INSERT OR IGNORE INTO programme_master(plant_id, name, prog_type, source)
                              VALUES(?,?,?,'New Requirement')''',
                           (pid, qr['programme_name'], qr['prog_type'] or ''))
        else:
            emp = db.execute(
                'SELECT name, collar, designation, department FROM employees '
                'WHERE plant_id=? AND emp_code=? AND is_active=1',
                (qr['plant_id'], emp_code)
            ).fetchone()
            if not emp:
                return _att_err(f'Employee code "{emp_code}" not found for {qr["plant_name"]}.')

            # Collar mismatch → ALLOW but flag (Central reviews via /anomalies)
            anom = []
            tgt = (qr['target_audience'] or '').strip()
            if tgt in ('Blue Collared', 'White Collared') and emp['collar'] and emp['collar'] != tgt:
                anom.append(f'collar_mismatch({emp["collar"]} vs {tgt})')

            emp_name = emp['name']
            month = _date_to_month(qr['plan_start'] or '')
            db.execute('''INSERT OR IGNORE INTO emp_training
                (plant_id, emp_code, session_code, programme_name, start_date, end_date,
                 hrs, prog_type, level, mode, cal_new, venue, month, anomaly_flags, created_at)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
                (qr['plant_id'], emp_code, qr['session_code'], qr['programme_name'],
                 qr['plan_start'] or '', qr['plan_end'] or '',
                 qr['duration_hrs'] or 0, qr['prog_type'] or '',
                 qr['level'] or '', qr['mode'] or '', 'Calendar Program',
                 '', month, ','.join(anom) if anom else None, scan_ts))

        changed = db.execute('SELECT changes()').fetchone()[0]
        # Refresh calendar actuals if this is a new insert
        if changed:
            _recompute_session_actuals(qr['plant_id'], qr['session_code'], db)
        db.commit()

        if changed == 0:
            return redirect(url_for('qr_thanks', token=token,
                                    msg='already_marked', emp_name=emp_name), 303)
        log_action('RECORD_ADD',
                   f"qr_attend:{emp_code}:{qr['session_code']}",
                   username=emp_code, plant_id=emp_plant if qr['plant_id'] == CENTRAL_PLANT_ID else qr['plant_id'])
        return redirect(url_for('qr_thanks', token=token,
                                msg='attendance_ok', emp_name=emp_name), 303)

    # ── PUBLIC: feedback form ─────────────────────────────────────────────────

    @app.route('/q/<token>/feedback', methods=['GET', 'POST'])
    def qr_feedback(token):
        db = get_db()
        try:
            qr = _validate_token(token, db)
        except Exception:
            return render_template('qr_error.html',
                                   msg='This QR code is invalid or has expired.'), 410

        lang = request.args.get('lang', request.form.get('lang', 'en'))
        if lang not in ('en', 'hi'):
            lang = 'en'

        if request.method == 'GET':
            return render_template('qr_feedback.html', qr=qr, token=token,
                                   error=None, lang=lang)

        if _is_cross_origin_post():
            return render_template('qr_feedback.html', qr=qr, token=token, lang=lang,
                                   error='Could not verify the request origin. Please rescan the QR and try again.')

        # Time gate: feedback opens once session starts
        if qr['plan_start']:
            today_iso = _today_ist().isoformat()
            if today_iso < qr['plan_start']:
                return render_template('qr_feedback.html', qr=qr, token=token, lang=lang,
                                       error=f'Feedback opens on {qr["plan_start"]}. Session has not started yet.')

        # Lock: feedback submit only while status='To Be Planned'.
        # Once 2C is saved, feedback aggregates feed verified programme_details —
        # accepting new responses would silently mutate Central-reviewed data.
        cal_status = db.execute(
            'SELECT status FROM calendar WHERE session_code=? AND plant_id=?',
            (qr['session_code'], qr['plant_id'])).fetchone()
        if cal_status and cal_status['status'] != 'To Be Planned':
            return render_template('qr_feedback.html', qr=qr, token=token, lang=lang,
                                   error=f'Feedback closed — session is "{cal_status["status"]}".')

        emp_code = request.form.get('emp_code', '').strip().upper() or None
        ip = request.remote_addr or ''
        if emp_code:
            if qr['plant_id'] == CENTRAL_PLANT_ID:
                ok = (
                    db.execute('SELECT 1 FROM corp_members WHERE emp_code=? AND is_active=1',
                               (emp_code,)).fetchone() or
                    db.execute('SELECT 1 FROM employees WHERE emp_code=? AND is_active=1',
                               (emp_code,)).fetchone()
                )
            else:
                ok = db.execute(
                    'SELECT 1 FROM employees WHERE plant_id=? AND emp_code=? AND is_active=1',
                    (qr['plant_id'], emp_code)
                ).fetchone()
            if not ok:
                return render_template('qr_feedback.html', qr=qr, token=token, lang=lang,
                                       error=f'Employee code "{emp_code}" not found.')
        else:
            # Anonymous: deduplicate by IP so same device can't spam
            already = db.execute(
                'SELECT 1 FROM feedback_response WHERE plant_id=? AND session_code=? AND emp_code IS NULL AND ip_address=?',
                (qr['plant_id'], qr['session_code'], ip)
            ).fetchone()
            if already:
                return redirect(url_for('qr_thanks', token=token, msg='feedback_ok'), 303)

        def _r(name):
            try:
                v = int(request.form.get(name, ''))
                return v if 1 <= v <= 4 else None
            except (ValueError, TypeError):
                return None

        # submitted_at written explicitly in IST. Schema default is
        # datetime('now','localtime') = UTC on Render (server TZ is UTC), so the
        # default drifts feedback timestamps 5.5h behind IST. Set it ourselves.
        # INSERT OR IGNORE (not REPLACE): once a (plant,session,emp_code) feedback
        # row exists it is immutable, so a person who guesses another's emp_code
        # cannot overwrite genuine feedback. First submission wins.
        db.execute('''INSERT OR IGNORE INTO feedback_response
            (plant_id, session_code, emp_code, submitted_at,
             q_obj_explained, q_well_structured, q_content_appropriate,
             q_presentation_quality, q_time_reasonable,
             q_inputs_appropriate, q_communication_clear,
             q_queries_responded, q_well_involved,
             key_learnings, suggestions, ip_address, lang)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
            (qr['plant_id'], qr['session_code'], emp_code, _now_iso(),
             _r('q1'), _r('q2'), _r('q3'), _r('q4'), _r('q5'),
             _r('q6'), _r('q7'), _r('q8'), _r('q9'),
             request.form.get('key_learnings', '').strip()[:1000],
             request.form.get('suggestions', '').strip()[:1000],
             request.remote_addr, lang))
        db.commit()
        _recompute_feedback_aggregates(qr['plant_id'], qr['session_code'], db)
        db.commit()
        log_action('RECORD_ADD',
                   f"qr_feedback:{emp_code or 'anon'}:{qr['session_code']}",
                   username=emp_code or 'anonymous', plant_id=qr['plant_id'])
        return redirect(url_for('qr_thanks', token=token, msg='feedback_ok'), 303)
