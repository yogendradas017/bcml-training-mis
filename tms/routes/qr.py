import io
import secrets
import datetime

import qrcode
from flask import (abort, flash, jsonify, redirect, render_template,
                   request, send_file, session, url_for)

from tms.db import get_db
from tms.decorators import spoc_required
from tms.helpers import _date_to_month


# ── helpers ──────────────────────────────────────────────────────────────────

def _now_iso():
    return datetime.datetime.now().isoformat(timespec='seconds')


def _avg(vals):
    v = [x for x in vals if x is not None]
    return round(sum(v) / len(v), 2) if v else None


def _validate_token(token, db, check_expiry=True):
    row = db.execute('''
        SELECT q.*, c.programme_name, c.prog_type, c.level, c.mode,
               c.plan_start, c.plan_end, c.duration_hrs, c.venue,
               c.time_from, c.time_to, c.target_audience,
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
                   course_feedback=?, faculty_feedback=?,
                   trainer_fb_participants=?, trainer_fb_facilities=?
                  WHERE plant_id=? AND session_code=?''',
               (prog_avg, trainer_avg, row['q8'], row['q9'], plant_id, session_code))


def _make_qr_png(url):
    img = qrcode.make(url)
    buf = io.BytesIO()
    img.save(buf, format='PNG')
    buf.seek(0)
    return buf


# ── register ─────────────────────────────────────────────────────────────────

def _register(app):

    # ── SPOC: generate QR for a calendar session ─────────────────────────────

    @app.route('/calendar/<int:cal_id>/qr/generate', methods=['POST'])
    @spoc_required
    def qr_generate(cal_id):
        plant_id = session['plant_id']
        db = get_db()
        cal = db.execute('SELECT * FROM calendar WHERE id=? AND plant_id=?',
                         (cal_id, plant_id)).fetchone()
        if not cal:
            flash('Session not found.', 'danger')
            return redirect(url_for('training_calendar'))

        stage = request.form.get('stage', 'attendance')
        if stage not in ('attendance', 'feedback'):
            stage = 'attendance'

        # Calculate expiry: plan_end + 1 day (or today + 7 days if no plan_end)
        try:
            plan_end = datetime.date.fromisoformat(cal['plan_end'])
            expires  = (plan_end + datetime.timedelta(days=1)).isoformat() + 'T23:59:59'
        except Exception:
            expires = (datetime.date.today() + datetime.timedelta(days=7)).isoformat() + 'T23:59:59'

        existing = db.execute(
            'SELECT id FROM session_qr WHERE plant_id=? AND session_code=? AND stage=?',
            (plant_id, cal['session_code'], stage)
        ).fetchone()

        if existing:
            # Reactivate + refresh token
            new_token = secrets.token_urlsafe(16)
            db.execute('''UPDATE session_qr SET token=?, is_active=1, expires_at=?,
                          created_at=datetime('now','localtime'), created_by=?
                          WHERE id=?''',
                       (new_token, expires, session.get('user_id'), existing['id']))
            db.commit()
            flash(f'QR regenerated for {stage}.', 'success')
        else:
            new_token = secrets.token_urlsafe(16)
            db.execute('''INSERT INTO session_qr
                (plant_id, session_code, token, stage, expires_at, created_by)
                VALUES(?,?,?,?,?,?)''',
                (plant_id, cal['session_code'], new_token, stage, expires,
                 session.get('user_id')))
            db.commit()
            flash(f'QR generated for {stage}.', 'success')

        return redirect(url_for('qr_poster', token=new_token))

    # ── SPOC: QR image (PNG stream) ───────────────────────────────────────────

    @app.route('/qr/<token>/image.png')
    @spoc_required
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
    @spoc_required
    def qr_poster(token):
        db = get_db()
        qr = _validate_token(token, db, check_expiry=False)
        qr_url = request.host_url + f'q/{token}'
        return render_template('qr_poster.html', qr=qr, token=token, qr_url=qr_url)

    # ── SPOC: revoke QR ───────────────────────────────────────────────────────

    @app.route('/qr/<int:qr_id>/revoke', methods=['POST'])
    @spoc_required
    def qr_revoke(qr_id):
        plant_id = session['plant_id']
        db = get_db()
        db.execute('UPDATE session_qr SET is_active=0 WHERE id=? AND plant_id=?',
                   (qr_id, plant_id))
        db.commit()
        flash('QR revoked — old QR no longer accepts scans.', 'warning')
        return redirect(url_for('training_calendar'))

    # ── SPOC: live attendance monitor ─────────────────────────────────────────

    @app.route('/calendar/<int:cal_id>/live')
    @spoc_required
    def qr_live(cal_id):
        plant_id = session['plant_id']
        db = get_db()
        cal = db.execute('SELECT * FROM calendar WHERE id=? AND plant_id=?',
                         (cal_id, plant_id)).fetchone()
        if not cal:
            flash('Session not found.', 'danger')
            return redirect(url_for('training_calendar'))

        qr_rows = db.execute(
            'SELECT * FROM session_qr WHERE plant_id=? AND session_code=? ORDER BY stage',
            (plant_id, cal['session_code'])
        ).fetchall()

        attendees = db.execute('''
            SELECT t.emp_code, t.created_at, e.name, e.designation, e.department, e.collar
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
                               attendees=attendees, fb_count=fb_count)

    # ── SPOC: live JSON poll ──────────────────────────────────────────────────

    @app.route('/api/qr/<int:cal_id>/live.json')
    @spoc_required
    def qr_live_json(cal_id):
        plant_id = session['plant_id']
        db = get_db()
        cal = db.execute('SELECT session_code, planned_pax FROM calendar WHERE id=? AND plant_id=?',
                         (cal_id, plant_id)).fetchone()
        if not cal:
            return jsonify({'error': 'not found'}), 404

        rows = db.execute('''
            SELECT t.emp_code, t.created_at, e.name, e.designation, e.department
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
        if qr['stage'] == 'attendance':
            return render_template('qr_attendance.html', qr=qr, token=token, error=None)
        return render_template('qr_feedback.html', qr=qr, token=token, error=None, lang='en')

    # ── PUBLIC: employee name search (uses token for plant resolution) ─────────

    @app.route('/q/<token>/emp-search')
    def qr_emp_search(token):
        db = get_db()
        row = db.execute('SELECT plant_id FROM session_qr WHERE token=? AND is_active=1',
                         (token,)).fetchone()
        if not row:
            return jsonify([])
        q = request.args.get('q', '').strip()
        if len(q) < 2:
            return jsonify([])
        rows = db.execute('''
            SELECT emp_code, name, designation, department
            FROM employees
            WHERE plant_id=? AND is_active=1
              AND (LOWER(name) LIKE LOWER(?) OR emp_code LIKE ?)
            ORDER BY name LIMIT 15
        ''', (row['plant_id'], f'%{q}%', f'%{q}%')).fetchall()
        return jsonify([dict(r) for r in rows])

    # ── PUBLIC: submit attendance ─────────────────────────────────────────────

    @app.route('/q/<token>/attend', methods=['POST'])
    def qr_attend(token):
        db = get_db()
        try:
            qr = _validate_token(token, db)
        except Exception:
            return render_template('qr_error.html',
                                   msg='This QR code is invalid or has expired.'), 410

        emp_code = request.form.get('emp_code', '').strip().upper()
        if not emp_code:
            return render_template('qr_attendance.html', qr=qr, token=token,
                                   error='Please enter or select your Employee Code.')

        emp = db.execute(
            'SELECT name, collar, designation, department FROM employees '
            'WHERE plant_id=? AND emp_code=? AND is_active=1',
            (qr['plant_id'], emp_code)
        ).fetchone()
        if not emp:
            return render_template('qr_attendance.html', qr=qr, token=token,
                                   error=f'Employee code "{emp_code}" not found for {qr["plant_name"]}.')

        month = _date_to_month(qr['plan_start'] or '')
        db.execute('''INSERT OR IGNORE INTO emp_training
            (plant_id, emp_code, session_code, programme_name, start_date, end_date,
             hrs, prog_type, level, mode, cal_new, venue, month)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)''',
            (qr['plant_id'], emp_code, qr['session_code'], qr['programme_name'],
             qr['plan_start'] or '', qr['plan_end'] or '',
             qr['duration_hrs'] or 0, qr['prog_type'] or '',
             qr['level'] or '', qr['mode'] or '', 'Calendar Program',
             qr['venue'] or '', month))
        changed = db.execute('SELECT changes()').fetchone()[0]
        db.commit()

        if changed == 0:
            return render_template('qr_thanks.html', qr=qr,
                                   msg='already_marked', emp_name=emp['name'])
        return render_template('qr_thanks.html', qr=qr,
                               msg='attendance_ok', emp_name=emp['name'])

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

        emp_code = request.form.get('emp_code', '').strip().upper() or None
        if emp_code:
            ok = db.execute(
                'SELECT 1 FROM employees WHERE plant_id=? AND emp_code=? AND is_active=1',
                (qr['plant_id'], emp_code)
            ).fetchone()
            if not ok:
                return render_template('qr_feedback.html', qr=qr, token=token, lang=lang,
                                       error=f'Employee code "{emp_code}" not found.')

        def _r(name):
            try:
                v = int(request.form.get(name, ''))
                return v if 1 <= v <= 4 else None
            except (ValueError, TypeError):
                return None

        db.execute('''INSERT OR REPLACE INTO feedback_response
            (plant_id, session_code, emp_code,
             q_obj_explained, q_well_structured, q_content_appropriate,
             q_presentation_quality, q_time_reasonable,
             q_inputs_appropriate, q_communication_clear,
             q_queries_responded, q_well_involved,
             key_learnings, suggestions, ip_address, lang)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
            (qr['plant_id'], qr['session_code'], emp_code,
             _r('q1'), _r('q2'), _r('q3'), _r('q4'), _r('q5'),
             _r('q6'), _r('q7'), _r('q8'), _r('q9'),
             request.form.get('key_learnings', '').strip()[:1000],
             request.form.get('suggestions', '').strip()[:1000],
             request.remote_addr, lang))
        db.commit()
        _recompute_feedback_aggregates(qr['plant_id'], qr['session_code'], db)
        db.commit()

        return render_template('qr_thanks.html', qr=qr,
                               msg='feedback_ok', emp_name=None)
