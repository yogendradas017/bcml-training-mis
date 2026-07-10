import json
import logging

from flask import render_template, request, redirect, url_for, session, flash

from tms.db import get_db
from tms.decorators import spoc_required, admin_required
from tms.audit import log_action, log_record_change
from tms.helpers import _now_ist, _today_ist, _fy_label


PAYLOAD_MAX_BYTES = 8 * 1024  # cap typed payload at 8 KB


def _build_payload(req_type, form, plant_id, db):
    """Validate + assemble typed payload per req_type. Returns (payload_dict,
    error_message). On success error_message is None."""
    if req_type == 'TNI_ADD':
        emp_code = (form.get('p_emp_code') or '').strip()
        prog_name = (form.get('p_programme_name') or '').strip()
        prog_type = (form.get('p_prog_type') or '').strip()
        if not (emp_code and prog_name and prog_type):
            return None, 'TNI Add requires Employee Code, Programme Name, and Programme Type.'
        emp = db.execute(
            "SELECT 1 FROM employees WHERE plant_id=? AND emp_code=? AND is_active=1 LIMIT 1",
            (plant_id, emp_code)).fetchone()
        if not emp:
            return None, f'Employee {emp_code} not found / not Active in this plant.'
        return {'emp_code': emp_code, 'programme_name': prog_name, 'prog_type': prog_type}, None

    if req_type == 'MARK_CONDUCTED':
        sc = (form.get('p_session_code') or '').strip()
        cdate = (form.get('p_conducted_date') or '').strip()
        faculty = (form.get('p_faculty') or '').strip()
        if not (sc and cdate and faculty):
            return None, 'Mark Conducted requires Session Code, Conducted Date, and Faculty.'
        cal = db.execute(
            'SELECT 1 FROM calendar WHERE plant_id=? AND session_code=? LIMIT 1',
            (plant_id, sc)).fetchone()
        if not cal:
            return None, f'Session {sc} not found in calendar for this plant.'
        return {'session_code': sc, 'conducted_date': cdate, 'faculty': faculty}, None

    if req_type == 'MANUAL_ATTENDANCE':
        sc = (form.get('p_session_code') or '').strip()
        emp_code = (form.get('p_emp_code') or '').strip()
        adate = (form.get('p_attendance_date') or '').strip()
        if not (sc and emp_code and adate):
            return None, 'Manual Attendance requires Session Code, Employee Code, and Attendance Date.'
        pd_row = db.execute(
            'SELECT programme_name, prog_type, hours_actual FROM programme_details '
            'WHERE plant_id=? AND session_code=? LIMIT 1',
            (plant_id, sc)).fetchone()
        if not pd_row:
            return None, f'Session {sc} has no 2C record yet — cannot back-fill attendance.'
        emp = db.execute(
            "SELECT 1 FROM employees WHERE plant_id=? AND emp_code=? AND is_active=1 LIMIT 1",
            (plant_id, emp_code)).fetchone()
        if not emp:
            return None, f'Employee {emp_code} not found / not Active in this plant.'
        return {'session_code': sc, 'emp_code': emp_code,
                'attendance_date': adate,
                'programme_name': pd_row['programme_name'],
                'prog_type': pd_row['prog_type'],
                'hrs': pd_row['hours_actual'] or 0}, None

    return {}, None


def _execute_request(req, payload, db, actor_username, actor_user_id):
    """Execute the approved request against req['plant_id']. Returns
    (success_bool, message). All writes are idempotent so re-approve is safe."""
    rt = req['request_type']
    plant_id = req['plant_id']
    provenance = f'via spoc_request:{req["id"]} plant:{plant_id} requested_by:{req["requested_by"]} approved_by:{actor_username}'

    if rt == 'TNI_ADD':
        fy = _fy_label()
        cur = db.execute(
            'INSERT OR IGNORE INTO tni '
            '(plant_id, emp_code, programme_name, prog_type, fy_year) '
            'VALUES (?, ?, ?, ?, ?)',
            (plant_id, payload['emp_code'], payload['programme_name'],
             payload['prog_type'], fy))
        if cur.rowcount:
            new_id = cur.lastrowid
            after = db.execute('SELECT * FROM tni WHERE id=?', (new_id,)).fetchone()
            log_record_change('RECORD_ADD', new_id, 'tni',
                              before=None,
                              after=dict(after) if after else None,
                              extra_detail=provenance)
            return True, f'TNI row added for {payload["emp_code"]} / {payload["programme_name"]}.'
        return True, 'TNI row already exists (idempotent — no duplicate created).'

    if rt == 'MARK_CONDUCTED':
        sc = payload['session_code']
        cdate = payload['conducted_date']
        cal_before = db.execute(
            'SELECT * FROM calendar WHERE plant_id=? AND session_code=?',
            (plant_id, sc)).fetchone()
        if not cal_before:
            return False, f'Session {sc} disappeared from calendar — aborting.'
        ts = _now_ist().isoformat(timespec='seconds')
        db.execute(
            "UPDATE calendar SET status='Conducted', conducted_at=?, conducted_by=?, "
            "verified_at=?, verified_by=? WHERE plant_id=? AND session_code=?",
            (ts, actor_user_id, ts, actor_user_id, plant_id, sc))
        cal_after = db.execute(
            'SELECT * FROM calendar WHERE plant_id=? AND session_code=?',
            (plant_id, sc)).fetchone()
        log_record_change('RECORD_EDIT', cal_before['id'], 'calendar',
                          before=dict(cal_before),
                          after=dict(cal_after) if cal_after else None,
                          extra_detail=provenance)
        existing_pd = db.execute(
            'SELECT 1 FROM programme_details WHERE plant_id=? AND session_code=? LIMIT 1',
            (plant_id, sc)).fetchone()
        if not existing_pd:
            db.execute(
                'INSERT INTO programme_details '
                '(plant_id, session_code, programme_name, prog_type, start_date, end_date, faculty_name, cal_new) '
                'VALUES (?, ?, ?, ?, ?, ?, ?, ?)',
                (plant_id, sc, cal_before['programme_name'], cal_before['prog_type'],
                 cdate, cdate, payload['faculty'], 'Calendar Program'))
            log_record_change('RECORD_ADD', sc, 'programme_details',
                              before=None, after=None,
                              extra_detail=provenance + ' (stub from override)')
        return True, f'Session {sc} marked Conducted.'

    if rt == 'MANUAL_ATTENDANCE':
        sc = payload['session_code']
        emp_code = payload['emp_code']
        dup = db.execute(
            'SELECT 1 FROM emp_training WHERE plant_id=? AND session_code=? AND emp_code=? LIMIT 1',
            (plant_id, sc, emp_code)).fetchone()
        if dup:
            return True, f'Attendance for {emp_code} on {sc} already recorded (idempotent).'
        cur = db.execute(
            'INSERT INTO emp_training '
            '(plant_id, emp_code, session_code, programme_name, prog_type, start_date, end_date, hrs) '
            'VALUES (?, ?, ?, ?, ?, ?, ?, ?)',
            (plant_id, emp_code, sc, payload['programme_name'], payload['prog_type'],
             payload['attendance_date'], payload['attendance_date'], payload['hrs']))
        new_id = cur.lastrowid
        after = db.execute('SELECT * FROM emp_training WHERE id=?', (new_id,)).fetchone()
        log_record_change('RECORD_ADD', new_id, 'emp_training',
                          before=None,
                          after=dict(after) if after else None,
                          extra_detail=provenance)
        return True, f'Manual attendance recorded for {emp_code} on {sc}.'

    if rt == 'OTHER':
        return True, 'Marked Approved — admin must take any required action manually.'

    return False, f'Unknown request type: {rt}'


def _register(app):

    @app.route('/requests/submit', methods=['GET', 'POST'])
    @spoc_required
    def spoc_submit_request():
        plant_id = session['plant_id']
        db = get_db()
        if request.method == 'POST':
            req_type = request.form.get('request_type', '').strip()
            details  = request.form.get('details', '').strip()
            valid_types = ('TNI_ADD', 'MARK_CONDUCTED', 'MANUAL_ATTENDANCE', 'OTHER')
            if req_type not in valid_types:
                flash('Invalid request type.', 'danger')
                return redirect(url_for('spoc_submit_request'))

            payload, err = _build_payload(req_type, request.form, plant_id, db)
            if err:
                flash(err, 'danger')
                return redirect(url_for('spoc_submit_request'))

            min_len = 20 if req_type == 'OTHER' else 10
            if not details or len(details) < min_len:
                flash(f'Please provide a description (at least {min_len} characters).', 'danger')
                return redirect(url_for('spoc_submit_request'))

            payload_str = json.dumps(payload) if payload else None
            if payload_str and len(payload_str.encode('utf-8')) > PAYLOAD_MAX_BYTES:
                flash('Payload too large — please simplify the request.', 'danger')
                return redirect(url_for('spoc_submit_request'))

            db.execute(
                'INSERT INTO spoc_requests(plant_id, requested_by, request_type, details, payload_json, ts) '
                'VALUES(?, ?, ?, ?, ?, ?)',
                (plant_id, session['username'], req_type, details[:2000], payload_str,
                 _now_ist().isoformat(timespec='seconds'))
            )
            db.commit()
            log_action('RECORD_ADD', f'spoc_request:{req_type}')
            flash('Override request submitted. Admin will review and respond shortly.', 'success')
            return redirect(url_for('spoc_submit_request'))

        my_requests = db.execute(
            'SELECT * FROM spoc_requests WHERE plant_id=? ORDER BY ts DESC LIMIT 50',
            (plant_id,)
        ).fetchall()
        return render_template('spoc_request_form.html', my_requests=my_requests)

    @app.route('/admin/requests')
    @admin_required
    def admin_requests():
        db = get_db()
        rows = db.execute('''
            SELECT r.*, p.name as plant_name
            FROM spoc_requests r
            LEFT JOIN plants p ON p.id = r.plant_id
            ORDER BY CASE r.status WHEN 'Pending' THEN 0 ELSE 1 END, r.ts DESC
            LIMIT 200
        ''').fetchall()
        return render_template('admin_requests.html', rows=rows)

    @app.route('/admin/requests/<int:req_id>/review', methods=['POST'])
    @admin_required
    def admin_review_request(req_id):
        action      = request.form.get('action', '')
        review_note = request.form.get('review_note', '').strip()
        if action not in ('approve', 'reject'):
            flash('Invalid action.', 'danger')
            return redirect(url_for('admin_requests'))

        db = get_db()
        req = db.execute('SELECT * FROM spoc_requests WHERE id=?', (req_id,)).fetchone()
        if not req:
            flash('Request not found.', 'danger')
            return redirect(url_for('admin_requests'))
        if req['status'] != 'Pending':
            flash(f'Request is already {req["status"]} — cannot re-review.', 'warning')
            return redirect(url_for('admin_requests'))

        now_iso = _now_ist().isoformat(timespec='seconds')
        actor_username = session['username']
        actor_user_id  = session.get('user_id')

        if action == 'reject':
            db.execute(
                "UPDATE spoc_requests SET status='Rejected', reviewed_by=?, reviewed_at=?, review_note=? WHERE id=?",
                (actor_username, now_iso, review_note[:500], req_id))
            db.commit()
            log_action('RECORD_EDIT', f'spoc_request:{req_id}:Rejected')
            flash(f'Request {req_id} rejected.', 'warning')
            return redirect(url_for('admin_requests'))

        # APPROVE path
        try:
            payload = json.loads(req['payload_json'] or '{}')
        except json.JSONDecodeError:
            flash('Request has no/invalid structured payload (legacy free-text request). '
                  'Reject and ask SPOC to resubmit via the typed form.', 'danger')
            return redirect(url_for('admin_requests'))

        original_plant = session.get('plant_id')
        session['plant_id'] = req['plant_id']
        exec_msg = ''
        try:
            try:
                ok, exec_msg = _execute_request(req, payload, db,
                                                actor_username, actor_user_id)
                if not ok:
                    db.rollback()
                    flash(f'Approve failed: {exec_msg} — request left Pending.', 'danger')
                    return redirect(url_for('admin_requests'))
                db.execute(
                    "UPDATE spoc_requests SET status='Approved', reviewed_by=?, reviewed_at=?, review_note=? WHERE id=?",
                    (actor_username, now_iso, review_note[:500], req_id))
                db.commit()
            except Exception:
                db.rollback()
                logging.exception('admin_review_request execute failed for req_id=%s', req_id)
                flash('Approve failed — transaction rolled back. Request left Pending. Check logs.',
                      'danger')
                return redirect(url_for('admin_requests'))
        finally:
            session['plant_id'] = original_plant

        log_action('RECORD_EDIT', f'spoc_request:{req_id}:Approved')
        flash(f'Request {req_id} approved. {exec_msg}', 'success')
        return redirect(url_for('admin_requests'))
