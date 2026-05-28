from datetime import datetime as _dt

from flask import render_template, request, redirect, url_for, session, flash

from tms.db import get_db
from tms.decorators import central_required
from tms.audit import log_action


def _register(app):

    @app.route('/verify-sessions')
    @central_required
    def verify_sessions():
        db = get_db()
        pending = db.execute('''
            SELECT c.id, c.session_code, c.programme_name, c.prog_type,
                   c.plant_id, p.name AS plant_name, p.unit_code,
                   c.plan_start, c.plan_end, c.planned_pax,
                   c.actual_pax, c.actual_hrs, c.conducted_at, c.conducted_by,
                   c.target_audience, c.duration_hrs,
                   u.username AS conducted_by_name,
                   (SELECT detail FROM verification_log v
                    WHERE v.session_code=c.session_code AND v.plant_id=c.plant_id
                      AND v.stage='2c_added'
                    ORDER BY v.id DESC LIMIT 1) AS anomalies
            FROM calendar c
            JOIN plants p ON p.id=c.plant_id
            LEFT JOIN users u ON u.id=c.conducted_by
            WHERE c.status='Awaiting Verification'
            ORDER BY c.conducted_at DESC
        ''').fetchall()
        return render_template('verify_sessions.html', pending=pending)

    @app.route('/verify-sessions/<session_code>/<int:plant_id>/approve', methods=['POST'])
    @central_required
    def verify_approve(session_code, plant_id):
        db = get_db()
        cal = db.execute('SELECT status FROM calendar WHERE session_code=? AND plant_id=?',
                         (session_code, plant_id)).fetchone()
        if not cal:
            flash('Session not found.', 'danger')
            return redirect(url_for('verify_sessions'))
        if cal['status'] != 'Awaiting Verification':
            flash(f'Session is "{cal["status"]}" — cannot approve.', 'warning')
            return redirect(url_for('verify_sessions'))
        now_iso  = _dt.now().isoformat(timespec='seconds')
        user_id  = session.get('user_id')
        username = session.get('username', '')
        note     = request.form.get('note', '').strip()[:500]
        db.execute(
            "UPDATE calendar SET status='Conducted', verified_at=?, verified_by=? "
            "WHERE session_code=? AND plant_id=?",
            (now_iso, user_id, session_code, plant_id))
        db.execute(
            'INSERT INTO verification_log (session_code, plant_id, stage, actor, actor_id, detail) '
            'VALUES (?,?,?,?,?,?)',
            (session_code, plant_id, 'verified', username, user_id, note or 'approved'))
        db.commit()
        log_action('VERIFY_APPROVE', f"session:{session_code}", plant_id=plant_id)
        flash(f'Session {session_code} verified — now counted as Conducted.', 'success')
        return redirect(url_for('verify_sessions'))

    @app.route('/verify-sessions/<session_code>/<int:plant_id>/reject', methods=['POST'])
    @central_required
    def verify_reject(session_code, plant_id):
        db = get_db()
        cal = db.execute('SELECT status FROM calendar WHERE session_code=? AND plant_id=?',
                         (session_code, plant_id)).fetchone()
        if not cal:
            flash('Session not found.', 'danger')
            return redirect(url_for('verify_sessions'))
        if cal['status'] != 'Awaiting Verification':
            flash(f'Session is "{cal["status"]}" — cannot reject.', 'warning')
            return redirect(url_for('verify_sessions'))
        note = request.form.get('note', '').strip()[:500]
        if not note:
            flash('Rejection note is required — explain why this is being sent back.', 'danger')
            return redirect(url_for('verify_sessions'))
        username = session.get('username', '')
        user_id  = session.get('user_id')
        db.execute(
            "UPDATE calendar SET status='To Be Planned', conducted_at=NULL, conducted_by=NULL, "
            "verified_at=NULL, verified_by=NULL, actual_pax=0, actual_hrs=0 "
            "WHERE session_code=? AND plant_id=?",
            (session_code, plant_id))
        db.execute('DELETE FROM programme_details WHERE session_code=? AND plant_id=?',
                   (session_code, plant_id))
        db.execute(
            'INSERT INTO verification_log (session_code, plant_id, stage, actor, actor_id, detail) '
            'VALUES (?,?,?,?,?,?)',
            (session_code, plant_id, 'rejected', username, user_id, note))
        db.commit()
        log_action('VERIFY_REJECT', f"session:{session_code}:{note[:100]}", plant_id=plant_id)
        flash(f'Session {session_code} rejected. 2C removed — SPOC must re-record.', 'warning')
        return redirect(url_for('verify_sessions'))

    @app.route('/verify-sessions/<session_code>/<int:plant_id>/trail')
    @central_required
    def verify_trail(session_code, plant_id):
        db = get_db()
        trail = db.execute(
            'SELECT v.*, u.username AS actor_name '
            'FROM verification_log v '
            'LEFT JOIN users u ON u.id=v.actor_id '
            'WHERE v.session_code=? AND v.plant_id=? ORDER BY v.ts DESC',
            (session_code, plant_id)).fetchall()
        cal = db.execute(
            'SELECT c.*, p.name AS plant_name FROM calendar c '
            'JOIN plants p ON p.id=c.plant_id '
            'WHERE c.session_code=? AND c.plant_id=?',
            (session_code, plant_id)).fetchone()
        attendees = db.execute(
            'SELECT t.emp_code, t.hrs, t.start_date, t.created_at, '
            'e.name, e.collar, e.designation, e.department '
            'FROM emp_training t '
            'LEFT JOIN employees e ON e.plant_id=t.plant_id AND e.emp_code=t.emp_code '
            'WHERE t.plant_id=? AND t.session_code=? '
            'ORDER BY t.created_at',
            (plant_id, session_code)).fetchall()
        return render_template('verify_trail.html', trail=trail, cal=cal, attendees=attendees)
