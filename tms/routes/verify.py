import logging

from flask import render_template, request, redirect, url_for, session, flash

from tms.db import get_db
from tms.decorators import central_required
from tms.audit import log_action, log_record_change
from tms.helpers import _now_ist


def seed_effectiveness_reviews(plant_id, session_code, after_snap, db):
    """SOP: Specialized programmes trigger 3-month post-training effectiveness
    review for every attendee. Seeds pending review rows so SPOC can track and
    central/admin sees pending count. Caller owns db.commit().

    Returns (seeded_count, due_date_iso). Returns (0, None) when programme is
    not Specialized or after_snap is missing. Set-based INSERT...SELECT keeps
    SQLite write-lock hold time at single-statement latency.
    """
    if not after_snap:
        return 0, None
    if (after_snap['category'] or 'General') != 'Specialized':
        return 0, None
    from datetime import date as _d, timedelta as _td
    conducted_date = (after_snap['plan_end'] or after_snap['plan_start'] or
                      _now_ist().isoformat(timespec='seconds')[:10])
    try:
        due_date = (_d.fromisoformat(conducted_date) + _td(days=90)).isoformat()
    except (ValueError, TypeError):
        due_date = conducted_date  # degenerate fallback
    cur = db.execute(
        'INSERT OR IGNORE INTO effectiveness_review '
        '(plant_id, session_code, emp_code, conducted_date, due_date) '
        'SELECT ?, ?, emp_code, ?, ? '
        'FROM (SELECT DISTINCT emp_code FROM emp_training '
        '      WHERE plant_id=? AND session_code=?)',
        (plant_id, session_code, conducted_date, due_date,
         plant_id, session_code))
    return cur.rowcount or 0, due_date


# Tier 6: verification checklist items — all must be ticked before approve.
# Adding/changing this list immediately tightens the gate everywhere.
VERIFY_CHECKLIST_ITEMS = [
    ('reviewed_attendance', 'Attendance (2A) reviewed — every attendee verified'),
    ('reviewed_feedback',   'Feedback responses reviewed — coverage and scores plausible'),
    ('reviewed_2c',         'Programme Details (2C) reviewed — faculty, duration, mode correct'),
    ('reviewed_anomalies',  'Anomaly flags reviewed and resolved'),
]
VERIFY_NOTE_MIN_LEN = 20


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
        cal = db.execute('SELECT * FROM calendar WHERE session_code=? AND plant_id=?',
                         (session_code, plant_id)).fetchone()
        if not cal:
            flash('Session not found.', 'danger')
            return redirect(url_for('verify_sessions'))
        if cal['status'] != 'Awaiting Verification':
            flash(f'Session is "{cal["status"]}" — cannot approve.', 'warning')
            return redirect(url_for('verify_sessions'))

        # Tier 6: enforce verification checklist + mandatory note
        missing_checks = [label for key, label in VERIFY_CHECKLIST_ITEMS
                          if request.form.get(key) != '1']
        note = request.form.get('note', '').strip()
        if missing_checks:
            flash(
                'Cannot approve. Required checklist items not confirmed: ' +
                ' · '.join(missing_checks),
                'danger')
            return redirect(url_for('verify_trail', session_code=session_code,
                                     plant_id=plant_id))
        if len(note) < VERIFY_NOTE_MIN_LEN:
            flash(
                f'Verification note must be at least {VERIFY_NOTE_MIN_LEN} characters '
                f'(got {len(note)}). State what you reviewed and any caveats.',
                'danger')
            return redirect(url_for('verify_trail', session_code=session_code,
                                     plant_id=plant_id))
        note = note[:500]

        now_iso  = _now_ist().isoformat(timespec='seconds')
        user_id  = session.get('user_id')
        username = session.get('username', '')
        before_snap_dict = dict(cal)

        # Atomic block: calendar flip + verification_log + effectiveness seeding
        # must all succeed or all roll back. Mid-loop failure that committed
        # 'Conducted' without seeding would leave the session stuck — the state
        # guard above would then block any retry.
        seeded = 0
        due_date = None
        after_snap = None
        try:
            db.execute(
                "UPDATE calendar SET status='Conducted', verified_at=?, verified_by=? "
                "WHERE session_code=? AND plant_id=?",
                (now_iso, user_id, session_code, plant_id))
            after_snap = db.execute(
                'SELECT * FROM calendar WHERE session_code=? AND plant_id=?',
                (session_code, plant_id)).fetchone()
            db.execute(
                'INSERT INTO verification_log (session_code, plant_id, stage, actor, actor_id, detail, ts) '
                'VALUES (?,?,?,?,?,?,?)',
                (session_code, plant_id, 'verified', username, user_id,
                 f'approved; checklist=all; note: {note}',
                 _now_ist().isoformat(timespec='seconds')))
            seeded, due_date = seed_effectiveness_reviews(
                plant_id, session_code, after_snap, db)
            db.commit()
        except Exception:
            db.rollback()
            logging.exception('verify_approve atomic block failed for %s/%s',
                              session_code, plant_id)
            flash('Approve failed — transaction rolled back. Try again or contact admin.',
                  'danger')
            return redirect(url_for('verify_trail', session_code=session_code,
                                     plant_id=plant_id))

        # Audit log AFTER the atomic commit (log_record_change commits its own
        # tx for the hash chain; inside the try it would prematurely flush a
        # partial write before seeding completes).
        log_record_change('VERIFY_APPROVE', cal['id'], 'calendar',
                          before=before_snap_dict,
                          after=dict(after_snap) if after_snap else None,
                          extra_detail=f'note:{note[:200]}')

        msg = f'Session {session_code} verified — now counted as Conducted.'
        if seeded:
            msg += (f' {seeded} effectiveness review(s) seeded — '
                    f'manager input due by {due_date}.')
        flash(msg, 'success')
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
        db.execute('DELETE FROM effectiveness_review WHERE session_code=? AND plant_id=?',
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
        return render_template(
            'verify_trail.html', trail=trail, cal=cal, attendees=attendees,
            checklist_items=VERIFY_CHECKLIST_ITEMS,
            verify_note_min_len=VERIFY_NOTE_MIN_LEN,
        )
