"""3-month post-training effectiveness review tracking.

SOP: 25% of yearly TNI programmes are tagged Specialized. For these, every
attendee needs a manager-led effectiveness review 90 days after the session
was Conducted. SPOC collects manager feedback offline + files into TMS.

Status states (derived, not stored):
    pending    — completed_date IS NULL AND today < due_date
    due        — completed_date IS NULL AND today >= due_date AND today <= due_date+30
    overdue    — completed_date IS NULL AND today > due_date + 30
    completed  — completed_date IS NOT NULL

Auto-seeded on verify_approve when category='Specialized'.
"""
from datetime import date as _d, timedelta as _td

from flask import render_template, request, redirect, url_for, session, flash, jsonify

from tms.db import get_db
from tms.decorators import spoc_required, central_required, login_required
from tms.audit import log_record_change
from tms.helpers import _today_ist, _now_ist


OVERDUE_DAYS = 30  # days past due_date before classifying as Overdue


def _eff_status(today_iso, due_date, completed_date):
    """Map raw fields → display status. Pure function for testability."""
    if completed_date:
        return 'completed'
    if not due_date:
        return 'pending'
    if today_iso < due_date:
        return 'pending'
    try:
        overdue_cutoff = (_d.fromisoformat(due_date) +
                          _td(days=OVERDUE_DAYS)).isoformat()
        if today_iso > overdue_cutoff:
            return 'overdue'
    except (ValueError, TypeError):
        pass
    return 'due'


def _eff_counts(db, plant_id=None):
    """Pending+due+overdue+completed counts. plant_id=None → all plants
    (central/admin view)."""
    today = _today_ist().isoformat()
    where = ''
    params = []
    if plant_id is not None:
        where = ' WHERE plant_id=?'
        params = [plant_id]
    rows = db.execute(
        f'SELECT completed_date, due_date FROM effectiveness_review{where}',
        params).fetchall()
    counts = {'pending': 0, 'due': 0, 'overdue': 0, 'completed': 0}
    for r in rows:
        counts[_eff_status(today, r['due_date'], r['completed_date'])] += 1
    counts['total'] = sum(counts.values())
    counts['open']  = counts['pending'] + counts['due'] + counts['overdue']
    return counts


def _register(app):

    @app.route('/effectiveness')
    @login_required
    def effectiveness():
        """SPOC view: own-plant effectiveness reviews to file.
        Central/admin: cross-plant view (filterable)."""
        db = get_db()
        role = session.get('role')
        plant_id = session.get('plant_id')

        today = _today_ist().isoformat()
        sel_status = request.args.get('status', '').strip()

        params = []
        where = '1=1'
        if role == 'spoc':
            where += ' AND e.plant_id=?'
            params.append(plant_id)
        # central/admin see all plants by default
        rows = db.execute(f'''
            SELECT e.id, e.plant_id, e.session_code, e.emp_code,
                   e.conducted_date, e.due_date, e.completed_date,
                   e.rating, e.behaviour_change, e.application_on_job,
                   e.comments, e.filed_at,
                   emp.name AS emp_name, emp.designation, emp.department, emp.collar,
                   c.programme_name, c.prog_type,
                   p.name AS plant_name, p.unit_code,
                   u.username AS filed_by_name
            FROM effectiveness_review e
            LEFT JOIN employees  emp ON emp.plant_id=e.plant_id AND emp.emp_code=e.emp_code
            LEFT JOIN calendar   c   ON c.plant_id=e.plant_id   AND c.session_code=e.session_code
            LEFT JOIN plants     p   ON p.id=e.plant_id
            LEFT JOIN users      u   ON u.id=e.filed_by
            WHERE {where}
            ORDER BY
                CASE WHEN e.completed_date IS NULL THEN 0 ELSE 1 END,
                e.due_date ASC
        ''', params).fetchall()

        # Compute status per row in Python (cheap, keeps SQL simple)
        decorated = []
        for r in rows:
            st = _eff_status(today, r['due_date'], r['completed_date'])
            if sel_status and st != sel_status:
                continue
            d = dict(r)
            d['status_'] = st
            decorated.append(d)

        counts = _eff_counts(db, plant_id if role == 'spoc' else None)
        return render_template('effectiveness.html',
                               rows=decorated, counts=counts,
                               sel_status=sel_status, today=today,
                               is_central=(role in ('central', 'admin')))

    @app.route('/effectiveness/<int:eff_id>/file', methods=['POST'])
    @spoc_required
    def effectiveness_file(eff_id):
        """SPOC files manager's review for one attendee."""
        db = get_db()
        role = session.get('role')
        plant_id = session.get('plant_id')

        eff = db.execute('SELECT * FROM effectiveness_review WHERE id=?',
                         (eff_id,)).fetchone()
        if not eff:
            flash('Effectiveness review not found.', 'danger')
            return redirect(url_for('effectiveness'))
        # Plant scope: SPOC and admin (acting as a plant) must match the
        # session plant_id. Cross-plant filing is not permitted from this
        # route — admin must switch plant first.
        if role in ('spoc', 'admin') and eff['plant_id'] != plant_id:
            flash('Not your plant. Switch plant to file this review.', 'danger')
            return redirect(url_for('effectiveness'))

        try:
            rating = int(request.form.get('rating', 0))
        except (ValueError, TypeError):
            rating = 0
        if rating < 1 or rating > 5:
            flash('Rating must be 1-5.', 'danger')
            return redirect(url_for('effectiveness'))

        behaviour = request.form.get('behaviour_change', '').strip()[:1000]
        application = request.form.get('application_on_job', '').strip()[:1000]
        comments = request.form.get('comments', '').strip()[:1000]

        # Min length on at least one observation field — prevents single-word entries
        if len(behaviour) < 10 and len(application) < 10:
            flash('Provide at least one observation (Behaviour Change OR Application on Job, min 10 chars).', 'danger')
            return redirect(url_for('effectiveness'))

        before = dict(eff)
        now_iso = _now_ist().isoformat(timespec='seconds')
        db.execute('''UPDATE effectiveness_review SET
            completed_date=?, rating=?, behaviour_change=?,
            application_on_job=?, comments=?,
            filed_by=?, filed_at=datetime('now','localtime')
            WHERE id=?''',
            (now_iso, rating, behaviour, application, comments,
             session.get('user_id'), eff_id))
        db.commit()
        after = db.execute('SELECT * FROM effectiveness_review WHERE id=?',
                           (eff_id,)).fetchone()
        log_record_change('EFFECTIVENESS_FILE', eff_id, 'effectiveness_review',
                          before=before, after=dict(after))
        flash(f'Review filed for {eff["emp_code"]} — rating {rating}/5.', 'success')
        return redirect(url_for('effectiveness'))

    @app.route('/api/effectiveness/counts')
    @login_required
    def api_effectiveness_counts():
        """Real-time pill counts for sidebar / dashboard. JSON. Cheap query."""
        db = get_db()
        role = session.get('role')
        pid  = session.get('plant_id') if role == 'spoc' else None
        return jsonify(_eff_counts(db, pid))
