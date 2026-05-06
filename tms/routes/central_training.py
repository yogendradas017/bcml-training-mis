from flask import render_template, request, redirect, url_for, session, flash, jsonify

from tms.constants import (CENTRAL_PLANT_ID, PROG_TYPES, MODES, LEVELS,
                            AUDIENCES, MONTHS_FY, STATUSES, TYPE_ABBREV)
from tms.db import get_db
from tms.decorators import central_required
from tms.helpers import (
    _get_or_create_prog_code, _new_session_code,
    _canonical_prog, _current_fy, _in_current_fy,
)


def _register(app):

    # ── Corp Members ─────────────────────────────────────────────────────────

    @app.route('/central/corp-members')
    @central_required
    def central_corp_members():
        db = get_db()
        members = db.execute(
            'SELECT * FROM corp_members ORDER BY name'
        ).fetchall()
        return render_template('central_corp_members.html', members=members)

    @app.route('/central/corp-members/add', methods=['POST'])
    @central_required
    def central_corp_member_add():
        f = request.form
        emp_code = f.get('emp_code', '').strip().upper()
        name = f.get('name', '').strip()
        if not emp_code or not name:
            flash('Employee code and name are required.', 'danger')
            return redirect(url_for('central_corp_members'))
        db = get_db()
        try:
            db.execute(
                '''INSERT INTO corp_members(emp_code, name, designation, department, email)
                   VALUES(?,?,?,?,?)''',
                (emp_code, name,
                 f.get('designation', '').strip(),
                 f.get('department', '').strip(),
                 f.get('email', '').strip())
            )
            db.commit()
            flash(f'Corp member {name} added.', 'success')
        except Exception:
            flash(f'Employee code {emp_code} already exists.', 'danger')
        return redirect(url_for('central_corp_members'))

    @app.route('/central/corp-members/<int:member_id>/edit', methods=['POST'])
    @central_required
    def central_corp_member_edit(member_id):
        f = request.form
        name = f.get('name', '').strip()
        if not name:
            flash('Name is required.', 'danger')
            return redirect(url_for('central_corp_members'))
        db = get_db()
        db.execute(
            '''UPDATE corp_members SET name=?, designation=?, department=?, email=?,
               is_active=? WHERE id=?''',
            (name, f.get('designation', '').strip(), f.get('department', '').strip(),
             f.get('email', '').strip(),
             1 if f.get('is_active') else 0, member_id)
        )
        db.commit()
        flash('Corp member updated.', 'success')
        return redirect(url_for('central_corp_members'))

    @app.route('/central/corp-members/<int:member_id>/delete', methods=['POST'])
    @central_required
    def central_corp_member_delete(member_id):
        db = get_db()
        db.execute('DELETE FROM corp_members WHERE id=?', (member_id,))
        db.commit()
        flash('Corp member removed.', 'warning')
        return redirect(url_for('central_corp_members'))

    # ── Central Programme Master ──────────────────────────────────────────────

    @app.route('/central/programmes')
    @central_required
    def central_programmes():
        db = get_db()
        programmes = db.execute(
            'SELECT * FROM programme_master WHERE plant_id=? ORDER BY name',
            (CENTRAL_PLANT_ID,)
        ).fetchall()
        return render_template('central_programmes.html', programmes=programmes,
                               prog_types=PROG_TYPES)

    @app.route('/central/programmes/add', methods=['POST'])
    @central_required
    def central_programme_add():
        f = request.form
        name = f.get('name', '').strip()
        if not name:
            flash('Programme name is required.', 'danger')
            return redirect(url_for('central_programmes'))
        db = get_db()
        prog_type = f.get('prog_type', '')
        try:
            db.execute(
                '''INSERT INTO programme_master(plant_id, name, prog_type, source)
                   VALUES(?,?,?,?)''',
                (CENTRAL_PLANT_ID, name, prog_type, 'New Requirement')
            )
            db.commit()
            flash(f'Programme "{name}" added.', 'success')
        except Exception:
            flash(f'Programme "{name}" already exists.', 'danger')
        return redirect(url_for('central_programmes'))

    @app.route('/central/programmes/<int:prog_id>/delete', methods=['POST'])
    @central_required
    def central_programme_delete(prog_id):
        db = get_db()
        db.execute('DELETE FROM programme_master WHERE id=? AND plant_id=?',
                   (prog_id, CENTRAL_PLANT_ID))
        db.commit()
        flash('Programme removed.', 'warning')
        return redirect(url_for('central_programmes'))

    # ── Central Calendar ──────────────────────────────────────────────────────

    @app.route('/central/calendar')
    @central_required
    def central_calendar():
        db = get_db()
        sessions = [dict(s) for s in db.execute(
            'SELECT * FROM calendar WHERE plant_id=? ORDER BY id DESC',
            (CENTRAL_PLANT_ID,)
        ).fetchall()]
        master_programmes = [r[0] for r in db.execute(
            'SELECT name FROM programme_master WHERE plant_id=? ORDER BY name',
            (CENTRAL_PLANT_ID,)
        ).fetchall()]

        qr_rows = db.execute(
            'SELECT session_code, stage, token, is_active FROM session_qr WHERE plant_id=?',
            (CENTRAL_PLANT_ID,)
        ).fetchall()
        qr_map = {}
        for q in qr_rows:
            qr_map.setdefault(q['session_code'], {})[q['stage']] = dict(q)

        fb_counts = {r['session_code']: r['cnt'] for r in db.execute(
            'SELECT session_code, COUNT(*) as cnt FROM feedback_response WHERE plant_id=? GROUP BY session_code',
            (CENTRAL_PLANT_ID,)
        ).fetchall()}

        return render_template('central_calendar.html',
                               sessions=sessions,
                               master_programmes=master_programmes,
                               prog_types=PROG_TYPES, modes=MODES,
                               levels=LEVELS, audiences=AUDIENCES,
                               months=MONTHS_FY, statuses=STATUSES,
                               qr_map=qr_map,
                               fb_counts=fb_counts)

    @app.route('/central/calendar/add', methods=['POST'])
    @central_required
    def central_calendar_add():
        f = request.form
        db = get_db()
        prog_name = _canonical_prog(f['programme_name'].strip(), CENTRAL_PLANT_ID, db)
        prog_type = f.get('prog_type', '')
        dur = float(f.get('duration_hrs') or 0)
        if dur <= 0:
            flash('Duration must be greater than 0 hours.', 'danger')
            return redirect(url_for('central_calendar'))

        fy_start, fy_end = _current_fy()
        for fld, lbl in [('plan_start', 'Plan Start'), ('plan_end', 'Plan End')]:
            val = f.get(fld, '')
            if val and not _in_current_fy(val):
                flash(f'{lbl} date must be within the current FY ({fy_start} to {fy_end}).', 'danger')
                return redirect(url_for('central_calendar'))

        prog_code    = _get_or_create_prog_code(CENTRAL_PLANT_ID, prog_name, prog_type, db)
        session_code = _new_session_code(CENTRAL_PLANT_ID, prog_code, db)

        db.execute('''INSERT INTO calendar
            (plant_id, prog_code, session_code, source, programme_name, prog_type,
             planned_month, plan_start, plan_end, time_from, time_to, duration_hrs,
             level, mode, target_audience, planned_pax, trainer_vendor, status, is_central)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
            (CENTRAL_PLANT_ID, prog_code, session_code, 'New Requirement',
             prog_name, prog_type,
             f.get('planned_month', ''), f.get('plan_start', ''), f.get('plan_end', ''),
             f.get('time_from', ''), f.get('time_to', ''), dur,
             f.get('level', ''), f.get('mode', ''), f.get('target_audience', ''),
             int(f.get('planned_pax') or 0), f.get('trainer_vendor', ''),
             'To Be Planned', 1))
        db.commit()
        flash(f'Central session {session_code} added.', 'success')
        return redirect(url_for('central_calendar'))

    @app.route('/central/calendar/<int:cal_id>/edit', methods=['POST'])
    @central_required
    def central_calendar_edit(cal_id):
        db = get_db()
        existing = db.execute('SELECT status FROM calendar WHERE id=? AND plant_id=?',
                              (cal_id, CENTRAL_PLANT_ID)).fetchone()
        if not existing:
            flash('Session not found.', 'danger')
            return redirect(url_for('central_calendar'))
        if existing['status'] == 'Conducted':
            flash('Conducted sessions cannot be edited.', 'danger')
            return redirect(url_for('central_calendar'))
        f = request.form
        dur = float(f.get('duration_hrs') or 0)
        if dur <= 0:
            flash('Duration must be greater than 0 hours.', 'danger')
            return redirect(url_for('central_calendar'))
        db.execute('''UPDATE calendar SET
            programme_name=?, prog_type=?, planned_month=?,
            plan_start=?, plan_end=?, time_from=?, time_to=?,
            duration_hrs=?, level=?, mode=?, target_audience=?,
            planned_pax=?, trainer_vendor=?, status=?
            WHERE id=? AND plant_id=?''',
            (f.get('programme_name', ''), f.get('prog_type', ''),
             f.get('planned_month', ''), f.get('plan_start', ''), f.get('plan_end', ''),
             f.get('time_from', ''), f.get('time_to', ''), dur,
             f.get('level', ''), f.get('mode', ''), f.get('target_audience', ''),
             int(f.get('planned_pax') or 0), f.get('trainer_vendor', ''),
             f.get('status', 'To Be Planned'),
             cal_id, CENTRAL_PLANT_ID))
        db.commit()
        flash('Session updated.', 'success')
        return redirect(url_for('central_calendar'))

    @app.route('/central/calendar/<int:cal_id>/delete', methods=['POST'])
    @central_required
    def central_calendar_delete(cal_id):
        db = get_db()
        cal = db.execute('SELECT session_code, status FROM calendar WHERE id=? AND plant_id=?',
                         (cal_id, CENTRAL_PLANT_ID)).fetchone()
        if not cal:
            flash('Session not found.', 'danger')
            return redirect(url_for('central_calendar'))
        db.execute('DELETE FROM session_qr WHERE plant_id=? AND session_code=?',
                   (CENTRAL_PLANT_ID, cal['session_code']))
        db.execute('DELETE FROM calendar WHERE id=? AND plant_id=?',
                   (cal_id, CENTRAL_PLANT_ID))
        db.commit()
        flash('Session deleted.', 'warning')
        return redirect(url_for('central_calendar'))

    # ── Central Live Monitor (delegates to qr_live) ──────────────────────────
    # The existing /calendar/<id>/live route works for plant_id=99 via
    # spoc_or_central_required decorator. Central users access it directly.

    # ── Central Attendance Summary ────────────────────────────────────────────

    @app.route('/central/attendance')
    @central_required
    def central_attendance():
        db = get_db()
        from tms.constants import PLANT_MAP, PLANTS
        rows = db.execute('''
            SELECT t.emp_code, t.session_code, t.programme_name, t.start_date,
                   t.end_date, t.hrs, t.plant_id,
                   COALESCE(e.name, cm.name) AS emp_name,
                   COALESCE(e.designation, cm.designation) AS designation,
                   COALESCE(e.department, cm.department) AS department,
                   p.name AS plant_name, p.unit_code
            FROM emp_training t
            LEFT JOIN employees e ON e.emp_code=t.emp_code AND e.plant_id=t.plant_id
            LEFT JOIN corp_members cm ON cm.emp_code=t.emp_code AND t.plant_id=99
            LEFT JOIN plants p ON p.id=t.plant_id
            WHERE t.host_plant_id=99 OR t.plant_id=99
            ORDER BY t.session_code, t.emp_code
        ''').fetchall()

        # Group by session_code
        from collections import defaultdict
        session_groups = defaultdict(list)
        for r in rows:
            session_groups[r['session_code']].append(dict(r))

        # Get session details
        sessions = db.execute(
            'SELECT * FROM calendar WHERE plant_id=? ORDER BY plan_start DESC',
            (CENTRAL_PLANT_ID,)
        ).fetchall()
        session_map = {s['session_code']: dict(s) for s in sessions}

        return render_template('central_attendance.html',
                               session_groups=session_groups,
                               session_map=session_map,
                               plant_map=PLANT_MAP)

    # ── AJAX: programme search for central forms ──────────────────────────────

    @app.route('/central/prog-search')
    @central_required
    def central_prog_search():
        q = request.args.get('q', '').strip()
        if len(q) < 1:
            return jsonify([])
        db = get_db()
        rows = db.execute(
            'SELECT name, prog_type FROM programme_master WHERE plant_id=? '
            'AND LOWER(name) LIKE LOWER(?) ORDER BY name LIMIT 20',
            (CENTRAL_PLANT_ID, f'%{q}%')
        ).fetchall()
        return jsonify([dict(r) for r in rows])
