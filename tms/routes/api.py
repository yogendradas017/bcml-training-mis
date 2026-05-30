from flask import request, session, jsonify, render_template

from tms.constants import PROG_TYPES
from tms.db import get_db
from tms.decorators import login_required, spoc_required
import datetime
from tms.helpers import _fy_label, _derive_audience, _calc_compliance, _current_fy, _now_ist


def _register(app):

    @app.route('/api/employee/<emp_code>')
    @login_required
    def api_employee(emp_code):
        plant_id = session.get('plant_id')
        if not plant_id:
            return jsonify({})
        db  = get_db()
        emp = db.execute('SELECT * FROM employees WHERE emp_code=? AND plant_id=? AND is_active=1',
                         (emp_code.strip(), plant_id)).fetchone()
        if not emp:
            return jsonify({})
        return jsonify({
            'name': emp['name'], 'designation': emp['designation'],
            'grade': emp['grade'], 'collar': emp['collar'],
            'department': emp['department'], 'section': emp['section'],
            'category': emp['category'], 'gender': emp['gender']
        })

    def _session_lookup(session_code, plant_id):
        db = get_db()
        cal = db.execute('SELECT * FROM calendar WHERE session_code=? AND plant_id=?',
                         (session_code.strip(), plant_id)).fetchone()
        if not cal:
            return jsonify({})
        return jsonify({
            'programme_name': cal['programme_name'], 'prog_type': cal['prog_type'],
            'level': cal['level'], 'mode': cal['mode'],
            'plan_start': cal['plan_start'], 'plan_end': cal['plan_end'],
            'duration_hrs': cal['duration_hrs'], 'target_audience': cal['target_audience']
        })

    @app.route('/api/session-info')
    @login_required
    def api_session_info():
        plant_id = session.get('plant_id')
        if not plant_id:
            return jsonify({})
        session_code = request.args.get('code', '').strip()
        if not session_code:
            return jsonify({})
        return _session_lookup(session_code, plant_id)

    @app.route('/api/session/<path:session_code>')
    @login_required
    def api_session(session_code):
        plant_id = session.get('plant_id')
        if not plant_id:
            return jsonify({})
        return _session_lookup(session_code, plant_id)

    @app.route('/api/employees_list')
    @login_required
    def api_employees_list():
        plant_id = session.get('plant_id')
        if not plant_id:
            return jsonify([])
        db   = get_db()
        emps = db.execute('SELECT emp_code, name FROM employees WHERE plant_id=? AND is_active=1 ORDER BY name',
                          (plant_id,)).fetchall()
        return jsonify([{'code': e['emp_code'], 'name': e['name']} for e in emps])

    @app.route('/api/emp-lookup')
    @spoc_required
    def api_emp_lookup():
        plant_id = session['plant_id']
        code = request.args.get('code', '').strip()
        if not code:
            return jsonify({'name': None})
        db  = get_db()
        emp = db.execute('SELECT name FROM employees WHERE emp_code=? AND plant_id=? AND is_active=1',
                         (code, plant_id)).fetchone()
        return jsonify({'name': emp['name'] if emp else None})

    @app.route('/api/programme-list')
    @spoc_required
    def api_programme_list():
        plant_id = session['plant_id']
        db = get_db()
        master = [r[0] for r in db.execute(
            'SELECT name FROM programme_master WHERE plant_id=? ORDER BY name', (plant_id,)
        ).fetchall()] or []
        return jsonify(master)

    @app.route('/api/tni-coverage')
    @spoc_required
    def api_tni_coverage():
        from difflib import get_close_matches as gcm
        plant_id  = session['plant_id']
        prog_name = request.args.get('q', '').strip()
        if not prog_name:
            return jsonify({})
        db  = get_db()
        fy  = _fy_label()
        fy_start, fy_end = _current_fy()

        canonical = prog_name
        exact = db.execute('SELECT 1 FROM tni WHERE plant_id=? AND programme_name=? LIMIT 1',
                           (plant_id, prog_name)).fetchone()
        if not exact:
            all_names = [r[0] for r in db.execute(
                'SELECT DISTINCT programme_name FROM tni WHERE plant_id=?', (plant_id,))]
            m = gcm(prog_name.lower(), [n.lower() for n in all_names], n=1, cutoff=0.65)
            if m:
                canonical = all_names[[n.lower() for n in all_names].index(m[0])]

        demand           = db.execute('SELECT COUNT(DISTINCT emp_code) FROM tni WHERE plant_id=? AND programme_name=?',
                                      (plant_id, canonical)).fetchone()[0]
        sessions_planned = db.execute('SELECT COUNT(*) FROM calendar WHERE plant_id=? AND programme_name=? AND session_code LIKE ?',
                                      (plant_id, canonical, f'%/{fy}/%')).fetchone()[0]
        covered          = db.execute(
            'SELECT COUNT(DISTINCT emp_code) FROM emp_training WHERE plant_id=? AND programme_name=? AND start_date BETWEEN ? AND ?',
            (plant_id, canonical, fy_start, fy_end)).fetchone()[0]
        uncovered = max(0, demand - covered)
        pct       = round(covered / demand * 100) if demand > 0 else 0

        meta = db.execute('''
            SELECT t.prog_type, t.mode, e.collar, COUNT(*) as cnt
            FROM tni t
            LEFT JOIN employees e ON e.emp_code=t.emp_code AND e.plant_id=t.plant_id
            WHERE t.plant_id=? AND t.programme_name=?
            GROUP BY t.prog_type, t.mode, e.collar
            ORDER BY cnt DESC LIMIT 1
        ''', (plant_id, canonical)).fetchone()

        month_row = db.execute('''
            SELECT target_month, COUNT(*) as cnt FROM tni
            WHERE plant_id=? AND programme_name=? AND target_month IS NOT NULL AND target_month != ''
            GROUP BY target_month ORDER BY cnt DESC LIMIT 1
        ''', (plant_id, canonical)).fetchone()

        hrs_row = db.execute('''
            SELECT AVG(planned_hours) as avg_hrs FROM tni
            WHERE plant_id=? AND programme_name=? AND planned_hours > 0
        ''', (plant_id, canonical)).fetchone()

        if not meta or not meta['prog_type']:
            pm = db.execute('SELECT prog_type, mode FROM programme_master WHERE plant_id=? AND LOWER(name)=LOWER(?)',
                            (plant_id, canonical)).fetchone()
        else:
            pm = None

        prog_type = (meta['prog_type'] if meta else '') or (pm['prog_type'] if pm else '')
        mode      = (meta['mode']      if meta else '') or (pm['mode']      if pm else '')

        audience   = _derive_audience(plant_id, canonical, db) or ''
        tni_locked = bool(audience)
        source     = 'TNI Driven' if demand > 0 else ''
        avg_hrs    = round(hrs_row['avg_hrs'], 1) if hrs_row and hrs_row['avg_hrs'] else 0

        return jsonify({
            'demand': demand, 'sessions_planned': sessions_planned,
            'covered': covered, 'uncovered': uncovered, 'pct': pct,
            'prog_type':    prog_type,
            'mode':         mode,
            'audience':     audience,
            'tni_locked':   tni_locked,
            'source':       source,
            'target_month': month_row['target_month'] if month_row else '',
            'avg_hrs':      avg_hrs,
        })

    @app.route('/api/dashboard-monthly')
    @spoc_required
    def api_dashboard_monthly():
        db  = get_db()
        pid = session['plant_id']

        # Per-month training attendance stats
        raw = db.execute('''
            SELECT strftime('%m', start_date) as mo,
                   COUNT(*) as seats,
                   COALESCE(SUM(hrs), 0) as hrs
            FROM emp_training
            WHERE plant_id=? AND start_date IS NOT NULL AND start_date!=''
            GROUP BY mo
        ''', (pid,)).fetchall()
        mo_map = {r['mo']: {'seats': r['seats'], 'hrs': round(r['hrs'], 1)} for r in raw}

        # Sessions conducted per month from programme_details
        sess_raw = db.execute('''
            SELECT strftime('%m', pd.start_date) as mo, COUNT(*) as n
            FROM programme_details pd
            JOIN calendar c ON pd.session_code = c.session_code
            WHERE c.plant_id=? AND pd.start_date IS NOT NULL AND pd.start_date!=''
            GROUP BY mo
        ''', (pid,)).fetchall()
        sess_map = {r['mo']: r['n'] for r in sess_raw}

        plan_raw = db.execute('''
            SELECT strftime('%m', plan_start) as mo, COUNT(*) as n
            FROM calendar WHERE plant_id=? AND plan_start IS NOT NULL AND plan_start!=''
            GROUP BY mo
        ''', (pid,)).fetchall()
        plan_map = {r['mo']: r['n'] for r in plan_raw}

        FY = [('04','April'),('05','May'),('06','June'),('07','July'),
              ('08','August'),('09','September'),('10','October'),('11','November'),
              ('12','December'),('01','January'),('02','February'),('03','March')]

        monthly = []
        for mo_num, mo_name in FY:
            d = mo_map.get(mo_num, {'seats': 0, 'hrs': 0.0})
            monthly.append({'month': mo_name, 'mo': mo_num,
                            'seats': d['seats'], 'hrs': d['hrs'],
                            'sessions': sess_map.get(mo_num, 0),
                            'planned': plan_map.get(mo_num, 0)})

        comp = _calc_compliance(pid, db)
        cur_mo = _now_ist().strftime('%m')

        bc_target    = comp.get('bc_mandate', 0)
        wc_target    = comp.get('wc_mandate', 0)
        bc_actual    = comp.get('bc_actual',  0)
        wc_actual    = comp.get('wc_actual',  0)
        total_actual = round(bc_actual + wc_actual, 1)
        total_target = bc_target + wc_target
        total_mh_pct = round(total_actual / total_target * 100, 1) if total_target else 0

        return jsonify({
            'monthly': monthly,
            'compliance': {
                'bc_pct':      round(comp.get('bc_pct',    0), 1),
                'wc_pct':      round(comp.get('wc_pct',    0), 1),
                'total_pct':   round(comp.get('total_pct', 0), 1),
                'bc_mh_pct':   round(comp.get('bc_pct',    0), 1),
                'wc_mh_pct':   round(comp.get('wc_pct',    0), 1),
                'total_mh_pct': total_mh_pct,
                'bc_actual':   bc_actual,   'bc_target':   bc_target,
                'wc_actual':   wc_actual,   'wc_target':   wc_target,
                'total_actual': total_actual, 'total_target': total_target,
            },
            'cur_mo': cur_mo
        })

    @app.route('/api/manhour-drilldown')
    @spoc_required
    def api_manhour_drilldown():
        pid    = session['plant_id']
        collar = request.args.get('collar', 'ALL').upper()
        db     = get_db()
        fy_start, fy_end = _current_fy()

        if collar == 'BC':
            collar_where = "AND e.collar='Blue Collared'"
        elif collar == 'WC':
            collar_where = "AND e.collar='White Collared'"
        else:
            collar_where = ''

        dept_rows = db.execute(f'''
            SELECT e.department, e.collar,
                   COUNT(DISTINCT e.emp_code)   AS emp_count,
                   COALESCE(SUM(et.hrs), 0)     AS actual_hrs
            FROM employees e
            LEFT JOIN emp_training et
                   ON et.emp_code=e.emp_code AND et.plant_id=e.plant_id
                  AND et.start_date BETWEEN ? AND ?
            WHERE e.plant_id=? AND e.is_active=1 {collar_where}
              AND e.department IS NOT NULL AND e.department!=''
            GROUP BY e.department, e.collar
            ORDER BY e.department
        ''', (fy_start, fy_end, pid)).fetchall()

        dept_map = {}
        for r in dept_rows:
            dept = r['department']
            if dept not in dept_map:
                dept_map[dept] = {'dept': dept, 'emp_count': 0, 'actual_hrs': 0.0, 'target_hrs': 0.0}
            d = dept_map[dept]
            d['emp_count']  += r['emp_count']
            d['actual_hrs'] += r['actual_hrs']
            d['target_hrs'] += r['emp_count'] * (12 if r['collar'] == 'Blue Collared' else 24 if r['collar'] == 'White Collared' else 0)

        departments = []
        for d in sorted(dept_map.values(), key=lambda x: x['dept']):
            pct = round(d['actual_hrs'] / d['target_hrs'] * 100, 1) if d['target_hrs'] else 0
            departments.append({'dept': d['dept'], 'emp_count': d['emp_count'],
                                'actual_hrs': round(d['actual_hrs'], 1),
                                'target_hrs': round(d['target_hrs'], 1), 'pct': pct})

        emp_rows = db.execute(f'''
            SELECT e.emp_code, e.name, e.department, e.collar,
                   COALESCE(SUM(et.hrs), 0) AS actual_hrs
            FROM employees e
            LEFT JOIN emp_training et
                   ON et.emp_code=e.emp_code AND et.plant_id=e.plant_id
                  AND et.start_date BETWEEN ? AND ?
            WHERE e.plant_id=? AND e.is_active=1 {collar_where}
            GROUP BY e.emp_code, e.name, e.department, e.collar
            ORDER BY e.department, e.name
        ''', (fy_start, fy_end, pid)).fetchall()

        employees = []
        for r in emp_rows:
            tgt    = 12 if r['collar'] == 'Blue Collared' else 24 if r['collar'] == 'White Collared' else 0
            actual = round(r['actual_hrs'], 1)
            pct    = round(actual / tgt * 100, 1) if tgt else 0
            status = 'Met' if pct >= 100 else ('Zero' if actual == 0 else 'Low')
            employees.append({'emp_code': r['emp_code'], 'name': r['name'],
                              'department': r['department'] or '', 'collar': r['collar'] or '',
                              'actual_hrs': actual, 'target_hrs': tgt, 'pct': pct, 'status': status})

        return jsonify({'departments': departments, 'employees': employees})

    @app.route('/intelligence')
    @spoc_required
    def programme_intelligence():
        plant_id = session['plant_id']
        db       = get_db()
        fy       = _fy_label()

        unique_progs = db.execute(
            'SELECT COUNT(DISTINCT programme_name) FROM tni WHERE plant_id=?',
            (plant_id,)).fetchone()[0]

        tni_rows = db.execute(
            'SELECT programme_name, COUNT(DISTINCT emp_code) as demand FROM tni WHERE plant_id=? GROUP BY programme_name ORDER BY demand DESC',
            (plant_id,)).fetchall()

        fy_start, fy_end = _current_fy()
        covered_map = {}
        for r in db.execute(
                'SELECT programme_name, COUNT(DISTINCT emp_code) as cnt FROM emp_training'
                ' WHERE plant_id=? AND start_date BETWEEN ? AND ? GROUP BY programme_name',
                (plant_id, fy_start, fy_end)):
            covered_map[r['programme_name']] = r['cnt']

        sessions_map = {}
        for r in db.execute('SELECT programme_name, COUNT(*) as cnt FROM calendar WHERE plant_id=? AND session_code LIKE ? GROUP BY programme_name',
                            (plant_id, f'%/{fy}/%')):
            sessions_map[r['programme_name']] = r['cnt']

        programmes = []
        total_needs = total_covered = 0
        for r in tni_rows:
            name      = r['programme_name']
            demand    = r['demand']
            covered   = covered_map.get(name, 0)
            planned   = sessions_map.get(name, 0)
            uncovered = max(0, demand - covered)
            pct       = round(covered / demand * 100) if demand > 0 else 0
            if demand < 30:  status = 'Small Group'
            elif pct >= 80:  status = 'On Track'
            elif pct >= 30:  status = 'In Progress'
            else:            status = 'Big Ticket'
            total_needs   += demand
            total_covered += min(covered, demand)
            programmes.append({'name': name, 'demand': demand, 'covered': covered,
                               'planned': planned, 'uncovered': uncovered, 'pct': pct, 'status': status})

        total_uncovered = max(0, total_needs - total_covered)
        progs_in_demand = sum(1 for p in programmes if p['uncovered'] > 0)
        big_tickets     = sum(1 for p in programmes if p['status'] == 'Big Ticket')

        _status_order = {'Big Ticket': 0, 'In Progress': 1, 'On Track': 2, 'Small Group': 3}
        programmes.sort(key=lambda p: (_status_order.get(p['status'], 9), -p['demand']))

        _mode_keys = ['Classroom', 'OJT', 'Online', 'SOP']
        mode_map   = {m: {'planned': 0, 'conducted': 0} for m in _mode_keys}
        mode_map['Other'] = {'planned': 0, 'conducted': 0}
        for r in db.execute(
                'SELECT mode, status, COUNT(*) as cnt FROM calendar WHERE plant_id=? AND session_code LIKE ? GROUP BY mode, status',
                (plant_id, f'%/{fy}/%')):
            key = r['mode'] if r['mode'] in _mode_keys else 'Other'
            mode_map[key]['planned'] += r['cnt']
            if r['status'] == 'Conducted':
                mode_map[key]['conducted'] += r['cnt']
        session_modes = [{'mode': k, **v, 'pct': round(v['conducted']/v['planned']*100) if v['planned'] else 0}
                         for k, v in mode_map.items() if v['planned'] > 0]
        total_sessions = sum(v['planned'] for v in mode_map.values())

        return render_template('intelligence.html', fy=fy,
                               unique_progs=unique_progs, progs_in_demand=progs_in_demand,
                               total_needs=total_needs, total_covered=total_covered,
                               total_uncovered=total_uncovered,
                               total_sessions=total_sessions, session_modes=session_modes,
                               big_tickets=big_tickets, programmes=programmes)
