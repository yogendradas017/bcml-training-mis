from flask import render_template, request, session
from collections import defaultdict

from tms.db import get_db
from tms.decorators import spoc_required
from tms.config import get_config


def _register(app):

    @app.route('/reports/training-hours')
    @spoc_required
    def training_hours_report():
        plant_id   = session['plant_id']
        db         = get_db()
        # Per-plant man-hour targets from config (not hardcoded 12/24) so this
        # report agrees with the compliance gauge under per-plant overrides.
        bc_t = get_config('mh_target_bc', 12, plant_id=plant_id)
        wc_t = get_config('mh_target_wc', 24, plant_id=plant_id)

        dept_filter   = request.args.get('dept', '').strip()
        collar_filter = request.args.get('collar', '').strip()
        status_filter = request.args.get('status', '').strip()
        q             = request.args.get('q', '').strip().lower()

        rows = db.execute('''
            SELECT
                COALESCE(e.department, '') AS department,
                e.emp_code,
                e.name,
                COALESCE(e.designation, '') AS designation,
                COALESCE(e.collar, '') AS collar,
                COALESCE(SUM(t.hrs), 0) AS actual_hrs
            FROM employees e
            LEFT JOIN emp_training t
                ON t.emp_code=e.emp_code AND t.plant_id=e.plant_id
            WHERE e.plant_id=? AND e.is_active=1
            GROUP BY e.id, e.emp_code, e.name, e.designation, e.collar, e.department
            ORDER BY COALESCE(e.department,'') COLLATE NOCASE,
                     e.collar,
                     e.name COLLATE NOCASE
        ''', (plant_id,)).fetchall()

        # Enrich + compute status
        employees = []
        for r in rows:
            target  = bc_t if r['collar'] == 'Blue Collared' else wc_t
            actual  = round(r['actual_hrs'], 1)
            pct     = round(actual / target * 100, 1) if target else 0
            if actual == 0:
                status = 'zero'
            elif actual < target:
                status = 'low'
            else:
                status = 'ok'
            employees.append({
                'department':  r['department'] or '(No Department)',
                'emp_code':    r['emp_code'],
                'name':        r['name'],
                'designation': r['designation'],
                'collar':      r['collar'],
                'target_hrs':  target,
                'actual_hrs':  actual,
                'pct':         pct,
                'status':      status,
            })

        # Apply filters
        if dept_filter:
            employees = [e for e in employees if e['department'] == dept_filter]
        if collar_filter:
            employees = [e for e in employees if e['collar'] == collar_filter]
        if status_filter:
            employees = [e for e in employees if e['status'] == status_filter]
        if q:
            employees = [e for e in employees
                         if q in e['emp_code'].lower() or q in e['name'].lower()]

        # Summary chips (before filter for totals)
        all_emps = [{'collar': r['collar'],
                     'actual': round(r['actual_hrs'], 1),
                     'target': r['target_hrs']} for r in rows]
        total_bc   = sum(1 for e in all_emps if e['collar'] == 'Blue Collared')
        total_wc   = sum(1 for e in all_emps if e['collar'] == 'White Collared')
        zero_count = sum(1 for e in all_emps if e['actual'] == 0)
        low_count  = sum(1 for e in all_emps if 0 < e['actual'] < e['target'])
        ok_count   = sum(1 for e in all_emps if e['actual'] >= e['target'] and e['target'] > 0)

        # Group filtered employees by department
        dept_groups = defaultdict(list)
        for e in employees:
            dept_groups[e['department']].append(e)

        # Dept summaries
        departments = []
        for dept_name, emps in sorted(dept_groups.items()):
            bc_count  = sum(1 for e in emps if e['collar'] == 'Blue Collared')
            wc_count  = sum(1 for e in emps if e['collar'] == 'White Collared')
            zero_c    = sum(1 for e in emps if e['status'] == 'zero')
            total_act = sum(e['actual_hrs'] for e in emps)
            total_tgt = sum(e['target_hrs'] for e in emps)
            avg_pct   = round(total_act / total_tgt * 100, 1) if total_tgt else 0
            departments.append({
                'name':       dept_name,
                'employees':  emps,
                'bc_count':   bc_count,
                'wc_count':   wc_count,
                'zero_count': zero_c,
                'total_act':  round(total_act, 1),
                'total_tgt':  total_tgt,
                'avg_pct':    avg_pct,
            })

        # Dept list for filter dropdown
        all_depts = sorted({r['department'] or '(No Department)' for r in rows})

        return render_template('training_hours_report.html',
                               departments=departments,
                               all_depts=all_depts,
                               total_bc=total_bc,
                               total_wc=total_wc,
                               zero_count=zero_count,
                               low_count=low_count,
                               ok_count=ok_count,
                               dept_filter=dept_filter,
                               collar_filter=collar_filter,
                               status_filter=status_filter,
                               q=q)
