import io

from flask import redirect, url_for, session, flash, send_file, render_template, request

from tms.constants import PLANT_MAP, MONTHS_FY
from tms.db import get_db
from tms.decorators import login_required

import openpyxl
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Font, PatternFill, Alignment


def _register(app):

    @app.route('/export/<int:plant_id>', methods=['GET', 'POST'])
    @login_required
    def export_excel(plant_id):
        if session.get('role') == 'spoc' and session.get('plant_id') != plant_id:
            flash('Access denied.', 'danger')
            return redirect(url_for('spoc_dashboard'))
        plant = PLANT_MAP.get(plant_id)
        if not plant:
            flash('Plant not found.', 'danger')
            return redirect(url_for('index'))

        if request.method == 'GET':
            db   = get_db()
            depts = [r[0] for r in db.execute(
                'SELECT DISTINCT department FROM employees WHERE plant_id=? AND department IS NOT NULL AND department!="" ORDER BY department',
                (plant_id,)).fetchall()]
            return render_template('export_config.html', plant=plant,
                                   months=MONTHS_FY, departments=depts)

        # ── POST: build filtered workbook ──────────────────────────────────────
        f        = request.form
        sheets   = f.getlist('sheets')
        month_f  = f.get('month', '')
        collar_f = f.get('collar', '')
        dept_f   = f.get('dept', '')
        src_f    = f.get('source', '')

        if not sheets:
            flash('Select at least one sheet to export.', 'warning')
            return redirect(url_for('export_excel', plant_id=plant_id))

        db    = get_db()
        fy    = '2026-27'
        wb    = openpyxl.Workbook(write_only=True)
        pname = plant['name'].upper()

        H_FONT  = Font(bold=True, color='FFFFFF', size=10)
        H_FILL  = PatternFill('solid', fgColor='1F4E79')
        H_ALIGN = Alignment(horizontal='center', vertical='center', wrap_text=True)
        T_FONT  = Font(bold=True, size=11)

        def hc(ws, val):
            c = WriteOnlyCell(ws, value=val)
            c.font = H_FONT; c.fill = H_FILL; c.alignment = H_ALIGN
            return c

        def tc(ws, val):
            c = WriteOnlyCell(ws, value=val)
            c.font = T_FONT
            return c

        # ── Sheet 1: Employee Master ───────────────────────────────────────────
        if 'emp_master' in sheets:
            ws1 = wb.create_sheet('EMP_MASTER')
            ws1.append([tc(ws1, 'BALRAMPUR CHINI MILLS LIMITED')])
            ws1.append([tc(ws1, f'{pname} — EMPLOYEE MASTER | FY {fy}')])
            ws1.append([])
            ws1.append([hc(ws1, h) for h in
                ['Sr.','Emp Code','Name','Designation','Grade','Collar',
                 'Department','Section','Category','Gender','PH',
                 'Exit Date','Exit Reason','Remarks']])
            where, params = ['plant_id=?'], [plant_id]
            if collar_f: where.append('collar=?'); params.append(collar_f)
            if dept_f:   where.append('department=?'); params.append(dept_f)
            for r, e in enumerate(db.execute(
                    f'SELECT * FROM employees WHERE {" AND ".join(where)} ORDER BY name', params), 1):
                ws1.append([r, e['emp_code'], e['name'], e['designation'] or '',
                            e['grade'] or '', e['collar'] or '', e['department'] or '',
                            e['section'] or '', e['category'] or '', e['gender'] or '',
                            e['physically_handicapped'] or '',
                            e['exit_date'] or '', e['exit_reason'] or '', e['remarks'] or ''])

        # ── Sheet 2: TNI Tracking ──────────────────────────────────────────────
        if 'tni' in sheets:
            ws2 = wb.create_sheet('TNI_Tracking')
            ws2.append([tc(ws2, 'BALRAMPUR CHINI MILLS LIMITED')])
            ws2.append([tc(ws2, f'{pname} — TNI TRACKING | FY {fy}')])
            ws2.append([])
            ws2.append([hc(ws2, h) for h in
                ['Sr.','Emp Code','Name','Designation','Grade','Collar','Dept',
                 'Section','Programme Name','Type','Mode','Target Month','Planned Hrs','Completed?','Source']])
            done_set = set(
                (row[0], row[1]) for row in db.execute(
                    'SELECT emp_code, programme_name FROM emp_training WHERE plant_id=?', (plant_id,)))
            where  = ['t.plant_id=?']; params = [plant_id]
            if collar_f: where.append('e.collar=?');      params.append(collar_f)
            if dept_f:   where.append('e.department=?');  params.append(dept_f)
            if src_f:    where.append('t.source=?');      params.append(src_f)
            for r, t in enumerate(db.execute(f'''
                    SELECT t.*,e.name,e.designation,e.grade,e.collar,e.department,e.section
                    FROM tni t LEFT JOIN employees e
                      ON e.emp_code=t.emp_code AND e.plant_id=t.plant_id
                    WHERE {" AND ".join(where)}''', params), 1):
                ws2.append([r, t['emp_code'], t['name'] or '', t['designation'] or '',
                            t['grade'] or '', t['collar'] or '', t['department'] or '',
                            t['section'] or '', t['programme_name'], t['prog_type'] or '',
                            t['mode'] or '', t['target_month'] or '', t['planned_hours'],
                            'Yes' if (t['emp_code'], t['programme_name']) in done_set else 'No',
                            t['source'] or ''])

        # ── Sheet 3: Calendar ─────────────────────────────────────────────────
        if 'calendar' in sheets:
            ws3 = wb.create_sheet('Cal_Plan_vs_Actual')
            ws3.append([tc(ws3, 'BALRAMPUR CHINI MILLS LIMITED')])
            ws3.append([tc(ws3, f'{pname} — TRAINING CALENDAR | FY {fy}')])
            ws3.append([])
            ws3.append([hc(ws3, h) for h in
                ['S/N','PROG CODE','SESSION CODE','Source','Programme Name','Type',
                 'Planned Month','Plan Start','Plan End','Duration(Hrs)','Level','Mode',
                 'Target Audience','Planned Pax','Trainer/Vendor','STATUS','Actual Date','Actual Pax']])
            act_pax_map = {row[0]: row[1] for row in db.execute(
                'SELECT session_code, COUNT(*) FROM emp_training WHERE plant_id=? GROUP BY session_code', (plant_id,))}
            pd_date_map = {row[0]: row[1] for row in db.execute(
                'SELECT session_code, start_date FROM programme_details WHERE plant_id=?', (plant_id,))}
            where  = ['plant_id=?']; params = [plant_id]
            if month_f: where.append('planned_month=?');  params.append(month_f)
            if src_f:   where.append('source=?');         params.append(src_f)
            for r, c in enumerate(db.execute(
                    f'SELECT * FROM calendar WHERE {" AND ".join(where)} ORDER BY id', params), 1):
                ws3.append([r, c['prog_code'], c['session_code'], c['source'] or '',
                            c['programme_name'], c['prog_type'] or '', c['planned_month'] or '',
                            c['plan_start'] or '', c['plan_end'] or '',
                            c['duration_hrs'], c['level'] or '', c['mode'] or '',
                            c['target_audience'] or '', c['planned_pax'],
                            c['trainer_vendor'] or '', c['status'] or '',
                            pd_date_map.get(c['session_code'], ''),
                            act_pax_map.get(c['session_code'], 0)])

        # ── Sheet 4: 2A Employee Training ─────────────────────────────────────
        if 'training_2a' in sheets:
            ws4 = wb.create_sheet('2A_Emp_Training')
            ws4.append([tc(ws4, 'BALRAMPUR CHINI MILLS LIMITED')])
            ws4.append([tc(ws4, f'{pname} — 2A: EMPLOYEE TRAINING DETAILS | FY {fy}')])
            ws4.append([])
            ws4.append([hc(ws4, h) for h in
                ['Sr.','Emp Code','Name','Designation','Grade','Collar','Dept','Section',
                 'Start Date','End Date','Hrs','Programme Name','Type','Level','Mode',
                 'Cal/New','Pre Rating','Post Rating','Venue','Month']])
            where  = ['t.plant_id=?']; params = [plant_id]
            if month_f:  where.append('t.month=?');       params.append(month_f)
            if collar_f: where.append('e.collar=?');      params.append(collar_f)
            if dept_f:   where.append('e.department=?');  params.append(dept_f)
            for r, t in enumerate(db.execute(f'''
                    SELECT t.*,e.name as emp_name,e.designation,e.grade,e.collar,
                           e.department,e.section
                    FROM emp_training t LEFT JOIN employees e
                      ON e.emp_code=t.emp_code AND e.plant_id=t.plant_id
                    WHERE {" AND ".join(where)} ORDER BY t.id''', params), 1):
                ws4.append([r, t['emp_code'], t['emp_name'] or '', t['designation'] or '',
                            t['grade'] or '', t['collar'] or '', t['department'] or '',
                            t['section'] or '', t['start_date'] or '', t['end_date'] or '',
                            t['hrs'], t['programme_name'], t['prog_type'] or '',
                            t['level'] or '', t['mode'] or '', t['cal_new'] or '',
                            t['pre_rating'], t['post_rating'], t['venue'] or '', t['month'] or ''])

        # ── Sheet 5: 2C Programme Details ─────────────────────────────────────
        if 'prog_2c' in sheets:
            ws5 = wb.create_sheet('2C_Programme')
            ws5.append([tc(ws5, 'BALRAMPUR CHINI MILLS LIMITED')])
            ws5.append([tc(ws5, f'{pname} — 2C: PROGRAMME-WISE DETAILS | FY {fy}')])
            ws5.append([])
            ws5.append([hc(ws5, h) for h in
                ['Sr.','Session Code','Programme Name','Type','Level','Cal/New','Mode',
                 'Start Date','End Date','Audience','Hours Actual','Faculty Name','Int/Ext',
                 'Cost (Rs.)','Venue','Course FB','Faculty FB','Trainer FB-Participants',
                 'Trainer FB-Facilities','Participants','Man-Hours']])
            pax_map = {row[0]: row[1] for row in db.execute(
                'SELECT session_code, COUNT(*) FROM emp_training WHERE plant_id=? GROUP BY session_code', (plant_id,))}
            hrs_map = {row[0]: row[1] for row in db.execute(
                'SELECT session_code, COALESCE(SUM(hrs),0) FROM emp_training WHERE plant_id=? GROUP BY session_code', (plant_id,))}
            where  = ['p.plant_id=?']; params = [plant_id]
            join   = 'LEFT JOIN calendar c ON c.session_code=p.session_code AND c.plant_id=p.plant_id'
            if month_f: where.append('c.planned_month=?'); params.append(month_f)
            for r, p in enumerate(db.execute(
                    f'SELECT p.* FROM programme_details p {join} WHERE {" AND ".join(where)} ORDER BY p.id', params), 1):
                ws5.append([r, p['session_code'], p['programme_name'], p['prog_type'] or '',
                            p['level'] or '', p['cal_new'] or '', p['mode'] or '',
                            p['start_date'] or '', p['end_date'] or '', p['audience'] or '',
                            p['hours_actual'], p['faculty_name'] or '', p['int_ext'] or '',
                            p['cost'], p['venue'] or '', p['course_feedback'],
                            p['faculty_feedback'], p['trainer_fb_participants'],
                            p['trainer_fb_facilities'],
                            pax_map.get(p['session_code'], 0),
                            round(hrs_map.get(p['session_code'], 0), 1)])

        # ── suffix for filename ────────────────────────────────────────────────
        suffix = f"_{month_f}" if month_f else ''
        suffix += f"_{collar_f.replace(' ','')}" if collar_f else ''
        output   = io.BytesIO()
        wb.save(output); output.seek(0)
        filename = f"BCML_{plant['unit_code']}_Training_MIS_{fy.replace('-','')}{suffix}.xlsx"
        return send_file(output, download_name=filename, as_attachment=True,
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
