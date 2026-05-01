import io
from datetime import datetime as _dt

from flask import render_template, request, redirect, url_for, session, flash, send_file

from tms.constants import MONTHS_FY
from tms.db import get_db
from tms.decorators import spoc_required
from tms.helpers import (
    _is_ajax, _canonical_prog, _date_to_month, _safe_float,
    _read_upload_file, _clean, _error_excel_response,
    _current_fy, _in_current_fy,
)

import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter


def _register(app):

    @app.route('/training')
    @spoc_required
    def emp_training():
        plant_id = session['plant_id']
        db = get_db()
        records = db.execute('''
            SELECT t.*, e.name as emp_name, e.designation, e.grade, e.collar,
                   e.department, e.section
            FROM emp_training t
            LEFT JOIN employees e ON e.emp_code=t.emp_code AND e.plant_id=t.plant_id
            WHERE t.plant_id=?
            ORDER BY t.id DESC
        ''', (plant_id,)).fetchall()
        emps = db.execute(
            'SELECT emp_code, name FROM employees WHERE plant_id=? AND is_active=1 ORDER BY name',
            (plant_id,)).fetchall()
        sessions_list = db.execute(
            "SELECT session_code, programme_name FROM calendar WHERE plant_id=? ORDER BY session_code",
            (plant_id,)).fetchall()
        return render_template('training_2a.html', records=records, employees=emps,
                               sessions=sessions_list, months=MONTHS_FY)

    @app.route('/training/add', methods=['POST'])
    @spoc_required
    def add_emp_training():
        plant_id = session['plant_id']
        f = request.form
        db = get_db()
        emp_code     = f['emp_code']
        session_code = f.get('session_code', '').strip()
        start_date   = f.get('start_date', '')
        end_date     = f.get('end_date', '')

        prog_name = _canonical_prog(f.get('programme_name', '').strip(), plant_id, db)
        prog_type = level = mode = cal_new = ''
        if session_code:
            cal = db.execute('SELECT * FROM calendar WHERE session_code=? AND plant_id=?',
                             (session_code, plant_id)).fetchone()
            if cal:
                prog_name  = cal['programme_name']
                prog_type  = cal['prog_type']
                level      = cal['level']
                mode       = cal['mode']
                cal_new    = 'Calendar Program'
                if not start_date: start_date = cal['plan_start'] or ''
                if not end_date:   end_date   = cal['plan_end'] or ''
            else:
                flash(f'Session code "{session_code}" not found in calendar — record saved without calendar link.', 'warning')

        if not prog_type and prog_name:
            mr = db.execute('SELECT prog_type FROM programme_master WHERE plant_id=? AND LOWER(name)=LOWER(?)',
                            (plant_id, prog_name)).fetchone()
            if mr and mr['prog_type']:
                prog_type = mr['prog_type']

        hrs = float(f.get('hrs') or 0)
        if hrs <= 0:
            flash('Training hours must be greater than 0.', 'danger')
            return redirect(url_for('emp_training'))
        fy_start, fy_end = _current_fy()
        if start_date and not _in_current_fy(start_date):
            flash(f'Training date must be within the current financial year ({fy_start} to {fy_end}).', 'danger')
            return redirect(url_for('emp_training'))

        month = _date_to_month(start_date)
        db.execute('''INSERT INTO emp_training
            (plant_id,emp_code,session_code,programme_name,start_date,end_date,
             hrs,prog_type,level,mode,cal_new,pre_rating,post_rating,venue,month)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
            (plant_id, emp_code, session_code, prog_name,
             start_date, end_date, hrs,
             prog_type, level, mode, cal_new,
             _safe_float(f.get('pre_rating')), _safe_float(f.get('post_rating')),
             f.get('venue',''), month))
        db.commit()
        flash('Training record added.', 'success')
        return redirect(url_for('emp_training'))

    @app.route('/training/<int:rec_id>/delete', methods=['POST'])
    @spoc_required
    def delete_emp_training(rec_id):
        db = get_db()
        db.execute('DELETE FROM emp_training WHERE id=? AND plant_id=?', (rec_id, session['plant_id']))
        db.commit()
        if _is_ajax():
            return '', 204
        flash('Training record deleted.', 'warning')
        return redirect(url_for('emp_training'))

    @app.route('/training/bulk-delete', methods=['POST'])
    @spoc_required
    def training_bulk_delete():
        plant_id = session['plant_id']
        ids = request.form.getlist('ids[]')
        if ids:
            db = get_db()
            deleted = 0
            for i in range(0, len(ids), 900):
                chunk = ids[i:i+900]
                ph = ','.join('?' * len(chunk))
                db.execute(f'DELETE FROM emp_training WHERE id IN ({ph}) AND plant_id=?', chunk + [plant_id])
                deleted += len(chunk)
            db.commit()
            flash(f'{deleted} training records deleted.', 'warning')
        return redirect(url_for('emp_training'))

    @app.route('/training/template')
    @spoc_required
    def training_template():
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = '2A_Bulk_Upload'
        headers = ['Employee Code', 'Session Code (optional)', 'Programme Name',
                   'Type of Programme', 'Start Date (DD-MM-YYYY)', 'End Date (DD-MM-YYYY)',
                   'Hours', 'Venue', 'Pre-Session Rating (1-5)', 'Post-Session Rating (1-5)']
        hdr_fill = PatternFill('solid', start_color='1F4E79')
        hdr_font = Font(bold=True, color='FFFFFF')
        for i, h in enumerate(headers, 1):
            cell = ws.cell(row=1, column=i, value=h)
            cell.fill = hdr_fill; cell.font = hdr_font
            ws.column_dimensions[get_column_letter(i)].width = 26
        ws.append(['21700011', 'BCM/EHS/001/B01', 'Fire Safety Training', 'EHS/HR', '10-06-2026', '10-06-2026', 4, 'Training Hall', 3.5, 4.2])
        ws.append(['21101568', '', 'MS Office Basics', 'IT/MIS', '05-07-2026', '06-07-2026', 8, 'Computer Lab', '', 4.0])
        ws['A5'] = 'NOTE: Session Code is optional. If provided, Programme Name/Type/Mode auto-fill from Calendar.'
        out = io.BytesIO()
        wb.save(out); out.seek(0)
        return send_file(out, download_name='2A_Training_Bulk_Upload_Template.xlsx', as_attachment=True,
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    @app.route('/training/bulk', methods=['POST'])
    @spoc_required
    def training_bulk_upload():
        plant_id = session['plant_id']
        f = request.files.get('file')
        if not f or f.filename == '':
            flash('No file selected.', 'danger')
            return redirect(url_for('emp_training'))
        try:
            df = _read_upload_file(f)
        except Exception as e:
            flash(f'Could not read file: {e}', 'danger')
            return redirect(url_for('emp_training'))

        def _parse_date(raw):
            """Normalize any date format to YYYY-MM-DD for storage."""
            s = str(raw).strip()[:10] if raw else ''
            if not s:
                return ''
            for fmt in ('%d-%m-%Y', '%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y'):
                try:
                    return _dt.strptime(s, fmt).strftime('%Y-%m-%d')
                except ValueError:
                    continue
            return s

        db = get_db(); inserted = 0; errors = []
        for i, row in df.iterrows():
            emp_code     = _clean(row, ['employee code', 'emp code', 'empcode'])
            session_code = _clean(row, ['session code', 'session code (optional)', 'sessioncode'])
            prog_name    = _clean(row, ['programme name', 'program name', 'training name'])
            file_type    = _clean(row, ['type of programme', 'type', 'prog type', 'programme type'])
            start_date   = _parse_date(_clean(row, ['start date', 'start date (dd-mm-yyyy)', 'start date (yyyy-mm-dd)', 'startdate', 'date']))
            end_date     = _parse_date(_clean(row, ['end date', 'end date (dd-mm-yyyy)', 'end date (yyyy-mm-dd)', 'enddate']))
            hrs          = _safe_float(_clean(row, ['hours', 'hrs', 'duration'])) or 0
            venue        = _clean(row, ['venue'])
            pre_r        = _safe_float(_clean(row, ['pre-session rating', 'pre rating', 'pre_rating']))
            post_r       = _safe_float(_clean(row, ['post-session rating', 'post rating', 'post_rating']))

            if not emp_code:
                errors.append(f'Row {i+2}: Employee Code is required.')
                continue
            emp = db.execute('SELECT 1 FROM employees WHERE emp_code=? AND plant_id=? AND is_active=1',
                             (emp_code, plant_id)).fetchone()
            if not emp:
                errors.append(f'Row {i+2}: Employee {emp_code} not found.')
                continue

            prog_type = level = mode = cal_new = ''
            if session_code:
                cal = db.execute('SELECT * FROM calendar WHERE session_code=? AND plant_id=?',
                                 (session_code, plant_id)).fetchone()
                if cal:
                    prog_name  = prog_name or cal['programme_name']
                    prog_type  = cal['prog_type']
                    level      = cal['level']
                    mode       = cal['mode']
                    cal_new    = 'Calendar Program'
                    start_date = start_date or _parse_date(cal['plan_start'] or '')
                    end_date   = end_date or _parse_date(cal['plan_end'] or '')

            if not prog_name:
                errors.append(f'Row {i+2}: Programme Name required (no session code matched).')
                continue

            if not prog_type:
                prog_type = file_type  # use column from file if session/master didn't fill it
            if not prog_type and prog_name:
                mr = db.execute('SELECT prog_type FROM programme_master WHERE plant_id=? AND LOWER(name)=LOWER(?)',
                                (plant_id, prog_name)).fetchone()
                if mr and mr['prog_type']:
                    prog_type = mr['prog_type']
            if not prog_type:
                errors.append(f'Row {i+2}: Type of Programme is mandatory — fill the "Type of Programme" column.')
                continue

            if not cal_new:
                prog_name = _canonical_prog(prog_name, plant_id, db)
            month = _date_to_month(start_date)
            db.execute('''INSERT INTO emp_training
                (plant_id,emp_code,session_code,programme_name,start_date,end_date,
                 hrs,prog_type,level,mode,cal_new,pre_rating,post_rating,venue,month)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
                (plant_id, emp_code, session_code, prog_name,
                 start_date, end_date, hrs, prog_type, level, mode, cal_new,
                 pre_r, post_r, venue, month))
            inserted += 1
        db.commit()
        if errors:
            if inserted:
                flash(f'Bulk upload complete: {inserted} records added. {len(errors)} rows had errors — downloading error report.', 'warning')
            return _error_excel_response(errors, inserted, 'Training2A_Upload_Errors.xlsx')
        flash(f'Bulk upload complete: {inserted} training records added successfully.', 'success')
        return redirect(url_for('emp_training'))
