import io
import os
from datetime import date
from flask import render_template, request, redirect, url_for, session, flash, send_file
import sqlite3

from tms.constants import GENDERS, GRADES, CATEGORIES, COLLARS, PH_OPTIONS, TEMP_UPLOAD_DIR
from tms.db import get_db
from tms.decorators import spoc_required
from tms.helpers import normalise_collar

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    _XLSX = True
except ImportError:
    _XLSX = False


def _plant_depts(db, plant_id):
    rows = db.execute(
        "SELECT DISTINCT UPPER(TRIM(department)) FROM employees "
        "WHERE plant_id=? AND department IS NOT NULL AND department!='' ORDER BY 1",
        (plant_id,)).fetchall()
    return [r[0] for r in rows]


def _plant_sections(db, plant_id):
    rows = db.execute(
        "SELECT DISTINCT UPPER(TRIM(section)) FROM employees "
        "WHERE plant_id=? AND section IS NOT NULL AND section!='' ORDER BY 1",
        (plant_id,)).fetchall()
    return [r[0] for r in rows]


def _validate_emp_fields(f, departments, sections):
    """Return list of error strings. Empty list = all good."""
    errs = []
    grade = (f.get('grade') or '').strip().upper()
    collar = normalise_collar(f.get('collar', ''))
    dept = (f.get('department') or '').strip().upper()
    sect = (f.get('section') or '').strip().upper()
    gender = (f.get('gender') or '').strip()
    ph = (f.get('physically_handicapped') or '').strip()

    if grade and grade not in GRADES:
        errs.append(f"Invalid grade '{f.get('grade')}'. Must be one of the predefined grades.")
    if collar and collar not in COLLARS:
        errs.append(f"Invalid collar '{f.get('collar')}'. Must be Blue Collared or White Collared.")
    # Dept/section: not strictly enforced — new values are allowed (e.g. new wings added)
    if gender and gender not in GENDERS:
        errs.append(f"Invalid gender '{gender}'. Must be Male, Female, or Others.")
    if ph and ph not in PH_OPTIONS:
        errs.append(f"Physically Handicapped must be Yes or No.")
    return errs


def _register(app):

    @app.route('/employees')
    @spoc_required
    def employees():
        plant_id = session['plant_id']
        db = get_db()
        show_exited = request.args.get('show_exited', '0') == '1'
        if show_exited:
            emps = db.execute('SELECT * FROM employees WHERE plant_id=? ORDER BY name', (plant_id,)).fetchall()
        else:
            emps = db.execute('SELECT * FROM employees WHERE plant_id=? AND is_active=1 ORDER BY name', (plant_id,)).fetchall()
        recent_exited = db.execute(
            "SELECT * FROM employees WHERE plant_id=? AND is_active=0 AND exit_date >= date('now','-7 days') ORDER BY exit_date DESC",
            (plant_id,)).fetchall()
        departments = _plant_depts(db, plant_id)
        sections = _plant_sections(db, plant_id)
        return render_template('employees.html',
                               employees=emps, show_exited=show_exited,
                               recent_exited=recent_exited,
                               genders=GENDERS, grades=GRADES, categories=CATEGORIES,
                               collars=COLLARS, ph_options=PH_OPTIONS,
                               departments=departments, sections=sections,
                               today=str(date.today()))

    @app.route('/employees/add', methods=['POST'])
    @spoc_required
    def add_employee():
        plant_id = session['plant_id']
        f = request.form
        db = get_db()
        departments = _plant_depts(db, plant_id)
        sections = _plant_sections(db, plant_id)
        errs = _validate_emp_fields(f, departments, sections)
        if errs:
            for e in errs:
                flash(e, 'danger')
            return redirect(url_for('employees'))
        collar = normalise_collar(f.get('collar', ''))
        grade = (f.get('grade') or '').strip().upper() or ''
        dept = (f.get('department') or '').strip().upper() or ''
        sect = (f.get('section') or '').strip().upper() or ''
        cat = (f.get('category') or '').strip().upper() or ''
        try:
            db.execute('''INSERT INTO employees
                (plant_id,emp_code,name,designation,grade,collar,department,section,
                 category,gender,physically_handicapped,remarks)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?)''',
                (plant_id, f['emp_code'].strip(), f['name'].strip(),
                 f.get('designation', '').strip(), grade, collar,
                 dept, sect, cat,
                 f.get('gender', ''), f.get('physically_handicapped', 'No'),
                 f.get('remarks', '')))
            db.commit()
            flash(f"Employee {f['name'].strip()} added successfully.", 'success')
        except sqlite3.IntegrityError:
            flash(f"Employee code {f['emp_code'].strip()} already exists.", 'danger')
        return redirect(url_for('employees'))

    @app.route('/employees/<int:emp_id>/exit', methods=['POST'])
    @spoc_required
    def exit_employee(emp_id):
        db = get_db()
        exit_date   = request.form.get('exit_date', str(date.today()))
        exit_reason = request.form.get('exit_reason', '')
        if exit_date > str(date.today()):
            flash('Exit date cannot be a future date.', 'danger')
            return redirect(url_for('employees'))
        if not exit_reason.strip():
            flash('Exit reason is mandatory for attrition analysis.', 'danger')
            return redirect(url_for('employees'))
        db.execute('UPDATE employees SET is_active=0, exit_date=?, exit_reason=? WHERE id=? AND plant_id=?',
                   (exit_date, exit_reason, emp_id, session['plant_id']))
        db.commit()
        flash('Employee marked as exited.', 'warning')
        return redirect(url_for('employees'))

    @app.route('/employees/<int:emp_id>/reactivate', methods=['POST'])
    @spoc_required
    def reactivate_employee(emp_id):
        db = get_db()
        db.execute('UPDATE employees SET is_active=1, exit_date=NULL, exit_reason=NULL WHERE id=? AND plant_id=?',
                   (emp_id, session['plant_id']))
        db.commit()
        flash('Employee reactivated.', 'success')
        return redirect(url_for('employees') + '?show_exited=1')

    @app.route('/employees/bulk-template')
    @spoc_required
    def emp_bulk_template():
        if not _XLSX:
            flash('openpyxl not installed.', 'danger')
            return redirect(url_for('employees'))
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = 'Employees'

        headers = ['Emp Code *', 'Full Name *', 'Designation', 'Grade', 'Collar *',
                   'Department', 'Section', 'Category', 'Gender', 'Physically Handicapped', 'Remarks']
        hdr_fill = PatternFill('solid', fgColor='1A3A5C')
        hdr_font = Font(color='FFFFFF', bold=True)
        for ci, h in enumerate(headers, 1):
            c = ws.cell(row=1, column=ci, value=h)
            c.fill = hdr_fill
            c.font = hdr_font
            c.alignment = Alignment(horizontal='center')

        # Reference sheet with allowed values
        ref = wb.create_sheet('Reference')
        ref_data = {
            'Grade': GRADES,
            'Collar': COLLARS,
            'Gender': GENDERS,
            'Physically Handicapped': PH_OPTIONS,
            'Category': CATEGORIES,
        }
        col = 1
        for header, vals in ref_data.items():
            ref.cell(row=1, column=col, value=header).font = Font(bold=True)
            for ri, v in enumerate(vals, 2):
                ref.cell(row=ri, column=col, value=v)
            col += 1

        for col_idx, width in zip(range(1, len(headers) + 1), [14, 30, 20, 22, 16, 20, 22, 22, 10, 20, 20]):
            ws.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = width

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return send_file(buf, as_attachment=True, download_name='employee_bulk_upload_template.xlsx',
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    @app.route('/employees/bulk-upload', methods=['POST'])
    @spoc_required
    def emp_bulk_upload():
        if not _XLSX:
            flash('openpyxl not installed.', 'danger')
            return redirect(url_for('employees'))

        plant_id = session['plant_id']
        f = request.files.get('bulk_file')
        if not f or not f.filename.endswith('.xlsx'):
            flash('Please upload a valid .xlsx file.', 'danger')
            return redirect(url_for('employees'))

        db = get_db()
        departments = _plant_depts(db, plant_id)
        sections = _plant_sections(db, plant_id)
        grades_upper = [g.upper() for g in GRADES]

        try:
            wb = openpyxl.load_workbook(f, read_only=True, data_only=True)
            ws = wb.active
        except Exception as e:
            flash(f'Could not read file: {e}', 'danger')
            return redirect(url_for('employees'))

        rows = list(ws.iter_rows(min_row=2, values_only=True))
        inserted = 0
        errors = []

        # Pre-load existing emp codes for this plant to detect DB duplicates upfront
        existing_codes = {
            r[0] for r in db.execute(
                'SELECT emp_code FROM employees WHERE plant_id=?', (plant_id,)).fetchall()
        }
        seen_in_file = {}  # emp_code -> first row number (detect within-file duplicates)

        for i, row in enumerate(rows, start=2):
            if not any(row):
                continue
            emp_code   = str(row[0]).strip() if row[0] else ''
            name       = str(row[1]).strip() if row[1] else ''
            desig      = str(row[2]).strip() if row[2] else ''
            grade_raw  = str(row[3]).strip() if row[3] else ''
            collar_raw = str(row[4]).strip() if row[4] else ''
            dept_raw   = str(row[5]).strip() if row[5] else ''
            sect_raw   = str(row[6]).strip() if row[6] else ''
            cat_raw    = str(row[7]).strip() if row[7] else ''
            gender_raw = str(row[8]).strip() if row[8] else ''
            ph_raw     = str(row[9]).strip() if row[9] else 'No'
            remarks    = str(row[10]).strip() if row[10] else ''

            row_errors = []

            if not emp_code:
                row_errors.append('Emp Code is required')
            if not name:
                row_errors.append('Full Name is required')

            grade  = grade_raw.upper()
            collar = normalise_collar(collar_raw)
            dept   = dept_raw.upper()
            sect   = sect_raw.upper()
            cat    = cat_raw.upper()
            gender = gender_raw
            ph     = ph_raw if ph_raw in PH_OPTIONS else ''

            if not collar_raw:
                row_errors.append('Collar is required')
            elif collar not in COLLARS:
                row_errors.append(f"Invalid collar '{collar_raw}' (must be Blue Collared / White Collared)")

            if grade_raw and grade not in grades_upper:
                row_errors.append(f"Invalid grade '{grade_raw}'")

            if gender_raw and gender not in GENDERS:
                row_errors.append(f"Invalid gender '{gender_raw}' (must be Male / Female / Others)")

            if ph_raw and ph_raw not in PH_OPTIONS:
                row_errors.append(f"Invalid PH value '{ph_raw}' (must be Yes / No)")

            # Duplicate checks — report before attempting insert
            if emp_code:
                if emp_code in existing_codes:
                    row_errors.append(f"Emp code {emp_code} already exists in the system — skipped")
                elif emp_code in seen_in_file:
                    row_errors.append(f"Emp code {emp_code} appears again (first at row {seen_in_file[emp_code]}) — skipped")
                else:
                    seen_in_file[emp_code] = i

            if row_errors:
                errors.append(f"Row {i} ({emp_code or '?'}): {'; '.join(row_errors)}")
                continue

            db.execute('''INSERT INTO employees
                (plant_id,emp_code,name,designation,grade,collar,department,section,
                 category,gender,physically_handicapped,remarks)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?)''',
                (plant_id, emp_code, name, desig, grade, collar,
                 dept, sect, cat, gender, ph or 'No', remarks))
            existing_codes.add(emp_code)
            inserted += 1

        db.commit()

        if inserted:
            flash(f'{inserted} employee(s) uploaded successfully.', 'success')
        if errors:
            for e in errors:
                flash(e, 'upload_error')
        if not inserted and not errors:
            flash('No data rows found in the file.', 'warning')

        return redirect(url_for('employees'))
