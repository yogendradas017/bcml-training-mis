from datetime import date
from flask import render_template, request, redirect, url_for, session, flash
import sqlite3

from tms.constants import GENDERS
from tms.db import get_db
from tms.decorators import spoc_required
from tms.helpers import normalise_collar


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
        return render_template('employees.html', employees=emps, show_exited=show_exited,
                               recent_exited=recent_exited,
                               genders=GENDERS, today=str(date.today()))

    @app.route('/employees/add', methods=['POST'])
    @spoc_required
    def add_employee():
        plant_id = session['plant_id']
        f = request.form
        db = get_db()
        collar = normalise_collar(f.get('collar', ''))
        try:
            db.execute('''INSERT INTO employees
                (plant_id,emp_code,name,designation,grade,collar,department,section,
                 category,gender,physically_handicapped,remarks)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?)''',
                (plant_id, f['emp_code'].strip(), f['name'].strip(),
                 f.get('designation',''), f.get('grade',''), collar,
                 f.get('department',''), f.get('section',''), f.get('category',''),
                 f.get('gender',''), f.get('physically_handicapped','No'),
                 f.get('remarks','')))
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
