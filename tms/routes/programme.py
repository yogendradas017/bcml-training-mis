import io
import logging

from flask import render_template, request, redirect, url_for, session, flash, send_file, jsonify

from tms.constants import PROG_TYPES, MODES, AUDIENCES, MONTHS_FY, INT_EXT
from tms.db import get_db
from tms.decorators import spoc_required
from tms.helpers import (
    _is_ajax, _smart_title, _prog_in_use, _canonical_prog,
    _read_upload_file, _clean, _safe_float, _error_excel_response,
    _sync_master_from_tni,
)

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


def _register(app):

    @app.route('/programme-master')
    @spoc_required
    def programme_master():
        plant_id = session['plant_id']
        db = get_db()
        db.execute('''CREATE TABLE IF NOT EXISTS programme_master (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            plant_id INTEGER NOT NULL,
            name TEXT NOT NULL,
            prog_type TEXT, mode TEXT,
            created_at TEXT DEFAULT (date('now')),
            UNIQUE(plant_id, name))''')
        progs = db.execute(
            'SELECT * FROM programme_master WHERE plant_id=? ORDER BY name', (plant_id,)).fetchall()
        return render_template('programme_master.html', progs=progs,
                               prog_types=PROG_TYPES, modes=MODES)

    @app.route('/programme-master/add', methods=['POST'])
    @spoc_required
    def programme_master_add():
        plant_id  = session['plant_id']
        name      = _smart_title(request.form.get('name', '').strip())
        prog_type = request.form.get('prog_type', '').strip()
        raw_src   = request.form.get('source', '').strip()
        source    = raw_src if raw_src in ('TNI Requirement', 'New Requirement') else 'TNI Requirement'
        if not name:
            flash('Programme name is required.', 'danger')
            return redirect(url_for('programme_master'))
        db = get_db()
        try:
            existing = db.execute('SELECT id FROM programme_master WHERE plant_id=? AND LOWER(name)=LOWER(?)',
                                  (plant_id, name)).fetchone()
            if existing:
                db.execute('UPDATE programme_master SET prog_type=?, source=? WHERE id=?',
                           (prog_type or None, source, existing['id']))
                db.commit()
                flash(f'"{name}" updated.', 'success')
            else:
                db.execute('INSERT INTO programme_master(plant_id,name,prog_type,source) VALUES(?,?,?,?)',
                           (plant_id, name, prog_type or None, source))
                db.commit()
                flash(f'"{name}" added to master list as {source}.', 'success')
        except Exception as e:
            logging.error(f'programme_master_add error: {e}')
            flash(f'Error: {e}', 'danger')
        return redirect(url_for('programme_master'))

    @app.route('/programme-master/<int:prog_id>/delete', methods=['POST'])
    @spoc_required
    def programme_master_delete(prog_id):
        plant_id = session['plant_id']
        db = get_db()
        prog = db.execute('SELECT name FROM programme_master WHERE id=? AND plant_id=?',
                          (prog_id, plant_id)).fetchone()
        if not prog:
            flash('Programme not found.', 'danger')
            return redirect(url_for('programme_master'))
        if _prog_in_use(prog['name'], plant_id, db):
            flash(f'Cannot delete "{prog["name"]}" — it is referenced in TNI, Calendar, or Training Records.', 'danger')
            return redirect(url_for('programme_master'))
        db.execute('DELETE FROM programme_master WHERE id=? AND plant_id=?', (prog_id, plant_id))
        db.commit()
        flash(f'"{prog["name"]}" removed from master list.', 'warning')
        return redirect(url_for('programme_master'))

    @app.route('/programme-master/<int:prog_id>/set-type', methods=['POST'])
    @spoc_required
    def programme_master_set_type(prog_id):
        plant_id = session['plant_id']
        data     = request.get_json(silent=True) or {}
        prog_type = data.get('prog_type', '').strip()
        from tms.constants import PROG_TYPES
        if prog_type and prog_type not in PROG_TYPES:
            return jsonify({'ok': False, 'error': 'Invalid type'}), 400
        db = get_db()
        row = db.execute('SELECT id FROM programme_master WHERE id=? AND plant_id=?',
                         (prog_id, plant_id)).fetchone()
        if not row:
            return jsonify({'ok': False, 'error': 'Not found'}), 404
        db.execute('UPDATE programme_master SET prog_type=? WHERE id=? AND plant_id=?',
                   (prog_type or None, prog_id, plant_id))
        db.commit()
        return jsonify({'ok': True, 'prog_type': prog_type})

    @app.route('/programme-master/<int:prog_id>/set-source', methods=['POST'])
    @spoc_required
    def programme_master_set_source(prog_id):
        plant_id = session['plant_id']
        data     = request.get_json(silent=True) or {}
        source   = data.get('source', '').strip()
        if source not in ('TNI Requirement', 'New Requirement'):
            return jsonify({'ok': False, 'error': 'Invalid source'}), 400
        db = get_db()
        row = db.execute('SELECT id FROM programme_master WHERE id=? AND plant_id=?',
                         (prog_id, plant_id)).fetchone()
        if not row:
            return jsonify({'ok': False, 'error': 'Not found'}), 404
        db.execute('UPDATE programme_master SET source=? WHERE id=? AND plant_id=?',
                   (source, prog_id, plant_id))
        db.commit()
        return jsonify({'ok': True, 'source': source})

    @app.route('/programme-master/bulk-delete', methods=['POST'])
    @spoc_required
    def programme_master_bulk_delete():
        plant_id = session['plant_id']
        ids = request.form.getlist('ids[]')
        if not ids:
            flash('No programmes selected.', 'warning')
            return redirect(url_for('programme_master'))
        try:
            ids_int = [int(i) for i in ids]
        except ValueError:
            flash('Invalid selection.', 'danger')
            return redirect(url_for('programme_master'))
        db = get_db()
        blocked = []; deleted = 0
        for i in ids_int:
            prog = db.execute('SELECT name FROM programme_master WHERE id=? AND plant_id=?',
                              (i, plant_id)).fetchone()
            if not prog:
                continue
            if _prog_in_use(prog['name'], plant_id, db):
                blocked.append(prog['name'])
            else:
                db.execute('DELETE FROM programme_master WHERE id=? AND plant_id=?', (i, plant_id))
                deleted += 1
        db.commit()
        if deleted:
            flash(f'{deleted} programme(s) deleted.', 'warning')
        if blocked:
            flash(f'{len(blocked)} programme(s) could not be deleted (in use in TNI/Calendar/Training): {", ".join(blocked[:5])}', 'danger')
        return redirect(url_for('programme_master'))

    @app.route('/programme-master/sync-from-tni', methods=['POST'])
    @spoc_required
    def programme_master_sync_from_tni():
        plant_id = session['plant_id']
        db = get_db()
        tni_progs = [r[0] for r in db.execute(
            'SELECT DISTINCT programme_name FROM tni WHERE plant_id=? AND programme_name IS NOT NULL AND programme_name != "" ORDER BY programme_name',
            (plant_id,)).fetchall()]
        if not tni_progs:
            flash('No TNI data found — master list unchanged.', 'warning')
            return redirect(url_for('programme_master'))
        existing = {r['name']: r['prog_type'] for r in db.execute(
            'SELECT name, prog_type FROM programme_master WHERE plant_id=?', (plant_id,)).fetchall()}
        tni_types = {r['programme_name']: r['top_type'] for r in db.execute('''
            SELECT programme_name,
                   (SELECT prog_type FROM tni t2
                    WHERE t2.plant_id=t.plant_id AND t2.programme_name=t.programme_name
                      AND t2.prog_type IS NOT NULL AND t2.prog_type != ""
                    GROUP BY prog_type ORDER BY COUNT(*) DESC LIMIT 1) AS top_type
            FROM tni t WHERE plant_id=? AND programme_name IS NOT NULL AND programme_name != ""
            GROUP BY programme_name
        ''', (plant_id,)).fetchall()}
        db.execute('DELETE FROM programme_master WHERE plant_id=?', (plant_id,))
        for name in tni_progs:
            prog_type = tni_types.get(name) or existing.get(name)
            db.execute('INSERT OR IGNORE INTO programme_master(plant_id, name, prog_type, source) VALUES(?,?,?,?)',
                       (plant_id, name, prog_type, 'TNI Requirement'))
        db.commit()
        flash(f'Programme Master rebuilt from TNI data — {len(tni_progs)} unique programme(s).', 'success')
        return redirect(url_for('programme_master'))

    @app.route('/programme-master/bulk', methods=['POST'])
    @spoc_required
    def programme_master_bulk():
        plant_id = session['plant_id']
        f = request.files.get('file')
        if not f or f.filename == '':
            flash('No file selected.', 'danger')
            return redirect(url_for('programme_master'))
        try:
            import pandas as _pd
            fname = f.filename.lower()
            if fname.endswith('.csv'):
                df = _pd.read_csv(f, dtype=str).fillna('')
            else:
                df = _pd.read_excel(f, dtype=str).fillna('')
        except Exception as e:
            flash(f'Could not read file: {e}', 'danger')
            return redirect(url_for('programme_master'))

        cols_lower = {c.strip().lower(): c for c in df.columns}
        name_col = next((cols_lower[k] for k in ['programme name','program name','name','training name','course name'] if k in cols_lower), None)
        type_col = next((cols_lower[k] for k in ['type of programme','type','prog type','programme type'] if k in cols_lower), None)

        if not name_col:
            flash(f'Could not find a "Programme Name" column. Columns found: {", ".join(df.columns.tolist()[:10])}', 'danger')
            return redirect(url_for('programme_master'))

        src_col = next((cols_lower[k] for k in ['source','requirement','req type'] if k in cols_lower), None)
        db = get_db()
        inserted = skipped = 0
        for _, row in df.iterrows():
            name = str(row.get(name_col, '')).strip()
            if not name or name.lower() in ('nan', 'none', ''):
                continue
            prog_type = str(row.get(type_col, '')).strip() if type_col else ''
            raw_src   = str(row.get(src_col, '')).strip() if src_col else ''
            source    = raw_src if raw_src in ('TNI Requirement', 'New Requirement') else 'TNI Requirement'
            try:
                db.execute('INSERT INTO programme_master(plant_id,name,prog_type,source) VALUES(?,?,?,?)',
                           (plant_id, name, prog_type or None, source))
                inserted += 1
            except Exception:
                skipped += 1
        db.commit()
        flash(f'{inserted} programmes added. {skipped} already existed (skipped).', 'success' if inserted else 'warning')
        return redirect(url_for('programme_master'))

    @app.route('/programme-master/template')
    @spoc_required
    def programme_master_template():
        from openpyxl.worksheet.datavalidation import DataValidation
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = 'Programme Master'
        hdr_fill = PatternFill('solid', fgColor='1A1F35')
        headers = ['Programme Name', 'Type of Programme', 'Source']
        widths  = [45, 30, 20]
        for ci, (h, w) in enumerate(zip(headers, widths), 1):
            c = ws.cell(row=1, column=ci, value=h)
            c.fill = hdr_fill
            c.font = Font(bold=True, color='FFFFFF', size=11)
            c.alignment = Alignment(horizontal='center')
            ws.column_dimensions[get_column_letter(ci)].width = w
        ws.row_dimensions[1].height = 22
        for row in [('Fire Safety', 'EHS/HR', 'TNI Requirement'),
                    ('5-S Management', 'EHS/HR', 'TNI Requirement'),
                    ('English Communication', 'Behavioural/Leadership', 'New Requirement')]:
            ws.append(row)
        dv_type = DataValidation(type='list', formula1=f'"{",".join(PROG_TYPES)}"', allow_blank=True)
        dv_type.sqref = 'B2:B500'
        dv_src = DataValidation(type='list', formula1='"TNI Requirement,New Requirement"', allow_blank=True)
        dv_src.sqref = 'C2:C500'
        ws.add_data_validation(dv_type)
        ws.add_data_validation(dv_src)
        ws.freeze_panes = 'A2'
        buf = io.BytesIO(); wb.save(buf); buf.seek(0)
        return send_file(buf, as_attachment=True, download_name='Programme_Master_Template.xlsx',
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    @app.route('/programme-master/export')
    @spoc_required
    def programme_master_export():
        plant_id   = session['plant_id']
        plant_name = session.get('plant_name', 'Plant')
        db = get_db()
        progs = db.execute('SELECT name, prog_type, source, created_at FROM programme_master WHERE plant_id=? ORDER BY name', (plant_id,)).fetchall()
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = 'Programme Master'
        hdr_fill = PatternFill('solid', fgColor='1A1F35')
        headers = ['#', 'Programme Name', 'Type of Programme', 'Source', 'Added On']
        widths  = [5, 45, 30, 20, 14]
        for ci, (h, w) in enumerate(zip(headers, widths), 1):
            c = ws.cell(row=1, column=ci, value=h)
            c.fill = hdr_fill
            c.font = Font(bold=True, color='FFFFFF', size=11)
            c.alignment = Alignment(horizontal='center')
            ws.column_dimensions[get_column_letter(ci)].width = w
        ws.row_dimensions[1].height = 22
        ws.freeze_panes = 'A2'
        for i, r in enumerate(progs, 1):
            ws.append([i, r['name'], r['prog_type'] or '', r['source'] or 'TNI Requirement', r['created_at'] or ''])
        buf = io.BytesIO(); wb.save(buf); buf.seek(0)
        return send_file(buf, as_attachment=True,
                         download_name=f'Programme_Master_{plant_name}.xlsx',
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    # ── Programme Details (2C) ────────────────────────────────────────────────

    @app.route('/programme')
    @spoc_required
    def programme_details():
        plant_id = session['plant_id']
        db = get_db()
        records = db.execute('''
            SELECT p.*,
                   (SELECT COUNT(*) FROM emp_training t WHERE t.session_code=p.session_code AND t.plant_id=p.plant_id) as participants,
                   (SELECT COALESCE(SUM(t.hrs),0) FROM emp_training t WHERE t.session_code=p.session_code AND t.plant_id=p.plant_id) as man_hours
            FROM programme_details p
            WHERE p.plant_id=?
            ORDER BY p.id DESC
        ''', (plant_id,)).fetchall()
        cal_sessions = db.execute(
            "SELECT session_code, programme_name FROM calendar WHERE plant_id=? ORDER BY session_code",
            (plant_id,)).fetchall()
        return render_template('programme_2c.html', records=records,
                               cal_sessions=cal_sessions,
                               int_ext=INT_EXT, audiences=AUDIENCES, months=MONTHS_FY)

    @app.route('/programme/add', methods=['POST'])
    @spoc_required
    def add_programme_details():
        plant_id = session['plant_id']
        f = request.form
        db = get_db()
        session_code = f['session_code'].strip()

        if db.execute('SELECT 1 FROM programme_details WHERE session_code=? AND plant_id=?',
                      (session_code, plant_id)).fetchone():
            flash(f'Session {session_code} already recorded. Edit the existing entry.', 'warning')
            return redirect(url_for('programme_details'))

        cal = db.execute('SELECT * FROM calendar WHERE session_code=? AND plant_id=?',
                         (session_code, plant_id)).fetchone()
        prog_name = cal['programme_name'] if cal else f.get('programme_name','')
        prog_type = cal['prog_type']      if cal else ''
        level     = cal['level']          if cal else ''
        cal_new   = 'Calendar Program'    if cal else 'New Program'
        mode      = cal['mode']           if cal else ''
        audience  = cal['target_audience'] if cal else ''

        db.execute('''INSERT INTO programme_details
            (plant_id,session_code,programme_name,prog_type,level,cal_new,mode,
             start_date,end_date,audience,hours_actual,faculty_name,int_ext,cost,
             venue,course_feedback,faculty_feedback,trainer_fb_participants,trainer_fb_facilities)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
            (plant_id, session_code, prog_name, prog_type, level, cal_new, mode,
             f.get('start_date',''), f.get('end_date',''), audience,
             float(f.get('hours_actual') or 0), f.get('faculty_name',''),
             f.get('int_ext',''), float(f.get('cost') or 0),
             f.get('venue',''),
             _safe_float(f.get('course_feedback')),
             _safe_float(f.get('faculty_feedback')),
             _safe_float(f.get('trainer_fb_participants')),
             _safe_float(f.get('trainer_fb_facilities'))))
        db.commit()
        db.execute("UPDATE calendar SET status='Conducted' WHERE session_code=? AND plant_id=?",
                   (session_code, plant_id))
        db.commit()
        flash(f'Programme {session_code} details saved.', 'success')
        return redirect(url_for('programme_details'))

    @app.route('/programme/<int:rec_id>/delete', methods=['POST'])
    @spoc_required
    def delete_programme(rec_id):
        db = get_db()
        rec = db.execute('SELECT session_code FROM programme_details WHERE id=? AND plant_id=?',
                         (rec_id, session['plant_id'])).fetchone()
        if rec:
            db.execute('DELETE FROM programme_details WHERE id=? AND plant_id=?', (rec_id, session['plant_id']))
            db.execute("UPDATE calendar SET status='To Be Planned' WHERE session_code=? AND plant_id=?",
                       (rec['session_code'], session['plant_id']))
            db.commit()
        if _is_ajax():
            return '', 204
        flash('Programme record deleted.', 'warning')
        return redirect(url_for('programme_details'))

    @app.route('/programme/bulk-delete', methods=['POST'])
    @spoc_required
    def programme_bulk_delete():
        plant_id = session['plant_id']
        ids = request.form.getlist('ids[]')
        if ids:
            ph = ','.join('?' * len(ids))
            db = get_db()
            recs = db.execute(f'SELECT session_code FROM programme_details WHERE id IN ({ph}) AND plant_id=?',
                              ids + [plant_id]).fetchall()
            for r in recs:
                db.execute("UPDATE calendar SET status='To Be Planned' WHERE session_code=? AND plant_id=?",
                           (r['session_code'], plant_id))
            db.execute(f'DELETE FROM programme_details WHERE id IN ({ph}) AND plant_id=?', ids + [plant_id])
            db.commit()
            flash(f'{len(ids)} programme records deleted.', 'warning')
        return redirect(url_for('programme_details'))

    @app.route('/programme/template')
    @spoc_required
    def programme_template():
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = '2C_Bulk_Upload'
        headers = ['Session Code', 'Actual Start Date (YYYY-MM-DD)', 'Actual End Date (YYYY-MM-DD)',
                   'Actual Hours', 'Faculty Name', 'Internal/External', 'Cost (Rs.)', 'Venue',
                   'Course Feedback (1-5)', 'Faculty Feedback (1-5)',
                   'Trainer FB Participants (1-5)', 'Trainer FB Facilities (1-5)']
        hdr_fill = PatternFill('solid', fgColor='6B3FA0')
        hdr_font = Font(bold=True, color='FFFFFF')
        for i, h in enumerate(headers, 1):
            cell = ws.cell(row=1, column=i, value=h)
            cell.fill = hdr_fill; cell.font = hdr_font
            ws.column_dimensions[get_column_letter(i)].width = 26
        ws.append(['BCM/EHS/001/B01', '2026-06-10', '2026-06-10', 4, 'Mr. Ramesh Kumar', 'Internal', 0, 'Training Hall', 4.2, 4.0, 3.8, 4.1])
        ws['A4'] = 'NOTE: Session Code must exist in Training Calendar. Internal/External options: Internal | External | Online'
        out = io.BytesIO(); wb.save(out); out.seek(0)
        return send_file(out, download_name='2C_Programme_Bulk_Upload_Template.xlsx', as_attachment=True,
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    @app.route('/programme/bulk', methods=['POST'])
    @spoc_required
    def programme_bulk_upload():
        plant_id = session['plant_id']
        f = request.files.get('file')
        if not f or f.filename == '':
            flash('No file selected.', 'danger')
            return redirect(url_for('programme_details'))
        try:
            df = _read_upload_file(f)
        except Exception as e:
            flash(f'Could not read file: {e}', 'danger')
            return redirect(url_for('programme_details'))
        db = get_db(); inserted = 0; errors = []
        for i, row in df.iterrows():
            sc         = _clean(row, ['session code', 'session_code'])
            start_date = _clean(row, ['actual start date (yyyy-mm-dd)', 'start date', 'actual start date'])
            end_date   = _clean(row, ['actual end date (yyyy-mm-dd)', 'end date', 'actual end date'])
            hrs        = _safe_float(_clean(row, ['actual hours', 'hours', 'hrs'])) or 0
            faculty    = _clean(row, ['faculty name', 'faculty'])
            int_ext    = _clean(row, ['internal/external', 'int/ext', 'internal external'])
            cost       = _safe_float(_clean(row, ['cost (rs.)', 'cost'])) or 0
            venue      = _clean(row, ['venue'])
            cfb        = _safe_float(_clean(row, ['course feedback (1-5)', 'course feedback', 'course fb']))
            ffb        = _safe_float(_clean(row, ['faculty feedback (1-5)', 'faculty feedback', 'faculty fb']))
            tfbp       = _safe_float(_clean(row, ['trainer fb participants (1-5)', 'trainer fb participants']))
            tfbf       = _safe_float(_clean(row, ['trainer fb facilities (1-5)', 'trainer fb facilities']))
            if not sc:
                errors.append(f'Row {i+2}: Session Code is required.')
                continue
            cal = db.execute('SELECT * FROM calendar WHERE session_code=? AND plant_id=?', (sc, plant_id)).fetchone()
            if not cal:
                errors.append(f'Row {i+2}: Session Code {sc} not found in Calendar.')
                continue
            if db.execute('SELECT 1 FROM programme_details WHERE session_code=? AND plant_id=?', (sc, plant_id)).fetchone():
                errors.append(f'Row {i+2}: Session {sc} already has programme details recorded.')
                continue
            if not start_date:
                errors.append(f'Row {i+2}: Actual Start Date is required.')
                continue
            db.execute('''INSERT INTO programme_details
                (plant_id,session_code,programme_name,prog_type,level,cal_new,mode,
                 start_date,end_date,audience,hours_actual,faculty_name,int_ext,cost,
                 venue,course_feedback,faculty_feedback,trainer_fb_participants,trainer_fb_facilities)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
                (plant_id, sc, cal['programme_name'], cal['prog_type'], cal['level'],
                 'Calendar Program', cal['mode'], start_date, end_date,
                 cal['target_audience'], hrs, faculty, int_ext, cost, venue, cfb, ffb, tfbp, tfbf))
            db.execute("UPDATE calendar SET status='Conducted' WHERE session_code=? AND plant_id=?",
                       (sc, plant_id))
            inserted += 1
        db.commit()
        if errors:
            if inserted:
                flash(f'Bulk upload complete: {inserted} programme records saved. {len(errors)} rows had errors — downloading error report.', 'warning')
            return _error_excel_response(errors, inserted, 'Programme2C_Upload_Errors.xlsx')
        flash(f'Bulk upload complete: {inserted} programme records saved.', 'success')
        return redirect(url_for('programme_details'))
