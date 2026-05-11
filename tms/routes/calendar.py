import io

from flask import render_template, request, redirect, url_for, session, flash, send_file

from tms.constants import PROG_TYPES, MODES, LEVELS, AUDIENCES, MONTHS_FY, STATUSES
from tms.db import get_db
from tms.decorators import spoc_required
from tms.helpers import (
    _is_ajax, _canonical_prog, _get_or_create_prog_code, _new_session_code,
    _derive_audience, _sync_calendar_from_2c,
    _read_upload_file, _clean, _safe_float, _error_excel_response,
    _current_fy, _in_current_fy, _parse_date_strict,
)

import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from tms.audit import log_action


def _register(app):

    @app.route('/calendar')
    @spoc_required
    def training_calendar():
        plant_id = session['plant_id']
        db = get_db()
        _sync_calendar_from_2c(plant_id, db)

        sessions = db.execute('SELECT * FROM calendar WHERE plant_id=? ORDER BY id DESC', (plant_id,)).fetchall()
        demand_map = {}
        for row in db.execute('SELECT programme_name, COUNT(DISTINCT emp_code) as cnt FROM tni WHERE plant_id=? GROUP BY programme_name', (plant_id,)):
            demand_map[row['programme_name']] = row['cnt']

        master_programmes = [r[0] for r in db.execute(
            'SELECT name FROM programme_master WHERE plant_id=? ORDER BY name', (plant_id,)).fetchall()] or []
        all_cal_programmes = master_programmes
        tni_programmes = [r[0] for r in db.execute(
            'SELECT DISTINCT programme_name FROM tni WHERE plant_id=? ORDER BY programme_name', (plant_id,))]

        cov_rows = []
        pax_map  = {}
        for s in sessions:
            p = s['programme_name']
            if p not in pax_map:
                pax_map[p] = {'sessions': 0, 'planned_pax': 0, 'conducted_pax': 0}
            pax_map[p]['sessions']     += 1
            pax_map[p]['planned_pax']  += (s['planned_pax'] or 0)
            if s['status'] == 'Conducted':
                pax_map[p]['conducted_pax'] += (s['planned_pax'] or 0)
        for prog, d in demand_map.items():
            pm = pax_map.get(prog, {'sessions': 0, 'planned_pax': 0, 'conducted_pax': 0})
            planned_pax = pm['planned_pax']
            gap         = max(0, d - planned_pax)
            pct         = min(100, round(planned_pax / d * 100)) if d > 0 else 0
            cov_rows.append({'name': prog, 'demand': d,
                             'sessions': pm['sessions'], 'planned_pax': planned_pax,
                             'conducted_pax': pm['conducted_pax'],
                             'gap': gap, 'pct': pct,
                             'over': max(0, planned_pax - d)})
        cov_rows.sort(key=lambda x: x['gap'], reverse=True)

        qr_rows = db.execute(
            'SELECT session_code, stage, token, is_active FROM session_qr WHERE plant_id=?',
            (plant_id,)
        ).fetchall()
        qr_map = {}
        for q in qr_rows:
            qr_map.setdefault(q['session_code'], {})[q['stage']] = dict(q)

        return render_template('calendar.html', sessions=sessions, demand_map=demand_map,
                               tni_programmes=tni_programmes,
                               all_cal_programmes=all_cal_programmes, cov_rows=cov_rows,
                               prog_types=PROG_TYPES, modes=MODES, levels=LEVELS,
                               audiences=AUDIENCES, months=MONTHS_FY, statuses=STATUSES,
                               qr_map=qr_map)

    @app.route('/calendar/add', methods=['POST'])
    @spoc_required
    def add_calendar():
        plant_id = session['plant_id']
        f = request.form
        db = get_db()
        prog_name_raw = f['programme_name'].strip()
        prog_name = _canonical_prog(prog_name_raw, plant_id, db, strict=True)
        if prog_name is None:
            flash(f'Programme "{prog_name_raw}" not found in Programme Master. Add it to the master list first.', 'danger')
            return redirect(url_for('training_calendar'))
        prog_type = f.get('prog_type', '')

        prog_code    = _get_or_create_prog_code(plant_id, prog_name, prog_type, db)
        session_code = _new_session_code(plant_id, prog_code, db)

        tni_audience  = _derive_audience(plant_id, prog_name, db)
        form_audience = f.get('target_audience', '')
        audience      = tni_audience if tni_audience else form_audience

        try:
            dur = float(f.get('duration_hrs') or 0)
        except (ValueError, TypeError):
            dur = 0
        if dur <= 0:
            flash('Duration must be greater than 0 hours.', 'danger')
            return redirect(url_for('training_calendar'))
        fy_start, fy_end = _current_fy()
        for fld, lbl in [('plan_start', 'Plan Start'), ('plan_end', 'Plan End')]:
            val = f.get(fld, '')
            if val and not _in_current_fy(val):
                flash(f'{lbl} date must be within the current financial year ({fy_start} to {fy_end}).', 'danger')
                return redirect(url_for('training_calendar'))

        db.execute('''INSERT INTO calendar
            (plant_id,prog_code,session_code,source,programme_name,prog_type,
             planned_month,plan_start,plan_end,time_from,time_to,duration_hrs,
             level,mode,target_audience,planned_pax,trainer_vendor,status)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
            (plant_id, prog_code, session_code,
             f.get('source','TNI Driven') if f.get('source','') in ('TNI Driven','New Requirement') else 'TNI Driven',
             prog_name, prog_type,
             f.get('planned_month',''), f.get('plan_start',''), f.get('plan_end',''),
             f.get('time_from',''), f.get('time_to',''),
             dur,
             f.get('level',''), f.get('mode',''), audience,
             int(f.get('planned_pax') or 0) if str(f.get('planned_pax') or '0').isdigit() else 0,
             f.get('trainer_vendor',''),
             'To Be Planned'))
        db.commit()
        log_action('RECORD_ADD', f"cal:{session_code}")
        msg = f'Session {session_code} added.'
        if tni_audience and form_audience and form_audience != tni_audience:
            msg += f' Audience set to "{tni_audience}" (locked from TNI — TNI has both BC & WC employees).'
        flash(msg, 'success')
        return redirect(url_for('training_calendar'))

    @app.route('/calendar/<int:cal_id>/delete', methods=['POST'])
    @spoc_required
    def delete_calendar(cal_id):
        db = get_db()
        cal = db.execute('SELECT session_code, status FROM calendar WHERE id=? AND plant_id=?',
                         (cal_id, session['plant_id'])).fetchone()
        if cal and cal['status'] == 'Conducted':
            if _is_ajax():
                return 'Conducted sessions cannot be deleted.', 403
            flash('Conducted sessions cannot be deleted.', 'danger')
            return redirect(url_for('training_calendar'))
        if cal:
            db.execute('DELETE FROM session_qr WHERE plant_id=? AND session_code=?',
                       (session['plant_id'], cal['session_code']))
        db.execute('DELETE FROM calendar WHERE id=? AND plant_id=?', (cal_id, session['plant_id']))
        db.commit()
        log_action('RECORD_DELETE', f"cal:{cal_id}")
        if _is_ajax():
            return '', 204
        flash('Calendar entry deleted.', 'warning')
        return redirect(url_for('training_calendar'))

    @app.route('/calendar/<int:cal_id>/edit', methods=['POST'])
    @spoc_required
    def edit_calendar(cal_id):
        plant_id = session['plant_id']
        f = request.form
        db = get_db()
        existing = db.execute('SELECT status, session_code FROM calendar WHERE id=? AND plant_id=?',
                              (cal_id, plant_id)).fetchone()
        if existing and existing['status'] == 'Conducted':
            flash('Conducted sessions cannot be edited.', 'danger')
            return redirect(url_for('training_calendar'))
        if f.get('status') == 'Conducted' and session.get('role') != 'admin':
            sc = existing['session_code'] if existing else None
            has_qr = sc and db.execute(
                'SELECT 1 FROM session_qr WHERE plant_id=? AND session_code=? AND is_active=1',
                (plant_id, sc)).fetchone()
            has_feedback = sc and db.execute(
                'SELECT 1 FROM feedback_response WHERE plant_id=? AND session_code=?',
                (plant_id, sc)).fetchone()
            if not has_qr or not has_feedback:
                flash("Can't mark Conducted. Process: Generate QR code → Mark Attendance → Collect Feedback. Contact Corporate L&D for assistance.", 'danger')
                return redirect(url_for('training_calendar'))
        edit_prog_raw     = f.get('programme_name','').strip()
        edit_prog         = _canonical_prog(edit_prog_raw, plant_id, db, strict=True)
        if edit_prog is None:
            flash(f'Programme "{edit_prog_raw}" not found in Programme Master. Add it to the master list first.', 'danger')
            return redirect(url_for('training_calendar'))
        tni_audience_edit = _derive_audience(plant_id, edit_prog, db)
        form_audience_edit = f.get('target_audience', '')
        edit_audience     = tni_audience_edit if tni_audience_edit else form_audience_edit

        try:
            dur = float(f.get('duration_hrs') or 0)
        except (ValueError, TypeError):
            dur = 0
        if dur <= 0:
            flash('Duration must be greater than 0 hours.', 'danger')
            return redirect(url_for('training_calendar'))
        fy_start, fy_end = _current_fy()
        for fld, lbl in [('plan_start', 'Plan Start'), ('plan_end', 'Plan End')]:
            val = f.get(fld, '')
            if val and not _in_current_fy(val):
                flash(f'{lbl} date must be within the current financial year ({fy_start} to {fy_end}).', 'danger')
                return redirect(url_for('training_calendar'))

        db.execute('''UPDATE calendar SET
            programme_name=?, prog_type=?, source=?, planned_month=?,
            plan_start=?, plan_end=?, time_from=?, time_to=?,
            duration_hrs=?, level=?, mode=?, target_audience=?,
            planned_pax=?, trainer_vendor=?, status=?
            WHERE id=? AND plant_id=?''',
            (edit_prog, f.get('prog_type',''),
             f.get('source','TNI Driven') if f.get('source','') in ('TNI Driven','New Requirement') else 'TNI Driven',
             f.get('planned_month',''),
             f.get('plan_start',''), f.get('plan_end',''),
             f.get('time_from',''), f.get('time_to',''),
             dur, f.get('level',''),
             f.get('mode',''), edit_audience,
             int(f.get('planned_pax') or 0) if str(f.get('planned_pax') or '0').isdigit() else 0,
             f.get('trainer_vendor',''),
             f.get('status','To Be Planned'),
             cal_id, plant_id))
        db.commit()
        log_action('RECORD_EDIT', f"cal:{cal_id}")
        msg = 'Session updated.'
        if tni_audience_edit and form_audience_edit and form_audience_edit != tni_audience_edit:
            msg += f' Audience locked to "{tni_audience_edit}" from TNI.'
        flash(msg, 'success')
        return redirect(url_for('training_calendar'))

    @app.route('/calendar/bulk-delete', methods=['POST'])
    @spoc_required
    def calendar_bulk_delete():
        plant_id = session['plant_id']
        ids = request.form.getlist('ids[]')
        if ids:
            db = get_db()
            deleted = 0
            for i in range(0, len(ids), 900):
                chunk = ids[i:i+900]
                ph = ','.join('?' * len(chunk))
                codes = db.execute(
                    f'SELECT session_code FROM calendar WHERE id IN ({ph}) AND plant_id=? AND status != "Conducted"',
                    chunk + [plant_id]
                ).fetchall()
                for c in codes:
                    db.execute('DELETE FROM session_qr WHERE plant_id=? AND session_code=?',
                               (plant_id, c['session_code']))
                db.execute(f'DELETE FROM calendar WHERE id IN ({ph}) AND plant_id=? AND status != "Conducted"', chunk + [plant_id])
                deleted += len(chunk)
            db.commit()
            log_action('BULK_DELETE', f"cal:{deleted}")
            flash(f'{deleted} calendar sessions deleted.', 'warning')
        return redirect(url_for('training_calendar'))

    @app.route('/calendar/template')
    @spoc_required
    def calendar_template():
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = 'Calendar_Bulk_Upload'
        headers = ['Programme Name', 'Type of Programme', 'Source', 'Planned Month',
                   'Plan Start (DD-MM-YYYY)', 'Plan End (DD-MM-YYYY)', 'Duration (Hrs)',
                   'Level', 'Mode', 'Target Audience', 'Planned Pax', 'Trainer/Vendor']
        hdr_fill = PatternFill('solid', fgColor='1F4E79')
        hdr_font = Font(bold=True, color='FFFFFF')
        for i, h in enumerate(headers, 1):
            cell = ws.cell(row=1, column=i, value=h)
            cell.fill = hdr_fill; cell.font = hdr_font
            ws.column_dimensions[get_column_letter(i)].width = 24
        ws.append(['Fire Safety Training', 'EHS/HR', 'TNI Driven', 'June', '10-06-2026', '10-06-2026', 4, 'General', 'Classroom', 'Blue Collared', 30, 'Internal Faculty'])
        ws.append(['Leadership Skills', 'Behavioural/Leadership', 'New Requirement', 'July', '05-07-2026', '06-07-2026', 8, 'Specialized', 'Classroom', 'White Collared', 20, 'External Vendor'])
        ws['A5'] = 'NOTE: Dates MUST be DD-MM-YYYY (e.g. 15-06-2026).'
        ws['A6'] = 'VALID Types: Behavioural/Leadership | Cane | Commercial | EHS/HR | IT | Technical'
        ws['A7'] = 'VALID Modes: Classroom | OJT | SOP | Online'
        ws['A8'] = 'VALID Audience: Blue Collared | White Collared | Common'
        ws['A9'] = 'VALID Months: April | May | June | July | August | September | October | November | December | January | February | March'
        out = io.BytesIO(); wb.save(out); out.seek(0)
        return send_file(out, download_name='Calendar_Bulk_Upload_Template.xlsx', as_attachment=True,
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    @app.route('/calendar/bulk', methods=['POST'])
    @spoc_required
    def calendar_bulk_upload():
        plant_id = session['plant_id']
        f = request.files.get('file')
        if not f or f.filename == '':
            flash('No file selected.', 'danger')
            return redirect(url_for('training_calendar'))
        try:
            df = _read_upload_file(f)
        except Exception as e:
            flash(f'Could not read file: {e}', 'danger')
            return redirect(url_for('training_calendar'))
        db = get_db(); inserted = 0; errors = []
        for i, row in df.iterrows():
            prog_name = _clean(row, ['programme name', 'programme_name', 'program name'])
            prog_type = _clean(row, ['type of programme', 'type', 'prog type'])
            raw_src   = _clean(row, ['source']) or ''
            source    = raw_src if raw_src in ('TNI Driven', 'New Requirement') else 'TNI Driven'
            month     = _clean(row, ['planned month', 'month'])
            raw_start = _clean(row, ['plan start (dd-mm-yyyy)', 'plan start (yyyy-mm-dd)', 'plan start', 'start date'])
            raw_end   = _clean(row, ['plan end (dd-mm-yyyy)', 'plan end (yyyy-mm-dd)', 'plan end', 'end date'])
            try:
                plan_start = _parse_date_strict(raw_start)
                plan_end   = _parse_date_strict(raw_end)
            except ValueError as e:
                errors.append(f'Row {i+2}: Date format error — {e}. Use DD-MM-YYYY (e.g. 15-06-2026).')
                continue
            duration  = _safe_float(_clean(row, ['duration (hrs)', 'duration', 'hrs'])) or 0
            level     = _clean(row, ['level'])
            mode      = _clean(row, ['mode'])
            audience  = _clean(row, ['target audience', 'audience'])
            pax       = int(_safe_float(_clean(row, ['planned pax', 'pax'])) or 0)
            trainer   = _clean(row, ['trainer/vendor', 'trainer', 'vendor'])
            if not prog_name:
                errors.append(f'Row {i+2}: Programme Name is required.')
                continue
            if not prog_type:
                errors.append(f'Row {i+2}: Type of Programme is required.')
                continue
            canonical = _canonical_prog(prog_name, plant_id, db, strict=True)
            if canonical is None:
                errors.append(f'Row {i+2}: Programme "{prog_name}" not in Programme Master — add it first.')
                continue
            prog_name    = canonical
            tni_aud      = _derive_audience(plant_id, prog_name, db)
            audience     = tni_aud if tni_aud else audience
            prog_code    = _get_or_create_prog_code(plant_id, prog_name, prog_type, db)
            session_code = _new_session_code(plant_id, prog_code, db)
            db.execute('''INSERT INTO calendar
                (plant_id,prog_code,session_code,source,programme_name,prog_type,
                 planned_month,plan_start,plan_end,duration_hrs,level,mode,
                 target_audience,planned_pax,trainer_vendor,status)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,'To Be Planned')''',
                (plant_id, prog_code, session_code, source, prog_name, prog_type,
                 month, plan_start, plan_end, duration, level, mode, audience, pax, trainer))
            inserted += 1
        db.commit()
        if errors:
            if inserted:
                flash(f'Bulk upload complete: {inserted} sessions added. {len(errors)} rows had errors — downloading error report.', 'warning')
            return _error_excel_response(errors, inserted, 'Calendar_Upload_Errors.xlsx')
        log_action('BULK_UPLOAD', f"cal:{inserted}")
        flash(f'Bulk upload complete: {inserted} sessions added to calendar.', 'success')
        return redirect(url_for('training_calendar'))
