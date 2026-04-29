import os
import re
import io
import glob
import time
import json as _json
import uuid as _uuid
from datetime import date, datetime
from difflib import get_close_matches

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from flask import request, send_file

from tms.constants import (
    BASE_DIR, TEMP_UPLOAD_DIR, PLANT_MAP, TYPE_ABBREV,
    PROG_TYPES, MODES, MONTHS_FY, MONTH_NUM, NON_TNI_SOURCES
)

# ── Upload helpers ────────────────────────────────────────────────────────────

def _is_ajax():
    return request.headers.get('X-Requested-With') == 'XMLHttpRequest'


def _read_upload_file(file_storage):
    import pandas as pd
    fname = file_storage.filename.lower()
    if fname.endswith('.csv'):
        return pd.read_csv(file_storage, dtype=str).fillna('')
    return pd.read_excel(file_storage, dtype=str).fillna('')


def _read_upload_file_path(path):
    import pandas as pd
    if path.lower().endswith('.csv'):
        return pd.read_csv(path, dtype=str).fillna('')
    return pd.read_excel(path, dtype=str).fillna('')


def _clean(row, keys):
    for k in keys:
        for col in row.index:
            if str(col).strip().lower() == k:
                val = str(row[col]).strip()
                return '' if val.lower() in ('nan', 'none', '') else val
    return ''


def normalise_collar(val):
    v = str(val).strip().upper()
    if any(x in v for x in ['WHITE', 'WC', 'W C']):
        return 'White Collared'
    if any(x in v for x in ['BLUE', 'BC', 'B C']):
        return 'Blue Collared'
    return val.strip()


def _safe_float(val):
    try:
        return float(val) if val and str(val).strip() != '' else None
    except (ValueError, TypeError):
        return None


def _current_fy():
    """Returns (fy_start, fy_end) as 'YYYY-MM-DD' strings for the current financial year (Apr–Mar)."""
    today = date.today()
    yr = today.year if today.month >= 4 else today.year - 1
    return f'{yr}-04-01', f'{yr+1}-03-31'


def _in_current_fy(date_str):
    """True if date_str falls within the current FY, or is empty/None."""
    if not date_str:
        return True
    try:
        d = date.fromisoformat(str(date_str)[:10])
        s, e = _current_fy()
        return date.fromisoformat(s) <= d <= date.fromisoformat(e)
    except (ValueError, TypeError):
        return False


def _date_to_month(date_str):
    if not date_str:
        return ''
    try:
        d = datetime.strptime(str(date_str)[:10], '%Y-%m-%d')
        return d.strftime('%B')
    except Exception:
        return ''


def _get_programme_names(plant_id, db):
    rows = db.execute(
        'SELECT DISTINCT programme_name FROM calendar WHERE plant_id=? ORDER BY programme_name',
        (plant_id,)).fetchall()
    return [r['programme_name'] for r in rows]


# ── Audience derivation ───────────────────────────────────────────────────────

def _derive_audience(plant_id, prog_name, db):
    collars = db.execute('''
        SELECT DISTINCT e.collar FROM tni t
        JOIN employees e ON e.emp_code=t.emp_code AND e.plant_id=t.plant_id
        WHERE t.plant_id=? AND LOWER(t.programme_name)=LOWER(?) AND e.collar IS NOT NULL AND e.collar != ''
    ''', (plant_id, prog_name)).fetchall()
    collar_set = {r['collar'] for r in collars}
    if not collar_set:
        return None
    if 'Blue Collared' in collar_set and 'White Collared' in collar_set:
        return 'Common'
    if 'Blue Collared' in collar_set:
        return 'Blue Collared'
    return 'White Collared'


# ── Session code helpers ──────────────────────────────────────────────────────

def _fy_label():
    """Returns short FY label like '26-27' for use in session codes."""
    today = date.today()
    y = today.year
    return f'{str(y-1)[2:]}-{str(y)[2:]}' if today.month < 4 else f'{str(y)[2:]}-{str(y+1)[2:]}'


def _get_or_create_prog_code(plant_id, prog_name, prog_type, db):
    existing = db.execute(
        'SELECT prog_code FROM calendar WHERE plant_id=? AND programme_name=? LIMIT 1',
        (plant_id, prog_name)).fetchone()
    if existing:
        return existing['prog_code']
    unit_code = PLANT_MAP[plant_id]['unit_code']
    abbrev    = TYPE_ABBREV.get(prog_type, 'GEN')
    count     = db.execute(
        "SELECT COUNT(DISTINCT prog_code) FROM calendar WHERE plant_id=? AND prog_code LIKE ?",
        (plant_id, f'{unit_code}/{abbrev}/%')).fetchone()[0]
    return f'{unit_code}/{abbrev}/{count+1:03d}'


def _new_session_code(plant_id, prog_code, db):
    fy    = _fy_label()
    count = db.execute(
        'SELECT COUNT(*) FROM calendar WHERE plant_id=? AND prog_code=? AND session_code LIKE ?',
        (plant_id, prog_code, f'{prog_code}/{fy}/%')).fetchone()[0]
    return f'{prog_code}/{fy}/B{count+1:02d}'


# ── Calendar sync ─────────────────────────────────────────────────────────────

def _sync_calendar_from_2c(plant_id, db):
    db.execute('''UPDATE calendar SET status='Conducted'
        WHERE plant_id=? AND session_code IN
        (SELECT session_code FROM programme_details WHERE plant_id=?)''', (plant_id, plant_id))
    programmes = db.execute(
        'SELECT DISTINCT programme_name FROM calendar WHERE plant_id=?', (plant_id,)).fetchall()
    for row in programmes:
        tni_aud = _derive_audience(plant_id, row['programme_name'], db)
        if tni_aud:
            db.execute('''UPDATE calendar SET target_audience=?
                WHERE plant_id=? AND LOWER(programme_name)=LOWER(?) AND target_audience!=?''',
                (tni_aud, plant_id, row['programme_name'], tni_aud))
    db.commit()


# ── Programme master helpers ──────────────────────────────────────────────────

def _sync_master_from_tni(plant_id, db):
    rows = db.execute('''
        SELECT programme_name,
               (SELECT prog_type FROM tni t2
                WHERE t2.plant_id=t.plant_id AND t2.programme_name=t.programme_name
                  AND t2.prog_type IS NOT NULL AND t2.prog_type != ""
                GROUP BY prog_type ORDER BY COUNT(*) DESC LIMIT 1) AS top_type,
               CASE WHEN EXISTS(
                   SELECT 1 FROM tni t3 WHERE t3.plant_id=t.plant_id
                   AND t3.programme_name=t.programme_name AND t3.source='TNI Driven'
               ) THEN 'TNI Requirement' ELSE 'New Requirement' END AS derived_source
        FROM tni t
        WHERE plant_id=? AND programme_name IS NOT NULL AND programme_name != ""
        GROUP BY programme_name
    ''', (plant_id,)).fetchall()
    for r in rows:
        db.execute('''INSERT OR IGNORE INTO programme_master(plant_id, name, prog_type, source)
                      VALUES(?,?,?,?)''',
                   (plant_id, r['programme_name'], r['top_type'], r['derived_source']))
        if r['top_type']:
            db.execute('''UPDATE programme_master SET prog_type=?
                          WHERE plant_id=? AND LOWER(name)=LOWER(?) AND (prog_type IS NULL OR prog_type="")''',
                       (r['top_type'], plant_id, r['programme_name']))
        db.execute('''UPDATE programme_master SET source=?
                      WHERE plant_id=? AND LOWER(name)=LOWER(?)''',
                   (r['derived_source'], plant_id, r['programme_name']))


def _prog_in_use(prog_name, plant_id, db):
    for table in ('tni', 'calendar', 'emp_training'):
        if db.execute(
                f'SELECT 1 FROM {table} WHERE plant_id=? AND LOWER(programme_name)=LOWER(?) LIMIT 1',
                (plant_id, prog_name)).fetchone():
            return True
    return False


# ── Summary calculations ──────────────────────────────────────────────────────

def _calc_summary(plant_id, month_filter, db):
    rows = []
    mn = MONTH_NUM.get(month_filter, '') if month_filter else ''
    month_clause_2c = f"AND strftime('%m', p.start_date) = '{mn}'" if mn else ("AND 1=0" if month_filter else "")

    for pt in PROG_TYPES:
        clause = "AND t.month=?" if month_filter else ""
        p_pt   = [plant_id, pt] + ([month_filter] if month_filter else [])
        p_bc   = [plant_id, pt, 'Blue Collared']  + ([month_filter] if month_filter else [])
        p_wc   = [plant_id, pt, 'White Collared'] + ([month_filter] if month_filter else [])

        pq = db.execute(f'''SELECT
            SUM(CASE WHEN audience='Blue Collared'  THEN 1 ELSE 0 END),
            SUM(CASE WHEN audience='White Collared' THEN 1 ELSE 0 END),
            SUM(CASE WHEN audience='Common'         THEN 1 ELSE 0 END),
            COUNT(DISTINCT session_code),
            SUM(CASE WHEN int_ext='Internal' THEN 1 ELSE 0 END),
            SUM(CASE WHEN int_ext='External' THEN 1 ELSE 0 END)
            FROM programme_details p
            WHERE p.plant_id=? AND p.prog_type=? {month_clause_2c}''',
            [plant_id, pt]).fetchone()
        bc_progs     = pq[0] or 0
        wc_progs     = pq[1] or 0
        common_progs = pq[2] or 0
        total_progs  = pq[3] or 0
        int_prog     = pq[4] or 0
        ext_prog     = pq[5] or 0

        bc_seats = db.execute(f'''SELECT COUNT(*) FROM emp_training t
            JOIN employees e ON e.emp_code=t.emp_code AND e.plant_id=t.plant_id
            WHERE t.plant_id=? AND t.prog_type=? AND e.collar=? {clause}''', p_bc).fetchone()[0]
        wc_seats = db.execute(f'''SELECT COUNT(*) FROM emp_training t
            JOIN employees e ON e.emp_code=t.emp_code AND e.plant_id=t.plant_id
            WHERE t.plant_id=? AND t.prog_type=? AND e.collar=? {clause}''', p_wc).fetchone()[0]
        bc_hrs = db.execute(f'''SELECT COALESCE(SUM(t.hrs),0) FROM emp_training t
            JOIN employees e ON e.emp_code=t.emp_code AND e.plant_id=t.plant_id
            WHERE t.plant_id=? AND t.prog_type=? AND e.collar=? {clause}''', p_bc).fetchone()[0]
        wc_hrs = db.execute(f'''SELECT COALESCE(SUM(t.hrs),0) FROM emp_training t
            JOIN employees e ON e.emp_code=t.emp_code AND e.plant_id=t.plant_id
            WHERE t.plant_id=? AND t.prog_type=? AND e.collar=? {clause}''', p_wc).fetchone()[0]

        bc_fixed = db.execute('''SELECT COUNT(DISTINCT t.emp_code) FROM tni t
            JOIN employees e ON e.emp_code=t.emp_code AND e.plant_id=t.plant_id
            WHERE t.plant_id=? AND t.prog_type=? AND e.collar='Blue Collared' ''',
            [plant_id, pt]).fetchone()[0]
        wc_fixed = db.execute('''SELECT COUNT(DISTINCT t.emp_code) FROM tni t
            JOIN employees e ON e.emp_code=t.emp_code AND e.plant_id=t.plant_id
            WHERE t.plant_id=? AND t.prog_type=? AND e.collar='White Collared' ''',
            [plant_id, pt]).fetchone()[0]

        bc_cum = db.execute('''SELECT COUNT(DISTINCT et.emp_code) FROM emp_training et
            JOIN employees e ON e.emp_code=et.emp_code AND e.plant_id=et.plant_id
            JOIN tni t ON t.emp_code=et.emp_code AND t.plant_id=et.plant_id
              AND LOWER(t.prog_type)=LOWER(et.prog_type)
            WHERE et.plant_id=? AND et.prog_type=? AND e.collar='Blue Collared' ''',
            [plant_id, pt]).fetchone()[0]
        wc_cum = db.execute('''SELECT COUNT(DISTINCT et.emp_code) FROM emp_training et
            JOIN employees e ON e.emp_code=et.emp_code AND e.plant_id=et.plant_id
            JOIN tni t ON t.emp_code=et.emp_code AND t.plant_id=et.plant_id
              AND LOWER(t.prog_type)=LOWER(et.prog_type)
            WHERE et.plant_id=? AND et.prog_type=? AND e.collar='White Collared' ''',
            [plant_id, pt]).fetchone()[0]

        bc_cov  = round(bc_cum  / bc_fixed  * 100, 1) if bc_fixed  else 0
        wc_cov  = round(wc_cum  / wc_fixed  * 100, 1) if wc_fixed  else 0
        tot_cov = round((bc_cum + wc_cum) / (bc_fixed + wc_fixed) * 100, 1) if (bc_fixed + wc_fixed) else 0

        rows.append({
            'prog_type':    pt,
            'bc_progs':     bc_progs,    'wc_progs':  wc_progs,
            'common_progs': common_progs,'total_progs': total_progs,
            'int_prog':     int_prog,    'ext_prog':  ext_prog,
            'bc_seats':     bc_seats,    'wc_seats':  wc_seats,
            'total_seats':  bc_seats + wc_seats,
            'bc_hrs':       round(bc_hrs, 1), 'wc_hrs': round(wc_hrs, 1),
            'total_hrs':    round(bc_hrs + wc_hrs, 1),
            'bc_fixed':     bc_fixed,    'wc_fixed':  wc_fixed,
            'bc_cum':       bc_cum,      'wc_cum':    wc_cum,
            'bc_cov':       bc_cov,      'wc_cov':    wc_cov,
            'tot_cov':      tot_cov,
        })
    return rows


def _calc_totals(rows):
    if not rows:
        return {}
    t = {k: 0 for k in rows[0]}
    t['prog_type'] = 'TOTAL'
    skip = {'prog_type', 'bc_cov', 'wc_cov', 'tot_cov'}
    for r in rows:
        for k, v in r.items():
            if k not in skip:
                t[k] = round(t.get(k, 0) + (v or 0), 1)
    t['bc_cov']  = round(t['bc_cum']  / t['bc_fixed']  * 100, 1) if t.get('bc_fixed')  else 0
    t['wc_cov']  = round(t['wc_cum']  / t['wc_fixed']  * 100, 1) if t.get('wc_fixed')  else 0
    t['tot_cov'] = round((t['bc_cum'] + t['wc_cum']) / (t['bc_fixed'] + t['wc_fixed']) * 100, 1) \
                   if (t.get('bc_fixed', 0) + t.get('wc_fixed', 0)) else 0
    return t


def _calc_compliance(plant_id, db):
    bc = db.execute(
        "SELECT COUNT(*) FROM employees WHERE plant_id=? AND is_active=1 AND collar='Blue Collared'",
        (plant_id,)).fetchone()[0]
    wc = db.execute(
        "SELECT COUNT(*) FROM employees WHERE plant_id=? AND is_active=1 AND collar='White Collared'",
        (plant_id,)).fetchone()[0]
    bc_act = db.execute('''SELECT COALESCE(SUM(t.hrs),0) FROM emp_training t
        JOIN employees e ON e.emp_code=t.emp_code AND e.plant_id=t.plant_id
        WHERE t.plant_id=? AND e.collar='Blue Collared' ''', (plant_id,)).fetchone()[0]
    wc_act = db.execute('''SELECT COALESCE(SUM(t.hrs),0) FROM emp_training t
        JOIN employees e ON e.emp_code=t.emp_code AND e.plant_id=t.plant_id
        WHERE t.plant_id=? AND e.collar='White Collared' ''', (plant_id,)).fetchone()[0]
    bc_mandate = bc * 12
    wc_mandate = wc * 24
    return {
        'bc_emp': bc, 'wc_emp': wc,
        'bc_mandate': bc_mandate, 'wc_mandate': wc_mandate,
        'bc_actual': round(bc_act, 1), 'wc_actual': round(wc_act, 1),
        'bc_pct': round(bc_act / bc_mandate * 100, 1) if bc_mandate else 0,
        'wc_pct': round(wc_act / wc_mandate * 100, 1) if wc_mandate else 0,
        'total_pct': round((bc_act + wc_act) / (bc_mandate + wc_mandate) * 100, 1)
                     if (bc_mandate + wc_mandate) else 0
    }


# ── Excel error response ──────────────────────────────────────────────────────

def _error_excel_response(errors, inserted, download_name='Upload_Errors.xlsx'):
    wb  = openpyxl.Workbook()
    ws  = wb.active
    ws.title = 'Failed Rows'
    ws.append([f'{inserted} rows imported successfully. {len(errors)} rows failed — details below.'])
    ws['A1'].font = Font(bold=True, size=12)
    ws.merge_cells('A1:C1')
    ws.append([])
    hdr = ['Row #', 'Error Reason', 'Tip']
    ws.append(hdr)
    for c, h in enumerate(hdr, 1):
        cell = ws.cell(row=3, column=c)
        cell.font      = Font(bold=True, color='FFFFFF')
        cell.fill      = PatternFill('solid', fgColor='C0392B')
        cell.alignment = Alignment(horizontal='center')
    for err in errors:
        parts   = err.split(':', 1)
        row_ref = parts[0].strip() if len(parts) == 2 else ''
        reason  = parts[1].strip() if len(parts) == 2 else err
        tip = ''
        rl  = reason.lower()
        if 'not found in your plant'   in rl: tip = 'Check employee code is registered under this plant in Employee Master'
        elif 'required'                in rl: tip = 'This column must not be empty'
        elif 'month'                   in rl: tip = 'Use: April / May / June / July / August / September / October / November / December / January / February / March'
        elif 'type' in rl or 'prog'    in rl: tip = 'Use: Behavioural/Leadership | Cane | Commercial | EHS/HR | IT | Technical'
        elif 'mode'                    in rl: tip = 'Use: Classroom | OJT | SOP | Online'
        elif 'date'                    in rl: tip = 'Use format YYYY-MM-DD'
        elif 'hours' in rl or 'hrs'    in rl: tip = 'Must be a number e.g. 4 or 2.5'
        elif 'session'                 in rl: tip = 'Session Code must already exist in Training Calendar'
        elif 'employee'                in rl: tip = 'Employee Code must exist in Employee Master for this plant'
        ws.append([row_ref, reason, tip])
    ws.column_dimensions['A'].width = 10
    ws.column_dimensions['B'].width = 55
    ws.column_dimensions['C'].width = 65
    for row in ws.iter_rows(min_row=4):
        if row[0].row % 2 == 0:
            for cell in row:
                cell.fill = PatternFill('solid', fgColor='FFF5F5')
    out = io.BytesIO()
    wb.save(out); out.seek(0)
    return send_file(out, download_name=download_name, as_attachment=True,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


# ── DB cleanse helpers (called from init_db) ──────────────────────────────────

def _cleanse_master_spelling(db):
    rows = db.execute('SELECT id, name FROM programme_master').fetchall()
    for row in rows:
        cleaned = _smart_title(_apply_word_fixes(row['name']))
        if cleaned != row['name']:
            clash = db.execute(
                'SELECT id FROM programme_master WHERE plant_id=(SELECT plant_id FROM programme_master WHERE id=?) AND LOWER(name)=LOWER(?) AND id!=?',
                (row['id'], cleaned, row['id'])
            ).fetchone()
            if not clash:
                db.execute('UPDATE programme_master SET name=? WHERE id=?', (cleaned, row['id']))
    db.commit()


def _cleanse_programme_names(db, plant_id=None):
    from difflib import get_close_matches as gcm
    report = {}
    plants = [plant_id] if plant_id else [
        r[0] for r in db.execute('SELECT DISTINCT plant_id FROM programme_master').fetchall()]
    for pid in plants:
        master = [r[0] for r in db.execute(
            'SELECT name FROM programme_master WHERE plant_id=? ORDER BY name', (pid,)).fetchall()]
        if not master:
            continue
        master_lower_map = {m.lower(): m for m in master}
        master_lower = list(master_lower_map.keys())
        fixed = 0; merged = 0
        for table in ('tni', 'emp_training', 'calendar'):
            rows = db.execute(f'SELECT id, programme_name FROM {table} WHERE plant_id=?', (pid,)).fetchall()
            for row in rows:
                raw = row['programme_name'] or ''
                if not raw:
                    continue
                raw_lower = raw.lower()
                if raw_lower in master_lower_map:
                    canonical = master_lower_map[raw_lower]
                    if canonical != raw:
                        db.execute(f'UPDATE {table} SET programme_name=? WHERE id=?', (canonical, row['id']))
                        fixed += 1
                else:
                    m = gcm(raw_lower, master_lower, n=1, cutoff=0.88)
                    if m:
                        canonical = master_lower_map[m[0]]
                        db.execute(f'UPDATE {table} SET programme_name=? WHERE id=?', (canonical, row['id']))
                        fixed += 1
        dupes = db.execute('''
            SELECT emp_code, programme_name, MIN(id) as keep_id, COUNT(*) as cnt
            FROM tni WHERE plant_id=?
            GROUP BY emp_code, programme_name HAVING cnt > 1
        ''', (pid,)).fetchall()
        for d in dupes:
            db.execute('DELETE FROM tni WHERE plant_id=? AND emp_code=? AND programme_name=? AND id != ?',
                       (pid, d['emp_code'], d['programme_name'], d['keep_id']))
            merged += 1
        db.commit()
        report[pid] = {'fixed': fixed, 'merged': merged}
    return report


def _cleanup_stale_analyze_files():
    pattern = os.path.join(BASE_DIR, 'data', 'tni_analyze_*.json')
    cutoff  = time.time() - 86400
    for path in glob.glob(pattern):
        try:
            if os.path.getmtime(path) < cutoff:
                os.remove(path)
        except Exception:
            pass


# ── Fresh TNI upload ──────────────────────────────────────────────────────────

def _poka_yoke_clean_prog(name):
    if not name:
        return ''
    s = re.sub(r'[\x00-\x1f\x7f]', '', str(name).strip())
    s = re.sub(r'\s+', ' ', s).strip()
    s = _apply_word_fixes(s)
    return _smart_title(s)


def _process_fresh_tni(df, plant_id, db):
    emp_rows = db.execute(
        'SELECT emp_code, name FROM employees WHERE plant_id=? AND is_active=1', (plant_id,)
    ).fetchall()
    emp_map   = {r['emp_code']: r['name'] for r in emp_rows}
    emp_upper = {k.upper(): k for k in emp_map}

    cols     = df.columns.tolist()
    col_emp  = _detect_col(cols, ['emp code','employee code','empcode','staff code','emp id','employee id','code'])
    col_prog = _detect_col(cols, ['programme name','program name','training name','course name','training need','training'])
    col_type = _detect_col(cols, ['type of programme','type','programme type','prog type','training type','category'])
    col_mode = _detect_col(cols, ['mode','training mode','delivery mode'])
    col_hrs  = _detect_col(cols, ['planned hours','hours','hrs','duration'])

    def gv(row, col):
        if not col: return ''
        v = str(row.get(col, '') or '').strip()
        return '' if v.lower() in ('nan', 'none', '0', '') else v

    valid_rows = []; error_rows = []; name_corrections = {}
    seen = set(); duplicate_count = 0

    for i, row in df.iterrows():
        raw_emp  = gv(row, col_emp)
        raw_prog = gv(row, col_prog)
        prog_type = gv(row, col_type)
        mode      = gv(row, col_mode)
        hours     = _safe_float(gv(row, col_hrs)) or 0

        if not raw_emp and not raw_prog:
            continue
        if not raw_emp:
            error_rows.append({'row': i+2, 'emp_code': '', 'prog_name': raw_prog, 'reason': 'Employee Code missing'})
            continue
        if not raw_prog:
            error_rows.append({'row': i+2, 'emp_code': raw_emp, 'prog_name': '', 'reason': 'Programme Name missing'})
            continue

        clean_emp = raw_emp
        if raw_emp not in emp_map:
            if raw_emp.upper() in emp_upper:
                clean_emp = emp_upper[raw_emp.upper()]
            else:
                error_rows.append({'row': i+2, 'emp_code': raw_emp, 'prog_name': raw_prog,
                                   'reason': f'Employee "{raw_emp}" not found in this plant'})
                continue

        cleaned = _poka_yoke_clean_prog(raw_prog)
        if cleaned != raw_prog:
            name_corrections[raw_prog] = cleaned

        key = (clean_emp, cleaned.lower())
        if key in seen:
            duplicate_count += 1
            continue
        seen.add(key)
        valid_rows.append({
            'emp_code': clean_emp, 'programme_name': cleaned,
            'prog_type': prog_type, 'mode': mode, 'hours': hours,
        })

    return {
        'total_rows':       len(df),
        'valid_rows':       valid_rows,
        'error_rows':       error_rows,
        'name_corrections': name_corrections,
        'unique_progs':     sorted(set(r['programme_name'] for r in valid_rows)),
        'duplicate_count':  duplicate_count,
    }


# ── MS Forms import ───────────────────────────────────────────────────────────

_MSFORMS_SKIP_HEADERS = {'id','start time','completion time','email','name','responder'}


def _parse_msforms_excel(file_storage, plant_id, db):
    import pandas as pd
    raw = file_storage.read()
    try:
        df = pd.read_excel(io.BytesIO(raw), dtype=str).fillna('')
    except Exception as e:
        raise ValueError(f'Could not read file: {e}')

    emp_rows = db.execute(
        'SELECT emp_code, name FROM employees WHERE plant_id=? AND is_active=1', (plant_id,)).fetchall()
    emp_map = {r['emp_code']: r['name'] for r in emp_rows}

    field_keywords = {
        'emp_code':       ['emp code','employee code','empcode','staff code','employee id'],
        'programme_name': ['programme name','program name','training name','course name','training need'],
        'prog_type':      ['type of programme','programme type','training type','type'],
        'mode':           ['mode','training mode','delivery mode'],
        'hours':          ['planned hours','hours','hrs','duration'],
    }
    col_map = {}
    for col in df.columns:
        cl = str(col).strip().lower()
        if cl in _MSFORMS_SKIP_HEADERS:
            continue
        for field, kws in field_keywords.items():
            if field not in col_map:
                for kw in kws:
                    if kw in cl or cl in kw:
                        col_map[field] = col
                        break

    inserted, errors = 0, []
    for i, row in df.iterrows():
        def gv(field):
            c = col_map.get(field)
            v = str(row.get(c, '') or '').strip() if c else ''
            return '' if v.lower() in ('nan','none') else v

        emp_code  = gv('emp_code')
        prog_name = gv('programme_name')
        prog_type = gv('prog_type')
        mode      = gv('mode')
        hours     = _safe_float(gv('hours')) or 0.0

        if not emp_code or not prog_name:
            errors.append(f'Row {i+2}: Employee Code and Programme Name are required.')
            continue
        if emp_code not in emp_map:
            errors.append(f'Row {i+2}: Employee code "{emp_code}" not found in your plant.')
            continue
        db.execute(
            'INSERT OR IGNORE INTO tni(plant_id,emp_code,programme_name,prog_type,mode,planned_hours) VALUES(?,?,?,?,?,?)',
            (plant_id, emp_code, prog_name, prog_type, mode, hours))
        inserted += 1

    db.commit()
    return inserted, errors


# ── Smart TNI Analyzer ────────────────────────────────────────────────────────

_ACRONYMS = {
    'PPE','SOP','EHS','OJT','DCS','UPS','VFD','DG','SLD','AC','DC','GST','ISO',
    'HR','IT','MBC','FFT','MIST','DM','ETP','CPU','CGCB','MSDS','OFSAM','ZFD',
    'STD2SD','FCS','RTD','TC','KNO3','MOP','PDM','5S','5-S','JCB','PM','R&M',
    'AI','ML','KPI','GMP','BOD','COD','TOC','TDS','ROI','MIS','SAP','ERR','ERB',
    'CCTV','GPS','QR','LED','LCD','CRM','ERP','LMS','HRM','WMS','PLC','SCADA',
}

_WORD_FIXES = {
    'techqnique':'Technique','tecnique':'Technique','techique':'Technique',
    'technqiue':'Technique','teqnique':'Technique','technque':'Technique',
    'grainding':'Grinding','graining':'Grinding','granding':'Grinding',
    'grindig':'Grinding','grindding':'Grinding',
    'hyigene':'Hygiene','hyegiene':'Hygiene','hygeine':'Hygiene','higiene':'Hygiene',
    'maitenance':'Maintenance','maintenace':'Maintenance','maintainance':'Maintenance',
    'mantenance':'Maintenance',
    'safty':'Safety','saftey':'Safety',
    'operartion':'Operation','operetion':'Operation','opertion':'Operation',
    'managment':'Management','managament':'Management','mangement':'Management',
    'trainning':'Training','traning':'Training',
    'awarness':'Awareness','awreness':'Awareness',
    'handeling':'Handling','handlng':'Handling',
    'electrcial':'Electrical','eletrical':'Electrical',
    'chemcial':'Chemical','chemicle':'Chemical',
    'equpiment':'Equipment','equipement':'Equipment','equipmnet':'Equipment',
    'proceudre':'Procedure','proceedure':'Procedure',
    'complience':'Compliance','compliace':'Compliance',
    'enviroment':'Environment','enviromental':'Environmental',
    'knowlege':'Knowledge','knwoledge':'Knowledge','knoweldge':'Knowledge',
    'monitorng':'Monitoring','monitering':'Monitoring',
    'buidling':'Building','buldling':'Building',
    'confind':'Confined','confinde':'Confined','condfined':'Confined',
    'chocking':'Choking','chocing':'Choking',
    'equipments':'Equipment',
    'lubricants':'Lubricants',
}

_STRIP_CHARS = '.,;:()/'


def _smart_title(s):
    _SMALL = frozenset({'a','an','the','and','or','but','nor','for','yet','so',
                        'at','by','in','of','on','to','as','is','it',
                        'with','from','into','onto','off','per','via'})

    def _tw(w, is_first):
        if not w:
            return w
        if '/' in w:
            parts = w.split('/')
            return '/'.join(_tw(p, is_first and j == 0) for j, p in enumerate(parts))
        lstripped = w.lstrip(_STRIP_CHARS)
        prefix    = w[:len(w) - len(lstripped)]
        core      = lstripped.rstrip(_STRIP_CHARS)
        suffix    = lstripped[len(core):]
        if not core:
            return w
        core_up  = core.upper()
        core_low = core.lower()
        if core_up in _ACRONYMS:
            return prefix + core_up + suffix
        if core == core_up and len(core) >= 2 and core.isalpha():
            return prefix + core_up + suffix
        if core_low == 'ph':
            return prefix + 'pH' + suffix
        if not is_first and core_low in _SMALL and not prefix and not suffix:
            return core_low
        return prefix + core.capitalize() + suffix

    result = ' '.join(_tw(w, i == 0) for i, w in enumerate(s.split()))
    return result.strip('.,;: ')


def _apply_word_fixes(s):
    if not s:
        return s
    from difflib import get_close_matches as _gcm
    words = s.split()
    out   = []
    for w in words:
        lstripped = w.lstrip(_STRIP_CHARS)
        prefix    = w[:len(w) - len(lstripped)]
        core      = lstripped.rstrip(_STRIP_CHARS)
        suffix    = lstripped[len(core):]
        core_low  = core.lower()

        if not core or len(core) < 4 or core.upper() in _ACRONYMS:
            out.append(w)
            continue

        fix = _WORD_FIXES.get(core_low)
        if fix:
            out.append(prefix + fix + suffix)
            continue

        if core_low in _MASTER_VOCAB:
            out.append(w)
            continue

        if _MASTER_VOCAB:
            m = _gcm(core_low, _MASTER_VOCAB.keys(), n=1, cutoff=0.82)
            if m:
                out.append(prefix + _MASTER_VOCAB[m[0]] + suffix)
                continue

        out.append(w)
    return ' '.join(out)


def _canonical_prog(raw_name, plant_id, db):
    if not raw_name or not raw_name.strip():
        return raw_name
    from difflib import get_close_matches as gcm
    master = [r[0] for r in db.execute(
        'SELECT name FROM programme_master WHERE plant_id=? ORDER BY name', (plant_id,)
    ).fetchall()] or []
    master_lower = [m.lower() for m in master]
    corrected = _apply_word_fixes(raw_name.strip())
    raw_lower = corrected.lower()
    if raw_lower in master_lower:
        return master[master_lower.index(raw_lower)]
    m = gcm(raw_lower, master_lower, n=1, cutoff=0.82)
    if m:
        return master[master_lower.index(m[0])]
    return _smart_title(corrected)


def _fuzzy_fix(val, valid_list):
    if not val: return '', False
    vl = val.strip().lower()
    for v in valid_list:
        if v.lower() == vl: return v, False
    for v in valid_list:
        if vl in v.lower() or v.lower() in vl: return v, True
    m = get_close_matches(vl, [v.lower() for v in valid_list], n=1, cutoff=0.55)
    if m:
        idx = [v.lower() for v in valid_list].index(m[0])
        return valid_list[idx], True
    return val, False


def _detect_col(columns, keywords):
    for col in columns:
        cl = str(col).strip().lower()
        for kw in keywords:
            if kw in cl or cl in kw:
                return col
    return None


MASTER_PROGRAMMES = [
    "5-S Management","5S Commercial","Advance Practice Pathology In Agriculture",
    "Advanced Excel","Alignment Of Pumps, Fans And Gear Boxes With Motors",
    "An Overview Of Bio-Pesticides And Its Classification",
    "Bagasse Feeding As Per Boiler Requirement","Bagasse Handling",
    "Basic Fire Safety Awareness","Basic Knowledge Of Hardware",
    "Basic Of DCS Maintenance","Basic Of Maintenance And Programming Of Electronic Governor",
    "Behaviour Based Safety","Boiler Operation & Maintenance","Breakdown Handling",
    "Budget Planning","CGCB Guideline","CPU Operation",
    "Cane Quality & Minimise The Cut To Crush Period",
    "Checking Of Tube Cleaning/Tube Choking","Chemical Safety","Communication Skill",
    "Concept Of Fitting Methodology","Condenser Maintenance And Testing",
    "Condition Monitoring & Log Book Maintenance",
    "Condition Monitoring System Of Equipment","Confined Space",
    "Control Of Insect Pest & Diseases","Control Of Maintenance Schedule",
    "DCS Maintenance & Programming","DM Plant Operation & Maintenance",
    "Dismantling And Fitting Of Pump, Gear Boxes And Fans",
    "ETP Operation","Economiser Maintenance","Electrical Safety",
    "Emergency Awareness During Running Plant","Emergency Management",
    "FFT Efficiency","Failure Analysis Of Process Material",
    "Fire Fighting Equipment Technique","Fire Safety","Flow Measurement",
    "GST","General Checking Of Equipment In Running Season",
    "General Safety Awareness","Good Knowledge Of Metals","HR SOP",
    "Handling Of All Testing Equipment","Handling Of Juice And Mud Removal System",
    "Health Monitoring/Condition Monitoring Of Running Equipment",
    "Hot Work (Gas Cutting, Building And Grinding)",
    "How To Collect Samples","How To Handle Emergency Situation During Running Plant",
    "How To Identify Cane Diseases","Hydraulic Testing And Vacuum Trial Of Pan",
    "ISO General Awareness","IT SOP","Implementation Of OFSAM",
    "Importance Of Machine Guarding","Improvement Of Fermentation Efficiency",
    "Improvement Steam Economy","Income Tax",
    "Industrial Hygiene And 5-S Management",
    "Inspections And Testing Of Lifting Tool And Tackles Lifting Operation",
    "Inventory Management","Irrigation Automation","JCB Operation","JCB Pumping",
    "Juice Analysis","Knowledge About Cleaning And Ability To Conduct Schedule Checking Of Engines",
    "Knowledge About Cleaning, Switch Gear Panels And Motors Checklist",
    "Knowledge About Lab Apparatus & Equipment",
    "Knowledge And Application Of Condensate Removal System",
    "Knowledge And Operations Of Turbine","Knowledge Of Boiler Water Treatment",
    "Knowledge Of DCS Hardware And Panel Wiring",
    "Knowledge Of Electrical Equipment (Motor, Transformers, DG), SLD And Electrical Logics",
    "Knowledge Of Field Instruments Like Pressure Transmitters, Temperature Transmitters, RTD/TC, I To P Converters, Control Valves, Loop Testing",
    "Knowledge Of Importance Of Aeration In Melt",
    "Knowledge Of Industrial Lubricants Properties",
    "Knowledge Of Lighting And AC Systems",
    "Knowledge Of MBC/Belt Conveyor Health Monitoring - Gearbox, Chain Condition, Rake Condition, Idlers And Belt",
    "Knowledge Of Maintenance Of Pumps","Knowledge Of Measuring Instruments",
    "Knowledge Of Molasses Brix & Purity",
    "Knowledge Of Operating Parameter And Quality Of Steam And Cooling Water",
    "Knowledge Of Operation Of MIST","Knowledge Of Pumps And Its Parts",
    "Knowledge Of Supersaturation Zones During Pan Boiling",
    "Knowledge Of Three Motion Hydraulic Cane Unloader Operations",
    "Knowledge Of Upgraded Technology Related IT",
    "Knowledge Of Wire Rope Sling Size By Weight And Job Wise",
    "Knowledge Of Working Tools, Tackles And Fasteners",
    "Labour Laws","Leadership Quality","Legal Compliance",
    "Maintain Brix Of Magma","Maintain Temperature From Juice Heater",
    "Maintaining Load Of Machine By Operating The Feed Valve",
    "Maintenance And Programming Of Electronic Governor",
    "Maintenance Effectiveness","Maintenance Of AC & DC Drives",
    "Maintenance Of Electrical Machine And Switch Gears",
    "Maintenance Of Power Turbine",
    "Maintenance Of Safety Valves And Checking For Its Perfection",
    "Maintenance Of UPS & Battery","Manufacturing Process",
    "Massecuite Curing Temperature And Its Impact","Material Handling",
    "Material Handling (Manual And Mechanical)","Measurement Maintenance",
    "Mill Efficiency","Molasses Conditioning Temperature And Brix",
    "Monitoring And Ensuring Proper Operations Of Bagasse Belt Conveyors",
    "New Methodology In Soil Testing","New Varietal Trial","New Wage Code",
    "Nil Safety","Operation & Maintenance Of Boiler",
    "Operation Of Boiler From Cold To Pressurization",
    "Operation Of Turbine Within Controlled Parameters",
    "Operational Behaviour Awareness",
    "Optimum Temperature And pH Adjustment During Defecation",
    "Organic Waste Management","Overhauling And Maintenance Of Centrifugal Machine",
    "Ownership","PPE Awareness Use Inspection And Handling","Personal Effectiveness",
    "Planning Of Different Massecuite Boiling To Control The Material Load",
    "PowerPoint","Premium Potash Fertilisers: SOP Vs KNO3 Vs MOP Vs PDM",
    "Preventive Maintenance",
    "Preventive Maintenance And Condition Monitoring Of Electrical Equipment",
    "Problem Solving","Process Parameter","Process Safety Management",
    "Purification Of Condensate Removal System","Qualitimetry",
    "Quality Communication & Industrial Security","Quality R&M And Operation",
    "Raw Material Analysis","Removal Of Juice And Filter Cake From The System",
    "Repair & Maintenance Of Workshop Machine","Reporting Of Non EHS Lapses",
    "Reporting Of Non Performance Of Chemicals","Road Safety & Defensive Driving",
    "SOP Boiler","SOP Distillation And Fermentation Operation","SOP ETP",
    "SOP Electrical","SOP Instrumentation","SOP Mill House","SOP PM Module",
    "SOP Sales","SOP Store","SOP Workshop","STD2SD","Safety Induction",
    "Sample Collection","Sample Collection Methods As Per SOP",
    "Sampling & Its Importance","Screen Checking And Molasses Purity Control",
    "Self Discipline","Start And Stop The Boiler","Switch Gear Maintenance & Testing",
    "Team Work","Theft Prevention","To Control Fermentation Process Parameter",
    "To Maintain Brix Of Massecuite Of Dropping Pan",
    "To Maintain Brix Of Syrup At Fix",
    "To Maintain Chemical Dosing As Per Requirement",
    "To Maintain Temperature & pH Of Juice",
    "To Maintain Temperature & pH Of Juice/Melt",
    "To Maintain Temperature And pH Of Juice","Treated Water Parameters",
    "Understanding Of DCS Logic And Graphics","Use Of MSDS",
    "Use Of Optimum Dose Of Flocculant","VFD Maintenance","VFD Operation & Maintenance",
    "Water Management","Withdrawal Of Scum From FCS Clarifier","ZFD",
]
_MASTER_LOWER = [p.lower() for p in MASTER_PROGRAMMES]


def _build_master_vocab():
    _skip = frozenset({'a','an','the','and','or','but','of','in','to','at','by','as','is','it',
                       'for','on','per','via','from','into','onto','off','with','how','its'})
    vocab = {}
    for prog in MASTER_PROGRAMMES:
        for w in re.split(r'[ /,;:()\-]+', prog):
            core = re.sub(r'[.,;:()/&\-]', '', w).strip()
            if core and len(core) >= 4 and core.upper() not in _ACRONYMS and core.lower() not in _skip:
                vocab[core.lower()] = core
    return vocab


_MASTER_VOCAB = _build_master_vocab()


def _smart_analyze_rows(df, plant_id, db):
    from difflib import get_close_matches as gcm
    _prog_cache = {}

    _active_master       = [r[0] for r in db.execute(
        'SELECT name FROM programme_master WHERE plant_id=? ORDER BY name', (plant_id,))]
    _active_master_lower = [p.lower() for p in _active_master]
    _has_master          = len(_active_master) > 0

    def _match_master(raw_lower):
        if raw_lower in _prog_cache:
            return _prog_cache[raw_lower]
        m = gcm(raw_lower, _active_master_lower, n=1, cutoff=0.65)
        result = _active_master[_active_master_lower.index(m[0])] if m else None
        _prog_cache[raw_lower] = result
        return result

    emp_rows  = db.execute(
        'SELECT emp_code, name FROM employees WHERE plant_id=? AND is_active=1', (plant_id,)).fetchall()
    emp_map   = {r['emp_code']: r['name'] for r in emp_rows}
    emp_upper = {k.upper(): k for k in emp_map}

    cols      = df.columns.tolist()
    col_emp   = _detect_col(cols, ['emp code','employee code','empcode','staff code','emp id','employee id','code'])
    col_prog  = _detect_col(cols, ['programme name','program name','training name','course name','training need','training'])
    col_type  = _detect_col(cols, ['type of programme','type','programme type','prog type','training type','category'])
    col_mode  = _detect_col(cols, ['mode','training mode','delivery mode'])
    col_hrs   = _detect_col(cols, ['planned hours','hours','hrs','duration'])

    if not col_emp and not col_prog:
        col_list = ', '.join(f'"{c}"' for c in cols[:15])
        raise ValueError(
            f'Could not detect Employee Code or Programme Name columns. '
            f'Columns found in file: {col_list}. '
            f'Try using "Skip top rows" if your file has a title row above the headers.'
        )

    def gv(row, col):
        if not col: return ''
        v = str(row.get(col, '') or '').strip()
        return '' if v.lower() in ('nan','none','0','') else v

    results = []
    for i, row in df.iterrows():
        raw_emp  = gv(row, col_emp)
        raw_prog = gv(row, col_prog)
        raw_type = gv(row, col_type)
        raw_mode = gv(row, col_mode)
        raw_hrs  = gv(row, col_hrs)

        if not any([raw_emp, raw_prog, raw_type, raw_mode]):
            continue

        fixes  = []
        issues = []
        status = 'ok'

        clean_emp = raw_emp
        if not raw_emp:
            issues.append('Employee Code is missing')
            status = 'error'
        elif raw_emp in emp_map:
            pass
        elif raw_emp.upper() in emp_upper:
            clean_emp = emp_upper[raw_emp.upper()]
            fixes.append({'field':'Employee Code','original':raw_emp,'fixed':clean_emp,'how':'Capitalisation corrected'})
            if status == 'ok': status = 'fixed'
        else:
            issues.append(f'Employee code "{raw_emp}" not found in this plant')
            status = 'error'

        emp_name = emp_map.get(clean_emp, '')

        clean_prog = raw_prog
        if not raw_prog:
            issues.append('Programme Name is missing')
            status = 'error'
        else:
            word_fixed = _apply_word_fixes(raw_prog.strip())
            raw_lower  = word_fixed.lower()
            best = _match_master(raw_lower)
            if best is not None:
                if best.lower() != raw_prog.strip().lower():
                    fixes.append({'field':'Programme Name','original':raw_prog,'fixed':best,'how':'Matched to master list'})
                    if status == 'ok': status = 'fixed'
                clean_prog = best
            else:
                titled = _smart_title(word_fixed)
                if titled != raw_prog:
                    fixes.append({'field':'Programme Name','original':raw_prog,'fixed':titled,'how':'Spelling/case corrected (not in master list)'})
                    clean_prog = titled
                    if status == 'ok': status = 'fixed'
                else:
                    clean_prog = titled
                if _has_master and status not in ('error',):
                    issues.append(f'"{clean_prog}" not found in Programme Master — verify spelling or add it to master list')
                    if status == 'ok': status = 'warning'

        clean_type, type_changed = _fuzzy_fix(raw_type, PROG_TYPES) if raw_type else ('', False)
        if raw_type and clean_type not in PROG_TYPES:
            issues.append(f'Unknown programme type: "{raw_type}" — could not auto-fix')
            if status == 'ok': status = 'error'
        elif raw_type and type_changed:
            fixes.append({'field':'Type of Programme','original':raw_type,'fixed':clean_type,'how':'Auto-matched to standard value'})
            if status == 'ok': status = 'fixed'

        clean_mode, mode_changed = _fuzzy_fix(raw_mode, MODES) if raw_mode else ('', False)
        if raw_mode and clean_mode not in MODES:
            issues.append(f'Unknown mode: "{raw_mode}" — could not auto-fix')
            if status == 'ok': status = 'error'
        elif raw_mode and mode_changed:
            fixes.append({'field':'Mode','original':raw_mode,'fixed':clean_mode,'how':'Auto-matched to standard value'})
            if status == 'ok': status = 'fixed'

        hours = _safe_float(raw_hrs) or 0.0

        results.append({
            'row_num':        i + 2,
            'status':         status,
            'fixes':          fixes,
            'issues':         issues,
            'emp_code':       clean_emp,
            'emp_name':       emp_name,
            'programme_name': clean_prog,
            'prog_type':      clean_type or raw_type,
            'mode':           clean_mode or raw_mode,
            'planned_hours':  hours,
        })
    return results


def _ai_validate_programme_names(uploaded_names, master_progs):
    """
    Second-pass AI check on programme names using Google Gemini (free tier).
    Catches non-canonical suffixes (year, FY, batch, unit codes) and
    semantic duplicates that fuzzy matching misses.

    Returns {name_lower: [{'type': str, 'msg': str, 'fix': str|None}]}.
    Returns {} silently if GEMINI_API_KEY not set or call fails.
    """
    import os
    if not os.environ.get('GEMINI_API_KEY') or not uploaded_names:
        return {}
    try:
        import google.generativeai as genai, json as _j
        genai.configure(api_key=os.environ['GEMINI_API_KEY'])
        model = genai.GenerativeModel('gemini-1.5-flash')

        master_text = '\n'.join(f'- {p}' for p in master_progs[:100]) or '(none yet)'
        names_text  = '\n'.join(f'{i+1}. "{n}"' for i, n in enumerate(uploaded_names))

        prompt = f"""You validate TNI (Training Needs Identification) data for BCML (Balrampur Chini Mills), an Indian sugar manufacturer. SPOCs upload Excel files with employee training plans.

Master list of canonical programme names:
{master_text}

Programme names from this upload:
{names_text}

Check ONLY these two issues:

1. SUFFIX — Name contains a non-canonical suffix that should be stripped: year numbers (2025, 2026), FY codes (FY25-26, 25-26), batch numbers (Batch 1, B2), plant unit codes (BCM, GCM, RCM, TCM, MZP, ACM, KCM, BBN, HCM, MCM), quarters (Q1, Q2, Q3, Q4), or date ranges (Jan-Mar). Suggest the clean canonical name.
   Examples:
   - "Fire Safety Training FY25-26" → fix: "Fire Safety Training"
   - "5S Housekeeping BCM Batch 2" → fix: "5S Housekeeping"
   - "Leadership Dev Workshop Q3 2025" → fix: "Leadership Development Workshop"

2. SEMANTIC_DUP — Two names in THIS list clearly refer to the same programme (>90% confident). Add dup_with as a list of the other 1-based index numbers.
   Examples:
   - "Fire Safety" and "Fire Safety Training" — probable dup
   - "5S Housekeeping" and "5-S House Keeping" — clear dup

Rules:
- Be conservative — only flag if clearly wrong
- Short clean names like "First Aid", "Fire Safety", "POSH Awareness" are fine
- Do NOT flag names just because they are not in master — that is handled elsewhere
- Indian industrial terms are valid: Boiler, Turbine, ETP, Cane, SOP, OJT

Respond with ONLY a compact JSON array, no markdown fences, no explanation:
[{{"idx":1,"issues":[]}},{{"idx":2,"issues":[{{"type":"suffix","msg":"Contains FY code 25-26","fix":"Fire Safety Training"}}]}}]"""

        resp = model.generate_content(prompt)
        raw  = resp.text.strip()
        if raw.startswith('```'):
            raw = raw.split('\n', 1)[1].rsplit('```', 1)[0].strip()

        data = _j.loads(raw)
        findings = {}
        for item in data:
            idx    = item.get('idx', 0) - 1
            issues = item.get('issues') or []
            if issues and 0 <= idx < len(uploaded_names):
                findings[uploaded_names[idx].lower()] = issues
        return findings
    except Exception:
        return {}


def _error_excel_for_tni(error_rows, dup_rows=None, plant_id=None, db=None):
    from openpyxl.worksheet.datavalidation import DataValidation
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Rows To Fix'

    hdr_fill = PatternFill('solid', fgColor='7F1D1D')
    hdr_font = Font(bold=True, color='FFFFFF', size=11)
    headers  = ['Row #','Employee Code','Programme Name','Type of Programme',
                'Mode','Planned Hours','Issue(s) Found']
    col_w    = [7, 16, 34, 22, 14, 14, 60]
    for ci, (h, w) in enumerate(zip(headers, col_w), 1):
        c = ws.cell(row=1, column=ci, value=h)
        c.fill = hdr_fill; c.font = hdr_font
        c.alignment = Alignment(horizontal='center', vertical='center')
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.row_dimensions[1].height = 24
    ws.freeze_panes = 'A2'

    err_fill = PatternFill('solid', fgColor='FEF2F2')
    for ri, r in enumerate(error_rows, 2):
        vals = [r['row_num'], r['emp_code'], r['programme_name'],
                r['prog_type'], r['mode'], r['planned_hours'],
                '; '.join(r['issues'])]
        for ci, v in enumerate(vals, 1):
            cell = ws.cell(row=ri, column=ci, value=v)
            cell.fill = err_fill
            if ci == 7:
                cell.font = Font(color='DC2626', size=10, italic=True)
                cell.alignment = Alignment(wrap_text=True)

    last_row = len(error_rows) + 1

    if plant_id and db and last_row > 1:
        master_progs = [r[0] for r in db.execute(
            'SELECT name FROM programme_master WHERE plant_id=? ORDER BY name', (plant_id,)
        ).fetchall()] or []
        ws_pl = wb.create_sheet('_ProgList')
        ws_pl.sheet_state = 'hidden'
        for idx, v in enumerate(master_progs, 1):
            ws_pl.cell(row=idx, column=1, value=v)
        dv_prog = DataValidation(type='list',
                                 formula1=f'_ProgList!$A$1:$A${len(master_progs)}',
                                 allow_blank=True, showDropDown=False)
        dv_prog.sqref = f'C2:C{last_row}'
        ws.add_data_validation(dv_prog)

    if last_row > 1:
        dv_type = DataValidation(type='list', formula1=f'"{",".join(PROG_TYPES)}"', allow_blank=True)
        dv_mode = DataValidation(type='list', formula1=f'"{",".join(MODES)}"',      allow_blank=True)
        dv_type.sqref = f'D2:D{last_row}'
        dv_mode.sqref = f'E2:E{last_row}'
        ws.add_data_validation(dv_type)
        ws.add_data_validation(dv_mode)

    ws2 = wb.create_sheet('How To Fix')
    tips = [
        'These rows could NOT be auto-fixed. Please correct column by column:',
        '',
        'Employee Code — must exactly match the code in TMS employee master',
        'Programme Name — use the dropdown (click the cell) to pick from the master list',
        'Type of Programme — use the dropdown (click the cell)',
        'Mode — use the dropdown',
        '',
        'After fixing, save and re-upload on the TNI Analyzer page.',
        'Column G (Issues) shows exactly what was wrong — read it before fixing.',
    ]
    for ri, t in enumerate(tips, 1):
        c = ws2.cell(row=ri, column=1, value=t)
        if ri == 1: c.font = Font(bold=True, color='7F1D1D', size=12)
        ws2.column_dimensions['A'].width = 80

    if dup_rows:
        ws3 = wb.create_sheet('Duplicates')
        dup_hdr_fill = PatternFill('solid', fgColor='92400E')
        dup_headers  = ['Row #', 'Employee Code', 'Employee Name', 'Programme Name',
                        'Type', 'Mode', 'Duplicate Type']
        dup_col_w    = [7, 16, 28, 34, 22, 14, 60]
        for ci, (h, w) in enumerate(zip(dup_headers, dup_col_w), 1):
            c = ws3.cell(row=1, column=ci, value=h)
            c.fill = dup_hdr_fill
            c.font = Font(bold=True, color='FFFFFF', size=11)
            c.alignment = Alignment(horizontal='center', vertical='center')
            ws3.column_dimensions[get_column_letter(ci)].width = w
        ws3.row_dimensions[1].height = 24
        ws3.freeze_panes = 'A2'
        dup_fill = PatternFill('solid', fgColor='FFFBEB')
        for ri, r in enumerate(dup_rows, 2):
            vals = [r.get('row_num', ''), r['emp_code'], r.get('emp_name', ''),
                    r['programme_name'], r.get('prog_type', ''), r.get('mode', ''),
                    r['dup_type']]
            for ci, v in enumerate(vals, 1):
                cell = ws3.cell(row=ri, column=ci, value=v)
                cell.fill = dup_fill
                if ci == 7:
                    cell.font = Font(color='92400E', size=10, italic=True)
        ws3.cell(row=len(dup_rows)+3, column=1,
                 value='These rows were NOT imported. Fix or confirm before re-uploading.').font = \
            Font(bold=True, color='92400E')

    buf = io.BytesIO()
    wb.save(buf); buf.seek(0)
    return buf
