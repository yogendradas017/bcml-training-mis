"""Yearly Training Planner — SPOC plans the full FY upfront, locks per-month
or all-at-once. Adherence (Plan vs Actual) computed later from emp_training.

Design:
- Single page /planner shows 12-month grid (Apr-Mar) × programmes from TNI.
- Cell = number of sessions planned for that programme in that month.
- Click cell → drawer to edit pax/hrs/faculty/notes for that programme-month.
- Past + current month auto-locked (cron) — cannot edit without Central override.
- Lock FY = batch-lock all 12 months once committed.
- Coverage projection computed live: if this plan executes, how much of TNI
  demand gets covered per collar (BC/WC).
"""
import json as _json
from datetime import date as _date, datetime as _dt
from flask import render_template, request, redirect, url_for, session, flash, jsonify

from tms.constants import MONTHS_FY, PLANT_MAP
from tms.db import get_db
from tms.decorators import spoc_required, central_required
from tms.audit import log_action


# ── Helpers ─────────────────────────────────────────────────────────────────

# Map FY month label -> YYYY-MM for a given FY start year (April = year, Jan-Mar = year+1)
def _fy_months(fy_start_year):
    """Return [('April', '2026-04'), ('May', '2026-05'), ... ('March', '2027-03')]."""
    out = []
    for i, m in enumerate(MONTHS_FY):
        mnum = (4 + i) if (4 + i) <= 12 else (4 + i - 12)
        yr = fy_start_year if (4 + i) <= 12 else fy_start_year + 1
        out.append((m, f'{yr}-{mnum:02d}'))
    return out


def _current_fy_year():
    """Returns the FY-start year (April = year, Jan-Mar = previous year)."""
    today = _date.today()
    return today.year if today.month >= 4 else today.year - 1


def _is_month_in_past_or_current(yyyymm):
    """True if the given YYYY-MM is the current month or earlier (=> auto-locked)."""
    today = _date.today()
    try:
        yr, mo = int(yyyymm[:4]), int(yyyymm[5:7])
    except (ValueError, IndexError):
        return False
    return (yr, mo) <= (today.year, today.month)


def _audit(db, plant_id, plan_month, actor, action, detail=''):
    try:
        db.execute(
            'INSERT INTO planner_audit(plant_id, plan_month, actor, action, detail)'
            ' VALUES(?,?,?,?,?)',
            (plant_id, plan_month, actor, action, str(detail)[:500])
        )
        db.commit()
    except Exception:
        pass


def _audience_from_tni(db, plant_id, prog_name):
    """Auto-derive audience from TNI nominations for this programme."""
    rows = db.execute(
        '''SELECT DISTINCT e.collar
           FROM tni t
           JOIN employees e ON e.emp_code=t.emp_code AND e.plant_id=t.plant_id
           WHERE t.plant_id=? AND LOWER(t.programme_name)=LOWER(?)''',
        (plant_id, prog_name)
    ).fetchall()
    collars = {r[0] for r in rows if r[0]}
    if not collars:
        return ''
    if collars == {'Blue Collared'}:  return 'Blue Collared'
    if collars == {'White Collared'}: return 'White Collared'
    return 'Common'


def _tni_demand(db, plant_id, fy_year):
    """Per-programme TNI demand: how many people nominated this FY."""
    rows = db.execute(
        '''SELECT programme_name, COUNT(DISTINCT emp_code) AS nominated
           FROM tni
           WHERE plant_id=? AND fy_year=?
           GROUP BY programme_name
           ORDER BY nominated DESC, programme_name''',
        (plant_id, fy_year)
    ).fetchall()
    return [{'programme_name': r['programme_name'], 'nominated': r['nominated']} for r in rows]


def _coverage_projection(db, plant_id, fy_year, plan_rows):
    """Project BC/WC coverage % if this plan executes. plan_rows = list of dicts."""
    bc_tot = db.execute(
        "SELECT COUNT(*) FROM employees WHERE plant_id=? AND is_active=1 AND collar='Blue Collared'",
        (plant_id,)).fetchone()[0]
    wc_tot = db.execute(
        "SELECT COUNT(*) FROM employees WHERE plant_id=? AND is_active=1 AND collar='White Collared'",
        (plant_id,)).fetchone()[0]
    bc_pax = sum(
        int(r.get('target_sessions') or 0) * int(r.get('pax_per_session') or 0)
        for r in plan_rows if r.get('audience') in ('Blue Collared', 'Common')
    )
    wc_pax = sum(
        int(r.get('target_sessions') or 0) * int(r.get('pax_per_session') or 0)
        for r in plan_rows if r.get('audience') in ('White Collared', 'Common')
    )
    bc_pct = round((bc_pax / bc_tot) * 100) if bc_tot else 0
    wc_pct = round((wc_pax / wc_tot) * 100) if wc_tot else 0
    return {
        'bc_pct': min(bc_pct, 100), 'wc_pct': min(wc_pct, 100),
        'bc_pax': bc_pax, 'wc_pax': wc_pax,
        'bc_total_emp': bc_tot, 'wc_total_emp': wc_tot,
    }


# ── Routes ─────────────────────────────────────────────────────────────────

def _register(app):

    @app.route('/planner')
    @spoc_required
    def planner():
        plant_id = session['plant_id']
        db = get_db()
        fy_start = _current_fy_year()
        # Match the short FY format used elsewhere in TMS (TNI, calendar, etc.)
        fy_label = f'{str(fy_start)[2:]}-{str(fy_start + 1)[2:]}'

        # Allow FY override via ?fy=2025 etc
        try:
            fy_param = int(request.args.get('fy', '') or 0)
            if fy_param > 2020:
                fy_start = fy_param
                fy_label = f'{str(fy_start)[2:]}-{str(fy_start + 1)[2:]}'
        except ValueError:
            pass

        months = _fy_months(fy_start)  # [(label, '2026-04'), ...]
        month_labels = [m[0] for m in months]
        month_keys   = [m[1] for m in months]
        today_yyyymm = _date.today().strftime('%Y-%m')

        # Pull existing plan rows for this FY
        rows = db.execute(
            '''SELECT * FROM planner_entries
               WHERE plant_id=? AND fy_year=?
               ORDER BY programme_name, plan_month''',
            (plant_id, fy_label)
        ).fetchall()
        plan_rows = [dict(r) for r in rows]

        # Build matrix: programme_name -> {month_key: row_dict}
        matrix = {}
        for r in plan_rows:
            matrix.setdefault(r['programme_name'], {})[r['plan_month']] = r

        # TNI demand (becomes the source of programme list)
        tni = _tni_demand(db, plant_id, fy_label)
        tni_by_prog = {t['programme_name']: t['nominated'] for t in tni}

        # Compute "planned vs nominated" gap per programme
        prog_summary = []
        # Use union of TNI programmes + any added in planner
        all_progs = list(dict.fromkeys(
            [t['programme_name'] for t in tni] + list(matrix.keys())
        ))
        # Pull canonical hours-per-session from TNI (most common value used per programme)
        tni_hrs = {r[0]: r[1] for r in db.execute(
            '''SELECT programme_name, planned_hours
               FROM tni
               WHERE plant_id=? AND fy_year=? AND planned_hours > 0
               GROUP BY programme_name
               ORDER BY COUNT(*) DESC''',
            (plant_id, fy_label)
        ).fetchall()}

        for pn in all_progs:
            row_map = matrix.get(pn, {})
            planned_sessions = sum(int(r.get('target_sessions') or 0) for r in row_map.values())
            planned_pax = sum(int(r.get('target_sessions') or 0) * int(r.get('pax_per_session') or 0)
                              for r in row_map.values())
            nominated = tni_by_prog.get(pn, 0)
            over_planned = (nominated > 0 and planned_pax > nominated)
            over_by_pct = round((planned_pax - nominated) / nominated * 100) if over_planned else 0
            if nominated == 0:
                tag = 'new'      # New Requirement (not in TNI)
            elif over_planned:
                tag = 'over'     # Over-planned vs TNI demand
            elif planned_pax == 0:
                tag = 'critical'
            elif planned_pax < nominated * 0.5:
                tag = 'under'
            elif planned_pax >= nominated:
                tag = 'ontrack'
            else:
                tag = 'partial'
            prog_summary.append({
                'programme_name': pn,
                'nominated': nominated,
                'planned_sessions': planned_sessions,
                'planned_pax': planned_pax,
                'over_planned': over_planned,
                'over_by_pct': over_by_pct,
                'standard_hrs': tni_hrs.get(pn, 4),
                'tag': tag,
            })

        # Sort: over (urgent waste) first, then critical, under, partial, ontrack, new
        tag_order = {'over': 0, 'critical': 1, 'under': 2, 'partial': 3, 'ontrack': 4, 'new': 5}
        prog_summary.sort(key=lambda x: (tag_order.get(x['tag'], 99), -x['nominated']))

        over_planned_count = sum(1 for p in prog_summary if p['over_planned'])

        # Total KPIs
        total_sessions = sum(int(r.get('target_sessions') or 0) for r in plan_rows)
        total_pax = sum(int(r.get('target_sessions') or 0) * int(r.get('pax_per_session') or 0)
                        for r in plan_rows)
        total_hrs = sum(int(r.get('target_sessions') or 0) * float(r.get('hours_per_session') or 0)
                        for r in plan_rows)
        coverage = _coverage_projection(db, plant_id, fy_label, plan_rows)

        # Month lock status
        month_locked = {}
        for mk in month_keys:
            # Auto-locked if past/current OR all rows for this month are locked
            auto_lock = _is_month_in_past_or_current(mk)
            explicit_lock = any(
                (r['plan_month'] == mk and r.get('status') == 'locked')
                for r in plan_rows
            )
            month_locked[mk] = auto_lock or explicit_lock

        months_locked_count = sum(1 for k, v in month_locked.items() if v)

        # JSON-friendly lookup for client-side validation in the drawer
        prog_meta_js = {
            p['programme_name']: {
                'nominated': p['nominated'],
                'planned_pax': p['planned_pax'],
                'planned_sessions': p['planned_sessions'],
                'standard_hrs': p['standard_hrs'],
                'over_planned': p['over_planned'],
                'tag': p['tag'],
            } for p in prog_summary
        }

        return render_template(
            'planner.html',
            fy_label=fy_label,
            fy_start_year=fy_start,
            month_labels=month_labels,
            month_keys=month_keys,
            today_yyyymm=today_yyyymm,
            matrix=matrix,
            prog_summary=prog_summary,
            prog_meta_js=prog_meta_js,
            total_sessions=total_sessions,
            total_pax=total_pax,
            total_hrs=int(total_hrs),
            coverage=coverage,
            over_planned_count=over_planned_count,
            month_locked=month_locked,
            months_locked_count=months_locked_count,
            fy_year_options=[_current_fy_year(), _current_fy_year() - 1, _current_fy_year() + 1],
        )

    @app.route('/planner/save', methods=['POST'])
    @spoc_required
    def planner_save():
        """Save draft plan rows. Body = JSON list of {programme_name, plan_month,
        target_sessions, pax_per_session, hours_per_session, faculty, notes, fy_year}.

        Server validates each row: cannot modify a locked month-row. Auto-derives
        audience from TNI. Upserts via UNIQUE(plant_id, plan_month, programme_name).
        """
        plant_id = session['plant_id']
        actor    = session.get('username', 'unknown')
        db       = get_db()
        try:
            payload = request.get_json(force=True, silent=True) or {}
        except Exception:
            return jsonify({'ok': False, 'error': 'invalid JSON body'}), 400

        rows     = payload.get('rows') or []
        fy_year  = (payload.get('fy_year') or '').strip()
        if not fy_year:
            return jsonify({'ok': False, 'error': 'fy_year required'}), 400

        # Pull current TNI demand + already-planned-other-months for cross-row validation
        tni_by_prog = {
            r['programme_name']: r['nominated']
            for r in [
                {'programme_name': x['programme_name'], 'nominated': x['nominated']}
                for x in _tni_demand(db, plant_id, fy_year)
            ]
        }
        existing_plan = {}
        for r in db.execute(
            'SELECT programme_name, plan_month, target_sessions, pax_per_session FROM planner_entries WHERE plant_id=? AND fy_year=?',
            (plant_id, fy_year)
        ).fetchall():
            existing_plan.setdefault(r['programme_name'], {})[r['plan_month']] = {
                'sessions': r['target_sessions'] or 0,
                'pax':      r['pax_per_session'] or 0,
            }

        saved = 0; skipped_locked = 0; errors = []; warnings = []
        for r in rows:
            try:
                prog   = (r.get('programme_name') or '').strip()
                pmonth = (r.get('plan_month') or '').strip()
                if not prog or not pmonth or len(pmonth) != 7:
                    continue
                # Hard reject impossible values
                ts_raw  = r.get('target_sessions')
                pps_raw = r.get('pax_per_session')
                hps_raw = r.get('hours_per_session')
                try:
                    ts  = int(ts_raw or 0)
                    pps = int(pps_raw or 20)
                    hps = float(hps_raw or 4)
                except (ValueError, TypeError):
                    errors.append(f'{prog}/{pmonth}: non-numeric input')
                    continue
                if ts < 0 or pps < 1 or hps <= 0:
                    errors.append(f'{prog}/{pmonth}: negative or zero values not allowed')
                    continue
                if ts > 50:
                    errors.append(f'{prog}/{pmonth}: max 50 sessions per cell (got {ts})')
                    continue
                if pps > 200:
                    errors.append(f'{prog}/{pmonth}: max 200 pax/session (got {pps})')
                    continue
                if hps > 40:
                    errors.append(f'{prog}/{pmonth}: max 40 hrs/session (got {hps})')
                    continue
                faculty = (r.get('faculty') or '').strip()[:120]
                notes   = (r.get('notes') or '').strip()[:500]

                # Soft validation: warn if total planned pax exceeds TNI nominated
                nominated = tni_by_prog.get(prog, 0)
                if nominated > 0:
                    other_months_pax = sum(
                        info['sessions'] * info['pax']
                        for mk, info in existing_plan.get(prog, {}).items()
                        if mk != pmonth
                    )
                    new_total = other_months_pax + (ts * pps)
                    if new_total > nominated:
                        warnings.append(
                            f'{prog} ({pmonth}): planned {new_total} pax exceeds TNI nominated {nominated} '
                            f'(over by {new_total - nominated})'
                        )

                # Block edits to locked months (unless caller has admin role)
                if (_is_month_in_past_or_current(pmonth)
                        and session.get('role') not in ('admin',)):
                    existing = db.execute(
                        'SELECT id FROM planner_entries WHERE plant_id=? AND plan_month=? AND programme_name=?',
                        (plant_id, pmonth, prog)).fetchone()
                    # Only block if there's existing data to protect
                    if existing or ts > 0:
                        skipped_locked += 1
                        continue

                # Zero sessions = delete the row (clean grid)
                if ts <= 0:
                    db.execute(
                        'DELETE FROM planner_entries WHERE plant_id=? AND plan_month=? AND programme_name=?',
                        (plant_id, pmonth, prog))
                    saved += 1
                    continue

                audience = _audience_from_tni(db, plant_id, prog)
                db.execute('''INSERT INTO planner_entries(
                    plant_id, fy_year, plan_month, programme_name,
                    target_sessions, pax_per_session, hours_per_session,
                    faculty, audience, notes, created_by, updated_at)
                    VALUES(?,?,?,?,?,?,?,?,?,?,?,datetime('now','localtime'))
                    ON CONFLICT(plant_id, plan_month, programme_name) DO UPDATE SET
                        target_sessions   = excluded.target_sessions,
                        pax_per_session   = excluded.pax_per_session,
                        hours_per_session = excluded.hours_per_session,
                        faculty           = excluded.faculty,
                        audience          = excluded.audience,
                        notes             = excluded.notes,
                        updated_at        = datetime('now','localtime')
                    ''',
                    (plant_id, fy_year, pmonth, prog,
                     ts, pps, hps, faculty, audience, notes, actor))
                saved += 1
            except Exception as e:
                errors.append(f'{r.get("programme_name", "?")}/{r.get("plan_month", "?")}: {e}')

        db.commit()
        _audit(db, plant_id, '', actor, 'save_draft',
               f'saved={saved} skipped_locked={skipped_locked} errors={len(errors)}')
        log_action('RECORD_EDIT', f'planner_save:{saved} rows')
        return jsonify({
            'ok': True, 'saved': saved,
            'skipped_locked': skipped_locked,
            'errors': errors,
            'warnings': warnings,
        })

    @app.route('/planner/lock', methods=['POST'])
    @spoc_required
    def planner_lock():
        """Lock either a specific month (?month=YYYY-MM) or whole FY (?fy_year=2026-27)."""
        plant_id = session['plant_id']
        actor    = session.get('username', 'unknown')
        db       = get_db()
        month    = (request.form.get('month') or '').strip()
        fy_year  = (request.form.get('fy_year') or '').strip()
        ack      = request.form.get('acknowledge_gap', '0') == '1'

        if month and len(month) == 7:
            cnt = db.execute(
                '''UPDATE planner_entries
                   SET status='locked', locked_at=datetime('now','localtime'), locked_by=?
                   WHERE plant_id=? AND plan_month=? AND status!='locked' ''',
                (actor, plant_id, month)).rowcount
            db.commit()
            _audit(db, plant_id, month, actor, 'lock_month',
                   f'rows_locked={cnt} ack_gap={ack}')
            log_action('RECORD_EDIT', f'planner_lock_month:{month}:{cnt}')
            flash(f'Locked {cnt} plan row(s) for {month}. Cannot be edited without Central override.',
                  'success')
            return redirect(url_for('planner'))

        if fy_year:
            cnt = db.execute(
                '''UPDATE planner_entries
                   SET status='locked', locked_at=datetime('now','localtime'), locked_by=?
                   WHERE plant_id=? AND fy_year=? AND status!='locked' ''',
                (actor, plant_id, fy_year)).rowcount
            db.commit()
            _audit(db, plant_id, '', actor, 'lock_fy',
                   f'rows_locked={cnt} ack_gap={ack}')
            log_action('RECORD_EDIT', f'planner_lock_fy:{fy_year}:{cnt}')
            flash(f'FY {fy_year} plan LOCKED. {cnt} row(s). Plan is now your committed baseline.',
                  'success')
            return redirect(url_for('planner'))

        flash('Specify either month=YYYY-MM or fy_year=YYYY-YY to lock.', 'danger')
        return redirect(url_for('planner'))

    @app.route('/planner/unlock', methods=['POST'])
    @central_required
    def planner_unlock():
        """Central-only override. Unlocks a specific plant+month for editing."""
        plant_id = request.form.get('plant_id', '').strip()
        month    = request.form.get('month', '').strip()
        reason   = request.form.get('reason', '').strip()
        actor    = session.get('username', 'central')
        if not plant_id.isdigit() or not month or len(reason) < 10:
            flash('Plant + month + reason (10+ chars) required.', 'danger')
            return redirect(url_for('planner'))
        db = get_db()
        cnt = db.execute(
            '''UPDATE planner_entries SET status='draft', locked_at=NULL, locked_by=NULL
               WHERE plant_id=? AND plan_month=? ''',
            (int(plant_id), month)).rowcount
        db.commit()
        _audit(db, int(plant_id), month, actor, 'edit_locked',
               f'central_unlock reason={reason} rows={cnt}')
        log_action('RECORD_EDIT', f'planner_unlock:{plant_id}:{month}:{cnt}')
        flash(f'Unlocked {cnt} row(s) for plant {plant_id} {month}. SPOC can now edit.',
              'warning')
        return redirect(url_for('central_dashboard'))
