from flask import render_template, request, redirect, url_for, session, flash

from tms.constants import PLANTS, PLANT_MAP, MONTHS_FY
from tms.db import get_db
from tms.decorators import central_required
from tms.helpers import _calc_summary, _calc_totals, _calc_compliance, _current_fy


def _register(app):

    @app.route('/central')
    @central_required
    def central_dashboard():
        db = get_db()
        fy_start, fy_end = _current_fy()
        plant_summaries = []
        for p in PLANTS:
            pid = p['id']
            bc  = db.execute("SELECT COUNT(*) FROM employees WHERE plant_id=? AND is_active=1 AND collar='Blue Collared'", (pid,)).fetchone()[0]
            wc  = db.execute("SELECT COUNT(*) FROM employees WHERE plant_id=? AND is_active=1 AND collar='White Collared'", (pid,)).fetchone()[0]
            # Plant's own calendar sessions
            own_sessions  = db.execute("SELECT COUNT(*) FROM calendar WHERE plant_id=?", (pid,)).fetchone()[0]
            own_conducted = db.execute("SELECT COUNT(*) FROM calendar WHERE plant_id=? AND status='Conducted'", (pid,)).fetchone()[0]
            # Distinct central-hosted sessions attended by this plant's employees — FY-scoped
            central_attended = db.execute(
                "SELECT COUNT(DISTINCT session_code) FROM emp_training "
                "WHERE plant_id=? AND host_plant_id=99 AND session_code IS NOT NULL AND session_code!=''"
                " AND start_date>=? AND start_date<=?",
                (pid, fy_start, fy_end)).fetchone()[0]
            sessions_cnt  = own_sessions + central_attended
            conducted_cnt = own_conducted + central_attended
            manhours = db.execute(
                "SELECT COALESCE(SUM(hrs),0) FROM emp_training WHERE plant_id=? AND start_date>=? AND start_date<=?",
                (pid, fy_start, fy_end)).fetchone()[0]
            bc_hrs   = db.execute(
                "SELECT COALESCE(SUM(t.hrs),0) FROM emp_training t JOIN employees e ON e.emp_code=t.emp_code AND e.plant_id=t.plant_id "
                "WHERE t.plant_id=? AND e.collar='Blue Collared' AND t.start_date>=? AND t.start_date<=?",
                (pid, fy_start, fy_end)).fetchone()[0]
            wc_hrs   = db.execute(
                "SELECT COALESCE(SUM(t.hrs),0) FROM emp_training t JOIN employees e ON e.emp_code=t.emp_code AND e.plant_id=t.plant_id "
                "WHERE t.plant_id=? AND e.collar='White Collared' AND t.start_date>=? AND t.start_date<=?",
                (pid, fy_start, fy_end)).fetchone()[0]
            bc_mandate = bc * 12
            wc_mandate = wc * 24
            bc_pct = round((bc_hrs / bc_mandate * 100), 1) if bc_mandate else 0
            wc_pct = round((wc_hrs / wc_mandate * 100), 1) if wc_mandate else 0
            plant_summaries.append({**p,
                'blue_collar': bc, 'white_collar': wc, 'total_emp': bc + wc,
                'sessions': sessions_cnt, 'conducted': conducted_cnt,
                'manhours': round(manhours, 1),
                'bc_pct': bc_pct, 'wc_pct': wc_pct
            })
        grand = {
            'total_emp': sum(p['total_emp'] for p in plant_summaries),
            'manhours':  round(sum(p['manhours'] for p in plant_summaries), 1),
            'sessions':  sum(p['sessions'] for p in plant_summaries),
            'conducted': sum(p['conducted'] for p in plant_summaries),
        }

        # Quarter date ranges: use start_date bounds instead of free-text month field
        Q_RANGES = [
            ('Q1 (Apr–Jun)', fy_start[:4]+'-04-01', fy_start[:4]+'-06-30'),
            ('Q2 (Jul–Sep)', fy_start[:4]+'-07-01', fy_start[:4]+'-09-30'),
            ('Q3 (Oct–Dec)', fy_start[:4]+'-10-01', fy_start[:4]+'-12-31'),
            ('Q4 (Jan–Mar)', str(int(fy_start[:4])+1)+'-01-01', str(int(fy_start[:4])+1)+'-03-31'),
        ]
        quarterly = []
        for qname, q_start, q_end in Q_RANGES:
            sc = db.execute(
                "SELECT COUNT(*) FROM calendar WHERE status='Conducted'"
                " AND plan_start>=? AND plan_start<=?", [q_start, q_end]).fetchone()[0]
            mh = db.execute(
                "SELECT COALESCE(SUM(hrs),0) FROM emp_training WHERE start_date>=? AND start_date<=?",
                [q_start, q_end]).fetchone()[0]
            plant_q = []
            for p in plant_summaries:
                pid  = p['id']
                own_sc_p = db.execute(
                    "SELECT COUNT(*) FROM calendar WHERE plant_id=? AND status='Conducted'"
                    " AND plan_start>=? AND plan_start<=?",
                    [pid, q_start, q_end]).fetchone()[0]
                cen_sc_p = db.execute(
                    "SELECT COUNT(DISTINCT session_code) FROM emp_training "
                    "WHERE plant_id=? AND host_plant_id=99 AND session_code IS NOT NULL AND session_code!=''"
                    " AND start_date>=? AND start_date<=?",
                    [pid, q_start, q_end]).fetchone()[0]
                sc_p = own_sc_p + cen_sc_p
                mh_p = db.execute(
                    "SELECT COALESCE(SUM(hrs),0) FROM emp_training WHERE plant_id=? AND start_date>=? AND start_date<=?",
                    [pid, q_start, q_end]).fetchone()[0]
                plant_q.append({'name': p['name'], 'unit_code': p['unit_code'], 'id': p['id'],
                                'sessions': sc_p, 'manhours': round(mh_p, 1)})
            quarterly.append({'quarter': qname, 'sessions': sc, 'manhours': round(mh, 1), 'plants': plant_q})

        return render_template('central.html', plants=plant_summaries, grand=grand, quarterly=quarterly)

    @app.route('/central/plant/<int:plant_id>')
    @central_required
    def central_plant_view(plant_id):
        if plant_id not in PLANT_MAP:
            flash('Plant not found.', 'danger')
            return redirect(url_for('central_dashboard'))
        plant        = PLANT_MAP[plant_id]
        db           = get_db()
        sel_month    = request.args.get('month', '')
        summary_rows = _calc_summary(plant_id, sel_month, db)
        totals       = _calc_totals(summary_rows, db=db, plant_id=plant_id)
        compliance   = _calc_compliance(plant_id, db)
        return render_template('central_plant.html', plant=plant,
                               summary_rows=summary_rows, totals=totals,
                               compliance=compliance, months=MONTHS_FY,
                               selected_month=sel_month)
