from flask import render_template, request, redirect, url_for, session, flash

from tms.constants import PLANTS, PLANT_MAP, MONTHS_FY
from tms.db import get_db
from tms.decorators import central_required
from tms.helpers import _calc_summary, _calc_totals, _calc_compliance, _current_fy


def _by_plant(rows, key='plant_id', val='cnt'):
    return {r[key]: r[val] for r in rows}


def _register(app):

    @app.route('/central')
    @central_required
    def central_dashboard():
        db = get_db()
        fy_start, fy_end = _current_fy()

        # ── 7 batched queries replace ~80 per-plant queries ─────────────────
        hc = {r['plant_id']: {'bc': r['bc'], 'wc': r['wc']} for r in db.execute(
            "SELECT plant_id, "
            "  SUM(CASE WHEN collar='Blue Collared'  THEN 1 ELSE 0 END) AS bc, "
            "  SUM(CASE WHEN collar='White Collared' THEN 1 ELSE 0 END) AS wc "
            "FROM employees WHERE is_active=1 GROUP BY plant_id"
        ).fetchall()}
        own_planned = _by_plant(db.execute(
            "SELECT plant_id, COUNT(*) AS cnt FROM calendar "
            "WHERE plan_start BETWEEN ? AND ? GROUP BY plant_id",
            (fy_start, fy_end)).fetchall())
        own_cond = _by_plant(db.execute(
            "SELECT plant_id, COUNT(*) AS cnt FROM calendar "
            "WHERE status='Conducted' AND plan_start BETWEEN ? AND ? GROUP BY plant_id",
            (fy_start, fy_end)).fetchall())
        cen_att = _by_plant(db.execute(
            "SELECT plant_id, COUNT(DISTINCT session_code) AS cnt FROM emp_training "
            "WHERE host_plant_id=99 AND session_code IS NOT NULL AND session_code!='' "
            "AND start_date>=? AND start_date<=? GROUP BY plant_id",
            (fy_start, fy_end)).fetchall())
        mh_all = _by_plant(db.execute(
            "SELECT plant_id, COALESCE(SUM(hrs),0) AS cnt FROM emp_training "
            "WHERE start_date>=? AND start_date<=? GROUP BY plant_id",
            (fy_start, fy_end)).fetchall())
        mh_bc = _by_plant(db.execute(
            "SELECT t.plant_id, COALESCE(SUM(t.hrs),0) AS cnt "
            "FROM emp_training t JOIN employees e "
            "  ON e.emp_code=t.emp_code AND e.plant_id=t.plant_id "
            "WHERE e.collar='Blue Collared' AND t.start_date>=? AND t.start_date<=? "
            "GROUP BY t.plant_id", (fy_start, fy_end)).fetchall())
        mh_wc = _by_plant(db.execute(
            "SELECT t.plant_id, COALESCE(SUM(t.hrs),0) AS cnt "
            "FROM emp_training t JOIN employees e "
            "  ON e.emp_code=t.emp_code AND e.plant_id=t.plant_id "
            "WHERE e.collar='White Collared' AND t.start_date>=? AND t.start_date<=? "
            "GROUP BY t.plant_id", (fy_start, fy_end)).fetchall())

        plant_summaries = []
        for p in PLANTS:
            pid = p['id']
            bc  = hc.get(pid, {}).get('bc', 0)
            wc  = hc.get(pid, {}).get('wc', 0)
            own_s = own_planned.get(pid, 0)
            own_c = own_cond.get(pid, 0)
            cen_c = cen_att.get(pid, 0)
            manhours = mh_all.get(pid, 0)
            bc_hrs   = mh_bc.get(pid, 0)
            wc_hrs   = mh_wc.get(pid, 0)
            bc_mandate = bc * 12
            wc_mandate = wc * 24
            bc_pct = round(bc_hrs / bc_mandate * 100, 1) if bc_mandate else 0
            wc_pct = round(wc_hrs / wc_mandate * 100, 1) if wc_mandate else 0
            plant_summaries.append({**p,
                'blue_collar': bc, 'white_collar': wc, 'total_emp': bc + wc,
                'sessions':  own_s + cen_c, 'conducted': own_c + cen_c,
                'own_sessions': own_s, 'own_conducted': own_c,
                'manhours': round(manhours, 1),
                'bc_pct': bc_pct, 'wc_pct': wc_pct,
            })

        plant_summaries.sort(key=lambda p: (p['bc_pct'] + p['wc_pct']) / 2, reverse=True)

        grand_central = db.execute(
            "SELECT COUNT(DISTINCT session_code) FROM emp_training "
            "WHERE host_plant_id=99 AND session_code IS NOT NULL AND session_code!='' "
            "AND start_date>=? AND start_date<=?",
            (fy_start, fy_end)).fetchone()[0]
        grand = {
            'total_emp': sum(p['total_emp'] for p in plant_summaries),
            'manhours':  round(sum(p['manhours'] for p in plant_summaries), 1),
            'sessions':  sum(p['own_sessions'] for p in plant_summaries) + grand_central,
            'conducted': sum(p['own_conducted'] for p in plant_summaries) + grand_central,
        }

        # ── Quarterly — 6 batched queries per quarter (was 50 per quarter) ──
        fy_yr = fy_start[:4]
        Q_RANGES = [
            ('Q1 (Apr–Jun)', f'{fy_yr}-04-01',         f'{fy_yr}-06-30'),
            ('Q2 (Jul–Sep)', f'{fy_yr}-07-01',         f'{fy_yr}-09-30'),
            ('Q3 (Oct–Dec)', f'{fy_yr}-10-01',         f'{fy_yr}-12-31'),
            ('Q4 (Jan–Mar)', f'{int(fy_yr)+1}-01-01',  f'{int(fy_yr)+1}-03-31'),
        ]
        quarterly = []
        for qname, q_start, q_end in Q_RANGES:
            sc = db.execute(
                "SELECT COUNT(*) FROM calendar WHERE status='Conducted' "
                "AND plan_start>=? AND plan_start<=?", (q_start, q_end)).fetchone()[0]
            mh = db.execute(
                "SELECT COALESCE(SUM(hrs),0) FROM emp_training "
                "WHERE start_date>=? AND start_date<=?", (q_start, q_end)).fetchone()[0]
            q_planned = _by_plant(db.execute(
                "SELECT plant_id, COUNT(*) AS cnt FROM calendar "
                "WHERE plan_start>=? AND plan_start<=? GROUP BY plant_id",
                (q_start, q_end)).fetchall())
            q_own_sc = _by_plant(db.execute(
                "SELECT plant_id, COUNT(*) AS cnt FROM calendar "
                "WHERE status='Conducted' AND plan_start>=? AND plan_start<=? GROUP BY plant_id",
                (q_start, q_end)).fetchall())
            q_cen_sc = _by_plant(db.execute(
                "SELECT plant_id, COUNT(DISTINCT session_code) AS cnt FROM emp_training "
                "WHERE host_plant_id=99 AND session_code IS NOT NULL AND session_code!='' "
                "AND start_date>=? AND start_date<=? GROUP BY plant_id",
                (q_start, q_end)).fetchall())
            q_mh = _by_plant(db.execute(
                "SELECT plant_id, COALESCE(SUM(hrs),0) AS cnt FROM emp_training "
                "WHERE start_date>=? AND start_date<=? GROUP BY plant_id",
                (q_start, q_end)).fetchall())
            q_bc = _by_plant(db.execute(
                "SELECT t.plant_id, COALESCE(SUM(t.hrs),0) AS cnt "
                "FROM emp_training t JOIN employees e "
                "  ON e.emp_code=t.emp_code AND e.plant_id=t.plant_id "
                "WHERE e.collar='Blue Collared' AND t.start_date>=? AND t.start_date<=? "
                "GROUP BY t.plant_id", (q_start, q_end)).fetchall())
            q_wc = _by_plant(db.execute(
                "SELECT t.plant_id, COALESCE(SUM(t.hrs),0) AS cnt "
                "FROM emp_training t JOIN employees e "
                "  ON e.emp_code=t.emp_code AND e.plant_id=t.plant_id "
                "WHERE e.collar='White Collared' AND t.start_date>=? AND t.start_date<=? "
                "GROUP BY t.plant_id", (q_start, q_end)).fetchall())

            plant_q = [{
                'name': p['name'], 'unit_code': p['unit_code'], 'id': p['id'],
                'planned':  q_planned.get(p['id'], 0),
                'sessions': q_own_sc.get(p['id'], 0) + q_cen_sc.get(p['id'], 0),
                'manhours': round(q_mh.get(p['id'], 0), 1),
                'bc_hrs':   round(q_bc.get(p['id'], 0), 1),
                'wc_hrs':   round(q_wc.get(p['id'], 0), 1),
            } for p in plant_summaries]

            quarterly.append({
                'quarter': qname, 'sessions': sc, 'manhours': round(mh, 1),
                'q_start': q_start, 'q_end': q_end,
                'plants': plant_q,
            })

        return render_template('central.html', plants=plant_summaries, grand=grand, quarterly=quarterly)

    @app.route('/central/duplicates', methods=['GET', 'POST'])
    @central_required
    def central_duplicates():
        """Programme master duplicate scanner + bulk merger.

        GET: scan all plants (or selected) and render clusters for review.
        POST: apply selected merges with cascade rename across tni/calendar/
              programme_details/emp_training.
        """
        from tms.master_dedup import find_duplicates, merge_cluster
        from tms.audit import log_action
        db = get_db()

        if request.method == 'POST':
            # Form fields:
            #   cluster_idx: 0,1,2,...
            #   winner_<idx>: master id of canonical row
            #   canonical_<idx>: free-text canonical name (defaults to winner's name)
            #   losers_<idx>: comma-separated master ids to merge in
            #   plant_<idx>: plant_id for this cluster
            plant_id = int(request.form.get('plant_id') or 0)
            merged_clusters = 0
            total_counts = {'tni': 0, 'calendar': 0, 'programme_details': 0, 'emp_training': 0,
                            'winner_renamed': 0, 'losers_deleted': 0}
            for key in request.form:
                if not key.startswith('winner_'):
                    continue
                idx = key[len('winner_'):]
                winner_id = int(request.form.get(f'winner_{idx}') or 0)
                losers_raw = request.form.get(f'losers_{idx}', '').strip()
                canonical  = request.form.get(f'canonical_{idx}', '').strip()
                cluster_pid = int(request.form.get(f'plant_{idx}') or plant_id or 0)
                if not winner_id or not losers_raw or not canonical or not cluster_pid:
                    continue
                loser_ids = [int(x) for x in losers_raw.split(',') if x.strip().isdigit()]
                if not loser_ids:
                    continue
                counts = merge_cluster(cluster_pid, winner_id, loser_ids, canonical, db,
                                       audit_log_fn=log_action)
                merged_clusters += 1
                for k, v in counts.items():
                    total_counts[k] = total_counts.get(k, 0) + v
            db.commit()
            if merged_clusters:
                flash(
                    f'Merged {merged_clusters} cluster(s). '
                    f'Cascaded: TNI {total_counts["tni"]}, Calendar {total_counts["calendar"]}, '
                    f'2C {total_counts["programme_details"]}, 2A {total_counts["emp_training"]}. '
                    f'Deleted {total_counts["losers_deleted"]} duplicate master rows.',
                    'success')
            else:
                flash('No merges selected.', 'warning')
            return redirect(url_for('central_duplicates'))

        # GET — scan
        plant_filter = request.args.get('plant_id', '').strip()
        try:
            plant_filter_id = int(plant_filter) if plant_filter else None
        except ValueError:
            plant_filter_id = None
        try:
            threshold = float(request.args.get('threshold', '0.85'))
            threshold = max(0.70, min(0.99, threshold))
        except (ValueError, TypeError):
            threshold = 0.85

        plant_clusters = []
        plants_to_scan = [plant_filter_id] if plant_filter_id else [
            p['id'] for p in PLANTS if p['id'] != 99]
        for pid in plants_to_scan:
            if pid not in PLANT_MAP:
                continue
            dupes = find_duplicates(pid, db, threshold=threshold)
            if dupes:
                plant_clusters.append({
                    'plant_id':   pid,
                    'plant_name': PLANT_MAP[pid]['name'],
                    'clusters':   dupes,
                })

        return render_template('central_duplicates.html',
                               plant_clusters=plant_clusters,
                               threshold=threshold,
                               plants=[p for p in PLANTS if p['id'] != 99],
                               plant_filter=plant_filter_id)

    @app.route('/central/tni-errors')
    @central_required
    def central_tni_errors():
        db = get_db()
        # Filters
        plant_filter = request.args.get('plant', '').strip()
        date_from    = request.args.get('from', '').strip()
        date_to      = request.args.get('to', '').strip()
        status_f     = request.args.get('status', '').strip()
        try:
            limit = min(int(request.args.get('limit', '500') or 500), 5000)
        except ValueError:
            limit = 500

        where = ['1=1']
        params = []
        if plant_filter and plant_filter.isdigit():
            where.append('plant_id=?')
            params.append(int(plant_filter))
        if date_from:
            where.append('ts >= ?')
            params.append(date_from + ' 00:00:00')
        if date_to:
            where.append('ts <= ?')
            params.append(date_to + ' 23:59:59')
        if status_f in ('error', 'duplicate'):
            where.append('row_status=?')
            params.append(status_f)

        wsql = ' AND '.join(where)

        # KPIs
        total_errors = db.execute(
            f'SELECT COUNT(*) FROM tni_upload_errors WHERE {wsql}', params).fetchone()[0]
        plants_involved = db.execute(
            f'SELECT COUNT(DISTINCT plant_id) FROM tni_upload_errors WHERE {wsql}', params).fetchone()[0]
        users_involved = db.execute(
            f'SELECT COUNT(DISTINCT username) FROM tni_upload_errors WHERE {wsql}', params).fetchone()[0]

        # Per-plant breakdown
        per_plant = db.execute(
            f'''SELECT plant_id, COUNT(*) AS cnt,
                       SUM(CASE WHEN row_status='error' THEN 1 ELSE 0 END) AS err_cnt,
                       SUM(CASE WHEN row_status='duplicate' THEN 1 ELSE 0 END) AS dup_cnt
                FROM tni_upload_errors WHERE {wsql}
                GROUP BY plant_id ORDER BY cnt DESC''', params).fetchall()
        per_plant = [{
            'plant_id': r['plant_id'],
            'plant_name': PLANT_MAP.get(r['plant_id'], {}).get('name', f'Plant {r["plant_id"]}'),
            'cnt': r['cnt'], 'err_cnt': r['err_cnt'] or 0, 'dup_cnt': r['dup_cnt'] or 0,
        } for r in per_plant]

        # Top error patterns (group by first 80 chars of issues)
        top_issues = db.execute(
            f'''SELECT SUBSTR(issues,1,80) AS issue_key, COUNT(*) AS cnt
                FROM tni_upload_errors WHERE {wsql} AND issues IS NOT NULL AND issues!=''
                GROUP BY issue_key ORDER BY cnt DESC LIMIT 15''', params).fetchall()

        # Top users by error count
        per_user = db.execute(
            f'''SELECT username, plant_id, COUNT(*) AS cnt
                FROM tni_upload_errors WHERE {wsql}
                GROUP BY username, plant_id ORDER BY cnt DESC LIMIT 20''', params).fetchall()
        per_user = [{
            'username': r['username'],
            'plant_id': r['plant_id'],
            'plant_name': PLANT_MAP.get(r['plant_id'], {}).get('name', '—'),
            'cnt': r['cnt'],
        } for r in per_user]

        # Recent rows
        rows = db.execute(
            f'''SELECT * FROM tni_upload_errors WHERE {wsql}
                ORDER BY ts DESC LIMIT ?''', params + [limit]).fetchall()
        rows = [{
            **dict(r),
            'plant_name': PLANT_MAP.get(r['plant_id'], {}).get('name', '—'),
        } for r in rows]

        return render_template('central_tni_errors.html',
                               total_errors=total_errors,
                               plants_involved=plants_involved,
                               users_involved=users_involved,
                               per_plant=per_plant,
                               top_issues=top_issues,
                               per_user=per_user,
                               rows=rows,
                               plants=[p for p in PLANTS if p['id'] != 99],
                               filters={'plant': plant_filter, 'from': date_from,
                                        'to': date_to, 'status': status_f, 'limit': limit})

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
