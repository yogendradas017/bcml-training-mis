from datetime import datetime as _dt

from flask import render_template, request, redirect, url_for, session, flash

from tms.db import get_db
from tms.decorators import spoc_required, admin_required
from tms.audit import log_action


def _register(app):

    @app.route('/requests/submit', methods=['GET', 'POST'])
    @spoc_required
    def spoc_submit_request():
        plant_id = session['plant_id']
        if request.method == 'POST':
            req_type = request.form.get('request_type', '').strip()
            details  = request.form.get('details', '').strip()
            valid_types = ('TNI_ADD', 'MARK_CONDUCTED', 'MANUAL_ATTENDANCE', 'OTHER')
            if req_type not in valid_types:
                flash('Invalid request type.', 'danger')
                return redirect(url_for('spoc_submit_request'))
            if not details or len(details) < 20:
                flash('Please provide a detailed description (at least 20 characters).', 'danger')
                return redirect(url_for('spoc_submit_request'))
            db = get_db()
            db.execute(
                '''INSERT INTO spoc_requests(plant_id, requested_by, request_type, details)
                   VALUES(?, ?, ?, ?)''',
                (plant_id, session['username'], req_type, details[:2000])
            )
            db.commit()
            log_action('RECORD_ADD', f"spoc_request:{req_type}")
            flash('Override request submitted. Admin will review and respond shortly.', 'success')
            return redirect(url_for('spoc_submit_request'))

        db = get_db()
        my_requests = db.execute(
            '''SELECT * FROM spoc_requests WHERE plant_id=? ORDER BY ts DESC LIMIT 50''',
            (plant_id,)
        ).fetchall()
        return render_template('spoc_request_form.html', my_requests=my_requests)

    @app.route('/admin/requests')
    @admin_required
    def admin_requests():
        db = get_db()
        rows = db.execute('''
            SELECT r.*, p.name as plant_name
            FROM spoc_requests r
            LEFT JOIN plants p ON p.id = r.plant_id
            ORDER BY CASE r.status WHEN 'Pending' THEN 0 ELSE 1 END, r.ts DESC
            LIMIT 200
        ''').fetchall()
        return render_template('admin_requests.html', rows=rows)

    @app.route('/admin/requests/<int:req_id>/review', methods=['POST'])
    @admin_required
    def admin_review_request(req_id):
        action      = request.form.get('action', '')
        review_note = request.form.get('review_note', '').strip()
        if action not in ('approve', 'reject'):
            flash('Invalid action.', 'danger')
            return redirect(url_for('admin_requests'))
        status = 'Approved' if action == 'approve' else 'Rejected'
        db = get_db()
        db.execute(
            '''UPDATE spoc_requests
               SET status=?, reviewed_by=?, reviewed_at=?, review_note=?
               WHERE id=?''',
            (status, session['username'], _dt.now().isoformat(timespec='seconds'),
             review_note[:500], req_id)
        )
        db.commit()
        log_action('RECORD_EDIT', f"spoc_request:{req_id}:{status}")
        flash(f'Request {req_id} {status.lower()}.', 'success' if status == 'Approved' else 'warning')
        return redirect(url_for('admin_requests'))
