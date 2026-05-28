from flask import render_template, request, session

from tms.db import get_db
from tms.decorators import central_required


def _register(app):

    @app.route('/anomalies')
    @central_required
    def anomalies_review():
        db = get_db()
        plant_filter = request.args.get('plant', '').strip()
        flag_filter  = request.args.get('flag', '').strip()

        ph_2a, ph_2c = [], []
        wh_2a, wh_2c = ["t.anomaly_flags IS NOT NULL AND t.anomaly_flags != ''"], \
                       ["pd.anomaly_flags IS NOT NULL AND pd.anomaly_flags != ''"]

        if plant_filter:
            try:
                pid = int(plant_filter)
                wh_2a.append("t.plant_id = ?"); ph_2a.append(pid)
                wh_2c.append("pd.plant_id = ?"); ph_2c.append(pid)
            except ValueError:
                pass
        if flag_filter:
            wh_2a.append("t.anomaly_flags LIKE ?"); ph_2a.append(f'%{flag_filter}%')
            wh_2c.append("pd.anomaly_flags LIKE ?"); ph_2c.append(f'%{flag_filter}%')

        rows_2a = db.execute(
            'SELECT t.id, t.plant_id, t.emp_code, t.session_code, t.programme_name, '
            '       t.start_date, t.hrs, t.anomaly_flags, t.created_at, '
            '       e.name AS emp_name, e.collar, '
            '       p.name AS plant_name, p.unit_code '
            'FROM emp_training t '
            'LEFT JOIN employees e ON e.plant_id=t.plant_id AND e.emp_code=t.emp_code '
            'LEFT JOIN plants p ON p.id=t.plant_id '
            'WHERE ' + ' AND '.join(wh_2a) + ' '
            'ORDER BY t.created_at DESC LIMIT 500',
            ph_2a).fetchall()

        rows_2c = db.execute(
            'SELECT pd.id, pd.plant_id, pd.session_code, pd.programme_name, '
            '       pd.start_date, pd.hours_actual, pd.faculty_name, pd.anomaly_flags, pd.created_at, '
            '       p.name AS plant_name, p.unit_code, '
            '       c.planned_pax, c.actual_pax, c.duration_hrs, c.target_audience '
            'FROM programme_details pd '
            'LEFT JOIN plants p ON p.id=pd.plant_id '
            'LEFT JOIN calendar c ON c.session_code=pd.session_code AND c.plant_id=pd.plant_id '
            'WHERE ' + ' AND '.join(wh_2c) + ' '
            'ORDER BY pd.created_at DESC LIMIT 500',
            ph_2c).fetchall()

        plants = db.execute(
            'SELECT id, name, unit_code FROM plants ORDER BY name'
        ).fetchall()

        flag_options = ['collar_mismatch', 'hours_over', 'hours_mismatch',
                        'date_outside', 'low_attendance']

        return render_template('anomalies.html',
                               rows_2a=rows_2a, rows_2c=rows_2c,
                               plants=plants, flag_options=flag_options,
                               plant_filter=plant_filter, flag_filter=flag_filter)
