from flask import render_template, request, session

from tms.constants import MONTHS_FY, PROG_TYPES
from tms.db import get_db
from tms.decorators import spoc_required
from tms.helpers import _calc_summary, _calc_totals, _calc_compliance


def _register(app):

    @app.route('/summary')
    @spoc_required
    def monthly_summary():
        plant_id   = session['plant_id']
        sel_month  = request.args.get('month', '')
        db         = get_db()
        summary_rows = _calc_summary(plant_id, sel_month, db)
        totals       = _calc_totals(summary_rows)
        compliance   = _calc_compliance(plant_id, db)
        return render_template('summary.html', summary_rows=summary_rows,
                               totals=totals, compliance=compliance,
                               months=MONTHS_FY, selected_month=sel_month,
                               prog_types=PROG_TYPES)
