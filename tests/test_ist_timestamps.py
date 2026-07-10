# -*- coding: utf-8 -*-
"""Regression test for the IST-vs-UTC timestamp bug (register #14/#19/#22).

Under a UTC server (as on Render), the audit tables used to rely on the SQLite
column DEFAULT datetime('now','localtime') = UTC, so timestamps landed ~5.5h off
and month rollups fell in the wrong month. The fix writes them explicitly via
_now_ist(). This test forces TZ=UTC, submits a SPOC override request, and asserts
the stored spoc_requests.ts is IST (matches _now_ist, NOT the UTC default).

Run:  TZ=UTC DATABASE_PATH=data/demo.db SECRET_KEY=t python tests/test_ist_timestamps.py
Exit 0 = pass.
"""
import os, sys
os.environ['TZ'] = 'UTC'
os.environ.setdefault('DATABASE_PATH', 'data/demo.db')
os.environ.setdefault('SECRET_KEY', 'ist-test')
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
try: sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception: pass
from datetime import datetime

import app as appmod
from tms.helpers import _now_ist

a = appmod.app
a.config['WTF_CSRF_ENABLED'] = False
a.testing = True
client = a.test_client()
with client.session_transaction() as s:
    s.update({'user_id': 3, 'username': 'balrampur', 'role': 'spoc', 'plant_id': 1,
              'plant_name': 'Balrampur', 'totp_enabled': True})

r = client.post('/requests/submit',
                data={'request_type': 'OTHER', 'details': 'QA IST regression probe'},
                follow_redirects=False)

with a.app_context():
    from tms.db import get_db
    row = get_db().execute(
        "SELECT ts FROM spoc_requests WHERE plant_id=1 AND details LIKE 'QA IST regression probe%' "
        "ORDER BY id DESC LIMIT 1").fetchone()

assert row and row['ts'], 'no spoc_requests row written (post status %s)' % r.status_code
stored = datetime.fromisoformat(row['ts'])
ist = _now_ist().replace(tzinfo=None)
skew_min = abs((stored - ist).total_seconds()) / 60
print('stored ts:', row['ts'])
print('_now_ist :', ist.isoformat(timespec='seconds'))
print('skew from IST: %.1f min' % skew_min)
assert skew_min < 3, 'FAIL: stored ts is not IST (skew %.1f min — likely UTC default)' % skew_min
print('PASS: spoc_requests.ts stored in IST under a UTC server')
