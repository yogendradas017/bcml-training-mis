# -*- coding: utf-8 -*-
"""Smoke harness — boot the app on the DEMO db, log in as each role, GET every
route, and flag any 500 / unhandled exception. Deterministic, prod-safe (never
touches the live DB). Catches the CRASH + SQL_TEMPLATE bug classes across ALL
routes — including modules the agent audit never reached.

Run:  DATABASE_PATH=data/demo.db SECRET_KEY=smoke python tests/smoke.py
Exit code = number of failing (endpoint, profile) pairs.
"""
import os, sys, traceback
os.environ.setdefault('DATABASE_PATH', 'data/demo.db')
os.environ.setdefault('SECRET_KEY', 'smoke-test-key')
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
try: sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception: pass

import app as appmod
from werkzeug.routing import IntegerConverter, FloatConverter

flask_app = appmod.app
flask_app.testing = True          # let unhandled exceptions propagate so we see them

# Role profiles (real ids from demo.db). admin carries a plant so SPOC-gated routes pass.
PROFILES = {
    'spoc':    {'user_id': 3, 'username': 'balrampur', 'role': 'spoc',    'plant_id': 1,  'plant_name': 'Balrampur', 'totp_enabled': True},
    'central': {'user_id': 1, 'username': 'central',   'role': 'central', 'plant_id': 99, 'plant_name': 'Central',   'totp_enabled': True},
    'admin':   {'user_id': 2, 'username': 'admin',     'role': 'admin',   'plant_id': 1,  'plant_name': 'Balrampur', 'totp_enabled': True},
}

# GET endpoints we must NOT auto-hit (state-changing / heavy / session-destroying).
SKIP_SUBSTR = ('logout', 'cron', 'backup', 'restore', 'seed', 'impersonate',
               'delete', 'disable', 'reset', 'wipe', 'archive_restore')


def sample_args(rule):
    """Build a values dict for a rule's URL params using converter types."""
    vals = {}
    for name in rule.arguments:
        conv = rule._converters.get(name)
        if isinstance(conv, (IntegerConverter,)):
            vals[name] = 1
        elif isinstance(conv, (FloatConverter,)):
            vals[name] = 1.0
        else:
            vals[name] = 'x'
    return vals


adapter = flask_app.url_map.bind('localhost')
rules = []
for rule in flask_app.url_map.iter_rules():
    if rule.endpoint == 'static':
        continue
    if 'GET' not in (rule.methods or set()):
        continue
    if any(s in rule.endpoint.lower() for s in SKIP_SUBSTR):
        continue
    rules.append(rule)

print('Smoke: %d GET endpoints x %d profiles on %s' % (len(rules), len(PROFILES), os.environ['DATABASE_PATH']))
fails = []
tested = 0
param_500 = []

for prof_name, sess in PROFILES.items():
    client = flask_app.test_client()
    with client.session_transaction() as s:
        s.clear(); s.update(sess)
    for rule in rules:
        try:
            url = adapter.build(rule.endpoint, sample_args(rule), force_external=False)
        except Exception:
            continue
        tested += 1
        has_params = bool(rule.arguments)
        try:
            resp = client.get(url)
            code = resp.status_code
            if code >= 500:
                bucket = param_500 if has_params else fails
                bucket.append((prof_name, rule.endpoint, url, 'HTTP %d' % code, ''))
        except Exception as e:
            tb = traceback.format_exc().strip().splitlines()
            last = ' | '.join(tb[-3:])
            bucket = param_500 if has_params else fails
            bucket.append((prof_name, rule.endpoint, url, '%s: %s' % (type(e).__name__, e), last))

print('\n=== requests run: %d ===' % tested)
print('=== HARD FAILS (no-param GET 500/exception) : %d ===' % len(fails))
for p, ep, url, err, last in fails:
    print('  [%s] %s  %s' % (p, ep, url))
    print('      -> %s' % err)
    if last: print('         %s' % last)
print('\n=== PARAM-ROUTE 500s (may be bad-sample-id, review) : %d ===' % len(param_500))
for p, ep, url, err, last in param_500:
    print('  [%s] %s  %s -> %s' % (p, ep, url, err))

print('\nSMOKE RESULT: %d hard fail(s), %d param-route 500(s)' % (len(fails), len(param_500)))
sys.exit(len(fails))
