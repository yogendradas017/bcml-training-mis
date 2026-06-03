import os
import logging
import subprocess
from datetime import timedelta
from flask import Flask, g, flash, redirect, request, url_for, render_template, session
from flask_wtf.csrf import CSRFProtect, CSRFError
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address
from flask_compress import Compress

from tms.constants import BASE_DIR
from tms.db import get_db, init_db

logging.basicConfig(
    level=logging.WARNING,
    format='%(asctime)s [%(levelname)s] %(message)s'
)

# Sentry error monitoring — opt-in via SENTRY_DSN env var.
# Silent no-op if SDK not installed OR DSN not set (free tier on prod only).
_sentry_dsn = os.environ.get('SENTRY_DSN', '').strip()
if _sentry_dsn:
    try:
        import sentry_sdk
        from sentry_sdk.integrations.flask import FlaskIntegration
        sentry_sdk.init(
            dsn=_sentry_dsn,
            integrations=[FlaskIntegration()],
            traces_sample_rate=0.05,
            send_default_pii=False,
            environment=os.environ.get('RENDER_GIT_BRANCH', 'local'),
            release=os.environ.get('RENDER_GIT_COMMIT', 'dev')[:7],
        )
    except ImportError:
        logging.warning('SENTRY_DSN set but sentry-sdk not installed; skip')

app = Flask(__name__)

# SECRET_KEY: required on production. Refuse to start without it on Render.
_secret = os.environ.get('SECRET_KEY')
_on_render = bool(os.environ.get('RENDER'))
if not _secret:
    if _on_render:
        raise RuntimeError(
            'SECRET_KEY environment variable is required in production. '
            'Set a random 32+ character value in Render dashboard → Environment.'
        )
    # Local dev fallback — explicitly random per process, not a fixed string.
    import secrets as _secrets
    _secret = _secrets.token_urlsafe(48)
    logging.warning('Using ephemeral SECRET_KEY for local dev. Set SECRET_KEY env var for stable sessions.')
app.secret_key = _secret

app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024
app.config['WTF_CSRF_TIME_LIMIT'] = 3600

csrf     = CSRFProtect(app)
limiter  = Limiter(get_remote_address, app=app, default_limits=[], storage_uri='memory://',
                   swallow_errors=True)
Compress(app)

# Session: 2 hours of inactivity then re-login
app.config['SESSION_PERMANENT']          = True
app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(hours=2)

# Security: only send session cookie over HTTPS in production
app.config['SESSION_COOKIE_SECURE']   = _on_render
app.config['SESSION_COOKIE_HTTPONLY'] = True
app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'

try:
    _sv = subprocess.check_output(['git', 'rev-parse', '--short', 'HEAD'],
                                   stderr=subprocess.DEVNULL).decode().strip()
except Exception:
    _sv = '1'
app.config['STATIC_VER'] = _sv


_CSP = (
    "default-src 'self'; "
    "script-src 'self' 'unsafe-inline' https://cdn.jsdelivr.net https://cdnjs.cloudflare.com; "
    "style-src  'self' 'unsafe-inline' https://cdn.jsdelivr.net https://fonts.googleapis.com; "
    "img-src    'self' data: blob: https:; "
    "font-src   'self' data: https://fonts.gstatic.com https://cdn.jsdelivr.net; "
    "connect-src 'self'; "
    "frame-ancestors 'self'; "
    "base-uri 'self'; "
    "form-action 'self'"
)


@app.after_request
def add_security_headers(response):
    response.headers.setdefault('X-Content-Type-Options', 'nosniff')
    response.headers.setdefault('X-Frame-Options', 'SAMEORIGIN')
    response.headers.setdefault('Referrer-Policy', 'strict-origin-when-cross-origin')
    response.headers.setdefault('Permissions-Policy', 'geolocation=(self), camera=(), microphone=(), payment=()')
    response.headers.setdefault('Content-Security-Policy', _CSP)
    if _on_render:
        response.headers.setdefault('Strict-Transport-Security',
                                    'max-age=31536000; includeSubDomains')
    return response


def _safe_redirect(target, fallback_endpoint='index'):
    """Block open-redirect: only follow same-host URLs from request.referrer/next.

    Allows relative paths or URLs whose host matches request.host. Anything else
    falls back to a known endpoint.
    """
    from urllib.parse import urlparse
    if not target:
        return redirect(url_for(fallback_endpoint))
    try:
        p = urlparse(target)
        # Relative path → safe
        if not p.netloc and not p.scheme:
            return redirect(target)
        if p.netloc == request.host:
            return redirect(target)
    except Exception:
        pass
    return redirect(url_for(fallback_endpoint))


@app.teardown_appcontext
def close_db(e=None):
    db = g.pop('db', None)
    if db:
        db.close()


@app.context_processor
def inject_fy_label():
    """Single source of truth for FY label in all templates.
    Avoids hardcoded 'FY 2026–27' strings that go stale on Apr 1."""
    from tms.helpers import _fy_label_long, _fy_label
    return {'fy_label_long': _fy_label_long(), 'fy_label_short': _fy_label()}


@app.context_processor
def inject_pending_verify_count():
    """Make pending verification + anomaly counts visible to base.html sidebar."""
    role = session.get('role')
    if role not in ('central', 'admin', 'spoc'):
        return {}
    try:
        db = get_db()
        if role in ('central', 'admin'):
            verify_cnt = db.execute(
                "SELECT COUNT(*) FROM calendar WHERE status='Awaiting Verification'"
            ).fetchone()[0]
            anom_cnt = db.execute(
                "SELECT (SELECT COUNT(*) FROM emp_training WHERE anomaly_flags IS NOT NULL AND anomaly_flags != '') + "
                "       (SELECT COUNT(*) FROM programme_details WHERE anomaly_flags IS NOT NULL AND anomaly_flags != '')"
            ).fetchone()[0]
            eff_scope = ''
            eff_params = []
        else:
            pid = session.get('plant_id')
            verify_cnt = db.execute(
                "SELECT COUNT(*) FROM calendar WHERE status='Awaiting Verification' AND plant_id=?",
                (pid,)).fetchone()[0]
            anom_cnt = db.execute(
                "SELECT (SELECT COUNT(*) FROM emp_training WHERE plant_id=? AND anomaly_flags IS NOT NULL AND anomaly_flags != '') + "
                "       (SELECT COUNT(*) FROM programme_details WHERE plant_id=? AND anomaly_flags IS NOT NULL AND anomaly_flags != '')",
                (pid, pid)).fetchone()[0]
            eff_scope = ' WHERE plant_id=?'
            eff_params = [pid]
        # Effectiveness review counts (open = pending+due+overdue; overdue separately)
        # Today computed in IST so day-rollover matches user wall clock.
        from tms.helpers import _today_ist
        today_iso = _today_ist().isoformat()
        overdue_cutoff = (
            __import__('datetime').date.fromisoformat(today_iso)
            - __import__('datetime').timedelta(days=30)
        ).isoformat()
        # Single-pass: open = all incomplete; overdue = incomplete AND due_date < cutoff.
        # One round-trip instead of two by using SUM(CASE WHEN ...) buckets.
        eff_row = db.execute(
            "SELECT "
            "  SUM(CASE WHEN completed_date IS NULL THEN 1 ELSE 0 END) AS open_cnt, "
            "  SUM(CASE WHEN completed_date IS NULL AND due_date < ? THEN 1 ELSE 0 END) AS overdue_cnt "
            f"FROM effectiveness_review{eff_scope}",
            [overdue_cutoff] + eff_params).fetchone()
        eff_open_cnt = (eff_row['open_cnt'] if eff_row and eff_row['open_cnt'] is not None else 0)
        eff_overdue_cnt = (eff_row['overdue_cnt'] if eff_row and eff_row['overdue_cnt'] is not None else 0)
        return {
            'pending_verify_count': verify_cnt,
            'anomaly_count': anom_cnt,
            'eff_open_count': eff_open_cnt,
            'eff_overdue_count': eff_overdue_cnt,
        }
    except Exception:
        return {
            'pending_verify_count': 0, 'anomaly_count': 0,
            'eff_open_count': 0, 'eff_overdue_count': 0,
        }


@app.template_filter('fmt_date')
def fmt_date(value):
    """Display any stored date (YYYY-MM-DD or YYYY-MM-DD HH:MM:SS) as DD-MM-YYYY."""
    if not value:
        return '—'
    from datetime import datetime as _dt
    s = str(value).strip()[:10]
    try:
        return _dt.strptime(s, '%Y-%m-%d').strftime('%d-%m-%Y')
    except ValueError:
        return s


@app.template_filter('fmt_dt')
def fmt_dt(value):
    """Display datetime (YYYY-MM-DD HH:MM:SS or ISO) as DD-MM-YYYY HH:MM."""
    if not value:
        return '—'
    from datetime import datetime as _dt
    s = str(value).strip().replace('T', ' ')[:16]
    try:
        return _dt.strptime(s, '%Y-%m-%d %H:%M').strftime('%d-%m-%Y %H:%M')
    except ValueError:
        try:
            return _dt.strptime(s[:10], '%Y-%m-%d').strftime('%d-%m-%Y')
        except ValueError:
            return s


@app.errorhandler(CSRFError)
def csrf_error(e):
    flash('Session expired or form was stale — please try again.', 'warning')
    return _safe_redirect(request.referrer, 'login'), 400


@app.errorhandler(413)
def upload_too_large(e):
    flash('File too large. Maximum upload size is 16 MB.', 'danger')
    return _safe_redirect(request.referrer, 'index')


@app.route('/favicon.ico')
def favicon():
    return app.send_static_file('favicon.ico')


@app.errorhandler(404)
def not_found(e):
    # Don't flash for non-HTML requests (browser favicon, fetch/XHR calls)
    wants_html = 'text/html' in request.accept_mimetypes.values()
    if 'user_id' in session:
        if wants_html:
            flash('Page not found.', 'warning')
        return redirect(url_for('index'))
    return redirect(url_for('login'))


@app.errorhandler(500)
def server_error(e):
    import sys, traceback as _tb
    sys.stderr.write(f'\n=== 500 ON {request.path} ===\n')
    _tb.print_exc(file=sys.stderr)
    sys.stderr.flush()
    logging.error(f'500 error: {e}', exc_info=True)
    wants_html = 'text/html' in request.accept_mimetypes.values()
    if wants_html:
        flash('Server error. Please try again or contact Corporate L&D.', 'danger')
    referrer = request.referrer or ''
    # Avoid redirect loop: don't go back to the URL that just crashed
    if referrer and request.path and request.path not in referrer:
        return _safe_redirect(referrer, 'index'), 302
    return redirect(url_for('index')), 302


@app.route('/health')
@csrf.exempt
def health():
    try:
        db = get_db()
        db.execute('SELECT 1')
        return {'status': 'ok'}, 200
    except Exception:
        # Do not leak exception detail to public endpoint
        return {'status': 'error'}, 500


# Register all routes (deferred imports — app is defined above)
from tms.routes import (auth, employees, tni, programme, calendar, training,
                        summary, central, export, api, qr, central_training, reports, requests,
                        verify, anomalies, effectiveness)

auth.             _register(app)
employees.        _register(app)
tni.              _register(app)
programme.        _register(app)
calendar.         _register(app)
training.         _register(app)
summary.          _register(app)
central.          _register(app)
export.           _register(app)
api.              _register(app)
qr.               _register(app)
central_training. _register(app)
reports.          _register(app)
requests.         _register(app)
verify.           _register(app)
anomalies.        _register(app)
effectiveness.    _register(app)

# Rate-limit login: 20/min per IP AND 5/min per username (botnet bypass mitigation)
def _login_user_key():
    return (request.form.get('username') or '').strip().lower() or get_remote_address()

limiter.limit('20 per minute')(app.view_functions['login'])
limiter.limit('5 per minute', key_func=_login_user_key,
              error_message='Too many login attempts for this account. Wait 1 minute.')(
    app.view_functions['login'])
limiter.limit('10 per minute')(app.view_functions['login_2fa'])

# Rate-limit public QR endpoints: 10 POST/min per IP (prevents spam/flooding)
for _vf in ('qr_attend', 'qr_feedback'):
    if _vf in app.view_functions:
        limiter.limit('10 per minute')(app.view_functions[_vf])

# CSRF-exempt public QR submission routes (no session on phone scan)
for _vf in ('qr_attend', 'qr_feedback'):
    if _vf in app.view_functions:
        csrf.exempt(app.view_functions[_vf])


try:
    init_db()
except Exception as _e:
    import logging
    logging.error(f'init_db failed: {_e}')

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.config['TEMPLATES_AUTO_RELOAD'] = True
    app.jinja_env.auto_reload = True
    app.run(host='0.0.0.0', port=port, debug=False)
