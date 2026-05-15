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

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'bcml-tms-2627-xK9pQ')
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024
app.config['WTF_CSRF_TIME_LIMIT'] = 3600

csrf     = CSRFProtect(app)
limiter  = Limiter(get_remote_address, app=app, default_limits=[], storage_uri='memory://',
                   swallow_errors=True)
Compress(app)

# Session: stays alive 8 hours; survives browser close
app.config['SESSION_PERMANENT']          = True
app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(hours=8)

# Security: only send session cookie over HTTPS in production
_on_render = bool(os.environ.get('RENDER'))
app.config['SESSION_COOKIE_SECURE']   = _on_render
app.config['SESSION_COOKIE_HTTPONLY'] = True
app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'

try:
    _sv = subprocess.check_output(['git', 'rev-parse', '--short', 'HEAD'],
                                   stderr=subprocess.DEVNULL).decode().strip()
except Exception:
    _sv = '1'
app.config['STATIC_VER'] = _sv


@app.after_request
def add_security_headers(response):
    response.headers.setdefault('X-Content-Type-Options', 'nosniff')
    response.headers.setdefault('X-Frame-Options', 'SAMEORIGIN')
    response.headers.setdefault('X-XSS-Protection', '1; mode=block')
    response.headers.setdefault('Referrer-Policy', 'strict-origin-when-cross-origin')
    return response


@app.teardown_appcontext
def close_db(e=None):
    db = g.pop('db', None)
    if db:
        db.close()


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


@app.errorhandler(CSRFError)
def csrf_error(e):
    flash('Session expired or form was stale — please try again.', 'warning')
    return redirect(request.referrer or url_for('login')), 400


@app.errorhandler(413)
def upload_too_large(e):
    flash('File too large. Maximum upload size is 16 MB.', 'danger')
    return redirect(request.referrer or url_for('index'))


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
    logging.error(f'500 error: {e}', exc_info=True)
    wants_html = 'text/html' in request.accept_mimetypes.values()
    if wants_html:
        flash('Server error. Please try again or contact Corporate L&D.', 'danger')
    referrer = request.referrer or ''
    # Avoid redirect loop: don't go back to the URL that just crashed
    if referrer and request.path and request.path not in referrer:
        return redirect(referrer), 302
    return redirect(url_for('index')), 302


@app.route('/health')
@csrf.exempt
def health():
    try:
        db = get_db()
        db.execute('SELECT 1')
        return {'status': 'ok'}, 200
    except Exception as e:
        return {'status': 'error', 'detail': str(e)}, 500


# Register all routes (deferred imports — app is defined above)
from tms.routes import (auth, employees, tni, programme, calendar, training,
                        summary, central, export, api, qr, central_training, reports, requests)

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

# Rate-limit login: 20 attempts/minute per IP
limiter.limit('20 per minute')(app.view_functions['login'])

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
