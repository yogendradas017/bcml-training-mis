from functools import wraps
from flask import session, flash, redirect, url_for, request


# Endpoints exempt from the "central/admin must enrol 2FA" gate.
# These are the pages the user needs to reach in order to set up 2FA / change password / log out.
_2FA_SETUP_ALLOWED = {
    'self_2fa_setup', 'admin_2fa_setup', 'admin_2fa_disable',
    'logout', 'change_password',
    'login', 'login_2fa', 'static', 'favicon', 'health'
}


def _require_2fa_if_privileged():
    """Central/Admin users must have 2FA enabled. If not, redirect to setup."""
    role = session.get('role')
    if role not in ('central', 'admin'):
        return None
    if session.get('totp_enabled'):
        return None
    # Check fresh from DB (session may be stale right after enrolment)
    try:
        from tms.db import get_db
        row = get_db().execute(
            'SELECT totp_enabled FROM users WHERE id=?', (session.get('user_id'),)
        ).fetchone()
        if row and row['totp_enabled']:
            session['totp_enabled'] = True
            return None
    except Exception:
        pass
    if request.endpoint in _2FA_SETUP_ALLOWED:
        return None
    flash('Two-factor authentication is mandatory for this role. Please set it up to continue.', 'warning')
    return redirect(url_for('self_2fa_setup'))


def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        r = _require_2fa_if_privileged()
        if r is not None:
            return r
        return f(*args, **kwargs)
    return decorated


def spoc_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        if session.get('role') not in ('spoc', 'admin'):
            flash('Access denied.', 'danger')
            return redirect(url_for('central_dashboard'))
        if session.get('role') == 'admin' and not session.get('plant_id'):
            flash('Please select a plant first to access SPOC functions.', 'warning')
            return redirect(url_for('central_dashboard'))
        r = _require_2fa_if_privileged()
        if r is not None:
            return r
        return f(*args, **kwargs)
    return decorated


def central_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        if session.get('role') not in ('central', 'admin'):
            flash('Access denied.', 'danger')
            return redirect(url_for('spoc_dashboard'))
        r = _require_2fa_if_privileged()
        if r is not None:
            return r
        return f(*args, **kwargs)
    return decorated


def spoc_or_central_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        role = session.get('role')
        if role not in ('spoc', 'central', 'admin'):
            flash('Access denied.', 'danger')
            return redirect(url_for('index'))
        r = _require_2fa_if_privileged()
        if r is not None:
            return r
        return f(*args, **kwargs)
    return decorated


def admin_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        if session.get('role') != 'admin':
            flash('This action requires admin access.', 'danger')
            return redirect(url_for('index'))
        r = _require_2fa_if_privileged()
        if r is not None:
            return r
        return f(*args, **kwargs)
    return decorated
