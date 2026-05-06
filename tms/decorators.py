from functools import wraps
from flask import session, flash, redirect, url_for


def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
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
        return f(*args, **kwargs)
    return decorated
