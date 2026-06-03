import json
import logging
from flask import render_template, request, session, redirect, url_for, flash
from tms.decorators import admin_required
from tms.db import get_db
from tms.config import set_config, CONFIG_REGISTRY

# Import limiter lazily to avoid circular import (app.py imports this module)
def _get_limiter():
    try:
        from app import limiter
        return limiter
    except Exception:
        return None


def _register(app):

    @app.route('/admin/settings', methods=['GET'])
    @admin_required
    def admin_settings():
        db = get_db()
        rows = db.execute(
            "SELECT key, value, value_type, label, category, updated_at, updated_by "
            "FROM org_config WHERE scope='global' AND plant_id IS NULL "
            "ORDER BY category, key"
        ).fetchall()
        by_cat = {}
        for r in rows:
            by_cat.setdefault(r['category'] or 'general', []).append(dict(r))
        # Ephemeral-disk advisory: data persists only on a mounted persistent disk.
        import os
        db_path = os.environ.get('DATABASE_PATH', '')
        ephemeral_warning = not db_path.startswith('/app/data') and not db_path.startswith('/var/data')
        return render_template('admin_settings.html',
                               config_by_category=by_cat,
                               ephemeral_warning=ephemeral_warning)

    @app.route('/admin/settings', methods=['POST'])
    @admin_required
    def admin_settings_save():
        # CSRF enforced globally via CSRFProtect — do NOT @csrf.exempt this route.
        try:
            changes = json.loads(request.form.get('changes_json') or '{}')
        except Exception:
            flash('Invalid changes payload.', 'danger')
            return redirect(url_for('admin_settings'))
        if not isinstance(changes, dict) or not changes:
            flash('No changes to save.', 'info')
            return redirect(url_for('admin_settings'))

        allowed = {k for (k, *_) in CONFIG_REGISTRY}
        username = session.get('username', '')
        user_id = session.get('user_id')
        saved, skipped, errors = [], [], []
        for key, new_val in changes.items():
            if key not in allowed:
                errors.append(f'{key}: not a known setting')
                continue
            try:
                result = set_config(key, new_val, scope='global', plant_id=None,
                                    user_id=user_id, username=username)
                if result.get('changed'):
                    saved.append(key)
                else:
                    skipped.append(key)
            except ValueError as e:
                # User-facing validation error — safe to show
                errors.append(f'{key}: {e}')
            except Exception as e:
                logging.exception('set_config failed for %s', key)
                errors.append(f'{key}: internal error (see logs)')

        if saved:
            flash(f'Saved {len(saved)} setting(s): {", ".join(saved)}', 'success')
        if skipped:
            flash(f'No change for: {", ".join(skipped)}', 'info')
        if errors:
            flash('Issues: ' + '; '.join(errors), 'warning')
        return redirect(url_for('admin_settings'))

    # Apply rate-limit to the write endpoint if limiter is wired.
    lim = _get_limiter()
    if lim is not None:
        try:
            lim.limit('30 per minute')(app.view_functions['admin_settings_save'])
        except Exception:
            logging.warning('could not apply rate limit to admin_settings_save')