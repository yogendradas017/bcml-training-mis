"""Tenant configuration — scoped key-value store with plant-level override.

Resolution order on read:
  1. (scope='plant', plant_id=<X>, key)
  2. (scope='global', plant_id=NULL, key)
  3. caller-supplied `default`

Cache strategy: request-scoped via flask.g (coherent across gunicorn workers
because it lives only for one request). No cross-request cache; the live DB
is the source of truth so admin edits take effect on the very next request
in every worker.

Write ordering: audit-first. We snapshot before, write the audit row (which
advances the hash chain in its own transaction), then commit the config
write. If the audit fails we abort and the config is unchanged.
"""
import json
import logging
from flask import session, g, has_request_context, has_app_context
from tms.db import get_db

_COERCE = {
    'string': lambda v: '' if v is None else str(v),
    'int':    lambda v: int(str(v).strip()),
    'float':  lambda v: float(str(v).strip()),
    'bool':   lambda v: str(v).strip().lower() in ('1', 'true', 'yes', 'on'),
    'json':   lambda v: json.loads(v) if isinstance(v, str) else v,
}

# (key, default, value_type, label, category, validator_or_None, is_secret)
# validator: callable(value)->None, raises ValueError on bad input.
def _int_bounds(lo, hi):
    def _v(x):
        n = int(x)
        if n < lo or n > hi:
            raise ValueError(f'must be a whole number between {lo} and {hi}')
    return _v

CONFIG_REGISTRY = [
    # key, default, type, label, category, validator, is_secret
    ('mh_target_bc', '12', 'int',
     'Annual man-hour target — Blue Collar', 'compliance',
     _int_bounds(1, 8760), False),
    ('mh_target_wc', '24', 'int',
     'Annual man-hour target — White Collar', 'compliance',
     _int_bounds(1, 8760), False),
]

# Backwards-compat alias used by db.py migration seed.
CONFIG_DEFAULTS = [(k, d, t, lab, cat) for (k, d, t, lab, cat, _v, _s) in CONFIG_REGISTRY]

_REGISTRY_BY_KEY = {r[0]: r for r in CONFIG_REGISTRY}

_UNSET = object()


def _cast(row, default, key=None):
    if not row:
        return default
    try:
        return _COERCE[row['value_type']](row['value'])
    except Exception as e:
        logging.warning('org_config bad value for %s: %r (type %s) — using default (%s)',
                        key, row['value'], row['value_type'], e)
        return default


def _load_all_into_g():
    """Single-query batch load of relevant config into flask.g for this request."""
    if not has_app_context():
        return None
    cache = getattr(g, '_org_config_cache', None)
    if cache is not None:
        return cache
    pid = None
    if has_request_context():
        try:
            pid = session.get('plant_id')
        except Exception:
            pid = None
    db = get_db()
    rows = db.execute(
        "SELECT scope, plant_id, key, value, value_type FROM org_config "
        "WHERE scope='global' OR (scope='plant' AND plant_id=?)",
        (pid,)
    ).fetchall()
    # Resolve: plant override wins over global per key.
    glob, plant = {}, {}
    for r in rows:
        bucket = plant if r['scope'] == 'plant' else glob
        bucket[r['key']] = {'value': r['value'], 'value_type': r['value_type']}
    cache = {'global': glob, 'plant': plant, 'plant_id': pid}
    g._org_config_cache = cache
    return cache


def invalidate_request_cache():
    if has_app_context():
        try:
            if hasattr(g, '_org_config_cache'):
                delattr(g, '_org_config_cache')
        except Exception:
            pass


def get_config(key, default=None, plant_id=_UNSET):
    """Read a config value.

    plant_id sentinel: when _UNSET, resolve from session.plant_id if a request
    context is active, else None (global-only). Explicit None skips plant lookup.
    Explicit int looks up that plant's override directly.
    """
    # Explicit plant_id != _UNSET means caller wants a specific scope; bypass g cache
    # because the cache is keyed to the request's session.plant_id.
    if plant_id is _UNSET:
        cache = _load_all_into_g()
        if cache is not None:
            row = cache['plant'].get(key) or cache['global'].get(key)
            return _cast(row, default, key) if row else default
        plant_id = None
    db = get_db()
    if plant_id is not None:
        row = db.execute(
            "SELECT value, value_type FROM org_config "
            "WHERE scope='plant' AND plant_id=? AND key=?",
            (plant_id, key)
        ).fetchone()
        if row is not None:
            return _cast(row, default, key)
    row = db.execute(
        "SELECT value, value_type FROM org_config "
        "WHERE scope='global' AND plant_id IS NULL AND key=?",
        (key,)
    ).fetchone()
    return _cast(row, default, key) if row is not None else default


def set_config(key, new_value, scope='global', plant_id=None,
               user_id=None, username=None, expected_updated_at=None):
    """Trusted-caller-only write. Caller MUST enforce @admin_required and MUST NOT
    forward scope/plant_id from untrusted request payloads without validation.

    Audit-first ordering: snapshot before -> write audit row (advances hash chain
    in its own commit) -> commit config write. If audit raises, config is not
    written.

    expected_updated_at: optimistic-concurrency token; if set and the current
    row's updated_at is different, raises ValueError('changed_by_other').
    """
    from tms.helpers import _now_ist
    from tms.audit import log_record_change

    reg = _REGISTRY_BY_KEY.get(key)
    if not reg:
        raise ValueError(f'unknown config key: {key}')
    _k, _default, vtype, _label, _cat, validator, is_secret = reg

    # Validate scope/plant_id (defence against future regressions)
    if scope not in ('global', 'plant'):
        raise ValueError('scope must be global or plant')
    if scope == 'plant':
        if plant_id is None:
            raise ValueError('plant scope requires plant_id')
        try:
            plant_id = int(plant_id)
        except Exception:
            raise ValueError('plant_id must be int')
    else:
        plant_id = None

    # Type cast + per-key validation
    try:
        casted = _COERCE[vtype](new_value)
    except Exception as e:
        raise ValueError(f'invalid value for {vtype}: {e}')
    if validator:
        validator(new_value)
    # Canonicalise storage form
    if vtype == 'bool':
        canonical = '1' if casted else '0'
    elif vtype == 'json':
        canonical = json.dumps(casted)
    else:
        canonical = str(casted)

    db = get_db()
    if scope == 'plant':
        before_row = db.execute(
            "SELECT value, updated_at FROM org_config WHERE scope='plant' AND plant_id=? AND key=?",
            (plant_id, key)
        ).fetchone()
    else:
        before_row = db.execute(
            "SELECT value, updated_at FROM org_config WHERE scope='global' AND plant_id IS NULL AND key=?",
            (key,)
        ).fetchone()

    # Optimistic concurrency
    if expected_updated_at is not None and before_row is not None:
        if (before_row['updated_at'] or '') != (expected_updated_at or ''):
            raise ValueError('changed_by_other')

    before_val = before_row['value'] if before_row else None
    # No-op short-circuit: avoid polluting the audit chain.
    if before_val == canonical:
        return {'changed': False}

    # Audit payloads (redact secrets)
    if is_secret:
        before_payload = {'value': '***'} if before_row else None
        after_payload = {'value': '***'}
    else:
        before_payload = {'value': before_val} if before_row else None
        after_payload = {'value': canonical}

    safe_plant = int(plant_id) if plant_id is not None else None
    extra = f'scope={scope} plant_id={safe_plant}'

    # AUDIT FIRST (writes + commits its own transaction)
    try:
        log_record_change('CONFIG_EDIT', key, 'org_config',
                          before=before_payload, after=after_payload,
                          extra_detail=extra)
    except Exception:
        logging.exception('audit write failed for org_config %s — aborting save', key)
        raise

    # NOW write the config. UPSERT both branches for safety (handles deleted seed rows).
    now = _now_ist()
    try:
        if scope == 'plant':
            db.execute(
                "INSERT INTO org_config(scope, plant_id, key, value, value_type, updated_at, updated_by) "
                "VALUES('plant', ?, ?, ?, ?, ?, ?) "
                "ON CONFLICT(scope, COALESCE(plant_id,-1), key) DO UPDATE SET "
                "  value=excluded.value, updated_at=excluded.updated_at, updated_by=excluded.updated_by",
                (plant_id, key, canonical, vtype, now, username or '')
            )
        else:
            db.execute(
                "INSERT INTO org_config(scope, plant_id, key, value, value_type, updated_at, updated_by) "
                "VALUES('global', NULL, ?, ?, ?, ?, ?) "
                "ON CONFLICT(scope, COALESCE(plant_id,-1), key) DO UPDATE SET "
                "  value=excluded.value, updated_at=excluded.updated_at, updated_by=excluded.updated_by",
                (key, canonical, vtype, now, username or '')
            )
        db.commit()
    except Exception:
        logging.exception('config write failed for %s after audit — manual reconciliation may be needed', key)
        raise

    invalidate_request_cache()
    return {'changed': True}