import hashlib
import json as _json
import logging
import os
from datetime import datetime, timezone, timedelta
from flask import request, session

_FALLBACK_LOG = os.path.join(os.path.dirname(__file__), '..', 'data', 'audit_fallback.log')
_IST = timezone(timedelta(hours=5, minutes=30))
_GENESIS = '0' * 64  # SHA-256 zero hash for first row


def _now_ist():
    return datetime.now(_IST).strftime('%Y-%m-%d %H:%M:%S')


def _compute_row_hash(prev_hash, ts, uid, username, plant_id, action, detail, ip, payload_hash=None):
    """SHA-256 over canonical concatenation. Any tampered row breaks the chain.

    payload_hash is included in the digest if present so any change to the
    detailed payload also breaks the chain (defence in depth).
    """
    h = hashlib.sha256()
    payload = '|'.join([
        prev_hash or _GENESIS,
        ts or '',
        str(uid) if uid is not None else '',
        username or '',
        str(plant_id) if plant_id is not None else '',
        action or '',
        detail or '',
        ip or '',
        payload_hash or '',
    ])
    h.update(payload.encode('utf-8'))
    return h.hexdigest()


def _compute_payload_hash(payload_json_str):
    """SHA-256 of the canonical payload JSON string."""
    if not payload_json_str:
        return ''
    return hashlib.sha256(payload_json_str.encode('utf-8')).hexdigest()


def _row_diff(before, after, ignore_keys=()):
    """Compute field-level diff dict {field: {before, after}} between two row dicts.
    Strips ignored keys + ts-style auto fields. Returns {} if no changes."""
    diff = {}
    keys = (set(before or {}) | set(after or {})) - set(ignore_keys or ())
    for k in keys:
        b = (before or {}).get(k)
        a = (after  or {}).get(k)
        if b != a:
            diff[k] = {'before': b, 'after': a}
    return diff


def log_action(action, detail='', user_id=None, username=None, plant_id=None,
               payload=None):
    try:
        from tms.db import get_db
        db = get_db()
        uid   = user_id  or session.get('user_id')
        uname = username or session.get('username', 'unknown')
        pid   = plant_id or session.get('plant_id')
        ip    = request.headers.get('X-Forwarded-For', request.remote_addr or '')
        if ',' in (ip or ''):
            ip = ip.split(',')[0].strip()
        ip = (ip or '')[:45]
        ts  = _now_ist()
        det = str(detail)[:500]

        # Serialise payload to canonical JSON (sorted keys, str default) so
        # the hash is reproducible at verify time.
        payload_json_str = None
        payload_hash = ''
        if payload is not None:
            try:
                payload_json_str = _json.dumps(payload, sort_keys=True, default=str,
                                                separators=(',', ':'))
                payload_hash = _compute_payload_hash(payload_json_str)
            except (TypeError, ValueError):
                payload_json_str = None
                payload_hash = ''

        # Tamper-evident chain: link to previous row's hash.
        # Use BEGIN IMMEDIATE to take a RESERVED lock before reading the tail,
        # so concurrent writers serialise and no two rows ever compute against
        # the same prev_hash (which would invalidate the chain).
        try:
            db.execute('BEGIN IMMEDIATE')
        except Exception:
            # Already in a transaction (autocommit off + pending stmt) — proceed;
            # SQLite will still serialise the upcoming write.
            pass
        try:
            last = db.execute(
                'SELECT row_hash FROM audit_log ORDER BY id DESC LIMIT 1'
            ).fetchone()
            prev_hash = (last['row_hash'] if last and last['row_hash'] else _GENESIS)
            row_hash  = _compute_row_hash(prev_hash, ts, uid, uname, pid, action, det, ip,
                                           payload_hash)

            db.execute(
                'INSERT INTO audit_log(ts,user_id,username,plant_id,action,detail,ip_address,'
                'prev_hash,row_hash,payload_json,payload_hash)'
                ' VALUES(?,?,?,?,?,?,?,?,?,?,?)',
                (ts, uid, uname, pid, action, det, ip,
                 prev_hash, row_hash, payload_json_str, payload_hash)
            )
            db.commit()
        except Exception:
            try:
                db.rollback()
            except Exception:
                pass
            raise
    except Exception as e:
        logging.warning(f'audit_log write failed: {e}')
        try:
            uname = username or (session.get('username', 'unknown') if session else 'unknown')
            with open(_FALLBACK_LOG, 'a', encoding='utf-8') as fh:
                fh.write(f"{_now_ist()} | {action} | {uname} | {detail} | err:{e}\n")
        except Exception:
            pass


def verify_chain(db, limit=None):
    """Recompute hashes for the entire audit_log and return list of broken row ids.
    Empty list = chain intact. Used by admin verification page.

    Verifies both the row_hash chain AND the payload_hash (if present).
    """
    q = ('SELECT id, ts, user_id, username, plant_id, action, detail, ip_address, '
         'prev_hash, row_hash, payload_json, payload_hash '
         'FROM audit_log ORDER BY id ASC')
    if limit:
        q += f' LIMIT {int(limit)}'
    broken = []
    prev = _GENESIS
    for r in db.execute(q):
        # 1. payload_hash check (if payload exists)
        stored_payload_hash = r['payload_hash'] if 'payload_hash' in r.keys() else ''
        payload_str = r['payload_json'] if 'payload_json' in r.keys() else None
        recomputed_payload_hash = _compute_payload_hash(payload_str) if payload_str else ''
        if stored_payload_hash and recomputed_payload_hash and stored_payload_hash != recomputed_payload_hash:
            broken.append(r['id'])
        # 2. row_hash chain check
        expected = _compute_row_hash(
            prev, r['ts'], r['user_id'], r['username'],
            r['plant_id'], r['action'], r['detail'] or '', r['ip_address'] or '',
            stored_payload_hash or ''
        )
        if r['row_hash'] and r['row_hash'] != expected:
            if r['id'] not in broken:
                broken.append(r['id'])
        if r['prev_hash'] and r['prev_hash'] != prev:
            if r['id'] not in broken:
                broken.append(r['id'])
        prev = r['row_hash'] or expected
    return broken


# Convenience wrapper for snapshot-on-write callers
def log_record_change(action, row_id, table, before=None, after=None,
                       extra_detail=''):
    """Convenience: log a row change with a structured payload.

    action: e.g. 'RECORD_EDIT' / 'RECORD_ADD' / 'RECORD_DELETE'
    row_id: scalar row identifier (e.g. cal_id)
    table:  source table name for context ('calendar', 'programme_details', etc.)
    before / after: dict snapshots; for ADD pass before=None, for DELETE pass after=None
    """
    diff = _row_diff(before, after,
                     ignore_keys=('id', 'created_at', 'updated_at',
                                  'audit_json', 'created_by'))
    payload = {
        'table':  table,
        'row_id': row_id,
        'before': before,
        'after':  after,
        'diff':   diff,
    }
    short_fields = ','.join(sorted(diff.keys())) if diff else ''
    if before and not after:
        detail = f'{table}:{row_id} DELETED'
    elif after and not before:
        detail = f'{table}:{row_id} CREATED'
    else:
        detail = f'{table}:{row_id} fields:{short_fields}'
    if extra_detail:
        detail = f'{detail} | {extra_detail}'
    log_action(action, detail=detail, payload=payload)
