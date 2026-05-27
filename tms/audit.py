import hashlib
import logging
import os
from datetime import datetime, timezone, timedelta
from flask import request, session

_FALLBACK_LOG = os.path.join(os.path.dirname(__file__), '..', 'data', 'audit_fallback.log')
_IST = timezone(timedelta(hours=5, minutes=30))
_GENESIS = '0' * 64  # SHA-256 zero hash for first row


def _now_ist():
    return datetime.now(_IST).strftime('%Y-%m-%d %H:%M:%S')


def _compute_row_hash(prev_hash, ts, uid, username, plant_id, action, detail, ip):
    """SHA-256 over canonical concatenation. Any tampered row breaks the chain."""
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
    ])
    h.update(payload.encode('utf-8'))
    return h.hexdigest()


def log_action(action, detail='', user_id=None, username=None, plant_id=None):
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

        # Tamper-evident chain: link to previous row's hash
        last = db.execute(
            'SELECT row_hash FROM audit_log ORDER BY id DESC LIMIT 1'
        ).fetchone()
        prev_hash = (last['row_hash'] if last and last['row_hash'] else _GENESIS)
        row_hash  = _compute_row_hash(prev_hash, ts, uid, uname, pid, action, det, ip)

        db.execute(
            'INSERT INTO audit_log(ts,user_id,username,plant_id,action,detail,ip_address,prev_hash,row_hash)'
            ' VALUES(?,?,?,?,?,?,?,?,?)',
            (ts, uid, uname, pid, action, det, ip, prev_hash, row_hash)
        )
        db.commit()
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
    Empty list = chain intact. Used by admin verification page."""
    q = 'SELECT id, ts, user_id, username, plant_id, action, detail, ip_address, prev_hash, row_hash FROM audit_log ORDER BY id ASC'
    if limit:
        q += f' LIMIT {int(limit)}'
    broken = []
    prev = _GENESIS
    for r in db.execute(q):
        expected = _compute_row_hash(
            prev, r['ts'], r['user_id'], r['username'],
            r['plant_id'], r['action'], r['detail'] or '', r['ip_address'] or ''
        )
        if r['row_hash'] and r['row_hash'] != expected:
            broken.append(r['id'])
        if r['prev_hash'] and r['prev_hash'] != prev:
            if r['id'] not in broken:
                broken.append(r['id'])
        prev = r['row_hash'] or expected
    return broken
