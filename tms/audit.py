import logging
import os
from flask import request, session

_FALLBACK_LOG = os.path.join(os.path.dirname(__file__), '..', 'data', 'audit_fallback.log')


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
        db.execute(
            'INSERT INTO audit_log(user_id,username,plant_id,action,detail,ip_address)'
            ' VALUES(?,?,?,?,?,?)',
            (uid, uname, pid, action, str(detail)[:500], (ip or '')[:45])
        )
        db.commit()
    except Exception as e:
        logging.warning(f'audit_log write failed: {e}')
        # Fallback: write to file so audit entries are never silently lost
        try:
            from datetime import datetime
            uid   = user_id  or (session.get('user_id') if session else None)
            uname = username or (session.get('username', 'unknown') if session else 'unknown')
            with open(_FALLBACK_LOG, 'a', encoding='utf-8') as fh:
                fh.write(f"{datetime.now().isoformat()} | {action} | {uname} | {detail} | err:{e}\n")
        except Exception:
            pass
