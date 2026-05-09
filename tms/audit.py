import logging
from flask import request, session


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
