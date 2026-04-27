import os
import subprocess
from flask import Flask, g, flash, redirect, request, url_for

from tms.constants import BASE_DIR
from tms.db import get_db, init_db

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'bcml-tms-2627-xK9pQ')
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

try:
    _sv = subprocess.check_output(['git', 'rev-parse', '--short', 'HEAD'],
                                   stderr=subprocess.DEVNULL).decode().strip()
except Exception:
    _sv = '1'
app.config['STATIC_VER'] = _sv


@app.teardown_appcontext
def close_db(e=None):
    db = g.pop('db', None)
    if db:
        db.close()


@app.errorhandler(413)
def upload_too_large(e):
    flash('File too large. Maximum upload size is 16 MB.', 'danger')
    return redirect(request.referrer or url_for('index'))


# Register all routes (deferred imports — app is defined above)
from tms.routes import auth, employees, tni, programme, calendar, training, summary, central, export, api

auth.       _register(app)
employees.  _register(app)
tni.        _register(app)
programme.  _register(app)
calendar.   _register(app)
training.   _register(app)
summary.    _register(app)
central.    _register(app)
export.     _register(app)
api.        _register(app)


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
