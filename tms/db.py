import os
import sqlite3
from flask import g
from werkzeug.security import generate_password_hash

from tms.constants import BASE_DIR, DB_PATH, PLANTS


def get_db():
    if 'db' not in g:
        g.db = sqlite3.connect(DB_PATH)
        g.db.row_factory = sqlite3.Row
        g.db.execute("PRAGMA journal_mode=WAL")
        g.db.execute("PRAGMA foreign_keys = ON")
    return g.db


def _ensure_indexes(db):
    db.executescript('''
        CREATE INDEX IF NOT EXISTS idx_emp_plant_code ON employees(plant_id, emp_code);
        CREATE INDEX IF NOT EXISTS idx_tni_prog ON tni(plant_id, programme_name);
        CREATE INDEX IF NOT EXISTS idx_et_lookup ON emp_training(plant_id, emp_code, programme_name);
    ''')
    db.commit()


def _migrate_tni_unique(db):
    has_unique = db.execute(
        "SELECT COUNT(*) FROM sqlite_master WHERE type='index' AND tbl_name='tni' "
        "AND sql LIKE '%emp_code%' AND sql LIKE '%programme_name%'"
    ).fetchone()[0]
    if has_unique:
        return
    db.executescript('''
        CREATE TABLE tni_new (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            plant_id INTEGER NOT NULL,
            emp_code TEXT NOT NULL,
            programme_name TEXT NOT NULL,
            prog_type TEXT, mode TEXT, target_month TEXT,
            planned_hours REAL DEFAULT 0,
            created_at TEXT DEFAULT (date('now')),
            UNIQUE(plant_id, emp_code, programme_name)
        );
        INSERT OR IGNORE INTO tni_new
            (plant_id, emp_code, programme_name, prog_type, mode, target_month, planned_hours, created_at)
        SELECT plant_id, emp_code, programme_name, prog_type, mode, target_month, planned_hours, created_at
        FROM tni ORDER BY id;
        DROP TABLE tni;
        ALTER TABLE tni_new RENAME TO tni;
        CREATE INDEX IF NOT EXISTS idx_tni_plant ON tni(plant_id);
        CREATE INDEX IF NOT EXISTS idx_tni_dedup ON tni(plant_id, emp_code, programme_name);
    ''')


def _migrate_tni_fy_year(db):
    """Add fy_year to tni and rebuild UNIQUE to include it."""
    import logging
    from tms.helpers import _fy_label

    cols = [row[1] for row in db.execute("PRAGMA table_info(tni)").fetchall()]
    if not cols:
        return

    # Phase 1 (critical): add column so all fy_year queries work
    if 'fy_year' not in cols:
        fy = _fy_label()
        try:
            db.execute("ALTER TABLE tni ADD COLUMN fy_year TEXT NOT NULL DEFAULT ''")
            db.execute("UPDATE tni SET fy_year=? WHERE fy_year=''", (fy,))
            db.commit()
        except Exception as e:
            logging.error(f'tni_fy_year phase1 failed: {e}')
            try: db.rollback()
            except Exception: pass
            return

    # Phase 2 (nice-to-have): rebuild UNIQUE to include fy_year
    unique_updated = bool(db.execute(
        "SELECT 1 FROM sqlite_master WHERE type='index' AND tbl_name='tni' AND sql LIKE '%fy_year%'"
    ).fetchone())
    if unique_updated:
        return

    try:
        existing_cols = [row[1] for row in db.execute("PRAGMA table_info(tni)").fetchall()]
        select_cols = [c for c in
                       ['id', 'plant_id', 'emp_code', 'programme_name', 'prog_type',
                        'mode', 'target_month', 'planned_hours', 'source', 'fy_year', 'created_at']
                       if c in existing_cols]
        cols_sql = ', '.join(select_cols)
        db.execute("DROP TABLE IF EXISTS tni_fy")
        db.commit()
        db.executescript(f'''
            CREATE TABLE tni_fy (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                plant_id INTEGER NOT NULL,
                emp_code TEXT NOT NULL,
                programme_name TEXT NOT NULL,
                prog_type TEXT, mode TEXT, target_month TEXT,
                planned_hours REAL DEFAULT 0,
                source TEXT DEFAULT 'TNI Driven',
                fy_year TEXT NOT NULL DEFAULT '',
                created_at TEXT DEFAULT (date('now')),
                UNIQUE(plant_id, emp_code, programme_name, fy_year)
            );
            INSERT OR IGNORE INTO tni_fy ({cols_sql})
            SELECT {cols_sql} FROM tni ORDER BY id;
            DROP TABLE tni;
            ALTER TABLE tni_fy RENAME TO tni;
            CREATE INDEX IF NOT EXISTS idx_tni_plant ON tni(plant_id);
            CREATE INDEX IF NOT EXISTS idx_tni_dedup ON tni(plant_id, emp_code, programme_name, fy_year);
            CREATE INDEX IF NOT EXISTS idx_tni_prog  ON tni(plant_id, programme_name);
            CREATE INDEX IF NOT EXISTS idx_tni_fy    ON tni(plant_id, fy_year);
        ''')
    except Exception as e:
        logging.warning(f'tni_fy_year phase2 (UNIQUE rebuild) failed — column exists, app ok: {e}')


def _migrate_tni_source(db):
    cols = [row[1] for row in db.execute("PRAGMA table_info(tni)").fetchall()]
    if 'source' not in cols:
        db.execute("ALTER TABLE tni ADD COLUMN source TEXT DEFAULT 'TNI Driven'")
    db.execute("""UPDATE tni SET source='New Requirement'
        WHERE source IN ('Corp Driven','Unit Driven','Compliance Driven',
                         'New- Unit Driven','New- Corporate Driven')""")
    db.execute("""UPDATE calendar SET source='New Requirement'
        WHERE source IN ('Corp Driven','Unit Driven','Compliance Driven',
                         'New- Unit Driven','New- Corporate Driven')""")
    db.execute("UPDATE tni SET source='TNI Driven' WHERE source IS NULL OR source=''")
    db.execute("UPDATE calendar SET source='TNI Driven' WHERE source IS NULL OR source='' OR source='TNI'")
    db.commit()


def _migrate_programme_master_source(db):
    cols = [row[1] for row in db.execute("PRAGMA table_info(programme_master)").fetchall()]
    if 'source' not in cols:
        db.execute("ALTER TABLE programme_master ADD COLUMN source TEXT DEFAULT 'TNI Requirement'")
    # Always re-derive source from TNI to ensure correctness
    db.execute('''
        UPDATE programme_master SET source = CASE
            WHEN EXISTS(
                SELECT 1 FROM tni t WHERE t.plant_id=programme_master.plant_id
                AND LOWER(t.programme_name)=LOWER(programme_master.name)
                AND t.source='TNI Driven'
            ) THEN 'TNI Requirement'
            WHEN EXISTS(
                SELECT 1 FROM tni t WHERE t.plant_id=programme_master.plant_id
                AND LOWER(t.programme_name)=LOWER(programme_master.name)
            ) THEN 'New Requirement'
            ELSE 'TNI Requirement'
        END
    ''')
    db.commit()


def _ensure_qr_tables(db):
    """Create session_qr and feedback_response if not present (schema.sql handles new DBs)."""
    db.executescript('''
        CREATE TABLE IF NOT EXISTS session_qr (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            plant_id INTEGER NOT NULL,
            session_code TEXT NOT NULL,
            token TEXT NOT NULL UNIQUE,
            stage TEXT NOT NULL DEFAULT 'attendance',
            created_at TEXT DEFAULT (datetime('now','localtime')),
            expires_at TEXT,
            is_active INTEGER DEFAULT 1,
            created_by INTEGER
        );
        CREATE TABLE IF NOT EXISTS feedback_response (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            plant_id INTEGER NOT NULL,
            session_code TEXT NOT NULL,
            emp_code TEXT,
            submitted_at TEXT DEFAULT (datetime('now','localtime')),
            q_obj_explained INTEGER, q_well_structured INTEGER,
            q_content_appropriate INTEGER, q_presentation_quality INTEGER,
            q_time_reasonable INTEGER,
            q_inputs_appropriate INTEGER, q_communication_clear INTEGER,
            q_queries_responded INTEGER, q_well_involved INTEGER,
            key_learnings TEXT, suggestions TEXT,
            ip_address TEXT, lang TEXT DEFAULT 'en',
            UNIQUE(plant_id, session_code, emp_code)
        );
        CREATE INDEX IF NOT EXISTS idx_qr_token   ON session_qr(token);
        CREATE INDEX IF NOT EXISTS idx_qr_session ON session_qr(plant_id, session_code);
        CREATE INDEX IF NOT EXISTS idx_fr_session ON feedback_response(plant_id, session_code);
    ''')


def _migrate_emp_training_dedup(db):
    """Add UNIQUE(plant_id, emp_code, programme_name, start_date) to emp_training."""
    import logging
    has_unique = db.execute(
        "SELECT 1 FROM sqlite_master WHERE type='index' AND tbl_name='emp_training' "
        "AND sql LIKE '%emp_code%' AND sql LIKE '%programme_name%' AND sql LIKE '%start_date%'"
    ).fetchone()
    if has_unique:
        return
    try:
        cols = [row[1] for row in db.execute("PRAGMA table_info(emp_training)").fetchall()]
        if not cols:
            return
        col_list = ', '.join(c for c in
            ['id','plant_id','emp_code','session_code','programme_name',
             'start_date','end_date','hrs','prog_type','level','mode',
             'cal_new','pre_rating','post_rating','venue','month','created_at']
            if c in cols)
        db.execute("DROP TABLE IF EXISTS emp_training_dedup")
        db.commit()
        db.executescript(f'''
            CREATE TABLE emp_training_dedup (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                plant_id INTEGER NOT NULL,
                emp_code TEXT NOT NULL,
                session_code TEXT,
                programme_name TEXT NOT NULL,
                start_date TEXT,
                end_date TEXT,
                hrs REAL DEFAULT 0,
                prog_type TEXT,
                level TEXT,
                mode TEXT,
                cal_new TEXT,
                pre_rating REAL,
                post_rating REAL,
                venue TEXT,
                month TEXT,
                created_at TEXT DEFAULT (date('now')),
                UNIQUE(plant_id, emp_code, programme_name, start_date)
            );
            INSERT OR IGNORE INTO emp_training_dedup ({col_list})
            SELECT {col_list} FROM emp_training ORDER BY id;
            DROP TABLE emp_training;
            ALTER TABLE emp_training_dedup RENAME TO emp_training;
            CREATE INDEX IF NOT EXISTS idx_training_plant ON emp_training(plant_id);
            CREATE INDEX IF NOT EXISTS idx_et_lookup ON emp_training(plant_id, emp_code, programme_name);
            CREATE INDEX IF NOT EXISTS idx_et_dedup ON emp_training(plant_id, emp_code, programme_name, start_date);
        ''')
    except Exception as e:
        logging.warning(f'emp_training dedup migration failed: {e}')


def _migrate_session_pin(db):
    cols = [r[1] for r in db.execute("PRAGMA table_info(calendar)").fetchall()]
    if 'session_pin' not in cols:
        db.execute("ALTER TABLE calendar ADD COLUMN session_pin TEXT")
        db.commit()


def _migrate_central_plant(db):
    db.execute("INSERT OR IGNORE INTO plants(id,name,unit_code) VALUES(99,'Central','CEN')")
    db.commit()


def _migrate_calendar_is_central(db):
    cols = [r[1] for r in db.execute("PRAGMA table_info(calendar)").fetchall()]
    if 'is_central' not in cols:
        db.execute("ALTER TABLE calendar ADD COLUMN is_central INTEGER NOT NULL DEFAULT 0")
    db.execute("CREATE INDEX IF NOT EXISTS idx_cal_central ON calendar(is_central, plant_id)")
    db.commit()


def _migrate_emp_training_host(db):
    import logging
    cols = [r[1] for r in db.execute("PRAGMA table_info(emp_training)").fetchall()]
    if not cols:
        return
    if 'host_plant_id' not in cols:
        db.execute("ALTER TABLE emp_training ADD COLUMN host_plant_id INTEGER")
        db.commit()
    # Rebuild UNIQUE to include session_code (prevents same emp attending two Central sessions
    # on same date for same programme from colliding)
    v2_done = bool(db.execute(
        "SELECT 1 FROM sqlite_master WHERE type='index' AND name='idx_et_v2_marker'"
    ).fetchone())
    if v2_done:
        return
    try:
        existing = [r[1] for r in db.execute("PRAGMA table_info(emp_training)").fetchall()]
        keep = [c for c in ['id','plant_id','emp_code','session_code','programme_name',
                             'start_date','end_date','hrs','prog_type','level','mode',
                             'cal_new','pre_rating','post_rating','venue','month',
                             'host_plant_id','created_at'] if c in existing]
        cols_sql = ', '.join(keep)
        db.execute("DROP TABLE IF EXISTS emp_training_v2")
        db.commit()
        db.executescript(f'''
            CREATE TABLE emp_training_v2 (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                plant_id INTEGER NOT NULL,
                emp_code TEXT NOT NULL,
                session_code TEXT,
                programme_name TEXT NOT NULL,
                start_date TEXT, end_date TEXT,
                hrs REAL DEFAULT 0,
                prog_type TEXT, level TEXT, mode TEXT,
                cal_new TEXT, pre_rating REAL, post_rating REAL,
                venue TEXT, month TEXT,
                host_plant_id INTEGER,
                created_at TEXT DEFAULT (date('now'))
            );
            INSERT OR IGNORE INTO emp_training_v2 ({cols_sql})
                SELECT {cols_sql} FROM emp_training ORDER BY id;
            DROP TABLE emp_training;
            ALTER TABLE emp_training_v2 RENAME TO emp_training;
            CREATE INDEX IF NOT EXISTS idx_training_plant ON emp_training(plant_id);
            CREATE INDEX IF NOT EXISTS idx_et_lookup      ON emp_training(plant_id, emp_code, programme_name);
            CREATE INDEX IF NOT EXISTS idx_et_host        ON emp_training(host_plant_id);
            CREATE INDEX IF NOT EXISTS idx_et_session     ON emp_training(session_code);
            CREATE UNIQUE INDEX IF NOT EXISTS idx_et_v2_dedup
                ON emp_training(plant_id, emp_code, programme_name, start_date, COALESCE(session_code,''));
            CREATE UNIQUE INDEX idx_et_v2_marker ON emp_training(id) WHERE 1=0;
        ''')
    except Exception as e:
        logging.warning(f'emp_training v2 migration failed: {e}')


def _migrate_corp_members(db):
    db.executescript('''
        CREATE TABLE IF NOT EXISTS corp_members (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            emp_code TEXT NOT NULL UNIQUE,
            name TEXT NOT NULL,
            designation TEXT DEFAULT '',
            department TEXT DEFAULT '',
            email TEXT DEFAULT '',
            is_active INTEGER NOT NULL DEFAULT 1,
            created_at TEXT DEFAULT (date('now'))
        );
        CREATE INDEX IF NOT EXISTS idx_corp_active ON corp_members(is_active, name);
    ''')
    db.commit()


def _migrate_central_user_plant(db):
    db.execute("UPDATE users SET plant_id=99 WHERE username='central' AND (plant_id IS NULL OR plant_id=0)")
    db.commit()


def init_db():
    from tms.helpers import (
        _cleanse_master_spelling, _cleanse_programme_names, _cleanup_stale_analyze_files
    )
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    db = sqlite3.connect(DB_PATH)
    db.row_factory = sqlite3.Row
    with open(os.path.join(BASE_DIR, 'schema.sql')) as f:
        db.executescript(f.read())
    _ensure_indexes(db)
    _migrate_tni_unique(db)
    _migrate_tni_fy_year(db)
    _migrate_tni_source(db)
    _migrate_programme_master_source(db)
    _migrate_emp_training_dedup(db)
    _ensure_qr_tables(db)
    _migrate_session_pin(db)
    _migrate_central_plant(db)
    _migrate_calendar_is_central(db)
    _migrate_emp_training_host(db)
    _migrate_corp_members(db)
    _migrate_central_user_plant(db)
    for p in PLANTS:
        db.execute('INSERT OR IGNORE INTO plants(id,name,unit_code) VALUES(?,?,?)',
                   (p['id'], p['name'], p['unit_code']))
    users = [('central', 'bcml@1234', 'central', 99),
             ('admin',   'admin@bcml', 'admin',   None)]
    for p in PLANTS:
        users.append((p['name'].lower(), 'bcml@1234', 'spoc', p['id']))
    for u in users:
        db.execute('INSERT OR IGNORE INTO users(username,password,role,plant_id) VALUES(?,?,?,?)',
                   (u[0], generate_password_hash(u[1]), u[2], u[3]))
    db.commit()
    _cleanse_master_spelling(db)
    _cleanse_programme_names(db)
    db.commit()
    db.close()
    _cleanup_stale_analyze_files()
