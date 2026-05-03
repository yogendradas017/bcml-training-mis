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
    """Add fy_year to tni and rebuild UNIQUE to include it — prevents re-nominations being silently dropped on FY rollover."""
    from tms.helpers import _fy_label

    cols = [row[1] for row in db.execute("PRAGMA table_info(tni)").fetchall()]
    if not cols:
        return  # tni table doesn't exist yet; schema.sql will create it

    fy_exists = 'fy_year' in cols
    unique_updated = bool(db.execute(
        "SELECT 1 FROM sqlite_master WHERE type='index' AND tbl_name='tni' AND sql LIKE '%fy_year%'"
    ).fetchone())

    if fy_exists and unique_updated:
        return  # fully migrated

    if not fy_exists:
        fy = _fy_label()
        db.execute("ALTER TABLE tni ADD COLUMN fy_year TEXT NOT NULL DEFAULT ''")
        db.execute("UPDATE tni SET fy_year=? WHERE fy_year=''", (fy,))
        db.commit()

    # Determine columns to copy — source may not exist on very old DBs
    existing_cols = [row[1] for row in db.execute("PRAGMA table_info(tni)").fetchall()]
    select_cols = ['id', 'plant_id', 'emp_code', 'programme_name', 'prog_type',
                   'mode', 'target_month', 'planned_hours', 'fy_year', 'created_at']
    if 'source' in existing_cols:
        select_cols.insert(8, 'source')
    cols_sql = ', '.join(select_cols)

    # Drop any leftover tni_fy from a previous partial run
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
        CREATE INDEX IF NOT EXISTS idx_tni_plant  ON tni(plant_id);
        CREATE INDEX IF NOT EXISTS idx_tni_dedup  ON tni(plant_id, emp_code, programme_name, fy_year);
        CREATE INDEX IF NOT EXISTS idx_tni_prog   ON tni(plant_id, programme_name);
        CREATE INDEX IF NOT EXISTS idx_tni_fy     ON tni(plant_id, fy_year);
    ''')


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
    for p in PLANTS:
        db.execute('INSERT OR IGNORE INTO plants(id,name,unit_code) VALUES(?,?,?)',
                   (p['id'], p['name'], p['unit_code']))
    users = [('central', 'bcml@1234', 'central', None),
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
