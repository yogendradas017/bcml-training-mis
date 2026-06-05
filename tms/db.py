import logging
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
        # 15s busy_timeout (default 5s). Multiple concurrent writers
        # otherwise hit "database is locked" cascade under SQLite + WAL.
        g.db.execute("PRAGMA busy_timeout=15000")
    return g.db


def _ensure_indexes(db):
    db.executescript('''
        CREATE INDEX IF NOT EXISTS idx_emp_plant_code ON employees(plant_id, emp_code);
        CREATE INDEX IF NOT EXISTS idx_tni_prog ON tni(plant_id, programme_name);
        CREATE INDEX IF NOT EXISTS idx_et_lookup ON emp_training(plant_id, emp_code, programme_name);
        CREATE INDEX IF NOT EXISTS idx_et_plant_date ON emp_training(plant_id, start_date);
        CREATE INDEX IF NOT EXISTS idx_et_session ON emp_training(plant_id, session_code);
        CREATE INDEX IF NOT EXISTS idx_cal_plant_plan ON calendar(plant_id, plan_start);
        CREATE INDEX IF NOT EXISTS idx_cal_session ON calendar(plant_id, session_code);
        CREATE INDEX IF NOT EXISTS idx_pd_session  ON programme_details(plant_id, session_code);
        CREATE INDEX IF NOT EXISTS idx_pd_prog     ON programme_details(plant_id, programme_name);
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
    import logging
    try:
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
    except Exception as e:
        logging.warning(f'_migrate_tni_source failed: {e}')


def _migrate_programme_master_source(db):
    import logging
    try:
        cols = [row[1] for row in db.execute("PRAGMA table_info(programme_master)").fetchall()]
        if 'source' not in cols:
            db.execute("ALTER TABLE programme_master ADD COLUMN source TEXT DEFAULT 'TNI Requirement'")
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
    except Exception as e:
        logging.warning(f'_migrate_programme_master_source failed: {e}')


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
            created_by INTEGER REFERENCES users(id) ON DELETE SET NULL
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
    import logging
    try:
        cols = [r[1] for r in db.execute("PRAGMA table_info(calendar)").fetchall()]
        if 'session_pin' not in cols:
            db.execute("ALTER TABLE calendar ADD COLUMN session_pin TEXT")
            db.commit()
    except Exception as e:
        logging.warning(f'_migrate_session_pin failed: {e}')


def _migrate_central_plant(db):
    import logging
    try:
        db.execute("INSERT OR IGNORE INTO plants(id,name,unit_code) VALUES(99,'Central','CEN')")
        db.commit()
    except Exception as e:
        logging.warning(f'_migrate_central_plant failed: {e}')


def _migrate_feedback_scale_to_4(db):
    """One-time: cap legacy 1-5 feedback values to 1-4 after scale change.
    Values >4 in course_feedback/faculty_feedback/trainer_fb_* are pre-change
    data from the 1-5 scale. Clamping to 4 prevents Feedback Reports showing
    >100% percentages. Idempotent."""
    import logging
    try:
        cols = ['course_feedback','faculty_feedback',
                'trainer_fb_participants','trainer_fb_facilities']
        total = 0
        for col in cols:
            n = db.execute(
                f"UPDATE programme_details SET {col}=4 WHERE {col} > 4").rowcount
            total += n
        if total > 0:
            db.commit()
            logging.info(f'_migrate_feedback_scale_to_4: clamped {total} cells to <=4')
    except Exception as e:
        logging.warning(f'_migrate_feedback_scale_to_4 failed: {e}')


def _migrate_mode_offline_to_classroom(db):
    """One-time data hygiene: collapse legacy 'Offline' mode to canonical 'Classroom'.
    Per tms.constants.MODES (Classroom/OJT/SOP/Online), 'Offline' is not a valid value —
    it was emitted by older seed_synthetic.py. Hygiene engine in helpers maps the same
    direction (Offline -> Classroom)."""
    import logging
    try:
        n_cal = db.execute(
            "UPDATE calendar SET mode='Classroom' WHERE mode='Offline'").rowcount
        n_pd = db.execute(
            "UPDATE programme_details SET mode='Classroom' WHERE mode='Offline'").rowcount
        n_pm = db.execute(
            "UPDATE programme_master SET mode='Classroom' WHERE mode='Offline'").rowcount
        n_tni = db.execute(
            "UPDATE tni SET mode='Classroom' WHERE mode='Offline'").rowcount
        if (n_cal + n_pd + n_pm + n_tni) > 0:
            db.commit()
            logging.info(
                f"_migrate_mode_offline_to_classroom: calendar={n_cal} "
                f"programme_details={n_pd} programme_master={n_pm} tni={n_tni}")
    except Exception as e:
        logging.warning(f'_migrate_mode_offline_to_classroom failed: {e}')


def _migrate_calendar_is_central(db):
    import logging
    try:
        cols = [r[1] for r in db.execute("PRAGMA table_info(calendar)").fetchall()]
        if 'is_central' not in cols:
            db.execute("ALTER TABLE calendar ADD COLUMN is_central INTEGER NOT NULL DEFAULT 0")
        db.execute("CREATE INDEX IF NOT EXISTS idx_cal_central ON calendar(is_central, plant_id)")
        db.commit()
    except Exception as e:
        logging.warning(f'_migrate_calendar_is_central failed: {e}')


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
    import logging
    try:
        db.execute("UPDATE users SET plant_id=99 WHERE username='central' AND (plant_id IS NULL OR plant_id=0)")
        db.commit()
    except Exception as e:
        logging.warning(f'_migrate_central_user_plant failed: {e}')


def _migrate_audit_lockout(db):
    import logging
    try:
        cols = {r[1] for r in db.execute("PRAGMA table_info(users)")}
        if 'failed_attempts' not in cols:
            db.execute("ALTER TABLE users ADD COLUMN failed_attempts INTEGER DEFAULT 0")
        if 'locked_until' not in cols:
            db.execute("ALTER TABLE users ADD COLUMN locked_until TEXT")
        db.commit()
    except Exception as e:
        logging.warning(f'_migrate_audit_lockout failed: {e}')


def _migrate_force_default_password_change(db):
    """For any user whose stored hash still matches the default seed password,
    set must_change_password=1 so they cannot use the app until they rotate.
    Idempotent — runs every startup but is a no-op once the user has rotated."""
    import logging
    try:
        from werkzeug.security import check_password_hash
        rows = db.execute(
            "SELECT id, username, password, role, must_change_password FROM users"
        ).fetchall()
        for r in rows:
            if r['must_change_password']:
                continue
            try:
                default = 'admin@bcml' if r['username'] == 'admin' else 'bcml@1234'
                if check_password_hash(r['password'], default):
                    db.execute('UPDATE users SET must_change_password=1 WHERE id=?', (r['id'],))
            except Exception:
                continue
        db.commit()
    except Exception as e:
        logging.warning(f'_migrate_force_default_password_change failed: {e}')


def _migrate_audit_hash_chain(db):
    """Tamper-evident audit log: add prev_hash + row_hash columns.
    Hash chain ensures any backdated/deleted/altered row is detectable by
    re-computing hashes from genesis row.

    Also adds payload_json + payload_hash for field-level evidence (Tier 3
    of Calendar audit). detail field stays short and human-readable; the full
    snapshot/diff goes in payload_json with SHA-256 of payload in payload_hash."""
    import logging
    try:
        cols = {r[1] for r in db.execute("PRAGMA table_info(audit_log)")}
        if 'prev_hash' not in cols:
            db.execute("ALTER TABLE audit_log ADD COLUMN prev_hash TEXT")
        if 'row_hash' not in cols:
            db.execute("ALTER TABLE audit_log ADD COLUMN row_hash TEXT")
        if 'payload_json' not in cols:
            db.execute("ALTER TABLE audit_log ADD COLUMN payload_json TEXT")
        if 'payload_hash' not in cols:
            db.execute("ALTER TABLE audit_log ADD COLUMN payload_hash TEXT")
        db.commit()
    except Exception as e:
        logging.warning(f'_migrate_audit_hash_chain failed: {e}')


def _migrate_totp(db):
    import logging
    try:
        cols = {r[1] for r in db.execute("PRAGMA table_info(users)")}
        if 'totp_secret' not in cols:
            db.execute("ALTER TABLE users ADD COLUMN totp_secret TEXT")
        if 'totp_enabled' not in cols:
            db.execute("ALTER TABLE users ADD COLUMN totp_enabled INTEGER DEFAULT 0")
        db.commit()
    except Exception as e:
        logging.warning(f'_migrate_totp failed: {e}')


def _migrate_anomaly_flags(db):
    """Add anomaly_flags column to emp_training and programme_details."""
    import logging
    try:
        et_cols = {r[1] for r in db.execute("PRAGMA table_info(emp_training)")}
        if 'anomaly_flags' not in et_cols:
            db.execute("ALTER TABLE emp_training ADD COLUMN anomaly_flags TEXT")
        pd_cols = {r[1] for r in db.execute("PRAGMA table_info(programme_details)")}
        if 'anomaly_flags' not in pd_cols:
            db.execute("ALTER TABLE programme_details ADD COLUMN anomaly_flags TEXT")
        db.commit()
    except Exception as e:
        logging.warning(f'_migrate_anomaly_flags failed: {e}')


def _migrate_session_time(db):
    """Standardise Start Time / End Time fields across emp_training and
    programme_details (2A + 2C). Calendar already has them. Adds columns
    idempotently so existing DBs match fresh-deploy schema."""
    import logging
    try:
        for table in ('emp_training', 'programme_details'):
            cols = {r[1] for r in db.execute(f"PRAGMA table_info({table})")}
            if 'time_from' not in cols:
                db.execute(f"ALTER TABLE {table} ADD COLUMN time_from TEXT")
            if 'time_to' not in cols:
                db.execute(f"ALTER TABLE {table} ADD COLUMN time_to TEXT")
        db.commit()
    except Exception as e:
        logging.warning(f'_migrate_session_time failed: {e}')


def _migrate_category_and_effectiveness(db):
    """Add Category (Specialized/General) on programme_master + calendar.
    Create effectiveness_review table for post-training 3-month manager
    review tracking (SOP: 25% of programmes are Specialized).
    Idempotent."""
    import logging
    try:
        pm_cols = {r[1] for r in db.execute("PRAGMA table_info(programme_master)")}
        if 'category' not in pm_cols:
            db.execute("ALTER TABLE programme_master ADD COLUMN category TEXT DEFAULT 'General'")
        cal_cols = {r[1] for r in db.execute("PRAGMA table_info(calendar)")}
        if 'category' not in cal_cols:
            db.execute("ALTER TABLE calendar ADD COLUMN category TEXT DEFAULT 'General'")
        db.executescript('''
            CREATE TABLE IF NOT EXISTS effectiveness_review (
                id              INTEGER PRIMARY KEY AUTOINCREMENT,
                plant_id        INTEGER NOT NULL,
                session_code    TEXT    NOT NULL,
                emp_code        TEXT    NOT NULL,
                conducted_date  TEXT    NOT NULL,
                due_date        TEXT    NOT NULL,
                completed_date  TEXT,
                rating          INTEGER,
                behaviour_change    TEXT,
                application_on_job  TEXT,
                comments        TEXT,
                filed_by        INTEGER,
                filed_at        TEXT,
                UNIQUE(plant_id, session_code, emp_code)
            );
            CREATE INDEX IF NOT EXISTS idx_eff_plant_status ON effectiveness_review(plant_id, completed_date, due_date);
            CREATE INDEX IF NOT EXISTS idx_eff_session ON effectiveness_review(session_code);
        ''')
        db.commit()
    except Exception as e:
        logging.warning(f'_migrate_category_and_effectiveness failed: {e}')


def _migrate_org_config(db):
    """Idempotent: scoped tenant configuration table + seed global defaults.
    The CREATE statements must succeed (table is required) so we raise on
    schema failure; only per-row seed errors are swallowed."""
    import logging
    db.executescript('''
        CREATE TABLE IF NOT EXISTS org_config (
            scope       TEXT    NOT NULL CHECK(scope IN ('global','plant')),
            plant_id    INTEGER,
            key         TEXT    NOT NULL,
            value       TEXT    NOT NULL,
            value_type  TEXT    NOT NULL DEFAULT 'string'
                        CHECK(value_type IN ('string','int','float','bool','json')),
            label       TEXT    DEFAULT '',
            category    TEXT    DEFAULT 'general',
            updated_at  TEXT,
            updated_by  TEXT
        );
        CREATE UNIQUE INDEX IF NOT EXISTS ux_org_config_scope_plant_key
            ON org_config(scope, COALESCE(plant_id,-1), key);
        CREATE INDEX IF NOT EXISTS idx_org_config_key ON org_config(key, scope);
    ''')
    try:
        from tms.config import CONFIG_DEFAULTS
        db.execute('BEGIN IMMEDIATE')
        for key, default, vtype, label, cat in CONFIG_DEFAULTS:
            db.execute(
                "INSERT OR IGNORE INTO org_config(scope, plant_id, key, value, value_type, label, category) "
                "VALUES('global', NULL, ?, ?, ?, ?, ?)",
                (key, default, vtype, label, cat)
            )
        db.commit()
        logging.info('org_config seeded %d default(s)', len(CONFIG_DEFAULTS))
    except Exception:
        try: db.rollback()
        except Exception: pass
        logging.exception('org_config seed failed (table exists, defaults skipped)')


def _migrate_calendar_verification(db):
    """Add verification audit columns to calendar table."""
    import logging
    try:
        cols = {r[1] for r in db.execute("PRAGMA table_info(calendar)")}
        adds = [
            ('conducted_at', 'TEXT'),
            ('conducted_by', 'INTEGER'),
            ('verified_at',  'TEXT'),
            ('verified_by',  'INTEGER'),
            ('actual_pax',   'INTEGER DEFAULT 0'),
            ('actual_hrs',   'REAL DEFAULT 0'),
        ]
        for name, typ in adds:
            if name not in cols:
                db.execute(f"ALTER TABLE calendar ADD COLUMN {name} {typ}")
        db.commit()
    except Exception as e:
        logging.warning(f'_migrate_calendar_verification failed: {e}')


def _migrate_verification_log(db):
    """Create verification_log table for audit chain of session lifecycle."""
    import logging
    try:
        db.executescript('''
            CREATE TABLE IF NOT EXISTS verification_log (
                id           INTEGER PRIMARY KEY AUTOINCREMENT,
                session_code TEXT    NOT NULL,
                plant_id     INTEGER NOT NULL,
                stage        TEXT    NOT NULL,
                actor        TEXT,
                actor_id     INTEGER,
                ts           TEXT    DEFAULT (datetime('now','localtime')),
                detail       TEXT
            );
            CREATE INDEX IF NOT EXISTS idx_vlog_session ON verification_log(session_code, plant_id);
            CREATE INDEX IF NOT EXISTS idx_vlog_stage   ON verification_log(stage);
        ''')
        db.commit()
    except Exception as e:
        logging.warning(f'_migrate_verification_log failed: {e}')


def _migrate_session_qr_fk(db):
    """Rebuild session_qr if created_by FK lacks ON DELETE SET NULL — prevents orphan QR rows."""
    import logging
    try:
        row = db.execute("SELECT sql FROM sqlite_master WHERE type='table' AND name='session_qr'").fetchone()
        if not row or not row[0]:
            return
        ddl = row[0]
        if 'ON DELETE SET NULL' in ddl.upper() or 'ON DELETE CASCADE' in ddl.upper():
            return
        # Idempotency guard: drop any orphaned __new table from a previously crashed run
        # before recreating it, otherwise the CREATE below fails with "table already exists".
        orphan = db.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='session_qr__new'").fetchone()
        if orphan:
            db.executescript('DROP TABLE IF EXISTS session_qr__new;')
        db.executescript('''
            PRAGMA foreign_keys=OFF;
            BEGIN;
            CREATE TABLE session_qr__new (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                plant_id INTEGER NOT NULL REFERENCES plants(id),
                session_code TEXT NOT NULL,
                token TEXT NOT NULL UNIQUE,
                stage TEXT NOT NULL DEFAULT 'attendance' CHECK(stage IN ('attendance','feedback')),
                created_at TEXT,
                expires_at TEXT,
                is_active INTEGER DEFAULT 1,
                created_by INTEGER REFERENCES users(id) ON DELETE SET NULL
            );
            INSERT INTO session_qr__new (id, plant_id, session_code, token, stage, created_at, expires_at, is_active, created_by)
                SELECT id, plant_id, session_code, token, stage, created_at, expires_at, is_active, created_by FROM session_qr;
            DROP TABLE session_qr;
            ALTER TABLE session_qr__new RENAME TO session_qr;
            CREATE INDEX IF NOT EXISTS idx_qr_token   ON session_qr(token);
            CREATE INDEX IF NOT EXISTS idx_qr_session ON session_qr(plant_id, session_code);
            COMMIT;
            PRAGMA foreign_keys=ON;
        ''')
    except Exception as e:
        logging.warning(f'_migrate_session_qr_fk failed: {e}')


def _migrate_spoc_requests(db):
    import logging
    try:
        db.executescript('''
            CREATE TABLE IF NOT EXISTS spoc_requests (
                id           INTEGER PRIMARY KEY AUTOINCREMENT,
                ts           TEXT    DEFAULT (datetime('now','localtime')),
                plant_id     INTEGER NOT NULL,
                requested_by TEXT    NOT NULL,
                request_type TEXT    NOT NULL CHECK(request_type IN ('TNI_ADD','MARK_CONDUCTED','MANUAL_ATTENDANCE','OTHER')),
                details      TEXT    NOT NULL,
                status       TEXT    DEFAULT 'Pending' CHECK(status IN ('Pending','Approved','Rejected')),
                reviewed_by  TEXT,
                reviewed_at  TEXT,
                review_note  TEXT
            );
            CREATE INDEX IF NOT EXISTS idx_spoc_req_plant  ON spoc_requests(plant_id, status);
            CREATE INDEX IF NOT EXISTS idx_spoc_req_status ON spoc_requests(status);
        ''')
        db.commit()
    except Exception as e:
        logging.warning(f'_migrate_spoc_requests failed: {e}')

    # Phase 2: payload_json column for typed request executor (idempotent ALTER).
    try:
        cols = {r[1] for r in db.execute('PRAGMA table_info(spoc_requests)').fetchall()}
        if 'payload_json' not in cols:
            db.execute('ALTER TABLE spoc_requests ADD COLUMN payload_json TEXT')
            db.commit()
    except Exception as e:
        logging.warning(f'_migrate_spoc_requests payload_json failed: {e}')


def _dedupe_tni_prog_variants(db):
    """Merge high-confidence spelling/plural variants of the SAME programme into one
    canonical name across every table that stores it (tni, programme_master, calendar,
    programme_details, emp_training). Companion to the _tni_canon_candidates guard that
    now PREVENTS new variants at entry — this cleans variants that predate the guard.

    Safety: only merges names that are >=0.90 similar AND share the same digit groups
    AND the same roman-numeral tokens. So 'Communication Skill(s)', 'Behaviour(al) Safety',
    'Storage'/'Store' merge, but 'Level 1'/'Level 2', 'Phase I'/'Phase II',
    'ISO 9001'/'ISO 14001' are NEVER merged (genuinely distinct).

    Idempotent — once clean it finds nothing. Runs on every deploy (incl. production).
    """
    import re
    from difflib import SequenceMatcher
    from collections import defaultdict
    try:
        rows = db.execute('SELECT plant_id, prog_type, programme_name, COUNT(*) c '
                          'FROM tni GROUP BY plant_id, prog_type, programme_name').fetchall()
    except Exception as e:
        logging.warning(f'_dedupe_tni_prog_variants skipped: {e}')
        return

    ROMAN = {'i', 'ii', 'iii', 'iv', 'v', 'vi', 'vii', 'viii', 'ix', 'x', 'xi', 'xii'}
    def _nums(s):   return tuple(re.findall(r'\d+', s.lower()))
    def _romans(s): return frozenset(t for t in re.findall(r'[a-z]+', s.lower()) if t in ROMAN)

    groups = defaultdict(list)
    for r in rows:
        groups[(r['plant_id'], r['prog_type'])].append((r['programme_name'], r['c']))

    plan = []  # (plant_id, variant, canonical)
    for (pid, _pt), items in groups.items():
        names = [n for n, _ in items]; cnt = dict(items); used = set()
        for i, a in enumerate(names):
            if not a or a in used:
                continue
            clust = [a]
            for b in names[i + 1:]:
                if not b or b in used:
                    continue
                if (SequenceMatcher(None, a.lower(), b.lower()).ratio() >= 0.90
                        and _nums(a) == _nums(b) and _romans(a) == _romans(b)):
                    clust.append(b); used.add(b)
            if len(clust) > 1:
                canon = max(clust, key=lambda n: (cnt[n], -len(n)))  # most frequent, then shortest
                for n in clust:
                    if n != canon:
                        plan.append((pid, n, canon))
    if not plan:
        return
    try:
        for pid, var, canon in plan:
            # tni: drop rows that would collide on UNIQUE(plant,emp,prog,fy), then re-point the rest
            db.execute('DELETE FROM tni WHERE plant_id=? AND programme_name=? AND EXISTS('
                       'SELECT 1 FROM tni x WHERE x.plant_id=tni.plant_id AND x.emp_code=tni.emp_code '
                       'AND x.fy_year=tni.fy_year AND x.programme_name=?)', (pid, var, canon))
            db.execute('UPDATE tni SET programme_name=? WHERE plant_id=? AND programme_name=?', (canon, pid, var))
            # programme_master: drop colliding variant row on UNIQUE(plant,name), then re-point
            db.execute('DELETE FROM programme_master WHERE plant_id=? AND name=? AND EXISTS('
                       'SELECT 1 FROM programme_master x WHERE x.plant_id=programme_master.plant_id '
                       'AND x.name=?)', (pid, var, canon))
            db.execute('UPDATE programme_master SET name=? WHERE plant_id=? AND name=?', (canon, pid, var))
            db.execute('UPDATE calendar SET programme_name=? WHERE plant_id=? AND programme_name=?', (canon, pid, var))
            db.execute('UPDATE programme_details SET programme_name=? WHERE plant_id=? AND programme_name=?', (canon, pid, var))
            db.execute('UPDATE emp_training SET programme_name=? WHERE plant_id=? AND programme_name=?', (canon, pid, var))
        db.commit()
        logging.info(f'_dedupe_tni_prog_variants: merged {len(plan)} spelling variant(s) into canonical names')
    except Exception as e:
        logging.warning(f'_dedupe_tni_prog_variants failed: {e}')


def init_db():
    from tms.helpers import (
        _cleanse_master_spelling, _cleanse_programme_names,
        _cleanse_emp_fields, _cleanup_stale_analyze_files
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
    _migrate_session_qr_fk(db)
    _migrate_session_pin(db)
    _migrate_central_plant(db)
    _migrate_calendar_is_central(db)
    _migrate_mode_offline_to_classroom(db)
    _migrate_feedback_scale_to_4(db)
    _migrate_emp_training_host(db)
    _migrate_corp_members(db)
    _migrate_central_user_plant(db)
    _migrate_audit_lockout(db)
    _migrate_audit_hash_chain(db)
    _migrate_totp(db)
    _migrate_force_default_password_change(db)
    _migrate_spoc_requests(db)
    _migrate_calendar_verification(db)
    _migrate_verification_log(db)
    _migrate_anomaly_flags(db)
    _migrate_session_time(db)
    _migrate_category_and_effectiveness(db)
    _migrate_org_config(db)
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
    _dedupe_tni_prog_variants(db)   # merge pre-guard spelling/plural duplicates (prod-safe, idempotent)
    try:
        _cleanse_emp_fields(db)
    except Exception as e:
        logging.warning(f'_cleanse_emp_fields failed: {e}')
    db.commit()
    db.close()
    _cleanup_stale_analyze_files()
