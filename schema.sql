-- BCML Training Management System — Database Schema

CREATE TABLE IF NOT EXISTS plants (
    id INTEGER PRIMARY KEY,
    name TEXT NOT NULL,
    unit_code TEXT NOT NULL UNIQUE
);

CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT NOT NULL UNIQUE,
    password TEXT NOT NULL,
    role TEXT NOT NULL CHECK(role IN ('spoc','central','admin')),
    plant_id INTEGER REFERENCES plants(id),
    must_change_password INTEGER DEFAULT 1,
    failed_attempts INTEGER DEFAULT 0,
    locked_until TEXT
);

CREATE TABLE IF NOT EXISTS employees (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    plant_id INTEGER NOT NULL REFERENCES plants(id),
    emp_code TEXT NOT NULL,
    name TEXT NOT NULL,
    designation TEXT,
    grade TEXT,
    collar TEXT,
    department TEXT,
    section TEXT,
    category TEXT,
    gender TEXT,
    physically_handicapped TEXT DEFAULT 'No',
    exit_date TEXT,
    exit_reason TEXT,
    remarks TEXT,
    is_active INTEGER DEFAULT 1,
    UNIQUE(plant_id, emp_code)
);

CREATE TABLE IF NOT EXISTS tni (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    plant_id INTEGER NOT NULL REFERENCES plants(id),
    emp_code TEXT NOT NULL,
    programme_name TEXT NOT NULL,
    prog_type TEXT,
    mode TEXT,
    target_month TEXT,
    planned_hours REAL DEFAULT 0,
    source TEXT DEFAULT 'TNI Driven',
    fy_year TEXT NOT NULL DEFAULT '',
    created_at TEXT DEFAULT (date('now')),
    UNIQUE(plant_id, emp_code, programme_name, fy_year)
);

CREATE TABLE IF NOT EXISTS calendar (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    plant_id INTEGER NOT NULL REFERENCES plants(id),
    prog_code TEXT,
    session_code TEXT UNIQUE,
    source TEXT DEFAULT 'TNI',
    programme_name TEXT NOT NULL,
    prog_type TEXT,
    planned_month TEXT,
    plan_start TEXT,
    plan_end TEXT,
    time_from TEXT,
    time_to TEXT,
    duration_hrs REAL DEFAULT 0,
    level TEXT,
    mode TEXT,
    target_audience TEXT,
    planned_pax INTEGER DEFAULT 0,
    trainer_vendor TEXT,
    status TEXT DEFAULT 'To Be Planned',
    session_pin TEXT,
    is_central INTEGER NOT NULL DEFAULT 0,
    created_at TEXT DEFAULT (date('now'))
);

CREATE TABLE IF NOT EXISTS emp_training (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    plant_id INTEGER NOT NULL REFERENCES plants(id),
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
    host_plant_id INTEGER,
    anomaly_flags TEXT,
    created_at TEXT DEFAULT (date('now'))
);

CREATE TABLE IF NOT EXISTS programme_details (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    plant_id INTEGER NOT NULL REFERENCES plants(id),
    session_code TEXT NOT NULL,
    programme_name TEXT,
    prog_type TEXT,
    level TEXT,
    cal_new TEXT,
    mode TEXT,
    start_date TEXT,
    end_date TEXT,
    audience TEXT,
    hours_actual REAL DEFAULT 0,
    faculty_name TEXT,
    int_ext TEXT,
    cost REAL DEFAULT 0,
    venue TEXT,
    course_feedback REAL,
    faculty_feedback REAL,
    trainer_fb_participants REAL,
    trainer_fb_facilities REAL,
    anomaly_flags TEXT,
    created_at TEXT DEFAULT (date('now'))
);

CREATE TABLE IF NOT EXISTS session_qr (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    plant_id INTEGER NOT NULL REFERENCES plants(id),
    session_code TEXT NOT NULL,
    token TEXT NOT NULL UNIQUE,
    stage TEXT NOT NULL DEFAULT 'attendance' CHECK(stage IN ('attendance','feedback')),
    created_at TEXT DEFAULT (datetime('now','localtime')),
    expires_at TEXT,
    is_active INTEGER DEFAULT 1,
    created_by INTEGER REFERENCES users(id)
);

CREATE TABLE IF NOT EXISTS feedback_response (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    plant_id INTEGER NOT NULL REFERENCES plants(id),
    session_code TEXT NOT NULL,
    emp_code TEXT,
    submitted_at TEXT DEFAULT (datetime('now','localtime')),
    q_obj_explained INTEGER,
    q_well_structured INTEGER,
    q_content_appropriate INTEGER,
    q_presentation_quality INTEGER,
    q_time_reasonable INTEGER,
    q_inputs_appropriate INTEGER,
    q_communication_clear INTEGER,
    q_queries_responded INTEGER,
    q_well_involved INTEGER,
    key_learnings TEXT,
    suggestions TEXT,
    ip_address TEXT,
    lang TEXT DEFAULT 'en',
    UNIQUE(plant_id, session_code, emp_code)
);

CREATE TABLE IF NOT EXISTS tni_archive (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    archive_token TEXT NOT NULL,
    archived_at TEXT NOT NULL,
    plant_id INTEGER NOT NULL,
    emp_code TEXT NOT NULL,
    programme_name TEXT NOT NULL,
    prog_type TEXT,
    mode TEXT,
    target_month TEXT,
    planned_hours REAL DEFAULT 0,
    source TEXT DEFAULT 'TNI Driven',
    fy_year TEXT NOT NULL DEFAULT ''
);

CREATE TABLE IF NOT EXISTS programme_master (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    plant_id INTEGER NOT NULL REFERENCES plants(id),
    name TEXT NOT NULL,
    prog_type TEXT,
    mode TEXT,
    source TEXT DEFAULT 'TNI Requirement',
    created_at TEXT DEFAULT (date('now')),
    UNIQUE(plant_id, name)
);

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

CREATE INDEX IF NOT EXISTS idx_qr_token    ON session_qr(token);
CREATE INDEX IF NOT EXISTS idx_qr_session  ON session_qr(plant_id, session_code);
CREATE INDEX IF NOT EXISTS idx_fr_session  ON feedback_response(plant_id, session_code);
CREATE INDEX IF NOT EXISTS idx_tni_archive_token ON tni_archive(archive_token);
CREATE INDEX IF NOT EXISTS idx_tni_archive_plant ON tni_archive(plant_id, fy_year);
CREATE INDEX IF NOT EXISTS idx_prog_master_plant ON programme_master(plant_id);
CREATE INDEX IF NOT EXISTS idx_emp_plant ON employees(plant_id, is_active);
CREATE INDEX IF NOT EXISTS idx_emp_code ON employees(emp_code);
CREATE INDEX IF NOT EXISTS idx_emp_plant_code ON employees(plant_id, emp_code);
CREATE INDEX IF NOT EXISTS idx_tni_plant ON tni(plant_id);
CREATE INDEX IF NOT EXISTS idx_tni_dedup ON tni(plant_id, emp_code, programme_name);
CREATE INDEX IF NOT EXISTS idx_tni_prog ON tni(plant_id, programme_name);
CREATE INDEX IF NOT EXISTS idx_cal_plant ON calendar(plant_id);
CREATE INDEX IF NOT EXISTS idx_training_plant ON emp_training(plant_id);
CREATE INDEX IF NOT EXISTS idx_et_lookup ON emp_training(plant_id, emp_code, programme_name);
CREATE INDEX IF NOT EXISTS idx_corp_active ON corp_members(is_active, name);
CREATE INDEX IF NOT EXISTS idx_prog_plant ON programme_details(plant_id);

CREATE TABLE IF NOT EXISTS audit_log (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    ts          TEXT    DEFAULT (datetime('now','localtime')),
    user_id     INTEGER,
    username    TEXT,
    plant_id    INTEGER,
    action      TEXT    NOT NULL,
    detail      TEXT,
    ip_address  TEXT,
    prev_hash   TEXT,
    row_hash    TEXT
);
CREATE INDEX IF NOT EXISTS idx_audit_ts   ON audit_log(ts);
CREATE INDEX IF NOT EXISTS idx_audit_user ON audit_log(username);


CREATE TABLE IF NOT EXISTS spoc_requests (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    ts           TEXT    DEFAULT (datetime('now','localtime')),
    plant_id     INTEGER NOT NULL REFERENCES plants(id),
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

CREATE TABLE IF NOT EXISTS planner_entries (
    id                INTEGER PRIMARY KEY AUTOINCREMENT,
    plant_id          INTEGER NOT NULL,
    fy_year           TEXT    NOT NULL,            -- '2026-27'
    plan_month        TEXT    NOT NULL,            -- '2026-06' (YYYY-MM)
    programme_name    TEXT    NOT NULL,
    target_sessions   INTEGER NOT NULL DEFAULT 0,
    pax_per_session   INTEGER NOT NULL DEFAULT 20,
    hours_per_session REAL    NOT NULL DEFAULT 4,
    faculty           TEXT,
    audience          TEXT,                        -- BC / WC / Common (auto from TNI)
    notes             TEXT,
    status            TEXT    DEFAULT 'draft',     -- draft | locked
    created_by        TEXT,
    created_at        TEXT    DEFAULT (datetime('now','localtime')),
    updated_at        TEXT    DEFAULT (datetime('now','localtime')),
    locked_at         TEXT,
    locked_by         TEXT,
    UNIQUE(plant_id, plan_month, programme_name)
);
CREATE INDEX IF NOT EXISTS idx_planner_plant_month ON planner_entries(plant_id, plan_month);
CREATE INDEX IF NOT EXISTS idx_planner_fy          ON planner_entries(plant_id, fy_year);

CREATE TABLE IF NOT EXISTS planner_audit (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    ts          TEXT    DEFAULT (datetime('now','localtime')),
    plant_id    INTEGER NOT NULL,
    plan_month  TEXT,
    actor       TEXT,
    action      TEXT    NOT NULL,                  -- save_draft | lock_month | lock_fy | edit_locked | acknowledge_gap
    detail      TEXT
);
CREATE INDEX IF NOT EXISTS idx_planner_audit_plant ON planner_audit(plant_id, ts);

CREATE TABLE IF NOT EXISTS tni_upload_errors (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    ts              TEXT    DEFAULT (datetime('now','localtime')),
    plant_id        INTEGER NOT NULL,
    username        TEXT,
    aid             TEXT,
    row_status      TEXT,
    row_num         INTEGER,
    emp_code        TEXT,
    emp_name        TEXT,
    programme_name  TEXT,
    prog_type       TEXT,
    mode            TEXT,
    planned_hours   REAL,
    issues          TEXT,
    garbage_class   TEXT
);
CREATE INDEX IF NOT EXISTS idx_tue_ts    ON tni_upload_errors(ts);
CREATE INDEX IF NOT EXISTS idx_tue_plant ON tni_upload_errors(plant_id, ts);
