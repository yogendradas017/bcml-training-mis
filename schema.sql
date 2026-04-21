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
    must_change_password INTEGER DEFAULT 1
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
    created_at TEXT DEFAULT (date('now'))
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
    created_at TEXT DEFAULT (date('now'))
);

CREATE TABLE IF NOT EXISTS programme_master (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    plant_id INTEGER NOT NULL REFERENCES plants(id),
    name TEXT NOT NULL,
    prog_type TEXT,
    mode TEXT,
    created_at TEXT DEFAULT (date('now')),
    UNIQUE(plant_id, name)
);

CREATE INDEX IF NOT EXISTS idx_prog_master_plant ON programme_master(plant_id);
CREATE INDEX IF NOT EXISTS idx_emp_plant ON employees(plant_id, is_active);
CREATE INDEX IF NOT EXISTS idx_emp_code ON employees(emp_code);
CREATE INDEX IF NOT EXISTS idx_tni_plant ON tni(plant_id);
CREATE INDEX IF NOT EXISTS idx_cal_plant ON calendar(plant_id);
CREATE INDEX IF NOT EXISTS idx_training_plant ON emp_training(plant_id);
CREATE INDEX IF NOT EXISTS idx_prog_plant ON programme_details(plant_id);
