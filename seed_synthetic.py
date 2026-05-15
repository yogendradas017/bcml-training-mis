"""
Seed 6 months (Apr–Sep 2026 = Q1+Q2 FY 26-27) of synthetic training data
for all 10 plants. Populates: programme_master, tni, calendar, emp_training,
programme_details.

Run: python seed_synthetic.py
"""
import sqlite3, random, math
from datetime import date, timedelta

DB_PATH = 'data/training.db'
random.seed(42)

PLANTS = [
    {'id': 1,  'name': 'Balrampur',  'unit_code': 'BCM'},
    {'id': 2,  'name': 'Babhnan',    'unit_code': 'BBN'},
    {'id': 3,  'name': 'Rauzagaon',  'unit_code': 'RCM'},
    {'id': 4,  'name': 'Maizapur',   'unit_code': 'MZP'},
    {'id': 5,  'name': 'Mankapur',   'unit_code': 'MCM'},
    {'id': 6,  'name': 'Gularia',    'unit_code': 'GCM'},
    {'id': 7,  'name': 'Tulsipur',   'unit_code': 'TCM'},
    {'id': 8,  'name': 'Kumbhi',     'unit_code': 'KCM'},
    {'id': 9,  'name': 'Haidergarh', 'unit_code': 'HCM'},
    {'id': 10, 'name': 'Akbarpur',   'unit_code': 'ACM'},
]

# Canonical programmes — name → (prog_type, audience)
PROGRAMMES = [
    # Technical
    ('5-S Management',                 'Technical',              'Common'),
    ('AC Maintenance',                 'Technical',              'Blue Collared'),
    ('Boiler Operation Safety',        'Technical',              'Blue Collared'),
    ('Electrical Safety',              'Technical',              'Blue Collared'),
    ('Fire Safety & Prevention',       'Technical',              'Common'),
    ('Hydraulic System Maintenance',   'Technical',              'Blue Collared'),
    ('Instrumentation & Control',      'Technical',              'Blue Collared'),
    ('Mechanical Preventive Maintenance','Technical',            'Blue Collared'),
    ('Pump & Compressor Maintenance',  'Technical',              'Blue Collared'),
    ('Sugar Process Technology',       'Technical',              'Blue Collared'),
    ('Water Treatment Operations',     'Technical',              'Blue Collared'),
    ('DCS & PLC Operations',           'Technical',              'Blue Collared'),
    # EHS/HR
    ('Accident Investigation',         'EHS/HR',                 'Common'),
    ('Behavioural Safety',             'EHS/HR',                 'Blue Collared'),
    ('Chemical Handling & Storage',    'EHS/HR',                 'Blue Collared'),
    ('Environmental Compliance',       'EHS/HR',                 'White Collared'),
    ('First Aid & Emergency Response', 'EHS/HR',                 'Common'),
    ('POSH Awareness',                 'EHS/HR',                 'Common'),
    ('PPE Awareness & Usage',          'EHS/HR',                 'Blue Collared'),
    ('Risk Assessment Techniques',     'EHS/HR',                 'White Collared'),
    ('ISO 14001 Awareness',            'EHS/HR',                 'White Collared'),
    # IT
    ('Advance Excel',                  'IT',                     'White Collared'),
    ('AI Prompting',                   'IT',                     'White Collared'),
    ('Cybersecurity Awareness',        'IT',                     'Common'),
    ('ERP System Usage',               'IT',                     'White Collared'),
    ('SAP Basics',                     'IT',                     'White Collared'),
    # Cane
    ('Cane Development Practices',     'Cane',                   'Blue Collared'),
    ('Cane Quality Assessment',        'Cane',                   'Blue Collared'),
    ('Drip Irrigation Technology',     'Cane',                   'Blue Collared'),
    ('Pest & Disease Management',      'Cane',                   'Blue Collared'),
    ('Ratoon Crop Management',         'Cane',                   'Blue Collared'),
    # Commercial
    ('Commercial Law Basics',          'Commercial',             'White Collared'),
    ('Contract Management',            'Commercial',             'White Collared'),
    ('GST & Taxation Update',          'Commercial',             'White Collared'),
    ('Purchase & Procurement Process', 'Commercial',             'White Collared'),
    # Behavioural/Leadership
    ('Communication Skills',           'Behavioural/Leadership', 'Common'),
    ('Leadership Development',         'Behavioural/Leadership', 'White Collared'),
    ('Problem Solving & Decision Making','Behavioural/Leadership','White Collared'),
    ('Team Building',                  'Behavioural/Leadership', 'Common'),
    ('Time Management',                'Behavioural/Leadership', 'White Collared'),
    ('Supervisory Skills',             'Behavioural/Leadership', 'Blue Collared'),
]

TYPE_ABBR = {
    'Technical': 'TEC', 'EHS/HR': 'EHS', 'IT': 'IT',
    'Cane': 'CAN', 'Commercial': 'COM', 'Behavioural/Leadership': 'BEH',
}

# 6-month window: Apr–Sep 2026 (Q1+Q2 FY 26-27)
MONTHS = [
    ('April 2026',     date(2026, 4, 1),  date(2026, 4, 30)),
    ('May 2026',       date(2026, 5, 1),  date(2026, 5, 31)),
    ('June 2026',      date(2026, 6, 1),  date(2026, 6, 30)),
    ('July 2026',      date(2026, 7, 1),  date(2026, 7, 31)),
    ('August 2026',    date(2026, 8, 1),  date(2026, 8, 31)),
    ('September 2026', date(2026, 9, 1),  date(2026, 9, 30)),
]


def rand_weekday(m_start, m_end):
    """Pick a random weekday within the month range."""
    delta = (m_end - m_start).days
    for _ in range(50):
        d = m_start + timedelta(days=random.randint(0, delta))
        if d.weekday() < 5:  # Mon–Fri
            return d
    return m_start


def seed(db):
    # ── 1. Programme Master — insert for all plants ──────────────────────
    print("Seeding programme_master...")
    for plant in PLANTS:
        for (pname, ptype, _audience) in PROGRAMMES:
            db.execute(
                "INSERT OR IGNORE INTO programme_master(plant_id, name, prog_type, mode, source) "
                "VALUES(?,?,?,?,?)",
                (plant['id'], pname, ptype, 'Offline', 'TNI Driven')
            )
    db.commit()
    print(f"  Done: {len(PROGRAMMES)} programmes × 10 plants")

    # ── 2. TNI — each BC-eligible employee gets BC programmes, WC gets WC ──
    print("Seeding TNI...")
    tni_count = 0
    bc_progs = [(n, t) for (n, t, a) in PROGRAMMES if a in ('Blue Collared', 'Common')]
    wc_progs = [(n, t) for (n, t, a) in PROGRAMMES if a in ('White Collared', 'Common')]

    for plant in PLANTS:
        pid = plant['id']
        # Get employees (limit to keep dataset manageable)
        emps_bc = db.execute(
            "SELECT emp_code FROM employees WHERE plant_id=? AND is_active=1 AND collar='Blue Collared'",
            (pid,)).fetchall()
        emps_wc = db.execute(
            "SELECT emp_code FROM employees WHERE plant_id=? AND is_active=1 AND collar='White Collared'",
            (pid,)).fetchall()

        # Each employee nominated for 4–8 programmes
        for row in emps_bc:
            emp = row[0]
            nominated = random.sample(bc_progs, min(random.randint(4, 8), len(bc_progs)))
            for (pname, ptype) in nominated:
                try:
                    db.execute(
                        "INSERT OR IGNORE INTO tni(plant_id, emp_code, programme_name, prog_type, source, fy_year) "
                        "VALUES(?,?,?,?,?,?)",
                        (pid, emp, pname, ptype, 'TNI Driven', '26-27'))
                    tni_count += 1
                except Exception:
                    pass

        for row in emps_wc:
            emp = row[0]
            nominated = random.sample(wc_progs, min(random.randint(3, 6), len(wc_progs)))
            for (pname, ptype) in nominated:
                try:
                    db.execute(
                        "INSERT OR IGNORE INTO tni(plant_id, emp_code, programme_name, prog_type, source, fy_year) "
                        "VALUES(?,?,?,?,?,?)",
                        (pid, emp, pname, ptype, 'TNI Driven', '26-27'))
                    tni_count += 1
                except Exception:
                    pass

        db.commit()
        print(f"  Plant {plant['name']}: TNI seeded")

    print(f"  Total TNI rows: {tni_count}")

    # ── 3. Calendar + emp_training + programme_details ───────────────────
    print("Seeding calendar + training records...")
    session_counters = {}  # (plant_id, type_abbr) → int
    cal_total = 0
    et_total  = 0
    pd_total  = 0

    for plant in PLANTS:
        pid  = plant['id']
        uc   = plant['unit_code']

        # Get employees by collar for this plant
        bc_emps = [r[0] for r in db.execute(
            "SELECT emp_code FROM employees WHERE plant_id=? AND is_active=1 AND collar='Blue Collared'",
            (pid,)).fetchall()]
        wc_emps = [r[0] for r in db.execute(
            "SELECT emp_code FROM employees WHERE plant_id=? AND is_active=1 AND collar='White Collared'",
            (pid,)).fetchall()]

        # Pick ~20 programmes to run across 6 months (mix of types)
        selected_progs = random.sample(PROGRAMMES, min(20, len(PROGRAMMES)))

        # Distribute: ~4-5 sessions per month
        sessions_plan = []
        for mi, (mname, m_start, m_end) in enumerate(MONTHS):
            # 4-5 sessions per month, choose from selected programmes
            month_progs = random.sample(selected_progs, random.randint(4, 5))
            for (pname, ptype, audience) in month_progs:
                start_d = rand_weekday(m_start, m_end)
                dur = random.choice([8, 8, 8, 16, 24])  # 1-day, 1-day, 1-day, 2-day, 3-day
                end_d = start_d + timedelta(days=max(0, dur // 8 - 1))
                sessions_plan.append({
                    'month': mname,
                    'pname': pname, 'ptype': ptype, 'audience': audience,
                    'start': start_d, 'end': end_d, 'dur': float(dur),
                })

        # Assign session codes and insert calendar
        for s in sessions_plan:
            ta = TYPE_ABBR.get(s['ptype'], 'GEN')
            key = (pid, ta)
            session_counters[key] = session_counters.get(key, 0) + 1
            seq = session_counters[key]
            prog_code    = f"{uc}/{ta}/{seq:03d}"
            session_code = f"{prog_code}/26-27/B01"

            # 75% conducted (past months fully, current partially)
            is_past = s['start'] <= date(2026, 5, 14)
            conducted = is_past or (random.random() < 0.5)
            status = 'Conducted' if conducted else 'To Be Planned'

            # Planned pax
            if s['audience'] == 'Blue Collared':
                pax = random.randint(20, 45)
            elif s['audience'] == 'White Collared':
                pax = random.randint(8, 20)
            else:
                pax = random.randint(15, 35)

            db.execute(
                "INSERT OR IGNORE INTO calendar("
                "plant_id, prog_code, session_code, source, programme_name, prog_type,"
                "planned_month, plan_start, plan_end, duration_hrs, level, mode,"
                "target_audience, planned_pax, trainer_vendor, status, is_central"
                ") VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                (pid, prog_code, session_code, 'TNI Driven', s['pname'], s['ptype'],
                 s['month'], str(s['start']), str(s['end']),
                 s['dur'], 'Basic', 'Offline',
                 s['audience'], pax, 'Internal Faculty', status, 0)
            )
            cal_total += 1

            if not conducted:
                continue

            # programme_details
            fb_course  = round(random.uniform(3.5, 5.0), 1)
            fb_faculty = round(random.uniform(3.8, 5.0), 1)
            db.execute(
                "INSERT OR IGNORE INTO programme_details("
                "plant_id, session_code, programme_name, prog_type, level, cal_new, mode,"
                "start_date, end_date, audience, hours_actual, faculty_name, int_ext,"
                "cost, venue, course_feedback, faculty_feedback"
                ") VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                (pid, session_code, s['pname'], s['ptype'], 'Basic', 'Calendar Program',
                 'Offline', str(s['start']), str(s['end']),
                 s['audience'], s['dur'],
                 random.choice(['Internal Faculty', 'External Trainer']),
                 random.choice(['Internal', 'External']),
                 random.choice([0, 0, 5000, 10000, 15000, 25000]),
                 f"{plant['name']} Training Hall",
                 fb_course, fb_faculty)
            )
            pd_total += 1

            # Attendees
            if s['audience'] == 'Blue Collared':
                pool = bc_emps
            elif s['audience'] == 'White Collared':
                pool = wc_emps
            else:
                bc_n = max(1, pax * 2 // 3)
                wc_n = max(1, pax // 3)
                pool = random.sample(bc_emps, min(bc_n, len(bc_emps))) + \
                       random.sample(wc_emps, min(wc_n, len(wc_emps)))

            actual_pax = min(pax, len(pool), random.randint(max(5, pax - 10), pax + 5))
            attendees  = random.sample(pool, min(actual_pax, len(pool)))

            for emp in attendees:
                pre  = round(random.uniform(2.0, 3.5), 1)
                post = round(random.uniform(3.5, 5.0), 1)
                db.execute(
                    "INSERT OR IGNORE INTO emp_training("
                    "plant_id, emp_code, session_code, programme_name, start_date, end_date,"
                    "hrs, prog_type, level, mode, cal_new, pre_rating, post_rating,"
                    "venue, month, host_plant_id"
                    ") VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                    (pid, emp, session_code, s['pname'],
                     str(s['start']), str(s['end']),
                     s['dur'], s['ptype'], 'Basic', 'Offline', 'Calendar Program',
                     pre, post,
                     f"{plant['name']} Training Hall",
                     s['month'], pid)
                )
                et_total += 1

        db.commit()
        print(f"  {plant['name']}: {len(sessions_plan)} sessions")

    return cal_total, et_total, pd_total


def main():
    db = sqlite3.connect(DB_PATH)
    db.execute("PRAGMA foreign_keys = OFF")
    db.execute("PRAGMA journal_mode = WAL")

    # Wipe existing synthetic data (keep employees & real TNI for Plant 1)
    print("Clearing old calendar/training data...")
    db.execute("DELETE FROM emp_training")
    db.execute("DELETE FROM programme_details")
    db.execute("DELETE FROM calendar")
    # Clear TNI for plants 2-10 (plant 1 has real TNI)
    db.execute("DELETE FROM tni WHERE plant_id != 1")
    # Clear programme_master for plants 2-10
    db.execute("DELETE FROM programme_master WHERE plant_id != 1")
    db.commit()

    cal, et, pd_ = seed(db)
    db.close()

    print(f"\n{'='*50}")
    print(f"SYNTHETIC DATA SEEDED")
    print(f"  Calendar sessions : {cal}")
    print(f"  Emp training rows : {et}")
    print(f"  Programme details : {pd_}")
    print(f"{'='*50}")
    print("Refresh the app — central dashboard should now show live data.")


if __name__ == '__main__':
    main()
