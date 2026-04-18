"""
Run once after app setup to import all 10 plant employee data.
Usage:  python seed.py
"""
import os, re, sqlite3
import pandas as pd
from werkzeug.security import generate_password_hash

BASE_DIR  = os.path.dirname(os.path.abspath(__file__))
DB_PATH   = os.path.join(BASE_DIR, 'data', 'training.db')
DATA_DIR  = r"C:\Users\yogendra.das\Desktop\Master Emp Data"

PLANTS = [
    {'id': 1,  'name': 'Balrampur',  'unit_code': 'BCM', 'file': 'Balrampur 26-27.xlsx'},
    {'id': 2,  'name': 'Babhnan',    'unit_code': 'BBN', 'file': 'Babhnan 26-27.xlsx'},
    {'id': 3,  'name': 'Rauzagaon',  'unit_code': 'RCM', 'file': 'Rauzagaon 26-27.xlsx'},
    {'id': 4,  'name': 'Maizapur',   'unit_code': 'MZP', 'file': 'Maizapur 26-27.xlsx'},
    {'id': 5,  'name': 'Mankapur',   'unit_code': 'MCM', 'file': 'Mankapur 26-27.xlsx'},
    {'id': 6,  'name': 'Gularia',    'unit_code': 'GCM', 'file': 'Gularia 26-27.xlsx'},
    {'id': 7,  'name': 'Tulsipur',   'unit_code': 'TCM', 'file': 'Tulsipur 26-27.xlsx'},
    {'id': 8,  'name': 'Kumbhi',     'unit_code': 'KCM', 'file': 'Kumbhi 26-27.xlsx'},
    {'id': 9,  'name': 'Haidergarh', 'unit_code': 'HCM', 'file': 'Haidergarh 26-27.xlsx'},
    {'id': 10, 'name': 'Akbarpur',   'unit_code': 'ACM', 'file': 'Akbarpur 26-27.xlsx'},
]

def normalise_collar(val):
    v = str(val).strip().upper()
    if any(x in v for x in ['WHITE', 'WC', 'W C']):
        return 'White Collared'
    if any(x in v for x in ['BLUE', 'BC', 'B C']):
        return 'Blue Collared'
    return str(val).strip()

def clean_designation(val):
    if not val or str(val).strip() == 'nan':
        return ''
    # Babhnan has "DESIGNATION (DESIGNATION_PLANT_...)" — strip the parenthetical
    return re.sub(r'\s*\(.*?\)\s*$', '', str(val)).strip()

def clean_val(val):
    if val is None:
        return ''
    s = str(val).strip()
    return '' if s.lower() in ('nan', 'none', '') else s

def find_header_row(df):
    """Find the row index that has 'Employee Code' or 'Sr.' as first meaningful column."""
    for i, row in df.iterrows():
        vals = [str(v).strip().lower() for v in row.tolist() if str(v).strip() not in ('', 'nan')]
        if any(x in vals for x in ['employee code', 'sr.', 'sr']):
            return i
    return 0

def load_plant_file(filepath):
    xl   = pd.ExcelFile(filepath)
    sheet = xl.sheet_names[0]
    raw  = pd.read_excel(filepath, sheet_name=sheet, header=None, dtype=str)

    hdr_row = find_header_row(raw)

    # Read again with correct header — skip title rows above header
    df = pd.read_excel(filepath, sheet_name=sheet, header=None,
                       skiprows=hdr_row, dtype=str)

    # Flatten multi-row headers (some files have 2 header rows)
    row0 = [str(v).strip() for v in df.iloc[0].tolist()]
    row1 = [str(v).strip() for v in df.iloc[1].tolist()] if len(df) > 1 else [''] * len(row0)

    # Build column names: combine row0 + row1 where row1 isn't blank/nan
    cols = []
    for a, b in zip(row0, row1):
        if b and b.lower() not in ('nan', ''):
            cols.append((a + ' ' + b).strip())
        else:
            cols.append(a)
    df.columns = cols

    # Drop header rows, keep only data rows
    df = df.iloc[2:].reset_index(drop=True)

    # Keep only rows where first column (Sr.) is a number
    def is_num(v):
        try:
            float(str(v).replace(',', ''))
            return True
        except Exception:
            return False

    df = df[df.iloc[:, 0].apply(is_num)].reset_index(drop=True)
    return df

def map_columns(df):
    """Map varied column names to standard names."""
    col_map = {}
    for col in df.columns:
        c = col.strip().lower()
        if 'employee code' in c or 'emp code' in c:
            col_map['emp_code'] = col
        elif col.strip().lower() in ('name',) or ('name' in c and 'designation' not in c):
            if 'emp_code' in col_map and 'name' not in col_map:
                col_map['name'] = col
            elif 'name' not in col_map:
                col_map['name'] = col
        elif 'designation' in c:
            col_map['designation'] = col
        elif 'grade' in c:
            col_map['grade'] = col
        elif 'collar' in c or 'blue/white' in c:
            col_map['collar'] = col
        elif 'department' in c or 'dept' in c:
            col_map['department'] = col
        elif 'section' in c:
            col_map['section'] = col
        elif 'category' in c:
            col_map['category'] = col
        elif 'gender' in c:
            col_map['gender'] = col
        elif 'physically' in c or 'handicap' in c or col.strip() == 'PH':
            col_map['ph'] = col
        elif 'remark' in c:
            col_map['remarks'] = col
    return col_map

def seed_plant(db, plant, df, col_map):
    inserted = 0
    skipped  = 0
    for _, row in df.iterrows():
        emp_code = clean_val(row.get(col_map.get('emp_code', ''), '')).split('.')[0].strip()
        name     = clean_val(row.get(col_map.get('name', ''), ''))
        if not emp_code or not name:
            skipped += 1
            continue

        desig    = clean_designation(row.get(col_map.get('designation', ''), ''))
        grade    = clean_val(row.get(col_map.get('grade', ''), ''))
        collar   = normalise_collar(row.get(col_map.get('collar', ''), ''))
        dept     = clean_val(row.get(col_map.get('department', ''), ''))
        section  = clean_val(row.get(col_map.get('section', ''), ''))
        category = clean_val(row.get(col_map.get('category', ''), ''))
        gender_raw = clean_val(row.get(col_map.get('gender', ''), '')).capitalize()
        gender   = gender_raw if gender_raw in ('Male', 'Female', 'Others') else ''
        ph_raw   = clean_val(row.get(col_map.get('ph', ''), '')).upper()
        ph       = 'Yes' if ph_raw in ('Y', 'YES') else 'No'
        remarks  = clean_val(row.get(col_map.get('remarks', ''), ''))

        try:
            db.execute('''INSERT OR IGNORE INTO employees
                (plant_id,emp_code,name,designation,grade,collar,department,
                 section,category,gender,physically_handicapped,remarks)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?)''',
                (plant['id'], emp_code, name, desig, grade, collar,
                 dept, section, category, gender, ph, remarks))
            inserted += 1
        except Exception as e:
            skipped += 1
    return inserted, skipped

def main():
    if not os.path.exists(DB_PATH):
        print("ERROR: Database not found. Run 'python app.py' once first to initialise the DB.")
        return

    db = sqlite3.connect(DB_PATH)
    db.execute("PRAGMA foreign_keys = ON")

    total_inserted = 0
    total_skipped  = 0

    for plant in PLANTS:
        filepath = os.path.join(DATA_DIR, plant['file'])
        if not os.path.exists(filepath):
            print(f"  SKIP {plant['name']}: file not found at {filepath}")
            continue

        print(f"\nLoading {plant['name']} ({plant['unit_code']})...")
        try:
            df      = load_plant_file(filepath)
            col_map = map_columns(df)
            ins, skp = seed_plant(db, plant, df, col_map)
            total_inserted += ins
            total_skipped  += skp
            print(f"  Inserted: {ins} | Skipped: {skp}")
        except Exception as e:
            print(f"  ERROR: {e}")

    db.commit()
    db.close()
    print(f"\nDone. Total inserted: {total_inserted} | Total skipped: {total_skipped}")
    print("All employee data loaded. You can now log in as any SPOC.")

if __name__ == '__main__':
    main()
