import os

BASE_DIR       = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_default_db    = os.path.join(BASE_DIR, 'data', 'training.db')
_env_db        = os.environ.get('DATABASE_PATH')
if _env_db:
    try:
        os.makedirs(os.path.dirname(_env_db), exist_ok=True)
        DB_PATH = _env_db
    except Exception:
        DB_PATH = _default_db
else:
    DB_PATH = _default_db

TEMP_UPLOAD_DIR = os.path.join(BASE_DIR, 'data', 'tmp_uploads')

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
PLANT_MAP = {p['id']: p for p in PLANTS}
PLANT_MAP[99] = {'id': 99, 'name': 'Central', 'unit_code': 'CEN'}

PROG_TYPES  = ['Behavioural/Leadership', 'Cane', 'Commercial', 'EHS/HR', 'IT', 'Technical']
MODES       = ['Classroom', 'OJT', 'SOP', 'Online']
LEVELS      = ['General', 'Specialized']
AUDIENCES   = ['Blue Collared', 'White Collared', 'Common']
STATUSES    = ['To Be Planned', 'Conducted', 'Re-Scheduled', 'Cancelled']
INT_EXT     = ['Internal', 'External', 'Online']
MONTHS_FY   = ['April','May','June','July','August','September',
               'October','November','December','January','February','March']
MONTH_NUM   = {m: f'{i+1:02d}' for i, m in enumerate(
               ['January','February','March','April','May','June',
                'July','August','September','October','November','December'])}
CAL_NEW     = ['Calendar Program', 'New Program']
GENDERS     = ['Male', 'Female', 'Others']
COLLARS     = ['Blue Collared', 'White Collared']
PH_OPTIONS  = ['No', 'Yes']

GRADES = [
    'MANAGERIAL 1A', 'MANAGERIAL 1B', 'MANAGERIAL 1C',
    'MANAGERIAL 2A', 'MANAGERIAL 2B', 'MANAGERIAL 2C', 'MANAGERIAL 2C-I',
    'MANAGERIAL 3A', 'MANAGERIAL 3B', 'MANAGERIAL 3C', 'MANAGERIAL 3D',
    'MANAGERIAL 4A', 'MANAGERIAL 4B', 'MANAGERIAL 4C',
    'SUPERVISORY A', 'SUPERVISORY B', 'SUPERVISORY C',
    'HIGHLY SKILLED',
    'SKILLED A', 'SKILLED B', 'SKILLED C',
    'SEMI-SKILLED', 'SEMI-SKILLED A', 'SEMI-SKILLED B',
    'UN SKILLED',
    'CLERICAL I', 'CLERICAL II', 'CLERICAL III', 'CLERICAL IV',
    'AGREEMENTAL I-A', 'AGREEMENTAL I-B',
    'AGREEMENTAL II-A', 'AGREEMENTAL II-B',
    'AGREEMENTAL III-A', 'AGREEMENTAL III-B', 'AGREEMENTAL III-C', 'AGREEMENTAL III-D',
    'AGREEMENTAL IV-A', 'AGREEMENTAL IV-B', 'AGREEMENTAL IV-C', 'AGREEMENTAL IV-D',
    'AGREEMENTAL SEASONAL III-A', 'AGREEMENTAL SEASONAL III-B',
    'AGREEMENTAL SEASONAL III-C', 'AGREEMENTAL SEASONAL III-D',
    'AGREEMENTAL SEASONAL IV-A', 'AGREEMENTAL SEASONAL IV-B',
    'AGREEMENTAL SEASONAL IV-C', 'AGREEMENTAL SEASONAL IV-D',
    'TRAINEE - MGT&GET', 'TRAINEE - OTHERS',
    'APPRENTICE', 'TEMPORARY', 'SEASONAL CONSOLIDATE',
]

CATEGORIES = [
    'PERMANENT', 'SEASONAL', 'AGREEMENTAL', 'AGREEMENTAL SEASONAL',
    'TEMPORARY', 'TRAINEE - MGT&GET', 'TRAINEE - OTHERS', 'APPRENTICE',
]
TYPE_ABBREV = {
    'Behavioural/Leadership': 'BEH', 'Cane': 'CAN', 'Commercial': 'COM',
    'EHS/HR': 'EHS', 'IT': 'ITC', 'Technical': 'TEC'
}
NON_TNI_SOURCES = ('New Requirement',)

CENTRAL_PLANT_ID = 99
