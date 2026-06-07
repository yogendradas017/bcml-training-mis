import re, sys, json, sqlite3, os, urllib.parse
import requests

BASE = "http://localhost:5000"
S = requests.Session()

def get_csrf(html):
    m = re.search(r'name="csrf_token"\s+value="([^"]+)"', html)
    if not m:
        m = re.search(r'name="csrf-token"\s+content="([^"]+)"', html)
    return m.group(1) if m else None

# --- Login as balrampur SPOC ---
r = S.get(f"{BASE}/login")
tok = get_csrf(r.text)
r = S.post(f"{BASE}/login", data={"csrf_token": tok, "username": "balrampur", "password": "TestPwd@2026Strong!"}, allow_redirects=False)
print("LOGIN status:", r.status_code, "->", r.headers.get('Location'))

# follow
r2 = S.get(f"{BASE}/spoc/dashboard", allow_redirects=False)
print("spoc dash:", r2.status_code)

# --- Test 1: GET /tni — verify panel ---
r = S.get(f"{BASE}/tni")
print("\n=== TEST 1: GET /tni ===")
print("status:", r.status_code, "len:", len(r.text))
has_panel = "Quick-Add TNI for New Employee" in r.text
has_collapse = 'id="quickAddPanel"' in r.text
collapse_default = 'class="collapse" id="quickAddPanel"' in r.text
expanded_attr = 'aria-expanded="false"' in r.text
print("Has panel header:", has_panel)
print("Has panel id:", has_collapse)
print("Collapsed by default:", collapse_default)
print("aria-expanded false:", expanded_attr)
tni_html = r.text
tni_csrf = get_csrf(tni_html)
print("csrf:", tni_csrf[:20] if tni_csrf else None)

# Get plant_id and 3 programme_master ids + 1 valid emp_code from DB
DB = os.path.join(os.path.dirname(__file__), "..", "data", "training.db")
con = sqlite3.connect(DB); con.row_factory = sqlite3.Row
plant_id = con.execute("SELECT id FROM plants WHERE LOWER(name)='balrampur'").fetchone()
if not plant_id:
    plant_id = con.execute("SELECT plant_id as id FROM users WHERE username='balrampur'").fetchone()
pid = plant_id['id']
print("plant_id:", pid)
prog_ids = [r['id'] for r in con.execute("SELECT id FROM programme_master WHERE plant_id=? ORDER BY id LIMIT 3", (pid,)).fetchall()]
emp = con.execute("SELECT emp_code FROM employees WHERE plant_id=? AND is_active=1 LIMIT 1", (pid,)).fetchone()
emp_code = emp['emp_code'] if emp else None
print("prog_ids:", prog_ids, "emp_code:", emp_code)

# Snapshot tni rows for this emp BEFORE
before = con.execute("SELECT COUNT(*) FROM tni WHERE plant_id=? AND emp_code=?", (pid, emp_code)).fetchone()[0]
print("tni rows before:", before)
con.close()

def get_flashes(html):
    # quick capture of alert messages
    return re.findall(r'<div[^>]*class="[^"]*alert[^"]*"[^>]*>(.*?)</div>', html, re.S)

# --- Test 2: POST quick-add with valid emp + 3 progs ---
print("\n=== TEST 2: POST quick-add valid ===")
data = [("csrf_token", tni_csrf), ("emp_code", emp_code), ("default_hours", "4")]
for pid_ in prog_ids:
    data.append(("prog_ids", str(pid_)))
r = S.post(f"{BASE}/tni/quick-add-for-employee", data=data, allow_redirects=True)
print("status:", r.status_code)
# Check inserted rows
con = sqlite3.connect(DB); con.row_factory = sqlite3.Row
after = con.execute("SELECT COUNT(*) FROM tni WHERE plant_id=? AND emp_code=?", (pid, emp_code)).fetchone()[0]
print("tni rows after:", after, "delta:", after - before)
# Check flashes on redirected page
fl = get_flashes(r.text)
print("flashes:", [re.sub(r'\s+', ' ', f).strip()[:200] for f in fl])
con.close()

# --- Test 3: POST same again - should report 3 duplicates skipped ---
print("\n=== TEST 3: POST same again (dupes) ===")
r = S.get(f"{BASE}/tni"); tni_csrf = get_csrf(r.text)
data = [("csrf_token", tni_csrf), ("emp_code", emp_code), ("default_hours", "4")]
for pid_ in prog_ids:
    data.append(("prog_ids", str(pid_)))
r = S.post(f"{BASE}/tni/quick-add-for-employee", data=data, allow_redirects=True)
print("status:", r.status_code)
con = sqlite3.connect(DB); con.row_factory = sqlite3.Row
after2 = con.execute("SELECT COUNT(*) FROM tni WHERE plant_id=? AND emp_code=?", (pid, emp_code)).fetchone()[0]
print("tni rows now:", after2, "delta from previous:", after2 - after)
fl = get_flashes(r.text)
print("flashes:", [re.sub(r'\s+', ' ', f).strip()[:300] for f in fl])
con.close()

# --- Test 4: POST with no prog_ids ---
print("\n=== TEST 4: POST no prog_ids ===")
r = S.get(f"{BASE}/tni"); tni_csrf = get_csrf(r.text)
data = [("csrf_token", tni_csrf), ("emp_code", emp_code), ("default_hours", "4")]
r = S.post(f"{BASE}/tni/quick-add-for-employee", data=data, allow_redirects=True)
print("status:", r.status_code)
fl = get_flashes(r.text)
print("flashes:", [re.sub(r'\s+', ' ', f).strip()[:300] for f in fl])

# --- Test 5: POST invalid emp_code ---
print("\n=== TEST 5: POST emp_code=ZZZ ===")
r = S.get(f"{BASE}/tni"); tni_csrf = get_csrf(r.text)
data = [("csrf_token", tni_csrf), ("emp_code", "ZZZ"), ("default_hours", "4"),
        ("prog_ids", str(prog_ids[0]))]
r = S.post(f"{BASE}/tni/quick-add-for-employee", data=data, allow_redirects=True)
print("status:", r.status_code)
fl = get_flashes(r.text)
print("flashes:", [re.sub(r'\s+', ' ', f).strip()[:300] for f in fl])

# --- Test 6: GET /programme-master — verify NO Sync button ---
print("\n=== TEST 6: GET /programme-master ===")
r = S.get(f"{BASE}/programme-master")
print("status:", r.status_code)
has_sync_btn = "Sync from TNI" in r.text
has_sync_url = "sync-from-tni" in r.text
print("Has 'Sync from TNI' text:", has_sync_btn)
print("Has sync-from-tni URL:", has_sync_url)

# --- Test 7: POST /programme-master/sync-from-tni as SPOC ---
print("\n=== TEST 7: POST sync-from-tni as SPOC ===")
pm_csrf = get_csrf(r.text)
r = S.post(f"{BASE}/programme-master/sync-from-tni",
           data={"csrf_token": pm_csrf}, allow_redirects=True)
print("status:", r.status_code)
fl = get_flashes(r.text)
print("flashes:", [re.sub(r'\s+', ' ', f).strip()[:400] for f in fl])

# Also verify nothing was wiped — check pm row count
con = sqlite3.connect(DB); con.row_factory = sqlite3.Row
pm_count = con.execute("SELECT COUNT(*) FROM programme_master WHERE plant_id=?", (pid,)).fetchone()[0]
print("programme_master row count for plant:", pm_count)
con.close()

# Cleanup — delete the tni rows we inserted to keep DB clean
con = sqlite3.connect(DB)
con.execute("DELETE FROM tni WHERE plant_id=? AND emp_code=? AND id > (SELECT COALESCE(MAX(id),0) FROM tni WHERE plant_id=? AND emp_code=?) - ?",
            (pid, emp_code, pid, emp_code, after2 - before))
# safer: just rely on prog names from snapshot
con.commit(); con.close()
print("\nDONE")
