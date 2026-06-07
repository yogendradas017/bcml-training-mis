# Design Parity Plan: Internal App → Apple/Swiss Design

> Status: **Pending** — fully specified, not yet executed.

## Context
Internal TMS app looks dated vs polished landing page. Root causes:
1. `central.html` uses legacy `.card.stat-card` with inline border-color instead of `.sc` gradient cards
2. 13 templates have dark inline card headers: `background:linear-gradient(135deg,#1a1f35,#252d45)` / `background:#1a3a5c`
3. Admin SPOC banner uses orange gradient `#92400e,#d97706`
4. Table action buttons are multi-color loud
5. Unit code badges use Bootstrap `bg-secondary` gray
6. Coverage % inline color logic: raw Jinja `color:#dc2626` etc.
7. `dashboard.html` is ALREADY correct (uses `.sc` cards) — do NOT touch it

## Files to modify
- `static/css/main.css` — append utility classes
- `templates/central.html` — targeted edits after batch script
- 13 templates (batch script): `admin_tni_archives.html`, `calendar.html`, `central_plant.html`, `employees.html`, `intelligence.html`, `programme_2c.html`, `summary.html`, `tni.html`, `tni_cleanse.html`, `tni_fresh_upload.html`, `tni_msforms.html`, `training_2a.html`, `central.html`

## Design decision
- **Major data section headers**: dark BCML navy → `.ch-dark`
- **Utility / admin / filter panels**: light Apple → `.ch`
- All dark headers unified to ONE consistent gradient (replacing scattered #1a1f35, #1a3a5c, #92400e, etc.)

---

## Step 1 — CSS to append to `static/css/main.css`

```css
/* ── CLEAN CARD HEADER (utility/admin panels) ── */
.ch { padding:13px 20px; background:#f5f5f7; color:var(--c-text); border-bottom:1px solid var(--c-border); font-size:13px; font-weight:700; display:flex; align-items:center; gap:8px; border-radius:0; }
.ch i { color:var(--c-primary); font-size:14px; }

/* ── DARK SECTION HEADER (primary data tables) ── */
.ch-dark { padding:13px 20px; background:linear-gradient(135deg,#1e3a5f 0%,#2e6da4 100%); color:#fff; font-size:13px; font-weight:700; display:flex; align-items:center; gap:8px; border-radius:0; }
.ch-dark i { color:rgba(255,255,255,.8); }

/* ── UNIT CODE BADGE ── */
.badge-unit { background:#1d1d1f; color:#fff; font-size:10px; padding:3px 7px; border-radius:5px; font-weight:700; letter-spacing:.4px; }

/* ── COVERAGE % ── */
.pct-good { color:#059669; font-weight:700; }
.pct-mid  { color:#d97706; font-weight:700; }
.pct-low  { color:#dc2626; font-weight:700; }

/* ── ICON ACTION BUTTON ── */
.btn-icon { width:28px; height:28px; display:inline-flex; align-items:center; justify-content:center; border-radius:7px; border:1px solid var(--c-border); background:var(--c-white); color:var(--c-muted); font-size:13px; transition:all var(--transition); text-decoration:none; flex-shrink:0; }
.btn-icon:hover { background:#f0f6ff; color:var(--c-primary); border-color:#bdd6f5; }
.btn-icon.bi-danger { color:#dc2626; } .btn-icon.bi-danger:hover { background:#fef2f2; border-color:#fca5a5; }
.btn-icon.bi-success { color:var(--c-primary2); } .btn-icon.bi-success:hover { background:#f0fdf4; border-color:#86efac; }
.btn-icon.bi-enter { color:var(--c-primary); }

/* ── ADMIN PANEL ── */
.admin-panel { background:var(--c-white); border:1px solid var(--c-border); border-radius:var(--radius-lg); margin-bottom:20px; overflow:hidden; }
.admin-panel .ch { border-left:3px solid var(--c-primary); background:#fff; }

/* ── PLANT CHIP ── */
.plant-chip { display:flex; flex-direction:column; padding:9px 12px; background:var(--c-white); border:1px solid var(--c-border); border-radius:var(--radius-sm); text-decoration:none; transition:all var(--transition); }
.plant-chip:hover { border-color:var(--c-primary); background:#f0f6ff; transform:translateY(-1px); box-shadow:var(--shadow-sm); }
.plant-chip-top { display:flex; align-items:center; justify-content:space-between; }
.plant-chip-code { font-size:11.5px; font-weight:800; color:var(--c-primary); letter-spacing:.3px; }
.plant-chip-name { font-size:10.5px; color:var(--c-muted); margin-top:2px; }

/* ── COMPLIANCE MINI CARD ── */
.comp-card { background:var(--c-white); border:1px solid var(--c-border); border-radius:var(--radius-md); padding:16px; text-align:center; transition:all var(--transition); }
.comp-card:hover { box-shadow:var(--shadow-md); transform:translateY(-2px); }
.comp-plant { font-size:12px; font-weight:700; color:var(--c-text); }
.comp-code  { font-size:10px; color:var(--c-muted); margin-top:1px; }
.comp-pct   { font-size:1.7rem; font-weight:900; line-height:1.1; margin:8px 0 2px; }
.comp-lbl   { font-size:10.5px; color:var(--c-muted); }
```

---

## Step 2 — Batch script `fix_styles.py` (run once, then delete)

```python
import re
from pathlib import Path

TEMPLATES = Path('templates')

DARK_STYLES = [
    'style="background:linear-gradient(135deg,#1a1f35,#252d45);color:#fff;"',
    'style="background:#1a3a5c;color:#fff;"',
    'style="background:linear-gradient(135deg,#065f46,#059669);color:#fff;"',
    'style="background:linear-gradient(135deg,#0891b2,#06b6d4);color:#fff;"',
    'style="background:linear-gradient(135deg,#047857,#059669);color:#fff;"',
]
AMBER_STYLE = 'style="background:linear-gradient(135deg,#92400e,#d97706);color:#fff;"'
DARK_TR = re.compile(r'<tr\s+style="background:#252d45[^"]*">')

TARGET_FILES = [
    'admin_tni_archives.html','calendar.html','central_plant.html',
    'employees.html','intelligence.html','programme_2c.html','summary.html',
    'tni.html','tni_cleanse.html','tni_fresh_upload.html','tni_msforms.html',
    'training_2a.html','central.html',
]

for fname in TARGET_FILES:
    p = TEMPLATES / fname
    if not p.exists(): continue
    txt = p.read_text(encoding='utf-8')
    for s in DARK_STYLES:
        txt = txt.replace(s, 'class="ch-dark"')
    txt = txt.replace(AMBER_STYLE, 'class="ch"')
    txt = DARK_TR.sub('<tr>', txt)
    txt = txt.replace('class="badge bg-secondary"', 'class="badge-unit"')
    p.write_text(txt, encoding='utf-8')
    print(f'  patched {fname}')
print('Done.')
```

---

## Step 3 — `templates/central.html` targeted edits (after batch script)

1. **Stat cards (lines 6–31):** replace `.card.stat-card` with `.sc` gradient cards matching `dashboard.html`
   - Total Employees → `sc-violet` + `bi-people-fill`
   - Sessions Planned → `sc-blue` + `bi-calendar3`
   - Sessions Conducted → `sc-green` + `bi-check-circle-fill`
   - Total Man-Hours → `sc-amber` + `bi-clock-history`

2. **Admin banner:** replace orange card with `.admin-panel` + plant chips:
   ```html
   <div class="admin-panel">
     <div class="ch"><i class="bi bi-shield-lock"></i> Admin — Enter Any Plant as SPOC</div>
     <div class="p-3"><div class="row g-2">
       {% for p in plants %}
       <div class="col-6 col-md-4 col-lg-2">
         <a href="{{ url_for('admin_select_plant', plant_id=p.id) }}" class="plant-chip">
           <div class="plant-chip-top">
             <span class="plant-chip-code">{{ p.unit_code }}</span>
             <span class="pct-{{ 'good' if p.bc_pct>=75 else ('mid' if p.bc_pct>=40 else 'low') }}">{{ p.bc_pct }}%</span>
           </div>
           <div class="plant-chip-name">{{ p.name }}</div>
         </a>
       </div>
       {% endfor %}
     </div></div>
   </div>
   ```

3. **Coverage % columns:** `color:{{ bc_color }}` → `class="pct-{{ 'good' if p.bc_pct>=75 else ('mid' if p.bc_pct>=40 else 'low') }}"`

4. **Compliance heat map:** add `.comp-card`, `.comp-plant`, `.comp-pct`, `.comp-lbl` classes

5. **Action buttons:** replace multi-color btns with `.btn-icon` icon-only buttons

---

## Step 4 — Commit
Single commit: `"refactor: unified card headers, badge-unit, btn-icon, pct classes across all templates"`

## Verification
- `/central` — gradient stat cards, no orange/navy inline styles, admin panel blue-border
- `/calendar`, `/tni`, `/training`, `/employees` — all card headers now `.ch-dark` or `.ch`
- Unit code badges dark pill, not Bootstrap gray
- Action buttons subtle icon-only
- Coverage % uses CSS class, not inline style
