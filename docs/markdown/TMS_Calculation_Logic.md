# TMS — Calculation Logic Reference

**Purpose.** This document is the authoritative, code-verified specification of every number the Training Management System (TMS) computes and displays. It is written for an external reviewing body that needs to trace each metric to its exact source formula, inputs, filters, and edge-case behaviour. Every entry cites the precise `file:line · function` from which the formula is taken. Formulas are reproduced **as implemented in the code**, not as an idealised intent — where the code diverges from naïve expectation (cross-year buckets, exit-policy asymmetry, double-rounding, mapping quirks), that divergence is documented explicitly.

**Last updated:** 2026-06-07. Reflects the current production logic including the most recent fixes (Monthly Summary conducted-only programme count and FY-bound seats/man-hours; the compliance exit policy; host-aware session-actuals recompute; and the COALESCE-only feedback fold into 2C).

**Stack.** Flask · raw `sqlite3` (no ORM) · helpers in `tms/helpers.py` · routes in `tms/routes/*.py` · config in `tms/config.py` · schema in `schema.sql`. Plant `99` is the **Central** pseudo-plant.

---

## Global Conventions

These hold across the whole document unless an entry states otherwise.

- **Financial year (FY): April → March.**
  - `_fy_label()` returns the short label `'YY-YY'` (e.g. `'26-27'`). For months Jan–Mar it returns `f'{(y-1)[2:]}-{y[2:]}'`; for Apr–Dec, `f'{y[2:]}-{(y+1)[2:]}'`. Source: `tms/helpers.py:260`.
  - `_current_fy()` returns the inclusive date bounds `('YYYY-04-01', 'YYYY+1-03-31')` where `yr = today.year if today.month >= 4 else today.year-1`. Source: `tms/helpers.py:68`.
  - `_in_current_fy(date_str)` treats empty/None as in-FY (returns `True`); bad/unparseable dates return `False`; otherwise checks `fy_start <= date_str[:10] <= fy_end`. Source: `tms/helpers.py:76`.
  - `_tni_is_locked()` returns `True` once `_today_ist() > fy_end` (TNI write window closes after Mar 31). Source: `tms/helpers.py:88`.

- **IST clock (server runs UTC).** All user-facing dates/times and every FY/lock/gate computation use IST via `_now_ist()` / `_today_ist()`. `_now_ist()` = `datetime.now(ZoneInfo('Asia/Kolkata')).replace(tzinfo=None)` (timezone-naive IST wall clock); fallback adds a fixed `+5:30` offset (India has no DST). Source: `tms/helpers.py:129, 141`. The single exception is the session-code last-resort uniqueness token, which uses `datetime.utcnow().strftime('%f')` (not a user-facing time).

- **Enums (exact strings).**
  - `source ∈ {TNI Driven, New Requirement}` — any other value is silently coerced to `TNI Driven` at calendar write time.
  - `collar ∈ {Blue Collared, White Collared}` — employees with any other/blank collar are excluded from collar-split aggregates.
  - `audience ∈ {Blue Collared, White Collared, Common}` — derived from TNI where TNI rows exist (see Audience Derivation); free-text only for New Requirement programmes with no TNI.
  - `prog_type` order list `PROG_TYPES = ['Behavioural/Leadership','Cane','Commercial','EHS/HR','IT','Technical']`.
  - `int_ext ∈ {Internal, External, Online}` — only `Internal`/`External` are matched in the programme split; `Online` is counted in neither.

- **Per-plant man-hour mandate targets.** Targets are **per-plant config**, never global constants: `get_config('mh_target_bc', 12, plant_id)` (default 12) and `get_config('mh_target_wc', 24, plant_id)` (default 24). Resolution order: plant-scoped override → global → hardcoded default; validated `1..8760` on write. Source: `tms/config.py` `get_config`.

- **Plant scoping.** SPOC views are scoped to `session['plant_id']`. Central (`/central`) aggregates the 10 real plants in `tms/constants.py PLANTS` (ids 1–10); plant 99 is **not** in `PLANTS` and is not summed into fleet grand totals, but it **is** in `PLANT_MAP`, so `/central/plant/99` renders Central as a "plant".

- **What "Conducted" means.** A programme/session counts as conducted when its calendar `status='Conducted'`. The Monthly Summary additionally counts legacy 2C rows that have **no** calendar row at all (`programme_details LEFT JOIN calendar`, gate `c.session_code IS NULL OR c.status='Conducted'`). Rows sitting in `'Awaiting Verification'` are **excluded** until Central verifies — matching the Dashboard and Export gates.

- **Exit policy (owner decision, documented explicitly).** For the man-hour compliance gauges: exited employees' **delivered man-hours STAY in the numerator** (the man-hours numerator join to `employees` has **no** `is_active` filter), but exited employees are **EXCLUDED from the headcount denominator** (`is_active=1`). This asymmetry means the gauge **CAN exceed 100% by design**. Note: TNI **coverage %** (Monthly Summary) does *not* apply this — it has no `is_active` filter on either side. The worst-cells gap analysis goes the other way — it filters `is_active=1` and so removes exited employees' nominations from both numerator and denominator.

- **Real-time vs cached.** Most figures are computed live per request. The QC analytics endpoint (`/api/dashboard-qc`) caches its six charts in-process for **60 seconds** per `(plant_id, fy_start)` — near-real-time, up to 60s stale within a gunicorn worker, never event-invalidated (TTL only).

---

## Monthly Summary — Programmes, Person Seats, Man-Hours

The Monthly Summary table (`/summary`, SPOC; and the Central per-plant card) is produced by `_calc_summary(plant_id, month_filter, db)` (`tms/helpers.py:424`) with a TOTAL footer from `_calc_totals(rows, db, plant_id)` (`tms/helpers.py:578`). It runs four GROUP BY queries: Query 1 = conducted programme counts; Query 2 = central-hosted programme tally; Query 3 = seats + man-hours; Query 4 = TNI coverage (documented in the Coverage section).

### No. of Programmes — BC / WC / Common / Total

- **What it is.** Count of DISTINCT conducted programmes for the plant, split by target audience, plus central-hosted programmes the plant's employees attended elsewhere (classified by attendee collar mix).
- **Formula.**
  ```sql
  -- Query 1 (base), GROUP BY p.prog_type
  bc_progs     = COUNT(DISTINCT CASE WHEN p.audience='Blue Collared'  THEN p.programme_name END)
  wc_progs     = COUNT(DISTINCT CASE WHEN p.audience='White Collared' THEN p.programme_name END)
  common_progs = COUNT(DISTINCT CASE WHEN p.audience='Common'         THEN p.programme_name END)
  total_progs  = COUNT(DISTINCT p.programme_name)
  FROM programme_details p LEFT JOIN calendar c
       ON c.session_code=p.session_code AND c.plant_id=p.plant_id
  WHERE p.plant_id=? {month}  AND (c.session_code IS NULL OR c.status='Conducted')
  ```
  Then add the central tally (Query 2): per `(prog_type, programme_name)` not already in `programme_details`, classify by attendee collar mix — `bc_cnt>0 AND wc_cnt>0` → +1 common & +1 total; bc only → +1 bc & +1 total; wc only → +1 wc & +1 total. Assembly adds `ct['bc'/'wc'/'common'/'total']` onto the base counts.
- **Inputs.** `programme_details` (`prog_type`, `audience`, `programme_name`, `session_code`, `plant_id`, `start_date`) LEFT JOIN `calendar` (`session_code`, `plant_id`, `status`); central add from `emp_training` LEFT JOIN `employees`.
- **Filters.** Query 1: `p.plant_id=plant_id`; conducted-only gate; optional month on `strftime('%m', p.start_date)`. **No FY-window filter on the programme count** (only seats/man-hours are FY-bound). Central add (Query 2): `et.plant_id=plant_id AND et.host_plant_id=99`, FY-bound `et.start_date BETWEEN fy_start AND fy_end`, optional month, and `NOT EXISTS` a matching `programme_details` row (case-insensitive name + prog_type, month-aligned when a month is selected).
- **Edge cases.** `'Awaiting Verification'` 2C rows excluded. `COUNT(DISTINCT programme_name)` dedups multiple sessions of one programme. Because Query 1 is not FY-bound while seats/man-hours are, a row's programme count and its seats/hrs can describe different time spans if conducted rows exist outside the FY. Central programmes excluded if an equivalent 2C row already exists (prevents double count). Audience values outside the three exact strings raise `total_progs` but no sub-column. TOTAL row = plain Python sum of each per-row column (round 1 dp).
- **Source.** `tms/helpers.py:447-462` (Query 1), `:466-497` (Query 2 tally), `:534-547` (assembly), `:578-587` (`_calc_totals`) · `_calc_summary` (`tms/helpers.py:424`).

### No. of Programmes — Internal / External

- **What it is.** Of the conducted programmes, how many were delivered by an internal vs external trainer. Central-hosted programmes are **not** split here.
- **Formula.**
  ```sql
  int_prog = COUNT(DISTINCT CASE WHEN p.int_ext='Internal' THEN p.programme_name END)
  ext_prog = COUNT(DISTINCT CASE WHEN p.int_ext='External' THEN p.programme_name END)
  GROUP BY p.prog_type
  ```
  TOTAL row = Python sum across prog_types.
- **Inputs.** `programme_details.int_ext`, `.programme_name`, `.prog_type` (Query 1 only).
- **Filters.** Same as programme count Query 1 (`p.plant_id=plant_id`; conducted-only; optional month). No FY window.
- **Edge cases.** `int_ext='Online'` is counted in **neither** column, so `int_prog + ext_prog` can be `< total_progs`. Central-hosted programmes never contribute here. DISTINCT-per-name dedup applies.
- **Source.** `tms/helpers.py:453-454` (CASE), `:540-541` (assembly) · `_calc_summary` (`tms/helpers.py:424`).

### Person Seats — BC / WC / Total

- **What it is.** Seats filled — one count per person per session attended, split by collar. Three sessions attended = three seats. FY-scoped.
- **Formula.**
  ```sql
  -- Query 3, GROUP BY t.prog_type, e.collar
  seats = COUNT(*) FROM emp_training t JOIN employees e
  -- assembly
  bc_seats    = seat_map[pt]['Blue Collared'][0]   (default 0)
  wc_seats    = seat_map[pt]['White Collared'][0]  (default 0)
  total_seats = bc_seats + wc_seats
  ```
  TOTAL row = Python sum across prog_types (round 1 dp).
- **Inputs.** `emp_training` (`prog_type`, `emp_code`, `plant_id`, `start_date`, `hrs`) JOIN `employees e` ON `e.emp_code=t.emp_code AND e.plant_id=t.plant_id` (`e.collar`).
- **Filters.** `t.plant_id=plant_id`; FY-bound `t.start_date BETWEEN fy_start AND fy_end`; `e.collar IN ('Blue Collared','White Collared')`; optional month on `strftime('%m', t.start_date)`; month-missing-from-`MONTH_NUM` → `AND 1=0`.
- **Edge cases.** INNER JOIN to `employees` **drops** attendance rows with no matching employee. Employees with collar outside the two literals are excluded entirely. Seats are **not** deduplicated by person — repeat attendance = multiple seats. New Requirement and non-TNI attendees **are** counted. Query 3 has **no host filter**, so central-hosted (`host_plant_id=99`) attendance is included alongside local. Each seat lands under the **employee's** collar, which may differ from the programme audience (audience-mismatch distortion).
- **Source.** `tms/helpers.py:499-511` (Query 3), `:549-551 & :566-567` (assembly), `:584-587` (`_calc_totals`) · `_calc_summary` (`tms/helpers.py:424`).

### Man-Hours — BC / WC / Total

- **What it is.** Total training hours delivered, summed across all attendance rows, split by collar. FY-scoped, rounded to 1 dp.
- **Formula.**
  ```sql
  -- Query 3 (same as seats), GROUP BY t.prog_type, e.collar
  hrs = COALESCE(SUM(t.hrs),0) FROM emp_training t JOIN employees e
  -- assembly
  bc_hrs    = round(seat_map[pt]['Blue Collared'][1], 1)
  wc_hrs    = round(seat_map[pt]['White Collared'][1], 1)
  total_hrs = round(bc_hrs + wc_hrs, 1)
  ```
  TOTAL row = Python sum of `bc_hrs`/`wc_hrs`/`total_hrs` across rows (round 1 dp).
- **Inputs.** `emp_training.hrs`, `.prog_type`, `.emp_code`/`.plant_id`/`.start_date` JOIN `employees.collar`.
- **Filters.** Identical to Person Seats Query 3.
- **Edge cases.** `COALESCE(SUM,0)` yields 0 for a prog_type/collar with no rows or NULL hrs. INNER JOIN drops orphan attendance. Per-row rounding to 1 dp then summed for TOTAL (small rounding drift possible vs raw sum). Hours land under the **employee's** collar (mismatch distortion). Central-hosted attendance hours are **included** (no host filter). FY bound is on `start_date` only — multi-day sessions are bucketed by their start date's FY/month.
- **Source.** `tms/helpers.py:499-511` (Query 3 `SUM(t.hrs)`), `:550-551 & :568-569` (assembly round), `:584-587` (`_calc_totals`) · `_calc_summary` (`tms/helpers.py:424`).

### Central-hosted programme inclusion

- **What it is.** When a plant's employees attend a Central-hosted (plant 99) programme with no local 2C record, it is credited to the plant's programme count; the collar split is decided by who from the plant attended.
- **Formula.**
  ```sql
  -- Query 2
  SELECT LOWER(et.prog_type) AS pt_lc, et.programme_name,
         SUM(CASE WHEN e.collar='Blue Collared'  THEN 1 ELSE 0 END) bc_cnt,
         SUM(CASE WHEN e.collar='White Collared' THEN 1 ELSE 0 END) wc_cnt
  ... GROUP BY pt_lc, et.programme_name
  -- Python tally per pt_lc
  bc_cnt>0 AND wc_cnt>0 -> common+=1, total+=1
  elif bc_cnt           -> bc+=1,     total+=1
  elif wc_cnt           -> wc+=1,     total+=1
  ```
  Added to the count columns at assembly, keyed by `pt.lower()`.
- **Inputs.** `emp_training et` (`prog_type`, `programme_name`, `emp_code`, `plant_id`, `host_plant_id`, `start_date`) LEFT JOIN `employees e` (`collar`); `NOT EXISTS` against `programme_details pd`.
- **Filters.** `et.plant_id=plant_id AND et.host_plant_id=99`; FY-bound; optional month; `NOT EXISTS` a `programme_details` row with same plant, `LOWER(programme_name)`, `LOWER(prog_type)`, and (if a month is selected) same month.
- **Edge cases.** Only `host_plant_id=99` (Central) is folded into the **programme count** here — sessions hosted by other plants are not added via this path (their **seats/hours** are still counted in Query 3, which has no host filter). LEFT JOIN to employees: an attendee with NULL/unmatched collar contributes to neither `bc_cnt` nor `wc_cnt`, so a programme attended only by such employees adds nothing. A programme is counted at most once per `(prog_type, programme_name)`. Dedup against 2C is case-insensitive on name+type, month-gated only when a month filter is active.
- **Source.** `tms/helpers.py:466-497` (Query 2 + tally), `:543-547` (assembly) · `_calc_summary` (`tms/helpers.py:424`).

### Month filter behaviour

- **What it is.** Selecting a calendar month narrows every figure to that month; "All" shows the whole FY; an unrecognised value blanks the page.
- **Formula.** `mn = MONTH_NUM.get(month_filter,'') if month_filter else ''`. Each clause: `"AND strftime('%m', <col>)='mm'"` if `mn`, else `"AND 1=0"` if a (bad) month was given, else `""`. Applied to `p.start_date` (Query 1), `t.start_date` (Query 3), `et.start_date` (Query 2), and `pd.start_date` (NOT EXISTS dedup).
- **Inputs.** `request.args.get('month','')` (`tms/routes/summary.py:15`); `MONTH_NUM` (Jan=01..Dec=12).
- **Filters.** Month matches the **calendar month number** (`strftime '%m'`), independent of FY ordering. Query 1 has no FY bound, so a month filter there matches that calendar month **across all years**, whereas Queries 2 & 3 also require the FY window.
- **Edge cases.** A month not in `MONTH_NUM` forces `AND 1=0` (all queries empty). Because Query 1 lacks an FY bound, month-filtered programme counts can include same-month conducted rows from prior FYs while their FY-bound seats/hrs are 0 — a possible count/seat divergence. Default ("All"/empty) applies no month restriction.
- **Source.** `tms/helpers.py:434-440` (clause construction), `:458 / :474-475 / :481 / :505` (usage); route `tms/routes/summary.py:15-18`.

---

## Coverage % (TNI fulfilment) — Monthly Summary table

Coverage answers: *of employees nominated in this FY's TNI for a programme type, what fraction have an attendance record for the matching programme.* Computed in Query 4 of `_calc_summary` and re-derived independently for the TOTAL footer in `_calc_totals`. Colour bands in `summary.html`: ≥100 green ("On track"), ≥75 amber ("Watch"), >0 red ("Critical"), else "—".

### Per-prog_type Blue Collared Coverage % (BC%)

- **What it is.** Of BC employees nominated in this FY's TNI for this prog_type, the fraction with a matching attendance record.
- **Formula.**
  ```sql
  -- Query 4, GROUP BY t.prog_type, e.collar
  bc_fixed = COUNT(*)                                         -- TNI rows
  bc_cum   = SUM(CASE WHEN EXISTS(
               SELECT 1 FROM emp_training et
               WHERE et.emp_code=t.emp_code AND et.plant_id=t.plant_id
                 AND LOWER(et.programme_name)=LOWER(t.programme_name)
             ) THEN 1 ELSE 0 END)
  bc_cov   = round(bc_cum / bc_fixed * 100, 1) if bc_fixed else 0
  ```
- **Inputs.** Denominator from `tni` (`prog_type`, `emp_code`, `plant_id`, `programme_name`, `fy_year`, `source`) JOIN `employees e` (`collar`). Numerator EXISTS against `emp_training`.
- **Filters.** `t.plant_id=?`; `t.fy_year=_fy_label()` (short `'YY-YY'`); `t.source='TNI Driven'` (New Requirement excluded from denominator); `e.collar IN ('Blue Collared','White Collared')`. The numerator EXISTS subquery has **no** additional filter — **not** FY-bound, date/month/status/host-bound: any `emp_training` row (any year, central or local) with the same emp+plant+lowercased programme name satisfies it. The screen's month filter does **not** affect coverage (Query 4 has no month clause).
- **Edge cases.** Zero-division guarded (0 when `bc_fixed=0`). Rounded to 1 dp. The `employees` join does **not** filter `is_active` — exited BC employees count in **both** numerator and denominator (differs from compliance). Non-TNI / New-Requirement attendees count in seats/man-hours but not here. Programme-name match is case-insensitive but otherwise exact string equality — no fuzzy match, so spelling variants between `tni` and `emp_training` fail to fulfil. Cannot exceed 100% (numerator counts only nominated rows).
- **Source.** `tms/helpers.py:514-530` (Query 4), `:553-557` (unpack + `bc_cov`) · `_calc_summary`.

### Per-prog_type White Collared Coverage % (WC%)

- **What it is.** As BC%, for White Collared employees.
- **Formula.** `wc_cov = round(wc_cum / wc_fixed * 100, 1) if wc_fixed else 0`, where `(wc_fixed, wc_cum) = tni_map[pt]['White Collared']` from the same Query 4.
- **Inputs / Filters.** Identical to BC% except collar resolves to `'White Collared'`.
- **Edge cases.** Zero-division guarded; round 1 dp; missing key defaults to `(0,0)`; `is_active` not filtered (exited WC included); case-insensitive exact programme-name match only.
- **Source.** `tms/helpers.py:514-530` (Query 4), `:555,558` (`wc_cov`).

### Per-prog_type Total Coverage % (Tot%)

- **What it is.** Combined BC+WC coverage for the prog_type.
- **Formula.** `tot_cov = round((bc_cum + wc_cum) / (bc_fixed + wc_fixed) * 100, 1) if (bc_fixed + wc_fixed) else 0`. **Pooled** numerator over pooled denominator — **not** an average of BC% and WC%.
- **Inputs.** Same Query 4 values.
- **Filters.** Same as BC%/WC%, both collars combined.
- **Edge cases.** Zero-division guarded on the combined denominator; round 1 dp. Weighted by nomination volume, not a simple mean.
- **Source.** `tms/helpers.py:559` · `_calc_summary`.

### TOTAL-row Coverage % (footer BC / WC / Tot)

- **What it is.** Plant-wide coverage across all programme types.
- **Formula.** **Recomputed from fresh DB queries**, not summed from per-row percentages: `bc_fixed`/`wc_fixed = COUNT(*)` of `tni` rows for the collar; `bc_cum`/`wc_cum = COUNT(*)` of those rows WHERE `EXISTS(matching emp_training)`. Then `bc_cov`/`wc_cov`/`tot_cov` exactly as the per-row formulas (zero-division guarded, round 1 dp). The keys `bc_cov/wc_cov/tot_cov/bc_fixed/wc_fixed/bc_cum/wc_cum` are in the `skip` set so they are **not** naively summed across rows.
- **Inputs.** Four independent `COUNT(*)` queries over `tni JOIN employees`; the cum ones add the EXISTS subquery against `emp_training`.
- **Filters.** Per query: `t.plant_id=?`, `t.fy_year=_fy_label()` (re-derived at line 589), `t.source='TNI Driven'`, `e.collar='Blue Collared'` (or White). Numerator EXISTS has no FY/date/month/status/host filter. FY-wide regardless of the month dropdown.
- **Edge cases.** Zero-division guarded per ratio; round 1 dp. `is_active` not filtered (exited counted). Computed in a second pass, so row-level rounding does not propagate. A fallback branch (db/plant_id None) would sum per-row values — not used by the live routes.
- **Source.** `tms/helpers.py:578-625` (`_calc_totals`); denominators `590-597`, numerators `598-615`, final cov `621-624`.

---

## Compliance % / Man-hour Mandate (Dashboard gauges)

Computed in `_calc_compliance(plant_id, db)` (`tms/helpers.py:628`), surfaced on the SPOC Dashboard (`auth.py:842-845`), the async JSON (`/api/dashboard-monthly`, `api.py:220-242`), and the Central per-plant view (`central.py:769`). The **exit policy** (Global Conventions) governs the asymmetry below.

### Blue Collar Compliance % (bc_pct)

- **What it is.** Of the annual training-hour target owed to all **currently-active** BC staff, the fraction actually delivered this FY. Can exceed 100% by design.
- **Formula.** `bc_pct = round(bc_act / bc_mandate * 100, 1) if bc_mandate else 0`, where `bc_mandate = bc * bc_target`, `bc_act = COALESCE(SUM(et.hrs),0)`, `bc = COUNT(*)` of active BC employees, `bc_target = get_config('mh_target_bc',12,plant_id)`. No DISTINCT — every attendance row's hours sum.
- **Inputs.** Numerator: `emp_training t` (`t.hrs`) JOIN `employees e` ON `emp_code+plant_id`. Denominator: `employees` count. Target: `org_config` via `tms/config.py get_config`.
- **Filters.** Numerator (`helpers.py:636-640`): `t.plant_id=plant_id AND e.collar='Blue Collared' AND t.start_date BETWEEN fy_start AND fy_end`; **no `is_active` filter** (exit-policy numerator). Denominator (`helpers.py:630-632`): `employees WHERE plant_id=? AND is_active=1 AND collar='Blue Collared'`. Target: per-plant override → global → default 12 (bounds 1..8760).
- **Edge cases.** Zero-division guarded (0 if `bc_mandate==0`). EXIT POLICY → gauge can exceed 100%. `COALESCE(SUM,0)` treats NULL hrs as 0. Round 1 dp. Inner JOIN drops orphan attendance (emp not in `employees`). **No source/status filter** — TNI Driven and New Requirement hours both count.
- **Source.** `tms/helpers.py:651` (bc_pct), `:636-640` (bc_act), `:630-632` (bc), `:649` (mandate), `:647` (target) · `_calc_compliance` (`tms/helpers.py:628`).

### White Collar Compliance % (wc_pct)

- **What it is.** As BC%, for White Collared staff (default target 24). Same exit-policy treatment.
- **Formula.** `wc_pct = round(wc_act / wc_mandate * 100, 1) if wc_mandate else 0`, `wc_mandate = wc * wc_target`, `wc_target = get_config('mh_target_wc',24,plant_id)`.
- **Inputs.** Numerator: `emp_training t` JOIN `employees e`, collar='White Collared'. Denominator: active WC count × target.
- **Filters.** Numerator (`helpers.py:641-645`): `t.plant_id=? AND e.collar='White Collared' AND t.start_date BETWEEN fy_start AND fy_end`; no `is_active`. Denominator (`helpers.py:633-635`): `employees WHERE plant_id=? AND is_active=1 AND collar='White Collared'`.
- **Edge cases.** Zero-division guarded; can exceed 100% (exit policy); `COALESCE(SUM,0)`; round 1 dp; orphan attendance dropped; no status/source filter.
- **Source.** `tms/helpers.py:652` (wc_pct), `:641-645` (wc_act), `:633-635` (wc), `:650` (mandate), `:648` (target) · `_calc_compliance`.

### Total / Combined Compliance % (total_pct)

- **What it is.** Combined BC+WC gauge.
- **Formula.** `total_pct = round((bc_act + wc_act) / (bc_mandate + wc_mandate) * 100, 1) if (bc_mandate + wc_mandate) else 0`. Pooled ratio — **not** an average of `bc_pct`/`wc_pct`.
- **Inputs.** Reuses `bc_act, wc_act, bc_mandate, wc_mandate`.
- **Filters.** Inherits all bc/wc filters (FY window, is_active=1 denominators, per-collar, plant scope, exit-policy asymmetry).
- **Edge cases.** Zero-division guarded; weighted by mandate size; can exceed 100%. The API separately computes `total_mh_pct = round(total_actual / total_target * 100, 1)` from the **rounded** `bc_actual`/`wc_actual` (each round()'d to 1 dp at `helpers.py:669`), so it can differ from `total_pct` (which uses raw sums) by a rounding hair.
- **Source.** `tms/helpers.py:653-654` (total_pct); recomputed at `tms/routes/api.py:223-229` (total_mh_pct) · `_calc_compliance`.

### Headline Compliance % and RAG (headline_pct / headline_rag)

- **What it is.** The single "worst of the two" dashboard headline — the lower of BC% and WC%, ignoring any collar with no active employees, with a traffic-light band.
- **Formula.** `candidates = [bc_pct if bc>0] + [wc_pct if wc>0]; headline_pct = min(candidates) if candidates else 0`. RAG: ≥75 → `'on-track'`; ≥50 → `'watch'`; else `'critical'`.
- **Inputs.** `bc_pct, wc_pct, bc, wc` from `_calc_compliance`.
- **Filters.** A collar is included in the MIN only if it has >0 active employees. Inherits upstream FY/collar/plant/is_active filters.
- **Edge cases.** Both collars 0 active → `headline_pct=0` → `'critical'`. Uses MIN, not average (a strong collar cannot mask a weak one). Thresholds 75/50 are hardcoded (not config-driven). Inherits the >100% possibility.
- **Source.** `tms/helpers.py:655-665` · `_calc_compliance`.

### BC/WC Man-hour Mandate (target hours owed)

- **What it is.** Hours the plant is expected to deliver this year: active BC × BC target + active WC × WC target. Denominator of all gauges.
- **Formula.** `bc_mandate = bc * bc_target; wc_mandate = wc * wc_target; target_hrs = bc_mandate + wc_mandate`. `bc_target=get_config('mh_target_bc',12,plant_id)`, `wc_target=get_config('mh_target_wc',24,plant_id)`.
- **Inputs.** `employees` (active counts per collar) and `org_config`.
- **Filters.** Headcount: `employees WHERE plant_id=? AND is_active=1 AND collar=<collar>`. Target resolution: plant override → global → default (12/24); validated 1..8760 on write.
- **Edge cases.** Targets are per-plant configurable (multi-tenant), not a global constant. Bad stored config falls back to default with a logged warning. No active employees in a collar → that mandate is 0 → gauge returns 0. Exited employees do **not** add to the mandate (`is_active=1`) — the denominator side of the exit policy.
- **Source.** `tms/helpers.py:646-650` (target + mandate); target source `tms/config.py get_config`; `target_hrs` aggregation `tms/routes/auth.py:831`.

### BC/WC Actual man-hours delivered (gauge numerator)

- **What it is.** Total FY training hours logged per collar — **including** hours for employees who have since left.
- **Formula.** `bc_actual = round(bc_act, 1); wc_actual = round(wc_act, 1)`, where `bc_act/wc_act = COALESCE(SUM(t.hrs),0)`. `total_actual = round(bc_actual + wc_actual, 1)` (`api.py:227`).
- **Inputs.** `emp_training t` (`t.hrs`) JOIN `employees e` ON `emp_code+plant_id`; collar from `employees`.
- **Filters.** `t.plant_id=plant_id`; `e.collar=<collar>`; `t.start_date BETWEEN fy_start AND fy_end`. **No `e.is_active` filter** (exit-policy numerator). No source/status filter.
- **Edge cases.** `COALESCE(SUM,0)` → 0 on no rows/NULL hrs. Per-row sum (no DISTINCT). Inner JOIN drops orphans. FY window is by `start_date`, not row-creation time. Rounded to 1 dp before exposure.
- **Source.** `tms/helpers.py:636-640` (bc_act), `:641-645` (wc_act), `:669` (round); FY bounds `:68-73`.

### Central Per-Plant Compliance % (on the Central dashboard)

- **What it is.** The same BC/WC compliance, computed for **every** plant in one batched pass for the central overview — a separate inline implementation, **not** a call to `_calc_compliance`.
- **Formula.** `bc_pct = round(bc_hrs / bc_mandate * 100, 1) if bc_mandate else 0` (WC analogous); `bc_mandate = bc * get_config('mh_target_bc',12,plant_id=pid)`. Ranking key = `(bc_pct + wc_pct)/2`.
- **Inputs.** Numerators `bc_hrs`/`wc_hrs`: batched `emp_training t` JOIN `employees e` (`CASE WHEN e.collar=... THEN t.hrs`, `central.py:91-100`). Denominators: batched active-employee counts (`central.py:65-70`). Targets: `org_config`.
- **Filters.** Numerator (`central.py:97`): `t.start_date BETWEEN fy_start AND fy_end`; JOIN on emp_code+plant_id; collar via CASE; **no `is_active`** on numerator (exit policy preserved). Denominator (`central.py:69`): `employees WHERE is_active=1 GROUP BY plant_id`. Per-plant target via `get_config(plant_id=pid)`.
- **Edge cases.** Mirrors `_calc_compliance` exit policy (can exceed 100%) but is a **separate** implementation — any change to the helper must be mirrored here (consistency risk). Zero-division guarded. Ranking uses an arithmetic **mean** of the two pct values (`central.py:126`), differing from the SPOC headline **MIN** logic. The un-joined `mh_all` scan (`central.py:84-87`) sums ALL `emp_training.hrs` (incl. orphans) for the `manhours` display column, so the manhours total can exceed `bc_hrs+wc_hrs`.
- **Source.** `tms/routes/central.py:115-116` (bc_pct/wc_pct), `:113-114` (mandates), `:91-100` (collar man-hour numerator), `:65-70` (active headcount), `:126` (ranking mean).

---

## Worst Cells / Gap Analysis (TNI coverage by prog_type × collar)

`_calc_worst_cells(plant_id, db, limit=3, min_nominated=3)` (`tms/helpers.py:679`) is called inside `_calc_compliance` and rendered into the Dashboard "Improvement Areas — Worst Cells" panel (server-rendered only; **not** exposed by `/api/dashboard-monthly`).

### Worst cell coverage % (cov_pct / pct)

- **What it is.** Per `(prog_type, collar)` cell, the share of nominated employees actually trained on that type; "worst cells" are the lowest-coverage combinations.
- **Formula.**
  ```sql
  cov_pct = ROUND(100.0 * SUM(CASE WHEN trained THEN 1 ELSE 0 END)
                  / NULLIF(COUNT(*),0), 1)
  GROUP BY prog_type, collar
  -- 'trained' is a per-TNI-row EXISTS boolean:
  EXISTS(SELECT 1 FROM emp_training et
         JOIN programme_details pd ON pd.session_code = et.session_code
         WHERE et.emp_code = t.emp_code
           AND pd.programme_name = t.programme_name
           AND pd.plant_id = t.plant_id)
  ORDER BY cov_pct ASC LIMIT 3
  ```
  Numerator = trained TNI rows in the cell; denominator = `COUNT(*)` of TNI rows (nominated). Output keys `pct` and `cov_pct` are the same value; `trained` and `trained_cnt` are the same value.
- **Inputs.** `tni t` JOIN `employees e` ON `emp_code+plant_id` (collar, is_active); `trained` subquery joins `emp_training et` to `programme_details pd` on `session_code`, matched on `programme_name` and `pd.plant_id = t.plant_id`.
- **Filters.** `t.plant_id = ?`; **`e.is_active = 1`** (nominations of inactive/exited employees excluded from both numerator and denominator); `GROUP BY prog_type, collar`; `HAVING nominated >= min_nominated` (default 3, cells with <3 suppressed); `LIMIT` (default 3). **No `fy_year` filter on tni and no date filter on emp_training/programme_details** — coverage here is **lifetime/all-time**, not FY-bound. **No `t.source` filter** — both TNI Driven and other sources counted (unlike Monthly Summary coverage which filters source). The trained-match is `programme_name` equality (exact, case-sensitive in SQL) and `pd.plant_id = t.plant_id`.
- **Edge cases.** `NULLIF(COUNT(*),0)` guards divide-by-zero; `HAVING nominated>=3` makes COUNT(*) ≥3 so `cov_pct` is never NULL via this path; Python falls back `r['cov_pct'] or 0`. Whole function wrapped in `try/except` returning `[]` on any DB error (silent failure → empty panel). `min_nominated`/`limit` are parameters but the only caller uses defaults (not per-plant configurable). Active-only denominator means an exited employee's nomination is removed (opposite of the compliance exit policy). The `trained` EXISTS does **not** restrict by date or `host_plant_id`, so any historical attendance counts as trained. RAG per cell: ≥75 `'on-track'`, ≥50 `'watch'`, else `'critical'`. Ties broken by SQLite's arbitrary order (no secondary sort).
- **Source.** `tms/helpers.py:679-733` (SQL `686-709`; post-processing `714-733`); called from `tms/helpers.py:675`.

### Worst cell 'nominated' count (denominator)

- **What it is.** How many employees were nominated (have a TNI entry) for that prog_type+collar — the size the coverage % is measured against.
- **Formula.** `nominated = COUNT(*)` over the inner per-TNI-row set, grouped by `(prog_type, collar)`. One TNI row (one emp + programme) = one unit. `trained_cnt = SUM(CASE WHEN trained THEN 1 ELSE 0 END)`; Coverage% = `trained_cnt / nominated * 100`.
- **Inputs.** `tni t` JOIN `employees e` (collar, is_active).
- **Filters.** `t.plant_id = ?`; `e.is_active = 1`; `GROUP BY prog_type, collar`; `HAVING nominated >= 3`. No fy_year/source/date filters. **No DISTINCT** — counts raw TNI rows; one employee with multiple programmes of the same prog_type contributes multiple to `nominated` (tni UNIQUE is on `plant_id+emp_code+programme_name`).
- **Edge cases.** `nominated` is a count of (employee × programme) demands, **not** a headcount. Cells under `min_nominated=3` are hidden (not shown as 0%). Inactive employees' nominations excluded.
- **Source.** `tms/helpers.py:690` (COUNT(*) AS nominated), `:706` (HAVING), `:730` (output).

---

## Dashboard APIs — Monthly Chart + Man-hour Drilldown

`/api/dashboard-monthly` (`tms/routes/api.py`, `api_dashboard_monthly()`) feeds the 12-bucket Apr→Mar chart and the gauge block; `/api/manhour-drilldown` feeds the gauge click-through.

### Per-month Seats

- **What it is.** Per calendar month, the number of attendance records (one seat per person per session). **All-time** for the plant — not limited to the current FY.
- **Formula.** `seats = COUNT(*)` of `emp_training` rows `GROUP BY strftime('%m', start_date)`; per-month via `mo_map.get(mo_num, {seats:0})`. Iterated in FY display order [04..03]; missing months → 0.
- **Inputs.** `emp_training` (`start_date`, row count). Scoped by `emp_training.plant_id = session['plant_id']`.
- **Filters.** `WHERE plant_id=? AND start_date IS NOT NULL AND start_date != ''`. **No FY window**, no status/collar/source/is_active filter.
- **Edge cases.** Cross-year contamination: 03/2024 and 03/2026 both land in "March". Rows with empty/NULL start_date dropped. `seats` not rounded. `host_plant_id` not considered — central-hosted attendees stored under home plant appear under the home plant only.
- **Source.** `tms/routes/api.py:181-189, 213-218`, `api_dashboard_monthly()`.

### Per-month Man-hours

- **What it is.** Total man-hours per month, summed across attendance, **all-time** (not FY-bound).
- **Formula.** `hrs = COALESCE(SUM(emp_training.hrs), 0) GROUP BY strftime('%m', start_date)`, then `round(hrs, 1)`; per-month via `mo_map.get(mo_num, {hrs:0.0})`.
- **Inputs.** `emp_training.hrs`, `start_date`. Plant scoped.
- **Filters.** `WHERE plant_id=pid AND start_date IS NOT NULL AND start_date != ''`. No FY/status/collar/source/is_active filter (differs from the FY-bound compliance block in the **same** endpoint).
- **Edge cases.** Same cross-year collision. NULL hrs → 0 via COALESCE. Round 1 dp. Does **not** reconcile with the gauge totals (gauge is FY-bound; these bars are all-time).
- **Source.** `tms/routes/api.py:181-189, 214-216`.

### Per-month Sessions Conducted

- **What it is.** Sessions with a 2C details row in each month, counting only sessions on this plant's calendar.
- **Formula.** `sessions = COUNT(*)` of `programme_details pd JOIN calendar c ON pd.session_code=c.session_code GROUP BY strftime('%m', pd.start_date)`; per-month via `sess_map.get(mo_num, 0)`.
- **Inputs.** `programme_details` (`session_code`, `start_date`) JOIN `calendar` (`session_code`, `plant_id`).
- **Filters.** `WHERE c.plant_id=pid AND pd.start_date IS NOT NULL AND pd.start_date != ''`. INNER JOIN on `session_code` **excludes** orphan 2C rows. **No explicit `status='Conducted'` filter** — presence of a `programme_details` row is treated as conducted. No FY window.
- **Edge cases.** Inner join drops conducted sessions whose `session_code` is missing from `calendar` (ad-hoc 2C), so this can undercount vs a raw `programme_details` count. Cross-year collision. Counts on `pd.start_date`, not `calendar.plan_start`.
- **Source.** `tms/routes/api.py:192-199, 217`.

### Per-month Sessions Planned

- **What it is.** Calendar rows with a planned start date in each month for this plant.
- **Formula.** `planned = COUNT(*)` of `calendar` rows `GROUP BY strftime('%m', plan_start)`; per-month via `plan_map.get(mo_num, 0)`.
- **Inputs.** `calendar.plan_start`, `calendar.plant_id`.
- **Filters.** `WHERE plant_id=pid AND plan_start IS NOT NULL AND plan_start != ''`. **No status filter** (planned + conducted + to-be-planned all counted). No FY/source/audience filter.
- **Edge cases.** Buckets by `plan_start` month only — cross-year collision. Empty `plan_start` dropped. Counts all calendar rows regardless of status.
- **Source.** `tms/routes/api.py:201-206, 218`.

### BC / WC / Total compliance % in the monthly JSON

- **What it is.** The man-hour gauges, surfaced via the API: `compliance.bc_pct` (alias `bc_mh_pct`), `compliance.wc_pct` (alias `wc_mh_pct`), `compliance.total_pct`, and a separately recomputed `compliance.total_mh_pct`.
- **Formula.** Values come from `_calc_compliance` (see Compliance section): `bc_pct`/`wc_pct` as documented there; the API exposes each twice under both names (identical values). `total_mh_pct = round(total_actual / total_target * 100, 1)` where `total_actual = round(bc_actual + wc_actual, 1)` (pre-rounded collar actuals) and `total_target = bc_mandate + wc_mandate`; `total_pct` (from `_calc_compliance`) uses **un-rounded** raw act sums.
- **Inputs / Filters.** Inherit all `_calc_compliance` filters: FY-bound numerators (`start_date BETWEEN _current_fy()`), collar split, plant scope; `is_active=1` denominators; exit-policy asymmetry.
- **Edge cases.** Can exceed 100% (exit policy). `total_mh_pct` can differ from `total_pct` by a rounding hair (pre-rounded vs raw). Zero-division guarded throughout.
- **Source.** `tms/helpers.py:628-654` (`_calc_compliance`); surfaced `tms/routes/api.py:220-239`.

### Gauge labels + cur_mo

- **What it is.** Raw hour totals and mandates behind each gauge plus the current-month highlight.
- **Formula.** `bc_actual=round(bc_act,1)`; `bc_target=bc_mandate=bc_headcount×config_bc`; WC analogous; `total_actual=round(bc_actual+wc_actual,1)`; `total_target=bc_target+wc_target`. `cur_mo = _now_ist().strftime('%m')`.
- **Inputs.** `_calc_compliance` dict; `cur_mo` from IST wall clock.
- **Filters.** Actuals FY-bound + collar + plant; targets from active headcount × per-plant config.
- **Edge cases.** `cur_mo` uses `_now_ist()` (not UTC), preventing the 18:30-IST month rollover bug. `total_actual` double-rounds (sum of already-rounded collar actuals). Targets are integers from config (validated 1..8760).
- **Source.** `tms/routes/api.py:221-242`; `_now_ist` `tms/helpers.py:129-138`.

### Man-hour Drilldown — Department actual hours (FY-bound)

- **What it is.** Per department, total FY man-hours delivered to active employees of the selected collar.
- **Formula.**
  ```sql
  SELECT e.department, e.collar, COUNT(DISTINCT e.emp_code) emp_count,
         COALESCE(SUM(et.hrs),0) actual_hrs
  FROM employees e
  LEFT JOIN emp_training et
    ON et.emp_code=e.emp_code AND et.plant_id=e.plant_id
       AND et.start_date BETWEEN fy_start AND fy_end
  ... GROUP BY e.department, e.collar
  -- dept_map accumulates: d['actual_hrs'] += r['actual_hrs']; rounded to 1 dp at output
  ```
- **Inputs.** `employees` (department, collar, emp_code, plant_id, is_active) LEFT JOIN `emp_training` (hrs, start_date).
- **Filters.** `WHERE e.plant_id=pid AND e.is_active=1 [AND e.collar=<BC|WC>] AND e.department IS NOT NULL AND e.department!=''`. The FY window is in the **JOIN ON** (not WHERE), so out-of-FY hours simply don't join and the employee row survives via LEFT JOIN.
- **Edge cases.** Employees with zero FY hours still appear with `actual_hrs=0`. Blank-department employees **excluded** here (but included in the employee-level list). **Only `is_active=1`** — exited staff hours are **not** shown in the drilldown (drilldown differs from the gauge numerator, which keeps exited hours). `host_plant_id` not matched.
- **Source.** `tms/routes/api.py:266-288`, `api_manhour_drilldown()`.

### Man-hour Drilldown — Department target hours + pct

- **What it is.** Each department's mandated hours and compliance %.
- **Formula.** `target_hrs += emp_count × (bc_t if collar='Blue Collared' else wc_t if collar='White Collared' else 0)`, accumulated across the dept's collar-rows; `pct = round(actual_hrs/target_hrs*100,1) if target_hrs else 0`. `bc_t=get_config('mh_target_bc',12,plant_id=pid)`, `wc_t=get_config('mh_target_wc',24,plant_id=pid)`.
- **Inputs.** `emp_count = COUNT(DISTINCT e.emp_code)` per dept+collar; config targets.
- **Filters.** `emp_count` from active employees, collar-scoped by the collar param. For `collar='ALL'`, both BC and WC rows contribute their respective per-collar targets.
- **Edge cases.** Any collar value other than the two canonical strings contributes 0 to target (defensive). `target_hrs=0 → pct=0`. Per-plant config honoured so the drilldown reconciles with the gauge mandate. Round 1 dp.
- **Source.** `tms/routes/api.py:256-257, 280-295`.

### Man-hour Drilldown — Employee actual / target / pct / status

- **What it is.** Per active employee: FY hours, individual annual target (by collar), % achieved, and status (Met / Zero / Low).
- **Formula.** `actual = round(COALESCE(SUM(et.hrs),0),1)` (LEFT JOIN, FY-bound). `tgt = bc_t if collar='Blue Collared' else wc_t if collar='White Collared' else 0`. `pct = round(actual/tgt*100,1) if tgt else 0`. `status = 'Met' if pct>=100 else ('Zero' if actual==0 else 'Low')`.
- **Inputs.** `employees` (emp_code, name, department, collar) LEFT JOIN `emp_training` (hrs) with the FY date window in the ON; config targets.
- **Filters.** `WHERE e.plant_id=pid AND e.is_active=1 [AND collar filter when BC/WC]`. `emp_training` joined only for `et.start_date BETWEEN _current_fy()`. **No department-not-blank filter** here (unlike the department table) — blank-department employees are included (department shown as `''`).
- **Edge cases.** Status order: `pct>=100` wins first. `'Zero'` requires `actual == 0` after rounding. `tgt=0` (unknown/blank collar) → `pct=0`, so status is `'Zero'` if `actual==0` else `'Low'` (never `'Met'`). Exit policy **not** applied — only `is_active=1` listed. `host_plant_id` not matched. Tiny <0.05 hours round to 0.0 and read as `'Zero'`.
- **Source.** `tms/routes/api.py:297-317`.

### Man-hour Drilldown — FY window & collar param decoding

- **What it is.** The shared FY window and collar-chip → SQL mapping for both drilldown tables.
- **Formula.** `fy_start, fy_end = _current_fy()`. `collar = request.args.get('collar','ALL').upper()`; `collar_where = "AND e.collar='Blue Collared'" (BC) | "AND e.collar='White Collared'" (WC) | '' (anything else incl. ALL)`.
- **Inputs.** `_current_fy()` via `_today_ist()` (IST); `collar` arg.
- **Filters.** FY window injected as `et.start_date BETWEEN fy_start AND fy_end` inside the LEFT JOIN ON of both queries; `collar_where` appended to the employees WHERE.
- **Edge cases.** FY computed in IST (no 5.5h early roll on Render). Unknown collar param falls through to `''` = ALL behaviour (no error). FY filter is in JOIN ON, not WHERE — preserves zero-hour employees via LEFT JOIN.
- **Source.** `tms/routes/api.py:250-264`; `_current_fy` `tms/helpers.py:68-73`.

---

## QC Dashboard Analytics

`/api/dashboard-qc` (`tms/routes/auth.py:849`, `@spoc_required`) builds six charts for the logged-in plant and current FY, with two shared pre-passes and a 60s in-process cache.

### Orchestrator + 60s cache

- **What it is.** One JSON endpoint that builds all six QC charts; cached per `(plant_id, fy_start)` for 60s.
- **Formula.** Builds `out` by calling, in order, `_qc_pareto, _qc_histogram, _qc_dept_compliance, _qc_cumulative, _qc_cumulative_hours, _qc_heatmap`. Shared pre-passes: `emp_rows=_emp_fy_hours(...)` (histogram + dept), `trained=_trained_pairs(...)` (pareto + heatmap). Cache: `key=(plant_id, fy_start)`; return cached payload if its expiry > now, else compute and store `(now+60, out)`.
- **Inputs.** `session['plant_id']`; `_current_fy()`; module globals `_QC_CACHE`, `_QC_TTL` (60). Tables: `employees`, `emp_training`, `programme_details`, `tni`; `org_config`.
- **Filters.** Plant scope = `session['plant_id']` (single-plant; **not** central/host-aware). FY window = `_current_fy()` (IST). `@spoc_required`.
- **Edge cases.** Cache is per-worker and time-only (no event invalidation) — a SPOC who edits data may see the old chart for up to 60s. Cache key omits `fy_end` (derived from the same FY). No `try/except` around aggregations — a DB error 500s rather than returning partial JSON.
- **Source.** `tms/routes/auth.py:849-874`; cache config `:21-22`.

### `_emp_fy_hours` (shared per-employee FY hours)

- **What it is.** Per active BC/WC employee, total FY training hours; never-trained employees get 0.
- **Formula.**
  ```sql
  SELECT e.collar, COALESCE(NULLIF(e.department,''),'Unassigned') AS dept,
         COALESCE(SUM(t.hrs),0) AS hrs
  FROM employees e
  LEFT JOIN emp_training t
    ON t.emp_code=e.emp_code AND t.plant_id=e.plant_id
       AND t.start_date BETWEEN fy_start AND fy_end
  WHERE e.plant_id=? AND e.is_active=1
    AND e.collar IN ('Blue Collared','White Collared')
  GROUP BY e.emp_code
  ```
- **Inputs.** `employees`, `emp_training`.
- **Filters.** `plant_id`; `is_active=1`; collar IN the two values; FY date filter in the JOIN ON. **No status/source filter** — all FY attendance counts (incl. New Requirement).
- **Edge cases.** Empty/NULL department → `'Unassigned'`. NULL hrs → 0. Inactive employees entirely excluded (denominator = active headcount). Blank/unrecognised collar excluded.
- **Source.** `tms/routes/auth.py:60-75`.

### `_trained_pairs` (shared trained-set)

- **What it is.** Distinct set of `(emp_code, programme_name)` the plant has trained (attended ≥1 session of that programme), used to decide who is "covered".
- **Formula.** Python set from `SELECT DISTINCT et.emp_code, pd.programme_name FROM emp_training et JOIN programme_details pd ON pd.session_code=et.session_code AND pd.plant_id=et.plant_id WHERE et.plant_id=?`.
- **Inputs.** `emp_training`, `programme_details`.
- **Filters.** `plant_id`. **No FY filter** — all-time trained pairs (unlike `_emp_fy_hours`). Inner JOIN means attendance with no matching `programme_details` row (2A without a 2C) is dropped.
- **Edge cases.** Duplicates collapse (DISTINCT). A pair trained in a prior FY still counts as trained (Pareto/Heatmap are cumulative, not FY-bound).
- **Source.** `tms/routes/auth.py:78-88`.

### QC Pareto — uncovered headcount per programme

- **What it is.** Per programme, how many nominated employees have **not** been trained on it; top 8 gaps, largest first.
- **Formula.** For each TNI nomination `(prog, emp)` with the employee active: if `(emp, prog) NOT in trained-set` → `unc[prog] += 1`. Then `top = sort items with unc>0 by (-count, prog_name)[:8]`. Output `paretoLabels`/`paretoData`.
- **Inputs.** `tni`, `employees`, trained-set from `_trained_pairs`.
- **Filters.** `tni: plant_id=? AND fy_year=_fy_label()`. `employees: is_active=1`. Coverage universe = TNI nominations only. Trained-set has **no FY bound**, so any-time training counts as covered. Only programmes with >0 uncovered shown; cap 8.
- **Edge cases.** Inactive nominee dropped (inner join `is_active=1`). Tie-break alphabetical. Nominations FY-filtered but "trained" is all-time — an emp trained in a prior FY counts as covered for a current-FY nomination. Empty result → empty arrays.
- **Source.** `tms/routes/auth.py:91-112` (`_fy_label()` at `:95`; uncovered test `:105`).

### QC Histogram — employees per FY-hours bucket (BC vs WC) + avg + target lines

- **What it is.** Distribution of employees by FY hours in 8 fixed buckets, split BC/WC; plus per-collar average hours/employee and the annual target for reference lines.
- **Formula.** `EDGES=[2,4,6,8,12,16,24, inf]`. For each emp row, `h=hrs or 0`; `idx = first i where h<EDGES[i], else 7` (h≥24 → bucket 7). Increment `bc[idx]`/`wc[idx]` by collar; accumulate sums/counts. `histAvgBc=round(bc_sum/bc_n,1) if bc_n else 0` (WC analogous). `histTargetBc=get_config('mh_target_bc',12,plant_id)`, `histTargetWc=get_config('mh_target_wc',24,plant_id)`.
- **Inputs.** `emp_rows` from `_emp_fy_hours`; `org_config`.
- **Filters.** Inherits `_emp_fy_hours` filters. Bucketing is strict-less-than the upper edge (0-2 = h<2; 24+ = h≥24). Collar branch: `'Blue Collared'` → bc, else → wc (already restricted to the two collars upstream).
- **Edge cases.** Never-trained employees (h=0) → bucket 0. Average denominator = count of active employees of that collar (incl. 0-hour). Division guarded. Targets per-plant config with fallbacks 12/24.
- **Source.** `tms/routes/auth.py:115-143` (EDGES `:121`; bucket loop `:125-136`; avg+target `:137-143`).

### QC Department Compliance — % of dept meeting collar target

- **What it is.** Per department, how many active employees hit their own collar's annual target vs fell short, and the % who met it.
- **Formula.** `bc_t=get_config('mh_target_bc',12,plant_id)`, `wc_t=get_config('mh_target_wc',24,plant_id)`. Per emp row: `target = bc_t if collar=='Blue Collared' else wc_t`; `if (hrs or 0) >= target: dept.met += 1 else dept.below += 1`. Per dept: `total=met+below; pct = round(met/total*100) if total else 0`. Output sorted by `pct` descending.
- **Inputs.** `emp_rows` from `_emp_fy_hours`; config targets.
- **Filters.** Inherits `_emp_fy_hours` filters. Met threshold is `>=` target. Department already normalised to `'Unassigned'` for blank/NULL.
- **Edge cases.** `>=` means exactly hitting target counts as met. `pct` uses `round()` (integer percent, Python banker's rounding). Department with 0 employees cannot occur (built from emp rows). Inactive employees excluded from numerator and denominator.
- **Source.** `tms/routes/auth.py:146-165` (target `:156`; met/below `:158`; pct `:163`; sort `:164`).

### QC Heatmap — prog_type × collar TNI coverage % matrix

- **What it is.** A 2-row (BC/WC) × 6-column (prog_type) grid of coverage %, each cell colour-banded.
- **Formula.** `DISPLAY_COLS = [Technical, EHS/HR, Behavioural/Leadership, Cane, Commercial, IT]`; rows `['Blue Collared','White Collared']`. For each TNI row: `k=(collar, prog_type); tot[k]+=1; if (emp,prog) in trained-set: cov[k]+=1`. Cell `= round(100*cov[k]/tot[k],1) if tot[k] else None`. Display: `pct = int(round(cell)) if not None else 0`; `cls = _qc_hclass(pct)`.
- **Inputs.** `tni`, `employees`, trained-set from `_trained_pairs`.
- **Filters.** `tni: plant_id=?`. `employees: is_active=1`. **No `fy_year` filter on TNI** here (unlike Pareto/Cumulative) — the heatmap counts **all** TNI rows for the plant regardless of FY. Trained-set is all-time. Coverage universe = TNI nominations only.
- **Edge cases.** Denominator per cell = nominated rows for that `(collar, prog_type)`; never averaged across cells. A 0-nomination cell → matrix `None` → displayed `pct=0` with class `'h0'` (indistinguishable from a real 0%). NULL collar/prog_type coerced to `''` (won't match a column, effectively hidden). Final `pct = int(round(...))` loses the .1 precision. **This chart has no FY filter on TNI** — a key difference vs Pareto.
- **Source.** `tms/routes/auth.py:263-301` (TNI query `:279-285` no fy_year; cell math `:291-292`; display `:294-300`).

### `_qc_hclass` — heatmap colour bands

- **What it is.** Maps a coverage percent to one of six bands.
- **Formula.** `pct>=100 → 'h100'; >=90 → 'h90'; >=75 → 'h75'; >=50 → 'h50'; >=25 → 'h25'; else 'h0'`. Bands: h0=[0,25), h25=[25,50), h50=[50,75), h75=[75,90), h90=[90,100), h100=[100,∞).
- **Inputs.** `pct` (integer from `_qc_heatmap`).
- **Filters.** None — pure function.
- **Edge cases.** Thresholds are `>=` (lower-inclusive). 99.x rounds to 100 (heatmap passes `int(round(pct))`) → `h100`. No-nomination cells arrive as 0 → `h0`.
- **Source.** `tms/routes/auth.py:253-260`.

### QC Cumulative Coverage — BC/WC TNI coverage across 12 FY months

- **What it is.** For each FY month-end (Apr..Mar), the running % of nominated employees (separately BC and WC) trained on their nominated programme by that date. Monotonic non-decreasing.
- **Formula.** `month_ends` = 12 hard-coded month-end dates from the fy_start year (Apr30..Mar31, **Feb=YYYY-02-28**). `trained = {(emp,prog): MIN(et.start_date)}` over `emp_training JOIN programme_details`, FY-bound. For each TNI nomination: `denom[collar]+=1; firsts[collar].append(trained.get((emp,prog)))`. Per collar per month_end `me`: `round(100.0 * count(ft for ft in firsts if ft and ft<=me) / denom, 1)`; if `denom==0` → `[0]*12`.
- **Inputs.** `emp_training`, `programme_details`, `tni`, `employees`.
- **Filters.** Trained query: `plant_id=? AND et.start_date BETWEEN fy_start AND fy_end` (**FY-bound**, unlike `_trained_pairs`). Nominations: `tni.plant_id=? AND tni.fy_year=_fy_label() AND employees.is_active=1 AND collar IN (BC,WC)`. Separate denominators per collar (never averaged).
- **Edge cases.** Feb month-end hard-coded `YYYY-02-28` — does **not** handle leap years (a Feb-29 training in a leap FY is only picked up from the Mar cutoff). MIN(start_date) is the coverage event date. `denom=0` → all zeros. FY-bound here, so prior-FY training does **not** count (differs from Pareto/Heatmap).
- **Source.** `tms/routes/auth.py:168-215` (month_ends `:173-178`; trained MIN `:183-190`; denom/firsts `:192-203`; cum_series `:205-210`).

### QC Cumulative Man-Hours — delivered vs target vs even-pace

- **What it is.** Running total of man-hours by each FY month-end, against the flat annual target and an even-pace line.
- **Formula.** `bc_t=get_config('mh_target_bc',12)`, `wc_t=get_config('mh_target_wc',24)`; active `bc_n`/`wc_n` from `employees`; `target = bc_n*bc_t + wc_n*wc_t`. `monthly[12]` (Apr=0..Mar=11): per FY `emp_training` row, `m=int(start_date[5:7]); idx = m-4 if m>=4 else m+8; monthly[idx] += hrs`. Cumulative running sum → `hoursCumulative = [round(run) per month]`. `hoursPace[i] = round(target*(i+1)/12)`. `hoursTarget = int(target)`.
- **Inputs.** `employees`, `emp_training`, `org_config`.
- **Filters.** Headcount: `plant_id=? AND is_active=1 AND collar IN (BC,WC)`. Man-hours: `emp_training.plant_id=? AND start_date BETWEEN fy_start AND fy_end`. No source/status filter (all FY hours, incl. New Requirement). Pace = linear `target/12` per month.
- **Edge cases.** Rows with no/short (<7 char) start_date skipped. Month parsed from chars `[5:7]` (assumes ISO `YYYY-MM-DD`). `idx` mapping Apr(4)→0..Dec(12)→8, Jan(1)→9, Feb(2)→10, Mar(3)→11; out-of-range idx skipped. hrs COALESCEd to 0 in SQL, re-OR'd in Python. **Target uses current active headcount** (snapshot at request time), not as-of each month — an employee added/removed mid-year shifts the whole target line. Cumulative values are `round()`'d (integer hours displayed).
- **Source.** `tms/routes/auth.py:218-250` (target `:224-232`; bucketing `:234-244`; cum/pace `:245-250`).

---

## Feedback Aggregates & CSAT (QR participant feedback)

Participant feedback (POST `/q/<token>/feedback`) is 9 Likert questions (1–4). The per-session Feedback Report computes per-question stats and three subtotals; aggregates also fold into the conducted 2C row.

### Question scale & per-response validity (q1–q9)

- **What it is.** Each participant answers 9 questions on a 1–4 scale (1=Strongly Disagree … 4=Strongly Agree). An answer counts only if it is an integer in `[1,4]`; anything else (blank, 0, out of range) is dropped.
- **Formula.** On submit, `_r(name)`: `v=int(form[name])`, kept only if `1<=v<=4` else `None`. On the report, `_analyse` keeps `r[field]` only if `r[field] and 1<=r[field]<=4`. The 9 fields in order: q1=`q_obj_explained`, q2=`q_well_structured`, q3=`q_content_appropriate`, q4=`q_presentation_quality`, q5=`q_time_reasonable`, q6=`q_inputs_appropriate`, q7=`q_communication_clear`, q8=`q_queries_responded`, q9=`q_well_involved`.
- **Inputs.** `feedback_response.q_obj_explained..q_well_involved` (INTEGER, `schema.sql:153-161`). Form fields q1..q9.
- **Filters.** Per-question validity: non-null integer `1<=v<=4`. NULL/0/out-of-range silently excluded. No FY/month/status/collar filter at the per-response level. Scope is per `(plant_id, session_code)`.
- **Edge cases.** 0 is excluded (treated as "no answer"). Each question analysed independently — a response may contribute to some questions and not others (partial responses allowed). Anonymous responses deduped by IP. Storage uses `INSERT OR REPLACE` keyed on `UNIQUE(plant_id,session_code,emp_code)` — a re-submit by the same emp overwrites.
- **Source.** `tms/routes/qr.py:811-816` (`_r`), `:392-405` (`_analyse` filter); `schema.sql:147-167`.

### Per-question analysis — SD/D/A/SA counts, average, score %

- **What it is.** Per question: counts of each option, the average rating, and a percentage out of max 4.
- **Formula.** `_analyse(field,rows)`: `vals = [r[field] for r in rows if r[field] and 1<=r[field]<=4]`. If empty → `{sd:0,d:0,a:0,sa:0,avg:None,pct:None,n:0}`. Else `sd=vals.count(1)`, `d=vals.count(2)`, `a=vals.count(3)`, `sa=vals.count(4)`, `avg=round(sum(vals)/len(vals),2)`, `pct=round(sum(vals)/len(vals)/4*100,1)`, `n=len(vals)`. So **Score% = (avg/4)*100** to 1 dp.
- **Inputs.** `feedback_response` rows for `(plant_id, session_code)`.
- **Filters.** `r.plant_id=? AND r.session_code=?`. For central/admin the calendar is fetched by id (any plant); for SPOC the calendar must match `session['plant_id']`. No FY/month/status filter — all feedback rows for that session. Per-question 1–4 validity applies.
- **Edge cases.** `n` is per-question (count of valid answers), not total responses. A question with zero valid answers → `avg`/`pct` are None (rendered "—"). Template colours `pct`: ≥75 green, ≥50 amber, else red. `avg` to 2 dp, `pct` to 1 dp.
- **Source.** `tms/routes/qr.py:404-418`; template `templates/feedback_report.html:93-203`.

### Programme Score subtotal (q1–q5)

- **What it is.** Overall course score = the average of the per-question averages of the first 5 questions, shown out of 4 and as a percentage.
- **Formula.** `_subtotal(q_stats[:5])`: `avgs = [s['avg'] for _,s in slice if s['avg'] is not None]`; if none → `(None,None)`; else `avg = round(sum(avgs)/len(avgs),2)`, `pct = round(avg/4*100,1)`. **Mean of the 5 question-level averages (equal weight per question), NOT a pooled mean over responses.**
- **Inputs.** `q_stats[0..4]`.
- **Filters.** Same per-session, per-plant scope. A question is included only if its avg is not None.
- **Edge cases.** Equal-weight-per-question (not response-weighted). If all 5 avgs are None → None. The 2C-fold version (`course_feedback`) uses a **different** (response-weighted) method.
- **Source.** `tms/routes/qr.py:420-427`.

### Trainer Score subtotal (q6–q9)

- **What it is.** Overall faculty score = the average of the per-question averages of the last 4 questions.
- **Formula.** `_subtotal(q_stats[5:])`: `avg=round(sum(avgs)/len(avgs),2)`, `pct=round(avg/4*100,1)`. Same equal-weight-per-question method.
- **Inputs.** `q_stats[5..8]`.
- **Filters.** Same per-session, per-plant scope; a question included only if its avg is not None.
- **Edge cases.** Equal-weight-per-question. None if all 4 avgs are None. Distinct from `trainer_fb_participants`/`trainer_fb_facilities` (single-question values folded into 2C).
- **Source.** `tms/routes/qr.py:420-428`.

### Overall Satisfaction (de-facto CSAT, all 9 questions)

- **What it is.** The single overall satisfaction index — the average of all 9 per-question averages, out of 4 and as a percentage. This is the headline CSAT figure.
- **Formula.** `_subtotal(q_stats)`: `avg = round(mean of the 9 question-level averages, 2)`; `pct = round(overall_avg/4*100,1)`. Because it averages the 9 question-**averages** (not the 5-q and 4-q subtotals), the overall index is **not generally equal** to `mean(prog_avg, trainer_avg)` when response counts differ.
- **Inputs.** All 9 `q_stats` entries.
- **Filters.** Same per-session, per-plant scope; each question contributes only if its avg is not None.
- **Edge cases.** Equal-weight-per-question across the 9. None if every question avg is None. There is **no field literally named "CSAT"** in code — `overall_pct` is the de-facto CSAT index ("CSAT" appears only in `login.html` marketing copy).
- **Source.** `tms/routes/qr.py:420-429`; template `templates/feedback_report.html:54-61,197-203`.

### Feedback Reports Index average score (per session)

- **What it is.** On the list of sessions with feedback, each session shows one average score, computed in SQL as the average of the sum of all 9 answers ÷ 9, shown as a percentage.
- **Formula.** SQL: `avg_score = AVG( ((CASE WHEN q_obj_explained>0 THEN q_obj_explained END) + ... + (CASE WHEN q_well_involved>0 THEN q_well_involved END)) / 9.0 )`. Template: `pct = (min(avg_score,4)/4*100) if avg_score else None`. `fb_count = COUNT(f.id)`.
- **Inputs.** `calendar c JOIN feedback_response f ON f.session_code=c.session_code AND f.plant_id=c.plant_id`. GROUP BY `c.id`, ORDER BY `c.plan_start DESC`.
- **Filters.** `WHERE c.plant_id=?` (CENTRAL(99) for central; `session['plant_id']` or 99 for admin; `session['plant_id']` for SPOC). INNER JOIN means only sessions with ≥1 feedback row appear. No FY/month/status filter.
- **Edge cases.** **This index formula differs from (and is arguably buggier than) the per-session report.** `CASE WHEN q>0 THEN q END` returns NULL (not 0) on a 0/missing answer, and in SQLite `NULL + number = NULL`, so **any single missing/0 answer makes the whole 9-term row-sum NULL**, dropping that row from `AVG`. Effectively `avg_score` is the response-weighted mean of `(sum of 9)/9` computed **only over responses that answered all 9 questions**; partial responses are excluded entirely. This differs from the report's equal-weight-per-question `overall_avg`.
- **Source.** `tms/routes/qr.py:342-360`; template `templates/feedback_reports_index.html:28,41`.

### Feedback fold into 2C (programme_details) — COALESCE-only

- **What it is.** On QR feedback submit (and again when the SPOC saves the 2C row), session-level feedback averages are written into the conducted-session record — but only into feedback fields the SPOC left blank, never overwriting manual entries.
- **Formula.** Per-question pooled averages via SQL: `q1..q9 = AVG(NULLIF(q_col,0))` over `feedback_response WHERE plant_id=? AND session_code=?` (0 excluded; **response-weighted mean per question**, unlike the report's equal-weight subtotals). Then Python `prog_avg = _avg([q1..q5])`, `trainer_avg = _avg([q6..q9])` (`_avg` drops Nones, `round(sum/len,2)` or None). Then:
  ```sql
  UPDATE programme_details
     SET course_feedback        = COALESCE(course_feedback, prog_avg),
         faculty_feedback       = COALESCE(faculty_feedback, trainer_avg),
         trainer_fb_participants = COALESCE(trainer_fb_participants, q8),
         trainer_fb_facilities   = COALESCE(trainer_fb_facilities, q9)
   WHERE plant_id=? AND session_code=?
  ```
- **Inputs.** `feedback_response` (all q columns) → `programme_details.{course_feedback, faculty_feedback, trainer_fb_participants, trainer_fb_facilities}` (REAL, `schema.sql:125-128`).
- **Filters.** Scope `plant_id=? AND session_code=?`. COALESCE guard: only NULL target columns filled — non-NULL (SPOC-entered) values preserved. `NULLIF(col,0)` excludes 0 answers per question.
- **Edge cases.** COALESCE-only, **no stub**: if no 2C row exists yet, the UPDATE matches 0 rows (deliberate no-op) — responses stay in `feedback_response` and fold later when `add_programme_details` runs this same function post-insert. **Mapping quirk to flag:** `trainer_fb_participants` is fed from **q8 = q_queries_responded** and `trainer_fb_facilities` from **q9 = q_well_involved** — the column **names** (participants/facilities) do **not** match the question semantics; they store single-question response-weighted averages, not the "participants"/"facilities" concepts the names imply. `course_feedback` here uses response-weighted per-question means then averaged, which can differ numerically from the report's `prog_avg` (equal-weight-per-question).
- **Source.** `tms/routes/qr.py:47-88` (`_recompute_feedback_aggregates`), `:23-25` (`_avg`); called at `:836` after feedback submit and from `add_programme_details` on 2C save.

### Anonymous IP dedup (submission integrity)

- **What it is.** A participant submitting feedback without an employee code is blocked from a second anonymous submission from the same IP for that session, preventing one device from spamming anonymous responses.
- **Formula.** If `emp_code` blank: `SELECT 1 FROM feedback_response WHERE plant_id=? AND session_code=? AND emp_code IS NULL AND ip_address=?` — if a row exists, redirect to thanks (silently treated as already submitted, no new row). If `emp_code` present: validate it exists & is_active (corp_members or employees for central; employees for plant), then `INSERT OR REPLACE` keyed on `UNIQUE(plant_id,session_code,emp_code)`.
- **Inputs.** `feedback_response.ip_address` (`request.remote_addr`), `emp_code`.
- **Filters.** Dedup scope: `plant_id + session_code + emp_code IS NULL + ip_address`. Identified (emp_code) submissions dedup via the UNIQUE constraint with `INSERT OR REPLACE` (a re-submit **overwrites** the prior response).
- **Edge cases.** Multiple participants behind one NAT/shared IP submitting anonymously: only the **first** anonymous response per IP per session is recorded — subsequent distinct people on that IP are silently dropped. emp_code submissions are not IP-deduped, only constraint-deduped (same person re-submitting replaces prior answers). Time gate: feedback blocked before `plan_start`. Status lock: feedback blocked unless `calendar.status='To Be Planned'`.
- **Source.** `tms/routes/qr.py:784-841`; `schema.sql:166` UNIQUE.

---

## Effectiveness Review Dates & Eligibility

3-month post-training effectiveness reviews for **Specialized** programmes. Seeding in `tms/routes/verify.py`; status/counts in `tms/routes/effectiveness.py`.

### Eligibility — which sessions/attendees get a review seeded

- **What it is.** Only programmes tagged `'Specialized'` create reviews. When such a session is approved/conducted, one review row is created per distinct attendee.
- **Formula.** **Gate:** seed only if `after_snap` (calendar row) is non-null AND `COALESCE(calendar.category,'General') == 'Specialized'`; else returns `(0, None)`. **Seed:**
  ```sql
  INSERT OR IGNORE INTO effectiveness_review
    (plant_id, session_code, emp_code, conducted_date, due_date)
  SELECT ?, ?, emp_code, ?, ?
  FROM (SELECT DISTINCT emp_code FROM emp_training WHERE plant_id=? AND session_code=?)
  ```
  Returns `(cur.rowcount or 0, due_date)`.
- **Inputs.** `calendar` (category, plan_end, plan_start via `after_snap`); `emp_training` (emp_code) by `plant_id+session_code`; target `effectiveness_review` with `UNIQUE(plant_id, session_code, emp_code)`.
- **Filters.** `category='Specialized'` (default `'General'` coerced); `emp_training WHERE plant_id=? AND session_code=?`. `INSERT OR IGNORE` prevents duplicates via the UNIQUE constraint. Plant scope = the session's own `plant_id` (passed by caller).
- **Edge cases.** Non-Specialized → 0 seeded. Missing calendar row → 0. `INSERT OR IGNORE` is idempotent. DISTINCT collapses duplicate attendance. **Known gap:** seeding fires only at the verify/2C-conduct moment, reading `emp_training` as it stands then — attendees added to 2A **after** approval are **not** retro-seeded. **Known gap:** centrally-hosted Specialized sessions whose attendees sit under a different `plant_id` are not captured by the strict `plant_id=?` filter.
- **Source.** `tms/routes/verify.py:11-39` (gate `:20-23`, conducted_date `:25-26`, due_date `:27-30`, INSERT…SELECT `:31-39`). Call sites: `verify.py:136-138`, `programme.py:640-643`, `programme.py:983`.

### Conducted date (basis date)

- **What it is.** The date the session was held — taken from calendar plan end, then plan start, then today (IST).
- **Formula.** `conducted_date = COALESCE(after_snap.plan_end, after_snap.plan_start, _now_ist().isoformat(timespec='seconds')[:10])`. Uses the **planned** end/start dates, not a separate actual-conducted field.
- **Inputs.** `calendar.plan_end`, `calendar.plan_start`; `_now_ist()` as last resort.
- **Filters.** None beyond the eligibility gate; per-session.
- **Edge cases.** If both plan dates are NULL/empty → today's IST date, so the 90-day clock anchors to the seeding day, not the real session day.
- **Source.** `tms/routes/verify.py:25-26`.

### Due date (review deadline)

- **What it is.** The review is due 90 days after the conducted date.
- **Formula.** `due_date = (date.fromisoformat(conducted_date) + timedelta(days=90)).isoformat()`. On parse failure → `due_date = conducted_date`.
- **Inputs.** `conducted_date`.
- **Filters.** None; per-row.
- **Edge cases.** Non-ISO `conducted_date` → fallback sets `due_date = conducted_date` (due immediately), via `except (ValueError, TypeError)`. Fixed 90-day window (not configurable).
- **Source.** `tms/routes/verify.py:27-30`.

### Overdue cutoff (grace window)

- **What it is.** A review becomes "overdue" once more than 30 days past its due date; between due and +30 days it shows as "due".
- **Formula.** `OVERDUE_DAYS = 30`. `overdue_cutoff = (date.fromisoformat(due_date) + timedelta(days=30)).isoformat()`. Overdue when `today_iso > overdue_cutoff`.
- **Inputs.** `due_date`; `OVERDUE_DAYS`.
- **Filters.** None.
- **Edge cases.** Un-parseable `due_date` → exception swallowed, status falls through to `'due'` (never `'overdue'`). 30 days hardcoded module-level (not per-plant config).
- **Source.** `tms/routes/effectiveness.py:25` (OVERDUE_DAYS); `:36-43` (`_eff_status`).

### Status derivation (pending / due / overdue / completed)

- **What it is.** Derived at read time (never stored). Filed → completed; before due → pending; due..+30 days → due; beyond → overdue. "Today" is the IST date.
- **Formula.** `_eff_status(today_iso, due_date, completed_date)`: `if completed_date → 'completed'; elif not due_date → 'pending'; elif today_iso < due_date → 'pending'; elif today_iso > (due_date + 30d) → 'overdue'; else → 'due'`. `today_iso = _today_ist().isoformat()`. String comparison on ISO `'YYYY-MM-DD'` (lexicographic == chronological).
- **Inputs.** `effectiveness_review.completed_date`, `.due_date`; `_today_ist()`.
- **Filters.** In the route, status is computed per row then optionally filtered to `sel_status` (request arg) in Python. Plant scope for SPOC: `WHERE e.plant_id=?`; central/admin see all plants.
- **Edge cases.** `completed_date` takes precedence. Boundaries: `today == due_date` → `'due'` (uses `<`); `today == overdue_cutoff` → `'due'` (uses strict `>`). Missing `due_date` → `'pending'`. IST 'today' aligns boundaries to India clock.
- **Source.** `tms/routes/effectiveness.py:28-43` (`_eff_status`); today derivation `:49`, `:77`; per-row apply `:108-114`.

### Counts (pending / due / overdue / completed / total / open)

- **What it is.** Tallies of reviews by status; `open` = pending+due+overdue; `total` = all reviews in scope.
- **Formula.** Fetch `SELECT completed_date, due_date FROM effectiveness_review [WHERE plant_id=?]`; per row increment `counts[_eff_status(today, due_date, completed_date)]`. `counts={'pending','due','overdue','completed'}` (start 0). `total = sum of the four`. `open = pending + due + overdue`.
- **Inputs.** `effectiveness_review` (completed_date, due_date); `_today_ist()`.
- **Filters.** `plant_id` scope: passed plant_id → `WHERE plant_id=?`; None → all plants (central/admin). SPOC passes session plant_id; central/admin pass None. **No FY/month/status filter** — counts span all years' reviews in scope.
- **Edge cases.** Recomputed live each request (no caching). Same IST 'today' and 30-day rule. No dedup needed (rows unique per plant+session+emp). `total` deliberately = the 4 buckets only; `open` added after the sum.
- **Source.** `tms/routes/effectiveness.py:46-63` (`_eff_counts`; total `:61`, open `:62`); API `:176-183`; route usage `:116`.

### Filing a review (completion / Kirkpatrick L2-L3 capture)

- **What it is.** The SPOC records the manager's review (1–5 rating + observations); filing stamps the completion date so the review counts as completed.
- **Formula.** On valid submit: `UPDATE effectiveness_review SET completed_date=now_iso, rating=?, behaviour_change=?, application_on_job=?, comments=?, filed_by=user_id, filed_at=now_iso WHERE id=?`. `now_iso = _now_ist().isoformat(timespec='seconds')`. Validations: `rating` int must be 1–5 (else reject); text fields trimmed to `[:1000]`; require `len(behaviour)>=10 OR len(application)>=10` (≥1 observation, min 10 chars).
- **Inputs.** Form `rating, behaviour_change, application_on_job, comments`; `effectiveness_review` row by id.
- **Filters.** Plant scope: `role in ('spoc','admin') AND eff.plant_id != session plant_id` → rejected ("Switch plant to file"). `@spoc_required`.
- **Edge cases.** Non-numeric rating → coerced to 0 then rejected (<1). DB CHECK allows NULL or 1–5. `completed_date` once set makes status `'completed'` permanently (no un-file path). Audit logged via `log_record_change('EFFECTIVENESS_FILE', ...)`. Cross-plant filing blocked for spoc/admin.
- **Source.** `tms/routes/effectiveness.py:122-174` (validation `:142-157`; UPDATE `:161-167`).

---

## Central Rollup — Cross-plant Dashboard `/central` and Drill-in `/central/plant`

`central_dashboard()` (`tms/routes/central.py:60`) runs 7 batched queries + a grand-central scalar over the 10 plants in `PLANTS`. FY bounds from `_current_fy()` (`central.py:62`).

### Per-plant headcount — BC (bc) and WC (wc)

- **Formula.** `SELECT plant_id, SUM(CASE WHEN collar='Blue Collared' THEN 1 ELSE 0 END) bc, SUM(CASE WHEN collar='White Collared' THEN 1 ELSE 0 END) wc FROM employees ... GROUP BY plant_id`. Per plant `bc=hc.get(pid).get('bc',0)`, `wc=...('wc',0)`, `total_emp=bc+wc`.
- **Inputs.** `employees` (collar, plant_id, is_active).
- **Filters.** `WHERE is_active=1`. No FY filter (point-in-time). Any collar other than the two strings counted in neither.
- **Edge cases.** Missing plant → 0. NULL/other collar silently excluded (not in `total_emp`).
- **Source.** `tms/routes/central.py:65-70, 105-106, 118`.

### Per-plant OWN planned sessions (own_sessions)

- **Formula.** `SELECT plant_id, COUNT(*) cnt FROM calendar WHERE plan_start BETWEEN ? AND ? GROUP BY plant_id`. `own_s = own_planned.get(pid,0)`.
- **Inputs.** `calendar` (plant_id, plan_start).
- **Filters.** `plan_start BETWEEN fy_start AND fy_end`. **No status filter** — all calendar rows in the FY window.
- **Edge cases.** Counts every calendar row in the window (any status). NULL/out-of-FY plan_start excluded. Missing plant → 0.
- **Source.** `tms/routes/central.py:71-74, 107`.

### Per-plant OWN conducted sessions (own_conducted)

- **Formula.** `SELECT plant_id, COUNT(*) cnt FROM calendar WHERE status='Conducted' AND plan_start BETWEEN ? AND ? GROUP BY plant_id`. `own_c = own_cond.get(pid,0)`.
- **Inputs.** `calendar` (plant_id, status, plan_start).
- **Filters.** `status='Conducted' AND plan_start BETWEEN fy_start AND fy_end`.
- **Edge cases.** Only exact `'Conducted'` qualifies (`'Awaiting Verification'` excluded). Missing plant → 0.
- **Source.** `tms/routes/central.py:75-78, 108`.

### Per-plant central-hosted sessions attended (cen_att / cen_c)

- **Formula.** `SELECT plant_id, COUNT(DISTINCT session_code) cnt FROM emp_training WHERE host_plant_id=99 AND session_code IS NOT NULL AND session_code!='' AND start_date>=? AND start_date<=? GROUP BY plant_id`. `cen_c = cen_att.get(pid,0)`. Then plant `sessions = own_s + cen_c` and `conducted = own_c + cen_c`.
- **Inputs.** `emp_training` (plant_id, host_plant_id, session_code, start_date).
- **Filters.** `host_plant_id=99`; non-empty `session_code`; `start_date BETWEEN fy_start AND fy_end`. `plant_id` = attendee's home plant.
- **Edge cases.** DISTINCT session_code de-dupes multiple attendees of one central session into one count per plant. Blank/NULL session_code excluded. A central session is counted once **per attending plant** — the same physical session can be counted under several plants. Missing plant → 0.
- **Source.** `tms/routes/central.py:79-83, 109, 119`.

### Per-plant total man-hours (manhours)

- **Formula.** `SELECT plant_id, COALESCE(SUM(hrs),0) cnt FROM emp_training WHERE start_date>=? AND start_date<=? GROUP BY plant_id`. `manhours = mh_all.get(pid,0)`, displayed `round(manhours,1)`.
- **Inputs.** `emp_training` (plant_id, hrs, start_date).
- **Filters.** `start_date BETWEEN fy_start AND fy_end`. **No join to employees** — deliberately un-joined so it also counts hours for orphan attendance rows.
- **Edge cases.** COALESCE → 0. Includes hours regardless of collar/active status. Can exceed `bc_hrs+wc_hrs` because orphan/uncollared attendees contribute here but not the collar split. Round 1 dp at display.
- **Source.** `tms/routes/central.py:84-87, 110, 121`.

### Per-plant man-hours split by collar (bc_hrs / wc_hrs)

- **Formula.** `SELECT t.plant_id, COALESCE(SUM(CASE WHEN e.collar='Blue Collared' THEN t.hrs END),0) bc, COALESCE(SUM(CASE WHEN e.collar='White Collared' THEN t.hrs END),0) wc FROM emp_training t JOIN employees e ON e.emp_code=t.emp_code AND e.plant_id=t.plant_id WHERE t.start_date>=? AND t.start_date<=? GROUP BY t.plant_id`. `mh_bc[pid]=r['bc']`, `mh_wc[pid]=r['wc']`.
- **Inputs.** `emp_training` JOIN `employees`.
- **Filters.** `start_date BETWEEN fy_start AND fy_end`. INNER JOIN on emp_code+plant_id (orphans excluded from the collar split). **Join does not filter `is_active`** — exited employees' hours still flow in (exit-policy numerator).
- **Edge cases.** Inner JOIN drops orphan attendance (still in `mh_all`). Other-collar hours dropped from both. COALESCE → 0. Missing plant → 0.
- **Source.** `tms/routes/central.py:91-100, 111-112`.

### Per-plant BC% / WC% mandate fulfilment (bc_pct / wc_pct)

- **Formula.** `bc_mandate = bc * get_config('mh_target_bc',12,plant_id=pid)`; `wc_mandate = wc * get_config('mh_target_wc',24,plant_id=pid)`; `bc_pct = round(bc_hrs/bc_mandate*100,1) if bc_mandate else 0`; `wc_pct = round(wc_hrs/wc_mandate*100,1) if wc_mandate else 0`. `target_hrs = bc_mandate + wc_mandate`.
- **Inputs.** bc/wc active headcount; bc_hrs/wc_hrs (FY collar man-hours); per-plant config targets.
- **Filters.** Headcount `is_active=1`; man-hours FY-bound + INNER JOIN employees (no is_active on hours). Targets per-plant.
- **Edge cases.** Zero-division guarded (0 when mandate 0). **Exit-policy asymmetry** — denominator `is_active=1`, numerator includes exited hours → pct **can exceed 100%** by design. Round 1 dp. Target is per-plant config, not a constant.
- **Source.** `tms/routes/central.py:113-116, 123, 126`.

### GRAND TOTAL — central sessions (grand_central)

- **Formula.** `SELECT COUNT(DISTINCT session_code) FROM emp_training WHERE host_plant_id=99 AND session_code IS NOT NULL AND session_code!='' AND start_date>=? AND start_date<=?`. Single scalar.
- **Inputs.** `emp_training` (host_plant_id, session_code, start_date).
- **Filters.** `host_plant_id=99`; non-empty session_code; FY window. No GROUP BY — global DISTINCT across all plants.
- **Edge cases.** DISTINCT de-dupes a central session attended by multiple plants into **one** — so `grand_central` is **not** the sum of per-plant `cen_att` (a session attended by 3 plants counts 3 in `sum(cen_att)` but 1 here). Intentional, to avoid double-counting central sessions fleet-wide.
- **Source.** `tms/routes/central.py:128-132`.

### GRAND TOTAL — fleet headcount (total_emp, blue_collar, white_collar)

- **Formula.** `grand_bc = sum(p['blue_collar'] for p in plant_summaries)`; `grand_wc = sum(p['white_collar'] ...)`; `grand_total = grand_bc + grand_wc`.
- **Inputs.** Per-plant bc/wc headcount (already `is_active=1`).
- **Filters.** Inherits `is_active=1`. Iterates only `PLANTS` (ids 1–10); Central plant 99 is not in `PLANTS` so not summed.
- **Edge cases.** Only the 10 `PLANTS` are aggregated; any plant_id in DB but absent from `PLANTS` excluded. Uncollared employees excluded.
- **Source.** `tms/routes/central.py:133-135, 138-141`; `constants.py:17-28`.

### GRAND TOTAL — fleet sessions and conducted

- **Formula.** `grand['sessions'] = sum(p['own_sessions']) + grand_central`; `grand['conducted'] = sum(p['own_conducted']) + grand_central`.
- **Inputs.** Per-plant `own_sessions` + `own_conducted` + `grand_central`.
- **Filters.** Sessions: calendar `plan_start` in FY (all statuses). Conducted: calendar `status='Conducted'` + FY. grand_central: `host_plant_id=99` + FY + distinct session_code.
- **Edge cases.** Central sessions added via `grand_central` (counted once fleet-wide), **not** via per-plant `cen_c` — so `grand['sessions'] != sum(plant['sessions'])`. The same `grand_central` value is added to both sessions and conducted (a central session is treated as both planned and conducted at fleet level).
- **Source.** `tms/routes/central.py:143-144`.

### GRAND TOTAL — fleet man-hours, target, hrs_pct, avg_hrs_per_emp

- **Formula.** `grand_manhours = round(sum(p['manhours']),1)`; `grand_target = sum(p['target_hrs'])`; `grand['hrs_pct'] = round(grand_manhours/grand_target*100,0) if grand_target else 0`; `grand['avg_hrs_per_emp'] = round(grand_manhours/grand_total,1) if grand_total else 0`.
- **Inputs.** Per-plant `manhours` (mh_all, total FY hrs incl. orphans), per-plant `target_hrs`, `grand_total` headcount.
- **Filters.** Man-hours: `emp_training.start_date` in FY, **all attendees** (un-joined — orphans + exited + all collars). Target: per-plant active headcount × per-plant config targets.
- **Edge cases.** **Numerator/denominator mismatch:** `grand_manhours` (numerator) uses `mh_all` = ALL FY hours (orphans, exited, uncollared all in), while `grand_target` is built only from active+collared headcount. Combined with the exit policy, `hrs_pct` can exceed 100%. `hrs_pct` rounded to 0 dp (integer-like); `avg_hrs_per_emp` to 1 dp. Both zero-division guarded.
- **Source.** `tms/routes/central.py:136-137, 145-147`.

### Plant-summary table sort order

- **Formula.** `plant_summaries.sort(key=lambda p: (p['bc_pct'] + p['wc_pct'])/2, reverse=True)`.
- **Inputs.** Per-plant `bc_pct`, `wc_pct`.
- **Filters.** None additional.
- **Edge cases.** Simple arithmetic mean of the two percentages (not headcount-weighted). A plant with one collar at 0 employees (pct=0) is penalised in the average even though that collar has no mandate.
- **Source.** `tms/routes/central.py:126`.

### Current FY date bounds used by all `/central` queries

- **Formula.** `_current_fy()`: `today=_today_ist(); yr = today.year if today.month>=4 else today.year-1; return (f'{yr}-04-01', f'{yr+1}-03-31')`.
- **Inputs.** System clock via `_now_ist()`/`_today_ist()` (Asia/Kolkata).
- **Filters.** Applied as `start_date`/`plan_start BETWEEN fy_start AND fy_end` across the 7 batched queries and grand_central.
- **Edge cases.** IST wall clock — FY rolls at IST midnight; before April the window belongs to the prior calendar year. String date comparison on `'YYYY-MM-DD'` (BETWEEN inclusive of both endpoints).
- **Source.** `tms/helpers.py:68-73`; `:129-143`; called at `tms/routes/central.py:62`.

### Per-plant drill-in `/central/plant` — read-only reuse of plant engines

- **What it is.** Clicking a plant shows that plant's own Monthly Summary, totals, and compliance using exactly the same engine the SPOC sees.
- **Formula.** `summary_rows=_calc_summary(plant_id, sel_month, db)`; `totals=_calc_totals(summary_rows, db=db, plant_id=plant_id)`; `compliance=_calc_compliance(plant_id, db)`; `mh_target_bc=get_config('mh_target_bc',12,plant_id)`; `mh_target_wc=get_config('mh_target_wc',24,plant_id)`. `sel_month` from `?month=`.
- **Inputs.** `programme_details`, `calendar`, `emp_training`, `employees`, `tni`, `org_config` — via the three shared helpers.
- **Filters.** Plant scope = path `plant_id` (validated against `PLANT_MAP`, else redirect). Optional month filter. Inherits all helper filters (Summary FY-binds man-hours/seats and counts conducted-only; coverage uses source='TNI Driven' + fy_year; compliance uses is_active=1 denominator and FY-bound exit-policy numerator).
- **Edge cases.** `plant_id` not in `PLANT_MAP` → flash + redirect. Central pseudo-plant 99 has no `PLANTS` entry but **is** in `PLANT_MAP`, so `/central/plant/99` renders Central as a "plant". Month not in `MONTH_NUM` → empty result (`AND 1=0`). Coverage/compliance zero-division guarded → 0. These are the plant-level engines invoked unchanged for Central's read-only view.
- **Source.** `tms/routes/central.py:758-777`; helpers `_calc_summary:424-575`, `_calc_totals:578-625`, `_calc_compliance:628-676`.

---

## Calendar Coverage / Demand & Session Actuals

Calendar planning coverage (`training_calendar`, `tms/routes/calendar.py`), the live in-form coverage panel (`api_tni_coverage`, `tms/routes/api.py`), and the host-aware actuals recompute (`_recompute_session_actuals`, `tms/helpers.py:95`).

### Planning Demand (per programme)

- **What it is.** Per programme, the number of distinct employees at this plant nominated for it in TNI — the denominator the planner must cover.
- **Formula.** `demand = COUNT(DISTINCT emp_code)` grouped by `programme_name`: `SELECT programme_name, COUNT(DISTINCT emp_code) cnt FROM tni WHERE plant_id=? GROUP BY programme_name`.
- **Inputs.** `tni` (plant_id, emp_code, programme_name).
- **Filters.** `plant_id = session['plant_id']` ONLY. **No FY/status/collar/source filter** — every TNI nomination counts regardless of when raised. DISTINCT on emp_code dedups.
- **Edge cases.** Programmes in TNI but never scheduled still appear in `cov_rows`. A session whose programme has no TNI rows has demand=0 and is **not** added to `cov_rows` (the loop iterates `demand_map.items()`). Per-row display uses "—" when no TNI entry.
- **Source.** `tms/routes/calendar.py:61-63` (build), `:81-90` (cov_rows), `training_calendar`.

### Planned Pax (seats planned)

- **What it is.** Total seats scheduled for a programme — the sum of planned headcount across its calendar sessions in the current view.
- **Formula.** `planned_pax = SUM(calendar.planned_pax)` over the programme's sessions; accumulated `pax_map[p]['planned_pax'] += (s['planned_pax'] or 0)`.
- **Inputs.** `calendar.planned_pax`, iterated from the fetched `sessions` rowset.
- **Filters.** `plant_id = session['plant_id']`. In the **default** view the sessions query **excludes** `status='Lapsed'`, so Lapsed sessions' planned_pax don't count; `?include_lapsed=1` includes them. **No FY filter** on the sessions rowset. No `status='Conducted'` requirement (To Be Planned / Awaiting Verification / Conducted all included).
- **Edge cases.** NULL planned_pax → 0. Counts both TNI Driven and New Requirement sessions (source not filtered).
- **Source.** `tms/routes/calendar.py:73-83`, `training_calendar`.

### Conducted Pax (per programme)

- **What it is.** Of the planned seats, how many belong to sessions marked Conducted.
- **Formula.** `if s['status'] == 'Conducted': pax_map[p]['conducted_pax'] += (s['planned_pax'] or 0)`.
- **Inputs.** `calendar` (status, planned_pax) from the `sessions` rowset.
- **Filters.** plant_id scope + (default view) `status != 'Lapsed'`, PLUS `status == 'Conducted'` in Python. **Sums the session's PLANNED pax, NOT actual attendance** — it is planned seats of conducted sessions, not seats actually filled.
- **Edge cases.** NULL planned_pax → 0. Uses planned_pax even for Conducted sessions, so a session conducted with fewer/more actual attendees still contributes its planned figure.
- **Source.** `tms/routes/calendar.py:79-80, 88`, `training_calendar`.

### Coverage Gap (planning)

- **What it is.** How many more seats still need to be planned to meet TNI demand. Zero means fully (or over-) covered.
- **Formula.** `gap = max(0, demand - planned_pax)`.
- **Inputs.** `demand_map` (tni) and `pax_map` (calendar.planned_pax).
- **Filters.** Inherits demand filter (plant_id only, no FY) and planned_pax filter (plant_id, non-Lapsed in default view). Floored at 0.
- **Edge cases.** When `planned_pax > demand`, gap=0 and the surplus shows via the separate `over` field. `cov_rows` sorted by gap descending (largest unmet need first).
- **Source.** `tms/routes/calendar.py:84`, `training_calendar`.

### Over-plan (surplus seats)

- **What it is.** Seats planned beyond the TNI demand (over-scheduling).
- **Formula.** `over = max(0, planned_pax - demand)`.
- **Inputs.** `pax_map` and `demand_map`.
- **Filters.** Same scope as planned_pax and demand. Floored at 0.
- **Edge cases.** Mutually exclusive with gap (both 0 only when exactly equal).
- **Source.** `tms/routes/calendar.py:90`, `training_calendar`.

### Coverage % (planning, per programme)

- **What it is.** What fraction of a programme's TNI demand is covered by planned seats. 100% means enough seats scheduled to seat every nominee.
- **Formula.** `pct = min(100, round(planned_pax / demand * 100)) if demand > 0 else 0`.
- **Inputs.** `pax_map` (planned_pax, numerator) and `demand_map` (distinct TNI nominees, denominator).
- **Filters.** Denominator = ALL plant TNI distinct nominees (no FY filter). Numerator = summed planned_pax of non-Lapsed sessions (default view). Capped at 100. Division guarded (demand==0 → 0).
- **Edge cases.** **Distinct from the canonical Monthly Summary coverage:** this PLANNING coverage uses planned **SEATS** ÷ distinct TNI nominees, per programme; the Summary coverage (`_calc_summary`) uses distinct employees actually **TRAINED** ÷ distinct TNI nominees, grouped by prog_type+collar, FY-bound. They will diverge — planning coverage counts seats scheduled (can double-count an employee across sessions / count over-plan toward the 100% cap), Summary coverage counts unique people trained.
- **Source.** `tms/routes/calendar.py:85`, `training_calendar`.

### Live TNI Coverage Panel (in-form, per typed programme)

- **What it is.** While creating a session, shows for the chosen programme: demand, sessions already planned this FY, distinct employees already trained this FY, remaining, and % already trained.
- **Formula.** `demand = COUNT(DISTINCT emp_code) FROM tni WHERE plant_id=? AND programme_name=canonical`. `sessions_planned = COUNT(*) FROM calendar WHERE plant_id=? AND programme_name=canonical AND session_code LIKE '%/{fy}/%'`. `covered = COUNT(DISTINCT emp_code) FROM emp_training WHERE plant_id=? AND programme_name=canonical AND start_date BETWEEN fy_start AND fy_end`. `uncovered = max(0, demand - covered)`. `pct = round(covered / demand * 100) if demand>0 else 0` (**NOT capped at 100** here, unlike the table).
- **Inputs.** `tni`, `calendar`, `emp_training`. Programme name fuzzy-canonicalised against tni distinct names (`difflib get_close_matches`, cutoff 0.65) when no exact match.
- **Filters.** `plant_id = session['plant_id']` on all three queries. demand: **no FY filter**. sessions_planned: current-FY only via `session_code LIKE '%/{fy}/%'` (`fy=_fy_label()`). covered: current-FY only via `start_date BETWEEN _current_fy()`. `covered` counts ALL distinct attendees in FY (TNI and non-TNI alike — `emp_training` is not joined to `tni`), so `covered` can exceed `demand` and `pct` can exceed 100 here.
- **Edge cases.** Returns `{}` when q empty or demand falsy (panel hidden). `pct` is **not** `min(100,...)`-capped (differs from the cov_rows table). `covered` uses distinct employees TRAINED — a different metric than the table's planned_pax. `source` returned as `'TNI Driven'` iff `demand>0` else `''`. Auto-fills form pax = uncovered.
- **Source.** `tms/routes/api.py:96-172`, `api_tni_coverage`.

### Session Actuals — actual_pax (attendance count)

- **What it is.** Real number of attendees recorded for a session (one per attendance row), kept in sync with the attendance sheet. Recomputed on every 2A save/edit/delete and QR confirm.
- **Formula.** `actual_pax = COUNT(*) FROM emp_training WHERE session_code=? AND (plant_id=? OR host_plant_id=?)`. Written via `UPDATE calendar SET actual_pax=?, actual_hrs=? WHERE session_code=? AND plant_id=?`.
- **Inputs.** `emp_training` (session_code, plant_id, host_plant_id, hrs) → `calendar` (actual_pax, actual_hrs).
- **Filters.** Match key: `session_code AND (plant_id = target OR host_plant_id = target)`. **Host-aware:** for a Central (plant 99) session, attendees are stored under their HOME plant_id with `host_plant_id=99`, so matching plant_id alone would miss them; the OR captures them. For ordinary plant sessions `host_plant_id` is NULL so the OR is a no-op. Counts ALL emp_training rows — **no is_active/TNI/collar/FY filter**.
- **Edge cases.** `COUNT(*)` counts every attendance row including non-TNI/non-active attendees (no dedup on emp_code — duplicate rows would double-count). Idempotent: re-running recomputes from scratch. No-op when `session_code` is falsy (early return). The UPDATE targets `WHERE session_code=? AND plant_id=?`, so for plant-99 sessions the actuals land on the **central** calendar row.
- **Source.** `tms/helpers.py:95-111`, `_recompute_session_actuals`.

### Session Actuals — actual_hrs (delivered man-hours)

- **What it is.** Total man-hours delivered for a session = sum of per-person hours across attendees.
- **Formula.** `actual_hrs = COALESCE(SUM(hrs),0) FROM emp_training WHERE session_code=? AND (plant_id=? OR host_plant_id=?)`. Persisted in the same UPDATE as actual_pax.
- **Inputs.** `emp_training.hrs` → `calendar.actual_hrs`.
- **Filters.** Identical host-aware match key to actual_pax. No is_active/TNI/collar/FY filter.
- **Edge cases.** NULL hrs handled by `COALESCE(SUM(hrs),0)` → 0. Sums raw per-person hrs as entered in 2A (no recomputation from the time window). Same idempotent UPDATE; rerun overwrites prior value.
- **Source.** `tms/helpers.py:95-111`, `_recompute_session_actuals`.

---

## Session Code, Audience Derivation, FY/Month Helpers

### Session code generation

- **What it is.** Each planned session gets a unique code like `BCM/TEC/001/26-27/B01`: unit code / prog-type abbrev / 3-digit programme serial / FY / 2-digit batch.
- **Formula.** `session_code = prog_code + '/' + fy + '/B' + nn`, where `prog_code = UNIT_CODE/TYPE_ABBREV/NNN` (NNN = 3-digit zero-padded serial via `_get_or_create_prog_code`), `fy = _fy_label()`, `nn` = 2-digit batch = `MAX(existing batch suffix)+1+attempt`. Batch: `prefix = prog_code+'/'+fy+'/B'`; `nxt = (MAX(CAST(SUBSTR(session_code, len(prefix)+1) AS INTEGER)) over matching rows, or 0) + 1 + attempt`; `code = f'{prefix}{nxt:02d}'`. Retries up to 5 attempts pre-checking calendar; `UNIQUE(calendar.session_code)` is the final authority; last-resort fallback appends `'X'+utcnow microseconds`. Programme serial (`_get_or_create_prog_code`): reuse existing `prog_code` if `(plant_id, programme_name)` already in calendar; else `nxt = (MAX(...)+1)`, return `f'{unit_code}/{abbrev}/{nxt:03d}'`.
- **Inputs.** `calendar` (session_code, prog_code, plant_id); `constants.PLANT_MAP[plant_id]['unit_code']` (e.g. BCM, plant 99 = CEN); `constants.TYPE_ABBREV[prog_type]` (Behavioural/Leadership=BEH, Cane=CAN, Commercial=COM, EHS/HR=EHS, IT=ITC, Technical=TEC; default `'GEN'`); `_fy_label()`.
- **Filters.** Batch MAX lookup filtered by `plant_id=? AND prog_code=? AND session_code LIKE prefix%`. Serial MAX lookup by `plant_id=? AND prog_code LIKE prefix%`. Reuse lookup by `plant_id=? AND programme_name=?`. FY segment is the current FY only. No status/source/is_active filter. Central sessions pass `plant_id=99` explicitly.
- **Edge cases.** Zero-padding: serial `%03d`, batch `%02d`. Empty MAX coerced via `(row['mx'] or 0)`. Concurrency: MAX+1 is racy, so the 5-attempt retry loop bumps `nxt` and pre-checks; the UNIQUE constraint is the real guard; on persistent clash a microsecond-suffixed code (`datetime.utcnow().strftime('%f')`, NOT IST) is returned to avoid a crash. `prog_type` not in `TYPE_ABBREV` → `'GEN'`. Programme serial is reused for repeat batches of an existing programme name so batch numbers increment under one prog_code.
- **Source.** `tms/helpers.py:301` (`_new_session_code`), `:278` (`_get_or_create_prog_code`); `constants.py:18-30` (PLANT_MAP/unit_code), `:82-85` (TYPE_ABBREV); call sites `tms/routes/calendar.py:149-150, 497-498`, `tms/routes/central_training.py:188-189`.

### Audience derivation (`_derive_audience`)

- **What it is.** A session's audience is decided by who is nominated for that programme in TNI (current FY): both collars → Common; only blue → Blue Collared; only white → White Collared; nobody in TNI → None (form value used).
- **Formula.** `collar_set = DISTINCT e.collar` from `tni JOIN employees` for this programme in the current FY. If empty → `None` (audience NOT derived; caller keeps the form value). `elif 'Blue Collared' in set AND 'White Collared' in set → 'Common'`. `elif 'Blue Collared' in set → 'Blue Collared'`. `else → 'White Collared'`.
- **Inputs.** `tni` (plant_id, programme_name, fy_year) INNER JOIN `employees` (emp_code, plant_id, collar). `SELECT DISTINCT e.collar`.
- **Filters.** `t.plant_id = plant_id`; `LOWER(t.programme_name) = LOWER(prog_name)` (case-insensitive); `t.fy_year = _fy_label()` (current FY only); `e.collar IS NOT NULL AND e.collar != ''`. INNER JOIN matches only existing employees, but there is **no `is_active` filter** here and **no source filter**.
- **Edge cases.** Returns `None` (not a string) when no qualifying TNI collars exist — calendar write paths then leave the form-supplied `target_audience` untouched (New Requirement free audience). Blank/NULL collars excluded. Case-insensitive programme match. Only the current FY's TNI rows considered, so a programme nominated only in a prior FY derives no audience. Collar comparison is exact; any other collar value falls through to the final `'White Collared'` branch if present alone.
- **Source.** `tms/helpers.py:241` (`_derive_audience`).

### Financial-year short label (`_fy_label`)

- **What it is.** Current FY as `'YY-YY'` (e.g. `'26-27'`); FY is Apr–Mar, so Jan–Mar belongs to the prior calendar year's FY.
- **Formula.** `today=_today_ist(); y=today.year`. If `today.month < 4`: `f'{str(y-1)[2:]}-{str(y)[2:]}'`; else `f'{str(y)[2:]}-{str(y+1)[2:]}'`.
- **Inputs.** `_today_ist()` only.
- **Filters.** None — pure date computation. FY boundary is April 1.
- **Edge cases.** IST date → FY rolls at IST midnight (not UTC). Two-digit slicing `str(y)[2:]` assumes 4-digit years. Sibling `_fy_label_long()` returns the long en-dash form (e.g. `'2026–27'`) for UI display only.
- **Source.** `tms/helpers.py:260` (`_fy_label`); `:268` (`_fy_label_long`).

### Current FY date bounds (`_current_fy`)

- **What it is.** Start and end dates of the current FY (April 1 to March 31).
- **Formula.** `today=_today_ist(); yr = today.year if today.month >= 4 else today.year - 1; return (f'{yr}-04-01', f'{yr+1}-03-31')`.
- **Inputs.** `_today_ist()` only.
- **Filters.** None — pure computation. Start = April 1 of start year; End = March 31 of next year.
- **Edge cases.** IST-based to avoid rolling 5.5h early on UTC servers. Inclusive `'YYYY-MM-DD'` bounds (used as `BETWEEN start AND end`). `_in_current_fy(date_str)` treats empty/None as in-FY and parses `date_str[:10]`; bad dates → False. `_tni_is_locked()` compares `_today_ist() > fy_end`.
- **Source.** `tms/helpers.py:68` (`_current_fy`); `:76` (`_in_current_fy`); `:88` (`_tni_is_locked`).

### Date-to-month-name (`_date_to_month`)

- **What it is.** Converts `2026-06-07` into `'June'`; blank/unparseable → `''`.
- **Formula.** `d = datetime.strptime(date_str[:10], '%Y-%m-%d'); return d.strftime('%B')`. Empty/None → `''`; any parse exception → `''`.
- **Inputs.** A single date string (first 10 chars).
- **Filters.** None.
- **Edge cases.** Falsy input → `''` immediately. Only the first 10 chars parsed (full ISO timestamp tolerated). Any `strptime` exception swallowed → `''`. `%B` = locale's full month name. No timezone conversion — reflects the literal stored date.
- **Source.** `tms/helpers.py:114` (`_date_to_month`).

### IST wall-clock helpers (`_now_ist` / `_today_ist`)

- **What it is.** Current date/time in IST so the UTC server does not display or compute values 5.5h off.
- **Formula.** `_now_ist()`: `datetime.now(ZoneInfo('Asia/Kolkata')).replace(tzinfo=None)` — timezone-NAIVE datetime carrying IST wall-clock. Fallback if zoneinfo unavailable: `datetime.utcnow() + timedelta(hours=5, minutes=30)`. `_today_ist()`: `_now_ist().date()`.
- **Inputs.** System clock (UTC on Render) converted to Asia/Kolkata.
- **Filters.** None.
- **Edge cases.** Returns a NAIVE datetime (tzinfo stripped) — a drop-in for `datetime.now()`/`date.today()` but stable on UTC hosts. Fallback uses a fixed +5:30 offset (India has no DST). Single source of truth for `_fy_label`, `_current_fy`, `_tni_is_locked`, and session-start gates. The session-code last-resort microsecond fallback intentionally uses `datetime.utcnow()` (uniqueness token only, not a user-facing time).
- **Source.** `tms/helpers.py:129` (`_now_ist`); `:141` (`_today_ist`).

---

## Anomaly Detection & Validation (2C save + time/duration cross-check)

Validations and soft anomalies applied on 2C `programme_details` save (`tms/routes/programme.py`) and the cross-form time/duration check (`tms/helpers.py`).

### Hours mismatch anomaly (2C hours vs average 2A hours)

- **What it is.** After the SPOC enters the official session duration (2C Actual Hours), it is compared to the average per-person hours from attendance (2A). >25% off → flagged for verification (the row still saves).
- **Formula.** `avg_2a = AVG(emp_training.hrs)` over rows for the session where `hrs>0`. Flag when `avg_2a is not NULL AND avg_2a > 0 AND abs(hours - avg_2a) / avg_2a > 0.25` (relative difference strictly > 25%). `hours = float(form 'hours_actual')`. Anomaly string `f'hours_mismatch(2C={hours} vs avg2A={avg_2a:.1f})'`.
- **Inputs.** `emp_training.hrs`, `.session_code`, `.plant_id`, `.host_plant_id`; form `hours_actual`.
- **Filters.** AVG restricted to `emp_training WHERE session_code=? AND hrs > 0 AND (plant_id=? OR host_plant_id=?)` — **host-aware** so centrally-hosted attendees under their own plant_id are still included. `plant_id` bound to `session['plant_id']`. No FY filter (keyed by session_code). Threshold non-inclusive `> 0.25`.
- **Edge cases.** Zero-division guarded (requires `avg_2a` truthy AND `> 0` before dividing). If no 2A rows with `hrs>0` (`avg_2a` NULL or 0), the check is skipped (no flag). **Soft/warn only — never blocks the save.** The bulk path uses a plant_id-only AVG (no host_plant_id widening).
- **Source.** `tms/routes/programme.py:512-520` (single add); bulk re-check `:920-921`.

### Low attendance anomaly (attended vs planned)

- **What it is.** If fewer than half the planned participants attended, the session is flagged.
- **Formula.** Flag when `planned_pax > 0 AND attended < planned_pax * 0.5` (strictly < 50%). Anomaly string `f'low_attendance({attended}/{planned_pax})'`. `planned_pax = calendar.planned_pax` (0 if NULL); `attended` = count of attendance rows.
- **Inputs.** `calendar.planned_pax`; `attended = COUNT(*) FROM emp_training WHERE plant_id=? AND session_code=?` (single-add).
- **Filters.** Calendar row matched by `session_code AND plant_id`. `attended` counted over `emp_training WHERE plant_id=? AND session_code=?` (plant-scoped, **not host-aware** here, so centrally-hosted attendees under other plants are not counted in this denominator). Only evaluated when `planned_pax > 0`.
- **Edge cases.** `planned_pax` 0/NULL → check skipped. Threshold strict `<` 50% (exactly 50% does not flag). Soft/warn only. A hard precondition requires `attended >= 1` (else save rejected "No attendance records found"), so this anomaly only fires for `1 <= attended < planned_pax*0.5`.
- **Source.** `tms/routes/programme.py:527-529` (single add); attended `:469-471`; bulk re-check `:917-919`.

### Collar mismatch anomaly (attendees outside the target collar)

- **What it is.** For a single-collar-targeted session, count attendees of the OTHER collar; if any, flag.
- **Formula.** Only runs when `audience IN ('Blue Collared','White Collared')`. `mc = COUNT(*)` of attendees whose employee collar is set and differs from the session audience. Flag when `mc > 0`. Anomaly string `f'collar_mismatch({mc}_attendees)'`.
- **Inputs.** `calendar.target_audience` (audience); `emp_training` JOIN `employees` on `(plant_id, emp_code)`; `employees.collar`.
- **Filters.** `SELECT COUNT(*) FROM emp_training t JOIN employees e ON e.plant_id=t.plant_id AND e.emp_code=t.emp_code WHERE t.plant_id=? AND t.session_code=? AND e.collar IS NOT NULL AND e.collar != ?(audience)`. Guard: audience must be exactly `'Blue Collared'` or `'White Collared'` — `'Common'` or empty audience skips the check. Plant-scoped (**not host-aware**: join on `t.plant_id=e.plant_id` excludes cross-plant central attendees).
- **Edge cases.** Employees with NULL collar excluded (`e.collar IS NOT NULL`; note an empty-string `''` is technically NOT NULL so a blank-collar row would be compared and counted as a mismatch, but seed data uses non-empty collar). `'Common'`/blank audience never flagged. Soft/warn only.
- **Source.** `tms/routes/programme.py:530-537` (single add); bulk re-check `:922-930`.

### Time window vs Duration cross-check (`_validate_time_vs_duration`)

- **What it is.** Checks that (End Time − Start Time) × days ≈ Duration within 15 minutes. **HARD BLOCK** at 2C single save (red flash before any anomaly/DB write), at bulk 2C upload (row error, row skipped), and in the calendar bulk-row validator. 2A is intentionally exempt.
- **Formula.** `per_day_hrs = (minutes(time_to) - minutes(time_from)) / 60.0`; `days = max(1, (end_date - start_date).days + 1)` when both dates parse, else 1; `expected = per_day_hrs * days`; `diff_min = abs(expected - total_hours) * 60`; FAIL when `diff_min > tolerance_min` (default 15). Separate hard failures: (a) exactly one of time_from/time_to set → "{Start/End} Time is required when the other is provided."; (b) `time_to minutes <= time_from minutes` → "End Time must be after Start Time."
- **Inputs.** `time_from`, `time_to` (`'HH:MM'` → `_time_to_minutes`, valid only `0<=h<=23, 0<=m<=59`), `total_hours` (2C hours_actual / calendar duration), `start_date`, `end_date`.
- **Filters.** No DB filter — pure form-field validation. Tolerance = 15 min. At 2C single-add: time_from/time_to = form value else calendar fallback; total_hours = form hours_actual; dates from form. At bulk 2C the time window comes from calendar; hrs from the row.
- **Edge cases.** PASSES without checking when: both times empty, total_hours falsy/non-numeric, `total_hours <= 0`, or either time fails HH:MM parse. Multi-day handled: total_hours is treated as **cumulative** across days, `expected = per_day × days`. Threshold strict `> 15 min` (exactly 15 min slack passes). The required-both rule and end>start rule are hard fails independent of tolerance. In the calendar bulk validator there is ALSO a prior order check (tt <= tf) and a duration bounds check (`0 < dur <= 80` hrs) before the cross-check.
- **Source.** `tms/helpers.py:160-217` (`_validate_time_vs_duration`); minutes parser `:146-157` (`_time_to_minutes`); callers `tms/routes/programme.py:502-507` (2C single, hard block), `:933-938` (2C bulk, row error), `tms/helpers.py:2029-2031` (calendar bulk validator); duration bounds `tms/helpers.py:2021-2024`.

### Actual hours > 0 hard gate (2C)

- **What it is.** A 2C record cannot be saved with zero or blank actual hours.
- **Formula.** `hours = float(form 'hours_actual' or 0); if hours <= 0 -> reject`. Non-numeric input coerced to 0 via try/except → also rejected.
- **Inputs.** Form `hours_actual`.
- **Filters.** None — pure form check, runs first.
- **Edge cases.** Blank, non-numeric, 0, or negative all rejected. Precedes the FY-window check, time check, and all anomaly computation.
- **Source.** `tms/routes/programme.py:483-489`.

### FY-window hard gate on 2C start date (`_in_current_fy`)

- **What it is.** A 2C session's start date must fall inside the current FY (Apr 1 – Mar 31, IST).
- **Formula.** `fy_start, fy_end = _current_fy()`; reject when `start_date` is non-empty AND `NOT _in_current_fy(start_date)`.
- **Inputs.** Form `start_date`; `_current_fy()`/`_in_current_fy()` bounds (IST-based).
- **Filters.** Only enforced when `start_date` is provided (empty skips). Runs after the hours>0 gate and before the time-vs-duration check.
- **Edge cases.** Empty start_date is allowed past this gate. FY bounds in IST so Render (UTC) matches IST wall clock.
- **Source.** `tms/routes/programme.py:494-498`.

### Feedback score clamp (`_clamp_fb`, 0–4 scale)

- **What it is.** Sanitises the four feedback inputs (course_feedback, faculty_feedback, trainer_fb_participants, trainer_fb_facilities) before INSERT/UPDATE.
- **Formula.** `_clamp_fb(v)`: if `v is None` → None; `try float(v)` else None; if `v<0 or v>4` → None; return `v if v>0 else None`.
- **Inputs.** Form fields course_feedback, faculty_feedback, trainer_fb_participants, trainer_fb_facilities.
- **Filters.** Valid retained range is effectively `(0, 4]` — values `<0` or `>4` → NULL; exactly 0 → NULL (treated as blank).
- **Edge cases.** Non-numeric → NULL; 0 → NULL; out-of-range → NULL. On UPDATE of an existing 2C row, a blank/NULL/0 form value falls back to the pre-existing stored value (COALESCE-style) rather than overwriting it.
- **Source.** `tms/routes/programme.py:551-567`.

---

## Cross-screen Consistency

The same calculation engine backs multiple screens so the numbers agree. Reviewers can rely on the following identities:

- **Compliance gauges == Summary compliance == Central per-plant drill-in.** `_calc_compliance(plant_id, db)` (`tms/helpers.py:628`) is the single engine for the SPOC Dashboard gauges (`auth.py:842-845`), the async JSON (`/api/dashboard-monthly`, `api.py:220-242`), the Monthly Summary route (`summary.py:19`), and `/central/plant` (`central.py:769`). All four therefore share the exit policy (active-headcount denominator, all-attendees FY-bound numerator) and can exceed 100% identically.

- **Worst-cells targets == gauge plant scope.** `_calc_worst_cells` is invoked **inside** `_calc_compliance` (`tms/helpers.py:675`), so the Improvement Areas panel is scoped to the same plant as the gauges on the same page.

- **Monthly Summary table == Central per-plant Summary card.** Both render from `_calc_summary` + `_calc_totals` (`tms/helpers.py:424`, `:578`). `/central/plant` (`central.py:758-777`) calls these unchanged, so a plant's Summary numbers are identical whether viewed by its own SPOC or by Central.

- **Conducted gate is shared.** Monthly Summary (`programme_details LEFT JOIN calendar`, `status='Conducted' OR no calendar row`), the Dashboard monthly "sessions" series, and Export all treat `'Awaiting Verification'` as not-yet-conducted, so programme counts agree across those screens.

- **Single FY/IST source of truth.** Every FY-bound figure derives its window from `_current_fy()` and `_fy_label()` (both IST via `_today_ist()`), so the dashboard, summary, central rollup, calendar in-form panel, and TNI lock all roll the financial year at the same IST midnight.

- **Two coverage definitions are intentionally different — do not expect them to match.** (1) Monthly Summary / QC coverage = distinct employees **trained** (`emp_training`) ÷ TNI nominees, grouped by prog_type+collar (canonical). (2) Calendar planning coverage = planned **seats** (`calendar.planned_pax`) ÷ distinct TNI nominees, per programme, capped at 100%. They measure different things (people trained vs seats scheduled) and will legitimately diverge.

- **Charts that are all-time vs FY-bound — a documented split.** The `/api/dashboard-monthly` per-month seats/hours bars are **all-time** (no FY filter), whereas the gauge block in the *same* endpoint is **FY-bound**. Similarly, QC Pareto/Heatmap use all-time trained pairs while QC Cumulative Coverage is FY-bound. These will not reconcile with each other by design.

- **Central fleet totals deliberately differ from the sum of per-plant rows.** A central-hosted session is counted once per attending plant in the per-plant `sessions`/`conducted` columns, but only once fleet-wide via `grand_central`. So `grand['sessions'] != sum(plant['sessions'])` is expected, not a defect.
