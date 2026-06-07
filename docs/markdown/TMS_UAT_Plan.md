# TMS — Rigorous User-Experience / UAT Test Plan

Manual, end-to-end test script grounded in actual code (fields, validations, exact flash strings, edge cases). Tick `[x]` as you go. Each case = **Do → Expect → Try-to-break**. "Expect" strings are the real messages from code.

- **Live:** https://bcml-training-mis.onrender.com
- **Logins:** SPOC `balrampur` / `bcml@1234` · Central `central` / `bcml@1234` · Admin `admin` / `admin@bcml`
- **Run order:** do SPOC cycle top-to-bottom — later modules need earlier data.
- **Two sessions:** normal browser (SPOC) + incognito (Central) to test cross-role live.
- **Key constants:** lockout 5 fails / 15 min · session timeout 2 hr (5-min warning at 25 min) · password min 10 chars · time-vs-duration ±15 min · upload max 16 MB · BC target 12 h/yr, WC 24 h/yr · effectiveness due = conducted + 90 days, overdue = +30 past due.

---

## A. AUTH & SESSION

- [ ] **Valid login** → SPOC lands on Dashboard; Central/Admin land on All Plants Overview.
- [ ] **Wrong password ×1–4** → `Invalid username or password. N attempt(s) before lockout.` (counts down 4→1)
- [ ] **5th wrong** → `Too many failed attempts. Account locked for 15 minutes.`; retry while locked → `Account locked. Try again in 15 minutes.`
- [ ] **Lockout auto-clears** after 15 min → normal login works.
- [ ] **Unknown username** → `Invalid username or password.` (no count — doesn't reveal user exists).
- [ ] **must_change_password user** → forced to change-password, flash `Please set a new password before continuing.`
- [ ] **Password policy** (change-pw): try 9 chars / no digit / no symbol / = username → rejected; reuse current → `New password must be different…`; weak `bcml@1234` → banned.
- [ ] **2FA mandatory for Central/Admin** — if not enabled, any page → `Two-factor authentication is mandatory for this role. Please set it up to continue.` → /2fa/setup. Scan QR, enter 6-digit → `Two-factor authentication enabled successfully.`; wrong code → `Invalid code — scan QR again and retry.`
- [ ] **Session timeout** — idle ~25 min → `confirm` "session will expire in 5 minutes…"; ignore → auto-logout at 30 min.
- [ ] **Logout** clears session; browser Back shows no data (redirects to /login).
- [ ] **CSRF** — submit a stale form (open form, log out elsewhere, submit) → `Session expired or form was stale — please try again.` (HTTP 400).

---

## B. SPOC — EMPLOYEE MASTER

- [ ] **List cap** — banner `Showing first 200 of 1115 employees by name…`; click **Show all** → loads full (slower); toggle back.
- [ ] **Search + filters** (Collar/Dept/Section/Gender/Status) — correct rows; no match → `No employees match the current filters. Try clearing search or filters above, or use Add New Employee to create one.`
- [ ] **Add employee** — all required fields enforced per field (`{Field} is required.`); success `Employee {name} added successfully.`
- [ ] **Duplicate emp code** → `Employee code {code} already exists.`
- [ ] **Invalid enums** — bad grade/collar/category/gender/PH each give exact message (e.g. `Invalid collar 'X'. Must be Blue Collared or White Collared.`).
- [ ] **TomSelect fuzzy** (Designation/Dept/Section) — type near-duplicate → `Similar value exists: "X". Use it instead of "Y"?`
- [ ] **Bulk upload** — download template, fill, upload; non-xlsx → `Please upload a valid .xlsx file.`; row errors listed `Row N (code): … is required` / `…already exists…— skipped`; success `N employee(s) uploaded successfully.`
- [ ] **Bulk update** — blank cells = no change; missing Emp Code header → `Header "Emp Code" not found — required to match rows.`; success `N employee(s) updated successfully.`
- [ ] **Exit employee** — future exit date blocked; reason required; pending effectiveness reviews → blocked `Cannot exit {code}: N pending effectiveness review(s) open…`; force-tick → exits with warning.
- [ ] **Undo Exit** — exited within 7 days shows in undo panel; reactivate → returns to active (`?show_exited=1` view).
- [ ] **Verify count** vs DB: active total = `COUNT(*) WHERE is_active=1`.

---

## C. SPOC — YEARLY TNI

- [ ] **Add TNI Entry** — pick employee → Grade/Collar/Dept auto-fill (readonly); Programme dropdown (TomSelect) only lists Programme Master; Source defaults `TNI Driven`.
- [ ] **Required missing** → `Employee Code and Programme Name are required.`
- [ ] **Programme not in master** → `"X" is not in Programme Master. Add it to Programme Master first…`
- [ ] **Duplicate emp+programme** → `TNI entry for "X" already exists for employee {code}.` (UNIQUE held).
- [ ] **Source coercion** — any value other than `TNI Driven`/`New Requirement` silently becomes `TNI Driven`.
- [ ] **Negative hours** → `Planned hours must be a non-negative number…`
- [ ] **Quick-Add (multi-programme for one emp)** — no emp → `Pick an employee first.`; no programme → `Select at least one programme.`; hours ≤0 → coerced to 4 with warning; success `N TNI entries added for {code}. (M already existed — skipped)`.
- [ ] **Excel upload → analyze → confirm** at `/tni/analyze/confirm` — fuzzy-match preview; bad file → `Could not read file: …`; within-file dupes skipped.
- [ ] **Data-quality panel** appears when dupes / name-mismatches exist → `View & Remove Duplicates`, `Review & Fix Names`.
- [ ] **Edit Source inline** (non-TNI rows show dropdown) — change persists via AJAX.
- [ ] **Per-row delete** → confirm `Delete this TNI entry?`.
- [ ] **Bulk delete** — <50: `Delete N TNI entries? This cannot be undone.`; ≥50 or select-all: must type `DELETE N`; mismatch → `Confirmation text did not match. Delete cancelled.`
- [ ] **Pagination** — 30/page, search debounced 320 ms; empty → `No TNI entries match the current filters. Clear filters or upload a TNI file to get started.`
- [ ] **Mouse-free table** (your fix) — desktop: ◀▶ buttons appear when wide, click scrolls; click table + arrow keys scroll; Employee pinned left, Action pinned right.
- [ ] **Mobile** (<768): #, Mode, Planned Hrs hidden; (<560): also Collar, Dept, Type hidden; swipe hint shows.
- [ ] **TNI window closed** (admin-only to test off-FY) → `TNI window is closed (FY ended March 31)…`

---

## D. SPOC — PROGRAMME MASTER

- [ ] **Empty state** before any data → auto-build message shown.
- [ ] **Add programme** — name smart-titled; duplicate without override → `"X" already exists… re-submit with the Override checkbox…`; with override → `"X" overridden — Type/Source/Category updated.`; new → `"X" added to master as {source} · {category}.`
- [ ] **Gatekeeper** — confirm a programme NOT in master can't be picked in Calendar/2A/2C (dropdown shows nothing).
- [ ] **Category toggle** (click badge / right-click) — General↔Specialized; Specialized toast mentions 3-month effectiveness; demote with existing reviews → `Cannot demote to General — N effectiveness review row(s) already exist…` (409).
- [ ] **Delete in-use programme** → `Cannot delete "X" — it is referenced in TNI, Calendar, or Training Records.`; unused → `"X" removed from master list.`
- [ ] **Sync from TNI** (admin) — ⚠️ wipes TNI-Requirement rows, rebuilds; success `Programme Master synced from TNI — N… New Requirement entries preserved.` **Verify**: a manually-added New-Requirement programme survives; confirm warning understood.

---

## E. SPOC — TRAINING CALENDAR

- [ ] **Add session** — Programme dropdown shows TNI demand; Type auto-fills + locks from master (badge); session code auto = `BCM/TYP/26-27/B01`.
- [ ] **Audience auto-lock from TNI** — set a different audience → overridden, info `Audience set to "X" (locked from TNI).`; New-Requirement programme (no TNI) → audience editable.
- [ ] **Date range** — end < start → rejected.
- [ ] **Time vs duration** — set 09:00–10:00 but Duration 1.5 → `Time window does not match Duration. 09:00–10:00 = 1.00 hrs, but Duration is set to 1.5 hrs. Fix one…`; within ±15 min → accepted.
- [ ] **Only one of start/end time** → `Start Time is required when the other is provided.` (or End).
- [ ] **TNI coverage panel** — demand vs planned vs covered vs gap; coverage summary table status badges (Fully/Under/Not Planned).
- [ ] **Conducted lock** — try edit a Conducted session (non-admin) → `Conducted sessions cannot be edited.`; delete → blocked.
- [ ] **Mark Conducted gating** — without QR/2A/feedback → `Can't mark Conducted — missing: QR code, attendance (no 2A rows), feedback responses…`
- [ ] **Lapsed auto-archive** — To-Be-Planned sessions with plan_start < FY start → Lapsed; toggle `include_lapsed`.
- [ ] **Filters** — Status/Type/Mode/Month/Audience filter correctly (client-side).
- [ ] **Bulk upload calendar** — audience still derived from TNI on each row.

---

## F. SPOC — PROGRAMME DETAILS (2C)

- [ ] **Pick Session Code** → programme/type/dates/hours back-fill from calendar.
- [ ] **Actual Start Date required**; feedback scores 1–4 only.
- [ ] **Save** → calendar status flips to Conducted; **Delete** → back to To Be Planned.
- [ ] **Calendar vs new** — known code = Calendar Program; unknown = New Program.
- [ ] **Anomaly icon** appears on rows with flags; hover shows detail.
- [ ] **Man-Hours column** = participants × hours — cross-check one row by hand.

---

## G. SPOC — TRAINING RECORDS (2A / Attendance)

- [ ] **Add attendee** — employee auto-fills collar/dept; Session Code optional (auto-fills if found); `[CENTRAL]` sessions selectable.
- [ ] **Hours required >0** → `Training hours must be greater than 0.`
- [ ] **Inactive emp** → `Employee "{code}" not found or inactive for this plant.`
- [ ] **Session not in calendar** → warning `Session code "X" not found in calendar.`
- [ ] **Cancelled/Re-Scheduled session** → `Session X is {status}. Cannot record attendance.`
- [ ] **Future date** → `Cannot log future-dated training. Start date "X" is after today.`
- [ ] **Outside FY** → `Training date must be within the current financial year (start–end).`
- [ ] **Score out of 0–100** → `Pre/Post-Session Score must be between 0 and 100 (got X).`
- [ ] **Both/neither time** rule enforced (`Start Time and End Time must both be provided (or both blank).`).
- [ ] **Duplicate** (same emp+session+prog+date) → `Duplicate record — this employee already has a training entry for this programme on this date.`
- [ ] **Anomaly flags** (allowed, tagged for Central): add WC person to BC session → `collar_mismatch`; hours > session×1.25 → `hours_over`; out-of-window date → `date_outside`. Success then reads `Training record added with N anomaly flag(s) — Central L&D will review.`
- [ ] **Non-TNI attendee** — counts seats/man-hours but NOT coverage% (verify in Summary).
- [ ] **Empty state** → `No training records yet. Pick a Session Code from the form above (or use Bulk Upload)…`

---

## H. SPOC — DASHBOARD

- [ ] **KPI cards** — Active Workforce (BC/WC chips), Man-Hours vs FY target with progress bar (green ≥75 / yellow 50–75 / red <50). Counter animates.
- [ ] **Pending alert** if any → `N session(s) awaiting Central L&D verification…`
- [ ] **3 gauges** (BC/WC/Overall) animate; click → man-hour drilldown modal.
- [ ] **Drilldown** — dept cards → drill to employees; search filters; KPI strip (At/Below/Zero/Total/Overall%); breadcrumb back; empty → `No departments/employees match.`; fail → `Failed to load data.`
- [ ] **Monthly bar chart** — skeleton bars then real; metric toggle Man-Hrs/Seats/Sessions; click month → detail panel; error → `Unable to load monthly data.`
- [ ] **QC charts (6)** load from `/api/dashboard-qc`; error → `failed to load`:
  - [ ] Pareto uncovered (top 8) · Histogram hrs/emp (BC/WC, target+avg lines) · Dept man-hour compliance · Coverage heatmap (collar×type, 6 color tiers) · Cumulative coverage BC/WC · Cumulative man-hours vs target.
- [ ] **Verify a number** — Total Man-Hours = `SUM(hrs) FROM emp_training WHERE plant_id=? AND start_date in FY`.

---

## I. SPOC — MONTHLY SUMMARY

- [ ] **Month filter** — default Cumulative (All); pick month auto-submits; Clear resets.
- [ ] **3 compliance cards** (BC/WC/Overall) — %, actual, mandate (= emp×target), bar.
- [ ] **MIS table** — Programmes (BC/WC/Cmn/Tot/Int/Ext), Seats (BC/WC/Tot), Man-Hours, %Covered YTD; row color by coverage (green ≥100 / yellow 75–100 / red <75 / grey 0); TOTAL footer sums.
- [ ] **Empty** → `No training records yet` + "Data populates from Training Records (2A)…".
- [ ] **Cross-check** one prog_type row's seats/hours/coverage against 2A entries.
- [ ] **Export Excel** from here → file downloads, rows = current view.

---

## J. SPOC — FEEDBACK REPORTS

- [ ] **Index list** — Programme, Session, Date, Status, Responses count, Overall%; empty → `No feedback collected yet. Generate a Feedback QR from the Training Calendar and ask participants to scan.`
- [ ] **Report page** — 4 scorecards (Responses, Programme, Trainer, Overall %); 9-question analysis (Q1–5 programme, Q6–9 trainer) with SD/D/A/SA distribution bars, avg, %; subtotals; individual responses table; Key Learnings + Suggestions (only if present); Print hides buttons.
- [ ] **Empty report** → `No feedback responses yet for this session.`
- [ ] **Verify** one question avg = SUM(responses)/COUNT.

---

## K. SPOC — EFFECTIVENESS REVIEWS

- [ ] **4 scorecards** (Pending/Due/Overdue/Completed) — click filters list; badge in sidebar matches open count, red if overdue>0.
- [ ] **Status logic** — pending (today<due) / due (≤30 days past) / overdue (>30 past) / completed.
- [ ] **File Review modal** — Rating 1–5 required; need ≥1 of Behaviour-Change / Application-on-Job ≥10 chars → else `Provide at least one observation (Behaviour Change OR Application on Job, min 10 chars).`
- [ ] **Completed → View modal** shows filed data + filer.
- [ ] **Seeding** — only Specialized sessions seed reviews (due = conducted + 90 days). Empty → `Reviews are auto-seeded when a Specialized session is Conducted.`

---

## L. SPOC — EXPORT

- [ ] **Export config** — pick sheets (emp/tni/calendar/2a/2c) + month/collar/dept/source filters → Excel.
- [ ] **Filename** = `BCML_{unit}_Training_MIS_{fy}{filters}.xlsx`.
- [ ] **Spot-check** TNI sheet "Completed?" = Yes only where emp_training row exists.

---

## M. CENTRAL (login `central`, read-only)

- [ ] **All Plants Overview** — cross-plant KPIs.
- [ ] **Central Calendar / Attendance** — all plants; read-only (no add/edit/delete controls).
- [ ] **Review Queue** (your regrouped sidebar): yellow group renders together; counts match.
  - [ ] **Verify Sessions** — sessions Awaiting Verification appear; approve/reject; approve seeds effectiveness for Specialized; badge `pending_verify_count` updates.
  - [ ] **Anomalies Review** — 2A/QR anomaly-flagged rows appear; `anomaly_count` badge.
- [ ] **SPOC Upload Errors** — per-plant upload error log.
- [ ] **Programme Master (central) / Corp Members** — view.
- [ ] **Consolidated Export** (green) — multi-plant workbook, Summary sheet per plant + TOTAL.
- [ ] **Role enforcement** — paste a SPOC write URL as central → `Access denied.` → bounced.
- [ ] **2FA gate** — central without 2FA can't proceed (mandatory).

---

## N. ADMIN (login `admin`)

- [ ] **Impersonate plant** — `/admin/plant/<id>` → banner `Admin — viewing as SPOC for {plant} ({unit})`; flash `Now viewing as SPOC for {plant}…`; **Exit & Return to Central** clears plant.
- [ ] **User Management** — create/edit user (role, plant, must_change_password); reset password; enable/disable TOTP.
- [ ] **Override Requests** — approve/reject SPOC override (e.g. closed-TNI-window) requests.
- [ ] **TNI Archives** — view archived FY data.
- [ ] **Audit Log** — actions recorded (LOGIN_OK/FAIL, RECORD_*, BACKUP_*, qr_*); SHA-256 chain present; run verify_chain if exposed.
- [ ] **Backup & Restore** — download backup; restore is destructive → must confirm.
- [ ] **Organisation Settings** — change man-hour targets / config; confirm it flows to Dashboard/Summary (global-scalability: no hardcoded BCML values).
- [ ] **Role gates** — admin-only routes 403/redirect for spoc & central.

---

## O. QR PUBLIC (logged-out, phone)

- [ ] **Landing** `/q/<token>` — attendance or feedback form by stage; lang `en`/`hi`.
- [ ] **Expired/invalid token** → `qr_error.html` "invalid or has expired" (HTTP 410).
- [ ] **PIN gate** (if set) — wrong → `Incorrect session code. Ask your trainer for the 4-digit code.`
- [ ] **Time gate** — before plan_start → `Session has not started. Attendance opens on {date}.`
- [ ] **Status lock** — only To-Be-Planned accepts; else `Attendance closed — session is "{status}".`
- [ ] **Unknown emp** → `Employee code "X" not found[ for {plant}].`
- [ ] **Collar mismatch** — allowed but flagged for Central.
- [ ] **Duplicate scan** → thanks page "already marked".
- [ ] **Feedback** — 9 ratings (1–4), key learnings/suggestions (≤1000); anon dedup by IP; closed status → `Feedback closed — session is "{status}".`
- [ ] **QR generate/revoke** (SPOC) — only To-Be-Planned; generate → poster; revoke → `QR revoked — old QR no longer accepts scans.`

---

## P. CROSS-CUTTING

- [ ] **Dark mode** — toggle on every page; no white flash on reload; persists (localStorage `tms-theme`).
- [ ] **Responsive** — redo TNI / Employee / Calendar at phone width; hamburger menu; plant chip hides on small.
- [ ] **Cache-bust** — after a deploy, CSS/JS URL has new `?v=` (STATIC_VER).
- [ ] **Error pages** — hit `/nonexistent` logged-in → `Page not found.`→ home; logged-out → /login.
- [ ] **500** — (if reproducible) → `Server error. Please try again or contact Corporate L&D.`
- [ ] **Upload >16 MB** → `File too large. Maximum upload size is 16 MB.`
- [ ] **/health** → JSON 200 (load-balancer check).
- [ ] **Render-vs-local** — anything time-based reads correctly in IST on live (timezone trap).

---

### Sign-off
| Area | Tester | Date | Pass/Fail | Notes |
|---|---|---|---|---|
| Auth | | | | |
| SPOC cycle | | | | |
| Reports | | | | |
| Central | | | | |
| Admin | | | | |
| QR public | | | | |
| Cross-cutting | | | | |

*Strings/values sourced from code as of this commit. If a flash message or rule changes, update the matching line here.*
