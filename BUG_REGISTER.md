# TMS Bug Register

Single source of truth for the bug-fix campaign. Sources: adversarially-verified agent audit (`wf_85b29ad8-896`, 3-vote) + deterministic smoke/static harness (`tests/`). Status: `open` until fixed + verified.

See plan: `~/.claude/plans/continue-giggly-crane.md`.

## Status log
- **Batch 0 FIXED + demo-verified + deployed:** #5 (seed-demo now gated behind `ALLOW_DEMO_SEED=1` env + typed "SEED DEMO DATA" phrase — proven it cannot wipe: emp_training 21660→21660 on blocked POST), #1 (central_duplicates 500), S1 (training-hours report 500). Smoke harness now 0 hard fails.
- **#8 → FALSE POSITIVE:** admin_seed_demo.html copy already states "Plant 1 *TNI* preserved" + "all 10 plants" calendar/attendance/2C deleted — accurate, no contradiction. No code change to delete scope.
- Remaining: 39 agent-confirmed + Phase R (10 uncovered modules + 3 lenses).

## Coverage status (honest)
- **Agent audit: PARTIAL.** 41 confirmed. 10 module finders died on API 529 (auth, employees, dashboard, programme_master, tni, calendar, programme_2c, training_2a, effectiveness, feedback) + 3 lenses unverified (xc-schema-migration, xc-chain-logic, verify). Re-run after 17:30 IST (Phase R).
- **Deterministic net: smoke harness (87 GET routes × 3 roles) + url_for check — DONE.** Catches CRASH/SQL_TEMPLATE across ALL modules. Found 1 new bug the agents missed.

## Deterministic finds (smoke / static harness)
| # | Sev | Area | Bug | File | Status |
|---|---|---|---|---|---|
| S1 | CRITICAL | reports | `/reports/training-hours` 500 — `IndexError: r['target_hrs']` (SELECT missing column the route reads) — **NEW, agents missed it** | tms/routes/reports.py | open |
| S2 | — | — | confirms agent #1 (central_duplicates 500) | — | dup of #1 |

Smoke: 3 hard fails across 2 endpoints (central_duplicates ×2 roles + training_hours_report). url_for: 0 unresolved (214 checked).

## Agent-audit confirmed bugs (41) — severity order

| # | Sev | Area | Class | Bug | File | Status |
|---|---|---|---|---|---|---|
| 1 | CRITICAL | central | CRASH | central_duplicates.html crashes (500) whenever any duplicate cluster exists — Jinja `sum(attribute='clusters')` adds lis | templates/central_duplicates.html | open |
| 2 | CRITICAL | central | CRASH | merge_cluster() raises unhandled IntegrityError (500 + all merges lost) when same emp is nominated/trained under both th | tms/master_dedup.py | open |
| 3 | CRITICAL | intelligence | DOMAIN_LOGIC | Programme Intelligence aggregates TNI across ALL financial years but labels and measures coverage for the current FY onl | c:\Users\yogendra.das\Desktop\Training Management System\tms\routes\ap | open |
| 4 | CRITICAL | xc-consistency | DOMAIN_LOGIC | Coverage % is computed two incompatible ways — Monthly Summary vs Dashboard show different numbers for the same plant | tms/helpers.py (lines 528-545 _calc_summary Q4; 605-630 _calc_totals)  | open |
| 5 | CRITICAL | xc-security | SECURITY | admin_seed_demo wipes ALL plants' calendar/2A/2C on production | tms/routes/auth.py | open |
| 6 | HIGH | export | DOMAIN_LOGIC | Consolidated 2C sheet shows 0 Participants & 0 Man-Hours for all Central (plant 99) programmes | tms/routes/export.py | open |
| 7 | HIGH | admin | CONSISTENCY | SPOC override request form never collects the typed payload — entire typed-executor pipeline is dead for TNI_ADD / MARK_ | templates/spoc_request_form.html (lines 19-44) vs tms/routes/requests. | open |
| 8 | HIGH | admin | DOMAIN_LOGIC | Seed Demo wipes calendar / emp_training / programme_details for ALL plants including Plant 1, contradicting the 'Plant 1 | tms/routes/auth.py (admin_seed_demo lines 918-933) + templates/admin_s | open |
| 9 | HIGH | intelligence | DOMAIN_LOGIC | Coverage % counts every emp_training attendee, including New-Requirement / non-TNI-nominated employees, violating the do | c:\Users\yogendra.das\Desktop\Training Management System\tms\routes\ap | open |
| 10 | HIGH | xc-consistency | DOMAIN_LOGIC | Summary coverage counts prior-FY training toward current-FY TNI fulfilment (no FY date filter on the EXISTS) | tms/helpers.py:532-536 (_calc_summary Q4) and 613-630 (_calc_totals bc | open |
| 11 | MEDIUM | qr | SQL_TEMPLATE | QR poster / attendance / feedback pages never show Venue — calendar has no venue column and _validate_token can't select | tms/routes/qr.py:31-46 (_validate_token); templates/qr_poster.html:66- | open |
| 12 | MEDIUM | qr | DOMAIN_LOGIC | Feedback Reports index Overall % drops any response with a single unanswered question (NULL-poisoned SUM), diverging fro | tms/routes/qr.py:372-380 (feedback_reports_index); templates/feedback_ | open |
| 13 | MEDIUM | summary | DOMAIN_LOGIC | Month filter on programme count is year-agnostic and not FY-bound, so prior-FY programmes of the same month inflate the  | tms/helpers.py | open |
| 14 | MEDIUM | central | ENVIRONMENT | tni_upload_errors.ts uses SQL DEFAULT datetime('now','localtime') (UTC on Render) but central_tni_errors buckets/compare | schema.sql | open |
| 15 | MEDIUM | export | DOMAIN_LOGIC | Per-plant export sheets (TNI, Calendar, 2A, 2C) are all-time but the sheet header claims the current FY | tms/routes/export.py | open |
| 16 | MEDIUM | export | DOMAIN_LOGIC | Month filter silently drops New-Program (non-calendar) sessions from the 2C sheet in both exports | tms/routes/export.py | open |
| 17 | MEDIUM | export | DOMAIN_LOGIC | Consolidated Summary 'Sessions Conducted' excludes central-hosted sessions, disagreeing with the Central Dashboard | tms/routes/export.py | open |
| 18 | MEDIUM | export | CONSISTENCY | Consolidated TNI sheet exports all financial years under an 'FY {fy}' header, while its Completed? flag is FY-bound | tms/routes/export.py | open |
| 19 | MEDIUM | admin | ENVIRONMENT | spoc_requests.ts uses datetime('now','localtime') — stored in UTC on Render, displayed ~5.5h off and inconsistent with I | schema.sql line 249 (and tms/db.py _migrate_spoc_requests line 696) | open |
| 20 | MEDIUM | admin | REGRESSION_RISK | Approve path commits the executed record mid-transaction (via log_action) but treats the whole approve as one rollback-a | tms/routes/requests.py (admin_review_request lines 259-278) + tms/rout | open |
| 21 | MEDIUM | xc-consistency | DOMAIN_LOGIC | Calendar 'Demand' column counts all-FY TNI but the Audience column beside it is current-FY only | tms/routes/calendar.py:61-63 (demand_map) and 81-90 (cov_rows) vs tms/ | open |
| 22 | MEDIUM | xc-consistency | ENVIRONMENT | feedback_response / verification_log / spoc_requests / reschedule-history timestamps use datetime('now','localtime') = U | schema.sql:152 (feedback_response.submitted_at), 272 (verification_log | open |
| 23 | MEDIUM | xc-consistency | CONSISTENCY | Calendar bulk-upload INSERT omits the category column that single add/edit set from Programme Master | tms/routes/calendar.py:488-495 (calendar_bulk_upload INSERT) vs 163-17 | open |
| 24 | MEDIUM | xc-security | CRASH | qr_set_pin reads session['plant_id'] before role check — KeyError 500 for central with no plant | tms/routes/qr.py | open |
| 25 | MEDIUM | xc-security | SECURITY | qr_attend / qr_feedback are CSRF-exempt and PIN-gating is optional — forgeable public writes | tms/routes/qr.py | open |
| 26 | MEDIUM | xc-performance | PERF | find_duplicates runs 4 COUNT queries per programme inside clusters, looped over all 10 plants on the Central Duplicates  | tms/master_dedup.py | open |
| 27 | LOW | qr | UI | Public QR feedback/thanks/poster/error pages use var(--fs-*) CSS vars that are never defined (main.css not loaded) — tex | templates/qr_feedback.html (13 uses), templates/qr_poster.html:45, tem | open |
| 28 | LOW | qr | VALIDATION | Anonymous feedback IP dedup can be bypassed when remote_addr is empty — dedup checks '' but row stored as NULL | tms/routes/qr.py:822, 841-846, 874 | open |
| 29 | LOW | summary | PERF | _calc_compliance runs _calc_worst_cells subquery on every /summary load though summary.html never uses worst_cells/headl | tms/helpers.py | open |
| 30 | LOW | central | VALIDATION | central_programme_add / central_corp_member_add report 'already exists' for ANY exception, masking real DB errors | tms/routes/central_training.py | open |
| 31 | LOW | central | VALIDATION | central_programme_add accepts arbitrary prog_type (no server-side enum validation) — bad value propagates into Programme | tms/routes/central_training.py | open |
| 32 | LOW | export | UI | Consolidated Summary TOTAL row omits BC/WC man-hours, mandates and compliance % totals | tms/routes/export.py | open |
| 33 | LOW | export | PERF | Summary sheet issues ~40 uncached get_config queries (2 keys × 10 plants, plant-scoped bypasses request cache) | tms/routes/export.py | open |
| 34 | LOW | admin | CRASH | Legacy/empty-payload approval of a non-OTHER request raises KeyError instead of the intended friendly 'legacy free-text' | tms/routes/requests.py (admin_review_request lines 249-276, _execute_r | open |
| 35 | LOW | admin | VALIDATION | MANUAL_ATTENDANCE dup-guard checks (session_code) but emp_training UNIQUE is (programme_name,start_date) — a plain INSER | tms/routes/requests.py (_execute_request MANUAL_ATTENDANCE lines 130-1 | open |
| 36 | LOW | admin | CONSISTENCY | Ephemeral-storage warning text/logic references /app/data while the live persistent disk is /var/data | tms/routes/admin.py (admin_settings lines 31-37) + templates/admin_set | open |
| 37 | LOW | intelligence | UI | Per-programme coverage bar/label can render >100% ('167% covered') | c:\Users\yogendra.das\Desktop\Training Management System\templates\int | open |
| 38 | LOW | xc-consistency | CONSISTENCY | tni_set_source seeds programme_master without the matching source, leaving master source inconsistent with the TNI row | tms/routes/tni.py:381-385 (tni_set_source) | open |
| 39 | LOW | xc-consistency | DOMAIN_LOGIC | QR attendance for a Central session stores full session duration as the attendee's hrs and never flags anomalies | tms/routes/qr.py:725-733 (central-session INSERT branch) | open |
| 40 | LOW | xc-security | SQL_TEMPLATE | feedback_reports_index SELECT alias drift — AVG over CASE expressions yields wrong/garbled avg_score | tms/routes/qr.py | open |
| 41 | LOW | xc-security | SECURITY | qr_landing / qr_thanks / public scan pages render plant directory data with no authn (token-gated only) | tms/routes/qr.py | open |
