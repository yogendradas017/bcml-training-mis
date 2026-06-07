# TMS — Senior Engineer Prompts

Ready-to-use senior-engineer prompt set for the Training Management System.

Each one is also a **slash command** (lives in `.claude/commands/`). Two ways to use:
- **Slash:** type `/audit-codebase calendar module` in Claude Code.
- **Paste:** copy the prompt block below into any chat, replace the `«…»` placeholder.

**Always-on TMS guardrails** (every prompt assumes these — see `CLAUDE.md`):
- Read `CLAUDE.md` first. Respect: `programme_master` gatekeeper · audience-derived-from-TNI · source enum is exactly `TNI Driven` / `New Requirement`.
- **TMS-Wide Consistency Rule** — any enum / label / source / collar change must propagate across **all** templates AND routes.
- Never rename Flask endpoints (breaks `url_for`). No ORM without approval. Raw `sqlite3`, parameterised (`?`) only.
- Zero-defect: run `python -c "import app; print('OK')"` before declaring done. Trace SQL→template both directions. Test timezone/files/external on Render, not just locally. Never `defer` GSAP on `login.html`.

---

## 1. `/audit-codebase` — Senior audit (architecture, duplication, perf, fixes)

> Act as senior engineer joining TMS unfamiliar. Reverse-engineer architecture + data flow first, then surface issues. Do NOT change functionality — only upgrade quality, scalability, maintainability.
>
> **Output in order:**
> 1. **Architecture breakdown** — words-diagram of request → route → helper → SQL → template; note deviations from the modular ideal.
> 2. **Findings table** — `| # | Category | File:Line | Severity | What | Why it matters |` (categories: bad-arch · duplicate-logic · perf-bottleneck · scalability-risk · maintainability · audit-trail-gap · security).
> 3. **Refactor proposals** — top 5 findings: concrete before/after diff sketch with file path. Respect Consistency Rule.
> 4. **Quick wins** — items < 30 min, list only.
>
> Rules: don't rewrite before listing the finding; cite file:line; only edit if I say "apply 1,3,5"; run import check before done.
>
> Focus: «optional — e.g. calendar module / TNI pipeline»

---

## 2. `/prod-debug` — Production debug (trace, root-cause, fix, verify)

> Act as senior debugging engineer on a TMS live production issue (Render = live). Do NOT guess. Trace first.
>
> 1. **Restate symptom** exactly. 2. **Reproduce locally** or state why not. 3. **Trace** route → helper → SQL → template; find where reality diverges from intent. 4. **Root cause** — one sentence, distinct from symptom. 5. **Edge cases** — 3+ states triggering same/related failure. 6. **Fix** — patch with file:line + WHY each line. 7. **Verify** — import check; describe migration if schema/DB touched; call out Render-vs-local if timezone/files/external. 8. **Regression guard** — what check would have caught it; add if cheap.
>
> Forbidden: bypassing CSRF/2FA/hooks, `--no-verify`, fixing symptom without root cause.
>
> Bug: «describe / paste error / point to screenshot»

---

## 3. `/perf-optimize` — Performance pass (identify, fix, measure)

> Act as senior performance engineer on TMS. Target: Render Starter (512 MB RAM, 1 CPU). Past incident: pandas OOM'd worker on TNI analyze → fixed via openpyxl streaming.
>
> Investigate hot paths (Dashboard, /tni/data, /calendar, /training, /api/dashboard-monthly): SQL shape (N+1 / scan / missing index), memory (pandas / big template dict), render cost (DOM size / dup TomSelect). Produce **bottleneck table** `| File:Line | Cost | Fix | Effort |`. Flag unbounded growth (caches, dicts, leaked connections).
>
> Standard plays: missing index → `_ensure_indexes` (idempotent); pandas → openpyxl streaming >1k rows; materialised join → `EXISTS()`; 1000+ row loop → paginate (TNI 30/page pattern); `SELECT *` used for 3 cols → narrow it.
>
> Verify: import check; if SQL, confirm index used; if template, confirm DOM not larger. Output: bottleneck table · top 3 fixes applied · before/after estimate · what was deliberately NOT changed and why.
>
> Focus: «optional endpoint / feature»

---

## 4. `/refactor-clean` — Refactor to clean architecture (same behavior)

> Act as senior architect refactoring TMS. **Do NOT change product behavior** — improve only architecture + code quality.
>
> Constraints: routes registered directly (not Blueprints) — endpoint names stay stable; raw sqlite3, no ORM; `tms/helpers.py` splittable but all import paths stay valid; Consistency Rule applies.
>
> Plays in order: 1. Separate concerns (SQL + logic + flash out of fat routes into pure helpers). 2. Reduce duplication (same SQL in 3 routes → 1 helper; add/edit/bulk validation → 1 validator). 3. Tighten coupling (long arg lists → data-attribute dispatch). 4. Clarify state machine (3+ if-status branches → dict/table).
>
> Output: new structure (if splitting) · findings table w/ before/after top 5 · refactor applied · endpoint-name list confirming all preserved · import check passes.
>
> Forbidden: renaming endpoints, changing column/label/enum without consistency sweep, premature abstraction (3 similar lines ≠ helper).
>
> Area: «which file / module»

---

## 5. `/squad` — 4-agent pipeline (Architect → Engineer → Reviewer → Optimizer)

> Run a 4-role Workflow on a TMS task. Pipeline: **Architect** (3-step plan + files + risks) → **Engineer** (implement, edit files) → **Reviewer** (senior review — findings only, file:line + severity, CLAUDE.md rules) → **Optimizer** (apply findings + tighten perf/clarity, final diff).
>
> Rules: Reviewer must not rubber-stamp (if clean, say why with evidence); Optimizer keeps behavior identical; final output 4 sections collapsed + recommended commit message.
>
> Task: «the task to run the squad on»

---

## 6. `/ui-build` — Senior frontend build (Bootstrap 5 + Jinja2 component)

> Act as senior frontend engineer on TMS UI. Stack: Jinja2 + Bootstrap 5 + Bootstrap Icons + TomSelect + Chart.js. No React/Vue/build step.
>
> Conventions: pages extend `base.html` via `{% block content %}` + `{% block scripts %}`; every POST form needs `<input type="hidden" name="csrf_token" value="{{ csrf_token() }}">`; tables use `TBL.init('id', pageSize)` and inline handlers need `window.TBL`; never put `transform` on a modal ancestor; live validation = `.is-invalid` + `.invalid-feedback`; never `defer` GSAP on login.
>
> Output: 1. component sketch (ascii + states: loading/empty/error/populated) 2. markup (Jinja2 + BS classes, no inline style >5 props) 3. server contract (route → context → template vars, traced both ways) 4. minimal vanilla JS (extract to `static/js/` if 30+ lines) 5. a11y (keyboard, aria, contrast) 6. cache-bust via `STATIC_VER`.
>
> Verify: import check + boot server + manually test golden path + 1 edge case.
>
> Component: «component or page to build»

---

## 7. `/security-audit` — TMS security audit (codebase-wide depth)

> Act as senior security engineer auditing TMS (deeper than the branch-diff `/security-review`).
>
> Context: username/password + optional TOTP, lockout 5 fails/15min, global CSRFProtect, Flask-Limiter on /login; roles spoc/central/admin (decorators in `tms/decorators.py`); QR public `/q/<token>/*` CSRF-exempt + status-locked; audit chain SHA-256 in `audit_log`; Render behind X-Forwarded-For.
>
> Inspect: AuthN (hashing, 2FA, session fixation, must_change_password) · AuthZ (every route decorated? cross-plant leak? impersonation safe?) · CSRF (every POST tokened? exemptions justified?) · Injection (all `?` params, no f-string SQL) · SSRF/file paths · sensitive-data exposure (logs/flash/errors) · audit-chain integrity (`verify_chain()` reachable/bypassable?) · rate limiting beyond /login · headers (CSP/HSTS/X-Frame/Referrer).
>
> Output: vuln report `| # | Severity (CRIT/HIGH/MED/LOW) | File:Line | What | Attack scenario | Fix |`, grouped by severity; CRIT/HIGH get inline patch; end with "Already-mitigated" list. No speculation without code citation; no defense-in-depth without a real attacker model.
>
> Focus: «optional area»

---

## 8. `/devops-check` — DevOps + deployment audit

> Act as senior DevOps engineer for TMS. Stack: Render Starter, GitHub auto-deploy from `main`, gunicorn, SQLite on Render disk.
>
> Inspect: Procfile (cmd/timeout/workers vs RAM) · health check (`/health` 200 + DB?) · persistent disk (`DATABASE_PATH`, backup) · secrets in Render env not repo (`SECRET_KEY`, `GEMINI_API_KEY`, `CRON_SECRET`) · CI/CD (Actions? else propose lint+import+smoke) · logging (structured, PII-redacted) · monitoring (uptime, slow-endpoint, error alerts) · cron jobs (monthly error email parked 89959ce — activated?) · rollback (last-good SHA, one-command revert) · scaling cliff (~50 concurrent on Starter).
>
> Output: posture GREEN/AMBER/RED w/ reasons · punchlist by blast-radius × ease · top 3 this week · "going to prod" gates. Use Render MCP shortcuts ("render status/logs/deploys/errors/memory").
>
> Focus: «optional — cron / secrets / scaling»

---

## 9. `/tech-lead` — Tech lead decision mode (clarify, challenge, recommend, plan)

> Act as senior tech lead on TMS, 5+ year horizon. Before any code:
>
> **Ask** — up to 3 clarifying questions if scope ambiguous; challenge if the request is likely wrong (cite reason; my go-ahead overrides). **Analyze** — tradeoff table `| Option | Pros | Cons | Risk |` (min 2 options); scaling risk at 10×/100×; maintenance load (handoff in 6 months?); reversibility. **Recommend** — one option + one-sentence why; numbered plan w/ file paths + effort; explicit out-of-scope.
>
> Flag anti-patterns: "just add a flag", "migrate later", "edge case ignore" (zero-defect standard), single-caller abstraction.
>
> Output: Decision → Tradeoff → Plan → Out-of-Scope. No code until I confirm the plan.
>
> Question: «the design question or proposal»

---

*Source of truth: `.claude/commands/*.md`. If a skill changes there, regenerate this file.*
