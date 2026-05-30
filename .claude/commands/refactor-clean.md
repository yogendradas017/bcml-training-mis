---
description: Refactor messy code to clean architecture — same behavior, better shape
---

Act as senior architect refactoring TMS. **Do NOT change product behavior.** Improve only architecture + code quality.

## Constraints
- TMS uses thin `app.py` + `tms/routes/<domain>.py` registering routes directly (NOT Blueprints). Templates call `url_for('endpoint_name')` — endpoint names MUST stay stable.
- Raw `sqlite3` — no ORM. Don't introduce one without explicit user approval.
- `tms/helpers.py` is ~2000 lines; splitting it is fine, but every import path stays valid.
- TMS-Wide Consistency Rule (CLAUDE.md): any enum / label / source change must propagate across ALL templates + routes.

## Plays in order

1. **Separate concerns** — fat routes containing SQL + business logic + flash strings → extract pure functions to helpers.
2. **Reduce duplication** — same SQL in 3 routes → one helper. Same validation in add/edit/bulk → one validator (see `validate_calendar_row` pattern).
3. **Tighten coupling** — long arg lists (`openEdit(id, prog, type, source, …, status)`) → data-attribute dispatch (see Tier 7 edit refactor).
4. **Clarify state machine** — if logic threads through 3+ if-status branches, lift to a small dict or table.

## Output
- New folder/file structure (only if splitting).
- Findings table with before/after snippet for top 5.
- Refactored code applied.
- Endpoint name list — confirm all preserved.
- `python -c "import app; print('OK')"` passes.

## Forbidden
- Renaming Flask endpoints (breaks template `url_for`).
- Changing DB column names / labels / enums without consistency sweep.
- Introducing premature abstractions (3 similar lines > premature helper).

$ARGUMENTS  # which file/module/area
