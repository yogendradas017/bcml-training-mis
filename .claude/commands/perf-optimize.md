---
description: Performance pass — identify bottlenecks, fix, measure
---

Act as senior performance engineer on TMS. Target: fits Render Starter (512 MB RAM, 1 CPU). Past incidents: pandas OOM'd worker on TNI analyze (fixed via openpyxl streaming).

## Investigate

1. **Hot paths** — list top 5 endpoints by likely traffic (Dashboard, /tni/data, /calendar, /training, /api/dashboard-monthly). For each, note:
   - SQL query shape (N+1? full table scan? missing index?)
   - Memory profile (pandas? large dict in template?)
   - Rendering cost (DOM size? duplicated TomSelect init?)
2. **Bottleneck table** — `| File:Line | Cost | Fix | Effort |`
3. **Memory leaks / unbounded growth** — long-lived caches, accumulating dicts, large session payloads, leaked DB connections.

## Apply standard plays
- Index missing — add via `_ensure_indexes` in `tms/db.py`, idempotent.
- Pandas → openpyxl streaming for any upload > 1k rows.
- Materialised joins → `EXISTS()` subquery for completed-status checks (pattern already used).
- Template loops over 1000+ rows → paginate or virtualise (TNI pattern, 30/page debounced search).
- `SELECT *` followed by template using 3 cols → narrow the SELECT.

## Verify
- `python -c "import app; print('OK')"`.
- If touching SQL: explain query, confirm uses index.
- If touching template: confirm browser DOM count unchanged or smaller.

## Output
- Bottleneck table · top 3 fixes applied · before/after metric estimate · what was deliberately NOT changed and why.

$ARGUMENTS  # optional: specific endpoint or feature
