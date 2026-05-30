---
description: DevOps + deployment audit — Render config, CI/CD, monitoring, rollout
---

Act as senior DevOps engineer for TMS. Stack: Render Starter ($7/mo), GitHub auto-deploy from `main`, gunicorn, SQLite on Render disk.

## Inspect

1. **Procfile** — start command sane? Timeout high enough? Workers count vs RAM?
2. **Health check** — does `/health` exist? Returns 200 + DB check? Render configured to use it?
3. **Persistent disk** — mounted? `DATABASE_PATH` env set? Backup strategy?
4. **Secrets** — `SECRET_KEY`, `GEMINI_API_KEY`, `CRON_SECRET` all in Render env vars (NOT in repo)? Use `mcp__render__update_environment_variables` to check.
5. **CI/CD** — any GitHub Actions? Pre-merge tests? If none, propose minimal (lint + import-check + smoke).
6. **Logging** — structured? PII redacted? Render log retention sufficient?
7. **Monitoring** — uptime ping? Slow-endpoint alerts? Error rate alerting?
8. **Cron jobs** — Render cron list (`mcp__render__list_workspaces`, then cron). Monthly error email parked (89959ce) — activated?
9. **Rollback** — last good deploy SHA documented? One-command revert path?
10. **Scaling cliff** — at what user count does Starter break? (~50 concurrent users typical).

## Output
- Deployment posture: GREEN / AMBER / RED with reasons.
- Punchlist sorted by blast-radius × ease.
- Top 3 to fix this week.
- "Going to prod" gates list (see `project_prod_rollout_checklist.md` if user asks).

## Use Render MCP shortcuts
Per `reference_render_shortcuts.md`: "render status / logs / deploys / errors / memory" → call MCP directly.

$ARGUMENTS  # optional focus (cron / secrets / scaling / etc)
