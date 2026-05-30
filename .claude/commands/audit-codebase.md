---
description: Senior-engineer audit of TMS — architecture, duplication, perf risks, fixes
---

Act as senior engineer joining TMS unfamiliar. Reverse-engineer architecture + data flow first. Then surface issues. Do NOT change functionality. Only upgrade quality, scalability, maintainability.

## TMS context
Flask 3.1 + raw sqlite3 + Jinja2, deployed Render. Modular routes under `tms/routes/`, helpers in `tms/helpers.py`, schema `schema.sql`. Read `CLAUDE.md` first — it documents conventions you MUST respect (programme_master gatekeeper, audience-from-TNI, source enum, etc).

## Output (in order)

1. **Architecture breakdown** — diagram-in-words of: request → route → helper → SQL → template. Note where it deviates from the modular ideal.
2. **Findings table** — `| # | Category | File:Line | Severity | What | Why it matters |`
   Categories: bad-arch · duplicate-logic · perf-bottleneck · scalability-risk · maintainability · audit-trail-gap · security
3. **Refactor proposals** — for top 5 findings: concrete diff sketch (file path + before/after snippet). Respect existing TMS-Wide Consistency Rule.
4. **Quick wins** — items < 30 min each, list only.

## Rules
- Do NOT rewrite without listing the finding first.
- Cite file:line, not vague references.
- If user wants to apply: they'll say "apply 1,3,5" → only then edit.
- Run `python -c "import app; print('OK')"` before declaring done if you edited anything.

$ARGUMENTS  # optional: focus area, e.g. "calendar module" or "TNI pipeline"
