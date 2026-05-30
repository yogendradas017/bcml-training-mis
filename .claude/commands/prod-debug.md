---
description: Production-grade debug — trace, root-cause, fix, verify
---

Act as senior debugging engineer investigating a TMS live production issue. Render is the live env. Do NOT guess. Trace first.

## Steps

1. **Restate the symptom** — exactly what user reported.
2. **Reproduce locally** if possible. If not, state why (data-dependent, prod-only, etc).
3. **Trace** — Read relevant route → helper → SQL → template. Identify where reality diverges from intent.
4. **Root cause** — one-sentence statement. Distinguish from symptom.
5. **Edge cases** — list 3+ inputs/states that could trigger same / related failure.
6. **Fix** — propose patch with file:line. Include WHY each line changes.
7. **Verify** — run `python -c "import app; print('OK')"`. If schema or DB touched, describe migration path. If timezone, files, or external resource — call out Render-vs-local difference (see CLAUDE.md zero-defect rule).
8. **Regression guard** — what tests/checks would have caught this. Add the check if cheap.

## Forbidden
- Bypassing CSRF / 2FA / hooks to "make it pass".
- `--no-verify` style shortcuts.
- Fixing the symptom (e.g. catching exception) without root-cause statement.

## Output
- Symptom · Reproduce · Trace (file:line walk) · Root Cause · Edge Cases · Patch · Verify · Guard.

$ARGUMENTS  # describe the bug / paste the error / point to the screenshot
