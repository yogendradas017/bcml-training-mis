---
description: Senior frontend build — Bootstrap 5 + Jinja2 reusable component
---

Act as senior frontend engineer on TMS UI. Stack: Jinja2 + Bootstrap 5 + Bootstrap Icons + TomSelect + Chart.js. No React/Vue/build step.

## TMS UI conventions
- `templates/base.html` is the shell — sidebar + topbar. Every page extends it via `{% block content %}` + `{% block scripts %}`.
- Forms POST to a Flask endpoint. CSRF token hidden field MANDATORY: `<input type="hidden" name="csrf_token" value="{{ csrf_token() }}">` (see CLAUDE.md).
- Tables use `TBL.init('id', pageSize)` from main.js. Inline `oninput`/`onchange` handlers expect `window.TBL` global ([[feedback-tbl-global]]).
- Modals: avoid `transform` on any ancestor or modal traps behind backdrop ([[feedback-bootstrap-modal-stacking]]).
- Live validation: red border via `.is-invalid` + `<div class="invalid-feedback">` (see Add Calendar form).
- GSAP scripts on login.html — NEVER add `defer`.

## Output
1. **Component sketch** — ascii layout + states (loading, empty, error, populated).
2. **Markup** — Jinja2 + Bootstrap classes. No inline styles unless < 5 properties.
3. **Server contract** — route → context vars → template vars (trace both directions per CLAUDE.md zero-defect rule).
4. **JS** — minimal vanilla. If 30+ lines, extract to `static/js/<name>.js`.
5. **CSP / a11y** — keyboard nav, aria labels, contrast.
6. **Cache-bust** — confirm static assets use `STATIC_VER` query param.

## Verify
- Run `python -c "import app; print('OK')"`.
- Boot server, manually test golden + 1 edge case (CLAUDE.md UI rule: "type checking verifies code correctness, not feature correctness").

$ARGUMENTS  # component or page to build
