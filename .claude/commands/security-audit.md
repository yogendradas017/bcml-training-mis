---
description: TMS-targeted security audit (deeper than built-in /security-review)
---

Act as senior security engineer auditing TMS. (Note: `/security-review` built-in covers branch diff; this command does codebase-wide depth.)

## Scope — TMS specifics
- Auth: username/password + optional TOTP, account lockout (5 fails / 15min), CSRFProtect global, Flask-Limiter on /login.
- Roles: spoc / central / admin. Decorators in `tms/decorators.py`.
- QR public endpoints (`/q/<token>/*`) — CSRF-exempt. PIN gate optional. Status-lock now enforced (To Be Planned only).
- Audit chain: SHA-256 row_hash + payload_hash in `audit_log` (see `tms/audit.py`).
- Render env: behind X-Forwarded-For; production secrets in Render env vars; SECRET_KEY default is dev-only.

## Inspect
1. **AuthN** — login flow, password hashing (werkzeug), 2FA enforcement, session fixation, must_change_password gate.
2. **AuthZ** — every route has a decorator? Cross-plant data leakage? Admin impersonation safe?
3. **CSRF** — every POST form has token? Exemptions justified (only QR public)?
4. **Injection** — all SQL uses `?` parameterisation? Any f-string SQL?
5. **SSRF / file paths** — uploads validated, paths normalised?
6. **Sensitive data exposure** — logs, flash messages, error pages leaking PII?
7. **Audit chain integrity** — `verify_chain()` actually called somewhere? Can it be bypassed?
8. **Rate limiting** — only /login? Other endpoints brute-forceable?
9. **Headers** — CSP, HSTS, X-Frame-Options, Referrer-Policy.

## Output
- **Vulnerability report** — `| # | Severity (CRIT/HIGH/MED/LOW) | File:Line | What | Attack scenario | Fix |`
- Group by severity, highest first.
- Critical / High → propose patch inline.
- End with "Already-mitigated" list (don't re-flag known-handled items).

## Forbidden
- Speculation without code citation.
- "Defense in depth" findings that don't have a real attacker model.

$ARGUMENTS  # optional focus area
