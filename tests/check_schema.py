# -*- coding: utf-8 -*-
"""Static check: schema.sql (fresh-deploy path) vs the fully-migrated demo DB
(existing-DB path). A column present in ONE but not the other = the
SCHEMA_MIGRATION drift CLAUDE.md warns about. Prod-safe, no agents."""
import sqlite3, os, sys
try: sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception: pass
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def cols_of(conn):
    out = {}
    for (t,) in conn.execute("SELECT name FROM sqlite_master WHERE type='table'"):
        if t.startswith('sqlite_'):
            continue
        out[t] = {r[1] for r in conn.execute('PRAGMA table_info(%s)' % t)}
    return out


# schema.sql -> fresh in-memory DB
fresh = sqlite3.connect(':memory:')
fresh.executescript(open(os.path.join(ROOT, 'schema.sql'), encoding='utf-8').read())
schema_cols = cols_of(fresh)

# migrated reference
demo = sqlite3.connect(os.path.join(ROOT, 'data', 'demo.db'))
demo_cols = cols_of(demo)

print('schema.sql tables: %d | demo.db tables: %d' % (len(schema_cols), len(demo_cols)))
gaps = []
migration_only = []
for t in sorted(set(schema_cols) | set(demo_cols)):
    s = schema_cols.get(t, set())
    d = demo_cols.get(t, set())
    if t not in demo_cols:
        print('  TABLE only in schema.sql (never created in a real DB?): %s' % t); continue
    if t not in schema_cols:
        print('  TABLE only in demo.db (missing from schema.sql -> fresh deploy lacks it!): %s' % t)
        gaps.append((t, '<whole table>')); continue
    only_schema = s - d   # in schema.sql CREATE TABLE but missing from a migrated DB -> existing DBs lack it
    only_demo = d - s     # in DB via migration but not in schema.sql -> fresh deploy relies on migration running
    for c in sorted(only_schema):
        gaps.append((t, c))
    for c in sorted(only_demo):
        migration_only.append((t, c))

print('\n=== MIGRATION GAPS (in schema.sql but NOT in a migrated DB -> existing DBs broken): %d ===' % len(gaps))
for t, c in gaps:
    print('  %s.%s' % (t, c))
print('\n=== MIGRATION-ONLY columns (in DB but NOT in schema.sql -> fresh deploy needs the migration to run): %d ===' % len(migration_only))
for t, c in migration_only:
    print('  %s.%s' % (t, c))
print('\nNote: migration-only is OK IF init_db runs that ALTER every startup; review each against db.py _migrate_*.')
sys.exit(len(gaps))
