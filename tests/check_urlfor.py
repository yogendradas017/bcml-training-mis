# -*- coding: utf-8 -*-
"""Static check: every url_for('endpoint') used in a template must resolve to a
real Flask endpoint. Catches typos that only 500 on a conditional/POST branch
the smoke harness's GET pass wouldn't render. Prod-safe, no agents."""
import os, re, sys, glob
os.environ.setdefault('DATABASE_PATH', 'data/demo.db')
os.environ.setdefault('SECRET_KEY', 'smoke-test-key')
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
try: sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception: pass

import app as appmod
endpoints = set(appmod.app.view_functions.keys())

pat = re.compile(r"""url_for\(\s*['"]([^'"]+)['"]""")
bad = []
n = 0
for path in glob.glob('templates/**/*.html', recursive=True):
    txt = open(path, encoding='utf-8').read()
    for m in pat.finditer(txt):
        ep = m.group(1)
        n += 1
        if ep not in endpoints:
            line = txt[:m.start()].count('\n') + 1
            bad.append((path, line, ep))

print('checked %d url_for() calls across templates' % n)
print('=== UNRESOLVED url_for endpoints: %d ===' % len(bad))
for path, line, ep in bad:
    print('  %s:%d -> url_for(%r) : NO SUCH ENDPOINT' % (path, line, ep))
sys.exit(len(bad))
