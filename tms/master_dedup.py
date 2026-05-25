"""Programme master duplicate detector + merger.

Finds near-identical rows in programme_master so Central can merge them.

Detection
---------
- Pairwise similarity via rapidfuzz token_set_ratio
- Pre-normalisation:
    * data_hygiene.normalise() — unicode / whitespace / punctuation
    * & ↔ and equivalence
    * lowercase, strip
- Clustering: greedy union-find (each row joins first cluster it matches)
- Default threshold: 0.85 (catches "Temperature & PH" vs "Temperature and PH")

Merge
-----
- Central picks one canonical row per cluster.
- UPDATE programme_master SET name = canonical WHERE id IN (loser_ids)
  is replaced by a real delete-and-rename:
    1. Cascade rename in tni / calendar / programme_details / emp_training
       (case-insensitive match on programme_name = any loser name)
    2. DELETE FROM programme_master WHERE id IN (loser_ids)
    3. UPDATE programme_master SET name = canonical_name WHERE id = winner_id
- Single transaction.
- Audit log entry per cluster.
"""

import re

try:
    from rapidfuzz import fuzz
    HAS_RAPIDFUZZ = True
except ImportError:
    HAS_RAPIDFUZZ = False
    from difflib import SequenceMatcher

from tms.data_hygiene import normalise


_AND_RE = re.compile(r'\s*&\s*|\s+and\s+', re.IGNORECASE)


def _dedup_key(name):
    """Canonicalise a programme name for similarity comparison.

    - normalise() for unicode / whitespace / trailing notes
    - & / and / + treated identically
    - punctuation removed
    - lowercase
    """
    if not name:
        return ''
    s = normalise(name)
    s = _AND_RE.sub(' and ', s)
    s = re.sub(r'[^\w\s]', ' ', s)
    s = re.sub(r'\s+', ' ', s).strip().lower()
    return s


def _similarity(a, b):
    """Conservative similarity for duplicate detection.

    Uses fuzz.ratio (full Levenshtein) — does NOT inflate when one string
    is a strict subset of the other. token_set_ratio is too generous here
    because every "X Maintenance" looks similar to every "Y Maintenance".
    """
    if HAS_RAPIDFUZZ:
        return fuzz.ratio(a, b) / 100.0
    return SequenceMatcher(None, a, b).ratio()


def _length_similar(a, b, min_ratio=0.6):
    """Reject pairs with very different lengths regardless of fuzz score."""
    la, lb = len(a), len(b)
    if not la or not lb:
        return False
    return min(la, lb) / max(la, lb) >= min_ratio


def find_duplicates(plant_id, db, threshold=0.85):
    """Return list of clusters. Each cluster is a list of dicts:
        {id, name, key, usage_count, tni_count, cal_count, pd_count, et_count}
    Only clusters with len >= 2 are returned.
    """
    rows = db.execute(
        'SELECT id, name FROM programme_master WHERE plant_id=? ORDER BY name',
        (plant_id,)
    ).fetchall()
    if len(rows) < 2:
        return []

    items = []
    for r in rows:
        items.append({
            'id':   r['id'],
            'name': r['name'],
            'key':  _dedup_key(r['name']),
        })

    # Greedy clustering. A row joins a cluster only if BOTH:
    #   1. fuzz.ratio >= 0.92    (very strong)
    #   OR
    #   2. fuzz.ratio >= threshold AND token-difference <= 1
    #      (catches typo / plural / hyphen variants without lumping
    #      "VFD Operation & Maintenance" with "Boiler Operation & Maintenance")
    #   3. One key is a strict substring of the other (catches
    #      "Control Process Losses" vs "How to Control Process Losses")
    clusters = []  # list[list[item]]
    for it in items:
        if not it['key']:
            continue
        placed = False
        for c in clusters:
            rep = c[0]['key']
            if not _length_similar(it['key'], rep):
                continue
            score = _similarity(it['key'], rep)
            it_tokens  = set(it['key'].split())
            rep_tokens = set(rep.split())
            token_diff = len(it_tokens.symmetric_difference(rep_tokens))
            is_substring = (it['key'] in rep) or (rep in it['key'])
            if score >= 0.92 or \
               (score >= threshold and token_diff <= 1) or \
               (score >= threshold and is_substring):
                c.append(it)
                placed = True
                break
        if not placed:
            clusters.append([it])

    # Keep only clusters with duplicates
    dupes = [c for c in clusters if len(c) >= 2]

    # Decorate with usage counts so Central can pick the canonical name
    # by "highest usage" if they want.
    for c in dupes:
        for it in c:
            tni_n = db.execute(
                'SELECT COUNT(*) FROM tni WHERE plant_id=? AND lower(programme_name)=lower(?)',
                (plant_id, it['name'])).fetchone()[0]
            cal_n = db.execute(
                'SELECT COUNT(*) FROM calendar WHERE plant_id=? AND lower(programme_name)=lower(?)',
                (plant_id, it['name'])).fetchone()[0]
            pd_n  = db.execute(
                'SELECT COUNT(*) FROM programme_details WHERE plant_id=? AND lower(programme_name)=lower(?)',
                (plant_id, it['name'])).fetchone()[0]
            et_n  = db.execute(
                'SELECT COUNT(*) FROM emp_training WHERE plant_id=? AND lower(programme_name)=lower(?)',
                (plant_id, it['name'])).fetchone()[0]
            it['tni_count'] = tni_n
            it['cal_count'] = cal_n
            it['pd_count']  = pd_n
            it['et_count']  = et_n
            it['usage_count'] = tni_n + cal_n + pd_n + et_n

    # Sort each cluster: highest usage first (suggested canonical)
    for c in dupes:
        c.sort(key=lambda x: -x['usage_count'])

    # Sort clusters by total cluster usage (most-impactful first)
    dupes.sort(key=lambda c: -sum(x['usage_count'] for x in c))

    return dupes


def merge_cluster(plant_id, winner_id, loser_ids, canonical_name, db, audit_log_fn=None):
    """Merge loser rows into winner. Returns dict of update counts.

    Caller is responsible for transaction commit.
    """
    # Get all names (winner + losers) for cascade
    all_ids = [winner_id] + list(loser_ids)
    rows = db.execute(
        f'SELECT id, name FROM programme_master WHERE id IN ({",".join("?"*len(all_ids))})',
        all_ids
    ).fetchall()
    name_map = {r['id']: r['name'] for r in rows}
    winner_name_orig = name_map.get(winner_id)
    loser_names = [name_map[i] for i in loser_ids if i in name_map]
    all_loser_lower = [n.lower() for n in loser_names + [winner_name_orig]
                       if n.lower() != canonical_name.lower()]

    counts = {'tni': 0, 'calendar': 0, 'programme_details': 0, 'emp_training': 0,
              'winner_renamed': 0, 'losers_deleted': 0}

    if all_loser_lower:
        placeholders = ','.join('?' * len(all_loser_lower))
        # Cascade rename in transactional tables
        for table in ('tni', 'calendar', 'programme_details', 'emp_training'):
            cur = db.execute(
                f'UPDATE {table} SET programme_name=? '
                f'WHERE plant_id=? AND lower(programme_name) IN ({placeholders})',
                [canonical_name, plant_id] + all_loser_lower
            )
            counts[table] = cur.rowcount

    # Delete losers from master
    if loser_ids:
        db.execute(
            f'DELETE FROM programme_master WHERE id IN ({",".join("?"*len(loser_ids))})',
            list(loser_ids)
        )
        counts['losers_deleted'] = len(loser_ids)

    # Rename winner to canonical if changed
    if winner_name_orig and winner_name_orig != canonical_name:
        db.execute(
            'UPDATE programme_master SET name=? WHERE id=?',
            (canonical_name, winner_id)
        )
        counts['winner_renamed'] = 1

    if audit_log_fn:
        try:
            audit_log_fn(
                'BULK_DELETE',
                f'master_dedup:winner={canonical_name};losers={loser_names};'
                f'cascaded=tni={counts["tni"]},cal={counts["calendar"]},'
                f'pd={counts["programme_details"]},et={counts["emp_training"]}'
            )
        except Exception:
            pass

    return counts
