"""Data hygiene engine — multi-layer detector for TNI / programme garbage.

Four layers run in sequence:

  Layer 1  normalise()           — deterministic cleanup (unicode, whitespace,
                                   quotes, control chars, trailing notes,
                                   duration suffixes, slash spacing)
  Layer 2  validate()            — reject obvious garbage (placeholder tokens,
                                   too short, no letters, paragraph-length)
  Layer 3  progressive_match()   — exact → word-order → abbr-expand →
                                   rapidfuzz token_set → phonetic metaphone
  Layer 4  suggest_top_n()       — always surface top-N candidates so the
                                   SPOC can pick even when no layer matched

The public entry point is `analyze_programme_name(raw, master_list)` which
returns a single dict the caller can fold into the existing row schema:

    {
      'normalised':   '<cleaned raw>',
      'valid':        bool,
      'reject_reason': str | None,
      'best_match':   str | None,     # canonical name if confident enough
      'confidence':   float,          # 0..1
      'method':       str,            # 'exact' | 'word-order' | 'abbr-expand'
                                      # | 'levenshtein-strong' | 'levenshtein-medium'
                                      # | 'phonetic' | 'synonym' | 'none'
      'suggestions':  [(name, score, method), ...]   # top 5
      'garbage_class': str | None,    # 'whitespace' | 'case' | 'punctuation'
                                      # | 'unicode' | 'placeholder' | 'typo'
                                      # | 'abbreviation' | 'word-order' | 'unknown'
    }

Also exposes per-field helpers for prog_type / mode / audience / emp_code so
the same engine validates every enum / lookup in the TNI workflow.
"""

import re
import unicodedata

try:
    from rapidfuzz import fuzz
    HAS_RAPIDFUZZ = True
except ImportError:
    HAS_RAPIDFUZZ = False
    from difflib import SequenceMatcher

try:
    from metaphone import doublemetaphone
    HAS_METAPHONE = True
except ImportError:
    HAS_METAPHONE = False

try:
    from unidecode import unidecode
    HAS_UNIDECODE = True
except ImportError:
    HAS_UNIDECODE = False


# ── Layer 1 — Normaliser ──────────────────────────────────────────────────────

_QUOTE_TRANSLATE = str.maketrans({
    '“': '"', '”': '"',          # curly double quotes
    '‘': "'", '’': "'",          # curly single quotes
    '–': '-', '—': '-',          # en-dash, em-dash
    ' ': ' ',                         # NBSP
    '​': '',  '﻿': '',           # zero-width / BOM
})

_TRAILING_PAREN_RE   = re.compile(r'\s*\([^)]*\)\s*$')
_TRAILING_BRACKET_RE = re.compile(r'\s*\[[^\]]*\]\s*$')
_TRAILING_DUR_RE     = re.compile(r'\s+\d+\s*(hr|hrs|hour|hours|min|mins|day|days|week|weeks)$', re.IGNORECASE)
_TRAILING_YEAR_RE    = re.compile(r'\s+(20\d{2}|19\d{2})\s*$')
_MULTI_WS_RE         = re.compile(r'\s+')
_SLASH_SPACE_RE      = re.compile(r'\s*/\s*')
_DASH_SPACE_RE       = re.compile(r'\s*-\s*')


def normalise(raw):
    """Deterministic cleanup. Returns cleaned string ready for matching."""
    if raw is None:
        return ''
    s = str(raw)
    # 1. Strip control / format chars (zero-width, BOM, etc.)
    s = ''.join(c for c in s if unicodedata.category(c)[0] != 'C')
    # 2. NFKC: fullwidth → ASCII, ligatures expand, etc.
    s = unicodedata.normalize('NFKC', s)
    # 3. Curly quotes / dashes / NBSP → ASCII
    s = s.translate(_QUOTE_TRANSLATE)
    # 4. Strip wrapping quotes
    s = s.strip().strip('"').strip("'").strip()
    # 5. Strip trailing parenthetical notes (Fire Safety (urgent) → Fire Safety)
    s = _TRAILING_PAREN_RE.sub('', s)
    s = _TRAILING_BRACKET_RE.sub('', s)
    # 6. Strip trailing duration marker (Fire Safety 8hrs → Fire Safety)
    s = _TRAILING_DUR_RE.sub('', s)
    # 7. Strip trailing year (5S Audit 2024 → 5S Audit)
    s = _TRAILING_YEAR_RE.sub('', s)
    # 8. Standardise slash + dash spacing
    s = _SLASH_SPACE_RE.sub('/', s)
    # 9. Collapse whitespace
    s = _MULTI_WS_RE.sub(' ', s).strip()
    # 10. Strip leading/trailing punctuation
    s = s.strip('-:/_.,; ')
    return s


# ── Layer 2 — Validator ───────────────────────────────────────────────────────

GARBAGE_TOKENS = {
    'na', 'n/a', 'nil', '-', '--', '---', '?', '??', '???',
    'tbd', 'todo', 'xxx', 'test', 'na.', 'none', 'null',
    'pending', 'unknown', 'others', 'other', 'misc',
}


def validate(s):
    """Returns (is_valid, reject_reason). reject_reason is None when valid."""
    if not s:
        return False, 'empty'
    s_low = s.lower().strip()
    if s_low in GARBAGE_TOKENS:
        return False, 'placeholder'
    if len(s) < 3:
        return False, 'too short'
    if not re.search(r'[A-Za-z]', s):
        return False, 'no letters'
    digit_ratio = sum(1 for c in s if c.isdigit()) / max(len(s), 1)
    if digit_ratio > 0.6:
        return False, 'mostly digits'
    if re.match(r'^[^A-Za-z0-9]+$', s):
        return False, 'symbols only'
    if len(s) > 200:
        return False, 'too long (paragraph?)'
    return True, None


# ── Layer 3 — Progressive matcher ─────────────────────────────────────────────
#
# Abbreviation map and synonyms are intentionally conservative — only well-known
# domain expansions go here. Add more over time as Central observes new typos.

ABBR_MAP = {
    'ehs':       'environment health safety',
    'sop':       'standard operating procedure',
    'sops':      'standard operating procedures',
    'ojt':       'on the job training',
    'qms':       'quality management system',
    'ems':       'environment management system',
    'mgmt':      'management',
    'mgt':       'management',
    'pdm':       'predictive maintenance',
    'pm':        'preventive maintenance',
    'cm':        'corrective maintenance',
    'tpm':       'total productive maintenance',
    'r&m':       'repair and maintenance',
    'safty':     'safety',
    'saftey':    'safety',
    'procidure': 'procedure',
    'trainning': 'training',
    'maintainance': 'maintenance',
    'mantenance':   'maintenance',
    'opr':       'operation',
    'opn':       'operation',
}

SYNONYMS = {
    'bagasse handling':  ['bagasse mgmt', 'bagasse storage', 'bagasse management'],
    'cane handling':     ['cane management', 'cane operations'],
    'fire safety':       ['fire fighting', 'fire prevention', 'firefighting'],
    'first aid':         ['first-aid', 'firstaid'],
    'work permit':       ['permit to work', 'ptw'],
}

# Build reverse map for synonym lookup
_REVERSE_SYNONYMS = {}
for canonical, alts in SYNONYMS.items():
    for alt in alts:
        _REVERSE_SYNONYMS[alt.lower()] = canonical


_TOKEN_SPLIT_RE = re.compile(r'[/\-_\s]+')


def _tokens_sorted(s):
    return tuple(sorted(t for t in _TOKEN_SPLIT_RE.split(s.lower()) if t))


def _expand_abbr(s_lower):
    parts = re.split(r'(\s+)', s_lower)
    out = []
    for p in parts:
        if p.strip() and p.lower() in ABBR_MAP:
            out.append(ABBR_MAP[p.lower()])
        else:
            out.append(p)
    return ''.join(out)


def _fuzzy_score(a, b):
    """0..1 weighted ratio. Uses rapidfuzz WRatio (combines Levenshtein,
    partial, and token-based) for balanced scoring. Falls back to
    SequenceMatcher when rapidfuzz unavailable."""
    if HAS_RAPIDFUZZ:
        # WRatio = weighted combination; less prone to false 1.0 from token_set
        # when one string is a strict subset of the other.
        return fuzz.WRatio(a, b) / 100.0
    return SequenceMatcher(None, a, b).ratio()


def _substring_bonus(short, long_str):
    """If short string (>=3 chars) is a token in long_str, return a confidence
    boost. Catches cases like 'EHS' → 'EHS/HR Training'."""
    if len(short) < 3:
        return 0.0
    short_l = short.lower()
    tokens = set(t.lower() for t in _TOKEN_SPLIT_RE.split(long_str) if t)
    if short_l in tokens:
        return 0.85   # confident but not exact
    return 0.0


def _phonetic_key(s):
    if not HAS_METAPHONE:
        return ''
    primary, _ = doublemetaphone(s)
    return primary or ''


def progressive_match(raw, master_list):
    """Return (best_match, confidence, method).

    method ∈ {'exact', 'word-order', 'abbr-expand', 'levenshtein-strong',
              'levenshtein-medium', 'phonetic', 'synonym', 'none'}
    confidence:   1.00 exact, 0.95 word-order, 0.92 abbr, 0.90+ strong fuzzy,
                  0.75-0.90 medium fuzzy, 0.85 phonetic, 0.88 synonym.
    """
    if not raw or not master_list:
        return None, 0.0, 'none'

    s = normalise(raw)
    sl = s.lower()
    master_lower = [m.lower() for m in master_list]

    # 1. Exact (case-insensitive after normalise)
    if sl in master_lower:
        return master_list[master_lower.index(sl)], 1.00, 'exact'

    # 2. Word-order / slash-order invariant
    s_tokens = _tokens_sorted(sl)
    for m, ml in zip(master_list, master_lower):
        if _tokens_sorted(ml) == s_tokens:
            return m, 0.95, 'word-order'

    # 3. Synonym dictionary
    if sl in _REVERSE_SYNONYMS:
        canonical = _REVERSE_SYNONYMS[sl]
        for m, ml in zip(master_list, master_lower):
            if ml == canonical:
                return m, 0.88, 'synonym'

    # 4. Expand abbreviations + retry exact + fuzzy
    s_expanded = _expand_abbr(sl)
    if s_expanded != sl:
        if s_expanded in master_lower:
            return master_list[master_lower.index(s_expanded)], 0.92, 'abbr-expand'
        # Also try fuzzy after expansion
        for m, ml in zip(master_list, master_lower):
            if _fuzzy_score(s_expanded, ml) >= 0.90:
                return m, 0.90, 'abbr-expand'

    # 5. Rapidfuzz weighted ratio + substring-token bonus
    best = None
    best_score = 0.0
    best_method = 'levenshtein-strong'
    for m, ml in zip(master_list, master_lower):
        score = _fuzzy_score(sl, ml)
        # Boost if normalised text is a complete token in master name
        sub_score = _substring_bonus(sl, m)
        if sub_score > score:
            score = sub_score
        if score > best_score:
            best, best_score = m, score

    if best_score >= 0.90:
        return best, best_score, best_method

    # 6. Phonetic (catches Saftey → Safety even if fuzzy missed)
    if HAS_METAPHONE:
        s_phon = _phonetic_key(sl)
        if s_phon:
            for m, ml in zip(master_list, master_lower):
                m_phon = _phonetic_key(ml)
                if m_phon and m_phon == s_phon:
                    return m, 0.85, 'phonetic'

    # 7. Medium fuzzy — return as suggestion-grade
    if best_score >= 0.75:
        return best, best_score, 'levenshtein-medium'

    return best, best_score, 'none'


# ── Layer 4 — Suggester ───────────────────────────────────────────────────────

def suggest_top_n(raw, master_list, n=5, min_score=0.40):
    """Always return top-N master candidates ranked by similarity.

    Returned list: [(master_name, score, method_hint), ...]
    Methods are hints only — for UI display.
    """
    if not raw or not master_list:
        return []
    s = normalise(raw).lower()
    s_expanded = _expand_abbr(s)
    s_tokens = _tokens_sorted(s)
    s_phon = _phonetic_key(s) if HAS_METAPHONE else ''

    scored = []
    for m in master_list:
        ml = m.lower()
        # Combine multiple signals, take max
        s_fuzzy = _fuzzy_score(s, ml)
        s_expanded_fuzzy = _fuzzy_score(s_expanded, ml) if s_expanded != s else 0
        score = max(s_fuzzy, s_expanded_fuzzy)
        method = 'fuzzy'
        if _tokens_sorted(ml) == s_tokens and score < 0.95:
            score = max(score, 0.95); method = 'word-order'
        if HAS_METAPHONE and s_phon and _phonetic_key(ml) == s_phon and score < 0.85:
            score = max(score, 0.85); method = 'phonetic'
        if score >= min_score:
            scored.append((m, round(score, 3), method))

    scored.sort(key=lambda x: -x[1])
    return scored[:n]


# ── Public entry point ───────────────────────────────────────────────────────

def analyze_programme_name(raw, master_list):
    """Composite hygiene call for one programme name. Returns the dict
    documented in the module docstring."""
    out = {
        'normalised':    '',
        'valid':         False,
        'reject_reason': None,
        'best_match':    None,
        'confidence':    0.0,
        'method':        'none',
        'suggestions':   [],
        'garbage_class': None,
    }
    if raw is None or str(raw).strip() == '':
        out['reject_reason'] = 'empty'
        out['garbage_class'] = 'empty'
        return out

    out['normalised'] = normalise(raw)
    raw_str = str(raw)

    # Detect garbage class from raw → normalised diff
    if raw_str != out['normalised']:
        if raw_str.strip() != raw_str:
            out['garbage_class'] = 'whitespace'
        elif _MULTI_WS_RE.search(raw_str.strip()):
            out['garbage_class'] = 'whitespace'
        elif raw_str.lower() != out['normalised'].lower():
            out['garbage_class'] = 'unicode'
        else:
            out['garbage_class'] = 'punctuation'

    ok, reason = validate(out['normalised'])
    out['valid'] = ok
    if not ok:
        out['reject_reason'] = reason
        out['garbage_class'] = 'placeholder' if reason == 'placeholder' else (out['garbage_class'] or reason)
        return out

    best, conf, method = progressive_match(out['normalised'], master_list)
    out['best_match'] = best
    out['confidence'] = round(conf, 3)
    out['method']     = method
    if method == 'abbr-expand' and not out['garbage_class']:
        out['garbage_class'] = 'abbreviation'
    elif method == 'word-order' and not out['garbage_class']:
        out['garbage_class'] = 'word-order'
    elif method in ('levenshtein-medium', 'phonetic') and not out['garbage_class']:
        out['garbage_class'] = 'typo'

    # Always include top-5 suggestions when below 1.00 — gives SPOC choice.
    if conf < 1.00:
        out['suggestions'] = suggest_top_n(out['normalised'], master_list, n=5)

    return out


# ── Per-field enum helpers ───────────────────────────────────────────────────

_PROG_TYPE_ABBR = {
    'tec': 'Technical', 'tech': 'Technical', 'technical': 'Technical',
    'beh': 'Behavioural/Leadership', 'behavioural': 'Behavioural/Leadership',
    'leadership': 'Behavioural/Leadership', 'beh/lead': 'Behavioural/Leadership',
    'comm': 'Commercial', 'commercial': 'Commercial',
    'ehs': 'EHS/HR', 'hr': 'EHS/HR', 'ehs/hr': 'EHS/HR', 'hr/ehs': 'EHS/HR',
    'it': 'IT', 'i.t.': 'IT', 'it.': 'IT',
    'cane': 'Cane',
}

_MODE_ABBR = {
    'class': 'Classroom', 'classroom': 'Classroom', 'cls': 'Classroom',
    'offline': 'Classroom',  # legacy
    'ojt': 'OJT', 'on-the-job': 'OJT', 'on the job': 'OJT', 'on_the_job': 'OJT',
    'sop': 'SOP', 'standard operating procedure': 'SOP',
    'online': 'Online', 'virtual': 'Online', 'vc': 'Online',
    'zoom': 'Online', 'teams': 'Online', 'webinar': 'Online',
}

_AUDIENCE_ABBR = {
    'bc': 'Blue Collared', 'blue': 'Blue Collared', 'blu': 'Blue Collared',
    'blue collar': 'Blue Collared', 'blue-collar': 'Blue Collared',
    'blue collared': 'Blue Collared', 'workman': 'Blue Collared',
    'workmen': 'Blue Collared', 'worker': 'Blue Collared',
    'wc': 'White Collared', 'white': 'White Collared',
    'white collar': 'White Collared', 'white-collar': 'White Collared',
    'white collared': 'White Collared', 'executive': 'White Collared',
    'staff': 'White Collared', 'officer': 'White Collared',
    'all': 'Common', 'both': 'Common', 'common': 'Common',
    'mixed': 'Common', 'any': 'Common',
}


def normalise_enum(raw, abbr_map, valid_values):
    """Generic enum cleanup: lookup in abbr map, then fuzzy match against valid
    list. Returns (canonical_value | None, confidence, method)."""
    if raw is None or str(raw).strip() == '':
        return None, 0.0, 'empty'
    s = normalise(raw).lower()
    # Also try dot-stripped variant (catches 'I.T.' → 'it')
    s_no_dot = s.replace('.', '').strip()
    if s in abbr_map:
        return abbr_map[s], 0.95, 'abbr'
    if s_no_dot != s and s_no_dot in abbr_map:
        return abbr_map[s_no_dot], 0.95, 'abbr'
    # Direct match (case-insensitive) against valid list
    for v in valid_values:
        if v.lower() == s:
            return v, 1.00, 'exact'
    # Word-order (HR/EHS vs EHS/HR)
    s_tokens = _tokens_sorted(s)
    for v in valid_values:
        if _tokens_sorted(v.lower()) == s_tokens:
            return v, 0.95, 'word-order'
    # Fuzzy
    best = None
    best_score = 0.0
    for v in valid_values:
        score = _fuzzy_score(s, v.lower())
        if score > best_score:
            best, best_score = v, score
    if best_score >= 0.85:
        return best, best_score, 'fuzzy'
    return None, best_score, 'no-match'


def analyze_prog_type(raw, valid_values):
    return normalise_enum(raw, _PROG_TYPE_ABBR, valid_values)


def analyze_mode(raw, valid_values):
    return normalise_enum(raw, _MODE_ABBR, valid_values)


def analyze_audience(raw, valid_values):
    return normalise_enum(raw, _AUDIENCE_ABBR, valid_values)


def analyze_emp_code(raw, plant_emp_codes):
    """Returns (canonical_emp_code | None, confidence, garbage_class).
    plant_emp_codes: set or list of valid emp_codes for the plant."""
    if raw is None or str(raw).strip() == '':
        return None, 0.0, 'empty'
    s = str(raw).strip().upper()
    if s in plant_emp_codes:
        return s, 1.00, None

    # Strip common prefixes
    s_clean = re.sub(r'^EMP[-_]?', '', s)
    s_clean = re.sub(r'^0+', '', s_clean)
    if s_clean in plant_emp_codes:
        return s_clean, 0.95, 'leading-zeros'

    # Strip punctuation
    s_squeeze = re.sub(r'[\s\-_./]', '', s)
    for c in plant_emp_codes:
        if re.sub(r'[\s\-_./]', '', c) == s_squeeze:
            return c, 0.92, 'punct-mismatch'

    # Single-char typo (same length)
    for c in plant_emp_codes:
        if len(c) == len(s) and sum(a != b for a, b in zip(c, s)) == 1:
            return c, 0.85, 'single-char-typo'

    return None, 0.0, 'no-match'
