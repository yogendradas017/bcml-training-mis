/* ═══════════════════════════════════════════════════
   PAGE LOADER
═══════════════════════════════════════════════════ */
window.addEventListener('load', () => {
  const l = document.getElementById('page-loader');
  if (l) { l.classList.add('done'); setTimeout(() => l.remove(), 400); }
});

/* ═══════════════════════════════════════════════════
   MOBILE SIDEBAR TOGGLE
═══════════════════════════════════════════════════ */
const menuBtn  = document.getElementById('menu-btn');
const sidebar  = document.querySelector('.sidebar');
const overlay  = document.getElementById('sb-overlay');
if (menuBtn && sidebar && overlay) {
  menuBtn.addEventListener('click', () => {
    sidebar.classList.toggle('open');
    overlay.classList.toggle('show');
  });
  overlay.addEventListener('click', () => {
    sidebar.classList.remove('open');
    overlay.classList.remove('show');
  });
  // close on nav link click (mobile)
  sidebar.querySelectorAll('.sb-nav a').forEach(a => {
    a.addEventListener('click', () => {
      sidebar.classList.remove('open');
      overlay.classList.remove('show');
    });
  });
}

/* ═══════════════════════════════════════════════════
   COUNTER ANIMATION
═══════════════════════════════════════════════════ */
function animateCount(el) {
  const raw = el.textContent.replace(/,/g, '').trim();
  const target = parseFloat(raw);
  if (isNaN(target) || target === 0) return;
  const duration = 900;
  const start = performance.now();
  const isInt = Number.isInteger(target);
  function step(now) {
    const p = Math.min((now - start) / duration, 1);
    const ease = 1 - Math.pow(1 - p, 3);
    const val = ease * target;
    el.textContent = isInt
      ? Math.round(val).toLocaleString()
      : val.toFixed(0);
    if (p < 1) requestAnimationFrame(step);
    else el.textContent = isInt ? target.toLocaleString() : target.toFixed(0);
  }
  requestAnimationFrame(step);
}

// Trigger counter when stat cards enter viewport
const countObserver = new IntersectionObserver((entries) => {
  entries.forEach(e => {
    if (e.isIntersecting) {
      e.target.querySelectorAll('.sc-val').forEach(animateCount);
      countObserver.unobserve(e.target);
    }
  });
}, { threshold: .3 });
document.querySelectorAll('.sc').forEach(sc => countObserver.observe(sc));

/* ── Close user dropdown on outside click ── */
document.addEventListener('click', e => {
  const menu = document.getElementById('userMenu');
  if (menu && !menu.contains(e.target)) menu.classList.remove('open');
});

/* ═══════════════════════════════════════════════════
   RIPPLE EFFECT ON BUTTONS
═══════════════════════════════════════════════════ */
document.addEventListener('click', e => {
  const btn = e.target.closest('.btn');
  if (!btn) return;
  const r = document.createElement('span');
  r.className = 'ripple';
  const rect = btn.getBoundingClientRect();
  const size = Math.max(rect.width, rect.height);
  r.style.cssText = `width:${size}px;height:${size}px;left:${e.clientX - rect.left - size/2}px;top:${e.clientY - rect.top - size/2}px`;
  btn.appendChild(r);
  setTimeout(() => r.remove(), 600);
});

/* ═══════════════════════════════════════════════════
   BUTTON LOADING ON FORM SUBMIT
═══════════════════════════════════════════════════ */
document.addEventListener('submit', e => {
  const form = e.target;
  const submitBtn = form.querySelector('[type=submit]');
  if (submitBtn && !submitBtn.classList.contains('no-loading')) {
    submitBtn.classList.add('btn-loading');
    setTimeout(() => submitBtn.classList.remove('btn-loading'), 8000);
  }
});

/* ═══════════════════════════════════════════════════
   INLINE FORM VALIDATION FEEDBACK
═══════════════════════════════════════════════════ */
document.querySelectorAll('input[required], select[required], textarea[required]').forEach(field => {
  field.addEventListener('blur', () => {
    if (field.value.trim()) {
      field.classList.remove('is-invalid'); field.classList.add('is-valid');
    } else {
      field.classList.remove('is-valid'); field.classList.add('is-invalid');
    }
  });
  field.addEventListener('input', () => {
    if (field.value.trim()) {
      field.classList.remove('is-invalid'); field.classList.add('is-valid');
    }
  });
});

/* ═══════════════════════════════════════════════════
   PROG_AC — Programme Name Autocomplete (shared)
   Usage: PROG_AC.setup(inputEl)
   Fetches master list once, shows dropdown + warns
   if value is not in master list.
═══════════════════════════════════════════════════ */
const PROG_AC = (() => {
  let _list = null;
  async function _load() {
    if (_list) return _list;
    try {
      const r = await fetch('/api/programme-list');
      _list = await r.json();
    } catch(e) { _list = []; }
    return _list;
  }

  function _sim(a, b) {
    a = a.toLowerCase(); b = b.toLowerCase();
    if (a === b) return 1;
    if (!a.length || !b.length) return 0;
    if (b.includes(a) || a.includes(b)) return 0.85;
    const dp = Array.from({length: b.length+1}, (_,i) => i);
    for (let i = 1; i <= a.length; i++) {
      let prev = i;
      for (let j = 1; j <= b.length; j++) {
        const v = a[i-1]===b[j-1] ? dp[j-1] : 1 + Math.min(dp[j-1],dp[j],prev);
        dp[j-1]=prev; prev=v;
      }
      dp[b.length]=prev;
    }
    return 1 - dp[b.length]/Math.max(a.length,b.length);
  }

  function setup(input) {
    if (!input) return;
    input.setAttribute('autocomplete','off');

    // Wrap input to position dropdown
    const wrap = document.createElement('div');
    wrap.style.cssText = 'position:relative;display:block;';
    input.parentNode.insertBefore(wrap, input);
    wrap.appendChild(input);

    const drop = document.createElement('div');
    drop.style.cssText = 'display:none;position:absolute;top:100%;left:0;right:0;z-index:9999;background:#fff;border:1px solid #e5e9f2;border-radius:8px;box-shadow:0 4px 20px rgba(0,0,0,.12);max-height:220px;overflow-y:auto;';
    wrap.appendChild(drop);

    const allowNew = input.dataset.allowNew === '1';

    const warn = document.createElement('div');
    warn.style.cssText = 'display:none;font-size:11px;margin-top:4px;';
    wrap.appendChild(warn);

    // Hidden flag sent to backend when user confirms add-to-master
    let _autoAddInput = null;

    function _lockSourceToNewRequirement(form, lock) {
      const sel = form.querySelector('[name=source]');
      if (!sel) return;
      if (lock) {
        sel.value = 'New Requirement';
        sel.disabled = true;
        sel.title = 'Locked to New Requirement — programme does not exist in TNI.';
      } else {
        sel.disabled = false;
        sel.title = '';
      }
    }

    function _showNotInMasterWarning() {
      if (allowNew) {
        const form = input.closest('form');
        _lockSourceToNewRequirement(form, true);
        warn.style.color = '#92400e';
        warn.innerHTML =
          '<i class="bi bi-exclamation-triangle me-1"></i>Not in Programme Master — source locked to <strong>New Requirement</strong>. ' +
          '<a href="#" id="prog-ac-add-master-btn" style="color:#1d4ed8;font-weight:600;text-decoration:underline;">' +
          '➕ Add to Programme Master &amp; Continue</a>';
        warn.style.display = 'block';
        const addBtn = warn.querySelector('#prog-ac-add-master-btn');
        if (addBtn) {
          addBtn.onclick = function(e) {
            e.preventDefault();
            if (!_autoAddInput) {
              _autoAddInput = document.createElement('input');
              _autoAddInput.type = 'hidden';
              _autoAddInput.name = 'auto_add_to_master';
              _autoAddInput.value = '1';
              form.appendChild(_autoAddInput);
            }
            warn.style.color = '#065f46';
            warn.innerHTML = '<i class="bi bi-check-circle me-1"></i>Will be added to Programme Master as <strong>New Requirement</strong> on save.';
            input.style.borderColor = '#10b981';
            _setSubmitBlocked(false);
          };
        }
      } else {
        warn.style.color = '#c2410c';
        warn.innerHTML = '<i class="bi bi-exclamation-circle me-1"></i>Not in Programme Master — check spelling or add to master list first.';
        warn.style.display = 'block';
      }
    }

    let _chosen = false;

    input.addEventListener('input', async function() {
      _chosen = false;
      warn.style.display = 'none';
      const q = this.value.trim();
      drop.innerHTML = ''; drop.style.display = 'none';
      if (q.length < 2) return;
      const list = await _load();
      const ql = q.toLowerCase();
      const scored = list
        .map(n => ({n, s: _sim(ql, n.toLowerCase())}))
        .filter(x => x.s > 0.35 || x.n.toLowerCase().includes(ql))
        .sort((a,b) => b.s - a.s)
        .slice(0, 8);
      if (!scored.length) return;
      scored.forEach(({n, s}) => {
        const d = document.createElement('div');
        const isExact = n.toLowerCase() === ql;
        const tag = isExact ? '' : s >= 0.75
          ? '<span style="font-size:10px;background:#fef3c7;color:#92400e;padding:1px 6px;border-radius:10px;margin-left:6px;">similar</span>' : '';
        d.innerHTML = `<span style="font-size:13px;">${n}</span>${tag}`;
        d.style.cssText = 'padding:8px 14px;cursor:pointer;border-bottom:1px solid #f1f5f9;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;';
        d.onmouseenter = () => d.style.background='#f8fafc';
        d.onmouseleave = () => d.style.background='';
        d.onmousedown = (e) => { e.preventDefault(); input.value=n; _chosen=true; drop.style.display='none'; warn.style.display='none'; input.dispatchEvent(new Event('input')); };
        drop.appendChild(d);
      });
      drop.style.display = 'block';
    });

    function _setSubmitBlocked(blocked) {
      const form = input.closest('form');
      if (!form) return;
      const btn = form.querySelector('[type=submit]');
      if (!btn) return;
      btn.disabled = blocked;
      btn.style.opacity = blocked ? '0.45' : '';
      btn.title = blocked ? 'Programme not in master list — add it to Programme Master first.' : '';
    }

    input.addEventListener('blur', async function() {
      setTimeout(async () => {
        drop.style.display = 'none';
        if (_chosen || !this.value.trim()) { warn.style.display='none'; _setSubmitBlocked(false); return; }
        const list = await _load();
        if (!list.length) return;
        const exact = list.find(n => n.toLowerCase() === this.value.trim().toLowerCase());
        if (exact) {
          warn.style.display = 'none';
          input.style.borderColor = '';
          _setSubmitBlocked(false);
        } else {
          input.style.borderColor = '#f97316';
          _showNotInMasterWarning();
          _setSubmitBlocked(true);
        }
      }, 150);
    });

    input.addEventListener('focus', () => {
      input.style.borderColor = '';
      warn.style.display = 'none';
      _setSubmitBlocked(false);
      if (_autoAddInput) { _autoAddInput.remove(); _autoAddInput = null; }
      if (allowNew) _lockSourceToNewRequirement(input.closest('form'), false);
    });

    document.addEventListener('click', e => {
      if (!wrap.contains(e.target)) drop.style.display='none';
    });
  }

  return { setup };
})();

/* ═══════════════════════════════════════════════════
   TBL — Universal table search + pagination engine
═══════════════════════════════════════════════════ */
const TBL = {
  _st: {},

  init(id, pageSize) {
    this._st[id] = { page:1, pageSize: pageSize||10, search:'' };
    this._render(id);
  },

  search(id, q) {
    if (!this._st[id]) return;
    this._st[id].search = q.trim().toLowerCase(); this._st[id].page = 1; this._render(id);
  },

  filter(id, key, val) {
    if (!this._st[id]) return;
    this._st[id][key] = val.toLowerCase(); this._st[id].page = 1; this._render(id);
  },

  go(id, p) {
    if (!this._st[id]) return;
    this._st[id].page = p; this._render(id);
  },

  _visible(id) {
    const st = this._st[id];
    const tbl = document.getElementById(id);
    if (!tbl) return [];
    const q = st.search || '';
    return Array.from(tbl.querySelectorAll('tbody tr')).filter(r => {
      if (r.dataset.empty) return false;
      const txt = r.textContent.toLowerCase();
      if (q && !txt.includes(q)) return false;
      for (const [k, v] of Object.entries(st)) {
        if (['page','pageSize','search'].includes(k) || !v) continue;
        const dv = r.dataset[k];
        if (dv !== undefined) {
          if (dv.toLowerCase() !== v.toLowerCase()) return false;
        } else {
          if (!txt.includes(v.toLowerCase())) return false;
        }
      }
      return true;
    });
  },

  _slCol(id) {
    const tbl = document.getElementById(id);
    if (!tbl) return -1;
    const ths = tbl.querySelectorAll('thead th');
    for (let i = 0; i < ths.length; i++) {
      const t = ths[i].textContent.trim();
      if (t === '#' || t === 'Sl' || t === 'SL' || t === 'S.No' || t === 'No.') return i;
    }
    return -1;
  },

  _render(id) {
    const st = this._st[id];
    const tbl = document.getElementById(id);
    if (!tbl) return;
    const all = Array.from(tbl.querySelectorAll('tbody tr'));
    const emptyRow = all.find(r => r.dataset.empty);
    const vis = this._visible(id);

    all.forEach(r => r.style.display = 'none');

    if (vis.length === 0) {
      if (emptyRow) emptyRow.style.display = '';
      this._pager(id, 0, 1, 1); return;
    }

    const ps = st.pageSize;
    const pages = Math.max(1, Math.ceil(vis.length / ps));
    st.page = Math.max(1, Math.min(st.page, pages));
    const pageRows = vis.slice((st.page-1)*ps, st.page*ps);
    const slIdx = this._slCol(id);
    pageRows.forEach((r, i) => {
      r.style.display = '';
      if (slIdx >= 0) {
        const cell = r.querySelectorAll('td')[slIdx];
        if (cell) cell.textContent = (st.page - 1) * ps + i + 1;
      }
    });
    this._pager(id, vis.length, st.page, pages);
  },

  _pager(id, total, page, pages) {
    const st = this._st[id];
    let el = document.getElementById(id+'_pager');
    if (!el) {
      el = document.createElement('div');
      el.id = id + '_pager'; el.className = 'tbl-pager';
      const tbl = document.getElementById(id);
      const wrap = tbl.closest('.table-responsive') || tbl.parentElement;
      wrap.insertAdjacentElement('afterend', el);
    }
    if (total === 0) { el.innerHTML = ''; return; }
    const s = (page-1)*st.pageSize+1, e2 = Math.min(page*st.pageSize, total);
    const info = `<span class="tbl-info">Showing <strong>${s}–${e2}</strong> of <strong>${total}</strong> records</span>`;
    const b = (label, p, dis, active) =>
      `<button class="pager-btn${active?' pg-active':''}" onclick="TBL.go('${id}',${p})" ${dis?'disabled':''}>${label}</button>`;
    let btns = `<div class="pager-btns">`;
    btns += b('<i class="bi bi-chevron-double-left"></i>', 1, page===1, false);
    btns += b('<i class="bi bi-chevron-left"></i>', page-1, page===1, false);
    const st2 = Math.max(1, Math.min(page-2, pages-4));
    const en = Math.min(pages, st2+4);
    for (let i=st2; i<=en; i++) btns += b(i, i, false, i===page);
    btns += b('<i class="bi bi-chevron-right"></i>', page+1, page===pages, false);
    btns += b('<i class="bi bi-chevron-double-right"></i>', pages, page===pages, false);
    btns += `</div>`;
    el.innerHTML = info + btns;
  },

  buildDynamic(tableId, dataKey, selectEl, label) {
    const tbl = document.getElementById(tableId);
    if (!tbl || !selectEl) return;
    const vals = new Set();
    tbl.querySelectorAll('tbody tr:not([data-empty])').forEach(r => {
      const v = r.dataset[dataKey];
      if (v && v.trim() && v.trim() !== '—') vals.add(v.trim());
    });
    [...vals].sort().forEach(v => {
      const o = document.createElement('option');
      o.value = v; o.textContent = v; selectEl.appendChild(o);
    });
  }
};
window.TBL = TBL;

/* ═══════════════════════════════════════════════════
   AJAX ROW DELETE
═══════════════════════════════════════════════════ */
function ajaxDelete(btn, confirmMsg) {
  if (!confirm(confirmMsg || 'Delete this record? This cannot be undone.')) return;
  const form = btn.closest('form');
  const row  = btn.closest('tr');
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner-border spinner-border-sm" style="width:12px;height:12px;border-width:2px;"></span>';
  fetch(form.action, {
    method: 'POST',
    headers: { 'X-Requested-With': 'XMLHttpRequest' }
  }).then(r => {
    if (r.status === 204 || r.ok) {
      row.style.transition = 'opacity .18s ease, transform .18s ease';
      row.style.opacity = '0';
      row.style.transform = 'translateX(16px)';
      setTimeout(() => row.remove(), 200);
      showToast('Deleted successfully.', 'warning');
    } else {
      btn.disabled = false;
      btn.innerHTML = '<i class="bi bi-trash"></i>';
      showToast('Delete failed. Please try again.', 'danger');
    }
  }).catch(() => {
    btn.disabled = false;
    btn.innerHTML = '<i class="bi bi-trash"></i>';
    showToast('Network error.', 'danger');
  });
}

function showToast(msg, type) {
  const t = document.createElement('div');
  t.className = `alert alert-${type}`;
  t.style.cssText = 'position:fixed;top:76px;right:20px;z-index:9999;min-width:220px;max-width:340px;animation:fadeUp .2s ease both;box-shadow:0 8px 24px rgba(0,0,0,.15);';
  t.innerHTML = `<i class="bi bi-${type==='warning'?'check-circle-fill':'exclamation-triangle-fill'}"></i> ${msg}`;
  document.body.appendChild(t);
  setTimeout(() => { t.style.transition = 'opacity .3s'; t.style.opacity = '0'; setTimeout(() => t.remove(), 320); }, 2400);
}

/* ═══════════════════════════════════════════════════
   BULK SELECT UTILITY
═══════════════════════════════════════════════════ */
const BulkSelect = {
  _tableId: null,
  _actionUrl: null,

  init(tableId, actionUrl) {
    this._tableId   = tableId;
    this._actionUrl = actionUrl;
    const tbl   = document.getElementById(tableId);
    if (!tbl) return;
    const allCb = tbl.querySelector('.cb-all');

    tbl.addEventListener('change', e => {
      if (e.target.classList.contains('cb-row') || e.target.classList.contains('cb-all')) {
        if (e.target.classList.contains('cb-all')) {
          const visible = [...tbl.querySelectorAll('tbody tr')]
            .filter(r => r.style.display !== 'none' && !r.dataset.empty);
          visible.forEach(r => { const c = r.querySelector('.cb-row'); if (c) c.checked = e.target.checked; });
        }
        this._update(tbl, allCb, actionUrl);
      }
    });

    document.getElementById('bulk-delete-btn').addEventListener('click', () => {
      const checked = [...tbl.querySelectorAll('.cb-row:checked')];
      if (!checked.length) return;
      if (!confirm(`Delete ${checked.length} selected record(s)? This cannot be undone.`)) return;
      const form = document.createElement('form');
      form.method = 'POST'; form.action = this._actionUrl;
      checked.forEach(cb => {
        const inp = document.createElement('input');
        inp.type = 'hidden'; inp.name = 'ids[]'; inp.value = cb.value;
        form.appendChild(inp);
      });
      document.body.appendChild(form);
      form.submit();
    });
  },

  _update(tbl, allCb, actionUrl) {
    const all     = [...tbl.querySelectorAll('tbody .cb-row')];
    const checked = all.filter(c => c.checked);
    const bar     = document.getElementById('bulk-bar');
    document.getElementById('bulk-count').textContent = checked.length + ' selected';
    bar.dataset.action = actionUrl;
    if (checked.length > 0) { bar.classList.add('show'); }
    else { bar.classList.remove('show'); if (allCb) { allCb.checked = false; allCb.indeterminate = false; } }
    if (allCb) {
      allCb.indeterminate = checked.length > 0 && checked.length < all.length;
      allCb.checked = all.length > 0 && checked.length === all.length;
    }
  },

  clear() {
    if (!this._tableId) return;
    const tbl = document.getElementById(this._tableId);
    if (tbl) tbl.querySelectorAll('.cb-row, .cb-all').forEach(c => { c.checked = false; c.indeterminate = false; });
    document.getElementById('bulk-bar').classList.remove('show');
  }
};
