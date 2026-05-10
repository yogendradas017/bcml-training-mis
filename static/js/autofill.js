// Auto-fill employee details from data attributes on the <option> element
function fillEmpDetails(sel, prefix) {
  const opt = sel.options[sel.selectedIndex];
  if (!opt || !opt.value) return;
  const set = (id, v) => { const el = document.getElementById(id); if (el) el.value = v || ''; };
  set(prefix + '_grade',  opt.dataset.grade);
  set(prefix + '_collar', opt.dataset.collar);
  set(prefix + '_dept',   opt.dataset.dept);
}

// Auto-dismiss success/warning alerts after 5 seconds
document.addEventListener('DOMContentLoaded', () => {
  document.querySelectorAll('.alert-success, .alert-warning').forEach(el => {
    setTimeout(() => {
      el.classList.remove('show');
      el.classList.add('fade');
    }, 5000);
  });
});
