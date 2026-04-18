// Auto-fill employee details from API
function fillEmpDetails(empCode, prefix) {
  if (!empCode) return;
  fetch('/api/employee/' + encodeURIComponent(empCode))
    .then(r => r.json())
    .then(d => {
      const set = (id, val) => { const el = document.getElementById(id); if (el) el.value = val || ''; };
      set(prefix + '_grade',  d.grade);
      set(prefix + '_collar', d.collar);
      set(prefix + '_dept',   d.department);
    });
}

// Auto-fill session details into Training 2A form
// Session codes contain '/' so encode each segment separately
function fillSessionDetails(code) {
  if (!code) return;
  fetch('/api/session/' + code.split('/').map(encodeURIComponent).join('/'))
    .then(r => r.json())
    .then(d => {
      const set = (id, val) => { const el = document.getElementById(id); if (el) el.value = val || ''; };
      set('tr_prog_name',      d.programme_name);
      set('tr_prog_type_disp', d.prog_type);
      set('tr_mode_disp',      d.mode);
      set('tr_hrs',            d.duration_hrs);
      if (d.plan_start) set('tr_start_date', d.plan_start);
      if (d.plan_end)   set('tr_end_date',   d.plan_end);
    });
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
