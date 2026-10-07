// UI helpers: toasts, confirm dialog, drawer, view switching, string normalizing.
export const $ = id => document.getElementById(id);
export const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
// Normalize names: strip "Gram Panchayat"/"Panchayat Samiti", collapse spaces, lowercase.
export const norm = s => String(s ?? '').replace(/gram\s+panchayat|panchayat\s+samiti/gi, '').replace(/\s+/g, ' ').trim().toLowerCase();

export function showToast(msg, type = '') {
  const t = document.createElement('div'); t.className = 'toast ' + type; t.textContent = msg;
  $('toasts').appendChild(t); setTimeout(() => t.remove(), 4000);
}
export function confirmDialog(msg) {
  return new Promise(res => {
    const d = document.createElement('dialog');
    d.innerHTML = `<p>${esc(msg)}</p><button value="n">Cancel</button> <button class="primary" value="y">Confirm</button>`;
    d.onclick = e => { if (e.target.value) { d.close(); res(e.target.value === 'y'); d.remove(); } };
    document.body.appendChild(d); d.showModal();
  });
}
export function fillSelect(sel, items, placeholder) { // items: [{v,t}]
  sel.innerHTML = (placeholder ? `<option value="">${esc(placeholder)}</option>` : '') +
    items.map(i => `<option value="${esc(i.v)}">${esc(i.t)}</option>`).join('');
}
export function setupNav(onView) {
  const toggle = o => { $('drawer').classList.toggle('open', o); $('scrim').classList.toggle('open', o); };
  $('menuBtn').onclick = () => toggle(!$('drawer').classList.contains('open'));
  $('scrim').onclick = () => toggle(false);
  document.addEventListener('click', e => {
    const el = e.target.closest && e.target.closest('[data-view]'); if (!el) return;
    const v = el.dataset.view;
    document.querySelectorAll('.view').forEach(s => s.classList.toggle('active', s.id === 'v-' + v));
    document.querySelectorAll('#drawer a').forEach(a => a.classList.toggle('on', a.dataset.view === v));
    window.scrollTo({ top: 0 }); toggle(false); onView && onView(v);
  });
}
// Convert raw errors into friendly messages
export function friendly(e) {
  const m = String(e && (e.code || e.message) || e);
  if (/permission-denied/.test(m)) return 'Permission denied. Please log in as admin or check Firestore rules.';
  if (/unavailable|network|offline/i.test(m)) return 'Network problem. Please check your internet connection and try again.';
  if (/TEMPLATE_MISSING/.test(m)) return 'Word template not found at template/covering-letter-template.docx.';
  if (/auth\/(invalid|wrong|user)/.test(m)) return 'Incorrect email or password.';
  return 'Something went wrong. Please try again.';
}
