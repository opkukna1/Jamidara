import * as FS from './js/firestore.js';
import * as UI from './js/ui.js';
import { generateDocx, download } from './js/docx-generator.js';
import { buildAmsReport, parseDate, normPS } from './js/ams-report.js';
const { $, norm, esc, showToast } = UI;

// Login button is wired FIRST, so it works even if something later fails. The real error code is shown to help debugging.
$('gBtn').onclick = () => {
  $('lerr').textContent = '';
  FS.googleLogin().catch(e => { console.error(e); $('lerr').textContent = UI.friendly(e) + ' [' + (e.code || e.message) + ']'; });
};
window.addEventListener('error', ev => { if (!document.body.classList.contains('authed')) $('lerr').textContent = 'Error: ' + ev.message; });

// ---- Field mapping from Firestore dispatch fields to template placeholders (edit if needed) ----
const F = { dispatchNo: 'Audit Party No', dispatchDate: 'Planned Start Date', paraNo: 'Converted to Para' };
// Admin Google accounts (lowercase). Only these see Upload. Firestore rules must list the same emails.
const ADMIN_EMAILS = ['opsiddh41@gmail.com'];
// Report types: template file inside /template and download filename prefix.
const REPORTS = { ams: { tpl: 'ams-report.docx', prefix: 'AMS_Report' }, ybc: { tpl: 'year-book-closing.docx', prefix: 'Year_Book_Closing' }, intim: { tpl: 'intimation-report.docx', prefix: 'Intimation_Report' } };
const REQUIRED = ['Unit Name', 'Unit ID', 'Parent Name', 'District Name'];
const DEFAULT_YEARS = ['2025-26', '2026-27'];
const S = { maps: [], disp: [], mapIdx: new Map(), tree: {}, years: [], admin: false, rtype: 'ams', amsMode: 1, last: null, rows: [], ids: new Set() };

// ---------- Data + in-memory indexes ----------
async function loadAll() {
  try {
    [S.maps, S.disp] = await Promise.all([FS.loadMappings(), FS.loadDispatchData()]);
    S.ids = new Set(S.disp.map(d => d._id));
    buildIndexes(); initSelectors();
  } catch (e) { showToast(UI.friendly(e), 'err'); }
}
const mkey = (d, p, g) => [norm(d), norm(p), norm(g)].join('|');
function findMapping(r) { return S.mapIdx.get(mkey(r['District Name'], r['Parent Name'], r['Unit Name'])); }
function buildIndexes() {
  S.mapIdx = new Map(S.maps.map(m => [mkey(m.DIST_EN, m.PS_EN, m.GP_EN), m]));
  S.tree = {};
  for (const r of S.disp) { // year -> district -> PS -> [records]
    const y = r.Year || '', d = r['District Name'], p = r['Parent Name'];
    if (!String(p || '').trim()) continue; // PS rows (no parent) are only used by the AMS report
    (((S.tree[y] ??= {})[d] ??= {})[p] ??= []).push(r);
  }
  S.years = [...new Set([...DEFAULT_YEARS, ...Object.keys(S.tree).filter(Boolean)])].sort();
  $('stD').textContent = S.disp.length; $('stM').textContent = S.maps.length; $('stY').textContent = S.years.length;
}
const hi = (r, f, k) => { const m = findMapping(r); return m && m[k] ? m[k] : null; };
const label = (en, hiTxt) => hiTxt ? `${en} (${hiTxt})` : en;
const opts = (arr) => arr.map(x => ({ v: x, t: x }));

// ---------- Dropdowns ----------
function initSelectors() {
  UI.fillSelect($('year'), opts(S.years), 'Select year'); populateDistricts(); renderAmsPs();
  UI.fillSelect($('upYear'), opts(S.years));
  UI.fillSelect($('fYear'), opts(S.years), 'All years'); renderData();
}
function recs(y, d, p) { return ((S.tree[y] || {})[d] || {})[p] || []; }
function populateDistricts() {
  const t = S.tree[$('year').value] || {};
  const items = Object.keys(t).sort().map(d => {
    const m = S.maps.find(m => norm(m.DIST_EN) === norm(d)); return { v: d, t: label(d, m && m.DIST_HI) };
  });
  UI.fillSelect($('dist'), items, 'Select district'); populatePS();
}
function populatePS() {
  const t = ((S.tree[$('year').value] || {})[$('dist').value]) || {};
  const items = Object.keys(t).sort().map(p => {
    const m = S.maps.find(m => norm(m.PS_EN) === norm(p) && norm(m.DIST_EN) === norm($('dist').value));
    return { v: p, t: label(p.replace(/panchayat samiti/i, '').trim() || p, m && m.PS_HI) };
  });
  UI.fillSelect($('ps'), items, 'Select panchayat samiti'); populateGPs();
}
function populateGPs() {
  const list = recs($('year').value, $('dist').value, $('ps').value);
  const items = list.map(r => ({ v: r._id, t: label(r['Unit Name'].replace(/gram panchayat/i, '').trim(), hi(r, 0, 'GP_HI')) }))
    .sort((a, b) => a.t.localeCompare(b.t));
  UI.fillSelect($('gp'), [{ v: '__ALL__', t: 'All Gram Panchayats' }, ...items], 'Select gram panchayat');
  showInfo();
}
function findDispatchRecord(id) { return S.disp.find(r => r._id === id); }
function showInfo() {
  const v = $('gp').value; $('warn').innerHTML = ''; $('result').hidden = true;
  if (!v) { $('info').innerHTML = ''; return; }
  const list = v === '__ALL__' ? recs($('year').value, $('dist').value, $('ps').value) : [findDispatchRecord(v)];
  const r = list[0];
  $('info').innerHTML = v === '__ALL__' ? `<b>${list.length}</b> Gram Panchayats will be included.` :
    `<b>Dispatch:</b> ${esc(r['Dispatch Name'] || '–')}<br><b>Audit Party No:</b> ${esc(r['Audit Party No'] || '–')}<br><b>Planned:</b> ${esc(r['Planned Start Date'] || '–')} to ${esc(r['Planned End Date'] || '–')}<br><b>Status:</b> ${esc(r['Report Status'] || '–')}`;
}

// ---------- Generate ----------
const pick = (r, names) => { const ks = Object.keys(r);
  for (const n of names) { const k = ks.find(k => k.toLowerCase() === n.toLowerCase()); if (k) return r[k]; }
  for (const n of names) { const k = ks.find(k => k.toLowerCase().includes(n.toLowerCase())); if (k) return r[k]; } return ''; };
const pad2 = x => String(x).padStart(2, '0');
/** Same placeholder values as CoverGeneretor.py */
function letterValues(r, m, year) {
  const dname = String(pick(r, ['Dispatch Name', 'Dispatch']) ?? '').trim();
  const dm = dname.match(/-(\d+)-/) || dname.match(/\b(\d+)\b/);
  const d = parseDate(pick(r, ['Report Approval Date', 'Dispatch Date', 'Approval Date', 'Date']));
  const para = String(pick(r, ['Converted to Para in Nos', 'Converted to Para', 'Para']) ?? '').trim(), n = parseFloat(para);
  const eng = String(m.GP_EN).replace(/^\s*gram\s+panchayat\s+/i, '').trim();
  return { YEAR: year, GP_NAME_EN: m.GP_EN, GP_NAME_ENG: eng, GP_NAME_HI: `${m.GP_HI} (${eng})`, PS_NAME_EN: m.PS_EN, PS_NAME_HI: m.PS_HI,
    DISTRICT_EN: m.DIST_EN, DISTRICT_HI: m.DIST_HI, DISPATCH_NO: dm ? dm[1] : dname, DISPATCH_NAME: dname,
    DATE: d ? `${pad2(d.getDate())}.${pad2(d.getMonth() + 1)}.${String(d.getFullYear()).slice(2)}` : '',
    PARA_COUNT: para ? (isNaN(n) ? para : String(Math.trunc(n) + 1)) : '', OFFICE_NAME: m.DIST_HI, DIVISION_NAME: m.DIST_HI,
    CONSTITUTION_OBJECTION: '0', SERIOUS_OBJECTION: '0', PARA_BREAKUP: para };
}
const genLabel = () => S.rtype === 'ams' ? 'Generate AMS Report' : 'Generate Cover Letter';
const safe = s => String(s).replace(/[^\w-]+/g, '_');
function generateCoveringLetter() { return S.rtype === 'ams' ? generateAms() : generateLetters(); }
async function withBusy(work) {
  const btn = $('genBtn'); btn.disabled = true; btn.textContent = 'Preparing your file...';
  try {
    const out = await work();
    if (out) { S.last = out; $('rname').textContent = out.name; $('result').hidden = false; $('result').scrollIntoView({ behavior: 'smooth', block: 'center' }); showToast('Generated successfully.', 'ok'); }
  } catch (e) {
    console.error(e);
    showToast(e.message === 'NO_PS' ? 'No data found for the selected Panchayat Samiti.' : e.properties && e.properties.errors ? 'The Word template has a formatting problem in its placeholders.' : UI.friendly(e), 'err');
  } finally { btn.disabled = false; btn.textContent = genLabel(); }
}
async function generateLetters() {
  const [y, d, p, g] = ['year', 'dist', 'ps', 'gp'].map(i => $(i).value);
  const miss = [!y && 'Year', !d && 'District', !p && 'Panchayat Samiti', !g && 'Gram Panchayat'].filter(Boolean);
  if (miss.length) return showToast('Please select: ' + miss.join(', '), 'err');
  const list = g === '__ALL__' ? recs(y, d, p) : [findDispatchRecord(g)].filter(Boolean);
  const ok = [], bad = [];
  list.forEach(r => { const m = findMapping(r); m ? ok.push([r, m]) : bad.push(r); });
  $('warn').innerHTML = bad.map(r => `⚠ Hindi mapping not found for:<br>${esc(r['Unit Name'])} (skipped)`).join('<br>');
  if (!ok.length) return showToast('No records with Hindi mapping found for this selection.', 'err');
  return withBusy(async () => {
    const blob = await generateDocx(ok.map(([r, m]) => letterValues(r, m, y)), REPORTS[S.rtype].tpl), m0 = ok[0][1];
    return { blob, name: `${REPORTS[S.rtype].prefix}_${safe(g === '__ALL__' ? m0.PS_EN : m0.GP_EN)}_${safe(m0.DIST_EN)}_${y}.docx` };
  });
}
function renderAmsPs() { // PS checklist built from the dispatch data of the chosen year
  const y = $('year').value, set = new Set();
  for (const r of S.disp) { if (y && r.Year !== y) continue; const n = normPS(String(r['Parent Name'] || '').trim() ? r['Parent Name'] : r['Unit Name']); if (n) set.add(n); }
  const list = [...set].sort(), title = n => n.toLowerCase().replace(/\b\w/g, c => c.toUpperCase());
  $('amsPs').innerHTML = list.length ? `<label><input type="checkbox" id="amsAll" checked><b>All Panchayat Samitis (${list.length})</b></label>` +
    list.map(n => `<label><input type="checkbox" data-ps="${esc(n)}" checked>${esc(title(n))}</label>`).join('') : '<div class="empty">Select a year with data.</div>';
  const boxes = () => [...$('amsPs').querySelectorAll('input[data-ps]')];
  if ($('amsAll')) {
    $('amsAll').onchange = e => boxes().forEach(b => b.checked = e.target.checked);
    boxes().forEach(b => b.onchange = () => $('amsAll').checked = boxes().every(x => x.checked));
  }
}
async function generateAms() {
  const y = $('year').value; if (!y) return showToast('Please select: Year', 'err');
  const picked = [...$('amsPs').querySelectorAll('input[data-ps]:checked')].map(i => i.dataset.ps);
  if (!picked.length) return showToast('Select at least one Panchayat Samiti.', 'err');
  const data = S.disp.filter(r => r.Year === y);
  const need = ['Audit Party No', 'Parent Name', 'Planned End Date', 'Report Status', 'Report Approval Date', 'Unit Name'];
  const miss = data.length ? need.filter(c => !(c in data[0])) : need;
  if (miss.length) return showToast('Dispatch data is missing columns: ' + miss.join(', '), 'err');
  return withBusy(async () => {
    const res = buildAmsReport(data, picked, S.amsMode);
    $('warn').innerHTML = `${res.psCount} Panchayat Samiti · ${res.totalGp} rows counted (${S.amsMode === 1 ? 'With PS' : 'Without PS'})`;
    return { blob: res.blob, name: `AMS_Report_${y}.docx` };
  });
}

// ---------- Dispatch Data page ----------
function renderData() {
  const y = $('fYear').value, d = $('fDist').value, p = $('fPs').value, st = $('fStatus').value, q = norm($('fSearch').value);
  const years = y ? [y] : Object.keys(S.tree);
  // refresh dependent filter lists
  const dists = [...new Set(years.flatMap(k => Object.keys(S.tree[k] || {})))].sort();
  if ($('fDist').options.length - 1 !== dists.length) UI.fillSelect($('fDist'), opts(dists), 'All districts'), $('fDist').value = d;
  const pss = [...new Set(years.flatMap(k => Object.keys((S.tree[k] || {})[$('fDist').value] || {})))].sort();
  if ($('fPs').options.length - 1 !== pss.length) UI.fillSelect($('fPs'), opts(pss), 'All samitis'), $('fPs').value = p;
  const sts = [...new Set(S.disp.map(r => r['Report Status']).filter(Boolean))].sort();
  if ($('fStatus').options.length - 1 !== sts.length) UI.fillSelect($('fStatus'), opts(sts), 'All statuses'), $('fStatus').value = st;
  const rows = S.disp.filter(r => (!y || r.Year === y) && (!$('fDist').value || r['District Name'] === $('fDist').value) &&
    (!$('fPs').value || r['Parent Name'] === $('fPs').value) && (!$('fStatus').value || r['Report Status'] === $('fStatus').value) &&
    (!q || norm(r['Unit Name']).includes(q) || String(r['Audit Party No'] || '').toLowerCase().includes(q) || String(r['Dispatch Name'] || '').toLowerCase().includes(q)));
  if (!rows.length) { $('dataOut').innerHTML = '<div class="empty">No matching records.</div>'; return; }
  const shown = rows.slice(0, 300);
  $('dataOut').innerHTML = `<p>${rows.length} records${rows.length > 300 ? ' (showing first 300)' : ''}</p>
  <div class="scroll tbl"><table><tr><th>GP</th><th>PS</th><th>District</th><th>Dispatch</th><th>Party No</th><th>Status</th></tr>${shown.map(r =>
    `<tr><td>${esc(r['Unit Name'])}</td><td>${esc(r['Parent Name'])}</td><td>${esc(r['District Name'])}</td><td>${esc(r['Dispatch Name'])}</td><td>${esc(r['Audit Party No'])}</td><td>${esc(r['Report Status'])}</td></tr>`).join('')}</table></div>
  <div class="cards">${shown.map(r => `<div class="rc"><b>${esc(r['Unit Name'])}</b><br>${esc(r['Parent Name'])} · ${esc(r['District Name'])}<br>Dispatch: ${esc(r['Dispatch Name'])}<br>Party No: ${esc(r['Audit Party No'])}<br>Status: ${esc(r['Report Status'])}</div>`).join('')}</div>`;
}

// ---------- Dispatch upload (admin only) ----------
function pad(n) { return String(n).padStart(2, '0'); }
const fmt = v => v instanceof Date ? `${pad(v.getDate())}/${pad(v.getMonth() + 1)}/${v.getFullYear()}` : (typeof v === 'string' ? v.trim().replace(/\s+/g, ' ') : v);
const canon = k => { k = k.trim(); const l = k.toLowerCase(); return l === 'unit id' ? 'Unit ID' : l === 'parent unit id' ? 'Parent Unit ID' : k; };
async function readExcel(file, hdr = 'Unit ID') {
  const wb = XLSX.read(await file.arrayBuffer(), { cellDates: true });
  // detect the right sheet: "Users" if present, else first sheet containing "Unit ID"
  const names = wb.SheetNames;
  const pick = names.find(n => n.toLowerCase() === 'users') || names.find(n => (XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1 })[0] || []).some(c => String(c).trim().toLowerCase() === hdr.toLowerCase()));
  if (!pick) throw new Error('NO_SHEET');
  return XLSX.utils.sheet_to_json(wb.Sheets[pick], { defval: '' }).map(r => Object.fromEntries(Object.entries(r).map(([k, v]) => [canon(k), fmt(v)])));
}
function validateDispatchFile(rows) {
  if (!rows.length) return 'The file has no data rows.';
  const miss = REQUIRED.filter(c => !(c in rows[0]));
  return miss.length ? 'Missing required columns: ' + miss.join(', ') : '';
}
async function onFile() {
  S.rows = []; $('upBtn').disabled = $('prevBtn').disabled = true; $('preview').innerHTML = '';
  const f = $('file').files[0]; if (!f) return;
  try {
    const rows = await readExcel(f), err = validateDispatchFile(rows);
    if (err) return $('fileInfo').innerHTML = '⚠ ' + esc(err), showToast(err, 'err');
    S.rows = rows;
    $('fileInfo').innerHTML = `Selected file:<br><b>${esc(f.name)}</b><br>Rows detected: <b>${rows.length.toLocaleString()}</b><br>Columns detected: <b>${Object.keys(rows[0]).length}</b>`;
    $('prevBtn').disabled = $('upBtn').disabled = false;
  } catch (e) { showToast('This does not look like a valid Dispatch Excel file.', 'err'); $('fileInfo').textContent = ''; }
}
function previewData() {
  const cols = Object.keys(S.rows[0]).slice(0, 8);
  $('preview').innerHTML = `<table><tr>${cols.map(c => `<th>${esc(c)}</th>`).join('')}</tr>${S.rows.slice(0, 15).map(r => `<tr>${cols.map(c => `<td>${esc(r[c])}</td>`).join('')}</tr>`).join('')}</table>`;
}
async function uploadDispatchFile() {
  const year = $('upYear').value;
  if (!await UI.confirmDialog(`Upload ${S.rows.length} records for ${year}? Existing records will be updated; nothing is deleted.`)) return;
  $('upBtn').disabled = true;
  try {
    const r = await FS.uploadDispatchRows(S.rows, year, S.ids, (a, b) => $('upRes').textContent = `Uploading ${a}/${b}...`);
    $('upRes').innerHTML = `<b>Upload completed</b><br>New records: ${r.inserted}<br>Updated records: ${r.updated.toLocaleString()}<br>Skipped: ${r.skipped}<br>Errors: ${r.errors}`;
    showToast('Upload completed', r.errors ? 'err' : 'ok'); await loadAll();
  } catch (e) { showToast(UI.friendly(e), 'err'); $('upRes').textContent = ''; }
  $('upBtn').disabled = false;
}

// ---------- Mapping upload (admin only) ----------
const MAP_COLS = ['GP_EN', 'GP_HI', 'PS_EN', 'PS_HI', 'DIST_EN', 'DIST_HI'];
async function onMapFile() {
  S.mapRows = []; $('mBtn').disabled = true; $('mRes').textContent = '';
  const f = $('mFile').files[0]; if (!f) return;
  try {
    const rows = await readExcel(f, 'GP_EN');
    const miss = rows.length ? MAP_COLS.filter(c => !(c in rows[0])) : MAP_COLS;
    if (miss.length) { $('mInfo').innerHTML = '⚠ Missing columns: ' + esc(miss.join(', ')); return showToast('Mapping file has missing columns.', 'err'); }
    S.mapRows = rows;
    $('mInfo').innerHTML = `Selected file:<br><b>${esc(f.name)}</b><br>Rows detected: <b>${rows.length.toLocaleString()}</b>`;
    $('mBtn').disabled = false;
  } catch (e) { $('mInfo').textContent = ''; showToast('This does not look like a valid Mapping Excel file (sheet needs GP_EN column).', 'err'); }
}
async function uploadMappingFile() {
  if (!await UI.confirmDialog(`Upload ${S.mapRows.length} mapping rows? Existing rows will be updated; nothing is deleted.`)) return;
  $('mBtn').disabled = true;
  try {
    const r = await FS.uploadMappingRows(S.mapRows, new Set(S.maps.map(m => m._id)), (a, b) => $('mRes').textContent = `Uploading ${a}/${b}...`);
    $('mRes').innerHTML = `<b>Upload completed</b><br>New records: ${r.inserted}<br>Updated records: ${r.updated.toLocaleString()}<br>Skipped: ${r.skipped}<br>Errors: ${r.errors}`;
    showToast('Mapping upload completed', r.errors ? 'err' : 'ok'); await loadAll();
  } catch (e) { showToast(UI.friendly(e), 'err'); $('mRes').textContent = ''; }
  $('mBtn').disabled = false;
}

// ---------- Wiring ----------
$('mFile').onchange = onMapFile; $('mBtn').onclick = uploadMappingFile;
UI.setupNav(v => { if (v === 'data') renderData(); });
document.querySelector('.view').classList.add('active');
$('year').onchange = () => { populateDistricts(); renderAmsPs(); }; $('dist').onchange = populatePS; $('ps').onchange = populateGPs; $('gp').onchange = showInfo;
$('genBtn').onclick = generateCoveringLetter;
['fYear', 'fDist', 'fPs', 'fStatus'].forEach(i => $(i).onchange = renderData); $('fSearch').oninput = renderData;
$('file').onchange = onFile; $('prevBtn').onclick = previewData; $('upBtn').onclick = uploadDispatchFile; $('reloadBtn').onclick = loadAll;
// ---------- WhatsApp share ----------
async function shareWhatsApp() {
  const { blob, name } = S.last, file = new File([blob], name, { type: blob.type });
  try { // phones: share the actual Word file straight into WhatsApp
    if (navigator.canShare && navigator.canShare({ files: [file] })) { await navigator.share({ files: [file], text: name }); return; }
  } catch (e) { if (e.name === 'AbortError') return; }
  download(blob, name); // fallback: download, then open WhatsApp to attach manually
  window.open('https://wa.me/?text=' + encodeURIComponent('Please find the attached letter: ' + name), '_blank');
  showToast('File downloaded. Attach it in WhatsApp.');
}
$('rtypes').onclick = e => {
  const b = e.target.closest('button'); if (!b) return;
  S.rtype = b.dataset.t; [...$('rtypes').children].forEach(c => c.classList.toggle('on', c === b)); $('result').hidden = true; $('warn').innerHTML = '';
  $('letterSel').hidden = S.rtype === 'ams'; $('amsBox').hidden = S.rtype !== 'ams'; $('genBtn').textContent = genLabel();
};
$('dlBtn').onclick = () => S.last && download(S.last.blob, S.last.name);
$('waBtn').onclick = () => S.last && shareWhatsApp();

$('amsMode').onclick = e => { const b = e.target.closest('button'); if (!b) return; S.amsMode = +b.dataset.m; [...$('amsMode').children].forEach(c => c.classList.toggle('on', c === b)); };
// AMS is the default report type: show its options first
$('letterSel').hidden = true; $('amsBox').hidden = false; $('genBtn').textContent = genLabel();

// ---------- Google sign-in ----------
$('logoutBtn').onclick = () => { FS.logout(); S.disp = []; S.maps = []; };
FS.onUser(u => {
  document.body.classList.toggle('authed', !!u); if (!u) return;
  S.admin = ADMIN_EMAILS.includes((u.email || '').toLowerCase());
  const nm = u.displayName || u.email || 'User';
  $('hello').textContent = 'नमस्ते, ' + nm.split(' ')[0] + ' 👋'; $('dname').textContent = nm; $('demail').textContent = u.email || '';
  $('avatar').src = $('davatar').src = u.photoURL || '';
  $('navUpload').hidden = $('tileUpload').hidden = $('upBox').hidden = !S.admin; $('noAdmin').hidden = S.admin;
  loadAll(); // data is read only after sign-in
});
