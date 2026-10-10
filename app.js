/* =========================================================
   Cover Letter Generator - static version (no Firebase, no login)
   - Mapping:  mapping.json in the repo (or a mapping Excel stored on this device)
   - Dispatch: Excel files chosen in the browser, stored in this browser only
========================================================= */

import * as UI from './js/ui.js';
import { generateDocx } from './js/docx-generator.js';
import { buildAmsReport, normPS } from './js/ams-report.js';
import { cacheGet, cacheSet, cacheClear } from './js/cache.js';

const $ = UI.$;

const S = {
  mappings: [],
  mappingSource: '',
  mappingJson: null,
  files: [],          // loaded Dispatch files [{year, name, rows, t}]
  dispatches: [],     // all rows of all loaded files
  tree: {},           // year -> district -> PS -> [GP]
  unmapped: { count: 0, names: [] },
  rtype: 'ybc',
  year: '',
  district: '',
  ps: '',
  gp: '',
  selectedRows: [],
  dispatchPreview: [],
  dispatchFileName: '',
  generatedBlob: null,
  generatedFilename: '',
  amsTicked: new Set(),
  amsInclude: 1,
  amsSearch: ''
};

const REPORTS = {
  ams: { tpl: 'ams-report.docx', prefix: 'AMS_Report' },
  ybc: { tpl: 'covering-letter-template.docx', prefix: 'Cover_Letter' },
  intim: { tpl: 'intimation-report.docx', prefix: 'Intimation_Report' }
};

const KEY_FILES = 'dispatch_files_v2';
const KEY_MAPPING = 'mapping_local_v1';
const WL_KEY = 'jamidara_ams_watchlist_v1';

/* =========================================================
   HELPERS (names of fields in the Excel / mapping rows)
========================================================= */

function safe(value) {
  return String(value ?? '').trim();
}

function norm(value) {
  return safe(value)
    .toLowerCase()
    .replace(/gram\s+panchayat/gi, '')
    .replace(/panchayat\s+samiti/gi, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function normalizeYear(value) {

  let v = safe(value);

  if (!v) return '';

  v = v
    .replace(/[–—]/g, '-')
    .replace(/\s+/g, '')
    .trim();


  const match = v.match(/^(\d{4})-(\d{2}|\d{4})$/);

  if (match) {

    const first = match[1];
    const second = match[2];

    return second.length === 4
      ? `${first}-${second.slice(-2)}`
      : `${first}-${second}`;

  }


  return v;

}

function uniqueSorted(values) {

  return [
    ...new Set(
      values
        .map(safe)
        .filter(Boolean)
    )
  ].sort((a, b) =>
    a.localeCompare(
      b,
      undefined,
      {
        numeric: true,
        sensitivity: 'base'
      }
    )
  );

}

function getDispatchYear(row) {

  return safe(
    row.Year ??
    row.year ??
    row['Year'] ??
    ''
  );

}

function getDispatchUnitId(row) {

  return safe(
    row['Unit ID'] ??
    row['Unit Id'] ??
    row.UnitID ??
    row.unitId ??
    row['UnitID'] ??
    ''
  );

}

function getDispatchDistrict(row) {

  return safe(
    row['District Name'] ??
    row.District ??
    row.DIST_EN ??
    row.District_EN ??
    row['District'] ??
    ''
  );

}

function getDispatchPS(row) {

  return safe(
    row['Parent Name'] ??
    row.PS ??
    row.PS_EN ??
    row['PS'] ??
    ''
  );

}

function getDispatchGP(row) {

  return safe(
    row['Unit Name'] ??
    row.GP ??
    row.GP_EN ??
    row['GP'] ??
    row['Gram Panchayat'] ??
    ''
  );

}

function getMappingDistrict(row) {

  return safe(
    row.DIST_EN ??
    row.District_EN ??
    row.District ??
    row.DISTRICT ??
    row['District EN'] ??
    ''
  );

}

function getMappingDistrictHindi(row) {

  return safe(
    row.DIST_HI ??
    row.District_HI ??
    row.DISTRICT_HI ??
    row['District HI'] ??
    ''
  );

}

function getMappingPS(row) {

  return safe(
    row.PS_EN ??
    row.Panchayat_Samiti_EN ??
    row.PS ??
    row['PS EN'] ??
    row['Panchayat Samiti'] ??
    ''
  );

}

function getMappingPSHindi(row) {

  return safe(
    row.PS_HI ??
    row.Panchayat_Samiti_HI ??
    row['PS HI'] ??
    ''
  );

}

function getMappingGP(row) {

  return safe(
    row.GP_EN ??
    row.GP ??
    row['GP EN'] ??
    row['Gram Panchayat'] ??
    row['GP Name'] ??
    ''
  );

}

function getMappingGPHindi(row) {

  return safe(
    row.GP_HI ??
    row.GPHI ??
    row['GP HI'] ??
    row['Gram Panchayat HI'] ??
    ''
  );

}

function dispatchNumberFrom(value) {
  // BIKANER-4580-LFAD -> 4580  (same as CoverGeneretor.py)
  const s = safe(value);
  let m = s.match(/-(\d+)-/);
  if (m) return m[1];
  m = s.match(/\b(\d+)\b/);
  return m ? m[1] : s;
}

function formatDateDMY(value) {
  // Report Approval Date -> dd.mm.yy  (24/09/2026 -> 24.09.26)
  if (value === null || value === undefined || safe(value) === '') return '';
  const pad = n => String(n).padStart(2, '0');
  if (typeof value === 'number') { // Excel serial date
    const d = new Date(Math.round((value - 25569) * 86400 * 1000));
    return `${pad(d.getUTCDate())}.${pad(d.getUTCMonth() + 1)}.${String(d.getUTCFullYear()).slice(-2)}`;
  }
  if (value instanceof Date) {
    return `${pad(value.getDate())}.${pad(value.getMonth() + 1)}.${String(value.getFullYear()).slice(-2)}`;
  }
  const m = safe(value).match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2,4})$/);
  if (m) return `${pad(m[1])}.${pad(m[2])}.${m[3].slice(-2)}`;
  return safe(value);
}

function paraPlusOne(value) {
  const s = safe(value);
  if (!s) return '';
  const n = parseFloat(s);
  return Number.isNaN(n) ? s : String(Math.trunc(n) + 1);
}

function prepareWordData(row) {

  const mapping = row.__mapping || findMappingForRow(row) || {};

  const gpEn = getMappingGP(mapping);
  const gpHi = getMappingGPHindi(mapping);
  const psEn = getMappingPS(mapping);
  const psHi = getMappingPSHindi(mapping);
  const distEn = getMappingDistrict(mapping);
  const distHi = getMappingDistrictHindi(mapping);

  const gpEngShort = gpEn.replace(/^\s*gram\s+panchayat\s+/i, '').trim();

  const dispatchName = safe(row['Dispatch Name']);
  const paraOriginal = safe(row['Converted to Para in Nos']);

  return {
    YEAR: getDispatchYear(row),
    GP_NAME_EN: gpEn,
    GP_NAME_ENG: gpEngShort,
    // Hindi naam ke aage English: भुरासर (Bhurasar)
    GP_NAME_HI: gpHi ? `${gpHi} (${gpEngShort})` : gpEngShort,
    PS_NAME_EN: psEn,
    PS_NAME_HI: psHi || psEn,
    DISTRICT_EN: distEn,
    DISTRICT_HI: distHi || distEn,
    DISPATCH_NO: dispatchNumberFrom(dispatchName),
    DISPATCH_NAME: dispatchName,
    DATE: formatDateDMY(row['Report Approval Date']),
    PARA_COUNT: paraPlusOne(paraOriginal),
    OFFICE_NAME: distHi || distEn,
    DIVISION_NAME: distHi || distEn,
    CONSTITUTION_OBJECTION: '0',
    SERIOUS_OBJECTION: '0',
    PARA_BREAKUP: paraOriginal
  };

}

function renderRowsPreview(rows) {

  const box = $('rowsPreview');

  if (!box) return;

  if (S.rtype === 'ams' || !rows || !rows.length) {
    box.innerHTML = '';
    return;
  }

  const body = rows.map((row, i) => {
    const d = prepareWordData(row);
    return `<tr>` +
      `<td>${i + 1}</td>` +
      `<td>${UI.esc(d.GP_NAME_HI)}</td>` +
      `<td>${UI.esc(d.DISPATCH_NO)}</td>` +
      `<td>${UI.esc(d.DATE)}</td>` +
      `<td>${UI.esc(d.PARA_COUNT)}${d.PARA_BREAKUP ? ` <small>(${UI.esc(d.PARA_BREAKUP)}+1)</small>` : ''}</td>` +
      `</tr>`;
  }).join('');

  box.innerHTML =
    `<div class="table-wrap"><table class="pv">` +
    `<thead><tr><th>#</th><th>Gram Panchayat</th><th>Dispatch No</th><th>Date</th><th>Para</th></tr></thead>` +
    `<tbody>${body}</tbody></table></div>`;

}

function titleCase(t) {
  return String(t || '').toLowerCase().replace(/\b\w/g, c => c.toUpperCase());
}

async function shareFileWhatsApp(blob, filename, text) {
  if (!blob) {
    UI.showToast('Please generate the report first.', 'error');
    return;
  }
  // Mobile: the Word file is shared directly (choose WhatsApp)
  try {
    const file = new File([blob], filename, { type: blob.type });
    if (navigator.canShare && navigator.canShare({ files: [file] })) {
      await navigator.share({ files: [file], text });
      return;
    }
  } catch (e) {
    if (e && e.name === 'AbortError') return;
  }
  // Fallback (desktop): download the file, then open WhatsApp and attach it there
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 3000);
  UI.showToast('File downloaded. Attach it in WhatsApp to send.', 'success');
  window.open('https://wa.me/?text=' + encodeURIComponent(text), '_blank');
}

async function readExcelFile(file) {

  if (!file) {

    throw new Error(
      'Please select an Excel file.'
    );

  }


  if (!window.XLSX) {

    throw new Error(
      'Excel library not loaded. Please refresh the page.'
    );

  }


  const buffer =
    await file.arrayBuffer();


  const workbook =
    XLSX.read(
      buffer,
      {
        type: 'array'
      }
    );


  if (
    !workbook.SheetNames.length
  ) {

    throw new Error(
      'Excel file has no worksheet.'
    );

  }


  const sheet =
    workbook.Sheets[
      workbook.SheetNames[0]
    ];


  return XLSX.utils.sheet_to_json(
    sheet,
    {
      defval: ''
    }
  );

}

function makeTable(rows) {

  if (!rows?.length) {
    return '<p>No data.</p>';
  }


  const headers =
    uniqueSorted(
      rows.flatMap(
        row =>
          Object.keys(row)
      )
    );


  const head =
    headers
      .map(
        h =>
          `<th>${UI.esc(h)}</th>`
      )
      .join('');


  const body =
    rows
      .map(row => {

        const cells =
          headers
            .map(
              h =>
                `<td>${UI.esc(
                  row[h]
                )}</td>`
            )
            .join('');


        return `<tr>${cells}</tr>`;

      })
      .join('');


  return `
    <div class="table-wrap">
      <table>
        <thead>
          <tr>${head}</tr>
        </thead>
        <tbody>${body}</tbody>
      </table>
    </div>
  `;

}

function showError(error) {

  console.error(error);

  const message =
    UI.friendly
      ? UI.friendly(error)
      : (error?.message || String(error));

  UI.showToast(message, 'error');

}


/* =========================================================
   MAPPING
========================================================= */

const MAP_COLS = ['GP_EN', 'GP_HI', 'PS_EN', 'PS_HI', 'DIST_EN', 'DIST_HI'];

function expandMapping(json) {
  if (Array.isArray(json)) return json;
  if (json && Array.isArray(json.cols) && Array.isArray(json.rows)) {
    return json.rows.map(r => Object.fromEntries(json.cols.map((c, i) => [c, r[i]])));
  }
  return [];
}

function compactMapping(rows) {
  const out = [];
  rows.forEach(r => {
    const row = [
      getMappingGP(r), getMappingGPHindi(r),
      getMappingPS(r), getMappingPSHindi(r),
      getMappingDistrict(r), getMappingDistrictHindi(r)
    ];
    if (row[0] && row[2] && row[4]) out.push(row);
  });
  return { v: 1, updated: Date.now(), cols: MAP_COLS, rows: out };
}

let mapIdx = new Map();
let mapIdx2 = new Map();
let mapCache = new WeakMap();

function buildMapIndex() {
  mapIdx = new Map();
  mapIdx2 = new Map();
  mapCache = new WeakMap();
  S.mappings.forEach(m => {
    const d = norm(getMappingDistrict(m));
    const p = norm(getMappingPS(m));
    const g = norm(getMappingGP(m));
    const k = `${d}|${p}|${g}`;
    if (!mapIdx.has(k)) mapIdx.set(k, m);
    const k2 = `${p}|${g}`;
    mapIdx2.set(k2, mapIdx2.has(k2) ? null : m); // null = ambiguous
  });
}

function findMappingForRow(row) {
  if (mapCache.has(row)) return mapCache.get(row);
  const d = norm(getDispatchDistrict(row));
  const p = norm(getDispatchPS(row));
  const g = norm(getDispatchGP(row));
  let found = null;
  if (d || p || g) {
    found = mapIdx.get(`${d}|${p}|${g}`) || mapIdx2.get(`${p}|${g}`) || null;
  }
  mapCache.set(row, found);
  return found;
}

async function loadMapping() {
  S.mappings = [];
  S.mappingSource = '';
  const local = await cacheGet(KEY_MAPPING);
  if (local && local.json) {
    S.mappings = expandMapping(local.json);
    S.mappingSource = `This device (${local.name || 'Excel'})`;
  } else {
    try {
      const r = await fetch('mapping.json', { cache: 'no-cache' });
      if (r.ok) {
        S.mappings = expandMapping(await r.json());
        S.mappingSource = 'mapping.json';
      }
    } catch (e) { /* no mapping.json yet */ }
  }
  buildMapIndex();
}

/* =========================================================
   DISPATCH FILES (stored in this browser)
========================================================= */

async function loadDispatchFiles() {
  const d = await cacheGet(KEY_FILES);
  S.files = Array.isArray(d?.files) ? d.files : [];
}

async function saveDispatchFiles() {
  await cacheSet(KEY_FILES, { files: S.files });
}

function rebuildDispatches() {
  S.dispatches = S.files.flatMap(f => f.rows);
  rebuildTree();
}

function rebuildTree() {
  const sets = {};
  let unmapped = 0;
  const names = [];
  S.dispatches.forEach(row => {
    if (safe(row['Parent Name']) === '') return; // Panchayat Samiti rows
    const y = normalizeYear(getDispatchYear(row));
    if (!y) return;
    const m = findMappingForRow(row);
    if (!m) {
      unmapped++;
      if (names.length < 15) {
        names.push(`${getDispatchDistrict(row)} / ${getDispatchPS(row)} / ${getDispatchGP(row)}`);
      }
      return;
    }
    const d = getMappingDistrict(m), p = getMappingPS(m), g = getMappingGP(m);
    (((sets[y] ??= {})[d] ??= {})[p] ??= new Set()).add(g);
  });
  const tree = {};
  Object.keys(sets).forEach(y => {
    tree[y] = {};
    Object.keys(sets[y]).forEach(d => {
      tree[y][d] = {};
      Object.keys(sets[y][d]).forEach(p => { tree[y][d][p] = uniqueSorted([...sets[y][d][p]]); });
    });
  });
  S.tree = tree;
  S.unmapped = { count: unmapped, names };
}

/* =========================================================
   DROPDOWNS
========================================================= */

function treeYear() {
  return S.tree[normalizeYear(S.year)] || {};
}

function getAvailableYears() {
  return uniqueSorted(S.dispatches.map(r => normalizeYear(getDispatchYear(r))));
}

function getAvailableDistricts() {
  return uniqueSorted(Object.keys(treeYear()));
}

function getAvailablePS() {
  const ty = treeYear();
  const d = Object.keys(ty).find(k => norm(k) === norm(S.district));
  return d === undefined ? [] : uniqueSorted(Object.keys(ty[d]));
}

function getAvailableGP() {
  const ty = treeYear();
  const d = Object.keys(ty).find(k => norm(k) === norm(S.district));
  if (d === undefined) return [];
  const p = Object.keys(ty[d]).find(k => norm(k) === norm(S.ps));
  return p === undefined ? [] : ty[d][p];
}

function fillSel(id, values, placeholder, current) {
  const sel = $(id);
  if (!sel) return '';
  UI.fillSelect(sel, values.map(v => ({ v, t: v })), placeholder);
  const cur = safe(current);
  const found = cur ? values.find(v => norm(v) === norm(cur)) : undefined;
  sel.value = found ?? '';
  return found ?? '';
}

function populateYear() {
  const years = getAvailableYears();
  let keep = S.year;
  if (!safe(keep) && years.length === 1) keep = years[0]; // only one year loaded: select it
  S.year = fillSel('year', years, 'Select Year', keep);
}
function populateDistrict() {
  const list = getAvailableDistricts();
  let keep = S.district;
  if (!safe(keep) && list.length === 1) keep = list[0];
  S.district = fillSel('dist', list, 'Select District', keep);
}
function populatePS() {
  S.ps = fillSel('ps', getAvailablePS(), 'Select Panchayat Samiti', S.ps);
}
function populateGP() {
  S.gp = fillSel('gp', getAvailableGP(), 'All Gram Panchayats', S.gp);
}

/* =========================================================
   COVER LETTER
========================================================= */

function getSelectedRows() {
  const y = normalizeYear(S.year);
  const d = norm(S.district);
  const p = norm(S.ps);
  const g = norm(S.gp);
  if (!y || !d || !p) {
    S.selectedRows = [];
    return [];
  }
  const out = [];
  S.dispatches.forEach(row => {
    if (safe(row['Parent Name']) === '') return;
    if (normalizeYear(getDispatchYear(row)) !== y) return;
    const m = findMappingForRow(row);
    if (!m) return;
    if (norm(getMappingDistrict(m)) !== d || norm(getMappingPS(m)) !== p) return;
    if (g && norm(getMappingGP(m)) !== g) return;
    out.push({ ...row, __mapping: m });
  });
  S.selectedRows = out;
  return out;
}

function diagText() {
  try {
    const y = normalizeYear(S.year);
    const yr = S.dispatches.filter(r => normalizeYear(getDispatchYear(r)) === y);
    const gp = yr.filter(r => safe(r['Parent Name']) !== '');
    const mapped = gp.filter(r => findMappingForRow(r));
    return `Details: dispatch rows for this year: ${yr.length} (Gram Panchayat rows: ${gp.length}, with mapping: ${mapped.length}); mapping rows loaded: ${S.mappings.length}.`;
  } catch (e) {
    return '';
  }
}

function updateDataNotice() {
  const box = $('dataNotice');
  if (!box) return;
  if (!S.dispatches.length) {
    box.innerHTML = '<div class="warn">No Dispatch file is loaded yet. <button type="button" class="primary" data-view="files" style="margin-top:8px">Open Data Files</button></div>';
  } else if (!S.mappings.length) {
    box.innerHTML = '<div class="warn">mapping.json was not found, so Hindi names are missing. <button type="button" class="primary" data-view="files" style="margin-top:8px">Open Data Files</button></div>';
  } else {
    box.innerHTML = '';
  }
}

function updateInfo() {
  const info = $('info');
  const warn = $('warn');
  updateDataNotice();

  if (S.rtype === 'ams') {
    if (info) info.innerHTML = '';
    if (warn) warn.innerHTML = '';
    renderRowsPreview([]);
    renderAmsPs();
    return;
  }

  const rows = getSelectedRows();
  renderRowsPreview(rows);

  if (info) {
    info.textContent = (!S.year || !S.district || !S.ps)
      ? 'Select Year, District and Panchayat Samiti.'
      : `${rows.length} dispatch record(s) available`;
  }

  if (warn) {
    warn.textContent = '';
    if (S.year && S.district && S.ps && !rows.length) {
      warn.textContent = 'No dispatch record found for selected filters. ' + diagText();
    }
  }
}

function showResult(blob, filename, label, waText) {
  S.generatedBlob = blob;
  S.generatedFilename = filename;
  const result = $('result');
  const rname = $('rname');
  if (rname) rname.textContent = label || filename;
  if (result) { result.style.display = ''; result.hidden = false; }

  const dl = $('dlBtn');
  if (dl) {
    dl.onclick = () => {
      if (!S.generatedBlob) return;
      const url = URL.createObjectURL(S.generatedBlob);
      const a = document.createElement('a');
      a.href = url;
      a.download = S.generatedFilename;
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 3000);
    };
  }
  const wa = $('waBtn');
  if (wa) wa.onclick = () => shareFileWhatsApp(S.generatedBlob, S.generatedFilename, waText);
}

async function generateReport() {
  try {
    updateInfo();
    const rows = getSelectedRows();

    if (!S.dispatches.length) throw new Error('No Dispatch file is loaded. Add one in Data Files.');
    if (!S.year) throw new Error('Please select Year.');
    if (!S.district) throw new Error('Please select District.');
    if (!S.ps) throw new Error('Please select Panchayat Samiti.');
    if (!rows.length) throw new Error('No dispatch record found for selected filters. ' + diagText());

    const report = REPORTS[S.rtype];
    if (!report) throw new Error('Invalid report type.');
    if (S.rtype !== 'ybc') {
      throw new Error(`Template for "${S.rtype}" is not available yet. For now select "Cover Letter".`);
    }

    UI.showToast(`Generating ${rows.length} document(s)...`);

    const blob = await generateDocx(rows.map(prepareWordData), report.tpl);
    if (!blob) throw new Error('Word file generation failed.');

    const filename = `${report.prefix}_${safe(S.year)}_${safe(S.district)}_${safe(S.ps)}.docx`;

    showResult(
      blob, filename, filename,
      `Cover Letter\nYear: ${S.year}\nDistrict: ${S.district}\nPanchayat Samiti: ${S.ps}\nGram Panchayat: ${S.gp || 'All Gram Panchayats'}`
    );

    UI.showToast('Word file generated successfully.', 'success');
  } catch (error) {
    console.error('GENERATION ERROR:', error);
    showError(error);
  }
}

function setReportType(type) {
  if (!REPORTS[type]) return;
  S.rtype = type;
  const isAms = type === 'ams';
  const ls = $('letterSel'), ab = $('amsBox'), gb = $('genBtn'), rs = $('result');
  if (ls) ls.hidden = isAms;
  if (ab) ab.hidden = !isAms;
  if (gb) gb.textContent = isAms ? 'Generate AMS Report' : 'Generate Cover Letter';
  if (rs) { rs.hidden = true; rs.style.display = 'none'; }
  document.querySelectorAll('#rtypes [data-t]').forEach(b => {
    const on = b.dataset.t === type;
    b.classList.toggle('on', on);
    b.classList.toggle('active', on);
  });
  updateInfo();
}

/* =========================================================
   AMS REPORT + WATCH LIST
========================================================= */

function loadWatch() {
  try {
    const a = JSON.parse(localStorage.getItem(WL_KEY) || '[]');
    return Array.isArray(a) ? a : [];
  } catch (e) { return []; }
}

function saveWatch(list) {
  try { localStorage.setItem(WL_KEY, JSON.stringify(list)); return true; } catch (e) { return false; }
}

function amsRecords() {
  const y = normalizeYear(S.year);
  return S.dispatches.filter(r => !y || normalizeYear(getDispatchYear(r)) === y);
}

function amsPsOptions() {
  const map = new Map();
  amsRecords().forEach(r => {
    const parent = safe(r['Parent Name']);
    const isPS = parent === '';
    const name = normPS(isPS ? r['Unit Name'] : parent);
    if (!name) return;
    const o = map.get(name) || { name, dist: safe(r['District Name']), gp: 0 };
    if (!isPS) o.gp++;
    map.set(name, o);
  });
  return [...map.values()].sort((a, b) => (a.name < b.name ? -1 : 1));
}

function renderAmsSaved() {
  const box = $('amsSaved');
  if (!box) return;
  const saved = loadWatch();
  box.innerHTML = saved.length
    ? saved.map(n => `<span class="chip">${UI.esc(titleCase(n))} <b data-rm="${UI.esc(n)}" title="Remove">×</b></span>`).join('')
    : '<span class="muted">Your watch list is empty. Tick Panchayat Samitis above and press Save.</span>';
}

function renderAmsPs() {
  const box = $('amsPs');
  if (!box) return;
  if (!S.amsTicked.size && !S.amsInit) {
    S.amsTicked = new Set(loadWatch());
    S.amsInit = true;
  }
  if (!S.year) {
    box.innerHTML = '<div style="padding:14px" class="muted">Please select a Year first.</div>';
    renderAmsSaved();
    return;
  }
  const q = norm(S.amsSearch);
  const list = amsPsOptions().filter(o => !q || o.name.toLowerCase().includes(q));
  box.innerHTML = list.length
    ? list.map(o =>
        `<label><input type="checkbox" data-n="${UI.esc(o.name)}" ${S.amsTicked.has(o.name) ? 'checked' : ''}>` +
        `<span>PS ${UI.esc(titleCase(o.name))}<small>${UI.esc(o.dist)}${o.gp ? ` · ${o.gp} GP` : ''}</small></span></label>`
      ).join('')
    : '<div style="padding:14px" class="muted">No Panchayat Samiti found for this Year.</div>';
  const c = $('amsCount');
  if (c) c.textContent = `${S.amsTicked.size} ticked`;
  renderAmsSaved();
}

async function generateAms() {
  try {
    if (!S.year) throw new Error('Please select Year.');
    const saved = loadWatch();
    if (!saved.length) throw new Error('Your watch list is empty. Tick Panchayat Samitis and press "Save Watch List".');
    const recs = amsRecords();
    if (!recs.length) throw new Error('No dispatch data found for this Year. Add a Dispatch file in Data Files.');

    let res;
    try {
      res = buildAmsReport(recs, saved, S.amsInclude, { year: safe(S.year) });
    } catch (e) {
      if (e && e.message === 'NO_PS') {
        throw new Error('The Panchayat Samitis in your watch list were not found for this Year. Check the Year or change the watch list.');
      }
      throw e;
    }

    const have = new Set(amsPsOptions().map(o => o.name));
    const missing = saved.filter(n => !have.has(normPS(n)));
    const stamp = new Date().toISOString().slice(0, 10);
    const filename = `AMS_Report_${safe(S.year)}_${stamp}.docx`;

    showResult(
      res.blob, filename,
      `${filename} · ${res.psCount} PS · ${res.totalGp} GP` +
        (missing.length ? ` · Not found for this Year: ${missing.map(titleCase).join(', ')}` : ''),
      `AMS Report\nYear: ${S.year}\nPS: ${res.psCount}\nGP: ${res.totalGp}`
    );
    UI.showToast('AMS report ready.', 'success');
  } catch (error) {
    console.error('AMS ERROR:', error);
    showError(error);
  }
}

function initAms() {
  S.amsTicked = new Set(loadWatch());
  S.amsInit = true;

  $('amsSearch')?.addEventListener('input', e => {
    S.amsSearch = e.target.value;
    renderAmsPs();
  });

  $('amsPs')?.addEventListener('change', e => {
    const t = e.target;
    if (!t?.dataset || t.dataset.n === undefined) return;
    if (t.checked) S.amsTicked.add(t.dataset.n); else S.amsTicked.delete(t.dataset.n);
    const c = $('amsCount');
    if (c) c.textContent = `${S.amsTicked.size} ticked`;
  });

  const setVisible = on => {
    const q = norm(S.amsSearch);
    amsPsOptions()
      .filter(o => !q || o.name.toLowerCase().includes(q))
      .forEach(o => (on ? S.amsTicked.add(o.name) : S.amsTicked.delete(o.name)));
    renderAmsPs();
  };
  $('amsAll')?.addEventListener('click', () => setVisible(true));
  $('amsNone')?.addEventListener('click', () => setVisible(false));

  $('amsSave')?.addEventListener('click', () => {
    const list = [...S.amsTicked].sort();
    if (!list.length) { UI.showToast('Please tick at least one Panchayat Samiti.', 'error'); return; }
    if (saveWatch(list)) {
      renderAmsSaved();
      UI.showToast(`Watch list saved (${list.length} PS).`, 'success');
    } else {
      UI.showToast('The watch list could not be saved (browser storage is blocked).', 'error');
    }
  });

  $('amsSaved')?.addEventListener('click', e => {
    const n = e.target?.dataset?.rm;
    if (n === undefined) return;
    saveWatch(loadWatch().filter(x => x !== n));
    S.amsTicked.delete(n);
    renderAmsPs();
  });

  document.querySelectorAll('#amsMode [data-m]').forEach(b => {
    b.addEventListener('click', () => {
      S.amsInclude = +b.dataset.m;
      document.querySelectorAll('#amsMode [data-m]').forEach(x => x.classList.toggle('on', x === b));
    });
  });
}

/* =========================================================
   DATA FILES PAGE
========================================================= */

function detectYearFromRows(rows) {
  for (const r of rows) {
    const m = String(r['Audit Party No'] ?? '').match(/\d{4}-\d{2}\b/);
    if (m) return m[0];
  }
  return '';
}

async function onDispatchFileChosen() {
  try {
    const file = $('file')?.files?.[0];
    if (!file) return;
    const rows = await readExcelFile(file);
    if (!rows.length) throw new Error('The file has no data rows.');

    S.dispatchPreview = rows;
    S.dispatchFileName = file.name;

    const yr = $('upYear');
    if (yr && !safe(yr.value)) yr.value = detectYearFromRows(rows);

    const info = $('fileInfo');
    if (info) info.textContent = `${rows.length} rows read from ${file.name}`;
    const preview = $('preview');
    if (preview) preview.innerHTML = makeTable(rows.slice(0, 20));
    const btn = $('useLocalBtn');
    if (btn) btn.disabled = false;
  } catch (error) {
    showError(error);
  }
}

async function useDispatchFile() {
  try {
    if (!S.dispatchPreview.length) throw new Error('Please choose a Dispatch Excel file first.');
    const year = safe($('upYear')?.value);
    if (!year) throw new Error('Please enter the Year for this file (for example 2026-27).');

    const rows = S.dispatchPreview
      .filter(r => safe(r['Unit Name']) !== '' || getDispatchUnitId(r) !== '')
      .map(r => ({ ...r, Year: year }));
    if (!rows.length) throw new Error('No rows with a Unit Name were found in this file.');

    const y = normalizeYear(year);
    S.files = S.files.filter(f => normalizeYear(f.year) !== y);
    S.files.push({ year, name: S.dispatchFileName || 'Dispatch file', rows, t: Date.now() });
    await saveDispatchFiles();

    rebuildDispatches();
    afterDataChange();

    S.dispatchPreview = [];
    $('file').value = '';
    const btn = $('useLocalBtn');
    if (btn) btn.disabled = true;
    const info = $('fileInfo');
    if (info) info.textContent = '';
    const preview = $('preview');
    if (preview) preview.innerHTML = '';

    UI.showToast(`${rows.length} rows loaded for ${year}.`, 'success');
  } catch (error) {
    showError(error);
  }
}

function renderLoadedFiles() {
  const box = $('loadedFiles');
  if (!box) return;
  if (!S.files.length) {
    box.innerHTML = '<p class="muted">No Dispatch file is loaded yet.</p>';
    return;
  }
  box.innerHTML = '<div class="lab">Loaded files</div>' + S.files.map(f =>
    `<div class="fileitem"><div><b>${UI.esc(f.year)}</b><span>${UI.esc(f.name)} · ${f.rows.length} rows · ${UI.esc(new Date(f.t).toLocaleString())}</span></div>` +
    `<button type="button" data-rmfile="${UI.esc(f.year)}">Remove</button></div>`
  ).join('');
}

async function removeDispatchFile(year) {
  const y = normalizeYear(year);
  S.files = S.files.filter(f => normalizeYear(f.year) !== y);
  await saveDispatchFiles();
  rebuildDispatches();
  afterDataChange();
  UI.showToast(`Removed ${year}.`, 'success');
}

async function onMappingFileChosen() {
  try {
    const file = $('mFile')?.files?.[0];
    if (!file) return;
    const rows = await readExcelFile(file);
    const json = compactMapping(rows);
    if (!json.rows.length) {
      throw new Error('No mapping rows found. The file needs the columns GP_EN, PS_EN and DIST_EN (and GP_HI, PS_HI, DIST_HI).');
    }
    S.mappingJson = json;
    S.mappingFileName = file.name;
    const info = $('mInfo');
    if (info) info.textContent = `${json.rows.length} mapping rows ready from ${file.name}`;
    ['mDlBtn', 'mUseBtn'].forEach(id => { const b = $(id); if (b) b.disabled = false; });
  } catch (error) {
    showError(error);
  }
}

function downloadMappingJson() {
  if (!S.mappingJson) return;
  const blob = new Blob([JSON.stringify(S.mappingJson)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'mapping.json';
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 3000);
  UI.showToast('mapping.json downloaded. Upload it to your GitHub repo.', 'success');
}

async function useMappingHere() {
  if (!S.mappingJson) return;
  await cacheSet(KEY_MAPPING, { name: S.mappingFileName || 'Excel', json: S.mappingJson, t: Date.now() });
  await loadMapping();
  rebuildTree();
  afterDataChange();
  UI.showToast('Mapping is now used on this device.', 'success');
}

async function clearLocalMapping() {
  await cacheClear(KEY_MAPPING);
  await loadMapping();
  rebuildTree();
  afterDataChange();
  UI.showToast('Local mapping removed.', 'success');
}

function renderMapStatus() {
  const box = $('mapStatus');
  if (!box) return;
  const districts = new Set(S.mappings.map(m => norm(getMappingDistrict(m))));
  const ps = new Set(S.mappings.map(m => `${norm(getMappingDistrict(m))}|${norm(getMappingPS(m))}`));
  const rows = [
    ['Status', S.mappings.length ? '<span class="pill good">Loaded</span>' : '<span class="pill wait">Not found</span>'],
    ['Source', UI.esc(S.mappingSource || '—')],
    ['Gram Panchayats', S.mappings.length],
    ['Panchayat Samitis', ps.size],
    ['Districts', districts.size]
  ];
  box.innerHTML = rows.map(([k, v]) => `<div><span>${k}</span><b>${v}</b></div>`).join('');
  const clear = $('clearMapBtn');
  if (clear) clear.hidden = !S.mappingSource.startsWith('This device');

  const un = $('unmapped');
  if (un) {
    un.innerHTML = S.unmapped.count
      ? `<details class="help"><summary>${S.unmapped.count} Dispatch row(s) have no mapping</summary><p>These Gram Panchayats are not in the dropdowns because their names were not found in the mapping. First ones:<br>${S.unmapped.names.map(UI.esc).join('<br>')}</p></details>`
      : '';
  }
}

function renderSetup() {
  const box = $('setupBox');
  if (!box) return;
  const rows = S.files.reduce((a, f) => a + f.rows.length, 0);
  const item = (ok, title, text) =>
    `<div class="setrow ${ok ? 'ok' : 'todo'}"><i>${ok ? '✓' : '!'}</i><div><b>${title}</b><span>${text}</span></div>` +
    `<button type="button" data-view="files">${ok ? 'Manage' : 'Add'}</button></div>`;
  box.innerHTML =
    item(S.files.length > 0, 'Dispatch file',
      S.files.length ? `${S.files.map(f => f.year).join(', ')} · ${rows} rows` : 'Not added yet') +
    item(S.mappings.length > 0, 'Mapping',
      S.mappings.length ? `${UI.esc(S.mappingSource)} · ${S.mappings.length} Gram Panchayats` : 'mapping.json not found');
}

function updateStats() {
  const set = (id, v) => { const e = $(id); if (e) e.textContent = v; };
  const years = Object.keys(S.tree);
  const d = new Set(), p = new Set();
  years.forEach(y => Object.keys(S.tree[y]).forEach(dn => {
    d.add(norm(dn));
    Object.keys(S.tree[y][dn]).forEach(pn => p.add(`${norm(dn)}|${norm(pn)}`));
  }));
  set('stY', years.length || '–');
  set('stD', d.size || '–');
  set('stM', p.size || '–');
}

/* =========================================================
   DISPATCH DATA PAGE
========================================================= */

let dataViewBound = false;

function setupDataFilters() {
  const fill = (id, values, placeholder) => {
    const el = $(id);
    if (!el) return;
    const prev = el.value;
    UI.fillSelect(el, values.map(v => ({ v, t: v })), placeholder);
    if (prev && values.includes(prev)) el.value = prev;
  };

  const allDistricts = uniqueSorted(Object.values(S.tree).flatMap(t => Object.keys(t)));
  fill('fYear', Object.keys(S.tree).sort(), 'All Years');
  fill('fDist', allDistricts, 'All Districts');
  fillDataPs();
  fill('fStatus', uniqueSorted(S.dispatches.map(r => r['Report Status'])), 'All Statuses');

  if (dataViewBound) return;
  dataViewBound = true;
  $('fDist')?.addEventListener('change', () => { fillDataPs(); renderDataView(); });
  ['fYear', 'fPs', 'fStatus'].forEach(id => $(id)?.addEventListener('change', renderDataView));
  $('fSearch')?.addEventListener('input', renderDataView);
}

function fillDataPs() {
  const el = $('fPs');
  if (!el) return;
  const prev = el.value;
  const d = norm($('fDist')?.value);
  const list = [];
  Object.values(S.tree).forEach(t => Object.keys(t).forEach(dn => {
    if (d && norm(dn) !== d) return;
    list.push(...Object.keys(t[dn]));
  }));
  const values = uniqueSorted(list);
  UI.fillSelect(el, values.map(v => ({ v, t: v })), 'All Panchayat Samitis');
  if (prev && values.includes(prev)) el.value = prev;
}

function renderDataView() {
  const out = $('dataOut');
  if (!out) return;
  if (!S.dispatches.length) {
    out.innerHTML = '<p class="muted">No Dispatch file is loaded yet. Add one in Data Files.</p>';
    return;
  }
  const y = normalizeYear($('fYear')?.value);
  const d = norm($('fDist')?.value);
  const p = norm($('fPs')?.value);
  const st = norm($('fStatus')?.value);
  const q = norm($('fSearch')?.value);

  const rows = S.dispatches.filter(row => {
    if (y && normalizeYear(getDispatchYear(row)) !== y) return false;
    if (d || p) {
      const m = findMappingForRow(row);
      if (!m) return false;
      if (d && norm(getMappingDistrict(m)) !== d) return false;
      if (p && norm(getMappingPS(m)) !== p) return false;
    }
    if (st && norm(row['Report Status']) !== st) return false;
    if (q && !norm(Object.values(row).join(' ')).includes(q)) return false;
    return true;
  });

  const cols = ['District Name', 'Parent Name', 'Unit Name', 'Dispatch Name', 'Audit Party No', 'Report Status', 'Report Approval Date', 'Converted to Para in Nos'];
  const view = rows.slice(0, 200).map(r => Object.fromEntries(cols.filter(c => c in r).map(c => [c, r[c]])));
  out.innerHTML = `<p>${rows.length} record(s)${rows.length > 200 ? ' (first 200 shown)' : ''}</p>` + makeTable(view);
}

/* =========================================================
   WIRING
========================================================= */

function afterDataChange() {
  populateYear();
  populateDistrict();
  populatePS();
  populateGP();
  setupDataFilters();
  renderLoadedFiles();
  renderMapStatus();
  renderSetup();
  updateStats();
  updateInfo();
}

function setupEvents() {
  document.querySelectorAll('#rtypes [data-t]').forEach(b => {
    b.addEventListener('click', () => setReportType(b.dataset.t));
  });

  $('year')?.addEventListener('change', e => {
    S.year = normalizeYear(e.target.value);
    populateDistrict();
    populatePS();
    populateGP();
    updateInfo();
  });
  $('dist')?.addEventListener('change', e => {
    S.district = e.target.value;
    S.ps = '';
    S.gp = '';
    populatePS();
    populateGP();
    updateInfo();
  });
  $('ps')?.addEventListener('change', e => {
    S.ps = e.target.value;
    S.gp = '';
    populateGP();
    updateInfo();
  });
  $('gp')?.addEventListener('change', e => {
    S.gp = e.target.value;
    updateInfo();
  });

  $('genBtn')?.addEventListener('click', () => (S.rtype === 'ams' ? generateAms() : generateReport()));

  $('file')?.addEventListener('change', onDispatchFileChosen);
  $('useLocalBtn')?.addEventListener('click', useDispatchFile);
  $('loadedFiles')?.addEventListener('click', e => {
    const y = e.target?.dataset?.rmfile;
    if (y !== undefined) removeDispatchFile(y);
  });

  $('mFile')?.addEventListener('change', onMappingFileChosen);
  $('mDlBtn')?.addEventListener('click', downloadMappingJson);
  $('mUseBtn')?.addEventListener('click', useMappingHere);
  $('clearMapBtn')?.addEventListener('click', clearLocalMapping);
}

function setupNavigation() {
  UI.setupNav(view => {
    if (view === 'generate') updateInfo();
    if (view === 'data') { setupDataFilters(); renderDataView(); }
    if (view === 'files') { renderLoadedFiles(); renderMapStatus(); }
    if (view === 'dashboard') { renderSetup(); updateStats(); }
  });
}

async function init() {
  document.body.classList.add('authed');
  setupEvents();
  setupNavigation();
  initAms();
  setReportType('ybc');
  try {
    await loadMapping();
    await loadDispatchFiles();
    rebuildDispatches();
  } catch (error) {
    showError(error);
  }
  afterDataChange();
}

init();
