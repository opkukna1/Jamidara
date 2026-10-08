import {
  loadMappings,
  loadDispatchData,
  googleLogin,
  logout,
  onUser,
  uploadDispatchRows,
  uploadMappingRows
} from './js/firestore.js';

import * as UI from './js/ui.js';
import { generateDocx } from './js/docx-generator.js';

const $ = UI.$;


/* =========================================================
   STATE
========================================================= */

const S = {
  user: null,

  mappings: [],
  dispatches: [],

  rtype: 'ybc',

  year: '',
  district: '',
  ps: '',
  gp: '',

  amsPs: [],
  amsMode: 'all',

  selectedRows: [],

  dispatchPreview: [],
  mappingPreview: []
};


/* =========================================================
   REPORT TYPES
========================================================= */

const REPORTS = {

  ams: {
    tpl: 'ams-report.docx',
    prefix: 'AMS_Report'
  },

  ybc: {
    tpl: 'covering-letter-template.docx',
    prefix: 'Year_Book_Closing'
  },

  intim: {
    tpl: 'intimation-report.docx',
    prefix: 'Intimation_Report'
  }

};


/* =========================================================
   HELPERS
========================================================= */

function safe(v) {
  return String(v ?? '').trim();
}


function norm(v) {
  return safe(v)
    .toLowerCase()
    .replace(/gram\s+panchayat|panchayat\s+samiti/g, '')
    .replace(/[^a-z0-9\u0900-\u097f]+/g, ' ')
    .trim();
}


function unique(arr) {
  return [...new Set(arr.filter(Boolean))];
}


function showError(e) {

  console.error(e);

  const msg = UI.friendly
    ? UI.friendly(e)
    : (e?.message || String(e));

  UI.showToast(msg, 'error');

}


/* =========================================================
   LOGIN / LOGOUT UI
========================================================= */

function showApp() {

  const login = $('login');

  if (login) {
    login.style.display = 'none';
  }

  /*
    IMPORTANT:
    style.css contains:

    body.authed #login { display:none }

    body:not(.authed) .top,
    body:not(.authed) main {
      visibility:hidden
    }

    इसलिए body.authed लगाना जरूरी है।
  */

  document.body.classList.add('authed');


  /*
    सभी views hide करो
    फिर dashboard activate करो
  */

  document
    .querySelectorAll('.view')
    .forEach(v => {
      v.classList.remove('active');
    });


  const dashboard = $('v-dashboard');

  if (dashboard) {
    dashboard.classList.add('active');
  }


  /*
    Header / drawer / main visible
  */

  const top = document.querySelector('.top');
  const drawer = $('drawer');
  const main = document.querySelector('main');

  if (top) {
    top.style.visibility = 'visible';
    top.style.display = '';
  }

  if (drawer) {
    drawer.style.visibility = 'visible';
    drawer.style.display = '';
  }

  if (main) {
    main.style.visibility = 'visible';
    main.style.display = '';
  }

}


function showLogin() {

  const login = $('login');

  document.body.classList.remove('authed');


  if (login) {
    login.style.display = 'grid';
  }


  const top = document.querySelector('.top');
  const main = document.querySelector('main');

  if (top) {
    top.style.visibility = 'hidden';
  }

  if (main) {
    main.style.visibility = 'hidden';
  }

}


/* =========================================================
   USER INFO
========================================================= */

function updateUserUI(user) {

  const name =
    user?.displayName ||
    user?.email ||
    'User';

  const email =
    user?.email ||
    '';

  /*
    Different possible IDs/classes in UI.
  */

  const selectors = [
    '#userName',
    '#profileName',
    '[data-user-name]'
  ];

  selectors.forEach(selector => {

    document
      .querySelectorAll(selector)
      .forEach(el => {
        el.textContent = name;
      });

  });


  document
    .querySelectorAll('#userEmail,[data-user-email]')
    .forEach(el => {
      el.textContent = email;
    });

}


/* =========================================================
   LOAD FIRESTORE DATA
========================================================= */

async function loadData() {

  try {

    UI.showToast('Loading data...');

    const [mappings, dispatches] = await Promise.all([
      loadMappings(),
      loadDispatchData()
    ]);

    S.mappings = Array.isArray(mappings)
      ? mappings
      : [];

    S.dispatches = Array.isArray(dispatches)
      ? dispatches
      : [];


    console.log('Mappings:', S.mappings.length);
    console.log('Dispatches:', S.dispatches.length);


    populateYear();
    populateDistrict();

    refreshGenerator();


    UI.showToast(
      `Data loaded: ${S.dispatches.length} dispatch + ${S.mappings.length} mappings`,
      'success'
    );


  } catch (e) {

    console.error('LOAD DATA ERROR:', e);

    showError(e);

    /*
      App shell फिर भी खुला रहेगा।
    */

  }

}


/* =========================================================
   YEAR
========================================================= */

function getYears() {

  const values = [];

  S.dispatches.forEach(r => {

    const y =
      r.Year ??
      r.year ??
      r['Year'];

    if (safe(y)) {
      values.push(safe(y));
    }

  });


  S.mappings.forEach(r => {

    const y =
      r.Year ??
      r.year ??
      r['Year'];

    if (safe(y)) {
      values.push(safe(y));
    }

  });


  const defaults = [
    '2025-26',
    '2026-27'
  ];

  return unique([
    ...values,
    ...defaults
  ]);
}


function populateYear() {

  const el = $('year');

  if (!el) return;

  const years = getYears();

  const current =
    S.year ||
    years[0] ||
    '2025-26';

  S.year = current;


  UI.fillSelect(
    el,
    years.map(v => ({
      v,
      t: v
    })),
    'Select Year'
  );


  el.value = current;

}


/* =========================================================
   DISTRICT
========================================================= */

function districtFromMapping(r) {

  return safe(
    r.DIST_EN ??
    r.District ??
    r.DISTRICT ??
    r['District EN']
  );

}


function districtFromDispatch(r) {

  return safe(
    r.District ??
    r.DIST_EN ??
    r['District']
  );

}


function getDistricts() {

  const values = [];

  S.mappings.forEach(r => {
    const v = districtFromMapping(r);
    if (v) values.push(v);
  });

  S.dispatches.forEach(r => {
    const v = districtFromDispatch(r);
    if (v) values.push(v);
  });

  return unique(values).sort();

}


function populateDistrict() {

  const el = $('dist');

  if (!el) return;

  const districts = getDistricts();

  let selected =
    S.district ||
    districts.find(x => norm(x) === 'bikaner') ||
    districts[0] ||
    '';


  S.district = selected;


  UI.fillSelect(
    el,
    districts.map(v => ({
      v,
      t: v
    })),
    'Select District'
  );


  el.value = selected;

}


/* =========================================================
   PS
========================================================= */

function psFromMapping(r) {

  return safe(
    r.PS_EN ??
    r.PS ??
    r['PS EN'] ??
    r['Panchayat Samiti']
  );

}


function getPSList() {

  const dist = norm(S.district);

  const values = [];

  S.mappings.forEach(r => {

    if (
      !dist ||
      norm(districtFromMapping(r)) === dist
    ) {

      const ps = psFromMapping(r);

      if (ps) {
        values.push(ps);
      }

    }

  });


  S.dispatches.forEach(r => {

    const d = districtFromDispatch(r);

    if (
      !dist ||
      norm(d) === dist
    ) {

      const ps = safe(
        r.PS_EN ??
        r.PS ??
        r['PS']
      );

      if (ps) {
        values.push(ps);
      }

    }

  });


  return unique(values).sort();

}


function populatePS() {

  const el = $('ps');

  if (!el) return;

  const list = getPSList();

  let selected =
    S.ps ||
    list.find(x => norm(x) === 'khajuwala') ||
    list.find(x => norm(x) === 'bikaner') ||
    list[0] ||
    '';


  S.ps = selected;


  UI.fillSelect(
    el,
    list.map(v => ({
      v,
      t: v
    })),
    'Select Panchayat Samiti'
  );


  el.value = selected;

}


/* =========================================================
   GP
========================================================= */

function gpFromMapping(r) {

  return safe(
    r.GP_EN ??
    r.GP ??
    r['GP EN'] ??
    r['Gram Panchayat']
  );

}


function getGPList() {

  const dist = norm(S.district);
  const ps = norm(S.ps);

  const values = [];

  S.mappings.forEach(r => {

    const sameDist =
      !dist ||
      norm(districtFromMapping(r)) === dist;

    const samePS =
      !ps ||
      norm(psFromMapping(r)) === ps;

    if (sameDist && samePS) {

      const gp = gpFromMapping(r);

      if (gp) {
        values.push(gp);
      }

    }

  });


  return unique(values).sort();

}


function populateGP() {

  const el = $('gp');

  if (!el) return;

  const list = getGPList();


  UI.fillSelect(
    el,
    list.map(v => ({
      v,
      t: v
    })),
    'All Gram Panchayats'
  );


  /*
    Blank means ALL
  */

  S.gp = '';

  el.value = '';

}


/* =========================================================
   GENERATOR REFRESH
========================================================= */

function refreshGenerator() {

  populatePS();
  populateGP();
  updateInfo();

}


/* =========================================================
   INFO / WARNINGS
========================================================= */

function updateInfo() {

  const info = $('info');
  const warn = $('warn');

  const rows = getSelectedRows();

  if (info) {

    info.textContent =
      `${rows.length} record(s) selected`;

  }


  if (warn) {

    warn.textContent = '';

    if (!S.year) {
      warn.textContent = 'Please select year.';
      return;
    }

    if (!S.district) {
      warn.textContent = 'Please select district.';
      return;
    }

    if (!S.ps) {
      warn.textContent = 'Please select Panchayat Samiti.';
      return;
    }

    if (!rows.length) {
      warn.textContent =
        'No dispatch records found for selected filters.';
    }

  }

}


/* =========================================================
   SELECT DISPATCH ROWS
========================================================= */

function getSelectedRows() {

  const year = norm(S.year);
  const dist = norm(S.district);
  const ps = norm(S.ps);
  const gp = norm(S.gp);


  let rows = S.dispatches.filter(r => {

    const ry = norm(
      r.Year ??
      r.year ??
      ''
    );

    const rd = norm(
      r.District ??
      r.DIST_EN ??
      ''
    );

    const rp = norm(
      r.PS ??
      r.PS_EN ??
      ''
    );


    /*
      Year
    */

    if (
      year &&
      ry &&
      ry !== year
    ) {
      return false;
    }


    /*
      District
    */

    if (
      dist &&
      rd &&
      rd !== dist
    ) {
      return false;
    }


    /*
      PS
    */

    if (
      ps &&
      rp &&
      rp !== ps
    ) {
      return false;
    }


    /*
      GP filter
    */

    if (gp) {

      const rg = norm(
        r.GP ??
        r.GP_EN ??
        r['Gram Panchayat'] ??
        ''
      );

      if (rg !== gp) {
        return false;
      }

    }


    return true;

  });


  /*
    If dispatch file doesn't contain district/PS,
    try matching through Unit ID / mapping.
  */

  if (!rows.length && S.dispatches.length) {

    rows = S.dispatches.filter(r => {

      const unit =
        safe(
          r['Unit ID'] ??
          r.UnitID ??
          r.unitId
        );

      if (!unit) return false;

      const m = S.mappings.find(x => {

        const xUnit =
          safe(
            x['Unit ID'] ??
            x.UnitID ??
            x.unitId
          );

        return xUnit && xUnit === unit;

      });


      if (!m) return false;


      const md = norm(districtFromMapping(m));
      const mp = norm(psFromMapping(m));
      const mg = norm(gpFromMapping(m));


      if (dist && md !== dist) return false;
      if (ps && mp !== ps) return false;
      if (gp && mg !== gp) return false;


      return true;

    });

  }


  S.selectedRows = rows;

  return rows;

}


/* =========================================================
   MAP DISPATCH + MAPPING DATA
========================================================= */

function findMappingForRow(row) {

  const unitId = safe(
    row['Unit ID'] ??
    row.UnitID ??
    row.unitId
  );


  /*
    First try Unit ID.
  */

  if (unitId) {

    const exact = S.mappings.find(m => {

      const mid = safe(
        m['Unit ID'] ??
        m.UnitID ??
        m.unitId
      );

      return mid && mid === unitId;

    });

    if (exact) return exact;

  }


  /*
    Then GP + PS + District.
  */

  const rg = norm(
    row.GP ??
    row.GP_EN ??
    row['Gram Panchayat'] ??
    ''
  );

  const rp = norm(
    row.PS ??
    row.PS_EN ??
    ''
  );

  const rd = norm(
    row.District ??
    row.DIST_EN ??
    ''
  );


  return S.mappings.find(m => {

    const mg = norm(gpFromMapping(m));
    const mp = norm(psFromMapping(m));
    const md = norm(districtFromMapping(m));


    return (
      (!rg || mg === rg) &&
      (!rp || mp === rp) &&
      (!rd || md === rd)
    );

  }) || null;

}


/* =========================================================
   PREPARE WORD DATA
========================================================= */

function prepareWordData(row) {

  const mapping = findMappingForRow(row);


  /*
    Keep original fields also.
    This makes the generator compatible with
    different Word template placeholders.
  */

  const data = {
    ...row
  };


  /*
    English fields
  */

  data.YEAR =
    S.year ||
    row.Year ||
    '';

  data.DISTRICT =
    districtFromMapping(mapping || {}) ||
    row.District ||
    S.district ||
    '';

  data.PS =
    psFromMapping(mapping || {}) ||
    row.PS ||
    S.ps ||
    '';

  data.GP =
    gpFromMapping(mapping || {}) ||
    row.GP ||
    row.GP_EN ||
    S.gp ||
    '';


  /*
    Hindi mapping fields.
    Different possible column names are supported.
  */

  data.DISTRICT_HI =
    mapping?.DIST_HI ??
    mapping?.DISTRICT_HI ??
    mapping?.District_HI ??
    mapping?.['DIST_HI'] ??
    '';

  data.PS_HI =
    mapping?.PS_HI ??
    mapping?.PSHI ??
    mapping?.['PS_HI'] ??
    '';

  data.GP_HI =
    mapping?.GP_HI ??
    mapping?.GPHI ??
    mapping?.['GP_HI'] ??
    '';


  /*
    Dispatch number/date.
  */

  data.DISPATCH_NO =
    row['Dispatch No'] ??
    row['Dispatch Number'] ??
    row['Dispatch Sankhya'] ??
    row['Dispatch Sankhya'] ??
    row['Dispatch No.'] ??
    row.DispatchNo ??
    '';


  data.DATE =
    row.Date ??
    row.date ??
    row['Dispatch Date'] ??
    row['Date'] ??
    '';


  /*
    Common alternate placeholders.
  */

  data.DISPATCH_NUMBER = data.DISPATCH_NO;
  data.DISPATCH_DATE = data.DATE;

  data.YEAR_EN = data.YEAR;
  data.DIST_EN = data.DISTRICT;
  data.PS_EN = data.PS;
  data.GP_EN = data.GP;


  /*
    Hindi fallbacks.
  */

  if (!data.DISTRICT_HI) {
    data.DISTRICT_HI = data.DISTRICT;
  }

  if (!data.PS_HI) {
    data.PS_HI = data.PS;
  }

  if (!data.GP_HI) {
    data.GP_HI = data.GP;
  }


  return data;

}


/* =========================================================
   REPORT TYPE
========================================================= */

function setReportType(type) {

  if (!REPORTS[type]) {
    return;
  }

  S.rtype = type;


  document
    .querySelectorAll('#rtypes [data-t]')
    .forEach(btn => {

      btn.classList.toggle(
        'on',
        btn.dataset.t === type
      );

      btn.classList.toggle(
        'active',
        btn.dataset.t === type
      );

    });


  updateInfo();

}


/* =========================================================
   GENERATE DOCUMENT
========================================================= */

async function generateReport() {

  try {

    const rows = getSelectedRows();

    if (!S.year) {
      throw new Error('Please select year.');
    }

    if (!S.district) {
      throw new Error('Please select district.');
    }

    if (!S.ps) {
      throw new Error('Please select Panchayat Samiti.');
    }

    if (!rows.length) {
      throw new Error(
        'No records found for selected filters.'
      );
    }


    const report = REPORTS[S.rtype];

    if (!report) {
      throw new Error('Invalid report type.');
    }


    /*
      Currently repository contains only:
      covering-letter-template.docx

      Therefore other report types are intentionally
      blocked until their templates are uploaded.
    */

    if (S.rtype !== 'ybc') {

      throw new Error(
        `Template for "${S.rtype}" is not uploaded yet. ` +
        `For now select "Year Book Closing".`
      );

    }


    UI.showToast(
      `Generating ${rows.length} document(s)...`
    );


    const items = rows.map(prepareWordData);


    /*
      generateDocx handles the template.
    */

    const blob = await generateDocx(
      items,
      report.tpl
    );


    if (!blob) {
      throw new Error(
        'Word file generation failed.'
      );
    }


    const filename =
      `${report.prefix}_${safe(S.year)}_` +
      `${safe(S.district)}_` +
      `${safe(S.ps)}.docx`;


    const result = $('result');
    const rname = $('rname');
    const dlBtn = $('dlBtn');
    const waBtn = $('waBtn');


    if (rname) {
      rname.textContent = filename;
    }


    if (result) {
      result.style.display = '';
      result.hidden = false;
    }


    /*
      Store blob for download.
    */

    S.generatedBlob = blob;
    S.generatedFilename = filename;


    if (dlBtn) {

      dlBtn.onclick = () => {

        const url =
          URL.createObjectURL(S.generatedBlob);

        const a =
          document.createElement('a');

        a.href = url;
        a.download = S.generatedFilename;

        document.body.appendChild(a);

        a.click();

        a.remove();

        setTimeout(
          () => URL.revokeObjectURL(url),
          3000
        );

      };

    }


    /*
      WhatsApp share.
    */

    if (waBtn) {

      waBtn.onclick = () => {

        const text =
          `Year Book Closing report ready\n` +
          `Year: ${S.year}\n` +
          `District: ${S.district}\n` +
          `PS: ${S.ps}\n` +
          `GP: ${S.gp || 'All Gram Panchayats'}\n` +
          `File: ${filename}`;

        window.open(
          'https://wa.me/?text=' +
          encodeURIComponent(text),
          '_blank'
        );

      };

    }


    UI.showToast(
      'Word file generated successfully.',
      'success'
    );


  } catch (e) {

    console.error(
      'GENERATION ERROR:',
      e
    );

    showError(e);

  }

}


/* =========================================================
   DISPATCH FILE PREVIEW
========================================================= */

async function readExcelFile(file) {

  if (!file) {
    throw new Error('Please select an Excel file.');
  }


  if (!window.XLSX) {
    throw new Error(
      'Excel library not loaded. Please refresh page.'
    );
  }


  const buffer =
    await file.arrayBuffer();


  const wb =
    XLSX.read(buffer, {
      type: 'array'
    });


  const first =
    wb.Sheets[wb.SheetNames[0]];


  const rows =
    XLSX.utils.sheet_to_json(
      first,
      {
        defval: ''
      }
    );


  return rows;

}


/* =========================================================
   DISPATCH PREVIEW
========================================================= */

async function previewDispatchFile() {

  try {

    const file = $('file')?.files?.[0];

    if (!file) {
      throw new Error(
        'Please select Dispatch Excel file.'
      );
    }


    const rows =
      await readExcelFile(file);


    S.dispatchPreview = rows;


    const info = $('fileInfo');

    if (info) {

      info.textContent =
        `${rows.length} rows loaded from ${file.name}`;

    }


    const preview = $('preview');

    if (preview) {

      preview.innerHTML =
        makeTable(rows.slice(0, 20));

    }


    const prevBtn = $('prevBtn');

    if (prevBtn) {
      prevBtn.disabled = false;
    }


    UI.showToast(
      `${rows.length} rows ready for upload.`,
      'success'
    );


  } catch (e) {

    showError(e);

  }

}


/* =========================================================
   DISPATCH UPLOAD
========================================================= */

async function uploadDispatchFile() {

  try {

    if (!S.dispatchPreview.length) {

      await previewDispatchFile();

    }


    if (!S.dispatchPreview.length) {
      return;
    }


    const year =
      safe(
        $('uploadYear')?.value ||
        $('year')?.value ||
        S.year
      );


    if (!year) {
      throw new Error(
        'Please select/upload year.'
      );
    }


    const existingIds =
      new Set(
        S.dispatches.map(
          r => r._id
        )
      );


    const res =
      await uploadDispatchRows(
        S.dispatchPreview,
        year,
        existingIds,
        (done, total) => {

          const el = $('upRes');

          if (el) {
            el.textContent =
              `Uploading ${done}/${total}...`;
          }

        }
      );


    const el = $('upRes');

    if (el) {

      el.textContent =
        `Inserted: ${res.inserted}, ` +
        `Updated: ${res.updated}, ` +
        `Skipped: ${res.skipped}, ` +
        `Errors: ${res.errors}`;

    }


    UI.showToast(
      'Dispatch upload completed.',
      'success'
    );


    await loadData();


  } catch (e) {

    showError(e);

  }

}


/* =========================================================
   MAPPING PREVIEW
========================================================= */

async function previewMappingFile() {

  try {

    const file = $('mFile')?.files?.[0];

    if (!file) {
      throw new Error(
        'Please select Mapping Excel file.'
      );
    }


    const rows =
      await readExcelFile(file);


    S.mappingPreview = rows;


    const info = $('mInfo');

    if (info) {

      info.textContent =
        `${rows.length} mapping rows loaded from ${file.name}`;

    }


    UI.showToast(
      `${rows.length} mapping rows ready.`,
      'success'
    );


  } catch (e) {

    showError(e);

  }

}


/* =========================================================
   MAPPING UPLOAD
========================================================= */

async function uploadMappingFile() {

  try {

    if (!S.mappingPreview.length) {
      await previewMappingFile();
    }


    if (!S.mappingPreview.length) {
      return;
    }


    const existingIds =
      new Set(
        S.mappings.map(
          r => r._id
        )
      );


    const res =
      await uploadMappingRows(
        S.mappingPreview,
        existingIds,
        (done, total) => {

          const el = $('mRes');

          if (el) {
            el.textContent =
              `Uploading ${done}/${total}...`;
          }

        }
      );


    const el = $('mRes');

    if (el) {

      el.textContent =
        `Inserted: ${res.inserted}, ` +
        `Updated: ${res.updated}, ` +
        `Skipped: ${res.skipped}, ` +
        `Errors: ${res.errors}`;

    }


    UI.showToast(
      'Mapping upload completed.',
      'success'
    );


    await loadData();


  } catch (e) {

    showError(e);

  }

}


/* =========================================================
   SIMPLE HTML TABLE
========================================================= */

function makeTable(rows) {

  if (!rows?.length) {
    return '<p>No data.</p>';
  }


  const headers =
    unique(
      rows.flatMap(
        r => Object.keys(r)
      )
    );


  const head =
    headers
      .map(
        h => `<th>${UI.esc(h)}</th>`
      )
      .join('');


  const body =
    rows
      .map(r => {

        return `
          <tr>
            ${
              headers
                .map(
                  h =>
                    `<td>${UI.esc(r[h])}</td>`
                )
                .join('')
            }
          </tr>
        `;

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


/* =========================================================
   DATA VIEW
========================================================= */

function renderDataView() {

  const out = $('dataOut');

  if (!out) return;


  const year =
    norm($('fYear')?.value);

  const dist =
    norm($('fDist')?.value);

  const ps =
    norm($('fPs')?.value);

  const status =
    norm($('fStatus')?.value);

  const search =
    norm($('fSearch')?.value);


  const rows =
    S.dispatches.filter(r => {

      const text =
        norm(
          Object
            .values(r)
            .join(' ')
        );


      if (
        year &&
        norm(r.Year) !== year
      ) {
        return false;
      }


      if (
        dist &&
        norm(
          r.District ??
          r.DIST_EN
        ) !== dist
      ) {
        return false;
      }


      if (
        ps &&
        norm(
          r.PS ??
          r.PS_EN
        ) !== ps
      ) {
        return false;
      }


      if (
        status &&
        norm(
          r.Status ??
          r.status
        ) !== status
      ) {
        return false;
      }


      if (
        search &&
        !text.includes(search)
      ) {
        return false;
      }


      return true;

    });


  out.innerHTML =
    `<p>${rows.length} record(s)</p>` +
    makeTable(rows.slice(0, 200));

}


/* =========================================================
   DATA FILTER DROPDOWNS
========================================================= */

function setupDataFilters() {

  const fy = $('fYear');
  const fd = $('fDist');
  const fp = $('fPs');


  if (fy) {

    UI.fillSelect(
      fy,
      getYears().map(v => ({
        v,
        t: v
      })),
      'All Years'
    );

  }


  if (fd) {

    UI.fillSelect(
      fd,
      getDistricts().map(v => ({
        v,
        t: v
      })),
      'All Districts'
    );

  }


  if (fp) {

    UI.fillSelect(
      fp,
      getPSList().map(v => ({
        v,
        t: v
      })),
      'All PS'
    );

  }


  [
    fy,
    fd,
    fp,
    $('fStatus'),
    $('fSearch')
  ]
    .filter(Boolean)
    .forEach(el => {

      el.addEventListener(
        'input',
        renderDataView
      );

      el.addEventListener(
        'change',
        renderDataView
      );

    });

}


/* =========================================================
   EVENTS
========================================================= */

function setupEvents() {

  /*
    Google Login
  */

  const gBtn = $('gBtn');

  if (gBtn) {

    gBtn.onclick = async () => {

      try {

        gBtn.disabled = true;

        gBtn.textContent =
          'Signing in...';

        await googleLogin();

      } catch (e) {

        console.error(e);

        const lerr = $('lerr');

        if (lerr) {

          lerr.textContent =
            e?.message ||
            'Login failed.';

        }

      } finally {

        gBtn.disabled = false;

        gBtn.textContent =
          'Continue with Google';

      }

    };

  }


  /*
    Logout
  */

  document.addEventListener(
    'click',
    e => {

      const el =
        e.target.closest(
          '[data-action="logout"],#logoutBtn'
        );

      if (!el) return;

      logout().catch(showError);

    }
  );


  /*
    Report type buttons
  */

  document
    .querySelectorAll('#rtypes [data-t]')
    .forEach(btn => {

      btn.addEventListener(
        'click',
        () => {

          setReportType(
            btn.dataset.t
          );

        }
      );

    });


  /*
    Year
  */

  const year = $('year');

  if (year) {

    year.addEventListener(
      'change',
      () => {

        S.year = year.value;

        populateDistrict();
        populatePS();
        populateGP();
        updateInfo();

      }
    );

  }


  /*
    District
  */

  const dist = $('dist');

  if (dist) {

    dist.addEventListener(
      'change',
      () => {

        S.district = dist.value;

        populatePS();
        populateGP();
        updateInfo();

      }
    );

  }


  /*
    PS
  */

  const ps = $('ps');

  if (ps) {

    ps.addEventListener(
      'change',
      () => {

        S.ps = ps.value;

        populateGP();
        updateInfo();

      }
    );

  }


  /*
    GP
  */

  const gp = $('gp');

  if (gp) {

    gp.addEventListener(
      'change',
      () => {

        S.gp = gp.value;

        updateInfo();

      }
    );

  }


  /*
    Generate
  */

  const genBtn = $('genBtn');

  if (genBtn) {

    genBtn.addEventListener(
      'click',
      generateReport
    );

  }


  /*
    Dispatch file
  */

  const file = $('file');

  if (file) {

    file.addEventListener(
      'change',
      previewDispatchFile
    );

  }


  const prevBtn = $('prevBtn');

  if (prevBtn) {

    prevBtn.addEventListener(
      'click',
      previewDispatchFile
    );

  }


  const upBtn = $('upBtn');

  if (upBtn) {

    upBtn.addEventListener(
      'click',
      uploadDispatchFile
    );

  }


  /*
    Mapping
  */

  const mFile = $('mFile');

  if (mFile) {

    mFile.addEventListener(
      'change',
      previewMappingFile
    );

  }


  const mBtn = $('mBtn');

  if (mBtn) {

    mBtn.addEventListener(
      'click',
      uploadMappingFile
    );

  }


  /*
    Reload
  */

  const reloadBtn = $('reloadBtn');

  if (reloadBtn) {

    reloadBtn.addEventListener(
      'click',
      async () => {

        await loadData();

        setupDataFilters();

      }
    );

  }

}


/* =========================================================
   NAVIGATION
========================================================= */

function setupNavigation() {

  if (!UI.setupNav) {
    return;
  }


  UI.setupNav(
    view => {

      console.log(
        'Opened view:',
        view
      );


      if (view === 'generate') {

        refreshGenerator();

      }


      if (view === 'data') {

        setupDataFilters();
        renderDataView();

      }

    }
  );

}


/* =========================================================
   AUTH STATE
========================================================= */

function setupAuth() {

  onUser(
    async user => {

      console.log(
        'AUTH STATE:',
        user
      );


      if (user) {

        S.user = user;

        updateUserUI(user);

        /*
          VERY IMPORTANT:
          This fixes the blank page after login.
        */

        showApp();


        /*
          Load Firestore after shell is visible.
        */

        await loadData();


        setupDataFilters();


      } else {

        S.user = null;

        showLogin();

      }

    }
  );

}


/* =========================================================
   INITIALIZATION
========================================================= */

function init() {

  console.log(
    'Jamidara app initializing...'
  );


  /*
    Initially show login state.
  */

  showLogin();


  /*
    Event listeners.
  */

  setupEvents();


  /*
    Drawer/navigation.
  */

  setupNavigation();


  /*
    Default report.
  */

  setReportType('ybc');


  /*
    Firebase Auth listener.
  */

  setupAuth();

}


/* =========================================================
   START
========================================================= */

if (
  document.readyState === 'loading'
) {

  document.addEventListener(
    'DOMContentLoaded',
    init
  );

} else {

  init();

}
