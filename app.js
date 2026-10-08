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
import { buildAmsReport, normPS } from './js/ams-report.js';

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

  selectedRows: [],

  dispatchPreview: [],
  mappingPreview: [],

  generatedBlob: null,
  generatedFilename: ''
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
    prefix: 'Cover_Letter'
  },

  intim: {
    tpl: 'intimation-report.docx',
    prefix: 'Intimation_Report'
  }
};


/* =========================================================
   BASIC HELPERS
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


/*
  Year को compare करने के लिए normalize करते हैं.

  2025-26
  2025 – 26
  2025-2026

  इन formats को comparison में compatible बनाया जाएगा.
*/

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


function showError(error) {

  console.error(error);

  const message =
    UI.friendly
      ? UI.friendly(error)
      : (error?.message || String(error));

  UI.showToast(message, 'error');

}


/* =========================================================
   FIELD HELPERS
========================================================= */

/*
  Mapping Excel / Firestore में column names थोड़ा
  अलग हों तो भी application काम करे.
*/

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


function getDispatchYear(row) {

  return safe(
    row.Year ??
    row.year ??
    row['Year'] ??
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


/* =========================================================
   LOGIN / APP VISIBILITY
========================================================= */

function showApp() {

  const login = $('login');

  if (login) {
    login.style.display = 'none';
  }


  /*
    CSS में body.authed के आधार पर application
    दिखाई जाती है.
  */

  document.body.classList.add('authed');


  /*
    सभी views बंद करो.
  */

  document
    .querySelectorAll('.view')
    .forEach(view => {
      view.classList.remove('active');
    });


  /*
    Dashboard खोलो.
  */

  const dashboard = $('v-dashboard');

  if (dashboard) {
    dashboard.classList.add('active');
  }


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

  document.body.classList.remove('authed');


  const login = $('login');

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
   USER UI
========================================================= */

function updateUserUI(user) {

  const name =
    user?.displayName ||
    user?.email ||
    'User';

  const email =
    user?.email ||
    '';


  document
    .querySelectorAll(
      '#userName,[data-user-name],#profileName'
    )
    .forEach(el => {
      el.textContent = name;
    });


  document
    .querySelectorAll(
      '#userEmail,[data-user-email]'
    )
    .forEach(el => {
      el.textContent = email;
    });

}


/* =========================================================
   FIREBASE DATA LOAD
========================================================= */

async function loadData() {

  try {

    console.log('Loading Firestore data...');


    const [
      mappings,
      dispatches
    ] = await Promise.all([
      loadMappings(),
      loadDispatchData()
    ]);


    S.mappings =
      Array.isArray(mappings)
        ? mappings
        : [];


    S.dispatches =
      Array.isArray(dispatches)
        ? dispatches
        : [];


    console.log(
      'Mappings:',
      S.mappings.length
    );

    console.log(
      'Dispatches:',
      S.dispatches.length
    );


    /*
      IMPORTANT:

      पहले year Firebase dispatch data से.
      फिर district Firebase mappings से.
      फिर PS/GP cascading तरीके से.
    */

    populateYear();


    /*
      Year set होने के बाद district.
    */

    populateDistrict();


    /*
      District set होने के बाद PS.
    */

    populatePS();


    /*
      PS set होने के बाद GP.
    */

    populateGP();


    updateInfo();


    console.log(
      'Firebase data initialized successfully.'
    );


  } catch (error) {

    console.error(
      'Firestore loading error:',
      error
    );

    showError(error);

  }

}


/* =========================================================
   YEAR FROM FIREBASE ONLY
========================================================= */

/*
  Year केवल dispatches collection से आएगा.

  कोई default year नहीं.
  कोई hard-coded year नहीं.
*/

function getAvailableYears() {

  const years =
    S.dispatches
      .map(row =>
        normalizeYear(
          getDispatchYear(row)
        )
      )
      .filter(Boolean);


  return uniqueSorted(years);

}


function populateYear() {

  const select = $('year');

  if (!select) return;


  const years =
    getAvailableYears();


  /*
    Current selection अगर अभी भी Firebase में है
    तो वही रखें.
  */

  const previous =
    normalizeYear(S.year);


  UI.fillSelect(
    select,

    years.map(year => ({
      v: year,
      t: year
    })),

    'Select Year'
  );


  if (
    previous &&
    years.includes(previous)
  ) {

    S.year = previous;
    select.value = previous;

  } else {

    /*
      कोई hard-coded/default selection नहीं.
    */

    S.year = '';
    select.value = '';

  }


  console.log(
    'Available Firebase Years:',
    years
  );

}


/* =========================================================
   DISTRICT FROM FIREBASE MAPPING
========================================================= */

/*
  District mapping collection से आएगा.
  कोई hard-coded district नहीं.
*/

function getAvailableDistricts() {

  const values = [];


  S.mappings.forEach(row => {

    const district =
      getMappingDistrict(row);

    if (district) {
      values.push(district);
    }

  });


  return uniqueSorted(values);

}


function populateDistrict() {

  const select = $('dist');

  if (!select) return;


  const districts =
    getAvailableDistricts();


  const previous =
    safe(S.district);


  UI.fillSelect(
    select,

    districts.map(district => ({
      v: district,
      t: district
    })),

    'Select District'
  );


  if (
    previous &&
    districts.some(
      x => norm(x) === norm(previous)
    )
  ) {

    const actual =
      districts.find(
        x => norm(x) === norm(previous)
      );

    S.district = actual;
    select.value = actual;

  } else {

    S.district = '';
    select.value = '';

  }


  console.log(
    'Available Firebase Districts:',
    districts
  );

}


/* =========================================================
   PS FROM FIREBASE MAPPING
========================================================= */

function getAvailablePS() {

  const district =
    norm(S.district);


  if (!district) {
    return [];
  }


  const values = [];


  S.mappings.forEach(row => {

    const rowDistrict =
      norm(
        getMappingDistrict(row)
      );


    if (
      rowDistrict === district
    ) {

      const ps =
        getMappingPS(row);

      if (ps) {
        values.push(ps);
      }

    }

  });


  return uniqueSorted(values);

}


function populatePS() {

  const select = $('ps');

  if (!select) return;


  const psList =
    getAvailablePS();


  const previous =
    safe(S.ps);


  UI.fillSelect(
    select,

    psList.map(ps => ({
      v: ps,
      t: ps
    })),

    'Select Panchayat Samiti'
  );


  if (
    previous &&
    psList.some(
      x => norm(x) === norm(previous)
    )
  ) {

    const actual =
      psList.find(
        x => norm(x) === norm(previous)
      );

    S.ps = actual;
    select.value = actual;

  } else {

    S.ps = '';
    select.value = '';

  }


  console.log(
    'Available Firebase PS:',
    psList
  );

}


/* =========================================================
   GP FROM FIREBASE MAPPING
========================================================= */

function getAvailableGP() {

  const district =
    norm(S.district);

  const ps =
    norm(S.ps);


  if (
    !district ||
    !ps
  ) {
    return [];
  }


  const values = [];


  S.mappings.forEach(row => {

    const rowDistrict =
      norm(
        getMappingDistrict(row)
      );

    const rowPS =
      norm(
        getMappingPS(row)
      );


    if (
      rowDistrict === district &&
      rowPS === ps
    ) {

      const gp =
        getMappingGP(row);

      if (gp) {
        values.push(gp);
      }

    }

  });


  return uniqueSorted(values);

}


function populateGP() {

  const select = $('gp');

  if (!select) return;


  const gps =
    getAvailableGP();


  const previous =
    safe(S.gp);


  UI.fillSelect(
    select,

    gps.map(gp => ({
      v: gp,
      t: gp
    })),

    /*
      Blank value = all GPs
    */

    'All Gram Panchayats'
  );


  if (
    previous &&
    gps.some(
      x => norm(x) === norm(previous)
    )
  ) {

    const actual =
      gps.find(
        x => norm(x) === norm(previous)
      );

    S.gp = actual;
    select.value = actual;

  } else {

    S.gp = '';
    select.value = '';

  }


  console.log(
    'Available Firebase GPs:',
    gps
  );

}


/* =========================================================
   MAPPING LOOKUP
========================================================= */

function findMappingForRow(row) {

  /*
    सबसे पहले Unit ID.
  */

  const unitId =
    getDispatchUnitId(row);


  if (unitId) {

    const exact =
      S.mappings.find(mapping => {

        const mappingUnitId =
          safe(
            mapping['Unit ID'] ??
            mapping.UnitID ??
            mapping.unitId ??
            mapping['UnitID'] ??
            ''
          );


        return (
          mappingUnitId &&
          mappingUnitId === unitId
        );

      });


    if (exact) {
      return exact;
    }

  }


  /*
    अगर Unit ID mapping में नहीं मिला,
    तो District + PS + GP से कोशिश करें.
  */

  const rowDistrict =
    norm(
      getDispatchDistrict(row)
    );

  const rowPS =
    norm(
      getDispatchPS(row)
    );

  const rowGP =
    norm(
      getDispatchGP(row)
    );


  if (
    !rowDistrict &&
    !rowPS &&
    !rowGP
  ) {
    return null;
  }


  return (
    S.mappings.find(mapping => {

      const md =
        norm(
          getMappingDistrict(mapping)
        );

      const mp =
        norm(
          getMappingPS(mapping)
        );

      const mg =
        norm(
          getMappingGP(mapping)
        );


      return (
        (!rowDistrict || !md || md === rowDistrict) &&
        (!rowPS || !mp || mp === rowPS) &&
        (!rowGP || !mg || mg === rowGP)
      );

    }) ||
    null
  );

}


/* =========================================================
   FILTER DISPATCH RECORDS
========================================================= */

/*
  IMPORTANT:

  Year = dispatch Firebase

  District/PS/GP =
  mapping Firebase

  इसलिए dispatch record में District/PS/GP
  होना जरूरी नहीं.
*/

function getSelectedRows() {

  const selectedYear =
    normalizeYear(S.year);

  const selectedDistrict =
    norm(S.district);

  const selectedPS =
    norm(S.ps);

  const selectedGP =
    norm(S.gp);


  console.log(
    'Generating with filters:',
    {
      year: selectedYear,
      district: selectedDistrict,
      ps: selectedPS,
      gp: selectedGP
    }
  );


  if (!selectedYear) {

    S.selectedRows = [];

    return [];

  }


  if (!selectedDistrict) {

    S.selectedRows = [];

    return [];

  }


  if (!selectedPS) {

    S.selectedRows = [];

    return [];

  }


  /*
    पहले उस Year के dispatch records.
  */

  const yearRows =
    S.dispatches.filter(row => {

      const rowYear =
        normalizeYear(
          getDispatchYear(row)
        );


      return (
        rowYear === selectedYear
      );

    });


  console.log(
    'Dispatch rows for selected year:',
    yearRows.length
  );


  /*
    अब mapping के आधार पर District/PS/GP filter.
  */

  const result = [];


  yearRows.forEach(row => {

    const mapping =
      findMappingForRow(row);


    /*
      Mapping जरूरी है क्योंकि
      District/PS/GP वहीं से आ रहे हैं.
    */

    if (!mapping) {
      return;
    }


    const mappingDistrict =
      norm(
        getMappingDistrict(mapping)
      );

    const mappingPS =
      norm(
        getMappingPS(mapping)
      );

    const mappingGP =
      norm(
        getMappingGP(mapping)
      );


    /*
      District
    */

    if (
      mappingDistrict !== selectedDistrict
    ) {
      return;
    }


    /*
      Panchayat Samiti
    */

    if (
      mappingPS !== selectedPS
    ) {
      return;
    }


    /*
      GP

      Blank GP = ALL
    */

    if (
      selectedGP &&
      mappingGP !== selectedGP
    ) {
      return;
    }


    /*
      Mapping information temporarily attach
      करते हैं ताकि document generation में
      दुबारा lookup न करना पड़े.
    */

    result.push({
      ...row,
      __mapping: mapping
    });

  });


  S.selectedRows = result;


  console.log(
    'FINAL SELECTED ROWS:',
    result.length,
    result
  );


  return result;

}


/* =========================================================
   LIVE PREVIEW (bina Generate kiye details dekhne ke liye)
========================================================= */

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


/* =========================================================
   GENERATOR INFO
========================================================= */

function updateInfo() {

  const info = $('info');
  const warn = $('warn');


  if (!info && !warn) {
    return;
  }


  if (S.rtype === 'ams') {
    if (info) info.innerHTML = '';
    if (warn) warn.innerHTML = '';
    renderRowsPreview([]);
    renderAmsPs();
    return;
  }


  const rows =
    getSelectedRows();

  renderRowsPreview(rows);


  if (info) {

    if (
      !S.year ||
      !S.district ||
      !S.ps
    ) {

      info.textContent =
        'Select Year, District and Panchayat Samiti.';

    } else {

      info.textContent =
        `${rows.length} dispatch record(s) available`;

    }

  }


  if (warn) {

    warn.textContent = '';


    if (!S.year) {

      warn.textContent =
        'Please select Year from Firebase data.';

      return;

    }


    if (!S.district) {

      warn.textContent =
        'Please select District from Firebase mapping data.';

      return;

    }


    if (!S.ps) {

      warn.textContent =
        'Please select Panchayat Samiti from Firebase mapping data.';

      return;

    }


    if (!rows.length) {

      warn.textContent =
        'No dispatch record found for selected filters.';

    }

  }

}


/* =========================================================
   PREPARE WORD DATA
========================================================= */

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


/* =========================================================
   REPORT TYPE
========================================================= */

function setReportType(type) {

  if (!REPORTS[type]) {
    return;
  }


  S.rtype = type;

  const isAms = type === 'ams';
  const ls = $('letterSel'), ab = $('amsBox'), gb = $('genBtn'), rs = $('result');
  if (ls) ls.hidden = isAms;
  if (ab) ab.hidden = !isAms;
  if (gb) gb.textContent = isAms ? 'Generate AMS Report' : 'Generate Cover Letter';
  if (rs) { rs.hidden = true; rs.style.display = 'none'; }


  document
    .querySelectorAll('#rtypes [data-t]')
    .forEach(button => {

      const active =
        button.dataset.t === type;


      button.classList.toggle(
        'on',
        active
      );

      button.classList.toggle(
        'active',
        active
      );

    });

  updateInfo();

}


/* =========================================================
   AMS REPORT + WATCH LIST
========================================================= */

const WL_KEY = 'jamidara_ams_watchlist_v1';

function loadWatch() {
  try {
    const a = JSON.parse(localStorage.getItem(WL_KEY) || '[]');
    return Array.isArray(a) ? a : [];
  } catch (e) {
    return [];
  }
}

function saveWatch(list) {
  try {
    localStorage.setItem(WL_KEY, JSON.stringify(list));
    return true;
  } catch (e) {
    return false;
  }
}

function titleCase(t) {
  return String(t || '').toLowerCase().replace(/\b\w/g, c => c.toUpperCase());
}

function amsState() {
  if (!S.amsTicked) S.amsTicked = new Set(loadWatch());
  if (!S.amsInclude) S.amsInclude = 1;
  if (S.amsSearch === undefined) S.amsSearch = '';
  return S;
}

function amsYearRecords() {
  const y = normalizeYear(S.year);
  return S.dispatches.filter(r => !y || normalizeYear(getDispatchYear(r)) === y);
}

function amsPsOptions() {
  const map = new Map();
  amsYearRecords().forEach(r => {
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
    ? saved.map(n => `<span class="chip">${UI.esc(titleCase(n))} <b data-rm="${UI.esc(n)}" title="Hatao">×</b></span>`).join('')
    : '<span class="muted">Watch list khali hai. Upar PS tick karke Save karo.</span>';
}

function renderAmsPs() {
  amsState();
  const box = $('amsPs');
  if (!box) return;

  if (!S.year) {
    box.innerHTML = '<div style="padding:14px" class="muted">Pehle Year select karo.</div>';
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
    : '<div style="padding:14px" class="muted">Is Year me koi PS nahi mila.</div>';

  const c = $('amsCount');
  if (c) c.textContent = `${S.amsTicked.size} tick`;
  renderAmsSaved();
}

async function shareFileWhatsApp(blob, filename, text) {
  try {
    const file = new File([blob], filename, { type: blob.type });
    if (navigator.canShare && navigator.canShare({ files: [file] })) {
      await navigator.share({ files: [file], text });
      return;
    }
  } catch (e) {
    if (e && e.name === 'AbortError') return;
  }
  window.open(
    'https://wa.me/?text=' + encodeURIComponent(text + '\n(Word file download karke attach karein)'),
    '_blank'
  );
}

async function generateAms() {
  try {
    amsState();
    if (!S.year) throw new Error('Please select Year.');

    const saved = loadWatch();
    if (!saved.length) {
      throw new Error('Watch list khali hai. PS tick karke "Save Watch List" dabao.');
    }

    const recs = amsYearRecords();
    if (!recs.length) throw new Error('Is Year ka dispatch data nahi mila.');

    let res;
    try {
      res = buildAmsReport(recs, saved, S.amsInclude);
    } catch (e) {
      if (e && e.message === 'NO_PS') {
        throw new Error('Watch list ke PS is Year me nahi mile. Year check karo ya watch list badlo.');
      }
      throw e;
    }

    const stamp = new Date().toISOString().slice(0, 10);
    const filename = `AMS_Report_${safe(S.year)}_${stamp}.docx`;

    S.generatedBlob = res.blob;
    S.generatedFilename = filename;

    const result = $('result');
    const rname = $('rname');
    const dlBtn = $('dlBtn');
    const waBtn = $('waBtn');

    const have = new Set(amsPsOptions().map(o => o.name));
    const missing = saved.filter(n => !have.has(normPS(n)));

    if (rname) {
      rname.textContent =
        `${filename} · ${res.psCount} PS · ${res.totalGp} GP` +
        (missing.length ? ` · Is Year me nahi mili: ${missing.map(titleCase).join(', ')}` : '');
    }
    if (result) { result.style.display = ''; result.hidden = false; }

    if (dlBtn) {
      dlBtn.onclick = () => {
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

    if (waBtn) {
      waBtn.onclick = () => shareFileWhatsApp(
        S.generatedBlob,
        S.generatedFilename,
        `AMS Report\nYear: ${S.year}\nPS: ${res.psCount}\nGP: ${res.totalGp}`
      );
    }

    UI.showToast('AMS report ready.', 'success');

  } catch (error) {
    console.error('AMS ERROR:', error);
    showError(error);
  }
}

function initAms() {
  amsState();

  const search = $('amsSearch');
  if (search) {
    search.addEventListener('input', () => {
      S.amsSearch = search.value;
      renderAmsPs();
    });
  }

  const ps = $('amsPs');
  if (ps) {
    ps.addEventListener('change', e => {
      const t = e.target;
      if (!t || !t.dataset || t.dataset.n === undefined) return;
      if (t.checked) S.amsTicked.add(t.dataset.n);
      else S.amsTicked.delete(t.dataset.n);
      const c = $('amsCount');
      if (c) c.textContent = `${S.amsTicked.size} tick`;
    });
  }

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
    if (!list.length) {
      UI.showToast('Pehle kam se kam ek PS tick karo.', 'error');
      return;
    }
    if (saveWatch(list)) {
      renderAmsSaved();
      UI.showToast(`Watch list saved (${list.length} PS).`, 'success');
    } else {
      UI.showToast('Watch list save nahi ho payi (browser storage band hai).', 'error');
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
   GENERATE WORD
========================================================= */

async function generateReport() {

  try {

    updateInfo();


    const rows =
      getSelectedRows();


    if (!S.year) {

      throw new Error(
        'Please select Year.'
      );

    }


    if (!S.district) {

      throw new Error(
        'Please select District.'
      );

    }


    if (!S.ps) {

      throw new Error(
        'Please select Panchayat Samiti.'
      );

    }


    if (!rows.length) {

      throw new Error(
        'No dispatch record found for selected filters.'
      );

    }


    const report =
      REPORTS[S.rtype];


    if (!report) {

      throw new Error(
        'Invalid report type.'
      );

    }


    /*
      अभी repository में केवल
      covering-letter-template.docx है.
    */

    if (S.rtype !== 'ybc') {

      throw new Error(
        `Template for "${S.rtype}" is not uploaded yet. ` +
        `For now select "Cover Letter".`
      );

    }


    UI.showToast(
      `Generating ${rows.length} document(s)...`
    );


    const wordData =
      rows.map(
        prepareWordData
      );


    const blob =
      await generateDocx(
        wordData,
        report.tpl
      );


    if (!blob) {

      throw new Error(
        'Word file generation failed.'
      );

    }


    const filename =
      `${report.prefix}_` +
      `${safe(S.year)}_` +
      `${safe(S.district)}_` +
      `${safe(S.ps)}.docx`;


    S.generatedBlob =
      blob;

    S.generatedFilename =
      filename;


    const result =
      $('result');

    const rname =
      $('rname');

    const dlBtn =
      $('dlBtn');

    const waBtn =
      $('waBtn');


    if (rname) {
      rname.textContent =
        filename;
    }


    if (result) {

      result.style.display = '';
      result.hidden = false;

    }


    if (dlBtn) {

      dlBtn.onclick =
        () => {

          if (!S.generatedBlob) {
            return;
          }


          const url =
            URL.createObjectURL(
              S.generatedBlob
            );


          const a =
            document.createElement('a');


          a.href = url;
          a.download =
            S.generatedFilename;


          document.body.appendChild(a);

          a.click();

          a.remove();


          setTimeout(
            () =>
              URL.revokeObjectURL(url),
            3000
          );

        };

    }


    if (waBtn) {

      waBtn.onclick =
        () => {

          const text =
            `Cover Letter\n` +
            `Year: ${S.year}\n` +
            `District: ${S.district}\n` +
            `Panchayat Samiti: ${S.ps}\n` +
            `Gram Panchayat: ${
              S.gp ||
              'All Gram Panchayats'
            }\n` +
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


  } catch (error) {

    console.error(
      'GENERATION ERROR:',
      error
    );

    showError(error);

  }

}


/* =========================================================
   EXCEL READER
========================================================= */

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


/* =========================================================
   SIMPLE TABLE
========================================================= */

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


/* =========================================================
   DISPATCH PREVIEW
========================================================= */

async function previewDispatchFile() {

  try {

    const file =
      $('file')?.files?.[0];


    if (!file) {

      throw new Error(
        'Please select Dispatch Excel file.'
      );

    }


    const rows =
      await readExcelFile(file);


    S.dispatchPreview =
      rows;


    const info =
      $('fileInfo');


    if (info) {

      info.textContent =
        `${rows.length} rows loaded from ${file.name}`;

    }


    const preview =
      $('preview');


    if (preview) {

      preview.innerHTML =
        makeTable(
          rows.slice(0, 20)
        );

    }


    const prevBtn =
      $('prevBtn');


    if (prevBtn) {
      prevBtn.disabled = false;
    }


    UI.showToast(
      `${rows.length} rows ready for upload.`,
      'success'
    );


  } catch (error) {

    showError(error);

  }

}


/* =========================================================
   DISPATCH UPLOAD
========================================================= */

async function uploadDispatchFile() {

  try {

    if (
      !S.dispatchPreview.length
    ) {

      await previewDispatchFile();

    }


    if (
      !S.dispatchPreview.length
    ) {
      return;
    }


    /*
      Upload year भी Firebase-generated list से
      selected किया जाएगा.
    */

    const uploadYear =
      safe(
        $('uploadYear')?.value ||
        $('year')?.value
      );


    if (!uploadYear) {

      throw new Error(
        'Please select a Year before uploading dispatch data.'
      );

    }


    const existingIds =
      new Set(
        S.dispatches.map(
          row => row._id
        )
      );


    const result =
      await uploadDispatchRows(
        S.dispatchPreview,
        uploadYear,
        existingIds,
        (done, total) => {

          const output =
            $('upRes');


          if (output) {

            output.textContent =
              `Uploading ${done}/${total}...`;

          }

        }
      );


    const output =
      $('upRes');


    if (output) {

      output.textContent =
        `Inserted: ${result.inserted}, ` +
        `Updated: ${result.updated}, ` +
        `Skipped: ${result.skipped}, ` +
        `Errors: ${result.errors}`;

    }


    UI.showToast(
      'Dispatch upload completed.',
      'success'
    );


    await loadData();


  } catch (error) {

    showError(error);

  }

}


/* =========================================================
   MAPPING PREVIEW
========================================================= */

async function previewMappingFile() {

  try {

    const file =
      $('mFile')?.files?.[0];


    if (!file) {

      throw new Error(
        'Please select Mapping Excel file.'
      );

    }


    const rows =
      await readExcelFile(file);


    S.mappingPreview =
      rows;


    const info =
      $('mInfo');


    if (info) {

      info.textContent =
        `${rows.length} mapping rows loaded from ${file.name}`;

    }


    UI.showToast(
      `${rows.length} mapping rows ready.`,
      'success'
    );


  } catch (error) {

    showError(error);

  }

}


/* =========================================================
   MAPPING UPLOAD
========================================================= */

async function uploadMappingFile() {

  try {

    if (
      !S.mappingPreview.length
    ) {

      await previewMappingFile();

    }


    if (
      !S.mappingPreview.length
    ) {
      return;
    }


    const existingIds =
      new Set(
        S.mappings.map(
          row => row._id
        )
      );


    const result =
      await uploadMappingRows(
        S.mappingPreview,
        existingIds,
        (done, total) => {

          const output =
            $('mRes');


          if (output) {

            output.textContent =
              `Uploading ${done}/${total}...`;

          }

        }
      );


    const output =
      $('mRes');


    if (output) {

      output.textContent =
        `Inserted: ${result.inserted}, ` +
        `Updated: ${result.updated}, ` +
        `Skipped: ${result.skipped}, ` +
        `Errors: ${result.errors}`;

    }


    UI.showToast(
      'Mapping upload completed.',
      'success'
    );


    await loadData();


  } catch (error) {

    showError(error);

  }

}


/* =========================================================
   DATA VIEW
========================================================= */

function setupDataFilters() {

  const year =
    $('fYear');

  const district =
    $('fDist');

  const ps =
    $('fPs');


  if (year) {

    UI.fillSelect(
      year,

      getAvailableYears()
        .map(v => ({
          v,
          t: v
        })),

      'All Years'
    );

  }


  if (district) {

    UI.fillSelect(
      district,

      getAvailableDistricts()
        .map(v => ({
          v,
          t: v
        })),

      'All Districts'
    );

  }


  if (ps) {

    UI.fillSelect(
      ps,

      uniqueSorted(
        S.mappings
          .map(getMappingPS)
      )
        .map(v => ({
          v,
          t: v
        })),

      'All Panchayat Samitis'
    );

  }


  [
    year,
    district,
    ps,
    $('fStatus'),
    $('fSearch')
  ]
    .filter(Boolean)
    .forEach(element => {

      element.addEventListener(
        'change',
        renderDataView
      );

      element.addEventListener(
        'input',
        renderDataView
      );

    });

}


function renderDataView() {

  const output =
    $('dataOut');


  if (!output) {
    return;
  }


  const year =
    normalizeYear(
      $('fYear')?.value
    );


  const district =
    norm(
      $('fDist')?.value
    );


  const ps =
    norm(
      $('fPs')?.value
    );


  const status =
    norm(
      $('fStatus')?.value
    );


  const search =
    norm(
      $('fSearch')?.value
    );


  const rows =
    S.dispatches.filter(row => {

      const rowYear =
        normalizeYear(
          getDispatchYear(row)
        );


      if (
        year &&
        rowYear !== year
      ) {
        return false;
      }


      /*
        Data view में district/PS mapping से resolve करें.
      */

      const mapping =
        findMappingForRow(row);


      const rowDistrict =
        norm(
          mapping
            ? getMappingDistrict(mapping)
            : getDispatchDistrict(row)
        );


      const rowPS =
        norm(
          mapping
            ? getMappingPS(mapping)
            : getDispatchPS(row)
        );


      if (
        district &&
        rowDistrict !== district
      ) {
        return false;
      }


      if (
        ps &&
        rowPS !== ps
      ) {
        return false;
      }


      if (status) {

        const rowStatus =
          norm(
            row.Status ??
            row.status ??
            ''
          );


        if (
          rowStatus !== status
        ) {
          return false;
        }

      }


      if (search) {

        const text =
          norm(
            Object
              .values(row)
              .join(' ')
          );


        if (
          !text.includes(search)
        ) {
          return false;
        }

      }


      return true;

    });


  output.innerHTML =
    `<p>${rows.length} record(s)</p>` +
    makeTable(
      rows.slice(0, 200)
    );

}


/* =========================================================
   EVENTS
========================================================= */

function setupEvents() {

  /*
    Google Login
  */

  const loginButton =
    $('gBtn');


  if (loginButton) {

    loginButton.onclick =
      async () => {

        try {

          loginButton.disabled = true;

          loginButton.textContent =
            'Signing in...';


          await googleLogin();

        } catch (error) {

          console.error(error);


          const loginError =
            $('lerr');


          if (loginError) {

            loginError.textContent =
              error?.message ||
              'Login failed.';

          }

        } finally {

          loginButton.disabled = false;

          loginButton.textContent =
            'Continue with Google';

        }

      };

  }


  /*
    Logout
  */

  document.addEventListener(
    'click',
    event => {

      const button =
        event.target.closest(
          '[data-action="logout"],#logoutBtn'
        );


      if (!button) {
        return;
      }


      logout().catch(
        showError
      );

    }
  );


  /*
    Report types
  */

  document
    .querySelectorAll(
      '#rtypes [data-t]'
    )
    .forEach(button => {

      button.addEventListener(
        'click',
        () => {

          setReportType(
            button.dataset.t
          );

        }
      );

    });


  /*
    Year
  */

  const year =
    $('year');


  if (year) {

    year.addEventListener(
      'change',
      () => {

        S.year =
          normalizeYear(
            year.value
          );


        /*
          Year change होने पर
          District/PS/GP selection
          refresh करें.
        */

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

  const district =
    $('dist');


  if (district) {

    district.addEventListener(
      'change',
      () => {

        S.district =
          district.value;


        S.ps = '';
        S.gp = '';


        populatePS();
        populateGP();

        updateInfo();

      }
    );

  }


  /*
    Panchayat Samiti
  */

  const ps =
    $('ps');


  if (ps) {

    ps.addEventListener(
      'change',
      () => {

        S.ps =
          ps.value;


        S.gp = '';


        populateGP();

        updateInfo();

      }
    );

  }


  /*
    GP
  */

  const gp =
    $('gp');


  if (gp) {

    gp.addEventListener(
      'change',
      () => {

        S.gp =
          gp.value;


        updateInfo();

      }
    );

  }


  /*
    Generate
  */

  const generateButton =
    $('genBtn');


  if (generateButton) {

    generateButton.addEventListener(
      'click',
      () => (S.rtype === 'ams' ? generateAms() : generateReport())
    );

  }


  /*
    Dispatch Excel
  */

  const file =
    $('file');


  if (file) {

    file.addEventListener(
      'change',
      previewDispatchFile
    );

  }


  const previewButton =
    $('prevBtn');


  if (previewButton) {

    previewButton.addEventListener(
      'click',
      previewDispatchFile
    );

  }


  const uploadButton =
    $('upBtn');


  if (uploadButton) {

    uploadButton.addEventListener(
      'click',
      uploadDispatchFile
    );

  }


  /*
    Mapping Excel
  */

  const mappingFile =
    $('mFile');


  if (mappingFile) {

    mappingFile.addEventListener(
      'change',
      previewMappingFile
    );

  }


  const mappingButton =
    $('mBtn');


  if (mappingButton) {

    mappingButton.addEventListener(
      'click',
      uploadMappingFile
    );

  }


  /*
    Reload
  */

  const reload =
    $('reloadBtn');


  if (reload) {

    reload.addEventListener(
      'click',
      async () => {

        await loadData();

        setupDataFilters();

        renderDataView();

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


      if (
        view === 'generate'
      ) {

        /*
          Firebase data से dropdowns
          हमेशा refresh.
        */

        populateYear();
        populateDistrict();
        populatePS();
        populateGP();

        updateInfo();

      }


      if (
        view === 'data'
      ) {

        setupDataFilters();
        renderDataView();

      }

    }
  );

}


/* =========================================================
   AUTH
========================================================= */

function setupAuth() {

  onUser(
    async user => {

      console.log(
        'AUTH STATE:',
        user
      );


      if (user) {

        S.user =
          user;


        updateUserUI(
          user
        );


        /*
          Login के बाद blank-page fix.
        */

        showApp();


        /*
          Firestore data load.
        */

        await loadData();


        setupDataFilters();


      } else {

        S.user =
          null;


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
    'Jamidara initializing...'
  );


  /*
    Login screen initially.
  */

  showLogin();


  /*
    Events.
  */

  setupEvents();


  /*
    Navigation.
  */

  setupNavigation();


  /*
    Default report type.
    यह Firebase data नहीं है,
    सिर्फ report का UI default है.
  */

  initAms();

  setReportType(
    'ybc'
  );


  /*
    Firebase Auth listener.
  */

  setupAuth();

}


/* =========================================================
   START APPLICATION
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
