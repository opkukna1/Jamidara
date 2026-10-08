import { db, auth } from './js/firebase-config.js';
import {
  collection,
  getDocs,
  query,
  where,
  orderBy
} from 'https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js';

import {
  signInWithEmailAndPassword,
  onAuthStateChanged,
  signOut
} from 'https://www.gstatic.com/firebasejs/10.12.2/firebase-auth.js';

import * as UI from './js/ui.js';
import { generateDocx } from './js/docx-generator.js';


// ============================================================
// CONFIG
// ============================================================

const ADMIN_EMAILS = [
  'opsiddh42@gmail.com'
];

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


// ============================================================
// STATE
// ============================================================

const S = {
  user: null,
  mappings: [],
  dispatches: [],
  filteredMappings: [],
  rtype: 'ams',
  year: '2025-26',
  district: '',
  ps: '',
  gp: 'ALL',
  last: null
};


// ============================================================
// SHORT HELPERS
// ============================================================

const $ = id => document.getElementById(id);

function safe(v) {
  return v == null ? '' : String(v).trim();
}

function norm(v) {
  return safe(v)
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim();
}

function slug(v) {
  return safe(v)
    .replace(/[^\w\u0900-\u097F-]+/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_+|_+$/g, '');
}

function today() {
  const d = new Date();

  const dd = String(d.getDate()).padStart(2, '0');
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const yy = d.getFullYear();

  return `${dd}-${mm}-${yy}`;
}

function showError(msg) {
  console.error(msg);
  UI.showToast(msg, 'err');
}


// ============================================================
// AUTH
// ============================================================

onAuthStateChanged(auth, async user => {

  S.user = user || null;

  if (!user) {
    console.log('No logged-in user');

    if ($('login')) {
      $('login').hidden = false;
    }

    if ($('app')) {
      $('app').hidden = true;
    }

    return;
  }

  console.log('Logged in:', user.email);

  if ($('login')) {
    $('login').hidden = true;
  }

  if ($('app')) {
    $('app').hidden = false;
  }

  await loadData();
});


// ============================================================
// LOGIN
// ============================================================

if ($('loginForm')) {

  $('loginForm').addEventListener('submit', async e => {

    e.preventDefault();

    const email = safe($('email')?.value);
    const password = $('password')?.value || '';

    if (!email || !password) {
      showError('Email और password डालें.');
      return;
    }

    try {

      await signInWithEmailAndPassword(
        auth,
        email,
        password
      );

    } catch (e) {

      console.error('LOGIN ERROR:', e);

      showError(
        e?.message ||
        'Login failed.'
      );
    }
  });
}


// ============================================================
// LOGOUT
// ============================================================

if ($('logoutBtn')) {

  $('logoutBtn').addEventListener('click', async () => {

    try {
      await signOut(auth);
    } catch (e) {
      console.error(e);
    }

  });
}


// ============================================================
// LOAD FIRESTORE DATA
// ============================================================

async function loadData() {

  try {

    console.log('Loading mappings...');

    const mappingSnap = await getDocs(
      collection(db, 'mappings')
    );

    S.mappings = mappingSnap.docs.map(d => ({
      id: d.id,
      ...d.data()
    }));

    console.log(
      'Mappings loaded:',
      S.mappings.length
    );


    console.log('Loading dispatches...');

    const dispatchSnap = await getDocs(
      collection(db, 'dispatches')
    );

    S.dispatches = dispatchSnap.docs.map(d => ({
      id: d.id,
      ...d.data()
    }));

    console.log(
      'Dispatches loaded:',
      S.dispatches.length
    );


    buildDistricts();

    updateUI();

  } catch (e) {

    console.error('FIRESTORE LOAD ERROR:', e);

    UI.showToast(
      e?.message ||
      'Firestore data load failed.',
      'err'
    );
  }
}


// ============================================================
// DISTRICT
// ============================================================

function buildDistricts() {

  const districts = [
    ...new Set(
      S.mappings
        .map(x =>
          safe(
            x.DIST_EN ||
            x.DISTRICT_EN ||
            x.district ||
            x.District
          )
        )
        .filter(Boolean)
    )
  ].sort();

  const select = $('district');

  if (!select) return;

  select.innerHTML =
    `<option value="">Select District</option>` +
    districts
      .map(d =>
        `<option value="${escapeHtml(d)}">${escapeHtml(d)}</option>`
      )
      .join('');

}


// ============================================================
// DISTRICT CHANGE
// ============================================================

if ($('district')) {

  $('district').addEventListener(
    'change',
    () => {

      S.district = $('district').value;
      S.ps = '';
      S.gp = 'ALL';

      buildPS();
      buildGP();

      updateCount();
    }
  );
}


// ============================================================
// BUILD PS
// ============================================================

function buildPS() {

  const select = $('ps');

  if (!select) return;

  const rows = S.mappings.filter(x => {

    const district =
      safe(
        x.DIST_EN ||
        x.DISTRICT_EN ||
        x.district ||
        x.District
      );

    return norm(district) === norm(S.district);
  });


  const psList = [
    ...new Set(
      rows
        .map(x =>
          safe(
            x.PS_EN ||
            x.PS_NAME_EN ||
            x.PS ||
            x.ps
          )
        )
        .filter(Boolean)
    )
  ].sort();


  select.innerHTML =
    `<option value="">Select Panchayat Samiti</option>` +
    psList
      .map(p =>
        `<option value="${escapeHtml(p)}">${escapeHtml(p)}</option>`
      )
      .join('');
}


// ============================================================
// PS CHANGE
// ============================================================

if ($('ps')) {

  $('ps').addEventListener(
    'change',
    () => {

      S.ps = $('ps').value;
      S.gp = 'ALL';

      buildGP();

      updateCount();
    }
  );
}


// ============================================================
// BUILD GP
// ============================================================

function buildGP() {

  const select = $('gp');

  if (!select) return;


  const rows = S.mappings.filter(x => {

    const district =
      safe(
        x.DIST_EN ||
        x.DISTRICT_EN ||
        x.district ||
        x.District
      );

    const ps =
      safe(
        x.PS_EN ||
        x.PS_NAME_EN ||
        x.PS ||
        x.ps
      );

    return (
      norm(district) === norm(S.district) &&
      norm(ps) === norm(S.ps)
    );
  });


  const gps = [
    ...new Set(
      rows
        .map(x =>
          safe(
            x.GP_EN ||
            x.GP_NAME_EN ||
            x.GP ||
            x.gp
          )
        )
        .filter(Boolean)
    )
  ].sort();


  select.innerHTML =
    `<option value="ALL">All Gram Panchayats</option>` +
    gps
      .map(g =>
        `<option value="${escapeHtml(g)}">${escapeHtml(g)}</option>`
      )
      .join('');
}


// ============================================================
// GP CHANGE
// ============================================================

if ($('gp')) {

  $('gp').addEventListener(
    'change',
    () => {

      S.gp = $('gp').value;

      updateCount();
    }
  );
}


// ============================================================
// YEAR
// ============================================================

if ($('year')) {

  $('year').addEventListener(
    'change',
    () => {

      S.year = $('year').value;

      updateCount();
    }
  );
}


// ============================================================
// REPORT TYPE
// ============================================================

document
  .querySelectorAll('[data-t]')
  .forEach(btn => {

    btn.addEventListener(
      'click',
      () => {

        document
          .querySelectorAll('[data-t]')
          .forEach(x =>
            x.classList.remove('active')
          );

        btn.classList.add('active');

        S.rtype =
          btn.dataset.t ||
          'ams';

        const title =
          $('reportTitle');

        if (title) {

          const names = {
            ams: 'AMS Report',
            ybc: 'Year Book Closing',
            intim: 'Intimation Report'
          };

          title.textContent =
            names[S.rtype] ||
            'Report';
        }

        updateUI();
      }
    );
  });


// ============================================================
// UPDATE COUNT
// ============================================================

function updateCount() {

  const rows = S.mappings.filter(x => {

    const district =
      safe(
        x.DIST_EN ||
        x.DISTRICT_EN ||
        x.district ||
        x.District
      );

    const ps =
      safe(
        x.PS_EN ||
        x.PS_NAME_EN ||
        x.PS ||
        x.ps
      );

    const gp =
      safe(
        x.GP_EN ||
        x.GP_NAME_EN ||
        x.GP ||
        x.gp
      );

    const districtOK =
      !S.district ||
      norm(district) === norm(S.district);

    const psOK =
      !S.ps ||
      norm(ps) === norm(S.ps);

    const gpOK =
      S.gp === 'ALL' ||
      !S.gp ||
      norm(gp) === norm(S.gp);

    return districtOK && psOK && gpOK;
  });


  S.filteredMappings = rows;


  if ($('count')) {
    $('count').textContent =
      rows.length;
  }

  if ($('gpCount')) {
    $('gpCount').textContent =
      rows.length;
  }
}


// ============================================================
// UPDATE UI
// ============================================================

function updateUI() {

  updateCount();

  const btn = $('genBtn');

  if (btn) {
    btn.textContent =
      genLabel();
  }
}


// ============================================================
// GENERATE BUTTON
// ============================================================

if ($('genBtn')) {

  $('genBtn').addEventListener(
    'click',
    () => {

      withBusy(
        generateLetters
      );

    }
  );
}


// ============================================================
// GENERATE LABEL
// ============================================================

function genLabel() {

  if (S.rtype === 'ams') {
    return 'Generate AMS Report';
  }

  if (S.rtype === 'ybc') {
    return 'Generate Cover Letter';
  }

  if (S.rtype === 'intim') {
    return 'Generate Intimation Report';
  }

  return 'Generate Report';
}


// ============================================================
// GENERATE LETTERS
// ============================================================

async function generateLetters() {

  console.log('================================');
  console.log('START GENERATION');
  console.log('Report Type:', S.rtype);
  console.log('Year:', S.year);
  console.log('District:', S.district);
  console.log('PS:', S.ps);
  console.log('GP:', S.gp);
  console.log('================================');


  if (!S.user) {
    throw new Error(
      'Please login first.'
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


  const report =
    REPORTS[S.rtype];


  if (!report) {
    throw new Error(
      'Invalid report type: ' +
      S.rtype
    );
  }


  console.log(
    'Template:',
    report.tpl
  );


  // ----------------------------------------------------------
  // FILTER MAPPINGS
  // ----------------------------------------------------------

  const rows =
    S.mappings.filter(m => {

      const district =
        safe(
          m.DIST_EN ||
          m.DISTRICT_EN ||
          m.district ||
          m.District
        );

      const ps =
        safe(
          m.PS_EN ||
          m.PS_NAME_EN ||
          m.PS ||
          m.ps
        );

      const gp =
        safe(
          m.GP_EN ||
          m.GP_NAME_EN ||
          m.GP ||
          m.gp
        );


      const districtOK =
        norm(district) ===
        norm(S.district);


      const psOK =
        norm(ps) ===
        norm(S.ps);


      const gpOK =
        S.gp === 'ALL' ||
        norm(gp) === norm(S.gp);


      return (
        districtOK &&
        psOK &&
        gpOK
      );
    });


  console.log(
    'Filtered mapping rows:',
    rows.length
  );


  if (!rows.length) {
    throw new Error(
      'No Gram Panchayat mapping found.'
    );
  }


  // ----------------------------------------------------------
  // CREATE VALUES
  // ----------------------------------------------------------

  const values =
    rows.map(m => {

      const gpEn =
        safe(
          m.GP_EN ||
          m.GP_NAME_EN ||
          m.GP ||
          m.gp
        );


      const gpHi =
        safe(
          m.GP_HI ||
          m.GP_NAME_HI ||
          m.GP_HINDI ||
          m.gp_hi
        );


      const psEn =
        safe(
          m.PS_EN ||
          m.PS_NAME_EN ||
          m.PS ||
          m.ps
        );


      const psHi =
        safe(
          m.PS_HI ||
          m.PS_NAME_HI ||
          m.PS_HINDI ||
          m.ps_hi
        );


      const distEn =
        safe(
          m.DIST_EN ||
          m.DISTRICT_EN ||
          m.district ||
          m.District
        );


      const distHi =
        safe(
          m.DIST_HI ||
          m.DISTRICT_HI ||
          m.DISTRICT_HINDI ||
          m.dist_hi
        );


      const dispatch =
        findDispatch(
          gpEn,
          psEn,
          distEn
        );


      const dispatchNo =
        dispatch
          ? getDispatchNumber(dispatch)
          : '';


      const dispatchName =
        dispatch
          ? getDispatchName(dispatch)
          : '';


      const para =
        dispatch
          ? getPara(dispatch)
          : '';


      const englishName =
        gpEn;


      const hindiName =
        gpHi;


      const values = {

        YEAR: S.year,

        GP_NAME_EN:
          gpEn,

        GP_NAME_ENG:
          englishName,

        GP_NAME_HI:
          hindiName,

        GP_NAME_HINDI:
          hindiName,

        GP_NAME:
          hindiName,

        PS_NAME_EN:
          psEn,

        PS_NAME_HI:
          psHi,

        PS_NAME:
          psHi,

        DISTRICT_EN:
          distEn,

        DISTRICT_HI:
          distHi,

        DISTRICT:
          distHi,

        DISPATCH_NO:
          dispatchNo,

        DISPATCH_NAME:
          dispatchName,

        DATE:
          today(),

        PARA_COUNT:
          para,

        OFFICE_NAME:
          distHi,

        DIVISION_NAME:
          distHi,

        CONSTITUTION_OBJECTION:
          '0',

        SERIOUS_OBJECTION:
          '0',

        PARA_BREAKUP:
          para
      };


      return values;
    });


  console.log(
    'Records to generate:',
    values.length
  );


  console.log(
    'FIRST RECORD:',
    values[0]
  );


  // ----------------------------------------------------------
  // GENERATE DOCX
  // ----------------------------------------------------------

  console.log(
    'Calling generateDocx...'
  );


  const blob =
    await generateDocx(
      values,
      report.tpl
    );


  console.log(
    'generateDocx returned:',
    blob
  );


  if (!blob) {
    throw new Error(
      'Document generator returned empty file.'
    );
  }


  const filename =
    `${report.prefix}_${slug(S.district)}_${slug(S.ps)}_${S.year}.docx`;


  console.log(
    'Final filename:',
    filename
  );


  return {
    blob,
    name: filename
  };
}


// ============================================================
// DISPATCH FINDER
// ============================================================

function findDispatch(
  gp,
  ps,
  district
) {

  const gpN =
    norm(gp);

  const psN =
    norm(ps);

  const distN =
    norm(district);


  // First try exact GP
  let found =
    S.dispatches.find(d => {

      const dgp =
        norm(
          d.GP_NAME_EN ||
          d.GP_EN ||
          d.GP ||
          d.gp ||
          d['GP Name'] ||
          d['Gram Panchayat']
        );


      return dgp === gpN;
    });


  if (found) {
    return found;
  }


  // Try GP + PS
  found =
    S.dispatches.find(d => {

      const dgp =
        norm(
          d.GP_NAME_EN ||
          d.GP_EN ||
          d.GP ||
          d.gp ||
          d['GP Name'] ||
          d['Gram Panchayat']
        );


      const dps =
        norm(
          d.PS_NAME_EN ||
          d.PS_EN ||
          d.PS ||
          d.ps ||
          d['PS Name'] ||
          d['Panchayat Samiti']
        );


      return (
        dgp === gpN &&
        (!dps || dps === psN)
      );
    });


  if (found) {
    return found;
  }


  // Try GP + district
  found =
    S.dispatches.find(d => {

      const dgp =
        norm(
          d.GP_NAME_EN ||
          d.GP_EN ||
          d.GP ||
          d.gp ||
          d['GP Name'] ||
          d['Gram Panchayat']
        );


      const ddist =
        norm(
          d.DIST_EN ||
          d.DISTRICT_EN ||
          d.DISTRICT ||
          d.district ||
          d['District']
        );


      return (
        dgp === gpN &&
        (!ddist || ddist === distN)
      );
    });


  return found || null;
}


// ============================================================
// DISPATCH NUMBER
// ============================================================

function getDispatchNumber(d) {

  const candidates = [

    d.DISPATCH_NO,
    d.Dispatch_No,
    d.DISPATCH,
    d.Dispatch,
    d['Dispatch No'],
    d['Dispatch Number'],
    d['Dispatch Sankhya'],
    d['dispatch no'],
    d['dispatch number'],
    d['dispatch sankhya'],
    d.Sankhya,
    d.sankhya

  ];


  for (const x of candidates) {

    if (safe(x)) {
      return safe(x);
    }

  }


  // fallback: scan object keys
  for (const key of Object.keys(d || {})) {

    const k =
      key
        .toLowerCase()
        .replace(/\s+/g, ' ');


    if (
      k.includes('dispatch') ||
      k.includes('sankhya')
    ) {

      if (safe(d[key])) {
        return safe(d[key]);
      }
    }
  }


  return '';
}


// ============================================================
// DISPATCH NAME
// ============================================================

function getDispatchName(d) {

  const candidates = [

    d.DISPATCH_NAME,
    d.Dispatch_Name,
    d.NAME,
    d.Name,
    d['Dispatch Name'],
    d['Dispatch File'],
    d['Dispatch'],
    d.File,
    d.file

  ];


  for (const x of candidates) {

    if (safe(x)) {
      return safe(x);
    }

  }


  return '';
}


// ============================================================
// PARA
// ============================================================

function getPara(d) {

  const candidates = [

    d.PARA_COUNT,
    d.PARA,
    d.Para,
    d['Para Sankhya'],
    d['Para Count'],
    d['Para'],
    d['para sankhya'],
    d['para count']

  ];


  for (const x of candidates) {

    if (safe(x)) {
      return safe(x);
    }

  }


  // Search dynamically
  for (const key of Object.keys(d || {})) {

    const k =
      key
        .toLowerCase()
        .replace(/\s+/g, ' ');


    if (
      k.includes('para') ||
      k.includes('sankhya')
    ) {

      if (safe(d[key])) {
        return safe(d[key]);
      }
    }
  }


  return '';
}


// ============================================================
// BUSY / ERROR HANDLER
// ============================================================

async function withBusy(work) {

  const btn =
    $('genBtn');


  if (btn) {

    btn.disabled = true;

    btn.textContent =
      'Preparing your file...';
  }


  try {

    console.log(
      '========== GENERATE START =========='
    );


    const out =
      await work();


    console.log(
      'GENERATION OUTPUT:',
      out
    );


    if (out) {

      S.last =
        out;


      if ($('rname')) {

        $('rname').textContent =
          out.name;
      }


      if ($('result')) {

        $('result').hidden =
          false;

        $('result').scrollIntoView({
          behavior: 'smooth',
          block: 'center'
        });
      }


      showToast(
        'Generated successfully.',
        'ok'
      );
    }


  } catch (e) {

    // VERY IMPORTANT:
    // Do NOT hide the real error.

    console.error(
      '========== GENERATE ERROR =========='
    );

    console.error(
      e
    );

    console.error(
      'Message:',
      e?.message
    );

    console.error(
      'Code:',
      e?.code
    );

    console.error(
      'Properties:',
      e?.properties
    );

    console.error(
      'Stack:',
      e?.stack
    );

    console.error(
      '===================================='
    );


    let msg = 'Unknown error';


    // Docxtemplater errors
    if (
      e?.properties?.errors?.length
    ) {

      msg =
        e.properties.errors
          .map(x =>
            x.properties?.explanation ||
            x.message ||
            String(x)
          )
          .join(' | ');
    }


    // Normal JS error
    else if (e?.message) {

      msg =
        e.message;
    }


    else if (e?.code) {

      msg =
        e.code;
    }


    else {

      msg =
        String(e);
    }


    showToast(
      'Generate Error: ' + msg,
      'err'
    );

  } finally {

    if (btn) {

      btn.disabled = false;

      btn.textContent =
        genLabel();
    }
  }
}


// ============================================================
// DOWNLOAD
// ============================================================

if ($('downloadBtn')) {

  $('downloadBtn').addEventListener(
    'click',
    () => {

      if (!S.last?.blob) {

        showError(
          'No generated file available.'
        );

        return;
      }


      const url =
        URL.createObjectURL(
          S.last.blob
        );


      const a =
        document.createElement('a');


      a.href =
        url;

      a.download =
        S.last.name;


      document.body.appendChild(a);

      a.click();

      a.remove();


      setTimeout(
        () => URL.revokeObjectURL(url),
        1000
      );
    }
  );
}


// ============================================================
// TOAST
// ============================================================

function showToast(
  message,
  type = 'ok'
) {

  if (
    UI &&
    typeof UI.showToast === 'function'
  ) {

    UI.showToast(
      message,
      type
    );

    return;
  }


  alert(message);
}


// ============================================================
// HTML ESCAPE
// ============================================================

function escapeHtml(value) {

  return safe(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}


// ============================================================
// INITIAL UI
// ============================================================

updateUI();

console.log(
  'Jamidara application loaded.'
);

console.log(
  'Report configuration:',
  REPORTS
);
