// ============================================================
// JAMIDARA - COVER LETTER GENERATOR
// ============================================================

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

import {
  generateDocx
} from './js/docx-generator.js';


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

  rtype: 'ybc',

  year: '2025-26',

  district: '',

  ps: '',

  gp: 'ALL',

  last: null,

  amsMode: '1',

  selectedAmsPs: []

};


// ============================================================
// HELPERS
// ============================================================

const $ = id =>
  document.getElementById(id);


function safe(v) {

  return v == null
    ? ''
    : String(v).trim();

}


function norm(v) {

  return safe(v)
    .replace(
      /gram\s+panchayat|panchayat\s+samiti/gi,
      ''
    )
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();

}


function slug(v) {

  return safe(v)
    .replace(/[^\w\u0900-\u097F-]+/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_+|_+$/g, '');

}


function today() {

  const d = new Date();

  const dd =
    String(d.getDate()).padStart(2, '0');

  const mm =
    String(d.getMonth() + 1).padStart(2, '0');

  const yyyy =
    d.getFullYear();

  return `${dd}-${mm}-${yyyy}`;

}


function isAdmin() {

  return !!(
    S.user &&
    ADMIN_EMAILS
      .map(x => x.toLowerCase())
      .includes(
        safe(S.user.email).toLowerCase()
      )
  );

}


function toast(
  msg,
  type = ''
) {

  try {

    UI.showToast(
      msg,
      type
    );

  } catch {

    alert(msg);

  }

}


// ============================================================
// AUTH
// ============================================================

onUser(async user => {

  console.log(
    'AUTH STATE:',
    user
  );

  S.user =
    user || null;


  if (!user) {

    showLogin();

    return;

  }


  showApp();


  setUserUI(user);


  try {

    await loadData();

  } catch (e) {

    console.error(
      'Initial data load failed:',
      e
    );

    toast(
      'Data load error: ' +
      getRealError(e),
      'err'
    );

  }

});


// ============================================================
// LOGIN SCREEN
// ============================================================

function showLogin() {

  const login =
    $('login');

  if (login) {
    login.style.display =
      'flex';
  }

  document
    .querySelectorAll(
      'header.top, nav#drawer, main'
    )
    .forEach(el => {

      el.style.display =
        'none';

    });

}


function showApp() {

  const login =
    $('login');

  if (login) {

    login.style.display =
      'none';

  }

  document
    .querySelectorAll(
      'header.top, nav#drawer, main'
    )
    .forEach(el => {

      el.style.display =
        '';

    });

}


// ============================================================
// GOOGLE LOGIN
// ============================================================

const googleButton =
  $('gBtn');


if (googleButton) {

  googleButton.addEventListener(
    'click',
    async () => {

      const err =
        $('lerr');

      if (err) {
        err.textContent = '';
      }


      googleButton.disabled =
        true;

      googleButton.style.opacity =
        '0.7';


      try {

        await googleLogin();

      } catch (e) {

        console.error(
          'GOOGLE LOGIN ERROR:',
          e
        );


        if (err) {

          err.textContent =
            getRealError(e);

        }

      } finally {

        googleButton.disabled =
          false;

        googleButton.style.opacity =
          '';

      }

    }
  );

}


// ============================================================
// LOGOUT
// ============================================================

const logoutButton =
  $('logoutBtn');


if (logoutButton) {

  logoutButton.addEventListener(
    'click',
    async () => {

      try {

        await logout();

      } catch (e) {

        console.error(
          'LOGOUT ERROR:',
          e
        );

      }

    }
  );

}


// ============================================================
// USER UI
// ============================================================

function setUserUI(user) {

  const name =
    safe(
      user.displayName ||
      user.email ||
      'User'
    );

  const email =
    safe(
      user.email
    );


  if ($('dname')) {

    $('dname').textContent =
      name;

  }


  if ($('demail')) {

    $('demail').textContent =
      email;

  }


  if ($('hello')) {

    $('hello').textContent =
      `नमस्ते ${name} 👋`;

  }


  const avatar =
    user.photoURL || '';


  if ($('avatar')) {

    $('avatar').src =
      avatar;

  }


  if ($('davatar')) {

    $('davatar').src =
      avatar;

  }


  const admin =
    isAdmin();


  const uploadNav =
    $('navUpload');


  const uploadTile =
    $('tileUpload');


  if (uploadNav) {

    uploadNav.style.display =
      admin
        ? ''
        : 'none';

  }


  if (uploadTile) {

    uploadTile.style.display =
      admin
        ? ''
        : 'none';

  }


  const noAdmin =
    $('noAdmin');


  const upBox =
    $('upBox');


  if (noAdmin) {

    noAdmin.style.display =
      admin
        ? 'none'
        : '';

  }


  if (upBox) {

    upBox.hidden =
      !admin;

  }

}


// ============================================================
// LOAD DATA
// ============================================================

async function loadData() {

  console.log(
    'Loading Firestore mappings...'
  );


  const mappings =
    await loadMappings();


  console.log(
    'Mappings:',
    mappings.length
  );


  const dispatches =
    await loadDispatchData();


  console.log(
    'Dispatches:',
    dispatches.length
  );


  S.mappings =
    mappings;


  S.dispatches =
    dispatches;


  updateStats();


  buildYears();


  buildDistricts();


  buildFilterControls();


  updateGenerateControls();


  updateInfo();

}


// ============================================================
// STATS
// ============================================================

function updateStats() {

  if ($('stD')) {

    $('stD').textContent =
      S.dispatches.length;

  }


  if ($('stM')) {

    $('stM').textContent =
      S.mappings.length;

  }


  const years =
    new Set();


  S.dispatches.forEach(d => {

    const y =
      safe(
        d.Year ||
        d.YEAR ||
        d.year
      );

    if (y) {
      years.add(y);
    }

  });


  S.mappings.forEach(d => {

    const y =
      safe(
        d.Year ||
        d.YEAR ||
        d.year
      );

    if (y) {
      years.add(y);
    }

  });


  if ($('stY')) {

    $('stY').textContent =
      years.size || '1';

  }

}


// ============================================================
// YEAR LIST
// ============================================================

function buildYears() {

  const years =
    new Set([
      '2025-26',
      '2026-27'
    ]);


  S.dispatches.forEach(d => {

    const y =
      safe(
        d.Year ||
        d.YEAR ||
        d.year
      );

    if (y) {
      years.add(y);
    }

  });


  const list =
    [...years].sort();


  if ($('year')) {

    $('year').innerHTML =
      list
        .map(y =>
          `<option value="${UI.esc(y)}">${UI.esc(y)}</option>`
        )
        .join('');


    if (
      list.includes(S.year)
    ) {

      $('year').value =
        S.year;

    } else {

      S.year =
        list[0] || '2025-26';

      $('year').value =
        S.year;

    }

  }


  if ($('upYear')) {

    $('upYear').innerHTML =
      list
        .map(y =>
          `<option value="${UI.esc(y)}">${UI.esc(y)}</option>`
        )
        .join('');

  }


  if ($('fYear')) {

    $('fYear').innerHTML =
      `<option value="">All Years</option>` +
      list
        .map(y =>
          `<option value="${UI.esc(y)}">${UI.esc(y)}</option>`
        )
        .join('');

  }

}


// ============================================================
// DISTRICT LIST
// ============================================================

function getDistrict(m) {

  return safe(
    m.DIST_EN ||
    m.DISTRICT_EN ||
    m.District ||
    m.district
  );

}


function getDistrictHi(m) {

  return safe(
    m.DIST_HI ||
    m.DISTRICT_HI ||
    m.District_HI ||
    m.district_hi
  );

}


function getPS(m) {

  return safe(
    m.PS_EN ||
    m.PS_NAME_EN ||
    m.PS ||
    m.ps
  );

}


function getPSHi(m) {

  return safe(
    m.PS_HI ||
    m.PS_NAME_HI ||
    m.PS_HINDI ||
    m.ps_hi
  );

}


function getGP(m) {

  return safe(
    m.GP_EN ||
    m.GP_NAME_EN ||
    m.GP ||
    m.gp
  );

}


function getGPHI(m) {

  return safe(
    m.GP_HI ||
    m.GP_NAME_HI ||
    m.GP_HINDI ||
    m.gp_hi
  );

}


function buildDistricts() {

  const list =
    [...new Set(
      S.mappings
        .map(getDistrict)
        .filter(Boolean)
    )].sort();


  const select =
    $('dist');


  if (!select) return;


  select.innerHTML =
    `<option value="">Select District</option>` +
    list
      .map(d =>
        `<option value="${UI.esc(d)}">${UI.esc(d)}</option>`
      )
      .join('');


  if (S.district) {

    select.value =
      S.district;

  }

}


// ============================================================
// DISTRICT CHANGE
// ============================================================

if ($('dist')) {

  $('dist').addEventListener(
    'change',
    () => {

      S.district =
        $('dist').value;

      S.ps =
        '';

      S.gp =
        'ALL';

      buildPS();

      buildGP();

      updateInfo();

    }
  );

}


// ============================================================
// PS LIST
// ============================================================

function buildPS() {

  const select =
    $('ps');


  if (!select) return;


  const list =
    [...new Set(
      S.mappings
        .filter(m =>
          norm(getDistrict(m)) ===
          norm(S.district)
        )
        .map(getPS)
        .filter(Boolean)
    )].sort();


  select.innerHTML =
    `<option value="">Select Panchayat Samiti</option>` +
    list
      .map(p =>
        `<option value="${UI.esc(p)}">${UI.esc(p)}</option>`
      )
      .join('');


  if (S.ps) {

    select.value =
      S.ps;

  }

}


// ============================================================
// PS CHANGE
// ============================================================

if ($('ps')) {

  $('ps').addEventListener(
    'change',
    () => {

      S.ps =
        $('ps').value;

      S.gp =
        'ALL';

      buildGP();

      updateInfo();

    }
  );

}


// ============================================================
// GP LIST
// ============================================================

function buildGP() {

  const select =
    $('gp');


  if (!select) return;


  const list =
    [...new Set(
      S.mappings
        .filter(m =>
          norm(getDistrict(m)) ===
          norm(S.district) &&
          norm(getPS(m)) ===
          norm(S.ps)
        )
        .map(getGP)
        .filter(Boolean)
    )].sort();


  select.innerHTML =
    `<option value="ALL">All Gram Panchayats</option>` +
    list
      .map(g =>
        `<option value="${UI.esc(g)}">${UI.esc(g)}</option>`
      )
      .join('');


  select.value =
    S.gp || 'ALL';

}


// ============================================================
// GP CHANGE
// ============================================================

if ($('gp')) {

  $('gp').addEventListener(
    'change',
    () => {

      S.gp =
        $('gp').value;

      updateInfo();

    }
  );

}


// ============================================================
// YEAR CHANGE
// ============================================================

if ($('year')) {

  $('year').addEventListener(
    'change',
    () => {

      S.year =
        $('year').value;

      updateInfo();

    }
  );

}


// ============================================================
// REPORT TYPE
// ============================================================

document
  .querySelectorAll(
    '#rtypes [data-t]'
  )
  .forEach(btn => {

    btn.addEventListener(
      'click',
      () => {

        document
          .querySelectorAll(
            '#rtypes [data-t]'
          )
          .forEach(x =>
            x.classList.remove('on')
          );


        btn.classList.add('on');


        S.rtype =
          btn.dataset.t;


        updateGenerateControls();

      }
    );

  });


// ============================================================
// REPORT TYPE UI
// ============================================================

function updateGenerateControls() {

  const amsBox =
    $('amsBox');


  const letterSel =
    $('letterSel');


  if (S.rtype === 'ams') {

    if (amsBox) {
      amsBox.hidden = false;
    }

    if (letterSel) {
      letterSel.hidden = true;
    }

    buildAmsPS();

  } else {

    if (amsBox) {
      amsBox.hidden = true;
    }

    if (letterSel) {
      letterSel.hidden = false;
    }

  }


  const btn =
    $('genBtn');


  if (!btn) return;


  if (S.rtype === 'ams') {

    btn.textContent =
      'Generate AMS Report';

  } else if (
    S.rtype === 'ybc'
  ) {

    btn.textContent =
      'Generate Cover Letter';

  } else {

    btn.textContent =
      'Generate Intimation Report';

  }

}


// ============================================================
// AMS PS
// ============================================================

function buildAmsPS() {

  const box =
    $('amsPs');


  if (!box) return;


  const psList =
    [...new Set(
      S.mappings
        .filter(m =>
          norm(getDistrict(m)) ===
          norm(S.district)
        )
        .map(getPS)
        .filter(Boolean)
    )].sort();


  box.innerHTML =
    psList
      .map((p, i) => {

        const checked =
          S.selectedAmsPs.includes(p)
            ? 'checked'
            : '';


        return `
          <label class="check">
            <input
              type="checkbox"
              value="${UI.esc(p)}"
              ${checked}
            >
            <span>${UI.esc(p)}</span>
          </label>
        `;

      })
      .join('');


  box
    .querySelectorAll(
      'input[type="checkbox"]'
    )
    .forEach(ch => {

      ch.addEventListener(
        'change',
        () => {

          S.selectedAmsPs =
            [
              ...box.querySelectorAll(
                'input:checked'
              )
            ]
              .map(x => x.value);

          updateInfo();

        }
      );

    });

}


// ============================================================
// AMS MODE
// ============================================================

document
  .querySelectorAll(
    '#amsMode [data-m]'
  )
  .forEach(btn => {

    btn.addEventListener(
      'click',
      () => {

        document
          .querySelectorAll(
            '#amsMode [data-m]'
          )
          .forEach(x =>
            x.classList.remove('on')
          );


        btn.classList.add('on');


        S.amsMode =
          btn.dataset.m;


        updateInfo();

      }
    );

  });


// ============================================================
// INFO
// ============================================================

function updateInfo() {

  const info =
    $('info');


  if (!info) return;


  if (
    S.rtype === 'ams'
  ) {

    info.innerHTML =
      `Select one or more Panchayat Samiti for AMS report.`;

    return;

  }


  if (!S.district) {

    info.innerHTML =
      'Please select District.';

    return;

  }


  if (!S.ps) {

    info.innerHTML =
      'Please select Panchayat Samiti.';

    return;

  }


  const rows =
    getSelectedMappings();


  info.innerHTML =
    `<b>${rows.length}</b> Gram Panchayat selected`;


  const warn =
    $('warn');


  if (warn) {

    warn.textContent =
      '';

  }

}


// ============================================================
// SELECTED MAPPINGS
// ============================================================

function getSelectedMappings() {

  return S.mappings.filter(m => {

    const districtOK =
      norm(getDistrict(m)) ===
      norm(S.district);


    const psOK =
      norm(getPS(m)) ===
      norm(S.ps);


    const gpOK =
      S.gp === 'ALL' ||
      norm(getGP(m)) ===
      norm(S.gp);


    return (
      districtOK &&
      psOK &&
      gpOK
    );

  });

}


// ============================================================
// DISPATCH HELPERS
// ============================================================

function getDispatchYear(d) {

  return safe(
    d.Year ||
    d.YEAR ||
    d.year
  );

}


function getDispatchGP(d) {

  return safe(
    d['GP Name'] ||
    d['Gram Panchayat'] ||
    d.GP_NAME_EN ||
    d.GP_EN ||
    d.GP ||
    d.gp
  );

}


function getDispatchPS(d) {

  return safe(
    d['PS Name'] ||
    d['Panchayat Samiti'] ||
    d.PS_NAME_EN ||
    d.PS_EN ||
    d.PS ||
    d.ps
  );

}


function getDispatchDistrict(d) {

  return safe(
    d.District ||
    d.DIST_EN ||
    d.DISTRICT_EN ||
    d.district
  );

}


function getDispatchNumber(d) {

  const keys = [

    'Dispatch No',
    'Dispatch Number',
    'Dispatch Sankhya',
    'dispatch no',
    'dispatch number',
    'dispatch sankhya',

    'DISPATCH_NO',
    'Dispatch_No',
    'DISPATCH',

    'Sankhya',
    'sankhya'

  ];


  for (
    const key of keys
  ) {

    if (
      safe(d[key])
    ) {

      return safe(d[key]);

    }

  }


  for (
    const key of Object.keys(d || {})
  ) {

    const k =
      key
        .toLowerCase();


    if (
      k.includes('dispatch') ||
      k.includes('sankhya')
    ) {

      if (
        safe(d[key])
      ) {

        return safe(d[key]);

      }

    }

  }


  return '';

}


function getDispatchName(d) {

  const keys = [

    'Dispatch Name',
    'Dispatch File',
    'DISPATCH_NAME',
    'Dispatch_Name',
    'Name',
    'NAME',
    'File',
    'file'

  ];


  for (
    const key of keys
  ) {

    if (
      safe(d[key])
    ) {

      return safe(d[key]);

    }

  }


  return '';

}


function getPara(d) {

  const keys = [

    'Para Sankhya',
    'Para Count',
    'Para',
    'para sankhya',
    'para count',

    'PARA_COUNT',
    'PARA',
    'Para_Count'

  ];


  for (
    const key of keys
  ) {

    if (
      safe(d[key])
    ) {

      return safe(d[key]);

    }

  }


  for (
    const key of Object.keys(d || {})
  ) {

    const k =
      key
        .toLowerCase();


    if (
      k.includes('para') ||
      k.includes('sankhya')
    ) {

      if (
        safe(d[key])
      ) {

        return safe(d[key]);

      }

    }

  }


  return '';

}


// ============================================================
// FIND DISPATCH
// ============================================================

function findDispatch(
  gp,
  ps,
  district,
  year
) {

  const gpN =
    norm(gp);

  const psN =
    norm(ps);

  const distN =
    norm(district);


  // Exact GP + year
  let found =
    S.dispatches.find(d => {

      const dg =
        norm(getDispatchGP(d));

      const dy =
        norm(getDispatchYear(d));


      return (
        dg === gpN &&
        (!year || dy === norm(year))
      );

    });


  if (found) {
    return found;
  }


  // GP + PS + year
  found =
    S.dispatches.find(d => {

      const dg =
        norm(getDispatchGP(d));

      const dp =
        norm(getDispatchPS(d));

      const dy =
        norm(getDispatchYear(d));


      return (
        dg === gpN &&
        (!dp || dp === psN) &&
        (!year || dy === norm(year))
      );

    });


  if (found) {
    return found;
  }


  // GP + district
  found =
    S.dispatches.find(d => {

      const dg =
        norm(getDispatchGP(d));

      const dd =
        norm(getDispatchDistrict(d));

      const dy =
        norm(getDispatchYear(d));


      return (
        dg === gpN &&
        (!dd || dd === distN) &&
        (!year || dy === norm(year))
      );

    });


  return found || null;

}


// ============================================================
// LETTER VALUES
// ============================================================

function letterValues(
  mapping
) {

  const gpEn =
    getGP(mapping);

  const gpHi =
    getGPHI(mapping);

  const psEn =
    getPS(mapping);

  const psHi =
    getPSHi(mapping);

  const distEn =
    getDistrict(mapping);

  const distHi =
    getDistrictHi(mapping);


  const dispatch =
    findDispatch(
      gpEn,
      psEn,
      distEn,
      S.year
    );


  return {

    YEAR:
      S.year,

    GP_NAME_EN:
      gpEn,

    GP_NAME_ENG:
      gpEn,

    GP_NAME_HI:
      gpHi,

    GP_NAME_HINDI:
      gpHi,

    GP_NAME:
      gpHi,

    GP_NAME_HI_EN:
      `${gpHi} (${gpEn})`,

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
      dispatch
        ? getDispatchNumber(dispatch)
        : '',

    DISPATCH_NAME:
      dispatch
        ? getDispatchName(dispatch)
        : '',

    DATE:
      today(),

    PARA_COUNT:
      dispatch
        ? getPara(dispatch)
        : '',

    PARA_BREAKUP:
      dispatch
        ? getPara(dispatch)
        : '',

    OFFICE_NAME:
      distHi,

    DIVISION_NAME:
      distHi,

    CONSTITUTION_OBJECTION:
      '0',

    SERIOUS_OBJECTION:
      '0'

  };

}


// ============================================================
// GENERATE BUTTON
// ============================================================

const genButton =
  $('genBtn');


if (genButton) {

  genButton.addEventListener(
    'click',
    generate
  );

}


// ============================================================
// GENERATE
// ============================================================

async function generate() {

  if (!S.user) {

    toast(
      'Please sign in first.',
      'err'
    );

    return;

  }


  genButton.disabled =
    true;

  const oldText =
    genButton.textContent;


  genButton.textContent =
    'Preparing your file...';


  try {

    console.log(
      '===== GENERATION START ====='
    );


    console.log(
      'User:',
      S.user.email
    );

    console.log(
      'Report:',
      S.rtype
    );

    console.log(
      'Year:',
      S.year
    );

    console.log(
      'District:',
      S.district
    );

    console.log(
      'PS:',
      S.ps
    );

    console.log(
      'GP:',
      S.gp
    );


    // --------------------------------------------------------
    // CURRENTLY SUPPORTED COVERING LETTER
    // --------------------------------------------------------

    if (
      S.rtype !== 'ybc'
    ) {

      throw new Error(
        `Template for "${S.rtype}" is not uploaded yet. ` +
        `For now select "Year Book Closing".`
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


    const mappings =
      getSelectedMappings();


    if (!mappings.length) {

      throw new Error(
        'No Gram Panchayat found for selected location.'
      );

    }


    console.log(
      'Selected GPs:',
      mappings.length
    );


    const values =
      mappings.map(
        letterValues
      );


    console.log(
      'First letter data:',
      values[0]
    );


    const report =
      REPORTS[S.rtype];


    console.log(
      'Using template:',
      report.tpl
    );


    const blob =
      await generateDocx(
        values,
        report.tpl
      );


    if (!blob) {

      throw new Error(
        'Word generator returned an empty file.'
      );

    }


    const filename =
      `${report.prefix}_${slug(S.district)}_${slug(S.ps)}_${slug(S.year)}.docx`;


    S.last = {

      blob,

      name:
        filename

    };


    if ($('rname')) {

      $('rname').textContent =
        filename;

    }


    if ($('result')) {

      $('result').hidden =
        false;

      $('result').scrollIntoView({
        behavior: 'smooth',
        block: 'center'
      });

    }


    toast(
      `Successfully generated ${values.length} letter(s).`,
      'ok'
    );


    console.log(
      '===== GENERATION SUCCESS ====='
    );


  } catch (e) {

    console.error(
      '===== GENERATION ERROR ====='
    );

    console.error(e);

    console.error(
      'message:',
      e?.message
    );

    console.error(
      'code:',
      e?.code
    );

    console.error(
      'properties:',
      e?.properties
    );

    console.error(
      'stack:',
      e?.stack
    );


    let message =
      getRealError(e);


    if (
      e?.properties?.errors?.length
    ) {

      message =
        e.properties.errors
          .map(x =>
            x.properties?.explanation ||
            x.message ||
            String(x)
          )
          .join(' | ');

    }


    toast(
      'Generate Error: ' + message,
      'err'
    );


  } finally {

    genButton.disabled =
      false;

    genButton.textContent =
      oldText;

  }

}


// ============================================================
// DOWNLOAD WORD
// ============================================================

const downloadButton =
  $('dlBtn');


if (downloadButton) {

  downloadButton.addEventListener(
    'click',
    () => {

      if (!S.last?.blob) {

        toast(
          'No generated file available.',
          'err'
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
        () =>
          URL.revokeObjectURL(url),
        1000
      );

    }
  );

}


// ============================================================
// WHATSAPP
// ============================================================

const waButton =
  $('waBtn');


if (waButton) {

  waButton.addEventListener(
    'click',
    () => {

      const text =
        S.last
          ? `Generated Word file: ${S.last.name}`
          : 'Covering letter generated.';


      window.open(
        `https://wa.me/?text=${encodeURIComponent(text)}`,
        '_blank'
      );

    }
  );

}


// ============================================================
// NAVIGATION
// ============================================================

UI.setupNav(
  view => {

    console.log(
      'View:',
      view
    );


    if (
      view === 'data'
    ) {

      renderDispatchData();

    }


    if (
      view === 'upload'
    ) {

      setupUploadView();

    }

  }
);


// ============================================================
// FILTER DATA
// ============================================================

function buildFilterControls() {

  const years =
    [...new Set(
      S.dispatches
        .map(getDispatchYear)
        .filter(Boolean)
    )].sort();


  if ($('fYear')) {

    $('fYear').innerHTML =
      `<option value="">All Years</option>` +
      years
        .map(y =>
          `<option value="${UI.esc(y)}">${UI.esc(y)}</option>`
        )
        .join('');

  }


  const districts =
    [...new Set(
      S.dispatches
        .map(getDispatchDistrict)
        .filter(Boolean)
    )].sort();


  if ($('fDist')) {

    $('fDist').innerHTML =
      `<option value="">All Districts</option>` +
      districts
        .map(x =>
          `<option value="${UI.esc(x)}">${UI.esc(x)}</option>`
        )
        .join('');

  }


  if ($('fPs')) {

    $('fPs').innerHTML =
      `<option value="">All Panchayat Samiti</option>`;

  }


  if ($('fStatus')) {

    $('fStatus').innerHTML = `
      <option value="">All Status</option>
      <option value="with">With Dispatch</option>
      <option value="without">Without Dispatch</option>
    `;

  }


  [
    'fYear',
    'fDist',
    'fPs',
    'fStatus',
    'fSearch'
  ]
    .forEach(id => {

      const el =
        $(id);

      if (!el) return;


      el.addEventListener(
        'input',
        renderDispatchData
      );

      el.addEventListener(
        'change',
        renderDispatchData
      );

    });

}


// ============================================================
// RENDER DISPATCH DATA
// ============================================================

function renderDispatchData() {

  const out =
    $('dataOut');


  if (!out) return;


  const year =
    safe(
      $('fYear')?.value
    );


  const district =
    safe(
      $('fDist')?.value
    );


  const ps =
    safe(
      $('fPs')?.value
    );


  const status =
    safe(
      $('fStatus')?.value
    );


  const search =
    norm(
      $('fSearch')?.value
    );


  let rows =
    S.dispatches.filter(d => {

      const y =
        getDispatchYear(d);

      const dist =
        getDispatchDistrict(d);

      const p =
        getDispatchPS(d);

      const gp =
        getDispatchGP(d);

      const no =
        getDispatchNumber(d);


      if (
        year &&
        norm(y) !== norm(year)
      ) {
        return false;
      }


      if (
        district &&
        norm(dist) !== norm(district)
      ) {
        return false;
      }


      if (
        ps &&
        norm(p) !== norm(ps)
      ) {
        return false;
      }


      if (status === 'with' && !no) {
        return false;
      }


      if (status === 'without' && no) {
        return false;
      }


      if (search) {

        const hay =
          norm(
            `${gp} ${p} ${dist} ${no}`
          );


        if (
          !hay.includes(search)
        ) {

          return false;

        }

      }


      return true;

    });


  rows =
    rows.slice(0, 500);


  if (!rows.length) {

    out.innerHTML =
      `<p class="muted">No records found.</p>`;

    return;

  }


  const keys =
    Object.keys(rows[0])
      .filter(
        k => k !== '_id'
      )
      .slice(0, 10);


  out.innerHTML = `
    <div class="scroll">
      <table>
        <thead>
          <tr>
            ${keys
              .map(k =>
                `<th>${UI.esc(k)}</th>`
              )
              .join('')}
          </tr>
        </thead>

        <tbody>

          ${rows
            .map(row => `
              <tr>
                ${keys
                  .map(k =>
                    `<td>${UI.esc(row[k])}</td>`
                  )
                  .join('')}
              </tr>
            `)
            .join('')}

        </tbody>
      </table>
    </div>

    <p class="muted">
      Showing ${rows.length} record(s)
    </p>
  `;

}


// ============================================================
// UPLOAD VIEW
// ============================================================

function setupUploadView() {

  if (!isAdmin()) {

    if ($('noAdmin')) {
      $('noAdmin').style.display =
        '';
    }

    if ($('upBox')) {
      $('upBox').hidden =
        true;
    }

    return;

  }


  if ($('noAdmin')) {

    $('noAdmin').style.display =
      'none';

  }


  if ($('upBox')) {

    $('upBox').hidden =
      false;

  }

}


// ============================================================
// DISPATCH FILE PREVIEW
// ============================================================

let dispatchRows = [];


const dispatchFile =
  $('file');


if (dispatchFile) {

  dispatchFile.addEventListener(
    'change',
    async () => {

      const file =
        dispatchFile.files?.[0];


      if (!file) return;


      try {

        dispatchRows =
          await readExcel(file);


        if ($('fileInfo')) {

          $('fileInfo').textContent =
            `${dispatchRows.length} rows loaded`;

        }


        if ($('prevBtn')) {

          $('prevBtn').disabled =
            !dispatchRows.length;

        }


        if ($('upBtn')) {

          $('upBtn').disabled =
            !dispatchRows.length;

        }


      } catch (e) {

        console.error(e);

        toast(
          'Excel error: ' +
          getRealError(e),
          'err'
        );

      }

    }
  );

}


// ============================================================
// EXCEL READER
// ============================================================

async function readExcel(file) {

  if (
    typeof XLSX ===
    'undefined'
  ) {

    throw new Error(
      'XLSX library is not loaded.'
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


  const sheet =
    workbook.Sheets[
      workbook.SheetNames[0]
    ];


  if (!sheet) {

    throw new Error(
      'Excel sheet not found.'
    );

  }


  return XLSX.utils.sheet_to_json(
    sheet,
    {
      defval: ''
    }
  );

}


// ============================================================
// PREVIEW BUTTON
// ============================================================

if ($('prevBtn')) {

  $('prevBtn').addEventListener(
    'click',
    () => {

      const preview =
        $('preview');


      if (!preview) return;


      const rows =
        dispatchRows.slice(0, 20);


      if (!rows.length) {

        preview.innerHTML =
          '<p>No data.</p>';

        return;

      }


      const keys =
        Object.keys(rows[0]);


      preview.innerHTML = `
        <table>
          <thead>
            <tr>
              ${keys
                .map(k =>
                  `<th>${UI.esc(k)}</th>`
                )
                .join('')}
            </tr>
          </thead>

          <tbody>

            ${rows
              .map(r => `
                <tr>
                  ${keys
                    .map(k =>
                      `<td>${UI.esc(r[k])}</td>`
                    )
                    .join('')}
                </tr>
              `)
              .join('')}

          </tbody>
        </table>
      `;

    }
  );

}


// ============================================================
// DISPATCH UPLOAD
// ============================================================

if ($('upBtn')) {

  $('upBtn').addEventListener(
    'click',
    async () => {

      if (!isAdmin()) {

        toast(
          'Admin access required.',
          'err'
        );

        return;

      }


      if (!dispatchRows.length) {

        toast(
          'Please select an Excel file first.',
          'err'
        );

        return;

      }


      const year =
        safe(
          $('upYear')?.value
        ) ||
        S.year;


      const button =
        $('upBtn');


      button.disabled =
        true;


      try {

        const existing =
          new Set(
            S.dispatches
              .map(d => d._id)
              .filter(Boolean)
          );


        const result =
          await uploadDispatchRows(
            dispatchRows,
            year,
            existing,
            (done, total) => {

              button.textContent =
                `Uploading ${done}/${total}`;

            }
          );


        toast(
          `Upload complete. Added: ${result.inserted}, Updated: ${result.updated}, Skipped: ${result.skipped}`,
          'ok'
        );


        button.textContent =
          'Upload / Update Firestore';


        await loadData();


      } catch (e) {

        console.error(e);

        toast(
          'Upload error: ' +
          getRealError(e),
          'err'
        );

      } finally {

        button.disabled =
          false;

        button.textContent =
          'Upload / Update Firestore';

      }

    }
  );

}


// ============================================================
// MAPPING FILE
// ============================================================

let mappingRows =
  [];


const mappingFile =
  $('mFile');


if (mappingFile) {

  mappingFile.addEventListener(
    'change',
    async () => {

      const file =
        mappingFile.files?.[0];


      if (!file) return;


      try {

        mappingRows =
          await readExcel(file);


        if ($('mInfo')) {

          $('mInfo').textContent =
            `${mappingRows.length} mapping rows loaded`;

        }


        if ($('mBtn')) {

          $('mBtn').disabled =
            !mappingRows.length;

        }


      } catch (e) {

        console.error(e);

        toast(
          'Mapping Excel error: ' +
          getRealError(e),
          'err'
        );

      }

    }
  );

}


// ============================================================
// MAPPING UPLOAD
// ============================================================

if ($('mBtn')) {

  $('mBtn').addEventListener(
    'click',
    async () => {

      if (!isAdmin()) {

        toast(
          'Admin access required.',
          'err'
        );

        return;

      }


      if (!mappingRows.length) {

        toast(
          'Please select mapping Excel first.',
          'err'
        );

        return;

      }


      const button =
        $('mBtn');


      button.disabled =
        true;


      try {

        const existing =
          new Set(
            S.mappings
              .map(d => d._id)
              .filter(Boolean)
          );


        const result =
          await uploadMappingRows(
            mappingRows,
            existing,
            (done, total) => {

              button.textContent =
                `Uploading ${done}/${total}`;

            }
          );


        toast(
          `Mapping upload complete. Added: ${result.inserted}, Updated: ${result.updated}, Skipped: ${result.skipped}`,
          'ok'
        );


        button.textContent =
          'Upload Mapping to Firestore';


        await loadData();


      } catch (e) {

        console.error(e);

        toast(
          'Mapping upload error: ' +
          getRealError(e),
          'err'
        );

      } finally {

        button.disabled =
          false;

        button.textContent =
          'Upload Mapping to Firestore';

      }

    }
  );

}


// ============================================================
// RELOAD DATA
// ============================================================

if ($('reloadBtn')) {

  $('reloadBtn').addEventListener(
    'click',
    async () => {

      try {

        $('reloadBtn').disabled =
          true;

        $('reloadBtn').textContent =
          'Reloading...';


        await loadData();


        toast(
          'Data reloaded successfully.',
          'ok'
        );


      } catch (e) {

        toast(
          'Reload error: ' +
          getRealError(e),
          'err'
        );


      } finally {

        $('reloadBtn').disabled =
          false;

        $('reloadBtn').textContent =
          '🔄 Reload data from Firestore';

      }

    }
  );

}


// ============================================================
// REAL ERROR
// ============================================================

function getRealError(e) {

  if (
    e?.properties?.errors?.length
  ) {

    return e.properties.errors
      .map(x =>
        x.properties?.explanation ||
        x.message ||
        String(x)
      )
      .join(' | ');

  }


  if (e?.message) {

    return e.message;

  }


  if (e?.code) {

    return e.code;

  }


  return String(e);

}


// ============================================================
// INITIAL STATE
// ============================================================

console.log(
  'Jamidara app.js loaded successfully.'
);

console.log(
  'Waiting for Firebase authentication...'
);
