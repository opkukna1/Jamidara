// UI helpers: toasts, confirm dialog, drawer, view switching, string normalizing.

export const $ = id => document.getElementById(id);

export const esc = s =>
  String(s ?? '').replace(
    /[&<>"]/g,
    c => ({
      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;',
      '"': '&quot;'
    }[c])
  );


// Normalize names
export const norm = s =>
  String(s ?? '')
    .replace(
      /gram\s+panchayat|panchayat\s+samiti/gi,
      ''
    )
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();


export function showToast(msg, type = '') {

  const t =
    document.createElement('div');

  t.className =
    'toast ' + type;

  t.textContent =
    msg;

  const container =
    $('toasts');

  if (container) {
    container.appendChild(t);
  } else {
    alert(msg);
  }

  setTimeout(
    () => t.remove(),
    8000
  );
}


export function confirmDialog(msg) {

  return new Promise(res => {

    const d =
      document.createElement('dialog');

    d.innerHTML = `
      <p>${esc(msg)}</p>
      <button value="n">Cancel</button>
      <button class="primary" value="y">Confirm</button>
    `;

    d.onclick = e => {

      if (e.target.value) {

        d.close();

        res(
          e.target.value === 'y'
        );

        d.remove();
      }
    };

    document.body.appendChild(d);

    d.showModal();
  });
}


export function fillSelect(
  sel,
  items,
  placeholder
) {

  if (!sel) return;

  sel.innerHTML =
    (placeholder
      ? `<option value="">${esc(placeholder)}</option>`
      : '') +

    items
      .map(
        i =>
          `<option value="${esc(i.v)}">${esc(i.t)}</option>`
      )
      .join('');
}


export function setupNav(onView) {

  const toggle = o => {

    const drawer =
      $('drawer');

    const scrim =
      $('scrim');

    if (drawer) {
      drawer.classList.toggle(
        'open',
        o
      );
    }

    if (scrim) {
      scrim.classList.toggle(
        'open',
        o
      );
    }
  };


  const menuBtn =
    $('menuBtn');

  if (menuBtn) {

    menuBtn.onclick = () =>
      toggle(
        !$('drawer')
          .classList
          .contains('open')
      );
  }


  const scrim =
    $('scrim');

  if (scrim) {
    scrim.onclick = () =>
      toggle(false);
  }


  document.addEventListener(
    'click',
    e => {

      const el =
        e.target.closest &&
        e.target.closest(
          '[data-view]'
        );

      if (!el) return;

      const v =
        el.dataset.view;


      document
        .querySelectorAll('.view')
        .forEach(s =>
          s.classList.toggle(
            'active',
            s.id === 'v-' + v
          )
        );


      document
        .querySelectorAll(
          '#drawer a'
        )
        .forEach(a =>
          a.classList.toggle(
            'on',
            a.dataset.view === v
          )
        );


      window.scrollTo({
        top: 0
      });


      toggle(false);


      if (onView) {
        onView(v);
      }
    }
  );
}


// ============================================================
// REAL ERROR DISPLAY
// ============================================================

export function friendly(e) {

  console.error(
    '========== FRIENDLY ERROR =========='
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

  console.error(
    '===================================='
  );


  const m =
    String(
      e &&
      (
        e.code ||
        e.message
      ) ||
      e
    );


  // Firestore
  if (
    /permission-denied/i.test(m)
  ) {

    return (
      'Permission denied. ' +
      'Please login as admin or check Firestore rules.'
    );
  }


  // Network
  if (
    /unavailable|network|offline/i.test(m)
  ) {

    return (
      'Network problem. ' +
      'Please check your internet connection.'
    );
  }


  // Template missing
  if (
    /TEMPLATE_MISSING/i.test(m)
  ) {

    return (
      'Word template not found. ' +
      'Check template/covering-letter-template.docx'
    );
  }


  // Authentication
  if (
    /auth\/(invalid|wrong|user)/i.test(m)
  ) {

    return (
      'Incorrect email or password.'
    );
  }


  // Docxtemplater
  if (
    e?.properties?.errors?.length
  ) {

    return (
      'Word template error: ' +
      e.properties.errors
        .map(
          x =>
            x.properties?.explanation ||
            x.message ||
            String(x)
        )
        .join(' | ')
    );
  }


  // Normal JavaScript error
  if (e?.message) {

    return (
      'Generate Error: ' +
      e.message
    );
  }


  if (e?.code) {

    return (
      'Generate Error: ' +
      e.code
    );
  }


  return (
    'Generate Error: ' +
    String(e)
  );
}
