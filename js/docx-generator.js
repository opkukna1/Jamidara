// Browser port of CoverGeneretor.py: fills the template, applies the same page-fit tweaks
// (date on right corner, signature gap, line spacing, extra gap), bolds the GP name, merges letters.
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
export const SETTINGS = { EXTRA_SPACE_PT: 3, LINE_SPACING: 1.3, SIGN_SPACE_PT: 25 }; // same as the Python script
const ORDER = ['pStyle','keepNext','keepLines','pageBreakBefore','framePr','widowControl','numPr','suppressLineNumbers','pBdr','shd','tabs','suppressAutoHyphens','kinsoku','wordWrap','overflowPunct','topLinePunct','autoSpaceDE','autoSpaceDN','bidi','adjustRightInd','snapToGrid','spacing','ind','contextualSpacing','mirrorIndents','suppressOverlap','jc','textDirection','textAlignment','textboxTightWrap','outlineLvl','divId','cnfStyle','rPr','sectPr','pPrChange'];

const kids = (el, n) => [...el.children].filter(c => c.localName === n);
const ptext = el => [...el.getElementsByTagNameNS(W, 't')].map(t => t.textContent).join('');
const wa = (el, n, v) => el.setAttributeNS(W, 'w:' + n, String(v));
function pPrOf(p) { let x = kids(p, 'pPr')[0]; if (!x) { x = p.ownerDocument.createElementNS(W, 'w:pPr'); p.insertBefore(x, p.firstChild); } return x; }
function pChild(p, name) { // get/create a pPr child in schema order
  const pr = pPrOf(p); let el = kids(pr, name)[0]; if (el) return el;
  el = p.ownerDocument.createElementNS(W, 'w:' + name);
  const i = ORDER.indexOf(name), next = [...pr.children].find(c => ORDER.indexOf(c.localName) > i);
  pr.insertBefore(el, next || null); return el;
}
function boldRun(r) {
  let rp = kids(r, 'rPr')[0];
  if (!rp) { rp = r.ownerDocument.createElementNS(W, 'w:rPr'); r.insertBefore(rp, r.firstChild); }
  let b = kids(rp, 'b')[0];
  if (b) { b.removeAttributeNS(W, 'val'); return; }
  b = r.ownerDocument.createElementNS(W, 'w:b');
  const after = [...rp.children].filter(c => ['rStyle', 'rFonts'].includes(c.localName)).pop();
  rp.insertBefore(b, after ? after.nextSibling : rp.firstChild);
}
function render(buf, data) {
  const z = new PizZip(buf);
  const f = z.file('word/document.xml'); // {{(GP_NAME_ENG)}} -> {{GP_NAME_ENG}}
  z.file('word/document.xml', f.asText().replace(/\{\{\(([A-Za-z_]+)\)\}\}/g, '{{$1}}'));
  const d = new docxtemplater(z, { delimiters: { start: '{{', end: '}}' }, paragraphLoop: true, linebreaks: true, nullGetter: () => '' });
  d.render(data); return d.getZip();
}
function isEmptyP(el) {
  if (el.localName !== 'p' || ptext(el).trim()) return false;
  return !['drawing', 'pict', 'br', 'object', 'sectPr', 'fldChar', 'fldSimple'].some(n => el.getElementsByTagNameNS(W, n).length);
}
function postProcess(zip, v) {
  const dom = new DOMParser().parseFromString(zip.file('word/document.xml').asText(), 'application/xml');
  const body = dom.getElementsByTagNameNS(W, 'body')[0];
  for (const r of dom.getElementsByTagNameNS(W, 'r')) { // bold the village name (Hindi + English)
    const t = ptext(r); if (t && ((v.GP_NAME_HI && t.includes(v.GP_NAME_HI)) || (v.GP_NAME_ENG && t.includes(v.GP_NAME_ENG)))) boldRun(r);
  }
  const sect = kids(body, 'sectPr')[0], pgSz = sect && kids(sect, 'pgSz')[0], pgMar = sect && kids(sect, 'pgMar')[0];
  const g = (el, n, d) => el && el.getAttributeNS(W, n) ? +el.getAttributeNS(W, n) : d;
  const right = g(pgSz, 'w', 11906) - g(pgMar, 'left', 1440) - g(pgMar, 'right', 1440);
  for (const p of kids(body, 'p')) {
    const text = ptext(p).trim();
    if (text.startsWith('क्रमांक') && text.includes('दिनांक')) { // "दिनांक" to the right corner via right tab
      const runs = kids(p, 'r');
      for (let i = 1; i < runs.length; i++) {
        if (ptext(runs[i]).trim() !== 'दिनांक') continue;
        const ts = runs[i - 1].getElementsByTagNameNS(W, 't');
        if (ts.length) ts[ts.length - 1].textContent = ts[ts.length - 1].textContent.replace(/ +$/, '');
        runs[i - 1].appendChild(dom.createElementNS(W, 'w:tab'));
        kids(pPrOf(p), 'tabs').forEach(x => x.remove());
        const tab = dom.createElementNS(W, 'w:tab'); wa(tab, 'val', 'right'); wa(tab, 'pos', right);
        pChild(p, 'tabs').appendChild(tab); break;
      }
    }
    if (text.startsWith('सहायक लेखाधिकारी')) wa(pChild(p, 'spacing'), 'before', SETTINGS.SIGN_SPACE_PT * 20);
    const sp = kids(pPrOf(p), 'spacing')[0], line = sp && sp.getAttributeNS(W, 'line'), rule = sp && sp.getAttributeNS(W, 'lineRule');
    if (line && (!rule || rule === 'auto') && +line / 240 > SETTINGS.LINE_SPACING) wa(sp, 'line', Math.round(SETTINGS.LINE_SPACING * 240));
    if (!kids(p, 'pPr')[0].children.length) kids(p, 'pPr')[0].remove();
  }
  let els = [...body.children].filter(e => e.localName !== 'sectPr');
  while (els.length && isEmptyP(els[els.length - 1])) { if (els.length > 1 && els[els.length - 2].localName === 'tbl') break; els.pop(); }
  els.filter(e => e.localName === 'p').slice(0, -1).forEach(e => { // extra gap under each paragraph (except last)
    const sp = pChild(e, 'spacing'); wa(sp, 'after', (+sp.getAttributeNS(W, 'after') || 0) + SETTINGS.EXTRA_SPACE_PT * 20);
  });
  return { dom, body, sect, els };
}

/** items: array of placeholder-value objects (one per GP). tpl: file name inside /template. Returns a Blob. */
export async function generateDocx(items, tpl) {
  const r = await fetch(new URL('./template/' + tpl, document.baseURI).href);
  if (!r.ok) throw new Error('TEMPLATE_MISSING:' + tpl);
  const buf = await r.arrayBuffer();
  const zips = items.map(v => render(buf, v));
  const parts = zips.map((z, i) => postProcess(z, items[i]));
  const base = parts[0];
  [...base.body.children].filter(e => e.localName !== 'sectPr' && !base.els.includes(e)).forEach(e => e.remove());
  parts.slice(1).forEach(pt => {
    const first = pt.els[0];
    if (first && first.localName === 'p') pChild(first, 'pageBreakBefore');
    else if (first) { const p = pt.dom.createElementNS(W, 'w:p'), rr = pt.dom.createElementNS(W, 'w:r'), br = pt.dom.createElementNS(W, 'w:br'); wa(br, 'type', 'page'); rr.appendChild(br); p.appendChild(rr); pt.els.unshift(p); }
    pt.els.forEach(e => base.body.insertBefore(base.dom.importNode(e, true), base.sect || null));
  });
  let xml = new XMLSerializer().serializeToString(base.dom);
  if (!xml.startsWith('<?xml')) xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + xml;
  zips[0].file('word/document.xml', xml);
  return zips[0].generate({ type: 'blob', mimeType: MIME });
}
export function download(blob, name) {
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name;
  document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 5000);
}
