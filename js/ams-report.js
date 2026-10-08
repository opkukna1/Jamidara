// Browser port of amsreportmaker.py: builds the AMS table (landscape A4) as a .docx.
const MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
export const normPS = n => String(n ?? '').trim().replace(/\s+/g, ' ').toUpperCase().replace(/^(PANCHAYAT SAMITI|PS)\s+/, '').replace(/\s+(PANCHAYAT SAMITI|PS)$/, '').trim();
export function parseDate(v) {
  if (v instanceof Date) return v;
  if (typeof v === 'number' && v > 20000) return new Date(Math.round((v - 25569) * 86400000) + new Date().getTimezoneOffset() * 60000);
  const s = String(v ?? '').trim(); let m = s.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2,4})/), d;
  if (m) d = new Date(+m[3] < 100 ? 2000 + +m[3] : +m[3], +m[2] - 1, +m[1]);
  else if ((m = s.match(/^(\d{4})-(\d{2})-(\d{2})/))) d = new Date(+m[1], +m[2] - 1, +m[3]);
  return d && !isNaN(d) ? d : null;
}
const p2 = x => String(x).padStart(2, '0'), esc = s => String(s).replace(/[&<>]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]));

/** recs: dispatch records. selectedPS: names or ['ALL']. include: 1 = With PS (PS row counted), 2 = Without PS. */
export function buildAmsReport(recs, selectedPS, include = 1) {
  const s = v => String(v ?? '').trim();
  const rows = recs.map(r => { const parent = s(r['Parent Name']), isPS = parent === '';
    const m = s(r['Audit Party No']).match(/\/(\d+)\//);
    return { isPS, ps: normPS(isPS ? r['Unit Name'] : parent), status: s(r['Report Status']).toUpperCase(),
      end: parseDate(r['Planned End Date']), appr: parseDate(r['Report Approval Date']), party: m ? +m[1] : 999 }; });
  const all = selectedPS.some(x => String(x).toUpperCase() === 'ALL'), want = new Set(selectedPS.map(normPS));
  const sel = all ? rows : rows.filter(r => want.has(r.ps));
  if (!sel.length) throw new Error('NO_PS');
  const cnt = d => { const c = st => d.filter(r => r.status === st).length;
    return { ap: c('APPROVED'), dr: c('DRAFT'), se: c('SENT_TO_AUDIT_REPORT_VIEWER'), re: c('REJECTED_BY_AUDIT_REPORT_VIEWER'), on: d.filter(r => r.appr).length }; };
  const order = [...new Set(sel.map(r => r.ps))].map(n => { const a = sel.filter(r => r.ps === n), gp = a.filter(r => !r.isPS);
    return { n, party: Math.min(...(gp.length ? gp : a).map(r => r.party)) }; })
    .sort((a, b) => a.party - b.party || (a.n < b.n ? -1 : a.n > b.n ? 1 : 0));
  const out = []; let psAms = 0, psOn = 0;
  for (const o of order) {
    const a = sel.filter(r => r.ps === o.n), gp = a.filter(r => !r.isPS), psr = a.filter(r => r.isPS);
    const ends = (gp.length ? gp : a).map(r => r.end).filter(Boolean), end = ends.length ? new Date(Math.max(...ends)) : null;
    const base = include === 1 ? a : gp, c = cnt(base);
    const yAms = psr.some(r => r.status === 'APPROVED'), yOn = psr.some(r => r.appr);
    psAms += yAms; psOn += yOn;
    const pend = Math.max(base.length - c.ap - c.dr - c.se - c.re, 0);
    out.push([o.party !== 999 ? o.party : '', 'PS ' + o.n, end ? `${p2(end.getDate())}-${p2(end.getMonth() + 1)}-${String(end.getFullYear()).slice(2)}` : '',
      yAms ? 'Yes' : 'No', yOn ? 'Yes' : 'No', base.length, c.ap, c.dr, c.se, c.re, pend ? `Pending ${pend} GP` : '-', c.on]);
  }
  const tb = include === 1 ? sel : sel.filter(r => !r.isPS), ca = cnt(tb);
  const total = ['TOTAL', order.length, '', psAms, psOn, tb.length, ca.ap, ca.dr, ca.se, ca.re, Math.max(tb.length - ca.ap - ca.dr - ca.se - ca.re, 0), ca.on];

  const heads = ['Party no', 'Name Of PS', 'End date', 'PS approved\non AMS', 'PS approved on\nAuditonline', 'Total GP', 'Approved on AMS', 'Draft', 'Sent to audit\nreport viewer', 'Rejected by audit\nreport viewer', 'Pending\nReason', 'Approved on\nAuditonline'];
  const widths = [0.55, 1.15, 0.75, 0.75, 0.9, 0.55, 0.75, 0.55, 0.9, 0.9, 0.95, 0.75].map(i => Math.round(i * 1440));
  const cell = (t, w, b) => `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/></w:pPr>` +
    String(t).split('\n').map((x, i) => (i ? '<w:r><w:br/></w:r>' : '') + `<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>${b ? '<w:b/>' : ''}<w:sz w:val="14"/></w:rPr><w:t xml:space="preserve">${esc(x)}</w:t></w:r>`).join('') + '</w:p></w:tc>';
  const tr = (r, b) => `<w:tr><w:trPr><w:cantSplit/></w:trPr>${r.map((t, i) => cell(t, widths[i], b)).join('')}</w:tr>`;
  const bd = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(x => `<w:${x} w:val="single" w:sz="4" w:space="0" w:color="000000"/>`).join('');
  const doc = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
    `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="26"/></w:rPr><w:t>AMS REPORT</w:t></w:r></w:p>` +
    `<w:tbl><w:tblPr><w:jc w:val="center"/><w:tblW w:w="0" w:type="auto"/><w:tblBorders>${bd}</w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid>${widths.map(w => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>` +
    tr(heads, true) + out.map(r => tr(r, false)).join('') + tr(total, true) + `</w:tbl><w:p/>` +
    `<w:sectPr><w:pgSz w:w="16834" w:h="11906" w:orient="landscape"/><w:pgMar w:top="504" w:right="432" w:bottom="504" w:left="432" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr></w:body></w:document>`;
  const z = new PizZip();
  z.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  z.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  z.file('word/document.xml', doc);
  return { blob: z.generate({ type: 'blob', mimeType: MIME }), psCount: order.length, totalGp: tb.length };
}
