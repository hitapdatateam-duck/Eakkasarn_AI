// Field extraction for the three template kinds. Pure functions over file bytes;
// build-templates.mjs merges the results with templates.config.mjs and writes the catalog.
import JSZip from 'jszip';
import * as pdfjs from 'pdfjs-dist/legacy/build/pdf.js';

const DOTS = '.…';
const isDot = c => DOTS.includes(c);
const clean = s => s.replace(/[.…_]+/g, ' ').replace(/\s+/g, ' ').replace(/^[\s:(（,]+|[\s:)）,]+$/g, '').trim();

/* =============================== DOCX =============================== */
const RUN = /<w:r[ >][\s\S]*?<\/w:r>/g;
const T = /(<w:t(?: [^>]*)?>)([^<]*)(<\/w:t>)/g;
const SYM = /<w:sym [^>]*\/>/;
const SYM_CHAR = '';
const runText = r => { let s = ''; for (const m of r.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>|<w:sym [^>]*\/>|<w:tab\/>/g)) s += m[1] !== undefined ? m[1] : m[0].startsWith('<w:sym') ? SYM_CHAR : '\t'; return s; };
const xmlUnesc = s => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');

/** Replace characters [s,e) of the paragraph's text with `token` (placed in the first run touched). */
function spliceRuns(p, s, e, token) {
  const runs = [...p.matchAll(RUN)];
  let pos = 0, placed = false, out = '', last = 0;
  for (const m of runs) {
    const r = m[0];
    const t = runText(r), a = pos, b = pos + t.length;
    pos = b;
    let nr = r;
    if (b > s && a < e && t.length) {
      // rebuild this run's text pieces keeping only chars outside [s,e)
      let idx = a;
      nr = r.replace(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>|<w:sym [^>]*\/>|<w:tab\/>/g, (all, txt) => {
        const len = txt !== undefined ? xmlUnesc(txt).length : 1;
        const from = idx, to = idx + len; idx = to;
        if (to <= s || from >= e) return all;
        let keep = '';
        if (txt !== undefined) {
          const raw = xmlUnesc(txt);
          keep = raw.slice(0, Math.max(0, s - from)) + raw.slice(Math.max(0, e - from));
          const before = raw.slice(0, Math.max(0, s - from)), after = raw.slice(Math.max(0, e - from));
          const ins = !placed && from <= s ? token : '';
          if (ins) placed = true;
          const text = before + ins + after;
          return text ? `<w:t xml:space="preserve">${text.replace(/&/g, '&amp;').replace(/</g, '&lt;')}</w:t>` : '';
        }
        if (!placed) { placed = true; return `<w:t xml:space="preserve">${token}</w:t>`; }
        return '';
      });
      if (!placed && a >= s) { nr = nr.replace(/<\/w:r>$/, `<w:t xml:space="preserve">${token}</w:t></w:r>`); placed = true; }
    }
    out += p.slice(last, m.index) + nr; last = m.index + r.length;
  }
  return out + p.slice(last);
}

export async function extractDocx(buf) {
  const zip = await JSZip.loadAsync(buf);
  let xml = await zip.file('word/document.xml').async('string');
  const fields = [], labels = {};
  let n = 0, prevLabel = '', lastSig = null;
  xml = xml.replace(/<w:p[ >][\s\S]*?<\/w:p>/g, p => {
    const runs = [...p.matchAll(RUN)].map(m => m[0]);
    const text = runs.map(runText).join('');
    if (!/[.…]{3}|/.test(text)) { if (text.trim()) prevLabel = ''; return p; }
    // spans: date (../../..), dotted leaders, checkbox symbols
    const spans = [];
    const re = /([.…][.…\s]*\/[.…\s]*[.…][.…\s]*\/[.…\s]*[.…]+)|([.…][.…]*[.…])|()/g;
    for (const m of text.matchAll(re)) {
      if (m[2] && m[2].length < 3) continue;
      spans.push({ s: m.index, e: m.index + m[0].length, kind: m[1] ? 'date' : m[3] ? 'check' : 'text', raw: m[0] });
    }
    if (!spans.length) return p;
    const isSigPara = /^\s*\(ลงชื่อ\)/.test(text);
    const isNamePara = /^\s*\(\s*[.…]/.test(text) && /\)\s*$/.test(text.trim());
    const symXml = [...p.matchAll(/<w:sym [^>]*\/>/g)].map(m => m[0]);
    let symIdx = 0, cursor = 0;
    const meta = spans.map((sp, i) => {
      const before = clean(text.slice(cursor, sp.s));
      const after = clean(text.slice(sp.e, i + 1 < spans.length ? spans[i + 1].s : text.length));
      cursor = sp.e;
      n++;
      let key, kind = sp.kind, label;
      if (kind === 'check') { key = `chk${n}`; label = after.split(/\s{2,}/)[0]; labels[key] = symXml[symIdx++]; }
      else if (isSigPara && kind === 'text') { key = `sigline_s${n}`; kind = 'sig'; label = after || 'ลงชื่อ'; labels[key] = sp.raw; lastSig = key; }
      else if (isNamePara) { key = `f${n}`; kind = 'signame'; label = 'ชื่อผู้ลงนาม'; labels[key] = sp.raw; }
      else {
        key = `f${n}`; labels[key] = sp.raw;
        label = before || (text.trim() === sp.raw.trim() && prevLabel ? `${prevLabel} (ต่อ)` : '') || after;
        if (kind === 'date' && !before && lastSig) label = 'วันที่ลงนาม';
      }
      if (kind === 'text' || kind === 'date') prevLabel = label.replace(/ \(ต่อ\)$/, '');
      return { key, kind, label, line: clean(text).slice(0, 140), cont: kind === 'text' && !before && text.trim() === sp.raw.trim() };
    });
    fields.push(...meta);
    // splice from the end so earlier offsets stay valid
    let out = p;
    for (let i = spans.length - 1; i >= 0; i--) {
      const sp = spans[i], m = meta[i];
      if (sp.kind === 'check') {
        let k = -1;
        out = out.replace(/<w:sym [^>]*\/>/g, all => (++k === symIdxOf(meta, i) ? `<w:t xml:space="preserve">{{${m.key}}}</w:t>` : all));
      } else out = spliceRuns(out, sp.s, sp.e, `{{${m.key}}}`);
    }
    return out;
  });
  zip.file('word/document.xml', xml);
  const tokenised = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  return { fields, labels, tokenised };
}
const symIdxOf = (meta, i) => meta.slice(0, i).filter(m => m.kind === 'check').length;

/* =============================== PDF =============================== */
export async function extractPdf(buf) {
  const doc = await pdfjs.getDocument({ data: new Uint8Array(buf), isEvalSupported: false, disableFontFace: true, verbosity: 0 }).promise;
  const pages = [], fields = [];
  const objs = (await doc.getFieldObjects()) || {};
  for (let p = 1; p <= doc.numPages; p++) {
    const page = await doc.getPage(p);
    const [, , w, h] = page.view;
    pages.push({ w, h });
    const items = (await page.getTextContent()).items.filter(i => i.str.trim())
      .map(i => ({ s: i.str, x: i.transform[4], y: i.transform[5], w: i.width, h: i.height || 10 }));
    for (const a of await page.getAnnotations()) {
      if (a.subtype !== 'Widget' || a.pushButton || !a.fieldName) continue;
      const [x1, y1, x2, y2] = a.rect;
      const cy = (y1 + y2) / 2, band = Math.max(6, (y2 - y1) / 2 + 3);
      const sameLine = items.filter(i => Math.abs(i.y + i.h * 0.3 - cy) < band).sort((a, b) => a.x - b.x);
      // One text item can span several fields ("Name…… Organization……"), so estimate per-character x.
      const cut = (i, from, to) => { const n = i.s.length, cw = i.w / Math.max(1, n); const a = Math.max(0, Math.round((from - i.x) / cw)), b = Math.min(n, Math.round((to - i.x) / cw)); return b > a ? i.s.slice(a, b) : ''; };
      const leftStr = sameLine.map(i => cut(i, -1e9, x1 + 2)).join(' ');
      const rightStr = sameLine.map(i => cut(i, x2 - 2, 1e9)).join(' ');
      const leftText = clean(leftStr.split(/[.…_]{3,}/).filter(s => clean(s)).pop() || '');
      const rightText = clean(rightStr.split(/[.…_]{3,}/).filter(s => clean(s))[0] || '');
      const fo = (objs[a.fieldName] || []).find(o => o.id === a.id) || (objs[a.fieldName] || [])[0] || {};
      const type = a.checkBox ? 'check' : a.radioButton ? 'radio' : a.combo || a.listBox ? 'choice' : 'text';
      fields.push({
        name: a.fieldName, id: a.id, page: p, rect: [x1, y1, x2 - x1, y2 - y1].map(v => Math.round(v * 100) / 100),
        type, multiline: !!a.multiLine, readOnly: !!a.readOnly, maxLen: a.maxLen || 0,
        options: a.options ? a.options.map(o => ({ v: o.exportValue, t: o.displayValue })) : undefined,
        exportValue: a.exportValue || (a.checkBox ? 'Yes' : undefined),
        value: a.fieldValue ?? '', defaultValue: fo.defaultValue ?? undefined,
        align: a.textAlignment ?? 0, fontSize: a.textFontSize || 0,
        actions: fo.actions || a.actions || undefined,
        autoLabel: { left: leftText, right: rightText },
      });
    }
  }
  return { pages, fields };
}

/* =============================== XLSX =============================== */
export async function extractXlsx(buf, { inputFill = 'FFCCFFFF' } = {}) {
  const zip = await JSZip.loadAsync(buf);
  const read = p => zip.file(p).async('string');
  const wb = await read('xl/workbook.xml');
  const rels = await read('xl/_rels/workbook.xml.rels');
  const relMap = Object.fromEntries([...rels.matchAll(/<Relationship [^>]*>/g)].map(m => [/Id="([^"]+)"/.exec(m[0])[1], /Target="([^"]+)"/.exec(m[0])[1].replace(/^\/?xl\//, '')]));
  const sst = zip.file('xl/sharedStrings.xml') ? [...(await read('xl/sharedStrings.xml')).matchAll(/<si>([\s\S]*?)<\/si>/g)].map(m => xmlUnesc([...m[1].matchAll(/<t[^>]*>([^<]*)<\/t>/g)].map(x => x[1]).join(''))) : [];
  const styles = await read('xl/styles.xml');
  const fills = [...(/<fills[^>]*>([\s\S]*?)<\/fills>/.exec(styles)[1]).matchAll(/<fill>[\s\S]*?<\/fill>|<fill\/>/g)].map(m => (/fgColor rgb="([^"]+)"/.exec(m[0]) || [])[1] || '');
  const xfs = [...(/<cellXfs[^>]*>([\s\S]*?)<\/cellXfs>/.exec(styles)[1]).matchAll(/<xf [^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g)].map(m => ({ fill: +((/fillId="(\d+)"/.exec(m[0]) || [])[1] || 0), locked: !/locked="0"/.test(m[0]), numFmt: +((/numFmtId="(\d+)"/.exec(m[0]) || [])[1] || 0) }));
  const sheets = [];
  for (const m of wb.matchAll(/<sheet [^>]*>/g)) {
    const name = xmlUnesc(/name="([^"]+)"/.exec(m[0])[1]), rid = /r:id="([^"]+)"/.exec(m[0])[1], hidden = /state="hidden"/.test(m[0]);
    const path = 'xl/' + relMap[rid];
    const x = await read(path);
    const merges = [...x.matchAll(/<mergeCell ref="([^"]+)"/g)].map(mm => mm[1]);
    const cells = {};
    for (const c of x.matchAll(/<c r="([A-Z]+\d+)"([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
      const s = +((/ s="(\d+)"/.exec(c[2]) || [])[1] || 0), t = (/ t="(\w+)"/.exec(c[2]) || [])[1];
      const v = (/<v>([\s\S]*?)<\/v>/.exec(c[3] || '') || [])[1];
      const f = /<f/.test(c[3] || '');
      cells[c[1]] = { s, t, v: t === 's' && v != null ? sst[+v] : v != null ? xmlUnesc(v) : '', f, fill: fills[xfs[s]?.fill] || '', locked: xfs[s]?.locked !== false };
    }
    sheets.push({ name, path, hidden, merges, cells });
  }
  // Input cells: the template marks them with a light-blue fill. Collapse merged ranges to their anchor.
  const inputs = [];
  for (const sh of sheets) {
    const covered = new Set();
    const anchorOf = {};
    for (const r of sh.merges) { const [a, b] = r.split(':'); for (const ref of rangeRefs(a, b)) { anchorOf[ref] = a; if (ref !== a) covered.add(ref); } }
    for (const [ref, c] of Object.entries(sh.cells)) {
      if (c.fill !== inputFill || c.f || covered.has(ref)) continue;
      inputs.push({ sheet: sh.name, ref, merge: sh.merges.find(r => r.startsWith(ref + ':')) || null, label: rowLabel(sh, ref, covered) });
    }
  }
  return { sheets: sheets.map(s => ({ name: s.name, hidden: s.hidden })), inputs, sheetsFull: sheets };
}
export const colNum = c => [...c].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
export const colName = n => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
export const splitRef = r => { const m = /^([A-Z]+)(\d+)$/.exec(r); return [colNum(m[1]), +m[2]]; };
export function rangeRefs(a, b) {
  const [c1, r1] = splitRef(a), [c2, r2] = splitRef(b || a), out = [];
  for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) out.push(colName(c) + r);
  return out;
}
function rowLabel(sh, ref, covered) {
  const [c, r] = splitRef(ref);
  for (let cc = c - 1; cc >= 1; cc--) {
    const cell = sh.cells[colName(cc) + r];
    if (cell && cell.v && !cell.f && cell.fill !== 'FFCCFFFF') return clean(String(cell.v));
  }
  return '';
}
