// Turns one template file + its config into a catalog entry and a payload. Used by the Node build
// and by the admin page (browser), so it only touches libraries through extract.mjs' setLibs().
import { extractDocx, extractPdf, extractXlsx, libs } from './extract.mjs';

const L = (th, en = th) => ({ th, en });
const AMOUNT_LABEL = /^(amount|เปนเงิน|เป็นเงิน|บาท|baht)$/i;
const typeOfAcro = f => {
  const fmt = f.actions && f.actions.Format && f.actions.Format.join(';');
  if (f.actions && f.actions.Calculate) return 'calc';
  if (f.type === 'check') return 'check';
  if (f.type === 'choice') return 'choice';
  if (fmt && /AFNumber_Format\(2/.test(fmt)) return 'money';
  if (fmt && /AFNumber_Format/.test(fmt)) return 'number';
  if (fmt && /AFDate_Format/.test(fmt)) return 'date';
  return f.multiline ? 'textarea' : 'text';
};

export async function buildPdf(cfg, buf) {
  const { pages, fields: widgets } = await extractPdf(buf);
  const fields = [], used = new Set(), byName = {};
  for (const w of widgets) (byName[w.name] = byName[w.name] || []).push(w);
  const widgetOut = w => ({ name: w.name, page: w.page, rect: w.rect, multiline: w.multiline, align: w.align, size: w.fontSize, exp: w.exportValue, maxLen: w.maxLen || undefined });

  // COI: pair each detail box with the Yes/No check boxes on the same row.
  for (const [name, th, en] of cfg.coi || []) {
    const d = (byName[name] || [])[0];
    if (!d) throw new Error(`${cfg.id}: COI field ${name} missing`);
    const [x, y, w, h] = d.rect;
    const boxes = widgets.filter(c => c.type === 'check' && c.page === d.page && !used.has(c.name)
      && c.rect[1] + c.rect[3] / 2 > y - 2 && c.rect[1] + c.rect[3] / 2 < y + h + 2 && c.rect[0] < x)
      .sort((a, b) => a.rect[0] - b.rect[0]);
    if (boxes.length >= 2) {
      const [yes, no] = [boxes[0], boxes[boxes.length - 1]];
      used.add(yes.name); used.add(no.name);
      fields.push({ key: `q_${name}`, label: L(`${th} — ใช่/ไม่ใช่`, `${en} — yes/no`), type: 'yesno', group: 'coi', ask: 100 + fields.length,
        pdf: [{ ...widgetOut(yes), role: 'yes' }, { ...widgetOut(no), role: 'no' }] });
    }
    used.add(name);
    fields.push({ key: name, label: L(`${th} — รายละเอียด`, `${en} — details`), type: 'textarea', group: 'coi', pdf: (byName[name]).map(widgetOut) });
  }
  const sigFields = new Set((cfg.signers || []).map(s => s.field).filter(Boolean));
  // Generic AcroForm fields, grouped by name (same name = same value).
  for (const [name, ws] of Object.entries(byName)) {
    if (used.has(name) || sigFields.has(name)) continue;
    const w0 = ws[0], o = (cfg.labels || {})[name] || {};
    let label = o.label;
    if (!label) {
      let t = w0.type === 'check' ? (w0.autoLabel.right || w0.autoLabel.left) : (w0.autoLabel.left || w0.autoLabel.right);
      t = t.replace(/\s+(amount|เปนเงิน|เป็นเงิน)$/i, '');
      if (!t || AMOUNT_LABEL.test(t)) {
        // amount column: borrow the row's description (check box text on the same line)
        const row = widgets.find(c => c !== w0 && c.page === w0.page && Math.abs((c.rect[1] + c.rect[3] / 2) - (w0.rect[1] + w0.rect[3] / 2)) < 7 && c.autoLabel.right && c.rect[0] < w0.rect[0]);
        t = row ? `${row.autoLabel.right.replace(/\s+(amount|เปนเงิน|เป็นเงิน).*$/i, '')} (${cfg.docLang === 'th' ? 'จำนวนเงิน' : 'amount'})` : name;
      }
      label = L(t, t);
    }
    const type = o.type || typeOfAcro(w0);
    const f = { key: o.key || name, label, type, pdf: ws.map(widgetOut) };
    if (w0.options) f.options = w0.options;
    if (type === 'check') f.exp = w0.exportValue || 'Yes';
    const v = Array.isArray(w0.value) ? w0.value[0] : w0.value;
    if (v && String(v).trim() && v !== 'Off') f.value = type === 'check' ? true : String(v).trim();
    for (const k of ['ask', 'tag', 'def', 'mirror', 'hide', 'kw', 'q', 'ph']) if (o[k] != null) f[k] = o[k];
    fields.push(f);
  }
  // Calculation / format scripts, run in the browser with a small Acrobat shim.
  const scripts = {};
  for (const w of widgets) if (w.actions) {
    const s = {};
    for (const k of ['Calculate', 'Format']) if (w.actions[k]) s[k.toLowerCase()] = w.actions[k].join('\n');
    if (Object.keys(s).length) scripts[w.name] = s;
  }
  for (const f of cfg.overlay || []) fields.push({ ...f, type: f.type || 'text', ov: { page: 1, rect: f.rect, size: f.size || 14 }, rect: undefined });
  const signers = (cfg.signers || []).map(s => {
    if (s.field) { const w = byName[s.field][0]; return { ...s, pdf: { page: w.page, rect: w.rect }, field: undefined }; }
    return { ...s, pdf: { page: s.page || 1, rect: s.rect }, rect: undefined };
  });
  return { pages, fields, signers, scripts, file: buf };
}

export async function buildDocx(cfg, buf) {
  const { fields: auto, labels, tokenised } = await extractDocx(buf);
  const map = cfg.map || {}, rename = {}, fields = [], signers = [];
  let lastText = null;
  for (const a of auto) {
    const o = map[a.key] || {};
    const key = o.key || a.key;
    rename[a.key] = key;
    if (a.kind === 'sig') { signers.push({ role: key.replace(/^sigline_/, ''), label: o.label || L(a.label), token: key }); continue; }
    const type = o.type || (a.kind === 'check' ? 'check' : a.kind === 'date' ? 'date' : 'text');
    const f = { key, label: o.label || L(a.label || a.line.slice(0, 40)), type };
    if (a.cont && !map[a.key] && lastText) { f.cont = lastText; f.hide = true; }
    if (a.kind === 'signame' && !o.mirror) f.label = o.label || L('ชื่อใต้ลายเซ็น', 'Name under signature');
    for (const k of ['ask', 'tag', 'def', 'mirror', 'hide', 'kw', 'q', 'ph']) if (o[k] != null) f[k] = o[k];
    if (type !== 'check' && !f.cont) lastText = key;
    fields.push(f);
  }
  // signer names: the "(......)" line right after each signature line
  for (const s of signers) {
    const i = auto.findIndex(a => rename[a.key] === s.token);
    const next = auto[i + 1];
    if (next && next.kind === 'signame') s.nameKey = rename[next.key];
  }
  const { JSZip } = libs();
  const zip = await JSZip.loadAsync(tokenised);
  let xml = await zip.file('word/document.xml').async('string');
  xml = xml.replace(/\{\{(\w+)\}\}/g, (m, k) => `{{${rename[k] || k}}}`);
  zip.file('word/document.xml', xml);
  const file = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  const lab = Object.fromEntries(Object.entries(labels).map(([k, v]) => [rename[k] || k, v]));
  return { fields, signers, labels: lab, file };
}

export async function buildXlsx(cfg, buf) {
  const x = await extractXlsx(buf);
  const fields = [], cells = cfg.cells || {};
  const sheetIdx = Object.fromEntries(x.sheets.map((s, i) => [s.name, i]));
  const seen = new Set();
  const add = (sheet, ref, merge, o, autoLabel) => {
    const id = `${sheet}!${ref}`;
    if (seen.has(id)) return; seen.add(id);
    const table = o.table || (cfg.tableSheets || []).includes(sheet);
    const key = o.key || `s${sheetIdx[sheet] + 1}_${ref}`;
    const f = { key, label: o.label || L(autoLabel ? `${autoLabel} (${ref})` : `${sheet} ${ref}`), type: o.type || 'text', xl: { sheet, ref, merge } };
    if (table) f.table = true;
    for (const k of ['ask', 'tag', 'def', 'mirror', 'hide', 'kw', 'q', 'ph']) if (o[k] != null) f[k] = o[k];
    fields.push(f);
  };
  for (const [sheet, map] of Object.entries(cells)) for (const [ref, o] of Object.entries(map)) {
    const sh = x.sheetsFull.find(s => s.name === sheet);
    if (!sh) throw new Error(`${cfg.id}: sheet ${sheet} missing`);
    add(sheet, ref, sh.merges.find(r => r.startsWith(ref + ':')) || null, o);
  }
  if (!cfg.manualCells) for (const i of x.inputs) if (i.sheet !== 'Instruction') add(i.sheet, i.ref, i.merge, {}, i.label);
  const signers = (cfg.signers || []).map(s => ({ ...s, xl: { sheet: s.sheet, range: s.range }, sheet: undefined, range: undefined }));
  return { fields, signers, sheets: x.sheets, file: buf };
}


/** cfg: {id, kind, file, cat, docLang, dateFmt, title, desc, …kind options}. Returns {entry, payload}. */
export async function compileTemplate(cfg, bytes) {
  const r = cfg.kind === 'pdf' ? await buildPdf(cfg, bytes) : cfg.kind === 'docx' ? await buildDocx(cfg, bytes) : await buildXlsx(cfg, bytes);
  const { file, labels, scripts, ...meta } = r;
  const entry = {
    id: cfg.id, kind: cfg.kind, file: cfg.file, cat: cfg.cat, docLang: cfg.docLang, dateFmt: cfg.dateFmt || (cfg.docLang === 'th' ? 'dmy-be' : 'dmy'),
    title: cfg.title, desc: cfg.desc, ...meta,
  };
  entry.fields = entry.fields.map(f => JSON.parse(JSON.stringify(f)));
  return { entry, payload: { file: file instanceof Uint8Array ? file : new Uint8Array(file), labels, scripts } };
}
