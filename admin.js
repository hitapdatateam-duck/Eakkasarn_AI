/* Admin back-office: list / hide / delete templates, upload new ones, manage admins.
   New files are analysed in the browser with the same compiler the build uses (build/compile.mjs),
   then stored in Supabase (Storage bucket "templates" + table public.templates). Writes are allowed
   for admins only by RLS (public.is_admin()). */
import { setLibs } from './build/extract.mjs';
import { compileTemplate } from './build/compile.mjs';

const HD = window.HD, U = HD.util;
const B = () => HD.appBridge;
const tr = (th, en) => B().tr(th, en);
const $ = (s, el = document) => el.querySelector(s);
const esc = U.htmlEsc;
const KIND = { docx: 'Word', pdf: 'PDF', xlsx: 'Excel' };
const MIME = { docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', pdf: 'application/pdf', xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' };
const TAGS = ['', 'name', 'position', 'unit', 'employee_code', 'phone', 'email', 'id_card', 'address', 'organization', 'project_code', 'activity_code'];
const TAG_LABEL = { '': '—', name: 'ชื่อ', position: 'ตำแหน่ง', unit: 'ฝ่าย', employee_code: 'รหัสพนักงาน', phone: 'โทรศัพท์', email: 'อีเมล', id_card: 'เลขบัตร', address: 'ที่อยู่', organization: 'หน่วยงาน', project_code: 'รหัสโครงการ', activity_code: 'รหัสกิจกรรม' };
let tab = 'list', draft = null, busy = false;

const db = () => HD.cloud.client();
async function ensureLibs() {
  if (!window.pdfjsLib) await U.loadScript(U.LIBS.pdfjs);
  window.pdfjsLib.GlobalWorkerOptions.workerSrc = U.LIBS.pdfjsWorker;
  setLibs({ JSZip: window.JSZip, pdfjs: window.pdfjsLib });
}
const b64 = u8 => { let s = ''; for (let i = 0; i < u8.length; i += 32768) s += String.fromCharCode.apply(null, u8.subarray(i, i + 32768)); return btoa(s); };

/* ======================= shell ======================= */
function render() {
  const root = $('#adminRoot');
  if (!HD.cloud || !HD.cloud.client()) { root.innerHTML = `<div class="admin-empty">${tr('กำลังเชื่อมต่อ…', 'Connecting…')}</div>`; return; }
  if (!HD.cloud.user()) { root.innerHTML = `<div class="admin-empty">${tr('เข้าสู่ระบบด้วยบัญชีผู้ดูแลก่อนครับ', 'Sign in with an admin account first.')}</div>`; return; }
  if (!HD.cloud.isAdmin()) { root.innerHTML = `<div class="admin-empty">${tr('บัญชีนี้ไม่มีสิทธิ์ผู้ดูแลระบบ', 'This account is not an admin.')}</div>`; return; }
  root.innerHTML = `
    <h1 class="page-title">${tr('ผู้ดูแลระบบ', 'Admin')}</h1>
    <p class="page-sub">${tr('เพิ่ม ซ่อน หรือลบแม่แบบเอกสาร และจัดการผู้ดูแล', 'Add, hide or delete templates and manage admins')}</p>
    <div class="seg" role="tablist">
      <button type="button" data-tab="list" class="${tab === 'list' ? 'active' : ''}">${tr('แม่แบบทั้งหมด', 'Templates')}</button>
      <button type="button" data-tab="add" class="${tab === 'add' ? 'active' : ''}">${tr('เพิ่มแม่แบบ', 'Add template')}</button>
      <button type="button" data-tab="admins" class="${tab === 'admins' ? 'active' : ''}">${tr('ผู้ดูแล', 'Admins')}</button>
    </div>
    <div id="adminBody"></div>`;
  root.querySelectorAll('[data-tab]').forEach(b => b.addEventListener('click', () => { tab = b.dataset.tab; render(); }));
  ({ list: renderList, add: renderAdd, admins: renderAdmins })[tab]();
}

/* ======================= templates list ======================= */
function renderList() {
  const L = document.documentElement.dataset.lang === 'en' ? 'en' : 'th';
  const builtin = new Set(B().builtinIds());
  const cats = Object.fromEntries(B().categories().map(c => [c.id, c[L]]));
  const rows = B().templates().map(t => ({ t, isBuiltin: builtin.has(t.id) }));
  $('#adminBody').innerHTML = `
    <div class="admin-table">
      <div class="at-head"><span>${tr('แม่แบบ', 'Template')}</span><span>${tr('ประเภท', 'Type')}</span><span>${tr('ที่มา', 'Source')}</span><span>${tr('สถานะ', 'Status')}</span><span></span></div>
      ${rows.map(({ t, isBuiltin }) => `
        <div class="at-row${t.hidden ? ' is-hidden' : ''}">
          <span class="at-title"><b>${esc(t.title[L] || t.title.th)}</b><small>${esc(cats[t.cat] || t.cat || '')}</small></span>
          <span><span class="badge">${KIND[t.kind]}</span></span>
          <span class="muted">${isBuiltin ? tr('ในระบบ', 'Built-in') : tr('อัปโหลด', 'Uploaded')}</span>
          <span class="muted">${t.hidden ? tr('ซ่อนอยู่', 'Hidden') : tr('แสดง', 'Visible')}</span>
          <span class="at-actions">
            <button type="button" class="btn btn-light btn-sm" data-toggle="${esc(t.id)}">${t.hidden ? tr('แสดง', 'Show') : tr('ซ่อน', 'Hide')}</button>
            ${isBuiltin ? '' : `<button type="button" class="btn btn-danger btn-sm" data-del="${esc(t.id)}">${tr('ลบ', 'Delete')}</button>`}
          </span>
        </div>`).join('')}
    </div>
    <p class="help">${tr('แม่แบบในระบบลบถาวรไม่ได้ (อยู่ในโค้ดของเว็บ) แต่ซ่อนจากผู้ใช้ได้ · แม่แบบที่อัปโหลดลบได้ทั้งไฟล์และข้อมูล', 'Built-in templates live in the site code, so they can be hidden but not deleted · uploaded templates can be deleted with their files')}</p>`;
  $('#adminBody').querySelectorAll('[data-toggle]').forEach(b => b.addEventListener('click', () => toggleHidden(b.dataset.toggle)));
  $('#adminBody').querySelectorAll('[data-del]').forEach(b => b.addEventListener('click', () => removeTemplate(b.dataset.del)));
}
async function toggleHidden(id) {
  const t = B().templates().find(x => x.id === id);
  const isBuiltin = B().builtinIds().includes(id);
  const { error } = isBuiltin
    ? await db().from('templates').upsert({ id, kind: t.kind, builtin: true, hidden: !t.hidden, meta: {} }, { onConflict: 'id' })
    : await db().from('templates').update({ hidden: !t.hidden, updated_at: new Date().toISOString() }).eq('id', id);
  if (error) return B().toast(tr('บันทึกไม่สำเร็จ: ', 'Failed: ') + error.message);
  await HD.cloud.reloadTemplates();
  B().toast(t.hidden ? tr('แสดงแม่แบบแล้ว', 'Template shown') : tr('ซ่อนแม่แบบแล้ว', 'Template hidden'));
  render();
}
async function removeTemplate(id) {
  const t = B().templates().find(x => x.id === id);
  if (!confirm(tr(`ลบแม่แบบ “${t.title.th}” ถาวร? เอกสารที่ผู้ใช้สร้างจากแม่แบบนี้จะเปิดไม่ได้อีก`, `Delete “${t.title.en}” permanently? Documents made from it will no longer open.`))) return;
  const paths = [t.dataPath, t.originalPath].filter(Boolean);
  if (paths.length) {
    const { error } = await db().storage.from('templates').remove(paths);
    if (error) return B().toast(tr('ลบไฟล์ไม่สำเร็จ: ', 'Failed to delete files: ') + error.message);
  }
  const { error } = await db().from('templates').delete().eq('id', id);
  if (error) return B().toast(tr('ลบไม่สำเร็จ: ', 'Failed: ') + error.message);
  await HD.cloud.reloadTemplates();
  B().toast(tr('ลบแม่แบบแล้ว', 'Template deleted'));
  render();
}

/* ======================= add template ======================= */
function renderAdd() {
  const L = document.documentElement.dataset.lang === 'en' ? 'en' : 'th';
  $('#adminBody').innerHTML = `
    <form class="admin-card" id="addForm">
      <label class="field"><span>${tr('ไฟล์แบบฟอร์ม (.docx, .pdf, .xlsx)', 'Form file (.docx, .pdf, .xlsx)')}<span class="req">*</span></span>
        <input type="file" id="afFile" accept=".docx,.pdf,.xlsx" required></label>
      <div class="row2">
        <label class="field"><span>${tr('ชื่อแม่แบบ (ไทย)', 'Title (Thai)')}<span class="req">*</span></span><input id="afTitleTh" required></label>
        <label class="field"><span>${tr('ชื่อแม่แบบ (อังกฤษ)', 'Title (English)')}</span><input id="afTitleEn"></label>
      </div>
      <label class="field"><span>${tr('คำอธิบาย', 'Description')}</span><input id="afDesc"></label>
      <div class="row2">
        <label class="field"><span>${tr('หมวด', 'Category')}</span><select id="afCat">${B().categories().map(c => `<option value="${c.id}">${esc(c[L])}</option>`).join('')}</select></label>
        <label class="field"><span>${tr('ภาษาของเอกสาร', 'Document language')}</span><select id="afLang"><option value="th">ไทย</option><option value="en">English</option></select></label>
      </div>
      <div class="admin-actions"><button type="submit" class="btn btn-dark" id="afAnalyse">${tr('วิเคราะห์ไฟล์', 'Analyse file')}</button></div>
      <p class="help">${tr('Word: เส้นจุด …… และช่อง ☐ จะกลายเป็นช่องกรอก · PDF: ใช้ช่องกรอกในไฟล์ หรือคลิกบนหน้าเพื่อวางช่องเอง · Excel: ช่องพื้นสีฟ้าอ่อน หรือระบุเซลล์เอง', 'Word: dotted lines and ☐ boxes become fields · PDF: uses its form fields, or click the page to place fields · Excel: light-blue cells, or list cells yourself')}</p>
    </form>
    <div id="draft"></div>`;
  $('#addForm').addEventListener('submit', e => { e.preventDefault(); analyse(); });
  if (draft) renderDraft();
}
async function analyse() {
  if (busy) return;
  const f = $('#afFile').files[0];
  if (!f) return;
  const ext = f.name.split('.').pop().toLowerCase();
  if (!MIME[ext]) return B().toast(tr('รองรับเฉพาะ .docx .pdf .xlsx', 'Only .docx .pdf .xlsx are supported'));
  busy = true; $('#afAnalyse').disabled = true; $('#afAnalyse').textContent = tr('กำลังวิเคราะห์…', 'Analysing…');
  try {
    await ensureLibs();
    const bytes = new Uint8Array(await f.arrayBuffer());
    const titleTh = $('#afTitleTh').value.trim() || f.name.replace(/\.[^.]+$/, '');
    const cfg = { id: 'u-' + Date.now().toString(36), kind: ext, file: f.name, cat: $('#afCat').value, docLang: $('#afLang').value,
      title: { th: titleTh, en: $('#afTitleEn').value.trim() || titleTh }, desc: { th: $('#afDesc').value.trim(), en: $('#afDesc').value.trim() } };
    const { entry, payload } = await compileTemplate(cfg, bytes);
    if (ext === 'xlsx') {
      // rows with 4+ input cells are tables (filled in place on the sheet, not listed as questions)
      const perRow = {};
      for (const f of entry.fields) if (f.xl) { const k = f.xl.sheet + '!' + f.xl.ref.replace(/[A-Z]+/, ''); perRow[k] = (perRow[k] || 0) + 1; }
      for (const f of entry.fields) if (f.xl && perRow[f.xl.sheet + '!' + f.xl.ref.replace(/[A-Z]+/, '')] >= 4) f.table = true;
    }
    draft = { cfg, entry, payload, bytes, ext, fileName: f.name };
    renderDraft();
  } catch (err) {
    console.error(err);
    B().toast(tr('วิเคราะห์ไฟล์ไม่สำเร็จ: ', 'Could not analyse the file: ') + err.message);
  } finally { busy = false; const b = $('#afAnalyse'); if (b) { b.disabled = false; b.textContent = tr('วิเคราะห์ไฟล์', 'Analyse file'); } }
}
const editable = f => !f.hide && !f.table && !f.cont && f.type !== 'calc';
function renderDraft() {
  const { entry, ext } = draft;
  const fields = entry.fields.filter(editable);
  const tableCount = entry.fields.filter(f => f.table).length;
  $('#draft').innerHTML = `
    <div class="admin-card">
      <h3>${tr('ช่องที่พบ', 'Detected fields')} · ${fields.length}${tableCount ? tr(` (+ ตาราง ${tableCount} ช่อง)`, ` (+ ${tableCount} table cells)`) : ''}</h3>
      <p class="help">${tr('แก้ชื่อช่องให้อ่านง่าย ติ๊ก “ถาม” สำหรับช่องที่ให้ผู้ช่วยในแชทถาม (ตามลำดับ) และเลือก “ข้อมูลส่วนตัว” เพื่อเติมให้อัตโนมัติจากเอกสารก่อนหน้า', 'Rename fields, tick “Ask” for fields the assistant should ask (in order), and pick a profile item to prefill it')}</p>
      ${fields.length ? `<div class="df-table">
        <div class="df-head"><span>${tr('ชื่อช่อง', 'Label')}</span><span>${tr('ชนิด', 'Type')}</span><span>${tr('ข้อมูลส่วนตัว', 'Profile')}</span><span>${tr('ถาม', 'Ask')}</span><span></span></div>
        ${fields.map(f => `<div class="df-row" data-key="${esc(f.key)}">
          <input class="df-label" value="${esc(f.label.th)}">
          ${f.type === 'check' || f.type === 'yesno' || f.type === 'choice' ? `<span class="muted">${f.type}</span>` : `<select class="df-type">${['text', 'textarea', 'date', 'money', 'number'].map(t => `<option ${t === f.type ? 'selected' : ''}>${t}</option>`).join('')}</select>`}
          <select class="df-tag">${TAGS.map(t => `<option value="${t}" ${t === (f.tag || '') ? 'selected' : ''}>${TAG_LABEL[t]}</option>`).join('')}</select>
          <input type="checkbox" class="df-ask" ${f.ask ? 'checked' : ''} ${f.type === 'check' ? 'disabled' : ''}>
          <button type="button" class="icon-btn df-del" title="${tr('ลบช่องนี้', 'Remove')}">✕</button>
        </div>`).join('')}
      </div>` : `<p class="admin-warn">${tr('ยังไม่พบช่องกรอก', 'No fields found yet')}${ext === 'pdf' ? tr(' — คลิกบนหน้าเอกสารด้านล่างเพื่อวางช่อง', ' — click the page below to place fields') : ext === 'xlsx' ? tr(' — ระบุเซลล์ด้านล่าง', ' — list the cells below') : tr(' — ไฟล์ Word ต้องมีเส้นจุด …… หรือช่อง ☐', ' — Word files need dotted lines or ☐ boxes')}</p>`}
      ${ext === 'xlsx' ? `<label class="field"><span>${tr('เพิ่มเซลล์เอง (บรรทัดละ 1 เซลล์: ชื่อชีท!เซลล์ = ชื่อช่อง)', 'Add cells (one per line: Sheet!Cell = Label)')}</span>
        <textarea id="xlCells" rows="3" placeholder="${esc((entry.sheets && entry.sheets[0] && entry.sheets[0].name) || 'Sheet1')}!B5 = ชื่อ-นามสกุล"></textarea>
        <span><button type="button" class="btn btn-light btn-sm" id="xlAdd">${tr('เพิ่มเซลล์', 'Add cells')}</button></span></label>` : ''}
      ${entry.signers && entry.signers.length ? `<h3>${tr('ช่องลงนาม', 'Signatures')} · ${entry.signers.length}</h3>
        <div class="df-table">${entry.signers.map((s, i) => `<div class="df-row sig" data-sig="${i}"><input class="sig-label" value="${esc(s.label.th)}"><span class="muted">${tr('ลายเซ็น', 'signature')}</span><span></span><span></span><button type="button" class="icon-btn sig-del">✕</button></div>`).join('')}</div>` : ''}
    </div>
    ${ext === 'pdf' ? `<div class="admin-card"><h3>${tr('วางช่องบนหน้า PDF', 'Place fields on the PDF')}</h3>
      <div class="seg small" id="placeMode"><button type="button" data-mode="text" class="active">+ ${tr('ช่องข้อความ', 'Text field')}</button><button type="button" data-mode="sig">+ ${tr('ช่องลงนาม', 'Signature')}</button><button type="button" data-mode="none">${tr('ดูอย่างเดียว', 'View only')}</button></div>
      <p class="help">${tr('คลิกตรงจุดเริ่มของเส้นที่ต้องการกรอก · กรอบเทา = ช่องกรอกเดิมในไฟล์ · เหลือง = ช่องที่เพิ่ม · เขียว = ช่องลงนาม', 'Click where a line starts · grey = existing form fields · yellow = added · green = signatures')}</p>
      <div id="pdfPages" class="pdf-pages"></div></div>` : ''}
    <div class="admin-actions sticky"><button type="button" class="btn btn-light" id="draftCancel">${tr('ยกเลิก', 'Cancel')}</button><button type="button" class="btn btn-dark" id="draftSave">${tr('บันทึกและเผยแพร่', 'Save & publish')}</button></div>`;
  const root = $('#draft');
  root.querySelectorAll('.df-del').forEach(b => b.addEventListener('click', () => {
    const key = b.closest('.df-row').dataset.key;
    syncDraft();
    draft.entry.fields = draft.entry.fields.filter(f => f.key !== key);
    renderDraft();
  }));
  root.querySelectorAll('.sig-del').forEach(b => b.addEventListener('click', () => { syncDraft(); draft.entry.signers.splice(+b.closest('.df-row').dataset.sig, 1); renderDraft(); }));
  root.querySelectorAll('input, select').forEach(el => el.addEventListener('change', syncDraft));
  if ($('#xlAdd')) $('#xlAdd').addEventListener('click', addXlCells);
  $('#draftCancel').addEventListener('click', () => { draft = null; renderAdd(); });
  $('#draftSave').addEventListener('click', saveDraft);
  if (ext === 'pdf') renderPdfPlacement();
}
/** Copy the editor's inputs back into draft.entry (labels, type, profile tag, ask order). */
function syncDraft() {
  if (!draft) return;
  let order = 1;
  document.querySelectorAll('#draft .df-row[data-key]').forEach(row => {
    const f = draft.entry.fields.find(x => x.key === row.dataset.key);
    if (!f) return;
    const label = row.querySelector('.df-label').value.trim() || f.label.th;
    f.label = { th: label, en: f.label.en && f.label.en !== f.label.th ? f.label.en : label };
    const type = row.querySelector('.df-type');
    if (type) f.type = type.value;
    const tag = row.querySelector('.df-tag').value;
    if (tag) f.tag = tag; else delete f.tag;
    if (row.querySelector('.df-ask').checked) f.ask = order++; else delete f.ask;
  });
  document.querySelectorAll('#draft .df-row.sig').forEach(row => {
    const s = draft.entry.signers[+row.dataset.sig];
    if (s) { const v = row.querySelector('.sig-label').value.trim() || s.label.th; s.label = { th: v, en: v }; }
  });
}
function addXlCells() {
  syncDraft();
  const lines = $('#xlCells').value.split('\n').map(l => l.trim()).filter(Boolean);
  let n = draft.entry.fields.length;
  for (const line of lines) {
    const m = /^(.+?)!\$?([A-Z]{1,3})\$?(\d+)\s*(?:=\s*(.+))?$/.exec(line);
    if (!m) { B().toast(tr(`อ่านบรรทัดนี้ไม่ได้: ${line}`, `Could not read: ${line}`)); continue; }
    const [, sheet, col, row, label] = m;
    if (!(draft.entry.sheets || []).some(s => s.name === sheet)) { B().toast(tr(`ไม่พบชีท ${sheet}`, `No sheet ${sheet}`)); continue; }
    n++;
    draft.entry.fields.push({ key: `c${n}`, label: { th: label || `${sheet} ${col}${row}`, en: label || `${sheet} ${col}${row}` }, type: 'text', xl: { sheet, ref: col + row, merge: null } });
  }
  renderDraft();
}

/* ---------- click-to-place on PDF pages ---------- */
let placeMode = 'text';
async function renderPdfPlacement() {
  const box = $('#pdfPages');
  document.querySelectorAll('#placeMode [data-mode]').forEach(b => b.addEventListener('click', () => {
    placeMode = b.dataset.mode;
    document.querySelectorAll('#placeMode [data-mode]').forEach(x => x.classList.toggle('active', x === b));
  }));
  const pdf = await window.pdfjsLib.getDocument({ data: draft.bytes.slice(), isEvalSupported: false }).promise;
  const scale = 1.2;
  for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);
    const vp = page.getViewport({ scale }), H = page.getViewport({ scale: 1 }).height;
    const wrap = document.createElement('div');
    wrap.className = 'place-page';
    wrap.style.width = vp.width + 'px'; wrap.style.height = vp.height + 'px';
    const c = document.createElement('canvas');
    c.width = vp.width * 2; c.height = vp.height * 2; c.style.width = vp.width + 'px'; c.style.height = vp.height + 'px';
    await page.render({ canvasContext: c.getContext('2d'), viewport: page.getViewport({ scale: scale * 2 }), intent: 'print', annotationMode: window.pdfjsLib.AnnotationMode.DISABLE }).promise;
    wrap.append(c);
    const box2 = (rect, cls, title) => {
      const d = document.createElement('div');
      d.className = 'place-box ' + cls; d.title = title || '';
      Object.assign(d.style, { left: rect[0] * scale + 'px', top: (H - rect[1] - rect[3]) * scale + 'px', width: rect[2] * scale + 'px', height: rect[3] * scale + 'px' });
      wrap.append(d);
    };
    for (const f of draft.entry.fields) {
      for (const w of f.pdf || []) if (w.page === p) box2(w.rect, 'acro', f.label.th);
      if (f.ov && f.ov.page === p) box2(f.ov.rect, 'ov', f.label.th);
    }
    for (const s of draft.entry.signers || []) if (s.pdf && s.pdf.page === p) box2(s.pdf.rect, 'sig', s.label.th);
    wrap.addEventListener('click', e => {
      if (placeMode === 'none') return;
      const r = wrap.getBoundingClientRect();
      const x = (e.clientX - r.left) / scale, y = H - (e.clientY - r.top) / scale;
      syncDraft();
      if (placeMode === 'sig') {
        const n = (draft.entry.signers || []).length + 1;
        (draft.entry.signers = draft.entry.signers || []).push({ role: `sig${n}`, label: { th: `ผู้ลงนาม ${n}`, en: `Signer ${n}` }, pdf: { page: p, rect: [x, y - 6, 150, 28] } });
      } else {
        const n = draft.entry.fields.filter(f => f.ov).length + 1;
        draft.entry.fields.push({ key: `ov${n}`, label: { th: `ช่องที่ ${n}`, en: `Field ${n}` }, type: 'text', ov: { page: p, rect: [x, y - 6, 180, 17], size: 14 } });
      }
      renderDraft();
    });
    box.append(wrap);
  }
}

async function saveDraft() {
  if (busy) return;
  syncDraft();
  const { cfg, entry, payload, bytes, ext, fileName } = draft;
  if (!entry.fields.length) return B().toast(tr('ต้องมีอย่างน้อย 1 ช่องกรอก', 'Add at least one field'));
  busy = true; $('#draftSave').disabled = true; $('#draftSave').textContent = tr('กำลังบันทึก…', 'Saving…');
  try {
    const id = cfg.id, dataPath = `${id}/payload.json`, originalPath = `${id}/original.${ext}`;
    const store = db().storage.from('templates');
    let r = await store.upload(originalPath, new Blob([bytes], { type: MIME[ext] }), { contentType: MIME[ext], upsert: true });
    if (r.error) throw r.error;
    const json = JSON.stringify({ file: b64(payload.file), labels: payload.labels, scripts: payload.scripts });
    r = await store.upload(dataPath, new Blob([json], { type: 'application/json' }), { contentType: 'application/json', upsert: true });
    if (r.error) throw r.error;
    const meta = { ...entry, file: fileName };
    const { error } = await db().from('templates').insert({ id, kind: ext, builtin: false, hidden: false, meta, data_path: dataPath, original_path: originalPath });
    if (error) throw error;
    await HD.cloud.reloadTemplates();
    B().toast(tr('เผยแพร่แม่แบบแล้ว', 'Template published'));
    draft = null; tab = 'list'; render();
  } catch (err) {
    console.error(err);
    B().toast(tr('บันทึกไม่สำเร็จ: ', 'Save failed: ') + err.message);
  } finally { busy = false; const b = $('#draftSave'); if (b) { b.disabled = false; b.textContent = tr('บันทึกและเผยแพร่', 'Save & publish'); } }
}

/* ======================= admins ======================= */
async function renderAdmins() {
  const body = $('#adminBody');
  body.innerHTML = `<div class="admin-card"><div id="adminList" class="muted">${tr('กำลังโหลด…', 'Loading…')}</div>
    <form id="adminAdd" class="admin-inline"><input type="email" id="adminEmail" placeholder="name@hitap.net" required><button class="btn btn-dark btn-sm" type="submit">${tr('เพิ่มผู้ดูแล', 'Add admin')}</button></form>
    <p class="help">${tr('ผู้ดูแลต้องเข้าสู่ระบบด้วยอีเมลนี้ (ลิงก์ทางอีเมล) จึงจะเห็นเมนูผู้ดูแลระบบ', 'Admins sign in with this email (magic link) to see the Admin menu')}</p></div>`;
  const { data, error } = await db().from('admins').select('email, added_at').order('added_at');
  if (error) { $('#adminList').textContent = error.message; return; }
  const me = (HD.cloud.user() || {}).email;
  $('#adminList').innerHTML = data.map(a => `<div class="admin-row"><span>${esc(a.email)}${a.email === me ? ` <small class="muted">(${tr('คุณ', 'you')})</small>` : ''}</span>
    <button type="button" class="btn btn-light btn-sm" data-rm="${esc(a.email)}" ${data.length < 2 ? 'disabled' : ''}>${tr('นำออก', 'Remove')}</button></div>`).join('');
  $('#adminList').querySelectorAll('[data-rm]').forEach(b => b.addEventListener('click', async () => {
    if (!confirm(tr(`นำ ${b.dataset.rm} ออกจากผู้ดูแล?`, `Remove ${b.dataset.rm}?`))) return;
    const { error: e } = await db().from('admins').delete().eq('email', b.dataset.rm);
    if (e) return B().toast(e.message);
    if (b.dataset.rm === me) location.reload(); else renderAdmins();
  }));
  $('#adminAdd').addEventListener('submit', async e => {
    e.preventDefault();
    const email = $('#adminEmail').value.trim().toLowerCase();
    const { error: er } = await db().from('admins').insert({ email });
    if (er) return B().toast(er.message);
    B().toast(tr('เพิ่มผู้ดูแลแล้ว', 'Admin added'));
    renderAdmins();
  });
}

HD.admin = { render };
if (location.hash.startsWith('#/admin')) render();
if (HD.cloud) HD.cloud.onChange(() => { if (location.hash.startsWith('#/admin') && !draft) render(); });
