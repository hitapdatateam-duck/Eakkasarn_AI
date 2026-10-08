/* PDF engine: pdf.js renders the pages, inline inputs sit on each field, pdf-lib writes the result.
   Fillable PDFs: values go into their AcroForm fields (and the form's own Calculate scripts run through a
   small Acrobat shim). Flat PDFs: values are drawn at the coordinates listed in the template config. */
(() => {
  'use strict';
  const HD = window.HD, U = HD.util;
  const PT = 96 / 72;               // CSS px per PDF point
  const RENDER_SCALE = 2;           // canvas resolution multiplier
  const cache = {};

  async function ensurePdfjs() {
    if (!window.pdfjsLib) await U.loadScript(U.LIBS.pdfjs);
    window.pdfjsLib.GlobalWorkerOptions.workerSrc = U.LIBS.pdfjsWorker;
    return window.pdfjsLib;
  }
  async function ensurePdfLib() {
    if (!window.PDFLib) await U.loadScript(U.LIBS.pdflib);
    if (!window.fontkit) await U.loadScript(U.LIBS.fontkit);
  }
  async function getPdf(ctx) {
    const lib = await ensurePdfjs();
    if (!cache[ctx.tpl.id]) cache[ctx.tpl.id] = lib.getDocument({ data: U.b64ToBytes(ctx.data.file), isEvalSupported: false }).promise;
    return cache[ctx.tpl.id];
  }

  /* ---------------- Acrobat calculation shim ---------------- */
  function fieldIndex(tpl) {
    const idx = {};
    for (const f of tpl.fields) for (const p of f.pdf || []) idx[p.name] = { f, p };
    return idx;
  }
  const asAcroValue = v => { if (typeof v === 'number') return v; const s = String(v ?? ''); const n = U.num(s); return n != null && s.trim() !== '' ? n : s; };

  /** Run the form's Calculate scripts against current values; returns {fieldName: value}. */
  function computeCalcs(ctx) {
    const scripts = ctx.data.scripts || {}, idx = fieldIndex(ctx.tpl), out = {};
    const names = Object.keys(scripts).filter(n => scripts[n].calculate);
    if (!names.length) return out;
    const rawOf = name => {
      if (name in out) return out[name];
      const e = idx[name];
      if (!e) return '';
      const { f, p } = e;
      if (f.type === 'check') return ctx.raw(f.key) ? (p.exp || 'Yes') : 'Off';
      if (f.type === 'yesno') return ctx.raw(f.key) === p.role ? (p.exp || 'Yes') : 'Off';
      return ctx.raw(f.key) ?? '';
    };
    let assigned = null;
    const fieldObj = name => ({
      get value() { return asAcroValue(rawOf(name)); },
      set value(v) { out[name] = v; if (assigned) assigned.add(name); },
      get valueAsString() { return String(rawOf(name)); },
    });
    const AFSimple_Calculate = (op, list) => {
      const nums = [].concat(list).map(n => +String(asAcroValue(rawOf(n))).replace(/,/g, '') || 0);
      const r = op === 'SUM' ? nums.reduce((a, b) => a + b, 0) : op === 'PRD' ? nums.reduce((a, b) => a * b, 1)
        : op === 'AVG' ? nums.reduce((a, b) => a + b, 0) / (nums.length || 1) : op === 'MIN' ? Math.min(...nums) : Math.max(...nums);
      ctx._event.value = r;
    };
    const noop = () => {};
    for (let pass = 0; pass < 3; pass++) {
      for (const name of names) {
        const before = rawOf(name);
        const event = { value: before, target: fieldObj(name), rc: true };
        ctx._event = event;
        assigned = new Set();
        const doc = { getField: fieldObj };
        try {
          // The scripts come from the template PDF itself (as in Acrobat); only the shim globals are exposed.
          const fn = new Function('event', 'getField', 'AFSimple_Calculate', 'AFNumber_Format', 'AFNumber_Keystroke', 'AFDate_FormatEx', 'AFDate_KeystrokeEx', 'app', 'util', scripts[name].calculate);
          fn.call(doc, event, fieldObj, AFSimple_Calculate, noop, noop, noop, noop, { alert: noop }, {});
          // scripts either set event.value or assign this.getField(name).value directly
          if (!assigned.has(name)) out[name] = event.value;
        } catch (e) { /* a broken script just leaves its field empty */ }
      }
    }
    return out;
  }
  /** AFNumber_Format(dec, sepStyle, …) → formatted display, as Acrobat would show it. */
  function formatFor(ctx, name, v) {
    const s = ctx.data.scripts && ctx.data.scripts[name] && ctx.data.scripts[name].format;
    if (v === '' || v == null) return '';
    const m = s && /AFNumber_Format\(\s*(\d+)\s*,\s*(\d+)/.exec(s);
    if (m) {
      const n = U.num(v);
      if (n == null) return String(v);
      const dec = +m[1], sep = +m[2];
      const str = n.toLocaleString('en-US', { minimumFractionDigits: dec, maximumFractionDigits: dec });
      return sep === 1 || sep === 3 ? str.replace(/,/g, '') : str;
    }
    return String(v);
  }
  function displayOf(ctx, f, p, calcs) {
    if (f.type === 'calc') return formatFor(ctx, p.name, calcs[p.name] ?? '');
    if (f.type === 'date') return ctx.value(f.key);
    const v = ctx.value(f.key);
    return p ? formatFor(ctx, p.name, v) : v;
  }

  /* ---------------- preview ---------------- */
  async function render(ctx) {
    const pdf = await getPdf(ctx);
    await document.fonts.load('16px "TH SarabunPSK"').catch(() => {});
    const root = document.createElement('div');
    root.className = 'pdf-doc';
    const dpr = window.devicePixelRatio || 1;
    for (let i = 1; i <= pdf.numPages; i++) {
      const page = await pdf.getPage(i);
      const vp1 = page.getViewport({ scale: 1 });
      const W = vp1.width, H = vp1.height;
      const vp = page.getViewport({ scale: PT * Math.max(RENDER_SCALE, dpr * 1.5) });
      const wrap = document.createElement('section');
      wrap.className = 'hd-page pdf-page';
      wrap.style.width = W * PT + 'px'; wrap.style.height = H * PT + 'px';
      wrap.dataset.page = i; wrap.dataset.h = H;
      const canvas = document.createElement('canvas');
      canvas.width = vp.width; canvas.height = vp.height;
      // intent 'print' renders without requestAnimationFrame, which stalls in hidden tabs
      await page.render({ canvasContext: canvas.getContext('2d'), viewport: vp, intent: 'print', annotationMode: window.pdfjsLib.AnnotationMode.DISABLE }).promise;
      wrap.append(canvas);
      const layer = document.createElement('div');
      layer.className = 'pdf-layer';
      wrap.append(layer);
      root.append(wrap);
    }
    buildOverlay(ctx, root);
    return root;
  }
  const place = (el, rect, H) => {
    const [x, y, w, h] = rect;
    el.style.left = x * PT + 'px'; el.style.top = (H - y - h) * PT + 'px';
    el.style.width = w * PT + 'px'; el.style.height = h * PT + 'px';
  };
  function buildOverlay(ctx, root) {
    const layers = {};
    root.querySelectorAll('.pdf-page').forEach(p => { layers[p.dataset.page] = { layer: p.querySelector('.pdf-layer'), H: +p.dataset.h }; });
    const add = (page, rect, el) => { const L = layers[page]; if (!L) return; place(el, rect, L.H); L.layer.append(el); };
    for (const f of ctx.tpl.fields) {
      if (f.hide) continue;
      const spots = f.pdf || (f.ov ? [{ ...f.ov, ov: true }] : []);
      for (const p of spots) {
        const h = p.rect[3], size = p.size || (f.ov ? f.ov.size : 0) || (p.multiline || h > 34 ? 13 : Math.min(15, Math.max(9, h * 0.62)));
        let el;
        if (f.type === 'check' || f.type === 'yesno') {
          el = document.createElement('button');
          el.type = 'button'; el.className = 'pf pf-chk';
          el.dataset.key = f.key; if (p.role) el.dataset.role = p.role;
        } else if (f.type === 'choice') {
          el = document.createElement('select');
          el.className = 'pf pf-sel'; el.dataset.key = f.key;
          el.innerHTML = (f.options || []).map(o => `<option value="${U.htmlEsc(o.v)}">${U.htmlEsc(o.t)}</option>`).join('');
        } else if (f.type === 'calc') {
          el = document.createElement('div'); el.className = 'pf pf-calc'; el.dataset.key = f.key; el.dataset.name = p.name;
        } else if (f.type === 'date') {
          el = document.createElement('button'); el.type = 'button'; el.className = 'pf pf-date'; el.dataset.key = f.key;
        } else {
          const multi = p.multiline || f.type === 'textarea' || h > 34;
          el = document.createElement(multi ? 'textarea' : 'input');
          el.className = 'pf pf-text' + (multi ? ' multi' : ''); el.dataset.key = f.key; el.spellcheck = false;
          if (p.name) el.dataset.name = p.name;
          if (p.maxLen) el.maxLength = p.maxLen;
          if (f.type === 'money' || f.type === 'number') el.inputMode = 'decimal';
        }
        el.style.fontSize = size * PT + 'px';
        if (p.align === 1) el.style.textAlign = 'center';
        if (p.align === 2 || f.type === 'money' || f.type === 'calc') el.style.textAlign = 'right';
        el.title = ctx.label(f.key);
        add(p.page || 1, p.rect, el);
      }
    }
    for (const s of ctx.tpl.signers || []) {
      const el = document.createElement('button');
      el.type = 'button'; el.className = 'pf-sig'; el.dataset.role = s.role;
      add(s.pdf.page, s.pdf.rect, el);
    }
    update(ctx, root);
  }
  /** Refresh every overlay from current values (cheap: no re-render of the pages). */
  function update(ctx, root) {
    const calcs = computeCalcs(ctx);
    root.querySelectorAll('.pf').forEach(el => {
      const f = ctx.field(el.dataset.key);
      if (!f) return;
      if (el.classList.contains('pf-chk')) {
        const on = f.type === 'yesno' ? ctx.raw(f.key) === el.dataset.role : !!ctx.raw(f.key);
        el.classList.toggle('on', on); el.setAttribute('aria-pressed', String(on));
        el.textContent = on ? '✓' : '';
      } else if (el.classList.contains('pf-sel')) {
        el.value = ctx.raw(f.key) ?? '';
        el.classList.toggle('empty', !String(el.value).trim());
      } else if (el.classList.contains('pf-calc')) {
        el.textContent = formatFor(ctx, el.dataset.name, calcs[el.dataset.name] ?? '');
      } else if (el.classList.contains('pf-date')) {
        const v = ctx.value(f.key);
        el.textContent = v; el.classList.toggle('empty', !v);
      } else if (document.activeElement !== el) {
        const p = (f.pdf || []).find(x => x.name === el.dataset.name) || (f.pdf || [])[0];
        el.value = displayOf(ctx, f, p, calcs);
        el.classList.toggle('empty', !el.value);
      }
    });
    root.querySelectorAll('.pf-sig').forEach(el => {
      const sig = ctx.doc.signs[el.dataset.role];
      el.classList.toggle('signed', !!sig);
      el.innerHTML = sig ? `<img src="${sig.data}" alt="signature">` : `<span class="sig-ask">✍ ${ctx.t('ลงนาม', 'Sign')}</span>`;
    });
  }

  /* ---------------- final PDF ---------------- */
  const fontSizeFor = p => p.size || (p.multiline || p.rect[3] > 34 ? 13 : Math.min(15, Math.max(9, p.rect[3] * 0.62)));
  async function build(ctx) {
    await ensurePdfLib();
    const { PDFDocument, rgb } = window.PDFLib;
    const pdfDoc = await PDFDocument.load(U.b64ToBytes(ctx.data.file), { ignoreEncryption: true });
    pdfDoc.registerFontkit(window.fontkit);
    const font = await pdfDoc.embedFont(await U.thaiFontBytes(), { subset: false });
    const calcs = computeCalcs(ctx);
    let form = null;
    try { form = pdfDoc.getForm(); } catch (e) { form = null; }
    const pages = pdfDoc.getPages();
    for (const f of ctx.tpl.fields) {
      for (const p of f.pdf || []) {
        if (!form) break;
        try {
          if (f.type === 'check') { const cb = form.getCheckBox(p.name); ctx.raw(f.key) ? cb.check() : cb.uncheck(); }
          else if (f.type === 'yesno') { const cb = form.getCheckBox(p.name); ctx.raw(f.key) === p.role ? cb.check() : cb.uncheck(); }
          else if (f.type === 'choice') {
            const dd = form.getDropdown(p.name), v = ctx.raw(f.key);
            if (v && String(v).trim()) { if (!dd.getOptions().includes(v)) dd.addOptions(v); dd.select(v); }
            dd.setFontSize(fontSizeFor(p));
          } else {
            const tf = form.getTextField(p.name);
            const text = displayOf(ctx, f, p, calcs);
            if (tf.getMaxLength() && text.length > tf.getMaxLength()) tf.setMaxLength(text.length);
            tf.setText(text || '');
            // auto-size makes Thai names huge; use the same size as the preview (shrink to fit the width)
            let size = fontSizeFor(p);
            if (!p.multiline) while (size > 6 && text && font.widthOfTextAtSize(text, size) > p.rect[2] - 4) size -= 0.5;
            tf.setFontSize(size);
          }
        } catch (e) { console.warn('field', p.name, e.message); }
      }
      if (f.ov) {
        const text = ctx.value(f.key);
        if (!text) continue;
        const page = pages[(f.ov.page || 1) - 1], [x, y, w, h] = f.ov.rect;
        let size = (f.ov.size || 14) + 1;
        while (size > 7 && font.widthOfTextAtSize(text, size) > w - 4) size -= 0.5;
        page.drawText(text, { x: x + 2, y: y + h * 0.36, size, font, color: rgb(0, 0, 0) });
      }
    }
    // the signature boxes that sit on a text field (COI) must not also print text
    for (const s of ctx.tpl.signers || []) {
      const sig = ctx.doc.signs[s.role];
      if (!sig) continue;
      const png = await pdfDoc.embedPng(U.dataUrlBytes(sig.data));
      const page = pages[(s.pdf.page || 1) - 1], [x, y, w, h] = s.pdf.rect;
      const scale = Math.min(w / png.width, (h + 10) / png.height);
      const dw = png.width * scale, dh = png.height * scale;
      page.drawImage(png, { x: x + (w - dw) / 2, y: y - 2, width: dw, height: dh });
    }
    if (form) {
      try { form.updateFieldAppearances(font); } catch (e) { console.warn('appearances', e.message); }
      // drop the template's Print / Reset buttons from the output
      for (const fl of form.getFields()) if (fl.constructor.name === 'PDFButton') { try { form.removeField(fl); } catch (e) { /* keep */ } }
    }
    const bytes = await pdfDoc.save();
    return new Blob([bytes], { type: 'application/pdf' });
  }

  HD.engines = HD.engines || {};
  HD.engines.pdf = { ext: 'pdf', mime: 'application/pdf', inline: true, render, update, build, computeCalcs };
})();
