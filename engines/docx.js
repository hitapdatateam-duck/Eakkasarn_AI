/* Word engine: fills {{token}} runs in a tokenised .docx and previews it with docx-preview.
   Tokens: {{key}} text/date/money · {{chk…}}/check fields · {{sigline_<role>[_tabs]}} signature lines. */
(() => {
  'use strict';
  const HD = window.HD, U = HD.util;
  const RUN_RE = /<w:r(?:\s[^>]*)?>(?:(?!<\/w:r>)[\s\S])*?\{\{(\w+)\}\}[\s\S]*?<\/w:r>/g;
  const TOKEN_T = /<w:t[^>]*>\{\{\w+\}\}<\/w:t>/;
  const CHECKED_SYM = '<w:sym w:font="Wingdings 2" w:char="F052"/>';
  const sigRole = key => key.replace(/^sigline_/, '').replace(/_\d+$/, '');
  const MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

  function drawingRun(rid, n, sig) {
    const w = Math.min(1650000, Math.round(480000 * sig.w / sig.h)), h = Math.round(w * sig.h / sig.w);
    return `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${w}" cy="${h}"/><wp:docPr id="${7000 + n}" name="Signature ${n}"/>`
      + `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">`
      + `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="${7000 + n}" name="signature${n}.png"/><pic:cNvPicPr/></pic:nvPicPr>`
      + `<pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>`
      + `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${w}" cy="${h}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>`
      + `</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
  }

  /** Word draws a list number with its paragraph mark's font; docx-preview only reads the numbering level. */
  async function patchNumberingForBrowser(zip, docXml) {
    const nf = zip.file('word/numbering.xml');
    if (!nf) return;
    let nb = await nf.async('string');
    const absOf = {};
    for (const m of nb.matchAll(/<w:num w:numId="(\d+)"[^>]*>[\s\S]*?<w:abstractNumId w:val="(\d+)"\/>/g)) absOf[m[1]] = m[2];
    const mark = {};
    for (const p of docXml.match(/<w:p[ >][\s\S]*?<\/w:p>/g) || []) {
      const np = /<w:numPr>[\s\S]*?<\/w:numPr>/.exec(p);
      const ppr = np && /<w:pPr>[\s\S]*?<\/w:pPr>/.exec(p);
      const rpr = ppr && /<w:rPr>[\s\S]*?<\/w:rPr>/.exec(ppr[0]);
      const numId = np && (/<w:numId w:val="(\d+)"/.exec(np[0]) || [])[1];
      if (!rpr || !absOf[numId]) continue;
      const key = `${absOf[numId]}:${(/<w:ilvl w:val="(\d+)"/.exec(np[0]) || [0, '0'])[1]}`;
      mark[key] = mark[key] || rpr[0];
    }
    nb = nb.replace(/<w:abstractNum [^>]*w:abstractNumId="(\d+)"[\s\S]*?<\/w:abstractNum>/g, (abs, aid) =>
      abs.replace(/<w:lvl [^>]*w:ilvl="(\d+)"[^>]*>[\s\S]*?<\/w:lvl>/g, (lvl, il) => {
        const m = mark[`${aid}:${il}`];
        const lr = /<w:rPr>([\s\S]*?)<\/w:rPr>/.exec(lvl);
        if (!m || (lr && /w:ascii=/.test(lr[1]))) return lvl;
        const pick = re => (re.exec(m) || [''])[0];
        const keep = lr ? lr[1].replace(/<w:rFonts [^>]*\/>/, '') : '';
        const rpr = `<w:rPr>${pick(/<w:rFonts [^>]*\/>/)}${keep}${/<w:sz /.test(keep) ? '' : pick(/<w:sz [^>]*\/>/)}${/<w:szCs /.test(keep) ? '' : pick(/<w:szCs [^>]*\/>/)}</w:rPr>`;
        return lr ? lvl.replace(lr[0], rpr) : lvl.replace('</w:lvl>', rpr + '</w:lvl>');
      }));
    zip.file('word/numbering.xml', nb);
  }

  /** mode 'preview' → clickable markers; 'final' → the Word file to download. forBrowser patches render-only quirks. */
  async function build(ctx, mode, forBrowser = false) {
    const { doc, data } = ctx, labels = data.labels || {};
    const zip = await JSZip.loadAsync(data.file, { base64: true });
    const images = [];
    const fill = (run, key) => {
      if (key.startsWith('sigline_')) {
        const role = sigRole(key), tabs = (/_(\d+)$/.exec(key) || [])[1];
        if (mode === 'preview') return run.replace(`{{${key}}}`, `⟦s:${role}⟧`);
        const sig = doc.signs && doc.signs[role];
        if (sig) { images.push(sig); return drawingRun(`rIdHitapSig${images.length}`, images.length, sig); }
        if (labels[key]) return run.replace(`{{${key}}}`, U.xmlEsc(labels[key]));
        return run.replace(TOKEN_T, '<w:tab/>'.repeat(+tabs || 3));
      }
      const f = ctx.field(key) || {};
      if (f.type === 'check') {
        if (mode === 'preview') return run.replace(`{{${key}}}`, `⟦c:${key}⟧`);
        return run.replace(TOKEN_T, ctx.raw(key) ? CHECKED_SYM : (labels[key] || '<w:t>☐</w:t>'));
      }
      if (f.cont) return run.replace(`{{${key}}}`, ctx.value(f.cont) ? '' : U.xmlEsc(labels[key] || ''));
      if (mode === 'preview') return run.replace(/<w:highlight w:val="yellow"\/>/, '').replace(`{{${key}}}`, `⟦f:${key}⟧`);
      const v = ctx.value(key);
      // values that replace a dotted leader keep a space either side ("จัดโดย HITAP ณ")
      const padded = v && /^[.…\s]+$/.test(labels[key] || '') ? ` ${v} ` : v;
      if (v) return run.replace(/<w:highlight w:val="yellow"\/>/, '').replace(`{{${key}}}`, U.xmlEsc(padded));
      if (labels[key] != null) return run.replace(`{{${key}}}`, U.xmlEsc(labels[key]));
      if (key.startsWith('sig_')) return run.replace(TOKEN_T, '<w:tab/><w:tab/><w:tab/>');
      return run.replace(`{{${key}}}`, '');
    };
    const parts = Object.keys(zip.files).filter(n => /^word\/(document|header\d*|footer\d*)\.xml$/.test(n));
    let body = '';
    for (const path of parts) {
      const xml = await zip.file(path).async('string');
      if (!xml.includes('{{')) { if (path === 'word/document.xml') body = xml; continue; }
      // a run can hold several tokens (e.g. "ตั้งแต่ {{from}} ถึง {{to}}"), so repeat until none are left
      let out = xml, prev;
      do { prev = out; out = out.replace(RUN_RE, fill); } while (out !== prev && /\{\{\w+\}\}/.test(out));
      zip.file(path, out);
      if (path === 'word/document.xml') body = out;
    }
    if (forBrowser) await patchNumberingForBrowser(zip, body);
    if (images.length) {
      const relsPath = 'word/_rels/document.xml.rels';
      let rels = await zip.file(relsPath).async('string');
      images.forEach((sig, i) => {
        zip.file(`word/media/hitap_signature${i + 1}.png`, sig.data.split(',')[1], { base64: true });
        rels = rels.replace('</Relationships>', `<Relationship Id="rIdHitapSig${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/hitap_signature${i + 1}.png"/></Relationships>`);
      });
      zip.file(relsPath, rels);
      let ct = await zip.file('[Content_Types].xml').async('string');
      if (!/Extension="png"/i.test(ct)) ct = ct.replace(/(<Types[^>]*>)/, '$1<Default Extension="png" ContentType="image/png"/>');
      zip.file('[Content_Types].xml', ct);
    }
    return zip.generateAsync({ type: 'blob', compression: 'DEFLATE', mimeType: MIME });
  }

  /* ---------- preview fix-ups ---------- */
  function decorate(root, ctx) {
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT), nodes = [];
    while (walker.nextNode()) if (walker.currentNode.nodeValue.includes('⟦')) nodes.push(walker.currentNode);
    for (const node of nodes) {
      const frag = document.createDocumentFragment(), s = node.nodeValue, re = /⟦([fsc]):(\w+)⟧/g;
      let last = 0, m;
      while ((m = re.exec(s))) {
        if (m.index > last) frag.append(s.slice(last, m.index));
        frag.append(m[1] === 'f' ? fieldEl(m[2], ctx) : m[1] === 'c' ? checkEl(m[2], ctx) : sigEl(m[2], ctx));
        last = re.lastIndex;
      }
      if (last < s.length) frag.append(s.slice(last));
      node.replaceWith(frag);
    }
  }
  function fieldEl(key, ctx) {
    const el = document.createElement('span'), v = ctx.value(key);
    el.className = 'fld ' + (v ? 'filled' : 'empty');
    el.dataset.key = key; el.tabIndex = 0; el.setAttribute('role', 'button');
    el.textContent = v || `[${ctx.chipLabel(key)}]`;
    return el;
  }
  function checkEl(key, ctx) {
    const el = document.createElement('span'), on = !!ctx.raw(key);
    el.className = 'chk' + (on ? ' on' : '');
    el.dataset.check = key; el.tabIndex = 0; el.setAttribute('role', 'checkbox'); el.setAttribute('aria-checked', String(on));
    el.textContent = on ? '☑' : '☐';
    return el;
  }
  function sigEl(role, ctx) {
    const el = document.createElement('span'), sig = ctx.doc.signs[role];
    el.className = 'sigslot' + (sig ? ' signed' : '');
    el.dataset.role = role; el.tabIndex = 0; el.setAttribute('role', 'button');
    if (sig) { const img = new Image(); img.src = sig.data; img.alt = 'signature'; el.append(img); }
    else el.innerHTML = `<span class="sig-ask">✍ ${ctx.t('คลิกเพื่อลงนาม', 'Click to sign')}</span>`;
    return el;
  }
  /** docx-preview prints cached PAGE/NUMPAGES values; number pages as Word would. */
  function fixPageNumbers(root) {
    const pages = [...root.querySelectorAll('section.docx')];
    pages.forEach((sec, i) => {
      const foot = sec.querySelector('footer');
      if (!foot) return;
      const w = document.createTreeWalker(foot, NodeFilter.SHOW_TEXT), nodes = [];
      while (w.nextNode()) if (w.currentNode.nodeValue.trim()) nodes.push(w.currentNode);
      let state = 0;
      for (const n of nodes) {
        const t = n.nodeValue.trim();
        if (state === 0 && /(หน้าที่|page)$/i.test(t)) { state = 1; continue; }
        if (state === 1 && /^\d+$/.test(t)) { n.nodeValue = n.nodeValue.replace(/\d+/, i + 1); state = 2; continue; }
        if (state === 2 && /^\/\s*\d+$/.test(t)) { n.nodeValue = n.nodeValue.replace(/\d+/, pages.length); break; }
        if (state === 2 && t === '/') continue;
        if (state === 2 && /^\d+$/.test(t)) { n.nodeValue = n.nodeValue.replace(/\d+/, pages.length); break; }
      }
    });
  }
  /** Tabbed header/footer lines → left / centre / right columns (docx-preview ignores tab stops). */
  function layoutHeaderFooterTabs(root) {
    root.querySelectorAll('header p, footer p').forEach(p => {
      const isTab = n => n.nodeType === 1 && (n.classList.contains('docx-tab-stop') || n.querySelector('.docx-tab-stop'));
      const kids = [...p.childNodes];
      if (!kids.some(isTab)) return;
      const groups = [[]];
      for (const n of kids) {
        if (isTab(n)) {
          if (groups[groups.length - 1].length || groups.length === 1) groups.push([]);
          if (!n.classList.contains('docx-tab-stop')) { n.querySelectorAll('.docx-tab-stop').forEach(t => t.remove()); if (n.textContent.trim()) groups[groups.length - 1].push(n); }
          continue;
        }
        groups[groups.length - 1].push(n);
      }
      const nonEmpty = g => g.some(n => n.textContent.trim());
      const left = groups[0], right = groups.length > 1 ? groups[groups.length - 1] : [];
      const middle = groups.slice(1, -1).flat();
      const col = (nodes, align) => { const c = document.createElement('span'); c.className = 'hf-col'; c.style.textAlign = align; c.append(...nodes); return c; };
      p.replaceChildren(col(left, 'left'), col(nonEmpty(middle) ? middle : [], 'center'), col(right, 'right'));
      p.classList.add('hf-tabbed');
    });
  }

  async function render(ctx, box, mode = 'preview') {
    await document.fonts.load('16px "TH SarabunPSK"').catch(() => {});
    const blob = await build(ctx, mode, true);
    const tmp = document.createElement('div');
    await docx.renderAsync(blob, tmp, null, { className: 'docx', inWrapper: true, breakPages: true, ignoreLastRenderedPageBreak: false, experimental: true, renderHeaders: true, renderFooters: true });
    if (mode === 'preview') decorate(tmp, ctx);
    fixPageNumbers(tmp);
    layoutHeaderFooterTabs(tmp);
    return tmp;
  }

  HD.engines = HD.engines || {};
  HD.engines.docx = { ext: 'docx', mime: MIME, inline: false, build: (ctx) => build(ctx, 'final'), render };
})();
