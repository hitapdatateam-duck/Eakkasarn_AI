/* Excel engine: parses the workbook, renders each sheet as an HTML grid (styles, merges, images),
   evaluates the formulas the templates use, and writes values / signatures back into the .xlsx. */
(() => {
  'use strict';
  const HD = window.HD, U = HD.util;
  const MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  const cache = {};
  const unesc = s => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&#(\d+);/g, (m, n) => String.fromCharCode(+n)).replace(/&amp;/g, '&');
  const attr = (s, a) => { const m = new RegExp(`\\b${a}="([^"]*)"`).exec(s); return m ? unesc(m[1]) : null; };
  const colNum = c => [...c].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
  const colName = n => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
  const splitRef = r => { const m = /^\$?([A-Z]+)\$?(\d+)$/.exec(r); return m ? [colNum(m[1]), +m[2]] : null; };
  const EPOCH = Date.UTC(1899, 11, 30);
  const serialToISO = n => { const d = new Date(EPOCH + Math.round(n) * 86400000); return d.toISOString().slice(0, 10); };
  const isoToSerial = iso => { const p = U.isoParts(iso); return p ? (Date.UTC(p.y, p.m - 1, p.d) - EPOCH) / 86400000 : null; };

  /* ======================= parse ======================= */
  async function parse(b64) {
    const zip = await JSZip.loadAsync(b64, { base64: true });
    const read = async p => (zip.file(p) ? zip.file(p).async('string') : '');
    const relsOf = async p => {
      const dir = p.replace(/[^/]+$/, ''), name = p.split('/').pop();
      const x = await read(`${dir}_rels/${name}.rels`);
      const out = {};
      for (const m of x.matchAll(/<Relationship [^>]*>/g)) {
        const t = attr(m[0], 'Target');
        out[attr(m[0], 'Id')] = t.startsWith('/') ? t.slice(1) : new URL(t, 'http://x/' + dir).pathname.slice(1);
      }
      return out;
    };
    const sst = [...(await read('xl/sharedStrings.xml')).matchAll(/<si>([\s\S]*?)<\/si>/g)].map(m => unesc([...m[1].matchAll(/<t[^>]*>([^<]*)<\/t>/g)].map(x => x[1]).join('')));
    const st = await read('xl/styles.xml');
    const sec = tag => (new RegExp(`<${tag}[^>]*>([\\s\\S]*?)</${tag}>`).exec(st) || [])[1] || '';
    const numFmts = {};
    for (const m of sec('numFmts').matchAll(/<numFmt [^>]*>/g)) numFmts[attr(m[0], 'numFmtId')] = attr(m[0], 'formatCode');
    const colorOf = x => x ? { rgb: attr(x, 'rgb'), theme: attr(x, 'theme'), tint: attr(x, 'tint'), indexed: attr(x, 'indexed') } : null;
    const fonts = [...sec('fonts').matchAll(/<font>([\s\S]*?)<\/font>|<font\/>/g)].map(m => {
      const f = m[1] || '';
      return { name: attr((/<name [^>]*>/.exec(f) || [''])[0], 'val'), sz: +attr((/<sz [^>]*>/.exec(f) || [''])[0], 'val') || 11,
        b: /<b\/>|<b val="1"/.test(f), i: /<i\/>/.test(f), u: /<u\/>|<u val="single"/.test(f), strike: /<strike\/>/.test(f), color: colorOf((/<color [^>]*>/.exec(f) || [])[0]) };
    });
    const fills = [...sec('fills').matchAll(/<fill>([\s\S]*?)<\/fill>|<fill\/>/g)].map(m => {
      const f = m[1] || '';
      return /patternType="(?!none)/.test(f) ? colorOf((/<fgColor [^>]*>/.exec(f) || [])[0]) : null;
    });
    const side = (b, s) => { const m = new RegExp(`<${s}( [^>]*)?(?:/>|>([\\s\\S]*?)</${s}>)`).exec(b); return m && attr(m[0], 'style') ? { style: attr(m[0], 'style'), color: colorOf((/<color [^>]*>/.exec(m[2] || '') || [])[0]) } : null; };
    const borders = [...sec('borders').matchAll(/<border[^s>]*>([\s\S]*?)<\/border>|<border[^s>]*\/>/g)].map(m => {
      const b = m[1] || '';
      return { l: side(b, 'left'), r: side(b, 'right'), t: side(b, 'top'), b: side(b, 'bottom') };
    });
    const xfs = [...sec('cellXfs').matchAll(/<xf [^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g)].map(m => {
      const al = (/<alignment [^>]*>/.exec(m[0]) || [''])[0];
      return { numFmt: +attr(m[0], 'numFmtId') || 0, font: +attr(m[0], 'fontId') || 0, fill: +attr(m[0], 'fillId') || 0, border: +attr(m[0], 'borderId') || 0,
        h: attr(al, 'horizontal'), v: attr(al, 'vertical'), wrap: attr(al, 'wrapText') === '1', indent: +attr(al, 'indent') || 0, shrink: attr(al, 'shrinkToFit') === '1' };
    });
    const theme = await read('xl/theme/theme1.xml');
    const scheme = (/<a:clrScheme[^>]*>([\s\S]*?)<\/a:clrScheme>/.exec(theme) || [])[1] || '';
    const tc = ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'].map(n => {
      const m = new RegExp(`<a:${n}>([\\s\\S]*?)</a:${n}>`).exec(scheme);
      return m ? (attr(m[1], 'lastClr') || attr(m[1], 'val') || '000000') : '000000';
    });
    const themeColors = [tc[1], tc[0], tc[3], tc[2], ...tc.slice(4)];

    const wb = await read('xl/workbook.xml');
    const wbRels = await relsOf('xl/workbook.xml');
    const printAreas = {};
    for (const m of wb.matchAll(/<definedName ([^>]*)>([^<]*)<\/definedName>/g)) if (attr(m[1], 'name') === '_xlnm.Print_Area') printAreas[attr(m[1], 'localSheetId')] = unesc(m[2]);
    const sheets = [];
    let i = 0;
    for (const m of wb.matchAll(/<sheet [^>]*>/g)) {
      const path = wbRels[attr(m[0], 'r:id')];
      const xml = await read(path);
      const sh = { name: attr(m[0], 'name'), hidden: (attr(m[0], 'state') || '') !== '' && attr(m[0], 'state') !== 'visible', path, index: i };
      const fmtPr = (/<sheetFormatPr [^>]*>/.exec(xml) || [''])[0];
      sh.defColW = +attr(fmtPr, 'defaultColWidth') || (+attr(fmtPr, 'baseColWidth') || 8) + 0.71;
      sh.defRowH = +attr(fmtPr, 'defaultRowHeight') || 15;
      sh.cols = [];
      for (const c of xml.matchAll(/<col [^>]*>/g)) sh.cols.push({ min: +attr(c[0], 'min'), max: +attr(c[0], 'max'), width: +attr(c[0], 'width'), hidden: attr(c[0], 'hidden') === '1' });
      sh.rows = {}; sh.cells = {};
      const shared = {};
      for (const r of xml.matchAll(/<row ([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g)) {
        const rn = +attr(r[1], 'r');
        sh.rows[rn] = { ht: +attr(r[1], 'ht') || null, hidden: attr(r[1], 'hidden') === '1' };
        for (const c of (r[2] || '').matchAll(/<c ([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
          const ref = attr(c[1], 'r'), t = attr(c[1], 't'), s = +attr(c[1], 's') || 0, inner = c[2] || '';
          const fm = /<f([^>]*)>([\s\S]*?)<\/f>|<f([^>]*)\/>/.exec(inner);
          let f = null;
          if (fm) {
            const fa = fm[1] || fm[3] || '', text = fm[2] != null ? unesc(fm[2]) : '';
            if (attr(fa, 't') === 'shared') {
              const si = attr(fa, 'si');
              if (text) shared[si] = { f: text, ref };
              f = text || (shared[si] ? shiftFormula(shared[si].f, shared[si].ref, ref) : null);
            } else f = text;
          }
          const vm = /<v>([\s\S]*?)<\/v>/.exec(inner);
          let v = vm ? unesc(vm[1]) : null;
          if (t === 's' && v != null) v = sst[+v];
          else if (t === 'inlineStr') v = unesc([...inner.matchAll(/<t[^>]*>([^<]*)<\/t>/g)].map(x => x[1]).join(''));
          else if (v != null && t !== 'str' && t !== 'e' && t !== 'b' && v !== '' && !isNaN(+v)) v = +v;
          sh.cells[ref] = { s, v, f, t };
        }
      }
      sh.merges = [...xml.matchAll(/<mergeCell ref="([^"]+)"/g)].map(x => x[1]);
      sh.printArea = printAreas[i] || null;
      sh.dimension = attr((/<dimension [^>]*>/.exec(xml) || [''])[0], 'ref');
      // images
      sh.images = [];
      const dm = /<drawing r:id="([^"]+)"/.exec(xml);
      if (dm) {
        const sheetRels = await relsOf(path), dpath = sheetRels[dm[1]];
        sh.drawingPath = dpath;
        const dx = await read(dpath), dRels = await relsOf(dpath);
        for (const a of dx.matchAll(/<xdr:(twoCellAnchor|oneCellAnchor)[\s\S]*?<\/xdr:\1>/g)) {
          const blip = /r:embed="([^"]+)"/.exec(a[0]);
          if (!blip || !dRels[blip[1]]) continue;
          const pos = tag => { const x = new RegExp(`<xdr:${tag}>([\\s\\S]*?)</xdr:${tag}>`).exec(a[0]); if (!x) return null; const g = n => +((new RegExp(`<xdr:${n}>(-?\\d+)</xdr:${n}>`).exec(x[1]) || [])[1] || 0); return { col: g('col'), colOff: g('colOff'), row: g('row'), rowOff: g('rowOff') }; };
          const ext = /<xdr:ext cx="(\d+)" cy="(\d+)"/.exec(a[0]) || /<a:ext cx="(\d+)" cy="(\d+)"/.exec(a[0]);
          const media = zip.file(dRels[blip[1]]);
          if (!media) continue;
          const ext2 = dRels[blip[1]].split('.').pop().toLowerCase();
          sh.images.push({ from: pos('from'), to: pos('to'), ext: ext ? { cx: +ext[1], cy: +ext[2] } : null, src: `data:image/${ext2 === 'jpg' ? 'jpeg' : ext2};base64,` + await media.async('base64') });
        }
      }
      sheets.push(sh);
      i++;
    }
    return { zip, sheets, styles: { numFmts, fonts, fills, borders, xfs, themeColors } };
  }
  /** Move relative refs of a shared formula from its master cell to `to`. */
  function shiftFormula(f, from, to) {
    const [c1, r1] = splitRef(from), [c2, r2] = splitRef(to), dc = c2 - c1, dr = r2 - r1;
    return f.replace(/('[^']+'!|[A-Za-z_][\w.]*!)?(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\w(])/g, (m, sh, ca, c, ra, r) =>
      `${sh || ''}${ca}${ca ? c : colName(colNum(c) + dc)}${ra}${ra ? r : +r + dr}`);
  }

  /* ======================= number formats ======================= */
  const BUILTIN = { 0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00', 9: '0%', 10: '0.00%', 14: 'm/d/yyyy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy', 18: 'h:mm AM/PM', 20: 'h:mm', 21: 'h:mm:ss', 22: 'm/d/yyyy h:mm', 37: '#,##0 ;(#,##0)', 38: '#,##0 ;[Red](#,##0)', 39: '#,##0.00;(#,##0.00)', 40: '#,##0.00;[Red](#,##0.00)', 43: '#,##0.00', 44: '#,##0.00', 49: '@' };
  const codeOf = (styles, xf) => styles.numFmts[xf.numFmt] || BUILTIN[xf.numFmt] || 'General';
  const isDateCode = code => { const c = code.replace(/"[^"]*"|\[[^\]]*\]|\\./g, ''); return /[dmyhs]/i.test(c) && !/^[#0.,%\s]*$/.test(c) && !/General/i.test(c); };
  function fmtValue(v, code) {
    if (v == null || v === '') return '';
    if (typeof v !== 'number') { const sections = code.split(';'); return sections[3] ? sections[3].replace(/"([^"]*)"/g, '$1').replace('@', v) : String(v); }
    if (/^General$/i.test(code)) return Number.isInteger(v) ? String(v) : String(+v.toPrecision(10));
    const locale = ((/\[\$-([0-9A-Fa-f]+)\]/.exec(code) || [])[1] || '').toUpperCase();
    // [$-107041E]: LCID 041E (Thai) with calendar 07 (Thai Buddhist)
    const thai = /41E$/.test(locale), buddhist = locale.length > 4 && locale.slice(-6, -4) === '07';
    let sections = code.split(';');
    let sec = sections[0];
    if (v < 0 && sections[1]) { sec = sections[1]; v = -v; } else if (v === 0 && sections[2]) sec = sections[2];
    sec = sec.replace(/\[[^\]]*\]/g, '');
    if (isDateCode(sec)) return fmtDateCode(v, sec, thai, buddhist);
    const lit = [];
    let s = sec.replace(/"([^"]*)"/g, (m, x) => { lit.push(x); return `\u0001${lit.length - 1}\u0001`; }).replace(/\\(.)/g, (m, x) => { lit.push(x); return `\u0001${lit.length - 1}\u0001`; })
      .replace(/_./g, ' ').replace(/\*./g, '');
    const pct = s.includes('%');
    if (pct) v *= 100;
    const m = /[#0?][#0?,]*(\.[0#?]+)?/.exec(s);
    if (!m) return s.replace(/\u0001(\d+)\u0001/g, (x, i) => lit[i]);
    const dec = m[1] ? m[1].length - 1 : 0, comma = m[0].includes(',');
    let num = Math.abs(v).toLocaleString('en-US', { minimumFractionDigits: dec, maximumFractionDigits: dec, useGrouping: comma });
    if (/^0+$/.test(m[0].split('.')[0].replace(/,/g, '')) === false && /^#/.test(m[0]) && Math.abs(v) < 1 && !dec) num = num === '0' ? '' : num;
    s = s.slice(0, m.index) + (v < 0 && !sections[1] ? '-' : '') + num + s.slice(m.index + m[0].length);
    return s.replace(/\u0001(\d+)\u0001/g, (x, i) => lit[i]).trim();
  }
  function fmtDateCode(serial, code, thai, buddhist) {
    const d = new Date(EPOCH + Math.floor(serial) * 86400000), frac = serial - Math.floor(serial);
    const Y = d.getUTCFullYear() + (buddhist ? 543 : 0), M = d.getUTCMonth(), D = d.getUTCDate();
    const hh = Math.floor(frac * 24), mi = Math.round((frac * 24 - hh) * 60);
    const mon = thai ? U.TH_MONTHS : U.EN_MONTHS, monS = thai ? U.TH_ABBR : U.EN_MONTHS.map(x => x.slice(0, 3));
    const lit = [];
    let s = code.replace(/"([^"]*)"/g, (m, x) => { lit.push(x); return `\u0001${lit.length - 1}\u0001`; }).replace(/\\(.)/g, (m, x) => { lit.push(x); return `\u0001${lit.length - 1}\u0001`; }).replace(/@|;/g, '');
    s = s.replace(/yyyy|yy|bbbb|bb|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|AM\/PM/gi, t => {
      const k = t.toLowerCase();
      if (k === 'yyyy' || k === 'bbbb') return String(k === 'bbbb' ? d.getUTCFullYear() + 543 : Y);
      if (k === 'yy' || k === 'bb') return String(k === 'bb' ? d.getUTCFullYear() + 543 : Y).slice(-2);
      if (k === 'mmmm') return mon[M];
      if (k === 'mmm') return monS[M];
      if (k === 'mm') return /h/i.test(code) && false ? String(mi).padStart(2, '0') : String(M + 1).padStart(2, '0');
      if (k === 'm') return String(M + 1);
      if (k === 'dddd' || k === 'ddd') return '';
      if (k === 'dd') return String(D).padStart(2, '0');
      if (k === 'd') return String(D);
      if (k === 'hh') return String(hh).padStart(2, '0');
      if (k === 'h') return String(hh);
      return '';
    });
    return s.replace(/\u0001(\d+)\u0001/g, (x, i) => lit[i]).trim();
  }

  /* ======================= formulas ======================= */
  function tokenize(f) {
    const out = []; let i = 0;
    while (i < f.length) {
      const c = f[i];
      if (/\s/.test(c)) { i++; continue; }
      if (c === '"') { let j = i + 1, s = ''; while (j < f.length) { if (f[j] === '"' && f[j + 1] === '"') { s += '"'; j += 2; continue; } if (f[j] === '"') break; s += f[j++]; } out.push({ t: 'str', v: s }); i = j + 1; continue; }
      let m = /^((?:'[^']+'|[A-Za-z_฀-๿][\w.฀-๿]*)!)?(\$?[A-Z]{1,3}\$?\d+)(?::(\$?[A-Z]{1,3}\$?\d+))?/.exec(f.slice(i));
      if (m && !/^\(/.test(f.slice(i + m[0].length))) { out.push({ t: 'ref', sheet: m[1] ? m[1].slice(0, -1).replace(/^'|'$/g, '') : null, a: m[2].replace(/\$/g, ''), b: m[3] ? m[3].replace(/\$/g, '') : null }); i += m[0].length; continue; }
      m = /^\d+(\.\d+)?(E[+-]?\d+)?/i.exec(f.slice(i));
      if (m) { out.push({ t: 'num', v: +m[0] }); i += m[0].length; continue; }
      m = /^[A-Za-z_][\w.]*(?=\()/.exec(f.slice(i));
      if (m) { out.push({ t: 'fn', v: m[0].toUpperCase() }); i += m[0].length; continue; }
      m = /^(TRUE|FALSE)\b/i.exec(f.slice(i));
      if (m) { out.push({ t: 'bool', v: /true/i.test(m[0]) }); i += m[0].length; continue; }
      m = /^(<>|<=|>=|[-+*/^&=<>(),%:])/.exec(f.slice(i));
      if (m) { out.push({ t: 'op', v: m[0] }); i += m[0].length; continue; }
      throw new Error('formula token: ' + f.slice(i, i + 10));
    }
    return out;
  }
  function evaluate(formula, sheet, wbx) {
    const toks = tokenize(formula.replace(/^=/, '')); let p = 0;
    const peek = () => toks[p], next = () => toks[p++];
    const isOp = v => peek() && peek().t === 'op' && peek().v === v;
    const val = x => (x && x.range ? (x.cells[0] ?? '') : x);
    const numv = x => { x = val(x); if (typeof x === 'number') return x; if (x === '' || x == null) return 0; if (typeof x === 'boolean') return +x; const n = U.num(x); return n == null ? 0 : n; };
    const strv = x => { x = val(x); return x == null ? '' : typeof x === 'number' ? String(+x.toPrecision(12)) : String(x); };
    const cmp = () => { let a = cat(); while (peek() && peek().t === 'op' && ['=', '<>', '<', '>', '<=', '>='].includes(peek().v)) { const o = next().v, b = cat(); const x = val(a), y = val(b); a = o === '=' ? x == y : o === '<>' ? x != y : o === '<' ? x < y : o === '>' ? x > y : o === '<=' ? x <= y : x >= y; } return a; };
    const cat = () => { let a = add(); while (isOp('&')) { next(); a = strv(a) + strv(add()); } return a; };
    const add = () => { let a = mul(); while (isOp('+') || isOp('-')) { const o = next().v; const b = mul(); a = o === '+' ? numv(a) + numv(b) : numv(a) - numv(b); } return a; };
    const mul = () => { let a = pow(); while (isOp('*') || isOp('/')) { const o = next().v; const b = pow(); a = o === '*' ? numv(a) * numv(b) : numv(a) / numv(b); } return a; };
    const pow = () => { let a = unary(); while (isOp('^')) { next(); a = Math.pow(numv(a), numv(unary())); } return a; };
    const unary = () => { if (isOp('-')) { next(); return -numv(unary()); } if (isOp('+')) { next(); return unary(); } let a = prim(); if (isOp('%')) { next(); a = numv(a) / 100; } return a; };
    const prim = () => {
      const t = next();
      if (!t) return '';
      if (t.t === 'num' || t.t === 'str' || t.t === 'bool') return t.v;
      if (t.t === 'ref') {
        const sh = t.sheet ? wbx.sheet(t.sheet) : sheet;
        if (!t.b) return wbx.get(sh, t.a);
        const [c1, r1] = splitRef(t.a), [c2, r2] = splitRef(t.b), cells = [];
        for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) cells.push(wbx.get(sh, colName(c) + r));
        return { range: true, cells };
      }
      if (t.t === 'op' && t.v === '(') { const v = cmp(); next(); return v; }
      if (t.t === 'fn') {
        next(); const args = [];
        if (!isOp(')')) { do { if (isOp(',')) next(); args.push(cmp()); } while (isOp(',')); }
        next();
        return call(t.v, args);
      }
      throw new Error('bad formula');
    };
    const flat = args => args.flatMap(a => (a && a.range ? a.cells : [a]));
    const call = (name, a) => {
      switch (name) {
        case 'SUM': return flat(a).reduce((s, x) => s + (typeof x === 'number' ? x : U.num(x) || 0), 0);
        case 'MAX': return Math.max(...flat(a).map(numv)); case 'MIN': return Math.min(...flat(a).map(numv));
        case 'AVERAGE': { const v = flat(a).filter(x => x !== '' && x != null).map(numv); return v.reduce((s, x) => s + x, 0) / (v.length || 1); }
        case 'ROUND': { const k = Math.pow(10, numv(a[1])); return Math.round(numv(a[0]) * k) / k; }
        case 'TODAY': return isoToSerial(U.todayISO());
        case 'DAY': return new Date(EPOCH + numv(a[0]) * 86400000).getUTCDate();
        case 'MONTH': return new Date(EPOCH + numv(a[0]) * 86400000).getUTCMonth() + 1;
        case 'YEAR': return new Date(EPOCH + numv(a[0]) * 86400000).getUTCFullYear();
        case 'DATE': return (Date.UTC(numv(a[0]), numv(a[1]) - 1, numv(a[2])) - EPOCH) / 86400000;
        case 'TEXT': return fmtValue(typeof val(a[0]) === 'number' ? val(a[0]) : numv(a[0]), strv(a[1]));
        case 'BAHTTEXT': return U.bahtText(numv(a[0]));
        case 'IF': return val(a[0]) ? val(a[1]) ?? true : val(a[2]) ?? false;
        case 'AND': return flat(a).every(x => !!val(x)); case 'OR': return flat(a).some(x => !!val(x));
        case 'CONCATENATE': return a.map(strv).join('');
        case 'LEN': return strv(a[0]).length;
        case 'ABS': return Math.abs(numv(a[0]));
        default: return '';
      }
    };
    const r = cmp();
    return val(r);
  }

  /** Workbook value access: user input overrides constants; formulas evaluate (memoised). */
  function evaluator(wb, ctx) {
    const inputs = {};
    for (const f of ctx.tpl.fields) if (f.xl) inputs[`${f.xl.sheet}!${f.xl.ref}`] = f;
    const memo = {}, busy = {};
    const api = {
      sheet: name => wb.sheets.find(s => s.name === name),
      input: (sh, ref) => inputs[`${sh.name}!${ref}`],
      get(sh, ref) {
        if (!sh) return '';
        const id = `${sh.name}!${ref}`;
        if (id in memo) return memo[id];
        const f = inputs[id];
        if (f) {
          const raw = ctx.raw(f.key);
          if (raw == null || raw === '') return (memo[id] = '');
          if (f.type === 'date') return (memo[id] = isoToSerial(raw) ?? raw);
          const n = U.num(raw);
          return (memo[id] = n != null && (f.type === 'money' || f.type === 'number' || /^-?[\d,.]+$/.test(String(raw).trim())) ? n : raw);
        }
        const c = sh.cells[ref];
        if (!c) return '';
        if (c.f) {
          if (busy[id]) return 0;
          busy[id] = true;
          try { memo[id] = evaluate(c.f, sh, api); } catch (e) { memo[id] = c.v ?? ''; }
          busy[id] = false;
          return memo[id];
        }
        return (memo[id] = c.v ?? '');
      },
    };
    return api;
  }

  /* ======================= render ======================= */
  const BORDER = { thin: '1px solid', hair: '1px solid', dotted: '1px dotted', dashed: '1px dashed', mediumDashed: '2px dashed', dashDot: '1px dashed', mediumDashDot: '2px dashed', dashDotDot: '1px dotted', slantDashDot: '2px dashed', medium: '2px solid', thick: '3px solid', double: '3px double' };
  function css(color, styles, def) {
    if (!color) return def;
    let hex = null;
    if (color.rgb) hex = color.rgb.slice(-6);
    else if (color.theme != null) hex = styles.themeColors[+color.theme] || '000000';
    else if (color.indexed != null) { const ix = +color.indexed; hex = ix === 64 ? (def === 'transparent' ? null : '000000') : ix === 9 || ix === 1 ? (ix === 9 ? 'FFFFFF' : 'FFFFFF') : ix === 8 || ix === 0 ? '000000' : null; }
    if (!hex) return def;
    let [r, g, b] = [0, 2, 4].map(i => parseInt(hex.slice(i, i + 2), 16));
    const tint = +color.tint || 0;
    if (tint) [r, g, b] = [r, g, b].map(x => Math.round(tint < 0 ? x * (1 + tint) : x + (255 - x) * tint));
    return `rgb(${r},${g},${b})`;
  }
  const FONT_MAP = n => {
    if (!n) return "'TH SarabunPSK', sans-serif";
    if (/sarabun|angsana|cordia|browallia|eucrosia|jasmine|lily|freesia|iris|niramit|kodchasan/i.test(n)) return `'${n}', 'TH SarabunPSK', sans-serif`;
    return `'${n}', 'TH SarabunPSK', Arial, sans-serif`;
  };
  const colPx = (sh, c) => {
    const col = sh.cols.find(x => c >= x.min && c <= x.max);
    if (col && col.hidden) return 0;
    const w = col && col.width ? col.width : sh.defColW;
    return Math.round(w * 7 + 5);
  };
  const rowPx = (sh, r) => { const row = sh.rows[r]; if (row && row.hidden) return 0; return Math.round(((row && row.ht) || sh.defRowH) * 4 / 3); };
  function usedRange(sh) {
    const area = sh.printArea && /\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)/.exec(sh.printArea.split(',')[0]);
    if (area) return [colNum(area[1]), +area[2], colNum(area[3]), +area[4]];
    let mc = 1, mr = 1;
    for (const [ref, c] of Object.entries(sh.cells)) { if (c.v === null && !c.f && !c.s) continue; const [col, row] = splitRef(ref); if (c.v !== '' && c.v != null || c.f) { mc = Math.max(mc, col); mr = Math.max(mr, row); } }
    for (const m of sh.merges) { const [, b] = m.split(':'); const [c, r] = splitRef(b); mc = Math.max(mc, c); mr = Math.max(mr, r); }
    return [1, 1, Math.min(mc, 60), Math.min(mr, 400)];
  }

  async function getWb(ctx) {
    if (!cache[ctx.tpl.id]) cache[ctx.tpl.id] = parse(ctx.data.file);
    return cache[ctx.tpl.id];
  }

  async function render(ctx) {
    const wb = await getWb(ctx);
    await document.fonts.load('16px "TH SarabunPSK"').catch(() => {});
    const ev = evaluator(wb, ctx), S = wb.styles;
    const root = document.createElement('div');
    root.className = 'xl-doc';
    for (const sh of wb.sheets) {
      if (sh.hidden) continue;
      const [c1, r1, c2, r2] = usedRange(sh);
      const merged = {}, covered = new Set();
      for (const m of sh.merges) {
        const [a, b] = m.split(':'), [ca, ra] = splitRef(a), [cb, rb] = splitRef(b);
        if (ca > c2 || ra > r2) continue;
        merged[a] = { cs: Math.min(cb, c2) - ca + 1, rs: Math.min(rb, r2) - ra + 1, end: colName(Math.min(cb, c2)) + Math.min(rb, r2) };
        for (let r = ra; r <= rb; r++) for (let c = ca; c <= cb; c++) if (r !== ra || c !== ca) covered.add(colName(c) + r);
      }
      const page = document.createElement('section');
      page.className = 'hd-page xl-page';
      page.dataset.sheet = sh.name;
      const title = document.createElement('div');
      title.className = 'xl-tab';
      title.textContent = sh.name;
      page.append(title);
      const grid = document.createElement('div');
      grid.className = 'xl-grid';
      const widths = []; for (let c = c1; c <= c2; c++) widths.push(colPx(sh, c));
      const heights = []; for (let r = r1; r <= r2; r++) heights.push(rowPx(sh, r));
      const table = document.createElement('table');
      table.className = 'xl';
      table.style.width = widths.reduce((a, b) => a + b, 0) + 'px';
      table.innerHTML = `<colgroup>${widths.map(w => `<col style="width:${w}px">`).join('')}</colgroup>`;
      const tbody = document.createElement('tbody');
      for (let r = r1; r <= r2; r++) {
        const tr = document.createElement('tr');
        tr.style.height = heights[r - r1] + 'px';
        if (!heights[r - r1]) tr.style.display = 'none';
        for (let c = c1; c <= c2; c++) {
          const ref = colName(c) + r;
          if (covered.has(ref)) continue;
          const cell = sh.cells[ref], xf = S.xfs[cell ? cell.s : 0] || {};
          const td = document.createElement('td');
          const mg = merged[ref];
          if (mg) { if (mg.cs > 1) td.colSpan = mg.cs; if (mg.rs > 1) td.rowSpan = mg.rs; }
          const font = S.fonts[xf.font] || {};
          const st = td.style;
          st.fontFamily = FONT_MAP(font.name);
          st.fontSize = (font.sz || 11) + 'pt';
          if (font.b) st.fontWeight = '700';
          if (font.i) st.fontStyle = 'italic';
          if (font.u) st.textDecoration = 'underline';
          st.color = css(font.color, S, '#000');
          const fill = S.fills[xf.fill];
          if (fill) st.background = css(fill, S, 'transparent');
          const bTop = S.borders[xf.border] || {};
          const endCell = mg ? sh.cells[mg.end] : null, bEnd = endCell ? S.borders[(S.xfs[endCell.s] || {}).border] || {} : bTop;
          const bd = (b, side) => b && BORDER[b.style] ? `${BORDER[b.style]} ${css(b.color, S, '#000')}` : '';
          if (bd(bTop.l)) st.borderLeft = bd(bTop.l);
          if (bd(bTop.t)) st.borderTop = bd(bTop.t);
          if (bd(bEnd.r || bTop.r)) st.borderRight = bd(bEnd.r || bTop.r);
          if (bd(bEnd.b || bTop.b)) st.borderBottom = bd(bEnd.b || bTop.b);
          st.verticalAlign = xf.v === 'center' ? 'middle' : xf.v === 'top' ? 'top' : 'bottom';
          if (xf.wrap) td.classList.add('wrap');
          const input = ev.input(sh, ref);
          const val = ev.get(sh, ref);
          const code = codeOf(S, xf);
          const text = typeof val === 'number' || val ? fmtValue(val === '' ? null : val, code) : '';
          const align = xf.h === 'centerContinuous' ? 'center' : xf.h && xf.h !== 'general' ? xf.h : typeof val === 'number' ? 'right' : 'left';
          st.textAlign = align === 'fill' || align === 'justify' || align === 'distributed' ? 'left' : align;
          if (xf.indent) st.paddingLeft = xf.indent * 9 + 2 + 'px';
          if (input && !input.hide) {
            td.classList.add('xl-input');
            const f = input;
            if (f.type === 'date') {
              const b = document.createElement('button');
              b.type = 'button'; b.className = 'xi xi-date'; b.dataset.key = f.key;
              td.append(b);
            } else {
              // the text spills into empty neighbours like Excel; the input only shows while editing
              const view = document.createElement('span');
              view.className = 'xl-v xi-view'; view.dataset.key = f.key;
              const el = document.createElement('input');
              el.className = 'xi'; el.dataset.key = f.key; el.spellcheck = false;
              if (f.type === 'money' || f.type === 'number') el.inputMode = 'decimal';
              el.dataset.code = code;
              el.style.textAlign = st.textAlign;
              td.append(view, el);
            }
          } else if (text) {
            const span = document.createElement('span');
            span.className = 'xl-v';
            span.textContent = text;
            if (cell && cell.f) span.dataset.f = ref;
            td.append(span);
          } else if (cell && cell.f) {
            const span = document.createElement('span'); span.className = 'xl-v'; span.dataset.f = ref; td.append(span);
          }
          td.dataset.ref = ref;
          tr.append(td);
        }
        tbody.append(tr);
      }
      table.append(tbody);
      grid.append(table);
      // pixel offsets for images and signatures
      const xAt = c => widths.slice(0, Math.max(0, c - c1)).reduce((a, b) => a + b, 0);
      const yAt = r => heights.slice(0, Math.max(0, r - r1)).reduce((a, b) => a + b, 0);
      for (const im of sh.images) {
        if (!im.from || im.from.col + 1 > c2 || im.from.row + 1 > r2) continue;
        const img = new Image();
        img.src = im.src; img.className = 'xl-img';
        const x = xAt(im.from.col + 1) + im.from.colOff / 9525, y = yAt(im.from.row + 1) + im.from.rowOff / 9525;
        let w, h;
        if (im.to) { w = xAt(im.to.col + 1) + im.to.colOff / 9525 - x; h = yAt(im.to.row + 1) + im.to.rowOff / 9525 - y; }
        else if (im.ext) { w = im.ext.cx / 9525; h = im.ext.cy / 9525; }
        Object.assign(img.style, { left: x + 'px', top: y + 'px', width: w + 'px', height: h + 'px' });
        grid.append(img);
      }
      for (const s of ctx.tpl.signers || []) {
        if (!s.xl || s.xl.sheet !== sh.name) continue;
        const [a, b] = s.xl.range.split(':'), [ca, ra] = splitRef(a), [cb, rb] = splitRef(b || a);
        const el = document.createElement('button');
        el.type = 'button'; el.className = 'pf-sig xl-sig'; el.dataset.role = s.role;
        Object.assign(el.style, { left: xAt(ca) + 'px', top: yAt(ra) - 18 + 'px', width: xAt(cb + 1) - xAt(ca) + 'px', height: yAt(rb + 1) - yAt(ra) + 18 + 'px' });
        grid.append(el);
      }
      page.style.width = table.style.width;
      page.append(grid);
      root.append(page);
    }
    update(ctx, root);
    return root;
  }

  /** Refresh inputs and formula results without rebuilding the grid. */
  async function update(ctx, root) {
    const wb = await getWb(ctx), ev = evaluator(wb, ctx);
    root.querySelectorAll('.xl-page').forEach(page => {
      const sh = wb.sheets.find(s => s.name === page.dataset.sheet);
      page.querySelectorAll('.xl-v[data-f]').forEach(span => {
        const ref = span.dataset.f, cell = sh.cells[ref], xf = wb.styles.xfs[cell.s] || {};
        const v = ev.get(sh, ref);
        span.textContent = v === '' || v == null ? '' : fmtValue(v, codeOf(wb.styles, xf));
      });
    });
    root.querySelectorAll('.xi').forEach(el => {
      const f = ctx.field(el.dataset.key);
      if (!f) return;
      if (el.classList.contains('xi-date')) { const v = ctx.value(f.key); el.textContent = v; el.classList.toggle('empty', !v); return; }
      const raw = ctx.raw(f.key) ?? '';
      const shown = (f.type === 'money' || f.type === 'number') && U.num(raw) != null ? fmtValue(U.num(raw), el.dataset.code === 'General' ? '#,##0.00' : el.dataset.code) : (ctx.value(f.key) || raw);
      const view = el.parentElement.querySelector('.xi-view');
      if (view) view.textContent = shown;
      el.parentElement.classList.toggle('empty', !String(raw).trim());
      if (document.activeElement === el) return;
      el.value = (f.type === 'money' || f.type === 'number') ? raw : shown;
      el.classList.toggle('empty', !String(raw).trim());
    });
    root.querySelectorAll('.pf-sig').forEach(el => {
      const sig = ctx.doc.signs[el.dataset.role];
      el.classList.toggle('signed', !!sig);
      el.innerHTML = sig ? `<img src="${sig.data}" alt="signature">` : `<span class="sig-ask">✍ ${ctx.t('ลงนาม', 'Sign')}</span>`;
    });
  }

  /* ======================= write ======================= */
  const cellXml = (ref, s, value, kind) => {
    const sa = s ? ` s="${s}"` : '';
    if (kind === 'num') return `<c r="${ref}"${sa}><v>${value}</v></c>`;
    if (kind === 'fstr') return value;
    return `<c r="${ref}"${sa} t="inlineStr"><is><t xml:space="preserve">${U.xmlEsc(value)}</t></is></c>`;
  };
  function setCell(xml, ref, build) {
    const [c, r] = splitRef(ref);
    const cellRe = new RegExp(`<c r="${ref}"(?: [^>]*?)?(?:/>|>[\\s\\S]*?</c>)`);
    const m = cellRe.exec(xml);
    if (m) { const s = attr(m[0], 's'); return xml.replace(m[0], build(s)); }
    const rowRe = new RegExp(`<row r="${r}"([^>]*?)(?:/>|>([\\s\\S]*?)</row>)`);
    const rm = rowRe.exec(xml);
    // a missing cell borrows the style of its left neighbour, so the text matches the form
    const left = rm ? [...(rm[2] || '').matchAll(/<c r="([A-Z]+)\d+"([^>]*?)(?:\/>|>)/g)].filter(x => colNum(x[1]) < c).pop() : null;
    const newCell = build(left ? attr(left[2], 's') : null);
    if (rm) {
      const cells = rm[2] || '';
      const parts = [...cells.matchAll(/<c r="([A-Z]+)\d+"[\s\S]*?(?:\/>|<\/c>)/g)];
      const after = parts.find(p => colNum(p[1]) > c);
      const inner = after ? cells.slice(0, after.index) + newCell + cells.slice(after.index) : cells + newCell;
      return xml.replace(rm[0], `<row r="${r}"${rm[1].replace(/\/$/, '')}>${inner}</row>`);
    }
    const rows = [...xml.matchAll(/<row r="(\d+)"/g)];
    const after = rows.find(x => +x[1] > r);
    const rowXml = `<row r="${r}">${newCell}</row>`;
    if (after) return xml.slice(0, after.index) + rowXml + xml.slice(after.index);
    return xml.replace('</sheetData>', rowXml + '</sheetData>').replace('<sheetData/>', `<sheetData>${rowXml}</sheetData>`);
  }

  async function build(ctx) {
    const wb = await getWb(ctx), ev = evaluator(wb, ctx);
    const zip = await JSZip.loadAsync(ctx.data.file, { base64: true });
    const S = wb.styles;
    for (const sh of wb.sheets) {
      let xml = await zip.file(sh.path).async('string');
      for (const f of ctx.tpl.fields) {
        if (!f.xl || f.xl.sheet !== sh.name) continue;
        const raw = ctx.raw(f.key), shown = ctx.value(f.key);
        if ((raw == null || raw === '') && !shown) continue;
        const cell = sh.cells[f.xl.ref], xf = S.xfs[cell ? cell.s : 0] || {}, dateCell = isDateCode(codeOf(S, xf));
        xml = setCell(xml, f.xl.ref, s => {
          if (f.type === 'date') {
            const serial = isoToSerial(raw);
            return dateCell && serial != null ? cellXml(f.xl.ref, s, serial, 'num') : cellXml(f.xl.ref, s, shown, 'str');
          }
          const n = U.num(raw);
          if (n != null && (f.type === 'money' || f.type === 'number' || /^-?[\d,.]+$/.test(String(raw).trim()))) return cellXml(f.xl.ref, s, n, 'num');
          return cellXml(f.xl.ref, s, shown || raw, 'str');
        });
      }
      // refresh cached formula results so viewers that don't recalculate still show totals
      xml = xml.replace(/<c r="([A-Z]+\d+)"([^>]*)>(<f[^>]*>[\s\S]*?<\/f>|<f[^>]*\/>)(?:<v>[\s\S]*?<\/v>)?<\/c>/g, (m, ref, attrs, f) => {
        const v = ev.get(sh, ref);
        const a = attrs.replace(/ t="[^"]*"/, '');
        if (typeof v === 'number' && isFinite(v)) return `<c r="${ref}"${a}>${f}<v>${v}</v></c>`;
        if (typeof v === 'boolean') return `<c r="${ref}"${a} t="b">${f}<v>${v ? 1 : 0}</v></c>`;
        return `<c r="${ref}"${a} t="str">${f}<v>${U.xmlEsc(v ?? '')}</v></c>`;
      });
      zip.file(sh.path, xml);
    }
    await addSignatures(zip, wb, ctx);
    let wbx = await zip.file('xl/workbook.xml').async('string');
    wbx = /<calcPr[^>]*>/.test(wbx) ? wbx.replace(/<calcPr([^>]*?)\/>/, (m, a) => `<calcPr${a.replace(/ fullCalcOnLoad="[^"]*"/, '')} fullCalcOnLoad="1"/>`) : wbx.replace('</workbook>', '<calcPr fullCalcOnLoad="1"/></workbook>');
    zip.file('xl/workbook.xml', wbx);
    return zip.generateAsync({ type: 'blob', compression: 'DEFLATE', mimeType: MIME });
  }

  /** Signature images as twoCellAnchor pictures in each sheet's drawing part. */
  async function addSignatures(zip, wb, ctx) {
    const signed = (ctx.tpl.signers || []).filter(s => s.xl && ctx.doc.signs[s.role]);
    if (!signed.length) return;
    let ct = await zip.file('[Content_Types].xml').async('string');
    if (!/Extension="png"/i.test(ct)) ct = ct.replace(/(<Types[^>]*>)/, '$1<Default Extension="png" ContentType="image/png"/>');
    let n = 0;
    for (const sh of wb.sheets) {
      const mine = signed.filter(s => s.xl.sheet === sh.name);
      if (!mine.length) continue;
      const dir = sh.path.replace(/[^/]+$/, ''), base = sh.path.split('/').pop();
      let dpath = sh.drawingPath;
      if (!dpath) {
        dpath = `xl/drawings/drawing_hitap_${sh.index + 1}.xml`;
        zip.file(dpath, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"></xdr:wsDr>');
        ct = ct.replace('</Types>', `<Override PartName="/${dpath}" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>`);
        const relPath = `${dir}_rels/${base}.rels`;
        let rels = zip.file(relPath) ? await zip.file(relPath).async('string') : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
        rels = rels.replace('</Relationships>', `<Relationship Id="rIdHitapDr" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/${dpath.split('/').pop()}"/></Relationships>`);
        zip.file(relPath, rels);
        let sx = await zip.file(sh.path).async('string');
        sx = sx.replace(/(<\/sheetData>[\s\S]*?)(<legacyDrawing|<tableParts|<extLst|<\/worksheet>)/, (m, a, b) => `${a}<drawing r:id="rIdHitapDr"/>${b}`);
        if (!/xmlns:r=/.test(sx.slice(0, 600))) sx = sx.replace('<worksheet ', '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ');
        zip.file(sh.path, sx);
      }
      const ddir = dpath.replace(/[^/]+$/, ''), dbase = dpath.split('/').pop(), drelPath = `${ddir}_rels/${dbase}.rels`;
      let drels = zip.file(drelPath) ? await zip.file(drelPath).async('string') : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
      let dx = await zip.file(dpath).async('string');
      for (const s of mine) {
        n++;
        const sig = ctx.doc.signs[s.role];
        const media = `xl/media/hitap_signature${n}.png`;
        zip.file(media, sig.data.split(',')[1], { base64: true });
        drels = drels.replace('</Relationships>', `<Relationship Id="rIdHitapSig${n}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/hitap_signature${n}.png"/></Relationships>`);
        const [a, b] = s.xl.range.split(':'), [ca, ra] = splitRef(a), [cb] = splitRef(b || a);
        // centre a ~1.6cm-high image over the range, sitting on its bottom line
        const widthPx = (() => { let w = 0; for (let c = ca; c <= cb; c++) w += colPx(sh, c); return w; })();
        const hPx = 44, wPx = Math.min(widthPx, Math.round(hPx * sig.w / sig.h));
        const offX = Math.max(0, (widthPx - wPx) / 2);
        let col = ca, rem = offX;
        while (rem > colPx(sh, col) && col < cb) { rem -= colPx(sh, col); col++; }
        // twoCellAnchor: the picture's bottom edge sits on the bottom of the signature row
        let toCol = col, toColOff = rem + wPx;
        while (toColOff > colPx(sh, toCol) && toCol < 16384) { toColOff -= colPx(sh, toCol); toCol++; }
        let fromRow = ra, fromRowOff = rowPx(sh, ra) - hPx;
        while (fromRowOff < 0 && fromRow > 1) { fromRow--; fromRowOff += rowPx(sh, fromRow); }
        fromRowOff = Math.max(0, fromRowOff);
        // some viewers place pictures by the absolute offset rather than the anchor, so fill it in too
        let absX = rem, absY = fromRowOff;
        for (let c = 1; c < col; c++) absX += colPx(sh, c);
        for (let r = 1; r < fromRow; r++) absY += rowPx(sh, r);
        absX = Math.round(absX * 9525); absY = Math.round(absY * 9525);
        const pic = `<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>${col - 1}</xdr:col><xdr:colOff>${Math.round(rem * 9525)}</xdr:colOff><xdr:row>${fromRow - 1}</xdr:row><xdr:rowOff>${Math.round(fromRowOff * 9525)}</xdr:rowOff></xdr:from>`
          + `<xdr:to><xdr:col>${toCol - 1}</xdr:col><xdr:colOff>${Math.round(toColOff * 9525)}</xdr:colOff><xdr:row>${ra - 1}</xdr:row><xdr:rowOff>${Math.round(rowPx(sh, ra) * 9525)}</xdr:rowOff></xdr:to><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${9000 + n}" name="Signature ${n}"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>`
          + `<xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rIdHitapSig${n}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>`
          + `<xdr:spPr><a:xfrm><a:off x="${absX}" y="${absY}"/><a:ext cx="${wPx * 9525}" cy="${hPx * 9525}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:twoCellAnchor>`;
        dx = dx.replace('</xdr:wsDr>', pic + '</xdr:wsDr>');
      }
      zip.file(dpath, dx);
      zip.file(drelPath, drels);
    }
    zip.file('[Content_Types].xml', ct);
  }

  HD.engines = HD.engines || {};
  HD.engines.xlsx = { ext: 'xlsx', mime: MIME, inline: true, render, update, build, _fmtValue: fmtValue, _evaluate: evaluate };
})();
