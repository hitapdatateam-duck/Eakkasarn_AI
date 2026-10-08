/* Shared helpers for the document engines (no dependencies). */
(() => {
  'use strict';
  const HD = (window.HD = window.HD || {});
  const U = (HD.util = {});

  U.TH_MONTHS = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];
  U.TH_ABBR = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
  U.EN_MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  const pad2 = n => String(n).padStart(2, '0');
  U.todayISO = () => { const d = new Date(); return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`; };
  U.isoParts = iso => { const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || ''); return m ? { y: +m[1], m: +m[2], d: +m[3] } : null; };

  /** fmt: th-long (8 ตุลาคม 2569) · en-long (8 October 2026) · dmy-be (8/10/2569) · dmy (08/10/2026) · mdy-short (10/8/26) */
  U.fmtDate = (iso, fmt) => {
    const p = U.isoParts(iso);
    if (!p) return iso || '';
    switch (fmt) {
      case 'th-long': return `${p.d} ${U.TH_MONTHS[p.m - 1]} ${p.y + 543}`;
      case 'en-long': return `${p.d} ${U.EN_MONTHS[p.m - 1]} ${p.y}`;
      case 'dmy-be': return `${p.d}/${p.m}/${p.y + 543}`;
      case 'mdy-short': return `${p.m}/${p.d}/${String(p.y).slice(2)}`;
      default: return `${pad2(p.d)}/${pad2(p.m)}/${p.y}`;
    }
  };
  U.money = v => {
    const n = U.num(v);
    return n == null ? (v || '') : n.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  };
  U.num = v => {
    if (v == null || v === '') return null;
    if (typeof v === 'number') return v;
    const s = String(v).replace(/[,\s฿]|บาท|baht/gi, '');
    return /^-?\d+(\.\d+)?$/.test(s) ? +s : null;
  };
  U.thaiDigits = s => s.replace(/[๐-๙]/g, c => '๐๑๒๓๔๕๖๗๘๙'.indexOf(c));
  const reEsc = s => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const TH_MONTH_ALT = U.TH_MONTHS.map((m, i) => [m, U.TH_ABBR[i], U.TH_ABBR[i].replace(/\./g, '')]).flat().sort((a, b) => b.length - a.length).map(reEsc).join('|');
  const thMonthIndex = s => { const t = s.replace(/\./g, ''); return U.TH_MONTHS.findIndex((m, i) => m === s || U.TH_ABBR[i].replace(/\./g, '') === t); };
  function mkDate(y, m, d) {
    y = +y; m = +m; d = +d;
    if (y < 100) y += 2500;
    if (y > 2400) y -= 543;
    const dt = new Date(y, m - 1, d);
    if (dt.getFullYear() !== y || dt.getMonth() !== m - 1 || dt.getDate() !== d) return null;
    return `${y}-${pad2(m)}-${pad2(d)}`;
  }
  U.parseDate = raw => {
    const s = U.thaiDigits(raw || '');
    if (/วันนี้|today/i.test(s)) return U.todayISO();
    let m;
    if ((m = /(\d{4})-(\d{1,2})-(\d{1,2})/.exec(s))) return mkDate(m[1], m[2], m[3]);
    if ((m = new RegExp(`(\\d{1,2})\\s*(${TH_MONTH_ALT})\\s*(?:พ\\.?ศ\\.?|ค\\.?ศ\\.?)?\\s*(\\d{2,4})`).exec(s))) return mkDate(m[3], thMonthIndex(m[2]) + 1, m[1]);
    const enAlt = 'jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|jul(?:y)?|aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?';
    const enIdx = w => ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'].indexOf(w.slice(0, 3).toLowerCase()) + 1;
    if ((m = new RegExp(`(\\d{1,2})(?:st|nd|rd|th)?\\s+(${enAlt})\\.?,?\\s+(\\d{4})`, 'i').exec(s))) return mkDate(m[3], enIdx(m[2]), m[1]);
    if ((m = new RegExp(`(${enAlt})\\.?\\s+(\\d{1,2})(?:st|nd|rd|th)?,?\\s+(\\d{4})`, 'i').exec(s))) return mkDate(m[3], enIdx(m[1]), m[2]);
    if ((m = /(\d{1,2})[/.-](\d{1,2})[/.-](\d{2,4})/.exec(s))) return mkDate(m[3], m[2], m[1]);
    return null;
  };
  U.thaiIdStatus = raw => {
    const t = (raw || '').trim(), d = t.replace(/\D/g, '');
    if (!/^[\d\s-]+$/.test(t) || d.length !== 13) return null;
    let sum = 0;
    for (let i = 0; i < 12; i++) sum += +d[i] * (13 - i);
    return (11 - (sum % 11)) % 10 === +d[12] ? 'ok' : 'bad';
  };

  /** Excel BAHTTEXT */
  U.bahtText = n => {
    n = Math.round((+n || 0) * 100) / 100;
    const digit = ['ศูนย์', 'หนึ่ง', 'สอง', 'สาม', 'สี่', 'ห้า', 'หก', 'เจ็ด', 'แปด', 'เก้า'];
    const unit = ['', 'สิบ', 'ร้อย', 'พัน', 'หมื่น', 'แสน'];
    const read = num => {
      if (num === 0) return '';
      const s = String(num);
      if (s.length > 6) return read(Math.floor(num / 1e6)) + 'ล้าน' + read(num % 1e6);
      let out = '';
      for (let i = 0; i < s.length; i++) {
        const d = +s[i], pos = s.length - i - 1;
        if (!d) continue;
        if (pos === 1 && d === 1) out += 'สิบ';
        else if (pos === 1 && d === 2) out += 'ยี่สิบ';
        else if (pos === 0 && d === 1 && s.length > 1) out += 'เอ็ด';
        else out += digit[d] + unit[pos];
      }
      return out;
    };
    const neg = n < 0; n = Math.abs(n);
    const baht = Math.floor(n), satang = Math.round((n - baht) * 100);
    let t = '';
    if (baht) t += read(baht) + 'บาท';
    if (satang) t += read(satang) + 'สตางค์'; else t += (baht ? '' : 'ศูนย์บาท') + 'ถ้วน';
    return (neg ? 'ลบ' : '') + t;
  };

  U.xmlEsc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  U.htmlEsc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  U.oneLine = s => (s || '').replace(/\s*\n\s*/g, ' ').trim();
  U.b64ToBytes = b64 => { const bin = atob(b64), a = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) a[i] = bin.charCodeAt(i); return a; };
  U.dataUrlBytes = url => U.b64ToBytes(url.split(',')[1]);

  const scripts = {};
  /** Load a classic script once (works from file:// too). */
  U.loadScript = src => scripts[src] || (scripts[src] = new Promise((res, rej) => {
    const s = document.createElement('script');
    s.src = src; s.onload = res; s.onerror = () => { delete scripts[src]; rej(new Error('load failed: ' + src)); };
    document.head.appendChild(s);
  }));
  /** Template payload {file, labels, scripts} written by the build into templates/data/<id>.js */
  U.loadTemplateData = async tpl => {
    window.HD_DATA = window.HD_DATA || {};
    if (!window.HD_DATA[tpl.id]) await U.loadScript(tpl.data);
    return window.HD_DATA[tpl.id];
  };
  U.LIBS = {
    pdfjs: 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js',
    pdfjsWorker: 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js',
    pdflib: 'https://cdnjs.cloudflare.com/ajax/libs/pdf-lib/1.17.1/pdf-lib.min.js',
    fontkit: 'https://unpkg.com/@pdf-lib/fontkit@1.1.1/dist/fontkit.umd.min.js',
  };
  U.thaiFontBytes = async () => {
    if (!window.HD_FONT_SARABUN) await U.loadScript('fonts/THSarabun.font.js');
    return U.b64ToBytes(window.HD_FONT_SARABUN);
  };
})();
