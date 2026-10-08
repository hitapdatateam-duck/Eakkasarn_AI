/* HITAP Docs — free, client-only document drafting for HITAP's templates.
   A rule-based assistant (no paid API, no credits) asks for the key details, the engines in engines/
   fill the original Word / PDF / Excel files, and the preview is the real document with editable fields. */
(() => {
  'use strict';
  const HD = window.HD, U = HD.util;
  const LS = { lang: 'hitap_nda_lang', docs: 'hitap_nda_docs', profile: 'hitap_docs_profile' };
  const $ = (s, el = document) => el.querySelector(s);
  const $$ = (s, el = document) => [...el.querySelectorAll(s)];
  const store = {
    get(k, d) { try { const v = localStorage.getItem(k); return v == null ? d : JSON.parse(v); } catch { return d; } },
    set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); return true; } catch { return false; } },
  };
  const uiLang = () => (document.documentElement.dataset.lang === 'en' ? 'en' : 'th');
  const tr = (th, en) => (uiLang() === 'en' ? en : th);
  const L = (th, en = th) => ({ th, en });
  const { htmlEsc, oneLine } = U;

  /* ================= Built-in NDA templates (tokenised by build_templates.py) ================= */
  const NDA_FIELDS = [
    { key: 'contract_no', type: 'text', label: L('เลขที่สัญญา', 'Contract / reference no.'), chip: L('ระบุเลขที่สัญญา', 'Reference number'),
      ph: L('HITAP-NDA 001/2569', 'HITAP-NDA 001/2026'), note: L('อยู่ที่หัวกระดาษทุกหน้า — เว้นว่างได้ จะเป็นเส้นจุดให้เขียนเอง', 'Printed in the header of every page — leave blank to keep the dotted line'),
      kw: ['เลขที่สัญญา', 'สัญญาเลขที่', 'เลขสัญญา', 'contract number', 'contract no', 'reference number', 'reference no', 'ref no'], group: 0 },
    { key: 'agreement_date', type: 'date', ask: 9, def: 'today', label: L('วันที่ทำสัญญา', 'Agreement date'), chip: L('ระบุวันที่ทำสัญญา', 'Agreement date'),
      q: L('สัญญาลงวันที่เท่าไหร่ครับ', 'What date is the agreement made?'), kw: ['วันที่ทำสัญญา', 'ทำสัญญาวันที่', 'วันทำสัญญา', 'สัญญาลงวันที่', 'agreement date', 'date of agreement', 'signing date'], group: 0 },
    { key: 'recipient_name', type: 'text', ask: 1, label: L('ชื่อ-นามสกุลผู้รับข้อมูล', 'Recipient full name'), chip: L('ระบุชื่อผู้รับข้อมูล', 'Recipient name'), nameField: true,
      ph: L('นางสาวสมหญิง ใจดี', 'Ms. Jane Doe'), q: L('ผู้รับข้อมูลชื่ออะไรครับ (ใส่คำนำหน้าด้วย เช่น นางสาวสมหญิง ใจดี)', 'What is the recipient\'s full name (with title, e.g. Ms. Jane Doe)?'),
      kw: ['ชื่อผู้รับข้อมูล', 'ผู้รับข้อมูลชื่อ', 'ชื่อ-นามสกุล', 'ชื่อนามสกุล', 'ชื่อผู้รับ', 'ชื่อ', 'recipient name', 'full name', 'name'], group: 1 },
    { key: 'recipient_id', type: 'text', ask: 2, label: L('เลขบัตรประชาชน / หนังสือเดินทาง', 'National ID / passport no.'), chip: L('ระบุเลขบัตรประชาชน', 'ID/Passport number'), idField: true,
      ph: L('1-2345-67890-12-3', 'AB1234567'), q: L('เลขบัตรประชาชน (หรือเลขหนังสือเดินทาง) ของผู้รับข้อมูลครับ', 'What is the recipient\'s national ID or passport number?'),
      kw: ['เลขบัตรประจำตัวประชาชน', 'บัตรประจำตัวประชาชนเลขที่', 'เลขประจำตัวประชาชน', 'เลขบัตรประชาชน', 'บัตรประชาชน', 'เลขที่บัตร', 'เลขบัตร', 'เลขหนังสือเดินทาง', 'หนังสือเดินทาง', 'พาสปอร์ต', 'passport number', 'passport no', 'passport', 'national id', 'id number', 'id no', 'id'], group: 1 },
    { key: 'recipient_address', type: 'textarea', ask: 3, label: L('ที่อยู่ตามบัตรประชาชน', 'Residential address'), chip: L('ระบุที่อยู่ตามบัตรประชาชน', 'Address of the recipient'),
      ph: L('99/1 ถนนติวานนท์ ตำบลตลาดขวัญ อำเภอเมืองนนทบุรี จังหวัดนนทบุรี 11000', '99/1 Tiwanon Rd, Muang, Nonthaburi 11000'), q: L('ที่อยู่ตามบัตรประชาชนของผู้รับข้อมูลครับ', 'What is the recipient\'s residential address?'),
      kw: ['ที่อยู่ตามบัตรประชาชน', 'ที่อยู่ตามบัตร', 'อาศัยอยู่ที่', 'ที่อยู่', 'residing at', 'lives at', 'address'], group: 1 },
    { key: 'position', type: 'text', ask: 4, label: L('ตำแหน่ง', 'Position'), chip: L('ระบุตำแหน่ง', 'Position'), ph: L('นักวิจัย', 'Researcher'),
      q: L('ผู้รับข้อมูลอยู่ในตำแหน่งอะไรครับ', 'What is the recipient\'s position?'), kw: ['ในตำแหน่ง', 'ตำแหน่งงาน', 'ตำแหน่ง', 'position', 'job title', 'role'], group: 1 },
    { key: 'employment_date', type: 'date', ask: 5, label: L('วันที่ในสัญญาจ้าง', 'Employment agreement date'), chip: L('ระบุวันที่ในสัญญาจ้าง', 'Employment agreement date'),
      q: L('สัญญาจ้างลงวันที่เท่าไหร่ครับ (เช่น 1 ต.ค. 2569)', 'What date is the employment agreement? (e.g. 1 Oct 2026)'),
      kw: ['วันที่ในสัญญาจ้าง', 'สัญญาจ้างแรงงานลงวันที่', 'สัญญาจ้างลงวันที่', 'สัญญาจ้างวันที่', 'วันที่สัญญาจ้าง', 'วันเริ่มงาน', 'เริ่มงานวันที่', 'เริ่มงาน', 'บรรจุวันที่', 'employment agreement dated', 'employment agreement date', 'employment date', 'employed on', 'employed since', 'hired on', 'hired', 'start date', 'started on'], group: 1 },
    { key: 'project', type: 'text', ask: 6, label: L('ชื่อโครงการวิจัย', 'Research project title'), chip: L('ระบุชื่อโครงการวิจัย', 'Project'),
      ph: L('การประเมินความคุ้มค่าของวัคซีน HPV', 'Economic evaluation of HPV vaccination'), q: L('ชื่อโครงการวิจัยคืออะไรครับ', 'What is the research project title?'),
      kw: ['ชื่อโครงการวิจัย', 'ชื่อโครงการ', 'โครงการวิจัยเรื่อง', 'โครงการวิจัย', 'โครงการ', 'research project', 'project title', 'project'], group: 2 },
    { key: 'information', type: 'textarea', ask: 7, label: L('รายละเอียดข้อมูลที่เปิดเผย', 'Information disclosed'), chip: L('ระบุรายละเอียดข้อมูล', 'Project/Information'),
      ph: L('ข้อมูลการเบิกจ่ายผู้ป่วยใน ปี 2565–2567', 'inpatient claims data, 2022–2024'), q: L('ข้อมูลที่จะเปิดเผยให้ผู้รับข้อมูลคืออะไรครับ (เช่น ข้อมูลการเบิกจ่ายผู้ป่วยใน ปี 2565–2567)', 'What information will be disclosed? (e.g. inpatient claims data, 2022–2024)'),
      kw: ['รายละเอียดข้อมูล', 'ข้อมูลที่เปิดเผย', 'ข้อมูลที่ใช้', 'ข้อมูลที่ขอ', 'ข้อมูลเกี่ยวกับ', '+ใช้ข้อมูล', '+ขอข้อมูล', '+ขอใช้ข้อมูล', '+เข้าถึงข้อมูล', 'information about', 'information on', 'information', 'data on', 'data about', 'access to', 'needs', 'uses'], group: 2 },
    { key: 'objective', type: 'textarea', ask: 8, label: L('วัตถุประสงค์ในการใช้ข้อมูล', 'Purpose of use'), chip: L('ระบุวัตถุประสงค์ในการขอข้อมูล', 'Objective'),
      ph: L('วิเคราะห์ความคุ้มค่าของวัคซีน HPV', 'cost-effectiveness analysis of HPV vaccination'), q: L('นำข้อมูลไปใช้เพื่ออะไรครับ', 'What is the information used for?'),
      kw: ['วัตถุประสงค์ในการขอข้อมูล', 'วัตถุประสงค์ในการใช้ข้อมูล', 'วัตถุประสงค์', 'เพื่อใช้ใน', 'เพื่อใช้', 'เพื่อ', 'for the purpose of', 'purpose of use', 'purpose', 'objective', 'in order to'], group: 2 },
    { key: 'sig_recipient', type: 'text', mirror: 'recipient_name', hide: true, label: L('ชื่อผู้รับข้อมูล (ใต้ลายเซ็น)', 'Recipient name (signature)'), chip: L('ระบุชื่อผู้รับข้อมูล', 'Recipient name') },
    { key: 'sig_witness1', type: 'text', label: L('ชื่อพยานคนที่ 1', 'Witness 1 name'), chip: L('ระบุชื่อพยาน', 'Witness name'),
      kw: ['พยานคนที่ 1', 'พยานคนที่1', 'พยานคนแรก', 'พยาน 1', 'พยาน1', 'witness 1', 'first witness'], group: 3 },
    { key: 'sig_witness2', type: 'text', label: L('ชื่อพยานคนที่ 2', 'Witness 2 name'), chip: L('ระบุชื่อพยาน', 'Witness name'),
      kw: ['พยานคนที่ 2', 'พยานคนที่2', 'พยานคนที่สอง', 'พยาน 2', 'พยาน2', 'witness 2', 'second witness'], group: 3 },
  ];
  const NDA_GROUPS = [L('ข้อมูลสัญญา', 'Agreement'), L('ผู้รับข้อมูล', 'Recipient'), L('ขอบเขตข้อมูล', 'Information scope'), L('พยาน (ไม่บังคับ)', 'Witnesses (optional)')];
  const DISCLOSER = { th: 'นางสาววรรณฤดี อิสรานุวัฒน์ชัย', en: 'Wanrudee Isaranuwatchai' };
  const ndaTemplate = lang => ({
    id: `nda-${lang}`, kind: 'docx', cat: 'contract', docLang: lang, dateFmt: lang === 'th' ? 'th-long' : 'en-long', builtin: true, witnessKw: true,
    file: window.NDA_TEMPLATES[lang].file,
    title: lang === 'th' ? L('สัญญาให้เก็บรักษาข้อมูลไว้เป็นความลับ (NDA) — ฉบับภาษาไทย', 'Non-Disclosure Agreement — Thai version') : L('Non-Disclosure Agreement — ฉบับภาษาอังกฤษ', 'Non-Disclosure Agreement — English version'),
    short: lang === 'th' ? L('สัญญา NDA (ฉบับไทย)', 'NDA (Thai)') : L('สัญญา NDA (ฉบับอังกฤษ)', 'NDA (English)'),
    desc: lang === 'th' ? L('สำหรับพนักงาน/ผู้รับจ้างของ HITAP ที่ต้องเข้าถึงข้อมูลอันเป็นความลับเพื่อใช้ในโครงการวิจัย', 'For HITAP staff/contractors who need confidential information for a research project')
      : L('ฉบับภาษาอังกฤษ สำหรับผู้รับข้อมูลชาวต่างชาติหรือเอกสารที่ต้องใช้ภาษาอังกฤษ', 'English version for foreign recipients or English-language files'),
    fields: NDA_FIELDS, groups: NDA_GROUPS,
    signers: [
      { role: 'discloser', label: L('ผู้ให้ข้อมูล', 'Discloser'), fixedName: DISCLOSER[lang] },
      { role: 'recipient', label: L('ผู้รับข้อมูล', 'Recipient'), nameKey: 'recipient_name' },
      { role: 'witness1', label: L('พยาน 1', 'Witness 1'), nameKey: 'sig_witness1' },
      { role: 'witness2', label: L('พยาน 2', 'Witness 2'), nameKey: 'sig_witness2' },
    ],
    loadData: async () => ({ file: window.NDA_TEMPLATES[lang].template, labels: window.NDA_TEMPLATES[lang].labels }),
    original: () => new Blob([U.b64ToBytes(window.NDA_TEMPLATES[lang].original)], { type: HD.engines.docx.mime }),
  });

  /* ================= Catalog ================= */
  const TAG_KW = {
    name: ['ชื่อ-นามสกุล', 'ชื่อ-สกุล', 'ชื่อนามสกุล', 'ชื่อ', 'full name', 'name'], position: ['ตำแหน่ง', 'position', 'job title'],
    unit: ['ฝ่าย', 'unit', 'department'], employee_code: ['รหัสพนักงาน', 'employee code', 'employee id'], phone: ['เบอร์โทรศัพท์', 'เบอร์โทร', 'โทรศัพท์', 'เบอร์', 'tel', 'phone', 'mobile'],
    email: ['อีเมล', 'อีเมล์', 'email', 'e-mail'], id_card: ['เลขบัตรประชาชน', 'บัตรประชาชน', 'เลขบัตร', 'หนังสือเดินทาง', 'passport', 'national id', 'id card'],
    address: ['ที่อยู่', 'address'], organization: ['หน่วยงาน', 'องค์กร', 'สังกัด', 'organization', 'organisation'],
    project_code: ['รหัสโครงการ', 'project code'], activity_code: ['รหัสกิจกรรม', 'activity code'],
  };
  const KIND_LABEL = { docx: L('ไฟล์ Word', 'Word'), pdf: L('PDF', 'PDF'), xlsx: L('ไฟล์ Excel', 'Excel') };
  const CATS = [{ id: 'all', th: 'ทั้งหมด', en: 'All' }, ...(window.HD_CATEGORIES || [{ id: 'contract', th: 'สัญญา/NDA', en: 'Contracts / NDA' }])];

  function normalize(t) {
    const tpl = { ...t };
    tpl.engine = HD.engines[tpl.kind];
    tpl.fields = tpl.fields.map(f => ({ ...f }));
    tpl.byKey = Object.fromEntries(tpl.fields.map(f => [f.key, f]));
    tpl.ask = tpl.fields.filter(f => f.ask && !f.hide).sort((a, b) => a.ask - b.ask).map(f => f.key);
    tpl.signers = tpl.signers || [];
    tpl.short = tpl.short || tpl.title;
    if (!tpl.loadData) tpl.loadData = () => U.loadTemplateData(tpl);
    if (!tpl.original && !tpl.originalHref) tpl.originalHref = `templates/${tpl.file}`;
    for (const f of tpl.fields) {
      if (f.tag === 'name') f.nameField = true;
      if (f.tag === 'id_card') f.idField = true;
    }
    tpl.kwList = buildKeywords(tpl);
    return tpl;
  }
  // TEMPLATES holds every known template (built-in + uploaded by admins); hidden ones stay resolvable
  // for documents that already use them but are left out of the catalog.
  const TEMPLATES = [ndaTemplate('th'), ndaTemplate('en'), ...(window.HD_CATALOG || [])].map(normalize);
  const TPL = id => TEMPLATES.find(t => t.id === id);
  const VISIBLE = () => TEMPLATES.filter(t => !t.hidden);
  const HD_BUILTIN_IDS = new Set(TEMPLATES.map(t => t.id));
  let remoteLoaded = false;
  const visibleFields = tpl => tpl.fields.filter(f => !f.hide && !f.table && !f.cont && f.type !== 'calc');

  /* ================= Values ================= */
  function rawOf(d, key, depth = 0) {
    const tpl = TPL(d.tpl), f = tpl.byKey[key];
    let v = d.values[key];
    if ((v == null || v === '') && f && f.mirror && depth < 3) v = rawOf(d, f.mirror, depth + 1);
    return v;
  }
  function valueOf(d, key) {
    const tpl = TPL(d.tpl), f = tpl.byKey[key] || {};
    const v = rawOf(d, key);
    if (v == null || v === '' || v === false) return '';
    if (f.type === 'date') return U.fmtDate(v, tpl.dateFmt);
    if (f.type === 'money' && tpl.kind === 'docx') return U.money(v);
    if (f.type === 'check' || f.type === 'yesno') return String(v);
    return oneLine(String(v));
  }
  const labelOf = (tpl, key) => { const f = tpl.byKey[key]; return f ? f.label[uiLang()] : key; };
  const filled = (d, key) => { const v = rawOf(d, key); return v != null && v !== '' && v !== false; };
  const missingFields = d => TPL(d.tpl).ask.filter(k => !filled(d, k));

  function defaults(tpl) {
    const vals = {}, today = U.isoParts(U.todayISO()), profile = store.get(LS.profile, {});
    for (const f of tpl.fields) {
      if (f.value != null && f.value !== '') vals[f.key] = f.value;
      if (f.def === 'today' && f.type === 'date') vals[f.key] = U.todayISO();
      if (f.type === 'choice' && f.def && f.options) {
        const want = f.def === 'day' ? [String(today.d)] : f.def === 'month' ? [U.EN_MONTHS[today.m - 1], U.TH_MONTHS[today.m - 1]] : [String(today.y), String(today.y + 543)];
        const opt = f.options.find(o => want.includes(String(o.v).trim()) || want.includes(String(o.t).trim()));
        if (opt) vals[f.key] = opt.v;
      }
      if (f.tag && profile[f.tag]) vals[f.key] = profile[f.tag];
    }
    return vals;
  }
  function rememberProfile(d, keys) {
    const tpl = TPL(d.tpl), profile = store.get(LS.profile, {});
    let changed = false;
    for (const k of keys) { const f = tpl.byKey[k]; if (f && f.tag && d.values[k]) { profile[f.tag] = d.values[k]; changed = true; } }
    if (changed) { store.set(LS.profile, profile); if (HD.cloud) HD.cloud.saveProfile(profile); }
  }

  /* ================= Free-text extraction (rule based, free) ================= */
  function stripLabel(s) { return s.replace(/^(A\d|งน\.\d+|Form \d+):\s*/i, '').replace(/\s*\((?:บาท|baht|ต่อ|รายละเอียด|details)\)\s*$/i, '').replace(/\s+—.*$/, '').trim().toLowerCase(); }
  function buildKeywords(tpl) {
    const out = [];
    for (const f of tpl.fields) {
      if (f.hide || f.table || f.cont || f.type === 'calc' || f.type === 'check') continue;
      const words = new Set();
      for (const w of f.kw || []) words.add(w);
      if (!f.kw) {
        for (const l of [f.label.th, f.label.en]) { const s = stripLabel(l); if (s.length >= 3) words.add(s); }
        for (const w of TAG_KW[f.tag] || []) words.add(w);
      }
      // 'ขอ>ลา' matches "ขอลา…" and keeps "ลา" at the start of the value
      for (const w of words) { const [a, keep = ''] = w.replace(/^\+/, '').split('>'); out.push({ key: f.key, kw: (a + keep).toLowerCase(), keep, keepData: w.startsWith('+'), tag: f.tag }); }
    }
    if (tpl.witnessKw) out.push(...['พยาน', 'witnesses', 'witness'].map(kw => ({ key: '_witnesses', kw })));
    return out;
  }
  const isLatin = c => /[a-z0-9]/i.test(c || '');
  function findKeywords(tpl, text) {
    const low = text.toLowerCase(), hits = [];
    for (const k of tpl.kwList) {
      let i = low.indexOf(k.kw);
      while (i >= 0) {
        const latin = /^[a-z]/.test(k.kw);
        if (!latin || (!isLatin(low[i - 1]) && !isLatin(low[i + k.kw.length]))) hits.push({ ...k, start: i, end: i + k.kw.length });
        i = low.indexOf(k.kw, i + 1);
      }
    }
    hits.sort((a, b) => a.start - b.start || b.end - a.end);
    const out = [];
    for (const h of hits) {
      const last = out[out.length - 1];
      if (!last || h.start >= last.end) out.push({ ...h, keys: [h.key] });
      else if (last.start === h.start && last.end === h.end && !last.keys.includes(h.key)) last.keys.push(h.key);
    }
    return out;
  }
  function clean(v) {
    let s = (v || '').replace(/\s+/g, ' ').trim(), prev;
    do {
      prev = s;
      s = s.replace(/^[\s:=：,;\-–—"“”']+/, '').replace(/^(คือ|เป็น|ว่า|ได้แก่|ของ|is|are|was|of|as)\s+/i, '').replace(/^(คือ|เป็น|ว่า|ได้แก่)/, '')
        .replace(/[\s,;:"“”']+$/, '').replace(/\s*(และ|ครับ|ค่ะ|คะ|นะ|จ้า|ด้วย|and|for)$/i, '').replace(/[.。]$/, '');
    } while (s !== prev);
    return s.trim();
  }
  const NAME_TH = /((?:นางสาว|นาง|นาย|น\.ส\.|ดร\.|ผศ\.\s?ดร\.|รศ\.\s?ดร\.|ศ\.\s?ดร\.|ผศ\.|รศ\.|ศ\.)\s?[฀-๿]+\s+[฀-๿]+)/;
  const NAME_EN = /\b((?:Mr|Ms|Mrs|Miss|Dr|Prof)\.?\s+[A-Z][\w'-]+(?:\s+[A-Z][\w'-]+){1,2})/;
  const ID_RE = /(?<!\d)(\d[\s-]?\d{4}[\s-]?\d{5}[\s-]?\d{2}[\s-]?\d)(?!\d)/;

  /** Pull field values out of free text; `pending` = the field the assistant just asked for. */
  function extract(tpl, text, pending) {
    // "โครงการ การพัฒนา…สำหรับโครงการ…": a later hit for the same field belongs to the value
    const hits = findKeywords(tpl, text).filter((h, i, all) => !(i && all[i - 1].keys.join() === h.keys.join())), vals = {}, errors = [];
    const prefix = text.slice(0, hits.length ? hits[0].start : text.length);
    hits.forEach((h, i) => {
      let v = clean(text.slice(h.end, i + 1 < hits.length ? hits[i + 1].start : text.length).split('\n')[0]);
      if (h.keepData && v) v = (tpl.docLang === 'th' ? 'ข้อมูล' : '') + v;
      if (h.keep && v) v = h.keep + v;
      if (!v) return;
      for (const k of h.keys) if (!vals[k]) vals[k] = v;
    });
    if (pending && clean(prefix) && !vals[pending]) vals[pending] = clean(prefix);
    if (vals._witnesses) {
      const parts = vals._witnesses.split(/\s*(?:,|และ|\band\b|&)\s*/).map(clean).filter(Boolean);
      if (parts[0] && !vals.sig_witness1) vals.sig_witness1 = parts[0];
      if (parts[1] && !vals.sig_witness2) vals.sig_witness2 = parts[1];
      delete vals._witnesses;
    }
    const nameKeys = tpl.fields.filter(f => f.nameField && !f.hide).map(f => f.key);
    const idKeys = tpl.fields.filter(f => f.idField && !f.hide).map(f => f.key);
    const plain = hits.filter(h => /witness|^sig_/.test(h.key)).reduce((t, h) => t.replace(text.slice(h.start), ' '), text);
    for (const k of nameKeys) {
      if (!vals[k]) { const m = NAME_TH.exec(plain) || NAME_EN.exec(plain); if (m && (!pending || pending === k || !nameKeys.includes(pending))) vals[k] = m[1].trim(); }
      else if (pending !== k) { const m = NAME_TH.exec(vals[k]) || NAME_EN.exec(vals[k]); if (m) vals[k] = m[1].trim(); }
    }
    for (const k of idKeys) {
      if (vals[k]) { const m = ID_RE.exec(vals[k]) || /[A-Z]{0,3}\d[A-Z0-9-]{4,}/i.exec(vals[k]); if (m) vals[k] = (m[1] || m[0]).trim(); }
      else { const m = ID_RE.exec(plain); if (m) vals[k] = m[1]; }
    }
    for (const [k, v] of Object.entries(vals)) {
      const f = tpl.byKey[k];
      if (!f) { delete vals[k]; continue; }
      if (f.type === 'date') { const iso = U.parseDate(v); if (iso) vals[k] = iso; else { errors.push(k); delete vals[k]; } }
      else if (f.type === 'money' || f.type === 'number') { const n = U.num(v.split(' ')[0]); if (n != null) vals[k] = String(n); }
      else if (f.type === 'yesno') { const yn = yesNo(v); if (yn) vals[k] = yn; else delete vals[k]; }
      else if (f.type === 'choice') { const o = (f.options || []).find(o => String(o.t).trim().toLowerCase() === v.toLowerCase() || String(o.v).trim() === v); if (o) vals[k] = o.v; else delete vals[k]; }
    }
    return { vals, errors };
  }
  // \b doesn't work next to Thai letters, so Thai answers are matched as prefixes
  const YN_PREFIX = /^(ไม่ใช่|ไม่มี|ไม่|ใช่|มี|(?:no|none|nope|n|yes|yeah|yep|y)(?![a-z]))/i;
  const yesNo = s => { const m = YN_PREFIX.exec(s.trim()); return !m ? null : /^(ไม่|no|n(?![a-z])|none|nope)/i.test(m[1]) ? 'no' : 'yes'; };

  /* ================= Template detection on the home composer ================= */
  const hasThai = t => /[฀-๿]/.test(t);
  function detectTemplate(text, langPref) {
    const lang = langPref !== 'auto' ? langPref : /ภาษาอังกฤษ|ฉบับอังกฤษ|อังกฤษ|english/i.test(text) ? 'en' : /ภาษาไทย|ฉบับไทย|thai/i.test(text) ? 'th' : hasThai(text) || !text.trim() ? 'th' : 'en';
    const t = text.toLowerCase();
    const rules = [
      [/สสส|hpie|ใบรับรองการจ่ายเงิน|งน\.?\s?(1|11)\b/, () => 'hpie'],
      [/ยืมเงิน|ทดรองจ่าย|borrow|คืนเงินยืม|return money|cash advance/, () => 'borrow'],
      [/\bcoi\b|ผลประโยชน์ทับซ้อน|ขัดกันของผลประโยชน์|conflict of interest/, () => `coi-${lang}`],
      [/ใบสำคัญรับเงิน|receipt|voucher|ค่าตอบแทน|เบี้ยเลี้ยง|เบิกค่าเดินทาง/, () => (lang === 'th' ? 'receipt-th-pcht' : /pcht/.test(t) ? 'receipt-en-pcht' : 'receipt-en')],
      [/\bar1\b|รายงานผล|ดูงาน|ฝึกอบรม|study visit|training report/, () => 'ar1'],
      [/\ba1\b|\ba2\b|ประชุมวิชาการ|conference|ส่งผลงาน/, () => 'a1a2'],
      [/^ลา|ขอลา|ใบลา|การลา|วันลา|ลาพัก|ลาป่วย|ลากิจ|ลาคลอด|ลาบวช|ลาอุปสมบท|ลาไป|\bleave\b|มอบหมายงาน/, () => (/fm-hr|hr-004/.test(t) ? 'leave-hr' : 'leave-adm')],
      [/nda|non-?disclosure|เก็บรักษาข้อมูล|ความลับ|สัญญา/, () => `nda-${lang}`],
    ];
    for (const [re, pick] of rules) if (re.test(t)) { const id = pick(); if (TPL(id) && !TPL(id).hidden) return id; }
    // uploaded templates: match on their title words
    const custom = VISIBLE().find(x => !x.builtin && x.cat !== 'contract' && [x.title.th, x.title.en].some(ti => ti && ti.length > 3 && t.includes(ti.toLowerCase().slice(0, 12))));
    if (custom) return custom.id;
    return null;
  }

  /* ================= Documents (persisted) ================= */
  // documents of uploaded templates are kept even before the catalog arrives from Supabase
  let docs = store.get(LS.docs, []).map(d => (d.tpl ? d : { ...d, tpl: `nda-${d.lang || 'th'}` }));
  let current = null;
  const saveDocs = () => {
    docs.sort((a, b) => b.updated - a.updated);
    if (!store.set(LS.docs, docs.slice(0, 40))) toast(tr('พื้นที่จัดเก็บในเบราว์เซอร์เต็ม — ลบเอกสารเก่าหรือลายเซ็นบางส่วน', 'Browser storage is full'));
    if (HD.cloud && current) HD.cloud.saveDoc(current);
  };
  function touch() { if (!current) return; current.updated = Date.now(); saveDocs(); renderHistory(); }

  /* ================= Engine context ================= */
  let ctxCache = null;
  async function ctxFor(d) {
    const tpl = TPL(d.tpl);
    if (!ctxCache || ctxCache.doc !== d) {
      const data = await tpl.loadData();
      ctxCache = {
        tpl, doc: d, data,
        field: k => tpl.byKey[k], raw: k => rawOf(d, k), value: k => valueOf(d, k), label: k => labelOf(tpl, k),
        chipLabel: k => { const f = tpl.byKey[k]; return f ? (f.chip ? f.chip[tpl.docLang] : f.label[tpl.docLang === 'en' ? 'en' : 'th'].replace(/^(A\d|งน\.\d+):\s*/, '')) : k; },
        t: tr,
      };
    }
    return ctxCache;
  }

  /* ================= Preview ================= */
  let renderSeq = 0, renderTimer = null;
  const schedulePreview = (ms = 250) => { clearTimeout(renderTimer); renderTimer = setTimeout(() => renderPreview(), ms); };
  async function renderPreview(mode = 'preview') {
    if (!current) return;
    const seq = ++renderSeq, box = $('#preview'), scroll = $('#paperScroll'), tpl = TPL(current.tpl);
    if (!box.firstChild) box.innerHTML = `<div class="preview-msg">${tr('กำลังจัดหน้าเอกสาร…', 'Laying out the document…')}</div>`;
    try {
      const ctx = await ctxFor(current);
      const node = await tpl.engine.render(ctx, null, mode);
      if (seq !== renderSeq) return;
      box.dataset.doclang = tpl.docLang;
      box.dataset.kind = tpl.kind;
      box.classList.toggle('nda-doc', !!tpl.builtin);
      const top = scroll.scrollTop;
      box.replaceChildren(...(tpl.kind === 'docx' ? node.childNodes : [node]));
      fitPreview();
      scroll.scrollTop = top;
    } catch (err) {
      console.error(err);
      box.innerHTML = `<div class="preview-msg">${tr('แสดงตัวอย่างไม่ได้', 'Preview unavailable')}: ${htmlEsc(err.message)}</div>`;
    }
  }
  /** Inline engines (PDF/Excel) refresh their inputs and calculated values without a full re-render. */
  async function refreshInline() {
    const tpl = TPL(current.tpl);
    if (tpl.engine.inline && tpl.engine.update && $('#preview').firstChild) tpl.engine.update(await ctxFor(current), $('#preview'));
    else schedulePreview(60);
  }
  function fitPreview() {
    const page = $('#preview .hd-page, #preview section.docx');
    if (!page) return;
    const avail = $('#paperScroll').clientWidth - 32;
    $('#preview').style.zoom = Math.min(1, avail / page.offsetWidth);
  }

  /* ================= Field editing ================= */
  function setValue(key, v, { render = true } = {}) {
    current.values[key] = v;
    rememberProfile(current, [key]);
    touch(); updateChrome();
    const inp = $(`#ndaForm [name="${CSS.escape(key)}"]`);
    if (inp && inp !== document.activeElement) { if (inp.type === 'checkbox') inp.checked = !!v; else inp.value = v ?? ''; }
    updateFieldState(key);
    if (render) TPL(current.tpl).engine.inline ? refreshInline() : schedulePreview(60);
  }
  function afterManualEdit(key) {
    if (current.pending === key && filled(current, key)) { current.pending = null; botNext(false); renderThread(); }
  }

  let popKey = null;
  function openPop(key, anchor) {
    const tpl = TPL(current.tpl);
    const f = tpl.byKey[key];
    if (f && f.mirror && !current.values[key] && f.hide) key = f.mirror;
    popKey = key;
    const fld = tpl.byKey[key], v = current.values[key] ?? rawOf(current, key) ?? '';
    $('#popLabel').textContent = fld.label[uiLang()];
    const ph = fld.ph ? fld.ph[tpl.docLang] : '';
    $('#popInput').innerHTML = fld.type === 'textarea'
      ? `<textarea id="popField" rows="3" placeholder="${htmlEsc(ph)}">${htmlEsc(v)}</textarea>`
      : `<input id="popField" type="${fld.type === 'date' ? 'date' : 'text'}" value="${htmlEsc(v)}" placeholder="${htmlEsc(ph)}">`;
    const pop = $('#pop');
    pop.hidden = false;
    const r = anchor.getBoundingClientRect(), pw = pop.offsetWidth, ph2 = pop.offsetHeight;
    const left = Math.min(Math.max(12, r.left), window.innerWidth - pw - 12);
    let top = r.bottom + 8;
    if (top + ph2 > window.innerHeight - 12) top = Math.max(12, r.top - ph2 - 8);
    pop.style.left = left + 'px'; pop.style.top = top + 'px';
    const input = $('#popField');
    input.focus(); if (input.select && fld.type !== 'date') input.select();
    updatePopNote();
    input.addEventListener('input', updatePopNote);
    input.addEventListener('keydown', e => {
      if (e.key === 'Escape') closePop();
      if (e.key === 'Enter' && (fld.type !== 'textarea' || e.metaKey || e.ctrlKey)) { e.preventDefault(); savePop(); }
    });
  }
  function noteFor(key, v) {
    const tpl = TPL(current.tpl), f = tpl.byKey[key];
    if (f.type === 'date' && v) return tr('ในเอกสาร: ', 'In document: ') + U.fmtDate(v, tpl.dateFmt);
    if (f.idField) { const s = U.thaiIdStatus(v); if (s === 'bad') return tr('⚠ เลขบัตรไม่ผ่านการตรวจเลขหลักสุดท้าย', '⚠ Thai ID checksum failed'); if (s === 'ok') return tr('✓ เลขบัตรถูกต้องตามรูปแบบ', '✓ Valid Thai ID'); }
    if (f.note) return f.note[uiLang()];
    if (f.nameField && v && tpl.signers.some(s => s.nameKey === key)) return tr('ชื่อนี้จะใส่ใต้ช่องลงนามด้วย', 'Also printed under the signature');
    return '';
  }
  function updatePopNote() { $('#popNote').textContent = noteFor(popKey, $('#popField').value); }
  function closePop() { $('#pop').hidden = true; popKey = null; }
  function savePop() { const key = popKey; if (!key) return; setValue(key, $('#popField').value.trim()); closePop(); afterManualEdit(key); }

  /* ================= Chat engine ================= */
  const BOT_ICON = '<svg class="bot-icon" viewBox="0 0 24 24" fill="currentColor"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" opacity=".9"/><path d="M8 12h8M8 16h5" stroke="#fff" stroke-width="1.8" stroke-linecap="round"/><path d="M19.5 1l.8 2 2 .8-2 .7-.8 2-.7-2-2-.7 2-.8z"/></svg>';
  const bot = msg => current.msgs.push({ r: 'b', ...msg });
  const shown = (tpl, k, v) => { const f = tpl.byKey[k]; return f.type === 'date' ? U.fmtDate(v, tpl.dateFmt) : f.type === 'yesno' ? (v === 'yes' ? tr('ใช่', 'Yes') : tr('ไม่ใช่', 'No')) : f.type === 'money' ? U.money(v) : v; };
  function questionFor(tpl, key) {
    const f = tpl.byKey[key];
    if (f.q) return f.q[uiLang()];
    const lab = stripLabel(f.label[uiLang()]).replace(/ — .*$/, '');
    const base = f.label[uiLang()].replace(/^(A\d|งน\.\d+|Form \d+):\s*/, '').replace(/ — (ใช่\/ไม่ใช่|yes\/no)$/, '');
    if (f.type === 'yesno') return tr(`ข้อ “${base}” — ใช่หรือไม่ใช่ครับ? (ถ้าใช่ พิมพ์รายละเอียดต่อท้ายได้ เช่น “ใช่ ได้รับทุนจาก…”)`, `“${base}” — yes or no? (if yes, add details, e.g. “yes, funded by …”)`);
    if (f.type === 'date') return tr(`${base} — วันที่เท่าไหร่ครับ (เช่น 1 ต.ค. 2569)`, `${base} — which date? (e.g. 1 Oct 2026)`);
    return tr(`${base} คืออะไรครับ`, `What is the ${lab || base}?`);
  }
  function botNext(withList) {
    const tpl = TPL(current.tpl);
    const miss = missingFields(current).filter(k => !current.skipped.includes(k));
    if (miss.length) {
      current.pending = miss[0];
      const q = questionFor(tpl, miss[0]);
      if (withList && miss.length > 1) {
        const list = miss.slice(0, 10).map(k => labelOf(tpl, k).replace(/ — (ใช่\/ไม่ใช่|yes\/no)$/, ''));
        if (miss.length > 10) list.push(tr(`…และอีก ${miss.length - 10} ข้อ`, `…and ${miss.length - 10} more`));
        bot({ lead: tr('ขอข้อมูลเพิ่มอีกนิดครับ:', 'A few more details, please:'), list, t: tr(`เริ่มจาก — ${q}`, `First — ${q}`) });
      } else bot({ t: miss.length > 1 ? tr(`ต่อไปครับ (เหลือ ${miss.length} ข้อ) — ${q}`, `Next (${miss.length} left) — ${q}`) : q });
      return;
    }
    current.pending = null;
    if (!missingFields(current).length && !current.doneSaid) {
      current.doneSaid = true;
      const more = tpl.builtin
        ? tr('\nถ้าต้องการใส่ชื่อพยาน พิมพ์ เช่น “พยาน นายสมชาย ใจดี และ นางสาวสมศรี รักงาน”\nใส่เลขที่สัญญาได้โดยพิมพ์ เช่น “เลขที่สัญญา HITAP-NDA 001/2569” หรือคลิกที่หัวกระดาษ', '\nTo add witnesses, type e.g. “witness Mr. John Smith and Ms. Mary Lee”.\nAdd a reference number with e.g. “reference number HITAP-NDA 001/2026”, or click the page header.')
        : tr('\nช่องอื่น ๆ (เช่น ตารางค่าใช้จ่าย) คลิกกรอกในเอกสารด้านขวาได้เลย ระบบคำนวณยอดรวมให้', '\nOther fields (e.g. expense tables) can be filled right in the document — totals are calculated for you.');
      bot({ lead: tr('ข้อมูลหลักครบแล้วครับ ✓', 'The key details are complete ✓'),
        t: tr(`ตรวจทานเอกสารด้านขวา${tpl.signers.length ? ' แล้วกด “ลงนาม” เพื่อเซ็นชื่อ' : ''} หรือกด “${tpl.engine.ext.toUpperCase()}” เพื่อดาวน์โหลดได้เลยครับ`, `Review the document on the right${tpl.signers.length ? ', press “Sign” to sign' : ''}, or press “${tpl.engine.ext.toUpperCase()}” to download.`) + more });
    } else if (missingFields(current).length && !current.skipNoted) {
      current.skipNoted = true;
      bot({ t: tr('ยังเหลือช่องที่ข้ามไว้ — คลิกช่องในเอกสารเพื่อกรอกได้ทุกเมื่อครับ', 'Some fields were skipped — click them in the document any time.') });
    }
  }

  function startDoc(text, tplId) {
    const tpl = TPL(tplId);
    const d = { id: 'd' + Date.now().toString(36), tpl: tpl.id, values: defaults(tpl), signs: {}, msgs: [], pending: null, skipped: [],
      title: text ? oneLine(text).slice(0, 60) : tpl.short[uiLang()], created: Date.now(), updated: Date.now() };
    current = d; ctxCache = null; docs.unshift(d);
    let got = {};
    if (text) {
      d.msgs.push({ r: 'u', t: text });
      got = extract(tpl, text, null).vals;
      Object.assign(d.values, got);
      rememberProfile(d, Object.keys(got));
    }
    bot({ lead: tr('ร่างเอกสารเสร็จแล้วครับ', 'Your draft is ready'),
      t: tr('ถ้าต้องการแก้ไขหรือเพิ่มเติมข้อมูล พิมพ์บอกได้เลยครับ หรือคลิกช่องในเอกสารเพื่อแก้โดยตรง', 'Type to add or change details, or click any field in the document.'), card: true });
    const prof = Object.keys(d.values).filter(k => tpl.byKey[k] && tpl.byKey[k].tag && !got[k] && !tpl.byKey[k].hide);
    const kv = Object.entries(got).filter(([k]) => tpl.byKey[k]).map(([k, v]) => [labelOf(tpl, k), shown(tpl, k, v)]);
    if (kv.length) bot({ t: tr(`กรอกจากข้อความของคุณให้แล้ว ${kv.length} ช่อง:`, `Filled ${kv.length} field(s) from your message:`), kv });
    if (prof.length) bot({ t: tr('ใช้ข้อมูลของคุณจากเอกสารก่อนหน้า:', 'Reused your details from earlier documents:'), kv: prof.map(k => [labelOf(tpl, k), shown(tpl, k, d.values[k])]) });
    botNext(true);
    saveDocs();
    location.hash = `#/doc/${d.id}`;
  }

  function handleUser(text) {
    const t = text.trim();
    if (!t || !current) return;
    const tpl = TPL(current.tpl);
    current.msgs.push({ r: 'u', t });
    const p = current.pending, pf = p && tpl.byKey[p];
    if (p && /^(ข้าม|ข้ามไป|ไม่ทราบ|ยังไม่มี|ไว้ก่อน|skip|later|n\/?a)$/i.test(t)) {
      current.skipped.push(p);
      bot({ t: tr(`ข้าม “${labelOf(tpl, p)}” ไว้ก่อนนะครับ`, `Skipped “${labelOf(tpl, p)}” for now.`) });
      botNext(false);
    } else if (/^(ลงนาม|เซ็น|เซ็นชื่อ|ลงชื่อ|sign)/i.test(t) && t.length < 20 && tpl.signers.length) {
      bot({ t: tr('เปิดหน้าต่างลงนามให้แล้วครับ — วาด พิมพ์ หรืออัปโหลดลายเซ็นได้', 'Opened the signing window — draw, type or upload a signature.') });
      openSign();
    } else if (/^(ดาวน์โหลด|โหลด|download|docx|pdf|xlsx)/i.test(t) && t.length < 30) {
      bot({ t: tr('กำลังดาวน์โหลดไฟล์ให้ครับ', 'Downloading the file.') });
      download();
    } else if (/^(ok|โอเค|ตกลง|เรียบร้อย|ขอบคุณ|thanks?|thank you)/i.test(t) && !p) {
      bot({ t: tr('ยินดีครับ ถ้าต้องการแก้ไขอะไรเพิ่ม พิมพ์บอกได้เลย', 'You\'re welcome — tell me if anything needs changing.') });
    } else {
      let vals = {}, errors = [];
      if (pf && pf.type === 'yesno') {
        const yn = yesNo(t);
        if (yn) {
          vals[p] = yn;
          const detailKey = p.replace(/^q_/, ''), rest = clean(t.trim().replace(YN_PREFIX, ''));
          if (rest && tpl.byKey[detailKey]) vals[detailKey] = rest;
        } else ({ vals, errors } = extract(tpl, t, null));
      } else ({ vals, errors } = extract(tpl, t, p));
      const keys = Object.keys(vals);
      keys.forEach(k => { current.values[k] = vals[k]; current.skipped = current.skipped.filter(x => x !== k); });
      rememberProfile(current, keys);
      if (keys.length) {
        bot({ t: tr('บันทึกแล้วครับ ✓', 'Saved ✓'), kv: keys.map(k => [labelOf(tpl, k), shown(tpl, k, vals[k])]) });
        const idk = keys.find(k => tpl.byKey[k].idField);
        if (idk && U.thaiIdStatus(vals[idk]) === 'bad') bot({ t: tr('⚠ เลขบัตรประชาชนนี้ไม่ผ่านการตรวจเลขหลักสุดท้าย ลองตรวจสอบอีกครั้งนะครับ', '⚠ This Thai ID fails the checksum — please double-check.') });
      }
      if (errors.length) bot({ t: tr(`อ่าน${errors.map(k => labelOf(tpl, k)).join(', ')}ไม่ออกครับ ลองพิมพ์แบบ “1 ต.ค. 2569” หรือ “1/10/2569”`, `Couldn't read the ${errors.map(k => labelOf(tpl, k)).join(', ')} — try “1 Oct 2026” or “2026-10-01”.`) });
      if (!keys.length && !errors.length) {
        bot({ t: pf && pf.type === 'yesno' ? tr('ตอบ “ใช่” หรือ “ไม่ใช่” ได้เลยครับ', 'Please answer “yes” or “no”.')
          : tr('ขอโทษครับ ยังจับข้อมูลจากข้อความนี้ไม่ได้ ลองระบุหัวข้อด้วย เช่น “ตำแหน่ง นักวิจัย” หรือคลิกช่องในเอกสารเพื่อกรอกเอง', 'Sorry, I couldn\'t pick anything up. Try naming the field, e.g. “position Researcher”, or click the field in the document.') });
      } else if (!errors.includes(p)) botNext(false);
    }
    touch(); updateChrome(); renderThread();
    tpl.engine.inline ? refreshInline() : schedulePreview(30);
    if ($('#editDrawer').classList.contains('open')) renderForm();
  }

  function renderThread() {
    const th = $('#thread'), tpl = TPL(current.tpl);
    const head = (filled(current, tpl.signers[1] && tpl.signers[1].nameKey) && valueOf(current, tpl.signers[1].nameKey))
      || (tpl.fields.find(f => f.nameField && filled(current, f.key)) && valueOf(current, tpl.fields.find(f => f.nameField && filled(current, f.key)).key)) || tpl.title[uiLang()];
    th.innerHTML = current.msgs.map(m => {
      if (m.r === 'u') return `<div class="msg-user">${htmlEsc(m.t)}</div>`;
      let body = '';
      if (m.lead) body += `<p class="lead">${htmlEsc(m.lead)}</p>`;
      if (m.t && !m.list) body += `<p class="${m.lead ? 'soft' : ''}">${htmlEsc(m.t)}</p>`;
      if (m.card) body += `<button type="button" class="doc-card" data-open-doc><span class="ic"><svg viewBox="0 0 24 24" width="22" height="22" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/><path d="M14 3v5h5M9 13h6M9 17h4"/></svg></span><span class="t"><small>${tr('กำลังเปิดอยู่', 'Open now')} · ${htmlEsc(tpl.short[uiLang()].slice(0, 48))}</small><b>${htmlEsc(head)}</b></span><svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2"><path d="m9 6 6 6-6 6"/></svg></button>`;
      if (m.list) body += `<ul>${m.list.map(x => `<li>${htmlEsc(x)}</li>`).join('')}</ul>${m.t ? `<p>${htmlEsc(m.t)}</p>` : ''}`;
      if (m.kv) body += `<dl class="kv">${m.kv.map(([k, v]) => `<dt>${htmlEsc(k)}</dt><dd>${htmlEsc(v)}</dd>`).join('')}</dl>`;
      return `<div class="msg-bot">${BOT_ICON}<span class="bot-name">HITAP Docs</span><div class="bot-body">${body}</div></div>`;
    }).join('');
    th.scrollTop = th.scrollHeight;
    renderQuick();
  }
  function renderQuick() {
    const tpl = TPL(current.tpl), p = current.pending, q = [];
    const pf = p && tpl.byKey[p];
    if (pf && pf.type === 'yesno') q.push([tr('ใช่', 'Yes'), tr('ใช่', 'yes')], [tr('ไม่ใช่', 'No'), tr('ไม่ใช่', 'no')]);
    if (pf && pf.type === 'date') q.push([tr('วันนี้', 'Today'), tr('วันนี้', 'today')]);
    if (p) q.push([tr('ข้าม', 'Skip'), tr('ข้าม', 'skip')]);
    if (!p) { if (tpl.signers.length) q.push([tr('ลงนาม', 'Sign'), tr('ลงนาม', 'sign')]); q.push([tr(`ดาวน์โหลด ${tpl.engine.ext.toUpperCase()}`, `Download ${tpl.engine.ext.toUpperCase()}`), 'download']); }
    $('#quick').innerHTML = q.map(([label, send]) => `<button type="button" data-send="${htmlEsc(send)}">${htmlEsc(label)}</button>`).join('');
  }

  /* ================= Workspace chrome ================= */
  function updateChrome() {
    if (!current) return;
    const tpl = TPL(current.tpl), miss = missingFields(current).length, st = $('#docStatus');
    st.classList.toggle('ready', !miss);
    st.textContent = miss ? tr(`กำลังร่าง · เหลือ ${miss} ข้อ`, `Drafting · ${miss} left`) : tr('เอกสารพร้อมตรวจ', 'Ready for review');
    $('#chatTitle').textContent = current.title;
    $('#docName').textContent = tpl.short[uiLang()];
    $('#docMade').innerHTML = `<svg viewBox="0 0 24 24" width="13" height="13" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round"><circle cx="12" cy="12" r="9"/><path d="m8 12 3 3 5-6"/></svg>${tr('สร้างแล้ว', 'Created')}`;
    const n = tpl.signers.filter(s => current.signs[s.role]).length;
    $('#signCount').textContent = n ? `${n}/${tpl.signers.length}` : '';
    $('#signBtn').hidden = !tpl.signers.length;
    $('#printBtn').hidden = tpl.kind === 'pdf';
    $('#downloadLabel').textContent = tpl.engine.ext.toUpperCase();
  }
  function renderForm() {
    const tpl = TPL(current.tpl), Lg = uiLang(), fields = visibleFields(tpl);
    const groups = tpl.groups
      ? tpl.groups.map((g, i) => ({ title: g[Lg], fields: fields.filter(f => f.group === i) }))
      : [{ title: tr('ข้อมูลหลัก', 'Key details'), fields: fields.filter(f => f.ask) }, { title: tr('ช่องอื่น ๆ', 'Other fields'), fields: fields.filter(f => !f.ask) }];
    $('#ndaForm').innerHTML = groups.filter(g => g.fields.length).map(g => `<div class="group-title">${htmlEsc(g.title)}</div>` + g.fields.map(f => {
      const v = current.values[f.key] ?? '', ph = f.ph ? f.ph[tpl.docLang] : f.mirror ? valueOf(current, f.mirror) : '';
      let input;
      if (f.type === 'textarea') input = `<textarea name="${htmlEsc(f.key)}" rows="2" placeholder="${htmlEsc(ph)}">${htmlEsc(v)}</textarea>`;
      else if (f.type === 'check') input = `<span class="check-row"><input type="checkbox" name="${htmlEsc(f.key)}" ${v ? 'checked' : ''}> ${tr('เลือก', 'Selected')}</span>`;
      else if (f.type === 'yesno') input = `<span class="yn"><label><input type="radio" name="${htmlEsc(f.key)}" value="yes" ${v === 'yes' ? 'checked' : ''}> ${tr('ใช่', 'Yes')}</label><label><input type="radio" name="${htmlEsc(f.key)}" value="no" ${v === 'no' ? 'checked' : ''}> ${tr('ไม่ใช่', 'No')}</label></span>`;
      else if (f.type === 'choice') input = `<select name="${htmlEsc(f.key)}">${(f.options || []).map(o => `<option value="${htmlEsc(o.v)}" ${String(o.v) === String(v) ? 'selected' : ''}>${htmlEsc(String(o.t).trim() || '—')}</option>`).join('')}</select>`;
      else input = `<input name="${htmlEsc(f.key)}" type="${f.type === 'date' ? 'date' : 'text'}" value="${htmlEsc(v)}" placeholder="${htmlEsc(ph)}" ${f.type === 'money' || f.type === 'number' ? 'inputmode="decimal"' : ''}>`;
      return `<label class="field" data-key="${htmlEsc(f.key)}"><span>${htmlEsc(f.label[Lg])}${f.ask ? '<span class="req">*</span>' : ''}</span>${input}<span class="note"></span></label>`;
    }).join('')).join('');
    fields.forEach(f => updateFieldState(f.key));
  }
  function updateFieldState(key) {
    const field = $(`#ndaForm .field[data-key="${CSS.escape(key)}"]`);
    if (!field || !current) return;
    const tpl = TPL(current.tpl), f = tpl.byKey[key];
    field.classList.toggle('missing', !filled(current, key) && !!f.ask);
    const note = $('.note', field), text = noteFor(key, current.values[key] || '');
    note.className = 'note' + (/⚠/.test(text) ? ' warn' : /✓/.test(text) ? ' ok' : '');
    note.textContent = text;
  }
  function toggleDrawer(open) {
    const d = $('#editDrawer');
    open = open ?? !d.classList.contains('open');
    if (open) renderForm();
    d.classList.toggle('open', open);
    d.setAttribute('aria-hidden', String(!open));
    $('#editBtn').setAttribute('aria-pressed', String(open));
  }

  async function download() {
    const tpl = TPL(current.tpl), miss = missingFields(current);
    if (miss.length && !confirm(tr(`ยังไม่ได้กรอก ${miss.length} ข้อ: ${miss.slice(0, 8).map(k => tpl.byKey[k].label.th).join(', ')}${miss.length > 8 ? ' …' : ''}\nดาวน์โหลดต่อหรือไม่?`,
      `${miss.length} field(s) still empty: ${miss.slice(0, 8).map(k => tpl.byKey[k].label.en).join(', ')}${miss.length > 8 ? ' …' : ''}\nDownload anyway?`))) return;
    try {
      toast(tr('กำลังสร้างไฟล์…', 'Building the file…'));
      const blob = await tpl.engine.build(await ctxFor(current));
      const nameField = tpl.fields.find(f => f.nameField && filled(current, f.key));
      const who = nameField ? valueOf(current, nameField.key) : '';
      const base = tpl.builtin ? `NDA_${tpl.docLang.toUpperCase()}` : tpl.file.replace(/\.[^.]+$/, '').slice(0, 50);
      saveBlob(blob, `${safeName(base)}${who ? '_' + safeName(who) : ''}.${tpl.engine.ext}`);
      toast(tr('ดาวน์โหลดแล้ว', 'Downloaded'));
    } catch (err) {
      console.error(err);
      toast(tr('สร้างไฟล์ไม่สำเร็จ: ', 'Could not build the file: ') + err.message);
    }
  }
  async function printDoc() {
    const tpl = TPL(current.tpl);
    if (tpl.kind === 'docx') await renderPreview('final');
    document.body.classList.add('printing');
    const restore = () => { window.removeEventListener('afterprint', restore); document.body.classList.remove('printing'); if (tpl.kind === 'docx') renderPreview(); };
    window.addEventListener('afterprint', restore);
    window.print();
  }
  function saveBlob(blob, name) {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob); a.download = name;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 4000);
  }
  const safeName = s => (s || '').replace(/[\\/:*?"<>|]+/g, '').replace(/\s+/g, '_').slice(0, 60);

  /* ================= Signatures ================= */
  let signRole = null, signMode = 'draw', padDirty = false, ink = '#1b2a6b', uploadCanvas = null;
  const pad = () => $('#pad');
  const signersOf = () => TPL(current.tpl).signers;
  const signerName = s => s.fixedName || (s.nameKey ? valueOf(current, s.nameKey) : '');
  function renderSigners() {
    $('#signers').innerHTML = signersOf().map(s => `<button type="button" class="signer${s.role === signRole ? ' active' : ''}" data-role="${htmlEsc(s.role)}">
      <small>${htmlEsc(s.label[uiLang()])}</small><b>${htmlEsc(signerName(s) || '—')}</b>
      <span class="st${current.signs[s.role] ? ' done' : ''}">${current.signs[s.role] ? tr('✓ ลงนามแล้ว', '✓ Signed') : tr('ยังไม่ลงนาม', 'Not signed')}</span></button>`).join('');
    $('#signRemove').style.visibility = current.signs[signRole] ? 'visible' : 'hidden';
  }
  function openSign(role) {
    if (!current || !signersOf().length) return;
    signRole = role || (signersOf().find(s => !current.signs[s.role]) || signersOf()[0]).role;
    $('#signDlg').showModal();
    setMode('draw'); renderSigners(); resetPad();
    $('#typeName').value = signerName(signersOf().find(s => s.role === signRole));
    renderTypePreview();
    uploadCanvas = null; $('#uploadPreview').hidden = true; $('#uploadFile').value = '';
  }
  function setMode(mode) {
    signMode = mode;
    $$('.sign-modes button').forEach(b => b.classList.toggle('active', b.dataset.mode === mode));
    $$('.sign-mode').forEach(el => { el.hidden = el.dataset.mode !== mode; });
    if (mode === 'draw') resetPad(true);
  }
  function resetPad(keep) {
    const c = pad(), r = c.getBoundingClientRect(), dpr = window.devicePixelRatio || 1;
    if (keep && c.width === Math.round(r.width * dpr) && padDirty) return;
    c.width = Math.round(r.width * dpr); c.height = Math.round(r.height * dpr);
    const ctx = c.getContext('2d');
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0); ctx.lineCap = 'round'; ctx.lineJoin = 'round';
    padDirty = false; $('#padHint').style.opacity = 1;
  }
  function bindPad() {
    const c = pad();
    let drawing = false, pts = [];
    const pos = e => { const r = c.getBoundingClientRect(); return { x: e.clientX - r.left, y: e.clientY - r.top, t: performance.now() }; };
    c.addEventListener('pointerdown', e => {
      drawing = true; c.setPointerCapture(e.pointerId); pts = [pos(e)];
      const ctx = c.getContext('2d'); ctx.fillStyle = ink;
      ctx.beginPath(); ctx.arc(pts[0].x, pts[0].y, 1.3, 0, Math.PI * 2); ctx.fill();
      padDirty = true; $('#padHint').style.opacity = 0;
    });
    c.addEventListener('pointermove', e => {
      if (!drawing) return;
      pts.push(pos(e));
      if (pts.length < 3) return;
      const [a, b, cc] = pts.slice(-3);
      const speed = Math.hypot(cc.x - b.x, cc.y - b.y) / Math.max(1, cc.t - b.t);
      const ctx = c.getContext('2d');
      ctx.strokeStyle = ink; ctx.lineWidth = Math.max(1.4, Math.min(3.4, 3.6 - speed * 1.2));
      ctx.beginPath(); ctx.moveTo((a.x + b.x) / 2, (a.y + b.y) / 2); ctx.quadraticCurveTo(b.x, b.y, (b.x + cc.x) / 2, (b.y + cc.y) / 2); ctx.stroke();
    });
    const end = () => { drawing = false; pts = []; };
    c.addEventListener('pointerup', end); c.addEventListener('pointercancel', end);
  }
  function renderTypePreview() { const el = $('#typePreview'); el.textContent = $('#typeName').value || ' '; el.style.color = ink; }
  function trimCanvas(src) {
    const ctx = src.getContext('2d'), { width: w, height: h } = src;
    const data = ctx.getImageData(0, 0, w, h).data;
    let x0 = w, y0 = h, x1 = -1, y1 = -1;
    for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) if (data[(y * w + x) * 4 + 3] > 16) { if (x < x0) x0 = x; if (x > x1) x1 = x; if (y < y0) y0 = y; if (y > y1) y1 = y; }
    if (x1 < 0) return null;
    const padPx = 8; x0 = Math.max(0, x0 - padPx); y0 = Math.max(0, y0 - padPx); x1 = Math.min(w - 1, x1 + padPx); y1 = Math.min(h - 1, y1 + padPx);
    const cw = x1 - x0 + 1, ch = y1 - y0 + 1, scale = Math.min(1, 180 / ch);
    const out = document.createElement('canvas');
    out.width = Math.round(cw * scale); out.height = Math.round(ch * scale);
    out.getContext('2d').drawImage(src, x0, y0, cw, ch, 0, 0, out.width, out.height);
    return { data: out.toDataURL('image/png'), w: out.width, h: out.height };
  }
  async function typedCanvas(text) {
    await document.fonts.load('96px Charmonman').catch(() => {});
    const c = document.createElement('canvas'), ctx = c.getContext('2d');
    ctx.font = '96px Charmonman';
    c.width = Math.ceil(ctx.measureText(text).width) + 60; c.height = 200;
    ctx.font = '96px Charmonman'; ctx.fillStyle = ink; ctx.textBaseline = 'middle';
    ctx.fillText(text, 30, 100);
    return c;
  }
  function loadUpload(file) {
    const img = new Image();
    img.onload = () => {
      const scale = Math.min(1, 1200 / img.width), c = document.createElement('canvas');
      c.width = Math.round(img.width * scale); c.height = Math.round(img.height * scale);
      const ctx = c.getContext('2d'); ctx.drawImage(img, 0, 0, c.width, c.height);
      const id = ctx.getImageData(0, 0, c.width, c.height), d = id.data;
      for (let i = 0; i < d.length; i += 4) { const lum = (d[i] + d[i + 1] + d[i + 2]) / 3; if (lum > 225) d[i + 3] = 0; else if (lum > 180) d[i + 3] = Math.round(d[i + 3] * (225 - lum) / 45); }
      ctx.putImageData(id, 0, 0);
      uploadCanvas = c;
      $('#uploadPreview').src = c.toDataURL('image/png'); $('#uploadPreview').hidden = false;
      URL.revokeObjectURL(img.src);
    };
    img.src = URL.createObjectURL(file);
  }
  async function applySignature() {
    let sig = null;
    if (signMode === 'draw') { if (!padDirty) return toast(tr('ยังไม่ได้วาดลายเซ็น', 'Draw a signature first')); sig = trimCanvas(pad()); }
    if (signMode === 'type') { const t = $('#typeName').value.trim(); if (!t) return toast(tr('พิมพ์ชื่อก่อนครับ', 'Type a name first')); sig = trimCanvas(await typedCanvas(t)); }
    if (signMode === 'upload') { if (!uploadCanvas) return toast(tr('เลือกรูปลายเซ็นก่อนครับ', 'Choose an image first')); sig = trimCanvas(uploadCanvas); }
    if (!sig) return toast(tr('ไม่พบลายเซ็นในภาพ', 'No signature found'));
    current.signs[signRole] = sig;
    const s = signersOf().find(x => x.role === signRole);
    bot({ t: tr(`ใส่ลายเซ็น${s.label.th}แล้วครับ ✓`, `${s.label.en} signature added ✓`) });
    afterSignChange();
    const next = signersOf().find(x => !current.signs[x.role]);
    if (next) { signRole = next.role; renderSigners(); resetPad(); $('#typeName').value = signerName(next); renderTypePreview(); uploadCanvas = null; $('#uploadPreview').hidden = true; }
    else $('#signDlg').close();
    toast(tr('ใส่ลายเซ็นในเอกสารแล้ว', 'Signature placed in the document'));
  }
  function afterSignChange() { touch(); updateChrome(); renderThread(); renderSigners(); TPL(current.tpl).engine.inline ? refreshInline() : schedulePreview(30); }

  /* ================= Home / forms / history ================= */
  let homeLang = 'auto';
  const STARTERS = ['nda-th', 'leave-adm', 'receipt-th-pcht', 'coi-th'];
  function renderHome() {
    $('#starters').innerHTML = STARTERS.map(TPL).filter(Boolean).map(t => `<button type="button" class="starter" data-tpl="${t.id}"><small>${htmlEsc((CATS.find(c => c.id === t.cat) || {})[uiLang()] || '')} · ${KIND_LABEL[t.kind][uiLang()]}</small><b>${htmlEsc(t.short[uiLang()])}</b></button>`).join('');
  }
  let activeCat = 'all';
  function renderForms() {
    const q = ($('#formSearch').value || '').trim().toLowerCase(), Lg = uiLang();
    const match = t => (activeCat === 'all' || t.cat === activeCat) && (!q || (t.title.th + t.title.en + t.desc.th + t.desc.en + t.file).toLowerCase().includes(q));
    $('#formCount').textContent = tr(`${VISIBLE().length} แบบ`, `${VISIBLE().length} forms`);
    $('#cats').innerHTML = CATS.map(c => `<button type="button" class="cat${c.id === activeCat ? ' active' : ''}" data-cat="${c.id}">${c[Lg]} <em>${VISIBLE().filter(t => c.id === 'all' || t.cat === c.id).length}</em></button>`).join('');
    const list = VISIBLE().filter(match);
    $('#catLabel').textContent = `${CATS.find(c => c.id === activeCat)[Lg]} · ${tr(`${list.length} แบบ`, `${list.length} forms`)}`;
    $('#formList').innerHTML = list.length ? list.map(t => {
      const n = visibleFields(t).length;
      return `<article class="form-card"><h3>${htmlEsc(t.title[Lg])}</h3>
      <div class="badges"><span class="badge badge-${t.kind}">${KIND_LABEL[t.kind][Lg]}</span><span class="badge">HITAP</span><span class="badge">${tr(`${n} ช่องกรอก`, `${n} fields`)}</span>${t.signers.length ? `<span class="badge">${tr(`ลงนาม ${t.signers.length} ท่าน`, `${t.signers.length} signatures`)}</span>` : ''}</div>
      <p>${htmlEsc(t.desc[Lg])}</p>
      <div class="card-actions">
        <button type="button" class="btn btn-green btn-sm" data-dl="${t.id}"><svg viewBox="0 0 24 24" width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M12 3v12M7 10l5 5 5-5M5 21h14"/></svg>${tr('ดาวน์โหลดแบบฟอร์มเปล่า', 'Download blank')}</button>
        <button type="button" class="btn btn-light btn-sm" data-fill="${t.id}"><svg viewBox="0 0 24 24" width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3l1.8 4.6L18.5 9l-4.7 1.6L12 15l-1.8-4.4L5.5 9l4.7-1.4z"/></svg>${tr('กรอกข้อมูลให้ (ฟรี)', 'Fill it in (free)')}</button>
      </div></article>`;
    }).join('') : `<div class="no-result">${tr('ไม่พบแบบฟอร์ม', 'No forms found')}</div>`;
  }
  function renderHistory() {
    const q = ($('#historySearch').value || '').trim().toLowerCase();
    const items = docs.filter(d => TPL(d.tpl) && (!q || (d.title + JSON.stringify(d.values)).toLowerCase().includes(q)));
    $('#historyList').innerHTML = items.length
      ? items.map(d => `<li><a href="#/doc/${d.id}" class="${current && d.id === current.id && location.hash.startsWith('#/doc') ? 'active' : ''}" title="${htmlEsc(d.title)}">${htmlEsc(d.title)}</a></li>`).join('')
      : `<li class="empty">${q ? tr('ไม่พบแชท', 'No matches') : tr('ยังไม่มีประวัติแชท', 'No chats yet')}</li>`;
  }
  function applyPlaceholders() { $$('[data-ph-th]').forEach(el => { el.placeholder = el.dataset[uiLang() === 'en' ? 'phEn' : 'phTh']; }); }
  function setLang(l) {
    document.documentElement.dataset.lang = l; document.documentElement.lang = l;
    store.set(LS.lang, l);
    $$('.lang-switch button').forEach(b => b.classList.toggle('active', b.dataset.lang === l));
    applyPlaceholders(); route();
    if (HD.cloud) renderAccount({ user: HD.cloud.user(), status: HD.cloud.status() });
  }
  function toast(msg) {
    const t = $('#toast'); t.textContent = msg; t.classList.add('show');
    clearTimeout(toast.t); toast.t = setTimeout(() => t.classList.remove('show'), 3200);
  }

  /* ================= Router ================= */
  function show(page) {
    $$('.page').forEach(p => p.classList.toggle('active', p.id === `page-${page}`));
    $('#sidebar').classList.remove('open'); $('#scrim').classList.remove('show');
    closePop();
  }
  function route() {
    const h = location.hash || '#/';
    let m;
    if ((m = /^#\/new\/([\w-]+)$/.exec(h)) && TPL(m[1])) { startDoc('', m[1]); return; }
    if ((m = /^#\/doc\/(\w+)/.exec(h))) {
      const d = docs.find(x => x.id === m[1]);
      if (d && !TPL(d.tpl) && !remoteLoaded) { show('doc'); $('#preview').innerHTML = `<div class="preview-msg">${tr('กำลังโหลดแม่แบบ…', 'Loading template…')}</div>`; return; }
      if (!d || !TPL(d.tpl)) { location.hash = '#/'; return; }
      if (current !== d) { current = d; ctxCache = null; $('#preview').replaceChildren(); $('#paperScroll').scrollTop = 0; toggleDrawer(false); }
      d.signs = d.signs || {}; d.skipped = d.skipped || []; d.msgs = d.msgs || [];
      show('doc'); updateChrome(); renderThread(); renderPreview(); renderHistory();
      return;
    }
    if (h.startsWith('#/admin')) { show('admin'); if (HD.admin) HD.admin.render(); else $('#adminRoot').innerHTML = `<div class="preview-msg">${tr('กำลังโหลด…', 'Loading…')}</div>`; }
    else if (h.startsWith('#/forms')) { show('forms'); renderForms(); }
    else { show('home'); renderHome(); }
    renderHistory();
  }

  /* ================= Events ================= */
  function bind() {
    window.addEventListener('hashchange', route);
    window.addEventListener('resize', () => { clearTimeout(bind.r); bind.r = setTimeout(fitPreview, 120); closePop(); });
    $$('.lang-switch button').forEach(b => b.addEventListener('click', () => setLang(b.dataset.lang)));
    $('#menuBtn').addEventListener('click', () => { $('#sidebar').classList.add('open'); $('#scrim').classList.add('show'); });
    $('#scrim').addEventListener('click', () => { $('#sidebar').classList.remove('open'); $('#scrim').classList.remove('show'); });
    $('#historySearch').addEventListener('input', renderHistory);
    $('#helpBtn').addEventListener('click', () => $('#helpDlg').showModal());
    $$('[data-close]').forEach(b => b.addEventListener('click', () => b.closest('dialog').close()));

    // Home
    $('#homeLang').addEventListener('click', e => {
      const b = e.target.closest('button'); if (!b) return;
      homeLang = b.dataset.v; $$('#homeLang button').forEach(x => x.classList.toggle('active', x === b));
    });
    const autoGrow = ta => { ta.style.height = 'auto'; ta.style.height = Math.min(ta.scrollHeight, 200) + 'px'; };
    const submitOnEnter = (ta, form) => {
      ta.addEventListener('keydown', e => { if (e.key === 'Enter' && !e.shiftKey && !e.isComposing) { e.preventDefault(); form.requestSubmit(); } });
      ta.addEventListener('input', () => autoGrow(ta));
      form.addEventListener('submit', () => setTimeout(() => autoGrow(ta)));
    };
    submitOnEnter($('#homeText'), $('#homeComposer'));
    submitOnEnter($('#chatText'), $('#chatComposer'));
    $('#homeComposer').addEventListener('submit', e => {
      e.preventDefault();
      const t = $('#homeText').value.trim();
      const id = t ? detectTemplate(t, homeLang) : `nda-${homeLang === 'en' ? 'en' : 'th'}`;
      if (!id) {
        $('#homeHint').hidden = false;
        $('#homeHint').innerHTML = `${tr('ยังไม่แน่ใจว่าต้องการเอกสารแบบไหน — เลือกแม่แบบ:', 'Not sure which document you need — pick a template:')} ${VISIBLE().map(x => `<button type="button" class="chip-tpl" data-tpl="${x.id}">${htmlEsc(x.short[uiLang()].slice(0, 40))}</button>`).join('')}`;
        return;
      }
      $('#homeText').value = ''; $('#homeHint').hidden = true;
      startDoc(t, id);
    });
    $('#homeHint').addEventListener('click', e => {
      const b = e.target.closest('[data-tpl]'); if (!b) return;
      const t = $('#homeText').value.trim();
      $('#homeText').value = ''; $('#homeHint').hidden = true;
      startDoc(t, b.dataset.tpl);
    });
    $('#starters').addEventListener('click', e => { const b = e.target.closest('[data-tpl]'); if (b) startDoc('', b.dataset.tpl); });

    // Forms
    $('#formSearch').addEventListener('input', renderForms);
    $('#cats').addEventListener('click', e => { const b = e.target.closest('[data-cat]'); if (b) { activeCat = b.dataset.cat; renderForms(); } });
    $('#formList').addEventListener('click', e => {
      const dl = e.target.closest('[data-dl]'), fill = e.target.closest('[data-fill]');
      if (dl) {
        const t = TPL(dl.dataset.dl);
        if (t.original) saveBlob(t.original(), t.file);
        else { const a = document.createElement('a'); a.href = t.originalHref; a.download = t.file; document.body.appendChild(a); a.click(); a.remove(); }
      }
      if (fill) startDoc('', fill.dataset.fill);
    });

    // Chat
    $('#chatComposer').addEventListener('submit', e => { e.preventDefault(); const t = $('#chatText').value; $('#chatText').value = ''; handleUser(t); });
    $('#quick').addEventListener('click', e => { const b = e.target.closest('[data-send]'); if (b) handleUser(b.dataset.send); });
    $('#thread').addEventListener('click', e => { if (e.target.closest('[data-open-doc]')) setPane('doc'); });
    $$('.ws-tabs button').forEach(b => b.addEventListener('click', () => setPane(b.dataset.pane)));

    // Document: chips (Word), inline inputs (PDF/Excel), check boxes, signatures
    const box = $('#preview');
    box.addEventListener('click', e => {
      const f = e.target.closest('.fld'), s = e.target.closest('.sigslot, .pf-sig'), c = e.target.closest('.chk, .pf-chk'), dt = e.target.closest('.pf-date, .xi-date');
      if (f) openPop(f.dataset.key, f);
      else if (dt) openPop(dt.dataset.key, dt);
      else if (s) openSign(s.dataset.role);
      else if (c) {
        const key = c.dataset.check || c.dataset.key, fld = TPL(current.tpl).byKey[key];
        if (fld.type === 'yesno') setValue(key, current.values[key] === c.dataset.role ? '' : c.dataset.role);
        else setValue(key, !current.values[key]);
        if (TPL(current.tpl).kind === 'docx') { c.classList.toggle('on', !!current.values[key]); c.textContent = current.values[key] ? '☑' : '☐'; }
        afterManualEdit(key);
      }
    });
    box.addEventListener('keydown', e => { if ((e.key === 'Enter' || e.key === ' ') && e.target.matches('.fld,.sigslot,.chk')) { e.preventDefault(); e.target.click(); } });
    box.addEventListener('input', e => {
      const el = e.target; if (!el.dataset || !el.dataset.key || !el.matches('.pf-text, .xi')) return;
      current.values[el.dataset.key] = el.value;
      el.classList.toggle('empty', !el.value);
      clearTimeout(bind.i); bind.i = setTimeout(() => { rememberProfile(current, [el.dataset.key]); touch(); updateChrome(); refreshInline(); updateFieldState(el.dataset.key); }, 250);
    });
    box.addEventListener('change', e => {
      const el = e.target; if (!el.dataset || !el.dataset.key) return;
      if (el.matches('.pf-sel')) setValue(el.dataset.key, el.value);
      else if (el.matches('.pf-text, .xi')) { setValue(el.dataset.key, el.value); afterManualEdit(el.dataset.key); }
    });
    box.addEventListener('focusout', e => { if (e.target.matches && e.target.matches('.pf-text, .xi')) setTimeout(refreshInline, 0); });
    $('#paperScroll').addEventListener('scroll', closePop);
    $('#popSave').addEventListener('click', savePop);
    $('#popCancel').addEventListener('click', closePop);
    document.addEventListener('mousedown', e => { if (!$('#pop').hidden && !e.target.closest('#pop, .fld, .pf-date, .xi-date')) closePop(); });
    $('#editBtn').addEventListener('click', () => toggleDrawer());
    $('#closeDrawer').addEventListener('click', () => toggleDrawer(false));
    $('#ndaForm').addEventListener('input', e => {
      const el = e.target; if (!el.name) return;
      const v = el.type === 'checkbox' ? el.checked : el.value;
      current.values[el.name] = v;
      updateFieldState(el.name); updateChrome();
      TPL(current.tpl).engine.inline ? refreshInline() : schedulePreview();
      clearTimeout(bind.s); bind.s = setTimeout(() => { rememberProfile(current, [el.name]); touch(); }, 500);
    });
    $('#ndaForm').addEventListener('change', e => { if (e.target.name) afterManualEdit(e.target.name); });
    $('#downloadBtn').addEventListener('click', download);
    $('#printBtn').addEventListener('click', printDoc);
    $('#signBtn').addEventListener('click', () => openSign());

    // Signature dialog
    bindPad();
    $('#signers').addEventListener('click', e => {
      const b = e.target.closest('[data-role]'); if (!b) return;
      signRole = b.dataset.role; renderSigners(); resetPad();
      $('#typeName').value = signerName(signersOf().find(s => s.role === signRole)); renderTypePreview();
    });
    $$('.sign-modes button').forEach(b => b.addEventListener('click', () => setMode(b.dataset.mode)));
    $('#inks').addEventListener('click', e => {
      const b = e.target.closest('[data-ink]'); if (!b) return;
      ink = b.dataset.ink; $$('#inks .ink').forEach(x => x.classList.toggle('active', x === b)); renderTypePreview();
    });
    $('#padClear').addEventListener('click', () => resetPad());
    $('#typeName').addEventListener('input', renderTypePreview);
    $('#uploadFile').addEventListener('change', e => { const f = e.target.files[0]; if (f) loadUpload(f); });
    $('#signSave').addEventListener('click', applySignature);
    $('#signRemove').addEventListener('click', () => {
      delete current.signs[signRole];
      const s = signersOf().find(x => x.role === signRole);
      bot({ t: tr(`ลบลายเซ็น${s.label.th}แล้ว`, `${s.label.en} signature removed`) });
      afterSignChange();
    });
    $('#signClose').addEventListener('click', () => $('#signDlg').close());
  }
  function setPane(p) {
    $('#ws').dataset.pane = p;
    $$('.ws-tabs button').forEach(b => b.classList.toggle('active', b.dataset.pane === p));
    if (p === 'doc') setTimeout(fitPreview, 30);
  }

  /* ================= Cloud bridge (cloud.js) ================= */
  HD.appBridge = {
    getDocs: () => docs,
    getProfile: () => store.get(LS.profile, {}),
    replaceAll(list, profile) {
      docs = list.filter(d => d && d.tpl);
      if (current) current = docs.find(d => d.id === current.id) || current;
      ctxCache = null;
      store.set(LS.docs, docs.slice(0, 40));
      store.set(LS.profile, profile || {});
      route();
    },
    /** rows from public.templates: hide built-ins, add/replace uploaded templates */
    setRemoteTemplates(rows, urlOf) {
      for (const t of TEMPLATES) if (t.builtin || HD_BUILTIN_IDS.has(t.id)) t.hidden = false;
      for (let i = TEMPLATES.length - 1; i >= 0; i--) if (TEMPLATES[i].remote) TEMPLATES.splice(i, 1);
      for (const r of rows) {
        const existing = TPL(r.id);
        if (existing && HD_BUILTIN_IDS.has(r.id)) { existing.hidden = !!r.hidden; continue; }
        if (r.builtin || !r.meta || !r.meta.fields) continue;
        const t = normalize({ ...r.meta, id: r.id, kind: r.kind, remote: true, hidden: !!r.hidden, dataPath: r.data_path, originalPath: r.original_path,
          dataUrl: r.data_path ? urlOf(r.data_path) : null, originalHref: r.original_path ? urlOf(r.original_path) : null });
        TEMPLATES.push(t);
      }
      remoteLoaded = true;
      if (!location.hash.startsWith('#/admin')) route(); else renderHistory();
    },
    templates: () => TEMPLATES,
    builtinIds: () => [...HD_BUILTIN_IDS],
    categories: () => CATS.filter(c => c.id !== 'all'),
    setAdmin() { if (location.hash.startsWith('#/admin') && HD.admin) HD.admin.render(); },
    toast: msg => toast(msg),
    tr: (a, b) => tr(a, b),
    remoteFailed() { remoteLoaded = true; route(); },
    clearLocal() {
      docs = []; current = null; ctxCache = null;
      try { localStorage.removeItem(LS.docs); localStorage.removeItem(LS.profile); } catch (e) { /* storage blocked */ }
      location.hash = '#/'; route();
    },
  };
  function renderAccount({ user, status }) {
    const box = $('#account');
    if (!box) return;
    if (!HD.cloud || !HD.cloud.enabled()) { box.hidden = true; return; }
    box.hidden = false;
    const label = {
      syncing: tr('กำลังซิงก์…', 'Syncing…'), saving: tr('กำลังบันทึก…', 'Saving…'), synced: tr('บันทึกบนคลาวด์แล้ว', 'Saved to cloud'),
      error: tr('ซิงก์ไม่สำเร็จ — ลองใหม่', 'Sync failed — retry'), offline: tr('ออฟไลน์', 'Offline'),
    }[status] || '';
    $('#avatar').textContent = user ? user.email[0].toUpperCase() : 'H';
    $('#whoName').textContent = user ? user.email : 'HITAP';
    $('#whoSub').textContent = user ? label : tr('ยังไม่ได้เข้าสู่ระบบ · เก็บไว้ในเครื่องนี้', 'Not signed in · saved on this device');
    $('#authBtn').textContent = user ? tr('ออกจากระบบ', 'Sign out') : tr('เข้าสู่ระบบ', 'Sign in');
    $('#authBtn').dataset.mode = user ? 'out' : 'in';
  }
  function bindAccount() {
    $('#authBtn').addEventListener('click', async () => {
      if ($('#authBtn').dataset.mode === 'out') {
        if (confirm(tr('ออกจากระบบ? เอกสารในเครื่องนี้จะถูกล้าง (ยังอยู่บนคลาวด์)', 'Sign out? Documents on this device are cleared (they stay in the cloud).'))) await HD.cloud.signOut();
        return;
      }
      $('#loginMsg').textContent = ''; $('#loginPassword').value = ''; $('#loginDlg').showModal(); $('#loginEmail').focus();
    });
    $('#whoSub').addEventListener('click', () => { if (HD.cloud && HD.cloud.status() === 'error') HD.cloud.sync(); });
    const setLoginMode = m => {
      $('#loginForm').dataset.mode = m; $('#loginMsg').textContent = '';
      $('#loginPassword').autocomplete = m === 'up' ? 'new-password' : 'current-password';
    };
    $('#loginToggle').addEventListener('click', () => setLoginMode($('#loginForm').dataset.mode === 'in' ? 'up' : 'in'));
    $('#loginForm').addEventListener('submit', async e => {
      e.preventDefault();
      const email = $('#loginEmail').value.trim(), password = $('#loginPassword').value;
      const signUp = $('#loginForm').dataset.mode === 'up';
      if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) { $('#loginMsg').textContent = tr('อีเมลไม่ถูกต้อง', 'Invalid email'); return; }
      if (password.length < 6) { $('#loginMsg').textContent = tr('รหัสผ่านอย่างน้อย 6 ตัวอักษร', 'Password must be at least 6 characters'); return; }
      $('#loginSend').disabled = true;
      $('#loginMsg').textContent = signUp ? tr('กำลังสร้างบัญชี…', 'Creating account…') : tr('กำลังเข้าสู่ระบบ…', 'Signing in…');
      try {
        const res = signUp ? await HD.cloud.signUp(email, password) : await HD.cloud.signInPassword(email, password);
        if (signUp && !res.session) {
          $('#loginMsg').textContent = tr('สร้างบัญชีแล้ว — ต้องยืนยันอีเมลก่อนจึงจะเข้าสู่ระบบได้', 'Account created — confirm your email before signing in');
          return;
        }
        $('#loginDlg').close();
        toast(tr('เข้าสู่ระบบแล้ว', 'Signed in'));
      } catch (err) {
        const m = err.message || '';
        $('#loginMsg').textContent = /invalid/i.test(m) ? tr('อีเมลหรือรหัสผ่านไม่ถูกต้อง', 'Wrong email or password')
          : /not confirmed/i.test(m) ? tr('บัญชียังไม่ได้ยืนยันอีเมล', 'This account email is not confirmed yet')
          : /already registered/i.test(m) ? tr('อีเมลนี้มีบัญชีแล้ว — เข้าสู่ระบบแทน', 'This email already has an account — sign in instead')
          : /rate|too many/i.test(m) ? tr('ลองหลายครั้งเกินไป รอสักครู่แล้วลองใหม่', 'Too many attempts — wait a moment')
          : m;
      } finally { $('#loginSend').disabled = false; }
    });
    $('#loginClose').addEventListener('click', () => $('#loginDlg').close());
    // cloud.js loads after this file and calls HD.onCloudReady once HD.cloud exists
    HD.onCloudReady = () => HD.cloud.onChange(renderAccount);
    if (HD.cloud) HD.onCloudReady();
  }

  /* theme: system (default) → light → dark */
  function bindTheme() {
    const order = ['system', 'light', 'dark'];
    const label = m => ({ system: tr('ธีม: ตามระบบ', 'Theme: system'), light: tr('ธีม: สว่าง', 'Theme: light'), dark: tr('ธีม: มืด', 'Theme: dark') })[m];
    const cur = () => document.documentElement.dataset.theme || 'system';
    const apply = m => {
      if (m === 'system') delete document.documentElement.dataset.theme; else document.documentElement.dataset.theme = m;
      try { m === 'system' ? localStorage.removeItem('hitap_theme') : localStorage.setItem('hitap_theme', m); } catch (e) { /* storage blocked */ }
      $('#themeBtn').title = label(m); $('#themeBtn').setAttribute('aria-label', label(m));
    };
    $('#themeBtn').addEventListener('click', () => { const m = order[(order.indexOf(cur()) + 1) % 3]; apply(m); toast(label(m)); });
    apply(cur());
  }

  bind();
  bindTheme();
  bindAccount();
  setLang(uiLang());
})();
