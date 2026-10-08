// Human-curated details for each template in templates/. Everything not listed here is
// auto-extracted by extract.mjs (field positions, dotted lines, blue Excel cells, PDF form fields).
//
// Field options: label {th,en}, type (text|textarea|date|money|number|check|choice|yesno),
// ask (assistant asks for it, in this order), tag (shared profile value: name, position, unit,
// employee_code, phone, email, id_card, address, organization), def ('today'), hide, mirror (key).

const L = (th, en = th) => ({ th, en });

/* ---------- Leave form (no form fields: values are drawn at these PDF coordinates) ---------- */
const leaveFields = () => {
  const f = [
    { key: 'emp_name', label: L('ชื่อ-สกุล', 'Name'), rect: [108, 703, 184, 17], tag: 'name', ask: 1 },
    { key: 'emp_code', label: L('รหัสพนักงาน', 'Employee code'), rect: [441, 703, 132, 17], tag: 'employee_code', ask: 2 },
    { key: 'emp_position', label: L('ตำแหน่ง', 'Position'), rect: [119, 680, 173, 17], tag: 'position', ask: 3 },
    { key: 'emp_unit', label: L('ฝ่าย', 'Unit'), rect: [347, 680, 226, 17], tag: 'unit', ask: 4 },
  ];
  for (let i = 0; i < 5; i++) {
    const y = 605 - 23.3 * i;
    f.push({ key: `work${i + 1}`, label: L(`โครงการ/งานที่ยังไม่แล้วเสร็จ ${i + 1}`, `Unfinished work ${i + 1}`), rect: [47, y, 243, 17], group: 'work', ask: i === 0 ? 8 : 0 });
    f.push({ key: `work${i + 1}_todo`, label: L(`ส่วนที่ต้องทำต่อ/ข้อมูลที่เกี่ยวข้อง ${i + 1}`, `Work to do / archives ${i + 1}`), rect: [296, y, 186, 17], group: 'work' });
    f.push({ key: `work${i + 1}_person`, label: L(`ผู้รับผิดชอบแทน ${i + 1}`, `Responsible person ${i + 1}`), rect: [488, y, 84, 17], group: 'work', size: 13 });
  }
  f.push(
    { key: 'leave_type', label: L('ประเภทการลา', 'Leave type'), rect: [216, 467, 76, 17], ask: 5, ph: L('ลาพักผ่อน', 'Annual leave'), kw: ['ประเภทการลา', 'ขอ>ลา', 'wish to leave', 'leave type'] },
    { key: 'leave_from', label: L('ลาตั้งแต่วันที่', 'Leave since'), type: 'date', rect: [373, 467, 56, 17], ask: 6, size: 13, kw: ['ตั้งแต่วันที่', 'ตั้งแต่', 'since', 'from'] },
    { key: 'leave_to', label: L('ลาถึงวันที่', 'Leave until'), type: 'date', rect: [494, 467, 79, 17], ask: 7, size: 13, kw: ['ถึงวันที่', 'ถึง', 'until', 'to'] },
  );
  for (let i = 0; i < 10; i++) {
    const base = 410.7 - 20.3 * (i % 5), x = i < 5 ? 47 : 333, w = i < 5 ? 243 : 240;
    f.push({ key: `coord${i + 1}`, label: L(`เจ้าหน้าที่ร่วมโครงการ ${i + 1}`, `Project co-ordinator ${i + 1}`), rect: [x, base - 5, w, 17], group: 'coord' });
  }
  f.push(
    { key: 'contact_tel', label: L('เบอร์โทรติดต่อระหว่างลา', 'Tel. during leave'), rect: [77, 252, 215, 17], tag: 'phone', ask: 9 },
    { key: 'contact_email', label: L('อีเมลติดต่อระหว่างลา', 'Email during leave'), rect: [333, 252, 240, 17], tag: 'email', ask: 10 },
    { key: 'contact_person', label: L('บุคคลที่ติดต่อได้', 'Contact person'), rect: [184, 228, 108, 17] },
    { key: 'contact_relation', label: L('เกี่ยวข้องเป็น', 'Relationship'), rect: [423, 228, 150, 17] },
    { key: 'contact_person_tel', label: L('เบอร์โทรบุคคลที่ติดต่อ', 'Contact person tel.'), rect: [77, 204, 215, 17] },
    { key: 'contact_person_email', label: L('อีเมลบุคคลที่ติดต่อ', 'Contact person email'), rect: [333, 204, 240, 17] },
    { key: 'sign_position', label: L('ตำแหน่ง (ผู้ลงนาม)', 'Position (signer)'), rect: [123, 151, 119, 17], mirror: 'emp_position' },
    { key: 'sign_date', label: L('วันที่ (ผู้ขอลา)', 'Date (employee)'), type: 'date', rect: [91, 130, 151, 17], def: 'today' },
    { key: 'hr_date', label: L('วันที่ (ฝ่ายบุคคล)', 'Date (HR)'), type: 'date', rect: [77, 47, 165, 17] },
    { key: 'head_date', label: L('วันที่ (หัวหน้าฝ่าย)', 'Date (Head of Unit)'), type: 'date', rect: [331, 47, 154, 17] },
  );
  return f;
};
const leaveSigners = [
  { role: 'employee', label: L('พนักงานที่ขอลา', 'Employee on leave'), nameKey: 'emp_name', rect: [104, 171, 138, 28] },
  { role: 'hr', label: L('ฝ่ายบุคคล', 'Human Resource Officer'), rect: [77, 88, 138, 28] },
  { role: 'head', label: L('หัวหน้าฝ่าย', 'Head of Unit'), rect: [331, 88, 154, 28] },
];

/* ---------- COI declaration: each question = Yes/No check boxes + a detail box ---------- */
const COI_EN = [
  ['Text3', 'การจ้างงาน', 'Employment'],
  ['Text6', 'การเป็นที่ปรึกษา', 'Consultancies / advisor'],
  ['Text7', 'การสนับสนุนการวิจัย', 'Research support'],
  ['Text11', 'การสนับสนุนที่ไม่ใช่ตัวเงิน / ค่าตอบแทนวิทยากร', 'Non-monetary support / speaker honoraria'],
  ['Text12', 'หุ้น ตราสารหนี้ หลักทรัพย์', 'Stocks, bonds, securities'],
  ['Text13', 'ผลประโยชน์ในกิจการธุรกิจ', 'Commercial business interests'],
  ['Text16', 'สิทธิบัตร เครื่องหมายการค้า ลิขสิทธิ์', 'Patents, trademarks, copyrights'],
  ['Text17', 'สิทธิในความรู้ความชำนาญ/เทคโนโลยี', 'Proprietary know-how'],
  ['Text18', 'ให้ความเห็นในฐานะผู้เชี่ยวชาญ/พยาน', 'Expert opinion or testimony'],
  ['Text19', 'ดำรงตำแหน่งที่เกี่ยวข้อง', 'Office or position held'],
  ['Text20', 'ทำงานให้คู่แข่ง/เข้าถึงข้อมูลคู่แข่ง', 'Work for a competitor'],
  ['Text21', 'ผลงานกระทบผลประโยชน์ของผู้อื่น', 'Outcome affects others you are close to'],
  ['Text22', 'กิจกรรมวิชาชีพที่อาจขัดกัน', 'Professional activities that may conflict'],
  ['Text23', 'ค่าตอบแทนการพูดเกี่ยวกับงาน HITAP', 'Payments for speaking about HITAP work'],
  ['Text24', 'ของขวัญหรือการต้อนรับ', 'Gifts or hospitality'],
  ['Text25', 'เรื่องอื่นที่อาจกระทบความเป็นกลาง', 'Anything else affecting independence'],
];
const COI_TH = [
  ['Text1', 'การจ้างงาน', 'Employment'],
  ['Text2', 'การให้คำปรึกษา / การเป็นที่ปรึกษา', 'Consultancies / advisor'],
  ['Text3', 'การสนับสนุนการวิจัย ทุน ค่าตอบแทนวิทยากร', 'Research support / honoraria'],
  ['Text4', 'หุ้น ตราสารหนี้ หลักทรัพย์', 'Stocks, bonds, securities'],
  ['Text5', 'ผลประโยชน์ในกิจการธุรกิจ', 'Commercial business interests'],
  ['Text6', 'สิทธิบัตร เครื่องหมายการค้า ลิขสิทธิ์', 'Patents, trademarks, copyrights'],
  ['Text7', 'สิทธิในความรู้ความชำนาญ/เทคโนโลยี', 'Proprietary know-how'],
  ['Text8', 'ให้ความเห็นในฐานะผู้เชี่ยวชาญ/พยาน', 'Expert opinion or testimony'],
  ['Text9', 'ดำรงตำแหน่งที่เกี่ยวข้อง', 'Office or position held'],
  ['Text10', 'ทำงานให้คู่แข่ง/เข้าถึงข้อมูลคู่แข่ง', 'Work for a competitor'],
  ['Text11', 'ผลงานกระทบผลประโยชน์ของผู้อื่น', 'Outcome affects others you are close to'],
  ['Text12', 'กิจกรรมวิชาชีพที่อาจขัดกัน', 'Professional activities that may conflict'],
  ['Text13', 'ค่าตอบแทนการพูดเกี่ยวกับงานมูลนิธิ', 'Payments for speaking about the work'],
  ['Text14', 'ของขวัญหรือการต้อนรับ', 'Gifts or hospitality'],
  ['Text15', 'เรื่องอื่นที่อาจกระทบความเป็นกลาง', 'Anything else affecting independence'],
];

/* ---------- Receipt vouchers (fillable PDFs) ---------- */
const receiptEN = {
  Text1: { label: L('รหัสโครงการ', 'Project code'), ask: 1, tag: 'project_code' },
  Text2: { label: L('รหัสกิจกรรม', 'Activity code'), ask: 2, tag: 'activity_code' },
  16: { label: L('วันที่ (วัน)', 'Day'), def: 'day' }, 17: { label: L('เดือน', 'Month'), def: 'month' }, 18: { label: L('ปี', 'Year'), def: 'year' },
  19: { label: L('ชื่อผู้รับเงิน', 'Name (receiver)'), ask: 3, tag: 'name' },
  20: { label: L('หน่วยงาน', 'Organization'), ask: 4, tag: 'organization' },
  21: { label: L('เลขหนังสือเดินทาง/บัตรประชาชน', 'Passport/ID #'), ask: 5, tag: 'id_card' },
  22: { label: L('ที่อยู่', 'Address'), ask: 6, tag: 'address' },
  23: { label: L('สถานที่ประชุม/เมือง/ประเทศ', 'Meeting venue/City/Country'), ask: 7 },
  24: { label: L('วัตถุประสงค์อื่น (ระบุ)', 'Other purpose (specify)') }, 25: { label: L('วัตถุประสงค์ (ต่อ)', 'Purpose (continued)') },
  26: { label: L('วันที่ออกเดินทาง', 'Date of departure') }, 27: { label: L('เวลาออกเดินทาง', 'Departure time') },
  28: { label: L('วันที่กลับถึง', 'Date of arrival') }, 29: { label: L('เวลากลับถึง', 'Arrival time') },
  30: { label: L('รวมระยะเวลา (วัน)', 'Total period (days)') }, 31: { label: L('รวมระยะเวลา (ชั่วโมง)', 'Total period (hours)') },
  56: { label: L('จำนวนเงิน (ตัวอักษร)', 'Amount in words') }, T19: { label: L('ยอดรวม', 'Total') }, T20: { label: L('รวมเป็นเงินทั้งสิ้น', 'Grand amount') },
  57: { label: L('ชื่อผู้ตรวจสอบ', 'Checked by (name)') }, 58: { label: L('ชื่อผู้จ่ายเงิน', 'Payer (name)') },
  59: { label: L('วันที่ (ผู้รับเงิน)', 'Date (receiver)') }, 60: { label: L('วันที่ (ผู้ตรวจสอบ)', 'Date (checked by)') }, 61: { label: L('วันที่ (ผู้จ่ายเงิน)', 'Date (payer)') },
  Text3: { label: L('ภาษีหัก ณ ที่จ่าย (%)', 'Withholding tax (%)') }, Text5: { label: L('ภาษีหัก ณ ที่จ่าย (บาท)', 'Withholding tax (Baht)') },
};
const receiptSignersEN = [
  { role: 'receiver', label: L('ผู้รับเงิน', 'Receiver'), nameKey: '19', rect: [86, 74, 126, 20] },
  { role: 'checker', label: L('ผู้ตรวจสอบ', 'Checked by'), nameKey: '57', rect: [282, 74, 110, 20] },
  { role: 'payer', label: L('ผู้จ่ายเงิน', 'Payer'), nameKey: '58', rect: [445, 74, 124, 20] },
];
const receiptTH = {
  Text1: { label: L('รหัสโครงการ', 'Project code'), ask: 1, tag: 'project_code' },
  Text2: { label: L('รหัสกิจกรรม', 'Activity code'), ask: 2, tag: 'activity_code' },
  16: { label: L('วันที่ (วัน)', 'Day'), def: 'day' }, 17: { label: L('เดือน', 'Month'), def: 'month' }, 18: { label: L('พ.ศ.', 'Year (B.E.)'), def: 'year' },
  19: { label: L('ชื่อผู้รับเงิน (ข้าพเจ้า)', 'Name (receiver)'), ask: 3, tag: 'name' },
  20: { label: L('หน่วยงาน', 'Organization'), ask: 4, tag: 'organization' },
  21: { label: L('เลขที่บัตรประชาชน', 'National ID no.'), ask: 5, tag: 'id_card' },
  26: { label: L('ที่อยู่', 'Address'), ask: 6, tag: 'address' },
  30: { label: L('ได้เดินทางไปปฏิบัติงาน ณ', 'Travelled to'), ask: 7 },
  31: { label: L('วัตถุประสงค์ / หัวข้อ', 'Purpose / topic') }, 32: { label: L('วัตถุประสงค์ (ต่อ)', 'Purpose (continued)') },
  33: { label: L('ออกเดินทางตั้งแต่วันที่', 'Departure date') }, 34: { label: L('เวลาออกเดินทาง', 'Departure time') },
  35: { label: L('กลับถึงที่พักวันที่', 'Return date') }, 36: { label: L('เวลากลับถึง', 'Return time') },
  37: { label: L('รวมระยะเวลา (วัน)', 'Total (days)') }, 38: { label: L('รวมระยะเวลา (ชั่วโมง)', 'Total (hours)') },
  59: { label: L('จำนวนเงิน (ตัวอักษร)', 'Amount in words') }, Text19: { label: L('ยอดรวม', 'Total') }, T20: { label: L('รวมเป็นเงินทั้งสิ้น', 'Grand amount') },
  61: { label: L('ชื่อผู้ตรวจสอบ', 'Checked by (name)') }, 62: { label: L('ชื่อผู้จ่ายเงิน', 'Payer (name)') },
  Text3: { label: L('ภาษีหัก ณ ที่จ่าย (%)', 'Withholding tax (%)') }, Text4: { label: L('ภาษีหัก ณ ที่จ่าย (บาท)', 'Withholding tax (Baht)') },
};
const receiptSignersTH = [
  { role: 'receiver', label: L('ผู้รับเงิน', 'Receiver'), nameKey: '19', rect: [54, 60, 112, 20] },
  { role: 'checker', label: L('ผู้ตรวจสอบ', 'Checked by'), nameKey: '61', rect: [233, 60, 104, 20] },
  { role: 'payer', label: L('ผู้จ่ายเงิน', 'Payer'), nameKey: '62', rect: [421, 60, 107, 20] },
];

/* ---------- Word forms: auto keys (f#, chk#, sigline_s#) renamed here ---------- */
const a1a2 = {
  f1: { key: 'a1_name', label: L('A1: ชื่อ-นามสกุล', 'A1: Full name'), tag: 'name', ask: 1 },
  f2: { key: 'a1_position', label: L('A1: ตำแหน่ง', 'A1: Position'), tag: 'position', ask: 2 },
  f3: { key: 'a1_paper', label: L('A1: ชื่อผลงานที่ส่ง', 'A1: Paper title'), ask: 3, kw: ['ชื่อผลงาน', 'ส่งผลงานเรื่อง', 'ผลงานเรื่อง', 'paper title', 'paper'] },
  f4: { key: 'a1_conf', label: L('A1: ชื่องานประชุมวิชาการ', 'A1: Conference name'), ask: 4, kw: ['ชื่องานประชุมวิชาการ', 'ชื่องานประชุม', 'ชื่องาน', 'งานประชุม', 'conference name', 'conference'] },
  f5: { key: 'a1_org', label: L('A1: จัดโดย', 'A1: Organised by'), ask: 5, kw: ['จัดโดย', 'organised by', 'organized by', 'organiser', 'organizer'] },
  f6: { key: 'a1_province', label: L('A1: จังหวัด/รัฐ', 'A1: Province/state'), kw: ['จังหวัด', 'รัฐ', 'province', 'state'] },
  f7: { key: 'a1_city', label: L('A1: เมือง', 'A1: City'), kw: ['เมือง', 'city'] },
  f8: { key: 'a1_country', label: L('A1: ประเทศ', 'A1: Country'), ask: 6, kw: ['ประเทศ', 'country'] },
  f9: { key: 'a1_from', label: L('A1: ประชุมตั้งแต่วันที่', 'A1: Conference from'), type: 'date', ask: 7, kw: ['ประชุมตั้งแต่วันที่', 'ตั้งแต่วันที่', 'ตั้งแต่', 'from'] },
  f10: { key: 'a1_to', label: L('A1: ถึงวันที่', 'A1: Conference to'), type: 'date', ask: 8, kw: ['ถึงวันที่', 'ถึง', 'until', 'to'] },
  chk11: { key: 'a1_poster', label: L('A1: นำเสนอแบบ Poster', 'A1: Poster presentation') },
  chk12: { key: 'a1_oral', label: L('A1: นำเสนอแบบ Oral', 'A1: Oral presentation') },
  f13: { key: 'a1_expected', label: L('A1: ผลที่คาดว่าจะได้รับ', 'A1: Expected outcome') },
  f15: { key: 'a1_times', label: L('A1: จำนวนครั้งที่ได้รับอนุมัติแล้วในปีนี้', 'A1: Times approved this year'), type: 'number' },
  f16: { key: 'a1_funded', label: L('A1: ค่าใช้จ่ายที่ได้รับสนับสนุนแล้ว (บาท)', 'A1: Amount already funded (Baht)'), type: 'money' },
  sigline_s17: { key: 'sigline_a1_requester', label: L('A1: ผู้ขออนุมัติ', 'A1: Requester') },
  f18: { key: 'a1_sign_name', label: L('A1: ชื่อผู้ขออนุมัติ (ใต้ลายเซ็น)', 'A1: Requester name'), mirror: 'a1_name' },
  f19: { key: 'a1_sign_date', label: L('A1: วันที่ (ผู้ขออนุมัติ)', 'A1: Date (requester)'), type: 'date', def: 'today' },
  chk20: { key: 'a1_approve', label: L('A1: อนุมัติ', 'A1: Approved') }, chk21: { key: 'a1_reject', label: L('A1: ไม่อนุมัติ', 'A1: Not approved') },
  f22: { key: 'a1_reject_reason', label: L('A1: เหตุผลที่ไม่อนุมัติ', 'A1: Reason not approved') },
  sigline_s23: { key: 'sigline_a1_approver', label: L('A1: ผู้อนุมัติ', 'A1: Approver') },
  f24: { key: 'a1_approver_name', label: L('A1: ชื่อผู้อนุมัติ', 'A1: Approver name') },
  f25: { key: 'a1_approver_date', label: L('A1: วันที่ (ผู้อนุมัติ)', 'A1: Date (approver)'), type: 'date' },
  f26: { key: 'a2_name', label: L('A2: ชื่อ-นามสกุล', 'A2: Full name'), tag: 'name' },
  f27: { key: 'a2_position', label: L('A2: ตำแหน่ง', 'A2: Position'), tag: 'position' },
  f28: { key: 'a2_conf', label: L('A2: ชื่อเรื่อง/งานประชุมวิชาการ', 'A2: Conference title'), kw: ['ชื่องานประชุมวิชาการ', 'ชื่องานประชุม', 'ชื่องาน', 'งานประชุม', 'conference name', 'conference'] },
  f29: { key: 'a2_org', label: L('A2: จัดโดย', 'A2: Organised by'), kw: ['จัดโดย', 'organised by', 'organized by', 'organiser', 'organizer'] },
  f30: { key: 'a2_city', label: L('A2: จังหวัด/เมือง', 'A2: Province/city'), kw: ['จังหวัด', 'เมือง', 'city', 'province'] },
  f31: { key: 'a2_country', label: L('A2: ประเทศ', 'A2: Country'), kw: ['ประเทศ', 'country'] },
  f32: { key: 'a2_from', label: L('A2: ประชุมตั้งแต่วันที่', 'A2: Conference from'), type: 'date', kw: ['ประชุมตั้งแต่วันที่', 'ตั้งแต่วันที่', 'ตั้งแต่', 'from'] },
  f33: { key: 'a2_to', label: L('A2: ถึงวันที่', 'A2: Conference to'), type: 'date', kw: ['ถึงวันที่', 'ถึง', 'until', 'to'] },
  chk34: { key: 'a2_poster', label: L('A2: นำเสนอแบบ Poster', 'A2: Poster') }, chk35: { key: 'a2_oral', label: L('A2: นำเสนอแบบ Oral', 'A2: Oral') },
  chk36: { key: 'a2_attend', label: L('A2: เข้าร่วมโดยไม่นำเสนอผลงาน', 'A2: Attend only') },
  f37: { key: 'a2_expected', label: L('A2: ผลที่คาดว่าจะได้รับ', 'A2: Expected outcome') },
  chk39: { key: 'a2_nofund', label: L('A2: ไม่ขอเบิกเงินสนับสนุน', 'A2: No funding requested') },
  chk40: { key: 'a2_fund', label: L('A2: ขอเบิกเงินสนับสนุน', 'A2: Funding requested') },
  f41: { key: 'a2_total', label: L('A2: ค่าใช้จ่ายทั้งหมดประมาณ (บาท)', 'A2: Estimated total cost (Baht)'), type: 'money' },
  chk42: { key: 'a2_c_reg', label: L('A2: ค่าลงทะเบียน', 'A2: Registration fee') }, f43: { key: 'a2_reg', label: L('A2: ค่าอบรม/ลงทะเบียน (บาท)', 'A2: Registration (Baht)'), type: 'money' },
  chk44: { key: 'a2_c_travel', label: L('A2: ค่าเดินทาง', 'A2: Travel') }, f45: { key: 'a2_travel', label: L('A2: ค่าเดินทาง (บาท)', 'A2: Travel (Baht)'), type: 'money' },
  chk46: { key: 'a2_c_hotel', label: L('A2: ค่าที่พัก', 'A2: Accommodation') }, f47: { key: 'a2_hotel', label: L('A2: ค่าที่พัก (บาท)', 'A2: Accommodation (Baht)'), type: 'money' },
  chk48: { key: 'a2_c_perdiem', label: L('A2: ค่าเบี้ยเลี้ยง', 'A2: Per diem') }, f49: { key: 'a2_perdiem', label: L('A2: ค่าเบี้ยเลี้ยง (บาท)', 'A2: Per diem (Baht)'), type: 'money' },
  chk50: { key: 'a2_c_other', label: L('A2: ค่าใช้จ่ายอื่นๆ', 'A2: Other cost') }, f51: { key: 'a2_other_desc', label: L('A2: ค่าใช้จ่ายอื่นๆ (ระบุ)', 'A2: Other (specify)') },
  f52: { key: 'a2_other', label: L('A2: ค่าใช้จ่ายอื่นๆ (บาท)', 'A2: Other (Baht)'), type: 'money' },
  f53: { key: 'a2_times', label: L('A2: จำนวนครั้งที่ได้รับอนุมัติแล้วในปีนี้', 'A2: Times approved this year'), type: 'number' },
  f54: { key: 'a2_funded', label: L('A2: ค่าใช้จ่ายที่ได้รับสนับสนุนแล้ว (บาท)', 'A2: Amount already funded (Baht)'), type: 'money' },
};
const ar1 = {
  f1: { key: 'name', label: L('ชื่อ-นามสกุล', 'Full name'), tag: 'name', ask: 1 },
  f2: { key: 'position', label: L('ตำแหน่ง', 'Position'), tag: 'position', ask: 2 },
  chk3: { key: 'c_conf', label: L('การเข้าร่วมประชุมวิชาการ', 'Academic conference') }, f4: { key: 'conf', label: L('ชื่อการประชุมวิชาการ', 'Conference name'), ask: 3, kw: ['ชื่องานประชุมวิชาการ', 'ชื่องานประชุม', 'ชื่องาน', 'งานประชุม', 'ประชุมวิชาการ', 'conference'] },
  chk5: { key: 'c_training', label: L('การลาฝึกอบรม', 'Training') }, f6: { key: 'training', label: L('ชื่อการฝึกอบรม', 'Training name'), kw: ['ฝึกอบรม', 'อบรม', 'training'] },
  chk7: { key: 'c_research', label: L('การปฏิบัติการวิจัย', 'Research practice') }, f8: { key: 'research', label: L('การปฏิบัติการวิจัย (ระบุ)', 'Research practice (detail)') },
  chk9: { key: 'c_visit', label: L('การศึกษาดูงาน', 'Study visit') }, f10: { key: 'visit', label: L('การศึกษาดูงาน (ระบุ)', 'Study visit (detail)'), kw: ['ศึกษาดูงาน', 'ดูงาน', 'study visit'] },
  f11: { key: 'org', label: L('จัดโดย', 'Organised by'), ask: 4, kw: ['จัดโดย', 'organised by', 'organized by'] },
  f12: { key: 'city', label: L('จังหวัด/เมือง', 'Province/city'), ask: 5 },
  f13: { key: 'country', label: L('ประเทศ', 'Country'), ask: 6 },
  f14: { key: 'from', label: L('ตั้งแต่วันที่', 'From'), type: 'date', ask: 7, kw: ['ตั้งแต่วันที่', 'ตั้งแต่', 'from'] },
  f15: { key: 'to', label: L('ถึงวันที่', 'To'), type: 'date', ask: 8, kw: ['ถึงวันที่', 'ถึง', 'until', 'to'] },
  f16: { key: 'activities', label: L('กิจกรรม/หัวข้อที่เข้าร่วมในแต่ละวัน', 'Activities attended each day'), type: 'textarea', ask: 9 },
  f17: { key: 'benefits', label: L('ทักษะหรือประโยชน์ที่ได้รับ', 'Skills or benefits gained'), type: 'textarea', ask: 10 },
  sigline_s19: { key: 'sigline_reporter', label: L('ผู้รายงาน', 'Reporter') },
  f20: { key: 'sign_name', label: L('ชื่อผู้รายงาน (ใต้ลายเซ็น)', 'Reporter name'), mirror: 'name' },
  f21: { key: 'sign_date', label: L('วันที่ลงนาม', 'Date signed'), type: 'date', def: 'today' },
};

/* ---------- Excel forms ---------- */
const borrowCells = {
  'ใบยืมเงิน Borrow Money Form': {
    E6: { key: 'project_code', label: L('รหัสโครงการ', 'Project code'), tag: 'project_code', ask: 1 },
    X6: { key: 'activity_code', label: L('รหัสกิจกรรม', 'Activity code'), tag: 'activity_code', ask: 2 },
    E9: { key: 'project_name', label: L('โครงการ (ชื่อย่อ)', 'Project name'), ask: 3 },
    D15: { key: 'name', label: L('ชื่อผู้ยืม (ข้าพเจ้า)', 'Borrower name'), tag: 'name', ask: 4 },
    W15: { key: 'position', label: L('ตำแหน่ง', 'Position'), tag: 'position', ask: 5 },
    G18: { key: 'purpose', label: L('ขออนุมัติยืมเงินเพื่อ', 'Purpose of the advance'), ask: 6 },
    G21: { key: 'from', label: L('ระหว่างวันที่', 'From date'), type: 'date', ask: 7, kw: ['ตั้งแต่วันที่', 'ตั้งแต่', 'from'] },
    P21: { key: 'to', label: L('ถึงวันที่', 'To date'), type: 'date', ask: 8, kw: ['ถึงวันที่', 'ถึง', 'until', 'to'] },
    AA21: { key: 'time', label: L('เวลา', 'Time') },
    G24: { key: 'place', label: L('สถานที่ดำเนินการ', 'Place'), ask: 9 },
    G27: { key: 'need_date', label: L('ต้องการรับเงินยืมในวันที่', 'Money needed by'), type: 'date', ask: 10 },
    E39: { key: 'd_honor', label: L('ค่าตอบแทน (รายละเอียด)', 'Honorarium (details)') }, AB39: { key: 'a_honor', label: L('ค่าตอบแทน (บาท)', 'Honorarium (Baht)'), type: 'money' },
    E42: { key: 'd_perdiem', label: L('ค่าเบี้ยเลี้ยง (รายละเอียด)', 'Per diem (details)') }, AB42: { key: 'a_perdiem', label: L('ค่าเบี้ยเลี้ยง (บาท)', 'Per diem (Baht)'), type: 'money' },
    E44: { key: 'd_air', label: L('ค่าเครื่องบิน (รายละเอียด)', 'Airfare (details)') }, AB44: { key: 'a_air', label: L('ค่าเครื่องบิน (บาท)', 'Airfare (Baht)'), type: 'money' },
    E46: { key: 'd_travel', label: L('ค่าเดินทาง (รายละเอียด)', 'Travel (details)') }, AB46: { key: 'a_travel', label: L('ค่าเดินทาง (บาท)', 'Travel (Baht)'), type: 'money' },
    E48: { key: 'd_hotel', label: L('ค่าที่พัก (รายละเอียด)', 'Accommodation (details)') }, AB48: { key: 'a_hotel', label: L('ค่าที่พัก (บาท)', 'Accommodation (Baht)'), type: 'money' },
    E50: { key: 'd_food', label: L('ค่าอาหารและเครื่องดื่ม (รายละเอียด)', 'Food & beverage (details)') }, AB50: { key: 'a_food', label: L('ค่าอาหารและเครื่องดื่ม (บาท)', 'Food & beverage (Baht)'), type: 'money' },
    E52: { key: 'd_other1', label: L('ค่าใช้จ่ายอื่นๆ 1 (รายละเอียด)', 'Other 1 (details)') }, AB52: { key: 'a_other1', label: L('ค่าใช้จ่ายอื่นๆ 1 (บาท)', 'Other 1 (Baht)'), type: 'money' },
    E53: { key: 'd_other2', label: L('ค่าใช้จ่ายอื่นๆ 2 (รายละเอียด)', 'Other 2 (details)') }, AB53: { key: 'a_other2', label: L('ค่าใช้จ่ายอื่นๆ 2 (บาท)', 'Other 2 (Baht)'), type: 'money' },
    E54: { key: 'd_other3', label: L('ค่าใช้จ่ายอื่นๆ 3 (รายละเอียด)', 'Other 3 (details)') }, AB54: { key: 'a_other3', label: L('ค่าใช้จ่ายอื่นๆ 3 (บาท)', 'Other 3 (Baht)'), type: 'money' },
    D62: { key: 'sign_name', label: L('ชื่อผู้ยืม (ใต้ลายเซ็น)', 'Borrower name (signature)'), mirror: 'name' },
  },
};
// HPIE form has no colour coding: list its input cells explicitly.
const hpieRows = (sheet, rows, cols) => Object.fromEntries(rows.flatMap(r => cols.map(([c, k, th, en, type]) => [`${c}${r}`, { key: `${k}${r}`, label: L(`${th} แถว ${r}`, `${en} row ${r}`), table: true, type }])));
const hpieCells = {
  'งน.11-ใบรับรองการจ่ายเงิน': {
    K5: { key: 'cert_date', label: L('งน.11: วันที่', 'Form 11: Date'), type: 'date', def: 'today', ask: 4 },
    C7: { key: 'agreement_no', label: L('ข้อตกลงเลขที่', 'Agreement no.'), ask: 1 },
    C8: { key: 'project', label: L('โครงการ', 'Project'), ask: 2 },
    ...hpieRows('11', [11, 12, 13, 14, 15, 16, 17, 18, 19, 20], [['A', 'c_date', 'วัน เดือน ปี', 'Date'], ['C', 'c_desc', 'รายละเอียดการจ่าย', 'Payment detail'], ['I', 'c_amt', 'จำนวนเงิน', 'Amount', 'money'], ['K', 'c_note', 'หมายเหตุ', 'Note']]),
    C24: { key: 'cert_name', label: L('งน.11: ชื่อผู้รับรอง (ข้าพเจ้า)', 'Form 11: Certifier name'), tag: 'name', ask: 3 },
    J24: { key: 'cert_position', label: L('งน.11: ตำแหน่ง', 'Form 11: Position'), tag: 'position' },
    H31: { key: 'cert_sign_name', label: L('งน.11: ชื่อใต้ลายเซ็น', 'Form 11: Name under signature'), mirror: 'cert_name' },
  },
  'งน.1 - ใบสำคัญรับเงิน(ต้นฉบับ)': {
    J5: { key: 'agreement_no_1', label: L('งน.1: ข้อตกลงเลขที่', 'Form 1: Agreement no.'), mirror: 'agreement_no' },
    I6: { key: 'rcpt_date', label: L('งน.1: วันที่', 'Form 1: Date'), type: 'date', def: 'today' },
    B8: { key: 'rcpt_name', label: L('งน.1: ชื่อผู้รับเงิน (ข้าพเจ้า)', 'Form 1: Receiver name'), tag: 'name' },
    I8: { key: 'rcpt_id', label: L('งน.1: เลขประจำตัวบัตรประชาชน', 'Form 1: National ID no.'), tag: 'id_card' },
    B9: { key: 'rcpt_address', label: L('งน.1: ที่อยู่', 'Form 1: Address'), tag: 'address' },
    C11: { key: 'rcpt_project', label: L('งน.1: ได้รับเงินจากโครงการ', 'Form 1: Project'), mirror: 'project' },
    ...hpieRows('1', [14, 15, 16, 17, 18, 19, 20, 21, 22, 23], [['A', 'r_no', 'ลำดับ', 'No.'], ['B', 'r_desc', 'รายการ', 'Item'], ['J', 'r_amt', 'จำนวนเงิน', 'Amount', 'money']]),
    B31: { key: 'rcpt_sign_name', label: L('งน.1: ชื่อผู้รับเงิน (ใต้ลายเซ็น)', 'Form 1: Receiver name (signature)'), mirror: 'rcpt_name' },
    H31: { key: 'payer_name', label: L('งน.1: ชื่อผู้จ่ายเงิน', 'Form 1: Payer name') },
  },
};

export default [
  { id: 'leave-adm', file: 'FM-ADM-03-05 บันทึกการลา มอบหมายงานระหว่างลา (กรณีลาติดต่อกันตั้งแต่ 5 วันขึ้นไปโดยนับรวมวันหยุด).pdf', kind: 'pdf', cat: 'hr', docLang: 'th', dateFmt: 'dmy-be',
    title: L('ใบบันทึกการลา / มอบหมายงานระหว่างลา (FM-ADM-03-05)', 'Leave form & work handover (FM-ADM-03-05)'),
    desc: L('กรณีลาติดต่อกันตั้งแต่ 5 วันขึ้นไปโดยนับรวมวันหยุด ระบุงานที่มอบหมายและผู้รับผิดชอบแทน', 'Consecutive leave of 5+ days incl. holidays — hand over unfinished work'),
    overlay: leaveFields(), signers: leaveSigners },
  { id: 'leave-hr', file: 'Leave Form Consecutive leave from 5 days up, including holidays.pdf', kind: 'pdf', cat: 'hr', docLang: 'en', dateFmt: 'dmy',
    title: L('ใบบันทึกการลา / มอบหมายงานระหว่างลา (FM-HR-004-4)', 'Leave form & work handover (FM-HR-004-4)'),
    desc: L('แบบฟอร์มลาติดต่อกัน 5 วันขึ้นไป ฉบับ FM-HR-004-4', 'Consecutive leave form, FM-HR-004-4 edition'),
    overlay: leaveFields(), signers: leaveSigners },
  { id: 'coi-th', file: '(TH) COI declaration_2024_Revised.pdf', kind: 'pdf', cat: 'compliance', docLang: 'th', dateFmt: 'dmy-be',
    title: L('แบบแถลงการขัดกันของผลประโยชน์ประจำปี (COI) — ฉบับไทย', 'Annual conflict of interest declaration (COI) — Thai'),
    desc: L('ตอบ ใช่/ไม่ใช่ พร้อมรายละเอียด 15 หัวข้อ แล้วลงนามแถลง', 'Answer 15 yes/no questions with details, then sign'),
    coi: COI_TH, labels: { Text17: { key: 'name', label: L('ชื่อ-นามสกุล', 'Full name'), tag: 'name', ask: 1 }, Text18: { key: 'date', label: L('วันที่', 'Date'), type: 'date', def: 'today' } },
    signers: [{ role: 'declarant', label: L('ผู้แถลง', 'Declarant'), nameKey: 'name', field: 'Text16' }] },
  { id: 'coi-en', file: '(EN) COI declaration_2024_Revised.pdf', kind: 'pdf', cat: 'compliance', docLang: 'en', dateFmt: 'mdy-short',
    title: L('แบบแถลงการขัดกันของผลประโยชน์ประจำปี (COI) — ฉบับอังกฤษ', 'Annual conflict of interest declaration (COI) — English'),
    desc: L('ฉบับภาษาอังกฤษ 16 หัวข้อ', 'English edition, 16 questions'),
    coi: COI_EN, labels: { Text26: { key: 'name', label: L('ชื่อ-นามสกุล', 'Elaborate name'), tag: 'name', ask: 1 }, Date25_af_date: { key: 'date', label: L('วันที่', 'Date'), type: 'date', def: 'today' } },
    signers: [{ role: 'declarant', label: L('ผู้แถลง', 'Declarant'), nameKey: 'name', field: 'Text1' }] },
  { id: 'receipt-en', file: 'Receipt voucher-EN.pdf', kind: 'pdf', cat: 'finance', docLang: 'en',
    title: L('ใบสำคัญรับเงิน (Receipt voucher) — ฉบับอังกฤษ', 'Receipt voucher / Certificate of receipt — English'),
    desc: L('FM-FIN-01-06 Rev.02 คำนวณยอดรวมและจำนวนเงินตัวอักษรให้อัตโนมัติ', 'FM-FIN-01-06 Rev.02 — totals and amount in words are calculated for you'),
    labels: receiptEN, signers: receiptSignersEN },
  { id: 'receipt-en-pcht', file: 'Receipt voucher-EN-PCHT-2566-003.pdf', kind: 'pdf', cat: 'finance', docLang: 'en',
    title: L('ใบสำคัญรับเงิน (Receipt voucher) — อังกฤษ PCHT-2566-003', 'Receipt voucher — English (PCHT-2566-003)'),
    desc: L('ฉบับสำหรับโครงการ PCHT-2566-003', 'Edition for project PCHT-2566-003'),
    labels: receiptEN, signers: receiptSignersEN },
  { id: 'receipt-th-pcht', file: 'Receipt voucher-TH-PCHT-2566-003.pdf', kind: 'pdf', cat: 'finance', docLang: 'th',
    title: L('ใบสำคัญรับเงิน / ใบรับรองแทนใบเสร็จ — ไทย PCHT-2566-003', 'Receipt voucher — Thai (PCHT-2566-003)'),
    desc: L('FM-FIN-01-06 Rev.02 คำนวณยอดรวมและจำนวนเงินตัวอักษรภาษาไทยให้อัตโนมัติ', 'Thai edition — totals and Thai amount in words calculated'),
    labels: receiptTH, signers: receiptSignersTH },
  { id: 'a1a2', file: 'A1 A2 request form for Academic Conference participation [TH].docx', kind: 'docx', cat: 'academic', docLang: 'th', dateFmt: 'dmy-be',
    title: L('แบบฟอร์มขออนุมัติส่งผลงาน/เข้าร่วมประชุมวิชาการ (A1/A2)', 'Academic conference request forms (A1/A2)'),
    desc: L('A1 ขออนุมัติส่งผลงาน และ A2 ขออนุมัติเข้าร่วมประชุมวิชาการ พร้อมประมาณการค่าใช้จ่าย', 'A1 submit a paper, A2 attend a conference with cost estimate'),
    map: a1a2 },
  { id: 'ar1', file: 'AR1_แบบรายงานผลการประชุม ดูงาน ฝึกอบรม.docx', kind: 'docx', cat: 'academic', docLang: 'th', dateFmt: 'dmy-be',
    title: L('แบบรายงานผลการประชุม อบรม ปฏิบัติการวิจัย ดูงาน (AR1)', 'Report on a conference, training or study visit (AR1)'),
    desc: L('รายงานกิจกรรมที่เข้าร่วมและประโยชน์ที่ได้รับ ต่อคณะกรรมการวิชาการ', 'Report activities and benefits to the academic committee'),
    map: ar1 },
  { id: 'borrow', file: 'ใบยืมเงินทดรองจ่าย (Borrow - Return Money Form).xlsx', kind: 'xlsx', cat: 'finance', docLang: 'th',
    title: L('ใบยืมเงินทดรองจ่าย และสรุปคืนเงินยืม', 'Borrow money & return money form'),
    desc: L('กรอกช่องสีฟ้า ระบบคำนวณยอดรวมตามสูตรในไฟล์ (ชีทคำแนะนำ / ใบยืมเงิน / สรุปคืนเงินยืม)', 'Fill the blue cells — totals follow the workbook formulas'),
    cells: borrowCells, tableSheets: ['สรุปคืนเงินยืมReturn money form'], hideSheets: [] },
  { id: 'hpie', file: 'ใบสำคัญรับเงิน หรือ ใบรับรองการจ่ายเงิน HPIE (สสส).xlsx', kind: 'xlsx', cat: 'finance', docLang: 'th',
    title: L('ใบรับรองการจ่ายเงิน (งน.11) / ใบสำคัญรับเงิน (งน.1) — HPIE (สสส.)', 'Payment certificate (Form 11) / Receipt (Form 1) — HPIE (ThaiHealth)'),
    desc: L('ใช้เมื่อไม่สามารถขอใบเสร็จได้ คำนวณยอดรวมและจำนวนเงินตัวอักษร (BAHTTEXT) ให้', 'When no receipt is available — totals and Thai amount in words calculated'),
    cells: hpieCells, manualCells: true,
    signers: [
      { role: 'certifier', label: L('ผู้รับรอง (งน.11)', 'Certifier (Form 11)'), nameKey: 'cert_name', sheet: 'งน.11-ใบรับรองการจ่ายเงิน', range: 'H30:K30' },
      { role: 'receiver', label: L('ผู้รับเงิน (งน.1)', 'Receiver (Form 1)'), nameKey: 'rcpt_name', sheet: 'งน.1 - ใบสำคัญรับเงิน(ต้นฉบับ)', range: 'B30:E30' },
      { role: 'payer', label: L('ผู้จ่ายเงิน (งน.1)', 'Payer (Form 1)'), nameKey: 'payer_name', sheet: 'งน.1 - ใบสำคัญรับเงิน(ต้นฉบับ)', range: 'H30:K30' },
    ] },
];

export const CATEGORIES = [
  { id: 'contract', th: 'สัญญา/NDA', en: 'Contracts / NDA' },
  { id: 'hr', th: 'บุคคล/การลา', en: 'HR / Leave' },
  { id: 'finance', th: 'การเงิน', en: 'Finance' },
  { id: 'academic', th: 'วิชาการ', en: 'Academic' },
  { id: 'compliance', th: 'ผลประโยชน์ทับซ้อน', en: 'Conflict of interest' },
];
