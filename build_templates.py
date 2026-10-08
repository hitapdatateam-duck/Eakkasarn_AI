"""Convert the HITAP NDA .docx files into token templates ({{key}}) for the web app.

Each yellow-highlighted [placeholder] becomes one run containing {{key}} (highlight kept,
so unfilled fields stay visible). Blank name lines "(......)" become {{sig_*}} and the
dotted "sign here" lines in the signature table become {{sigline_<role>_<tabs>}}; the dotted
contract-number line in the header becomes {{contract_no}}.
Output: templates.js with base64 of the tokenised docx and the original docx.
"""
import base64, io, json, re, sys, zipfile
from pathlib import Path

SRC = Path(sys.argv[1] if len(sys.argv) > 1 else "/Users/pakawat/Desktop/NDA")
OUT = Path(__file__).parent / "templates.js"

# (placeholder text, key) in document order; repeated placeholders map in sequence.
MAPS = {
    "th": ("TH_NDA_HITAPer.docx", [
        ("[วันที่ทำสัญญา]", "agreement_date"),
        ("[ชื่อผู้รับข้อมูล]", "recipient_name"),
        ("[หมายเลขบัตรประชาชนของผู้รับข้อมูล]", "recipient_id"),
        ("[ที่อยู่ของผู้รับข้อมูลตามบัตรประชาชน]", "recipient_address"),
        ("[วันที่ในสัญญาจ้าง]", "employment_date"),
        ("[ชื่อตำแหน่ง]", "position"),
        ("[รายละเอียดข้อมูล]", "information"),
        ("[วัตถุประสงค์ในการขอข้อมูล]", "objective"),
        ("[รายละเอียดข้อมูล]", "information"),
        ("[ชื่อโครงการวิจัย]", "project"),
    ], ["sig_recipient", "sig_witness1", "sig_witness2"]),
    "en": ("EN_NDA_HITAPer.docx", [
        ("[Date]", "agreement_date"),
        ("[Name of Contractor/Recipient]", "recipient_name"),
        ("[ID/Passport number]", "recipient_id"),
        ("[Address of the Contractor/Recipient]", "recipient_address"),
        ("[xxx]", "employment_date"),
        ("[xxx]", "position"),
        ("[Project/Information]", "information"),
        ("[Objective]", "objective"),
        ("[Project/Information]", "information"),
        ("[Project]", "project"),
    ], ["sig_recipient", "sig_witness1", "sig_witness2"]),
}

RUN = re.compile(r"<w:r[ >].*?</w:r>", re.S)
TXT = re.compile(r"(<w:t(?: [^>]*)?>)([^<]*)(</w:t>)")


def run_text(r):
    return "".join(m.group(2) for m in TXT.finditer(r))


def set_run_text(r, text):
    """Replace all <w:t> in run with a single one holding text (tabs removed)."""
    r = re.sub(r"<w:tab/>", "", r)
    first = [True]
    def rep(m):
        if first[0]:
            first[0] = False
            return '<w:t xml:space="preserve">' + text + "</w:t>"
        return ""
    r2 = TXT.sub(rep, r)
    if first[0]:  # run had no <w:t>
        r2 = r2.replace("</w:r>", '<w:t xml:space="preserve">' + text + "</w:t></w:r>")
    return r2


def tokenise_para(p, ph, token):
    """Replace first occurrence of placeholder ph (may span runs) with token."""
    runs = list(RUN.finditer(p))
    texts = [run_text(m.group()) for m in runs]
    full = "".join(texts)
    s = full.find(ph)
    if s < 0:
        return None
    e = s + len(ph)
    new_runs, pos = [], 0
    for m, t in zip(runs, texts):
        a, b = pos, pos + len(t)
        pos = b
        r = m.group()
        if b <= s or a >= e or not t:
            new_runs.append(r if not (a >= s and b <= e and t) else set_run_text(r, ""))
            continue
        before = t[: max(0, s - a)] if a < s else ""
        after = t[e - a:] if b > e else ""
        mid = token if a <= s < b or (a > s and not any(token in x for x in new_runs)) else ""
        new_runs.append(set_run_text(r, before + mid + after))
    out, last = [], 0
    for m, nr in zip(runs, new_runs):
        out.append(p[last:m.start()]); out.append(nr); last = m.end()
    out.append(p[last:])
    return "".join(out)


def sig_para(p, key):
    """'(' + dotted tabs + ')' -> '(' run, token run with dotted underline, ')' run."""
    runs = list(RUN.finditer(p))
    rpr = re.search(r"<w:rPr>.*?</w:rPr>", runs[1].group(), re.S).group()
    token_run = "<w:r>" + rpr + '<w:t xml:space="preserve">{{' + key + "}}</w:t></w:r>"
    return p[: runs[1].start()] + token_run + p[runs[-1].start():]


SIGNERS = ["discloser", "recipient", "witness1", "witness2"]


def sign_line_para(p, role):
    """Dotted tab runs -> one run '{{sigline_<role>_<tabs>}}' keeping the dotted underline."""
    runs = list(RUN.finditer(p))
    tabs = sum(m.group().count("<w:tab/>") for m in runs)
    rpr = re.search(r"<w:rPr>.*?</w:rPr>", runs[0].group(), re.S).group()
    token_run = "<w:r>" + rpr + '<w:t xml:space="preserve">{{sigline_%s_%d}}</w:t></w:r>' % (role, tabs)
    return p[: runs[0].start()] + token_run + p[runs[-1].end():]


def build(lang):
    fname, maps, sigs = MAPS[lang]
    raw = (SRC / fname).read_bytes()
    zin = zipfile.ZipFile(io.BytesIO(raw))
    xml = zin.read("word/document.xml").decode("utf8")
    paras = list(re.finditer(r"<w:p[ >].*?</w:p>", xml, re.S))
    texts = ["".join(run_text(r.group()) for r in RUN.finditer(m.group())) for m in paras]
    new = [m.group() for m in paras]
    labels = {}
    i = 0
    for ph, key in maps:
        while i < len(new):
            r = tokenise_para(new[i], ph, "{{" + key + "}}")
            if r is not None:
                new[i] = r; labels.setdefault(key, ph); break
            i += 1
        else:
            raise SystemExit(f"{lang}: placeholder not found {ph}")
    # Dotted "sign here" lines inside the signature table, in table order.
    tbl = re.search(r"<w:tbl>.*?</w:tbl>", xml, re.S)
    lines = [k for k, m in enumerate(paras)
             if tbl.start() < m.start() < tbl.end() and texts[k].strip() == ""
             and re.search(r"<w:u [^>]*/>", m.group()) and "<w:tab/>" in m.group()]
    assert len(lines) == len(SIGNERS), (lang, lines)
    for k, role in zip(lines, SIGNERS):
        new[k] = sign_line_para(new[k], role)
    blanks = [k for k, t in enumerate(texts) if t.strip() == "()"]
    assert len(blanks) == len(sigs), (lang, blanks)
    for k, key in zip(blanks, sigs):
        new[k] = sig_para(new[k], key)
    out, last = [], 0
    for m, n in zip(paras, new):
        out.append(xml[last:m.start()]); out.append(n); last = m.end()
    out.append(xml[last:])
    xml2 = "".join(out)
    # Drop runs emptied by tokenising, and un-highlight leftover whitespace-only runs.
    def tidy(m):
        r = m.group()
        if "{{" in r or "<w:tab/>" in r or "<w:br" in r or "<w:drawing" in r or "<w:fldChar" in r or "<w:instrText" in r:
            return r
        t = run_text(r)
        if "<w:t" in r and t == "":
            return ""
        if t.strip() == "":
            return r.replace('<w:highlight w:val="yellow"/>', "")
        return r
    xml2 = RUN.sub(tidy, xml2)
    left = re.findall(r"\[[^\]<]{1,60}\]", "".join(run_text(r.group()) for r in RUN.finditer(xml2)))
    assert not left, (lang, left)
    # Contract/reference number: the dotted run after "เลขที่สัญญา" / "Reference Number:" in the header.
    hdr = zin.read("word/header1.xml").decode("utf8")
    def contract_no(m):
        r = m.group()
        t = run_text(r)
        d = re.fullmatch(r"(\s*)(\.{5,})\s*", t)
        if not d or "contract_no" in labels:
            return r
        labels["contract_no"] = d.group(2)
        return set_run_text(r, d.group(1) + "{{contract_no}}")
    hdr = RUN.sub(contract_no, hdr)
    assert "contract_no" in labels, (lang, "contract number dots not found in header")
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = {"word/document.xml": xml2, "word/header1.xml": hdr}.get(item.filename)
            zout.writestr(item, data.encode("utf8") if data is not None else zin.read(item))
    return {
        "file": fname,
        "labels": labels,
        "template": base64.b64encode(buf.getvalue()).decode(),
        "original": base64.b64encode(raw).decode(),
    }


data = {lang: build(lang) for lang in MAPS}
OUT.write_text("// Generated by build_templates.py — do not edit by hand.\nwindow.NDA_TEMPLATES = "
               + json.dumps(data, ensure_ascii=False) + ";\n", encoding="utf8")
print("wrote", OUT, OUT.stat().st_size, "bytes")
for l, d in data.items():
    print(l, d["labels"])
