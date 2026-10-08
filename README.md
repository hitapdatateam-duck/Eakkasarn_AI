# HITAP Docs

Free, browser-only drafting for HITAP's forms: describe what you need in the chat, answer a few
questions or type straight into the document, sign, and download the file in its original format.

## Run

```bash
python3 -m http.server 8791
```

Then open http://localhost:8791 (double-clicking `index.html` also works on a machine that has
TH SarabunPSK installed).

## Templates

| Kind | How it is filled |
|------|------------------|
| Word (`.docx`) | dotted lines `……` and check-box symbols become fields; values replace them in the Word file |
| PDF with form fields | the PDF's own fields (and its calculation scripts) are filled; TH Sarabun is embedded |
| PDF without form fields | values are drawn at coordinates listed in `build/templates.config.mjs` |
| Excel (`.xlsx`) | the light-blue input cells (or cells listed in the config) are written; formulas recalculate |

NDA (Thai/English) templates are built by `build_templates.py`; everything in `templates/` is built by:

```bash
npm install
npm run build
```

To add a form: copy the file into `templates/`, add an entry to `build/templates.config.mjs`
(title, category, and labels for the fields the assistant should ask about), then run `npm run build`.
