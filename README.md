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

## Admin back-office

Admins (rows in `public.admins`, by email) sign in with email + password (or the password-only gate at `#/admin`) and get an **ผู้ดูแลระบบ / Admin**
menu (`#/admin`):

- **Templates** — hide/show any template; delete uploaded ones (built-in templates live in this repo, so they
  can only be hidden).
- **Add template** — upload a `.docx`, `.pdf` or `.xlsx`; it is analysed in the browser with the same
  `build/compile.mjs` the build uses. Rename fields, choose which ones the assistant asks, link profile items,
  click on flat PDFs to place text/signature boxes, list Excel cells. Files go to the Supabase Storage bucket
  `templates`, metadata to `public.templates`.
- **Admins** — add or remove admin emails.

RLS: anyone can read `templates`; only `public.is_admin()` can write it or upload to the bucket.

## Theme

Light/dark follows the operating system; the button next to the language switch cycles
system → light → dark (remembered per browser).
