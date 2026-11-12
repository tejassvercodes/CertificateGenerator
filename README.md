# Certificate Generator

A browser-based tool for bulk-generating certificates from pre-loaded Word (DOCX) templates. Pick one of the four journals (**IJARCCE**, **IARJSET**, **IJIREEICE**, **IMRJR** — internally keyed as CCE/SET/ICE/RJR), paste in a block of recipient data, and download the result as a combined file, individual files, or a single group certificate (via docxtemplater/PizZip). A companion local server (`server.js`) can convert a generated DOCX to a pixel-faithful PDF using Microsoft Word itself (COM automation via `convert-docx-to-pdf.ps1`).

There is no upload step — the templates ship with the app in `templates/`.

## Running it

This is a static site with no build step or backend, but it **must be served over HTTP** — templates are loaded with `fetch()`, which browsers block on `file://`, so opening `index.html` directly will not work.

```
node server.js
```
then visit http://localhost:8080. Any static file server works equally well (`npx serve .`, `python -m http.server`).

## Templates

The four journal templates live in `templates/` (`CCE.docx`, `SET.docx`, `ICE.docx`, `RJR.docx`) and are fetched on demand — only the template you click is downloaded, and it's cached for the rest of the session. Each journal's logo lives in `logos/` (`CCE.jpg`, `SET.jpg`, `ICE.jpg`, `RJR.jpg`).

The `TEMPLATES` array at the top of `app.js` is the single source of truth for the picker — `index.html` just holds an empty `<div id="template-grid">`, and `renderTemplateGrid()` builds the four cards (logo + full journal name) from the array on load.

- **To update a template or logo**: replace the file in `templates/` or `logos/`, keeping the same name.
- **To add a journal**: drop the `.docx` in `templates/` and the logo in `logos/`, then add one `{ key, fullName, file, logo }` entry to `TEMPLATES` in `app.js` — no HTML changes needed.

Placeholder tags are matched case- and spacing-insensitively (`{Paper Title}`, `{paper_title}`, and `{PAPERTITLE}` are equivalent), so a template with an inconsistently-spelled tag still fills correctly — the RJR template, for instance, spells its name tag `{NAMe}`. An unrecognized tag renders as empty text rather than breaking the export.

## How it works

1. **Choose a journal template** — IJARCCE, IARJSET, IJIREEICE, or IMRJR. All four contain the same placeholder tags: `{NAME}`, `{Designation}`, `{Paper Title}`, `{DOI}`, `{vol}`, `{issue}`, `{year}`, and `{month}` (docxtemplater performs the substitution). Each DOI line is wrapped in `{#hasDoi}...{/hasDoi}` so the whole line, including its "DOI:" label, disappears in Without DOI mode; **a replacement template must keep that wrapper**. You can switch templates at any point without reloading; the next export uses the newly selected one.

2. **Paste mapping data** into the text box, one category per line. Each line is parsed by `parseDataInput`/`parseMappingLine` in `app.js`:
   - **Names**: a comma-separated list of names, each with a numeric suffix indicating which certificate(s) it maps to, e.g. `John Smith1, Jane Doe2`. A range suffix like `3-5` applies the same value to certificates 3 through 5, and a comma-separated list like `1,2,3` (or a mix, e.g. `1,3-5`) applies to exactly those certificates. The name line determines how many certificates are generated (the highest suffix number). Superscript numerals (e.g. `Name¹²`) are also recognized as suffix digits.
   - **Designations**: same comma+suffix format, e.g. `Professor, Example University1-2`. Lines are heuristically classified as a designation if they contain keywords such as "student", "professor", "department", "college", "university", etc. Two affiliations for one author are both kept.
   - **Paper Title**: a line without keyword matches is treated as a paper title and applied to every certificate. A title wrapped onto two lines is rejoined. If there are several separate title lines with suffixes, they're mapped per-certificate the same way as names.
   - **Numbers that aren't suffixes**: a suffix is attached to the text (`India1`). PIN codes, years and numbers after a space, hyphen or dot (`Chennai - 600001`, `COVID-19`, `Industry 4.0`, `Phase 2`) stay part of the text.
   - The **Parsed Preview** table shows what will be printed on each certificate, with warnings above it. Check it before downloading.
   - Use the **"Load Example Data"** button to see a sample of the expected input format.

3. **Publication details** — volume, issue, year, month and the file number, supplied in one of two modes:

   - **With DOI** (default): paste the DOI and everything else is derived from it. The final numeric segment packs **volume (2 digits) + issue + file number (2 digits)**, and the segment before it is the year. The issue is whatever sits between volume and file number: one digit for issues 1–9, two digits for issues 10–12. So `10.17148/IJARCCE.2026.15817` → `{year}` `2026`, `{vol}` `15`, `{issue}` `8`, `{month}` `August`, file number `17`. A pasted `DOI:` prefix or `https://doi.org/` link is removed. A 3-digit file number or 1-digit volume is recognised when the usual layout can't be right, and a DOI with two possible readings shows a warning.
   - **Without DOI**: enter **volume**, **issue**, **file number** and **year** by hand. The DOI line is left off the certificate.

   In both modes the issue number *is* the month of publication, so `{month}` is always derived rather than typed — issue 7 is July, issue 9 is September. `getPublicationInfo()` in `app.js` resolves all of this for whichever mode is active, so the rest of the app never has to care which was used.

4. **Download** — one of three modes. All are disabled until a template is picked, data is parsed, and the publication details are complete and readable. Long text can be cut off by the fixed-size certificate text box, so a warning appears beyond the measured limits in `SAFE_TEXT_SIZES` (`app.js`):
   - **All in One**: every certificate combined into a single multi-page Word file (`buildCombinedDocxBlob` + `handleDownloadAllInOne`), named `"{file number} {author list}"` — e.g. `17 Author1, Author2.docx` (`buildCombinedFilename`).
   - **Multiple Downloads**: one separate file per recipient (`renderDocxBlobForRecord` + `handleDownloadMultiple`), triggered as sequential browser downloads staggered ~250ms apart. Each filename shares the same base as "All in One" but ends in that recipient's 1-based certificate index — e.g. `17 Author1, Author2 1.docx`, `17 Author1, Author2 2.docx` (`buildIndividualFilename`). Chrome may show a one-time "allow multiple downloads" prompt after the second file.
   - **Group Download**: every recipient combined onto a single certificate instead of one each (`buildGroupRecord` + `handleDownloadGroup`) — names are comma-joined, and Designation/Paper Title are de-duplicated (each distinct value appears once, in first-occurrence order) then comma-joined, since the same designation is often shared by several recipients. Same filename pattern as "All in One".

   The file number is omitted from any filename if it isn't available in the active mode (e.g. no DOI entered and no manual file number given).

**PDF export.** Every button in section 4 (All in One, Multiple Downloads, Group Download) downloads both a Word (.docx) file and a PDF under the same name -- there is no format switch. PDF works on the hosted site too, with nothing to install or configure. If the PDF cannot be made, the Word file is still downloaded and the log says why.

- **In the browser (default).** Word renders each certificate once into `templates/pdf/<KEY>.pdf` with the text box that holds the changing fields left empty. `pdf-export.js` (using pdf-lib) lays that box out itself -- word wrap, line spacing, colours, shadow, taken from `templates/pdf/<KEY>.json` -- draws the text at 300 dpi with the fonts installed on the visitor's computer, and stamps it onto the page. Checked against Word's own PDFs on all four templates: every line within 0.4 pt of Word, same line breaks. The text is an image (not selectable) and each certificate adds about 0.3 MB.
- **Fonts.** The templates use Microsoft Office fonts (Algerian, Calisto MT, Nirmala UI, Calibri, ...). Any computer without one prints that text in a different typeface; the page checks for this and shows a warning with the way out (download the .docx and use Word's File > Save As > PDF).
- **Exact, on a PC with Word.** Running the local server (`start.bat` or `node server.js`, Windows with Word installed) is preferred automatically: `POST /api/convert`, backed by `convert-docx-to-pdf.ps1`, has Word itself make the PDFs, so they match the Word file exactly. The server listens on `127.0.0.1` only (set `HOST` to change that) and serves only the site's own files.
- **When a template changes.** The background PDFs and layout files are generated, not hand-written. After editing anything in `templates/`, run `node tools/build-pdf-assets.js` on a Windows PC with Microsoft Word and commit the new `templates/pdf/` files.

## Dependencies

- [PizZip](https://github.com/open-xml-templating/pizzip) + [docxtemplater](https://docxtemplater.com/) (loaded from CDN) — filling placeholders inside the DOCX template
- [docx-preview](https://github.com/VolodymyrBaydalka/docxjs) + [JSZip](https://stuk.github.io/jszip/) (loaded from CDN) — the Live Preview panel
- [Font Awesome](https://fontawesome.com/) and [Google Fonts](https://fonts.google.com/) (loaded from CDN) — UI typography
- Node.js standard library only for `server.js` — no npm dependencies
- Microsoft Word (Windows, COM automation) for PDF export (optional)
