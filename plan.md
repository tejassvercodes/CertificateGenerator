# Full audit & repair of the Certificate Generator

## Context

The app generates journal certificates by filling four pre-loaded DOCX templates
(`templates/CCE|SET|ICE|RJR.docx`) with recipient data pasted as free text, then downloading
them as one combined file, one file per recipient, or a single group certificate.

A full read of `app.js`, `index.html`, `styles.css`, `server.js`, `convert-docx-to-pdf.ps1` and
the unpacked template XML — plus running `parseDataInput` against 15 realistic inputs — found
that the app does not fail loudly. It fails **silently**: a paper title lands in the wrong
field, a certificate comes out blank, volume/issue/month render empty, and the user only
discovers it after opening the `.docx`. Several of these reproduce on ordinary, well-formed
input.

Goal: fix every defect below, make failure modes visible instead of silent, and document the
whole audit in a new `Fix.md` at the repo root.

Decisions already made:
- The unreachable PDF backend (`server.js` `/api/convert` + `convert-docx-to-pdf.ps1`) **stays
  as-is**; only the docs that over-claim it are corrected.
- A **parsed-data preview table** is added so classification is verifiable before download.
- Blank certificates are **skipped with a warning**, and a cap rejects runaway suffix ranges.

---

## Confirmed defects

Severity: **P0** = wrong output on ordinary input · **P1** = wrong output on plausible input ·
**P2** = robustness/correctness hardening · **P3** = docs/dead code.

### Parser — `parseDataInput` / `parseMappingLine` / `normalizeSuperscripts` (app.js:223–436)

| # | Sev | Defect | Reproduction (verified) |
|---|-----|--------|-------------------------|
| 1 | **P0** | A paper title containing an affiliation keyword is classified as a Designation, so `{Paper Title}` renders **empty** on every certificate. Rule 2 (`hasKeyword`, app.js:322) runs before the title rules and matches on substrings — `student`, `india`, `university`, `studies`, `college` all appear in normal titles. | `A Study on Student Performance in Indian Universities` / `Ravi Kumar1, Anita Rao2` / `Professor, Dept of CS, ABC College, Chennai, India1-2` → both records have **no `PaperTitle` key at all**. |
| 2 | **P0** | Superscript **ranges and separators** are not normalized. `SUPERSCRIPT_MAP` (app.js:256) covers only `⁰–⁹`, not superscript minus `⁻` (U+207B), plus `⁺` (U+207A), or the superscript comma. | `Professor, ABC College, India¹⁻²` → yields the literal value `"Professor, ABC College, India1⁻"` mapped to record 2 only; record 1 gets **no Designation**. |
| 3 | **P1** | A gap in the index sequence produces a **completely blank certificate page/file**. | `Alice1, Bob3` → 3 records; record 2 is `{}` and renders an empty certificate. |
| 4 | **P1** | No upper bound on `maxIdx`. A typo'd or pasted range builds thousands of records and freezes the tab. | `Alice1-500` → 500 records generated, no warning. |
| 5 | **P1** | Text after the last suffixed segment on a line is **silently dropped** — `parseMappingLine` discards a non-empty `bufferVal` when the loop ends (app.js:250). | `Professor, Anna University1, Chennai` → `", Chennai"` vanishes. |
| 6 | **P1** | `applyMappedLine`'s fallback (app.js:419) stamps the **raw line including its suffix digits** onto every certificate when no suffix is in range. | `…India3` with `maxIdx = 2` prints `"…India3"` on the certificate. |
| 7 | **P2** | An inverted range (`5-2`) maps nothing, silently. `parseSuffix` (app.js:444) only loops upward. | — |
| 8 | **P2** | `parseDataInput` writes `state.authorListLine` as a side effect mid-parse (app.js:375); if it later throws, a stale author list survives into the next filename. | — |

### DOI / publication details — `parseDoiParts` / `getPublicationInfo` (app.js:465–520)

| # | Sev | Defect | Reproduction (verified) |
|---|-----|--------|-------------------------|
| 9 | **P0** | An unparseable or empty DOI **silently yields blank** vol/issue/month with downloads still enabled — the certificate prints empty fields. Nothing warns. | `10.17148/IMRJR.2026.0207` (final run < 5 digits, app.js:489) → `{vol:"", issue:"", month:""}`, buttons stay green. Same for an empty DOI box. |
| 10 | **P1** | The year is assumed to be the second-to-last numeric run, with no validation. Any trailing digit breaks it. | `10.17148/IJARCCE.2026.15817v2` → `year: "15817"`, `vol/issue` garbage. |
| 11 | **P1** | Volume keeps its leading zero (`last.slice(0,2)`) while issue is int-parsed — volume 5 prints as `05`. | `…2026.05817` → `vol: "05"`. |
| 12 | **P1** | Manual mode has no validation: a non-numeric volume, an issue outside 1–12, or an empty year all pass through and render blank/garbage. | Issue `13` → `month: ""` silently. |

### DOCX rendering & combining (app.js:612–722)

| # | Sev | Defect | Notes |
|---|-----|--------|-------|
| 13 | **P1** | `buildCombinedDocxBlob` concatenates each record's body verbatim, duplicating every `wp:docPr/@id`, `w14:paraId`, `w14:textId` and `w:bookmarkStart/@w:id`. OOXML requires these to be unique; Word can respond with an "unreadable content / repair" prompt. Verified counts per template: CCE 10 `docPr` + 43 `paraId`, SET 10 + 39 + 2 bookmarks, ICE 9 + 52, RJR 18 + 33. |
| 14 | **P1** | The section-break separator `<w:p><w:pPr>{sectPr}</w:pPr></w:p>` (app.js:714) carries **default paragraph formatting**. With `w:top`/`w:bottom` page margins of `0` in every template, that extra line can overflow into a spurious blank page between certificates. Must be given zero spacing and a minimal font size. |
| 15 | **P2** | `neutralizeBraces` is applied to NAME/Designation/PaperTitle only — `DOI`, `vol`, `issue`, `year`, `month` bypass it (app.js:618, 632–634, 658–665). A brace in any of those breaks the render. |
| 16 | **P2** | Template-switch **race**: `selectTemplate` awaits `fetch` without a request token. Click CCE then SET; if CCE resolves second it overwrites `state.docxBuffer` while SET's card shows active — exports use the wrong journal. The `catch` branch's `markActiveCard('')` (app.js:180) has the same flaw. |
| 17 | **P2** | If the PizZip/docxtemplater CDN is blocked, nothing is detected until the first download click, which then fails with `PizZip is not a constructor`. |
| 18 | **P2** | docxtemplater render errors surface as the bare message `"Multi error"`; its `properties.errors[]` (tag name, offset) is discarded (app.js:744). |
| 19 | **P2** | Filenames embed the entire author list untruncated and are never length-capped; `sanitizeFilename` (app.js:523) also ignores control characters. Four authors easily exceed the ~255-char filename limit and the download fails or is mangled. |

### UI / docs / dead code

| # | Sev | Defect |
|---|-----|--------|
| 20 | **P3** | `index.html:7` meta description claims output "as a pixel-accurate PDF" — the UI has no PDF path. |
| 21 | **P3** | `README.md` presents `/api/convert` as a working step; it is never called by the front end. |
| 22 | **P3** | `styles.css:439–498` is dead: `.preview-card`, `.preview-pager`, `.btn-pager`, `#preview-index`, `.preview-viewport`, `.docx-render-preview` belong to a removed feature. `.btn-primary` is also unused. |
| 23 | **P3** | `server.js` `serveStatic` happily serves `server.js`, `.claude/settings.local.json` and anything else under the project root. |
| 24 | **P3** | `convert-docx-to-pdf.ps1:30` — `$files.Count` on a `$null` result (empty input dir) never equals `0`, so `NO_FILES` is unreachable and Word is started for nothing. Needs `@($files).Count`. |

---

## Implementation

### 1. `app.js` — parser rewrite of the classification step

Replace the rule ladder in `parseDataInput` (app.js:297–354). New order, evaluated per line:

1. **Honorific name** — unchanged (app.js:315), it is the strongest signal.
2. **Prose title** — `commaCount <= 1 && words >= 4` wins over the keyword rule, *unless* the
   line starts with an affiliation keyword or contains 2+ of them. This is the fix for #1:
   `A Study on Student Performance in Indian Universities` (0 commas, 7 words, 1 keyword, does
   not start with one) → title; `Department of CS Anna University Chennai India1` (starts with
   `department`, 3 keywords) → designation.
3. **Keyword** → Designation (unchanged).
4. **Suffixed + comma-free + short** → NAME (unchanged, app.js:332).
5. **Any other suffixed line** → Designation.
6. **No suffix** → PaperTitle.

Compute `commaCount` from the trimmed line and keep the existing `parsed`/`hasSuffix`/
`hadSuperscript`/`words` locals — no new parsing pass.

Other parser changes:
- Extend `SUPERSCRIPT_MAP` / the regex in `normalizeSuperscripts` (app.js:256–259) with
  `⁻`→`-`, `⁺`→`+`, `˒`/`﹐`→`,` (fixes #2).
- `parseMappingLine`: flush a leftover `bufferVal` after the loop onto the last item's value
  (or as a suffix-less item) instead of dropping it (#5).
- `parseSuffix`: normalize inverted ranges by swapping `start`/`end` (#7).
- Add `const MAX_CERTIFICATES = 200;` — throw a clear parse error when `maxIdx` exceeds it (#4).
- After mapping, drop records with no `NAME`, collect their original indexes, and expose them
  as a warning (#3). Records are renumbered 1..N for output; the preview shows both.
- `applyMappedLine` fallback: use `items.map(i => i.val).join(', ')` when `items` is non-empty,
  falling back to the raw line only when there were no suffixed items at all (#6).
- Make `parseDataInput` pure: return `{ records, authorListLine, warnings }` and let
  `handleDataInput` commit it to `state` (#8).

### 2. `app.js` — publication details

- `parseDoiParts`: require the year run to be exactly 4 digits and the final run to be ≥5
  digits; strip the volume's leading zero; return an explicit `error` string (e.g.
  `"DOI does not encode a volume/issue"`) instead of silently blank fields (#9, #10, #11).
- `getPublicationInfo`: return the same `{ …, error }` shape for manual mode — non-numeric
  volume, issue outside 1–12, or a non-4-digit year each produce a message (#12).
- `handleDataInput` surfaces that `error` in `parser-status`, in the log, and in the preview —
  as a **warning, not a block** (a certificate with no DOI is legitimate).

### 3. `app.js` — DOCX generation

- **`renderDocxZipForRecord`**: run `neutralizeBraces` over every value, including the DOI and
  the numeric fields (#15).
- **`buildCombinedDocxBlob`**: add a `remapDocxIds(bodyXml, offset)` helper applied to every
  body after the first — regex-rewrite `wp:docPr id="N"`, `w14:paraId="HEX"`,
  `w14:textId="HEX"`, `<w:bookmarkStart w:id="N"` / `<w:bookmarkEnd w:id="N"` and `<v:shape
  id="…">` with a per-copy offset so ids stay unique (#13).
- Give the separator paragraph explicit zero geometry (#14):
  `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="1" w:lineRule="exact"/><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr>{sectPr}</w:pPr></w:p>`
- **`selectTemplate`**: bump a `state.loadToken` on entry and ignore a resolved fetch (both
  success and error paths) whose token is stale (#16).
- **`init`**: check `window.PizZip` / `window.docxtemplater` up front; if missing, log an error
  and keep all download buttons disabled (#17).
- **`runDownload`**: when `err.properties?.errors` exists, log each entry's `id`, `explanation`
  and offending tag rather than `"Multi error"` (#18).
- **`sanitizeFilename`**: strip control characters, collapse whitespace, trim trailing dots and
  spaces, and truncate the base to ~150 chars so `base + " N.docx"` stays well under the limit
  (#19). `buildCombinedFilename` / `buildIndividualFilename` keep their current shape.

### 4. `index.html` + `styles.css` + `app.js` — parsed preview table

Add a card between step 2 and step 3 containing `<div id="parse-warnings">` and a
`<table id="preview-table">`, rebuilt by a new `renderPreview()` called from `handleDataInput`:

- One row per output certificate: `#`, NAME, Designation, Paper Title.
- A footer line with the resolved DOI / vol / issue / month / year and the file number, taken
  from the existing `getPublicationInfo()`.
- Warnings rendered above it in the existing error colour: skipped blank certificates, missing
  paper title, unparseable DOI/manual fields, record count near the cap.
- All cell content set with `textContent` (matching the XSS-safe pattern already used in
  `log()` at app.js:58) — never `innerHTML`.
- Hidden entirely when there are no records, so the empty state is unchanged.
- Styles reuse the existing `--bg-input` / `--border` / `--text-muted` tokens; delete the dead
  `.preview-*` / `.docx-render-preview` / `.btn-primary` blocks at `styles.css:376–498` (#22)
  and add the table rules in their place.

### 5. Docs & server hygiene

- `index.html:7` — meta description describes DOCX output only (#20).
- `README.md` — state plainly that `/api/convert` is a server-side utility not wired to any UI
  button; document the new preview table, the certificate cap, and blank-record skipping (#21).
- `server.js` `serveStatic` — deny dotfiles and a small denylist (`server.js`,
  `convert-docx-to-pdf.ps1`, `*.md`, `*.ps1`) before reading (#23).
- `convert-docx-to-pdf.ps1:30` — `@($files).Count` (#24).

### 6. `Fix.md` (new, repo root)

Written last, from the finished work: every defect above with its severity, the reproduction
that demonstrated it, the fix applied, and the file/function touched — plus a "Verified test
cases" section listing the inputs below and their expected output.

---

## Verification

**Parser (no browser needed).** The existing `app.js` can be evaluated in Node with a stubbed
`document`, which is how every P0/P1 parser defect above was confirmed. Re-run the same harness
after the change and assert:

| Input | Expected |
|-------|----------|
| Demo data (`loadDemoData`, app.js:126) | 2 records, correct NAME/Designation/PaperTitle — unchanged from today |
| `A Study on Student Performance in Indian Universities` + 2 names + 1 designation | `PaperTitle` present on **both** records (#1) |
| `Professor, ABC College, India¹⁻²` | Designation `"Professor, ABC College, India"` on **both** records (#2) |
| `Alice1, Bob3` | 2 records (Alice, Bob), warning naming the skipped index (#3) |
| `Alice1-500` | Parse error citing the 200 cap (#4) |
| `Professor, Anna University1, Chennai` | `", Chennai"` retained (#5) |
| `Ravi Kumar1` (single author) | 1 record — must not regress |
| `Impact of AI` + a designation wrapped across two lines ending in `,` | Stitched into one designation — must not regress |
| `10.17148/IJARCCE.2026.15817` | `vol 15 / issue 8 / August / 2026`, file `17` — must not regress |
| `10.17148/IMRJR.2026.0207`, `…15817v2`, empty DOI | Explicit `error` message, no silent blanks (#9, #10) |
| `…2026.05817` | `vol: "5"` (#11) |
| Manual mode, issue `13` | Explicit validation message (#12) |

**End to end.** `node server.js`, open <http://localhost:8080>, then for **each of the four
templates**: load the demo data, confirm the preview table matches, and exercise all three
download modes. Repeat with a 3-author input in With-DOI and Without-DOI modes.

**Combined-document integrity (the #13/#14 fix).** Generate a 3-record "All in One" DOCX, then
convert it with the existing script:

```
powershell -File convert-docx-to-pdf.ps1 -InputDir <dir> -OutputDir <dir>
```

Word must open it **without a repair prompt**, and the PDF must have exactly **3 pages** — no
blank page between certificates, every page laid out like page 1. Verify for all four
templates; RJR is the important one for layout since it is the only template without
`<w:titlePg/>`.

**Regression on switching.** Click two templates in quick succession and confirm the exported
file matches the card that is highlighted (#16).


---

# PDF export (Word-only)

- Section 4 gets a **Word (.docx) / PDF** switch; the three buttons (All in One, Multiple Downloads,
  Group Download) work for either format. DOCX blobs are built exactly as before, then for PDF are
  sent to `POST /api/convert` (Microsoft Word via `convert-docx-to-pdf.ps1`) and the returned PDFs
  are downloaded.
- `GET /api/convert/status` (server.js) reports whether Word conversion is possible; the page probes
  it on load. PDF buttons stay disabled unless it says yes (so on Vercel PDF is off, Word stays on).
- **Rejected: in-browser PDF** (docx-preview + html2canvas + jsPDF). Tested on all four templates:
  valid PDFs, but names/titles missing and backgrounds misplaced, because docx-preview mis-lays these
  templates. Printing wrong certificates is worse than no PDF, so it was removed.
- Also: Live Preview placeholder/note text darkened (#8a8a7a -> #5f5f52) for contrast.
- Open option for PDF on a hosted site: run the converter on a Windows host with Word, or a
  LibreOffice-based service, and point the page at it.

# In-browser PDF export (hosted site)

- Replaces the rejected docx-preview route. Each template is split into (a) a Word-rendered background PDF with the dynamic text box emptied and (b) a JSON spec of that box (position, insets, resolved run formatting, text shadow, font line metrics). `pdf-export.js` lays the box out and stamps it as a 300 dpi PNG with pdf-lib; `tools/build-pdf-assets.js` (+ `extract-spec.ps1`, `zip-dir.ps1`, `list-fonts.ps1`) regenerate the assets with Word.
- Layout rules matched to Word: 1.5x line spacing is added *below* a line, the first baseline sits one font ascent under the line top; font ascent includes GDI external leading (win ascent/descent + hhea line gap rule), read from the installed font files; no kerning/ligatures; breaks allowed after hyphens; a line is dropped if it does not fit inside the box.
- Verified: 12 sample certificates (4 templates x short/long/no-DOI) compared against Word PDFs -- baselines within 0.4 pt, line starts within 0.3 pt, same wraps; visual side-by-side on all templates. The full UI path (template, paste data, PDF, All in One / Multiple / Group) run in headless Chrome on a static server and on the local Word server.
- Known limits: text is an image; ~0.3 MB per certificate; visitors lacking the Office fonts see a warning and a different typeface; the font metrics come from the PC that ran the build tool.

# Word + PDF downloaded together

- The Word/PDF switch is gone. `deliverFiles` always downloads the .docx and, right after it, the .pdf with the same stem (interleaved per certificate for Multiple Downloads). If the PDF step fails (e.g. pdf-lib did not load), the Word files are still downloaded and an error line is logged.
- Verified in headless Chrome on a static server (browser route) and on the local server (Word route): All in One = 1 docx + 1 pdf, Group = 1 + 1, Multiple (2 records) = 2 docx + 2 pdf; with the PDF library removed, the Word file still downloads.
