// Build-time tool for the in-browser PDF export. Run it on a Windows PC with Microsoft Word whenever
// a certificate template in templates/ changes:
//
//     node tools/build-pdf-assets.js            (all templates)
//     node tools/build-pdf-assets.js CCE ICE    (some)
//
// For every template it writes, into templates/pdf/:
//   <KEY>.pdf   the certificate as Word renders it, with the one text box that holds the changing
//               fields ({NAME}, {Designation}, ...) emptied -- vector, fonts embedded by Word
//   <KEY>.json  that text box's geometry and resolved formatting, so pdf-export.js can lay the
//               fields out itself and stamp them onto the page
//
// Needs: Windows, Microsoft Word, Node. No npm packages.

const fs = require('fs');
const os = require('os');
const path = require('path');
const { execFileSync } = require('child_process');

const ROOT = path.resolve(__dirname, '..');
const OUT_DIR = path.join(ROOT, 'templates', 'pdf');
const TAR = path.join(process.env.SystemRoot || 'C:\\Windows', 'System32', 'tar.exe'); // bsdtar: reads and writes zip
const KEYS = ['CCE', 'ICE', 'SET', 'RJR'];
const NAME_FIELD = /\{NAM[Ee]\}/;

function ps(script, args) {
    return execFileSync('powershell.exe', ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', script, ...args],
        { encoding: 'utf8', maxBuffer: 64 * 1024 * 1024 });
}

// Every <w:txbxContent> span in document.xml, as [start, end) offsets.
function textBoxSpans(xml) {
    const spans = [];
    let from = 0;
    for (;;) {
        const s = xml.indexOf('<w:txbxContent>', from);
        if (s < 0) break;
        const e = xml.indexOf('</w:txbxContent>', s) + '</w:txbxContent>'.length;
        spans.push([s, e]);
        from = e;
    }
    return spans;
}

// Empty the text box(es) holding the NAME field -- both the DrawingML copy and the VML fallback.
function blankDynamicBox(xml) {
    const spans = textBoxSpans(xml).filter(([s, e]) => NAME_FIELD.test(xml.slice(s, e)));
    if (spans.length === 0) throw new Error('No text box with a NAME field found.');
    let out = xml;
    for (const [s, e] of spans.reverse()) {
        out = out.slice(0, s) + '<w:txbxContent><w:p/></w:txbxContent>' + out.slice(e);
    }
    return out;
}

// Word reports theme colours as unresolved negative numbers; for those, take the colour written
// in the XML run that carries the same text.
function xmlRunColors(xml) {
    const [s, e] = textBoxSpans(xml).find(([a, b]) => NAME_FIELD.test(xml.slice(a, b)));
    const box = xml.slice(s, e);
    const runs = [];
    for (const r of box.matchAll(/<w:r[ >][\s\S]*?<\/w:r>/g)) {
        const text = [...r[0].matchAll(/<w:t[^>]*>([^<]*)</g)].map((m) => m[1]).join('');
        const m = r[0].match(/<w:color w:val="([0-9A-Fa-f]{6})"/);
        runs.push({ text, color: m ? '#' + m[1].toLowerCase() : '#000000' });
    }
    return runs;
}

// A text box with no fill and an outer shadow casts that shadow on its text. (With a fill, the shadow
// belongs to the box -- already part of the background PDF -- and the text has none.)
function textShadow(xml) {
    const i = xml.search(NAME_FIELD);
    const s = xml.lastIndexOf('<wps:spPr', i);
    const e = xml.indexOf('</wps:spPr>', s);
    const sp = xml.slice(s, e);
    const fill = sp.replace(/<a:effectLst>[\s\S]*?<\/a:effectLst>/, '').replace(/<a:ln>[\s\S]*?<\/a:ln>/, '');
    if (!/<a:noFill\/>/.test(fill) || /<a:(solidFill|gradFill|blipFill|pattFill)/.test(fill)) return null;
    const m = sp.match(/<a:outerShdw([^>]*)>[\s\S]*?<a:alpha val="(\d+)"/);
    if (!m) return null;
    const attr = (name, dflt) => { const a = m[1].match(new RegExp(name + '="(-?\\d+)"')); return a ? Number(a[1]) : dflt; };
    const dist = attr('dist', 0) / 12700;
    const dir = (attr('dir', 0) / 60000) * Math.PI / 180; // clockwise from east
    return { dx: +(dist * Math.cos(dir)).toFixed(3), dy: +(dist * Math.sin(dir)).toFixed(3), blur: attr('blurRad', 0) / 12700, alpha: Number(m[2]) / 100000 };
}

// ---- Font line metrics -------------------------------------------------------------------------
// Word lays lines out with each font's Windows metrics (OS/2 usWinAscent / usWinDescent). Browsers
// report other values for some fonts (Algerian, notably), which shifts every line below it, so the
// exact ratios are read from the installed font files here and stored in the spec.

let installedFonts = null;
function fontFile(family, bold, italic) {
    if (!installedFonts) installedFonts = JSON.parse(ps(path.join(__dirname, 'list-fonts.ps1'), []));
    const style = (bold ? ' Bold' : '') + (italic ? ' Italic' : '');
    const names = [family + style, family + (bold ? ' Bold' : ''), family];
    for (const name of names) {
        for (const suffix of [' (TrueType)', ' (OpenType)']) {
            if (installedFonts[name + suffix]) return installedFonts[name + suffix];
        }
    }
    // Collections list several faces in one entry, e.g. "Cambria & Cambria Math (TrueType)".
    const hit = Object.keys(installedFonts).find((k) => k.split(' & ').some((part) => part.replace(/ \((TrueType|OpenType)\)$/, '') === family + style));
    return hit ? installedFonts[hit] : null;
}

function readSfntMetrics(file) {
    const buf = fs.readFileSync(file);
    let base = 0;
    if (buf.toString('latin1', 0, 4) === 'ttcf') base = buf.readUInt32BE(12); // first face of a collection
    const numTables = buf.readUInt16BE(base + 4);
    const tables = {};
    for (let i = 0; i < numTables; i++) {
        const o = base + 12 + i * 16;
        tables[buf.toString('latin1', o, o + 4)] = buf.readUInt32BE(o + 8);
    }
    if (!tables.head || !tables['OS/2'] || !tables.hhea) return null;
    const upm = buf.readUInt16BE(tables.head + 18);
    const winAscent = buf.readUInt16BE(tables['OS/2'] + 74);
    const winDescent = buf.readUInt16BE(tables['OS/2'] + 76);
    const hheaAscent = buf.readInt16BE(tables.hhea + 4);
    const hheaDescent = buf.readInt16BE(tables.hhea + 6);
    const hheaGap = buf.readInt16BE(tables.hhea + 8);
    // The Windows (GDI) rule Word follows: external leading is whatever line gap the Windows
    // ascent/descent do not already cover, and it is added above the ascent.
    const leading = Math.max(0, hheaGap - (winAscent + winDescent - (hheaAscent - hheaDescent)));
    return { ascent: (winAscent + leading) / upm, descent: winDescent / upm };
}

function fontMetricsFor(rawParagraphs) {
    const out = {};
    for (const p of rawParagraphs) {
        for (const r of p.runs) {
            const family = r.font.replace(/\s+(Bold|Italic|Regular)$/i, '');
            const bold = r.bold, italic = r.italic;
            const key = family + '|' + (bold ? 1 : 0) + '|' + (italic ? 1 : 0);
            if (key in out) continue;
            const file = fontFile(family, bold, italic);
            const m = file && fs.existsSync(file) ? readSfntMetrics(file) : null;
            if (!m) console.warn('  (no font metrics for ' + key + ' -- the page falls back to the browser values)');
            out[key] = m ? { ascent: +m.ascent.toFixed(5), descent: +m.descent.toFixed(5) } : null;
        }
    }
    return out;
}

function normalize(raw, xmlRuns, shadow) {
    const toHex = (bgr) => '#' + [bgr & 255, (bgr >> 8) & 255, (bgr >> 16) & 255].map((v) => v.toString(16).padStart(2, '0')).join('');
    const pageW = raw.page.width, pageH = raw.page.height;
    const b = raw.box;
    // Word's relative frames: horizontal 0 margin / 1 page / 2 column; vertical 0 margin / 1 page / 2 paragraph.
    const x = (b.relH === 1 ? 0 : raw.page.marginLeft) + b.left;
    const y = b.relV === 1 ? b.top : b.relV === 0 ? raw.page.marginTop + b.top : b.anchorParaTopOnPage + b.top;
    const align = ['left', 'center', 'right', 'justify'];
    const paragraphs = raw.paragraphs.map((p) => {
        let line;
        if (p.lineRule === 0) line = { rule: 'multiple', value: 1 };
        else if (p.lineRule === 1) line = { rule: 'multiple', value: 1.5 };
        else if (p.lineRule === 2) line = { rule: 'multiple', value: 2 };
        else if (p.lineRule === 5) line = { rule: 'multiple', value: p.lineSpacing / 12 };
        else if (p.lineRule === 3) line = { rule: 'atLeast', value: p.lineSpacing };
        else line = { rule: 'exact', value: p.lineSpacing };
        return {
            align: align[p.align] || 'left', spaceBefore: p.spaceBefore, spaceAfter: p.spaceAfter, line,
            leftIndent: p.leftIndent, rightIndent: p.rightIndent, firstLineIndent: p.firstLineIndent,
            runs: p.runs.map((r) => {
                let color;
                if (r.color >= 0) color = toHex(r.color);
                else {
                    // The XML run that ends the same way this run does ("includes" would also match the
                    // earlier "{#hasDoi}DOI" run, since "{/hasDoi}" and "{#hasDoi}" share their last characters).
                    const tail = r.text.slice(-3);
                    const hit = [...xmlRuns].reverse().find((xr) => xr.text.endsWith(tail)) || xmlRuns.find((xr) => xr.text.startsWith(r.text.slice(0, 3)));
                    color = hit ? hit.color : '#000000';
                }
                return {
                    text: r.text, font: r.font, size: r.size, bold: r.bold, italic: r.italic, color,
                    spacing: r.spacing, caps: r.allCaps, smallCaps: r.smallCaps, underline: r.underline !== 0,
                    superscript: r.superscript, subscript: r.subscript,
                };
            }),
        };
    });
    return {
        page: { width: pageW, height: pageH },
        box: { x, y, width: b.width, height: b.height, insetLeft: b.insetLeft, insetTop: b.insetTop, insetRight: b.insetRight, insetBottom: b.insetBottom },
        textShadow: shadow,
        fontMetrics: fontMetricsFor(raw.paragraphs),
        paragraphs,
    };
}

function buildOne(key, work) {
    const template = path.join(ROOT, 'templates', `${key}.docx`);
    const dir = path.join(work, key);
    fs.mkdirSync(path.join(dir, 'x'), { recursive: true });
    fs.mkdirSync(path.join(dir, 'in'), { recursive: true });
    fs.mkdirSync(path.join(dir, 'out'), { recursive: true });

    ps(path.join(__dirname, 'extract-spec.ps1'), ['-Docx', template, '-Out', path.join(dir, 'raw.json')]);
    const raw = JSON.parse(fs.readFileSync(path.join(dir, 'raw.json'), 'utf8').replace(/^\uFEFF/, ''));

    execFileSync(TAR, ['-xf', template, '-C', path.join(dir, 'x')]);
    const docPath = path.join(dir, 'x', 'word', 'document.xml');
    const xml = fs.readFileSync(docPath, 'utf8');
    const spec = normalize(raw, xmlRunColors(xml), textShadow(xml));
    fs.writeFileSync(docPath, blankDynamicBox(xml));

    const blank = path.join(dir, 'in', `${key}.docx`);
    ps(path.join(__dirname, 'zip-dir.ps1'), ['-Dir', path.join(dir, 'x'), '-Out', blank]);
    ps(path.join(ROOT, 'convert-docx-to-pdf.ps1'), ['-InputDir', path.join(dir, 'in'), '-OutputDir', path.join(dir, 'out')]);

    fs.mkdirSync(OUT_DIR, { recursive: true });
    fs.copyFileSync(path.join(dir, 'out', `${key}.pdf`), path.join(OUT_DIR, `${key}.pdf`));
    fs.writeFileSync(path.join(OUT_DIR, `${key}.json`), JSON.stringify(spec, null, 1) + '\n');
    console.log(`${key}: ${(fs.statSync(path.join(OUT_DIR, `${key}.pdf`)).size / 1024).toFixed(0)} KB background, ${spec.paragraphs.length} paragraphs`);
}

const wanted = process.argv.length > 2 ? process.argv.slice(2) : KEYS;
const work = fs.mkdtempSync(path.join(os.tmpdir(), 'certgen-pdf-assets-'));
try {
    for (const key of wanted) {
        if (!KEYS.includes(key)) throw new Error(`Unknown template "${key}".`);
        buildOne(key, work);
    }
} finally {
    fs.rmSync(work, { recursive: true, force: true });
}
