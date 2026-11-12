// In-browser PDF export.
//
// A certificate is a Word page made of floating shapes and Microsoft Office fonts, which no browser
// renderer reproduces. So the PDF is assembled from two parts instead:
//
//   1. templates/pdf/<KEY>.pdf -- the certificate exactly as Word renders it, with the one text box that
//      holds the changing fields (name, designation, paper title, volume line, DOI) left empty.
//   2. That text box, laid out here: templates/pdf/<KEY>.json holds its position and Word's resolved
//      formatting; the fields are word-wrapped and drawn to a transparent 300 dpi image with the
//      fonts installed on this computer (the same ones Word uses), then stamped onto the page.
//
// Both files are produced by tools/build-pdf-assets.js. Needs pdf-lib (index.html loads it).

(function () {
    'use strict';

    const PX_PER_PT = 300 / 72; // the text image is drawn at 300 dpi
    const assetCache = {};

    function loadAssets(key) {
        if (!assetCache[key]) {
            assetCache[key] = (async () => {
                const [specRes, pdfRes] = await Promise.all([fetch(`templates/pdf/${key}.json`), fetch(`templates/pdf/${key}.pdf`)]);
                if (!specRes.ok || !pdfRes.ok) throw new Error(`PDF assets for ${key} are missing.`);
                return { spec: await specRes.json(), pdf: await pdfRes.arrayBuffer() };
            })();
            assetCache[key].catch(() => { delete assetCache[key]; });
        }
        return assetCache[key];
    }

    // Word names some faces oddly ("Arial Bold" is Arial with the bold flag set).
    function familyOf(name) {
        return name.replace(/\s+(Bold|Italic|Regular)$/i, '');
    }

    function cssFont(run, px) {
        const family = familyOf(run.font);
        const generic = /times|serif|georgia|garamond|cambria|calisto|constantia|book/i.test(family) && !/sans/i.test(family) ? 'serif' : 'sans-serif';
        return `${run.italic ? 'italic ' : ''}${run.bold ? 'bold ' : ''}${px}px "${family}", ${generic}`;
    }

    // Which of the template's fonts this computer does not have. A missing font falls back to a
    // different typeface, so the certificate would not look like the Word file.
    function missingFonts(spec) {
        const ctx = document.createElement('canvas').getContext('2d');
        const probe = 'mmmmmmmmmmlliWWWW 0123456789';
        const widthOf = (font) => { ctx.font = font; return ctx.measureText(probe).width; };
        const families = new Set();
        spec.paragraphs.forEach((p) => p.runs.forEach((r) => families.add(familyOf(r.font))));
        return [...families].filter((family) => {
            const mono = widthOf('72px monospace');
            const sans = widthOf('72px sans-serif');
            const serif = widthOf('72px serif');
            const mine = (fallback) => widthOf(`72px "${family}", ${fallback}`);
            return mine('monospace') === mono && mine('sans-serif') === sans && mine('serif') === serif;
        });
    }

    // ---- Filling the template's fields ---------------------------------------------------------

    function normTag(tag) {
        return String(tag).toLowerCase().replace(/[\s_]+/g, '');
    }

    function fieldValues(record) {
        return {
            name: (record.NAME || '').toUpperCase(), // the Word export upper-cases the name too
            designation: record.Designation || '',
            papertitle: record.PaperTitle || '',
            doi: record.DOI || '',
            vol: record.Volume || '',
            volume: record.Volume || '',
            issue: record.Issue || '',
            year: record.Year || '',
            month: record.Month || '',
        };
    }

    // A template paragraph as characters that each remember their run's formatting, with the
    // {tags} replaced by the record's values. Returns null for a paragraph that is dropped (the
    // DOI line when there is no DOI, as in the Word export).
    function fillParagraph(paragraph, values) {
        let chars = [];
        paragraph.runs.forEach((run) => { for (const ch of run.text) chars.push({ ch, run }); });

        const text = () => chars.map((c) => c.ch).join('');
        const cond = /\{#hasDoi\}([\s\S]*?)\{\/hasDoi\}/.exec(text());
        if (cond) {
            if (!values.doi) return null;
            const start = cond.index;
            const innerStart = start + '{#hasDoi}'.length;
            const innerEnd = start + cond[0].length - '{/hasDoi}'.length;
            chars = chars.slice(0, start).concat(chars.slice(innerStart, innerEnd), chars.slice(start + cond[0].length));
        }

        const out = [];
        const re = /\{([^{}]+)\}/g;
        const source = text();
        let last = 0;
        let m;
        while ((m = re.exec(source))) {
            for (let i = last; i < m.index; i++) out.push(chars[i]);
            const run = chars[m.index].run;
            const value = values[normTag(m[1])];
            for (const ch of String(value === undefined ? '' : value)) out.push({ ch, run });
            last = m.index + m[0].length;
        }
        for (let i = last; i < chars.length; i++) out.push(chars[i]);
        return out;
    }

    // ---- Laying the box out ---------------------------------------------------------------------

    function applyRun(ctx, run) {
        ctx.font = cssFont(run, run.size * PX_PER_PT);
        ctx.letterSpacing = `${(run.spacing || 0) * PX_PER_PT}px`;
        // Word does not kern or ligate by default; matching that keeps line widths (and so line breaks) the same.
        ctx.fontKerning = 'none';
        ctx.textRendering = 'optimizeSpeed';
    }

    function displayChar(ch, run) {
        return run.caps ? ch.toUpperCase() : ch;
    }

    // Width of a character sequence, drawn in the runs they carry.
    function measure(ctx, chars) {
        let width = 0;
        let i = 0;
        while (i < chars.length) {
            let j = i;
            while (j < chars.length && chars[j].run === chars[i].run) j++;
            applyRun(ctx, chars[i].run);
            width += ctx.measureText(chars.slice(i, j).map((c) => displayChar(c.ch, c.run)).join('')).width;
            i = j;
        }
        return width;
    }

    function fontMetrics(ctx, run, table) {
        const exact = table && table[`${familyOf(run.font)}|${run.bold ? 1 : 0}|${run.italic ? 1 : 0}`];
        if (exact) return { ascent: exact.ascent * run.size * PX_PER_PT, descent: exact.descent * run.size * PX_PER_PT };
        applyRun(ctx, run);
        const m = ctx.measureText('Hg');
        const size = run.size * PX_PER_PT;
        const ascent = m.fontBoundingBoxAscent != null ? m.fontBoundingBoxAscent : size * 0.9;
        const descent = m.fontBoundingBoxDescent != null ? m.fontBoundingBoxDescent : size * 0.25;
        return { ascent, descent };
    }

    // Greedy word wrap: break at spaces (and inside a word that is wider than the line).
    function wrap(ctx, chars, available) {
        const lines = [];
        let line = [];
        let word = [];
        const lineWidth = (arr) => measure(ctx, arr);
        const pushLine = () => { lines.push(line); line = []; };

        const flushWord = () => {
            if (!word.length) return;
            const candidate = line.concat(word);
            if (line.length && lineWidth(candidate) > available) {
                while (line.length && line[line.length - 1].ch === ' ') line.pop();
                pushLine();
            }
            if (!line.length && lineWidth(word) > available) {
                // A single word wider than the line: break it by characters.
                let piece = [];
                for (const c of word) {
                    if (piece.length && lineWidth(piece.concat(c)) > available) { line = piece; pushLine(); piece = []; }
                    piece.push(c);
                }
                line = piece;
            } else {
                line = line.concat(word);
            }
            word = [];
        };

        for (let i = 0; i < chars.length; i++) {
            const c = chars[i];
            if (c.ch === '\n') { flushWord(); while (line.length && line[line.length - 1].ch === ' ') line.pop(); pushLine(); continue; }
            if (c.ch === ' ') { flushWord(); line.push(c); continue; }
            word.push(c);
            // Word may also break a line right after a hyphen or dash that sits inside a word.
            const next = chars[i + 1];
            if (/[-‐–—]/.test(c.ch) && next && next.ch !== ' ' && next.ch !== '\n' && word.length > 1) flushWord();
        }
        flushWord();
        while (line.length && line[line.length - 1].ch === ' ') line.pop();
        lines.push(line);
        return lines;
    }

    function layoutBox(ctx, spec, record) {
        const values = fieldValues(record);
        const box = spec.box;
        const contentWidth = (box.width - box.insetLeft - box.insetRight) * PX_PER_PT;
        const laid = [];
        let y = box.insetTop * PX_PER_PT;
        spec.paragraphs.forEach((paragraph) => {
            const chars = fillParagraph(paragraph, values);
            if (chars === null) return;
            const available = contentWidth - (paragraph.leftIndent + paragraph.rightIndent) * PX_PER_PT;
            y += paragraph.spaceBefore * PX_PER_PT;
            // An empty paragraph still takes up one line, as in Word.
            const lines = wrap(ctx, chars, available - Math.max(0, paragraph.firstLineIndent) * PX_PER_PT);
            lines.forEach((line) => {
                const runs = line.length ? [...new Set(line.map((c) => c.run))] : [paragraph.runs[0]];
                let ascent = 0, descent = 0;
                runs.forEach((run) => { const m = fontMetrics(ctx, run, spec.fontMetrics); ascent = Math.max(ascent, m.ascent); descent = Math.max(descent, m.descent); });
                const natural = ascent + descent;
                const spacing = paragraph.line;
                const height = spacing.rule === 'multiple' ? natural * spacing.value
                    : spacing.rule === 'atLeast' ? Math.max(natural, spacing.value * PX_PER_PT)
                        : spacing.value * PX_PER_PT;
                // Word puts a line's extra 1.5x spacing below it: the baseline sits one font ascent under the line top.
                laid.push({ chars: line, width: measure(ctx, line), top: y, height, natural, baseline: y + ascent, paragraph });
                y += height;
            });
            y += paragraph.spaceAfter * PX_PER_PT;
        });
        return { lines: laid, bottom: y };
    }

    function drawBox(spec, record) {
        const box = spec.box;
        const measureCanvas = document.createElement('canvas');
        const mctx = measureCanvas.getContext('2d');
        const layout = layoutBox(mctx, spec, record);

        const shadow = spec.textShadow;
        const pad = shadow ? Math.ceil((shadow.blur + Math.abs(shadow.dx) + Math.abs(shadow.dy)) * PX_PER_PT) : 0;
        const width = Math.ceil(box.width * PX_PER_PT);
        // Only as tall as the text: the rest of the box is empty, and a transparent image of the whole
        // box would bloat the PDF for nothing. The box clips whatever does not fit, like Word.
        const boxBottom = Math.floor(box.height * PX_PER_PT);
        const height = Math.max(1, Math.min(Math.ceil(layout.bottom + pad), Math.ceil(box.height * PX_PER_PT)));
        const canvas = document.createElement('canvas');
        canvas.width = width;
        canvas.height = height;
        const ctx = canvas.getContext('2d');
        ctx.textBaseline = 'alphabetic';
        ctx.textAlign = 'left';
        if (shadow) {
            ctx.shadowColor = `rgba(0,0,0,${shadow.alpha})`;
            ctx.shadowOffsetX = shadow.dx * PX_PER_PT;
            ctx.shadowOffsetY = shadow.dy * PX_PER_PT;
            ctx.shadowBlur = shadow.blur * PX_PER_PT;
        }

        const contentLeft = box.insetLeft * PX_PER_PT;
        const contentWidth = (box.width - box.insetLeft - box.insetRight) * PX_PER_PT;
        layout.lines.forEach((line) => {
            if (!line.chars.length) return;
            // Word shows a line only if the line itself (not its extra spacing) fits inside the box.
            if (line.top + line.natural > boxBottom + 1) return;
            const p = line.paragraph;
            const left = contentLeft + p.leftIndent * PX_PER_PT;
            const room = contentWidth - (p.leftIndent + p.rightIndent) * PX_PER_PT;
            let x = left;
            if (p.align === 'center') x = left + (room - line.width) / 2;
            else if (p.align === 'right') x = left + room - line.width;
            let i = 0;
            while (i < line.chars.length) {
                let j = i;
                while (j < line.chars.length && line.chars[j].run === line.chars[i].run) j++;
                const run = line.chars[i].run;
                const text = line.chars.slice(i, j).map((c) => displayChar(c.ch, run)).join('');
                applyRun(ctx, run);
                ctx.fillStyle = run.color;
                let rise = 0;
                if (run.superscript) rise = -run.size * PX_PER_PT * 0.35;
                else if (run.subscript) rise = run.size * PX_PER_PT * 0.15;
                ctx.fillText(text, x, line.baseline + rise);
                const w = ctx.measureText(text).width;
                if (run.underline) ctx.fillRect(x, line.baseline + 2, w, Math.max(1, run.size * PX_PER_PT / 18));
                x += w;
                i = j;
            }
        });
        return canvas;
    }

    function canvasToPngBytes(canvas) {
        return new Promise((resolve, reject) => {
            canvas.toBlob(async (blob) => {
                if (!blob) { reject(new Error('Could not draw the certificate text.')); return; }
                resolve(new Uint8Array(await blob.arrayBuffer()));
            }, 'image/png');
        });
    }

    // One PDF containing a page per record.
    async function build(key, records) {
        if (!window.PDFLib) throw new Error('The PDF library (pdf-lib) failed to load.');
        const { PDFDocument } = window.PDFLib;
        const { spec, pdf } = await loadAssets(key);
        const out = await PDFDocument.create();
        const [background] = await out.embedPdf(pdf, [0]);
        const pageW = spec.page.width;
        const pageH = spec.page.height;
        for (const record of records) {
            const canvas = drawBox(spec, record);
            const image = await out.embedPng(await canvasToPngBytes(canvas));
            const page = out.addPage([pageW, pageH]);
            page.drawPage(background, { x: 0, y: 0, width: pageW, height: pageH });
            const h = canvas.height / PX_PER_PT;
            page.drawImage(image, { x: spec.box.x, y: pageH - spec.box.y - h, width: canvas.width / PX_PER_PT, height: h });
        }
        const bytes = await out.save();
        return new Blob([bytes], { type: 'application/pdf' });
    }

    window.PdfExport = {
        available: () => !!window.PDFLib,
        build,
        missingFontsFor: async (key) => missingFonts((await loadAssets(key)).spec),
        // Test hook: baselines (pt from the page top) of every laid-out line.
        _baselines: async (key, record) => {
            const { spec } = await loadAssets(key);
            const ctx = document.createElement('canvas').getContext('2d');
            return layoutBox(ctx, spec, record).lines.filter((l) => l.chars.length).map((l) => ({ text: l.chars.map((c) => c.ch).join('').slice(0, 30), base: +(spec.box.y + l.baseline / PX_PER_PT).toFixed(2), x: +(spec.box.x + (spec.box.insetLeft + (spec.box.width - spec.box.insetLeft - spec.box.insetRight - l.width / PX_PER_PT) / 2)).toFixed(1) }));
        },
    };
})();
