// Render docs/IO_Summary.md as a plain, unbranded Word document.
// Built-in Word styles only - no template, no logos, no company assets.
const fs = require('fs');
const d = require('docx');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell,
  WidthType, AlignmentType, BorderStyle, ShadingType, LevelFormat, PageOrientation,
} = d;

const MD = process.argv[2] || 'docs/IO_Summary.md';
const OUT = process.argv[3] || 'docs/IO_Summary_plain.docx';

const raw = fs.readFileSync(MD, 'utf8');
const head = raw.slice(0, raw.indexOf('-->'));
const meta = {};
for (const m of head.matchAll(/^\s*(Title|Subtitle|Date):\s*(.+?)\s*$/gm)) meta[m[1]] = m[2];
const bodyMd = raw.slice(raw.indexOf('-->') + 3).replace(/^\n+/, '');

// ---------------------------------------------------------------- parse
const lines = bodyMd.split('\n');
const blocks = [];
let i = 0;
while (i < lines.length) {
  const s = lines[i].trim();
  if (s === '' || s === '---') { i++; continue; }
  if (s.startsWith('|')) {
    const rows = [];
    while (i < lines.length && lines[i].trim().startsWith('|')) {
      const cells = lines[i].trim().replace(/^\||\|$/g, '').split('|').map(c => c.trim());
      if (!cells.every(c => /^:?-{2,}:?$/.test(c))) rows.push(cells);
      i++;
    }
    blocks.push({ k: 'table', v: rows }); continue;
  }
  if (s.startsWith('### ')) { blocks.push({ k: 'h3', v: s.slice(4) }); i++; continue; }
  if (s.startsWith('## '))  { blocks.push({ k: 'h2', v: s.slice(3) }); i++; continue; }
  if (s.startsWith('# '))   { blocks.push({ k: 'h1', v: s.slice(2) }); i++; continue; }
  let m = s.match(/^\*Table (\d+)\s+(.+?)\*$/);
  if (m) { blocks.push({ k: 'caption', n: m[1], v: m[2] }); i++; continue; }
  if (/^\*Source:.+\*$/.test(s)) { blocks.push({ k: 'source', v: s.replace(/^\*|\*$/g, '') }); i++; continue; }
  if (s.startsWith('- ')) {
    const buf = [s.slice(2)]; i++;
    while (i < lines.length && lines[i].startsWith('  ') && lines[i].trim()) { buf.push(lines[i].trim()); i++; }
    blocks.push({ k: 'bullet', v: buf.join(' ') }); continue;
  }
  const buf = [s]; i++;
  while (i < lines.length && lines[i].trim() && !/^(- |#|\||---)/.test(lines[i].trim())
         && !/^\*(Table \d+.*|Source:.*)\*$/.test(lines[i].trim())) { buf.push(lines[i].trim()); i++; }
  blocks.push({ k: 'para', v: buf.join(' ') });
}

// ------------------------------------------------------------- inline
function inline(text, opts = {}) {
  const out = [];
  let pos = 0;
  const re = /\*\*(.+?)\*\*|\*(.+?)\*/g;
  let m;
  while ((m = re.exec(text)) !== null) {
    if (m.index > pos) out.push(new TextRun({ text: text.slice(pos, m.index), ...opts }));
    if (m[1] !== undefined) out.push(new TextRun({ text: m[1], bold: true, ...opts }));
    else out.push(new TextRun({ text: m[2], italics: true, ...opts }));
    pos = re.lastIndex;
  }
  if (pos < text.length) out.push(new TextRun({ text: text.slice(pos), ...opts }));
  return out;
}

// -------------------------------------------------------------- tables
const TOTAL = 9350;                      // dxa, A4 with 2cm margins
function buildTable(rows) {
  const n = rows[0].length;
  const widths = n === 2 ? [2200, TOTAL - 2200] : [2200, (TOTAL - 2200) / 2, (TOTAL - 2200) / 2];
  return new Table({
    columnWidths: widths,
    width: { size: TOTAL, type: WidthType.DXA },
    rows: rows.map((r, ri) => new TableRow({
      tableHeader: ri === 0,
      cantSplit: true,
      children: widths.map((w, ci) => new TableCell({
        width: { size: w, type: WidthType.DXA },
        shading: ri === 0 ? { type: ShadingType.CLEAR, fill: 'E7E6E6' } : undefined,
        margins: { top: 60, bottom: 60, left: 108, right: 108 },
        children: [new Paragraph({
          spacing: { before: 20, after: 20 },
          children: inline(r[ci] ?? '', ri === 0 ? { bold: true } : {}),
        })],
      })),
    })),
  });
}

// ------------------------------------------------------------- content
const kids = [
  new Paragraph({ text: meta.Title, heading: HeadingLevel.TITLE }),
  new Paragraph({ children: [new TextRun({ text: meta.Subtitle, italics: true, size: 24, color: '595959' })],
                  spacing: { after: 80 } }),
  new Paragraph({ children: [new TextRun({ text: meta.Date, size: 20, color: '595959' })],
                  spacing: { after: 360 },
                  border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: 'BFBFBF', space: 8 } } }),
];

let tableSeen = 0;
for (const b of blocks) {
  if (b.k === 'h1') continue;                       // the title block already carries it
  else if (b.k === 'h2') kids.push(new Paragraph({ text: b.v, heading: HeadingLevel.HEADING_1, spacing: { before: 320, after: 140 } }));
  else if (b.k === 'h3') kids.push(new Paragraph({ text: b.v, heading: HeadingLevel.HEADING_2, spacing: { before: 240, after: 120 } }));
  else if (b.k === 'caption') kids.push(new Paragraph({
      children: [new TextRun({ text: `Table ${b.n}  `, bold: true, size: 19 }),
                 ...inline(b.v, { size: 19 })],
      spacing: { before: 160, after: 80 } }));
  else if (b.k === 'source') kids.push(new Paragraph({
      children: [new TextRun({ text: b.v, size: 17, color: '595959' })], spacing: { after: 200 } }));
  else if (b.k === 'bullet') {
    let t = b.v;
    const m = t.match(/^\*\*(.+?)\*\*([\s\S]*)$/);
    let runs;
    if (m) {
      let lead = m[1], rest = m[2];
      if (lead.endsWith('.')) { lead = lead.slice(0, -1); rest = '.' + rest; }
      runs = [new TextRun({ text: lead, bold: true }), ...inline(rest)];
    } else runs = inline(t);
    kids.push(new Paragraph({ children: runs, numbering: { reference: 'bullets', level: 0 },
                              spacing: { after: 80 } }));
  }
  else if (b.k === 'table') { tableSeen++; kids.push(buildTable(b.v)); }
  else if (b.k === 'para') {
    const isIntro = b.v.startsWith('*Introduction.*');
    const text = isIntro ? b.v.slice('*Introduction.*'.length).trim() : b.v;
    kids.push(new Paragraph({
      children: inline(text, isIntro ? { size: 23 } : {}),
      spacing: { after: 140 },
    }));
  }
}

const doc = new Document({
  title: meta.Title,
  subject: meta.Subtitle,
  description: meta.Subtitle,
  styles: {
    default: {
      document: { run: { font: 'Calibri', size: 21 }, paragraph: { spacing: { line: 264 } } },
      title: { run: { font: 'Calibri Light', size: 44, bold: false, color: '1F3864' },
               paragraph: { spacing: { after: 60 } } },
      heading1: { run: { font: 'Calibri Light', size: 28, bold: true, color: '1F3864' } },
      heading2: { run: { font: 'Calibri Light', size: 23, bold: true, color: '2E5496' } },
    },
  },
  numbering: {
    config: [{
      reference: 'bullets',
      levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT,
                 style: { paragraph: { indent: { left: 340, hanging: 200 } } } }],
    }],
  },
  sections: [{
    properties: { page: { size: { orientation: PageOrientation.PORTRAIT },
                          margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 } } },
    children: kids,
  }],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(OUT, buf);
  console.log(`wrote ${OUT} (${buf.length.toLocaleString()} bytes)`);
  console.log(`blocks: ${blocks.length}  tables: ${tableSeen}  paragraphs+tables emitted: ${kids.length}`);
});
