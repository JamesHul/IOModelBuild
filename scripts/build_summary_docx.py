"""Render docs/IO_Summary.md into the ACIL Allen Word template used by the
CGE companion note, so the two documents are typographically identical.

The template is the CGE .docx itself: styles, numbering, theme, headers,
footers and the contact block are all reused unchanged. Only the body
content between the title block and the contact block is replaced.
"""
import argparse, re, shutil, subprocess, sys, tempfile, zipfile
from pathlib import Path

ap = argparse.ArgumentParser(description=__doc__,
                             formatter_class=argparse.RawDescriptionHelpFormatter)
ap.add_argument('template', help='a .docx built on the house template - the CGE '
                                 'companion note is the one this was written against')
ap.add_argument('-i', '--input', default='docs/IO_Summary.md')
ap.add_argument('-o', '--output', default='docs/IO_Summary.docx')
args = ap.parse_args()

SRC_MD = Path(args.input).resolve()
OUT    = Path(args.output).resolve()
_tmp   = Path(tempfile.mkdtemp(prefix='io_summary_docx_'))
TPL    = _tmp / 'template'
WORK   = _tmp / 'build'
with zipfile.ZipFile(args.template) as z:
    z.extractall(TPL)
for link in TPL.rglob('*'):          # source .docx is untrusted - no symlinks
    if link.is_symlink():
        link.unlink()

GRIDS = {2: [1985, 6520], 3: [1985, 3260, 3260]}   # dxa, must sum to 8505
GRID_T3 = [2127, 6378]

# ----------------------------------------------------------------- helpers
def esc(s):
    return (s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;'))

def runs(text, base_bold=False):
    """Inline **bold** / *italic* -> a sequence of w:r elements."""
    out, i = [], 0
    for m in re.finditer(r'\*\*(.+?)\*\*|\*(.+?)\*', text):
        if m.start() > i:
            out.append((text[i:m.start()], base_bold, False))
        if m.group(1) is not None:
            out.append((m.group(1), True, False))
        else:
            out.append((m.group(2), base_bold, True))
        i = m.end()
    if i < len(text):
        out.append((text[i:], base_bold, False))
    xml = []
    for t, b, it in out:
        if t == '':
            continue
        rpr = ''
        if b or it:
            rpr = '<w:rPr>' + ('<w:b/><w:bCs/>' if b else '') + \
                  ('<w:i/><w:iCs/>' if it else '') + '</w:rPr>'
        xml.append(f'<w:r>{rpr}<w:t xml:space="preserve">{esc(t)}</w:t></w:r>')
    return ''.join(xml)

def para(style, text):
    return f'<w:p><w:pPr><w:pStyle w:val="{style}"/></w:pPr>{runs(text)}</w:p>'

def bullet(text):
    """Bold lead-in convention of the companion: bold the phrase, period outside."""
    m = re.match(r'\*\*(.+?)\*\*(.*)$', text, re.S)
    if m:
        lead, rest = m.group(1), m.group(2)
        if lead.endswith('.'):
            lead, rest = lead[:-1], '.' + rest
        body = (f'<w:r><w:rPr><w:b/><w:bCs/></w:rPr>'
                f'<w:t xml:space="preserve">{esc(lead)}</w:t></w:r>' + runs(rest))
    else:
        body = runs(text)
    return f'<w:p><w:pPr><w:pStyle w:val="ListBullet"/></w:pPr>{body}</w:p>'

def caption(num, text):
    return (
        '<w:p><w:pPr><w:pStyle w:val="Caption"/>'
        '<w:ind w:left="1200" w:hanging="1200"/></w:pPr>'
        '<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Table </w:t></w:r>'
        '<w:r><w:rPr><w:b/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>'
        '<w:r><w:rPr><w:b/></w:rPr><w:instrText xml:space="preserve">'
        ' SEQ Table \\* ARABIC \\s 1 \\* MERGEFORMAT </w:instrText></w:r>'
        '<w:r><w:rPr><w:b/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r>'
        f'<w:r><w:rPr><w:b/><w:noProof/></w:rPr><w:t>{num}</w:t></w:r>'
        '<w:r><w:rPr><w:b/></w:rPr><w:fldChar w:fldCharType="end"/></w:r>'
        f'<w:r><w:tab/></w:r>{runs(text)}</w:p>')

def cell(w, style, text):
    return (f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/></w:tcPr>'
            f'<w:p><w:pPr><w:pStyle w:val="{style}"/></w:pPr>{runs(text)}</w:p></w:tc>')

def table(rows, grid):
    ncol = len(grid)
    hdr_cnf = ('<w:cnfStyle w:val="100000000000" w:firstRow="1" w:lastRow="0"'
               ' w:firstColumn="0" w:lastColumn="0" w:oddVBand="0" w:evenVBand="0"'
               ' w:oddHBand="0" w:evenHBand="0" w:firstRowFirstColumn="0"'
               ' w:firstRowLastColumn="0" w:lastRowFirstColumn="0"'
               ' w:lastRowLastColumn="0"/>')
    xml = ['<w:tbl><w:tblPr><w:tblStyle w:val="aacTableBasic"/>'
           '<w:tblW w:w="4907" w:type="pct"/><w:tblLayout w:type="fixed"/>'
           '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1"'
           ' w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr><w:tblGrid>'
           + ''.join(f'<w:gridCol w:w="{w}"/>' for w in grid) + '</w:tblGrid>']
    for ri, row in enumerate(rows):
        if ri == 0:
            xml.append(f'<w:tr><w:trPr>{hdr_cnf}<w:cantSplit/><w:tblHeader/></w:trPr>')
            style = 'Tablecolumnheadings'
        else:
            xml.append('<w:tr><w:trPr><w:cantSplit/></w:trPr>')
            style = 'Tabletext'
        for ci in range(ncol):
            xml.append(cell(grid[ci], style, row[ci] if ci < len(row) else ''))
        xml.append('</w:tr>')
    xml.append('</w:tbl>')
    return ''.join(xml)

# ------------------------------------------------------------------ parse
raw = SRC_MD.read_text(encoding='utf8')
meta = dict(re.findall(r'^\s*(Title|Subtitle|Date):\s*(.+?)\s*$',
                       raw[:raw.index('-->')], re.M))
body_md = raw[raw.index('-->') + 3:].strip('\n')

lines = body_md.split('\n')
blocks, i, tnum = [], 0, 0
while i < len(lines):
    ln = lines[i]
    s = ln.strip()
    if s == '' or s == '---':
        i += 1; continue
    if s.startswith('|'):
        rows = []
        while i < len(lines) and lines[i].strip().startswith('|'):
            cells = [c.strip() for c in lines[i].strip().strip('|').split('|')]
            if not all(re.fullmatch(r':?-{2,}:?', c) for c in cells):
                rows.append(cells)
            i += 1
        blocks.append(('table', rows)); continue
    if s.startswith('### '):
        blocks.append(('h3', s[4:])); i += 1; continue
    if s.startswith('## '):
        blocks.append(('h2', s[3:])); i += 1; continue
    if s.startswith('# '):
        blocks.append(('h1', s[2:])); i += 1; continue
    m = re.fullmatch(r'\*Table (\d+)\s+(.+?)\*', s)
    if m:
        tnum = int(m.group(1))
        blocks.append(('caption', (tnum, m.group(2)))); i += 1; continue
    if re.fullmatch(r'\*Source:.+\*', s):
        blocks.append(('source', s.strip('*'))); i += 1; continue
    if s.startswith('- '):
        buf = [s[2:]]; i += 1
        while i < len(lines) and lines[i].startswith('  ') and lines[i].strip():
            buf.append(lines[i].strip()); i += 1
        blocks.append(('bullet', ' '.join(buf))); continue
    buf = [s]; i += 1
    while i < len(lines) and lines[i].strip() and not lines[i].strip().startswith(
            ('- ', '#', '|', '---')) and not re.fullmatch(
            r'\*(Table \d+.*|Source:.*)\*', lines[i].strip()):
        buf.append(lines[i].strip()); i += 1
    blocks.append(('para', ' '.join(buf)))

# ------------------------------------------------------------------ emit
xml, seen_h1, tbl_i = [], False, 0
for kind, val in blocks:
    if kind == 'h1':
        xml.append(para('Heading1ES', val)); seen_h1 = True
    elif kind == 'h2':
        xml.append(para('Heading2ES', val))
    elif kind == 'h3':
        xml.append(para('Heading3ES', val))
    elif kind == 'caption':
        xml.append(caption(*val))
    elif kind == 'source':
        xml.append(para('Source', val))
    elif kind == 'bullet':
        xml.append(bullet(val))
    elif kind == 'table':
        tbl_i += 1
        grid = GRID_T3 if (tbl_i == 3) else GRIDS[len(val[0])]
        xml.append(table(val, grid))
    elif kind == 'para':
        if val.startswith('*Introduction.*'):
            xml.append(para('Introduction', val[len('*Introduction.*'):].strip()))
        else:
            xml.append(para('BodyText', val))
CONTENT = ''.join(xml)

# ------------------------------------------------------------------ splice
shutil.copytree(TPL, WORK)
doc = (WORK / 'word/document.xml').read_text(encoding='utf8')

# 1. title block text
for old, new in [('Tasman Global CGE model overview', meta['Title']),
                 ('Short internal note on our CGE modelling capability', meta['Subtitle']),
                 ('3 September 2026', meta['Date'])]:
    assert old in doc, old
    doc = doc.replace(old, esc(new), 1)

# 2. body content: everything between the title table and the trailing
#    empty paragraphs that precede the contact block
start = doc.index('</w:tbl>') + len('</w:tbl>')

# Walk top-level <w:tbl> elements to find the contact block (the last one),
# then step back over the empty spacer paragraphs that precede it.
depth, tops = 0, []
for m in re.finditer(r'<w:tbl>|</w:tbl>', doc):
    if m.group(0) == '<w:tbl>':
        if depth == 0:
            top_start = m.start()
        depth += 1
    else:
        depth -= 1
        if depth == 0:
            tops.append((top_start, m.end()))
contact_start = tops[-1][0]
assert 'acilallen.com.au' in doc[contact_start:tops[-1][1]], 'last table is not the contact block'

tail_start = contact_start
spacer = re.compile(r'<w:p [^>]*><w:pPr><w:pStyle w:val="BodyText"/></w:pPr></w:p>$')
while True:
    m = spacer.search(doc, 0, tail_start)
    if m and m.end() == tail_start:
        tail_start = m.start()
    else:
        break

doc = doc[:start] + CONTENT + doc[tail_start:]
(WORK / 'word/document.xml').write_text(doc, encoding='utf8')

# 2b. footers cache the title/subtitle as STYLEREF field results. Word
#     refreshes them on open, but fix the cached text so it is never stale.
def fix_fldsimple(xml, key, first, rest):
    pat = re.compile(
        r'(<w:fldSimple w:instr=" STYLEREF &quot;' + key + r'&quot;[^"]*">)(.*?)(</w:fldSimple>)',
        re.S)
    def repl(m):
        body = ('<w:r><w:rPr><w:b/><w:bCs/><w:noProof/></w:rPr>'
                f'<w:t xml:space="preserve">{esc(first)}</w:t></w:r>'
                '<w:r><w:rPr><w:noProof/></w:rPr>'
                f'<w:t xml:space="preserve">{esc(rest)}</w:t></w:r>')
        return m.group(1) + body + m.group(3)
    return pat.subn(repl, xml)

title_first, _, title_rest = meta['Title'].partition(' ')
sub_words = meta['Subtitle'].split(' ')
sub_first, sub_rest = ' '.join(sub_words[:2]), ' ' + ' '.join(sub_words[2:])
fixed = 0
for f in sorted((WORK / 'word').glob('footer*.xml')) + sorted((WORK / 'word').glob('header*.xml')):
    x = f.read_text(encoding='utf8')
    x, a = fix_fldsimple(x, 'Cp Title', title_first, ' ' + title_rest)
    x, b = fix_fldsimple(x, 'Cp SubTitle', sub_first, sub_rest)
    if a or b:
        f.write_text(x, encoding='utf8')
        fixed += a + b
print(f'footer/header STYLEREF fields refreshed: {fixed}')

# 3. document properties
core = (WORK / 'docProps/core.xml').read_text(encoding='utf8')
core = re.sub(r'<dc:title>.*?</dc:title>', f'<dc:title>{esc(meta["Title"])}</dc:title>', core)
core = re.sub(r'<dc:subject>.*?</dc:subject>', f'<dc:subject>{esc(meta["Subtitle"])}</dc:subject>', core)
(WORK / 'docProps/core.xml').write_text(core, encoding='utf8')

if OUT.exists():
    OUT.unlink()
subprocess.run(['zip', '-Xrq', str(OUT), '.'], cwd=WORK, check=True)
shutil.rmtree(_tmp, ignore_errors=True)
print(f'wrote {OUT}  ({OUT.stat().st_size:,} bytes)')
print(f'blocks: {len(blocks)}  tables: {tbl_i}  content xml: {len(CONTENT):,} chars')
