"""
Write the stacked RAW layer. Importable, so the model generator can put these
same tabs into the model workbook.

    python scripts/build_raw_stacked.py      # RAW-only workbook, for inspection

One tab per data type, every block stacked inside it behind a key column. This
is a deliberate design choice, not just tidiness:

  * CLAUDE.md bans INDIRECT and OFFSET and requires INDEX/MATCH on the code.
    With one tab per region a region switch has only two possible shapes - a
    nine-deep nested IF repeated in every formula, or INDIRECT to build the tab
    name. Stacked, the region becomes a value you can MATCH on and the whole
    switch collapses to one pattern.

  * Table 5 and Table 8 disagree about where the primary inputs sit (T5 has P6
    at source row 148; T8 has P5 there and P6 at 151). The RowType column makes
    the output denominator a SUMIF over row type instead of a hardcoded row
    range, which is what CLAUDE.md means by "sum the P-rows dynamically".

  * Lookups are by key, not by position, so a block pasted a row out still
    resolves. Position only matters to the paste workflow and its QA gate.

Layout of every stacked tab - keys first, then the source block verbatim:

    Region  Code  RowType  SrcRow  Key      | source block, pasted verbatim
    NSW     0101  Industry 13      NSW|0101 | (source col A lands in F, data H+)

The Key column is a formula joining the identity columns. It is what lets every
downstream lookup be a single-criterion MATCH rather than an array formula.

Keys sit OUTSIDE the pasted rectangle, so re-pasting a block cannot disturb
them and they cannot disturb the data.

RAW holds source data verbatim - rule 1. Nothing here trims, pads, re-spines,
blanks or coerces. 'n.a.' stays text, Total columns stay, Dummy rows stay,
primary-input rows stay. All reshaping belongs in the MAP_ layer.
"""
import pickle
from datetime import date
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter as CL

ROOT = Path(__file__).resolve().parent.parent
SOURCES = ROOT / 'build' / 'sources.pkl'
OUTFILE = ROOT / 'output' / 'IO_RAW_Stacked.xlsx'

REGIONS = ['Aus', 'NSW', 'Vic', 'QLD', 'SA', 'WA', 'Tas', 'NT', 'ACT']
HDR_ROWS = 2
FIRST_DATA_ROW = 4 + HDR_ROWS + 1     # banner(3) + key header(1) + 2 source headers

H1 = Font(bold=True, size=13, color='FFFFFF')
H2 = Font(bold=True, size=10)
KEYFONT = Font(bold=False, size=9, color='333333')
BLUE = Font(color='0000CC')            # hardcoded input, per the workbook colour key
BLACK = Font(color='000000')           # formula
GREEN = Font(color='006100')           # cross-sheet link
NOTE = Font(italic=True, size=9, color='666666')
TITLEFILL = PatternFill('solid', fgColor='1F3864')
KEYFILL = PatternFill('solid', fgColor='DDEBF7')
HDRFILL = PatternFill('solid', fgColor='F2F2F2')
BANDFILL = PatternFill('solid', fgColor='FFF2CC')
YELLOW = PatternFill('solid', fgColor='FFF2CC')
THIN = Side(style='thin', color='BFBFBF')

# Column geometry of the source blocks, 1-based within the source sheet.
# Provider Table 5 / Table 8:  C..DV industries (114 real + 10 dummy),
#   DW Total Industry Uses, DX..ED Q1..Q7, EE Final Uses, EF Total Supply.
FLOW_IND_FIRST, FLOW_IND_LAST = 3, 126
FLOW_Q_FIRST = 128                     # Q1 Households
# ABS margin / tax cubes: C..DM the 115 industry columns, DN Total Industry
#   Uses, DO..DU Q1..Q7, DV Final Uses, DW Total Supply.
MARG_IND_FIRST, MARG_IND_LAST = 3, 117
MARG_Q_FIRST = 119

REGION_KEYS = ['Region', 'Code', 'RowType', 'SrcRow', 'Key']
MARGIN_KEYS = ['ABS_Table', 'Margin', 'Code', 'RowType', 'SrcRow', 'Key']
FLOW_VERB_COL = len(REGION_KEYS) + 1   # source col 1 lands here (F = 6)
MARG_VERB_COL = len(MARGIN_KEYS) + 1   # source col 1 lands here (G = 7)


def flow_col(src_col):
    """Model column for a source column of a flow / multiplier block."""
    return len(REGION_KEYS) + src_col


def marg_col(src_col):
    """Model column for a source column of an ABS margin block."""
    return len(MARGIN_KEYS) + src_col


def write_stacked(wb, tabname, title, blocks, note, keydefs=None, hdr_src=(0, 1)):
    """
    blocks:  list of (block_id, loaded_block). Every block gets the same band
             height so a short one leaves trailing rows empty rather than
             shifting the ones below it.
    keydefs: list of (header, fn(block_id, meta, rownum) -> value). A value
             beginning with '=' is written as a formula.
    """
    if keydefs is None:
        keydefs = [('Region', lambda b, m, r: b),
                   ('Code', lambda b, m, r: m['code'] or m['raw_code']),
                   ('RowType', lambda b, m, r: m['row_type']),
                   ('SrcRow', lambda b, m, r: m['src_row']),
                   ('Key', lambda b, m, r: '=A%d&"|"&B%d' % (r, r))]
    nkey = len(keydefs)
    ws = wb.create_sheet(tabname)
    band = max(len(b['verbatim']) for _, b in blocks)
    width = max(len(r) for _, b in blocks for r in b['verbatim'])

    ws.cell(row=1, column=1, value=title).font = H1
    for c in range(1, nkey + width + 1):
        ws.cell(row=1, column=c).fill = TITLEFILL
    src = blocks[0][1].get('source', '')
    ws.cell(row=2, column=1, value=(
        f"Source: {src.split(' :: ')[0]}   |   vintage {blocks[0][1].get('vintage', '')}"
        f"   |   loaded {date.today():%Y-%m-%d}   |   {len(blocks)} blocks"
        f"   |   band height {band} rows")).font = NOTE
    ws.cell(row=3, column=1, value=note).font = NOTE

    hr = 4
    for i, (k, _) in enumerate(keydefs, 1):
        c = ws.cell(row=hr, column=i, value=k)
        c.font = H2
        c.fill = KEYFILL
        c.border = Border(bottom=THIN)
    ws.cell(row=hr, column=nkey + 1, value='-- source block, verbatim -->').font = H2
    for j, hi in enumerate(hdr_src):
        hrow = blocks[0][1]['header'][hi]
        r = hr + 1 + j
        for ci, v in enumerate(hrow, 1):
            cell = ws.cell(row=r, column=nkey + ci, value=v)
            cell.font = H2
            cell.fill = HDRFILL
            cell.alignment = Alignment(horizontal='center')

    bands = []
    row = FIRST_DATA_ROW
    for bid, b in blocks:
        start = row
        for m, vals in zip(b['meta'], b['verbatim']):
            for i, (name, fn) in enumerate(keydefs, 1):
                c = ws.cell(row=row, column=i, value=fn(bid, m, row))
                c.font = BLACK if name == 'Key' else KEYFONT
                c.fill = KEYFILL
                if name == 'Code':
                    c.number_format = '@'      # codes are TEXT, always
            for ci, v in enumerate(vals, 1):
                cell = ws.cell(row=row, column=nkey + ci, value=v)
                cell.font = BLUE
                if ci == 1:
                    cell.number_format = '@'
            row += 1
        while row - start < band:              # pad short blocks
            for i, (name, fn) in enumerate(keydefs, 1):
                blank = {'code': '', 'raw_code': '', 'row_type': '', 'src_row': None}
                v = '' if name in ('Code', 'RowType', 'SrcRow', 'Key') else fn(bid, blank, row)
                c = ws.cell(row=row, column=i, value=v)
                c.font = KEYFONT
                c.fill = KEYFILL
            row += 1
        bands.append((bid, start, row - 1, b.get('source', '')))
        ws.cell(row=start, column=1).fill = BANDFILL

    ws.freeze_panes = ws.cell(row=FIRST_DATA_ROW, column=nkey + 1)
    ws.auto_filter.ref = f"A{hr}:{CL(nkey)}{row - 1}"
    widths = {'Region': 8, 'Code': 9, 'RowType': 11, 'SrcRow': 8,
              'ABS_Table': 10, 'Margin': 15, 'Key': 14}
    for i, (name, _) in enumerate(keydefs, 1):
        ws.column_dimensions[CL(i)].width = widths.get(name, 10)
    ws.column_dimensions[CL(nkey + 1)].width = 9
    ws.column_dimensions[CL(nkey + 2)].width = 34
    return {'bands': bands, 'band': band, 'width': width, 'last': row - 1,
            'first': FIRST_DATA_ROW, 'nkey': nkey}


def write_raw_tabs(wb, src, verbose=True):
    """Add every RAW tab to `wb`. Returns geometry the model layer needs."""
    flows, mult = src['flows'], src['multipliers']
    geo = {}
    for kind, tab, title, note in [
        ('T5', 'RAW_T5', 'RAW - Table 5, all regions stacked. Industry by industry, '
         'DIRECT allocation of imports (domestic only), $m 2022-23',
         'Verbatim. Domestic multipliers come from this table. Imports sit in the P6 '
         'primary-input row, not in the intermediate quadrant.'),
        ('T8', 'RAW_T8', 'RAW - Table 8, all regions stacked. Industry by industry, '
         'INDIRECT allocation of imports (imports embedded), $m 2022-23',
         'Verbatim. Used ONLY to derive imports as T8 less T5, cell by cell. '
         'Never build multipliers off this table.'),
    ]:
        blocks = [(r, flows[(kind, r)]) for r in REGIONS if (kind, r) in flows]
        geo[tab] = write_stacked(wb, tab, title, blocks, note)
        if verbose:
            print(f"  {tab:16s} {len(geo[tab]['bands'])} regions, band "
                  f"{geo[tab]['band']}, last row {geo[tab]['last']}")

    mblocks = [(r, mult['regions'][r]) for r in REGIONS if r in mult.get('regions', {})]
    geo['RAW_Multipliers'] = write_stacked(
        wb, 'RAW_Multipliers',
        'RAW - Multiplier set, all regions stacked. 15 measure blocks x 11 effects, '
        'FY23 (2022-23)', mblocks,
        "Verbatim, INCLUDING the 'n.a.' text cells - the MAP layer decides what they "
        "mean. Multipliers are an input and are never derived. 'Value added multipliers' "
        "is the basic-prices block and the ABS headline; 'at market prices' includes "
        "taxes on products and is NOT the headline.")
    if verbose:
        print(f"  RAW_Multipliers  {len(geo['RAW_Multipliers']['bands'])} regions, "
              f"last row {geo['RAW_Multipliers']['last']}")

    abs_t = src.get('abs') or {}
    mt = src['margin_tables']
    names = dict({k: v[0] for k, v in mt.items()}, **{'35': 'NetTaxes'})
    mgblocks = [(k, abs_t['T' + k]) for k in sorted(mt, key=int) + ['35']
                if 'T' + k in abs_t]
    if mgblocks:
        keydefs = [('ABS_Table', lambda b, m, r: int(b)),
                   ('Margin', lambda b, m, r: names.get(b, '')),
                   ('Code', lambda b, m, r: m['code'] or m['raw_code']),
                   ('RowType', lambda b, m, r: m['row_type']),
                   ('SrcRow', lambda b, m, r: m['src_row']),
                   ('Key', lambda b, m, r: '=B%d&"|"&C%d' % (r, r))]
        geo['RAW_Margins'] = write_stacked(
            wb, 'RAW_Margins',
            'RAW - ABS Tables 23-34 (margins) and 35 (net taxes), stacked. '
            'Product by using industry and final use, $m 2023-24. NATIONAL ONLY',
            mgblocks,
            'Verbatim, including the Re-exports row and the Total row and columns. '
            'Filter on ABS_Table or Margin to isolate one table. Tables 23-34 sum to '
            '$421,410m on the 115-code spine and $422,034m including re-exports; '
            'Table 35 is $168,673m. Margin rates are national - the ABS publishes no '
            'state margin or tax matrices.', keydefs=keydefs)
        if verbose:
            print(f"  RAW_Margins      {len(geo['RAW_Margins']['bands'])} tables, "
                  f"last row {geo['RAW_Margins']['last']}")
    if 'T21' in abs_t:
        geo['RAW_T21_Control'] = write_stacked(
            wb, 'RAW_T21_Control',
            'RAW - ABS Table 21. Composition of supply by margin commodity, '
            '$m 2023-24. The independent control total', [('21', abs_t['T21'])],
            'Verbatim. Margin commodity total is $422,034m, which is Tables 23-34 '
            'including their Re-exports rows. Use as the gate on any re-paste.',
            keydefs=[('ABS_Table', lambda b, m, r: int(b)),
                     ('Code', lambda b, m, r: m['code'] or m['raw_code']),
                     ('RowType', lambda b, m, r: m['row_type']),
                     ('SrcRow', lambda b, m, r: m['src_row']),
                     ('Key', lambda b, m, r: '=A%d&"|"&B%d' % (r, r))],
            hdr_src=(0,))
    return geo


def write_margin_map(wb, src):
    """The margin -> earning IOIG assignment. A mapping, so NOT in a RAW tab."""
    mt = src['margin_tables']
    abs_t = src.get('abs') or {}
    ctrl = src.get('margin_control', {})
    lm = wb.create_sheet('Lists_MarginMap')
    lm.cell(row=1, column=1, value='Lists - margin table to earning IOIG').font = \
        Font(bold=True, size=13)
    lm.cell(row=2, column=1, value=(
        'This is a mapping, not source data, so it lives here rather than in a RAW tab '
        '(rule 2). Verified against ABS Table 21 to the dollar. Join to RAW_Margins on '
        'ABS_Table, or to the strip on Margin.')).font = NOTE
    for i, h in enumerate(['ABS_Table', 'Margin', 'Earning IOIG', 'Spine $m',
                           'Re-exports $m'], 1):
        c = lm.cell(row=4, column=i, value=h)
        c.font = H2
        c.fill = KEYFILL
    r = 4
    for k in sorted(mt, key=int):
        r += 1
        name, earner = mt[k]
        lm.cell(row=r, column=1, value=int(k))
        lm.cell(row=r, column=2, value=name)
        e = lm.cell(row=r, column=3, value=earner)
        e.number_format = '@'
        lm.cell(row=r, column=4, value=ctrl.get(name))
        if 'T' + k in abs_t:
            rex = next((row[126] for row, m in zip(abs_t['T' + k]['verbatim'],
                                                   abs_t['T' + k]['meta'])
                        if m['row_type'] == 'ReExports'), None)
            lm.cell(row=r, column=5, value=rex)
    first_map, last_map = 5, r
    r += 1
    lm.cell(row=r, column=2, value='TOTAL 23-34').font = H2
    lm.cell(row=r, column=4, value=src.get('margin_total_spine')).font = H2
    r += 2
    lm.cell(row=r, column=2, value='Net taxes (Table 35)')
    lm.cell(row=r, column=3, value='n/a - leakage to government')
    lm.cell(row=r, column=4, value=src.get('net_tax_total_spine'))
    for col, w in zip('ABCDE', [11, 16, 26, 12, 14]):
        lm.column_dimensions[col].width = w
    return {'first': first_map, 'last': last_map}


def main():
    if not SOURCES.exists():
        raise SystemExit(f"{SOURCES} not found. Run scripts/load_sources.py first.")
    src = pickle.load(open(SOURCES, 'rb'))
    if not src['flows'] or not src['multipliers']:
        raise SystemExit("sources.pkl has no supplied data. Run load_sources.py.")
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    write_raw_tabs(wb, src)
    write_margin_map(wb, src)
    OUTFILE.parent.mkdir(exist_ok=True)
    wb.save(OUTFILE)
    print(f"\nwrote {OUTFILE}  ({OUTFILE.stat().st_size / 1e6:.1f} MB)")


if __name__ == '__main__':
    main()
