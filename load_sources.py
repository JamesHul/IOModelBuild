"""
The ONLY file that knows the layout of the input spreadsheets.

Everything downstream consumes build/sources.pkl, so when a file's shape changes
this is the single place to fix.

Two kinds of thing are stored for every source table:

  'verbatim'  the block exactly as it appears in the file, values untouched,
              including Total columns, re-exports rows, primary-input rows and
              any 'n.a.' text. This is what gets written to the RAW_ tabs.
  'index'     row and column labels, so the MAP layer can locate the 115-code
              spine inside the verbatim block without anyone editing the block.

Rule: never drop, pad, re-order or coerce anything on the way in. See CLAUDE.md.

    python scripts/load_sources.py
"""
import pickle
import re
import sys
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
ABS_DIR = ROOT / 'data' / 'abs'
SUP_DIR = ROOT / 'data' / 'supplied'
OUT = ROOT / 'build' / 'sources.pkl'
CODE_RE = re.compile(r'^\d{3,4}$')


def norm_code(v):
    """IOIG code as 4-char text. ABS files store 0101 as the number 101 in places,
    which is exactly how leading-zero codes get lost. Always normalise on read."""
    if v is None:
        return None
    s = str(v).strip()
    if s.endswith('.0'):
        s = s[:-2]
    return s.zfill(4) if CODE_RE.match(s) else None


REGIONS = ['Aus', 'NSW', 'Vic', 'QLD', 'SA', 'WA', 'Tas', 'NT', 'ACT']

# ABS margin and tax tables -> the IOIG that earns the margin. Verified against Table 21.
MARGIN_TABLES = {
    '23': ('Wholesale', '3301'), '24': ('Retail', '3901'), '25': ('RestHotelClub', '4501'),
    '26': ('Road', '4601'), '27': ('Rail', '4701'), '28': ('Pipeline', '4801'),
    '29': ('Water', '4801'), '30': ('Air', '4901'), '31': ('PortHandling', '5201'),
    '32': ('MarineIns', '6301'), '33': ('Gas', '2701'), '34': ('Electricity', '2605'),
}
# 2023-24 control totals on the 115-code spine, $m. Used by check_sources.py.
MARGIN_CONTROL = {'Wholesale': 182778, 'Retail': 146374, 'RestHotelClub': 4790, 'Road': 44056,
                  'Rail': 7171, 'Pipeline': 2392, 'Water': 120, 'Air': 1151, 'PortHandling': 1057,
                  'MarineIns': 23, 'Gas': 3608, 'Electricity': 27890}
MARGIN_TOTAL_SPINE = 421410       # published 422,034 less the re-exports row
NET_TAX_TOTAL_SPINE = 168673


def grid(ws, max_row=400, max_col=200):
    """Read a sheet into a list of tuples, values exactly as stored."""
    return [r for r in ws.iter_rows(min_row=1, max_row=max_row, max_col=max_col, values_only=True)]


def find_code_column(rows, min_hits=100):
    hits = {}
    for row in rows:
        for ci, v in enumerate(row, 1):
            if norm_code(v):
                hits[ci] = hits.get(ci, 0) + 1
    if not hits:
        return None
    col = max(hits, key=hits.get)
    return col if hits[col] >= min_hits else col


def spine_index(rows, code_col):
    """Map IOIG code -> row number (1-based) for the first occurrence of each code."""
    out = {}
    for ri, row in enumerate(rows, 1):
        if len(row) >= code_col:
            c = norm_code(row[code_col - 1])
            if c:
                out.setdefault(c, ri)
    return out


# ---------------------------------------------------------------- ABS loaders
def load_abs_table(path, sheet=None):
    """Load an ABS data cube verbatim. Works for Tables 5, 8 and 23-35."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    name = sheet or (wb.sheetnames[1] if len(wb.sheetnames) > 1 else wb.sheetnames[0])
    rows = grid(wb[name])
    wb.close()
    code_col = find_code_column(rows)
    return {'verbatim': rows, 'sheet': name, 'code_col': code_col,
            'row_index': spine_index(rows, code_col) if code_col else {},
            'source': str(path.name)}


def load_all_abs():
    """Everything in data/abs/. Filenames are matched loosely so renames survive."""
    got = {}
    for p in sorted(ABS_DIR.glob('*.xls*')):
        stem = p.stem
        m = re.search(r'5209055001(\d{2})', stem)
        key = None
        if m:
            key = 'T' + str(int(m.group(1)))
        else:
            m2 = re.search(r'[Tt]able[ _-]?(\d{1,2})', stem)
            if m2:
                key = 'T' + m2.group(1)
        if key is None:
            print(f"  ? could not identify {p.name} - name it like 520905500123.xlsx or 'Table 23.xlsx'")
            continue
        got[key] = load_abs_table(p)
        print(f"  {key:5s} <- {p.name}  sheet '{got[key]['sheet']}'  "
              f"{len(got[key]['verbatim'])} rows, {len(got[key]['row_index'])} codes")
    return got


# ------------------------------------------------------- supplied data loaders
def load_supplied_flows():
    """
    Table 5 and Table 8 for all nine regions, as supplied.

    TODO ADAPT ME. Run scripts/inspect_inputs.py first, then fill this in.

    Return {('T5'|'T8', region): {'verbatim': rows, 'code_col': int,
                                  'row_index': {code: row}, 'source': str}}

    Watch for:
      - region blocks stacked in one sheet, keyed by a label column, versus one
        sheet or one file per region
      - primary-input rows named P3a/P3b/P3c/P3d and P4a/P4b in state blocks but
        P3/P4 nationally. Keep whatever is there; check_sources.py sums the
        P-rows dynamically
      - a 114-code spine with no 6700. Do NOT pad it here. The MAP layer bridges
    """
    out = {}
    files = sorted(SUP_DIR.glob('*.xls*'))
    if not files:
        print("  ! nothing in data/supplied/")
        return out
    print("  ! load_supplied_flows() is a stub - adapt it to your files, then rerun")
    return out


def load_supplied_multipliers():
    """
    The multiplier set for all nine regions, as supplied.

    TODO ADAPT ME.

    Return {'blocks': [measure names], 'effects': [11 effect names],
            'data': {'REGION|CODE': [values in block-major order]},
            'verbatim': rows, 'source': str}

    Watch for:
      - 'n.a.' text cells. KEEP THEM AS TEXT. The MAP layer converts, RAW does not
      - measure blocks spaced 12 columns apart with 11 effect columns and a spacer
      - the 11 effects, in order: Initial, First round, Simple, Industrial support,
        Production-induced, Type 1A, Type 1B, Total, Consumption-induced,
        Type 2A, Type 2B
    """
    print("  ! load_supplied_multipliers() is a stub - adapt it to your files, then rerun")
    return {}


def main():
    ROOT.joinpath('build').mkdir(exist_ok=True)
    print("ABS source tables:")
    abs_tables = load_all_abs()
    print("\nSupplied flow tables:")
    flows = load_supplied_flows()
    print("\nSupplied multipliers:")
    mult = load_supplied_multipliers()

    sources = {'abs': abs_tables, 'flows': flows, 'multipliers': mult,
               'regions': REGIONS, 'margin_tables': MARGIN_TABLES,
               'margin_control': MARGIN_CONTROL,
               'margin_total_spine': MARGIN_TOTAL_SPINE,
               'net_tax_total_spine': NET_TAX_TOTAL_SPINE}
    with open(OUT, 'wb') as f:
        pickle.dump(sources, f)
    print(f"\nwrote {OUT}")
    if not flows or not mult:
        print("Loaders still incomplete - check_sources.py will tell you what is missing.")


if __name__ == '__main__':
    main()
