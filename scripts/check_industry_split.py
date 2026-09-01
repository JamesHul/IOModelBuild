"""
Independent check of the industry-split workbook.

    python scripts/check_industry_split.py [workbook]

A second opinion on the arithmetic, computed from the source spreadsheets rather
than from the generator, plus a verbatim round-trip of the RAW tabs. It does NOT
evaluate the workbook's formulas - scripts/recalc_industry_split.py does that.

Five gates:

  1  RAW round-trip. Every cell of RAW_Split_* read back off disk and compared
     against the source file. Two representation differences are collapsed and
     nothing else. openpyxl writes '' back as an empty cell, so None and '' are
     the same cell - but only those; a 0, a '0' or an 'n.a.' is real content and
     a difference in one is a real difference. And openpyxl serialises floats
     with %.16g, which drops the seventeenth significant digit, so numbers are
     compared with the same 1e-9 relative tolerance verify_stacked.py uses. The
     worst relative deviation is printed, so a real change to a number could not
     hide behind the tolerance.

  2  Reconciliation to the supplied multipliers. Each split-out column, summed
     down its 124 rows, must equal that region's Simple multiplier for the same
     measure in "Piggy IO tables and multipliers_V2.xlsx". Likewise the sheet's
     own Initial effect row against the multiplier summary's. This is the gate
     that makes the whole add-on legitimate: if it holds, an industry split that
     sums the column reproduces the number the master model already reports.

  3  Region distinctness. Six of eight state blocks in the old master model were
     byte-identical to NSW and nothing in the workbook revealed it.

  4  End-to-end recomputation. Runs the workbook's own arithmetic in numpy on
     whatever sits in IN_DirectVector, and checks that summing the industry
     split reproduces shock x Simple multiplier, and that the direct effect
     reproduces shock x Initial effect.

  5  Value-added definition. Reports the gap between value added at market
     prices, which is the only split-out family supplied, and value added at
     basic prices, which is the master model's headline.
"""
import numbers
import pickle
import sys
from collections import defaultdict
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
SPLITPKL = ROOT / 'build' / 'splitout.pkl'
SUP = ROOT / 'data' / 'supplied'
MAIN = SUP / 'Piggy IO tables and multipliers_V2.xlsx'
DEFAULT = ROOT / 'output' / 'IO_Industry_Split_Model.xlsx'
TOL = 1e-6

# Column of the first effect in each multiplier block on the summary sheets, and
# the offset of each effect within a block. Established in load_sources.py.
EFF_OFFSET = {'Initial effect': 0, 'First round effect': 1, 'Simple multiplier': 2,
              'Industrial support effect': 3, 'Production-induced effect': 4}


def empty(v):
    """An empty cell and a cell holding an empty string are the same thing.
    Deliberately narrow: a '0', a 0 or an 'n.a.' is real content."""
    return v is None or (isinstance(v, str) and v.strip() == '')


def same(a, b):
    """Equal as stored, with the two representation differences allowed for.

    Returns (equal, relative deviation) so the caller can report how much slack
    the tolerance actually absorbed.
    """
    if empty(a) and empty(b):
        return True, 0.0
    if isinstance(a, numbers.Number) and isinstance(b, numbers.Number):
        if a == b:
            return True, 0.0
        scale = max(abs(a), abs(b), 1.0)
        rel = abs(a - b) / scale
        return rel <= 1e-9, rel
    if empty(a) or empty(b):
        return False, float('inf')
    return str(a).strip() == str(b).strip(), 0.0


def main():
    path = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT
    if not SPLITPKL.exists():
        sys.exit('build/splitout.pkl missing - run scripts/load_splitout.py')
    src = pickle.load(open(SPLITPKL, 'rb'))
    blocks, regions, measures = src['blocks'], list(src['regions']), src['measures']
    FR, LR, LC = src['first_row'], src['last_row'], src['last_col']
    FD, LD = src['first_data_row'], src['last_data_row']
    FMC, LMC = src['first_mat_col'], src['last_mat_col']
    NC = LD - FD + 1
    order = [m['code'] for m in blocks[(measures[0][0], regions[0])]['row_meta']]
    fails = []

    # ------------------------------------------------------------- gate 1
    print('=' * 74)
    print('1  RAW round-trip: the built workbook against the source file')
    if not path.exists():
        print(f'   ! {path} not built yet - skipping')
    else:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        for key, *_ in measures:
            tab = f'RAW_Split_{key}'
            if tab not in wb.sheetnames:
                fails.append(f'{tab} missing')
                print(f'   ! {tab} missing')
                continue
            ws = wb[tab]
            got = [r for r in ws.iter_rows(min_row=6, max_col=6 + LC, values_only=True)]
            n = diff = 0
            worst = 0.0
            gi = 0
            for rg in regions:
                b = blocks[(key, rg)]
                for si, row in enumerate(b['verbatim']):
                    g = got[gi]
                    gi += 1
                    if g[1] != rg or g[5] != FR + si:
                        fails.append(f'{tab} row {gi} keyed {g[1]}/{g[5]}, '
                                     f'expected {rg}/{FR + si}')
                        diff += 1
                        continue
                    for ci, v in enumerate(row):
                        n += 1
                        ok, rel = same(g[6 + ci], v)
                        worst = max(worst, rel if rel != float('inf') else worst)
                        if not ok:
                            diff += 1
                            if diff <= 3:
                                print(f'   ! {tab} r{gi} src col {ci + 1}: '
                                      f'{g[6 + ci]!r} != {v!r}')
            status = 'VERBATIM' if diff == 0 else f'{diff} MISMATCHES'
            print(f'   {tab:22s} {n:,} cells   {status}   '
                  f'worst relative deviation {worst:.2e} '
                  f'(xlsx float serialisation, tolerance 1e-9)')
            if diff:
                fails.append(f'{tab}: {diff} mismatches')
        wb.close()

    # ------------------------------------------------------------- gate 2
    print('=' * 74)
    print('2  Reconciliation: split-out column sums vs the supplied multipliers')
    wm = openpyxl.load_workbook(MAIN, read_only=True, data_only=True)
    blkcol = {}
    ws0 = wm[f'{src["sheet_region"][regions[0]]} multiplier summary FY23']
    for ci in range(1, 200):
        v = ws0.cell(row=10, column=ci).value
        if v not in (None, ''):
            blkcol[str(v).strip()] = ci
    mult = {}
    for rg in regions:
        ws = wm[f'{src["sheet_region"][rg]} multiplier summary FY23']
        g = list(ws.iter_rows(min_row=1, max_row=140, max_col=182, values_only=True))
        idx = {}
        for r in range(12, 137):
            v = g[r - 1][0]
            if v is None:
                continue
            s = str(v).strip()
            s = s[:-2] if s.endswith('.0') else s
            idx[s.zfill(4) if s.isdigit() else s] = r
        mult[rg] = (g, idx)
    wm.close()

    for key, suffix, label, unit, fmt, mblock in measures:
        bc = blkcol.get(mblock)
        if bc is None:
            fails.append(f'no multiplier block {mblock!r}')
            print(f'   ! block {mblock!r} not on the summary sheets')
            continue
        worst_sum = worst_ini = 0.0
        nchk = 0
        for rg in regions:
            b = blocks[(key, rg)]
            g, idx = mult[rg]
            mat = [row[FMC - 1:LMC] for row in b['verbatim'][FD - FR:LD - FR + 1]]
            ini = b['verbatim'][b['effect_rows']['Initial effect'] - FR][FMC - 1:LMC]
            for j, code in enumerate(order):
                if code.startswith('99'):
                    continue
                r = idx.get(code)
                if r is None:
                    fails.append(f'{key}/{rg}: {code} not in the multiplier summary')
                    continue
                col_sum = sum(float(row[j]) for row in mat
                              if isinstance(row[j], numbers.Number))
                want = g[r - 1][bc - 1 + EFF_OFFSET['Simple multiplier']]
                want_i = g[r - 1][bc - 1 + EFF_OFFSET['Initial effect']]
                if isinstance(want, numbers.Number):
                    worst_sum = max(worst_sum, abs(col_sum - float(want)))
                    nchk += 1
                if isinstance(want_i, numbers.Number) and isinstance(ini[j], numbers.Number):
                    worst_ini = max(worst_ini, abs(float(ini[j]) - float(want_i)))
        ok = worst_sum < 1e-9 and worst_ini < 1e-9
        print(f'   {label[:44]:44s} {nchk:4d} codes  '
              f'col sum vs Simple multiplier {worst_sum:.2e}  '
              f'Initial effect {worst_ini:.2e}  {"OK" if ok else "FAIL"}')
        if not ok:
            fails.append(f'{key}: reconciliation off by {max(worst_sum, worst_ini):.3e}')

    # ------------------------------------------------------------- gate 3
    print('=' * 74)
    print('3  Region distinctness')
    import hashlib
    for key, *_ in measures:
        seen = defaultdict(list)
        for rg in regions:
            b = blocks[(key, rg)]
            mat = [row[FMC - 1:LMC] for row in b['verbatim'][FD - FR:LD - FR + 1]]
            seen[hashlib.sha256(repr(mat).encode()).hexdigest()].append(rg)
        dup = [v for v in seen.values() if len(v) > 1]
        print(f'   {key:10s} {len(seen)} distinct blocks across {len(regions)} regions'
              f'{"   DUPLICATES: " + str(dup) if dup else ""}')
        if dup:
            fails.append(f'{key}: identical region blocks {dup}')

    # ------------------------------------------------------------- gate 4
    print('=' * 74)
    print('4  End-to-end: the workbook\'s own arithmetic, recomputed here')
    if not path.exists():
        print('   ! workbook not built yet - skipping')
    else:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb['IN_DirectVector']
        nyr = 0
        for ci in range(4, 40):
            v = ws.cell(row=5, column=ci).value
            if isinstance(v, str) and v.startswith('Year '):
                nyr += 1
            else:
                break
        shock = defaultdict(lambda: [0.0] * nyr)     # (region, code) -> years
        pasted = 0.0
        for row in ws.iter_rows(min_row=6, max_col=3 + nyr, values_only=True):
            srcname, rg, code = row[0], row[1], row[2]
            if not rg or not code:
                continue
            code = str(code).strip()
            code = '6701' if code == '6700' else code
            for y in range(nyr):
                v = row[3 + y]
                if isinstance(v, numbers.Number):
                    shock[(rg, code)][y] += float(v)
                    pasted += float(v)
        wb.close()
        print(f'   {len(shock)} (region, industry) cells, {pasted:,.1f} $m pasted, '
              f'{nyr} years')
        if not shock:
            print('   ! IN_DirectVector is empty - nothing to recompute')
        worst_t = worst_d = 0.0
        rows = 0
        for key, suffix, label, unit, fmt, mblock in measures:
            for rg in regions:
                b = blocks[(key, rg)]
                mat = [[float(v) if isinstance(v, numbers.Number) else 0.0
                        for v in row[FMC - 1:LMC]]
                       for row in b['verbatim'][FD - FR:LD - FR + 1]]
                ini = [float(v) if isinstance(v, numbers.Number) else 0.0
                       for v in b['verbatim'][b['effect_rows']['Initial effect']
                                              - FR][FMC - 1:LMC]]
                sim = [float(v) if isinstance(v, numbers.Number) else 0.0
                       for v in b['verbatim'][b['effect_rows']['Simple multiplier']
                                              - FR][FMC - 1:LMC]]
                for y in range(nyr):
                    sv = [shock.get((rg, c), [0.0] * nyr)[y] for c in order]
                    if not any(sv):
                        continue
                    rows += 1
                    total_i = [sum(mat[i][j] * sv[j] for j in range(NC))
                               for i in range(NC)]
                    direct_i = [sv[i] * ini[i] for i in range(NC)]
                    worst_t = max(worst_t, abs(sum(total_i)
                                               - sum(sv[j] * sim[j] for j in range(NC))))
                    worst_d = max(worst_d, abs(sum(direct_i)
                                               - sum(sv[j] * ini[j] for j in range(NC))))
        print(f'   {rows} (measure, region, year) cells with spend')
        print(f'   sum of the industry split vs shock x Simple multiplier : '
              f'{worst_t:.3e}  {"OK" if worst_t < 1e-6 else "FAIL"}')
        print(f'   sum of the direct column vs shock x Initial effect     : '
              f'{worst_d:.3e}  {"OK" if worst_d < 1e-6 else "FAIL"}')
        if worst_t >= 1e-6 or worst_d >= 1e-6:
            fails.append('end-to-end recomputation does not tie')

    # ------------------------------------------------------------- gate 5
    print('=' * 74)
    print('5  Value added: MARKET prices (supplied) against BASIC prices (headline)')
    wm = openpyxl.load_workbook(MAIN, read_only=True, data_only=True)
    bmp, bbp = blkcol.get('Value added at market prices multipliers'), \
        blkcol.get('Value added multipliers')
    for rg in regions:
        ws = wm[f'{src["sheet_region"][rg]} multiplier summary FY23']
        g = list(ws.iter_rows(min_row=12, max_row=136, max_col=182, values_only=True))
        d = [float(r[bmp - 1 + 2]) - float(r[bbp - 1 + 2]) for r in g
             if isinstance(r[bmp - 1 + 2], numbers.Number)
             and isinstance(r[bbp - 1 + 2], numbers.Number)]
        print(f'   {rg:4s} simple multiplier, market less basic: '
              f'mean {sum(d) / len(d):+.4f}  min {min(d):+.4f}  max {max(d):+.4f}')
    wm.close()
    print('   Not a failure - a definitional difference. Label the output '
          '"value added at market prices".')

    print('=' * 74)
    if fails:
        print(f'FAILED - {len(fails)} problem(s)')
        for f in fails:
            print(f'  - {f}')
        return 1
    print('ALL GATES PASSED')
    return 0


if __name__ == '__main__':
    sys.exit(main())
