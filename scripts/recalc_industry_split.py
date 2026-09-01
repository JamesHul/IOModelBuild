"""
Recalculate the industry-split workbook and scan every cell for error values.

    python scripts/recalc_industry_split.py [workbook]

A build that has not been recalculated is not verified. LibreOffice cannot load
any xlsx in this container, so this uses the `formulas` engine, which evaluates
from scratch and trusts no cached value openpyxl wrote.

Run it on the subset build - IOSPLIT_SUBSET=1 python scripts/build_industry_split.py -
which is structurally identical but small enough to finish.
"""
import re
import sys
import warnings
from collections import Counter
from pathlib import Path

warnings.filterwarnings('ignore')
import formulas                                     # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
DEFAULT = ROOT / 'build' / 'IO_Industry_Split_subset.xlsx'
ERRS = ('#REF!', '#VALUE!', '#DIV/0!', '#N/A', '#NAME?', '#NUM!', '#NULL!', '#ERROR!')
CELL = re.compile(r"^'?\[.*?\]([^']+)'?!([A-Z]+\d+)$", re.I)


def scan(path):
    print(f'loading {path.name} ...', flush=True)
    xl = formulas.ExcelModel().loads(str(path)).finish()
    print('calculating ...', flush=True)
    sol = xl.calculate()
    print(f'{len(sol):,} cells evaluated', flush=True)
    bad, examples, values = Counter(), {}, {}
    for k, v in sol.items():
        m = CELL.match(k)
        if not m:
            continue
        sheet, ref = m.group(1).strip("'").upper(), m.group(2).upper()
        try:
            val = v.value[0, 0]
        except Exception:
            continue
        values[(sheet, ref)] = val
        s = str(val).strip().upper()
        if s in ERRS:
            bad[(sheet, s)] += 1
            examples.setdefault((sheet, s), ref)
    return values, bad, examples


def main():
    path = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT
    if not path.exists():
        sys.exit(f'missing {path}')
    values, bad, examples = scan(path)

    print('\n' + '=' * 72)
    if bad:
        print('ERROR CELLS')
        for (sheet, err), n in sorted(bad.items(), key=lambda kv: -kv[1]):
            print(f'  {sheet:22s} {err:9s} {n:6,d}   e.g. {examples[(sheet, err)]}')
    else:
        print('NO ERROR CELLS - every formula evaluated cleanly')
    print('=' * 72)

    def g(sheet, ref):
        return values.get((sheet.upper(), ref.upper()))

    print('\nQA_CHECKS')
    print(f"  OVERALL: {g('QA_Checks', 'C4')}")
    for r in range(8, 30):
        name, res = g('QA_Checks', f'B{r}'), g('QA_Checks', f'C{r}')
        if name is None or str(name).strip() == '':
            continue
        print(f'    {str(res):6s} {str(name)[:66]}')

    print('\nQA_RECONCILE  (split total vs the supplied Simple multiplier)')
    worst_t = worst_d = 0.0
    n = 0
    for r in range(6, 400):
        meas = g('QA_Reconcile', f'A{r}')
        if meas is None or str(meas).strip() == '':
            continue
        n += 1
        for col, keep in (('F', 't'), ('I', 'd')):
            v = g('QA_Reconcile', f'{col}{r}')
            if isinstance(v, (int, float)):
                if keep == 't':
                    worst_t = max(worst_t, abs(float(v)))
                else:
                    worst_d = max(worst_d, abs(float(v)))
        if n <= 8:
            print(f"    {str(meas):9s} {str(g('QA_Reconcile', f'B{r}')):5s} "
                  f"yr{g('QA_Reconcile', f'C{r}')}  "
                  f"split={float(g('QA_Reconcile', f'D{r}') or 0):14,.4f}  "
                  f"mult={float(g('QA_Reconcile', f'E{r}') or 0):14,.4f}  "
                  f"diff={float(g('QA_Reconcile', f'F{r}') or 0):.2e}")
    print(f'    ... {n} rows.  worst TOTAL diff {worst_t:.3e}, '
          f'worst DIRECT diff {worst_d:.3e}')
    sys.exit(1 if bad else 0)


if __name__ == '__main__':
    main()
