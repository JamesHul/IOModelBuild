"""
Recalculate the generated workbook and scan every cell for error values.

    python scripts/recalc_model.py [workbook]

A build that has not been recalculated is not verified. LibreOffice cannot load
any xlsx in this container (it fails on a three-cell file), so this uses the
`formulas` engine instead. It evaluates the workbook from scratch - it does not
trust any cached value openpyxl wrote.

Reports every #REF!, #VALUE!, #DIV/0!, #N/A, #NAME?, #NUM! and #NULL!, then
prints the QA gates and the headline results so the numbers can be eyeballed.
"""
import re
import sys
import warnings
from collections import Counter
from pathlib import Path

warnings.filterwarnings('ignore')
import formulas                                     # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
DEFAULT = ROOT / 'output' / 'IO_Impact_Model_v0-3.xlsx'
ERRS = ('#REF!', '#VALUE!', '#DIV/0!', '#N/A', '#NAME?', '#NUM!', '#NULL!', '#ERROR!')
CELL = re.compile(r"^'?\[.*?\]([^']+)'?!([A-Z]+\d+)$", re.I)


def scan(path):
    print(f'loading {path.name} ...', flush=True)
    xl = formulas.ExcelModel().loads(str(path)).finish()
    print('calculating ...', flush=True)
    sol = xl.calculate()
    print(f'{len(sol):,} cells evaluated', flush=True)

    bad = Counter()
    examples = {}
    values = {}
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

    print('\n' + '=' * 68)
    if bad:
        print('ERROR CELLS')
        for (sheet, err), n in sorted(bad.items(), key=lambda kv: -kv[1]):
            print(f'  {sheet:20s} {err:9s} {n:6,d}   e.g. {examples[(sheet, err)]}')
    else:
        print('NO ERROR CELLS - every formula evaluated cleanly')
    print('=' * 68)

    def g(sheet, ref):
        return values.get((sheet.upper(), ref.upper()))

    print('\nQA_CHECKS')
    print(f"  STATUS: {g('QA_Checks','B4')}   ({g('QA_Checks','C4')})")
    for r in range(7, 40):
        name, res = g('QA_Checks', f'B{r}'), g('QA_Checks', f'C{r}')
        if name is None or str(name).strip() == '':
            continue
        print(f'    {str(res):8s} {str(name)[:62]}')

    print('\nOUT_SUMMARY  (year 1, then total)')
    for r in range(6, 40):
        meas, eff = g('OUT_Summary', f'A{r}'), g('OUT_Summary', f'B{r}')
        if not meas or str(meas).strip() == '':
            continue
        y1, tot = g('OUT_Summary', f'C{r}'), g('OUT_Summary', f'K{r}')
        try:
            print(f'    {str(meas)[:34]:34s} {str(eff)[:30]:30s} {float(y1):14,.1f} {float(tot):14,.1f}')
        except Exception:
            print(f'    {str(meas)[:34]:34s} {str(eff)[:30]:30s} {y1} {tot}')

    print('\nCONVERSION / VECTOR')
    print(f"  total spend, year 1      : {g('IN_Shock','F48')}")
    print(f"  direct domestic, year 1  : {g('CALC_Vector','C123')}")
    sys.exit(1 if bad else 0)


if __name__ == '__main__':
    main()
