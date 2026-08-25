"""
Independently compute the whole superannuation-spending study from
build/sources.pkl and the supplied spending file, and check every output.

    python scripts/check_super_shock.py

Written without reference to the workbook's formulas, so agreement between the
two is real evidence rather than a restatement.
"""
import numbers
import pickle
from collections import defaultdict
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
import os
SHOCK = Path(os.environ.get('IOMODEL_SHOCK',
             ROOT / 'output' / 'Super_shock_CORRECTED_by_state.xlsx'))
FF, FL, MF, ML = 128, 128, 119, 119          # Q1_HFCE columns
MARGINS = ['Wholesale', 'Retail', 'RestHotelClub', 'Road', 'Rail', 'Pipeline',
           'Water', 'Air', 'PortHandling', 'MarineIns', 'Gas', 'Electricity']
RMAP = {'National': 'Aus', 'NSW': 'NSW', 'VIC': 'Vic', 'QLD': 'QLD', 'SA': 'SA',
        'WA': 'WA', 'TAS': 'Tas', 'NT': 'NT', 'ACT': 'ACT'}
ORDER = ['NSW', 'Vic', 'QLD', 'WA', 'SA', 'Tas', 'ACT', 'NT']


def num(v):
    return float(v) if isinstance(v, numbers.Number) else 0.0


def rs(row, a, b):
    return sum(num(row[c - 1]) for c in range(a, b + 1) if c - 1 < len(row))


def main():
    src = pickle.load(open(ROOT / 'build' / 'sources.pkl', 'rb'))
    abs_t, flows, mult = src['abs'], src['flows'], src['multipliers']
    mt = src['margin_tables']
    tbl_of = {v[0]: k for k, v in mt.items()}
    earner = {v[0]: v[1] for v in mt.values()}
    blocks = {b['name']: b['first_col'] for b in mult['blocks']}
    eff = mult['effects']

    wb = openpyxl.load_workbook(SHOCK, data_only=True)
    spend = defaultdict(float)
    for r in wb['ABS spending long'].iter_rows(min_row=3, values_only=True):
        if not r or not r[0]:
            continue
        g, _t, code, _c, _bn, m = r[:6]
        if not isinstance(m, numbers.Number) or m == 0:
            continue
        spend[(RMAP[str(g).strip()], str(code).strip().zfill(4))] += float(m)
    wb.close()

    def mrow(margin, code):
        t = abs_t['T' + tbl_of[margin]] if margin in tbl_of else abs_t['T35']
        i = next((i for i, m in enumerate(t['meta']) if m['code'] == code), None)
        return t, i

    # national rates, computed once per code
    natcache = {}

    def nat_rates(code):
        if code in natcache:
            return natcache[code]
        t5, t8 = flows[('T5', 'Aus')], flows[('T8', 'Aus')]
        i5 = next((i for i, m in enumerate(t5['meta']) if m['code'] == code), None)
        i8 = next((i for i, m in enumerate(t8['meta']) if m['code'] == code), None)
        d = rs(t5['verbatim'][i5], FF, FL) if i5 is not None else 0.0
        im = (rs(t8['verbatim'][i8], FF, FL) - d) if i8 is not None else 0.0
        comp = {}
        t35, i35 = mrow('NetTaxes', code)
        comp['NetTaxes'] = rs(t35['verbatim'][i35], MF, ML) if i35 is not None else 0.0
        for mn in MARGINS:
            t, i = mrow(mn, code)
            comp[mn] = rs(t['verbatim'][i], MF, ML) if i is not None else 0.0
        pp = d + im + sum(comp.values())
        res = (pp, (d + im) / pp if pp else 0.0,
               {k: (v / pp if pp else 0.0) for k, v in comp.items()})
        natcache[code] = res
        return res

    def reg_split(region, code):
        t5, t8 = flows[('T5', region)], flows[('T8', region)]
        i5 = next((i for i, m in enumerate(t5['meta']) if m['code'] == code), None)
        i8 = next((i for i, m in enumerate(t8['meta']) if m['code'] == code), None)
        d = rs(t5['verbatim'][i5], FF, FL) if i5 is not None else 0.0
        im = (rs(t8['verbatim'][i8], FF, FL) - d) if i8 is not None else 0.0
        return d / (d + im) if (d + im) else 0.0

    vector = defaultdict(float)
    spend_by_reg = defaultdict(float)
    comp_by_reg = defaultdict(lambda: defaultdict(float))
    worst_share = 0.0
    for (reg, code), amt in spend.items():
        pp, basic, rates = nat_rates(code)
        split = reg_split(reg, code)
        dsh = basic * split
        ish = basic * (1 - split)
        tot = dsh + ish + sum(rates.values())
        worst_share = max(worst_share, abs(tot - 1))
        spend_by_reg[reg] += amt
        vector[(reg, code)] += amt * dsh
        comp_by_reg[reg]['domestic'] += amt * dsh
        comp_by_reg[reg]['imports'] += amt * ish
        comp_by_reg[reg]['tax'] += amt * rates['NetTaxes']
        for mn in MARGINS:
            v = amt * rates[mn]
            comp_by_reg[reg]['margins'] += v
            if v:
                vector[(reg, earner[mn])] += v

    print(f'CHECK 1  shares sum to 1 on every line: max deviation {worst_share:.2e}')
    print(f'CHECK 2  lines loaded: {len(spend)}   codes: '
          f'{len({c for _, c in spend})}   regions: {len({r for r, _ in spend})}')
    tot_states = sum(spend_by_reg[r] for r in ORDER)
    print(f'CHECK 3  spend: states ${tot_states:,.1f}m vs Aus ${spend_by_reg["Aus"]:,.1f}m '
          f'(diff {tot_states - spend_by_reg["Aus"]:+,.4f})')
    print()
    print(f'{"Region":7s} {"spend $m":>11s} {"domestic":>11s} {"imports":>10s} '
          f'{"tax":>9s} {"margins":>10s} {"recon":>9s} {"convert":>8s}')
    for reg in ORDER + ['Aus']:
        c = comp_by_reg[reg]
        recon = c['domestic'] + c['imports'] + c['tax'] + c['margins']
        dv = sum(v for (r, _), v in vector.items() if r == reg)
        print(f'{reg:7s} {spend_by_reg[reg]:11,.1f} {c["domestic"]:11,.1f} '
              f'{c["imports"]:10,.1f} {c["tax"]:9,.1f} {c["margins"]:10,.1f} '
              f'{recon - spend_by_reg[reg]:+9,.4f} {dv / spend_by_reg[reg]:8.2%}')
    print('   recon = components less spend; must be 0. convert = direct domestic / spend')

    print('\nIMPACTS')
    hdr = f'{"Region":7s} {"spend":>10s} ' + ' '.join(f'{x:>12s}' for x in
          ['Output dir', 'Output TOT', 'GVA TOT', 'Wages TOT', 'FTE TOT'])
    print(hdr)
    agg = defaultdict(float)
    for reg in ORDER + ['Aus']:
        md = mult['regions'][reg]
        rows = {m['code']: i for i, m in enumerate(md['meta']) if m['row_type'] == 'Industry'}
        res = {}
        for bn, lab in [('Output multipliers', 'Output'), ('Value added multipliers', 'GVA'),
                        ('Income multipliers', 'Wages'), ('Employed multipliers', 'FTE')]:
            for en in ('Initial effect', 'Total multiplier'):
                j = blocks[bn] - 1 + eff.index(en)
                s = sum(v * num(md['verbatim'][rows[c if c in rows else '6701']][j])
                        for (r, c), v in vector.items() if r == reg)
                res[(lab, en)] = s
                if reg != 'Aus':
                    agg[(lab, en)] += s
        print(f'{reg:7s} {spend_by_reg[reg]:10,.0f} '
              f'{res[("Output", "Initial effect")]:12,.0f} {res[("Output", "Total multiplier")]:12,.0f} '
              f'{res[("GVA", "Total multiplier")]:12,.0f} {res[("Wages", "Total multiplier")]:12,.0f} '
              f'{res[("FTE", "Total multiplier")]:12,.0f}')
    print(f'{"SUM(st)":7s} {tot_states:10,.0f} '
          f'{agg[("Output", "Initial effect")]:12,.0f} {agg[("Output", "Total multiplier")]:12,.0f} '
          f'{agg[("GVA", "Total multiplier")]:12,.0f} {agg[("Wages", "Total multiplier")]:12,.0f} '
          f'{agg[("FTE", "Total multiplier")]:12,.0f}')
    print('\nCHECK 4  sum of states must be BELOW the Aus run (states leak more):')
    for lab in ['Output', 'GVA', 'Wages', 'FTE']:
        a = agg[(lab, 'Total multiplier')]
        n = sum(v * num(mult['regions']['Aus']['verbatim'][
            {m['code']: i for i, m in enumerate(mult['regions']['Aus']['meta'])
             if m['row_type'] == 'Industry'}[c if c != '6700' else '6701']][
            blocks[{'Output': 'Output multipliers', 'GVA': 'Value added multipliers',
                    'Wages': 'Income multipliers', 'FTE': 'Employed multipliers'}[lab]]
            - 1 + eff.index('Total multiplier')])
            for (r, c), v in vector.items() if r == 'Aus')
        flag = 'ok' if a < n else 'UNEXPECTED'
        print(f'   {lab:7s} sum(states) {a:14,.0f}   Aus {n:14,.0f}   '
              f'ratio {a / n:5.2%}  {flag}')


if __name__ == '__main__':
    main()
