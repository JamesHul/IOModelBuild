"""
Independently reproduce the multi-region model's arithmetic from sources.pkl.

    python scripts/check_multiregion_numbers.py

Second opinion on the per-line-region design, written without reading the
workbook's formulas. The built-in example shocks the SAME product and column
group (3101 into Q5_GFCF_GG) in both NSW and Aus, so if the region is really
driving the lookups the two lines must strip and multiply differently.
"""
import numbers
import pickle
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
GROUPS = {'Intermediate': (3, 126, 3, 117), 'Q1_HFCE': (128, 128, 119, 119),
          'Q2_GGFCE': (129, 129, 120, 120), 'Q3_GFCF_Priv': (130, 130, 121, 121),
          'Q4_GFCF_PubCorp': (131, 131, 122, 122), 'Q5_GFCF_GG': (132, 132, 123, 123),
          'Q6_Inventories': (133, 133, 124, 124), 'Q7_Exports': (134, 134, 125, 125)}
MARGINS = ['Wholesale', 'Retail', 'RestHotelClub', 'Road', 'Rail', 'Pipeline',
           'Water', 'Air', 'PortHandling', 'MarineIns', 'Gas', 'Electricity']
EXAMPLE = [('NSW', '3101', 'Q5_GFCF_GG', 100.0), ('Vic', '1101', 'Q1_HFCE', 20.0),
           ('QLD', '6901', 'Intermediate', 15.0), ('Aus', '3101', 'Q5_GFCF_GG', 50.0)]


def num(v):
    return float(v) if isinstance(v, numbers.Number) else 0.0


def rowsum(row, a, b):
    return sum(num(row[c - 1]) for c in range(a, b + 1) if c - 1 < len(row))


def main():
    src = pickle.load(open(ROOT / 'build' / 'sources.pkl', 'rb'))
    abs_t, flows, mult = src['abs'], src['flows'], src['multipliers']
    mt = src['margin_tables']
    tbl_of = {v[0]: k for k, v in mt.items()}
    earner = {v[0]: v[1] for v in mt.values()}
    mcodes = {m['code'] for m in mult['regions']['Aus']['meta'] if m['row_type'] == 'Industry'}
    blocks = {b['name']: b['first_col'] for b in mult['blocks']}
    eff = mult['effects']

    def mrow(margin, code):
        t = abs_t['T' + tbl_of[margin]] if margin in tbl_of else abs_t['T35']
        for i, m in enumerate(t['meta']):
            if m['code'] == code:
                return t, i
        return t, None

    vector = {}          # (region, code) -> $m
    per_region_spend = {}
    print('Per-line strip, each line using ITS OWN region:\n')
    for reg, code, grp, amt in EXAMPLE:
        ff, fl, mf, ml = GROUPS[grp]
        use = code if code in mcodes else '6701'
        t5, t8 = flows[('T5', reg)], flows[('T8', reg)]
        i5 = next((i for i, m in enumerate(t5['meta']) if m['code'] == use), None)
        i8 = next((i for i, m in enumerate(t8['meta']) if m['code'] == use), None)
        dom = rowsum(t5['verbatim'][i5], ff, fl) if i5 is not None else 0.0
        imp = (rowsum(t8['verbatim'][i8], ff, fl) - dom) if i8 is not None else 0.0
        comp = {'Domestic': dom, 'Imports': imp}
        t35, i35 = mrow('NetTaxes', code)
        comp['NetTaxes'] = rowsum(t35['verbatim'][i35], mf, ml) if i35 is not None else 0.0
        for mn in MARGINS:
            t, i = mrow(mn, code)
            comp[mn] = rowsum(t['verbatim'][i], mf, ml) if i is not None else 0.0
        pp = sum(comp.values())
        dshare = comp['Domestic'] / pp if pp else 0
        print(f'  {reg:4s} {code} {grp:15s} ${amt:6,.1f}m   PP cell ${pp:12,.1f}m   '
              f'domestic {dshare:6.2%} -> ${amt * dshare:8,.3f}m')
        vector[(reg, use)] = vector.get((reg, use), 0.0) + amt * dshare
        per_region_spend[reg] = per_region_spend.get(reg, 0.0) + amt
        for mn in MARGINS:
            v = amt * comp[mn] / pp if pp else 0
            if v:
                k = (reg, earner[mn])
                vector[k] = vector.get(k, 0.0) + v

    print('\nDirect domestic vector, by (region, industry):')
    for (reg, code), v in sorted(vector.items()):
        if abs(v) > 1e-9:
            print(f'    {reg:4s} {code}  ${v:9,.3f}m')

    print('\nImpacts by region (own-region multipliers):')
    tot = {}
    for reg in ['NSW', 'Vic', 'QLD', 'Aus']:
        md = mult['regions'][reg]
        rows = {m['code']: i for i, m in enumerate(md['meta']) if m['row_type'] == 'Industry'}
        out = {}
        for bn, lab in [('Output multipliers', 'Output $m'),
                        ('Value added multipliers', 'Value added $m'),
                        ('Employed multipliers', 'Employment FTE')]:
            for en in ('Initial effect', 'Total multiplier'):
                j = blocks[bn] - 1 + eff.index(en)
                s = 0.0
                for (r2, code), v in vector.items():
                    if r2 != reg:
                        continue
                    i = rows.get(code if code in rows else '6701')
                    if i is not None:
                        s += v * num(md['verbatim'][i][j])
                out[(lab, en)] = s
                tot[(lab, en)] = tot.get((lab, en), 0.0) + s
        sp = per_region_spend.get(reg, 0)
        dv = sum(v for (r2, _), v in vector.items() if r2 == reg)
        print(f'  {reg:4s} spend ${sp:6,.1f}m  direct domestic ${dv:8,.3f}m  '
              f'conversion {dv / sp if sp else 0:6.2%}')
        for lab in ('Output $m', 'Value added $m', 'Employment FTE'):
            print(f'         {lab:16s} direct {out[(lab, "Initial effect")]:10,.2f}   '
                  f'TOTAL {out[(lab, "Total multiplier")]:10,.2f}')
    print('\n  ALL REGIONS (sum of the above - NOT the national impact):')
    for lab in ('Output $m', 'Value added $m', 'Employment FTE'):
        print(f'         {lab:16s} direct {tot[(lab, "Initial effect")]:10,.2f}   '
              f'TOTAL {tot[(lab, "Total multiplier")]:10,.2f}')

    # the decisive check: same product+group, different region
    print('\nDoes region actually change the answer? 3101 into Q5_GFCF_GG:')
    for reg in ('NSW', 'Aus'):
        t5, t8 = flows[('T5', reg)], flows[('T8', reg)]
        i5 = next(i for i, m in enumerate(t5['meta']) if m['code'] == '3101')
        i8 = next(i for i, m in enumerate(t8['meta']) if m['code'] == '3101')
        d = rowsum(t5['verbatim'][i5], 132, 132)
        im = rowsum(t8['verbatim'][i8], 132, 132) - d
        md = mult['regions'][reg]
        rows = {m['code']: i for i, m in enumerate(md['meta']) if m['row_type'] == 'Industry'}
        j = blocks['Output multipliers'] - 1 + eff.index('Total multiplier')
        print(f'    {reg:4s} T5 cell ${d:12,.2f}m   imports ${im:10,.2f}m   '
              f'output Type II multiplier {num(md["verbatim"][rows["3101"]][j]):.4f}')


if __name__ == '__main__':
    main()
