"""
Independently reproduce the model's arithmetic straight from build/sources.pkl.

    python scripts/check_model_numbers.py

This is a second opinion, deliberately written without looking at the workbook's
formulas: it walks the same economics by a different route, so if the workbook
and this script agree, both are probably right, and if they disagree one of them
has a bug worth finding.

Mirrors the built-in example on IN_Shock.
"""
import numbers
import pickle
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / 'build' / 'sources.pkl'

GROUPS = {          # (flow first, flow last, margin first, margin last), 1-based
    'Intermediate': (3, 126, 3, 117),
    'Q1_HFCE': (128, 128, 119, 119),
    'Q2_GGFCE': (129, 129, 120, 120),
    'Q3_GFCF_Priv': (130, 130, 121, 121),
    'Q4_GFCF_PubCorp': (131, 131, 122, 122),
    'Q5_GFCF_GG': (132, 132, 123, 123),
    'Q6_Inventories': (133, 133, 124, 124),
    'Q7_Exports': (134, 134, 125, 125),
}
MARGINS = ['Wholesale', 'Retail', 'RestHotelClub', 'Road', 'Rail', 'Pipeline',
           'Water', 'Air', 'PortHandling', 'MarineIns', 'Gas', 'Electricity']
EXAMPLE = [('3101', 'Q5_GFCF_GG', 100.0), ('1101', 'Q1_HFCE', 20.0),
           ('6901', 'Intermediate', 15.0)]
REGION = 'Aus'


def num(v):
    return float(v) if isinstance(v, numbers.Number) else 0.0


def rowsum(row, first, last):
    return sum(num(row[c - 1]) for c in range(first, last + 1) if c - 1 < len(row))


def main():
    src = pickle.load(open(SRC, 'rb'))
    abs_t, flows, mult = src['abs'], src['flows'], src['multipliers']
    mt = src['margin_tables']
    name_of = {v[0]: k for k, v in mt.items()}
    earner = {v[0]: v[1] for v in mt.values()}

    spine = [(m['code'], m['label']) for m in abs_t['T23']['meta']
             if m['row_type'] == 'Product']
    mult_rows = {m['code']: i for i, m in enumerate(mult['regions'][REGION]['meta'])
                 if m['row_type'] == 'Industry'}
    t5 = flows[('T5', REGION)]
    t8 = flows[('T8', REGION)]
    t5row = {m['code']: i for i, m in enumerate(t5['meta']) if m['row_type'] == 'Industry'}
    t8row = {m['code']: i for i, m in enumerate(t8['meta']) if m['row_type'] == 'Industry'}

    def margin_row(margin, code):
        t = abs_t['T' + name_of[margin]] if margin in name_of else abs_t['T35']
        for i, m in enumerate(t['meta']):
            if m['code'] == code:
                return t, i
        return t, None

    print(f'Region {REGION}. Independent recomputation of the built-in example.\n')
    vector = {}
    margin_tot = {m: 0.0 for m in MARGINS}
    total_spend = 0.0
    for code, grp, amt in EXAMPLE:
        ff, fl, mf, ml = GROUPS[grp]
        use = code if code in t5row else '6701'
        dom = rowsum(t5['verbatim'][t5row[use]], ff, fl) if use in t5row else 0.0
        imp = (rowsum(t8['verbatim'][t8row[use]], ff, fl) - dom) if use in t8row else 0.0
        comp = {'Domestic': dom, 'Imports': imp}
        t35, i35 = margin_row('NetTaxes', code)
        comp['NetTaxes'] = rowsum(t35['verbatim'][i35], mf, ml) if i35 is not None else 0.0
        for mn in MARGINS:
            t, i = margin_row(mn, code)
            comp[mn] = rowsum(t['verbatim'][i], mf, ml) if i is not None else 0.0
        pp = sum(comp.values())
        total_spend += amt
        print(f'  {code} in {grp}: PP cell ${pp:,.1f}m')
        for k in ('Domestic', 'Imports', 'NetTaxes'):
            print(f'      {k:10s} {comp[k] / pp if pp else 0:7.2%}  -> ${amt * comp[k] / pp if pp else 0:9,.3f}m')
        mshown = {k: comp[k] for k in MARGINS if comp[k]}
        if mshown:
            print('      margins    ' + ', '.join(
                f'{k} {comp[k] / pp:.2%}' for k in mshown))
        vector[use] = vector.get(use, 0.0) + (amt * comp['Domestic'] / pp if pp else 0)
        for mn in MARGINS:
            margin_tot[mn] += amt * comp[mn] / pp if pp else 0

    for mn, v in margin_tot.items():
        if v:
            e = earner[mn]
            vector[e] = vector.get(e, 0.0) + v
    direct = sum(vector.values())
    print(f'\n  total spend            ${total_spend:,.1f}m')
    print(f'  direct domestic vector ${direct:,.3f}m')
    print(f'  conversion ratio        {direct / total_spend:.2%}')

    # impacts
    blocks = {b['name']: b['first_col'] for b in mult['blocks']}
    eff = mult['effects']
    md = mult['regions'][REGION]
    print('\n  Impacts (year 1):')
    for bn, label in [('Output multipliers', 'Output $m'),
                      ('Income multipliers', 'Wages $m'),
                      ('Value added multipliers', 'Value added $m'),
                      ('Employed multipliers', 'Employment FTE')]:
        out = {}
        for en in ('Initial effect', 'Production-induced effect',
                   'Consumption-induced effect', 'Total multiplier'):
            j = blocks[bn] - 1 + eff.index(en)
            tot = 0.0
            for code, v in vector.items():
                mr = mult_rows.get(code if code in mult_rows else '6701')
                if mr is None:
                    continue
                tot += v * num(md['verbatim'][mr][j])
            out[en] = tot
        print(f'    {label:18s} direct {out["Initial effect"]:12,.2f}   '
              f'indirect {out["Production-induced effect"]:12,.2f}   '
              f'induced {out["Consumption-induced effect"]:12,.2f}   '
              f'TOTAL {out["Total multiplier"]:12,.2f}')


if __name__ == '__main__':
    main()
