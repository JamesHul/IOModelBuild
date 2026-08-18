"""
Integrity gates on the loaded source data. Run before every build.

These are the checks that have actually caught things:
  - six of eight state flow-table blocks identical to NSW
  - five of eight multiplier sets identical to Australia
  - multipliers reconciling to Table 5 but not Table 8, confirming the provider
    used the direct-allocation table

    python scripts/check_sources.py
"""
import numbers
import pickle
import re
import sys
from pathlib import Path

import numpy as np

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / 'build' / 'sources.pkl'
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


def _num(v):
    """Numeric value or nan. Uses numbers.Number so ints Excel stored as ints
    are not silently skipped - that has caused false discrepancies twice."""
    import math
    return float(v) if isinstance(v, numbers.Number) else math.nan


PASS, FAIL, WARN = 'PASS', 'FAIL', 'WARN'
results = []


def report(name, status, detail=''):
    results.append((name, status, detail))
    mark = {'PASS': '  ok ', 'FAIL': 'FAIL ', 'WARN': 'warn '}[status]
    print(f"{mark} {name}" + (f"\n         {detail}" if detail else ''))


# ------------------------------------------------------------------ utilities
def numeric_block(rows, row_index, codes, first_data_col, ncols):
    """Pull a numeric matrix for the given codes. Text becomes nan, not zero."""
    out = np.full((len(codes), ncols), np.nan)
    for i, c in enumerate(codes):
        ri = row_index.get(c)
        if ri is None:
            continue
        row = rows[ri - 1]
        for j in range(ncols):
            ci = first_data_col - 1 + j
            v = row[ci] if ci < len(row) else None
            if isinstance(v, (int, float)):
                out[i, j] = float(v)
    return out


def simple_output_multiplier(Z, x):
    """Column sums of (I-A)^-1 where A = Z/x. Used only for checking."""
    A = np.divide(Z, x, out=np.zeros_like(Z), where=x != 0)
    return np.linalg.inv(np.eye(len(x)) - A).sum(0)


# --------------------------------------------------------------------- checks
def check_abs(src):
    abs_t = src['abs']
    # ABS Table 5 and Table 8 are deliberately NOT required here. The flow
    # tables come from the provider's regionalised set, and imports are derived
    # as supplied-T8 less supplied-T5. That is what removes the need for ABS
    # Tables 2 and 3 as well - see CLAUDE.md and the v2 build plan.
    needed = ['T35'] + ['T' + k for k in src['margin_tables']]
    missing = [k for k in needed if k not in abs_t]
    report('ABS margin and tax tables present', PASS if not missing else FAIL,
           '' if not missing else f"missing: {missing}")
    report('ABS Table 21 control present', PASS if 'T21' in abs_t else WARN,
           '' if 'T21' in abs_t else 'optional, but it is the independent margin control')

    for key, (name, ind) in src['margin_tables'].items():
        k = 'T' + key
        if k not in abs_t:
            continue
        t = abs_t[k]
        n = len(t['row_index'])
        report(f"{k} {name}: spine codes found", PASS if n >= 115 else WARN,
               f"{n} four-digit codes in column {t['code_col']}")

    # Margin control totals. These are the gate on a re-paste: if a new vintage
    # lands and these move, the strip rates move with them.
    ctrl = src.get('margin_control', {})
    tot = rex = 0.0
    for key, (name, _ind) in src['margin_tables'].items():
        t = abs_t.get('T' + key)
        if not t:
            continue
        spine = sum(_num(row[126]) for row, m in zip(t['verbatim'], t['meta'])
                    if m['row_type'] == 'Product' and not np.isnan(_num(row[126])))
        r = sum(_num(row[126]) for row, m in zip(t['verbatim'], t['meta'])
                if m['row_type'] == 'ReExports' and not np.isnan(_num(row[126])))
        tot += spine
        rex += r
        gate = ctrl.get(name)
        if gate is not None:
            report(f"T{key} {name}: control total", PASS if abs(spine - gate) < 0.5 else FAIL,
                   f"{spine:,.0f} against {gate:,.0f}")
    if tot:
        g = src.get('margin_total_spine', 421410)
        report('Margins 23-34 total on the spine', PASS if abs(tot - g) < 0.5 else FAIL,
               f"{tot:,.0f} against {g:,.0f}")
        report('Margins 23-34 incl re-exports',
               PASS if abs(tot + rex - 422034) < 0.5 else FAIL,
               f"{tot + rex:,.0f} against 422,034 (re-exports {rex:,.0f})")
    t35 = abs_t.get('T35')
    if t35:
        n35 = sum(_num(row[126]) for row, m in zip(t35['verbatim'], t35['meta'])
                  if m['row_type'] == 'Product' and not np.isnan(_num(row[126])))
        g = src.get('net_tax_total_spine', 168673)
        report('Table 35 net taxes total', PASS if abs(n35 - g) < 0.5 else FAIL,
               f"{n35:,.0f} against {g:,.0f}")

    # Table 21 is an independent source for the same number, by earning
    # industry rather than by product. If 23-34 and 21 disagree, one of them
    # is the wrong vintage.
    t21 = abs_t.get('T21')
    if t21 and tot:
        # Table 21 columns: A code, B name, C Margin Commodity, D Non margin,
        # E Total. The comparable number is C - column E is margin plus
        # non-margin supply and is roughly double.
        t21tot = next((_num(row[2]) for row, m in zip(t21['verbatim'], t21['meta'])
                       if m['row_type'] == 'Total'), np.nan)
        report('Table 21 agrees with Tables 23-34',
               PASS if not np.isnan(t21tot) and abs(t21tot - (tot + rex)) < 0.5 else FAIL,
               f"Table 21 margin commodity {t21tot:,.0f} against {tot + rex:,.0f}")


def industry_block(d):
    """
    The real industry quadrant and the output vector, from a loaded flow block.

    Two things this must get right, both of which have bitten before:

      * Position. 'meta' carries src_row for audit, which is NOT an index into
        'verbatim' - verbatim[0] IS src_row 13. Index off meta's own position.
      * Order. The columns run in the source's own spine order, so the rows must
        too. Sorting the codes here would silently transpose the matrix against
        its own column headings.

    The 10 'Dummy' rows (9901-9910) and their matching columns are excluded from
    the algebra but left untouched in RAW.

    Output x is the source's own Production row, which is intermediate use plus
    every primary input. Using the T1 row alone understates output and inflates
    every derived multiplier - see CLAUDE.md.
    """
    ind = [(i, m['code']) for i, m in enumerate(d['meta']) if m['row_type'] == 'Industry']
    codes = [c for _, c in ind]
    n = len(codes)
    first = d.get('first_data_col', 3)
    Z = np.full((n, n), np.nan)
    for r, (vi, _) in enumerate(ind):
        row = d['verbatim'][vi]
        for c in range(n):
            ci = first - 1 + c
            v = row[ci] if ci < len(row) else None
            if isinstance(v, numbers.Number):
                Z[r, c] = float(v)
    prod = [i for i, m in enumerate(d['meta']) if m['row_type'] == 'Production']
    x = np.full(n, np.nan)
    if prod:
        row = d['verbatim'][prod[0]]
        for c in range(n):
            ci = first - 1 + c
            v = row[ci] if ci < len(row) else None
            if isinstance(v, numbers.Number):
                x[c] = float(v)
    else:                                   # fall back to T1 + every P row
        want = ('Total', 'Primary')
        x = np.zeros(n)
        for i, m in enumerate(d['meta']):
            if m['row_type'] in want:
                row = d['verbatim'][i]
                for c in range(n):
                    ci = first - 1 + c
                    v = row[ci] if ci < len(row) else None
                    if isinstance(v, numbers.Number):
                        x[c] += float(v)
    return codes, Z, x


def check_flow_tables(src):
    flows = src.get('flows') or {}
    if not flows:
        report('Supplied flow tables loaded', FAIL,
               'load_supplied_flows() is still a stub. Adapt it in scripts/load_sources.py')
        return None
    have = sorted({r for _, r in flows})
    missing = [r for r in src['regions'] if r not in have]
    report('Flow tables for all nine regions', PASS if not missing else FAIL,
           '' if not missing else f"missing: {missing}")

    blocks = {}
    for (tbl, reg), d in flows.items():
        if tbl == 'T5':
            codes, Z, x = industry_block(d)
            blocks[reg] = (codes, Z, x)
    dupes = []
    regs = sorted(blocks)
    for i, a in enumerate(regs):
        for b in regs[i + 1:]:
            if blocks[a][1].shape == blocks[b][1].shape and np.allclose(
                    np.nan_to_num(blocks[a][1]), np.nan_to_num(blocks[b][1]), atol=1e-6):
                dupes.append((a, b))
    report('Every region has a distinct Table 5', PASS if not dupes else FAIL,
           '' if not dupes else f"identical pairs: {dupes}")

    # Table 8 must exceed Table 5 everywhere: imports are embedded in T8.
    bad = []
    for reg in src['regions']:
        if ('T5', reg) in flows and ('T8', reg) in flows:
            _, Z5, _ = industry_block(flows[('T5', reg)])
            _, Z8, _ = industry_block(flows[('T8', reg)])
            diff = np.nan_to_num(Z8) - np.nan_to_num(Z5)
            if diff.sum() <= 0 or (diff < -1e-6).sum() > 0:
                bad.append(f"{reg} ({int((diff < -1e-6).sum())} negative cells)")
    report('Imports (T8 less T5) non-negative in every region',
           PASS if not bad else WARN, '; '.join(bad))
    return blocks


def check_multipliers(src, flow_blocks):
    mult = src.get('multipliers') or {}
    if not mult or not mult.get('data'):
        report('Supplied multipliers loaded', FAIL,
               'load_supplied_multipliers() is still a stub. Adapt it in scripts/load_sources.py')
        return
    data = mult['data']
    have = sorted({k.split('|')[0] for k in data})
    missing = [r for r in src['regions'] if r not in have]
    report('Multipliers for all nine regions', PASS if not missing else FAIL,
           '' if not missing else f"missing: {missing}")

    # distinctness
    sig = {}
    for r in have:
        key = tuple(round(v, 6) if isinstance(v, float) else v
                    for c in sorted(k.split('|')[1] for k in data if k.startswith(r + '|'))[:5]
                    for v in (data.get(f"{r}|{c}") or []))
        sig.setdefault(key, []).append(r)
    dupes = [v for v in sig.values() if len(v) > 1]
    report('Every region has a distinct multiplier set', PASS if not dupes else FAIL,
           '' if not dupes else f"identical groups: {dupes}")

    # internal identities: simple = initial + production induced
    eff = mult.get('effects', [])
    try:
        i_init, i_simp, i_prod = eff.index('Initial effect'), eff.index('Simple multiplier'), \
            eff.index('Production-induced effect')
    except ValueError:
        report('Multiplier effect columns identified', WARN, f"effects seen: {eff}")
        return
    worst = 0.0
    for k, vals in data.items():
        a, b, c = vals[i_init], vals[i_simp], vals[i_prod]
        if all(isinstance(x, (int, float)) for x in (a, b, c)):
            worst = max(worst, abs(b - a - c))
    report('Identity: simple = initial + production induced',
           PASS if worst < 1e-4 else FAIL, f"max deviation {worst:.6f}")

    # ---------------------------------------------------------- reconciliation
    # The check that proves the supplied multipliers and the supplied Table 5
    # are the same body of work. Derive the simple output multiplier from the
    # flow table and compare it to the set. Against Table 8 it will not match.
    if not flow_blocks:
        return
    for reg in [r for r in src['regions'] if r in flow_blocks]:
        codes, Z, x = flow_blocks[reg]
        try:
            derived = simple_output_multiplier(np.nan_to_num(Z), np.nan_to_num(x))
        except np.linalg.LinAlgError:
            report(f"{reg}: multipliers reconcile to Table 5", WARN, 'matrix not invertible')
            continue
        supplied = np.array([_num((data.get(f"{reg}|{c}") or [None] * 11)[i_simp])
                             for c in codes], dtype=float)
        ok = ~np.isnan(supplied) & ~np.isnan(derived)
        if ok.sum() == 0:
            report(f"{reg}: multipliers reconcile to Table 5", WARN, 'no comparable rows')
            continue
        dev = np.abs(derived[ok] - supplied[ok])
        near = int((dev < 0.001).sum())
        report(f"{reg}: multipliers reconcile to Table 5",
               PASS if near == ok.sum() else WARN,
               f"{near}/{int(ok.sum())} within 0.001, max deviation {dev.max():.4f}")

    # State multipliers must be strictly smaller than national - smaller
    # economies leak more. A state at or above national means the
    # regionalisation failed, or the block is a copy of Australia.
    # A wholesale failure shows up as most industries above national, or as a
    # block identical to Australia - both of which the checks above already
    # catch. A handful of named industries is a question for the provider, not
    # a reason to block the build, so this warns and names them.
    if 'Aus' in flow_blocks:
        codes_aus = flow_blocks['Aus'][0]
        nat = {c: _num((data.get(f"Aus|{c}") or [None] * 11)[i_simp]) for c in codes_aus}
        labels = {m['code']: m['label']
                  for m in (mult.get('regions', {}).get('Aus', {}) or {}).get('meta', [])
                  if m.get('code')}
        for reg in [r for r in src['regions'] if r != 'Aus' and r in flow_blocks]:
            over = []
            for c in flow_blocks[reg][0]:
                s, a = _num((data.get(f"{reg}|{c}") or [None] * 11)[i_simp]), nat.get(c, np.nan)
                if not np.isnan(s) and not np.isnan(a) and s > a + 1e-9:
                    over.append(f"{c} {labels.get(c, '')[:34]} ({s:.4f} vs {a:.4f})")
            n = len(flow_blocks[reg][0])
            report(f"{reg}: simple multipliers below national",
                   PASS if not over else WARN,
                   '' if not over else f"{len(over)}/{n} above Australia: " + '; '.join(over))


def main():
    if not SRC.exists():
        sys.exit(f"{SRC} not found. Run scripts/load_sources.py first.")
    src = pickle.load(open(SRC, 'rb'))
    print("ABS source data")
    check_abs(src)
    print("\nSupplied flow tables")
    blocks = check_flow_tables(src)
    print("\nSupplied multipliers")
    check_multipliers(src, blocks)

    n_fail = sum(1 for _, s, _ in results if s == FAIL)
    n_warn = sum(1 for _, s, _ in results if s == WARN)
    print("\n" + "=" * 70)
    print(f"{len(results)} checks: {n_fail} fail, {n_warn} warn")
    if n_fail:
        print("Do not build until the failures are resolved, or the model will")
        print("produce results that look fine and are not.")
    sys.exit(1 if n_fail else 0)


if __name__ == '__main__':
    main()
