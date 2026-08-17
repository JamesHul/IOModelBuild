"""
Integrity gates on the loaded source data. Run before every build.

These are the checks that have actually caught things:
  - six of eight state flow-table blocks identical to NSW
  - five of eight multiplier sets identical to Australia
  - multipliers reconciling to Table 5 but not Table 8, confirming the provider
    used the direct-allocation table

    python scripts/check_sources.py
"""
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
    needed = ['T5', 'T8', 'T35'] + ['T' + k for k in src['margin_tables']]
    missing = [k for k in needed if k not in abs_t]
    report('ABS tables present', PASS if not missing else FAIL,
           '' if not missing else f"missing: {missing}")

    for key, (name, ind) in src['margin_tables'].items():
        k = 'T' + key
        if k not in abs_t:
            continue
        t = abs_t[k]
        n = len(t['row_index'])
        report(f"{k} {name}: spine codes found", PASS if n >= 115 else WARN,
               f"{n} four-digit codes in column {t['code_col']}")


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

    # region distinctness on Table 5
    blocks = {}
    for (tbl, reg), d in flows.items():
        if tbl != 'T5':
            continue
        codes = sorted(c for c in d['row_index'] if norm_code(c))
        first_col = d.get('first_data_col', (d['code_col'] or 0) + 2)
        blocks[reg] = (codes, numeric_block(d['verbatim'], d['row_index'], codes,
                                            first_col, len(codes)))
    dupes = []
    regs = sorted(blocks)
    for i, a in enumerate(regs):
        for b in regs[i + 1:]:
            if blocks[a][1].shape == blocks[b][1].shape and np.allclose(
                    np.nan_to_num(blocks[a][1]), np.nan_to_num(blocks[b][1]), atol=1e-6):
                dupes.append((a, b))
    report('Every region has a distinct Table 5', PASS if not dupes else FAIL,
           '' if not dupes else f"identical pairs: {dupes}")
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

    # reconciliation against the flow tables
    if not flow_blocks:
        return
    for reg, (codes, Z) in flow_blocks.items():
        x = Z.sum(0)  # TODO replace with intermediate use + ALL primary input rows
        try:
            derived = simple_output_multiplier(np.nan_to_num(Z), x)
        except np.linalg.LinAlgError:
            report(f"{reg}: multipliers reconcile to Table 5", WARN, 'matrix not invertible')
            continue
        supplied = np.array([(data.get(f"{reg}|{c}") or [np.nan] * 11)[i_simp] for c in codes],
                            dtype=float)
        ok = ~np.isnan(supplied)
        if ok.sum() == 0:
            continue
        d = np.abs(derived[ok] - supplied[ok])
        near = int((d < 0.001).sum())
        report(f"{reg}: multipliers reconcile to Table 5",
               PASS if near == ok.sum() else WARN,
               f"{near}/{int(ok.sum())} within 0.001, max deviation {d.max():.4f}. "
               f"If this fails, check the output denominator includes every P-row.")


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
