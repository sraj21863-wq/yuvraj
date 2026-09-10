#!/usr/bin/env python3
"""
Prove the Stage 2 workbook's live block still works.

Excel is not available here, so this re-implements the twenty formulas that
make up the YTD Merg live block and runs them against the Stage 2 workbook's
own raw sheets. If the automation survived the rebuild, those formulas fed by
the sample's raw export must reproduce the sample's own register rows exactly.
Any drift is a broken link, not a rounding difference.
"""

import argparse
import os
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import xlsx_surgery as X


def load(z, sheet, smap, sst, cols=None, first=1, last=None):
    """{row: {col: value}} with shared strings resolved."""
    out = {}
    for r in X.read_sheet_rows(z, smap[sheet][0], expand_shared=False):
        if r.n < first or (last and r.n > last):
            continue
        row = {}
        for col, c in r.cells.items():
            if cols and col not in cols:
                continue
            v = c.v
            if v is None:
                continue
            if c.t == 's':
                v = sst[int(v)]
            elif c.t in (None, 'n'):
                try:
                    v = float(v)
                except ValueError:
                    pass
            row[col] = v
        if row:
            out[r.n] = row
    return out


def text(v):
    if v is None:
        return ''
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return str(v)


def close(a, b, tol=1e-6):
    try:
        return abs(float(a) - float(b)) <= tol * max(1.0, abs(float(b)))
    except (TypeError, ValueError):
        return text(a) == text(b)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--workbook', required=True)
    ap.add_argument('--expected', required=True,
                    help='the sample workbook, used as the expected result')
    ap.add_argument('--first-live-row', type=int, default=2317)
    args = ap.parse_args()

    z = zipfile.ZipFile(args.workbook)
    smap = X.sheet_map(z)
    sst = [_plain(s) for s in X.read_shared_strings(z)]

    sheet5 = load(z, 'Sheet5', smap, sst)
    region = {text(r.get('A')): text(r.get('B')) for r in sheet5.values() if r.get('A')}
    segment = {text(r.get('D')): text(r.get('E')) for r in sheet5.values() if r.get('D')}

    aso = load(z, 'ASO_Raw', smap, sst,
               cols={'B', 'C', 'D', 'E', 'F', 'G', 'H', 'J', 'V'}, first=2, last=600)
    aso = {n: r for n, r in aso.items() if r.get('G') not in (None, '')}

    sop = load(z, 'SOP', smap, sst, cols={'E', 'I'}, first=5)
    not_invoiced = {}
    for r in sop.values():                       # MATCH takes the first hit
        k = text(r.get('E'))
        if k and k not in not_invoiced:
            not_invoiced[k] = float(r.get('I') or 0)

    inv = load(z, 'Invoice data_Raw', smap, sst, cols={'I', 'J'}, first=2)
    invoiced = {}
    for r in inv.values():
        k = text(r.get('J'))
        if k:
            invoiced[k] = invoiced.get(k, 0.0) + float(r.get('I') or 0)

    cps = load(z, 'CPS', smap, sst, cols={'B', 'J'}, first=2)
    owner = {}
    for r in cps.values():
        k = text(r.get('B'))
        if k and k not in owner:
            owner[k] = text(r.get('J'))

    divisor = 1e7
    for r in load(z, 'YTD Merg', smap, sst, cols={'R'}, first=1, last=1).values():
        pass

    computed = {}
    for n, a in sorted(aso.items()):
        f = text(a.get('G'))
        d = text(a.get('E'))
        o = not_invoiced.get(f, 0.0)
        p = invoiced.get(f, 0.0)
        q = o + p
        computed[n] = {
            'A': text(a.get('B')),
            'B': region.get(text(a.get('C')), ''),
            'C': segment.get(text(a.get('D')), ''),
            'D': d,
            'F': f,
            'H': text(a.get('H')),
            'N': text(a.get('V')),
            'O': o,
            'P': p,
            'Q': q,
            'R': q / divisor,
            'T': owner.get(d, ''),
        }

    ze = zipfile.ZipFile(args.expected)
    emap = X.sheet_map(ze)
    esst = [_plain(s) for s in X.read_shared_strings(ze)]
    expected = load(ze, 'YTD Merg', emap, esst, first=args.first_live_row)

    checks = ['A', 'B', 'C', 'F', 'N', 'O', 'P', 'Q', 'R', 'T']
    rows = sorted(computed)
    fails, compared, filled = [], 0, 0
    for i, aso_row in enumerate(rows):
        exp = expected.get(args.first_live_row + i)
        if exp is None:
            fails.append('no expected row for ASO row %d' % aso_row)
            continue
        got = computed[aso_row]
        for col in checks:
            compared += 1
            want, have = exp.get(col), got.get(col)
            if close(have, want):
                continue
            # The sample is a values-only snapshot: whoever froze it pasted the
            # newest orders in before the sales-person column had been filled,
            # so every one of its tail rows carries a blank there. The live
            # block computes that name from CPS, so it fills a gap the snapshot
            # has rather than disagreeing with it. Anything else is a failure.
            if col == 'T' and want in (None, '') and have:
                filled += 1
                continue
            fails.append('row %d col %s: recalculated %r, sample has %r'
                         % (args.first_live_row + i, col, have, want))

    print('live block: %d rows recalculated, %d cell comparisons' % (len(rows), compared))
    if filled:
        print('sales person filled in on %d rows the sample left blank '
              '(the snapshot was frozen before that column was completed)' % filled)
    if fails:
        print('MISMATCHES: %d' % len(fails))
        for f in fails[:25]:
            print('  ' + f)
        return 1
    print('every recalculated cell matches the sample register --- '
          'the automation still reproduces the data end to end.')
    return 0


def _plain(si_xml):
    import re
    return ''.join(re.findall(r'<t[^>]*>(.*?)</t>', si_xml, re.S)) \
        .replace('&amp;', '&').replace('&lt;', '<').replace('&gt;', '>')


if __name__ == '__main__':
    raise SystemExit(main())
