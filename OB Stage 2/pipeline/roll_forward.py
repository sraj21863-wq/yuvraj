#!/usr/bin/env python3
"""
Roll the workbook forward: freeze the live block, lay a fresh one under it.

The live block in YTD Merg is 500 rows of formulas that read ASO_Raw rows
2..501. That is the whole trick and also the whole limit: it can only ever show
one export at a time, and the moment you paste the next export over ASO_Raw the
previous one stops existing. Before that happens the rows it produced have to
become ordinary values -- history -- and a fresh block of formulas has to be
laid out underneath them, ready for the next paste.

Doing that by hand in Excel means paste-special-values over exactly the right
rows and then drag 500 rows of formulas down without disturbing the references.
This does it instead:

  1. reads how many rows of the live block the last export actually filled
  2. turns those rows into values, keeping the numbers Excel calculated
  3. writes a fresh 500-row block directly beneath them, reading ASO_Raw 2..501
  4. clears the ASO_Raw paste zone, ready for the next export
  5. keeps a timestamped backup of the workbook it started from

Run it AFTER you have opened the workbook in Excel and saved it, because the
numbers it freezes are the ones Excel last calculated. It checks, and refuses
rather than freezing a block of blanks.
"""

import argparse
import datetime as dt
import os
import re
import shutil
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import xlsx_surgery as X

REGISTER = 'YTD Merg'
RAW = 'ASO_Raw'
RAW_FIRST_ROW = 2
RAW_LAST_COL = 'AA'
WINDOW = 500
LIVE_MARK = re.compile(r'ASO_Raw!', re.I)
KEY_COL = 'F'          # sales order: filled means the row carries an order


def live_rows(rows):
    return sorted(n for n, r in rows.items()
                  if any(c.f and LIVE_MARK.search(c.f) for c in r.cells.values()))


def filled(row):
    c = row.cells.get(KEY_COL)
    return c is not None and c.v not in (None, '')


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--workbook', default=None)
    ap.add_argument('--keep-export', action='store_true',
                    help='leave the ASO_Raw paste zone alone instead of clearing it')
    ap.add_argument('--dry-run', action='store_true',
                    help='say what would happen and change nothing')
    ap.add_argument('--window', type=int, default=WINDOW)
    args = ap.parse_args()

    here = os.path.dirname(os.path.abspath(__file__))
    wb_path = args.workbook
    if wb_path is None:
        import json
        with open(os.path.join(here, 'config.json'), encoding='utf-8') as fh:
            wb_path = json.load(fh)['workbook']
        wb_path = os.path.normpath(os.path.join(here, wb_path))
    if not os.path.exists(wb_path):
        raise SystemExit('workbook not found: %s' % wb_path)

    z = zipfile.ZipFile(wb_path)
    smap = X.sheet_map(z)
    rows = {r.n: r for r in X.read_sheet_rows(z, smap[REGISTER][0],
                                              self_sheet_names=(REGISTER,))}
    block = live_rows(rows)
    if not block:
        raise SystemExit('no live block found in %s -- nothing to roll forward' % REGISTER)
    start, end = block[0], block[-1]

    have_values = any(rows[n].cells.get(KEY_COL) is not None
                      and rows[n].cells[KEY_COL].v is not None for n in block)
    if not have_values:
        raise SystemExit(
            'the live block at rows %d-%d has no calculated values in it.\n'
            'Open the workbook in Excel and save it first -- this freezes the numbers\n'
            'Excel calculated, and there are none to freeze yet.' % (start, end))

    done = [n for n in block if filled(rows[n])]
    if not done:
        print('the live block at rows %d-%d is empty -- nothing to freeze.' % (start, end))
        print('Paste the export into %s A%d:%s and save, then run this again.'
              % (RAW, RAW_FIRST_ROW, RAW_LAST_COL))
        return 0
    # the filled rows must be the top of the block; a gap means the export was
    # pasted somewhere unexpected and freezing would bury live rows under history
    expected = list(range(start, start + len(done)))
    if done != expected:
        raise SystemExit(
            'the filled rows in the live block are not a solid run from row %d.\n'
            'Filled: %s ... %s. Fix the paste in %s before rolling forward.'
            % (start, done[:3], done[-3:], RAW))

    k = len(done)
    new_start = start + k
    new_end = new_start + args.window - 1
    if new_end > X.MAX_ROW:
        raise SystemExit('a new %d-row block would run past the end of the sheet'
                         % args.window)

    print('%s: history ends at row %d, live block at rows %d-%d'
          % (REGISTER, start - 1, start, end))
    print('  %d of its %d rows carry an order -- these become history (rows %d-%d)'
          % (k, len(block), start, start + k - 1))
    print('  a fresh %d-row live block goes in at rows %d-%d, reading %s rows %d-%d'
          % (args.window, new_start, new_end, RAW, RAW_FIRST_ROW, RAW_FIRST_ROW + args.window - 1))
    print('  the %s paste zone A%d:%s is %s'
          % (RAW, RAW_FIRST_ROW, RAW_LAST_COL,
             'left alone' if args.keep_export else 'cleared, ready for the next export'))
    if args.dry_run:
        print('\ndry run -- nothing was changed.')
        return 0

    template = rows[start]
    live_cols = sorted((c for c, cell in template.cells.items() if cell.f), key=X.col_num)

    out_rows = []
    for n in sorted(rows):
        if n > end:
            continue
        r = rows[n]
        if start <= n < start + k:
            # freeze: keep exactly what Excel calculated, drop the formula
            nr = X.Row(n, dict(r.attrs), {})
            for col, c in r.cells.items():
                nc = X.Cell(col, c.s, c.t)
                nc.v, nc.is_xml = c.v, c.is_xml
                if nc.t == 'str':
                    nc.t = None if _numeric(nc.v) else 'str'
                nr.cells[col] = nc
            out_rows.append(nr)
        elif n >= start:
            continue                      # the rest of the old block is replaced
        else:
            out_rows.append(r)

    delta = new_start - start
    for i in range(args.window):
        n = new_start + i
        nr = X.Row(n, dict(template.attrs), {})
        for col in live_cols:
            cell = template.cells[col]
            c = X.Cell(col, cell.s, 'str')
            f = X.shift_rows(cell.f, i)                      # follow the export down
            c.f = X.shift_rows(f, delta, self_only=True, self_names=(REGISTER,))
            c.fattrs = dict(cell.fattrs)
            c.v = ''                                         # empty until the next paste
            nr.cells[col] = c
        out_rows.append(nr)

    maxcol = max((X.col_num(c) for r in out_rows for c in r.cells), default=1)
    xml = z.read(smap[REGISTER][0]).decode('utf-8')
    xml = X.replace_sheet_data(xml, out_rows)
    xml = X.set_dimension(xml, 'A1:%s%d' % (X.col_name(maxcol), new_end))
    rewritten = {smap[REGISTER][0]: xml.encode('utf-8')}

    raw_last = RAW_FIRST_ROW
    if not args.keep_export:
        paste = [X.col_name(i) for i in range(1, X.col_num(RAW_LAST_COL) + 1)]
        raw_rows = []
        for r in X.read_sheet_rows(z, smap[RAW][0], expand_shared=False):
            if r.n >= RAW_FIRST_ROW:
                for col in paste:
                    c = r.cells.get(col)
                    if c is not None and c.f is None:
                        if c.v not in (None, ''):
                            raw_last = max(raw_last, r.n)
                        c.v, c.t = None, None        # keep the formatting, drop the value
            raw_rows.append(r)
        rxml = z.read(smap[RAW][0]).decode('utf-8')
        rxml = X.replace_sheet_data(rxml, raw_rows)
        rewritten[smap[RAW][0]] = rxml.encode('utf-8')

    # autofilter range and the ASO_Raw table follow the data
    wbx = z.read('xl/workbook.xml').decode('utf-8')
    wbx = re.sub(r"('%s'!\$A\$2:\$([A-Z]+)\$)\d+" % re.escape(REGISTER),
                 lambda m: '%s%d' % (m.group(1), new_end), wbx)
    rewritten['xl/workbook.xml'] = wbx.encode('utf-8')

    backup = os.path.join(os.path.dirname(wb_path), 'backups')
    os.makedirs(backup, exist_ok=True)
    stamp = dt.datetime.now().strftime('%Y%m%d-%H%M%S')
    keep = os.path.join(backup, os.path.basename(wb_path).replace(
        '.xlsx', '--before-roll-%s.xlsx' % stamp))
    shutil.copy2(wb_path, keep)

    tmp = wb_path + '.tmp'
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED, compresslevel=6) as zo:
        for item in z.infolist():
            data = rewritten.pop(item.filename, None)
            zo.writestr(item.filename, data if data is not None else z.read(item.filename))
        for name, data in rewritten.items():
            zo.writestr(name, data)
    z.close()
    shutil.move(tmp, wb_path)

    print('\ndone.')
    print('  backup   %s' % keep)
    print('  history  rows 3-%d' % (new_start - 1))
    print('  live     rows %d-%d' % (new_start, new_end))
    print('  next step: paste the new export into %s A%d:%s and save.'
          % (RAW, RAW_FIRST_ROW, RAW_LAST_COL))
    return 0


def _numeric(v):
    if v is None:
        return False
    try:
        float(v)
        return True
    except (TypeError, ValueError):
        return False


if __name__ == '__main__':
    raise SystemExit(main())
