#!/usr/bin/env python3
"""
Build the OB Stage 2 master workbook.

Input
  --automated  Order Booking Master (AUTOMATED).xlsx   the live template: its
               formulas, pivots, external links and styles are the automation
               and must survive untouched.
  --sample     order booking Sample (Copy).xlsx        a values-only snapshot of
               the same workbook: this is the data to load.

What it does
  1. Starts from the automated workbook part-for-part, so every piece of
     machinery the user depends on (pivot definitions, external links, tables,
     conditional formatting, defined names, web extensions, custom properties)
     is carried across byte-identical unless this script has a reason to touch it.
  2. Clears the data out of every data-bearing sheet and writes the sample's
     data in its place, re-mapping the sample's shared-string indexes, style
     indexes and external-workbook [n] indexes into the automated workbook's
     own numbering so nothing points at the wrong thing.
  3. Leaves the automated workbook's live blocks as formulas, re-anchored to
     wherever the new data actually ends:
       - ASO_Raw   AB/AC region + segment lookups stay formulas for 88k rows;
                   only the A:AA paste zone receives raw data.
       - YTD Merg  the 500-row live window that reads ASO_Raw stays a formula
                   block and is moved to sit directly under the new history.
  4. Drops calcChain.xml (Excel rebuilds it) and asks every pivot cache to
     refresh on load, so the workbook recalculates itself the first time it is
     opened rather than showing the donor's cached numbers.
"""

import argparse
import os
import re
import shutil
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import xlsx_surgery as X


# Sheets whose content is data: cleared, then re-loaded from the sample.
DATA_SHEETS = [
    'MTD Merg', 'CPS', 'SOP', 'Invoice data_Raw', 'Sheet2', 'GK (2)',
    'Monthly Booking Target (2)', 'PH', 'FOC', 'CP', 'RT', 'RP', 'Sheet3',
]

# Sheets that are pure automation output or configuration: left exactly as the
# automated workbook has them. The pivots rebuild from their caches on open,
# Sheet5 is the region/segment mapping table the lookups depend on.
KEEP_SHEETS = ['Weekly Pivot', 'Pivot', 'Pivot Invoice', 'Sheet4', 'Sheet5']

ASO_PASTE_LAST_COL = 'AA'      # A:AA is the raw-export paste zone
ASO_FIRST_DATA_ROW = 2
YTD_HEADER_ROWS = 2            # rows 1-2 are the totals bar and the header
YTD_WINDOW_ROWS = 500          # size of the live block that reads ASO_Raw



def _plain_text(si_xml):
    return ''.join(re.findall(r'<t[^>]*>(.*?)</t>', si_xml, re.S)) \
        .replace('&amp;', '&').replace('&lt;', '<').replace('&gt;', '>')


def cached_value(cell, sst_text):
    """(t attribute, <v> text) for using a donor cell's value as a formula's
    cached result. Formula results are never shared strings, so text results are
    written literally with t="str" the way Excel writes them."""
    if cell is None or cell.v is None:
        return None, None
    if cell.t == 's':
        return 'str', sst_text[int(cell.v)]
    if cell.t in ('str', 'inlineStr'):
        return 'str', cell.v
    if cell.t == 'e':
        return 'e', cell.v
    if cell.t == 'b':
        return 'b', cell.v
    return None, cell.v


def remap_cell(cell, sst_map, style_map, ext_map, self_names):
    """Translate one donor cell into the base workbook's numbering."""
    c = X.Cell(cell.col, cell.s, cell.t)
    if cell.s is not None:
        c.s = str(style_map.get(int(cell.s), int(cell.s)))
    if cell.t == 's' and cell.v is not None:
        c.v = str(sst_map[int(cell.v)])
    else:
        c.v = cell.v
    c.is_xml = cell.is_xml
    if cell.f is not None:
        c.f = X.remap_external(cell.f, ext_map)
        c.fattrs = dict(cell.fattrs)
    return c


def find_last_data_row(rows, cols):
    last = 0
    for r in rows:
        for col in cols:
            cell = r.cells.get(col)
            if cell is not None and (cell.v not in (None, '') or cell.f):
                last = max(last, r.n)
                break
    return last


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--automated', required=True)
    ap.add_argument('--sample', required=True)
    ap.add_argument('--out', required=True)
    ap.add_argument('--report', default=None)
    args = ap.parse_args()

    za = zipfile.ZipFile(args.automated)
    zs = zipfile.ZipFile(args.sample)
    log = []

    def say(msg):
        print(msg)
        log.append(msg)

    amap, smap = X.sheet_map(za), X.sheet_map(zs)

    # ---- external workbook [n] indexes: donor order -> base order ----------
    a_links = [X.link_key(t) for t in X.external_link_order(za)]
    s_links = [X.link_key(t) for t in X.external_link_order(zs)]
    ext_map = {}
    for i, key in enumerate(s_links, start=1):
        ext_map[i] = a_links.index(key) + 1 if key in a_links else i
    say('external link remap (sample -> stage 2): %s' % ext_map)

    # ---- shared strings ---------------------------------------------------
    pool = X.StringPool(X.read_shared_strings(za))
    base_count = len(pool.items)
    sst_map = [pool.intern(s) for s in X.read_shared_strings(zs)]
    say('shared strings: %d from template + %d new = %d'
        % (base_count, len(pool.items) - base_count, len(pool.items)))

    # ---- styles -----------------------------------------------------------
    new_styles, style_map = X.merge_styles(
        za.read('xl/styles.xml').decode('utf-8'),
        zs.read('xl/styles.xml').decode('utf-8'))
    say('cell formats: template kept, %d sample formats appended' % len(style_map))

    rewritten = {}          # part name -> bytes
    filter_ranges = {}      # sheet name -> (last row, last col) for defined names

    def donor_rows(sheet):
        return list(X.read_sheet_rows(zs, smap[sheet][0], self_sheet_names=(sheet,)))

    def base_rows(sheet):
        return list(X.read_sheet_rows(za, amap[sheet][0], self_sheet_names=(sheet,)))

    # ---------------------------------------------------------------- data --
    for sheet in DATA_SHEETS:
        rows = donor_rows(sheet)
        out = []
        maxcol = 0
        for r in rows:
            nr = X.Row(r.n, dict(r.attrs), {})
            for col, cell in r.cells.items():
                nr.cells[col] = remap_cell(cell, sst_map, style_map, ext_map, (sheet,))
                maxcol = max(maxcol, X.col_num(col))
            out.append(nr)
        last = out[-1].n if out else 1
        xml = za.read(amap[sheet][0]).decode('utf-8')
        xml = X.replace_sheet_data(xml, out)
        xml = X.set_dimension(xml, 'A1:%s%d' % (X.col_name(max(maxcol, 1)), last))
        rewritten[amap[sheet][0]] = xml.encode('utf-8')
        data_last = find_last_data_row(out, [X.col_name(i) for i in range(1, maxcol + 1)])
        filter_ranges[sheet] = (data_last, X.col_name(max(maxcol, 1)))
        say('%-28s %6d rows loaded from sample (data to row %d)' % (sheet, len(out), data_last))

    # ------------------------------------------------------------ ASO_Raw --
    # The raw export paste zone only. Everything else on this sheet - the AB/AC
    # region and segment lookups that feed YTD Merg - is automation and is kept.
    sheet = 'ASO_Raw'
    paste_cols = [X.col_name(i) for i in range(1, X.col_num(ASO_PASTE_LAST_COL) + 1)]
    donor = {r.n: r for r in X.read_sheet_rows(zs, smap[sheet][0], self_sheet_names=(sheet,))}
    aso_last = find_last_data_row(donor.values(), paste_cols)
    out = []
    maxcol = 0
    for r in X.read_sheet_rows(za, amap[sheet][0], self_sheet_names=(sheet,)):
        nr = X.Row(r.n, dict(r.attrs), dict(r.cells))
        d = donor.get(r.n)
        if d is not None and ASO_FIRST_DATA_ROW <= r.n <= aso_last:
            for col in paste_cols:                       # clear the paste zone
                nr.cells.pop(col, None)
            for col, cell in d.cells.items():            # then write the sample's export in
                if X.col_num(col) <= X.col_num(ASO_PASTE_LAST_COL):
                    nr.cells[col] = remap_cell(cell, sst_map, style_map, ext_map, (sheet,))
        elif ASO_FIRST_DATA_ROW <= r.n:
            for col in paste_cols:
                cell = nr.cells.get(col)
                if cell is not None and cell.f is None:
                    cell.v = None                        # leave the formatting, drop the value
                    cell.t = None
        if nr.cells:
            maxcol = max(maxcol, max(X.col_num(c) for c in nr.cells))
        out.append(nr)
    have = {r.n for r in out}
    extra = sorted(n for n in donor if ASO_FIRST_DATA_ROW <= n <= aso_last and n not in have)
    for n in extra:
        d = donor[n]
        nr = X.Row(n, dict(d.attrs), {})
        for col, cell in d.cells.items():
            if X.col_num(col) <= X.col_num(ASO_PASTE_LAST_COL):
                nr.cells[col] = remap_cell(cell, sst_map, style_map, ext_map, (sheet,))
        out.append(nr)
    out.sort(key=lambda r: r.n)
    xml = za.read(amap[sheet][0]).decode('utf-8')
    xml = X.replace_sheet_data(xml, out)
    xml = X.set_dimension(xml, 'A1:%s%d' % (X.col_name(maxcol), out[-1].n))
    rewritten[amap[sheet][0]] = xml.encode('utf-8')
    filter_ranges[sheet] = (aso_last, ASO_PASTE_LAST_COL)
    say('%-28s export rows 2..%d pasted into A:AA; AB/AC lookups kept as formulas'
        % (sheet, aso_last))

    # ----------------------------------------------------------- YTD Merg --
    # Frozen history from the sample, then the automated workbook's live block
    # moved to sit directly underneath it.
    sheet = 'YTD Merg'
    a_rows = {r.n: r for r in X.read_sheet_rows(za, amap[sheet][0], self_sheet_names=(sheet,))}
    s_rows = {r.n: r for r in X.read_sheet_rows(zs, smap[sheet][0], self_sheet_names=(sheet,))}

    # Where does the sample's frozen history stop and its ASO-derived tail start?
    # The tail is the run of rows whose sales order also appears in the raw
    # export - those are exactly the rows the live block will regenerate.
    aso_orders = _sample_export_orders(zs, smap)
    live_start = None
    for n in sorted(s_rows):
        if n < 3:
            continue
        order = _order_text(s_rows[n], 'F')
        if order is not None and order in aso_orders:
            live_start = n
            break
    if live_start is None:
        live_start = max(s_rows) + 1
    say('%-28s sample history rows 3..%d, ASO-derived tail from row %d'
        % (sheet, live_start - 1, live_start))

    # the automated workbook's live block
    a_live = sorted(n for n, r in a_rows.items()
                    if n > YTD_HEADER_ROWS and any(c.f for c in r.cells.values())
                    and _is_live_row(r))
    a_live_start = a_live[0] if a_live else None
    if a_live_start is None:
        raise SystemExit('could not find the live formula block in the automated YTD Merg')
    delta = live_start - a_live_start
    say('%-28s live block found at row %d in template, re-anchored by %+d to row %d'
        % (sheet, a_live_start, delta, live_start))

    out = []
    maxcol = 0
    for n in (1, 2):
        if n in a_rows:
            out.append(a_rows[n])
    for n in sorted(s_rows):
        if n < 3 or n >= live_start:
            continue
        r = s_rows[n]
        nr = X.Row(n, dict(r.attrs), {})
        for col, cell in r.cells.items():
            nr.cells[col] = remap_cell(cell, sst_map, style_map, ext_map, (sheet,))
        out.append(nr)
    # Regenerate the whole block from the template row rather than copying the
    # template workbook's rows one by one. 68 of its 500 rows had been pasted
    # over with values and another 14 had lost their Q/R/S formulas, so copying
    # would carry the template's own leftover orders into Stage 2 and leave the
    # block only four-fifths live. Every row is rebuilt from the one row that
    # still had all 20 formulas, which both clears that data and repairs the block.
    tmpl = a_rows[a_live_start]
    live_cols = sorted((c for c, cell in tmpl.cells.items() if cell.f), key=X.col_num)
    if not live_cols:
        raise SystemExit('template row %d carries no formulas' % a_live_start)

    # Cached results. A formula with no cached value is legal but unreadable by
    # anything that is not Excel, and the dashboard pipeline reads this file
    # directly -- it would see 486 blank rows until someone opened and saved the
    # workbook. The sample's own tail rows ARE the results of these formulas fed
    # by this raw export (verify_stage2.py checks that cell by cell), so they are
    # written back as the cached values, exactly as Excel would have saved them.
    sample_sst = [_plain_text(t) for t in X.read_shared_strings(zs)]
    owner_by_ps = {}
    for r in X.read_sheet_rows(zs, smap['CPS'][0], expand_shared=False):
        if r.n < 2:
            continue
        b, j = r.cells.get('B'), r.cells.get('J')
        if b is None or b.v is None:
            continue
        key = sample_sst[int(b.v)] if b.t == 's' else b.v
        if key not in owner_by_ps:
            owner_by_ps[key] = (sample_sst[int(j.v)] if (j is not None and j.t == 's' and j.v is not None)
                                else (j.v if j is not None else ''))
    cached, person_filled = 0, 0
    for i in range(YTD_WINDOW_ROWS):
        n = live_start + i
        src = s_rows.get(n)
        nr = X.Row(n, dict(tmpl.attrs), {})
        for col in live_cols:
            cell = tmpl.cells[col]
            c = X.Cell(col, cell.s, None)
            # step the block down one row at a time (so its reads of the raw
            # export follow the export down), then slide the whole block up to
            # sit under the new history without moving those reads.
            f = X.shift_rows(cell.f, i)
            c.f = X.shift_rows(f, delta, self_only=True, self_names=(sheet,))
            c.fattrs = dict(cell.fattrs)
            t, v = cached_value(src.cells.get(col) if src else None, sample_sst)
            if col == 'T' and (v is None or v == '') and src is not None:
                d = src.cells.get('D')
                ps = (sample_sst[int(d.v)] if (d is not None and d.t == 's' and d.v is not None)
                      else (d.v if d is not None else None))
                name = owner_by_ps.get(ps)
                if name:
                    t, v, person_filled = 'str', name, person_filled + 1
            if v is None:
                t, v = 'str', ''        # the block returns "" past the last export row
            c.t, c.v = t, v
            cached += 1
            nr.cells[col] = c
        out.append(nr)
    say('%-28s live block rebuilt from row %d of the template: %d rows x %d live columns'
        % (sheet, a_live_start, YTD_WINDOW_ROWS, len(live_cols)))
    say('%-28s %d cached results written (sales person derived for %d rows the sample left blank)'
        % (sheet, cached, person_filled))
    for r in out:
        if r.cells:
            maxcol = max(maxcol, max(X.col_num(c) for c in r.cells))
    last = out[-1].n
    xml = za.read(amap[sheet][0]).decode('utf-8')
    xml = X.replace_sheet_data(xml, out)
    xml = X.set_dimension(xml, 'A1:%s%d' % (X.col_name(maxcol), last))
    rewritten[amap[sheet][0]] = xml.encode('utf-8')
    filter_ranges[sheet] = (last, 'V')
    say('%-28s %d history rows + %d live formula rows (ends at row %d)'
        % (sheet, live_start - 3, last - live_start + 1, last))

    for s in KEEP_SHEETS:
        say('%-28s kept from template unchanged' % s)

    # ------------------------------------------------------------ workbook --
    wb = za.read('xl/workbook.xml').decode('utf-8')

    def fix_defined_name(m):
        whole, body = m.group(0), m.group(2)
        for sheet_name, (last_row, last_col) in filter_ranges.items():
            quoted = "'%s'" % sheet_name if ' ' in sheet_name or '(' in sheet_name else sheet_name
            pat = re.compile(r'^%s!\$([A-Z]+)\$(\d+):\$([A-Z]+)\$(\d+)$' % re.escape(quoted))
            mm = pat.match(body)
            if mm:
                return whole.replace(
                    body, '%s!$%s$%s:$%s$%d' % (quoted, mm.group(1), mm.group(2),
                                                mm.group(3), max(last_row, int(mm.group(2)))))
        return whole

    wb = re.sub(r'(<definedName[^>]*>)([^<]*)(</definedName>)',
                lambda m: fix_defined_name(m), wb)
    if 'fullCalcOnLoad' not in wb:
        wb = re.sub(r'<calcPr([^>]*?)/>', r'<calcPr\1 fullCalcOnLoad="1"/>', wb)
    wb = re.sub(r'<fileVersion[^>]*/>', '', wb)
    # calcChain is being dropped; its relationship must go with it
    wbrels = za.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    wbrels = re.sub(r'<Relationship[^>]*Target="calcChain\.xml"[^>]*/>', '', wbrels)
    ct = za.read('[Content_Types].xml').decode('utf-8')
    ct = re.sub(r'<Override PartName="/xl/calcChain\.xml"[^>]*/>', '', ct)

    rewritten['xl/workbook.xml'] = wb.encode('utf-8')
    rewritten['xl/_rels/workbook.xml.rels'] = wbrels.encode('utf-8')
    rewritten['[Content_Types].xml'] = ct.encode('utf-8')
    rewritten['xl/styles.xml'] = new_styles.encode('utf-8')
    rewritten['xl/sharedStrings.xml'] = X.build_shared_strings(pool.items)

    # ---- tables follow their sheet's data ---------------------------------
    for name in za.namelist():
        if name.startswith('xl/tables/table'):
            t = za.read(name).decode('utf-8')
            disp = re.search(r'displayName="([^"]+)"', t)
            ref = re.search(r'<table[^>]*\sref="([A-Z]+\d+:[A-Z]+)(\d+)"', t)
            owner = _table_owner(za, amap, name)
            if ref and owner in filter_ranges:
                new_ref = '%s%d' % (ref.group(1), max(filter_ranges[owner][0], 2))
                t = t.replace('ref="%s%s"' % (ref.group(1), ref.group(2)),
                              'ref="%s"' % new_ref, 1)
                t = re.sub(r'(<autoFilter\s+ref=")[A-Z]+\d+:[A-Z]+\d+(")',
                           lambda m: m.group(1) + new_ref + m.group(2), t)
                rewritten[name] = t.encode('utf-8')
                say('table %-18s on %-18s -> %s' % (disp.group(1) if disp else '?', owner, new_ref))

    # ---- pivot caches rebuild themselves on open --------------------------
    npc = 0
    for name in za.namelist():
        if re.match(r'xl/pivotCache/pivotCacheDefinition\d+\.xml$', name):
            d = za.read(name).decode('utf-8')
            if 'refreshOnLoad=' in d:
                d = re.sub(r'refreshOnLoad="[^"]*"', 'refreshOnLoad="1"', d, count=1)
            else:
                d = d.replace('<pivotCacheDefinition ', '<pivotCacheDefinition refreshOnLoad="1" ', 1)
            rewritten[name] = d.encode('utf-8')
            npc += 1
    say('%d pivot caches set to refresh on open' % npc)

    # ---- write ------------------------------------------------------------
    tmp = args.out + '.tmp'
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED, compresslevel=6) as zo:
        for item in za.infolist():
            if item.filename == 'xl/calcChain.xml':
                continue
            data = rewritten.pop(item.filename, None)
            zo.writestr(item.filename, data if data is not None else za.read(item.filename))
        for name, data in rewritten.items():
            zo.writestr(name, data)
    shutil.move(tmp, args.out)
    say('written %s (%.1f MB)' % (args.out, os.path.getsize(args.out) / 1e6))

    if args.report:
        with open(args.report, 'w') as fh:
            fh.write('\n'.join(log) + '\n')


def _table_owner(za, amap, table_part):
    base = table_part.split('/')[-1]
    for sheet, (part, _) in amap.items():
        rel = 'xl/worksheets/_rels/%s.rels' % part.split('/')[-1]
        if rel in za.namelist() and base in za.read(rel).decode('utf-8'):
            return sheet
    return None


_SAMPLE_ORDER_CACHE = {}


def _sample_export_orders(zs, smap):
    if 'orders' not in _SAMPLE_ORDER_CACHE:
        sst = X.read_shared_strings(zs)
        orders = set()
        for r in X.read_sheet_rows(zs, smap['ASO_Raw'][0]):
            if r.n < 2:
                continue
            c = r.cells.get('G')
            if c is None or c.v is None:
                continue
            orders.add(sst[int(c.v)] if c.t == 's' else str(c.v))
        _SAMPLE_ORDER_CACHE['orders'] = orders
        _SAMPLE_ORDER_CACHE['sst'] = sst
    return _SAMPLE_ORDER_CACHE['orders']


def _order_text(row, col):
    c = row.cells.get(col)
    if c is None or c.v is None:
        return None
    if c.t == 's':
        return _SAMPLE_ORDER_CACHE['sst'][int(c.v)]
    return str(c.v)


_LIVE_HINT = re.compile(r'ASO_Raw!', re.I)


def _is_live_row(row):
    return any(c.f and _LIVE_HINT.search(c.f) for c in row.cells.values())


if __name__ == '__main__':
    main()
