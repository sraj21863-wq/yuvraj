#!/usr/bin/env python3
"""
Order Booking Stage 2 -- workbook to dashboard payload.

Reads the Stage 2 master workbook and writes the JSON the dashboard renders.
Every rule the pipeline applies lives in config.json, so a column moving in the
workbook is a config edit, not a code change.

Two principles run through this file:

  Nothing is dropped in silence. Booking that falls outside the reporting grid
  (EXPORT, INTERNAL, FOC, Others, Quality) is carried in its own bucket and
  shown on the dashboard. Sales-person names that appear in the register but
  not in the target register are written to reconciliation.txt and raised as a
  warning on screen. A row that cannot be placed in a month is counted and
  reported. The payload never quietly becomes smaller than the workbook.

  Booked value and money are different measures and are never mixed. Column Q
  is the value of an order at release; columns O and P are the not-invoiced and
  invoiced amounts. The dashboard shows both, labelled, side by side.
"""

import argparse
import datetime as dt
import hashlib
import json
import os
import re
import shutil
import sys
import unicodedata

try:
    import openpyxl
except ImportError:                                          # pragma: no cover
    sys.exit('openpyxl is required:  pip install -r requirements.txt')

HERE = os.path.dirname(os.path.abspath(__file__))


# ---------------------------------------------------------------- helpers --

def col_index(letter):
    n = 0
    for ch in letter.strip().upper():
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def load_config(path):
    with open(path, 'r', encoding='utf-8') as fh:
        cfg = json.load(fh)
    return cfg


def resolve(path, base):
    return path if os.path.isabs(path) else os.path.normpath(os.path.join(base, path))


def clean(v):
    if v is None:
        return ''
    if isinstance(v, str):
        return v.strip()
    return v


def as_float(v):
    if v is None or v == '':
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(',', '').strip())
    except ValueError:
        return 0.0


def month_of(v):
    """Booking month from the register's month-start column."""
    if isinstance(v, dt.datetime):
        return v.year, v.month
    if isinstance(v, dt.date):
        return v.year, v.month
    if isinstance(v, (int, float)) and v > 1:
        base = dt.datetime(1899, 12, 30) + dt.timedelta(days=float(v))
        return base.year, base.month
    if isinstance(v, str):
        m = re.match(r'\s*(\d{4})-(\d{2})', v)
        if m:
            return int(m.group(1)), int(m.group(2))
    return None, None


def name_tokens(name):
    """The words in a person's name, lower-cased, with punctuation and bare
    initials removed. The two registers spell the same person several ways --
    "Uthayan B, Uthayan" against "B, Uthayan", "P, Sharath Kumar" against
    "Kumar,Sharath" -- so identity is compared as a set of real words, not as
    a string."""
    s = unicodedata.normalize('NFKD', str(name)).encode('ascii', 'ignore').decode()
    s = re.sub(r'\(deleted\)', ' ', s, flags=re.I)
    s = re.sub(r'[^A-Za-z0-9 ]+', ' ', s).lower()
    return frozenset(w for w in s.split() if len(w) > 1)


def norm_key(name):
    return name_tokens(name)


def match_people(register_names, target_names, aliases):
    """
    Pair each register name with a target name.

    Exact word-set first; then a containment match ("Nikhil" against "Sunil
    Ponkshe, Nikhil"), accepted only where it is unambiguous in both directions.
    Anything left over is returned unmatched rather than guessed at -- a wrong
    guess here quietly credits one person's booking to another.
    """
    by_key = {}
    for name in target_names:
        by_key.setdefault(name_tokens(name), []).append(name)
    exact = {k: v[0] for k, v in by_key.items() if len(v) == 1 and k}

    matched, fuzzy, unmatched = {}, {}, []
    pending = []
    for name in register_names:
        if name in aliases:
            matched[name] = aliases[name]
            continue
        key = name_tokens(name)
        hit = exact.get(key)
        if hit:
            matched[name] = hit
            continue
        pending.append((name, key))

    free_targets = [(n, name_tokens(n)) for n in target_names
                    if n not in set(matched.values())]
    for name, key in pending:
        if not key:
            unmatched.append(name)
            continue
        cands = [n for n, tk in free_targets if tk and (tk <= key or key <= tk)]
        rivals = [rn for rn, rk in pending
                  if rn != name and cands and any(
                      tk <= rk or rk <= tk for n, tk in free_targets if n in cands)]
        if len(cands) == 1 and len(rivals) == 0:
            matched[name] = cands[0]
            fuzzy[name] = cands[0]
        else:
            unmatched.append(name)
    return matched, fuzzy, unmatched


# ------------------------------------------------------------------ read ---

def read_register(ws, cfg, warn):
    c = {k: col_index(v) for k, v in cfg['register']['columns'].items()}
    width = max(c.values()) + 1
    div = float(cfg['register']['amountDivisor'])
    rows = []
    for raw in ws.iter_rows(min_row=cfg['register']['firstDataRow'],
                            max_col=width, values_only=True):
        region = clean(raw[c['region']])
        segment = clean(raw[c['segment']])
        order = clean(raw[c['salesOrder']])
        amount = raw[c['amountInr']]
        if not region and not segment and not order and amount in (None, ''):
            continue
        year, month = month_of(raw[c['bookingMonth']])
        rows.append({
            'region': str(region).upper(),
            'segment': str(segment).upper(),
            'ps': str(clean(raw[c['psNumber']])),
            'order': str(order),
            'customer': str(clean(raw[c['customer']])),
            'status': str(clean(raw[c['status']])),
            'notInvoiced': as_float(raw[c['notInvoiced']]) / div,
            'invoiced': as_float(raw[c['invoiced']]) / div,
            'value': as_float(raw[c['amountInr']]) / div,
            'year': year,
            'month': month,
            'person': str(clean(raw[c['salesPerson']])),
        })
    return rows


def read_targets(ws, cfg, warn):
    c = {k: col_index(v) for k, v in cfg['targets']['columns'].items()}
    first = c['firstMonth']
    width = first + 12
    out = []
    for raw in ws.iter_rows(min_row=cfg['targets']['firstDataRow'],
                            max_col=width, values_only=True):
        person = clean(raw[c['salesPerson']])
        region = clean(raw[c['region']])
        segment = clean(raw[c['segment']])
        if not person and not region:
            continue
        if not region and str(person).strip().lower() in ('total', 'grand total', 'sum'):
            continue          # the sheet's own footer row, not a person
        year = raw[c['year']]
        try:
            year = int(year)
        except (TypeError, ValueError):
            year = None
        out.append({
            'person': str(person),
            'region': str(region).upper(),
            'segment': str(segment).upper(),
            'territory': str(clean(raw[c['territory']])),
            'year': year,
            'monthly': [as_float(raw[first + i]) for i in range(12)],
        })
    return out


# ------------------------------------------------------------------ build --

def build(cfg, workbook_path, source_via, warn):
    wb = openpyxl.load_workbook(workbook_path, data_only=True, read_only=True)
    try:
        register = read_register(wb[cfg['register']['sheet']], cfg, warn)
        targets = read_targets(wb[cfg['targets']['sheet']], cfg, warn)
    finally:
        wb.close()

    nreg = {k.upper(): v for k, v in cfg['normalisation']['regions'].items()}
    nseg = {k.upper(): v for k, v in cfg['normalisation']['segments'].items()}
    regions = cfg['grid']['regions']
    segments = cfg['grid']['segments']
    fy = int(cfg['fiscalYear'])

    actual, count, status, status_count = {}, {}, {}, {}
    soip, invamt = {}, {}
    pactual, pcount = {}, {}
    excluded, excluded_by_month = {}, {}
    excluded_total = 0.0
    people_seen = {}
    wrong_year = undated = 0
    all_statuses = {}

    def add(d, k, v):
        d[k] = d.get(k, 0.0) + v

    for r in register:
        region = nreg.get(r['region'], r['region'] or '(unmapped)')
        segment = nseg.get(r['segment'], r['segment'] or '(unmapped)')
        month = r['month']
        if month is None:
            undated += 1
            continue
        if r['year'] != fy:
            wrong_year += 1
            continue
        all_statuses[r['status'] or '(blank)'] = all_statuses.get(r['status'] or '(blank)', 0) + 1

        if region in regions and segment in segments:
            key = '%d|%s|%s' % (month, region, segment)
            add(actual, key, r['value'])
            count[key] = count.get(key, 0) + 1
            add(soip, key, r['notInvoiced'])
            add(invamt, key, r['invoiced'])
            skey = key + '|' + (r['status'] or '(blank)')
            add(status, skey, r['value'])
            status_count[skey] = status_count.get(skey, 0) + 1
            person = r['person'] or '(unassigned)'
            people_seen[person] = people_seen.get(person, 0) + 1
            pkey = '%s|%s' % (key, person)
            add(pactual, pkey, r['value'])
            pcount[pkey] = pcount.get(pkey, 0) + 1
        else:
            label = '%s / %s' % (region, segment)
            add(excluded, label, r['value'])
            add(excluded_by_month, '%d|%s|%s' % (month, region, segment), r['value'])
            excluded_total += r['value']

    if undated:
        warn('%d register rows carry no booking month and are not in any figure' % undated)
    if wrong_year:
        warn('%d register rows fall outside FY %d and are not in any figure' % (wrong_year, fy))

    # -- targets -----------------------------------------------------------
    target, ptarget, ptarget_month = {}, {}, {}
    target_people = {}
    for t in targets:
        if t['year'] not in (None, fy):
            continue
        region = nreg.get(t['region'], t['region'])
        segment = nseg.get(t['segment'], t['segment'])
        if region not in regions or segment not in segments:
            if any(t['monthly']):
                warn('target row for %s (%s / %s) is outside the reporting grid'
                     % (t['person'], t['region'], t['segment']))
            continue
        target_people[norm_key(t['person'])] = t['person']
        for i, v in enumerate(t['monthly']):
            if not v:
                continue
            key = '%d|%s|%s' % (i + 1, region, segment)
            target[key] = target.get(key, 0.0) + v
            pk = '%s|%s|%s' % (region, segment, t['person'])
            ptarget[pk] = ptarget.get(pk, 0.0) + v / 12.0
            pmk = '%d|%s|%s|%s' % (i + 1, region, segment, t['person'])
            ptarget_month[pmk] = ptarget_month.get(pmk, 0.0) + v

    # -- reconcile the two name registers ----------------------------------
    aliases = {k: v for k, v in cfg['normalisation']['personAliases'].items()
               if not k.startswith('_')}
    register_names = sorted(p for p in people_seen if p != '(unassigned)')
    matched, fuzzy, unmatched_names = match_people(
        register_names, sorted(set(target_people.values())), aliases)
    unmatched = {p: people_seen[p] for p in unmatched_names}
    booked_without_target = {}
    for person in unmatched:
        booked_without_target[person] = sum(
            v for k, v in pactual.items() if k.rsplit('|', 1)[1] == person)

    # people the target register knows but who booked nothing
    claimed = set(matched.values())
    idle = sorted(name for name in set(target_people.values()) if name not in claimed)

    if unmatched:
        warn('%d sales people book against no target (%.2f Cr) -- see reconciliation.txt'
             % (len(unmatched), sum(booked_without_target.values())))
    unassigned = people_seen.get('(unassigned)', 0)
    if unassigned:
        warn('%d orders carry no sales person and sit under "(unassigned)"' % unassigned)

    # a person's targets are keyed by the TARGET register's spelling, so
    # re-key the booking side onto it wherever the two differ
    if matched:
        remap = {p: t for p, t in matched.items() if p != t}
        if remap:
            for store in (pactual, pcount):
                for key in list(store):
                    head, person = key.rsplit('|', 1)
                    if person in remap:
                        nk = head + '|' + remap[person]
                        store[nk] = store.get(nk, 0) + store.pop(key)

    people = sorted(set(list(target_people.values())) |
                    {k.rsplit('|', 1)[1] for k in pactual})

    latest = max([int(k.split('|')[0]) for k in actual] or [1])

    src_stat = os.stat(workbook_path)
    payload = {
        'schema': 'ob-dashboard/2',
        'fiscalYear': fy,
        'generatedAt': dt.datetime.now().astimezone().isoformat(timespec='seconds'),
        'sourceFile': os.path.basename(cfg['_resolved_workbook']),
        'sourceModified': dt.datetime.fromtimestamp(src_stat.st_mtime).astimezone()
                            .isoformat(timespec='seconds'),
        'sourceVia': source_via,
        'sourceSize': src_stat.st_size,
        'months': cfg['months'],
        'regions': regions,
        'segments': segments,
        'latestMonth': latest,
        'actual': actual,
        'target': target,
        'count': count,
        'status': status,
        'statusCount': status_count,
        'allStatuses': all_statuses,
        'openStatuses': cfg['statuses']['open'],
        'invoicedStatuses': cfg['statuses']['invoiced'],
        'soip': soip,
        'invoicedAmount': invamt,
        'people': people,
        'pactual': pactual,
        'pcount': pcount,
        'ptarget': ptarget,
        'ptargetByMonth': ptarget_month,
        'excluded': excluded,
        'excludedTotal': excluded_total,
        'excludedByMonth': excluded_by_month,
        'rowsRead': len(register),
        'targetRowsRead': len(targets),
        'unmatchedPeople': booked_without_target,
        'fuzzyMatchedPeople': fuzzy,
        'peopleWithoutBooking': idle,
        'undatedRows': undated,
        'otherYearRows': wrong_year,
    }
    return payload, {
        'matched': matched, 'fuzzy': fuzzy, 'unmatched': unmatched, 'idle': idle,
        'booked_without_target': booked_without_target,
        'statuses': all_statuses,
    }


# ---------------------------------------------------------------- outputs --

def write_reconciliation(path, payload, detail):
    lines = []
    add = lines.append
    add('Order Booking Stage 2 -- reconciliation')
    add('generated %s' % payload['generatedAt'])
    add('source     %s (saved %s)' % (payload['sourceFile'], payload['sourceModified']))
    add('')
    add('REGISTER')
    add('  rows read from %s: %d' % ('YTD Merg', payload['rowsRead']))
    add('  rows with no booking month: %d' % payload['undatedRows'])
    add('  rows outside FY %d: %d' % (payload['fiscalYear'], payload['otherYearRows']))
    add('  booking inside the %d x %d grid: %.2f Cr'
        % (len(payload['regions']), len(payload['segments']),
           sum(payload['actual'].values())))
    add('  booking outside the grid:      %.2f Cr' % payload['excludedTotal'])
    for label, v in sorted(payload['excluded'].items(), key=lambda kv: -abs(kv[1])):
        add('      %-28s %10.2f Cr' % (label, v))
    add('')
    add('STATUS FLAGS SEEN (column N)')
    for s, n in sorted(detail['statuses'].items(), key=lambda kv: -kv[1]):
        bucket = ('open' if s in payload['openStatuses']
                  else 'invoiced' if s in payload['invoicedStatuses'] else 'NEITHER')
        add('  %-16s %6d rows   counted as: %s' % (s, n, bucket))
    add('')
    add('TARGETS')
    add('  rows read from %s: %d' % ('Monthly Booking Target (2)', payload['targetRowsRead']))
    add('  full-year target inside the grid: %.2f Cr' % sum(payload['target'].values()))
    add('')
    add('SALES PEOPLE')
    add('  names matched between the register and the target sheet: %d' % len(detail['matched']))
    renamed = {k: v for k, v in detail['matched'].items() if k != v}
    if renamed:
        add('  matched across a spelling difference -- check these, and pin any that are')
        add('  wrong with an entry in config.json normalisation.personAliases:')
        for k, v in sorted(renamed.items()):
            how = 'word overlap' if k in detail.get('fuzzy', {}) else 'same words'
            add('      %-32s -> %-28s (%s)' % (k, v, how))
    if detail['unmatched']:
        add('  booking with no target row (their booking still counts in every regional')
        add('  figure; only their personal target column is blank):')
        for k, n in sorted(detail['unmatched'].items(),
                           key=lambda kv: -detail['booked_without_target'].get(kv[0], 0)):
            add('      %-32s %4d orders  %8.2f Cr'
                % (k, n, detail['booked_without_target'].get(k, 0.0)))
    if detail['idle']:
        add('  carrying a target but no booking yet:')
        for k in detail['idle']:
            add('      %s' % k)
    text = '\n'.join(lines) + '\n'
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, 'w', encoding='utf-8') as fh:
        fh.write(text)
    return text


def write_payload(payload, json_path, js_path):
    blob = json.dumps(payload, ensure_ascii=False, separators=(',', ':'), sort_keys=True)
    payload['checksum'] = hashlib.sha256(blob.encode('utf-8')).hexdigest()[:16]
    pretty = json.dumps(payload, ensure_ascii=False, indent=1, sort_keys=True)
    for path in (json_path, js_path):
        if path:
            os.makedirs(os.path.dirname(path), exist_ok=True)
    if json_path:
        with open(json_path, 'w', encoding='utf-8') as fh:
            fh.write(pretty)
    if js_path:
        with open(js_path, 'w', encoding='utf-8') as fh:
            fh.write('/* Offline snapshot written by ob_pipeline.py -- do not edit by hand.\n'
                     '   The dashboard uses this only when the data service cannot be\n'
                     '   reached, and always labels it on screen as a snapshot with its age. */\n')
            fh.write('window.OB_SNAPSHOT = ')
            fh.write(pretty)
            fh.write(';\n')
    return payload


# ------------------------------------------------------------------- main --

def run(config_path, quiet=False):
    cfg = load_config(config_path)
    base = os.path.dirname(os.path.abspath(config_path))
    warnings = []

    def warn(msg):
        warnings.append(msg)
        if not quiet:
            print('  warning: ' + msg)

    workbook = resolve(cfg['workbook'], base)
    cfg['_resolved_workbook'] = workbook
    if not os.path.exists(workbook):
        raise SystemExit('workbook not found: %s' % workbook)

    # Excel holds an exclusive lock while the file is open. Read a copy so a
    # refresh never fails just because someone has the workbook on screen --
    # and say so in the payload, because a fallback nobody is told about is how
    # a pipeline starts lying.
    source_via = 'direct'
    read_from = workbook
    try:
        with open(workbook, 'rb'):
            pass
    except OSError:
        copy = resolve(cfg['workingCopy'], base)
        os.makedirs(os.path.dirname(copy), exist_ok=True)
        shutil.copy2(workbook, copy)
        read_from, source_via = copy, 'working copy'

    payload, detail = build(cfg, read_from, source_via, warn)
    payload['warnings'] = warnings
    payload['freshness'] = cfg['freshness']
    write_payload(payload,
                  resolve(cfg['outputJson'], base) if cfg.get('outputJson') else None,
                  resolve(cfg['output'], base) if cfg.get('output') else None)
    write_reconciliation(resolve(cfg['reconciliation'], base), payload, detail)
    return payload


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--config', default=os.path.join(HERE, 'config.json'))
    ap.add_argument('--quiet', action='store_true')
    args = ap.parse_args()

    payload = run(args.config, quiet=args.quiet)
    if args.quiet:
        return 0
    a = sum(payload['actual'].values())
    t = sum(payload['target'].values())
    ytd_t = sum(v for k, v in payload['target'].items()
                if int(k.split('|')[0]) <= payload['latestMonth'])
    print('read     %d register rows, %d target rows'
          % (payload['rowsRead'], payload['targetRowsRead']))
    print('latest   %s %d' % (payload['months'][payload['latestMonth'] - 1], payload['fiscalYear']))
    print('booking  %.2f Cr in grid, %.2f Cr outside it'
          % (a, payload['excludedTotal']))
    print('target   %.2f Cr YTD, %.2f Cr full year' % (ytd_t, t))
    print('people   %d' % len(payload['people']))
    print('written  data.js + data.json + reconciliation.txt   (checksum %s)'
          % payload['checksum'])
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
