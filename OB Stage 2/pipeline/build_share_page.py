#!/usr/bin/env python3
"""
Fold the dashboard and its three scripts into one file you can mail to someone.

The everyday dashboard is four files that a local web server ties together. That
is right for a live page and wrong for sending to a colleague, who will save one
attachment, double-click it, and get a broken page if the scripts are missing.

This writes a single self-contained HTML file with the config, the validator and
the current snapshot inlined. It deliberately clears the service endpoint: there
is no service on the recipient's machine, so the page goes straight to the
snapshot instead of waiting eight seconds to find that out. It renders exactly
as the live page does, and says SNAPSHOT with the data's age, because that is
what it is.

Run refresh.cmd first if you want the snapshot to be current.
"""

import argparse
import os
import re

HERE = os.path.dirname(os.path.abspath(__file__))
DASH = os.path.normpath(os.path.join(HERE, '..', 'dashboard'))


def read(name):
    with open(os.path.join(DASH, name), encoding='utf-8') as fh:
        return fh.read()


def build(strip_document_shell=False):
    html = read('Order_Booking_Dashboard.html')
    parts = {
        'dashboard-config.js': read('dashboard-config.js'),
        'validator.js': read('validator.js'),
        'data.js': read('data.js'),
    }
    # the shared copy has no service to talk to, so don't make it wait for one
    parts['dashboard-config.js'] = re.sub(
        r"endpoint:\s*'[^']*'", "endpoint: ''", parts['dashboard-config.js'], count=1)

    for name, body in parts.items():
        tag = '<script src="%s"></script>' % name
        if tag not in html:
            raise SystemExit('could not find %s in the dashboard' % tag)
        html = html.replace(tag, '<script>\n/* ---- %s ---- */\n%s\n</script>' % (name, body), 1)

    if strip_document_shell:
        head = re.search(r'<head[^>]*>(.*?)</head>', html, re.S).group(1)
        body = re.search(r'<body[^>]*>(.*?)</body>', html, re.S).group(1)
        keep = '\n'.join(
            line for line in head.splitlines()
            if not re.match(r'\s*<meta\s+(charset|name="viewport")', line))
        html = keep.strip() + '\n' + body.strip() + '\n'
    return html


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--out', default=os.path.join(DASH, 'Order_Booking_Dashboard_standalone.html'))
    ap.add_argument('--fragment', action='store_true',
                    help='emit head+body content only, for hosts that supply the document shell')
    args = ap.parse_args()
    html = build(strip_document_shell=args.fragment)
    with open(args.out, 'w', encoding='utf-8') as fh:
        fh.write(html)
    print('written %s (%.0f KB)' % (args.out, os.path.getsize(args.out) / 1024))


if __name__ == '__main__':
    main()
