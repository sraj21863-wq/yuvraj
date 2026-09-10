#!/usr/bin/env python3
"""
The Stage 2 data service.

Serves the dashboard and the payload it reads:

  GET /data.json   the payload. Rebuilt from the workbook whenever the workbook
                   has been saved since the last build, so the page is live
                   without anyone having to remember to refresh anything.
  GET /health      what the service can see right now: whether it can reach the
                   workbook, when the workbook was last saved, when the payload
                   was last built, and any warnings the last build raised.
  GET /            the dashboard itself.

A failed rebuild does not take the service down and does not serve a stale
payload as if it were fresh: /data.json returns the last good payload with its
real build time, and /health reports the failure, so the page's own freshness
badge tells the truth on its own.
"""

import json
import mimetypes
import os
import sys
import threading
import traceback
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import ob_pipeline  # noqa: E402

CONFIG = os.path.join(HERE, 'config.json')
DASHBOARD = os.path.normpath(os.path.join(HERE, '..', 'dashboard'))

_lock = threading.Lock()
_state = {'payload': None, 'mtime': None, 'error': None, 'builds': 0}


def workbook_path():
    cfg = ob_pipeline.load_config(CONFIG)
    return ob_pipeline.resolve(cfg['workbook'], HERE)


def current(force=False):
    """Return (payload, error). Rebuild only when the workbook has moved on."""
    path = workbook_path()
    try:
        mtime = os.path.getmtime(path)
    except OSError as exc:
        with _lock:
            _state['error'] = 'workbook not readable: %s' % exc
            return _state['payload'], _state['error']
    with _lock:
        fresh = _state['payload'] is not None and _state['mtime'] == mtime and not force
        if fresh:
            return _state['payload'], None
    try:
        payload = ob_pipeline.run(CONFIG, quiet=True)
        with _lock:
            _state.update(payload=payload, mtime=mtime, error=None,
                          builds=_state['builds'] + 1)
        return payload, None
    except Exception:                                       # noqa: BLE001
        err = traceback.format_exc(limit=3)
        with _lock:
            _state['error'] = err
            return _state['payload'], err


class Handler(BaseHTTPRequestHandler):
    server_version = 'OBStage2/2.0'

    def log_message(self, fmt, *args):
        sys.stderr.write('  %s %s\n' % (self.address_string(), fmt % args))

    def _send(self, code, body, ctype='application/json; charset=utf-8', extra=None):
        if isinstance(body, str):
            body = body.encode('utf-8')
        self.send_response(code)
        self.send_header('Content-Type', ctype)
        self.send_header('Content-Length', str(len(body)))
        self.send_header('Cache-Control', 'no-store')
        self.send_header('Access-Control-Allow-Origin', '*')
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        if self.command != 'HEAD':
            self.wfile.write(body)

    def do_HEAD(self):
        self.do_GET()

    def do_GET(self):
        route = urlparse(self.path).path
        if route in ('/data.json', '/data'):
            payload, err = current()
            if payload is None:
                self._send(503, json.dumps({'error': err or 'no payload has been built yet'}))
                return
            self._send(200, json.dumps(payload, ensure_ascii=False))
            return

        if route == '/health':
            payload, err = current()
            path = workbook_path()
            body = {
                'ok': err is None and payload is not None,
                'workbook': path,
                'workbookReadable': os.path.exists(path),
                'workbookModified': (ob_pipeline.dt.datetime
                                     .fromtimestamp(os.path.getmtime(path)).astimezone()
                                     .isoformat(timespec='seconds')) if os.path.exists(path) else None,
                'payloadBuiltAt': payload.get('generatedAt') if payload else None,
                'checksum': payload.get('checksum') if payload else None,
                'builds': _state['builds'],
                'warnings': payload.get('warnings', []) if payload else [],
                'error': err,
            }
            self._send(200 if body['ok'] else 503, json.dumps(body, indent=1))
            return

        if route == '/refresh':
            payload, err = current(force=True)
            self._send(200 if payload else 503,
                       json.dumps({'rebuilt': err is None, 'error': err,
                                   'checksum': payload.get('checksum') if payload else None}))
            return

        name = 'Order_Booking_Dashboard.html' if route in ('/', '') else route.lstrip('/')
        target = os.path.normpath(os.path.join(DASHBOARD, name))
        if not target.startswith(DASHBOARD) or not os.path.isfile(target):
            self._send(404, json.dumps({'error': 'not found: ' + route}))
            return
        ctype = mimetypes.guess_type(target)[0] or 'application/octet-stream'
        if ctype.startswith('text/') or ctype.endswith('javascript'):
            ctype += '; charset=utf-8'
        with open(target, 'rb') as fh:
            self._send(200, fh.read(), ctype)


def main():
    cfg = ob_pipeline.load_config(CONFIG)
    host = cfg['service'].get('host', '127.0.0.1')
    port = int(cfg['service'].get('port', 8787))
    payload, err = current(force=True)
    if err:
        print('first build failed:\n' + err, file=sys.stderr)
    elif payload:
        print('first build ok: %d register rows, checksum %s'
              % (payload['rowsRead'], payload['checksum']))
    print('Order Booking Stage 2 service')
    print('  dashboard  http://%s:%d/' % (host, port))
    print('  payload    http://%s:%d/data.json' % (host, port))
    print('  health     http://%s:%d/health' % (host, port))
    print('  watching   %s' % workbook_path())
    print('  Ctrl-C to stop.')
    ThreadingHTTPServer((host, port), Handler).serve_forever()


if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        print('\nstopped.')
