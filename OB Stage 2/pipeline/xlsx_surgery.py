"""
Low-level OOXML helpers for rebuilding a workbook's data while leaving its
automation (formulas, pivot definitions, external links, styles, conditional
formatting, tables, defined names) intact.

Everything here works on the raw parts inside the .xlsx zip. openpyxl is not
used on purpose: it silently drops web extensions, rich-data parts, custom
properties and some pivot metadata, and it rewrites styles.xml -- all of which
would change the "automated" workbook in ways nobody asked for.
"""

import re
import zipfile
from xml.etree import ElementTree as ET

NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
NS = '{%s}' % NS_MAIN
MAX_ROW = 1048576

ET.register_namespace('', NS_MAIN)


# --------------------------------------------------------------------------
# workbook part discovery
# --------------------------------------------------------------------------

def sheet_map(z):
    """{sheet name: (part name, state)} in workbook order."""
    wb = z.read('xl/workbook.xml').decode('utf-8')
    rels = z.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    target = {m.group(1): m.group(2) for m in
              re.finditer(r'<Relationship Id="([^"]+)"[^>]*Target="([^"]+)"', rels)}
    out = {}
    for m in re.finditer(
            r'<sheet name="([^"]+)" sheetId="\d+"(?: state="(\w+)")? r:id="([^"]+)"', wb):
        t = target[m.group(3)].lstrip('/')
        out[m.group(1)] = ('xl/' + t if not t.startswith('xl/') else t, m.group(2))
    return out


def external_link_order(z):
    """[target of [1], target of [2], ...] -- the order Excel's [n] indexes use."""
    wb = z.read('xl/workbook.xml').decode('utf-8')
    rels = z.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    target = {m.group(1): m.group(2) for m in
              re.finditer(r'<Relationship Id="([^"]+)"[^>]*Target="([^"]+)"', rels)}
    block = re.search(r'<externalReferences>(.*?)</externalReferences>', wb, re.S)
    if not block:
        return []
    out = []
    for m in re.finditer(r'r:id="([^"]+)"', block.group(1)):
        part = 'xl/' + target[m.group(1)].lstrip('/')
        lrels = part.rsplit('/', 1)[0] + '/_rels/' + part.rsplit('/', 1)[1] + '.rels'
        tgt = ''
        if lrels in z.namelist():
            mm = re.search(r'Target="([^"]+)"', z.read(lrels).decode('utf-8'))
            if mm:
                tgt = mm.group(1)
        out.append(tgt)
    return out


def link_key(target):
    """Normalise an external-link target so the same workbook matches across files."""
    t = target.replace('\\', '/').split('/')[-1]
    return t.lower()


# --------------------------------------------------------------------------
# shared strings
# --------------------------------------------------------------------------

def read_shared_strings(z):
    """Returns the list of <si> inner XML fragments (rich text preserved)."""
    if 'xl/sharedStrings.xml' not in z.namelist():
        return []
    raw = z.read('xl/sharedStrings.xml').decode('utf-8')
    return re.findall(r'<si>(.*?)</si>', raw, re.S)


def build_shared_strings(items):
    body = ''.join('<si>%s</si>' % s for s in items)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            '<sst xmlns="%s" count="%d" uniqueCount="%d">%s</sst>'
            % (NS_MAIN, len(items), len(items), body)).encode('utf-8')


class StringPool:
    """Automated workbook's strings, kept at their original indices; new ones append."""

    def __init__(self, base_items):
        self.items = list(base_items)
        self.index = {}
        for i, s in enumerate(self.items):
            self.index.setdefault(s, i)

    def intern(self, si_xml):
        i = self.index.get(si_xml)
        if i is None:
            i = len(self.items)
            self.items.append(si_xml)
            self.index[si_xml] = i
        return i


# --------------------------------------------------------------------------
# styles merge
# --------------------------------------------------------------------------

def _collection(raw, tag):
    """Return (prefix, [child xml, ...], suffix) for a <tag count=..> block."""
    m = re.search(r'<%s(?:\s+count="\d+")?\s*>(.*?)</%s>' % (tag, tag), raw, re.S)
    if not m:
        m2 = re.search(r'<%s(?:\s+count="\d+")?\s*/>' % tag, raw)
        if m2:
            return m2.span(), []
        return None, []
    inner = m.group(1)
    depth, start, out = 0, None, []
    i = 0
    while i < len(inner):
        if inner[i] == '<':
            j = inner.index('>', i)
            token = inner[i:j + 1]
            if token.startswith('</'):
                depth -= 1
                if depth == 0:
                    out.append(inner[start:j + 1])
                    start = None
            elif token.endswith('/>'):
                if depth == 0:
                    out.append(token)
            else:
                if depth == 0:
                    start = i
                depth += 1
            i = j + 1
        else:
            i += 1
    return m.span(), out


def merge_styles(base_raw, add_raw):
    """
    Append the donor workbook's formatting records to the base workbook's
    styles.xml so donor cells keep their own look. Returns
    (new styles.xml text, {donor xf id: base xf id}).
    """
    base = base_raw
    out_map = {}

    def counts(raw, tag):
        span, items = _collection(raw, tag)
        return span, items

    # -- numFmts (custom ids >= 164 may collide, so re-id the donor's) -------
    b_span, b_numfmts = counts(base, 'numFmts')
    _, a_numfmts = counts(add_raw, 'numFmts')
    b_codes = {}
    for x in b_numfmts:
        m = re.search(r'numFmtId="(\d+)"\s+formatCode="(.*?)"\s*/>', x, re.S)
        if m:
            b_codes[m.group(2)] = int(m.group(1))
    next_id = max(list(b_codes.values()) + [163]) + 1
    numfmt_map = {}
    added_numfmts = []
    for x in a_numfmts:
        m = re.search(r'numFmtId="(\d+)"\s+formatCode="(.*?)"\s*/>', x, re.S)
        if not m:
            continue
        old, code = int(m.group(1)), m.group(2)
        if code in b_codes:
            numfmt_map[old] = b_codes[code]
        else:
            numfmt_map[old] = next_id
            b_codes[code] = next_id
            added_numfmts.append('<numFmt numFmtId="%d" formatCode="%s"/>' % (next_id, code))
            next_id += 1

    def extend(raw, tag, extra, self_closing_ok=True):
        """Append `extra` children to <tag>, returning (new raw, offset for donor ids)."""
        span, items = _collection(raw, tag)
        offset = len(items)
        if not extra:
            return raw, offset
        new_items = items + extra
        block = '<%s count="%d">%s</%s>' % (tag, len(new_items), ''.join(new_items), tag)
        if span is None:
            return raw, offset
        return raw[:span[0]] + block + raw[span[1]:], offset

    if added_numfmts:
        span, items = _collection(base, 'numFmts')
        new_items = items + added_numfmts
        block = '<numFmts count="%d">%s</numFmts>' % (len(new_items), ''.join(new_items))
        if span is None:
            base = base.replace('<styleSheet', '<styleSheet', 1)
            insert_at = base.index('>', base.index('<styleSheet')) + 1
            base = base[:insert_at] + block + base[insert_at:]
        else:
            base = base[:span[0]] + block + base[span[1]:]

    # -- fonts / fills / borders -------------------------------------------
    offsets = {}
    for tag in ('fonts', 'fills', 'borders'):
        _, donor = _collection(add_raw, tag)
        base, off = extend(base, tag, donor)
        offsets[tag] = off

    # -- cellStyleXfs is referenced by cellXfs' xfId ------------------------
    _, donor_csx = _collection(add_raw, 'cellStyleXfs')
    base, csx_off = extend(base, 'cellStyleXfs', donor_csx)

    # -- cellXfs ------------------------------------------------------------
    _, base_xfs = _collection(base, 'cellXfs')
    _, donor_xfs = _collection(add_raw, 'cellXfs')
    xf_off = len(base_xfs)
    remapped = []
    for x in donor_xfs:
        def fix(attr, off):
            def sub(m):
                return '%s="%d"' % (attr, int(m.group(1)) + off)
            return re.sub(r'%s="(\d+)"' % attr, sub, x)

        y = x
        for attr, tag in (('fontId', 'fonts'), ('fillId', 'fills'), ('borderId', 'borders')):
            o = offsets[tag]
            y = re.sub(r'%s="(\d+)"' % attr,
                       lambda m, a=attr, o=o: '%s="%d"' % (a, int(m.group(1)) + o), y)
        y = re.sub(r'xfId="(\d+)"',
                   lambda m: 'xfId="%d"' % (int(m.group(1)) + csx_off), y)
        y = re.sub(r'numFmtId="(\d+)"',
                   lambda m: 'numFmtId="%d"' % numfmt_map.get(int(m.group(1)), int(m.group(1))), y)
        remapped.append(y)
    base, _ = extend(base, 'cellXfs', remapped)
    for i in range(len(donor_xfs)):
        out_map[i] = xf_off + i
    return base, out_map


# --------------------------------------------------------------------------
# formula rewriting
# --------------------------------------------------------------------------

_QUAL = r"(?:(?:'(?:[^']|'')+'|\[\d+\][A-Za-z0-9_.\- ]+|[A-Za-z_][A-Za-z0-9_.]*)!)"
_REF = re.compile(
    r"(?<![A-Za-z0-9_.$!])(%s?)(\$?)([A-Z]{1,3})(\$?)([0-9]{1,7})(?![0-9(A-Za-z_])" % _QUAL)
_STR = re.compile(r'"(?:[^"]|"")*"')


def _map_outside_strings(formula, fn):
    out, last = [], 0
    for m in _STR.finditer(formula):
        out.append(fn(formula[last:m.start()]))
        out.append(m.group(0))
        last = m.end()
    out.append(fn(formula[last:]))
    return ''.join(out)


def shift_rows(formula, delta, self_only=False, self_names=()):
    """
    Add `delta` to the row of every relative A1 reference.

    self_only=True restricts the shift to references that live on the formula's
    own sheet -- unqualified refs, or refs qualified with one of `self_names`.
    That is what re-anchoring a live block needs: the block's self-references
    follow it down the sheet while its references into the raw-paste sheet stay
    pinned where they are.
    """
    lowered = tuple(n.lower() for n in self_names)

    def one(chunk):
        def sub(m):
            qual, dollar_c, col, dollar_r, row = m.groups()
            if dollar_r:
                return m.group(0)
            if self_only:
                if qual:
                    name = qual[:-1].strip()
                    if name.startswith("'") and name.endswith("'"):
                        name = name[1:-1].replace("''", "'")
                    if name.lower() not in lowered:
                        return m.group(0)
            r = int(row) + delta
            if r < 1:
                r = 1
            if r > MAX_ROW:
                r = MAX_ROW
            return '%s%s%s%s%d' % (qual, dollar_c, col, dollar_r, r)
        return _REF.sub(sub, chunk)

    return _map_outside_strings(formula, one)


def remap_external(formula, mapping):
    """Rewrite [n] external-workbook indexes, e.g. donor [1] -> base [3]."""
    if not mapping:
        return formula

    def one(chunk):
        return re.sub(r'\[(\d+)\]',
                      lambda m: '[%d]' % mapping.get(int(m.group(1)), int(m.group(1))), chunk)
    return _map_outside_strings(formula, one)


# --------------------------------------------------------------------------
# sheet reading / writing
# --------------------------------------------------------------------------

COL_RE = re.compile(r'([A-Z]+)')


def col_of(ref):
    return COL_RE.match(ref).group(1)


def col_num(col):
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - 64)
    return n


def col_name(n):
    s = ''
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


class Cell:
    __slots__ = ('col', 's', 't', 'f', 'fattrs', 'v', 'is_xml', 'cm', 'vm', 'ph')

    def __init__(self, col, s=None, t=None, f=None, fattrs=None, v=None, is_xml=None):
        self.col, self.s, self.t = col, s, t
        self.f, self.fattrs, self.v, self.is_xml = f, fattrs or {}, v, is_xml
        self.cm = self.vm = self.ph = None


class Row:
    __slots__ = ('n', 'attrs', 'cells')

    def __init__(self, n, attrs, cells):
        self.n, self.attrs, self.cells = n, attrs, cells


def root_prefixes(z, part):
    """{namespace uri: prefix} from a part's root element, so attribute names
    ElementTree hands back expanded ({uri}local) can be written back as
    prefix:local -- otherwise the emitted XML is not well formed."""
    with z.open(part) as fh:
        head = fh.read(8192).decode('utf-8', errors='replace')
    m = re.search(r'<(?:\w+:)?worksheet\b[^>]*>', head, re.S)
    scope = m.group(0) if m else head
    return {uri: pfx for pfx, uri in re.findall(r'xmlns:([A-Za-z0-9_.\-]+)="([^"]+)"', scope)}


def qualify(attrs, prefixes):
    """Turn {'{uri}local': v} back into {'pfx:local': v}; drop undeclared ones."""
    out = {}
    for k, v in attrs.items():
        if k.startswith('{'):
            uri, local = k[1:].split('}', 1)
            pfx = prefixes.get(uri)
            if pfx is None:
                continue
            k = '%s:%s' % (pfx, local)
        out[k] = v
    return out


def read_sheet_rows(z, part, expand_shared=True, self_sheet_names=()):
    """Yield Row objects. Shared formulas are expanded into ordinary ones."""
    masters = {}
    prefixes = root_prefixes(z, part)
    with z.open(part) as fh:
        for _, el in ET.iterparse(fh, events=('end',)):
            if el.tag != NS + 'row':
                continue
            n = int(el.get('r'))
            attrs = qualify({k: v for k, v in el.attrib.items() if k != 'r'}, prefixes)
            cells = {}
            for c in el.findall(NS + 'c'):
                ref = c.get('r')
                col = col_of(ref)
                cell = Cell(col, c.get('s'), c.get('t'))
                for extra in ('cm', 'vm', 'ph'):
                    if c.get(extra) is not None:
                        setattr(cell, extra, c.get(extra))
                fe = c.find(NS + 'f')
                if fe is not None:
                    fattrs = qualify(dict(fe.attrib), prefixes)
                    text = fe.text or ''
                    if expand_shared and fattrs.get('t') == 'shared':
                        si = fattrs.get('si')
                        if text:
                            masters[si] = (n, text)
                            cell.f = text
                        else:
                            base_row, base_text = masters.get(si, (n, ''))
                            cell.f = shift_rows(base_text, n - base_row) if base_text else ''
                        fattrs.pop('t', None)
                        fattrs.pop('si', None)
                        fattrs.pop('ref', None)
                    else:
                        cell.f = text
                    cell.fattrs = fattrs
                ve = c.find(NS + 'v')
                if ve is not None:
                    cell.v = ve.text
                ise = c.find(NS + 'is')
                if ise is not None:
                    cell.is_xml = ''.join(ET.tostring(k, encoding='unicode')
                                          for k in list(ise)).replace(NS, '')
                cells[col] = cell
            yield Row(n, attrs, cells)
            el.clear()


def esc(s):
    return (s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;'))


def cell_xml(row_n, cell):
    ref = '%s%d' % (cell.col, row_n)
    a = ' r="%s"' % ref
    if cell.s is not None:
        a += ' s="%s"' % cell.s
    if cell.t:
        a += ' t="%s"' % cell.t
    for extra in ('cm', 'vm', 'ph'):
        val = getattr(cell, extra, None)
        if val is not None:
            a += ' %s="%s"' % (extra, val)
    body = ''
    if cell.f is not None:
        fa = ''.join(' %s="%s"' % (k, esc(v)) for k, v in cell.fattrs.items())
        body += '<f%s>%s</f>' % (fa, esc(cell.f)) if cell.f else '<f%s/>' % fa
    if cell.v is not None:
        body += '<v>%s</v>' % esc(cell.v)
    if cell.is_xml is not None:
        body += '<is>%s</is>' % cell.is_xml
    if not body:
        return '<c%s/>' % a
    return '<c%s>%s</c>' % (a, body)


def row_xml(row):
    cols = sorted(row.cells.values(), key=lambda c: col_num(c.col))
    a = ' r="%d"' % row.n
    if cols:
        a += ' spans="%d:%d"' % (col_num(cols[0].col), col_num(cols[-1].col))
    for k, v in row.attrs.items():
        if k == 'spans':
            continue
        a += ' %s="%s"' % (k, esc(v))
    if not cols:
        return '<row%s/>' % a
    return '<row%s>%s</row>' % (a, ''.join(cell_xml(row.n, c) for c in cols))


SHEETDATA_RE = re.compile(r'<sheetData\s*/>|<sheetData>.*?</sheetData>', re.S)


def replace_sheet_data(sheet_xml, rows_iter):
    parts = ['<sheetData>']
    for r in rows_iter:
        parts.append(row_xml(r))
    parts.append('</sheetData>')
    body = ''.join(parts)
    return SHEETDATA_RE.sub(lambda _: body, sheet_xml, count=1)


def set_dimension(sheet_xml, ref):
    if '<dimension' in sheet_xml:
        return re.sub(r'<dimension ref="[^"]*"/>', '<dimension ref="%s"/>' % ref,
                      sheet_xml, count=1)
    return sheet_xml
