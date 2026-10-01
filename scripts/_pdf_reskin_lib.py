#!/usr/bin/env python3
"""Shared builder for reskinning an instrument whose source is a PDF.

A PDF has no reusable Word XML, so the instrument's content is transcribed
(from a PyMuPDF text dump, verbatim) into a small block list in the per-document
script, and this module pours it into an unpacked CoPP template package.

Inline markup inside any text string:
    **bold**   //italic//   __underline__   [text](https://url | mailto:x | #bookmark)
    {fn:N}     footnote N (text supplied in the document's FOOTNOTES dict)
    \n         line break inside the same paragraph

Block types (tuples):
    ('h1', text[, bookmark])  ('h2', text)
    ('p', text)               ('pk', text)   paragraph kept with the next one
    ('b', [items])            bullet list (template list, numId 15)
    ('alpha', [items])        (a), (b), (c) list
    ('table', spec)           see table()
    ('callout', title, [items])
    ('img', path, alt_title, alt_descr, width_twips)
"""
import copy, os, re, shutil
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
W14 = 'http://schemas.microsoft.com/office/word/2010/wordml'
RELNS = 'http://schemas.openxmlformats.org/package/2006/relationships'
XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'
NS = {'w': W, 'r': R, 'wp': WP, 'a': A, 'pic': PIC}
ns = {'w': W}

HEADER_FILL = 'F68B1F'          # Process template orange
CALLOUT_FILL = 'FEF1E4'         # light tint of the same orange
BORDER = 'D0D0D0'               # template table border grey
BULLET_NUMID = '15'             # the template's own body bullet list
TEXT_W = 9026                   # A4 portrait text width at the template margins


def q(tag):
    p, local = tag.split(':')
    return f'{{{NS[p]}}}{local}'


def E(tag, attrib=None, *children, text=None):
    el = etree.Element(q(tag))
    for k, v in (attrib or {}).items():
        el.set(q(k), str(v))
    for c in children:
        el.append(c)
    if text is not None:
        el.text = text
    return el


def wt(text):
    el = E('w:t', text=text)
    el.set(XML_SPACE, 'preserve')
    return el


RPR_ORDER = ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike',
             'dstrike', 'outline', 'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid',
             'vanish', 'webHidden', 'color', 'spacing', 'w', 'kern', 'position', 'sz',
             'szCs', 'highlight', 'u', 'effect', 'bdr', 'shd', 'fitText', 'vertAlign']
PPR_ORDER = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr',
             'widowControl', 'numPr', 'suppressLineNumbers', 'pBdr', 'shd', 'tabs',
             'suppressAutoHyphens', 'kinsoku', 'wordWrap', 'overflowPunct',
             'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd',
             'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents',
             'suppressOverlap', 'jc', 'textDirection', 'textAlignment',
             'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr']


def make_rpr(props):
    """props: dict of rPr child name -> attrib dict (or None for a bare flag)."""
    rpr = E('w:rPr')
    for name in RPR_ORDER:
        if name in props:
            rpr.append(E('w:' + name, props[name] or {}))
    return rpr


def make_ppr(props):
    ppr = E('w:pPr')
    for name in PPR_ORDER:
        if name in props:
            v = props[name]
            ppr.append(v if isinstance(v, etree._Element) else E('w:' + name, v or {}))
    return ppr


TOKEN = re.compile(r'\[([^\]]+)\]\(([^)\s]+)\)|\{fn:(\d+)\}|\*\*|//|__')


class Builder:
    def __init__(self, tpl_dir, build_dir, footnotes=None):
        shutil.rmtree(build_dir, ignore_errors=True)
        shutil.copytree(tpl_dir, build_dir)
        self.build = build_dir
        self.doc = etree.parse(os.path.join(build_dir, 'word', 'document.xml'))
        self.body = self.doc.getroot().find('w:body', ns)
        self.tblocks = list(self.body)
        self.rels_path = os.path.join(build_dir, 'word', '_rels', 'document.xml.rels')
        self.rels = etree.parse(self.rels_path)
        self.fn_path = os.path.join(build_dir, 'word', 'footnotes.xml')
        self.fn_doc = etree.parse(self.fn_path)
        self.num_path = os.path.join(build_dir, 'word', 'numbering.xml')
        self.num_doc = etree.parse(self.num_path)
        self.footnote_text = footnotes or {}
        self.footnote_ids = {}
        self.next_rid = 300
        self.url_rids = {}
        self.bm_seq = 0
        self.docpr_id = 5000
        self.toc = []                       # (title, bookmark)
        self.alpha_numid = None

    # ------------------------------------------------------------ package parts
    def add_rel(self, rtype, target, external=False):
        rid = f'rId{self.next_rid}'
        self.next_rid += 1
        el = etree.SubElement(self.rels.getroot(), f'{{{RELNS}}}Relationship')
        el.set('Id', rid)
        el.set('Type', f'http://schemas.openxmlformats.org/officeDocument/2006/relationships/{rtype}')
        el.set('Target', target)
        if external:
            el.set('TargetMode', 'External')
        return rid

    def url_rid(self, url):
        if url not in self.url_rids:
            self.url_rids[url] = self.add_rel('hyperlink', url, external=True)
        return self.url_rids[url]

    def footnote_ref(self, n):
        if n not in self.footnote_ids:
            fid = str(len(self.footnote_ids) + 2)        # template reserves -1, 0, 1
            self.footnote_ids[n] = fid
            fn = E('w:footnote', {'w:id': fid})
            p = E('w:p')
            p.append(make_ppr({'pStyle': {'w:val': 'FootnoteText'}}))
            r1 = E('w:r', None, make_rpr({'rStyle': {'w:val': 'FootnoteReference'}}), E('w:footnoteRef'))
            p.append(r1)
            p.append(E('w:r', None, wt(' ')))
            for r in self.runs(self.footnote_text[n], {}):
                p.append(r)
            fn.append(p)
            self.fn_doc.getroot().append(fn)
        return E('w:r', None, make_rpr({'rStyle': {'w:val': 'FootnoteReference'}}),
                 E('w:footnoteReference', {'w:id': self.footnote_ids[n]}))

    def ensure_alpha_list(self):
        """(a), (b), (c) list - the template has none, so add one."""
        if self.alpha_numid:
            return self.alpha_numid
        root = self.num_doc.getroot()
        abs_ids = [int(a.get(q('w:abstractNumId'))) for a in root.findall('w:abstractNum', ns)]
        num_ids = [int(n.get(q('w:numId'))) for n in root.findall('w:num', ns)]
        aid, nid = max(abs_ids) + 1, max(num_ids) + 1
        absn = E('w:abstractNum', {'w:abstractNumId': aid},
                 E('w:multiLevelType', {'w:val': 'singleLevel'}))
        lvl = E('w:lvl', {'w:ilvl': 0},
                E('w:start', {'w:val': 1}), E('w:numFmt', {'w:val': 'lowerLetter'}),
                E('w:lvlText', {'w:val': '(%1)'}), E('w:lvlJc', {'w:val': 'left'}),
                E('w:pPr', None, E('w:ind', {'w:left': 720, 'w:hanging': 436})))
        absn.append(lvl)
        first_num = root.find('w:num', ns)
        first_num.addprevious(absn)          # abstractNums must precede nums
        last_num = root.findall('w:num', ns)[-1]   # and nums precede numIdMacAtCleanup
        last_num.addnext(E('w:num', {'w:numId': nid}, E('w:abstractNumId', {'w:val': aid})))
        self.alpha_numid = str(nid)
        return self.alpha_numid

    # ------------------------------------------------------------ inline
    def runs(self, text, base):
        """Parse inline markup into a list of w:r / w:hyperlink elements."""
        out = []
        state = {'b': False, 'i': False, 'u': False}

        def emit(seg, link=None):
            if not seg:
                return
            props = dict(base)
            if state['b']:
                props['b'] = None; props['bCs'] = None
            if state['i']:
                props['i'] = None; props['iCs'] = None
            if state['u']:
                props['u'] = {'w:val': 'single'}
            if link:
                props['rStyle'] = {'w:val': 'Hyperlink'}
            parts = seg.split('\n')
            rs = []
            for k, part in enumerate(parts):
                r = E('w:r')
                if props:
                    r.append(make_rpr(props))
                if k:
                    r.append(E('w:br'))
                if part:
                    r.append(wt(part))
                rs.append(r)
            if link:
                if link.startswith('#'):
                    hl = E('w:hyperlink', {'w:anchor': link[1:], 'w:history': 1})
                else:
                    hl = E('w:hyperlink', {'r:id': self.url_rid(link), 'w:history': 1})
                for r in rs:
                    hl.append(r)
                out.append(hl)
            else:
                out.extend(rs)

        pos = 0
        for m in TOKEN.finditer(text):
            emit(text[pos:m.start()])
            tok = m.group(0)
            if m.group(1) is not None:
                emit(m.group(1), link=m.group(2))
            elif m.group(3) is not None:
                out.append(self.footnote_ref(m.group(3)))
            elif tok == '**':
                state['b'] = not state['b']
            elif tok == '//':
                state['i'] = not state['i']
            elif tok == '__':
                state['u'] = not state['u']
            pos = m.end()
        emit(text[pos:])
        assert not any(state.values()), f'unbalanced markup in: {text!r}'
        return out

    # ------------------------------------------------------------ blocks
    def para(self, text, ppr=None, rpr=None):
        p = E('w:p')
        p.append(make_ppr(ppr or {}))
        for r in self.runs(text, rpr or {}):
            p.append(r)
        return p

    def heading(self, text, level, bookmark=None):
        style = {1: 'Heading1', 2: 'Heading2'}[level]
        p = E('w:p')
        p.append(make_ppr({'pStyle': {'w:val': style}}))
        if level == 1:
            self.bm_seq += 1
            bm = bookmark or f'_Toc9100{self.bm_seq:04d}'
            bid = str(9100 + self.bm_seq)
            p.append(E('w:bookmarkStart', {'w:id': bid, 'w:name': bm}))
            for r in self.runs(text, {}):
                p.append(r)
            p.append(E('w:bookmarkEnd', {'w:id': bid}))
            self.toc.append((text, bm))
        else:
            for r in self.runs(text, {}):
                p.append(r)
        return p

    def bullet(self, text, in_table=False, numid=None, rpr=None, ilvl=0):
        numpr = E('w:numPr', None, E('w:ilvl', {'w:val': ilvl}),
                  E('w:numId', {'w:val': numid or BULLET_NUMID}))
        props = {'pStyle': {'w:val': 'ListParagraph'}, 'numPr': numpr, 'jc': {'w:val': 'left'}}
        if in_table:
            props['spacing'] = {'w:before': 20, 'w:after': 40, 'w:line': 240, 'w:lineRule': 'auto'}
            props['ind'] = {'w:left': 284, 'w:hanging': 284}
            props['contextualSpacing'] = {'w:val': 0}
        base = dict(rpr or {})
        if in_table:
            base.setdefault('sz', {'w:val': 20}); base.setdefault('szCs', {'w:val': 20})
        p = E('w:p')
        p.append(make_ppr(props))
        for r in self.runs(text, base):
            p.append(r)
        return p

    def cell_para(self, text, jc=None, rpr=None, keep=False):
        props = {'spacing': {'w:before': 40, 'w:after': 40, 'w:line': 252, 'w:lineRule': 'auto'}}
        if jc:
            props['jc'] = {'w:val': jc}
        if keep:
            props['keepNext'] = None
        base = {'sz': {'w:val': 20}, 'szCs': {'w:val': 20}}
        base.update(rpr or {})
        return self.para(text, props, base)

    def _cell(self, content, width, header=False, fill=None, jc=None, cell_rpr=None):
        tc = E('w:tc')
        tcpr = E('w:tcPr', None, E('w:tcW', {'w:w': width, 'w:type': 'dxa'}))
        if header or fill:
            tcpr.append(E('w:shd', {'w:val': 'clear', 'w:color': 'auto',
                                    'w:fill': HEADER_FILL if header else fill}))
        tc.append(tcpr)
        items = content if isinstance(content, list) else [content]
        if not items:
            items = ['']
        for it in items:
            if isinstance(it, str):
                rpr = dict(cell_rpr or {})
                if header:
                    rpr.update({'b': None, 'bCs': None, 'color': {'w:val': 'FFFFFF'}})
                tc.append(self.cell_para(it, jc=jc, rpr=rpr, keep=header))
            elif it[0] == 'b':
                for b in it[1]:
                    tc.append(self.bullet(b, in_table=True, rpr=cell_rpr))
            else:
                raise ValueError(it)
        return tc

    def table(self, spec):
        """spec: cols=[widths], header=[...] (optional), rows=[[cell,...],...],
        jc={col: 'center'}. A cell is a string or a list of strings / ('b', [items])."""
        cols = spec['cols']
        assert sum(cols) == TEXT_W, (sum(cols), cols)
        tbl = E('w:tbl')
        borders = E('w:tblBorders')
        for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            borders.append(E('w:' + edge, {'w:val': 'single', 'w:sz': 4, 'w:space': 0, 'w:color': BORDER}))
        tbl.append(E('w:tblPr', None,
                     E('w:tblStyle', {'w:val': 'TableGrid'}),
                     E('w:tblW', {'w:w': TEXT_W, 'w:type': 'dxa'}),
                     borders,
                     E('w:tblLayout', {'w:type': 'fixed'}),
                     E('w:tblLook', {'w:val': '04A0', 'w:firstRow': 1, 'w:lastRow': 0,
                                     'w:firstColumn': 1, 'w:lastColumn': 0,
                                     'w:noHBand': 0, 'w:noVBand': 1})))
        grid = E('w:tblGrid')
        for c in cols:
            grid.append(E('w:gridCol', {'w:w': c}))
        tbl.append(grid)
        jcs = spec.get('jc', {})
        if spec.get('header'):
            tr = E('w:tr', None, E('w:trPr', None, E('w:cantSplit'), E('w:tblHeader')))
            for ci, h in enumerate(spec['header']):
                tr.append(self._cell(h, cols[ci], header=True))
            tbl.append(tr)
        for row in spec['rows']:
            tr = E('w:tr')
            if spec.get('cant_split', True):
                tr.append(E('w:trPr', None, E('w:cantSplit')))
            for ci, c in enumerate(row):
                tr.append(self._cell(c, cols[ci], jc=jcs.get(ci)))
            tbl.append(tr)
        return tbl

    def callout(self, title, items):
        """Examples box: one-cell table in the template orange."""
        tbl = E('w:tbl')
        borders = E('w:tblBorders')
        for edge in ('top', 'left', 'bottom', 'right'):
            borders.append(E('w:' + edge, {'w:val': 'single', 'w:sz': 8, 'w:space': 0, 'w:color': HEADER_FILL}))
        tbl.append(E('w:tblPr', None,
                     E('w:tblW', {'w:w': TEXT_W, 'w:type': 'dxa'}),
                     borders,
                     E('w:tblLayout', {'w:type': 'fixed'}),
                     E('w:tblCellMar', None,
                       E('w:top', {'w:w': 80, 'w:type': 'dxa'}), E('w:left', {'w:w': 140, 'w:type': 'dxa'}),
                       E('w:bottom', {'w:w': 80, 'w:type': 'dxa'}), E('w:right', {'w:w': 140, 'w:type': 'dxa'})),
                     E('w:tblLook', {'w:val': '0000'})))
        tbl.append(E('w:tblGrid', None, E('w:gridCol', {'w:w': TEXT_W})))
        tr = E('w:tr', None, E('w:trPr', None, E('w:cantSplit')))
        tc = E('w:tc', None, E('w:tcPr', None, E('w:tcW', {'w:w': TEXT_W, 'w:type': 'dxa'}),
                                E('w:shd', {'w:val': 'clear', 'w:color': 'auto', 'w:fill': CALLOUT_FILL})))
        tc.append(self.cell_para(f'**//{title}//**', keep=True))
        for it in items:
            tc.append(self.bullet(f'//{it}//', in_table=True))
        tr.append(tc)
        tbl.append(tr)
        return tbl

    def image(self, path, title, descr, width_twips):
        from PIL import Image
        name = os.path.basename(path)
        shutil.copy(path, os.path.join(self.build, 'word', 'media', name))
        rid = self.add_rel('image', f'media/{name}')
        w_px, h_px = Image.open(path).size
        cx = width_twips * 635                       # twips -> EMU
        cy = round(cx * h_px / w_px)
        self.docpr_id += 1
        inline = E('wp:inline', {})
        for k in ('distT', 'distB', 'distL', 'distR'):
            inline.set(k, '0')
        ext = etree.SubElement(inline, q('wp:extent')); ext.set('cx', str(cx)); ext.set('cy', str(cy))
        ee = etree.SubElement(inline, q('wp:effectExtent'))
        for k in ('l', 't', 'r', 'b'):
            ee.set(k, '0')
        dp = etree.SubElement(inline, q('wp:docPr'))
        dp.set('id', str(self.docpr_id)); dp.set('name', title); dp.set('title', title); dp.set('descr', descr)
        cnv = etree.SubElement(inline, q('wp:cNvGraphicFramePr'))
        lk = etree.SubElement(cnv, q('a:graphicFrameLocks')); lk.set('noChangeAspect', '1')
        g = etree.SubElement(inline, q('a:graphic'))
        gd = etree.SubElement(g, q('a:graphicData')); gd.set('uri', PIC)
        pic = etree.SubElement(gd, q('pic:pic'))
        nv = etree.SubElement(pic, q('pic:nvPicPr'))
        c = etree.SubElement(nv, q('pic:cNvPr')); c.set('id', '0'); c.set('name', name); c.set('descr', descr)
        etree.SubElement(nv, q('pic:cNvPicPr'))
        bf = etree.SubElement(pic, q('pic:blipFill'))
        blip = etree.SubElement(bf, q('a:blip')); blip.set(q('r:embed'), rid)
        st = etree.SubElement(bf, q('a:stretch')); etree.SubElement(st, q('a:fillRect'))
        sp = etree.SubElement(pic, q('pic:spPr'))
        xf = etree.SubElement(sp, q('a:xfrm'))
        off = etree.SubElement(xf, q('a:off')); off.set('x', '0'); off.set('y', '0')
        ex = etree.SubElement(xf, q('a:ext')); ex.set('cx', str(cx)); ex.set('cy', str(cy))
        pg = etree.SubElement(sp, q('a:prstGeom')); pg.set('prst', 'rect'); etree.SubElement(pg, q('a:avLst'))
        p = E('w:p')
        p.append(make_ppr({'spacing': {'w:before': 120, 'w:after': 120}, 'jc': {'w:val': 'center'}}))
        p.append(E('w:r', None, E('w:drawing', None, inline)))
        return p

    def render(self, blocks):
        out = []
        for blk in blocks:
            kind = blk[0]
            if kind == 'h1':
                out.append(self.heading(blk[1], 1, blk[2] if len(blk) > 2 else None))
            elif kind == 'h2':
                out.append(self.heading(blk[1], 2))
            elif kind == 'p':
                out.append(self.para(blk[1]))
            elif kind == 'pk':                       # run-in sub-heading, kept with next
                out.append(self.para(blk[1], {'keepNext': None, 'spacing': {'w:before': 200, 'w:after': 80}}))
            elif kind == 'b':
                for it in blk[1]:
                    out.append(self.bullet(it))
            elif kind == 'alpha':
                nid = self.ensure_alpha_list()
                for it in blk[1]:
                    out.append(self.bullet(it, numid=nid))
            elif kind == 'table':
                if out and etree.QName(out[-1]).localname == 'tbl':
                    out.append(E('w:p'))             # Word merges adjacent tables
                out.append(self.table(blk[1]))
            elif kind == 'callout':
                if out and etree.QName(out[-1]).localname == 'tbl':
                    out.append(E('w:p'))
                out.append(self.callout(blk[1], blk[2]))
            elif kind == 'img':
                out.append(self.image(blk[1], blk[2], blk[3], blk[4]))
            else:
                raise ValueError(kind)
            # a table directly followed by a heading or another table needs breathing room
        spaced = []
        for i, el in enumerate(out):
            spaced.append(el)
            nxt = out[i + 1] if i + 1 < len(out) else None
            if etree.QName(el).localname == 'tbl' and nxt is not None and etree.QName(nxt).localname == 'p':
                ps = nxt.find('w:pPr/w:pStyle', ns)
                if ps is None or not ps.get(q('w:val')).startswith('Heading'):
                    spaced.append(E('w:p', None, make_ppr({'spacing': {'w:after': 0}})))
        return spaced

    # ------------------------------------------------------------ front matter
    @staticmethod
    def strip_ids(el):
        for e in el.iter():
            for k in list(e.attrib):
                if k.startswith(f'{{{W14}}}') or 'rsid' in k:
                    del e.attrib[k]

    def value_para(self, text):
        """Plain Normal 10pt paragraph matching the template's table cells."""
        return self.para(text, {'spacing': {'w:before': 40, 'w:after': 40}},
                         {'rFonts': {'w:cs': 'Arial'}, 'sz': {'w:val': 20}, 'szCs': {'w:val': 20}})

    def fill_governance(self, tbl, values):
        """values: list of 7 strings in template row order ('' leaves the cell empty)."""
        rows = tbl.findall('w:tr', ns)
        assert len(rows) == len(values) == 7
        for tr, val in zip(rows, values):
            tc = tr.findall('w:tc', ns)[1]
            for c in list(tc):
                if etree.QName(c).localname != 'tcPr':
                    tc.remove(c)
            tc.append(self.value_para(val))
        return tbl

    def fill_history(self, tbl, rows_values):
        """rows_values: list of 5-tuples (Version, Date Approved, Review type, Changes, Approved by)."""
        rows = tbl.findall('w:tr', ns)
        proto = rows[1]
        for tr in rows[1:]:
            tbl.remove(tr)
        for vals in rows_values:
            tr = copy.deepcopy(proto)
            for tc, val in zip(tr.findall('w:tc', ns), vals):
                for c in list(tc):
                    if etree.QName(c).localname != 'tcPr':
                        tc.remove(c)
                tc.append(self.value_para(val))
            tbl.append(tr)
        return tbl

    def toc_paras(self, pages):
        toc_first, toc_mid, toc_end = self.tblocks[42], self.tblocks[43], self.tblocks[55]
        out = []
        for idx, (title, bm) in enumerate(self.toc):
            src = toc_first if idx == 0 else toc_mid
            p = copy.deepcopy(src)
            self.strip_ids(p)
            ppr = p.find('w:pPr', ns)
            for c in list(p):
                if c is not ppr:
                    p.remove(c)
            runs = []

            def fld(kind):
                return E('w:r', None, E('w:fldChar', {'w:fldCharType': kind}))

            def instr(text):
                it = E('w:instrText', text=text)
                it.set(XML_SPACE, 'preserve')
                return E('w:r', None, it)
            if idx == 0:
                runs += [fld('begin'), instr(' TOC \\o "1-1" \\h \\z \\u '), fld('separate')]
            hl = E('w:hyperlink', {'w:anchor': bm, 'w:history': 1})
            plain = re.sub(r'\*\*|//|__', '', title)
            for r in [E('w:r', None, wt(f'{idx + 1}.')), E('w:r', None, E('w:tab')),
                      E('w:r', None, wt(plain)), E('w:r', None, E('w:tab')),
                      fld('begin'), instr(f' PAGEREF {bm} \\h '), fld('separate'),
                      E('w:r', None, wt(str(pages.get(bm, '')))), fld('end')]:
                hl.append(r)
            runs.append(hl)
            for r in runs:
                p.append(r)
            out.append(p)
        end = copy.deepcopy(toc_end)
        self.strip_ids(end)
        out.append(end)
        return out

    @staticmethod
    def page_break():
        return E('w:p', None, E('w:r', None, E('w:br', {'w:type': 'page'})))

    # ------------------------------------------------------------ write
    def finish(self, title, version, front, body, out_docx, cover_title_size=None):
        tb = self.tblocks
        sdt = tb[0]
        for t in sdt.iter(q('w:t')):
            if t.text == 'Process Guide Title':
                t.text = title
            elif t.text == 'Version X.X':
                t.text = f'Version {version}'
        if cover_title_size:
            for t in sdt.iter(q('w:t')):
                if t.text == title:
                    rpr = t.getparent().find('w:rPr', ns)
                    for tag in ('sz', 'szCs'):
                        rpr.find(f'w:{tag}', ns).set(q('w:val'), str(cover_title_size))
        keep = [tb[0]] + tb[5:25] + front + body + [self.body.find('w:sectPr', ns)]
        for c in list(self.body):
            self.body.remove(c)
        for c in keep:
            self.body.append(c)
        self.doc.write(os.path.join(self.build, 'word', 'document.xml'),
                       xml_declaration=True, encoding='UTF-8', standalone=True)
        # drop relationships and media that only the deleted template scaffolding used
        doc_xml = open(os.path.join(self.build, 'word', 'document.xml'), encoding='utf-8').read()
        used = set(re.findall(r'r:(?:id|embed|link)="([^"]+)"', doc_xml))
        for rel in list(self.rels.getroot()):
            if rel.get('Type').rsplit('/', 1)[-1] in ('hyperlink', 'image') and rel.get('Id') not in used:
                self.rels.getroot().remove(rel)
        self.rels.write(self.rels_path, xml_declaration=True, encoding='UTF-8', standalone=True)
        rels_dir = os.path.join(self.build, 'word', '_rels')
        targets = ''.join(open(os.path.join(rels_dir, f), encoding='utf-8').read() for f in os.listdir(rels_dir))
        media = os.path.join(self.build, 'word', 'media')
        for f in os.listdir(media):
            if f'media/{f}"' not in targets:
                os.remove(os.path.join(media, f))
        self.fn_doc.write(self.fn_path, xml_declaration=True, encoding='UTF-8', standalone=True)
        self.num_doc.write(self.num_path, xml_declaration=True, encoding='UTF-8', standalone=True)

        def sub(path, pairs):
            path = os.path.join(self.build, path)
            x = open(path, encoding='utf-8').read()
            for a, b in pairs:
                assert a in x, (path, a)
                x = x.replace(a, b)
            open(path, 'w', encoding='utf-8').write(x)
        esc = title.replace('&', '&amp;')
        sub('[Content_Types].xml', [('wordprocessingml.template.main+xml', 'wordprocessingml.document.main+xml')])
        sub('word/header1.xml', [('&lt;Document title&gt;', esc)])
        sub('word/footer1.xml', [('&lt;Process Guide Title&gt;  |  Version &lt;X.X&gt;', f'{esc}  |  Version {version}')])
        sub('docProps/core.xml', [('<dc:title></dc:title>', f'<dc:title>{esc}</dc:title>')])
        sub('docProps/app.xml', [('<Template>Template - Process v2.dotx</Template>', '<Template>Normal.dotm</Template>')])
        if os.path.exists(out_docx):
            os.remove(out_docx)
        os.system(f'cd "{self.build}" && zip -qXr "{os.path.abspath(out_docx)}" .')
        return out_docx


def front_matter(b, governance_values, history_rows, toc_pages):
    """Document Governance + History tables, then the Table of Contents page."""
    tb = b.tblocks
    gov = b.fill_governance(tb[27], governance_values)
    hist = b.fill_history(tb[31], history_rows)
    front = [tb[25], gov, tb[28], tb[29], hist, b.page_break(), tb[40]]
    front += b.toc_paras(toc_pages)
    front.append(b.page_break())
    return front


# ---------------------------------------------------------------- render + ToC page harvest
SOFFICE_HELPER = os.environ.get(
    'SOFFICE_HELPER_DIR',
    '/root/.claude/skills/synced/222cd086-5dec-4741-a778-8fa3daab3971_c8359c85-a1d1-4519-b8a9-c846f57defad/docx/scripts')


def render_pdf(docx_path, out_dir):
    """Convert with LibreOffice (headless) and return the PDF path."""
    sys_path = list(__import__('sys').path)
    __import__('sys').path.insert(0, SOFFICE_HELPER)
    from office.soffice import run_soffice
    __import__('sys').path[:] = sys_path
    run_soffice(['--headless', '--convert-to', 'pdf', '--outdir', out_dir, docx_path],
                capture_output=True, text=True, timeout=300)
    return os.path.join(out_dir, os.path.splitext(os.path.basename(docx_path))[0] + '.pdf')


def harvest_pages(pdf_path, toc):
    """Find each Heading 1 in the rendered PDF. The template numbers pages from 0
    on the cover (pgNumType start=0, titlePg), so the printed number equals the
    0-based PDF page index."""
    import pymupdf
    doc = pymupdf.open(pdf_path)
    norm = lambda s: re.sub(r'\s+', ' ', s).strip()
    heads = []                                   # (page_index, text) of 20pt headings
    for i, pg in enumerate(doc):
        for blk in pg.get_text('dict')['blocks']:
            for ln in blk.get('lines', []):
                txt = norm(''.join(sp['text'] for sp in ln['spans']))
                if ln['spans'] and max(sp['size'] for sp in ln['spans']) >= 19:
                    heads.append((i, txt))
    pages, start = {}, 0
    for n, (title, bm) in enumerate(toc, 1):
        want = norm(re.sub(r'\*\*|//|__', '', title))
        for k in range(start, len(heads)):
            i, txt = heads[k]
            if txt == want or txt == f'{n}. {want}' or txt.endswith(want):
                pages[bm] = i
                start = k + 1
                break
        else:
            print('MISS', n, title)
    return pages
