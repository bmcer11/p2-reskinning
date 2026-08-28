#!/usr/bin/env python3
"""Reskin the Councillor Individual Briefing Guidelines into the CoPP 'Process' template.

Strategy (same method as the Councillor Interaction Protocol -> 'Other' template job):
- Start from template package; keep cover, CoPP page, governance/history tables, ToC skeleton,
  headers/footers, styles/theme.
- Delete template-instruction table, all red guidance paragraphs, Copilot tables and
  placeholder body sections.
- Insert the instrument's content verbatim (text unchanged), restyled with template
  Heading1/Heading2 + the template's own multilevel decimal numbering (numId 2).
- Tables restyled to the template look (orange F68B1F headers / D0D0D0 borders); the source's
  own "Shared understanding" callout box is preserved as a callout but recoloured into the
  orange palette (it isn't a semantic legend like Table 2 in the Protocol job, just an old
  blue accent that doesn't belong in a Process Guide).
- The source's own Document Governance / Review History table (its own metadata, distinct
  from the template's placeholder governance tables) is carried over verbatim as unnumbered
  headings at the top of the body, same treatment as the Attachments in the prior job.
"""
import copy, os, re, shutil, sys
from lxml import etree

WORK = os.path.dirname(os.path.abspath(__file__))
TPL = os.path.join(WORK, 'template_unpacked')
INS = os.path.join(WORK, 'instrument_unpacked')
BUILD = os.path.join(WORK, 'build2')

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
W14 = 'http://schemas.microsoft.com/office/word/2010/wordml'
ns = {'w': W, 'r': R, 'a': A, 'wp': WP}

TITLE = 'Councillor Individual Briefing Guidelines'
VERSION_LABEL = 'Version 1.0, March 2026'
FOOTER_VERSION = 'Version 1.0'
HEADER_FILL = 'F68B1F'   # this template's orange accent (vs 38322F dark grey in 'Other')
CALLOUT_BORDER = 'F68B1F'
CALLOUT_FILL = 'FCEEDD'

def q(tag):
    p, local = tag.split(':')
    return f'{{{ {"w": W, "r": R, "a": A, "wp": WP}[p] }}}{local}'

def E(tag, attrib=None, children=(), text=None):
    el = etree.Element(q(tag))
    if attrib:
        for k, v in attrib.items():
            el.set(q(k), v)
    for c in children:
        el.append(c)
    if text is not None:
        el.text = text
    return el

def wt(text):
    el = E('w:t', text=text)
    el.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    return el

# ---------------------------------------------------------------- load packages
shutil.rmtree(BUILD, ignore_errors=True)
shutil.copytree(TPL, BUILD)

tdoc = etree.parse(os.path.join(BUILD, 'word', 'document.xml'))
tbody = tdoc.getroot().find('w:body', ns)
tblocks = list(tbody)

idoc = etree.parse(os.path.join(INS, 'word', 'document.xml'))
ibody = idoc.getroot().find('w:body', ns)
iblocks = list(ibody)

# ---------------------------------------------------------------- numbering remap
inum = etree.parse(os.path.join(INS, 'word', 'numbering.xml'))
inum_root = inum.getroot()
i_num2abs = {}
for num in inum_root.iterfind('w:num', ns):
    i_num2abs[num.get(q('w:numId'))] = num.find('w:abstractNumId', ns).get(q('w:val'))

tnum = etree.parse(os.path.join(BUILD, 'word', 'numbering.xml'))
tnum_root = tnum.getroot()

template_style_ids = set()
tstyles = etree.parse(os.path.join(BUILD, 'word', 'styles.xml'))
for s in tstyles.getroot().iterfind('w:style', ns):
    template_style_ids.add(s.get(q('w:styleId')))

# All source bullet lists (numIds 5-16,18) collapse onto the template's single
# ListParagraph bullet list (numId 3). Only the decimal steps list in section 9
# (numId 17) needs its own imported abstract, since the template has nothing
# equivalent free.
BULLET_SRC_IDS = ['5', '6', '7', '8', '9', '10', '11', '12', '13', '14', '15', '16', '18']
NUMMAP = {nid: '3' for nid in BULLET_SRC_IDS}
NEEDED = ['17']

new_abs_id = 100
new_num_id = 100
first_num = tnum_root.find('w:num', ns)
for nid in NEEDED:
    abs_id = i_num2abs.get(nid)
    if abs_id is None:
        continue
    src_abs = None
    for an in inum_root.iterfind('w:abstractNum', ns):
        if an.get(q('w:abstractNumId')) == abs_id:
            src_abs = copy.deepcopy(an)
            break
    src_abs.set(q('w:abstractNumId'), str(new_abs_id))
    for ps in src_abs.iterfind('.//w:pStyle', ns):
        if ps.get(q('w:val')) not in template_style_ids:
            ps.getparent().remove(ps)
    for el in src_abs.iter():
        for attr in ('themeColor', 'themeTint', 'themeShade', 'themeFill', 'themeFillTint', 'themeFillShade'):
            if el.get(q('w:' + attr)) is not None:
                del el.attrib[q('w:' + attr)]
    first_num.addprevious(src_abs)
    num_el = E('w:num', {'w:numId': str(new_num_id)},
               [E('w:abstractNumId', {'w:val': str(new_abs_id)})])
    last_num = tnum_root.findall('w:num', ns)[-1]
    last_num.addnext(num_el)
    NUMMAP[nid] = str(new_num_id)
    new_abs_id += 1
    new_num_id += 1

tnum.write(os.path.join(BUILD, 'word', 'numbering.xml'),
           xml_declaration=True, encoding='UTF-8', standalone=True)

# ---------------------------------------------------------------- relationships
trels_path = os.path.join(BUILD, 'word', '_rels', 'document.xml.rels')
trels = etree.parse(trels_path)
RELNS = 'http://schemas.openxmlformats.org/package/2006/relationships'
rel_root = trels.getroot()

RELMAP = {}
next_rid = 200
def add_rel(rtype, target, mode=None):
    global next_rid
    rid = f'rId{next_rid}'; next_rid += 1
    el = etree.SubElement(rel_root, f'{{{RELNS}}}Relationship')
    el.set('Id', rid); el.set('Type', rtype); el.set('Target', target)
    if mode:
        el.set('TargetMode', mode)
    return rid

# flowchart image (only external asset the body content needs)
shutil.copy(os.path.join(INS, 'word', 'media', 'image5.png'),
            os.path.join(BUILD, 'word', 'media', 'image10.png'))
FIG_RID = add_rel('http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
                  'media/image10.png')
RELMAP['rId17'] = FIG_RID

trels.write(trels_path, xml_declaration=True, encoding='UTF-8', standalone=True)

# ---------------------------------------------------------------- content types
ct_path = os.path.join(BUILD, '[Content_Types].xml')
ct = open(ct_path).read()
ct = ct.replace('application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml',
                'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml')
open(ct_path, 'w').write(ct)

# ---------------------------------------------------------------- headers / footers / props
def sub_file(path, pairs):
    x = open(path).read()
    for a, b in pairs:
        x = x.replace(a, b)
    open(path, 'w').write(x)

sub_file(os.path.join(BUILD, 'word', 'header1.xml'),
         [('&lt;Document title&gt;', TITLE)])
sub_file(os.path.join(BUILD, 'word', 'footer1.xml'),
         [('&lt;Process Guide Title&gt;  |  Version &lt;X.X&gt;', f'{TITLE}  |  {FOOTER_VERSION}')])
core_path = os.path.join(BUILD, 'docProps', 'core.xml')
core = open(core_path).read()
core = re.sub(r'<dc:title>[^<]*</dc:title>', f'<dc:title>{TITLE}</dc:title>', core)
if '<dc:title>' not in core:
    core = core.replace('</cp:coreProperties>', f'<dc:title>{TITLE}</dc:title></cp:coreProperties>')
open(core_path, 'w').write(core)

set_path = os.path.join(BUILD, 'word', 'settings.xml')
st = open(set_path).read()
if '<w:updateFields' not in st:
    for anchor in ('<w:hdrShapeDefaults', '<w:footnotePr', '<w:endnotePr', '<w:compat'):
        if anchor in st:
            st = st.replace(anchor, '<w:updateFields w:val="true"/>' + anchor, 1)
            break
open(set_path, 'w').write(st)

# ---------------------------------------------------------------- helpers: cleaning
STRIP_RPR_BODY = {'rFonts', 'lang', 'noProof', 'spacing', 'kern', 'sz', 'szCs'}
STRIP_RPR_TABLE = {'rFonts', 'lang', 'noProof', 'spacing', 'kern'}

def strip_theme_attrs(el):
    for e in el.iter():
        for attr in ('themeColor', 'themeTint', 'themeShade', 'themeFill',
                     'themeFillTint', 'themeFillShade'):
            k = q('w:' + attr)
            if e.get(k) is not None:
                del e.attrib[k]

def clean_rpr(rpr, table=False):
    if rpr is None:
        return
    strip = STRIP_RPR_TABLE if table else STRIP_RPR_BODY
    for c in list(rpr):
        if etree.QName(c).localname in strip:
            rpr.remove(c)
    strip_theme_attrs(rpr)

def clean_runs(el, table=False):
    for r in el.iter(q('w:r')):
        clean_rpr(r.find('w:rPr', ns), table=table)
        for lrb in r.findall('w:lastRenderedPageBreak', ns):
            r.remove(lrb)
    for hl in el.iter(q('w:hyperlink')):
        rid = hl.get(q('r:id'))
        if rid and rid in RELMAP:
            hl.set(q('r:id'), RELMAP[rid])

def strip_ids(el):
    for e in el.iter():
        for k in list(e.attrib):
            if k.startswith(f'{{{W14}}}') or 'rsid' in k:
                del e.attrib[k]
    etree.strip_elements(el, q('w:proofErr'), with_tail=False)
    etree.strip_elements(el, q('w:bookmarkStart'), with_tail=False)
    etree.strip_elements(el, q('w:bookmarkEnd'), with_tail=False)

def remap_numids(el):
    for numid in el.iterfind('.//w:numPr/w:numId', ns):
        old = numid.get(q('w:val'))
        if old in NUMMAP:
            numid.set(q('w:val'), NUMMAP[old])
        elif old != '0':
            print(f'  WARNING: unmapped numId {old}')

def para_text(p):
    return ''.join(t.text or '' for t in p.iter(q('w:t')))

PPR_ORDER = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr',
             'widowControl', 'numPr', 'suppressLineNumbers', 'pBdr', 'shd', 'tabs',
             'suppressAutoHyphens', 'kinsoku', 'wordWrap', 'overflowPunct',
             'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd',
             'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents',
             'suppressOverlap', 'jc', 'textDirection', 'textAlignment',
             'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr']

TBLPR_ORDER = ['tblStyle', 'tblpPr', 'tblOverlap', 'bidiVisual', 'tblStyleRowBandSize',
               'tblStyleColBandSize', 'tblW', 'jc', 'tblCellSpacing', 'tblInd',
               'tblBorders', 'shd', 'tblLayout', 'tblCellMar', 'tblLook']
TCPR_ORDER = ['cnfStyle', 'tcW', 'gridSpan', 'hMerge', 'vMerge', 'tcBorders', 'shd',
              'noWrap', 'tcMar', 'textDirection', 'tcFitText', 'vAlign', 'hideMark']

def _ordered_insert(parent, el, order):
    name = etree.QName(el).localname
    rank = order.index(name)
    for c in parent:
        cname = etree.QName(c).localname
        crank = order.index(cname) if cname in order else 999
        if crank > rank:
            c.addprevious(el)
            return
    parent.append(el)

def insert_in_order(ppr, el):
    _ordered_insert(ppr, el, PPR_ORDER)

def insert_tblpr(tblpr, el):
    _ordered_insert(tblpr, el, TBLPR_ORDER)

def insert_tcpr(tcpr, el):
    _ordered_insert(tcpr, el, TCPR_ORDER)

# ---------------------------------------------------------------- builders
bookmark_seq = [0]
TOC_ENTRIES = []  # (kind, title, bookmark) — kind 'n' numbered, None unnumbered

def heading(p_src, style, bookmark=False, unnumbered=False, page_break=False, text_override=None):
    text = text_override if text_override is not None else para_text(p_src)
    p = E('w:p')
    ppr = etree.SubElement(p, q('w:pPr'))
    etree.SubElement(ppr, q('w:pStyle')).set(q('w:val'), style)
    if page_break:
        etree.SubElement(ppr, q('w:pageBreakBefore'))
    if unnumbered:
        numpr = etree.SubElement(ppr, q('w:numPr'))
        etree.SubElement(numpr, q('w:numId')).set(q('w:val'), '0')
        ind = etree.SubElement(ppr, q('w:ind'))
        ind.set(q('w:left'), '0'); ind.set(q('w:hanging'), '0')
    bm_name = None
    if bookmark:
        bookmark_seq[0] += 1
        bm_id = str(9000 + bookmark_seq[0])
        bm_name = f'_Toc9000000{bookmark_seq[0]:02d}'
        bs = E('w:bookmarkStart', {'w:id': bm_id, 'w:name': bm_name})
        p.append(bs)
    r = etree.SubElement(p, q('w:r'))
    r.append(wt(text))
    if bookmark:
        p.append(E('w:bookmarkEnd', {'w:id': str(9000 + bookmark_seq[0])}))
    return p, bm_name

def plain_para(p_src, default_left=None, page_break=False, drop_negative_ind=True):
    p = copy.deepcopy(p_src)
    strip_ids(p)
    clean_runs(p, table=False)
    remap_numids(p)
    ppr = p.find('w:pPr', ns)
    had_ind = False
    if ppr is not None:
        for c in list(ppr):
            name = etree.QName(c).localname
            if name == 'pStyle':
                ppr.remove(c)
            elif name == 'rPr':
                clean_rpr(c, table=False)
            elif name == 'ind':
                left = c.get(q('w:left'))
                if drop_negative_ind and left and int(left) < 0:
                    ppr.remove(c)
                else:
                    had_ind = True
            elif name in ('sectPr',):
                ppr.remove(c)
            elif name == 'numPr':
                had_ind = True
    else:
        ppr = etree.SubElement(p, q('w:pPr'))
        p.remove(ppr); p.insert(0, ppr)
    if page_break:
        insert_in_order(ppr, E('w:pageBreakBefore'))
    if not had_ind and default_left:
        insert_in_order(ppr, E('w:ind', {'w:left': default_left}))
    return p

def page_break_para():
    p = E('w:p')
    r = etree.SubElement(p, q('w:r'))
    etree.SubElement(r, q('w:br')).set(q('w:type'), 'page')
    return p

# ---------------------------------------------------------------- table transforms
def make_borders():
    tb = E('w:tblBorders')
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        e = etree.SubElement(tb, q('w:' + edge))
        e.set(q('w:val'), 'single'); e.set(q('w:sz'), '4')
        e.set(q('w:space'), '0'); e.set(q('w:color'), 'D0D0D0')
    return tb

def scale_table(tbl, target=9770):
    grid = tbl.findall('w:tblGrid/w:gridCol', ns)
    widths = [int(g.get(q('w:w'))) for g in grid]
    total = sum(widths)
    if total == 0:
        return
    factor = target / total
    new = [round(w * factor) for w in widths]
    new[-1] += target - sum(new)
    for g, nw in zip(grid, new):
        g.set(q('w:w'), str(nw))
    tblw = tbl.find('w:tblPr/w:tblW', ns)
    if tblw is None:
        tblw = E('w:tblW')
        insert_tblpr(tbl.find('w:tblPr', ns), tblw)
    tblw.set(q('w:w'), str(target)); tblw.set(q('w:type'), 'dxa')
    for tc in tbl.iterfind('.//w:tc', ns):
        tcw = tc.find('w:tcPr/w:tcW', ns)
        if tcw is not None and tcw.get(q('w:type')) == 'dxa':
            tcw.set(q('w:w'), str(round(int(tcw.get(q('w:w'))) * factor)))

def style_cell_dark(tc, fill=HEADER_FILL):
    tcpr = tc.find('w:tcPr', ns)
    if tcpr is None:
        tcpr = E('w:tcPr'); tc.insert(0, tcpr)
    old = tcpr.find('w:shd', ns)
    if old is not None:
        tcpr.remove(old)
    insert_tcpr(tcpr, E('w:shd', {'w:val': 'clear', 'w:color': 'auto', 'w:fill': fill}))
    for p in tc.iterfind('.//w:p', ns):
        ppr = p.find('w:pPr', ns)
        if ppr is None:
            ppr = E('w:pPr'); p.insert(0, ppr)
        prpr = ppr.find('w:rPr', ns)
        if prpr is None:
            prpr = etree.SubElement(ppr, q('w:rPr'))
        for el in (prpr,) + tuple(r.find('w:rPr', ns) for r in p.iterfind('.//w:r', ns)):
            if el is None:
                continue
            for tag in ('color', 'b', 'bCs'):
                for c in el.findall(f'w:{tag}', ns):
                    el.remove(c)
            el.append(E('w:b')); el.append(E('w:bCs'))
            el.append(E('w:color', {'w:val': 'FFFFFF'}))
        for r in p.iterfind('.//w:r', ns):
            if r.find('w:rPr', ns) is None:
                rpr = E('w:rPr')
                rpr.append(E('w:b')); rpr.append(E('w:bCs'))
                rpr.append(E('w:color', {'w:val': 'FFFFFF'}))
                r.insert(0, rpr)

def default_cell_size(tbl):
    for p in tbl.iterfind('.//w:p', ns):
        ppr = p.find('w:pPr', ns)
        if ppr is None:
            ppr = E('w:pPr'); p.insert(0, ppr)
        prpr = ppr.find('w:rPr', ns)
        if prpr is None:
            prpr = etree.SubElement(ppr, q('w:rPr'))
        if prpr.find('w:sz', ns) is None:
            prpr.append(E('w:sz', {'w:val': '20'}))
            prpr.append(E('w:szCs', {'w:val': '20'}))
        for r in p.iterfind('.//w:r', ns):
            rpr = r.find('w:rPr', ns)
            if rpr is None:
                rpr = E('w:rPr'); r.insert(0, rpr)
            if rpr.find('w:sz', ns) is None:
                rpr.append(E('w:sz', {'w:val': '20'}))
                rpr.append(E('w:szCs', {'w:val': '20'}))

def fix_cell_paragraph_styles(tbl):
    for p in tbl.iterfind('.//w:p', ns):
        ppr = p.find('w:pPr', ns)
        if ppr is None:
            continue
        ps = ppr.find('w:pStyle', ns)
        if ps is not None:
            ppr.remove(ps)
        if ppr.find('w:spacing', ns) is None:
            insert_in_order(ppr, E('w:spacing', {'w:before': '40', 'w:after': '40',
                                                 'w:line': '240', 'w:lineRule': 'auto'}))

def transform_table(tbl_src, mode):
    """mode: 'header-row' | 'label-col' | 'callout'"""
    tbl = copy.deepcopy(tbl_src)
    strip_ids(tbl)
    strip_theme_attrs(tbl)
    clean_runs(tbl, table=True)
    remap_numids(tbl)
    fix_cell_paragraph_styles(tbl)
    tblpr = tbl.find('w:tblPr', ns)
    tblppr = tblpr.find('w:tblpPr', ns)
    if tblppr is not None:
        tblpr.remove(tblppr)

    if mode in ('header-row', 'label-col'):
        old_b = tblpr.find('w:tblBorders', ns)
        if old_b is not None:
            tblpr.remove(old_b)
        style_el = tblpr.find('w:tblStyle', ns)
        if style_el is None:
            style_el = E('w:tblStyle', {'w:val': 'TableGrid'})
            tblpr.insert(0, style_el)
        else:
            style_el.set(q('w:val'), 'TableGrid')
        insert_tblpr(tblpr, make_borders())
        for tcb in tbl.iterfind('.//w:tcBorders', ns):
            tcb.getparent().remove(tcb)
        rows = tbl.findall('w:tr', ns)
        if mode == 'header-row':
            for tc in rows[0].findall('w:tc', ns):
                style_cell_dark(tc)
            hdr = rows[0].find('w:trPr', ns)
            if hdr is None:
                hdr = E('w:trPr'); rows[0].insert(0, hdr)
            if hdr.find('w:tblHeader', ns) is None:
                hdr.append(E('w:tblHeader'))
        else:
            for tr in rows:
                tcs = tr.findall('w:tc', ns)
                if tcs:
                    style_cell_dark(tcs[0])
    elif mode == 'callout':
        # single highlighted box: recolour the source's old blue accent into
        # the template's orange, keep the callout shape/border weight
        for tc in tbl.iterfind('.//w:tc', ns):
            tcpr = tc.find('w:tcPr', ns)
            old_b = tcpr.find('w:tcBorders', ns)
            if old_b is not None:
                tcpr.remove(old_b)
            tcb = E('w:tcBorders')
            for edge in ('top', 'left', 'bottom', 'right'):
                e = etree.SubElement(tcb, q('w:' + edge))
                e.set(q('w:val'), 'single'); e.set(q('w:sz'), '8')
                e.set(q('w:space'), '0'); e.set(q('w:color'), CALLOUT_BORDER)
            insert_tcpr(tcpr, tcb)
            old_shd = tcpr.find('w:shd', ns)
            if old_shd is not None:
                tcpr.remove(old_shd)
            insert_tcpr(tcpr, E('w:shd', {'w:val': 'clear', 'w:color': 'auto', 'w:fill': CALLOUT_FILL}))

    default_cell_size(tbl)
    scale_table(tbl, 9770)
    return tbl

# ---------------------------------------------------------------- figure
def transform_figure(p_src):
    p = copy.deepcopy(p_src)
    strip_ids(p)
    clean_runs(p, table=False)
    ppr = p.find('w:pPr', ns)
    if ppr is not None:
        for c in list(ppr):
            if etree.QName(c).localname in ('ind', 'pStyle'):
                ppr.remove(c)
    MAXW = 5731510  # content width 9026 twips, in EMU
    for ext in p.iter(f'{{{WP}}}extent'):
        cx, cy = int(ext.get('cx')), int(ext.get('cy'))
        if cx > MAXW:
            factor = MAXW / cx
            ext.set('cx', str(MAXW)); ext.set('cy', str(round(cy * factor)))
    for ext in p.iter(f'{{{A}}}ext'):
        cx, cy = int(ext.get('cx') or 0), int(ext.get('cy') or 0)
        if cx > MAXW:
            factor = MAXW / cx
            ext.set('cx', str(MAXW)); ext.set('cy', str(round(cy * factor)))
    for blip in p.iter(f'{{{A}}}blip'):
        rid = blip.get(q('r:embed'))
        if rid in RELMAP:
            blip.set(q('r:embed'), RELMAP[rid])
    return p

# ---------------------------------------------------------------- classify instrument blocks
TABLE_MODE = {2: 'label-col', 4: 'header-row', 68: 'callout'}
SKIP = {0, 31}   # empty cover-page sdt shell; the manual page-break paragraph (-> pageBreakBefore on next heading)

out = []
for i, el in enumerate(iblocks):
    tag = etree.QName(el).localname
    if tag == 'sectPr':
        continue
    if i in SKIP:
        continue
    if i in TABLE_MODE:
        out.append(transform_table(el, TABLE_MODE[i]))
        continue
    if tag != 'p':
        continue
    text = para_text(el)
    if i == 30:  # flowchart
        out.append(transform_figure(el))
        continue
    if i in (1, 3):  # Document Governance / Review History — the source's OWN metadata
        p, bm = heading(el, 'Heading1', bookmark=True, unnumbered=True)
        TOC_ENTRIES.append((None, text.strip(), bm))
        out.append(p)
        continue
    if not text.strip():
        continue
    if i == 69:  # '4.1 Practical Examples...' sub-heading -> Heading2 (template auto-numbers '4.1')
        stripped = re.sub(r'^4\.1\s*', '', text)
        p, _ = heading(el, 'Heading2', text_override=stripped)
        out.append(p)
        continue
    ps = el.find('w:pPr/w:pStyle', ns)
    style_name = ps.get(q('w:val')) if ps is not None else None
    if style_name == 'Heading1':
        p, bm = heading(el, 'Heading1', bookmark=True, page_break=(i == 32))
        TOC_ENTRIES.append(('n', text.strip(), bm))
        out.append(p)
        continue
    out.append(plain_para(el))

seq = 0
final_entries = []
for kind, title, bm in TOC_ENTRIES:
    if kind == 'n':
        seq += 1
        final_entries.append((f'{seq}.', title, bm))
    else:
        final_entries.append((None, title, bm))
print(f'{len(out)} body blocks; ToC entries:')
for num, t, bm in final_entries:
    print(f'  {num or "":4s} {t}  [{bm}]')

# ---------------------------------------------------------------- ToC rebuild
toc_first = tblocks[42]
toc_mid = tblocks[43]
toc_end = tblocks[55]

def make_toc_entry(num, title, bm, page, first=False):
    src = toc_first if first else toc_mid
    p = copy.deepcopy(src)
    strip_ids(p)
    ppr = p.find('w:pPr', ns)
    for c in list(p):
        if c is not ppr:
            p.remove(c)
    runs = []
    if first:
        r1 = E('w:r')
        etree.SubElement(r1, q('w:fldChar')).set(q('w:fldCharType'), 'begin')
        r2 = E('w:r')
        it = etree.SubElement(r2, q('w:instrText'))
        it.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        it.text = ' TOC \\o "1-1" \\h \\z \\u '
        r3 = E('w:r')
        etree.SubElement(r3, q('w:fldChar')).set(q('w:fldCharType'), 'separate')
        runs += [r1, r2, r3]
    if num:
        rn = E('w:r'); rn.append(wt(num))
        rt = E('w:r'); etree.SubElement(rt, q('w:tab'))
        runs += [rn, rt]
    rtitle = E('w:r'); rtitle.append(wt(title))
    rtab2 = E('w:r'); etree.SubElement(rtab2, q('w:tab'))
    runs += [rtitle, rtab2]
    rb = E('w:r'); etree.SubElement(rb, q('w:fldChar')).set(q('w:fldCharType'), 'begin')
    ri = E('w:r')
    it = etree.SubElement(ri, q('w:instrText'))
    it.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    it.text = f' PAGEREF {bm} \\h '
    rs = E('w:r'); etree.SubElement(rs, q('w:fldChar')).set(q('w:fldCharType'), 'separate')
    rp = E('w:r'); rp.append(wt(str(page)))
    re_ = E('w:r'); etree.SubElement(re_, q('w:fldChar')).set(q('w:fldCharType'), 'end')
    runs += [rb, ri, rs, rp, re_]
    for r in runs:
        p.append(r)
    return p

toc_pages = {}
toc_pages_path = os.path.join(WORK, 'toc_pages.json')
if os.path.exists(toc_pages_path):
    import json
    toc_pages = json.load(open(toc_pages_path))
toc_paras = []
for idx, (num, title, bm) in enumerate(final_entries):
    toc_paras.append(make_toc_entry(num, title, bm, page=toc_pages.get(bm, 99),
                                    first=(idx == 0)))
toc_paras.append(copy.deepcopy(toc_end))

# ---------------------------------------------------------------- assemble template body
sdt = tblocks[0]
for t in sdt.iter(q('w:t')):
    if t.text == 'Process Guide Title':
        t.text = TITLE
    elif t.text == 'Version X.X':
        t.text = VERSION_LABEL

keep = []
keep.append(tblocks[0])                       # cover
# skip 1 (instructions table), 2-4 (empties)
keep.extend(tblocks[5:25])                    # logo/CoPP contact page incl. trailing pagebreak para
keep.append(tblocks[25])                      # 'Document Governance' heading (template's own)
# skip 26 red guidance
keep.append(tblocks[27])                      # governance table (template placeholders)
keep.append(tblocks[28])                      # spacer
keep.append(tblocks[29])                      # 'Document History' heading (template's own)
# skip 30 red guidance
keep.append(tblocks[31])                      # history table (template placeholders)
keep.append(page_break_para())                # -> ToC page
keep.append(tblocks[40])                      # 'Table of Contents' heading
# skip 41 red guidance
keep.extend(toc_paras)                        # rebuilt ToC (fields cached)
keep.append(page_break_para())                # -> body
keep.extend(out)                              # instrument content
keep.append(tbody.find('w:sectPr', ns))       # final section props

for c in list(tbody):
    tbody.remove(c)
for c in keep:
    tbody.append(c)

tdoc.write(os.path.join(BUILD, 'word', 'document.xml'),
           xml_declaration=True, encoding='UTF-8', standalone=True)
print('document.xml written')

# ---------------------------------------------------------------- zip
out_docx = os.path.join(WORK, 'out2.docx')
if os.path.exists(out_docx):
    os.remove(out_docx)
os.system(f'cd "{BUILD}" && zip -qXr "{out_docx}" .')
print('wrote', out_docx, os.path.getsize(out_docx), 'bytes')
