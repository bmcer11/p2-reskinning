#!/usr/bin/env python3
"""Reskin the Councillor Interaction and Request Protocol into the CoPP 'Other' template.

Strategy:
- Start from template package; keep cover, CoPP page, governance/history tables, ToC skeleton,
  headers/footers, styles/theme.
- Delete template-instruction table, all red guidance paragraphs, Copilot tables and
  placeholder body sections.
- Insert the instrument's content verbatim (text unchanged), restyled with template
  Heading1/Heading2 + the template's own multilevel decimal numbering (numId 2).
- Lettered a) items: template abstractNum 1 level 3 is repurposed to lowerLetter "%4)"
  (nothing in the template used level 3).
- Tables: restyled to template look (dark 38322F headers / D0D0D0 borders) except Table 2 +
  its colour legend, whose fills are semantic and preserved exactly (theme refs resolved to
  literal colours since the theme changes).
"""
import copy, os, re, shutil, sys
from lxml import etree

WORK = os.path.dirname(os.path.abspath(__file__))
TPL = os.path.join(WORK, 'template_unpacked')
INS = os.path.join(WORK, 'instrument_unpacked')
BUILD = os.path.join(WORK, 'build')

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
W14 = 'http://schemas.microsoft.com/office/word/2010/wordml'
ns = {'w': W, 'r': R, 'a': A, 'wp': WP}

TITLE = 'Councillor Interaction and Request Protocol'
VERSION = '2'
DOCTYPE = 'PROTOCOL'

def q(tag):
    p, local = tag.split(':')
    return f'{{{ {"w": W, "r": R, "a": A, "wp": WP}[p] }}}{local}'

def E(tag, attrib=None, children=(), text=None):
    el = etree.SubElement if False else etree.Element(q(tag))
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
# instrument numId -> abstractNumId
inum = etree.parse(os.path.join(INS, 'word', 'numbering.xml'))
inum_root = inum.getroot()
i_num2abs = {}
for num in inum_root.iterfind('w:num', ns):
    i_num2abs[num.get(q('w:numId'))] = num.find('w:abstractNumId', ns).get(q('w:val'))

tnum = etree.parse(os.path.join(BUILD, 'word', 'numbering.xml'))
tnum_root = tnum.getroot()

# 1. Repurpose template abstractNum 1 level 3 -> lowerLetter "a)"
for an in tnum_root.iterfind('w:abstractNum', ns):
    if an.get(q('w:abstractNumId')) == '1':
        for lvl in an.iterfind('w:lvl', ns):
            if lvl.get(q('w:ilvl')) == '3':
                lvl.find('w:numFmt', ns).set(q('w:val'), 'lowerLetter')
                lvl.find('w:lvlText', ns).set(q('w:val'), '%4)')
                isLgl = lvl.find('w:isLgl', ns)   # isLgl forces decimal display
                if isLgl is not None:
                    lvl.remove(isLgl)
                ind = lvl.find('w:pPr/w:ind', ns)
                if ind is not None:
                    ind.set(q('w:left'), '1985'); ind.set(q('w:hanging'), '567')

# 2. Import instrument abstracts needed by kept content (lists other than the main one)
NEEDED = ['2', '3', '4', '16', '17', '18', '19', '20', '21', '22', '23', '24', '25']
template_style_ids = set()
tstyles = etree.parse(os.path.join(BUILD, 'word', 'styles.xml'))
for s in tstyles.getroot().iterfind('w:style', ns):
    template_style_ids.add(s.get(q('w:styleId')))

NUMMAP = {'1': '2'}  # instrument main list -> template heading list
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
    # drop style links to styles the template doesn't have; strip theme attrs
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
existing_ids = {rel.get('Id') for rel in rel_root}

irels = etree.parse(os.path.join(INS, 'word', '_rels', 'document.xml.rels'))
i_rel = {rel.get('Id'): rel for rel in irels.getroot()}

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

# figure image
shutil.copy(os.path.join(INS, 'word', 'media', 'image2.jpg'),
            os.path.join(BUILD, 'word', 'media', 'image10.jpg'))
FIG_RID = add_rel('http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
                  'media/image10.jpg')
RELMAP['rId13'] = FIG_RID
# hyperlinks
for old_id in ['rId15', 'rId16', 'rId17', 'rId18', 'rId19']:
    src = i_rel[old_id]
    RELMAP[old_id] = add_rel(src.get('Type'), src.get('Target'), src.get('TargetMode'))

trels.write(trels_path, xml_declaration=True, encoding='UTF-8', standalone=True)

# ---------------------------------------------------------------- content types
ct_path = os.path.join(BUILD, '[Content_Types].xml')
ct = open(ct_path).read()
ct = ct.replace('application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml',
                'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml')
if 'Extension="jpg"' not in ct:
    ct = ct.replace('</Types>', '<Default Extension="jpg" ContentType="image/jpeg"/></Types>')
open(ct_path, 'w').write(ct)

# ---------------------------------------------------------------- headers / footers / props
def sub_file(path, pairs):
    x = open(path).read()
    for a, b in pairs:
        x = x.replace(a, b)
    open(path, 'w').write(x)

sub_file(os.path.join(BUILD, 'word', 'header1.xml'),
         [('&lt;Document title&gt;', TITLE)])
sub_file(os.path.join(BUILD, 'word', 'header2.xml'),
         [('DOCUMENT TYPE', DOCTYPE)])
sub_file(os.path.join(BUILD, 'word', 'footer1.xml'),
         [('&lt;Instrument Title&gt;  |  Version &lt;X.X&gt;', f'{TITLE}  |  Version {VERSION}')])
core_path = os.path.join(BUILD, 'docProps', 'core.xml')
core = open(core_path).read()
core = re.sub(r'<dc:title>[^<]*</dc:title>', f'<dc:title>{TITLE}</dc:title>', core)
if '<dc:title>' not in core:
    core = core.replace('</cp:coreProperties>', f'<dc:title>{TITLE}</dc:title></cp:coreProperties>')
open(core_path, 'w').write(core)

# settings: ask Word to refresh fields (ToC) on open — schema position is just
# before hdrShapeDefaults/footnotePr/endnotePr/compat, whichever comes first
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
    """Clean every run under el in place."""
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
        else:
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
TOC_ENTRIES = []  # (number or None, title, bookmark)

def heading(p_src, style, bookmark=False, unnumbered=False, page_break=False):
    text = para_text(p_src)
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

CLAUSE_IND = {1: ('709', '709'), 2: ('1418', '709'), 3: ('1985', '567')}

def clause(p_src, ilvl):
    p = copy.deepcopy(p_src)
    strip_ids(p)
    clean_runs(p, table=False)
    old_ppr = p.find('w:pPr', ns)
    keep_children = []
    if old_ppr is not None:
        for c in old_ppr:
            if etree.QName(c).localname in ('rPr',):
                clean_rpr(c, table=False)
                keep_children.append(c)
        p.remove(old_ppr)
    ppr = E('w:pPr')
    numpr = etree.SubElement(ppr, q('w:numPr'))
    etree.SubElement(numpr, q('w:ilvl')).set(q('w:val'), str(ilvl))
    etree.SubElement(numpr, q('w:numId')).set(q('w:val'), '2')
    ind = etree.SubElement(ppr, q('w:ind'))
    left, hang = CLAUSE_IND[ilvl]
    ind.set(q('w:left'), left); ind.set(q('w:hanging'), hang)
    for c in keep_children:
        ppr.append(c)
    p.insert(0, ppr)
    return p

def plain_para(p_src, default_left=None, page_break=False, drop_negative_ind=True):
    """Non-numbered paragraph: keep runs (cleaned), keep explicit ind/jc/spacing,
    drop template-foreign styles."""
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
                ppr.remove(c)  # ListParagraph etc. -> Normal
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
                had_ind = True  # numbering will position it
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
TPL_BORDER = None  # built below

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
    # scale cell widths proportionally
    for tc in tbl.iterfind('.//w:tc', ns):
        tcw = tc.find('w:tcPr/w:tcW', ns)
        if tcw is not None and tcw.get(q('w:type')) == 'dxa':
            tcw.set(q('w:w'), str(round(int(tcw.get(q('w:w'))) * factor)))

def style_cell_dark(tc):
    tcpr = tc.find('w:tcPr', ns)
    if tcpr is None:
        tcpr = E('w:tcPr'); tc.insert(0, tcpr)
    old = tcpr.find('w:shd', ns)
    if old is not None:
        tcpr.remove(old)
    insert_tcpr(tcpr, E('w:shd', {'w:val': 'clear', 'w:color': 'auto', 'w:fill': '38322F'}))
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
    """Give any table run/paragraph-mark with no explicit size sz 20 (10pt)."""
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
    """Remove template-foreign pStyles in cells; keep spacing tight."""
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
    """mode: 'header-row' | 'label-col' | 'preserve'"""
    tbl = copy.deepcopy(tbl_src)
    strip_ids(tbl)
    strip_theme_attrs(tbl)
    clean_runs(tbl, table=True)
    remap_numids(tbl)
    fix_cell_paragraph_styles(tbl)
    tblpr = tbl.find('w:tblPr', ns)
    tblppr = tblpr.find('w:tblpPr', ns)   # floating position -> inline flow
    if tblppr is not None:
        tblpr.remove(tblppr)
    # ensure a table style + tblLook survive; keep original tblLayout etc.
    if mode in ('header-row', 'label-col'):
        old_b = tblpr.find('w:tblBorders', ns)
        if old_b is not None:
            tblpr.remove(old_b)
        # order inside tblPr matters: tblStyle, tblW, tblBorders ... keep it simple:
        style_el = tblpr.find('w:tblStyle', ns)
        if style_el is None:
            style_el = E('w:tblStyle', {'w:val': 'TableGrid'})
            tblpr.insert(0, style_el)
        else:
            style_el.set(q('w:val'), 'TableGrid')
        insert_tblpr(tblpr, make_borders())
        # drop per-cell borders so the uniform grid shows
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
    # rescale to content width 9026 twips = 5731510 EMU
    MAXW = 5731510
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
H2_IDX = {41, 46, 54, 71, 83, 105, 116, 135, 162, 174, 192, 196, 201, 207, 211, 218}
TABLE_MODE = {1: 'label-col', 68: 'header-row', 149: 'preserve', 151: 'preserve',
              189: 'header-row', 257: 'header-row', 273: 'label-col', 324: 'header-row'}
SKIP = {0, 2, 63, 69, 70, 275, 353, 354}          # empties handled generically; these are structural
PAGEBREAK_BEFORE = {64, 148, 276, 314}             # figure caption, Table 2 caption, attachments

def get_num(p):
    numpr = p.find('w:pPr/w:numPr', ns)
    if numpr is None:
        return None, None
    numid = numpr.find('w:numId', ns)
    ilvl = numpr.find('w:ilvl', ns)
    return (numid.get(q('w:val')) if numid is not None else None,
            int(ilvl.get(q('w:val'))) if ilvl is not None else 0)

out = []
for i, el in enumerate(iblocks):
    tag = etree.QName(el).localname
    if tag == 'sectPr':
        continue
    if i in SKIP:
        continue
    if tag == 'tbl':
        out.append(transform_table(el, TABLE_MODE[i]))
        continue
    if tag != 'p':
        continue
    text = para_text(el)
    numid, ilvl = get_num(el)
    if i == 65:   # Figure 1 image
        out.append(transform_figure(el))
        continue
    if i == 66:   # explicit page break between figure and Table 1
        out.append(page_break_para())
        continue
    if not text.strip():
        continue  # spacer paragraphs: template styles provide the rhythm
    if i == 276:  # ATTACHMENT 1
        p, bm = heading(el, 'Heading1', bookmark=True, unnumbered=True, page_break=True)
        TOC_ENTRIES.append((None, text.strip(), bm))
        out.append(p)
        continue
    if i == 314:  # ATTACHMENT 2 (strip its inline page-break run; use pageBreakBefore)
        el2 = copy.deepcopy(el)
        for br in el2.iterfind('.//w:br', ns):
            br.getparent().remove(br)
        p, bm = heading(el2, 'Heading1', bookmark=True, unnumbered=True, page_break=True)
        TOC_ENTRIES.append((None, text.strip(), bm))
        out.append(p)
        continue
    if numid == '1' and ilvl == 0:
        p, bm = heading(el, 'Heading1', bookmark=True)
        TOC_ENTRIES.append(('n', text.strip(), bm))
        out.append(p)
        continue
    if numid == '1' and i in H2_IDX:
        p, _ = heading(el, 'Heading2')
        out.append(p)
        continue
    if numid == '1':
        out.append(clause(el, ilvl))
        continue
    # everything else: captions, bold labels, attachment lists, continuation paragraphs
    default_left = None
    ppr = el.find('w:pPr', ns)
    has_ind = ppr is not None and ppr.find('w:ind', ns) is not None
    ps = ppr.find('w:pStyle', ns) if ppr is not None else None
    if not has_ind and numid is None and ps is not None and ps.get(q('w:val')) == 'ListParagraph':
        default_left = '720'
    out.append(plain_para(el, default_left=default_left, page_break=(i in PAGEBREAK_BEFORE)))

# number the numbered TOC entries
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
# template blocks: [040] first entry (has TOC field begin), [041] middle entry, [058] field end
toc_first = tblocks[40]
toc_mid = tblocks[41]
toc_end = tblocks[58]

def make_toc_entry(num, title, bm, page, first=False):
    """Clone template entry structure, swap content."""
    src = toc_first if first else toc_mid
    p = copy.deepcopy(src)
    strip_ids(p)
    # collect and remove all runs/hyperlinks after pPr; rebuild
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
    # PAGEREF field
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
# patch cover text
sdt = tblocks[0]
for t in sdt.iter(q('w:t')):
    if t.text == 'Document Title':
        t.text = TITLE
    elif t.text == 'Version X.X':
        t.text = f'Version {VERSION}'

keep = []
keep.append(tblocks[0])                       # cover
# skip 1..3 (empty, instructions table, empty)
keep.extend(tblocks[4:24])                    # CoPP contact page incl. trailing pagebreak para
keep.append(tblocks[24])                      # Document Governance heading
# skip 25 red guidance
keep.append(tblocks[26])                      # governance table
keep.append(tblocks[27])                      # spacer
keep.append(tblocks[28])                      # Document History heading
# skip 29 red guidance
keep.append(tblocks[30])                      # history table
keep.append(page_break_para())                # -> ToC page
keep.append(tblocks[38])                      # Table of Contents heading
# skip 39 red guidance
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
out_docx = os.path.join(WORK, 'out.docx')
if os.path.exists(out_docx):
    os.remove(out_docx)
os.system(f'cd "{BUILD}" && zip -qXr "{out_docx}" .')
print('wrote', out_docx, os.path.getsize(out_docx), 'bytes')
