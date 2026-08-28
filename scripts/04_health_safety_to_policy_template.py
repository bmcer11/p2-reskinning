#!/usr/bin/env python3
"""Reskin the Health and Safety Policy (V2.2) into the CoPP 'Policy' template.

Same method as the three prior jobs. Notes specific to this source:
- The source carried tracked changes (the V2.2 update itself: version 1->2.2,
  date -> 12/08/2026, new psychosocial regulations). Those were accepted first,
  so this builds from the current V2.2 text.
- Source section titles carry literal numbers ("1. Purpose", "4.1 Health and
  Safety Policy overview"); the prefixes are stripped so the template's own
  auto-numbering supplies them.
- Old-skin decoration is dropped: the cover photograph, the target icon in the
  Purpose heading, and a hairline rule shape - the template supplies its own.
- The source's real sub-bullets (an 'o' list) map onto the template's bullet
  list level 1, which uses the same glyph. Literal "* " and "1. " glyphs typed
  into body text are left exactly as written (no words changed).
"""
import copy, os, re, shutil
from lxml import etree

WORK = os.path.dirname(os.path.abspath(__file__))
TPL = os.path.join(WORK, 'template_unpacked')
INS = os.path.join(WORK, 'instrument_unpacked')
BUILD = os.path.join(WORK, 'build4')

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
W14 = 'http://schemas.microsoft.com/office/word/2010/wordml'
ns = {'w': W, 'r': R, 'a': A, 'wp': WP}

TITLE = 'Health and Safety Policy'
VERSION_LABEL = 'Version 2.2'
FOOTER_VERSION = 'Version 2.2'
HEADER_FILL = '20ABAD'        # this template's teal table-header fill
BULLET_NUMID = '3'            # template's own bullet list

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

# ---------------------------------------------------------------- relationships
trels_path = os.path.join(BUILD, 'word', '_rels', 'document.xml.rels')
trels = etree.parse(trels_path)
RELNS = 'http://schemas.openxmlformats.org/package/2006/relationships'
rel_root = trels.getroot()
irels = etree.parse(os.path.join(INS, 'word', '_rels', 'document.xml.rels'))

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

for rel in irels.getroot():           # the three external hyperlinks
    if rel.get('Type').endswith('/hyperlink'):
        RELMAP[rel.get('Id')] = add_rel(rel.get('Type'), rel.get('Target'), rel.get('TargetMode'))

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

sub_file(os.path.join(BUILD, 'word', 'header1.xml'), [('&lt;Document title&gt;', TITLE)])
sub_file(os.path.join(BUILD, 'word', 'footer1.xml'),
         [('&lt;Policy Title&gt;  |  Version &lt;X.X&gt;', f'{TITLE}  |  {FOOTER_VERSION}')])
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

# ---------------------------------------------------------------- helpers
STRIP_RPR_BODY = {'rFonts', 'lang', 'noProof', 'spacing', 'kern', 'sz', 'szCs', 'color'}
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

def drop_graphics(el):
    """Remove old-skin decoration (cover photo, heading icon, rule shapes)."""
    for tag in ('w:drawing', 'w:pict', 'w:object'):
        for g in list(el.iter(q(tag))):
            g.getparent().remove(g)

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

def insert_in_order(ppr, el): _ordered_insert(ppr, el, PPR_ORDER)
def insert_tblpr(tblpr, el):  _ordered_insert(tblpr, el, TBLPR_ORDER)
def insert_tcpr(tcpr, el):    _ordered_insert(tcpr, el, TCPR_ORDER)

# ---------------------------------------------------------------- builders
bookmark_seq = [0]
TOC_ENTRIES = []
NUM_PREFIX = re.compile(r'^\s*\d+(?:\.\d+)*\.?\s+')

def heading(p_src, style, bookmark=False, unnumbered=False):
    text = NUM_PREFIX.sub('', para_text(p_src)).strip()
    p = E('w:p')
    ppr = etree.SubElement(p, q('w:pPr'))
    etree.SubElement(ppr, q('w:pStyle')).set(q('w:val'), style)
    if unnumbered:
        numpr = etree.SubElement(ppr, q('w:numPr'))
        etree.SubElement(numpr, q('w:numId')).set(q('w:val'), '0')
        ind = etree.SubElement(ppr, q('w:ind'))
        ind.set(q('w:left'), '0'); ind.set(q('w:hanging'), '0')
    bm_name = None
    if bookmark:
        bookmark_seq[0] += 1
        bm_name = f'_Toc9000000{bookmark_seq[0]:02d}'
        p.append(E('w:bookmarkStart', {'w:id': str(9000 + bookmark_seq[0]), 'w:name': bm_name}))
    r = etree.SubElement(p, q('w:r'))
    r.append(wt(text))
    if bookmark:
        p.append(E('w:bookmarkEnd', {'w:id': str(9000 + bookmark_seq[0])}))
    return p, bm_name, text

def body_para(p_src, bullet_level=None):
    p = copy.deepcopy(p_src)
    strip_ids(p)
    drop_graphics(p)
    clean_runs(p, table=False)
    ppr = p.find('w:pPr', ns)
    if ppr is None:
        ppr = E('w:pPr'); p.insert(0, ppr)
    for c in list(ppr):
        name = etree.QName(c).localname
        if name in ('pStyle', 'numPr', 'ind', 'tabs', 'pBdr', 'sectPr'):
            ppr.remove(c)
        elif name == 'rPr':
            clean_rpr(c, table=False)
    if bullet_level is not None:
        numpr = E('w:numPr')
        etree.SubElement(numpr, q('w:ilvl')).set(q('w:val'), str(bullet_level))
        etree.SubElement(numpr, q('w:numId')).set(q('w:val'), BULLET_NUMID)
        insert_in_order(ppr, numpr)
    return p

def page_break_para():
    p = E('w:p')
    r = etree.SubElement(p, q('w:r'))
    etree.SubElement(r, q('w:br')).set(q('w:type'), 'page')
    return p

# ---------------------------------------------------------------- table transform
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
        tblw = E('w:tblW'); insert_tblpr(tbl.find('w:tblPr', ns), tblw)
    tblw.set(q('w:w'), str(target)); tblw.set(q('w:type'), 'dxa')
    for tc in tbl.iterfind('.//w:tc', ns):
        tcw = tc.find('w:tcPr/w:tcW', ns)
        if tcw is not None and tcw.get(q('w:type')) == 'dxa':
            tcw.set(q('w:w'), str(round(int(tcw.get(q('w:w'))) * factor)))

def style_header_cell(tc):
    tcpr = tc.find('w:tcPr', ns)
    if tcpr is None:
        tcpr = E('w:tcPr'); tc.insert(0, tcpr)
    old = tcpr.find('w:shd', ns)
    if old is not None:
        tcpr.remove(old)
    insert_tcpr(tcpr, E('w:shd', {'w:val': 'clear', 'w:color': 'auto', 'w:fill': HEADER_FILL}))
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
            prpr.append(E('w:sz', {'w:val': '20'})); prpr.append(E('w:szCs', {'w:val': '20'}))
        for r in p.iterfind('.//w:r', ns):
            rpr = r.find('w:rPr', ns)
            if rpr is None:
                rpr = E('w:rPr'); r.insert(0, rpr)
            if rpr.find('w:sz', ns) is None:
                rpr.append(E('w:sz', {'w:val': '20'})); rpr.append(E('w:szCs', {'w:val': '20'}))

def fix_cell_paragraphs(tbl):
    for p in tbl.iterfind('.//w:p', ns):
        ppr = p.find('w:pPr', ns)
        if ppr is None:
            continue
        for c in list(ppr):
            if etree.QName(c).localname in ('pStyle', 'ind', 'numPr', 'tabs'):
                ppr.remove(c)
        if ppr.find('w:spacing', ns) is None:
            insert_in_order(ppr, E('w:spacing', {'w:before': '40', 'w:after': '40',
                                                 'w:line': '240', 'w:lineRule': 'auto'}))

def transform_table(tbl_src):
    tbl = copy.deepcopy(tbl_src)
    strip_ids(tbl)
    strip_theme_attrs(tbl)
    drop_graphics(tbl)
    clean_runs(tbl, table=True)
    fix_cell_paragraphs(tbl)
    tblpr = tbl.find('w:tblPr', ns)
    tblppr = tblpr.find('w:tblpPr', ns)
    if tblppr is not None:
        tblpr.remove(tblppr)
    old_b = tblpr.find('w:tblBorders', ns)
    if old_b is not None:
        tblpr.remove(old_b)
    style_el = tblpr.find('w:tblStyle', ns)
    if style_el is None:
        style_el = E('w:tblStyle', {'w:val': 'TableGrid'}); tblpr.insert(0, style_el)
    else:
        style_el.set(q('w:val'), 'TableGrid')
    insert_tblpr(tblpr, make_borders())
    for tcb in tbl.iterfind('.//w:tcBorders', ns):
        tcb.getparent().remove(tcb)
    rows = tbl.findall('w:tr', ns)
    for tc in rows[0].findall('w:tc', ns):
        style_header_cell(tc)
    trpr = rows[0].find('w:trPr', ns)
    if trpr is None:
        trpr = E('w:trPr'); rows[0].insert(0, trpr)
    if trpr.find('w:tblHeader', ns) is None:
        trpr.append(E('w:tblHeader'))
    default_cell_size(tbl)
    scale_table(tbl, 9770)
    return tbl

# ---------------------------------------------------------------- classify + convert
H1        = {8, 16, 35, 40}      # 1. Purpose / 2. Policy / 3. Related documents / 4. Internal Overview
H2        = {41, 57}             # 4.1 overview / 4.2 document history
H1_UNNUM  = {68}                 # 'Here to help.' closing block (not a numbered policy section)
SUBBULLET = {20, 21, 22}         # the source's real 'o' list -> template bullet level 1
TABLES    = {36, 38, 58}
SKIP      = {0, 1, 2, 3, 4,      # cover title block -> template cover
             5,                  # 'REVISION NO. 2.2 | ISSUED: ...' -> cover/footer
             6, 7,               # source ToC + its page break -> template ToC
             56}                 # hairline rule shape (old skin)

out = []
for i, el in enumerate(iblocks):
    name = etree.QName(el).localname
    if name == 'sectPr' or i in SKIP:
        continue
    if i in TABLES:
        out.append(transform_table(el))
        continue
    if name != 'p':
        continue
    if not para_text(el).strip():
        continue
    if i in H1:
        p, bm, text = heading(el, 'Heading1', bookmark=True)
        TOC_ENTRIES.append(text)
        out.append(p)
    elif i in H2:
        p, _, _ = heading(el, 'Heading2')
        out.append(p)
    elif i in H1_UNNUM:
        p, _, _ = heading(el, 'Heading1', unnumbered=True)
        out.append(p)
    else:
        out.append(body_para(el, bullet_level=1 if i in SUBBULLET else None))

# Word merges two tables that are adjacent siblings — keep a spacer between them
spaced = []
for blk in out:
    if (spaced and etree.QName(blk).localname == 'tbl'
            and etree.QName(spaced[-1]).localname == 'tbl'):
        spaced.append(E('w:p'))
    spaced.append(blk)
out = spaced

final_entries = [(f'{n+1}.', t, f'_Toc9000000{n+1:02d}') for n, t in enumerate(TOC_ENTRIES)]
print(f'{len(out)} body blocks; ToC entries:')
for num, t, bm in final_entries:
    print(f'  {num:4s} {t}  [{bm}]')

# ---------------------------------------------------------------- ToC rebuild
toc_first, toc_mid, toc_end = tblocks[41], tblocks[42], tblocks[49]

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
        r1 = E('w:r'); etree.SubElement(r1, q('w:fldChar')).set(q('w:fldCharType'), 'begin')
        r2 = E('w:r')
        it = etree.SubElement(r2, q('w:instrText'))
        it.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        it.text = ' TOC \\o "1-1" \\h \\z \\u '
        r3 = E('w:r'); etree.SubElement(r3, q('w:fldChar')).set(q('w:fldCharType'), 'separate')
        runs += [r1, r2, r3]
    rn = E('w:r'); rn.append(wt(num))
    rt = E('w:r'); etree.SubElement(rt, q('w:tab'))
    rtitle = E('w:r'); rtitle.append(wt(title))
    rtab2 = E('w:r'); etree.SubElement(rtab2, q('w:tab'))
    runs += [rn, rt, rtitle, rtab2]
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
toc_paras = [make_toc_entry(num, title, bm, toc_pages.get(bm, 99), first=(idx == 0))
             for idx, (num, title, bm) in enumerate(final_entries)]
toc_paras.append(copy.deepcopy(toc_end))

# ---------------------------------------------------------------- assemble
for t in tblocks[1].iter(q('w:t')):          # cover content control
    if t.text == 'Policy Title':
        t.text = TITLE
    elif t.text == 'Version X.X':
        t.text = VERSION_LABEL

keep = []
keep.append(tblocks[0])                       # leading empty paragraph
keep.append(tblocks[1])                       # cover
# skip [2] template instructions table
keep.extend(tblocks[3:26])                    # CoPP contact page + 'Document Governance' heading
# skip [26] red guidance
keep.append(tblocks[27])                      # governance table (placeholders kept)
keep.append(tblocks[28])                      # spacer
keep.append(tblocks[29])                      # 'Document History' heading
# skip [30] red guidance
keep.append(tblocks[31])                      # history table (placeholders kept)
keep.append(page_break_para())                # -> ToC page (the template relied on spacer
keep.append(tblocks[39])                      #    paragraphs sized around its guidance text)
# skip [40] red guidance
keep.extend(toc_paras)
keep.append(tblocks[50])                      # page break -> body
keep.extend(out)
keep.append(tbody.find('w:sectPr', ns))

for c in list(tbody):
    tbody.remove(c)
for c in keep:
    tbody.append(c)

tdoc.write(os.path.join(BUILD, 'word', 'document.xml'),
           xml_declaration=True, encoding='UTF-8', standalone=True)
print('document.xml written')

out_docx = os.path.join(WORK, 'out4.docx')
if os.path.exists(out_docx):
    os.remove(out_docx)
os.system(f'cd "{BUILD}" && zip -qXr "{out_docx}" .')
print('wrote', out_docx, os.path.getsize(out_docx), 'bytes')
