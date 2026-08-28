#!/usr/bin/env python3
"""Reskin the City of Port Phillip Governance Framework into the CoPP 'Framework' template.

This one is inverted from the earlier jobs. The source is a designed 56-page
LANDSCAPE publication - 27 sections driving two-column layouts, ~20 inline
images and infographics, running section banners in the headers - so the base
package is the instrument itself (keeping that design intact) and the
template's styling is applied over it, rather than pouring content into the
portrait template.

What the template contributes:
- its style definitions for every shared style (Normal, BodyText, ListParagraph,
  Heading 1-5 and their linked character styles, TOC 1-2, table styles),
- its theme (so theme-driven table styles pick up the framework palette),
- its heading palette: the source's teal family is mapped to the template's
  magenta/plum,
- its Document Governance and Document History tables, which the source lacked.

What is deliberately preserved from the source:
- landscape orientation, all 27 sections, columns, images and header banners,
- the author's heading levels exactly as written (H1-H5), and every word,
- unnumbered headings: the template's heading styles carry automatic numbering,
  which is stripped on import because this framework uses named sections
  ("Executive Summary", "Roles and Powers") rather than numbered clauses.
"""
import copy, os, re, shutil
from lxml import etree

WORK = os.path.dirname(os.path.abspath(__file__))
TPL = os.path.join(WORK, 'template_unpacked')
INS = os.path.join(WORK, 'instrument_unpacked')
BUILD = os.path.join(WORK, 'build7')

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
ns = {'w': W, 'r': R}

TITLE = 'City of Port Phillip Governance Framework'

# the source's teal accents -> the framework template's magenta/plum
PALETTE = {
    '2EBBB8': '9D2270', '008080': '9D2270', '009999': '9D2270',
    '3E988D': '9D2270', '007184': '9D2270', '2BBBB0': '9D2270',
    '00A3AD': '9D2270',
}

def q(tag):
    p, local = tag.split(':')
    return f'{{{ {"w": W, "r": R}[p] }}}{local}'

def E(tag, attrib=None):
    el = etree.Element(q(tag))
    for k, v in (attrib or {}).items():
        el.set(q(k), v)
    return el

# ---------------------------------------------------------------- base package
shutil.rmtree(BUILD, ignore_errors=True)
shutil.copytree(INS, BUILD)                      # keep the designed landscape doc

# ---------------------------------------------------------------- 1. styles
tstyles = etree.parse(os.path.join(TPL, 'word', 'styles.xml'))
istyles = etree.parse(os.path.join(BUILD, 'word', 'styles.xml'))
T = {s.get(q('w:styleId')): s for s in tstyles.getroot().iterfind('w:style', ns)}
iroot = istyles.getroot()

HEADINGISH = re.compile(r'^(Heading[1-9]|Style1)')

def detemplate_heading(style):
    """Take the template's heading look, minus its automatic clause numbering."""
    ppr = style.find('w:pPr', ns)
    if ppr is None:
        return
    for numpr in ppr.findall('w:numPr', ns):
        ppr.remove(numpr)                        # source headings are unnumbered
    for ind in ppr.findall('w:ind', ns):
        ppr.remove(ind)                          # ...so no hanging indent either

swapped = []
for s in list(iroot.iterfind('w:style', ns)):
    sid = s.get(q('w:styleId'))
    if sid in T:
        new = copy.deepcopy(T[sid])
        if HEADINGISH.match(sid or ''):
            detemplate_heading(new)
        s.addprevious(new)
        iroot.remove(s)
        swapped.append(sid)

# instrument-only heading styles (Style1) are based on Heading1 but pin their own
# teal; drop that so they inherit the template's heading colour
for s in iroot.iterfind('w:style', ns):
    sid = s.get(q('w:styleId')) or ''
    base = s.find('w:basedOn', ns)
    based_on = base.get(q('w:val')) if base is not None else ''
    if HEADINGISH.match(sid) or HEADINGISH.match(based_on or '') \
            or (based_on or '').startswith('Heading'):
        rpr = s.find('w:rPr', ns)
        if rpr is not None:
            for c in rpr.findall('w:color', ns):
                rpr.remove(c)

# docDefaults from the template too, so body text matches
tdd = tstyles.getroot().find('w:docDefaults', ns)
idd = iroot.find('w:docDefaults', ns)
if tdd is not None and idd is not None:
    idd.addprevious(copy.deepcopy(tdd))
    iroot.remove(idd)

istyles.write(os.path.join(BUILD, 'word', 'styles.xml'),
              xml_declaration=True, encoding='UTF-8', standalone=True)
print(f'styles swapped from template: {len(swapped)} -> {sorted(swapped)}')

# ---------------------------------------------------------------- 2. theme
shutil.copy(os.path.join(TPL, 'word', 'theme', 'theme1.xml'),
            os.path.join(BUILD, 'word', 'theme', 'theme1.xml'))

# ---------------------------------------------------------------- 3. recolour
def recolour(path, strip_heading_colour=True):
    """Map the teal family onto the template palette; let heading styles govern
    heading colour by removing direct colour overrides on heading runs."""
    tree = etree.parse(path)
    root = tree.getroot()
    n_map = n_strip = 0
    for p in root.iter(q('w:p')):
        ps = p.find('w:pPr/w:pStyle', ns)
        sty = ps.get(q('w:val')) if ps is not None else ''
        is_heading = bool(HEADINGISH.match(sty or ''))
        for r in p.iterfind('.//w:r', ns):
            rpr = r.find('w:rPr', ns)
            if rpr is None:
                continue
            c = rpr.find('w:color', ns)
            if c is None:
                continue
            if is_heading and strip_heading_colour:
                rpr.remove(c); n_strip += 1
            else:
                v = (c.get(q('w:val')) or '').upper()
                if v in PALETTE:
                    c.set(q('w:val'), PALETTE[v]); n_map += 1
    # paragraph-mark colours and any other colour element (shading, borders)
    for el in root.iter():
        for attr in ('w:val', 'w:fill', 'w:color'):
            v = el.get(q(attr))
            if v and v.upper() in PALETTE:
                el.set(q(attr), PALETTE[v.upper()]); n_map += 1
    tree.write(path, xml_declaration=True, encoding='UTF-8', standalone=True)
    return n_map, n_strip

doc_path = os.path.join(BUILD, 'word', 'document.xml')
m, s_ = recolour(doc_path)
print(f'document.xml: {m} colours re-paletted, {s_} heading overrides stripped')

import glob
hf = 0
for f in glob.glob(os.path.join(BUILD, 'word', 'header*.xml')) + \
         glob.glob(os.path.join(BUILD, 'word', 'footer*.xml')):
    a, b = recolour(f, strip_heading_colour=False)   # keep banner styling, recolour only
    hf += a
print(f'headers/footers: {hf} colours re-paletted')

# ---------------------------------------------------------------- 4. front matter
tdoc = etree.parse(os.path.join(TPL, 'word', 'document.xml'))
tblocks = list(tdoc.getroot().find('w:body', ns))
gov_heading, gov_table = tblocks[25], tblocks[27]        # 'Document Governance' + table
hist_heading, hist_table = tblocks[29], tblocks[31]      # 'Document History' + table

bdoc = etree.parse(doc_path)
bbody = bdoc.getroot().find('w:body', ns)
bblocks = list(bbody)

def page_break_para():
    p = E('w:p')
    r = etree.SubElement(p, q('w:r'))
    etree.SubElement(r, q('w:br')).set(q('w:type'), 'page')
    return p

def widen(tbl, target=13000):
    """The template tables are sized for portrait; give them a landscape width."""
    grid = tbl.findall('w:tblGrid/w:gridCol', ns)
    widths = [int(g.get(q('w:w'))) for g in grid]
    total = sum(widths) or 1
    factor = target / total
    new = [round(x * factor) for x in widths]
    new[-1] += target - sum(new)
    for g, nw in zip(grid, new):
        g.set(q('w:w'), str(nw))
    tw = tbl.find('w:tblPr/w:tblW', ns)
    if tw is not None:
        tw.set(q('w:w'), str(target)); tw.set(q('w:type'), 'dxa')
    for tc in tbl.iterfind('.//w:tc', ns):
        cw = tc.find('w:tcPr/w:tcW', ns)
        if cw is not None and cw.get(q('w:type')) == 'dxa':
            cw.set(q('w:w'), str(round(int(cw.get(q('w:w'))) * factor)))
    return tbl

# the front matter needs its OWN single-column section: the contents section
# that follows is two-column, and dropping tables into it clips them and breaks
# the running text. Section [53] is the document's single-column landscape setup.
single_col = copy.deepcopy(bblocks[53].find('w:pPr/w:sectPr', ns))
t = single_col.find('w:type', ns)
if t is None:
    t = E('w:type'); single_col.insert(0, t)
t.set(q('w:val'), 'nextPage')          # start the block on a fresh page

closer = E('w:p')
cppr = etree.SubElement(closer, q('w:pPr'))
cppr.append(single_col)                # this paragraph closes my new section

CONTENT_W = 16840 - 709 - 709          # landscape page less its margins
front = [copy.deepcopy(gov_heading), widen(copy.deepcopy(gov_table), 14000),
         E('w:p'),
         copy.deepcopy(hist_heading), widen(copy.deepcopy(hist_table), 14000),
         closer]

# insert after the paragraph at [038], which carries a sectPr and so closes the
# contact-page section - the new blocks then form a section of their own
anchor = bblocks[38]
for el in reversed(front):
    anchor.addnext(el)

bdoc.write(doc_path, xml_declaration=True, encoding='UTF-8', standalone=True)
print('inserted Document Governance + Document History before the contents page')

# ---------------------------------------------------------------- 5. doc props
core_path = os.path.join(BUILD, 'docProps', 'core.xml')
core = open(core_path).read()
core = re.sub(r'<dc:title>[^<]*</dc:title>', f'<dc:title>{TITLE}</dc:title>', core)
if '<dc:title>' not in core:
    core = core.replace('</cp:coreProperties>', f'<dc:title>{TITLE}</dc:title></cp:coreProperties>')
open(core_path, 'w').write(core)

# ---------------------------------------------------------------- zip
out_docx = os.path.join(WORK, 'out7.docx')
if os.path.exists(out_docx):
    os.remove(out_docx)
os.system(f'cd "{BUILD}" && zip -qXr "{out_docx}" .')
print('wrote', out_docx, os.path.getsize(out_docx), 'bytes')
