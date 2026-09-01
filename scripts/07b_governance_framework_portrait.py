#!/usr/bin/env python3
"""Reskin the CoPP Governance Framework into the Framework template - PORTRAIT.

Second pass on instrument 7, at the requester's direction: the standard
portrait reflow like the other six instruments, replacing the earlier
keep-the-landscape-design build.

- Base package is the TEMPLATE: its cover (FRAMEWORK chip), contact page,
  Document Governance / Document History tables, headers and footers.
- The source's 27 landscape sections and two-column layouts are dissolved;
  content flows single-column portrait in reading order (Word column order:
  left column then right column per section).
- The author's paragraph styles - and therefore her heading levels H1-H5 -
  travel with each paragraph. The template's Heading styles supply the look;
  their automatic clause numbering is stripped because this framework uses
  named sections, and instrument-only styles (Style1, ListTable4-Accent5) are
  imported.
- Her curated Contents control is kept verbatim (with its bookmarks, so
  internal "back to contents" links still work); cached page numbers are
  re-pointed at the portrait layout in a second pass.
- The 29-row x 10-column Principles Matrix cannot survive a portrait squeeze,
  so that one section is placed on landscape pages; everything else is
  portrait. The two 2-column tables scale to the template's table width.
- Teal accents map to the template's magenta/plum; heading runs drop their
  direct colours so the template styles govern.
- No version is stated anywhere in the source, so the cover keeps the
  template's "Version X.X" placeholder for the author.
"""
import copy, glob, json, os, re, shutil, sys
from lxml import etree

WORK = os.path.dirname(os.path.abspath(__file__))
TPL = os.path.join(WORK, 'template_unpacked')
INS = os.path.join(WORK, 'instrument_unpacked')
BUILD = os.path.join(WORK, 'build8')

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
ns = {'w': W, 'r': R, 'a': A, 'wp': WP}

TITLE = 'City of Port Phillip Governance Framework'

PALETTE = {  # source teal family -> template magenta
    '2EBBB8': '9D2270', '008080': '9D2270', '009999': '9D2270',
    '3E988D': '9D2270', '007184': '9D2270', '2BBBB0': '9D2270',
    '00A3AD': '9D2270', '5DA5AF': '9D2270',
}
HEADINGISH = re.compile(r'^(Heading[1-9]|Style1)')
MAXW_EMU = 5731510            # portrait content width 9026 twips, in EMU
PORTRAIT_TBL = 9770           # the template family's own table width
LANDSCAPE_TBL = 15945         # the matrix's own width - it slightly overflows the text area, as in the source

MATRIX_FIRST, MATRIX_LAST = 248, 256    # Principles Matrix heading .. section tail
BODY_START = 45                         # source body after its old front matter
TOC_BOOKMARKS = range(39, 44)           # _Contents / _Toc anchors before her sdt
SDT_TOC = 44

def q(tag):
    p, local = tag.split(':')
    return f'{{{ {"w": W, "r": R, "a": A, "wp": WP}[p] }}}{local}'

def E(tag, attrib=None):
    el = etree.Element(q(tag))
    for k, v in (attrib or {}).items():
        el.set(q(k), v)
    return el

# ---------------------------------------------------------------- base package
shutil.rmtree(BUILD, ignore_errors=True)
shutil.copytree(TPL, BUILD)

tdoc = etree.parse(os.path.join(BUILD, 'word', 'document.xml'))
tbody = tdoc.getroot().find('w:body', ns)
tblocks = list(tbody)

idoc = etree.parse(os.path.join(INS, 'word', 'document.xml'))
iblocks = list(idoc.getroot().find('w:body', ns))

# ---------------------------------------------------------------- styles
tstyles = etree.parse(os.path.join(BUILD, 'word', 'styles.xml'))
sroot = tstyles.getroot()
have = {s.get(q('w:styleId')) for s in sroot.iterfind('w:style', ns)}

# headings unnumbered: this framework uses named sections
for s in sroot.iterfind('w:style', ns):
    sid = s.get(q('w:styleId')) or ''
    if re.match(r'^Heading[1-5]$', sid):
        ppr = s.find('w:pPr', ns)
        if ppr is not None:
            for el in ppr.findall('w:numPr', ns) + ppr.findall('w:ind', ns):
                ppr.remove(el)

# import instrument-only styles the content references
istyles = etree.parse(os.path.join(INS, 'word', 'styles.xml'))
S = {s.get(q('w:styleId')): s for s in istyles.getroot().iterfind('w:style', ns)}
used = set()
for el in iblocks[BODY_START:] + [iblocks[SDT_TOC]]:
    for x in el.iter(q('w:pStyle')):
        used.add(x.get(q('w:val')))
    for x in el.iter(q('w:rStyle')):
        used.add(x.get(q('w:val')))
    for x in el.iter(q('w:tblStyle')):
        used.add(x.get(q('w:val')))
missing = sorted(u for u in used if u and u not in have)
for sid in list(missing):
    for extra in (sid + 'Char',):        # bring linked character styles along
        if extra in S and extra not in have and extra not in missing:
            missing.append(extra)
for sid in missing:
    new = copy.deepcopy(S[sid])
    base = new.find('w:basedOn', ns)
    if HEADINGISH.match(sid) or (base is not None and
                                 (base.get(q('w:val')) or '').startswith('Heading')):
        rpr = new.find('w:rPr', ns)
        if rpr is not None:              # inherit the template heading colour
            for c in rpr.findall('w:color', ns):
                rpr.remove(c)
        ppr = new.find('w:pPr', ns)
        if ppr is not None:
            for el in ppr.findall('w:numPr', ns):
                ppr.remove(el)
    sroot.append(new)
print('imported styles:', missing)
tstyles.write(os.path.join(BUILD, 'word', 'styles.xml'),
              xml_declaration=True, encoding='UTF-8', standalone=True)

# ---------------------------------------------------------------- numbering
inum = etree.parse(os.path.join(INS, 'word', 'numbering.xml'))
inum_root = inum.getroot()
i_num2abs = {n.get(q('w:numId')): n.find('w:abstractNumId', ns).get(q('w:val'))
             for n in inum_root.iterfind('w:num', ns)}
i_abs = {a.get(q('w:abstractNumId')): a
         for a in inum_root.iterfind('w:abstractNum', ns)}

used_nids = set()
for el in iblocks[BODY_START:]:
    for nid in el.iterfind('.//w:numPr/w:numId', ns):
        v = nid.get(q('w:val'))
        if v and v != '0':
            used_nids.add(v)

tnum = etree.parse(os.path.join(BUILD, 'word', 'numbering.xml'))
tnum_root = tnum.getroot()
first_num = tnum_root.find('w:num', ns)
NUMMAP, abs_map = {}, {}
next_abs, next_num = 200, 200
for nid in sorted(used_nids, key=int):
    aid = i_num2abs.get(nid)
    if aid is None:
        continue
    if aid not in abs_map:                    # copy each abstract once
        src = copy.deepcopy(i_abs[aid])
        src.set(q('w:abstractNumId'), str(next_abs))
        for ps in src.iterfind('.//w:pStyle', ns):   # no style links across packages
            ps.getparent().remove(ps)
        first_num.addprevious(src)
        abs_map[aid] = str(next_abs)
        next_abs += 1
    num_el = E('w:num', {'w:numId': str(next_num)})
    etree.SubElement(num_el, q('w:abstractNumId')).set(q('w:val'), abs_map[aid])
    tnum_root.findall('w:num', ns)[-1].addnext(num_el)
    NUMMAP[nid] = str(next_num)
    next_num += 1
print(f'numbering: {len(NUMMAP)} lists imported over {len(abs_map)} abstracts')
tnum.write(os.path.join(BUILD, 'word', 'numbering.xml'),
           xml_declaration=True, encoding='UTF-8', standalone=True)

# ---------------------------------------------------------------- rels + media
trels_path = os.path.join(BUILD, 'word', '_rels', 'document.xml.rels')
trels = etree.parse(trels_path)
RELNS = 'http://schemas.openxmlformats.org/package/2006/relationships'
rel_root = trels.getroot()
irels = etree.parse(os.path.join(INS, 'word', '_rels', 'document.xml.rels'))
i_rel = {rel.get('Id'): rel for rel in irels.getroot()}

used_rids = set()
for el in iblocks[BODY_START:] + [iblocks[SDT_TOC]]:
    x = etree.tostring(el).decode()
    used_rids |= set(re.findall(r'r:(?:embed|id|link)="(rId\d+)"', x))

RELMAP = {}
next_rid = 300
media_n = 0
for rid in sorted(used_rids, key=lambda s: int(s[3:])):
    src = i_rel.get(rid)
    if src is None:
        continue
    rtype, target = src.get('Type'), src.get('Target')
    if not (rtype.endswith('/image') or rtype.endswith('/hyperlink')):
        continue                # the source's own headers/footers are dropped
    new_id = f'rId{next_rid}'; next_rid += 1
    el = etree.SubElement(rel_root, f'{{{RELNS}}}Relationship')
    el.set('Id', new_id); el.set('Type', rtype)
    if rtype.endswith('/image'):
        media_n += 1
        ext = os.path.splitext(target)[1]
        newname = f'media/imageF{media_n:02d}{ext}'
        shutil.copy(os.path.join(INS, 'word', target),
                    os.path.join(BUILD, 'word', newname))
        el.set('Target', newname)
    else:
        el.set('Target', target)
        if src.get('TargetMode'):
            el.set('TargetMode', src.get('TargetMode'))
    RELMAP[rid] = new_id
print(f'rels: {len(RELMAP)} imported ({media_n} images)')
trels.write(trels_path, xml_declaration=True, encoding='UTF-8', standalone=True)

# .wdp (HD photo) sits in source media but only as a fallback inside one image;
# copy it blindly if referenced
ct_path = os.path.join(BUILD, '[Content_Types].xml')
ct = open(ct_path).read()
ct = ct.replace('application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml',
                'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml')
open(ct_path, 'w').write(ct)

# ---------------------------------------------------------------- transforms
def strip_ids(el):
    for e in el.iter():
        for k in list(e.attrib):
            if k.endswith('}paraId') or k.endswith('}textId') or 'rsid' in k:
                del e.attrib[k]
    etree.strip_elements(el, q('w:proofErr'), with_tail=False)
    etree.strip_elements(el, q('w:lastRenderedPageBreak'), with_tail=False)

def remap(el):
    for nid in el.iterfind('.//w:numPr/w:numId', ns):
        v = nid.get(q('w:val'))
        if v in NUMMAP:
            nid.set(q('w:val'), NUMMAP[v])
    x_attrs = [f'{{{R}}}embed', f'{{{R}}}id', f'{{{R}}}link']
    for e in el.iter():
        for a in x_attrs:
            v = e.get(a)
            if v in RELMAP:
                e.set(a, RELMAP[v])

def recolour(el):
    """Teal -> magenta on ordinary runs; heading runs drop their colour."""
    for p in ([el] if etree.QName(el).localname == 'p' else el.iter(q('w:p'))):
        ps = p.find('w:pPr/w:pStyle', ns)
        is_heading = bool(HEADINGISH.match((ps.get(q('w:val')) if ps is not None else '') or ''))
        for r in p.iterfind('.//w:r', ns):
            rpr = r.find('w:rPr', ns)
            if rpr is None:
                continue
            c = rpr.find('w:color', ns)
            if c is None:
                continue
            if is_heading:
                rpr.remove(c)
            else:
                v = (c.get(q('w:val')) or '').upper()
                if v in PALETTE:
                    c.set(q('w:val'), PALETTE[v])
    for e in el.iter():
        for a in ('w:val', 'w:fill', 'w:color'):
            v = e.get(q(a))
            if v and v.upper() in PALETTE:
                e.set(q(a), PALETTE[v.upper()])
    for sc in el.iter(f'{{{A}}}srgbClr'):    # DrawingML shape fills (the house
        v = (sc.get('val') or '').upper()    # diagram etc.) live in a:, not w:
        if v in PALETTE:
            sc.set('val', PALETTE[v])

def cap_images(el, maxw=MAXW_EMU):
    for tag in (f'{{{WP}}}extent', f'{{{A}}}ext'):
        for ext in el.iter(tag):
            cx, cy = int(ext.get('cx') or 0), int(ext.get('cy') or 0)
            if cx > maxw:
                ext.set('cx', str(maxw)); ext.set('cy', str(round(cy * maxw / cx)))

def drop_column_indents(el):
    """72 body paragraphs carry big direct right indents (3000-5200 twips) that
    squeezed text into the landscape design's narrow columns; in the portrait
    flow they just pinch paragraphs to half width, so drop them. Small insets
    (quotes etc.) stay."""
    for ind in el.iter(q('w:ind')):
        r = ind.get(q('w:right'))
        if r and int(r) > 700:
            del ind.attrib[q('w:right')]

# ---------------------------------------------------------- floating shapes
# The source's floating shapes are positioned for its landscape spreads: the
# 24 "OUR PRINCIPLES" side badges sit at x~8M EMU (off a portrait page - they
# rendered nowhere), and the Governance Framework Diagram's five shapes and the
# principles wheel land on top of the reflowed text. Re-seat each against its
# own paragraph and make it reserve space.
MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'

def modernize(el):
    """Drop the VML fallback so a repositioned anchor cannot disagree with a
    stale fallback position; every consumer since Word 2007 uses the choice."""
    for ac in list(el.iter(f'{{{MC}}}AlternateContent')):
        choice = ac.find(f'{{{MC}}}Choice')
        drawing = choice.find(q('w:drawing')) if choice is not None else None
        if drawing is not None:
            ac.getparent().replace(ac, drawing)

def repos(el, x=None, y=0, halign=None, wrap='topAndBottom', behind=None):
    for a in el.iter(f'{{{WP}}}anchor'):
        ph = a.find(f'{{{WP}}}positionH')
        for c in list(ph):
            ph.remove(c)
        if halign is not None:
            ph.set('relativeFrom', 'margin')
            etree.SubElement(ph, f'{{{WP}}}align').text = halign
        else:
            ph.set('relativeFrom', 'column')
            etree.SubElement(ph, f'{{{WP}}}posOffset').text = str(int(x))
        pv = a.find(f'{{{WP}}}positionV')
        for c in list(pv):
            pv.remove(c)
        pv.set('relativeFrom', 'paragraph')
        etree.SubElement(pv, f'{{{WP}}}posOffset').text = str(int(y))
        old = [c for c in a if etree.QName(c).localname.startswith('wrap')]
        if wrap == 'topAndBottom':
            nw = etree.Element(f'{{{WP}}}wrapTopAndBottom',
                               {'distT': '91440', 'distB': '91440'})
        else:                                   # 'square-left': text wraps left
            nw = etree.Element(f'{{{WP}}}wrapSquare',
                               {'wrapText': 'left', 'distL': '114300'})
        a.replace(old[0], nw)
        for extra in old[1:]:
            a.remove(extra)
        if behind is not None:
            a.set('behindDoc', behind)

# Governance Framework Diagram, restacked for portrait: house centred (roof /
# Council / CEO), the two arrow-callouts side by side beneath it, arrows
# converging. Offsets in EMU within the 5731510 EMU portrait column.
HOUSE_HOST = 156                               # left callout paragraph hosts all
HOUSE_XY = {156: (0, 1913890),                 # rightArrowCallout (community)
            157: (1303973, 0),                 # roof triangle
            161: (1318895, 1005840),           # Council bar
            163: (1293178, 1389380),           # CEO / Organisation bar
            160: (2886075, 1913890)}           # leftArrowCallout (legislative)
EMPLOYEES_BOX, WHEEL = 99, 221

def scale_table(tbl, target):
    grid = tbl.findall('w:tblGrid/w:gridCol', ns)
    widths = [int(g.get(q('w:w'))) for g in grid]
    total = sum(widths) or 1
    f = target / total
    new = [round(x * f) for x in widths]
    new[-1] += target - sum(new)
    for g, nw in zip(grid, new):
        g.set(q('w:w'), str(nw))
    tw = tbl.find('w:tblPr/w:tblW', ns)
    if tw is None:
        tw = E('w:tblW'); tbl.find('w:tblPr', ns).insert(0, tw)
    tw.set(q('w:w'), str(target)); tw.set(q('w:type'), 'dxa')
    for tc in tbl.iterfind('.//w:tc', ns):
        cw = tc.find('w:tcPr/w:tcW', ns)
        if cw is not None and cw.get(q('w:type')) == 'dxa':
            cw.set(q('w:w'), str(round(int(cw.get(q('w:w'))) * f)))

AFTER_SZ = {'szCs', 'highlight', 'u', 'effect', 'bdr', 'shd', 'fitText',
            'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout',
            'specVanish', 'oMath'}

def pin_run_size(tbl, half_points):
    """The source set this table at its Normal's 10pt; the template's Normal is
    11pt, which makes the matrix header words wrap mid-word. Pin the source
    size on every run that doesn't set its own."""
    for r in tbl.iterfind('.//w:r', ns):
        rpr = r.find('w:rPr', ns)
        if rpr is None:
            rpr = E('w:rPr'); r.insert(0, rpr)
        if rpr.find('w:sz', ns) is not None:
            continue
        sz = E('w:sz', {'w:val': str(half_points)})
        szcs = E('w:szCs', {'w:val': str(half_points)})
        anchor = None
        for el in rpr:
            if etree.QName(el).localname in AFTER_SZ:
                anchor = el
                break
        if anchor is not None:
            anchor.addprevious(sz)
        else:
            rpr.append(sz)
        sz.addnext(szcs)

def magenta_header_row(tbl):
    rows = tbl.findall('w:tr', ns)
    if not rows:
        return
    for tc in rows[0].findall('w:tc', ns):
        tcpr = tc.find('w:tcPr', ns)
        if tcpr is None:
            tcpr = E('w:tcPr'); tc.insert(0, tcpr)
        for old in tcpr.findall('w:shd', ns):
            tcpr.remove(old)
        shd = E('w:shd', {'w:val': 'clear', 'w:color': 'auto', 'w:fill': '9D2270'})
        anchor = None                    # shd sits after tcBorders in the schema
        for tag in ('w:tcW', 'w:gridSpan', 'w:hMerge', 'w:vMerge', 'w:tcBorders'):
            el2 = tcpr.find(tag, ns)
            if el2 is not None:
                anchor = el2
        (anchor.addnext(shd) if anchor is not None else tcpr.insert(0, shd))
        for r in tc.iterfind('.//w:r', ns):
            rpr = r.find('w:rPr', ns)
            if rpr is None:
                rpr = E('w:rPr'); r.insert(0, rpr)
            for tag in ('color', 'b', 'bCs'):
                for c in rpr.findall(f'w:{tag}', ns):
                    rpr.remove(c)
            rpr.append(E('w:b')); rpr.append(E('w:bCs'))
            rpr.append(E('w:color', {'w:val': 'FFFFFF'}))
    trpr = rows[0].find('w:trPr', ns)
    if trpr is None:
        trpr = E('w:trPr'); rows[0].insert(0, trpr)
    if trpr.find('w:tblHeader', ns) is None:
        trpr.append(E('w:tblHeader'))

def sect_closer(landscape=False, first=False):
    """Paragraph whose sectPr ends the section above it, cloned from the template."""
    sect = copy.deepcopy(tbody.find('w:sectPr', ns))
    pg = sect.find('w:pgSz', ns)
    if landscape:
        pg.set(q('w:w'), '16838'); pg.set(q('w:h'), '11906')
        pg.set(q('w:orient'), 'landscape')
        mar = sect.find('w:pgMar', ns)      # the source's own landscape margins,
        mar.set(q('w:left'), '709')         # so the Principles Matrix keeps its
        mar.set(q('w:right'), '709')        # column breathing room
    if not first:                     # numbering continues; no title page
        for tag in ('w:pgNumType', 'w:titlePg'):
            el = sect.find(tag, ns)
            if el is not None:
                sect.remove(el)
    p = E('w:p')
    ppr = etree.SubElement(p, q('w:pPr'))
    ppr.append(sect)
    return p

def page_break():
    p = E('w:p')
    r = etree.SubElement(p, q('w:r'))
    etree.SubElement(r, q('w:br')).set(q('w:type'), 'page')
    return p

# ---------------------------------------------------------------- body content
out = []
for i in range(BODY_START, len(iblocks)):
    el = iblocks[i]
    tag = etree.QName(el).localname
    if tag == 'sectPr':
        continue
    if i == MATRIX_FIRST:                       # portrait section ends here
        out.append(sect_closer(landscape=False, first=True))
    if i == MATRIX_LAST + 1:                    # landscape matrix section ends
        out.append(sect_closer(landscape=True))
    blk = copy.deepcopy(el)
    strip_ids(blk); remap(blk); recolour(blk); drop_column_indents(blk)
    in_matrix = MATRIX_FIRST <= i <= MATRIX_LAST
    cap_images(blk, LANDSCAPE_TBL * 635 if in_matrix else MAXW_EMU)
    if tag == 'tbl':
        if i == 252:                            # Principles Matrix
            pin_run_size(blk, 20)               # the source's 10pt, so the
            magenta_header_row(blk)             # header words fit as designed
            scale_table(blk, LANDSCAPE_TBL)
        elif i == 639:                          # Strategic direction / Vision
            magenta_header_row(blk)
            scale_table(blk, PORTRAIT_TBL)
        else:                                   # themed obligations table etc.
            scale_table(blk, PORTRAIT_TBL)
        out.append(blk)
        continue
    if tag != 'p':
        out.append(blk)                         # body-level bookmarks etc.
        continue
    ppr = blk.find('w:pPr', ns)
    if ppr is not None:                         # dissolve landscape sections
        for s in ppr.findall('w:sectPr', ns):
            ppr.remove(s)
    if blk.find(f'.//{{{WP}}}anchor') is not None:
        modernize(blk)
        text_all = ''.join(t.text or '' for t in blk.iter(q('w:t')))
        if i in (EMPLOYEES_BOX, WHEEL):
            repos(blk, halign='center', y=114300, behind='0')
        elif i == HOUSE_HOST:
            repos(blk, x=HOUSE_XY[i][0], y=HOUSE_XY[i][1])
            house_host = blk
        elif i in HOUSE_XY:
            repos(blk, x=HOUSE_XY[i][0], y=HOUSE_XY[i][1])
            for d in list(blk.iterfind('.//w:drawing', ns)):
                run = d.getparent()             # move the shape's whole run
                house_host.append(run)          # into the diagram's host para
            continue                            # drop the emptied paragraph
        elif text_all.startswith('OUR PRINCIPLES'):
            repos(blk, halign='right', y=0, wrap='square-left')
    text = ''.join(t.text or '' for t in blk.iter(q('w:t')))
    has_media = any(True for _ in blk.iter(q('w:drawing'))) or \
                any(True for _ in blk.iter(q('w:pict')))
    if not text.strip() and not has_media:
        # keep bookmark integrity: hoist any markers out of dropped spacers
        for bm in blk.iter(q('w:bookmarkStart')):
            out.append(copy.deepcopy(bm))
        for bm in blk.iter(q('w:bookmarkEnd')):
            out.append(copy.deepcopy(bm))
        continue
    out.append(blk)
print(f'{len(out)} body blocks')

# ---------------------------------------------------------------- assemble
for t in tblocks[0].iter(q('w:t')):             # cover: title in, version stays X.X
    if t.text == 'Framework Title':
        t.text = TITLE
sub = os.path.join(BUILD, 'word')
for fname, old in (('header1.xml', '&lt;Document title&gt;'),
                   ('footer1.xml', '&lt;Framework Title&gt;')):
    p = os.path.join(sub, fname)
    x = open(p).read().replace(old, TITLE)
    open(p, 'w').write(x)

set_p = os.path.join(BUILD, 'word', 'settings.xml')
st = open(set_p).read()
if '<w:updateFields' not in st:
    for anchor_tag in ('<w:hdrShapeDefaults', '<w:footnotePr', '<w:endnotePr', '<w:compat'):
        if anchor_tag in st:
            st = st.replace(anchor_tag, '<w:updateFields w:val="true"/>' + anchor_tag, 1)
            break
open(set_p, 'w').write(st)

# the template's webSettings carries optimizeForBrowser, under which the ToC
# entries' webHidden tab runs are treated as hidden and the dotted leaders
# vanish; it only matters for save-as-web-page, so drop it
web_p = os.path.join(BUILD, 'word', 'webSettings.xml')
if os.path.exists(web_p):
    wx = open(web_p).read().replace('<w:optimizeForBrowser/>', '')
    open(web_p, 'w').write(wx)

core_p = os.path.join(BUILD, 'docProps', 'core.xml')
core = open(core_p).read()
core = re.sub(r'<dc:title>[^<]*</dc:title>', f'<dc:title>{TITLE}</dc:title>', core)
if '<dc:title>' not in core:
    core = core.replace('</cp:coreProperties>', f'<dc:title>{TITLE}</dc:title></cp:coreProperties>')
open(core_p, 'w').write(core)

sdt = copy.deepcopy(iblocks[SDT_TOC])           # her curated Contents entries
strip_ids(sdt); remap(sdt); recolour(sdt)
# her 'Contents' heading carries a run-split rendering artifact from the source;
# the template's own 'Table of Contents' heading replaces it (as in jobs 1-6)
first_p = sdt.find('.//w:p', ns)
if first_p is not None and 'Contents' in ''.join(t.text or '' for t in first_p.iter(q('w:t'))):
    first_p.getparent().remove(first_p)

# entries carried a tab stop sized for the old narrow landscape column; use the
# template's dotted right tab, and re-point cached page numbers at this layout
toc_pages = {}
tp = os.path.join(WORK, 'toc_pages8.json')
if os.path.exists(tp):
    toc_pages = json.load(open(tp))
for p in sdt.iterfind('.//w:p', ns):
    ppr = p.find('w:pPr', ns)
    ps = ppr.find('w:pStyle', ns) if ppr is not None else None
    if ps is None or ps.get(q('w:val')) not in ('TOC1', 'TOC2', 'TOC3'):
        continue
    old_tabs = ppr.find('w:tabs', ns)
    if old_tabs is not None:
        ppr.remove(old_tabs)
    ntabs = E('w:tabs')
    t1 = etree.SubElement(ntabs, q('w:tab'))
    t1.set(q('w:val'), 'right'); t1.set(q('w:leader'), 'dot'); t1.set(q('w:pos'), '9026')
    ps.addnext(ntabs)
    wts = [t for t in p.iter(q('w:t')) if t.text]
    if wts and wts[-1].text.strip().isdigit():
        title = ''.join(t.text for t in wts[:-1]).strip()
        if title in toc_pages:
            wts[-1].text = str(toc_pages[title])
# LibreOffice treats the entries' webHidden tab runs as hidden, swallowing the
# dotted leaders; webHidden only affects Word's web layout view, so drop it
for wh in list(sdt.iter(q('w:webHidden'))):
    wh.getparent().remove(wh)
anchors = [copy.deepcopy(iblocks[i]) for i in TOC_BOOKMARKS]

final_sect = tbody.find('w:sectPr', ns)         # last section: numbering continues
for tag in ('w:pgNumType', 'w:titlePg'):
    el = final_sect.find(tag, ns)
    if el is not None:
        final_sect.remove(el)

keep = [tblocks[0]]                             # cover
keep += tblocks[2:25]                           # CoPP contact page (+ page break)
keep += [tblocks[25], tblocks[27],              # Document Governance
         tblocks[28], tblocks[29], tblocks[31]] # Document History
keep.append(page_break())
keep.append(tblocks[40])                        # template 'Table of Contents' heading
keep += anchors                                 # _Contents / _Toc bookmark anchors
keep.append(sdt)                                # her Contents entries
keep.append(page_break())
keep += out
keep.append(final_sect)

for c in list(tbody):
    tbody.remove(c)
for c in keep:
    tbody.append(c)
tdoc.write(os.path.join(BUILD, 'word', 'document.xml'),
           xml_declaration=True, encoding='UTF-8', standalone=True)
print('document.xml written')

out_docx = os.path.join(WORK, 'out8.docx')
if os.path.exists(out_docx):
    os.remove(out_docx)
os.system(f'cd "{BUILD}" && zip -qXr "{out_docx}" .')
print('wrote', out_docx, os.path.getsize(out_docx), 'bytes')
