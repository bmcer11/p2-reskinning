#!/usr/bin/env python3
"""Verify every source text block (body, from block 45 on) survives into the
portrait build, in order. The Contents control (block 44) is checked separately
because its cached page numbers are deliberately re-pointed."""
from lxml import etree
import copy, re

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
ns = {'w': W}
q = lambda t: f'{{{W}}}{t}'
BODY_START, SDT_TOC = 45, 44

def choice_only(el):
    """Text lives twice in the source's shapes (DrawingML choice + VML
    fallback); the build keeps the choice, so compare against that."""
    el = copy.deepcopy(el)
    etree.strip_elements(el, f'{{{MC}}}Fallback', with_tail=False)
    return el

def norm(s):
    return re.sub(r'\s+', ' ', s).replace(' ', ' ').strip()

def block_texts(el):
    """Flatten one top-level block into its non-empty paragraph texts."""
    out = []
    tag = etree.QName(el).localname
    if tag == 'p':
        t = norm(''.join(x.text or '' for x in el.iter(q('t'))))
        if t:
            out.append(t)
    elif tag in ('tbl', 'sdt'):
        for p in el.iterfind('.//w:p', ns):
            t = norm(''.join(x.text or '' for x in p.iter(q('t'))))
            if t:
                out.append(t)
    return out

ibody = etree.parse('instrument_unpacked/word/document.xml').getroot().find('w:body', ns)
iblocks = list(ibody)
src = []
for el in iblocks[BODY_START:]:
    src += block_texts(choice_only(el))

bbody = etree.parse('build8/word/document.xml').getroot().find('w:body', ns)
bblocks = list(bbody)
sdt_i = next(i for i, el in enumerate(bblocks)
             if etree.QName(el).localname == 'sdt'
             and any(ps.get(q('val')) in ('TOC1', 'TOC2')
                     for ps in el.findall('.//w:p/w:pPr/w:pStyle', ns)))
dst = []
for el in bblocks[sdt_i + 1:]:
    dst += block_texts(el)
joined = ''.join(dst)

pos, missing, moved = 0, [], []
for t in src:
    i = joined.find(t, pos)
    if i < 0:
        (moved if joined.find(t) >= 0 else missing).append(t)
    else:
        pos = i + len(t)

print(f'source body blocks : {len(src)}')
print(f'output body blocks : {len(dst)}')
print(f'missing            : {len(missing)}')
for m in missing[:15]:
    print('   MISSING', repr(m[:95]))
print(f'out of order       : {len(moved)}')
for m in moved[:15]:
    print('   MOVED  ', repr(m[:95]))

# Contents control: titles must match the source's entries one-for-one
def sdt_entries(blocks, idx=None, sdt_el=None):
    el = blocks[idx] if idx is not None else sdt_el
    out = []
    for p in el.iterfind('.//w:p', ns):
        ps = p.find('w:pPr/w:pStyle', ns)
        if ps is None or ps.get(q('val')) not in ('TOC1', 'TOC2', 'TOC3'):
            continue
        wts = [t.text for t in p.iter(q('t')) if t.text]
        if wts and wts[-1].strip().isdigit():
            out.append((ps.get(q('val')), norm(''.join(wts[:-1]))))
    return out

src_toc = sdt_entries(iblocks, idx=SDT_TOC)
dst_toc = sdt_entries(bblocks, sdt_el=bblocks[sdt_i])
print(f'ToC entries        : source {len(src_toc)}, output {len(dst_toc)}, '
      f'{"MATCH" if src_toc == dst_toc else "DIFFER"}')
if src_toc != dst_toc:
    for a, b in zip(src_toc, dst_toc):
        if a != b:
            print('   ', a, '->', b)

# heading style census: every body paragraph keeps its source style
def style_seq(blocks):
    seq = []
    for el in blocks:
        for p in ([el] if etree.QName(el).localname == 'p' else el.iterfind('.//w:p', ns)):
            ps = p.find('w:pPr/w:pStyle', ns)
            sty = ps.get(q('val')) if ps is not None else 'Normal'
            t = norm(''.join(x.text or '' for x in p.iter(q('t'))))
            if t and re.match(r'^(Heading[1-9]|Style1)', sty):
                seq.append((sty, t[:60]))
    return seq

hsrc = style_seq([choice_only(el) for el in iblocks[BODY_START:]])
hdst = style_seq(bblocks[sdt_i + 1:])
print(f'headings           : source {len(hsrc)}, output {len(hdst)}, '
      f'{"MATCH" if hsrc == hdst else "DIFFER"}')
if hsrc != hdst:
    shown = 0
    for k in range(max(len(hsrc), len(hdst))):
        a = hsrc[k] if k < len(hsrc) else None
        b = hdst[k] if k < len(hdst) else None
        if a != b:
            print(f'   [{k}]', a, '->', b)
            shown += 1
            if shown >= 8:
                break
