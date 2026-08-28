#!/usr/bin/env python3
"""Dump a readable outline of word/document.xml: blocks with style, numbering, text."""
import sys, re
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
ns = {'w': W}

def para_info(p):
    style = p.find('w:pPr/w:pStyle', ns)
    style = style.get(f'{{{W}}}val') if style is not None else ''
    numpr = p.find('w:pPr/w:numPr', ns)
    num = ''
    if numpr is not None:
        ilvl = numpr.find('w:ilvl', ns)
        numid = numpr.find('w:numId', ns)
        ilvl = ilvl.get(f'{{{W}}}val') if ilvl is not None else '?'
        numid = numid.get(f'{{{W}}}val') if numid is not None else '?'
        num = f'num={numid}/l{ilvl}'
    texts = []
    for t in p.iter(f'{{{W}}}t'):
        texts.append(t.text or '')
    for _ in p.iter(f'{{{W}}}drawing'):
        texts.append('[DRAWING]')
    for _ in p.iter(f'{{{W}}}pict'):
        texts.append('[PICT]')
    for fld in p.iter(f'{{{W}}}instrText'):
        texts.append(f'[FLD:{(fld.text or "").strip()}]')
    sect = '[SECTPR]' if p.find('w:pPr/w:sectPr', ns) is not None else ''
    brk = ''
    for b in p.iter(f'{{{W}}}br'):
        if b.get(f'{{{W}}}type') == 'page':
            brk = '[PAGEBRK]'
    return style, num, ''.join(texts), sect, brk

def cell_text(tc):
    parts = []
    for p in tc.iterfind('.//w:p', ns):
        s, n, t, _, _ = para_info(p)
        parts.append(t)
    return ' / '.join(x for x in parts if x)

def main(path, maxtext=200):
    tree = etree.parse(path)
    body = tree.getroot().find('w:body', ns)
    i = 0
    for el in body:
        tag = etree.QName(el).localname
        if tag == 'p':
            style, num, text, sect, brk = para_info(el)
            flags = ' '.join(x for x in [style, num, sect, brk] if x)
            print(f'[{i:03d}] P {flags}: {text[:maxtext]!r}')
        elif tag == 'tbl':
            rows = el.findall('w:tr', ns)
            tblstyle = el.find('w:tblPr/w:tblStyle', ns)
            tblstyle = tblstyle.get(f'{{{W}}}val') if tblstyle is not None else ''
            print(f'[{i:03d}] TBL style={tblstyle} rows={len(rows)}')
            for ri, tr in enumerate(rows[:60]):
                cells = [cell_text(tc)[:70] for tc in tr.findall('w:tc', ns)]
                print(f'      r{ri}: {cells}')
        elif tag == 'sectPr':
            print(f'[{i:03d}] SECTPR')
        else:
            print(f'[{i:03d}] {tag}')
        i += 1

if __name__ == '__main__':
    main(sys.argv[1], int(sys.argv[2]) if len(sys.argv) > 2 else 200)
