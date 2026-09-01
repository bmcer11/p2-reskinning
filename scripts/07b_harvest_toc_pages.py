#!/usr/bin/env python3
"""Harvest the ToC page numbers that LibreOffice resolved from the PAGEREF
bookmarks in out8.pdf, and cache them into toc_pages8.json so the next
build8.py run bakes them in as the entries' cached field results."""
import fitz, json, re, sys
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
ns = {'w': W}
q = lambda t: f'{{{W}}}{t}'
norm = lambda s: re.sub(r'\s+', ' ', s).strip()

root = etree.parse('build8/word/document.xml').getroot()
titles = []
for sdt in root.find('w:body', ns).iterfind('w:sdt', ns):
    for p in sdt.iterfind('.//w:p', ns):
        ps = p.find('w:pPr/w:pStyle', ns)
        if ps is None or ps.get(q('val')) not in ('TOC1', 'TOC2', 'TOC3'):
            continue
        wts = [t.text for t in p.iter(q('t')) if t.text]
        if wts and wts[-1].strip().isdigit():
            titles.append(''.join(wts[:-1]).strip())
    if titles:
        break

doc = fitz.open('out8.pdf')
toc_pdf_pages = [i for i, pg in enumerate(doc)
                 if 'Table of Contents' in pg.get_text()
                 or 'PAGEREF' in pg.get_text()]
# the ToC occupies the pages right after the front matter; find them by content
lines = []
for i, pg in enumerate(doc):
    t = pg.get_text()
    if 'Acknowledgement of Country' in t and 'Roles and Powers' in t:
        lines += t.splitlines()
        if i + 1 < len(doc):
            lines += doc[i + 1].get_text().splitlines()
        break

pages, li = {}, 0
for t in titles:
    tt, found, j = norm(t), False, li
    while j < len(lines):
        s = norm(lines[j])
        if s == tt:                       # bare title; number on a later line
            k = j + 1
            while k < len(lines):
                s2 = norm(lines[k])
                if s2.isdigit():
                    pages[t] = int(s2); li = k + 1; found = True
                    break
                if s2 and set(s2) - set('. '):   # skip blank and dot-leader lines
                    break
                k += 1
            break
        m = re.match(re.escape(tt) + r'\s+(\d+)$', s)   # glued 'Title N'
        if m:
            pages[t] = int(m.group(1)); li = j + 1; found = True
            break
        j += 1
    if not found:
        print('MISS:', repr(t))

print(f'harvested {len(pages)}/{len(titles)}')
if len(pages) == len(titles):
    json.dump(pages, open('toc_pages8.json', 'w'), indent=1)
    print('toc_pages8.json written')
else:
    sys.exit(1)
