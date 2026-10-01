#!/usr/bin/env python3
"""Fidelity check for a PDF-sourced reskin: every word of the source body must
appear in the reskinned body, and vice versa.

Compares word multisets (table cells read column-wise in a PDF, row-wise in
Word, so order is not comparable) after normalising quotes, dashes, bullets and
line-break hyphenation. Running headers/footers are dropped by position, the
source's literal heading numbers by pattern.

Usage: _verify_pdf_reskin.py source.pdf first_body_page last_body_page out.docx
       (pages 1-based, inclusive)
"""
import re, sys, zipfile
from collections import Counter
import pymupdf
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
ns = {'w': W}


def norm_words(text):
    text = (text.replace('’', "'").replace('‘', "'").replace('“', '"').replace('”', '"')
                .replace('–', '-').replace('—', '-').replace('•', ' ').replace(' ', ' '))
    text = re.sub(r'-\s*\n\s*', '-', text)          # line-end hyphen stays a hyphen
    return [w for w in text.split() if w]


def pdf_words(path, first, last):
    doc = pymupdf.open(path)
    lines = []
    for pno in range(first - 1, last):
        pg = doc[pno]
        h = pg.rect.height
        for b in pg.get_text('dict')['blocks']:
            for ln in b.get('lines', []):
                y = ln['bbox'][1]
                if y < 0.105 * h or y > 0.925 * h:     # running header / footer band
                    continue
                t = ''.join(sp['text'] for sp in ln['spans'])
                if ln['spans'] and max(sp['size'] for sp in ln['spans']) >= 19:
                    t = re.sub(r'^\s*\d+\.\s*', '', t)   # literal heading number
                t = re.sub(r'^\s*\d+(\.\d+)+\.?\s+', '', t)  # 4.1. sub-heading number
                lines.append(t)
    return norm_words('\n'.join(lines))


def docx_words(path):
    z = zipfile.ZipFile(path)
    root = etree.fromstring(z.read('word/document.xml'))
    body = root.find('w:body', ns)
    started, chunks = False, []
    for p in body.iter(f'{{{W}}}p'):
        ps = p.find('w:pPr/w:pStyle', ns)
        if ps is not None and ps.get(f'{{{W}}}val') == 'Heading1':
            started = True
        if not started:
            continue
        parts = []
        for el in p.iter(f'{{{W}}}t', f'{{{W}}}tab', f'{{{W}}}br'):
            parts.append(el.text or '' if el.tag.endswith('}t') else ' ')
        chunks.append(''.join(parts))
    fn = etree.fromstring(z.read('word/footnotes.xml'))
    for f in fn.findall('w:footnote', ns):
        if int(f.get(f'{{{W}}}id')) > 1:
            chunks.append(''.join(t.text or '' for t in f.iter(f'{{{W}}}t')))
    return norm_words('\n'.join(chunks))


if __name__ == '__main__':
    src, first, last, out = sys.argv[1], int(sys.argv[2]), int(sys.argv[3]), sys.argv[4]
    a, b = Counter(pdf_words(src, first, last)), Counter(docx_words(out))
    missing, extra = a - b, b - a
    print(f'source words {sum(a.values())}, output words {sum(b.values())}')
    print('in source, not in output:', dict(missing) or 'none')
    print('in output, not in source:', dict(extra) or 'none')
