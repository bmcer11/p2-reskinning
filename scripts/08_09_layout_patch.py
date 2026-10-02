#!/usr/bin/env python3
"""Layout patch for the shipped PDF-source reskins (instruments 8 and 9).

The first renders of these files lost content: LibreOffice 24.2 dropped rows
2-7 of the PID Procedure's 'Overview of the disclosure process' table (a
cantSplit table breaking between rows just above a Heading 1), and the ToC
page numbers had been harvested from that broken layout.

_pdf_reskin_lib.py now emits the fix itself: keep_together on that table, and
lead-in paragraphs ending in ':' kept with the table or list they introduce.
The template package a full rebuild needs was not available, so this script
applies the same two rules to the already-built .docx, then renders it and
re-bakes each ToC entry's cached PAGEREF result from the new layout. Text is
not touched; the zip member order of the package is preserved.

Usage: 08_09_layout_patch.py <in.docx> <out.docx>
"""
import os, sys, tempfile, zipfile
from lxml import etree
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _pdf_reskin_lib import keep_with_next, render_pdf, harvest_pages

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
ns = {'w': W}
q = lambda t: f'{{{W}}}{t}'
DOC = 'word/document.xml'


def text(el):
    return ''.join(t.text or '' for t in el.iter(q('t')))


def style(p):
    ps = p.find('w:pPr/w:pStyle', ns)
    return (ps.get(q('val')) or '') if ps is not None else ''


def body_start(blocks):
    """Index of the first Heading 1 after the ToC: the transcribed body."""
    last_toc = max(i for i, el in enumerate(blocks)
                   if el.tag == q('p') and style(el).upper().startswith('TOC'))
    return next(i for i in range(last_toc + 1, len(blocks))
                if blocks[i].tag == q('p') and style(blocks[i]) == 'Heading1')


def patch_layout(body):
    blocks = list(body)
    n_lead = n_rows = 0
    for i in range(body_start(blocks), len(blocks) - 1):
        el, nxt = blocks[i], blocks[i + 1]
        if el.tag != q('p') or style(el) or el.find('w:pPr/w:numPr', ns) is not None:
            continue                                  # plain body paragraphs only
        if not text(el).rstrip().endswith(':'):
            continue
        if nxt.tag == q('tbl') or (nxt.tag == q('p') and style(nxt) == 'ListParagraph'):
            keep_with_next(el)
            n_lead += 1
    for tbl in body.iterfind('w:tbl', ns):
        rows = tbl.findall('w:tr', ns)
        if [text(tc) for tc in rows[0].findall('w:tc', ns)] == ['Step', 'Action', 'Responsibility']:
            for tr in rows[:-1]:                      # keep_together, as the builder does
                for p in tr.iter(q('p')):
                    keep_with_next(p)
            n_rows = len(rows) - 1
    return n_lead, n_rows


def toc_entries(body):
    out = []
    for p in body.iterfind('w:p', ns):
        if not style(p).upper().startswith('TOC'):
            continue
        hl = p.find('w:hyperlink', ns)
        if hl is None:                                # the TOC field's closing paragraph
            continue
        wts = [t.text for t in hl.iter(q('t')) if t.text]
        out.append((p, hl.get(q('anchor')), ''.join(wts[1:-1])))   # 'N.', title, page
    return out


def set_cached_page(p, page):
    """Rewrite the PAGEREF field's result. The first entry also sits inside the
    outer TOC field, so only a 'separate' that follows a PAGEREF code counts."""
    seen_pageref = in_result = False
    for r in p.iter(q('r')):
        it = r.find('w:instrText', ns)
        if it is not None and 'PAGEREF' in (it.text or ''):
            seen_pageref = True
        fc = r.find('w:fldChar', ns)
        if fc is not None:
            kind = fc.get(q('fldCharType'))
            if kind == 'separate' and seen_pageref:
                in_result = True
            elif kind == 'end' and in_result:
                return
            continue
        wt = r.find('w:t', ns)
        if in_result and wt is not None:
            wt.text = str(page)


def write_package(src_zip, out_path, doc_tree):
    data = etree.tostring(doc_tree, xml_declaration=True, encoding='UTF-8', standalone=True)
    with zipfile.ZipFile(src_zip) as zin, \
            zipfile.ZipFile(out_path + '.tmp', 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            zout.writestr(item, data if item.filename == DOC else zin.read(item.filename))
    os.replace(out_path + '.tmp', out_path)


def main(src, out):
    tree = etree.fromstring(zipfile.ZipFile(src).read(DOC)).getroottree()
    body = tree.getroot().find('w:body', ns)
    n_lead, n_rows = patch_layout(body)
    print(f'{os.path.basename(src)}: {n_lead} lead-ins kept with next, '
          f'{n_rows} overview rows kept together')
    write_package(src, out, tree)
    entries = toc_entries(body)
    toc = [(title, bm) for _, bm, title in entries]
    work = tempfile.mkdtemp()
    previous = None
    for attempt in range(3):
        pages = harvest_pages(render_pdf(out, work), toc)
        if len(pages) != len(toc):
            sys.exit(f'only {len(pages)}/{len(toc)} headings found in the render')
        for p, bm, _ in entries:
            set_cached_page(p, pages[bm])
        write_package(src, out, tree)
        print(f'  pass {attempt + 1}: ToC pages {[pages[bm] for _, bm in toc]}')
        if pages == previous:
            break
        previous = pages


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
