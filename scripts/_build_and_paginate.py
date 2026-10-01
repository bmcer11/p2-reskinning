#!/usr/bin/env python3
"""Two-pass build: build, render, harvest Heading 1 page numbers into the ToC
cache, rebuild so the ToC's cached PAGEREF results match the pagination.
Usage: _build_and_paginate.py <script.py> <tag>   (tag 8 -> out8.docx, toc_pages8.json)"""
import json, os, subprocess, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _pdf_reskin_lib import render_pdf, harvest_pages

script, tag = sys.argv[1], sys.argv[2]
work = os.environ.get('RESKIN_WORK', os.getcwd())
cache = os.path.join(work, f'toc_pages{tag}.json')
for attempt in range(3):
    subprocess.run([sys.executable, script], check=True)
    docx = os.path.join(work, f'out{tag}.docx')
    pdf = render_pdf(docx, work)
    # recover the ToC (title, bookmark) list by re-running the block render
    import importlib.util
    spec = importlib.util.spec_from_file_location('m', script)
    m = importlib.util.module_from_spec(spec); spec.loader.exec_module(m)
    from _pdf_reskin_lib import Builder
    b = Builder(m.TPL, m.BUILD + '_probe', m.FOOTNOTES)
    b.render(m.BLOCKS)
    pages = harvest_pages(pdf, b.toc)
    old = json.load(open(cache)) if os.path.exists(cache) else {}
    json.dump(pages, open(cache, 'w'), indent=1)
    print('pass', attempt + 1, pages)
    if pages == old and len(pages) == len(b.toc):
        break
subprocess.run([sys.executable, script], check=True)
render_pdf(os.path.join(work, f'out{tag}.docx'), work)
print('done')
