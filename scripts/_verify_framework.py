from lxml import etree
import re
W='http://schemas.openxmlformats.org/wordprocessingml/2006/main'; ns={'w':W}
def norm(s): return re.sub(r'\s+',' ',s).replace(' ',' ').strip()
def texts(path):
    root=etree.parse(path).getroot()
    out=[]
    for p in root.iter(f'{{{W}}}p'):
        t=norm(''.join(v.text or '' for v in p.iter(f'{{{W}}}t')))
        if t: out.append(t)
    return out
src=texts('instrument_unpacked/word/document.xml')
dst=texts('build7/word/document.xml')
joined=''.join(dst)
pos=0; missing=[]; moved=[]
for t in src:
    i=joined.find(t,pos)
    if i<0: (moved if joined.find(t)>=0 else missing).append(t)
    else: pos=i+len(t)
print(f'source paragraphs : {len(src)}')
print(f'output paragraphs : {len(dst)}')
print(f'missing           : {len(missing)}')
for m in missing[:12]: print('   MISSING', repr(m[:90]))
print(f'out of order      : {len(moved)}')
for m in moved[:12]: print('   MOVED  ', repr(m[:90]))
# heading level preservation
def heads(path):
    root=etree.parse(path).getroot(); out=[]
    for p in root.iter(f'{{{W}}}p'):
        ps=p.find('w:pPr/w:pStyle',ns)
        s=ps.get(f'{{{W}}}val') if ps is not None else ''
        if s.startswith('Heading') or s=='Style1':
            t=norm(''.join(v.text or '' for v in p.iter(f'{{{W}}}t')))
            out.append((s,t))
    return out
hs, hd = heads('instrument_unpacked/word/document.xml'), heads('build7/word/document.xml')
print(f'\nheadings source/output: {len(hs)}/{len(hd)}   levels identical: {hs==hd}')
