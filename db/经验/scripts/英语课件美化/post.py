"""XML pass after the COM build: every explicit run font -> Comic Sans MS (latin),
East-Asian slot -> 微软雅黑 for Chinese text / Comic Sans MS for English-only text,
and any leftover size below 24pt (empty paragraph marks, list-style defaults) floored to 24pt.
usage: python post.py in.pptx out.pptx"""
import sys, re, zipfile
from lxml import etree

A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
q = lambda t: f'{{{A}}}{t}'
LAT, EA = "Comic Sans MS", "微软雅黑"

def has_cjk(t):
    return any(0x2E80 <= ord(c) <= 0x9FFF or 0xF900 <= ord(c) <= 0xFAFF or 0xFF00 <= ord(c) <= 0xFFEF
               or 0x3000 <= ord(c) <= 0x303F for c in t)

def text_of(el):
    return ''.join(el.itertext())

stats = {'latin': 0, 'ea': 0, 'sz': 0}

SMALL_OK = ("不时之需",)          # slide 21 gloss: deliberately below 24pt at the user's request

def fix_props(pr, cjk, floor=True):
    sz = pr.get('sz')
    if floor and sz is not None and int(sz) < 2400:
        pr.set('sz', '2400'); stats['sz'] += 1
    lat = pr.find(q('latin'))
    if lat is not None and not lat.get('typeface', '').startswith('+') and lat.get('typeface') != LAT:
        lat.set('typeface', LAT); stats['latin'] += 1
        for k in ('panose', 'pitchFamily', 'charset'):
            lat.attrib.pop(k, None)
    ea = pr.find(q('ea'))
    if ea is not None and not ea.get('typeface', '').startswith('+'):
        want = EA if cjk else LAT
        if ea.get('typeface') != want:
            ea.set('typeface', want); stats['ea'] += 1
            for k in ('panose', 'pitchFamily', 'charset'):
                ea.attrib.pop(k, None)

def process(xml):
    root = etree.fromstring(xml)
    for p in root.iter(q('p')):
        ptxt = text_of(p)
        pc = has_cjk(ptxt)
        fl = not any(k in ptxt for k in SMALL_OK)
        for child in p:
            if child.tag in (q('r'), q('fld')):
                rpr = child.find(q('rPr'))
                if rpr is not None:
                    fix_props(rpr, has_cjk(text_of(child)), fl)
            elif child.tag == q('endParaRPr'):
                fix_props(child, pc, fl)
            elif child.tag == q('pPr'):
                d = child.find(q('defRPr'))
                if d is not None:
                    fix_props(d, True)
    for lst in root.iter(q('lstStyle')):
        for d in lst.iter(q('defRPr')):
            fix_props(d, True)
    return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)

src, dst = sys.argv[1], sys.argv[2]
zin = zipfile.ZipFile(src)
zout = zipfile.ZipFile(dst, 'w')
pat = re.compile(r'ppt/(slides/slide\d+|slideLayouts/slideLayout\d+|slideMasters/slideMaster\d+)\.xml$')
for info in zin.infolist():
    data = zin.read(info.filename)
    if pat.match(info.filename):
        data = process(data)
    elif re.match(r'ppt/theme/theme\d+\.xml$', info.filename):
        t = data.decode('utf-8')
        t = re.sub(r'(<a:font script="(?:Hans|Hant)" typeface=")[^"]*(")', r'\g<1>微软雅黑\g<2>', t)
        data = t.encode('utf-8')
    zout.writestr(info, data, compress_type=info.compress_type)
zout.close()
print(stats)
