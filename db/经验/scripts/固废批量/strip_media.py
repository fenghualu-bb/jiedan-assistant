# Remove jumps/hyperlinks, video/audio, controls, action buttons; freeze GIFs. In-place on a .pptx.
import sys, re, io, zipfile, shutil
from lxml import etree
from PIL import Image
src = sys.argv[1]; tmp = src + ".tmp"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"; P = "http://schemas.openxmlformats.org/presentationml/2006/main"
LINKWORD = re.compile(r"视频|动画|录像|演示|flash|点击|播放", re.I)
stats = {"link": 0, "linktext": 0, "media": 0, "ctrl": 0, "gif": 0}
def fix_slide(xml):
    root = etree.fromstring(xml)
    for p in root.iter("{%s}p" % A):
        runs = p.findall("{%s}r" % A); removed = False
        for r in runs:
            rpr = r.find("{%s}rPr" % A)
            if rpr is None: continue
            hl = rpr.findall("{%s}hlinkClick" % A) + rpr.findall("{%s}hlinkMouseOver" % A)
            if not hl: continue
            t = r.find("{%s}t" % A); txt = t.text or ""
            for h in hl: rpr.remove(h)
            stats["link"] += 1
            if LINKWORD.search(txt):
                p.remove(r); removed = True; stats["linktext"] += 1
            else:   # keep word, normal underline off
                rpr.attrib.pop("u", None)
        if removed:   # tidy empty brackets left behind
            rs = [r for r in p.findall("{%s}r" % A)]
            full = "".join((r.find("{%s}t" % A).text or "") for r in rs)
            for a, b in (("（", "）"), ("(", ")")):
                if a + b in full.replace(" ", ""):
                    for i, r in enumerate(rs):
                        t = r.find("{%s}t" % A)
                        if t.text and t.text.rstrip().endswith(a):
                            # find next non-space run starting with b
                            for r2 in rs[i + 1:]:
                                t2 = r2.find("{%s}t" % A)
                                if (t2.text or "").strip() == "": t2.text = ""; continue
                                if t2.text.lstrip().startswith(b):
                                    t.text = t.text.rstrip()[:-1]; t2.text = t2.text.lstrip()[1:]
                                break
    # shape-level hyperlinks / actions
    for tag in ("hlinkClick", "hlinkMouseOver", "hlinkHover", "snd"):
        for h in list(root.iter("{%s}%s" % (A, tag))):
            h.getparent().remove(h); stats["link"] += 1
    # video/audio pics -> plain pictures
    for vf in list(root.iter("{%s}videoFile" % A)) + list(root.iter("{%s}audioFile" % A)) + list(root.iter("{%s}quickTimeFile" % A)):
        vf.getparent().remove(vf); stats["media"] += 1
    for ext in list(root.iter("{%s}ext" % P)):
        if ext.get("uri") == "{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}":   # p14:media
            ext.getparent().remove(ext)
    # ActiveX controls (flash) & action buttons
    for c in list(root.iter("{%s}controls" % P)):
        c.getparent().remove(c); stats["ctrl"] += 1
    for sp in list(root.iter("{%s}sp" % P)):
        g = sp.find(".//{%s}prstGeom" % A)
        if g is not None and g.get("prst", "").startswith("actionButton"):
            sp.getparent().remove(sp); stats["ctrl"] += 1
    for tag in ("timing", "transition"):
        for e in list(root.iter("{%s}%s" % (P, tag))): e.getparent().remove(e)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
MEDIA_REL = re.compile(r'<Relationship [^>]*Type="[^"]*/(video|audio|media|hyperlink|control)"[^>]*/>')
with zipfile.ZipFile(src) as zi, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
    for it in zi.infolist():
        d = zi.read(it.filename); n = it.filename
        if re.match(r"ppt/slides/slide\d+\.xml$", n): d = fix_slide(d)
        elif re.match(r"ppt/slides/_rels/slide\d+\.xml\.rels$", n):
            d = MEDIA_REL.sub("", d.decode("utf8")).encode("utf8")
        elif n.lower().endswith(".gif"):
            im = Image.open(io.BytesIO(d))
            if getattr(im, "n_frames", 1) > 1:
                b = io.BytesIO(); im.seek(0); im.convert("RGBA").save(b, "GIF"); d = b.getvalue(); stats["gif"] += 1
        elif re.search(r"\.(avi|mp4|wmv|swf|wav|mp3|mov)$", n, re.I):
            continue
        zo.writestr(it, d)
shutil.move(tmp, src); print(src, stats)
