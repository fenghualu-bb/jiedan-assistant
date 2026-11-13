# -*- coding: utf-8 -*-
"""Generic restyle engine built from chapter 1's standard.
usage: python engine.py src.pptx out.pptx assets_dir deck_config.py
"""
import sys, os, re, math
HERE = os.path.dirname(os.path.abspath(__file__))
SRC, OUT, ASSETS, CFG = sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4]
exec(open(os.path.join(HERE, "lib.py"), encoding="utf-8").read())
exec(open(os.path.join(HERE, "lib169.py"), encoding="utf-8").read())

# ---------------- deck config (defaults, overridden by CFG)
CHAPTER = ""            # e.g. "第三章  固体废物物理处理"
COVER = 1
TOC = {}                # n: (title, [items])
DIVIDERS = {}           # n: (label, title, carry_title)
TITLES = {}             # n: forced title text
NOTITLE_TOP = set()     # slides whose top text must not be taken as title
FIX = None              # per-deck fix file
exec(open(CFG, encoding="utf-8").read())
COURSE = "固体废物处置与资源化"
FOOTER = COURSE + "  ·  " + CHAPTER.replace("  ", " ")
OW, OH = 10.0, 7.5

# ---------------- theme fonts (all masters)
for m in prs.slide_masters:
    for rel in m.part.rels.values():
        if rel.reltype == RT.THEME:
            th = etree.fromstring(rel.target_part.blob)
            for tag in ("majorFont", "minorFont"):
                for f in th.iter(A(tag)):
                    for t in ("latin", "ea"):
                        e = f.find(A(t))
                        if e is not None:
                            e.set("typeface", FONT)
                            for k in ("panose", "pitchFamily", "charset"): e.attrib.pop(k, None)
            rel.target_part._blob = etree.tostring(th, xml_declaration=True, encoding="UTF-8", standalone=True)

def set_bg(el):
    cSld = el.find(P("cSld")); bg = cSld.find(P("bg"))
    if bg is not None: cSld.remove(bg)
    bg = oxml_parser.makeelement(P("bg")); bgPr = sub(bg, P("bgPr")); bgPr.append(solid(BG)); sub(bgPr, A("effectLst"))
    cSld.insert(0, bg)

for m in prs.slide_masters:
    set_bg(m._element)
    for l in m.slide_layouts:
        cSld = l._element.find(P("cSld")); bg = cSld.find(P("bg"))
        if bg is not None: cSld.remove(bg)
    # strip decorative pictures/shapes from masters (logo, bands)
    for sp in list(m.shapes):
        if not sp.is_placeholder: sp._element.getparent().remove(sp._element)

master = prs.slide_masters[0]
def has_title_ph(l):
    return any(s.is_placeholder and s.placeholder_format.type in (1, 3) for s in l.placeholders)
cands = [l for l in master.slide_layouts if has_title_ph(l)]
layout = ([l for l in cands if "标题和内容" in l.name] or cands)[0]
blank = ([l for l in master.slide_layouts if l.name.endswith("空白")] or
         [l for l in master.slide_layouts if l is not layout])[0]
for sp in list(blank.shapes): sp._element.getparent().remove(sp._element)

TB = dict(x=0.6, y=0.42, w=6.3, h=0.74)
def style_title_ph(sp):
    spPr = sp.find(P("spPr"))
    for ch in list(spPr): spPr.remove(ch)
    xfrm = sub(spPr, A("xfrm")); sub(xfrm, A("off"), x=I(TB["x"]), y=I(TB["y"]))
    sub(xfrm, A("ext"), cx=I(TB["w"]), cy=I(TB["h"]))
    g = sub(spPr, A("prstGeom"), prst="rect"); sub(g, A("avLst"))
    spPr.append(solid(BLUE)); ln = sub(spPr, A("ln")); sub(ln, A("noFill"))
    tx = sp.find(P("txBody")); bp = tx.find(A("bodyPr"))
    for k in list(bp.attrib): del bp.attrib[k]
    for ch in list(bp): bp.remove(ch)
    bp.set("lIns", str(I(0.32))); bp.set("rIns", str(I(0.2))); bp.set("tIns", "0"); bp.set("bIns", "0")
    bp.set("anchor", "ctr"); bp.set("wrap", "none"); sub(bp, A("noAutofit"))

for sp in master.shapes:
    if sp.is_placeholder and sp.placeholder_format.type in (1, 3): style_title_ph(sp._element)
    elif sp.is_placeholder: place(sp, 0.9, 1.5, 11.5, 5.4)
txs = master._element.find(P("txStyles"))
for lvl in txs.find(P("titleStyle")):
    if etree.QName(lvl).localname.startswith("lvl"):
        lvl.set("algn", "l"); d = lvl.find(A("defRPr"))
        if d is not None:
            d.set("sz", "2400"); d.set("b", "1")
            for e in list(d):
                if etree.QName(e).localname in ("solidFill", "effectLst", "gradFill"): d.remove(e)
            fix_rpr(d, force=WHITE)
for lvl in txs.find(P("bodyStyle")):
    d = lvl.find(A("defRPr"))
    if d is not None:
        for e in list(d):
            if etree.QName(e).localname == "solidFill": d.remove(e)
        fix_rpr(d, force=TEXT)
        if d.get("sz") and int(d.get("sz")) > 2400: d.set("sz", "2400")
    bc = lvl.find(A("buClr"))
    if bc is not None:
        for ch in list(bc): bc.remove(ch)
        sub(bc, A("srgbClr"), val=BLUE)
for sp in list(layout.shapes):
    if sp.is_placeholder and sp.placeholder_format.type in (1, 3): style_title_ph(sp._element)
    elif sp.is_placeholder: place(sp, 0.9, 1.5, 11.5, 5.4)
    else: sp._element.getparent().remove(sp._element)

def deco(target):
    add_box(target, TB["x"] + 0.14, TB["y"] + TB["h"], 0.055, 0.22, fill=ORANGE, name="橙色标记")
    add_text(target, 0.6, 7.02, 7, 0.3, FOOTER, size=10, color=MUTED, name="页脚")
    nb = add_text(target, 11.73, 7.02, 1.0, 0.3, "", size=10, color=MUTED, align=PP_ALIGN.RIGHT, name="页码")
    p = nb.text_frame.paragraphs[0]
    for r in list(p._p.findall(A("r"))): p._p.remove(r)
    fld = etree.SubElement(p._p, A("fld")); fld.set("id", "{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}"); fld.set("type", "slidenum")
    rpr = sub(fld, A("rPr"), lang="en-US", sz="1000"); rpr.append(solid(MUTED))
    sub(rpr, A("latin"), typeface=FONT); sub(rpr, A("ea"), typeface=FONT)
    t = sub(fld, A("t")); t.text = "‹#›"

tmp = prs.slides.add_slide(layout)
for ph in list(tmp.placeholders): ph._element.getparent().remove(ph._element)
deco(tmp)
for el in list(tmp.shapes._spTree)[2:]: layout.shapes._spTree.append(el)
sl = prs.slides._sldIdLst; last = sl[-1]; prs.part.drop_rel(last.rId); sl.remove(last)

# ---------------- helpers
slides = list(prs.slides)
def S(n): return slides[n - 1]
def shp(n, name, k=0): return [s for s in S(n).shapes if s.name == name][k]
def by_text(n, start):
    for s in S(n).shapes:
        if s.has_text_frame and s.text_frame.text.strip().startswith(start): return s
    raise KeyError((n, start))

def relink(slide, lay):
    for rel in slide.part.rels.values():
        if rel.reltype == RT.SLIDE_LAYOUT: rel._target = lay.part

def clear(slide):
    for sp in list(slide.shapes): sp._element.getparent().remove(sp._element)

def drop_slide_bg(slide):
    cSld = slide._element.find(P("cSld")); bg = cSld.find(P("bg"))
    if bg is not None: cSld.remove(bg)

def build_cover(slide):
    clear(slide)
    add_box(slide, 0, 2.45, 13.333, 2.55, fill=BLUE, name="主标题底")
    add_text(slide, 0.95, 2.75, 11.5, 1.1, COURSE, size=46, color=WHITE, bold=True, anchor=MSO_ANCHOR.MIDDLE, name="主标题")
    add_text(slide, 0.95, 3.95, 11.5, 0.7, CHAPTER, size=24, color="D6E6F5", anchor=MSO_ANCHOR.MIDDLE, name="副标题")
    add_box(slide, 0.95, 5.0, 0.07, 0.36, fill=ORANGE, name="橙色标记")

def build_divider(slide, label, title):
    clear(slide)
    add_box(slide, 0, 2.45, 13.333, 2.55, fill=BLUE, name="节标题底")
    add_text(slide, 0.95, 2.8, 11.5, 0.6, label, size=24, color="D6E6F5", anchor=MSO_ANCHOR.MIDDLE, name="节序")
    add_text(slide, 0.95, 3.45, 11.5, 1.1, title, size=40, color=WHITE, bold=True, anchor=MSO_ANCHOR.MIDDLE, name="节标题")
    add_box(slide, 0.95, 5.0, 0.07, 0.36, fill=ORANGE, name="橙色标记")
    add_text(slide, 0.6, 7.02, 7, 0.3, FOOTER, size=10, color=MUTED, name="页脚")

def build_toc(slide, title, items):
    clear(slide); set_title(slide, title)
    n = len(items)
    if n <= 5:
        gap = 0.22; h = min(1.1, (5.2 - (n - 1) * gap) / n); y = TOP + 0.1
        for i, it in enumerate(items):
            add_box(slide, L, y, W, h, fill=WHITE, name="目录行")
            add_text(slide, L, y, h * 1.1, h, "%02d" % (i + 1), size=26, color=WHITE, bold=True, align=PP_ALIGN.CENTER,
                     anchor=MSO_ANCHOR.MIDDLE, shape=MSO_SHAPE.RECTANGLE, fill=BLUE, name="序号")
            add_text(slide, L + h * 1.1 + 0.4, y, W - h * 1.1 - 0.6, h, it, size=24, color=TEXT, bold=True, anchor=MSO_ANCHOR.MIDDLE)
            y += h + gap
    else:
        rows = math.ceil(n / 2); gap = 0.25; h = min(1.0, (5.2 - (rows - 1) * gap) / rows)
        for i, it in enumerate(items):
            c, r = i // rows, i % rows
            x = L + c * (C2 + 0.3); y = TOP + 0.1 + r * (h + gap)
            add_box(slide, x, y, C2, h, fill=WHITE, name="目录行")
            add_text(slide, x, y, h * 1.1, h, "%02d" % (i + 1), size=24, color=WHITE, bold=True, align=PP_ALIGN.CENTER,
                     anchor=MSO_ANCHOR.MIDDLE, shape=MSO_SHAPE.RECTANGLE, fill=BLUE, name="序号")
            add_text(slide, x + h * 1.1 + 0.3, y, C2 - h * 1.1 - 0.4, h, it, size=22, color=TEXT, bold=True, anchor=MSO_ANCHOR.MIDDLE)

C2 = (W - 0.3) / 2; C3 = (W - 0.6) / 3

def text_w_in(t, pt):
    w = 0.0
    for ch in t:
        w += (pt / 72.0) * (1.0 if ord(ch) > 0x2E80 else 0.55)
    return w

def para_sz(p, default):
    s = [int(r.get("sz")) / 100 for r in p.iter(A("rPr")) if r.get("sz")]
    return max(s) if s else default

def est_height(el, width_in, default=18):
    isph = el.find('.//' + P('ph')) is not None
    if isph: default = 22
    tx = el.find(".//" + P("txBody")) if el.tag != P("txBody") else el
    if tx is None: return 0.3
    bp = tx.find(A("bodyPr"))
    li = int(bp.get("lIns", 91440)) / IN + int(bp.get("rIns", 91440)) / IN if bp is not None else 0.2
    ti = int(bp.get("tIns", 45720)) / IN + int(bp.get("bIns", 45720)) / IN if bp is not None else 0.1
    avail = max(0.8, width_in - li)
    h = ti
    for p in tx.findall(A("p")):
        pt = para_sz(p, default)
        ppr = p.find(A("pPr")); ind = 0; f = 1.2
        if ppr is not None:
            ind = int(ppr.get("marL", 0)) / IN
            ls = ppr.find(A("lnSpc") + "/" + A("spcPct"))
            if ls is not None: f = 1.2 * int(ls.get("val")) / 100000
            for tag in ("spcBef", "spcAft"):
                e = ppr.find(A(tag))
                if e is not None:
                    pp = e.find(A("spcPts")); pc = e.find(A("spcPct"))
                    if pp is not None: h += int(pp.get("val")) / 100 / 72
                    if pc is not None: h += int(pc.get("val")) / 100000 * pt / 72
        t = text_of(p)
        lines = max(1, math.ceil(text_w_in(t, pt) / max(0.5, avail - ind))) if t else 1
        h += lines * pt / 72 * f * (1.3 if isph else 1.08)
    return h

def scale_fonts(el, f):
    for tag in ("rPr", "endParaRPr", "defRPr"):
        for r in el.iter(A(tag)):
            if r.get("sz"): r.set("sz", str(max(1000, int(int(r.get("sz")) * f))))

TEXT_GEOMS = {None, "rect", "roundRect", "wedgeRectCallout", "wedgeRoundRectCallout", "flowChartProcess",
              "snip1Rect", "round2SameRect", "plaque", "flowChartAlternateProcess"}

def geom(sp):
    return [(sp.left or 0) / IN, (sp.top or 0) / IN, (sp.width or 0) / IN, (sp.height or 0) / IN]

def is_text_item(sp):
    if sp.shape_type in (13, 6, 3, 7, 10, 19) or sp.has_chart: return False
    if getattr(sp, "has_table", False) and sp.has_table: return False
    if sp._element.tag != P("sp"): return False
    if not (sp.has_text_frame and sp.text_frame.text.strip()): return False
    return prst(sp._element) in TEXT_GEOMS

def union(bs):
    x0 = min(b[0] for b in bs); y0 = min(b[1] for b in bs)
    x1 = max(b[0] + b[2] for b in bs); y1 = max(b[1] + b[3] for b in bs)
    return [x0, y0, x1 - x0, y1 - y0]

def inter_area(a, b):
    w = min(a[0] + a[2], b[0] + b[2]) - max(a[0], b[0]); h = min(a[1] + a[3], b[1] + b[3]) - max(a[1], b[1])
    return max(0, w) * max(0, h)

def setg(sp, x, y, w, h):
    if sp._element.tag == P("grpSp"):
        xf = sp._element.find(P("grpSpPr") + "/" + A("xfrm"))
        xf.find(A("off")).set("x", str(I(x))); xf.find(A("off")).set("y", str(I(y)))
        xf.find(A("ext")).set("cx", str(I(max(0.05, w)))); xf.find(A("ext")).set("cy", str(I(max(0.05, h))))
        return
    sp.left = I(x); sp.top = I(y); sp.width = I(max(0.05, w)); sp.height = I(max(0.05, h))

def layout_text(items, x0, x1, top, maxh, allow_grow=True):
    """reflow text items (with orig geometry) into column [x0,x1] starting at top. returns used height."""
    if not items: return 0
    tb = union([g for _, g in items]); fx = (x1 - x0) / max(0.5, tb[2])
    new = []
    for sp, g in items:
        nx = x0 + (g[0] - tb[0]) * fx; nw = g[2] * fx
        new.append([sp, g, nx, nw, est_height(sp._element, nw)])
    def flow():
        rows = []
        for it in sorted(new, key=lambda r: r[1][1]):
            g = it[1]
            def xover(r_):
                return any(min(g[0] + g[2], o[1][0] + o[1][2]) - max(g[0], o[1][0]) > 0.15 * min(g[2], o[1][2]) for o in r_["items"])
            if rows and g[1] < rows[-1]["y1"] - 0.15 and g[1] + 0.05 >= rows[-1]["y0"] and not xover(rows[-1]):
                rows[-1]["items"].append(it); rows[-1]["y1"] = max(rows[-1]["y1"], g[1] + g[3])
            else:
                rows.append({"y0": g[1], "y1": g[1] + g[3], "items": [it]})
        y = top; prev = None
        for r in rows:
            if prev is not None:
                y += min(0.5, max(0.3, r["y0"] - prev)) + (0.3 if any(it[0].is_placeholder for it in rows[rows.index(r)-1]["items"]) else 0)
            hh = max(it[4] for it in r["items"])
            for it in r["items"]: it.append(y)
            y += hh; prev = r["y1"]
        return y - top
    used = flow()
    if allow_grow and used < maxh * 0.7:
        f = min(1.6, (maxh * 0.7 / max(0.3, used)) ** 0.5)
        for it in new:
            scale_fonts(it[0]._element, f); it[4] = est_height(it[0]._element, it[3]); del it[5:]
        used = flow()
    elif used > maxh:
        f = max(0.82, maxh / used)
        for it in new:
            scale_fonts(it[0]._element, f); it[4] = est_height(it[0]._element, it[3]); del it[5:]
        used = flow()
    for sp, g, nx, nw, eh, ny in new:
        setg(sp, nx, ny, nw, max(eh, 0.3))
    return used

def layout_vis(items, ox, oy, s, vb):
    for sp, g in items:
        setg(sp, ox + (g[0] - vb[0]) * s, oy + (g[1] - vb[1]) * s, g[2] * s, g[3] * s)
        if abs(s - 1) > 0.05 and sp._element.find(".//" + A("t")) is not None:
            scale_fonts(sp._element, min(s, 1.35))

def auto_layout(slide, items):
    """items: [(shape, orig_geom)] for one slide (title excluded). Arrange into 16:9 content area."""
    H = BOT - TOP
    texts = [(s, g) for s, g in items if is_text_item(s)]
    for s_, g_ in texts:           # floor body size, then widen narrow labels so they never wrap per-character
        for r_ in s_._element.iter(A("rPr")):
            if r_.get("sz") and int(r_.get("sz")) < 1400: r_.set("sz", "1400")
    for s_, g_ in sorted(texts, key=lambda t: t[1][0]):
        tt = text_of(s_._element).replace(chr(10), "")
        if g_[2] < 1.7 and len(tt) > 3:
            need = min(2.6, max(g_[2], text_w_in(tt, max_sz(s_)) + 0.35)) if len(tt) <= 14 else 2.6
            d = need - g_[2]
            if d > 0:
                for s2, g2 in texts:
                    if s2 is not s_ and g2[0] >= g_[0] + g_[2] - 0.05 and g2[1] < g_[1] + g_[3] and g2[1] + g2[3] > g_[1]:
                        g2[0] += d
                g_[2] = need
    vis = [(s, g) for s, g in items if not is_text_item(s)]
    if vis:
        for _ in range(2):  # labels overlapping visuals belong to the visual
            vb = union([g for _, g in vis])
            moved = [(s, g) for s, g in texts if inter_area(g, vb) > 0.3 * max(0.01, g[2] * g[3])]
            for t in moved: texts.remove(t); vis.append(t)
    if not vis:
        used = layout_text(texts, L, R, TOP + 0.05, H); return "text"
    vis = [(s_, g) for s_, g in vis if g[2] > 0.01 and g[3] > 0.01] or vis
    if all(g[2] <= 0.01 or g[3] <= 0.01 for _, g in vis):
        return layout_text(texts, L, R, TOP + 0.05, H) and 'text'
    vb = union([g for _, g in vis]); has_pic = any(s.shape_type == 13 for s, _ in vis)
    smax = 1.7 if has_pic else 1.45
    if not texts:
        s = min(W / vb[2], H / vb[3], smax)
        layout_vis(vis, L + (W - vb[2] * s) / 2, TOP + (H - vb[3] * s) / 2 * 0.6, s, vb); return "vis"
    tb = union([g for _, g in texts])
    ovx = max(0, min(tb[0] + tb[2], vb[0] + vb[2]) - max(tb[0], vb[0])) / max(0.1, min(tb[2], vb[2]))
    vis_left = vb[0] + vb[2] / 2 < tb[0] + tb[2] / 2
    aspect = vb[2] / vb[3]
    side = ovx < 0.35
    if not side and aspect <= 1.9:
        # try converting stacked -> side by side
        colv = min(W * 0.52, H * aspect * 0.95)
        colt = W - colv - 0.4
        if colt >= 4.6:
            tot = sum(est_height(s._element, g[2] * colt / max(0.5, tb[2])) for s, g in texts)
            side = tot <= H * 0.95
            vis_left = False
    if side:
        share = vb[2] / (vb[2] + tb[2])
        colv = max(3.8, min(W * 0.58, W * share * 1.1))
        s = min(colv / vb[2], H / vb[3], smax); vw, vh = vb[2] * s, vb[3] * s
        colt = W - vw - 0.4
        if vis_left:
            vx, tx0 = L, L + vw + 0.4
        else:
            tx0, vx = L, L + colt + 0.4
        used = layout_text(texts, tx0, tx0 + colt, TOP + 0.05, H)
        block = max(used, vh)
        ty = TOP + 0.05 + max(0, (block - used) / 2)
        if ty > TOP + 0.06:
            for sp, _ in texts: sp.top = sp.top + I(ty - TOP - 0.05)
        layout_vis(vis, vx, TOP + 0.05 + max(0, (block - vh) / 2), s, vb)
        return "side"
    # stacked: keep original order
    text_first = tb[1] + tb[3] / 2 < vb[1] + vb[3] / 2
    th = sum(est_height(s._element, g[2] * W / max(0.5, tb[2])) for s, g in texts) + 0.15
    s = min(W / vb[2], max(1.0, H - th - 0.3) / vb[3], smax); vw, vh = vb[2] * s, vb[3] * s
    if text_first:
        used = layout_text(texts, L, R, TOP + 0.05, H - vh - 0.3, allow_grow=False)
        layout_vis(vis, L + (W - vw) / 2, TOP + 0.05 + used + 0.25, s, vb)
    else:
        layout_vis(vis, L + (W - vw) / 2, TOP + 0.05, s, vb)
        layout_text(texts, L, R, TOP + 0.05 + vh + 0.25, H - vh - 0.3, allow_grow=False)
    return "stack"

def set_title(slide, text):
    text = re.sub(r"\s+", " ", text).strip()
    el = add_title_ph(slide, text)
    wneed = text_w_in(text, 24) + 0.75
    if wneed > TB["w"]:
        spPr = el.find(P("spPr"))
        xfrm = sub(spPr, A("xfrm")); sub(xfrm, A("off"), x=I(TB["x"]), y=I(TB["y"]))
        if wneed > 11.4:
            for r in el.iter(A("rPr")): r.set("sz", "2000")
            wneed = text_w_in(text, 20) + 0.75
        sub(xfrm, A("ext"), cx=I(min(11.6, wneed)), cy=I(TB["h"]))
    return el

def max_sz(sp):
    s = [int(r.get("sz")) / 100 for r in sp._element.iter(A("rPr")) if r.get("sz")]
    return max(s) if s else 20

NUM_RE = re.compile(r"^\s*\d+\.\d+")
DIV_RE = re.compile(r"^第[一二三四五六七八九十]+节\s*")

# ---------------- slides
carry = ""
MODES = {}
for idx, slide in enumerate(slides, 1):
    remove_anim(slide); drop_slide_bg(slide)
    if idx == COVER:
        relink(slide, blank); build_cover(slide); continue
    if idx in DIVIDERS:
        lab, tit, carry = DIVIDERS[idx]; relink(slide, blank); build_divider(slide, lab, tit); continue
    if idx in TOC:
        relink(slide, layout); build_toc(slide, *TOC[idx]); continue
    items = []; title = None; title_sp = []
    shapes = list(slide.shapes)
    cands = []
    for sp in shapes:
        g = geom(sp); txt = sp.text_frame.text.strip() if sp.has_text_frame else ""
        # logos / home buttons / header bands
        if sp.shape_type == 13 and g[1] < 1.4 and g[0] < 1.7 and g[2] < 2.2 and g[3] < 2.0:
            sp._element.getparent().remove(sp._element); continue
        if prst(sp._element) and prst(sp._element).startswith("actionButton"):
            sp._element.getparent().remove(sp._element); continue
        if not txt and sp.shape_type not in (13, 6, 3, 7, 10, 19) and g[1] < 0.35 and g[3] < 2.0 and g[2] > 4:
            sp._element.getparent().remove(sp._element); continue
        if not txt and sp._element.tag in (P("cxnSp"), P("sp")) and g[1] < 2.0 and g[3] < 0.05 and g[0] < 0.2 and g[2] < 2.6:
            sp._element.getparent().remove(sp._element); continue
        if not txt and sp._element.tag == P("sp") and g[2] >= 9.0 and g[3] < 0.06 and sp.shape_type not in (13, 6):
            sp._element.getparent().remove(sp._element); continue
        if sp.is_placeholder and sp.placeholder_format.type in (1, 3):
            if txt: cands.append((0, sp, txt))
            else: sp._element.getparent().remove(sp._element); continue
        elif txt and idx not in NOTITLE_TOP and sp._element.tag == P("sp") and g[1] < 1.25 and len(txt) <= 36 \
                and txt.count("\n") <= 1 and max_sz(sp) >= 18:
            cands.append((1, sp, txt))
        items.append((sp, g))
    if idx in TITLES:
        title = TITLES[idx]
        for _, sp, t in cands:
            if t.replace(" ", "") in title.replace(" ", "") or title.replace(" ", "") in t.replace(" ", ""):
                title_sp.append(sp)
    elif cands:
        cands.sort(key=lambda c: (c[0], -max_sz(c[1]), c[1].top))
        chosen = cands[0]
        title = chosen[2]; title_sp.append(chosen[1])
        if DIV_RE.match(title) and len(cands) > 1:   # prefer the specific heading over "第X节 …"
            title = cands[1][2]; title_sp.append(cands[1][1])
        for _, sp, t in cands[1:]:
            if DIV_RE.match(t) or t.replace(" ", "") in title.replace(" ", ""):
                title_sp.append(sp)
    if title:
        title = DIV_RE.sub("", title.replace("\n", " ")) or title
        if NUM_RE.match(title): carry = title
    else:
        title = carry or CHAPTER
    items = [(s, g) for s, g in items if s not in title_sp]
    for sp in title_sp:
        if sp._element.getparent() is not None: sp._element.getparent().remove(sp._element)
    relink(slide, layout)
    for sp, g in items:   # pin geometry before anything inherits from the new layout
        setg(sp, *g)
    for sp, g in items:
        if is_text_item(sp):
            bp = sp._element.find(".//" + A("bodyPr"))
            if bp is not None:
                bp.set("wrap", "square")
                bp.attrib.pop("vert", None); bp.attrib.pop("rot", None)
                for e in list(bp):
                    if etree.QName(e).localname in ("spAutoFit", "normAutofit", "noAutofit"): bp.remove(e)
    set_title(slide, title)
    items = [(sp, g) for sp, g in items if sp._element.getparent() is not None]
    def _is_ole(sp): return sp._element.tag == P("graphicFrame") or "oleObj" in sp._element.xml[:3000]
    def _is_line(sp): return sp._element.tag == P("cxnSp") or prst(sp._element) in ("line", "straightConnector1") or (not sp.has_text_frame or not sp.text_frame.text.strip()) and sp._element.tag == P("sp") and (sp.height or 0) < 40000
    ntext = sum(1 for sp, g in items if is_text_item(sp)); nline = sum(1 for sp, g in items if _is_line(sp))
    complex_ = any(sp.shape_type == 6 for sp, g in items) or len(items) >= 7 or (ntext >= 1 and any(_is_ole(sp) for sp, g in items)) or (nline >= 1 and ntext >= 2)
    if items and complex_ and not (len(items) == ntext):
        gs = [g for sp, g in items if g[2] > 0.02 or g[3] > 0.02]
        vb = union(gs); sc = min(W / max(0.5, vb[2]), (BOT - TOP - 0.3) / max(0.5, vb[3]), 1.35)
        ox = L + (W - vb[2] * sc) / 2
        for sp, g in items:
            setg(sp, ox + (g[0] - vb[0]) * sc, TOP + (g[1] - vb[1]) * sc, g[2] * sc, g[3] * sc)
            if abs(sc - 1) > 0.05 and sp._element.find(".//" + A("t")) is not None: scale_fonts(sp._element, min(sc, 1.3))
        MODES[idx] = "uniform"
    else:
        MODES[idx] = auto_layout(slide, items) if items else "empty"
    # white card behind the text block (chapter-1 look)
    tshapes = [sp for sp, g in items if is_text_item(sp) and sp._element.getparent() is not None]
    if tshapes and MODES[idx] in ("text", "side", "stack"):
        gs = [geom(sp) for sp in tshapes]
        if sum(len(text_of(sp._element)) for sp in tshapes) > 25:
            u = union(gs); pad = 0.18
            card = add_box(slide, max(0.3, u[0] - pad), max(TOP - 0.05, u[1] - pad), min(12.73, u[0] + u[2] + pad) - max(0.3, u[0] - pad),
                           min(u[3] + 2 * pad, 6.95 - max(TOP - 0.05, u[1] - pad)), fill=WHITE, name="文字卡片")
            tree = slide.shapes._spTree; tree.remove(card._element)
            tree.insert(3, card._element)
            for sp in tshapes:           # text on card: no own fill
                spPr = sp._element.find(P("spPr"))
                if eff_fill(sp._element) not in (None,) : pass
    for sp, g in items:
        restyle_shape(sp._element)

def est_box(sp):
    g = geom(sp); return g[0], g[1], g[2], max(g[3], est_height(sp._element, g[2]))
def resolve_overlaps(slide):
    tx = [sp for sp in slide.shapes if is_text_item(sp) and sp.name not in ("页脚", "页码", "文字卡片") and not (sp.is_placeholder and sp.placeholder_format.type in (1, 3))]
    tx.sort(key=lambda sp: (sp.top, sp.left))
    for i, a in enumerate(tx):
        for b in tx[:i]:
            ax, ay, aw, ah = est_box(a); bx, by, bw, bh = est_box(b)
            ox = min(ax + aw, bx + bw) - max(ax, bx); oy = min(ay + ah, by + bh) - max(ay, by)
            if ox > 0.2 and oy > 0.08 and ay >= by - 0.05:
                a.top = I(by + bh + 0.1)
    low = max([(sp.top / IN + est_box(sp)[3]) for sp in tx] or [0])
    if low > 6.95:
        f = max(0.75, (6.95 - TOP) / (low - TOP))
        for sp in tx:
            scale_fonts(sp._element, f); sp.top = I(TOP + (sp.top / IN - TOP) * f)
for idx_, sl_ in enumerate(slides, 1):
    if idx_ != COVER and idx_ not in DIVIDERS and idx_ not in TOC and MODES.get(idx_) != "uniform": resolve_overlaps(sl_)

# ---------------- charts: palette
pal = ["1A5F9A", "3B7DBA", "5E9BD2", "8BB8E3", "B5D2EE", "EE8742"]
for slide in slides:
    for sp in slide.shapes:
        if getattr(sp, "has_chart", False) and sp.has_chart:
            try:
                for plot in sp.chart.plots:
                    for i, ser in enumerate(plot.series):
                        if plot.__class__.__name__.startswith(("Pie", "Doughnut")):
                            for j in range(len(list(plot.categories))):
                                pt = ser.points[j]; pt.format.fill.solid()
                                pt.format.fill.fore_color.rgb = RGBColor.from_string(pal[j % len(pal)])
                        else:
                            ser.format.fill.solid(); ser.format.fill.fore_color.rgb = RGBColor.from_string(pal[i % len(pal)])
            except Exception as e:
                print("chart skip", e)

MC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
def fix_ole(slide):
    """Use the OOXML fallback (oleObj + pic) instead of the VML choice, and sync the preview pic to the frame."""
    for gf in slide.shapes._spTree.iter(P("graphicFrame")):
        xf = gf.find(P("xfrm"))
        for ac in list(gf.iter(MC + "AlternateContent")):
            fb = ac.find(MC + "Fallback")
            if fb is None: continue
            parent = ac.getparent(); i = list(parent).index(ac)
            for ch in list(fb): parent.insert(i, ch); i += 1
            parent.remove(ac)
        pics = list(gf.iter(P("pic")))
        if pics and xf is not None:
            pic = copy.deepcopy(pics[0])
            pxf = pic.find(P("spPr") + "/" + A("xfrm"))
            pxf.find(A("off")).attrib.update(xf.find(A("off")).attrib)
            pxf.find(A("ext")).attrib.update(xf.find(A("ext")).attrib)
            nm = gf.find(P("nvGraphicFramePr") + "/" + P("cNvPr"))
            c = pic.find(P("nvPicPr") + "/" + P("cNvPr")); c.set("id", nm.get("id")); c.set("name", "公式" if "Equation" in (gf.find(".//" + P("oleObj")).get("progId", "") if gf.find(".//" + P("oleObj")) is not None else "") else "嵌入图")
            gf.addprevious(pic); gf.getparent().remove(gf)
for slide in slides:
    fix_ole(slide)
    for rid, rel in list(slide.part.rels.items()):
        if rel.reltype.endswith("/vmlDrawing"):
            slide.part.rels.pop(rid)
    for ld in slide._element.findall(P("legacyDrawing")):
        rid = ld.get("{%s}id" % NS["r"]); slide._element.remove(ld)
        try: slide.part.drop_rel(rid)
        except Exception: pass

if FIX: exec(open(FIX, encoding="utf-8").read())

# ---------------- type scale (generic part of chapter 1 typo.py)
SIZE_MAP = {11: 12, 12: 12, 13: 14, 14: 14, 15: 16, 16: 16, 17: 16, 18: 18, 19: 18, 20: 20, 21: 20, 22: 20,
            23: 24, 24: 24, 25: 24, 26: 24, 27: 28, 28: 28, 29: 28, 30: 28, 31: 28, 32: 28}
def run_color(rpr):
    f = rpr.find(A("solidFill"))
    if f is None: return None
    c = f.find(A("srgbClr")); return c.get("val") if c is not None else None
def recolor_run(rpr, c):
    for e in rpr.findall(A("solidFill")): rpr.remove(e)
    _insert_rpr_fill(rpr, solid(c))
for idx, slide in enumerate(slides, 1):
    if idx == COVER or idx in DIVIDERS: continue
    for sp_el in slide.shapes._spTree.iter(P("sp")):
        ph = sp_el.find(".//" + P("ph"))
        if ph is not None and ph.get("type") in ("title", "ctrTitle"): continue
        tx = sp_el.find(P("txBody"))
        if tx is None: continue
        for p in tx.findall(A("p")):
            for tag in ("rPr", "endParaRPr"):
                for r in p.iter(A(tag)):
                    s = r.get("sz")
                    if s:
                        pt = round(int(s) / 100)
                        if pt in SIZE_MAP: r.set("sz", str(SIZE_MAP[pt] * 100))
                        elif pt > 32: r.set("sz", "2800")
            runs = p.findall(A("r")); plen = len(text_of(p))
            szs = [int(r.find(A("rPr")).get("sz", "0")) / 100 for r in runs if r.find(A("rPr")) is not None]
            if plen > 30 and (max(szs) if szs else 0) <= 24:
                for r in runs:
                    rpr = r.find(A("rPr"))
                    if rpr is None: continue
                    c = run_color(rpr); t = (r.find(A("t")).text or "")
                    if c == BLUE and len(t.strip()) > 20: recolor_run(rpr, TEXT); c = TEXT
                    if c != EMPH: rpr.set("b", "0")

prs.save(OUT)
import subprocess; subprocess.run([sys.executable, os.path.join(HERE, "fontfix.py"), OUT], check=True)
import collections; print("modes", collections.Counter(MODES.values()))
print("saved", OUT)
