# -*- coding: utf-8 -*-
"""Restyle 第1章.pptx: 16:9, blue title boxes (ref image 3), flat minimal content."""
import sys, re, copy, colorsys, os
from lxml import etree
from pptx import Presentation
from pptx.util import Emu, Pt, Inches
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn
from pptx.oxml import parse_xml, oxml_parser

IN = 914400
def I(v): return int(round(v * IN))

BLUE = "1A5F9A"; ORANGE = "EE8742"; TEXT = "333333"; MUTED = "8A96A3"
BG = "E8F0F9"; LIGHT = "FFFFFF"; LINE = "9DBFE0"; EMPH = "C0392B"; WHITE = "FFFFFF"
FONT = "微软雅黑"
NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
A = lambda t: "{%s}%s" % (NS["a"], t)
P = lambda t: "{%s}%s" % (NS["p"], t)

prs = Presentation(SRC)
OLD_W = prs.slide_width
prs.slide_width = 12192000; prs.slide_height = I(7.5)
sz = prs.part._element.find(P("sldSz")); sz.attrib.pop("type", None)
DX = I(1.6665); DY = I(-0.5)

SECTION_TITLES = {"1.1": "1.1  固体废物定义与特性", "1.2": "1.2  固体废物分类",
                  "1.3": "1.3  固体废物污染与控制", "1.4": "1.4  固体废物管理原则"}

# ---------------------------------------------------------------- helpers
def sub(parent, tag, **attrs):
    el = etree.SubElement(parent, tag)
    for k, v in attrs.items(): el.set(k, str(v))
    return el

def solid(color):
    f = oxml_parser.makeelement(A("solidFill")); sub(f, A("srgbClr"), val=color); return f

SPPR_ORDER = ["xfrm", "custGeom", "prstGeom", "noFill", "solidFill", "gradFill", "blipFill",
              "pattFill", "grpFill", "ln", "effectLst", "effectDag", "scene3d", "sp3d", "extLst"]
FILL_TAGS = [A(t) for t in ("noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill")]

def sppr_insert(spPr, el):
    name = etree.QName(el).localname
    idx = SPPR_ORDER.index(name)
    for i, ch in enumerate(spPr):
        cn = etree.QName(ch).localname
        if cn in SPPR_ORDER and SPPR_ORDER.index(cn) > idx:
            ch.addprevious(el); return el
    spPr.append(el); return el

def set_fill(spPr, color):
    for t in FILL_TAGS:
        for e in spPr.findall(t): spPr.remove(e)
    sppr_insert(spPr, solid(color) if color else oxml_parser.makeelement(A("noFill")))

def set_line(spPr, color=None, w=None):
    for e in spPr.findall(A("ln")): spPr.remove(e)
    ln = oxml_parser.makeelement(A("ln"))
    if w: ln.set("w", str(w))
    if color: ln.append(solid(color))
    else: sub(ln, A("noFill"))
    sppr_insert(spPr, ln); return ln

def flatten(spPr):
    for t in ("effectLst", "effectDag", "scene3d", "sp3d"):
        for e in spPr.findall(A(t)): spPr.remove(e)
    sppr_insert(spPr, oxml_parser.makeelement(A("effectLst")))

def hexcolor_of(fill_el):
    if fill_el is None: return None
    c = fill_el.find(A("srgbClr"))
    if c is not None: return c.get("val")
    c = fill_el.find(A("schemeClr"))
    if c is not None: return "scheme:" + c.get("val")
    return None

def classify(hexv):
    """map an arbitrary text color to palette."""
    if hexv is None: return None
    if hexv.startswith("scheme:"):
        v = hexv[7:]
        if v in ("bg1", "lt1"): return WHITE
        if v in ("tx1", "dk1"): return TEXT
        return BLUE
    r, g, b = (int(hexv[i:i + 2], 16) / 255 for i in (0, 2, 4))
    h, s, v = colorsys.rgb_to_hsv(r, g, b)
    if s < 0.18:
        return WHITE if v > 0.92 else TEXT
    if h < 0.10 or h > 0.80: return EMPH
    if 0.10 <= h < 0.18: return EMPH  # orange / yellow-ish highlights
    return BLUE

def text_of(el):
    return "".join(t.text or "" for t in el.iter(A("t"))).strip()

def fix_rpr(rpr, force=None, white_to_dark=False):
    """fonts -> YaHei, drop text effects, recolor."""
    for e in rpr.findall(A("effectLst")): rpr.remove(e)
    for e in rpr.findall(A("highlight")): rpr.remove(e)
    for tag in ("latin", "ea"):
        el = rpr.find(A(tag))
        if el is None:
            el = oxml_parser.makeelement(A(tag))
            # order: ln, fill, effect, highlight, uLnTx.., latin, ea, cs, sym, hlinkClick...
            after = None
            for t in ("cs", "sym", "hlinkClick", "hlinkMouseOver", "rtl", "extLst"):
                after = rpr.find(A(t))
                if after is not None: break
            if tag == "latin" and rpr.find(A("ea")) is not None: after = rpr.find(A("ea"))
            if after is not None: after.addprevious(el)
            else: rpr.append(el)
        el.set("typeface", FONT)
        for k in ("panose", "pitchFamily", "charset"): el.attrib.pop(k, None)
    fill = rpr.find(A("solidFill"))
    for t in ("gradFill", "pattFill"):
        for e in rpr.findall(A(t)):
            rpr.remove(e)
    if force:
        if fill is not None: rpr.remove(fill)
        new = solid(force); _insert_rpr_fill(rpr, new)
        return
    if fill is not None:
        c = classify(hexcolor_of(fill))
        if c == WHITE and white_to_dark: c = TEXT
        if c:
            rpr.remove(fill); _insert_rpr_fill(rpr, solid(c))
    elif white_to_dark:
        _insert_rpr_fill(rpr, solid(TEXT))

def _insert_rpr_fill(rpr, fill):
    ln = rpr.find(A("ln"))
    if ln is not None: ln.addnext(fill)
    else: rpr.insert(0, fill)

def fix_text(txBody, force=None, white_to_dark=False):
    if txBody is None: return
    for p in txBody.iter(A("p")):
        ppr = p.find(A("pPr"))
        if ppr is not None:
            bc = ppr.find(A("buClr"))
            if bc is not None:
                for ch in list(bc): bc.remove(ch)
                sub(bc, A("srgbClr"), val=BLUE)
        for tag in ("rPr", "endParaRPr", "defRPr"):
            for rpr in p.iter(A(tag)):
                fix_rpr(rpr, force=force, white_to_dark=white_to_dark)
        # runs without rPr
        for r in p.findall(A("r")):
            if r.find(A("rPr")) is None:
                rpr = oxml_parser.makeelement(A("rPr")); rpr.set("lang", "zh-CN"); r.insert(0, rpr)
                fix_rpr(rpr, force=force, white_to_dark=white_to_dark)
    lst = txBody.find(A("lstStyle"))
    if lst is not None:
        for rpr in lst.iter(A("defRPr")): fix_rpr(rpr, force=force, white_to_dark=white_to_dark)

def eff_fill(sp):
    spPr = sp.find(P("spPr"))
    for t in FILL_TAGS:
        e = spPr.find(t) if spPr is not None else None
        if e is not None:
            n = etree.QName(e).localname
            if n == "noFill": return None
            if n == "solidFill": return hexcolor_of(e) or "x"
            if n == "gradFill":
                c = e.find(".//" + A("srgbClr"))
                return c.get("val") if c is not None else "scheme:accent1"
            return "x"
    st = sp.find(P("style"))
    if st is not None:
        fr = st.find(A("fillRef"))
        if fr is not None and fr.get("idx") not in ("0", None): return "scheme:accent1"
    return None

def has_line(sp):
    spPr = sp.find(P("spPr"))
    ln = spPr.find(A("ln")) if spPr is not None else None
    if ln is not None:
        if ln.find(A("noFill")) is not None: return False
        if ln.find(A("solidFill")) is not None or ln.find(A("gradFill")) is not None: return True
    st = sp.find(P("style"))
    if st is not None:
        lr = st.find(A("lnRef"))
        if lr is not None and lr.get("idx") not in ("0", None): return True
    return False

def is_whiteish(c):
    if c is None: return True
    if c.startswith("scheme:"): return c[7:] in ("bg1", "lt1")
    if not re.fullmatch(r"[0-9A-Fa-f]{6}", c): return False
    r, g, b = (int(c[i:i + 2], 16) for i in (0, 2, 4))
    return min(r, g, b) > 238

def prst(sp):
    g = sp.find(P("spPr") + "/" + A("prstGeom"))
    return g.get("prst") if g is not None else None

def set_prst(sp, name, adj=None):
    g = sp.find(P("spPr") + "/" + A("prstGeom"))
    if g is None: return
    g.set("prst", name)
    av = g.find(A("avLst"))
    if av is None: av = sub(g, A("avLst"))
    for ch in list(av): av.remove(ch)
    if adj is not None: sub(av, A("gd"), name="adj", fmla="val %d" % adj)

ARROWS = {"downArrow", "rightArrow", "leftArrow", "upArrow", "leftBrace", "rightBrace"}

def restyle_shape(el, mode="card"):
    """el: p:sp / p:cxnSp / p:pic / p:graphicFrame / p:grpSp"""
    tag = etree.QName(el).localname
    if tag == "grpSp":
        gp = el.find(P("grpSpPr"))
        if gp is not None:
            for t in ("effectLst", "scene3d", "sp3d"):
                for e in gp.findall(A(t)): gp.remove(e)
        for ch in el:
            if etree.QName(ch).localname in ("sp", "cxnSp", "pic", "grpSp", "graphicFrame"):
                restyle_shape(ch, mode)
        return
    if tag == "pic":
        spPr = el.find(P("spPr")); flatten(spPr); return
    if tag == "graphicFrame":
        restyle_table(el); return
    spPr = el.find(P("spPr"))
    if spPr is None: return
    flatten(spPr)
    for d in spPr.iter(A("prstDash")): d.set("val", "solid")
    st = el.find(P("style"))
    if st is not None:
        er = st.find(A("effectRef"))
        if er is not None: er.set("idx", "0")
    txt = text_of(el)
    p = prst(el)
    if tag == "cxnSp" or p == "line" or p in ("straightConnector1", "bentConnector3"):
        ln = spPr.find(A("ln"))
        w = ln.get("w") if ln is not None else None
        head = ln.find(A("headEnd")) if ln is not None else None
        tail = ln.find(A("tailEnd")) if ln is not None else None
        new = set_line(spPr, "7FA7CF", w or 12700)
        for e in (head, tail):
            if e is not None: new.append(copy.deepcopy(e))
        return
    if p in ("cube",) or p is None and spPr.find(A("custGeom")) is None:
        pass
    if p == "cube": set_prst(el, "rect")
    if p in ("wedgeRectCallout",): set_prst(el, "roundRect", 6000)
    if p == "wedgeEllipseCallout": set_prst(el, "ellipse")
    f = eff_fill(el); line = has_line(el)
    txBody = el.find(P("txBody"))
    if p in ARROWS:
        if p == "leftBrace":
            set_fill(spPr, None); set_line(spPr, BLUE, 19050)
        else:
            set_fill(spPr, LINE); set_line(spPr, None)
        fix_text(txBody); return
    if txt:
        if (f and not is_whiteish(f)) or line:
            set_fill(spPr, LIGHT); set_line(spPr, None)
            fix_text(txBody, white_to_dark=True)
        else:
            fix_text(txBody, white_to_dark=(f is None))
    else:
        if f and not is_whiteish(f):
            set_fill(spPr, LIGHT); set_line(spPr, None)
        elif line:
            ln = spPr.find(A("ln")); w = ln.get("w") if ln is not None else None
            set_line(spPr, "7FA7CF", w)
        fix_text(txBody)

def restyle_table(gf):
    tbl = gf.find(".//" + A("tbl"))
    if tbl is None:
        return
    rows = tbl.findall(A("tr"))
    for ri, tr in enumerate(rows):
        for tc in tr.findall(A("tc")):
            txBody = tc.find(A("txBody"))
            header = ri == 0
            fix_text(txBody, force=WHITE if header else TEXT)
            if header and txBody is not None:
                for rpr in txBody.iter(A("rPr")): rpr.set("b", "1")
            tcPr = tc.find(A("tcPr"))
            if tcPr is None: tcPr = sub(tc, A("tcPr"))
            for ch in list(tcPr):
                if etree.QName(ch).localname in ("lnL", "lnR", "lnT", "lnB", "lnTlToBr", "lnBlToTr",
                                                 "noFill", "solidFill", "gradFill", "blipFill",
                                                 "pattFill", "grpFill", "cell3D"):
                    tcPr.remove(ch)
            tcPr.set("anchor", "ctr")
            for side in ("lnL", "lnR", "lnT", "lnB"):
                ln = oxml_parser.makeelement(A(side)); ln.set("w", "9525")
                ln.append(solid("D3DEEA" if not header else BLUE))
                tcPr.insert(len([c for c in tcPr if etree.QName(c).localname.startswith("ln")]), ln)
            fillc = BLUE if header else (WHITE if ri % 2 == 1 else "F5F9FD")
            tcPr.append(solid(fillc))
    tp = tbl.find(A("tblPr"))
    if tp is not None:
        for e in tp.findall(A("tableStyleId")): tp.remove(e)
        for e in list(tp):
            if etree.QName(e).localname in ("effectLst", "solidFill", "gradFill"): tp.remove(e)

def remove_anim(slide):
    el = slide._element
    for t in ("timing", "transition"):
        for e in el.findall(P(t)): el.remove(e)
    for e in el.findall("{http://schemas.openxmlformats.org/markup-compatibility/2006}AlternateContent"):
        el.remove(e)

def move(shape, dx, dy):
    if shape.left is not None: shape.left = shape.left + dx
    if shape.top is not None: shape.top = shape.top + dy

def place(shape, x=None, y=None, w=None, h=None):
    if x is not None: shape.left = I(x)
    if y is not None: shape.top = I(y)
    if w is not None: shape.width = I(w)
    if h is not None: shape.height = I(h)

def add_title_ph(slide, text):
    spTree = slide.shapes._spTree
    ids = [int(e.get("id")) for e in spTree.iter(P("cNvPr"))] + [1]
    xml = ('<p:sp xmlns:p="%s" xmlns:a="%s"><p:nvSpPr><p:cNvPr id="%d" name="标题"/>'
           '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>'
           '<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="zh-CN" altLang="en-US" dirty="0"><a:latin typeface="微软雅黑"/><a:ea typeface="微软雅黑"/></a:rPr>'
           '<a:t></a:t></a:r></a:p></p:txBody></p:sp>') % (NS["p"], NS["a"], max(ids) + 1)
    el = parse_xml(xml)
    el.find(".//" + A("t")).text = text
    # insert right after grpSpPr so it is at the back
    spTree.insert(2, el)
    return el

def add_box(slide, x, y, w, h, fill=None, line=None, shape=MSO_SHAPE.RECTANGLE, name=None):
    s = slide.shapes.add_shape(shape, I(x), I(y), I(w), I(h))
    spPr = s._element.spPr
    set_fill(spPr, fill); set_line(spPr, line); flatten(spPr)
    st = s._element.find(P("style"))
    if st is not None: s._element.remove(st)
    if name: s.name = name
    return s

def add_text(slide, x, y, w, h, runs, size=18, color=TEXT, bold=False, align=PP_ALIGN.LEFT,
             anchor=MSO_ANCHOR.TOP, shape=None, fill=None, name=None, margin=0.0, spacing=None):
    if shape is None:
        s = slide.shapes.add_textbox(I(x), I(y), I(w), I(h))
    else:
        s = add_box(slide, x, y, w, h, fill=fill, shape=shape)
    tf = s.text_frame; tf.word_wrap = True
    tf.margin_left = tf.margin_right = I(margin); tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    paras = runs if isinstance(runs, list) else [runs]
    for i, ptxt in enumerate(paras):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = align
        if spacing: para.space_after = Pt(spacing)
        items = ptxt if isinstance(ptxt, tuple) else ((ptxt, {}),)
        for t, opt in items:
            r = para.add_run(); r.text = t
            f = r.font; f.size = Pt(opt.get("size", size)); f.bold = opt.get("bold", bold)
            f.color.rgb = RGBColor.from_string(opt.get("color", color))
            f.name = FONT
            rpr = r._r.get_or_add_rPr()
            rpr.set("lang", "zh-CN"); rpr.set("altLang", "en-US")
            ea = rpr.find(A("ea"))
            if ea is None: ea = sub(rpr, A("ea"))
            ea.set("typeface", FONT)
    if name: s.name = name
    return s

