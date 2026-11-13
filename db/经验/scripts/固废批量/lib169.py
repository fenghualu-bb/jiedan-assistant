# -*- coding: utf-8 -*-
# True 16:9 re-layout; executed inside restyle.py after layout_fix.py
L, R = 0.6, 12.73
W = R - L
TOP, BOT = 1.55, 6.85

def delete(sp): sp._element.getparent().remove(sp._element)

def card_style(sp, fill=WHITE, inset=0.22, anchor=None):
    spPr = sp._element.spPr
    set_fill(spPr, fill); set_line(spPr, None)
    bp = sp._element.find(".//" + A("bodyPr"))
    if bp is not None:
        for k, v in (("lIns", inset), ("rIns", inset), ("tIns", 0.14), ("bIns", 0.14)):
            bp.set(k, str(I(v)))
        if anchor: bp.set("anchor", anchor)
        for e in list(bp):
            if etree.QName(e).localname in ("spAutoFit", "normAutofit"): bp.remove(e)

def set_sz(el, pt):
    for tag in ("rPr", "endParaRPr"):
        for r in el.iter(A(tag)): r.set("sz", str(int(pt * 100)))

def scale_sz(el, f):
    for tag in ("rPr", "endParaRPr"):
        for r in el.iter(A(tag)):
            if r.get("sz"): r.set("sz", str(int(int(r.get("sz")) * f)))

def no_bullet(p):
    ppr = p.find(A("pPr"))
    if ppr is None:
        ppr = oxml_parser.makeelement(A("pPr")); p.insert(0, ppr)
    ppr.set("marL", "0"); ppr.set("indent", "0")
    for t in ("buNone", "buChar", "buAutoNum", "buFont", "buClr", "buSzPct", "buBlip"):
        for e in ppr.findall(A(t)): ppr.remove(e)
    bn = oxml_parser.makeelement(A("buNone"))
    # buNone goes after lnSpc/spcBef/spcAft
    idx = 0
    for i, ch in enumerate(ppr):
        if etree.QName(ch).localname in ("lnSpc", "spcBef", "spcAft"): idx = i + 1
    ppr.insert(idx, bn)

def strip_lead_space(p):
    ts = p.findall(".//" + A("t"))
    if ts and ts[0].text: ts[0].text = ts[0].text.lstrip(" 　")

def para_box(slide, x, y, w, h, p_elems, fill=None, anchor="t", inset=0.22, bullets=False, size=None, name=None):
    s = add_text(slide, x, y, w, h, "", name=name)
    if fill is not None:
        card_style(s, fill, inset)
    txBody = s._element.find(P("txBody"))
    for p in txBody.findall(A("p")): txBody.remove(p)
    for p in p_elems:
        np_ = copy.deepcopy(p)
        if not bullets: no_bullet(np_)
        strip_lead_space(np_)
        txBody.append(np_)
    if size: set_sz(txBody, size)
    bp = txBody.find(A("bodyPr")); bp.set("anchor", anchor); bp.set("wrap", "square")
    return s

def paras_of(sp):
    return [p for p in sp._element.find(P("txBody")).findall(A("p")) if text_of(p)]

def number_card(slide, x, y, w, h, num):
    add_box(slide, x, y, w, h, fill=WHITE, name="卡片")
    add_text(slide, x + 0.25, y + 0.2, 1.5, 0.6, num, size=28, color=BLUE, bold=True, name="序号")

def stretch(n, src0=1.667, src1=11.667, keep_pics=True):
    f = W / (src1 - src0)
    for sp in S(n).shapes:
        if sp.is_placeholder and sp.placeholder_format.type == 1: continue
        if sp.left is None: continue
        x, w = sp.left / IN, sp.width / IN
        if sp.shape_type == 13 and keep_pics:
            c = L + (x + w / 2 - src0) * f
            sp.left = I(c - w / 2)
        else:
            sp.left = I(L + (x - src0) * f); sp.width = I(w * f)
            if sp.has_table:
                for col in sp.table.columns: col.width = int(col.width * f)

def shift(n, names, dx):
    for nm in names:
        sp = shp(n, nm); sp.left = sp.left + I(dx)

def left_align(sp):
    for p in sp.text_frame.paragraphs: p.alignment = PP_ALIGN.LEFT

