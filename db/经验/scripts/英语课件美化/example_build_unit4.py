"""Restyle the Unit 4 deck in place via PowerPoint COM.
Keeps every shape id (so animations / triggers / media survive), never touches textbook pictures.
usage: python build.py SRC.pptx OUT.pptx BG.png
"""
import sys, os, win32com.client
from PIL import ImageFont

SRC, OUT, BG = [os.path.abspath(p) for p in sys.argv[1:4]]
LAT, EA = "Comic Sans MS", "微软雅黑"
G, GD, GL, GXL, MAT = "3E7C4F", "2E5E3B", "B9D8AE", "F1F8EE", "DCEED3"
OR, ORL, ORB, TX, RED, GRAY = "E9862A", "FFF5E6", "F3C383", "333333", "E53935", "7A7A7A"

_R = ImageFont.truetype("C:/Windows/Fonts/comic.ttf", 1000)
_B = ImageFont.truetype("C:/Windows/Fonts/comicbd.ttf", 1000)
def tw(s, pt, bold=False):
    w = 0.0
    f = _B if bold else _R
    for ch in s:
        w += pt if ord(ch) > 0x2E80 else f.getlength(ch) * pt / 1000
    return w

def rgb(h):
    h = h.lstrip('#'); return int(h[0:2], 16) + int(h[2:4], 16) * 256 + int(h[4:6], 16) * 65536

app = win32com.client.gencache.EnsureDispatch("PowerPoint.Application")
if "Microsoft Office" not in app.Path:
    raise SystemExit(f"ABORT: PowerPoint automation is being served by {app.Path} (WPS?). Close WPS and re-run.")
_mine = {os.path.abspath(a).lower() for a in sys.argv[1:3] if a.lower().endswith('.pptx')}
for _p in list(app.Presentations):
    if _p.FullName.lower() in _mine:
        _p.Close()
pres = app.Presentations.Open(SRC, False, False, False)
log = []

def all_shapes(sl):
    for s in sl.Shapes:
        yield s
        if s.Type == 6:
            for g in s.GroupItems:
                yield g

def S(sl, sid):
    for s in all_shapes(sl):
        if s.Id == sid:
            return s
    raise KeyError((sl.SlideIndex, sid))

def delete(sl, *ids):
    for i in ids:
        try:
            S(sl, i).Delete()
        except KeyError:
            log.append(f"S{sl.SlideIndex}: missing id {i} for delete")

def geo(s, l=None, t=None, w=None, h=None):
    if w is not None: s.Width = w
    if h is not None: s.Height = h
    if l is not None: s.Left = l
    if t is not None: s.Top = t

def txt(s, size=None, bold=None, color=None, align=None, anchor=None, wrap=None, auto=None,
        m=None, spacing=None, italic=None, nospace=False):
    tf = s.TextFrame2
    if auto is not None: tf.AutoSize = auto
    if wrap is not None: tf.WordWrap = -1 if wrap else 0
    if m is not None:
        tf.MarginLeft, tf.MarginRight, tf.MarginTop, tf.MarginBottom = m
    if anchor: tf.VerticalAnchor = {'t': 1, 'm': 3, 'b': 4}[anchor]
    tr = tf.TextRange
    if size: tr.Font.Size = size
    if bold is not None: tr.Font.Bold = -1 if bold else 0
    if italic is not None: tr.Font.Italic = -1 if italic else 0
    if color: tr.Font.Fill.ForeColor.RGB = rgb(color)
    if align: tr.ParagraphFormat.Alignment = {'l': 1, 'c': 2, 'r': 3}[align]
    if spacing:
        tr.ParagraphFormat.LineRuleWithin = -1
        tr.ParagraphFormat.SpaceWithin = spacing
    if nospace:
        tr.ParagraphFormat.SpaceBefore = 0
        tr.ParagraphFormat.SpaceAfter = 0

def style(s, fill=None, line=None, lw=None, noline=False, nofill=False, shadow=None, rounded=False):
    if rounded:
        try: s.AutoShapeType = 5
        except Exception as e: log.append(f"rounded fail {s.Name}: {e}")
    if nofill: s.Fill.Visible = 0
    elif fill:
        s.Fill.Visible = -1; s.Fill.Solid(); s.Fill.ForeColor.RGB = rgb(fill); s.Fill.Transparency = 0
    if noline: s.Line.Visible = 0
    elif line:
        s.Line.Visible = -1; s.Line.ForeColor.RGB = rgb(line); s.Line.DashStyle = 1
        if lw: s.Line.Weight = lw
    if shadow is not None: s.Shadow.Visible = -1 if shadow else 0

def recolor_dark(s, color):
    """recolor runs that are black / gray / navy (keeps red / orange highlight runs)"""
    tr = s.TextFrame.TextRange
    n = tr.Runs().Count
    for k in range(1, n + 1):
        r = tr.Runs(k, 1)
        c = r.Font.Color.RGB
        R, Gc, B = c & 255, (c >> 8) & 255, (c >> 16) & 255
        if not (R > 170 and R - Gc > 60):        # not a red/orange highlight
            r.Font.Color.RGB = rgb(color)

# ---------------------------------------------------------------- tag / header
def pill(sl, T, R=None, size=26):
    t = S(sl, T)
    if R is not None:
        S(sl, R).Delete()
    try:
        t.AutoShapeType = 5
    except Exception as e:
        log.append(f"S{sl.SlideIndex} pill autoshape fail: {e}")
    style(t, fill=G, noline=True, shadow=False)
    text = t.TextFrame2.TextRange.Text.strip()
    w = round(tw(text, size, True) + 44)
    txt(t, size=size, bold=True, color="FFFFFF", align='c', anchor='m', wrap=False, auto=0,
        m=(8, 8, 0, 0), nospace=True)
    f2 = t.TextFrame2.TextRange.Font
    try:
        f2.Line.Visible = 0
        f2.Shadow.Visible = 0
        f2.Glow.Radius = 0
    except Exception as e:
        log.append(f"S{sl.SlideIndex} tag text effects: {e}")
    geo(t, 80, 20, w, 40)
    return 80 + w

def place_icons(hx, ct):
    hx.LockAspectRatio = -1; hx.Width = 42; hx.Left = 30; hx.Top = 19
    ct.LockAspectRatio = -1; ct.Width = 20; ct.Left = 41; ct.Top = 30

def icons(sl, hexid, cartid):
    place_icons(S(sl, hexid), S(sl, cartid))

def add_icons(sl):
    src = pres.Slides(5)
    got = []
    for sid in (557, 558):
        S(src, sid).Copy()
        r = sl.Shapes.Paste()
        got.append(r.Item(1))
    place_icons(*got)

# ---------------------------------------------------------------- global
def has_cjk(t):
    return any(0x2E80 <= ord(c) <= 0x9FFF or 0xF900 <= ord(c) <= 0xFAFF or 0xFF00 <= ord(c) <= 0xFFEF
               or 0x3000 <= ord(c) <= 0x303F for c in t)

def global_fonts(sl):
    for s in all_shapes(sl):
        if s.Type == 6 or not s.HasTextFrame:
            continue
        tr = s.TextFrame.TextRange
        if tr.Length == 0:
            continue
        ea = EA if has_cjk(tr.Text) else LAT
        try:
            tr.Font.Name = LAT
            tr.Font.NameFarEast = ea
        except Exception as e:
            n = tr.Runs().Count
            for k in range(1, n + 1):
                try:
                    tr.Runs(k, 1).Font.Name = LAT; tr.Runs(k, 1).Font.NameFarEast = ea
                except Exception as e2:
                    log.append(f"S{sl.SlideIndex} font fail {s.Name} run{k} {tr.Runs(k,1).Text!r}: {e2}")
        try:
            s.TextFrame2.TextRange.Font.NameFarEast = ea
        except Exception:
            pass
        if sl.SlideIndex not in (1, 27):
            s.TextFrame2.TextRange.Font.Spacing = 0
        n = tr.Runs().Count
        for k in range(1, n + 1):
            r = tr.Runs(k, 1)
            if 0 < r.Font.Size < 24:
                r.Font.Size = 24
        npar = tr.Paragraphs().Count
        for k in range(1, npar + 1):
            p = tr.Paragraphs(k, 1)
            if 0 < p.Font.Size < 24:
                p.Font.Size = 24

def fit(s):
    """shrink/grow a single-line text shape to its text (PowerPoint metrics); returns (w, h)"""
    tf = s.TextFrame2
    tf.WordWrap = 0
    tf.AutoSize = 1
    w, h = s.Width, s.Height
    tf.AutoSize = 0
    return w, h

BULB = os.path.join(os.path.dirname(BG), "tips_bulb.png")
TIPS_TX, TIPS_ZH = "444444", "8A7356"
import re as _re

def bulb(sl, x, yc, size=30, sync_with=None):
    """the one Tips icon; if the tips text is animated, the icon enters together with it"""
    pic = sl.Shapes.AddPicture(BULB, 0, -1, x, yc - size / 2, size, size)
    pic.Name = "tipsIcon"
    if sync_with is not None:
        seq = sl.TimeLine.MainSequence
        for k in range(1, seq.Count + 1):
            e = seq.Item(k)
            if e.Shape.Id == sync_with.Id and not e.Exit:
                ne = seq.AddEffect(pic, e.EffectType, 0, 2, -1)    # 2 = with previous
                ne.MoveTo(k + 1)
                ne.Timing.TriggerType = 2
                ne.Timing.Duration = e.Timing.Duration
                ne.Timing.TriggerDelayTime = e.Timing.TriggerDelayTime
                break
    return pic

def tips_style(t, text=None, label=None, size=24):
    """borderless tips text: dark gray body, orange bold label + item numbers, muted Chinese"""
    if text is not None:
        t.TextFrame.TextRange.Text = text
    style(t, nofill=True, noline=True)
    txt(t, size=size, bold=False, italic=False, color=TIPS_TX, align='l', wrap=False, auto=0,
        m=(0, 0, 0, 0), nospace=True)
    tr = t.TextFrame.TextRange
    tr.Font.Name = LAT; tr.Font.NameFarEast = EA
    full = tr.Text
    for m_ in _re.finditer(r'[\u2e80-\u9fff\uff00-\uffef\u3000-\u303f]+', full):
        if m_.group(0).strip('：！？，。（）'):
            tr.Characters(m_.start() + 1, m_.end() - m_.start()).Font.Color.RGB = rgb(TIPS_ZH)
    for m_ in _re.finditer(r'(?<![\w.])\d+\.', full):
        c = tr.Characters(m_.start() + 1, m_.end() - m_.start())
        c.Font.Color.RGB = rgb(OR); c.Font.Bold = -1
    if label:
        i = full.find(label)
        c = tr.Characters(i + 1, len(label))
        c.Font.Color.RGB = rgb(OR); c.Font.Bold = -1; c.Font.Size = 26

def tips_label(t):
    """a stand-alone Tips heading (TASK TIPS / Acting Tips)"""
    style(t, nofill=True, noline=True)
    txt(t, size=26, bold=True, italic=False, color=OR, align='l', anchor='m', wrap=False, auto=0,
        m=(0, 0, 0, 0), nospace=True)

def setup_master():
    m = pres.SlideMaster
    m.Background.Fill.UserPicture(BG)
    for i in range(1, m.CustomLayouts.Count + 1):
        L = m.CustomLayouts(i)
        for k in range(L.Shapes.Count, 0, -1):
            s = L.Shapes(k)
            if s.Type != 14:            # keep placeholders, drop decorative frames / bars
                s.Delete()
        L.FollowMasterBackground = -1
    try:
        fs = pres.Designs(1).SlideMaster.Theme.ThemeFontScheme
        for part in (fs.MajorFont, fs.MinorFont):
            part.Item(1).Name = LAT
            part.Item(3).Name = EA
    except Exception as e:
        pres.Close()
        raise SystemExit(f"ABORT: PowerPoint COM is in a bad state (theme access failed: {e}). Re-run the build.")
    for sl in pres.Slides:
        sl.FollowMasterBackground = -1

setup_master()
P = {i: pres.Slides(i) for i in range(1, pres.Slides.Count + 1)}
for i, sl in P.items():
    if i != 10:
        global_fonts(sl)

# ================================================================ S1 cover
s = P[1]
style(S(s, 20), rounded=True)
ph = S(s, 19)
ph.LockAspectRatio = -1
ph.Height = 500; ph.Left = (960 - ph.Width) / 2; ph.Top = 20
try:
    ph.AutoShapeType = 5
    ph.Adjustments.SetItem(1, 0.035)
except Exception as e:
    log.append(f"cover photo rounding: {e}")
t = S(s, 10); recolor_dark(t, GD)
t = S(s, 251)
w = tw("How do people make choices?", 28, True) + 20
txt(t, size=28, bold=True, italic=True, color=OR, align='c', wrap=False, auto=0, m=(0, 0, 0, 0))
geo(t, 480 - w / 2, 300, w, 40)
t = S(s, 11)
style(t, fill=G)
txt(t, size=36, bold=True, color="FFFFFF")

# ================================================================ S5 (first: icon source) — click-any-card mind map
s = P[5]
delete(s, 540)
pill(s, 542, 541); icons(s, 557, 558)
t = S(s, 543); txt(t, size=34, bold=True, color=GD, align='c', wrap=False, auto=0, anchor='m'); geo(t, 20, 72, 920, 50)
C = win32com.client.constants
seq = s.TimeLine.MainSequence                     # old fixed-order reveal goes away
for k in range(seq.Count, 0, -1):
    seq.Item(k).Delete()
delete(s, 544)                                    # old curve artwork (purple/blue)
OR_D, G_T, OR_T = "C2671A", "EEF6EA", "FDF1E3"
HUB = (380, 282, 200, 100)
hub = S(s, 545)
hub.AutoShapeType = 5; hub.Adjustments.SetItem(1, 0.5)
style(hub, fill=G, noline=True)
geo(hub, *HUB)
hub.Shadow.Visible = -1; hub.Shadow.OffsetX = 0; hub.Shadow.OffsetY = 4; hub.Shadow.Blur = 10
hub.Shadow.ForeColor.RGB = rgb(GD); hub.Shadow.Transparency = 0.7
t = S(s, 546); txt(t, size=28, bold=True, color="FFFFFF", align='c', anchor='m', wrap=True, auto=0,
                   m=(10, 10, 0, 0), spacing=0.9, nospace=True); geo(t, *HUB)
CW, CH = 240, 104
cards = ((547, 11, (548,), 40, 160, G, GD, G_T),
         (549, 2, (550,), 40, 400, G, GD, G_T),
         (551, 4, (552, 553), 680, 160, OR, OR_D, OR_T),
         (554, 3, (555, 556), 680, 400, OR, OR_D, OR_T))
for box, mask, texts, x, y, col, dark, tint in cards:
    left = x < 480
    hx = HUB[0] + 4 if left else HUB[0] + HUB[2] - 4
    hy = HUB[1] + HUB[3] / 2
    cx = x + CW if left else x
    cy = y + CH / 2
    mid = (hx + cx) / 2
    ff = s.Shapes.BuildFreeform(C.msoEditingCorner, hx, hy)
    ff.AddNodes(C.msoSegmentCurve, C.msoEditingCorner, mid, hy, mid, cy, cx, cy)
    br = ff.ConvertToShape(); br.Name = "branch"
    br.Fill.Visible = 0; br.Line.Visible = -1; br.Line.ForeColor.RGB = rgb(col); br.Line.Weight = 4
    br.Shadow.Visible = 0
    br.ZOrder(1)
    b = S(s, box)
    b.AutoShapeType = 5; b.Adjustments.SetItem(1, 0.16)
    style(b, fill="FFFFFF", line=col, lw=2.5, shadow=False); geo(b, x, y, CW, CH)
    if len(texts) == 1:
        t = S(s, texts[0]); txt(t, size=36, bold=True, color=dark, align='c', anchor='m', wrap=False, auto=0, m=(0, 0, 0, 0))
        geo(t, x, y, CW, CH)
    else:
        t = S(s, texts[0]); txt(t, size=32, bold=True, color=dark, align='c', anchor='m', wrap=False, auto=0, m=(0, 0, 0, 0))
        geo(t, x, y + 10, CW, 48)
        t = S(s, texts[1]); txt(t, size=24, bold=False, color="6B6B6B", align='c', anchor='m', wrap=False, auto=0, m=(0, 0, 0, 0))
        geo(t, x, y + 58, CW, 36)
    mk = S(s, mask)
    mk.AutoShapeType = 5; mk.Adjustments.SetItem(1, 0.16)
    style(mk, fill=tint, line=col, lw=2.5, shadow=False); geo(mk, x, y, CW, CH)
    mk.Line.DashStyle = C.msoLineDash
    mk.TextFrame.TextRange.Text = "?"
    txt(mk, size=48, bold=True, color=col, align='c', anchor='m', wrap=False, auto=0, m=(0, 0, 0, 0))
    mk.TextFrame.TextRange.Font.Name = LAT; mk.TextFrame.TextRange.Font.NameFarEast = LAT
    dot = s.Shapes.AddShape(9, cx - 7, cy - 7, 14, 14); dot.Name = "branchDot"
    style(dot, fill=col, line="FFFFFF", lw=2, shadow=False)
    # click this card (in any order) -> the "?" cover zooms away and the answer pops in
    isq = s.TimeLine.InteractiveSequences.Add()
    e = isq.AddTriggerEffect(mk, C.msoAnimEffectFadedZoom, C.msoAnimTriggerOnShapeClick, mk)
    e.Exit = -1; e.Timing.Duration = 0.35
    for tid in texts:
        e2 = isq.AddEffect(S(s, tid), C.msoAnimEffectFadedZoom, 0, C.msoAnimTriggerWithPrevious)
        e2.Timing.TriggerType = C.msoAnimTriggerWithPrevious
        e2.Timing.Duration = 0.5; e2.Timing.TriggerDelayTime = 0.2

# ================================================================ S2
s = P[2]
delete(s, 12)
pill(s, 11); add_icons(s)
style(S(s, 474), line=GL, lw=2)
card = S(s, 476); style(card, fill=GXL, line=GL, lw=1.5); geo(card, 24, 352, 450, 166)
t = S(s, 477); txt(t, size=24, bold=False, color=GD, align='c', anchor='m', wrap=False, auto=0, m=(0, 0, 0, 0)); geo(t, 24, 362, 450, 40)
t = S(s, 478); txt(t, size=24, color=TX, align='l', anchor='m', wrap=False, auto=0, m=(0, 0, 0, 0)); geo(t, 104, 410, 400, 36)
t = S(s, 480); txt(t, size=24, color=TX, align='l', anchor='m', wrap=False, auto=0, m=(0, 0, 0, 0)); geo(t, 104, 456, 400, 36)
tk = S(s, 60); tk.LockAspectRatio = -1; tk.Width = 44; geo(tk, 52, 406)
card = S(s, 483); style(card, fill=GXL, line=GL, lw=1.5); geo(card, 486, 352, 450, 166)
t = S(s, 484); txt(t, size=24, bold=False, color=GD, align='c', anchor='m', wrap=False, auto=0, m=(0, 0, 0, 0)); geo(t, 486, 362, 450, 40)
t = S(s, 485); txt(t, size=24, color=TX, align='l', anchor='m', wrap=False, auto=0, m=(0, 0, 0, 0)); geo(t, 566, 410, 368, 36)
t = S(s, 486); txt(t, size=24, color=TX, align='l', anchor='m', wrap=False, auto=0, m=(0, 0, 0, 0)); geo(t, 566, 456, 368, 36)
tk = S(s, 8); tk.LockAspectRatio = -1; tk.Width = 44; geo(tk, 514, 452)

# ================================================================ S3
s = P[3]
delete(s, 12)
pill(s, 6); add_icons(s)
style(S(s, 494), line=GL, lw=2); style(S(s, 496), line=GL, lw=2)
for card, q, a, x in ((501, 502, 503, 33), (505, 506, 507, 507)):
    c = S(s, card); style(c, fill=GXL, line=GL, lw=1.5); geo(c, x, 372, 420, 140)
    t = S(s, q); txt(t, size=24, align='c', anchor='m', wrap=False, auto=0, m=(0, 0, 0, 0)); geo(t, x, 386, 420, 40)
    t = S(s, a); txt(t, size=24, align='c', anchor='m', wrap=False, auto=0, m=(0, 0, 0, 0)); geo(t, x, 438, 420, 40)
for pid in (13, 2):
    t = S(s, pid); cx, cy = t.Left + t.Width / 2, t.Top + t.Height / 2
    txt(t, size=26, bold=True, align='c', anchor='m', wrap=False, auto=0, m=(0, 0, 0, 0)); geo(t, cx - 52, cy - 23, 104, 46)

# ================================================================ S4
s = P[4]
delete(s, 11, 8)
pill(s, 12); add_icons(s)
style(S(s, 2), line=GL, lw=2)
c = S(s, 5); style(c, fill=GXL, line=GL, lw=1.5); geo(c, 60, 312, 840, 184)
t = S(s, 6); txt(t, size=24, color=GD, align='c', anchor='m', wrap=False, auto=0, m=(0, 0, 0, 0)); geo(t, 60, 322, 840, 40)
t = S(s, 7); txt(t, size=24, color=TX, align='l', anchor='t', wrap=False, auto=0, m=(0, 0, 0, 0), spacing=1.15, nospace=True); geo(t, 196, 366, 690, 78)
t = S(s, 10); txt(t, size=24, color=TX, align='l', anchor='t', wrap=False, auto=0, m=(0, 0, 0, 0), spacing=1.15, nospace=True); geo(t, 196, 444, 690, 40)
tk = S(s, 9); geo(tk, 148, 401)

# ================================================================ S6 — Ben's photo wall (twine, pegs, fairy lights, polaroids)
import math
s = P[6]
delete(s, 560, 561, 562, 563)
pill(s, 564); add_icons(s)
C = win32com.client.constants

def rot_place(shape, cx, cy, dx, dy, ang, w=None, h=None):
    a = math.radians(ang)
    px = cx + dx * math.cos(a) - dy * math.sin(a)
    py = cy + dx * math.sin(a) + dy * math.cos(a)
    if w is not None: shape.Width = w
    if h is not None: shape.Height = h
    shape.Rotation = 0
    shape.Left = px - shape.Width / 2; shape.Top = py - shape.Height / 2
    shape.Rotation = ang % 360

def polaroid(fr, cx, cy, W, H, ang):
    fr.AutoShapeType = 1
    style(fr, fill="FFFFFF", line="E6E0D4", lw=0.75)
    fr.Shadow.Visible = -1; fr.Shadow.OffsetX = 2; fr.Shadow.OffsetY = 4; fr.Shadow.Blur = 9
    fr.Shadow.ForeColor.RGB = rgb("6B5B45"); fr.Shadow.Transparency = 0.7
    rot_place(fr, cx, cy, 0, 0, ang, W, H)

def twine(x0, x1, y0, sag):
    k = 4 / 3 * sag
    ff = s.Shapes.BuildFreeform(C.msoEditingCorner, x0, y0)
    ff.AddNodes(C.msoSegmentCurve, C.msoEditingCorner, x0 + (x1 - x0) / 3, y0 + k, x0 + 2 * (x1 - x0) / 3, y0 + k, x1, y0)
    ln = ff.ConvertToShape(); ln.Name = "twine"
    ln.Fill.Visible = 0; ln.Line.Visible = -1; ln.Line.ForeColor.RGB = rgb("A67C52"); ln.Line.Weight = 2.25
    ln.Shadow.Visible = 0
    ln.ZOrder(1)
    for x in (x0, x1):
        pin = s.Shapes.AddShape(9, x - 5, y0 - 5, 10, 10); pin.Name = "twinePin"
        style(pin, fill="8C6239", noline=True, shadow=False)
    return lambda x: y0 + 4 * sag * ((x - x0) / (x1 - x0)) * (1 - (x - x0) / (x1 - x0))

def lights(yf, xs):
    cols = ("FFC845", "F2994A", "8CC084", "FFC845")
    for i, x in enumerate(xs):
        y = yf(x)
        sock = s.Shapes.AddShape(1, x - 3, y - 1, 6, 6); sock.Name = "lightCap"
        style(sock, fill="5B6B4A", noline=True, shadow=False)
        bl = s.Shapes.AddShape(9, x - 5.5, y + 4, 11, 14); bl.Name = "fairyLight"
        style(bl, fill=cols[i % 4], noline=True, shadow=False)
        try:
            bl.Glow.Radius = 7; bl.Glow.Color.RGB = rgb(cols[i % 4]); bl.Glow.Transparency = 0.55
        except Exception as e:
            log.append(f"glow: {e}")

M, MB = 9, 20
XS = (385, 600, 815)
ROWS = ((76, 14, ((573, 572, -5), (575, 574, 3), (577, 576, -3))),
        (232, 14, ((579, 578, 4), (581, 580, -2), (583, 582, 5))))
clips = []
for y0, sag, items in ROWS:
    yf = twine(262, 926, y0, sag)
    lights(yf, (284, 492, 707, 904))
    for (pid, fid, ang), xc in zip(items, XS):
        ph, fr = S(s, pid), S(s, fid)
        W, H = ph.Width + 2 * M, ph.Height + M + MB
        cy = yf(xc) + 7 + H / 2
        polaroid(fr, xc, cy, W, H, ang)
        rot_place(ph, xc, cy, 0, (M + ph.Height / 2) - H / 2, ang)
        clip = s.Shapes.AddShape(5, 0, 0, 9, 28); clip.Name = "clothespin"
        style(clip, fill="D8A35D", line="A87332", lw=0.75, shadow=False)
        rot_place(clip, xc, cy, 0, -H / 2 + 3, ang)
        clips.append((pid, clip))

# Ben's own big polaroid, taped to the wall
ben, bfr, lab = S(s, 566), S(s, 565), S(s, 567)
BW, BH = ben.Width + 24, ben.Height + 12 + 54
BX, BY, BA = 140, 236, -4
polaroid(bfr, BX, BY, BW, BH, BA)
rot_place(ben, BX, BY, 0, (12 + ben.Height / 2) - BH / 2, BA)
txt(lab, size=30, bold=True, color=GD, align='c', anchor='m', wrap=False, auto=0, m=(0, 0, 0, 0))
rot_place(lab, BX, BY, 0, BH / 2 - 27, BA, w=ben.Width, h=44)
tape = s.Shapes.AddShape(1, 0, 0, 90, 26); tape.Name = "washiTape"
style(tape, fill="F6C36B", noline=True, shadow=False); tape.Fill.Transparency = 0.2
rot_place(tape, BX, BY, 0, -BH / 2 + 2, BA - 6)

t = S(s, 568); txt(t, size=34, bold=True, color=GD, align='c', anchor='m', wrap=False, auto=0); geo(t, 20, 398, 920, 50)
c = S(s, 569); style(c, fill="FFFFFF", line=GL, lw=2)
try:
    c.Adjustments.SetItem(1, 0.5)
except Exception:
    pass
geo(c, 340, 456, 280, 48)
t = S(s, 571); txt(t, size=24, align='c', anchor='m', wrap=False, auto=0, m=(0, 0, 0, 0)); geo(t, 340, 456, 280, 48)

# same clicks, same order — photos now pop onto the wall one after another
seq = s.TimeLine.MainSequence
PH = (573, 575, 577, 579, 581, 583); FR = (572, 574, 576, 578, 580, 582)
delay = {566: 0, 567: 0.2}
for i, (p_, f_) in enumerate(zip(PH, FR)):
    delay[p_] = delay[f_] = 0.15 * i
delay[564] = 0.15 * 6 + 0.1
for k in range(1, seq.Count + 1):
    e = seq.Item(k); sid = e.Shape.Id
    if sid in delay:
        e.EffectType = C.msoAnimEffectFadedZoom; e.Timing.Duration = 0.4
        if e.Timing.TriggerType == 2:
            e.Timing.TriggerDelayTime = delay[sid]
    else:
        e.EffectType = C.msoAnimEffectFade; e.Timing.Duration = 0.5

def add_with(shape, after_id, d, dur=0.4):
    for k in range(1, seq.Count + 1):
        if seq.Item(k).Shape.Id == after_id:
            ne = seq.AddEffect(shape, C.msoAnimEffectFadedZoom, 0, C.msoAnimTriggerWithPrevious, -1)
            ne.MoveTo(k + 1)
            ne.Timing.TriggerType = 2; ne.Timing.TriggerDelayTime = d; ne.Timing.Duration = dur
            return
for pid, clip in clips:
    add_with(clip, pid, delay[pid] + 0.1, 0.3)
add_with(bfr, 566, 0)
add_with(tape, 566, 0.25, 0.3)

# ================================================================ S7
s = P[7]
delete(s, 585)
pill(s, 590, 589); icons(s, 598, 599)
q1, q2 = S(s, 593), S(s, 596)
txt(q1, size=32, bold=True, color=GD, align='l', anchor='m', m=(0, 0, 0, 0))
txt(q2, size=32, bold=True, color=OR, align='l', anchor='m', m=(0, 0, 0, 0))
w1, _ = fit(q1); w2, _ = fit(q2)
x0 = 480 - (w1 + 24 + w2) / 2
geo(q1, x0, 63, w1, 42); geo(q2, x0 + w1 + 24, 63, w2, 42)

# ================================================================ S8
s = P[8]
delete(s, 2)
right = pill(s, 4, 3); icons(s, 11, 12)
style(S(s, 5), fill=MAT, noline=True)
b = S(s, 8); geo(b, 296, 106, 328, 46)
t = S(s, 9); txt(t, size=24, align='c', anchor='m', wrap=False, auto=0, m=(4, 4, 0, 0)); geo(t, 296, 106, 328, 46)
mk = S(s, 13); geo(mk, 300, 109, 320, 40)
for q in (10, 16):
    t = S(s, q); txt(t, size=24, bold=True, color=GD, align='l', anchor='m', wrap=False, auto=0, m=(0, 0, 0, 0)); geo(t, right + 14, 20, 360, 40)
t = S(s, 15)
parts = [x.strip() for x in t.TextFrame.TextRange.Text.replace('\x0b', '\r').split('\r') if x.strip()]
lab, rest = parts[0].split('：', 1)
tips_style(t, f"{lab}：{rest.strip()}\r{' '.join(parts[1:])}", label=f"{lab}：")
txt(t, anchor='t', spacing=0.9); geo(t, 748, 18, 190, 62)
t.TextFrame.TextRange.Paragraphs(2, 1).IndentLevel = 2
ind = tw(lab, 26, True) + 26
t.TextFrame.Ruler.Levels(2).FirstMargin = ind; t.TextFrame.Ruler.Levels(2).LeftMargin = ind
bulb(s, 712, 35, sync_with=t)
geo(S(s, 5), 70, 78, 820, 389)

# ================================================================ S9 / S11
for idx in (9, 11):
    s = P[idx]
    delete(s, 44, 52, 42)
    right = pill(s, 16, 43); icons(s, 51, 17)
s = P[9]
b = S(s, 45); style(b, fill="FFFFFF", line=G, lw=2); geo(b, 350, 20, 576, 92)
t = S(s, 2); txt(t, size=24, align='c', anchor='m', wrap=True, auto=0, m=(6, 6, 2, 2), spacing=0.95, nospace=True); geo(t, 350, 20, 576, 92)
t.TextFrame.TextRange.Paragraphs(1, 1).Font.Color.RGB = rgb(GD)

# ================================================================ S12
s = P[12]
delete(s, 74, 75, 76)
t = S(s, 81); recolor_dark(t, GD)
t = S(s, 82); txt(t, size=28, bold=True, color=OR, align='c', anchor='m', wrap=False, auto=0); geo(t, 180, 294, 600, 44)

# ================================================================ S13
s = P[13]
delete(s, 62)
pill(s, 64, 63); icons(s, 71, 72)

# ================================================================ S14
s = P[14]
delete(s, 85, 86, 11, 12, 110, 90, 95, 100)
pill(s, 14, 13); icons(s, 15, 16)
t = S(s, 89); tips_label(t); geo(t, 88, 116, 260, 36)
bulb(s, 50, 134)
recolor_dark(S(s, 108), GD)
rows = ((91, 92, 93, 94), (96, 97, 98, 99), (101, 102, 103, 104))
for k, (circ, ico, en, zh) in enumerate(rows):
    top = 158 + k * 74
    geo(S(s, circ), 50, top + 10, 44, 44)
    ic = S(s, ico); geo(ic, 50 + (44 - ic.Width) / 2, top + 10 + (44 - ic.Height) / 2)
    t = S(s, en); tips_style(t, size=26); txt(t, anchor='t'); geo(t, 110, top, 480, 38)
    t = S(s, zh); tips_style(t); txt(t, anchor='t'); geo(t, 110, top + 38, 480, 32)
style(S(s, 109), fill="FFFFFF", line=GL, lw=2)
ln = S(s, 113); ln.Width = ln.Left + ln.Width - 224; ln.Left = 224
ln = S(s, 116); ln.Width = ln.Left + ln.Width - 322; ln.Left = 322

# ================================================================ S15 / S16 / S13-like headers
s = P[15]; delete(s, 118, 119); pill(s, 123, 122); icons(s, 120, 121)
s = P[16]; delete(s, 132, 133); pill(s, 137, 136); icons(s, 134, 135)
style(S(s, 138), fill=MAT, noline=True); style(S(s, 139), fill="FFFFFF", noline=True)

# ================================================================ S17
s = P[17]
delete(s, 149, 133, 65)
right = pill(s, 154, 153); icons(s, 151, 152)
t = S(s, 155); txt(t, size=24, color="8A7356", align='l', anchor='m', wrap=False, auto=0, m=(0, 0, 0, 0)); geo(t, 82, 62, 300, 34)
# small round avatars (head crops) re-spaced so the 24pt names don't collide
for aid, cx in ((12, 52), (15, 142), (18, 232), (21, 330)):
    a = S(s, aid); a.Left = cx - a.Width / 2
labels = {13: (52, 64), 16: (142, 84), 19: (232, 84), 22: (330, 120)}
for tid, (cx, w) in labels.items():
    t = S(s, tid); txt(t, size=24, color=TX, align='c', anchor='t', wrap=False, auto=0, m=(0, 0, 0, 0), spacing=0.9, nospace=True)
    geo(t, cx - w / 2, 165, w, 66 if tid == 22 else 34)
# scene thumbnails: same size, aligned inside the frame (was x=2..11)
for iid in (24, 28, 32):
    S(s, iid).Left = 24
XC = 404
t = S(s, 36); txt(t, size=24, align='l', anchor='m', wrap=False, auto=0, m=(4, 4, 0, 0)); geo(t, XC, 64, 300, 36)
t = S(s, 35); tips_label(t); geo(t, XC + 38, 106, 300, 36)
bulb(s, XC, 124)
t = S(s, 38)
paras = [p.strip() for p in t.TextFrame2.TextRange.Text.replace('\x0b', '\r').split('\r') if p.strip()]
merged = [f"{paras[i]}  {paras[i + 1]}" for i in range(0, len(paras) - 1, 2)]
t.TextFrame2.TextRange.Text = '\r'.join(merged)
tips_style(t); txt(t, anchor='t', spacing=0.95)
geo(t, XC + 38, 144, 936 - XC - 38, 92)
bars = ((23, 25, 26, 238, 79), (27, 29, 30, 328, 77), (31, 33, 34, 417, 82))
for bar, lab, dia, y, h in bars:
    b = S(s, bar); style(b, fill=GXL, line=GL, lw=1.25); geo(b, 36, y, 900, h)
    t = S(s, lab); txt(t, size=24, bold=True, color=OR, align='l', anchor='m', wrap=False, auto=0, m=(0, 0, 0, 0)); geo(t, 184, y, 104, h)
    t = S(s, dia); txt(t, size=24, align='l', anchor='m', wrap=False, auto=0, m=(0, 0, 0, 0), spacing=0.95, nospace=True); geo(t, 290, y, 636, h)

# ================================================================ S18
s = P[18]
delete(s, 203)
pill(s, 208, 207); icons(s, 205, 206)
t = S(s, 18); txt(t, size=28, bold=True, color=GD, align='l', anchor='m', m=(0, 0, 0, 0))
w, _ = fit(t); geo(t, 80, 88, w, 50)
t = S(s, 17)
parts = [_re.sub(r'^(\d\.)(\S)', r'\1 \2', x.strip()) for x in t.TextFrame.TextRange.Text.replace('\x0b', '\r').split('\r') if x.strip()]
tips_style(t, f"{parts[0]}{parts[1]}\r" + "\r".join(parts[2:]), label=parts[0])
txt(t, anchor='t', spacing=0.9); geo(t, 650, 20, 290, 136)
bulb(s, 614, 37)
t = S(s, 15); txt(t, size=24, align='c', anchor='m', wrap=False, auto=0, m=(0, 0, 0, 0)); geo(t, 150, 422, 508, 36)

# ================================================================ S19
s = P[19]
delete(s, 231, 232)
pill(s, 236, 235); icons(s, 233, 234)
t = S(s, 237); txt(t, size=32, bold=True, color=GD, align='c', anchor='m', wrap=False, auto=0); geo(t, 60, 66, 840, 48)
style(S(s, 238), fill=MAT, noline=True); style(S(s, 239), fill="FFFFFF", noline=True)

# ================================================================ S20 / S22
for idx in (20, 22):
    s = P[idx]
    delete(s, 260, 261)
    pill(s, 265, 264); icons(s, 262, 263)
    style(S(s, 266), fill=MAT, noline=True); style(S(s, 267), fill="FFFFFF", noline=True)
    t = S(s, 276); txt(t, size=24, align='c', anchor='m', wrap=True, auto=0, m=(4, 4, 0, 0), spacing=0.95, nospace=True); geo(t, 148, 188, 274, 76)
    t = S(s, 278); txt(t, size=24, align='c', anchor='m', wrap=True, auto=0, m=(4, 4, 0, 0), spacing=0.95, nospace=True); geo(t, 565, 206, 232, 92)

# ================================================================ S21
s = P[21]
delete(s, 2, 3, 4, 5, 6, 7)
t1, t2 = S(s, 8), S(s, 9)
txt(t1, size=34, bold=True, align='l', anchor='m', m=(0, 0, 0, 0))
recolor_dark(t1, GD)
txt(t2, size=24, color=GRAY, align='l', anchor='m', m=(0, 0, 0, 0))
w1, _ = fit(t1); w2, _ = fit(t2)
x0 = 480 - (w1 + 8 + w2) / 2
geo(t1, x0, 18, w1, 44); geo(t2, x0 + w1 + 8, 19, w2, 42)
for qb, qt, x, y in ((12, 13, 744, 138), (18, 19, 744, 334), (24, 25, 298, 138), (30, 31, 298, 334)):
    b = S(s, qb); geo(b, x, y, 164, 96)
    t = S(s, qt); txt(t, size=24, bold=False, align='c', anchor='m', wrap=True, auto=0, m=(2, 2, 0, 0), spacing=0.9, nospace=True); geo(t, x + 6, y + 2, 152, 92)
b = S(s, 34); geo(b, 42, 448, 876, 70)
t = S(s, 35); txt(t, size=24, align='c', anchor='m', wrap=False, auto=0, m=(0, 0, 0, 0), spacing=0.85, nospace=True); geo(t, 52, 452, 856, 32)
t = S(s, 36); txt(t, size=24, align='c', anchor='m', wrap=False, auto=0, m=(0, 0, 0, 0), spacing=0.85, nospace=True); geo(t, 52, 484, 856, 30)

# ================================================================ S23
s = P[23]; delete(s, 306); pill(s, 308, 307); icons(s, 317, 318)

# ================================================================ S24
s = P[24]
delete(s, 320, *range(321, 336), 336)
pill(s, 337); add_icons(s)
t = S(s, 338); txt(t, size=48, bold=True, color=GD, align='c', anchor='m', wrap=False, auto=0); geo(t, 100, 72, 760, 66)
t = S(s, 339); txt(t, size=26, color=OR, align='c', anchor='m', wrap=False, auto=0); geo(t, 180, 138, 600, 40)
t = S(s, 342); txt(t, size=24, wrap=False, auto=0, anchor='m'); t.Width = 120
t = S(s, 351); txt(t, size=26, bold=True, color=GD, align='c', anchor='m', wrap=False, auto=0, m=(0, 0, 0, 0)); geo(t, 666, 462, 100, 38)

# ================================================================ S25
s = P[25]
delete(s, 353, *range(354, 361), 361)
right = pill(s, 362); add_icons(s)
t = S(s, 363); txt(t, size=32, bold=True, color=GD, align='l', anchor='m', wrap=False, auto=0, m=(0, 0, 0, 0)); geo(t, right + 20, 20, 400, 40)
c = S(s, 364); style(c, fill="FFFFFF", line=GL, lw=2); geo(c, 24, 74, 360, 404)
c = S(s, 365); style(c, fill=ORL, line=ORB, lw=1); geo(c, 44, 90, 150, 36)
t = S(s, 366); txt(t, size=24, bold=True, color=OR, align='c', anchor='m', wrap=False, auto=0, m=(0, 0, 0, 0)); geo(t, 44, 90, 150, 36)
geo(S(s, 367), 332, 88); geo(S(s, 368), 332, 88)
L = 44
rowsy = {369: 132, 370: 178, 371: 220, 372: 264, 374: 264, 375: 308, 377: 352, 378: 352, 380: 396}
for tid, y in rowsy.items():
    t = S(s, tid)
    txt(t, size=28 if tid == 369 else 24, align='l', anchor='m', wrap=False, auto=0, m=(0, 0, 0, 0), nospace=True)
    if tid == 369: txt(t, bold=True, color=GD)
    geo(t, None, y, None, 38)
geo(S(s, 369), L, None, 300)
geo(S(s, 370), L, None, 330); geo(S(s, 371), L, None, 330)
txt(S(s, 378), bold=True)
wb, _ = fit(S(s, 372)); wt, _ = fit(S(s, 374)); ww, _ = fit(S(s, 375)); fit(S(s, 377)); wn, _ = fit(S(s, 378))
for tid in (372, 374, 375, 377, 378):
    S(s, tid).Height = 38
geo(S(s, 372), L, None); geo(S(s, 373), L + wb + 6, 296, 100, 2); geo(S(s, 374), L + wb + 112)
geo(S(s, 375), L); geo(S(s, 376), L + ww + 6, 340, 220, 2)
geo(S(s, 377), L); geo(S(s, 378), L + ww + 8); geo(S(s, 379), L + ww + 8 + wn + 6, 384, 160, 2)
geo(S(s, 380), L, None, 160); geo(S(s, 381), L, 458, 170, 2)
txt(S(s, 378), bold=True)
XR, WR = 398, 538
for k, (grp, rect, emo, lab) in enumerate(((382, 383, 384, 385), (386, 387, 388, 389), (390, 391, 392, 393))):
    y = 74 + k * 50
    geo(S(s, rect), XR, y, WR, 42)
    t = S(s, emo); txt(t, size=24, align='c', anchor='m', wrap=False, auto=0, m=(0, 0, 0, 0)); geo(t, XR + 10, y + 3, 36, 36)
    t = S(s, lab); txt(t, size=24, align='l', anchor='m', wrap=False, auto=0, m=(0, 0, 0, 0)); geo(t, XR + 52, y + 3, WR - 64, 36)
t = S(s, 394); txt(t, size=26, bold=True, color=OR, align='l', anchor='m', wrap=False, auto=0, m=(0, 0, 0, 0)); geo(t, XR + 2, 228, 170, 38)
t = S(s, 395); txt(t, size=24, color=GRAY, align='l', anchor='m', wrap=False, auto=0, m=(0, 0, 0, 0)); geo(t, XR + 176, 228, 120, 38)
chips = (((396, 397), (398, 399), (400, 401)), ((402, 403), (404, 405), (406, 407)))
colx = ((XR, 244), (XR + 250, 288))
for ci, col in enumerate(chips):
    x, w = colx[ci]
    for ri, (bg, tt) in enumerate(col):
        y = 272 + ri * 50
        geo(S(s, bg), x, y, w, 42)
        t = S(s, tt); txt(t, size=24, align='c', anchor='m', wrap=False, auto=0, m=(0, 0, 0, 0)); geo(t, x, y, w, 42)
        t.TextFrame2.TextRange.Font.Spacing = -0.75
t = S(s, 408); txt(t, size=24, color=GRAY, align='l', anchor='m', wrap=False, auto=0, m=(0, 0, 0, 0)); geo(t, XR + 2, 426, 536, 38)

# ================================================================ S26
s = P[26]
delete(s, 410, 411)
t = S(s, 424); txt(t, size=48, bold=True, color=OR, align='c', anchor='m', wrap=False, auto=0); geo(t, 280, 30, 400, 66)
import math
def put_line(parts, x0, y0):
    x = x0
    for tid, w in parts:
        t = S(s, tid)
        txt(t, wrap=False, auto=0, m=(0, 0, 0, 0), anchor='m')
        geo(t, x, y0 - (x - x0) * math.tan(math.radians(2)), w, 32)
        x += w
for tid in (428, 429, 430, 431, 432, 434, 436, 437, 439):
    txt(S(s, tid), size=24)
txt(S(s, 428), size=26, bold=True)
W = {tid: fit(S(s, tid))[0] for tid in (429, 430, 431, 432, 434, 436, 437, 439)}
put_line(((429, W[429]), (430, W[430]), (431, W[431])), 284, 194)
put_line(((432, W[432]),), 285, 248)
geo(S(s, 2), 285 + W[432] + 6, 272)
put_line(((434, W[434]),), 287, 304)
geo(S(s, 435), 287 + W[434] + 6, 318); geo(S(s, 436), 287 + W[434] + 6 + 135 + 6, 296)
put_line(((437, W[437]),), 289, 360)
geo(S(s, 438), 289 + W[437] + 6, 374)

# ================================================================ S27
s = P[27]
delete(s, 471)
geo(S(s, 472), 170, 40, 470, 470)
geo(S(s, 473), 450, 160, 350, 350)
S(s, 489).Top = 404
def reflow(ids, cx=480):
    ws = []
    for tid in ids:
        t = S(s, tid); top, h = t.Top, t.Height
        t.TextFrame2.MarginLeft = 0; t.TextFrame2.MarginRight = 0
        w, _ = fit(t); t.Top = top; t.Height = h
        ws.append(w)
    x = cx - sum(ws) / 2
    for tid, w in zip(ids, ws):
        S(s, tid).Left = x; x += w
for punct, word in ((484, 483), (487, 486)):
    a, b = S(s, punct), S(s, word)
    fa, fb = a.TextFrame2, b.TextFrame2
    fa.TextRange.Font.Size = fb.TextRange.Font.Size
    fa.TextRange.Font.Bold = fb.TextRange.Font.Bold
    fa.TextRange.Font.Fill.ForeColor.RGB = fb.TextRange.Font.Fill.ForeColor.RGB
    fa.VerticalAnchor = fb.VerticalAnchor
    fa.MarginTop, fa.MarginBottom = fb.MarginTop, fb.MarginBottom
    a.Top, a.Height = b.Top, b.Height
reflow((482, 483, 484)); reflow((485, 486, 487))

# ================================================================ Q&A unified style
# Every question and every answer stem (答语) in the deck: Comic Sans MS, bold, title green.
# A question and its answer always share one size: 28pt on open slides, 24pt inside cards /
# header rows / bubbles. Only teaching highlights (is/are, rainy day, NEED, FUTURE) keep their colour.
QA_BIG, QA_SMALL = 28, 24

def qa(t, size, align=None, spacing=None):
    txt(t, size=size, bold=True, italic=False, align=align, nospace=True)
    if spacing is not None:
        t.TextFrame2.TextRange.Font.Spacing = spacing
    recolor_dark(t, GD)

def qa_line(t, size, x, y, w, h, align='c'):
    qa(t, size, align=align)
    txt(t, anchor='m', wrap=False, auto=0, m=(0, 0, 0, 0))
    geo(t, x, y, w, h)

# S1 cover: the unit question
t = S(P[1], 251); qa(t, QA_BIG, align='c'); t.TextFrame.TextRange.Font.Color.RGB = rgb(GD); w, _ = fit(t); geo(t, 480 - w / 2, 298, w, 44)

# S2 / S3 / S4 picture cards: question + options / answer frame, 24pt
s = P[2]
for tid in (477, 484):
    qa(S(s, tid), QA_SMALL, align='c', spacing=-0.8)
for tid in (478, 480, 485, 486):
    qa(S(s, tid), QA_SMALL, align='l', spacing=-0.8)
s = P[3]
for qid in (502, 506):
    qa(S(s, qid), QA_SMALL, align='c')
for aid, word, qid, key in ((503, "It's", 502, " is "), (507, "They're", 506, " are ")):
    qtr = S(s, qid).TextFrame.TextRange
    kred = qtr.Characters(qtr.Text.find(key) + 2, len(key.strip())).Font.Color.RGB
    t = S(s, aid); qa(t, QA_SMALL, align='c')
    t.TextFrame.TextRange.Font.Color.RGB = rgb(GD)
    t.TextFrame.TextRange.Characters(1, len(word)).Font.Color.RGB = kred
s = P[4]
qa(S(s, 6), QA_SMALL, align='c')
qa(S(s, 7), QA_SMALL, align='l'); qa(S(s, 10), QA_SMALL, align='l')

# S5 page question
qa_line(S(P[5], 543), QA_BIG, 20, 74, 920, 44)

# S6 question + answer stem, identical, stacked (the pill that made the answer look different goes)
s = P[6]
delete(s, 569)
qa_line(S(s, 568), QA_BIG, 20, 404, 920, 44)
qa_line(S(s, 571), QA_BIG, 20, 452, 920, 44)

# S7 question + answer stem on one line, identical
s = P[7]
q1, q2 = S(s, 593), S(s, 596)
qa(q1, QA_BIG, align='l'); qa(q2, QA_BIG, align='l')
q1.TextFrame.TextRange.Font.Color.RGB = rgb(GD); q2.TextFrame.TextRange.Font.Color.RGB = rgb(GD)
w1, _ = fit(q1); w2, _ = fit(q2)
x0 = 480 - (w1 + 28 + w2) / 2
geo(q1, x0, 62, w1, 42); geo(q2, x0 + w1 + 28, 62, w2, 42)

# S8 header-row questions (swap animation), 24pt
for qid in (10, 16):
    qa(S(P[8], qid), QA_SMALL, align='l')

# S9 question bubble: English + Chinese both in the question style
t = S(P[9], 2); qa(t, QA_SMALL, align='c'); t.TextFrame.TextRange.Font.Color.RGB = rgb(GD)

# S14 question + "I will / I will NOT" answer frames, 28pt; wavy underline re-fitted under "save up"
s = P[14]
q = S(s, 87); qa_line(q, QA_BIG, 20, 64, 920, 44)
wq, _ = fit(q); qx = 480 - wq / 2; geo(q, qx, 64, wq, 44)
k = wq / tw("How will you save up your money?", QA_BIG, True)
wv = S(s, 88); wv.LockAspectRatio = 0
geo(wv, qx + tw("How will you ", QA_BIG, True) * k - 4, 100, tw("save up", QA_BIG, True) * k + 8, 14)
for tid, icon, ln in ((112, 111, 113), (115, 114, 116)):
    t = S(s, tid); qa(t, QA_BIG, align='l')
    w, _ = fit(t)
    ic = S(s, icon); cy = ic.Top + ic.Height / 2
    geo(t, 117, cy - 21, w, 42)
    l = S(s, ln); right = l.Left + l.Width
    l.Left = 117 + w + 12; l.Width = right - l.Left; l.Top = cy + 14

# S18 question + the two answer captions, 28pt (zoom card widened a little to hold them)
s = P[18]
qa(S(s, 18), QA_BIG, align='l')
zb = S(s, 13); geo(zb, 141, 179, 545, 287)
t = S(s, 15)
txt_now = t.TextFrame.TextRange.Text
parts = [x for x in txt_now.split('  ') if x.strip()]
t.TextFrame.TextRange.Text = parts[0].strip() + "   " + parts[-1].strip()
qa_line(t, QA_BIG, 146, 425, 535, 40)

# S19 page question
qa_line(S(P[19], 237), QA_BIG, 60, 68, 840, 44)

# S21 question (title + Chinese gloss) and its answer box: one size, 24pt (the answer sentence is long)
s = P[21]
t1, t2 = S(s, 8), S(s, 9)
qa(t1, QA_SMALL, align='l'); qa(t2, QA_SMALL, align='l'); t2.TextFrame.TextRange.Font.Color.RGB = rgb(GD)
w1, _ = fit(t1); w2, _ = fit(t2)
x0 = 480 - (w1 + 10 + w2) / 2
geo(t1, x0, 20, w1, 40); geo(t2, x0 + w1 + 10, 20, w2, 40)
for tid, y in ((35, 451), (36, 484)):
    qa_line(S(s, tid), QA_SMALL, 52, y, 856, 32)
    txt(S(s, tid), spacing=0.85)

# ================================================================ ROUND 5
import math
C = win32com.client.constants
UMBRELLA = os.path.join(os.path.dirname(BG), "umbrella.png")

def front(*shapes):
    for sh in shapes:
        sh.ZOrder(0)

def soft_shadow(sh, blur=16, dy=6, tr=0.68, col="000000"):
    sh.Shadow.Visible = -1; sh.Shadow.OffsetX = 0; sh.Shadow.OffsetY = dy; sh.Shadow.Blur = blur
    sh.Shadow.ForeColor.RGB = rgb(col); sh.Shadow.Transparency = tr

def eff_of(seq, sid, exit_=False):
    for k in range(1, seq.Count + 1):
        e = seq.Item(k)
        if e.Shape.Id == sid and bool(e.Exit) == exit_:
            return k, e
    return None, None

def add_after(seq, shape, after_k, effect, delay, dur):
    ne = seq.AddEffect(shape, effect, 0, C.msoAnimTriggerWithPrevious, -1)
    ne.MoveTo(after_k + 1)
    ne.Timing.TriggerType = C.msoAnimTriggerWithPrevious
    ne.Timing.TriggerDelayTime = delay; ne.Timing.Duration = dur
    return ne

# ---------------------------------------------------------------- 4. later layers fully cover earlier ones
# S3: the white "dim" veils cover the whole shop picture (they left a clear strip right / bottom)
s = P[3]
for veil, pic in ((43, 495), (3, 497)):
    v, pc = S(s, veil), S(s, pic)
    geo(v, pc.Left, pc.Top, pc.Width, pc.Height)
# S8: the second question swaps in on the same spot and now hides the first one completely
s = P[8]
q_a, q_b = S(s, 10), S(s, 16)
wq = max(q_a.Width, q_b.Width)
for q in (q_a, q_b):
    q.Width = wq
style(q_b, fill="FFFFFF", noline=True)
# S9 / S11: blank answer bubble covers the whole printed bubble (its full stop was peeking out)
for idx in (9, 11):
    mk = S(P[idx], 12); geo(mk, 444, 363, 326, 64)
    mk.Line.Visible = -1; mk.Line.ForeColor.RGB = rgb("C9C9C9"); mk.Line.Weight = 1.25
    mk.Line.DashStyle = C.msoLineDash
# S20 / S22: speech bubbles and their masks are a size larger than the printed bubbles underneath
for idx in (20, 22):
    s = P[idx]
    for sid in (269, 275):
        geo(S(s, sid), 130, 178, 300, 92)
    geo(S(s, 276), 136, 182, 288, 84)
    for sid in (271, 277):
        geo(S(s, sid), 548, 192, 266, 108)
    geo(S(s, 278), 554, 196, 254, 100)
geo(S(P[20], 279), 128, 294, 232, 60)
geo(S(P[22], 273), 128, 294, 232, 60); geo(S(P[22], 274), 135, 300, 218, 48)

# S18: the zoom card hides the whole picture while it is up, sits above the underline marks,
#      and its "close" step (lost when PowerPoint read the original file) is rebuilt
s = P[18]
zb, zi, zt = S(s, 13), S(s, 14), S(s, 15)
geo(zb, 78, 164, 730, 308)
try:
    zb.Adjustments.SetItem(1, 0.06)
except Exception:
    pass
style(zb, fill="FFFFFF", line=GL, lw=2.5); soft_shadow(zb, blur=14, dy=4, tr=0.75)
zi.Left = 78 + (730 - zi.Width) / 2; zi.Top = 174
qa_line(zt, QA_BIG, 88, 424, 710, 42)
front(zb, zi, zt)
seq = s.TimeLine.MainSequence
for k in range(seq.Count, 0, -1):
    if seq.Item(k).Exit:
        seq.Item(k).Delete()
e = seq.AddEffect(zb, C.msoAnimEffectFade, 0, C.msoAnimTriggerOnPageClick, -1); e.Exit = -1; e.Timing.Duration = 0.5
for shp in (zi, zt):
    e = seq.AddEffect(shp, C.msoAnimEffectFade, 0, C.msoAnimTriggerWithPrevious, -1)
    e.Exit = -1; e.Timing.TriggerType = C.msoAnimTriggerWithPrevious; e.Timing.Duration = 0.5

# ---------------------------------------------------------------- 5. S21: bigger English title, smaller gloss, centred pop-up answer
s = P[21]
DY = 44
for sid in range(10, 34):
    S(s, sid).Top = S(s, sid).Top + DY
for ov, card in ((26, 22), (32, 28)):
    c = S(s, card); geo(S(s, ov), c.Left, c.Top, c.Width, c.Height)
t1, t2 = S(s, 8), S(s, 9)
txt(t1, size=36, bold=True)
txt(t2, size=20, bold=False, color=TIPS_ZH)
w1, _ = fit(t1); w2, _ = fit(t2)
x0 = 480 - (w1 + 14 + w2) / 2
geo(t1, x0, 24, w1, 52); geo(t2, x0 + w1 + 14, 30, w2, 46)
# pop-up: dim veil + white card + umbrella badge
veil = s.Shapes.AddShape(5, 15, 15, 930, 510); veil.Name = "popupVeil"
try:
    veil.Adjustments.SetItem(1, 0.04)
except Exception:
    pass
style(veil, fill="1E3A28", noline=True, shadow=False); veil.Fill.Transparency = 0.42
card = S(s, 34)
card.AutoShapeType = 5
try:
    card.Adjustments.SetItem(1, 0.12)
except Exception:
    pass
style(card, fill="FFFFFF", line=OR, lw=3); soft_shadow(card, blur=24, dy=8, tr=0.6)
geo(card, 130, 150, 700, 240)
a1, a2 = S(s, 35), S(s, 36)
for t_ in (a1, a2):
    txt(t_, size=36, bold=True, align='c', anchor='m', auto=0, m=(0, 0, 0, 0), spacing=0.95, nospace=True)
txt(a1, wrap=True); geo(a1, 155, 196, 650, 100)
txt(a2, wrap=False); geo(a2, 155, 304, 650, 52)
tr2 = a2.TextFrame.TextRange
i = tr2.Text.find("FUTURE")
if i >= 0:
    tr2.Characters(i + 1, 6).Font.Color.RGB = rgb(OR)
front(card, a1, a2)
badge = s.Shapes.AddShape(9, 440, 108, 80, 80); badge.Name = "popupBadge"
style(badge, fill="FFFFFF", line=OR, lw=3); soft_shadow(badge, blur=10, dy=3, tr=0.7)
umb = s.Shapes.AddPicture(UMBRELLA, 0, -1, 452, 120, 56, 56); umb.Name = "popupUmbrella"
# timing: 1 s after the last tick finishes, everything pops up together
seq = s.TimeLine.MainSequence
k, e = eff_of(seq, 34)
e.EffectType = C.msoAnimEffectFadedZoom; e.Timing.TriggerType = C.msoAnimTriggerAfterPrevious
e.Timing.TriggerDelayTime = 1.0; e.Timing.Duration = 0.5
for sid in (35, 36):
    _, e2 = eff_of(seq, sid)
    e2.EffectType = C.msoAnimEffectFadedZoom; e2.Timing.TriggerType = C.msoAnimTriggerWithPrevious
    e2.Timing.TriggerDelayTime = 1.1; e2.Timing.Duration = 0.5
k, _ = eff_of(seq, 34)
add_after(seq, veil, k, C.msoAnimEffectFade, 1.0, 0.4)
k, _ = eff_of(seq, 36)
add_after(seq, badge, k, C.msoAnimEffectFadedZoom, 1.2, 0.4)
k, _ = eff_of(seq, badge.Id)
add_after(seq, umb, k, C.msoAnimEffectFadedZoom, 1.2, 0.4)

# ---------------------------------------------------------------- 1 + 3. S25: bigger postcard text, compact rules, hint removed
s = P[25]
delete(s, 408)
c = S(s, 364); geo(c, 24, 74, 384, 412)
geo(S(s, 365), 44, 90, 150, 36); geo(S(s, 366), 44, 90, 150, 36)
geo(S(s, 367), 352, 88); geo(S(s, 368), 352, 88)
L = 44
t = S(s, 369); txt(t, size=30, bold=True, color=GD, anchor='m', wrap=False, auto=0, m=(0, 0, 0, 0)); geo(t, L, 134, 320, 42)
for tid, y in ((370, 180), (371, 224), (372, 268), (374, 268), (375, 312), (377, 356), (378, 356), (380, 400)):
    t = S(s, tid); txt(t, size=28, anchor='m', wrap=False, auto=0, m=(0, 0, 0, 0), nospace=True)
    t.Top = y; t.Height = 40
txt(S(s, 378), bold=True)
W = {tid: fit(S(s, tid))[0] for tid in (370, 371, 372, 374, 375, 377, 378, 380)}
for tid, y in ((370, 180), (371, 224), (372, 268), (374, 268), (375, 312), (377, 356), (378, 356), (380, 400)):
    S(s, tid).Top = y; S(s, tid).Height = 40
geo(S(s, 370), L); geo(S(s, 371), L); geo(S(s, 372), L)
geo(S(s, 373), L + W[372] + 6, 300, 84, 2); geo(S(s, 374), L + W[372] + 96)
geo(S(s, 375), L); geo(S(s, 376), L + W[375] + 6, 344, 220, 2)
geo(S(s, 377), L); geo(S(s, 378), L + W[377] + 8); geo(S(s, 379), L + W[377] + 8 + W[378] + 6, 388, 170, 2)
geo(S(s, 380), L); geo(S(s, 381), L, 468, 180, 2)
XR, WR = 420, 516
for kk, (rect, emo, lab) in enumerate(((383, 384, 385), (387, 388, 389), (391, 392, 393))):
    y = 80 + kk * 48
    geo(S(s, rect), XR, y, WR, 40)
    t = S(s, emo); geo(t, XR + 10, y + 2, 36, 36)
    t = S(s, lab); txt(t, size=24); geo(t, XR + 52, y + 2, WR - 64, 36)
t = S(s, 394); txt(t, size=24, bold=True); geo(t, XR + 2, 228, 150, 36)
t = S(s, 395); txt(t, size=24); geo(t, XR + 150, 228, 120, 36)
for ci, (x, w) in enumerate(((XR, 238), (XR + 242, 274))):
    for ri, (bg, tt) in enumerate(chips[ci]):
        y = 274 + ri * 48
        geo(S(s, bg), x, y, w, 40)
        t = S(s, tt); geo(t, x, y, w, 40); t.TextFrame2.TextRange.Font.Spacing = -1.0 if ci == 0 else -1.5

# ---------------------------------------------------------------- 1 + 2. S26: bigger postcard text, missing "I will NOT" row added
s = P[26]
NOT_RED = S(P[25], 378).TextFrame.TextRange.Font.Color.RGB
grp = S(s, 425)
grp.Ungroup()
CX, CY, CW_, CH_, ANG = 480, 310, 480, 380, -2

def place(sh, dx, dy, w=None, h=None):
    a = math.radians(ANG)
    if w is not None: sh.Width = w
    if h is not None: sh.Height = h
    sh.Rotation = 0
    px = CX + dx * math.cos(a) - dy * math.sin(a)
    py = CY + dx * math.sin(a) + dy * math.cos(a)
    sh.Left = px - sh.Width / 2; sh.Top = py - sh.Height / 2
    sh.Rotation = ANG % 360

place(S(s, 426), 0, 0, CW_, CH_)
place(S(s, 427), 0, 0, CW_ - 16, CH_ - 16)
XL = -CW_ / 2 + 36
ROW = [-CH_ / 2 + 44 + 46 * i for i in range(7)]

def text_at(sh, x_left, row_y, size, bold=False, color=None):
    sh.Rotation = 0
    txt(sh, size=size, bold=bold, color=color, anchor='m', wrap=False, auto=0, m=(0, 0, 0, 0), nospace=True)
    w, _ = fit(sh)
    place(sh, x_left + w / 2, row_y, w, 40)
    return w

def line_at(sh, x_left, row_y, length):
    place(sh, x_left + length / 2, row_y + 17, length, 3)

text_at(S(s, 428), XL, ROW[0], 30, True, GD)
x = XL
for tid in (429, 430, 431):
    x += text_at(S(s, tid), x, ROW[1], 28)
w = text_at(S(s, 432), XL, ROW[2], 28); line_at(S(s, 2), XL + w + 4, ROW[2], 150)
w = text_at(S(s, 434), XL, ROW[3], 28); line_at(S(s, 435), XL + w + 6, ROW[3], 130)
text_at(S(s, 436), XL + w + 6 + 130 + 6, ROW[3], 28)
w = text_at(S(s, 437), XL, ROW[4], 28); line_at(S(s, 438), XL + w + 6, ROW[4], 210)
nt = S(s, 437).Duplicate().Item(1); nt.Name = "I will NOT"
nt.TextFrame.TextRange.Text = "I will NOT"
nt.TextFrame.TextRange.Characters(8, 3).Font.Color.RGB = NOT_RED
nt.TextFrame.TextRange.Characters(8, 3).Font.Bold = -1
w = text_at(nt, XL, ROW[5], 28)
nl = S(s, 438).Duplicate().Item(1); nl.Name = "I will NOT line"
line_at(nl, XL + w + 6, ROW[5], 170)
text_at(S(s, 439), XL, ROW[6], 28)
line_at(S(s, 440), XL, ROW[6] + 22, 170)

# ================================================================ ROUND 6
# S21: "Save it for the FUTURE!" becomes its own take-away banner and waits for a click
s = P[21]
a2 = S(s, 36)
txt(a2, size=36, bold=True, color="FFFFFF", align='c', anchor='m', wrap=False, auto=0, m=(0, 0, 0, 0))
tr2 = a2.TextFrame.TextRange
i = tr2.Text.find("FUTURE")
if i >= 0:
    tr2.Characters(i + 1, 6).Font.Color.RGB = rgb("FFD54F")
wa, _ = fit(a2)
BW, BH, BY = wa + 64, 56, 302
banner = s.Shapes.AddShape(5, 480 - BW / 2, BY, BW, BH); banner.Name = "takeawayBanner"
banner.Adjustments.SetItem(1, 0.5)
style(banner, fill=G, noline=True); soft_shadow(banner, blur=8, dy=3, tr=0.72, col=GD)
geo(a2, 480 - wa / 2, BY, wa, BH)
front(a2)
seq = s.TimeLine.MainSequence
k, e = eff_of(seq, 36)
e.MoveTo(seq.Count)
k, e = eff_of(seq, 36)
e.Timing.TriggerType = C.msoAnimTriggerOnPageClick
e.Timing.TriggerDelayTime = 0; e.Timing.Duration = 0.5; e.EffectType = C.msoAnimEffectFadedZoom
add_after(seq, banner, k, C.msoAnimEffectFadedZoom, 0, 0.5)

# S22: the blank answer bubble wears the same green edge as the blank on S20
src_b, dst_b = S(P[20], 279), S(P[22], 273)
dst_b.Line.Visible = -1; dst_b.Line.ForeColor.RGB = src_b.Line.ForeColor.RGB
dst_b.Line.Weight = src_b.Line.Weight; dst_b.Line.DashStyle = src_b.Line.DashStyle
dst_b.Shadow.Visible = src_b.Shadow.Visible

# ================================================================ ROUND 8
A_DIR = os.path.dirname(BG)

# ---------- S6: twine, pins and fairy lights arrive together with Ben's photo (click 1)
s = P[6]
seq = s.TimeLine.MainSequence
k, _ = eff_of(seq, 566)
for sh in list(s.Shapes):
    if sh.Name in ("twine", "twinePin", "lightCap", "fairyLight"):
        add_after(seq, sh, k, C.msoAnimEffectFade, 0, 0.4)

# ---------- S25 (round 9): bigger postcard on the left, lighter word banks on the right
s = P[25]
delete(s, 394, 395, *range(396, 408))
CL, CW = 24, 460
geo(S(s, 364), CL, 72, CW, 446)
geo(S(s, 365), CL + 20, 88, 160, 38); geo(S(s, 366), CL + 20, 88, 160, 38)
geo(S(s, 367), CL + CW - 50, 86); geo(S(s, 368), CL + CW - 50, 86)
L = CL + 22

def row(tid, y, size=30, bold=False, color=None):
    t = S(s, tid)
    txt(t, size=size, bold=bold, color=color, anchor='m', wrap=False, auto=0, m=(0, 0, 0, 0), nospace=True)
    w, _ = fit(t); t.Top = y; t.Height = 42
    return t, w

ROW = [136, 186, 234, 282, 330, 378, 426]
t, _ = row(369, ROW[0], 32, True, GD); t.Left = L
t, _ = row(370, ROW[1]); t.Left = L
ya = S(s, 371)
ya.TextFrame.TextRange.Text = "You are"
t, wy = row(371, ROW[2]); t.Left = L
ln = S(s, 373).Duplicate().Item(1); ln.Name = "you are line"
geo(ln, L + wy + 8, ROW[2] + 34, 200, 2)
ex = S(s, 374).Duplicate().Item(1); ex.Name = "you are !"
ex.TextFrame.TextRange.Text = "!"
t, we = row(ex.Id, ROW[2]); t.Left = L + wy + 8 + 200 + 4
t, wb = row(372, ROW[3]); t.Left = L
geo(S(s, 373), L + wb + 8, ROW[3] + 34, 120, 2)
t, wt = row(374, ROW[3]); t.Left = L + wb + 8 + 120 + 6
t, ww = row(375, ROW[4]); t.Left = L
geo(S(s, 376), L + ww + 8, ROW[4] + 34, 260, 2)
t, ww2 = row(377, ROW[5]); t.Left = L
t, wn = row(378, ROW[5], 30, True); t.Left = L + ww2 + 8
geo(S(s, 379), L + ww2 + 8 + wn + 8, ROW[5] + 34, 200, 2)
t, _ = row(380, ROW[6]); t.Left = L
geo(S(s, 381), L, ROW[6] + 72, 200, 2)

XR, WR = CL + CW + 16, 936 - (CL + CW + 16)
for kk, (rect, emo, lab) in enumerate(((383, 384, 385), (387, 388, 389), (391, 392, 393))):
    y = 76 + kk * 42
    geo(S(s, rect), XR, y, WR, 36)
    geo(S(s, emo), XR + 8, y, 36, 36)
    t = S(s, lab); txt(t, size=24); geo(t, XR + 46, y, WR - 54, 36)

def bank_pill(icon, text, zh, y, fill, line, col):
    tx = text + (" " + zh if zh else "")
    w = tw(text, 24, True) + (tw(" " + zh, 20) + 14 if zh else 0) + 66
    p_ = s.Shapes.AddShape(5, XR, y, w, 34); p_.Name = "bankPill"
    p_.Adjustments.SetItem(1, 0.5)
    style(p_, fill=fill, line=line, lw=1.25, shadow=False)
    s.Shapes.AddPicture(os.path.join(A_DIR, icon), 0, -1, XR + 10, y + 6, 22, 22).Name = "bankIcon"
    h = s.Shapes.AddTextbox(1, XR + 38, y, w - 44, 34); h.Name = "bankTitle"
    h.TextFrame.TextRange.Text = tx
    txt(h, size=24, bold=True, color=col, align='l', anchor='m', wrap=False, auto=0, m=(0, 0, 0, 0), nospace=True)
    tr = h.TextFrame.TextRange; tr.Font.Name = LAT; tr.Font.NameFarEast = EA
    if zh:
        c_ = tr.Characters(len(text) + 2, len(zh)); c_.Font.Size = 20; c_.Font.Bold = 0
    return y + 34

def bank_words(words, y, col):
    b = s.Shapes.AddTextbox(1, XR + 6, y + 4, WR - 6, 30); b.Name = "bankWords"
    b.TextFrame.TextRange.Text = " / ".join(w_.replace(" ", " ") for w_ in words)
    txt(b, size=24, bold=False, color=col, align='l', anchor='t', wrap=True, auto=1, m=(0, 0, 0, 0), spacing=0.9, nospace=True)
    tr = b.TextFrame.TextRange; tr.Font.Name = LAT; tr.Font.NameFarEast = LAT
    return b.Top + b.Height

y = 212
y = bank_pill("ic_star.png", "Describe Ben", "(评价Ben)", y, "FFF4C7", "E8C14A", "7A5A10")
y = bank_words(["great", "helpful", "kind", "hard-working"], y, "7A5A10") + 10
y = bank_pill("ic_check.png", "I will...", "", y, "E3F2DF", "8CC084", GD)
y = bank_words(["save up money", "work hard", "help parents"], y, GD) + 10
y = bank_pill("ic_cross.png", "I will NOT...", "", y, "FDE6E5", "E99A9A", "A93226")
y = bank_words(["buy too many snacks", "buy new toys", "buy ice cream"], y, "A93226")
if y > 518:
    log.append(f"S25 word banks end at {y:.0f}pt — too low")

# ---------- S21: the centre pop-up waits for a click like everything else
s = P[21]
seq = s.TimeLine.MainSequence
for kk in range(1, seq.Count + 1):
    e = seq.Item(kk); sh = e.Shape
    if sh.Id == 34:
        e.Timing.TriggerType = C.msoAnimTriggerOnPageClick; e.Timing.TriggerDelayTime = 0
    elif sh.Name == "popupVeil":
        e.Timing.TriggerDelayTime = 0
    elif sh.Id == 35:
        e.Timing.TriggerDelayTime = 0.1
    elif sh.Name in ("popupBadge", "popupUmbrella"):
        e.Timing.TriggerDelayTime = 0.2

pres.SaveAs(OUT)
print("\n".join(log) if log else "no warnings")
pres.Close()
