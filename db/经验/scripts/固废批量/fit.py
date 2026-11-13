# PowerPoint COM post-pass: measure real text heights, fix verticals/labels/overlaps/overflow/cards.
import sys, os, win32com.client, pythoncom
src = os.path.abspath(sys.argv[1]); out = os.path.abspath(sys.argv[2])
TOP, BOT, L, R = 111.0, 494.0, 43.0, 917.0
pythoncom.CoInitialize()
app = win32com.client.DispatchEx("PowerPoint.Application")
pres = app.Presentations.Open(src, ReadOnly=False, Untitled=False, WithWindow=False)
SKIPN = {"主标题底", "节标题底", "目录行"}
def is_title(s):
    try: return s.Type == 14 and s.PlaceholderFormat.Type in (1, 3)
    except Exception: return False
def has_text(s):
    try: return s.HasTextFrame and s.TextFrame.HasText
    except Exception: return False
def txt(s): return s.TextFrame.TextRange.Text.strip()
def ov(a, b):
    ox = min(a.Left + a.Width, b.Left + b.Width) - max(a.Left, b.Left)
    oy = min(a.Top + a.Height, b.Top + b.Height) - max(a.Top, b.Top)
    return ox, oy
log = []
for sl in pres.Slides:
    names = [s.Name for s in sl.Shapes]
    if any(n in SKIPN for n in names): continue
    card = None; texts = []; others = []
    for s in sl.Shapes:
        if s.Name == "文字卡片": card = s; continue
        if s.Name in ("页脚", "页码", "橙色标记") or is_title(s): continue
        if s.Type != 6 and has_text(s): texts.append(s)
        else: others.append(s)
    if not texts: continue
    for o in list(others):
        try:
            if (o.Type == 13 and o.Width < 64 and o.Height < 64 and 0.75 < o.Width / max(1, o.Height) < 1.33 and o.Name != "公式") or (o.Type == 1 and o.AutoShapeType == 9 and o.Width < 48 and o.Height < 48 and not has_text(o)):
                o.Delete(); others.remove(o)
        except Exception: pass
    diagram = any(o.Type in (6, 9) or (o.Type == 1 and o.Height < 3) for o in others) or len(texts) > 7
    # 1 horizontal text, wrapping on
    for t in texts:
        try:
            if t.TextFrame.Orientation != 1: t.TextFrame.Orientation = 1
            t.TextFrame.WordWrap = True
            if len(txt(t)) <= 20 and t.TextFrame.TextRange.Paragraphs().Count == 1 and t.Width < 400: t.TextFrame.WordWrap = False
        except Exception: pass
    # 2 side labels -> headings (text-only slides)
    if not diagram:
        for t in list(texts):
            if (t.Width < 125 or (t.Width < 210 and len(txt(t)) <= 12)) and len(txt(t)) >= 3:
                bodies = [b for b in texts if b is not t and b.Left > t.Left + 10 and b.Width > 220
                          and ov(t, b)[1] > 5]
                if bodies:
                    b = min(bodies, key=lambda b: b.Top)
                    nl = min(t.Left, b.Left); nw = b.Left + b.Width - nl
                    t.Left = nl; t.Width = nw; t.Top = b.Top
                    t.TextFrame.TextRange.ParagraphFormat.Alignment = 1
                    t.TextFrame.TextRange.Text = txt(t).replace("\r", "").replace("\n", "")
                    t.TextFrame.AutoSize = 1
                    for o in bodies:
                        o.Left = nl; o.Width = nw
                    log.append((sl.SlideIndex, "label"))
    for t in texts:
        try:
            pf = t.TextFrame.TextRange.ParagraphFormat
            for k in range(1, t.TextFrame.TextRange.Paragraphs().Count + 1):
                pp = t.TextFrame.TextRange.Paragraphs(k).ParagraphFormat
                if not pp.LineRuleWithin: pp.LineRuleWithin = -1; pp.SpaceWithin = 1.15
                elif pp.SpaceWithin < 1.0: pp.SpaceWithin = 1.1
        except Exception: pass
    # 3 real heights
    for t in texts:
        try: t.TextFrame.AutoSize = 1
        except Exception: pass
    def resolve():
        ordered = sorted(texts, key=lambda s: (s.Top, s.Left))
        for i, a in enumerate(ordered):
            for b in ordered[:i]:
                ox, oy = ov(a, b)
                if ox > 12 and oy > 2 and a.Top >= b.Top - 6 and (len(txt(a)) >= 4 or len(txt(b)) >= 4):
                    a.Top = b.Top + b.Height + 6
    for it in range(5):
        if not diagram: resolve()
        items = texts + [o for o in others if o.Name not in ("页脚", "页码")]
        bottom = max(s.Top + s.Height for s in items)
        if bottom <= BOT + 1: break
        top0 = min(s.Top for s in items)
        f = max(0.6, (BOT - top0) / max(1, bottom - top0))
        for s in items:
            s.Top = top0 + (s.Top - top0) * f
        for t in texts:
            try:
                tr = t.TextFrame.TextRange
                for k in range(1, tr.Runs().Count + 1):
                    r = tr.Runs(k); r.Font.Size = max(11, round(r.Font.Size * f * 2) / 2)
            except Exception: pass
        for o in others:
            if o.Type in (13, 6, 7, 3) or o.Type == 1:
                cx = o.Left; o.LockAspectRatio = 0; o.Width = o.Width * f; o.Height = o.Height * f
        log.append((sl.SlideIndex, "shrink%.2f" % f))
    for s in texts + others:          # keep inside page horizontally
        if s.Left + s.Width > R + 5: s.Left = max(L, R - s.Width)
        if s.Left < L - 15 and s.Width < R - L: s.Left = L
    # 3a tiny decorative pictures (balls, icons) out
    for o in list(others):
        try:
            if (o.Type == 13 and o.Width < 64 and o.Height < 64 and 0.75 < o.Width / max(1, o.Height) < 1.33 and o.Name != "公式") or (o.Type == 1 and o.AutoShapeType == 9 and o.Width < 48 and o.Height < 48 and not has_text(o)):
                o.Delete(); others.remove(o)
        except Exception: pass
    # 3c text must not sit on pictures: choose side-by-side or stacked by which keeps the picture larger
    pics = [o for o in others if o.Type in (13, 6, 7, 3) and o.Width > 60 and o.Height > 40 and o.Name != "公式"]
    def cwidth(t):
        try: return min(t.Width, t.TextFrame.TextRange.BoundWidth + t.TextFrame.MarginLeft + t.TextFrame.MarginRight + 6)
        except Exception: return t.Width
    # short labels sitting on a picture edge: slide them beside the picture
    for t in texts:
        if len(txt(t)) > 12: continue
        for o in pics:
            ox, oy = ov(t, o)
            if ox > 6 and oy > 4 and (ox * oy) / max(1, t.Width * t.Height) <= 0.8:
                if o.Left - 10 - t.Width >= L - 5: t.Left = o.Left - 10 - t.Width
                else: t.Top = o.Top - t.Height - 6
    hitp = []; block = []
    for o in pics:
        for t in texts:
            ox, oy = ov(t, o)
            if len(txt(t)) > 12 and ox > 6 and oy > 4:
                if o not in hitp: hitp.append(o)
                if t not in block: block.append(t)
    if hitp:
        for t in texts:
            if len(txt(t)) > 12 and t not in block and any(ov(t, o)[0] > 0 for o in hitp): block.append(t)
        ux0 = min(o.Left for o in hitp); uy0 = min(o.Top for o in hitp)
        ux1 = max(o.Left + o.Width for o in hitp); uy1 = max(o.Top + o.Height for o in hitp)
        ow, oh = ux1 - ux0, uy1 - uy0
        bx0 = min(t.Left for t in block); by0 = min(t.Top for t in block)
        bx1 = max(t.Left + cwidth(t) for t in block); by1 = max(t.Top + t.Height for t in block)
        s_side = min(1.3, (R - bx1 - 16) / ow, (BOT - TOP) / oh)
        s_below = min(1.3, (BOT - by1 - 8) / oh, (R - L) / ow)
        def place(sc, nx, ny):
            for o in hitp:
                o.LockAspectRatio = -1
                rx, ry = o.Left - ux0, o.Top - uy0
                o.Width = o.Width * sc; o.Left = nx + rx * sc; o.Top = ny + ry * sc
        if s_side >= s_below and s_side >= 0.45:
            for t in block: t.Width = max(60, min(t.Width, bx1 - t.Left))
            place(s_side, bx1 + 16 + max(0, (R - bx1 - 16 - ow * s_side) / 2), max(TOP, min(BOT - oh * s_side, (by0 + by1) / 2 - oh * s_side / 2)))
            log.append((sl.SlideIndex, "pic-side"))
        elif s_below >= 0.45:
            place(s_below, L + (R - L - ow * s_below) / 2, by1 + 8)
            log.append((sl.SlideIndex, "pic-below"))
        else:
            sc = min((R - L) * 0.42 / ow, (BOT - TOP) / oh, 1.0)
            place(sc, R - ow * sc, max(TOP, (by0 + by1) / 2 - oh * sc / 2))
            for t in block:
                t.Width = max(120, R - ow * sc - 16 - t.Left); t.TextFrame.WordWrap = True; t.TextFrame.AutoSize = 1
            log.append((sl.SlideIndex, "pic-squeeze"))
    # orphan short label that ended up under its body: lift it to the top as a heading
    if not diagram and 2 <= len(texts) <= 3:
        last = max(texts, key=lambda t: t.Top)
        rest = [t for t in texts if t is not last]
        if len(txt(last)) <= 8 and all(t.Top < last.Top for t in rest):
            top0 = min(t.Top for t in rest)
            last.Top = top0; last.Left = min(t.Left for t in rest)
            for t in rest: t.Top = t.Top + last.Height + 6
            resolve()
    # 3b grow small text to use the page (real measurement, revert when it collides)
    def fsizes(ts):
        out = []
        for t in ts:
            tr = t.TextFrame.TextRange
            for k in range(1, tr.Runs().Count + 1): out.append(tr.Runs(k).Font.Size)
        return out
    def scale(ts, f):
        for t in ts:
            tr = t.TextFrame.TextRange
            for k in range(1, tr.Runs().Count + 1):
                r = tr.Runs(k); r.Font.Size = round(r.Font.Size * f * 2) / 2
    def clash():
        n = 0.0
        for t in texts:
            for o in others:
                if o.Type in (13, 6, 7, 3) or (o.Type == 1 and o.Width > 30 and o.Height > 30):
                    ox, oy = ov(t, o)
                    if ox > 0 and oy > 0: n += ox * oy
        return n
    try:
        for t in texts:
            tr = t.TextFrame.TextRange
            for k in range(1, tr.Runs().Count + 1):
                if tr.Runs(k).Font.Size < 12: tr.Runs(k).Font.Size = 12
        base = clash()
        for it in range(8):
            szs = fsizes(texts)
            if not szs or max(szs) >= 28 or sum(szs) / len(szs) >= 22: break
            snap = [(t.Top, t.Height) for t in texts]
            hs = [t.Height for t in texts]
            scale(texts, 1.12)
            for t in texts: t.TextFrame.AutoSize = 1
            if not diagram: resolve()
            bottom = max(t.Top + t.Height for t in texts)
            wrapped = any(len(txt(t)) <= 20 and t.Height > h0 * 1.4 for t, h0 in zip(texts, hs))
            if wrapped or bottom > BOT - 4 or clash() > base + 1 or any(t.Left + t.Width > R + 6 for t in texts):
                scale(texts, 1 / 1.12)
                for t, (tp, h) in zip(texts, snap): t.TextFrame.AutoSize = 1; t.Top = tp
                break
            log.append((sl.SlideIndex, "grow"))
    except Exception as e:
        print("grow skip", sl.SlideIndex, e)
    # 4 card follows real text
    if card is not None and any(o.Type in (13, 6, 7, 3) and o.Width > 60 and o.Height > 40 for o in others):
        card.Delete(); card = None
    if card is not None:
        inside = [t for t in texts if ov(card, t)[0] > 5 and ov(card, t)[1] > 2]
        if inside:
            x0 = min(t.Left for t in inside) - 12; y0 = min(t.Top for t in inside) - 10
            x1 = max(t.Left + t.Width for t in inside) + 12; y1 = max(t.Top + t.Height for t in inside) + 10
            card.Left = max(30, x0); card.Top = max(100, y0); card.Width = min(930, x1) - card.Left
            card.Height = min(BOT + 8, y1) - card.Top
            for o in others:   # card must not cover pictures
                if o.Type in (13, 6, 7) and ov(card, o)[0] > 20 and ov(card, o)[1] > 20 and o.Left > card.Left + 30:
                    card.Width = max(80, o.Left - 12 - card.Left)
        else:
            card.Delete()
for sl in pres.Slides:   # titles must fit their box
    for s in sl.Shapes:
        if is_title(s) and has_text(s):
            tr = s.TextFrame.TextRange
            while tr.BoundWidth > s.Width - 10 and tr.Font.Size > 14:
                tr.Font.Size = tr.Font.Size - 1
            if tr.BoundWidth > s.Width - 10 and s.Width < 860: s.Width = min(870, tr.BoundWidth + 40)
pres.SaveAs(out); pres.Close()
print("fit done", len(log))
