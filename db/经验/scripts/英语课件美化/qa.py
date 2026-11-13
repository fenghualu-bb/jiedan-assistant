"""QA: min font size, text overflow (PowerPoint's own text metrics), off-slide shapes, fonts used."""
import sys, os, win32com.client
win32com.client.gencache.EnsureModule('{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}', 0, 2, 8)
app = win32com.client.gencache.EnsureDispatch("PowerPoint.Application")
if "Microsoft Office" not in app.Path:
    raise SystemExit(f"ABORT: PowerPoint automation is being served by {app.Path} (WPS?). Close WPS and re-run.")
_mine = {os.path.abspath(a).lower() for a in sys.argv[1:3] if a.lower().endswith('.pptx')}
for _p in list(app.Presentations):
    if _p.FullName.lower() in _mine:
        _p.Close()
pres = app.Presentations.Open(os.path.abspath(sys.argv[1]), True, False, False)
W, H = pres.PageSetup.SlideWidth, pres.PageSetup.SlideHeight
fonts = {}
issues = []
minsize = 999
for sl in pres.Slides:
    i = sl.SlideIndex
    shapes = []
    for s in sl.Shapes:
        shapes.append(s)
        if s.Type == 6:
            shapes.extend(list(s.GroupItems))
    for s in shapes:
        if s.Type == 6:
            continue
        if s.Type != 16 and i != 10 and s.Visible and (s.Left < 18 or s.Top < 16 or s.Left + s.Width > 942 or s.Top + s.Height > 524):
            issues.append(f"S{i} near/over frame [{s.Id}] {s.Name} t{s.Type} ({s.Left:.0f},{s.Top:.0f},{s.Width:.0f},{s.Height:.0f})")
        if not s.HasTextFrame:
            continue
        tr = s.TextFrame.TextRange
        if tr.Length == 0 or not tr.Text.strip():
            continue
        for k in range(1, tr.Runs().Count + 1):
            r = tr.Runs(k, 1)
            if not r.Text.strip():
                continue
            sz = r.Font.Size
            minsize = min(minsize, sz)
            if sz < 24:
                issues.append(f"S{i} small font {sz} [{s.Id}] {r.Text[:30]!r}")
            fonts[(r.Font.Name, r.Font.NameFarEast)] = fonts.get((r.Font.Name, r.Font.NameFarEast), 0) + 1
            if r.Font.Name != "Comic Sans MS":
                issues.append(f"S{i} latin font {r.Font.Name} [{s.Id}] {r.Text[:30]!r}")
        tf2 = s.TextFrame2
        bw, bh = tf2.TextRange.BoundWidth, tf2.TextRange.BoundHeight
        aw = s.Width - tf2.MarginLeft - tf2.MarginRight
        ah = s.Height - tf2.MarginTop - tf2.MarginBottom
        if tf2.AutoSize == 0 and abs(s.Rotation) < 0.1:
            if bw > aw + 2 or bh > ah + 3:
                issues.append(f"S{i} overflow [{s.Id}] {s.Name} text {bw:.0f}x{bh:.0f} > box {aw:.0f}x{ah:.0f} {tr.Text[:40]!r}")
print("min font size:", minsize)
print("fonts (latin, eastasian):", fonts)
print("\n".join(issues) if issues else "no issues")
pres.Close()
