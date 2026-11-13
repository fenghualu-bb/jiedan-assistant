"""Export each overlay slide twice: as-is, and with its masks / bubbles hidden (temp copy, never saved)."""
import sys, os, win32com.client
win32com.client.gencache.EnsureModule('{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}', 0, 2, 8)
app = win32com.client.gencache.EnsureDispatch("PowerPoint.Application")
src, outdir = os.path.abspath(sys.argv[1]), os.path.abspath(sys.argv[2])
os.makedirs(outdir, exist_ok=True)
HIDE = {8: (13, 8, 9), 9: (11, 12, 3), 11: (12,), 13: (12,), 15: (12, 2), 16: (12, 3),
        20: (269, 271, 275, 276, 277, 278, 279), 22: (269, 271, 273, 274, 275, 276, 277, 278)}
pres = app.Presentations.Open(src, True, False, False)
for si, ids in HIDE.items():
    sl = pres.Slides(si)
    sl.Export(os.path.join(outdir, f"s{si:02d}_a.png"), "PNG", 1280, 720)
    for s in sl.Shapes:
        if s.Id in ids:
            print(f"S{si} hide [{s.Id}] {s.Name} ({s.Left:.0f},{s.Top:.0f},{s.Width:.0f},{s.Height:.0f}) fill={s.Fill.Visible} line={s.Line.Visible}")
            s.Visible = 0
    sl.Export(os.path.join(outdir, f"s{si:02d}_b.png"), "PNG", 1280, 720)
pres.Close()
