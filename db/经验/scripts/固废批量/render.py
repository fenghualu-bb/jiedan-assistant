import sys, os, win32com.client, pythoncom
src = os.path.abspath(sys.argv[1]); out = os.path.abspath(sys.argv[2]); w = int(sys.argv[3]) if len(sys.argv)>3 else 960
os.makedirs(out, exist_ok=True)
for f in os.listdir(out):
    if f.lower().endswith('.png'): os.remove(os.path.join(out,f))
pythoncom.CoInitialize()
app = win32com.client.DispatchEx("PowerPoint.Application")
pres = app.Presentations.Open(src, ReadOnly=True, Untitled=False, WithWindow=False)
ratio = pres.PageSetup.SlideHeight / pres.PageSetup.SlideWidth
for i, s in enumerate(pres.Slides, 1):
    s.Export(os.path.join(out, f"s{i:02d}.png"), "PNG", w, int(w*ratio))
print("slides", pres.Slides.Count, pres.PageSetup.SlideWidth, pres.PageSetup.SlideHeight)
pres.Close()
