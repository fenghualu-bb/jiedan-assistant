import sys, os, win32com.client, pythoncom
pythoncom.CoInitialize()
app = win32com.client.DispatchEx("PowerPoint.Application")
pres = app.Presentations.Open(os.path.abspath(sys.argv[1]), ReadOnly=True, WithWindow=False)
for n in map(int, sys.argv[2:]):
    sl = pres.Slides(n); print("== slide", n)
    for s in sl.Shapes:
        info = ""
        try:
            if s.HasTextFrame and s.TextFrame.HasText:
                tr = s.TextFrame.TextRange
                info = "sizes=%s auto=%s text=%r" % (sorted(set(tr.Runs(k).Font.Size for k in range(1, tr.Runs().Count + 1))), s.TextFrame.AutoSize, tr.Text[:25])
        except Exception as e: info = "ERR %s" % e
        print("  ", s.Type, s.Name, round(s.Left), round(s.Top), round(s.Width), round(s.Height), info)
pres.Close()
