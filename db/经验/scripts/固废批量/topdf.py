import sys, os, win32com.client, pythoncom
pythoncom.CoInitialize(); app = win32com.client.DispatchEx("PowerPoint.Application")
for src in sys.argv[1:]:
    p = app.Presentations.Open(os.path.abspath(src), ReadOnly=True, WithWindow=False)
    p.SaveAs(os.path.abspath(src)[:-5] + ".pdf", 32); p.Close(); print("pdf", src)
