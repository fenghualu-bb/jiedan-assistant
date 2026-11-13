# Force every text font slot (latin/ea/cs) to YaHei in slides, layouts, masters, charts, themes.
import sys, zipfile, re, shutil, os
src = sys.argv[1]; tmp = src + ".tmp"
pat = re.compile(r'<a:(latin|ea|cs) typeface="[^"]*"((?: [a-zA-Z]+="[^"]*")*)\s*/>')
def fix(x):
    return pat.sub(lambda m: '<a:%s typeface="微软雅黑"/>' % m.group(1), x)
with zipfile.ZipFile(src) as zi, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
    n = 0
    for it in zi.infolist():
        data = zi.read(it.filename)
        if re.match(r'ppt/(slides|slideLayouts|slideMasters|charts|theme)/[^/]+\.xml$', it.filename):
            x = data.decode("utf8"); y = fix(x)
            if it.filename.startswith("ppt/charts/chart"):   # give chart text an explicit face
                y = re.sub(r'<a:defRPr([^>]*)/>', r'<a:defRPr\1><a:latin typeface="微软雅黑"/><a:ea typeface="微软雅黑"/></a:defRPr>', y)
            if "/theme/" in it.filename:
                y = re.sub(r'<a:(hlink|folHlink)>.*?</a:(?:hlink|folHlink)>', lambda m: '<a:%s><a:srgbClr val="1A5F9A"/></a:%s>' % (m.group(1), m.group(1)), y, flags=re.S)
            n += x != y; data = y.encode("utf8")
        zo.writestr(it, data)
shutil.move(tmp, src); print("fontfix parts changed:", n)
