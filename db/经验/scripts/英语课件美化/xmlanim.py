"""Compare animation effects at XML level (what any player will read), slide by slide."""
import sys, re, zipfile

def effects(path):
    z = zipfile.ZipFile(path)
    out = {}
    for n in z.namelist():
        m = re.match(r'ppt/slides/slide(\d+)\.xml$', n)
        if not m:
            continue
        x = z.read(n).decode('utf-8')
        i = x.find('<p:timing')
        t = x[i:] if i >= 0 else ''
        effs = []
        for em in re.finditer(r'<p:cTn id="\d+"([^>]*presetClass="(\w+)"[^>]*)>(.*?)</p:cTn>\s*</p:par>', t, re.S):
            attrs, cls, body = em.group(1), em.group(2), em.group(3)
            node = re.search(r'nodeType="(\w+)"', attrs)
            sp = re.search(r'<p:spTgt spid="(\d+)"', body)
            effs.append((cls, node.group(1) if node else '?', sp.group(1) if sp else '?'))
        trig = re.findall(r'<p:cond evt="onClick"[^>]*>\s*<p:tgtEl>\s*<p:spTgt spid="(\d+)"', t)
        out[int(m.group(1))] = (effs, trig)
    return out

a, b = effects(sys.argv[1]), effects(sys.argv[2])
for k in sorted(a):
    ea, eb = a[k], b.get(k, ([], []))
    if ea != eb:
        print(f"S{k}: original {len(ea[0])} effects / triggers {ea[1]}  ->  new {len(eb[0])} / {eb[1]}")
        print("   orig:", ea[0])
        print("   new :", eb[0])
