import sys, win32com.client
app = win32com.client.Dispatch("PowerPoint.Application")
pres = app.Presentations.Open(sys.argv[1], True, False, False)
m = pres.SlideMaster
print("MASTER shapes:", [(s.Name, s.Type, round(s.Left), round(s.Top), round(s.Width), round(s.Height)) for s in m.Shapes])
for i in range(1, m.CustomLayouts.Count+1):
    L = m.CustomLayouts(i)
    print("LAYOUT", i, L.Name, "followbg", L.FollowMasterBackground, [(s.Name, s.Type) for s in L.Shapes])
for d in range(1, pres.Designs.Count+1):
    print("DESIGN", d, pres.Designs(d).Name)
for s in pres.Slides:
    out=[]
    seq = s.TimeLine.MainSequence
    for k in range(1, seq.Count+1):
        e = seq(k)
        try: nm = e.Shape.Name
        except Exception: nm = '?'
        out.append(f"{nm}:{e.EffectType}{'X' if e.Exit else ''}")
    inter=[]
    for q in range(1, s.TimeLine.InteractiveSequences.Count+1):
        sq = s.TimeLine.InteractiveSequences(q)
        for k in range(1, sq.Count+1):
            e = sq(k)
            try: trg = e.Timing.TriggerShape.Name
            except Exception: trg='?'
            inter.append(f"[{trg}]->{e.Shape.Name}:{e.EffectType}{'X' if e.Exit else ''}")
    print(f"S{s.SlideIndex} bgfollow={s.FollowMasterBackground} design={s.Design.Name} layout={s.CustomLayout.Name} main={out} inter={inter}")
pres.Close()
