for n, lab, ow, oh in ((35, "Text Box 15", 4.59, 4.76), (36, None, 4.42, 3.22)):
    g = [s for s in S(n).shapes if s.shape_type == 6][0]
    s_ = min(4.4 / oh, 6.0 / ow, 1.55)
    setg(g, 0.9, 1.7, ow * s_, oh * s_)
    for t in S(n).shapes:
        if t.has_text_frame and t.name in ("Text Box 15", "Text Box 4") and t.shape_type == 17 and text_of(t._element).startswith(("流态化", "渗滤")):
            setg(t, 0.9, 1.7 + oh * s_ + 0.1, ow * s_, 0.5)
    for t in S(n).shapes:
        if t.has_text_frame and t.shape_type == 17 and t.left / IN > 5.5:
            setg(t, 0.9 + ow * s_ + 0.5, 1.9, 12.4 - (0.9 + ow * s_ + 0.5), 4.8)
for n in (35, 36):
    for t in list(S(n).shapes):
        if t.name == "文字卡片": t._element.getparent().remove(t._element)
