import sys, os
from PIL import Image, ImageDraw
d=sys.argv[1]; pref=sys.argv[2]; cols=int(sys.argv[3]) if len(sys.argv)>3 else 3
fs=sorted(f for f in os.listdir(d) if f.endswith('.png'))
ims=[Image.open(os.path.join(d,f)) for f in fs]
w,h=ims[0].size; per=cols*4
for k in range(0,len(ims),per):
    chunk=ims[k:k+per]; rows=(len(chunk)+cols-1)//cols
    g=Image.new('RGB',(cols*(w+10),rows*(h+30)),'white'); dr=ImageDraw.Draw(g)
    for j,im in enumerate(chunk):
        x=(j%cols)*(w+10); y=(j//cols)*(h+30)
        g.paste(im,(x,y+25)); dr.text((x+5,y+5),fs[k+j],fill='red')
    p=f"{pref}-{k//per+1}.jpg"; g.save(p,quality=85); print(os.path.abspath(p))
