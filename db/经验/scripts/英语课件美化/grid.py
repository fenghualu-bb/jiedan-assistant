import sys, glob, os
from PIL import Image, ImageDraw, ImageFont
src, prefix = sys.argv[1], sys.argv[2]
per = int(sys.argv[3]) if len(sys.argv)>3 else 4
files = sorted(glob.glob(os.path.join(src, "s*.png")))
cols = 2
tw, th = 800, 450
for gi in range(0, len(files), per):
    chunk = files[gi:gi+per]
    rows = (len(chunk)+cols-1)//cols
    im = Image.new("RGB", (cols*tw+(cols+1)*10, rows*(th+30)+10), "white")
    d = ImageDraw.Draw(im)
    for k, f in enumerate(chunk):
        r, c = divmod(k, cols)
        t = Image.open(f).convert("RGB").resize((tw, th))
        x, y = 10+c*(tw+10), 10+r*(th+30)
        im.paste(t, (x, y+22)); d.rectangle([x, y+22, x+tw-1, y+22+th-1], outline="gray")
        d.text((x, y+4), os.path.basename(f), fill="black", font=ImageFont.truetype("arial.ttf", 18))
    out = f"{prefix}-{gi//per+1}.jpg"; im.save(out, quality=85); print(os.path.abspath(out))
