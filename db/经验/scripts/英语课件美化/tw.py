import sys
from PIL import ImageFont
R = ImageFont.truetype("C:/Windows/Fonts/comic.ttf", 1000)
B = ImageFont.truetype("C:/Windows/Fonts/comicbd.ttf", 1000)
def w(s, pt, bold=False): return (B if bold else R).getlength(s)*pt/1000
for line in sys.stdin:
    line=line.rstrip('\n')
    if not line: continue
    pt, bold, s = line.split('|',2)
    print(f"{w(s,float(pt),bold=='b'):7.1f}  {pt}{bold} {s}")
