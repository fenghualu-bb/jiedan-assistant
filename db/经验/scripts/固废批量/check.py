import sys
from pptx import Presentation
E=914400
p=Presentation(sys.argv[1]); bad=[]
def box(sh): return sh.left/E, sh.top/E, sh.width/E, sh.height/E
for i,s in enumerate(p.slides,1):
    tx=[sh for sh in s.shapes if sh.has_text_frame and sh.text_frame.text.strip() and sh.shape_type!=13 and not sh.is_placeholder and sh.name not in ('页脚','页码')]
    issues=[]
    for a in range(len(tx)):
        for b in range(a+1,len(tx)):
            A,B=box(tx[a]),box(tx[b])
            ox=min(A[0]+A[2],B[0]+B[2])-max(A[0],B[0]); oy=min(A[1]+A[3],B[1]+B[3])-max(A[1],B[1])
            if ox>0.3 and oy>0.25 and ox*oy>0.15*min(A[2]*A[3],B[2]*B[3]): issues.append(('overlap',tx[a].name,tx[b].name))
    for sh in s.shapes:
        if sh.name in ('页脚','页码'): continue
        x,y,w,h=box(sh)
        if x<-0.05 or y<-0.05 or x+w>13.4 or y+h>7.55: issues.append(('offpage',sh.name,round(x+w,1),round(y+h,1)))
    if issues: bad.append((i,issues[:3]))
print(sys.argv[1],len(bad),'flagged'); [print(' ',b) for b in bad]
