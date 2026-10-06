import sys,glob
from PIL import Image
files=sys.argv[2:]; out=sys.argv[1]
ims=[Image.open(f).convert('RGB') for f in files]
W=max(i.width for i in ims); H=sum(i.height+8 for i in ims)
m=Image.new('RGB',(W,H),(200,200,200)); y=0
for i in ims: m.paste(i,(0,y)); y+=i.height+8
if W>1100: m=m.resize((1100,int(H*1100/W)))
m.save(out)
