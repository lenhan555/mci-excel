#!/bin/bash
# usage: snap2.sh OUTNAME x1 y1 x2 y2 (toa do khong gian anh chup man hinh 1389 rong); chup rieng cua so Excel
R="$HOME/mnt/MCI Excel"
rm -f "$R/2_Dataset/_generator/vba/_snap.done"
echo "raw_$1" > "$R/2_Dataset/_generator/vba/_snap.trigger"
for i in $(seq 1 30); do [ -f "$R/2_Dataset/_generator/vba/_snap.done" ] && break; sleep 0.5; done
sleep 1
python3 - "$@" <<'P'
import sys
from PIL import Image
n,x1,y1,x2,y2=sys.argv[1],*map(float,sys.argv[2:6])
R="/sessions/rcw-01bprbfzzqkp2qyshcc2byk9/mnt/MCI Excel/"
d=open(R+"2_Dataset/_generator/vba/_snap.done").read().split()
sw,L,T=int(d[1]),int(d[2]),int(d[3])
im=Image.open(R+"Materials/B6/Images/_raw/raw_"+n)
k=sw/1389
box=(int(x1*k-L),int(y1*k-T),int(x2*k-L),int(y2*k-T))
im.crop(box).save(R+"Materials/B6/Images/"+n)
print(n,im.size,box)
P
