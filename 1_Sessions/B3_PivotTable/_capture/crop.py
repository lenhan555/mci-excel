# crop.py <full.png> <out.png> x0 y0 x1 y1  (toa do pixel man hinh, x1/y1 khong tinh)
import sys
from PIL import Image
src, dst, x0, y0, x1, y1 = sys.argv[1], sys.argv[2], *map(int, sys.argv[3:7])
im = Image.open(src).convert("RGB").crop((x0, y0, x1, y1))
im.save(dst, optimize=True)
print(dst, im.size)
