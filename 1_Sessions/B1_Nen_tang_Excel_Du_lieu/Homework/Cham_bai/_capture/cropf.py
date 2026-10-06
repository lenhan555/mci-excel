# cropf.py <name> <x0> <y0> <x1> <y1> [dst]  -- toa do theo khung 1389x868 cua screenshot; chon ban it cam nhat
import sys, glob, os
from PIL import Image
K = 1920 / 1389
name = sys.argv[1]; x0, y0, x1, y1 = [round(float(v) * K) for v in sys.argv[2:6]]
dst = sys.argv[6] if len(sys.argv) > 6 else f"shots/{name}.png"
best = None
for p in sorted(glob.glob(f"out/{name}_*.png")):
    im = Image.open(p).convert("RGB").crop((x0, y0, x1, y1))
    w, h = im.size; px = im.load()
    edge = [px[x, 0] for x in range(w)] + [px[x, h - 1] for x in range(w)] + [px[0, y] for y in range(h)] + [px[w - 1, y] for y in range(h)]
    orange = sum(1 for r, g, b in edge if r > 200 and 90 < g < 180 and b < 110)
    if best is None or orange < best[0]: best = (orange, p, im)
best[2].save(dst, optimize=True)
print(dst, best[2].size, "from", os.path.basename(best[1]), "orange_edge", best[0])
