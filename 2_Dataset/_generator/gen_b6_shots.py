# -*- coding: utf-8 -*-
"""Sinh vba/build_b6_homework.bas (ASCII thuần, tiếng Việt qua ChrW) từ b6shots_template.bas.txt.
Chạy: python3 gen_b6_hw_vba.py [tên_module]  (mặc định B6HW). Cách dùng: xem đầu file .bas."""
import os, re, sys
HERE = os.path.dirname(os.path.abspath(__file__))
MODNAME = sys.argv[1] if len(sys.argv) > 1 else "B6Shots"

def lit(m):
    s = m.group(1)
    if all(ord(c) < 128 for c in s):
        return m.group(0)
    return 'U("' + "".join(c if ord(c) < 128 else "\\u%04X" % ord(c) for c in s) + '")'

def to_ascii(src):
    lines = []
    for ln in src.split("\n"):
        code, cmt, inq = ln, "", False
        for i, ch in enumerate(ln):
            if ch == '"': inq = not inq
            if ch == "'" and not inq:
                code, cmt = ln[:i], ln[i:]
                break
        code = re.sub(r'"((?:[^"]|"")*)"', lit, code)
        if cmt and not all(ord(c) < 128 for c in cmt):
            cmt = "'" + "".join(c if ord(c) < 128 else "?" for c in cmt[1:])
        lines.append(code + cmt)
    return "\n".join(lines)

tpl = open(os.path.join(HERE, "b6shots_template.bas.txt"), encoding="utf-8").read().replace("@@MOD@@", MODNAME)
out = to_ascii(tpl).replace("\r\n", "\n").replace("\n", "\r\n")
assert all(ord(c) < 128 for c in out)
for n, ln in enumerate(out.split("\r\n"), 1):
    assert len(ln) < 1000, ("dòng quá dài", n, len(ln))
os.makedirs(os.path.join(HERE, "vba"), exist_ok=True)
path = os.path.join(HERE, "vba", MODNAME + ".bas")
open(path, "w", encoding="ascii", newline="").write(out)
print("ok", path, len(out.split("\r\n")), "dòng")
