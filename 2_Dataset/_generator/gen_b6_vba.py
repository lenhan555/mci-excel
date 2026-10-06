# -*- coding: utf-8 -*-
"""Sinh vba/build_b6_dashboard.bas (ASCII thuần, tiếng Việt qua ChrW) từ mẫu VBA + số liệu build_b6_meta.json.
Chạy: python3 gen_b6_vba.py [tên_module]  → ghi vba/<tên_module>.bas (mặc định B6Build); mẫu: b6_template.bas.txt
Cách dùng .bas: xem 2_Dataset/README.md (mục Buổi 6, Practice)."""
import json, re, sys, os

HERE = os.path.dirname(os.path.abspath(__file__))
META = json.load(open(os.path.join(HERE, "build_b6_meta.json"), encoding="utf-8"))
MODNAME = sys.argv[1] if len(sys.argv) > 1 else "B6Build"

def lit(m):
    s = m.group(1)
    if all(ord(c) < 128 for c in s):
        return m.group(0)
    out, buf = [], ""
    for c in s:
        if ord(c) < 128:
            buf += c
        else:
            if buf:
                out.append('"' + buf + '"'); buf = ""
            out.append("ChrW(%d)" % ord(c))
    if buf:
        out.append('"' + buf + '"')
    return " & ".join(out)

def to_ascii(src):
    # chuyển mọi chuỗi "..." chứa ký tự không ASCII thành chuỗi nối ChrW (bỏ qua dòng chú thích '...)
    lines = []
    for ln in src.split("\n"):
        code, cmt = ln, ""
        # tách chú thích ' ngoài chuỗi
        inq = False
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

def vba_num(x):
    return repr(float(x))

# ---- dữ liệu cho Kiem_tra_B6 ----
card = META["card"]
KT = []   # (nhãn, công thức, đáp án, dung sai, định dạng)
def mem(m): return '"[Calendar].[Thang].&[%d]"' % m
NAM = '"[Calendar].[Nam].&[2025]"'
def cv(measure, *members):
    return '=CUBEVALUE("ThisWorkbookDataModel","[Measures].[%s]",%s)' % (measure, ",".join(members))
names = [("rev", "Total Revenue", "Doanh thu", 0.5, "#,##0"), ("orders", "Total Orders", u"Số đơn", 0.5, "#,##0"),
         ("aov", "Avg Order Value", "AOV", 0.5, "#,##0"), ("gm", "Gross Margin %", "GM%", 0.0005, "0.0%"),
         ("yoy", "Revenue YoY %", "YoY", 0.0005, "0.0%"), ("cust", "Active Customers", u"Khách hoạt động", 0.5, "#,##0")]
for m in (2, 3, 4):
    c = card[str(m)]
    for k, meas, lab, tol, fmt in names:
        KT.append((u"%s tháng %d/2025" % (lab, m), cv(meas, NAM, mem(m)), c[k], tol, fmt))
    KT.append((u"Target tháng %d/2025" % m, cv("Total Target", NAM, mem(m)), c["target"], 0.5, "#,##0"))
KT.append((u"Waterfall tháng 3: tổng Variance (thực tế − target)", cv("Variance", NAM, mem(3)), card["3"]["rev"] - card["3"]["target"], 0.5, "#,##0"))
for p in META["top5_t4"]:
    KT.append((u"Top 5 tháng 4: %s (%s)" % (p["name"], p["code"]), cv("Total Revenue", NAM, mem(4), '"[Product].[Product_Code].&[%s]"' % p["code"]), p["rev"], 0.5, "#,##0"))
for m in (1, 2, 3, 4):
    KT.append((u"Combo: doanh thu tháng %d" % m, cv("Total Revenue", NAM, mem(m)), META["actual_month"][str(m)], 0.5, "#,##0"))
    KT.append((u"Combo: target tháng %d" % m, cv("Total Target", NAM, mem(m)), META["target_month"][str(m)], 0.5, "#,##0"))

kt_lines = []
for i, (lab, f, exp, tol, fmt) in enumerate(KT):
    r = 5 + i
    kt_lines.append('    ws.Cells(%d, 1).Value = "%s"' % (r, lab.replace('"', '""')))
    kt_lines.append('    ws.Cells(%d, 3).Formula = "%s"' % (r, f.replace('"', '""')))
    kt_lines.append('    ws.Cells(%d, 4).Value = %s' % (r, vba_num(exp)))
    kt_lines.append('    ws.Cells(%d, 5).Formula = "=IF(ISNUMBER(C%d),IF(ABS(C%d-D%d)<=%s,""✔"",""✘""),""✘"")"' % (r, r, r, r, vba_num(tol)))
    kt_lines.append('    ws.Range("C%d:D%d").NumberFormat = "%s"' % (r, r, fmt))
KT_N = len(KT)
KT_LAST = 4 + KT_N

TEMPLATE = open(os.path.join(HERE, "b6_template.bas.txt"), encoding="utf-8").read()
TEMPLATE = TEMPLATE.replace("@@MOD@@", MODNAME).replace("@@KT_LINES@@", "\n".join(kt_lines)).replace("@@KT_LAST@@", str(KT_LAST))
out = to_ascii(TEMPLATE).replace("\r\n", "\n").replace("\n", "\r\n")
assert all(ord(c) < 128 for c in out), "còn ký tự non-ASCII"
# giới hạn 1023 ký tự/dòng
for n, ln in enumerate(out.split("\r\n"), 1):
    assert len(ln) < 1000, ("dòng quá dài", n, len(ln))
os.makedirs(os.path.join(HERE, "vba"), exist_ok=True)
path = os.path.join(HERE, "vba", MODNAME + ".bas")
open(path, "w", encoding="ascii", newline="").write(out)
print("ok", path, len(out.split("\r\n")), "dòng;", KT_N, "kiểm tra")
