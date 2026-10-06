"""Dựng deck Buổi 6: Slide_B6.pptx từ Slide_Script_B6.md, theo style của Slide_B1.pptx (cùng khung với build_slides_b4.py).

- Master, layout, theme và slide bìa lấy từ Slide_B1.pptx (không sửa file B1; chỉ đọc làm khuôn).
- Chữ trên slide khai báo trong script này, khớp phần "Trên slide" của kịch bản.
- Speaker notes lấy NGUYÊN VĂN từ kịch bản (Lời giảng, Ghi chú giảng viên), bỏ định dạng markdown,
  nên sửa lời giảng thì sửa Slide_Script_B6.md rồi chạy lại script này.
- Quiz: shape đáp án tên "Answer …", lớp mờ "Dim …", hiện khi bấm (build_student_slides.py sẽ xóa ở bản _HocVien).
- Ảnh chụp thật lấy từ Materials/B6/Images (tên S<số slide>_…png, danh sách: _DANH_SACH_ANH.md); chưa có thì để khung [CHỤP].
- Số liệu khớp build_b6_meta.json (build_b6_meta.py).

    python3 _generator/build_slides_b6.py     # cần python-pptx
    python3 _generator/build_student_slides.py B6
"""
import copy, os, re, sys
for s_ in (sys.stdout, sys.stderr):
    if hasattr(s_, "reconfigure"): s_.reconfigure(encoding="utf-8", errors="replace")
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.dml import MSO_LINE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from lxml import etree

ROOT = os.environ.get("MCI_ROOT") or os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
TEMPLATE = os.path.join(ROOT, "1_Sessions", "B1_Nen_tang_Excel_Du_lieu", "Slide_B1.pptx")
B6 = os.path.join(ROOT, "1_Sessions", "B6_Dashboard")
SCRIPT = os.path.join(B6, "Slide_Script_B6.md")
OUT = os.path.join(B6, "Slide_B6.pptx")
IMG = os.path.join(ROOT, "Materials", "B6", "Images")        # ảnh chụp thật: S<số slide>_<chủ đề>.png
MISSING = []

NAVY, INK, INK2, MUTED = "203864", "1B2230", "585858", "6B6F76"
RED, GOLD, ALT, HAIR, WHITE = "C00000", "FFC000", "F3F5F9", "D9DEE8", "FFFFFF"
GOOD, GOOD_BG, BAD_TXT, BAD_BG, ACCENT = "2E7D32", "E8F5E9", "9C0006", "FFC7CE", "3398DA"
L, R, W = 1.67, 12.74, 11.07          # lề trái, lề phải, bề rộng vùng nội dung (in)

# ---------------------------------------------------------------------
# Khuôn: mở B1, giữ slide bìa, xóa các slide còn lại
# ---------------------------------------------------------------------
prs = Presentation(TEMPLATE)
sldIdLst = prs.slides._sldIdLst
for sldId in list(sldIdLst)[1:]:
    prs.part.drop_rel(sldId.rId); sldIdLst.remove(sldId)
LAYOUT = {l.name: l for l in prs.slide_layouts}

def rgb(h): return RGBColor.from_string(h)

def _runs(p, parts, size, color, bold):
    for part in parts:
        if isinstance(part, str): part = (part, {})
        text, f = part
        r = p.add_run(); r.text = text
        r.font.size = Pt(f.get("size", size)); r.font.bold = f.get("b", bold)
        r.font.color.rgb = rgb(f.get("color", color)); r.font.name = "Arial"
        if f.get("i"): r.font.italic = True

def tb(s, x, y, w, h, paras, size=18, color=INK, bold=False, anchor="t", align=None, name="TextBox", bullet=False, space=6):
    """paras: str | list[para]; para: str | list[run]; run: str | (text, {b, color, size, i})."""
    sh = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h)); sh.name = name
    tf = sh.text_frame; tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = {"t": MSO_ANCHOR.TOP, "ctr": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}[anchor]
    if isinstance(paras, str): paras = [paras]
    for i, para in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        if align: p.alignment = {"l": PP_ALIGN.LEFT, "ctr": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}[align]
        if i: p.space_before = Pt(space)
        if isinstance(para, str): para = [para]
        if bullet: para = [("•  ", {"color": NAVY, "b": True})] + list(para)
        _runs(p, para, size, color, bold)
    return sh

def box(s, x, y, w, h, fill=None, line=None, lw=1, shape=MSO_SHAPE.RECTANGLE, radius=None, name="Box", alpha=None):
    sh = s.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h)); sh.name = name
    if radius is not None and shape == MSO_SHAPE.ROUNDED_RECTANGLE: sh.adjustments[0] = radius
    if fill: sh.fill.solid(); sh.fill.fore_color.rgb = rgb(fill)
    else: sh.fill.background()
    if line: sh.line.color.rgb = rgb(line); sh.line.width = Pt(lw)
    else: sh.line.fill.background()
    sh.shadow.inherit = False
    if alpha is not None:
        clr = sh.fill._xPr.find(qn("a:solidFill")).find(qn("a:srgbClr"))
        a = etree.SubElement(clr, qn("a:alpha")); a.set("val", str(alpha))
    return sh

def chip(s, x, y, w, text, fill=None, line=NAVY, color=NAVY, size=14, name="Chip", h=0.4):
    sh = box(s, x, y, w, h, fill=fill, line=line, lw=1, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.2, name=name)
    tf = sh.text_frame; tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER; _runs(p, [text], size, color, True)
    return sh

def vline(s, x, y, h):
    ln = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x), Inches(y), Inches(x), Inches(y + h))
    ln.line.color.rgb = rgb(HAIR); ln.line.width = Pt(1); return ln

def content(title, rail):
    s = prs.slides.add_slide(LAYOUT["Title Only"])
    s.shapes.title.text = title
    tb(s, 0, 1.5, 1.07, 5.25, rail, size=14, color=WHITE, bold=True, anchor="ctr", align="ctr", name="Rail label")
    s.shapes[-1].text_frame._bodyPr.set("vert", "vert270")
    return s

def section(title, body=None):
    s = prs.slides.add_slide(LAYOUT["Section Header"])
    s.shapes.title.text = title
    for ph in list(s.placeholders):
        if ph.placeholder_format.idx == 1:
            if body: ph.text = body
            else: ph._element.getparent().remove(ph._element)
    return s

def tip(s, x, y, w, h, text, pro=False):
    box(s, x, y, w, h, fill=ALT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06, name=("PRO TIP" if pro else "TIP") + " callout")
    if pro:
        chip(s, x + 0.12, y + 0.12, 1.1, "PRO TIP", fill=GOLD, line=None, color=INK, size=12, name="PRO TIP chip", h=0.35)
        tb(s, x + 1.32, y + 0.1, w - 1.44, h - 0.2, text, size=16, name="PRO TIP text")
    else:
        tb(s, x + 0.12, y + 0.12, 0.6, 0.35, "TIP", size=14, color=NAVY, bold=True, anchor="ctr")
        tb(s, x + 0.75, y + 0.1, w - 0.87, h - 0.2, text, size=16, name="TIP text")

def tool(s, x, y, w, name, purpose, h_purpose=0.7, size=24):
    tb(s, x, y, w, 0.5, name, size=size, color=RED, bold=True, name="Tool name")
    tb(s, x, y + 0.5, w, h_purpose, purpose, size=16, color=INK2, name="Tool purpose")

def table(s, x, y, colw, rows, size=16, head_size=None, rowh=0.45, bold_first=False, name="Table",
          header=True, fills=None, colors=None, align=None):
    """rows[0] là tiêu đề (nền navy). fills/colors: {(r,c): hex}. Chữ trong ô có thể là list run."""
    nr, nc = len(rows), len(rows[0])
    gf = s.shapes.add_table(nr, nc, Inches(x), Inches(y), Inches(sum(colw)), Inches(rowh * nr)); gf.name = name
    tbl = gf.table
    tbl._tbl.tblPr.set("firstRow", "1"); tbl._tbl.tblPr.set("bandRow", "0")
    style = tbl._tbl.tblPr.find(qn("a:tableStyleId"))
    if style is not None: tbl._tbl.tblPr.remove(style)
    for j, cw in enumerate(colw): tbl.columns[j].width = Inches(cw)
    for i in range(nr):
        tbl.rows[i].height = Inches(rowh)
        for j in range(nc):
            c = tbl.cell(i, j); v = rows[i][j]
            c.margin_left = c.margin_right = Inches(0.08); c.margin_top = c.margin_bottom = Inches(0.03)
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf = c.text_frame; tf.word_wrap = True
            p = tf.paragraphs[0]
            if align and align[j]: p.alignment = {"l": PP_ALIGN.LEFT, "ctr": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}[align[j]]
            head = header and i == 0
            col = WHITE if head else (colors or {}).get((i, j), INK)
            parts = v if isinstance(v, list) else [str(v)]
            _runs(p, parts, head_size or size, col, head or (bold_first and j == 0))
            tcPr = c._tc.get_or_add_tcPr()
            for side in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
                ln = etree.SubElement(tcPr, qn(side))
                if side in ("a:lnL", "a:lnR"):
                    ln.set("w", "0"); etree.SubElement(ln, qn("a:noFill"))
                else:
                    ln.set("w", "12700"); sf = etree.SubElement(ln, qn("a:solidFill"))
                    etree.SubElement(sf, qn("a:srgbClr")).set("val", HAIR)
            f = NAVY if head else (fills or {}).get((i, j))
            if f:
                sf = etree.SubElement(tcPr, qn("a:solidFill")); etree.SubElement(sf, qn("a:srgbClr")).set("val", f)
            else:
                etree.SubElement(tcPr, qn("a:noFill"))
    return gf

def fx(s, x, y, w, formula, size=16, name="Formula bar"):
    """Thanh công thức mô phỏng Excel: ô 'fx' + công thức."""
    box(s, x, y, w, 0.45, fill=WHITE, line=HAIR, name=name)
    tb(s, x + 0.1, y, 0.45, 0.45, [[("fx", {"i": True})]], size=14, color=MUTED, anchor="ctr")
    tb(s, x + 0.6, y, w - 0.7, 0.45, formula, size=size, color=INK, anchor="ctr", name="Formula")

def stat(s, x, y, w, h, label, value, sub=None, name="Stat card", vcolor=NAVY, vsize=32):
    box(s, x, y, w, h, fill=WHITE, line=HAIR, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06, name=name)
    paras = [[(label, {"size": 14, "color": INK2})], [(value, {"size": vsize, "b": True, "color": vcolor})]]
    if sub: paras.append([(sub, {"size": 14, "color": INK})])
    tb(s, x + 0.2, y + 0.1, w - 0.4, h - 0.2, paras, anchor="ctr", space=2)

def shot(s, x, y, w, h, what, where, img, name="Image placeholder [CHỤP]"):
    """Ảnh chụp thật (Excel 365 Windows, zoom 100%) — xem CLAUDE.md, mục Ảnh minh họa.
    Có file Materials/B6/Images/<img> thì chèn ảnh vào khung (giữ tỉ lệ, canh giữa, viền hairline);
    chưa có thì để khung chờ [CHỤP] ghi nội dung, file · sheet · vùng và tên file ảnh cần lưu."""
    path = os.path.join(IMG, img)
    if os.path.exists(path):
        pic = s.shapes.add_picture(path, Inches(x), Inches(y), width=Inches(w))
        if pic.height > Inches(h):
            ratio = Inches(h) / pic.height; pic.height = Inches(h); pic.width = int(pic.width * ratio)
        pic.left = Inches(x) + (Inches(w) - pic.width) // 2; pic.top = Inches(y) + (Inches(h) - pic.height) // 2
        pic.line.color.rgb = rgb(HAIR); pic.line.width = Pt(0.75); pic.name = f"Screenshot {img[:-4]}"
        return pic
    sh = box(s, x, y, w, h, fill=ALT, line="A6A6A6", lw=1.25, name=name)
    sh.line.dash_style = MSO_LINE.DASH
    tb(s, x + 0.2, y + 0.1, w - 0.4, h - 0.2, [[("[CHỤP] ", {"b": True, "color": NAVY}), (what, {"color": INK2})],
                                               [(f"{where} → {img}", {"color": MUTED, "size": 12})]], size=14, anchor="ctr", align="ctr", space=4)
    MISSING.append(img)
    return sh

def footnote(s, text):
    tb(s, L, 6.8, W, 0.35, text, size=12, color=MUTED, anchor="b", name="Footnote")

# ---------------------------------------------------------------------
# Speaker notes từ kịch bản
# ---------------------------------------------------------------------
md = open(SCRIPT, encoding="utf-8").read()
blocks = re.split(r"^### Slide (\d+) · ", md, flags=re.M)
NOTES = {}

def clean(t):
    t = re.sub(r"\*\*(.+?)\*\*", r"\1", t); t = re.sub(r"(?<![\w*])\*(?=\S)(.+?)(?<=\S)\*(?![\w*])", r"\1", t)
    return t.replace("`", "").strip()

def fmt_min(m):
    return (f"{m:g}").replace(".", ",")

for k in range(1, len(blocks), 2):
    n = int(blocks[k]); body = blocks[k + 1]
    head, rest = body.split("\n", 1)
    body = re.split(r"^(?:---|## )", rest, flags=re.M)[0]
    m = re.search(r"· ([\d.]+)'\s*$", head) or re.search(r"\((\d+) phút\)", head)
    out = []
    ans = re.search(r"\*\*([A-D])\)", body)
    if head.startswith("QUESTION"):
        out.append(f"Khoảng 35 giây. Đáp án: {ans.group(1)}.")
    elif m:
        out.append(f"Thời lượng: {fmt_min(float(m.group(1)))} phút.")
    for sec in re.finditer(r"^\*\*(Lời giảng[^*]*|Ghi chú giảng viên[^*]*):\*\*\n(.*?)(?=^\*\*(?:Lời giảng|Ghi chú giảng viên)|\Z)", body, flags=re.M | re.S):
        label, txt = sec.group(1), sec.group(2)
        head_, _, tail_ = label.partition(" (")
        out.append(head_.upper() + (f" ({tail_}" if tail_ else ""))
        for line in txt.splitlines():
            line = line.rstrip()
            if line.startswith(">"): line = line[1:].strip()
            elif line.lstrip().startswith("- "): line = line.lstrip()[2:]
            else: line = line.strip()
            if line: out.append(clean(line))
    NOTES[n] = "\n".join(out)
assert len(NOTES) == 41, sorted(NOTES)

def notes(s, n):
    s.notes_slide.notes_text_frame.text = NOTES[n]

# ---------------------------------------------------------------------
# Quiz: 4 phương án, đáp án hiện khi bấm (Q1 = slide 34 → nhãn "Question 1")
# ---------------------------------------------------------------------
def quiz(n, topic, question, options, answer, rail="PHẦN 8 · QUIZ & TỔNG KẾT", qsize=22, osize=18):
    s = content(f"Question {n - 35}: {topic}", rail)
    tb(s, L, 1.65, W, 1.0, question, size=qsize, color=INK)
    pos = {"A": (L, 2.85), "B": (7.33, 2.85), "C": (L, 4.35), "D": (7.33, 4.35)}
    ow, oh = 5.41, 1.3
    for letter, (x, y) in pos.items():
        box(s, x, y, ow, oh, fill=WHITE, line=HAIR, name=f"Option {letter}")
        if letter == answer:
            box(s, x, y, ow, oh, fill=GOOD_BG, line=GOOD, lw=3, name=f"Answer highlight {letter}")
        lt = box(s, x + 0.2, y + (oh - 0.6) / 2, 0.6, 0.6, fill=NAVY, shape=MSO_SHAPE.OVAL, name=f"Letter {letter}")
        _center_text(lt, letter)
        if letter == answer:
            al = box(s, x + 0.2, y + (oh - 0.6) / 2, 0.6, 0.6, fill=GOOD, shape=MSO_SHAPE.OVAL, name=f"Answer letter {letter}")
            _center_text(al, letter)
        tb(s, x + 1.0, y, ow - 1.1, oh, options[letter], size=osize, anchor="ctr")
    for letter, (x, y) in pos.items():
        if letter != answer: box(s, x, y, ow, oh, fill=WHITE, name=f"Dim {letter}", alpha=60000)
    ax, ay = pos[answer]
    chip(s, ax + ow - 1.6, ay - 0.2, 1.45, "✓ Đáp án", fill=GOOD, line=None, color=WHITE, size=14, name="Answer chip", h=0.32)
    _reveal(s, [sh for sh in s.shapes if sh.name.startswith(("Answer", "Dim"))])
    notes(s, n)

def _center_text(sh, text):
    tf = sh.text_frame; tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER; _runs(p, [text], 20, WHITE, True)

def _reveal(s, shapes):
    """Một lần bấm: mọi shape đáp án/làm mờ hiện dần (fade), giống deck B1."""
    ids = [sh.shape_id for sh in shapes]
    cid = [3]
    def nid(): cid[0] += 1; return cid[0]
    pars = []
    for k, spid in enumerate(ids):
        a, b = nid(), nid(); c = nid()
        node = "clickEffect" if k == 0 else "withEffect"
        pars.append(f'<p:par><p:cTn id="{a}" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" nodeType="{node}">'
                    f'<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:set><p:cBhvr><p:cTn id="{b}" dur="1" fill="hold">'
                    f'<p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>'
                    f'<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>'
                    f'<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="{c}" dur="400"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cBhvr></p:animEffect>'
                    f'</p:childTnLst></p:cTn></p:par>')
    bld = "".join(f'<p:bldP spid="{i}" grpId="0" animBg="1"/>' for i in ids)
    xml = ('<p:timing xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:tnLst><p:par>'
           '<p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek">'
           '<p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst><p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>'
           f'<p:childTnLst><p:par><p:cTn id="{nid() + 900}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>{"".join(pars)}</p:childTnLst></p:cTn></p:par>'
           '</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
           '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst>'
           f'<p:bldLst>{bld}</p:bldLst></p:timing>')
    s._element.append(etree.fromstring(xml))

# =====================================================================
# SLIDES
# =====================================================================
BLUE, ORANGE = "2A78D6", "EB6834"
R0 = "PHẦN 0 · MỞ ĐẦU"; R1 = "PHẦN 1 · DATA MODEL"; R2 = "PHẦN 2 · RELATIONSHIP"
R3 = "PHẦN 3 · COLUMN VS MEASURE"; R4 = "PHẦN 4 · FILTER CONTEXT"; R5 = "PHẦN 5 · CALCULATE & ALL"
R6 = "PHẦN 6 · TIME INTELLIGENCE"; R7 = "PHẦN 7 · KPI DICTIONARY"; R8 = "PHẦN 8 · THỰC HÀNH"; R9 = "PHẦN 9 · QUIZ & TỔNG KẾT"

def card(s, x, y, w, h, title, body, size=15, fill=WHITE, line=HAIR, title_color=NAVY, name="Card"):
    box(s, x, y, w, h, fill=fill, line=line, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05, name=name)
    paras = [[(title, {"b": True, "color": title_color, "size": size + 1})]] + [p if isinstance(p, list) else [p] for p in body]
    tb(s, x + 0.18, y + 0.12, w - 0.36, h - 0.24, paras, size=size, space=5)

def bricks(s, done, nxt):
    box(s, 9.07, 4.6, 3.67, 1.0, fill=NAVY, name=f"Brick {done[:2]}")
    tb(s, 9.17, 4.6, 3.47, 1.0, "✓  " + done, size=18, color=WHITE, bold=True, anchor="ctr")
    box(s, 9.07, 3.3, 3.67, 1.0, fill=WHITE, line=BLUE, lw=2.25, name=f"Brick {nxt[:2]}")
    tb(s, 9.17, 3.3, 3.47, 1.0, nxt, size=18, color=NAVY, bold=True, anchor="ctr")

def arrow(s, x1, y1, x2, y2, color=ACCENT, w=2.25, name="Arrow"):
    a = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    a.line.color.rgb = rgb(color); a.line.width = Pt(w); a.name = name
    ln = a.line._get_or_add_ln(); te = etree.SubElement(ln, qn("a:tailEnd")); te.set("type", "triangle")
    return a

def tbl_box(s, x, y, w, h, title, sub, fill=WHITE, line=NAVY, tcolor=NAVY, name="Model table"):
    """[VẼ] một bảng trong mô hình: tên bảng + số dòng / vai trò."""
    box(s, x, y, w, h, fill=fill, line=line, lw=1.5, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.08, name=name)
    tb(s, x + 0.1, y + 0.05, w - 0.2, h - 0.1, [[(title, {"b": True, "color": tcolor, "size": 15})], [(sub, {"size": 12, "color": INK2})]],
       anchor="ctr", align="ctr", space=2)

def numbered(s, x, y, w, h, items, size=16, space=12):
    tb(s, x, y, w, h, [[(f"{k}. ", {"b": True, "color": NAVY}), t] if isinstance(t, str) else [(f"{k}. ", {"b": True, "color": NAVY})] + t
                       for k, t in enumerate(items, 1)], size=size, space=space)

R0 = "PHẦN 0 · MỞ ĐẦU"; R1 = "PHẦN 1 · BRIEF"; R2 = "PHẦN 2 · RÕ RÀNG & ĐƠN GIẢN"; R3 = "PHẦN 3 · CHỌN CHART"
R4 = "PHẦN 4 · BỐ CỤC & CÔNG CỤ"; R5 = "PHẦN 5 · DỰ BÁO"; R6 = "PHẦN 6 · BUILD"; R7 = "PHẦN 7 · KỂ CHUYỆN & PHÁT HÀNH"; R8 = "PHẦN 8 · QUIZ & TỔNG KẾT"
GREEN, ORANGE_, WARN_BG = "1BAF7A", "EB6834", "FFEB9C"

def practice(n, title, rail, steps, check_title, check_lines, foot=None, size=16, steps_w=6.9):
    s = content(title, rail)
    numbered(s, L, 1.6, steps_w, 5.0, steps, size=size, space=10)
    box(s, L + steps_w + 0.3, 1.6, W - steps_w - 0.3, 5.05, fill=ALT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.03, name="Check box")
    tb(s, L + steps_w + 0.5, 1.75, W - steps_w - 0.7, 4.8,
       [[(check_title, {"b": True, "color": NAVY, "size": 17})]] + check_lines, size=15, space=8)
    if foot: footnote(s, foot)
    notes(s, n)
    return s

def tool_slide(n, title, rail, name, purpose):
    s = content(title, rail)
    tool(s, L, 1.5, W, name, purpose)
    return s

# 1 · Bìa
s1 = prs.slides[0]
for sh in s1.shapes:
    if sh.name.startswith("subTitle"):
        sh.text_frame.paragraphs[0].runs[0].text = "Buổi 6: Thiết kế & Xây dựng Dashboard"
        for r in sh.text_frame.paragraphs[0].runs[1:]: r.text = ""
    elif sh.has_text_frame and sh.text_frame.text.startswith("Yêu cầu máy"):
        p = sh.text_frame.paragraphs[0]
        p.runs[0].text = "Bắt buộc máy Windows có Power Pivot (CUBEVALUE cần Data Model)"
        for r in p.runs[1:]: r.text = ""
        p2 = copy.deepcopy(p._p); sh.text_frame._txBody.append(p2)
        p2r = sh.text_frame.paragraphs[-1].runs
        p2r[0].text = "Mở sẵn: file khởi đầu B5_MCI_Mart_Answer.xlsx (đáp án Buổi 5)"
        for r in p2r[1:]: r.text = ""
        sh.height = Inches(0.9)
    elif sh.has_text_frame and sh.text_frame.text.startswith("Mở sẵn"):
        p = sh.text_frame.paragraphs[0]
        p.runs[0].text = "Mở sẵn: file khởi đầu B5_MCI_Mart_Answer.xlsx (đáp án Buổi 5)"
        for r in p.runs[1:]: r.text = ""
s1.notes_slide.notes_text_frame.text = NOTES[1]

# 2 · Câu hỏi mở buổi
s = content("Câu hỏi mở buổi", R0)
tb(s, L, 1.75, W, 1.2, [[("“Giám đốc bán hàng có 5 giây. Họ cần thấy gì đầu tiên?”", {"b": True, "color": NAVY})]], size=28)
chip(s, L, 3.05, 4.3, "Tầng 4 — Prescriptive · nhập môn Tầng 5", fill=WHITE, line=NAVY, color=NAVY, size=14, name="Tier chip")
box(s, L, 3.75, 5.2, 2.2, fill=WHITE, line=BAD_TXT, lw=1.5, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05, name="Dashboard A")
tb(s, L + 0.2, 3.85, 4.8, 2.0, [[("Dashboard A", {"b": True, "color": BAD_TXT, "size": 18})],
                              "12 biểu đồ · pie 3D · trục cắt khúc", "Không có target", "Người xem không biết nhìn đâu"], size=15, space=5)
tb(s, 6.97, 3.75, 0.8, 2.2, [[("≠", {"b": True, "color": BAD_TXT, "size": 44})]], anchor="ctr", align="ctr", name="Not equal")
box(s, 7.54, 3.75, 5.2, 2.2, fill=WHITE, line=GOOD, lw=1.5, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05, name="Dashboard B")
tb(s, 7.74, 3.85, 4.8, 2.0, [[("Dashboard B", {"b": True, "color": GOOD, "size": 18})],
                           "6 thẻ KPI · mốc target · ✔ ▲ ✘", "Thứ bậc: KPI → Trend → Detail", "Nhìn 5 giây là biết đạt hay chưa"], size=15, space=5)
tb(s, L, 6.15, W, 0.5, "Tháng 4/2025: 296.965.000 · 103,1% target ✔", size=18, color=NAVY, bold=True)
notes(s, 2)

# 3 · Recap B5 & file khởi đầu
s = content("Recap Buổi 5 & file khởi đầu", R0)
tb(s, L, 1.75, 7.1, 4.8, [
    [("Viên gạch B5: ", {"b": True, "color": NAVY}), "mô hình ngôi sao 6 bảng, 6 relationship; 14 measure dùng chung; KPI dictionary; đối soát 64 ô với KPI_SUMIFS khớp 0."],
    [("Bài về nhà B5: ", {"b": True, "color": NAVY}), "báo cáo của phòng Kinh doanh có 4 lỗi — Target chưa nối Calendar, tháng lấy từ Sales[Thang], GM% trung bình của biên từng dòng, số đơn đếm dòng. Đối soát trước, kết luận sau."],
    [("File khởi đầu hôm nay ", {"b": True, "color": NAVY}), "= đáp án B5 (B5_MCI_Mart_Answer.xlsx). Lưu thành B6_<HoTen>.xlsx. KHÔNG bấm Refresh All. Mọi con số trên dashboard đến từ các measure này."],
], size=18, bullet=True, space=16)
bricks(s, "B5 · Star schema + DAX", "B6 · Dashboard")
footnote(s, "Demo: ../B5_Data_Model_DAX/Practice/B5_MCI_Mart_Answer.xlsx (sheet KPI_Dictionary) · ../B5_Data_Model_DAX/Homework/B5_BTVN_MCI_Mart_KPI_T1_T4_2025_Answer.xlsx (sheet Dap_an_loi)")
notes(s, 3)

# 4 · Chia phần 1
notes(section("1. Dashboard brief: brief trước, chart sau"), 4)

# 5 · Dashboard brief
s = content("Dashboard brief", R1)
tb(s, L, 1.5, W, 0.5, [[("Thứ tự: ", {"b": True, "color": NAVY}), "Audience → Câu hỏi / quyết định → KPI → Layout → Chart → Build → Story → Phát hành"]], size=16)
table(s, L, 2.1, [2.0, 3.3, 5.77], [
    ["Mục", "Câu hỏi", "Dashboard MCI Mart"],
    ["Người xem", "Executive, operational hay self-service?", "Sales Director — executive, xem trong họp ban điều hành"],
    ["Câu hỏi (3–5)", "Chỉ số · chiều · thời gian · mốc so sánh", "Q1 Tháng này đạt bao nhiêu % target, so với tháng trước và cùng kỳ? · Q2 Cửa hàng nào dưới 90% target? · Q3 Ngày nào bất thường? · Q4 5 sản phẩm đầu chiếm bao nhiêu % doanh thu? · Q5 Thành phố nào đóng góp nhiều nhất?"],
    ["Tần suất refresh", "Bao lâu một lần?", "Hằng tháng, khi có file tháng mới"],
    ["Quyết định", "Dashboard giúp quyết định gì?", "Hỗ trợ cửa hàng nào trước; điều chỉnh target"],
], size=13, rowh=0.62, bold_first=True)
tip(s, L, 5.85, W, 0.85, "Câu hỏi yếu: “Xem doanh thu”, “Phân tích khách hàng” — chỉ là chủ đề, thiếu chỉ số, chiều, mốc so sánh.")
footnote(s, "Mẫu brief: sheet Brief trong B6_<HoTen>.xlsx (tạo ở Slide 6) · Tiêu chí C1: ../B7_Capstone/Rubric_B7_Capstone.md")
notes(s, 5)

# 6 · Thực hành brief
practice(6, "Thực hành: viết brief của bạn", R1, [
    "Tạo sheet Brief trong B6_<HoTen>.xlsx.",
    "Chép bảng 4 mục của Slide 5, rồi viết lại 2 câu hỏi theo cách của bạn: đủ chỉ số · chiều · thời gian · mốc so sánh.",
    "Đổi cặp với bạn bên cạnh: câu nào còn là chủ đề thì gạch và viết lại.",
], "Kiểm tra (5 phút)", ["Mỗi câu hỏi có đủ 4 thành phần.", "Không câu nào bắt đầu bằng “Xem…” hoặc “Phân tích…”.",
                         [("Brief này là bài tập của tiêu chí C1 trong capstone.", {"b": True, "color": NAVY})]])

# 7 · Chia phần 2
notes(section("2. Rõ ràng và đơn giản: dashboard tốt, xấu, “ugly”"), 7)

# 8 · Bài kiểm tra 5 giây
s = content("Bài kiểm tra 5 giây", R2)
shot(s, L, 1.55, 7.2, 4.55, "Dashboard xấu: 12 visual, pie 3D, trục cắt khúc, không target (toàn trang)", "Demo/B6_Demo_DashboardXau.xlsx · Dashboard_Xau", "S08_dashboard_xau.png")
box(s, 9.07, 1.55, 3.67, 4.55, fill=ALT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.04, name="Prompt box")
tb(s, 9.25, 1.7, 3.3, 4.3, [[("Bạn có 5 giây. Ghi lại:", {"b": True, "color": NAVY, "size": 17})],
                           "1. KPI chính là gì?", "2. Đạt hay chưa đạt target?", "3. Nên nhìn vào đâu tiếp?",
                           [("Sau đó: đếm có bao nhiêu vấn đề.", {"b": True})]], size=16, space=10)
tip(s, L, 6.15, W, 0.6, "Mỗi vấn đề tìm được sẽ thành một dòng của checklist chấm chéo (Slide 34).")
footnote(s, "Ảnh: Demo/B6_Demo_DashboardXau.xlsx, sheet Dashboard_Xau (toàn trang)")
notes(s, 8)

# 9 · Tốt · Xấu · Ugly
s = content("Tốt · xấu · “ugly”: 5 nguyên tắc", R2)
table(s, L, 1.5, [0.45, 2.7, 3.3], [
    ["#", "Nguyên tắc", "Ví dụ trong dữ liệu"],
    ["1", "Trục bar/column bắt đầu từ 0", "T3 vs T4/2024: chênh 0,96%; trục từ 236 triệu → cột cao gấp 2,9 lần"],
    ["2", "Không 3D", "Chiều sâu làm méo phần trăm"],
    ["3", "Pie ≤ 5 phần", "6 ngành: 34,5 · 22,5 · 18,0 · 11,6 · 11,6 · 1,8% → dùng bar"],
    ["4", "Mỗi KPI có mốc so sánh", "296 triệu một mình không nói gì; so target 288 triệu → “đạt”"],
    ["5", "Màu theo thực thể", "Bắc luôn xanh dương; target luôn xám nét đứt"],
], size=13, rowh=0.74)
shot(s, 8.3, 1.5, 4.44, 4.45, "Hai biểu đồ cột T3/2024 và T4/2024 cạnh nhau: trục từ 236.000.000 (xấu) và từ 0 (tốt)", "Demo/B6_Demo_DashboardXau.xlsx · Demo_TrucY", "S09_truc_y.png")
tip(s, L, 6.1, W, 0.65, "Line chart được phép cắt trục (ta nhìn hình dáng); bar và column thì không (ta so độ cao).")
footnote(s, "Ảnh: Demo/B6_Demo_DashboardXau.xlsx, sheet Demo_TrucY (hai biểu đồ cột cạnh nhau)")
notes(s, 9)

# 10 · Chia phần 3
notes(section("3. Chọn chart: 3 câu hỏi, 8 loại lõi"), 10)

# 11 · 3 câu hỏi + bảng chọn chart
s = content("3 câu hỏi then chốt · bảng chọn chart", R3)
for k, t in enumerate(["1. Dữ liệu loại gì? (số, danh mục, thời gian, địa lý)", "2. Muốn truyền đạt gì? (so sánh, cơ cấu, xu hướng, phân phối, quan hệ)", "3. Ai xem? (CEO, analyst, khách hàng)"]):
    box(s, L + k * 3.73, 1.5, 3.61, 0.95, fill=ALT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06, name=f"Question {k+1}")
    tb(s, L + k * 3.73 + 0.12, 1.55, 3.37, 0.85, t, size=13, anchor="ctr", bold=True, color=NAVY)
table(s, L, 2.65, [4.4, 3.5, 3.17], [
    ["Câu hỏi dữ liệu", "Chart lõi", "Thay thế"],
    ["Xu hướng theo thời gian", [("Line", {"b": True, "color": RED})], "Area"],
    ["So sánh giữa các nhóm", [("Bar / Column", {"b": True, "color": RED})], "Bullet"],
    ["Cơ cấu / tỷ trọng", [("Stacked bar", {"b": True, "color": RED})], "Donut (≤ 5 phần), Treemap"],
    ["Hai thước đo khác đơn vị; Actual vs Target", [("Combo 2 trục", {"b": True, "color": RED})], "Column + đường target"],
    ["Phân phối, outlier", [("Histogram", {"b": True, "color": RED})], "Box & Whisker"],
    ["Quan hệ giữa hai biến", [("Scatter (+ trendline)", {"b": True, "color": RED})], "Bubble"],
    ["Địa lý", [("Map", {"b": True, "color": RED})], "Bar theo vùng"],
    ["Chênh lệch cộng dồn (target → actual)", [("Waterfall", {"b": True, "color": RED})], "—"],
], size=14, rowh=0.44)
footnote(s, "Chart ít dùng (Area, Pie, Bubble, Radar, Funnel, Sunburst…): xem handout B6")
notes(s, 11)

# 12 · Line · Column · Combo
s = content("Line · Column · Combo 2 trục", R3)
tool(s, L, 1.45, 5.3, "Insert → Combo", "Hai thước đo khác đơn vị (VNĐ và %) trên cùng một chart; hoặc Actual (cột) với Target (đường).")
tb(s, L, 2.75, 5.4, 2.6, [
    [("Line: ", {"b": True, "color": NAVY}), "xu hướng · ", ("Column/Bar: ", {"b": True, "color": NAVY}), "so sánh nhóm (trục từ 0)"],
    [("Combo ngành, T1–T4/2025: ", {"b": True, "color": NAVY}), "cột = Doanh thu (Bánh kẹo 426.916.000 … Sữa 22.790.000); đường = Gross Margin % ở trục phụ (Bánh kẹo 39,4% … Thực phẩm khô 21,6%)."],
    [("Insight: ", {"b": True, "color": NAVY}), "Thực phẩm khô lớn thứ ba (222.348.000) nhưng biên thấp nhất (21,6%)."],
], size=14, space=8)
shot(s, 7.4, 1.5, 5.34, 3.9, "Combo: cột Doanh thu + đường GM% (trục phụ) theo 6 ngành", "Demo/B6_Demo_Chart.xlsx · Demo_Combo · A3:C9", "S12_combo_nganh.png")
tip(s, L, 5.65, W, 0.95, "Cột trục chính (VNĐ), đường trục phụ (%). Cho trục phụ bắt đầu từ 0%: hai trục lệch nhau rất dễ gây hiểu lầm.")
footnote(s, "Ảnh: Demo/B6_Demo_Chart.xlsx, sheet Demo_Combo (A3:C9)")
notes(s, 12)

# 13 · Cơ cấu · Waterfall · Map
s = content("Cơ cấu · Waterfall · Map", R3)
tool(s, L, 1.45, 5.3, "Insert → Waterfall", "Từ Target đến Actual: ai kéo lên, ai kéo xuống.")
tb(s, L, 2.6, 5.3, 3.0, [
    [("Cơ cấu: ", {"b": True, "color": NAVY}), "Stacked bar thay pie khi hơn 5 phần; Treemap khi cần thấy kích thước."],
    [("Waterfall T3/2025: ", {"b": True, "color": NAVY}), "Target 285.000.000 → từng cửa hàng (ST12 −11.749.000, ST07 −10.461.000, ST08 +10.733.000 …) → Actual 269.608.000 (tổng −15.392.000). Cột đầu và cuối: Set as Total."],
    [("Map: ", {"b": True, "color": NAVY}), "theo thành phố (TP.HCM 108.885.000, T4). Cần Internet; không có thì dùng bar."],
], size=14, space=8)
shot(s, 7.2, 1.5, 5.54, 3.9, "Waterfall từ Target 285.000.000 qua 12 cửa hàng đến Actual 269.608.000, hai cột Total", "Demo/B6_Demo_Chart.xlsx · Demo_Waterfall · A3:B17", "S13_waterfall.png")
box(s, L, 5.65, W, 1.0, fill=BAD_BG, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06, name="Limit callout")
tb(s, L + 0.2, 5.7, W - 0.4, 0.9, [[("Giới hạn: ", {"b": True, "color": BAD_TXT}), "PivotChart không làm được Waterfall, Map, Histogram, Treemap — chúng lấy dữ liệu từ một vùng ô. Muốn chạy theo Slicer, lấp vùng ô bằng CUBEVALUE (Slide 18, 29)."]], size=15, anchor="ctr")
footnote(s, "Ảnh: Demo/B6_Demo_Chart.xlsx, sheet Demo_Waterfall (A3:B17)")
notes(s, 13)

# 14 · Phân phối · Quan hệ
s = content("Histogram · Scatter", R3)
tool(s, L, 1.4, 11.0, "Insert → Histogram · Scatter · CORREL", "Phân phối và quan hệ giữa hai biến.")
tb(s, L, 2.55, 4.4, 3.1, [
    [("Histogram, giá trị đơn Q1: ", {"b": True, "color": NAVY}), "3.144 đơn · trung vị 198.000 · trung bình 299.445 · ngưỡng 896.500 · 138 đơn vượt. Bin width 100.000."],
    [("Scatter, tuổi cửa hàng ↔ doanh thu T1–T4: ", {"b": True, "color": NAVY}), "12 cửa hàng, r = 0,78. Hà Nội + TP.HCM: 4,3 năm · 137,5 triệu; còn lại 2,7 năm · 78,7 triệu."],
    [("Tương quan ≠ nhân quả.", {"b": True, "color": BAD_TXT})],
], size=13, space=8)
shot(s, 6.3, 2.5, 3.2, 3.0, "Histogram giá trị đơn Q1, bin 100.000, đuôi phải", "Demo_Phan_phoi", "S14a_histogram.png")
shot(s, 9.6, 2.5, 3.14, 3.0, "Scatter 12 cửa hàng + trendline + R²", "Demo_Scatter", "S14b_scatter.png")
tip(s, L, 5.75, W, 0.95, "Histogram: Format Axis → Bin width. Scatter: Add Trendline → Display R-squared; ô riêng gõ =CORREL(x, y).")
footnote(s, "Ảnh: Demo/B6_Demo_Chart.xlsx, sheet Demo_Phan_phoi (cột B) và Demo_Scatter (A3:C15)")
notes(s, 14)

# 15 · Nghỉ
s = section("Nghỉ giữa giờ", "10 phút · Sau giờ nghỉ: bố cục, KPI card bằng CUBEVALUE, RAG, Slicer")
notes(s, 15)

# 16 · Chia phần 4
notes(section("4. Bố cục & công cụ: CUBEVALUE · RAG · Slicer"), 16)

# 17 · Thứ bậc bố cục
s = content("Thứ bậc: KPI → Trend → Detail", R4)
box(s, L, 1.55, 8.0, 5.1, fill=WHITE, line=NAVY, lw=1.5, name="Dashboard frame")
tb(s, L + 0.15, 1.6, 7.7, 0.35, "Slicer: Tháng · Miền", size=12, color=MUTED)
for k in range(6):
    box(s, L + 0.15 + k * 1.3, 2.0, 1.2, 0.95, fill=NAVY if k == 0 else ALT, line=HAIR, name=f"KPI card {k+1}")
    tb(s, L + 0.15 + k * 1.3, 2.0, 1.2, 0.95, ["KPI chính" if k == 0 else f"KPI {k+1}"], size=12, color=WHITE if k == 0 else INK, bold=True, anchor="ctr", align="ctr")
tb(s, L + 0.15, 3.0, 4, 0.3, "Hàng 1 — KPI (góc trên trái = quan trọng nhất)", size=11, color=MUTED)
box(s, L + 0.15, 3.35, 3.8, 1.25, fill=ALT, line=HAIR, name="Line"); tb(s, L + 0.15, 3.35, 3.8, 1.25, "Line theo ngày", size=13, anchor="ctr", align="ctr", bold=True, color=NAVY)
box(s, L + 4.1, 3.35, 3.75, 1.25, fill=ALT, line=HAIR, name="Combo"); tb(s, L + 4.1, 3.35, 3.75, 1.25, "Combo Actual vs Target", size=13, anchor="ctr", align="ctr", bold=True, color=NAVY)
tb(s, L + 0.15, 4.65, 4, 0.3, "Hàng 2 — Trend", size=11, color=MUTED)
for k, t in enumerate(["Waterfall theo cửa hàng", "Top 5 sản phẩm", "Map theo thành phố"]):
    box(s, L + 0.15 + k * 2.65, 5.0, 2.5, 1.45, fill=ALT, line=HAIR, name=f"Detail {k+1}")
    tb(s, L + 0.15 + k * 2.65, 5.0, 2.5, 1.45, t, size=13, anchor="ctr", align="ctr", bold=True, color=NAVY)
tb(s, 9.9, 1.6, 2.84, 5.0, [
    [("Tối đa 5–8 visual", {"b": True, "color": NAVY})], "Hôm nay: 7",
    [("Mỗi visual trả lời một câu hỏi brief", {"b": True, "color": NAVY})], "Thẻ KPI + combo → Q1 · Waterfall → Q2 · Line → Q3 · Top 5 → Q4 · Map → Q5",
    [("Đòi thêm chart?", {"b": True, "color": NAVY})], "Hỏi: “Nó trả lời câu nào trong brief?”"], size=14, space=8)
notes(s, 17)

# 18 · CUBEVALUE
s = content("CUBEVALUE · CUBEMEMBER · KPI card", R4)
tool(s, L, 1.4, 8.0, "CUBEVALUE · CUBEMEMBER", "Đọc một con số từ measure trong Data Model vào một ô — nền của thẻ KPI.")
fx(s, L, 2.55, W, '=CUBEVALUE("ThisWorkbookDataModel","[Measures].[Total Revenue]","[Calendar].[Nam].&[2025]",Slicer_Thang,Slicer_Region)', size=12)
tb(s, L, 3.2, 5.1, 3.4, [
    [("Đối số: ", {"b": True, "color": NAVY}), "kết nối · measure · các bộ lọc (member hoặc Slicer) = filter context của B5."],
    [("Bắc × Tháng 3: ", {"b": True, "color": NAVY}), "90.210.000 (= DoiSoat_B2, B5)."],
    [("Tháng 4/2025: ", {"b": True, "color": NAVY}), "Doanh thu 296.965.000 · Target 288.000.000 · Revenue vs Target % 103,1%."],
    [("PivotTable Analyze → OLAP Tools → Convert to Formulas: ", {"b": True, "color": NAVY}), "Excel tự viết CUBEVALUE từ Pivot."],
], size=14, space=8)
shot(s, 6.97, 3.2, 2.7, 2.2, "Ô CUBEVALUE, gợi ý autocomplete", "Data_Dash · B3", "S18a_cubevalue.png")
shot(s, 9.77, 3.2, 2.97, 2.2, "Data_Dash: 6 KPI giá trị · target · mức đạt (T4)", "Data_Dash · A3:E9", "S18b_data_dash.png")
tip(s, L, 5.65, W, 1.0, "Gõ sai tên measure → #N/A. Dùng gợi ý khi gõ, hoặc Convert to Formulas. Mức đạt = giá trị ÷ target là phép so sánh với hằng số, không phải định nghĩa KPI.")
footnote(s, "Ảnh: B6_<HoTen>.xlsx, sheet Data_Dash (A3:E9). Mẫu: Practice/B5_MCI_Mart_Answer.xlsx, sheet DoiSoat_B2 (ô B6)")
notes(s, 18)

# 19 · RAG
s = content("RAG: màu + biểu tượng + chữ", R4)
tool(s, L, 1.4, 8.0, "Conditional Formatting → Icon Sets", "Trạng thái đạt / cảnh báo / không đạt — không bao giờ chỉ dùng màu.")
table(s, L, 2.5, [1.6, 2.3, 1.5, 1.6], [
    ["Mức", "Mức đạt", "Biểu tượng", "Chữ"],
    ["Đạt", "≥ 100%", [("✔", {"b": True, "color": GOOD})], "Đạt"],
    ["Cảnh báo", "90% – < 100%", [("▲", {"b": True, "color": "B7791F"})], "Cảnh báo"],
    ["Không đạt", "< 90%", [("✘", {"b": True, "color": BAD_TXT})], "Không đạt"],
], size=14, rowh=0.5, fills={(1, 0): GOOD_BG, (2, 0): WARN_BG, (3, 0): BAD_BG})
tb(s, L, 4.65, 7.0, 1.1, [
    [("Tháng 4: ", {"b": True, "color": NAVY}), "Doanh thu 103,1% ✔ · Đơn 96,6% ▲ · AOV 102,5% ✔ · GM ✔ · YoY ✔ · Khách 95,7% ▲"],
    [("Tháng 2: ", {"b": True, "color": NAVY}), "Doanh thu 82,5% ▲ · YoY −0,9% ✘"],
], size=14, space=6)
shot(s, 8.97, 2.5, 3.77, 3.0, "6 thẻ KPI tháng 4/2025 với ✔ ▲ ✘ và nhãn chữ", "Dashboard (hàng KPI)", "S19_rag.png")
tip(s, L, 5.75, 6.9, 0.95, "Bộ 3 Symbols (Circled): ba hình khác nhau, người mù màu vẫn đọc được. Ngưỡng ở sheet Cai_dat.")
tip(s, 8.77, 5.75, 3.97, 0.95, "Edit Rule → Type = Formula, trỏ ô Cai_dat.", pro=True)
footnote(s, "Ảnh: B6_<HoTen>.xlsx, sheet Dashboard (hàng KPI, tháng 4/2025). Ngưỡng: sheet Cai_dat (C13:C14)")
notes(s, 19)

# 20 · Slicer · Report Connections
s = content("Slicer nối nhiều Pivot", R4)
tool(s, L, 1.4, 8.0, "Report Connections", "Một Slicer điều khiển mọi PivotTable / PivotChart của dashboard.")
tb(s, L, 2.55, 5.4, 3.1, [
    "Slicer Tháng (Calendar[Thang]) và Miền (Store[Region]) sinh ra từ một Pivot; mặc định chỉ lọc Pivot đó.",
    [("Chuột phải Slicer → Report Connections → tick từng Pivot.", {"b": True, "color": NAVY})],
    [("Combo 4 tháng: ", {"b": True, "color": NAVY}), "chỉ nối Miền; ghi trên tiêu đề “Cả 4 tháng · không đổi theo Slicer Tháng”."],
    [("Thẻ KPI, waterfall, map: ", {"b": True, "color": NAVY}), "nghe Slicer qua đối số của CUBEVALUE."],
], size=14, space=8)
shot(s, 7.2, 2.5, 5.54, 3.15, "Hộp Report Connections của Slicer Tháng", "chuột phải Slicer Tháng", "S20_report_connections.png")
tip(s, L, 5.85, W, 0.85, "Pivot đổi cỡ cột khi Refresh làm vỡ bố cục → PivotTable Options → bỏ Autofit column widths on update.")
footnote(s, "Ảnh: B6_<HoTen>.xlsx, chuột phải Slicer Tháng → Report Connections")
notes(s, 20)

# 21 · Sparklines · Alt text
s = content("Sparklines · Alt text", R4)
tool(s, L, 1.4, 8.0, "Insert → Sparklines · Edit Alt Text", "Xu hướng thu nhỏ trong một ô; mô tả chart cho người đọc màn hình.")
tb(s, L, 2.55, 5.3, 3.2, [
    [("Sparkline ", {"b": True, "color": NAVY}), "dưới mỗi thẻ KPI: T1–T4 (Doanh thu 425,9 → 246,0 → 269,6 → 297,0 triệu). Dữ liệu từ hàng T1–T4 trong Data_Dash (CUBEVALUE theo từng tháng)."],
    [("Alt text: ", {"b": True, "color": NAVY}), "một câu cho mỗi chart, ví dụ “Biểu đồ cột và đường: doanh thu thực tế so với target, T1–T4/2025”."],
], size=14, space=10)
shot(s, 7.2, 2.5, 5.54, 3.2, "Thẻ Doanh thu có sparkline T1–T4", "Dashboard (thẻ Doanh thu)", "S21_sparkline.png")
footnote(s, "Ảnh: B6_<HoTen>.xlsx, sheet Dashboard (thẻ Doanh thu, sparkline) · Data_Dash (hàng T1–T4)")
notes(s, 21)

# 22 · Chia phần 5
notes(section("5. Dự báo nhập môn: sắp tới sẽ thế nào?"), 22)

# 23 · Trendline
s = content("Trendline tuyến tính + Forward", R5)
tool(s, L, 1.4, 8.0, "Trendline → Linear · Forecast Forward", "Kéo dài xu hướng bằng đường thẳng.")
tb(s, L, 2.5, 5.2, 3.3, [
    [("Bảng Ngày–Doanh thu: ", {"b": True, "color": NAVY}), "486 dòng, 01/01/2024–30/04/2025, không thiếu ngày."],
    [("Linear: ", {"b": True, "color": NAVY}), "độ dốc +4.742 VNĐ/ngày, R² = 0,045 — giải thích 4,5% biến động."],
    [("Forward 30 ngày: ", {"b": True, "color": NAVY}), "≈ 10,3 triệu/ngày, như nhau cả cuối tuần."],
    [("Thực tế: ", {"b": True, "color": NAVY}), "Thứ Hai 8,4 triệu → Chủ nhật 10,5 triệu (+25,1%)."],
    [("Đường này có tin được không?", {"b": True, "color": BAD_TXT})],
], size=14, space=8)
shot(s, 7.0, 2.5, 5.74, 3.2, "Line doanh thu ngày + Linear Trendline + Forward 30 + R² 0,045", "Demo/B6_Demo_DuBao.xlsx · Demo_Trendline", "S23_trendline.png")
tip(s, L, 5.9, W, 0.8, "Tin được cho mức cả tháng (kiểm chứng ngược: +4,1%), không tin cho từng ngày (sai trung bình gần 20%).")
footnote(s, "Ảnh: Demo/B6_Demo_DuBao.xlsx, sheet Demo_Trendline (tblNgay, biểu đồ line)")
notes(s, 23)

# 24 · Forecast Sheet
s = content("Forecast Sheet (ETS) · giới hạn", R5)
tool(s, L, 1.4, 11.0, "Data → Forecast Sheet · FORECAST.ETS · FORECAST.ETS.CONFINT", "Dự báo có nhịp lặp, kèm dải tin cậy.")
tb(s, L, 2.45, 5.2, 3.7, [
    [("Hộp thoại: ", {"b": True, "color": NAVY}), "Forecast End 30/05/2025 · Confidence 95% · Seasonality = Set Manually 7 (tự phát hiện ra 121, không phải 7) · Forecast Start có thể lùi để so với thực tế."],
    [("Kết quả T5/2025: ", {"b": True, "color": NAVY}), "237,7 triệu (thấp hơn T4 thực tế cỡ 20%); cuối tuần +24%; dải 95% ±6,6 triệu ngày đầu, ±18,0 triệu trung bình — rộng hơn cả giá trị dự báo. Kiểm chứng ngược T4: 280,0 vs 297,0 triệu (−5,7%)."],
    [("Giới hạn: ", {"b": True, "color": BAD_TXT}), "16 tháng chưa đủ 2 chu kỳ năm · Tết đổi ngày theo âm lịch (10/02/2024 · 29/01/2025 · 17/02/2026) · luôn kèm dải + giả định."],
    [("Dự báo = mẫu hình quá khứ + giả định.", {"b": True, "color": NAVY})],
], size=13, space=7)
shot(s, 6.9, 2.45, 2.7, 3.3, "Hộp thoại Create Forecast Worksheet (Options)", "Data → Forecast Sheet", "S24a_forecast_dialog.png")
shot(s, 9.7, 2.45, 3.04, 3.3, "Biểu đồ dự báo 30 ngày + hai đường Confidence Bound", "Demo_ETS", "S24b_forecast_chart.png")
tip(s, L, 6.0, W, 0.75, "Hai hàm: FORECAST.ETS cho điểm dự báo; FORECAST.ETS.CONFINT cho nửa chiều rộng dải.")
footnote(s, "Ảnh: Demo/B6_Demo_DuBao.xlsx (Data → Forecast Sheet trên tblNgay; hàm ETS ở sheet Demo_ETS)")
notes(s, 24)

# 25 · Chia phần 6
notes(section("6. Build: dựng dashboard từng bước"), 25)

# 26–29 · Build
practice(26, "Build 1: khung, Slicer, tham số", R6, [
    "Sheet Cai_dat: target KPI (Orders 1.000 · AOV 300.000 · GM 33% · YoY +10% · Khách 700) và ngưỡng RAG (100% · 90%).",
    "Sheet Dashboard: View → bỏ Gridlines; độ rộng cột đều; tiêu đề lớn ở A1.",
    "Sheet Data_Dash: vùng lấy dữ liệu cho thẻ KPI, waterfall, map (sẽ ẩn ở Slide 33).",
    "Pivot trống từ Data Model → chèn Slicer Tháng (Calendar[Thang]) và Miền (Store[Region]); đặt lên Dashboard; chọn tháng 4.",
    "Đặt tên Slicer Settings: Slicer_Thang, Slicer_Region.",
], "Kiểm tra (8 phút)", ["Slicer Tháng đang chọn 4.", "Cai_dat có 5 target + 2 ngưỡng.", "Điểm hay vấp: Slicer cần một Pivot làm gốc."])
practice(27, "Build 2: hàng KPI card", R6, [
    "Data_Dash: 6 KPI × (giá trị · target · mức đạt · nhãn) bằng CUBEVALUE (Slicer_Thang, Slicer_Region) + hàng T1–T4 cho sparkline.",
    "Dashboard: 6 thẻ (ô gộp, viền hairline, chữ lớn); thẻ Doanh thu ở góc trên trái.",
    "Mỗi thẻ: tên KPI · giá trị · “Target …” · mức đạt + biểu tượng + nhãn · sparkline.",
    "Bấm Slicer sang tháng 3, tháng 2 và quay lại tháng 4.",
], "Kiểm tra — tháng 4/2025", ["Doanh thu 296.965.000 (103,1%) ✔", "Đơn 966 (96,6%) ▲", "AOV 307.417 (102,5%) ✔", "GM 33,8% ✔ · YoY +24,0% ✔", "Khách 670 (95,7%) ▲",
                            [("Tháng 2: Doanh thu 82,5% ▲ · YoY −0,9% ✘", {"b": True})]], size=14)
practice(28, "Build 3: line ngày · combo vs Target", R6, [
    "PivotChart Line: Filters Calendar[Nam] = 2025 · Rows Calendar[Date] · Values Total Revenue → Line. Nối Slicer Tháng và Miền.",
    "PivotChart Combo: Rows Calendar[Thang] · Values Total Revenue + Total Target → Revenue = Column, Target = Line xám nét đứt. Rows: Value Filters → Total Revenue > 0. Nối chỉ Slicer Miền.",
    "Tiêu đề chart là kết luận; thêm chú thích chỉ vào điểm bất thường.",
], "Kiểm tra (8 phút)", ["Combo: T1 120,6% · T2 82,5% · T3 94,6% · T4 103,1%", "Target T1–T4: 353 · 298 · 285 · 288 triệu", "Line tháng 4: 30 ngày",
                       [("Combo = câu chuyện số 4: Tết 29/01 dồn mua sắm sang T1.", {"b": True, "color": NAVY})]])
practice(29, "Build 4: Top 5 · Waterfall · Map", R6, [
    "Top 5: Rows Product[Product_Name] · Values Total Revenue → Value Filters → Top 10 → 5 → PivotChart Bar; nối cả hai Slicer.",
    "Waterfall: Data_Dash A20:B33 — Target (Total Target), 12 Variance từng cửa hàng (CUBEMEMBER + CUBEVALUE), Actual (Total Revenue) → Insert → Waterfall; Set as Total đầu và cuối.",
    "Map hoặc Bar: doanh thu từng thành phố (CUBEVALUE, Slicer_Thang) → Filled Map; không có Internet → Bar.",
], "Kiểm tra — tháng 4", ["Top 5: Hộp quà Tết 35.100.000 · Chảo 29.155.000 · Dầu gội 26.565.000 · Mì tôm 19.712.000 · Bột giặt 19.276.000",
                        "Waterfall: Target 288.000.000 → Actual 296.965.000 (+8.965.000)", [("Tháng 3: ST07 −10.461.000", {"b": True, "color": NAVY})]], size=14)

# 30 · Build 5
s = content("Build 5: ráp trang · kiểm tra 5 giây", R6)
numbered(s, L, 1.55, 5.6, 4.2, [
    "Sắp 7 visual theo bố cục Slide 17; căn lề bằng Align; màu theo thực thể; target xám nét đứt.",
    "Report Connections: Tháng nối mọi Pivot trừ combo; Miền nối tất cả.",
    "Alt text cho từng chart; tiêu đề chart là kết luận.",
    "Kiểm tra 5 giây: đổi cặp — KPI chính · đạt hay chưa · nhìn đâu tiếp.",
], size=14, space=10)
shot(s, 7.4, 1.55, 5.34, 3.85, "Dashboard hoàn chỉnh, tháng 4/2025 (toàn trang)", "Practice/B6_MCI_Mart_Dashboard_Answer.xlsx · Dashboard", "S30_dashboard_t4.png")
tip(s, L, 5.65, W, 1.0, "Kiểm tra: bấm Slicer Tháng 4 → 3 → 2: thẻ, line, waterfall, Top 5, map đều đổi; combo không đổi. Tối đa 8 visual.")
footnote(s, "Ảnh: Practice/B6_MCI_Mart_Dashboard_Answer.xlsx, sheet Dashboard (tháng 4/2025)")
notes(s, 30)

# 31 · Chia phần 7
notes(section("7. Kể chuyện (SCR) & phát hành dashboard"), 31)

# 32 · SCR
s = content("Storytelling SCR", R7)
for k, (lab, body) in enumerate([
    ("S — Situation", "Con số có mốc so sánh. Tháng 3: 269.608.000, đạt 94,6% target (hụt 15.392.000), nhưng +13,7% so với cùng kỳ."),
    ("C — Complication", "Nguyên nhân đã kiểm chứng bằng drill-down. ST07 hụt 10.461.000 (44,9%), chiếm 68% khoản hụt, do cải tạo 10–31/03. ST12 dưới 90% ba tháng liền (76,9% · 53,0% · 65,2%)."),
    ("R — Resolution", "Đề xuất cho đúng người xem: (1) đánh giá ST07 theo số ngày mở cửa; (2) rà soát cách lập target ST12; (3) target 2026 theo lịch Tết âm (17/02/2026), đánh giá T1+T2 gộp."),
]):
    card(s, L + k * 3.73, 1.55, 3.61, 3.9, lab, [body], size=16, name=f"SCR {k+1}", title_color=NAVY)
tip(s, L, 5.65, W, 1.0, "Chú thích trực tiếp trên chart, chỉ vào điểm bất thường (ST07, T2). Giả thuyết chưa kiểm chứng phải ghi là giả thuyết.")
notes(s, 32)

# 33 · Phát hành
s = content("Phát hành dashboard", R7)
tool(s, L, 1.4, 11.0, "Refresh on open · Hide · Protect Sheet · Export PDF · Share", "Đóng gói để gửi cho người xem.")
table(s, L, 2.5, [1.9, 5.8], [
    ["Việc", "Cách làm"],
    ["Refresh on open", "Data → Queries & Connections → chuột phải query → Properties → Refresh data when opening the file (không thử trên lớp)"],
    ["Ẩn sheet kỹ thuật", "Hide Data_Dash, Cai_dat, các sheet Pivot"],
    ["Protect Sheet", "Review → Protect Sheet; bỏ Locked của Slicer, tick Use PivotTable & PivotChart"],
    ["Xuất PDF", "Landscape · Fit to 1 page → File → Export → Create PDF"],
    ["Chia sẻ", "OneDrive / SharePoint → Share; chú ý dữ liệu khách hàng, Sensitivity label"],
    [[("Giới hạn", {"b": True, "color": BAD_TXT})], "Excel không tự refresh theo lịch → cần Power BI"],
], size=12, rowh=0.62, bold_first=True)
shot(s, 9.55, 2.5, 3.19, 3.6, "Hộp Protect Sheet, tick Use PivotTable & PivotChart", "Review → Protect Sheet", "S33_protect_sheet.png")
footnote(s, "Ảnh: B6_<HoTen>.xlsx, Review → Protect Sheet (trên sheet Dashboard)")
notes(s, 33)

# 34 · Chấm chéo
s = content("Chấm chéo bằng checklist", R7)
items34 = ["Mỗi visual trả lời một câu trong brief", "KPI chính ở góc trên trái; KPI → Trend → Detail", "5–8 visual", "Mọi KPI có mốc so sánh và RAG kèm biểu tượng/nhãn",
           "Mọi số từ measure (CUBEVALUE/Pivot); không công thức KPI ở tầng hiển thị", "Trục bar/column từ 0; không 3D; không pie quá 5 phần",
           "Slicer điều khiển mọi visual (trừ visual ngữ cảnh đã ghi chú)", "Có chú thích chỉ vào điểm bất thường", "Bài kiểm tra 5 giây",
           "Phát hành: Refresh on open, ẩn sheet, Protect Sheet, alt text, PDF", "Nếu có dự báo: dải tin cậy và giả định"]
for k, t in enumerate(items34):
    col, row = divmod(k, 6)
    x, y = L + col * 5.6, 1.55 + row * 0.72
    box(s, x, y + 0.13, 0.3, 0.3, fill=WHITE, line=NAVY, lw=1.5, name=f"Checkbox {k+1}")
    tb(s, x + 0.45, y, 5.0, 0.62, t, size=14, anchor="ctr")
tip(s, L, 6.0, W, 0.7, "Chỉ chấm, không sửa. Ghi 1 điểm làm tốt và 1 điểm cần sửa.")
footnote(s, "Checklist: Course_Plan.md, Phụ lục D · Tiêu chí C5: ../B7_Capstone/Rubric_B7_Capstone.md")
notes(s, 34)

# 35 · Quiz cover
notes(section("Excel Quiz Challenge", "4 câu hỏi nhanh · câu 3–4 là tình huống"), 35)
quiz(36, "Chọn chart", "Muốn đặt doanh thu (VNĐ) và biên lợi nhuận gộp (%) của sáu ngành lên cùng một biểu đồ, nên dùng gì?",
     {"A": "Pie", "B": "Combo: cột doanh thu, đường biên % ở trục phụ", "C": "Scatter", "D": "Waterfall"}, "B", osize=16)
quiz(37, "Slicer", "Dashboard có ba PivotChart. Bạn bấm Slicer Miền, chỉ một chart đổi. Nguyên nhân và cách xử lý?",
     {"A": "Phải Refresh All", "B": "Slicer chưa nối các Pivot còn lại: chuột phải Slicer → Report Connections, tick đủ",
      "C": "Phải tạo Slicer riêng cho từng chart", "D": "Đổi Slicer thành Timeline"}, "B", osize=15)
quiz(38, "Tình huống — Dự báo Tết", "Giám đốc xem Forecast Sheet và hỏi: “Doanh thu Tết 2026 là bao nhiêu?” Câu trả lời đúng?",
     {"A": "Đọc số trên đường Forecast, chắc chắn đúng", "B": "Chạy lại với dải tin cậy 99%",
      "C": "Không nên dùng: mới 16 tháng dữ liệu, Tết đổi ngày theo âm lịch (17/02/2026); chỉ dự báo ngắn hạn, kèm dải tin cậy và giả định", "D": "Dùng Linear Trendline cho chắc"}, "C", qsize=20, osize=13)
quiz(39, "Tình huống — Trục Y", "Bar chart so doanh thu tháng 3 (237.192.000) và tháng 4 (239.466.000) năm 2024: cột tháng 4 cao gần gấp ba. Vì sao và sửa thế nào?",
     {"A": "Doanh thu thật sự tăng gấp ba", "B": "Trục Y bắt đầu từ 236.000.000: đặt trục từ 0 (chênh thật chỉ 0,96%)",
      "C": "Đổi sang pie 3D", "D": "Thêm đường lưới"}, "B", qsize=20, osize=15)

# 40 · Tổng kết + insight
s = content("Tổng kết & insight của bạn", R8)
tb(s, L, 1.65, 5.3, 0.45, "Điểm chính cần nhớ", size=18, color=NAVY, bold=True)
tb(s, L, 2.2, 5.3, 4.5, ["(1) Brief trước, chart sau: người xem, 3–5 câu hỏi đo lường được, tần suất, quyết định.",
                        "(2) Thứ bậc KPI → Trend → Detail; 5–8 visual; mỗi visual trả lời một câu.",
                        "(3) Một nguồn sự thật: mọi số từ measure (CUBEVALUE hoặc Pivot); KPI nào cũng có mốc so sánh và RAG có biểu tượng.",
                        "(4) Insight theo SCR, có drill-down kiểm chứng và đề xuất cho đúng người.",
                        "(5) Dự báo = mẫu hình quá khứ + giả định: luôn kèm dải tin cậy."], size=15, space=9)
box(s, 6.97, 1.65, 5.77, 3.8, fill=ALT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.03, name="Insight box")
tb(s, 7.17, 1.8, 5.37, 0.45, "Insight của bạn (3 câu SCR)", size=18, color=NAVY, bold=True)
tb(s, 7.17, 2.35, 5.37, 2.5, ["S: ___ (con số có mốc so sánh)", "C: ___ (nguyên nhân đã kiểm chứng)", "R: ___ (đề xuất cho Sales Director)"], size=18, space=14)
for k, t in enumerate(["Con số", "Nguyên nhân", "Hành động"]):
    chip(s, 7.17 + k * 1.84, 4.8, 1.69, t, fill=WHITE, line=HAIR, name=f"Tag {t}", h=0.45)
tb(s, 6.97, 5.6, 5.77, 0.9, "Chọn tháng bạn thích. Câu “vậy thì sao?” là câu hỏi chấm điểm C6.", size=14, color=INK2)
notes(s, 40)

# 41 · Giao capstone & bài về nhà
s = content("Giao capstone & bài về nhà", R8)
tb(s, L, 1.65, 7.0, 0.9, "Buổi 7: ôn tập, hoàn thiện dự án, thuyết trình 5–7 phút, nhận xét theo rubric.", size=17, color=NAVY, bold=True)
tb(s, L, 2.65, 7.0, 4.0, [[("Capstone (cá nhân, làm ở nhà): ", {"b": True, "color": NAVY}), "brief dự án + dataset mới + rubric bản học viên (../B7_Capstone/Rubric_B7_Capstone.md). Chuỗi Brief → Power Query → Data Model → KPI measure → Dashboard → Insight."],
                          [("Bài tập về nhà B6 (bắt buộc): ", {"b": True, "color": NAVY}), "dashboard một trang cho Marketing Director (Active Customers, phân khúc) trên file khởi đầu; tập dượt C1 · C5 · C6. Nộp trước Buổi 7 tại hocvien.mcivietnam.com."],
                          "Đọc handout B6 · tick “Checklist trước Buổi 7”"], size=14, space=10)
bricks(s, "B6 · Dashboard", "B7 · Capstone")
tb(s, 9.07, 5.7, 3.67, 0.9, "C5 Dashboard và C2 Dữ liệu là điều kiện bắt buộc để đạt Level 3 — Builder.", size=13, color=INK2)
notes(s, 41)

assert len(prs.slides) == 41, len(prs.slides)
prs.save(OUT)
print("ok", OUT, len(prs.slides), "slide;", f"còn {len(MISSING)} khung [CHỤP] chưa có ảnh:" if MISSING else "đủ ảnh.", " ".join(MISSING))
