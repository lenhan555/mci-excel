"""Dựng handout tự học Buổi 5 (Data Model và DAX): Handout/B5_Handout.docx bằng python-docx.

Số liệu lấy từ build_b5_meta.json (build_b5_meta.py). Ảnh chụp thật lấy từ Materials/B5/Images (H<nn>_…png);
chưa có ảnh thì để khung [CHỤP]. Xuất PDF bằng Microsoft Word (xem 2_Dataset/README.md).
    python3 _generator/build_docs_b5.py
"""
import json, os
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.image.image import Image as DocxImage

ROOT = os.path.normpath(os.environ.get("MCI_ROOT") or os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
GEN = os.path.dirname(os.path.abspath(__file__))
OUT = f"{ROOT}/1_Sessions/B5_Data_Model_DAX/Handout"
IMG = f"{ROOT}/Materials/B5/Images"   # ảnh chụp thật, danh sách: _DANH_SACH_ANH.md
MISSING = []
os.makedirs(OUT, exist_ok=True)
M = json.load(open(f"{GEN}/build_b5_meta.json", encoding="utf-8"))
H = M["hw"]; I = M["info"]
RED, NAVY, GOLD, GREY = RGBColor(0xC0, 0, 0), RGBColor(0x20, 0x38, 0x64), RGBColor(0x8A, 0x65, 0x00), RGBColor(0x58, 0x58, 0x58)
vn = lambda x: f"{round(x):,}".replace(",", ".")
pct = lambda x: f"{x * 100:.1f}%".replace(".", ",")


def base_doc(footer="MCI – Magic Code Institute · Excel for Business Intelligence Analyst · Buổi 5"):
    d = Document()
    sec = d.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    for m in ("left_margin", "right_margin"): setattr(sec, m, Cm(2))
    sec.top_margin = sec.bottom_margin = Cm(1.8)
    st = d.styles["Normal"]; st.font.name = "Arial"; st.font.size = Pt(10.5)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    for name, size, color in (("Heading 1", 16, NAVY), ("Heading 2", 13, NAVY), ("Heading 3", 11.5, RED)):
        h = d.styles[name]; h.font.name = "Arial"; h.font.size = Pt(size); h.font.color.rgb = color; h.font.bold = True
        rf = h.element.rPr.rFonts
        for k in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
            if rf.get(qn(k)) is not None: del rf.attrib[qn(k)]
        for k in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"): rf.set(qn(k), "Arial")
    fp = sec.footer.paragraphs[0]
    fp.text = footer
    fp.runs[0].font.size = Pt(8); fp.runs[0].font.color.rgb = GREY
    return d


def para(d, text="", bold=False, color=None, size=None, italic=False):
    p = d.add_paragraph()
    if text:
        r = p.add_run(text); r.bold = bold; r.italic = italic
        if color: r.font.color.rgb = color
        if size: r.font.size = Pt(size)
    return p


def rich(d, parts):
    p = d.add_paragraph()
    for t, f in parts:
        r = p.add_run(t); r.bold = f.get("bold", False); r.italic = f.get("italic", False)
        if f.get("color"): r.font.color.rgb = f["color"]
        if f.get("mono"): r.font.name = "Consolas"; r.font.size = Pt(9.5)
    return p


def bullet(d, text):
    p = d.add_paragraph(style="List Bullet"); p.add_run(text); return p


def code(d, text):
    """Khối mã M: chữ Consolas trên nền xám nhạt."""
    for line in text.splitlines():
        p = d.add_paragraph(); p.paragraph_format.space_after = Pt(0); p.paragraph_format.left_indent = Cm(0.5)
        r = p.add_run(line); r.font.name = "Consolas"; r.font.size = Pt(9)
        pPr = p._p.get_or_add_pPr(); shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), "F3F5F9"); pPr.append(shd)
    d.add_paragraph().paragraph_format.space_after = Pt(0)


def shade(cell, hex_):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd"); shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), hex_)
    tcPr.append(shd)


def table(d, rows, widths, header=True, size=9.5, first_bold=False):
    t = d.add_table(rows=len(rows), cols=len(rows[0]))
    t.style = "Table Grid"; t.alignment = WD_TABLE_ALIGNMENT.CENTER; t.autofit = False
    for j, w in enumerate(widths): t.columns[j].width = Cm(w)
    for i, row in enumerate(rows):
        for j, v in enumerate(row):
            c = t.cell(i, j); c.width = Cm(widths[j]); c.text = ""
            r = c.paragraphs[0].add_run(str(v)); r.font.size = Pt(size)
            if i < len(rows) - 1: c.paragraphs[0].paragraph_format.keep_with_next = True   # không tách bảng qua trang
            if header and i == 0:
                r.bold = True; r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF); shade(c, "203864")
            elif first_bold and j == 0:
                r.bold = True
    d.add_paragraph()
    return t


def figure(d, what, where, img, width=15):
    """Chèn ảnh chụp thật từ Materials/B5/Images nếu đã có; chưa có thì để khung chờ [CHỤP] (CLAUDE.md, mục Ảnh minh họa)."""
    p = d.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if os.path.exists(f"{IMG}/{img}"):
        px = DocxImage.from_file(f"{IMG}/{img}")
        w = min(width, px.px_width * 2.54 / 96)
        if w * px.px_height / px.px_width > 10: w = 10 * px.px_width / px.px_height
        p.add_run().add_picture(f"{IMG}/{img}", width=Cm(w)); return p
    MISSING.append(img)
    r = p.add_run(f"[CHỤP] {what}"); r.italic = True; r.font.color.rgb = GREY; r.font.size = Pt(9)
    r = p.add_run(f"\n{where} → {img}"); r.font.color.rgb = GREY; r.font.size = Pt(8)
    pPr = p._p.get_or_add_pPr(); bdr = OxmlElement("w:pBdr")
    for side in ("top", "left", "bottom", "right"):
        e = OxmlElement(f"w:{side}"); e.set(qn("w:val"), "dashed"); e.set(qn("w:sz"), "6"); e.set(qn("w:color"), "A6A6A6"); e.set(qn("w:space"), "8"); bdr.append(e)
    pPr.append(bdr)
    return p


def tool(d, name, purpose, how=None, example=None, tip=None, protip=None, shot=None, m=None):
    """Khối công cụ: tên (đỏ) → mục đích → cách dùng → bảng ví dụ trên case → mã M → ảnh → TIP / PRO TIP."""
    d.add_heading(name, level=3).paragraph_format.keep_with_next = True
    para(d, purpose).paragraph_format.keep_with_next = True
    for h in how or []: bullet(d, h)
    if example: table(d, example[0], example[1])
    if m: code(d, m)
    if shot: figure(d, *shot)
    if tip: rich(d, [("TIP: ", {"bold": True, "color": NAVY}), (tip, {})])
    if protip: rich(d, [("PRO TIP: ", {"bold": True, "color": GOLD}), (protip, {})])



pm = lambda k: pct(I[k])          # % đạt target tháng 1..4
STAT = {1: "Đạt", 2: "Cảnh báo", 3: "Gần đạt", 4: "Đạt"}
RAG = M["rag"]
n90 = {m: len(RAG[m]["duoi_80"]) + len(RAG[m]["80_90"]) for m in "1234"}
n80 = {m: len(RAG[m]["duoi_80"]) for m in "1234"}

# =====================================================================
d = base_doc()
para(d, "MCI – MAGIC CODE INSTITUTE", bold=True, color=GREY, size=9)
d.add_heading("Handout Buổi 5: Data Model và DAX", level=1)
para(d, "Tài liệu tự học đi kèm Buổi 5. Phần Tóm tắt nhắc lại mô hình và 14 measure đã dựng trên lớp; mục 1 mở rộng phần Hide from Client Tools đã làm ở Slide 20, "
        "mục 2–8 là công cụ KHÔNG giảng chi tiết trên lớp. Đọc sau buổi học rồi làm bài tập về nhà (bắt buộc) ở mục 9. "
        "Ví dụ làm trên BẢN SAO của file đáp án Practice/B5_MCI_Mart_Answer.xlsx (file đáp án là file khởi đầu Buổi 6, giữ nguyên bản gốc). "
        "Tên lệnh giữ tiếng Anh như trên Ribbon. Số trong ảnh chụp theo định dạng en-US: dấu phẩy ngăn nghìn, dấu chấm thập phân "
        "(425,850,000 trong ảnh là 425.850.000 trong bảng).", italic=True, color=GREY)

d.add_heading("Mục lục", level=2)
for s in ["Tóm tắt: mô hình ngôi sao và 14 measure", "1. Ẩn cột khỏi Client Tools, bỏ cột thừa, giảm granularity", "2. Hierarchy trên Calendar: Nam → Quy → Thang → Date",
          "3. Lookup-to-lookup: Customer[Home_Store_Code] → Store, quan hệ active / inactive, USERELATIONSHIP",
          "4. Bad cardinality: vì sao không nối thẳng Sales với Target", "5. PRO TIP iterator: SUMX thay cột COGS", "6. PRO TIP SWITCH(TRUE(), …): nhãn trạng thái target",
          "7. Hàm văn bản trong DAX: FORMAT, LEFT, &", "8. Disconnected slicer: bảng ngưỡng 80% / 90% / 100%", "9. Bài tập về nhà (bắt buộc) và checklist trước Buổi 6"]:
    bullet(d, s)

# Tóm tắt
d.add_heading("Tóm tắt: mô hình ngôi sao và 14 measure", level=2)
para(d, f"Data Model của buổi gồm 6 bảng nạp từ Buổi 4: hai bảng fact (Sales {vn(I['sales_rows'])} dòng, Target {I['target_rows']} dòng) và bốn bảng lookup "
        f"(Product {I['product']}, Store {I['store']}, Customer {vn(I['customer'])}, Calendar {I['calendar_rows']}). Sales nối đủ bốn lookup; Target nối Store và Calendar. "
        "Quy tắc dùng Pivot: hàng, cột, bộ lọc, slicer luôn lấy cột của bảng lookup; bảng fact chỉ cung cấp giá trị (measure).")
table(d, [["Measure", "Công thức DAX", "Ý nghĩa"],
          ["Total Revenue", "SUM(Sales[Revenue])", "Measure cơ sở"],
          ["Total COGS", "SUMX(Sales, Sales[Quantity] * RELATED(Product[Unit_Cost]))", "Giá vốn (mục 5)"],
          ["Gross Margin %", "DIVIDE([Total Revenue] - [Total COGS], [Total Revenue])", "Biên lợi nhuận gộp"],
          ["Revenue All Stores", "CALCULATE([Total Revenue], ALL(Store))", "Mẫu số tỷ trọng"],
          ["% Revenue Share", "DIVIDE([Total Revenue], [Revenue All Stores])", "Tỷ trọng doanh thu"],
          ["Total Target", "SUM(Target[Target])", "Target (grain tháng × cửa hàng)"],
          ["Revenue vs Target %", "DIVIDE([Total Revenue], [Total Target])", "% đạt target"],
          ["Variance", "IF(ISBLANK([Total Revenue]), BLANK(), [Total Revenue] - [Total Target])", "Chênh lệch so với target"],
          ["Revenue YTD", "CALCULATE([Total Revenue], DATESYTD(Calendar[Date]))", "Lũy kế từ đầu năm"],
          ["Revenue LY", "CALCULATE([Total Revenue], DATEADD(Calendar[Date], -1, YEAR))", "Cùng kỳ năm trước"],
          ["Revenue YoY %", "IF(ISBLANK([Total Revenue]), BLANK(), DIVIDE([Total Revenue] - [Revenue LY], [Revenue LY]))", "Tăng trưởng so với cùng kỳ"],
          ["Total Orders", "DISTINCTCOUNT(Sales[Order_ID])", "Số đơn (một đơn nhiều dòng)"],
          ["Avg Order Value", "DIVIDE([Total Revenue], [Total Orders])", "Giá trị đơn trung bình"],
          ["Active Customers", "DISTINCTCOUNT(Sales[Customer_ID])", "Số khách có mua"]],
      [3.6, 9.2, 3.2], size=8.5, first_bold=True)
para(d, "Lưu ý Total COGS: file đáp án viết bằng SUMX (mục 5); trên lớp, Total COGS = SUM(Sales[COGS]) trên cột COGS đã tạo ở Slide 15. Hai cách cùng ra 619.663.000.", size=9)
para(d, f"Số tự đối chiếu: Quý 1/2025 doanh thu {vn(I['q1_rev'])}, {vn(I['q1_orders'])} đơn, AOV {vn(I['q1_aov'])}, GM {pct(I['q1_gm'])}; "
        f"% đạt target tháng 1–4: {pm('t1_pct')} · {pm('t2_pct')} · {pm('t3_pct')} · {pm('t4_pct')}; YTD đến tháng 4 {vn(I['ytd_t4'])}, YoY {pct(I['ytd_yoy_t4'])}.",
     italic=True, color=GREY)

# 1 Hide
d.add_heading("1. Ẩn cột khỏi Client Tools, bỏ cột thừa, giảm granularity", level=2)
para(d, "Slide 20 đã ẩn 9 cột khóa và cột ngày của Sales và Target; mục này nhắc lại cách làm và thêm hai việc chưa làm trên lớp: bỏ cột thừa, giảm granularity. "
        "Pivot liệt kê mọi cột của mọi bảng. Cột khóa và cột ngày của bảng fact (Store_Code, Order_Date…) rất dễ bị kéo nhầm vào hàng hoặc cột: "
        "Pivot vẫn chạy nhưng cho số sai (Target lặp lại, YTD / YoY trống). Giấu chúng đi để chỉ còn cột của bảng lookup trong danh sách.")
tool(d, "Hide from Client Tools", "Giấu cột (hoặc cả bảng) khỏi PivotTable Fields. Cột vẫn còn trong mô hình: relationship và công thức DAX vẫn dùng được.",
     how=["Cửa sổ Power Pivot → Data View → chọn bảng Sales → giữ Ctrl chọn tiêu đề các cột Nam, Thang, Order_Date, Store_Code, Product_Code, Customer_ID → chuột phải → Hide from Client Tools.",
          "Làm tương tự bảng Target: Store_Code, Thang, Ngay_Dau_Thang.",
          "Cột đã ẩn chuyển sang màu xám. Muốn hiện lại: chuột phải → Unhide from Client Tools.",
          "Giữ Order_ID (và Line_No) hiển thị hoặc ẩn đều được, nhưng không xóa: Total Orders đếm Order_ID."],
     shot=("Data View bảng Sales: các cột khóa / ngày màu xám, menu chuột phải có Unhide from Client Tools",
           "Practice/B5_MCI_Mart_Answer.xlsx · Power Pivot → Data View → Sales", "H01_hide_client_tools.png"),
     tip="Sau khi ẩn, mở lại Pivot: PivotTable Fields không còn Sales[Thang] nữa, chỉ còn Calendar[Thang]. Lỗi \"Target lặp lại mọi tháng\" từ đó không tái diễn.")
tool(d, "Bỏ cột thừa và giảm granularity", "Mỗi cột và mỗi dòng nạp vào Data Model đều tốn bộ nhớ và làm file chậm hơn.",
     how=["Cột không dùng trong Pivot, relationship hay measure thì bỏ ngay ở Power Query (Remove Columns) trước khi nạp, đừng chỉ ẩn.",
          "Cột có nhiều giá trị khác nhau (ngày giờ, mã đơn) nén kém nhất; ngày giờ thì tách phần ngày, bỏ phần giờ nếu không dùng.",
          "Granularity (độ chi tiết) của bảng fact chỉ nên chi tiết đến mức câu hỏi cần. Nếu mọi báo cáo chỉ xem theo ngày × cửa hàng × sản phẩm thì gộp dòng bằng Group By ở Power Query là đủ. "
          "Sales của lớp giữ mức dòng hàng của đơn vì cần đếm đơn và khách."])

# 2 Hierarchy
d.add_heading("2. Hierarchy trên Calendar: Nam → Quy → Thang → Date", level=2)
para(d, "Hierarchy gom các cột của một bảng thành một cây để kéo một lần vào hàng của Pivot, rồi mở rộng / thu gọn bằng nút +/−. Thứ tự các mức phải từ lớn đến nhỏ.")
tool(d, "Create Hierarchy", "Tạo cây thời gian trên bảng Calendar.",
     how=["Power Pivot → Diagram View → chuột phải tiêu đề bảng Calendar → Create Hierarchy (hoặc nút ở góc phải tiêu đề bảng). Đặt tên Cay_Thoi_Gian.",
          "Kéo lần lượt vào hierarchy: Nam, Quy, Thang, Date (thứ tự này).",
          "Trong Pivot (From Data Model), kéo Cay_Thoi_Gian vào Rows: bấm ⊞ để mở Năm → Quý → Tháng → Ngày."],
     shot=("Diagram View: bảng Calendar có hierarchy Cay_Thoi_Gian (Nam, Quy, Thang, Date); bên phải Pivot Rows = Cay_Thoi_Gian, Values = Total Revenue",
           "Bản sao Practice/B5_MCI_Mart_Answer.xlsx · Diagram View (Create Hierarchy) và sheet PV_Hier (tự tạo)", "H02_hierarchy.png"),
     tip=f"Pivot trong ảnh: 2024 = 3.198.430.500; 2025 mở ra Quý 1 gồm tháng 1 = 425.850.000 · tháng 2 = 245.996.000 · tháng 3 = 269.608.000; Quý 2 chỉ có tháng 4 = 296.965.000. "
         f"Grand Total = {vn(M['source']['revenue'])}.",
     protip="Mỗi mức của hierarchy vẫn là cột riêng của Calendar: có thể kéo Quy hoặc Thang riêng ra Columns. Hierarchy chỉ tiện khi cần drill từ năm xuống ngày.")

# 3 Lookup to lookup
d.add_heading("3. Lookup-to-lookup: Customer[Home_Store_Code] → Store, active / inactive", level=2)
para(d, "Có khi một bảng lookup cần nối tiếp sang bảng lookup khác. Ví dụ Customer có cột Home_Store_Code (cửa hàng gốc của khách) trỏ về Store[Store_Code]. "
        "Quan hệ này cho phép hỏi \"khách của cửa hàng nào\", khác với \"đơn mua tại cửa hàng nào\" (Sales[Store_Code] → Store).")
tool(d, "Quan hệ active và inactive", "Giữa hai bảng chỉ có tối đa một đường lọc hoạt động (active). Quan hệ thêm vào được đánh dấu inactive: nét đứt trong Diagram View.",
     how=["Tạo quan hệ như thường lệ: Design → Create Relationship (hoặc kéo cột sang bảng kia trong Diagram View). Cột bên Store (một) là khóa duy nhất; bên Customer (nhiều) trùng lặp.",
          "Store đã có đường lọc tới Sales (trực tiếp); thêm một đường nữa qua Customer sẽ cho hai đường giữa Store và Sales, mơ hồ. Vì vậy quan hệ Customer[Home_Store_Code] → Store ở trạng thái inactive.",
          "Bỏ tick Active trong hộp Create Relationship nếu muốn chủ động đặt inactive ngay từ đầu."],
     shot=("Diagram View: quan hệ Customer[Home_Store_Code] → Store[Store_Code] hiện nét đứt (inactive)",
           "Bản sao Practice/B5_MCI_Mart_Answer.xlsx · Diagram View", "H03_lookup_to_lookup.png"))
tool(d, "USERELATIONSHIP", "Bật một quan hệ inactive cho riêng một measure, không đổi mô hình.",
     m='Khach Cua Hang Goc = CALCULATE(\n    DISTINCTCOUNT(Customer[Customer_ID]),\n    USERELATIONSHIP(Customer[Home_Store_Code], Store[Store_Code])\n)',
     how=["Đặt Store[Store_Name] vào hàng của Pivot: measure đếm số khách có Home_Store_Code là cửa hàng đó. Không có USERELATIONSHIP, quan hệ inactive bị bỏ qua."],
     tip="Muốn mô hình vẫn hiểu theo cách \"mua tại cửa hàng\" thì giữ Sales → Store là quan hệ active. Chỉ bật quan hệ thứ hai khi cần, và chỉ trong measure đó.")

# 4 Bad cardinality
d.add_heading("4. Bad cardinality: vì sao không nối thẳng Sales với Target", level=2)
para(d, "Data Model chỉ nhận quan hệ một–nhiều (1–*): một đầu là khóa duy nhất (bảng lookup), đầu kia lặp lại (bảng fact). "
        "Cả Sales[Store_Code] và Target[Store_Code] đều lặp lại (mỗi cửa hàng có nhiều dòng bán hàng và 12 dòng target) nên không nối được với nhau, đó là quan hệ nhiều–nhiều (many-to-many).")
tool(d, "Thử nối thẳng Sales[Store_Code] với Target[Store_Code]", "Chạy thử để nhận ra thông báo và biết hướng sửa.",
     how=["Design → Create Relationship → bảng trên: Sales, cột Store_Code; bảng dưới: Target, cột Store_Code. Hộp thoại hiện cảnh báo vàng và nút OK bị khóa (xám).",
          "Cách sửa: không nối hai bảng fact với nhau. Nối từng bảng fact vào bảng lookup chung (Store, Calendar). Store và Calendar lúc đó lọc cả Sales lẫn Target, và measure Revenue vs Target % đặt hai fact cạnh nhau theo cửa hàng, theo tháng."],
     shot=("Hộp Create Relationship Sales[Store_Code] – Target[Store_Code] với cảnh báo many-to-many, nút OK bị khóa",
           "Bản sao Practice/B5_MCI_Mart_Answer.xlsx · Design → Create Relationship", "H04_bad_cardinality.png"),
     tip="Lỗi tương tự gặp khi lookup của bạn còn khóa trùng (ví dụ Product còn mã P015 xuất hiện hai lần ở Buổi 4): Remove Duplicates ở Power Query rồi mới nối.")

# 5 SUMX
d.add_heading("5. PRO TIP iterator: SUMX thay cột COGS", level=2)
para(d, "Trên lớp tính COGS bằng calculated column rồi SUM cột đó. Iterator (SUMX, AVERAGEX, MAXX…) cho cùng kết quả mà không thêm cột vào bảng: "
        "duyệt từng dòng của bảng, tính một biểu thức, rồi cộng lại.")
tool(d, "SUMX(Sales, Sales[Quantity] * RELATED(Product[Unit_Cost]))", "Với mỗi dòng của Sales: lấy Quantity nhân giá vốn của sản phẩm đó (RELATED đi từ Sales sang Product theo relationship), rồi cộng mọi dòng.",
     m='Total COGS = SUMX(\n    Sales,\n    Sales[Quantity] * RELATED(Product[Unit_Cost])\n)',
     example=([["Cách làm", "Công thức", "Ưu / nhược"],
               ["Calculated column + SUM", "COGS = Sales[Quantity] * RELATED(Product[Unit_Cost]); Total COGS = SUM(Sales[COGS])", "Dễ hiểu, xem được từng dòng; tốn bộ nhớ, thêm một cột"],
               ["Iterator SUMX", "SUMX(Sales, Sales[Quantity] * RELATED(Product[Unit_Cost]))", "Không thêm cột, nhẹ hơn; không xem được giá trị từng dòng"]], [4.0, 7.0, 5.0]),
     tip=f"Hai cách phải ra cùng số. Kiểm tra: Total COGS Quý 1/2025 = {vn(M['q1']['cogs'])}, GM {pct(I['q1_gm'])}.",
     protip="Cột RELATED chỉ dùng được khi Sales đã nối Product. RELATED lấy một giá trị từ bảng phía một; muốn lấy từ nhiều dòng của bảng phía nhiều thì dùng RELATEDTABLE.")

# 6 SWITCH
d.add_heading("6. PRO TIP SWITCH(TRUE(), …): nhãn trạng thái target", level=2)
para(d, "Nhiều điều kiện lồng nhau bằng IF rất khó đọc. SWITCH(TRUE(), điều kiện 1, kết quả 1, điều kiện 2, kết quả 2, …, kết quả còn lại) chạy từ trên xuống, "
        "trả kết quả của điều kiện ĐÚNG đầu tiên.")
tool(d, "Measure Trang thai Target", "Gán nhãn theo ngưỡng của KPI Dictionary: Warning 90%, Critical 80%.",
     m=('Trang thai Target =\nVAR p = [Revenue vs Target %]\nRETURN\n    SWITCH(\n        TRUE(),\n        ISBLANK(p), BLANK(),\n        p >= 1,    "Đạt",\n        p >= 0.9,  "Gần đạt",\n        p >= 0.8,  "Cảnh báo",\n        "Nguy cấp"\n    )'),
     example=([["Tháng 2025", "Revenue vs Target %", "Trang thai Target"]] + [[f"Tháng {k}", pm(f"t{k}_pct"), STAT[k]] for k in (1, 2, 3, 4)], [4.0, 5.0, 5.0]),
     tip="Thứ tự điều kiện quan trọng: viết điều kiện hẹp nhất (p >= 1) trước, vì SWITCH dừng ngay ở điều kiện đúng đầu tiên. Đặt p >= 0.8 lên đầu thì mọi tháng ≥ 80% đều ra \"Cảnh báo\".",
     protip="VAR p = … tính một lần, dùng nhiều lần: nhanh hơn và dễ đọc hơn việc viết [Revenue vs Target %] ở mỗi dòng. Nhãn là chữ nên không cộng dồn được; dùng cho Pivot, KPI card, không dùng làm giá trị tính tiếp.")

# 7 Text
d.add_heading("7. Hàm văn bản trong DAX: FORMAT, LEFT, &", level=2)
table(d, [["Hàm", "Dùng để", "Ví dụ"],
          ["&", "Nối chuỗi", "\"Doanh thu: \" & [Total Revenue]"],
          ["FORMAT(giá trị, \"định dạng\")", "Biến số hoặc ngày thành chuỗi theo định dạng", "FORMAT([Total Revenue] / 1000000, \"#,##0.0\") → 425.9 (máy en-US) hoặc 425,9 (máy vi-VN) cho tháng 1"],
          ["LEFT(chuỗi, n)", "Lấy n ký tự đầu", "LEFT(Calendar[Thu], 3) → 3 ký tự đầu của tên thứ"]],
      [4.2, 4.6, 7.2], first_bold=True, size=9)
tool(d, "Nhãn KPI bằng chữ", "Gộp số và đơn vị thành một chuỗi hiển thị trong ô KPI card.",
     m='Nhan Doanh Thu =\n"Doanh thu: " & FORMAT([Total Revenue] / 1000000, "#,##0.0") & " triệu"',
     tip="FORMAT trả về chữ: dùng để hiển thị, không dùng để tính tiếp. Dấu thập phân và ngăn nghìn của FORMAT theo locale của máy mở file, nên cùng một file hiển thị khác nhau trên máy en-US và vi-VN.")

# 8 Disconnected slicer
d.add_heading("8. Disconnected slicer: bảng ngưỡng target 80% / 90% / 100%", level=2)
para(d, "Slicer thông thường lọc theo relationship. Disconnected slicer đọc từ một bảng KHÔNG nối với mô hình: nó không lọc dữ liệu, chỉ cho người dùng chọn một giá trị, "
        "và measure đọc giá trị đó để tự tính. Ở đây người dùng chọn ngưỡng 80% / 90% / 100%, measure đếm số cửa hàng có % đạt target thấp hơn ngưỡng đã chọn.")
tool(d, "Dựng bảng ngưỡng, measure và slicer", "Làm trên bản sao của Practice/B5_MCI_Mart_Answer.xlsx, một sheet mới (ví dụ PV_Nguong).",
     how=["Bảng ngưỡng: trên sheet mới, ô G1 gõ Muc; G2:G4 gõ '80%, '90%, '100% (dấu nháy đơn đầu để Excel lưu thành chữ) → Ctrl+T tạo Table, đặt tên Muc.",
          "Đưa vào mô hình: Power Pivot → Add to Data Model. KHÔNG tạo relationship nào cho bảng Muc.",
          "Measure: Power Pivot → Measures → New Measure; Table name: Target; Measure name: Stores Duoi Nguong; dán công thức bên dưới.",
          "Pivot From Data Model: Filters = Calendar[Nam] chọn 2025; Rows = Calendar[Thang]; Values = Stores Duoi Nguong.",
          "Slicer: Insert → Slicer → bảng Muc, cột Muc. Chọn 90% trên slicer, Pivot đổi số theo."],
     m=('Stores Duoi Nguong =\nVAR t = IF(\n    HASONEVALUE(Muc[Muc]),\n    VALUE(SUBSTITUTE(VALUES(Muc[Muc]), "%", "")) / 100,\n    0.9\n)\nRETURN\n    IF(\n        ISBLANK([Total Revenue]), BLANK(),\n        COUNTROWS(\n            FILTER(\n                VALUES(Store[Store_Code]),\n                NOT ISBLANK([Total Target]) &&\n                DIVIDE([Total Revenue], [Total Target]) < t\n            )\n        ) + 0\n    )'),
     shot=("Slicer Nguong (80% · 90% · 100%, đang chọn 90%) bên trái; Pivot Filters Nam = 2025, Rows Thang 1–4, Values Stores Duoi Nguong; bảng Muc ở cột G",
           "Bản sao Practice/B5_MCI_Mart_Answer.xlsx · sheet PV_Nguong (dựng theo các bước trên)", "H05_disconnected_slicer.png"))
table(d, [["Ngưỡng chọn trên slicer", "Tháng 1", "Tháng 2", "Tháng 3", "Tháng 4"],
          ["80% (cửa hàng dưới 80% target)", n80["1"], n80["2"], n80["3"], n80["4"]],
          ["90% (cửa hàng dưới 90% target)", n90["1"], n90["2"], n90["3"], n90["4"]]], [6.0, 2.5, 2.5, 2.5, 2.5], first_bold=True)
para(d, "Cửa hàng dưới 80%: tháng 2: " + ", ".join(RAG["2"]["duoi_80"]) + "; tháng 3: " + ", ".join(RAG["3"]["duoi_80"]) + "; tháng 4: " + ", ".join(RAG["4"]["duoi_80"]) +
     ". Ở ngưỡng 90% thêm các cửa hàng nằm giữa 80% và 90%.", italic=True, color=GREY)
rich(d, [("Đọc dòng Grand Total: ", {"bold": True, "color": NAVY}),
         ("ở ngưỡng 90% dòng Grand Total ra 12, KHÔNG phải tổng 0 + 8 + 6 + 4. Dòng này so doanh thu tháng 1–4 với target cả năm 2025 "
          f"({vn(I['target_total'])}) nên cả 12 cửa hàng đều dưới ngưỡng. Với measure kiểu này, đọc từng tháng; đừng dùng Grand Total làm kết luận.", {})])
rich(d, [("TIP: ", {"bold": True, "color": NAVY}),
         ("Chưa chọn gì trên slicer (hoặc chọn nhiều ngưỡng) thì HASONEVALUE ra FALSE và measure dùng ngưỡng mặc định 90%. "
          "Tháng chưa có doanh thu (tháng 5–12) cho kết quả trống nhờ ISBLANK([Total Revenue]); \"+ 0\" ở cuối giúp tháng 1 hiện số 0 thay vì trống.", {})])
rich(d, [("PRO TIP: ", {"bold": True, "color": GOLD}),
         ("Cùng kỹ thuật dùng cho what-if: bảng tham số (ví dụ % tăng giá 0%, 5%, 10%) không nối mô hình, measure đọc giá trị được chọn để giả lập kịch bản.", {})])

# 9 Homework
d.add_heading("9. Bài tập về nhà (bắt buộc) và checklist trước Buổi 6", level=2)
rich(d, [("Tình huống: ", {"bold": True}), ("Phòng Kinh doanh gửi Sales Director báo cáo tháng 1–4/2025; Phòng Tài chính báo số khác. Bạn là BI analyst được giao kiểm tra số nào đúng, "
                                              "vì sao lệch và sửa để hai phòng dùng chung một định nghĩa.", {})])
rich(d, [("File đề: ", {"bold": True}), ("Homework/B5_BTVN_MCI_Mart_KPI_T1_T4_2025.xlsx. Làm ngay trên file đề, nộp trước Buổi 6, tên file B5_BTVN_<HoTen>.xlsx. Sheet De_bai ghi đủ yêu cầu.", {})])
rich(d, [("Nộp bài: ", {"bold": True, "color": NAVY}), ("https://hocvien.mcivietnam.com/auth#/ (tài khoản như Buổi 4).", {})])
for s in ["Đối soát: so BaoCao_KinhDoanh với So_lieu_Tai_chinh; ghi Nhat_ky_doi_soat đủ 4 lỗi gốc, mỗi lỗi một dòng: chỉ tiêu, số hai bên, nguyên nhân, cách sửa.",
          "Sửa mô hình thành hình ngôi sao: nối Target vào Calendar, ẩn cột kỹ thuật của hai bảng fact (mục 1), Mark as Date Table; dán ảnh Diagram View vào Brief_Insight.",
          "Viết lại KPI bằng explicit measure theo tầng (measure cơ sở → measure dẫn xuất); thêm like-for-like: Revenue LFL, Revenue LFL LY, YoY LFL %; xóa 6 measure KD.",
          "Dựng lại BaoCao_KinhDoanh từ measure của bạn (tháng lấy từ Calendar[Thang]); tạo sheet BaoCao_LFL; đối soát lại với So_lieu_Tai_chinh, chênh lệch phải bằng 0.",
          "KPI_Dictionary: một dòng cho mỗi measure, đủ DAX, grain, dimensions, owner, target và giới hạn của chỉ số.",
          "Brief_Insight: 3 câu hỏi đo lường được cho Sales Director (chỉ số, chiều, thời gian, mốc so sánh), mỗi câu trả lời bằng một Pivot từ measure; 1 câu insight (con số, nguyên nhân khả dĩ, hàm ý)."]:
    bullet(d, s)
rich(d, [("Tự kiểm tra trước khi nộp: ", {"bold": True, "color": NAVY}),
         (f"tháng 4 doanh thu {vn(H['months']['4']['tc_rev'])}, {vn(H['months']['4']['tc_orders'])} đơn, AOV {vn(H['months']['4']['tc_aov'])}, GM {pct(H['months']['4']['tc_gm'])}, đạt {pct(H['months']['4']['tc_vs_target'])} target; "
          f"YoY T1–T4 cả chuỗi {pct(H['ytd_yoy_t4'])}, like-for-like {pct(H['lfl_ytd_yoy'])}.", {})])
h = d.add_heading("Checklist trước Buổi 6", level=3); h.paragraph_format.keep_with_next = True
for s in [f"Mô hình 6 bảng, 6 relationship hình ngôi sao: Sales nối Product, Store, Customer, Calendar; Target nối Store và Calendar.",
          "Calendar đã Mark as Date Table; cột khóa và cột ngày của Sales và Target đã Hide from Client Tools.",
          f"14 measure chạy đúng: Quý 1/2025 doanh thu {vn(I['q1_rev'])}, GM {pct(I['q1_gm'])}; % đạt target tháng 1–4 {pm('t1_pct')} · {pm('t2_pct')} · {pm('t3_pct')} · {pm('t4_pct')}.",
          "Pivot đúng khi đổi bộ lọc: hàng / cột / slicer luôn lấy cột của bảng lookup.",
          "Đã thử các mục 1–8 trên bản sao file đáp án (hierarchy, quan hệ inactive, SUMX, SWITCH, disconnected slicer).",
          "Đã làm và nộp bài tập về nhà (B5_BTVN_<HoTen>.xlsx).",
          "Nếu file chưa đạt: Buổi 6 lớp phát file khởi đầu (đáp án B5, Practice/B5_MCI_Mart_Answer.xlsx)."]:
    bullet(d, "☐  " + s).paragraph_format.keep_with_next = True
d.paragraphs[-1].paragraph_format.keep_with_next = False
d.save(f"{OUT}/B5_Handout.docx")
print("ok;", f"còn {len(MISSING)} khung [CHỤP] chưa có ảnh:" if MISSING else "đủ ảnh.", " ".join(MISSING))
