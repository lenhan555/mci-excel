"""Build B1 handout and machine checklist (.docx) with python-docx."""
import os
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

ROOT = os.path.normpath(os.environ.get("MCI_ROOT") or os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
OUT = f"{ROOT}/1_Sessions/B1_Nen_tang_Excel_Du_lieu/Handout"
IMG = f"{ROOT}/Materials/B1/Images"
RED, NAVY, GOLD, GREY = RGBColor(0xC0, 0, 0), RGBColor(0x20, 0x38, 0x64), RGBColor(0x8A, 0x65, 0x00), RGBColor(0x58, 0x58, 0x58)

def base_doc():
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
    # footer
    fp = sec.footer.paragraphs[0]
    fp.text = "MCI – Magic Code Institute · Excel for Business Intelligence Analyst · Buổi 1"
    fp.runs[0].font.size = Pt(8); fp.runs[0].font.color.rgb = GREY
    return d

def para(d, text="", bold=False, color=None, size=None, italic=False, style=None, align=None):
    p = d.add_paragraph(style=style)
    if text:
        r = p.add_run(text); r.bold = bold; r.italic = italic
        if color: r.font.color.rgb = color
        if size: r.font.size = Pt(size)
    if align: p.alignment = align
    return p

def rich(d, parts, style=None):
    """parts: list of (text, dict(bold, color, italic, code))"""
    p = d.add_paragraph(style=style)
    for t, f in parts:
        r = p.add_run(t)
        r.bold = f.get("bold", False); r.italic = f.get("italic", False)
        if f.get("color"): r.font.color.rgb = f["color"]
        if f.get("code"): r.font.name = "Consolas"
    return p

def bullet(d, text, level=0):
    p = d.add_paragraph(style="List Bullet" if level == 0 else "List Bullet 2")
    p.add_run(text)
    return p

def shade(cell, hex_):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd"); shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), hex_)
    tcPr.append(shd)

def table(d, rows, widths, header=True):
    t = d.add_table(rows=len(rows), cols=len(rows[0]))
    t.style = "Table Grid"; t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, row in enumerate(rows):
        for j, v in enumerate(row):
            c = t.cell(i, j); c.width = Cm(widths[j])
            c.text = ""
            r = c.paragraphs[0].add_run(str(v)); r.font.size = Pt(9.5)
            if header and i == 0:
                r.bold = True; r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF); shade(c, "203864")
    d.add_paragraph()
    return t

def figure(d, image, img=None, width=None, height=None):
    """Chèn ảnh từ Materials/B1/Images nếu đã có; chưa có thì giữ khung placeholder."""
    p = d.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if img and os.path.exists(f"{IMG}/{img}"):
        p.add_run().add_picture(f"{IMG}/{img}", width=Cm(width) if width else None, height=Cm(height) if height else None)
        return p
    r = p.add_run(f"[ẢNH MINH HỌA — cần chụp: {image}]"); r.italic = True; r.font.color.rgb = GREY; r.font.size = Pt(9)
    pPr = p._p.get_or_add_pPr(); bdr = OxmlElement("w:pBdr")
    for side in ("top", "left", "bottom", "right"):
        e = OxmlElement(f"w:{side}"); e.set(qn("w:val"), "dashed"); e.set(qn("w:sz"), "6"); e.set(qn("w:color"), "A6A6A6"); e.set(qn("w:space"), "8"); bdr.append(e)
    pPr.append(bdr)
    return p

def tool(d, name, purpose, how=None, tip=None, protip=None, image=None, img=None, width=None, height=None):
    """Tool block following the course style: name (red) → purpose → image → TIP/PRO TIP."""
    d.add_heading(name, level=3)
    para(d, purpose)
    if how:
        for h in how: bullet(d, h)
    if image:
        figure(d, image, img, width, height)
    if tip:
        rich(d, [("TIP: ", {"bold": True, "color": NAVY}), (tip, {})])
    if protip:
        rich(d, [("PRO TIP: ", {"bold": True, "color": GOLD}), (protip, {})])

# =====================================================================
# HANDOUT B1
# =====================================================================
d = base_doc()
para(d, "MCI – MAGIC CODE INSTITUTE", bold=True, color=GREY, size=9)
d.add_heading("Handout Buổi 1 — Excel nền tảng cho người làm BI", level=1)
para(d, "Tài liệu tự học đi kèm Buổi 1. Nội dung ở đây KHÔNG giảng chi tiết trên lớp; hãy đọc sau buổi học và làm bài tập về nhà (bắt buộc) ở mục 10. "
        "Mọi ví dụ dùng file case MCI Mart (B1_MCI_Mart_Q1_2025_Answer.xlsx).", italic=True, color=GREY)
para(d, "Lưu ý dấu phân cách: máy cài định dạng vùng Việt Nam thường dùng dấu chấm phẩy ( ; ) giữa các đối số, ví dụ =SUMIFS(C:C; A:A; \"Bắc\"). "
        "Tài liệu này viết theo dấu phẩy ( , ) như tài liệu tiếng Anh — hãy đổi nếu máy bạn dùng ( ; ).", color=RED)

d.add_heading("Mục lục", level=2)
for s in ["1. Các tab Ribbon — dùng tab nào cho việc gì", "2. Phím tắt Windows và tương đương trên Mac",
          "3. Quick Access Toolbar", "4. Freeze Panes", "5. Tắt GETPIVOTDATA (chuẩn bị cho Buổi 3)",
          "6. Mười hàm dùng nhiều nhất trong BI", "7. Flash Fill", "8. Fill Series nâng cao",
          "9. Tham chiếu sang workbook khác", "10. Bài tập về nhà (bắt buộc)", "11. Checklist trước Buổi 2"]:
    bullet(d, s)

# 1 Ribbon
d.add_heading("1. Các tab Ribbon — dùng tab nào cho việc gì", level=2)
para(d, "Không cần thuộc từng nút. Hãy nhớ quy tắc: \"cần làm việc X thì vào tab Y\". Cột cuối cho biết khóa học dùng tab đó nhiều nhất ở buổi nào.")
table(d, [
    ["Tab", "Nhóm lệnh chính", "Việc thường làm trong BI", "Buổi"],
    ["File", "Save As, Options, Account, Export", "Bật add-in Power Pivot; kiểm tra phiên bản; xuất PDF", "B1, B6"],
    ["Home", "Clipboard, Font, Number, Conditional Formatting, Format as Table, Find & Select", "Định dạng số, tô màu theo điều kiện, Go To Special", "Mọi buổi"],
    ["Insert", "Table, PivotTable, Charts, Sparklines, Slicer, Timeline", "Tạo Table, Pivot, biểu đồ, slicer", "B1, B3, B6"],
    ["Page Layout", "Themes, Page Setup, Print Area, Scale to Fit", "Chuẩn bị dashboard để in/xuất PDF", "B6"],
    ["Formulas", "Function Library, Name Manager, Formula Auditing, Calculation Options", "Viết hàm, đặt tên vùng, dò lỗi, Watch Window", "B1, B2"],
    ["Data", "Get & Transform Data, Queries & Connections, Sort & Filter, Data Tools", "Data Validation (B1), Refresh All (B3), Power Query (B4), Remove Duplicates, Text to Columns", "B1, B3, B4"],
    ["Review", "Comments, Protect Sheet, Protect Workbook", "Khóa dashboard trước khi gửi", "B6"],
    ["View", "Freeze Panes, Gridlines, Zoom, New Window", "Cố định tiêu đề, ẩn gridline cho dashboard", "B1, B6"],
    ["Power Pivot*", "Manage, Measures, KPIs", "Mô hình dữ liệu và DAX (*chỉ xuất hiện sau khi bật add-in)", "B5–B6"],
], [2.6, 5.2, 6.4, 1.8])
figure(d, "dải Ribbon Excel 365, mỗi tab đóng khung và đánh số", "H01_ribbon.png", width=17)

# 2 Shortcuts
d.add_heading("2. Phím tắt Windows và tương đương trên Mac", level=2)
para(d, "Bảng dưới chỉ gồm phím tắt dùng trong khóa học. Cột Mac theo tài liệu tham khảo (Excel Formulas & Functions, trang 13) và các tổ hợp Command tiêu chuẩn của macOS.")
table(d, [
    ["Việc cần làm", "Windows", "Mac"],
    ["Nhảy tới mép vùng dữ liệu", "Ctrl + Mũi tên", "Command + Mũi tên"],
    ["Chọn tới mép vùng dữ liệu", "Ctrl + Shift + Mũi tên", "Command + Shift + Mũi tên"],
    ["Chuyển sheet", "Ctrl + PgUp / PgDn", "Command + Fn + ↑ / ↓"],
    ["Sửa ô, tô màu vùng được tham chiếu", "F2", "Control + U"],
    ["Đổi kiểu tham chiếu A1 → $A$1 → A$1 → $A1", "F4 (khi đang sửa công thức)", "Command + T"],
    ["Lặp lại thao tác cuối", "F4 (khi không sửa công thức)", "Command + Y"],
    ["Tạo Excel Table", "Ctrl + T", "Control + T"],
    ["Flash Fill", "Ctrl + E", "(Data → Flash Fill)"],
    ["Hiện/ẩn công thức", "Ctrl + `", "Control + `"],
    ["Dán giá trị", "Alt → H → V → V", "(Edit → Paste Special → Values)"],
    ["Bật/tắt Filter", "Alt → A → T", "Command + Shift + F"],
    ["Evaluate Formula", "Alt → M → V", "(Formulas → Evaluate Formula)"],
    ["Copy / Dán / Hoàn tác / Lưu / Tìm", "Ctrl + C / V / Z / S / F", "Command + C / V / Z / S / F"],
], [6.4, 4.8, 4.8])
para(d, "Nhắc lại: khóa học yêu cầu Excel trên Windows (Power Pivot không có trên Mac và Excel Online). Nếu dùng Mac, hãy cài máy ảo Windows (Parallels Desktop hoặc VMware Fusion) hoặc chuẩn bị một laptop Windows khác trước Buổi 2, để quen giao diện và phím tắt Windows ngay từ Buổi 2; hạn chót là trước Buổi 4. Bảng Mac ở trên chỉ để bạn tạm dùng trong lúc chuẩn bị máy.", italic=True, color=GREY)
para(d, "Danh sách đầy đủ: exceljet.net/keyboard-shortcuts", italic=True, color=GREY)

# 3 QAT
d.add_heading("3. Quick Access Toolbar", level=2)
tool(d, "Quick Access Toolbar (QAT)",
     "Thanh lệnh nhỏ luôn hiển thị, dùng để ghim các lệnh hay dùng — không cần tìm trong Ribbon.",
     how=["Cách 1: chuột phải vào một lệnh trên Ribbon → Add to Quick Access Toolbar.",
          "Cách 2: File → Options → Quick Access Toolbar → chọn lệnh → Add.",
          "Đặt QAT dưới Ribbon cho gần vùng làm việc: mũi tên cuối QAT → Show Below the Ribbon."],
     image="QAT có Refresh All, Paste Values, Freeze Panes", img="H02_qat.png", width=10,
     tip="Mỗi lệnh trên QAT có phím tắt Alt + số thứ tự (Alt+1, Alt+2…).",
     protip="Ghim sẵn Refresh All ngay từ bây giờ — bạn sẽ dùng nó liên tục từ Buổi 3.")

# 4 Freeze panes
d.add_heading("4. Freeze Panes", level=2)
tool(d, "Freeze Panes",
     "Cố định dòng tiêu đề (và/hoặc cột đầu) khi cuộn bảng dài — không bao giờ mất tên cột.",
     how=["View → Freeze Panes → Freeze Top Row (cố định dòng 1).",
          "Muốn cố định cả dòng và cột: chọn ô ngay dưới-phải vùng cần cố định (ví dụ B2) → Freeze Panes.",
          "Bỏ cố định: View → Freeze Panes → Unfreeze Panes."],
     image="file case 6.187 dòng cuộn xuống dưới, dòng tiêu đề vẫn hiện", img="H03_freeze_panes.png", width=16,
     tip="Excel Table tự thay tên cột A, B, C… bằng tên trường khi cuộn — một lý do nữa để dùng Table.")

# 5 GETPIVOTDATA
d.add_heading("5. Tắt GETPIVOTDATA (chuẩn bị cho Buổi 3)", level=2)
tool(d, "GETPIVOTDATA",
     "Khi bạn gõ = rồi bấm vào một ô của PivotTable, Excel tự sinh hàm GETPIVOTDATA thay vì tham chiếu ô thường. "
     "Hàm này không đổi khi kéo công thức, gây kết quả lặp lại khó hiểu.",
     how=["Tắt theo từng Pivot: PivotTable Analyze → Options (mũi tên) → bỏ chọn Generate GetPivotData.",
          "Tắt cho mọi file: File → Options → Formulas → bỏ chọn \"Use GetPivotData functions for PivotTable references\"."],
     image="hộp File → Options → Formulas với ô GetPivotData bỏ chọn", img="H04_getpivotdata.png", width=11,
     protip="GETPIVOTDATA không xấu — nó hữu ích khi bố cục Pivot thay đổi. Chỉ tắt khi bạn cần kéo công thức.")

# 6 Top 10
d.add_heading("6. Mười hàm dùng nhiều nhất trong BI", level=2)
para(d, "Ví dụ viết trên bảng tblSales của file đáp án Buổi 1 (cột Region, Store_Code, Quantity, Unit_Price, Revenue…). Buổi 2 sẽ học sâu nhóm 6–10.")
table(d, [
    ["#", "Hàm", "Dùng để", "Ví dụ trên case"],
    ["1", "SUM", "Cộng", "=SUM(tblSales[Revenue])"],
    ["2", "COUNT / COUNTA", "Đếm ô số / ô không trống", "=COUNTA(tblSales[Order_ID])"],
    ["3", "AVERAGE", "Trung bình", "=AVERAGE(tblSales[Unit_Price])"],
    ["4", "MIN / MAX", "Nhỏ nhất / lớn nhất", "=MAX(tblSales[Order_Date])"],
    ["5", "ROUND", "Làm tròn", "=ROUND(SUM(tblSales[Revenue])/1e6, 1)  → triệu đồng"],
    ["6", "IF", "Rẽ nhánh theo điều kiện", "=IF([@Revenue]>=500000, \"Đơn lớn\", \"Đơn thường\")"],
    ["7", "IFERROR", "Thay lỗi bằng giá trị khác", "=IFERROR([@Revenue]/[@Quantity], 0)"],
    ["8", "SUMIFS", "Cộng có điều kiện", "=SUMIFS(tblSales[Revenue], tblSales[Region], \"Bắc\")"],
    ["9", "COUNTIFS", "Đếm có điều kiện", "=COUNTIFS(tblSales[Store_Code], \"ST07\")"],
    ["10", "XLOOKUP", "Tra cứu từ bảng khác", "=XLOOKUP([@Store_Code], tblStore[Store_Code], tblStore[City])"],
], [0.8, 3.2, 4.2, 7.8])
rich(d, [("PRO TIP: ", {"bold": True, "color": GOLD}), ("Trong Table, [@Revenue] nghĩa là \"ô Revenue của chính dòng này\". Công thức đọc được bằng tên cột dễ kiểm tra hơn nhiều so với K2*J2.", {})])

# 7 Flash fill
d.add_heading("7. Flash Fill", level=2)
tool(d, "Flash Fill (Ctrl + E)",
     "Excel đoán quy luật từ 1–2 ví dụ bạn gõ tay rồi điền cho cả cột — tách, ghép, đổi dạng chuỗi mà không cần công thức.",
     how=["Ví dụ: cột Order_ID có dạng HD2501-00001. Ở cột bên cạnh gõ 2501 cho dòng đầu.",
          "Chọn ô tiếp theo → Ctrl + E (hoặc Data → Flash Fill) → Excel điền kỳ \"năm-tháng\" cho mọi dòng."],
     image="cột mới được Flash Fill điền 2501, 2502, 2503", img="H05_flash_fill.png", height=10,
     tip="Flash Fill cho ra GIÁ TRỊ, không phải công thức: dữ liệu mới thêm vào sẽ không tự cập nhật. Việc lặp lại mỗi tháng nên làm bằng công thức (Buổi 2) hoặc Power Query (Buổi 4).")

# 8 Fill series
d.add_heading("8. Fill Series nâng cao", level=2)
tool(d, "Fill Series",
     "Tạo nhanh chuỗi số hoặc ngày theo quy luật — hữu ích khi dựng bảng tháng, bảng ngày làm việc.",
     how=["Gõ giá trị đầu (ví dụ 01/01/2025) → Home → Fill → Series.",
          "Chọn Series in (Rows/Columns), Type = Date, Date unit = Day / Weekday / Month / Year, Step value, Stop value.",
          "Ví dụ: Date unit = Month, Step = 1, Stop = 01/12/2025 → 12 ngày đầu tháng của năm 2025.",
          "Date unit = Weekday → chỉ sinh ngày thứ Hai–thứ Sáu."],
     image="hộp thoại Series với Type = Date, Date unit = Month", img="H06_fill_series.png", width=15,
     protip="Ở Buổi 4 bạn sẽ tạo bảng Calendar tự động bằng Power Query — Fill Series là cách làm tay tương đương.")

# 9 External refs
d.add_heading("9. Tham chiếu sang workbook khác", level=2)
tool(d, "External Reference",
     "Lấy giá trị từ một file Excel khác, ví dụ lấy Order_ID từ bảng tblSales của file case: =B1_MCI_Mart_Q1_2025_Answer.xlsx!tblSales[@[Order_ID]].",
     how=["Mở cả hai file, gõ = ở file đích, chuyển sang file nguồn và bấm ô cần lấy, Enter.",
          "Quản lý liên kết: Data → Queries & Connections → Workbook Links (bản cũ: Data → Edit Links)."],
     image="công thức chứa B1_MCI_Mart_Q1_2025_Answer.xlsx!tblSales và hộp Workbook Links", img="H07_external_ref.png", width=17,
     tip="Liên kết dễ gãy khi đổi tên hoặc di chuyển file nguồn (lỗi #REF!).",
     protip="Với dữ liệu cần cập nhật định kỳ từ file khác, Power Query (Buổi 4) an toàn và dễ kiểm soát hơn tham chiếu ngoài.")

# 10 Homework (bắt buộc) — file đề/đáp án dựng bằng build_homework.py
d.add_heading("10. Bài tập về nhà (bắt buộc)", level=2)
rich(d, [("File đề: ", {"bold": True}), ("B1_BTVN_MCI_Mart_T4_2025.xlsx", {"code": True}),
         (" — báo cáo bán hàng tháng 4/2025 của MCI Mart, xuất từ hệ thống POS theo từng cửa hàng. "
          "Nộp trước Buổi 2, đặt tên file B1_BTVN_<HoTen>.xlsx. Sheet De_bai trong file ghi đủ yêu cầu.", {})])
rich(d, [("Nộp bài: ", {"bold": True, "color": NAVY}), ("https://hocvien.mcivietnam.com/auth#/", {"code": True}),
         (" — Tài khoản: email bạn đăng ký tại MCI · Mật khẩu: hv12345678", {})])
for s_ in ["Làm sạch cấu trúc: xóa tiêu đề báo cáo, dòng \"Cửa hàng …\", dòng \"Cộng …\", \"TỔNG CỘNG\", dòng tiêu đề bị lặp lại và dòng trống.",
           "Chuyển Quantity và Order_Date đang ở dạng text thành số và ngày.",
           "Ctrl + T, đặt tên bảng tblSales_T4; thêm cột Revenue = [@Quantity]*[@Unit_Price].",
           "Tạo ô tổng doanh thu; thêm cột Pct_Share = Revenue ÷ ô tổng (dùng $).",
           "Data Validation cho Store_Code (List từ sheet Store), Product_Code (List từ sheet Product), Quantity (số nguyên > 0); dùng Circle Invalid Data để tìm và sửa các ô sai.",
           "Viết 1 câu insight có đủ con số, nguyên nhân, hàm ý (theo mẫu slide 40)."]:
    bullet(d, s_)
rich(d, [("Tự kiểm tra trước khi nộp: ", {"bold": True, "color": NAVY}),
         ("1.919 dòng dữ liệu; tổng doanh thu 296.965.000 VNĐ; tổng Quantity 3.448; tổng Pct_Share = 100%; Circle Invalid Data không còn ô nào.", {})])

# 11 Checklist
# Giữ trọn mục 11 trên một trang (keep_with_next cho tiêu đề và mọi dòng trừ dòng cuối)
h11 = d.add_heading("11. Checklist trước Buổi 2", level=2)
h11.paragraph_format.keep_with_next = True
for s in ["Máy đã qua kiểm tra: có Get Data, có tab Power Pivot, biết bản 32/64-bit.",
          "Nếu dùng Mac hoặc máy không có Power Pivot: đã cài máy ảo Windows (Parallels Desktop hoặc VMware Fusion) hoặc đã có một laptop Windows khác để mang theo.",
          "File B1 của bạn: là Table tên tblSales, không ô gộp, không dòng tổng, không dòng trống.",
          "Có cột Revenue và Pct_Share; tổng Pct_Share = 100%.",
          "Tổng doanh thu Q1/2025 = 941.454.000 VNĐ; 6.187 dòng. Nếu chưa khớp: dùng file khởi đầu Buổi 2 do lớp phát.",
          "Đã viết câu insight cuối Buổi 1.",
          "Đã nộp bài tập về nhà (B1_BTVN_<HoTen>.xlsx), tự kiểm tra khớp 1.919 dòng và 296.965.000 VNĐ."]:
    bullet(d, "☐  " + s).paragraph_format.keep_with_next = True
d.paragraphs[-1].paragraph_format.keep_with_next = False
d.save(f"{OUT}/B1_Handout.docx")

# =====================================================================
# MACHINE CHECKLIST
# =====================================================================
d = base_doc()
para(d, "MCI – MAGIC CODE INSTITUTE", bold=True, color=GREY, size=9)
d.add_heading("Checklist kiểm tra máy — Buổi 1", level=1)
para(d, "Khóa học cần Excel có Power Query và Power Pivot. Hãy tự làm 3 bước dưới đây trên chính máy bạn sẽ dùng cho khóa học, trước Buổi 1. Nếu bước nào KHÔNG ĐẠT hoặc bạn không làm được, hãy liên hệ trợ giảng ngay trong Buổi 1 để được hỗ trợ.")
table(d, [
    ["Bước", "Kiểm tra", "Cách làm", "Đạt khi"],
    ["1", "Power Query", "Mở tab Data, tìm nút Get Data (nhóm Get & Transform Data)", "Thấy nút Get Data → From File / From Database…"],
    ["2", "Power Pivot", "File → Options → Add-ins → ô Manage chọn COM Add-ins → Go → tick Microsoft Power Pivot for Excel → OK",
     "Trên Ribbon xuất hiện tab Power Pivot"],
    ["3", "32 / 64-bit", "File → Account → About Excel, đọc dòng đầu", "Ghi lại 32-bit hoặc 64-bit (khuyến nghị 64-bit)"],
], [1.2, 2.8, 7.2, 4.8])
d.add_heading("Nếu không đạt: liên hệ trợ giảng trong Buổi 1", level=2)
for s in ["Dùng Mac hoặc Excel Online → không có Power Pivot: khuyến nghị cài phần mềm máy ảo (Parallels Desktop hoặc VMware Fusion), trong đó cài Windows và Excel bản Windows; nếu không cài được, tự chuẩn bị một laptop Windows khác (lớp không có máy cho mượn). Nên có trước Buổi 2 để quen giao diện Windows; bắt buộc trước Buổi 4 (buổi bắt đầu dùng Data Model).",
          "Không thấy Microsoft Power Pivot trong COM Add-ins → phiên bản Office không kèm Power Pivot: cần Microsoft 365 hoặc bản Office có Power Pivot.",
          "Có Power Pivot nhưng bản 32-bit → vẫn học được; báo trợ giảng nếu máy chậm khi mở mô hình dữ liệu ở Buổi 5."]:
    bullet(d, s)
d.save(f"{OUT}/B1_Checklist_KiemTraMay.docx")

print("ok")
