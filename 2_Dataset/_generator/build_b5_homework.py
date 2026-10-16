r"""Bài tập về nhà BẮT BUỘC Buổi 5 — kiểm toán báo cáo cũ của phòng Kinh doanh (tập dượt capstone C3 Mô hình, C4 KPI; D2).

  Homework/B5_BTVN_MCI_Mart_KPI_T1_T4_2025.xlsx          đề (phát học viên)
  Homework/B5_BTVN_MCI_Mart_KPI_T1_T4_2025_Answer.xlsx   đáp án (chỉ giảng viên / trợ giảng): Kiem_tra, Dap_an_loi, Insight_mau

ĐỀ  = đáp án B4 (Data Model 6 bảng) bỏ các sheet B1–B4, thêm đúng mô hình "của phòng Kinh doanh":
      Sales nối Product, Store, Customer, Calendar; Target CHỈ nối Store (chưa nối Calendar).
      6 measure KD (KD Doanh thu, KD Target, KD % dat, KD So don = COUNTROWS, KD AOV, KD GM % = AVERAGEX trung bình
      từng dòng) và sheet BaoCao_KinhDoanh (Pivot From Data Model, tháng lấy từ Sales[Thang]).
      Excel không cho script tạo calculated column nên "cột GM_Dong" của kịch bản được thay bằng measure AVERAGEX
      cùng kết quả (trung bình biên từng dòng).
ĐÁP ÁN = file đáp án Practice B5 (6 relationship, 14 measure) bỏ sheet Pivot cũ, thêm 3 measure like-for-like,
      BaoCao_KinhDoanh dựng lại từ measure, BaoCao_LFL, Nhat_ky_doi_soat / Brief_Insight mẫu, Kiem_tra, Dap_an_loi,
      Insight_mau. KPI_Dictionary giữ từ file Practice, thêm 3 dòng LFL.

Số liệu đọc từ build_b5_meta.json (khóa hw). CHỈ CHẠY TRÊN WINDOWS (Excel qua COM, datamodel_com.run).
    python 2_Dataset\_generator\build_b5_homework.py            # dựng đề + đáp án
    python 2_Dataset\_generator\build_b5_homework.py --check    # mở lại, kiểm số liệu BaoCao_KinhDoanh và Kiem_tra
Trước khi chạy: đóng B4_MCI_Mart_Answer.xlsx, B5_MCI_Mart_Answer.xlsx và 2 file bài tập B5 nếu đang mở.
"""
import json, os, sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from datamodel_com import Rel, Measure, CubePivot, Cells, run

for s_ in (sys.stdout, sys.stderr):
    if hasattr(s_, "reconfigure"):
        s_.reconfigure(encoding="utf-8", errors="replace")

GEN = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.normpath(os.environ.get("MCI_ROOT") or os.path.dirname(os.path.dirname(GEN)))
B4_ANSWER = os.path.join(ROOT, "1_Sessions", "B4_Power_Query", "Practice", "B4_MCI_Mart_Answer.xlsx")
B5 = os.path.join(ROOT, "1_Sessions", "B5_Data_Model_DAX")
B5_ANSWER = os.path.join(B5, "Practice", "B5_MCI_Mart_Answer.xlsx")
HW = os.path.join(B5, "Homework")
HW_DE = os.path.join(HW, "B5_BTVN_MCI_Mart_KPI_T1_T4_2025.xlsx")
HW_ANS = os.path.join(HW, "B5_BTVN_MCI_Mart_KPI_T1_T4_2025_Answer.xlsx")
M = json.load(open(os.path.join(GEN, "build_b5_meta.json"), encoding="utf-8"))
H = M["hw"]
MON = H["months"]


def n(x):
    return f"{x:,.0f}".replace(",", ".")


def p(x, d=1):
    return f"{x * 100:.{d}f}".replace(".", ",") + "%"


def sp(x, d=1):
    return ("+" if x >= 0 else "−") + p(abs(x), d)


NAVY, NOTE_C = "203864", "585858"
DROP_OLD = ["KPI_Card", "Kiem_tra_B4", "PV_Refresh", "PV_MienThang", "PV_NganhHang", "PV_Thang", "PV_Ngay", "PV_Thu",
            "PV_CuaHang", "ST07_T3_ChiTiet", "Kiem_tra", "Sales", "Product", "Store", "KPI_SUMIFS"]
DROP_PRACTICE = ["PV_Loc", "PV_Target", "PV_Target_CuaHang", "PV_YoY", "PV_GM_Nam", "PV_KPI_Thang", "PV_DoiSoat_B2",
                 "DoiSoat_B2", "Kiem_tra_B5"] + DROP_OLD

NAM, THANG, QUY = "[Calendar].[Nam]", "[Calendar].[Thang]", "[Calendar].[Quy]"
SNAM, STHANG = "[Sales].[Nam]", "[Sales].[Thang]"
M14 = ["1", "2", "3", "4"]

# ---------------------------------------------------------------- Đề: mô hình của phòng Kinh doanh
DE_RELS = [
    Rel("Sales", "Product_Code", "Product", "Product_Code"),
    Rel("Sales", "Store_Code", "Store", "Store_Code"),
    Rel("Sales", "Customer_ID", "Customer", "Customer_ID"),
    Rel("Sales", "Order_Date", "Calendar", "Date"),
    Rel("Target", "Store_Code", "Store", "Store_Code"),          # Target chưa nối Calendar: lỗi cài sẵn #1
]
KD = [
    Measure("KD Doanh thu", "Sales", "SUM(Sales[Revenue])", "int", "Phòng Kinh doanh: tổng doanh thu."),
    Measure("KD Target", "Target", "SUM(Target[Target])", "int", "Phòng Kinh doanh: tổng target."),
    Measure("KD % dat", "Sales", "DIVIDE([KD Doanh thu], [KD Target])", "pct1", "Phòng Kinh doanh: doanh thu ÷ target."),
    Measure("KD So don", "Sales", "COUNTROWS(Sales)", "int", "Phòng Kinh doanh: số đơn."),
    Measure("KD AOV", "Sales", "DIVIDE([KD Doanh thu], [KD So don])", "int", "Phòng Kinh doanh: doanh thu ÷ số đơn."),
    Measure("KD GM %", "Sales", "AVERAGEX(Sales, DIVIDE(Sales[Revenue] - Sales[Quantity] * RELATED(Product[Unit_Cost]), Sales[Revenue]))", "pct2",
            "Phòng Kinh doanh: biên lợi nhuận gộp (trung bình biên từng dòng)."),
]
DE_PIVOT = CubePivot("BaoCao_KinhDoanh", "BaoCao_KinhDoanh", anchor="A8", rows=[STHANG],
                     values=[m.name for m in KD], filters={SNAM: ["2025"]}, row_items={STHANG: M14},
                     title="BÁO CÁO KINH DOANH THÁNG 1–4/2025 (phòng Kinh doanh gửi Sales Director)",
                     note="Tháng lấy từ cột Sales[Thang]. Số của phòng Tài chính: sheet So_lieu_Tai_chinh.")

# ---------------------------------------------------------------- helpers ô
def block(sheet, r0, header, rows, widths=None, wrap_cols="ABCDEF", fill_header=True, last_col=None, **kw):
    cols = "ABCDEFGHIJ"
    lc = last_col or cols[len(header) - 1]
    items = []
    for j, h in enumerate(header):
        items.append((f"{cols[j]}{r0}", h))
    for i, row in enumerate(rows, start=r0 + 1):
        for j, v in enumerate(row):
            if v not in (None, ""):
                items.append((f"{cols[j]}{i}", v))
    last = r0 + len(rows)
    return items, last, lc


def cells(sheet, items, widths, bold=(), wrap=(), fill=None, font_color=None, numfmt=None, title=True):
    return Cells(sheet, items, widths=widths, bold=list(bold), wrap=list(wrap), fill=fill or {}, font_color=font_color or {},
                 numfmt=numfmt or {}, title="x" if title else "")


# ---------------------------------------------------------------- So_lieu_Tai_chinh
tc_rows = []
for m in M14:
    d = MON[m]
    tc_rows.append((int(m), d["tc_rev"], d["tc_orders"], round(d["tc_aov"]), d["tc_gm"], d["tc_target"], d["tc_vs_target"]))
tc_items = [("A1", 'SỐ LIỆU CHÍNH THỨC CỦA PHÒNG TÀI CHÍNH, THÁNG 1–4/2025'),
            ("A2", "Báo cáo của bạn phải khớp các con số này. Số nào lệch: ghi vào Nhat_ky_doi_soat và tìm nguyên nhân.")]
it, last, _ = block("So_lieu_Tai_chinh", 4, ["Tháng", "Doanh thu (VNĐ)", "Số đơn", "AOV (VNĐ)", "Gross Margin %", "Target (VNĐ)", "% đạt target"], tc_rows)
tc_items += it
tc_sheet = cells("So_lieu_Tai_chinh", tc_items, {"A": 8, "B": 17, "C": 10, "D": 13, "E": 15, "F": 16, "G": 13},
                 bold=["A4:G4"], fill={"A4:G4": NAVY}, font_color={"A4:G4": "FFFFFF"},
                 numfmt={f"B5:D{last}": "#,##0", f"F5:F{last}": "#,##0", f"E5:E{last}": "0.00%", f"G5:G{last}": "0.0%"})

# ---------------------------------------------------------------- De_bai
TASKS = [('1', 'Đối soát', 'So từng chỉ tiêu, từng tháng của BaoCao_KinhDoanh với So_lieu_Tai_chinh: doanh thu, số đơn, AOV, GM%, % đạt target. Mỗi lỗi gốc ghi một dòng vào Nhat_ky_doi_soat: chỉ tiêu, số báo cáo cũ, số Tài chính, chênh lệch, nguyên nhân, cách sửa. Chỉ tiêu bị kéo theo (ví dụ AOV kéo theo số đơn) ghi chung dòng với lỗi gốc.', 'Nhat_ky_doi_soat có đúng 4 dòng, mỗi dòng một lỗi gốc và nêu các chỉ tiêu bị ảnh hưởng. Doanh thu khớp Tài chính ở cả 4 tháng. Lỗi nhỏ nhất chỉ lệch dưới nửa điểm phần trăm, nên phải so từng ô mới thấy.', 'Tháng 4: báo cáo cũ ghi đạt 8,1% target, Tài chính ghi 103,1%. Doanh thu cả 4 tháng không lệch.'),
    ('2', 'Sửa mô hình thành hình ngôi sao', 'Power Pivot → Diagram View. Kéo Target[Ngay_Dau_Thang] thả lên Calendar[Date]. Ẩn khỏi Client Tools các cột khóa và cột ngày của hai bảng fact (Sales: Nam, Thang, Order_Date, Store_Code, Product_Code, Customer_ID; Target: Store_Code, Thang, Ngay_Dau_Thang). Calendar → Design → Mark as Date Table, chọn cột Date. Chụp Diagram View, dán vào sheet Brief_Insight (bên dưới dòng 3).', 'Diagram View là hình ngôi sao: hai bảng fact Sales và Target cùng nối Store và Calendar. Danh sách trường của Pivot không còn các cột đã ẩn.', '6 bảng, 6 relationship. Calendar hiện biểu tượng bảng ngày. Ảnh Diagram View nằm trên sheet Brief_Insight.'),
    ('3', 'Viết lại KPI bằng explicit measure', 'Viết measure theo tầng (measure cơ sở → measure dẫn xuất): Total Revenue, Total Orders (DISTINCTCOUNT của Order_ID, không đếm dòng), Avg Order Value, Gross Margin % (tổng ÷ tổng, dùng Total COGS đã làm ở Slide 15 và 17), Active Customers, Revenue YoY %, Total Target, Revenue vs Target %, Variance. Thêm like-for-like (chỉ cửa hàng mở trước 01/01/2024): Revenue LFL = CALCULATE([Total Revenue], Store[Open_Date] < DATE(2024, 1, 1)), Revenue LFL LY, YoY LFL %. Cuối cùng xóa 6 measure KD.', 'Mô hình có đủ 12 measure nêu ở cột Yêu cầu và không còn measure KD. Mỗi measure dẫn xuất gọi measure cơ sở bằng [tên], không gõ lại công thức.', 'Tháng 4/2025: Total Orders = 966, Gross Margin % = 33,8%, Revenue LFL = 280.658.000.'),
    ('4', 'Dựng lại báo cáo', 'Sheet BaoCao_KinhDoanh: xóa Pivot cũ, dựng Pivot mới từ measure của bạn. Rows = Calendar[Thang], Filters = Calendar[Nam] chọn 2025, tháng 1–4. Cột: Total Revenue, Total Target, Revenue vs Target %, Total Orders, Avg Order Value, Gross Margin %. Thêm một Pivot cửa hàng × tháng cho Revenue vs Target % (bên dưới Pivot chính) để tìm cửa hàng dưới 80% target. Tạo sheet mới BaoCao_LFL: Revenue YoY % và YoY LFL % theo tháng. Đối soát lại với So_lieu_Tai_chinh.', 'BaoCao_KinhDoanh khớp So_lieu_Tai_chinh ở 4 tháng × 5 chỉ tiêu, chênh lệch bằng 0. BaoCao_LFL có 4 dòng tháng và dòng Grand Total.', 'BaoCao_LFL, Grand Total: YoY cả chuỗi +21,0%, like-for-like +13,3%. Cửa hàng dưới 80% target: T2 có 5, T3 có 5, T4 có 2.'),
    ('5', 'KPI_Dictionary', 'Mỗi measure ở việc 3 một dòng: định nghĩa, công thức DAX, grain, dimensions, owner, target, ngưỡng Warning / Critical và cột Giới hạn. Cột Giới hạn nói rõ khi nào KPI dễ hiểu sai: Revenue vs Target % không theo ngành hàng; Gross Margin % dùng một bảng giá vốn cho mọi năm; YoY cả chuỗi gồm ST12 chưa có kỳ gốc.', 'KPI_Dictionary có 12 dòng, không ô nào trống ở cột DAX và Giới hạn. Giới hạn của ba KPI nêu trên viết bằng lời của bạn.', 'Công thức DAX trong từ điển giống hệt measure trong mô hình.'),
    ('6', 'Brief_Insight', 'Viết 3 câu hỏi đo lường được cho Sales Director. Mỗi câu có đủ chỉ số · chiều · thời gian · mốc so sánh và ghi tên sheet chứa Pivot trả lời. Ví dụ: "Doanh thu theo vùng, tháng 4/2025, vùng nào tăng dưới 10% so với tháng 4/2024?". Sau đó viết 1 câu insight gồm con số có so sánh · nguyên nhân khả dĩ · hàm ý.', 'Brief_Insight có 3 câu hỏi, mỗi câu đủ 4 thành phần và ghi tên sheet Pivot; 1 câu insight đủ 3 phần; con số trong insight khớp Pivot.', 'Insight có đủ ba vế (so sánh, nguyên nhân, hàm ý), không dừng ở việc đọc lại con số.')]
SELF = [
    f"Tháng 4/2025: doanh thu {n(MON['4']['tc_rev'])} VNĐ · {n(MON['4']['tc_orders'])} đơn · AOV {n(MON['4']['tc_aov'])} · GM {p(MON['4']['tc_gm'])} · đạt {p(MON['4']['tc_vs_target'])} target",
    f"% đạt target tháng 1–4: {' · '.join(p(MON[m]['tc_vs_target']) for m in M14)}",
    f"YoY cả chuỗi tháng 1–4 (lũy kế): {sp(H['ytd_yoy_t4'])}; like-for-like ({H['lfl_stores']} cửa hàng): {sp(H['lfl_ytd_yoy'])}",
    f"YoY like-for-like theo tháng: {' · '.join('T' + m + ' ' + sp(MON[m]['lfl_yoy']) for m in M14)}",
    "Cửa hàng đạt dưới 80% target: " + " · ".join(f"T{m}: {', '.join(H['below80'][m]) if H['below80'][m] else 'không có'}" for m in M14),
    "Diagram View là hình ngôi sao: 6 bảng, 6 relationship, hai bảng fact (Sales, Target) cùng nối Store và Calendar",
]
de_items = [("A1", 'BÀI TẬP VỀ NHÀ BUỔI 5 (BẮT BUỘC): KIỂM TOÁN BÁO CÁO CỦA PHÒNG KINH DOANH'),
            ("A2", 'Nộp trước Buổi 6. Đặt tên file: B5_BTVN_<HoTen>.xlsx (làm ngay trên file này). Tình huống: phòng Kinh doanh gửi Sales Director báo cáo tháng 1–4/2025 (sheet BaoCao_KinhDoanh), còn phòng Tài chính báo số khác (sheet So_lieu_Tai_chinh). Bạn là BI analyst được giao kiểm tra: số nào đúng, vì sao lệch, và sửa thế nào để hai phòng dùng chung một định nghĩa. Bài tập dượt capstone B7: tiêu chí C3 (Mô hình dữ liệu) và C4 (KPI); việc 6 dượt thêm C1 (Câu hỏi) và C6 (Insight).'),
            ("A3", 'Nộp bài tại https://hocvien.mcivietnam.com/auth#/ — Tài khoản: email bạn đăng ký tại MCI · Mật khẩu: hv12345678'),
            ("A4", 'Không bấm Refresh All trong file này: dữ liệu đã nằm sẵn trong Data Model (6 bảng: Sales, Product, Store, Customer, Calendar, Target). Chỉ dựng Pivot từ 6 bảng này, không kéo trường từ các bảng khác nếu thấy trong tab All.')]
it, last, _ = block("De_bai", 5, ['#', 'Việc cần làm', 'Yêu cầu', 'Kết quả kỳ vọng', 'Tự kiểm tra nhanh'], TASKS)
de_items += it
kr = last + 2
de_items.append((f"A{kr}", "Tự kiểm tra trước khi nộp"))
for i, tx in enumerate(SELF, start=kr + 1):
    de_items.append((f"B{i}", "☐ " + tx))
de_items.append((f"A{kr + len(SELF) + 2}", 'Gợi ý: xem handout B5, mục 9 (bài tập về nhà) và KPI Definition Card của Buổi 2.'))
de_sheet = cells("De_bai", de_items, {"A": 5, "B": 30, "C": 75, "D": 55, "E": 42}, bold=["A5:E5", f"A{kr}"], fill={"A5:E5": NAVY},
                 font_color={"A5:E5": "FFFFFF"}, wrap=[f"A6:E{last}", "A2:E3"])

# ---------------------------------------------------------------- Nhat_ky_doi_soat / KPI_Dictionary / Brief_Insight (mẫu trống cho học viên)
nk_head = ["Chỉ tiêu", "Số báo cáo cũ", "Số Tài chính", "Chênh lệch", "Nguyên nhân", "Cách sửa"]
nk_items = [("A1", 'NHẬT KÝ ĐỐI SOÁT: BÁO CÁO KINH DOANH ↔ SỐ TÀI CHÍNH'),
            ("A2", 'Mỗi lỗi gốc một dòng; chỉ tiêu bị kéo theo ghi chung dòng với lỗi gốc. Ghi cả tháng bạn so sánh.')]
it, last, _ = block("Nhat_ky_doi_soat", 4, nk_head, [()] * 8)
nk_items += it
nk_sheet = cells("Nhat_ky_doi_soat", nk_items, {"A": 28, "B": 18, "C": 16, "D": 14, "E": 55, "F": 55},
                 bold=["A4:F4"], fill={"A4:F4": NAVY}, font_color={"A4:F4": "FFFFFF"})
kd_head = ["KPI", "Measure", "DAX", "Định dạng", "Grain", "Dimensions", "Owner", "Target", "Warning / Critical", "Giới hạn / ghi chú"]
kd_items = [("A1", 'KPI DICTIONARY: MCI Mart'),
            ("A2", "Một dòng cho mỗi measure đã viết ở việc 3. Cột Giới hạn: nói rõ khi nào KPI này không được dùng hoặc dễ hiểu sai.")]
it, last, _ = block("KPI_Dictionary", 4, kd_head, [()] * 12)
kd_items += it
kd_sheet = cells("KPI_Dictionary", kd_items, {"A": 20, "B": 20, "C": 52, "D": 10, "E": 14, "F": 28, "G": 15, "H": 15, "I": 14, "J": 60},
                 bold=["A4:J4"], fill={"A4:J4": NAVY}, font_color={"A4:J4": "FFFFFF"})
bi_items = [("A1", "BRIEF CHO SALES DIRECTOR"),
            ("A3", "Ảnh Diagram View của bạn (dán dưới đây, bên dưới dòng này):"),
            ("A22", '3 câu hỏi đo lường được (chỉ số · chiều · thời gian · mốc so sánh), mỗi câu ghi tên sheet có Pivot trả lời:'),
            ("A23", "1."), ("A25", "2."), ("A27", "3."),
            ("A30", "Câu insight (con số có so sánh · nguyên nhân khả dĩ · hàm ý):")]
bi_sheet = cells("Brief_Insight", bi_items, {"A": 130}, bold=["A3", "A22", "A30"])

DE_SHEETS = [de_sheet, tc_sheet, nk_sheet, kd_sheet, bi_sheet]

# ---------------------------------------------------------------- Đáp án
LFL = [
    Measure("Revenue LFL", "Sales", "CALCULATE([Total Revenue], Store[Open_Date] < DATE(2024, 1, 1))", "int",
            "Doanh thu like-for-like: chỉ cửa hàng mở trước 01/01/2024 (11 cửa hàng; loại ST12)."),
    Measure("Revenue LFL LY", "Sales", "CALCULATE([Revenue LFL], DATEADD(Calendar[Date], -1, YEAR))", "int", "Doanh thu like-for-like cùng kỳ năm trước."),
    Measure("YoY LFL %", "Sales", "IF(ISBLANK([Revenue LFL]), BLANK(), DIVIDE([Revenue LFL] - [Revenue LFL LY], [Revenue LFL LY]))", "pct1",
            "Tăng trưởng like-for-like so với cùng kỳ năm trước."),
]
ANS_PIVOTS = [
    CubePivot("BaoCao_KinhDoanh", "BaoCao_KinhDoanh", anchor="A8", rows=[THANG],
              values=["Total Revenue", "Total Target", "Revenue vs Target %", "Total Orders", "Avg Order Value", "Gross Margin %"],
              filters={NAM: ["2025"]}, row_items={THANG: M14},
              title="BÁO CÁO KINH DOANH THÁNG 1–4/2025 — DỰNG LẠI TỪ MEASURE (đáp án)",
              note="Tháng lấy từ Calendar[Thang]; Target nối Calendar; số đơn đếm Order_ID không trùng; GM% tổng ÷ tổng. Khớp So_lieu_Tai_chinh."),
    CubePivot("BaoCao_LFL", "BaoCao_LFL", anchor="A8", rows=[THANG],
              values=["Total Revenue", "Revenue LY", "Revenue YoY %", "Revenue LFL", "Revenue LFL LY", "YoY LFL %"],
              filters={NAM: ["2025"]}, row_items={THANG: M14},
              title="YoY cả chuỗi và YoY like-for-like theo tháng (đáp án)",
              note="Cả chuỗi gồm ST12 (khai trương 01/06/2024, chưa có kỳ gốc) nên cao hơn like-for-like."),
]


def mem(field_, key):
    return f"{field_}.&[{key}]"


def cv(measure, *members):
    args = ",".join(f'"{x}"' for x in [f"[Measures].[{measure}]", *members])
    return f'=CUBEVALUE("ThisWorkbookDataModel",{args})'


Y25 = mem(NAM, 2025)
MONEY, PCT = 0.5, 0.00005
kt_rows = []
for m in M14:
    d = MON[m]
    mm = mem(THANG, m)
    kt_rows += [
        (f"Total Revenue T{m}/2025 (= Tài chính)", cv("Total Revenue", Y25, mm), d["tc_rev"], MONEY, "#,##0"),
        (f"Revenue vs Target % T{m}/2025 — cần Target nối Calendar", cv("Revenue vs Target %", Y25, mm), d["tc_vs_target"], PCT, "0.00%"),
        (f"Total Orders T{m}/2025 — đếm Order_ID không trùng", cv("Total Orders", Y25, mm), d["tc_orders"], 0, "#,##0"),
        (f"Avg Order Value T{m}/2025", cv("Avg Order Value", Y25, mm), d["tc_aov"], MONEY, "#,##0"),
        (f"Gross Margin % T{m}/2025 — tổng ÷ tổng", cv("Gross Margin %", Y25, mm), d["tc_gm"], PCT, "0.00%"),
        (f"YoY like-for-like T{m}/2025", cv("YoY LFL %", Y25, mm), d["lfl_yoy"], PCT, "0.00%"),
    ]
kt_items = [("A1", "KIỂM TRA / ĐỐI SOÁT — Bài tập về nhà B5"),
            ("A2", "Cột C: CUBEVALUE đọc measure trong Data Model của file đáp án. Chấm bài: đọc giá trị Pivot BaoCao_KinhDoanh / BaoCao_LFL của học viên (cột tương ứng), so với cột Đáp án."),
            ("A4", "Chỉ tiêu"), ("C4", "Kết quả"), ("D4", "Đáp án"), ("E4", "Khớp?")]
kt_fmt = {}
for i, (label, formula, exp, tol, nf) in enumerate(kt_rows, start=5):
    kt_items += [(f"A{i}", label), (f"C{i}", formula), (f"D{i}", float(exp)),
                 (f"E{i}", f'=IF(ISNUMBER(C{i}),IF(ABS(C{i}-D{i})<={tol},"✔","✘"),"✘")')]
    kt_fmt[f"C{i}:D{i}"] = nf
KT_LAST = 4 + len(kt_rows)
kt_sheet = cells("Kiem_tra", kt_items, {"A": 66, "B": 2, "C": 18, "D": 18, "E": 8}, bold=["A4:E4"],
                 fill={"A4:E4": NAVY}, font_color={"A4:E4": "FFFFFF"}, numfmt=kt_fmt)
CHECKS = [f"Kiem_tra!E{i}" for i in range(5, KT_LAST + 1)]

# Dap_an_loi
m1, m2, m3, m4 = (MON[m] for m in M14)
errs = [("#", "Lỗi cài sẵn", "Báo cáo cũ (T1 · T2 · T3 · T4) và số Tài chính", "Vì sao", "Cách sửa", "Tiêu chí"),
        ("1", "Target chưa nối Calendar",
         f"Target tháng = {n(m1['cu_target'])} (target cả năm) ở mọi tháng → % đạt {' · '.join(p(MON[m]['cu_vs_target']) for m in M14)}. Tài chính: {' · '.join(p(MON[m]['tc_vs_target']) for m in M14)}.",
         "Target chỉ nối Store; bộ lọc tháng không chảy từ Calendar sang Target, nên Pivot đưa cả tổng target vào mọi dòng (bài học Slide 12 và 20).",
         "Diagram View: kéo Target[Ngay_Dau_Thang] thả lên Calendar[Date].", "C3"),
        ("2", "Báo cáo lấy tháng từ Sales[Thang] (cột của bảng fact)",
         "Cùng triệu chứng lỗi 1. Sửa lỗi 1 mà không sửa lỗi này thì Target vẫn lặp.",
         "Hàng, cột, bộ lọc phải lấy từ bảng lookup (Calendar). Thang của Sales không lọc được Target.",
         "Rows = Calendar[Thang], Filters = Calendar[Nam]. Ẩn Sales[Nam], Sales[Thang] khỏi Client Tools.", "C3"),
        ("3", "GM% = trung bình biên từng dòng (AVERAGEX)",
         f"{' · '.join(p(MON[m]['cu_gm_avg'], 2) for m in M14)}. Tài chính: {' · '.join(p(MON[m]['tc_gm'], 2) for m in M14)} (lệch 0,1–0,5 điểm). Khó thấy nhất.",
         "Trung bình của tỷ lệ: dòng chai trà xanh mười nghìn đồng có trọng số ngang dòng hộp quà Tết sáu trăm năm mươi nghìn đồng.",
         "Measure tổng ÷ tổng: DIVIDE([Total Revenue] - [Total COGS], [Total Revenue]).", "C4"),
        ("4", "Số đơn = đếm dòng (COUNTROWS)",
         f"{' · '.join(n(MON[m]['cu_orders_rows']) for m in M14)} đơn → AOV {' · '.join(n(MON[m]['cu_aov_rows']) for m in M14)}. "
         f"Tài chính: {' · '.join(n(MON[m]['tc_orders']) for m in M14)} đơn, AOV {' · '.join(n(MON[m]['tc_aov']) for m in M14)}.",
         "Một đơn có nhiều dòng (Line_No). Số đơn phải đếm Order_ID không trùng; AOV kéo theo sai về một nửa.",
         "Total Orders := DISTINCTCOUNT(Sales[Order_ID]); AOV = DIVIDE([Total Revenue], [Total Orders]).", "C4"),
        ("", "", "", "", "", ""),
        ("#", "Lỗi hay gặp khi làm bài", "Dấu hiệu", "Vì sao", "Cách sửa", "Tiêu chí"),
        ("a", "Nối Target với Sales (hoặc nối Target[Store_Code] với Sales[Store_Code])", "Không tạo được relationship, hoặc báo nhiều-nhiều", "Hai cột đều có giá trị trùng; hai fact chỉ gặp nhau qua lookup", "Nối Target vào Store và Calendar", "C3"),
        ("b", "Quên Mark as Date Table", "Revenue LY, YoY LFL % trống hoặc sai", "DATEADD cần bảng ngày liên tục được đánh dấu", "Calendar → Design → Mark as Date Table", "C3"),
        ("c", "Total Orders dùng COUNTROWS hoặc COUNT", "Số đơn gấp đôi, AOV một nửa", "Đếm dòng, không đếm đơn", "DISTINCTCOUNT(Sales[Order_ID])", "C4"),
        ("d", "Revenue LFL lọc bằng tên cửa hàng gõ tay", "Thêm cửa hàng mới thì measure sai", "Điều kiện nên theo thuộc tính: Store[Open_Date] < DATE(2024, 1, 1)", "CALCULATE([Total Revenue], Store[Open_Date] < DATE(2024, 1, 1))", "C4"),
        ("e", "Insight chỉ nêu con số, không nêu nguyên nhân hay hàm ý", "Câu insight một vế", "Thiếu hai trong ba thành phần", "Con số có so sánh · nguyên nhân khả dĩ · hàm ý", "C6")]
da_items = [("A1", "ĐÁP ÁN — LỖI CÀI SẴN VÀ LỖI HAY GẶP")]
da_wrap = []
for i, row in enumerate(errs, start=3):
    for j, v in enumerate(row):
        if v != "":
            da_items.append((f"{'ABCDEF'[j]}{i}", v))
da_last = 2 + len(errs)
da_sheet = cells("Dap_an_loi", da_items, {"A": 4, "B": 40, "C": 62, "D": 52, "E": 48, "F": 9}, bold=["A3:F3", "A9:F9"],
                 fill={"A3:F3": NAVY, "A9:F9": NAVY}, font_color={"A3:F3": "FFFFFF", "A9:F9": "FFFFFF"}, wrap=[f"A3:F{da_last}"])

# Insight_mau
st12 = M["st12"]
ins_text = (f"\"Tháng 1–4/2025 doanh thu cả chuỗi tăng {sp(H['ytd_yoy_t4'])} so với cùng kỳ, nhưng chỉ tăng {sp(H['lfl_ytd_yoy'])} nếu so like-for-like ({H['lfl_stores']} cửa hàng); "
            f"phần chênh đến từ ST12 (khai trương 01/06/2024, đóng góp {n(st12['rev_2025_t1_t4'])} VNĐ không có kỳ gốc). "
            f"Báo cáo cũ ghi đạt {p(MON['4']['cu_vs_target'])} target tháng 4 vì Target chưa nối Calendar, số đúng là {p(MON['4']['tc_vs_target'])}; "
            f"tháng 2 hụt ({p(MON['2']['tc_vs_target'])}) phần lớn do target lập theo cùng tháng dương lịch 2024 nên lệch mùa Tết. "
            "Từ nay mọi báo cáo dùng chung một bộ measure và KPI dictionary, và so tăng trưởng nên dùng like-for-like khi có cửa hàng mới.\"")
ins_items = [("A1", "INSIGHT MẪU (tham khảo khi chấm; học viên viết theo cách của mình)"), ("A3", ins_text),
             ("A5", "Chấm: đủ 3 thành phần (con số có so sánh · nguyên nhân khả dĩ · hàm ý), con số khớp dữ liệu. Nhật ký đối soát nêu đủ 4 lỗi gốc, mỗi lỗi có nguyên nhân và cách sửa; "
                    "Diagram View là hình ngôi sao; KPI_Dictionary có cột Giới hạn cho Revenue vs Target %, Gross Margin % và YoY.")]
ins_sheet = cells("Insight_mau", ins_items, {"A": 130}, wrap=["A3", "A5"])

# Nhat_ky mẫu (điền sẵn)
nk_rows = [
    ("% đạt target (T1–T4)", " · ".join(p(MON[m]["cu_vs_target"]) for m in M14), " · ".join(p(MON[m]["tc_vs_target"]) for m in M14), "Lệch hàng chục điểm",
     "Lỗi 1: Target chỉ nối Store, chưa nối Calendar, nên mọi tháng lấy target cả năm (3.665.000.000)",
     "Diagram View: kéo Target[Ngay_Dau_Thang] thả lên Calendar[Date]"),
    ("% đạt target (T1–T4), cùng triệu chứng với dòng trên", " · ".join(p(MON[m]["cu_vs_target"]) for m in M14), " · ".join(p(MON[m]["tc_vs_target"]) for m in M14), "Lệch hàng chục điểm",
     "Lỗi 2: Pivot lấy tháng từ Sales[Thang] (cột của bảng fact), không lọc được Target; sửa lỗi 1 mà chưa sửa lỗi này thì Target vẫn lặp",
     "Rows = Calendar[Thang], Filters = Calendar[Nam]; ẩn Sales[Nam], Sales[Thang] khỏi Client Tools"),
    ("Gross Margin % (T1–T4)", " · ".join(p(MON[m]["cu_gm_avg"], 2) for m in M14), " · ".join(p(MON[m]["tc_gm"], 2) for m in M14), "−0,1 đến −0,5 điểm",
     "Lỗi 3: trung bình biên từng dòng, dòng hàng rẻ có trọng số ngang dòng hàng đắt",
     "Measure tổng ÷ tổng: DIVIDE([Total Revenue] - [Total COGS], [Total Revenue])"),
    ("Số đơn và AOV (T1–T4)",
     "Số đơn: " + " · ".join(n(MON[m]["cu_orders_rows"]) for m in M14) + ". AOV: " + " · ".join(n(MON[m]["cu_aov_rows"]) for m in M14),
     "Số đơn: " + " · ".join(n(MON[m]["tc_orders"]) for m in M14) + ". AOV: " + " · ".join(n(MON[m]["tc_aov"]) for m in M14),
     "Số đơn gần gấp đôi, AOV khoảng một nửa",
     "Lỗi 4: đếm dòng thay vì đếm đơn (một đơn có nhiều dòng Line_No); AOV sai theo vì lấy doanh thu chia số đơn đã phồng",
     "Total Orders := DISTINCTCOUNT(Sales[Order_ID]); AOV = DIVIDE([Total Revenue], [Total Orders])"),
]
nk_ans_items = [("A1", "NHẬT KÝ ĐỐI SOÁT: BÁO CÁO KINH DOANH ↔ SỐ TÀI CHÍNH (bài làm mẫu)"),
                ("A2", "Mỗi lỗi gốc một dòng; chỉ tiêu bị kéo theo ghi chung dòng với lỗi gốc. Lỗi 1 và 2 cùng triệu chứng (% đạt target) nên tách thành hai dòng theo nguyên nhân.")]
it, last, _ = block("Nhat_ky_doi_soat", 4, nk_head, nk_rows)
nk_ans_items += it
nk_ans = cells("Nhat_ky_doi_soat", nk_ans_items, {"A": 28, "B": 40, "C": 40, "D": 22, "E": 55, "F": 55},
               bold=["A4:F4"], fill={"A4:F4": NAVY}, font_color={"A4:F4": "FFFFFF"}, wrap=[f"A5:F{last}"])
# Brief_Insight mẫu
bi_ans_items = [("A1", "BRIEF CHO SALES DIRECTOR (bài làm mẫu)"),
                ("A3", "3 câu hỏi đo lường được:"),
                ("A4", "1. Revenue vs Target % theo cửa hàng, tháng 2–4/2025, cửa hàng nào dưới 80% target (mốc so sánh: 100% target; Pivot Store_Code × Thang)."),
                ("A5", "2. YoY like-for-like của doanh thu theo tháng, tháng 1–4/2025, so với cùng tháng 2024 (sheet BaoCao_LFL)."),
                ("A6", "3. Gross Margin % theo ngành hàng, Quý 1/2025, so với biên mục tiêu 33% (lưu ý một bảng giá vốn)."),
                ("A8", "Câu insight: xem sheet Insight_mau.")]
bi_ans = cells("Brief_Insight", bi_ans_items, {"A": 130}, bold=["A3"], wrap=["A4:A6"])
# KPI_Dictionary: thêm 3 dòng LFL vào sheet đã có từ file Practice (hàng 21–23)
kd_add = [("A21", "Revenue like-for-like"), ("B21", "Revenue LFL"), ("C21", "CALCULATE([Total Revenue], Store[Open_Date] < DATE(2024, 1, 1))"), ("D21", "#,##0"),
          ("E21", "tháng"), ("F21", "region, month"), ("G21", "Finance"), ("J21", "Chỉ 11 cửa hàng mở trước 01/01/2024; loại ST12."),
          ("A22", ""), ("B22", "Revenue LFL LY"), ("C22", "CALCULATE([Revenue LFL], DATEADD(Calendar[Date], -1, YEAR))"), ("D22", "#,##0"), ("J22", "Cần Calendar là bảng ngày."),
          ("A23", "Revenue YoY like-for-like"), ("B23", "YoY LFL %"), ("C23", "IF(ISBLANK([Revenue LFL]), BLANK(), DIVIDE([Revenue LFL] - [Revenue LFL LY], [Revenue LFL LY]))"),
          ("D23", "0,0%"), ("E23", "tháng"), ("F23", "region, month"), ("G23", "Finance"), ("J23", "So sánh công bằng khi có cửa hàng mới; trống khi không có kỳ gốc.")]
kd_add = [(a, v) for a, v in kd_add if v != ""]
kd_ans = Cells("KPI_Dictionary", kd_add, wrap=["C21:C23", "J21:J23"])

ANS_SHEETS = [kd_ans, tc_sheet, nk_ans, bi_ans, kt_sheet, da_sheet, ins_sheet]


def main():
    if sys.platform != "win32":
        sys.exit("build_b5_homework.py chỉ chạy trên Windows (Excel qua COM).")
    os.makedirs(HW, exist_ok=True)
    if "--check" in sys.argv:
        import subprocess, tempfile
        for f in (HW_DE, HW_ANS):
            print("==", os.path.basename(f), "tồn tại" if os.path.exists(f) else "THIẾU")
        if os.path.exists(HW_ANS):
            _, log = run(HW_ANS, HW_ANS, checks=CHECKS)
        return
    for f in (B4_ANSWER, B5_ANSWER):
        if not os.path.exists(f):
            sys.exit(f"Không thấy file nguồn: {f}")
    print("Dựng đề (từ đáp án B4)...")
    _, log = run(B4_ANSWER, HW_DE, delete_sheets=DROP_OLD, rels=DE_RELS, measures=KD, pivots=[DE_PIVOT], cells=DE_SHEETS,
                 first_sheets=["De_bai", "BaoCao_KinhDoanh", "So_lieu_Tai_chinh", "Nhat_ky_doi_soat", "KPI_Dictionary", "Brief_Insight"],
                 active_sheet="De_bai")
    for line in log:
        print("  ·", line)
    print("Dựng đáp án (từ đáp án Practice B5)...")
    result, log = run(B5_ANSWER, HW_ANS, delete_sheets=DROP_PRACTICE, measures=LFL, pivots=ANS_PIVOTS, cells=ANS_SHEETS, checks=CHECKS,
                      first_sheets=["Kiem_tra", "BaoCao_KinhDoanh", "BaoCao_LFL", "So_lieu_Tai_chinh", "Nhat_ky_doi_soat", "KPI_Dictionary",
                                    "Brief_Insight", "Dap_an_loi", "Insight_mau"], active_sheet="Kiem_tra")
    for line in log:
        print("  ·", line)
    bad = [(kt_rows[i][0], v) for i, v in enumerate(result) if v != "✔"]
    print(f"Kiem_tra BTVN: {len(result) - len(bad)}/{len(result)} ✔")
    for label, v in bad:
        print(f"  ✘ {label}: {v!r}")
    if bad:
        sys.exit(1)


if __name__ == "__main__":
    main()
