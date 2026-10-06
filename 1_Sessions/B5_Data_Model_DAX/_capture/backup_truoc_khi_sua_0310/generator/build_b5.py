r"""Dựng file đáp án Buổi 5 (Data Model & DAX): Practice/B5_MCI_Mart_Answer.xlsx — file khởi đầu Buổi 6.

Bắt đầu từ đáp án B4 (copy nguyên file: Data Model 6 bảng, các sheet B1–B4, KPI_SUMIFS), rồi THÊM đúng những gì
dựng trên lớp (Slide_Script_B5.md):
  6 relationship   Sales → Product, Store, Customer, Calendar (Order_Date → Date); Target → Store, Calendar (Ngay_Dau_Thang → Date)
  14 measure       bảng "Measure dựng trong buổi" của kịch bản (trừ Revenue Mien Nam — chỉ để demo, xóa trên lớp)
  7 Pivot          From Data Model: PV_Loc, PV_Target, PV_Target_CuaHang, PV_YoY, PV_GM_Nam, PV_KPI_Thang, PV_DoiSoat_B2
  DoiSoat_B2       lưới CUBEVALUE Miền × Tháng × 4 KPI đặt đúng vị trí như KPI_SUMIFS, cạnh đó là chênh lệch (64 ô)
  KPI_Dictionary   7 KPI (Course Plan, Phụ lục C) + measure hỗ trợ, kèm DAX và giới hạn
  Kiem_tra_B5      24 dòng: mỗi relationship được kiểm bằng một CUBEVALUE cần đúng quan hệ đó, đối soát 64 ô, KPI, target, YTD/LY

Excel không có API cho calculated column, Mark as Date Table, Hide from Client Tools, nên:
  - Total COGS viết bằng SUMX(Sales, Sales[Quantity] * RELATED(Product[Unit_Cost])) — cùng kết quả với cột COGS
    (calculated column) + SUM(Sales[COGS]) làm trên lớp (Slide 15, 17); KPI_Dictionary ghi chú điều này.
  - Mark as Date Table, Hide from Client Tools: bước làm tay sau khi dựng (script in hướng dẫn). Nếu các dòng
    YTD / LY của Kiem_tra_B5 ra ✘ thì làm Mark as Date Table, lưu, rồi chạy lại với --check.

KHÔNG Refresh: query Power Query của file trỏ tới thư mục máy dựng B4 (Data_Sales chỉ còn 4 file). Dữ liệu đã có
sẵn trong Data Model (Sales 30.247 dòng, tới 04/2025).

Số liệu đáp án đọc từ build_b5_meta.json (chạy build_b5_meta.py trước). CHỈ CHẠY TRÊN WINDOWS (Excel qua COM).

Chạy từ thư mục MCI Excel:
    python 2_Dataset\_generator\build_b5.py            # dựng lại đáp án từ đáp án B4
    python 2_Dataset\_generator\build_b5.py --check    # chỉ mở file đáp án, tính lại, đọc Kiem_tra_B5 (sau khi làm tay)
    python 2_Dataset\_generator\build_b5.py --print-ps # in mã PowerShell sẽ chạy (xem trên máy bất kỳ), không mở Excel
Trước khi chạy: đóng B4_MCI_Mart_Answer.xlsx và B5_MCI_Mart_Answer.xlsx nếu đang mở.
"""
import json, os, shutil, sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from datamodel_com import Rel, Measure, CubePivot, Cells, build_ps, run

for s_ in (sys.stdout, sys.stderr):
    if hasattr(s_, "reconfigure"):
        s_.reconfigure(encoding="utf-8", errors="replace")

ROOT = os.path.normpath(os.environ.get("MCI_ROOT") or os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
GEN = os.path.dirname(os.path.abspath(__file__))
B4_ANSWER = os.path.join(ROOT, "1_Sessions", "B4_Power_Query", "Practice", "B4_MCI_Mart_Answer.xlsx")
PRACTICE = os.path.join(ROOT, "1_Sessions", "B5_Data_Model_DAX", "Practice")
ANSWER = os.path.join(PRACTICE, "B5_MCI_Mart_Answer.xlsx")
M = json.load(open(os.path.join(GEN, "build_b5_meta.json"), encoding="utf-8"))

# ---------------------------------------------------------------- Data Model
rels = [
    Rel("Sales", "Product_Code", "Product", "Product_Code"),
    Rel("Sales", "Store_Code", "Store", "Store_Code"),
    Rel("Sales", "Customer_ID", "Customer", "Customer_ID"),
    Rel("Sales", "Order_Date", "Calendar", "Date"),
    Rel("Target", "Store_Code", "Store", "Store_Code"),
    Rel("Target", "Ngay_Dau_Thang", "Calendar", "Date"),
]

measures = [
    Measure("Total Revenue", "Sales", "SUM(Sales[Revenue])", "int", "Tổng doanh thu = SUM(Quantity × Unit_Price). Measure cơ sở."),
    Measure("Total COGS", "Sales", "SUMX(Sales, Sales[Quantity] * RELATED(Product[Unit_Cost]))", "int",
            "Tổng giá vốn. Trên lớp: cột COGS (calculated column) + SUM(Sales[COGS]); cùng kết quả. Giá vốn: một bảng Product[Unit_Cost] cho mọi năm."),
    Measure("Gross Margin %", "Sales", "DIVIDE([Total Revenue] - [Total COGS], [Total Revenue])", "pct1",
            "(Doanh thu − giá vốn) ÷ doanh thu, tổng chia tổng. Giới hạn: một bảng giá vốn cho mọi năm → GM% 2024 thấp giả."),
    Measure("Revenue All Stores", "Sales", "CALCULATE([Total Revenue], ALL(Store))", "int", "Doanh thu mọi cửa hàng (bỏ lọc Store), giữ các bộ lọc khác."),
    Measure("% Revenue Share", "Sales", "DIVIDE([Total Revenue], [Revenue All Stores])", "pct1", "Tỷ trọng doanh thu trong tổng các cửa hàng."),
    Measure("Total Target", "Target", "SUM(Target[Target])", "int", "Target doanh thu. Grain tháng × cửa hàng: chỉ lọc theo store, region, month."),
    Measure("Revenue vs Target %", "Sales", "DIVIDE([Total Revenue], [Total Target])", "pct1",
            "Doanh thu ÷ target. KHÔNG dùng theo product/category (Target không có chiều sản phẩm)."),
    Measure("Variance", "Sales", "IF(ISBLANK([Total Revenue]), BLANK(), [Total Revenue] - [Total Target])", "int",
            "Doanh thu − target; để trống khi chưa có doanh thu (tháng tương lai)."),
    Measure("Revenue YTD", "Sales", "CALCULATE([Total Revenue], DATESYTD(Calendar[Date]))", "int", "Lũy kế từ đầu năm (Calendar)."),
    Measure("Revenue LY", "Sales", "CALCULATE([Total Revenue], DATEADD(Calendar[Date], -1, YEAR))", "int", "Doanh thu cùng kỳ năm trước."),
    Measure("Revenue YoY %", "Sales", "IF(ISBLANK([Total Revenue]), BLANK(), DIVIDE([Total Revenue] - [Revenue LY], [Revenue LY]))", "pct1",
            "Tăng trưởng so với cùng kỳ. Trống khi chưa có doanh thu hoặc không có kỳ gốc (ST12 tháng 1–5/2025). Cả chuỗi gồm cửa hàng mới."),
    Measure("Total Orders", "Sales", "DISTINCTCOUNT(Sales[Order_ID])", "int", "Số đơn = đếm Order_ID không trùng (một đơn nhiều dòng)."),
    Measure("Avg Order Value", "Sales", "DIVIDE([Total Revenue], [Total Orders])", "int", "Giá trị đơn trung bình = doanh thu ÷ số đơn."),
    Measure("Active Customers", "Sales", "DISTINCTCOUNT(Sales[Customer_ID])", "int", "Số khách có mua. Không cộng dồn qua các nhóm."),
]

# ---------------------------------------------------------------- Pivot từ Data Model
NAM, THANG, QUY = "[Calendar].[Nam]", "[Calendar].[Thang]", "[Calendar].[Quy]"
REGION, STORE = "[Store].[Region]", "[Store].[Store_Code]"
T14 = ["1", "2", "3", "4"]
pivots = [
    CubePivot("PV_Loc", "PV_Loc", rows=[REGION], values=["Total Revenue", "Revenue All Stores", "% Revenue Share"],
              filters={NAM: ["2025"], QUY: ["1"]}, title="Filter context · CALCULATE · ALL (Slide 19, 22, 23)",
              note="Quý 1/2025 theo miền: % Revenue Share = Pct_Share của Buổi 1."),
    CubePivot("PV_Target", "PV_Target", rows=[THANG], values=["Total Revenue", "Total Target", "Revenue vs Target %", "Variance"],
              filters={NAM: ["2025"]}, title="Revenue vs Target % theo tháng 2025 (Slide 24)",
              note="Tháng 5–12 chỉ có target; Variance để trống. T1 vượt, T2 hụt: target lập theo cùng tháng 2024 × 110%, Tết lệch tháng."),
    CubePivot("PV_Target_CuaHang", "PV_Target_CuaHang", rows=[STORE], cols=[THANG], values=["Revenue vs Target %"],
              filters={NAM: ["2025"]}, row_items={THANG: T14}, title="Revenue vs Target % — cửa hàng × tháng 1–4/2025 (Slide 25)",
              note="ST07 tháng 3 (cải tạo 10–31/03) và tháng 4 (hồi phục)."),
    CubePivot("PV_YoY", "PV_YoY", rows=[STORE], values=["Total Revenue", "Revenue LY", "Revenue YoY %"],
              filters={NAM: ["2025"], THANG: T14}, title="YoY tháng 1–4/2025 theo cửa hàng (Slide 28)",
              note="ST12 khai trương 01/06/2024: không có kỳ gốc → LY và YoY trống. Like-for-like: bỏ chọn ST12 ở Store_Code."),
    CubePivot("PV_GM_Nam", "PV_GM_Nam", rows=[NAM], values=["Gross Margin %"], title="Gross Margin % theo năm (Slide 30, câu chuyện #5)",
              note="Giá vốn chỉ có một bảng cho cả hai năm, giá bán 2025 cao hơn 2024 → GM% 2024 thấp giả."),
    CubePivot("PV_KPI_Thang", "PV_KPI_Thang", rows=[THANG],
              values=["Total Revenue", "Total Orders", "Avg Order Value", "Gross Margin %", "Revenue YoY %", "Active Customers", "Revenue YTD"],
              filters={NAM: ["2025"]}, row_items={THANG: T14}, title="6 KPI theo tháng 1–4/2025 (Thực hành 1, Slide 31)"),
    CubePivot("PV_DoiSoat_B2", "PV_DoiSoat_B2", rows=[REGION], cols=[THANG],
              values=["Total Revenue", "Total Orders", "Avg Order Value", "Gross Margin %"],
              filters={NAM: ["2025"], QUY: ["1"]}, title="Báo cáo Miền × Tháng Quý 1/2025 từ Data Model (Thực hành 2, Slide 32)",
              note="Bản Data Model của PV_MienThang (B3). Đối soát từng ô với KPI_SUMIFS: sheet DoiSoat_B2."),
]


# ---------------------------------------------------------------- CUBEVALUE
def mem(field_, key):
    return f"{field_}.&[{key}]"


def cv(measure, *members):
    args = ",".join(f'"{x}"' for x in [f"[Measures].[{measure}]", *members])
    return f'=CUBEVALUE("ThisWorkbookDataModel",{args})'


Y25, Q1 = mem(NAM, 2025), mem(QUY, 1)
REG = ["Bắc", "Trung", "Nam"]
COLS = ["B", "C", "D", "E"]                     # tháng 1, 2, 3, Quý 1 — đúng vị trí của KPI_SUMIFS
DCOLS = ["G", "H", "I", "J"]                    # chênh lệch
BLOCKS = [("Doanh thu (VNĐ)", "Total Revenue", 4, "#,##0"), ("Số đơn", "Total Orders", 11, "#,##0"),
          ("Giá trị đơn trung bình – AOV (VNĐ)", "Avg Order Value", 18, "#,##0"), ("Gross Margin %", "Gross Margin %", 25, "0.00%")]

ds = [("A1", "ĐỐI SOÁT DATA MODEL ↔ KPI_SUMIFS (Buổi 2) — Miền × Tháng, Quý 1/2025"),
      ("A2", "Cột B–E: CUBEVALUE đọc measure trong Data Model, đặt đúng vị trí ô như sheet KPI_SUMIFS. Cột G–J: chênh lệch với KPI_SUMIFS (phải bằng 0).")]
ds_fmt, ds_bold = {}, ["A4", "A11", "A18", "A25"]
for title, meas, r0, nf in BLOCKS:
    ds += [(f"A{r0}", title), (f"A{r0 + 1}", "Miền"), (f"B{r0 + 1}", 1), (f"C{r0 + 1}", 2), (f"D{r0 + 1}", 3), (f"E{r0 + 1}", "Quý 1"),
           (f"G{r0 + 1}", "Chênh 1"), (f"H{r0 + 1}", "Chênh 2"), (f"I{r0 + 1}", "Chênh 3"), (f"J{r0 + 1}", "Chênh Quý 1")]
    ds_bold.append(f"A{r0 + 1}:J{r0 + 1}")
    for i, reg in enumerate(REG + ["Tổng"]):
        r = r0 + 2 + i
        ds.append((f"A{r}", reg))
        for j, col in enumerate(COLS):
            members = [Y25] + ([Q1] if j == 3 else [mem(THANG, j + 1)]) + ([] if reg == "Tổng" else [mem(REGION, reg)])
            ds.append((f"{col}{r}", cv(meas, *members)))
            ds.append((f"{DCOLS[j]}{r}", f"={col}{r}-KPI_SUMIFS!{col}{r}"))
    ds_fmt[f"B{r0 + 2}:E{r0 + 5}"] = nf
    ds_fmt[f"G{r0 + 2}:J{r0 + 5}"] = "0.000000" if nf == "0.00%" else "#,##0.00"
doisoat = Cells("DoiSoat_B2", ds, widths={"A": 34, "B": 15, "C": 15, "D": 15, "E": 15, "F": 3, "G": 12, "H": 12, "I": 12, "J": 13},
                numfmt=ds_fmt, bold=ds_bold, title="x")

# ---------------------------------------------------------------- KPI dictionary
kpi_rows = [
    ("KPI", "Measure", "DAX", "Định dạng", "Grain", "Dimensions", "Owner", "Target", "Warning / Critical", "Giới hạn / ghi chú"),
    ("Total Revenue", "Total Revenue", "SUM(Sales[Revenue])", "#,##0", "tháng", "region, store, category, month", "Sales Director",
     "Bằng Total Target", "90% / 80%", "Measure cơ sở cho mọi KPI doanh thu."),
    ("Total Orders", "Total Orders", "DISTINCTCOUNT(Sales[Order_ID])", "#,##0", "tháng", "region, store, month", "Sales Director",
     "1.000 đơn/tháng", "90% / 80%", "Đếm Order_ID không trùng; COUNTROWS đếm dòng là sai."),
    ("Avg Order Value", "Avg Order Value", "DIVIDE([Total Revenue], [Total Orders])", "#,##0", "tháng", "region, store, month", "Sales Director",
     "300.000", "90% / 80%", "Lọc theo ngành: giá trị phần ngành đó trong mỗi đơn có mua ngành đó."),
    ("Gross Margin %", "Gross Margin %", "DIVIDE([Total Revenue] - [Total COGS], [Total Revenue])", "0,0%", "tháng", "category, region, month", "Finance",
     "33%", "90% / 80%", "Một bảng giá vốn (Product[Unit_Cost]) cho mọi năm → GM% 2024 thấp giả, không so được giữa các năm."),
    ("Revenue YoY Growth", "Revenue YoY %", "IF(ISBLANK([Total Revenue]), BLANK(), DIVIDE([Total Revenue] - [Revenue LY], [Revenue LY]))", "0,0%",
     "tháng", "region, category, month", "Finance", "+10%", "50% / 20%", "Cả chuỗi gồm ST12 chưa có kỳ gốc; so like-for-like khi cần."),
    ("Active Customers", "Active Customers", "DISTINCTCOUNT(Sales[Customer_ID])", "#,##0", "tháng", "region, segment, month", "Marketing",
     "700 KH/tháng", "90% / 80%", "Không cộng dồn: tổng các miền lớn hơn dòng tổng."),
    ("Revenue vs Target %", "Revenue vs Target %", "DIVIDE([Total Revenue], [Total Target])", "0,0%", "tháng × cửa hàng", "store, region, month",
     "Sales Director", "100%", "90% / 80%", "KHÔNG theo product/category. Target lập theo cùng tháng 2024 × 110% → lệch mùa Tết (T1, T2)."),
    ("", "", "", "", "", "", "", "", "", ""),
    ("Measure hỗ trợ", "", "", "", "", "", "", "", "", ""),
    ("", "Total COGS", "SUMX(Sales, Sales[Quantity] * RELATED(Product[Unit_Cost]))", "#,##0", "", "", "", "", "",
     "Trên lớp: cột COGS (calculated column) + SUM(Sales[COGS]). Excel không cho script tạo calculated column; SUMX cho cùng kết quả."),
    ("", "Total Target", "SUM(Target[Target])", "#,##0", "tháng × cửa hàng", "", "", "", "", "Nằm ở bảng Target."),
    ("", "Variance", "IF(ISBLANK([Total Revenue]), BLANK(), [Total Revenue] - [Total Target])", "#,##0", "", "", "", "", "", "Chặn tháng chưa có dữ liệu."),
    ("", "Revenue All Stores", "CALCULATE([Total Revenue], ALL(Store))", "#,##0", "", "", "", "", "", "Mẫu số cho % Revenue Share."),
    ("", "% Revenue Share", "DIVIDE([Total Revenue], [Revenue All Stores])", "0,0%", "", "", "", "", "", ""),
    ("", "Revenue YTD", "CALCULATE([Total Revenue], DATESYTD(Calendar[Date]))", "#,##0", "", "", "", "", "", "Cần Calendar liên tục; hàng/cột lấy từ Calendar."),
    ("", "Revenue LY", "CALCULATE([Total Revenue], DATEADD(Calendar[Date], -1, YEAR))", "#,##0", "", "", "", "", "", ""),
]
kd = [("A1", "KPI DICTIONARY — MCI Mart (Course Plan, Phụ lục C)"),
      ("A2", "Mỗi KPI một measure, một định nghĩa. Owner, target, ngưỡng theo Phụ lục C; mẫu KPI Definition Card ở Buổi 2.")]
for i, row in enumerate(kpi_rows):
    for j, v in enumerate(row):
        if v != "":
            kd.append((f"{'ABCDEFGHIJ'[j]}{4 + i}", v))
kpi_dict = Cells("KPI_Dictionary", kd, widths={"A": 20, "B": 20, "C": 52, "D": 10, "E": 14, "F": 28, "G": 15, "H": 15, "I": 14, "J": 60},
                 bold=["A4:J4", "A13", "A5:A11"], fill={"A4:J4": "203864"}, font_color={"A4:J4": "FFFFFF"}, wrap=["C5:C21", "F5:F21", "J5:J21"], title="x")

# ---------------------------------------------------------------- Kiem_tra_B5
q1, sm = M["q1"], M["store_month"]
MONEY, PCT = 0.5, 0.00005
kt_rows = [  # (chỉ tiêu, công thức, đáp án, sai số, định dạng)
    ("Doanh thu Quý 1/2025 (Total Revenue)", cv("Total Revenue", Y25, Q1), q1["rev"], MONEY, "#,##0"),
    ("Đối soát KPI_SUMIFS: số ô khớp (4 KPI × 3 miền + tổng × 3 tháng + Quý 1)",
     "=" + "+".join(f"SUMPRODUCT(--(ROUND(ABS(DoiSoat_B2!G{r0 + 2}:J{r0 + 5}),6)=0))" for _, _, r0, _ in BLOCKS), M["recon_cells"], 0, "0"),
    ("Bắc × T3/2025 — cần Sales→Store, Sales→Calendar", cv("Total Revenue", Y25, mem(THANG, 3), mem(REGION, "Bắc")), M["bac_t3_2025"], MONEY, "#,##0"),
    ("Bánh kẹo, Quý 1 — cần Sales→Product", cv("Total Revenue", Y25, Q1, mem("[Product].[Category]", "Bánh kẹo")),
     M["aov_cat_q1"]["Bánh kẹo"]["rev"], MONEY, "#,##0"),
    ("Active Customers VIP, Quý 1 — cần Sales→Customer", cv("Active Customers", Y25, Q1, mem("[Customer].[Segment]", "VIP")),
     M["hw"]["seg_cust_q1"]["VIP"], 0, "#,##0"),
    ("Total Target T1/2025 — cần Target→Calendar", cv("Total Target", Y25, mem(THANG, 1)), M["target_by_month"]["1"], MONEY, "#,##0"),
    ("Total Target miền Bắc 2025 — cần Target→Store", cv("Total Target", Y25, mem(REGION, "Bắc")), M["target_by_region"]["Bắc"], MONEY, "#,##0"),
    ("Gross Margin % Quý 1/2025", cv("Gross Margin %", Y25, Q1), q1["gm"], PCT, "0.00%"),
    ("Total Orders Quý 1/2025", cv("Total Orders", Y25, Q1), q1["orders"], 0, "#,##0"),
    ("Avg Order Value Quý 1/2025", cv("Avg Order Value", Y25, Q1), q1["aov"], MONEY, "#,##0"),
    ("Active Customers Quý 1/2025", cv("Active Customers", Y25, Q1), q1["cust"], 0, "#,##0"),
    ("Active Customers Quý 1: Bắc + Trung + Nam (lớn hơn dòng tổng — không cộng dồn)",
     "=" + "+".join(cv("Active Customers", Y25, Q1, mem(REGION, r))[1:] for r in REG), M["cust_sum_regions_q1"], 0, "#,##0"),
    ("% Revenue Share miền Bắc, Quý 1 (= Pct_Share B1)", cv("% Revenue Share", Y25, Q1, mem(REGION, "Bắc")), M["share_q1"]["Bắc"], PCT, "0.00%"),
    ("Revenue All Stores ở ô miền Bắc, Quý 1 (ALL bỏ lọc Store)", cv("Revenue All Stores", Y25, Q1, mem(REGION, "Bắc")), q1["rev"], MONEY, "#,##0"),
    ("Revenue vs Target % T1/2025", cv("Revenue vs Target %", Y25, mem(THANG, 1)), M["vs_target_month"]["1"], PCT, "0.00%"),
    ("Revenue vs Target % T2/2025", cv("Revenue vs Target %", Y25, mem(THANG, 2)), M["vs_target_month"]["2"], PCT, "0.00%"),
    ("Revenue vs Target % ST07 T3/2025 (câu chuyện #2)", cv("Revenue vs Target %", Y25, mem(THANG, 3), mem(STORE, "ST07")), sm["ST07"]["3"]["pct"], PCT, "0.00%"),
    ("Variance T5/2025 để trống (0 = trống)", f'=IF({cv("Variance", Y25, mem(THANG, 5))[1:]}="",0,1)', 0, 0, "0"),
    ("Revenue YTD T4/2025 (cần Calendar là bảng ngày)", cv("Revenue YTD", Y25, mem(THANG, 4)), M["ytd_2025"]["4"], MONEY, "#,##0"),
    ("Revenue LY T4/2025 (= tháng 4/2024)", cv("Revenue LY", Y25, mem(THANG, 4)), M["rev_2024"]["4"], MONEY, "#,##0"),
    ("Revenue YoY % T1/2025", cv("Revenue YoY %", Y25, mem(THANG, 1)), M["yoy_month"]["1"], PCT, "0.00%"),
    ("Revenue YoY % ST12 T1/2025 để trống — không có kỳ gốc (0 = trống)",
     f'=IF({cv("Revenue YoY %", Y25, mem(THANG, 1), mem(STORE, "ST12"))[1:]}="",0,1)', 0, 0, "0"),
    ("Gross Margin % năm 2024 (câu chuyện #5)", cv("Gross Margin %", mem(NAM, 2024)), M["gm_year"]["2024"], PCT, "0.00%"),
    ("AOV Bánh kẹo, Quý 1 (kiểm tay: doanh thu ÷ số đơn)", cv("Avg Order Value", Y25, Q1, mem("[Product].[Category]", "Bánh kẹo")),
     M["aov_cat_q1"]["Bánh kẹo"]["aov"], MONEY, "#,##0"),
]
kt = [("A1", "KIỂM TRA / ĐỐI SOÁT — Brick output B5 (Data Model & DAX)"),
      ("A2", "Cột C: CUBEVALUE đọc measure trong Data Model. Cột D: đáp án tính lại từ dữ liệu (build_b5_meta.json). "
             "Mỗi relationship được kiểm bằng một dòng chỉ đúng khi có quan hệ đó. YTD / LY cần Calendar là bảng ngày (Mark as Date Table)."),
      ("A4", "Chỉ tiêu"), ("C4", "Kết quả"), ("D4", "Đáp án"), ("E4", "Khớp?")]
kt_fmt = {}
for i, (label, formula, expected, tol, nf) in enumerate(kt_rows, start=5):
    kt += [(f"A{i}", label), (f"C{i}", formula), (f"D{i}", float(expected)),
           (f"E{i}", f'=IF(ISNUMBER(C{i}),IF(ABS(C{i}-D{i})<={tol},"✔","✘"),"✘")')]
    kt_fmt[f"C{i}:D{i}"] = nf
last = 4 + len(kt_rows)
kiemtra = Cells("Kiem_tra_B5", kt, widths={"A": 78, "B": 2, "C": 18, "D": 18, "E": 8}, numfmt=kt_fmt,
                bold=["A4:E4"], fill={"A4:E4": "203864"}, font_color={"A4:E4": "FFFFFF"}, title="x")
checks = [f"Kiem_tra_B5!E{i}" for i in range(5, last + 1)]

ops = dict(rels=rels, measures=measures, pivots=pivots, cells=[doisoat, kpi_dict, kiemtra], checks=checks, active_sheet="Kiem_tra_B5")

if "--print-ps" in sys.argv:
    print(build_ps("<work>.xlsx", "<work>.json", **ops))
    sys.exit(0)
if sys.platform != "win32":
    sys.exit("build_b5.py chỉ chạy trên Windows (Excel qua COM). Dùng --print-ps để xem mã PowerShell trên máy khác.")

if "--check" in sys.argv:
    if not os.path.exists(ANSWER):
        sys.exit(f"Chưa có {ANSWER}")
    result, log = run(ANSWER, ANSWER, checks=checks)
else:
    if not os.path.exists(B4_ANSWER):
        sys.exit(f"Không thấy file khởi đầu: {B4_ANSWER}")
    os.makedirs(PRACTICE, exist_ok=True)
    print(f"Dựng {ANSWER}\n  từ {B4_ANSWER}\nExcel qua COM — có thể mất vài phút...")
    result, log = run(B4_ANSWER, ANSWER, **ops)
for line in log:
    print("  ·", line)
bad = [(kt_rows[i][0], v) for i, v in enumerate(result) if v != "✔"]
print(f"Kiem_tra_B5: {len(result) - len(bad)}/{len(result)} ✔")
for label, v in bad:
    print(f"  ✘ {label}: {v!r}")
print("""
Bước làm tay sau khi dựng (Excel không có API cho các việc này), trên Windows:
  1. Mở Practice/B5_MCI_Mart_Answer.xlsx → Data → Manage Data Model.
  2. Tab Calendar → Design → Mark as Date Table → Date → OK.
  3. Data View: ẩn khỏi Client Tools — Sales: Nam, Thang, Order_Date, Store_Code, Product_Code, Customer_ID;
     Target: Store_Code, Thang, Ngay_Dau_Thang (chuột phải tiêu đề cột → Hide from Client Tools).
  4. Lưu, đóng, rồi: python 2_Dataset\\_generator\\build_b5.py --check""")
if bad:
    sys.exit(1)
