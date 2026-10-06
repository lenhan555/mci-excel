r"""Dựng file đáp án Buổi 4 (Power Query): B4_MCI_Mart_Answer.xlsx — file khởi đầu Buổi 5.

Bắt đầu từ đáp án B3 (copy nguyên file, giữ mọi sheet tblSales/KPI_SUMIFS/PV_.../Kiem_tra), rồi THÊM vào đúng các
query dựng trên lớp (Slide_Script_B4.md, bản 1.3):
  DataFolder  tham số (Manage Parameters): thư mục Practice của máy dựng file, có dấu \ ở cuối
  Sales       From Folder DataFolder & "Data_Sales" → Combine & Transform Data, kèm 4 query trợ giúp như khi bấm trên
              giao diện (Sample File, Parameter1, Transform Sample File, Transform File — nhóm "Transform File from
              Data_Sales" / "Helper Queries"; Sample File cũng đọc DataFolder) → bỏ Source.Name, Trim Store_Code, Using
              Locale vi-VN cho Order_Date và Unit_Price, Remove Duplicates, Revenue / Nam / Thang. Load Table (sheet
              Sales_PQ) + Data Model, tắt background refresh
  Product     B4_Lookups.xlsx, Remove Duplicates theo Product_Code (P015 trùng)       Connection Only + Data Model
  Store, Customer                                                                      Connection Only + Data Model
  Target      B4_Target_2025.xlsx, Unpivot → Store_Code | Thang (số) | Target | Ngay_Dau_Thang    Connection + Data Model
  Calendar    List.Dates 01/01/2024–31/12/2025 (731 ngày) → To Table → Add Column → Date → Year / Month / Quarter /
              Name of Day → đổi tên Nam / Thang / Quy / Thu (tên thứ theo locale: máy tiếng Anh ra Monday)
                                                                                       Connection Only + Data Model
Pivot PV_Refresh (Rows Nam, Thang; Values So_dong = Count of Order_ID, Doanh_thu = Sum of Revenue; Filters Store_Code)
và sheet Kiem_tra_B4 (10 dòng).

Đáp án thể hiện trạng thái SAU Thực hành 1 (đã thả file tháng 4): lúc dựng, script chép tạm
Practice/New_Month/Sales_2025-04.csv vào Practice/Data_Sales/, Refresh, rồi xóa lại, để thư mục Data_Sales phát cho
học viên vẫn chỉ có 4 file. Refresh file đáp án với thư mục 4 file thì ra 28.328 dòng (trạng thái trước Thực hành 1).

openpyxl không tạo được Power Query — toàn bộ việc thêm query/load/Pivot làm bằng Excel qua COM (PowerShell),
xem `powerquery_com.py`. COM không đặt được nhóm query và liên kết Transform File ↔ Transform Sample File, nên sau
lần dựng đầu script gắn hai thứ đó vào gói DataMashup (`set_query_groups`), rồi mở lại bằng Excel để Refresh All và
lưu. CHỈ CHẠY ĐƯỢC TRÊN WINDOWS, cần Excel đã cài Power Query/Power Pivot.

Số liệu đối soát: build_b4_meta.json (build_b4_homework.py) — Sales 30.247 dòng, 4.436.849.500; tháng 4/2025
1.919 dòng; Product 40, Store 12, Customer 1.500, Target 144, Calendar 731.

File mốc (--moc): Practice/B4_MCI_Mart_Moc_ThucHanh.xlsx — cùng các query, load và Pivot như đáp án nhưng ở trạng thái
TRƯỚC Thực hành 1 (cuối slide 27: Data_Sales 4 file, Sales 28.328 dòng). Dùng khi file của học viên bị lỗi hoặc chưa
làm kịp: học viên sửa DataFolder, Refresh All rồi làm Thực hành 1 như cả lớp. Không có Kiem_tra_B4; thay bằng sheet
Huong_dan_Moc (đứng đầu, file mở ở sheet này): cách dùng + bảng tự kiểm trạng thái mốc, số lấy từ build_b4_meta.json.

Chạy từ thư mục MCI Excel:
    Windows:  python 2_Dataset\_generator\build_b4.py          # đáp án
              python 2_Dataset\_generator\build_b4.py --moc    # file mốc (Data_Sales phải đúng 4 file)
    Thêm --print-m để chỉ in mã M của từng query ra màn hình (kiểm tra cú pháp), không mở Excel.
Trước khi chạy: đóng B4_MCI_Mart_Answer.xlsx / B4_MCI_Mart_Moc_ThucHanh.xlsx nếu đang mở trong Excel.
"""
import json, os, re, shutil, sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from powerquery_com import Query, TableLoad, ModelLoad, Pivot, KiemTra, run, set_query_groups, read_mashup

for s_ in (sys.stdout, sys.stderr):
    if hasattr(s_, "reconfigure"):
        s_.reconfigure(encoding="utf-8", errors="replace")

ROOT = os.path.normpath(os.environ.get("MCI_ROOT") or os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
B3_ANSWER = os.path.join(ROOT, "1_Sessions", "B3_PivotTable", "Practice", "B3_MCI_Mart_Q1_2025_Answer.xlsx")
B4 = os.path.join(ROOT, "1_Sessions", "B4_Power_Query")
PRACTICE = os.path.join(B4, "Practice")
DATA_SALES = os.path.join(PRACTICE, "Data_Sales")
NEW_MONTH = os.path.join(PRACTICE, "New_Month")
LOOKUPS = os.path.join(PRACTICE, "B4_Lookups.xlsx")
TARGET_FILE = os.path.join(PRACTICE, "B4_Target_2025.xlsx")
ANSWER = os.path.join(PRACTICE, "B4_MCI_Mart_Answer.xlsx")
MOC = os.path.join(PRACTICE, "B4_MCI_Mart_Moc_ThucHanh.xlsx")
META = os.path.join(os.path.dirname(os.path.abspath(__file__)), "build_b4_meta.json")
MOC_MODE = "--moc" in sys.argv

if not os.path.exists(B3_ANSWER):
    sys.exit(f"Không thấy file khởi đầu: {B3_ANSWER}")
for f in (LOOKUPS, TARGET_FILE):
    if not os.path.exists(f):
        sys.exit(f"Không thấy file: {f}")


# ---------------------------------------------------------------- M code ---

APRIL = os.path.join(NEW_MONTH, "Sales_2025-04.csv")
for f in [os.path.join(DATA_SALES, n) for n in ("Sales_2024.csv", "Sales_2025-01.csv", "Sales_2025-02.csv", "Sales_2025-03.csv")] + [APRIL]:
    if not os.path.exists(f):
        sys.exit(f"Không thấy file: {f}")
extra = sorted(set(os.listdir(DATA_SALES)) - {"Sales_2024.csv", "Sales_2025-01.csv", "Sales_2025-02.csv", "Sales_2025-03.csv", "Sales_2025-04.csv", "desktop.ini"})
if extra:
    sys.exit(f"Thư mục Data_Sales có file lạ, hãy dọn trước khi dựng: {extra}")
if MOC_MODE and os.path.exists(os.path.join(DATA_SALES, "Sales_2025-04.csv")):
    sys.exit("File mốc là trạng thái trước Thực hành 1: xóa Data_Sales/Sales_2025-04.csv (bản gốc ở New_Month/) rồi chạy lại.")

M_DATAFOLDER = '"' + PRACTICE.rstrip("\\/") + os.sep + '" meta [IsParameterQuery = true, Type = "Text", IsParameterQueryRequired = true]'

# Sales dựng đúng như thao tác trên lớp (Slide 21): From Folder → Combine & Transform Data (File mẫu: First file).
# Combine tạo 4 query trợ giúp; mã M dưới đây chép từ file Excel 365 tạo bằng giao diện, chỉ thay chuỗi đường dẫn
# bằng tham số DataFolder ở CẢ Sales và Sample File (Sample File cũng đọc thư mục — sót nó thì mở trên máy khác vẫn lỗi).
M_SAMPLE_FILE = """let
    Source = Folder.Files(DataFolder & "Data_Sales"),
    Navigation1 = Source{0}[Content]
in
    Navigation1"""

M_PARAMETER1 = '#"Sample File" meta [IsParameterQuery=true, BinaryIdentifier=#"Sample File", Type="Binary", IsParameterQueryRequired=true]'

M_TRANSFORM_SAMPLE_FILE = """let
    Source = Csv.Document(Parameter1,[Delimiter=",", Columns=8, Encoding=65001, QuoteStyle=QuoteStyle.None]),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true])
in
    #"Promoted Headers\""""

M_TRANSFORM_FILE = """let
    Source = (Parameter1) => let
        Source = Csv.Document(Parameter1,[Delimiter=",", Columns=8, Encoding=65001, QuoteStyle=QuoteStyle.None]),
        #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true])
    in
        #"Promoted Headers"
in
    Source"""

# Sau Combine: xóa bước Changed Type tự động, bỏ Source.Name, Trim, Using Locale vi-VN, kiểu dữ liệu,
# Remove Duplicates (Ctrl+A), Custom Column Revenue, Add Column → Date → Year / Month rồi đổi tên Nam / Thang.
M_SALES = """let
    Source = Folder.Files(DataFolder & "Data_Sales"),
    #"Filtered Hidden Files1" = Table.SelectRows(Source, each [Attributes]?[Hidden]? <> true),
    #"Invoke Custom Function1" = Table.AddColumn(#"Filtered Hidden Files1", "Transform File", each #"Transform File"([Content])),
    #"Renamed Columns1" = Table.RenameColumns(#"Invoke Custom Function1", {"Name", "Source.Name"}),
    #"Removed Other Columns1" = Table.SelectColumns(#"Renamed Columns1", {"Source.Name", "Transform File"}),
    #"Expanded Table Column1" = Table.ExpandTableColumn(#"Removed Other Columns1", "Transform File", Table.ColumnNames(#"Transform File"(#"Sample File"))),
    #"Removed Columns" = Table.RemoveColumns(#"Expanded Table Column1", {"Source.Name"}),
    #"Trimmed Text" = Table.TransformColumns(#"Removed Columns", {{"Store_Code", Text.Trim, type text}}),
    #"Changed Type with Locale" = Table.TransformColumnTypes(#"Trimmed Text", {{"Order_Date", type date}}, "vi-VN"),
    #"Changed Type with Locale1" = Table.TransformColumnTypes(#"Changed Type with Locale", {{"Unit_Price", Int64.Type}}, "vi-VN"),
    #"Changed Type" = Table.TransformColumnTypes(#"Changed Type with Locale1", {{"Order_ID", type text}, {"Line_No", Int64.Type}, {"Customer_ID", type text}, {"Product_Code", type text}, {"Quantity", Int64.Type}}),
    #"Removed Duplicates" = Table.Distinct(#"Changed Type"),
    #"Added Custom" = Table.AddColumn(#"Removed Duplicates", "Revenue", each [Quantity] * [Unit_Price], Int64.Type),
    #"Inserted Year" = Table.AddColumn(#"Added Custom", "Year", each Date.Year([Order_Date]), Int64.Type),
    #"Renamed Columns" = Table.RenameColumns(#"Inserted Year", {{"Year", "Nam"}}),
    #"Inserted Month" = Table.AddColumn(#"Renamed Columns", "Month", each Date.Month([Order_Date]), Int64.Type),
    #"Renamed Columns2" = Table.RenameColumns(#"Inserted Month", {{"Month", "Thang"}})
in
    #"Renamed Columns2\""""

M_PRODUCT = """let
    Source = Excel.Workbook(File.Contents(DataFolder & "B4_Lookups.xlsx"), null, true),
    Product_Table = Source{[Item = "Product", Kind = "Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(Product_Table, {{"Product_Code", type text}, {"Product_Name", type text}, {"Category", type text}, {"Unit_Price", Int64.Type}, {"Unit_Cost", Int64.Type}}),
    #"Removed Duplicates" = Table.Distinct(#"Changed Type", {"Product_Code"})
in
    #"Removed Duplicates\""""

M_STORE = """let
    Source = Excel.Workbook(File.Contents(DataFolder & "B4_Lookups.xlsx"), null, true),
    Store_Table = Source{[Item = "Store", Kind = "Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(Store_Table, {{"Store_Code", type text}, {"Store_Name", type text}, {"City", type text}, {"Region", type text}, {"Store_Type", type text}, {"Open_Date", type date}})
in
    #"Changed Type\""""

M_CUSTOMER = """let
    Source = Excel.Workbook(File.Contents(DataFolder & "B4_Lookups.xlsx"), null, true),
    Customer_Table = Source{[Item = "Customer", Kind = "Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(Customer_Table, {{"Customer_ID", type text}, {"Full_Name", type text}, {"Gender", type text}, {"Birth_Year", Int64.Type}, {"Segment", type text}, {"Join_Date", type date}, {"Home_Store_Code", type text}})
in
    #"Changed Type\""""

M_TARGET = """let
    Source = Excel.Workbook(File.Contents(DataFolder & "B4_Target_2025.xlsx"), null, true),
    Target_2025_Sheet = Source{[Item = "Target_2025", Kind = "Sheet"]}[Data],
    #"Removed Top Rows" = Table.Skip(Target_2025_Sheet, 2),
    #"Promoted Headers" = Table.PromoteHeaders(#"Removed Top Rows", [PromoteAllScalars = true]),
    #"Unpivoted Other Columns" = Table.UnpivotOtherColumns(#"Promoted Headers", {"Store_Code", "Store_Name"}, "Attribute", "Value"),
    #"Renamed Columns" = Table.RenameColumns(#"Unpivoted Other Columns", {{"Attribute", "Thang"}, {"Value", "Target"}}),
    #"Extracted Text After Delimiter" = Table.TransformColumns(#"Renamed Columns", {{"Thang", each Text.AfterDelimiter(_, " "), type text}}),
    #"Changed Type" = Table.TransformColumnTypes(#"Extracted Text After Delimiter", {{"Store_Code", type text}, {"Thang", Int64.Type}, {"Target", Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type", {"Store_Name"}),
    #"Added Custom" = Table.AddColumn(#"Removed Columns", "Ngay_Dau_Thang", each #date(2025, [Thang], 1), type date)
in
    #"Added Custom\""""

M_CALENDAR = """let
    Source = List.Dates(#date(2024, 1, 1), 731, #duration(1, 0, 0, 0)),
    #"Converted to Table" = Table.FromList(Source, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    #"Renamed Columns" = Table.RenameColumns(#"Converted to Table", {{"Column1", "Date"}}),
    #"Changed Type" = Table.TransformColumnTypes(#"Renamed Columns", {{"Date", type date}}),
    #"Inserted Year" = Table.AddColumn(#"Changed Type", "Year", each Date.Year([Date]), Int64.Type),
    #"Inserted Month" = Table.AddColumn(#"Inserted Year", "Month", each Date.Month([Date]), Int64.Type),
    #"Inserted Quarter" = Table.AddColumn(#"Inserted Month", "Quarter", each Date.QuarterOfYear([Date]), Int64.Type),
    #"Inserted Day Name" = Table.AddColumn(#"Inserted Quarter", "Day Name", each Date.DayOfWeekName([Date]), type text),
    #"Renamed Columns1" = Table.RenameColumns(#"Inserted Day Name", {{"Year", "Nam"}, {"Month", "Thang"}, {"Quarter", "Quy"}, {"Day Name", "Thu"}})
in
    #"Renamed Columns1\""""

queries = [
    Query("DataFolder", M_DATAFOLDER, "Tham so: thu muc Practice (co dau \\ o cuoi). Mo file tren may khac: Manage Parameters -> sua gia tri nay"),
    Query("Sample File", M_SAMPLE_FILE),                      # 4 query trợ giúp của Combine (không mô tả,
    Query("Parameter1", M_PARAMETER1),                        #  giống hệt khi tạo bằng giao diện)
    Query("Transform Sample File", M_TRANSFORM_SAMPLE_FILE),
    Query("Transform File", M_TRANSFORM_FILE, is_function=True),
    Query("Sales", M_SALES, "From Folder Data_Sales (Combine): Trim, Using Locale vi-VN, Remove Duplicates, Revenue/Nam/Thang"),
    Query("Product", M_PRODUCT, "Danh muc san pham, da Remove Duplicates theo Product_Code (P015)"),
    Query("Store", M_STORE, "Danh muc cua hang"),
    Query("Customer", M_CUSTOMER, "Danh muc khach hang"),
    Query("Target", M_TARGET, "Ke hoach 2025 da Unpivot: Store_Code, Thang, Target, Ngay_Dau_Thang"),
    Query("Calendar", M_CALENDAR, "Bang ngay 01/01/2024-31/12/2025 (731 ngay)"),
]

# Nhóm query như khung Queries & Connections sau Combine (ảnh S27): "Transform File from Data_Sales" chứa
# "Helper Queries" (Sample File, Parameter1, Transform File) và Transform Sample File; các query còn lại ở "Other Queries".
G_COMBINE, G_HELPER = "f9b2bfa8-08a5-4938-9d5b-8ab1db2ea31f", "0c050e38-bc81-4f3b-aadd-702f34194e2b"
query_groups = [{"id": G_COMBINE, "name": "Transform File from Data_Sales"},
                {"id": G_HELPER, "name": "Helper Queries", "parent": G_COMBINE}]
group_members = {"Sample File": G_HELPER, "Parameter1": G_HELPER, "Transform File": G_HELPER, "Transform Sample File": G_COMBINE}
function_bindings = {"Transform File": "Transform Sample File"}

table_loads = [TableLoad(query="Sales", sheet="Sales_PQ", table_name="Sales", anchor="A1", also_model=True, background=False)]
model_loads = [ModelLoad("Product"), ModelLoad("Store"), ModelLoad("Customer"), ModelLoad("Target"), ModelLoad("Calendar")]

pivots = [Pivot(sheet="PV_Refresh", name="PV_Refresh", source_table="Sales", anchor="A3",
                rows=["Nam", "Thang"], values=[("Order_ID", "So_dong", "count"), ("Revenue", "Doanh_thu", "sum")],
                filters=["Store_Code"])]

kiemtra = KiemTra(
    sheet="Kiem_tra_B4",
    title="KIỂM TRA / ĐỐI SOÁT — Brick output B4 (Power Query)",
    note=("Cột C: công thức đọc bảng Sales / Pivot PV_Refresh, hoặc số dòng của bảng trong Data Model (đọc lúc dựng file). Cột D: đáp án tính từ dataset "
          "(build_b4_meta.json). File ở trạng thái SAU khi thả tháng 4. Mở trên máy khác: Data → Get Data → Launch Power Query Editor → Home → "
          "Manage Parameters → sửa DataFolder thành thư mục Practice của máy bạn (có dấu \\ ở cuối) → Close & Load → Refresh All."),
    rows=[
        ("Sales: số dòng (01/2024–04/2025, đã làm sạch)", "=ROWS(Sales[Order_ID])", None, 30247),
        ("Sales: tổng doanh thu (VNĐ)", "=SUM(Sales[Revenue])", None, 4436849500),
        ("PV_Refresh: tổng doanh thu − tổng Sales (phải = 0)", '=GETPIVOTDATA("Doanh_thu",PV_Refresh!$A$3)-SUM(Sales[Revenue])', None, 0),
        ("PV_Refresh: số dòng tháng 4/2025 (file vừa thả vào)", '=GETPIVOTDATA("So_dong",PV_Refresh!$A$3,"Nam",2025,"Thang",4)', None, 1919),
        ("Sales: ST12 không có dòng nào trước 06/2024 (câu chuyện #3)", '=COUNTIFS(Sales[Store_Code],"ST12",Sales[Nam],2024,Sales[Thang],"<6")', None, 0),
        ("Product: số dòng trong Data Model (đã bỏ P015 trùng)", None, "model:Product", 40),
        ("Store: số dòng trong Data Model", None, "model:Store", 12),
        ("Customer: số dòng trong Data Model", None, "model:Customer", 1500),
        ("Target: số dòng trong Data Model (12 cửa hàng × 12 tháng)", None, "model:Target", 144),
        ("Calendar: số dòng trong Data Model (01/01/2024–31/12/2025)", None, "model:Calendar", 731),
    ],
)

checks = ["Sales_PQ!A1"] + [f"Kiem_tra_B4!E{r}" for r in range(5, 15)]

# File mốc: trạng thái cuối slide 27 (trước Thực hành 1). Số mong đợi lấy từ build_b4_meta.json (build_b4_homework.py).
meta = json.load(open(META, encoding="utf-8"))
vn = lambda x: f"{x:,.0f}".replace(",", ".")
kiemtra_moc = KiemTra(
    sheet="Huong_dan_Moc",
    title="FILE MỐC BUỔI 4 — trạng thái cuối slide 27, trước Thực hành 1",
    note=("Dùng file này khi file B4_<HoTen>.xlsx của bạn bị lỗi hoặc chưa làm kịp. File đã có đủ các query làm trên lớp tới slide 27: "
          "tham số DataFolder, Sales (From Folder, 4 file CSV), Product, Store, Customer, Target, Calendar, Pivot PV_Refresh. "
          "Làm theo các bước ở dưới bảng, rồi làm tiếp Thực hành 1 cùng cả lớp."),
    expected_label="Mong đợi",
    rows=[
        ("Sales: số dòng (01/2024–03/2025, chưa có tháng 4)", "=ROWS(Sales[Order_ID])", None, meta["q_rows"]),
        ("Sales: tổng doanh thu (VNĐ)", "=SUM(Sales[Revenue])", None, meta["q_rev"]),
        ("PV_Refresh: tổng doanh thu − tổng Sales (phải = 0)", '=GETPIVOTDATA("Doanh_thu",PV_Refresh!$A$3)-SUM(Sales[Revenue])', None, 0),
        ("Sales: số dòng tháng 4/2025 (chưa thả file tháng 4)", '=COUNTIFS(Sales[Nam],2025,Sales[Thang],4)', None, 0),
        ("Sales: ST12 không có dòng nào trước 06/2024", '=COUNTIFS(Sales[Store_Code],"ST12",Sales[Nam],2024,Sales[Thang],"<6")', None, meta["st12_before_jun"]),
        ("Product: số dòng trong Data Model (đã bỏ P015 trùng)", None, "model:Product", meta["product"]),
        ("Store: số dòng trong Data Model", None, "model:Store", meta["store"]),
        ("Customer: số dòng trong Data Model", None, "model:Customer", meta["customer"]),
        ("Target: số dòng trong Data Model (12 cửa hàng × 12 tháng)", None, "model:Target", meta["target_rows"]),
        ("Calendar: số dòng trong Data Model (01/01/2024–31/12/2025)", None, "model:Calendar", meta["calendar_rows"]),
    ],
    steps=[
        "1. File → Save As, lưu thành B4_<HoTen>.xlsx để dùng thay file cũ.",
        "2. Sửa đường dẫn: Data → Get Data → Launch Power Query Editor → Home → Manage Parameters → DataFolder: dán đường dẫn "
        "thư mục Practice trên máy bạn (thư mục chứa Data_Sales, B4_Lookups.xlsx, B4_Target_2025.xlsx), có dấu \\ ở cuối → OK → Close & Load.",
        f"3. Data → Refresh All. Cột Khớp? ở bảng trên phải toàn ✔ (Sales {vn(meta['q_rows'])} dòng, doanh thu {vn(meta['q_rev'])}). "
        "Có ✘: xem lại đường dẫn ở bước 2, và thư mục Data_Sales phải có đúng 4 file CSV.",
        "4. Làm Thực hành 1 (slide 28): chép Practice/New_Month/Sales_2025-04.csv vào Practice/Data_Sales/, rồi Refresh All. "
        "Sau bước này các dòng Sales ở bảng trên chuyển sang ✘ vì đã có thêm tháng 4: đối chiếu số mới với slide 29.",
    ],
)
checks_moc = ["Sales_PQ!A1"] + [f"Huong_dan_Moc!E{r}" for r in range(5, 15)]

if "--print-m" in sys.argv:
    for q in queries:
        print(f"\n=== {q.name} ===\n{q.formula}\n")
    sys.exit(0)

if sys.platform != "win32":
    sys.exit("build_b4.py chỉ chạy trên Windows (cần Excel qua COM để tạo Power Query). Dùng --print-m để xem mã M trên máy khác.")


def build(path, kiemtra, checks, with_april, first_sheet=None):
    print(f"Copy {B3_ANSWER}\n  -> {path}")
    shutil.copyfile(B3_ANSWER, path)
    print("Dựng Power Query bằng Excel (COM) — có thể mất vài phút...")
    apr_tmp = os.path.join(DATA_SALES, "Sales_2025-04.csv")
    had_april = os.path.exists(apr_tmp)
    if with_april and not had_april:
        shutil.copyfile(APRIL, apr_tmp)          # trạng thái sau Thực hành 1
    try:
        result = run(path, queries=queries, table_loads=table_loads, model_loads=model_loads,
                     pivots=pivots, kiemtra=kiemtra, checks=checks, first_sheet=first_sheet)
        print("Kết quả kiểm tra (lần 1):", result)
        # COM không tạo được nhóm query và liên kết Transform File ↔ Transform Sample File: gắn thẳng vào gói
        # DataMashup, rồi mở lại bằng Excel (Refresh All, lưu) để chắc Excel đọc được và lưu lại gói chuẩn.
        print("Gắn nhóm query trợ giúp, mở lại bằng Excel để Refresh All và lưu...")
        set_query_groups(path, query_groups, group_members, function_bindings)
        result = run(path, queries=[], checks=checks)
    finally:
        if with_april and not had_april and os.path.exists(apr_tmp):
            os.remove(apr_tmp)                   # Data_Sales phát cho học viên chỉ có 4 file
    print("Kết quả kiểm tra (sau khi mở lại):", result)
    if any(v != "✔" for v in result[1:]):
        sys.exit(f"{kiemtra.sheet} có dòng ✘ — xem sheet {kiemtra.sheet}")
    sec, meta_xml = read_mashup(path)
    def _group_of(q):
        key = f"<ItemPath>Section1/{q.replace(' ', '%20')}</ItemPath>"
        if key not in meta_xml:
            return None
        m = re.search(r'Type="QueryGroupID" Value="s([^"]+)"', meta_xml.split(key, 1)[1].split("</Item>", 1)[0])
        return m.group(1) if m else None
    miss = [q for q, g in group_members.items() if _group_of(q) != g]
    if 'Type="QueryGroups"' not in meta_xml or miss or "FunctionQueryBinding" not in sec:
        sys.exit(f"Excel không giữ nhóm query / liên kết hàm sau khi lưu lại (thiếu: {miss or 'QueryGroups/FunctionQueryBinding'})")
    print("Nhóm query trợ giúp và liên kết Transform File ↔ Transform Sample File: OK")
    print(f"\nXong: {path} — {kiemtra.sheet} {len(result) - 1}/{len(result) - 1} ✔")


if MOC_MODE:
    build(MOC, kiemtra_moc, checks_moc, with_april=False, first_sheet=kiemtra_moc.sheet)
else:
    build(ANSWER, kiemtra, checks, with_april=True)
