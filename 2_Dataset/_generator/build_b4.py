r"""Dựng file đáp án Buổi 4 (Power Query): B4_MCI_Mart_Answer.xlsx — file khởi đầu Buổi 5.

Bắt đầu từ đáp án B3 (copy nguyên file, giữ mọi sheet tblSales/KPI_SUMIFS/PV_.../Kiem_tra), rồi THÊM vào đúng các
query dựng trên lớp (Slide_Script_B4.md, bản 1.3):
  DataFolder  tham số (Manage Parameters): thư mục Practice của máy dựng file, có dấu \ ở cuối
  Sales       From Folder DataFolder & "Data_Sales" → nối mọi CSV → Trim Store_Code, Using Locale vi-VN cho Order_Date
              và Unit_Price, Remove Duplicates, Revenue / Nam / Thang. Load Table (sheet Sales_PQ) + Data Model,
              tắt background refresh
  Product     B4_Lookups.xlsx, Remove Duplicates theo Product_Code (P015 trùng)       Connection Only + Data Model
  Store, Customer                                                                      Connection Only + Data Model
  Target      B4_Target_2025.xlsx, Unpivot → Store_Code | Thang (số) | Target | Ngay_Dau_Thang    Connection + Data Model
  Calendar    List.Dates 01/01/2024–31/12/2025 (731 ngày), Nam / Thang / Quy / Thu       Connection Only + Data Model
Pivot PV_Refresh (Rows Nam, Thang; Values So_dong = Count of Order_ID, Doanh_thu = Sum of Revenue; Filters Store_Code)
và sheet Kiem_tra_B4 (10 dòng).

Đáp án thể hiện trạng thái SAU Thực hành 1 (đã thả file tháng 4): lúc dựng, script chép tạm
Practice/New_Month/Sales_2025-04.csv vào Practice/Data_Sales/, Refresh, rồi xóa lại, để thư mục Data_Sales phát cho
học viên vẫn chỉ có 4 file. Refresh file đáp án với thư mục 4 file thì ra 28.328 dòng (trạng thái trước Thực hành 1).

openpyxl không tạo được Power Query — toàn bộ việc thêm query/load/Pivot làm bằng Excel qua COM (PowerShell),
xem `powerquery_com.py`. CHỈ CHẠY ĐƯỢC TRÊN WINDOWS, cần Excel đã cài Power Query/Power Pivot.

Số liệu đối soát: build_b4_meta.json (build_b4_homework.py) — Sales 30.247 dòng, 4.436.849.500; tháng 4/2025
1.919 dòng; Product 40, Store 12, Customer 1.500, Target 144, Calendar 731.

Chạy từ thư mục MCI Excel:
    Windows:  python 2_Dataset\_generator\build_b4.py
    Thêm --print-m để chỉ in mã M của từng query ra màn hình (kiểm tra cú pháp), không mở Excel.
Trước khi chạy: đóng B4_MCI_Mart_Answer.xlsx nếu đang mở trong Excel.
"""
import os, shutil, sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from powerquery_com import Query, TableLoad, ModelLoad, Pivot, KiemTra, run

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

M_DATAFOLDER = '"' + PRACTICE.rstrip("\\/") + os.sep + '" meta [IsParameterQuery = true, Type = "Text", IsParameterQueryRequired = true]'

M_SALES = """let
    Source = Folder.Files(DataFolder & "Data_Sales"),
    #"Filtered CSV" = Table.SelectRows(Source, each Text.Lower([Extension]) = ".csv" and not Text.StartsWith([Name], "~")),
    #"Parsed CSV" = Table.AddColumn(#"Filtered CSV", "Data", each Table.PromoteHeaders(Csv.Document([Content], [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.None]), [PromoteAllScalars = true])),
    Combined = Table.Combine(#"Parsed CSV"[Data]),
    #"Trimmed Text" = Table.TransformColumns(Combined, {{"Store_Code", Text.Trim, type text}}),
    #"Changed Type with Locale" = Table.TransformColumnTypes(#"Trimmed Text", {{"Order_Date", type date}, {"Unit_Price", Int64.Type}}, "vi-VN"),
    #"Changed Type" = Table.TransformColumnTypes(#"Changed Type with Locale", {{"Order_ID", type text}, {"Line_No", Int64.Type}, {"Customer_ID", type text}, {"Product_Code", type text}, {"Quantity", Int64.Type}}),
    #"Removed Duplicates" = Table.Distinct(#"Changed Type"),
    #"Added Revenue" = Table.AddColumn(#"Removed Duplicates", "Revenue", each [Quantity] * [Unit_Price], Int64.Type),
    #"Inserted Year" = Table.AddColumn(#"Added Revenue", "Nam", each Date.Year([Order_Date]), Int64.Type),
    #"Inserted Month" = Table.AddColumn(#"Inserted Year", "Thang", each Date.Month([Order_Date]), Int64.Type)
in
    #"Inserted Month\""""

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
    #"Added Ngay_Dau_Thang" = Table.AddColumn(#"Removed Columns", "Ngay_Dau_Thang", each #date(2025, [Thang], 1), type date)
in
    #"Added Ngay_Dau_Thang\""""

M_CALENDAR = """let
    Source = List.Dates(#date(2024, 1, 1), 731, #duration(1, 0, 0, 0)),
    #"Converted to Table" = Table.FromList(Source, Splitter.SplitByNothing(), {"Date"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Converted to Table", {{"Date", type date}}),
    #"Inserted Year" = Table.AddColumn(#"Changed Type", "Nam", each Date.Year([Date]), Int64.Type),
    #"Inserted Month" = Table.AddColumn(#"Inserted Year", "Thang", each Date.Month([Date]), Int64.Type),
    #"Inserted Quarter" = Table.AddColumn(#"Inserted Month", "Quy", each Date.QuarterOfYear([Date]), Int64.Type),
    #"Inserted Day Name" = Table.AddColumn(#"Inserted Quarter", "Thu", each Date.ToText([Date], "ddd", "vi-VN"), type text)
in
    #"Inserted Day Name\""""

queries = [
    Query("DataFolder", M_DATAFOLDER, "Tham so: thu muc Practice (co dau \\ o cuoi). Mo file tren may khac: Manage Parameters -> sua gia tri nay"),
    Query("Sales", M_SALES, "From Folder Data_Sales: Trim, Using Locale vi-VN, Remove Duplicates, Revenue/Nam/Thang"),
    Query("Product", M_PRODUCT, "Danh muc san pham, da Remove Duplicates theo Product_Code (P015)"),
    Query("Store", M_STORE, "Danh muc cua hang"),
    Query("Customer", M_CUSTOMER, "Danh muc khach hang"),
    Query("Target", M_TARGET, "Ke hoach 2025 da Unpivot: Store_Code, Thang, Target, Ngay_Dau_Thang"),
    Query("Calendar", M_CALENDAR, "Bang ngay 01/01/2024-31/12/2025 (731 ngay)"),
]

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

if "--print-m" in sys.argv:
    for q in queries:
        print(f"\n=== {q.name} ===\n{q.formula}\n")
    sys.exit(0)

if sys.platform != "win32":
    sys.exit("build_b4.py chỉ chạy trên Windows (cần Excel qua COM để tạo Power Query). Dùng --print-m để xem mã M trên máy khác.")

print(f"Copy {B3_ANSWER}\n  -> {ANSWER}")
shutil.copyfile(B3_ANSWER, ANSWER)
print("Dựng Power Query bằng Excel (COM) — có thể mất vài phút...")
apr_tmp = os.path.join(DATA_SALES, "Sales_2025-04.csv")
had_april = os.path.exists(apr_tmp)
if not had_april:
    shutil.copyfile(APRIL, apr_tmp)          # trạng thái sau Thực hành 1
try:
    result = run(ANSWER, queries=queries, table_loads=table_loads, model_loads=model_loads,
                 pivots=pivots, kiemtra=kiemtra, checks=checks)
finally:
    if not had_april and os.path.exists(apr_tmp):
        os.remove(apr_tmp)                   # Data_Sales phát cho học viên chỉ có 4 file
print("Kết quả kiểm tra:", result)
if any(v != "✔" for v in result[1:]):
    sys.exit("Kiem_tra_B4 có dòng ✘ — xem sheet Kiem_tra_B4")
print(f"\nXong: {ANSWER} — Kiem_tra_B4 {len(result) - 1}/{len(result) - 1} ✔")
