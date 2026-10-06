r"""Dựng 2 file demo Buổi 5 (theo mẫu B3: mỗi file có bản giảng viên + bản _HocVien, sheet Huong_dan):

  Demo/B5_Demo_Model.xlsx   Data Model 6 bảng, 0 relationship, 0 measure        — Slide 5, 8, 10, 11, 12
  Demo/B5_Demo_DAX.xlsx     Data Model 6 bảng, 6 relationship, 0 measure        — Slide 15–17, 19–20, 22–25, 27–28, 30

Bắt đầu từ đáp án B4 (Practice/B4_MCI_Mart_Answer.xlsx), bỏ các sheet Pivot / kiểm tra của B2–B4, giữ Data Model nguyên
(không Refresh), giữ sheet Sales (tblSales), Product, Store, KPI_SUMIFS và Sales_PQ. DAX: thêm 6 relationship như
build_b5.py. Mọi con số trong Huong_dan đọc từ build_b5_meta.json.

Bản _HocVien: Huong_dan đứng đầu, tiêu đề "VÍ DỤ". Bản giảng viên mở ở sheet Sales, Huong_dan để cuối.
Excel không có API cho calculated column, Mark as Date Table, Hide from Client Tools: các bước này nằm trong Huong_dan
để làm trên lớp (Demo_DAX chưa có cột COGS, chưa Mark as Date Table, chưa ẩn cột).

CHỈ CHẠY TRÊN WINDOWS (Excel qua COM, datamodel_com.run). Từ thư mục MCI Excel:
    python 2_Dataset\_generator\build_b5_demos.py           # dựng cả 4 file
    python 2_Dataset\_generator\build_b5_demos.py --check   # chỉ mở 4 file, in bảng/relationship/measure/sheet
Trước khi chạy: đóng B4_MCI_Mart_Answer.xlsx và các file Demo B5 nếu đang mở.
"""
import json, os, subprocess, sys, tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from datamodel_com import Rel, Cells, run

for s_ in (sys.stdout, sys.stderr):
    if hasattr(s_, "reconfigure"):
        s_.reconfigure(encoding="utf-8", errors="replace")

GEN = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.normpath(os.environ.get("MCI_ROOT") or os.path.dirname(os.path.dirname(GEN)))
B4_ANSWER = os.path.join(ROOT, "1_Sessions", "B4_Power_Query", "Practice", "B4_MCI_Mart_Answer.xlsx")
DEMO = os.path.join(ROOT, "1_Sessions", "B5_Data_Model_DAX", "Demo")
M = json.load(open(os.path.join(GEN, "build_b5_meta.json"), encoding="utf-8"))

DELETE = ["KPI_Card", "Kiem_tra_B4", "PV_Refresh", "PV_MienThang", "PV_NganhHang", "PV_Thang", "PV_Ngay", "PV_Thu",
          "PV_CuaHang", "ST07_T3_ChiTiet", "Kiem_tra"]
RELS = [
    Rel("Sales", "Product_Code", "Product", "Product_Code"),
    Rel("Sales", "Store_Code", "Store", "Store_Code"),
    Rel("Sales", "Customer_ID", "Customer", "Customer_ID"),
    Rel("Sales", "Order_Date", "Calendar", "Date"),
    Rel("Target", "Store_Code", "Store", "Store_Code"),
    Rel("Target", "Ngay_Dau_Thang", "Calendar", "Date"),
]


def n(x):
    return f"{x:,.0f}".replace(",", ".")


def p(x, d=1):
    return f"{x * 100:.{d}f}".replace(".", ",") + "%"


def sg(x):
    return ("+" if x >= 0 else "−") + n(abs(x))


q1, tm, reg = M["q1"], M["target_by_month"], M["target_by_region"]
vt, vm, yy = M["vs_target_month"], M["variance_month"], M["yoy_month"]
cat = M["cat_vs_target_q1"]
sm = M["store_month"]
act = M["actual_2025"]
gmy = M["gm_year"]
share = M["share_q1"]
nam_q1 = q1["rev"] - M["reg_q1"]["Bắc"]["rev"] - M["reg_q1"]["Trung"]["rev"]

# ---------------------------------------------------------------- Huong_dan (slide, thao tác, kết quả mong đợi)
MODEL_STEPS = [
    ("5", "Data → Manage Data Model (cửa sổ Power Pivot). Nhìn các tab ở dưới cùng, chọn tab Sales.",
     f"6 tab: Sales, Product, Store, Customer, Target, Calendar. Bảng Sales có {n(M['tables']['Sales'])} dòng."),
    ("5", "Home → Diagram View.", "6 hộp bảng rời nhau (Sales, Product, Store, Customer, Target, Calendar); chưa có đường nối nào."),
    ("8", "Chỉ đọc số liệu, không làm trên Excel: so sánh Merge với Relationship.",
     f"Merge cả Product, Store, Customer vào Sales: {n(M['merged_cells'])} ô lặp. Giữ lookup riêng: {n(M['lookup_cells'])} ô."),
    ("10", "Diagram View: kéo Sales[Product_Code] thả lên Product[Product_Code]; Sales[Store_Code] lên Store[Store_Code]; Sales[Customer_ID] lên Customer[Customer_ID].",
     "3 đường nối. Ký hiệu 1 ở đầu lookup, dấu * ở đầu Sales."),
    ("10", "Design → Create Relationship: bảng Sales, cột Order_Date; bảng Calendar, cột Date → OK.", "Đường nối thứ 4: Sales → Calendar."),
    ("10", "Design → Manage Relationships.", "Danh sách 4 quan hệ, đều đang Active."),
    ("11", "Kéo hộp Sales vào giữa, bốn lookup xung quanh.", "Hình ngôi sao: Sales ở giữa, 4 lookup xung quanh. Bộ lọc chảy từ lookup xuống Sales."),
    ("12", "Insert → PivotTable ▾ → From Data Model → New Worksheet; đặt tên sheet PV_Grain. Rows: Product → Category; Values: Target → Target.",
     f"6 dòng (6 ngành), dòng nào cũng {n(M['target_total'])}. PivotTable Fields báo thanh vàng thiếu relationship. Target chưa nối với bảng nào, và cũng không có cột sản phẩm."),
    ("12", "Diagram View: kéo Target[Store_Code] thả lên Store[Store_Code]. Design → Create Relationship: bảng Target, cột Ngay_Dau_Thang; bảng Calendar, cột Date → OK.",
     "Tổng 6 relationship. Sales và Target cùng nối vào Store và Calendar (hai fact, hai lookup chung)."),
    ("12", "Quay lại PV_Grain: bỏ Category khỏi Rows, kéo Store → Region vào Rows.",
     f"Bắc {n(reg['Bắc'])} · Trung {n(reg['Trung'])} · Nam {n(reg['Nam'])} (tổng {n(M['target_total'])})."),
    ("12", "Kéo Category trở lại Rows, dưới Region.", "Ngành hàng vẫn lặp cùng một số trong mỗi miền: đúng, vì Target chỉ lọc được theo cửa hàng, miền, tháng."),
]

DAX_STEPS = [
    ("15", "Power Pivot → Data View, tab Sales: bấm ô đầu tiên của cột Add Column, gõ trên thanh công thức =Sales[Quantity]*RELATED(Product[Unit_Cost]) → Enter; đổi tên cột thành COGS.",
     f"Cột COGS có giá trị cho cả {n(M['tables']['Sales'])} dòng. Tổng COGS Quý 1/2025 = {n(q1['cogs'])}."),
    ("16", "Ví dụ cho kết quả sai. Data View, tab Sales: Add Column =DIVIDE(Sales[Revenue]-Sales[COGS],Sales[Revenue]), đổi tên GM_Dong. Insert → PivotTable ▾ → From Data Model → New Worksheet, đặt tên PV_Loc; Filters: Calendar → Nam = 2025, Calendar → Quy = 1; Values: GM_Dong, đổi sang Average, định dạng %.",
     f"Average of GM_Dong = {p(M['gm_wrong_avg_q1'], 2)}. Đổi sang Sum ra {n(M['gm_wrong_sum_q1'] * 100)}%. Số đúng: sheet KPI_SUMIFS, ô E30 = {p(q1['gm'], 2)}."),
    ("16", "Data View: chuột phải tiêu đề cột GM_Dong → Delete Columns. Giữ lại sheet PV_Loc.", "Cột GM_Dong biến mất, trường này cũng biến khỏi Pivot."),
    ("17", "Trên PV_Loc: tick Sales → Revenue.", "Pivot hiện Sum of Revenue: implicit measure, không tên, không định dạng."),
    ("17", "Power Pivot → Measures → New Measure: Table name Sales; Measure name Total Revenue; Formula =SUM(Sales[Revenue]); Check formula; Category Number, tick Use 1000 separator, Decimal places 0 → OK. Tick vào Values.",
     f"Measure có biểu tượng fx trong danh sách trường. Quý 1/2025 = {n(q1['rev'])}."),
    ("17", "New Measure: Total COGS =SUM(Sales[COGS]). New Measure: Gross Margin % =DIVIDE([Total Revenue]-[Total COGS],[Total Revenue]), Category Percentage, 1 decimal. Tick vào Values.",
     f"Total COGS = {n(q1['cogs'])}; Gross Margin % = {p(q1['gm'])}, khớp KPI_SUMIFS!E30."),
    ("19", "Trên PV_Loc: Values chỉ giữ Total Revenue; Rows: Store → Region; Columns: Calendar → Thang; Filters: Nam = 2025, Quy = 1. Bấm ô Bắc × 3.",
     f"Bắc × 3 = {n(M['bac_t3_2025'])} (sheet KPI_SUMIFS, ô D6). Grand Total = {n(q1['rev'])} (KPI_SUMIFS!E9)."),
    ("20", "Insert → PivotTable ▾ → From Data Model → New Worksheet, đặt tên PV_Lookup_vs_Fact. Filters: Calendar → Nam = 2025; Rows: Calendar → Thang; Values: Total Revenue, Target → Target.",
     f"T1: doanh thu {n(act['1'])}, target {n(tm['1'])}. T2: {n(act['2'])} và {n(tm['2'])}. T5–T12 chỉ có target."),
    ("20", "Chép Pivot sang cột F (Ctrl+C, dán ở F). Trong Pivot mới: bỏ Calendar → Thang, kéo Sales → Thang vào Rows. Xem xong thì xóa Pivot bên phải.",
     f"Doanh thu vẫn đúng nhưng Sum of Target = {n(M['target_total'])} ở mọi dòng: Thang của Sales không lọc được Target."),
    ("20", "Data View, bảng Sales: giữ Ctrl chọn tiêu đề Nam, Thang, Order_Date, Store_Code, Product_Code, Customer_ID → chuột phải → Hide from Client Tools. Làm tương tự cho Target: Store_Code, Thang, Ngay_Dau_Thang.",
     "Cột chuyển màu xám trong Data View và biến mất khỏi danh sách trường của Pivot."),
    ("22", "Trên PV_Loc: bỏ Calendar → Thang khỏi Columns. New Measure: Revenue Mien Nam =CALCULATE([Total Revenue], Store[Region]=\"Nam\"). Tick vào Values.",
     f"Cột Revenue Mien Nam = {n(nam_q1)} ở cả ba dòng Bắc, Trung, Nam. Bộ lọc Region của CALCULATE thay bộ lọc của hàng; Quý 1 vẫn còn."),
    ("22", "Power Pivot → Measures → Manage Measures → chọn Revenue Mien Nam → Delete.", "Measure minh họa bị xóa."),
    ("23", "New Measure: Revenue All Stores =CALCULATE([Total Revenue], ALL(Store)). New Measure: % Revenue Share =DIVIDE([Total Revenue], [Revenue All Stores]), Category Percentage. Tick vào Values.",
     f"Revenue All Stores = {n(q1['rev'])} ở mọi dòng. % Revenue Share: Bắc {p(share['Bắc'])} · Trung {p(share['Trung'])} · Nam {p(share['Nam'])} (cùng Pct_Share Buổi 1)."),
    ("24", "New Measure (Table name Target): Total Target =SUM(Target[Target]). New Measure (Table name Sales): Revenue vs Target % =DIVIDE([Total Revenue],[Total Target]), Percentage. Insert → PivotTable ▾ → From Data Model, đặt tên PV_Target; Filters: Nam = 2025; Rows: Calendar → Thang; Values: Total Revenue, Total Target, Revenue vs Target %.",
     f"T1 {p(vt['1'])} · T2 {p(vt['2'])} · T3 {p(vt['3'])} · T4 {p(vt['4'])}. T1 + T2 gộp {p(M['t1_t2_pct'])}."),
    ("24", "New Measure: Variance =[Total Revenue]-[Total Target], quan sát tháng 5. Rồi Manage Measures → Edit Variance: =IF(ISBLANK([Total Revenue]), BLANK(), [Total Revenue]-[Total Target]).",
     f"Trước khi sửa: T1 {sg(vm['1'])}, T5 {sg(M['variance_naive']['5'])} (chưa có doanh thu). Sau khi sửa: T2 {sg(vm['2'])}, T3 {sg(vm['3'])}, T4 {sg(vm['4'])}; T5–T12 để trống."),
    ("25", "Trên PV_Target: Rows → Product → Category; Filters thêm Calendar → Quy = 1.",
     f"Bánh kẹo {p(cat['Bánh kẹo'])}, Sữa {p(cat['Sữa'])}: vô nghĩa, vì Target không có chiều ngành hàng, mẫu số ở mọi dòng là target của cả chuỗi."),
    ("25", "Bỏ Quy khỏi Filters; Rows: Store → Store_Code; Columns: Calendar → Thang; Values chỉ giữ Revenue vs Target %; lọc tháng 1–4. Bấm ô ST07 × 3.",
     f"ST07 tháng 3 = {p(sm['ST07']['3']['pct'])} (doanh thu {n(sm['ST07']['3']['rev'])}, target {n(sm['ST07']['3']['target'])}); tháng 4 = {p(sm['ST07']['4']['pct'])}."),
    ("27", "Cửa sổ Power Pivot, tab Calendar → Design → Mark as Date Table → Date column: Date → OK.",
     f"Calendar thành bảng ngày ({n(M['tables']['Calendar'])} ngày liên tục, 01/01/2024–31/12/2025)."),
    ("27", "New Measure: Revenue YTD =CALCULATE([Total Revenue], DATESYTD(Calendar[Date])). New Measure: Revenue LY =CALCULATE([Total Revenue], DATEADD(Calendar[Date], -1, YEAR)). Insert → PivotTable ▾ → From Data Model, đặt tên PV_YoY; Filters: Nam = 2025; Rows: Calendar → Thang; Values: Total Revenue, Revenue LY, Revenue YTD.",
     f"T1: {n(act['1'])}; cùng kỳ {n(M['rev_2024']['1'])}. YTD tới T4 = {n(M['ytd_2025']['4'])}."),
    ("27", "New Measure: Revenue YoY % =DIVIDE([Total Revenue]-[Revenue LY],[Revenue LY]), Percentage. Quan sát T1, T2 rồi T5. Sau đó Edit thành =IF(ISBLANK([Total Revenue]), BLANK(), DIVIDE([Total Revenue]-[Revenue LY],[Revenue LY])) và lọc Thang 1–4.",
     f"T1 {'+' if yy['1'] >= 0 else '−'}{p(abs(yy['1']))} · T2 {'+' if yy['2'] >= 0 else '−'}{p(abs(yy['2']))} · T3 +{p(yy['3'])} · T4 +{p(yy['4'])}. Trước khi sửa, T5 = −100%. YTD tới T4 so cùng kỳ: +{p(M['ytd_yoy_t4'])}."),
    ("28", "Trên PV_YoY: Rows → Store → Store_Code; Filters: Calendar → Thang chọn 1, 2, 3, 4; Values: Total Revenue, Revenue LY, Revenue YoY %.",
     f"ST12 có hai ô trống (Revenue LY và Revenue YoY %): khai trương 01/06/2024, không có kỳ gốc. Dòng tổng +{p(M['ytd_yoy_t4'])}."),
    ("28", "Filters: thêm Store → Store_Code, bỏ chọn ST12. Xem xong thì bỏ lọc.",
     f"Like-for-like ({M['hw']['lfl_stores']} cửa hàng): +{p(M['hw']['lfl_ytd_yoy'])}."),
    ("30", "Manage Measures → Edit Gross Margin % → ô Description: (Doanh thu − giá vốn) ÷ doanh thu. Giá vốn: một bảng Product[Unit_Cost] cho mọi năm.",
     "Định nghĩa và giới hạn nằm trong mô hình."),
    ("30", "Trên PV_YoY: xóa mọi bộ lọc ở vùng Filters (Nam, Thang, Store_Code); Rows: Calendar → Nam; Values chỉ giữ Gross Margin %.",
     f"2024: {p(gmy['2024'])} · 2025: {p(gmy['2025'])}. Một bảng giá vốn cho cả hai năm nên biên 2024 thấp giả."),
]

SPECS = {
    "Model": dict(rels=[], steps=MODEL_STEPS, slides="5–12",
                  head="DỮ LIỆU: Data Model 6 bảng, CHƯA có relationship, chưa có measure (như file khởi đầu Buổi 5). Làm theo thứ tự các dòng bên dưới.",
                  what="DỰNG DATA MODEL: BẢNG, KHÓA, RELATIONSHIP, FACT THỨ HAI"),
    "DAX": dict(rels=RELS, steps=DAX_STEPS, slides="15–30",
                head="DỮ LIỆU: Data Model đã nối đủ 6 relationship (Sales và Target nối Store, Calendar). CHƯA có measure, chưa có cột COGS, chưa Mark as Date Table, chưa ẩn cột. Làm theo thứ tự các dòng bên dưới.",
                what="CALCULATED COLUMN, MEASURE, FILTER CONTEXT, CALCULATE, TARGET, TIME INTELLIGENCE"),
}


def huong_dan(kind, student):
    sp = SPECS[kind]
    word = "VÍ DỤ" if student else "DEMO"
    items = [("A1", f"{word} B5 — {sp['what']} (slide {sp['slides']})"),
             ("A2", sp["head"]),
             ("A3", "Chỉ dùng 6 bảng của Data Model (tab All của PivotTable Fields còn hiện các bảng tblSales, tblProduct, tblStore của B1–B3: không kéo trường của chúng). "
                    "Không bấm Refresh All: tham số DataFolder trỏ tới máy dựng file, dữ liệu đã nằm sẵn trong Data Model. Sheet KPI_SUMIFS (Buổi 2) dùng để đối chiếu."),
             ("A5", "Slide"), ("B5", "Thao tác"), ("C5", "Kết quả mong đợi")]
    for i, (sl, step, res) in enumerate(sp["steps"], start=6):
        items += [(f"A{i}", sl), (f"B{i}", step), (f"C{i}", res)]
    last = 5 + len(sp["steps"])
    return Cells("Huong_dan", items, widths={"A": 8, "B": 80, "C": 70}, bold=["A5:C5"], fill={"A5:C5": "203864"},
                 font_color={"A5:C5": "FFFFFF"}, wrap=[f"A6:C{last}", "A2:C3"], title="x")


def title_only(kind, student):
    sp = SPECS[kind]
    word = "VÍ DỤ" if student else "DEMO"
    return Cells("Huong_dan", [("A1", f"{word} B5 — {sp['what']} (slide {sp['slides']})")], title="x")


def check_ps(path):
    p_ = path.replace("'", "''")
    return r"""
$ErrorActionPreference = 'Stop'
function Com([scriptblock]$b) { for ($i = 0; ; $i++) { try { return (& $b) } catch { if ($_.Exception.ToString() -match '0x80010001|0x8001010A|RPC_E_CALL_REJECTED' -and $i -lt 300) { Start-Sleep -Milliseconds 500 } else { throw } } } }
$before = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$x = New-Object -ComObject Excel.Application
$mine = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | ForEach-Object { $_.Id })
try {
  Com { $x.Visible = $false; $x.DisplayAlerts = $false; $x.AskToUpdateLinks = $false }
  $wb = Com { $x.Workbooks.Open('""" + p_ + r"""', 0, $true) }
  $m = $null
  for ($i=0; $i -lt 40; $i++) { $c = 0; try { $m = $wb.Model; $c = $m.ModelTables.Count } catch { }; if ($c -gt 0) { break }; Start-Sleep -Seconds 3 }
  $t = @(); foreach ($mt in $m.ModelTables) { $t += ($mt.Name + '=' + $mt.RecordCount) }
  'tables: ' + ($t -join ', ')
  'relationships: ' + (Com { $m.ModelRelationships.Count }) + '  measures: ' + (Com { $m.ModelMeasures.Count })
  $s = @(); foreach ($w in $wb.Worksheets) { $s += $w.Name }
  'sheets: ' + ($s -join ' | ')
  'pivots: ' + (@($wb.Worksheets | ForEach-Object { $_.PivotTables().Count } | Measure-Object -Sum).Sum)
  Com { $wb.Close($false) }
} finally {
  if ($mine.Count -gt 0) { try { $x.Quit() } catch { } }
  Start-Sleep -Milliseconds 800
  foreach ($id in $mine) { Stop-Process -Id $id -Force -ErrorAction SilentlyContinue }
}
"""


def main():
    if sys.platform != "win32":
        sys.exit("build_b5_demos.py chỉ chạy trên Windows (Excel qua COM).")
    files = [(k, os.path.join(DEMO, f"B5_Demo_{k}.xlsx"), os.path.join(DEMO, f"B5_Demo_{k}_HocVien.xlsx")) for k in ("Model", "DAX")]
    if "--check" in sys.argv:
        for k, teacher, student in files:
            for f in (teacher, student):
                if not os.path.exists(f):
                    print(f"[thiếu] {f}")
                    continue
                ps1 = os.path.join(tempfile.gettempdir(), "mci_check_demo.ps1")
                with open(ps1, "w", encoding="utf-8-sig") as fh:
                    fh.write(check_ps(f))
                r = subprocess.run(["powershell.exe", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", ps1],
                                   capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=600)
                print(f"== {os.path.basename(f)}\n{r.stdout}{r.stderr}")
        return
    if not os.path.exists(B4_ANSWER):
        sys.exit(f"Không thấy file khởi đầu: {B4_ANSWER}")
    os.makedirs(DEMO, exist_ok=True)
    for k, teacher, student in files:
        sp = SPECS[k]
        print(f"Dựng {os.path.basename(teacher)} (từ đáp án B4, Excel qua COM, vài phút)...")
        _, log = run(B4_ANSWER, teacher, delete_sheets=DELETE, rels=sp["rels"], cells=[huong_dan(k, False)], active_sheet="Sales")
        for line in log:
            print("  ·", line)
        print(f"Dựng {os.path.basename(student)} (từ bản giảng viên: Huong_dan lên đầu, tiêu đề VÍ DỤ)...")
        _, log = run(teacher, student, cells=[title_only(k, True)], first_sheets=["Huong_dan"], active_sheet="Huong_dan")
        for line in log:
            print("  ·", line)
    print("\nXong. Chạy lại với --check để in số bảng / relationship / measure / sheet của 4 file.")


if __name__ == "__main__":
    main()
