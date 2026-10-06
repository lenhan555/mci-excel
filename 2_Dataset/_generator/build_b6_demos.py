"""Buổi 6: dựng 3 file Demo (Demo/ của B6), mỗi file có bản giảng viên + bản _HocVien và sheet Huong_dan.

  B6_Demo_DashboardXau.xlsx  Slide 2, 8, 9     dashboard "xấu" (12 chart) + Demo_TrucY (trục Y cắt khúc)
  B6_Demo_Chart.xlsx         Slide 11-14       Demo_Combo · Demo_Waterfall · Demo_Phan_phoi · Demo_Scatter · Demo_Map
  B6_Demo_DuBao.xlsx         Slide 23-24       Table tblNgay (486 ngày) · Demo_Trendline · Demo_ETS · Thong_ke

Không cần Data Model: chạy được trên Mac (openpyxl). Số liệu tính từ dataset (b6_data.py), không gõ tay.
Chart của Demo_Combo/Waterfall/Phan_phoi/Scatter/DuBao KHÔNG dựng sẵn: giảng viên tạo trước lớp (đó là nội dung slide).
Dashboard_Xau và Demo_TrucY dựng sẵn chart (đó là "vật" để chê).

    /Users/leopham/anaconda3/bin/python _generator/build_b6_demos.py
Sau khi dựng: mở từng file bằng Excel và kiểm kết quả mong đợi (xem phần kiểm tra trong Slide_Script_B6.md, Phụ lục File Demo).
"""
import os, sys
import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, PieChart, PieChart3D, LineChart, AreaChart, DoughnutChart, RadarChart, ScatterChart, Reference, Series, BarChart3D
from openpyxl.chart.series import DataPoint
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

for s_ in (sys.stdout, sys.stderr):
    if hasattr(s_, "reconfigure"): s_.reconfigure(encoding="utf-8", errors="replace")
from b6_data import ROOT, s, store, daily

DEMO = f"{ROOT}/1_Sessions/B6_Dashboard/Demo"
os.makedirs(DEMO, exist_ok=True)

F = "Arial"
HDR_FONT = Font(name=F, bold=True, color="FFFFFF")
HDR_FILL = PatternFill("solid", fgColor="203864")
TITLE = Font(name=F, size=14, bold=True, color="203864")
BODY = Font(name=F, size=10)
BOLD = Font(name=F, size=10, bold=True)
INPUT = Font(name=F, size=10, color="0000FF")
NOTE = Font(name=F, size=10, italic=True, color="585858")
NUM, PCT, DATE = "#,##0", "0.0%", "dd/mm/yyyy"
RAINBOW = ["FF0000", "FF8000", "FFFF00", "80FF00", "00FFFF", "0000FF", "8000FF", "FF00FF", "00FF80", "FF0080", "808000", "008080"]


def widths(ws, ws_):
    for i, w in enumerate(ws_, start=1): ws.column_dimensions[get_column_letter(i)].width = w


def header(ws, row, labels, col=1):
    for j, h in enumerate(labels, start=col):
        c = ws.cell(row=row, column=j, value=h); c.font = HDR_FONT; c.fill = HDR_FILL
        c.alignment = Alignment(wrap_text=True, vertical="center")


def put(ws, ref, v, font=BODY, fmt=None):
    ws[ref] = v; ws[ref].font = font
    if fmt: ws[ref].number_format = fmt


def guide(ws, title, steps):
    ws["A1"] = title; ws["A1"].font = TITLE
    header(ws, 3, ["Slide", "Thao tác", "Kết quả mong đợi"])
    for i, st in enumerate(steps, start=4):
        for j, v in enumerate(st, start=1):
            c = ws.cell(row=i, column=j, value=v); c.font = BODY
            c.alignment = Alignment(wrap_text=True, vertical="top")
    widths(ws, [8, 80, 55])


# Câu chỉ dành cho giảng viên → viết lại thành bước làm của người học (bản _HocVien)
HV_REPLACE = []


def save_both(wb, name, work_sheet):
    """Bản giảng viên: mở ra ở sheet làm việc, Huong_dan để cuối. Bản _HocVien: Huong_dan đứng đầu, tiêu đề VÍ DỤ,
    sheet Demo_* đổi thành Vi_du_*, không còn chữ dành cho giảng viên (kiểm bằng assert)."""
    g = wb["Huong_dan"]
    wb.move_sheet(g, offset=len(wb.sheetnames) - 1 - wb.sheetnames.index("Huong_dan"))
    wb.active = wb.sheetnames.index(work_sheet)
    for w in wb.worksheets: w.sheet_view.tabSelected = w.title == work_sheet
    wb.save(f"{DEMO}/{name}.xlsx")
    g["A1"] = g["A1"].value.replace("DEMO B6", "VÍ DỤ B6")
    for row in g.iter_rows(min_row=4):
        for c in row:
            if isinstance(c.value, str):
                for a, b in HV_REPLACE: c.value = c.value.replace(a, b)
                for w in wb.worksheets:
                    if w.title.startswith("Demo_"): c.value = c.value.replace(w.title, "Vi_du_" + w.title[5:])
    for w in wb.worksheets:
        if w.title.startswith("Demo_"): w.title = "Vi_du_" + w.title[5:]
    wb.move_sheet(g, offset=-wb.sheetnames.index("Huong_dan"))
    wb.active = 0
    for w in wb.worksheets: w.sheet_view.tabSelected = w.title == "Huong_dan"
    wb.save(f"{DEMO}/{name}_HocVien.xlsx")
    banned = ("giảng viên", "trước lớp", "cho lớp xem", "DEMO")
    for row in g.iter_rows():
        for c in row:
            if isinstance(c.value, str): assert not any(b in c.value for b in banned), (name, c.coordinate, c.value)


d25 = s[s.Y == 2025]
d24 = s[s.Y == 2024]
CATS = list(d25[d25.M <= 4].groupby("Category").Revenue.sum().sort_values(ascending=False).index)
cat = {c: (int(d25[(d25.M <= 4) & (d25.Category == c)].Revenue.sum()),
           float((d25[(d25.M <= 4) & (d25.Category == c)].Revenue.sum() - d25[(d25.M <= 4) & (d25.Category == c)].COGS.sum())
                 / d25[(d25.M <= 4) & (d25.Category == c)].Revenue.sum())) for c in CATS}
ytd = int(d25[d25.M <= 4].Revenue.sum())
rev24 = {m: int(d24[d24.M == m].Revenue.sum()) for m in range(1, 13)}
rev25 = {m: int(d25[d25.M == m].Revenue.sum()) for m in range(1, 5)}

# =====================================================================
# 1. B6_Demo_DashboardXau.xlsx
# =====================================================================
wb = Workbook(); ws = wb.active; ws.title = "Dashboard_Xau"
dl = wb.create_sheet("Du_lieu")
dl["A1"] = "Dữ liệu cho dashboard xấu (từ MCI Mart; số dựng bằng script, không gõ tay)"; dl["A1"].font = TITLE
# A: ngành T1–T4/2025
header(dl, 3, ["Ngành", "Doanh thu T1–T4/2025"])
for k, c in enumerate(CATS, start=4):
    put(dl, f"A{k}", c); put(dl, f"B{k}", cat[c][0], INPUT, NUM)
# D: 12 tháng 2024
header(dl, 3, ["Tháng 2024", "Doanh thu 2024"], col=4)
for m in range(1, 13):
    put(dl, f"D{m + 3}", f"T{m}"); put(dl, f"E{m + 3}", rev24[m], INPUT, NUM)
# G: miền
header(dl, 3, ["Miền", "Q1/2025"], col=7)
reg_q1 = d25[d25.Q == 1].groupby("Region").Revenue.sum()
for k, r in enumerate(["Bắc", "Trung", "Nam"], start=4):
    put(dl, f"G{k}", r); put(dl, f"H{k}", int(reg_q1[r]), INPUT, NUM)
# J: cửa hàng
header(dl, 3, ["Cửa hàng", "T1–T4/2025"], col=10)
st = d25[d25.M <= 4].groupby("Store_Code").Revenue.sum()
for k, sc in enumerate(store.Store_Code, start=4):
    put(dl, f"J{k}", sc); put(dl, f"K{k}", int(st[sc]), INPUT, NUM)
# M: tháng 2025
header(dl, 3, ["Tháng 2025", "Doanh thu 2025"], col=13)
for m in range(1, 5):
    put(dl, f"M{m + 3}", f"T{m}"); put(dl, f"N{m + 3}", rev25[m], INPUT, NUM)
widths(dl, [18, 22, 3, 14, 18, 3, 10, 16, 3, 12, 16, 3, 14, 18])


def ref(sheet, col, r1, r2): return Reference(sheet, min_col=col, min_row=r1, max_row=r2)


def rainbow(chart, n):
    s_ = chart.series[0]
    for i in range(n):
        pt = DataPoint(idx=i); pt.graphicalProperties = GraphicalProperties(solidFill=RAINBOW[i % len(RAINBOW)])
        s_.dPt.append(pt)


def place(chart, anchor, title, w=11, h=6.5):
    chart.title = title; chart.width, chart.height = w, h
    ws.add_chart(chart, anchor)


ws["A1"] = "DASHBOARD BÁN HÀNG MCI MART"; ws["A1"].font = Font(name=F, size=20, bold=True, color="FF0000")
ws["A2"] = "(ví dụ xấu — đừng làm theo)"; ws["A2"].font = NOTE
# 1 Pie 3D ngành (6 phần)
c = PieChart3D(); c.add_data(Reference(dl, min_col=2, min_row=3, max_row=9), titles_from_data=True); c.set_categories(ref(dl, 1, 4, 9))
place(c, "A4", "Ngành")
# 2 Bar cột, trục Y cắt khúc
c = BarChart(); c.type = "col"; c.add_data(Reference(dl, min_col=5, min_row=3, max_row=15), titles_from_data=True); c.set_categories(ref(dl, 4, 4, 15))
c.y_axis.scaling.min = 230_000_000; c.y_axis.scaling.max = 325_000_000; rainbow(c, 12); c.legend = None
place(c, "G4", "Doanh thu")
# 3 Line 3D
c = LineChart(); c.add_data(Reference(dl, min_col=5, min_row=3, max_row=15), titles_from_data=True); c.set_categories(ref(dl, 4, 4, 15))
place(c, "M4", "Xu hướng")
# 4 Bar 3D theo miền
c = BarChart3D(); c.add_data(Reference(dl, min_col=8, min_row=3, max_row=6), titles_from_data=True); c.set_categories(ref(dl, 7, 4, 6)); rainbow(c, 3)
place(c, "S4", "Miền")
# 5 Area
c = AreaChart(); c.add_data(Reference(dl, min_col=5, min_row=3, max_row=15), titles_from_data=True); c.set_categories(ref(dl, 4, 4, 15))
place(c, "A18", "Diện tích")
# 6 Doughnut miền
c = DoughnutChart(); c.add_data(Reference(dl, min_col=8, min_row=3, max_row=6), titles_from_data=True); c.set_categories(ref(dl, 7, 4, 6))
place(c, "G18", "Cơ cấu")
# 7 Radar ngành
c = RadarChart(); c.add_data(Reference(dl, min_col=2, min_row=3, max_row=9), titles_from_data=True); c.set_categories(ref(dl, 1, 4, 9))
place(c, "M18", "Radar")
# 8 Scatter cửa hàng (chỉ số thứ tự × doanh thu: không có ý nghĩa)
c = ScatterChart(); c.style = 13
x = Reference(dl, min_col=10, min_row=4, max_row=15)
c.series.append(Series(ref(dl, 11, 4, 15), Reference(dl, min_col=11, min_row=4, max_row=15), title="Điểm"))
place(c, "S18", "Phân tán")
# 9 Cột 12 cửa hàng cầu vồng, trục cắt khúc
c = BarChart(); c.type = "col"; c.add_data(Reference(dl, min_col=11, min_row=3, max_row=15), titles_from_data=True); c.set_categories(ref(dl, 10, 4, 15))
c.y_axis.scaling.min = 50_000_000; rainbow(c, 12); c.legend = None
place(c, "A32", "Cửa hàng")
# 10 Pie 3D miền
c = PieChart3D(); c.add_data(Reference(dl, min_col=8, min_row=3, max_row=6), titles_from_data=True); c.set_categories(ref(dl, 7, 4, 6))
place(c, "G32", "Miền (pie)")
# 11 Bar ngang ngành
c = BarChart(); c.type = "bar"; c.add_data(Reference(dl, min_col=2, min_row=3, max_row=9), titles_from_data=True); c.set_categories(ref(dl, 1, 4, 9)); rainbow(c, 6)
place(c, "M32", "So sánh")
# 12 Line 2025
c = LineChart(); c.add_data(Reference(dl, min_col=14, min_row=3, max_row=7), titles_from_data=True); c.set_categories(ref(dl, 13, 4, 7))
place(c, "S32", "Tháng 2025")
ws.sheet_view.showGridLines = False
ws.sheet_view.zoomScale = 60
assert len(ws._charts) == 12

# Demo_TrucY
ty = wb.create_sheet("Demo_TrucY")
ty["A1"] = "Trục Y: cùng một dữ liệu, hai cách vẽ"; ty["A1"].font = TITLE
header(ty, 3, ["Tháng", "Doanh thu (VNĐ)"])
put(ty, "A4", "T3/2024"); put(ty, "B4", rev24[3], INPUT, NUM)
put(ty, "A5", "T4/2024"); put(ty, "B5", rev24[4], INPUT, NUM)
put(ty, "D3", "Mốc trục bị cắt", BOLD); put(ty, "E3", 236_000_000, INPUT, NUM)
put(ty, "D4", "Chênh lệch thật (T4 so T3)", BOLD); put(ty, "E4", "=B5/B4-1", BODY, "0.00%")
put(ty, "D5", "Chiều cao cột T4 ÷ T3 khi trục bắt đầu từ mốc cắt", BOLD); put(ty, "E5", "=(B5-E3)/(B4-E3)", BODY, "0.00")
put(ty, "D6", "Chiều cao cột T4 ÷ T3 khi trục bắt đầu từ 0", BOLD); put(ty, "E6", "=B5/B4", BODY, "0.00")
widths(ty, [12, 18, 3, 52, 16])
for k, (mn, ttl, anc) in enumerate([(236_000_000, "Trục bắt đầu từ 236.000.000 (XẤU)", "A9"), (0, "Trục bắt đầu từ 0 (TỐT)", "H9")]):
    c = BarChart(); c.type = "col"; c.add_data(Reference(ty, min_col=2, min_row=3, max_row=5), titles_from_data=True); c.set_categories(ref(ty, 1, 4, 5))
    c.y_axis.scaling.min = mn; c.y_axis.scaling.max = 240_000_000 if mn else None; c.legend = None; c.title = ttl; c.width, c.height = 12, 8
    c.y_axis.number_format = "#,##0"; c.y_axis.delete = False; c.x_axis.delete = False
    ty.add_chart(c, anc)

g = wb.create_sheet("Huong_dan")
guide(g, "DEMO B6 — DASHBOARD XẤU (slide 2, 8, 9)", [
    (8, "Mở sheet Dashboard_Xau. Chiếu 5 giây rồi che màn hình; mỗi người ghi: KPI chính là gì · đạt hay chưa đạt target · nên nhìn đâu tiếp. Sau đó chiếu lại và liệt kê vấn đề.",
     "Không ai nói được đạt hay chưa (không có target). Đếm được ≥ 6 vấn đề: 12 chart; pie 3D; pie 6 phần (Sữa 1,8%); trục Y cắt khúc ở hai chart cột; màu cầu vồng; không có KPI, thứ bậc, slicer; tiêu đề chart là nhãn chung chung."),
    (8, "Mở sheet Du_lieu, đối chiếu số trên pie ngành với dữ liệu.",
     f"Tỷ trọng ngành T1–T4/2025 trong pie: " + " · ".join(f"{c} {cat[c][0] / ytd * 100:.1f}%".replace(".", ",") for c in CATS) + "; tổng " + f"{ytd:,}".replace(",", ".") + "."),
    (9, "Mở sheet Demo_TrucY, xem hai biểu đồ cột cạnh nhau (cùng dữ liệu T3 và T4/2024).",
     f"Chênh lệch thật {(rev24[4] / rev24[3] - 1) * 100:.2f}%".replace(".", ",") + f" (ô E4). Trục từ 236.000.000: cột T4 cao gấp {(rev24[4] - 236_000_000) / (rev24[3] - 236_000_000):.2f} lần cột T3 (ô E5)".replace(".", ",") + "; trục từ 0: gấp 1,01 lần (ô E6)."),
])
save_both(wb, "B6_Demo_DashboardXau", "Dashboard_Xau")

# =====================================================================
# 2. B6_Demo_Chart.xlsx
# =====================================================================
wb = Workbook(); ws = wb.active; ws.title = "Demo_Combo"
ws["A1"] = "Combo 2 trục: doanh thu (cột) + Gross Margin % (đường, trục phụ) theo ngành, T1–T4/2025"; ws["A1"].font = TITLE
header(ws, 3, ["Ngành", "Doanh thu (VNĐ)", "Gross Margin %"])
for k, c_ in enumerate(CATS, start=4):
    put(ws, f"A{k}", c_); put(ws, f"B{k}", cat[c_][0], INPUT, NUM); put(ws, f"C{k}", round(cat[c_][1], 4), INPUT, PCT)
put(ws, "A11", "Tổng", BOLD); put(ws, "B11", "=SUM(B4:B9)", BOLD, NUM)
put(ws, "A12", "GM% toàn chuỗi", BOLD); put(ws, "C12", "=SUMPRODUCT(B4:B9,C4:C9)/B11", BOLD, PCT)
widths(ws, [20, 18, 16])

# Waterfall tháng 3
wf = wb.create_sheet("Demo_Waterfall")
wf["A1"] = "Waterfall: từ Target đến Actual, tháng 3/2025, theo cửa hàng"; wf["A1"].font = TITLE
header(wf, 3, ["Mục", "Giá trị (VNĐ)"])
tg = pd.read_excel(f"{ROOT}/1_Sessions/B4_Power_Query/Practice/B4_Target_2025.xlsx", sheet_name="Target_2025", header=2).set_index("Store_Code")
t3 = {sc: int(tg.loc[sc, "Tháng 3"]) for sc in store.Store_Code}
a3 = d25[d25.M == 3].groupby("Store_Code").Revenue.sum()
put(wf, "A4", "Target T3"); put(wf, "B4", sum(t3.values()), INPUT, NUM)
for k, sc in enumerate(store.Store_Code, start=5):
    put(wf, f"A{k}", sc); put(wf, f"B{k}", int(a3[sc]) - t3[sc], INPUT, NUM)
put(wf, "A17", "Actual T3"); put(wf, "B17", "=B4+SUM(B5:B16)", BOLD, NUM)
put(wf, "D3", "Kiểm tra", BOLD)
put(wf, "D4", "Tổng chênh lệch (B5:B16)"); put(wf, "E4", "=SUM(B5:B16)", BODY, NUM)
put(wf, "D5", "Actual T3 (từ dữ liệu bán hàng)"); put(wf, "E5", int(d25[d25.M == 3].Revenue.sum()), INPUT, NUM)
put(wf, "D6", "B17 − E5 (phải = 0)"); put(wf, "E6", "=B17-E5", BODY, NUM)
put(wf, "D7", "ST07 ÷ tổng chênh lệch"); put(wf, "E7", "=B11/E4", BODY, PCT)
widths(wf, [14, 16, 3, 32, 16])

# Phân phối
ph = wb.create_sheet("Demo_Phan_phoi")
q1 = d25[d25.Q == 1].groupby("Order_ID").Revenue.sum()
header(ph, 1, ["Order_ID", "Gia_tri_don"])
for k, (oid, v) in enumerate(q1.items(), start=2):
    ph.cell(row=k, column=1, value=oid).font = BODY
    c = ph.cell(row=k, column=2, value=int(v)); c.font = BODY; c.number_format = NUM
last = len(q1) + 1
header(ph, 1, ["Chỉ số", "Giá trị"], col=4)
stats = [("Số đơn", f"=COUNT(B2:B{last})", NUM), ("Trung bình", f"=AVERAGE(B2:B{last})", NUM), ("Trung vị", f"=MEDIAN(B2:B{last})", NUM),
         ("Q1", f"=_xlfn.QUARTILE.INC(B2:B{last},1)", NUM), ("Q3", f"=_xlfn.QUARTILE.INC(B2:B{last},3)", NUM),
         ("Ngưỡng outlier = Q3 + 1,5 × IQR", "=E6+1.5*(E6-E5)", NUM), ("Số đơn vượt ngưỡng", f"=COUNTIF(B2:B{last},\">\"&E7)", NUM),
         ("Độ lệch (SKEW)", f"=SKEW(B2:B{last})", "0.00")]
for k, (a, f_, fm) in enumerate(stats, start=2):
    put(ph, f"D{k}", a); put(ph, f"E{k}", f_, BODY, fm)
widths(ph, [12, 14, 3, 34, 14])

# Scatter
sc_ = wb.create_sheet("Demo_Scatter")
sc_["A1"] = "Scatter: tuổi cửa hàng (năm, tính tới 30/04/2025) × doanh thu T1–T4/2025"; sc_["A1"].font = TITLE
header(sc_, 3, ["Cửa hàng", "Tuổi (năm)", "Doanh thu T1–T4 (VNĐ)", "Thành phố"])
stt = d25[d25.M <= 4].groupby("Store_Code").Revenue.sum()
for k, r in enumerate(store.itertuples(), start=4):
    age = round((pd.Timestamp("2025-04-30") - r.Open_Date).days / 365.25, 2)
    put(sc_, f"A{k}", r.Store_Code); put(sc_, f"B{k}", age, INPUT, "0.00"); put(sc_, f"C{k}", int(stt[r.Store_Code]), INPUT, NUM); put(sc_, f"D{k}", r.City)
put(sc_, "F3", "Hệ số tương quan r", BOLD); put(sc_, "G3", "=CORREL(B4:B15,C4:C15)", BODY, "0.00")
put(sc_, "F4", "R²", BOLD); put(sc_, "G4", "=G3^2", BODY, "0.00")
put(sc_, "F5", "r khi bỏ ST12 (hàng 15)", BOLD); put(sc_, "G5", "=CORREL(B4:B14,C4:C14)", BODY, "0.00")
widths(sc_, [12, 12, 22, 14, 3, 30, 12])
header(sc_, 17, ["Nhóm", "Tuổi TB (năm)", "Doanh thu TB (VNĐ)"])
put(sc_, "A18", "Hà Nội + TP.HCM"); put(sc_, "B18", '=(SUMIFS(B4:B15,D4:D15,"Hà Nội")+SUMIFS(B4:B15,D4:D15,"TP.HCM"))/(COUNTIFS(D4:D15,"Hà Nội")+COUNTIFS(D4:D15,"TP.HCM"))', BODY, "0.0")
put(sc_, "C18", '=(SUMIFS(C4:C15,D4:D15,"Hà Nội")+SUMIFS(C4:C15,D4:D15,"TP.HCM"))/(COUNTIFS(D4:D15,"Hà Nội")+COUNTIFS(D4:D15,"TP.HCM"))', BODY, NUM)
put(sc_, "A19", "Các thành phố còn lại"); put(sc_, "B19", "=(SUM(B4:B15)-B18*(COUNTIFS(D4:D15,\"Hà Nội\")+COUNTIFS(D4:D15,\"TP.HCM\")))/(12-COUNTIFS(D4:D15,\"Hà Nội\")-COUNTIFS(D4:D15,\"TP.HCM\"))", BODY, "0.0")
put(sc_, "C19", "=(SUM(C4:C15)-C18*(COUNTIFS(D4:D15,\"Hà Nội\")+COUNTIFS(D4:D15,\"TP.HCM\")))/(12-COUNTIFS(D4:D15,\"Hà Nội\")-COUNTIFS(D4:D15,\"TP.HCM\"))", BODY, NUM)

# Map
mp = wb.create_sheet("Demo_Map")
mp["A1"] = "Map theo thành phố: doanh thu tháng 4/2025"; mp["A1"].font = TITLE
header(mp, 3, ["Thành phố", "Doanh thu (VNĐ)"])
city = d25[d25.M == 4].groupby("City").Revenue.sum().sort_values(ascending=False)
for k, (cname, v) in enumerate(city.items(), start=4):
    put(mp, f"A{k}", cname); put(mp, f"B{k}", int(v), INPUT, NUM)
put(mp, "A13", "Tổng", BOLD); put(mp, "B13", "=SUM(B4:B12)", BOLD, NUM)
put(mp, "D3", "Nếu Filled Map không nhận tên (cần Internet): vẽ Bar theo cột A:B.", NOTE)
widths(mp, [16, 18, 3, 60])

g = wb.create_sheet("Huong_dan")
guide(g, "DEMO B6 — CHART LÕI (slide 11–14)", [
    (12, "Sheet Demo_Combo: chọn A3:C9 → Insert → Recommended Charts → All Charts → Combo. Doanh thu: Clustered Column; Gross Margin %: Line, tích Secondary Axis → OK. Cho trục phụ bắt đầu từ 0%.",
     "Cột doanh thu (trục trái) + đường GM% (trục phải). Thực phẩm khô: cột lớn thứ ba (222.348.000) nhưng GM thấp nhất 21,6%; Bánh kẹo 426.916.000 và GM 39,4%. Ô C12 (GM toàn chuỗi T1–T4) kiểm bằng SUMPRODUCT."),
    (13, "Sheet Demo_Waterfall: chọn A3:B17 → Insert → Waterfall. Chuột phải cột Target → Set as Total; cột Actual T3 → Set as Total.",
     "Target 285.000.000 → 12 bước cửa hàng → Actual 269.608.000 (ô B17). Tổng chênh lệch −15.392.000 (E4); B17 − E5 = 0 (E6). ST12 −11.749.000, ST07 −10.461.000, ST08 +10.733.000; ST07 chiếm 68% khoản hụt (E7)."),
    (14, "Sheet Demo_Phan_phoi: chọn cột B (từ B2) → Insert → Histogram. Chuột phải trục ngang → Format Axis → Bin width 100000. Đọc khối D1:E9.",
     "3.144 đơn; trung bình 299.445; trung vị 198.000; Q1 89.000; Q3 412.000; ngưỡng 896.500; 138 đơn vượt; SKEW ≈ 1,77. Cột cao nhất ở nhóm 0–100.000 (905 đơn), đuôi phải kéo dài tới hơn 2 triệu."),
    (14, "Sheet Demo_Scatter: chọn B3:C15 → Insert → Scatter. Add Trendline → Display R-squared. Đọc ô G3:G5 và B18:C19.",
     "r = 0,78 (G3), R² = 0,61 (G4); bỏ ST12 thì r = 0,90 (G5). Hà Nội + TP.HCM: tuổi TB 4,3 năm, doanh thu TB 137.519.200; các thành phố còn lại: 2,7 năm, 78.689.000 → thành phố lớn là yếu tố thứ ba."),
    (13, "Sheet Demo_Map: chọn A3:B12 → Insert → Maps → Filled Map (cần Internet).",
     "TP.HCM 108.885.000 đứng đầu, tiếp theo Hà Nội 62.411.000; tổng B13 = 296.965.000 (= doanh thu tháng 4). Không nhận được tên thì vẽ Bar theo thành phố."),
])
save_both(wb, "B6_Demo_Chart", "Demo_Combo")

# =====================================================================
# 3. B6_Demo_DuBao.xlsx
# =====================================================================
wb = Workbook(); ws = wb.active; ws.title = "Demo_Trendline"
ws["A1"] = "Bảng Ngày–Doanh thu MCI Mart: 01/01/2024–30/04/2025 (486 ngày liên tục)"; ws["A1"].font = TITLE
header(ws, 3, ["Ngay", "Doanh thu"])
for k, (d, v) in enumerate(daily.items(), start=4):
    c = ws.cell(row=k, column=1, value=d.to_pydatetime()); c.number_format = DATE; c.font = BODY
    c = ws.cell(row=k, column=2, value=int(v)); c.number_format = NUM; c.font = BODY
LAST = 3 + len(daily)
assert LAST == 489
tbl = Table(displayName="tblNgay", ref=f"A3:B{LAST}")
tbl.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
ws.add_table(tbl)
widths(ws, [14, 16])

# Thong_ke
tk = wb.create_sheet("Thong_ke")
tk["A1"] = "Thống kê trên bảng Ngày–Doanh thu"; tk["A1"].font = TITLE
header(tk, 3, ["STT", "Thu (1 = Thứ Hai)"])
for k in range(len(daily)):
    put(tk, f"A{k + 4}", k + 1); put(tk, f"B{k + 4}", f"=WEEKDAY(Demo_Trendline!A{k + 4},2)")
header(tk, 3, ["Thứ", "Doanh thu TB mỗi ngày"], col=4)
names = ["Thứ Hai", "Thứ Ba", "Thứ Tư", "Thứ Năm", "Thứ Sáu", "Thứ Bảy", "Chủ nhật"]
for i, nme in enumerate(names, start=1):
    put(tk, f"D{i + 3}", nme); put(tk, f"E{i + 3}", f"=AVERAGEIFS(tblNgay[Doanh thu],$B$4:$B${LAST},{i})", BODY, NUM)
rows_ = [("Chủ nhật ÷ Thứ Hai − 1", "=E10/E4-1", PCT), ("Cuối tuần ÷ ngày thường − 1", '=AVERAGEIFS(tblNgay[Doanh thu],$B$4:$B$489,">=6")/AVERAGEIFS(tblNgay[Doanh thu],$B$4:$B$489,"<=5")-1', PCT),
         ("Số ngày", "=COUNT(tblNgay[Doanh thu])", NUM), ("Tổng doanh thu", "=SUM(tblNgay[Doanh thu])", NUM),
         ("Độ lệch chuẩn doanh thu ngày", "=STDEV.S(tblNgay[Doanh thu])", NUM),
         ("Độ dốc Trendline (VNĐ/ngày, x = STT)", f"=SLOPE(tblNgay[Doanh thu],$A$4:$A${LAST})", NUM),
         ("R² của Trendline tuyến tính", f"=RSQ(tblNgay[Doanh thu],$A$4:$A${LAST})", "0.000"),
         ("Dự báo tuyến tính ngày STT 487 (01/05/2025)", f"=FORECAST.LINEAR(487,tblNgay[Doanh thu],$A$4:$A${LAST})", NUM)]
for i, (a, f_, fm) in enumerate(rows_, start=13):
    put(tk, f"D{i}", a); put(tk, f"E{i}", f_.replace("FORECAST.LINEAR", "_xlfn.FORECAST.LINEAR").replace("STDEV.S", "_xlfn.STDEV.S"), BODY, fm)
widths(tk, [8, 18, 3, 44, 22])

# Demo_ETS
es = wb.create_sheet("Demo_ETS")
es["A1"] = "Hàm dự báo ETS trên bảng tblNgay (Forecast Sheet: Data → Forecast → Forecast Sheet, tạo sheet mới)"; es["A1"].font = TITLE
put(es, "A3", "Ngày dự báo", BOLD); put(es, "B3", __import__("datetime").datetime(2025, 5, 1), INPUT, DATE)
put(es, "A4", "FORECAST.ETS (mùa 7)", BOLD); put(es, "B4", "=_xlfn.FORECAST.ETS(B3,tblNgay[Doanh thu],tblNgay[Ngay],7)", BODY, NUM)
put(es, "A5", "FORECAST.ETS.CONFINT 95% (mùa 7)", BOLD); put(es, "B5", "=_xlfn.FORECAST.ETS.CONFINT(B3,tblNgay[Doanh thu],tblNgay[Ngay],0.95,7)", BODY, NUM)
put(es, "A6", "FORECAST.ETS.SEASONALITY (tự phát hiện)", BOLD); put(es, "B6", "=_xlfn.FORECAST.ETS.SEASONALITY(tblNgay[Doanh thu],tblNgay[Ngay])", BODY, "0")
put(es, "A7", "Cận dưới 95%", BOLD); put(es, "B7", "=B4-B5", BODY, NUM)
put(es, "A8", "Cận trên 95%", BOLD); put(es, "B8", "=B4+B5", BODY, NUM)
widths(es, [42, 18])

g = wb.create_sheet("Huong_dan")
guide(g, "DEMO B6 — DỰ BÁO NHẬP MÔN (slide 23–24)", [
    (23, "Sheet Demo_Trendline: bảng tblNgay (Ngay, Doanh thu). Chọn bảng → Insert → Line chart. Chart Elements → Trendline → Linear; Format Trendline → Forecast Forward 30 periods; tích Display R-squared.",
     "486 dòng, liên tục 01/01/2024–30/04/2025 (đủ 486 ngày). R² = 0,045; độ dốc ≈ +4.742 VNĐ/ngày (sheet Thong_ke, E18:E19). Dự báo ngày 487 (01/05/2025) ≈ 10,28 triệu (E20)."),
    (23, "Sheet Thong_ke: đọc doanh thu trung bình theo thứ (D4:E10).",
     "Thứ Hai 8.434.264 · Ba 8.643.886 · Tư 8.696.329 · Năm 8.816.072 · Sáu 8.722.935 · Bảy 10.067.797 · Chủ nhật 10.547.341. Chủ nhật hơn thứ Hai 25,1%; cuối tuần hơn ngày thường 19,0%."),
    (24, "Chọn một ô trong tblNgay → Data → Forecast → Forecast Sheet → Options: Forecast End 30/05/2025; Confidence Interval 95%; Seasonality → Set Manually 7 → Create.",
     "Sheet mới có cột Forecast, Lower Confidence Bound, Upper Confidence Bound và biểu đồ có răng cưa cuối tuần. Thử thêm lần Detect Automatically để so với Set Manually 7."),
    (24, "Sheet Demo_ETS: đọc B3:B8 (hàm FORECAST.ETS, FORECAST.ETS.CONFINT, FORECAST.ETS.SEASONALITY).",
     "B4 là điểm dự báo ngày 01/05/2025; B5 là nửa chiều rộng dải 95%; B6 là chu kỳ Excel tự phát hiện. Ghi lại các số này để so với thực tế."),
    (24, "Kiểm chứng ngược: trong hộp Forecast Sheet, đặt Forecast Start = 01/04/2025 (trước ngày cuối dữ liệu), so đường Forecast với thực tế.",
     "Thực tế tháng 4/2025 = 296.965.000 (trung bình 9.898.833 mỗi ngày). Ghi tổng 30 ngày dự báo để so sánh."),
])
save_both(wb, "B6_Demo_DuBao", "Demo_Trendline")
print("ok:", sorted(os.listdir(DEMO)))
