"""Sinh dataset capstone B7: "Hạt Việt Coffee" (chuỗi cà phê giả lập, 10 chi nhánh, 3 miền).

Dữ liệu giả lập, seed cố định (2026) nên chạy lại luôn ra đúng cùng một bộ số. Mọi tên đều hư cấu.
    python3 _generator/gen_capstone.py            # từ thư mục 2_Dataset (cần openpyxl)

Đầu ra:
  1_Sessions/B7_Capstone/Du_lieu_Capstone/            PHÁT HỌC VIÊN (cuối B6)
      Sales/Sales_2024.csv, Sales_2025-01.csv … Sales_2025-11.csv
      HatViet_Lookups.xlsx        (Table: tblProduct, tblBranch, tblChannel, tblCustomer)
      HatViet_Target_2025.xlsx    (dạng cột tháng, tiêu đề ở dòng 3, dòng tổng cuối)
  1_Sessions/B7_Capstone/GiangVien/                   KHÔNG PHÁT
      New_Month/Sales_2025-12.csv  (file mới cho bài kiểm tra Refresh, C2)
      Dap_an_Capstone_HatViet.xlsx (số đối soát, 4 câu chuyện, kiểm Refresh)
  2_Dataset/_generator/build_capstone_meta.json         (số liệu cho brief, đáp án)

Bốn câu chuyện cài sẵn: (1) kênh giao hàng tăng nhanh nhưng hoa hồng 22–25% kéo biên sau hoa hồng xuống; (2) bão Miền Trung
07–20/10/2025; (3) HV10 Cần Thơ khai trương 01/09/2025 (YoY trống, like-for-like); (4) Cold Brew ra mắt 01/06/2025 (mix, một phần ăn
vào Cà phê sữa đá). Hiệu ứng nền: Tết (10/02/2024, 29/01/2025), tăng giá nhóm Cà phê 5% từ 01/07/2025, cuối tuần cao hơn.
Lỗi dữ liệu cài sẵn (D2): xem DATA_ISSUES ở cuối file này và sheet Dap_an của file giảng viên.
"""
import bisect, csv, datetime as dt, json, os, random
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

ROOT = os.path.normpath(os.environ.get("MCI_ROOT") or os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
GEN = os.path.dirname(os.path.abspath(__file__))
B7 = f"{ROOT}/1_Sessions/B7_Capstone"
OUT_STU = f"{B7}/Du_lieu_Capstone"
OUT_INS = f"{B7}/GiangVien"
for p in (f"{OUT_STU}/Sales", f"{OUT_INS}/New_Month"):
    os.makedirs(p, exist_ok=True)
rng = random.Random(2026)
D = dt.date

# ---------------------------------------------------------------- danh mục
BRANCH = [  # code, name, city, region, open, base orders/day (cửa hàng đã trưởng thành), diện tích m2, số chỗ
    ("HV01", "Hạt Việt Hoàn Kiếm", "Hà Nội", "Bắc", D(2022, 3, 1), 12, 85, 48),
    ("HV02", "Hạt Việt Cầu Giấy", "Hà Nội", "Bắc", D(2022, 9, 15), 11, 78, 44),
    ("HV03", "Hạt Việt Lê Chân", "Hải Phòng", "Bắc", D(2023, 4, 10), 8, 66, 36),
    ("HV04", "Hạt Việt Hải Châu", "Đà Nẵng", "Trung", D(2022, 6, 1), 10, 80, 46),
    ("HV05", "Hạt Việt Vĩnh Ninh", "Huế", "Trung", D(2023, 1, 20), 8, 70, 40),
    ("HV06", "Hạt Việt Quy Nhơn", "Quy Nhơn", "Trung", D(2023, 8, 5), 7, 64, 34),
    ("HV07", "Hạt Việt Quận 1", "TP.HCM", "Nam", D(2021, 11, 15), 14, 92, 52),
    ("HV08", "Hạt Việt Quận 7", "TP.HCM", "Nam", D(2022, 5, 20), 12, 88, 50),
    ("HV09", "Hạt Việt Thủ Đức", "TP.HCM", "Nam", D(2023, 10, 1), 10, 75, 42),
    ("HV10", "Hạt Việt Ninh Kiều", "Cần Thơ", "Nam", D(2025, 9, 1), 9, 72, 40),
]
BCODES = [b[0] for b in BRANCH]
TRUNG = {"HV04", "HV05", "HV06"}
BASE = {b[0]: b[5] for b in BRANCH}
OPEN = {b[0]: b[4] for b in BRANCH}

CHANNEL = [("DINE", "Tại quán", 0.0), ("TAKE", "Mang đi", 0.0), ("GRAB", "GrabFood", 0.25), ("SHOPEE", "ShopeeFood", 0.22)]
COMM = {c[0]: c[2] for c in CHANNEL}

# sku, tên, nhóm, giá niêm yết 2024, giá vốn, trọng số bán, ngày ra mắt
PRODUCT = [
    ("CF01", "Cà phê đen đá", "Cà phê", 29000, 8000, 9, D(2021, 1, 1)),
    ("CF02", "Cà phê sữa đá", "Cà phê", 32000, 9000, 20, D(2021, 1, 1)),
    ("CF03", "Bạc xỉu", "Cà phê", 35000, 10000, 11, D(2021, 1, 1)),
    ("CF04", "Cappuccino", "Cà phê", 49000, 14000, 6, D(2021, 1, 1)),
    ("CF05", "Latte", "Cà phê", 52000, 15000, 6, D(2021, 1, 1)),
    ("CF06", "Americano", "Cà phê", 42000, 10000, 4, D(2021, 1, 1)),
    ("CF07", "Espresso", "Cà phê", 35000, 8000, 2, D(2021, 1, 1)),
    ("CF08", "Cold Brew", "Cà phê", 55000, 13000, 0, D(2025, 6, 1)),
    ("TR01", "Trà đào cam sả", "Trà", 45000, 12000, 12, D(2021, 1, 1)),
    ("TR02", "Trà vải", "Trà", 42000, 11000, 7, D(2021, 1, 1)),
    ("TR03", "Trà sen vàng", "Trà", 45000, 13000, 6, D(2021, 1, 1)),
    ("TR04", "Trà chanh", "Trà", 32000, 7000, 6, D(2021, 1, 1)),
    ("TR05", "Trà ô long sữa", "Trà", 48000, 13000, 5, D(2021, 1, 1)),
    ("TR06", "Matcha latte", "Trà", 55000, 18000, 4, D(2022, 1, 1)),
    ("TR07", "Trà hoa cúc", "Trà", 38000, 9000, 2, D(2021, 1, 1)),
    ("DX01", "Matcha đá xay", "Đá xay", 59000, 19000, 4, D(2021, 1, 1)),
    ("DX02", "Cookie đá xay", "Đá xay", 62000, 21000, 4, D(2021, 1, 1)),
    ("DX03", "Socola đá xay", "Đá xay", 59000, 20000, 3, D(2021, 1, 1)),
    ("DX04", "Dâu đá xay", "Đá xay", 58000, 19000, 2, D(2021, 1, 1)),
    ("DX05", "Cà phê đá xay", "Đá xay", 55000, 15000, 3, D(2021, 1, 1)),
    ("BK01", "Croissant bơ", "Bánh", 39000, 16000, 5, D(2021, 1, 1)),
    ("BK02", "Bánh tiramisu", "Bánh", 55000, 24000, 3, D(2021, 1, 1)),
    ("BK03", "Cheesecake chanh dây", "Bánh", 58000, 25000, 3, D(2021, 1, 1)),
    ("BK04", "Bánh mì que", "Bánh", 25000, 9000, 4, D(2021, 1, 1)),
    ("BK05", "Cookie socola", "Bánh", 30000, 12000, 3, D(2021, 1, 1)),
    ("BK06", "Bông lan trứng muối", "Bánh", 45000, 19000, 3, D(2021, 1, 1)),
    ("HT01", "Hạt rang Robusta 250g", "Hạt rang & quà", 120000, 62000, 0.7, D(2021, 1, 1)),
    ("HT02", "Hạt rang Arabica 250g", "Hạt rang & quà", 165000, 92000, 0.5, D(2021, 1, 1)),
    ("HT03", "Drip bag hộp 10 gói", "Hạt rang & quà", 110000, 52000, 0.5, D(2021, 1, 1)),
    ("HT04", "Hộp quà Hạt Việt", "Hạt rang & quà", 450000, 230000, 0, D(2021, 1, 1)),
]
PINFO = {p[0]: p for p in PRODUCT}
PRICE_RISE = D(2025, 7, 1)          # Cà phê tăng 5%
COLD_LAUNCH = D(2025, 6, 1)


def price_on(sku, d):
    p = PINFO[sku][3]
    if PINFO[sku][2] == "Cà phê" and d >= PRICE_RISE:
        p = round(p * 1.05 / 1000) * 1000
    return p


def price_current(sku):
    return price_on(sku, D(2025, 12, 31))


TET = {2024: D(2024, 2, 10), 2025: D(2025, 1, 29)}
GIFT_WINDOWS = [(D(2024, 1, 10), D(2024, 2, 8)), (D(2024, 12, 15), D(2025, 1, 27)), (D(2025, 11, 15), D(2025, 12, 31))]
SEASON = {1: 1.00, 2: 0.95, 3: 1.00, 4: 1.02, 5: 1.06, 6: 1.10, 7: 1.12, 8: 1.10, 9: 1.02, 10: 0.98, 11: 0.96, 12: 1.08}
DOW = [0.92, 0.93, 0.95, 0.98, 1.08, 1.18, 1.15]          # thứ Hai … Chủ nhật
GROWTH_2025 = 1.04


def delivery_share(d):
    """Tỷ trọng đơn của hai kênh giao hàng theo tháng (câu chuyện 1)."""
    if d.year == 2024:
        return 0.07 + 0.02 * (d.month - 1) / 11
    return 0.10 + 0.21 * ((d.month - 1) / 11)           # 10% (T1/2025) → 31% (T12/2025)


def event_factor(code, d):
    f = 1.0
    for y, t in TET.items():
        k = (d - t).days
        if -2 <= k <= 4:
            f *= 0.05 if k in (0, 1) else 0.2
        elif -14 <= k <= -3:
            f *= 1.10
    if code in TRUNG and D(2025, 10, 7) <= d <= D(2025, 10, 20):    # bão Miền Trung (câu chuyện 2)
        f *= 0.0 if d <= D(2025, 10, 9) else (0.35 if d <= D(2025, 10, 14) else 0.7)
    if code == "HV10":                                              # khai trương (câu chuyện 3)
        f *= min(1.0, 0.45 + 0.55 * (d - OPEN["HV10"]).days / 90)
    return f


def in_gift(d):
    return any(a <= d <= b for a, b in GIFT_WINDOWS)


_WCACHE = {}


def weights(d, channel):
    """Trọng số sản phẩm theo ngày/kênh; Cold Brew ăn một phần vào CF02, CF03 (câu chuyện 4)."""
    gift = in_gift(d)
    key = (d.year, d.month if d >= COLD_LAUNCH else 0, channel in ("GRAB", "SHOPEE"), gift)
    if key in _WCACHE:
        return _WCACHE[key]
    cb_level = 0.0
    if d >= COLD_LAUNCH:
        cb_level = {6: 0.25, 7: 0.4, 8: 0.55, 9: 0.7, 10: 0.8, 11: 0.9, 12: 1.0}[d.month]
    skus, w = [], []
    for sku, name, cat, price, cost, wt, launch in PRODUCT:
        if d < launch:
            continue
        x = wt
        if sku == "CF08":
            x = 10 * cb_level
        if sku in ("CF02", "CF03") and cb_level:
            x *= (1 - 0.25 * cb_level) if sku == "CF02" else (1 - 0.10 * cb_level)
        if sku == "HT04":
            x = 1.2 if gift else 0
        if channel in ("GRAB", "SHOPEE"):
            if cat in ("Trà", "Đá xay"):
                x *= 1.35
            if sku in ("CF01", "CF07", "CF06"):
                x *= 0.6
            if cat == "Hạt rang & quà":
                x *= 0.4
        skus.append(sku); w.append(x)
    _WCACHE[key] = (skus, w)
    return skus, w


# ---------------------------------------------------------------- khách hàng
CUST = []
n_cust = 1200
home_w = [BASE[c] for c in BCODES]
for i in range(1, n_cust + 1):
    hb = rng.choices(BCODES, home_w)[0]
    lo = max(D(2022, 1, 1), OPEN[hb])
    hi = D(2025, 10, 31)
    jd = lo + dt.timedelta(days=int(rng.random() ** 1.3 * (hi - lo).days))
    tier = rng.choices(["Standard", "Silver", "Gold"], [70, 22, 8])[0]
    CUST.append((f"C{i:04d}", jd, tier, hb))
CUST_BY_BRANCH = defaultdict(list)
for c in CUST:
    CUST_BY_BRANCH[c[3]].append(c)
for k in CUST_BY_BRANCH:
    CUST_BY_BRANCH[k].sort(key=lambda c: c[1])
ALL_SORTED = sorted(CUST, key=lambda c: c[1])


def pick_customer(branch, d):
    pool = CUST_BY_BRANCH[branch] if rng.random() < 0.65 else ALL_SORTED
    jds = [c[1] for c in pool]
    k = bisect.bisect_right(jds, d)
    if k == 0:
        return ""
    return pool[rng.randrange(k)][0]


# ---------------------------------------------------------------- sinh dòng bán hàng
HOUR_W = [0, 0, 0, 0, 0, 0, 1, 6, 10, 8, 6, 5, 7, 8, 6, 6, 7, 6, 4, 3, 1, 0, 0, 0]


def gen_day(d):
    rows = []
    for code in BCODES:
        if d < OPEN[code]:
            continue
        exp = BASE[code] * DOW[d.weekday()] * SEASON[d.month] * event_factor(code, d)
        if d.year == 2025:
            exp *= GROWTH_2025
        n = int(exp)
        if rng.random() < exp - n:
            n += 1
        n = max(0, n + rng.choice([-1, 0, 0, 1]) if exp > 3 else n)
        if event_factor(code, d) == 0.0:
            n = 0
        share = delivery_share(d)
        for s in range(1, n + 1):
            r = rng.random()
            if r < share:
                ch = "GRAB" if rng.random() < 0.72 else "SHOPEE"
            else:
                ch = "DINE" if rng.random() < 0.66 else "TAKE"
            deliv = ch in ("GRAB", "SHOPEE")
            lines = rng.choices([1, 2, 3, 4], [0.30, 0.40, 0.20, 0.10] if deliv else [0.55, 0.33, 0.12, 0.0])[0]
            if deliv:
                disc = rng.choices([0, 10, 20], [55, 25, 20])[0]
            else:
                disc = rng.choices([0, 10, 20], [88, 8, 4])[0]
            hr = rng.choices(range(24), HOUR_W)[0]
            ts = dt.datetime(d.year, d.month, d.day, hr, rng.randrange(60))
            cid = pick_customer(code, d) if rng.random() < (0.04 if deliv else 0.42) else ""
            oid = f"{code}-{d:%y%m%d}-{s:04d}"
            skus, w = weights(d, ch)
            chosen = []
            for ln in range(1, lines + 1):
                sku = rng.choices(skus, w)[0]
                if sku in chosen and sku != "HT04":
                    continue
                chosen.append(sku)
            for ln, sku in enumerate(chosen, 1):
                q = rng.choices([1, 2, 3], [82, 15, 3])[0]
                if sku == "HT04":
                    q = 1
                rows.append([oid, ln, ts, code, ch, cid, sku, q, price_on(sku, d), disc])
    return rows


def daterange(a, b):
    d = a
    while d <= b:
        yield d
        d += dt.timedelta(days=1)


ALL = []
for d in daterange(D(2024, 1, 1), D(2025, 12, 31)):
    ALL.extend(gen_day(d))
print("tổng dòng", len(ALL))


def revenue(r):
    return r[7] * r[8] * (100 - r[9]) / 100


def cogs(r):
    return r[7] * PINFO[r[6]][4]


# ---------------------------------------------------------------- xuất CSV (kèm lỗi cài sẵn cho các file 2025)
DATA_ISSUES = []


def fmt_row(r, sku_case=None, pad_branch=False):
    oid, ln, ts, code, ch, cid, sku, q, price, disc = r
    return [oid, ln, ts.strftime("%d/%m/%Y %H:%M"), code + (" " if pad_branch else ""), ch, cid,
            sku.lower() if sku_case == "lower" else sku, q, f"{price:,}".replace(",", "."), disc]


HEADER = ["Order_ID", "Line_No", "Order_Time", "Branch_Code", "Channel_Code", "Customer_ID", "SKU", "Qty", "Unit_Price", "Discount_Pct"]


def write_csv(path, rows, messy):
    """messy=True: cài lỗi — Branch_Code dính dấu cách, SKU viết thường, 1 dòng trùng tuyệt đối."""
    idx = list(range(len(rows)))
    pad = set(rng.sample(idx, 18)) if messy else set()
    low = set(rng.sample(idx, 5)) if messy else set()
    out = [fmt_row(r, "lower" if i in low else None, i in pad) for i, r in enumerate(rows)]
    dup_at = None
    if messy:
        dup_at = rng.randrange(len(out))
        out.insert(dup_at + 1, list(out[dup_at]))
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f); w.writerow(HEADER); w.writerows(out)
    return {"file": os.path.basename(path), "raw_rows": len(out), "clean_rows": len(rows), "pad": len(pad), "lower": len(low), "dup": 1 if messy else 0}


files = []
y2024 = [r for r in ALL if r[2].year == 2024]
files.append(write_csv(f"{OUT_STU}/Sales/Sales_2024.csv", y2024, False))
months = {}
for r in ALL:
    if r[2].year == 2025:
        months.setdefault(r[2].month, []).append(r)
for m in range(1, 12):
    files.append(write_csv(f"{OUT_STU}/Sales/Sales_2025-{m:02d}.csv", months[m], True))
files.append(write_csv(f"{OUT_INS}/New_Month/Sales_2025-12.csv", months[12], True))

# ---------------------------------------------------------------- Lookups
BOLD = Font(name="Arial", bold=True, color="FFFFFF")
HFILL = PatternFill("solid", fgColor="203864")
NORM = Font(name="Arial", size=10)


def add_table(ws, name, header, rows, start_row=1, widths=None):
    for j, h in enumerate(header, 1):
        c = ws.cell(row=start_row, column=j, value=h); c.font = BOLD; c.fill = HFILL
    for i, row in enumerate(rows, start_row + 1):
        for j, v in enumerate(row, 1):
            c = ws.cell(row=i, column=j, value=v); c.font = NORM
    ref = f"A{start_row}:{get_column_letter(len(header))}{start_row + len(rows)}"
    t = Table(displayName=name, ref=ref); t.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
    ws.add_table(t)
    for j in range(1, len(header) + 1):
        ws.column_dimensions[get_column_letter(j)].width = (widths or {}).get(j, 18)


wb = Workbook(); ws = wb.active; ws.title = "Product"
prod_rows = []
for sku, name, cat, price, cost, wt, launch in PRODUCT:
    prod_rows.append([sku, name, cat, price_current(sku), cost, launch, D(2024, 6, 1) if sku == "BK03" else launch])
# dòng cũ trùng khóa BK03: đặt TRƯỚC bản hiện hành (D2: khóa lookup phải duy nhất)
bk = next(i for i, r in enumerate(prod_rows) if r[0] == "BK03")
prod_rows.insert(bk, ["BK03", "Cheesecake chanh dây", "Bánh", 52000, 23000, D(2021, 1, 1), D(2023, 11, 1)])
for r in prod_rows:
    r[5] = dt.datetime.combine(r[5], dt.time()); r[6] = dt.datetime.combine(r[6], dt.time())
add_table(ws, "tblProduct", ["SKU", "Product_Name", "Category", "Unit_Price", "Unit_Cost", "Launch_Date", "Updated_At"], prod_rows,
          widths={1: 8, 2: 26, 3: 18, 4: 12, 5: 12, 6: 13, 7: 13})
for row in ws.iter_rows(min_row=2, min_col=6, max_col=7):
    for c in row: c.number_format = "dd/mm/yyyy"
for row in ws.iter_rows(min_row=2, min_col=4, max_col=5):
    for c in row: c.number_format = "#,##0"
ws2 = wb.create_sheet("Branch")
add_table(ws2, "tblBranch", ["Branch_Code", "Branch_Name", "City", "Region", "Open_Date", "Area_m2", "Seats"],
          [[b[0], b[1], b[2], b[3], dt.datetime.combine(b[4], dt.time()), b[6], b[7]] for b in BRANCH], widths={2: 24})
for row in ws2.iter_rows(min_row=2, min_col=5, max_col=5):
    for c in row: c.number_format = "dd/mm/yyyy"
ws3 = wb.create_sheet("Channel")
add_table(ws3, "tblChannel", ["Channel_Code", "Channel_Name", "Commission_Rate"], [[c[0], c[1], c[2]] for c in CHANNEL])
for row in ws3.iter_rows(min_row=2, min_col=3, max_col=3):
    for c in row: c.number_format = "0%"
ws4 = wb.create_sheet("Customer")
add_table(ws4, "tblCustomer", ["Customer_ID", "Join_Date", "Tier", "Home_Branch"],
          [[c[0], dt.datetime.combine(c[1], dt.time()), c[2], c[3]] for c in CUST])
for row in ws4.iter_rows(min_row=2, min_col=2, max_col=2):
    for c in row: c.number_format = "dd/mm/yyyy"
wb.save(f"{OUT_STU}/HatViet_Lookups.xlsx")

# ---------------------------------------------------------------- tổng hợp số liệu (cho Target, đáp án, câu chuyện)
def ym(r): return (r[2].year, r[2].month)
rev_bm = defaultdict(float)                       # (branch, y, m)
for r in ALL:
    rev_bm[(r[3], r[2].year, r[2].month)] += revenue(r)

# Target 2025 = doanh thu cùng tháng 2024 × 1,08 (không điều chỉnh Tết/bão) — làm tròn 100.000; HV10 từ tháng 9
TARGET = {}
for code in BCODES:
    for m in range(1, 13):
        if code == "HV10":
            if m >= 9:
                peers = [rev_bm[(c, 2024, m)] for c in ("HV08", "HV09")]
                TARGET[(code, m)] = round(sum(peers) / 2 * 1.08 * 0.70 / 100000) * 100000
            else:
                TARGET[(code, m)] = None
        else:
            TARGET[(code, m)] = round(rev_bm[(code, 2024, m)] * 1.08 / 100000) * 100000
wb = Workbook(); ws = wb.active; ws.title = "Target_2025"
ws["A1"] = "TARGET DOANH THU 2025 — HẠT VIỆT COFFEE"; ws["A1"].font = Font(name="Arial", bold=True, size=13)
ws["A2"] = "Đơn vị: VNĐ. Lập tháng 12/2024 theo doanh thu cùng tháng năm 2024. HV10 khai trương 01/09/2025, target từ tháng 9."; ws["A2"].font = Font(name="Arial", italic=True, size=9)
hdr = ["Branch_Code", "Branch_Name"] + [f"Tháng {m}" for m in range(1, 13)]
for j, h in enumerate(hdr, 1):
    c = ws.cell(row=3, column=j, value=h); c.font = BOLD; c.fill = HFILL
for i, b in enumerate(BRANCH, 4):
    ws.cell(row=i, column=1, value=b[0]); ws.cell(row=i, column=2, value=b[1])
    for m in range(1, 13):
        v = TARGET[(b[0], m)]
        if v is not None:
            ws.cell(row=i, column=2 + m, value=v).number_format = "#,##0"
tot_row = 4 + len(BRANCH)
ws.cell(row=tot_row, column=1, value="TỔNG CỘNG").font = Font(name="Arial", bold=True)
for m in range(1, 13):
    ws.cell(row=tot_row, column=2 + m, value=sum(TARGET[(c, m)] or 0 for c in BCODES)).number_format = "#,##0"
ws.column_dimensions["A"].width = 14; ws.column_dimensions["B"].width = 24
for j in range(3, 15):
    ws.column_dimensions[get_column_letter(j)].width = 14
wb.save(f"{OUT_STU}/HatViet_Target_2025.xlsx")

# ---------------------------------------------------------------- số liệu đối soát và câu chuyện
def agg(rows):
    rev = sum(revenue(r) for r in rows); cg = sum(cogs(r) for r in rows)
    comm = sum(revenue(r) * COMM[r[4]] for r in rows)
    orders = len({r[0] for r in rows})
    return {"rows": len(rows), "orders": orders, "qty": sum(r[7] for r in rows), "rev": round(rev), "cogs": round(cg),
            "comm": round(comm), "gm": (rev - cg) / rev if rev else None, "gm_net": (rev - cg - comm) / rev if rev else None,
            "aov": rev / orders if orders else None}


META = {"files": files, "months": {}, "totals": {}}
for y in (2024, 2025):
    META["totals"][str(y)] = agg([r for r in ALL if r[2].year == y and (y == 2024 or r[2].month <= 11)])
META["totals"]["2025_12"] = agg(months[12])
for y in (2024, 2025):
    for m in range(1, 13):
        META["months"][f"{y}-{m:02d}"] = agg([r for r in ALL if ym(r) == (y, m)])
# 1) kênh giao hàng
ch = {}
for m in range(1, 13):
    rows = [r for r in ALL if ym(r) == (2025, m)]
    tot = sum(revenue(r) for r in rows)
    ch[m] = {"deliv_share_rev": sum(revenue(r) for r in rows if r[4] in ("GRAB", "SHOPEE")) / tot,
             "deliv_share_orders": len({r[0] for r in rows if r[4] in ("GRAB", "SHOPEE")}) / len({r[0] for r in rows}),
             "gm": META["months"][f"2025-{m:02d}"]["gm"], "gm_net": META["months"][f"2025-{m:02d}"]["gm_net"],
             "rev": META["months"][f"2025-{m:02d}"]["rev"]}
META["story1_delivery"] = ch
META["story1_aov_nov"] = {c: agg([r for r in ALL if ym(r) == (2025, 11) and r[4] == c])["aov"] for c in COMM}
# 2) bão
st2 = {}
for m in (9, 10, 11):
    for code in sorted(TRUNG):
        a = rev_bm[(code, 2025, m)]; t = TARGET[(code, m)]
        st2[f"{code}-{m}"] = {"rev": round(a), "target": t, "pct": a / t}
    nt = [c for c in BCODES if c not in TRUNG and c != "HV10"]
    a = sum(rev_bm[(c, 2025, m)] for c in nt); t = sum(TARGET[(c, m)] for c in nt)
    st2[f"khac-{m}"] = {"rev": round(a), "target": t, "pct": a / t}
st2["orders_days_off"] = {code: sum(1 for d in daterange(D(2025, 10, 7), D(2025, 10, 9)) if not any(r[3] == code and r[2].date() == d for r in ALL if r[2].date() == d)) for code in sorted(TRUNG)}
META["story2_storm"] = st2
# 3) HV10
st3 = {}
for m in range(9, 13):
    a = rev_bm[("HV10", 2025, m)]; t = TARGET[("HV10", m)]
    all25 = sum(rev_bm[(c, 2025, m)] for c in BCODES); all24 = sum(rev_bm[(c, 2024, m)] for c in BCODES)
    lfl25 = sum(rev_bm[(c, 2025, m)] for c in BCODES if c != "HV10")
    st3[m] = {"rev": round(a), "target": t, "pct": a / t, "yoy_all": all25 / all24 - 1, "yoy_lfl": lfl25 / all24 - 1}
META["story3_hv10"] = st3
# 4) Cold Brew
st4 = {}
for m in range(6, 13):
    rows = [r for r in ALL if ym(r) == (2025, m) and PINFO[r[6]][2] == "Cà phê"]
    cf = sum(revenue(r) for r in rows)
    cb = sum(revenue(r) for r in rows if r[6] == "CF08")
    c2 = sum(revenue(r) for r in rows if r[6] == "CF02")
    gm_cf = (cf - sum(cogs(r) for r in rows)) / cf
    st4[m] = {"cf_rev": round(cf), "cb_share_of_cf": cb / cf, "cf02_share_of_cf": c2 / cf, "cf_gm": gm_cf}
META["story4_coldbrew"] = st4
META["story4_cf02_share_may"] = None
rows_may = [r for r in ALL if ym(r) == (2025, 5) and PINFO[r[6]][2] == "Cà phê"]
META["story4_cf02_share_may"] = sum(revenue(r) for r in rows_may if r[6] == "CF02") / sum(revenue(r) for r in rows_may)
# tổng kết đối soát theo file
rec = []
for f in files:
    rec.append(f)
META["checks"] = {
    "sales_files_students": ["Sales_2024.csv"] + [f"Sales_2025-{m:02d}.csv" for m in range(1, 12)],
    "total_rows_clean_students": META["totals"]["2024"]["rows"] + META["totals"]["2025"]["rows"],
    "total_rev_students": META["totals"]["2024"]["rev"] + META["totals"]["2025"]["rev"],
    "dec_rows": META["totals"]["2025_12"]["rows"], "dec_rev": META["totals"]["2025_12"]["rev"],
    "dec_orders": META["totals"]["2025_12"]["orders"],
    "total_rows_after_dec": META["totals"]["2024"]["rows"] + META["totals"]["2025"]["rows"] + META["totals"]["2025_12"]["rows"],
    "target_total_2025": sum(v or 0 for v in TARGET.values()),
}
# trùng khóa Product: GM% BK03 nếu lấy bản cũ (23.000) thay vì 25.000
META["bk03_cost"] = {"old": 23000, "current": 25000}
bk03 = [r for r in ALL if r[6] == "BK03"]
META["bk03_gm_effect"] = {"qty": sum(r[7] for r in bk03), "cogs_diff": sum(r[7] for r in bk03) * (25000 - 23000)}
json.dump(META, open(f"{GEN}/build_capstone_meta.json", "w", encoding="utf-8"), ensure_ascii=False, indent=1, default=str)

# ---------------------------------------------------------------- file giảng viên
vn = lambda x: f"{round(x):,}".replace(",", ".")
pc = lambda x: f"{x * 100:.1f}%".replace(".", ",")
wb = Workbook(); ws = wb.active; ws.title = "Doi_soat"
ws["A1"] = "ĐỐI SOÁT DATASET CAPSTONE — HẠT VIỆT COFFEE (chỉ giảng viên / trợ giảng)"; ws["A1"].font = Font(name="Arial", bold=True, size=13)
ws["A2"] = "Số sau khi làm sạch (bỏ dòng trùng, Trim Branch_Code, đổi SKU về chữ hoa). Revenue = Qty × Unit_Price × (1 − Discount_Pct/100), Unit_Price lấy từ Sales (không lấy từ lookup)."
rows = [["File", "Dòng thô", "Dòng sạch", "Dòng trùng", "Branch_Code dính dấu cách", "SKU viết thường", "Doanh thu sạch (VNĐ)", "Số đơn"]]
for f in files:
    nm = f["file"]; key = nm.replace("Sales_", "").replace(".csv", "")
    a = META["totals"]["2024"] if key == "2024" else META["months"][key]
    rows.append([nm, f["raw_rows"], f["clean_rows"], f["dup"], f["pad"], f["lower"], a["rev"], a["orders"]])
rows.append(["Tổng học viên nhận (2024 + 2025-01..11)", sum(r[1] for r in rows[1:-1]), META["checks"]["total_rows_clean_students"], "", "", "", META["checks"]["total_rev_students"], ""])
for i, row in enumerate(rows, 4):
    for j, v in enumerate(row, 1):
        c = ws.cell(row=i, column=j, value=v); c.font = NORM
        if i == 4: c.font = BOLD; c.fill = HFILL
        elif isinstance(v, (int, float)) and j >= 2: c.number_format = "#,##0"
ws.cell(row=len(rows) + 5, column=1, value="Kiểm tra Refresh (C2): thả New_Month/Sales_2025-12.csv vào folder Sales rồi Refresh. Sau Refresh:").font = Font(name="Arial", bold=True)
ws.cell(row=len(rows) + 6, column=1, value=f"Số dòng = {vn(META['checks']['total_rows_after_dec'])} (thêm {vn(META['checks']['dec_rows'])}); doanh thu tháng 12 = {vn(META['checks']['dec_rev'])}; số đơn tháng 12 = {vn(META['checks']['dec_orders'])}; tổng doanh thu 2025 (12 tháng) = {vn(META['totals']['2025']['rev'] + META['checks']['dec_rev'])}.")
ws.column_dimensions["A"].width = 46
for j in range(2, 9):
    ws.column_dimensions[get_column_letter(j)].width = 20

ws = wb.create_sheet("Loi_cai_san")
issues = [
    ["#", "Lỗi cài sẵn", "Ở đâu", "Cách sửa", "Tiêu chí"],
    [1, "Order_Time dạng chữ dd/mm/yyyy hh:mm; Unit_Price có dấu chấm ngăn nghìn kiểu Việt Nam (\"45.000\")", "Mọi file Sales", "Locale vi-VN (Using Locale) khi đổi kiểu cột", "C2"],
    [2, "Branch_Code dính dấu cách cuối (18 dòng mỗi file 2025)", "Sales_2025-01…12", "Transform → Format → Trim", "C2, C3"],
    [3, "SKU viết thường (5 dòng mỗi file 2025)", "Sales_2025-01…12", "Transform → Format → UPPERCASE", "C2, C3"],
    [4, "1 dòng trùng tuyệt đối mỗi file 2025", "Sales_2025-01…12", "Remove Duplicates (sau khi Trim/UPPERCASE)", "C2"],
    [5, "tblProduct trùng khóa BK03 (dòng cũ: giá vốn 23.000, Updated_At 01/11/2023; dòng hiện hành 25.000, 01/06/2024)", "HatViet_Lookups.xlsx · Product", "Giữ dòng Updated_At mới nhất (Sort giảm dần + Remove Duplicates) rồi nối 1-*", "C2, C3, C4"],
    [6, "Target: tiêu đề ở dòng 3, dòng TỔNG CỘNG cuối, tháng dạng cột, HV10 trống tháng 1–8", "HatViet_Target_2025.xlsx", "Bỏ dòng tổng, Promote Headers, Unpivot, bỏ null", "C2, C3"],
    [7, "Unit_Price trong Sales là giá thực tế tại thời điểm bán (Cà phê +5% từ 01/07/2025); tblProduct.Unit_Price là giá hiện hành", "Sales / Product", "Revenue tính từ Sales[Unit_Price], KHÔNG nhân giá lookup", "C4"],
    [8, "Customer_ID trống = khách không dùng thẻ thành viên (không phải lỗi)", "Sales", "Giữ nguyên; Active Customers chỉ đếm khách thành viên; dòng trống nối vào Customer thành (Blank)", "C3, C4"],
]
for i, row in enumerate(issues, 1):
    for j, v in enumerate(row, 1):
        c = ws.cell(row=i, column=j, value=v); c.font = BOLD if i == 1 else NORM
        if i == 1: c.fill = HFILL
        c.alignment = Alignment(wrap_text=True, vertical="top")
for col, w in zip("ABCDE", (5, 70, 30, 55, 12)):
    ws.column_dimensions[col].width = w

ws = wb.create_sheet("Cau_chuyen")
ws["A1"] = "BỐN CÂU CHUYỆN CÀI SẴN (không nói trước; chấm C6 theo việc học viên tự gọi tên và kiểm chứng)"; ws["A1"].font = Font(name="Arial", bold=True, size=13)
s1 = META["story1_delivery"]; s3 = META["story3_hv10"]; s4 = META["story4_coldbrew"]; s2 = META["story2_storm"]
stories = [
    ["#", "Câu chuyện", "Bằng chứng (số từ dataset)", "Công cụ phát hiện", "Hàm ý / đề xuất hợp lý"],
    [1, "Kênh giao hàng tăng nhanh, hoa hồng 22–25% kéo biên sau hoa hồng xuống",
     f"Tỷ trọng doanh thu kênh giao hàng: T1/2025 {pc(s1[1]['deliv_share_rev'])} → T11 {pc(s1[11]['deliv_share_rev'])}; GM% trước hoa hồng {pc(s1[1]['gm'])} → {pc(s1[11]['gm'])}, sau hoa hồng {pc(s1[1]['gm_net'])} → {pc(s1[11]['gm_net'])}. AOV tháng 11: " + "; ".join(f"{k} {vn(v)}" for k, v in META['story1_aov_nov'].items()),
     "Pivot Kênh × Tháng; measure Net GM % dùng tblChannel[Commission_Rate]", "Giá menu riêng cho app, giới hạn voucher, đàm phán hoa hồng; theo dõi biên sau hoa hồng, không chỉ doanh thu"],
    [2, "Bão Miền Trung 07–20/10/2025: HV04, HV05, HV06 đóng cửa 07–09/10, hạn chế đến 20/10",
     "Đạt target tháng 10 (%): " + ", ".join(f"{c} {pc(s2[f'{c}-10']['pct'])}" for c in ("HV04", "HV05", "HV06")) + f"; các chi nhánh còn lại {pc(s2['khac-10']['pct'])}. Tháng 11 hồi phục: " + ", ".join(f"{c} {pc(s2[f'{c}-11']['pct'])}" for c in ("HV04", "HV05", "HV06")),
     "Pivot Chi nhánh × Tháng, drill xuống ngày; Revenue vs Target % theo Region", "Không kết luận chi nhánh làm kém; ghi nhận sự kiện, cân nhắc điều chỉnh target và kế hoạch dự phòng mùa bão"],
    [3, "HV10 Cần Thơ khai trương 01/09/2025: YoY trống, tăng trưởng chuỗi bị thổi phồng bởi chi nhánh mới",
     "YoY toàn chuỗi vs like-for-like (không tính HV10) — " + "; ".join(f"T{m}: {pc(s3[m]['yoy_all'])} vs {pc(s3[m]['yoy_lfl'])}" for m in (9, 10, 11)) + ". Đạt target HV10: " + ", ".join(f"T{m} {pc(s3[m]['pct'])}" for m in (9, 10, 11)),
     "Đối soát số dòng Chi nhánh × Tháng; measure YoY và like-for-like", "Báo cáo hai số: toàn chuỗi và LFL; theo dõi HV10 theo đường tăng dần, không so với chi nhánh trưởng thành"],
    [4, "Cold Brew ra mắt 01/06/2025: lên tỷ trọng Cà phê, biên cao hơn, ăn một phần vào Cà phê sữa đá",
     f"Tỷ trọng Cold Brew trong doanh thu Cà phê: T6 {pc(s4[6]['cb_share_of_cf'])} → T11 {pc(s4[11]['cb_share_of_cf'])}; CF02 từ {pc(META['story4_cf02_share_may'])} (T5) xuống {pc(s4[11]['cf02_share_of_cf'])} (T11); GM% nhóm Cà phê {pc(s4[5 if 5 in s4 else 6]['cf_gm'])} (T6) → {pc(s4[11]['cf_gm'])} (T11)",
     "Pivot Sản phẩm × Tháng, % of Column Total; measure GM% theo nhóm", "Đẩy mạnh Cold Brew (biên cao); cân nhắc cannibalization khi định giá và khuyến mãi"],
]
for i, row in enumerate(stories, 3):
    for j, v in enumerate(row, 1):
        c = ws.cell(row=i, column=j, value=v); c.font = BOLD if i == 3 else NORM
        if i == 3: c.fill = HFILL
        c.alignment = Alignment(wrap_text=True, vertical="top")
for col, w in zip("ABCDE", (5, 48, 90, 38, 55)):
    ws.column_dimensions[col].width = w
ws["A9"] = "Hiệu ứng nền (không phải câu chuyện chính): Tết 10/02/2024 và 29/01/2025 (target lập theo cùng tháng 2024 nên lệch mùa), Cà phê tăng giá 5% từ 01/07/2025, cuối tuần cao hơn."
ws["A9"].font = Font(name="Arial", italic=True, size=9)

ws = wb.create_sheet("KPI_Tham_chieu")
ws["A1"] = "KPI THAM CHIẾU THEO THÁNG (đã làm sạch)"; ws["A1"].font = Font(name="Arial", bold=True, size=13)
hdr = ["Tháng", "Doanh thu", "Số đơn", "AOV", "COGS", "GM %", "Hoa hồng", "GM % sau hoa hồng", "Target", "Revenue vs Target %"]
for j, h in enumerate(hdr, 1):
    c = ws.cell(row=3, column=j, value=h); c.font = BOLD; c.fill = HFILL
for i, m in enumerate(range(1, 13), 4):
    a = META["months"][f"2025-{m:02d}"]; t = sum(TARGET[(c, m)] or 0 for c in BCODES)
    vals = [f"2025-{m:02d}", a["rev"], a["orders"], round(a["aov"]), a["cogs"], a["gm"], a["comm"], a["gm_net"], t, a["rev"] / t]
    for j, v in enumerate(vals, 1):
        c = ws.cell(row=i, column=j, value=v); c.font = NORM
        if j in (2, 3, 4, 5, 7, 9): c.number_format = "#,##0"
        if j in (6, 8, 10): c.number_format = "0.0%"
ws.cell(row=17, column=1, value="Tháng 12/2025 là file kiểm tra Refresh (New_Month); học viên không có dữ liệu này khi làm bài.").font = Font(name="Arial", italic=True, size=9)
for j in range(1, 11):
    ws.column_dimensions[get_column_letter(j)].width = 18
wb.move_sheet("Doi_soat", offset=0)
wb.save(f"{OUT_INS}/Dap_an_Capstone_HatViet.xlsx")
print("đã dựng dataset capstone:", len(ALL), "dòng sạch (2024–2025)")
for f in files:
    print(f)
