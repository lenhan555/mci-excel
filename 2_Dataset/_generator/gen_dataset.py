"""Generate the MCI Mart case dataset (synthetic) for the Excel for BI course.
Deterministic (seed=2025). Outputs go to MCI Excel/Materials/.
"""
import json, os, random, datetime as dt
import numpy as np
import pandas as pd

SEED = 2025
rng = np.random.default_rng(SEED)
random.seed(SEED)

ROOT = os.path.normpath(os.environ.get("MCI_ROOT") or os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
DS = os.path.join(ROOT, "2_Dataset")
os.makedirs(DS, exist_ok=True)

# ---------------------------------------------------------------- Store
stores = pd.DataFrame([
    ("ST01", "MCI Mart Cầu Giấy",   "Hà Nội",     "Bắc",   "Mall",   "2021-03-15", 1.30),
    ("ST02", "MCI Mart Hoàn Kiếm",  "Hà Nội",     "Bắc",   "Street", "2020-09-01", 1.10),
    ("ST03", "MCI Mart Hải Phòng",  "Hải Phòng",  "Bắc",   "Street", "2022-01-10", 0.80),
    ("ST04", "MCI Mart Hạ Long",    "Quảng Ninh", "Bắc",   "Street", "2022-07-01", 0.60),
    ("ST05", "MCI Mart Hải Châu",   "Đà Nẵng",    "Trung", "Mall",   "2021-05-20", 1.00),
    ("ST06", "MCI Mart Huế",        "Huế",        "Trung", "Street", "2022-11-11", 0.60),
    ("ST07", "MCI Mart Nha Trang",  "Khánh Hòa",  "Trung", "Street", "2021-12-01", 0.70),
    ("ST08", "MCI Mart Quận 1",     "TP.HCM",     "Nam",   "Mall",   "2020-06-01", 1.50),
    ("ST09", "MCI Mart Thủ Đức",    "TP.HCM",     "Nam",   "Mall",   "2021-08-08", 1.20),
    ("ST10", "MCI Mart Quận 7",     "TP.HCM",     "Nam",   "Street", "2021-02-01", 1.00),
    ("ST11", "MCI Mart Cần Thơ",    "Cần Thơ",    "Nam",   "Street", "2023-03-01", 0.70),
    ("ST12", "MCI Mart Biên Hòa",   "Đồng Nai",   "Nam",   "Street", "2024-06-01", 0.80),
], columns=["Store_Code", "Store_Name", "City", "Region", "Store_Type", "Open_Date", "_factor"])
stores["Open_Date"] = pd.to_datetime(stores["Open_Date"])

# ---------------------------------------------------------------- Product
P = [
 # code, name, category, price_2025, cost_ratio
 ("P001","Nước suối 500ml (lốc 6 chai)","Đồ uống",32000,0.72),
 ("P002","Nước ngọt có ga 330ml (lốc 6 lon)","Đồ uống",58000,0.74),
 ("P003","Trà xanh đóng chai 455ml","Đồ uống",10000,0.70),
 ("P004","Cà phê hòa tan 3in1 (hộp 20 gói)","Đồ uống",62000,0.68),
 ("P005","Cà phê rang xay 500g","Đồ uống",135000,0.62),
 ("P006","Nước ép trái cây 1L","Đồ uống",45000,0.70),
 ("P007","Bia lon 330ml (thùng 24 lon)","Đồ uống",345000,0.82),
 ("P008","Sữa tươi tiệt trùng 1L","Sữa",36000,0.78),
 ("P009","Sữa tươi 180ml (lốc 4 hộp)","Sữa",32000,0.78),
 ("P010","Sữa chua ăn (lốc 4 hộp)","Sữa",28000,0.72),
 ("P011","Sữa bột trẻ em 800g","Sữa",485000,0.80),
 ("P012","Sữa đặc có đường 380g","Sữa",26000,0.76),
 ("P013","Mì ăn liền tôm chua cay (thùng 30 gói)","Thực phẩm khô",128000,0.80),
 ("P014","Gạo thơm 5kg","Thực phẩm khô",145000,0.82),
 ("P015","Dầu ăn thực vật 1L","Thực phẩm khô",52000,0.80),
 ("P016","Nước mắm 500ml","Thực phẩm khô",38000,0.70),
 ("P017","Đường trắng 1kg","Thực phẩm khô",27000,0.84),
 ("P018","Hạt nêm 400g","Thực phẩm khô",42000,0.72),
 ("P019","Bún khô 500g","Thực phẩm khô",24000,0.70),
 ("P020","Cá hộp sốt cà 155g","Thực phẩm khô",22000,0.72),
 ("P021","Bánh quy bơ hộp thiếc 454g","Bánh kẹo",115000,0.60),
 ("P022","Kẹo trái cây hộp 250g","Bánh kẹo",48000,0.58),
 ("P023","Hạt điều rang muối 500g","Bánh kẹo",185000,0.65),
 ("P024","Mứt Tết thập cẩm 500g","Bánh kẹo",95000,0.55),
 ("P025","Socola thanh 100g","Bánh kẹo",42000,0.62),
 ("P026","Bánh snack khoai tây 90g","Bánh kẹo",19000,0.60),
 ("P027","Hộp quà Tết cao cấp","Bánh kẹo",650000,0.60),
 ("P028","Dầu gội 650ml","Chăm sóc cá nhân",165000,0.62),
 ("P029","Sữa tắm 800ml","Chăm sóc cá nhân",145000,0.62),
 ("P030","Kem đánh răng 230g","Chăm sóc cá nhân",39000,0.64),
 ("P031","Bàn chải đánh răng (vỉ 2)","Chăm sóc cá nhân",35000,0.55),
 ("P032","Khăn giấy rút (lốc 4 gói)","Chăm sóc cá nhân",56000,0.70),
 ("P033","Nước rửa tay 500ml","Chăm sóc cá nhân",62000,0.60),
 ("P034","Bột giặt 3kg","Gia dụng",158000,0.75),
 ("P035","Nước rửa chén 750ml","Gia dụng",32000,0.70),
 ("P036","Nước lau sàn 1L","Gia dụng",45000,0.68),
 ("P037","Túi đựng rác (cuộn 3)","Gia dụng",25000,0.55),
 ("P038","Hộp đựng thực phẩm (bộ 3)","Gia dụng",89000,0.50),
 ("P039","Chảo chống dính 26cm","Gia dụng",245000,0.55),
 ("P040","Nước xả vải 1.6L","Gia dụng",99000,0.72),
]
products = pd.DataFrame(P, columns=["Product_Code", "Product_Name", "Category", "Unit_Price", "_cost_ratio"])
products["Unit_Cost"] = (products["Unit_Price"] * products["_cost_ratio"] / 500).round() * 500
products["Unit_Cost"] = products["Unit_Cost"].astype(int)

# base popularity per product (weights), quantity profile
pop = rng.gamma(2.0, 1.0, len(products))
cheap = products["Unit_Price"] < 60000
pop[cheap.values] *= 1.8
products["_pop"] = pop
TET_BOOST = {"Bánh kẹo": 3.0, "Đồ uống": 1.6, "Thực phẩm khô": 1.3}

# ---------------------------------------------------------------- Customer (synthetic names)
HO = ["Nguyễn","Trần","Lê","Phạm","Hoàng","Huỳnh","Phan","Vũ","Võ","Đặng","Bùi","Đỗ","Hồ","Ngô","Dương","Lý"]
DEM_M = ["Văn","Hữu","Đức","Minh","Quốc","Thành","Gia","Hoàng"]
DEM_F = ["Thị","Ngọc","Thu","Thanh","Minh","Bảo","Khánh","Mỹ"]
TEN_M = ["An","Bình","Cường","Dũng","Hải","Hiếu","Hùng","Khoa","Long","Nam","Phong","Quân","Sơn","Tâm","Thắng","Tuấn","Việt","Vinh"]
TEN_F = ["Anh","Chi","Dung","Giang","Hà","Hạnh","Hương","Lan","Linh","Mai","Nga","Ngân","Nhung","Phương","Quỳnh","Thảo","Trang","Vy"]
N_CUST = 1500
cust_rows = []
store_w = stores["_factor"].values / stores["_factor"].sum()
for i in range(1, N_CUST + 1):
    g = "Nữ" if rng.random() < 0.58 else "Nam"
    name = f"{random.choice(HO)} {random.choice(DEM_F if g=='Nữ' else DEM_M)} {random.choice(TEN_F if g=='Nữ' else TEN_M)}"
    home = rng.choice(stores["Store_Code"].values, p=store_w)
    open_d = stores.set_index("Store_Code").loc[home, "Open_Date"]
    start = max(open_d, pd.Timestamp("2020-06-01"))
    span = (pd.Timestamp("2025-03-31") - start).days
    join = start + pd.Timedelta(days=int(rng.integers(0, max(span, 1))))
    seg = rng.choice(["Thường", "Thân thiết", "VIP"], p=[0.70, 0.25, 0.05])
    cust_rows.append((f"C{i:05d}", name, g, int(rng.integers(1960, 2006)), seg, join.normalize(), home))
customers = pd.DataFrame(cust_rows, columns=["Customer_ID","Full_Name","Gender","Birth_Year","Segment","Join_Date","Home_Store_Code"])
seg_w = customers["Segment"].map({"Thường": 1.0, "Thân thiết": 2.5, "VIP": 5.0}).values

# ---------------------------------------------------------------- Sales
START, END = pd.Timestamp("2024-01-01"), pd.Timestamp("2025-04-30")
TET = {2024: pd.Timestamp("2024-02-10"), 2025: pd.Timestamp("2025-01-29")}
MONTH_SEAS = {1:1.00,2:0.95,3:0.92,4:0.95,5:0.97,6:1.00,7:1.02,8:1.00,9:1.03,10:1.02,11:1.05,12:1.15}

def day_mult(d):
    m = MONTH_SEAS[d.month]
    tet = TET[d.year] if d.year in TET else None
    tet_phase = None
    if tet is not None:
        delta = (tet - d).days
        if 1 <= delta <= 20:
            m *= 1.6; tet_phase = "pre"
        elif -4 <= delta <= 0:
            m *= 0.35; tet_phase = "holiday"
    if d.weekday() >= 5:
        m *= 1.25
    if d.year == 2025:
        m *= 1.08          # underlying growth
    return m, tet_phase

prod_codes = products["Product_Code"].values
price_2025 = products.set_index("Product_Code")["Unit_Price"]
cat_of = products.set_index("Product_Code")["Category"]
base_pop = products["_pop"].values
tet_pop = base_pop * products["Category"].map(lambda c: TET_BOOST.get(c, 1.0)).values
cust_by_store = {s: np.where(customers["Home_Store_Code"].values == s)[0] for s in stores["Store_Code"]}

rows = []
order_seq = {}
for d in pd.date_range(START, END, freq="D"):
    mult, phase = day_mult(d)
    for _, st in stores.iterrows():
        if d < st["Open_Date"]:
            continue
        f = st["_factor"]
        # anomaly for storytelling: ST07 renovation 2025-03-10 .. 2025-03-31 (reduced area)
        if st["Store_Code"] == "ST07" and pd.Timestamp("2025-03-10") <= d <= pd.Timestamp("2025-03-31"):
            f *= 0.35
        # ramp-up for new store ST12 during first 3 months
        if st["Store_Code"] == "ST12":
            months_open = (d - st["Open_Date"]).days / 30
            f *= min(1.0, 0.5 + 0.17 * months_open)
        lam = 2.6 * f * mult
        n_orders = rng.poisson(lam)
        for _ in range(n_orders):
            key = d.strftime("%y%m")
            order_seq[key] = order_seq.get(key, 0) + 1
            oid = f"HD{key}-{order_seq[key]:05d}"
            # customer: 85% home-store customers, 15% others
            pool = cust_by_store[st["Store_Code"]]
            if len(pool) == 0 or rng.random() < 0.15:
                pool = np.arange(N_CUST)
            w = seg_w[pool] / seg_w[pool].sum()
            cust = customers["Customer_ID"].values[rng.choice(pool, p=w)]
            n_lines = rng.choice([1, 2, 3, 4], p=[0.40, 0.33, 0.18, 0.09])
            pw = tet_pop if phase == "pre" else base_pop
            chosen = rng.choice(len(prod_codes), size=n_lines, replace=False, p=pw / pw.sum())
            for ln, pi in enumerate(chosen, start=1):
                pc = prod_codes[pi]
                price = int(price_2025[pc])
                if d.year == 2024:          # 2025 list prices rose ~4% vs 2024
                    price = int(round(price / 1.04 / 500) * 500)
                qty = int(rng.choice([1, 2, 3, 4, 6], p=[0.55, 0.25, 0.10, 0.06, 0.04]))
                if price >= 300000:
                    qty = 1
                rows.append((oid, ln, d, st["Store_Code"], cust, pc, qty, price))

sales = pd.DataFrame(rows, columns=["Order_ID","Line_No","Order_Date","Store_Code","Customer_ID","Product_Code","Quantity","Unit_Price"])
sales = sales.sort_values(["Order_Date","Order_ID","Line_No"]).reset_index(drop=True)
sales["Revenue"] = sales["Quantity"] * sales["Unit_Price"]

# ---------------------------------------------------------------- Target 2025 (wide, month columns)
s24 = sales[sales["Order_Date"].dt.year == 2024].copy()
s24["m"] = s24["Order_Date"].dt.month
act24 = s24.groupby(["Store_Code","m"])["Revenue"].sum().unstack(fill_value=0)
tgt = pd.DataFrame(index=stores["Store_Code"])
for m in range(1, 13):
    col = f"2025-{m:02d}"
    vals = []
    for s in stores["Store_Code"]:
        a = act24.loc[s, m] if (s in act24.index and m in act24.columns) else 0
        if s == "ST12" and m < 9:        # ST12 had partial/ramp-up 2024 months -> plan from its Sep-Dec 2024 average
            a = act24.loc["ST12", [9, 10, 11, 12]].mean()
        vals.append(round(a * 1.10 / 1e6) * 1e6)   # +10% vs same month 2024, rounded to million
    tgt[col] = np.array(vals, dtype=np.int64)
tgt = tgt.reset_index().merge(stores[["Store_Code","Store_Name"]], on="Store_Code")
tgt = tgt[["Store_Code","Store_Name"] + [f"2025-{m:02d}" for m in range(1, 13)]]

# ---------------------------------------------------------------- save pickles for downstream builders
os.makedirs(os.path.join(os.path.dirname(__file__), "build"), exist_ok=True)
B = os.path.join(os.path.dirname(__file__), "build")
stores.drop(columns=["_factor"]).to_pickle(f"{B}/stores.pkl")
products.drop(columns=["_cost_ratio","_pop"]).to_pickle(f"{B}/products.pkl")
customers.to_pickle(f"{B}/customers.pkl")
sales.to_pickle(f"{B}/sales.pkl")
tgt.to_pickle(f"{B}/target.pkl")

# quick stats
q1 = sales[(sales.Order_Date >= "2025-01-01") & (sales.Order_Date <= "2025-03-31")]
stats = {
    "sales_rows_total": len(sales),
    "orders_total": sales.Order_ID.nunique(),
    "rows_2024": int((sales.Order_Date.dt.year == 2024).sum()),
    "rows_q1_2025": len(q1),
    "orders_q1_2025": q1.Order_ID.nunique(),
    "revenue_q1_2025": int(q1.Revenue.sum()),
    "revenue_by_month": {k.strftime("%Y-%m"): int(v) for k, v in sales.groupby(sales.Order_Date.dt.to_period("M").dt.to_timestamp())["Revenue"].sum().items()},
    "aov_q1_2025": round(q1.Revenue.sum() / q1.Order_ID.nunique()),
}
json.dump(stats, open(f"{B}/stats.json", "w"), ensure_ascii=False, indent=1)
print(json.dumps(stats, ensure_ascii=False, indent=1))
