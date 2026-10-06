"""Buổi 6 (Dashboard): số liệu dùng trong buổi, ghi _generator/build_b6_meta.json.

Chạy được trên Mac lẫn Windows (openpyxl, pandas, numpy; statsmodels chỉ để ước lượng ETS — không có thì bỏ qua phần đó).
Không dựng file nào trong thư mục buổi.

Nguồn = chính file khởi đầu của B5 là đáp án B4 + file lookup / target B4 (giống build_b5_meta.py), nên số của B6 khớp
số của B5 (đối chiếu với build_b5_meta.json: lệch thì dừng). Mỗi con số trong kịch bản B6 được tính lại đúng như measure
của B5 sẽ tính (tỷ lệ = tổng ÷ tổng; BLANK khi không có dữ liệu).

SỐ ĐO TRÊN EXCEL (Mac 365, 03/10/2026, hàm FORECAST.ETS, seasonality 7): tháng 5/2025 tổng 237.672.111; dải 95% nửa rộng TB 18.036.080;
kiểm chứng ngược tháng 4 (dữ liệu tới 31/03) 280.041.676 vs 296.965.000; FORECAST.ETS.SEASONALITY tự phát hiện = 121 (cả chuỗi) / 35 (tới 31/03).
Các số ETS của statsmodels bên dưới KHÔNG dùng trong kịch bản (lệch nhiều so với Excel); chỉ Trendline/thống kê theo thứ là chính xác.

PHẦN DỰ BÁO (khóa "forecast"): Linear Trendline và thống kê theo thứ là chính xác (tính bằng OLS; Excel Trendline đặt
x = 1..n). ETS ước lượng bằng statsmodels (Holt-Winters cộng, mùa 7): **Excel Forecast Sheet dùng thuật toán AAA riêng,
nên số ETS ở đây chỉ để định cỡ (đúng cỡ chục %), phải chụp lại trên Excel 365 Windows trước khi đưa vào slide.**

    /Users/leopham/anaconda3/bin/python _generator/build_b6_meta.py     # Mac;  Windows: python _generator\\build_b6_meta.py
"""
import json, os, sys, warnings
import numpy as np
import pandas as pd
from openpyxl import load_workbook

warnings.filterwarnings("ignore")
for s_ in (sys.stdout, sys.stderr):
    if hasattr(s_, "reconfigure"): s_.reconfigure(encoding="utf-8", errors="replace")

ROOT = os.path.normpath(os.environ.get("MCI_ROOT") or os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
GEN = os.path.dirname(os.path.abspath(__file__))
B4P = os.path.join(ROOT, "1_Sessions", "B4_Power_Query", "Practice")
STARTER = os.path.join(B4P, "B4_MCI_Mart_Answer.xlsx")
MASTER = os.path.join(ROOT, "2_Dataset", "MCI_Mart_Master.xlsx")
B5META = json.load(open(os.path.join(GEN, "build_b5_meta.json"), encoding="utf-8"))

# ---------------------------------------------------------------- 6 bảng (đúng như query B4)
wb = load_workbook(STARTER, read_only=True, data_only=True)
rows = list(wb["Sales_PQ"].iter_rows(values_only=True))
sales = pd.DataFrame(rows[1:], columns=rows[0])
sales["Order_Date"] = pd.to_datetime(sales.Order_Date)
wb.close()
lk = pd.ExcelFile(os.path.join(B4P, "B4_Lookups.xlsx"))
product = lk.parse("Product").drop_duplicates("Product_Code")
store = lk.parse("Store")
customer = lk.parse("Customer")
tw = pd.read_excel(os.path.join(B4P, "B4_Target_2025.xlsx"), sheet_name="Target_2025", header=2)
target = tw.melt(id_vars=["Store_Code", "Store_Name"], var_name="Attribute", value_name="Target")
target["Thang"] = target.Attribute.str.split(" ").str[-1].astype(int)
target = target[["Store_Code", "Thang", "Target"]]

master = pd.read_excel(MASTER, sheet_name="Sales")
assert len(sales) == len(master) == 30247 and int(sales.Revenue.sum()) == 4436849500

s = sales.merge(product[["Product_Code", "Product_Name", "Category", "Unit_Cost", "Unit_Price"]], on="Product_Code", how="left") \
         .merge(store[["Store_Code", "Store_Name", "City", "Region", "Store_Type", "Open_Date"]], on="Store_Code", how="left") \
         .merge(customer[["Customer_ID", "Segment"]], on="Customer_ID", how="left")
s["COGS"] = s.Quantity * s.Unit_Cost
s["Y"], s["M"], s["Q"] = s.Order_Date.dt.year, s.Order_Date.dt.month, s.Order_Date.dt.quarter
t = target.merge(store[["Store_Code", "Region"]], on="Store_Code")
d25 = s[s.Y == 2025]
d24 = s[s.Y == 2024]
REG = ["Bắc", "Trung", "Nam"]
STORES = list(store.Store_Code)
CATS = list(product.Category.drop_duplicates())

vn = lambda x: f"{x:,.0f}".replace(",", ".")
pct = lambda x, d=1: "(trống)" if x is None else f"{x * 100:.{d}f}%".replace(".", ",")


def kpi(d):
    if len(d) == 0:
        return {"rev": None, "orders": None, "aov": None, "cogs": None, "gm": None, "cust": None}
    rev, orders, cogs = int(d.Revenue.sum()), int(d.Order_ID.nunique()), int(d.COGS.sum())
    return {"rev": rev, "orders": orders, "aov": rev / orders, "cogs": cogs, "gm": (rev - cogs) / rev, "cust": int(d.Customer_ID.nunique())}


def div(a, b):
    return None if a is None or not b else a / b


# ---------------------------------------------------------------- khớp B5 (cùng nguồn): lệch thì dừng
for m in range(1, 5):
    k = kpi(d25[d25.M == m])
    b5 = B5META["kpi_month_2025"][str(m)] if str(m) in B5META["kpi_month_2025"] else B5META["kpi_month_2025"][m - 1]
    assert abs(k["rev"] - (b5["rev"] if "rev" in b5 else b5["Total Revenue"])) < 1, (m, k, b5)
assert B5META["ytd_2025"]["4"] == int(d25[d25.M <= 4].Revenue.sum())

# ---------------------------------------------------------------- KPI card: giá trị, target, mức đạt, RAG
# Target KPI theo Course_Plan Phụ lục C: Revenue = bảng Target; Orders 1.000/tháng; AOV 300.000; GM% 33%; YoY +10%; Active Customers 700/tháng.
TARGETS = {"orders": 1000, "aov": 300000, "gm": 0.33, "yoy": 0.10, "cust": 700}
RAG_OK, RAG_WARN = 1.0, 0.9        # Theo KPI Dictionary (Target 100%, Warning 90%): ✔ ≥ 100% · ▲ 90% – <100% · ✘ < 90% (Critical 80% chỉ nằm ở KPI Dictionary, không có màu riêng)


def rag(att):
    return "ok" if att >= RAG_OK - 1e-12 else "warn" if att >= RAG_WARN - 1e-12 else "bad"


rev24 = {m: int(d24[d24.M == m].Revenue.sum()) for m in range(1, 13)}
tgt_month = {int(m): int(v) for m, v in target.groupby("Thang").Target.sum().items()}
card = {}
for m in range(1, 5):
    k = kpi(d25[d25.M == m])
    yoy = (k["rev"] - rev24[m]) / rev24[m]
    c = {"rev": k["rev"], "target": tgt_month[m], "orders": k["orders"], "aov": k["aov"], "gm": k["gm"], "yoy": yoy, "cust": k["cust"]}
    att = {"rev": k["rev"] / tgt_month[m], "orders": k["orders"] / TARGETS["orders"], "aov": k["aov"] / TARGETS["aov"],
           "gm": k["gm"] / TARGETS["gm"], "yoy": yoy / TARGETS["yoy"], "cust": k["cust"] / TARGETS["cust"]}
    c["att"] = att
    c["rag"] = {kk: rag(v) for kk, v in att.items()}
    c["variance"] = k["rev"] - tgt_month[m]
    card[m] = c

# ---------------------------------------------------------------- combo Actual vs Target theo tháng (2025) và theo cửa hàng
actual_month = {m: int(d25[d25.M == m].Revenue.sum()) for m in range(1, 5)}
pct_month = {m: actual_month[m] / tgt_month[m] for m in range(1, 5)}
ytd_actual = sum(actual_month.values())
ytd_target = sum(tgt_month[m] for m in range(1, 5))
store_month = {}
for sc in STORES:
    store_month[sc] = {}
    for m in range(1, 5):
        a = int(d25[(d25.Store_Code == sc) & (d25.M == m)].Revenue.sum())
        tg = int(target[(target.Store_Code == sc) & (target.Thang == m)].Target.iloc[0])
        store_month[sc][m] = {"rev": a, "target": tg, "var": a - tg, "pct": a / tg}
store_ytd = {sc: {"rev": sum(store_month[sc][m]["rev"] for m in range(1, 5)),
                  "target": sum(store_month[sc][m]["target"] for m in range(1, 5))} for sc in STORES}
for sc in STORES:
    store_ytd[sc]["var"] = store_ytd[sc]["rev"] - store_ytd[sc]["target"]
    store_ytd[sc]["pct"] = store_ytd[sc]["rev"] / store_ytd[sc]["target"]
region_month = {r: {m: {"rev": int(d25[(d25.Region == r) & (d25.M == m)].Revenue.sum()),
                        "target": int(t[(t.Region == r) & (t.Thang == m)].Target.sum())} for m in range(1, 5)} for r in REG}
below80 = {m: [sc for sc in STORES if store_month[sc][m]["pct"] < RAG_WARN - 1e-12] for m in range(1, 5)}   # tên biến cũ; nay là cửa hàng đỏ (< 90%)
below80_strict = {m: [sc for sc in STORES if store_month[sc][m]["pct"] < 0.8] for m in range(1, 5)}   # mức Critical 80% của KPI Dictionary

# ---------------------------------------------------------------- Top sản phẩm, thành phố, ngành (tháng 4 và T1–T4)
def top_products(d, n=5):
    g = d.groupby(["Product_Code", "Product_Name", "Category"]).Revenue.sum().sort_values(ascending=False).head(n)
    return [{"code": i[0], "name": i[1], "cat": i[2], "rev": int(v)} for i, v in g.items()]


top_t4 = top_products(d25[d25.M == 4])
top_t3 = top_products(d25[d25.M == 3])
top_t1t4 = top_products(d25[d25.M <= 4])
top5_share_t4 = sum(x["rev"] for x in top_t4) / actual_month[4]
city_t4 = {c: int(v) for c, v in d25[d25.M == 4].groupby("City").Revenue.sum().sort_values(ascending=False).items()}
region_t4 = {r: int(d25[(d25.M == 4) & (d25.Region == r)].Revenue.sum()) for r in REG}
cat_ytd = {}
for c in CATS:
    d = d25[(d25.M <= 4) & (d25.Category == c)]
    cat_ytd[c] = {"rev": int(d.Revenue.sum()), "gm": kpi(d)["gm"], "share": d.Revenue.sum() / ytd_actual}
cat_ytd = dict(sorted(cat_ytd.items(), key=lambda kv: -kv[1]["rev"]))

# ---------------------------------------------------------------- Dashboard "xấu": trục Y cắt khúc, pie 6 phần
t3_24, t4_24 = rev24[3], rev24[4]
axis_start = 236_000_000
trunc = {"t3_2024": t3_24, "t4_2024": t4_24, "chenh_pct": (t4_24 - t3_24) / t3_24, "axis_start": axis_start,
         "ty_le_hinh_hoc": (t4_24 - axis_start) / (t3_24 - axis_start)}
pie_shares_ytd = {c: v["share"] for c, v in cat_ytd.items()}

# ---------------------------------------------------------------- Phân phối giá trị đơn (Q1/2025, = B2): histogram, outlier
q1 = d25[d25.Q == 1]
ov = q1.groupby("Order_ID").Revenue.sum()
q1_, q3_ = float(ov.quantile(0.25)), float(ov.quantile(0.75))
fence = q3_ + 1.5 * (q3_ - q1_)
bins = list(range(0, 1_000_001, 100_000))
hist = np.histogram(ov, bins=bins + [float("inf")])[0].tolist()
dist = {"n_orders": int(len(ov)), "mean": float(ov.mean()), "median": float(ov.median()), "q1": q1_, "q3": q3_, "fence": fence,
        "n_outlier": int((ov > fence).sum()), "max": float(ov.max()), "skew": float(ov.skew()),
        "hist_bin_100k": hist, "hist_labels": [f"{b // 1000}-{(b + 100_000) // 1000}k" for b in bins[:-1]] + ["≥1.000k"]}
assert dist["n_orders"] == 3144 and dist["n_outlier"] == 138 and round(fence) == 896500, dist

# ---------------------------------------------------------------- Scatter: tuổi cửa hàng ↔ doanh thu T1–T4/2025 (tương quan ≠ nhân quả)
st = d25[d25.M <= 4].groupby("Store_Code").agg(rev=("Revenue", "sum")).join(store.set_index("Store_Code")[["Open_Date", "Store_Type", "City", "Region"]])
st["age_years"] = (pd.Timestamp("2025-04-30") - st.Open_Date).dt.days / 365.25
r_all = float(np.corrcoef(st.age_years, st.rev)[0, 1])
r_no12 = float(np.corrcoef(st.drop("ST12").age_years, st.drop("ST12").rev)[0, 1])
big_city = st.City.isin(["Hà Nội", "TP.HCM"])
scatter = {"r_all": r_all, "r2_all": r_all ** 2, "r_without_ST12": r_no12,
           "avg_rev_big_city": float(st[big_city].rev.mean()), "avg_rev_other": float(st[~big_city].rev.mean()),
           "avg_age_big_city": float(st[big_city].age_years.mean()), "avg_age_other": float(st[~big_city].age_years.mean()),
           "n": int(len(st)), "points": {sc: {"age": round(float(r.age_years), 2), "rev": int(r.rev)} for sc, r in st.iterrows()}}

# ---------------------------------------------------------------- Dự báo: bảng Ngày–Doanh thu (486 ngày)
daily = s.groupby("Order_Date").Revenue.sum().astype(float)
assert len(daily) == 486 and daily.index.min() == pd.Timestamp("2024-01-01") and daily.index.max() == pd.Timestamp("2025-04-30")
assert len(pd.date_range(daily.index.min(), daily.index.max())) == len(daily)          # không thiếu ngày
dow_name = ["Thứ Hai", "Thứ Ba", "Thứ Tư", "Thứ Năm", "Thứ Sáu", "Thứ Bảy", "Chủ nhật"]
dow_mean = {dow_name[i]: float(daily[daily.index.dayofweek == i].mean()) for i in range(7)}
wk = daily[daily.index.dayofweek >= 5].mean() / daily[daily.index.dayofweek < 5].mean()
sun_vs_mon = dow_mean["Chủ nhật"] / dow_mean["Thứ Hai"]


def ols(y):
    x = np.arange(1, len(y) + 1, dtype=float)
    b, a = np.polyfit(x, y, 1)
    pred = a + b * x
    r2 = 1 - ((y - pred) ** 2).sum() / ((y - y.mean()) ** 2).sum()
    return float(a), float(b), float(r2)


a_all, b_all, r2_all = ols(daily.values)
n = len(daily)
lin_next30 = [a_all + b_all * (n + i) for i in range(1, 31)]
# kiểm chứng ngược (backtest): cắt dữ liệu tới 31/03/2025, dự báo 30 ngày tháng 4/2025 rồi so với thực tế
train = daily[:"2025-03-31"]
test = daily["2025-04-01":"2025-04-30"]
a_tr, b_tr, r2_tr = ols(train.values)
lin_bt = [a_tr + b_tr * (len(train) + i) for i in range(1, 31)]
forecast = {
    "n_days": int(n), "first": "2024-01-01", "last": "2025-04-30",
    "dow_mean": dow_mean, "weekend_vs_weekday": float(wk), "sunday_vs_monday": float(sun_vs_mon),
    "linear_all": {"intercept": a_all, "slope_per_day": b_all, "r2": r2_all, "next30_mean": float(np.mean(lin_next30)),
                   "next30_first": lin_next30[0], "next30_last": lin_next30[-1]},
    "daily_actual_april_mean": float(test.mean()), "daily_actual_april_min": float(test.min()), "daily_actual_april_max": float(test.max()),
    "daily_mean_all": float(daily.mean()), "daily_std_all": float(daily.std()),
    "backtest_linear": {"slope_per_day": b_tr, "r2": r2_tr, "pred_april_sum": float(sum(lin_bt)), "actual_april_sum": float(test.sum()),
                        "err_pct": float((sum(lin_bt) - test.sum()) / test.sum()),
                        "mae": float(np.mean(np.abs(np.array(lin_bt) - test.values))),
                        "mae_pct_of_mean": float(np.mean(np.abs(np.array(lin_bt) - test.values)) / test.mean())},
    "tet_days": {"2024": "2024-02-10", "2025": "2025-01-29", "2026": "2026-02-17"},
}
try:
    from statsmodels.tsa.holtwinters import ExponentialSmoothing
    fit = ExponentialSmoothing(train.values, trend="add", damped_trend=True, seasonal="add", seasonal_periods=7).fit()
    pred = fit.forecast(30)
    resid = train.values - fit.fittedvalues
    sigma = float(np.std(resid))
    forecast["backtest_ets_approx"] = {
        "pred_april_sum": float(pred.sum()), "actual_april_sum": float(test.sum()),
        "err_pct": float((pred.sum() - test.sum()) / test.sum()),
        "mae": float(np.mean(np.abs(pred - test.values))), "mae_pct_of_mean": float(np.mean(np.abs(pred - test.values)) / test.mean()),
        "sigma_resid": sigma, "band95_halfwidth_daily_approx": 1.96 * sigma,
        "band95_halfwidth_pct_of_mean": 1.96 * sigma / float(train.mean()),
        "coverage_95": float(np.mean(np.abs(pred - test.values) <= 1.96 * sigma)),
        "pred_by_dow_april": {dow_name[i]: float(np.mean([p for p, d in zip(pred, test.index) if d.dayofweek == i])) for i in range(7)},
        "note": "statsmodels Holt-Winters cộng, mùa 7; Excel Forecast Sheet (AAA) có thể lệch. Chụp lại trên Excel trước khi đưa vào slide.",
    }
    fit2 = ExponentialSmoothing(daily.values, trend="add", damped_trend=True, seasonal="add", seasonal_periods=7).fit()
    p2 = fit2.forecast(30)
    s2 = float(np.std(daily.values - fit2.fittedvalues))
    forecast["ets_next30_approx"] = {"first_day": "2025-05-01", "sum": float(p2.sum()), "mean_daily": float(p2.mean()),
                                     "band95_halfwidth_daily_approx": 1.96 * s2, "band95_halfwidth_pct_of_mean": 1.96 * s2 / float(p2.mean()),
                                     "sat_sun_vs_weekday": float(np.mean([p for p, d in zip(p2, pd.date_range('2025-05-01', periods=30)) if d.dayofweek >= 5]) /
                                                                 np.mean([p for p, d in zip(p2, pd.date_range('2025-05-01', periods=30)) if d.dayofweek < 5]))}
except Exception as e:                                                          # pragma: no cover
    forecast["ets_error"] = str(e)

# tự tương quan theo độ trễ (vì sao ETS đoán mùa 7): bỏ xu hướng bằng trung bình trượt 29 ngày (dài hơn 1 chu kỳ tuần, không xóa nhịp tuần)
ma = pd.Series(daily.values).rolling(29, center=True, min_periods=15).mean().values
x = daily.values - ma
acf = {lag: float(np.corrcoef(x[:-lag], x[lag:])[0, 1]) for lag in range(1, 15)}
forecast["acf_detrended"] = acf
forecast["acf_peak_lag"] = int(max(range(2, 15), key=lambda l: acf[l]))


def dow_r2(seg):
    g = seg.groupby(seg.index.dayofweek).transform("mean")
    return float(1 - ((seg - g) ** 2).sum() / ((seg - seg.mean()) ** 2).sum())


# nhịp tuần tồn tại (cuối tuần +19%) nhưng nhỏ so với biến động ngày: R² theo thứ chỉ vài % → Forecast Sheet có thể không tự phát hiện mùa 7
forecast["dow_r2"] = {"all": dow_r2(daily), "mar_dec_2024": dow_r2(daily["2024-03-01":"2024-12-31"]), "mar_apr_2025": dow_r2(daily["2025-03-01":"2025-04-30"])}
forecast["daily_std_mar_dec_2024"] = float(daily["2024-03-01":"2024-12-31"].std())

# ---------------------------------------------------------------- Insight SCR mẫu (YTD T1–T4)
scr = {
    "ytd_actual": ytd_actual, "ytd_target": ytd_target, "ytd_pct": ytd_actual / ytd_target, "ytd_var": ytd_actual - ytd_target,
    "t1_t2_pct": (actual_month[1] + actual_month[2]) / (tgt_month[1] + tgt_month[2]),
    "t1_t2_var": actual_month[1] + actual_month[2] - tgt_month[1] - tgt_month[2],
    "st07_t3": store_month["ST07"][3], "st07_t3_var": store_month["ST07"][3]["var"],
    "st07_t3_share_of_t3_gap": store_month["ST07"][3]["var"] / (actual_month[3] - tgt_month[3]),
    "st12_ytd": store_ytd["ST12"], "yoy_ytd": B5META["ytd_yoy_t4"], "yoy_lfl_ytd": B5META["lfl_ytd_yoy"],
}

# ---------------------------------------------------------------- Tình huống bài về nhà B6 (dashboard Marketing — Active Customers)
seg = {}
for sg in ["Thường", "Thân thiết", "VIP"]:
    d = d25[(d25.M == 4) & (d25.Segment == sg)]
    seg[sg] = {"cust": int(d.Customer_ID.nunique()), "rev": int(d.Revenue.sum())}
cust_t4_reg = {r: int(d25[(d25.M == 4) & (d25.Region == r)].Customer_ID.nunique()) for r in REG}
cust_total_t4 = int(d25[d25.M == 4].Customer_ID.nunique())
hw = {"cust_t1_t4": {m: card[m]["cust"] for m in range(1, 5)}, "cust_target": TARGETS["cust"],
      "cust_t4_by_segment": seg, "cust_t4_by_region": cust_t4_reg, "cust_t4_total": cust_total_t4,
      "cust_t4_sum_regions": sum(cust_t4_reg.values()), "cust_t4_sum_segments": sum(v["cust"] for v in seg.values())}

meta = {
    "source": "B4_MCI_Mart_Answer.xlsx + B4_Lookups.xlsx + B4_Target_2025.xlsx (cùng nguồn build_b5_meta.py)",
    "targets": TARGETS, "rag_thresholds": {"ok": RAG_OK, "warn": RAG_WARN},
    "card": card, "target_month": tgt_month, "actual_month": actual_month, "pct_month": pct_month,
    "ytd": {"actual": ytd_actual, "target": ytd_target, "pct": ytd_actual / ytd_target, "var": ytd_actual - ytd_target},
    "store_month": store_month, "store_ytd": store_ytd, "region_month": region_month, "stores_below_80": below80, "stores_below_90": below80, "stores_below_80_critical": below80_strict,
    "top5_t4": top_t4, "top5_t3": top_t3, "top5_t1t4": top_t1t4, "top5_share_t4": top5_share_t4,
    "city_t4": city_t4, "region_t4": region_t4, "cat_ytd": cat_ytd, "pie_shares_ytd": pie_shares_ytd,
    "truncated_axis": trunc, "dist": dist, "scatter": scatter, "forecast": forecast, "scr": scr, "hw": hw,
}
with open(os.path.join(GEN, "build_b6_meta.json"), "w", encoding="utf-8") as f:
    json.dump(meta, f, ensure_ascii=False, indent=1, default=lambda o: float(o) if isinstance(o, (np.floating,)) else int(o) if isinstance(o, np.integer) else str(o))

# ---------------------------------------------------------------- in tóm tắt
print("KPI card theo tháng 2025 (giá trị · target · mức đạt · RAG):")
for m in range(1, 5):
    c = card[m]
    print(f" T{m}: rev {vn(c['rev'])} / {vn(c['target'])} = {pct(c['att']['rev'])} {c['rag']['rev']} | orders {c['orders']} {pct(c['att']['orders'])} {c['rag']['orders']}"
          f" | AOV {vn(c['aov'])} {pct(c['att']['aov'])} {c['rag']['aov']} | GM {pct(c['gm'])} {c['rag']['gm']} | YoY {pct(c['yoy'])} {c['rag']['yoy']}"
          f" | cust {c['cust']} {pct(c['att']['cust'])} {c['rag']['cust']} | var {vn(c['variance'])}")
print("YTD:", vn(ytd_actual), vn(ytd_target), pct(ytd_actual / ytd_target), vn(ytd_actual - ytd_target))
print("Cửa hàng < 90% target (đỏ) theo tháng:", below80)
print("Cửa hàng < 80% target (Critical) theo tháng:", below80_strict)
print("Variance YTD theo cửa hàng:", {sc: vn(v["var"]) for sc, v in store_ytd.items()})
print("Top 5 T4:", [(x["name"], vn(x["rev"])) for x in top_t4], pct(top5_share_t4))
print("Thành phố T4:", {k: vn(v) for k, v in city_t4.items()})
print("Ngành YTD:", {k: (vn(v["rev"]), pct(v["gm"]), pct(v["share"])) for k, v in cat_ytd.items()})
print("Trục cắt khúc:", trunc)
print("Phân phối đơn Q1:", {k: v for k, v in dist.items() if k not in ("hist_bin_100k", "hist_labels")}, dist["hist_bin_100k"])
print("Scatter:", {k: v for k, v in scatter.items() if k != "points"})
print("Dự báo:", json.dumps({k: v for k, v in forecast.items() if k not in ("acf_detrended",)}, ensure_ascii=False, indent=1, default=str))
print("ACF lag 7:", forecast["acf_detrended"][7], "peak lag", forecast["acf_peak_lag"])
print("SCR:", scr)
print("HW:", hw)
