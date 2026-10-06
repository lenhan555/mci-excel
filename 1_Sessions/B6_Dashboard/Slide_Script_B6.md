# Kịch bản slide — Buổi 6: Thiết kế & Xây dựng Dashboard — "Giám đốc bán hàng có 5 giây"

> **Phiên bản 1.0 (2026-10-03) — bản nháp chờ duyệt.** Soạn theo `Course_Plan.md` v3.3 (B6: timetable 150 phút, Hướng dẫn giảng dạy B6, Phụ lục D, Phụ lục F), khuôn mẫu `Slide_Script_B5.md`. Lời giảng viết thành **văn nói đầy đủ**. Mọi con số tính từ chính file khởi đầu (đáp án B5, dữ liệu như `Practice/B5_MCI_Mart_Answer.xlsx`) bằng `2_Dataset/_generator/build_b6_meta.py` (lưu ở `build_b6_meta.json`), cùng nguồn và khớp số của `build_b5_meta.json`.
> **Trạng thái 03/10/2026:** đã có kịch bản, deck (`Slide_B6.pptx` + `_HocVien`, dựng bằng `build_slides_b6.py`, đã chèn đủ ảnh chụp) và **3 file Demo** (`Demo/`, dựng bằng `build_b6_demos.py` trên Mac, đã mở bằng Excel kiểm số). **Đã dựng 04/10/2026:** handout (`Handout/B6_Handout.docx`, `build_docs_b6.py`), infographic (`Handout/B6_Infographic_TomTat.png`), brief capstone (`../B7_Capstone/Brief/Capstone_Brief_HatViet.docx`) và dataset capstone Hạt Việt Coffee (`../B7_Capstone/Du_lieu_Capstone/`). **Đã dựng 04/10/2026 (Excel Windows, qua macro VBA nhập vào Excel):** đáp án `Practice/B6_MCI_Mart_Dashboard_Answer.xlsx` (`Kiem_tra_B6` 35/35 ✔) và bài tập về nhà `Homework/B6_BTVN_MCI_Mart_Dashboard_Marketing.xlsx` + `_Answer.xlsx` (`Kiem_tra` 10/10 ✔). **Đã chụp 04/10/2026:** 23 ảnh trong `Materials/B6/Images/` (deck + handout), đã chèn vào `Slide_B6.pptx`, `_HocVien` và `B6_Handout.docx`. Waterfall S13 chụp với trục bắt đầu từ 240.000.000 để thấy chênh lệch (khác thông điệp S09: ghi chú khi giảng). **Số dự báo ETS ở Slide 24 đo bằng hàm `FORECAST.ETS` trên Excel for Mac 365** (cùng thuật toán Forecast Sheet); kiểm lại một lần trên Excel Windows khi chụp ảnh.

## Cách dùng kịch bản

- **Lời giảng** là lời nói mẫu. Giảng viên không cần đọc nguyên văn; nên nói theo giọng của mình nhưng giữ đủ **ý chính** và **câu hỏi cho lớp** (in đậm).
- Dòng *(→ …)* trong lời giảng là **chỉ dẫn hành động**. Đường dẫn file tính từ thư mục buổi này (`1_Sessions/B6_Dashboard/`); file khởi đầu nằm ở `../B5_Data_Model_DAX/Practice/`.
- Mỗi slide có thời lượng dự kiến. Nếu lớp chậm, cắt ở những chỗ đánh dấu **[có thể rút gọn]**.
- **File khởi đầu:** `../B5_Data_Model_DAX/Practice/B5_MCI_Mart_Answer.xlsx` (đáp án B5), lưu thành `B6_<HoTen>.xlsx`. Data Model đã có 6 bảng, 6 relationship, 14 measure, sheet `KPI_Dictionary`. Dashboard dựng **ngay trong file này** (sheet mới `Dashboard`, `Data_Dash`, `Cai_dat`, `Brief`); lớp làm cùng giảng viên từng bước ở Phần 6.
- **Không bấm Refresh All** (tham số `DataFolder` của file khởi đầu trỏ tới máy dựng file; xem lưu ý ở `Slide_Script_B5.md`). Dữ liệu đã nằm sẵn trong Data Model, nên dashboard chạy bình thường. Phần "Refresh on open" ở Slide 33 chỉ **nói và bật**, không thử trên lớp.
- **Chỉ dùng 6 bảng của Data Model** khi dựng Pivot (tab All của PivotTable Fields còn hiện bảng worksheet của B1–B3).
- **Bật add-in Power Pivot** và kiểm tra trước giờ học: `CUBEVALUE` chỉ chạy khi file có Data Model. Máy trình chiếu cần Excel 365 Windows (Mac không có Data Model đầy đủ).
- **Map chart** cần Internet (Bing). Máy phòng học không có mạng thì dùng phương án bar chart theo thành phố (Slide 13, 29).
- **Dashboard trong buổi nhìn một tháng một lần.** Target của Orders (1.000 đơn) và Active Customers (700 khách) là target **theo tháng**, nên Slicer Tháng chọn **một** tháng. Mặc định tháng 4/2025; lật sang tháng 3 và tháng 2 để kể chuyện.

## Quy ước

Loại slide và nguồn ảnh giống `Slide_Script_B5.md`: BÌA · CHIA PHẦN · KHÁI NIỆM · CÔNG CỤ · THỰC HÀNH · QUIZ · TỔNG KẾT; **[CHỤP]** ảnh chụp Excel 365 Windows trên file của khóa, **[DV PDF pX]** trang `4_Reference/Excel_Data_Visualization.pdf` (chỉ để trong kịch bản, không ghi lên slide), **[VẼ]** vẽ bằng shape, chỉ cho sơ đồ khái niệm và thẻ số.

**Quy tắc RAG của buổi** (theo KPI Dictionary B5, Warning 90%): **✔ Đạt** khi mức đạt ≥ 100% · **▲ Cảnh báo** khi 90% đến dưới 100% · **✘ Không đạt** khi dưới 90%. *Mức đạt = giá trị ÷ target* (YoY: YoY ÷ 10%). Không dùng màu đơn độc: luôn kèm biểu tượng và chữ. Ngưỡng nằm ở sheet `Cai_dat`, đổi được. *(Ngưỡng theo KPI Dictionary B5: Target 100%, Warning 90%; mức Critical 80% chỉ nằm trong KPI Dictionary, dashboard không có màu riêng cho mức này — xem "Ghi chú so với Course Plan".)*

**File dùng trong buổi** (đường dẫn tính từ thư mục buổi; `../B5_Data_Model_DAX/` là thư mục Buổi 5):

| File | Công dụng | Slide |
|---|---|---|
| `../B5_Data_Model_DAX/Practice/B5_MCI_Mart_Answer.xlsx` | **File khởi đầu** (đáp án B5): Data Model 6 bảng, 6 relationship, 14 measure, `KPI_Dictionary`, `DoiSoat_B2` (mẫu CUBEVALUE). Lưu thành `B6_<HoTen>.xlsx`; dựng dashboard ngay trong file này | 3, 17–21, 26–30, 32–34 |
| `../B5_Data_Model_DAX/Homework/B5_BTVN_MCI_Mart_KPI_T1_T4_2025_Answer.xlsx` | Đáp án bài về nhà B5, chữa nhanh đầu buổi | 3 |
| `Demo/B6_Demo_DashboardXau.xlsx` (+ `_HocVien`) | **Đã dựng 03/10/2026** (`build_b6_demos.py`). Sheet `Dashboard_Xau`: 12 chart cố ý xấu (pie 3D 6 phần, trục Y cắt khúc, không target, màu cầu vồng); `Demo_TrucY`: hai cột T3/T4/2024; `Du_lieu`; `Huong_dan`. Để làm bài kiểm tra 5 giây | 2, 8, 9 |
| `Demo/B6_Demo_Chart.xlsx` (+ `_HocVien`) | **Đã dựng 03/10/2026.** Dữ liệu cho chart lõi (giảng viên vẽ chart trên lớp): `Demo_Combo` (ngành), `Demo_Waterfall` (cửa hàng, T3), `Demo_Phan_phoi` (3.144 giá trị đơn Q1), `Demo_Scatter` (12 cửa hàng), `Demo_Map` (thành phố, T4) | 11–14 |
| `Demo/B6_Demo_DuBao.xlsx` (+ `_HocVien`) | **Đã dựng 03/10/2026.** Table `tblNgay` (Ngay, Doanh thu; 486 ngày, 01/01/2024–30/04/2025) ở `Demo_Trendline`; `Thong_ke` (theo thứ, độ dốc, R²); `Demo_ETS` (3 hàm ETS) | 23, 24 |
| `Practice/B6_MCI_Mart_Dashboard_Answer.xlsx` | **Đã dựng (04/10/2026).** Đáp án B6: file khởi đầu + dashboard hoàn chỉnh (`Dashboard`, `Data_Dash`, `Cai_dat`, `Brief`, `Du_bao`), `Kiem_tra_B6`. Phát cho học viên chưa làm kịp; là file mẫu để chấm chéo | 30, 34 |
| `Homework/B6_BTVN_MCI_Mart_Dashboard_Marketing.xlsx` (+ `_Answer`) | **Đã dựng (04/10/2026).** Bài tập về nhà bắt buộc, tập dượt C1, C5, C6 (đặc tả ở Phụ lục); đề có `Dashboard_Cu` với 5 lỗi cài sẵn | 41 |
| `../B7_Capstone/Rubric_B7_Capstone.md` + brief + dataset capstone | Phát cuối buổi (bản Student-Friendly). Brief và dataset capstone đã dựng (xem trên) | 41 |
| `Handout/B6_Handout.docx` / `.pdf` · `Handout/B6_Infographic_TomTat.png` | **Đã dựng** (`build_docs_b6.py` → `.docx`; PDF 12 trang xuất bằng Word 05/10/2026; infographic dựng lại cùng ngày sau khi đổi ngưỡng RAG sang 90%) | 41 |

**Số liệu dùng trong buổi** (`build_b6_meta.json`):

| Chỉ tiêu | Giá trị |
|---|---|
| KPI card tháng 4/2025 (mặc định) | Doanh thu **296.965.000** / target **288.000.000** = **103,1%** ✔ · **966** đơn (96,6% của 1.000) ▲ · AOV **307.417** (102,5% của 300.000) ✔ · GM **33,8%** (102,4% của 33%) ✔ · YoY **+24,0%** (240% của +10%) ✔ · **670** khách (95,7% của 700) ▲ |
| Tháng 3/2025 | 269.608.000 / 285.000.000 = **94,6%** ▲ (hụt **15.392.000**) · 923 đơn 92,3% ▲ · AOV 292.100 97,4% ▲ · GM 33,1% ✔ · YoY +13,7% ✔ · 637 khách 91,0% ▲ |
| Tháng 2/2025 | 245.996.000 / 298.000.000 = **82,5%** ▲ (hụt 52.004.000) · 823 đơn 82,3% ▲ · AOV 298.902 99,6% ▲ · GM 33,7% ✔ · **YoY −0,9% ✘** · 605 khách 86,4% ▲ |
| Tháng 1/2025 | 425.850.000 / 353.000.000 = **120,6%** ✔ (vượt 72.850.000) · 1.398 đơn 139,8% · AOV 304.614 · GM 35,2% · YoY +42,8% · 831 khách — **tất cả ✔** |
| Lũy kế T1–T4/2025 | Doanh thu **1.238.419.000** / target **1.224.000.000** = **101,2%** (vượt **14.419.000**); T1+T2 gộp **103,2%** (vượt 20.846.000) |
| Actual vs Target theo tháng 2025 (combo) | T1 **120,6%** · T2 **82,5%** · T3 **94,6%** · T4 **103,1%** (target T1–T4: 353 · 298 · 285 · 288 triệu) |
| Chênh lệch so với target theo cửa hàng, **tháng 3** (waterfall; tổng −15.392.000) | ST01 +8.095.000 · ST02 −9.028.000 · ST03 −2.505.000 · ST04 −1.352.000 · ST05 +5.653.000 · ST06 −5.232.000 · **ST07 −10.461.000** · ST08 +10.733.000 · ST09 −2.923.000 · ST10 +9.446.000 · ST11 −6.069.000 · **ST12 −11.749.000** |
| ST07 tháng 3 | 8.539.000 / 19.000.000 = **44,9%**; riêng ST07 = **68%** khoản hụt −15.392.000 của cả chuỗi |
| ST12 theo tháng (% target) | T1 122,4% · T2 76,9% · T3 53,0% · T4 65,2% → **dưới 90% ba tháng liền**; lũy kế T1–T4: 79.399.000 / 100.000.000 = 79,4% |
| Chênh lệch theo cửa hàng, **tháng 4** (tổng +8.965.000) | ST01 −4.303.000 · ST02 +11.714.000 · ST03 +4.980.000 · ST04 −5.353.000 · ST05 +1.253.000 · ST06 +1.140.000 · ST07 −527.000 · ST08 +4.175.000 · ST09 −5.136.000 · ST10 +8.846.000 · ST11 +869.000 · ST12 −8.693.000 |
| Cửa hàng < 90% target (đỏ) | T2: ST01, ST05, ST06, ST07, ST08, ST09, ST11, ST12 · T3: ST02, ST04, ST06, ST07, ST11, ST12 · T4: ST01, ST04, ST09, ST12 (dưới 80%, mức Critical: T2 ST01, ST05, ST08, ST09, ST11, ST12 · T3 ST02, ST06, ST07, ST11, ST12 · T4 ST04, ST12) |
| Top 5 sản phẩm tháng 4/2025 | Hộp quà Tết cao cấp **35.100.000** · Chảo chống dính 26cm 29.155.000 · Dầu gội 650ml 26.565.000 · Mì ăn liền tôm chua cay (thùng 30 gói) 19.712.000 · Bột giặt 3kg 19.276.000 — 5 sản phẩm = **43,7%** doanh thu tháng |
| Thành phố, tháng 4/2025 | TP.HCM 108.885.000 · Hà Nội 62.411.000 · Đà Nẵng 24.253.000 · Hải Phòng 20.980.000 · Khánh Hòa 19.473.000 · Cần Thơ 16.869.000 · Đồng Nai 16.307.000 · Huế 16.140.000 · Quảng Ninh 11.647.000 |
| Doanh thu và GM% theo ngành, T1–T4/2025 (combo 2 trục) | Bánh kẹo 426.916.000 · **39,4%** · Gia dụng 278.397.000 · 38,5% · Thực phẩm khô 222.348.000 · **21,6%** · Chăm sóc cá nhân 143.991.000 · 37,5% · Đồ uống 143.977.000 · 27,6% · Sữa 22.790.000 · 22,5% |
| Tỷ trọng ngành T1–T4 (pie "xấu") | Bánh kẹo 34,5% · Gia dụng 22,5% · Thực phẩm khô 18,0% · Chăm sóc cá nhân 11,6% · Đồ uống 11,6% · Sữa 1,8% (6 phần, vượt giới hạn 5) |
| Trục Y cắt khúc (Good/Bad/Ugly) | Doanh thu T3/2024 **237.192.000**, T4/2024 **239.466.000**: chênh **0,96%**. Trục bắt đầu từ 236.000.000 → cột T4 cao gấp **2,9** lần cột T3 |
| Phân phối giá trị đơn Q1/2025 (= B2) | **3.144** đơn · trung bình **299.445** · trung vị **198.000** · Q1 89.000 · Q3 412.000 · ngưỡng outlier **896.500** · **138** đơn vượt ngưỡng · lệch phải (skew **1,77**); nhóm 0–100k: 905 đơn, ≥ 1.000k: 93 đơn |
| Scatter: tuổi cửa hàng ↔ doanh thu T1–T4/2025 | 12 cửa hàng, **r = 0,78** (R² 0,61). Cửa hàng ở Hà Nội + TP.HCM: tuổi TB **4,3 năm**, doanh thu TB **137,5 triệu**; cửa hàng còn lại: **2,7 năm**, **78,7 triệu** → thành phố lớn là yếu tố thứ ba |
| Doanh thu theo thứ (bảng Ngày–Doanh thu, 486 ngày) | Thứ Hai 8.434.264 · Ba 8.643.886 · Tư 8.696.329 · Năm 8.816.072 · Sáu 8.722.935 · **Bảy 10.067.797 · Chủ nhật 10.547.341** (trung bình mỗi ngày). Chủ nhật hơn thứ Hai **25,1%**; cuối tuần hơn ngày thường **19,0%** |
| Nhịp tuần nhỏ so với biến động ngày | Độ lệch chuẩn doanh thu ngày **3.144.731**; R² của "thứ trong tuần" chỉ **5,9%** (cả chuỗi), **7,7%** (03–12/2024), **15,3%** (03–04/2025) |
| Linear Trendline (x = 1..486) | độ dốc **+4.742 VNĐ/ngày**, **R² = 0,045**; Forward 30 ngày: **10.284.071 → 10.421.598** mỗi ngày (TB 10.352.834; 30 ngày ≈ **310,6 triệu**). Thực tế tháng 4: TB 9.898.833/ngày (min 6.011.000, max 15.243.000) |
| Kiểm chứng ngược Trendline (cắt dữ liệu tới 31/03/2025, dự báo tháng 4) | dự báo cả tháng **309.004.435** vs thực **296.965.000** (**+4,1%**); nhưng sai trung bình từng ngày **19,8%** giá trị trung bình ngày |
| **ETS — đo bằng hàm trên Excel for Mac 365** (seasonality 7; dữ liệu tới 30/04/2025) | Dự báo 30 ngày tháng 5/2025: tổng **237.672.111** (TB **7.922.404**/ngày — **thấp hơn 20%** so với thực tế tháng 4: 9.898.833/ngày); cuối tuần ÷ ngày thường = **1,237**. **Dải 95%**: nửa chiều rộng ngày đầu (01/05) **6.644.736**, trung bình cả 30 ngày **18.036.080** — **rộng hơn cả giá trị dự báo** (cận dưới âm) |
| **Kiểm chứng ngược ETS** (cắt dữ liệu tới 31/03/2025, seasonality 7, dự báo 30 ngày tháng 4) | Tổng dự báo **280.041.676** vs thực **296.965.000** (**−5,7%**); sai trung bình từng ngày **2.162.244** (**21,8%** giá trị TB); nửa chiều rộng dải TB **8.588.378**; **30/30 ngày** nằm trong dải 95%. (Trendline tuyến tính cùng phép thử: +4,1%, nhưng cũng sai ~20% từng ngày) |
| **`FORECAST.ETS.SEASONALITY` (tự phát hiện)** | Cả chuỗi 486 ngày: **121** · dữ liệu tới 31/03/2025: **35** · chỉ 03–12/2024: **7** → *Detect Automatically* **không** cho 7; phải **Set Manually = 7**. Dự báo điểm 01/05/2025 (thứ Năm) **6.731.497** |
| Tết | Mùng 1: **10/02/2024**, **29/01/2025**, **17/02/2026** |
| Bài về nhà B6 (Marketing) tháng 4/2025 | Active Customers **670**; theo miền 234 + 151 + 328 = **713** ≠ 670; theo phân khúc Thường **360** · Thân thiết **251** · VIP **59** (cộng = 670); doanh thu theo phân khúc 138.190.000 · 119.096.000 · 39.679.000 |

**Công thức dùng trong buổi** (`ThisWorkbookDataModel`, tên slicer `Slicer_Thang`, `Slicer_Region` — tên thật xem hộp Slicer Settings):

| Dùng cho | Công thức |
|---|---|
| KPI card (giá trị) | `=CUBEVALUE("ThisWorkbookDataModel","[Measures].[Total Revenue]","[Calendar].[Nam].&[2025]",Slicer_Thang,Slicer_Region)` |
| Target của card Doanh thu | `=CUBEVALUE("ThisWorkbookDataModel","[Measures].[Total Target]","[Calendar].[Nam].&[2025]",Slicer_Thang,Slicer_Region)` |
| Waterfall (một cửa hàng) | `=CUBEVALUE("ThisWorkbookDataModel","[Measures].[Variance]",$A5,"[Calendar].[Nam].&[2025]",Slicer_Thang)` với `$A5` = `=CUBEMEMBER("ThisWorkbookDataModel","[Store].[Store_Code].&[ST07]")` |
| Mức đạt (so sánh với hằng số) | `=B5/Cai_dat!$C$3` (giá trị ÷ target); nhãn `=IF(F5>=Cai_dat!$C$13,"Đạt",IF(F5>=Cai_dat!$C$14,"Cảnh báo","Không đạt"))` |
| Dự báo (`Demo/B6_Demo_DuBao.xlsx`, sheet `Demo_ETS`, ô B3 = 01/05/2025) | `=FORECAST.ETS(B3, tblNgay[Doanh thu], tblNgay[Ngay], 7)` · `=FORECAST.ETS.CONFINT(B3, tblNgay[Doanh thu], tblNgay[Ngay], 0.95, 7)` · `=FORECAST.ETS.SEASONALITY(tblNgay[Doanh thu], tblNgay[Ngay])` |

**Bố cục dashboard của buổi** (sheet `Dashboard`, ẩn gridline; 7 visual, ≤ 8):

| # | Visual | Nguồn | Trả lời câu hỏi brief | Slicer |
|---|---|---|---|---|
| 1 | Hàng 6 KPI card (Doanh thu · Đơn · AOV · GM% · YoY · Khách) + sparkline T1–T4, mốc target, RAG | CUBEVALUE (`Data_Dash`) | Q1: tháng này đạt bao nhiêu % target? | Tháng, Miền |
| 2 | Line: doanh thu theo ngày của tháng đang chọn | PivotChart | Q3: ngày nào bất thường? | Tháng, Miền |
| 3 | Combo: Actual (cột) vs Target (đường nét đứt) theo tháng | PivotChart | Q1: xu hướng đạt target | Miền (**không** nối Tháng: ghi "cả 4 tháng" trên tiêu đề) |
| 4 | Waterfall: Target → chênh lệch từng cửa hàng → Actual | CUBEVALUE | Q2: cửa hàng nào kéo tụt? | Tháng, Miền |
| 5 | Bar: Top 5 sản phẩm | PivotChart (Top 10 → 5) | Q4: 5 sản phẩm chiếm bao nhiêu %? | Tháng, Miền |
| 6 | Map (hoặc bar) theo thành phố | CUBEVALUE | Q5: miền / thành phố nào đóng góp? | Tháng |

**Tổng:** 41 slide, 150 phút.

---

## Phần 0 — Mở đầu: 5 giây (5 phút · slide 1–3)

### Slide 1 · BÌA · 0.5'
- **Trên slide:** MCI / Consulting & Analytics / **Excel for Business Intelligence Analyst** / Buổi 6: Thiết kế & Xây dựng Dashboard / "Bắt buộc máy Windows có Power Pivot" / "Mở sẵn: file khởi đầu B5_MCI_Mart_Answer.xlsx (đáp án Buổi 5)"
- **Ảnh:** [VẼ] nền theo Style spec.

**Lời giảng:**
> Chào cả lớp, chào mừng đến Buổi 6. Năm buổi qua các bạn đã có dữ liệu sạch, mô hình ngôi sao và một bộ KPI viết một lần. Hôm nay mình biến tất cả thành thứ mà người ra quyết định thực sự nhìn: **một trang dashboard**.

### Slide 2 · KHÁI NIỆM · Câu hỏi mở buổi · 2'
- **Trên slide:** *"Giám đốc bán hàng có 5 giây. Họ cần thấy gì đầu tiên?"* · nhãn Tầng 4 — Prescriptive (nhập môn Tầng 5).
- **Thẻ số:** hai khung cạnh nhau — **Dashboard A** (12 visual, không target) và **Dashboard B** (6 thẻ KPI, mốc target, ✔ ▲ ✘). Dòng dưới: "Tháng 4/2025: 296.965.000 · 103,1% target ✔".
- **Ảnh:** [VẼ] hai khung phác (không phải ảnh Excel).

**Lời giảng:**
> Giám đốc bán hàng bước vào phòng họp, mở máy. Họ có năm giây trước khi chuyển sang việc khác. **Câu hỏi cho lớp: trong năm giây đó họ cần thấy gì đầu tiên?** *(→ chờ 2–3 câu trả lời; ghi lên bảng)* Gần như bạn nào cũng nói cùng một điều: *tháng này đạt hay chưa đạt target*. Không ai nói "tôi cần thấy biểu đồ cột 3D mười hai màu".
>
> Đây là câu hỏi Tầng 4 — **vậy thì sao, nên làm gì**. Năm buổi trước trả lời "chuyện gì đã xảy ra, ở đâu, vì sao". Dashboard là nơi con số gặp **quyết định**: đạt thì thôi, chưa đạt thì cửa hàng nào, ngành nào, và ai cần làm gì. Cuối buổi các bạn có một trang như thế, và một sheet dự báo ngắn hạn cho tầng cuối: **sắp tới sẽ thế nào**.
>
> *(→ chỉ khung B trên slide)* Mục tiêu: sau năm giây, người xem nói được "Tháng 4 đạt 103% target, ổn". Dashboard A bên trái sẽ làm ngược lại — và mình sẽ cho các bạn xem nó ở Phần 2.

### Slide 3 · KHÁI NIỆM · Recap B5 & file khởi đầu · 2.5'
- **Trên slide:**
  - Viên gạch B5: mô hình ngôi sao 6 bảng, 6 relationship; **14 measure** dùng chung; KPI dictionary; đối soát 64 ô với KPI_SUMIFS khớp 0.
  - Bài về nhà B5: báo cáo của phòng Kinh doanh có **4 lỗi** — Target chưa nối Calendar, tháng lấy từ `Sales[Thang]`, GM% trung bình của biên từng dòng, số đơn đếm dòng. Đối soát trước, kết luận sau.
  - Hôm nay nhận **file khởi đầu** = đáp án B5 (`B5_MCI_Mart_Answer.xlsx`), lưu thành `B6_<HoTen>.xlsx`. **Không bấm Refresh All.** Mọi con số trên dashboard đến từ các measure này.
- **Ảnh:** [VẼ] viên gạch B5 (đã xong) nằm dưới viên gạch B6 (sáng lên).
- **Chân slide:** Demo: ../B5_Data_Model_DAX/Practice/B5_MCI_Mart_Answer.xlsx (sheet KPI_Dictionary) · ../B5_Data_Model_DAX/Homework/B5_BTVN_MCI_Mart_KPI_T1_T4_2025_Answer.xlsx (sheet Dap_an_loi)

**Lời giảng:**
> *(→ mở `../B5_Data_Model_DAX/Practice/B5_MCI_Mart_Answer.xlsx`, sheet KPI_Dictionary)* Nhắc nhanh buổi trước: mười bốn measure, mỗi KPI một định nghĩa. Hôm nay mình **không viết thêm một dòng DAX nào cho KPI**. Mọi con số trên dashboard sẽ gọi tên các measure này. Đó là nguyên tắc "một nguồn sự thật" đi tới tầng hiển thị.
>
> Bài về nhà B5: bốn lỗi trong báo cáo của phòng Kinh doanh. *(→ nếu còn thời gian, mở `../B5_Data_Model_DAX/Homework/B5_BTVN_MCI_Mart_KPI_T1_T4_2025_Answer.xlsx`, sheet Dap_an_loi)* Lỗi khó thấy nhất là GM% lệch 0,1 đến 0,5 điểm — chỉ người đối soát từng ô mới bắt được. Thói quen đó hôm nay còn dùng ở Slide 34.
>
> Các bạn lưu file khởi đầu thành `B6_<HoTen>.xlsx`. Và vẫn như buổi trước: **không bấm Refresh All**.

---

## Phần 1 — Dashboard brief: brief trước, chart sau (10 phút · slide 4–6)

### Slide 4 · CHIA PHẦN · 0.5'
- **Trên slide:** MCI / **1. Dashboard brief: brief trước, chart sau**

**Lời giảng:**
> Phần 1. Sai lầm phổ biến nhất khi làm dashboard là mở Excel ra vẽ chart trước khi biết ai xem và xem để làm gì.

### Slide 5 · KHÁI NIỆM · Dashboard brief · 4.5'
- **Trên slide:** Thứ tự thiết kế: **Audience → Câu hỏi/quyết định → KPI → Layout → Chart → Build → Story → Phát hành.** Brief gồm 4 mục:

| Mục | Câu hỏi | Dashboard MCI Mart |
|---|---|---|
| Người xem | Executive, operational hay self-service? | **Sales Director** — executive, xem trong họp ban điều hành |
| Câu hỏi (3–5, đo lường được) | Chỉ số · chiều · thời gian · mốc so sánh | Q1 Doanh thu tháng này đạt bao nhiêu % target, so với tháng trước và cùng kỳ? · Q2 Cửa hàng nào dưới 90% target? · Q3 Ngày nào trong tháng bất thường? · Q4 5 sản phẩm đầu chiếm bao nhiêu % doanh thu? · Q5 Thành phố nào đóng góp nhiều nhất? |
| Tần suất refresh | Bao lâu một lần? | Hằng tháng, khi có file tháng mới |
| Quyết định | Dashboard giúp quyết định gì? | Hỗ trợ cửa hàng nào trước; điều chỉnh target |

  - Câu hỏi **yếu**: "Xem doanh thu", "Phân tích khách hàng" (chủ đề, không có chỉ số, chiều, mốc so sánh).
- **Ảnh:** [VẼ] bảng brief.
- **Chân slide:** Mẫu brief: sheet Brief trong B6_<HoTen>.xlsx (tạo ở Slide 6). Rubric C1: ../B7_Capstone/Rubric_B7_Capstone.md

**Lời giảng:**
> Brief là một trang giấy, bốn mục. **Người xem**: Giám đốc bán hàng — executive, ít thời gian. **Câu hỏi**: ba đến năm câu, và câu nào cũng phải *đo lường được*. **Tần suất refresh**: hằng tháng. **Quyết định**: dashboard này giúp quyết định điều gì.
>
> **Câu hỏi cho lớp: "Xem doanh thu" có phải là câu hỏi tốt không?** *(→ chờ)* Không. Đó là một chủ đề. Câu hỏi tốt có bốn thành phần: chỉ số, chiều, thời gian, mốc so sánh. "Cửa hàng nào có doanh thu tháng này dưới 90% target?" — chỉ số là doanh thu, chiều là cửa hàng, thời gian là tháng này, mốc so sánh là 90% target. Có câu hỏi như vậy thì mới biết cần **chart nào**, **measure nào**.
>
> Hôm nay dashboard của lớp trả lời năm câu trên slide. Giữ chúng lại: mỗi visual phải trả lời đúng một câu. Visual nào không trả lời câu nào thì bỏ.

### Slide 6 · THỰC HÀNH · Viết brief của bạn · 5'
- **Trên slide:**
  1. Tạo sheet **Brief** trong `B6_<HoTen>.xlsx`.
  2. Chép bảng 4 mục của Slide 5, rồi **viết lại 2 câu hỏi theo cách của bạn** (đủ chỉ số · chiều · thời gian · mốc so sánh).
  3. Đổi cặp với bạn bên cạnh: câu nào còn là chủ đề thì gạch và viết lại.
- **Kiểm tra:** mỗi câu hỏi của bạn có đủ 4 thành phần; không câu nào bắt đầu bằng "Xem…" hoặc "Phân tích…".

**Lời giảng (giao bài):**
> *(→ học viên tự làm 4 phút, giảng viên đi vòng)* Các bạn có năm phút. Tạo sheet Brief, chép bảng, và viết lại hai câu hỏi. Rồi đổi cặp: bạn bên cạnh đọc, nếu thấy câu nào chỉ là chủ đề thì gạch đi. *(→ gọi 1–2 bạn đọc câu hỏi; giảng viên hỏi "mốc so sánh ở đâu?")* Brief này là bài tập của tiêu chí C1 trong capstone, nên mình làm kỹ ngay từ bây giờ.

---

## Phần 2 — Clarity & simplicity: dashboard xấu và nguyên tắc (10 phút · slide 7–9)

### Slide 7 · CHIA PHẦN · 0.5'
- **Trên slide:** MCI / **2. Rõ ràng và đơn giản: dashboard tốt, xấu, "ugly"**

**Lời giảng:**
> Phần 2: mình cho các bạn xem một dashboard xấu, và các bạn tự rút nguyên tắc.

### Slide 8 · KHÁI NIỆM · Bài kiểm tra 5 giây · 4.5'
- **Trên slide:** *(Ảnh dashboard xấu: 12 visual, pie 3D, trục cắt khúc, không target)* · "Bạn có 5 giây. Ghi lại: (1) KPI chính là gì? (2) Đạt hay chưa đạt target? (3) Nên nhìn vào đâu tiếp?"
- **Ảnh:** [CHỤP] S08: sheet `Dashboard_Xau` — 12 visual cùng lúc; tên file `S08_dashboard_xau.png`.
- **Chân slide:** Ảnh: Demo/B6_Demo_DashboardXau.xlsx, sheet Dashboard_Xau (toàn trang)

**Lời giảng:**
> *(→ mở `Demo/B6_Demo_DashboardXau.xlsx`, sheet Dashboard_Xau, chiếu 5 giây rồi che màn hình hoặc chuyển slide)* Bài kiểm tra năm giây. Mình chiếu năm giây, rồi các bạn ghi: KPI chính là gì, đạt hay chưa đạt target, và nên nhìn đâu tiếp. *(→ chiếu 5 giây, rồi ẩn)* Ai ghi được "đạt hay chưa đạt target"? *(→ chờ; thường không ai)*
>
> *(→ chiếu lại, đi qua từng vấn đề, ghi lên bảng)* Mười hai biểu đồ, không cái nào là chủ. Pie 3D sáu phần. Cột có trục bắt đầu từ hai trăm ba mươi sáu triệu. Không có target ở đâu cả. Màu cầu vồng — màu nào là miền nào? **Câu hỏi cho lớp: các bạn đếm được bao nhiêu vấn đề?** *(→ chờ, ghi)* Từng vấn đề sẽ có một nguyên tắc đi kèm, và các nguyên tắc đó nằm trong checklist ở Slide 34.

**Ghi chú giảng viên:**
- Đáp án mong đợi (≥ 6): quá nhiều visual; pie 3D; pie quá 5 phần (6 phần, Sữa chỉ 1,8%); trục Y không bắt đầu từ 0; không có target / mốc so sánh; không có thứ bậc (KPI chính không ở góc trên trái); màu không theo thực thể (Bắc đổi màu giữa các chart); không có slicer; tiêu đề chart là nhãn, không phải kết luận.
- Dashboard xấu dựng từ chính dữ liệu của lớp để học viên nhận ra con số quen thuộc.

### Slide 9 · KHÁI NIỆM · Tốt · Xấu · "Ugly" · 5'
- **Trên slide:** 5 nguyên tắc, mỗi nguyên tắc một ví dụ:
  1. **Trục bar/column bắt đầu từ 0.** T3/2024 237.192.000 vs T4/2024 239.466.000: chênh **0,96%**. Trục từ 236.000.000 → cột T4 **cao gấp 2,9 lần**.
  2. **Không 3D.** Chiều sâu làm méo phần trăm.
  3. **Pie ≤ 5 phần.** Ngành T1–T4: 34,5% · 22,5% · 18,0% · 11,6% · 11,6% · **1,8%** — sáu phần, phần Sữa không đọc được. Dùng bar.
  4. **Mỗi KPI có mốc so sánh** (target, kỳ trước).
  5. **Màu theo thực thể**, mốc so sánh luôn là đường xám nét đứt.
- **Ảnh:** [CHỤP] S09: hai cột T3/2024 và T4/2024 — trục bắt đầu từ 236.000.000 (xấu) và từ 0 (tốt), đặt cạnh nhau; `S09_truc_y.png`.
- **Chân slide:** Ảnh: Demo/B6_Demo_DashboardXau.xlsx, sheet Demo_TrucY (hai biểu đồ cột cạnh nhau)

**Lời giảng:**
> Năm nguyên tắc, mỗi nguyên tắc một con số. Cái nguy hiểm nhất là **trục Y**. *(→ chỉ ảnh S09)* Doanh thu tháng 3 và tháng 4 năm 2024 chênh nhau chưa tới **một phần trăm**: 237,2 triệu và 239,5 triệu. Nhưng nếu trục bắt đầu từ 236 triệu thay vì 0, cột tháng 4 cao gấp **gần ba lần** cột tháng 3. Cùng một dữ liệu, một bên nói "không đổi", một bên nói "tăng vọt". Quy tắc: **bar và column luôn bắt đầu từ 0**. Line chart thì được phép cắt trục, vì ta nhìn *hình dáng*, không so độ cao.
>
> Pie: sáu ngành, phần nhỏ nhất là Sữa 1,8% — các bạn có đọc được không? Quá năm phần thì đổi sang bar. 3D thì không bao giờ: chiều sâu làm phần phía trước to hơn. **Mốc so sánh**: số 296 triệu một mình thì chẳng nói gì; 296 triệu so với target 288 triệu mới nói được "đạt". Và màu đi theo thực thể: Bắc luôn là xanh dương, không đổi giữa các chart.
>
> Nguyên tắc chung của cả phần này một câu: **mỗi thứ trên trang phải xứng đáng với sự chú ý mà nó chiếm**.

---

## Phần 3 — Chọn chart bằng 3 câu hỏi (20 phút · slide 10–14)

### Slide 10 · CHIA PHẦN · 0.5'
- **Trên slide:** MCI / **3. Chọn chart: 3 câu hỏi, 8 loại lõi**

**Lời giảng:**
> Phần 3: vì sao chart này mà không phải chart kia — và vì sao phần lớn câu trả lời nằm ở hai câu hỏi.

### Slide 11 · KHÁI NIỆM · 3 câu hỏi then chốt · bảng chọn chart · 4'
- **Trên slide:** **3 câu hỏi then chốt** [DV PDF p5]: (1) Dữ liệu thuộc loại nào? (số, danh mục, thời gian, địa lý) · (2) Muốn truyền đạt điều gì? (so sánh, cơ cấu, xu hướng, phân phối, quan hệ) · (3) Ai xem?

| Câu hỏi dữ liệu | Chart lõi | Thay thế |
|---|---|---|
| Xu hướng theo thời gian | **Line** | Area |
| So sánh giữa các nhóm | **Bar / Column** | Bullet |
| Cơ cấu / tỷ trọng | **Stacked bar** | Donut (≤ 5 phần), Treemap |
| Hai thước đo khác đơn vị; Actual vs Target | **Combo 2 trục** | Column + đường target |
| Phân phối, outlier | **Histogram** | Box & Whisker |
| Quan hệ giữa hai biến | **Scatter** (+ trendline) | Bubble |
| Địa lý | **Map** | Bar theo vùng |
| Chênh lệch cộng dồn (target → actual) | **Waterfall** | — |

- **Ảnh:** [VẼ] bảng chọn chart (mỗi dòng một biểu tượng chart nhỏ).

**Lời giảng:**
> Ba câu hỏi. **Dữ liệu của bạn là loại gì**: số, danh mục, thời gian, địa lý. **Bạn muốn người xem nhận ra điều gì**: so sánh, cơ cấu, xu hướng, phân phối, quan hệ. **Ai xem**: giám đốc cần ít chart và chữ to hơn analyst. Hai câu đầu gần như quyết định chart; câu thứ ba quyết định *bao nhiêu* chart.
>
> Bảng bên dưới là tám loại chart lõi của khóa. Các bạn đừng học thuộc cả bảng — chỉ cần nhớ cột trái là **câu hỏi**, cột giữa là chart. Hôm nay mình làm bốn nhóm: so sánh theo thời gian và nhóm, cơ cấu, phân phối và quan hệ, rồi chênh lệch cộng dồn. Các chart ít dùng hơn — Area, Pie, Bubble, Radar, Funnel, Sunburst — nằm trong handout.

### Slide 12 · CÔNG CỤ · Line · Column/Bar · Combo 2 trục · 5.5'
- **Tên (đỏ):** **Insert → Charts → Combo** · **Change Chart Type → Combo → Secondary Axis**
- **Mục đích:** Đặt hai thước đo khác đơn vị (VNĐ và %) trên cùng một chart; hoặc Actual (cột) với Target (đường).
- **Trên slide:** 
  - Line: xu hướng theo thời gian · Column/Bar: so sánh nhóm (trục từ 0).
  - **Combo ngành, T1–T4/2025:** cột = Doanh thu (Bánh kẹo 426.916.000 … Sữa 22.790.000), đường = Gross Margin % (Bánh kẹo **39,4%** … Thực phẩm khô **21,6%**), đường nằm ở **trục phụ**.
  - Insight: Thực phẩm khô là ngành lớn thứ ba (222.348.000) nhưng biên **thấp nhất** (21,6%).
- **TIP:** Cột trục chính (VNĐ), đường trục phụ (%). Cho trục phụ bắt đầu từ 0%.
- **Ảnh:** [CHỤP] S12: combo cột Doanh thu + đường GM% theo ngành; `S12_combo_nganh.png`.
- **Chân slide:** Ảnh: Demo/B6_Demo_Chart.xlsx, sheet Demo_Combo (A3:C9)

**Lời giảng:**
> *(→ mở `Demo/B6_Demo_Chart.xlsx`, sheet Demo_Combo)* Bảng nhỏ: sáu ngành, doanh thu và biên lợi nhuận gộp T1–T4. Mình muốn nhìn hai thứ cùng lúc, nhưng một cái tính bằng triệu đồng, một cái tính bằng phần trăm. *(→ chọn A3:C9 → Insert → Recommended Charts → All Charts → Combo; Doanh thu: Clustered Column; Gross Margin %: Line, tích Secondary Axis → OK)* Cột là doanh thu theo trục trái, đường là biên theo trục phải.
>
> **Câu hỏi cho lớp: nhìn chart này, ngành nào đáng chú ý?** *(→ chờ)* Thực phẩm khô: cột cao thứ ba nhưng đường rơi xuống thấp nhất, 21,6%. Bánh kẹo bán nhiều nhất và biên cao nhất, 39,4%. Hai thước đo cùng một chart cho thấy điều mà từng chart riêng lẻ không cho thấy: **bán nhiều không có nghĩa lãi nhiều**. Đó là lý do combo hai trục có mặt trong tám loại lõi.
>
> Một lưu ý: trục phụ rất dễ gây hiểu lầm nếu bạn cho hai trục bắt đầu từ chỗ khác nhau. Giữ trục phần trăm bắt đầu từ 0.

### Slide 13 · CÔNG CỤ · Cơ cấu · Chênh lệch cộng dồn · Địa lý · 4.5'
- **Tên (đỏ):** **Stacked Bar** · **Waterfall (Insert → Waterfall)** · **Filled Map**
- **Trên slide:**
  - **Cơ cấu:** Stacked bar thay pie khi hơn 5 phần; Treemap khi cần thấy kích thước tương đối.
  - **Waterfall — từ Target đến Actual, tháng 3/2025:** Target **285.000.000** → chênh lệch từng cửa hàng (ST07 **−10.461.000**, ST12 **−11.749.000**, ST08 +10.733.000 …) → Actual **269.608.000** (tổng −15.392.000). Đặt cột đầu và cột cuối là **Total**.
  - **Map:** theo thành phố (TP.HCM 108.885.000 đứng đầu, T4). Cần Internet; không có thì dùng bar.
  - **Giới hạn:** PivotChart **không** làm được Waterfall, Map, Histogram, Treemap — những chart này lấy dữ liệu từ **một vùng ô**.
- **TIP:** Muốn waterfall chạy theo Slicer, lấp vùng ô bằng **CUBEVALUE** (Slide 18, 29).
- **Ảnh:** [CHỤP] S13: waterfall target → actual tháng 3 theo cửa hàng; `S13_waterfall.png`.
- **Chân slide:** Ảnh: Demo/B6_Demo_Chart.xlsx, sheet Demo_Waterfall (A3:B17)

**Lời giảng:**
> *(→ `Demo/B6_Demo_Chart.xlsx`, sheet Demo_Waterfall)* Cơ cấu: pie sáu phần các bạn đã thấy là xấu; stacked bar hoặc treemap làm tốt hơn. Mình đi vào chart quan trọng nhất của dashboard hôm nay: **waterfall**.
>
> Câu hỏi nó trả lời: *từ target đến actual, ai kéo lên, ai kéo xuống?* Tháng 3, target chuỗi là 285 triệu, thực tế 269,6 triệu, hụt 15,4 triệu. *(→ chọn A3:B17 → Insert → Waterfall; bấm chuột phải cột Target → Set as Total; cột Actual → Set as Total)* Mười hai cửa hàng, mỗi cái một bước. **Câu hỏi cho lớp: cửa hàng nào kéo tụt nhiều nhất?** *(→ chờ)* ST12 trừ 11,7 triệu, rồi ST07 trừ 10,5 triệu. Nhưng nhìn cả bốn cái kéo lên, ST08, ST10, ST01, ST05, cộng lại gần bù hết. Waterfall cho thấy cả bức tranh.
>
> Một giới hạn quan trọng. *(→ chỉ dòng "Giới hạn")* Waterfall, map, histogram, treemap **không có trong PivotChart**. PivotChart chỉ hỗ trợ line, column, bar, pie, combo và vài loại nữa. Muốn các chart kia thì dữ liệu phải nằm trong **một vùng ô bình thường**. Điều đó dẫn tới một cách làm: dùng CUBEVALUE để lấp vùng ô từ Data Model — mình làm ở Slide 18 và 29. Map thì cần Internet; bạn nào không có mạng thì thay bằng bar theo thành phố.

### Slide 14 · CÔNG CỤ · Phân phối · Quan hệ: Histogram · Box · Scatter · 5.5'
- **Tên (đỏ):** **Insert → Histogram / Box & Whisker** · **Insert → Scatter** + **Trendline** · **CORREL**
- **Trên slide:**
  - **Histogram — giá trị đơn Q1/2025 (B2):** 3.144 đơn; trung vị **198.000**, trung bình **299.445**, lệch phải; ngưỡng outlier **896.500**, **138** đơn vượt. Bin width 100.000.
  - **Scatter — tuổi cửa hàng ↔ doanh thu T1–T4/2025:** 12 điểm, **r = 0,78**. Cửa hàng Hà Nội + TP.HCM: tuổi TB **4,3 năm**, doanh thu TB **137,5 triệu**; còn lại: **2,7 năm**, **78,7 triệu**.
  - **Tương quan ≠ nhân quả.**
- **TIP:** Histogram: Format Axis → Bin width. Scatter: Add Trendline → Display R-squared.
- **Ảnh:** [CHỤP] S14a: histogram giá trị đơn, bin 100.000; [CHỤP] S14b: scatter 12 cửa hàng + trendline + R²; `S14a_histogram.png`, `S14b_scatter.png`.
- **Chân slide:** Ảnh: Demo/B6_Demo_Chart.xlsx, sheet Demo_Phan_phoi (cột B) và Demo_Scatter (A3:C15)

**Lời giảng:**
> *(→ `Demo/B6_Demo_Chart.xlsx`, sheet Demo_Phan_phoi, cột B là giá trị từng đơn Quý 1)* Các bạn nhớ Buổi 2: trung bình đơn hàng 299 nghìn, trung vị 198 nghìn, và **138 đơn** vượt ngưỡng 896 nghìn. Giờ nhìn bằng mắt. *(→ chọn cột B → Insert → Histogram; chuột phải trục ngang → Format Axis → Bin width 100000)* Cột cao nhất ở nhóm 0–100 nghìn; đuôi kéo dài sang phải tới hơn 2 triệu. Đây là hình dạng của dữ liệu lệch phải, và là lý do trung bình bị kéo lên cao hơn trung vị.
>
> Quan hệ giữa hai biến: scatter. *(→ sheet Demo_Scatter; chọn B3:C15 → Insert → Scatter; Add Trendline → Display R-squared; ô G3 `=CORREL(B4:B15,C4:C15)`)* Mỗi chấm là một cửa hàng: trục ngang tuổi cửa hàng, trục dọc doanh thu T1–T4. Hệ số tương quan **0,78**. Cửa hàng càng lâu năm, doanh thu càng cao.
>
> **Câu hỏi cho lớp: vậy muốn tăng doanh thu thì cho cửa hàng "già" đi à?** *(→ chờ, cười)* Tất nhiên không. Nhóm cửa hàng ở Hà Nội và TP.HCM vừa mở sớm hơn, trung bình 4,3 năm, vừa nằm ở thành phố lớn — doanh thu 137 triệu so với 79 triệu của các tỉnh. **Thành phố là yếu tố thứ ba** làm cả tuổi lẫn doanh thu cùng cao. **Tương quan không phải nhân quả.** Đây là câu mà giám đốc sẽ hỏi lại các bạn ở buổi thuyết trình B7.

**Ghi chú giảng viên:**
- Thiếu thời gian: bỏ histogram, chỉ làm scatter.
- Nếu muốn mở rộng: bỏ ST12 (chỉ 0,9 năm tuổi) thì r tăng lên **0,90** — một điểm bất thường có thể làm tăng hoặc giảm hệ số; đọc scatter bằng mắt trước khi tin con số.

---

## Nghỉ giữa giờ (10 phút · slide 15)

### Slide 15 · NGHỈ · 10'
- **Trên slide:** NGHỈ GIỮA GIỜ · 10 phút · "Sau giờ nghỉ: bố cục, KPI card bằng CUBEVALUE, RAG, Slicer"

**Lời giảng:**
> Nghỉ mười phút. Sau giờ nghỉ mình bắt đầu dựng thật: bố cục, thẻ KPI bằng CUBEVALUE và đèn RAG. Các bạn nhớ giữ file `B6_<HoTen>.xlsx` mở sẵn.

---

## Phần 4 — Bố cục và công cụ dựng (20 phút · slide 16–21)

### Slide 16 · CHIA PHẦN · 0.5'
- **Trên slide:** MCI / **4. Bố cục & công cụ: CUBEVALUE · RAG · Slicer**

**Lời giảng:**
> Phần 4: bốn công cụ dựng nên dashboard. Trước hết, đặt cái gì ở đâu.

### Slide 17 · KHÁI NIỆM · Thứ bậc bố cục: KPI → Trend → Detail · 3.5'
- **Trên slide:** Sơ đồ bố cục 1 trang (7 visual): 
  - **Hàng 1 — KPI:** 6 thẻ, KPI quan trọng nhất ở **góc trên trái**.
  - **Hàng 2 — Trend:** Line theo ngày · Combo Actual vs Target.
  - **Hàng 3 — Detail:** Waterfall theo cửa hàng · Top 5 sản phẩm · Map theo thành phố.
  - **Bên trái hoặc trên cùng:** Slicer Tháng, Miền.
  - Tối đa **5–8 visual**; mỗi visual trả lời một câu trong brief (Q1–Q5).
- **Ảnh:** [VẼ] wireframe 3 hàng (không phải ảnh Excel).

**Lời giảng:**
> Mắt người đọc từ trên xuống, trái sang phải, như đọc một trang báo. Vậy **thứ quan trọng nhất ở góc trên trái**. Với Giám đốc bán hàng đó là *doanh thu so với target*. Hàng đầu tiên là sáu thẻ KPI: trả lời "đạt hay chưa" trong năm giây. Hàng thứ hai là xu hướng: line theo ngày, combo thực tế và target. Hàng cuối là chi tiết: cửa hàng nào, sản phẩm nào, thành phố nào. Từ tổng quát xuống chi tiết — đó là thứ bậc **KPI → Trend → Detail**.
>
> Và số lượng: tối đa tám visual. Mình dùng bảy. Mỗi visual gắn với một câu trong brief của Slide 5: thẻ KPI và combo trả lời Q1, waterfall trả lời Q2, line trả lời Q3, Top 5 trả lời Q4, map trả lời Q5. **Câu hỏi cho lớp: nếu có người xin thêm một chart nữa, bạn hỏi gì?** *(→ chờ)* "Chart đó trả lời câu hỏi nào trong brief?" Không có câu trả lời thì không thêm.

### Slide 18 · CÔNG CỤ · CUBEVALUE · CUBEMEMBER · KPI card · 6'
- **Tên (đỏ):** **CUBEVALUE** · **CUBEMEMBER** · **PivotTable Analyze → OLAP Tools → Convert to Formulas**
- **Mục đích:** Đọc **một con số** từ measure trong Data Model vào **một ô** — nền của thẻ KPI.
- **Công thức (trên slide):** `=CUBEVALUE("ThisWorkbookDataModel","[Measures].[Total Revenue]","[Calendar].[Nam].&[2025]",Slicer_Thang,Slicer_Region)`
- **Trên slide:**
  - Đối số: kết nối · **measure** · các **bộ lọc** (member hoặc **Slicer**).
  - Quý 1, Bắc × Tháng 3 = **90.210.000** (đối soát với `DoiSoat_B2`, B5).
  - Tháng 4/2025: Doanh thu **296.965.000** · Target **288.000.000** · Revenue vs Target % **103,1%**.
  - **TIP:** gõ sai tên measure → `#N/A`; dùng gợi ý khi gõ, hoặc Convert to Formulas để Excel tự viết.
- **Ảnh:** [CHỤP] S18a: ô CUBEVALUE trên thanh công thức (gợi ý autocomplete đang hiện); [CHỤP] S18b: 6 ô KPI thô ở sheet Data_Dash; `S18a_cubevalue.png`, `S18b_data_dash.png`.
- **Chân slide:** Ảnh: B6_<HoTen>.xlsx (từ file khởi đầu), sheet Data_Dash (A3:E9). Mẫu: Practice/B5_MCI_Mart_Answer.xlsx, sheet DoiSoat_B2 (ô B6)

**Lời giảng:**
> Có hai cách đưa số từ Data Model ra dashboard. Cách một: PivotTable — các bạn đã quen. Cách hai: **CUBEVALUE**, hàm lấy *đúng một con số* từ measure và đặt vào *đúng một ô*. Thẻ KPI cần đúng điều đó: một ô, một số, tự do định dạng và đặt ở đâu tùy ý.
>
> *(→ mở sheet `DoiSoat_B2` của file khởi đầu, ô B6 → chỉ thanh công thức)* Các bạn đã thấy hàm này ở cuối B5: ô Bắc × tháng 1. Cấu trúc: tên kết nối, luôn là `"ThisWorkbookDataModel"`; tên measure, dạng `"[Measures].[Total Revenue]"`; rồi các bộ lọc, mỗi bộ lọc là một member như `"[Calendar].[Nam].&[2025]"`. Đây chính là **filter context của Buổi 5**, viết thành công thức: mỗi đối số là một tọa độ.
>
> *(→ tạo sheet `Data_Dash`; Insert → Slicer: Calendar → Thang, đổi tên Slicer_Thang; Store → Region, Slicer_Region; chọn tháng 4 | ô B3 gõ công thức Total Revenue như trên slide)* Điều thú vị: đối số cuối có thể là **chính cái Slicer**. Bấm Slicer sang tháng 3, ô B3 đổi theo. *(→ bấm tháng 3, rồi trở lại tháng 4)* 269.608.000, rồi 296.965.000. Mình không viết lại công thức; mình chỉ đổi bộ lọc.
>
> *(→ B4: thay `Total Revenue` bằng `Total Target` → 288.000.000)* Đây là target. *(→ B5 `=B3/B4` → 103,1%?)* À — các bạn để ý: mình chỉ chia hai ô cho nhau. Mức đạt này là **phép so sánh**, không phải định nghĩa KPI; định nghĩa vẫn nằm ở measure `Revenue vs Target %`. Để chắc, ta gọi thẳng measure đó thêm một lần nữa và đối chiếu: 103,1%. Hai con số phải khớp. *(→ viết `Revenue vs Target %` vào B5, đối chiếu)*
>
> Còn ai sợ gõ sai? *(→ Pivot bất kỳ từ Data Model → PivotTable Analyze → OLAP Tools → Convert to Formulas)* Chuyển cả Pivot thành CUBEVALUE chỉ bằng một nút — Excel tự viết hết. Và nếu gõ sai tên measure, hàm trả `#N/A`: không phải lỗi mô hình, mà là gõ sai chữ. Hãy dùng gợi ý khi gõ.

**Ghi chú giảng viên:**
- Cú pháp slicer làm đối số: tên trong hộp Slicer Settings (Name to use in formulas); thường là `Slicer_Thang`, `Slicer_Region`. Nếu hàm báo lỗi, kiểm tra tên ở **Slicer → Slicer Settings**.
- Slicer cần tạo từ một PivotTable của Data Model; nếu chưa có Pivot thì Insert → Slicer chưa bật được: tạo Pivot trống trước (Insert → PivotTable → From Data Model) rồi mới chèn Slicer.
- Không đặt công thức KPI (chia, nhân) ở tầng hiển thị trừ phép **so sánh với hằng số target** ở Slide 19.

### Slide 19 · CÔNG CỤ · RAG: Conditional Formatting + biểu tượng · 4.5'
- **Tên (đỏ):** **Home → Conditional Formatting → Icon Sets** · **Rules Manager**
- **Mục đích:** Báo trạng thái đạt / cảnh báo / không đạt — **bằng màu, biểu tượng và chữ cùng lúc**.
- **Trên slide:**

| Mức | Điều kiện (mức đạt) | Biểu tượng | Chữ |
|---|---|---|---|
| Đạt | ≥ 100% | ✔ (xanh) | "Đạt" |
| Cảnh báo | 90% – < 100% | ▲ (vàng) | "Cảnh báo" |
| Không đạt | < 90% | ✘ (đỏ) | "Không đạt" |

  - Tháng 4/2025: Doanh thu 103,1% ✔ · Đơn 96,6% ▲ · AOV 102,5% ✔ · GM 102,4% ✔ · YoY ✔ · Khách 95,7% ▲.
  - Tháng 2/2025: Doanh thu 82,5% ▲ · **YoY −0,9% ✘**.
  - Ngưỡng 100% và 90% đặt ở sheet `Cai_dat`; quy tắc trỏ tới ô đó.
- **TIP:** Dùng bộ **3 Symbols (Circled)**: ba hình khác nhau (dấu tick, chấm than, dấu chéo), người mù màu vẫn phân biệt được.
- **PRO TIP:** Rules Manager → Edit Rule → Type = **Formula**, trỏ `=Cai_dat!$C$13` để đổi ngưỡng một nơi.
- **Ảnh:** [CHỤP] S19: 6 thẻ KPI tháng 4/2025 với biểu tượng ✔ ▲ ✘ và nhãn chữ; `S19_rag.png`.
- **Chân slide:** Ảnh: B6_<HoTen>.xlsx, sheet Dashboard (hàng KPI, tháng 4/2025). Ngưỡng: sheet Cai_dat (C13:C14)

**Lời giảng:**
> Thẻ KPI có số rồi, giờ cần **trạng thái**. *(→ sheet `Cai_dat`: A12 "Ngưỡng Đạt" B12 100%; "Ngưỡng Cảnh báo" 90% → đặt tên sheet này là tham số của dashboard)* Mình đặt ngưỡng ra một sheet riêng, để sau này đổi ngưỡng không phải sửa từng thẻ.
>
> *(→ ô mức đạt của thẻ Doanh thu → Home → Conditional Formatting → Icon Sets → 3 Symbols (Circled) → Manage Rules → Edit Rule: icon thứ nhất "when value is ≥" Type Number, Value 1 → icon thứ hai ≥ 0,9)* Xanh có dấu tick khi đạt 100% trở lên; vàng có dấu chấm than khi từ 90%; đỏ có dấu chéo khi dưới 90%. **Câu hỏi cho lớp: sao không chỉ tô ô màu đỏ-xanh cho nhanh?** *(→ chờ)* Khoảng **tám phần trăm đàn ông không phân biệt được đỏ và xanh**. Hình dạng khác nhau — tick, chấm than, chéo — thì ai cũng đọc được. Và kèm nhãn chữ "Đạt", "Cảnh báo", "Không đạt" cho chắc. Quy tắc: **không bao giờ chỉ dùng màu để truyền nghĩa**.
>
> *(→ bấm Slicer sang tháng 2)* Tháng 2: doanh thu 82,5% vàng. YoY: trừ 0,9% — **đỏ có dấu chéo**. Doanh thu chỉ "cảnh báo" nhưng tăng trưởng so với năm ngoái đã âm; chỉ nhìn một thẻ thì không thấy. Nhìn cả hàng thì thấy ngay.

**Ghi chú giảng viên:**
- Tỷ lệ "8% nam giới mù màu đỏ-xanh" là con số phổ biến (khoảng 1 trong 12 nam giới); nói "khoảng 8%".
- YoY: mức đạt = YoY ÷ 10% (target +10%). Tháng 2: −0,9% ÷ 10% = −9% → ✘. Giải thích là phép so với mục tiêu tăng trưởng.
- Excel có thể hiển thị nhãn "Cảnh báo" bằng Custom Number Format hoặc công thức `IF`; mình dùng `IF` vì dễ đọc.

### Slide 20 · CÔNG CỤ · Slicer nối nhiều Pivot · Report Connections · 4'
- **Tên (đỏ):** **PivotTable Analyze → Insert Slicer** · **Report Connections**
- **Mục đích:** Một Slicer điều khiển **mọi** PivotTable / PivotChart của dashboard — không chỉ cái đang chọn.
- **Trên slide:**
  - Slicer Tháng (Calendar[Thang]) và Miền (Store[Region]); tạo từ Pivot, **bắt buộc** nối thêm bằng Report Connections.
  - Chuột phải Slicer → **Report Connections** → tick từng Pivot.
  - Combo Actual vs Target (4 tháng): **chỉ nối Miền**, bỏ tick Tháng — và ghi trên tiêu đề "Cả 4 tháng · không đổi theo Slicer Tháng".
  - KPI card (CUBEVALUE) và waterfall/map (vùng ô): lọc bằng **Slicer là đối số của công thức**, không qua Report Connections.
- **TIP:** Pivot đổi cỡ cột khi Refresh làm vỡ bố cục → PivotTable Options → bỏ **Autofit column widths on update**.
- **Ảnh:** [CHỤP] S20: hộp Report Connections (Slicer_Thang) với danh sách Pivot; `S20_report_connections.png`.
- **Chân slide:** Ảnh: B6_<HoTen>.xlsx, chuột phải Slicer Tháng → Report Connections

**Lời giảng:**
> Slicer sinh ra từ một Pivot, và mặc định chỉ điều khiển **chính Pivot đó**. Lỗi quen thuộc: bấm Slicer, một nửa dashboard đổi, nửa kia đứng im. *(→ chuột phải Slicer_Thang → Report Connections)* Hộp này liệt kê mọi Pivot trong file: tick vào Pivot nào, Slicer điều khiển Pivot đó. Mình tick hết ngoại trừ một cái.
>
> Cái đó là **combo Actual vs Target theo tháng**. Nếu Slicer cắt nó về một tháng thì combo chỉ còn một cột — mất ý nghĩa xu hướng. Quy tắc hơi mềm: Slicer điều khiển mọi visual, **trừ visual cho ngữ cảnh**, và visual đó phải ghi rõ trên tiêu đề "cả 4 tháng, không đổi theo Slicer Tháng". Người xem phải biết cái nào đổi, cái nào không.
>
> Còn thẻ KPI, waterfall, map không có Pivot — chúng nghe Slicer qua **đối số của CUBEVALUE**, như các bạn vừa viết ở Slide 18. Hai đường dẫn khác nhau, cùng một Slicer.

### Slide 21 · CÔNG CỤ · Sparklines · Alt text · Tùy chọn Pivot · 1.5'
- **Tên (đỏ):** **Insert → Sparklines** · **Chuột phải → Edit Alt Text**
- **Trên slide:**
  - **Sparkline** là biểu đồ nhỏ trong **một ô**: dưới mỗi thẻ KPI, vẽ T1–T4 (Doanh thu 425,9 → 246,0 → 269,6 → 297,0 triệu).
  - **Alt text:** mỗi chart có mô tả ngắn cho người dùng đọc màn hình ("Biểu đồ cột và đường: doanh thu thực tế so với target, T1–T4/2025").
  - Lấy dữ liệu sparkline từ hàng T1–T4 trong `Data_Dash` (CUBEVALUE theo từng tháng).
- **Ảnh:** [CHỤP] S21: một thẻ KPI có sparkline T1–T4; `S21_sparkline.png`.
- **Chân slide:** Ảnh: B6_<HoTen>.xlsx, sheet Dashboard (thẻ Doanh thu, sparkline) · Data_Dash (hàng T1–T4)

**Lời giảng:**
> Hai thứ nhỏ nhưng ăn tiền. **Sparkline**: biểu đồ thu nhỏ nằm trong một ô, đặt dưới con số của thẻ KPI. Thẻ doanh thu: 296 triệu, kèm một đường nhỏ cho thấy tháng 1 cao, tháng 2 rơi, rồi hồi dần. Người xem thấy ngay "đang đi lên hay xuống" mà không cần chart riêng. *(→ Insert → Sparklines → Line → Data Range: Data_Dash!B12:E12; Location: ô dưới thẻ)*
>
> **Alt text**: mô tả chart bằng một câu cho người dùng đọc màn hình, và cho cả bạn khi chart hỏng không hiện. *(→ chuột phải chart → Edit Alt Text → gõ mô tả)* Đây là một dòng trong checklist chấm chéo và là yêu cầu của bài capstone.

---

## Phần 5 — Dự báo nhập môn (Tầng 5) (10 phút · slide 22–24)

### Slide 22 · CHIA PHẦN · 0.5'
- **Trên slide:** MCI / **5. Dự báo nhập môn: sắp tới sẽ thế nào?**

**Lời giảng:**
> Phần 5: tầng cuối của thang — "sắp tới sẽ thế nào". Mình chỉ làm nhập môn: hai công cụ và một giới hạn.

### Slide 23 · CÔNG CỤ · Trendline tuyến tính + Forward · 4.5'
- **Tên (đỏ):** **Chart Elements → Trendline → Linear** · **Forecast Forward**
- **Mục đích:** Kéo dài xu hướng bằng đường thẳng.
- **Trên slide:**
  - Bảng **Ngày–Doanh thu**: 486 dòng, 01/01/2024–30/04/2025 (không thiếu ngày).
  - Linear Trendline: độ dốc **+4.742 VNĐ/ngày**, **R² = 0,045** — đường thẳng giải thích **4,5%** biến động.
  - Forward 30 ngày: **≈ 10,3 triệu/ngày**, đều như nhau cả cuối tuần.
  - Thực tế theo thứ: Thứ Hai **8,4** triệu → Chủ nhật **10,5** triệu (+25,1%).
- **Câu hỏi:** "Đường này có tin được không?"
- **Ảnh:** [CHỤP] S23: line chart doanh thu ngày + Linear Trendline + Forward 30; `S23_trendline.png`.
- **Chân slide:** Ảnh: Demo/B6_Demo_DuBao.xlsx, sheet Demo_Trendline (tblNgay, biểu đồ line)

**Lời giảng:**
> *(→ mở `Demo/B6_Demo_DuBao.xlsx`, sheet Demo_Trendline; bảng tblNgay: Ngày, Doanh thu, 486 dòng)* Bảng Ngày–Doanh thu của cả hai năm, mỗi ngày một dòng. *(→ chọn bảng → Insert → Line chart; Chart Elements → Trendline → Linear; Format Trendline → Forecast Forward 30 periods; Display R-squared)* Đường thẳng dốc nhẹ lên, kéo dài thêm 30 ngày. **Câu hỏi cho lớp: đường này có tin được không?** *(→ chờ)*
>
> Nhìn R-bình-phương: **0,045**. Đường thẳng này chỉ giải thích được bốn phẩy năm phần trăm biến động của doanh thu ngày. Nó dự báo mỗi ngày khoảng 10,3 triệu, thứ Hai cũng như Chủ nhật. Trong khi dữ liệu thật: thứ Hai 8,4 triệu, Chủ nhật 10,5 triệu — **Chủ nhật hơn thứ Hai 25%**. Đường thẳng bỏ qua hoàn toàn nhịp tuần.
>
> Công bằng với đường thẳng: mình thử dự báo ngược. Cắt dữ liệu tới hết tháng 3, vẽ đường thẳng, dự báo tháng 4: tổng cả tháng ra 309 triệu, thực tế 297 triệu, lệch chỉ 4%. Nhưng nếu hỏi *từng ngày*, sai trung bình gần **20%**. Đường thẳng tạm được cho cả tháng, tệ cho từng ngày. Hãy nhớ phân biệt "dự báo mức" và "dự báo hình dáng".

**Ghi chú giảng viên:**
- Trendline của Excel trên line chart dùng x = 1..n (không phải serial ngày): độ dốc +4.742 VNĐ/ngày và R² 0,045 (số tính bằng OLS cùng quy ước).
- Kiểm chứng ngược bằng Trendline: học viên cắt bảng tới 31/03/2025 (bản sao bảng), vẽ trendline, đọc phương trình `y = 5.027x + …`; dự báo tháng 4 = 309.004.435. Nếu thiếu thời gian chỉ nói kết quả.

### Slide 24 · CÔNG CỤ · Forecast Sheet (ETS) · dải tin cậy · giới hạn · 5'
- **Tên (đỏ):** **Data → Forecast → Forecast Sheet** · `FORECAST.ETS` · `FORECAST.ETS.CONFINT` · `FORECAST.ETS.SEASONALITY`
- **Mục đích:** Dự báo **có nhịp lặp** (mùa vụ), kèm **dải tin cậy**.
- **Trên slide:**
  - Hộp thoại: **Forecast End** (30/05/2025) · **Confidence Interval** 95% · **Seasonality: Set Manually = 7** (Detect Automatically **không** ra 7) · Timeline Range, Values Range · **Forecast Start** (có thể chọn ngày trước ngày cuối dữ liệu để so với thực tế).
  - Kết quả tháng 5/2025 (seasonality 7): tổng 30 ngày **237,7 triệu** — thấp hơn tháng 4 thực tế (297,0 triệu) cỡ **20%**; cuối tuần cao hơn ngày thường **24%**; dải 95% rộng **±6,6 triệu ngày đầu, ±18,0 triệu trung bình** — rộng hơn cả giá trị dự báo.
  - Kiểm chứng ngược tháng 4: dự báo **280,0 triệu** vs thực **297,0 triệu** (−5,7%).
  - **Giới hạn:** (1) 16 tháng chưa đủ **2 chu kỳ năm**; (2) **Tết đổi ngày theo âm lịch**: 10/02/2024, 29/01/2025, **17/02/2026**; (3) đừng trình bày một con số — luôn **kèm dải**.
  - Dự báo = **mẫu hình quá khứ + giả định**.
- **Ảnh:** [CHỤP] S24a: hộp thoại Create Forecast Worksheet (Options mở rộng); [CHỤP] S24b: biểu đồ dự báo với đường Forecast và hai đường Confidence Bound; `S24a_forecast_dialog.png`, `S24b_forecast_chart.png`.
- **Chân slide:** Ảnh: Demo/B6_Demo_DuBao.xlsx (Data → Forecast Sheet trên tblNgay; hàm ở sheet Demo_ETS)

**Lời giảng:**
> *(→ chọn bảng tblNgay → Data → Forecast → Forecast Sheet → Options)* Forecast Sheet dùng thuật toán **ETS** — làm mượt theo hàm mũ, có tính thành phần mùa. Có ô **Seasonality**, mặc định *Detect Automatically*. *(→ sheet Demo_ETS, ô B6: `=FORECAST.ETS.SEASONALITY(tblNgay[Doanh thu], tblNgay[Ngay])`)* Thử để Excel tự đoán chu kỳ: nó trả **121**. Không phải 7. Nhịp tuần của MCI Mart có thật nhưng nhỏ so với biến động từng ngày — nên Excel không tự nhận ra. Mình biết chu kỳ là *tuần*, nên đặt **Set Manually, 7**. *(→ Forecast End 30/05, Confidence 95% → Create)* Bài học nhỏ: công cụ tự động không thay được hiểu biết về dữ liệu.
>
> *(→ chỉ biểu đồ)* Đường dự báo giờ **có răng cưa**: cuối tuần cao hơn ngày thường khoảng 24% — đúng nhịp mà đường thẳng bỏ qua. Nhưng nhìn hai đường bao quanh: **Lower và Upper Confidence Bound** — dải tin cậy 95%. Dải này **rất rộng**: nửa chiều rộng ngày đầu khoảng 6,6 triệu, và đi xa hơn thì phình ra, trung bình hơn 18 triệu mỗi ngày — **rộng hơn cả con số dự báo**, cận dưới còn âm. Mô hình đang nói thẳng: "tôi không chắc". Và tổng dự báo tháng 5 chỉ 237,7 triệu, thấp hơn tháng 4 thực tế gần hai mươi phần trăm — không có lý do kinh doanh nào cho việc đó.
>
> **Câu hỏi cho lớp: vậy dự báo này có dùng được không?** *(→ chờ)* Mình thử kiểm chứng ngược. *(→ Forecast Start = 01/04/2025 → dự báo tháng 4 từ dữ liệu tới hết tháng 3)* Tổng dự báo 280 triệu, thực tế 297 triệu: lệch 5,7%, và cả 30 ngày thực tế nằm trong dải — nhưng dải rộng nên dĩ nhiên nằm trong. Nghĩa là: dự báo tạm được cho **mức cả tháng**, không đáng tin cho **từng ngày**. Giống Trendline ở Slide 23.
>
> *(→ sheet Demo_ETS: B4 `=FORECAST.ETS(B3, …, 7)` và B5 `=FORECAST.ETS.CONFINT(B3, …, 0.95, 7)`)* Hai hàm tương ứng: một cho điểm dự báo, một cho nửa chiều rộng dải. Dùng khi cần đưa số vào dashboard.
>
> Bây giờ câu tình huống: **"Giám đốc hỏi doanh thu Tết 2026. Dùng được dự báo này không?"** *(→ chờ)* Không. Thứ nhất, mình mới có 16 tháng dữ liệu — chưa đủ hai chu kỳ năm để mô hình học mùa vụ theo năm. Thứ hai, Tết **đổi ngày theo âm lịch**: 10/02 năm 2024, 29/01 năm 2025, 17/02 năm 2026. ETS học nhịp *lặp đều*, không biết sự kiện dịch chuyển. Kết luận: dự báo ngắn hạn, mô hình đơn giản, có dải tin cậy, ghi rõ giả định.

**Ghi chú giảng viên:**
- **Số đo trên Excel for Mac 365** bằng hàm `FORECAST.ETS` (Forecast Sheet dùng cùng thuật toán), xem bảng "Số liệu". **Kiểm lại một lần bằng hộp thoại Forecast Sheet trên Excel Windows** khi chụp ảnh S24a, S24b: tổng 30 ngày tháng 5 ≈ 237,7 triệu; Detect Automatically có thể ra chu kỳ khác 7.
- **Nhịp tuần thật nhưng nhỏ:** cuối tuần +19% so với độ lệch chuẩn doanh thu ngày 3,1 triệu; R² của "thứ trong tuần" chỉ 5,9%–15,3%. `FORECAST.ETS.SEASONALITY` trả 121 (cả chuỗi), 35 (tới 31/03), chỉ ra 7 khi cắt riêng 03–12/2024.
- Số dự báo tháng 5 thấp bất thường (7,9 triệu/ngày, trong khi tháng 4 thực tế 9,9 triệu) là **hiện tượng thật của mô hình** trên dữ liệu có đỉnh Tết, không phải lỗi gõ: dùng làm bằng chứng cho "luôn kèm dải và giả định". Nếu muốn chạy thử một dự báo "đẹp" hơn, cắt dữ liệu từ 01/03/2024 (bỏ Tết 2024) và so sánh.
- Nếu học viên hỏi "sao không chạy theo tháng, mùa 12": đủ 2 chu kỳ năm mới chạy được (xem bảng lỗi thường gặp).

---

## Phần 6 — Build dashboard có hướng dẫn (40 phút · slide 25–30)

### Slide 25 · CHIA PHẦN · 0.5'
- **Trên slide:** MCI / **6. Build: dựng dashboard từng bước**

**Lời giảng:**
> Phần 6, bốn mươi phút, và là phần chính của buổi. Mình làm cùng nhau từng bước: các bạn dựng trên file của mình, giảng viên dựng trên máy chiếu. Cuối phần này mỗi bạn có một dashboard một trang.

### Slide 26 · THỰC HÀNH · Build 1 — Khung, Slicer, tham số · 8'
- **Trên slide:**
  1. Sheet mới **Cai_dat**: Target KPI (Orders 1.000 · AOV 300.000 · GM 33% · YoY +10% · Khách 700) và ngưỡng RAG (100% · 90%).
  2. Sheet mới **Dashboard**: View → bỏ **Gridlines**; đặt độ rộng cột đều; tiêu đề lớn ô A1.
  3. Sheet mới **Data_Dash** (sẽ ẩn ở Slide 33): vùng lấy dữ liệu cho thẻ KPI, waterfall, map.
  4. Pivot trống từ Data Model (sheet tạm) → chèn **Slicer Tháng** (Calendar[Thang]) và **Slicer Miền** (Store[Region]); đặt Slicer lên `Dashboard`; chọn tháng **4**.
  5. Đặt tên trong **Slicer Settings**: `Slicer_Thang`, `Slicer_Region`.
- **Kiểm tra:** Slicer Tháng đang chọn 4 · `Cai_dat` có 5 target + 2 ngưỡng.

**Lời giảng (giao bài):**
> *(→ giảng viên làm trước 3 phút, lớp làm theo 5 phút)* Bước một: khung. Ba sheet: **Cai_dat** giữ target và ngưỡng RAG — mọi hằng số nằm một chỗ, đổi một lần. **Dashboard** là trang người xem thấy; mình tắt đường lưới cho giống một bảng báo cáo. **Data_Dash** là phòng máy: nơi CUBEVALUE lấy số. Sau đó hai Slicer: Tháng và Miền. Slicer Tháng mình chọn tháng 4 làm mặc định.
>
> **Điểm hay vấp:** Slicer cần một Pivot làm gốc. Các bạn tạo một Pivot trống từ Data Model ở sheet tạm, chèn Slicer rồi dọn sau. *(→ đi vòng lớp, kiểm tra tên Slicer)*

**Ghi chú giảng viên:**
- Sheet `Cai_dat` mẫu: B3:C8 target; C13 = 100%, C14 = 90% (khớp công thức ở bảng "Công thức").
- Slicer hiển thị tháng 1–12 vì Calendar có đủ 2025; tháng 5–12 sẽ trống — Slicer Settings → **Hide items with no data**.

### Slide 27 · THỰC HÀNH · Build 2 — Hàng KPI card (CUBEVALUE + RAG + sparkline) · 10'
- **Trên slide:**
  1. `Data_Dash`: 6 KPI × (giá trị · target · mức đạt · nhãn) bằng **CUBEVALUE** (Slicer_Thang, Slicer_Region) + hàng T1–T4 cho sparkline.
  2. `Dashboard`: 6 thẻ (ô gộp, viền hairline, chữ lớn 28); thẻ **Doanh thu** ở góc trên trái.
  3. Mỗi thẻ: tên KPI · **giá trị** · "Target …" · mức đạt + **biểu tượng + nhãn** (Slide 19) · sparkline (Slide 21).
  4. Bấm Slicer sang tháng 3, tháng 2 và quay lại tháng 4 để kiểm tra.
- **Kiểm tra — tháng 4/2025:** Doanh thu **296.965.000** (103,1%) ✔ · Đơn **966** (96,6%) ▲ · AOV **307.417** (102,5%) ✔ · GM **33,8%** ✔ · YoY **+24,0%** ✔ · Khách **670** (95,7%) ▲. Tháng 2: Doanh thu 82,5% ▲, YoY −0,9% ✘.

**Lời giảng (giao bài):**
> *(→ giảng viên làm thẻ Doanh thu, lớp làm các thẻ còn lại)* Bước hai là bước nhiều công việc nhất. Sáu thẻ, mỗi thẻ có bốn thứ: giá trị, target, mức đạt, nhãn trạng thái. Các bạn làm thẻ **Doanh thu** cùng mình, rồi nhân bản công thức sang năm thẻ còn lại — chỉ đổi tên measure. Target lấy từ `Cai_dat`. Thẻ Doanh thu dùng measure `Total Target` vì target doanh thu đến từ bảng Target.
>
> Kiểm tra khi xong: tháng 4, sáu thẻ phải ra đúng bảng trên slide. *(→ bấm Slicer sang tháng 2)* Tháng 2: doanh thu 82,5% vàng, **YoY đỏ có dấu chéo**. Nếu thẻ YoY ra trống: bạn đã chọn một tháng mà năm trước không có dữ liệu, hoặc measure `Revenue YoY %` đang bị chặn bởi bộ lọc `Nam`; xem ghi chú.

**Ghi chú giảng viên:**
- **Khách (Active Customers) ở tháng 4: 670**, target 700 → 95,7% ▲. Không cộng tay Active Customers theo miền (234 + 151 + 328 = **713** ≠ 670): đây là bẫy của B5 quay lại; nhắc để học viên đọc ô tổng của measure.
- Thẻ YoY: filter `Nam = 2025` kèm Slicer_Thang; measure `Revenue YoY %` lấy cùng kỳ qua `DATEADD`; ô trống nếu không có `Revenue LY`.
- Nếu máy chậm, rút gọn sparkline (làm ở nhà) — thẻ vẫn đủ giá trị, target, RAG.

### Slide 28 · THỰC HÀNH · Build 3 — Line theo ngày và Combo Actual vs Target · 8'
- **Trên slide:**
  1. **PivotChart Line:** Pivot từ Data Model → Filters `Calendar[Nam]` = 2025 · Rows `Calendar[Date]` · Values `Total Revenue` → Insert PivotChart → Line. Nối Slicer Tháng và Miền (Report Connections).
  2. **PivotChart Combo:** Rows `Calendar[Thang]` · Values `Total Revenue` và `Total Target` → Combo: Revenue = Clustered Column, Target = Line **nét đứt, xám** (series-reference). Rows: **Value Filters → Greater Than → Total Revenue > 0** (chỉ giữ tháng đã có doanh thu). Nối **chỉ Slicer Miền**.
  3. Tiêu đề chart là **kết luận**; thêm **chú thích** chỉ vào điểm bất thường.
- **Kiểm tra:** Combo: T1 120,6% · T2 82,5% · T3 94,6% · T4 103,1% (target 353 · 298 · 285 · 288 triệu). Line tháng 4: 30 ngày.

**Lời giảng (giao bài):**
> Hai chart xu hướng. *(→ làm cùng)* Line theo ngày: Pivot với Calendar[Date] ở hàng. Nếu để tất cả ngày thì ra 730 điểm; Slicer Tháng sẽ cắt về một tháng. Tiêu đề chart: không phải "Doanh thu theo ngày" mà là một **kết luận**: ví dụ "Doanh thu tháng 4 cao nhất vào cuối tuần".
>
> Combo thực tế vs target. Cột là Total Revenue, đường là Total Target — **xám, nét đứt**: mốc so sánh không bao giờ dùng màu của series. Lọc hàng để chỉ giữ tháng có doanh thu: dùng **Value Filter** thay vì chọn tay tháng 1–4, vì tháng 5 có dữ liệu thì chart tự thêm, không phải sửa — đó là tư duy **tự động hóa** D3.
>
> *(→ chỉ combo)* Đây là chart **nhìn thấy câu chuyện số 4**: tháng 1 vượt target 21%, tháng 2 hụt 17%, gộp hai tháng 103,2%. Tết 2025 rơi vào 29/01 trong khi target lập theo cùng tháng 2024, nơi Tết rơi vào 10/02. **Câu hỏi cho lớp: nên thêm chú thích gì lên chart này?** *(→ chờ)* Một ô chữ nhỏ chỉ vào hai cột: "Tết 29/01 dồn mua sắm sang T1; target lập theo Tết 10/02".

### Slide 29 · THỰC HÀNH · Build 4 — Top 5 sản phẩm và Waterfall theo cửa hàng · 8'
- **Trên slide:**
  1. **Top 5:** Pivot Rows `Product[Product_Name]` · Values `Total Revenue` → Row Labels → Value Filters → **Top 10** → đổi thành **5** → PivotChart Bar (sắp từ lớn đến nhỏ, trục từ 0). Nối cả hai Slicer.
  2. **Waterfall:** `Data_Dash` A20:B33 — dòng đầu **Target** (`Total Target`), 12 dòng **Variance** của từng cửa hàng (`CUBEMEMBER` + `CUBEVALUE`), dòng cuối **Actual** (`Total Revenue`) → Insert → Waterfall; **Set as Total** cho dòng đầu và dòng cuối.
  3. **Map hoặc Bar:** `Data_Dash` — doanh thu từng thành phố (CUBEVALUE, Slicer_Thang) → Insert → Maps → **Filled Map**; không có Internet → Bar.
- **Kiểm tra — tháng 4:** Top 5: Hộp quà Tết cao cấp **35.100.000** · Chảo chống dính 26cm **29.155.000** · Dầu gội 650ml **26.565.000** · Mì tôm chua cay **19.712.000** · Bột giặt 3kg **19.276.000**. Waterfall tháng 4: Target 288.000.000 → Actual **296.965.000** (tổng +8.965.000). Tháng 3: ST07 **−10.461.000**.

**Lời giảng (giao bài):**
> *(→ làm cùng)* **Top 5**: một Pivot, rồi Value Filter Top 10 đổi thành 5. Năm sản phẩm đầu chiếm 43,7% doanh thu tháng 4. Hộp quà Tết cao cấp vẫn đứng đầu trong tháng 4 — **câu hỏi cho lớp: vì sao hộp quà Tết vẫn bán chạy nhất sau Tết?** *(→ chờ; đây là câu "vì sao" nảy sinh, ghi lên bảng — mình không có đủ dữ liệu để trả lời hôm nay, và đó là một điều đáng nói trong brief: **câu hỏi dữ liệu hiện có chưa trả lời được**)*
>
> **Waterfall**: không có Pivot, nên ta lấp vùng ô bằng CUBEVALUE. Cột A là tên cửa hàng viết bằng `CUBEMEMBER`, cột B là `CUBEVALUE` đo measure **Variance** cho từng cửa hàng và đọc Slicer Tháng. Dòng đầu là Target, dòng cuối là Actual. *(→ chọn vùng → Insert → Waterfall → chuột phải cột đầu: Set as Total, cột cuối: Set as Total)* Bấm Slicer sang **tháng 3**: ST07 trừ 10,5 triệu — **câu chuyện số 2** hiện ra mà không cần chú thích; và ST12 trừ 11,7 triệu — một cái tên mà ta sẽ quay lại ở phần kể chuyện.
>
> **Map**: cùng cách. Bing nhận tên thành phố; nếu máy lớp không ra, chuyển sang bar theo thành phố. TP.HCM 108,9 triệu đứng đầu.

**Ghi chú giảng viên:**
- Hộp quà Tết tháng 4: dữ liệu giả lập, Top 5 theo doanh thu — không cố giải thích. Dùng làm ví dụ "câu hỏi mới nảy ra từ dashboard".
- Map: [CẦN KIỂM TRA TRÊN EXCEL 365 WINDOWS] Bing có nhận đủ 9 tên (Hà Nội, TP.HCM, Hải Phòng, Quảng Ninh, Đà Nẵng, Huế, Khánh Hòa, Cần Thơ, Đồng Nai) và đúng vị trí không (một số tỉnh đã sáp nhập từ 07/2025); nếu không, dùng bar. Map chart không chạy trên Excel for the web / Mac cũ.
- Variance tính qua measure có chặn ô trống (B5): cửa hàng chưa có doanh thu trả `BLANK`, waterfall bỏ qua.

### Slide 30 · THỰC HÀNH · Build 5 — Ráp trang, kiểm tra 5 giây · 5.5'
- **Trên slide:**
  1. Sắp 7 visual vào bố cục Slide 17; căn lề bằng **Align**; màu theo thực thể (Bắc xanh dương, Trung cam, Nam xanh lá); mốc target xám nét đứt.
  2. **Report Connections** (Slide 20): Slicer Tháng nối mọi Pivot **trừ combo**; Slicer Miền nối **tất cả**.
  3. **Alt text** cho từng chart; tiêu đề chart là kết luận.
  4. **Kiểm tra 5 giây:** đổi cặp, người kia nhìn 5 giây rồi nói: KPI chính là gì · đạt hay chưa · nhìn đâu tiếp.
- **Kiểm tra:** bấm Slicer Tháng 4 → 3 → 2: thẻ, line, waterfall, Top 5, map đều đổi; combo không đổi. ≤ 8 visual.
- **Ảnh:** [CHỤP] S30: dashboard hoàn chỉnh, tháng 4/2025 (toàn trang); `S30_dashboard_t4.png`.
- **Chân slide:** Ảnh: Practice/B6_MCI_Mart_Dashboard_Answer.xlsx, sheet Dashboard (tháng 4/2025)

**Lời giảng (giao bài):**
> *(→ giảng viên chiếu `Dashboard` của mình, tháng 4)* Ráp trang. Căn lề cho đều: Home → Format → Align. Màu: Bắc xanh dương, Trung cam, Nam xanh lá, giống mọi chart khác của khóa; target luôn xám nét đứt. Và chạy **Report Connections** lần cuối: Slicer Tháng nối hết, trừ combo.
>
> **Bài kiểm tra năm giây:** đổi cặp. Người xem nhìn đúng năm giây rồi **quay đi** và nói ba điều: KPI chính là gì, đạt hay chưa đạt, nhìn đâu tiếp theo. Nếu bạn ngồi cạnh không nói được "doanh thu 297 triệu, đạt 103% target, ổn" thì dashboard chưa xong. *(→ cho lớp 2 phút)*

**Ghi chú giảng viên:**
- Lớp chưa làm kịp: phát `Practice/B6_MCI_Mart_Dashboard_Answer.xlsx` để vẫn tham gia chấm chéo ở Slide 34.
- Chỗ dễ lố giờ: Slide 27 (6 thẻ) và Slide 29 (3 chart). Lớp chậm: bỏ sparkline, bỏ map (thay bằng bar), làm các phần đó ở nhà.

---

## Phần 7 — Kể chuyện và phát hành (15 phút · slide 31–34)

### Slide 31 · CHIA PHẦN · 0.5'
- **Trên slide:** MCI / **7. Kể chuyện (SCR) & phát hành dashboard**

**Lời giảng:**
> Phần 7: dashboard đúng và đẹp chưa đủ — người xem phải biết *phải làm gì*. Rồi mình đóng gói để gửi đi.

### Slide 32 · KHÁI NIỆM · Storytelling SCR · chú thích trên chart · 5.5'
- **Trên slide:** **Situation → Complication → Resolution**
  - **S — Situation:** con số **có mốc so sánh**.
  - **C — Complication:** vấn đề và nguyên nhân **đã kiểm chứng bằng drill-down**.
  - **R — Resolution:** đề xuất hành động **cho đúng người xem trong brief**.
  - **Mẫu (tháng 3/2025, Sales Director):**
    - **S:** Tháng 3 doanh thu **269.608.000**, đạt **94,6%** target (hụt **15.392.000**), nhưng tăng **13,7%** so với cùng kỳ.
    - **C:** Riêng **ST07 Nha Trang** hụt **10.461.000** (đạt **44,9%**), chiếm **68%** khoản hụt, do cải tạo 10–31/03. **ST12 Biên Hòa** dưới 90% target ba tháng liền (**76,9% · 53,0% · 65,2%**).
    - **R:** (1) Đánh giá ST07 theo số ngày mở cửa, không phạt target tháng 3; (2) rà soát cách lập target ST12 (đang lấy trung bình T9–T12/2024); (3) lập **target 2026 theo lịch Tết âm** (Tết 17/02/2026): đánh giá T1 và T2 gộp.
  - **Chú thích trực tiếp trên chart** chỉ vào điểm bất thường (ST07, T2).
- **Ảnh:** [VẼ] ba khối S–C–R nối nhau.

**Lời giảng:**
> Dashboard cho người ta **thấy**. Kể chuyện cho người ta **hành động**. Khung ba bước: **Situation** — con số có mốc so sánh. **Complication** — vấn đề, và nguyên nhân *đã kiểm chứng* bằng drill-down, không phải đoán. **Resolution** — đề xuất cụ thể, gửi *đúng người xem trong brief*.
>
> *(→ bấm Slicer sang tháng 3, chiếu dashboard)* Mẫu cho Giám đốc bán hàng. **Situation**: tháng 3, 269,6 triệu, đạt 94,6% target, hụt 15,4 triệu — nhưng tăng 13,7% so với tháng 3 năm ngoái. **Complication**: nhìn waterfall, riêng ST07 Nha Trang hụt 10,5 triệu, **chiếm hơn hai phần ba** khoản hụt; cửa hàng này cải tạo từ 10 đến 31 tháng 3. Kiểm chứng bằng drill-down: Pivot ST07 theo ngày. Và còn một cái tên khác: **ST12** Biên Hòa dưới 90% ba tháng liền. **Resolution**: ba đề xuất, mỗi đề xuất một việc cho Giám đốc. Đặc biệt đề xuất cuối: *lập target 2026 theo lịch Tết âm* — vì Tết 2026 rơi vào 17/02, lệch ngày 12 ngày so với 2025.
>
> **Câu hỏi cho lớp: tại sao đề xuất 2 nói "rà soát cách lập target" thay vì "ST12 làm việc kém"?** *(→ chờ)* Vì ta chưa kiểm chứng: ST12 mới khai trương từ tháng 6/2024, target của nó lấy trung bình T9–T12/2024 — giai đoạn cuối năm mùa cao. Đó là **giả thuyết cần kiểm chứng**, không phải kết luận. Phân biệt giữa "điều dữ liệu cho thấy" và "điều ta đoán" là kỹ năng C6 mức Level 3 trở lên.

**Ghi chú giảng viên:**
- Câu mẫu dùng số của bảng "Số liệu"; mọi con số lấy từ measure, đối soát bằng Pivot. ST12: T2–T4 đều dưới 90% (76,9% · 53,0% · 65,2%), lũy kế 79,4%.
- ST07 chiếm 68% khoản hụt T3: −10.461.000 / −15.392.000. Các cửa hàng khác bù lẫn nhau (ST08 +10,7 triệu, ST10 +9,4 triệu).
- Câu chuyện #4 (Tết): đề xuất 3 là hành động Tầng 4 (Course Plan, Phụ lục F).

### Slide 33 · CÔNG CỤ · Phát hành: Refresh on open · ẩn sheet · Protect · PDF · chia sẻ · 6'
- **Tên (đỏ):** **Query Properties → Refresh data when opening the file** · **Hide Sheet** · **Review → Protect Sheet** · **File → Export → PDF** · **Share**
- **Trên slide:**
  1. **Refresh on open:** Data → Queries & Connections → chuột phải query → Properties → tick *Refresh data when opening the file*. (Không thử trên lớp.)
  2. **Ẩn sheet kỹ thuật:** `Data_Dash`, `Cai_dat`, các sheet Pivot.
  3. **Protect Sheet** cho `Dashboard`; bỏ khóa Slicer (Properties → bỏ *Locked*) và tick *Use PivotTable & PivotChart* để Slicer vẫn bấm được.
  4. **Xuất PDF:** Page Layout → Landscape · **Fit to 1 page** → File → Export → Create PDF.
  5. **Chia sẻ:** OneDrive / SharePoint → Share; lưu ý dữ liệu khách hàng và **sensitivity label** (Home → Sensitivity).
  6. **Giới hạn:** Excel **không tự refresh theo lịch** → cần Power BI (hoặc Power Automate / Office Scripts).
- **Ảnh:** [CHỤP] S33: hộp Protect Sheet với tùy chọn Use PivotTable & PivotChart; `S33_protect_sheet.png`.
- **Chân slide:** Ảnh: B6_<HoTen>.xlsx, Review → Protect Sheet (trên sheet Dashboard)

**Lời giảng:**
> Dashboard xong, chưa gửi được ngay. Sáu việc. **Một**: để file tự cập nhật khi mở, bật *Refresh data when opening the file* trong Properties của query. Hôm nay mình chỉ bật chứ không thử, vì đường dẫn của file khởi đầu chưa trỏ đúng.
>
> **Hai**: **ẩn các sheet kỹ thuật** — Data_Dash, Cai_dat, các sheet Pivot — để người xem chỉ thấy Dashboard. **Ba**: **Protect Sheet** cho khỏi bị kéo lệch, nhưng chú ý: Slicer phải bỏ khóa và tick "Use PivotTable & PivotChart", nếu không người xem không bấm được. *(→ demo)* **Bốn**: xuất **PDF** — nhớ Fit to 1 page, nằm ngang. **Năm**: chia sẻ qua OneDrive hoặc SharePoint; dashboard này có dữ liệu khách hàng nên chú ý quyền xem và **sensitivity label** nếu công ty bạn có.
>
> **Sáu — giới hạn, và phải nói thẳng**: Excel **không tự refresh theo lịch**. Mở file thì refresh, nhưng không ai mở thì dữ liệu cứ cũ. Muốn dashboard cập nhật mỗi sáng 7 giờ, bạn cần Power BI. Đây không phải "Excel kém", mà là biết khi nào nên chuyển công cụ.

**Ghi chú giảng viên:**
- Chặn sửa Data Model: Protect Workbook (Structure) giữ nguyên số sheet; ẩn sheet bằng **Hide** là đủ cho bài này (không dùng Very Hidden).
- Cảnh báo Excel for the web: Data Model và Pivot từ Data Model hạn chế; **[CẦN KIỂM TRA TRONG MÔI TRƯỜNG CỦA TỔ CHỨC]** trước khi hứa người xem mở được dashboard trên trình duyệt.

### Slide 34 · THỰC HÀNH · Chấm chéo bằng checklist · 3'
- **Trên slide:** Đổi máy với bạn bên cạnh. Tick từng dòng (Phụ lục D):
  - [ ] Mỗi visual trả lời một câu trong brief
  - [ ] KPI chính ở góc trên trái; thứ tự KPI → Trend → Detail
  - [ ] 5–8 visual
  - [ ] Mọi KPI có mốc so sánh và RAG **kèm biểu tượng hoặc nhãn**
  - [ ] Mọi con số từ measure (CUBEVALUE / Pivot); không công thức KPI ở tầng hiển thị
  - [ ] Trục bar/column bắt đầu từ 0; không 3D; không pie quá 5 phần
  - [ ] Slicer điều khiển mọi visual (trừ visual ngữ cảnh đã ghi chú)
  - [ ] Có chú thích chỉ vào điểm bất thường
  - [ ] Bài kiểm tra 5 giây
  - [ ] Phát hành: Refresh on open, ẩn sheet kỹ thuật, Protect Sheet, alt text, xuất PDF thử
  - [ ] Nếu có dự báo: dải tin cậy và ghi rõ giả định
- **Kiểm tra:** ghi 1 điểm làm tốt và 1 điểm cần sửa cho dashboard của bạn.

**Lời giảng (giao bài):**
> *(→ đổi máy 3 phút)* Mười một dòng, các bạn chấm dashboard của người bên cạnh. **Chỉ chấm, không sửa.** Cuối cùng ghi một điều làm tốt và một điều cần sửa. Checklist này cũng là tiêu chí C5 của bài capstone: các bạn tự kiểm tra bài cuối khóa bằng chính tờ này. *(→ gọi 1–2 bạn nêu một điều họ học được từ dashboard của bạn mình)*

---

## Phần 8 — Quiz, insight, giao capstone (10 phút · slide 35–41)

Quiz kiểu hỏi nhanh, khoảng 35 giây mỗi câu: **4 câu, trong đó 2 câu tình huống (câu 3–4).**

### Slide 35 · QUIZ · EXCEL QUIZ CHALLENGE · 0.2'

**Lời giảng:**
> Bốn câu hỏi nhanh. Câu 3 và 4 là hai tình huống các bạn sẽ gặp ngoài đời.

### Slide 36 · QUESTION 1: CHỌN CHART
- Muốn đặt doanh thu (VNĐ) và biên lợi nhuận gộp (%) của sáu ngành lên cùng một biểu đồ, nên dùng gì? A) Pie · **B) Combo: cột doanh thu, đường biên % ở trục phụ** · C) Scatter · D) Waterfall

**Lời giảng:**
> B. Hai thước đo khác đơn vị: combo hai trục. Pie là cơ cấu; scatter là quan hệ hai biến số; waterfall là chênh lệch cộng dồn.

### Slide 37 · QUESTION 2: SLICER
- Dashboard có ba PivotChart. Bạn bấm Slicer Miền, chỉ một chart đổi. Nguyên nhân và cách xử lý? A) Phải Refresh All · **B) Slicer chưa nối với các Pivot còn lại: chuột phải Slicer → Report Connections, tick đủ** · C) Phải tạo Slicer riêng cho từng chart · D) Đổi Slicer thành Timeline

**Lời giảng:**
> B. Slicer mặc định chỉ điều khiển Pivot mà nó sinh ra. Report Connections cho nối thêm. Với thẻ KPI bằng CUBEVALUE thì Slicer là đối số của công thức.

### Slide 38 · QUESTION 3: TÌNH HUỐNG — DỰ BÁO TẾT
- Giám đốc xem Forecast Sheet và hỏi: "Doanh thu Tết 2026 là bao nhiêu?" Câu trả lời đúng? A) Đọc số trên đường Forecast, chắc chắn đúng · B) Chạy lại với dải tin cậy 99% · **C) Không nên dùng: mới có 16 tháng dữ liệu chưa đủ hai chu kỳ năm, Tết đổi ngày theo âm lịch (17/02/2026), nên dự báo ngắn hạn kèm dải tin cậy và giả định** · D) Dùng Linear Trendline cho chắc

**Lời giảng:**
> C. ETS học nhịp lặp đều như cuối tuần; nó không biết sự kiện dịch chuyển như Tết. Và một con số dự báo không kèm dải tin cậy dễ bị hiểu là cam kết.

### Slide 39 · QUESTION 4: TÌNH HUỐNG — TRỤC Y
- Bar chart so doanh thu tháng 3 (237.192.000) và tháng 4 (239.466.000) năm 2024: cột tháng 4 cao gần gấp ba. Vì sao và sửa thế nào? A) Doanh thu thật sự tăng gấp ba · **B) Trục Y bắt đầu từ 236.000.000; đặt trục bắt đầu từ 0 (chênh thật chỉ 0,96%)** · C) Đổi sang pie 3D · D) Thêm đường lưới

**Lời giảng:**
> B. Trục cắt khúc phóng đại chênh lệch nhỏ. Bar và column luôn bắt đầu từ 0. Line chart mới được cắt trục.

### Slide 40 · TỔNG KẾT + INSIGHT · 2.5'
- **Trên slide:**
  - **Điểm chính cần nhớ:**
    - (1) **Brief trước, chart sau:** người xem, 3–5 câu hỏi đo lường được, tần suất, quyết định.
    - (2) **Thứ bậc** KPI → Trend → Detail; 5–8 visual; mỗi visual trả lời một câu.
    - (3) **Một nguồn sự thật:** mọi số từ measure (CUBEVALUE hoặc Pivot); KPI nào cũng có **mốc so sánh** và RAG có biểu tượng.
    - (4) **Insight theo SCR** có drill-down kiểm chứng và đề xuất cho đúng người.
    - (5) **Dự báo = mẫu hình quá khứ + giả định:** luôn kèm dải tin cậy.
  - **Insight của bạn (3 câu SCR):** "S: ___ (con số có mốc so sánh). C: ___ (nguyên nhân đã kiểm chứng). R: ___ (đề xuất cho Sales Director)."
- **Ảnh:** [VẼ] khung insight.

**Lời giảng:**
> Năm điều cần nhớ *(→ đọc lướt 5 ý trên slide)*. Viết câu insight theo SCR trong hai phút, cho **tháng bạn chọn** — không nhất thiết tháng 3. *(→ chờ, gọi 1–2 bạn đọc; hỏi "vậy thì sao?" nếu câu dừng ở mô tả)* Câu hỏi "vậy thì sao" là câu hỏi chấm điểm C6: nguyên nhân chưa kiểm chứng thì chỉ ở Level 2.

### Slide 41 · GIAO CAPSTONE & BÀI VỀ NHÀ · 5'
- **Trên slide:**
  - *Buổi sau (B7):* Ôn tập, hoàn thiện dự án, **thuyết trình 5–7 phút**, nhận xét theo rubric.
  - **Giao hôm nay — Capstone (cá nhân):** brief dự án + dataset mới + rubric bản học viên (`../B7_Capstone/Rubric_B7_Capstone.md`). Chuỗi: Brief → Power Query → Data Model → KPI measure → Dashboard → Insight.
  - **Bài tập về nhà B6 (bắt buộc):** dashboard một trang cho **Marketing Director** (Active Customers, phân khúc) trên file khởi đầu. Tập dượt **C1 · C5 · C6**. Nộp **trước Buổi 7** tại hocvien.mcivietnam.com.
  - Đọc handout B6 · tick "Checklist trước Buổi 7".
- **Ảnh:** [VẼ] viên gạch B6 (đã xong) → viên gạch B7 (sáng lên).

**Lời giảng:**
> Hôm nay các bạn đã đi hết chuỗi: brief, mô hình, KPI, dashboard, câu chuyện. Năng lực D1 — bắt đầu từ câu hỏi — và D4 — truyền đạt insight — giờ ở mức Competent.
>
> Hai việc giao hôm nay. **Một: dự án capstone**, làm cá nhân, ở nhà, trên bộ dữ liệu mới. Các bạn nhận brief, dataset và rubric bản học viên. Rubric có sáu tiêu chí; **tiêu chí C5 dashboard và C2 dữ liệu là điều kiện bắt buộc** để đạt Level 3 — Builder. Hãy đọc rubric trước khi bắt đầu, không phải sau. **Hai: bài tập về nhà** — một dashboard một trang cho Giám đốc Marketing, khác người xem với hôm nay. Đây là phần tập dượt của C1, C5 và C6. Nộp trước Buổi 7.
>
> Buổi 7 là buổi thuyết trình: mỗi bạn năm đến bảy phút, và người nghe sẽ hỏi "vậy thì sao?". Hẹn gặp lại các bạn.

---

## Phụ lục kịch bản

### Kiểm tra thời lượng

| Phần | Slide | Phút | Timetable B6 |
|---|---|---|---|
| 0 Mở đầu, recap | 1–3 | 5 | 5 |
| 1 Dashboard brief | 4–6 | 10 | 10 |
| 2 Clarity & simplicity, Good/Bad/Ugly | 7–9 | 10 | 10 |
| 3 3 câu hỏi + 8 chart lõi | 10–14 | 20 | 20 |
| Nghỉ | 15 | 10 | 10 |
| 4 Layout, CUBEVALUE, RAG, Slicer, Sparkline, Alt text | 16–21 | 20 | 20 |
| 5 Dự báo nhập môn | 22–24 | 10 | 10 |
| 6 Build | 25–30 | 40 | 40 |
| 7 Storytelling, phát hành, chấm chéo | 31–34 | 15 | 15 |
| 8 Quiz, insight, giao capstone | 35–41 | 10 | 10 |
| **Tổng** | **41** | **150** | **150** |

Phần 8 cộng từng slide đúng 10 phút (slide 35: 0,2'; 4 câu × khoảng 35 giây ≈ 2,3'; slide 40: 2,5'; slide 41: 5'), không có dự phòng.

**Chỗ dễ lố giờ:** Slide 27 (6 thẻ KPI), Slide 29 (3 chart), Slide 24 (hộp thoại Forecast). Lớp chậm: bỏ histogram ở Slide 14; bỏ sparkline và map (thay bằng bar) ở Slide 27, 29; Slide 24 chỉ làm Forecast Sheet, bỏ hai hàm `FORECAST.ETS`.

### Lỗi thường gặp → cách xử lý (Course Plan, Hướng dẫn giảng dạy B6)

| Lỗi | Dấu hiệu | Xử lý | Slide |
|---|---|---|---|
| Quá nhiều visual | Người xem không biết nhìn đâu | Tối đa 5–8; bỏ visual không trả lời câu hỏi nào trong brief | 8, 17 |
| Bar chart có trục Y không bắt đầu từ 0 | Chênh lệch nhỏ trông rất lớn (T3 vs T4/2024: 0,96% → cao gấp 2,9) | Trục bar/column bắt đầu từ 0 | 9, 39 |
| Slicer không nối với mọi Pivot | Một phần dashboard không đổi khi lọc | Report Connections | 20, 37 |
| Pivot co giãn cột khi Refresh, làm vỡ bố cục | Layout xô lệch | Tắt "Autofit column widths on update" | 20 |
| Gõ sai tên measure/member trong CUBEVALUE | `#N/A` | Dùng gợi ý autocomplete hoặc Convert to Formulas | 18 |
| RAG chỉ dùng đỏ-xanh | Người mù màu không đọc được | Thêm biểu tượng và nhãn chữ | 19 |
| Dùng Trendline tuyến tính cho dữ liệu có mùa vụ | Dự báo phẳng, bỏ qua đỉnh và đáy | Forecast Sheet (ETS) với seasonality phù hợp | 23, 24 |
| Chạy ETS theo tháng với seasonality 12 khi chưa đủ 2 năm | Mùa vụ đoán sai, dải tin cậy rất rộng | Dự báo theo ngày (seasonality 7), hoặc chờ đủ dữ liệu | 24 |
| Chuỗi ngày bị thiếu hoặc trùng | Forecast Sheet báo lỗi hoặc tự nội suy âm thầm | Tạo bảng Ngày–Doanh thu bằng Group By trong Power Query, đối chiếu với Calendar | 23 |
| Trình bày một con số dự báo, không có khoảng | Người xem hiểu là cam kết | Luôn kèm dải tin cậy và giả định | 24, 38 |
| Slicer không tạo được | Nút Insert Slicer mờ | Tạo Pivot từ Data Model trước, rồi chèn Slicer | 18, 26 |
| Slicer bị khóa sau Protect Sheet | Người xem không bấm được | Bỏ *Locked* của Slicer; tick *Use PivotTable & PivotChart* | 33 |
| Cộng tay Active Customers theo miền | 713 thay vì 670 | Đọc ô tổng của measure (DISTINCTCOUNT không cộng dồn, B5) | 27 |
| Lỡ Refresh All | Báo lỗi đường dẫn `DataFolder` | Không Refresh trong buổi; dữ liệu đã nằm sẵn | 3, 33 |

### Danh sách ảnh cần chụp [CHỤP] (Excel 365 Windows, zoom 100%)

**Chưa chụp.** Lưu vào `Materials/B6/Images/` đúng tên ở cột Tên file (danh sách đầy đủ: `Materials/B6/Images/_DANH_SACH_ANH.md`). Chụp trên file Demo và trên `Practice/B6_MCI_Mart_Dashboard_Answer.xlsx` khi đã dựng (các ảnh dashboard S18–S21, S30, S33); chụp xong đóng file **không lưu**.

| Tên file | Nội dung cần chụp | File · sheet · vùng | Slide |
|---|---|---|---|
| S08_dashboard_xau.png | Dashboard xấu: 12 visual, pie 3D, trục cắt khúc, không target (toàn trang) | `Demo/B6_Demo_DashboardXau.xlsx` · Dashboard_Xau | 8 |
| S09_truc_y.png | Hai biểu đồ cột T3/2024 và T4/2024 cạnh nhau: trục từ 236.000.000 và từ 0 | `Demo/B6_Demo_DashboardXau.xlsx` · Demo_TrucY | 9 |
| S12_combo_nganh.png | Combo: cột Doanh thu + đường GM% (trục phụ) theo 6 ngành | `Demo/B6_Demo_Chart.xlsx` · Demo_Combo · A3:C9 | 12 |
| S13_waterfall.png | Waterfall từ Target 285.000.000 qua 12 cửa hàng đến Actual 269.608.000, hai cột Total | `Demo/B6_Demo_Chart.xlsx` · Demo_Waterfall · A3:B17 | 13 |
| S14a_histogram.png | Histogram giá trị đơn Q1, bin 100.000, đuôi phải | `Demo/B6_Demo_Chart.xlsx` · Demo_Phan_phoi | 14 |
| S14b_scatter.png | Scatter 12 cửa hàng (tuổi × doanh thu) + trendline + R², ô CORREL = 0,78 | `Demo/B6_Demo_Chart.xlsx` · Demo_Scatter | 14 |
| S18a_cubevalue.png | Ô CUBEVALUE trên thanh công thức, gợi ý autocomplete đang hiện | `B6_<HoTen>.xlsx` · Data_Dash · B3 | 18 |
| S18b_data_dash.png | Vùng `Data_Dash` 6 KPI: giá trị · target · mức đạt (tháng 4) | Như trên · A3:E9 | 18 |
| S19_rag.png | 6 thẻ KPI tháng 4/2025 với biểu tượng ✔ ▲ ✘ và nhãn chữ | `Practice/B6_MCI_Mart_Dashboard_Answer.xlsx` · Dashboard | 19 |
| S20_report_connections.png | Hộp Report Connections của Slicer Tháng: tick các Pivot, bỏ Pivot combo | `Practice/…` · chuột phải Slicer | 20 |
| S21_sparkline.png | Thẻ Doanh thu có sparkline T1–T4 | `Practice/…` · Dashboard | 21 |
| S23_trendline.png | Line chart doanh thu ngày + Linear Trendline + Forward 30 + R² 0,045 | `Demo/B6_Demo_DuBao.xlsx` · Demo_Trendline | 23 |
| S24a_forecast_dialog.png | Hộp Create Forecast Worksheet với Options mở rộng (Seasonality = 7, 95%) | `Demo/B6_Demo_DuBao.xlsx` · Data → Forecast Sheet | 24 |
| S24b_forecast_chart.png | Biểu đồ dự báo 30 ngày, đường Forecast và hai đường Confidence Bound | `Demo/B6_Demo_DuBao.xlsx` · Demo_ETS | 24 |
| S30_dashboard_t4.png | Dashboard hoàn chỉnh, tháng 4/2025, toàn trang | `Practice/B6_MCI_Mart_Dashboard_Answer.xlsx` · Dashboard | 30 |
| S33_protect_sheet.png | Hộp Protect Sheet, tick Use PivotTable & PivotChart | `Practice/…` · Review → Protect Sheet | 33 |

### File Demo — đã dựng 03/10/2026

Theo `CLAUDE.md`: mỗi file có sheet `Huong_dan` (slide · thao tác · kết quả mong đợi), bản `_HocVien` (Huong_dan đứng đầu, "DEMO" → "VÍ DỤ", sheet `Demo_*` → `Vi_du_*`), số liệu tính từ dataset. Dựng bằng `2_Dataset/_generator/build_b6_demos.py` (nạp dữ liệu qua `b6_data.py`); **không cần Data Model**, chạy trên Mac bằng openpyxl. Chart của file Chart và DuBao **không dựng sẵn** (giảng viên vẽ trước lớp — đó là nội dung slide); chỉ `Dashboard_Xau` và `Demo_TrucY` có chart dựng sẵn.

| File | Nội dung | Slide | Kết quả mong đợi (đã kiểm bằng Excel for Mac) |
|---|---|---|---|
| `Demo/B6_Demo_DashboardXau.xlsx` | `Dashboard_Xau`: 12 chart cố ý xấu (pie 3D, bar 3D, line, area, doughnut, radar, scatter vô nghĩa, hai cột trục cắt khúc, màu cầu vồng, không target); `Du_lieu`; `Demo_TrucY`: hai cột T3/T4/2024 cạnh nhau (trục từ 236.000.000 và từ 0) | 2, 8, 9 | `Demo_TrucY`: E4 = 0,96% · E5 = 2,91 · E6 = 1,01; pie ngành: 34,5% · 22,5% · 18,0% · 11,6% · 11,6% · 1,8% |
| `Demo/B6_Demo_Chart.xlsx` | `Demo_Combo` (A3:C9); `Demo_Waterfall` (A3:B17: Target, 12 Variance T3, Actual); `Demo_Phan_phoi` (cột B = 3.144 giá trị đơn Q1; thống kê D1:E9); `Demo_Scatter` (B3:C15 + `=CORREL`); `Demo_Map` (A3:B12) | 11–14 | Combo: GM% toàn chuỗi T1–T4 34,1% (C12); Waterfall: B17 = 269.608.000, E4 = −15.392.000, E6 = 0, E7 = 68% (ST07); Phân phối: 3.144 · 299.445 · 198.000 · 89.000 · 412.000 · 896.500 · 138 · SKEW 1,77; Scatter: r = 0,78 · R² = 0,61 · bỏ ST12: 0,90 · Hà Nội+TP.HCM 4,3 năm / 137.519.200 · còn lại 2,7 năm / 78.689.000; Map: B13 = 296.965.000 |
| `Demo/B6_Demo_DuBao.xlsx` | Table `tblNgay` (A3:B489) ở `Demo_Trendline`; `Thong_ke` (trung bình theo thứ D4:E10, E13:E20); `Demo_ETS` (B3:B8: 3 hàm ETS) | 23, 24 | tblNgay: 486 ngày, tổng 4.436.849.500; theo thứ: 8.434.264 → 10.547.341; Chủ nhật/Thứ Hai +25,1%; cuối tuần/ngày thường +19,0%; độ dốc 4.742 · R² 0,045 · dự báo tuyến tính 01/05/2025 10.284.071; ETS: điểm 6.731.497 · CONFINT 6.644.736 · SEASONALITY 121 |

Dashboard (Slide 17–21, 26–30, 32–33) dựng trên file khởi đầu `../B5_Data_Model_DAX/Practice/B5_MCI_Mart_Answer.xlsx` — **chạy được nguyên trạng** (Data Model, 14 measure), nên không tạo file demo riêng (`CLAUDE.md`: được dùng lại file Answer khi demo chạy nguyên trạng). Bản dựng sẵn để chấm chéo là `Practice/B6_MCI_Mart_Dashboard_Answer.xlsx`: dựng bằng macro VBA chạy trong Excel Windows (`2_Dataset/_generator/gen_b6_vba.py` → `vba/B6Build.bas`, module `B6Build`, chạy `B6Build.RunAll`; Excel COM từ Python không dùng được trong môi trường này): sheet `Dashboard`, `Data_Dash`, `Cai_dat`, `Brief`, `Du_bao`, `Kiem_tra_B6` (6 thẻ × tháng 2/3/4 khớp bảng Số liệu; waterfall tháng 3 tổng −15.392.000; Top 5 tháng 4; combo 4 tháng).

### Bài tập về nhà (bắt buộc) — tập dượt capstone C1, C5, C6 (đã dựng 04/10/2026)

Nguyên tắc 11 (`Course_Plan.md`): đặt trong một tình huống công việc, có lỗi cài sẵn chỉ lộ ra khi đối soát, gắn với tiêu chí rubric B7: **C1 Brief · C5 Dashboard · C6 Insight**.

**Tình huống.** Marketing Director xin một dashboard theo dõi khách hàng tháng 4/2025. Bạn nhận một dashboard cũ do đồng nghiệp làm (sheet `Dashboard_Cu`) và phải kiểm tra rồi làm lại.

**Đề `Homework/B6_BTVN_MCI_Mart_Dashboard_Marketing.xlsx`:** file khởi đầu (đáp án B5) + sheet `Dashboard_Cu` (lỗi cài sẵn) + `De_bai` + `Brief` (mẫu) + `Nhat_ky_doi_soat` + `Insight`.

**Lỗi cài sẵn trong `Dashboard_Cu`** (lộ ra khi đối soát với measure):

| # | Lỗi | Số trong dashboard cũ | Số đúng | Tiêu chí |
|---|---|---|---|---|
| 1 | Thẻ "Active Customers" cộng tay ba miền (gõ cứng, không CUBEVALUE) | 713 | **670** (DISTINCTCOUNT không cộng dồn) | C5 |
| 2 | Thẻ KPI không có target và RAG; hoặc RAG chỉ màu | — | target 700; 95,7% ▲ | C5 |
| 3 | Pie tỷ trọng doanh thu theo phân khúc kèm 3D; trục bar cắt khúc | — | Thường 46,5% · Thân thiết 40,1% · VIP 13,4% | C5 |
| 4 | Slicer chỉ nối một Pivot | — | Report Connections | C5 |
| 5 | Câu hỏi dạng chủ đề, không đo lường được ("Xem khách hàng") | — | viết lại có chỉ số · chiều · thời gian · mốc | C1 |

**Việc phải làm:** (1) **Đối soát** `Dashboard_Cu` với measure, ghi `Nhat_ky_doi_soat` (lỗi · số cũ · số đúng · cách sửa) *(D2)*. (2) Viết **brief** cho Marketing Director: người xem, 3–5 câu hỏi đo lường được, tần suất, quyết định *(C1)*. (3) Dựng **dashboard 1 trang** ≤ 8 visual: thẻ KPI (CUBEVALUE, target, RAG có biểu tượng), Slicer nối đủ, alt text, tiêu đề là kết luận *(C5)*. (4) Viết **insight SCR** 3 câu kèm hành động cho Marketing Director *(C6)*.

**Số tự kiểm tra (tháng 4/2025):** Active Customers **670** (target 700 → 95,7% ▲); theo phân khúc: Thường **360** · Thân thiết **251** · VIP **59** (cộng 670); doanh thu theo phân khúc 138.190.000 · 119.096.000 · 39.679.000; theo miền 234 + 151 + 328 = 713 ≠ 670.

**Đáp án `_Answer`:** dashboard đã làm lại, `Brief` mẫu, `Nhat_ky_doi_soat` mẫu, `Insight_mau`, `Kiem_tra` (số ở trên); chấm bằng agent `homework-grader`. Dựng bằng `2_Dataset/_generator/gen_b6_hw_vba.py` → `vba/B6HW1.bas` (macro `B6HW1.BuildDe` cho đề, `B6HW1.BuildAnswer` cho đáp án; xem đầu file .bas). `Kiem_tra` 10/10 ✔: 670 · 360/251/59 · 138.190.000/119.096.000/39.679.000 · 296.965.000 · 95,7% · 713. Slicer Tháng đã kiểm: chọn tháng 3 thì 3 thẻ KPI và 3 biểu đồ cùng đổi (637 khách hoạt động).

### Handout — dàn ý dự kiến (Course Plan: mục Handout B6)

1. Chart ít dùng hơn: Area, Pie, Bubble, Radar, Surface/Contour, Stock, Funnel, Sunburst, Histogram/Pareto nâng cao [DV PDF p7, 9, 12, 16–19]
2. Brief mẫu và 3 câu hỏi then chốt (có bản in trắng)
3. CUBEVALUE / CUBEMEMBER: cú pháp, Convert to Formulas, lỗi `#N/A`
4. Conditional Formatting Rules Manager, Icon Sets, Theme cho dashboard
5. Form Controls, dynamic ranges, custom number format
6. `FORECAST.ETS.SEASONALITY`, `FORECAST.ETS.STAT`, Data Completion, Aggregate Duplicates của Forecast Sheet
7. Sneak peek Power BI (PQ PDF p111–112): khi nào nên chuyển công cụ
8. Checklist chấm chéo (Phụ lục D) + checklist dự báo; bài tập về nhà; checklist trước Buổi 7

Ảnh handout dự kiến ở `Materials/B6/Images/_DANH_SACH_ANH.md` (bảng Handout); chốt tên và vùng khi viết `build_docs_b6.py`.

### Việc còn lại (theo thứ tự)

1. **Duyệt kịch bản.**
2. ~~**Đo ETS**~~ — đã đo bằng hàm trên Excel for Mac 03/10/2026 (Slide 24); kiểm lại bằng hộp thoại Forecast Sheet trên Windows khi chụp ảnh.
3. ~~**Dựng file Demo**~~ — đã dựng (`build_b6_demos.py`, Mac). ~~Đáp án `Practice/`~~ — đã dựng 04/10/2026 (macro `B6Build`).
4. ~~**Bài tập về nhà**~~ — đã dựng 04/10/2026 (macro `B6HW1`). ~~Hoàn thiện file đáp án~~ — đã làm 04/10/2026 (macro `B6Polish`): `Dashboard` của cả `Practice/` và `Homework/..._Answer.xlsx` đã **Protect Sheet** (không mật khẩu, Slicer mở khóa, cho dùng PivotTable); Top 5 đổi trục sang "tr"; waterfall bỏ nhãn số vì chồng nhau (số nằm ở `Data_Dash`). **Còn mở:** màu biểu tượng ✔ ▲ ✘ lấy theo icon set mặc định; ngưỡng RAG **đã chốt 05/10/2026 theo KPI Dictionary: ✔ ≥ 100% · ▲ 90% – <100% · ✘ < 90%** (Cai_dat C14 = 90%; mức Critical 80% chỉ nằm trong KPI Dictionary). Màu icon giữ mặc định (Leo: không cần đổi).
5. ~~**Brief và dataset capstone**~~ — đã dựng 04/10/2026 (`gen_capstone.py`, `build_brief_b7.py`); còn xuất PDF brief bằng Word và để giảng viên đọc lại thư khách hàng.
6. `session-builder` chế độ `khung` → `screenshot-capturer` (16 ảnh deck; ảnh handout H01–…) → `chot` → handout (đã có bản nháp, còn chèn ảnh H01–H07), infographic (đã dựng) → `consistency-checker`.

### Còn cần làm trước khi dạy

- **Forecast Sheet trên bảng Ngày–Doanh thu (mục chưa tick trong `Course_Plan.md`):** đã đo bằng hàm trên Excel for Mac (Slide 24: tự phát hiện = 121, không phải 7; dải 95% rộng hơn giá trị dự báo; kiểm chứng ngược −5,7%). **Còn lại:** lặp bằng hộp thoại Forecast Sheet trên Windows (Detect Automatically có ra chu kỳ khác 7 không; Forecast Start = 01/04/2025).
- **Tên Slicer thật** (`Slicer_Thang`, `Slicer_Region`) và cú pháp truyền Slicer làm đối số CUBEVALUE trên Excel 365 Windows.
- **Map (Filled Map)** nhận đủ 9 tên thành phố không; nếu không, chốt phương án bar.
- **Icon Set với Type = Formula** trỏ `Cai_dat` hoạt động đúng trên Excel 365.
- Chạy thử toàn bộ Phần 6 từ file khởi đầu; ghi thời gian thật của Slide 27 và 29.

### Ghi chú so với Course Plan

- **Ngưỡng RAG (chốt 05/10/2026):** theo KPI Dictionary (Phụ lục C và sheet `KPI_Dictionary`: Target 100%, **Warning 90% / Critical 80%**), dashboard 3 mức: ✔ Đạt ≥ 100% · ▲ Cảnh báo 90% – <100% · ✘ Không đạt < 90%. Mức Critical 80% chỉ nằm trong KPI Dictionary (B5 vẫn chia 4 nhãn). YoY: mức đạt = YoY ÷ 10% (KPI Dictionary ghi 50% / 20%, không dùng ở dashboard). Design System đã sửa theo.
- **Slicer "điều khiển mọi visual" (Phụ lục D, rubric C5 Level 3):** kịch bản cho phép **một visual ngữ cảnh** (combo 4 tháng) không nối Slicer Tháng, với điều kiện ghi rõ trên tiêu đề. Nếu rubric muốn nghiêm, đổi combo sang chạy theo một Slicer khác hoặc ghi ngoại lệ này vào rubric.
- **Cuối tuần "khoảng 25%" (README dataset, Phụ lục F #6):** đúng khi so **Chủ nhật với thứ Hai** (+25,1%); cuối tuần so ngày thường là **+19,0%**. Kịch bản nêu cả hai.
- **ETS và mùa 7:** nhịp tuần có thật nhưng nhỏ so với biến động ngày (R² theo thứ 5,9%–15,3%); đo trên Excel: `FORECAST.ETS.SEASONALITY` trả **121** (không phải 7) → dạy **Set Manually = 7**. Course Plan viết "seasonality tự phát hiện" — **không đúng** với dữ liệu này. Dải 95% rộng hơn cả giá trị dự báo; tổng tháng 5 (237,7 triệu) thấp hơn tháng 4 thực tế cỡ 20%.
- **"Map theo vùng"** (Course Plan): Bắc/Trung/Nam không phải đơn vị địa lý Bing nhận; kịch bản dùng **thành phố** (cột `Store[City]`), phương án dự phòng là bar theo miền.
- **KPI row:** Course Plan nói "hàng KPI card"; kịch bản dùng **6 thẻ** (Doanh thu kèm target, Đơn, AOV, GM%, YoY, Khách). `Revenue vs Target %` hiển thị là mức đạt của thẻ Doanh thu.
- **Waterfall theo cửa hàng** thay "waterfall chênh lệch target" chung chung; **combo Actual vs Target** theo tháng cho câu chuyện #4; waterfall tháng 3 cho câu chuyện #2.
- **Giới hạn PivotChart** (không có Waterfall, Map, Histogram) là điểm giảng thêm so với Course Plan, vì nó giải thích vì sao có hai đường dẫn (Pivot, CUBEVALUE).
- **Bài tập về nhà B6 và capstone** cùng giao ở cuối buổi. Bài về nhà (Marketing) là bước tập dượt nhỏ trước capstone; capstone vẫn là bài lớn nộp ở B7.
