# Ảnh minh họa Buổi 5 (deck + handout)

Chụp trên Excel 365 **Windows**, zoom 100%, trên đúng file, sheet và cửa sổ ghi dưới đây (đường dẫn tính từ `1_Sessions/B5_Data_Model_DAX/`). Lưu thành **PNG** vào thư mục này, **đúng tên file**. Ảnh được script tự chèn vào khung [CHỤP]; không chèn tay vào `.pptx` / `.docx` (bị ghi đè khi dựng lại).

Buổi 5 có 2 file demo theo mẫu B3 (đã dựng 02/10/2026, đặc tả ở Phụ lục `Slide_Script_B5.md`): `Demo/B5_Demo_Model.xlsx` cho S05a/b, S10a/b, S12a/b và `Demo/B5_Demo_DAX.xlsx` cho S15 trở đi. Phần lớn ảnh là cửa sổ **Power Pivot** (Data View, Diagram View, hộp Measure) và PivotTable From Data Model. Chụp trên file demo (bản `_HocVien`, để tên sheet khớp file học viên), làm đúng các bước ở lời giảng của slide tương ứng rồi chụp. Cột File trong bảng dưới ghi `B5_<HoTen>.xlsx` hoặc `Demo/…`: ảnh S05–S12 chụp trên `Demo/B5_Demo_Model_HocVien.xlsx`, ảnh S15 trở đi trên `Demo/B5_Demo_DAX_HocVien.xlsx` (sheet theo lời giảng). Ảnh trạng thái cuối buổi (S10b, S12b, S17, S19, S22–S28, S30) có thể chụp trên `Practice/B5_MCI_Mart_Answer.xlsx` khi đã dựng. **S05b và S12a phải chụp trước khi tạo relationship.** Chụp xong đóng file **không lưu**. Không bấm Refresh All trên bản sao (tham số `DataFolder` trỏ tới máy dựng file).

Sau khi thêm ảnh, chạy trong `2_Dataset/`:

```bash
python3 _generator/build_slides_b5.py           # in ra các khung [CHỤP] còn thiếu ảnh
python3 _generator/build_student_slides.py B5
python3 _generator/build_docs_b5.py             # rồi xuất lại B5_Handout.pdf bằng Word
```

## Deck `Slide_B5.pptx` (S = số slide)

Bảng này trùng với bảng "Danh sách ảnh cần chụp" trong `Slide_Script_B5.md`.

| Tên file | Nội dung cần chụp | File · sheet · vùng | Slide |
|---|---|---|---|
| S05a_data_view.png | Cửa sổ Power Pivot, Data View, tab Sales; thấy 6 tab bảng ở dưới và dòng Record 1 of 30,247 | `Demo/B5_Demo_Model_HocVien.xlsx` → Data → Manage Data Model | 5 |
| S05b_diagram_roi.png | Diagram View: 6 hộp Sales, Product, Store, Customer, Target, Calendar, **chưa có đường nối** | Như trên, Home → Diagram View (trước Slide 10) | 5 |
| S10a_create_relationship.png | Hộp Create Relationship: Sales / Order_Date ↔ Calendar / Date | Power Pivot → Design → Create Relationship | 10 |
| S10b_diagram_sales.png | Diagram View: Sales ở giữa, 4 đường nối tới Product, Store, Customer, Calendar; ký hiệu 1 và * trên đường nối | Sau khi tạo 4 relationship của Sales | 10 |
| S12a_target_lap.png | Pivot Rows = Category, Values = Sum of Target: 6 dòng đều 3.665.000.000; khung PivotTable Fields có thanh vàng "Relationships between tables may be needed" | Sheet `PV_Grain`, **trước** khi nối Target | 12 |
| S12b_diagram_hai_fact.png | Diagram View đủ 6 bảng, 6 relationship; Sales và Target cùng nối vào Store và Calendar | Sau Slide 12 | 12 |
| S15_cot_cogs.png | Data View bảng Sales, cột COGS ở cuối, thanh công thức `=Sales[Quantity]*RELATED(Product[Unit_Cost])` | Power Pivot → Data View → Sales | 15 |
| S16_gm_dong.png | Pivot lọc Nam = 2025, Quy = 1: Average of GM_Dong 33,77% và Sum of GM_Dong | Sheet `PV_Loc` (cột GM_Dong thêm tạm) | 16 |
| S17_measure_dialog.png | Hộp Measure: Table name Sales, Measure name Gross Margin %, Formula `=DIVIDE([Total Revenue]-[Total COGS],[Total Revenue])`, Category Number, Format Percentage (ảnh chụp lúc Decimal places còn 2; giảng viên đổi thành 1 trên lớp) | Power Pivot → Measures → New Measure | 17 |
| S19_pivot_mien_thang.png | Pivot Rows Region, Columns Thang (1–3), Filters Nam = 2025 và Quy = 1, Values Total Revenue; ô Bắc × 3 = 90.210.000 đang chọn | Sheet `PV_Loc` | 19 |
| S20_lookup_vs_fact.png | Hai Pivot cạnh nhau: trái Rows Calendar[Thang] (T5–T12 chỉ có target); phải Rows Sales[Thang], chỉ còn cột Sum of Target = 3.665.000.000 ở mọi dòng | Sheet `PV_Lookup_vs_Fact`, trước khi ẩn cột | 20 |
| S22_calculate.png | Pivot Rows Region (Quý 1): Total Revenue và Revenue Mien Nam = 427.663.000 ở cả 3 dòng | Sheet `PV_Loc` | 22 |
| S23_pct_of_total.png | Pivot Rows Region (Quý 1): Total Revenue, Revenue All Stores = 941.454.000 mọi dòng, % Revenue Share 34,4% · 20,2% · 45,4% | Sheet `PV_Loc` | 23 |
| S24_target_thang.png | Pivot Filters Nam = 2025, Rows Calendar[Thang]: Total Revenue, Total Target, Revenue vs Target %, Variance; T1 120,6%, T5–T12 chỉ có target | Sheet `PV_Target` (Variance đã chặn ô trống) | 24 |
| S25_store_thang.png | Pivot Rows Store_Code, Columns Thang 1–4, Values Revenue vs Target %; ô ST07 × 3 = 44,9% đang chọn | Sheet `PV_Target_CuaHang` | 25 |
| S27a_mark_date_table.png | Hộp Mark as Date Table, Date column = Date | Power Pivot, tab Calendar → Design → Mark as Date Table | 27 |
| S27b_yoy_ytd.png | Pivot Filters Nam = 2025, Rows Thang 1–4: Total Revenue, Revenue LY, Revenue YTD, Revenue YoY %; T4 YTD 1.238.419.000 | Sheet `PV_YoY` | 27 |
| S28_yoy_cua_hang.png | Pivot Rows Store_Code, Filters Nam = 2025 và Thang 1–4: Total Revenue, Revenue LY, Revenue YoY %; dòng ST12 hai ô trống | Sheet `PV_YoY` | 28 |
| S30_gm_theo_nam.png | Pivot Rows Calendar[Nam], không lọc, Values Gross Margin %: 2024 31,2% · 2025 34,1% | Sheet `PV_YoY` | 30 |

## Handout `B5_Handout.docx` (H = thứ tự ảnh)

Theo dàn ý handout ở Phụ lục kịch bản B5. Tên và vùng chốt khi viết `build_docs_b5.py`; cập nhật bảng này cùng lúc.

| Tên file | Nội dung cần chụp | File · sheet · vùng | Mục trong handout |
|---|---|---|---|
| H01_hide_client_tools.png | Data View bảng Sales: các cột khóa / ngày đã ẩn (màu xám) và menu chuột phải có Hide from Client Tools / Unhide from Client Tools | `Practice/B5_MCI_Mart_Answer.xlsx` · Power Pivot → Data View → Sales | 1. Ẩn cột, bỏ cột thừa |
| H02_hierarchy.png | Diagram View: bảng Calendar có hierarchy Nam → Quy → Thang → Date; bên cạnh Pivot dùng hierarchy, nút +/− drill | `Practice/B5_MCI_Mart_Answer.xlsx` · Diagram View (Create Hierarchy) | 2. Hierarchy |
| H03_lookup_to_lookup.png | Diagram View: quan hệ `Customer[Home_Store_Code]` → `Store[Store_Code]` hiện nét đứt (inactive) | Bản sao `Practice/B5_MCI_Mart_Answer.xlsx` · Diagram View | 3. Lookup-to-lookup, active / inactive |
| H04_bad_cardinality.png | Thông báo lỗi khi tạo quan hệ thẳng `Sales[Store_Code]` – `Target[Store_Code]` (cả hai cột có giá trị trùng) | Bản sao `Practice/B5_MCI_Mart_Answer.xlsx` · Design → Create Relationship | 4. Bad cardinality |
| H05_disconnected_slicer.png | Slicer bảng ngưỡng (80% / 90% / 100%) và Pivot đếm số cửa hàng dưới ngưỡng theo tháng | `Practice/B5_MCI_Mart_Answer.xlsx` (sheet minh họa sẽ thêm khi dựng handout) | 8. Disconnected slicer |

Ảnh chưa có thì deck và handout giữ khung "[CHỤP] …" (viền nét đứt). Trước khi phát tài liệu cho học viên, không còn khung nào.

**Trạng thái (03/10/2026):** đã chụp đủ 24/24 ảnh (19 ảnh deck S-series + 5 ảnh handout H01–H05), chụp trên bản sao trong `1_Sessions/B5_Data_Model_DAX/_capture/` (file gốc Demo/Practice không bị sửa). Ghi chú: số liệu trong ảnh theo định dạng en-US (dấu phẩy ngăn nghìn, dấu chấm thập phân); H02 gồm cửa sổ Power Pivot + Pivot hierarchy; H03 dùng bản sao Answer có thêm quan hệ inactive Customer[Home_Store_Code]→Store; H05 dựng trên bản sao Demo (bảng `Muc` 80/90/100% không nối relationship, measure `Stores Duoi Nguong`, slicer chọn 90% → tháng 1–4 = 0 · 8 · 6 · 4, tổng 12 cửa hàng) – sheet minh họa này chưa có trong file Practice, cần dựng khi viết handout/bài tập.


**03/10/2026:** đã chạy lại `build_slides_b5.py` → `build_student_slides.py B5` (deck 39 slide, đủ ảnh) và viết + chạy `build_docs_b5.py` (handout 9 mục, 5/5 ảnh, không còn [CHỤP]). Mục 8 handout hướng dẫn học viên dựng sheet `PV_Nguong` (bảng `Muc`, measure `Stores Duoi Nguong`, slicer) trên bản sao file đáp án; sheet này chưa nằm trong `Practice/`. Grand Total = 12 trong ảnh H05 là so doanh thu T1–T4 với target cả năm, không phải tổng các tháng. `B5_Handout.pdf` đã xuất bằng Word (8 trang).
