# Ảnh minh họa Buổi 4 (deck + handout)

Chụp trên Excel 365 **Windows**, zoom 100%, trên đúng file và query ghi dưới đây (đường dẫn tính từ `1_Sessions/B4_Power_Query/`). Lưu thành **PNG** vào thư mục này, **đúng tên file**. Ảnh được script tự chèn vào khung [CHỤP]; không chèn tay vào `.pptx` / `.docx` (bị ghi đè khi dựng lại).

Buổi 4 không có file Demo: hầu hết ảnh là cửa sổ **Power Query Editor**. Chụp trên một bản sao file khởi đầu (`../B3_PivotTable/Practice/B3_MCI_Mart_Q1_2025_Answer.xlsx` → `B4_<HoTen>.xlsx`), làm đúng các bước ở lời giảng của slide tương ứng trong `Slide_Script_B4.md` rồi chụp. Ảnh query Sales, Product, Target, Calendar có thể chụp trên `Practice/B4_MCI_Mart_Answer.xlsx` (Data → Get Data → Launch Power Query Editor, bấm vào bước cần xem). Chụp xong đóng file **không lưu**.

Máy chụp để Region **English (United States)**, để ảnh S11a (bước Changed Type tự động) hiện đúng lỗi Locale như lời giảng. Trước khi chụp, `Practice/Data_Sales/` phải có đúng 4 file (chưa chép file tháng 4).

Sau khi thêm ảnh, chạy trong `2_Dataset/`:

```bash
python3 _generator/build_slides_b4.py           # in ra các khung [CHỤP] còn thiếu ảnh
python3 _generator/build_student_slides.py B4
python3 _generator/build_docs_b4.py             # rồi xuất lại B4_Handout.pdf bằng Word
```

## Deck `Slide_B4.pptx` (S = số slide)

Bảng này trùng với bảng "Danh sách ảnh cần chụp" trong `Slide_Script_B4.md`.

| Tên file | Nội dung cần chụp | File · query · vùng | Slide |
|---|---|---|---|
| S04_ba_file_csv.png | Ba cửa sổ Excel mở `Sales_2025-01/02/03.csv` xếp cạnh nhau (View → Arrange All → Vertical), thấy cột Order_Date và Unit_Price | `Practice/Data_Sales/Sales_2025-01.csv`, `-02.csv`, `-03.csv` | 4 |
| S09_query_editor.png | Query Editor đầy đủ: bảng xem trước, thanh công thức, Applied Steps 3 bước Source · Promoted Headers · Changed Type, ô Name = Sales_T1 | Query `Sales_T1` (`Practice/Data_Sales/Sales_2025-01.csv`) ngay sau Transform Data | 9 |
| S11a_column_quality.png | View: Column quality + Column profile (entire data set); cột Unit_Price chọn: 100% Valid, Min 10, Max 650; Order_Date có vạch đỏ | Query `Sales_T1`, bước Changed Type tự động | 11 |
| S11b_column_distribution.png | Cột Store_Code chọn: Column statistics Distinct 23 | Query `Sales_T1`, trước Trim | 11 |
| S12a_locale_date.png | Hộp Change Type with Locale (Date, Vietnamese (Vietnam)) và cột Order_Date 100% Valid | Query `Sales_T1` | 12 |
| S12b_locale_price.png | Unit_Price sau Using Locale → Whole Number, Vietnamese: 19000, 42000… | Query `Sales_T1` | 12 |
| S13a_trim.png | Store_Code sau Trim: Column statistics Distinct 12 | Query `Sales_T1` | 13 |
| S13b_conditional_column.png | Hộp Add Conditional Column: Nhom_SL, Quantity is greater than 3 → "Nhiều", Else "Ít" | Query `Sales_T1` | 13 |
| S14a_remove_dup_sales.png | Applied Steps có Removed Duplicates, bước Counted Rows hiện 2725 | Query `Sales_T1` | 14 |
| S14b_remove_dup_product.png | Product: Removed Duplicates theo Product_Code, Counted Rows = 40 | Query `Product` (`Practice/B4_Lookups.xlsx`, bảng Product) | 14 |
| S15_keep_errors.png | Order_Date đọc bằng Locale English (US), Keep Errors: bảng chỉ còn dòng lỗi, khung chi tiết lỗi của một ô | Query `Sales_T1 (2)` (Duplicate của Sales_T1) | 15 |
| S18a_target_before.png | Target trước Unpivot: dòng 1 tiêu đề báo cáo, dòng 3 tên cột, Tháng 1…12 nằm ngang | Query `Target` (`Practice/B4_Target_2025.xlsx`), bước Navigation | 18 |
| S18b_target_after.png | Target sau Unpivot và làm gọn: Store_Code, Thang, Target, Ngay_Dau_Thang | Query `Target`, bước cuối | 18 |
| S19_group_by.png | Hộp Group By (Advanced): Store_Code; So_dong = Count Rows; Doanh_thu = Sum of Revenue | Query `DoiSoat_T1` (Reference của Sales_T1) | 19 |
| S21_from_folder.png | Hộp danh sách file của From Folder (4 file) và nút Combine ▾ đang mở | Data → Get Data → From Folder → `Practice/Data_Sales/` | 21 |
| S22a_merge_dialog.png | Hộp Merge: Sales × Product theo Product_Code, Left Outer, "The selection matches 28328 of 28328 rows" | Home → Merge Queries ▾ → Merge Queries as New | 22 |
| S22b_merge_expand.png | Query `Sales_Merge_Demo` sau Expand, cột Category | Query `Sales_Merge_Demo` | 22 |
| S25_calendar.png | Calendar: Date, Nam, Thang, Quy, Thu; thanh công thức hoặc Applied Steps thấy List.Dates | Query `Calendar` | 25 |
| S27_queries_pane.png | Khung Queries & Connections: 6 query + nhóm trợ giúp của From Folder, "28,328 rows loaded" dưới Sales | `B4_<HoTen>.xlsx`, Data → Queries & Connections (trước khi thả file tháng 4) | 27 |

## Handout `B4_Handout.docx` (H = thứ tự ảnh)

| Tên file | Nội dung cần chụp | File · query · vùng | Mục trong handout |
|---|---|---|---|
| H01_from_text_csv.png | Hộp From Text/CSV: File Origin 65001 (UTF-8), Delimiter Comma, Data Type Detection, nút Load / Transform Data | `Practice/Data_Sales/Sales_2025-01.csv` → Data → Get Data → From Text/CSV | 1. Tóm tắt |
| H02_query_options_locale.png | Hộp Query Options, CURRENT WORKBOOK → Regional Settings, ô Locale | Data → Get Data → Query Options (trên `B4_<HoTen>.xlsx`) | 2. Locale |
| H03_count_rows_statistics.png | Tab Transform: nhóm Table (Count Rows) và Number Column (Statistics ▾ mở, thấy Sum) | Query `Sales` | 3. Đối soát |
| H04_index_column.png | Hộp Add Index Column (Custom: Starting Index 2, Increment 1) và cột Index mới | Query `Sales_T1` | 4. Index Column |
| H05_pivot_column.png | Hộp Pivot Column: Values Column = Target, Advanced options: Don't Aggregate; kết quả 12 dòng | Reference của query `Target` | 5. Pivot Column |
| H06_queries_pane_menu.png | Khung Queries & Connections với menu chuột phải của query Sales (Edit, Duplicate, Reference, Load To…, Properties) | `B4_<HoTen>.xlsx`, Data → Queries & Connections | 6. Quản lý query |
| H07_source_name.png | Query gộp từ thư mục Cua_hang: cột Source.Name và hộp Text Before Delimiter "_" | `Homework/B4_BTVN_Du_lieu/Cua_hang` (From Folder) trên `Homework/B4_BTVN_MCI_Mart_T4_2025.xlsx` | 7. From Folder nâng cao |
| H08_manage_parameters.png | Hộp Manage Parameters: tham số DataFolder, Type Text, Current Value | `Practice/B4_MCI_Mart_Answer.xlsx` · Launch Power Query Editor → Manage Parameters | 8. PRO TIP |
| H09_privacy_levels.png | Data Source Settings → Edit Permissions: ô Privacy Level | `B4_<HoTen>.xlsx` · Data → Get Data → Data Source Settings | 9. Xử lý sự cố |

Ảnh chưa có thì deck và handout giữ khung "[CHỤP] …" (viền nét đứt). Trước khi phát tài liệu cho học viên, không còn khung nào.

**Trạng thái (30/09/2026):** đã chụp đủ 28 ảnh (19 deck + 9 handout) trên Excel 365 Windows, máy Region English (United States); deck, bản `_HocVien` và handout đã dựng lại, không còn khung [CHỤP]. Bản sao dùng để chụp và ảnh gốc toàn màn hình nằm trong `1_Sessions/B4_Power_Query/_capture/`.
