# Dataset MCI Mart & file dữ liệu theo buổi

Dữ liệu **giả lập**, sinh bằng mã có seed cố định (2025), nên chạy lại luôn ra đúng cùng một bộ số liệu. Mọi tên người, cửa hàng và sản phẩm đều là hư cấu.

## Vị trí các file

Đường dẫn tính từ thư mục gốc của khóa học.

| Thư mục / file | Nội dung | Dùng ở |
|---|---|---|
| `2_Dataset/MCI_Mart_Master.xlsx` | Bản gốc cho giảng viên, gồm các sheet README, Sales (30.247 dòng, 01/2024–04/2025), Product (40), Store (12), Customer (1.500), Target_2025 | Tham chiếu, soạn đáp án |
| `1_Sessions/B1_Nen_tang_Excel_Du_lieu/Practice/B1_MCI_Mart_Q1_2025_Dirty.xlsx` | File "bẩn" dạng báo cáo Quý 1/2025: 6.187 dòng, cài sẵn 9 loại lỗi (cấu trúc và giá trị) + định dạng thừa (không tính là lỗi dữ liệu) | B1: phát cho học viên |
| `1_Sessions/B1_Nen_tang_Excel_Du_lieu/Practice/B1_MCI_Mart_Q1_2025_Answer.xlsx` | Đáp án B1: tblSales + Revenue + Pct_Share + Validation; sheet `Kiem_tra` đối soát (8/8 ✔); sheet `Dap_an_loi` | B1: chữa bài; **file khởi đầu B2** |
| `1_Sessions/B1_Nen_tang_Excel_Du_lieu/Slide_B1_HocVien.pptx` | Deck B1 bỏ speaker notes và đáp án quiz, dựng từ `Slide_B1.pptx` | B1: phát học viên |
| `1_Sessions/B1_Nen_tang_Excel_Du_lieu/Homework/B1_BTVN_MCI_Mart_T4_2025.xlsx` | Bài tập về nhà bắt buộc: báo cáo tháng 4/2025 theo cửa hàng (1.919 dòng), cài 10 loại lỗi; sheet `De_bai` | B1: phát cuối buổi, nộp trước B2 |
| `1_Sessions/B1_Nen_tang_Excel_Du_lieu/Homework/B1_BTVN_MCI_Mart_T4_2025_Answer.xlsx` | Đáp án: tblSales_T4 + Revenue + Pct_Share + Validation; `Kiem_tra` (9/9 ✔), `Dap_an_loi`, `Insight_mau` | Chấm bài (chỉ giảng viên / trợ giảng) |
| `1_Sessions/B1_Nen_tang_Excel_Du_lieu/Demo/B1_Demo_PctShare.xlsx` | Bảng % đóng góp 3 miền: bản sai (không `$`, ra #DIV/0!), bản đúng, ô IFERROR; sheet `Huong_dan` | B1: demo slide 19, 21, 22 (bản `_HocVien` phát học viên) |
| `1_Sessions/B1_Nen_tang_Excel_Du_lieu/Demo/B1_Demo_Table_Validation.xlsx` | Dữ liệu Q1 đã tidy nhưng chưa là Table, còn 5 ô sai (3 ST3, 2 Quantity âm); sheet `Them_dong` (5 dòng T4/2025) để dán thử; sheet `Huong_dan` | B1: demo slide 26, 27, 28 (bản `_HocVien` phát học viên) |
| `1_Sessions/B1_Nen_tang_Excel_Du_lieu/Handout/B1_Handout.docx` / `.pdf` | Handout tự học B1 (11 mục) | B1: phát sau buổi |
| `1_Sessions/B1_Nen_tang_Excel_Du_lieu/Handout/B1_Checklist_KiemTraMay.docx` / `.pdf` | Checklist kiểm tra Power Query / Power Pivot / 64-bit; học viên tự làm, không đạt thì liên hệ trợ giảng trong Buổi 1 | Gửi học viên trước B1; slide 8 |
| `1_Sessions/B1_Nen_tang_Excel_Du_lieu/Handout/B1_Infographic_TomTat.png` | Infographic tóm tắt B1, 1 trang A4 dọc (2480 × 3508 px): câu hỏi mở buổi, quy trình BI, 6 nội dung chính, số tự đối chiếu, câu insight, việc trước B2. Nguồn `_generator/infographic_b1.html` | B1: phát sau buổi, cùng handout |
| `1_Sessions/B2_Ham_va_KPI/Practice/B2_Product_DuplicateKey.xlsx` | Danh mục sản phẩm 41 dòng, P021 trùng khóa (dòng 22 là bản ghi cũ, Unit_Cost 62.000, đứng trước). Không còn ghi chú giảng viên, phát thẳng | B2: slide 23, Thực hành 1 |
| `1_Sessions/B2_Ham_va_KPI/Practice/B2_MCI_Mart_Q1_2025_Answer.xlsx` | Đáp án B2: tblSales làm giàu (Category, Unit_Cost, COGS, Gross_Margin, Thang, Quy, Thu, Phan_khuc), tblProduct sạch, `KPI_SUMIFS` (Miền × Tháng: doanh thu, số đơn, AOV, GM%), `KPI_Card` (3 card), `Kiem_tra` (10/10 ✔). Đã lưu bằng Excel | B2: chữa bài; **file khởi đầu B3** |
| `1_Sessions/B2_Ham_va_KPI/Slide_B2.pptx` / `Slide_B2_HocVien.pptx` | Deck B2 (41 slide) dựng từ `Slide_Script_B2.md`; bản học viên bỏ notes và đáp án quiz | B2 |
| `1_Sessions/B2_Ham_va_KPI/Demo/B2_Demo_Logic.xlsx` | 40 sản phẩm + 2 sản phẩm chưa có giá: IF, IFS, AND/OR/NOT, IFERROR; sheet `Huong_dan` | B2: demo slide 7, 8, 9 (bản `_HocVien` phát học viên) |
| `1_Sessions/B2_Ham_va_KPI/Demo/B2_Demo_ThongKe.xlsx` | 3.144 đơn Q1/2025 (giá trị đơn): mean/median, STDEV, SKEW, outlier IQR; sheet `Demo_RANK` 12 cửa hàng | B2: demo slide 11, 12, 13 |
| `1_Sessions/B2_Ham_va_KPI/Demo/B2_Demo_SUMIFS.xlsx` | tblSales Q1 có cột Thang, COGS; SUMIFS/COUNTIFS/AVERAGEIFS, bảng Miền × Tháng, GM% trung bình của trung bình; sheet `Demo_DynamicArray` (PRO TIP) | B2: demo slide 15, 16, 17 |
| `1_Sessions/B2_Ham_va_KPI/Demo/B2_Demo_Lookup.xlsx` | Product có P021 trùng; VLOOKUP khớp gần đúng, INDEX/MATCH, XLOOKUP, khóa trùng (143 dòng P021), khóa khác kiểu (TRIM/VALUE) | B2: demo slide 20–24 |
| `1_Sessions/B2_Ham_va_KPI/Demo/B2_Demo_Text_Date.xlsx` | LEFT/MID/RIGHT, TRIM, &, TEXT; EOMONTH, WEEKDAY, NETWORKDAYS có ngày lễ Tết 2025 | B2: demo slide 26, 27 |
| `1_Sessions/B2_Ham_va_KPI/Homework/B2_BTVN_MCI_Mart_T4_2025.xlsx` | Bài tập về nhà bắt buộc: tblSales_T4 sạch (1.919 dòng), 4 mã Product_Code dính dấu cách cuối; Store, Product; sheet `De_bai` | B2: phát cuối buổi, nộp trước B3 |
| `1_Sessions/B2_Ham_va_KPI/Homework/B2_BTVN_MCI_Mart_T4_2025_Answer.xlsx` | Đáp án: XLOOKUP Region/Category/Unit_Cost, KPI Miền × Ngành, KPI theo Miền, card AOV; `Kiem_tra` (8/8 ✔), `Dap_an_loi`, `Insight_mau`. Đã lưu bằng Excel | Chấm bài (chỉ giảng viên / trợ giảng) |
| `1_Sessions/B2_Ham_va_KPI/Handout/B2_Handout.docx` / `.pdf` | Handout tự học B2 (10 mục: hàm bổ sung, PRO TIP Dynamic Arrays, bài về nhà, checklist trước B3); 7 khung chờ [CHỤP], ảnh lấy từ `Materials/B2/Images/` khi có | B2: phát sau buổi |
| `1_Sessions/B2_Ham_va_KPI/Handout/B2_KPI_Definition_Card.docx` / `.pdf` | Mẫu KPI Definition Card 1 trang (10 trường) | B2: in 3 tờ / học viên; slide 5, 30 |
| `1_Sessions/B3_PivotTable/Practice/B3_MCI_Mart_Q1_2025_Answer.xlsx` | Đáp án B3: cấu trúc đáp án B2 + cột So_don; 6 sheet Pivot (PV_MienThang có khối đối soát GETPIVOTDATA ↔ KPI_SUMIFS, PV_NganhHang % of Column Total, PV_Thang MoM / lũy kế / AOV / GM% bằng Calculated Field, PV_Ngay, PV_Thu, PV_CuaHang); Slicer Store_Code nối 3 Pivot, Timeline Order_Date, 3 PivotChart; ST07_T3_ChiTiet (kết quả Show Details); `Kiem_tra` (14/14 ✔). Đã lưu bằng Excel | B3: chữa bài; **file khởi đầu B4** |
| `1_Sessions/B3_PivotTable/Slide_B3.pptx` / `Slide_B3_HocVien.pptx` | Deck B3 (42 slide) dựng từ `Slide_Script_B3.md`; bản học viên bỏ notes và đáp án quiz. Đủ 24 ảnh chụp thật (`Materials/B3/Images/`) | B3 |
| `1_Sessions/B3_PivotTable/Demo/B3_Demo_Pivot.xlsx` | tblSales Q1 giống file khởi đầu (20 cột), KPI_SUMIFS (doanh thu, số đơn), Product, Store; dựng mọi Pivot trực tiếp theo sheet `Huong_dan` | B3: demo slide 9–29 (bản `_HocVien` phát học viên) |
| `1_Sessions/B3_PivotTable/Demo/B3_Demo_Refresh.xlsx` | Cùng dữ liệu dạng vùng thường (Sales_Vung) và Table (Sales_Table), Pivot dựng sẵn PV_Vung / PV_Table, 10 dòng tháng 4 để dán (Them_dong) | B3: demo slide 12; handout mục 6 (Subtotal) |
| `1_Sessions/B3_PivotTable/Demo/B3_Demo_Loi_Du_lieu.xlsx` | tblSalesLoi: Revenue có 2 ô chữ + 1 ô trống (Pivot mặc định Count, Sum thiếu 974.000), Order_Date có 3 ô ngày dạng chữ (Cannot group that selection) | B3: demo slide 19, 23 |
| `1_Sessions/B3_PivotTable/Homework/B3_BTVN_MCI_Mart_T1_T4_2025.xlsx` | Bài tập về nhà bắt buộc: tblSales4T tháng 1–4/2025 (8.106 dòng) đã làm giàu, cài sẵn 6 ô Quantity tháng 4 dạng chữ; `So_lieu_he_thong` để đối soát; Store; `De_bai` | B3: phát cuối buổi, nộp trước B4 |
| `1_Sessions/B3_PivotTable/Homework/B3_BTVN_MCI_Mart_T1_T4_2025_Answer.xlsx` | Đáp án: Pivot đối soát, Miền × Tháng (% so tháng trước), tỷ trọng ngành, Cửa hàng × Tháng, AOV (Calculated Field), Slicer Region, 3 PivotChart (câu 2 Miền, câu 3 Ngành, câu 4 Cửa hàng/ST07); `Kiem_tra` (9/9 ✔), `Dap_an_loi`, `Insight_mau`. Đã lưu bằng Excel | Chấm bài (chỉ giảng viên / trợ giảng) |
| `1_Sessions/B3_PivotTable/Handout/B3_Handout.docx` / `.pdf` | Handout tự học B3 (10 mục: Show Values As mở rộng, Calculated Item, Report Filter Pages, Subtotal, PRO TIP, GETPIVOTDATA, bài về nhà, checklist trước B4); đủ 8 ảnh chụp thật (`Materials/B3/Images/`) | B3: phát sau buổi |
| `1_Sessions/B3_PivotTable/Handout/B3_Infographic_TomTat.png` | Infographic tóm tắt B3, nguồn `_generator/infographic_b3.html`, số liệu từ `build_b3_meta.json` | B3: phát sau buổi |
| `1_Sessions/B4_Power_Query/Practice/Data_Sales/Sales_2024.csv` | 22.141 dòng, doanh thu 3.198.430.500; ngày dd/mm/yyyy, giá không có dấu ngăn nghìn, không trùng, không khoảng trắng | B4: nguồn From Folder |
| `1_Sessions/B4_Power_Query/Practice/Data_Sales/Sales_2025-01/02/03.csv` | Có lỗi cài sẵn: giá dạng "19.000" (dấu chấm ngăn nghìn kiểu Việt Nam), ngày dd/mm/yyyy, 25 dòng Store_Code dính khoảng trắng thừa, 1 dòng trùng tuyệt đối mỗi file. Sau khi làm sạch: 2.725 / 1.595 / 1.867 dòng, doanh thu 425.850.000 / 245.996.000 / 269.608.000 (Quý 1: 6.187 dòng, 941.454.000, khớp B2/B3). Thư mục `Data_Sales/` phát cho học viên có **đúng 4 file** (chưa có tháng 4) | B4: query tập `Sales_T1` (file tháng 1), From Folder, Using Locale, Trim, Remove Duplicates |
| `1_Sessions/B4_Power_Query/Practice/New_Month/Sales_2025-04.csv` | File tháng mới, cùng dạng lỗi 3 file trên; 1.920 dòng thô → 1.919 dòng sạch, doanh thu 296.965.000 (khớp đáp án bài về nhà B2) | B4: demo Refresh (thả vào `Data_Sales/` giữa buổi) |
| `1_Sessions/B4_Power_Query/Practice/B4_Lookups.xlsx` | Bảng (Excel Table) Product (41 dòng: 40 sản phẩm + P015 trùng 1 dòng; Merge khi chưa bỏ trùng: 28.328 → 29.408 dòng), Store (12), Customer (1.500) | B4: Remove Duplicates, Merge minh họa, load riêng (Only Create Connection + Data Model) |
| `1_Sessions/B4_Power_Query/Practice/B4_Target_2025.xlsx` | Target dạng cột Tháng 1…12, tiêu đề ở dòng 3, 12 cửa hàng | B4: Unpivot → 144 dòng (Store_Code, Thang, Target, Ngay_Dau_Thang); B5: fact thứ hai |
| `1_Sessions/B4_Power_Query/Slide_Script_B4.md` | Kịch bản 39 slide, 150 phút (bản 1.3); dạy trực tiếp trên các file `Practice/` ở trên (B4 không có file `Demo/`) | B4: kịch bản giảng dạy |
| `1_Sessions/B4_Power_Query/Practice/B4_MCI_Mart_Answer.xlsx` | Đáp án B4, dựng bằng Power Query thật (Excel COM trên Windows, `_generator/build_b4.py`): tham số `DataFolder`; Sales = From Folder `Data_Sales` (Trim, Using Locale vi-VN, Remove Duplicates, Revenue/Nam/Thang), load Table `Sales` (sheet `Sales_PQ`) + Data Model, tắt background refresh; Product (bỏ P015 trùng), Store, Customer, Target (Unpivot, Thang số, Ngay_Dau_Thang), Calendar 01/01/2024–31/12/2025 (731 dòng): Only Create Connection + Data Model; Pivot `PV_Refresh` (Nam × Thang: số dòng, doanh thu; lọc Store_Code); `Kiem_tra_B4` (10 dòng). Trạng thái sau khi thả tháng 4: Sales 30.247 dòng, 4.436.849.500. Giữ nguyên mọi sheet đáp án B3. Mở trên máy khác: Manage Parameters → sửa DataFolder. Đã dựng lại trên Windows theo kịch bản 1.3 (30/09/2026): `Kiem_tra_B4` 10/10 ✔ | B4: chữa bài; **file khởi đầu B5** |
| `1_Sessions/B4_Power_Query/Slide_B4.pptx` / `Slide_B4_HocVien.pptx` | Deck B4 (39 slide) dựng từ `Slide_Script_B4.md` bằng `_generator/build_slides_b4.py`; bản học viên (`build_student_slides.py B4`) bỏ notes và đáp án quiz. Đủ 19 ảnh chụp thật (`Materials/B4/Images/`, danh sách `_DANH_SACH_ANH.md`), không còn khung [CHỤP] | B4 |
| `1_Sessions/B4_Power_Query/Homework/B4_BTVN_Du_lieu/` | `Cua_hang/`: 12 CSV tháng 4/2025, mỗi cửa hàng một file (`ST01_2025-04.csv`…), không có cột Store_Code (nằm trong tên file); cài sẵn 12 dòng "Tổng cộng" cuối file, 8 Product_Code sai dạng (4 viết thường, 4 dính dấu cách), 3 dòng trùng; ngày dd/mm/yyyy, giá "19.000". File ST07 chỉ có ngày 1–15; `Bo_sung/ST07_2025-04_phan2.csv` chứa ngày 16–30 (61 dòng, 8.611.000) — phát hiện qua đối soát | B4: bài tập về nhà (phát cả thư mục) |
| `1_Sessions/B4_Power_Query/Homework/B4_BTVN_MCI_Mart_T4_2025.xlsx` | Đề bài tập về nhà bắt buộc: `De_bai` (7 việc), `So_lieu_he_thong` (12 cửa hàng: số dòng, Quantity, doanh thu), `Product` (40 dòng, Table), `Target_2025` (dạng cột tháng), `Insight` | B4: phát cuối buổi, nộp trước B5 |
| `1_Sessions/B4_Power_Query/Homework/B4_BTVN_MCI_Mart_T4_2025_Answer.xlsx` | Đáp án: bảng kết quả của các query (`Sales_T4` 1.919 dòng, 296.965.000; `KetQua_T4` % đạt target theo cửa hàng — ST07 54,3% khi thiếu file, 97,4% khi đủ; cả chuỗi 103,1%), `Dap_an_M` (mã M), `Kiem_tra` (11/11 ✔), `Dap_an_loi`, `Insight_mau`. Đã lưu bằng Excel | Chấm bài (chỉ giảng viên / trợ giảng) |
| `1_Sessions/B4_Power_Query/Handout/B4_Handout.docx` / `.pdf` | Handout tự học B4 (11 mục: Locale, đối soát trong Power Query, Index Column, công cụ cột số, Pivot Column / Transpose, quản lý query, Source.Name và Transform Sample File, PRO TIP tham số thư mục / Rolling Calendar, Formula.Firewall và privacy levels, bài về nhà, checklist trước B5); đủ 9 ảnh chụp thật (`Materials/B4/Images/`) | B4: phát sau buổi |
| `1_Sessions/B4_Power_Query/Handout/B4_Infographic_TomTat.png` | Infographic tóm tắt B4, nguồn `_generator/infographic_b4.html`, số liệu từ `build_b4_meta.json` | B4: phát sau buổi |
| `1_Sessions/B5_Data_Model_DAX/Slide_Script_B5.md` | Kịch bản 39 slide, 150 phút (bản nháp 1.0, chờ duyệt); dạy trên bản sao file khởi đầu `B4_MCI_Mart_Answer.xlsx` (B5 không có file `Demo/`). Số liệu từ `_generator/build_b5_meta.py` → `build_b5_meta.json`, đối soát 64/64 ô với `KPI_SUMIFS`. Deck `Slide_B5.pptx` (+ `_HocVien`) dựng bằng `_generator/build_slides_b5.py`, còn 19 khung [CHỤP]. Đáp án `Practice/B5_MCI_Mart_Answer.xlsx`, bài về nhà, handout, infographic: chưa dựng | B5: kịch bản giảng dạy |
| `2_Dataset/_generator/` | Mã sinh dữ liệu và dựng file (Python) | Tái tạo hoặc chỉnh dataset |

## Từ điển dữ liệu

**Sales** (grain: 1 dòng = 1 sản phẩm trong 1 đơn hàng):

| Cột | Mô tả |
|---|---|
| `Order_ID` | Mã đơn, dạng `HDyymm-nnnnn` |
| `Line_No` | Số thứ tự dòng trong đơn |
| `Order_Date` | Ngày bán |
| `Store_Code` | Mã cửa hàng |
| `Customer_ID` | Mã khách hàng |
| `Product_Code` | Mã sản phẩm |
| `Quantity` | Số lượng |
| `Unit_Price` | Giá bán thực tế tại thời điểm giao dịch; giá 2024 thấp hơn 2025 khoảng 4% |

File B1 có thêm `Region` và `Product_Name` (dạng báo cáo phẳng). File CSV ở B4 không có hai cột này; đó là dạng chuẩn hóa để dùng ở B5.

**Các bảng khác:**

| Bảng | Các cột |
|---|---|
| Product | `Product_Code`, `Product_Name`, `Category` (6 ngành), `Unit_Price` (niêm yết 2025), `Unit_Cost` |
| Store | `Store_Code`, `Store_Name`, `City`, `Region` (Bắc/Trung/Nam), `Store_Type` (Mall/Street), `Open_Date` |
| Customer | `Customer_ID`, `Full_Name`, `Gender`, `Birth_Year`, `Segment` (Thường/Thân thiết/VIP), `Join_Date`, `Home_Store_Code` |
| Target_2025 | `Store_Code`, `Store_Name`, 12 cột tháng. Giá trị = doanh thu cùng tháng 2024 × 110%, làm tròn triệu. ST12 dùng trung bình T9–T12/2024. |

## Số liệu tham chiếu

| Tháng | Doanh thu (VNĐ) | Số đơn | Target |
|---|---|---|---|
| 01/2025 | 425.850.000 | 1.398 | 353.000.000 |
| 02/2025 | 245.996.000 | 823 | 298.000.000 |
| 03/2025 | 269.608.000 | 923 | 285.000.000 |
| 04/2025 | 296.965.000 | 966 | 288.000.000 |

**Quý 1/2025:** 6.187 dòng · 3.144 đơn · doanh thu 941.454.000 · AOV khoảng 299.000.

| Miền | Doanh thu Quý 1/2025 |
|---|---|
| Bắc | 323.978.000 |
| Trung | 189.813.000 |
| Nam | 427.663.000 |

## "Câu chuyện" cài sẵn trong dữ liệu (để khai thác insight ở B3–B7)

1. **Tết:** doanh thu tăng khoảng 3 tuần trước Tết (29/01/2025; 10/02/2024) và giảm sâu trong kỳ nghỉ. Bánh kẹo tăng tỷ trọng (50,8% doanh thu 20 ngày trước Tết 2025, so với 27,9% ngày thường); Đồ uống bán gấp khoảng 1,8 lần mỗi ngày nhưng tỷ trọng gần như không đổi (11,7% so với 11,5%).
2. **ST07 Nha Trang** cải tạo từ 10 đến 31/03/2025: tháng 3 chỉ đạt khoảng 45% target.
3. **ST12 Biên Hòa** khai trương 01/06/2024 và tăng dần trong 3 tháng đầu. Không có dữ liệu cùng kỳ cho T1–T5, nên YoY trống.
4. **Target lệch mùa Tết:** target lập theo cùng tháng 2024, nên T1/2025 vượt target (~121%) và T2/2025 hụt (~83%).
5. **Biên lợi nhuận gộp** 2025 (khoảng 33–35%) cao hơn 2024 (khoảng 31%), do giá bán tăng trong khi dataset dùng một bảng giá vốn duy nhất (giả định của dataset).
6. **Cuối tuần** doanh thu cao hơn khoảng 25%.

## Tái tạo dữ liệu

Chạy trong thư mục `2_Dataset/`. Các script ghi thẳng file vào đúng thư mục buổi ở trên.

```bash
python3 -m venv venv && venv/bin/pip install openpyxl pandas numpy python-docx python-pptx
venv/bin/python _generator/gen_dataset.py     # sinh dữ liệu (thư mục build/ cạnh script)
venv/bin/python _generator/build_files.py     # dựng các file .xlsx / .csv
venv/bin/python _generator/build_docs.py      # dựng handout và checklist .docx
venv/bin/python _generator/build_demos.py     # dựng file demo, bản giảng viên + bản _HocVien (đọc file Answer/Dirty/Master đã có)
venv/bin/python _generator/build_homework.py  # dựng bài tập về nhà: đề + đáp án (đọc MCI_Mart_Master.xlsx)
venv/bin/python _generator/build_student_slides.py  # Slide_B<n>.pptx → Slide_B<n>_HocVien.pptx, bỏ speaker notes và đáp án quiz (cần python-pptx); thêm B2 để chỉ dựng một buổi
python3 _generator/build_infographic.py B1   # infographic_b<n>.html → Handout/B<n>_Infographic_TomTat.png (Chrome headless, không cần venv); mỗi buổi một file. Con số viết `{{khóa|n}}` / `{{khóa|pct}}` được điền từ `build_b<n>_meta.json`
# Buổi 2 (đọc MCI_Mart_Master.xlsx và file Answer B1, không cần build/)
venv/bin/python _generator/build_b2.py            # Practice (Answer, Product trùng khóa), Demo (+ _HocVien), Homework; số liệu in ra và lưu build_b2_meta.json
venv/bin/python _generator/build_slides_b2.py     # Slide_B2.pptx: khuôn Slide_B1.pptx, notes lấy nguyên văn từ Slide_Script_B2.md (cần python-pptx)
venv/bin/python _generator/build_docs_b2.py       # B2_Handout.docx, B2_KPI_Definition_Card.docx
# Buổi 3 (đọc MCI_Mart_Master.xlsx và file Answer B1; cần Microsoft Excel trên Mac hoặc Windows để bước lưu lại chạy được)
venv/bin/python _generator/build_b3.py            # Practice, Demo (+ _HocVien), Homework; Pivot/Slicer/Timeline/PivotChart ghi bằng pivot_xml.py,
                                                  # rồi tự mở từng file bằng Excel (Mac: AppleScript; Windows: COM qua PowerShell), Refresh All, lưu,
                                                  # và kiểm tra sheet Kiem_tra; số liệu lưu build_b3_meta.json. Đường dẫn tự lấy theo vị trí script.
                                                  # Windows, từ thư mục MCI Excel: python 2_Dataset\_generator\build_b3.py (đóng các file B3 đang mở trước khi chạy)
venv/bin/python _generator/build_slides_b3.py     # Slide_B3.pptx (cần python-pptx)
venv/bin/python _generator/build_student_slides.py B3
venv/bin/python _generator/build_docs_b3.py       # B3_Handout.docx
python3 _generator/build_infographic.py B3
# Buổi 4 (đọc các file Practice/ của B4; không cần build/)
venv/bin/python _generator/build_b4_homework.py   # số liệu build_b4_meta.json + bài tập về nhà (thư mục CSV, đề, đáp án); tự mở đề và đáp án bằng Excel để lưu giá trị,
                                                  # kiểm tra Kiem_tra 11/11 ✔ (--no-excel để bỏ bước này). Chạy được trên Mac và Windows
python 2_Dataset\_generator\build_b4.py          # CHỈ WINDOWS, từ thư mục MCI Excel: Practice/B4_MCI_Mart_Answer.xlsx với Power Query thật (Excel qua COM,
                                                  # powerquery_com.py); chép tạm file tháng 4 vào Data_Sales rồi xóa lại; Kiem_tra_B4 phải 10/10 ✔. --print-m: chỉ in mã M
venv/bin/python _generator/build_slides_b4.py     # Slide_B4.pptx (cần python-pptx)
venv/bin/python _generator/build_student_slides.py B4
venv/bin/python _generator/build_docs_b4.py       # B4_Handout.docx (rồi xuất PDF bằng Word)
python3 _generator/build_infographic.py B4
# Buổi 5 (đọc file khởi đầu B4_MCI_Mart_Answer.xlsx: sheet Sales_PQ, KPI_SUMIFS; B4_Lookups.xlsx, B4_Target_2025.xlsx; đối chiếu MCI_Mart_Master.xlsx)
/Users/leopham/anaconda3/bin/python _generator/build_b5_meta.py   # chỉ tính số liệu → build_b5_meta.json (measure tính lại bằng pandas, đối soát 64 ô KPI_SUMIFS); không ghi file nào trong thư mục buổi
                                                  # Windows: python 2_Dataset\_generator\build_b5_meta.py. Đáp án có Data Model (build_b5.py) chưa viết: chỉ chạy được trên Windows
python 2_Dataset\_generator\build_b5.py           # CHỈ WINDOWS, từ thư mục MCI Excel: Practice/B5_MCI_Mart_Answer.xlsx từ đáp án B4 — 6 relationship, 14 measure,
                                                  # 7 Pivot From Data Model, DoiSoat_B2 (CUBEVALUE), KPI_Dictionary, Kiem_tra_B5 (24 dòng) qua datamodel_com.py. Không Refresh query.
                                                  # Sau đó làm tay Mark as Date Table + Hide from Client Tools, rồi --check. --print-ps: chỉ in mã PowerShell
python3 _generator/build_slides_b5.py             # Slide_B5.pptx (cần python-pptx); in ra các khung [CHỤP] còn thiếu ảnh
python3 _generator/build_student_slides.py B5
```

Sau khi dựng lại, mở `B1_MCI_Mart_Q1_2025_Answer.xlsx` bằng Excel và lưu lại một lần, để các ô công thức có sẵn giá trị khi xem trước. Với B2 làm như vậy cho `B2_MCI_Mart_Q1_2025_Answer.xlsx` và `B2_BTVN_MCI_Mart_T4_2025_Answer.xlsx`. Buổi 3 không cần bước này: `build_b3.py` tự mở và lưu từng file bằng Excel. Lý do: openpyxl không tạo được PivotTable, nên `pivot_xml.py` chỉ ghi phần khai báo (cache rỗng, refresh khi mở) và Excel dựng lại Pivot khi mở file; file được chép vào thư mục sandbox của Excel (`~/Library/Containers/com.microsoft.Excel/Data/Documents`) rồi mới mở, tránh Excel lấy bản cũ trên OneDrive. Bản PDF của handout xuất bằng Microsoft Word (Save As PDF).
