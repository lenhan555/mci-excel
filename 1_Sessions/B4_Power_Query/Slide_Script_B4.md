# Kịch bản slide — Buổi 4: Tiền xử lý dữ liệu với Power Query

> **Phiên bản 1.4 (2026-10-01):** soạn theo `Course_Plan.md` v3.3, khuôn mẫu `Slide_Script_B3.md`. Lời giảng viết thành **văn nói đầy đủ**. Mọi con số tính từ chính các file `Practice/` bằng `2_Dataset/_generator/build_b4_homework.py` (lưu ở `build_b4_meta.json`) và khớp file đáp án dựng bằng Power Query thật (`build_b4.py`, Excel qua COM trên Windows).
> Căn cứ: `Course_Plan.md` v3.3 (B4: timetable 150 phút, Hướng dẫn giảng dạy B4, Nguyên tắc 1/6/7/9/10, Phụ lục A chuỗi A và B, Phụ lục F câu chuyện #3 và #4).
>
> **Thay đổi so với 1.2:** thao tác Power Query viết lại cho đúng giao diện Excel 365 (CSV mở ra có 3 bước tự động; Query Editor không có Undo, muốn bỏ thì xóa bước; bước Changed Type tự động đoán kiểu theo locale máy nên phải xóa rồi Using Locale); lỗi Locale tái hiện được trên mọi máy bằng Using Locale → English (United States); đối soát ngay trong Power Query bằng Count Rows và Statistics → Sum; làm sạch lại trên query gộp từ Folder thay vì trong Transform Sample File; Merge minh họa bằng **Merge Queries as New** (không đụng vào Sales); Target thêm cột Thang dạng số và `Ngay_Dau_Thang` để B5 nối với Calendar; **Calendar 01/01/2024–31/12/2025 = 731 dòng** (phủ trọn hai năm, vì Target có đủ 12 tháng 2025; Course Plan Phụ lục F ghi "Calendar 2 năm"); Pivot `PV_Refresh` đổi thành Nam × Thang (số dòng, doanh thu) và lọc Store_Code, để đối soát từng tháng và thấy ST12; tắt background refresh để Refresh All cập nhật Pivot trong một lần bấm. Bài tập về nhà, handout, infographic đã dựng. File đáp án `Practice/B4_MCI_Mart_Answer.xlsx` đã dựng lại trên Windows theo bản 1.3 (30/09/2026, `Kiem_tra_B4` 10/10 ✔). Đã chụp đủ 19 ảnh deck + 9 ảnh handout, dựng lại deck, bản `_HocVien`, handout và xuất PDF (30/09/2026). **01/10/2026:** file đáp án dựng lại cho khớp đúng thao tác trên lớp: Sales có 4 query trợ giúp của Combine (nhóm `Transform File from Data_Sales` / `Helper Queries`, Transform File gắn với Transform Sample File), tham số `DataFolder` dùng ở cả Sales và Sample File; Calendar theo Add Column → Date → … → Name of Day (cột Thu ra Monday… trên máy tiếng Anh); `Kiem_tra_B4` 10/10 ✔; tên bước của Sales và Target theo đúng thao tác trên lớp (Using Locale hai bước, `Added Custom`); chụp lại H08, S14b, S18a, S18b (chỉ hiện các query đã dạy); quiz 6 đổi thành tình huống "lỗi ngầm" (đã thử trên Excel: file mới đổi tên cột thì Refresh không báo lỗi, cột để trống); slide 12 tách 991 / 927 dòng; câu Locale ở slide 25. **02/10/2026:** thêm **file mốc** `Practice/B4_MCI_Mart_Moc_ThucHanh.xlsx` (`build_b4.py --moc`, chỉ Windows): trạng thái cuối slide 27, để học viên có file lỗi hoặc chưa làm kịp vẫn vào được Thực hành 1 (ghi chú giảng viên slide 28, bảng file, bảng lỗi thường gặp).

## Cách dùng kịch bản

- **Lời giảng** là lời nói mẫu. Giảng viên không cần đọc nguyên văn; nên nói theo giọng của mình nhưng giữ đủ **ý chính** và **câu hỏi cho lớp** (in đậm).
- Dòng *(→ …)* trong lời giảng là **chỉ dẫn hành động**. Đường dẫn file tính từ thư mục buổi này (`1_Sessions/B4_Power_Query/`); file khởi đầu nằm ở `../B3_PivotTable/Practice/`.
- Mỗi slide có thời lượng dự kiến. Nếu lớp chậm, cắt ở những chỗ đánh dấu **[có thể rút gọn]**.
- **Buổi này không có file `Demo/`.** Power Query là xây pipeline trên chính dữ liệu thật, nên mọi demo chạy nguyên trạng trên các file `Practice/` (CSV bẩn, `B4_Lookups.xlsx`, `B4_Target_2025.xlsx`) — đúng quy tắc dùng lại file đã có trong `CLAUDE.md`. Giảng viên và học viên cùng làm trên **bản sao file khởi đầu** của mình (`B4_<HoTen>.xlsx`), nên mọi con số trên màn hình giống nhau.
- **File mốc khi học viên bị tụt lại.** Cả buổi là một chuỗi liền trên một file: hỏng một query ở slide 12 hay 21 là không còn gì để làm Thực hành 1. `Practice/B4_MCI_Mart_Moc_ThucHanh.xlsx` là trạng thái **cuối slide 27** (đủ 6 bảng, `PV_Refresh`, Sales 28.328 dòng, chưa có tháng 4), mở ra ở sheet `Huong_dan_Moc`: cách dùng và bảng tự kiểm. Học viên phải sửa tham số `DataFolder` về thư mục `Practice` trên máy mình rồi Refresh All; bước này chưa dạy trên lớp (handout mục 8), nên trợ giảng làm giúp. Không dùng file đáp án thay file mốc: đáp án đã có tháng 4 và sheet `Kiem_tra_B4`.
- **File khởi đầu:** bản sao `../B3_PivotTable/Practice/B3_MCI_Mart_Q1_2025_Answer.xlsx`, lưu thành `B4_<HoTen>.xlsx`. Các sheet tblSales, KPI_SUMIFS, PV_… của B3 **giữ nguyên**; query và bảng mới hôm nay được **thêm vào** file này (B5 sẽ đối soát measure với KPI_SUMIFS ngay trong cùng file).
- **Query dựng trong buổi:** `Sales_T1` (query tập, một file tháng 1; Phần 2–4), `DoiSoat_T1` (Reference của Sales_T1, Group By), `Product` · `Store` · `Customer` (Phần 3), `Target` (Phần 4), `Sales` (From Folder, Phần 5), `Sales_Merge_Demo` (Phần 5, xóa ngay sau khi xem), `Calendar` (Phần 6). Phần 7 xóa các query tập, còn đúng 6 bảng: **Sales, Product, Store, Customer, Target, Calendar** (cộng nhóm query trợ giúp mà From Folder tự tạo).
- **Pivot "tự cập nhật"** là Pivot mới `PV_Refresh` dựng trên bảng Sales do Power Query nạp ra (Phần 7), không phải các Pivot cũ của B3 (nguồn của chúng là tblSales Quý 1, không đọc thư mục `Data_Sales`). Nói rõ điều này để học viên không đi tìm Pivot B3 sau khi Refresh.
- **Locale.** Ngày trong mọi file CSV là dd/mm/yyyy, giá các file 2025 viết kiểu Việt Nam ("19.000"). Máy để Region **English (United States)** thì bước Changed Type tự động đã gây lỗi ngay khi mở file; máy để **Vietnamese** thì không thấy lỗi nào. Kịch bản tái hiện lỗi trên **mọi máy** bằng Change Type → Using Locale → English (United States), nên không phụ thuộc máy trình chiếu. Kiểm tra Region của máy trình chiếu trước giờ dạy để biết học viên sẽ thấy gì khi mở file.
- **Trước giờ dạy:** thư mục `Practice/Data_Sales/` phải có **đúng 4 file** (`Sales_2024.csv`, `Sales_2025-01/02/03.csv`). File tháng 4 nằm ở `Practice/New_Month/` và chỉ được chép vào lúc Thực hành 1. Dạy xong lớp trước, xóa `Sales_2025-04.csv` khỏi `Data_Sales/` trên máy trình chiếu.

## Quy ước

Loại slide và nguồn ảnh giống `Slide_Script_B3.md`: BÌA · CHIA PHẦN · KHÁI NIỆM · CÔNG CỤ · THỰC HÀNH · QUIZ · TỔNG KẾT; **[CHỤP]** ảnh chụp Excel 365 Windows trên file của khóa, **[PQ PDF pX]** trang PDF nguồn (`4_Reference/PowerQuery_PowerPivot_DAX.pdf`, chỉ để trong kịch bản, không ghi lên slide), **[VẼ]** vẽ bằng shape, chỉ cho sơ đồ khái niệm và thẻ số (theo `CLAUDE.md`, mục Ảnh minh họa).

**File dùng trong buổi** (đường dẫn tính từ thư mục buổi; `../B3_PivotTable/` là thư mục Buổi 3):

| File | Công dụng | Slide |
|---|---|---|
| `../B3_PivotTable/Practice/B3_MCI_Mart_Q1_2025_Answer.xlsx` | **File khởi đầu** (đáp án B3): tblSales Quý 1/2025 (6.187 dòng), KPI_SUMIFS, 6 Pivot, `Kiem_tra` 14/14 ✔. Lưu thành `B4_<HoTen>.xlsx`, làm mọi việc hôm nay trong file này | 1, 3, 9–29 |
| `../B3_PivotTable/Homework/B3_BTVN_MCI_Mart_T1_T4_2025_Answer.xlsx` | Đáp án bài về nhà B3, chữa nhanh đầu buổi | 3 |
| `Practice/Data_Sales/Sales_2025-01.csv` | Tháng 1/2025, 2.726 dòng thô: giá "19.000", ngày dd/mm/yyyy, 25 mã cửa hàng dính khoảng trắng, 1 dòng trùng. Nguồn của query tập `Sales_T1` | 4, 9, 11–15, 19 |
| `Practice/Data_Sales/Sales_2025-02.csv`, `Sales_2025-03.csv` | Tháng 2, 3 (1.596 / 1.868 dòng thô, cùng dạng lỗi) | 4, 21 |
| `Practice/Data_Sales/` (cả thư mục: `Sales_2024.csv` 22.141 dòng + 3 file 2025) | Nguồn của **From Folder** → query `Sales` | 21, 22, 27 |
| `Practice/B4_Lookups.xlsx` (bảng Product, Store, Customer) | Product 41 dòng (**P015 trùng**): Remove Duplicates, Merge minh họa; Store, Customer load riêng | 14, 22, 23, 27 |
| `Practice/B4_Target_2025.xlsx` (sheet Target_2025) | Target dạng cột Tháng 1…Tháng 12, tiêu đề ở dòng 3: Unpivot | 18, 19 |
| `Practice/New_Month/Sales_2025-04.csv` | Tháng 4/2025 (1.920 dòng thô, cùng dạng lỗi): chép vào `Data_Sales/` ở Thực hành 1 | 28, 29 |
| `Practice/B4_MCI_Mart_Moc_ThucHanh.xlsx` | **File mốc** (`build_b4.py --moc`, Power Query thật): giống đáp án nhưng ở trạng thái **cuối slide 27** — Data_Sales 4 file, Sales 28.328 dòng · 4.139.884.500, 5 bảng Connection Only + Data Model, `PV_Refresh`; không có `Kiem_tra_B4`. Sheet `Huong_dan_Moc` (đứng đầu): 4 bước dùng file (Save As, sửa `DataFolder`, Refresh All, làm Thực hành 1) và bảng tự kiểm 10 dòng. Chỉ phát cho học viên có file lỗi hoặc chưa làm kịp | 28 |
| `Practice/B4_MCI_Mart_Answer.xlsx` | **Đáp án B4** (`build_b4.py`, Power Query thật): 6 query + tham số `DataFolder` + 4 query trợ giúp của Combine (nhóm `Transform File from Data_Sales`; `DataFolder` dùng ở cả Sales và Sample File), Sales (Table `Sales` trên sheet `Sales_PQ` + Data Model, đã gồm tháng 4), 5 bảng Connection Only + Data Model, Pivot `PV_Refresh`, sheet `Kiem_tra_B4`. Giữ nguyên các sheet B3. **File khởi đầu Buổi 5** | 29, 39 |
| `Homework/B4_BTVN_MCI_Mart_T4_2025.xlsx` + thư mục `Homework/B4_BTVN_Du_lieu/` (`Cua_hang/` 12 CSV, `Bo_sung/` 1 CSV) | **Bài tập về nhà bắt buộc** (đề): tháng 4/2025 mỗi cửa hàng một file; sheet De_bai, So_lieu_he_thong, Product, Target_2025, Insight | 39 |
| `Homework/B4_BTVN_MCI_Mart_T4_2025_Answer.xlsx` | Đáp án bài về nhà: bảng kết quả Sales_T4, KetQua_T4, mã M (Dap_an_M), `Kiem_tra` (11/11 ✔), Dap_an_loi, Insight_mau. Chỉ giảng viên / trợ giảng | — |
| `Handout/B4_Handout.docx` / `.pdf` | Handout tự học (11 mục) | 39 |
| `Handout/B4_Infographic_TomTat.png` | Infographic tóm tắt buổi (nguồn `_generator/infographic_b4.html`) | 39 |

**Số liệu dùng trong buổi** (`build_b4_meta.json`; khớp `Kiem_tra_B4` của file đáp án):

| Chỉ tiêu | Giá trị |
|---|---|
| `Sales_2025-01.csv` | 2.726 dòng thô → **2.725** sau Remove Duplicates; doanh thu **425.850.000** |
| Order_Date tháng 1 có ngày ≥ 13 (thành Error nếu đọc kiểu mm/dd) | **1.735 / 2.726** dòng; 991 dòng còn lại không báo lỗi, trong đó **927 dòng bị đảo ngày–tháng** (64 dòng ngày 01/01 đọc kiểu nào cũng đúng) |
| Unit_Price tháng 1 đọc kiểu Mỹ (dấu chấm là thập phân) | từ **10 đến 650** (đúng: 10.000 đến 650.000) — không báo lỗi |
| Store_Code tháng 1 trước Trim | **23 giá trị khác nhau** (12 mã + 11 mã có dấu cách cuối); sau Trim còn 12. Mỗi file 2025 có 25 dòng dính dấu cách |
| Group By tháng 1 theo Store_Code | 12 dòng; cao nhất ST08 345 dòng / 52.832.000, thấp nhất ST04 136 dòng / 18.991.000 |
| `B4_Lookups.xlsx` | Product **41 → 40** dòng (P015 trùng); Store 12; Customer 1.500 |
| Merge Sales × Product **chưa** bỏ P015 trùng | **29.408** dòng thay vì 28.328 (1.080 dòng P015 bị nhân đôi) |
| Merge Left Anti Sales × Product | **0** dòng (mọi mã sản phẩm hợp lệ) |
| Target sau Unpivot | 12 cửa hàng × 12 tháng = **144** dòng; cả năm 3.665.000.000 |
| Target tháng 1 / tháng 2 (12 cửa hàng) | **353.000.000** (thực tế 425.850.000 = 120,6%) / **298.000.000** (thực tế 245.996.000 = 82,5%) |
| `Sales` sau From Folder (2024 + Quý 1/2025, đã làm sạch) | **28.328** dòng, doanh thu **4.139.884.500** (Quý 1/2025: 6.187 dòng, 941.454.000 = KPI_SUMIFS B2) |
| `New_Month/Sales_2025-04.csv` | 1.920 thô → **1.919** dòng; **296.965.000** (khớp đáp án bài về nhà B2) |
| Sau khi thả tháng 4 và Refresh | **30.247** dòng, **4.436.849.500** (khớp `MCI_Mart_Master.xlsx`) |
| Calendar 01/01/2024–31/12/2025 | **731** dòng (2024 nhuận 366 + 2025 365). Mùng 1 Tết: 10/02/2024 và 29/01/2025 |
| ST12 (câu chuyện #3) | 0 dòng trước 06/2024; tháng 6/2024 có **75** dòng (khai trương 01/06/2024) |

Bảng đối soát theo tháng (Pivot `PV_Refresh`, Nam × Thang):

| Tháng | Số dòng | Doanh thu | Tháng | Số dòng | Doanh thu |
|---|---|---|---|---|---|
| 01/2024 | 2.097 | 298.203.000 | 09/2024 | 1.857 | 271.399.500 |
| 02/2024 | 1.632 | 248.304.000 | 10/2024 | 1.913 | 276.707.000 |
| 03/2024 | 1.693 | 237.192.000 | 11/2024 | 1.929 | 276.029.000 |
| 04/2024 | 1.639 | 239.466.000 | 12/2024 | 2.278 | 322.354.500 |
| 05/2024 | 1.608 | 239.810.000 | 01/2025 | 2.725 | 425.850.000 |
| 06/2024 | 1.774 | 249.780.000 | 02/2025 | 1.595 | 245.996.000 |
| 07/2024 | 1.768 | 250.600.000 | 03/2025 | 1.867 | 269.608.000 |
| 08/2024 | 1.953 | 288.585.500 | 04/2025 (sau Refresh) | 1.919 | 296.965.000 |

**Tổng:** 39 slide, 150 phút.

---

## Phần 0 — Mở đầu: từ Pivot đẹp đến nỗi đau ba file CSV (10 phút · slide 1–4)

### Slide 1 · BÌA · 0.5'
- **Trên slide:** MCI / Consulting & Analytics / **Excel for Business Intelligence Analyst** / Buổi 4: Tiền xử lý dữ liệu với Power Query / "Bắt buộc máy Windows có Power Query (Data → Get Data)" / "Mở sẵn: file khởi đầu B3_MCI_Mart_Q1_2025_Answer.xlsx (đáp án Buổi 3)"
- **Ảnh:** [VẼ] nền theo Style spec.

**Lời giảng:**
> Chào cả lớp, chào mừng đến Buổi 4. Ba buổi trước mình đã có một bảng dữ liệu sạch, một bộ KPI, và một báo cáo Pivot rất đẹp. Nhưng có một điều bí mật: dữ liệu đó **đã được ai đó dọn sẵn**. Hôm nay mình học công cụ để **tự dọn**, và dọn đúng một lần.

### Slide 2 · KHÁI NIỆM · Câu hỏi mở buổi · 2'
- **Trên slide:** *"Làm sao để việc làm sạch dữ liệu chỉ làm **một lần**, từ đó chỉ cần bấm Refresh?"* · khung giải thích: "Tháng nào công ty cũng có file bán hàng mới. Copy-paste, sửa ngày, sửa mã cửa hàng bằng tay mỗi tháng là công việc không có hồi kết — càng làm nhiều lần càng dễ sai."
- **Ảnh:** [VẼ] nút Refresh lớn.

**Lời giảng:**
> Tháng nào công ty cũng có file bán hàng mới. Nếu tháng nào cũng copy-paste, sửa ngày, sửa mã cửa hàng bằng tay, thì đây là công việc **không có hồi kết**, và càng làm nhiều lần càng dễ sai. Câu hỏi hôm nay: làm sao dọn dữ liệu **một lần**, ghi lại các bước đó, để tháng sau chỉ cần bấm một nút?
>
> Buổi này không có tầng câu hỏi riêng trên thang 5 tầng. Nó là **nền móng**: nếu dữ liệu không sạch và không tự cập nhật, mọi con số ở các tầng sau đều không đáng tin.

### Slide 3 · KHÁI NIỆM · Recap B3 & file khởi đầu · 3'
- **Trên slide:**
  - Viên gạch B3: Pivot Miền × Tháng (tổng **941.454.000**), Slicer nối 3 Pivot, PivotChart theo ngày.
  - Bài về nhà B3 (4 tháng đầu năm): đối soát Pivot với số liệu hệ thống, **6 ô Quantity dạng chữ** cài sẵn tháng 4, câu hỏi "ST07 đã hồi phục chưa".
  - Hôm nay nhận **file khởi đầu** = đáp án B3 (`B3_MCI_Mart_Q1_2025_Answer.xlsx`), lưu thành `B4_<HoTen>.xlsx`.
- **Ảnh:** [VẼ] viên gạch B3 (đã xong) nằm dưới viên gạch B4 (sáng lên).

**Lời giảng:**
> Nhắc nhanh buổi trước. *(→ mở `../B3_PivotTable/Practice/B3_MCI_Mart_Q1_2025_Answer.xlsx`, sheet Kiem_tra)* Đây là báo cáo Pivot hoàn chỉnh: tổng doanh thu Quý 1 941 triệu 454 nghìn, 14 dòng đối soát đều tick xanh.
>
> Bài về nhà Buổi 3 các bạn vừa nộp: dữ liệu 4 tháng, có sẵn 6 ô Quantity dạng chữ, và câu hỏi ST07 đã hồi phục sau đợt cải tạo tháng 3 chưa. *(→ nếu còn thời gian: mở `../B3_PivotTable/Homework/B3_BTVN_MCI_Mart_T1_T4_2025_Answer.xlsx`, sheet Kiem_tra, chữa nhanh câu 1)*
>
> File khởi đầu hôm nay vẫn là đáp án Buổi 3. Các bạn lưu thành `B4_<HoTen>.xlsx`, **giữ nguyên** mọi sheet cũ. Hôm nay mình chỉ **thêm vào**, không xóa gì cả.

### Slide 4 · KHÁI NIỆM · "Nỗi đau" — làm tay ba file CSV · 4.5'
- **Trên slide:**
  - 3 file: `Sales_2025-01.csv` (2.726 dòng) · `Sales_2025-02.csv` (1.596 dòng) · `Sales_2025-03.csv` (1.868 dòng).
  - Việc phải làm bằng tay: mở từng file → copy → dán nối tiếp → sửa 25 mã cửa hàng dính dấu cách mỗi file → sửa ngày dd/mm và giá "19.000" → tìm và xóa dòng trùng → làm lại **mỗi tháng**.
- **Ảnh:** [CHỤP] ba cửa sổ Excel mở đồng thời `Sales_2025-01.csv`, `-02.csv`, `-03.csv`, xếp cạnh nhau.
- **Chân slide:** Ảnh: Practice/Data_Sales/Sales_2025-01.csv, -02.csv, -03.csv mở bằng Excel.

**Lời giảng:**
> *(→ mở lần lượt `Practice/Data_Sales/Sales_2025-01.csv`, `Sales_2025-02.csv`, `Sales_2025-03.csv` bằng Excel, View → Arrange All → Vertical)* Ba file, ba tháng, tổng gần sáu nghìn hai trăm dòng. Nếu làm tay: mở từng file, copy, dán nối tiếp, rồi sửa. Sửa gì? *(→ trong `Sales_2025-01.csv`, Ctrl+F tìm `ST07 ` có dấu cách)* 25 dòng mỗi file có mã cửa hàng dính khoảng trắng thừa ở cuối — nhìn bằng mắt không thấy, nhưng máy tính coi `"ST07 "` khác `"ST07"`. *(→ trỏ cột Order_Date và Unit_Price)* Ngày viết kiểu Việt Nam ngày/tháng/năm, giá viết "19.000" có dấu chấm ngăn nghìn. Máy các bạn để Region tiếng Anh thì Excel đã đọc sai ngay lúc mở: giá thành 19, ngày 13 trở đi thành chữ. Và mỗi file có đúng 1 dòng bị trùng hoàn toàn.
>
> **Câu hỏi cho lớp: nếu tháng sau có file thứ tư, các bạn phải làm gì?** *(→ chờ)* Làm lại đúng như vậy, từ đầu. Đây là giới hạn Buổi 3 đã nói tới: dữ liệu sạch không tự nhiên mà có. Buổi 1 các bạn đã dọn tay một file Quý 1 bẩn; hôm nay Power Query ghi lại các bước dọn đó **một lần**, tháng sau chỉ bấm Refresh. Cuối buổi các bạn sẽ tự tay chứng kiến. *(→ đóng ba file, không lưu)*

---

## Phần 1 — Quy trình "Power Excel" (10 phút · slide 5–7)

### Slide 5 · CHIA PHẦN · 0.5'
- **Trên slide:** MCI / **1. Quy trình Power Excel: từ dữ liệu thô đến báo cáo**

**Lời giảng:**
> Phần 1: nhìn toàn cảnh quy trình, trước khi bấm vào chi tiết.

### Slide 6 · KHÁI NIỆM · Bốn khối: Raw Data → Power Query → Data Model → Power Pivot · 5'
- **Trên slide:** 4 khối nối tiếp [PQ PDF p11]:
  1. **Raw Data:** file CSV/TXT, Excel Table, database, cả một **thư mục** file, web…
  2. **Power Query** (Get & Transform): kết nối, làm sạch, biến đổi — ETL (Extract – Transform – Load).
  3. **Data Model:** các bảng đã sạch, nối bằng quan hệ, không cần gộp thành một bảng khổng lồ (Buổi 5).
  4. **Power Pivot & Pivot:** phân tích, viết DAX (Buổi 5–6).
- **Ảnh:** [VẼ] 4 khối, mũi tên nối, khối 2 (Power Query) sáng lên.
- **TIP:** Tất cả nằm sẵn trong Excel (tab **Data**), không cần cài thêm phần mềm.

**Lời giảng:**
> Đây là bức tranh lớn. Dữ liệu thô ở bất cứ đâu — file, thư mục, database — đi qua **Power Query** để làm sạch và biến đổi. Sạch rồi thì vào **Data Model**, nơi các bảng được nối với nhau bằng quan hệ, không phải bằng XLOOKUP hay gộp thành một bảng khổng lồ. Cuối cùng, **Power Pivot** đọc từ Data Model để phân tích và viết công thức DAX.
>
> Hôm nay mình ở khối thứ hai: **Power Query**. Buổi 5 mới sang khối ba và bốn. Nhớ thứ tự: sạch trước, nối sau, phân tích cuối cùng.

### Slide 7 · KHÁI NIỆM · Khi nào dùng Power Query / Power Pivot · 4.5'
- **Trên slide:** Dùng Power Query/Power Pivot khi [PQ PDF p12–17]:
  - Dữ liệu **nhiều hơn một triệu dòng** (vượt giới hạn worksheet).
  - Dữ liệu nằm ở **nhiều nguồn** (nhiều file, database, thư mục).
  - Cần **tự động hoá**: mỗi lần có dữ liệu mới, chỉ Refresh.
  - Cần kết hợp nhiều bảng mà không muốn nhân bản dữ liệu.
- **Thẻ số:** Worksheet tối đa **1.048.576** dòng · Data Model: hàng chục triệu dòng, giới hạn bởi bộ nhớ máy (nên dùng Excel 64-bit).
- **TIP:** Power Pivot **không chạy trên Mac** và Excel Online — lý do B1 đã nhắc yêu cầu máy Windows từ B4. Chưa chuẩn bị được: gặp trợ giảng ngay giờ nghỉ.

**Lời giảng:**
> Vì sao không cứ SUMIFS và Pivot mà dùng? Vì ba buổi trước dữ liệu của các bạn đã sạch, và chỉ có một file. Ngoài đời, dữ liệu thật thường là **nhiều file, mỗi tháng một file, có khi hàng triệu dòng**. Một worksheet chỉ chứa được khoảng một triệu dòng; Data Model thì chứa hàng chục triệu dòng, giới hạn chỉ còn là bộ nhớ máy.
>
> Và đây là lý do Buổi 1 mình đã nhắc: **Power Pivot không chạy trên Mac**. Bạn nào dùng Mac mà chưa có máy Windows hoặc máy ảo, đây là hạn chót — gặp trợ giảng ngay giờ nghỉ.

---

## Phần 2 — Get Data, Query Editor, Applied Steps (10 phút · slide 8–9)

### Slide 8 · CHIA PHẦN · 0.5'
- **Trên slide:** MCI / **2. Get Data & Query Editor: nơi mọi bước làm sạch được ghi lại**

**Lời giảng:**
> Phần 2: mở cánh cửa vào Power Query, và làm quen với màn hình Query Editor.

### Slide 9 · CÔNG CỤ · Get Data · Query Editor · Applied Steps · 9.5'

| Tên (đỏ) | Mục đích |
|---|---|
| **Data → Get Data → From File → From Text/CSV** | Kết nối tới một file; bấm **Transform Data** để mở Query Editor trước khi nạp |
| **Query Editor** (Power Query Editor) | Cửa sổ riêng: bảng xem trước ở giữa, **Applied Steps** bên phải, 4 tab Home / Transform / Add Column / View |
| **Applied Steps** | Danh sách các bước đã làm, chạy lại **tự động** mỗi lần Refresh — như một macro ghi sẵn |

- **Ảnh:** [CHỤP] Query Editor đầy đủ của query `Sales_T1` (`Practice/Data_Sales/Sales_2025-01.csv`): bảng xem trước, thanh công thức, Applied Steps có 3 bước tự động Source · Promoted Headers · Changed Type; [PQ PDF p21–22].
- **TIP:** Query Editor **không có Undo**. Muốn bỏ một bước: bấm dấu **X** cạnh bước đó. Bấm vào một bước ở giữa để xem dữ liệu **tại thời điểm** bước đó.
- **PRO TIP:** Đặt tên query rõ ràng **ngay từ đầu** (ô Name ở khung Query Settings) — đổi tên sau khi query khác đã tham chiếu tới sẽ rất phiền [PQ PDF p37].
- **Chân slide:** Ảnh: Practice/Data_Sales/Sales_2025-01.csv → Data → Get Data → From Text/CSV → Transform Data (query Sales_T1).

**Lời giảng:**
> *(→ trên `B4_<HoTen>.xlsx`: Data → Get Data → From File → From Text/CSV → chọn `Practice/Data_Sales/Sales_2025-01.csv`)* Đây là cửa vào Power Query. Hộp thoại cho xem trước dữ liệu, rồi có hai nút: **Load** nạp thẳng, hoặc **Transform Data** để mở Query Editor trước. Mình luôn chọn Transform Data, vì mình cần sửa trước khi nạp.
>
> *(→ bấm Transform Data; ô Name bên phải đổi `Sales_2025-01` thành `Sales_T1`)* Đây là **Query Editor**, một cửa sổ riêng. Giữa màn hình là **bảng xem trước**. Bên phải là ô quan trọng nhất buổi hôm nay: **Applied Steps**. Power Query đã tự làm ba bước: `Source` (mở file), `Promoted Headers` (lấy dòng 1 làm tên cột) và `Changed Type` (tự đoán kiểu dữ liệu mỗi cột). *(→ bấm lần lượt vào từng bước, chỉ thanh công thức phía trên bảng)* Mỗi bước là một dòng công thức, viết bằng ngôn ngữ **M**. Bấm vào bước nào, bảng hiện dữ liệu tại thời điểm đó.
>
> **Câu hỏi cho lớp: bước Changed Type do máy tự đoán. Máy đoán dựa vào đâu?** *(→ chờ)* Dựa vào 200 dòng đầu và **định dạng vùng miền của máy đang chạy**. Nghĩa là cùng một file, máy tiếng Anh và máy tiếng Việt có thể đoán khác nhau. Giữ câu hỏi này, lát nữa ở Phần 3 mình sẽ thấy nó quan trọng thế nào.
>
> Buổi 2 các bạn viết công thức trong ô; hôm nay Power Query ghi công thức thành một **danh sách bước**, và chạy lại toàn bộ danh sách đó mỗi khi bấm Refresh. Một lưu ý: Query Editor **không có Undo**. Muốn bỏ một bước thì bấm dấu X cạnh bước đó.

---

## Phần 3 — Data profiling, làm sạch, Remove Duplicates, Keep/Remove Errors (25 phút · slide 10–15)

### Slide 10 · CHIA PHẦN · 0.5'
- **Trên slide:** MCI / **3. Làm sạch dữ liệu: nhìn trước khi sửa**

**Lời giảng:**
> Phần dài nhất buổi. Trước khi sửa, mình phải **nhìn thấy** dữ liệu đang sai ở đâu — đó là data profiling.

### Slide 11 · CÔNG CỤ · Data profiling: Column quality · Distribution · Profile · 4.5'
- **Tên (đỏ):** **View → Column quality / Column distribution / Column profile**
- **Mục đích:** Xem nhanh sức khỏe từng cột: % **Valid** / **Error** / **Empty**; số giá trị **khác nhau** (distinct); Min, Max; các giá trị hay gặp — trước khi mất công sửa mò.
- **Ảnh:** [CHỤP] Column quality và Column profile của `Sales_T1` (bước Changed Type tự động, profiling trên toàn bộ dữ liệu): Unit_Price 100% Valid nhưng Min 10, Max 650 (máy English US). [CHỤP] Column distribution của Store_Code trước Trim: **23 distinct** thay vì 12.
- **TIP:** Góc dưới trái màn hình ghi "Column profiling based on top 1000 rows" — bấm vào, chọn **entire data set** để đếm trên cả file.
- **Chân slide:** Ảnh: Query Editor, query Sales_T1 (Practice/Data_Sales/Sales_2025-01.csv), bước Changed Type tự động.

**Lời giảng:**
> *(→ trong Query Editor của `Sales_T1`: View → tick Column quality, Column distribution, Column profile; bấm dòng "Column profiling based on top 1000 rows" ở góc dưới → chọn entire data set)* Ba công cụ này là bản đồ sức khỏe dữ liệu. Ngay dưới tên mỗi cột là một vạch: xanh là **Valid**, đỏ là **Error**, xám là **Empty**.
>
> *(→ bấm tiêu đề cột Store_Code, nhìn Column distribution và khung Column statistics phía dưới)* Nhìn Store_Code: đáng lẽ chỉ có 12 cửa hàng, ST01 đến ST12, nhưng profiling đếm ra **23 giá trị khác nhau**. Vì `"ST07"` và `"ST07 "` có dấu cách bị đếm là hai. Đây là cách phát hiện lỗi **trước khi** đọc từng dòng bằng mắt.
>
> *(→ bấm cột Unit_Price)* Máy nào để Region tiếng Anh sẽ thấy cột giá 100% xanh — không lỗi nào. Nhưng nhìn Min, Max: **từ 10 đến 650**. Một sản phẩm giá 10 đồng? *(→ bấm cột Order_Date)* Và Order_Date có hơn nửa cột đỏ. Máy để tiếng Việt thì các cột này đều đúng. Đó chính là bước Changed Type tự đoán theo máy mà mình vừa hỏi. Slide sau mình sửa tận gốc.
>
> **Câu hỏi cho lớp: profiling giúp gì mà Ctrl+F không giúp được?** *(→ chờ)* Ctrl+F phải biết trước mình tìm gì. Profiling cho biết cột nào có vấn đề **mà chưa cần đoán trước lỗi**. Và "100% Valid" chưa chắc là đúng — phải nhìn cả Min, Max. Đây là thói quen D2: hoài nghi dữ liệu, nhìn tổng quan trước, sửa sau.

### Slide 12 · CÔNG CỤ · Change Type · Using Locale · 6'
- **Tên (đỏ):** **Change Type → Using Locale…**
- **Mục đích:** Đổi kiểu dữ liệu **theo đúng định dạng vùng miền của dữ liệu nguồn** — không phải định dạng của máy đang chạy Excel.
- **Ảnh:** [CHỤP] hộp Change Type with Locale: Data Type = Date, Locale = **Vietnamese (Vietnam)**, cột Order_Date của `Sales_T1` 100% Valid. [CHỤP] Unit_Price sau Using Locale → Whole Number, Vietnamese: 19.000 → **19000**.
- **Thẻ số:** Đọc Order_Date kiểu Mỹ (mm/dd): **1.735 / 2.726** dòng Error (ngày ≥ 13); 991 dòng còn lại không báo lỗi, trong đó **927 dòng bị đảo ngày–tháng**. Đọc Unit_Price kiểu Mỹ: **19.000 → 19**, sai 1.000 lần, không báo lỗi.
- **TIP:** Chuột phải tiêu đề cột → Change Type → **Using Locale…** (hoặc Transform → Data Type ▾ → Using Locale). Quy tắc: **xóa bước Changed Type tự động, rồi đổi kiểu cột ngày và cột số bằng Using Locale.**
- **Chân slide:** Ảnh: Query Editor, query Sales_T1 (Practice/Data_Sales/Sales_2025-01.csv), cột Order_Date và Unit_Price.

**Lời giảng:**
> Đây là lỗi các bạn **chắc chắn gặp** khi làm việc với dữ liệu Việt Nam. Trước hết, *(→ trong Applied Steps của `Sales_T1`: bấm X cạnh bước Changed Type)* mình xóa bước Changed Type tự động. Mọi cột trở về dạng chữ, biểu tượng ABC trên tiêu đề cột. Giờ mình tự đổi kiểu, có kiểm soát.
>
> Để cả lớp cùng thấy lỗi, dù máy để tiếng Anh hay tiếng Việt, mình cố tình đọc theo kiểu Mỹ: *(→ chuột phải Order_Date → Change Type → Using Locale → Data Type: Date, Locale: English (United States) → OK)* Nhìn Column quality: cột đỏ gần hai phần ba. Đúng **1.735 trên 2.726 dòng** lỗi, vì kiểu Mỹ đọc tháng trước ngày sau, mà không có tháng 13. Còn nguy hiểm hơn: *(→ trỏ một dòng ngày 02/01/2025)* ngày 2 tháng 1 bị đọc thành **1 tháng 2** — không báo lỗi gì cả. 927 dòng như thế.
>
> *(→ bấm X xóa bước vừa tạo; chuột phải Order_Date → Change Type → Using Locale → Date, Vietnamese (Vietnam) → OK)* Nói cho Excel biết dữ liệu này viết kiểu Việt Nam, ngày trước tháng sau. Cột xanh 100%, và ngày 2 tháng 1 đúng là ngày 2 tháng 1.
>
> Cột giá: *(→ chuột phải Unit_Price → Change Type → Using Locale → Whole Number, English (United States))* kiểu Mỹ coi dấu chấm là dấu thập phân: `19.000` thành **19**. Không một dấu đỏ nào. **Đây là lỗi nguy hiểm nhất buổi: sai 1.000 lần mà không có gì báo.** *(→ xóa bước; Using Locale → Whole Number, Vietnamese (Vietnam))* Đúng Locale thì `19.000` thành **19000**. *(→ chọn Line_No và Quantity, Ctrl, chuột phải → Change Type → Whole Number)* Hai cột số nguyên không có dấu ngăn nghìn thì đổi kiểu thường là đủ.
>
> Quy tắc nhớ đời: **xóa Changed Type tự động; mọi cột ngày và cột số của dữ liệu Việt Nam, đổi kiểu bằng Using Locale.** Làm vậy thì file chạy đúng trên máy nào cũng được, kể cả máy đồng nghiệp để tiếng Anh.

### Slide 13 · CÔNG CỤ · Use First Row as Headers · Trim · Conditional Column · 5'

| Tên (đỏ) | Mục đích |
|---|---|
| **Home → Use First Row as Headers** | Biến dòng đầu của bảng thành tên cột |
| **Transform → Format → Trim / Clean** | Xóa khoảng trắng đầu/cuối (Trim) hoặc ký tự ẩn (Clean) |
| **Add Column → Conditional Column** | Tạo cột mới theo quy tắc IF/THEN, không cần viết hàm |

- **Ảnh:** [CHỤP] Store_Code sau Trim: Column distribution còn đúng **12** giá trị. [CHỤP] hộp Add Conditional Column: Nhom_SL — If Quantity is greater than 3 then "Nhiều" else "Ít"; [PQ PDF p24–25, 31].
- **TIP:** Trim/Clean có ở cả **Transform** (sửa cột đang có) và **Add Column** (tạo cột mới): chọn đúng tab theo việc mình muốn.
- **Chân slide:** Ảnh: Query Editor, query Sales_T1 (Practice/Data_Sales/Sales_2025-01.csv).

**Lời giảng:**
> Với file CSV có tên cột ở dòng 1, bước **Promoted Headers** Power Query đã tự làm — mình vừa thấy trong Applied Steps. Khi nào phải bấm tay **Use First Row as Headers**? Khi file có dòng tiêu đề báo cáo phía trên, như bảng Target sau giờ nghỉ.
>
> *(→ chuột phải cột Store_Code → Transform → Trim)* 25 dòng có khoảng trắng thừa giờ đã sạch. *(→ nhìn lại Column distribution / Column statistics)* Đếm lại: đúng **12 giá trị**. Bước này quan trọng hơn vẻ ngoài: khoảng trắng thừa sẽ làm **Merge không khớp** ở Phần 5 — hai mã trông giống hệt nhau trên màn hình nhưng máy không khớp được.
>
> **[có thể rút gọn]** *(→ Add Column → Conditional Column; New column name: Nhom_SL; If Quantity is greater than 3 then "Nhiều", Else "Ít" → OK)* Conditional Column là IF/THEN không cần gõ công thức: chọn cột, điều kiện, kết quả. Buổi 2 các bạn viết `=IF(...)`; đây là bản kéo-thả của chính hàm đó, áp cho cả cột mỗi lần Refresh.

### Slide 14 · CÔNG CỤ · Remove Duplicates · 4.5'
- **Tên (đỏ):** **Home → Remove Rows → Remove Duplicates**
- **Mục đích:** Xóa các dòng **giống hệt nhau trên các cột đang chọn** — dùng cho bảng giao dịch (dòng trùng tuyệt đối) lẫn bảng lookup (khóa trùng).
- **Ảnh:** [CHỤP] `Sales_T1` sau Remove Duplicates (chọn mọi cột) và **Transform → Count Rows** = **2.725** (trước là 2.726). [CHỤP] query `Product` (`Practice/B4_Lookups.xlsx`): chọn Product_Code → Remove Duplicates → Count Rows = **40** (trước là 41).
- **Thẻ số:** Merge Sales với Product **chưa** bỏ P015 trùng: **29.408** dòng thay vì 28.328 — 1.080 dòng P015 bị nhân đôi.
- **TIP:** Đếm dòng ngay trong Power Query: **Transform → Count Rows**, xem số, rồi xóa bước đó. Remove Duplicates chỉ xét **các cột đang chọn**: chọn cả bảng (Ctrl+A) cho bảng giao dịch, chỉ cột khóa cho bảng lookup.
- **Chân slide:** Ảnh: query Sales_T1 (Practice/Data_Sales/Sales_2025-01.csv) · query Product (Practice/B4_Lookups.xlsx, bảng Product).

**Lời giảng:**
> Hai cách dùng Remove Duplicates, phải phân biệt rõ. *(→ trên `Sales_T1`: bấm vào bảng xem trước, Ctrl+A chọn mọi cột → Home → Remove Rows → Remove Duplicates)* Cách một: xóa dòng **trùng tuyệt đối** trong bảng giao dịch. Số dòng? Góc dưới chỉ ghi "999+ rows", nên mình đếm: *(→ Transform → Count Rows)* **2.725** — bớt đúng 1 dòng so với 2.726. *(→ bấm X xóa bước Counted Rows)* Đếm xong nhớ xóa bước đếm, kẻo query chỉ còn một con số.
>
> *(→ trong Query Editor: Home → New Source → File → Excel Workbook → `Practice/B4_Lookups.xlsx` → tick "Select multiple items" → tick 3 bảng Product, Store, Customer → OK)* Cách hai, quan trọng hơn cho Buổi 5. Mình lấy luôn ba bảng lookup một lượt. Bảng Product có 41 dòng nhưng chỉ 40 sản phẩm — **P015 bị trùng**. *(→ trên query Product: chọn cột Product_Code → Remove Duplicates → Count Rows = 40 → xóa bước đếm)* Còn 40.
>
> **Câu hỏi cho lớp: nếu bỏ qua bước này, lát nữa Merge Sales với Product sẽ ra sao?** *(→ chờ)* Mỗi dòng Sales mua P015 khớp với **cả hai** dòng P015, nên bị nhân đôi: bảng Sales 28.328 dòng sẽ phình thành **29.408** dòng. Doanh thu cũng bị cộng thừa. Quy tắc: **Remove Duplicates ở lookup trước khi Merge hoặc trước khi tạo Relationship** — Buổi 5 dùng lại nguyên vẹn.

### Slide 15 · CÔNG CỤ · Keep Errors · Remove Errors · 4.5'
- **Tên (đỏ):** **Home → Keep Rows → Keep Errors** / **Home → Remove Rows → Remove Errors**
- **Mục đích:** Tách riêng các dòng đang có ô lỗi (`Error`) để **xem trước khi quyết định xóa hay sửa** — không xóa mù.
- **Ảnh:** [CHỤP] bản sao `Sales_T1 (2)`: Order_Date đọc bằng Locale English (United States) rồi Keep Errors còn **1.735** dòng; bấm một ô Error, khung chi tiết hiện lý do (không đọc được ngày).
- **TIP:** **Keep Errors** để hiểu lỗi trước. **Remove Errors** chỉ dùng khi chắc chắn dòng đó nên bỏ. Lỗi do sai Locale thì sửa Locale (Slide 12), **không** Remove Errors — xóa là mất 1.735 dòng doanh thu mà không ai biết.
- **Chân slide:** Ảnh: bản sao query Sales_T1 (Practice/Data_Sales/Sales_2025-01.csv), Order_Date đọc bằng Locale English (United States).

**Lời giảng:**
> Tình huống thật: đồng nghiệp gửi file đã đổi kiểu sai, cột ngày có hơn một nghìn bảy trăm ô lỗi. Làm gì? Mình dựng lại tình huống đó trên một bản sao, để không làm hỏng query chính. *(→ khung Queries bên trái: chuột phải `Sales_T1` → Duplicate; trong `Sales_T1 (2)`: chuột phải bước ngay sau Promoted Headers → Delete Until End (xóa bước đó và mọi bước sau, còn lại Source và Promoted Headers); chuột phải Order_Date → Change Type → Using Locale → Date, English (United States))* Cột ngày lại đỏ.
>
> *(→ Home → Keep Rows → Keep Errors → Count Rows = 1.735 → xóa bước đếm)* Bảng chỉ còn **1.735** dòng lỗi, để xem chúng là gì. *(→ bấm vào khoảng trắng cạnh chữ Error trong một ô)* Khung chi tiết cho biết: không đọc được chuỗi ngày theo định dạng đang chọn. Nhìn các ngày: toàn ngày 13 trở đi.
>
> **Xử lý đúng ở đây là Remove Errors hay sửa lại kiểu dữ liệu?** *(→ chờ)* Sửa lại Locale, như Slide 12. Remove Errors ở đây là xóa mất 1.735 dòng doanh thu — tổng tháng 1 hụt hơn nửa mà không ai biết. **Remove Errors chỉ dùng khi dòng đó thật sự nên bỏ**, ví dụ vài dòng rác không sửa được. Lỗi hiện ra là **cơ hội tìm nguyên nhân gốc**. *(→ chuột phải `Sales_T1 (2)` → Delete)*

---

### Slide 16 · NGHỈ · 10'
- **Trên slide:** Nghỉ 10 phút. Trước khi quay lại: mở sẵn Query Editor (Data → Get Data → Launch Power Query Editor).

**Lời giảng:**
> Nghỉ 10 phút. Nửa sau buổi mình xoay bảng kế hoạch, gộp hai năm dữ liệu bằng From Folder, nối bảng bằng Merge, và dựng Calendar.
>
> *(→ trong giờ nghỉ: đi vòng kiểm tra `Sales_T1` của học viên đã xóa Changed Type tự động, Order_Date và Unit_Price đổi bằng Using Locale Vietnamese, đã Trim Store_Code, Remove Duplicates; Product còn 40 dòng)*

---

## Phần 4 — Unpivot Target; Group By (15 phút · slide 17–19)

### Slide 17 · CHIA PHẦN · 0.5'
- **Trên slide:** MCI / **4. Unpivot & Group By: đổi hình dạng bảng**

**Lời giảng:**
> Phần 4: xoay bảng kế hoạch từ nằm ngang thành đứng dọc — đúng tư duy tidy data của Buổi 1.

### Slide 18 · CÔNG CỤ · Unpivot Columns · 8.5'
- **Tên (đỏ):** **Transform → Unpivot Columns ▾ → Unpivot Other Columns**
- **Mục đích:** Biến **nhiều cột cùng loại** (Tháng 1 … Tháng 12) thành **hai cột**: tên cột cũ (Attribute) và giá trị (Value) — ngược lại của Pivot.
- **Ảnh:** [CHỤP] `Practice/B4_Target_2025.xlsx` trong Query Editor trước Unpivot: tiêu đề báo cáo ở dòng 1, tên cột ở dòng 3, Tháng 1…Tháng 12 nằm ngang. [CHỤP] sau Unpivot và làm gọn: Store_Code · Thang (số) · Target · Ngay_Dau_Thang, **144** dòng.
- **Thẻ số:** 12 cửa hàng × 12 tháng = **144** dòng.
- **TIP:** Chọn **các cột giữ nguyên** (Store_Code, Store_Name) rồi **Unpivot Other Columns**: tháng mới thêm vào file vẫn tự được xoay, không phải sửa query. Chuỗi A lần 4: Table (B1) → Dimension/Measure (B2) → Pivot (B3) → **Unpivot (B4)** → Star schema (B5).
- **Chân slide:** Ảnh: Practice/B4_Target_2025.xlsx, sheet Target_2025 (query Target).

**Lời giảng:**
> *(→ Home → New Source → File → Excel Workbook → `Practice/B4_Target_2025.xlsx` → chọn sheet Target_2025 → OK; đổi tên query thành `Target`)* Bảng kế hoạch nằm ngang: mỗi cửa hàng một dòng, mỗi tháng một cột. Đẹp để in, nhưng **không tidy**: Buổi 1 mình học mỗi dòng một bản ghi, mỗi cột một biến — ở đây "tháng" đang nằm trong **tên cột**, không phải giá trị.
>
> Dòng 1 là tiêu đề báo cáo, tên cột thật ở dòng 3. *(→ nếu Applied Steps có bước Promoted Headers / Changed Type tự động: xóa cả hai; Home → Remove Rows → Remove Top Rows → 2; Home → Use First Row as Headers)* Đây là lúc dùng tay Use First Row as Headers.
>
> *(→ chọn cột Store_Code, giữ Ctrl chọn Store_Name → Transform → Unpivot Columns ▾ → Unpivot Other Columns)* Chọn hai cột **giữ nguyên**, rồi xoay mọi cột còn lại. *(→ đổi tên: Attribute → Thang, Value → Target)* Cột Thang đang là chữ "Tháng 1". *(→ chọn Thang → Transform → Extract → Text After Delimiter → dấu cách → OK; đổi kiểu Whole Number; Target → Whole Number; chuột phải Store_Name → Remove)* Lấy phần số, bỏ Store_Name vì tên cửa hàng đã có ở bảng Store. *(→ Add Column → Custom Column: tên `Ngay_Dau_Thang`, công thức `#date(2025, [Thang], 1)`, Data type: Date → OK)* Cột ngày đầu tháng này để Buổi 5 nối Target với Calendar. *(→ Transform → Count Rows = 144 → xóa bước đếm)* **144 dòng**: 12 cửa hàng nhân 12 tháng.
>
> **Câu hỏi cho lớp: nếu để nguyên dạng cột-tháng, Buổi 5 có làm được Pivot "Doanh thu so với Target theo tháng" không?** *(→ chờ)* Rất khó, vì Pivot cần "tháng" nằm trong **một cột giá trị** để lọc và nối, không phải rải ra 12 cột. Đây là Chuỗi A, lần thứ tư.

### Slide 19 · CÔNG CỤ · Group By · 6'
- **Tên (đỏ):** **Transform → Group By** (Basic / Advanced)
- **Mục đích:** Tổng hợp dữ liệu **lên mức chi tiết khác** — từ dòng giao dịch lên theo cửa hàng, theo tháng — ngay trong Power Query, ra một **bảng mới**.
- **Ảnh:** [CHỤP] hộp Group By (Advanced) trên `DoiSoat_T1`: Group by Store_Code; So_dong = Count Rows; Doanh_thu = Sum of Revenue; kết quả 12 dòng.
- **Thẻ số:** Target tháng 1 **353.000.000** (thực tế 425.850.000, 120,6%) · tháng 2 **298.000.000** (thực tế 245.996.000, 82,5%).
- **PRO TIP:** **Reference** tạo query mới đọc kết quả của query gốc, gốc đổi thì nó đổi theo; **Duplicate** chép rời toàn bộ các bước. Bảng đối soát, bảng tổng hợp: dùng Reference.
- **Chân slide:** Ảnh: query DoiSoat_T1 (Reference của Sales_T1, Practice/Data_Sales/Sales_2025-01.csv).

**Lời giảng:**
> Nhìn nhanh tổng Target theo tháng. *(→ trên query Target: Transform → Group By → Basic; Group by Thang; New column name Tong_Target, Operation Sum, Column Target → OK)* Group By kiểu Basic: một cột nhóm, một phép tính. 12 dòng. Tháng 1 kế hoạch **353 triệu**, thực tế 425,85 triệu — vượt khoảng 21%. Tháng 2 kế hoạch **298 triệu**, thực tế chỉ 246 triệu — hụt khoảng 17%. Cùng một công ty, hai tháng liền kề, một tháng vượt xa, một tháng hụt. Vì sao? Nhớ hai con số này, cuối buổi mình quay lại. *(→ bấm X xóa bước Grouped Rows: Target trở lại 144 dòng)*
>
> Giờ dùng Group By để **đối soát**. *(→ trên `Sales_T1`: Add Column → Custom Column: tên Revenue, công thức `[Quantity] * [Unit_Price]`, Data type: Whole Number → OK)* Thêm cột doanh thu. *(→ khung Queries: chuột phải `Sales_T1` → Reference; đổi tên query mới thành `DoiSoat_T1`)* **Reference** tạo một query mới đọc kết quả của Sales_T1 — Sales_T1 đổi thì nó đổi theo. *(→ Transform → Group By → Advanced; Group by Store_Code; So_dong = Count Rows; Add aggregation: Doanh_thu = Sum of Revenue → OK)* 12 dòng: mỗi cửa hàng một số dòng và một tổng doanh thu tháng 1. Cao nhất ST08, 345 dòng, 52,8 triệu. Cộng lại đúng 2.725 dòng, 425.850.000. Đây là bảng đối soát: có số liệu hệ thống thì so từng cửa hàng **trước khi** nạp dữ liệu đi đâu cả.
>
> Khác Pivot: Group By tạo một **bảng cố định**, không kéo thả được nữa. Dùng khi cần một bảng tổng hợp để đối soát, để Merge, hoặc làm nguồn cho biểu đồ — ví dụ bảng doanh thu theo ngày cho Forecast ở Buổi 6.

---

## Phần 5 — Append từ Folder; Merge (20 phút · slide 20–23)

### Slide 20 · CHIA PHẦN · 0.5'
- **Trên slide:** MCI / **5. Append & Merge: nối dọc và nối ngang**

**Lời giảng:**
> Phần 5, trái tim của buổi hôm nay: gộp nhiều file bằng From Folder, và nối bảng bằng Merge.

### Slide 21 · CÔNG CỤ · From Folder · Append Queries · 8'

| Tên (đỏ) | Mục đích |
|---|---|
| **Data → Get Data → From File → From Folder** | Đọc và nối **mọi file cùng cấu trúc** trong một thư mục; file mới thả vào thư mục tự được đọc khi Refresh |
| **Home → Append Queries** | Nối dọc hai (hoặc nhiều) query có **cùng tên cột** |

- **Ảnh:** [CHỤP] hộp From Folder trỏ `Practice/Data_Sales/`: danh sách 4 file và nút **Combine → Combine & Transform Data**; [PQ PDF p34, 36].
- **Thẻ số:** Sau Combine và làm sạch: **28.328** dòng · doanh thu **4.139.884.500** (Count Rows; Statistics → Sum).
- **TIP:** Combine tạo nhóm query trợ giúp (Sample File, Transform File…) — **đừng xóa**. Làm sạch trên query gộp `Sales`: mỗi bước viết một lần, áp cho mọi file, kể cả file thả vào sau này.
- **PRO TIP:** Tham số hóa đường dẫn thư mục (Home → Manage Parameters) để chuyển máy chỉ sửa một chỗ — xem handout.
- **Chân slide:** Ảnh: Data → Get Data → From File → From Folder → Practice/Data_Sales/ · đếm: Transform → Count Rows; tổng: Statistics → Sum.

**Lời giảng:**
> Giờ làm đúng việc của Slide 4, nhưng cho cả thư mục. *(→ Home → New Source → File → Folder → chọn `Practice/Data_Sales/` → OK)* Power Query liệt kê đủ 4 file: năm 2024 và ba tháng 2025. *(→ Combine ▾ → Combine & Transform Data → File mẫu: First file → OK; đổi tên query `Data_Sales` thành `Sales`)* Power Query nối cả 4 file thành một bảng. Bên trái có thêm một nhóm query trợ giúp — đừng xóa, chúng là "máy đọc từng file".
>
> Bảng gộp này vẫn bẩn y như Sales_T1 lúc đầu. Mình làm sạch lại — lần này nhanh vì các bạn đã biết từng bước, và đây là **lần cuối cùng**: *(→ trên `Sales`: bấm X xóa bước Changed Type cuối cùng; chuột phải cột Source.Name → Remove; Store_Code → Transform → Trim; Order_Date → Using Locale → Date, Vietnamese; Unit_Price → Using Locale → Whole Number, Vietnamese; Line_No, Quantity → Whole Number; Order_ID, Store_Code, Customer_ID, Product_Code → Text; Ctrl+A → Remove Duplicates; Add Column → Custom Column Revenue = `[Quantity] * [Unit_Price]`, Whole Number; chọn Order_Date → Add Column → Date → Year → Year, đổi tên Nam; Add Column → Date → Month → Month, đổi tên Thang)*
>
> Đối soát ngay tại đây. *(→ Transform → Count Rows)* **28.328 dòng**: 22.141 của năm 2024 cộng 2.725, 1.595, 1.867 của ba tháng 2025. *(→ xóa bước đếm; chọn cột Revenue → Transform → Statistics → Sum)* Doanh thu **4.139.884.500**. *(→ xóa bước Sum)* Ghi hai con số này lại: lát nữa thả file tháng 4 vào, mình so trước và sau.
>
> **Append Queries** là nối dọc các query đã có, ví dụ Sales_T1 với một query tháng 2. From Folder làm việc đó cho cả thư mục, **và tự đón file mới**. Có một cái bẫy: nếu nối file CSV mà không lấy dòng đầu làm tên cột cho từng file, tên cột của file sau — chữ "Order_ID", "Store_Code" — sẽ lẫn vào giữa dữ liệu. Combine & Transform đã lo việc đó trong query trợ giúp; nếu gặp, lọc bỏ dòng có Order_ID bằng đúng chữ "Order_ID".

### Slide 22 · CÔNG CỤ · Merge Queries · 7.5'
- **Tên (đỏ):** **Home → Merge Queries ▾ → Merge Queries as New**
- **Mục đích:** Nối **ngang** hai bảng theo cột chung — đúng như XLOOKUP nhưng làm ở **cấp cả bảng**, không phải từng ô.
- **Ảnh:** [CHỤP] hộp Merge: Sales (Product_Code) × Product (Product_Code), Join Kind Left Outer, dòng xác nhận "The selection matches 28328 of 28328 rows". [CHỤP] `Sales_Merge_Demo` sau Expand: thêm cột Category, không dòng null; [PQ PDF p35].
- **Thẻ số:** Sau Merge: **28.328** dòng — không đổi, vì Product đã bỏ P015 trùng (nếu chưa bỏ: 29.408).
- **TIP:** Merge ở đây **chỉ để minh họa** "XLOOKUP cấp bảng". Dùng **Merge Queries as New** để không đụng vào Sales, xem xong xóa query minh họa. Product, Store, Customer **load riêng**, Buổi 5 nối bằng Relationship — không gộp thành một bảng khổng lồ.
- **Chân slide:** Ảnh: query Sales_Merge_Demo (Sales × Product, Practice/B4_Lookups.xlsx).

**Lời giảng:**
> Buổi 2 các bạn lấy Category từ bảng Product sang Sales bằng XLOOKUP, từng ô một. Merge làm việc đó cho **cả bảng cùng lúc**. *(→ chọn query `Sales` → Home → Merge Queries ▾ → Merge Queries as New)* Mình chọn "as New" để tạo một query riêng, không đụng vào Sales.
>
> *(→ bảng trên: Sales, bấm cột Product_Code; bảng dưới: Product, bấm cột Product_Code; Join Kind: Left Outer → nhìn dòng xác nhận ở dưới)* Chọn cột khóa ở hai bảng, chọn kiểu nối. Power Query báo ngay: khớp 28.328 trên 28.328 dòng. *(→ nếu hiện dải vàng "Information is required about data privacy": Continue → đặt cả hai nguồn là Organizational → Save; giải thích ở handout mục 9)* Hai bảng đến từ hai nguồn khác nhau — thư mục CSV và file Excel — nên lần đầu nối, Power Query có thể hỏi mức riêng tư của từng nguồn. *(→ OK; đổi tên query thành `Sales_Merge_Demo`; bấm nút hai mũi tên ở tiêu đề cột Product → chỉ tick Category, bỏ tick "Use original column name as prefix" → OK)* Mỗi dòng Sales giờ có thêm ngành hàng. *(→ Transform → Count Rows)* Vẫn **28.328 dòng** — nhờ Product đã bỏ P015 trùng. Chưa bỏ thì ra 29.408.
>
> **Câu hỏi cho lớp: số dòng tăng sau Merge nghĩa là gì?** *(→ chờ)* Nghĩa là bảng bên phải có **khóa trùng**. Đó là lý do phải đối soát số dòng sau mỗi bước: số dòng tăng sau Merge thì dừng lại, kiểm tra lookup ngay.
>
> Nhưng nhớ TIP quan trọng nhất slide này: mình Merge chỉ để **hiểu cơ chế**. *(→ chuột phải `Sales_Merge_Demo` → Delete)* Trong file thật, Product, Store, Customer đứng **riêng**. Gộp Category, Region, Segment… vào Sales thì mỗi dòng trong 28 nghìn dòng lặp lại cùng một thông tin. Buổi 5 dùng Relationship để nối mà không nhân bản dữ liệu.

### Slide 23 · KHÁI NIỆM · Các kiểu Join (Merge) · 4'
- **Trên slide:** 4 kiểu Merge hay dùng [PQ PDF p35]:

| Join Kind | Kết quả | Giống |
|---|---|---|
| **Left Outer** (mặc định) | Giữ hết bảng trái, thêm cột nếu khớp bảng phải | XLOOKUP (không thấy → null) |
| **Inner** | Chỉ giữ dòng **khớp ở cả hai** bảng | INNER JOIN (SQL) |
| **Right Outer** | Giữ hết bảng phải, thêm cột từ bảng trái nếu khớp | Left Outer đổi chiều |
| **Left Anti** | Chỉ giữ dòng bảng trái **không** khớp bảng phải | Tìm "dòng mồ côi" |

- **Ảnh:** [VẼ] 4 hình hai vòng tròn giao nhau, phần tô đậm là phần giữ lại.
- **TIP:** Kiểm tra chất lượng bằng **Left Anti**: Sales × Product ra **0** dòng nghĩa là mọi Product_Code trong Sales đều có trong danh mục.

**Lời giảng:**
> Merge có nhiều kiểu nối. *(→ chỉ bảng trên slide)* **Left Outer**, mình vừa dùng, giữ nguyên toàn bộ Sales. **Inner** chỉ giữ dòng khớp ở cả hai bên. **Right Outer** ngược lại với Left. Và một kiểu đặc biệt: **Left Anti** — chỉ giữ dòng **không** khớp.
>
> **Câu hỏi cho lớp: muốn kiểm tra có đơn hàng nào ghi mã sản phẩm sai, không có trong danh mục Product, dùng Join Kind nào?** *(→ chờ)* Left Anti. Merge Sales với Product bằng Left Anti: ra **0 dòng** nghĩa là mọi mã đều hợp lệ; ra vài dòng thì đó chính là các dòng cần sửa. Trên dữ liệu hôm nay kết quả là 0. Bài tập về nhà các bạn sẽ gặp trường hợp không phải 0.

---

## Phần 6 — Bảng Calendar (10 phút · slide 24–25)

### Slide 24 · CHIA PHẦN · 0.5'
- **Trên slide:** MCI / **6. Bảng Calendar: trục thời gian cho mọi bảng**

**Lời giảng:**
> Phần 6, ngắn nhưng rất quan trọng cho Buổi 5: dựng một bảng ngày tháng độc lập.

### Slide 25 · CÔNG CỤ · Blank Query → List.Dates → Calendar · 9.5'
- **Tên (đỏ):** **Home → New Source → Other Sources → Blank Query** · `List.Dates(...)` · **Add Column → Date**
- **Mục đích:** Tạo bảng có **đủ mọi ngày liên tục** trong các năm cần phân tích, kèm Năm / Tháng / Quý / Thứ — trục thời gian chung cho Sales và Target ở Buổi 5.
- **Công thức (trên slide):** `= List.Dates(#date(2024,1,1), 731, #duration(1,0,0,0))`
- **Ảnh:** [CHỤP] query `Calendar` sau To Table và Add Column → Date: cột Date, Nam, Thang, Quy, Thu; [PQ PDF p27–28].
- **Thẻ số:** 01/01/2024 – 31/12/2025: **731** dòng (2024 nhuận 366 ngày + 2025 365 ngày), không thiếu ngày nào.
- **TIP:** Calendar phải **liên tục** và **phủ trọn các năm** của mọi bảng fact (Target có đủ 12 tháng 2025). Không lấy từ cột Order_Date của Sales: ngày không có đơn sẽ bị thiếu.
- **PRO TIP:** Rolling calendar: điểm cuối tự lấy theo ngày hiện tại bằng `DateTime.LocalNow()` — xem handout.
- **Chân slide:** Ảnh: query Calendar (Blank Query) trong B4_<HoTen>.xlsx.

**Lời giảng:**
> *(→ Home → New Source → Other Sources → Blank Query; đổi tên thành `Calendar`)* Bảng Calendar không lấy từ file nào — mình **dựng bằng công thức**. *(→ gõ vào thanh công thức: `= List.Dates(#date(2024,1,1), 731, #duration(1,0,0,0))` → Enter)* Nghĩa là: bắt đầu từ 01/01/2024, sinh 731 giá trị, mỗi giá trị cách nhau 1 ngày. Vì sao 731? Năm 2024 là năm nhuận, 366 ngày, cộng 365 ngày của 2025.
>
> **Câu hỏi cho lớp: vì sao không lấy Calendar từ cột Order_Date của Sales cho nhanh?** *(→ chờ)* Vì ngày nào không có đơn hàng thì ngày đó **biến mất** khỏi Order_Date — Calendar bị thủng lỗ, và các phép lũy kế, so cùng kỳ ở Buổi 5 sai ngay từ những ngày trống. Và vì sao tới tận cuối năm 2025, khi dữ liệu bán hàng mới tới tháng 4? Vì **Target có đủ 12 tháng 2025**. Calendar phải phủ trọn các năm của mọi bảng cần phân tích.
>
> *(→ Transform → To Table → OK; đổi tên Column1 thành Date, kiểu Date; chọn Date → Add Column → Date → Year → Year (đổi tên Nam); Month → Month (Thang); Quarter → Quarter of Year (Quy); Day → Name of Day (Thu); Count Rows = 731 → xóa bước đếm)* Vài cú nhấp, mình có Năm, Tháng, Quý, Thứ cho một trục thời gian dùng chung. Tên thứ hiện theo Locale của workbook, mặc định lấy theo Region của máy lúc tạo query: máy tiếng Anh ra Monday, máy tiếng Việt (hoặc workbook đặt Locale Vietnamese trong Query Options) ra Thứ Hai.
>
> Một điều nhìn thấy ngay khi có Calendar hai năm: *(→ cuộn tới dòng 10/02/2024 và 29/01/2025 của query Calendar, hoặc chỉ trên slide)* Mùng 1 Tết năm 2024 là **10/02**, năm 2025 là **29/01**. Tết năm ngoái rơi vào tháng 2, năm nay rơi vào tháng 1. Nhớ lại hai con số Target tháng 1 và tháng 2 ở Phần 4 — mình sẽ nối hai điều này ở phần tổng kết.

---

## Phần 7 — Load options & đối soát (10 phút · slide 26–27)

### Slide 26 · CHIA PHẦN · 0.5'
- **Trên slide:** MCI / **7. Load: đưa các bảng sạch vào đúng chỗ**

**Lời giảng:**
> Phần 7: quyết định mỗi bảng nạp ra đâu, và chốt lại bằng một bước không thể bỏ qua — đối soát.

### Slide 27 · CÔNG CỤ · Close & Load To · Table / Only Create Connection / Add to Data Model · 9.5'
- **Tên (đỏ):** **Home → Close & Load ▾ → Close & Load To…** · chuột phải query → **Load To…**
- **Mục đích:** Chọn nơi mỗi query đi tới: **Table** (ra sheet), **Only Create Connection** (chỉ lưu các bước, không ra sheet), ☑ **Add this data to the Data Model** (nạp vào Data Model cho Power Pivot).
- **Ảnh:** [CHỤP] khung Queries & Connections: 6 query Sales, Product, Store, Customer, Target, Calendar (và nhóm trợ giúp của From Folder), dòng "… rows loaded" dưới Sales; [PQ PDF p23, 37].
- **Thẻ số (đối soát sau Load):** Sales **28.328** · Product **40** · Store **12** · Customer **1.500** · Target **144** · Calendar **731**.
- **TIP:** Sales → **Table** + Data Model (dựng Pivot ngay). Product, Store, Customer, Target, Calendar → **Only Create Connection** + Data Model, **không** Merge vào Sales. Sales: Properties → bỏ tick **Enable background refresh**, để Refresh All cập nhật cả Pivot trong một lần bấm.
- **Chân slide:** Ảnh: B4_<HoTen>.xlsx, Data → Queries & Connections.

**Lời giảng:**
> Trước khi nạp, dọn các query tập. *(→ khung Queries: chuột phải `DoiSoat_T1` → Delete, rồi chuột phải `Sales_T1` → Delete — xóa query phụ thuộc trước)* Sales_T1 đã hoàn thành nhiệm vụ: query Sales làm lại đúng các bước đó cho cả thư mục. Còn lại 6 bảng: Sales, Product, Store, Customer, Target, Calendar.
>
> *(→ Home → Close & Load ▾ → Close & Load To… → Only Create Connection → OK)* Chọn cách an toàn nhất: chỉ tạo kết nối, chưa bảng nào ra sheet. Giờ chỉnh nơi nạp của từng bảng trong khung Queries & Connections. *(→ Data → Queries & Connections; chuột phải Sales → Load To… → Table, New worksheet, tick Add this data to the Data Model → OK; đổi tên sheet mới thành `Sales_PQ`)* Sales ra **Table** để dựng Pivot ngay, đồng thời vào Data Model cho Buổi 5.
>
> *(→ lần lượt chuột phải Product, Store, Customer, Target, Calendar → Load To… → Only Create Connection, tick Add this data to the Data Model → OK)* Năm bảng còn lại: **chỉ kết nối + Data Model**. Không ra sheet, không Merge vào Sales, nhưng đã nằm sẵn trong Data Model — chuẩn bị cho mô hình hình ngôi sao Buổi 5.
>
> *(→ chuột phải Sales → Properties → bỏ tick Enable background refresh → OK)* Một thiết lập nhỏ nhưng quyết định khoảnh khắc lát nữa: tắt refresh chạy nền, để Refresh All chạy xong query **rồi mới** cập nhật Pivot.
>
> *(→ chọn một ô trong bảng Sales ở sheet Sales_PQ → Insert → PivotTable → New Worksheet, đổi tên sheet `PV_Refresh`; Rows: Nam, Thang; Values: Order_ID (Count, đổi tên So_dong), Revenue (Sum, đổi tên Doanh_thu); Filters: Store_Code)* Pivot đối soát theo tháng. Nhìn khung Queries & Connections: Sales ghi **28.328 rows loaded**. Grand Total của Pivot: **28.328 dòng, 4.139.884.500** — khớp hai con số mình ghi ở Phần 5. Nếu một con số không khớp — ví dụ 28.329 — dừng lại, tìm nguyên nhân trước khi làm tiếp.

---

## Phần 8 — Thực hành: thả tháng 4, bấm Refresh (20 phút · slide 28–29)

### Slide 28 · THỰC HÀNH 1 · Thả file tháng 4 vào Data_Sales (10 phút)
- **Trên slide:**
  1. Ghi lại **trước khi Refresh**: Grand Total của `PV_Refresh` (28.328 dòng · 4.139.884.500).
  2. Copy `Practice/New_Month/Sales_2025-04.csv` vào **đúng thư mục** `Practice/Data_Sales/`.
  3. **Data → Refresh All** (Ctrl+Alt+F5).
  4. Quan sát: `PV_Refresh` có thêm dòng **2025 · 4**, Grand Total tự đổi — **không sửa một bước nào**.
- **Kiểm tra:** Sau Refresh: **30.247 dòng · 4.436.849.500**; dòng tháng 4/2025: **1.919 dòng · 296.965.000**.

**Lời giảng (giao bài):**
> Đây là khoảnh khắc cả buổi hôm nay hướng tới. *(→ học viên tự làm 10 phút)* Trước tiên, ghi lại Grand Total của Pivot PV_Refresh. Sau đó copy file `Sales_2025-04.csv` từ `Practice/New_Month/` vào đúng thư mục `Practice/Data_Sales/` — thư mục mà query Sales của các bạn đang đọc.
>
> Rồi bấm **Refresh All**. Không mở Query Editor. Không sửa gì. Chỉ một nút. Xem điều gì xảy ra với Pivot.

**Ghi chú giảng viên:**
- Lỗi hay gặp nhất: chép nhầm thư mục (vào `New_Month` hoặc thư mục khác), Refresh không đổi gì. Kiểm tra đường dẫn: Data → Queries & Connections → di chuột lên query Sales.
- Học viên có file lỗi hoặc chưa làm kịp tới slide 27: phát `Practice/B4_MCI_Mart_Moc_ThucHanh.xlsx`, mở sheet `Huong_dan_Moc`, làm bước 1–3 (Save As, sửa tham số `DataFolder` về thư mục `Practice` trên máy học viên, Refresh All: bảng tự kiểm toàn ✔), rồi làm Thực hành 1 như cả lớp. Bước sửa `DataFolder` chưa dạy trên lớp, trợ giảng làm giúp.
- Pivot không đổi sau lần Refresh All đầu: chưa tắt Enable background refresh ở Slide 27 — bấm Refresh All lần nữa, rồi tắt thiết lập đó.
- File tháng 4 có cùng dạng lỗi (giá "19.000", ngày dd/mm, 25 mã dính khoảng trắng, 1 dòng trùng). Applied Steps viết theo **tên cột** nên tự làm sạch file mới. Refresh báo lỗi thì xem Applied Steps có bước nào gọi đích danh một **giá trị** cụ thể, hoặc còn sót bước Changed Type tự động.

### Slide 29 · THỰC HÀNH 2 · Đối soát & chữa bài nhanh (10 phút)
- **Trên slide:** Checklist brick output B4:
  - [ ] Sales sau Refresh: **30.247 dòng**, **4.436.849.500** (tăng đúng 1.919 dòng, 296.965.000).
  - [ ] `PV_Refresh` tự có dòng 2025 · 4 mà không sửa bước nào; lọc Store_Code = ST12: dòng đầu tiên là **06/2024**.
  - [ ] Product 40 dòng · Store 12 · Customer 1.500 · Target 144 · Calendar 731.
  - [ ] Product, Store, Customer, Target, Calendar: **Only Create Connection + Data Model**, không Merge vào Sales.
- **Chân slide:** Ghi số dòng và doanh thu trước / sau Refresh vào ô trống cạnh PV_Refresh. Đối chiếu: Practice/B4_MCI_Mart_Answer.xlsx, sheet Kiem_tra_B4.

**Lời giảng:**
> *(→ đi vòng lớp, xem số trên máy học viên)* Con số chuẩn: từ 28.328 lên **30.247** dòng, từ 4.139.884.500 lên **4.436.849.500**. Dòng mới 2025 tháng 4: 1.919 dòng, 296.965.000 — đúng bằng doanh thu tháng 4 các bạn đã tính bằng SUMIFS ở bài về nhà Buổi 2. Số khác thì đừng vội sửa query: kiểm tra đã chép đúng thư mục chưa, đã Refresh All chưa.
>
> *(→ hỏi lớp)* **Bạn nào Refresh xong không phải sửa một bước nào?** Đó là câu trả lời cho câu hỏi mở buổi: dọn dữ liệu một lần, từ đó chỉ bấm Refresh.
>
> *(→ mở `Practice/B4_MCI_Mart_Answer.xlsx`, sheet Kiem_tra_B4)* File đáp án có đủ 10 dòng đối soát. Việc cuối: 5 bảng còn lại phải là **Only Create Connection + Data Model**, đứng riêng, chờ Relationship của Buổi 5 nối lại. Và file đáp án này chính là file khởi đầu Buổi 5.

---

## Phần 9 — Quiz, insight, preview (10 phút · slide 30–39)

Quiz kiểu hỏi nhanh: khoảng 40 giây mỗi câu. **Câu 6–7 là câu tình huống.**

### Slide 30 · QUIZ · EXCEL QUIZ CHALLENGE · 0.2'

**Lời giảng:**
> 7 câu hỏi nhanh. Câu 6 và 7 là tình huống thật giống hôm nay các bạn vừa làm.

### Slide 31 · QUESTION 1: APPLIED STEPS
- Bạn Refresh một query sau khi file nguồn có dữ liệu mới. Các bước làm sạch (Trim, Using Locale, Remove Duplicates…) có phải làm lại bằng tay không? **A) Không, Applied Steps tự chạy lại theo đúng thứ tự đã ghi** · B) Có, mỗi lần Refresh phải làm lại từ đầu · C) Chỉ bước đầu tiên tự chạy lại · D) Chỉ khi mở lại Query Editor

**Lời giảng:**
> A. Applied Steps là danh sách bước được ghi lại và chạy lại tự động mỗi lần Refresh — như một macro.

### Slide 32 · QUESTION 2: LOCALE
- Cột ngày viết `13/01/2025` (dd/mm/yyyy) bị Power Query báo lỗi khi Change Type sang Date. Nguyên nhân khả dĩ nhất? A) File bị hỏng · B) Cột đang là số · **C) Đang đổi kiểu theo locale mặc định (mm/dd), không phải Using Locale kiểu Việt Nam** · D) Thiếu bước Trim

**Lời giảng:**
> C. Ngày 13 không phải tháng hợp lệ theo kiểu Mỹ mm/dd, nên báo lỗi. Sửa bằng Change Type → Using Locale → Vietnamese (Vietnam). Và nhớ: các ngày từ 12 trở xuống không báo lỗi nhưng bị đảo ngày–tháng.

### Slide 33 · QUESTION 3: REMOVE DUPLICATES
- Bảng Product có mã P015 xuất hiện 2 dòng. Nếu Merge Sales với Product mà **chưa** Remove Duplicates, hậu quả là gì? A) Merge báo lỗi ngay · B) Cột Category bị để trống · C) Không ảnh hưởng gì · **D) Số dòng của Sales sau Merge tăng lên (các dòng P015 bị nhân đôi)**

**Lời giảng:**
> D. Mỗi dòng Sales có Product_Code = P015 khớp với cả hai dòng P015 bên Product: 28.328 dòng thành 29.408. Đối soát số dòng sau Merge bắt được lỗi này.

### Slide 34 · QUESTION 4: UNPIVOT
- Bảng `Target_2025` có các cột Tháng 1, Tháng 2 … Tháng 12. Muốn đưa về dạng tidy (Store_Code, Thang, Target), dùng công cụ nào? A) Pivot Column · **B) Unpivot Columns** · C) Transpose · D) Group By

**Lời giảng:**
> B. Unpivot Columns biến nhiều cột cùng loại thành hai cột (tên cột cũ và giá trị). Transpose chỉ lật ngang-dọc cả bảng; Pivot Column làm chiều ngược lại.

### Slide 35 · QUESTION 5: APPEND VS MERGE
- Bốn file CSV theo tháng, **cùng 8 cột**, muốn gộp thành một bảng dài hơn. Nên dùng công cụ nào? **A) Append Queries (hoặc From Folder)** · B) Merge Queries · C) Group By · D) Unpivot

**Lời giảng:**
> A. Cùng cấu trúc cột, nối theo chiều dọc (thêm dòng) là Append; From Folder làm việc đó cho cả thư mục. Merge dùng khi nối theo chiều ngang (thêm cột) dựa trên khóa chung.

### Slide 36 · QUESTION 6: TÌNH HUỐNG — LỖI NGẦM
- Tháng sau, file CSV mới thả vào `Data_Sales` có tên cột `Store_Code` bị đổi thành `StoreCode` (không có dấu gạch dưới). Bạn Refresh All: **không báo lỗi gì**, nhưng bộ lọc Store_Code của Pivot có thêm mục **(blank)**. Nguyên nhân? A) File mới bị hỏng · B) Do chưa Remove Duplicates · C) Do thiếu Using Locale · **D) Combine lấy tên cột theo file mẫu (file đầu tiên); file mới không có cột `Store_Code` nên các dòng của nó để trống, còn cột `StoreCode` bị bỏ đi**

**Lời giảng:**
> D. Bước Expand của query gộp lấy danh sách cột từ **file mẫu** — file đầu tiên trong thư mục, ở đây là Sales_2024. File nào thiếu một cột trong danh sách đó thì cột ấy để trống; cột lạ như `StoreCode` thì bị bỏ. Power Query không báo lỗi gì cả, nên đây là loại sai nguy hiểm nhất: số tổng vẫn đúng, nhưng doanh thu tháng mới không gắn được cửa hàng nào. Cách bắt: **đối soát sau mỗi lần Refresh** — Count Rows theo tháng, lọc Store_Code trống, hoặc nhìn mục (blank) trong bộ lọc Store_Code của Pivot. Cách sửa: tốt nhất là yêu cầu bên xuất file giữ đúng tên cột; nếu không được, mở **Transform Sample File** → Advanced Editor, thêm một bước `Table.RenameColumns(..., {{"StoreCode", "Store_Code"}}, MissingField.Ignore)` (bước ở đó chạy trên từng file trước khi nối — handout mục 7). Phải có `MissingField.Ignore`: file mẫu và các file cũ không có cột `StoreCode`, thiếu nó thì Refresh báo lỗi; vì file mẫu không có cột này nên cũng không bấm Rename trên giao diện được.

### Slide 37 · QUESTION 7: TÌNH HUỐNG — CHỌN CÔNG CỤ
- Bạn nghi ngờ có vài dòng trong Sales ghi `Product_Code` sai chính tả, không khớp bảng Product nào. Cách kiểm tra nhanh nhất? A) Đọc từng dòng bằng mắt · B) Group By theo Product_Code · **C) Merge Sales với Product, Join Kind = Left Anti, xem có dòng nào không** · D) Xóa cột Product_Code

**Lời giảng:**
> C. Left Anti chỉ giữ lại các dòng bên trái không khớp được bên phải — ra 0 dòng nghĩa là mọi mã đều hợp lệ, ra vài dòng thì đó chính là các dòng cần sửa.

### Slide 38 · TỔNG KẾT + INSIGHT · 2'
- **Trên slide:**
  - **Điểm chính cần nhớ:**
    - (1) Applied Steps = công thức làm sạch được ghi lại, tự chạy lại khi Refresh.
    - (2) Xóa Changed Type tự động; ngày và số của dữ liệu Việt Nam đổi kiểu bằng **Using Locale**.
    - (3) Remove Duplicates ở lookup **trước** khi Merge hoặc tạo Relationship.
    - (4) Unpivot đưa bảng cột-tháng về tidy; Append / From Folder nối dọc, Merge nối ngang.
    - (5) Đối soát số dòng và tổng sau **mỗi bước**; lookup load riêng (Only Create Connection + Data Model).
  - **Insight của bạn (1 câu):** "Sau khi gộp và Refresh, ___ *(con số)*, phát hiện ___ *(bất thường)*, nên ___ *(hàm ý)*."
  - Gợi ý: lọc `PV_Refresh` theo từng cửa hàng — có một cửa hàng không có dòng nào suốt 5 tháng đầu 2024.

**Lời giảng:**
> Tổng kết 5 điều cần nhớ *(→ đọc lướt 5 ý trên slide)*.
>
> Giờ viết câu insight, 1 phút. Gợi ý: trên Pivot PV_Refresh, kéo Store_Code ở vùng Filters, chọn lần lượt từng cửa hàng — có một cửa hàng hoàn toàn không có dòng nào suốt 5 tháng đầu năm 2024.
>
> *(→ chờ 1 phút, gọi 2–3 bạn đọc)*
>
> Câu mẫu: "Sau khi gộp `Sales_2024` và Quý 1/2025 rồi thả thêm tháng 4, Sales có 30.247 dòng, 4.436.849.500, khớp đối soát; lọc ST12 thì dòng đầu tiên là **tháng 6/2024 với 75 dòng**, không có dòng nào trước đó — đây là cửa hàng **mới khai trương 01/06/2024**, không phải lỗi mất dữ liệu, nên Buổi 5 khi tính tăng trưởng cùng kỳ phải so riêng ST12 hoặc chỉ so các cửa hàng có đủ hai kỳ."
>
> Để ý: câu này dùng đúng kỹ năng D2 — thấy dòng trống thì tìm nguyên nhân trước khi kết luận "mất dữ liệu". Và một câu chuyện nữa đang chờ: Target tháng 1 là 353 triệu, thực tế 425,85 triệu, vượt 21%; tháng 2 ngược lại, hụt 17%. Calendar vừa cho thấy Tết năm ngoái rơi vào tháng 2, năm nay rơi vào tháng 1. Target được lập thế nào mà lệch như vậy? Buổi 5 mình đo chính xác bằng measure `Revenue vs Target %`.

### Slide 39 · PREVIEW B5 · 1'
- **Trên slide:**
  - *Buổi sau:* "Doanh thu tháng này đạt bao nhiêu % mục tiêu, và vì sao mỗi báo cáo cũ lại ra một con số khác nhau?" (Data Model & DAX)
  - Giới hạn hôm nay: Sales, Product, Store, Customer, Target, Calendar đã sạch nhưng **đứng rời nhau**. Muốn phân tích Sales và Target cùng lúc mà không gộp thành một bảng khổng lồ?
  - *Về nhà:* **Bài tập về nhà (bắt buộc)** `B4_BTVN_MCI_Mart_T4_2025.xlsx` + thư mục `B4_BTVN_Du_lieu`: 12 file CSV tháng 4, mỗi cửa hàng một file → From Folder, làm sạch, Left Anti, đối soát với hệ thống, % đạt target theo cửa hàng. Nộp trước Buổi 5 tại hocvien.mcivietnam.com · đọc handout B4 · tick "Checklist trước Buổi 5".
- **Ảnh:** [VẼ] viên gạch B4 (đã xong) → viên gạch B5 (sáng lên).

**Lời giảng:**
> Hôm nay các bạn đã tự tay biến bốn file CSV bẩn và một bảng kế hoạch nằm ngang thành sáu bảng sạch, ghi lại bằng Applied Steps, và thấy Refresh cập nhật toàn bộ mà không sửa một bước. Đó là năng lực D3 — tự động hóa — ở mức Competent.
>
> Nhưng nhìn lại: sáu bảng đang **đứng riêng**. Muốn biết doanh thu so với target theo cửa hàng, theo tháng, phải nối Sales với Target, với Calendar, với Store — nối bằng Merge thì ra một bảng khổng lồ, trùng lặp ở mọi dòng. Buổi 5 học cách nối đúng: **Relationship trong Data Model**, và viết DAX để mọi báo cáo dùng chung một định nghĩa KPI.
>
> Bài về nhà: tháng 4, hệ thống POS xuất mỗi cửa hàng một file. Các bạn dựng pipeline, làm sạch, và **đối soát với số liệu hệ thống trước khi tính % đạt target**. Có một bất ngờ đang chờ ở bước đối soát. Nộp trước Buổi 5. Cảm ơn cả lớp, hẹn gặp lại!

---

## Phụ lục kịch bản

### Kiểm tra thời lượng

| Phần | Slide | Phút | Timetable B4 |
|---|---|---|---|
| 0 Mở đầu, recap, nỗi đau 3 file CSV | 1–4 | 10 | 10 |
| 1 Quy trình Power Excel | 5–7 | 10 | 10 |
| 2 Get Data, Query Editor, Applied Steps | 8–9 | 10 | 10 |
| 3 Data profiling, Locale, Trim, Remove Duplicates, Keep/Remove Errors | 10–15 | 25 | 25 |
| Nghỉ | 16 | 10 | 10 |
| 4 Unpivot Target, Group By | 17–19 | 15 | 15 |
| 5 From Folder, Merge, các kiểu Join | 20–23 | 20 | 20 |
| 6 Calendar | 24–25 | 10 | 10 |
| 7 Load options, đối soát | 26–27 | 10 | 10 |
| 8 Thực hành: Refresh tháng 4 | 28–29 | 20 | 20 |
| 9 Quiz, insight, preview | 30–39 | 10 | 10 |
| **Tổng** | **39** | **150** | **150** |

Phần 9 cộng từng slide khoảng 8 phút (slide 30: 0,2'; 7 câu × khoảng 40 giây ≈ 4,7'; slide 38: 2'; slide 39: 1'); khoảng 2 phút dự phòng cho câu 6–7 và phần đọc insight.

**Chỗ dễ lố giờ:** Slide 21 (làm sạch lại trên query gộp) và Slide 27 (Load từng bảng + Pivot). Lớp chậm: bỏ Conditional Column (Slide 13), bỏ Group By Basic trên Target (Slide 19, nói miệng hai con số), và dựng sẵn `PV_Refresh` trên máy giảng viên để học viên chỉ cần làm theo.

### Lỗi thường gặp → cách xử lý (Course Plan, Hướng dẫn giảng dạy B4)

| Lỗi | Dấu hiệu | Xử lý | Slide |
|---|---|---|---|
| Bước Changed Type tự động đoán theo locale máy | Ngày ≥ 13 thành Error, giá 19.000 thành 19 | Xóa Changed Type tự động, đổi kiểu bằng Using Locale sau khi đã chuẩn hóa cột | 9, 12 |
| File mới (From Folder) đổi tên hoặc thêm cột | **Không báo lỗi**: cột sai tên để trống ở các dòng của file đó, cột mới bị bỏ (Combine lấy tên cột theo file mẫu); Pivot/bộ lọc có mục (blank) | Đối soát sau mỗi lần Refresh; yêu cầu nguồn giữ tên cột, hoặc đổi tên trong Transform Sample File với `MissingField.Ignore` | 36 |
| Tên cột của từng CSV lọt vào dữ liệu khi nối | Dòng có giá trị "Order_ID", "Store_Code" giữa bảng | Combine & Transform (tự lấy dòng đầu làm tên cột cho từng file); gặp thì lọc bỏ dòng có Order_ID = "Order_ID" | 21 |
| Ngày dd/mm bị hiểu thành mm/dd | Ngày ≥ 13 Error; ngày ≤ 12 đảo ngày–tháng | Change Type → Using Locale → Vietnamese (Vietnam) | 12, 15 |
| Dấu ngăn nghìn / thập phân | Giá nhỏ đi 1.000 lần, không báo lỗi | Using Locale; nhìn Min / Max ở Column profile | 11, 12 |
| Khoảng trắng thừa trong mã | Distinct 23 thay vì 12; Merge không khớp, ra null | Trim / Clean trước khi Merge | 11, 13 |
| Merge với lookup có khóa trùng | Số dòng tăng sau Merge (28.328 → 29.408) | Remove Duplicates ở lookup; đối soát số dòng | 14, 22 |
| Pivot không đổi sau Refresh All | Query đã có dữ liệu mới, Pivot vẫn số cũ | Tắt Enable background refresh của query nguồn; Refresh All lại | 27, 28 |
| Lỗi Formula.Firewall / privacy levels | Không Refresh được khi kết hợp nhiều nguồn | Handout mục 9 | 22 |
| Học viên làm hỏng query hoặc tụt lại | Tới slide 28 chưa có Sales From Folder / `PV_Refresh`, không còn giờ làm lại | File mốc `Practice/B4_MCI_Mart_Moc_ThucHanh.xlsx`: sửa `DataFolder` → Refresh All → làm tiếp Thực hành 1 (sheet `Huong_dan_Moc`) | 28 |

### Danh sách ảnh cần chụp [CHỤP] (Excel 365 Windows, zoom 100%)

**Đã chụp đủ (30/09/2026; chụp lại H08, S14b, S18a, S18b ngày 01/10/2026)**, deck và handout đã dựng lại, không còn khung [CHỤP]. Khi cần chụp lại: lưu vào `Materials/B4/Images/` đúng tên ở cột Tên file (danh sách đầy đủ, gồm ảnh handout: `Materials/B4/Images/_DANH_SACH_ANH.md`), rồi chạy lại `build_slides_b4.py`: ảnh tự chèn vào đúng khung. Chụp trên bản sao file khởi đầu, làm đúng các bước của lời giảng slide tương ứng (hoặc mở Query Editor của `Practice/B4_MCI_Mart_Answer.xlsx` cho các ảnh query Sales, Product, Target, Calendar). Máy chụp nên để Region **English (United States)** để thấy đúng lỗi Locale như lời giảng.

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
| S14b_remove_dup_product.png | Product: Removed Duplicates theo Product_Code, Counted Rows = 40; khung Queries chỉ có các query đã dạy (Sales_T1, Product, Store, Customer) | Query `Product` (`Practice/B4_Lookups.xlsx`, bảng Product) | 14 |
| S15_keep_errors.png | Order_Date đọc bằng Locale English (US), Keep Errors: bảng chỉ còn dòng lỗi, khung chi tiết lỗi của một ô | Query `Sales_T1 (2)` (Duplicate của Sales_T1) | 15 |
| S18a_target_before.png | Target trước Unpivot (Applied Steps chỉ có Source · Navigation): dòng 1 tiêu đề báo cáo, dòng 3 tên cột, Tháng 1…12 nằm ngang | Query `Target` (`Practice/B4_Target_2025.xlsx`), bước Navigation | 18 |
| S18b_target_after.png | Target sau Unpivot và làm gọn: Store_Code, Thang, Target, Ngay_Dau_Thang | Query `Target`, bước cuối | 18 |
| S19_group_by.png | Hộp Group By (Advanced): Store_Code; So_dong = Count Rows; Doanh_thu = Sum of Revenue | Query `DoiSoat_T1` (Reference của Sales_T1) | 19 |
| S21_from_folder.png | Hộp danh sách file của From Folder (4 file) và nút Combine ▾ đang mở | Data → Get Data → From Folder → `Practice/Data_Sales/` | 21 |
| S22a_merge_dialog.png | Hộp Merge: Sales × Product theo Product_Code, Left Outer, "The selection matches 28328 of 28328 rows" | Home → Merge Queries ▾ → Merge Queries as New | 22 |
| S22b_merge_expand.png | Query `Sales_Merge_Demo` sau Expand, cột Category | Query `Sales_Merge_Demo` | 22 |
| S25_calendar.png | Calendar làm theo đúng thao tác của slide (Blank Query → List.Dates → To Table → Add Column → Date → Year / Month / Quarter / Name of Day → đổi tên): cột Date, Nam, Thang, Quy, Thu (máy tiếng Anh: Monday…); Applied Steps từ Source đến Renamed Columns1. Công thức List.Dates in riêng trên slide | Query `Calendar` | 25 |
| S27_queries_pane.png | Khung Queries & Connections: 6 query + nhóm trợ giúp của From Folder, "28,328 rows loaded" dưới Sales | `B4_<HoTen>.xlsx`, Data → Queries & Connections (trước khi thả file tháng 4) | 27 |

### Còn cần làm trước khi dạy

- Chạy thử toàn bộ buổi **trên máy trình chiếu**: Region của máy, thư mục `Data_Sales` có đúng 4 file, Refresh All sau khi chép file tháng 4 cập nhật `PV_Refresh` trong một lần bấm.
- Thử file mốc `Practice/B4_MCI_Mart_Moc_ThucHanh.xlsx` trên một máy khác máy dựng, làm đúng 4 bước ở sheet `Huong_dan_Moc`: sau bước 3 bảng tự kiểm toàn ✔, sau bước 4 `PV_Refresh` có dòng 2025 · 4 (30.247 dòng). Chép sẵn file mốc cho trợ giảng.
- Mở file đáp án trên máy khác máy dựng: sửa tham số `DataFolder` (Data → Get Data → Launch Power Query Editor → Home → Manage Parameters) về thư mục `Practice` của máy đó trước khi Refresh.

### Ghi chú so với Course Plan

- Course Plan (mục Chuẩn bị B4) chỉ liệt kê file `Practice/`, không có file `Demo/`. Kịch bản dạy trực tiếp trên các file này; không cần bảng minh họa tách rời.
- **"Pivot B3 tự cập nhật"** (Course Plan, dòng Thực hành B4): Pivot cũ của B3 đọc tblSales Quý 1, không đọc thư mục `Data_Sales`, nên không tự đổi khi thả file tháng 4. Kịch bản dựng Pivot mới `PV_Refresh` (Nam × Thang, lọc Store_Code) trên bảng Power Query nạp ra — Pivot "tự cập nhật" đúng nghĩa, đồng thời là bảng đối soát "số dòng từng tháng và tổng doanh thu" của mục Kiểm tra cuối buổi.
- **Calendar hai năm trọn vẹn** (01/01/2024–31/12/2025, 731 dòng) theo Phụ lục F ("Calendar 2 năm") và vì Target có đủ 12 tháng 2025. Bản 1.2 dừng ở 30/04/2025 (486 dòng).
- **Target** giữ dạng Tháng | Cửa hàng | Target như Course Plan, thêm `Ngay_Dau_Thang` (ngày 1 của tháng) để Buổi 5 nối Target với Calendar; bỏ Store_Name vì đã có ở bảng Store.
- Câu chuyện #3 (ST12) và #4 (Target lệch mùa Tết) được **hé lộ** ở buổi này (Slide 19, 25, 38) bằng số liệu thô; phần giải thích và đo bằng measure thuộc Buổi 5, đúng Phụ lục F.
