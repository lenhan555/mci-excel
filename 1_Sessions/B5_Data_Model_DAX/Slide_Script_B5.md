# Kịch bản slide — Buổi 5: Mô hình dữ liệu & DAX — KPI dùng chung một nguồn sự thật

> **Phiên bản 1.0 (2026-10-01), bản nháp chờ duyệt.** Soạn theo `Course_Plan.md` v3.3, khuôn mẫu `Slide_Script_B4.md`. Lời giảng viết thành **văn nói đầy đủ**. Mọi con số tính từ chính file khởi đầu (`../B4_Power_Query/Practice/B4_MCI_Mart_Answer.xlsx`: bảng Sales trên sheet `Sales_PQ`, sheet `KPI_SUMIFS`) và các file lookup / target mà query của file đó đọc, bằng `2_Dataset/_generator/build_b5_meta.py` (lưu ở `build_b5_meta.json`). Script tính lại từng measure đúng như DAX sẽ tính: lọc qua bảng lookup, tỷ lệ = tổng ÷ tổng, ô trống khi không có dữ liệu. Đối soát với `KPI_SUMIFS` của B2: **64/64 ô khớp**.
> Căn cứ: `Course_Plan.md` v3.3 (B5: timetable 150 phút, Hướng dẫn giảng dạy B5, Phụ lục A chuỗi A/B/C, Phụ lục C KPI dictionary, Phụ lục F câu chuyện #2, #3, #4, #5).
>
> **Cập nhật 13/10/2026, sửa theo hai lượt rà soát `consistency-checker`:** giới thiệu hai file demo ở Slide 3 và 15; thứ tự tab Power Pivot (Slide 5); thứ tự bước Diagram View (Slide 10); demo GM_Dong bắt buộc, ví dụ trà xanh / hộp quà Tết (Slide 16); bốn bộ lọc (Slide 19); Pivot chép sang cột G (Slide 20); ALL(Store) chỉ bỏ lọc Store (Slide 23); định dạng số cho measure (Slide 23, 24, 27); ST12 target 25 triệu (Slide 24); lọc ST12 bằng nút lọc Row Labels và PRO TIP like-for-like (Slide 28); Revenue YoY % có IF(ISBLANK…) trong KPI dictionary (Slide 30); Slide 31–32 làm mẫu trên `Practice/B5_MCI_Mart_Answer.xlsx`, sheet `PV_DoiSoat_B2`; chân slide ghi đúng nguồn ảnh. Sửa kèm: `Huong_dan` trong `build_b5_demos.py`, `build_slides_b5.py`, `build_docs_b5.py`, `_DANH_SACH_ANH.md`. **Cần dựng lại:** demo (chỉ Windows), deck + `_HocVien`, handout + PDF, infographic.
>
> **Trạng thái 03/10/2026, B5 theo mô hình file của B3 (có `Demo/`).** Đã dựng: deck `Slide_B5.pptx` + `_HocVien` (39 slide, 19/19 ảnh deck đã chèn, **0 khung [CHỤP]**); đáp án `Practice/B5_MCI_Mart_Answer.xlsx` (24/24 ✔; bước làm tay Mark as Date Table và Hide from Client Tools chưa xác nhận); `Demo/` (`B5_Demo_Model`, `B5_Demo_DAX` + bản `_HocVien`, sheet `Huong_dan`); `Homework/` (đề + `_Answer`, `Kiem_tra` 24/24 ✔); `Handout/B5_Handout.docx` + `.pdf` (dựng và xuất 03/10/2026, 8 trang, 5 ảnh H01–H05); `Handout/B5_Infographic_TomTat.png`. **Chưa có:** `grade_b5.py` (agent `homework-grader` tự viết khi chấm). Dòng *(→ …)* và chân slide đã ghi đường dẫn `Demo/…`; speaker notes trong `.pptx` khớp kịch bản sau mỗi lần chạy `build_slides_b5.py`.

## Cách dùng kịch bản

- **Lời giảng** là lời nói mẫu. Giảng viên không cần đọc nguyên văn; nên nói theo giọng của mình nhưng giữ đủ **ý chính** và **câu hỏi cho lớp** (in đậm).
- Dòng *(→ …)* trong lời giảng là **chỉ dẫn hành động**. Đường dẫn file tính từ thư mục buổi này (`1_Sessions/B5_Data_Model_DAX/`); file khởi đầu nằm ở `../B4_Power_Query/Practice/`.
- Mỗi slide có thời lượng dự kiến. Nếu lớp chậm, cắt ở những chỗ đánh dấu **[có thể rút gọn]**.
- **Buổi này có 2 file demo (theo mẫu B3): `Demo/B5_Demo_Model.xlsx` (Slide 5–12) và `Demo/B5_Demo_DAX.xlsx` (Slide 15–30)**, mỗi file có bản `_HocVien` và sheet `Huong_dan`. Giảng viên mở file demo gọn thay vì file khởi đầu nhiều sheet của B2–B4. Học viên làm cùng trên file của mình (`B5_<HoTen>.xlsx`, bản sao file khởi đầu) từ đầu tới hết Thực hành 2; ai tụt lại ở phần relationship thì mở `Demo/B5_Demo_DAX_HocVien.xlsx` (đã nối sẵn) và làm tiếp trên file đó tới hết buổi. Slide 3 giới thiệu hai file demo, Slide 15 nói lý do đổi file. Hai demo cố ý ra kết quả sai (Target lặp lại ở Slide 12, cột GM từng dòng ở Slide 16) làm trên các file này rồi sửa hoặc xóa trước khi đi tiếp. Đã dựng 02/10/2026, đặc tả và cách dựng ở Phụ lục "File Demo".
- **File khởi đầu:** bản sao `../B4_Power_Query/Practice/B4_MCI_Mart_Answer.xlsx`, lưu thành `B5_<HoTen>.xlsx`. Data Model đã có 6 bảng: Sales 30.247 dòng (01/2024–04/2025), Product 40, Store 12, Customer 1.500, Target 144, Calendar 731. **Chưa có relationship, chưa có measure.** Sheet `KPI_SUMIFS` của B2 vẫn nằm trong file, nên đối soát làm ngay tại chỗ.
- **Không bấm Refresh All trong buổi này.** Tham số `DataFolder` của file khởi đầu trỏ tới thư mục `Practice` trên máy dựng file, còn dữ liệu thì đã nằm sẵn trong Data Model. Lỡ bấm mà báo lỗi đường dẫn thì bỏ qua: refresh lỗi không xóa dữ liệu cũ. **Đừng** sửa `DataFolder` về `../B4_Power_Query/Practice/`: thư mục `Data_Sales` ở đó chỉ có 4 file (chưa có tháng 4), Refresh xong Sales còn 28.328 dòng và mọi con số tháng 4 trong kịch bản sẽ sai.
- **Chỉ dùng 6 bảng của Data Model.** Tab **All** của PivotTable Fields còn hiện các bảng worksheet của B1–B3 (`tblSales`, `tblProduct`, `tblStore`). Kéo nhầm trường của chúng, Excel sẽ thêm bảng đó vào Data Model. Bảng `Sales` (sheet `Sales_PQ`) là bảng của Power Query, đúng bảng cần dùng.
- **Bảng Sales có sẵn cột `Nam`, `Thang`** (B4 thêm để đối soát). Hôm nay mọi chiều thời gian lấy từ `Calendar`; Slide 20 giải thích vì sao.
- **Gõ tiếng Việt trong DAX:** tên miền "Bắc", "Trung" có dấu. Ví dụ CALCULATE ở Slide 22 dùng "Nam" để không phải gõ dấu.
- **Máy trình chiếu:** bật add-in Power Pivot (File → Options → Add-ins → COM Add-ins → Microsoft Power Pivot for Excel), đã kiểm ở checklist B1.

## Quy ước

Loại slide và nguồn ảnh giống `Slide_Script_B4.md`: BÌA · CHIA PHẦN · KHÁI NIỆM · CÔNG CỤ · THỰC HÀNH · QUIZ · TỔNG KẾT; **[CHỤP]** ảnh chụp Excel 365 Windows trên file của khóa, **[PQ PDF pX]** trang PDF nguồn (`4_Reference/PowerQuery_PowerPivot_DAX.pdf`, chỉ để trong kịch bản, không ghi lên slide), **[VẼ]** vẽ bằng shape, chỉ cho sơ đồ khái niệm và thẻ số (theo `CLAUDE.md`, mục Ảnh minh họa).

**File dùng trong buổi** (đường dẫn tính từ thư mục buổi; `../B4_Power_Query/` là thư mục Buổi 4):

| File | Công dụng | Slide |
|---|---|---|
| `../B4_Power_Query/Practice/B4_MCI_Mart_Answer.xlsx` | **File khởi đầu** (đáp án B4): Data Model 6 bảng chưa nối, sheet `KPI_SUMIFS` (B2), `Kiem_tra_B4` 10/10 ✔. Lưu thành `B5_<HoTen>.xlsx`, làm mọi việc hôm nay trong file này | 1, 3, 5–32 |
| ↳ sheet `KPI_SUMIFS` trong file khởi đầu (có cùng địa chỉ ô trong hai file demo và file đáp án Practice; lời giảng gọi là "sheet KPI_SUMIFS trong file đang mở") | "Đáp án" đối soát của B2: doanh thu B6:E9, số đơn B13:E16, AOV B20:E23, GM% B27:E30 (Miền × Tháng 1–3, cột E = Quý 1) | 2, 16, 19, 32 |
| `Demo/B5_Demo_Model.xlsx` (+ `_HocVien`) | **Đã dựng 02/10/2026** (`build_b5_demos.py`, chỉ Windows). Data Model 6 bảng **chưa có relationship** (giống file khởi đầu), sheet `Sales`/`Product`/`Store`/`KPI_SUMIFS` như `B3_Demo_Pivot`, sheet `Huong_dan`. Giảng viên làm mẫu Data View, Diagram View, tạo relationship, Target lặp (`PV_Grain`); học viên làm theo trên `B5_<HoTen>.xlsx`, bản `_HocVien` để đọc `Huong_dan`. Giới thiệu ở Slide 3 | 3, 5, 8, 10, 11, 12 |
| `Demo/B5_Demo_DAX.xlsx` (+ `_HocVien`) | **Đã dựng 02/10/2026.** File mốc đúng trạng thái cuối Slide 12: mô hình ngôi sao đã nối (6 relationship), **chưa có measure**, chưa Mark as Date Table, chưa cột COGS; có sẵn sheet trống `PV_Loc` (ô A1 ghi "Đặt PivotTable tại ô A3"). Giảng viên đổi sang file này từ Slide 15; học viên tụt lại mở bản `_HocVien` làm tiếp tới hết buổi. Demo calculated column, measure, filter context, CALCULATE/ALL, Target, YTD/YoY, KPI dictionary. Giới thiệu ở Slide 3 và 15 | 3, 15–17, 19–20, 22–25, 27–28, 30 |
| `../B4_Power_Query/Homework/B4_BTVN_MCI_Mart_T4_2025_Answer.xlsx` | Đáp án bài về nhà B4 (sheet `KetQua_T4`), chữa nhanh đầu buổi | 3 |
| `Practice/B5_MCI_Mart_Answer.xlsx` | **Đáp án B5 - đã dựng 02/10/2026** (`_generator/build_b5.py`, chỉ Windows, Excel qua COM; `Kiem_tra_B5` 24/24 ✔): file khởi đầu + 6 relationship + 14 measure + 7 Pivot từ Data Model + `PV_KPI_Thang` (bảng kiểm tra Slide 31) + `PV_DoiSoat_B2` (Pivot Miền × Tháng ở A8:Q14; Slide 32 làm mẫu GETPIVOTDATA ở ô J11, N14) + `DoiSoat_B2` (64 ô CUBEVALUE − KPI_SUMIFS; Slide 32 mở sheet này) + `KPI_Dictionary` + `Kiem_tra_B5` (24 dòng). Ảnh S19, S22–S25, S27b, S28, S30 chụp trên bản sao file này. Excel không có API cho calculated column, nên Total COGS trong đáp án viết bằng `SUMX(Sales, Sales[Quantity] * RELATED(Product[Unit_Cost]))` - cùng kết quả với cột COGS làm trên lớp. Mark as Date Table, Hide from Client Tools: làm tay sau khi dựng (chưa xác nhận đã làm). **File khởi đầu Buổi 6** | 3, 31, 32, 39 |
| `Homework/B5_BTVN_MCI_Mart_KPI_T1_T4_2025.xlsx` (+ `_Answer`) | **Bài tập về nhà bắt buộc - đã dựng 02/10/2026** (`build_b5_homework.py`, chỉ Windows; đáp án `Kiem_tra` 24/24 ✔): kiểm toán báo cáo cũ của phòng Kinh doanh, tập dượt C3, C4 của capstone. Mô tả ở Phụ lục kịch bản | 39 |
| `Handout/B5_Handout.docx` / `.pdf` | Handout tự học - `.docx` đã dựng 03/10/2026 (`_generator/build_docs_b5.py`, 9 mục, 5 ảnh); `.pdf` xuất bằng Word 03/10/2026 (8 trang, không còn [CHỤP]) | 39 |
| `Handout/B5_Infographic_TomTat.png` | Infographic tóm tắt buổi - đã dựng 02/10/2026 (nguồn `_generator/infographic_b5.html`); rà lại khi chốt kịch bản | 39 |

**Số liệu dùng trong buổi** (`build_b5_meta.json`):

| Chỉ tiêu | Giá trị |
|---|---|
| Data Model trong file khởi đầu | Sales **30.247** · Product **40** · Store **12** · Customer **1.500** · Target **144** · Calendar **731**; khóa ngoại không khớp lookup: **0** |
| Merge cả 3 lookup vào Sales | 30.247 dòng × 15 cột mô tả = **453.705** ô lặp lại; giữ lookup riêng: **9.220** ô |
| Target cả năm 2025 | **3.665.000.000**; theo miền: Bắc 1.227.000.000 · Nam 1.685.000.000 · Trung 753.000.000 |
| Quý 1/2025 | doanh thu **941.454.000** · COGS **619.663.000** · **3.144** đơn · AOV **299.445** · Gross Margin **34,18%** · **1.208** khách |
| GM% tính sai từ cột GM từng dòng (Quý 1) | AVERAGE **33,77%** · SUM **208.904%** (đúng: 34,18%) |
| Bắc × Tháng 3/2025 | **90.210.000** (= `KPI_SUMIFS!D6`) |
| % doanh thu Quý 1 theo miền | Bắc **34,4%** · Nam **45,4%** · Trung **20,2%** (= Pct_Share B1) |
| Doanh thu miền Nam Quý 1 (ví dụ CALCULATE) | **427.663.000** |
| Active Customers Quý 1 theo miền | Bắc 492 + Nam 657 + Trung 307 = **1.456** ≠ tổng **1.208** (DISTINCTCOUNT không cộng dồn) |
| Revenue vs Target % theo tháng 2025 | T1 **120,6%** · T2 **82,5%** · T3 **94,6%** · T4 **103,1%**; T1+T2 gộp **103,2%**; Quý 1 **100,6%** (target Quý 1 936.000.000) |
| Variance theo tháng 2025 | T1 +72.850.000 · T2 −52.004.000 · T3 −15.392.000 · T4 +8.965.000; **không chặn ô trống** thì T5 = −288.000.000 |
| ST07 (câu chuyện #2) | T3/2025: 8.539.000 / 19.000.000 = **44,9%**; T4: **97,4%** |
| Revenue vs Target % theo ngành, Quý 1 (sai grain) | Bánh kẹo 36,2% · Gia dụng 21,7% · Thực phẩm khô 18,0% · Đồ uống 11,8% · Chăm sóc cá nhân 11,2% · Sữa 1,8% (doanh thu ngành ÷ target cả chuỗi) |
| Revenue LY · YoY theo tháng 2025 | T1 298.203.000 · **+42,8%** · T2 248.304.000 · **−0,9%** · T3 237.192.000 · **+13,7%** · T4 239.466.000 · **+24,0%**; không chặn ô trống thì T5 = **−100%** |
| Revenue YTD 2025 | T1 425.850.000 · T2 671.846.000 · T3 941.454.000 · T4 **1.238.419.000** (cùng kỳ 2024: 1.023.165.000 → **+21,0%**) |
| ST12 (câu chuyện #3) | T1–T4/2024: 0 · T1–T4/2025: **79.399.000** → YoY trống; like-for-like 11 cửa hàng: YTD T4 **+13,3%** |
| YoY T1–T4 theo cửa hàng | cao nhất ST10 **+37,9%**; âm: ST07 −3,2%, ST04 −1,5% |
| Gross Margin % theo năm (câu chuyện #5) | 2024 **31,2%** · 2025 **34,1%** (T1–T4: 31,3% → 34,1%); giá bán 2025 cao hơn 2024 khoảng **3,9%**, giá vốn chỉ có một bảng |
| AOV ngành Bánh kẹo, Quý 1 (kiểm tay) | 338.733.000 ÷ **1.739** đơn = **194.786**; AVERAGE doanh thu từng dòng = 153.830 (không phải AOV) |

KPI theo tháng (Thực hành 1, lọc `Calendar[Nam]` = 2025):

| Tháng | Total Revenue | Total Orders | Avg Order Value | Gross Margin % | Revenue YoY % | Active Customers |
|---|---|---|---|---|---|---|
| 1 | 425.850.000 | 1.398 | 304.614 | 35,2% | +42,8% | 831 |
| 2 | 245.996.000 | 823 | 298.902 | 33,7% | −0,9% | 605 |
| 3 | 269.608.000 | 923 | 292.100 | 33,1% | +13,7% | 637 |
| 4 | 296.965.000 | 966 | 307.417 | 33,8% | +24,0% | 670 |

**Measure dựng trong buổi** (bảng chứa: Sales, trừ `Total Target` nằm ở Target):

| Measure | DAX | Định dạng | Slide |
|---|---|---|---|
| Total Revenue | `SUM(Sales[Revenue])` | #,##0 | 17 |
| Total COGS | `SUM(Sales[COGS])` (cột COGS = `Sales[Quantity] * RELATED(Product[Unit_Cost])`) | #,##0 | 15, 17 |
| Gross Margin % | `DIVIDE([Total Revenue] - [Total COGS], [Total Revenue])` | 0,0% | 17 |
| Revenue Mien Nam *(xóa sau demo)* | `CALCULATE([Total Revenue], Store[Region] = "Nam")` | #,##0 | 22 |
| Revenue All Stores | `CALCULATE([Total Revenue], ALL(Store))` | #,##0 | 23 |
| % Revenue Share | `DIVIDE([Total Revenue], [Revenue All Stores])` | 0,0% | 23 |
| Total Target | `SUM(Target[Target])` | #,##0 | 24 |
| Revenue vs Target % | `DIVIDE([Total Revenue], [Total Target])` | 0,0% | 24 |
| Variance | `IF(ISBLANK([Total Revenue]), BLANK(), [Total Revenue] - [Total Target])` | #,##0 | 24 |
| Revenue YTD | `CALCULATE([Total Revenue], DATESYTD(Calendar[Date]))` | #,##0 | 27 |
| Revenue LY | `CALCULATE([Total Revenue], DATEADD(Calendar[Date], -1, YEAR))` | #,##0 | 27 |
| Revenue YoY % | `IF(ISBLANK([Total Revenue]), BLANK(), DIVIDE([Total Revenue] - [Revenue LY], [Revenue LY]))` | 0,0% | 27 |
| Total Orders | `DISTINCTCOUNT(Sales[Order_ID])` | #,##0 | 31 |
| Avg Order Value | `DIVIDE([Total Revenue], [Total Orders])` | #,##0 | 31 |
| Active Customers | `DISTINCTCOUNT(Sales[Customer_ID])` | #,##0 | 31 |

**Relationship dựng trong buổi:** `Sales[Product_Code]` → `Product[Product_Code]` · `Sales[Store_Code]` → `Store[Store_Code]` · `Sales[Customer_ID]` → `Customer[Customer_ID]` · `Sales[Order_Date]` → `Calendar[Date]` (Slide 10) · `Target[Store_Code]` → `Store[Store_Code]` · `Target[Ngay_Dau_Thang]` → `Calendar[Date]` (Slide 12).

**Sheet Pivot tạo trong buổi** (đều là PivotTable From Data Model): `PV_Grain` (Slide 12), `PV_Loc` (16, 17, 19, 22, 23; sheet trống có sẵn trong file demo, Pivot đặt vào ô A3), `PV_Lookup_vs_Fact` (20), `PV_Target` (24, 25), `PV_YoY` (27, 28, 30), `PV_KPI_Thang` (31), `PV_DoiSoat_B2` (32).

**Tổng:** 39 slide, 150 phút.

---

## Phần 0 — Mở đầu: hai phòng ban, hai con số (5 phút · slide 1–3)

### Slide 1 · BÌA · 0.5'
- **Trên slide:** MCI / Consulting & Analytics / **Excel for Business Intelligence Analyst** / Buổi 5: Mô hình dữ liệu & DAX / "Bắt buộc máy Windows có Power Pivot (Data → Manage Data Model)" / "Mở sẵn: file khởi đầu B4_MCI_Mart_Answer.xlsx (đáp án Buổi 4)"
- **Ảnh:** [VẼ] nền theo Style spec.

**Lời giảng:**
> Chào cả lớp, chào mừng đến Buổi 5. Buổi trước các bạn đã có sáu bảng sạch, tự cập nhật bằng một nút Refresh. Hôm nay mình nối sáu bảng đó thành một mô hình, và viết công thức KPI **một lần** để mọi báo cáo dùng chung.

### Slide 2 · KHÁI NIỆM · Câu hỏi mở buổi · 1.5'
- **Trên slide:** *"Doanh thu tháng này đạt bao nhiêu % mục tiêu, và vì sao mỗi báo cáo cũ lại ra một con số khác nhau?"* · nhãn Tầng 3, Diagnostic.
- **Thẻ số:** Gross Margin % Quý 1/2025, Báo cáo A: **34,18%** · Báo cáo B: **33,77%**. Cùng một dữ liệu.
- **Ảnh:** [VẼ] hai thẻ số đặt cạnh nhau, dấu "≠" ở giữa.

**Lời giảng:**
> Câu hỏi hôm nay có hai nửa. Nửa đầu dễ: tháng này đạt bao nhiêu phần trăm mục tiêu. Nửa sau mới là chuyện thật ở công ty: **vì sao mỗi báo cáo lại ra một con số khác nhau?**
>
> Ví dụ trên slide. Cùng dữ liệu Quý 1, báo cáo A ghi biên lợi nhuận gộp 34,18%, báo cáo B ghi 33,77%. Không ai bịa số. Báo cáo A lấy tổng lợi nhuận chia tổng doanh thu, đúng như bảng KPI_SUMIFS các bạn làm ở Buổi 2. Báo cáo B tính biên lợi nhuận cho từng dòng rồi lấy trung bình. Hai công thức, hai con số, và cuộc họp mất nửa tiếng để cãi xem số nào đúng.
>
> Đây là câu hỏi Tầng 3, **vì sao**. Cuối buổi các bạn sẽ có câu trả lời cho cả hai nửa: một mô hình dữ liệu, và mỗi KPI chỉ có **một định nghĩa** duy nhất.

### Slide 3 · KHÁI NIỆM · Recap B4 & file khởi đầu · 3'
- **Trên slide:**
  - Viên gạch B4: 6 bảng sạch trong Data Model, Sales **30.247** dòng (4.436.849.500), Product 40, Store 12, Customer 1.500, Target 144, Calendar 731. Refresh tự cập nhật.
  - Bài về nhà B4: đối soát trước khi tính % target, ST07 tháng 4 thiếu file bổ sung chỉ đạt **54,3%**, đủ file đạt **97,4%**; cả chuỗi **103,1%**.
  - Hôm nay nhận **file khởi đầu** = đáp án B4 (`B4_MCI_Mart_Answer.xlsx`), lưu thành `B5_<HoTen>.xlsx`. **Không bấm Refresh All.**
  - Giảng viên làm mẫu trên 2 file gọn: `Demo/B5_Demo_Model.xlsx` (Phần 1–2, chưa nối) · `Demo/B5_Demo_DAX.xlsx` (Phần 3–7, đã nối); mỗi file có bản `_HocVien`.
  - Cuối buổi: đối soát với KPI_SUMIFS (Thực hành 2), đáp án ở sheet DoiSoat_B2 của `Practice/B5_MCI_Mart_Answer.xlsx`; file này là file khởi đầu Buổi 6.
- **Ảnh:** [VẼ] viên gạch B4 (đã xong) nằm dưới viên gạch B5 (sáng lên).

**Lời giảng:**
> *(→ mở `../B4_Power_Query/Practice/B4_MCI_Mart_Answer.xlsx`, sheet Kiem_tra_B4)* Nhắc nhanh buổi trước: sáu bảng sạch, Sales ba mươi nghìn hai trăm bốn mươi bảy dòng, mười dòng đối soát đều tick xanh.
>
> Bài về nhà: tháng 4, mỗi cửa hàng một file. Bạn nào đối soát với số liệu hệ thống thì phát hiện ST07 thiếu nửa tháng, chưa thêm file bổ sung, ST07 chỉ đạt 54,3% target; thêm vào thì lên 97,4%. *(→ nếu còn thời gian: mở `../B4_Power_Query/Homework/B4_BTVN_MCI_Mart_T4_2025_Answer.xlsx`, sheet KetQua_T4)* Đối soát trước, kết luận sau, hôm nay mình còn dùng thói quen đó nhiều lần.
>
> File khởi đầu hôm nay là chính file đáp án Buổi 4. Các bạn lưu thành `B5_<HoTen>.xlsx`. Một lưu ý quan trọng: **hôm nay không bấm Refresh All**. Dữ liệu đã nằm sẵn trong Data Model; tham số thư mục trong file đang trỏ tới máy khác, bấm Refresh chỉ ra lỗi.
>
> Để màn hình chiếu gọn, mình làm mẫu trên hai file demo trong thư mục `Demo` của buổi, thay cho file khởi đầu có hơn hai mươi sheet của Buổi 2 đến Buổi 4. File thứ nhất, `Demo/B5_Demo_Model.xlsx`: đúng sáu bảng của Data Model, chưa nối gì, giống file của các bạn lúc này; mình dùng nó cho Phần 1 và Phần 2. File thứ hai, `Demo/B5_Demo_DAX.xlsx`: cùng sáu bảng đó nhưng đã nối xong, đúng trạng thái sau Phần 2; mình chuyển sang nó từ Phần 3. Các bạn làm theo trên file `B5_<HoTen>.xlsx` của mình. Mỗi file demo có bản `_HocVien` với sheet `Huong_dan` ghi từng bước theo slide; bạn nào lỡ một bước ở phần nối bảng thì mở `Demo/B5_Demo_DAX_HocVien.xlsx` và làm tiếp trên đó tới hết buổi.
>
> Để các bạn biết đích đến của buổi: cuối giờ mình đối soát bảng KPI của Buổi 2 ngay trong file này (Thực hành 2). Mình sẽ làm mẫu trên Pivot `PV_DoiSoat_B2` của file đáp án `Practice/B5_MCI_Mart_Answer.xlsx`; bảng đối soát đủ 64 ô nằm ngay trong file đó, sheet `DoiSoat_B2`. Đó là file đáp án của hôm nay, và cũng là file khởi đầu của Buổi 6.

---

## Phần 1 — Data Model: các bảng sống chung, không gộp làm một (20 phút · slide 4–8)

### Slide 4 · CHIA PHẦN · 0.5'
- **Trên slide:** MCI / **1. Data Model: nhiều bảng, một mô hình**

**Lời giảng:**
> Phần 1: mở Data Model ra xem, và hiểu vì sao mình giữ các bảng riêng thay vì gộp thành một bảng khổng lồ.

### Slide 5 · CÔNG CỤ · Data View · Diagram View · 5.5'

| Tên (đỏ) | Mục đích |
|---|---|
| **Data → Manage Data Model** (hoặc **Power Pivot → Manage**) | Mở cửa sổ Power Pivot - nơi chứa Data Model, tách riêng khỏi worksheet |
| **Data View** | Mỗi bảng một tab; xem dữ liệu, thêm calculated column; dòng **Record** dưới cùng ghi số dòng của bảng |
| **Diagram View** | Mỗi bảng là một hộp; relationship là đường nối giữa các hộp |

- **Ảnh:** [CHỤP] S05a: Data View, tab Sales đang chọn, 6 tab bảng ở dưới (Product, Store, Customer, Target, Calendar, Sales), dòng Record ghi 1 of 30,247. [CHỤP] S05b: Diagram View, 6 hộp bảng **chưa có đường nối nào**. [PQ PDF p39–41]
- **TIP:** Không thấy nút Manage Data Model / tab Power Pivot: bật add-in Power Pivot (File → Options → Add-ins → COM Add-ins). Data View **chỉ để xem**: không sửa từng ô như trên worksheet.
- **PRO TIP:** Trong Diagram View, kéo các bảng lookup bao quanh bảng fact, sơ đồ dễ đọc hơn hẳn.
- **Chân slide:** Ảnh: Demo/B5_Demo_Model.xlsx, Data → Manage Data Model (Data View, Diagram View).

**Lời giảng:**
> *(→ mở `Demo/B5_Demo_Model.xlsx`: Data → Manage Data Model)* File demo thứ nhất, sáu bảng chưa nối, giống file các bạn đang mở. Đây là **cửa sổ Power Pivot**, một cửa sổ riêng. Mọi thứ Buổi 4 nạp vào Data Model nằm ở đây. *(→ chỉ các tab dưới cùng)* Sáu tab, sáu bảng: Product, Store, Customer, Target, Calendar, và Sales ở cuối. *(→ bấm tab Sales, chỉ dòng Record ở dưới cùng: 1 of 30,247)* Sales: ba mươi nghìn hai trăm bốn mươi bảy dòng, đúng con số đối soát cuối Buổi 4.
>
> Đây là **Data View**: xem dữ liệu theo dạng bảng. Nó giống worksheet nhưng không sửa được từng ô; dữ liệu đến từ Power Query, muốn sửa thì sửa ở query.
>
> *(→ Home → Diagram View)* Còn đây là **Diagram View**. Mỗi bảng là một hộp, liệt kê tên cột. **Câu hỏi cho lớp: nhìn sơ đồ này, các bạn thấy thiếu gì?** *(→ chờ)* Thiếu đường nối. Sáu bảng đứng rời nhau, đúng giới hạn mình nói ở cuối Buổi 4. Nối chúng lại là việc của Phần 2.

### Slide 6 · KHÁI NIỆM · Normalization · Fact vs Lookup · 5'
- **Trên slide:**
  - **Normalization:** mỗi thông tin lưu đúng **một chỗ**, tên ngành hàng nằm ở Product, không lặp lại trên từng dòng bán hàng [PQ PDF p42].
  - **Fact (data) table:** ghi lại **sự kiện**, nhiều dòng, có cột số để cộng: **Sales** (30.247 dòng; grain: 1 sản phẩm trong 1 đơn) · **Target** (144 dòng; grain: tháng × cửa hàng) [PQ PDF p43].
  - **Lookup (dimension) table:** **mô tả** đối tượng, mỗi đối tượng một dòng: **Product** 40 · **Store** 12 · **Customer** 1.500 · **Calendar** 731.
- **Ảnh:** [VẼ] hai cột: Fact (Sales, Target) bên trái, Lookup (Product, Store, Customer, Calendar) bên phải; dưới mỗi bảng ghi grain và số dòng.

**Lời giảng:**
> Buổi 2 các bạn đã gặp chữ **grain**: một dòng của bảng là một cái gì. Hôm nay mình chia bảng thành hai loại theo vai trò.
>
> **Bảng fact** ghi sự kiện: mỗi lần bán một sản phẩm là một dòng Sales. Bảng fact dài, và có cột số để cộng, Quantity, Revenue. **Bảng lookup** mô tả đối tượng: mỗi sản phẩm một dòng, mỗi cửa hàng một dòng, mỗi ngày một dòng. Bảng lookup ngắn, chủ yếu là chữ.
>
> Nguyên tắc **normalization**: mỗi thông tin chỉ lưu một chỗ. Tên ngành "Bánh kẹo" nằm một lần ở bảng Product, không lặp lại trên hàng nghìn dòng bán hàng.
>
> **Câu hỏi cho lớp: bảng Target là fact hay lookup?** *(→ chờ)* Target ghi một con số, doanh thu kế hoạch, cho mỗi cửa hàng mỗi tháng. Nó là **fact**: fact thứ hai, với grain khác hẳn Sales. Sales chi tiết tới từng dòng đơn; Target chỉ tới tháng và cửa hàng, không có sản phẩm. Nhớ chi tiết này, Phần 2 nó sẽ gây ra một con số rất lạ.

### Slide 7 · KHÁI NIỆM · Primary key · Foreign key · 4'

| Bảng lookup | Primary key (duy nhất) | Foreign key ở bảng fact (lặp lại) |
|---|---|---|
| Product | `Product_Code` - 40 mã | `Sales[Product_Code]` |
| Store | `Store_Code` - 12 mã | `Sales[Store_Code]`, `Target[Store_Code]` |
| Customer | `Customer_ID` - 1.500 mã | `Sales[Customer_ID]` |
| Calendar | `Date` - 731 ngày | `Sales[Order_Date]`, `Target[Ngay_Dau_Thang]` |

- **Thẻ số:** Khóa chính phải **duy nhất**, P021 trùng ở B2, P015 trùng ở B4 · Khóa ngoại không có trong lookup: **0 dòng** (đã kiểm bằng Left Anti ở B4).
- **Ảnh:** [VẼ] bảng Product (mỗi mã một dòng) nối với cột Product_Code của Sales (mã lặp lại nhiều dòng). [PQ PDF p44]

**Lời giảng:**
> Muốn nối bảng thì cần **khóa**. Ở bảng lookup, cột khóa gọi là **primary key**, khóa chính: mỗi giá trị chỉ xuất hiện một lần. Bốn mươi sản phẩm, bốn mươi mã. Ở bảng fact, cùng cột đó gọi là **foreign key**, khóa ngoại: một mã sản phẩm xuất hiện ở hàng trăm dòng bán hàng.
>
> Các bạn đã gặp khóa ba lần rồi. Buổi 2, XLOOKUP sai vì P021 trùng. Buổi 4, Merge làm Sales phình ra vì P015 trùng. Hôm nay quy tắc vẫn vậy, chỉ chặt hơn: khóa chính trùng thì Data Model **không cho tạo** quan hệ.
>
> Để ý dòng Calendar: khóa chính là `Date`, mỗi ngày một dòng. Sales nối vào bằng Order_Date, còn Target nối vào bằng `Ngay_Dau_Thang`, cái cột ngày đầu tháng các bạn thêm ở Buổi 4. Hôm nay nó phát huy tác dụng.

### Slide 8 · KHÁI NIỆM · Relationships vs Merged tables · 5'
- **Trên slide:** [PQ PDF p45]
  - **Merge** (Buổi 4): kéo cột của lookup vào bảng fact → một bảng to, thông tin lặp lại trên mọi dòng; lookup trùng khóa thì số dòng phình ra (28.328 → 29.408).
  - **Relationship:** giữ các bảng riêng, nối bằng khóa; Pivot lấy cột từ bảng nào cũng được.
  - **Thẻ số:** Merge cả Product, Store, Customer vào Sales: 30.247 dòng × 15 cột mô tả = **453.705 ô** lặp lại · Giữ lookup riêng: **9.220 ô**.
- **Ảnh:** [VẼ] trái: một bảng phẳng rất rộng; phải: hình ngôi sao, Sales ở giữa, 4 lookup xung quanh.

**Lời giảng:**
> **Câu hỏi cho lớp: Buổi 4 mình đã biết Merge. Sao không Merge hết Product, Store, Customer vào Sales cho xong?** *(→ chờ)*
>
> Làm được, nhưng tốn. Merge cả ba lookup thì mỗi dòng trong ba mươi nghìn dòng Sales phải mang thêm mười lăm cột mô tả: tên sản phẩm, ngành, tên cửa hàng, miền, phân khúc khách… Cộng lại hơn bốn trăm năm mươi nghìn ô, toàn là thông tin lặp. Giữ các bảng riêng thì chỉ cần hơn chín nghìn ô. Và nhớ cái bẫy Buổi 4: lookup trùng một khóa, Merge làm số dòng phình từ 28.328 lên 29.408 mà không báo lỗi nào.
>
> **Relationship** làm điều ngược lại: không chép gì cả. Mỗi bảng ở yên chỗ của nó, chỉ nối với nhau bằng khóa. Pivot muốn lấy "ngành hàng" thì đi theo đường nối sang bảng Product. Hình dạng của mô hình này có tên riêng: **star schema**, hình ngôi sao, fact ở giữa, lookup xung quanh. Mình dựng nó ngay bây giờ.

---

## Phần 2 — Relationship, cardinality, fact thứ hai (15 phút · slide 9–12)

### Slide 9 · CHIA PHẦN · 0.5'
- **Trên slide:** MCI / **2. Relationship: nối bảng bằng khóa, không bằng công thức**

**Lời giảng:**
> Phần 2: tạo relationship, và đưa fact thứ hai, Target, vào mô hình.

### Slide 10 · CÔNG CỤ · Tạo và quản lý Relationship · 5.5'

| Tên (đỏ) | Mục đích |
|---|---|
| **Kéo thả trong Diagram View** | Kéo cột khóa ngoại của bảng fact, thả lên khóa chính của bảng lookup |
| **Design → Create Relationship** | Hộp thoại chọn bảng và cột ở hai bên - dùng khi sơ đồ quá rối để kéo |
| **Design → Manage Relationships** (ngoài Excel: **Data → Relationships**) | Xem, sửa, xóa, bật / tắt từng relationship |

- **Bốn relationship của Sales:** `Sales[Product_Code]` → `Product[Product_Code]` · `Sales[Store_Code]` → `Store[Store_Code]` · `Sales[Customer_ID]` → `Customer[Customer_ID]` · `Sales[Order_Date]` → `Calendar[Date]`.
- **Ảnh:** [CHỤP] S10a: hộp Create Relationship, Sales / Order_Date ↔ Calendar / Date. [CHỤP] S10b: Diagram View, Sales ở giữa nối với Product, Store, Customer, Calendar; mỗi đường có ký hiệu **1** ở đầu lookup và **\*** ở đầu Sales. [PQ PDF p46, 48]
- **TIP:** Hai cột phải **cùng kiểu dữ liệu**: Order_Date và Calendar[Date] đều là Date nhờ bước Using Locale ở B4. Báo lỗi trùng giá trị khi tạo quan hệ = khóa của lookup bị trùng, quay lại Remove Duplicates trong Power Query.
- **Chân slide:** Ảnh: Demo/B5_Demo_Model.xlsx, Power Pivot → Diagram View; Design → Create Relationship.

**Lời giảng:**
> *(→ Diagram View: kéo cột Product_Code trong hộp Sales, thả lên Product_Code trong hộp Product)* Kéo từ bảng fact, thả lên bảng lookup. Một đường nối xuất hiện. *(→ làm tương tự: Sales[Store_Code] → Store[Store_Code], Sales[Customer_ID] → Customer[Customer_ID])*
>
> Cái thứ tư mình làm bằng hộp thoại, để các bạn biết cách thứ hai. *(→ Design → Create Relationship: bảng Sales, cột Order_Date; bảng Calendar, cột Date → OK)* Cách này hữu ích khi sơ đồ có nhiều bảng, kéo thả dễ nhầm.
>
> *(→ kéo hộp Sales vào giữa, bốn lookup xung quanh)* Nhìn kỹ một đường nối: đầu phía Product có số **1**, đầu phía Sales có dấu **sao**. **Câu hỏi cho lớp: số 1 và dấu sao nói gì?** *(→ chờ)* Một sản phẩm, xuất hiện ở nhiều dòng bán hàng. Một – nhiều. Slide sau mình nói kỹ.
>
> *(→ Design → Manage Relationships)* Đây là danh sách mọi quan hệ, bốn dòng. Muốn sửa hay xóa thì vào đây.

### Slide 11 · KHÁI NIỆM · Cardinality 1-* · Filter direction · 3.5'
- **Trên slide:** [PQ PDF p50, 54–55]
  - **Cardinality** = mức độ duy nhất của khóa ở hai đầu: **1** (lookup, duy nhất) – **\*** (fact, lặp lại). Data Model chỉ nhận 1-1 và 1-\*; khóa trùng ở cả hai đầu (nhiều–nhiều) thì không tạo được (handout).
  - **Filter direction:** bộ lọc chảy **từ lookup xuống fact**, chọn Region = Nam ở Store thì Sales chỉ còn dòng của cửa hàng miền Nam. Chiều ngược lại không có.
  - **Không bao giờ nối thẳng hai bảng fact.** Hai fact gặp nhau qua lookup chung.
- **Ảnh:** [VẼ] Store ở trên, Sales ở dưới, mũi tên chảy xuống; Store_Code = "ST07" lọc ra các dòng ST07.

**Lời giảng:**
> **Cardinality** nghe khó nhưng ý đơn giản: ở mỗi đầu đường nối, khóa duy nhất hay lặp lại. Đầu lookup duy nhất, số 1. Đầu fact lặp lại, dấu sao. Quan hệ chuẩn trong Data Model là một – nhiều.
>
> Quan trọng hơn là **hướng lọc**. Bộ lọc chảy như nước, từ trên xuống: từ lookup xuống fact. Các bạn chọn miền Nam ở bảng Store, bộ lọc chảy theo đường nối xuống Sales, Sales chỉ còn các dòng của cửa hàng miền Nam. Chiều ngược lại thì không: lọc bảng Sales không làm bảng Store ngắn đi.
>
> Và một quy tắc vàng: **không nối thẳng hai bảng fact với nhau**. Sales và Target đều có Store_Code, nhưng chúng gặp nhau qua bảng Store, lookup chung. Slide sau sẽ cho thấy vì sao quy tắc này quan trọng.

### Slide 12 · CÔNG CỤ · Fact thứ hai: Target lọc qua lookup chung · 5.5'
- **Tên (đỏ):** **Insert → PivotTable ▾ → From Data Model** · relationship cho **Target**
- **Mục đích:** Pivot đọc thẳng Data Model, lấy trường từ nhiều bảng cùng lúc. Target chỉ lọc được qua **lookup chung** với Sales: Store và Calendar.
- **Hai relationship của Target:** `Target[Store_Code]` → `Store[Store_Code]` · `Target[Ngay_Dau_Thang]` → `Calendar[Date]`.
- **Thẻ số:** Chưa nối: mọi ngành đều **3.665.000.000** · Đã nối, theo miền: Bắc **1.227.000.000** · Nam **1.685.000.000** · Trung **753.000.000**.
- **Ảnh:** [CHỤP] S12a: Pivot Rows = Product[Category], Values = Sum of Target: cả 6 dòng đều 3.665.000.000, thanh vàng "Relationships between tables may be needed" trong PivotTable Fields. [CHỤP] S12b: Diagram View đủ 6 bảng, 6 relationship, Sales và Target cùng nối vào Store và Calendar. [PQ PDF p53–55]
- **TIP:** Pivot hiện **cùng một con số ở mọi dòng** = trường ở Rows không lọc được bảng của trường ở Values. Kiểm tra Diagram View.
- **Chân slide:** Ảnh: Demo/B5_Demo_Model.xlsx, sheet PV_Grain (Insert → PivotTable → From Data Model); Power Pivot → Diagram View.

**Lời giảng:**
> Giờ đưa Target vào báo cáo. Mình cố ý làm "sai" trước. *(→ Insert → PivotTable ▾ → From Data Model → New Worksheet; đổi tên sheet thành `PV_Grain`)* PivotTable From Data Model đọc thẳng mô hình, nên danh sách trường có đủ sáu bảng. *(→ Rows: Product → Category; Values: tick Target → Target)*
>
> **Câu hỏi cho lớp: sáu ngành hàng, sáu dòng, mà dòng nào cũng 3.665.000.000. Vì sao?** *(→ chờ)* Thanh vàng phía trên danh sách trường gợi ý: thiếu relationship. Target chưa nối với bảng nào. Nhưng sâu hơn: bảng Target **không có cột sản phẩm**, kế hoạch chỉ lập tới cửa hàng và tháng. Nên dù nối thế nào, ngành hàng cũng không lọc được Target. Pivot không biết chia, nên đưa cả tổng vào mọi dòng.
>
> Nối Target vào hai lookup nó có chung với Sales. *(→ Diagram View: kéo Target[Store_Code] thả lên Store[Store_Code]; Design → Create Relationship: Target / Ngay_Dau_Thang, Calendar / Date → OK)* Cột Ngay_Dau_Thang của Buổi 4 chính là cái móc để Target nối vào lịch.
>
> *(→ quay lại PV_Grain: bỏ Category khỏi Rows, kéo Store → Region vào Rows)* Giờ chia đúng: Bắc 1 tỷ 227 triệu, Nam 1 tỷ 685 triệu, Trung 753 triệu. *(→ kéo Category trở lại Rows, dưới Region)* Còn ngành hàng thì vẫn lặp lại, và **đó là đúng**. Target chỉ lọc được theo những chiều mà Target có: cửa hàng, miền, tháng. *(→ Diagram View: chỉ hình ngôi sao có hai bảng ở giữa)* Đây là brick output thứ nhất hôm nay: một mô hình ngôi sao với **hai bảng fact**, gặp nhau qua Store và Calendar.

---

### Slide 13 · NGHỈ · 10'
- **Trên slide:** Nghỉ 10 phút · "Trước khi quay lại: mở Diagram View, kiểm tra đủ **6 relationship**."

---

## Phần 3 — Calculated Column vs Measure (15 phút · slide 14–17)

### Slide 14 · CHIA PHẦN · 0.5'
- **Trên slide:** MCI / **3. Calculated Column và Measure: tính từng dòng hay tính từng ô**

**Lời giảng:**
> Phần 3: hai cách viết công thức trong Data Model, và vì sao KPI phải chọn đúng cách.

### Slide 15 · CÔNG CỤ · Calculated Column · RELATED · 5'

| Tên (đỏ) | Mục đích |
|---|---|
| **Calculated Column** (Data View → cột **Add Column**) | Công thức tính cho **từng dòng**, kết quả **lưu** thành một cột của bảng |
| **RELATED(**`<cột ở bảng lookup>`**)** | Lấy giá trị từ bảng lookup theo relationship - XLOOKUP của Data Model |

- **Công thức:** `= Sales[Quantity] * RELATED(Product[Unit_Cost])` → đổi tên cột thành **COGS**.
- **Ảnh:** [CHỤP] S15: Data View bảng Sales, cột COGS mới ở cuối, thanh công thức hiện `=Sales[Quantity]*RELATED(Product[Unit_Cost])`. [PQ PDF p66–68, 101–102]
- **TIP:** Không có địa chỉ ô kiểu A1: công thức viết theo tên cột và áp cho cả cột. Đổi tên cột: double-click tiêu đề.
- **PRO TIP:** RELATED chỉ chạy khi **đã có relationship**, và đi từ fact sang lookup (từ phía "nhiều" sang phía "một").
- **Chân slide:** Ảnh: Demo/B5_Demo_DAX.xlsx, Power Pivot → Data View, bảng Sales, cột COGS.

**Lời giảng:**
> *(→ mở `Demo/B5_Demo_DAX.xlsx`: mô hình đã nối đủ 6 relationship, giống cuối Slide 12; học viên nào chưa kịp nối thì mở `Demo/B5_Demo_DAX_HocVien.xlsx` làm tiếp)* Từ phần này mình đổi sang file demo thứ hai đã nói ở đầu buổi: cùng sáu bảng, đã nối đủ sáu relationship như các bạn vừa làm, và có sẵn một sheet trống tên `PV_Loc` cho các Pivot sắp tới. Bạn nào chưa nối xong thì mở bản `_HocVien` của file này và làm tiếp trên đó. Buổi 2 các bạn tính giá vốn bằng cách XLOOKUP Unit_Cost từ bảng Product về bảng Sales, rồi nhân với Quantity. Trong Data Model, việc đó gọn hơn nhiều.
>
> *(→ Data View, tab Sales, bấm vào ô đầu tiên của cột Add Column; gõ trên thanh công thức `=Sales[Quantity]*RELATED(Product[Unit_Cost])` → Enter)* **RELATED** đi theo đường nối sang bảng Product, lấy Unit_Cost của đúng sản phẩm trên dòng đó. Không cần XLOOKUP, không cần khóa tuyệt đối, vì relationship đã làm việc nối. *(→ double-click tiêu đề cột, đổi tên thành `COGS`)*
>
> Đây là **calculated column**: tính một lần cho từng dòng, kết quả nằm trong bảng như một cột thật. Ba mươi nghìn dòng, ba mươi nghìn giá trị COGS. Cột kiểu này hợp với thuộc tính của **từng dòng**: giá vốn của dòng đó, nhóm của dòng đó.

### Slide 16 · KHÁI NIỆM · Vì sao KPI tỷ lệ phải là measure · 4'
- **Trên slide:** [PQ PDF p68, 76]
  - Thử làm GM% bằng calculated column `GM_Dong = DIVIDE(Sales[Revenue] - Sales[COGS], Sales[Revenue])`, rồi đưa vào Pivot Quý 1/2025:
  - **Thẻ số:** Average of GM_Dong **33,77%** · Sum of GM_Dong **208.904%** · Đúng: (941.454.000 − 619.663.000) ÷ 941.454.000 = **34,18%** (`KPI_SUMIFS!E30`).

| | Calculated Column | Measure |
|---|---|---|
| Tính theo | từng dòng (row context) | từng ô Pivot (filter context) |
| Lưu trong mô hình | có - tốn bộ nhớ | không - tính khi cần |
| Dùng ở | Rows, Columns, Filters, Slicer | chỉ ở Values |
| Hợp với | thuộc tính của một dòng | KPI, tổng, **tỷ lệ** |

- **Ảnh:** [CHỤP] S16: Pivot PV_Loc lọc Quý 1/2025, Values = Average of GM_Dong (33,77%) và Sum of GM_Dong.
- **Chân slide:** Ảnh: Demo/B5_Demo_DAX.xlsx, sheet PV_Loc (cột GM_Dong thêm tạm, xóa sau demo). Đối chiếu: sheet KPI_SUMIFS, ô E30.
- **Ghi chú trên slide (TIP):** Pivot đặt ở sheet PV_Loc, ô A3 (Existing Worksheet). File demo có sẵn sheet này; file B5_<HoTen>.xlsx: bấm dấu + cạnh tab sheet, đặt tên PV_Loc.

**Lời giảng:**
> *(→ Data View, Sales: Add Column `=DIVIDE(Sales[Revenue]-Sales[COGS],Sales[Revenue])`, đổi tên `GM_Dong`; mở sheet `PV_Loc` (có sẵn trong file demo, đang trống), bấm ô A3; Insert → PivotTable ▾ → From Data Model → Existing Worksheet, Location `PV_Loc!$A$3`; Filters: Calendar → Nam = 2025, Calendar → Quy = 1; Values: GM_Dong, đổi sang Average, định dạng %)* Trên file `B5_<HoTen>.xlsx` của các bạn chưa có sheet `PV_Loc`: bấm dấu + cạnh các tab sheet để thêm sheet mới, đổi tên thành `PV_Loc`, rồi chèn Pivot vào ô A3 như mình. Mình thử tính biên lợi nhuận bằng calculated column: biên của từng dòng, rồi để Pivot lấy trung bình. Ra **33,77%**. *(→ kéo GM_Dong lần nữa vào Values, để Sum)* Còn nếu Pivot cộng lại thì ra hai trăm nghìn phần trăm, vô nghĩa.
>
> *(→ mở sheet KPI_SUMIFS trong file đang mở, chỉ ô E30)* Số đúng của Buổi 2 là **34,18%**: tổng lợi nhuận chia tổng doanh thu. Vậy 33,77% ở slide mở đầu từ đâu ra, chính là trung bình của các tỷ lệ, cái bẫy "trung bình của trung bình" Buổi 2 đã gặp. Một dòng bán chai trà xanh mười nghìn đồng có trọng số ngang với một dòng bán hộp quà Tết sáu trăm năm mươi nghìn đồng.
>
> Quy tắc: **KPI dạng tỷ lệ luôn là measure**, tính tổng trước, chia sau, cho từng ô Pivot. *(→ Data View: chuột phải tiêu đề GM_Dong → Delete Columns; giữ lại sheet PV_Loc)* Xóa cột tạm này đi. Bảng so sánh trên slide là thứ các bạn cần nhớ: cột tính từng dòng và nằm ở hàng, cột; measure tính từng ô và nằm ở Values.

### Slide 17 · CÔNG CỤ · Implicit vs Explicit measure · 5.5'

| Tên (đỏ) | Mục đích |
|---|---|
| **Implicit measure** | Tick một cột số vào Values → Excel tự tạo "Sum of …" - không đặt tên, không định dạng, **không dùng lại được** |
| **Power Pivot → Measures → New Measure…** (hoặc chuột phải bảng trong PivotTable Fields → **Add Measure…**) | **Explicit measure**: tự đặt tên, viết DAX, chọn định dạng; dùng lại trong measure khác và trong CUBEVALUE (Buổi 6) |
| **DIVIDE(**`tử, mẫu`**)** | Phép chia an toàn: mẫu bằng 0 hoặc trống thì trả về ô trống thay vì lỗi |

- **Ba measure đầu tiên** (Table name: Sales):
  - `Total Revenue := SUM(Sales[Revenue])`, Number, dấu ngăn nghìn, 0 chữ số thập phân
  - `Total COGS := SUM(Sales[COGS])`
  - `Gross Margin % := DIVIDE([Total Revenue] - [Total COGS], [Total Revenue])`, Category Number, Format Percentage, 1 chữ số thập phân
- **Thẻ số (Quý 1/2025):** 941.454.000 · 619.663.000 · **34,2%**.
- **Ảnh:** [CHỤP] S17: hộp Measure của Gross Margin %, Table name Sales, Measure name, Formula, nút Check formula, Category = Number, Format = Percentage (ảnh chụp lúc Decimal places còn 2; giảng viên đổi thành 1 trên lớp). [PQ PDF p70–72, 77]
- **TIP:** Cột luôn ghi kèm tên bảng `Sales[Revenue]`; measure ghi trần `[Total Revenue]`. Nhìn công thức là biết đâu là cột, đâu là measure.
- **Chân slide:** Ảnh: Demo/B5_Demo_DAX.xlsx, Power Pivot → Measures → New Measure (Gross Margin %).

**Lời giảng:**
> *(→ trên PV_Loc, hai trường GM_Dong đã tự biến mất khi xóa cột: tick Sales → Revenue)* Tick cột Revenue, Pivot tự hiện "Sum of Revenue". Đây là **implicit measure**, measure ngầm. Tiện, nhưng không tên, không định dạng, và không dùng lại được trong công thức khác. *(→ bỏ tick Revenue)*
>
> Mình viết **explicit measure**. *(→ Power Pivot → Measures → New Measure…; Table name: Sales; Measure name: `Total Revenue`; Formula: `=SUM(Sales[Revenue])`; bấm Check formula; Category: Number, tick Use 1000 separator, Decimal places 0 → OK)* Có tên rõ ràng, có định dạng. Nó xuất hiện trong danh sách trường với biểu tượng fx. *(→ tick Total Revenue vào Values)* 941.454.000, khớp Quý 1.
>
> *(→ New Measure: `Total COGS` = `SUM(Sales[COGS])`, định dạng như trên; tick vào Values)* Tổng giá vốn Quý 1/2025: 619.663.000. *(→ New Measure: `Gross Margin %` = `DIVIDE([Total Revenue]-[Total COGS],[Total Revenue])`, Category: Number, Format: Percentage, Decimal places 1 → OK; tick vào Values)* Và đây là điều hay nhất: measure **gọi lại measure**. Gross Margin % không cần biết doanh thu tính thế nào, nó dùng Total Revenue. Ra **34,2%**, khớp KPI_SUMIFS. Từ giờ ai cần biên lợi nhuận chỉ kéo measure này vào; không ai phải tự nghĩ công thức nữa.
>
> DIVIDE thay cho dấu gạch chéo: nếu ô nào doanh thu bằng 0, DIVIDE trả về ô trống thay vì lỗi chia cho 0.

---

## Phần 4 — Filter context (15 phút · slide 18–20)

### Slide 18 · CHIA PHẦN · 0.5'
- **Trên slide:** MCI / **4. Filter context: ô này đang bị lọc bởi những gì?**

**Lời giảng:**
> Phần 4: khái niệm khó nhất khóa học, lần thứ ba mình gặp nó, và lần này gọi đúng tên.

### Slide 19 · KHÁI NIỆM · Filter context: Chuỗi C lần 3 · 7'
- **Trên slide:** [PQ PDF p73–75]
  - **Hình xương sống:** ô **90.210.000** ← [`Store[Region]` = Bắc] + [`Calendar[Nam]` = 2025] + [`Calendar[Quy]` = 1] + [`Calendar[Thang]` = 3] ← `SUM(Sales[Revenue])`.
  - Ba lần một ý: B2, tiêu chí của SUMIFS (`KPI_SUMIFS!D6`) · B3, tọa độ ô Pivot · B5, **filter context** của measure.
  - **Measure tính một ô qua 3 bước:** (1) đọc tọa độ của ô, hàng, cột, Filters, Slicer; (2) bộ lọc chảy theo relationship từ lookup xuống Sales; (3) tính công thức trên những dòng còn lại.
- **Ảnh:** [VẼ] hình xương sống Chuỗi C (dùng lại hình B2/B3, thêm nhãn "B5: filter context"). [CHỤP] S19: Pivot PV_Loc, Rows Region, Columns Thang, Filters Nam = 2025 và Quy = 1, Values Total Revenue; ô Bắc × 3 = 90.210.000 đang chọn.
- **Chân slide:** Ảnh: bản sao Practice/B5_MCI_Mart_Answer.xlsx, sheet PV_Loc (trên lớp: Demo/B5_Demo_DAX.xlsx). Đối chiếu: sheet KPI_SUMIFS, ô D6 và E9.

**Lời giảng:**
> *(→ trên PV_Loc: bỏ Gross Margin % và Total COGS khỏi Values, giữ Total Revenue; Rows: Store → Region; Columns: Calendar → Thang)* Pivot quen thuộc: miền theo tháng, Quý 1/2025. *(→ bấm ô Bắc × 3)* Ô này: 90.210.000.
>
> *(→ hiện hình xương sống)* Các bạn đã gặp hình này hai lần. Buổi 2, nó là SUMIFS với hai tiêu chí: miền bằng Bắc, tháng bằng 3. *(→ mở sheet KPI_SUMIFS trong file đang mở, chỉ ô D6)* Đúng 90.210.000. Buổi 3, nó là tọa độ của một ô Pivot. Hôm nay nó có tên chính thức trong DAX: **filter context**, ngữ cảnh lọc.
>
> **Câu hỏi cho lớp: ô 90.210.000 đang bị lọc bởi những gì?** *(→ chờ)* Bốn bộ lọc: Region bằng Bắc từ hàng, Thang bằng 3 từ cột, Nam bằng 2025 và Quy bằng 1 từ vùng Filters. Measure tính ô này qua ba bước. Một: đọc tọa độ. Hai: bộ lọc chảy theo đường nối, Region = Bắc chảy từ Store xuống Sales, tháng và năm chảy từ Calendar xuống Sales. Ba: cộng Revenue của những dòng còn lại.
>
> **Câu hỏi tiếp: còn ô Grand Total góc dưới bên phải thì sao?** *(→ chờ)* Không có miền, không có tháng, chỉ còn Nam = 2025 và Quy = 1. Nên nó là cả Quý 1: 941.454.000, khớp ô E9 của KPI_SUMIFS. Cùng một công thức `SUM(Sales[Revenue])`, mỗi ô ra một số khác nhau, vì **mỗi ô có ngữ cảnh lọc riêng**. Đó là toàn bộ bí mật của measure.

### Slide 20 · CÔNG CỤ · Lọc bằng cột của lookup · 7.5'
- **Tên (đỏ):** **Hide from Client Tools** (Power Pivot: chuột phải tiêu đề cột) [PQ PDF p56]
- **Mục đích:** Ẩn cột khỏi danh sách trường của Pivot, để người dùng không kéo nhầm cột của bảng fact vào Rows / Columns.
- **Demo, hai Pivot cạnh nhau** (Filters: `Calendar[Nam]` = 2025; Values: Total Revenue, Sum of Target):
  - Rows = **Calendar[Thang]**: T1 425.850.000 / 353.000.000 · T2 245.996.000 / 298.000.000 … T5–T12 chỉ có target. **Đúng.**
  - Rows = **Sales[Thang]** (Pivot chép sang bên phải chỉ giữ Sum of Target, bỏ Total Revenue): Target **3.665.000.000 ở mọi dòng**. **Sai.**
- **Quy tắc:** Rows, Columns, Filters, Slicer lấy cột của **lookup** (Calendar, Store, Product, Customer). Ẩn khỏi Client Tools: `Sales[Nam]`, `Sales[Thang]`, `Sales[Order_Date]`, `Sales[Store_Code]`, `Sales[Product_Code]`, `Sales[Customer_ID]`, `Target[Store_Code]`, `Target[Thang]`, `Target[Ngay_Dau_Thang]`.
- **Ảnh:** [CHỤP] S20: sheet PV_Lookup_vs_Fact, hai Pivot cạnh nhau; Pivot bên phải chỉ còn cột Sum of Target, lặp 3.665.000.000 ở mọi dòng (đã bỏ Total Revenue).
- **Chân slide:** Ảnh: bản sao Demo/B5_Demo_DAX.xlsx, sheet PV_Lookup_vs_Fact (tạo trong demo).

**Lời giảng:**
> Một cái bẫy rất hay gặp. Bảng Sales có sẵn cột Nam và Thang từ Buổi 4. Calendar cũng có Nam và Thang. Dùng cột nào cũng được chứ?
>
> *(→ Insert → PivotTable ▾ → From Data Model → New Worksheet, đổi tên sheet `PV_Lookup_vs_Fact`; Filters: Calendar → Nam = 2025; Rows: Calendar → Thang; Values: Total Revenue, Target → Target)* Pivot bên trái lấy tháng từ Calendar: tháng 1 doanh thu 425 triệu 850, target 353 triệu; tháng 2 doanh thu 246 triệu, target 298 triệu. Từ tháng 5 trở đi chỉ có target, vì dữ liệu bán hàng mới tới tháng 4.
>
> *(→ chép Pivot sang bên phải: chọn cả Pivot, Ctrl+C, dán ở cột G; trong Pivot mới: bỏ Calendar[Thang] và Total Revenue, kéo Sales → Thang vào Rows; giữ Sum of Target)* Pivot bên phải lấy tháng từ bảng Sales. Cột target, 3.665.000.000 ở mọi dòng, tháng nào cũng bằng cả năm.
>
> **Câu hỏi cho lớp: vì sao?** *(→ chờ)* Nhớ hướng lọc: bộ lọc chảy từ lookup xuống fact. Thang của Calendar chảy xuống được cả Sales lẫn Target. Thang của Sales thì nằm ở đáy một bảng fact, nó lọc được chính Sales, nhưng không chảy sang Target, vì hai fact không nối thẳng với nhau.
>
> Quy tắc: hàng, cột, bộ lọc, slicer **luôn lấy cột của bảng lookup**. *(→ xóa Pivot bên phải: chọn cả Pivot → Delete)* Và để không ai kéo nhầm, mình giấu các cột khóa và cột ngày của bảng fact đi. *(→ Power Pivot, Data View, bảng Sales: giữ Ctrl chọn tiêu đề Nam, Thang, Order_Date, Store_Code, Product_Code, Customer_ID → chuột phải → Hide from Client Tools; làm tương tự cho Target: Store_Code, Thang, Ngay_Dau_Thang)* Cột chuyển màu xám trong Data View, và biến mất khỏi danh sách trường của Pivot.

---

## Phần 5 — CALCULATE, ALL, Target (20 phút · slide 21–25)

### Slide 21 · CHIA PHẦN · 0.5'
- **Trên slide:** MCI / **5. CALCULATE và ALL: đổi bộ lọc có chủ đích**

**Lời giảng:**
> Phần 5: nếu filter context là bộ lọc Pivot đặt lên mỗi ô, thì CALCULATE là cách **mình** đổi bộ lọc đó.

### Slide 22 · CÔNG CỤ · CALCULATE · 5.5'
- **Tên (đỏ):** **CALCULATE(**`<biểu thức>`, `<bộ lọc 1>`, `<bộ lọc 2>`, …**)**
- **Mục đích:** Tính biểu thức dưới bộ lọc mình đặt; bộ lọc trong CALCULATE **ghi đè** bộ lọc cùng cột của ô Pivot, các bộ lọc khác giữ nguyên.
- **Ví dụ:** `Revenue Mien Nam := CALCULATE([Total Revenue], Store[Region] = "Nam")`
- **Thẻ số (Quý 1/2025):** ô Bắc, ô Trung, ô Nam của cột Revenue Mien Nam đều **427.663.000**.
- **Ảnh:** [CHỤP] S22: PV_Loc, Rows Region, Values Total Revenue và Revenue Mien Nam (cùng 427.663.000 ở cả 3 dòng). [PQ PDF p92–94]
- **TIP:** Demo xong, xóa measure này: Power Pivot → Measures → Manage Measures → Delete.
- **Chân slide:** Ảnh: bản sao Practice/B5_MCI_Mart_Answer.xlsx, sheet PV_Loc (trên lớp: Demo/B5_Demo_DAX.xlsx).

**Lời giảng:**
> *(→ trên PV_Loc: bỏ Calendar[Thang] khỏi Columns)* Pivot còn ba miền, Quý 1. *(→ New Measure: `Revenue Mien Nam` = `CALCULATE([Total Revenue], Store[Region]="Nam")`, định dạng số → OK; tick vào Values)*
>
> **Câu hỏi cho lớp: dòng Bắc, cột Revenue Mien Nam ra 427.663.000, bằng đúng doanh thu miền Nam. Pivot có sai không?** *(→ chờ)* Không sai. Ô đó có ngữ cảnh lọc Region = Bắc từ hàng. Nhưng CALCULATE nói: "bỏ bộ lọc Region đang có, thay bằng Region = Nam". Bộ lọc của CALCULATE thắng. Còn bộ lọc Nam = 2025, Quý 1 ở vùng Filters thì CALCULATE không đụng tới, nên vẫn là Quý 1.
>
> CALCULATE là hàm quan trọng nhất của DAX: **giữ nguyên bộ lọc này, thay bộ lọc kia**. Hai slide tới đều dựa vào nó. *(→ Power Pivot → Measures → Manage Measures → chọn Revenue Mien Nam → Delete)* Measure này chỉ để minh họa, mình xóa đi.

### Slide 23 · CÔNG CỤ · ALL → % of Total · 5'
- **Tên (đỏ):** **ALL(**`<bảng hoặc cột>`**)**
- **Mục đích:** Bỏ mọi bộ lọc trên bảng (hoặc cột) đó, dùng bên trong CALCULATE để lấy "mẫu số chung".
- **Measure:**
  - `Revenue All Stores := CALCULATE([Total Revenue], ALL(Store))`
  - `% Revenue Share := DIVIDE([Total Revenue], [Revenue All Stores])`, Format Percentage, 1 chữ số thập phân
- **Thẻ số (Quý 1/2025):** Bắc **34,4%** · Nam **45,4%** · Trung **20,2%**, đúng Pct_Share của Buổi 1.
- **Ảnh:** [CHỤP] S23: PV_Loc, Rows Region, Values Total Revenue, Revenue All Stores (941.454.000 ở mọi dòng), % Revenue Share. [PQ PDF p99–100]
- **TIP:** `ALL(Store)` bỏ lọc mọi cột của Store (miền, cửa hàng, loại cửa hàng); bộ lọc của Calendar vẫn giữ, nên đây là % **trong Quý 1**. Tỷ trọng theo ngành: đổi thành `ALL(Product)`.
- **Chân slide:** Ảnh: bản sao Practice/B5_MCI_Mart_Answer.xlsx, sheet PV_Loc (trên lớp: Demo/B5_Demo_DAX.xlsx).

**Lời giảng:**
> Buổi 1 các bạn tính phần trăm đóng góp của từng miền: doanh thu miền chia ô tổng, và phải khóa ô tổng bằng dấu đô-la để kéo công thức không bị trượt. Trong DAX, "khóa mẫu số" làm bằng **ALL**.
>
> *(→ New Measure: `Revenue All Stores` = `CALCULATE([Total Revenue], ALL(Store))`, định dạng như Total Revenue: Category Number, tick Use 1000 separator, Decimal places 0 → OK; tick vào Values)* Mọi dòng đều 941.454.000. ALL(Store) xóa bộ lọc miền khỏi ô, nên mẫu số luôn là tổng các cửa hàng. Bộ lọc Quý 1 thì vẫn còn.
>
> *(→ New Measure: `% Revenue Share` = `DIVIDE([Total Revenue], [Revenue All Stores])`, Format Percentage → OK; tick vào Values)* Bắc 34,4%, Nam 45,4%, Trung 20,2%. **Câu hỏi cho lớp: các bạn đã thấy ba con số này ở đâu?** *(→ chờ)* Buổi 1, cột Pct_Share. Cùng một sự thật, giờ là một measure: đặt nó vào Pivot theo cửa hàng hay theo loại cửa hàng, nó tự tính lại, không cần dấu đô-la nào. ALL(Store) chỉ bỏ bộ lọc của bảng Store; muốn tỷ trọng theo ngành hàng thì mẫu số phải bỏ lọc bảng Product, tức là đổi ALL(Store) thành ALL(Product).

### Slide 24 · CÔNG CỤ · Measure cho Target · 6'
- **Tên (đỏ):** **IF(ISBLANK(**…**), BLANK(), …)**, chặn ô chưa có dữ liệu
- **Measure:**
  - `Total Target := SUM(Target[Target])`, Table name: **Target**
  - `Revenue vs Target % := DIVIDE([Total Revenue], [Total Target])`
  - `Variance := IF(ISBLANK([Total Revenue]), BLANK(), [Total Revenue] - [Total Target])`
- **Thẻ số (2025):** T1 **120,6%** (+72.850.000) · T2 **82,5%** (−52.004.000) · T3 **94,6%** · T4 **103,1%** · T1 + T2 gộp **103,2%**.
- **Ảnh:** [CHỤP] S24: sheet PV_Target, Filters Nam = 2025, Rows Calendar[Thang], Values Total Revenue, Total Target, Revenue vs Target %, Variance; T5–T12 chỉ có Total Target. [PQ PDF p92]
- **TIP:** Không có IF(ISBLANK…), Variance tháng 5 = **−288.000.000**, "hụt 288 triệu" chỉ vì tháng 5 chưa có dữ liệu.
- **Lưu ý dòng Grand Total:** ra **33,8%** và **−2.426.581.000** vì lấy doanh thu 4 tháng (1.238.419.000) chia cho target cả năm 3.665.000.000. Không đọc dòng tổng này như kết quả đạt target.
- **Chân slide:** Ảnh: bản sao Practice/B5_MCI_Mart_Answer.xlsx, sheet PV_Target (trên lớp: Demo/B5_Demo_DAX.xlsx).

**Lời giảng:**
> Giờ trả lời nửa đầu câu hỏi mở buổi. *(→ New Measure, Table name: Target: `Total Target` = `SUM(Target[Target])`, định dạng số như Total Revenue (Number, Use 1000 separator, 0 chữ số thập phân) → OK)* Measure target nằm ở bảng Target cho dễ tìm. *(→ New Measure, Table name: Sales: `Revenue vs Target %` = `DIVIDE([Total Revenue],[Total Target])`, Format Percentage → OK)*
>
> *(→ Insert → PivotTable ▾ → From Data Model → New Worksheet, đổi tên `PV_Target`; Filters: Calendar → Nam = 2025; Rows: Calendar → Thang; Values: Total Revenue, Total Target, Revenue vs Target %)* Tháng 1: 120,6%. Tháng 2: 82,5%. Tháng 3: 94,6%. Tháng 4: 103,1%.
>
> Thêm chênh lệch tuyệt đối. *(→ New Measure: `Variance` = `[Total Revenue]-[Total Target]`, định dạng số như Total Revenue → OK; tick vào Values)* Tháng 1 vượt 72 triệu 850. Nhưng nhìn tháng 5: **âm 288 triệu**. **Câu hỏi cho lớp: tháng 5 bán kém đến vậy sao?** *(→ chờ)* Tháng 5 chưa có dữ liệu. Doanh thu trống, trừ đi target thì ra số âm. Một con số như vậy mà lên báo cáo thì cả phòng kinh doanh bị gọi lên. *(→ Manage Measures → Edit Variance: `=IF(ISBLANK([Total Revenue]), BLANK(), [Total Revenue]-[Total Target])` → OK)* Chưa có doanh thu thì để trống. Tháng 5 trở đi giờ chỉ còn target. Dòng Grand Total ở đáy thì đừng đọc: 33,8% và âm 2 tỷ 426 triệu là doanh thu 4 tháng đem so với target cả năm, hai kỳ không cùng độ dài.
>
> Còn một chuyện lạ: tháng 1 vượt khoảng 21%, tháng 2 hụt khoảng 17%. *(→ chờ một nhịp)* Buổi 4 các bạn đã thấy trong Calendar: Tết 2025 rơi vào 29/01, Tết 2024 rơi vào 10/02. Năm nay cả mùa mua sắm trước Tết nằm trọn trong tháng 1, còn kỳ nghỉ Tết kéo tháng 2 xuống. Năm ngoái mùa mua sắm trải qua cả cuối tháng 1 lẫn đầu tháng 2. Target lại được lập bằng doanh thu **cùng tháng dương lịch năm trước** nhân 110% (ST12 chưa có năm 2024 nên target đặt cố định 25 triệu), nên tháng 1 bị đặt thấp, tháng 2 bị đặt cao. Gộp hai tháng lại thì đạt 103,2%. Tháng 2 không bán kém, **target lệch mùa Tết**. Đây là câu trả lời Tầng 3 cho chữ "vì sao".

### Slide 25 · KHÁI NIỆM · Target không lọc theo ngành · 3'
- **Trên slide:**
  - Rows = **Product[Category]**, Quý 1: Bánh kẹo 36,2% · Gia dụng 21,7% · … · Sữa 1,8% = doanh thu ngành ÷ target **cả chuỗi** (936.000.000) → **vô nghĩa**.
  - Rows = **Store[Store_Code]** × Columns = **Calendar[Thang]**: **ST07 tháng 3: 8.539.000 / 19.000.000 = 44,9%** (cải tạo 10–31/03) → tháng 4: **97,4%**.
  - KPI dictionary ghi rõ: Revenue vs Target % lọc theo store, region, month, **không** theo product/category.
- **Ảnh:** [CHỤP] S25: PV_Target, Rows Store_Code, Columns Thang (1–4), Values Revenue vs Target %; ô ST07 × 3 = 44,9% đang chọn.
- **Chân slide:** Ảnh: bản sao Practice/B5_MCI_Mart_Answer.xlsx, sheet PV_Target_CuaHang (trên lớp: sheet PV_Target, Demo/B5_Demo_DAX.xlsx).

**Lời giảng:**
> *(→ trên PV_Target: Rows đổi thành Product → Category; Filters thêm Calendar → Quy = 1)* Theo ngành, Quý 1: Bánh kẹo đạt 36%, Sữa đạt chưa tới 2%. **Câu hỏi cho lớp: ngành Sữa thảm hại vậy sao?** *(→ chờ)* Không. Target không có ngành, nên mẫu số ở mọi dòng là target của cả chuỗi. Con số này không có nghĩa gì. Measure không báo lỗi, **người dùng phải biết giới hạn của nó**, và mình ghi giới hạn đó vào KPI dictionary ở Phần 7.
>
> *(→ bỏ Quy khỏi Filters; Rows: Store → Store_Code; Columns: Calendar → Thang; Values chỉ giữ Revenue vs Target %; Row Labels/Column Labels lọc tháng 1–4)* Theo cửa hàng và tháng thì có nghĩa. *(→ bấm ô ST07 × 3)* ST07 tháng 3: 44,9%. Buổi 3 các bạn đã dùng Show Details để thấy cửa hàng Nha Trang cải tạo từ 10 đến 31/03. Tháng 4 lên lại 97,4%, đúng con số bài về nhà Buổi 4. Một câu chuyện, ba buổi, ba công cụ, cùng một kết luận.

---

## Phần 6 — Time intelligence: lũy kế và cùng kỳ (15 phút · slide 26–28)

### Slide 26 · CHIA PHẦN · 0.5'
- **Trên slide:** MCI / **6. Time intelligence: lũy kế và so với cùng kỳ**

**Lời giảng:**
> Phần 6: hai câu hỏi sếp nào cũng hỏi, từ đầu năm tới giờ bao nhiêu, và so với năm ngoái thế nào.

### Slide 27 · CÔNG CỤ · YTD · cùng kỳ năm trước · 8'

| Tên (đỏ) | Mục đích |
|---|---|
| **Design → Mark as Date Table** (cửa sổ Power Pivot) | Khai báo Calendar là bảng lịch của mô hình; cột Date phải có **mỗi ngày một dòng, liên tục, không trống** (Calendar: 731 ngày ✔) |
| **DATESYTD(**`Calendar[Date]`**)** | Các ngày từ đầu năm đến ngày cuối của ô đang tính - dùng trong CALCULATE để ra lũy kế |
| **DATEADD(**`Calendar[Date]`, -1, YEAR**)** | Dịch các ngày của ô lùi một năm - ra cùng kỳ năm trước |

- **Measure:**
  - `Revenue YTD := CALCULATE([Total Revenue], DATESYTD(Calendar[Date]))`
  - `Revenue LY := CALCULATE([Total Revenue], DATEADD(Calendar[Date], -1, YEAR))`
  - `Revenue YoY % := IF(ISBLANK([Total Revenue]), BLANK(), DIVIDE([Total Revenue] - [Revenue LY], [Revenue LY]))`
- **Thẻ số (2025):** YoY T1 **+42,8%** · T2 **−0,9%** · T3 **+13,7%** · T4 **+24,0%** · YTD tháng 4 **1.238.419.000** (cùng kỳ 2024: 1.023.165.000 → **+21,0%**).
- **Ảnh:** [CHỤP] S27a: hộp Mark as Date Table, Date column = Date. [CHỤP] S27b: sheet PV_YoY, Filters Nam = 2025, Rows Calendar[Thang] (lọc 1–4), Values Total Revenue, Revenue LY, Revenue YoY %, Revenue YTD. [PQ PDF p105–107]
- **TIP:** Time intelligence chỉ đúng khi hàng / cột lấy từ **Calendar**. Không chặn ô trống thì YoY tháng 5 = **−100%**.
- **PRO TIP:** `SAMEPERIODLASTYEAR(Calendar[Date])` cho cùng kết quả với DATEADD lùi một năm; DATESQTD, DATESMTD cho lũy kế quý, tháng.
- **Chân slide:** Ảnh: Demo/B5_Demo_DAX.xlsx (Design → Mark as Date Table); bản sao Practice/B5_MCI_Mart_Answer.xlsx, sheet PV_YoY.

**Lời giảng:**
> Trước hết, cho mô hình biết đâu là bảng lịch. *(→ cửa sổ Power Pivot, tab Calendar → Design → Mark as Date Table → Date column: Date → OK)* Điều kiện: mỗi ngày đúng một dòng, không bỏ sót ngày nào. Calendar Buổi 4 dựng từ 01/01/2024 đến 31/12/2025, bảy trăm ba mươi mốt ngày liên tục, đạt.
>
> *(→ New Measure: `Revenue YTD` = `CALCULATE([Total Revenue], DATESYTD(Calendar[Date]))`, định dạng số như Total Revenue → OK)* Lại là CALCULATE: giữ năm của ô, nhưng thay bộ lọc ngày bằng "từ 1/1 tới ngày cuối của ô". *(→ New Measure: `Revenue LY` = `CALCULATE([Total Revenue], DATEADD(Calendar[Date], -1, YEAR))`, định dạng số như Total Revenue → OK)* Và cùng kỳ năm trước: dịch các ngày của ô lùi một năm.
>
> *(→ Insert → PivotTable ▾ → From Data Model → New Worksheet, đổi tên `PV_YoY`; Filters: Calendar → Nam = 2025; Rows: Calendar → Thang; Values: Total Revenue, Revenue LY, Revenue YTD)* Tháng 1/2025: 425 triệu 850; cùng kỳ năm trước 298 triệu 203. Lũy kế tới tháng 4: 1 tỷ 238 triệu.
>
> *(→ New Measure: `Revenue YoY %` = `DIVIDE([Total Revenue]-[Revenue LY],[Revenue LY])`, Format Percentage → OK; tick vào Values)* Tháng 1 tăng 42,8%, tháng 2 giảm 0,9%. **Câu hỏi cho lớp: tháng 1 tăng gần 43% so với năm ngoái, có đáng mở tiệc không?** *(→ chờ)* Cũng là Tết: tháng 1 năm nay có Tết, tháng 1 năm ngoái thì không. So tháng 1 với tháng 1 là so hai thứ khác nhau. Gộp lũy kế tới tháng 4 cho bức tranh công bằng hơn: tăng 21%.
>
> Kéo xuống tháng 5: YoY **âm 100%**. Cùng cái bẫy với Variance, chưa có doanh thu, trừ năm ngoái thì ra âm hết. *(→ Edit Revenue YoY %: `=IF(ISBLANK([Total Revenue]), BLANK(), DIVIDE([Total Revenue]-[Revenue LY],[Revenue LY]))` → OK; Row Labels lọc tháng 1–4)* Chặn ô trống, rồi lọc bảng về bốn tháng có dữ liệu.

### Slide 28 · KHÁI NIỆM · Like-for-like: ST12 và ô YoY trống · 6.5'
- **Trên slide:**
  - YoY T1–T4/2025 theo cửa hàng: cao nhất **ST10 +37,9%** · âm: ST07 −3,2%, ST04 −1,5% · **ST12: Revenue LY trống, YoY trống**.
  - ST12 khai trương **01/06/2024**: T1–T4/2025 có **79.399.000**, T1–T4/2024 không có gì để so.
  - Tổng YTD +21,0% **gồm cả** 79.399.000 không có kỳ gốc → **like-for-like** (11 cửa hàng có đủ hai kỳ): **+13,3%**.
- **Ảnh:** [CHỤP] S28: PV_YoY, Rows Store_Code, Filters Nam = 2025 và Thang 1–4, Values Total Revenue, Revenue LY, Revenue YoY %; dòng ST12 có hai ô trống.
- **TIP:** D2: thấy ô trống, hỏi vì sao; thấy tổng đẹp, hỏi tổng gồm những gì. ST10 +37,9% · ST07 −3,2% · ST04 −1,5%.
- **PRO TIP (bài về nhà):** like-for-like thành measure, không cần lọc tay: `Revenue LFL := CALCULATE([Total Revenue], Store[Open_Date] < DATE(2024,1,1))`.
- **Chân slide:** Ảnh: bản sao Practice/B5_MCI_Mart_Answer.xlsx, sheet PV_YoY (trên lớp: Demo/B5_Demo_DAX.xlsx).

**Lời giảng:**
> *(→ trên PV_YoY: Rows đổi thành Store → Store_Code; Filters: Calendar → Thang chọn 1, 2, 3, 4; Values: Total Revenue, Revenue LY, Revenue YoY %)* YoY bốn tháng đầu năm, theo cửa hàng. ST10 tăng mạnh nhất, gần 38%. ST07 giảm 3%, tháng cải tạo kéo xuống. Và ST12: **hai ô trống**.
>
> **Câu hỏi cho lớp: ST12 mất dữ liệu năm ngoái à?** *(→ chờ)* Buổi 4 các bạn đã có câu trả lời: ST12 khai trương 01/06/2024. Bốn tháng đầu năm 2024 nó chưa tồn tại. Revenue LY trống, DIVIDE gặp mẫu số trống thì trả về trống, không phải lỗi, mà là sự thật.
>
> Nhưng chuyện chưa hết. Dòng tổng ghi tăng 21%. Trong đó có 79 triệu của ST12, doanh thu năm nay **không có kỳ gốc**. Cộng nó vào là so một chuỗi mười hai cửa hàng với một chuỗi mười một cửa hàng. *(→ bấm nút lọc ở ô Store_Code (Row Labels) của Pivot, bỏ tick ST12 → OK; đọc dòng Grand Total)* So **like-for-like**, chỉ các cửa hàng có đủ hai kỳ: tăng **13,3%**. Khác 21% rất nhiều.
>
> Đây là kỹ năng D2: thấy ô trống, hỏi vì sao trước khi kết luận; thấy tổng đẹp, hỏi tổng đó gồm những gì. *(→ bấm lại nút lọc Store_Code → tick (Select All) → OK)* Lọc tay như vừa rồi thì mỗi lần mở báo cáo lại phải nhớ bỏ ST12. Trong bài về nhà, các bạn biến nó thành measure: `CALCULATE([Total Revenue], Store[Open_Date] < DATE(2024,1,1))`, chỉ cộng các cửa hàng mở trước năm 2024. Công thức nằm ở PRO TIP trên slide.

---

## Phần 7 — KPI dictionary (10 phút · slide 29–30)

### Slide 29 · CHIA PHẦN · 0.5'
- **Trên slide:** MCI / **7. KPI dictionary: một định nghĩa, mọi báo cáo**

**Lời giảng:**
> Phần 7: quay lại nửa sau câu hỏi mở buổi, vì sao mỗi báo cáo một số.

### Slide 30 · KHÁI NIỆM · KPI dictionary · 9.5'
- **Trên slide:** (theo Course Plan, Phụ lục C)

| KPI | Measure (DAX) | Dimensions | Owner | Target |
|---|---|---|---|---|
| Total Revenue | `SUM(Sales[Revenue])` | region, store, category, month | Sales Director | = Total Target |
| Total Orders | `DISTINCTCOUNT(Sales[Order_ID])` | region, store, month | Sales Director | 1.000 đơn/tháng |
| Avg Order Value | `DIVIDE([Total Revenue], [Total Orders])` | region, store, month | Sales Director | 300.000 |
| Gross Margin % | `DIVIDE([Total Revenue] - [Total COGS], [Total Revenue])` | category, region, month | Finance | 33% |
| Revenue YoY % | `IF(ISBLANK([Total Revenue]), BLANK(), DIVIDE([Total Revenue] - [Revenue LY], [Revenue LY]))` | region, category, month | Finance | +10% |
| Active Customers | `DISTINCTCOUNT(Sales[Customer_ID])` | region, segment, month | Marketing | 700 KH/tháng |
| Revenue vs Target % | `DIVIDE([Total Revenue], [Total Target])` | store, region, month - **không** theo category | Sales Director | 100% |

  - Measure = **semantic layer**: KPI được định nghĩa **một lần** trong mô hình; mọi Pivot, mọi dashboard (Buổi 6) chỉ gọi tên measure.
  - Ô **Description** của hộp Measure: ghi định nghĩa và giới hạn ngay trong mô hình.
  - **D2, kiểm tra giả định (câu chuyện #5):** Gross Margin % theo năm: 2024 **31,2%** → 2025 **34,1%**.
- **Ảnh:** [VẼ] bảng KPI dictionary. [CHỤP] S30: PV_YoY, Rows Calendar[Nam], Values Gross Margin % (31,2% · 34,1%).
- **Chân slide:** Course Plan, Phụ lục C · Ảnh: bản sao Practice/B5_MCI_Mart_Answer.xlsx, sheet PV_GM_Nam (trên lớp: sheet PV_YoY, Demo/B5_Demo_DAX.xlsx).

**Lời giảng:**
> Quay lại hai con số đầu buổi: 34,18% và 33,77%. Hai báo cáo khác nhau vì **mỗi người tự viết công thức**. Cách chữa không phải là họp để chọn một công thức, mà là đặt công thức vào **một chỗ duy nhất**: measure trong Data Model. Từ nay ai cần biên lợi nhuận đều kéo measure Gross Margin %; không ai có cơ hội tính trung bình của trung bình nữa.
>
> Bảng trên slide gọi là **KPI dictionary**, từ điển KPI. Các bạn đã làm KPI Definition Card ở Buổi 2; giờ mỗi card có thêm một dòng quan trọng nhất: **measure nào** thực hiện định nghĩa đó. Để ý cột Dimensions: Revenue vs Target % ghi rõ không theo ngành, bài học của Slide 25. *(→ Manage Measures → Edit Gross Margin % → ô Description gõ: "(Doanh thu − giá vốn) ÷ doanh thu. Giá vốn: một bảng Product[Unit_Cost] cho mọi năm." → OK)* Định nghĩa và giới hạn nằm ngay trong mô hình.
>
> Nói đến giới hạn. *(→ PV_YoY: xóa mọi bộ lọc ở vùng Filters (Nam, Thang); Rows: Calendar → Nam; Values chỉ giữ Gross Margin %)* Biên lợi nhuận gộp năm 2024: 31,2%. Năm 2025: 34,1%. **Câu hỏi cho lớp: biên lợi nhuận tăng gần 3 điểm, báo cáo lên ban giám đốc được chưa?** *(→ chờ)* Khoan. Kiểm tra giả định trước. Giá vốn lấy từ đâu? Từ một bảng Product duy nhất, giá vốn hiện tại, áp cho cả hai năm. Trong khi giá bán năm 2024 thấp hơn năm nay khoảng 4%. Lấy giá vốn năm nay trừ vào giá bán năm ngoái thì biên năm ngoái tất nhiên thấp. Phần "tăng" này phần lớn là do **cách dữ liệu được lưu**, không phải do kinh doanh tốt hơn. Đó chính là thứ phải ghi vào cột "Giới hạn" của KPI card. Đây là kỹ năng D2 ở mức Competent: kiểm giả định trước khi kết luận.

---

## Phần 8 — Thực hành: 6 KPI và đối soát với Buổi 2 (20 phút · slide 31–32)

### Slide 31 · THỰC HÀNH 1 · Hoàn thành bộ 6 KPI (10 phút)
- **Trên slide:**
  1. Viết 3 measure còn lại (Table name: Sales): `Total Orders := DISTINCTCOUNT(Sales[Order_ID])` · `Avg Order Value := DIVIDE([Total Revenue], [Total Orders])` · `Active Customers := DISTINCTCOUNT(Sales[Customer_ID])`, định dạng số, dấu ngăn nghìn.
  2. Sheet mới **PV_KPI_Thang** (PivotTable From Data Model): Filters `Calendar[Nam]` = 2025 · Rows `Calendar[Thang]` (lọc 1–4) · Values: Total Revenue, Total Orders, Avg Order Value, Gross Margin %, Revenue YoY %, Active Customers.
  3. Kiểm tay AOV: Filters thêm `Product[Category]` = Bánh kẹo, `Calendar[Quy]` = 1 → đọc dòng Grand Total: Total Revenue ÷ Total Orders = Avg Order Value?
- **Kiểm tra:**

| Tháng | Revenue | Orders | AOV | GM % | YoY % | Active Customers |
|---|---|---|---|---|---|---|
| 1 | 425.850.000 | 1.398 | 304.614 | 35,2% | +42,8% | 831 |
| 2 | 245.996.000 | 823 | 298.902 | 33,7% | −0,9% | 605 |
| 3 | 269.608.000 | 923 | 292.100 | 33,1% | +13,7% | 637 |
| 4 | 296.965.000 | 966 | 307.417 | 33,8% | +24,0% | 670 |

  Bánh kẹo, Quý 1: 338.733.000 ÷ **1.739** đơn = **194.786**.

**Lời giảng (giao bài):**
> *(→ học viên tự làm 10 phút)* Bốn KPI đã có: doanh thu, biên lợi nhuận, YoY, và % target. Các bạn tự viết ba measure còn lại, công thức trên slide. Rồi dựng một Pivot KPI theo tháng và so với bảng kiểm tra.
>
> Bước 3 là bài kiểm tra quan trọng: lọc riêng ngành Bánh kẹo, Quý 1, rồi đọc dòng Grand Total. Lấy Total Revenue chia Total Orders bằng máy tính, có ra đúng Avg Order Value không?

**Ghi chú giảng viên:**
- Bảng kiểm tra trên slide lấy từ `Practice/B5_MCI_Mart_Answer.xlsx`, sheet `PV_KPI_Thang`. Chữa bài thì mở sheet này trên máy giảng viên.
- Đếm số đơn bằng **DISTINCTCOUNT**, không phải COUNTROWS: một đơn có nhiều dòng (Line_No). COUNTROWS cho Quý 1 ra 6.187 (số dòng), không phải 3.144 đơn.
- Bánh kẹo, Quý 1: 1.739 đơn có mua bánh kẹo; AOV 194.786 = giá trị phần bánh kẹo trong mỗi đơn đó. Học viên dùng `AVERAGE(Sales[Revenue])` sẽ ra **153.830**, trung bình một **dòng**, không phải một **đơn**.
- **Active Customers không cộng dồn:** Quý 1 theo miền 492 + 657 + 307 = 1.456, nhưng Grand Total là **1.208**, khách mua ở nhiều miền chỉ đếm một lần. Câu 4 của quiz dùng chi tiết này.
- Revenue YoY % trống ở mọi dòng: Rows đang dùng `Sales[Thang]` (chưa ẩn cột ở Slide 20), hoặc chưa Mark as Date Table.

### Slide 32 · THỰC HÀNH 2 · Đối soát với KPI_SUMIFS (10 phút)
- **Trên slide:**
  1. Sheet mới **PV_DoiSoat_B2** (PivotTable From Data Model): Filters `Calendar[Nam]` = 2025, `Calendar[Quy]` = 1 · Rows `Store[Region]` · Columns `Calendar[Thang]` · Values: Total Revenue (rồi lần lượt Total Orders, Avg Order Value, Gross Margin %). Đây là báo cáo PV_MienThang của Buổi 3, nay đọc từ Data Model, GM% không cần Calculated Field.
  2. Đối soát 4 ô then chốt bằng `GETPIVOTDATA` (gõ `=`, bấm vào ô Pivot) trừ ô tương ứng của `KPI_SUMIFS`: Bắc × 3 (D6) · Doanh thu Quý 1 (E9) · Số đơn Quý 1 (E16) · GM% Quý 1 (E30). Kết quả phải bằng **0**.
  3. Checklist brick output B5:
     - [ ] Diagram View: hình ngôi sao, **6 bảng, 6 relationship**, hai fact Sales và Target gặp nhau qua Store và Calendar.
     - [ ] 6 KPI + Total Target, Revenue vs Target %, Variance là **explicit measure**, có định dạng.
     - [ ] Đối soát với KPI_SUMIFS: chênh lệch **0**.
- **Kiểm tra:** Quý 1/2025: **941.454.000** · **3.144** đơn · AOV **299.445** · GM **34,18%** · Bắc × 3 **90.210.000**.
- **Chân slide:** Làm mẫu: Practice/B5_MCI_Mart_Answer.xlsx, sheet PV_DoiSoat_B2 (ô J11, N14) · đủ 64 ô Miền × Tháng × 4 KPI: sheet DoiSoat_B2 cùng file.

**Lời giảng:**
> *(→ đi vòng lớp)* Đây là khoảnh khắc chốt buổi hôm nay. Bảng KPI_SUMIFS các bạn làm ở Buổi 2 bằng bốn khối SUMIFS và công thức chia. Hôm nay cùng bảng đó ra từ Data Model, bằng bốn measure. **Hai cách làm khác hẳn nhau, số có khớp không?**
>
> *(→ trên máy giảng viên: mở `Practice/B5_MCI_Mart_Answer.xlsx`, sheet `PV_DoiSoat_B2`: Pivot Miền × Tháng dựng sẵn ở A8:Q14, đúng bước 1 trên slide)* Mình làm mẫu hai ô trên file đáp án. *(→ ô S11 gõ `=`, bấm ô J11 (Bắc × tháng 3, Total Revenue), gõ `-KPI_SUMIFS!D6` → Enter; sheet KPI_SUMIFS nằm trong file đang mở)* Bằng 0. *(→ ô S14 gõ `=`, bấm ô N14 (Grand Total, Total Revenue Quý 1), gõ `-KPI_SUMIFS!E9` → Enter)* Bằng 0. Hai ô còn lại các bạn tự làm trên file của mình: số đơn Quý 1 trừ ô E16, biên lợi nhuận Quý 1 trừ ô E30. Bạn nào có số khác 0 thì kiểm tra lại vùng Filters: phải là Nam 2025 **và** Quý 1, lấy từ Calendar.
>
> **Khớp nghĩa là Buổi 2 và Buổi 5 cùng nói một sự thật.** Từ nay các bạn có một nguồn duy nhất: mọi báo cáo mới chỉ cần kéo measure vào. *(→ trong cùng file, mở sheet `DoiSoat_B2`)* Sheet này đối soát đủ 64 ô, bốn KPI, ba miền và dòng tổng, ba tháng và Quý 1. Đây cũng là file khởi đầu Buổi 6. *(→ đóng file, không lưu: hai ô S11, S14 vừa gõ không được nằm lại trong file khởi đầu Buổi 6)*

---

## Phần 9 — Quiz, insight, preview (5 phút · slide 33–39)

Quiz kiểu hỏi nhanh, khoảng 35 giây mỗi câu. Buổi này chỉ còn 5 phút, nên quiz có **4 câu, câu 3–4 là tình huống**.

### Slide 33 · QUIZ · EXCEL QUIZ CHALLENGE · 0.2'

**Lời giảng:**
> Bốn câu hỏi nhanh. Câu 3 và 4 là hai con số lạ các bạn sẽ gặp ngoài đời.

### Slide 34 · QUESTION 1: MEASURE VS COLUMN
- KPI Gross Margin % nên viết thế nào trong Data Model? A) Calculated column GM từng dòng, rồi lấy Average · B) Calculated column GM từng dòng, rồi lấy Sum · **C) Measure: DIVIDE(tổng doanh thu − tổng giá vốn, tổng doanh thu)** · D) Implicit measure Average of Revenue

**Lời giảng:**
> C. Tỷ lệ phải tính tổng trước, chia sau, cho từng ô Pivot, đó là measure. Average của cột GM từng dòng ra 33,77% thay vì 34,18%; Sum thì vô nghĩa.

### Slide 35 · QUESTION 2: HAI BẢNG FACT
- Sales và Target đều có cột Store_Code. Nối hai bảng này thế nào để phân tích doanh thu so với target theo cửa hàng? A) Tạo relationship thẳng giữa Sales[Store_Code] và Target[Store_Code] · **B) Nối cả hai vào bảng lookup chung Store (và Calendar); lọc qua lookup** · C) Merge Target vào Sales bằng Power Query · D) Không nối được, phải dùng SUMIFS

**Lời giảng:**
> B. Hai fact không nối thẳng với nhau, Store_Code lặp lại ở cả hai bảng, quan hệ đó là nhiều–nhiều. Cả hai nối vào Store và Calendar; bộ lọc từ hai lookup chung chảy xuống cả hai fact.

### Slide 36 · QUESTION 3: TÌNH HUỐNG: CON SỐ LẶP LẠI
- Pivot Quý 1/2025: Rows = Product[Category], Values = Total Target. Sáu dòng đều ra **936.000.000**. Vì sao? A) Thiếu Mark as Date Table · B) Measure Total Target viết sai · **C) Target có grain tháng × cửa hàng, không có chiều sản phẩm; Product không lọc được Target** · D) Phải Refresh All

**Lời giảng:**
> C. Cùng một con số ở mọi dòng nghĩa là trường ở Rows không lọc được bảng ở Values. Với Target, đó là thiết kế: kế hoạch không lập theo ngành. Ghi giới hạn này vào KPI dictionary và đừng đặt ngành hàng cạnh Revenue vs Target %.

### Slide 37 · QUESTION 4: TÌNH HUỐNG: TỔNG LỆCH
- Pivot Quý 1/2025 theo miền: Active Customers Bắc 492, Nam 657, Trung 307, cộng lại 1.456. Nhưng Grand Total ghi **1.208**. Kết luận nào đúng? A) Pivot bị lỗi, phải Refresh · B) Thiếu relationship Sales – Customer · **C) Không lỗi: một khách mua ở nhiều miền chỉ được đếm một lần ở dòng tổng; DISTINCTCOUNT không cộng dồn** · D) Phải đổi sang SUM(Sales[Customer_ID])

**Lời giảng:**
> C. Dòng tổng có ngữ cảnh lọc riêng, "cả ba miền", và đếm khách không trùng trên cả ba miền. Đếm phân biệt, giống tỷ lệ, không cộng dồn được. Đây là lý do mọi KPI phải là measure: Pivot cộng các dòng thì sai, measure tính lại cho ô tổng thì đúng.

### Slide 38 · TỔNG KẾT + INSIGHT · 1.7'
- **Trên slide:**
  - **Điểm chính cần nhớ:**
    - (1) Star schema: fact ở giữa, lookup xung quanh, quan hệ 1-\*; bộ lọc chảy từ lookup xuống fact. Hai fact gặp nhau qua lookup chung.
    - (2) Calculated column tính từng dòng; **measure** tính từng ô. KPI tỷ lệ và đếm phân biệt luôn là measure.
    - (3) Filter context = tọa độ của ô. **CALCULATE** đổi bộ lọc; **ALL** bỏ bộ lọc.
    - (4) Time intelligence cần Calendar liên tục; hàng / cột lấy từ Calendar, không lấy từ fact.
    - (5) KPI dictionary: mỗi KPI một measure, một định nghĩa, kèm giới hạn, và đối soát với nguồn trước khi dùng.
  - **Insight của bạn (1 câu):** "Theo Revenue vs Target %, ___ *(con số)*, nguyên nhân khả dĩ là ___, nên ___ *(hàm ý)*."
  - Gợi ý: so tháng 1 với tháng 2, rồi mở Calendar xem Tết rơi vào ngày nào.

**Lời giảng:**
> Năm điều cần nhớ *(→ đọc lướt 5 ý trên slide)*. Viết câu insight trong một phút. *(→ chờ, gọi 1–2 bạn đọc)*
>
> Câu mẫu: "Tháng 1/2025 đạt 120,6% target nhưng tháng 2 chỉ 82,5%; gộp hai tháng đạt 103,2%. Nguyên nhân khả dĩ: Tết 2025 rơi vào 29/01 còn Tết 2024 rơi vào 10/02, trong khi target lập theo cùng tháng dương lịch năm trước nhân 110%, mua sắm Tết dồn sang tháng 1. Nên đánh giá tháng 1 và 2 gộp, không kết luận tháng 2 bán kém." Câu này đi từ con số, qua nguyên nhân đã kiểm chứng, tới hàm ý. Bước tiếp theo, **nên lập target thế nào**, là câu hỏi Tầng 4 của Buổi 6.

### Slide 39 · PREVIEW B6 · 0.8'
- **Trên slide:**
  - *Buổi sau:* "Giám đốc bán hàng có 5 giây. Họ cần thấy gì đầu tiên?" (Dashboard)
  - Hôm nay số đã đúng và thống nhất. Buổi 6: **trình bày** cho người ra quyết định, KPI card bằng CUBEVALUE đọc thẳng các measure hôm nay.
  - *Về nhà:* **Bài tập về nhà (bắt buộc)** `B5_BTVN_MCI_Mart_KPI_T1_T4_2025.xlsx`: báo cáo tháng 1–4 của phòng Kinh doanh lệch số Tài chính, đối soát, tìm 4 lỗi, sửa mô hình và measure, điền KPI dictionary, 3 câu hỏi cho Sales Director, 1 câu insight. Đây là phần **C3 Mô hình** và **C4 KPI** của bài capstone. Nộp trước Buổi 6 tại hocvien.mcivietnam.com · đọc handout B5 · tick "Checklist trước Buổi 6".
- **Ảnh:** [VẼ] viên gạch B5 (đã xong) → viên gạch B6 (sáng lên).

**Lời giảng:**
> Hôm nay các bạn đã nối sáu bảng thành một mô hình ngôi sao, viết bộ KPI một lần, và chứng minh nó khớp từng đồng với Buổi 2. Năng lực D2, đối soát, kiểm giả định, giờ ở mức Competent.
>
> Bài về nhà là chính câu hỏi mở buổi, ở ngoài đời: phòng Kinh doanh gửi một báo cáo, số lệch với Tài chính. Các bạn đối soát, tìm ra **bốn** lý do khiến báo cáo đó sai, sửa mô hình và measure, rồi viết KPI dictionary. Đây cũng là hai tiêu chí C3 Mô hình và C4 KPI của bài capstone Buổi 7, cộng lại 35% số điểm, làm tốt bài này là các bạn đã tập trước phần đó.
>
> Số đã đúng. Nhưng một bảng Pivot dày đặc phần trăm không phải thứ giám đốc muốn đọc. Buổi 6 mình biến các measure hôm nay thành một dashboard một trang, mỗi con số trên đó vẫn đến từ đúng các measure này. Cảm ơn cả lớp, hẹn gặp lại!

---

## Phụ lục kịch bản

### Kiểm tra thời lượng

| Phần | Slide | Phút | Timetable B5 |
|---|---|---|---|
| 0 Mở đầu, recap | 1–3 | 5 | 5 |
| 1 Data Model, Data / Diagram view, normalization, fact vs lookup, keys, relationship vs merge | 4–8 | 20 | 20 |
| 2 Relationship, cardinality, filter direction, fact thứ hai Target | 9–12 | 15 | 15 |
| Nghỉ | 13 | 10 | 10 |
| 3 Calculated column vs measure, implicit vs explicit | 14–17 | 15 | 15 |
| 4 Filter context | 18–20 | 15 | 15 |
| 5 CALCULATE, ALL, % of Total, Revenue vs Target %, Variance | 21–25 | 20 | 20 |
| 6 Time intelligence YTD, YoY | 26–28 | 15 | 15 |
| 7 KPI dictionary | 29–30 | 10 | 10 |
| 8 Thực hành: 6 KPI, đối soát với SUMIFS B2, dựng lại báo cáo B3 | 31–32 | 20 | 20 |
| 9 Quiz, insight, preview | 33–39 | 5 | 5 |
| **Tổng** | **39** | **150** | **150** |

Phần 9 cộng từng slide đúng 5 phút (slide 33: 0,2'; 4 câu × khoảng 35 giây ≈ 2,3'; slide 38: 1,7'; slide 39: 0,8'), không có dự phòng, quiz giữ 4 câu thay vì 7 như B4.

**Chỗ dễ lố giờ:** Slide 12 (Pivot đầu tiên từ Data Model, thêm 2 relationship), Slide 20 (hai Pivot + ẩn 9 cột), Slide 27 (Mark as Date Table + 3 measure + Pivot mới). Lớp chậm: bỏ phần Category ở Slide 25, và ở Slide 20 chỉ ẩn hai cột `Sales[Nam]`, `Sales[Thang]`. Demo GM_Dong ở Slide 16 là bắt buộc (Pivot dựng ở bước này trên sheet `PV_Loc` dùng tiếp ở Slide 17, 19, 22, 23; demo cũng là bằng chứng cho con số 33,77% của slide mở buổi), không cắt.

### Lỗi thường gặp → cách xử lý (Course Plan, Hướng dẫn giảng dạy B5)

| Lỗi | Dấu hiệu | Xử lý | Slide |
|---|---|---|---|
| Không có relationship / lọc qua bảng không liên quan | Cùng một giá trị lặp ở mọi dòng; thanh vàng "Relationships between tables may be needed" | Kiểm tra Diagram View; lọc bằng lookup chung | 12, 20, 36 |
| Khóa trùng ở lookup | Không tạo được quan hệ (báo cả hai cột đều có giá trị trùng) | Remove Duplicates trong Power Query (B4) | 7, 10 |
| Kiểu dữ liệu hai đầu khác nhau | Không tạo được quan hệ Order_Date – Date | Đổi kiểu Date trong Power Query (Using Locale, B4) | 10 |
| Đặt cột ngày / tháng của fact vào hàng, cột | Target lặp lại; YTD / YoY sai hoặc trống | Luôn dùng cột của Calendar; Hide from Client Tools các cột của fact | 20, 27 |
| Implicit measure (kéo thả Sum) | Không dùng lại được trong measure khác / CUBEVALUE | Viết explicit measure | 17 |
| KPI tỷ lệ bằng calculated column | Average ra 33,77%, Sum ra hàng trăm nghìn % | Measure dạng tổng ÷ tổng; DIVIDE | 16, 34 |
| Không chặn ô chưa có dữ liệu | Variance tháng 5 = −288.000.000; YoY tháng 5 = −100% | `IF(ISBLANK([Total Revenue]), BLANK(), …)` | 24, 27 |
| Đếm đơn bằng COUNTROWS | Quý 1 ra 6.187 thay vì 3.144 | DISTINCTCOUNT(Sales[Order_ID]) | 31 |
| Cộng các dòng Active Customers | 1.456 thay vì 1.208 | Đọc ô tổng của measure, không cộng tay | 31, 37 |
| Lỡ Refresh All / sửa DataFolder về thư mục B4 | Báo lỗi đường dẫn, hoặc Sales còn 28.328 dòng | Không Refresh trong buổi; dữ liệu đã có sẵn | 3 |

### Danh sách ảnh cần chụp [CHỤP] (Excel 365 Windows, zoom 100%)

**Đã chụp đủ 19/19 ảnh deck và 5 ảnh handout (03/10/2026), đã chèn bằng `build_slides_b5.py` và `build_docs_b5.py`.** Ảnh lưu ở `Materials/B5/Images/` đúng tên ở cột Tên file (danh sách đầy đủ, gồm ảnh handout: `Materials/B5/Images/_DANH_SACH_ANH.md`). Chụp trên `Demo/B5_Demo_Model.xlsx` (S05a/b, S10a/b, S12a/b) và `Demo/B5_Demo_DAX.xlsx` (S15 trở đi) khi đã dựng, làm đúng các bước của lời giảng tới slide tương ứng; hoặc chụp trên `Practice/B5_MCI_Mart_Answer.xlsx` (các ảnh trạng thái cuối: S10b, S12b, S17, S19, S22–S28, S30). Ảnh S05b và S12a phải chụp **trước** khi tạo relationship. Chụp xong đóng file **không lưu**.

| Tên file | Nội dung cần chụp | File · sheet · vùng | Slide |
|---|---|---|---|
| S05a_data_view.png | Cửa sổ Power Pivot, Data View, tab Sales; thấy 6 tab bảng ở dưới (Product, Store, Customer, Target, Calendar, Sales) và dòng Record 1 of 30,247 | `Demo/B5_Demo_Model_HocVien.xlsx` → Data → Manage Data Model | 5 |
| S05b_diagram_roi.png | Diagram View: 6 hộp Sales, Product, Store, Customer, Target, Calendar, **chưa có đường nối** | Như trên, Home → Diagram View (trước Slide 10) | 5 |
| S10a_create_relationship.png | Hộp Create Relationship: Sales / Order_Date ↔ Calendar / Date | Power Pivot → Design → Create Relationship | 10 |
| S10b_diagram_sales.png | Diagram View: Sales ở giữa, 4 đường nối tới Product, Store, Customer, Calendar; ký hiệu 1 và * trên đường nối | Sau khi tạo 4 relationship của Sales | 10 |
| S12a_target_lap.png | Pivot Rows = Category, Values = Sum of Target: 6 dòng đều 3.665.000.000; khung PivotTable Fields có thanh vàng "Relationships between tables may be needed" | Sheet `PV_Grain`, **trước** khi nối Target | 12 |
| S12b_diagram_hai_fact.png | Diagram View đủ 6 bảng, 6 relationship; Sales và Target cùng nối vào Store và Calendar | Sau Slide 12 | 12 |
| S15_cot_cogs.png | Data View bảng Sales, cột COGS ở cuối, thanh công thức `=Sales[Quantity]*RELATED(Product[Unit_Cost])` | Power Pivot → Data View → Sales | 15 |
| S16_gm_dong.png | Pivot lọc Nam = 2025, Quy = 1: Average of GM_Dong 33,77% và Sum of GM_Dong | Sheet `PV_Loc` (cột GM_Dong thêm tạm) | 16 |
| S17_measure_dialog.png | Hộp Measure: Table name Sales, Measure name Gross Margin %, Formula `=DIVIDE([Total Revenue]-[Total COGS],[Total Revenue])`, Category Number, Format Percentage (ảnh chụp lúc Decimal places còn 2) | Power Pivot → Measures → New Measure | 17 |
| S19_pivot_mien_thang.png | Pivot Rows Region, Columns Thang (1–3), Filters Nam = 2025 và Quy = 1, Values Total Revenue; ô Bắc × 3 = 90.210.000 đang chọn | Sheet `PV_Loc` | 19 |
| S20_lookup_vs_fact.png | Hai Pivot cạnh nhau: trái Rows Calendar[Thang] (T5–T12 chỉ có target); phải Rows Sales[Thang], chỉ còn cột Sum of Target = 3.665.000.000 ở mọi dòng | Sheet `PV_Lookup_vs_Fact`, trước khi ẩn cột | 20 |
| S22_calculate.png | Pivot Rows Region (Quý 1): Total Revenue và Revenue Mien Nam = 427.663.000 ở cả 3 dòng | Sheet `PV_Loc` | 22 |
| S23_pct_of_total.png | Pivot Rows Region (Quý 1): Total Revenue, Revenue All Stores = 941.454.000 mọi dòng, % Revenue Share Bắc 34,4% · Nam 45,4% · Trung 20,2% | Sheet `PV_Loc` | 23 |
| S24_target_thang.png | Pivot Filters Nam = 2025, Rows Calendar[Thang]: Total Revenue, Total Target, Revenue vs Target %, Variance; T1 120,6%, T5–T12 chỉ có target | Sheet `PV_Target` (Variance đã chặn ô trống) | 24 |
| S25_store_thang.png | Pivot Rows Store_Code, Columns Thang 1–4, Values Revenue vs Target %; ô ST07 × 3 = 44,9% đang chọn | Sheet `PV_Target` theo lời giảng (ảnh hiện có: bản sao `Practice/B5_MCI_Mart_Answer.xlsx`, sheet `PV_Target_CuaHang`) | 25 |
| S27a_mark_date_table.png | Hộp Mark as Date Table, Date column = Date | Power Pivot, tab Calendar → Design → Mark as Date Table | 27 |
| S27b_yoy_ytd.png | Pivot Filters Nam = 2025, Rows Thang 1–4: Total Revenue, Revenue LY, Revenue YTD, Revenue YoY %; T4 YTD 1.238.419.000 | Sheet `PV_YoY` | 27 |
| S28_yoy_cua_hang.png | Pivot Rows Store_Code, Filters Nam = 2025 và Thang 1–4: Total Revenue, Revenue LY, Revenue YoY %; dòng ST12 hai ô trống | Sheet `PV_YoY` | 28 |
| S30_gm_theo_nam.png | Pivot Rows Calendar[Nam], không lọc, Values Gross Margin %: 2024 31,2% · 2025 34,1% | Sheet `PV_YoY` theo lời giảng (ảnh hiện có: bản sao `Practice/B5_MCI_Mart_Answer.xlsx`, sheet `PV_GM_Nam`) | 30 |

### File Demo (theo mẫu B3): đã dựng 02/10/2026

**Vì sao.** B3 có `Demo/` (3 file, mỗi file có `_HocVien` và sheet `Huong_dan`; `B3_Demo_Pivot` dùng cùng dữ liệu với file khởi đầu nhưng gọn hơn). B4 và bản nháp đầu của B5 dùng ngoại lệ "dùng lại file Answer" trong `CLAUDE.md`, nên giảng viên phải demo trong file có hơn 20 sheet của B2–B4 và học viên không có `Huong_dan` theo slide. B5 quay về mô hình B3. Practice giữ như B3: học viên làm Slide 31–32 trên file của mình (khởi đầu = đáp án buổi trước), đáp án ở `Practice/`.

| File | Trạng thái khi mở | Slide | Kết quả mong đợi (kiểm bằng Excel sau khi dựng) |
|---|---|---|---|
| `Demo/B5_Demo_Model.xlsx` | Data Model 6 bảng (Sales 30.247 · Product 40 · Store 12 · Customer 1.500 · Target 144 · Calendar 731), **0 relationship, 0 measure**. Sheet `Sales` (tblSales), `Product`, `Store`, `KPI_SUMIFS` như `B3_Demo_Pivot`; `Huong_dan` ở cuối | 5, 8, 10, 11, 12 | Diagram View: 6 hộp rời. Sau 4 relationship của Sales: hình ngôi sao. `PV_Grain` (Category × Sum of Target, trước khi nối Target): 6 dòng đều 3.665.000.000. Sau khi nối Target: Bắc 1.227.000.000 · Nam 1.685.000.000 · Trung 753.000.000 (tổng 3.665.000.000) |
| `Demo/B5_Demo_DAX.xlsx` | Cùng 6 bảng, **6 relationship** (Sales và Target cùng nối Store và Calendar), **0 measure**, chưa Mark as Date Table, chưa cột COGS; `KPI_SUMIFS` giữ nguyên địa chỉ ô (D6, E9, E16, E30) để lời giảng không đổi; sheet trống `PV_Loc` (ô A1 ghi "Đặt PivotTable tại ô A3") cho Pivot của Slide 16 | 15–17, 19–20, 22–25, 27–28, 30 | COGS = Quantity × RELATED(Unit_Cost). `GM_Dong` trung bình Quý 1/2025 = 33,77% (KPI_SUMIFS!E30 = 34,18%). Total Revenue Bắc × 3 = 90.210.000; Quý 1 = 941.454.000. Revenue Mien Nam = 427.663.000 ở cả 3 dòng. % Revenue Share 34,4% · 45,4% · 20,2% (Bắc · Nam · Trung). Revenue vs Target % T1 120,6% · T2 82,5%; ST07 × T3 44,9%. YoY T1 +42,8% · T2 −0,9%; Revenue YTD T4 1.238.419.000; ST12 trống T1–T4. Gross Margin % 2024 31,2% · 2025 34,1% |

- **Bản `_HocVien`:** bỏ câu dành cho giảng viên, tiêu đề "DEMO" → "VÍ DỤ", `Huong_dan` đứng đầu; bản giảng viên mở ở sheet làm việc, `Huong_dan` ở cuối (như B3).
- **`Huong_dan`:** cột Slide · Thao tác · Kết quả mong đợi, soạn từ các dòng *(→ …)* của kịch bản, số liệu từ `build_b5_meta.json`.
- **Dựng (đã chạy 02/10/2026, ~2 phút, `--check` đạt):** `_generator/build_b5_demos.py`, **CHỈ WINDOWS** (Data Model và relationship chỉ tạo được bằng Excel qua COM; dùng lại `datamodel_com.py` của `build_b5.py`). Quy trình: chép đáp án B4 → bỏ sheet Pivot, `Kiem_tra*`, `ST07_T3_ChiTiet` → giữ Data Model nguyên 6 bảng (không Refresh) → với `Demo_DAX` thêm 6 relationship → ghi `Huong_dan` → lưu bằng Excel. `--check` mở file, đối chiếu số dòng 6 bảng, số relationship (0 và 6), số measure (0), rồi trên bản sao tạm thêm đúng các measure của lời giảng và so các con số ở cột "Kết quả mong đợi". Kết quả `--check`: Data Model đủ 6 bảng (Sales 30.247, Product 40, Store 12, Customer 1.500, Target 144, Calendar 731); Demo_Model 0 relationship, Demo_DAX 6; cả hai 0 measure, 0 Pivot. Sheet: `Sales`, `Product`, `Store`, `KPI_SUMIFS`, `Sales_PQ`, `Huong_dan` (đã giữ `Sales_PQ`). Các bước cần giao diện (cột COGS, GM_Dong, Mark as Date Table, Hide from Client Tools) làm trên lớp theo `Huong_dan`.
- **Việc đi kèm (02/10/2026):** đã đổi chân slide và dòng *(→ mở …)* ở Slide 5 và 15 sang `Demo/…`; bảng "File dùng trong buổi", `_DANH_SACH_ANH.md`, `2_Dataset/README.md` đã cập nhật. **Còn:** chạy lại `build_slides_b5.py` và `build_student_slides.py B5` (đóng PowerPoint trước) để speaker notes và chân slide trong `.pptx` khớp kịch bản.

### Bài tập về nhà (bắt buộc): tập dượt capstone C3, C4

Ghi vào `Course_Plan.md` (B5, mục Chuẩn bị; Nguyên tắc 11) ngày 01/10/2026. Số liệu trong mục này lấy từ `build_b5_meta.json` (khóa `hw`).

**Tình huống.** Phòng Kinh doanh gửi Sales Director báo cáo tháng 1–4/2025. Tài chính báo số khác. Học viên là BI analyst được giao kiểm tra: số nào đúng, vì sao lệch, và sửa để từ nay hai phòng dùng chung một định nghĩa. Đây là câu hỏi mở của B5, đặt vào một tình huống công việc.

**Đề `Homework/B5_BTVN_MCI_Mart_KPI_T1_T4_2025.xlsx`:**
- Data Model 6 bảng (cùng dữ liệu của lớp, tới 04/2025) do "phòng Kinh doanh" dựng: Sales nối đủ 4 lookup, **Target chỉ nối Store, chưa nối Calendar**; 6 measure `KD …` do phòng Kinh doanh viết: `KD Doanh thu`, `KD Target`, `KD % dat`, `KD So don` (= COUNTROWS), `KD AOV`, `KD GM %` (= AVERAGEX, trung bình biên từng dòng). Excel không cho script tạo calculated column, nên cột `GM_Dong` trong bản nháp được thay bằng measure AVERAGEX cho cùng kết quả.
- Sheet `BaoCao_KinhDoanh`: Pivot tháng 1–4 lấy tháng từ **`Sales[Thang]`**, cột Doanh thu, Target, % đạt, Số đơn (**`KD So don`** = đếm dòng), AOV (doanh thu ÷ số đơn đó), GM% (**`KD GM %`** = trung bình biên từng dòng).
- Sheet `So_lieu_Tai_chinh`: số chính thức theo tháng, doanh thu, số đơn, GM%, target.
- Sheet `De_bai`, `Nhat_ky_doi_soat` (mẫu: chỉ tiêu · số báo cáo cũ · số Tài chính · chênh lệch · nguyên nhân · cách sửa), `KPI_Dictionary` (mẫu trống 10 trường của KPI Definition Card + cột Measure DAX), `Brief_Insight`.

**Bốn lỗi cài sẵn** (mỗi lỗi lộ ra khi đối soát với `So_lieu_Tai_chinh`):

| # | Lỗi | Báo cáo cũ (T1 · T2 · T3 · T4) | Tài chính | Tiêu chí |
|---|---|---|---|---|
| 1 | Target chưa nối Calendar | target tháng = 3.665.000.000 (cả năm) → % đạt 11,6% · 6,7% · 7,4% · 8,1% | 120,6% · 82,5% · 94,6% · 103,1% | C3 |
| 2 | Báo cáo lấy tháng từ `Sales[Thang]` (sửa lỗi 1 mà chưa sửa lỗi này thì target vẫn lặp) | như trên | như trên | C3 |
| 3 | GM% = trung bình biên từng dòng (measure AVERAGEX) | 34,62% · 33,25% · 32,97% · 33,54% | 35,16% · 33,70% · 33,07% · 33,80% | C4 |
| 4 | Số đơn = đếm dòng | 2.725 · 1.595 · 1.867 · 1.919 đơn → AOV 156.275 · 154.229 · 144.407 · 154.750 | 1.398 · 823 · 923 · 966 đơn → AOV 304.614 · 298.902 · 292.100 · 307.417 | C4 |

Lỗi 3 là lỗi khó thấy nhất (lệch 0,1–0,5 điểm), chỉ ai đối soát từng ô mới bắt được. Đúng ý D2: số trông hợp lý chưa chắc đúng.

**Việc phải làm** (De_bai; sheet có 5 cột: # · Việc cần làm · Yêu cầu · Kết quả kỳ vọng · Tự kiểm tra nhanh, theo mẫu đề B4):
1. **Đối soát** báo cáo cũ với `So_lieu_Tai_chinh`; ghi `Nhat_ky_doi_soat` đủ 4 lỗi gốc, mỗi lỗi một dòng (chỉ tiêu bị kéo theo, như AOV theo số đơn, ghi chung dòng): chỉ tiêu, số hai bên, nguyên nhân, cách sửa. *(D2)*
2. **Sửa mô hình** thành hình ngôi sao: nối Target vào Calendar, ẩn cột kỹ thuật của hai fact, Mark as Date Table; dán ảnh Diagram View vào `Brief_Insight`. *(C3)*
3. **Viết lại KPI bằng explicit measure** theo tầng (measure cơ sở → measure dẫn xuất): 6 KPI của lớp, Total Target, Revenue vs Target %, Variance; thêm **like-for-like**: `Revenue LFL := CALCULATE([Total Revenue], Store[Open_Date] < DATE(2024, 1, 1))`, `Revenue LFL LY`, `YoY LFL %`; xóa 6 measure `KD …`. *(C4)*
4. **Dựng lại báo cáo:** `BaoCao_KinhDoanh` từ measure của bạn (tháng lấy từ `Calendar[Thang]`, Filters Nam = 2025), thêm Pivot cửa hàng × tháng cho Revenue vs Target %; sheet mới `BaoCao_LFL` (YoY cả chuỗi và YoY like-for-like theo tháng); đối soát lại với `So_lieu_Tai_chinh`, chênh lệch bằng 0. *(C3, C4)*
5. **KPI_Dictionary:** một dòng cho mỗi measure ở việc 3, gồm DAX, grain, dimensions, owner, target, ngưỡng và **giới hạn** (Revenue vs Target % không theo ngành; GM% dùng một bảng giá vốn; YoY cả chuỗi gồm ST12 chưa có kỳ gốc). *(C4)*
6. **Brief_Insight:** 3 câu hỏi đo lường được cho Sales Director (chỉ số, chiều, thời gian, mốc so sánh; đề có một câu ví dụ), mỗi câu trả lời bằng một Pivot từ measure; 1 câu insight (con số có so sánh, nguyên nhân khả dĩ, hàm ý). *(tập trước C1, C6)*

**Số tự kiểm tra** (in trong De_bai): tháng 4, doanh thu 296.965.000, 966 đơn, AOV 307.417, GM 33,8%, đạt 103,1% target; YoY T1–T4 cả chuỗi **+21,0%**, like-for-like **+13,3%** (T1 +32,5% · T2 −8,7% · T3 +8,1% · T4 +17,2%); cửa hàng dưới 80% target, T2: ST01, ST05, ST08, ST09, ST12 · T3: ST02, ST06, ST07, ST11, ST12 · T4: ST04, ST12.

**Đáp án `_Answer.xlsx`:** mô hình đã sửa, measure, `BaoCao_KinhDoanh` dựng lại, `Nhat_ky_doi_soat` mẫu, `KPI_Dictionary` mẫu, `Kiem_tra` (số liệu từ `build_b5_meta.json`), `Dap_an_loi`, `Insight_mau`. Chấm như các buổi trước: theo `Kiem_tra`, bằng agent `homework-grader`.

**Dựng (02/10/2026):** `_generator/build_b5_homework.py` (chỉ Windows, Excel qua COM, dùng `datamodel_com.py`). Đề dựng từ đáp án B4, đáp án dựng từ `Practice/B5_MCI_Mart_Answer.xlsx` (thêm `Revenue LFL`, `Revenue LFL LY`, `YoY LFL %`; `KPI_Dictionary` giữ từ file Practice, thêm 3 dòng LFL). Kết quả: `BaoCao_KinhDoanh` của đề ra đúng số báo cáo cũ ở bảng trên (T1 % đạt 11,6%, GM 34,62%, 2.725 đơn, AOV 156.275); `Kiem_tra` của đáp án 24/24 ✔ (4 tháng × doanh thu, % đạt target, số đơn, AOV, GM%, YoY like-for-like). Chưa có `grade_b5.py`: agent `homework-grader` viết khi chấm, đọc giá trị Pivot `BaoCao_KinhDoanh` / `BaoCao_LFL` của học viên đã lưu trong file (không đọc được measure của học viên bằng openpyxl), cùng `Nhat_ky_doi_soat`, `KPI_Dictionary` (có cột DAX) và `Brief_Insight`.

### Handout: dàn ý (Course Plan: mục Handout B5; đã dựng 03/10/2026)

1. Ẩn cột khỏi Client Tools và bỏ cột thừa, giảm granularity (performance & best practices) [PQ PDF p56, 108–109]
2. Hierarchy trên Calendar: Nam → Quy → Thang → Date [PQ PDF p57]
3. Lookup-to-lookup: `Customer[Home_Store_Code]` → Store; quan hệ active / inactive, USERELATIONSHIP [PQ PDF p47, 49]
4. Bad cardinality: thử nối thẳng `Sales[Store_Code]` với `Target[Store_Code]` [PQ PDF p51–52]
5. PRO TIP iterator: `SUMX(Sales, Sales[Quantity] * RELATED(Product[Unit_Cost]))` thay cột COGS [PQ PDF p103]
6. PRO TIP `SWITCH(TRUE(), …)`: nhãn trạng thái target [PQ PDF p88–89]
7. DAX text functions: FORMAT, LEFT, & [PQ PDF p90–91]
8. Disconnected slicer: bảng ngưỡng target 80% / 90% / 100% [PQ PDF p96–97]
9. Bài tập về nhà; checklist trước Buổi 6

Ảnh handout H01–H05 ở `Materials/B5/Images/_DANH_SACH_ANH.md` (bảng Handout); `build_docs_b5.py` tự chèn.

### Việc còn lại (theo thứ tự)

1. **Duyệt kịch bản và kế hoạch.**
2. ~~`build_b5.py` + `datamodel_com.py`~~, đã chạy 02/10/2026, `Kiem_tra_B5` 24/24 ✔. **Còn:** làm tay Mark as Date Table, Hide from Client Tools (Sales: Nam, Thang, Order_Date, Store_Code, Product_Code, Customer_ID; Target: Store_Code, Thang, Ngay_Dau_Thang), lưu, đóng, `build_b5.py --check`.
3. ~~**Demo:** `build_b5_demos.py`~~, đã dựng 4 file, `--check` đạt (02/10/2026).
4. ~~**Bài tập về nhà:** `build_b5_homework.py`~~, đã dựng đề + đáp án, `Kiem_tra` 24/24 ✔ (02/10/2026). `grade_b5.py`: agent `homework-grader` viết khi chấm.
5. ~~**Chạy lại `build_slides_b5.py` và `build_student_slides.py B5`**~~, đã chạy lại 03/10/2026 sau đợt rà soát; speaker notes khớp kịch bản.
6. ~~**Handout:** `build_docs_b5.py` (Mac) → PDF bằng Word~~, đã dựng 03/10/2026 (8 trang, 5 ảnh).
7. ~~`session-builder` `khung` → `screenshot-capturer` (19 ảnh deck + H01–H05) → `session-builder` `chot`~~, xong 03/10/2026, không còn [CHỤP]. **Còn:** rà lại infographic khi kịch bản được duyệt, rồi chạy lại `consistency-checker`.
8. **Sau đợt sửa 13/10/2026:** dựng lại `build_b5_demos.py` trên Windows (chữ `Huong_dan` đổi; sheet `PV_Loc` giữ nguyên nội dung), rồi `session-builder` `chot`: deck → `build_student_slides.py B5` → handout → PDF → infographic; chạy lại `consistency-checker`. Kiểm bằng Excel: Grand Total của `PV_YoY` sau khi bỏ tick ST12 ở nút lọc Row Labels ra +13,3%.

### Còn cần làm trước khi dạy

- Chạy thử toàn bộ buổi trên máy trình chiếu, từ file khởi đầu tới Slide 32; ghi lại thời gian thật của Slide 12, 20, 27.
- Kiểm tra trên Excel 365 Windows: thanh vàng "Relationships between tables may be needed" có hiện ở Slide 12 không; tab All của PivotTable Fields có hiện `tblSales` không; DAX `Store[Region] = "Nam"` chạy đúng.

### Ghi chú so với Course Plan

- **Câu chuyện #3 (ST12):** Course Plan ghi "YoY trống T1–T5"; dữ liệu khóa học chỉ tới 04/2025 nên kịch bản nói T1–T4.
- **Recap 5 phút, quiz 5 phút** đúng timetable; quiz rút còn 4 câu (2 câu tình huống) để vừa 5 phút.
- **"6 KPI khớp bảng SUMIFS của B2"** (Kiểm tra cuối buổi): KPI_SUMIFS có 4 KPI (doanh thu, số đơn, AOV, GM%), đối soát đủ 64 ô; Revenue YoY % và Active Customers không có ở B2, nên kiểm bằng bảng số liệu trong kịch bản (tính từ dữ liệu).
- **"Dựng lại báo cáo B3 bằng Power Pivot"** (Thực hành): làm ở Slide 32 bằng Pivot Miền × Tháng (bản Data Model của `PV_MienThang`), GM% bằng measure thay cho Calculated Field.
- **Câu chuyện #5:** Course Plan ghi 2025 "khoảng 33–35%", 2024 "khoảng 31%"; số tính được là 34,1% và 31,2%.
- **Time intelligence** dùng `DATEADD(…, -1, YEAR)` theo PQ PDF p107; SAMEPERIODLASTYEAR để ở PRO TIP.
