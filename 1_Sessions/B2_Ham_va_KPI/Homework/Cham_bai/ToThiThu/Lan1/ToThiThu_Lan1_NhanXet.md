**Điểm: 8/10**

# Nhận xét bài tập về nhà Buổi 2 — Tô Thị Thu

| # | Việc | Kết quả | Điểm |
|---|---|---|---|
| 1 | Tra cứu XLOOKUP (Region, Category, Unit_Cost) | Đạt | 2/2 |
| 2 | Cột tính toán (COGS, Gross_Margin, Phan_khuc) | Một phần | 1/2 |
| 3 | Bảng KPI 1 — Miền × Ngành hàng | Đạt | 1/1 |
| 4 | Bảng KPI 2 — theo Miền | Một phần | 0,5/1 |
| 5 | KPI Definition Card (AOV) | Một phần | 0,5/1 |
| 6 | Insight | Insight 3/3 thành phần | 3/3 |
| | **Tổng** | | **8/10** |

Thang điểm: việc 1 và 2 mỗi việc 2 điểm, việc 3, 4, 5 mỗi việc 1 điểm, việc 6 (insight) 3 điểm, mỗi thành phần đạt (con số, nguyên nhân, hàm ý) được 1 điểm; việc 1–5 Đạt được trọn điểm, Một phần được một nửa, chưa làm được 0.

## Tài liệu để xem lại

Slide nằm trong file `Slide_B2_HocVien.pptx` (thư mục Buổi 2). Các file Excel để làm lại thao tác tính từ thư mục Buổi 2; slide của Buổi 1 nằm trong `Slide_B1_HocVien.pptx` (thư mục Buổi 1).

| Slide | Nội dung | Làm lại trên file |
|---|---|---|
| Buổi 2, slide 4 | Dimension / Measure / Grain: một đơn có nhiều dòng sản phẩm | `../B1_Nen_tang_Excel_Du_lieu/Practice/B1_MCI_Mart_Q1_2025_Answer.xlsx`, sheet Sales, đơn HD2501-00002 (A1:K7) |
| Buổi 2, slide 5 | KPI Definition Card (10 trường) | `Handout/B2_KPI_Definition_Card.docx` |
| Buổi 2, slide 7 | IF, IFS: nhãn trả về do bạn tự đặt | `Demo/B2_Demo_Logic.xlsx`, sheet Demo_IF |
| Buổi 2, slide 16 | SUMIFS, COUNTIFS; đếm đơn bằng `Line_No = 1` | `Demo/B2_Demo_SUMIFS.xlsx`, sheet Demo_SUMIFS |
| Buổi 2, slide 17 | Bảng KPI 2 chiều; KPI tỷ lệ tính **tổng ÷ tổng** | `Demo/B2_Demo_SUMIFS.xlsx`, sheet Demo_SUMIFS (A11:E15 và khối Gross Margin % dòng 20–28) |
| Buổi 2, slide 30–31 | Thực hành 2 (bảng KPI & KPI Card) và chữa bài | `Practice/B2_MCI_Mart_Q1_2025_Answer.xlsx`, sheets KPI_Card và Kiem_tra |
| Buổi 2, slide 40 | Tổng kết + câu insight | Sheet `insight` trong file bạn nộp |
| Buổi 1, slide 26 | Structured Reference `tblSales_T4[Revenue]` | `../B1_Nen_tang_Excel_Du_lieu/Demo/B1_Demo_Table_Validation.xlsx`, sheet Sales |

## Đối soát nhanh với đáp án

Đáp án nằm ở `Homework/B2_BTVN_MCI_Mart_T4_2025_Answer.xlsx`, sheet `Kiem_tra`.

| Chỉ tiêu | Đáp án | File của bạn | Khớp? |
|---|---|---|---|
| Số dòng dữ liệu | 1.919 | 1.919 | ✔ |
| Ô `#N/A` (Region, Category, Unit_Cost) | 0 | 0 | ✔ |
| Tổng doanh thu | 296.965.000 | 296.965.000 | ✔ |
| Tổng bảng KPI 1 (Miền × Ngành) | 296.965.000 | 296.965.000 | ✔ |
| Tổng COGS | 196.595.000 | 196.595.000 | ✔ |
| Gross Margin % dòng Tổng | 33,8% | 33,8% (L5) | ✔ |
| Số đơn từng miền (Bắc / Trung / Nam) | 317 / 202 / 447 | 317 / 202 / 447 | ✔ |
| **Tổng số đơn** | **966** | Ô I5 **trống** | ✘ |
| **AOV dòng Tổng** | **307.417** | Ô J5 **trống** | ✘ |
| Nhãn Phan_khuc | Cao / Trung / Thấp | Cao / **trung bình** / **thấp** | ✘ |

## Làm tốt

- **Phần tra cứu sạch.** Region, Category, Unit_Cost dùng XLOOKUP, không còn ô `#N/A`, **không dùng IFERROR để che** lỗi, Product_Code không còn dấu cách thừa (4 mã có dấu cách ở cuối đã được xử lý). Số dòng 1.919 và tổng COGS 196.595.000 khớp số tự kiểm tra.
- **Bảng KPI 1 chỉ một công thức SUMIFS** kéo ra cả khối 3 × 6 với `$` hỗn hợp (`KPI!$A2`, `KPI!B$1`): đúng điều slide 17 dạy. Tổng doanh thu 296.965.000.
- **Gross Margin % dòng Tổng (L5) = tổng ÷ tổng = 33,8%**, đúng; bạn không dùng AVERAGE các miền (bẫy 31,2% ở slide 17).
- **Số đơn và AOV từng miền đúng** (317 + 202 + 447; COUNTIFS với `Line_No = 1`).
- **Câu insight đủ ba thành phần** (xem cuối).

![Hình 1 · Bảng KPI 1 của bạn](Images/NX01_kpi1_ban.png)
> **[CHỤP] NX01_kpi1_ban.png** · File: `Homework/Bai_nop/b2_btvn_to-thuXRBUKA.xlsx` · Sheet: `KPI` · Vùng: A1:H5 · Cần thấy: khối Miền × Ngành 3 × 6 + dòng/cột Tổng, H5 = 296.965.000; chọn B2 để thanh công thức hiện `SUMIFS(... $A2 ... B$1)`.

## Điểm cần lưu ý

### 1. Dòng Tổng của bảng KPI 2 thiếu Số đơn và AOV (Việc 4: 0,5/1)

**Bạn đã làm:** sheet `KPI`, ô **I5 (Số đơn, dòng Tổng)** và **J5 (AOV, dòng Tổng)** đang **trống**. Số đơn và AOV từng miền đã đúng (317 + 202 + 447).
**Đáp án:** dòng Tổng có **966 đơn** và AOV **307.417**.
**Hậu quả:** người đọc báo cáo không có con số "cả tháng" cho hai KPI quan trọng nhất của bảng; `Kiem_tra` của đáp án cũng đối chiếu đúng hai ô này (966 và 307.417).

**Cách sửa:**
1. I5: `=SUM(I2:I4)` → ra **966**.
2. J5: `=H5/I5` → ra **307.417** (**tổng ÷ tổng**, không dùng `AVERAGE(J2:J4)`).
3. Kiểm tra chéo: `AVERAGE` của ba AOV miền = (299.804 + 296.366 + 317.810) ÷ 3 ≈ 304.660, **khác** 307.417. Đây chính là bẫy "trung bình của trung bình": miền Nam có 447 đơn, nhiều gấp đôi miền Trung (202 đơn) nhưng bị tính "một phiếu" ngang nhau.

**Học lại:** **slide 16** (COUNTIFS + `Line_No = 1`) và **slide 17** (tổng ÷ tổng, `Slide_B2_HocVien.pptx`), làm trên `Demo/B2_Demo_SUMIFS.xlsx`, sheet Demo_SUMIFS (khối Gross Margin % dòng 20–28: AVERAGE ra 31,2% sai, tổng ÷ tổng ra 34,2% đúng).

![Hình 2 · Dòng Tổng bảng KPI 2 trong bài của bạn](Images/NX02a_dong_tong_ban.png)
> **[CHỤP] NX02a_dong_tong_ban.png** · File: `Homework/Bai_nop/b2_btvn_to-thuXRBUKA.xlsx` · Sheet: `KPI` · Vùng: A1:L5 (hoặc H1:L5 nếu quá rộng) · Cần thấy: dòng 5 "Tổng" với I5 và J5 trống; viền đỏ I5 và J5; cột Số đơn (317 / 202 / 447) và AOV từng miền ở trên.

![Hình 2 · Bảng 2 trong file đáp án](Images/NX02b_dong_tong_dap_an.png)
> **[CHỤP] NX02b_dong_tong_dap_an.png** · File: `Homework/B2_BTVN_MCI_Mart_T4_2025_Answer.xlsx` · Sheet: `KPI` · Vùng: A11:E16 · Cần thấy: dòng Tổng có C16 = 966, D16 = 307.417, E16 = 33,8%; chọn D16 để thanh công thức hiện `=B16/C16`.

![Hình 3 · Slide 17: AVERAGE sai, tổng ÷ tổng đúng](Images/NX03_demo_tong_chia_tong.png)
> **[CHỤP] NX03_demo_tong_chia_tong.png** · File: `Demo/B2_Demo_SUMIFS.xlsx` · Sheet: `Demo_SUMIFS` · Vùng: khối Gross Margin % theo ngành (dòng 20–28) · Cần thấy: ô AVERAGE = 31,2% và ô tổng ÷ tổng = 34,2% cạnh nhau.

### 2. Nhãn cột Phan_khuc không đúng nhãn đề (Việc 2: 1/2)

**Bạn đã làm:** sheet `Sales`, cột **P (Phan_khuc)**: ngưỡng đúng, nhưng nhãn bạn đặt là **"trung bình"** (910 dòng) và **"thấp"** (440 dòng, viết thường).
**Đề yêu cầu:** nhãn **"Cao" / "Trung" / "Thấp"** (≥ 100.000 "Cao", ≥ 30.000 "Trung", còn lại "Thấp").
**Hậu quả:** nhãn là **khóa để Pivot và lọc** từ Buổi 3. "trung bình" khác hẳn "Trung" nên bảng đối soát hoặc Pivot dựng sẵn theo đề sẽ không gom được; chữ "thấp" viết thường tuy Excel (Filter, COUNTIFS) coi như "Thấp", nhưng **Power Query ở Buổi 4 phân biệt hoa/thường**, nên "thấp" và "Thấp" sẽ thành hai nhóm khác nhau.

**Cách sửa:** sửa công thức IFS cho đúng nhãn đề:
`=IFS([@Unit_Price]>=100000,"Cao",[@Unit_Price]>=30000,"Trung",TRUE,"Thấp")`.
Nhánh cuối dùng `TRUE, "Thấp"` ("các trường hợp còn lại") thay vì lặp lại một điều kiện: ngắn hơn và không bao giờ bỏ sót dòng nào.
Kiểm tra: lọc cột P chỉ còn đúng 3 giá trị Cao / Trung / Thấp.

**Học lại:** **slide 7** (`Slide_B2_HocVien.pptx`), làm trên `Demo/B2_Demo_Logic.xlsx`, sheet Demo_IF (cột H, IFS; TIP: `TRUE` ở cuối là "các trường hợp còn lại").

![Hình 4 · Cột Phan_khuc của bạn](Images/NX04a_phan_khuc_ban.png)
> **[CHỤP] NX04a_phan_khuc_ban.png** · File: `Homework/Bai_nop/b2_btvn_to-thuXRBUKA.xlsx` · Sheet: `Sales` · Vùng: I1:P8 · Cần thấy: cột P có nhãn "trung bình" / "thấp"; chọn một ô P để thanh công thức hiện công thức IFS của bạn.

![Hình 4 · Bộ lọc cột Phan_khuc của bạn](Images/NX04b_loc_phan_khuc_ban.png)
> **[CHỤP] NX04b_loc_phan_khuc_ban.png** · Cùng file và sheet · Mở: bấm mũi tên Filter ở cột P · Cần thấy: các giá trị lọc "Cao", "thấp", "trung bình".

![Hình 4 · Cột Phan_khuc trong file đáp án](Images/NX04c_phan_khuc_dap_an.png)
> **[CHỤP] NX04c_phan_khuc_dap_an.png** · File: `Homework/B2_BTVN_MCI_Mart_T4_2025_Answer.xlsx` · Sheet: `Sales` · Vùng: I1:P8 · Cần thấy: cột P có nhãn "Cao" / "Trung" / "Thấp"; thanh công thức `=IFS(...,TRUE,"Thấp")`.

### 3. KPI_Card, mục Công thức (B6) và Nguồn dữ liệu (B7) (Việc 5: 0,5/1)

**Bạn đã làm:**
- B6: `=SUM(tblSales[Revenue]) /COUNTIFS(Sales!$K:$K,KPI!A2,Sales!$B:$B,1)`: mẫu số `COUNTIFS(...,KPI!A2,...)` chỉ đếm đơn của **một miền** (dòng 2), trong khi tử số là SUM của **cả bảng**.
- Tên bảng ghi `tblSales`; tên đúng là **`tblSales_T4`**.
- B7: cột "Order_ID_line" **không tồn tại**; cột đúng là `Order_ID` và `Line_No`.

**Đáp án (KPI_Card):** Công thức: *"Tổng doanh thu ÷ Số đơn = SUM(Revenue) ÷ COUNTIFS(Line_No, 1)"*; Nguồn: *"tblSales_T4 (Revenue, Order_ID, Line_No)"*.
**Hậu quả:** nếu ai copy công thức trong card, kết quả = 296.965.000 ÷ 317 ≈ 936.800 (chỉ chia cho đơn miền Bắc), không phải AOV 307.417. Card là "hợp đồng" giữa người làm và người đọc nên công thức phải chạy đúng khi dán.

**Cách sửa:**
- B6: `=SUM(tblSales_T4[Revenue]) / COUNTIFS(tblSales_T4[Line_No],1)` → ra 307.417.
- B7: "Bảng: tblSales_T4. Cột: Revenue, Order_ID, Line_No."

**Học lại:** **slide 5** (viết card trước, công thức sau; mẫu `Handout/B2_KPI_Definition_Card.docx`) và **slide 30** (`Slide_B2_HocVien.pptx`); so với `Practice/B2_MCI_Mart_Q1_2025_Answer.xlsx`, sheet KPI_Card (card mẫu của khóa).

![Hình 5 · KPI_Card của bạn, mục Công thức và Nguồn](Images/NX05a_card_cong_thuc_ban.png)
> **[CHỤP] NX05a_card_cong_thuc_ban.png** · File: `Homework/Bai_nop/b2_btvn_to-thuXRBUKA.xlsx` · Sheet: `KPI_Card` · Vùng: A3:B8 · Cần thấy: B6 (công thức có `KPI!A2`) và B7 (có "Order_ID_line"); viền đỏ hai ô này. Bật Wrap Text, kéo rộng cột B.

![Hình 5 · KPI_Card trong file đáp án](Images/NX05b_card_cong_thuc_dap_an.png)
> **[CHỤP] NX05b_card_cong_thuc_dap_an.png** · File: `Homework/B2_BTVN_MCI_Mart_T4_2025_Answer.xlsx` · Sheet: `KPI_Card` · Vùng: A3:B8 · Cần thấy: B6 và B7 theo đáp án.

### 4. KPI_Card, mục Dimensions (B9) và Target (B11) (Việc 5)

**Bạn đã làm:**
- B9 liệt kê **Category** và **Phan_khuc** là chiều *xem được*.
- B11: Target **304.000** không nêu nguồn.

**Đáp án:** Dimensions: *"Miền, cửa hàng, tháng (không theo ngành hàng: một đơn có nhiều ngành)"*.
**Vì sao Category, Phan_khuc không dùng cho AOV:** Grain của dữ liệu là *một sản phẩm trong một đơn*. Một đơn có nhiều sản phẩm, nhiều ngành, nhiều phân khúc. Chia AOV theo ngành hay phân khúc thì **một đơn bị đếm lặp** ở nhiều ngành, và tổng các nhóm vượt tổng thật. Chiều xem được phải là thuộc tính của **cả đơn**: Miền, cửa hàng, tháng.
**Target:** ghi rõ nguồn: *"giả định của người viết"* hoặc *"lấy từ kế hoạch quý/năm"*; đáp án mẫu dùng 300.000 (Warning 90% = 270.000, Critical 80% = 240.000). Bạn nêu Warning 90% / Critical 80% đúng khung.

**Học lại:** **slide 4** (Grain) trên `../B1_Nen_tang_Excel_Du_lieu/Practice/B1_MCI_Mart_Q1_2025_Answer.xlsx`, sheet Sales, đơn HD2501-00002 có 4 dòng (Line_No 1–4), và **slide 5** (`Slide_B2_HocVien.pptx`).

![Hình 6 · Một đơn có nhiều dòng sản phẩm](Images/NX06_don_nhieu_dong.png)
> **[CHỤP] NX06_don_nhieu_dong.png** · File: `../B1_Nen_tang_Excel_Du_lieu/Practice/B1_MCI_Mart_Q1_2025_Answer.xlsx` · Sheet: `Sales` · Vùng: A1:K7 (chọn dòng 3–6, ẩn tạm cột D–F, I–J) · Cần thấy: đơn HD2501-00002 chiếm 4 dòng Line_No 1–4.

![Hình 6 · KPI_Card của bạn, mục Dimensions và Target](Images/NX06b_card_dimensions_ban.png)
> **[CHỤP] NX06b_card_dimensions_ban.png** · File: `Homework/Bai_nop/b2_btvn_to-thuXRBUKA.xlsx` · Sheet: `KPI_Card` · Vùng: A9:B13 · Cần thấy: B9 (Region, Store_Code, Thang, Category, Phan_khuc) và B11 (304,000 VND). Bật Wrap Text.

### 5. Tham chiếu công thức (không trừ điểm)

**Bạn đã làm:** công thức dùng `Sales!$J:$J` và `Store!A:A`: chạy được.
**Gợi ý:** `tblSales_T4[Revenue]`, `tblStore[Region]` tự mở rộng khi thêm dữ liệu, đọc như tiếng Việt, và SUMIFS trên cột của Table không bao giờ bị `#VALUE!` do lệch kích thước vùng (lời giảng slide 16). Nên làm quen cho Buổi 3 trở đi, nhất là Pivot và measure ở Buổi 5.
**Học lại:** **slide 26 của Buổi 1** (`Slide_B1_HocVien.pptx`), `Demo/B1_Demo_Table_Validation.xlsx`, sheet Sales; **slide 16 của Buổi 2**.

## Câu insight

Câu của bạn: *"Doanh thu tháng 4/2025 là 296,965,000 VNĐ (với khu vực miền Nam dẫn đầu hệ thống đạt 142,061,000 VNĐ, chiếm gần 48% tổng doanh số), vì giá trị đơn hàng trung bình (AOV) tại miền Nam ở mức cao nhất toàn quốc (317,810 VNĐ/đơn) nhờ sức mua vượt trội ở hai ngành hàng chủ lực là Bánh kẹo và Gia dụng, nên phòng Marketing cần tiếp tục tập trung ngân sách tối ưu hóa các chiến dịch bán kèm combo sản phẩm tại khu vực miền Nam để giữ vững đà tăng trưởng dòng tiền."*

| Thành phần | Nhận xét |
|---|---|
| **Con số** | **Có, có so sánh và đúng.** 296.965.000, miền Nam 142.061.000 (47,8%), AOV miền Nam 317.810 cao nhất, đều khớp bảng KPI. |
| **Nguyên nhân** | **Có, nhưng chưa đủ sức giải thích.** Bánh kẹo và Gia dụng cũng là hai ngành lớn nhất ở miền Bắc và miền Trung, nên chưa thấy vì sao riêng miền Nam có AOV cao hơn. Miền Nam dẫn đầu chủ yếu nhờ **số đơn** (447, so với 317 và 202); AOV chỉ cao hơn miền Bắc khoảng 6% (317.810 so với 299.804). Thử nêu cả hai ý. |
| **Hàm ý** | **Có**, rõ đối tượng (Marketing) và việc làm (combo). "Giữ vững đà tăng trưởng" chưa có số liệu hỗ trợ vì bạn mới có một tháng; bỏ hoặc đổi thành điều bạn kiểm chứng được ở Buổi 3. |

**Mức hiện tại:** Competent (đủ ba thành phần, con số khớp). Câu đang dài, thử tách thành hai ý ngắn.
**Câu mẫu của giảng viên** (sheet Insight_mau trong file đáp án): *"Tháng 4/2025 MCI Mart bán 296.965.000 VNĐ qua 966 đơn (AOV 307.417), Gross Margin 33,8%; Miền Nam đóng góp nhiều nhất (47,8% doanh thu). Biên chênh lệch lớn giữa các ngành: Bánh kẹo 39,4% so với Thực phẩm khô 21,8%, nên cơ cấu ngành hàng của từng miền sẽ quyết định biên lợi nhuận — cần xem theo cửa hàng và theo ngày ở Buổi 3."* Để ý câu mẫu **không khẳng định điều chưa chứng minh**, kết bằng việc cần kiểm tra thêm: đúng tinh thần slide 40 (Buổi 2).

![Hình 7 · Doanh thu và số đơn theo miền (căn cứ cho nhận xét insight)](Images/NX07_kpi_theo_mien.png)
> **[CHỤP] NX07_kpi_theo_mien.png** · File: `Homework/Bai_nop/b2_btvn_to-thuXRBUKA.xlsx` · Sheet: `KPI` · Vùng: A1:L5 (bản rộng, hoặc ghép H1:L5 với A1:A5) · Cần thấy: Tổng theo miền (95.038.000 / 59.866.000 / 142.061.000), Số đơn (317 / 202 / 447), AOV (≈299.804 / 296.366 / 317.810).

![Hình 7 · Insight mẫu để so sánh](Images/NX07b_insight_mau.png)
> **[CHỤP] NX07b_insight_mau.png** · File: `Homework/B2_BTVN_MCI_Mart_T4_2025_Answer.xlsx` · Sheet: `Insight_mau` · Vùng: A1:A5 · Cần thấy: câu mẫu (bật Wrap Text, kéo rộng cột A).

## Việc cần làm trước Buổi 3

- [ ] Điền I5 `=SUM(I2:I4)` và J5 `=H5/I5`; kiểm tra ra 966 và 307.417.
- [ ] Sửa nhãn Phan_khuc: Cao / Trung / Thấp; lọc cột P chỉ còn 3 giá trị.
- [ ] Sửa B6, B7 của KPI_Card: công thức dùng `tblSales_T4`, `COUNTIFS(tblSales_T4[Line_No],1)`; nguồn đúng tên cột.
- [ ] Bỏ Category, Phan_khuc khỏi Dimensions của AOV; ghi nguồn cho Target.
- [ ] Làm quen viết công thức bằng tên cột Table (`tblSales_T4[Revenue]`).
