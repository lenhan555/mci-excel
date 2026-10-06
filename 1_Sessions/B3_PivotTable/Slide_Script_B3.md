# Kịch bản slide — Buổi 3: Tư duy dữ liệu & PivotTable

> **Phiên bản 1.0 (2026-09-28):** soạn theo `Course_Plan.md` v3.3, khuôn mẫu `Slide_Script_B2.md`. Lời giảng viết thành **văn nói đầy đủ**, số liệu tính từ dataset (script `2_Dataset/_generator/build_b3.py`; mọi file Excel đã được mở, Refresh và lưu bằng Excel, sheet `Kiem_tra` của đáp án 14/14 ✔).
> Căn cứ: `Course_Plan.md` v3.3 (B3: timetable 150 phút, Hướng dẫn giảng dạy B3, Nguyên tắc 1/7/9/10, Phụ lục A chuỗi A và C, Phụ lục D bảng chọn chart, Phụ lục F câu chuyện #1, #2, #6).

## Cách dùng kịch bản

- **Lời giảng** là lời nói mẫu. Giảng viên không cần đọc nguyên văn; nên nói theo giọng của mình nhưng giữ đủ **ý chính** và **câu hỏi cho lớp** (in đậm).
- Dòng *(→ …)* trong lời giảng là **chỉ dẫn hành động**. Đường dẫn file tính từ thư mục buổi này (`1_Sessions/B3_PivotTable/`).
- Mỗi slide có thời lượng dự kiến. Nếu lớp chậm, cắt ở những chỗ đánh dấu **[có thể rút gọn]**.
- Buổi này **làm cùng nhau** là chính: giảng viên dựng Pivot trên `Demo/B3_Demo_Pivot.xlsx`, học viên làm cùng lúc trên file khởi đầu của mình. Hai file có cùng bảng tblSales (6.187 dòng, cùng cột, cùng công thức), nên mọi con số trên màn hình giống nhau.
- Mỗi Pivot mới đặt ở sheet mới; đổi tên sheet ngay theo tên ghi trong lời giảng (PV_MienThang, PV_Thu…) để lúc cần quay lại dễ tìm.
- Cột Thu là `TEXT(ngày,"ddd")`: máy cài tiếng Anh hiện Mon…Sun, máy cài tiếng Việt hiện T2…CN. Số liệu không đổi. Excel 365 tự nhóm ngày theo tháng khi kéo Order_Date vào Rows; slide 19 và 29 nói cách nhóm lại theo ngày.

### Điều chỉnh cho lớp đã biết Excel sơ bộ (đề xuất, Course Plan chưa quy định cho B3)

Lớp đã quen Insert Pivot thì rút phần thao tác cơ bản, dùng thời gian cho phần phân tích. Tổng vẫn 150 phút.

| Slide | Bản chuẩn | Lớp đã biết Excel | Cách rút / cách dùng thêm |
|---|---|---|---|
| 9–10 Insert Pivot, Field List | 9' | 4' | Chỉ dựng nhanh Pivot Miền × Tháng, bỏ Recommended PivotTables |
| 14–15 Định dạng, Layout | 9.5' | 4.5' | Chỉ đổi Number Format và Outline; bỏ phần Style |
| Sau slide 26 | — | +10' | PRO TIP: Conditional Formatting trong Pivot (Color Scale trên Pivot Cửa hàng × Tháng), custom sort list, cached source data (handout mục 7) |

## Quy ước

Loại slide và nguồn ảnh giống `Slide_Script_B2.md`: BÌA · CHIA PHẦN · KHÁI NIỆM · CÔNG CỤ · THỰC HÀNH · QUIZ · TỔNG KẾT; **[CHỤP]** ảnh chụp Excel 365 Windows trên file của khóa, **[PDF pX]** trang PDF nguồn (`4_Reference/Excel_PivotTable_Slides.pdf`, trừ khi ghi khác), **[VẼ]** vẽ bằng shape, chỉ cho sơ đồ khái niệm và thẻ số (theo `CLAUDE.md`, mục Ảnh minh họa).

**File dùng trong buổi** (đường dẫn tính từ thư mục buổi; bản phát cho học viên có hậu tố `_HocVien`):

| File | Công dụng | Slide |
|---|---|---|
| `../B2_Ham_va_KPI/Practice/B2_MCI_Mart_Q1_2025_Answer.xlsx` | **File khởi đầu** (đáp án B2): tblSales làm giàu 6.187 dòng, KPI_SUMIFS. Phát đầu buổi | 3, 30 |
| `../B2_Ham_va_KPI/Homework/B2_BTVN_MCI_Mart_T4_2025_Answer.xlsx` | Đáp án bài về nhà B2, chữa nhanh đầu buổi | 3 |
| `Demo/B3_Demo_Pivot.xlsx` | File demo chính: tblSales giống file khởi đầu, KPI_SUMIFS (doanh thu, số đơn), Product, Store. Dựng mọi Pivot trực tiếp | 9, 10, 11, 14, 15, 18, 19, 20, 21, 23, 24, 25, 26, 28, 29 |
| `Demo/B3_Demo_Refresh.xlsx` | Cùng dữ liệu ở hai dạng: vùng thường (Sales_Vung) và Table (Sales_Table), hai Pivot dựng sẵn (PV_Vung, PV_Table), 10 dòng tháng 4 để dán (Them_dong) | 12 |
| `Demo/B3_Demo_Loi_Du_lieu.xlsx` | tblSalesLoi: cột Revenue có 2 ô chữ + 1 ô trống, cột Order_Date có 3 ô ngày dạng chữ | 19, 23 |
| `Practice/B3_MCI_Mart_Q1_2025_Answer.xlsx` | Đáp án B3: tblSales thêm So_don; 6 sheet Pivot (PV_…), Slicer Store_Code nối 3 Pivot, Timeline, 3 PivotChart, ST07_T3_ChiTiet, Kiem_tra (14/14 ✔). **File khởi đầu B4** | 32 |
| `Homework/B3_BTVN_MCI_Mart_T1_T4_2025.xlsx` | Bài tập về nhà bắt buộc (đề): tháng 1–4/2025, 8.106 dòng, cài sẵn 6 ô Quantity dạng chữ; sheet So_lieu_he_thong để đối soát | 42 |
| `Homework/B3_BTVN_MCI_Mart_T1_T4_2025_Answer.xlsx` | Đáp án bài về nhà: Pivot dựng sẵn, Kiem_tra 17 chỉ tiêu số liệu, CT_ST07_T3/T4, So_sanh (ST07, AOV, tỷ trọng), Huong_dan_cham đủ Câu 1–7, Dap_an_loi, Insight_mau. Chỉ giảng viên / trợ giảng | — |
| `Handout/B3_Handout.docx` / `.pdf` | Handout tự học (Show Values As mở rộng, Calculated Item, Report Filter Pages, Subtotal, PRO TIP, bài về nhà, checklist trước B4) | 42 |
| `Handout/B3_Infographic_TomTat.png` | Infographic tóm tắt buổi, phát riêng (không chèn vào handout); sửa ở `2_Dataset/_generator/infographic_b3.html` rồi chạy `build_infographic.py B3` | 42 |

Mọi file demo có sheet `Huong_dan` (Slide · Thao tác · Kết quả mong đợi). Bản giảng viên mở ra ở sheet làm việc, `Huong_dan` để cuối.

**Số liệu đáp án B3** (Quý 1/2025, dùng khi chữa bài):

| Chỉ tiêu | Giá trị |
|---|---|
| Tổng Pivot = tổng SUMIFS B2 | **941.454.000** |
| Ô Pivot Bắc × Tháng 3 = ô KPI_SUMIFS!D6 | **90.210.000** (635 dòng) |
| Sum of So_don / Count of Order_ID | **3.144 đơn** / 6.187 dòng |
| MoM (% Difference From) | T2 **−42,2%** · T3 **+9,6%** |
| Lũy kế (Running Total) | 425.850.000 · 671.846.000 · 941.454.000 |
| AOV (Calculated Field) | T1 304.614 · T2 298.902 · T3 292.100 · Quý 1 **299.445** |
| Gross Margin % (Calculated Field) | 35,2% · 33,7% · 33,1% · Quý 1 **34,2%** |

| Tỷ trọng ngành (% of Column Total) | Tháng 1 | Tháng 2 | Tháng 3 | Quý 1 |
|---|---|---|---|---|
| Bánh kẹo | **45,6%** | **30,9%** | 25,5% | 36,0% |
| Gia dụng | 17,7% | 23,3% | 25,9% | 21,5% |
| Thực phẩm khô | 15,0% | 18,8% | 21,7% | 17,9% |
| Đồ uống | 11,4% | 12,0% | 11,9% | 11,7% |
| Chăm sóc cá nhân | 9,1% | 13,7% | 12,1% | 11,1% |
| Sữa | 1,3% | 1,4% | 2,9% | 1,8% |

| Câu chuyện | Con số |
|---|---|
| #1 Tết (Group 7 ngày từ 01/01) | 01–07/01: 72.271.000 · **08–14/01: 118.778.000 · 15–21/01: 111.733.000 · 22–28/01: 112.944.000** · 29/01–04/02 (tuần Tết): **31.567.000** · các tuần sau: 53.025.000–73.803.000. Bình quân ngày: 09–28/01 16,5 triệu; 29/01–02/02 3,0 triệu; từ 03/02 9,0 triệu |
| #2 ST07 Nha Trang | T1 28.706.000 · T2 17.956.000 · T3 **8.539.000 (−52,4%, giảm sâu nhất)**; Show Details: 55 dòng, 28 đơn; 01–09/03: 13 đơn; 10–31/03: 15 đơn, 12/22 ngày không có đơn. ST04 −50,5% nhưng đơn rải đều |
| #6 Cuối tuần | Sun 159.845.000 · Sat 154.369.000 · Mon 137.855.000 · Wed 129.824.000 · Fri 126.513.000 · Tue 121.640.000 · Thu 111.408.000. Trung bình T7–CN 157,1 triệu so với T2–T6 125,4 triệu: **cao hơn khoảng 25%** |

**Tổng:** 42 slide, 150 phút.

---

## Phần 0 — Mở đầu: từ bảng SUMIFS đến câu hỏi "theo chiều nào" (10 phút · slide 1–4)

### Slide 1 · BÌA · 0.5'
- **Trên slide:** MCI / Consulting & Analytics / **Excel for Business Intelligence Analyst** / Buổi 3: Tư duy dữ liệu & PivotTable
- **Ảnh:** [VẼ] nền theo Style spec.

**Lời giảng:**
> Chào cả lớp, chào mừng các bạn đến Buổi 3. Buổi trước mình viết từng công thức SUMIFS để trả lời một câu hỏi. Hôm nay mình học công cụ giúp trả lời **hàng chục câu hỏi** chỉ bằng kéo thả: PivotTable.

### Slide 2 · KHÁI NIỆM · Câu hỏi mở buổi · 2.5'
- **Trên slide:** *"Làm sao trả lời 20 câu hỏi 'theo chiều nào' trong 20 phút?"* Góc slide có nhãn **Tầng 2 — Dimensional**.
- Dưới câu hỏi: 6 câu hỏi mẫu của Giám đốc bán hàng, mỗi câu một thẻ: *theo cửa hàng?* · *theo thứ trong tuần?* · *theo tuần?* · *theo ngành hàng?* · *cửa hàng nào giảm bất thường?* · *vì sao tháng 1 cao hơn tháng 2?*
- **Ảnh:** [VẼ] các thẻ câu hỏi, chữ "theo …" tô đậm.

**Lời giảng:**
> Tuần trước sếp hỏi "doanh thu theo miền, theo tháng", mình trả lời bằng một bảng SUMIFS. Tuần này sếp hỏi tiếp: theo cửa hàng thì sao, theo thứ trong tuần thì sao, tuần nào bán mạnh nhất, ngành nào chiếm bao nhiêu phần trăm. Và câu hỏi còn treo từ Buổi 2: **vì sao tháng 1 cao hơn hẳn tháng 2?**
>
> **Câu hỏi cho lớp: nếu dùng SUMIFS, mỗi câu hỏi mới các bạn phải làm gì?** *(→ chờ 2–3 câu trả lời)*
>
> Đúng rồi: dựng lại một bảng mới, gõ lại danh sách tiêu chí, kéo lại công thức. Hai mươi câu hỏi là hai mươi bảng. Chúng ta đang leo lên **tầng 2** của thang câu hỏi: không chỉ "bao nhiêu", mà "**ở đâu, khi nào, nhóm nào**". Tầng này cần một công cụ đổi chiều phân tích trong vài giây.

### Slide 3 · KHÁI NIỆM · Recap B2 & file khởi đầu · 4'
- **Trên slide:**
  - Viên gạch B2: tblSales làm giàu (Category, COGS, Gross_Margin, Thang, Thu…) · bảng KPI_SUMIFS: tổng **941.454.000** · T1 **425.850.000**, T2 **245.996.000**.
  - Bài về nhà B2 (tháng 4): doanh thu **296.965.000** · **966** đơn · lỗi cài sẵn: **4 mã Product_Code dính dấu cách cuối** (XLOOKUP ra #N/A; sửa bằng TRIM).
  - Hôm nay nhận **file khởi đầu** = đáp án B2 (`B2_MCI_Mart_Q1_2025_Answer.xlsx`), lưu thành `B3_<HoTen>.xlsx`.
- **Ảnh:** [VẼ] viên gạch B2 (đã xong) nằm dưới viên gạch B3 (sáng lên).

**Lời giảng:**
> Nhắc nhanh buổi trước. *(→ demo: mở `../B2_Ham_va_KPI/Practice/B2_MCI_Mart_Q1_2025_Answer.xlsx`, sheet KPI_SUMIFS, khối Doanh thu A4:E9)* Đây là bảng KPI chúng ta dựng bằng SUMIFS: tổng quý 941 triệu 454 nghìn, tháng 1 gần 426 triệu, tháng 2 chỉ 246 triệu. Câu hỏi "vì sao" vẫn đang treo; hôm nay mình trả lời.
>
> Bài về nhà Buổi 2: *(→ mở `../B2_Ham_va_KPI/Homework/B2_BTVN_MCI_Mart_T4_2025_Answer.xlsx`, sheet Kiem_tra)* tháng 4 doanh thu 296 triệu 965 nghìn, 966 đơn. Lỗi cài sẵn là 4 mã sản phẩm có dấu cách ở cuối, XLOOKUP ra #N/A. Bạn nào dùng IFERROR để che thì xem lại sheet Dap_an_loi: lỗi vẫn còn, chỉ là bị giấu đi.
>
> File khởi đầu hôm nay là đáp án Buổi 2, trợ giảng đã gửi trong nhóm chat. Các bạn mở ra, lưu thành tên mình, ví dụ `B3_NguyenVanA.xlsx`. **Giữ nguyên sheet KPI_SUMIFS**: lát nữa mình dùng nó để kiểm tra Pivot.

### Slide 4 · KHÁI NIỆM · Hôm nay leo lên tầng 2 · 3'
- **Trên slide:**
  - Thang 5 tầng: Tầng 1 Descriptive (B1–B2, đã xong) → **Tầng 2 Dimensional (hôm nay)** → chạm Tầng 3 Diagnostic bằng drill-down.
  - Ba câu hỏi sẽ tự tìm ra đáp án trong buổi: (1) Vì sao tháng 1 cao hơn tháng 2? (2) Cửa hàng nào giảm bất thường trong tháng 3? (3) Ngày nào trong tuần bán chạy nhất?
  - Viên gạch B3: **báo cáo Pivot tương tác** + bản nháp PivotChart.
- **Ảnh:** [VẼ] thang 5 tầng, tầng 2 sáng lên; 3 thẻ câu hỏi đánh số.

**Lời giảng:**
> Khóa học có thang 5 tầng câu hỏi. Hai buổi đầu ở tầng 1: chuyện gì đã xảy ra. Hôm nay lên **tầng 2**: xảy ra ở đâu, khi nào, nhóm nào. Và cuối buổi, nhờ một thao tác tên là drill-down, mình sẽ chạm vào **tầng 3**: vì sao.
>
> Có ba câu hỏi mình sẽ cùng đi tìm hôm nay. Mình **chưa nói đáp án**, các bạn tự tìm trong Pivot. Câu thứ nhất các bạn đã biết. Câu thứ hai: có một cửa hàng tháng 3 giảm rất bất thường, cửa hàng nào, và vì sao. Câu thứ ba: ngày nào trong tuần khách đi siêu thị nhiều nhất.
>
> Viên gạch hôm nay là một báo cáo Pivot có bộ lọc tương tác, cộng vài biểu đồ. Buổi 4 và Buổi 6 sẽ xây tiếp trên viên gạch này.

---

## Phần 1 — Vì sao cần PivotTable (10 phút · slide 5–7)

### Slide 5 · CHIA PHẦN · 0.5'
- **Trên slide:** MCI / **1. Vì sao cần PivotTable**

**Lời giảng:**
> Phần 1: SUMIFS đã tốt như vậy, vì sao còn cần Pivot?

### Slide 6 · KHÁI NIỆM · Giới hạn của SUMIFS · 4.5'
- **Trên slide:** so sánh hai cột.
  - **SUMIFS:** mỗi câu hỏi mới = một bảng mới. Cửa hàng × Tháng: gõ tay 12 mã cửa hàng, 3 tháng, kéo 36 ô. Theo tuần: bảng chưa có cột Tuần, phải thêm cột. Thêm một chiều = viết lại công thức.
  - **PivotTable:** kéo trường vào Rows / Columns / Values; đổi chiều phân tích trong vài giây; bấm đúp để xem dòng gốc.
  - Dòng cuối: 5 lợi ích của Pivot: Powerful · Beautiful · Fast · Accurate · Flexible [PDF p2–4].
- **Ảnh:** [VẼ] trái: lưới SUMIFS 12 × 3 với chữ "36 ô · 1 công thức / câu hỏi"; phải: 4 ô Filters / Columns / Rows / Values với mũi tên kéo thả.

**Lời giảng:**
> Mình thử làm phép tính. Sếp hỏi doanh thu theo cửa hàng và tháng. Với SUMIFS: gõ 12 mã cửa hàng, 3 tháng, viết công thức, khóa đô-la, kéo ra 36 ô. Mất khoảng 5 phút nếu không sai. Sếp hỏi tiếp theo tuần: tblSales chưa có cột tuần, phải thêm cột rồi mới SUMIFS. Rồi theo ngành hàng, theo thứ… **Mỗi chiều mới là làm lại từ đầu.**
>
> Đó là giới hạn chúng ta đã thấy cuối Buổi 2. PivotTable giải quyết đúng chỗ này: các bạn chỉ **kéo** một trường vào hàng, một trường vào cột, một số vào giá trị. Đổi câu hỏi thì kéo trường khác vào. Không viết công thức, không khóa đô-la, không sợ sót dòng.
>
> Nhưng mình muốn các bạn nhớ: Pivot **không phải phép màu**. Lát nữa các bạn sẽ thấy mỗi ô trong Pivot chính là một SUMIFS mà Excel viết giúp mình.

### Slide 7 · KHÁI NIỆM · Dữ liệu cho Pivot: tidy + Table · 5'
- **Trên slide:**
  - Điều kiện nguồn [PDF p5]: hình chữ nhật (mỗi cột một trường, mỗi dòng một bản ghi) · tiêu đề rõ, một dòng · không dòng tổng, không ô gộp · nên là **Excel Table** (B1).
  - Vai trò của cột → vùng Pivot (Chuỗi A: Dimension / Measure từ B2):

| Vai trò | Ví dụ trong tblSales | Kéo vào |
|---|---|---|
| **Dimension** ("theo cái gì") | Region, Store_Code, Thang, Thu, Category, Order_Date | Rows · Columns · Filters |
| **Measure** ("đo cái gì") | Revenue, Quantity, COGS, Gross_Margin | Values |
| **Grain** ("mỗi dòng là gì") | 1 sản phẩm trong 1 đơn | Quyết định Pivot đếm cái gì |

- **Ảnh:** [VẼ] đầu bảng tblSales, cột tô màu theo vai trò, mũi tên sang 4 vùng Pivot.
- **TIP:** Dữ liệu đã tidy ở B1 là lý do hôm nay chỉ cần kéo thả. Dữ liệu còn dòng tổng hay ô gộp thì Pivot cộng sai.

**Lời giảng:**
> Pivot chỉ chạy tốt trên dữ liệu **tidy**, đúng như Buổi 1: mỗi dòng một bản ghi, mỗi cột một trường, không có dòng tổng xen giữa, không ô gộp. Nếu file còn dòng "Tổng Miền Bắc" ở giữa, Pivot sẽ cộng cả dòng đó và các bạn nhận được con số gấp đôi. tblSales của chúng ta đã sạch từ Buổi 1, và đã là Excel Table.
>
> Còn nhớ ba từ của Buổi 2 không? **Dimension** là "theo cái gì": vùng, cửa hàng, tháng. Trong Pivot, dimension đi vào **Rows, Columns hoặc Filters**. **Measure** là "đo cái gì": doanh thu, số lượng. Measure đi vào **Values**. Và **grain**, mỗi dòng là một sản phẩm trong một đơn, sẽ quyết định Pivot đang đếm cái gì. Lát nữa các bạn sẽ thấy Pivot "đếm" ra 6.187 chứ không phải 3.144, đúng cái bẫy của Buổi 2.

---

## Phần 2 — Tạo Pivot và Pivot tính như thế nào (20 phút · slide 8–12)

### Slide 8 · CHIA PHẦN · 0.5'
- **Trên slide:** MCI / **2. Tạo PivotTable: mỗi ô là một bộ lọc**

**Lời giảng:**
> Phần 2: dựng Pivot đầu tiên, và hiểu Excel tính mỗi ô như thế nào.

### Slide 9 · CÔNG CỤ · Insert PivotTable · 4'

| Tên (đỏ) | Mục đích |
|---|---|
| **Insert PivotTable** | Tạo Pivot trống từ một Table / vùng dữ liệu |
| **Recommended PivotTables** | Excel gợi ý sẵn vài Pivot để bắt đầu |

- **Ảnh:** [CHỤP] hộp Create PivotTable: Table/Range = tblSales, chọn New Worksheet: `Demo/B3_Demo_Pivot.xlsx`, sheet Sales, chọn ô A2 → Insert → PivotTable → From Table/Range; [PDF p6].
- **TIP:** Chọn **một ô** trong Table rồi Insert → PivotTable: Excel tự điền tên Table. Phím tắt: Alt, N, V, T.

**Lời giảng:**
> *(→ demo: mở `Demo/B3_Demo_Pivot.xlsx`, sheet Sales; các bước ở sheet Huong_dan)* Các bạn làm cùng mình trên file khởi đầu. Bấm vào **một ô bất kỳ** trong bảng, không cần bôi đen cả bảng. Vào Insert, PivotTable, From Table/Range.
>
> Hộp thoại hỏi hai câu, đúng như slide: **dữ liệu ở đâu**, và **đặt Pivot ở đâu**. Ô Table/Range đã tự ghi tblSales, vì mình đang đứng trong Table. Chọn New Worksheet, OK. *(→ đổi tên sheet mới thành PV_MienThang)*
>
> Nút bên cạnh, **Recommended PivotTables**, là Excel đoán giúp vài Pivot. Tiện để khám phá, nhưng mình khuyên các bạn tự dựng, vì mình bắt đầu từ **câu hỏi** chứ không bắt đầu từ gợi ý của máy.

### Slide 10 · CÔNG CỤ · PivotTable Fields: 4 vùng · 5'
- **Tên (đỏ):** **PivotTable Fields**
- **Mục đích:** Danh sách trường của nguồn và 4 vùng: **Filters** (lọc cả Pivot) · **Columns** · **Rows** · **Values** (số cần tính).
- **Ảnh:** [CHỤP] Pivot Region × Thang ra 3 × 3 + Grand Total, và khung PivotTable Fields bên phải: Region ở Rows, Thang ở Columns, Sum of Revenue ở Values: `Demo/B3_Demo_Pivot.xlsx`, sheet PV_MienThang (sheet vừa tạo), A3:E8 + khung Fields; [PDF p7].
- **TIP:** Dimension vào Rows / Columns, Measure vào Values. Kéo nhầm thì kéo trường ra khỏi vùng là xong.

**Lời giảng:**
> Bên phải là khung **PivotTable Fields**: trên là 20 cột của tblSales, dưới là 4 vùng. *(→ trên `Demo/B3_Demo_Pivot.xlsx`, sheet PV_MienThang: kéo Region vào Rows)* Kéo Region vào Rows: Bắc, Nam, Trung. *(→ kéo Thang vào Columns)* Thang vào Columns: 1, 2, 3. *(→ kéo Revenue vào Values)* Và Revenue vào Values.
>
> Mười giây. Các bạn nhìn: đây chính là bảng Miền × Tháng Buổi 2 mình mất 12 phút để dựng. Grand Total góc dưới: **941 triệu 454 nghìn**. Tháng 1 **425.850.000**, tháng 2 **245.996.000**.
>
> **Câu hỏi cho lớp: muốn xem theo cửa hàng thay vì theo miền, làm gì?** *(→ chờ; kéo Region ra, kéo Store_Code vào Rows, rồi trả lại Region)* Đúng, kéo Region ra, kéo Store_Code vào. Đó là toàn bộ sức mạnh của Pivot.

### Slide 11 · KHÁI NIỆM · Pivot hoạt động thế nào: tọa độ = bộ lọc · 5'
- **Trên slide:** [VẼ] hình xương sống Chuỗi C (dùng lại nguyên hình Buổi 2, slide 15):
  - Bên phải: **ô Pivot Bắc × 3** = 90.210.000.
  - Hai "tọa độ" đi vào ô: **[Region = Bắc]** (hàng) và **[Thang = 3]** (cột).
  - Bên trái: cột Revenue → **phép tổng**.
  - Chú thích dưới: *B2: criteria của SUMIFS · **B3: tọa độ ô Pivot** · B5: filter context.*
- **Ảnh:** [CHỤP] Pivot Miền × Tháng và khối đối soát: ô B13 `=GETPIVOTDATA("Doanh thu",$A$4,"Region","Bắc","Thang",3)` = 90.210.000, SUMIFS B2 = 90.210.000, chênh lệch 0: `Practice/B3_MCI_Mart_Q1_2025_Answer.xlsx`, sheet PV_MienThang, A4:D15; [PDF p11].
- **TIP:** Gõ `=` rồi bấm vào ô Pivot, Excel tự viết **GETPIVOTDATA**. Kéo công thức đó không đổi kết quả; muốn tắt: PivotTable Analyze → Options ▾ → bỏ tick Generate GetPivotData.

**Lời giảng:**
> Giờ đến phần quan trọng nhất của Phần 2. Pivot tính ô Bắc, tháng 3 như thế nào? *(→ chỉ ô Bắc × 3 trên sheet PV_MienThang)* Excel làm đúng ba bước: xác định **tọa độ** của ô, Region bằng Bắc và Thang bằng 3; **lọc** những dòng thỏa cả hai; rồi **cộng** Revenue.
>
> Nghe quen không? Đó chính là hình xương sống Buổi 2. Buổi 2 gọi hai điều kiện đó là *criteria* của SUMIFS; hôm nay gọi là **tọa độ của ô Pivot**. **Cùng một ý, cái tên thứ hai.**
>
> Mình kiểm chứng. *(→ trên `Demo/B3_Demo_Pivot.xlsx`: View → New Window → Arrange → Vertical; cửa sổ trái sheet PV_MienThang, cửa sổ phải sheet KPI_SUMIFS)* Bên trái ô Pivot Bắc × 3, bên phải ô D6 của bảng SUMIFS: cả hai **90.210.000**. *(→ double-click ô Bắc × 3 trong Pivot)* Và khi bấm đúp, Excel mở một sheet mới chứa đúng **635 dòng** Bắc tháng 3: những dòng đã được cộng. *(→ xóa sheet chi tiết vừa tạo)*
>
> Một mẹo đối soát. *(→ ở ô trống G4 của PV_MienThang gõ `=`, bấm vào ô Bắc × 3, Enter)* Excel không viết `=D6` mà viết **GETPIVOTDATA**: lấy giá trị theo tên trường và tên mục. Hay ở chỗ Pivot có đổi bố cục thì công thức vẫn đúng ô. Nhưng nếu các bạn định kéo công thức sang ô bên cạnh thì nó không đổi kết quả; lúc đó tắt tính năng này như TIP trên slide.
>
> Quy tắc đối soát của hôm nay: **tổng Pivot phải bằng tổng SUMIFS của Buổi 2**. Trong file đáp án, sheet PV_MienThang có sẵn khối đối soát, chênh lệch bằng 0.

### Slide 12 · CÔNG CỤ · Refresh · Change Data Source · 5.5'

| Tên (đỏ) | Mục đích |
|---|---|
| **Refresh** (Alt+F5) · **Refresh All** (Ctrl+Alt+F5) | Cập nhật Pivot sau khi dữ liệu nguồn thay đổi |
| **Change Data Source** | Đổi / mở rộng vùng nguồn của Pivot |

- **Ảnh:** [CHỤP] PV_Table và PV_Vung sau khi dán 10 dòng tháng 4 và Refresh All: PV_Table tổng **942.937.000**, PV_Vung vẫn **941.454.000**: `Demo/B3_Demo_Refresh.xlsx`, sheet PV_Table và PV_Vung (A3:B7); [CHỤP] hộp Change PivotTable Data Source: `Demo/B3_Demo_Refresh.xlsx`, sheet PV_Vung; [PDF p8–10].
- **TIP:** Pivot đọc từ **bộ nhớ đệm (cache)**, không đọc thẳng ô: sửa dữ liệu xong phải Refresh.
- **PRO TIP:** Nguồn là **Table** thì dòng mới tự vào Pivot khi Refresh; nguồn là vùng thường `$A$1:$I$6188` thì không.

**Lời giảng:**
> *(→ demo: mở `Demo/B3_Demo_Refresh.xlsx`, sheet PV_Table; các bước ở sheet Huong_dan)* File này có cùng dữ liệu ở hai dạng: sheet Sales_Vung là vùng thường, Sales_Table là Excel Table. Mỗi dạng có một Pivot. Hiện hai Pivot giống hệt nhau, tổng 941.454.000.
>
> Thí nghiệm 1. *(→ sheet Sales_Table, ô G2: đổi Quantity từ 1 thành 101; quay lại PV_Table)* Mình vừa tăng số lượng của dòng đầu tiên thêm 100, tức là thêm 1 triệu 900 nghìn doanh thu cho Miền Bắc. **Pivot có đổi không?** *(→ chờ)* Không. Vì Pivot không đọc thẳng từ ô, nó đọc từ một **bản sao trong bộ nhớ**, gọi là cache. *(→ chuột phải trong Pivot → Refresh)* Refresh xong, Bắc tăng đúng 1.900.000. *(→ Ctrl+Z ở Sales_Table, Refresh lại)*
>
> Thí nghiệm 2. *(→ sheet Them_dong: copy A4:H13; dán vào A6189 của Sales_Table và A6189 của Sales_Vung; ở Sales_Vung kéo công thức I6188 xuống I6198)* Đây là 10 dòng tháng 4 vừa về. Mình dán xuống dưới cả hai bảng. *(→ Data → Refresh All)* Refresh All. **PV_Table** thêm 1.483.000, thành 942.937.000. **PV_Vung đứng yên** 941.454.000, vì nguồn của nó là vùng cố định từ dòng 1 đến dòng 6188.
>
> *(→ PV_Vung → PivotTable Analyze → Change Data Source → sửa thành `Sales_Vung!$A$1:$I$6198`)* Muốn sửa thì phải vào Change Data Source, và tháng nào cũng phải nhớ làm. Đây là lý do Buổi 1 mình chuyển dữ liệu thành **Table**: Table tự lớn lên, Pivot tự theo. Lỗi "quên Refresh" và "nguồn không phải Table" là hai lỗi Pivot hay gặp nhất ngoài đời.

---

## Phần 3 — Định dạng báo cáo (10 phút · slide 13–15)

### Slide 13 · CHIA PHẦN · 0.5'
- **Trên slide:** MCI / **3. Định dạng: Pivot dễ đọc cho người xem**

**Lời giảng:**
> Pivot đúng rồi, giờ làm cho nó dễ đọc. Phần này ngắn.

### Slide 14 · CÔNG CỤ · Value Field Settings · PivotTable Styles · 4.5'

| Tên (đỏ) | Mục đích |
|---|---|
| **Value Field Settings → Number Format** | Đặt tên và định dạng số cho cả trường giá trị |
| **PivotTable Styles · Grand Totals · Subtotals** | Kiểu bảng, bật / tắt dòng tổng |

- **Ảnh:** [CHỤP] hộp Value Field Settings (Custom Name: Doanh thu) mở Number Format (Number, 0 chữ số, Use 1000 Separator): `Demo/B3_Demo_Pivot.xlsx`, sheet PV_MienThang; [PDF p13–14].
- **TIP:** Định dạng qua Value Field Settings, **không** bôi đen ô rồi định dạng: bôi đen thì Pivot đổi bố cục là mất.
- **PRO TIP:** Conditional Formatting trong Pivot [PDF p17]: chọn "All cells showing 'Doanh thu' values for Store_Code and Thang" để quy tắc tự áp cho ô mới.

**Lời giảng:**
> *(→ trên `Demo/B3_Demo_Pivot.xlsx`, sheet PV_MienThang: chuột phải một số → Value Field Settings)* Hai việc: đổi tên "Sum of Revenue" thành **Doanh thu**, và bấm Number Format, chọn Number, 0 chữ số thập phân, tick dấu phân cách nghìn. *(→ OK)* Số giờ dễ đọc: 145.814.000.
>
> Vì sao không bôi đen rồi bấm định dạng như bảng thường? Vì Pivot thay đổi hình dạng liên tục; lúc thêm một trường, ô vừa định dạng có thể chuyển thành chỗ khác. Định dạng qua Value Field Settings là định dạng **cho cả trường**, đi theo trường đó.
>
> *(→ PivotTable Design → PivotTable Styles: chọn một style; Grand Totals → On for Columns Only, rồi trả lại)* Tab Design có kiểu bảng và nút bật tắt dòng tổng. **[có thể rút gọn]**

### Slide 15 · CÔNG CỤ · Report Layout: Compact · Outline · Tabular · 5'
- **Tên (đỏ):** **Report Layout**
- **Mục đích:** Cách xếp nhiều trường trên hàng: **Compact** (mặc định, một cột thụt lề) · **Outline** (mỗi trường một cột, nên dùng khi phân tích) · **Tabular** (mỗi trường một cột, tổng phụ ở dưới).
- **Ảnh:** [CHỤP] 3 ảnh cùng Pivot Region + Store_Code trên hàng, Doanh thu: Compact / Outline / Tabular (+ Repeat All Item Labels): `Demo/B3_Demo_Pivot.xlsx`, sheet PV_MienThang, A3:F20; [PDF p15–16].
- **TIP:** Muốn chép kết quả Pivot ra làm bảng dữ liệu mới: Tabular + **Repeat All Item Labels** + tắt Subtotals và Grand Totals.

**Lời giảng:**
> *(→ trên `Demo/B3_Demo_Pivot.xlsx`, sheet PV_MienThang: kéo Store_Code vào Rows, nằm dưới Region)* Khi có hai trường trên hàng, bố cục bắt đầu quan trọng. Mặc định là **Compact**: miền và cửa hàng chung một cột, thụt lề. Gọn, nhưng cột tiêu đề chỉ ghi "Row Labels".
>
> *(→ PivotTable Design → Report Layout → Show in Outline Form)* **Outline**: mỗi trường một cột có tên riêng, lọc riêng được. Tài liệu Pivot khuyên dùng bố cục này khi phân tích. *(→ Show in Tabular Form, rồi Repeat All Item Labels)* **Tabular** có đường kẻ như bảng thường; bật Repeat All Item Labels thì mỗi dòng có đủ tên miền, rất tiện khi cần copy kết quả sang chỗ khác.
>
> *(→ bỏ Store_Code khỏi Rows, để Pivot về Miền × Tháng; chọn lại Outline Form)* Trong file đáp án, mọi Pivot để Outline.

### Slide 16 · NGHỈ · 10'
- **Trên slide:** Nghỉ 10 phút. Trước khi quay lại: mở `Demo/B3_Demo_Loi_Du_lieu_HocVien.xlsx` (trợ giảng đã gửi).

**Lời giảng:**
> Mình nghỉ 10 phút. Khi quay lại, các bạn để sẵn file của mình và file `Demo/B3_Demo_Loi_Du_lieu_HocVien.xlsx` trợ giảng vừa gửi trong nhóm chat. Phần sau là sắp xếp, lọc, nhóm và đào sâu: phần tìm câu chuyện.
>
> *(→ trong giờ nghỉ: kiểm tra học viên đã có sheet PV_MienThang ra tổng 941.454.000)*

---

## Phần 4 — Sắp xếp, lọc, nhóm, đào sâu (20 phút · slide 17–21)

### Slide 17 · CHIA PHẦN · 0.5'
- **Trên slide:** MCI / **4. Sắp xếp, lọc, nhóm, đào sâu: đi tìm câu chuyện**

**Lời giảng:**
> Phần 4: dùng Pivot để đi tìm ba câu chuyện lúc đầu buổi.

### Slide 18 · CÔNG CỤ · Sort · Value Filters (Top 5) · 4'

| Tên (đỏ) | Mục đích |
|---|---|
| **Sort** | Sắp xếp mục theo tên hoặc theo giá trị |
| **Label / Value Filters** | Lọc mục theo tên (chứa, bắt đầu bằng…) hoặc theo giá trị (Top 10, lớn hơn…) |

- **Ảnh:** [CHỤP] Pivot Thu × Doanh thu sắp xếp giảm dần (Sun 159.845.000 … Thu 111.408.000): `Practice/B3_MCI_Mart_Q1_2025_Answer.xlsx`, sheet PV_Thu, A4:B12; [CHỤP] hộp Top 10 Filter (Top 5 Items by Sum of Revenue) và kết quả 5 sản phẩm: `Demo/B3_Demo_Pivot.xlsx`, sheet PV_Top5; [PDF p19].
- **Thẻ số:** T7–CN trung bình **157,1 triệu** / thứ · T2–T6 **125,4 triệu** → cuối tuần cao hơn **khoảng 25%**.
- **PRO TIP:** Thứ tự T2 → CN cố định: tạo **Custom List** (File → Options → Advanced → Edit Custom Lists), rồi Sort → More Sort Options → A to Z.

**Lời giảng:**
> Câu chuyện thứ ba trước, vì nó nhanh nhất. *(→ trên `Demo/B3_Demo_Pivot.xlsx`: Insert → PivotTable từ tblSales, New Worksheet, đổi tên sheet thành PV_Thu; Thu vào Rows, Revenue vào Values)* Doanh thu theo thứ trong tuần. *(→ chuột phải một số → Sort → Sort Largest to Smallest)* Sắp xếp giảm dần.
>
> **Chủ nhật** đứng đầu, 159,8 triệu, rồi **thứ Bảy** 154,4 triệu. Thấp nhất là thứ Năm, 111,4 triệu. Quý 1 có 13 ngày mỗi thứ, riêng thứ Ba 12 ngày, nên so tổng được. Trung bình hai ngày cuối tuần 157 triệu, trung bình năm ngày thường 125 triệu: **cuối tuần cao hơn khoảng 25%**. Câu hỏi số 3 đã có đáp án.
>
> *(→ Insert → PivotTable mới, đổi tên sheet PV_Top5; Product_Name vào Rows, Revenue vào Values; mũi tên ở ô tiêu đề hàng → Value Filters → Top 10 → sửa 10 thành 5 → OK)* Thêm một câu hỏi sếp hay hỏi: 5 sản phẩm bán nhiều tiền nhất. Value Filters, Top 10, sửa thành 5. Đứng đầu là **Hộp quà Tết cao cấp, 135,85 triệu**, rồi chảo chống dính, thùng mì, dầu gội, hạt điều. Một hộp quà Tết đứng đầu cả quý: thêm một manh mối cho câu hỏi số 1.

### Slide 19 · CÔNG CỤ · Group: theo ngày, theo nhóm tự chọn · 5.5'
- **Tên (đỏ):** **Group**
- **Mục đích:** Gộp mục thành nhóm mới: ngày → tuần / tháng / quý; số → khoảng; chữ → nhóm tự đặt tên.
- **Ảnh:** [CHỤP] hộp Grouping (By: Days, Number of days: 7) và Pivot 13 tuần: `Demo/B3_Demo_Pivot.xlsx`, sheet PV_Tuan, A3:B17; [CHỤP] thông báo "Cannot group that selection.": `Demo/B3_Demo_Loi_Du_lieu.xlsx`, sheet Pivot mới từ tblSalesLoi; [PDF p20].
- **Thẻ số:** 3 tuần trước Tết **118,8 · 111,7 · 112,9 triệu** · tuần Tết 29/01–04/02 **31,6 triệu** · tuần bình thường 53–74 triệu.
- **TIP:** "Cannot group that selection" = cột ngày có ô **chữ** hoặc ô trống → sửa ở nguồn, Refresh.

**Lời giảng:**
> Giờ đến câu hỏi số 1: vì sao tháng 1 cao hơn tháng 2? Theo tháng thì quá thô, mình cần nhìn **theo tuần**. tblSales không có cột tuần, nhưng Pivot tự nhóm được.
>
> *(→ trên `Demo/B3_Demo_Pivot.xlsx`: Insert → PivotTable mới, đổi tên sheet PV_Tuan; Order_Date vào Rows, Revenue vào Values)* Excel 365 tự nhóm ngày thành tháng. *(→ chuột phải một mục ngày → Group → bỏ chọn Months, chọn Days, Number of days: 7 → OK)* Mình đổi: nhóm theo **Days**, mỗi nhóm **7 ngày**, bắt đầu từ 01/01.
>
> Các bạn đọc cột số cùng mình. Tuần đầu tháng 1: 72 triệu. Ba tuần sau đó, **từ 8 đến 28 tháng 1: 119, 112, 113 triệu**, gần gấp đôi một tuần bình thường. Rồi tuần **29/01 đến 04/02: chỉ 31,6 triệu**. Mùng 1 Tết Ất Tỵ là 29/01/2025. Sau Tết, các tuần quay về khoảng 53 đến 74 triệu.
>
> **Vậy vì sao tháng 1 cao hơn tháng 2?** *(→ chờ)* Không phải tháng 2 bán kém: tháng 1 chứa trọn **ba tuần mua sắm Tết**, còn tháng 2 bắt đầu bằng **kỳ nghỉ Tết**. Câu hỏi số 1 có đáp án; lát nữa ta còn một bằng chứng nữa ở phần Show Values As.
>
> Group còn làm được với chữ. *(→ Pivot Store_Code × Revenue; giữ Ctrl chọn ST08, ST09, ST10 → chuột phải → Group; đổi tên Group1 thành TP.HCM)* Ba cửa hàng ở TP.HCM gộp lại: 309.645.000. **[có thể rút gọn]**
>
> Và một lỗi các bạn sẽ gặp. *(→ mở `Demo/B3_Demo_Loi_Du_lieu.xlsx`, sheet Sales_Loi: Insert → PivotTable, Order_Date vào Rows → chuột phải → Group)* Excel báo "**Cannot group that selection**". Nhìn lên đầu danh sách: ba mục 11/01/2025, 15/01/2025, 23/03/2025 đứng riêng, trước cả các ngày thật, vì đó là **chữ** trông giống ngày, ở dòng 1913, 2274, 3926 của Sales_Loi. Chỉ một ô ngày dạng chữ là Pivot không nhóm được. Sửa ở nguồn, gõ lại thành ngày thật, Refresh.

### Slide 20 · CÔNG CỤ · Show Details (drill-down) · 5'
- **Tên (đỏ):** **Show Details** (double-click ô giá trị)
- **Mục đích:** Mở sheet mới chứa **các dòng gốc** tạo nên con số: từ "bao nhiêu" sang "vì sao".
- **Ảnh:** [CHỤP] Pivot Store_Code × Thang, ô ST07 × 3 = 8.539.000 được chọn: `Practice/B3_MCI_Mart_Q1_2025_Answer.xlsx`, sheet PV_CuaHang, A4:E18; [CHỤP] sheet chi tiết sau khi double-click, sắp xếp theo Order_Date (55 dòng, từ 10/03 thưa hẳn): `Practice/B3_MCI_Mart_Q1_2025_Answer.xlsx`, double-click ô D12 (ST07 × 3) của PV_CuaHang; [PDF p11].
- **Thẻ số:** ST07: tháng 2 **51 đơn** · 01–09/03 **13 đơn** · 10–31/03 **15 đơn trong 22 ngày, 12 ngày không có đơn**.
- **TIP:** Sheet chi tiết là bản sao tĩnh: xem xong thì xóa, đừng sửa dữ liệu ở đó.

**Lời giảng:**
> Câu hỏi số 2: cửa hàng nào bất thường? *(→ trên `Demo/B3_Demo_Pivot.xlsx`: Insert → PivotTable mới, đổi tên sheet PV_CuaHang; Store_Code vào Rows, Thang vào Columns, Revenue vào Values)* Cửa hàng × Tháng: 36 ô, 10 giây.
>
> Tháng 3 cả công ty **tăng 9,6%** so với tháng 2, nhưng các bạn dò cột 3: ST07 Nha Trang chỉ còn **8,5 triệu**, tháng 2 là 18 triệu. Giảm hơn một nửa, sâu nhất trong 12 cửa hàng. **Vì sao?** Nhìn bảng tổng thì không thể biết.
>
> *(→ double-click ô ST07 × 3)* Double-click. Excel mở sheet mới với **55 dòng**: đúng những dòng làm nên con số 8,5 triệu. *(→ sắp xếp Order_Date tăng dần)* Các bạn nhìn cột ngày. Từ mùng 1 đến 9 tháng 3: 13 đơn, gần như ngày nào cũng có. **Từ 10/03 trở đi**: cả 22 ngày chỉ có 15 đơn, và 12 ngày không có đơn nào. Tháng 2 cửa hàng này bán 51 đơn.
>
> Dữ liệu chỉ nói được đến đây: **có chuyện gì đó bắt đầu từ ngày 10/03** ở ST07. Hỏi phòng vận hành mới biết: cửa hàng **cải tạo** từ 10 đến 31/03. Đây là tầng 3, "vì sao", và cách làm là: con số bất thường → drill-down → tìm mốc thời gian → hỏi người trong cuộc.
>
> Một bài học D2. *(→ xóa sheet chi tiết; double-click ô ST04 × 3)* ST04 Hạ Long cũng giảm 50%. Nhưng drill-down thì thấy 41 đơn rải đều cả tháng, không có quãng trống. **Cùng con số, khác câu chuyện.** Đừng kết luận trước khi bấm đúp. *(→ xóa sheet chi tiết)*

### Slide 21 · CÔNG CỤ · Slicer · Timeline · Report Connections · 5'

| Tên (đỏ) | Mục đích |
|---|---|
| **Slicer** | Nút bấm để lọc Pivot, nhìn là biết đang lọc gì |
| **Timeline** | Thanh thời gian để lọc theo ngày / tháng / quý / năm |
| **Report Connections** | Nối một Slicer / Timeline với nhiều Pivot cùng nguồn |

- **Ảnh:** [CHỤP] Slicer Store_Code (2 cột) nối 3 Pivot, hộp Report Connections: `Practice/B3_MCI_Mart_Q1_2025_Answer.xlsx`, sheet PV_MienThang, A4:I20 + hộp Report Connections; [CHỤP] Timeline Order_Date trên PV_Ngay: `Practice/B3_MCI_Mart_Q1_2025_Answer.xlsx`, sheet PV_Ngay, D4:N11; [PDF p21].
- **TIP:** Slicer chỉ nối được các Pivot dùng **chung cache** (cùng nguồn). Bỏ lọc: nút phễu có dấu X (Alt+C).

**Lời giảng:**
> Báo cáo gửi sếp cần lọc được, mà sếp không quen mũi tên lọc nhỏ xíu của Pivot. **Slicer** là bộ lọc dạng nút bấm. *(→ trên `Demo/B3_Demo_Pivot.xlsx`, sheet PV_Thu: PivotTable Analyze → Insert Slicer → tick Region → OK)* Bấm Nam: Pivot chỉ còn Miền Nam.
>
> Nhưng mình muốn **một** Slicer lọc **nhiều** Pivot. *(→ chuột phải Slicer → Report Connections → tick thêm Pivot ở sheet PV_CuaHang → OK; bấm Nam)* Report Connections: tick các Pivot muốn nối. Giờ bấm Nam, cả Pivot thứ lẫn Pivot cửa hàng cùng lọc. Đây là tiêu chí kiểm tra cuối buổi: **Slicer lọc được mọi Pivot liên quan**.
>
> *(→ Clear Filter; PivotTable Analyze → Insert Timeline → tick Order_Date → OK; kéo chọn tháng 2 và tháng 3)* **Timeline** là Slicer dành cho ngày tháng. Mình dùng nó để kiểm tra lại câu chuyện cuối tuần: bỏ tháng 1 có Tết ra, chỉ xem tháng 2 và 3. Chủ nhật 93,6 triệu, thứ Bảy 89 triệu vẫn đứng đầu. Câu chuyện cuối tuần **không phải do Tết**. Đó là cách một nhà phân tích tự phản biện kết luận của mình. *(→ Clear Filter trên Timeline)*

---

## Phần 5 — Tính toán trong Pivot (20 phút · slide 22–26)

### Slide 22 · CHIA PHẦN · 0.5'
- **Trên slide:** MCI / **5. Tính toán trong Pivot: tính trên kết quả đã tổng hợp**

**Lời giảng:**
> Phần 5: không chỉ cộng. Pivot đếm, lấy trung bình, tính tỷ trọng, tăng trưởng, và cả công thức riêng của mình.

### Slide 23 · CÔNG CỤ · Summarize Values By · 4.5'
- **Tên (đỏ):** **Summarize Values By**
- **Mục đích:** Chọn phép tổng hợp: Sum · Count · Average · Max · Min…
- **Ảnh:** [CHỤP] Pivot từ tblSalesLoi, Revenue vào Values ra **Count of Revenue** = 6.186 thay vì Sum: `Demo/B3_Demo_Loi_Du_lieu.xlsx`, sheet Pivot mới, A3:B7; [PDF p24].
- **Thẻ số (tblSales):** Sum **941.454.000** · Count **6.187 dòng** (không phải 3.144 đơn) · Average **152.166 / dòng** · Max **1.470.000**.
- **TIP:** Kéo một cột số vào Values mà Pivot hiện **Count** → cột đó có ô **trống hoặc chữ**. Đổi sang Sum thì Excel lặng lẽ bỏ qua các ô đó.

**Lời giảng:**
> *(→ trên `Demo/B3_Demo_Pivot.xlsx`, sheet PV_MienThang: chuột phải một số → Summarize Values By → Count)* Đổi từ Sum sang Count: tổng **6.187**. **Đây có phải số đơn không?** *(→ chờ)* Không: nhớ grain, Count đếm **dòng**. Số đơn là 3.144; lát nữa mình đếm đúng bằng Calculated Field. *(→ Average, rồi Max, rồi trả về Sum)* Average: 152 nghìn mỗi dòng. Max: 1 triệu 470 nghìn, dòng lớn nhất.
>
> Giờ lỗi hay gặp nhất. *(→ mở `Demo/B3_Demo_Loi_Du_lieu.xlsx`, sheet Sales_Loi: Insert → PivotTable, Region vào Rows, Revenue vào Values)* Cùng thao tác, nhưng Pivot hiện **Count of Revenue**. Vì sao? Tài liệu Pivot nói rõ: cột có **ô trống hoặc ô chữ** thì Excel mặc định Count.
>
> *(→ Summarize Values By → Sum)* Đổi sang Sum: **940.480.000**, thiếu 974.000 so với 941.454.000, không một lời cảnh báo. *(→ sheet Sales_Loi: lọc cột Revenue, dòng 1372, 4674 là số dạng chữ, dòng 5183 trống)* Hai ô là số lưu dạng chữ, một ô trống. *(→ Convert to Number cho 2 ô chữ; gõ 356000 vào ô I5183; Refresh)* Sửa xong, Refresh: 941.454.000. Bài học D2: **Pivot hiện Count là một lời cảnh báo**, đừng chỉ bấm sang Sum rồi đi tiếp.

### Slide 24 · CÔNG CỤ · Show Values As: % of Column Total · % of Grand Total · 5'
- **Tên (đỏ):** **Show Values As**
- **Mục đích:** Hiển thị con số dưới dạng khác (tỷ trọng, chênh lệch, lũy kế…) **mà không tạo dữ liệu mới**.
- **Ảnh:** [CHỤP] Pivot Category × Thang, % of Column Total (Bánh kẹo 45,6% → 30,9% → 25,5%): `Practice/B3_MCI_Mart_Q1_2025_Answer.xlsx`, sheet PV_NganhHang, A4:E12; [PDF p25–26].
- **Thẻ số:** 20 ngày trước Tết (09–28/01) Bánh kẹo chiếm **50,8%** doanh thu, ngày thường **27,9%**. Đồ uống bán gấp khoảng 1,8 lần mỗi ngày nhưng tỷ trọng gần như giữ nguyên (11,7% so với 11,5%).
- **TIP:** % of **Column** Total = mỗi cột cộng lại 100%. Tháng ở cột → so cơ cấu **trong từng tháng**.

**Lời giảng:**
> *(→ trên `Demo/B3_Demo_Pivot.xlsx`: Insert → PivotTable mới, đổi tên sheet PV_NganhHang; Category vào Rows, Thang vào Columns, Revenue vào Values; chuột phải một số → Show Values As → % of Column Total)* Bằng chứng thứ hai cho câu hỏi số 1. Mình đổi số tiền thành **tỷ trọng trong từng tháng**: mỗi cột cộng lại 100%.
>
> Tháng 1, **Bánh kẹo chiếm 45,6%** doanh thu. Tháng 2 còn 30,9%, tháng 3 còn 25,5%. Bánh kẹo là mứt Tết, hộp quà, hạt điều: hàng Tết. Nếu tính riêng 20 ngày trước Tết thì Bánh kẹo chiếm hơn một nửa doanh thu, 50,8%.
>
> **Câu hỏi cho lớp: Đồ uống cũng là hàng Tết, sao tỷ trọng Đồ uống tháng 1 lại không tăng?** *(→ chờ)* Vì tỷ trọng là phần của một cái bánh đang to ra. Trước Tết, Đồ uống bán gấp khoảng 1,8 lần mỗi ngày, nhưng cả cửa hàng cũng tăng gần bằng vậy, nên phần của Đồ uống vẫn khoảng 11–12%. Đọc tỷ trọng thì luôn nhớ: **tỷ trọng đứng yên không có nghĩa là doanh thu đứng yên**.
>
> *(→ bỏ Thang khỏi Columns, đổi Rows thành Region → Show Values As → % of Grand Total)* % of Grand Total: Nam 45,4%, Bắc 34,4%, Trung 20,2% doanh thu quý. Show Values As **không tạo dữ liệu mới**: nó tính lại trên chính các con số đã tổng hợp. *(→ trả Pivot về Category × Thang, % of Column Total)*

### Slide 25 · CÔNG CỤ · Show Values As: % Difference From · Running Total In · 4.5'

| Tên (đỏ) | Mục đích | Kết quả (Quý 1/2025) |
|---|---|---|
| **% Difference From** (previous) | Tăng trưởng so với mục liền trước (MoM) | T2 **−42,2%** · T3 **+9,6%** |
| **Running Total In** | Lũy kế theo thứ tự mục | 425.850.000 · 671.846.000 · 941.454.000 |

- **Ảnh:** [CHỤP] Pivot theo Thang: Doanh thu, MoM %, Lũy kế, Số đơn, AOV, Gross Margin %: `Practice/B3_MCI_Mart_Q1_2025_Answer.xlsx`, sheet PV_Thang, A4:G9; [PDF p26].
- **TIP:** Base item **(previous)** = so với tháng liền trước; tháng đầu tiên để trống vì không có tháng trước.

**Lời giảng:**
> *(→ trên `Demo/B3_Demo_Pivot.xlsx`: Insert → PivotTable mới, đổi tên sheet PV_Thang; Thang vào Rows; kéo Revenue vào Values 3 lần)* Ba cột giống nhau. Cột thứ nhất giữ nguyên. *(→ cột thứ hai: Show Values As → % Difference From → Base field: Thang, Base item: (previous))* Cột thứ hai là **tăng trưởng so với tháng trước**: tháng 2 **âm 42,2%**, tháng 3 **dương 9,6%**. Tháng 1 để trống, vì trước nó không có tháng nào trong dữ liệu.
>
> *(→ cột thứ ba: Show Values As → Running Total In → Base field: Thang)* Cột thứ ba là **lũy kế**: đến hết tháng 2 được 672 triệu, hết quý 941 triệu. Cột này hay dùng để so với kế hoạch năm.
>
> Để ý: ba cột cùng lấy từ **một** trường Revenue. Buổi 2 muốn tính MoM các bạn phải viết công thức riêng; ở đây chỉ là một lựa chọn trong menu. *(→ đổi tên 3 cột: Doanh thu, MoM %, Lũy kế; định dạng % 1 chữ số cho cột MoM)*

### Slide 26 · CÔNG CỤ · Calculated Field · 5.5'
- **Tên (đỏ):** **Calculated Field**
- **Mục đích:** Tạo trường mới bằng công thức trên **tổng** của các trường khác: `AOV = Revenue / So_don`.
- **Ảnh:** [CHỤP] hộp Insert Calculated Field: Name AOV, Formula `=Revenue/So_don`: `Demo/B3_Demo_Pivot.xlsx`, sheet PV_Thang; [PDF p28–29].
- **Thẻ số:** Count of Order_ID **6.187** ✘ · Sum of So_don **3.144** ✔ · AOV Quý 1 **299.445** = 941.454.000 ÷ 3.144 (tổng ÷ tổng).
- **PRO TIP:** Calculated Field luôn tính trên **tổng** (SUM) của các trường. Vì vậy KPI tỷ lệ tính bằng Calculated Field tự động là **tổng ÷ tổng**. Cần đếm thì thêm cột đếm (=1) vào nguồn [PDF p29].

**Lời giảng:**
> Pivot đếm ra 6.187 dòng, vậy làm sao đếm **đơn**? Mẹo trong tài liệu Pivot: thêm vào nguồn một cột đếm. *(→ trên `Demo/B3_Demo_Pivot.xlsx`, sheet Sales: thêm cột So_don ở U1, ô U2 gõ `=IF([@Line_No]=1,1,0)`)* Mỗi đơn có đúng một dòng Line_No bằng 1, nên cột So_don có số 1 ở dòng đầu của mỗi đơn, còn lại là 0. *(→ quay lại PV_Thang, Refresh, kéo So_don vào Values)* Sum of So_don: **3.144 đơn**. Đúng con số Buổi 2.
>
> Giờ AOV, giá trị đơn trung bình. *(→ PivotTable Analyze → Fields, Items & Sets → Calculated Field; Name: AOV; Formula: =Revenue/So_don → Add → OK)* Tạo trường mới tên AOV bằng doanh thu chia số đơn. Tháng 1: 304.614, tháng 2: 298.902, tháng 3: 292.100. **Grand Total: 299.445.**
>
> **Câu hỏi cho lớp: 299.445 là trung bình của ba con số AOV tháng, hay là gì?** *(→ chờ)* Trung bình ba tháng là khoảng 298.500, không khớp. 299.445 là **tổng doanh thu chia tổng số đơn**: Calculated Field luôn tính trên **tổng**. Đúng quy tắc "tổng chia tổng" của Buổi 2, và Pivot làm giúp mình.
>
> *(→ tạo thêm Calculated Field GM_pct = Gross_Margin/Revenue, định dạng %)* Làm tương tự với biên lợi nhuận: 34,2% cả quý, khớp bảng KPI Buổi 2.

---

## Phần 6 — PivotChart: chọn biểu đồ theo câu hỏi (10 phút · slide 27–29)

### Slide 27 · CHIA PHẦN · 0.5'
- **Trên slide:** MCI / **6. PivotChart: biểu đồ trả lời câu hỏi**

**Lời giảng:**
> Phần 6: biến Pivot thành biểu đồ, và chọn đúng loại biểu đồ. Đây là bản xem trước của Buổi 6.

### Slide 28 · KHÁI NIỆM · 3 câu hỏi chọn chart · 4'
- **Trên slide:**
  - 3 câu hỏi then chốt [DataViz PDF p5]: (1) **Dữ liệu loại gì?** (số, nhóm, thời gian, địa lý…) (2) **Muốn truyền đạt gì?** (xu hướng, so sánh, cơ cấu, phân phối, quan hệ) (3) **Ai xem?** (nhà phân tích, giám đốc, khách hàng…)
  - Bảng chọn nhanh (Course Plan, Phụ lục D):

| Câu hỏi dữ liệu | Chart |
|---|---|
| Xu hướng theo thời gian | **Line** |
| So sánh giữa các nhóm | **Column / Bar** |
| Cơ cấu / tỷ trọng | **Stacked column / bar** (Donut ≤ 5 phần) |

- **Ảnh:** [CHỤP] cùng Pivot Thang × Doanh thu vẽ bằng Pie (không thấy xu hướng) và bằng Line: `Demo/B3_Demo_Pivot.xlsx`, sheet PV_Thang (hai PivotChart).
- **TIP:** Tiêu đề biểu đồ là **câu kết luận**: "Doanh thu sụt 42% trong tháng Tết", không phải "Doanh thu theo tháng".

**Lời giảng:**
> Trước khi chèn biểu đồ, hỏi ba câu. Dữ liệu là loại gì? Mình muốn nói điều gì: xu hướng, so sánh hay cơ cấu? Và ai sẽ xem?
>
> *(→ trên `Demo/B3_Demo_Pivot.xlsx`, sheet PV_Thang: chọn Pivot → PivotTable Analyze → PivotChart → Pie)* Mình cố tình làm sai: doanh thu ba tháng vẽ bằng biểu đồ tròn. **Nhìn biểu đồ này, các bạn có thấy tháng 2 giảm 42% không?** *(→ chờ)* Rất khó. Biểu đồ tròn nói "mỗi tháng chiếm bao nhiêu phần của quý", không nói "tăng hay giảm". *(→ Change Chart Type → Line)* Câu hỏi là **xu hướng theo thời gian**, nên dùng **đường**: thấy ngay tháng 1 cao, tháng 2 sụt, tháng 3 nhích lên.
>
> Quy tắc thứ hai: tiêu đề biểu đồ là **câu kết luận**, để người xem biết ngay mình nên thấy gì. Buổi 6 mình học đủ 8 loại biểu đồ lõi; hôm nay chỉ cần ba dòng trên slide.

### Slide 29 · CÔNG CỤ · PivotChart · 5.5'
- **Tên (đỏ):** **PivotChart**
- **Mục đích:** Biểu đồ gắn với Pivot: đổi trường, lọc Slicer / Timeline thì biểu đồ đổi theo.
- **Ảnh:** [CHỤP] PivotChart đường doanh thu theo ngày Quý 1 (đỉnh 3 tuần trước Tết, đáy 29/01–02/02) và Timeline phía trên: `Practice/B3_MCI_Mart_Q1_2025_Answer.xlsx`, sheet PV_Ngay, D4:P33; [PDF p34–36].
- **TIP:** Ẩn các nút lọc trên biểu đồ: PivotChart Analyze → Field Buttons → Hide All (handout mục 5).

**Lời giảng:**
> *(→ trên `Demo/B3_Demo_Pivot.xlsx`: Insert → PivotTable mới, đổi tên sheet PV_Ngay; Order_Date vào Rows; chuột phải → Group → chỉ chọn Days (Number of days: 1) → OK; Revenue vào Values; PivotTable Analyze → PivotChart → Line)* Bằng chứng cuối cho câu hỏi số 1, dạng hình. Doanh thu **từng ngày** Quý 1, biểu đồ đường.
>
> Các bạn thấy hình dạng rất rõ: đầu tháng 1 khoảng 10 triệu mỗi ngày. **Từ 9/1 vọt lên**, trung bình 16,5 triệu mỗi ngày, trong 20 ngày. Rồi **rơi xuống đáy**: 29/01 đến 02/02 chỉ khoảng 3 triệu mỗi ngày, kỳ nghỉ Tết. Sau đó đi ngang quanh 9 triệu, có những gai nhỏ đều đặn: đó là các ngày cuối tuần, câu chuyện số 3.
>
> *(→ bấm vào tiêu đề biểu đồ, gõ: "3 tuần trước Tết mỗi ngày bán ~16,5 triệu, nghỉ Tết còn ~3 triệu")* Đặt tiêu đề là câu kết luận. *(→ PivotTable Analyze → Insert Timeline → Order_Date; chọn tháng 1)* Và vì là PivotChart, Timeline lọc Pivot thì biểu đồ đổi theo.
>
> Trong file đáp án, mỗi biểu đồ có một dòng "Câu hỏi:" phía trên. Tiêu chí kiểm tra cuối buổi: **mỗi PivotChart ghi rõ câu hỏi nó trả lời**.

---

## Phần 7 — Thực hành trên case (30 phút · slide 30–32)

### Slide 30 · THỰC HÀNH 1 · Báo cáo Pivot & đối soát (15 phút)
- **Trên slide:**
  1. File của bạn (`B3_<HoTen>.xlsx`): thêm cột **So_don** `=IF([@Line_No]=1,1,0)` vào tblSales.
  2. Sheet **PV_MienThang**: Region × Thang, Doanh thu (`#,##0`, Outline). Đối soát: `=GETPIVOTDATA(…"Bắc"…3)−KPI_SUMIFS!D6` = 0.
  3. Sheet **PV_NganhHang**: Category × Thang, **% of Column Total**.
  4. Sheet **PV_Thang**: Thang × Doanh thu, **MoM %** (% Difference From, previous), **Lũy kế**, **Số đơn** (Sum So_don), **AOV** (Calculated Field).
  5. **Slicer Store_Code**, Report Connections nối 3 Pivot.
- **Kiểm tra:** Tổng **941.454.000** · Bắc–T3 **90.210.000** · Bánh kẹo T1 **45,6%** · MoM T2 **−42,2%** · **3.144** đơn · AOV **299.445**.

**Lời giảng (giao bài):**
> Giờ các bạn tự dựng báo cáo, 15 phút, trên file của mình. Phần lớn thao tác mình đã làm cùng nhau; bài này gom lại thành một báo cáo gọn: mỗi Pivot một sheet, đặt tên đúng như slide.
>
> Bước 2 là bước **D2**: dựng xong Pivot Miền × Tháng, đối soát ngay với bảng SUMIFS Buổi 2 bằng GETPIVOTDATA. Chênh lệch phải bằng 0. Nếu tổng không phải 941.454.000 thì dừng lại tìm nguyên nhân trước khi làm tiếp.
>
> Bước 5: Slicer Store_Code, chuột phải, Report Connections, tick cả ba Pivot. Bấm thử ST07: cả ba Pivot phải cùng đổi.

**Ghi chú giảng viên:**
- Đi vòng lớp. **Ba lỗi hay gặp nhất:**
  - Thêm cột So_don nhưng quên Refresh → không thấy trường So_don trong Field List.
  - Report Connections không thấy Pivot khác → Pivot đó được tạo từ vùng khác (không phải tblSales), khác cache; tạo lại từ tblSales.
  - Calculated Field gõ `=SUM(Revenue)/SUM(So_don)` → sai cú pháp; chỉ cần `=Revenue/So_don`.
- Excel 2016 không tự nhóm ngày; Excel 365 tự nhóm Order_Date thành tháng (bài 1 không dùng Order_Date).
- Ai xong sớm: làm trước Thực hành 2.

### Slide 31 · THỰC HÀNH 2 · Tìm câu chuyện bằng PivotChart (12 phút)
- **Trên slide:**
  1. Sheet **PV_Ngay**: Order_Date (Group → Days) × Doanh thu → **PivotChart Line**; thêm **Timeline** Order_Date.
  2. Sheet **PV_Thu**: Thu × Doanh thu, sắp xếp giảm dần → **PivotChart Column**.
  3. Sheet **PV_CuaHang**: Store_Code × Thang; thêm cột **% so tháng trước**; double-click ô **ST07 × 3**, ghi số đơn trước và sau 10/03.
  4. Mỗi biểu đồ: tiêu đề là **câu kết luận**, dòng "Câu hỏi:" ở ô A2.
- **Kiểm tra:** CN **159.845.000** đứng đầu, T5 **111.408.000** cuối · ST07 T3 **8.539.000 (−52,4%)**, **55** dòng chi tiết · 3 tuần trước Tết ~**16,5 triệu/ngày**.

**Lời giảng (giao bài):**
> Bài 2 là ba câu chuyện của hôm nay, 12 phút. Lần này các bạn tự làm và tự viết tiêu đề biểu đồ bằng lời của mình. Tiêu đề phải là một câu kết luận có con số, ví dụ "Chủ nhật bán nhiều nhất tuần: 160 triệu", không phải "Doanh thu theo thứ".
>
> Phần ST07: bấm đúp, sắp xếp theo ngày, đếm số đơn trước và sau 10/03. Ghi con số vào ô ngay dưới Pivot. Đó là nguyên liệu cho câu insight cuối buổi.

**Ghi chú giảng viên:**
- Order_Date bị Excel 365 tự nhóm thành tháng → chuột phải → Group → bỏ Months, chọn Days. Muốn tắt hẳn: File → Options → Data → tick "Disable automatic grouping of Date/Time columns in PivotTables".
- PivotChart của Pivot đang lọc bằng Slicer cũng bị lọc theo; nhắc Clear Filter trước khi đọc số.
- Nếu Thu hiện T2…CN (máy tiếng Việt): số liệu vẫn giống slide.

### Slide 32 · KHÁI NIỆM · Chữa bài nhanh · 3'
- **Trên slide:** Checklist brick output B3:
  - [ ] Tổng Pivot = tổng SUMIFS B2 = **941.454.000** (khối đối soát chênh lệch 0)
  - [ ] PV_NganhHang, PV_Thang dùng Show Values As; AOV bằng Calculated Field (**299.445**)
  - [ ] Slicer Store_Code lọc được cả 3 Pivot
  - [ ] PivotChart theo ngày, theo thứ; tiêu đề là câu kết luận, có dòng "Câu hỏi:"
  - [ ] Đã drill-down ST07 tháng 3
- **Chân slide:** File đáp án `B3_MCI_Mart_Q1_2025_Answer.xlsx` (sheet Kiem_tra: 14 dòng đối soát) là file khởi đầu Buổi 4.

**Lời giảng:**
> *(→ mở `Practice/B3_MCI_Mart_Q1_2025_Answer.xlsx`, sheet Kiem_tra)* File đáp án có sheet Kiem_tra 14 dòng, đọc thẳng từ Pivot bằng GETPIVOTDATA, tất cả tick xanh. *(→ sheet PV_MienThang: bấm vài nút trên Slicer Store_Code)* Slicer lọc cùng lúc Pivot miền, ngành hàng và tháng. *(→ sheet ST07_T3_ChiTiet)* Sheet ST07_T3_ChiTiet là kết quả drill-down, các dòng từ 10/03 được tô hồng.
>
> File này là **file khởi đầu Buổi 4**. Lưu ý: nếu các bạn đang bấm Slicer thì số trên Kiem_tra sẽ đổi theo; bỏ lọc trước khi đối chiếu.
>
> *(→ hỏi lớp)* **Bạn nào tick đủ 5 ô?**

**Ghi chú giảng viên (kiểm tra cuối buổi, theo Course Plan):**
- Tổng Pivot = tổng SUMIFS của B2 (941.454.000); Slicer lọc được mọi Pivot liên quan; mỗi PivotChart ghi rõ câu hỏi nó trả lời.
- Học viên dùng Mac: **hạn chót có máy Windows là trước Buổi 4** (Power Query và Data Model). Hỏi lại từng bạn.
- D2 mong đợi: Developing → hướng tới Competent (tự đối soát Pivot với SUMIFS; drill-down trước khi kết luận).

---

## Phần 8 — Quiz, insight, preview (10 phút · slide 33–42)

Quiz kiểu hỏi nhanh: khoảng 40 giây mỗi câu. **Câu 6–7 là câu tình huống.**

### Slide 33 · QUIZ · EXCEL QUIZ CHALLENGE · 0.2'

**Lời giảng:**
> 7 câu hỏi nhanh. Câu 6 và 7 là tình huống: không có công thức nào để thuộc, các bạn phải suy luận.

### Slide 34 · QUESTION 1: PIVOT
- Ô Pivot Bắc × Tháng 3 = 90.210.000 tương đương công thức nào? A) `=SUM(tblSales[Revenue])` · **B) `=SUMIFS(tblSales[Revenue],tblSales[Region],"Bắc",tblSales[Thang],3)`** · C) `=COUNTIFS(tblSales[Region],"Bắc",tblSales[Thang],3)` · D) `=AVERAGEIFS(tblSales[Revenue],tblSales[Region],"Bắc")`

**Lời giảng:**
> B. Mỗi ô Pivot là một SUMIFS: tọa độ hàng và cột chính là các tiêu chí. C đếm dòng, D lấy trung bình và thiếu điều kiện tháng.

### Slide 35 · QUESTION 2: REFRESH
- Bạn sửa Quantity trong tblSales nhưng Pivot không đổi. Vì sao? A) Pivot bị khóa · **B) Pivot đọc từ bộ nhớ đệm, cần Refresh** · C) Excel đang tính thủ công · D) Pivot chỉ đọc cột Revenue

**Lời giảng:**
> B. Pivot đọc từ cache. Sửa nguồn xong bấm Refresh (Alt+F5) hoặc Refresh All (Ctrl+Alt+F5); hoặc bật "Refresh data when opening the file".

### Slide 36 · QUESTION 3: SUMMARIZE VALUES BY
- Kéo Revenue vào Values, Pivot hiện "Count of Revenue". Nguyên nhân khả dĩ nhất? A) Revenue là cột công thức · **B) Cột Revenue có ô trống hoặc ô chữ** · C) Pivot luôn mặc định Count · D) Chưa Refresh

**Lời giảng:**
> B. Cột số có ô trống hoặc chữ thì Excel mặc định Count. Đó là tín hiệu cần kiểm tra nguồn, không chỉ đổi sang Sum: Sum sẽ bỏ qua các ô đó mà không báo.

### Slide 37 · QUESTION 4: SHOW VALUES AS
- Tháng ở cột, ngành hàng ở hàng. Muốn biết mỗi ngành chiếm bao nhiêu % doanh thu **của từng tháng**? A) % of Grand Total · B) % of Row Total · **C) % of Column Total** · D) Running Total In

**Lời giảng:**
> C. Tháng nằm ở cột, nên mỗi cột cộng lại 100%. % of Row Total sẽ cho biết mỗi tháng chiếm bao nhiêu % doanh thu quý của một ngành: câu hỏi khác.

### Slide 38 · QUESTION 5: CALCULATED FIELD
- Calculated Field `AOV = Revenue / So_don`. Ô Grand Total tính thế nào? A) Trung bình AOV của các tháng · **B) Tổng Revenue ÷ tổng So_don** · C) AOV tháng lớn nhất · D) Tổng AOV các tháng

**Lời giảng:**
> B. Calculated Field luôn tính trên tổng: 941.454.000 chia 3.144 bằng 299.445. Đúng quy tắc tổng chia tổng của Buổi 2.

### Slide 39 · QUESTION 6: TÌNH HUỐNG — CHỌN CHART
- Sếp muốn nhìn doanh thu **90 ngày** Quý 1 để thấy ảnh hưởng của Tết. Chọn loại biểu đồ nào?
  - A) Pie
  - **B) Line**
  - C) Stacked column 100%
  - D) Scatter

**Lời giảng:**
> B. Câu hỏi là xu hướng theo thời gian: dùng đường. Pie 90 miếng không đọc được; stacked 100% nói cơ cấu, không nói tăng giảm.

### Slide 40 · QUESTION 7: TÌNH HUỐNG — CHỌN CÔNG CỤ
- Khi nào **vẫn nên** dùng SUMIFS thay vì Pivot?
  - A) Khi cần đổi chiều phân tích liên tục
  - **B) Khi báo cáo có bố cục cố định và ô kết quả được công thức khác tham chiếu (thẻ KPI, so với target)**
  - C) Khi dữ liệu có nhiều dòng
  - D) Khi cần xem các dòng gốc

**Lời giảng:**
> B. Pivot mạnh khi khám phá, đổi chiều liên tục; nhưng bố cục của nó thay đổi theo dữ liệu. Khi cần một ô cố định để công thức khác tham chiếu, ví dụ thẻ KPI hay bảng so với target, SUMIFS vẫn là lựa chọn tốt. A và D là lý do dùng Pivot; C thì cả hai đều làm được.

### Slide 41 · TỔNG KẾT + INSIGHT · 2'
- **Trên slide:**
  - **Điểm chính cần nhớ:**
    - (1) Mỗi ô Pivot = một SUMIFS: tọa độ là bộ lọc.
    - (2) Nguồn Pivot là Table; Pivot đọc cache → Refresh.
    - (3) Con số bất thường → drill-down trước khi kết luận.
    - (4) Show Values As và Calculated Field tính trên kết quả đã tổng hợp: tỷ lệ luôn tổng ÷ tổng.
    - (5) Chọn chart theo câu hỏi; tiêu đề là câu kết luận.
  - **Insight của bạn (1 câu):** "Doanh thu ___ là ___ *(con số, có so sánh)*, vì ___ *(nguyên nhân khả dĩ)*, nên ___ *(hàm ý)*."

**Lời giảng:**
> Tổng kết 5 điều cần nhớ *(→ đọc lướt 5 ý trên slide)*.
>
> Giờ viết câu insight, 1 phút. Lần này các bạn có ba câu chuyện để chọn.
>
> *(→ chờ 1 phút, gọi 2–3 bạn đọc)*
>
> Câu mẫu: "Doanh thu tháng 1/2025 là 425,9 triệu, cao hơn tháng 2 khoảng 73%, vì 3 tuần trước Tết (08–28/01) mỗi tuần bán 112–119 triệu, gần gấp đôi tuần bình thường, và Bánh kẹo chiếm 45,6% doanh thu tháng 1, rồi tuần Tết chỉ còn 31,6 triệu; nên tháng 1 và tháng 2 không so trực tiếp được, kế hoạch và đánh giá hai tháng này phải tính theo lịch Tết."
>
> Để ý: câu này đã nói "**vì Tết**", và lần này có bằng chứng: nhóm theo tuần, tỷ trọng ngành, biểu đồ theo ngày. Buổi 2 mình mới nghi ngờ; hôm nay mình chứng minh. Còn ST07 thì mình mới biết "có chuyện từ 10/03"; muốn biết cửa hàng đạt bao nhiêu phần trăm kế hoạch, cần ghép với bảng Target ở Buổi 4 và Buổi 5.

### Slide 42 · PREVIEW B4 · 1'
- **Trên slide:**
  - *Buổi sau:* "Làm sao để việc làm sạch dữ liệu chỉ làm **một lần**, từ đó chỉ cần bấm Refresh?" (Power Query)
  - Giới hạn hôm nay: dữ liệu tháng mới về dạng **CSV bẩn, tách nhiều file** → lại phải copy-paste, sửa tay, rồi mới Refresh được Pivot.
  - *Về nhà:* đọc handout B3 · làm **Bài tập về nhà (bắt buộc)** `B3_BTVN_MCI_Mart_T1_T4_2025.xlsx`, nộp trước Buổi 4 tại hocvien.mcivietnam.com, tên file `B3_BTVN_<HoTen>.xlsx` · tick "Checklist trước Buổi 4" (mục 10 handout).
  - **Buổi 4 bắt buộc máy Windows** có Power Query (Data → Get Data).
- **Ảnh:** [VẼ] viên gạch B3 (đã xong) → viên gạch B4 (sáng lên).

**Lời giảng:**
> Hôm nay các bạn đã trả lời được hai mươi câu hỏi "theo chiều nào" chỉ bằng kéo thả, và tìm ra ba câu chuyện trong dữ liệu. Nhưng để ý: dữ liệu hôm nay đã **sạch sẵn**. Ngoài đời, tháng 4 sẽ về dưới dạng vài file CSV, cột ngày lộn xộn, mã thừa dấu cách, như các bạn thấy ở Buổi 1. Mỗi tháng lại copy-paste và sửa tay rồi mới Refresh được Pivot. Buổi 4 học **Power Query**: làm sạch **một lần**, tháng sau chỉ bấm Refresh.
>
> Về nhà: đọc handout Buổi 3, có Calculated Item, Report Filter Pages, Subtotal. Làm **bài tập về nhà bắt buộc**: dữ liệu 4 tháng đầu năm, dựng Pivot trả lời 5 câu hỏi, trong đó có câu "ST07 đã hồi phục chưa". Việc đầu tiên là **đối soát** Pivot với số liệu hệ thống có sẵn trong file: sẽ có một cột lệch, các bạn phải tự tìm ra vì sao. Nộp trước Buổi 4 trên hocvien.mcivietnam.com.
>
> Và nhắc lần cuối: **Buổi 4 bắt buộc máy Windows**. Bạn nào dùng Mac mà chưa cài máy ảo thì liên hệ trợ giảng ngay tuần này. Cảm ơn cả lớp, hẹn gặp lại Buổi 4!

---

## Phụ lục kịch bản

### Kiểm tra thời lượng

| Phần | Slide | Phút | Timetable B3 |
|---|---|---|---|
| 0 Mở đầu, recap, câu hỏi BI | 1–4 | 10 | 10 |
| 1 Vì sao cần Pivot, tidy data | 5–7 | 10 | 10 |
| 2 Insert, Field List, Pivot tính thế nào, Refresh | 8–12 | 20 | 20 |
| 3 Number format, Style, Layout | 13–15 | 10 | 10 |
| Nghỉ | 16 | 10 | 10 |
| 4 Sort, Filter, Group, Show Details, Slicer, Timeline | 17–21 | 20 | 20 |
| 5 Summarize Values By, Show Values As, Calculated Field | 22–26 | 20 | 20 |
| 6 PivotChart + 3 câu hỏi chọn chart | 27–29 | 10 | 10 |
| 7 Thực hành | 30–32 | 30 | 30 |
| 8 Quiz, insight, preview | 33–42 | 10 | 10 |
| **Tổng** | **42** | **150** | **150** |

Phần 8 cộng từng slide khoảng 8 phút (slide 33: 0,2'; 7 câu × khoảng 40 giây; slide 41: 2'; slide 42: 1'); khoảng 2 phút dự phòng cho câu 6–7 và phần đọc insight.

### Danh sách ảnh cần chụp [CHỤP] (Excel 365 Windows, zoom 100%)

Theo `CLAUDE.md` (mục Ảnh minh họa): thao tác và kết quả Excel dùng ảnh chụp thật trên file của khóa. Đã chụp đủ 24 ảnh dưới đây (28/09/2026, Excel 365 Windows, zoom 100%) và chèn vào deck; bảng giữ lại để chụp lại khi file hoặc lời giảng đổi. Mọi slide CÔNG CỤ có ít nhất một ảnh chụp; slide KHÁI NIỆM có demo (11, 28) cũng có ảnh. Hình [VẼ] còn lại chỉ là sơ đồ khái niệm và thẻ số (slide 2, 3, 4, 6, 7, 11 hình xương sống, 18–20 và 24, 26 thẻ số, 42). Ảnh lưu vào `Materials/B3/Images/` đúng tên file ở cột 2; chạy lại `build_slides_b3.py` thì ảnh tự thay vào khung (danh sách cả deck và handout: `Materials/B3/Images/_DANH_SACH_ANH.md`).

Một số ảnh cần làm thao tác trước khi chụp (dán dòng, Group, double-click): làm đúng các bước ở sheet `Huong_dan` của file demo, rồi chụp.

| # | Tên file | Nội dung ảnh | File · sheet · vùng | Slide |
|---|---|---|---|---|
| 1 | S09_create_pivot.png | Hộp Create PivotTable: Table/Range = tblSales, New Worksheet | `Demo/B3_Demo_Pivot.xlsx` · Sales · chọn A2 → Insert → PivotTable → From Table/Range | 9 |
| 2 | S10_field_list.png | Pivot Region × Thang (3 × 3 + Grand Total) và khung PivotTable Fields: Region ở Rows, Thang ở Columns, Sum of Revenue ở Values | `Demo/B3_Demo_Pivot.xlsx` · PV_MienThang (sheet vừa tạo) · A3:E8 + khung Fields | 10 |
| 3 | S11_pivot_vs_sumifs.png | Pivot Miền × Tháng và khối đối soát; chọn B13, thanh công thức hiện `=GETPIVOTDATA("Doanh thu",$A$4,"Region","Bắc","Thang",3)`; chênh lệch 0 | `Practice/B3_MCI_Mart_Q1_2025_Answer.xlsx` · PV_MienThang · A4:D15 | 11 |
| 4 | S12a_refresh_table_vs_range.png | Sau khi dán 10 dòng tháng 4 và Refresh All: PV_Table tổng 942.937.000, PV_Vung vẫn 941.454.000 (hai cửa sổ cạnh nhau) | `Demo/B3_Demo_Refresh.xlsx` · PV_Table, PV_Vung · A3:B7 | 12 |
| 5 | S12b_change_source.png | Hộp Change PivotTable Data Source: `Sales_Vung!$A$1:$I$6198` | `Demo/B3_Demo_Refresh.xlsx` · PV_Vung | 12 |
| 6 | S14_value_field_settings.png | Hộp Value Field Settings (Custom Name: Doanh thu) và hộp Format Cells → Number, Use 1000 Separator | `Demo/B3_Demo_Pivot.xlsx` · PV_MienThang | 14 |
| 7 | S15a_compact.png | Region + Store_Code trên hàng, Compact Form | `Demo/B3_Demo_Pivot.xlsx` · PV_MienThang · A3:F20 | 15 |
| 8 | S15b_outline.png | Cùng Pivot, Outline Form | `Demo/B3_Demo_Pivot.xlsx` · PV_MienThang · A3:G20 | 15 |
| 9 | S15c_tabular.png | Cùng Pivot, Tabular Form + Repeat All Item Labels | `Demo/B3_Demo_Pivot.xlsx` · PV_MienThang · A3:G20 | 15 |
| 10 | S18a_sort_thu.png | Pivot Thu × Doanh thu sắp xếp giảm dần: Sun 159.845.000 … Thu 111.408.000 | `Practice/B3_MCI_Mart_Q1_2025_Answer.xlsx` · PV_Thu · A4:B12 | 18 |
| 11 | S18b_top5.png | Hộp Top 10 Filter (Top 5 Items by Sum of Revenue) và kết quả 5 sản phẩm, Hộp quà Tết cao cấp 135.850.000 đứng đầu | `Demo/B3_Demo_Pivot.xlsx` · PV_Top5 (sheet vừa tạo) · A3:B9 | 18 |
| 12 | S19a_group_7days.png | Hộp Grouping (Days, Number of days: 7) và Pivot 13 tuần (08–14/01 118.778.000; 29/01–04/02 31.567.000) | `Demo/B3_Demo_Pivot.xlsx` · PV_Tuan (sheet vừa tạo) · A3:B17 | 19 |
| 13 | S19b_cannot_group.png | Thông báo "Cannot group that selection." và 3 mục ngày dạng chữ ở đầu danh sách | `Demo/B3_Demo_Loi_Du_lieu.xlsx` · Pivot mới từ tblSalesLoi · Order_Date ở Rows | 19 |
| 14 | S20a_store_month.png | Pivot Store_Code × Thang, ô ST07 × 3 = 8.539.000 được chọn | `Practice/B3_MCI_Mart_Q1_2025_Answer.xlsx` · PV_CuaHang · A4:E18 | 20 |
| 15 | S20b_show_details.png | Sheet chi tiết sau khi double-click ô D12 (ST07 × 3), sắp xếp Order_Date: 55 dòng, từ 10/03 thưa hẳn | `Practice/B3_MCI_Mart_Q1_2025_Answer.xlsx` · sheet mới sau double-click PV_CuaHang!D12 · A1:J30 | 20 |
| 16 | S21a_slicer_connections.png | Slicer Store_Code (2 cột) cạnh Pivot Miền × Tháng và hộp Report Connections tick 3 Pivot | `Practice/B3_MCI_Mart_Q1_2025_Answer.xlsx` · PV_MienThang · A4:I20 | 21 |
| 17 | S21b_timeline.png | Timeline Order_Date (mức Months) phía trên PivotChart | `Practice/B3_MCI_Mart_Q1_2025_Answer.xlsx` · PV_Ngay · D4:N11 | 21 |
| 18 | S23_count_default.png | Pivot từ tblSalesLoi, Region × **Count of Revenue** (tổng 6.186) | `Demo/B3_Demo_Loi_Du_lieu.xlsx` · Pivot mới · A3:B7 + khung Values | 23 |
| 19 | S24_pct_column.png | Category × Thang, % of Column Total (Bánh kẹo 45,6% → 30,9% → 25,5%) | `Practice/B3_MCI_Mart_Q1_2025_Answer.xlsx` · PV_NganhHang · A4:E12 | 24 |
| 20 | S25_mom_running.png | Thang × Doanh thu, MoM %, Lũy kế, Số đơn, AOV, Gross Margin % | `Practice/B3_MCI_Mart_Q1_2025_Answer.xlsx` · PV_Thang · A4:G9 | 25 |
| 21 | S26_calc_field.png | Hộp Insert Calculated Field: Name AOV, Formula `=Revenue/So_don` | `Demo/B3_Demo_Pivot.xlsx` · PV_Thang (sau khi thêm cột So_don) | 26 |
| 22 | S28a_pie_wrong.png | PivotChart Pie doanh thu 3 tháng | `Demo/B3_Demo_Pivot.xlsx` · PV_Thang | 28 |
| 23 | S28b_line_right.png | Cùng dữ liệu, PivotChart Line | `Demo/B3_Demo_Pivot.xlsx` · PV_Thang | 28 |
| 24 | S29_pivotchart_ngay.png | PivotChart đường doanh thu theo ngày Quý 1 và Timeline | `Practice/B3_MCI_Mart_Q1_2025_Answer.xlsx` · PV_Ngay · D4:P33 | 29 |

Ảnh nên chụp thêm khi có thời gian (chưa có khung trên deck): hai cửa sổ Pivot và KPI_SUMIFS cạnh nhau bằng New Window (slide 11); Slicer Region nối Pivot Thu và Pivot cửa hàng (slide 21).

### Còn cần làm trước khi dạy
- Gửi trong nhóm chat trước buổi: file khởi đầu (`../B2_Ham_va_KPI/Practice/B2_MCI_Mart_Q1_2025_Answer.xlsx`); trong giờ nghỉ gửi `Demo/B3_Demo_Loi_Du_lieu_HocVien.xlsx`.
- Ảnh minh họa (xong 28/09/2026): 24 ảnh deck + 8 ảnh handout trong `Materials/B3/Images/`; đã dựng lại `Slide_B3.pptx`, `Slide_B3_HocVien.pptx`, `B3_Handout.docx` và xuất lại `B3_Handout.pdf` (không còn khung [CHỤP]).
- Chạy thử trên Excel 365 Windows (28/09/2026): `Practice/B3_MCI_Mart_Q1_2025_Answer.xlsx` mở được, Refresh khi mở, Kiem_tra 14/14 ✔, Slicer và Timeline chạy. **Còn sửa trong `build_b3.py`:** Slicer Store_Code quá hẹp (rộng 62 pt, nút hiện "…"); khối đối soát ở PV_MienThang bị cắt chữ cột A–B vì Pivot tự co cột khi Refresh (tắt Autofit column widths on update hoặc đặt khối đối soát sang cột khác).
- Slide 19, ngày dạng chữ (đã sửa 29/09/2026): trên Excel 365 Windows, 3 mục ngày dạng chữ nằm ở **đầu** danh sách Pivot (11/01/2025, 15/01/2025, 23/03/2025), còn trong ô lọc (AutoFilter) chúng nằm cuối danh sách. Đã sửa lời giảng slide 19, sheet `Huong_dan` (ô C7) của `Demo/B3_Demo_Loi_Du_lieu.xlsx` và bản `_HocVien` (sửa bằng Excel), `build_b3.py` và `_DANH_SACH_ANH.md`; đã dựng lại deck và bản học viên.

### Ghi chú so với Course Plan

- Course Plan (B3, câu chuyện #1) và `2_Dataset/README.md` viết "Bánh kẹo và Đồ uống tăng tỷ trọng". Tính trên dataset: Bánh kẹo tăng rõ (50,8% trong 20 ngày trước Tết so với 27,9% ngày thường; 45,6% tháng 1 so với 30,9% tháng 2), còn **Đồ uống tăng doanh thu (khoảng 1,8 lần mỗi ngày) nhưng tỷ trọng gần như không đổi** (11,7% so với 11,5%). Kịch bản nói đúng theo dữ liệu và dùng chính điểm này làm câu hỏi cho lớp (slide 24); đã sửa câu này trong Course Plan và README dataset.
- Course Plan (câu chuyện #2) viết "Double-click (Show Details) để thấy từ ngày 10/03". Dữ liệu ST07 sau 10/03 không về 0 (cửa hàng vẫn bán một phần diện tích): 15 đơn trong 22 ngày, 12 ngày không có đơn. Kịch bản mô tả đúng như vậy. ST04 cũng giảm 50,5% trong tháng 3 (do ngẫu nhiên của dữ liệu); kịch bản dùng điểm này cho bài học "drill-down trước khi kết luận".
- Nhóm theo tuần dùng Group → Days → 7 ngày tính từ 01/01/2025 (thứ Tư), nên tuần thứ 5 bắt đầu đúng mùng 1 Tết (29/01).
- Trình tự lỗi thường gặp (Course Plan, Hướng dẫn giảng dạy B3) được dạy ở: nguồn không phải Table và quên Refresh (slide 12), Count thay vì Sum (slide 23), không group được theo ngày (slide 19), GETPIVOTDATA (slide 11), Pie cho xu hướng (slide 28).
