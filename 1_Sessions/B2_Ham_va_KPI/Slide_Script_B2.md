# Kịch bản slide — Buổi 2: Hàm cho phân tích: từ dữ liệu giao dịch đến KPI

> **Phiên bản 1.0 (2026-09-27):** soạn theo `Course_Plan.md` v3.3, khuôn mẫu `Slide_Script_B1.md`. Lời giảng viết thành **văn nói đầy đủ**, số liệu tính từ dataset (script `2_Dataset/_generator/build_b2.py`, đã mở bằng Excel để kiểm tra).
> Căn cứ: `Course_Plan.md` v3.3 (B2: timetable 150 phút, Hướng dẫn giảng dạy B2, Nguyên tắc 1/7/9/10, Phụ lục A chuỗi B và C, Phụ lục C KPI dictionary, Phụ lục F câu chuyện #1).

## Cách dùng kịch bản

- **Lời giảng** là lời nói mẫu. Giảng viên không cần đọc nguyên văn; nên nói theo giọng của mình nhưng giữ đủ **ý chính** và **câu hỏi cho lớp** (in đậm).
- Dòng *(→ …)* trong lời giảng là **chỉ dẫn hành động**. Đường dẫn file tính từ thư mục buổi này (`1_Sessions/B2_Ham_va_KPI/`).
- Mỗi slide có thời lượng dự kiến. Nếu lớp chậm, cắt ở những chỗ đánh dấu **[có thể rút gọn]**.
- Máy cài định dạng vùng Việt Nam dùng dấu chấm phẩy `;` giữa các đối số. Slide và lời giảng viết theo dấu phẩy `,`; nhắc học viên đổi dấu khi chép công thức (đã nói ở B1, slide 18).

### Điều chỉnh cho lớp đã biết Excel sơ bộ

Theo Course Plan: rút Thống kê mô tả từ 15 xuống 5 phút (chỉ giữ mean vs median và outlier), dùng 10 phút cho PRO TIP Dynamic Arrays. Tổng vẫn 150 phút.

| Slide | Bản chuẩn | Lớp đã biết Excel | Cách rút / cách dùng thêm |
|---|---|---|---|
| 11–13 Thống kê mô tả | 14.5' | 4.5' | Slide 11 chỉ so AVERAGE với MEDIAN; slide 12 chỉ nêu ngưỡng outlier và số 138 đơn; bỏ slide 13 (RANK chuyển sang handout) |
| Sau slide 17 | — | +10' | PRO TIP Dynamic Arrays: mở `Demo/B2_Demo_SUMIFS.xlsx`, sheet Demo_DynamicArray; gõ `=SORT(UNIQUE(tblSales[Store_Code]))`, `=FILTER(…)`, rồi `=SUMIFS(tblSales[Revenue],tblSales[Store_Code],A4#)`. Kết quả mong đợi ở sheet Huong_dan |

## Quy ước

Loại slide và nguồn ảnh giống `Slide_Script_B1.md`: BÌA · CHIA PHẦN · KHÁI NIỆM · CÔNG CỤ · THỰC HÀNH · QUIZ · TỔNG KẾT; **[CHỤP]** ảnh chụp Excel 365 Windows trên file case, **[PDF pX]** trang PDF nguồn, **[VẼ]** vẽ bằng shape, chỉ cho sơ đồ khái niệm và thẻ số (theo `CLAUDE.md`, mục Ảnh minh họa: thao tác và kết quả Excel phải là ảnh chụp thật).

**File dùng trong buổi** (đường dẫn tính từ thư mục buổi; bản phát cho học viên có hậu tố `_HocVien`):

| File | Công dụng | Slide |
|---|---|---|
| `../B1_Nen_tang_Excel_Du_lieu/Practice/B1_MCI_Mart_Q1_2025_Answer.xlsx` | **File khởi đầu** (đáp án B1): tblSales 6.187 dòng. Phát đầu buổi | 3, 4, 29 |
| `Practice/B2_Product_DuplicateKey.xlsx` | Bảng Product 41 dòng, **P021 trùng khóa** (dòng 22 là bản ghi cũ, Unit_Cost 62.000). Phát cho học viên | 23 (nhắc), 29 |
| `Practice/B2_MCI_Mart_Q1_2025_Answer.xlsx` | Đáp án B2: tblSales làm giàu, tblProduct sạch, KPI_SUMIFS, KPI_Card, Kiem_tra (10/10 ✔). **File khởi đầu B3** | 31 |
| `Handout/B2_KPI_Definition_Card.docx` / `.pdf` | Mẫu KPI Definition Card 1 trang (in sẵn 3 tờ / học viên) | 5, 30 |
| `Demo/B2_Demo_Logic.xlsx` | Demo IF, IFS, AND/OR/NOT, IFERROR trên bảng 40 sản phẩm + 2 sản phẩm chưa có giá | 7, 8, 9 |
| `Demo/B2_Demo_ThongKe.xlsx` | Demo thống kê giá trị đơn (3.144 đơn), outlier; RANK 12 cửa hàng | 11, 12, 13 |
| `Demo/B2_Demo_SUMIFS.xlsx` | Demo SUMIFS/COUNTIFS/AVERAGEIFS, bảng 2 chiều, trung bình của trung bình; PRO TIP Dynamic Arrays | 15, 16, 17 |
| `Demo/B2_Demo_Lookup.xlsx` | Demo VLOOKUP, INDEX/MATCH, XLOOKUP, khóa trùng P021, khóa khác kiểu | 20, 21, 22, 23, 24 |
| `Demo/B2_Demo_Text_Date.xlsx` | Demo LEFT/MID/RIGHT, TRIM, &, TEXT; EOMONTH, WEEKDAY, NETWORKDAYS | 26, 27 |
| `Homework/B2_BTVN_MCI_Mart_T4_2025.xlsx` | Bài tập về nhà bắt buộc (đề), nộp trước Buổi 3 | 41 |
| `Homework/B2_BTVN_MCI_Mart_T4_2025_Answer.xlsx` | Đáp án bài về nhà (Kiem_tra 8/8 ✔, Dap_an_loi, Insight_mau). Chỉ giảng viên / trợ giảng | — |
| `Handout/B2_Handout.docx` / `.pdf` | Handout tự học (hàm bổ sung, PRO TIP, bài về nhà, checklist trước B3) | 41 |
| `Handout/B2_Infographic_TomTat.png` | Infographic tóm tắt buổi, phát riêng (không chèn vào handout); sửa ở `2_Dataset/_generator/infographic_b2.html` rồi chạy `build_infographic.py B2` | 41 |

Mọi file demo có sheet `Huong_dan` (Slide · Thao tác · Kết quả mong đợi). Bản giảng viên mở ra ở sheet làm việc, `Huong_dan` để cuối.

**Số liệu đáp án B2** (Quý 1/2025, dùng khi chữa bài):

| Chỉ tiêu | Giá trị |
|---|---|
| Số dòng / số đơn | **6.187 dòng / 3.144 đơn** |
| Doanh thu | **941.454.000** |
| COGS | **619.663.000** (nếu còn P021 cũ: 617.871.000) |
| Gross Margin % | **34,2%** (nếu còn P021 cũ: 34,4%) |
| AOV | **299.445** |

| Miền × Tháng (doanh thu) | Tháng 1 | Tháng 2 | Tháng 3 | Quý 1 |
|---|---|---|---|---|
| Bắc | 145.814.000 | 87.954.000 | 90.210.000 | 323.978.000 |
| Trung | 85.599.000 | 54.254.000 | 49.960.000 | 189.813.000 |
| Nam | 194.437.000 | 103.788.000 | 129.438.000 | 427.663.000 |
| **Tổng** | **425.850.000** | **245.996.000** | **269.608.000** | **941.454.000** |
| Số đơn | 1.398 | 823 | 923 | 3.144 |
| AOV | 304.614 | 298.902 | 292.100 | 299.445 |
| Gross Margin % | 35,2% | 33,7% | 33,1% | 34,2% |

**Tổng:** 41 slide, 150 phút.

---

## Phần 0 — Mở đầu: từ câu hỏi đến KPI (10 phút · slide 1–5)

### Slide 1 · BÌA · 0.5'
- **Trên slide:** MCI / Consulting & Analytics / **Excel for Business Intelligence Analyst** / Buổi 2: Hàm cho phân tích — từ dữ liệu giao dịch đến KPI
- **Ảnh:** [VẼ] nền theo Style spec.

**Lời giảng:**
> Chào cả lớp, chào mừng các bạn quay lại Buổi 2. Buổi trước chúng ta mới dọn dữ liệu cho sạch; hôm nay mình bắt đầu **tính ra con số để trả lời sếp**.

### Slide 2 · KHÁI NIỆM · Câu hỏi mở buổi · 1.5'
- **Trên slide:** *"Doanh thu, số đơn, biên lợi nhuận của từng vùng tháng này là bao nhiêu?"* Góc slide có nhãn **Tầng 1 — Descriptive**.
- Dưới câu hỏi: ba từ khóa đóng khung: **Doanh thu** · **Số đơn** · **Biên lợi nhuận**, và dòng nhỏ *"… của từng vùng … tháng này"*.
- **Ảnh:** [VẼ] bảng trống 3 miền × 3 KPI, các ô là dấu "?".

**Lời giảng:**
> Đây là câu hỏi Giám đốc bán hàng gửi cho chúng ta tuần này. Nghe rất đơn giản, đúng không? Nhưng mình hỏi các bạn một câu: **"số đơn" là đếm cái gì?** Đếm số dòng trong file, hay đếm cái gì khác?
>
> *(→ chờ 2–3 câu trả lời, KHÔNG chốt đáp án)*
>
> Các bạn thấy chưa, chỉ một chữ "số đơn" mà đã có nhiều cách hiểu. Còn "biên lợi nhuận" thì file hôm qua chưa có giá vốn. Vậy trước khi viết công thức nào, mình phải **định nghĩa KPI cho rõ**. Đó là việc đầu tiên hôm nay.

### Slide 3 · KHÁI NIỆM · Recap B1 & file khởi đầu · 1.5'
- **Trên slide:**
  - Viên gạch B1: `tblSales` · 6.187 dòng · 941.454.000 · không dòng tổng, không ô gộp.
  - Hôm nay nhận **file khởi đầu** = đáp án B1 (`B1_MCI_Mart_Q1_2025_Answer.xlsx`). Bài về nhà B1 chữa ở slide chữa bài và trong nhóm chat.
  - Hai hàm cần nhớ từ B1: `$` (F4) và `[@Cột]` trong Table.
- **Ảnh:** [VẼ] viên gạch B1 (đã xong) nằm dưới viên gạch B2 (sáng lên).

**Lời giảng:**
> Nhắc nhanh buổi trước. Viên gạch B1 là bảng tblSales: 6.187 dòng, tổng doanh thu 941 triệu 454 nghìn, không còn dòng tổng hay ô gộp.
>
> *(→ demo: mở `../B1_Nen_tang_Excel_Du_lieu/Practice/B1_MCI_Mart_Q1_2025_Answer.xlsx`, sheet Kiem_tra: 8 dòng ✔)*
>
> Đây là file khởi đầu hôm nay, trợ giảng đã gửi trong nhóm chat. Dù file của bạn đã khớp hay chưa, hôm nay **mọi người làm trên file này**, để cả lớp cùng một nền. Các bạn mở nó ra và lưu thành tên mình, ví dụ `B2_NguyenVanA.xlsx`.
>
> Hai thứ từ buổi trước sẽ dùng liên tục hôm nay: **dấu đô-la** và **công thức theo tên cột trong Table**. Ai còn lúng túng với F4 thì để ý kỹ ở phần SUMIFS.

### Slide 4 · KHÁI NIỆM · Dimension / Measure / Grain · 3'
- **Trên slide:** [VẼ] đầu bảng tblSales, tô màu theo vai trò:

| Vai trò | Câu hỏi | Cột trong tblSales |
|---|---|---|
| **Dimension** | "Theo cái gì?" | Region, Store_Code, Order_Date (→ Tháng), Product_Code (→ Category) |
| **Measure** | "Đo cái gì?" | Quantity, Revenue (→ COGS, Gross Margin) |
| **Grain** | "Mỗi dòng là gì?" | **1 dòng = 1 sản phẩm trong 1 đơn** (Order_ID × Line_No) |

- **Ảnh:** [CHỤP] đơn HD2501-00002 có 4 dòng (Line_No 1–4): `../B1_Nen_tang_Excel_Du_lieu/Practice/B1_MCI_Mart_Q1_2025_Answer.xlsx`, sheet Sales, A1:K7, chọn dòng 3–6 (ẩn tạm cột D–F, I–J).
- **Hộp nhấn:** 6.187 dòng ≠ 3.144 đơn. Đếm dòng không phải đếm đơn.

**Lời giảng:**
> Ba từ này các bạn sẽ gặp đến hết khóa, nên mình giới thiệu ngay.
>
> **Dimension** là "theo cái gì": theo vùng, theo cửa hàng, theo tháng. **Measure** là "đo cái gì": số lượng, doanh thu. Câu hỏi của sếp có cả hai: "doanh thu" là measure, "từng vùng", "tháng này" là dimension.
>
> Từ thứ ba quan trọng nhất: **Grain**, tức là *mỗi dòng trong bảng là gì*. *(→ mở `../B1_Nen_tang_Excel_Du_lieu/Practice/B1_MCI_Mart_Q1_2025_Answer.xlsx`, sheet Sales, chỉ vào đơn HD2501-00002 ở dòng 3–6)* Các bạn nhìn đơn này: một đơn hàng nhưng có bốn dòng, vì khách mua bốn sản phẩm. Vậy grain của tblSales là **một sản phẩm trong một đơn**.
>
> Hiểu grain thì trả lời được câu hỏi lúc nãy: bảng có 6.187 dòng nhưng chỉ có **3.144 đơn**. Ai đếm số dòng rồi báo "Quý 1 có 6.187 đơn" là báo sai gần gấp đôi. Lát nữa mình sẽ đếm đơn đúng cách.

### Slide 5 · KHÁI NIỆM · KPI Definition Card · 3.5'
- **Trên slide:** mẫu card 10 trường (Phụ lục C), điền sẵn card **Total Revenue**:

| Trường | Total Revenue |
|---|---|
| Mục đích | Tháng này bán được bao nhiêu tiền, ở đâu? |
| Công thức | SUM(Quantity × Unit_Price) |
| Nguồn dữ liệu | tblSales |
| Grain | Dữ liệu: 1 sản phẩm trong 1 đơn · KPI: tháng |
| Dimensions | Miền, cửa hàng, ngành hàng, tháng |
| Owner | Giám đốc bán hàng |
| Target / Ngưỡng | Theo bảng Target (từ B4) · Warning 90%, Critical 80% |
| Ghi chú giới hạn | Chưa trừ trả hàng, khuyến mại |

- **TIP:** Viết card **trước** khi viết công thức. Card là "hợp đồng" giữa bạn và người đọc báo cáo.
- **Tài liệu:** `Handout/B2_KPI_Definition_Card.docx` (phát bản in).

**Lời giảng:**
> Trong công ty, cùng một chữ "doanh thu" mà phòng Kinh doanh, phòng Tài chính tính ra hai con số khác nhau là chuyện rất thường gặp. Cách chữa là **KPI Definition Card**: một tờ ghi rõ KPI tên gì, tính thế nào, lấy dữ liệu ở đâu, theo chiều nào, ai chịu trách nhiệm.
>
> *(→ phát mẫu `Handout/B2_KPI_Definition_Card.docx` bản in; cùng lớp điền card Total Revenue theo slide)*
>
> Mình điền cùng các bạn card đầu tiên. Tên: Total Revenue. Mục đích: trả lời câu hỏi "bán được bao nhiêu tiền". Công thức: tổng số lượng nhân đơn giá. Để ý dòng **Grain**: dữ liệu ở mức từng sản phẩm trong đơn, nhưng KPI báo cáo theo tháng. Dòng **Ghi chú giới hạn** là chỗ các bạn thành thật với người đọc: con số này chưa trừ hàng trả lại.
>
> Hai card còn lại, **Total Orders** và **Gross Margin %**, các bạn tự viết trong phần thực hành. Quy tắc của khóa: **viết card trước, viết công thức sau.**

---

## Phần 1 — Hàm logic (15 phút · slide 6–9)

### Slide 6 · CHIA PHẦN · 0.5'
- **Trên slide:** MCI / **1. Hàm logic: dạy Excel ra quyết định**

**Lời giảng:**
> Phần 1: dạy Excel tự ra quyết định theo quy tắc kinh doanh, bằng IF và các hàm đi cùng nó.

### Slide 7 · CÔNG CỤ · IF, IFS · 6'

| Tên (đỏ) | Mục đích |
|---|---|
| **IF** | Một điều kiện, hai kết quả: `=IF(F2>=35%,"Biên cao","Biên thấp")` |
| **IFS** | Nhiều điều kiện xét lần lượt, gặp điều kiện đúng đầu tiên thì dừng: `=IFS(D2>=100000,"Cao",D2>=30000,"Trung",TRUE,"Thấp")` |

- **Ảnh:** [CHỤP] cột Margin_%, IF (G), IFS (H) sau khi kéo; P041 ra #DIV/0! và "Thấp": `Demo/B2_Demo_Logic.xlsx`, sheet Demo_IF, A1:H6 và dòng 42–43 (ẩn tạm dòng 7–41); [PDF p17–20].
- **TIP:** IFS xét **từ trái sang phải**: đặt ngưỡng cao trước. `TRUE` ở cuối là "các trường hợp còn lại".

**Lời giảng:**
> *(→ demo: mở `Demo/B2_Demo_Logic.xlsx`, sheet Demo_IF; các bước ở sheet Huong_dan)* Đây là danh mục 40 sản phẩm của MCI Mart, cộng 2 sản phẩm mới ở cuối chưa có giá bán. Cột F là biên lợi nhuận của từng sản phẩm: giá bán trừ giá vốn, chia giá bán.
>
> Phòng mua hàng muốn chia sản phẩm thành "biên cao" và "biên thấp" với ngưỡng 35%. Một điều kiện, hai kết quả: đó là **IF**. *(→ chọn G2:G43, Delete; gõ G2: `=IF(F2>=35%,"Biên cao","Biên thấp")`, kéo xuống)* Cấu trúc đọc như tiếng Việt: **nếu** biên từ 35% trở lên **thì** "Biên cao", **ngược lại** "Biên thấp". Kết quả: **16 sản phẩm biên cao, 24 biên thấp**.
>
> Khi có nhiều hơn hai nhóm thì sao? Phân khúc giá: từ 100 nghìn trở lên là "Cao", từ 30 nghìn là "Trung", còn lại "Thấp". Ngày trước phải lồng IF trong IF rất rối; giờ có **IFS**. *(→ H2: `=IFS(D2>=100000,"Cao",D2>=30000,"Trung",TRUE,"Thấp")`, kéo xuống)* 12 sản phẩm phân khúc Cao.
>
> **Câu hỏi cho lớp: nếu mình đảo thứ tự, đặt điều kiện "≥ 30.000" lên trước thì sản phẩm giá 135 nghìn ra gì?** *(→ chờ)* Ra "Trung", vì IFS dừng ở điều kiện đúng **đầu tiên**. Đặt ngưỡng cao trước.
>
> Còn một điều các bạn nhìn xuống cuối bảng: hai sản phẩm **chưa có giá** mà IFS vẫn xếp vào "Thấp". Excel coi ô trống là số 0. Công thức không sai cú pháp, nhưng **kết luận thì sai**. Nhớ ví dụ này, lát nữa gặp lại.

### Slide 8 · CÔNG CỤ · AND, OR, NOT · 4'

| Tên (đỏ) | Mục đích |
|---|---|
| **AND** | TRUE khi **mọi** điều kiện đúng |
| **OR** | TRUE khi **ít nhất một** điều kiện đúng |
| **NOT** | Đảo TRUE ↔ FALSE |

- **Ảnh:** thanh công thức `=AND(C2="Bánh kẹo",F2>=35%)` · `=OR(C2="Sữa",C2="Thực phẩm khô")` · `=NOT(J2)` và thẻ số 7 / 13 / 29; [CHỤP] cột I, J, K ra TRUE/FALSE trên các dòng Sữa và Bánh kẹo: `Demo/B2_Demo_Logic.xlsx`, sheet Demo_IF, A1:K1 + dòng 9–13 và 22–28; [PDF p21–22].
- **TIP:** Thường lồng vào IF: `=IF(AND(…),"Ưu tiên","")`.

**Lời giảng:**
> Trước Tết, phòng mua hàng muốn đánh dấu sản phẩm **ưu tiên nhập**: phải là Bánh kẹo **và** biên từ 35% trở lên. Hai điều kiện cùng đúng: dùng **AND**. *(→ trên `Demo/B2_Demo_Logic.xlsx`, sheet Demo_IF: I2 `=AND(C2="Bánh kẹo",F2>=35%)`, kéo xuống)* Có **7 sản phẩm** thỏa cả hai.
>
> Nhóm "hàng thiết yếu" thì khác: Sữa **hoặc** Thực phẩm khô, thuộc một trong hai là được. Đó là **OR**. *(→ J2 `=OR(C2="Sữa",C2="Thực phẩm khô")`)* 13 sản phẩm. Muốn lấy nhóm ngược lại, "không thiết yếu", thì bọc **NOT** bên ngoài. *(→ K2 `=NOT(J2)`)*
>
> Ba hàm này trả về TRUE/FALSE, nên thường được lồng vào IF để ra nhãn dễ đọc. Mẹo nhớ: **AND khó tính, OR dễ tính.**

### Slide 9 · CÔNG CỤ · IFERROR · 4.5'
- **Tên (đỏ):** **IFERROR**
- **Mục đích:** Nếu công thức ra lỗi thì trả về giá trị thay thế: `=IFERROR((D2-E2)/D2,"Chưa có giá")`.
- **Ảnh:** [CHỤP] dòng P041, P042: Margin_% ra `#DIV/0!`, cột L ra "Chưa có giá"; L42 sửa thành `IFERROR(…,0)` ra 0%: `Demo/B2_Demo_Logic.xlsx`, sheet Demo_IF, A41:L43.
- **PRO TIP:** IFERROR **che** lỗi chứ không **sửa** lỗi. Thay bằng chữ có nghĩa, đừng thay bằng 0.

**Lời giảng:**
> Hai sản phẩm mới chưa có giá, nên cột biên lợi nhuận ra **#DIV/0!**, chia cho 0. Báo cáo gửi sếp đầy lỗi thì không đẹp. **IFERROR** cho phép nói: "nếu ra lỗi thì hiện cái này thay thế". *(→ trên `Demo/B2_Demo_Logic.xlsx`, sheet Demo_IF: L2 `=IFERROR((D2-E2)/D2,"Chưa có giá")`, kéo xuống; P041, P042 ở dòng 42–43)*
>
> Nhưng mình muốn các bạn cẩn thận. Rất nhiều người viết IFERROR **trả về 0**. *(→ sửa L42 thành `=IFERROR((D42-E42)/D42,0)`)* Các bạn thấy không: bây giờ sản phẩm này hiện biên **0%**. Người đọc sẽ nghĩ "sản phẩm này bán hòa vốn", trong khi sự thật là **chưa có dữ liệu**. Một con số sai trông như thật còn nguy hiểm hơn một ô báo lỗi.
>
> *(→ Ctrl+Z; gõ giá 25000 vào D42)* Khi có giá, mọi cột tự tính lại: biên 44%. Quy tắc: **IFERROR để làm sạch hiển thị, không phải để giấu vấn đề.** Đây là bài học D2 hôm nay: số hiển thị chưa chắc đúng.

---

## Phần 2 — Thống kê mô tả (15 phút · slide 10–13)

### Slide 10 · CHIA PHẦN · 0.5'
- **Trên slide:** MCI / **2. Thống kê mô tả: một con số "trung bình" có thể nói dối**

**Lời giảng:**
> Phần 2 ngắn nhưng rất quan trọng: chọn đúng con số để mô tả dữ liệu.

### Slide 11 · CÔNG CỤ · Thống kê mô tả: trung bình hay trung vị? (COUNT, AVERAGE vs MEDIAN, MIN/MAX) · 5.5'

| Tên (đỏ) | Mục đích |
|---|---|
| **COUNT / COUNTA** | Đếm ô số / ô không trống |
| **AVERAGE vs MEDIAN** | Trung bình cộng vs giá trị ở giữa khi xếp thứ tự |
| **MIN / MAX** | Nhỏ nhất / lớn nhất |

- **Ảnh:** [VẼ] thẻ số: **Trung bình 299.445** · **Trung vị 198.000** · Nhỏ nhất 10.000 · Lớn nhất 2.464.000 (3.144 đơn Q1/2025); [CHỤP] khối I2:I6 và Status Bar khi chọn cột E: `Demo/B2_Demo_ThongKe.xlsx`, sheet Demo_ThongKe, E1:I13.
- **TIP:** Trung bình lớn hơn trung vị nhiều → có vài giá trị rất lớn kéo lên (lệch phải).

**Lời giảng:**
> *(→ demo: mở `Demo/B2_Demo_ThongKe.xlsx`, sheet Demo_ThongKe; mỗi dòng là một đơn hàng Q1, cột E là giá trị đơn)* Sếp hỏi: "Một khách vào MCI Mart thường chi bao nhiêu?" Phản xạ đầu tiên là lấy **AVERAGE**: 299 nghìn.
>
> Giờ mình tính thêm **MEDIAN**, trung vị: xếp 3.144 đơn từ nhỏ đến lớn và lấy đơn đứng giữa. *(→ chỉ ô I4)* **198 nghìn.** Chênh nhau hơn 50%!
>
> **Theo các bạn, con số nào đúng hơn cho câu "khách thường chi bao nhiêu"?** *(→ chờ)* Trung vị. Một nửa số đơn dưới 198 nghìn. Trung bình bị kéo lên bởi một số đơn rất lớn: đơn lớn nhất là 2 triệu 464 nghìn, trong khi đơn nhỏ nhất chỉ 10 nghìn.
>
> Nhưng lưu ý: KPI **Avg Order Value** trong dictionary vẫn dùng trung bình, vì nó nhân ngược được ra doanh thu: AOV × số đơn = doanh thu. Không có con số nào "đúng" tuyệt đối; **chọn con số theo câu hỏi**, và khi báo cáo trung bình thì nhìn thêm trung vị.

### Slide 12 · CÔNG CỤ · STDEV, skew, outlier · 5'

| Tên (đỏ) | Mục đích |
|---|---|
| **STDEV.S** | Độ phân tán quanh trung bình |
| **SKEW** | Độ lệch: > 0 là đuôi dài bên phải |
| **QUARTILE.INC** | Q1, Q3 để tính ngưỡng outlier = Q3 + 1,5 × (Q3 − Q1) |

- **Ảnh:** [VẼ] trục số giá trị đơn: Q1 89.000 · trung vị 198.000 · Q3 412.000 · ngưỡng **896.500**; 138 chấm vượt ngưỡng; [CHỤP] khối I7:I13: `Demo/B2_Demo_ThongKe.xlsx`, sheet Demo_ThongKe, G7:I13; [Statistics PDF p24–38].
- **PRO TIP:** Outlier không phải "lỗi để xóa". **Giải thích trước, quyết định sau.**

**Lời giảng:**
> Độ lệch chuẩn của giá trị đơn là 290 nghìn, gần bằng chính trung bình: dữ liệu **phân tán rất rộng**. SKEW bằng 1,77, dương, nghĩa là đuôi dài bên phải: nhiều đơn nhỏ, ít đơn rất to.
>
> Vậy đơn nào là "bất thường"? Cách phổ biến là quy tắc IQR: lấy tứ phân vị thứ ba cộng 1,5 lần khoảng tứ phân vị. *(→ trên `Demo/B2_Demo_ThongKe.xlsx`, sheet Demo_ThongKe, chỉ I9:I13)* Ngưỡng là 896.500 đồng, và có **138 đơn** vượt ngưỡng.
>
> **138 đơn này có nên xóa không?** *(→ chờ)* Không! Đây có thể là khách mua hộp quà Tết, mua sỉ; đúng là khách thật và doanh thu thật. Việc của người làm BI là **giải thích outlier trước khi báo cáo**: chúng đến từ đâu, có phải lỗi nhập liệu không. Xóa đi thì trung bình đẹp hơn nhưng báo cáo sai.
>
> *(→ [có thể rút gọn] sort cột E giảm dần để thấy đơn lớn nhất HD2502-00141, rồi Ctrl+Z)*

### Slide 13 · CÔNG CỤ · RANK.EQ · 4'
- **Tên (đỏ):** **RANK.EQ**
- **Mục đích:** Xếp hạng một giá trị trong danh sách: `=RANK.EQ(D2,$D$2:$D$13,0)` (0 = giảm dần).
- **Ảnh:** [CHỤP] cột E sau khi kéo RANK.EQ (hạng 1 ST08 Quận 1 130.448.000 … hạng 12 ST04 Hạ Long 45.105.000), F2 hiện vùng `$D$2:$D$13`: `Demo/B2_Demo_ThongKe.xlsx`, sheet Demo_RANK, A1:E13.
- **TIP:** Vùng xếp hạng phải khóa `$`, nếu không kéo xuống sẽ trôi.

**Lời giảng:**
> *(→ trên `Demo/B2_Demo_ThongKe.xlsx`, sheet Demo_RANK: E2 `=RANK.EQ(D2,$D$2:$D$13,0)`, kéo xuống)* Hàm cuối của phần này: **RANK.EQ**, xếp hạng cửa hàng theo doanh thu Quý 1. Đối số thứ hai là danh sách để so, và nó **phải khóa bằng đô-la**; các bạn thấy không, lại là bài học F4 hôm qua.
>
> Quận 1 đứng đầu với 130 triệu, Hạ Long đứng cuối với 45 triệu. Nhưng khoan kết luận Hạ Long "kém": một cửa hàng Mall ở trung tâm và một cửa hàng mặt phố ở tỉnh có quy mô khác nhau. Xếp hạng chỉ có nghĩa khi so **cùng loại**; câu hỏi "vì sao" thì để Buổi 3 trả lời.

---

## Phần 3 — Tổng hợp có điều kiện: COUNTIFS, SUMIFS, AVERAGEIFS (20 phút · slide 14–17)

### Slide 14 · CHIA PHẦN · 0.5'
- **Trên slide:** MCI / **3. SUMIFS: mỗi tiêu chí là một bộ lọc**

**Lời giảng:**
> Phần 3 là phần quan trọng nhất hôm nay. Khái niệm các bạn học ở đây sẽ quay lại ở Buổi 3 và Buổi 5 dưới hai cái tên khác.

### Slide 15 · KHÁI NIỆM · Xương sống Chuỗi C: bộ lọc + phép tổng · 5'
- **Trên slide:** [VẼ] hình xương sống (dùng lại nguyên vẹn ở B3, B5):
  - Bên phải: **một ô kết quả** = 90.210.000.
  - Hai "bộ lọc" đi vào ô: **[Vùng = Bắc]** và **[Tháng = 3]**.
  - Bên trái: cột Revenue → **phép tổng** (SUM).
  - Chú thích dưới: *B2 gọi là criteria của SUMIFS · B3 gọi là tọa độ ô Pivot · B5 gọi là filter context.*
- **TIP:** Kiểm chứng SUMIFS bằng Filter + Status Bar.

**Lời giảng:**
> Mình vẽ một hình và mong các bạn nhớ nó đến hết khóa. Bên phải là **một ô kết quả**: doanh thu Miền Bắc tháng 3. Để ra con số đó, Excel làm hai việc: **lọc** những dòng có Vùng = Bắc **và** Tháng = 3, rồi **cộng** cột Revenue của những dòng còn lại.
>
> Mỗi điều kiện là một **bộ lọc**. Hôm nay, trong SUMIFS, chúng ta gọi nó là *criteria*. Buổi 3, trong Pivot, nó là *tọa độ của ô*. Buổi 5, trong DAX, nó là *filter context*. **Cùng một ý, ba cái tên.** Hiểu nó hôm nay thì Buổi 5 nhẹ đi một nửa.
>
> *(→ demo: mở `Demo/B2_Demo_SUMIFS.xlsx`, sheet Sales; lọc Region = Bắc, Thang = 3; chọn cột Revenue, nhìn Status Bar)* Mình làm đúng việc đó bằng tay: lọc Bắc, lọc tháng 3, chọn cột Revenue. Status Bar: **90.210.000**, 635 dòng. *(→ Clear Filter; sang sheet Demo_SUMIFS, ô D4)* Và đây là SUMIFS làm cùng việc đó trong một công thức: cũng 90.210.000.

### Slide 16 · CÔNG CỤ · SUMIFS, COUNTIFS, AVERAGEIFS · 8.5'

| Tên (đỏ) | Mục đích | Ví dụ (Bắc, tháng 3) |
|---|---|---|
| **SUMIFS** | Tổng có điều kiện | `=SUMIFS(tblSales[Revenue],tblSales[Region],"Bắc",tblSales[Thang],3)` → 90.210.000 |
| **COUNTIFS** | Đếm có điều kiện | `=COUNTIFS(tblSales[Region],"Bắc",tblSales[Thang],3)` → 635 dòng; thêm `tblSales[Line_No],1` → **321 đơn** |
| **AVERAGEIFS** | Trung bình có điều kiện | `=AVERAGEIFS(tblSales[Revenue],…)` → 142.063 / dòng |

- **Ảnh:** dòng giải phẫu SUMIFS: ① vùng cần cộng (đứng **đầu**) ② cặp vùng tiêu chí – tiêu chí ③ cặp thứ hai…; [CHỤP] ScreenTip khi gõ `=SUMIFS(` trên tblSales: `Demo/B2_Demo_SUMIFS.xlsx`, sheet Demo_SUMIFS, ô trống; [PDF p33–34].
- **TIP:** Dùng cột của Table (`tblSales[Region]`) thì các vùng luôn cùng kích thước, không bị `#VALUE!`.
- **PRO TIP:** Đếm **đơn** trên bảng grain "dòng sản phẩm": COUNTIFS thêm điều kiện `Line_No = 1`.

**Lời giảng:**
> Cú pháp SUMIFS: đối số đầu tiên là **cột cần cộng**, sau đó là từng **cặp**: cột để lọc, giá trị lọc. Muốn thêm bộ lọc thì thêm cặp. *(→ trên `Demo/B2_Demo_SUMIFS.xlsx`, sheet Demo_SUMIFS: F2 vào D4, chỉ từng phần)* Ở đây giá trị lọc không gõ cứng "Bắc" mà trỏ vào ô B4, để mình đổi được.
>
> *(→ đổi B4:B7 thành Nam, C4:C7 thành 1)* Đổi thành Nam, tháng 1: 194 triệu 437 nghìn. Một công thức, trả lời được mọi cặp vùng – tháng.
>
> **COUNTIFS** cùng cấu trúc nhưng không có cột cần cộng, chỉ đếm. *(→ trả B4:B7 về Bắc, C4:C7 về 3; chỉ D5, D6)* Đếm Bắc tháng 3 ra 635. **Câu hỏi: 635 là số đơn phải không?** *(→ chờ)* Không, đó là số **dòng**, nhớ grain! Mỗi đơn có đúng một dòng Line_No bằng 1, nên thêm điều kiện Line_No = 1 là đếm được **đơn**: 321 đơn.
>
> **AVERAGEIFS** là trung bình có điều kiện: 142 nghìn mỗi dòng. Cẩn thận: đây là trung bình **mỗi dòng sản phẩm**, không phải giá trị đơn trung bình.
>
> Lỗi hay gặp: viết SUMIFS trên vùng thường mà vùng cộng và vùng tiêu chí lệch kích thước, Excel báo #VALUE!. Dùng cột của Table thì không bao giờ bị.

### Slide 17 · KHÁI NIỆM · Bảng KPI 2 chiều · tổng ÷ tổng ("trung bình của trung bình") · 6'
- **Trên slide:** trái: [CHỤP] bảng Miền × Tháng (3 × 3 + Tổng) sau khi kéo từ B12, F2 ở C13 hiện `$A13` và `C$11`, E15 = 941.454.000 (`Demo/B2_Demo_SUMIFS.xlsx`, sheet Demo_SUMIFS, A11:E15); phía trên là thanh công thức: `=SUMIFS(tblSales[Revenue],tblSales[Region],$A12,tblSales[Thang],B$11)` (tô `$A` và `B$`). Phải: thẻ so sánh:
  - AVERAGE Gross Margin % của 6 ngành = **31,2%** ✘
  - Tổng ÷ tổng = **34,2%** ✔
- **TIP:** KPI tỷ lệ (AOV, Margin %) luôn tính **tổng ÷ tổng**.

**Lời giảng:**
> Sếp không hỏi một ô mà hỏi cả bảng: 3 miền × 3 tháng. Viết 9 công thức SUMIFS? Không cần, **viết một công thức rồi kéo**, và đây là lúc dấu đô-la hôm qua phát huy. *(→ trên `Demo/B2_Demo_SUMIFS.xlsx`, sheet Demo_SUMIFS: xóa B12:D14, gõ B12 `=SUMIFS(tblSales[Revenue],tblSales[Region],$A12,tblSales[Thang],B$11)`, kéo sang D12 rồi xuống dòng 14)*
>
> Tên miền nằm ở **cột A**, nên khóa cột: đô-la A. Tháng nằm ở **dòng 11**, nên khóa dòng: B đô-la 11. Kéo ra 9 ô, và **đối soát**: ô tổng E15 phải bằng tổng cột Revenue, 941.454.000; ô E17 là hiệu hai số, bằng 0.
>
> Nhìn hàng Tổng: **tháng 1 là 425,9 triệu, tháng 2 chỉ còn 246 triệu.** Giảm hơn 40% chỉ sau một tháng. Mình ghi lại câu hỏi "vì sao?" lên góc bảng; hôm nay chưa trả lời, Buổi 3 sẽ trả lời.
>
> Một cái bẫy cuối. *(→ cuộn tới khối Gross Margin % theo ngành, dòng 20–28)* Muốn biên lợi nhuận toàn công ty, nhiều bạn lấy AVERAGE biên của 6 ngành: **31,2%**. Sai. Ngành Sữa rất nhỏ, chỉ 16 triệu doanh thu, nhưng được tính "một phiếu" ngang với Bánh kẹo 339 triệu. Cách đúng là **tổng lãi gộp chia tổng doanh thu: 34,2%**. Với mọi KPI dạng tỷ lệ: **tổng chia tổng, không lấy trung bình của trung bình.**

---

### Slide 18 · NGHỈ · 10'
- **Trên slide:** Nghỉ 10 phút. Trước khi quay lại: mở `B2_Product_DuplicateKey.xlsx`.

**Lời giảng:**
> Mình nghỉ 10 phút. Khi quay lại, các bạn mở sẵn file khởi đầu và file `Practice/B2_Product_DuplicateKey.xlsx`. Phần tiếp theo là nối bảng.
>
> *(→ trong giờ nghỉ: kiểm tra học viên đã mở được file khởi đầu, đổi tên file theo tên mình)*

---

## Phần 4 — Nối bảng: VLOOKUP → INDEX/MATCH → XLOOKUP (25 phút · slide 19–24)

### Slide 19 · CHIA PHẦN · 0.5'
- **Trên slide:** MCI / **4. Nối bảng: lấy thông tin từ bảng khác**

**Lời giảng:**
> tblSales có mã sản phẩm nhưng không có ngành hàng và giá vốn. Hai thứ đó nằm ở bảng Product. Phần 4 là cách nối hai bảng.

### Slide 20 · CÔNG CỤ · VLOOKUP · 5'
- **Tên (đỏ):** **VLOOKUP**
- **Mục đích:** Tìm giá trị ở **cột đầu** của bảng, trả về cột thứ n bên phải: `=VLOOKUP(A2,Product!$A$2:$E$42,5,FALSE)`.
- **Ảnh:** [CHỤP] cột B (FALSE) ra `#N/A` cho P041, P21; cột C (bỏ trống) ra **71.500** cho cả hai: `Demo/B2_Demo_Lookup.xlsx`, sheet Demo_VLOOKUP, A1:D7; [PDF p36–38].
- **TIP:** **Luôn** ghi đối số cuối là `FALSE` (hoặc 0). Bỏ trống = khớp gần đúng.

**Lời giảng:**
> *(→ demo: mở `Demo/B2_Demo_Lookup.xlsx`, sheet Demo_VLOOKUP; các bước ở sheet Huong_dan)* VLOOKUP là hàm tra cứu nổi tiếng nhất. Bốn đối số: **tìm cái gì**, **tìm trong bảng nào**, **lấy cột thứ mấy**, và **khớp kiểu gì**.
>
> Mình tra giá vốn, cột thứ 5 của bảng Product. Cột B viết đủ FALSE, cột C bỏ trống đối số cuối. Bốn mã đầu, hai cột giống nhau. Nhưng nhìn P041, mã hàng mới chưa có trong danh mục, và P21, mã gõ thiếu số 0. Cột B báo **#N/A**: đúng, không có mã này. Cột C trả về **71.500**. Không báo lỗi gì cả.
>
> 71.500 là gì? Là giá vốn của **P040**, dòng cuối cùng. Bỏ trống đối số cuối, VLOOKUP hiểu là **khớp gần đúng**: không thấy mã thì lấy mã gần nhất nhỏ hơn. Với bảng giá theo bậc thì hữu ích, còn tra mã sản phẩm thì là thảm họa. **Luôn viết FALSE.**
>
> VLOOKUP còn hai hạn chế: chỉ tìm được ở **cột đầu** và chỉ trả về cột **bên phải**; và số 5 là "số cứng", chèn thêm cột vào bảng là trả sai cột.

### Slide 21 · CÔNG CỤ · INDEX / MATCH · 4'
- **Tên (đỏ):** **INDEX / MATCH**
- **Mục đích:** MATCH tìm **vị trí**, INDEX lấy giá trị **ở vị trí đó**: `=INDEX(Product!$E$2:$E$42,MATCH(A2,Product!$A$2:$A$42,0))`.
- **Ảnh:** [CHỤP] cột B = INDEX/MATCH, F2 ở ô B2 (Excel chỉ tô màu ô A2; hai vùng trên sheet Product không được tô vì nằm ở sheet khác): `Demo/B2_Demo_Lookup.xlsx`, sheet Demo_XLOOKUP, A1:B7; bên cạnh là chữ giải thích MATCH → dòng 5 → INDEX → 83.500; [PDF p42–44].
- **TIP:** Chạy được trên **mọi** phiên bản Excel. Dùng khi máy chưa có XLOOKUP (Excel 2016/2019).

**Lời giảng:**
> Cách khắc phục cổ điển là cặp **INDEX / MATCH**. *(→ trên `Demo/B2_Demo_Lookup.xlsx`, sheet Demo_XLOOKUP, cột B)* MATCH trả lời "mã này nằm ở dòng thứ mấy", số 0 cuối là khớp chính xác. INDEX trả lời "cột giá vốn, dòng thứ mấy, là bao nhiêu".
>
> Vì cột tìm và cột trả về tách rời nhau nên không còn giới hạn "cột đầu, bên phải", chèn cột cũng không sai. Nhược điểm: dài và khó đọc. Bạn nào dùng Excel 2016 hoặc 2019 thì đây là cách chính của bạn; còn Microsoft 365 có cách gọn hơn.

### Slide 22 · CÔNG CỤ · XLOOKUP · 6'
- **Tên (đỏ):** **XLOOKUP**
- **Mục đích:** Tìm trong một cột, trả về cột khác, mặc định khớp chính xác: `=XLOOKUP(A2,Product!$A$2:$A$42,Product!$E$2:$E$42,"Không có mã")`.
- **Ảnh:** bảng đối số ① lookup_value ② lookup_array ③ return_array ④ [if_not_found] ⑤ [match_mode] ⑥ [search_mode]; [CHỤP] cột C ("Không có mã") và cột F (tra ngược "Gạo thơm 5kg" → P014): `Demo/B2_Demo_Lookup.xlsx`, sheet Demo_XLOOKUP, A1:F7; [PDF p45–46].
- **TIP:** Trong Table: `=XLOOKUP([@Product_Code],tblProduct[Product_Code],tblProduct[Unit_Cost])`.
- **PRO TIP:** Chỉ dùng `if_not_found` khi đã hiểu vì sao không tìm thấy; đừng dùng nó để che lỗi dữ liệu.

**Lời giảng:**
> **XLOOKUP** gộp mọi ưu điểm: tìm ở đâu, trả về đâu, chỉ cần chỉ ra hai cột. Mặc định là khớp chính xác, không có cái bẫy như VLOOKUP. *(→ trên `Demo/B2_Demo_Lookup.xlsx`, sheet Demo_XLOOKUP, cột C)* P005 ra 83.500, và P041, P21 hiện "Không có mã" nhờ đối số thứ tư.
>
> Vì cột trả về có thể nằm **bên trái**, mình tra ngược được: từ tên sản phẩm ra mã. *(→ cột F)* "Gạo thơm 5kg" ra P014.
>
> Trong phần thực hành, các bạn viết XLOOKUP ngay trong tblSales, theo tên cột, như trên slide. Viết một ô, Table tự điền 6.187 dòng. Và lưu ý PRO TIP: đối số "không tìm thấy" rất tiện, nhưng nếu dùng nó để giấu #N/A thì các bạn đang lặp lại sai lầm IFERROR bằng 0 lúc nãy.

### Slide 23 · KHÁI NIỆM · Laws of Lookups: khóa phải duy nhất · 6.5'
- **Trên slide:**
  - Luật 1: **Khóa ở bảng tra phải duy nhất.** Trùng khóa → XLOOKUP/VLOOKUP lấy **dòng đầu tiên**, không báo lỗi.
  - Luật 2: **Khóa hai bên phải cùng kiểu, cùng cách viết** (slide 24).
  - Luật 3: **Không thấy thì phải thấy #N/A**, đừng để khớp gần đúng.
  - [CHỤP] bảng Product dòng 22–23: P021 "(mã cũ – ngừng KD 2024)" Unit_Cost 62.000 đứng **trước** P021 Unit_Cost 69.000: `Demo/B2_Demo_Lookup.xlsx`, sheet Product, A1:E23 (ẩn tạm dòng 2–20).
  - Thẻ số: COGS của P021 trong Q1: **15.872.000** thay vì **17.664.000** (thiếu 1.792.000) · Gross Margin % toàn công ty 34,4% thay vì 34,2%.
- **TIP:** Kiểm tra trùng: `=COUNTIF(tblProduct[Product_Code],[@Product_Code])` > 1.

**Lời giảng:**
> Giờ đến bài học quan trọng nhất của phần này. *(→ trên `Demo/B2_Demo_Lookup.xlsx`, sheet Demo_KhoaTrung)* Đây là 143 dòng bán P021, bánh quy bơ hộp thiếc, trong Quý 1. Cột E là giá vốn tra bằng XLOOKUP: **62.000**. Không có ô nào báo lỗi.
>
> Nhưng giá vốn thật của P021 là 69.000. *(→ sang sheet Product, dòng 22–23)* Nhìn đây: P021 xuất hiện **hai lần**. Dòng trên là bản ghi cũ, sản phẩm đã ngừng kinh doanh năm 2024, còn sót lại trong danh mục. XLOOKUP gặp dòng nào **trước** thì lấy dòng đó, và im lặng.
>
> **Hậu quả là gì?** Doanh thu không đổi, vì giá bán lấy từ Sales. Nhưng giá vốn bị thấp đi gần 1,8 triệu, nên **biên lợi nhuận đẹp hơn thực tế**: 34,4% thay vì 34,2%. Chênh lệch nhỏ nên càng khó phát hiện; báo cáo lên sếp là con số sai mà ai cũng tin.
>
> Đây là **luật thứ nhất của lookup: khóa phải duy nhất.** Trước khi tra, luôn kiểm tra bảng tra bằng COUNTIF. *(→ ô D3, D4: P021 xuất hiện 2 lần; 2 dòng bị trùng)* *(→ xóa dòng 22 ở sheet Product; quay lại Demo_KhoaTrung: cột E thành 69.000, D5 = D6 = 17.664.000)* Xóa bản ghi cũ, mọi thứ đúng lại.
>
> Và đây là tinh thần D2 của khóa: **con số hiển thị chưa chắc đúng.** Không có lỗi không có nghĩa là không có sai.

### Slide 24 · CÔNG CỤ · Khóa khác kiểu: TRIM, VALUE · 3'

| Tên (đỏ) | Mục đích |
|---|---|
| **TRIM** | Bỏ dấu cách thừa ở đầu, cuối và giữa chuỗi |
| **VALUE / TEXT** | Đổi số dạng chữ `"1"` thành số `1` và ngược lại |

- **Ảnh:** [CHỤP] cột C ra #N/A với `"1"`, `"P005 "`, `" P013"`; cột E đúng sau VALUE / TRIM; ô `=LEN(A11)` = 5: `Demo/B2_Demo_Lookup.xlsx`, sheet Demo_KieuKhoa, A3:E14.
- **TIP:** `#N/A` dù "nhìn thấy có"? Kiểm tra `LEN` và căn lề (số căn phải, chữ căn trái).

**Lời giảng:**
> Luật thứ hai: **khóa hai bên phải giống nhau từng ký tự và cùng kiểu.** *(→ trên `Demo/B2_Demo_Lookup.xlsx`, sheet Demo_KieuKhoa)* Mã vùng số 1 trong bảng tra là **số**, nhưng giá trị tra lấy từ file CSV là chữ "1". Nhìn giống hệt, Excel vẫn báo #N/A. Sửa: VALUE để đổi chữ thành số.
>
> Mã "P005" có một dấu cách vô hình ở cuối. *(→ gõ `=LEN(A11)` ở ô trống: ra 5)* Bốn ký tự mà LEN ra 5. Sửa bằng **TRIM**. Bài về nhà hôm nay có đúng loại lỗi này, nên các bạn nhớ mẹo LEN.

---

## Phần 5 — Hàm Text và Date lõi (15 phút · slide 25–28)

### Slide 25 · CHIA PHẦN · 0.5'
- **Trên slide:** MCI / **5. Text & Date: chuẩn hóa để phân tích theo thời gian**

**Lời giảng:**
> Phần 5: vài hàm chữ và ngày để chuẩn bị cột Tháng, Quý, Thứ cho bảng KPI.

### Slide 26 · CÔNG CỤ · TRIM, LEFT/MID/RIGHT, TEXT, & · 6'

| Tên (đỏ) | Ví dụ | Kết quả |
|---|---|---|
| **LEFT / MID / RIGHT** | `=MID("HD2501-00001",3,4)` | `2501` (năm 25, tháng 01) |
| **TRIM** | `=TRIM("  MCI Mart   Cầu Giấy")` | `MCI Mart Cầu Giấy` |
| **&** | `=B2&" – "&TRIM(C2)` | `ST01 – MCI Mart Cầu Giấy` |
| **TEXT** | `=TEXT(D2,"#,##0")&" đ"` | `19,000 đ` (theo cài đặt vùng) |

- **Ảnh:** [CHỤP] cột E–J (LEFT, MID, RIGHT, TRIM, &, TEXT) cho 6 cửa hàng: `Demo/B2_Demo_Text_Date.xlsx`, sheet Demo_Text, A1:J7; [PDF p50–53].
- **TIP:** Kết quả của TEXT là **chữ**, không cộng trừ được. Chỉ dùng TEXT ở bước hiển thị.

**Lời giảng:**
> *(→ demo: mở `Demo/B2_Demo_Text_Date.xlsx`, sheet Demo_Text)* Mã đơn MCI Mart có quy luật: "HD", rồi năm-tháng, gạch ngang, số thứ tự. **LEFT** lấy ký tự bên trái, **RIGHT** bên phải, **MID** lấy ở giữa: từ ký tự thứ 3, lấy 4 ký tự, ra "2501".
>
> Tên cửa hàng gõ tay có dấu cách thừa lung tung. **TRIM** dọn sạch: bỏ cách ở đầu, cuối, và rút các cách ở giữa còn một. *(→ gõ `=LEN(C2)` và `=LEN(H2)`)* Độ dài từ 21 còn 17.
>
> Dấu **&** nối chuỗi: mã cửa hàng, gạch ngang, tên. Và **TEXT** định dạng số thành chữ theo mẫu. *(→ gõ `=J2+1`: ra #VALUE!)* Nhưng các bạn thấy: kết quả của TEXT là chữ, cộng vào là lỗi. **Tính toán bằng số, TEXT chỉ để trình bày.**

### Slide 27 · CÔNG CỤ · Hàm ngày lõi (EOMONTH, TEXT tháng/quý/thứ, NETWORKDAYS) · 6.5'

| Tên (đỏ) | Ví dụ (15/01/2025) | Kết quả |
|---|---|---|
| **EOMONTH** | `=EOMONTH(A2,0)` · `=EOMONTH(A2,-1)+1` | 31/01/2025 · 01/01/2025 |
| **MONTH / Quý** | `=MONTH(A2)` · `="Q"&ROUNDUP(MONTH(A2)/3,0)` | 1 · Q1 |
| **TEXT thứ** | `=TEXT(A2,"ddd")` · `=WEEKDAY(A2,2)` | Wed (hoặc T4) · 3 |
| **NETWORKDAYS** | `=NETWORKDAYS(B10,C10,$H$10:$H$15)` | Tháng 1/2025: 23 → **17** ngày làm việc khi trừ Tết |

- **Ảnh:** [CHỤP] bảng ngày B2:G5 và khối NETWORKDAYS (tháng 1: 23 → 17): `Demo/B2_Demo_Text_Date.xlsx`, sheet Demo_Date, A1:H15; [PDF p57–65].
- **TIP:** Ngày trong Excel là **số** (01/01/2025 = 45658). Ngày căn trái = ngày lưu dạng chữ, không group được.

**Lời giảng:**
> *(→ trên `Demo/B2_Demo_Text_Date.xlsx`, sheet Demo_Date)* Ngày trong Excel thật ra là một con số đếm từ năm 1900, nên cộng trừ được. **EOMONTH** trả về ngày cuối tháng: đối số 0 là tháng này, âm 1 là tháng trước. Cuối tháng trước cộng 1 là **ngày đầu tháng**, mẹo rất hay dùng.
>
> Để phân tích theo thời gian, bảng KPI cần cột **Tháng**, **Quý** và **Thứ**. MONTH ra số tháng. Quý thì chia tháng cho 3 rồi làm tròn lên. Thứ dùng TEXT với mẫu "ddd"; máy tiếng Anh ra "Wed", máy tiếng Việt ra "T4". Nếu cần con số để sắp xếp thì dùng WEEKDAY với kiểu 2: thứ Hai là 1, Chủ nhật là 7. Ngày 29/01/2025, mùng 1 Tết Ất Tỵ, là thứ Tư.
>
> Cuối cùng **NETWORKDAYS**: số ngày làm việc giữa hai ngày, có thể trừ danh sách ngày lễ. *(→ chỉ khối A10:E12)* Tháng 1/2025 có 23 ngày thứ Hai đến thứ Sáu, nhưng trừ Tết Dương lịch và 5 ngày nghỉ Tết thì chỉ còn **17 ngày**. Nhớ con số này; nó liên quan tới câu hỏi "vì sao" chúng ta vừa ghi lên bảng.

### Slide 28 · KHÁI NIỆM · Chọn hàm theo câu hỏi · 2'
- **Trên slide:**

| Câu hỏi kinh doanh | Nhóm hàm |
|---|---|
| "Sản phẩm này thuộc nhóm nào?" | IF / IFS / AND / OR |
| "Khách thường chi bao nhiêu? Có gì bất thường?" | MEDIAN, STDEV, QUARTILE |
| "Bao nhiêu, theo vùng / tháng?" | SUMIFS / COUNTIFS / AVERAGEIFS |
| "Lấy thông tin từ bảng khác" | XLOOKUP (INDEX/MATCH) |
| "Theo tháng / quý / thứ" | MONTH, EOMONTH, TEXT, WEEKDAY |

**Lời giảng:**
> Trước khi thực hành, chốt lại bằng một bảng: **bắt đầu từ câu hỏi, rồi mới chọn hàm.** Đừng hỏi "hàm này dùng làm gì", hãy hỏi "câu hỏi này cần hàm nào". Bảng này là cẩm nang cho phần thực hành ngay sau đây.

---

## Phần 6 — Thực hành trên case (30 phút · slide 29–31)

### Slide 29 · THỰC HÀNH 1 · Làm giàu tblSales (15 phút)
- **Trên slide:**
  1. Mở file khởi đầu, lưu thành `B2_<HoTen>.xlsx`. Mở `B2_Product_DuplicateKey.xlsx` → chuột phải tab **Product** → Move or Copy → chọn file của bạn, tick **Create a copy**.
  2. Kiểm tra khóa: `=COUNTIF(tblProduct[Product_Code],[@Product_Code])` → xóa bản ghi P021 cũ (dòng 22), xóa cột kiểm tra.
  3. Thêm vào tblSales: `Category`, `Unit_Cost` (XLOOKUP) · `COGS = [@Quantity]*[@Unit_Cost]` · `Gross_Margin = [@Revenue]-[@COGS]`.
  4. `Thang = MONTH([@Order_Date])` · `Quy = "Q"&ROUNDUP(MONTH(…)/3,0)` · `Thu = TEXT([@Order_Date],"ddd")` · `Phan_khuc = IFS(Unit_Price ≥ 100.000 "Cao", ≥ 30.000 "Trung", còn lại "Thấp")`.
- **Kiểm tra:** Không còn `#N/A`. Tổng COGS = **619.663.000**. Phan_khuc: Cao 1.814 · Trung 3.005 · Thấp 1.368 dòng.

**Lời giảng (giao bài):**
> Giờ đến lượt các bạn, 15 phút. Mở file khởi đầu *(→ `../B1_Nen_tang_Excel_Du_lieu/Practice/B1_MCI_Mart_Q1_2025_Answer.xlsx`)* và file sản phẩm *(→ `Practice/B2_Product_DuplicateKey.xlsx`)*. Việc đầu tiên là chép sheet Product sang file của bạn: chuột phải vào tab, Move or Copy, chọn file của bạn, nhớ tick "Create a copy".
>
> **Bước 2 là bước D2**: trước khi tra, kiểm tra khóa. Thêm một cột COUNTIF tạm vào bảng Product, lọc những dòng lớn hơn 1. Các bạn đã biết trước là P021; xóa dòng bản ghi cũ đi rồi xóa cột kiểm tra.
>
> Bước 3 và 4: thêm các cột vào tblSales bằng công thức theo tên cột. Gõ ở một ô, Table tự điền cả cột.
>
> Tự kiểm tra: chọn cột COGS, Status Bar phải ra **619.663.000**. Nếu bạn ra **617.871.000**, thiếu đúng 1.792.000, thì bạn chưa xóa bản ghi P021 cũ.

**Ghi chú giảng viên:**
- Đi vòng lớp. **Ba lỗi hay gặp nhất:**
  - Chép Product bằng copy–paste vùng thay vì Move or Copy → mất tên bảng `tblProduct`; đặt lại tên trong Table Design.
  - Quên xóa P021 cũ (COGS 617.871.000).
  - Cột Thu ra số (ví dụ 45658) → ô đang định dạng Number; đổi về General.
- Máy Excel 2016/2019 không có XLOOKUP, IFS: dùng INDEX/MATCH và IF lồng.
- Ai xong sớm: làm trước Thực hành 2.

### Slide 30 · THỰC HÀNH 2 · Bảng KPI & KPI Card (12 phút)
- **Trên slide:**
  1. Sheet mới `KPI_SUMIFS`: bảng **Doanh thu** Miền (A6:A8) × Tháng (B5:D5 = 1, 2, 3). **Một** công thức ở B6, kéo ra 9 ô (`$A6`, `B$5`); thêm cột Quý 1 và dòng Tổng.
  2. Bảng **Số đơn**: COUNTIFS thêm `tblSales[Line_No],1`.
  3. Bảng **AOV** = Doanh thu ÷ Số đơn · **Gross Margin %** = SUMIFS(Gross_Margin) ÷ Doanh thu. Dòng Tổng: tổng ÷ tổng.
  4. Viết 2 card: **Total Orders**, **Gross Margin %** (mẫu `B2_KPI_Definition_Card`).
- **Kiểm tra:** Tổng = **941.454.000** · T1 **425.850.000** · Số đơn **3.144** · AOV **299.445** · GM **34,2%**.

**Lời giảng (giao bài):**
> Bài 2 trả lời câu hỏi đầu buổi của sếp. 12 phút.
>
> Nhớ quy tắc: **một công thức, kéo ra cả bảng**. Nếu phải sửa từng ô là bạn đang khóa đô-la sai chỗ. Số đơn thì nhớ grain: thêm điều kiện Line_No bằng 1. AOV và biên lợi nhuận là tỷ lệ, nên dòng Tổng tính tổng chia tổng.
>
> Đáp án để tự kiểm: tổng bảng doanh thu 941.454.000, trùng tổng cột Revenue. Tháng 1 là 425 triệu 850 nghìn. Tổng số đơn 3.144. Nếu số đơn của bạn ra 6.187 thì bạn đang đếm dòng.
>
> Còn thời gian thì viết nốt hai KPI card. Không kịp thì hoàn thành ở nhà, cùng với bài về nhà.

### Slide 31 · KHÁI NIỆM · Chữa bài nhanh · 3'
- **Trên slide:** Checklist brick output B2:
  - [ ] tblSales có Category, Unit_Cost, COGS, Gross_Margin, Thang, Quy, Thu, Phan_khuc; không còn `#N/A`
  - [ ] tblProduct không còn mã trùng
  - [ ] Bảng KPI Miền × Tháng: tổng = 941.454.000; số đơn = 3.144
  - [ ] AOV, Gross Margin % dòng Tổng tính tổng ÷ tổng
  - [ ] 3 KPI Definition Card
- **Chân slide:** File đáp án `B2_MCI_Mart_Q1_2025_Answer.xlsx` (sheet Kiem_tra: 10 dòng đối soát) là file khởi đầu Buổi 3.

**Lời giảng:**
> *(→ mở `Practice/B2_MCI_Mart_Q1_2025_Answer.xlsx`, sheet Kiem_tra)* Đây là file đáp án, sheet Kiem_tra có 10 dòng đối soát, tất cả tick xanh. Để ý dòng cuối sheet: nếu COGS của bạn ra 617.871.000 thì là lỗi P021. File này là **file khởi đầu Buổi 3**, và bảng KPI_SUMIFS trong đó sẽ là "đáp án" để các bạn đối soát khi dựng Pivot và khi viết measure ở Buổi 5. **Đừng xóa nó.**
>
> *(→ hỏi lớp)* **Bạn nào tick đủ 5 ô?** Ai còn thiếu thì mở file đáp án so với file của mình để biết sót ở đâu.

**Ghi chú giảng viên (kiểm tra cuối buổi, theo Course Plan):**
- Tổng bảng KPI SUMIFS = tổng cột Revenue; không còn `#N/A`; đủ 3 KPI card (card nào chưa xong thì làm cùng bài về nhà).
- Nhắc lưu bảng SUMIFS: đối soát ở B3 và B5.
- Học viên dùng Mac: hỏi lại tiến độ máy Windows (hạn chót trước B4).
- D2 mong đợi: Developing (phát hiện P021 khi được nhắc; dùng IFERROR có chủ đích).

---

## Phần 7 — Quiz, insight, preview (10 phút · slide 32–41)

Quiz kiểu hỏi nhanh: khoảng 40 giây mỗi câu. **Câu 6–7 là câu tình huống.**

### Slide 32 · QUIZ · EXCEL QUIZ CHALLENGE · 0.2'

**Lời giảng:**
> 7 câu hỏi nhanh. Câu 6 và 7 là tình huống; các bạn phải suy luận.

### Slide 33 · QUESTION 1: IFS
- Giá bán 135.000 ra gì với `=IFS(D2>=30000,"Trung",D2>=100000,"Cao",TRUE,"Thấp")`? A) Cao · **B) Trung** · C) Thấp · D) #N/A

**Lời giảng:**
> Đáp án B, Trung. IFS dừng ở điều kiện đúng **đầu tiên**; 135 nghìn đã lớn hơn 30 nghìn nên không bao giờ tới điều kiện "Cao". Đặt ngưỡng cao trước.

### Slide 34 · QUESTION 2: STATISTICS
- Giá trị đơn Q1: trung bình 299.445, trung vị 198.000. Điều này cho biết gì? A) Dữ liệu có lỗi · **B) Có một số đơn rất lớn kéo trung bình lên** · C) Nửa số đơn trên 299.445 · D) Trung vị tính sai

**Lời giảng:**
> B. Dữ liệu lệch phải: vài đơn rất lớn kéo trung bình lên. Không phải lỗi, và C sai vì một nửa số đơn nằm dưới **trung vị**, không phải dưới trung bình.

### Slide 35 · QUESTION 3: SUMIFS
- Trong SUMIFS, đối số **đầu tiên** là gì? A) Vùng tiêu chí đầu tiên · B) Tiêu chí đầu tiên · **C) Vùng cần cộng** · D) Số tiêu chí

**Lời giảng:**
> C. SUMIFS: cột cần cộng đứng đầu, sau đó là từng cặp vùng – tiêu chí. Hàm SUMIF cũ, không có chữ S, thì ngược lại; đó là lý do nên dùng luôn SUMIFS.

### Slide 36 · QUESTION 4: VLOOKUP
- `=VLOOKUP(A2,Product!$A$2:$E$42,5)` bỏ trống đối số cuối thì tra kiểu gì? A) Khớp chính xác · **B) Khớp gần đúng** · C) Báo lỗi · D) Tìm từ dưới lên

**Lời giảng:**
> B. Và khớp gần đúng với mã sản phẩm là cái bẫy: P041 ra giá của P040. Luôn viết FALSE, hoặc dùng XLOOKUP.

### Slide 37 · QUESTION 5: DATE
- `=EOMONTH(DATE(2025,2,15),0)` trả về? A) 01/02/2025 · B) 15/03/2025 · **C) 28/02/2025** · D) 31/03/2025

**Lời giảng:**
> C, ngày cuối tháng 2 năm 2025, là 28 vì 2025 không nhuận. Muốn ngày đầu tháng thì EOMONTH với âm 1, cộng 1.

### Slide 38 · QUESTION 6: TÌNH HUỐNG — LOOKUP
- Sau khi XLOOKUP giá vốn, Gross Margin % cao hơn tháng trước dù không có ô lỗi nào. Việc **đầu tiên** nên làm?
  - A) Đổi sang VLOOKUP
  - **B) Kiểm tra khóa trùng trong bảng Product (COUNTIF)**
  - C) Bọc IFERROR
  - D) Định dạng lại cột %

**Lời giảng:**
> B. Không có lỗi không có nghĩa là không có sai. Khóa trùng làm XLOOKUP lấy dòng đầu tiên mà không báo gì, đúng như P021 hôm nay. **Khi con số "đẹp bất thường", nghi ngờ trước, ăn mừng sau.**

### Slide 39 · QUESTION 7: TÌNH HUỐNG — KPI TỶ LỆ
- Bạn có AOV của 3 miền; sếp hỏi AOV toàn công ty. Tính thế nào? A) AVERAGE 3 AOV · **B) Tổng doanh thu ÷ tổng số đơn** · C) MEDIAN 3 AOV · D) Lấy AOV miền lớn nhất

**Lời giảng:**
> B. KPI tỷ lệ luôn tổng chia tổng. Lấy trung bình 3 miền là coi miền nhỏ nặng ngang miền lớn, giống ví dụ biên lợi nhuận 31,2% và 34,2% lúc nãy.

### Slide 40 · TỔNG KẾT + INSIGHT · 2'
- **Trên slide:**
  - **Điểm chính cần nhớ:**
    - (1) Định nghĩa KPI (card) trước, công thức sau.
    - (2) Biết grain: 6.187 dòng ≠ 3.144 đơn.
    - (3) Mỗi tiêu chí SUMIFS là một bộ lọc.
    - (4) Khóa lookup phải duy nhất; không lỗi ≠ không sai.
    - (5) KPI tỷ lệ: tổng ÷ tổng.
  - **Insight của bạn (1 câu):** "Doanh thu tháng ___ là ___ *(con số, có so sánh)*, vì ___ *(nguyên nhân khả dĩ)*, nên ___ *(hàm ý)*."

**Lời giảng:**
> Tổng kết 5 điều cần nhớ *(→ đọc lướt 5 ý trên slide)*.
>
> Giờ viết câu insight, 1 phút, nhìn vào bảng KPI của bạn.
>
> *(→ chờ 1 phút, gọi 2–3 bạn đọc)*
>
> Câu mẫu: "Doanh thu tháng 1/2025 là 425,9 triệu, cao hơn tháng 2 khoảng 73% (246 triệu). Số đơn giảm từ 1.398 xuống 823 trong khi AOV gần như giữ nguyên, khoảng 300 nghìn, nên doanh thu giảm vì **ít lượt mua hơn**, không phải vì khách chi ít hơn mỗi lần; cần xem theo ngày để biết điều gì xảy ra quanh cuối tháng 1."
>
> Để ý: câu này chưa nói "vì Tết". Chúng ta **nghi** là Tết, và nhớ NETWORKDAYS tháng 1 chỉ còn 17 ngày làm việc, nhưng chưa chứng minh. Buổi 3, với Pivot theo ngày, các bạn sẽ tự kiểm chứng.

### Slide 41 · PREVIEW B3 · 1'
- **Trên slide:**
  - *Buổi sau:* "Doanh thu biến động thế nào theo cửa hàng, ngày trong tuần và ngành hàng — và vì sao tháng 1 cao hơn hẳn tháng 2?" (Tầng 2–3)
  - Giới hạn hôm nay: **mỗi chiều phân tích mới lại phải viết lại SUMIFS** → PivotTable.
  - *Về nhà:* đọc handout B2 · làm **Bài tập về nhà (bắt buộc)** `B2_BTVN_MCI_Mart_T4_2025.xlsx`, nộp trước Buổi 3 tại hocvien.mcivietnam.com, tên file `B2_BTVN_<HoTen>.xlsx` · hoàn thành 3 KPI card · tick "Checklist trước Buổi 3" (mục 10 handout).
- **Ảnh:** [VẼ] viên gạch B2 (đã xong) → viên gạch B3 (sáng lên).

**Lời giảng:**
> Hôm nay các bạn đã trả lời được câu hỏi của sếp theo miền và tháng. Nhưng nếu sếp hỏi tiếp "theo cửa hàng", "theo thứ trong tuần", "theo ngành hàng", thì mỗi chiều mới các bạn lại phải dựng một bảng SUMIFS mới. Đó là **giới hạn của SUMIFS**, và là lý do Buổi 3 học **PivotTable**: kéo thả là đổi chiều. Buổi sau chúng ta cũng sẽ trả lời câu hỏi "vì sao tháng 1 cao hơn hẳn tháng 2".
>
> Về nhà: đọc handout Buổi 2, có thêm nhiều hàm hay như MAXIFS, SWITCH, WORKDAY. Làm **bài tập về nhà bắt buộc**: dữ liệu tháng 4 đã sạch, các bạn tra Region, ngành hàng, giá vốn, dựng bảng doanh thu Miền × Ngành hàng bằng một công thức SUMIFS, bảng số đơn, AOV, biên lợi nhuận theo miền, viết card cho AOV và một câu insight. Trong file có một lỗi lookup được cài sẵn; nhớ mẹo LEN hôm nay. Nộp trước Buổi 3 trên hocvien.mcivietnam.com.
>
> Cảm ơn cả lớp, hẹn gặp lại Buổi 3!

---

## Phụ lục kịch bản

### Kiểm tra thời lượng

| Phần | Slide | Phút | Timetable B2 |
|---|---|---|---|
| 0 Mở đầu, recap, KPI Card, D/M/G | 1–5 | 10 | 10 |
| 1 Hàm logic | 6–9 | 15 | 15 |
| 2 Thống kê mô tả | 10–13 | 15 | 15 |
| 3 COUNTIFS / SUMIFS / AVERAGEIFS | 14–17 | 20 | 20 |
| Nghỉ | 18 | 10 | 10 |
| 4 Lookup | 19–24 | 25 | 25 |
| 5 Text & Date | 25–28 | 15 | 15 |
| 6 Thực hành | 29–31 | 30 | 30 |
| 7 Quiz, insight, preview | 32–41 | 10 | 10 |
| **Tổng** | **41** | **150** | **150** |

Phần 7 cộng từng slide khoảng 8 phút (slide 32: 0,2'; 7 câu × khoảng 40 giây; slide 40: 2'; slide 41: 1'); khoảng 2 phút dự phòng cho câu 6–7 và phần đọc insight.

### Danh sách ảnh cần chụp [CHỤP] (Excel 365 Windows, zoom 100%)

Theo `CLAUDE.md` (mục Ảnh minh họa): thao tác và kết quả Excel dùng ảnh chụp thật trên file của khóa. Deck `Slide_B2.pptx` đang để **khung chờ [CHỤP]** (viền nét đứt, ghi file · sheet · vùng) ở các slide dưới đây. Mọi slide CÔNG CỤ đều có ít nhất một ảnh chụp. Hình [VẼ] còn lại chỉ là sơ đồ khái niệm và thẻ số (slide 2, 3, 8, 11, 12 trục số, 15, 17 thẻ GM%, 41). Ảnh lưu vào `Materials/B2/Images/` đúng tên file ở cột 2; chạy lại `build_slides_b2.py` thì ảnh tự thay vào khung (danh sách cả deck và handout: `Materials/B2/Images/_DANH_SACH_ANH.md`).

| # | Tên file | Nội dung ảnh | File · sheet · vùng | Slide |
|---|---|---|---|---|
| 1 | S04_grain.png | Đơn HD2501-00002 có 4 dòng (Line_No 1–4), chọn dòng 3–6 | `../B1_Nen_tang_Excel_Du_lieu/Practice/B1_MCI_Mart_Q1_2025_Answer.xlsx` · Sales · A1:K7 (ẩn tạm cột D–F, I–J) | 4 |
| 2 | S07_if_ifs.png | Cột Margin_%, IF, IFS sau khi kéo; P041 ra #DIV/0! và "Thấp" | `Demo/B2_Demo_Logic.xlsx` · Demo_IF · A1:H6 + dòng 42–43 (ẩn tạm 7–41) | 7 |
| 3 | S08_and_or_not.png | Cột I, J, K (AND, OR, NOT) ra TRUE/FALSE trên các dòng Sữa và Bánh kẹo | `Demo/B2_Demo_Logic.xlsx` · Demo_IF · A1:K1 + dòng 9–13, 22–28 | 8 |
| 4 | S09_iferror.png | P041, P042: #DIV/0!, "Chưa có giá"; L42 = IFERROR(…,0) ra 0% | `Demo/B2_Demo_Logic.xlsx` · Demo_IF · A41:L43 | 9 |
| 5 | S11_average_median.png | Khối I2:I6 (COUNT, AVERAGE, MEDIAN, MIN, MAX) và Status Bar khi chọn cột E | `Demo/B2_Demo_ThongKe.xlsx` · Demo_ThongKe · E1:I13 + Status Bar | 11 |
| 6 | S12_stdev_outlier.png | Khối I7:I13: STDEV.S, SKEW, Q1, Q3, IQR, ngưỡng 896.500, 138 đơn | `Demo/B2_Demo_ThongKe.xlsx` · Demo_ThongKe · G7:I13 | 12 |
| 7 | S13_rank.png | RANK.EQ 12 cửa hàng, F2 hiện `$D$2:$D$13` | `Demo/B2_Demo_ThongKe.xlsx` · Demo_RANK · A1:E13 | 13 |
| 8 | S16_sumifs_screentip.png | ScreenTip khi gõ `=SUMIFS(` (sum_range, criteria_range1, criteria1…) | `Demo/B2_Demo_SUMIFS.xlsx` · Demo_SUMIFS · ô trống | 16 |
| 9 | S17_sumifs_grid.png | Bảng Miền × Tháng sau khi kéo; F2 ở C13 hiện `$A13`, `C$11` | `Demo/B2_Demo_SUMIFS.xlsx` · Demo_SUMIFS · A11:E15 | 17 |
| 10 | S20_vlookup.png | VLOOKUP FALSE vs bỏ trống (P041, P21 → 71.500) | `Demo/B2_Demo_Lookup.xlsx` · Demo_VLOOKUP · A1:D7 | 20 |
| 11 | S21_index_match.png | INDEX/MATCH, F2 ở B2: chỉ ô A2 được tô màu (vùng trên sheet Product không tô) | `Demo/B2_Demo_Lookup.xlsx` · Demo_XLOOKUP · A1:B7 | 21 |
| 12 | S22_xlookup.png | XLOOKUP "Không có mã" và tra ngược tên → mã | `Demo/B2_Demo_Lookup.xlsx` · Demo_XLOOKUP · A1:F7 | 22 |
| 13 | S23_p021_duplicate.png | Product dòng 22–23: P021 hai lần | `Demo/B2_Demo_Lookup.xlsx` · Product · A1:E23 (ẩn tạm 2–20) | 23 |
| 14 | S24_trim_value.png | Khóa khác kiểu: #N/A, sửa bằng VALUE / TRIM, LEN = 5 | `Demo/B2_Demo_Lookup.xlsx` · Demo_KieuKhoa · A3:E14 | 24 |
| 15 | S26_text.png | LEFT, MID, RIGHT, TRIM, &, TEXT cho 6 cửa hàng | `Demo/B2_Demo_Text_Date.xlsx` · Demo_Text · A1:J7 | 26 |
| 16 | S27_date.png | EOMONTH, TEXT, WEEKDAY và NETWORKDAYS (23 → 17) | `Demo/B2_Demo_Text_Date.xlsx` · Demo_Date · A1:H15 | 27 |

Ảnh nên chụp thêm khi có thời gian (chưa có khung trên deck): Filter Bắc + tháng 3, Status Bar Sum = 90.210.000 (slide 15, KHÁI NIỆM); hộp Move or Copy sheet Product (slide 29, THỰC HÀNH).

### Còn cần làm trước khi dạy
- In `Handout/B2_KPI_Definition_Card.pdf` (3 tờ / học viên).
- Gửi trong nhóm chat trước buổi: file khởi đầu (`B1_MCI_Mart_Q1_2025_Answer.xlsx`) và `B2_Product_DuplicateKey.xlsx`.
- Chụp 16 ảnh [CHỤP] ở bảng trên, lưu vào `Materials/B2/Images/`, chạy lại `build_slides_b2.py` và `build_student_slides.py B2`.
- Chụp 7 ảnh cho handout theo `Materials/B2/Images/_DANH_SACH_ANH.md` (H01–H07), chạy lại `build_docs_b2.py` và xuất PDF.

### Ghi chú so với Course Plan

- Course Plan (Cách dạy B2) nói "XLOOKUP ra giá sai → vì sao doanh thu lệch?". Vì `Unit_Price` đã có trong Sales, bản ghi P021 cũ chỉ làm sai **giá vốn**, nên kịch bản hỏi "vì sao biên lợi nhuận đẹp bất thường?" (COGS thiếu 1.792.000; GM 34,4% thay vì 34,2%). Đã cập nhật câu này trong Course Plan.
- Hằng số minh họa AOV và GM trong kịch bản tính trên 3.144 đơn Q1/2025; KPI dictionary (Phụ lục C) giữ target 300.000 và 33%.
