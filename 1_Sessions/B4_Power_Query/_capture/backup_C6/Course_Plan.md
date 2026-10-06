# Excel for Business Intelligence Analyst — Plan tổng thể 7 buổi (v3.3)

> MCI – Magic Code Institute.
> - **7 buổi × 150 phút.**
> - Tài liệu nguồn: `4_Reference/` (số trang là trang của file PDF).
> - Phong cách slide: mỗi công cụ hoặc hàm đi kèm một ảnh minh họa. Deck `Slide_S1_GioiThieuExcel.pptx` gốc **đã bị xóa** (2026-09-23); các thông số style của nó được giữ lại ở mục "Style spec" cuối file.
> - v2 áp dụng khuyến nghị từ `3_Design_Reviews/Review_01_Knowledge_Architecture.md` (R1) và `Review_02_BI_Practice.md` (R2). v2.1 áp dụng thêm `Review_04_Coverage_Audit.md` (R4), audit theo chuẩn Microsoft MO-211 và PL-300. v3 đưa toàn bộ chi tiết giảng dạy vào file này, nên có thể dùng độc lập mà không cần mở các file review (mỗi buổi có mục "Hướng dẫn giảng dạy", cùng các Phụ lục A–E). v3.1 đưa vào các phần tốt của bản nháp `Curriculum_Guide.html`: thang 5 tầng câu hỏi, câu chuyện dữ liệu theo buổi (Phụ lục F), dự báo nhập môn ở B6. Xem Changelog ở cuối file.

## Yêu cầu đầu vào *(R2 §1)*

- **Excel 365 hoặc Excel 2016+ bản Windows, ưu tiên 64-bit.** Power Pivot không chạy trên Mac và Excel Online (PQ PDF p5), nên không đáp ứng yêu cầu này thì không học được từ B4 (load vào Data Model) đến B7.
- **Thông báo sớm:** nêu yêu cầu này **trước khóa học** (email xác nhận đăng ký / email chào mừng) và nhắc lại ngay **phút đầu B1**, để học viên chuẩn bị máy từ đầu thay vì phát hiện ra giữa khóa.
- **Tự kiểm tra máy trước B1:** gửi `B1_Checklist_KiemTraMay` qua nhóm chat của lớp để học viên tự làm 3 bước. Không đạt hoặc không làm được thì liên hệ trợ giảng ngay trong B1.
- **Xác nhận ở B1** (học viên đã tự kiểm tra trước; bạn chưa đạt thì trợ giảng hỗ trợ ngay trong lớp):
  - Data → Get Data (có Power Query).
  - File → Options → Add-ins → COM Add-ins → bật *Microsoft Power Pivot*.
- Học viên dùng Mac: **khuyến nghị cài phần mềm máy ảo** (Parallels Desktop hoặc VMware Fusion) để chạy Windows và Excel bản Windows ngay trên máy Mac; nếu không cài được thì mang một laptop Windows khác.
  - **Học viên tự chuẩn bị máy.** Trung tâm không cung cấp phòng lab hay máy cho học viên; tài liệu không được hứa phương án này.
  - **Khuyến nghị:** có máy Windows **từ B2**, để làm quen giao diện và phím tắt Windows ngay trong B2–B3.
  - **Hạn chót:** trước B4.

## Mục tiêu đầu ra *(R2 §2.7, thang Self-service BI maturity)*

| Mức | Năng lực | Đạt ở buổi |
|---|---|---|
| 1 – Consumer | Xem và lọc báo cáo | B1 |
| 2 – Explorer | Phân tích ad-hoc, biểu đồ cơ bản | B2–B3 |
| 3 – Builder | Kết hợp nguồn dữ liệu, xây báo cáo và dashboard | B4, B6 |
| 4 – Modeler | Xây mô hình dữ liệu, định nghĩa KPI | B5 |

- **Cam kết:** 100% học viên đạt Level 3.
- Học viên khá đạt Level 4, đánh giá qua capstone B7.

## Nguyên tắc xuyên suốt

1. **BI trước, Excel sau.** Mỗi buổi **mở bằng một câu hỏi kinh doanh** trên case và **đóng bằng một câu insight** do học viên tự viết (con số, nguyên nhân khả dĩ, hàm ý). *(R1 §4.9)*
2. **Mỗi công cụ có một ảnh minh họa.** Mẫu slide: tên công cụ (đỏ) → mục đích một dòng → ảnh → TIP / PRO TIP.
3. **Ba tầng trong mỗi buổi:**
   - **Lõi:** giảng tại lớp, bắt buộc.
   - **Nâng cao:** PRO TIP cho học viên đã biết.
   - **Handout:** tự học, vẫn có ảnh minh họa theo style S1.
   Điều này giúp lớp hỗn hợp không bị quá tải. *(R1 §4.7)*
4. **Brick theo khái niệm, không chỉ theo buổi.** Ba chuỗi xương sống đi xuyên khóa học (mục tiếp theo).
5. **Một case xuyên suốt** *(đề xuất, cần chốt dữ liệu thật)*: chuỗi bán lẻ bối cảnh Việt Nam.
   - Các bảng: `Sales` (giao dịch), `Product`, `Store`, `Customer`, `Calendar`.
   - **`Target`**: mục tiêu theo Tháng × Cửa hàng, dạng cột tháng. *(R2 §2.2)*
   - File CSV bán hàng theo tháng, dùng cho B4.
6. **Lưới an toàn cho brick:** mỗi buổi phát **file khởi đầu**, tức là đáp án brick output của buổi trước. Học viên chưa hoàn thành buổi trước vẫn theo kịp buổi sau, và chuỗi brick không bị gãy.
7. **Quiz cuối buổi:** khoảng 7 câu trắc nghiệm, **trong đó 2 câu tình huống** dạng "Câu hỏi X nên giải bằng công cụ nào? Vì sao?". *(R1 §5)*
8. **Bốn năng lực thực hành (D1–D4)** được luyện ở mọi buổi và chấm ở B7:
   - D1 Bắt đầu từ câu hỏi kinh doanh
   - D2 Hoài nghi và đối soát dữ liệu
   - D3 Tự động hóa, tái sử dụng
   - D4 Truyền đạt insight
   Bảng mức độ chi tiết: **Phụ lục B**.
9. **Thang 5 tầng câu hỏi.** Quy trình BI 5 bước trả lời "làm theo thứ tự nào"; thang này trả lời "câu hỏi sâu đến đâu". Câu hỏi mở mỗi buổi được gắn tầng, để học viên thấy mình đang leo thang:

   | Tầng | Câu hỏi | Loại phân tích | Buổi |
   |---|---|---|---|
   | 1 | Chuyện gì đã xảy ra? | Descriptive | B1–B2 |
   | 2 | Ở đâu, khi nào, nhóm nào? | Dimensional | B3 |
   | 3 | Vì sao? | Diagnostic | B3 (drill-down), B5 |
   | 4 | Vậy thì sao, nên làm gì? | Prescriptive | B6 |
   | 5 | Sắp tới sẽ thế nào? | Predictive | B6 (nhập môn) |

   B4 không có tầng riêng: đó là nền móng giúp các tầng sau đáng tin. B7 đi đủ 5 tầng trên bộ dữ liệu mới.
10. **Câu chuyện cài sẵn trong dữ liệu.** Dataset MCI Mart có 6 "câu chuyện" (xem `2_Dataset/README.md`). Mỗi buổi mở khóa một phần, và câu insight đóng buổi nên chạm vào câu chuyện đó. Bảng phân bổ: **Phụ lục F**.
11. **Bài tập về nhà là bước tập dượt cho capstone.** Mỗi bài về nhà là một phần việc của dự án B7, đặt trong một tình huống công việc (một bộ dữ liệu mới, hoặc một báo cáo có vấn đề), có lỗi cài sẵn chỉ lộ ra khi đối soát (D2), và gắn với một hoặc vài tiêu chí của rubric B7. Cách chấm giữ như hiện nay (sheet `Kiem_tra` của file `_Answer`, agent `homework-grader`). Đến cuối B6, học viên đã tự làm đủ chuỗi Brief → Power Query → Data Model → KPI measure → Dashboard → Insight một lần, từng phần, trước khi nhận brief capstone.

   | Bài về nhà | Phần việc của capstone | Tiêu chí rubric B7 |
   |---|---|---|
   | B1 | Làm sạch, tidy, Excel Table, Validation | C2 |
   | B2 | KPI bằng SUMIFS, KPI card | C4 |
   | B3 | Pivot, PivotChart, Slicer; đối soát với số hệ thống | C4, C5 |
   | B4 | Pipeline Power Query, đối soát, % target | C2 |
   | B5 | Kiểm toán và sửa một báo cáo cũ: mô hình ngôi sao, measure, KPI dictionary, câu hỏi đo lường được | C3, C4 (tập trước C1) |
   | B6 | Brief, dashboard, insight | C1, C5, C6 |

   Bài nào cũng kết thúc bằng một câu insight (C6).

### Ba chuỗi xương sống *(bản chi tiết có tiên quyết cứng/mềm: Phụ lục A)*

```
A. Cấu trúc dữ liệu:  Tidy data + Excel Table (B1) → Dimension/Measure/Grain (B2) → Pivot (B3)
                      → Unpivot (B4) → Star schema: Fact vs Lookup (B5)
B. Nối bảng:          $ reference (B1) → XLOOKUP + "khóa phải duy nhất" (B2) → Merge (B4)
                      → Relationship + Cardinality (B5)
C. Tổng hợp có lọc:   SUMIFS: criteria = bộ lọc (B2) → Ô Pivot = tổng theo tọa độ (B3)
                      → Filter context, CALCULATE (B5) → KPI card bằng CUBEVALUE (B6)
```

**Hình minh họa xương sống cho Chuỗi C** (dùng lại ở B2, B3, B5): một ô kết quả → các tiêu chí lọc (Vùng = Bắc, Tháng = 3) → phép tổng.
- B2 gọi là *criteria của SUMIFS*.
- B3 gọi là *tọa độ ô Pivot* (Pivot PDF p11).
- B5 gọi là *filter context* (PQ PDF p73–75).

Cùng một khái niệm, học ba lần ở ba mức.

### Bản đồ Brick

```
B1 Dữ liệu gốc tidy (Excel Table) ─▶ B2 Bảng làm giàu + bảng KPI bằng SUMIFS ─▶ B3 Báo cáo Pivot
   ─▶ B4 Pipeline Power Query tự refresh ─▶ B5 Star schema + KPI dictionary (đã đối soát với B2)
   ─▶ B6 Dashboard có target/RAG ─▶ B7 Dự án BI end-to-end + thuyết trình
```

Mỗi buổi để lộ **giới hạn** của công cụ vừa học, và giới hạn đó là lý do cần buổi sau:
- SUMIFS phải viết lại cho mỗi chiều phân tích → cần Pivot.
- Dữ liệu bẩn, tách nhiều file → cần Power Query.
- Merge thành một bảng lớn thì nặng và trùng lặp → cần Data Model.
- Bảng số khó đọc → cần Dashboard.

---

## Buổi 1 — Giới thiệu Excel, Nền tảng công thức & Dữ liệu có cấu trúc

**Câu hỏi BI mở buổi:** "Công ty có một file bán hàng. Trước khi phân tích, làm sao biết file này dùng được?" *(Tầng 1 — Descriptive)*

**Mục tiêu** — học viên có thể:
- Mô tả quy trình BI 5 bước và vị trí của Excel (Data → Logic → Presentation).
- Xác nhận máy mình đủ điều kiện cho khóa học: Windows, có Power Query và Power Pivot.
- Thao tác nhanh bằng phím tắt; viết công thức đúng cú pháp; dùng đúng 4 kiểu tham chiếu.
- Đọc và sửa 6 lỗi phổ biến; dùng Auditing Tools để "debug" báo cáo.
- **Nhận biết dữ liệu tidy** và chuyển vùng dữ liệu thành **Excel Table**. *(R1 §4.2)*
- Kiểm soát chất lượng nhập liệu bằng Data Validation.

| Phút | Nội dung | Nguồn |
|---|---|---|
| 10 | Giới thiệu khóa học: **Quy trình BI 5 bước** (Hỏi → Thu thập → Làm sạch → Phân tích → Trình bày), giới thiệu case, bản đồ Brick | *mới* |
| 10 | Phiên bản Excel. So sánh 2016 / 365 / Online; **nêu rõ Mac và Online không đủ cho khóa này.** Xác nhận kiểm tra máy (học viên đã tự làm trước buổi) | PQ PDF p5 |
| 15 | Giao diện rút gọn: Workbook/Sheet/Cell, Name Box, Formula Bar. **Tour 8 tab Ribbon chuyển thành handout** | Ảnh chụp giao diện Excel 365 |
| 15 | Phím tắt: Ctrl, Fn, Alt Key Tips, Quick Access Toolbar | Formulas PDF p11–14; Pro Tips (Productivity) |
| 10 | *Nghỉ* | |
| 15 | Formula syntax; Relative/Absolute/Mixed reference (F4) | Formulas PDF p5–6 |
| 10 | Common Errors; Auditing (Trace, Show Formulas, Evaluate, Error Checking, **Watch Window**) | Formulas PDF p7–10 *(R4: MO-31)* |
| 10 | **Tidy data (Good vs Bad structure) + Excel Table (Ctrl+T), structured reference** | Pivot PDF p5 *(chuyển từ B3)* |
| 10 | Data Validation | Formulas PDF p15 |
| 35 | **Thực hành:** làm sạch cấu trúc file case → Excel Table; thêm cột Revenue và % đóng góp (dùng `$`); Data Validation + Circle Invalid Data; Watch Window | |
| 10 | Quiz + insight cuối buổi + preview B2 | |

- **Handout:** 8 tab Ribbon, Mac shortcuts, Bonus (GETPIVOTDATA, QAT), Top 10 hàm, **Flash Fill, Fill Series nâng cao, tham chiếu sang workbook khác** *(R4: MO-09, MO-10, MO-02)*.
- **Brick output:** File case dạng Excel Table, tidy, có Data Validation, không có lỗi công thức.
- **Bài tập về nhà (bắt buộc), nộp trước B2:** làm sạch báo cáo tháng 4/2025 xuất theo cửa hàng (`Homework/B1_BTVN_MCI_Mart_T4_2025.xlsx`), lặp lại đủ quy trình B1: tidy → Table → Revenue, Pct_Share (`$`) → Validation + Circle Invalid Data → 1 câu insight. Đáp án và sheet Kiem_tra để chấm ở file `_Answer`.

### Hướng dẫn giảng dạy B1

**Khái niệm then chốt:**
- **Tham chiếu `$`:** tiên quyết cứng cho Chuỗi B (lookup) và Chuỗi C (SUMIFS). Học viên chưa vững `$` sẽ vấp ở B2.
- **Tidy data:** 1 dòng = 1 bản ghi, 1 cột = 1 trường. Đây là tiên quyết cứng của mọi buổi sau.
- **Excel Table:** tự mở rộng khi thêm dòng. Đây là "hạt giống" của năng lực D3 (tự động hóa), sẽ được hái quả ở B4.

**Cách dạy:**
- Mở bằng một file "bad structure" thật: ô gộp, tiêu đề 2 dòng, dòng tổng nằm giữa dữ liệu. Hỏi: "Nếu phải tính doanh thu theo vùng, file này gây khó ở đâu?" → rút ra nguyên tắc tidy → chuyển thành Table.
- Với tham chiếu: cho kéo công thức sai trước, để học viên thấy kết quả lệch, rồi mới dạy F4.

**Lỗi thường gặp → cách xử lý:**

| Lỗi | Dấu hiệu | Xử lý |
|---|---|---|
| Quên `$` khi kéo công thức | Kết quả lệch dần theo dòng | F4; Trace Precedents để thấy vùng bị trôi |
| Số lưu dạng text | Căn trái, SUM = 0, có tam giác xanh | Error Checking → Convert to Number; VALUE |
| Ô gộp, ô trống giữa bảng, dòng tổng trong dữ liệu | Sort/Filter lỗi, tổng bị nhân đôi | Unmerge, xóa dòng tổng, chuyển sang Table |

**Kiểm tra cuối buổi:**
- File case đã là Table, không có ô gộp hay dòng tổng, có Data Validation.
- **D2 mong đợi:** Developing (phát hiện lỗi dữ liệu khi được nhắc).

**Điều chỉnh cho lớp đã biết Excel sơ bộ:** rút Giao diện + Phím tắt từ 30 xuống 15 phút; dồn 15 phút vào Auditing (Evaluate Formula, Watch Window) và tidy data.

**Chuẩn bị:**
- File case bản "bẩn".
- Checklist kiểm tra máy (Get Data, COM Add-in Power Pivot, 64-bit).
- Handout 8 tab Ribbon.

---

## Buổi 2 — Hàm cho phân tích: từ dữ liệu giao dịch đến KPI

**Câu hỏi BI mở buổi:** "Doanh thu, số đơn, biên lợi nhuận của từng vùng tháng này là bao nhiêu?" *(Tầng 1 — Descriptive)*

**Mục tiêu** — học viên có thể:
- **Định nghĩa một KPI** trước khi tính: công thức, grain, dimensions. **Phân biệt Dimension, Measure và Grain.** *(R1 §4.3, R2 §2.1)*
- Viết logic kinh doanh bằng IF/IFS/AND/OR/IFERROR.
- Chọn đúng thống kê mô tả: mean hay median, độ lệch chuẩn, outlier.
- Tính KPI có điều kiện bằng COUNTIFS/SUMIFS/AVERAGEIFS (**Chuỗi C, lần 1**).
- Nối bảng bằng XLOOKUP và hiểu "Laws of Lookups": khóa phải duy nhất (**Chuỗi B**).
- Chuẩn hóa text và date cho phân tích theo thời gian.

| Phút | Nội dung | Nguồn |
|---|---|---|
| 10 | Recap B1 + **KPI Definition Card** (tên, công thức, grain, dimensions) + **Dimension / Measure / Grain** | *mới*; Phụ lục C |
| 15 | IF, IFS, AND/OR, NOT, IFERROR | Formulas PDF p17–24 |
| 15 | Thống kê mô tả: COUNT, AVERAGE vs MEDIAN, MIN/MAX, STDEV, RANK; skew, outlier | Formulas PDF p25–27; Statistics PDF p24–38 |
| 20 | COUNTIFS / SUMIFS / AVERAGEIFS + **hình xương sống Chuỗi C** | Formulas PDF p33–34 |
| 10 | *Nghỉ* | |
| 25 | VLOOKUP → INDEX/MATCH → XLOOKUP; Laws of Lookups | Formulas PDF p36–46 |
| 15 | Text lõi (TRIM, LEFT/MID/RIGHT, TEXT, &) và Date lõi (EOMONTH, TEXT tháng/quý, NETWORKDAYS) | Formulas PDF p50–53, 57–65 |
| 30 | **Thực hành trên case:** XLOOKUP `Category` và `Unit_Cost` từ `Product` (giá bán `Unit_Price` đã có sẵn trong Sales), cột COGS và Gross Margin, Tháng/Quý, **Thứ** (`TEXT(ngày,"ddd")`, dùng ở B3), Phân khúc (IFS); **bảng KPI theo Vùng × Tháng bằng SUMIFS** (bảng này sẽ là "đáp án" đối soát ở B5) | |
| 10 | Quiz + insight + preview B3 | |

- **Nâng cao (PRO TIP):**
  - Dynamic Arrays (FILTER, SORT, UNIQUE, LET), Top 5 bằng SORT+FILTER (Formulas PDF p70–96)
  - SUMPRODUCT (p30–32)
  - Fuzzy-match lookup, Uniques & Duplicates (Pro Tips – Formula)
- **Handout:** HLOOKUP, CHOOSE, OFFSET, ROW/COLUMN, SEARCH/FIND, YEARFRAC, DATEDIF, RAND/**RANDARRAY**, **TODAY/NOW, WEEKDAY, WORKDAY**, **SWITCH, MAXIFS/MINIFS**, formula-based formatting, INDIRECT, HYPERLINK *(R4: MO-11, MO-20, MO-22, MO-23)*.
- **Brick output:** Bảng `Sales` đã làm giàu, bảng KPI SUMIFS và 3 KPI Definition Card.
- **Giới hạn dẫn sang B3:** Mỗi chiều phân tích mới lại phải viết lại SUMIFS.
- **Câu chuyện mở khóa (Phụ lục F):** #1 hé lộ. Bảng SUMIFS cho thấy T1/2025 (khoảng 426 triệu) cao gần gấp đôi T2 (khoảng 246 triệu). Để câu hỏi "vì sao?" mở, B3 sẽ trả lời.

### Hướng dẫn giảng dạy B2

**Khái niệm then chốt:**
- **KPI Definition Card:** tên, công thức, grain, dimensions. Viết card **trước** khi viết công thức.
- **Dimension / Measure / Grain:**
  - Dimension là "theo cái gì" (Vùng, Tháng).
  - Measure là "đo cái gì" (Doanh thu).
  - Grain là "mỗi dòng là gì" (một giao dịch).
- **Chuỗi C, lần 1:** mỗi tiêu chí của SUMIFS là một *bộ lọc*. Vẽ **hình xương sống**: một ô kết quả ← [Vùng = Bắc] + [Tháng = 3] ← phép tổng. Hình này dùng lại nguyên vẹn ở B3 và B5.
- **Chuỗi B:** khóa lookup phải duy nhất. Nếu khóa trùng, XLOOKUP trả về dòng đầu tiên *mà không báo lỗi*.

**Cách dạy:**
- Cài sẵn một mã sản phẩm bị trùng trong `Product`. Để học viên XLOOKUP ra giá vốn sai, rồi hỏi "vì sao biên lợi nhuận đẹp bất thường?" → rút ra Laws of Lookups. (Doanh thu không lệch vì `Unit_Price` lấy từ Sales; P021 cũ làm COGS Q1 thiếu 1.792.000, Gross Margin 34,4% thay vì 34,2%.)
- Đây cũng là bài học D2 đầu tiên: số hiển thị chưa chắc đúng.

**Lỗi thường gặp → cách xử lý:**

| Lỗi | Dấu hiệu | Xử lý |
|---|---|---|
| VLOOKUP để trống `range_lookup` (mặc định là khớp gần đúng) | Trả về giá của sản phẩm khác | Luôn dùng FALSE/0, hoặc chuyển sang XLOOKUP |
| Khóa text ≠ khóa số (`"001"` và `1`) | `#N/A` dù "nhìn thấy có" | TEXT/VALUE đồng bộ kiểu; TRIM khoảng trắng |
| Vùng tiêu chí SUMIFS lệch kích thước | `#VALUE!` | Dùng cột của Table (structured reference) |
| Trung bình của trung bình | AOV toàn công ty = AVERAGE(AOV từng vùng) | Tổng doanh thu ÷ tổng số đơn |
| Ngày lưu dạng text | Không group được theo tháng | DATEVALUE; kiểm tra căn phải |

**Kiểm tra cuối buổi:**
- Tổng bảng KPI SUMIFS = tổng cột Doanh thu.
- Không còn `#N/A`.
- Đủ 3 KPI Definition Card.
- **Lưu bảng SUMIFS lại:** đây là "đáp án" để đối soát ở B3 và B5.

**Điều chỉnh cho lớp đã biết Excel sơ bộ:** rút Thống kê mô tả từ 15 xuống 5 phút (chỉ giữ mean vs median và outlier); dùng 10 phút cho PRO TIP Dynamic Arrays (FILTER, SORT, UNIQUE).

**Chuẩn bị:**
- File khởi đầu (đáp án B1).
- Bảng `Product` có 1 khóa trùng cài sẵn: `1_Sessions/B2_Ham_va_KPI/Practice/B2_Product_DuplicateKey.xlsx` (P021 xuất hiện 2 lần, bản ghi cũ đứng trước).
- Mẫu KPI Definition Card (1 trang): `1_Sessions/B2_Ham_va_KPI/Handout/B2_KPI_Definition_Card.pdf`.
- Bài tập về nhà (bắt buộc), nộp trước B3: `Homework/B2_BTVN_MCI_Mart_T4_2025.xlsx` (làm giàu dữ liệu T4/2025, KPI Miền × Ngành hàng, card AOV, 1 câu insight); đáp án và sheet Kiem_tra ở file `_Answer`.

---

## Buổi 3 — Tư duy dữ liệu & PivotTable

**Câu hỏi BI mở buổi:** "Làm sao trả lời 20 câu hỏi 'theo chiều nào' trong 20 phút?" *(Tầng 2 — Dimensional)*

**Mục tiêu** — học viên có thể:
- Giải thích Pivot tính mỗi ô như thế nào: **tọa độ = bộ lọc** (**Chuỗi C, lần 2**).
- Dựng Pivot từ Excel Table B2, định dạng báo cáo dễ đọc.
- Group theo ngày/giá trị; lọc tương tác bằng Slicer và Timeline.
- Phân tích tỷ trọng và tăng trưởng bằng Show Values As; tạo Calculated Field.
- Tạo PivotChart và **chọn loại chart theo 3 câu hỏi then chốt** (preview B6). *(R1 §4.6)*

| Phút | Nội dung | Nguồn |
|---|---|---|
| 10 | Recap B2 (bảng KPI SUMIFS) + câu hỏi BI | |
| 10 | Vì sao cần Pivot; nhắc tidy data (B1) | Pivot PDF p2–5 |
| 20 | Insert Pivot, Field List, Analyze/Design, Refresh; **Pivot hoạt động thế nào** (so với SUMIFS) | Pivot PDF p6–11 |
| 10 | Number format, Style, Compact/Outline/Tabular | Pivot PDF p13–16 |
| 10 | *Nghỉ* | |
| 20 | Sort, Filter, Group; **Show Details / drill-down** (double-click vào ô giá trị để xem các dòng gốc); Slicer & Timeline | Pivot PDF p19–21 *(R4: PL-38, MO-48)* |
| 20 | Summarize Values By, Show Values As (% of total, % difference, running total); Calculated Field | Pivot PDF p24–26, 28 |
| 10 | PivotChart + **3 câu hỏi chọn chart** | Pivot PDF p34–36; DataViz PDF p5 |
| 30 | **Thực hành:** doanh thu Vùng × Tháng, % đóng góp theo danh mục, MoM bằng Show Values As, Slicer cửa hàng. **Đối chiếu một ô Pivot với ô SUMIFS của B2** | |
| 10 | Quiz + insight + preview B4 | |

- **Nâng cao (PRO TIP):** Conditional formatting trong Pivot (p17); custom sort lists, cached source data (Pro Tips – Pivot).
- **Handout:** Calculated Items, Solve Order, List Formulas, Report Filter Pages, Index calculation (p22, 27, 29–32), Pivot chart field buttons, **Data → Group/Outline và Subtotal** *(R4: MO-14, MO-15)*.
- **Brick output:** Báo cáo Pivot tương tác và bản nháp PivotChart.
- **Giới hạn dẫn sang B4:** Dữ liệu tháng mới về dạng CSV bẩn, tách file, nên lại phải copy-paste và sửa tay.
- **Câu chuyện mở khóa (Phụ lục F):**
  - #1 Tết: group theo ngày (7 ngày) thấy doanh thu tăng khoảng 3 tuần trước 29/01/2025 rồi giảm sâu trong kỳ nghỉ; % of total theo danh mục cho thấy Bánh kẹo tăng tỷ trọng (45,6% tháng 1, 30,9% tháng 2). Đồ uống tăng doanh thu nhưng tỷ trọng gần như không đổi (khoảng 11–12%).
  - #2 ST07: Pivot Cửa hàng × Tháng, ST07 thấp bất thường trong tháng 3 (−52,4% so với tháng 2). Double-click (Show Details) để thấy từ ngày 10/03 số đơn giảm còn một nửa, nhiều ngày không có đơn.
  - #6 Cuối tuần: Pivot theo cột Thứ, T7–CN cao hơn khoảng 25%.

### Hướng dẫn giảng dạy B3

**Khái niệm then chốt:**
- **Chuỗi C, lần 2:** mỗi ô giá trị của Pivot = một SUMIFS với tọa độ hàng/cột làm tiêu chí (Pivot PDF p11).
- **Pivot đọc từ cache**, nên phải Refresh sau khi dữ liệu nguồn thay đổi.
- **Show Values As** không tạo dữ liệu mới; nó tính lại trên kết quả đã tổng hợp.

**Cách dạy:**
- Đặt bảng SUMIFS của B2 cạnh Pivot vừa tạo. Chọn cùng một ô ở hai bên và chỉ ra tọa độ = tiêu chí.
- Double-click ô Pivot (Show Details) để thấy các dòng gốc tạo nên con số.
- Kết thúc bằng câu hỏi tình huống: "Khi nào vẫn nên dùng SUMIFS thay vì Pivot?" (bố cục tự do, công thức cần tham chiếu ô).

**Lỗi thường gặp → cách xử lý:**

| Lỗi | Dấu hiệu | Xử lý |
|---|---|---|
| Nguồn không phải Table | Dòng mới không vào Pivot sau Refresh | Nguồn Pivot = Table (B1) |
| Quên Refresh | Pivot không khớp dữ liệu | Refresh All; bật Refresh on open |
| Cột số có ô trống/text | Pivot mặc định dùng **Count** thay vì Sum | Sửa dữ liệu; đổi Summarize Values By |
| Không group được theo ngày | Báo "Cannot group that selection" | Cột ngày có giá trị text/trống → sửa nguồn |
| GETPIVOTDATA tự sinh khi tham chiếu ô Pivot | Kéo công thức không đổi kết quả | Tắt "Use GetPivotData functions" |
| Pie chart cho xu hướng theo tháng | Không đọc được xu hướng | 3 câu hỏi chọn chart → Line |

**Kiểm tra cuối buổi:**
- Tổng Pivot = tổng SUMIFS của B2.
- Slicer lọc được mọi Pivot liên quan.
- Mỗi PivotChart ghi rõ câu hỏi nó trả lời.

**Chuẩn bị:**
- File khởi đầu (đáp án B2): `1_Sessions/B2_Ham_va_KPI/Practice/B2_MCI_Mart_Q1_2025_Answer.xlsx`.
- File demo: `1_Sessions/B3_PivotTable/Demo/` (Pivot, Refresh, Loi_Du_lieu).
- Bài tập về nhà (bắt buộc), nộp trước B4: `Homework/B3_BTVN_MCI_Mart_T1_T4_2025.xlsx` (Pivot 4 tháng đầu năm 2025: đối soát với số liệu hệ thống, cài sẵn 6 ô Quantity dạng chữ; Miền × Tháng MoM, tỷ trọng ngành, ST07 hồi phục, AOV bằng Calculated Field, PivotChart + Slicer, 1 câu insight); đáp án và sheet Kiem_tra ở file `_Answer`.

---

## Buổi 4 — Tiền xử lý dữ liệu với Power Query

**Câu hỏi BI mở buổi:** "Làm sao để việc làm sạch dữ liệu chỉ làm **một lần**, từ đó chỉ cần bấm Refresh?" *(Nền móng cho các tầng)*

**Mục tiêu** — học viên có thể:
- Giải thích quy trình Power Excel: PQ → Data Model → Pivot.
- Làm sạch dữ liệu bằng Applied Steps: data type, header, text/date tools, conditional column.
- **Unpivot bảng `Target`** (file kế hoạch dạng cột tháng) về dạng tidy (**Chuỗi A**). *(R2 §2.2)*
- Append dữ liệu từ Folder; hiểu Merge như "XLOOKUP ở cấp bảng" (**Chuỗi B**).
- Tạo bảng Calendar.
- **Đối soát sau mỗi bước:** số dòng, tổng doanh thu so với nguồn. *(R2 §2.3)*

| Phút | Nội dung | Nguồn |
|---|---|---|
| 10 | Recap + demo "nỗi đau": 3 file CSV bẩn | |
| 10 | Quy trình Power Excel; khi nào dùng PQ/PP | PQ PDF p10–17 |
| 10 | Connections, Query Editor, Applied Steps | PQ PDF p19–22 |
| 25 | **Data profiling** (Column quality / distribution / profile); Transform: data type, header, text/date tools, conditional column; **Remove Duplicates** (khóa trùng sẽ làm hỏng relationship ở B5); **Keep/Remove Errors** | PQ PDF p24–27, 31 *(R4: PL-04, PL-06, MO-16)* |
| 10 | *Nghỉ* | |
| 15 | **Unpivot `Target`**; Group By | PQ PDF p32–33 |
| 20 | Append từ Folder; Merge (các kiểu join). **Merge chỉ để minh họa, các bảng lookup load riêng** *(R1 §4.5)* | PQ PDF p34–36 |
| 10 | Calendar table | PQ PDF p28 |
| 10 | Load options (**Connection only / Add to Data Model**) + **đối soát số dòng và tổng** | PQ PDF p23, 37 |
| 20 | **Thực hành:** thả tháng thứ 4 vào folder, bấm Refresh, Pivot B3 tự cập nhật. Ghi lại số dòng đối soát | |
| 10 | Quiz + insight + preview B5 | |

- **Nâng cao (PRO TIP):** Rolling calendar (p29); tham số hóa đường dẫn folder; **Reference vs Duplicate query** *(R4: PL-13)*.
- **Handout:** Index column (p30), number-specific tools (p26), modify workbook queries (p34), **xử lý sự cố privacy levels / lỗi Formula.Firewall khi Merge nhiều nguồn** (lỗi gần như chắc chắn gặp ở capstone), Pivot column và Transpose *(R4: PL-02, PL-10)*.
- **Brick output:** Các bảng sạch `Sales`, `Product`, `Store`, `Customer`, `Calendar`, `Target`, load riêng từng bảng vào Data Model và refresh được.
- **Câu chuyện mở khóa (Phụ lục F):** Sau khi Append 2024 + 2025, bảng đối soát số dòng theo Cửa hàng × Tháng cho thấy **#3 ST12** không có dữ liệu trước 06/2024 (cửa hàng mới, không phải lỗi mất dữ liệu). Calendar hai năm cho thấy Tết rơi vào tháng 2 năm 2024 nhưng tháng 1 năm 2025, là tiền đề của #4.
- **Giới hạn dẫn sang B5:** Các bảng rời nhau. Làm sao phân tích Sales và Target cùng lúc mà không merge thành một bảng khổng lồ?

### Hướng dẫn giảng dạy B4

**Khái niệm then chốt:**
- **Applied Steps** là "công thức làm sạch" được ghi lại, chạy lại mỗi lần Refresh.
- **Unpivot** đưa bảng dạng cột tháng về dạng tidy (Chuỗi A).
- **Merge** là "XLOOKUP ở cấp bảng" (Chuỗi B). **Append** là nối dọc.
- **Load Connection only / Data Model:** lookup được load riêng, để chuẩn bị cho B5.
- **Đối soát:** số dòng và tổng sau mỗi bước quan trọng.

**Cách dạy:**
- Làm tay việc gộp 3 file CSV trong 5 phút để học viên "thấy đau".
- Sau đó làm lại bằng Power Query.
- Cuối cùng thả file tháng 4 vào folder và bấm Refresh. Đây là **khoảnh khắc D3** của khóa học; hãy dành thời gian cho nó.

**Lỗi thường gặp → cách xử lý:**

| Lỗi | Dấu hiệu | Xử lý |
|---|---|---|
| Bước "Changed Type" ghi cứng tên cột | Refresh lỗi khi file mới đổi tên hoặc thêm cột | Đặt kiểu dữ liệu sau khi đã chuẩn hóa cột; kiểm tra Applied Steps |
| Tiêu đề của từng CSV lọt vào dữ liệu khi Append folder | Dòng có giá trị "Date", "Store"… trong cột dữ liệu | Filter bỏ dòng tiêu đề lặp |
| Ngày dd/mm/yyyy bị hiểu thành mm/dd (locale Việt Nam) | Ngày 13 trở lên thành lỗi; tháng và ngày bị đảo | Change Type → **Using Locale** → Vietnamese |
| Dấu thập phân phẩy/chấm | Số bị nhân 1.000 hoặc thành text | Using Locale; kiểm tra bằng column profile |
| Khoảng trắng thừa trong mã | Merge không khớp, ra null | Trim / Clean trước khi Merge |
| Merge với lookup có khóa trùng | **Số dòng tăng** sau Merge | Remove Duplicates ở lookup; đối soát số dòng sẽ bắt được lỗi này |
| Lỗi Formula.Firewall / privacy levels | Không Refresh được khi kết hợp nhiều nguồn | Xem handout xử lý sự cố |

**Kiểm tra cuối buổi:**
- Bảng đối soát: số dòng từng tháng và tổng doanh thu, khớp nguồn.
- Refresh với file tháng 4 chạy không lỗi.
- `Target` có dạng Tháng | Cửa hàng | Target (thêm cột ngày đầu tháng `Ngay_Dau_Thang` để B5 nối với Calendar).
- **D3 mong đợi:** Competent (pipeline tự refresh).

**Chuẩn bị:**
- `1_Sessions/B4_Power_Query/Practice/Data_Sales/`: `Sales_2024.csv` (22.141 dòng) và `Sales_2025-01/02/03.csv`, **đúng 4 file**. Các file 2025 cài sẵn lỗi: đơn giá có dấu chấm ngăn nghìn kiểu Việt Nam ("19.000"), ngày dd/mm/yyyy (mọi file), 25 mã cửa hàng có khoảng trắng thừa, 1 dòng trùng mỗi file.
- `1_Sessions/B4_Power_Query/Practice/B4_Lookups.xlsx`: Product (P015 bị trùng một dòng để luyện Remove Duplicates), Store, Customer.
- `1_Sessions/B4_Power_Query/Practice/B4_Target_2025.xlsx`: target dạng cột "Tháng 1…Tháng 12", tiêu đề ở dòng 3.
- `1_Sessions/B4_Power_Query/Practice/New_Month/Sales_2025-04.csv`: file tháng 4 để thả vào folder khi demo Refresh.
- File khởi đầu (đáp án B3): `1_Sessions/B3_PivotTable/Practice/B3_MCI_Mart_Q1_2025_Answer.xlsx`. Không có file `Demo/`: mọi demo chạy trên các file `Practice/` ở trên.
- Đáp án (file khởi đầu B5): `Practice/B4_MCI_Mart_Answer.xlsx` — tham số `DataFolder`, 6 bảng load riêng, Calendar 01/01/2024–31/12/2025 (731 ngày, phủ trọn hai năm vì Target có đủ 12 tháng 2025), Pivot `PV_Refresh` (Nam × Thang), `Kiem_tra_B4`.
- Bài tập về nhà (bắt buộc), nộp trước B5: `Homework/B4_BTVN_MCI_Mart_T4_2025.xlsx` + thư mục `Homework/B4_BTVN_Du_lieu/` (tháng 4/2025, mỗi cửa hàng một file CSV, Store_Code nằm trong tên file; cài sẵn dòng "Tổng cộng", mã sản phẩm sai dạng, dòng trùng, và file ST07 thiếu nửa tháng — phát hiện nhờ đối soát với số liệu hệ thống; % đạt target theo cửa hàng); đáp án và sheet Kiem_tra ở file `_Answer`.

---

## Buổi 5 — Mô hình dữ liệu & DAX: KPI dùng chung một nguồn sự thật

**Câu hỏi BI mở buổi:** "Doanh thu tháng này đạt bao nhiêu % mục tiêu, và vì sao mỗi báo cáo cũ lại ra một con số khác nhau?" *(Tầng 3 — Diagnostic)*

**Mục tiêu** — học viên có thể:
- Thiết kế **star schema**: Fact (`Sales`, `Target`) và Lookup; Primary/Foreign key; relationship thay cho merge.
- Nối **hai fact table khác grain** và biết chỉ lọc được qua lookup chung. *(R2 §2.2; PQ PDF p53–55)*
- Phân biệt Calculated Column và Measure; hiểu vì sao KPI dạng tỷ lệ phải là measure.
- Giải thích **filter context** (**Chuỗi C, lần 3**) và dùng CALCULATE, ALL.
- Viết time intelligence (YTD, YoY) với bảng Calendar.
- Xây **KPI dictionary** và **đối soát measure với bảng SUMIFS của B2**. *(R2 §2.1, §2.3)*

| Phút | Nội dung | Nguồn |
|---|---|---|
| 5 | Recap | |
| 20 | Data Model, Data vs Diagram view; normalization; Fact vs Lookup; keys; **Relationships vs Merged tables** | PQ PDF p39–45 |
| 15 | Tạo relationship; cardinality 1-*; filter direction; **fact thứ hai `Target`** | PQ PDF p46, 50, 53–55 |
| 10 | *Nghỉ* | |
| 15 | Calculated Column vs Measure; implicit vs explicit | PQ PDF p64–72, 76 |
| 15 | **Filter context** (dùng lại hình Chuỗi C) | PQ PDF p73–75 |
| 20 | CALCULATE, ALL → `% of Total`; `Revenue vs Target %`, `Variance` | PQ PDF p92–94, 99–100 |
| 15 | Time intelligence: YTD, YoY | PQ PDF p105–107 |
| 10 | **KPI dictionary:** measure là định nghĩa KPI dùng chung (semantic layer) | Phụ lục C |
| 20 | **Thực hành:** viết 6 KPI measure; **đối soát `Total Revenue` theo vùng với bảng SUMIFS của B2**; dựng lại báo cáo B3 bằng Power Pivot | |
| 5 | Quiz + insight + preview B6 | |

- **KPI của case** (định nghĩa đầy đủ ở **Phụ lục C**): Total Revenue, Total Orders, Avg Order Value, Gross Margin %, Revenue YoY Growth, Active Customers, cộng thêm Revenue vs Target %.
- **Nâng cao (PRO TIP):** RELATED (p101–102), X-functions SUMX (p103), SWITCH(TRUE) (p88–89), active/inactive relationship (p49).
- **Handout:** DAX text functions (p90–91), disconnected slicers (p96–97), hierarchies, hide fields (p56–57), lookup-to-lookup (p47), bad cardinality (p51–52), performance & best practices (p108–109): bỏ cột và dòng thừa, giảm granularity *(R4: PL-27, PL-28)*.
- **Brick output:** Star schema cùng KPI dictionary đã đối soát.
- **Câu chuyện mở khóa (Phụ lục F):**
  - #4 Target lệch mùa Tết: `Revenue vs Target %` T1/2025 khoảng 121%, T2 khoảng 83%. Nguyên nhân không nằm ở bán hàng mà ở cách lập target theo cùng tháng dương lịch.
  - #2 ST07 tháng 3 chỉ đạt khoảng 45% target.
  - #3 ST12: YoY trống T1–T5, nên phải so sánh like-for-like (chỉ cửa hàng có đủ hai kỳ).
  - #5 Gross Margin % 2025 cao hơn 2024. Đây là bài D2: kiểm tra giả định trước khi kết luận (dataset dùng một bảng giá vốn duy nhất).
- **Dẫn sang B6:** Số đã đúng và thống nhất. Tiếp theo là trình bày cho người ra quyết định.

### Hướng dẫn giảng dạy B5

**Khái niệm then chốt:**
- **Star schema:** fact ở giữa, lookup xung quanh, quan hệ 1-*. Bộ lọc chảy **từ lookup sang fact**.
- **Measure khác Calculated Column:**
  - Measure tính theo ngữ cảnh lọc.
  - Column tính theo từng dòng và được lưu lại.
  - KPI tỷ lệ phải là measure.
- **Chuỗi C, lần 3: filter context.** Dùng lại hình xương sống: "ô này đang bị lọc bởi những gì?". CALCULATE là cách *thay đổi* bộ lọc đó.
- **Hai fact khác grain:** `Target` (Tháng × Cửa hàng) chỉ lọc được qua `Calendar` và `Store`.

**Cách dạy:**
- Đưa `Target` vào Pivot cùng `Product` **trước** khi giải thích. Học viên sẽ thấy cùng một con số lặp ở mọi dòng, rồi tự hỏi vì sao → dạy lookup chung và grain.
- Kết thúc bằng đối soát: `Total Revenue` theo vùng phải khớp bảng SUMIFS của B2. Khớp nghĩa là B2 và B5 cùng nói một sự thật.

**Lỗi thường gặp → cách xử lý:**

| Lỗi | Dấu hiệu | Xử lý |
|---|---|---|
| Không có relationship / lọc qua bảng không liên quan | Cùng một giá trị lặp ở mọi dòng | Kiểm tra Diagram view; lọc bằng lookup chung |
| Khóa trùng ở lookup | Không tạo được quan hệ 1-*, hoặc thành nhiều-nhiều | Remove Duplicates ở B4 |
| Đặt cột ngày của fact vào hàng/cột thay vì Calendar | YTD/YoY sai hoặc trống | Luôn dùng cột của `Calendar`; Calendar phải liên tục |
| Implicit measure (kéo thả Sum) | Không dùng lại được trong CUBEVALUE/measure khác | Viết explicit measure |
| KPI tỷ lệ bằng calculated column | Tổng cộng các tỷ lệ | Measure dạng tổng ÷ tổng; dùng DIVIDE để tránh chia 0 |

**Kiểm tra cuối buổi:**
- 6 KPI khớp bảng SUMIFS của B2.
- Ảnh Diagram view là hình ngôi sao.
- AOV đúng khi lọc một danh mục (tính tay để so).

**Chuẩn bị:**
- File khởi đầu (đáp án B4): `1_Sessions/B4_Power_Query/Practice/B4_MCI_Mart_Answer.xlsx`, Data Model 6 bảng chưa có relationship. Không có file `Demo/`: mọi demo chạy trên bản sao file này.
- Đáp án (file khởi đầu B6): `Practice/B5_MCI_Mart_Answer.xlsx` — 6 relationship (hai fact Sales, Target nối qua Store và Calendar), 14 explicit measure (Total COGS bằng SUMX vì Excel không cho script tạo calculated column), `KPI_Dictionary`, `Kiem_tra_B5` (đối soát 64 ô với `KPI_SUMIFS`). Dựng bằng Excel trên Windows.
- Bài tập về nhà (bắt buộc), nộp trước B6, tập dượt C3 và C4 của capstone (Nguyên tắc 11): `Homework/B5_BTVN_MCI_Mart_KPI_T1_T4_2025.xlsx`. Tình huống: phòng Kinh doanh gửi báo cáo tháng 1–4/2025, số lệch với Tài chính. File có Data Model và sheet báo cáo dựng sẵn với 4 lỗi cài sẵn: Target chưa nối Calendar và báo cáo lấy tháng từ `Sales[Thang]` (% đạt target chỉ còn 6,7–11,6% thay vì 82,5–120,6%); GM% là trung bình của cột GM từng dòng (lệch 0,1–0,5 điểm); số đơn đếm bằng số dòng (AOV chỉ còn khoảng một nửa). Học viên đối soát với sheet số liệu Tài chính, ghi nhật ký lỗi và nguyên nhân, sửa thành mô hình ngôi sao, viết lại measure (thêm tăng trưởng like-for-like), điền KPI dictionary, viết 3 câu hỏi đo lường được cho Sales Director và 1 câu insight. Đáp án và sheet `Kiem_tra` ở file `_Answer`.

---

## Buổi 6 — Thiết kế & Xây dựng Dashboard

**Câu hỏi BI mở buổi:** "Giám đốc bán hàng có 5 giây. Họ cần thấy gì đầu tiên?" *(Tầng 4 — Prescriptive; nhập môn Tầng 5)*

**Mục tiêu** — học viên có thể:
- Viết **dashboard brief**: audience, 3–5 câu hỏi hoặc quyết định, tần suất refresh. *(R2 §2.4)*
- Áp dụng nguyên tắc rõ ràng và đơn giản; nhận biết biểu đồ gây hiểu lầm.
- Chọn chart bằng 3 câu hỏi then chốt và bảng chọn chart (8 loại lõi, gồm combo 2 trục).
- Bố cục theo thứ bậc (KPI → Trend → Detail), **tối đa 5–8 visual**. KPI card có **target và màu RAG**.
- Lấy mọi con số từ Data Model bằng **CUBEVALUE** và Pivot, không tính lại ở tầng hiển thị. *(R2 §2.5)*
- Trình bày insight theo **Situation → Complication → Resolution** và phát hành dashboard.
- **Dự báo ngắn hạn (Tầng 5, nhập môn):** so sánh Trendline tuyến tính với Forecast Sheet (ETS), đọc dải tin cậy và nêu giới hạn của dự báo.

| Phút | Nội dung | Nguồn |
|---|---|---|
| 5 | Recap: KPI dictionary B5 | |
| 10 | **Dashboard brief** | Phụ lục D |
| 10 | Clarity & simplicity; Good/Bad/Ugly (trục Y gây hiểu lầm) | DataViz PDF p1–4 |
| 20 | 3 Key Questions + 8 loại lõi: Line, Bar/Column, **Combo 2 trục** (Doanh thu cột + Margin % đường; Actual vs Target), Stacked bar/Donut/Treemap, Histogram/Box (phân phối, outlier), Scatter (tương quan ≠ nhân quả), Map, Waterfall | DataViz PDF p5–8, 10–11, 13–15, 20–21; Statistics PDF p34, 131 *(R4: MO-37, MO-38)* |
| 10 | *Nghỉ* | |
| 20 | Layout hierarchy; KPI card bằng **CUBEVALUE/CUBEMEMBER**; RAG bằng conditional formatting, **kèm biểu tượng hoặc nhãn (không chỉ dựa vào màu đỏ-xanh), alt text cho chart**; Slicer kết nối nhiều Pivot; Sparklines | Pro Tips (Visualization, Analytics) *(R4: PL-40)* |
| 10 | **Dự báo nhập môn (Tầng 5):** Trendline Linear + Forward trên chart doanh thu ngày → thấy đường thẳng bỏ qua nhịp cuối tuần; **Forecast Sheet (ETS)** trên bảng Ngày–Doanh thu, seasonality 7, dải tin cậy 95%; hàm `FORECAST.ETS`, `FORECAST.ETS.CONFINT`; **giới hạn**: 16 tháng chưa đủ 2 chu kỳ năm, Tết đổi ngày theo âm lịch | *mới* (từ `Curriculum_Guide.html`) *(R4: PL-42)* |
| 40 | **Build dashboard có hướng dẫn** trên case: hàng KPI card (có target/RAG), trend doanh thu, **combo Actual vs Target**, Top sản phẩm, map theo vùng, waterfall chênh lệch target | |
| 15 | Storytelling SCR, chú thích trên chart; **phát hành:** Refresh on open, ẩn sheet kỹ thuật, Protect Sheet, xuất PDF, chia sẻ OneDrive/SharePoint, lưu ý dữ liệu khách hàng và sensitivity label; nêu giới hạn: Excel không tự refresh theo lịch, cần Power BI *(R4: PL-46, PL-47)* | Phụ lục D |
| 10 | Quiz + insight + **giao brief và dataset capstone** | |

- **Nâng cao (PRO TIP):** Form Controls, dynamic ranges, custom number formats (Pro Tips – Visualization/Formatting); **Theme cho dashboard; Conditional Formatting Rules Manager** *(R4: PL-31, MO-19)*; `FORECAST.ETS.SEASONALITY`, `FORECAST.ETS.STAT`, tùy chọn *Data Completion* và *Aggregate Duplicates* của Forecast Sheet; sneak peek Power BI (PQ PDF p111–112).
- **Handout:** Area, Pie, Bubble, Radar, Surface/Contour, Stock, Funnel, Sunburst, Histogram/Pareto nâng cao (DataViz PDF p7, 9, 12, 16–19).
- **Brick output:** Dashboard bán hàng 1 trang, refresh được từ nguồn (B4 → B5 → B6), cùng một sheet dự báo 30 ngày tới có dải tin cậy.
- **Câu chuyện mở khóa (Phụ lục F):** #4 và #2 phải *nhìn thấy được* trên combo Actual vs Target, có chú thích; #6 là lý do ETS chọn seasonality 7; #1 là lý do không tin dự báo tháng Tết.

### Hướng dẫn giảng dạy B6

**Khái niệm then chốt:**
- **Brief trước, chart sau.**
- **Thứ bậc:** KPI → Trend → Detail.
- **Một nguồn sự thật:** mọi số lấy từ measure qua CUBEVALUE hoặc Pivot.
- **Mốc so sánh cho mọi KPI:** target, kỳ trước.
- **Insight theo SCR.**
- **Dự báo = mẫu hình quá khứ + giả định.** ETS học được nhịp lặp đều (cuối tuần); nó không biết sự kiện dịch chuyển (Tết). Luôn đọc dự báo cùng dải tin cậy và ghi giả định bên cạnh.

Quy tắc chi tiết: **Phụ lục D**.

**Cách dạy:**
- Mở bằng một dashboard "xấu" chuẩn bị sẵn: 12 visual, pie 3D, trục Y cắt khúc, không có target.
- Cho học viên làm **bài kiểm tra 5 giây** với nhau và liệt kê vấn đề → rút ra nguyên tắc → viết brief → build.
- Dùng output `dashboard_spec_generator.py` gợi ý line cho cả 6 KPI làm bài phản biện: vì sao dashboard tốt cần kết hợp card, trend và breakdown?
- **Dự báo:** vẽ Trendline trước, hỏi "đường này có tin được không?". Sau đó tạo Forecast Sheet trên cùng dữ liệu để học viên tự thấy răng cưa cuối tuần. Kết bằng câu tình huống: "Giám đốc hỏi doanh thu Tết 2026. Dùng được dự báo này không? Vì sao?" (không: chưa đủ 2 năm, và Tết 2026 rơi vào ngày khác).

**Lỗi thường gặp → cách xử lý:**

| Lỗi | Dấu hiệu | Xử lý |
|---|---|---|
| Quá nhiều visual | Người xem không biết nhìn đâu | Tối đa 5–8; bỏ visual không trả lời câu hỏi nào trong brief |
| Bar chart có trục Y không bắt đầu từ 0 | Chênh lệch nhỏ trông rất lớn | Trục bar/column bắt đầu từ 0 |
| Slicer không nối với mọi Pivot | Một phần dashboard không đổi khi lọc | Report Connections |
| Pivot co giãn cột khi Refresh, làm vỡ bố cục | Layout xô lệch | Tắt "Autofit column widths on update" |
| Gõ sai tên measure/member trong CUBEVALUE | `#N/A` | Dùng gợi ý autocomplete của CUBEMEMBER |
| RAG chỉ dùng đỏ-xanh | Người mù màu không đọc được | Thêm biểu tượng hoặc nhãn |
| Dùng Trendline tuyến tính cho dữ liệu có mùa vụ | Dự báo phẳng, bỏ qua đỉnh và đáy | Forecast Sheet (ETS) với seasonality phù hợp |
| Chạy ETS theo tháng với seasonality 12 khi chưa đủ 2 năm | Mùa vụ bị đoán sai, dải tin cậy rất rộng | Dự báo theo ngày (seasonality 7), hoặc chờ đủ dữ liệu |
| Chuỗi ngày bị thiếu hoặc trùng | Forecast Sheet báo lỗi hoặc tự nội suy âm thầm | Tạo bảng Ngày–Doanh thu bằng Group By trong Power Query, đối chiếu với Calendar |
| Trình bày một con số dự báo, không có khoảng | Người xem hiểu là cam kết | Luôn kèm dải tin cậy và giả định |

**Kiểm tra cuối buổi:**
- Học viên chấm chéo dashboard của nhau bằng checklist Phụ lục D.
- **Phát brief capstone, dataset và rubric bản Student-Friendly.**

**Chuẩn bị:**
- File khởi đầu (đáp án B5).
- Dashboard "xấu" mẫu.
- Brief, dataset capstone và `1_Sessions/B7_Capstone/Rubric_B7_Capstone.md` bản học viên.
- Bảng Ngày–Doanh thu 01/01/2024–30/04/2025 (query Group By từ B4, load ra sheet dạng Table) để làm Forecast Sheet.

---

## Buổi 7 — Ôn tập & Thu hoạch (Capstone)

**Mục tiêu:**
- Tự thực hiện dự án BI end-to-end trên **bộ dữ liệu mới**: Brief → Power Query → Data Model → KPI measure → Dashboard → Insight và đề xuất.
- Thể hiện D1–D4 và đạt tối thiểu Level 3 (Builder).

**Hình thức** *(R2 §3)*:
- Brief và dataset được **giao cuối B6**; học viên làm ở nhà, **cá nhân** (không làm nhóm).
- Brief theo mẫu "PROJECT BRIEF" của Statistics PDF (p41, 64, 99, 125, 155).

| Phút | Nội dung |
|---|---|
| 20 | Ôn tập: 3 chuỗi xương sống + mỗi buổi "3 điều cần nhớ"; **xếp 6 câu chuyện MCI Mart vào thang 5 tầng**; quiz tổng hợp |
| 40 | Hoàn thiện dự án, hỏi đáp, kiểm tra refresh và đối soát |
| 70 | Thuyết trình 5–7 phút mỗi học viên, kèm nhận xét theo rubric |
| 20 | Tổng kết khóa học, lộ trình tiếp theo (Power BI, thống kê suy luận) |

- **Rubric:** `1_Sessions/B7_Capstone/Rubric_B7_Capstone.md`.
  - **6 tiêu chí × 4 level:** C1 Brief, C2 Dữ liệu, C3 Mô hình, C4 KPI, C5 Dashboard, C6 Insight. D1–D4 được gắn vào các tiêu chí này.
  - **Level 3 – Builder:** không có tiêu chí nào ở Level 1, và C2 ≥ 3, C5 ≥ 3.
  - **Level 4 – Modeler:** đạt Builder, và C3 = 4, C4 = 4.
  - Phát bản Student-Friendly cùng brief ở cuối B6.
  - **Trong 40 phút hoàn thiện ở B7:** giảng viên đưa một file dữ liệu mới để học viên thử Refresh; kết quả dùng để chấm C2.

### Hướng dẫn giảng dạy B7

**Chuẩn bị:**
- File dữ liệu mới của dataset capstone, dùng cho bài kiểm tra Refresh (C2).
- 2 bài mẫu để hiệu chuẩn nếu có từ hai người chấm trở lên.
- Phiếu chấm theo 6 tiêu chí.

**Điểm quyết định khi chấm** (chi tiết: `Rubric_B7_Capstone.md`, mục Calibration Notes):
- **C2:** chấm bằng Refresh thật, không chấm bằng lời kể.
- **C4:** lọc một danh mục và tính tay AOV/Margin % để so.
- **C3 và C4 chấm độc lập:** không trừ hai lần cho cùng một lỗi gốc.
- **C6:** hỏi "vậy thì sao?". Nguyên nhân chưa kiểm chứng trên dashboard vẫn chỉ đạt Level 2.

---

## Bonus (ngoài 7 buổi, tùy chọn)

- **Thống kê suy luận:** phân phối chuẩn, CLT, khoảng tin cậy, kiểm định giả thuyết (Statistics PDF p42–125); hồi quy tuyến tính (p126–154).
- **Analytics nâng cao:** Goal Seek, Scenario Manager, Solver, Monte Carlo, phát hiện outlier (Pro Tips – Analytics).
- **Power Query/DAX nâng cao:** dữ liệu bán cấu trúc (JSON), semi-additive measure (số dư tồn kho/tài khoản) *(R4: PL-11, PL-26)*.

## Ngoài phạm vi & lộ trình chứng chỉ *(R4)*

Khóa học tập trung vào Excel **cho BI**. Các mục sau thuộc chuẩn Microsoft nhưng **không** được dạy, và cần nêu rõ với học viên ngay từ B1.

**Học viên muốn thi MO-211 (Excel Expert)** cần tự học thêm:
- Macro: ghi, đặt tên, sửa, copy, bật macro.
- Hàm tài chính: PMT, NPER.
- Consolidate. Trong khóa này, việc gộp dữ liệu được thay bằng Append trong Power Query.
- Quản lý phiên bản workbook, Restrict Editing, Protect Workbook Structure.

**Học viên muốn thi PL-300 (Power BI Data Analyst)** cần học thêm:
- Power BI Desktop/Service: workspace, app, gateway, scheduled refresh, row-level security.
- Bookmarks, drillthrough page, Copilot, calculation groups.

Phần Power Query, Data Model và DAX của khóa này chuyển thẳng sang Power BI được.

Mức độ bao phủ hiện tại (theo R4, trước v2.1):

| Chuẩn | Direct | Partial | Indirect | None |
|---|---|---|---|---|
| MO-211 | 17 / 48 | 12 / 48 | 7 / 48 | 12 / 48 |
| PL-300 (phần làm được bằng Excel) | 25 / 47 | 12 / 47 | 5 / 47 | 5 / 47 |

## Phụ lục A — Ba chuỗi xương sống (chi tiết)

Ký hiệu:
- **[cứng]:** không nắm khái niệm trước thì không học được khái niệm sau.
- **[mềm]:** học được, nhưng khó hơn.

**Chuỗi A — Cấu trúc dữ liệu**
```
Cell / Range / Reference (B1)
  ↓ [cứng]
Tidy data + Excel Table (B1)
  ↓ [cứng]
Cột tính toán trên Table — bảng làm giàu (B2)
  ↓ [cứng]
Dimension / Measure / Grain (B2 gọi tên → B3 dùng trong Pivot)
  ↓ [cứng]
Unpivot / Group By đưa dữ liệu về dạng tidy (B4)
  ↓ [cứng]
Fact vs Lookup, normalization, star schema; fact thứ hai khác grain (B5)
```

**Chuỗi B — Nối bảng**
```
Tham chiếu tuyệt đối $ (B1)
  ↓ [cứng]
VLOOKUP → INDEX/MATCH → XLOOKUP; khóa phải duy nhất (B2)
  ↓ [cứng]
Merge = lookup ở cấp bảng; Append = nối dọc; Remove Duplicates (B4)
  ↓ [cứng]
Primary/Foreign key → Relationship → Cardinality 1-* → Filter direction (B5)
  ↓ [mềm]
RELATED() trong DAX (B5, PRO TIP)
```

**Chuỗi C — Tổng hợp có lọc** (khái niệm khó nhất khóa học)
```
SUM / AVERAGE (B1–B2)
  ↓ [cứng]
SUMIFS: criteria = bộ lọc (B2)
  ↓ [cứng]
Ô Pivot = SUMIFS với tọa độ hàng/cột làm tiêu chí (B3; Pivot PDF p11)
  ↓ [cứng]
Show Values As / Calculated Field: tính trên kết quả đã tổng hợp (B3)
  ↓ [cứng]
Measure + filter context (B5; PQ PDF p73–75)
  ↓ [cứng]
CALCULATE / ALL: thay đổi bộ lọc có chủ đích (B5)
  ↓ [mềm]
Time intelligence YTD / YoY, cần Calendar từ B4 (B5)
  ↓ [mềm]
KPI card bằng CUBEVALUE (B6)
```

**Chuỗi phụ:**
- **Ngày tháng:** Date value (B2) → Group theo ngày (B3) → Calendar (B4) → Time intelligence (B5) → Timeline (B6).
- **Thống kê:** Mean/Median/StDev (B2) → Histogram, Box plot (B6) → Correlation (B6) → Trendline, Forecast ETS (B6) → Regression (Bonus). Đây là chuỗi mềm.

---

## Phụ lục B — Bốn năng lực thực hành D1–D4

Mỗi mức mô tả **hành vi quan sát được**. Giảng viên dùng làm từ vựng chung khi góp ý ở mọi buổi; rubric B7 chấm các năng lực này qua sản phẩm.

**D1 — Bắt đầu từ câu hỏi kinh doanh** (chấm ở rubric C1)

| Mức | Hành vi |
|---|---|
| Emerging | Mở file và kéo thả ngay. Mô tả kết quả bằng tên công cụ ("em đã làm Pivot"). |
| Developing | Viết câu hỏi trước khi làm, nhưng câu hỏi còn là chủ đề ("xem doanh thu"). |
| Competent | Câu hỏi đo lường được: có chỉ số, chiều, thời gian, mốc so sánh. Chọn công cụ theo câu hỏi. |
| Extending | Đặt câu hỏi "vì sao" tiếp theo từ kết quả; ưu tiên câu hỏi theo tác động kinh doanh. |

**D2 — Hoài nghi và đối soát dữ liệu** (chấm ở rubric C2, C4)

| Mức | Hành vi |
|---|---|
| Emerging | Tin mọi con số hiển thị; không nhận ra dòng trùng, ô trống, số dạng text. |
| Developing | Phát hiện lỗi khi được nhắc; dùng Data Validation, IFERROR. |
| Competent | Chủ động đối chiếu tổng Pivot/measure với nguồn, số dòng trước và sau Merge; giải thích outlier trước khi báo cáo. |
| Extending | Thiết kế bước kiểm tra tự động; ghi chú giả định dữ liệu trong báo cáo. |

**D3 — Tự động hóa, tái sử dụng** (chấm ở rubric C2)

| Mức | Hành vi |
|---|---|
| Emerging | Copy-paste và sửa tay mỗi kỳ báo cáo. |
| Developing | Dùng Excel Table và công thức tự mở rộng; biết Refresh. |
| Competent | Pipeline PQ → Model → Measure; thêm dữ liệu mới mà không phải sửa gì. |
| Extending | Tham số hóa; đặt tên và tổ chức query/measure để người khác bảo trì được. |

**D4 — Truyền đạt insight** (chấm ở rubric C5, C6)

| Mức | Hành vi |
|---|---|
| Emerging | Đọc lại con số trên bảng hoặc biểu đồ. |
| Developing | Nêu nhận xét tăng/giảm cho từng biểu đồ. |
| Competent | Dashboard có thứ bậc; mỗi insight gồm con số có so sánh, nguyên nhân khả dĩ và hàm ý (SCR). |
| Extending | Kể câu chuyện dẫn đến đề xuất cụ thể; lường trước câu hỏi phản biện; nêu giới hạn của dữ liệu. |

**Mức mong đợi theo buổi:**
- D2 đạt Developing sau B1–B2 và Competent sau B4–B5.
- D3 đạt Competent sau B4.
- D1 và D4 đạt Competent sau B6.

D2 và D3 chỉ có thể đạt Competent sau khi học xong Power Query (B4).

---

## Phụ lục C — KPI dictionary của case

Target đã được chỉnh theo quy mô dataset MCI Mart (2025: doanh thu khoảng 250–430 triệu/tháng, khoảng 820–1.400 đơn/tháng). Target doanh thu lấy từ bảng `Target`, không dùng hằng số.

| KPI | Công thức (khái niệm) | Aggregation | Grain | Owner | Dimensions | Target | Warning / Critical |
|---|---|---|---|---|---|---|---|
| Total Revenue | SUM(Quantity × Unit_Price) | sum | tháng | Sales Director | region, store, category, month | = SUM(Target) (285–354 triệu/tháng) | 90% / 80% |
| Total Orders | DISTINCTCOUNT(Order_ID) | count_distinct | tháng | Sales Director | region, store, month | 1.000 đơn/tháng | 90% / 80% |
| Avg Order Value | Revenue ÷ Orders | ratio | tháng | Sales Director | region, store, month | 300.000 | 90% / 80% |
| Gross Margin % | (Revenue − COGS) ÷ Revenue | ratio | tháng | Finance | category, region, month | 33% | 90% / 80% |
| Revenue YoY Growth | (Revenue − Revenue LY) ÷ Revenue LY | rate | tháng | Finance | region, category, month | +10% | 50% / 20% |
| Active Customers | DISTINCTCOUNT(Customer_ID) | count_distinct | tháng | Marketing | region, segment, month | 700 KH/tháng | 90% / 80% |
| Revenue vs Target % | Total Revenue ÷ SUM(Target) | ratio | Tháng × Cửa hàng | Sales Director | store, region, month (**không** theo product/category) | 100% | 90% / 80% |

**Ghi chú giảng dạy:**
- KPI dạng **ratio/rate** không được cộng dồn hay lấy trung bình của trung bình, nên luôn phải là measure (PQ PDF p68).
- **Revenue vs Target %** chỉ lọc được theo chiều mà `Target` có. Đây là ví dụ về grain mismatch ở B5.
- **Mẫu KPI Definition Card** (dùng từ B2): Tên · Mục đích · Công thức · Nguồn dữ liệu · Grain · Dimensions · Owner · Target · Ngưỡng Warning/Critical · Ghi chú giới hạn.

---

## Phụ lục D — Quy tắc dashboard & checklist chấm chéo

**Thứ tự thiết kế:** Audience → Câu hỏi/quyết định → KPI → Layout → Chart → Build → Story → Phát hành.

**Dashboard brief:**
- Người xem: executive, operational hay self-service.
- 3–5 câu hỏi đo lường được.
- Tần suất refresh.
- Quyết định mà dashboard hỗ trợ.

**Bảng chọn chart (8 loại lõi):**

| Câu hỏi dữ liệu | Chart | Thay thế |
|---|---|---|
| Xu hướng theo thời gian | Line | Area |
| So sánh giữa các nhóm | Bar / Column | Bullet |
| Cơ cấu / tỷ trọng | Stacked bar | Donut (≤ 5 phần), Treemap |
| Hai thước đo khác đơn vị, hoặc Actual vs Target | **Combo 2 trục** | Column + đường target |
| Phân phối, outlier | Histogram | Box & Whisker |
| Quan hệ giữa hai biến | Scatter (+ trendline) | Bubble |
| Địa lý | Map (choropleth) | Bar theo vùng |
| Chênh lệch cộng dồn (target → actual) | Waterfall | — |

**Checklist chấm chéo** (dùng ở B6 và để học viên tự kiểm tra trước B7):
- [ ] Mỗi visual trả lời một câu hỏi trong brief; không có visual "cho đẹp".
- [ ] KPI quan trọng nhất nằm góc trên trái; thứ tự KPI → Trend → Detail.
- [ ] 5–8 visual trên một trang.
- [ ] Mọi KPI có mốc so sánh (target hoặc kỳ trước) và trạng thái RAG, **kèm biểu tượng hoặc nhãn**.
- [ ] Mọi con số lấy từ measure (CUBEVALUE/Pivot); không có công thức KPI ở tầng hiển thị.
- [ ] Trục bar/column bắt đầu từ 0; không 3D; không pie quá 5 phần.
- [ ] Slicer/Timeline điều khiển mọi visual (Report Connections).
- [ ] Có chú thích trên chart chỉ vào điểm bất thường.
- [ ] Bài kiểm tra 5 giây: người lạ thấy ngay KPI chính và biết đạt hay chưa đạt target.
- [ ] Nếu có dự báo: hiển thị dải tin cậy và ghi rõ giả định (dữ liệu dùng, mùa vụ, sự kiện không được tính).
- [ ] Phát hành: Refresh on open, sheet kỹ thuật được ẩn, Protect Sheet, alt text, xuất PDF thử.

**Storytelling SCR:**
- **Situation:** con số có so sánh.
- **Complication:** vấn đề và nguyên nhân đã kiểm chứng bằng drill-down.
- **Resolution:** đề xuất hành động cho đúng người xem trong brief.

---

## Phụ lục E — Hệ thống tài liệu

Cấu trúc thư mục (chi tiết: `README.md` ở thư mục gốc):

```
MCI Excel/
├── README.md, Course_Plan.md
├── 1_Sessions/B1…B7/     tài liệu từng buổi (kịch bản slide, Handout/, Practice/)
├── 2_Dataset/            dataset gốc MCI Mart + _generator/
├── 3_Design_Reviews/     hồ sơ lý do (R1, R2, R4)
└── 4_Reference/          PDF nguồn và bảng tra
```

| File | Vai trò |
|---|---|
| `Course_Plan.md` | **Tài liệu gốc**: nội dung, thời lượng, cách dạy, phụ lục. Đủ để soạn deck và dạy. |
| `1_Sessions/B<n>_…/` | Tài liệu từng buổi: kịch bản slide ở gốc thư mục buổi, `Handout/` (handout, checklist), `Practice/` (file thực hành, đáp án). |
| `1_Sessions/B1_Nen_tang_Excel_Du_lieu/Slide_Script_B1.md` | Kịch bản từng slide của B1 (41 slide): nội dung, ảnh cần chụp, lời giảng, quiz, tài nguyên. Là khuôn mẫu cho B2–B7. |
| `1_Sessions/B2_Ham_va_KPI/Slide_Script_B2.md` | Kịch bản B2 (41 slide, 150 phút); deck, demo, practice, homework, handout B2 dựng bằng `build_b2.py`, `build_slides_b2.py`, `build_docs_b2.py`. |
| `1_Sessions/B3_PivotTable/Slide_Script_B3.md` | Kịch bản B3 (42 slide, 150 phút); demo, practice, homework dựng bằng `build_b3.py` (PivotTable, Slicer, Timeline, PivotChart ghi bằng `pivot_xml.py`, rồi mở và lưu bằng Excel); deck, handout bằng `build_slides_b3.py`, `build_docs_b3.py`. |
| `1_Sessions/B4_Power_Query/Slide_Script_B4.md` | Kịch bản B4 (39 slide, 150 phút), dạy trên file `Practice/` (không có Demo); số liệu, bài về nhà bằng `build_b4_homework.py`; đáp án có Power Query thật bằng `build_b4.py` (chỉ Windows, Excel qua COM); deck, handout bằng `build_slides_b4.py`, `build_docs_b4.py`. |
| `1_Sessions/B5_Data_Model_DAX/Slide_Script_B5.md` | Kịch bản B5 (39 slide, 150 phút), **bản nháp chờ duyệt**; dạy trên file khởi đầu (đáp án B4), không có Demo; số liệu bằng `build_b5_meta.py`. Đáp án có Data Model và measure thật sẽ dựng bằng `build_b5.py` (chỉ Windows). |
| `1_Sessions/B7_Capstone/Rubric_B7_Capstone.md` | Rubric đầy đủ 6 tiêu chí × 4 level, bản học viên, calibration notes. |
| `2_Dataset/` | Dataset MCI Mart (`MCI_Mart_Master.xlsx`), từ điển dữ liệu và câu chuyện cài sẵn (`README.md`), mã sinh dữ liệu và tài liệu (`_generator/`). |
| `3_Design_Reviews/Review_01_Knowledge_Architecture.md` | *Lý do* của thứ tự buổi và các chuỗi xương sống. |
| `3_Design_Reviews/Review_02_BI_Practice.md` | *Lý do* của các thay đổi B3–B6 theo thực tiễn BI. |
| `3_Design_Reviews/Review_04_Coverage_Audit.md` + `Coverage_Audit.csv` | Đối chiếu với chuẩn MO-211 / PL-300. |
| `4_Reference/Excel_PowerBI_Reference_Knowledge.md` | Bảng tra hàm và tính năng Excel / Power Query / Power Pivot / DAX, tổng hợp từ 5 PDF nguồn (không gồm Statistics PDF). Dùng khi soạn slide và handout. |
| `4_Reference/Curriculum_Guide.html` | Bản nháp trang web giới thiệu khóa, viết cho case khác (VinaSweetCo) và thứ tự buổi khác. **Chưa đồng bộ với plan**; các phần tốt đã đưa vào v3.1. |

Các file review là **hồ sơ lý do**, chỉ cần đọc khi muốn hiểu *vì sao*. Khi chỉnh plan, cập nhật `Course_Plan.md` trước.

## Phụ lục F — Câu chuyện dữ liệu theo buổi và theo tầng

Nguồn: mục "Câu chuyện cài sẵn trong dữ liệu" của `2_Dataset/README.md`. Ý tưởng "mỗi buổi mở khóa một câu chuyện" lấy từ bản nháp `Curriculum_Guide.html`.

| # | Câu chuyện | Hé lộ | Giải thích | Tầng | Công cụ phát hiện |
|---|---|---|---|---|---|
| 1 | Tết: tăng khoảng 3 tuần trước Tết, giảm sâu trong kỳ nghỉ; Bánh kẹo tăng tỷ trọng (Đồ uống tăng doanh thu, tỷ trọng gần như giữ nguyên) | B2 | B3 | 1 → 2 | SUMIFS theo tháng → Pivot group theo ngày, % of total |
| 2 | ST07 Nha Trang cải tạo 10–31/03/2025 | B3 | B5 | 2 → 3 | Pivot Cửa hàng × Tháng, Show Details → Revenue vs Target % |
| 3 | ST12 Biên Hòa khai trương 01/06/2024, tăng dần 3 tháng đầu | B4 | B5 | 2 → 3 | Đối soát số dòng Cửa hàng × Tháng → YoY trống, like-for-like |
| 4 | Target lập theo cùng tháng 2024 nên lệch mùa Tết (T1 ~121%, T2 ~83%) | B4 | B5–B6 | 3 → 4 | Calendar 2 năm → measure vs Target → combo Actual vs Target |
| 5 | Gross Margin % 2025 cao hơn 2024 (do giả định một bảng giá vốn) | B5 | B5 | 3 | Measure Gross Margin % theo năm; kiểm tra giả định (D2) |
| 6 | Cuối tuần cao hơn khoảng 25% | B3 | B6 | 2 → 5 | Pivot theo Thứ → seasonality 7 trong Forecast Sheet |

**Gợi ý dùng:**
- Không nói trước đáp án. Câu hỏi mở buổi dẫn tới câu chuyện; câu insight đóng buổi là nơi học viên tự gọi tên nó.
- Tầng 4 (nên làm gì) chỉ xuất hiện khi có người ra quyết định, nên hầu hết câu chuyện chỉ lên Tầng 4 ở B6. Ví dụ với #4: đề xuất lập target 2026 theo lịch âm cho các tháng quanh Tết.
- Dataset capstone (B7) nên có 3–4 câu chuyện tương tự, để chấm được C6 (Insight).

## Việc cần chốt

- [x] Dataset xuyên suốt: MCI Mart (giả lập), đã tạo ở `2_Dataset/MCI_Mart_Master.xlsx`.
- [ ] Dataset riêng cho capstone (công ty khác).
- [x] Target trong KPI dictionary (Phụ lục C) đã chỉnh theo dataset.
- [ ] Mỗi buổi có file Excel thực hành (bài tập + đáp án) hay không.
- [x] Phương án cho học viên dùng Mac: khuyến nghị máy ảo (Parallels Desktop hoặc VMware Fusion); dự phòng là học viên tự mang laptop Windows khác (trung tâm không có phòng lab).
- [x] Dựng lại deck B1 từ đầu (deck S1 cũ đã xóa), theo mục "Style spec": `Slide_B1.pptx` + bản học viên `Slide_B1_HocVien.pptx`.
- [ ] Dựng lại `Curriculum_Guide.html` theo plan v3.1 (case MCI Mart, đúng thứ tự buổi), hoặc bỏ file này.
- [ ] Kiểm tra Forecast Sheet trên bảng Ngày–Doanh thu MCI Mart khi soạn deck B6: seasonality tự phát hiện có ra 7 không, dải tin cậy rộng bao nhiêu.
- [ ] Mỗi buổi xuất infographic tóm tắt `Handout/B<n>_Infographic_TomTat.png` (nguồn `_generator/infographic_b<n>.html`, dựng bằng `build_infographic.py`). Đã có: B1, B2, B3, B4. Còn: B5–B7.
- [ ] Soạn handout cho mọi mục "Handout" đã liệt kê. Coverage của R4 chỉ được tính khi handout thực sự tồn tại, kèm ảnh minh họa.

## Changelog v2 (2026-09-23)

| Thay đổi | Nguồn |
|---|---|
| Thêm yêu cầu Windows, kiểm tra máy ở B1, sửa slide phiên bản S1 | R2 §1 |
| Mục tiêu đầu ra theo thang Self-service BI maturity | R2 §2.7 |
| Chuyển tidy data + Excel Table từ B3 về B1; rút tour Ribbon thành handout | R1 §4.2, §4.8 |
| Dimension/Measure/Grain và KPI Definition Card từ B2 | R1 §4.3, R2 §2.1 |
| Ba chuỗi xương sống + hình minh họa Chuỗi C dùng lại ở B2/B3/B5 | R1 §2, §4.4 |
| Thêm bảng `Target`: Unpivot (B4), fact thứ hai (B5), RAG/variance (B6) | R2 §2.2 |
| B4: merge chỉ để minh họa, load lookup riêng; đối soát số dòng/tổng | R1 §4.5, R2 §2.3 |
| B5: cắt về lõi; KPI dictionary; đối soát measure với SUMIFS B2 | R1 §4.7, R2 §2.1, §2.3 |
| B6: tái cấu trúc Brief → Layout → Chart → Build → Story → Publish; CUBEVALUE; 7 loại chart lõi; regression chuyển sang Bonus | R2 §2.4–2.6 |
| B3: preview 3 câu hỏi chọn chart | R1 §4.6 |
| Câu hỏi mở và insight đóng mỗi buổi; 2 câu tình huống trong quiz | R1 §4.9, §5 |
| Phân bổ 150 phút cho mọi buổi; tách Lõi / Nâng cao / Handout | R2 §3 |
| B7: capstone giao từ cuối B6, làm ở nhà; B7 dành cho hoàn thiện và thuyết trình | R2 §3 |

## Changelog v2.1 (2026-09-23)

| Thay đổi | Nguồn |
|---|---|
| Bỏ tham chiếu tới deck S1 đã xóa; lưu Style spec ở cuối file | — |
| B1: thêm Watch Window; Handout thêm Flash Fill, Fill Series, tham chiếu sang workbook khác | R4: MO-31, MO-09, MO-10, MO-02 |
| B2 Handout: TODAY/NOW, WEEKDAY, WORKDAY, SWITCH, MAXIFS/MINIFS, RANDARRAY | R4: MO-11, MO-20, MO-22, MO-23 |
| B3: Show Details / drill-down vào lõi; Handout thêm Group/Outline, Subtotal | R4: PL-38, MO-48, MO-14, MO-15 |
| B4: Transform 20' → 25' (thêm data profiling, Remove Duplicates, Keep/Remove Errors); Editor 15' → 10'; PRO TIP Reference vs Duplicate; Handout privacy levels, Pivot column/Transpose | R4: PL-04, PL-06, MO-16, PL-13, PL-02, PL-10 |
| B5 Handout: bỏ cột/dòng thừa, giảm granularity | R4: PL-27, PL-28 |
| B6: combo 2 trục vào chart lõi (8 loại) và vào bài build; RAG có biểu tượng/nhãn + alt text; sensitivity label, giới hạn refresh; PRO TIP Theme, CF Rules Manager, Forecast Sheet | R4: MO-37, MO-38, PL-40, PL-46, PL-47, PL-31, MO-19, PL-42 |
| Bonus: Scenario Manager, JSON, semi-additive measure | R4: MO-25, PL-11, PL-26 |
| Mục mới "Ngoài phạm vi & lộ trình chứng chỉ" | R4 |

## Changelog v3 (2026-09-23)

| Thay đổi | Lý do |
|---|---|
| Mỗi buổi B1–B7 có mục **Hướng dẫn giảng dạy**: khái niệm then chốt, cách dạy, lỗi thường gặp → cách xử lý, kiểm tra cuối buổi, chuẩn bị | Plan trước đó kém chi tiết hơn các file review |
| Nguyên tắc mới: **file khởi đầu** mỗi buổi (lưới an toàn cho brick) | Học viên bị tụt một buổi không làm gãy chuỗi brick |
| Phụ lục A–E: chuỗi xương sống chi tiết, bảng D1–D4, KPI dictionary, quy tắc và checklist dashboard, hệ thống tài liệu | Để plan dùng được độc lập |

## Changelog v3.1 (2026-09-24)

Đưa các phần tốt của bản nháp `Curriculum_Guide.html` vào plan, chuyển sang case MCI Mart. Tổng thời lượng mỗi buổi vẫn là 150 phút.

| Thay đổi | Nguồn |
|---|---|
| Nguyên tắc 9: **thang 5 tầng câu hỏi** (Descriptive → Dimensional → Diagnostic → Prescriptive → Predictive); câu hỏi mở mỗi buổi được gắn tầng | Curriculum_Guide: "5 Tầng Analytics" |
| Nguyên tắc 10 + **Phụ lục F**: câu chuyện dữ liệu mở khóa theo buổi; mục "Câu chuyện mở khóa" ở B2–B6; B7 ôn tập bằng cách xếp câu chuyện vào 5 tầng | Curriculum_Guide: "BI Insight / Story Unlocked"; `2_Dataset/README.md` |
| B6: **Dự báo nhập môn 10'** (Trendline vs Forecast Sheet ETS, FORECAST.ETS.CONFINT, giới hạn Tết âm lịch và 16 tháng dữ liệu). Chart lõi 25' → 20', Build 45' → 40'. Forecast Sheet chuyển từ PRO TIP vào lõi; thêm 4 dòng lỗi thường gặp, checklist dự báo ở Phụ lục D | Curriculum_Guide B06–B07; R4: PL-42 (partial → hướng tới direct) |
| B2 thêm cột Thứ (phục vụ câu chuyện cuối tuần ở B3 và seasonality ở B6) | Phụ lục F |
| B1, B2: mục "Điều chỉnh cho lớp đã biết Excel sơ bộ" | Curriculum_Guide: ghi chú 📌 ở B01–B02 |

**Không đưa vào** (vì plan đã làm tốt hơn, hoặc không hợp với case): làm sạch bằng công thức thay cho Power Query; OFFSET/INDIRECT vào lõi; nhồi Power Query và DAX vào B7; các con số dự báo mẫu của VinaSweetCo.

## Changelog v3.2 (2026-09-24)

| Thay đổi | Lý do |
|---|---|
| Yêu cầu đầu vào: thông báo yêu cầu Windows **trước khóa học** và nhắc lại phút đầu B1; học viên Mac nên có máy Windows từ B2, hạn chót trước B4 | Để học viên chuẩn bị máy từ sớm, không bị động giữa khóa |
| Sửa "không học được B5–B7" thành "từ B4 đến B7", vì B4 đã load vào Data Model | Thống nhất với nội dung B4 |
| Sắp xếp lại thư mục theo cấu trúc khóa học (`1_Sessions/`, `2_Dataset/`, `3_Design_Reviews/`, `4_Reference/`); cập nhật đường dẫn và Phụ lục E | Dễ tìm tài liệu theo buổi |

## Changelog v3.3 (2026-09-27)

| Thay đổi | Lý do |
|---|---|
| Checklist kiểm tra máy gửi học viên **tự làm trước B1**; không đạt thì liên hệ trợ giảng trong lớp | Không mất giờ học cho việc học viên làm được ở nhà |
| B1: "Bài tự luyện" (tùy chọn) đổi thành **Bài tập về nhà bắt buộc**, có file đề + đáp án | Bài tự luyện không bắt buộc thì ít người làm; B2 cần học viên vững quy trình B1 |
| Tài liệu phát cho học viên có bản riêng hậu tố `_HocVien`: slide không có speaker notes; file demo không có câu chỉ dẫn dành cho giảng viên | Giữ uy tín giảng viên; notes là tài liệu nội bộ |
| B7: dự án capstone làm **cá nhân**, không làm nhóm (Hình thức B7, Rubric mục 8) | Chốt hình thức capstone |
| Học viên dùng Mac: khuyến nghị cài máy ảo Parallels Desktop hoặc VMware Fusion (checklist, handout, slide 7) | Chốt phương án cho Mac |
| B1: phần Phiên bản & kiểm tra máy 15 → 10 phút (slide 8 chỉ còn xác nhận); Thực hành 30 → 35 phút (Thực hành 1: 15 → 20 phút) | Dồn thời gian cho phần học viên hay chậm nhất |
| (2026-09-28) Bỏ phiếu ghi nhận kiểm tra máy và danh sách gửi sau B1. Học viên tự chuẩn bị máy; trung tâm không cung cấp phòng lab. Tài liệu không dùng chữ "ban tổ chức" | Đây là lớp học, không phải sự kiện; máy móc là trách nhiệm của học viên |
| (2026-09-28) B3: soạn đủ tài liệu (kịch bản, deck, demo, đáp án có Pivot/Slicer/Timeline/PivotChart, bài tập về nhà bắt buộc, handout, infographic). Sửa câu chuyện #1: dữ liệu cho thấy Bánh kẹo tăng tỷ trọng quanh Tết, Đồ uống tăng doanh thu nhưng tỷ trọng gần như không đổi | Kiểm tra lại trên dataset khi soạn B3 |
| (2026-10-01) Nguyên tắc 11: bài tập về nhà là bước tập dượt cho capstone, mỗi bài gắn với tiêu chí rubric B7 | Bài về nhà các buổi phải dẫn tới bài capstone |
| (2026-10-01) B5: bài tập về nhà bắt buộc — kiểm toán báo cáo cũ của phòng Kinh doanh (4 lỗi cài sẵn), sửa mô hình, viết lại measure, KPI dictionary; tập dượt C3, C4 | Câu hỏi mở B5 "vì sao mỗi báo cáo cũ lại ra một con số khác nhau"; Nguyên tắc 11 |

## Style spec (trích từ deck S1 trước khi bị xóa)

- **Khổ trang:** 16:9, 20 × 11.25 in. Layout Blank; mọi thành phần là shape hoặc group vẽ tay.
- **Khung slide nội dung:**
  - Thanh dọc bên trái, rộng khoảng 1.6 in, tên phần xoay dọc.
  - Logo nhỏ ở góc dưới trái.
  - Slide công cụ có banner ảnh ở đầu trang.
- **Slide chia phần:** "MAGIC CODE INSTITUTE / MCI / `<n>. <tên phần>`". Slide bìa: "MCI / Consulting & Analytics / Excel for Business Intelligence Analyst". Deck cũ ghi sai thành "Consuting"; không lặp lại lỗi này.
- **Font và màu chữ:** chỉ dùng Arial. Màu chữ `000000`, `7E7E7E`, `585858`.
- **Màu nhấn:**
  - Tên công cụ: `C00000` (đỏ), cỡ 24 pt, giải thích tiếng Việt ở dòng dưới.
  - "PRO TIP:": `FFC000`.
  - Nhấn phụ: `3398DA` và `203864`.
  - "TIP:": cỡ 18 pt.
- **Slide công cụ:** Tên → mục đích 1 dòng → ảnh minh họa, chú thích "(Trong ảnh minh họa: …)" → TIP. Tối đa 3 công cụ cạnh nhau, ngăn bằng đường kẻ dọc mảnh.
- **Ngôn ngữ:** tiếng Việt. Tên hàm và tên giao diện Excel giữ tiếng Anh, kèm nghĩa tiếng Việt trong ngoặc khi cần.
- **Thành phần cố định mỗi deck:**
  - "THỰC HÀNH n / <bài> / (10 phút)"
  - slide "Tổng Kết – Điểm chính cần nhớ:"
  - "EXCEL QUIZ CHALLENGE" → "QUESTION n: TOPIC", lựa chọn A)–D)
