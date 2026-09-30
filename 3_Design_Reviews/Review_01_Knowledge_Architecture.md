# Knowledge Architecture Analysis: Excel for Business Intelligence Analyst

> Review #1 được thực hiện bằng skill `curriculum-knowledge-architecture-designer` (Bernstein, Muller, Maton). Đầu vào là `Course_Plan.md`, ngày 2026-09-23.

- **Input type:** course (7 buổi: 6 bài học + 1 capstone)
- **Learner stage:** người đi làm hoặc sinh viên năm cuối, trình độ hỗn hợp (từ basic đến advanced), học bằng tiếng Việt
- **Learning goals (tóm tắt):**
  - Dùng Excel như một nền tảng BI: công thức → Pivot → Power Query → Data Model/DAX → Dashboard.
  - Luôn bắt đầu từ câu hỏi kinh doanh, không phải từ menu công cụ.
  - Tự thực hiện một dự án BI end-to-end và trình bày insight.

---

## 1. Epistemic Diagnosis

**Architecture type:** Mixed, thiên về Hierarchical.

**Proportions:** 50% Hierarchical · 20% Horizontal · 30% Dispositional

**Reasoning**

**Hierarchical (50%)** gồm phần lớn nội dung kỹ thuật. Có ba dấu hiệu cho thấy đây là cấu trúc phân tầng:

- **Có chuỗi tiên quyết cứng.** Không thể viết SUMIFS đúng nếu chưa hiểu tham chiếu tuyệt đối/tương đối (B1). Không thể hiểu Merge (B4) và Relationship (B5) nếu chưa hiểu "khóa duy nhất" trong Laws of Lookups (B2). Không thể hiểu filter context của DAX (B5) nếu chưa hiểu cách Pivot tính theo tọa độ ô (B3).
- **Lỗi ở tầng dưới lan lên tầng trên.** Học viên không nắm "mỗi dòng là một bản ghi" (tidy data) sẽ gặp lỗi ở SUMIFS, ở Pivot, ở Unpivot và ở star schema.
- **Thứ tự các buổi gần như không thể đảo.** Đây là *conceptual coherence* theo Muller (2009).

**Horizontal (20%)** gồm các nội dung là "lăng kính" có nhiều cách tiếp cận hợp lệ, không có thứ tự tiên quyết cố định:

- Chọn loại biểu đồ theo "3 câu hỏi then chốt" (B6).
- Nguyên tắc trực quan hóa Good / Bad / Ugly.
- Thống kê như một lăng kính đọc dữ liệu: mean hay median, phân phối, tương quan (B2, B6).
- **Chọn công cụ cho một câu hỏi:** cùng một KPI có thể giải bằng SUMIFS, Pivot, Power Query Group By hoặc DAX measure. Năng lực "biết khi nào dùng cái nào" tăng dần qua các buổi, nhưng không có đáp án đúng duy nhất.

**Dispositional (30%)** gồm các năng lực mà khóa học tuyên bố là *mục tiêu cốt lõi* ("BI trước, Excel sau"). Chúng chỉ tồn tại khi được thực hành:

- Đặt câu hỏi kinh doanh trước khi mở công cụ.
- Hoài nghi dữ liệu: kiểm tra chất lượng, nguồn gốc, outlier.
- Tư duy tự động hóa và tái sử dụng ("làm một lần, Refresh mãi mãi").
- Truyền đạt insight bằng ngôn ngữ kinh doanh cho người ra quyết định.

Tỷ lệ 30% phản ánh *mục tiêu tuyên bố*. Trong `Course_Plan.md` hiện tại, năng lực dispositional chỉ được đánh giá ở B7 (xem mục 3 và 5).

---

## 2. Knowledge Architecture Map

### Hierarchical Elements

Khóa học có **ba chuỗi xương sống**. Mỗi chuỗi đi xuyên qua nhiều buổi. Đây là "brick" thực sự: brick không phải là *buổi học* mà là *khái niệm* được nâng cấp qua từng buổi.

**Chuỗi A — Cấu trúc dữ liệu (Data shape)**

```
Cell / Range / Reference (B1)
    ↓ [hard]
Tidy data: 1 dòng = 1 bản ghi, 1 cột = 1 trường; Excel Table   ← HIỆN DẠY Ở B3 (Pivot p5), CẦN ĐƯA VỀ B1
    ↓ [hard]
Cột tính toán trên bảng: enriched table (B2)
    ↓ [hard]
Dimension vs Measure; grain của bảng (B3)                       ← nên được GỌI TÊN từ B2
    ↓ [hard]
Unpivot / Group By để đưa dữ liệu về tidy (B4)
    ↓ [hard]
Fact vs Lookup table, normalization, star schema (B5)
```

**Chuỗi B — Nối bảng (Join)**

```
Tham chiếu tuyệt đối $ (B1)
    ↓ [hard]
VLOOKUP → INDEX/MATCH → XLOOKUP; "Laws of Lookups": khóa phải duy nhất (B2)
    ↓ [hard]
Merge Queries = lookup ở cấp bảng; Append = nối dọc (B4)
    ↓ [hard]
Primary/Foreign key → Relationship → Cardinality (1-*, *-*) → Filter direction (B5)
    ↓ [soft]
RELATED() trong DAX (B5)
```

**Chuỗi C — Tổng hợp có điều kiện (Aggregation → Filter context).** Đây là khái niệm có mật độ ngữ nghĩa cao nhất của khóa học.

```
SUM / AVERAGE (B1–B2)
    ↓ [hard]
SUMIFS: "criteria" = bộ lọc áp lên phép tổng (B2)
    ↓ [hard]
Pivot: mỗi ô giá trị = SUMIFS với tọa độ hàng/cột làm criteria (B3; Pivot PDF p11 "How do pivots actually work?")
    ↓ [hard]
Show Values As / Calculated Field: tính trên kết quả đã tổng hợp (B3)
    ↓ [hard]
DAX Measure + Filter context: "cell coordinates" (B5; PQ PDF p73–75)
    ↓ [hard]
CALCULATE / ALL: thay đổi bộ lọc có chủ đích (B5)
    ↓ [soft]
Time intelligence: YTD, YoY (B5, cần Calendar từ B4)
```

**Nhận xét:** Chuỗi C là lý do chính khiến B5 khó. Hai nguồn PDF đã có sẵn "cây cầu":

- Pivot PDF p11 giải thích Pivot hoạt động bằng cách *detect coordinates → filter → aggregate*.
- PQ PDF p74 giải thích filter context bằng *cell coordinates*.

Nếu B2, B3 và B5 dùng **cùng một hình minh họa** ("tiêu chí SUMIFS = tọa độ ô Pivot = filter context") thì học viên chỉ phải học một khái niệm, được đóng gói lại ba lần. Hiện `Course_Plan.md` chưa làm tường minh điểm này.

**Chuỗi phụ (độc lập, hội tụ muộn)**

- **Date:** Date value (B2) → Group theo ngày trong Pivot (B3) → Calendar table (B4) → Time intelligence (B5) → Timeline slicer (B6).
- **Thống kê:** Mean/Median/StDev (B2) → Histogram, Box plot (B6) → Correlation, Regression (B6). Đây là chuỗi *soft*, có thể đảo một phần.

### Horizontal Elements

**Hub 1 — "Câu hỏi kinh doanh này nên trả lời bằng công cụ nào?"**

```
                     ┌── Công thức (SUMIFS, XLOOKUP, FILTER): linh hoạt, từng ô
                     ├── PivotTable: khám phá nhanh, nhiều chiều
Câu hỏi KPI ─────────├── Power Query: làm sạch và lặp lại
                     ├── DAX measure: định nghĩa một lần, dùng mọi nơi
                     └── Biểu đồ / Dashboard: truyền đạt
```

**Hub 2 — "Dữ liệu này nói gì?" (lăng kính phân tích)**

```
                     ┌── Xu hướng theo thời gian (line, YoY)
                     ├── Cơ cấu / tỷ trọng (% of total, treemap)
Bộ dữ liệu bán hàng ─├── So sánh / xếp hạng (bar, RANK, Top N)
                     ├── Phân phối & ngoại lai (histogram, box plot, mean vs median)
                     └── Quan hệ (scatter, correlation, regression; tương quan ≠ nhân quả)
```

**Độ tinh vi phân tích phát triển qua các buổi**

| Mức | Học viên làm gì |
|---|---|
| Single-tool | "Tôi dùng SUMIFS để tính doanh thu theo vùng." |
| Multi-tool | "Có thể dùng SUMIFS hoặc Pivot; Pivot nhanh hơn khi cần xem nhiều chiều." |
| Integrative | "Nên làm sạch ở Power Query, định nghĩa `Revenue` bằng measure, rồi mới vẽ. Nếu tính trong biểu đồ thì mỗi báo cáo sẽ ra một con số khác nhau." |
| Systemic | Thiết kế pipeline sao cho câu hỏi mới chỉ cần thêm một measure hoặc một visual, không phải làm lại từ đầu. |

### Dispositional Elements

**D1 — Tư duy bắt đầu từ câu hỏi kinh doanh**

| Mức | Hành vi quan sát được |
|---|---|
| Emerging | Mở file và bắt đầu kéo thả, tạo công thức ngay. Mô tả kết quả bằng tên công cụ ("em đã làm Pivot"). |
| Developing | Viết được câu hỏi trước khi làm, nhưng câu hỏi còn chung chung ("xem doanh thu"). |
| Competent | Viết câu hỏi cụ thể, đo lường được ("Vùng nào giảm doanh thu MoM > 10% trong Q3?"), rồi chọn công cụ phù hợp. |
| Extending | Tự đặt câu hỏi tiếp theo từ kết quả ("Giảm do số đơn hay giá trị đơn?"), xây cây câu hỏi và ưu tiên theo tác động kinh doanh. |

**D2 — Hoài nghi và kiểm soát chất lượng dữ liệu**

| Mức | Hành vi quan sát được |
|---|---|
| Emerging | Tin mọi con số hiển thị. Không nhận ra dòng trùng, ô trống, text-số. |
| Developing | Phát hiện lỗi khi được nhắc; dùng Data Validation và IFERROR. |
| Competent | Chủ động đối chiếu tổng: tổng Pivot = tổng nguồn, số dòng trước và sau Merge. Giải thích được outlier trước khi báo cáo. |
| Extending | Thiết kế bước kiểm tra tự động (query kiểm tra, measure đối soát) và ghi chú giả định dữ liệu trong báo cáo. |

**D3 — Tư duy tự động hóa và tái sử dụng**

| Mức | Hành vi quan sát được |
|---|---|
| Emerging | Copy-paste và sửa tay mỗi kỳ báo cáo. |
| Developing | Dùng Excel Table và công thức tự mở rộng; biết bấm Refresh. |
| Competent | Xây pipeline PQ → Model → Measure; thêm dữ liệu tháng mới không cần sửa gì. |
| Extending | Tham số hóa (folder, ngày cắt), đặt tên và tổ chức query/measure để người khác bảo trì được. |

**D4 — Truyền đạt insight**

| Mức | Hành vi quan sát được |
|---|---|
| Emerging | Trình bày bảng số hoặc biểu đồ, đọc lại con số. |
| Developing | Nêu được một nhận xét cho mỗi biểu đồ ("doanh thu tăng"). |
| Competent | Dashboard có thứ bậc (KPI → xu hướng → chi tiết). Mỗi insight gồm con số, nguyên nhân khả dĩ và hàm ý. |
| Extending | Kể câu chuyện dẫn đến đề xuất hành động cụ thể, lường trước câu hỏi của người ra quyết định và nêu giới hạn của dữ liệu. |

D1–D4 là **knowledge-contingent dispositions**: không thể "hoài nghi dữ liệu" nếu không biết dữ liệu đúng trông như thế nào (Chuỗi A), và không thể "tự động hóa" nếu chưa có PQ (B4). Vì vậy D2 và D3 chỉ thực sự lên được mức Competent từ B4 trở đi.

---

## 3. Mixed Architecture Notes

### Kiểm định liên kết brick giữa các buổi (Hierarchical)

| Liên kết | Độ bền | Lý do / lỗ hổng |
|---|---|---|
| B1 → B2 | **Mạnh**, có 1 lỗ hổng | Syntax và reference là nền cho mọi hàm. Lỗ hổng: **tidy data và Excel Table (Ctrl+T) chưa có ở B1**, trong khi B2 đã làm cột tính toán và SUMIFS trên bảng. |
| B2 → B3 | **Trung bình – mạnh** | Động lực "SUMIFS phải viết lại cho mỗi chiều → cần Pivot" tốt. Lỗ hổng: khái niệm dimension/measure và grain xuất hiện lần đầu ở B3, nên B2 chưa có từ vựng để nói "bảng KPI theo vùng". |
| B3 → B4 | **Trung bình** | Động lực "dữ liệu mới, bẩn, tách file" chỉ là động lực *vận hành*, chưa phải *khái niệm*. Có thể gia cố: Unpivot (B4) sửa đúng loại "bad data structure" đã thấy ở B3 p5. |
| B4 → B5 | **Mạnh**, có 1 mâu thuẫn | PQ PDF p45 "Relationships vs Merged tables" là cầu nối hoàn hảo. Mâu thuẫn: thực hành B4 **Merge Product/Store vào Sales**, rồi B5 lại dạy **không nên merge**. Học viên học một cách rồi phải bỏ nó đi. |
| B5 → B6 | **Mạnh** | Measure trở thành KPI card. Lỗ hổng mềm: PivotChart đã xuất hiện ở B3 nhưng nguyên tắc chọn biểu đồ đến B6 mới dạy, nên học viên dựng biểu đồ "xấu" suốt ba buổi. |
| B6 → B7 | **Mạnh** | Capstone dùng dataset mới, kiểm tra được *transfer* thay vì học thuộc. |

### Tương tác giữa các loại tri thức

- **Hierarchical × Horizontal.** Hub "chọn công cụ nào" chỉ có nghĩa khi học viên đã có ít nhất hai công cụ để so sánh. Vì vậy câu hỏi so sánh công cụ nên bắt đầu từ B3 (SUMIFS vs Pivot) và lặp lại ở mỗi buổi sau, chứ không chờ đến B7.
- **Hierarchical × Dispositional.** D3 (tự động hóa) phụ thuộc cứng vào B4. Nếu B1–B3 không gieo mầm (Excel Table, tổng tự mở rộng), khoảnh khắc "Refresh là xong" ở B4 mất đi tác động.
- **Rủi ro tải nhận thức với lớp hỗn hợp:**
  - B2 dồn 84 trang PDF vào 9 mục.
  - B5 dồn 72 trang vào 10 mục và chứa khái niệm khó nhất (filter context).
  - Người mới dễ quá tải ở B2 và B5, trong khi người đã biết Excel sẽ chán ở B1 (tour giao diện dài khoảng 30 slide).
- **Căng thẳng chính.** Khóa học tuyên bố "BI trước, Excel sau", nhưng cả 6 quiz và phần lớn thực hành đo kỹ năng *công cụ* (hierarchical). Dispositional chỉ được chấm ở B7. Đúng như cảnh báo của khung phân tích: phần dễ đo sẽ chiếm lĩnh đánh giá.
- **Semantic waves (Maton).** B1 "phẳng ở mức cụ thể": nhiều ảnh giao diện, ít nguyên lý. B5 có nguy cơ "phẳng ở mức trừu tượng": DAX syntax nhiều, ít ví dụ gắn với case. Mỗi buổi nên có ít nhất một vòng *ví dụ cụ thể → nguyên lý → áp dụng lại vào case*.

---

## 4. Teaching & Sequencing Implications

1. **Giữ nguyên thứ tự B1→B6.** Chuỗi A, B, C đều ràng buộc cứng, và thứ tự của slide 2 trong S1 là đúng về mặt tri thức.
2. **Chuyển "Tidy data + Excel Table" từ B3 về B1.** Đây là thay đổi quan trọng nhất. Thêm 2–3 slide: Good vs Bad structure (Pivot PDF p5), Ctrl+T, structured reference. B3 chỉ nhắc lại.
3. **Gọi tên "Dimension / Measure / Grain" ngay ở B2** khi dựng bảng KPI bằng SUMIFS. B3 và B5 sẽ dùng lại đúng từ vựng này.
4. **Tạo một hình minh họa xương sống cho Chuỗi C**, dùng lại ở B2, B3 và B5: một ô kết quả → các tiêu chí lọc → phép tổng. Ở B5 chỉ cần nói "đây chính là filter context".
5. **Sửa mâu thuẫn B4 ↔ B5.** Trong thực hành B4, *Merge chỉ để minh họa*. Các bảng Product/Store/Calendar nên được load riêng ("Connection only / Add to Data Model") để B5 nối bằng relationship. Như vậy B4 → B5 trở thành *nâng cấp*, không phải *phủ định*.
6. **Preview nguyên tắc biểu đồ ở B3.** Chỉ cần 1 slide "3 câu hỏi then chốt" (DataViz PDF p5) khi dạy PivotChart. B6 đi sâu.
7. **Phân tầng rõ Nền tảng / Ứng dụng / Nâng cao cho B2 và B5**, để phục vụ lớp hỗn hợp:
   - B2: Dynamic Arrays, OFFSET, INDIRECT → *Nâng cao*.
   - B5: SUMX/X-functions, disconnected slicer, solve order → *PRO TIP*.
   - Lõi bắt buộc của B5 chỉ gồm: relationship, measure, filter context, CALCULATE, ALL, YTD/YoY.
8. **B1 cho người đã biết:** gom tour 8 tab thành bài tự học hoặc handout, dành thời gian lớp cho khung "Quy trình BI" và giới thiệu case.
9. **Dispositional là sợi chỉ liên tục.** Mỗi buổi mở bằng một câu hỏi kinh doanh trên case và đóng bằng một câu insight học viên tự viết. Không gán D1–D4 cho một buổi riêng.

---

## 5. Assessment Implications

**Auto-assessable elements (có thể chấm tự động)**

- Kiểm tra công thức trong file Excel nộp bài (B1–B2): kết quả ô, có dùng `$` đúng không, không có lỗi `#N/A`/`#REF!`. Có thể đối chiếu bằng script `openpyxl` với file đáp án.
- Pivot và measure cho ra đúng con số (B3, B5), so với đáp án.
- Query PQ load đúng số dòng và kiểu dữ liệu; Data Model có đúng relationship và cardinality (B4–B5).
- Quiz trắc nghiệm hiện có: đo tốt Chuỗi A, B, C ở mức nhớ và hiểu.

**Teacher-judgment elements (cần giảng viên chấm)**

- D1–D4: chất lượng câu hỏi kinh doanh, sự chủ động kiểm tra dữ liệu, mức độ tự động hóa, chất lượng insight và đề xuất. Cần quan sát xuyên suốt và chấm bài thuyết trình B7.
- Thiết kế dashboard ở mức Integrative/Systemic: đúng về kỹ thuật chưa đủ, phải *truyền đạt được*.

**Mixed elements (tự động một phần)**

- Chọn biểu đồ (B6): máy kiểm tra được loại biểu đồ có thuộc nhóm hợp lệ không; tính phù hợp với thông điệp cần người chấm.
- Câu insight cuối buổi: máy kiểm tra được có con số, có so sánh, có hàm ý; độ sâu cần người chấm.

**Đề xuất bổ sung vào quiz mỗi buổi:** 2 câu *tình huống* thuộc Hub horizontal, dạng "Câu hỏi X nên giải bằng công cụ nào? Vì sao?", để quiz không chỉ đo kiến thức công cụ. Rubric B7 nên dùng D1–D4 (bước 3 của lộ trình review: `criterion-referenced-rubric-generator`).

---

## 6. AI Tutoring Design Implications

Áp dụng nếu dùng Claude hoặc chatbot làm trợ giảng thực hành:

- **Hierarchical:** trợ giảng nên *kiểm tra tiên quyết* trước khi giải bài.
  - Ví dụ: học viên hỏi về CALCULATE thì hỏi lại "Ô này trong Pivot đang bị lọc bởi những gì?" (Chuỗi C).
  - Ví dụ: học viên lỗi Merge thì kiểm tra khóa có duy nhất không (Chuỗi B).
  - Phản hồi đúng/sai ngay lập tức, bài tập thích ứng theo độ khó.
- **Horizontal:** trợ giảng *gợi lăng kính*, không áp đáp án. Ví dụ: "Bạn đã xem xu hướng; thử nhìn theo cơ cấu, vùng nào đóng góp phần lớn mức giảm?" Hoặc "Nếu làm bằng Pivot thay vì SUMIFS thì được và mất gì?"
- **Dispositional:** trợ giảng **không chấm điểm** D1–D4. Chỉ làm ba việc:
  - hỏi phản tư cuối buổi ("Câu hỏi kinh doanh hôm nay của bạn là gì? Con số nào làm bạn nghi ngờ?");
  - tổng hợp câu trả lời cho giảng viên;
  - dùng bảng mức độ ở mục 2 làm từ vựng chung khi góp ý.
- **Chuyển chế độ:** trong một bài thực hành, học viên đi từ "viết measure đúng" (sửa lỗi, có đáp án), sang "chọn cách trình bày" (gợi mở), rồi "insight là gì" (phản tư). Trợ giảng phải đổi giọng tương ứng.

---

## Tóm tắt khuyến nghị (ưu tiên giảm dần)

1. Đưa **tidy data + Excel Table** về B1.
2. Làm tường minh **Chuỗi C** (SUMIFS → Pivot → Filter context) bằng một hình minh họa dùng lại ở B2, B3, B5.
3. Sửa mâu thuẫn **Merge (B4) ↔ Relationship (B5)** trong thực hành.
4. Gọi tên **Dimension / Measure / Grain** từ B2.
5. Giảm tải B2 và B5 bằng cách tách rõ lõi và phần Nâng cao/PRO TIP.
6. Thêm **sợi chỉ dispositional** (câu hỏi mở buổi, insight đóng buổi) và câu hỏi tình huống vào quiz.
7. Preview "3 câu hỏi chọn biểu đồ" ở B3.

*Giới hạn của phân tích:* tỷ lệ 50/20/30 là ước lượng diễn giải. Phân tích dựa trên kế hoạch *tuyên bố* trong `Course_Plan.md`; chất lượng thực tế còn phụ thuộc cách giảng dạy.
