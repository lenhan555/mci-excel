# Review #2 — Đối chiếu B3–B6 với thực tiễn BI

> Thực hiện bằng skill `business-intelligence` (KPI framework, dashboard design, reporting automation, data storytelling). Đầu vào: `Course_Plan.md` và `Review_01_Knowledge_Architecture.md`. Ngày: 2026-09-23.
> Ràng buộc: 7 buổi × **2.5 giờ (150 phút)**.

## Giả định

1. **Đối tượng dashboard trong case:** cấp quản lý bán hàng. Đây là dashboard *operational/executive*, cập nhật theo tháng.
2. **Nền tảng:** Excel 365 hoặc 2016+ trên Windows. Xem rủi ro Mac ở mục 1.
3. **Mỗi buổi 150 phút gồm:**
   - Khoảng 25–30 phút cố định: recap 5–10, nghỉ giải lao 10, quiz và preview 10.
   - Khoảng 30–45 phút thực hành.
   - Chỉ còn **khoảng 75–90 phút giảng thực sự**.

---

## 0. Kết luận nhanh

| Buổi | Mức sát thực tiễn BI | Vừa 150 phút? | Vấn đề chính |
|---|---|---|---|
| B3 Pivot | Tốt | Vừa, nếu cắt 4 mục | Thiếu "KPI là gì" trước khi pivot |
| B4 Power Query | Rất tốt | Vừa | Thiếu bước **đối soát dữ liệu** (reconciliation) |
| B5 Power Pivot/DAX | Tốt về kỹ thuật, **thiếu lớp KPI** | **Quá tải** (72 trang PDF) | Không có Target/Budget; measure chưa được định nghĩa như "KPI" |
| B6 Dashboard | **Yếu nhất** so với thực tiễn | **Quá tải** (16 loại chart + hồi quy + build) | Bắt đầu từ chart thay vì từ **audience/brief**; không có target/RAG; không có storytelling, không có phát hành |

---

## 1. Rủi ro chặn: Power Pivot không chạy trên Mac

`PowerQuery_PowerPivot_DAX.pdf` p5 ghi: *"Power Pivot is currently not available for Mac, and is only available in certain versions of Excel for Windows/PC"*.

Deck S1 có slide "Common Mac Shortcuts" và giới thiệu cả Excel Online, nên rất có thể lớp có học viên dùng Mac. Excel Online cũng không có Power Pivot, và Power Query trên Excel Online rất hạn chế.

**Hậu quả:**
- Học viên dùng Mac **không làm được B5**.
- B4 bị hạn chế một phần: Excel for Mac có Power Query nhưng thiếu một số connector.
- B6 và B7 phụ thuộc B5, nên cũng không làm được.

**Đề xuất:**
- Đưa yêu cầu "Excel 365 / 2016+ **Windows**, 64-bit" vào thông tin tuyển sinh.
- Kiểm tra máy học viên ở cuối B1 (menu Data → Get Data; File → Options → Add-ins → COM: Power Pivot).
- Có phương án phòng lab Windows hoặc máy ảo.
- Sửa slide S1 (slide 5–8, "chọn phiên bản") để nói rõ: *Excel Online và Mac không đủ cho khóa này.*

---

## 2. Khoảng trống so với thực tiễn BI

### 2.1 Thiếu **KPI framework**

Trong BI thực tế, mỗi KPI có định nghĩa chuẩn gồm:
- công thức (formula)
- nguồn dữ liệu (data source)
- mức chi tiết (granularity)
- chiều phân tích (dimensions)
- người chịu trách nhiệm (owner)
- mục tiêu (target)
- ngưỡng cảnh báo RAG (warning/critical)
- ghi chú giới hạn (caveats)

`Course_Plan.md` có nhắc "KPI" ở B2 và "measure" ở B5, nhưng **chưa bao giờ dạy cách định nghĩa một KPI**. Hệ quả là học viên biết *tính* nhưng không biết *định nghĩa*.

**Đề xuất:**
- **B2:** thêm 1 slide "KPI Definition Card" (công thức, grain, dimensions) trước khi viết SUMIFS.
- **B5:** biến bộ measure thành **KPI dictionary** của case, dạy "measure = định nghĩa KPI dùng chung cho mọi báo cáo" (semantic layer). Xem mẫu ở Phụ lục A.
  - Mẫu này đã được kiểm tra bằng `metric_validator.py --strict`: **PASS**, 6 KPI, 0 lỗi.
  - Lần chạy đầu phát hiện ngưỡng `critical_pct = 0` của YoY không hợp lệ. Có thể dùng chính lỗi này làm ví dụ "định nghĩa KPI cũng cần QA".

### 2.2 Thiếu **Target / Budget và phân tích chênh lệch (Actual vs Target)**

So sánh thực tế với kế hoạch là câu hỏi BI phổ biến nhất trong doanh nghiệp, nhưng case hiện chỉ có dữ liệu thực tế.

**Đề xuất:** thêm bảng `Target` (mục tiêu theo **Tháng × Cửa hàng**) vào dataset case.
- **B4:** Unpivot bảng Target. File kế hoạch thường có dạng cột tháng, tức là "bad structure", nên đây là ví dụ Unpivot tự nhiên nhất.
- **B5:** thêm một fact table thứ hai, grain khác với Sales. PQ PDF p53–55 (*Connecting multiple data tables*, *Filter direction*) dạy đúng điều này: chỉ lọc qua các lookup chung (`Calendar`, `Store`), không lọc được theo `Product`. Đây là bài học về **grain mismatch** rất thực tế. Measure mới: `Revenue vs Target %`, `Variance`.
- **B6:** KPI card có target và màu RAG bằng conditional formatting. Waterfall hiển thị chênh lệch.

### 2.3 Thiếu bước **đối soát** (validation / reconciliation)

Nguyên tắc BI: "Số trên dashboard phải khớp với nguồn". Kế hoạch hiện chưa dạy cách kiểm chứng.

**Đề xuất (tận dụng brick):**
- **B4:** so số dòng trước và sau mỗi Append/Merge; so tổng doanh thu với file nguồn.
- **B5:** **bảng KPI viết bằng SUMIFS ở B2 trở thành "đáp án" để kiểm tra measure DAX.** Nếu `Total Revenue` theo vùng khớp với SUMIFS thì measure đúng. Việc này vừa là QA thực tế, vừa nối B2 với B5 tốt hơn bất kỳ slide recap nào.

### 2.4 B6 bắt đầu sai điểm

Thực tiễn: **Audience → Câu hỏi/quyết định → KPI → Layout → Chart**.

Kế hoạch B6 hiện đi theo thứ tự: Nguyên tắc → Thư viện 16 loại chart → Kỹ thuật → Build.

**Đề xuất cấu trúc lại B6:**
1. **Dashboard brief:** audience (executive / operational / self-service), 3–5 câu hỏi dashboard phải trả lời, tần suất refresh.
2. **Layout hierarchy:**
   - KPI quan trọng nhất ở góc trên trái.
   - Thứ tự Summary cards → Trend → Detail, từ trên xuống.
   - **Tối đa 5–8 visual trên một trang.**
3. **Chart selection matrix** chỉ gồm các loại chart dùng nhiều trong kinh doanh: xu hướng dùng Line; cơ cấu dùng Bar xếp chồng hoặc Donut/Treemap; so sánh dùng Bar/Column (hoặc Bullet); phân phối dùng Histogram/Box; quan hệ dùng Scatter; địa lý dùng Map; chênh lệch dùng Waterfall.
   - Radar, Surface, Stock, Bubble và Funnel chuyển sang handout.
4. **Build:** KPI card (CUBEVALUE hoặc Pivot 1 ô), Pivot và PivotChart từ Data Model, Slicer kết nối nhiều Pivot (Report Connections), Timeline.
5. **Storytelling:** chú thích trực tiếp trên chart; mỗi insight theo khung **Situation → Complication → Resolution**.
6. **Phát hành:** Refresh All hoặc Refresh on open, ẩn các sheet kỹ thuật, Protect Sheet, xuất PDF, chia sẻ qua OneDrive/SharePoint, và nhắc về dữ liệu khách hàng nhạy cảm.

### 2.5 **Semantic layer**: một nguồn sự thật duy nhất

Lỗi phổ biến nhất trong BI là "mỗi báo cáo ra một con số khác nhau", thường do tính toán nằm rải rác ở tầng hiển thị.

**Quy tắc cần dạy tường minh từ B5:** mọi KPI được tính **chỉ trong measure**. B6 không được thêm công thức tính KPI bên cạnh Pivot.

**Kỹ thuật Excel tương ứng:** hàm **CUBEVALUE / CUBEMEMBER** (Pro Tips – Analytics). Các hàm này cho phép đặt KPI card tự do trên dashboard mà vẫn lấy số từ Data Model. Người làm BI bằng Excel dùng kỹ thuật này thường xuyên, nên nên đưa CUBE functions từ "PRO TIP B5" sang **phần chính của B6**.

### 2.6 Thống kê hồi quy ở B6: không cân xứng với thời lượng

Correlation và regression (Statistics PDF p126–154) có giá trị, nhưng ít xuất hiện trên dashboard vận hành, và B6 đã quá tải.

**Đề xuất:**
- Giữ lại **Histogram, Box plot** (phân phối, outlier) và **scatter + trendline 5 phút** với thông điệp "tương quan ≠ nhân quả".
- Chuyển regression sang Bonus.

### 2.7 Định vị năng lực đầu ra (Self-service BI maturity)

Dùng thang 4 mức để viết mục tiêu khóa học rõ ràng cho lớp hỗn hợp:

| Mức | Năng lực | Đạt ở buổi |
|---|---|---|
| 1 – Consumer | Xem và lọc dashboard | B1 |
| 2 – Explorer | Phân tích ad-hoc, chart cơ bản | B2–B3 |
| 3 – Builder | Kết hợp nguồn dữ liệu, xây báo cáo | B4, B6 |
| 4 – Modeler | Xây semantic model, định nghĩa KPI | B5 |

**Mục tiêu đầu ra đề xuất:**
- 100% học viên đạt tối thiểu **Level 3 (Builder)**.
- Học viên khá đạt **Level 4 (Modeler)**, đo bằng capstone B7.

---

## 3. Phân bổ thời gian đề xuất (150 phút/buổi)

Mục bị cắt khỏi lớp chuyển thành **handout tự học** (vẫn có ảnh minh họa theo style S1).

### B3 — Data & PivotTable

| Phút | Nội dung |
|---|---|
| 10 | Recap B2 (bảng KPI SUMIFS) + câu hỏi BI của buổi |
| 10 | Vì sao cần Pivot; tidy data (nhắc lại B1) |
| 20 | Tạo Pivot, Field List; **Pivot tính như thế nào (tọa độ = bộ lọc)** |
| 10 | Number format, layout (Compact/Tabular), style |
| 10 | *Nghỉ* |
| 20 | Sort, Filter, Group (ngày/giá trị); Slicer & Timeline |
| 20 | Show Values As (% of total, % difference, running total); Calculated Field |
| 10 | PivotChart + preview "3 câu hỏi chọn chart" |
| 30 | **Thực hành trên case** |
| 10 | Quiz + preview B4 (dữ liệu tháng mới bị bẩn) |

**Chuyển sang handout:** Calculated Items, Solve Order, List Formulas, Report Filter Pages, Index calculation.

### B4 — Power Query

| Phút | Nội dung |
|---|---|
| 10 | Recap + demo "nỗi đau": 3 file CSV bẩn |
| 10 | Quy trình Power Excel (PQ → Model → Pivot); kiểm tra phiên bản |
| 15 | Query Editor, Applied Steps (mỗi bước là một "công thức" được ghi lại) |
| 20 | Transform: data type, header, text/date tools, conditional column |
| 10 | *Nghỉ* |
| 15 | **Unpivot bảng Target** (file kế hoạch dạng cột tháng); Group By |
| 20 | Append từ Folder; Merge (các kiểu join); **merge chỉ để minh họa**, còn các bảng lookup load riêng |
| 10 | Calendar table |
| 10 | Load options (Connection only / Data Model) + **đối soát số dòng và tổng** |
| 20 | **Thực hành:** thả thêm tháng thứ 4 vào folder, bấm Refresh, thấy Pivot B3 tự cập nhật |
| 10 | Quiz + preview B5 |

**Chuyển sang handout:** rolling calendar, index column, number-specific tools.

### B5 — Data Model & DAX *(buổi quá tải nhất, cần cắt mạnh)*

| Phút | Nội dung |
|---|---|
| 5 | Recap |
| 20 | Data Model, star schema: Fact vs Lookup, Primary/Foreign key, **relationship thay cho merge** |
| 15 | Relationship: cardinality 1-*, filter direction; **fact table thứ hai (Target) với grain khác** |
| 10 | *Nghỉ* |
| 15 | Calculated Column vs Measure; implicit vs explicit |
| 15 | **Filter context** (dùng lại hình SUMIFS → Pivot từ B2/B3) |
| 20 | CALCULATE, ALL → `% of Total`; `Revenue vs Target %` |
| 15 | Time intelligence với Calendar: YTD, YoY |
| 10 | **KPI dictionary:** measure là định nghĩa KPI dùng chung |
| 20 | **Thực hành:** viết 6 KPI measure và **đối soát với bảng SUMIFS của B2** |
| 5 | Quiz + preview B6 |

**Chuyển sang PRO TIP/handout:** SWITCH(TRUE), DAX text functions, X-functions (SUMX…), disconnected slicers, hierarchies, hide fields, lookup-to-lookup.

### B6 — Visualization & Dashboard

| Phút | Nội dung |
|---|---|
| 5 | Recap: KPI dictionary B5 |
| 10 | **Dashboard brief:** audience, câu hỏi, tần suất refresh |
| 10 | Nguyên tắc: clarity; Good/Bad/Ugly (trục Y gây hiểu lầm) |
| 25 | 3 câu hỏi then chốt + chart matrix (Line, Bar, Stacked/Donut, Histogram/Box, Scatter, Map, Waterfall) |
| 10 | *Nghỉ* |
| 20 | Layout hierarchy, 5–8 visual; KPI card bằng **CUBEVALUE**; RAG bằng conditional formatting; Slicer kết nối nhiều Pivot |
| 45 | **Build dashboard theo hướng dẫn** trên case |
| 15 | Storytelling SCR + phát hành (protect, refresh on open, PDF, chia sẻ) |
| 10 | Quiz + giao brief capstone |

**Chuyển sang handout / Bonus:** Radar, Surface, Stock, Bubble, Funnel, Sunburst; regression.

### B7 — Lưu ý thời lượng

Làm một dự án end-to-end trên dataset **mới** trong 150 phút có thuyết trình là **không khả thi** nếu làm từ đầu tại lớp: chỉ còn khoảng 80 phút làm bài.

**Đề xuất:**
- **Giao brief và dataset cuối B6** để học viên làm ở nhà.
- **B7 dành cho:** ôn tập (20'), hoàn thiện và hỏi đáp (40'), thuyết trình 5–7 phút mỗi nhóm kèm nhận xét (70'), tổng kết (20').
- Rubric chấm sẽ làm ở bước 3 (`criterion-referenced-rubric-generator`).

---

## 4. Thay đổi đề xuất cho `Course_Plan.md` (tóm tắt)

1. **Yêu cầu nền tảng Windows** + kiểm tra máy ở B1. Sửa slide "chọn phiên bản" của S1.
2. Thêm **bảng `Target`** vào dataset case. Unpivot ở B4, fact thứ hai ở B5, RAG/variance ở B6.
3. Thêm **KPI Definition Card** (B2) và **KPI dictionary** (B5).
4. Thêm **đối soát:** B4 so số dòng và tổng; B5 so measure với SUMIFS của B2.
5. **Cấu trúc lại B6** theo thứ tự Brief → Layout → Chart → Build → Story → Publish. Đưa CUBE functions vào B6. Cắt 6 loại chart và regression.
6. **Cắt B5** xuống lõi (bảng thời gian ở mục 3).
7. **B7:** giao capstone từ cuối B6 để làm ở nhà.
8. Ghi mục tiêu đầu ra theo **thang Self-service BI maturity** (Level 3 cho tất cả, Level 4 cho học viên khá).

---

## Phụ lục A — KPI dictionary mẫu cho case (đã validate)

Kết quả `metric_validator.py --strict`: **PASS** (6 metric, 0 lỗi, 0 cảnh báo). Các giá trị target chỉ là **giá trị minh họa**; cần chỉnh theo dataset thật.

| KPI | Công thức (khái niệm) | Aggregation | Owner | Dimensions | Target | Warning / Critical |
|---|---|---|---|---|---|---|
| Total Revenue | SUM(Quantity × Unit_Price) | sum | Sales Director | region, store, category, month | 5 tỷ/tháng | 90% / 80% |
| Total Orders | DISTINCTCOUNT(Order_ID) | count_distinct | Sales Director | region, store, month | 20,000 | 90% / 80% |
| Avg Order Value | Revenue ÷ Orders | ratio | Sales Director | region, store, month | 250,000 | 90% / 80% |
| Gross Margin % | (Revenue − COGS) ÷ Revenue | ratio | Finance | category, region, month | 30% | 90% / 80% |
| Revenue YoY Growth | (Revenue − Revenue LY) ÷ Revenue LY | rate | Finance | region, category, month | +10% | 50% / 20% |
| Active Customers | DISTINCTCOUNT(Customer_ID) | count_distinct | Marketing | region, segment, month | 8,000 | 90% / 80% |

**Gợi ý giảng dạy:**
- Avg Order Value và Gross Margin % là **ratio**, không được cộng dồn hay lấy trung bình của trung bình. Đây là lỗi kinh điển, minh họa rất tốt vì sao cần measure thay vì calculated column (PQ PDF p68, "Calculated columns: good & bad").
- `dashboard_spec_generator.py` gợi ý **line chart cho cả 6 KPI** vì mọi KPI đều có trục tháng. Có thể dùng làm bài tập phản biện ở B6: một dashboard tốt cần kết hợp KPI card, trend và breakdown, không phải 6 đường line.
