# Rubric: Capstone B7 — Dự án BI end-to-end trên Excel

> Tạo bằng skill `criterion-referenced-rubric-generator` (Brookhart, Andrade, Sadler, Jonsson & Svingby). Đầu vào: `Course_Plan.md` v2. Ngày 2026-09-23.

- **Đánh giá mục tiêu:** Học viên tự thực hiện một dự án BI end-to-end trên bộ dữ liệu mới theo chuỗi Brief → Power Query → Data Model → KPI measure → Dashboard → Insight và đề xuất. Qua đó học viên thể hiện bốn năng lực D1–D4:
  - D1: bắt đầu từ câu hỏi
  - D2: đối soát dữ liệu
  - D3: tự động hóa
  - D4: truyền đạt insight
- **Đối tượng:** người đi làm hoặc sinh viên, trình độ hỗn hợp. Dự án được giao cuối B6, làm ở nhà, thuyết trình 5–7 phút ở B7.
- **Criteria:** 6 · **Levels:** 4

---

## Rubric

| Tiêu chí | Level 1 | Level 2 | Level 3 | Level 4 |
|---|---|---|---|---|
| **C1. Câu hỏi kinh doanh & Brief** (D1) | Không có brief. Bài mở đầu bằng mô tả công cụ đã dùng ("em làm Pivot doanh thu"). Không xác định ai là người xem dashboard. | Có brief nêu người xem và câu hỏi. Tuy nhiên câu hỏi ở dạng chủ đề ("xem doanh thu", "phân tích khách hàng") và không có chỉ số, chiều phân tích, khoảng thời gian hay mốc so sánh. | Brief nêu người xem, tần suất refresh và 3–5 câu hỏi đo lường được. Mỗi câu hỏi có chỉ số, chiều, thời gian và mốc so sánh, ví dụ "Vùng nào có doanh thu tháng 6 dưới 90% target?". Mỗi visual trên dashboard trả lời một câu hỏi trong brief. | Có mọi yếu tố của Level 3. Ngoài ra, câu hỏi được sắp theo quyết định kinh doanh mà chúng phục vụ, và có câu hỏi "vì sao" nảy sinh từ kết quả phân tích đầu tiên. Brief ghi rõ câu hỏi nào dữ liệu hiện có không trả lời được. |
| **C2. Chuẩn bị dữ liệu: Power Query, đối soát, refresh** (D2, D3) | Dữ liệu được copy-paste và sửa tay trên sheet, hoặc query còn sai kiểu dữ liệu, còn dòng tiêu đề lặp hay dòng trống. Khi có dữ liệu mới phải làm lại thủ công. | Mỗi nguồn có một query với kiểu dữ liệu và header đúng. Tuy nhiên vẫn còn bước sửa tay ngoài Power Query, hoặc Refresh báo lỗi hay cho số sai khi thêm file mới. Bài không có con số đối soát. | Toàn bộ bước làm sạch nằm trong Applied Steps và query được đặt tên theo nội dung. Thêm file mới rồi Refresh chạy không lỗi. Số dòng và tổng doanh thu được ghi lại và khớp với dữ liệu nguồn. | Có mọi yếu tố của Level 3. Ngoài ra, pipeline có tham số (đường dẫn folder, ngày cắt) hoặc có query/measure tự kiểm tra (trùng khóa, giá trị trống). Các dòng bị loại được ghi chú kèm lý do. |
| **C3. Mô hình dữ liệu** | Dữ liệu được merge thành một bảng phẳng, hoặc các bảng có trong Data Model nhưng không có relationship. Các Pivot lấy từ những nguồn rời nhau. | Có relationship giữa fact và lookup, nhưng còn lỗi cấu trúc: quan hệ nhiều-nhiều, khóa trùng ở bảng lookup, cột ngày chưa nối với Calendar, hoặc filter direction ngược. Vì vậy khi lọc theo một chiều, con số bị lặp hoặc sai. | Mô hình là star schema: fact nối với từng lookup bằng quan hệ 1-*, khóa của lookup là duy nhất, bảng Calendar liên tục và nối vào cột ngày của fact. Lọc theo bất kỳ lookup nào cũng ra số đúng. | Có mọi yếu tố của Level 3. Ngoài ra, mô hình có fact thứ hai khác grain (ví dụ `Target` theo Tháng × Cửa hàng) và chỉ nối qua các lookup chung. Học viên giải thích được vì sao không lọc được Target theo Product. Các cột kỹ thuật được ẩn khỏi client tools. |
| **C4. KPI & Measure** | KPI được tính bằng implicit measure (kéo thả Sum) hoặc bằng công thức đặt cạnh Pivot trên sheet. Không có định nghĩa KPI. | Có explicit measure cho KPI, nhưng ít nhất một KPI sai logic: lấy trung bình của trung bình, cộng các tỷ lệ, tính KPI tỷ lệ bằng calculated column, hoặc YoY/YTD ra sai khi đổi bộ lọc. Định nghĩa KPI không được ghi, hoặc không khớp với measure. | Mỗi KPI có một KPI Definition Card (công thức, grain, dimensions, target) và một explicit measure khớp với định nghĩa đó. KPI tỷ lệ được tính dạng tổng chia tổng. Ít nhất một KPI đã được đối soát với cách tính thứ hai (SUMIFS hoặc số nguồn) và cho kết quả khớp. | Có mọi yếu tố của Level 3. Ngoài ra, measure được xây theo tầng: measure cơ sở rồi đến measure dẫn xuất tái sử dụng. KPI so sánh (% of total, vs target, YoY) dùng CALCULATE, ALL hoặc time intelligence. Khi được hỏi, học viên giải thích được filter context của một ô bất kỳ. |
| **C5. Thiết kế dashboard** | Các bảng Pivot và biểu đồ được đặt cạnh nhau, không theo thứ tự nào. Loại chart không khớp với câu hỏi (ví dụ pie cho xu hướng theo thời gian), hoặc trục và tỷ lệ gây đọc sai. Dashboard không có bộ lọc tương tác. | Loại chart khớp với loại câu hỏi và có slicer. Tuy nhiên còn một trong các vấn đề sau: bố cục không có thứ bậc (KPI chính không nằm đầu trang), có hơn 8 visual, hoặc KPI hiển thị không có mốc so sánh (target, kỳ trước). | Bố cục đi từ KPI đến xu hướng rồi đến chi tiết, với 5–8 visual. KPI card có mốc so sánh và màu trạng thái RAG. Slicer/Timeline điều khiển mọi visual. Mọi con số lấy từ measure (qua CUBEVALUE hoặc Pivot); không có công thức tính KPI ở tầng hiển thị. | Có mọi yếu tố của Level 3. Ngoài ra, đã bỏ các yếu tố không mang thông tin (hiệu ứng 3D, legend trùng lặp, gridline thừa), và có chú thích trực tiếp trên chart chỉ vào điểm bất thường. Dashboard đã sẵn sàng phát hành: refresh on open, sheet kỹ thuật được ẩn, sheet được bảo vệ. Người xem tự tìm được câu trả lời cho các câu hỏi trong brief mà không cần giải thích thêm. |
| **C6. Insight & Thuyết trình** (D4) | Bài thuyết trình kể lại các bước kỹ thuật đã làm hoặc đọc lại con số trên biểu đồ ("doanh thu tháng 6 là 4,2 tỷ"). | Mỗi biểu đồ có một nhận xét (tăng/giảm, vùng cao nhất). Nhận xét dừng ở mức mô tả, không nêu nguyên nhân khả dĩ hay hàm ý cho quyết định. | Mỗi insight chính đi theo cấu trúc Situation → Complication → Resolution: con số có mốc so sánh, nguyên nhân khả dĩ được kiểm chứng bằng drill-down trên dashboard, và có đề xuất hành động gửi đúng người xem trong brief. | Có mọi yếu tố của Level 3. Ngoài ra, các insight được sắp theo tác động kinh doanh ước lượng (bằng tiền hoặc %). Học viên phân biệt tương quan với nhân quả, nêu giới hạn của dữ liệu, và trả lời câu hỏi phản biện bằng cách thao tác trên dashboard. |

### Trọng số và quy đổi

| Tiêu chí | Trọng số | Buổi hình thành |
|---|---|---|
| C1 Câu hỏi & Brief | 10% | B1 → B6 (câu hỏi mở mỗi buổi), B6 (dashboard brief) |
| C2 Chuẩn bị dữ liệu | 20% | B1 (tidy, Table), B4 |
| C3 Mô hình dữ liệu | 15% | B5 |
| C4 KPI & Measure | 20% | B2 (KPI card, SUMIFS), B3, B5 |
| C5 Dashboard | 20% | B3 (PivotChart), B6 |
| C6 Insight & Thuyết trình | 15% | Insight cuối mỗi buổi, B6 (SCR) |

**Điểm tổng** = Σ (Level × trọng số), thang 1.0–4.0.

**Quy đổi sang mục tiêu đầu ra** (thang Self-service BI maturity trong `Course_Plan.md`):

| Kết quả | Điều kiện |
|---|---|
| **Hoàn thành khóa, đạt Level 3 – Builder** | Không tiêu chí nào ở Level 1, **và** C2 ≥ Level 3, **và** C5 ≥ Level 3 |
| **Đạt Level 4 – Modeler** | Đạt điều kiện Builder, **và** C3 = Level 4, **và** C4 = Level 4 |
| **Chưa hoàn thành** | Có tiêu chí ở Level 1, hoặc C2 < 3, hoặc C5 < 3. Học viên được nộp lại một lần sau khi nhận phản hồi. |

C2 và C5 là điều kiện bắt buộc vì chúng thể hiện năng lực cốt lõi của Level 3: kết hợp nguồn dữ liệu và xây báo cáo.

---

## Design Rationale

**Vì sao chọn 6 tiêu chí:**
- Mỗi tiêu chí ứng với một mắt xích của quy trình BI và một brick trong khóa học.
- Rubric mặc định 4 tiêu chí không đủ, vì bài thu hoạch phải kiểm tra được cả chuỗi. Một học viên có thể dựng mô hình đúng (C3) nhưng viết KPI tỷ lệ sai (C4), hoặc làm dashboard đẹp (C5) nhưng không có insight (C6). Tách riêng từng mắt xích cho thấy chính xác mắt xích nào đang yếu.
- **Bốn năng lực D1–D4 được gắn vào tiêu chí sản phẩm**, không chấm như tiêu chí độc lập. Lý do: năng lực chỉ tồn tại khi được thực hành (R1 §2). Ví dụ D2 thể hiện qua con số đối soát ở C2 và C4, D3 thể hiện qua bài kiểm tra Refresh ở C2. Chấm "thái độ" tách rời sản phẩm sẽ kém tin cậy.
- **C3 và C4 được tách riêng** vì trong thực tế lỗi mô hình (quan hệ nhiều-nhiều) và lỗi measure (trung bình của trung bình) có nguyên nhân khác nhau và cần phản hồi khác nhau.

**Logic phân biệt các level** (áp dụng cho mọi tiêu chí):

| Bước | Ý nghĩa |
|---|---|
| Level 1 → 2 | Thành phần **có mặt** nhưng chưa vận hành đúng (có query nhưng Refresh lỗi; có measure nhưng sai logic). |
| Level 2 → 3 | Thành phần **vận hành đúng và kiểm chứng được** (Refresh chạy, số khớp khi đối soát, mỗi visual trả lời một câu hỏi). Level 3 là chuẩn Builder. |
| Level 3 → 4 | Thành phần được **thiết kế có chủ đích** và học viên **giải thích được lựa chọn**: tham số hóa, measure theo tầng, fact thứ hai khác grain, insight sắp theo tác động. |

---

## Student-Friendly Version

*Dùng để tự đánh giá trước khi nộp.*

**C1. Câu hỏi và brief của bạn**
- L1: Bạn chưa có brief; bạn bắt đầu bằng việc kể mình đã dùng công cụ gì.
- L2: Bạn có ghi câu hỏi, nhưng câu hỏi còn chung chung như "xem doanh thu".
- L3: Mỗi câu hỏi của bạn có chỉ số, chiều, thời gian và mốc so sánh. Mỗi biểu đồ trả lời đúng một câu hỏi đó.
- L4: Câu hỏi của bạn được sắp theo quyết định mà chúng phục vụ. Bạn có câu hỏi "vì sao" tiếp theo, và bạn ghi rõ điều gì dữ liệu chưa trả lời được.

**C2. Dữ liệu của bạn**
- L1: Bạn sửa tay trên sheet; có dữ liệu mới là phải làm lại.
- L2: Bạn có query, nhưng vẫn còn bước sửa tay, hoặc Refresh bị lỗi.
- L3: *Thử ngay:* thả thêm một file vào folder rồi bấm Refresh. Nếu không lỗi, và số dòng cùng tổng doanh thu bạn ghi lại khớp với nguồn, bạn đạt L3.
- L4: Pipeline của bạn có tham số hoặc query tự kiểm tra; các dòng bị loại được ghi chú kèm lý do.

**C3. Mô hình của bạn**
- L1: Bạn gộp hết thành một bảng lớn, hoặc các bảng không nối với nhau.
- L2: Các bảng có nối, nhưng lọc theo một chiều nào đó thì số bị lặp hoặc sai.
- L3: Mô hình của bạn là hình ngôi sao: fact ở giữa, lookup xung quanh, quan hệ 1-*, có Calendar. Lọc theo chiều nào cũng ra số đúng.
- L4: Bạn nối thêm được bảng Target (khác grain) và giải thích được vì sao Target không lọc được theo sản phẩm.

**C4. KPI của bạn**
- L1: Bạn kéo thả Sum hoặc viết công thức bên cạnh Pivot.
- L2: Bạn có measure, nhưng kiểm tra lại: KPI tỷ lệ của bạn có đang lấy trung bình của trung bình không?
- L3: Mỗi KPI có thẻ định nghĩa và measure khớp với thẻ. Bạn đã so ít nhất một KPI với SUMIFS hoặc số nguồn, và hai kết quả khớp nhau.
- L4: Measure của bạn xây theo tầng, bạn dùng CALCULATE/ALL/YoY, và bạn giải thích được "ô này đang bị lọc bởi những gì".

**C5. Dashboard của bạn**
- L1: Bảng và biểu đồ nằm lộn xộn; có chart sai loại, hoặc trục gây hiểu nhầm.
- L2: Chart đúng loại, nhưng KPI chính không nằm ở đầu trang, có quá 8 visual, hoặc con số không có mốc so sánh.
- L3: *Bài kiểm tra 5 giây:* đưa dashboard cho một người khác xem. Họ phải thấy ngay KPI chính, biết con số đạt hay chưa đạt target (màu RAG), và dùng slicer lọc được mọi visual.
- L4: Bạn đã bỏ mọi thứ thừa, có chú thích chỉ vào điểm bất thường, và dashboard đã sẵn sàng gửi đi (tự refresh, sheet kỹ thuật được ẩn và khóa).

**C6. Phần trình bày của bạn**
- L1: Bạn kể lại mình đã làm gì, hoặc đọc lại con số.
- L2: Bạn nói biểu đồ tăng hay giảm, nhưng chưa nói "vậy thì sao".
- L3: Mỗi insight của bạn có đủ: tình hình (con số có so sánh) → vấn đề (nguyên nhân bạn đã kiểm chứng) → đề xuất hành động.
- L4: Bạn xếp insight theo mức tác động, tách rõ tương quan và nhân quả, nêu giới hạn của dữ liệu, và trả lời câu hỏi bằng cách thao tác trực tiếp trên dashboard.

---

## Calibration Notes

**Các điểm quyết định mà người chấm dễ bất đồng:**

1. **C2, Level 2 hay 3: chấm bằng thao tác, không chấm bằng lời kể.**
   - Trong 40 phút hoàn thiện ở B7, giảng viên đưa cho học viên một file dữ liệu mới và yêu cầu Refresh.
   - Refresh chạy không lỗi và số đối soát khớp → Level 3.
   - Refresh cần sửa tay dù chỉ một bước → Level 2.

2. **C4, Level 2 hay 3: kiểm tra KPI tỷ lệ.**
   - Trên dashboard, lọc một danh mục, rồi so Avg Order Value hoặc Gross Margin % với phép tính tay (tổng doanh thu ÷ tổng số đơn của danh mục đó).
   - Lệch → Level 2, kể cả khi mọi KPI dạng tổng đều đúng.

3. **C3 và C4 phải chấm độc lập.**
   - Mô hình sai (Level 2) có thể khiến measure đúng logic cho số sai. Trong trường hợp đó, C4 vẫn chấm theo **logic của measure**, được kiểm tra trên một chiều không bị lỗi mô hình ảnh hưởng.
   - Không chấm trừ hai lần cho cùng một lỗi gốc.

4. **C5, Level 3 hay 4:**
   - Level 3 là "đúng chuẩn": thứ bậc, RAG, 5–8 visual, số từ measure.
   - Level 4 đòi hỏi **dấu hiệu của chủ đích**: chú thích chỉ vào điểm bất thường, và dashboard sẵn sàng phát hành.
   - Dashboard nhiều màu sắc và hiệu ứng **không** được tính là Level 4. Hiệu ứng 3D còn là dấu hiệu Level 2–3.

5. **C6, Level 2 hay 3: hỏi "vậy thì sao?"**
   - Nếu sau câu hỏi này học viên nêu được nguyên nhân đã kiểm chứng trên dữ liệu và một hành động cụ thể → Level 3.
   - Nếu học viên chỉ lặp lại mô tả → Level 2.
   - Nguyên nhân phỏng đoán nhưng không kiểm chứng trên dashboard → vẫn là Level 2.

6. **C1 thường bị chấm cao.**
   - Brief chỉ đạt Level 3 khi câu hỏi **có mốc so sánh**. "Doanh thu theo vùng tháng 6" là mô tả, chưa phải câu hỏi → Level 2.
   - Đối chiếu từng visual với brief: có visual không trả lời câu hỏi nào → không đạt Level 3.

7. **Lớp trình độ hỗn hợp: chấm sản phẩm, không chấm xuất phát điểm hay công sức.** Học viên mới học Excel mà đạt Level 3 ở mọi tiêu chí thì hoàn thành khóa ngang với học viên đã có kinh nghiệm.

8. **Bài làm cá nhân.**
   - Capstone làm cá nhân, không nhận bài nhóm. Cả 6 tiêu chí chấm cho từng học viên.
   - Trong phần Q&A, mỗi học viên trả lời ít nhất một câu hỏi filter context (C4) và một câu hỏi phản biện (C6).

9. **Hiệu chuẩn trước khi chấm.** Nếu có từ hai người chấm trở lên:
   - Cùng chấm độc lập 2 bài mẫu (một bài ở khoảng Level 2, một bài ở khoảng Level 3).
   - So kết quả theo từng tiêu chí, thống nhất các chỗ lệch nhau từ 2 level trở lên.
   - Sau đó mới chấm chính thức.

---

## Cách dùng trong khóa học

- **Phát rubric (bản Student-Friendly) cùng brief capstone ở cuối B6.** Theo Andrade, rubric phát huy tác dụng nhất khi học viên có nó trước khi làm bài.
- Mỗi tiêu chí đã được luyện dần trong các buổi trước (cột "Buổi hình thành"). Giảng viên có thể dùng đúng mô tả level của rubric khi góp ý phần thực hành ở B2–B6.
- **Giới hạn:** rubric mô tả *cái gì* ở mỗi level, không mô tả *cách* để lên level. Cần đi kèm phản hồi cụ thể, và cho phép học viên chưa hoàn thành nộp lại một lần.
