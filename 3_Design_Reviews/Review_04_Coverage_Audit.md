# Review #4 — Coverage Audit: Course_Plan v2 so với chuẩn kỹ năng Microsoft

> - **Công cụ:** skill `coverage-audit` (Webb 1997; Porter et al. 2002).
> - **Ngày:** 2026-09-23.
> - **Framework được kiểm tra:** `Course_Plan.md` v2.
> - **Bảng dữ liệu kèm theo:** `Coverage_Audit.csv` (95 dòng, lọc được bằng Excel).

## Nguồn requirement

1. **MO-211 Microsoft Excel Expert (Microsoft 365 Apps)**, "Skills Measured" (48 mục). Nguồn: [trang kỳ thi](https://learn.microsoft.com/en-us/credentials/certifications/exams/mo-211/), [PDF skills outline](https://arch-center.azureedge.net/Learning/Credentials/MO-211_OD_MOS365_ExcelExpert.pdf).
2. **PL-300 Microsoft Power BI Data Analyst**, skills measured as of April 20, 2026. Nguồn: [study guide](https://learn.microsoft.com/en-us/credentials/certifications/resources/study-guides/pl-300).
   - Chỉ giữ **47 mục có thể làm được bằng Excel** (Power Query, Power Pivot, PivotChart, dashboard).
   - **Loại khỏi audit** 31 mục chỉ tồn tại trong Power BI Service/Desktop:
     - DirectLake/DirectQuery, shared semantic model
     - quick measures, calculated tables, calculation groups, Performance Analyzer/DAX query view, visual calculations
     - 4 mục Copilot, paginated report, bookmarks, custom tooltips, navigation, Selection pane, mobile, personalization, automatic page refresh, Analyze feature, AI visuals
     - workspace/app/publish/subscriptions/alerts/promote/gateway
     - workspace roles, item/semantic model access, RLS (2 mục)

   Việc lọc này là **phán đoán của người audit**; danh sách loại trừ ở trên để giảng viên kiểm tra lại.

**Vì sao chọn hai chuẩn này:**
- MO-211 là chuẩn "Excel Expert" chính thức.
- PL-300 là chuẩn nghề "Data Analyst" chính thức của Microsoft; phần Power Query/Data Model/DAX của nó dùng chung công nghệ với Excel.

Không có khung năng lực BI Analyst bằng tiếng Việt được chuẩn hóa để thay thế. Nếu MCI có JD hoặc khung năng lực của doanh nghiệp đối tác, nên chạy lại audit với khung đó.

## Coverage statistics

| | Tổng | Direct | Partial | Indirect | None |
|---|---|---|---|---|---|
| MO-211 (Excel Expert) | 48 | 17 | 12 | 7 | 12 |
| PL-300 (phần làm được bằng Excel) | 47 | 25 | 12 | 5 | 5 |
| **Tổng** | **95** | **42 (44%)** | **24 (25%)** | **12 (13%)** | **17 (18%)** |

```json
{"total_requirements": 95, "directly_covered": 42, "partially_covered": 24, "indirectly_covered": 12, "not_covered": 17}
```

**Đọc số liệu:**
- Khóa học bám sát **PL-300** hơn MO-211: 25/47 direct so với 17/48. Điều này đúng với định vị "Excel **cho BI**".
- Phần lớn "partial" là mục **có trong plan nhưng chỉ ở tầng Nâng cao hoặc Handout**, không dạy ở lõi.
- Theo nguyên tắc audit bảo thủ, các mục này **không** được tính là direct.

## Gap summary

Có **17 mục none** và **12 mục indirect**. Chia theo 4 nhóm:

1. **Nằm ngoài phạm vi BI có chủ đích (10 mục none của MO-211):**
   - Macro: MO-01, 03, 34, 35, 36.
   - Hàm tài chính: MO-26 NPER, MO-27 PMT.
   - Quản lý phiên bản: MO-04.
   - Consolidate: MO-24, đã được Append trong PQ thay thế.
   - Advanced Fill Series: MO-10.

   Không phải lỗ hổng của khóa BI. Tuy nhiên plan nên **ghi rõ "ngoài phạm vi"** để học viên muốn thi MO-211 biết cần tự học thêm.

2. **Lỗ hổng thật, rẻ để vá, sát với BI (nên thêm vào lõi):**
   - **MO-37 dual-axis/combo chart** (và combo trong MO-38): biểu đồ "Doanh thu (cột) + Margin % (đường)" hoặc "Actual vs Target" là loại chart phổ biến nhất trong báo cáo quản trị. Nên thêm vào B6.
   - **PL-04 Data profiling** trong PQ (Column quality/distribution/profile): 5 phút ở B4, gắn trực tiếp với năng lực D2 (hoài nghi dữ liệu).
   - **PL-06 lỗi import** (Keep/Remove Errors) và **PL-02 privacy levels**: đây là hai lỗi học viên *chắc chắn* gặp khi Merge nhiều nguồn ở B4 và capstone. Nên có slide xử lý sự cố.
   - **MO-31 Watch Window**: 2 phút trong mục Auditing B1.
   - **MO-16 Remove Duplicates**: nên đưa vào lõi B4 (trong PQ), vì trùng khóa làm hỏng relationship ở B5.
   - **PL-40 Accessibility**: RAG đỏ-xanh khó đọc với người mù màu; thêm biểu tượng hoặc nhãn và alt text trong B6 (1 slide).

3. **Lỗ hổng nâng cao, phù hợp PRO TIP / Handout / Bonus:**
   - PL-13 reference vs duplicate queries
   - PL-11 semi-structured data (JSON)
   - PL-26 semi-additive measures
   - PL-28 giảm granularity để tăng hiệu năng
   - PL-38 drill-down / Show Details (nên dạy tường minh, vì rubric C6 yêu cầu học viên kiểm chứng nguyên nhân bằng drill-down)
   - PL-31 theme cho dashboard
   - PL-47 sensitivity label
   - MO-02 external references
   - MO-09 Flash Fill
   - MO-11 RANDARRAY
   - MO-14/15 Group-Outline, Subtotal
   - MO-19 quản lý CF rules
   - MO-25 Goal Seek/Scenario (đang ở Bonus)

4. **Lệch tầng (partial vì chỉ có ở Nâng cao/Handout):**
   - FILTER/SORTBY (MO-28, 29)
   - TODAY/WEEKDAY/WORKDAY (MO-22, 23)
   - custom number formats (MO-12)
   - calculation options (MO-08)

   Không cần đưa lên lõi, vì thời lượng B2 đã kín 150 phút. Nhưng **handout phải thực sự có** các mục này, kèm ảnh minh họa, thì coverage mới được tính.

## Khuyến nghị cập nhật Course_Plan (chưa áp dụng)

| # | Thay đổi | Buổi | Chi phí thời gian |
|---|---|---|---|
| 1 | Thêm **Combo / dual-axis chart** (Actual vs Target) vào chart matrix lõi | B6 | Gộp vào 25' chart matrix |
| 2 | Thêm **Column profiling + Remove Duplicates + Keep/Remove Errors** | B4 | Gộp vào 20' Transform |
| 3 | Slide xử lý sự cố **privacy levels / Formula.Firewall** | B4 Handout | 0' tại lớp |
| 4 | **Watch Window** | B1 Auditing | +0–2' |
| 5 | **Drill-down / Show Details** trong Pivot | B3 hoặc B6 | +3' |
| 6 | **Accessibility:** RAG kèm biểu tượng/nhãn, alt text | B6 | 1 slide |
| 7 | Mục **"Ngoài phạm vi & lộ trình thi chứng chỉ"**: macro, hàm tài chính, Consolidate → tự học cho MO-211; Power BI Service → PL-300 | Course_Plan | — |

## Giới hạn của audit

- **Coverage ở đây là "có mặt trên giấy"** trong `Course_Plan.md`. Audit không đánh giá được độ sâu nhận thức (Webb DoK), cách dạy thực tế, hay việc handout đã được soạn hay chưa.
- **Chuẩn MO-211 và PL-300 là chuẩn kỳ thi**, không phải khung năng lực nghề nghiệp. Chúng đo thao tác công cụ nhiều hơn tư duy BI (D1–D4), nên các điểm mạnh của khóa về tư duy BI **không** hiện lên trong audit này.
- Việc chọn 47/78 mục PL-300 "làm được bằng Excel" là phán đoán; cần người có chuyên môn Power BI rà lại.
- **Deck `Slide_S1_GioiThieuExcel.pptx` đã bị xóa.** Các ô "Covered by" của B1 dựa trên nội dung *dự kiến* trong plan, không dựa trên deck hiện có.

---

## Coverage table

| Req ID | Requirement | Covered by | Buổi | Evidence | Notes |
|---|---|---|---|---|---|
| MO-01 | Copy macros between workbooks | — | — | none | Macro/VBA ngoài phạm vi khóa BI. |
| MO-02 | Reference data in other workbooks | B4 Power Query kết nối file | B4 | indirect | Dùng PQ thay cho tham chiếu ngoài ('[Book]Sheet!A1'); tham chiếu ngoài không được dạy. |
| MO-03 | Enable macros in a workbook | — | — | none | Macro ngoài phạm vi. |
| MO-04 | Manage workbook versions | — | — | none | Version history (OneDrive) không có trong plan. |
| MO-05 | Restrict editing | B6 Phát hành: Protect Sheet | B6 | partial | Chỉ có Protect Sheet; không có Restrict Editing/Mark as Final. |
| MO-06 | Protect worksheets and cell ranges | B6 Phát hành: Protect Sheet | B6 | direct | Cell ranges (Allow Edit Ranges) không nêu riêng. |
| MO-07 | Protect workbook structure | B6 Phát hành: ẩn sheet kỹ thuật | B6 | partial | Ẩn sheet có, Protect Workbook Structure chưa nêu. |
| MO-08 | Configure formula calculation options | B2 PRO TIP Calculation modes | B2 | partial | Chỉ ở tầng Nâng cao. |
| MO-09 | Fill cells by using Flash Fill | B1 Pro Tips (Productivity) — chung chung | B1 | indirect | Plan không nêu Flash Fill cụ thể. |
| MO-10 | Fill cells by using advanced Fill Series options | — | — | none | Formulas PDF p59 có Fill Series nhưng plan không dùng. |
| MO-11 | Generate numeric data by using RANDARRAY() | B2 Handout RAND | B2 | indirect | RANDARRAY (Formulas PDF p92) không có trong plan. |
| MO-12 | Create custom number formats | B6 PRO TIP custom number formats | B6 | partial | Chỉ ở tầng Nâng cao. |
| MO-13 | Configure data validation | B1 Data Validation | B1 | direct |  |
| MO-14 | Group and ungroup data (outline) | B3 Group trong Pivot | B3 | indirect | Pivot grouping khác Data → Group (outline). |
| MO-15 | Calculate data by inserting subtotals and totals | B3 Pivot totals | B3 | indirect | Tính năng Subtotal / Table Total Row chưa có. |
| MO-16 | Remove duplicate records | B2 PRO TIP Uniques & Duplicates | B2 | partial | Nên có ở lõi B4 (Remove Duplicates trong PQ). |
| MO-17 | Create custom conditional formatting rules | B6 RAG bằng conditional formatting; B3 PRO TIP | B3, B6 | direct |  |
| MO-18 | Create conditional formatting rules that use formulas | B2 Handout formula-based formatting | B2 | partial | Chỉ ở Handout. |
| MO-19 | Manage conditional formatting rules | B6 RAG | B6 | indirect | Rules Manager không nêu. |
| MO-20 | Logical operations with nested IF, IFS, SWITCH, *IF(S), MAXIFS, MINIFS, AND, OR, NOT, LET | B2 Logic + COUNTIFS/SUMIFS/AVERAGEIFS; LET ở Nâng cao | B2 | direct | SWITCH, MAXIFS/MINIFS chưa nêu; LET chỉ ở Nâng cao. |
| MO-21 | Look up data with XLOOKUP, VLOOKUP, HLOOKUP, MATCH, INDEX | B2 VLOOKUP → INDEX/MATCH → XLOOKUP | B2 | direct | HLOOKUP ở Handout. |
| MO-22 | Reference date and time with NOW() and TODAY() | B2 Handout | B2 | partial | Date lõi B2 không liệt kê TODAY/NOW. |
| MO-23 | Calculate dates with WEEKDAY() and WORKDAY() | B2 Date lõi: NETWORKDAYS | B2 | partial | WEEKDAY/WORKDAY không ở lõi. |
| MO-24 | Summarize data from multiple ranges with Consolidate | — | — | none | Thay bằng Append trong PQ (B4) — có thể ghi rõ là thay thế. |
| MO-25 | What-if analysis with Goal Seek and Scenario Manager | Bonus (Pro Tips – Analytics) | Bonus | indirect | Ngoài 7 buổi. |
| MO-26 | Forecast data with AND(), IF(), NPER() | — | — | none | Hàm tài chính ngoài phạm vi. |
| MO-27 | Calculate financial data with PMT() | — | — | none | Hàm tài chính ngoài phạm vi. |
| MO-28 | Filter data by using FILTER() | B2 Nâng cao Dynamic Arrays | B2 | partial | Chỉ ở tầng Nâng cao. |
| MO-29 | Sort data by using SORTBY() | B2 Nâng cao Dynamic Arrays | B2 | partial | Chỉ ở tầng Nâng cao. |
| MO-30 | Trace precedence and dependence | B1 Auditing | B1 | direct |  |
| MO-31 | Monitor cells and formulas with the Watch Window | — | — | none | Chưa có; dễ thêm vào mục Auditing B1. |
| MO-32 | Validate formulas by using error checking rules | B1 Error Checking | B1 | direct |  |
| MO-33 | Evaluate formulas | B1 Evaluate Formula | B1 | direct |  |
| MO-34 | Record simple macros | — | — | none | Macro ngoài phạm vi. |
| MO-35 | Name simple macros | — | — | none | Macro ngoài phạm vi. |
| MO-36 | Edit simple macros | — | — | none | Macro ngoài phạm vi. |
| MO-37 | Create and modify dual-axis charts | — | — | none | Chưa có; combo Revenue (cột) + Margin % (đường) rất sát BI. |
| MO-38 | Box & Whisker, Combo, Funnel, Histogram, Sunburst, Waterfall charts | B6 lõi: Box, Histogram, Waterfall; Handout: Funnel, Sunburst | B6 | partial | Combo chưa có. |
| MO-39 | Create PivotTables | B3 | B3 | direct |  |
| MO-40 | Modify field selections and options | B3 Field List, Analyze/Design | B3 | direct |  |
| MO-41 | Create slicers | B3 Slicer & Timeline | B3 | direct |  |
| MO-42 | Group PivotTable data | B3 Group | B3 | direct |  |
| MO-43 | Add calculated fields | B3 Calculated Field | B3 | direct |  |
| MO-44 | Configure value field settings | B3 Summarize Values By, Show Values As | B3 | direct |  |
| MO-45 | Create PivotCharts | B3 PivotChart | B3 | direct |  |
| MO-46 | Manipulate options in existing PivotCharts | B3 PivotChart (Pivot PDF p34–36) | B3 | direct |  |
| MO-47 | Apply styles to PivotCharts | B3 PivotChart layouts & styles | B3 | direct |  |
| MO-48 | Drill down into PivotChart details | B3 Handout field buttons | B3 | partial | Chỉ ở Handout. |
| PL-01 | Identify and connect to data sources | B4 Connections, Folder, CSV | B4 | direct | Shared semantic model: loại trừ (Power BI). |
| PL-02 | Change data source settings, credentials and privacy levels | — | — | none | Privacy levels gây lỗi Formula.Firewall khi Merge nhiều nguồn — nên có trong B4 handout. |
| PL-03 | Create and modify parameters | B4 PRO TIP tham số hóa folder | B4 | partial | Chỉ ở tầng Nâng cao. |
| PL-04 | Evaluate data, including data statistics and column properties | B2 thống kê mô tả | B2 | partial | Column quality/distribution/profile trong PQ chưa có. |
| PL-05 | Resolve inconsistencies, unexpected or null values, data quality issues | B4 Transform + đối soát | B4 | direct |  |
| PL-06 | Resolve data import errors | B1 Common Errors (lỗi công thức) | B1 | indirect | Lỗi import trong PQ (Keep/Remove Errors) chưa có. |
| PL-07 | Select appropriate column data types | B4 data type | B4 | direct |  |
| PL-08 | Create and transform columns | B4 text/date tools, conditional column | B4 | direct |  |
| PL-09 | Group and aggregate rows | B4 Group By | B4 | direct |  |
| PL-10 | Pivot, unpivot, and transpose data | B4 Unpivot Target | B4 | partial | Pivot column và Transpose chưa nêu. |
| PL-11 | Convert semi-structured data to a table | — | — | none | JSON/nested chưa có. |
| PL-12 | Create fact tables and dimension tables | B4 load lookup riêng; B5 star schema | B4, B5 | direct |  |
| PL-13 | Reference vs duplicate queries and the impact | — | — | none | Chưa có. |
| PL-14 | Merge and append queries | B4 Append Folder, Merge | B4 | direct |  |
| PL-15 | Identify and create appropriate keys for relationships | B2 Laws of Lookups; B5 Primary/Foreign key | B2, B5 | direct |  |
| PL-16 | Configure data loading for queries | B4 Load options | B4 | direct |  |
| PL-17 | Configure table and column properties | B5 Handout hide fields | B5 | partial | Chỉ ở Handout. |
| PL-18 | Implement role-playing dimensions | B5 PRO TIP active/inactive relationship | B5 | partial | Chỉ ở tầng Nâng cao; USERELATIONSHIP chưa nêu. |
| PL-19 | Define cardinality and cross-filter direction | B5 Relationship | B5 | direct |  |
| PL-20 | Create a common date table | B4 Calendar; B5 | B4, B5 | direct |  |
| PL-21 | Use cases for calculated columns and calculated tables | B5 Calculated Column vs Measure | B5 | direct | Calculated tables không có trong Power Pivot Excel. |
| PL-22 | Create single aggregation measures | B5 explicit measure | B5 | direct |  |
| PL-23 | Use the CALCULATE function | B5 CALCULATE | B5 | direct |  |
| PL-24 | Implement time intelligence measures | B5 YTD, YoY | B5 | direct |  |
| PL-25 | Use basic statistical functions (DAX) | B2 thống kê (Excel) | B2 | partial | DAX stats (PQ PDF p82–85) không ở lõi B5. |
| PL-26 | Create semi-additive measures | — | — | none | Chưa có (số dư tồn kho, số dư tài khoản). |
| PL-27 | Remove unnecessary rows and columns for performance | B5 Handout performance | B5 | partial | Chỉ ở Handout. |
| PL-28 | Improve performance by reducing granularity | B4 Group By | B4 | indirect | Không được trình bày như kỹ thuật hiệu năng. |
| PL-29 | Select an appropriate visual | B3 preview; B6 3 Key Questions + chart matrix | B3, B6 | direct |  |
| PL-30 | Format and configure visuals | B3 PivotChart; B6 | B3, B6 | direct |  |
| PL-31 | Apply and customize a theme | B1 Handout tab Page Layout | B1 | indirect | Theme cho dashboard chưa nêu. |
| PL-32 | Apply conditional formatting | B6 RAG; B3 PRO TIP | B3, B6 | direct |  |
| PL-33 | Apply slicing and filtering | B3 Slicer/Timeline; B6 | B3, B6 | direct |  |
| PL-34 | Configure the report page | B6 layout hierarchy | B6 | partial | Page setup/print area cho dashboard chưa nêu. |
| PL-35 | Apply sorting to visuals | B3 Sort | B3 | direct |  |
| PL-36 | Configure sync slicers | B6 Slicer kết nối nhiều Pivot | B6 | direct | Tương đương Report Connections trong Excel. |
| PL-37 | Edit and configure interactions between visuals | B6 Slicer connections | B6 | partial | Chỉ qua slicer; chưa bàn chủ động chọn visual nào bị lọc. |
| PL-38 | Configure drillthrough navigation | Rubric C6 drill-down | B7 | indirect | Show Details / drill-down trong Pivot chưa dạy tường minh. |
| PL-39 | Configure export settings | B6 xuất PDF | B6 | direct |  |
| PL-40 | Design reports for accessibility | — | — | none | Alt text, bảng màu an toàn cho mù màu (RAG đỏ-xanh) chưa có. |
| PL-41 | Use grouping, binning, and clustering | B3 Group theo giá trị | B3 | partial | Clustering chưa có. |
| PL-42 | Use reference lines, error bars, and forecasting | B6 scatter + trendline | B6 | partial | Target line có qua KPI card; Forecast Sheet chưa có. |
| PL-43 | Detect outliers and anomalies | B2 outlier; B6 Box plot | B2, B6 | direct |  |
| PL-44 | Create dashboards | B6 | B6 | direct |  |
| PL-45 | Choose a distribution method | B6 OneDrive/SharePoint, PDF | B6 | direct |  |
| PL-46 | Configure a semantic model scheduled refresh | B6 Refresh on open | B6 | partial | Excel không có scheduled refresh trên desktop; nêu giới hạn và hướng Power BI. |
| PL-47 | Apply sensitivity labels | B6 lưu ý dữ liệu khách hàng | B6 | indirect | Sensitivity label (Microsoft Purview) chưa nêu. |
