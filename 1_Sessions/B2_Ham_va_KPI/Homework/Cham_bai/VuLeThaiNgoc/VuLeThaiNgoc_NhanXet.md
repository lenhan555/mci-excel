**Điểm: 8/10**

# Nhận xét bài tập về nhà Buổi 2 - Vũ Lê Thái Ngọc

File: `b2_btvn_vu-le-thai-ngoc.xlsx`

| # | Việc | Kết quả | Trọng số | % đạt / Thành phần |
|---|---|---|---:|---|
| 1 | Tra cứu XLOOKUP | Đạt | 2 | 100% |
| 2 | Cột tính toán | Đạt | 2 | 100% |
| 3 | Bảng KPI 1: Miền x Ngành hàng | Đạt | 1 | 100% |
| 4 | Bảng KPI 2: theo Miền | Đạt | 1 | 100% |
| 5 | KPI Definition Card (AOV) | Một phần | 1 | 50% |
| 6 | Insight | 1/3 thành phần | 3 | 1/3 |
| | **Tổng đã làm tròn** | | **10** | **8/10** |

Thang điểm: việc 1-2 mỗi việc 2 điểm; việc 3-5 mỗi việc 1 điểm. Đạt được 100%, Một phần được 50%, chưa có phần việc đáp ứng bài được 0%. Insight có ba thành phần, mỗi thành phần đạt được 1 điểm. Cộng các phần theo trọng số trước, rồi làm tròn tổng về số nguyên theo quy tắc từ nửa điểm trở lên làm tròn lên. Không cộng các điểm từng mục đã làm tròn.

## Tài liệu để xem lại

Slide nằm trong `Slide_B2_HocVien.pptx`. Đường dẫn Excel dưới đây tính từ thư mục Buổi 2.

| Slide | Nội dung | File để làm lại |
|---|---|---|
| 4-5 | Grain và KPI Definition Card | `Handout/B2_KPI_Definition_Card.docx` |
| 7 | IFS và nhãn phân khúc | `Demo/B2_Demo_Logic.xlsx`, sheet Demo_IF |
| 16-17 | COUNTIFS Line_No = 1 và tổng chia tổng | `Demo/B2_Demo_SUMIFS.xlsx`, sheet Demo_SUMIFS |
| 30-31 | Thực hành bảng KPI và card | `Practice/B2_MCI_Mart_Q1_2025_Answer.xlsx`, sheet KPI_Card |
| 40 | Câu insight có số, nguyên nhân, hàm ý | Sheet Insight trong bài của bạn |

## Đối soát nhanh

| Chỉ tiêu | Số cần đạt | File của bạn | Khớp? |
|---|---:|---:|---|
| Số dòng dữ liệu | 1.919 | 1.919 | Có |
| Số ô #N/A (Region, Category, Unit_Cost) | 0 | 0 | Có |
| Tổng doanh thu (VNĐ) | 296.965.000 | 296.965.000 | Có |
| Tổng bảng KPI 1 (Miền × Ngành) | 296.965.000 | 296.965.000 | Có |
| Tổng số đơn | 966 | 966 | Có |
| Tổng COGS (VNĐ) | 196.595.000 | 196.595.000 | Có |
| Gross Margin % tháng 4 | 33,8% | 33,8% | Có |
| AOV tháng 4 (VNĐ) | 307.417 | 307.417 | Có |

Số tiền trong bảng làm tròn đến VNĐ để dễ đọc; phần đối soát dùng giá trị đầy đủ đến sáu chữ số thập phân.

## Làm tốt

- Sales đủ 1.919 dòng; ba cột tra cứu và cột tính toán đều là công thức, khớp theo khóa dòng.
- KPI_SUMIFS!H9 = 296.965.000, B16 = 966; B23 = 307.417,184 VNĐ, C30 = 33,8%, đều dùng tổng chia tổng.

## Điểm cần lưu ý

### 1. Card còn thiếu ngưỡng và sai grain

**Bạn đã làm:** KPI_Card!B13 trống; B9 ghi "1 line_no = 1 đơn hàng". B8 ghi tblSales; B12 chưa có target AOV bằng số hoặc nguồn cụ thể.

**Đúng cần là:** Một dòng dữ liệu là một sản phẩm trong một đơn. Chỉ Line_No = 1 được dùng đếm đơn. Card cần đủ 10 trường, ghi target AOV và hai ngưỡng.

**Hậu quả:** Định nghĩa một dòng thành một đơn khiến người đọc có thể đếm số dòng thay số đơn. Thiếu ngưỡng nên chưa phân loại được KPI đạt hay cảnh báo.

**Cách sửa:** B9 sửa grain và giữ kỳ tháng; B8 ghi tblSales_T4 với Revenue, Order_ID, Line_No. Điền B12 bằng target hoặc nguồn target AOV, rồi B13 ghi Warning/Critical tương ứng.

**Học lại:** Slide 4-5, `Handout/B2_KPI_Definition_Card.docx`.

![Hình 1 · Card còn thiếu ngưỡng và sai grain](Images/NX01_card_grain.png)
> **[CHỤP] NX01_card_grain.png** · File: `Homework/Cham_bai/VuLeThaiNgoc/b2_btvn_vu-le-thai-ngoc.xlsx` · Sheet: `KPI_Card` · Vùng: A5:B14 · Cần thấy: Nội dung B8/B9/B12 và thấy B13 trống. · Chụp Excel 365 Windows, zoom 100%, chỉ thao tác trên bản sao.

### 2. Insight cần sửa tỷ lệ và thêm hành động

**Bạn đã làm:** Insight!A3 ghi miền Nam hơn Bắc 49,9% và Trung 137,2%, câu kết chỉ nhận xét biên lợi nhuận.

**Đúng cần là:** Từ KPI_SUMIFS!H6:H8, Nam hơn Bắc 49,4781% và Trung 137,2983%; làm tròn một chữ số là 49,5% và 137,3%. Insight còn cần hàm ý.

**Hậu quả:** Người đọc nhận tỷ lệ chưa đúng và chưa biết cần kiểm tra hoặc làm gì tiếp theo.

**Cách sửa:** Tính `(doanh thu Nam/doanh thu miền so sánh-1)` trong ô trước khi viết. Thêm hành động kiểm tra cửa hàng hoặc cơ cấu giỏ hàng để giải thích số đơn/AOV miền Nam và tìm cách áp dụng cho miền còn lại.

**Học lại:** Slide 17 và 40, `Demo/B2_Demo_SUMIFS.xlsx`, sheet Demo_SUMIFS.

![Hình 2 · Insight cần sửa tỷ lệ và thêm hành động](Images/NX02_insight.png)
> **[CHỤP] NX02_insight.png** · File: `Homework/Cham_bai/VuLeThaiNgoc/b2_btvn_vu-le-thai-ngoc.xlsx` · Sheet: `Insight` · Vùng: A2:A3 · Cần thấy: Nội dung nguyên văn hai tỷ lệ và phần kết chưa có hành động. · Chụp Excel 365 Windows, zoom 100%, chỉ thao tác trên bản sao.

## Câu insight

Câu của bạn tại Insight!A3:

> Miền Nam đạt doanh thu cao nhất 142,1 triệu VNĐ, cao hơn miền Bắc 49,9% và miền Trung 137,2%, chủ yếu nhờ số đơn hàng cao nhất (447 đơn) và AOV cũng dẫn đầu (317.810 VNĐ). Tuy nhiên, Gross Margin giữa ba miền khá đồng đều ở mức 33,3%–34,0%, cho thấy hiệu quả biên lợi nhuận không chênh lệch nhiều giữa các vùng.

| Thành phần | Nhận xét |
|---|---|
| Con số có so sánh | Chưa đạt: 49,9% và 137,2% cần sửa thành 49,5% và 137,3%. Các số doanh thu, số đơn và AOV còn lại đúng. |
| Nguyên nhân khả dĩ | Đạt: đã giải thích bằng số đơn 447 và AOV 317.810 của miền Nam. |
| Hàm ý | Chưa đạt: chưa có hành động hoặc hướng kiểm tra cụ thể sau nhận định biên lợi nhuận. |

**Mức D4:** Developing. Phần insight đạt 1/3 thành phần theo bài T4.

## Việc cần làm trước Buổi 3

- [ ] Card còn thiếu ngưỡng và sai grain.
- [ ] Insight cần sửa tỷ lệ và thêm hành động.

Ảnh minh họa hiện là khung chờ chụp, đã ghi rõ file, sheet và vùng.
