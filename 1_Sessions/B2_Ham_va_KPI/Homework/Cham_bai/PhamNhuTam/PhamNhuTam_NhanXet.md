**Điểm: 10/10**

# Nhận xét bài tập về nhà Buổi 2 - Phạm Như Tâm

File: `b2_btvn_xlsx_Pham Nhu Tam.xlsx`

| # | Việc | Kết quả | Trọng số | % đạt / Thành phần |
|---|---|---|---:|---|
| 1 | Tra cứu XLOOKUP | Đạt | 2 | 100% |
| 2 | Cột tính toán | Đạt | 2 | 100% |
| 3 | Bảng KPI 1: Miền x Ngành hàng | Đạt | 1 | 100% |
| 4 | Bảng KPI 2: theo Miền | Đạt | 1 | 100% |
| 5 | KPI Definition Card (AOV) | Một phần | 1 | 50% |
| 6 | Insight | 3/3 thành phần | 3 | 3/3 |
| | **Tổng đã làm tròn** | | **10** | **10/10** |

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

- Sales đủ 1.919 dòng; bốn mã Product_Code tại G53, G230, G1311, G1520 sạch và cột tính toán đúng.
- KPI!B3:G5 kéo SUMIFS đúng. H6/B13 = 296.965.000, C13 = 966, D13 = 307.417,184 VNĐ.

## Điểm cần lưu ý

### 1. Card cần nêu kỳ báo cáo

**Bạn đã làm:** ` KPI_Card`!B6 ghi dữ liệu một sản phẩm trong một đơn, KPI tính "ở cấp đơn hàng"; B7 có miền, cửa hàng, tháng.

**Đúng cần là:** AOV được tổng hợp cho một kỳ báo cáo, ở bài này là tháng 4/2025; đơn hàng là đối tượng được đếm ở mẫu số.

**Hậu quả:** Người đọc chưa phân biệt giá trị mỗi đơn với AOV trung bình của các đơn trong tháng.

**Cách sửa:** B6 bổ sung: dữ liệu một sản phẩm trong một đơn; KPI tổng hợp theo tháng, có thể xem theo miền và cửa hàng. Tên sheet có khoảng trắng đầu chỉ nên sửa để gọn, không trừ thêm điểm. Target B9 cần nêu nguồn hoặc giả định.

**Học lại:** Slide 4-5, `Handout/B2_KPI_Definition_Card.docx`.

![Hình 1 · Card cần nêu kỳ báo cáo](Images/NX01_card_ky.png)
> **[CHỤP] NX01_card_ky.png** · File: `Homework/Cham_bai/PhamNhuTam/b2_btvn_xlsx_Pham Nhu Tam.xlsx` · Sheet: ` KPI_Card` · Vùng: A2:B11 · Cần thấy: Nội dung B6 với mức KPI cấp đơn và B7 có chiều tháng. · Chụp Excel 365 Windows, zoom 100%, chỉ thao tác trên bản sao.

## Câu insight

Câu của bạn tại Insight!A1:

> Miền Nam có AOV cao nhất (317.810 đồng/đơn), đồng thời là miền có số đơn cao nhất (447 đơn). Doanh thu miền Nam tập trung nhiều ở nhóm Bánh kẹo (45.695.000 đồng) và Gia dụng (34.927.000 đồng), cho thấy nhu cầu mua hàng và giá trị sản phẩm trong mỗi đơn có thể là yếu tố góp phần tạo AOV cao hơn. Vì vậy có thể ưu tiên tìm hiểu các yếu tố tạo ra giá trị đơn hàng cao tại miền Nam để đối chiếu và tìm phương án cải thiện áp dụng cho hai miền còn lại.

| Thành phần | Nhận xét |
|---|---|
| Con số có so sánh | Đạt: AOV Nam 317.810, 447 đơn, Bánh kẹo 45.695.000 và Gia dụng 34.927.000 đều đúng; có so sánh cao nhất. |
| Nguyên nhân khả dĩ | Đạt: viết nhu cầu và giá trị sản phẩm dưới dạng yếu tố khả dĩ, giữ đúng mức bằng chứng. |
| Hàm ý | Đạt: đã đề xuất tìm hiểu nguyên nhân AOV miền Nam để tìm cách cải thiện hai miền còn lại. |

**Mức D4:** Competent. Phần insight đạt 3/3 thành phần theo bài T4.

## Việc cần làm trước Buổi 3

- [ ] Card cần nêu kỳ báo cáo.

Ảnh minh họa hiện là khung chờ chụp, đã ghi rõ file, sheet và vùng.
