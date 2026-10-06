**Điểm: 8/10**

# Nhận xét bài tập về nhà Buổi 2 - Tô Thị Thu

File: `b2_btvn_to-thu_lan2.xlsx`

| # | Việc | Kết quả | Trọng số | % đạt / Thành phần |
|---|---|---|---:|---|
| 1 | Tra cứu XLOOKUP | Đạt | 2 | 100% |
| 2 | Cột tính toán | Một phần | 2 | 50% |
| 3 | Bảng KPI 1: Miền x Ngành hàng | Đạt | 1 | 100% |
| 4 | Bảng KPI 2: theo Miền | Một phần | 1 | 50% |
| 5 | KPI Definition Card (AOV) | Một phần | 1 | 50% |
| 6 | Insight | 3/3 thành phần | 3 | 3/3 |
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
| AOV tháng 4 (VNĐ) | 307.417 | 913.981 | Chưa |

Số tiền trong bảng làm tròn đến VNĐ để dễ đọc; phần đối soát dùng giá trị đầy đủ đến sáu chữ số thập phân.

## Làm tốt

- Sales giữ đủ 1.919 dòng; Product_Code tại G53, G230, G1311 và G1520 đã sạch, các cột tra cứu không còn lỗi.
- KPI!B4:G6 kéo SUMIFS đúng với $ hỗn hợp, H7 ra 296.965.000 VNĐ; C14 ra đúng 966 đơn.

## Điểm cần lưu ý

### 1. AOV dòng Tổng đang cộng ba miền

**Bạn đã làm:** KPI!D14 dùng `=SUM(D11:D13)`, ra 913.980,596 VNĐ. Số đơn ở C14 đã đúng 966.

**Đúng cần là:** AOV toàn hệ thống = 296.965.000 / 966 = 307.417,184 VNĐ.

**Hậu quả:** Cộng ba AOV miền làm con số tổng lớn gần ba lần, dù AOV từng miền đúng.

**Cách sửa:** Đổi D14 thành `=B14/C14`. C11:C13 hiện SUMIFS cột LINE1 cho số đúng; bạn đổi về COUNTIFS với Line_No = 1 theo yêu cầu, ví dụ `=COUNTIFS(Sales!$L:$L,A11,Sales!$B:$B,1)`, rồi kéo xuống.

**Học lại:** Slide 16-17, `Demo/B2_Demo_SUMIFS.xlsx`, sheet Demo_SUMIFS.

![Hình 1 · AOV dòng Tổng đang cộng ba miền](Images/NX01_aov_tong.png)
> **[CHỤP] NX01_aov_tong.png** · File: `Homework/Cham_bai/ToThiThu/b2_btvn_to-thu_lan2.xlsx` · Sheet: `KPI` · Vùng: A10:G14 · Cần thấy: Chọn D14 để thấy SUM(D11:D13); cả số tổng doanh thu, số đơn và AOV phải hiện. · Chụp Excel 365 Windows, zoom 100%, chỉ thao tác trên bản sao.

### 2. Card AOV chưa thống nhất công thức và nguồn

**Bạn đã làm:** KPI_Card!B6 ghi `=SUM(tblSales[Revenue]) /COUNTIFS(Sales!$K:$K,KPI!A2,Sales!$B:$B,1)`; B7 ghi tblSales và Order_ID_line.

**Đúng cần là:** Bảng đang tên tblSales_T4, các cột đúng là Revenue, Order_ID, Line_No. Tử số và mẫu số cần cùng kỳ và cùng miền.

**Hậu quả:** B6 tham chiếu tên bảng khác và trộn tổng toàn bảng với điều kiện miền; trên bản lan2, Sales!K:K còn là Revenue và KPI!A2 trống. Người dùng card sẽ không tính đúng AOV.

**Cách sửa:** B6 ghi `=SUM(tblSales_T4[Revenue])/COUNTIFS(tblSales_T4[Line_No],1)`. B7 sửa tên bảng/cột. B8 ghi dữ liệu một sản phẩm trong một đơn, KPI báo cáo theo tháng. B9 giữ miền, cửa hàng, tháng; muốn xem theo Category/Phan_khuc cần định nghĩa cách phân bổ đơn nhiều ngành. Target tại B11 cần ghi nguồn hoặc đánh dấu giả định.

**Học lại:** Slide 4-5 và 30, mẫu `Handout/B2_KPI_Definition_Card.docx`.

![Hình 2 · Card AOV chưa thống nhất công thức và nguồn](Images/NX02_card_aov.png)
> **[CHỤP] NX02_card_aov.png** · File: `Homework/Cham_bai/ToThiThu/b2_btvn_to-thu_lan2.xlsx` · Sheet: `KPI_Card` · Vùng: A3:B13 · Cần thấy: Nội dung rõ B6, B7, B8 và B9; bật Wrap Text trên bản sao nếu cần. · Chụp Excel 365 Windows, zoom 100%, chỉ thao tác trên bản sao.

### 3. Nhãn Phan_khuc khác nhãn yêu cầu

**Bạn đã làm:** Sales!Q10 và 909 ô cùng phân khúc ghi "trung bình". Ngưỡng của IFS vẫn đúng.

**Đúng cần là:** Đề dùng nhãn Cao / Trung / Thấp.

**Hậu quả:** Nhóm "trung bình" có nghĩa đúng nhưng khác chuỗi "Trung", dễ làm bảng lọc hoặc đối chiếu dùng nhãn chuẩn bỏ sót nhóm. Chữ hoa/thường không bị trừ riêng.

**Cách sửa:** Đổi nhánh IFS thành "Trung". Có thể dùng `=IFS([@Unit_Price]>=100000,"Cao",[@Unit_Price]>=30000,"Trung",TRUE,"Thấp")`.

**Học lại:** Slide 7, `Demo/B2_Demo_Logic.xlsx`, sheet Demo_IF.

## Câu insight

Câu của bạn tại insight!A1:

> Doanh thu tháng 4/2025 là 296,965,000 VNĐ (với khu vực miền Nam dẫn đầu hệ thống đạt 142,061,000 VNĐ,
>  chiếm gần 48% tổng doanh số), vì giá trị đơn hàng trung bình (AOV) tại miền Nam ở mức cao nhất toàn quốc
>  (317,810 VNĐ/đơn) nhờ sức mua vượt trội ở hai ngành hàng chủ lực là Bánh kẹo và Gia dụng, nên phòng Marketing cần tiếp tục tập trung ngân sách tối ưu hóa các chiến dịch bán kèm combo sản phẩm tại khu vực miền Nam để giữ vững đà tăng trưởng dòng tiền.

| Thành phần | Nhận xét |
|---|---|
| Con số có so sánh | Đạt: doanh thu 296.965.000, Nam 142.061.000 (47,8%) và AOV Nam 317.810 đúng. |
| Nguyên nhân khả dĩ | Đạt ở mức giả thuyết: Bánh kẹo/Gia dụng được nêu như lý do khả dĩ. Hai ngành này cũng lớn ở các miền khác, nên cần kiểm tra cơ cấu giỏ hàng và số đơn trước khi kết luận. |
| Hàm ý | Đạt: đã nêu Marketing và combo. "Giữ vững đà tăng trưởng" chưa có căn cứ thời gian trong một tháng; đổi thành việc cần kiểm chứng tiếp. |

**Mức D4:** Competent. Phần insight đạt 3/3 thành phần theo bài T4.

## Việc cần làm trước Buổi 3

- [ ] AOV dòng Tổng đang cộng ba miền.
- [ ] Card AOV chưa thống nhất công thức và nguồn.
- [ ] Nhãn Phan_khuc khác nhãn yêu cầu.

Ảnh minh họa hiện là khung chờ chụp, đã ghi rõ file, sheet và vùng.
