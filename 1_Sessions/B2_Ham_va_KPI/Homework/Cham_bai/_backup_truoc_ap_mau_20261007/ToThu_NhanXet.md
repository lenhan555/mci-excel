Chưa đạt - nộp lại

File: b2_btvn_to-thu_lan2.xlsx

**Làm tốt:**

- Bạn đã giữ đủ 1.919 dòng ở Sales. Bốn mã Product_Code tại G53, G230, G1311 và G1520 đúng, không còn dấu cách thừa hay lỗi tra cứu.
- Bảng KPI!B4:G6 dùng SUMIFS và tham chiếu hỗn hợp đúng. Tổng tại H7 là 296.965.000 VNĐ; số đơn tại C14 là 966.

**Cần sửa:**

1. KPI!D14 đang dùng `=SUM(D11:D13)`, ra 913.980,596 VNĐ. AOV toàn hệ thống phải là 296.965.000 chia 966, tức 307.417,184 VNĐ. Đổi thành `=B14/C14`. Đây là cộng ba AOV miền, nên số tổng bị sai. Xem slide 17, mục tổng chia tổng.
2. KPI!C11:C13 dùng SUMIFS trên cột LINE1. Cách này đang cho số đúng, nhưng đề yêu cầu COUNTIFS với Line_No = 1. Chẳng hạn C11 dùng `=COUNTIFS(Sales!$L:$L,A11,Sales!$B:$B,1)`, rồi kéo xuống. Xem slide 16.
3. KPI_Card!B6 ghi `=SUM(tblSales[Revenue]) /COUNTIFS(Sales!$K:$K,KPI!A2,Sales!$B:$B,1)`. Bảng thực tế tên tblSales_T4, K là Revenue và A2 trên KPI trống; tử số toàn hệ thống cũng chưa cùng phạm vi với mẫu số theo miền. Bạn sửa về tổng Revenue chia COUNTIFS(Line_No,1), rồi cập nhật B7 bằng đúng tên bảng và cột. Xem slide 5 và 16.
4. KPI_Card!B9 cho phép xem AOV theo Category và Phan_khuc. Một đơn có thể có nhiều ngành hàng, nên bạn cần ghi rõ cách đếm và phân bổ trước khi dùng các chiều này. Giữ các chiều miền, cửa hàng và tháng cho card hiện tại; B8 bổ sung mức báo cáo theo tháng. Xem slide 4-5.
5. Sales!Q10 và 909 ô cùng phân khúc đang ghi "trung bình". Ngưỡng IFS đúng, nhưng nhãn đề yêu cầu "Trung". Bạn đổi nhãn trong công thức để phân khúc thống nhất. Xem slide 7.

**Câu insight:**

insight!A1 đạt 3/3 thành phần, mức D4 Competent. Tỷ trọng miền Nam gần 48% và AOV 317.810 VNĐ/đơn khớp dữ liệu. Câu đã có giả thuyết về Bánh kẹo, Gia dụng và đề xuất combo. Bạn nên viết nguyên nhân dưới dạng giả thuyết cần kiểm tra, vì cơ cấu ngành hàng chưa chứng minh tác động nhân quả đến AOV. Giữ ý này khi sửa AOV tổng.
