Đạt - cần sửa nhỏ

File: b2_btvn_vu-le-thai-ngoc.xlsx

**Làm tốt:**

- Bạn giữ đủ 1.919 dòng ở Sales; các cột tra cứu, COGS, Gross_Margin và Phan_khuc đều là công thức, không còn lỗi dữ liệu trong các cột này.
- KPI_SUMIFS!H9 ra 296.965.000 VNĐ, B16 ra 966 đơn. B23 và C30 dùng tổng chia tổng, lần lượt ra AOV 307.417,184 VNĐ và Gross Margin 33,8%.

**Cần sửa:**

1. Insight!A3 ghi miền Nam cao hơn Bắc 49,9% và Trung 137,2%. Từ 142.061.000, 95.038.000 và 59.866.000 VNĐ, hai tỷ lệ đúng khi làm tròn một chữ số là 49,5% và 137,3%. Bạn tính `(doanh thu Nam/doanh thu miền so sánh-1)` trong ô trước khi đưa vào câu. Xem slide 17 và mục Viết insight của bài tập về nhà.
2. KPI_Card!B9 ghi "1 line_no = 1 đơn hàng". Một dòng là một sản phẩm trong một đơn; chỉ dòng Line_No = 1 được dùng để đếm đơn. Bạn sửa grain này và cập nhật nguồn ở B8 từ tblSales thành tblSales_T4, nêu Revenue, Order_ID, Line_No. Xem slide 4-5.
3. KPI_Card!B13 còn trống; B12 chưa nêu target AOV bằng số hay một nguồn target AOV cụ thể. Bạn điền target và ngưỡng Warning/Critical tương ứng. Xem slide 5 và mẫu B2_KPI_Definition_Card.
4. Sheet đang tên KPI_SUMIFS, trong đề yêu cầu KPI. Bạn đổi tên sheet để thống nhất tên bảng nộp. Các công thức SUMIFS hiện tại kéo đúng. Xem mục Bảng KPI 1 và Bảng KPI 2 trong De_bai.

**Câu insight:**

Insight!A3 đạt 1/3 thành phần, mức D4 Developing. Câu đã nêu nguyên nhân khả dĩ bằng số đơn và AOV của miền Nam. Hai tỷ lệ so sánh cần sửa; câu còn thiếu hành động tiếp theo. Bạn có thể nêu việc phân tích cửa hàng hoặc cơ cấu giỏ hàng để kiểm tra vì sao số đơn và AOV cao, rồi xác định cách cải thiện hai miền còn lại. Xem mục Viết insight trong bài tập về nhà.
