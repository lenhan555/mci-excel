Chưa đạt - nộp lại

File: b2-le-lam-tuyen.xlsx

**Làm tốt:**

- Bạn đã có các cột Category, Unit_Cost, COGS và Gross_Margin bằng công thức trong Sales.
- KPI_SUMIFS!E20 tính AOV Quý 1 bằng `=E6/E13`, đúng cách lấy tổng chia tổng.

**Cần sửa:**

1. Sales có 6.187 dòng dữ liệu, tổng Revenue 941.454.000 VNĐ. KPI_SUMIFS!B2:D2 ghi tháng 1, 2, 3 và E2 ghi Quý 1. Bài về nhà yêu cầu tháng 4/2025, 1.919 dòng và 296.965.000 VNĐ. Bạn kiểm tra lại file đã gửi, rồi làm trên B2_BTVN_MCI_Mart_T4_2025.xlsx. Các chênh lệch số liệu hiện tại cùng xuất phát từ khác kỳ dữ liệu. Xem slide 39 và bảng tự kiểm tra trong De_bai.
2. KPI_SUMIFS!A2:E6 đang là Miền theo tháng của Quý 1. Trên dữ liệu tháng 4, bạn dựng lại ba miền theo sáu ngành hàng, viết một SUMIFS với tham chiếu hỗn hợp rồi kéo. Xem slide 17.
3. KPI CARD!A2:D3 đang có ba KPI Quý 1, chưa có card AOV đủ 10 trường. Bạn tạo KPI_Card và điền theo mẫu B2_KPI_Definition_Card, ghi rõ công thức, nguồn, grain và kỳ tháng. Xem slide 4-5.
4. File chưa có sheet Insight. KPI CARD!A6:A17 đã có nhận xét và đề xuất cho Quý 1, nhưng bạn cần viết lại một câu dựa trên bảng tháng 4, có so sánh, nguyên nhân khả dĩ và hàm ý. Xem slide 39, mục Viết insight của bài tập về nhà.

**Câu insight:**

Phần chữ hiện tại ở KPI CARD!A6:A17 chưa đáp ứng insight tháng 4, ghi 0/3 theo yêu cầu bài này, mức D4 Emerging. Bạn giữ cách dẫn số theo miền, đổi sang đúng dữ liệu tháng 4, rồi thêm một giả thuyết cần kiểm tra và hành động cụ thể gắn với giả thuyết đó.
