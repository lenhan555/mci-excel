Chưa đạt - nộp lại

File: b2_btvn_to-thuXRBUKA.xlsx (bản trước lần 2)

**Làm tốt:**

- Sales có đủ 1.919 dòng. Product_Code ở F53, F230, F1311 và F1520 đã sạch; ba cột tra cứu không còn ô lỗi.
- Khối KPI!B2:G4 kéo SUMIFS đúng, H5 ra 296.965.000 VNĐ. COGS toàn bảng là 196.595.000 VNĐ.

**Cần sửa:**

1. KPI!I5 và J5 còn trống. Bạn bổ sung tổng số đơn `=SUM(I2:I4)` để ra 966, rồi AOV tổng `=H5/I5` để ra 307.417,184 VNĐ. Xem slide 16-17.
2. KPI_Card!B6 dùng tổng doanh thu toàn hệ thống chia số đơn của một miền. Tên bảng tblSales trong B6:B7 cũng khác tblSales_T4, cột Order_ID_line tại B7 không có trong Sales. Bạn viết lại công thức và nguồn cho cùng phạm vi AOV. Xem slide 5 và 16.
3. KPI_Card!B9 nêu Category và Phan_khuc là chiều phân tích AOV nhưng chưa xử lý đơn nhiều ngành hàng. Bạn ghi rõ giới hạn này, giữ miền, cửa hàng và tháng cho card hiện tại. B8 bổ sung kỳ báo cáo theo tháng. Xem slide 4-5.
4. Sales!P10 và 909 ô cùng phân khúc ghi "trung bình". Ngưỡng đúng; bạn đổi nhãn thành "Trung" trong IFS. Xem slide 7.

**Câu insight:**

insight!A1 đạt 3/3, mức D4 Competent. Con số, so sánh miền Nam và đề xuất combo có đủ. Bạn diễn đạt sức mua hai ngành hàng như nguyên nhân khả dĩ cần kiểm tra thêm, tránh khẳng định đã xác định được nguyên nhân.
