Báo cáo lớp B2 - chấm ngày 07/10/2026

Chưa có thư mục con theo lớp. Đã chấm 5 file của 4 học viên. Tô Thu có hai bản, tổng hợp theo học viên dùng b2_btvn_to-thu_lan2.xlsx; bản XRBUKA có nhận xét và bằng chứng riêng.

| Phạm vi | Đạt | Đạt - cần sửa nhỏ | Chưa đạt - nộp lại |
|---|---:|---:|---:|
| 4 học viên, lấy bản mới của Tô Thu | 0 | 2 | 2 |
| 5 file, tính cả hai bản Tô Thu | 0 | 2 | 3 |

Không quy điểm. Thang /10 trong kết quả cũ là thang tạm, chưa có thang và trọng số do người dùng đưa ở lần chấm này. Kết quả cũ được lưu trong _backup_20261007, gồm ToThiThu/, báo cáo, CSV và grade_b2.py cũ. Các file cũ tại ToThiThu/ và _DANH_SACH_ANH_NX.md vẫn được giữ nguyên; kết quả hiện tại ở các file nhận xét tại thư mục Cham_bai gốc.

| Học viên / lần nộp | Chỉ tiêu khớp | Insight | D4 | Kết quả |
|---|---:|---:|---|---|
| Tô Thu, lan2 | 7/8 | 3/3 | Competent | Chưa đạt - nộp lại |
| Tô Thu, XRBUKA | 6/8 | 3/3 | Competent | Chưa đạt - nộp lại |
| Vũ Lê Thái Ngọc | 8/8 | 1/3 | Developing | Đạt - cần sửa nhỏ |
| Phạm Như Tâm | 8/8 | 3/3 | Competent | Đạt - cần sửa nhỏ |
| Lê Lâm Tuyên | 1/8 | 0/3 theo đề T4 | Emerging | Chưa đạt - nộp lại |

Lỗi phổ biến, tính trên 4 học viên theo bản được chọn:

- 1/4 sai AOV tổng: Tô Thu, KPI!D14 = SUM(D11:D13), ra 913.980,5964378964 thay vì 307.417,1842650103. Đây là cộng AOV ba miền, cùng họ lỗi tổng tỷ lệ ở Dap_an_loi dòng 7; đề xuất bổ sung trường hợp SUM tỷ lệ bên cạnh AVERAGE trong script dựng homework.
- 1/4 chưa dùng đúng COUNTIFS theo yêu cầu: Tô Thu, KPI!C11:C13 dùng SUMIFS cột Sales!C:C (LINE1). Số miền và tổng 966 đúng. Dap_an_loi dòng 6 nói lỗi đếm dòng thành đơn; bài này đã xử lý grain bằng cờ nhưng khác hàm yêu cầu, chỉ đánh giá Một phần cho phần phương pháp, không tính thêm lỗi số đơn.
- 3/4 có vấn đề trong card AOV: Tô Thu đủ 10 trường nhưng KPI_Card!B6 trộn phạm vi tử/mẫu, B7 sai bảng/cột, B9 dùng Category và Phan_khuc không định nghĩa đơn nhiều ngành; Ngọc KPI_Card!B13 trống, B9 grain sai, B8 nguồn sai và B12 chưa định lượng target AOV; Tâm ` KPI_Card`!B6 chưa nêu kỳ báo cáo tháng, tên sheet có khoảng trắng đầu. Lê Lâm Tuyên là một trường hợp khác: KPI CARD!A2:D3 là ba KPI Q1, chưa có card AOV đủ 10 trường. Chưa có trong Dap_an_loi; đề xuất bổ sung lỗi card ở script dựng homework, giữ phần đọc hiểu chấm tay.
- 1/4 dẫn sai phần trăm và thiếu hàm ý: Ngọc, Insight!A3 ghi 49,9% và 137,2%. Tính từ KPI_SUMIFS!H6:H8 ra 49,4781034954% và 137,2982995356%, làm tròn 49,5% và 137,3%. Có nguyên nhân bằng số đơn/AOV; chưa có hành động tiếp theo. Đề xuất bổ sung vào Dap_an_loi của script dựng homework.
- 1/4 nộp khác kỳ: Lê Lâm Tuyên, Sales 6.187 dòng, Revenue 941.454.000 VNĐ, KPI_SUMIFS!B2:E2 là tháng 1-3/Quý 1. Không có khóa Order_ID + Line_No chung với dữ liệu T4; thiếu 1.919 khóa T4 và có 6.187 khóa Q1. Mọi lệch số liệu ở bài này gộp về lỗi khác kỳ, không tính nhiều lỗi gốc. Đề xuất thêm bước xác nhận kỳ vào danh sách nộp bài.
- 1/4 khác nhãn Phan_khuc nhưng đúng ngưỡng: Tô Thu lan2, Sales!Q10 và 909 dòng khác ghi "trung bình" thay vì "Trung". Không đánh giá sai ngưỡng. Bản XRBUKA cùng vấn đề ở Sales!P10.

Ba điểm nên chữa đầu buổi sau:

1. AOV tổng và Gross Margin % phải tính tổng chia tổng. Dùng KPI!D14 của bản lan2 để chỉ ra lỗi SUM tỷ lệ, sau đó đối chiếu KPI!B14/C14.
2. Card AOV phải cùng phạm vi ở tử/mẫu, đúng grain một dòng sản phẩm trong đơn và có kỳ báo cáo. Dùng card Tô Thu, Ngọc để phân biệt dữ liệu gốc với mức tổng hợp KPI.
3. Tính lại số so sánh trước khi viết insight, rồi nêu nguyên nhân khả dĩ và hành động. Dùng hai tỷ lệ sai tại Insight!A3 của Ngọc.

Bài cần GV xem:

- Lê Lâm Tuyên: cần xác nhận gửi nhầm file Q1 hay chưa làm bài T4. COGS/Gross_Margin/Phan_khuc có công thức nhưng không có khóa chung để đối chiếu ngưỡng với bài T4, ghi Không kiểm được ở việc 2. Các chỉ tiêu KPI2 trong JSON đọc cột T1 (B13/B20/B27), vì bài chỉ có bảng Q1; Q1 tổng ở E13/E20/E27. Không trình bày các số T1 này như số T4.
- Tô Thu: ưu tiên bản lan2 đã có tổng số đơn nhưng AOV tổng sai. Không suy đoán lịch sử nộp ngoài hậu tố tên file.
- Tài liệu gốc: Answer thiếu Insight; đề hiện đã sửa lỗi cài sẵn và có nhiều phần lời giải, cần GV xem trước khi giao lại bài.

Dấu hiệu bất thường:

Không phát hiện sheet Dap_an_loi hoặc Insight_mau trong 5 bài nộp. docProps/core.xml được đọc ở cả 5 bài; không thấy người tạo/người sửa cuối trùng tên một học viên khác trong 4 học viên được chấm. Lam Ngoc Tran ở hai bản Tô Thu chưa khớp một học viên trong danh sách này, nên không kết luận danh tính từ metadata. Hai bản cùng Tô Thu có insight và card giống nhau, là lịch sử nộp của cùng người. Không dùng dữ liệu sạch giống nhau làm dấu hiệu. Không kết luận chép bài.

Tên file / định dạng:

Cả 5 file đều mở được dạng .xlsx; cả 5 không theo đúng chuỗi B2_BTVN_<HoTen>.xlsx: b2_btvn_to-thu_lan2.xlsx, b2_btvn_to-thuXRBUKA.xlsx, b2_btvn_vu-le-thai-ngoc.xlsx, b2_btvn_xlsx_Pham Nhu Tam.xlsx, b2-le-lam-tuyen.xlsx. Tên ghi theo tên file, chưa xác minh danh sách lớp để bổ sung họ tên. Không trừ thêm ở các chỉ tiêu số vì lỗi tên file.

Thử script và lệch tài liệu:

- grade_b2.py đã sửa: bỏ điểm tạm, kiểm toàn cột, tìm tiêu đề từ dòng 1-10, đọc ArrayFormula, bảng theo nhãn thay vì tên sheet, hỗ trợ KPI2 tách nhiều khối, đọc drawing XML và VML. Kiểm toàn khối 18 ô, dòng/cột Tổng, dữ liệu theo khóa, COUNTIFS Line_No=1, công thức tổng tỷ lệ và từng miền. Sau thay đổi cuối, chạy lại Answer, đề và cả 5 bài.
- Answer: 8/8 Kiem_tra khớp. Việc 1-4 Đạt; KPI_Card!B4:B13 đủ 10 trường đúng AOV nên việc 5 Đạt sau đọc tay. Việc 6 Chưa đạt vì thiếu sheet Insight, trái De_bai!C11; Insight_mau!A3/A5 chỉ là tham chiếu chấm. Đây là lệch của Answer, chấm học viên theo De_bai.
- Đề hiện tại: 6/8 chỉ tiêu có giá trị khớp; tổng KPI1 thiếu, GM% thiếu. Sales đã có cột Region/Category/UnitCost/COGS/Gross Margin/Phan_khuc và Sheet1 có SUMIFS, COUNTIFS, AOV. Bốn mã Product_Code tại Sales!F53/F230/F1311/F1520 đều đã sạch. Không còn lỗi 4 mã có dấu cách như Dap_an_loi dòng 4 mô tả. Vì đề đã làm một phần, phép thử đề lỗi cài sẵn không đạt điều kiện dữ liệu ban đầu; những việc còn thiếu vẫn trả Chưa đạt. Ghi rõ ngoại lệ này thay vì sửa tiêu chí để làm đề trượt giả. Không sửa đề hoặc Answer.
- Phần chữ được đọc trong ô và XML drawing/VML của cả 5 file. Không có text box chứa chữ trong 5 file; chữ học viên ở các ô Insight và KPI CARD. Không mở Excel và không lưu lại bài nộp.
- Kết quả dùng giá trị cache đã lưu kết hợp đọc công thức và tính lại tổng từ dữ liệu, COGS từ Quantity x Unit_Cost, so sánh từng dòng và từng ô KPI. Không chạy lại engine Excel; không chứng minh trạng thái giao diện hay lịch sử thao tác.

Đầu ra: 5 file *_NhanXet.md, Tong_hop_B2.csv (UTF-8 BOM, 4 dòng học viên), 5 file *_Bang_chung.json, cham_tay_B2.json, Thu_script_B2.json và _evidence_b2.json. Báo cáo cũ và điểm cũ nằm trong backup, không dùng thay cho kết quả hiện tại.
