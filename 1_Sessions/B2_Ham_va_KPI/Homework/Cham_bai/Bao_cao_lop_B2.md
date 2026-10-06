Báo cáo chấm bài B2 - áp dụng mẫu ToThiThu, 07/10/2026

Đã đọc 5 file của 4 học viên; Tô Thị Thu dùng bản lan2. Không có thư mục lớp riêng. Mỗi học viên có thư mục chứa bản sao .xlsx, nhận xét .md và .pdf. Bản XRBUKA của Thu nằm trong ToThiThu/Lan1; mẫu và đầu ra trước được sao lưu trước khi ghi mới.

| Học viên | Điểm /10 (số nguyên) | Insight | D4 | Ghi chú |
|---|---:|---|---|---|
| Tô Thị Thu | 8 | 3/3 | Competent | AOV tổng sai; V2, V4, V5 đạt 50% |
| Vũ Lê Thái Ngọc | 8 | 1/3 | Developing | Card đạt 50%; tỷ lệ insight cần sửa và thiếu hàm ý |
| Phạm Như Tâm | 10 | 3/3 | Competent | Card đạt 50%, tổng được làm tròn lên |
| Lê Lâm Tuyên | 0 | 0/3 theo đề T4 | Emerging | Cần xác nhận file Q1 có gửi nhầm |

Thang điểm áp dụng theo mẫu được người dùng yêu cầu: V1=2, V2=2, V3=1, V4=1, V5=1, insight=3 (mỗi thành phần 1). V1-V5 Đạt 100%, Một phần 50%, chưa có phần việc đáp ứng T4 0%. Cộng theo trọng số trước rồi làm tròn half-up về số nguyên. Bảng điểm thành phần chỉ hiển thị trọng số nguyên và mức hoàn thành, tránh làm tròn riêng từng mục. Điểm 10 của Tâm vẫn kèm phần card cần sửa; kết quả làm tròn không xóa nhận xét đó. Không dùng kết luận yêu cầu nộp lại.

KPI2 Thu: số đơn các miền và tổng, AOV từng miền, GM% đều đúng; sai AOV dòng Tổng nên chấm Một phần 50% theo cùng nguyên tắc của mẫu thiếu ô tổng. Tên sheet, khoảng trắng và hoa/thường không bị trừ riêng khi phần việc đúng. V2 Thu khác chuỗi "trung bình" với nhãn "Trung", đúng ngưỡng nên 50% theo mẫu. Với Tuyên, chưa có bằng chứng T4 nên không ghi nhận cột tính Q1 là hoàn thành T4; các chênh lệch được gộp về một lỗi khác kỳ.

Lỗi phổ biến, tính trên 4 học viên theo bản được chọn:

- 1/4 sai AOV tổng: Tô Thị Thu, KPI!D14 = SUM(D11:D13), ra 913.980,5964378964 thay vì 307.417,1842650103. Đây là cộng AOV ba miền, cùng họ lỗi tổng tỷ lệ ở Dap_an_loi dòng 7; đề xuất bổ sung trường hợp SUM tỷ lệ bên cạnh AVERAGE trong script dựng homework.
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

- Lê Lâm Tuyên: cần xác nhận gửi nhầm file Q1 hay chưa làm bài T4. COGS/Gross_Margin/Phan_khuc có công thức nhưng không có khóa chung để đối chiếu ngưỡng với bài T4, bằng chứng lượt trước ở b2-le-lam-tuyen_Bang_chung.json tại thư mục gốc ghi Không kiểm được ở việc 2. Bảng điểm mới ghi 0 cho phần này vì chưa có bằng chứng T4. Các chỉ tiêu KPI2 trong JSON đọc cột T1 (B13/B20/B27), vì bài chỉ có bảng Q1; Q1 tổng ở E13/E20/E27. Không trình bày các số T1 này như số T4.
- Tô Thị Thu: ưu tiên bản lan2 đã có tổng số đơn nhưng AOV tổng sai. Không suy đoán lịch sử nộp ngoài hậu tố tên file.
- Tài liệu gốc: Answer thiếu Insight; đề hiện đã sửa lỗi cài sẵn và có nhiều phần lời giải, cần GV xem trước khi giao lại bài.

Dấu hiệu bất thường:

Không phát hiện sheet Dap_an_loi hoặc Insight_mau trong 5 bài nộp. docProps/core.xml được đọc ở cả 5 bài; không thấy người tạo/người sửa cuối trùng tên một học viên khác trong 4 học viên được chấm. Lam Ngoc Tran ở hai bản Tô Thu chưa khớp một học viên trong danh sách này, nên không kết luận danh tính từ metadata. Hai bản cùng Tô Thu có insight và card giống nhau, là lịch sử nộp của cùng người. Không dùng dữ liệu sạch giống nhau làm dấu hiệu. Không kết luận chép bài.

Thử script và tài liệu gốc:

grade_b2.py đã kiểm toàn cột và từng ô bảng, drawing/VML. Bằng chứng chấm trước được dùng lại; không mở hay tính lại 7 workbook trong bước chuyển mẫu. Answer 8/8 chỉ tiêu khớp, việc 1-5 Đạt sau đọc card; thiếu sheet Insight so với De_bai!C11. Đề 6/8 chỉ tiêu khớp, đã có cột tính và bảng Sheet1, bốn mã Sales!F53/F230/F1311/F1520 đã sạch nên không còn đúng trạng thái lỗi cài sẵn Dap_an_loi dòng 4. Các ngoại lệ vẫn lưu ở Thu_script_B2.json; không sửa đề/Answer.

5 bài nộp đều mở được .xlsx; tên file không theo chuỗi B2_BTVN_<HoTen>.xlsx. Tên nhận xét Thu theo mẫu đã có; cần đối chiếu danh sách lớp nếu muốn xác nhận họ tên đầy đủ cho mọi học viên. Không phát hiện dấu hiệu bất thường đủ căn cứ giữa các học viên. Hai bản cùng Thu giống card/insight được coi là lịch sử cùng một người, không phải dấu hiệu giữa hai học viên.

Ảnh minh họa: 7 khung chờ, chưa có ảnh chụp. MD/PDF ghi rõ nguồn bản sao, sheet, vùng và trạng thái. Không dựng ảnh Excel giả. Danh sách tại _DANH_SACH_ANH_NX.md; Images/ có sẵn ở mỗi thư mục.

Đầu ra chính: ToThiThu/, VuLeThaiNgoc/, PhamNhuTam/, LeLamTuyen/, Tong_hop_B2.csv (UTF-8 BOM, 4 dòng, Diem_10 nguyên). PDF dùng build_nhanxet_pdf.py hiện có theo từng file nhận xét; không gọi main để tránh lấy thư mục backup. Kiểm PDF và trạng thái tạo được ghi trong Kiem_tra_PDF_B2.json.

Bản dùng hiện tại nằm trong 4 thư mục ToThiThu, VuLeThaiNgoc, PhamNhuTam, LeLamTuyen. Năm MD tại thư mục Cham_bai gốc là lượt chấm trước; không phát các bản này thay cho nhận xét đã áp mẫu.
