# Báo cáo lớp — Bài tập về nhà Buổi 1 (BTVN_MCI_Mart_T4_2025)

Dành cho giảng viên. Chấm bằng `2_Dataset/_generator/grade_b1.py`, đáp án lấy từ `Homework/B1_BTVN_MCI_Mart_T4_2025_Answer.xlsx`. File học viên chỉ đọc, không sửa. Khóa không có yêu cầu nộp lại, nên báo cáo chỉ chấm điểm, không có kết luận "nộp lại" hay "cần sửa nhỏ".

## 1. Bảng điểm

7 file nộp, không có file không mở được; **6 bài đúng đề tháng 4/2025** và **1 bài nộp nhầm đề** (Lê Lâm Tuyền, bộ Quý 1 của bài lớp, chưa tính là đã nộp đúng đề). Điểm tối đa 10; điểm trung bình lớp **8,0** (6 bài đúng đề, không tính bài Tuyền).

| Học viên | V1 Cấu trúc | V2 Kiểu dữ liệu | V3 Table | V4 Revenue | V5 C5 + Pct | V6 Validation | V7 Circle | V8 Insight | **Tổng /10** |
|---|---|---|---|---|---|---|---|---|---|
| Phạm Như Tâm | 1,25 | 1,25 | 1,25 | 1,25 | 1,25 | 0,63 | 1,25 | 1,25 (3/3) | **9,38** |
| Tô Thị Thu | 1,25 | 0,63 | 1,25 | 1,25 | 1,25 | 0,63 | 1,25 | 0,83 (2/3) | **8,33** |
| Nguyễn Thị Bích Trâm | 1,25 | 1,25 | 1,25 | 1,25 | 1,25 | 0,63 | 1,25 | 0,00 (0/3, trống) | **8,13** |
| Vũ Lê Thái Ngọc | 0,63 | 1,25 | 1,25 | 1,25 | 1,25 | 0,63 | 1,25 | 0,42 (1/3) | **7,92** |
| Nguyễn Đức Minh | 1,25 | 0,63 | 1,25 | 1,25 | 0,63 | 0,63 | 1,25 | 0,42 (1/3) | **7,29** |
| Vũ Thị Thanh Thủy | 1,25 | 0,63 | 0,63 | 1,25 | 0,63 | 0,63 | 1,25 | 0,42 (1/3) | **6,67** |
| Lê Lâm Tuyền (NỘP NHẦM ĐỀ: bài lớp Q1, chỉ tham khảo, không tính vào điểm lớp) | 1,25 | 0,00 | 0,63 | 1,25 | 0,63 | 0,63 | 0,00 | 0,83 (2/3) | **5,21** |

(Điểm tính chính xác đến 3 chữ số; bảng làm tròn 2 chữ số theo kiểu half-up.)

**Thang điểm: thang TẠM do trợ giảng tự dựng, cần giảng viên xác nhận.** Đề (`De_bai`) và `Kiem_tra` không ghi điểm hay trọng số. Cách tính đã dùng: 8 việc trong bảng "Việc cần làm", mỗi việc 1,25 điểm (tổng 10); Đạt = 100%, Một phần = 50%, Chưa đạt = 0; việc 8 (insight) = 1,25 × số thành phần đạt (con số, nguyên nhân, hàm ý) / 3. Các chỉ tiêu `Kiem_tra` không có điểm riêng: chỉ tiêu lệch đã thể hiện ở việc gây ra nó (một lỗi gốc chỉ tính một lần). Nếu giảng viên có thang khác, sửa hằng số `DIEM_MOI_VIEC` và hàm `diem_viec()` trong `grade_b1.py` rồi chấm lại.

Phần chữ (insight) do trợ giảng đọc và chấm tay; số thành phần ghi trong `Cham_bai/insight_cham_tay_B1.json`.

Chi tiết từng chỉ tiêu Kiem_tra và điểm từng việc nằm trong `Tong_hop_B1.csv`.

## 2. Lỗi phổ biến

Trong 6 bài làm đúng file đề (không tính Tuyền):

| Lỗi | Số bài | Dòng Dap_an_loi | Ghi chú |
|---|---|---|---|
| Order_Date còn dạng text dd/mm/yyyy | 3/6 (Minh 12 ô, Thu 12 ô, Thủy 7 ô) | #7 | Phổ biến nhất. Minh và Thu bỏ qua cả 12 ô; không ai bị đảo ngày/tháng khi đã chuyển |
| Quantity còn dạng text | 1/6 (Thủy 4/10 ô) | #6 | 5 bài còn lại dùng Convert to Number thành công |
| Còn tiêu đề báo cáo ở dòng 1–3 | 1/6 (Ngọc) | #1 | Table bắt đầu ở dòng 4; mọi chỉ tiêu vẫn khớp |
| Dòng Cửa hàng/Cộng/TỔNG CỘNG, dòng lặp tiêu đề, dòng trống | 0/6 | #2–#5 | Cả lớp làm sạch tốt |
| Sửa ST5, mã sản phẩm thiếu số 0, Quantity âm | 0/6 sai | #8–#10 | 6/6 sửa đúng giá trị, không ai xóa dòng |

**Lỗi mới, chưa có trong Dap_an_loi** (đề xuất bổ sung trong `build_homework.py`, phần dựng `Dap_an_loi`):

| Lỗi | Số bài | Ví dụ |
|---|---|---|
| Data Validation không có thông báo lỗi (Error Alert) ở ít nhất một cột | 7/7. Không bài nào đủ cả ba cột; Thu có Quantity, Tâm có Store_Code và Quantity | Trâm viết nội dung ở **Input Message** thay vì Error Alert |
| List không lấy từ sheet Store / Product | 2/7 | Minh gõ danh sách ở Sales!N1:N12 và P1:P40; Tuyền dùng sheet Store tự tạo 6.187 dòng |
| List bị ẩn mũi tên (In-cell dropdown bỏ tick) | 1/7 (Tâm) | `showDropDown="1"` ở Store_Code và Product_Code |
| Tên Table/sheet dữ liệu khác đề (trừ điểm) | 2/7 | Thủy `tbsale_T4`/`sale`; Tuyền `tblSales`. Chỉ khác hoa/thường (`sales`, `revenue`, `Pct_share`) không trừ |
| Ô tổng doanh thu không ở Kiem_tra!C5, hoặc gõ tay | 3/7 | Minh để ở Sales!L1 (không có sheet Kiem_tra); Thủy để ở C2 (C5 là tổng Pct) và gõ số 0 ở C6:C9; Tuyền gõ tay C5 |
| Insight ghi sai số dòng | 4/5 bài có insight | 1920 (Minh, Thủy), 1948 (Thu), 3448 (Ngọc, nhầm với tổng Quantity). Chỉ Tâm ghi 1.919 |
| Insight nêu "số loại lỗi" tùy ý | 5/6 có nêu (2, 3, 3, 3, 4) | Dap_an_loi có 10 loại; slide 40 không định nghĩa "loại lỗi" |
| Làm nhầm file | 1/7 (Tuyền) | Nộp file Quý 1 của buổi thực hành |

## 3. Ba điểm nên chữa đầu buổi sau

1. **Order_Date text và cách kiểm tra "còn sót ô nào".** 3/6 bài còn 7–12 ô. Chỉ cách kiểm: lọc cột ngày hoặc thêm cột `=ISTEXT()`; và cách chuyển không phụ thuộc cài đặt máy (`=DATE(RIGHT(…,4),MID(…,4,2),LEFT(…,2))`) bên cạnh Text to Columns DMY. Gắn với đối soát: ô "số ô Order_Date dạng text" phải ra 0.
2. **Data Validation cho đủ ba phần.** (a) Source trỏ sheet Store / Product, không gõ danh sách; (b) tab **Error Alert** (khác Input Message) có chữ tiếng Việt; (c) giữ tick In-cell dropdown. Đây là việc mất điểm ở cả 7 bài.
3. **Câu insight có con số đúng và có so sánh.** 4/5 bài có insight ghi sai số dòng; số loại lỗi mỗi người đếm khác nhau. Chỉ cách lấy số dòng (Count của Order_ID) và dùng con số **trước và sau khi sửa** làm phần so sánh; chỉ Tâm đạt đủ ba thành phần, và cũng chưa có con số so sánh.

## 4. Bài cần giảng viên xem

- **Lê Lâm Tuyền (NỘP NHẦM ĐỀ, chưa tính là đã nộp đúng đề; 5,21 chỉ là điểm tham khảo trên bài lớp Q1):** nộp file thực hành Quý 1 (6.187 dòng, khớp cả 6.187 khóa với file đáp án Q1), không có sheet De_bai và Product. Đã đánh dấu "CHƯA NỘP ĐÚNG ĐỀ" trong `Tong_hop_B1.csv` và đầu nhận xét; cần hỏi lại có tải nhầm file không. Điểm chấm trên dữ liệu Q1: việc 7 tính 0 vì còn 5 ô bị khoanh (3 ST3, 2 Quantity âm) và việc 2 tính 0 (15 Quantity text, 20 ngày text); script gốc ra "Không kiểm được" ở việc 7 vì khác bộ dữ liệu. Ô E6 ghi ST02 trong khi dữ liệu gốc là ST01. Giảng viên quyết định có tính điểm bài này không, hay chấm lại khi học viên gửi đúng file.
- **Nguyễn Thị Bích Trâm:** sheet `Question` có 2 câu hỏi cho giảng viên (chọn lại vùng đã validated; lọc Circle Invalid Data). Đã trả lời tạm trong nhận xét (Find & Select → Data Validation; không lọc được vòng đỏ, dùng cột kiểm tra hoặc Conditional Formatting + Filter by Color). Xin duyệt nội dung.
- **Phạm Như Tâm:** mức D4 Competent theo dòng "Chấm:" (đủ ba thành phần, số 1.919 và -2 khớp), nhưng "3 loại lỗi" lệch 10 loại cài sẵn và chưa có con số so sánh; giảng viên cân nhắc, điểm việc 8 hiện là trọn 1,25.
- **Tô Thị Thu:** `docProps/core.xml` ghi người sửa cuối "Lam Ngoc Tran", không trùng học viên nào trong lớp. Chỉ ghi nhận, không kết luận.
- Không có bài không mở được, không có file chứa sheet hoặc tiêu đề của file đáp án.

**Dấu hiệu bất thường (chỉ để giảng viên biết):** không có. Các câu insight giống nhau 0,60–0,65 (Thu–Ngọc, Ngọc–Thủy, Thu–Thủy, Thu–Tâm) vì cùng điền khung slide 40; không có câu nào giống hệt. Minh và Thu cùng bỏ qua 12 ô ngày: hệ quả của dữ liệu cài sẵn giống nhau, không phải dấu hiệu.

## 5. Tên file

Cả 7 file đều không đúng quy ước `B1_BTVN_<HoTen>.xlsx`: tên có hậu tố ngẫu nhiên do cổng nộp bài thêm vào, nên không tính là lỗi của học viên. Hai file có dạng khác hẳn: `b1_btvn_xlsx_PhamNhuTam.xlsx` và `bai-tap-ve-nha-buoi-1_-lam-tuyenGSIO38_LeLamTuyen.xlsx` (file sau không có "btvn" trong tên, cùng với việc làm trên dữ liệu Q1). Đều là `.xlsx`, mở được.

## 6. Lệch giữa tài liệu khóa và file đáp án

1. **File đề trong thư mục đã bị thay.** `Homework/B1_BTVN_MCI_Mart_T4_2025.xlsx` (sửa 30/09/2026, người sửa cuối Leo Pham) đang là bản đã làm sạch: sheet `BaoCao_T4_2025` có Table, Revenue, Pct_share, không còn lỗi cài sẵn. Bản gốc có lỗi là bản trong git (HEAD); phép thử script dùng bản HEAD. Nếu phát lại đề, dựng lại từ `build_homework.py` hoặc khôi phục bản git.
2. **Dap_an_loi mục #9** ghi ví dụ "P25 thay vì P025", nhưng 3 ô lỗi thật là `P19` (HD2504-00601 dòng 1), `P26` (HD2504-00383 dòng 2), `P12` (HD2504-00413 dòng 4). Nên sửa ví dụ cho khớp.
3. Dap_an_loi mục #1 ghi "Xóa dòng 1–3" cho tiêu đề và dòng trống: khớp với file (tiêu đề dòng 1–2, dòng trống dòng 3, tiêu đề cột ở dòng 4).
4. File `_Answer` thiếu sheet `Insight` (chỉ có `Insight_mau`), và các ô công thức không có giá trị đã lưu (ghi bằng openpyxl). Script tự tính lại và ghi chú. Data Validation của `_Answer` đủ, đúng điều đề yêu cầu.
5. Đề và `Kiem_tra` không có thang điểm (xem mục 1).
6. Số liệu đã kiểm lại trên file đề gốc: tổng Quantity cột H khi chưa làm sạch = 10.309 (khớp Insight_mau); dòng rác 27 (12 + 12 + 1 + 1 + 1); lỗi cài sẵn 10 + 12 + 2 + 3 + 2 đúng như Dap_an_loi.
7. Slide 40 và De_bai dùng cụm "File case" trong câu insight; học viên dùng cho file tháng 4 không sai, nhưng Tuyền ghi số của bộ Q1. Cân nhắc ghi rõ "file BTVN T4" trong khung câu.

## 7. Quyết định chấm (giảng viên đã chốt / đang áp dụng)

**Hai quy tắc giảng viên chốt, đã áp dụng cho toàn bộ lớp:**
1. **Tên sheet `Kiem tra`, `Kiem_tra`, `Kiểm tra` (và các biến thể không dấu / hoa thường) coi là giống nhau**: đạt, không trừ, không tính Một phần. Đã sửa trong `grade_b1.py` (sheet Kiểm tra được tìm theo tên chuẩn hóa, không còn trừ vì tên). Ảnh hưởng: Trâm và Tâm ở việc 5 từ Một phần lên Đạt.
2. **Khóa học không có yêu cầu nộp lại**: bỏ mọi kết luận "Chưa đạt – nộp lại" / "Đạt – cần sửa nhỏ"; chỉ chấm điểm. Nhận xét học viên không dùng ngôn ngữ yêu cầu làm lại; chỉ nêu điểm, chỗ mất điểm và cách xử lý cho lần sau.

**Các quy ước trợ giảng đang dùng (đề nghị duyệt):**
- Thang điểm tạm 8 × 1,25 (mục 1).
- Tên Table / sheet dữ liệu khác đề (`tbsale_T4`, `sale`, `tblSales`) vẫn tính Một phần ở việc 3; chỉ khác hoa/thường thì không trừ. Tên sheet `Insight` viết sai chính tả (`insinght`) không trừ điểm.
- Table bật Total Row (Thu) hợp lệ, không phải dòng thừa.
- Một lỗi gốc chỉ tính một lần: Quantity text tính ở việc 2, không trừ lại ở việc 7; dòng tiêu đề còn lại tính ở việc 1, không trừ lại ở việc 3.
- Insight có nhưng thiếu thành phần được điểm theo tỉ lệ; không có = 0.

## 8. Script và kiểm thử

- `2_Dataset/_generator/grade_b1.py`, đã sửa ở lượt này: quy tắc tên sheet Kiem_tra, bỏ kết luận nộp lại, thêm tính điểm (`diem_viec`, `ap_dung_insight`, tùy chọn `--insight-file`). Cả lớp được chấm lại bằng bản mới.
- **Phép thử 1, trên file `_Answer`:** 9/9 chỉ tiêu khớp, việc 1–7 Đạt (8,75/10 chưa tính insight); việc 8 ra "Cần chấm tay".
- **Phép thử 2, trên file đề gốc (git HEAD):** việc 1–8 đều Chưa đạt, điểm 0; 9/9 chỉ tiêu lệch đúng như Dap_an_loi (10 Quantity text, 12 ngày text, 2 ST5, 3 mã sản phẩm sai, 2 Quantity âm; tổng Quantity chưa sạch 10.309; 27 dòng rác).
- `--check-seed` xác nhận danh sách lỗi cài sẵn trong script trùng kết quả so file đề với `_Answer`.

## 9. Thư mục kết quả

- `Cham_bai/<HoTen>/<HoTen>_NhanXet.md` (mỗi học viên một thư mục) kèm một bản sao file đã nộp; file gốc trong `Bai_nop/` không bị sửa hay di chuyển.
- `Cham_bai/Tong_hop_B1.csv` (UTF-8 có BOM): có cột Diem_10 (tổng điểm) và điểm từng việc.
- `Cham_bai/insight_cham_tay_B1.json`: số thành phần insight trợ giảng chấm tay.
- `Cham_bai/Bao_cao_lop_B1.md` (file này).
