# Báo cáo lớp — Bài tập về nhà Buổi 2 (BTVN_MCI_Mart_T4_2025)

Dành cho giảng viên. Chấm lại ngày 04/10/2026 theo cách của B1 (thang /10, mỗi học viên một thư mục, không có kết luận "nộp lại"). Chấm bằng `2_Dataset/_generator/grade_b2.py`, đáp án lấy từ `Homework/B2_BTVN_MCI_Mart_T4_2025_Answer.xlsx`. File học viên chỉ đọc, không sửa. Bản chấm lượt trước (kết luận "Chưa đạt – nộp lại") được thay bằng bản này.

## 1. Bảng điểm

1 bài nộp, mở được. Điểm tối đa 10; điểm trung bình lớp 8 (n = 1).

| Học viên | V1 Tra cứu | V2 Cột tính | V3 KPI 1 | V4 KPI 2 | V5 KPI_Card | V6 Insight | **Tổng /10** |
|---|---|---|---|---|---|---|---|
| Tô Thị Thu (suy từ tên file) | 2 | 1 | 1 | 0,5 | 0,5 | 3 (3/3) | **8** |

**Thang điểm: thang TẠM do trợ giảng tự dựng, cần giảng viên xác nhận.** Đề (`De_bai`) và `Kiem_tra` không ghi điểm hay trọng số. Cách tính (trọng số nguyên): việc 1 và 2 mỗi việc 2 điểm, việc 3, 4, 5 mỗi việc 1 điểm; Đạt = 100%, Một phần = 50%, Chưa đạt = 0; việc 6 (insight) = 3 điểm, mỗi thành phần đạt (con số, nguyên nhân, hàm ý) được 1 điểm. Các chỉ tiêu `Kiem_tra` không có điểm riêng: chỉ tiêu lệch thể hiện ở việc gây ra nó (một lỗi gốc chỉ tính một lần). Nếu giảng viên có thang khác, sửa `TRONG_SO` và `diem_viec()` trong `grade_b2.py` rồi chấm lại.

Việc 5 (KPI_Card) và việc 6 (insight) do trợ giảng đọc và chấm tay; kết quả ghi trong `Cham_bai/cham_tay_B2.json`.

## 2. Lỗi phổ biến (n = 1 bài)

| Lỗi | Số bài | Dòng Dap_an_loi | Ghi chú |
|---|---|---|---|
| Dòng Tổng của bảng KPI 2 để trống Số đơn và AOV (KPI!I5, J5) | 1/1 | Gần #3 và #4 | Lỗi **bỏ trống**, không phải ra 1.919 hay dùng AVERAGE. GM% dòng Tổng đúng (tổng ÷ tổng). Làm Tổng số đơn và AOV lệch 2 chỉ tiêu Kiem_tra; tính một lần ở V4 |
| Nhãn Phan_khuc khác đề: "trung bình" (910 dòng), "thấp" (440 dòng) thay vì "Trung", "Thấp"; ngưỡng đúng | 1/1 | Chưa có | V2 Một phần |
| KPI_Card: công thức tham chiếu một miền (`KPI!A2`) với tử số toàn bảng; sai tên bảng (`tblSales`) và cột (`Order_ID_line`) | 1/1 | Chưa có | V5 Một phần |
| KPI_Card: Dimensions liệt kê Category, Phan_khuc cho AOV; Target 304.000 không nêu nguồn | 1/1 | Chưa có | Đáp án mẫu: AOV không xem theo ngành (một đơn có nhiều ngành), Target 300.000 |
| Tham chiếu cột nguyên (`Sales!$J:$J`, `Store!A:A`) thay vì tên bảng | 1/1 | Chưa có | Số đúng, không trừ điểm |

Đã làm đúng Dap_an_loi #1 (mã dính dấu cách, không còn #N/A, không IFERROR) và #2 (`$` hỗn hợp, tổng khớp 296.965.000).

**Lỗi mới, chưa có trong Dap_an_loi** (đề xuất bổ sung trong script dựng homework B2): nhãn Phan_khuc sai chữ; AOV bị liệt kê theo ngành hàng trong KPI_Card; dòng Tổng bỏ trống thay vì tính sai.

## 3. Ba điểm nên chữa đầu Buổi 3

1. **Dòng Tổng của bảng tỷ lệ là tổng ÷ tổng, và không bỏ trống.** AOV tổng = 296.965.000 ÷ 966 = 307.417, khác trung bình ba AOV miền; Số đơn tổng cộng theo miền (317 + 202 + 447 = 966), không đếm số dòng (1.919).
2. **Nhãn phân loại do IFS sinh ra là khóa của Pivot và bộ lọc.** "Trung" khác "trung bình" là hai nhóm khác nhau. Buổi 3 dùng Phan_khuc ngay nên dễ lộ lỗi.
3. **KPI Card AOV: grain và Dimensions.** Grain 1 dòng = 1 sản phẩm trong 1 đơn, nên AOV không chia theo ngành hay phân khúc (đơn bị đếm lặp); công thức ghi cho cả khoảng báo cáo, không theo một miền; Target phải nêu nguồn hoặc ghi là giả định.

## 4. Bài cần giảng viên xem

- **Họ tên không có trên tên file.** `b2_btvn_to-thuXRBUKA.xlsx` cùng phần "to-thu" với file B1 của Tô Thị Thu, và `docProps/core.xml` ghi người sửa cuối "Lam Ngoc Tran", trùng với file B1 của bạn Thu. Đây là suy luận, giảng viên xác nhận trước khi gửi nhận xét.
- **Quyết định V4:** thiếu hai ô Tổng nhưng từng miền đúng nên tính **Một phần** (0,5), không phải Chưa đạt. Lượt chấm trước tính Chưa đạt vì quy tắc cũ "thiếu ô là Chưa đạt"; script giờ phân biệt "thiếu cột" (Chưa đạt) và "thiếu ô ở dòng Tổng" (Một phần). Giảng viên xác nhận.
- **Insight 3/3 có điều kiện:** con số và hàm ý đạt; nguyên nhân (Bánh kẹo, Gia dụng) có nhưng chưa phân biệt được miền Nam, vì hai ngành này là hai ngành lớn nhất ở cả ba miền (đọc từ bảng KPI 1 của bài). Nếu giảng viên chấm chặt thì nguyên nhân là Một phần, insight 2/3, tổng giảm còn 7.
- **Target 304.000** ở KPI_Card không có trong đề; đáp án mẫu dùng 300.000. Giảng viên quyết định có chấp nhận con số tự đặt hay không (đã nêu trong nhận xét: cần ghi nguồn hoặc giả định).
- Không có dấu hiệu bất thường: không có sheet `Dap_an_loi` hay `Insight_mau`. Một bài nên không so được độ giống giữa các bài.

## 5. Tên file

`b2_btvn_to-thuXRBUKA.xlsx` không đúng quy ước `B2_BTVN_<HoTen>.xlsx`: có hậu tố ngẫu nhiên do cổng nộp bài thêm vào, không tính là lỗi của học viên. Là `.xlsx`, mở được.

## 6. Lệch giữa tài liệu khóa và file đáp án

1. `Kiem_tra` của `_Answer` trỏ `KPI!H8`, `KPI!C16`, `KPI!E16`, `KPI!D16` theo bố cục riêng của đáp án; học viên đặt bảng KPI ở vị trí khác nên script tìm theo nhãn hàng/cột. Nếu muốn chấm theo ô thì đề cần ghi rõ bố cục bảng.
2. `Kiem_tra` không có chỉ tiêu cho Phan_khuc (ví dụ số dòng "Cao", "Trung", "Thấp"); nên thêm để bắt lỗi nhãn bằng số. Chưa có lỗi này trong `Dap_an_loi`.
3. `_Answer` không có sheet `Insight` (chỉ `Insight_mau`), nên việc 6 ra "Chưa đạt" khi chạy script trên chính file này; chỉ ảnh hưởng phép thử.
4. `De_bai` nói "từ tblStore", "từ tblProduct" nhưng không bắt buộc dùng tham chiếu bảng; bài nộp dùng `Store!A:A`, kết quả đúng. Cân nhắc ghi rõ nếu muốn học viên dùng tên bảng.
5. Đề và `Kiem_tra` không có thang điểm (xem mục 1).

## 7. Quyết định chấm (đang áp dụng, đề nghị duyệt)

Hai quy tắc giảng viên đã chốt ở B1, áp dụng tiếp cho B2:
1. Tên sheet khác hoa/thường hoặc dấu (`insight` thay vì `Insight`) coi là giống nhau, không trừ điểm.
2. Khóa không có yêu cầu nộp lại: chỉ chấm điểm, nhận xét nêu chỗ mất điểm và cách xử lý cho lần sau.

Quy ước trợ giảng dùng thêm ở B2:
- Thang tạm 6 × 10/6 (mục 1).
- Thiếu ô ở dòng Tổng nhưng từng miền đúng: Một phần. Thiếu hẳn cột Số đơn, AOV hoặc Gross Margin %: Chưa đạt.
- Nhãn Phan_khuc đúng ngưỡng nhưng khác chữ: Một phần ở V2, không tính lại ở V6.
- Một lỗi gốc chỉ tính một lần: hai ô Tổng bỏ trống tính ở V4, không trừ lại ở Kiem_tra.
- Insight có nhưng thiếu thành phần được điểm theo tỉ lệ; không có sheet = 0.

## 8. Script và kiểm thử

- `2_Dataset/_generator/grade_b2.py`, sửa ở lượt này: đổi tên hàm chấm một bài thành `grade`; thêm `--folder` (chấm cả thư mục) và `--cham-tay <json>` (kết quả chấm tay V5, V6); thêm tính điểm (`TRONG_SO`, `diem_viec`, `ap_cham_tay`), so độ giống insight; V4 phân biệt thiếu cột và thiếu ô Tổng. Script chưa được commit vào git; bản trước lượt này không còn lưu.
- **Phép thử 1, trên file `_Answer`:** 8/8 chỉ tiêu Kiem_tra khớp, việc 1–4 Đạt; việc 5 cần chấm tay; việc 6 ra Chưa đạt vì `_Answer` không có sheet Insight (mục 6.3).
- **Phép thử 2, trên file đề (chưa làm):** việc 1–6 đều Chưa đạt; 7/8 chỉ tiêu lệch (số dòng 1.919 vẫn khớp vì đề đã có đủ dòng).
- **Bài nộp:** 6/8 chỉ tiêu khớp (lệch: Tổng số đơn và AOV, do I5, J5 trống). Bảng số liệu tự tính lại từ sheet Sales của bài: Bắc 95.038.000 / 317 đơn, Trung 59.866.000 / 202 đơn, Nam 142.061.000 / 447 đơn, khớp bảng KPI của học viên.

## 9. Thư mục kết quả

- `Cham_bai/ToThiThu/ToThiThu_NhanXet.md` kèm một bản sao file đã nộp; file gốc trong `Bai_nop/` không bị sửa hay di chuyển.
- `Cham_bai/Tong_hop_B2.csv` (UTF-8 có BOM): cột Diem_10, các chỉ tiêu Kiem_tra, kết quả từng việc.
- `Cham_bai/cham_tay_B2.json`: kết quả chấm tay V5 và số thành phần insight.
- `Cham_bai/Bao_cao_lop_B2.md` (file này).
