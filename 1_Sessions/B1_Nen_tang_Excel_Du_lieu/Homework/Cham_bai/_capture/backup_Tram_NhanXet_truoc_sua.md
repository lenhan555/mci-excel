**Điểm: 8,1/10**

# Nhận xét bài tập về nhà Buổi 1 — Nguyễn Thị Bích Trâm

| # | Việc | Kết quả | Điểm |
|---|---|---|---|
| 1 | Làm sạch cấu trúc | Đạt | 1,25/1,25 |
| 2 | Kiểu dữ liệu (Quantity, Order_Date) | Đạt | 1,25/1,25 |
| 3 | Excel Table + tên sheet | Đạt | 1,25/1,25 |
| 4 | Cột Revenue | Đạt | 1,25/1,25 |
| 5 | Kiem_tra!C5 + Pct_Share | Đạt | 1,25/1,25 |
| 6 | Data Validation | Một phần | 0,63/1,25 |
| 7 | Circle Invalid Data | Đạt | 1,25/1,25 |
| 8 | Insight | Insight 0/3 thành phần | 0,00/1,25 |
| | **Tổng** | | **8,13/10** |

Thang điểm: 8 việc, mỗi việc 1,25 điểm; Đạt được trọn điểm, Một phần được một nửa, chưa làm được 0; việc 8 tính theo số thành phần đạt của câu insight (con số, nguyên nhân, hàm ý). Chênh lệch nhỏ khi giảng viên chốt lại thang điểm là có thể.

## Tài liệu để xem lại

Slide nằm trong file `Slide_B1_HocVien.pptx` (thư mục Buổi 1). Các file Excel để làm lại thao tác tính từ thư mục Buổi 1.

| Slide | Nội dung | Làm lại trên file |
|---|---|---|
| 27 | Data Validation, tab Error Alert | `Demo/B1_Demo_Table_Validation.xlsx`, sheet Sales |
| 28 | Circle Invalid Data | `Demo/B1_Demo_Table_Validation.xlsx`, sheet Sales |
| 40 | Tổng kết + câu insight (con số, nguyên nhân, hàm ý) | Sheet `Insight` trong file bạn nộp; câu mẫu ở `Homework/B1_BTVN_MCI_Mart_T4_2025_Answer.xlsx`, sheet Insight_mau |

## Đối soát nhanh với đáp án

Đáp án nằm ở `Homework/B1_BTVN_MCI_Mart_T4_2025_Answer.xlsx`, sheet `Kiem_tra`.

| Chỉ tiêu | Đáp án | File của bạn | Khớp? |
|---|---|---|---|
| Số dòng dữ liệu | 1.919 | 1.919 | ✔ |
| Tổng doanh thu | 296.965.000 | 296.965.000 | ✔ |
| Tổng Quantity | 3.448 | 3.448 | ✔ |
| Tổng Pct_Share | 100% | 100% | ✔ |
| Ô Quantity / Order_Date còn là text | 0 / 0 | 0 / 0 | ✔ |
| Thông báo lỗi tiếng Việt ở tab Error Alert | Có, cả 3 cột | Chưa có | ✘ |
| Sheet Insight | 1 câu đủ 3 phần | Trống | ✘ |

## Làm tốt

- **Phần dữ liệu làm rất chắc.** Bảng sạch 1.919 dòng; đổi đủ 10 ô Quantity và 12 ô ngày về đúng số và đúng ngày (không đảo ngày/tháng); sửa đúng 7 ô sai mã và số âm mà không xóa dòng nào. Mọi chỉ tiêu đối soát ở bảng trên đều khớp.
- **Công thức đúng kỹ thuật.** `Kiểm tra`!C5 là công thức tổng doanh thu, cột Pct_Share dùng `$C$5` nên kéo xuống không bị lệch.
- **Biết tự đặt câu hỏi.** Bạn hỏi về Validation và Circle Invalid Data trong sheet `Question`; đây là thói quen học tốt (phần trả lời ở cuối).

![Hình 1 · Bảng sạch và hai ô đối soát](Images/NX01_bang_sach_doi_soat.png)
> **[CHỤP] NX01_bang_sach_doi_soat.png** · File: `Homework/Bai_nop/b1_btvn_nguyen-thi-bich-tramBZHK3N_NguyenThiBichTram.xlsx` · Sheet: `Kiểm tra` · Vùng: A1:E13 · Cần thấy: ô C5 (công thức tổng doanh thu) và các dòng đối soát cho thấy 296.965.000 / 1.919 / 3.448 / 100%; bấm vào C5 để thanh công thức hiện `=SUM(...)`.

## Điểm cần lưu ý

### 1. Sheet `Insight` đang trống (Việc 8: 0/1,25)

**Bạn đã làm:** sheet `Insight` chưa có chữ nào.
**Cần có:** một câu gồm ba phần theo mẫu ở **slide 40** (`Slide_B1_HocVien.pptx`): *"File case có ___ dòng, tôi tìm thấy ___ loại lỗi dữ liệu (con số). Lỗi nguy hiểm nhất là ___ (nguyên nhân), nên nếu không sửa thì ___ (hàm ý)."*
**Câu mẫu của giảng viên** (sheet Insight_mau trong file đáp án): *"Báo cáo tháng 4/2025 có 1.919 dòng giao dịch và tôi tìm thấy 10 loại lỗi. Lỗi nguy hiểm nhất là 13 dòng cộng ("Cộng STxx" và "TỔNG CỘNG") nằm lẫn trong dữ liệu, nên nếu không sửa thì SUM cột Quantity ra 10.309 thay vì 3.448, gần gấp ba, và mọi báo cáo theo cửa hàng đều sai theo."*
**Cách làm nhanh:** bạn đã sửa đủ lỗi nên chỉ cần chọn một lỗi, lấy con số **trước** và **sau** khi sửa (ví dụ tổng Quantity khi còn ô text so với 3.448), rồi nối bằng "nên nếu không sửa thì …".

![Hình 2 · Sheet Insight của bạn](Images/NX02a_insight_trong.png)
> **[CHỤP] NX02a_insight_trong.png** · File: `Homework/Bai_nop/b1_btvn_nguyen-thi-bich-tramBZHK3N_NguyenThiBichTram.xlsx` · Sheet: `Insight` · Vùng: A1:E8 · Cần thấy: sheet trống.

![Hình 2 · Insight mẫu để so sánh](Images/NX02b_insight_mau.png)
> **[CHỤP] NX02b_insight_mau.png** · File: `Homework/B1_BTVN_MCI_Mart_T4_2025_Answer.xlsx` · Sheet: `Insight_mau` · Vùng: A1:A5 · Cần thấy: câu mẫu có đủ con số, nguyên nhân, hàm ý (bật Wrap Text, kéo rộng cột A cho đọc được cả câu).

### 2. Data Validation chưa có thông báo lỗi tiếng Việt (Việc 6: 0,63/1,25)

**Bạn đã làm:** ở cột Quantity bạn nhập "Số lượng quantity nhập số nguyên lớn hơn 0." nhưng nội dung này nằm ở tab **Input Message**, tức là lời nhắc hiện **khi bấm vào ô**. Store_Code và Product_Code chưa có thông báo ở tab nào.
**Đề yêu cầu:** thông báo hiện **khi nhập sai**, nằm ở tab **Error Alert** (đề bài, việc 6: "Viết thông báo lỗi tiếng Việt").
**Khác nhau thế nào:** Input Message là lời nhắc *trước* khi nhập; Error Alert là thông báo chặn *sau* khi nhập sai. Người dùng nhập "abc" vào Quantity chỉ thấy cửa sổ lỗi tiếng Anh mặc định của Excel.

**Cách sửa (làm cho cả ba cột):**
1. Chọn cột (ví dụ Quantity) → **Data → Data Validation**.
2. Sang tab **Error Alert**, tick **Show error alert after invalid data is entered**.
3. Style chọn **Stop**; Title ghi "Số lượng không hợp lệ"; Error message ghi "Chỉ nhập số nguyên lớn hơn 0".
4. Lặp lại cho Store_Code (List từ sheet Store, thông báo: "Chỉ nhập mã trong danh sách cửa hàng ST01–ST12") và Product_Code (List từ sheet Product).
5. Thử nhập sai vào một ô để xem cửa sổ lỗi hiện đúng chữ bạn viết.

**Học lại:** **slide 27** (`Slide_B1_HocVien.pptx`), làm trên `Demo/B1_Demo_Table_Validation.xlsx`, sheet Sales. TIP trong slide: tab Error Alert cho phép tự viết thông báo lỗi bằng tiếng Việt.

![Hình 3 · Tab bạn đã điền (Input Message)](Images/NX03a_dv_input_message.png)
> **[CHỤP] NX03a_dv_input_message.png** · File: `Homework/Bai_nop/b1_btvn_nguyen-thi-bich-tramBZHK3N_NguyenThiBichTram.xlsx` · Sheet: `Sales` · Mở: chọn cột Quantity → Data → Data Validation → tab **Input Message** · Cần thấy: hộp thoại có chữ "Số lượng quantity nhập số nguyên lớn hơn 0." và tên tab đang chọn.

![Hình 3 · Tab Error Alert của bạn (đang trống)](Images/NX03b_dv_error_alert_trong.png)
> **[CHỤP] NX03b_dv_error_alert_trong.png** · Cùng file, cột Quantity, tab **Error Alert** · Cần thấy: ô "Show error alert after invalid data is entered" chưa tick / Title và Error message để trống.

![Hình 3 · Đáp án: tab Error Alert đã điền](Images/NX03c_dv_error_alert_dap_an.png)
> **[CHỤP] NX03c_dv_error_alert_dap_an.png** · File: `Homework/B1_BTVN_MCI_Mart_T4_2025_Answer.xlsx` · Sheet: `Sales` · Mở: cột Quantity → Data Validation → tab **Error Alert** · Cần thấy: ô đã tick, Style Stop, có Title và Error message tiếng Việt.

## Câu insight

Sheet `Insight` đang trống nên chưa có câu để nhận xét. Một câu insight tốt có đủ ba phần:
- **Con số cụ thể**, ví dụ số dòng 1.919 hoặc con số trước và sau khi sửa một lỗi.
- **Nguyên nhân khả dĩ**: lỗi nào gây ra chênh lệch.
- **Hàm ý "nếu không sửa thì …"**: con số nào sẽ sai và sai bao nhiêu.

Gợi ý khung để bạn viết ngay: *"Báo cáo tháng 4/2025 có 1.919 dòng, tôi tìm thấy ___ loại lỗi. Lỗi nguy hiểm nhất là ___ nên nếu không sửa thì tổng ___ ra ___ thay vì ___."*

## Trả lời hai câu hỏi trong sheet `Question`

**1. "Chọn lại được vùng đã validated không?"** Được.
- Muốn **chọn** mọi ô đang có validation: **Home → Find & Select → Data Validation**.
- Muốn **đổi quy tắc** cho cả vùng: bấm vào một ô, mở Data Validation, đổi quy tắc và tick **Apply these changes to all other cells with the same settings**.

![Hình 4 · Find & Select → Data Validation](Images/NX04_find_select_data_validation.png)
> **[CHỤP] NX04_find_select_data_validation.png** · File: `Homework/B1_BTVN_MCI_Mart_T4_2025_Answer.xlsx` · Sheet: `Sales` · Cần thấy: menu Home → Find & Select đang mở, mục Data Validation được trỏ; các ô Quantity/Store_Code/Product_Code đã được chọn (nền xám) ở phía sau.

**2. "Filter được Circle Invalid Data không?"** Không. Vòng đỏ chỉ là hình hiển thị tạm, không phải định dạng của ô nên Filter không nhận ra. Muốn lọc các ô sai:
- Thêm cột kiểm tra, ví dụ `=COUNTIF(Store!$A$2:$A$13,[@Store_Code])=0`, rồi lọc **TRUE**; hoặc
- Tô màu bằng Conditional Formatting rồi **Filter by Color**.

Hướng dẫn thao tác Circle Invalid Data gốc: **slide 28**, `Demo/B1_Demo_Table_Validation.xlsx`, sheet Sales.

## Việc cần làm trước Buổi 2

- [ ] Viết câu insight vào sheet `Insight` theo khung slide 40.
- [ ] Điền tab **Error Alert** cho Quantity, Store_Code, Product_Code rồi thử nhập sai để kiểm tra.
- [ ] Mở `Practice/B1_MCI_Mart_Q1_2025_Answer.xlsx` (file khởi đầu Buổi 2) để làm quen sheet Kiem_tra.
