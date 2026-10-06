**Điểm: 0/10**

# Nhận xét bài tập về nhà Buổi 2 - Lê Lâm Tuyên

File: `b2-le-lam-tuyen.xlsx`

Điểm hiện tại phản ánh phần việc có bằng chứng trên đúng dữ liệu T4. Bạn kiểm tra lại file đã gửi để xác nhận có nhầm với bài thực hành Q1 không.

| # | Việc | Kết quả | Trọng số | % đạt / Thành phần |
|---|---|---|---:|---|
| 1 | Tra cứu XLOOKUP | Chưa có phần việc đáp ứng bài T4 | 2 | 0% |
| 2 | Cột tính toán | Chưa có phần việc đáp ứng bài T4 | 2 | 0% |
| 3 | Bảng KPI 1: Miền x Ngành hàng | Chưa có phần việc đáp ứng bài T4 | 1 | 0% |
| 4 | Bảng KPI 2: theo Miền | Chưa có phần việc đáp ứng bài T4 | 1 | 0% |
| 5 | KPI Definition Card (AOV) | Chưa có phần việc đáp ứng bài T4 | 1 | 0% |
| 6 | Insight | 0/3 thành phần | 3 | 0/3 |
| | **Tổng đã làm tròn** | | **10** | **0/10** |

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
| Số dòng dữ liệu | 1.919 | Chưa có chỉ tiêu T4 | Chưa |
| Số ô #N/A (Region, Category, Unit_Cost) | 0 | Chưa có chỉ tiêu T4 | Chưa |
| Tổng doanh thu (VNĐ) | 296.965.000 | Chưa có chỉ tiêu T4 | Chưa |
| Tổng bảng KPI 1 (Miền × Ngành) | 296.965.000 | Chưa có chỉ tiêu T4 | Chưa |
| Tổng số đơn | 966 | Chưa có chỉ tiêu T4 | Chưa |
| Tổng COGS (VNĐ) | 196.595.000 | Chưa có chỉ tiêu T4 | Chưa |
| Gross Margin % tháng 4 | 33,8% | Chưa có chỉ tiêu T4 | Chưa |
| AOV tháng 4 (VNĐ) | 307.417 | Chưa có chỉ tiêu T4 | Chưa |

Số tiền trong bảng làm tròn đến VNĐ để dễ đọc; phần đối soát dùng giá trị đầy đủ đến sáu chữ số thập phân.

## Làm tốt

- Sales có công thức Category, Unit_Cost, COGS và Gross_Margin; đã thực hành cách làm giàu bảng.
- KPI_SUMIFS!E20 dùng `=E6/E13`, đúng cách tổng chia tổng cho AOV của Quý 1. Những phần này được ghi nhận là thao tác đã làm, chưa thay cho bằng chứng của bài T4.

## Điểm cần lưu ý

### 1. File đang là Quý 1, khác dữ liệu bài tháng 4

**Bạn đã làm:** Sales có 6.187 dòng, tổng Revenue 941.454.000 VNĐ. KPI_SUMIFS!B2:D2 ghi 1, 2, 3 và E2 ghi Quý 1.

**Đúng cần là:** Bài yêu cầu tháng 4/2025, 1.919 dòng, doanh thu 296.965.000 VNĐ và 966 đơn.

**Hậu quả:** Không có khóa Order_ID + Line_No nào chung với dữ liệu tháng 4 nên chưa có bằng chứng để chấm các phần số liệu của bài T4. Đây là một lỗi gốc khác kỳ dữ liệu; không tách các chênh lệch thành nhiều lỗi độc lập.

**Cách sửa:** Bạn kiểm tra có gửi nhầm file không. Nếu đây là file thực hành trên lớp, gửi đúng bài đã làm từ `Homework/B2_BTVN_MCI_Mart_T4_2025.xlsx`. Với dữ liệu T4, dựng Miền x sáu ngành hàng và bảng KPI miền bằng COUNTIFS/tổng chia tổng.

**Học lại:** Mục bài tập về nhà trong De_bai và slide 16-17, `Demo/B2_Demo_SUMIFS.xlsx`, sheet Demo_SUMIFS.

![Hình 1 · File đang là Quý 1, khác dữ liệu bài tháng 4](Images/NX01_sai_ky.png)
> **[CHỤP] NX01_sai_ky.png** · File: `Homework/Cham_bai/LeLamTuyen/b2-le-lam-tuyen.xlsx` · Sheet: `KPI_SUMIFS` · Vùng: A1:E6 · Cần thấy: tiêu đề Quý 1, tháng 1-3 và tổng 941.454.000. · Chụp Excel 365 Windows, zoom 100%, chỉ thao tác trên bản sao.

### 2. Chưa có card AOV và insight của bài tháng 4

**Bạn đã làm:** KPI CARD!A2:D3 có ba KPI Q1; A6:A17 có nhận xét và đề xuất cho Q1. File chưa có sheet Insight.

**Đúng cần là:** Bài T4 cần card AOV đủ 10 trường và insight gồm con số có so sánh, nguyên nhân khả dĩ, hàm ý.

**Hậu quả:** Phần chữ hiện tại chưa trả lời yêu cầu của dữ liệu T4, nên chưa ghi nhận điểm ở việc 5 và 6.

**Cách sửa:** Sau khi xác nhận đúng file, tạo KPI_Card và Insight. Điền card theo mẫu, viết câu dựa trên bảng T4; giữ số so sánh, thêm nguyên nhân khả dĩ và một hành động cụ thể.

**Học lại:** Slide 4-5 và 40, `Handout/B2_KPI_Definition_Card.docx`.

![Hình 2 · Chưa có card AOV và insight của bài tháng 4](Images/NX02_card_q1.png)
> **[CHỤP] NX02_card_q1.png** · File: `Homework/Cham_bai/LeLamTuyen/b2-le-lam-tuyen.xlsx` · Sheet: `KPI CARD` · Vùng: A1:D17 · Cần thấy: ba KPI Q1, phần chữ A6:A17 và chưa có card AOV đủ 10 trường. · Chụp Excel 365 Windows, zoom 100%, chỉ thao tác trên bản sao.

## Câu insight

File chưa có sheet Insight. Phần chữ hiện tại ở KPI CARD!A6:A7:

> Q1 ghi nhận tổng doanh thu 941,45 triệu VNĐ.
> Miền Nam đang dẫn đầu với 427,66 triệu VNĐ.
> Miền Bắc 323,98 triệu VNĐ.
> Miền Trung 189,81 triệu VNĐ.
> Đề xuất sử dụng kết quả Q1 làm baseline để xây dựng KPI Quý 2, 
> đồng thời có cơ chế ghi nhận/khen thưởng đối với quản lý miền đạt hoặc vượt KPI nhằm tạo động lực và thúc đẩy tăng trưởng đồng đều giữa các khu vực.

| Thành phần | Nhận xét |
|---|---|
| Con số có so sánh | Chưa có con số cho bài T4. KPI CARD!A6 chỉ dẫn doanh thu Q1 và số theo miền. |
| Nguyên nhân khả dĩ | Chưa có nguyên nhân khả dĩ cho so sánh ở bài T4. |
| Hàm ý | KPI CARD!A7/A11:A17 có đề xuất cho Q1, nhưng chưa gắn với số liệu T4 nên chưa ghi nhận trong phần insight của bài này. |

**Mức D4:** Emerging. Phần insight đạt 0/3 thành phần theo bài T4.

## Việc cần làm trước Buổi 3

- [ ] File đang là Quý 1, khác dữ liệu bài tháng 4.
- [ ] Chưa có card AOV và insight của bài tháng 4.

Ảnh minh họa hiện là khung chờ chụp, đã ghi rõ file, sheet và vùng.
