# Kịch bản slide — Buổi 1: Giới thiệu Excel, Nền tảng công thức & Dữ liệu có cấu trúc

> **Phiên bản 3.1:** đồng bộ với `Course_Plan.md` v3.3 (xem các bảng "Thay đổi" ở cuối file). Lời giảng viết thành **văn nói đầy đủ**, số liệu lấy từ dataset thật.
> Căn cứ: `Course_Plan.md` v3.3 (Yêu cầu đầu vào, B1: timetable 150 phút, Hướng dẫn giảng dạy B1, Nguyên tắc 1/9/10, Style spec). Dữ liệu: `2_Dataset/` (xem `2_Dataset/README.md`).

## Cách dùng kịch bản

- **Lời giảng** là lời nói mẫu. Giảng viên không cần đọc nguyên văn; nên nói theo giọng của mình nhưng giữ đủ **ý chính** và **câu hỏi cho lớp** (in đậm).
- Dòng *(→ …)* trong lời giảng là **chỉ dẫn hành động**: bấm, chuyển màn hình, chờ lớp làm.
- Mỗi slide có thời lượng dự kiến. Nếu lớp chậm, cắt ở những chỗ đánh dấu **[có thể rút gọn]**.

### Điều chỉnh cho lớp đã biết Excel sơ bộ

Theo Course Plan: rút Giao diện + Phím tắt từ 30 xuống 15 phút, dồn 15 phút vào Auditing và tidy data. Tổng vẫn 150 phút.

| Slide | Bản chuẩn | Lớp đã biết Excel | Cách rút / cách dùng thêm |
|---|---|---|---|
| 9–11 Giao diện | 15' | 6' | Slide 10 chỉ giữ Name Box và Status Bar; Slide 11 chỉ nêu quy tắc "việc X → tab Y", bỏ câu hỏi nhanh |
| 12–15 Phím tắt | 15' | 9' | Slide 13 giữ Ctrl + ↓ (phát hiện dòng trống); Slide 14 giữ F4; Slide 15 chỉ nêu 3 tổ hợp Alt |
| 21–22 Auditing | 5' | 12' | Demo Evaluate Formula trên một công thức lồng (ví dụ `=IFERROR(B2/$B$5,0)`); cả lớp tự đặt 2–3 ô vào Watch Window |
| 25–26 Tidy data & Table | 9.5' | 17.5' | Lớp tự liệt kê hết lỗi tidy của file Dirty theo nhóm 2 người trước khi giảng viên chữa; demo Table tự mở rộng bằng cách dán thêm dữ liệu |

## Quy ước

**Loại slide** (theo Style spec):

| Loại | Ý nghĩa |
|---|---|
| BÌA | Slide bìa |
| CHIA PHẦN | Slide mở một phần mới |
| KHÁI NIỆM | Slide giải thích khái niệm |
| CÔNG CỤ | Slide dạy công cụ: tên đỏ → mục đích → ảnh → TIP / PRO TIP |
| THỰC HÀNH | Bài làm tại lớp |
| QUIZ | Câu hỏi trắc nghiệm |
| TỔNG KẾT | Chốt kiến thức cuối buổi |

**Nguồn ảnh:**

| Ký hiệu | Ý nghĩa |
|---|---|
| **[CHỤP]** | Tự chụp Excel 365 Windows trên file case |
| **[PDF pX]** | Tham khảo trang X của PDF nguồn |
| **[VẼ]** | Sơ đồ vẽ bằng shape |

**File dùng trong buổi** (trong thư mục buổi này: `Practice/`, `Handout/`, `Demo/` và `Homework/`; bản phát cho học viên có hậu tố `_HocVien`):

| File | Công dụng |
|---|---|
| `Practice/B1_MCI_Mart_Q1_2025_Dirty.xlsx` | File "bẩn" phát cho học viên |
| `Practice/B1_MCI_Mart_Q1_2025_Answer.xlsx` | Đáp án; đồng thời là file khởi đầu B2 |
| `Handout/B1_Checklist_KiemTraMay.docx` | Checklist kiểm tra máy: gửi qua nhóm chat của lớp trước khóa học để học viên tự làm; không đạt thì liên hệ trợ giảng trong lớp |
| `Homework/B1_BTVN_MCI_Mart_T4_2025.xlsx` | Bài tập về nhà bắt buộc (đề), phát cuối buổi, nộp trước Buổi 2 tại hocvien.mcivietnam.com (tài khoản = email đăng ký MCI) |
| `Homework/B1_BTVN_MCI_Mart_T4_2025_Answer.xlsx` | Đáp án bài tập về nhà (sheet Kiem_tra, Dap_an_loi, Insight_mau). Chỉ giảng viên / trợ giảng |
| `Demo/B1_Demo_PctShare.xlsx` | Demo slide 19 (tham chiếu `$`), 21 (Trace), 22 (Evaluate Formula, Watch Window). Chỉ giảng viên |
| `Demo/B1_Demo_Table_Validation.xlsx` | Demo slide 26 (Ctrl + T, Table tự mở rộng), 27 (Data Validation), 28 (Circle Invalid Data). Chỉ giảng viên |
| `Handout/B1_Handout.docx` | Handout tự học |
| `Handout/B1_Infographic_TomTat.png` | Infographic tóm tắt buổi, phát riêng (không chèn vào handout) |

**Số liệu đáp án B1** (dùng khi chữa bài):

| Chỉ tiêu | Giá trị |
|---|---|
| Số dòng dữ liệu sạch | **6.187 dòng** |
| Số đơn hàng | **3.144 đơn** |
| Tổng số lượng | **10.990** |
| Tổng doanh thu Q1/2025 | **941.454.000 VNĐ** |

| Miền | Số dòng | Doanh thu |
|---|---|---|
| Bắc | 2.114 | 323.978.000 |
| Trung | 1.228 | 189.813.000 |
| Nam | 2.845 | 427.663.000 |

**Tổng:** 41 slide, 150 phút.

---

## Phần 0 — Mở đầu (10 phút · slide 1–5)

### Slide 1 · BÌA · 1'
- **Trên slide:** MCI / Consulting & Analytics / **Excel for Business Intelligence Analyst** / Buổi 1
  - Dòng nhỏ cuối slide: *Yêu cầu máy: Excel 365 (hoặc 2016+) trên **Windows**. Dùng Mac? Báo giảng viên ngay hôm nay.*
- **Ảnh:** [VẼ] nền theo Style spec.
- **Ghi chú giảng viên:** yêu cầu này đã được gửi trong email trước khóa học (Course Plan, "Yêu cầu đầu vào"). Nhắc lại ở đây để ai chưa đọc email cũng biết ngay từ phút đầu.

**Lời giảng:**
> Chào cả lớp, chào mừng các bạn đến với khóa Excel for Business Intelligence Analyst của MCI.
>
> Trước khi vào bài, một thông tin quan trọng về máy tính, dù email trước khóa học đã nhắc: **khóa này cần Excel trên Windows.** Bạn nào đang dùng Mac thì nhớ giúp mình điều này; chút nữa ở Phần 1 mình sẽ nói kỹ vì sao và cách chuẩn bị. Nói ngay từ đầu để các bạn có đủ thời gian sắp xếp, chứ không phải đến giữa khóa mới phát hiện.
>
> Và thêm một điều để các bạn khỏi bất ngờ: buổi hôm nay chúng ta **chưa làm phân tích gì cả**. Hôm nay mình học một kỹ năng mà người làm BI dùng mỗi ngày nhưng ít ai dạy, đó là *nhìn vào một file dữ liệu và biết nó có dùng được hay không*. Nghe thì đơn giản, nhưng một nửa số báo cáo sai mình từng gặp là sai từ bước này.

### Slide 2 · KHÁI NIỆM · Câu hỏi mở buổi · 1.5'
- **Trên slide:** *"Công ty có một file bán hàng. Trước khi phân tích, làm sao biết file này dùng được?"* Góc slide có nhãn nhỏ **Tầng 1 — Descriptive**.
- **Ảnh:** [CHỤP] `B1_MCI_Mart_Q1_2025_Dirty.xlsx`, phần đầu file có:
  - tiêu đề báo cáo và dòng "Xuất từ hệ thống POS";
  - tiêu đề 2 tầng tô vàng;
  - cột Vùng bị gộp ô;
  - màu nền theo miền.

**Lời giảng:**
> Đây là file bán hàng Quý 1 năm 2025 của MCI Mart, một chuỗi bán lẻ mà các bạn sẽ đồng hành suốt khóa học. File này được xuất thẳng từ hệ thống bán hàng, đúng kiểu file mà sếp hay chuyển tiếp cho các bạn kèm một dòng: "Em xem giúp anh doanh thu từng miền."
>
> Mình cho các bạn 30 giây. **Nhìn vào file này, bạn thấy có điều gì có thể làm mình tính sai không?**
>
> *(→ chờ 30 giây, gọi 2–3 bạn, ghi ý của các bạn lên bảng, KHÔNG nhận xét đúng sai)*
>
> Cảm ơn các bạn. Mình giữ các ý này trên bảng. Cuối buổi, ở Phần 5, mình sẽ quay lại xem các bạn đã tìm ra được bao nhiêu vấn đề. Hiện file này có **ít nhất 9 loại vấn đề** được cài sẵn.

### Slide 3 · KHÁI NIỆM · Quy trình BI 5 bước · 2.5'
- **Trên slide:** Hỏi → Thu thập → Làm sạch → Phân tích → Trình bày. Dưới mỗi bước là công cụ Excel tương ứng: *Câu hỏi kinh doanh / Get Data / Power Query / Pivot–DAX / Dashboard*.
- **Ảnh:** [VẼ] 5 khối mũi tên. Bước "Hỏi" tô navy đậm, các bước còn lại tô nhạt (đỏ chỉ dành cho tên công cụ).

**Lời giảng:**
> Trước khi vào Excel, mình muốn các bạn nhìn thấy bức tranh lớn. Làm BI, dù bằng Excel, Power BI hay bất kỳ công cụ nào, luôn đi qua 5 bước.
>
> **Bước 1 là Hỏi:** mình cần trả lời câu hỏi kinh doanh gì. **Bước 2 là Thu thập:** lấy dữ liệu từ đâu. **Bước 3 là Làm sạch:** sửa những lỗi như trong file vừa rồi. **Bước 4 là Phân tích:** tính toán, so sánh. **Bước 5 là Trình bày:** biến con số thành thông điệp mà người ra quyết định hiểu được.
>
> Các bạn để ý mình tô đậm bước đầu tiên. Người mới làm BI thường nhảy thẳng vào bước 4: mở file là kéo Pivot ngay. Nhưng nếu câu hỏi sai hoặc mơ hồ thì phân tích đẹp đến đâu cũng vô ích.
>
> Vì vậy khóa này có một quy tắc: **mỗi buổi đều mở bằng một câu hỏi kinh doanh và kết thúc bằng một câu insight các bạn tự viết.** Khóa học không dạy "Excel có những nút gì", mà dạy "làm BI bằng Excel".

### Slide 4 · KHÁI NIỆM · Bản đồ Brick 7 buổi & thang 5 tầng câu hỏi · 3.5'
- **Trên slide:** bảng 3 cột:

| Buổi | Viên gạch (brick output) | Tầng câu hỏi |
|---|---|---|
| B1 | Dữ liệu gốc tidy (Excel Table) | 1 — Chuyện gì đã xảy ra? |
| B2 | Bảng làm giàu + bảng KPI bằng SUMIFS | 1 — Chuyện gì đã xảy ra? |
| B3 | Báo cáo Pivot | 2 — Ở đâu, khi nào, nhóm nào? · 3 — Vì sao? (drill-down) |
| B4 | Pipeline Power Query tự refresh | *Nền móng: giúp các tầng sau đáng tin* |
| B5 | Star schema + KPI dictionary | 3 — Vì sao? |
| B6 | Dashboard có target / RAG | 4 — Nên làm gì? · 5 — Sắp tới thế nào? (nhập môn) |
| B7 | Dự án BI end-to-end + thuyết trình (bộ dữ liệu mới) | Đủ 5 tầng |

- **Ảnh:** [VẼ] 7 viên gạch xếp chồng, B1 ở dưới cùng và tô đậm; bên phải là cầu thang 5 bậc, bậc 1 tô đỏ `C00000`.

**Lời giảng:**
> Khóa học có 7 buổi và mình gọi mỗi buổi là một viên gạch. Viên gạch hôm nay nằm ở dưới cùng: một file dữ liệu tidy, dạng Excel Table. Buổi 2 làm giàu chính bảng đó và tính KPI bằng SUMIFS. Buổi 3 dựng báo cáo Pivot trên cùng bảng dữ liệu. Buổi 4 tự động hóa toàn bộ bằng Power Query, Buổi 5 dựng mô hình dữ liệu, và đến Buổi 6 các bạn có một dashboard hoàn chỉnh. Buổi 7 các bạn tự làm lại toàn bộ trên dữ liệu của một công ty khác, làm cá nhân.
>
> Điều này có nghĩa là **sản phẩm hôm nay là nguyên liệu của buổi sau**. Nếu file hôm nay còn dòng tổng lẫn trong dữ liệu, tổng số lượng ở Buổi 2 sẽ bị cộng gần gấp ba mà các bạn không hề biết.
>
> Cột bên phải là **thang 5 tầng câu hỏi**. Quy trình 5 bước vừa rồi cho biết *làm theo thứ tự nào*; thang này cho biết *câu hỏi sâu đến đâu*. Tầng 1 hỏi "chuyện gì đã xảy ra", tầng cao nhất hỏi "sắp tới sẽ thế nào". Câu hỏi hôm nay ở **Tầng 1**, và mỗi buổi các bạn sẽ leo thêm một bậc.
>
> Một lưu ý để các bạn yên tâm: nếu vì lý do gì đó bạn chưa làm xong, đầu buổi sau lớp sẽ phát **file khởi đầu**, tức là đáp án của buổi trước, để bạn theo tiếp. Không ai bị bỏ lại vì lỡ một viên gạch.

### Slide 5 · KHÁI NIỆM · Giới thiệu case MCI Mart · 1.5'
- **Trên slide:**
  - MCI Mart: **12 cửa hàng, 3 miền** (Bắc 4, Trung 3, Nam 5), **40 sản phẩm, 6 ngành hàng**, **1.500 khách hàng** thành viên.
  - Người dùng báo cáo: **Giám đốc bán hàng**.
  - Dữ liệu: 01/2024 – 04/2025. Các bảng: Sales · Product · Store · Customer · Target.
- **Ảnh:** [VẼ] 5 hộp bảng, chưa nối quan hệ.

**Lời giảng:**
> Giới thiệu nhanh công ty của chúng ta. MCI Mart có 12 cửa hàng trải khắp 3 miền, từ Cầu Giấy, Hoàn Kiếm ngoài Hà Nội, qua Đà Nẵng, Huế, Nha Trang, vào tới Quận 1, Thủ Đức và Cần Thơ. Công ty bán 40 sản phẩm thuộc 6 ngành hàng: đồ uống, sữa, thực phẩm khô, bánh kẹo, chăm sóc cá nhân và gia dụng.
>
> Người dùng báo cáo của chúng ta là **Giám đốc bán hàng**. Mỗi buổi, anh ấy sẽ hỏi một câu và chúng ta phải trả lời.
>
> Đây là dữ liệu giả lập nhưng được làm rất giống thật: có mùa Tết, có cửa hàng mới mở, có cửa hàng gặp sự cố. Những "câu chuyện" này nằm sẵn trong dữ liệu. Mỗi buổi sẽ mở khóa một phần, bắt đầu từ Buổi 2, và nhiệm vụ của các bạn là tự tìm ra và gọi tên chúng trong câu insight cuối buổi.
>
> Năm bảng trên slide hiện đứng riêng lẻ. Đến Buổi 4, các bạn sẽ tự tạo thêm bảng thứ sáu là bảng lịch (Calendar), và đến Buổi 5 thì nối tất cả lại thành một mô hình.

---

## Phần 1 — Phiên bản Excel & kiểm tra máy (10 phút · slide 6–8)

### Slide 6 · CHIA PHẦN · 0.5'
- **Trên slide:** MAGIC CODE INSTITUTE / MCI / **1. Excel dùng cho BI: phiên bản & chuẩn bị máy**

**Lời giảng:**
> Phần đầu tiên rất thực tế: đảm bảo máy của các bạn học được hết khóa này.

### Slide 7 · KHÁI NIỆM · So sánh phiên bản · 6'
- **Trên slide:**

| | Excel 2016+ (Windows) | Microsoft 365 (Windows) | Excel Online | Excel cho Mac |
|---|---|---|---|---|
| Power Query | Có | Có | Hạn chế | Có, thiếu một số connector |
| Power Pivot / Data Model | Chỉ một số bản | Có | **Không** | **Không** |
| Dynamic Arrays (FILTER, XLOOKUP…) | Không (2016/2019) | Có | Có | Có (365) |
| **Học được khóa này?** | Được nếu có Power Pivot | **Khuyến nghị** | Không (từ B4) | Không (từ B4) |

- **TIP:** Bạn dùng Mac? Nên cài máy ảo Windows (Parallels Desktop hoặc VMware Fusion) và báo giảng viên **ngay hôm nay**. **Nên có máy từ Buổi 2**; hạn chót trước Buổi 4.

**Lời giảng:**
> Có nhiều phiên bản Excel, và với người làm BI chúng khác nhau rất nhiều. Mình không so giá, vì giá thay đổi liên tục. Mình chỉ so những thứ ảnh hưởng tới khóa học.
>
> Các bạn nhìn dòng thứ hai: **Power Pivot**, hay Data Model. Đây là công cụ dựng mô hình dữ liệu. Buổi 4 chúng ta bắt đầu load dữ liệu vào Data Model, và từ Buổi 5 đến Buổi 7 mọi thứ xoay quanh nó. Nó chỉ có trên Excel cho Windows, và không phải bản Windows nào cũng có. **Excel trên Mac và Excel Online không có Power Pivot.** Đây là thông tin chính thức từ tài liệu của Microsoft, không phải ý kiến của mình.
>
> *(→ hỏi lớp)* **Ai đang dùng Mac giơ tay giúp mình?** Cảm ơn các bạn. Từ Buổi 4 các bạn bắt buộc phải có Excel chạy trên Windows. Cách mình khuyến nghị là cài phần mềm máy ảo như **Parallels Desktop hoặc VMware Fusion** để chạy Windows và Excel bản Windows ngay trên máy Mac; nếu không cài được thì các bạn chuẩn bị một laptop Windows khác để mang theo. Máy móc là phần các bạn tự chuẩn bị, lớp không có máy cho mượn. Nhưng mình khuyên đừng đợi đến Buổi 4: hãy chuẩn bị xong **trước Buổi 2**. Lý do là giao diện, phím tắt và cả một số menu trên Mac khác Windows. Nếu Buổi 2 và Buổi 3 các bạn đã quen với Windows, thì đến Buổi 4, khi bài khó hơn hẳn, các bạn chỉ phải tập trung vào nội dung chứ không phải vừa học Power Query vừa làm quen máy mới.
>
> Hôm nay vẫn học được trên Mac; handout có bảng phím tắt tương đương để các bạn tạm dùng.
>
> Dòng thứ ba là Dynamic Arrays, gồm các hàm mới như XLOOKUP, FILTER. Excel 2016 và 2019 không có. Nếu đang dùng bản cũ, bạn vẫn học được nhưng sẽ cần cách làm thay thế mà mình chỉ ở Buổi 2.
>
> Tóm lại: **khuyến nghị Microsoft 365 trên Windows.**

### Slide 8 · CÔNG CỤ · Xác nhận kiểm tra máy · 3.5'
Học viên đã tự làm 3 bước theo `B1_Checklist_KiemTraMay` (gửi qua nhóm chat của lớp trước khóa học). Slide này chỉ nhắc lại 3 bước và xác nhận ai chưa đạt. Ba công cụ trên một slide, ngăn bằng đường kẻ dọc:

| # | Tên (đỏ) | Mục đích | Ảnh [CHỤP] | TIP |
|---|---|---|---|---|
| 1 | **Get Data** | Có Power Query không? | Tab Data → Get Data | |
| 2 | **COM Add-ins** | Có Power Pivot không? | File → Options → Add-ins → Manage: COM Add-ins → Go → tick *Microsoft Power Pivot for Excel* | Tick xong, Ribbon có tab **Power Pivot** |
| 3 | **About Excel** | 32-bit hay 64-bit? | File → Account → About Excel | **PRO TIP:** 64-bit xử lý mô hình lớn tốt hơn |

- **Chân slide:** Checklist đã gửi trong nhóm chat trước buổi học. Chưa đạt hoặc chưa làm được: báo trợ giảng ngay.

**Lời giảng:**
> Trước buổi học, các bạn đã nhận checklist kiểm tra máy trong nhóm chat của lớp. Đây là 3 bước trong checklist: **Get Data** để biết máy có Power Query, **COM Add-ins** để bật Power Pivot, và **About Excel** để xem máy là bản 32-bit hay 64-bit.
>
> *(→ hỏi lớp)* **Bạn nào đã làm đủ 3 bước và thấy tab Power Pivot trên Ribbon, giơ tay giúp mình.** Cảm ơn các bạn. **Còn bạn nào chưa làm kịp, hoặc có bước không đạt?** Cứ giơ tay, trợ giảng sẽ tới làm cùng bạn ngay.
>
> *(→ trợ giảng đi vòng; bạn chưa làm thì làm bước 2 cùng trợ giảng. Giảng viên không chờ, chuyển sang Phần 2)*
>
> Bạn nào **không thấy** dòng Power Pivot trong danh sách thì đừng lo. Thường là do phiên bản Office không kèm Power Pivot. Các bạn chuẩn bị phương án máy (máy ảo Windows hoặc một laptop Windows khác) trước Buổi 2; cần hỏi thêm thì nhắn trợ giảng trong nhóm chat của lớp.
>
> Các bạn để ý: kiểm tra máy trước khi học chính là **đối soát**, tức là kiểm tra trước chứ không giả định. Trong BI, giả định "chắc là đúng" là nguồn gốc của rất nhiều sai sót. Đây là thói quen đầu tiên mình muốn các bạn mang theo.

---

## Phần 2 — Giao diện rút gọn (15 phút · slide 9–11)

### Slide 9 · CHIA PHẦN · 0.5'
- **Trên slide:** MCI / **2. Giao diện Excel: những gì người làm BI dùng hằng ngày**

**Lời giảng:**
> Phần này nhiều bạn đã biết rồi. Mình sẽ đi nhanh và chỉ dừng lại ở những chỗ người làm BI dùng nhiều nhất.

### Slide 10 · CÔNG CỤ · Giao diện chính · 7'
- **Trên slide:** ảnh giao diện có 7 chú thích:
  - ① Ribbon
  - ② Name Box
  - ③ Formula Bar
  - ④ Ô / Dòng / Cột
  - ⑤ Sheet tabs
  - ⑥ **Status Bar**
  - ⑦ Workbook = file, Sheet = trang tính
- **Ảnh:** [CHỤP] Excel 365 mở file Answer, chọn cột Revenue; Status Bar hiện Sum = 941.454.000.
- **TIP:** Chọn một cột số, **Status Bar hiện ngay Sum / Average / Count**. Đây là cách đối soát nhanh nhất trong Excel.

**Lời giảng:**
> Mình đi nhanh qua màn hình chính. **Số 1** là Ribbon, dải lệnh phía trên. **Số 2** là Name Box: nó cho biết bạn đang đứng ở ô nào, và bạn có thể gõ địa chỉ vào đó để nhảy tới. Thử gõ A5000 rồi Enter xem. **Số 3** là Formula Bar: ô đang chứa công thức gì sẽ hiện ở đây. **Số 4** là lưới ô, dòng, cột. **Số 5** là các tab sheet ở dưới cùng.
>
> Giờ mình muốn các bạn dừng lại ở **số 6: Status Bar**, dải nhỏ ở đáy màn hình mà ít ai để ý. Các bạn thử bôi đen một cột số bất kỳ. Thấy không? Status Bar hiện ngay tổng, trung bình và số ô được đếm, không cần viết công thức nào.
>
> *(→ demo: mở `Practice/B1_MCI_Mart_Q1_2025_Answer.xlsx`, sheet Sales, chọn cột Revenue; Status Bar hiện Sum 941.454.000)*
>
> Vì sao mình nhấn mạnh chỗ này? Vì từ Buổi 3 trở đi, mỗi lần dựng xong một báo cáo, các bạn sẽ bôi đen dữ liệu gốc, nhìn Status Bar và so với tổng trên báo cáo. Khớp thì yên tâm; lệch là có vấn đề. **Đó là cách đối soát nhanh nhất trong Excel.**
>
> Cuối cùng, **số 7:** Workbook là cả file, còn Sheet là từng trang tính bên trong. Hai từ này sẽ xuất hiện liên tục.

### Slide 11 · KHÁI NIỆM · 8 tab Ribbon: tab nào dùng ở buổi nào · 7.5'
- **Trên slide:**

| Tab | Dùng ở |
|---|---|
| Home | Mọi buổi |
| Insert | B1 (Table), B3 (Pivot), B6 (Chart) |
| Page Layout | B6 |
| Formulas | B1 (Auditing), B2 |
| Data | B1 (Validation), B3 (Refresh All), B4 (Get Data) |
| Review | B6 (Protect Sheet) |
| View | B1 (Freeze Panes), B6 (ẩn gridline) |
| Power Pivot | B5–B6 (bật ở B1, dùng từ B5) |

- **Ảnh:** [CHỤP] dải Ribbon, mỗi tab đóng khung và đánh số.

**Lời giảng:**
> Excel có hàng trăm nút lệnh. Mình sẽ **không** đi qua từng nút, vì tour chi tiết từng tab đã có trong handout, các bạn đọc ở nhà. Hôm nay các bạn chỉ cần nhớ một quy tắc: **"cần làm việc X thì vào tab Y."**
>
> Muốn định dạng, tô màu: tab **Home**. Muốn tạo bảng, Pivot, biểu đồ: tab **Insert**. Muốn viết hàm hay dò lỗi công thức: tab **Formulas**. Muốn lấy dữ liệu, lọc, kiểm soát nhập liệu: tab **Data**. Muốn khóa file trước khi gửi sếp: tab **Review**. Muốn cố định dòng tiêu đề: tab **View**.
>
> Cột bên phải cho biết mỗi tab được dùng nhiều ở buổi nào. Ví dụ tab Data: hôm nay dùng cho Data Validation, đến Buổi 4 thì nó trở thành "cửa chính" của Power Query.
>
> **[có thể rút gọn]** *(→ hỏi nhanh)* **Muốn tìm nút Freeze Panes, bạn vào tab nào?** Đúng rồi, View.

---

## Phần 3 — Phím tắt năng suất (15 phút · slide 12–15)

### Slide 12 · CHIA PHẦN · 0.5'
- **Trên slide:** MCI / **3. Phím tắt: làm việc với dữ liệu lớn không cần chuột**

**Lời giảng:**
> File quý 1 của chúng ta có hơn 6.000 dòng. Cả năm sẽ là hơn 20.000 dòng. Kéo chuột qua từng đó dòng là không khả thi, nên chúng ta cần phím tắt.

### Slide 13 · CÔNG CỤ · Ctrl Shortcuts · 5'

| Tên (đỏ) | Mục đích |
|---|---|
| **Ctrl + Arrow** | Nhảy tới ô cuối vùng dữ liệu theo hướng mũi tên |
| **Ctrl + Shift + Arrow** | Chọn tới ô cuối vùng dữ liệu |
| **Ctrl + Home / End** | Về ô đầu / ô cuối có dữ liệu |

- **Ảnh:** [PDF p11] bố cục; [CHỤP] file Dirty: đứng ở dòng dữ liệu đầu tiên, nhấn Ctrl + ↓ và dừng ngay trước dòng trống giữa khối Miền Trung.
- **TIP:** Ctrl + PgUp / PgDn để chuyển sheet.

**Lời giảng:**
> Phím tắt quan trọng nhất với dữ liệu lớn là **Ctrl + mũi tên**. Các bạn mở file `B1_MCI_Mart_Q1_2025_Dirty.xlsx` *(→ `Practice/B1_MCI_Mart_Q1_2025_Dirty.xlsx`)*, đứng ở ô B6 (dòng dữ liệu đầu tiên) rồi nhấn **Ctrl + mũi tên xuống**.
>
> *(→ chờ lớp làm)*
>
> Bạn nào thấy con trỏ dừng lại **giữa chừng** chứ không xuống tới cuối file? Dừng ở khoảng dòng 2.700 phải không? Nhìn kỹ xem, ngay bên dưới là gì? **Một dòng trống.**
>
> Đây là điều mình muốn các bạn nhớ: Ctrl + mũi tên không chỉ để di chuyển nhanh, nó còn là **công cụ phát hiện lỗi**. Dữ liệu sạch thì Ctrl + ↓ chạy một mạch xuống dòng cuối. Nếu nó dừng giữa chừng, dữ liệu có lỗ hổng. Mình ghi lỗi này lên bảng; nó sẽ được sửa ở phần thực hành.
>
> Thêm **Shift** thành **Ctrl + Shift + ↓** thì Excel bôi đen từ ô hiện tại xuống cuối vùng. **Ctrl + Home** đưa bạn về ô A1, còn **Ctrl + End** đưa tới ô cuối cùng có dữ liệu.

### Slide 14 · CÔNG CỤ · Fn Shortcuts · 5'

| Tên (đỏ) | Mục đích |
|---|---|
| **F2** | Sửa ô đang chọn; tô màu các ô công thức tham chiếu |
| **F4** | Khi đang sửa công thức: đổi A1 → $A$1 → A$1 → $A1. Khi không sửa: lặp lại thao tác cuối |
| **F9** | Tính giá trị phần công thức đang bôi đen |

- **Ảnh:** [PDF p12]; [CHỤP] ô công thức đang F2, vùng tham chiếu có màu; [VẼ] vòng 4 trạng thái của F4.
- **TIP:** Dùng F9 trên Formula Bar xong, nhấn **Esc** để không ghi đè công thức.

**Lời giảng:**
> Ba phím chức năng người làm BI dùng nhiều nhất.
>
> **F2** để sửa ô. Nhưng điều hay hơn là khi bạn F2 vào một ô công thức, Excel **tô màu** các ô mà công thức đang dùng. Đây là cách nhanh nhất để trả lời câu hỏi "công thức này đang lấy số ở đâu?".
>
> **F4** là phím mình sẽ dùng rất nhiều ngay trong Phần 4. Khi đang sửa công thức, mỗi lần nhấn F4 là tham chiếu đổi dạng: A1, rồi đô-la A đô-la 1, rồi A đô-la 1, rồi đô-la A 1, rồi quay về A1. Tạm thời các bạn chỉ cần nhớ phím này; ý nghĩa của dấu đô-la mình giải thích ngay sau giờ nghỉ.
>
> **F9** là mẹo cho người thích "mổ xẻ" công thức: bôi đen một phần công thức trên Formula Bar, nhấn F9, Excel cho bạn biết phần đó bằng bao nhiêu. Nhưng nhớ nhấn **Esc** để thoát; nếu lỡ nhấn Enter, phần đó bị thay bằng con số cứng.

### Slide 15 · CÔNG CỤ · Alt Key Tips & Quick Access Toolbar · 4.5'

| Tên (đỏ) | Mục đích |
|---|---|
| **Alt Key Tips** | Nhấn rồi thả Alt, Ribbon hiện chữ cái; gõ lần lượt để đi sâu vào menu |
| **Quick Access Toolbar** | Ghim lệnh hay dùng lên thanh trên cùng |

- **Ảnh:** [PDF p14]; [CHỤP] Ribbon hiện Key Tips; [CHỤP] QAT có nút Refresh All.
- **PRO TIP:** 3 tổ hợp đáng nhớ nhất:
  - **Alt H V V:** dán giá trị (Paste Values)
  - **Alt A T:** bật/tắt Filter
  - **Alt M V:** Evaluate Formula

**Lời giảng:**
> Các bạn thử nhấn phím **Alt** rồi **thả ra**, đừng giữ. Thấy các chữ cái nhỏ hiện trên Ribbon chưa? Đó là Key Tips: mỗi chữ cái đưa bạn vào một tab, gõ tiếp thì vào sâu hơn.
>
> Không ai thuộc hết. Mình chỉ đề nghị ba tổ hợp. **Alt, H, V, V** để dán giá trị: bạn sẽ dùng nó mỗi khi muốn "đóng băng" kết quả công thức thành số. **Alt, A, T** để bật tắt bộ lọc. **Alt, M, V** để mở Evaluate Formula, công cụ mình giới thiệu ngay sau đây.
>
> Còn **Quick Access Toolbar** là thanh nhỏ trên cùng. Bạn chuột phải vào bất kỳ nút nào và chọn "Add to Quick Access Toolbar" là ghim được. Mình khuyên ghim ngay nút **Refresh All**, vì từ Buổi 3 các bạn sẽ bấm nó liên tục.
>
> Lời khuyên của mình: mỗi buổi chọn **một** phím tắt và ép mình dùng nó đến khi thành phản xạ. Sau 7 buổi là 7 phím tắt, đủ để bạn nhanh hơn hẳn đồng nghiệp.

---

### Slide 16 · NGHỈ · 10'
- **Trên slide:** Nghỉ 10 phút. Trước khi quay lại: mở sẵn file `B1_MCI_Mart_Q1_2025_Dirty.xlsx`.

**Lời giảng:**
> Mình nghỉ 10 phút. Khi quay lại, các bạn mở sẵn file Dirty nhé. Phần tiếp theo là phần quan trọng nhất hôm nay.
>
> *(→ trong giờ nghỉ: trợ giảng hỗ trợ tiếp các bạn chưa bật được Power Pivot)*

---

## Phần 4 — Công thức: cú pháp, tham chiếu, lỗi, kiểm tra (25 phút · slide 17–23)

### Slide 17 · CHIA PHẦN · 0.5'
- **Trên slide:** MCI / **4. Công thức: viết đúng, sao chép đúng, tìm lỗi nhanh**

**Lời giảng:**
> Phần 4 có ba ý: viết công thức đúng cú pháp, sao chép công thức mà không bị sai, và khi có lỗi thì tìm ra nguyên nhân thay vì đoán.

### Slide 18 · KHÁI NIỆM · Formula Syntax · 5'
- **Trên slide:** `= MATCH(lookup_value, lookup_array, [match_type])`, có 4 chú thích:
  - ① luôn bắt đầu bằng `=`
  - ② tên hàm (khoảng 500 hàm, không phân biệt hoa thường)
  - ③ đối số trong ngoặc tròn, ngăn bằng **dấu phẩy hoặc chấm phẩy tùy cài đặt vùng**
  - ④ đối số trong `[ ]` là tùy chọn
- **Ảnh:** [PDF p5]; [CHỤP] ScreenTip khi gõ `=XLOOKUP(`.
- **PRO TIP:** Đọc ScreenTip khi gõ hàm; nó hướng dẫn từng đối số.

**Lời giảng:**
> Mọi công thức trong Excel đều có cùng một khung. Luôn bắt đầu bằng **dấu bằng**. Tiếp theo là **tên hàm**; Excel có khoảng 500 hàm và gõ hoa hay thường đều được. Sau đó là **các đối số** trong ngoặc tròn. Đối số nào nằm trong **ngoặc vuông** là không bắt buộc.
>
> Giờ đến một chi tiết mà người Việt hay vấp. *(→ hỏi lớp)* **Các bạn thử gõ `=SUM(1,2)` rồi Enter xem máy báo gì?**
>
> *(→ chờ)*
>
> Máy một số bạn ra 3, máy một số bạn báo lỗi đúng không? Đó là vì máy cài định dạng vùng Việt Nam dùng **dấu chấm phẩy** để ngăn đối số, còn máy cài vùng Mỹ dùng **dấu phẩy**. Tài liệu tiếng Anh, kể cả slide khóa này, viết theo dấu phẩy. Nếu máy bạn dùng chấm phẩy thì khi chép công thức, bạn đổi dấu. Chỉ vậy thôi, không có gì sai cả.
>
> Mẹo cuối: khi gõ tên hàm và mở ngoặc, Excel hiện một **dòng gợi ý** ngay bên dưới. Đối số đang in đậm là đối số bạn đang nhập. Hãy tập đọc dòng gợi ý này; nó là người hướng dẫn miễn phí tốt nhất.

### Slide 19 · CÔNG CỤ · Reference Types · 9.5'
- **Trên slide:** 4 ô minh họa:
  - Relative `A1`
  - Fixed row `A$1`
  - Fixed column `$A1`
  - Fixed both `$A$1`

  Mũi tên cho thấy tham chiếu thay đổi thế nào khi kéo công thức.
- **Ảnh:** [PDF p6]; [CHỤP] cột **% đóng góp**: phiên bản sai (không có `$`) và phiên bản đúng (có `$`).
- **PRO TIP:** Nắm vững 4 kiểu tham chiếu là mẹo số 1 để làm việc hiệu quả với công thức.

**Lời giảng:**
> Đây là slide quan trọng nhất trước phần thực hành. Mình sẽ không giải thích lý thuyết trước; các bạn làm cùng mình trước đã.
>
> *(→ demo: mở `Demo/B1_Demo_PctShare.xlsx`, sheet Demo_PctShare (các bước ở sheet Huong_dan): 3 dòng doanh thu theo miền Bắc 323.978.000, Trung 189.813.000, Nam 427.663.000, ô tổng 941.454.000 ở B5)*
>
> Câu hỏi của Giám đốc: **"Mỗi miền đóng góp bao nhiêu phần trăm doanh thu?"** Dễ thôi: doanh thu miền chia cho tổng. Ở dòng Bắc, mình gõ bằng ô Bắc chia cho ô tổng, Enter: 34,4%. Giờ mình kéo công thức xuống cho Trung và Nam.
>
> *(→ cột C: gõ =B2/B5 ở C2, kéo xuống C4; Trung và Nam ra #DIV/0!)*
>
> Ơ, Trung ra lỗi chia cho 0. **Tại sao?** Các bạn bấm vào ô đó và nhấn F2.
>
> *(→ chờ lớp nhìn)*
>
> Thấy chưa: tử số đúng là ô Trung, nhưng mẫu số đã **trôi xuống** một ô, trỏ vào ô trống bên dưới ô tổng. Khi kéo công thức, Excel mặc định dịch **mọi** tham chiếu theo. Kiểu tham chiếu mặc định này gọi là **tương đối**. Với tử số thì dịch là đúng, vì mỗi dòng một miền. Nhưng mẫu số phải **đứng yên** ở ô tổng.
>
> Cách sửa: quay lại công thức đầu tiên, đặt con trỏ vào ô tổng trong công thức và nhấn **F4**. Ô tổng giờ có đô-la ở cả cột và dòng, tức là **khóa cứng**. Kéo lại xuống: Bắc 34,4%, Trung 20,2%, Nam 45,4%, cộng lại 100%.
>
> Bốn kiểu trên slide chỉ là bốn cách khóa: **không khóa gì**, **khóa dòng**, **khóa cột**, **khóa cả hai**. Dấu đô-la đứng trước phần nào thì phần đó bị khóa.
>
> Vì sao đây là "mẹo số 1"? Buổi sau các bạn sẽ viết SUMIFS và XLOOKUP: gần như mọi công thức đều có một vùng cần khóa. **Không vững dấu đô-la hôm nay thì Buổi 2 sẽ rất vất vả.**

### Slide 20 · KHÁI NIỆM · Common Errors · 4'
- **Trên slide:**

| Lỗi | Nghĩa | Cách sửa |
|---|---|---|
| `######` | Cột quá hẹp | Kéo hoặc double-click mép cột |
| `#NAME?` | Excel không nhận ra tên trong công thức | Kiểm tra chính tả tên hàm, dấu ngoặc kép |
| `#VALUE!` | Sai kiểu đối số | Kiểm tra phép tính trên text / số dạng text |
| `#DIV/0!` | Chia cho 0 hoặc ô trống | Kiểm tra mẫu số; IF / IFERROR |
| `#REF!` | Tham chiếu tới ô đã bị xóa | Không xóa ô đang được tham chiếu |
| `#N/A` | Không tìm thấy giá trị | Kiểm tra giá trị tra cứu có tồn tại |

**Lời giảng:**
> Sáu lỗi này các bạn sẽ gặp suốt khóa học. Mình đi nhanh.
>
> **Dấu thăng liên tiếp** không phải lỗi thật; cột quá hẹp thôi, double-click vào mép cột là xong. **#NAME?** nghĩa là Excel không hiểu bạn viết gì, thường là gõ sai tên hàm. **#VALUE!** là sai kiểu dữ liệu, ví dụ đem một ô chữ nhân với số. Lỗi này hay gặp khi **số bị lưu dạng chữ**; file Dirty có 15 ô như vậy, lát nữa các bạn sẽ gặp. **#DIV/0!** chính là lỗi vừa rồi. **#REF!** xuất hiện khi bạn xóa một ô đang được công thức khác dùng. **#N/A** là "không tìm thấy", các bạn sẽ gặp nó rất nhiều khi học XLOOKUP ở buổi sau.
>
> Mình muốn các bạn đổi cách nhìn: **lỗi không phải kẻ thù, lỗi là tín hiệu.** Một ô báo lỗi đang cố nói với bạn rằng dữ liệu hoặc logic có vấn đề. Điều đáng sợ hơn nhiều là những con số **sai mà không báo lỗi**, và phần còn lại của khóa học sẽ dạy các bạn bắt chúng.

### Slide 21 · CÔNG CỤ · Auditing Tools (1) · 2.5'

| Tên (đỏ) | Mục đích |
|---|---|
| **Trace Precedents** | Mũi tên từ các ô *ảnh hưởng tới* ô đang chọn |
| **Trace Dependents** | Mũi tên tới các ô *bị ảnh hưởng bởi* ô đang chọn |
| **Show Formulas** | Hiện toàn bộ công thức thay cho kết quả (Ctrl + `) |

- **Ảnh:** [PDF p8]; [CHỤP] trên bảng % đóng góp.
- **TIP:** Ctrl + G → Special → Formulas để chọn mọi ô chứa công thức.

**Lời giảng:**
> *(→ demo trên `Demo/B1_Demo_PctShare.xlsx`, sheet Demo_PctShare: D2 → Trace Precedents; B5 → Trace Dependents)* Khi một file có hàng trăm công thức, làm sao biết ô nào phụ thuộc vào ô nào? Vào tab **Formulas**, nhóm **Formula Auditing**.
>
> Chọn ô % của miền Bắc, bấm **Trace Precedents**: mũi tên xanh chỉ từ các ô "nuôi" nó, là ô doanh thu Bắc và ô tổng. Ngược lại, chọn ô tổng rồi bấm **Trace Dependents**: mũi tên chỉ tới mọi ô phần trăm đang dùng ô tổng, tức là nếu ô tổng sai thì tất cả các ô này sai theo.
>
> **Show Formulas**, hay Ctrl + phím huyền, cho hiện toàn bộ công thức thay vì kết quả. Rất tiện khi nhận một file lạ và muốn biết ô nào là số gõ tay, ô nào là công thức.

### Slide 22 · CÔNG CỤ · Auditing Tools (2) · 2.5'

| Tên (đỏ) | Mục đích |
|---|---|
| **Evaluate Formula** | Chạy công thức từng bước |
| **Error Checking** | Quét sheet tìm lỗi |
| **Watch Window** | Theo dõi giá trị các ô quan trọng khi đang làm ở chỗ khác |

- **Ảnh:** [PDF p9], [PDF p10]; [CHỤP] Watch Window chứa ô Tổng doanh thu.
- **PRO TIP:** Evaluate Formula là công cụ số 1 để hiểu một công thức lạ.

**Lời giảng:**
> *(→ demo trên `Demo/B1_Demo_PctShare.xlsx`, sheet Demo_PctShare: Evaluate Formula ở ô E2; Watch Window thêm ô B5)* Ba công cụ nữa. **Evaluate Formula** cho bạn chạy công thức từng bước một, giống xem phim quay chậm. Với công thức dài lồng nhiều hàm, đây là cách nhanh nhất để biết nó hỏng ở đâu.
>
> **Error Checking** quét cả sheet và dẫn bạn tới từng ô lỗi. Lát nữa các bạn dùng nó để tìm những ô số bị lưu dạng chữ.
>
> Còn **Watch Window**, mình rất thích dùng khi làm BI. Bạn đặt ô **Tổng doanh thu** vào cửa sổ này. Sau đó dù bạn đang sửa dữ liệu ở đâu, cửa sổ vẫn hiện con số tổng. Xóa nhầm một dòng là thấy tổng nhảy ngay. Hôm nay các bạn sẽ đặt ô tổng 941.454.000 vào Watch Window trong lúc thực hành.

### Slide 23 · KHÁI NIỆM · Chốt Phần 4 · 1'
- **Trên slide:**
  - `$` quyết định công thức có tái sử dụng được không.
  - Lỗi là tín hiệu.
  - Auditing giúp tìm nguyên nhân thay vì đoán.

**Lời giảng:**
> Chốt lại Phần 4 bằng ba câu: **Dấu đô-la** quyết định công thức có sao chép được hay không. **Lỗi là tín hiệu**, đừng vội giấu nó đi. Và khi có vấn đề, **dùng Auditing để tìm nguyên nhân** thay vì ngồi đoán.

---

## Phần 5 — Dữ liệu có cấu trúc: Tidy data & Excel Table (10 phút · slide 24–26)

### Slide 24 · CHIA PHẦN · 0.5'
- **Trên slide:** MCI / **5. Dữ liệu "sạch về cấu trúc": nền móng của mọi phân tích**

**Lời giảng:**
> Giờ mình quay lại câu hỏi đầu buổi: làm sao biết file này dùng được?

### Slide 25 · KHÁI NIỆM · Good vs Bad Data Structure · 4.5'
- **Trên slide:**

| GOOD | BAD |
|---|---|
| Hình chữ nhật: biến là cột, bản ghi là dòng | Đảo chiều, biến nằm theo dòng |
| Không định dạng thừa | Ô gộp, tô màu thừa |
| Chỉ gồm dimension & measure | Cột tính toán lẫn vào |
| Tiêu đề cột rõ ràng, 1 dòng | Tiêu đề 2 tầng |
| Không dòng tổng, không tiêu đề phụ | Có dòng tổng, subtotal giữa dữ liệu |

- **Ảnh:** [CHỤP] đặt cạnh nhau: file Dirty và sheet Sales của file Answer. [PDF Pivot p5] bố cục GOOD/BAD.

**Lời giảng:**
> Người làm dữ liệu gọi dữ liệu "sạch về cấu trúc" là **tidy data**. Quy tắc gói gọn trong một câu: **mỗi dòng là một bản ghi, mỗi cột là một trường thông tin.** Với file bán hàng, một dòng là một sản phẩm trong một đơn hàng.
>
> Giờ đối chiếu với những gì các bạn ghi trên bảng lúc đầu buổi. *(→ đọc lại các ý trên bảng, tick từng ý)* Các bạn đã thấy ô gộp ở cột Vùng, tiêu đề hai tầng, màu nền, đúng hết.
>
> Nhưng có một lỗi nguy hiểm nhất mà ít người để ý. *(→ mở `Practice/B1_MCI_Mart_Q1_2025_Dirty.xlsx`, cuộn tới dòng "Tổng Miền Bắc", dòng 2120)* **Dòng này.** File có 3 dòng tổng miền và 1 dòng tổng cộng nằm ngay trong dữ liệu. Nếu các bạn cộng cả cột Quantity, **mỗi sản phẩm bị đếm tới ba lần**: một lần ở dòng chi tiết, một lần ở dòng tổng miền, một lần ở dòng tổng cộng. Tổng số lượng thật là 10.990, nhưng SUM cả cột sẽ ra **32.931**. Và Excel **không báo lỗi gì cả**; nó chỉ lặng lẽ cho bạn một con số sai gần gấp ba.
>
> Vì sao file lại có dạng này? Vì nó được thiết kế để **người đọc**, không phải để **máy tính**. Đẹp để in ra, nhưng không dùng để phân tích được. Việc đầu tiên của người làm BI luôn là chuyển file "để đọc" thành file "để tính".

### Slide 26 · CÔNG CỤ · Excel Table · 5'

| Tên (đỏ) | Mục đích |
|---|---|
| **Ctrl + T (Format as Table)** | Biến vùng tidy thành Table |
| **Table Name** | Đặt tên bảng có nghĩa: `tblSales` |
| **Structured Reference** | Công thức theo tên cột: `=[@Quantity]*[@Unit_Price]`, tự điền cả cột |

- **Ảnh:** [CHỤP] hộp Create Table; Table Design → Table Name; cột Revenue tự điền.
- **PRO TIP:** Table **tự mở rộng** khi thêm dòng, nên công thức, Pivot (B3) và Power Query (B4) tự nhận dữ liệu mới.

**Lời giảng:**
> *(→ demo: mở `Demo/B1_Demo_Table_Validation.xlsx`, sheet Sales; các bước ở sheet Huong_dan)* Khi dữ liệu đã tidy, việc tiếp theo là biến nó thành **Excel Table**. Đứng trong vùng dữ liệu, nhấn **Ctrl + T**, tick "My table has headers", OK. Sau đó vào tab **Table Design** và đặt tên bảng. Chúng ta đặt là **tblSales**: "tbl" để biết đây là bảng, "Sales" để biết nội dung.
>
> Table có hai điểm lợi lớn. Thứ nhất, công thức đọc được bằng tên cột. Thay vì `=I2*J2`, bạn viết `=[@Quantity]*[@Unit_Price]`, đọc là "số lượng của dòng này nhân đơn giá của dòng này". Gõ ở một ô, Excel tự điền cả 6.187 dòng.
>
> Thứ hai, và đây là điều quan trọng nhất: **Table tự mở rộng.** *(→ demo: copy 5 dòng ở sheet Them_dong, dán ngay dưới bảng trong sheet Sales)* Thấy chưa, dòng mới tự thuộc về bảng, cột Revenue tự tính. Đến Buổi 3, Pivot dựng trên Table sẽ tự nhận dữ liệu mới; đến Buổi 4, Power Query cũng vậy.
>
> Đây là bước đầu tiên của thứ mình gọi là **tư duy tự động hóa**: làm một lần, dùng mãi.

---

## Phần 6 — Data Validation (10 phút · slide 27–28)

### Slide 27 · CÔNG CỤ · Data Validation · 6'
- **Tên (đỏ):** **Data Validation**
- **Mục đích:** Giới hạn giá trị được phép nhập: Whole/Decimal, khoảng giá trị, **List**, Date, Text length, Custom.
- **Ảnh:**
  - [PDF p15]
  - [CHỤP] Data → Data Validation: cột `Store_Code`, List = `=Store!$A$2:$A$13`
  - [CHỤP] cột `Quantity`: Whole number > 0
  - [CHỤP] tab Error Alert có câu tiếng Việt
- **TIP:** Tab **Error Alert** cho phép tự viết thông báo lỗi bằng tiếng Việt.

**Lời giảng:**
> Dọn dữ liệu bẩn tốn thời gian. Cách rẻ nhất là **đừng để dữ liệu bẩn lọt vào ngay từ đầu.** Đó là việc của **Data Validation**.
>
> *(→ demo tiếp trên `Demo/B1_Demo_Table_Validation.xlsx`, sheet Sales)* Ví dụ cột Store_Code: MCI Mart chỉ có 12 cửa hàng, từ ST01 đến ST12. Mình chọn cả cột, vào **Data → Data Validation**, phần Allow chọn **List**, phần Source trỏ tới danh sách 12 mã trong sheet Store. Từ giờ ô đó có một mũi tên xổ xuống, và gõ sai mã là Excel chặn lại.
>
> Với Quantity, mình chọn **Whole number**, **greater than 0**: chỉ nhận số nguyên dương.
>
> Mẹo nhỏ: sang tab **Error Alert** và tự viết thông báo, ví dụ "Chỉ nhập mã có trong danh sách cửa hàng ST01–ST12". Người nhập liệu đọc là hiểu phải sửa gì, thay vì một thông báo tiếng Anh khó hiểu.

### Slide 28 · CÔNG CỤ · Circle Invalid Data · 4'
- **Tên (đỏ):** **Circle Invalid Data**
- **Mục đích:** Khoanh đỏ các ô **đã có sẵn** vi phạm quy tắc validation.
- **Ảnh:** [CHỤP] Data → Data Validation ▾ → Circle Invalid Data. File case có **5 ô bị khoanh**: 3 ô `ST3`, 2 ô Quantity âm.
- **PRO TIP:** Validation chỉ chặn dữ liệu *nhập mới*. Với dữ liệu *có sẵn*, dùng Circle Invalid Data.

**Lời giảng:**
> Có một điều ít người biết: Data Validation **chỉ chặn dữ liệu nhập mới.** Những ô đã sai từ trước vẫn nằm im đó. Vậy làm sao tìm chúng?
>
> Bấm vào mũi tên nhỏ cạnh nút Data Validation và chọn **Circle Invalid Data**. *(→ demo trên `Demo/B1_Demo_Table_Validation.xlsx`, sheet Sales, sau khi đã đặt Validation ở slide 27)* Excel khoanh đỏ mọi ô vi phạm quy tắc. Trong file của chúng ta có **5 ô**: ba ô ghi mã cửa hàng là "ST3" thay vì "ST03", và hai ô số lượng bị âm.
>
> **Theo các bạn, "ST3" nên sửa thành gì?** Đúng, ST03. Cả ba dòng đều thuộc Miền Bắc, và ST03 là Hải Phòng, cũng ở Miền Bắc. Còn số lượng âm thì nhiều khả năng là lỗi gõ dấu, nên sửa thành số dương.
>
> Các bạn thấy không: mình vừa làm một bài **đối soát** nhỏ: đặt quy tắc, rồi để Excel chỉ ra chỗ vi phạm. Đây là tinh thần của người làm BI: **không tin số liệu cho đến khi đã kiểm tra.**

---

## Phần 7 — Thực hành (35 phút · slide 29–31)

### Slide 29 · THỰC HÀNH 1 · Làm sạch cấu trúc file case (20 phút)
- **Trên slide:**
  1. Xóa dòng tiêu đề báo cáo và dòng nhóm; tiêu đề chỉ còn 1 dòng; cột A đặt tên `Region`.
  2. Bỏ gộp ô cột Region và điền đầy giá trị (Bắc / Trung / Nam).
  3. Xóa 4 dòng "Tổng…" và dòng trống.
  4. Chuyển Quantity và Order_Date đang ở dạng text thành số và ngày.
  5. Ctrl + T, đặt tên `tblSales`.
- **Kiểm tra:** Ctrl + ↓ chạy thẳng tới dòng cuối. Chọn ô Order_ID đầu tiên → Ctrl + Shift + ↓ (không gồm tiêu đề). Status Bar: Count = **6.187**.

**Lời giảng (giao bài):**
> Giờ đến lượt các bạn. Mở file Dirty và làm theo 5 bước trên slide, thời gian 20 phút.
>
> Mình gợi ý hai chỗ khó. **Bước 2**, điền lại cột Region sau khi bỏ gộp ô: bỏ gộp xong chỉ ô đầu mỗi khối có giá trị, còn lại trống. Các bạn chọn cả cột, nhấn **Ctrl + G → Special → Blanks**, gõ dấu bằng rồi trỏ vào ô ngay trên, sau đó nhấn **Ctrl + Enter**. Cuối cùng copy và dán giá trị (Alt H V V) để bỏ công thức. Nhớ đổi "Miền Bắc" thành "Bắc" cho gọn.
>
> **Bước 4**, số dạng text: ô có tam giác xanh ở góc là dấu hiệu. Dùng **Error Checking → Convert to Number**. Với ngày dạng text, dùng **Text to Columns** với định dạng **DMY**, hoặc hàm DATEVALUE.
>
> Làm xong, các bạn tự kiểm tra: đứng ở dòng đầu, nhấn Ctrl + ↓. Nếu con trỏ chạy một mạch xuống dòng cuối là đạt. Bấm ô Order_ID đầu tiên, nhấn Ctrl + Shift + ↓ để chọn cả cột dữ liệu, không lấy ô tiêu đề, rồi nhìn Status Bar: **Count phải là 6.187.** Nếu bấm chữ cái trên đầu cột thì Excel đếm cả ô tiêu đề và ra 6.188.

**Ghi chú giảng viên:**
- Đi vòng lớp.
- **Ba lỗi hay gặp nhất:**
  - quên dòng "TỔNG CỘNG" ở cuối file (Count ra 6.188 khi đã chọn đúng vùng không gồm tiêu đề; nếu học viên chọn cả cột thì 6.188 là bình thường, hỏi lại cách chọn trước);
  - quên dòng trống giữa khối Miền Trung;
  - để lại chữ "Miền" trong Region.
- Ai làm xong sớm: cho làm trước Thực hành 2.

### Slide 30 · THỰC HÀNH 2 · Công thức & kiểm soát (10 phút)
- **Trên slide:**
  1. Thêm cột `Revenue = [@Quantity]*[@Unit_Price]`.
  2. Tạo ô tổng doanh thu, rồi thêm cột `Pct_Share` (% đóng góp) = Revenue ÷ ô tổng (dùng `$`).
  3. Data Validation cho `Store_Code` (List từ sheet Store) và `Quantity` (> 0).
  4. Circle Invalid Data → sửa 5 ô bị khoanh.
  5. Đặt ô Tổng doanh thu vào Watch Window.
- **Kiểm tra:** Tổng doanh thu = **941.454.000**; tổng Quantity = **10.990**; tổng cột Pct_Share = **100%**; Circle Invalid Data không còn ô nào.

**Lời giảng (giao bài):**
> Bài 2 dùng mọi thứ vừa học: Table, dấu đô-la, Validation, Watch Window. Thời gian 10 phút.
>
> Đáp án để các bạn tự kiểm: tổng doanh thu Quý 1 phải đúng **941 triệu 454 nghìn đồng**. Tổng số lượng phải là **10.990**. Nếu bạn ra một con số lớn hơn nhiều, ví dụ hơn 20 nghìn, thì mình biết ngay bạn còn sót dòng tổng. Nếu doanh thu thấp hơn đáp án khoảng 140 nghìn, nhiều khả năng bạn còn ô số lượng âm chưa sửa.
>
> Đây chính là giá trị của việc biết đáp án trước: **con số lệch sẽ chỉ cho bạn loại lỗi.**

### Slide 31 · KHÁI NIỆM · Chữa bài nhanh · 5'
- **Trên slide:** Checklist brick output B1:
  - [ ] Table `tblSales`, 6.187 dòng
  - [ ] Không ô gộp, không dòng tổng, không dòng trống
  - [ ] Có Validation, không còn ô bị khoanh
  - [ ] Không còn lỗi công thức
  - [ ] Tổng doanh thu = 941.454.000; tổng Pct_Share = 100%
- **Chân slide:** File đáp án B1_MCI_Mart_Q1_2025_Answer.xlsx (sheet Kiem_tra: 8 dòng đối soát) là file khởi đầu Buổi 2.

**Lời giảng:**
> Các bạn tự tick checklist này. *(→ mở `Practice/B1_MCI_Mart_Q1_2025_Answer.xlsx`, sheet Kiem_tra)* Đây là file đáp án. Sheet Kiem_tra có 8 dòng đối soát, mỗi dòng so kết quả công thức với đáp án và hiện dấu tick xanh khi khớp. Các bạn sẽ được nhận file này làm **file khởi đầu Buổi 2**.
>
> *(→ hỏi lớp)* **Bạn nào tick đủ 5 ô?** Tuyệt vời. Bạn nào còn thiếu cũng không sao; buổi sau các bạn bắt đầu từ file đáp án. Nhưng hãy mở file đáp án ra so với file của mình để biết mình sót gì. Đó là cách học nhanh nhất.

**Ghi chú giảng viên (kiểm tra cuối buổi, theo Course Plan):**
- File của học viên đã là Table, không ô gộp, không dòng tổng, có Data Validation.
- Năng lực **D2** (hoài nghi và đối soát) mong đợi ở mức **Developing**: học viên phát hiện được lỗi dữ liệu khi được nhắc.

---

## Phần 8 — Quiz, insight, preview (10 phút · slide 32–41)

Quiz kiểu hỏi nhanh: khoảng 40 giây mỗi câu, cả lớp giơ tay A/B/C/D hoặc dùng Kahoot/Mentimeter. **Câu 6–7 là câu tình huống.**

### Slide 32 · QUIZ · EXCEL QUIZ CHALLENGE · 0.2'

**Lời giảng:**
> Trước khi về, 7 câu hỏi nhanh. Câu 6 và 7 là tình huống thực tế, không có trong slide; các bạn phải suy luận.

### Slide 33 · QUESTION 1: CELL REFERENCES
- Ký hiệu dùng để tạo tham chiếu tuyệt đối? A) # · **B) $** · C) @ · D) !

**Lời giảng:**
> Đáp án B, dấu đô-la. Dễ đúng không?

### Slide 34 · QUESTION 2: CELL REFERENCES
- Tham chiếu nào cố định **dòng** nhưng để cột tương đối? A) $A$1 · B) A1 · **C) A$1** · D) $A1

**Lời giảng:**
> Đáp án C. Nhớ quy tắc: **đô-la đứng trước phần nào thì phần đó bị khóa.** Ở đây đô-la đứng trước số 1, nên khóa dòng.

### Slide 35 · QUESTION 3: SHORTCUTS
- Tổ hợp nào dán giá trị? **A) Alt H V V** · B) Ctrl V · C) Ctrl Alt V · D) Alt V V

**Lời giảng:**
> A. Còn Ctrl + Alt + V mở hộp Paste Special; cũng dán được giá trị, nhưng phải thêm bước chọn.

### Slide 36 · QUESTION 4: AUDITING
- Công cụ nào vẽ mũi tên tới các ô **bị ảnh hưởng** bởi ô đang chọn? A) Error Checking · B) Trace Precedents · **C) Trace Dependents** · D) Evaluate Formula

**Lời giảng:**
> C. Mẹo nhớ: **Dependents** là "những ô phụ thuộc vào tôi", **Precedents** là "những ô đứng trước tôi".

### Slide 37 · QUESTION 5: COMMON ERRORS
- Lỡ xóa một ô đang được công thức tham chiếu, lỗi nào xuất hiện? A) #NAME? · B) #N/A · C) #VALUE! · **D) #REF!**

**Lời giảng:**
> D, #REF!, viết tắt của *reference*: tham chiếu đã mất.

### Slide 38 · QUESTION 6: TÌNH HUỐNG — DATA STRUCTURE
- Sếp gửi file bán hàng có ô gộp theo vùng và dòng "Tổng" sau mỗi vùng, nhờ tính doanh thu từng vùng. Việc **đầu tiên** nên làm là gì?
  - A) Viết SUM cho từng vùng ngay
  - **B) Đưa về dạng tidy rồi chuyển thành Table**
  - C) Tô màu từng vùng
  - D) Chụp màn hình gửi lại sếp

**Lời giảng:**
> Đáp án B. Nhưng mình muốn hỏi thêm: **nếu làm A thì chuyện gì xảy ra?**
>
> *(→ chờ 1 bạn trả lời)*
>
> Đúng rồi: SUM cả cột sẽ cộng luôn dòng tổng, nên doanh thu bị **nhân đôi**, và Excel không báo lỗi gì. Đó là lý do việc đầu tiên luôn là sửa cấu trúc.

### Slide 39 · QUESTION 7: TÌNH HUỐNG — AUTOMATION
- Mỗi tháng bạn dán thêm dữ liệu mới vào cuối bảng và muốn cột Revenue **tự tính**. Nên dùng gì? A) Data Validation · **B) Excel Table** · C) Freeze Panes · D) Show Formulas

**Lời giảng:**
> B. Table tự mở rộng và tự điền công thức cho dòng mới. Đây là hạt giống của tự động hóa; đến Buổi 4 các bạn sẽ thấy nó "nở hoa" với Power Query.

### Slide 40 · TỔNG KẾT + INSIGHT · 2'
- **Trên slide:**
  - **Điểm chính cần nhớ:**
    - (1) Kiểm tra máy trước.
    - (2) `$` cho phép tái sử dụng công thức.
    - (3) Lỗi là tín hiệu.
    - (4) Tidy data + Table là nền móng.
    - (5) Validation chặn lỗi từ gốc.
  - **Insight của bạn (1 câu):** "File case có ___ dòng, tôi tìm thấy ___ loại lỗi dữ liệu *(con số)*. Lỗi nguy hiểm nhất là ___ *(nguyên nhân)*, nên nếu không sửa thì ___ *(hàm ý)*."

**Lời giảng:**
> Tổng kết 5 điều cần nhớ hôm nay *(→ đọc lướt 5 ý trên slide)*.
>
> Giờ là thói quen mà mình muốn các bạn giữ suốt khóa học: **viết một câu insight.** Điền vào câu mẫu trên slide, trong 1 phút.
>
> *(→ chờ 1 phút, gọi 2–3 bạn đọc)*
>
> Một câu mẫu để các bạn tham khảo: "File có 6.187 dòng giao dịch và tôi tìm thấy 9 loại lỗi. Lỗi nguy hiểm nhất là 4 dòng tổng nằm lẫn trong dữ liệu, nên nếu không sửa thì tổng số lượng bán Quý 1 sẽ bị báo gần gấp ba, 32.931 thay vì 10.990, và mọi báo cáo theo vùng đều sai theo."
>
> Các bạn để ý câu này có ba thứ: **con số**, **nguyên nhân** và **hàm ý**, tức là "vậy thì sao". Một insight tốt luôn có đủ cả ba. Hôm nay chúng ta mới ở Tầng 1, "chuyện gì đã xảy ra", nhưng từ Buổi 2 trở đi các câu chuyện trong dữ liệu MCI Mart sẽ lần lượt lộ ra, và câu insight của các bạn sẽ là nơi gọi tên chúng.

### Slide 41 · PREVIEW B2 · 1'
- **Trên slide:**
  - *Buổi sau:* "Doanh thu, số đơn, biên lợi nhuận của từng vùng tháng này là bao nhiêu?"
  - Nguyên liệu: `tblSales` hôm nay + bảng `Product`.
  - Công cụ: KPI Definition Card, Dimension / Measure / Grain, SUMIFS, XLOOKUP.
  - *Về nhà:* đọc handout B1 · ôn dấu `$` · làm **Bài tập về nhà (bắt buộc)**: file `B1_BTVN_MCI_Mart_T4_2025.xlsx` · nộp trước Buổi 2 tại hocvien.mcivietnam.com, tên file `B1_BTVN_<HoTen>.xlsx` · tick "Checklist trước Buổi 2" (mục 11 handout).
  - *Dùng Mac:* báo giảng viên và chốt phương án máy Windows **trước Buổi 2**.
- **Ảnh:** [VẼ] viên gạch B1 (đã xong) → viên gạch B2 (sáng lên).

**Lời giảng:**
> Buổi sau, Giám đốc bán hàng hỏi: "Doanh thu, số đơn và biên lợi nhuận của từng miền tháng này là bao nhiêu?" Vẫn là câu hỏi Tầng 1, nhưng lần này phải định nghĩa KPI cho rõ trước khi tính. Chúng ta sẽ trả lời bằng chính bảng tblSales các bạn vừa làm sạch, cộng thêm bảng sản phẩm.
>
> Về nhà các bạn có mấy việc. Một, **đọc handout Buổi 1**, nhất là phần tour Ribbon mà hôm nay mình đã lướt qua. Hai, **ôn lại dấu đô-la**, buổi sau sẽ cần nó rất nhiều. Ba, làm **bài tập về nhà**, đây là bài **bắt buộc**. File đề là báo cáo bán hàng tháng 4 xuất theo từng cửa hàng, các bạn làm lại đúng quy trình hôm nay: làm sạch, chuyển thành Table, thêm Revenue và Pct_Share, đặt Validation rồi viết một câu insight. Yêu cầu và con số tự kiểm tra có ở mục 10 handout. Các bạn nộp trước Buổi 2 trên hocvien.mcivietnam.com, đăng nhập bằng email đã đăng ký với MCI, đặt tên file là B1_BTVN kèm họ tên. Tiếp theo, tick **Checklist trước Buổi 2** ở mục 11 handout; nếu file của bạn chưa khớp 941.454.000 thì đừng lo, đầu buổi sau lớp phát file khởi đầu.
>
> Và nhắc lại lần cuối: bạn nào dùng Mac, hãy **báo mình và chốt máy Windows trước Buổi 2**, để buổi sau các bạn làm bài trên đúng môi trường của cả khóa.
>
> Cảm ơn cả lớp, hẹn gặp lại Buổi 2!

---

## Phụ lục kịch bản

### Kiểm tra thời lượng

| Phần | Slide | Phút | Timetable B1 |
|---|---|---|---|
| 0 Mở đầu | 1–5 | 10 | 10 |
| 1 Phiên bản & kiểm tra máy | 6–8 | 10 | 10 |
| 2 Giao diện | 9–11 | 15 | 15 |
| 3 Phím tắt | 12–15 | 15 | 15 |
| Nghỉ | 16 | 10 | 10 |
| 4 Công thức, lỗi, Auditing | 17–23 | 25 | 15 + 10 |
| 5 Tidy data & Table | 24–26 | 10 | 10 |
| 6 Data Validation | 27–28 | 10 | 10 |
| 7 Thực hành | 29–31 | 35 | 35 |
| 8 Quiz, insight, preview | 32–41 | 10 | 10 |
| **Tổng** | **41** | **150** | **150** |

Phần 8 cộng từng slide khoảng 8 phút (slide 32: 0,2'; 7 câu quiz × khoảng 40 giây; slide 40: 2'; slide 41: 1'). Khoảng 2 phút còn lại là dự phòng cho câu 6–7 (tình huống) và phần đọc insight.

### Danh sách ảnh cần chụp [CHỤP] (Excel 365 Windows)

| # | Nội dung ảnh | Slide |
|---|---|---|
| 1 | File Dirty: phần đầu và dòng "Tổng Miền Bắc" | 2, 25 |
| 2 | Data → Get Data; COM Add-ins có Power Pivot; About Excel 64-bit | 8 |
| 3 | Giao diện có 7 chú thích; Status Bar Sum = 941.454.000 khi chọn cột Revenue (file Answer) | 10 |
| 4 | Dải Ribbon 8 tab và tab Power Pivot | 11 |
| 5 | Ctrl + ↓ trên file Dirty dừng trước dòng trống ở khối Miền Trung | 13 |
| 6 | F2 tô màu vùng tham chiếu | 14 |
| 7 | Alt Key Tips; QAT có Refresh All | 15 |
| 8 | ScreenTip khi gõ `=XLOOKUP(` | 18 |
| 9 | Bảng % đóng góp 3 miền: sai (không `$`, ra #DIV/0!) và đúng (34,4% / 20,2% / 45,4%) | 19 |
| 10 | Trace Precedents / Dependents trên bảng % đóng góp | 21 |
| 11 | Watch Window chứa ô Tổng doanh thu | 22 |
| 12 | Dirty và sheet Sales của Answer đặt cạnh nhau | 25 |
| 13 | Hộp Create Table; Table Name `tblSales`; cột Revenue tự điền | 26 |
| 14 | Validation List cho Store_Code; Whole number > 0 cho Quantity; Error Alert tiếng Việt | 27 |
| 15 | Circle Invalid Data khoanh 5 ô | 28 |

### Tài nguyên đã có (thư mục `Practice/` và `Handout/` của buổi này)
- `B1_MCI_Mart_Q1_2025_Dirty.xlsx`: 6.187 dòng dữ liệu, cài sẵn 9 loại lỗi + định dạng thừa (sheet `Dap_an_loi` trong file Answer liệt kê chi tiết).
- `B1_MCI_Mart_Q1_2025_Answer.xlsx`: gồm các sheet `Sales` (tblSales + Revenue + Pct_Share + Validation), `Store`, `Kiem_tra` (8 dòng đối soát, tất cả ✔) và `Dap_an_loi`. Đây cũng là **file khởi đầu B2**.
- `Handout/B1_Checklist_KiemTraMay.docx` / `.pdf` (gửi học viên trước Buổi 1)
- `Handout/B1_Handout.docx` / `.pdf` (cập nhật 2026-09-27: chèn đủ 7 ảnh; mục 10 là bài tập về nhà bắt buộc)
- `Handout/B1_Infographic_TomTat.png` (2026-09-27): infographic tóm tắt buổi, phát riêng (không chèn vào handout); sửa nội dung ở `2_Dataset/_generator/infographic_b1.html` rồi chạy `build_infographic.py B1`

### Còn cần làm trước khi dạy
- ~~Chụp 15 ảnh [CHỤP] ở bảng trên, thay vào slide và handout.~~ Đã xong: deck có đủ ảnh, handout đã chèn 7 ảnh (H01–H07).
- ~~Dựng file `.pptx` B1 từ kịch bản này theo Style spec.~~ Đã có: `Slide_B1.pptx` (41 slide, lời giảng nằm trong speaker notes) và bản học viên `Slide_B1_HocVien.pptx` (không notes, không đáp án quiz).

### Thay đổi so với phiên bản 2 (2026-09-24, đồng bộ Course Plan v3.2)

| Slide | Thay đổi | Căn cứ trong Course Plan |
|---|---|---|
| Đầu file | Thêm bảng "Điều chỉnh cho lớp đã biết Excel sơ bộ" | Hướng dẫn giảng dạy B1 (v3.1) |
| 2 | Gắn nhãn Tầng 1 — Descriptive cho câu hỏi mở buổi | Nguyên tắc 9 |
| 4 | Nhãn brick theo Bản đồ Brick; thêm cột thang 5 tầng; sửa "Pivot dựng trên bảng KPI" thành "trên cùng bảng dữ liệu"; B7 là dự án end-to-end, làm cá nhân. 3' → 3.5' | Bản đồ Brick, Nguyên tắc 9, Buổi 7 |
| 5 | Câu chuyện dữ liệu mở khóa từ B2; nhắc bảng Calendar tạo ở B4. 2' → 1.5' | Nguyên tắc 5, 10; Phụ lục F |
| 1, 3 | Nhắc yêu cầu Windows ngay phút đầu (dòng nhỏ trên bìa + 1 câu lời giảng). Slide 1: 0.5' → 1'; Slide 3: 3' → 2.5' | Yêu cầu đầu vào (v3.2): thông báo sớm |
| 7 | Mac / Online: "Không (từ B4)"; Data Model bắt đầu từ B4, trọng tâm B5–B7; học viên Mac nên có máy Windows **từ B2**, hạn chót trước B4 | Yêu cầu đầu vào (v3.2) |
| 8 | Học viên chưa đạt được liên hệ ngay sau buổi, có phương án trước B2 | Yêu cầu đầu vào (v3.2) |
| 11 | Tab Data thêm B3 (Refresh All); Power Pivot B5–B6 | B3, B5, B6 |
| 31 | Ghi chú kiểm tra cuối buổi: danh sách máy chưa đạt, D2 mức Developing | Kiểm tra cuối buổi B1 |
| 40 | Insight đủ 3 thành phần: con số, nguyên nhân, hàm ý | Nguyên tắc 1 |
| 41 | Preview thêm Dimension / Measure / Grain; bài về nhà khớp handout (Checklist trước B2) | Mục tiêu B2; Handout B1 |

Tổng thời lượng và số slide không đổi (41 slide, 150 phút).

### Thay đổi 2026-09-27 (đồng bộ Course Plan v3.3)

| Slide | Thay đổi | Căn cứ trong Course Plan |
|---|---|---|
| 8 | Học viên tự kiểm tra máy trước buổi theo checklist gửi qua nhóm chat; slide chỉ còn xác nhận, trợ giảng ghi phiếu. 8.5' → 3.5' | Yêu cầu đầu vào (v3.3) |
| 29 | Thực hành 1: 15 → 20 phút (nhận 5 phút từ slide 8) | Timetable B1 (v3.3) |
| 41 | "Bài tự luyện A/B" thay bằng **Bài tập về nhà bắt buộc** (`Homework/B1_BTVN_MCI_Mart_T4_2025.xlsx`), nộp trước Buổi 2 | Bài tập về nhà (v3.3) |
| 3, 4 | Slide 3: bước "Hỏi" tô navy đậm (đỏ chỉ dành cho tên công cụ). Slide 4: Buổi 7 làm cá nhân | Style spec; Hình thức B7 (v3.3) |

Tổng vẫn 41 slide, 150 phút.

### Thay đổi 2026-09-28

| Slide | Thay đổi | Căn cứ trong Course Plan |
|---|---|---|
| 7, 8, 16, 31 | Bỏ phiếu ghi nhận kiểm tra máy và danh sách gửi sau buổi. Học viên tự chuẩn bị máy (máy ảo hoặc laptop Windows khác); lớp không có máy cho mượn. Không dùng chữ "ban tổ chức" | Yêu cầu đầu vào (v3.3) |
