# MCI Excel BI — Design System

> Bản Markdown của design system MCI Excel BI. Nguồn chuẩn: https://claude.ai/artifact/Es8iAxiYzhSV6q8D7isomX (xuất ngày 25/09/2026).
> File theme Office đi kèm: `MCI_Excel_BI_Colors.xml`, `MCI_Excel_BI_Fonts.xml` (xem `README.md` trong cùng thư mục).

## Tổng quan

Hệ thống thiết kế cho khóa **Excel for Business Intelligence Analyst** của MCI – Magic Code Institute. Áp dụng cho bốn loại sản phẩm: **slide** (.pptx 16:9), **handout** (.docx A4), **file Excel** thực hành và **dashboard** học viên dựng ở B6–B7.

Hệ thống kế thừa deck S1 cũ của khóa học: font Arial, navy `#203864`, tên công cụ đỏ `#C00000`, nhãn PRO TIP vàng `#FFC000`, xanh nhấn `#3398DA`. Phần mới được thêm vào gồm bảng màu dữ liệu đã kiểm tra an toàn mù màu, màu trạng thái, quy ước Excel và theme tối.

### Nguyên tắc

1. **Mỗi màu một vai.** Một màu chỉ làm đúng một việc. Học viên học cách đọc màu, nên màu không được đổi nghĩa giữa các buổi.
2. **Hệ thống vuông như lưới bảng tính.** Mặc định `radius-0`. Chỉ KPI card và callout được bo `radius-md`.
3. **Dữ liệu mới là nhân vật chính.** Chrome (rail, header, khung) dùng navy và xám; màu bão hòa dành cho dữ liệu và trạng thái.
4. **Không bao giờ chỉ dùng màu để truyền nghĩa.** Trạng thái luôn kèm icon và chữ; chuỗi dữ liệu luôn có legend hoặc nhãn trực tiếp.

### Nội dung và giọng văn

- **Tiếng Việt, xưng "mình – các bạn"** trong slide và lời giảng; handout dùng "bạn".
- **Tên giao diện và tên hàm Excel giữ tiếng Anh**, nghĩa tiếng Việt để trong ngoặc khi cần: "Special… (Đặc biệt…)".
- **Tên công cụ** viết đúng như trên Ribbon: "Trace Precedents", không viết "trace precedents".
- **Tiêu đề biểu đồ là câu kết luận** khi có thể: "Tháng 1 vượt target 21% nhờ Tết đến sớm", thay vì "Doanh thu theo tháng".
- **Số theo kiểu Việt Nam:** dấu chấm ngăn nghìn, dấu phẩy thập phân (941.454.000; 34,4%). Đơn vị ghi ở tiêu đề cột: "Doanh thu (triệu VNĐ)".
- Không dùng emoji. Ký hiệu trạng thái dùng glyph ✔ ▲ ✘.

### Màu

#### Chrome và chữ

| Vai trò | Token | Quy tắc |
|---|---|---|
| Nền | `surface`, `surface-alt` | `surface-alt` cho dải xen kẽ bảng và nền callout. |
| Chữ | `ink` → `ink-secondary` → `ink-muted` | Ba cấp, cấp nào cũng đạt ≥4.5:1 trên cả hai nền. **Không dùng lại `#7E7E7E` của deck cũ cho chữ nhỏ** (chỉ 4,1:1); thay bằng `ink-muted`. |
| Thương hiệu | `navy` | Tiêu đề slide, heading, nền header bảng, số KPI. Chữ trên nền navy dùng `on-navy`. |
| Rail | `rail` | Thanh dọc trái của slide nội dung và nền slide chia phần. |
| Tên công cụ | `tool-red` | **Chỉ dành cho tên công cụ/hàm.** Không dùng cho biểu đồ, lỗi, cảnh báo. Đây là "chữ ký" của phong cách khóa học. |
| PRO TIP | `protip-gold` | Là **nền** của nhãn PRO TIP, chữ trên nó dùng `on-gold`. Khi không có nhãn nền, viết "PRO TIP:" bằng `protip-text`. **Không đặt chữ vàng trên nền trắng.** |
| Nhấn phụ | `accent` | Mũi tên và số chú thích trên ảnh chụp màn hình, viền ô đang chọn. Không dùng cho chữ nhỏ. |
| Liên kết, focus | `link`, `focus-ring` | Focus là vòng đặc 2px `focus-ring`, cách mép 2px. |

#### Dữ liệu (biểu đồ, dashboard)

Bảng màu categorical 6 slot đã chạy bộ kiểm tra màu (dataviz validator) cho cả hai theme: không màu nào ra ngoài dải độ sáng quy định, đủ độ bão hòa, cặp liền kề tách biệt với người mù màu (ΔE ≥ 8,4) và với người nhìn bình thường (ΔE ≥ 19,3).

- **Thứ tự slot cố định, không xoay vòng. Màu đi theo thực thể, không theo thứ hạng.** Lọc bớt chuỗi thì các chuỗi còn lại giữ nguyên màu.

| Slot | Token | Miền | Ngành hàng |
|---|---|---|---|
| 1 | `series-1` | Bắc | Đồ uống |
| 2 | `series-2` | Trung | Sữa |
| 3 | `series-3` | Nam | Thực phẩm khô |
| 4 | `series-4` | — | Bánh kẹo |
| 5 | `series-5` | — | Chăm sóc cá nhân |
| 6 | `series-6` | — | Gia dụng |

- **Biểu đồ phân tán, bubble và bản đồ** (mọi cặp màu đều đứng cạnh nhau) tối đa **3 chuỗi**: slot 1–3 đã kiểm tra mọi cặp. Nhiều hơn thì gộp thành "Khác" hoặc tách thành nhiều biểu đồ nhỏ.
- **Theme sáng:** `series-3`, `series-4`, `series-5` dưới 3:1 so với nền, nên **bắt buộc có nhãn trực tiếp hoặc bảng số** kèm biểu đồ.
- **Mốc so sánh** (Target, trung bình, kỳ trước) luôn là `series-reference`, xám nét đứt, không bao giờ dùng màu series.
- **Thang tuần tự:** một sắc xanh `seq-100` → `seq-300` → `seq-500` (color scale 3 màu của Excel).
- **Thang phân kỳ:** `div-negative` – `div-midpoint` – `div-positive`, dùng cho % so target và YoY. Điểm giữa luôn là xám.

#### Trạng thái (RAG)

| Mức | Điều kiện | Mark / icon | Ô Excel (conditional formatting) |
|---|---|---|---|
| Đạt | ≥ 100% target | `status-good` + ✔ | `cell-good-fill` + `cell-good-text` (style *Good* có sẵn của Excel) |
| Cảnh báo | 90% – < 100% | `status-warning` + ▲ | `cell-warn-fill` + `cell-warn-text` (style *Neutral*) |
| Không đạt | < 90% | `status-critical` + ✘ | `cell-bad-fill` + `cell-bad-text` (style *Bad*) |

- Ngưỡng 100% (Target) và 90% (Warning) lấy từ KPI dictionary của khóa; mức Critical 80% chỉ ghi trong KPI dictionary, dashboard không có màu riêng cho mức này.
- Màu trạng thái **dành riêng**, không bao giờ làm màu series.
- `status-critical` khác `tool-red`: một bên là số liệu xấu, một bên là tên công cụ.

#### Quy ước Excel

- **Ô nhập tay hoặc giả định:** chữ `excel-input-text` trên nền `excel-input-fill`. Ô công thức giữ màu `ink`.
- **Header bảng:** nền `navy`, chữ `on-navy` đậm, style `cell`.
- **Kiểm tra / đối soát:** cột kết quả là ✔ / ✘ trên nền `cell-good-fill` / `cell-bad-fill`.
- **Table style:** *TableStyleMedium2* (theo accent 1 của theme). Tên bảng có tiền tố `tbl`: `tblSales`, `tblStore`.

### Chữ

Toàn hệ thống dùng một font: **Arial**. Font này có sẵn trên mọi máy Office, nên slide, handout và Excel hiển thị giống nhau ở mọi nơi. Phân cấp tạo bằng cỡ chữ và độ đậm, không đổi font.

- **Slide:** `slide-title` 40pt · `slide-section` 32pt · `tool-name` 24pt · `slide-body` 20pt · `slide-tip` 18pt · `slide-footnote` 12pt.
- **Handout:** `doc-h1` 16 · `doc-h2` 13 · `doc-h3` 11,5 (tên công cụ, `tool-red`) · `doc-body` 10,5 · `doc-caption` 9.
- **Excel / dashboard:** `kpi-value` 28 · `sheet-title` 14 · `chart-title` 12 · `cell` 10 · `axis-label` 9.
- Số trong bảng căn phải và dùng chữ số đều cột (tabular-nums). Số lớn đứng riêng (KPI) dùng chữ số thường.

### Bố cục

- **Slide 20 × 11,25 in** (`slide-width`, `slide-height`). Slide nội dung có `rail` rộng `rail-width` bên trái, tên phần xoay dọc; lề `slide-margin`.
- **Slide công cụ** theo thứ tự tên (`tool-name`) → mục đích (`slide-body`) → ảnh minh họa → TIP/PRO TIP. Tối đa `tools-per-slide` = 3 công cụ, ngăn bằng hairline dọc.
- **Ảnh chụp màn hình** không bo góc, viền `hairline` 1px. Chú thích trên ảnh dùng số tròn `accent` ① ② ③.
- **Dashboard:** KPI card ở hàng đầu, trend ở giữa, chi tiết ở dưới; tối đa `visuals-per-dashboard` = 8 visual. KPI quan trọng nhất đặt góc trên trái.
- **Khoảng cách:** `space-1` … `space-8` (4–32px).

### Mẫu slide

Mỗi slide của khóa chọn một trong các mẫu dưới đây (nhóm *Slide templates*). Nội dung mẫu lấy từ kịch bản B1 và dataset MCI Mart.

| Loại slide | Component | Dùng khi |
|---|---|---|
| BÌA | `SlideTitle` | Slide 1 mỗi buổi. |
| CHIA PHẦN | `SlideSection` | Mở một phần mới; toàn slide nền `rail`. |
| CÂU HỎI | `SlideQuestion` | Câu hỏi kinh doanh mở buổi. |
| KHÁI NIỆM – quy trình | `SlideProcess` | Nội dung là một chuỗi bước có thứ tự. |
| KHÁI NIỆM – lộ trình | `SlideBrickMap` | Đầu buổi: buổi hiện tại nằm ở đâu trong 7 viên gạch. |
| CÔNG CỤ | `SlideTool` | 1–3 công cụ cạnh nhau, mỗi công cụ có ảnh minh họa. |
| SO SÁNH | `SlideCompare` | Bảng so sánh phương án theo tiêu chí. |
| CÔNG THỨC | `SlideFormula` | Giải phẫu cú pháp một hàm với chú thích ①–④. |
| ĐÚNG / SAI | `SlideGoodBad` | Đặt cách làm đúng và sai cạnh nhau. |
| THỰC HÀNH | `SlidePractice` | Giao bài: các bước, thời gian, đáp án tự kiểm. |
| QUIZ | `SlideQuiz` | Câu hỏi A–D; trình chiếu thành 2 slide (câu hỏi → đáp án). |
| TỔNG KẾT | `SlideSummary` | Điểm chính + khung viết insight. |
| DASHBOARD | `SlideDashboard` | Trình bày dashboard ở B6–B7: tiêu đề insight, KPI, 1 biểu đồ. |

Mọi slide nội dung dùng khung chung trong `components/bundle.css` (lớp `.s`, `.s-rail`, `.s-main`, `.s-title`, `.s-tool`…), vẽ ở tỷ lệ 960 × 540 px cho slide 20 × 11,25 in (1pt ≈ 0,667px).

### Biểu tượng và hình ảnh

- Không dùng bộ icon riêng. Icon Ribbon của Office xuất hiện qua ảnh chụp màn hình; trạng thái dùng glyph ✔ ▲ ✘.
- **Ảnh minh họa luôn chụp trên dữ liệu case MCI Mart** (Excel 365 Windows, zoom 100%), để ảnh khớp đúng file học viên đang làm.
- Chưa có logo MCI dạng file. Nơi cần logo, viết chữ "MCI" đậm màu `navy`.

### Áp dụng vào Office

Theme màu Office tương ứng: **Accent 1–6 = `series-1` … `series-6`**, Dark 2 = `navy`, Light 2 = `surface-alt`, Hyperlink = `link`, font Headings / Body = Arial. Áp theme này thì biểu đồ Excel và PowerPoint tự nhận đúng màu dữ liệu.

## Tokens

### Màu

| Token | Light | Dark | Cách dùng |
|---|---|---|---|
| `surface` | `#FFFFFF` | `#151A24` | Nền trang: slide, handout, sheet Excel, dashboard. Mọi màu chữ bên dưới đều đạt ≥4.5:1 trên nền này. |
| `surface-alt` | `#F3F5F9` | `#1E2532` | Nền phụ: dải xen kẽ của bảng, nền callout, nền ô nhóm. Chữ ink, ink-secondary, navy vẫn đạt ≥4.5:1. |
| `rail` | `#203864` | `#2A4576` | Thanh dọc bên trái slide nội dung (rộng rail-width) và nền slide chia phần. Chữ trên nó dùng on-navy. |
| `hairline` | `#D9DEE8` | `#2F3746` | Đường kẻ bảng, viền khung ảnh minh họa, vạch ngăn 3 công cụ trên cùng một slide. |
| `ink` | `#1B2230` | `#EEF1F6` | Chữ chính (thân bài, số liệu) trên surface và surface-alt. 15.9:1 / 15.4:1. |
| `ink-secondary` | `#585858` | `#B9BDC6` | Chữ phụ: mô tả dưới tên công cụ, chú thích bảng. Giữ màu xám của deck S1 cũ. 7.1:1 trên surface. |
| `ink-muted` | `#6B6F76` | `#9AA0AA` | Caption, footnote, nguồn ảnh [PDF pX]. Thay cho #7E7E7E của deck S1 (chỉ 4.1:1, không đạt với chữ nhỏ). |
| `navy` | `#203864` | `#9DB4E0` | Màu thương hiệu MCI: tiêu đề slide, heading handout, header bảng Excel (nền), KPI value. Trên dark theme chuyển sang tông sáng để đọc được. |
| `on-navy` | `#FFFFFF` | `#FFFFFF` | Chữ trên nền navy hoặc rail (11.6:1). |
| `tool-red` | `#C00000` | `#FF7A7A` | CHỈ dùng cho TÊN CÔNG CỤ / HÀM trên slide và handout (24pt đậm). Không dùng cho biểu đồ, không dùng để báo lỗi. 6.5:1 trên surface. |
| `protip-gold` | `#FFC000` | `#FFC000` | Nền nhãn PRO TIP và điểm nhấn vàng. Là màu NỀN, không dùng làm màu chữ trên nền trắng (1.6:1). Chữ trên nó dùng on-gold. |
| `on-gold` | `#1B2230` | `#1B2230` | Chữ trên nền protip-gold (9.7:1). |
| `protip-text` | `#8A6500` | `#FFC000` | Khi cần viết chữ 'PRO TIP:' mà không có nhãn nền (handout, ghi chú): 5.3:1 trên surface. |
| `accent` | `#3398DA` | `#4FA9E6` | Màu nhấn phụ của deck S1: mũi tên chú thích, số thứ tự chú thích ảnh, đường viền ô đang chọn. Không dùng cho chữ nhỏ trên nền sáng (3.2:1). |
| `link` | `#1F6FB2` | `#7FB8EA` | Liên kết và tên file trong handout, tài liệu (5.3:1). |
| `focus-ring` | `#1F6FB2` | `#7FB8EA` | Vòng focus bàn phím 2px, đặc, trên mọi nền (≥4.8:1). |
| `excel-input-text` | `#0000FF` | `#8AB4FF` | Quy ước Excel: số nhập tay / giả định (chữ xanh dương). Formula giữ màu ink. |
| `excel-input-fill` | `#FFF2CC` | `#3A3320` | Quy ước Excel: ô học viên được sửa (ví dụ ô Target, tham số). |
| `series-1` | `#2A78D6` | `#3987E5` | Dữ liệu slot 1 — Miền Bắc; ngành Đồ uống; hoặc chuỗi duy nhất (Actual). Thứ tự slot cố định, không xoay vòng. |
| `series-2` | `#EB6834` | `#D95926` | Dữ liệu slot 2 — Miền Trung; ngành Sữa. |
| `series-3` | `#1BAF7A` | `#199E70` | Dữ liệu slot 3 — Miền Nam; ngành Thực phẩm khô. Light: 2.8:1 với nền, phải có nhãn trực tiếp hoặc bảng số. |
| `series-4` | `#EDA100` | `#C98500` | Dữ liệu slot 4 — ngành Bánh kẹo. Light: 2.2:1, phải có nhãn trực tiếp. |
| `series-5` | `#E87BA4` | `#D55181` | Dữ liệu slot 5 — ngành Chăm sóc cá nhân. Light: 2.7:1, phải có nhãn trực tiếp. |
| `series-6` | `#4A3AA7` | `#9085E9` | Dữ liệu slot 6 — ngành Gia dụng. |
| `series-reference` | `#52514E` | `#C3C2B7` | Đường/điểm Target, trung bình, kỳ trước: xám trung tính, nét đứt. Không bao giờ dùng màu series cho mốc so sánh. |
| `chart-grid` | `#E5E7EC` | `#2A303C` | Gridline mảnh của biểu đồ. |
| `chart-axis` | `#C3C7CF` | `#3A4150` | Trục và đường baseline. |
| `seq-100` | `#CDE2FB` | `#184F95` | Thang tuần tự (heatmap, color scale): giá trị thấp. |
| `seq-300` | `#6DA7EC` | `#2A78D6` | Thang tuần tự: giữa. |
| `seq-500` | `#256ABF` | `#86B6EF` | Thang tuần tự: cao. Color scale Excel 3 màu = seq-100 → seq-300 → seq-500. |
| `div-negative` | `#E34948` | `#E66767` | Thang phân kỳ: âm (dưới target, giảm YoY). |
| `div-midpoint` | `#F0EFEC` | `#383835` | Thang phân kỳ: điểm giữa (= 0 hoặc = 100% target). Luôn xám. |
| `div-positive` | `#2A78D6` | `#3987E5` | Thang phân kỳ: dương (vượt target, tăng YoY). Xanh–đỏ thay vì xanh lá–đỏ để an toàn mù màu. |
| `status-good` | `#0CA30C` | `#0CA30C` | Trạng thái ĐẠT (≥100% target): chấm/icon ✔. Luôn kèm chữ hoặc icon, không bao giờ chỉ màu. |
| `status-warning` | `#FAB219` | `#FAB219` | Trạng thái CẢNH BÁO (80–<100%): icon ▲. 1.8:1 trên nền sáng nên luôn kèm nhãn. |
| `status-critical` | `#D03B3B` | `#E66767` | Trạng thái NGUY HIỂM (<80%): icon ✘. Khác tool-red; không dùng thay cho nhau. |
| `cell-good-fill` | `#C6EFCE` | `#16391D` | Conditional formatting ô ĐẠT (đúng style 'Good' có sẵn của Excel). Chữ cell-good-text. |
| `cell-good-text` | `#006100` | `#8FE59A` | Chữ trên cell-good-fill (6.1:1). |
| `cell-warn-fill` | `#FFEB9C` | `#3D3312` | Conditional formatting ô CẢNH BÁO (style 'Neutral' của Excel). Chữ cell-warn-text. |
| `cell-warn-text` | `#9C5700` | `#F5C86B` | Chữ trên cell-warn-fill (4.7:1). |
| `cell-bad-fill` | `#FFC7CE` | `#3F1A1E` | Conditional formatting ô KHÔNG ĐẠT (style 'Bad' của Excel). Chữ cell-bad-text. |
| `cell-bad-text` | `#9C0006` | `#FF9AA5` | Chữ trên cell-bad-fill (5.9:1). |

### Chữ

Font: Arial, "Helvetica Neue", Helvetica, sans-serif

**Slide (20 × 11.25 in)**

| Style | Cỡ | Line-height | Độ đậm | Cách dùng |
|---|---|---|---|---|
| `slide-title` | 53.3px | 1.1 | 700 | 40pt — tiêu đề slide, màu navy. |
| `slide-section` | 42.7px | 1.15 | 700 | 32pt — tên phần trên slide CHIA PHẦN, màu on-navy trên rail. |
| `tool-name` | 32px | 1.2 | 700 | 24pt đậm — tên công cụ/hàm, màu tool-red. Dòng dưới là mục đích bằng slide-body. |
| `slide-body` | 26.7px | 1.35 | 400 | 20pt — nội dung, mô tả công cụ, màu ink / ink-secondary. |
| `slide-tip` | 24px | 1.35 | 400 | 18pt — TIP / PRO TIP / chú thích ảnh '(Trong ảnh minh họa: …)'. |
| `slide-footnote` | 16px | 1.3 | 400 | 12pt — nguồn, số trang, ghi chú nhỏ, màu ink-muted. |

**Handout (A4)**

| Style | Cỡ | Line-height | Độ đậm | Cách dùng |
|---|---|---|---|---|
| `doc-h1` | 21.3px | 1.25 | 700 | 16pt — tiêu đề handout, navy. |
| `doc-h2` | 17.3px | 1.3 | 700 | 13pt — mục lớn, navy. |
| `doc-h3` | 15.3px | 1.3 | 700 | 11.5pt — tên công cụ trong handout, tool-red. |
| `doc-body` | 14px | 1.45 | 400 | 10.5pt — thân bài, ink. |
| `doc-caption` | 12px | 1.4 | 400 | 9pt — chú thích bảng, khung ảnh chờ, footer; ink-muted. |

**Excel & Dashboard**

| Style | Cỡ | Line-height | Độ đậm | Cách dùng |
|---|---|---|---|---|
| `kpi-value` | 37.3px | 1.05 | 700 | 28pt — số KPI trên card dashboard, navy; dùng số thường (proportional). |
| `sheet-title` | 18.7px | 1.2 | 700 | 14pt đậm — tiêu đề sheet Excel ở A1, tool-red chỉ cho file giảng viên; file học viên dùng navy. |
| `chart-title` | 16px | 1.25 | 700 | 12pt đậm — tiêu đề biểu đồ, viết thành câu kết luận khi có thể. |
| `cell` | 13.3px | 1.3 | 400 | 10pt — ô dữ liệu và header bảng (header đậm, nền navy, chữ on-navy). Số căn phải, dùng tabular-nums. |
| `axis-label` | 12px | 1.2 | 400 | 9pt — nhãn trục, legend, nhãn trực tiếp trên biểu đồ; ink-secondary. |

### Khoảng cách

| Token | Giá trị | Cách dùng |
|---|---|---|
| `space-1` | 4px | Khoảng cách icon–chữ trong chip trạng thái. |
| `space-2` | 8px | Khoảng giữa dòng mục đích và ảnh minh họa; gap của chip. |
| `space-3` | 12px | Padding ô bảng, padding callout TIP. |
| `space-4` | 16px | Padding KPI card; gap giữa các card. |
| `space-6` | 24px | Khoảng giữa các khối trên slide / dashboard. |
| `space-8` | 32px | Lề trong của khung nội dung slide. |

### Bo góc

| Token | Giá trị | Cách dùng |
|---|---|---|
| `radius-0` | 0 | Mặc định: bảng, ô Excel, khung slide, ảnh minh họa — hệ thống vuông như lưới bảng tính. |
| `radius-sm` | 4px | Đầu cột dữ liệu trong biểu đồ, chip trạng thái, nhãn PRO TIP. |
| `radius-md` | 8px | KPI card và callout — mức bo lớn nhất của hệ thống. |

### Bố cục

| Token | Giá trị | Cách dùng |
|---|---|---|
| `slide-width` | 20in | Khổ slide 16:9 (18288000 EMU). |
| `slide-height` | 11.25in | Khổ slide 16:9 (10287000 EMU). |
| `rail-width` | 1.6in | Thanh rail bên trái mọi slide nội dung, tên phần xoay dọc. |
| `slide-margin` | 0.5in | Lề an toàn trong vùng nội dung, bên phải rail. |
| `tools-per-slide` | 3 | Tối đa 3 công cụ trên một slide, ngăn bằng hairline dọc. |
| `visuals-per-dashboard` | 8 | Tối đa 8 visual trên một trang dashboard (khuyến nghị 5–8). |

## Components

### Callout

Hộp nhấn cho TIP (mẹo dùng ngay) và PRO TIP (mẹo nâng cao cho học viên đã biết), dùng trên slide và trong handout.

**Người dựng cung cấp:** loại callout (TIP hoặc PRO TIP) và một đến hai câu nội dung.

**Quy tắc:**
- **TIP:** nhãn chữ `navy` đậm, nền `surface-alt`.
- **PRO TIP:** nhãn là một chip nền `protip-gold` với chữ `on-gold`. Nội dung dùng `ink`.
- Khi không đặt được chip (ví dụ ghi chú trong Word), viết "PRO TIP:" bằng màu `protip-text`.
- Bo góc `radius-md`, padding `space-3`. **Không** dùng viền trái tô màu.
- Mỗi slide tối đa 1 PRO TIP; PRO TIP là tầng "Nâng cao" của buổi học.

**Không nên:** đặt chữ vàng trên nền trắng (1,6:1, không đọc được).

### ExcelTable

Quy ước trình bày bảng trong file Excel của khóa học: header, dải xen kẽ, ô nhập liệu, ô công thức và cột đối soát.

**Người dựng file cung cấp:** dữ liệu dạng tidy (1 dòng = 1 bản ghi), tên bảng có tiền tố `tbl`, và giá trị đáp án cho cột đối soát.

**Quy tắc:**
- **Header:** nền `navy`, chữ `on-navy` đậm, style `cell` 10pt Arial.
- **Dòng dữ liệu:** xen kẽ `surface` / `surface-alt`. Số căn phải, dấu chấm ngăn nghìn.
- **Ô nhập tay:** chữ `excel-input-text` trên nền `excel-input-fill`. **Ô công thức:** chữ `ink`.
- **Cột đối soát:** ✔ trên `cell-good-fill`, ✘ trên `cell-bad-fill`.
- Không gộp ô trong vùng dữ liệu, không đặt dòng tổng trong bảng (quy tắc tidy data của B1).

### KpiCard

Thẻ KPI ở hàng đầu dashboard: một con số, mốc so sánh và trạng thái.

**Người dựng cung cấp:**
- tên KPI (theo KPI dictionary);
- giá trị lấy từ measure (CUBEVALUE hoặc ô Pivot);
- target hoặc kỳ trước;
- % đạt và trạng thái.

**Quy tắc:**
- Giá trị dùng `kpi-value` màu `navy`, rút gọn đơn vị ("425,9 tr").
- Dòng so sánh dùng `ink-secondary`. Trạng thái dùng StatusChip.
- Bo `radius-md`, nền `surface`, viền `hairline`, padding `space-4`.
- **Mọi KPI card phải có mốc so sánh.** Một con số đứng một mình không phải insight.
- Không tính KPI ngay trong card; số phải đến từ measure (một nguồn sự thật).

### StatusChip

Chip trạng thái RAG (Đạt / Cảnh báo / Không đạt) cho KPI, bảng đối soát và dashboard.

**Người dựng cung cấp:** mức trạng thái, tính từ % đạt target (≥100% Đạt; 90–<100% Cảnh báo; <90% Không đạt), và nhãn chữ.

**Quy tắc:**
- Chấm màu `status-good` / `status-warning` / `status-critical` **luôn đi kèm icon** (✔ ▲ ✘) **và chữ**. Không bao giờ chỉ dùng màu.
- Chữ trong chip dùng `ink`, không dùng màu trạng thái.
- Bo `radius-sm`, padding `space-1` / `space-2`.
- Màu trạng thái không bao giờ dùng làm màu chuỗi dữ liệu.

### TargetBarChart

Biểu đồ cột Actual so với Target trên cùng một trục.

**Người dựng cung cấp:**
- giá trị Actual theo kỳ (từ measure `Total Revenue`);
- Target cùng kỳ (từ bảng `Target`);
- tiêu đề viết dưới dạng câu kết luận.

**Quy tắc:**
- Cột Actual dùng `series-1`, đầu cột bo `radius-sm`, cột bắt đầu từ 0.
- Target là vạch ngang nét đứt màu `series-reference` vắt qua cột, không phải một chuỗi màu.
- Nhãn giá trị đặt trực tiếp trong cột. Legend luôn có (Actual + Target).
- Grid dùng `chart-grid`, trục dùng `chart-axis`, chữ dùng `ink-secondary` với style `axis-label`. Tiêu đề dùng `chart-title`.
- Mỗi cột có tooltip khi rê chuột (trong Excel: data label hoặc tooltip mặc định).

**Không nên:** tô cột theo trạng thái; trạng thái đặt ở KpiCard hoặc StatusChip.

## Mẫu slide

### SlideTitle

Slide bìa của mỗi buổi học (slide 1 trong kịch bản).

**Người dựng cung cấp:** tên khóa học, số buổi và tên buổi.

**Quy tắc:**
- Bên trái: dòng thương hiệu "MCI · Consulting & Analytics" và tên khóa bằng `slide-title` cỡ lớn màu `navy`. "Buổi n" dùng `tool-red`, tên buổi dùng `ink-secondary`.
- Bên phải: khối `rail` với họa tiết lưới ô bảng tính (`seq-100` mờ), trong đó có một ô gộp `tool-red`, một ô `protip-gold`, một ô `accent`, giống họa tiết ở cover của design system.
- Không đặt ảnh chụp màn hình trên slide bìa.
- Viết đúng chính tả "Consulting" (deck S1 cũ ghi sai thành "Consuting").

### SlideSection

Slide CHIA PHẦN, mở một phần mới trong buổi (ví dụ slide 6, 9, 12, 17, 24 của kịch bản B1).

**Người dựng cung cấp:** số thứ tự phần và tên phần. Tên phần viết thành một câu nói rõ học viên sẽ làm được gì.

**Quy tắc:**
- Toàn slide nền `rail`. Dòng kicker "MAGIC CODE INSTITUTE · MCI" và tên phần dùng `on-navy`.
- Số phần rất lớn màu `protip-gold`, kèm một vạch ngắn `tool-red`. Đây là chỗ duy nhất vàng và đỏ được dùng làm trang trí.
- Không đặt nội dung, ảnh hay bullet trên slide chia phần.

### SlideQuestion

Slide đặt câu hỏi kinh doanh mở đầu mỗi buổi (nguyên tắc "BI trước, Excel sau").

**Người dựng cung cấp:** câu hỏi của Giám đốc bán hàng, ảnh chụp dữ liệu liên quan, và một câu hỏi tương tác cho lớp.

**Quy tắc:**
- Câu hỏi dùng chữ `ink` cỡ lớn, dấu ngoặc kép lớn `tool-red`. Người hỏi ghi bằng `ink-secondary`.
- Câu hỏi tương tác nằm trong hộp `surface-alt`, bo `radius-md`.
- Ảnh bên phải chụp trên dữ liệu case MCI Mart. Chưa trả lời câu hỏi trên slide này; đáp án đến ở cuối buổi.

### SlideProcess

Slide khái niệm dạng quy trình có thứ tự: 3 đến 6 bước đặt ngang.

**Người dựng cung cấp:** tên các bước theo đúng thứ tự, một dòng mô tả và công cụ Excel tương ứng cho mỗi bước, và bước cần nhấn mạnh.

**Quy tắc:**
- Chỉ dùng khi nội dung thật sự là một trình tự. Nhãn "BƯỚC n" chỉ xuất hiện vì thứ tự có ý nghĩa.
- Bước được nhấn mạnh có viền 2px và tên màu `tool-red`. Các bước khác viền `hairline`, tên màu `navy`.
- Công cụ tương ứng đặt ở đáy thẻ, trên nền `surface-alt`.
- Tối đa 6 bước. Nhiều hơn thì chia thành 2 slide.

### SlideBrickMap

Bản đồ 7 buổi dạng các viên gạch xếp chồng, B1 ở dưới cùng. Dùng ở đầu mỗi buổi để học viên thấy mình đang ở đâu.

**Người dựng cung cấp:** buổi hiện tại và brick output của từng buổi.

**Quy tắc:**
- Viên gạch của buổi hiện tại nền `navy`, chữ `on-navy`, có dấu ◀ màu `protip-gold`. Các viên khác nền `surface`, viền `hairline`.
- Cột phải ghi sản phẩm hoặc công cụ chính bằng `ink-secondary`.
- Nội dung phải khớp bản đồ Brick trong `Course_Plan.md`.

### SlideTool

Bố cục slide CÔNG CỤ của khóa học: rail bên trái, tiêu đề, và tối đa 3 công cụ đặt cạnh nhau.

**Người dựng slide cung cấp:**
- tên phần (hiện xoay dọc trên rail);
- tiêu đề slide;
- với mỗi công cụ: tên công cụ, một dòng mục đích, một ảnh chụp màn hình trên dữ liệu case, và TIP hoặc PRO TIP nếu có.

**Quy tắc:**
- Tên công cụ dùng `tool-name` màu `tool-red`; mục đích dùng `slide-body` màu `ink-secondary`.
- Ảnh không bo góc, viền `hairline`. Chú thích trong ảnh viết "(Trong ảnh minh họa: …)" bằng `slide-tip`.
- Các công cụ ngăn nhau bằng hairline dọc. **Tối đa 3 công cụ**; nhiều hơn thì tách sang slide mới.
- Rail dùng `rail`, chữ trên rail dùng `on-navy`.

**Không nên:**
- nhồi thêm đoạn văn giải thích dài; lời giải thích thuộc về lời giảng;
- dùng `tool-red` cho bất cứ thứ gì khác ngoài tên công cụ.

### SlideCompare

Slide bảng so sánh: 2 đến 5 phương án theo 3 đến 6 tiêu chí.

**Người dựng cung cấp:** các phương án (cột), các tiêu chí (dòng) và dòng kết luận.

**Quy tắc:**
- Dùng bảng chuẩn ExcelTable: header nền `navy`, dải xen kẽ `surface-alt`.
- Dòng cuối là kết luận, in đậm.
- Ô tốt / xấu dùng `cell-good-text` / `cell-bad-text` **kèm ✔ / ✘**, không chỉ dùng màu.
- Không ghi giá bản quyền hay con số dễ lỗi thời lên slide.

### SlideFormula

Slide giải phẫu một công thức: công thức lớn, các phần được đánh số và chú thích bên dưới.

**Người dựng cung cấp:** công thức mẫu và 2 đến 4 chú thích theo thứ tự đọc.

**Quy tắc:**
- Công thức đặt trong khối `surface-alt`, font monospace; dấu `=` màu `tool-red`, tên hàm màu `navy`.
- Chú thích dùng số tròn `accent` ①–④. Mỗi chú thích tối đa một dòng.
- Công thức viết bằng dấu phẩy như tài liệu gốc; lời giảng nhắc học viên đổi thành `;` nếu máy dùng vùng Việt Nam.

### SlideGoodBad

Slide đặt hai cách làm cạnh nhau: GOOD bên trái, BAD bên phải, mỗi bên một ảnh và 2 đến 4 ý.

**Người dựng cung cấp:** ảnh minh họa cho mỗi bên và các ý so sánh theo từng cặp tương ứng.

**Quy tắc:**
- Tiêu đề mỗi cột có icon ✔ / ✘ trên chấm `status-good` / `status-critical`, chữ `cell-good-text` / `cell-bad-text`. Màu luôn đi kèm icon và chữ.
- Hai ảnh cùng kích thước, cùng tỷ lệ, chụp trên cùng một bộ dữ liệu.
- Nếu có số liệu hậu quả (ví dụ 32.970 thay vì 10.990), đặt nó ở cột BAD.

### SlidePractice

Slide THỰC HÀNH: tên bài, thời gian, các bước làm và đáp án để học viên tự kiểm tra.

**Người dựng cung cấp:** số và tên bài, số phút, 3 đến 6 bước, file dùng, và con số đáp án.

**Quy tắc:**
- Nhãn "THỰC HÀNH n" nền `navy`, chữ `on-navy`. Thời gian dùng `tool-red` đậm.
- Các bước là danh sách đánh số (có thứ tự thật). Thao tác phím và tên bảng in đậm.
- Hộp "Tự kiểm tra" nền `surface-alt`: luôn có ít nhất một con số đáp án cụ thể (style `kpi-value` thu nhỏ, `navy`).
- Ghi tên file ở footnote.

### SlideQuiz

Slide câu hỏi trắc nghiệm A–D, gồm hai trạng thái: chỉ có câu hỏi, và đã hiện đáp án kèm lời giải thích.

**Người dựng cung cấp:** loại câu (kiến thức hay tình huống), chủ đề, câu hỏi, 4 lựa chọn, đáp án đúng và một câu giải thích.

**Quy tắc:**
- Dòng đầu "QUESTION n · CHỦ ĐỀ" màu `navy`; số câu "n / 7" ở góc phải.
- Mỗi buổi khoảng 7 câu, **trong đó 2 câu tình huống** (theo `Course_Plan.md`).
- Đáp án đúng: viền 2px `status-good`, nền `cell-good-fill`, nhãn "✔ Đáp án". Không chỉ đổi màu.
- Khi trình chiếu, tách thành 2 slide (câu hỏi → đáp án) để lớp trả lời trước khi thấy đáp án.

### SlideSummary

Slide tổng kết cuối buổi: các điểm chính cần nhớ bên trái, khung viết câu insight bên phải.

**Người dựng cung cấp:** 3 đến 5 điểm chính, câu mẫu insight có chỗ trống, và một câu ví dụ có số liệu thật.

**Quy tắc:**
- Điểm chính đánh số màu `tool-red`, chữ `ink`. Mỗi điểm tối đa một dòng.
- Khung insight nền `navy`, chữ `on-navy`, nhãn `protip-gold`, bo `radius-md`. Chỗ trống là gạch chân chấm màu vàng.
- Câu ví dụ luôn có **con số** và **hệ quả**; đây là tiêu chuẩn insight của khóa (rubric C6).

### SlideDashboard

Slide trình bày dashboard trong bài thuyết trình B6–B7: tiêu đề là câu kết luận, hàng KPI card, rồi một biểu đồ chính.

**Người dựng cung cấp:**
- câu kết luận làm tiêu đề;
- 2 đến 4 KPI, mỗi KPI có mốc so sánh và trạng thái;
- một biểu đồ trả lời đúng câu kết luận.

**Quy tắc:**
- Tiêu đề là insight ("Tháng 1 vượt target nhờ Tết đến sớm; ST07 cần theo dõi"), không phải tên biểu đồ.
- KPI card và biểu đồ dùng đúng các component KpiCard, StatusChip, TargetBarChart.
- Tối đa 1 biểu đồ chính + 3 KPI trên một slide trình bày. Dashboard đầy đủ để trong file Excel, slide chỉ trích phần kể chuyện.
- Số liệu phải lấy từ measure / KPI dictionary, không gõ tay.

## Phụ lục — khung slide (components/bundle.css)

```css
/* MCI Excel BI — slide frame for template previews.
   Layout scale: a 20 x 11.25 in slide drawn at 960 x 540 px (48 px/in, 1pt = 0.667px). */
.s{position:relative;width:960px;height:540px;background:var(--surface);color:var(--ink);font-family:var(--font-sans);overflow:hidden;border:1px solid var(--hairline);box-sizing:border-box;}
.s *{box-sizing:border-box;}
.s-rail{position:absolute;left:0;top:0;bottom:0;width:77px;background:var(--rail);}
.s-rail .lbl{position:absolute;left:50%;top:50%;transform:translate(-50%,-50%) rotate(-90deg);white-space:nowrap;color:var(--on-navy);font-weight:700;font-size:15px;letter-spacing:.06em;text-transform:uppercase;}
.s-rail .logo{position:absolute;left:0;right:0;bottom:16px;text-align:center;color:var(--on-navy);font-weight:700;font-size:13px;letter-spacing:.08em;}
.s-main{position:absolute;left:101px;right:28px;top:24px;bottom:22px;display:flex;flex-direction:column;gap:14px;}
.s-title{margin:0;font-size:27px;line-height:1.1;font-weight:700;color:var(--navy);text-wrap:balance;}
.s-sub{margin:-6px 0 0;font-size:13px;color:var(--ink-secondary);}
.s-body{font-size:13.3px;line-height:1.4;color:var(--ink-secondary);}
.s-ink{color:var(--ink);}
.s-tool{font-size:16px;line-height:1.2;font-weight:700;color:var(--tool-red);}
.s-tip{font-size:12px;line-height:1.4;color:var(--ink);}
.s-tip b{color:var(--navy);}
.s-pro{display:inline-block;background:var(--protip-gold);color:var(--on-gold);font-weight:700;font-size:10px;letter-spacing:.06em;padding:2px 7px;border-radius:var(--radius-sm);margin-right:6px;vertical-align:1px;}
.s-foot{margin-top:auto;font-size:9px;color:var(--ink-muted);}
.s-shot{position:relative;border:1px solid var(--hairline);background:var(--surface-alt);min-height:80px;display:grid;place-items:center;padding:8px;text-align:center;font-size:10px;line-height:1.35;color:var(--ink-muted);}
.s-n{display:inline-grid;place-items:center;width:18px;height:18px;border-radius:50%;background:var(--accent);color:#fff;font-size:10px;font-weight:700;}
.s-shot .s-n{position:absolute;top:6px;right:6px;}
.s-cols{display:grid;gap:0;flex:1;min-height:0;}
.s-col{display:flex;flex-direction:column;gap:8px;padding:0 16px;border-left:1px solid var(--hairline);min-height:0;}
.s-col:first-child{border-left:0;padding-left:0;}
.s-col:last-child{padding-right:0;}
.s-table{width:100%;border-collapse:collapse;font-size:13px;}
.s-table th{background:var(--navy);color:var(--on-navy);text-align:left;font-weight:700;padding:10px 12px;}
.s-table td{padding:11px 12px;border-bottom:1px solid var(--hairline);color:var(--ink);}
.s-table tr:nth-child(even) td{background:var(--surface-alt);}
.s-table .hl td{font-weight:700;}
.s-kicker{font-size:11px;font-weight:700;letter-spacing:.12em;text-transform:uppercase;color:var(--ink-muted);}
```
