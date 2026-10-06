# Quy tắc chung: tell bề mặt

Nguồn: chắt lọc từ `viet-dieuchinh-vanphong/quy-tac-chung.md` (MIT) và typography T1-T6 của `vietnamizer` (MIT).

Áp cho mọi sắc thái, trừ vùng không đụng ở SKILL.md.

## Ký tự: cấm và thay

| Ký tự | Thay bằng |
|---|---|
| `—` `–` (gạch dài) | gạch thường `-` cách hai bên, dấu hai chấm, ngoặc đơn, hoặc tách câu |
| `…` | `...` (ba dấu chấm gõ tay) |
| ngoặc kép cong `“ ”`, `‘ ’` | ngoặc thẳng `"..."`, `'...'` (T3) |
| `•`, `×`, `™`, `®` | `-`, `x` hoặc bỏ |
| emoji, ký hiệu trang trí (tích, tên lửa, tia lửa) | bỏ hẳn (T6) |
| `→` trong văn xuôi | "rồi", "thành", "dẫn đến" |

Ngoại lệ của dự án: `→` trong `(→ demo …)`, `·` trong chân slide và bảng, `✔` trong kết quả `Kiem_tra` (ví dụ "11/11 ✔"). Giữ nguyên.

## Số theo chuẩn Việt

- Ngăn nghìn bằng chấm: `1.200.000`, không `1,200,000`.
- Thập phân bằng phẩy: `3,5`, không `3.5`. (Số trong công thức Excel và trích từ ô giữ nguyên theo giao diện Excel của máy.)
- Phần trăm dính số: `30%`, không `30 %`.
- Tiền: "2 triệu", "300k", hoặc `2.000.000đ`.

## Định dạng

- Bôi đậm hiếm, chỉ vài chỗ thật cần trong cả bài.
- Không bẻ mọi thứ thành gạch đầu dòng; nhân quả viết thành câu liền.
- Bỏ khuôn lặp "**Nhãn đậm:** mô tả" ở mọi gạch đầu dòng (T5).
- Tiêu đề chỉ hoa chữ đầu, không Title Case từng từ (T1).
- Không rải đường kẻ ngang giữa các mục.
- Dấu phẩy trước "và" trong liệt kê, đứng giữa hai vế ngắn: bỏ (T4).
- Chấm phẩy `;` dùng dè.

## Cụm cấm

- Mở bài sáo: "Trong bối cảnh ... không ngừng phát triển", "Trong thời đại số ngày nay", "Không thể phủ nhận rằng", "Dưới đây là...", "Hãy cùng tìm hiểu...".
- Kết bài máy móc: "Tóm lại", "Kết luận là", "Nhìn chung", "Có thể thấy rằng", "Hy vọng ... hữu ích", "Chúc bạn thành công".
- Chuyển đoạn cơ khí: "Hơn nữa", "Bên cạnh đó", "Thêm vào đó", "Đáng chú ý là", "Không những thế".
- Khuôn lặp: "không chỉ ... mà còn", "tưởng chừng ... nhưng".
- Từ rỗng: "đóng vai trò quan trọng", "mang lại nhiều lợi ích", "tạo giá trị", "góp phần nâng cao", "giải pháp toàn diện", "tối ưu hóa" khi không cụ thể.
- Trạng từ thừa: "vô cùng", "hết sức", "đáng kể", "một cách hiệu quả", "vượt trội", "mạnh mẽ" khi sáo.

## Regex quét (đọc-only)

Dùng cho `Grep` hoặc `grep -nP`:

| Tell | Mẫu |
|---|---|
| gạch dài | `[—–]` |
| chấm lửng ký tự | `…` |
| ngoặc cong | `[“”‘’]` |
| số kiểu Mỹ | `\b\d{1,3}(,\d{3})+\b` |
| phần trăm có cách | `\d\s%` |
| emoji phổ biến | `[\x{1F300}-\x{1FAFF}\x{2600}-\x{27BF}]` (loại `✔` U+2714 ở kết quả `Kiem_tra`) |
| từ nối dịch ở đầu câu | `(^\|[.!?]\s)(Do đó\|Vì vậy\|Bên cạnh đó\|Hơn nữa\|Chính vì thế)` |
| khung đối lập giả | `không phải[^.]{0,40}mà (là\|còn)\|không chỉ[^.]{0,40}mà còn` |
| câu trợ lý còn sót | `Chắc chắn rồi\|Tất nhiên rồi\|Hy vọng (điều này\|bài viết)` |
| từ bị người dùng cấm | `tay nghề` |

Mọi kết quả khớp phải xét theo vùng không đụng trước khi báo (tên file, công thức, ký hiệu quy ước).
