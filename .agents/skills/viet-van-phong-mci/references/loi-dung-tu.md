# Lỗi dùng từ và cấu trúc

Nguồn: chọn lọc từ V1-V25 và T1-T6 của `vietnamizer` v0.9.8 (MIT). Chỉ giữ những lỗi hay gặp trong văn bản hướng dẫn, nhận xét và báo cáo của khóa. Bỏ các luật riêng cho blog cá nhân và học thuật.

Nguyên tắc từ vietnamizer: **gọi tên lỗi trước, chỉ sửa chỗ có lỗi, giữ nguyên ý, dữ kiện và giọng người viết.** Câu đã tự nhiên thì để yên.

| Mã | Lỗi | Ví dụ | Sửa |
|---|---|---|---|
| V1 | Thiếu bổ ngữ kết quả sau động từ | "không tìm công thức sai" | "không tìm ra công thức sai"; "không giải quyết được vấn đề" |
| V2 | Cặp liên từ thiếu vế | "Vì ô trống, kết quả sai" | "Vì ô trống nên kết quả sai" |
| V3 | Thiếu "là" khi định nghĩa | "cách nhanh nhất dùng Ctrl+T" | "cách nhanh nhất là dùng Ctrl+T" |
| V4 | Dấu thời/trạng thái ("đã", "đang") thiếu hoặc lặp | "đã đang làm" | chọn một |
| V5 | Sai loại từ | "ba con cột" | "ba cột" |
| V7 | Thừa "của" trong cụm danh từ | "bảng của dữ liệu của tháng 4" | "bảng dữ liệu tháng 4" |
| V8 | "các/những" lẫn lộn | "những các ô lỗi" | "các ô lỗi" |
| V9 | Giới từ kéo dài | "trong quá trình thực hiện" | "khi làm" |
| V10 | Trạng ngữ đặt mơ hồ | "chỉ sửa ô A1 và B1" (chỉ sửa, hay chỉ A1 và B1?) | đặt "chỉ" sát vế nó nhấn |
| V11 | Câu mở rỗng | "Điều quan trọng cần lưu ý là" | vào thẳng ý |
| V12 | Đầu câu lặp một từ nối | "Sau đó... Sau đó... Sau đó..." | xen "rồi", "xong", bỏ bớt |
| V13 | Cùng khuôn chủ-vị-tân qua nhiều câu | "Bạn đã... Bạn đã... Bạn đã..." | đổi chủ ngữ hoặc gộp |
| V14 | Danh từ đứng một mình thành câu | "Một lỗi phổ biến." | "Đây là lỗi phổ biến." hoặc nối vào câu trước |
| V15 | Danh từ hóa + động từ yếu | "thực hiện việc kiểm tra" | "kiểm tra" |
| V16 | Bị động "bởi" | "ô được tính bởi công thức" | "công thức tính ô" |
| V17 | Cụm dài thay "là" | "có thể được xem là" | "là" |
| V18 | Chủ thể hoặc quy chiếu mơ hồ giữa các vế | "Nó sai vì nó chưa..." | gọi tên đối tượng |
| V19 | Thuật ngữ Anh không hợp người đọc | "dedupe cột key" cho người mới | "xóa dòng trùng theo cột mã" (xem `tu-vung-khoa.md`) |
| V20 | Từ bị cụt | "rười", "hiểm thị" | từ đầy đủ |
| V21 | Câu trợ lý còn sót | "Chắc chắn rồi!", "Dưới đây là..." | xóa |
| V22 | Rào rồi khẳng định điều chưa kiểm | "có lẽ đúng, tổng là 1.234" | kiểm trước, hoặc ghi "chưa kiểm chứng" |
| V23 | Bác bỏ ý chưa ai nêu | "Đừng nghĩ đây là lỗi hiếm" | xóa, hoặc nêu ý bị bác bỏ trước |
| V24 | Chồng tình thái | "có thể có lẽ" | một từ |

## Typography (T1-T6)

Đã gộp vào `quy-tac-chung.md`: T1 Title Case, T2 gạch dài, T3 ngoặc kép lẫn lộn, T4 phẩy trước "và", T5 nhãn trước dấu hai chấm không thêm nghĩa, T6 emoji trang trí.

## Áp dụng cho khóa

- Kịch bản và notes: lời giảng được **đọc thành tiếng**; ưu tiên V9, V11, V15, V16, V19. Câu phải đọc một hơi được.
- Nhận xét học viên: ưu tiên V1, V3, V10, V18. Nêu rõ đối tượng ("cột Quantity", "ô C5"), tránh "nó", "cái này".
- Báo cáo rà soát: ưu tiên V11, V15, V22, V24.
