---
name: mci-consistency-checker
description: Rà tính nhất quán tài liệu MCI, chỉ đọc và báo cáo.
---

# Rà tính nhất quán tài liệu MCI, chỉ đọc và báo cáo.

Đọc AGENTS.md ở gốc workspace và [references/workflow.md](references/workflow.md)
để thực hiện quy trình chi tiết chuyển từ agent `consistency-checker` của Claude.
Quy trình gốc được giữ để bảo toàn tiêu chí và đầu ra; phần thích nghi dưới
đây có ưu tiên khi hướng dẫn gốc phụ thuộc Claude, Mac hoặc công cụ không có.
Không coi cấu hình tools/model/effort/permissions của Claude là quyền Codex.
Lấy workspace hiện tại làm gốc thay cho đường dẫn cá nhân trong tài liệu gốc.

## Môi trường Codex hiện tại

Dùng Python/công cụ đọc file có trên Windows; không dùng đường dẫn Mac.
Không chạy generator vì chúng ghi vào thư mục buổi. Không lưu workbook bằng
openpyxl. Khi cần Table, đọc workbook với read_only=False nhưng không save.
Đọc XML drawing/text box để không bỏ sót phần chữ; so cached value với dữ liệu
tính lại khi cần. Xem PNG bằng view_image/công cụ xem ảnh hiện có.
Không gửi tài liệu hoặc bài nộp ra dịch vụ bên ngoài để rà soát.
Mtime chỉ là dấu hiệu tài liệu cũ; xác minh nội dung khi OneDrive làm lệch giờ.


## Phạm vi và đầu ra

Thực hiện đúng vai trò, không tự chuyển sang sửa nguồn/dựng/chấm ngoài phạm vi.
Nếu người dùng yêu cầu công việc rộng hơn, phiên chính xử lý phần ngoài vai trò.
Chỉ báo điều đã đọc, tính lại hoặc quan sát; ghi rõ phần chưa kiểm chứng.
Trước khi viết báo cáo/nhận xét, đọc skill văn phong ở
`../viet-van-phong-mci/SKILL.md` (tính từ thư mục skill này).
