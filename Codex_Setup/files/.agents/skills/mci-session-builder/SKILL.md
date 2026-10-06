---
name: mci-session-builder
description: Dựng hoặc chốt tài liệu một buổi MCI Excel bằng script hiện có.
---

# Dựng hoặc chốt tài liệu một buổi MCI Excel bằng script hiện có.

Đọc AGENTS.md ở gốc workspace và [references/workflow.md](references/workflow.md)
để thực hiện quy trình chi tiết chuyển từ agent `session-builder` của Claude.
Quy trình gốc được giữ để bảo toàn tiêu chí và đầu ra; phần thích nghi dưới
đây có ưu tiên khi hướng dẫn gốc phụ thuộc Claude, Mac hoặc công cụ không có.
Không coi cấu hình tools/model/effort/permissions của Claude là quyền Codex.
Lấy workspace hiện tại làm gốc thay cho đường dẫn cá nhân trong tài liệu gốc.

## Môi trường Codex hiện tại

Workspace chạy trên Windows/PowerShell. Xác định Python bằng Get-Command python;
không dùng đường dẫn anaconda của Mac. Đọc script trước khi chạy để phát hiện
AppleScript, osascript, đường dẫn cố định và dependency. Lệnh CHỈ WINDOWS được
chạy trên Windows khi thuộc buổi được giao. Script chỉ hỗ trợ Mac thì báo phần
chưa chạy được; không tự thay backend trong vai trò dựng tài liệu.
Không khởi động hai tác vụ cùng điều khiển Excel/Word/PowerPoint.

Kiểm tra file khóa ~$ và tài liệu đang mở qua COM nếu công cụ cho phép.
Chỉ yêu cầu lưu/đóng khi đúng file sắp ghi đang mở, hoặc có xung đột OneDrive
chưa giải quyết; không hỏi xác nhận đóng mọi ứng dụng theo mặc định.
Xuất PDF trên Windows bằng Word COM nếu có, trên bản sao trong thư mục tạm.
Nếu Office automation cần quyền ngoài sandbox, xin quyền đúng lệnh đó.
Không coi dữ liệu cached rỗng là kết quả số bằng 0.


## Phạm vi và đầu ra

Thực hiện đúng vai trò, không tự chuyển sang sửa nguồn/dựng/chấm ngoài phạm vi.
Nếu người dùng yêu cầu công việc rộng hơn, phiên chính xử lý phần ngoài vai trò.
Chỉ báo điều đã đọc, tính lại hoặc quan sát; ghi rõ phần chưa kiểm chứng.
Trước khi viết báo cáo/nhận xét, đọc skill văn phong ở
`../viet-van-phong-mci/SKILL.md` (tính từ thư mục skill này).
