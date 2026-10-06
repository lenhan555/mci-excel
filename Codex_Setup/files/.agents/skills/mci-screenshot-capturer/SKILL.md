---
name: mci-screenshot-capturer
description: Chụp ảnh Excel Windows từ file thật cho slide và handout MCI.
---

# Chụp ảnh Excel Windows từ file thật cho slide và handout MCI.

Đọc AGENTS.md ở gốc workspace và [references/workflow.md](references/workflow.md)
để thực hiện quy trình chi tiết chuyển từ agent `screenshot-capturer` của Claude.
Quy trình gốc được giữ để bảo toàn tiêu chí và đầu ra; phần thích nghi dưới
đây có ưu tiên khi hướng dẫn gốc phụ thuộc Claude, Mac hoặc công cụ không có.
Không coi cấu hình tools/model/effort/permissions của Claude là quyền Codex.
Lấy workspace hiện tại làm gốc thay cho đường dẫn cá nhân trong tài liệu gốc.

## Môi trường Codex hiện tại

Đang chạy trực tiếp trên Windows. Đọc lib.ps1 và watcher.ps1 của buổi trước
khi dùng. Ưu tiên chạy công cụ chụp hiện có trực tiếp nếu có quyền truy cập GUI;
watcher chỉ cần cho điều khiển từ máy khác hoặc khi đã được người dùng bật.
Không tự bật watcher chạy nền hay tăng quyền để né sandbox.
Nếu dùng watcher, giữ nguyên giao thức cmd/out/DONE; mỗi lượt chờ tối đa
60 giây rồi kiểm tra log và cập nhật tiến độ, tổng timeout tối đa 3 phút.
Giữ số thứ tự lệnh tăng dần; không dùng tên biến của watcher.
Get-XlFor không kết nối được thì kiểm tra mức quyền thay vì lặp vô hạn.
Dùng view_image hoặc công cụ xem ảnh hiện có để xem PNG. Cắt ảnh chụp bằng
crop.py/Pillow hiện có; không dùng AI tạo hoặc chỉnh nội dung ảnh Excel.
Chỉ yêu cầu người dùng dựng hộp thoại khi công cụ hiện có không thao tác được.
Không dừng watcher do người dùng mở trừ khi được yêu cầu hoặc đã được giao
quản lý vòng đời watcher đó. Không đóng tài liệu/cửa sổ khác của người dùng.


## Phạm vi và đầu ra

Thực hiện đúng vai trò, không tự chuyển sang sửa nguồn/dựng/chấm ngoài phạm vi.
Nếu người dùng yêu cầu công việc rộng hơn, phiên chính xử lý phần ngoài vai trò.
Chỉ báo điều đã đọc, tính lại hoặc quan sát; ghi rõ phần chưa kiểm chứng.
Trước khi viết báo cáo/nhận xét, đọc skill văn phong ở
`../viet-van-phong-mci/SKILL.md` (tính từ thư mục skill này).
