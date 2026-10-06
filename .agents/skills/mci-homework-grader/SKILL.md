---
name: mci-homework-grader
description: Chấm bài Excel MCI, viết nhận xét và báo cáo lớp có bằng chứng.
---

# Chấm bài Excel MCI, viết nhận xét và báo cáo lớp có bằng chứng.

Đọc AGENTS.md ở gốc workspace và [references/workflow.md](references/workflow.md)
để thực hiện quy trình chi tiết chuyển từ agent `homework-grader` của Claude.
Quy trình gốc được giữ để bảo toàn tiêu chí và đầu ra; phần thích nghi dưới
đây có ưu tiên khi hướng dẫn gốc phụ thuộc Claude, Mac hoặc công cụ không có.
Không coi cấu hình tools/model/effort/permissions của Claude là quyền Codex.
Lấy workspace hiện tại làm gốc thay cho đường dẫn cá nhân trong tài liệu gốc.

## Môi trường Codex hiện tại

Dùng Python sẵn có trên Windows, không dùng đường dẫn Python của Mac.
Đọc cả ô, công thức và XML drawing/text box trong xlsx: insight có thể nằm
trong shape, openpyxl không đọc hết. Nếu không đọc được, ghi Không kiểm được,
cần GV xem; không kết luận thiếu insight chỉ từ dữ liệu ô.
Không mặc định có Claude Cowork hoặc quyền thao tác GUI.
Khi chạy grade_b<n>.py, đọc giao diện CLI của script trước, không đoán tham số.
Kiểm chứng trên đáp án và đề trước khi chấm. Không mở/lưu bài gốc bằng Excel.
Nếu cần tính lại bằng Office, chỉ dùng bản sao, không sửa bài nộp.
Giữ thang điểm người dùng đã chốt trong phiên; chỉ hỏi khi thật sự chưa có.


## Phạm vi và đầu ra

Thực hiện đúng vai trò, không tự chuyển sang sửa nguồn/dựng/chấm ngoài phạm vi.
Nếu người dùng yêu cầu công việc rộng hơn, phiên chính xử lý phần ngoài vai trò.
Chỉ báo điều đã đọc, tính lại hoặc quan sát; ghi rõ phần chưa kiểm chứng.
Trước khi viết báo cáo/nhận xét, đọc skill văn phong ở
`../viet-van-phong-mci/SKILL.md` (tính từ thư mục skill này).
