# Cấu hình Codex cho MCI Excel

AGENTS.md ở gốc chứa quy tắc chung. Agent TOML trong agents/ kế thừa model và
reasoning của phiên chính; consistency-checker mặc định read-only.
Skill ở .agents/skills/ được phát hiện theo workspace.

Ví dụ yêu cầu:
- Dùng session-builder dựng khung B6.
- Dùng screenshot-capturer chụp các ảnh B6 còn thiếu.
- Dùng homework-grader chấm bài B2.
- Dùng consistency-checker rà tài liệu B5 trước khi phát.
- Dùng $viet-van-phong-mci kiểm văn phong nhận xét.

Hai hook SessionStart và SubagentStart chỉ nhắc quy tắc, không ghi file và
không chạy generator. Sau khi cài, mở phiên mới trong workspace để nạp cấu
hình/agent/skill. Trong Codex CLI, dùng /hooks để xem và trust hai định nghĩa
hook. Codex bỏ qua hook chưa được người dùng trust; bộ cài không tự trust.
Nếu app chưa hỗ trợ review hook, dùng CLI để review. AGENTS.md và skill vẫn
có thể đọc độc lập khi hook chưa hoạt động.

Lệnh hook trỏ tới Python và workspace Windows hiện tại để chạy được cả khi
mở từ thư mục con. Nếu chuyển workspace/máy/Python, sửa đường dẫn hooks.json
và review lại hook. Không nhập permissions.allow từ Claude.

Nguồn chuyển đổi: .claude/agents/ (4 vai trò), .claude/skills/viet-van-phong-mci/.
Các file .claude/settings*.json hiện không có hooks, chỉ có permissions.allow;
hai hook trên là phần bổ sung cho Codex. Bản quy trình chi tiết nằm trong
references/workflow.md của mỗi skill; hướng dẫn thích nghi Windows tại
SKILL.md có ưu tiên hơn phần gốc chỉ dành cho Mac/Claude.

Tài liệu chính thức đã đối chiếu:
- https://learn.chatgpt.com/docs/agent-configuration/subagents
- https://learn.chatgpt.com/docs/build-skills
- https://learn.chatgpt.com/docs/hooks

Kiểm chứng bộ cài: TOML/JSON hợp lệ, 5 skill hợp lệ, hook chạy trực tiếp với
payload mẫu và không ghi file. Việc Codex phát hiện agent/skill và hook chạy
trong phiên mới cần kiểm tra sau khi reload và trust; chưa coi là đã chạy
end-to-end dựng tài liệu, chụp ảnh hay chấm bài.
