"""Read-only Codex lifecycle context for the MCI Excel workspace."""
import json
from pathlib import Path
import sys


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    try:
        event = json.load(sys.stdin)
    except (ValueError, UnicodeError):
        return
    if not isinstance(event, dict):
        return
    name = event.get("hook_event_name")
    if name not in {"SessionStart", "SubagentStart"}:
        return
    root = Path(__file__).resolve().parents[2]
    cwd = Path(event.get("cwd") or root).resolve()
    if not cwd.is_relative_to(root):
        return
    context = (
        "MCI Excel: đọc AGENTS.md, README.md và quy ước liên quan trong CLAUDE.md. "
        "Course_Plan.md là nguồn chương trình; tra lệnh dựng trong 2_Dataset/README.md. "
        "Skill dự án ở .agents/skills/: mci-session-builder, mci-screenshot-capturer, "
        "mci-homework-grader, mci-consistency-checker, viet-van-phong-mci. "
        "Chạy trên Windows: kiểm tra backend/dependency, không dùng đường dẫn Mac. "
        "Không sửa bài nộp; không lưu workbook có Pivot/Model/validation mở rộng bằng openpyxl. "
        "Tài liệu phát cần hết [CHỤP] và bản học viên không có notes/đáp án. "
        "Chỉ điều khiển một tác vụ Office tại một thời điểm; báo phần chưa kiểm chứng."
    )
    print(json.dumps({"hookSpecificOutput": {"hookEventName": name,
                                               "additionalContext": context}},
                     ensure_ascii=False))


if __name__ == "__main__":
    main()
