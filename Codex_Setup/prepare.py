"""Prepare and validate project-local Codex configuration; never change Claude files."""
from pathlib import Path
import json
import hashlib
import re
import shutil
import sys
import tomllib
from datetime import datetime

ROOT = Path(__file__).resolve().parent.parent
STAGE = ROOT / 'Codex_Setup' / 'files'

ROLES = {
    'session-builder': ('mci-session-builder', 'Dựng hoặc chốt tài liệu một buổi MCI Excel bằng script hiện có.', 'workspace-write'),
    'screenshot-capturer': ('mci-screenshot-capturer', 'Chụp ảnh Excel Windows từ file thật cho slide và handout MCI.', 'workspace-write'),
    'homework-grader': ('mci-homework-grader', 'Chấm bài Excel MCI, viết nhận xét và báo cáo lớp có bằng chứng.', 'workspace-write'),
    'consistency-checker': ('mci-consistency-checker', 'Rà tính nhất quán tài liệu MCI, chỉ đọc và báo cáo.', 'read-only'),
}

ADAPTATIONS = {
    'session-builder': '''
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
''',
    'screenshot-capturer': '''
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
''',
    'homework-grader': '''
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
''',
    'consistency-checker': '''
## Môi trường Codex hiện tại

Dùng Python/công cụ đọc file có trên Windows; không dùng đường dẫn Mac.
Không chạy generator vì chúng ghi vào thư mục buổi. Không lưu workbook bằng
openpyxl. Khi cần Table, đọc workbook với read_only=False nhưng không save.
Đọc XML drawing/text box để không bỏ sót phần chữ; so cached value với dữ liệu
tính lại khi cần. Xem PNG bằng view_image/công cụ xem ảnh hiện có.
Không gửi tài liệu hoặc bài nộp ra dịch vụ bên ngoài để rà soát.
Mtime chỉ là dấu hiệu tài liệu cũ; xác minh nội dung khi OneDrive làm lệch giờ.
''',
}

AGENTS = '''# MCI Excel — hướng dẫn cho Codex

Workspace xây dựng và vận hành khóa Excel for Business Intelligence Analyst
của MCI: 7 buổi × 150 phút, case MCI Mart, capstone B7.
Trả lời và viết tài liệu bằng tiếng Việt.

## Nguồn và quy tắc

- Đọc README.md và các phần liên quan của CLAUDE.md trước khi sửa tài liệu.
  CLAUDE.md là quy ước nội dung dùng chung; phần chỉ dành cho Claude/Mac được
  thay bằng hướng dẫn Codex/Windows trong các skill tương ứng.
- Course_Plan.md là nguồn tổng thể. Thay đổi chương trình thì sửa plan trước,
  sau đó kịch bản và script sinh tài liệu. Khi đối soát con số/hành vi Excel,
  dùng file và ảnh thật làm bằng chứng; báo sai lệch của tài liệu nguồn.
- Tái sử dụng 2_Dataset/_generator/; tra lệnh trong 2_Dataset/README.md.
  Không sửa tay sản phẩm sinh từ script rồi bỏ qua nguồn.
- Không ghi lại bằng openpyxl file đã được Excel lưu nếu có validation mở
  rộng, Pivot, Power Query hoặc Data Model. Dùng Excel COM hoặc dựng lại bằng
  backend phù hợp, có sao lưu trước khi ghi đè sản phẩm.
- Bài học viên nộp chỉ đọc. Không dùng _capture/ hoặc bản backup làm nguồn
  phát học viên. Tách bản giảng viên/học viên theo quy ước hiện có.
- Dùng ảnh chụp Excel thật. Tài liệu phát phải hết [CHỤP], bản _HocVien không
  lộ notes/đáp án, PDF và infographic được dựng từ nguồn hiện hành.

## Skill và agent

Skill dự án nằm trong .agents/skills/. Chọn đúng phạm vi và đọc SKILL.md:

| Yêu cầu | Skill | Agent |
|---|---|---|
| Dựng khung/chốt B<n> | mci-session-builder | session_builder |
| Chụp ảnh Excel B<n> | mci-screenshot-capturer | screenshot_capturer |
| Chấm bài B<n> | mci-homework-grader | homework_grader |
| Rà tài liệu | mci-consistency-checker | consistency_checker |
| Viết/kiểm văn phong | viet-van-phong-mci | Dùng trong vai trò liên quan |

Khi dựng trọn một buổi, phiên chính điều phối lần lượt: khung → chụp ảnh thiếu
→ chốt → rà nhất quán. Có thể giao agent đúng vai trò khi công cụ delegation
có sẵn; không spawn để chỉ trả lời câu hỏi đơn giản. Agent chỉ ghi trong phạm
vi vai trò được giao, không tự gọi thêm agent. Không chạy song song tác vụ
cùng ghi một file hay điều khiển Office. Phiên chính xử lý lỗi nguồn nằm ngoài
phạm vi agent, theo yêu cầu người dùng.

## Windows và quyền thực thi

Phát hiện Python/dependency và khả năng COM trước khi chạy. Không dùng đường
dẫn Mac cố định, không giả định AppleScript chạy trên Windows. Không nhập
permissions.allow của Claude thành quyền Codex; tuân thủ sandbox hiện tại.
Chỉ yêu cầu lưu/đóng tài liệu khi phát hiện đúng file sắp ghi đang mở hoặc có
xung đột chưa giải quyết. Không hỏi lại quyền đã được cấp trong phiên.
Hook dự án chỉ bổ sung ngữ cảnh, không dựng tài liệu/chấm bài/tự cấp quyền.
'''

HOOK = '''"""Read-only Codex lifecycle context for the MCI Excel workspace."""
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
'''


def write(rel, content):
    if str(rel).endswith('.md'):
        for old in ROLES:
            content = content.replace(old.replace('-', '_'), old)
    path = STAGE / rel
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content.strip() + '\n', encoding='utf-8')


def prepare():
    existing_agents = ROOT / 'AGENTS.md'
    if existing_agents.exists():
        existing = existing_agents.read_text(encoding='utf-8')
        existing = re.sub(r'\.Codex/agents/([a-z-]+)\.md', r'.codex/agents/\1.toml', existing)
        existing = existing.replace('.Codex/skills/', '.agents/skills/')
        existing = re.sub(r'- \*\*Khuyến khích chấm bằng Codex Cowork\.\*\*[^\n]*',
                          '- Đọc cả ô và XML drawing/text box để kiểm tra insight; '
                          'openpyxl chỉ đọc ô có thể bỏ sót phần chữ. Không giả định có Cowork. '
                          'Phần không đọc được ghi rõ Không kiểm được, cần GV xem.', existing)
        write('AGENTS.md', existing + '\n\n---\n\n'
              'Các hướng dẫn Codex/Windows dưới đây thay thế phần trên nếu có '
              'mâu thuẫn về môi trường, đường dẫn công cụ hoặc việc xác nhận file đang mở.\n\n' + AGENTS)
    else:
        write('AGENTS.md', AGENTS)
    write('.codex/config.toml', '[agents]\nenabled = true\nmax_concurrent_threads_per_session = 3')
    for old, (skill, description, sandbox) in ROLES.items():
        src = (ROOT / '.claude/agents' / (old + '.md')).read_text(encoding='utf-8')
        old_toml = ROOT / '.codex/agents' / (old + '.toml')
        prior = tomllib.loads(old_toml.read_text(encoding='utf-8')) if old_toml.exists() else {}
        body = prior.get('developer_instructions') or re.sub(r'\A---\s*\n.*?\n---\s*\n', '', src, count=1, flags=re.S)
        body = body.replace('.claude/skills/viet-van-phong-mci', '.agents/skills/viet-van-phong-mci')
        body = body.replace('.Codex/skills/', '.agents/skills/')
        body = body.replace('bằng `Read`', 'bằng công cụ xem ảnh/đọc file hiện có')
        role = old
        instruction = (
            f'Bạn đảm nhiệm vai trò {old} cho workspace MCI Excel. Trả lời bằng tiếng Việt. '
            f'Đọc .agents/skills/{skill}/SKILL.md và references/workflow.md trước khi làm. '
            'Chỉ làm phạm vi được giao; không tự gọi agent khác. '
            'Không áp dụng tools/model/effort/permissions của Claude. '
            'Ghi rõ bằng chứng, đầu ra và phần chưa kiểm chứng.'
        )
        toml = '\n'.join([
            f'name = {json.dumps(role)}',
            f'description = {json.dumps(description, ensure_ascii=False)}',
            f'sandbox_mode = {json.dumps(sandbox)}',
            f'developer_instructions = {json.dumps(instruction, ensure_ascii=False)}',
        ])
        write(f'.codex/agents/{old}.toml', toml)
        entry = f'''---
name: {skill}
description: {description}
---

# {description}

Đọc AGENTS.md ở gốc workspace và [references/workflow.md](references/workflow.md)
để thực hiện quy trình chi tiết chuyển từ agent `{old}` của Claude.
Quy trình gốc được giữ để bảo toàn tiêu chí và đầu ra; phần thích nghi dưới
đây có ưu tiên khi hướng dẫn gốc phụ thuộc Claude, Mac hoặc công cụ không có.
Không coi cấu hình tools/model/effort/permissions của Claude là quyền Codex.
Lấy workspace hiện tại làm gốc thay cho đường dẫn cá nhân trong tài liệu gốc.
{ADAPTATIONS[old]}

## Phạm vi và đầu ra

Thực hiện đúng vai trò, không tự chuyển sang sửa nguồn/dựng/chấm ngoài phạm vi.
Nếu người dùng yêu cầu công việc rộng hơn, phiên chính xử lý phần ngoài vai trò.
Chỉ báo điều đã đọc, tính lại hoặc quan sát; ghi rõ phần chưa kiểm chứng.
Trước khi viết báo cáo/nhận xét, đọc skill văn phong ở
`../../viet-van-phong-mci/SKILL.md` (tính từ thư mục skill này).
'''
        # sibling skill location relative to SKILL.md
        entry = entry.replace('../../viet-van-phong-mci/SKILL.md', '../viet-van-phong-mci/SKILL.md')
        write(f'.agents/skills/{skill}/SKILL.md', entry)
        write(f'.agents/skills/{skill}/references/workflow.md',
              '# Quy trình nguồn đã chuyển từ Claude\n\n'
              'Đọc phần thích nghi trong ../SKILL.md trước; phần đó thay thế mọi\n'
              'chỉ dẫn Mac/Claude không phù hợp bên dưới.\n\n' + body)
    style = ROOT / '.claude/skills/viet-van-phong-mci'
    for source in style.rglob('*'):
        if source.is_file():
            rel = source.relative_to(style)
            existing = ROOT / '.agents/skills/viet-van-phong-mci' / rel
            content = (existing if existing.exists() else source).read_text(encoding='utf-8')
            write(str(Path('.agents/skills/viet-van-phong-mci') / rel), content)
    write('.codex/hooks/workspace_context.py', HOOK)
    interpreter = Path(sys.executable).as_posix()
    script = (ROOT / '.codex/hooks/workspace_context.py').as_posix()
    handler = {'type': 'command', 'command': f'"{interpreter}" "{script}"',
               'timeout': 10, 'additionalContextLimit': 1500}
    hooks = {'description': 'MCI Excel: bổ sung ngữ cảnh, chỉ đọc, không cấp quyền.',
             'hooks': {
                 'SessionStart': [{'matcher': '^(startup|resume|clear|compact)$', 'hooks': [handler]}],
                 'SubagentStart': [{'hooks': [handler]}],
             }}
    write('.codex/hooks.json', json.dumps(hooks, ensure_ascii=False, indent=2))
    write('.codex/README.md', '''# Cấu hình Codex cho MCI Excel

AGENTS.md ở gốc chứa quy tắc chung. Agent TOML trong agents/ kế thừa model và
reasoning của phiên chính; consistency_checker mặc định read-only.
Skill ở .agents/skills/ được phát hiện theo workspace.

Ví dụ yêu cầu:
- Dùng session_builder dựng khung B6.
- Dùng screenshot_capturer chụp các ảnh B6 còn thiếu.
- Dùng homework_grader chấm bài B2.
- Dùng consistency_checker rà tài liệu B5 trước khi phát.
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
''')
    print(f'Prepared {sum(p.is_file() for p in STAGE.rglob("*"))} files in {STAGE}')
    baseline = {}
    for staged in STAGE.rglob('*'):
        if staged.is_file():
            rel = staged.relative_to(STAGE)
            target = ROOT / rel
            baseline[str(rel)] = hashlib.sha256(target.read_bytes()).hexdigest() if target.exists() else None
    (ROOT / 'Codex_Setup/baseline.json').write_text(json.dumps(baseline, indent=2), encoding='utf-8')


def validate():
    for path in STAGE.rglob('*.toml'):
        parsed = tomllib.loads(path.read_text(encoding='utf-8'))
        if path.parent.name == 'agents':
            assert all(k in parsed for k in ('name', 'description', 'developer_instructions'))
    json.loads((STAGE / '.codex/hooks.json').read_text(encoding='utf-8'))
    compile(HOOK, 'workspace_context.py', 'exec')
    for skill in (STAGE / '.agents/skills').iterdir():
        body = (skill / 'SKILL.md').read_text(encoding='utf-8')
        assert body.startswith('---\n') and f'name: {skill.name}\n' in body
    print('TOML, JSON, hook syntax and 5 skill entrypoints valid.')


def install():
    validate()
    files = [p for p in STAGE.rglob('*') if p.is_file()]
    destinations = [(p, ROOT / p.relative_to(STAGE)) for p in files]
    baseline = json.loads((ROOT / 'Codex_Setup/baseline.json').read_text(encoding='utf-8'))
    for _, target in destinations:
        assert target.resolve().is_relative_to(ROOT.resolve())
        actual = hashlib.sha256(target.read_bytes()).hexdigest() if target.exists() else None
        if actual != baseline[str(target.relative_to(ROOT))]:
            raise SystemExit(f'Configuration changed since preparation; refusing overwrite: {target}')
    backup = ROOT / 'Codex_Setup/backups' / datetime.now().strftime('%Y%m%d_%H%M%S')
    for source, target in destinations:
        if target.exists() and target.read_bytes() != source.read_bytes():
            saved = backup / target.relative_to(ROOT)
            saved.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(target, saved)
    for source, target in destinations:
        if target.exists() and target.read_bytes() == source.read_bytes():
            continue
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, target)
    print(f'Installed {len(files)} project files. Claude files untouched. Hooks not auto-trusted.')


if __name__ == '__main__':
    if '--install' in sys.argv:
        install()
    else:
        prepare()
        validate()
