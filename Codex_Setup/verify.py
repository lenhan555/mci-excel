"""Validate staged skills and lifecycle hook behavior without starting a model."""
from pathlib import Path
import json
import os
import subprocess
import sys
import tomllib

ROOT = Path(__file__).resolve().parent.parent
FILES = ROOT / 'Codex_Setup/files'
VALIDATOR = Path('C:/Users/Admin/.codex/skills/.system/skill-creator/scripts/quick_validate.py')

for skill in sorted((FILES / '.agents/skills').iterdir()):
    result = subprocess.run([sys.executable, '-X', 'utf8', str(VALIDATOR), str(skill)],
                            capture_output=True, text=True, encoding='utf-8')
    print(skill.name + ': ' + result.stdout.strip())
    if result.returncode:
        raise SystemExit(result.stderr or result.stdout)

hook = FILES / '.codex/hooks/workspace_context.py'
before = {p.relative_to(FILES): (p.stat().st_size, p.stat().st_mtime_ns)
          for p in FILES.rglob('*') if p.is_file()}
for name in ('SessionStart', 'SubagentStart'):
    for cwd in (FILES, FILES / '1_Sessions/B6_Dashboard'):
        result = subprocess.run([sys.executable, str(hook)],
                                input=json.dumps({'hook_event_name': name, 'cwd': str(cwd)}),
                                text=True, capture_output=True, encoding='utf-8', check=True)
        parsed = json.loads(result.stdout)['hookSpecificOutput']
        assert parsed['hookEventName'] == name
        assert 'Course_Plan.md' in parsed['additionalContext']
for payload in ('not JSON', '[]', json.dumps({'hook_event_name': 'Stop'}),
                json.dumps({'hook_event_name': 'SessionStart', 'cwd': str(ROOT.parent)})):
    result = subprocess.run([sys.executable, str(hook)], input=payload,
                            capture_output=True, text=True, encoding='utf-8', check=True)
    assert result.stdout == ''
after = {p.relative_to(FILES): (p.stat().st_size, p.stat().st_mtime_ns)
         for p in FILES.rglob('*') if p.is_file()}
assert before == after, 'Hook unexpectedly wrote a file'
reviewer = tomllib.loads((FILES / '.codex/agents/consistency-checker.toml').read_text(encoding='utf-8'))
assert reviewer['sandbox_mode'] == 'read-only'
config = json.loads((FILES / '.codex/hooks.json').read_text(encoding='utf-8'))
assert set(config['hooks']) == {'SessionStart', 'SubagentStart'}
print('PASS: 5 skills, hook lifecycle payloads, nested cwd, invalid/unrelated events, no writes, reviewer read-only.')
