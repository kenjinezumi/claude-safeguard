"""Install/migrate owned settings, retaining unrelated keys. Never reads user data."""
from pathlib import Path
import json
import sys

ROOT = Path(__file__).resolve().parent
EVENTS = ('SessionStart', 'UserPromptSubmit', 'PreToolUse')
DENY = ['Bash', 'PowerShell', 'Grep', 'Glob', 'Write', 'Edit', 'NotebookEdit',
        'WebFetch', 'WebSearch', 'Agent', 'Task', 'Skill']

def load(path):
    def unique(pairs):
        result = {}
        for key, value in pairs:
            if key in result:
                raise ValueError('Duplicate settings key')
            result[key] = value
        return result
    return json.loads(path.read_text(encoding='utf-8-sig'), object_pairs_hook=unique)

def handler(powershell):
    return {'type': 'command', 'command': powershell,
            'args': ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', str(ROOT / 'guard.ps1')],
            'timeout': 20, 'onFailure': 'block'}

def server():
    return {'type': 'stdio', 'command': str(ROOT / 'runtime' / 'Scripts' / 'python.exe'),
            'args': [str(ROOT / 'spreadsheet.py')]}

def update(powershell, fresh=False):
    settings_path = ROOT / '.claude' / 'settings.json'
    mcp_path = ROOT / '.mcp.json'
    settings = {} if fresh else load(settings_path)
    mcp = {} if fresh else load(mcp_path)
    if settings.get('disableAllHooks'):
        raise ValueError('Hooks are disabled in existing settings. Ask IT to review.')
    hooks = settings.setdefault('hooks', {})
    guard = str(ROOT / 'guard.ps1')
    old_command = 'powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "' + guard + '"'
    for event in EVENTS:
        entries = []
        for entry in hooks.get(event, []):
            remaining = [h for h in entry.get('hooks', [])
                         if not (h.get('command') == old_command or guard in h.get('args', []))]
            if remaining:
                entries.append(dict(entry, hooks=remaining))
        entries.append({'hooks': [handler(powershell)]})
        hooks[event] = entries
    env = settings.setdefault('env', {})
    env.update(ENABLE_TOOL_SEARCH='false', CLAUDE_CODE_DISABLE_NONESSENTIAL_TRAFFIC='1')
    permissions = settings.setdefault('permissions', {})
    permissions['defaultMode'] = 'default'
    permissions['deny'] = list(dict.fromkeys(permissions.get('deny', []) + DENY))
    mcp.setdefault('mcpServers', {})['claude_safeguard'] = server()
    settings_path.parent.mkdir(exist_ok=True)
    # UTF-8 without BOM, including when called by Windows PowerShell 5.1.
    settings_path.write_text(json.dumps(settings, indent=2) + '\n', encoding='utf-8')
    mcp_path.write_text(json.dumps(mcp, indent=2) + '\n', encoding='utf-8')
    check(powershell)

def check(powershell):
    settings = load(ROOT / '.claude' / 'settings.json')
    mcp = load(ROOT / '.mcp.json')
    if settings.get('disableAllHooks'):
        raise ValueError('Hooks disabled')
    for event in EVENTS:
        if not any(not e.get('matcher') and handler(powershell) in e.get('hooks', [])
                   for e in settings.get('hooks', {}).get(event, [])):
            raise ValueError('Required guard configuration missing or changed')
    if not set(DENY).issubset(settings.get('permissions', {}).get('deny', [])):
        raise ValueError('Required tool restrictions missing')
    if mcp.get('mcpServers', {}).get('claude_safeguard') != server():
        raise ValueError('Bundled MCP configuration missing or changed')

if __name__ == '__main__':
    try:
        mode, powershell = sys.argv[1:]
        if mode == '--check':
            check(powershell)
        elif mode in ('--install', '--repair'):
            update(powershell, fresh=mode == '--install')
        else:
            raise ValueError('Unknown mode')
        print('PASS: package settings.')
    except Exception:
        print('FAIL: package settings. Close Claude and run Repair.cmd; ask IT if repair stops.')
        sys.exit(1)
