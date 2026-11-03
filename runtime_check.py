"""Local checks using only synthetic data; no Claude or Microsoft requests."""
from pathlib import Path
import asyncio
import importlib
import importlib.metadata
import json
import shutil
import sys
import tempfile

ROOT = Path(__file__).resolve().parent

def imports(sharepoint=False):
    files = ['requirements.txt'] + (['requirements-sharepoint.txt'] if sharepoint else [])
    for filename in files:
        for line in (ROOT / filename).read_text(encoding='utf-8').splitlines():
            if not line or line.startswith('#'):
                continue
            name, version = line.split('==')
            if importlib.metadata.version(name) != version:
                raise RuntimeError('Dependency versions do not match the tested package.')
    from mcp.server.fastmcp import FastMCP
    for name in ['pypdf', 'openpyxl', 'pandas', 'spreadsheet', 'finance', 'folder_access', 'convert', 'sharepoint']:
        importlib.import_module(name)
    if sharepoint:
        import msal
        from msal_extensions import build_encrypted_persistence, PersistedTokenCache

async def smoke():
    from mcp import ClientSession, StdioServerParameters
    from mcp.client.stdio import stdio_client
    import anyio
    with tempfile.TemporaryDirectory(prefix='safeguard smoke ') as temp:
        root = Path(temp)
        for name in ['spreadsheet.py', 'finance.py', 'folder_access.py', 'sharepoint.py', 'convert.py']:
            shutil.copy2(ROOT / name, root / name)
        params = StdioServerParameters(command=sys.executable, args=[str(root / 'spreadsheet.py')])
        with anyio.fail_after(45):
            async with stdio_client(params) as (read, write):
                async with ClientSession(read, write) as session:
                    await session.initialize()
                    names = [tool.name for tool in (await session.list_tools()).tools]
                    if names != ['spreadsheet']:
                        raise RuntimeError('Unexpected MCP tools.')
                    async def call(request):
                        result = await session.call_tool('spreadsheet', {'request': request})
                        if result.isError:
                            raise RuntimeError('MCP call failed.')
                        return json.loads(result.content[0].text)
                    created = await call({'operation': 'create', 'sheet': 'Data',
                                          'rows': [['Department', 'Budget', 'Actual'], ['A', 100, 90]]})
                    result = await call({'operation': 'dashboard', 'workbook': created['workbook'],
                                         'sheet': 'Data', 'group_by': ['Department'], 'budget_column': 'Budget',
                                         'actual_column': 'Actual', 'currency': 'GBP'})
                    if result.get('report', {}).get('variance') != '-10':
                        raise RuntimeError('Synthetic dashboard result incorrect.')
                    blocked = await call({'operation': 'create', 'rows': [['test@example.invalid']]})
                    if 'error' not in blocked or 'test@example.invalid' in json.dumps(blocked):
                        raise RuntimeError('Synthetic sensitive-data check failed.')

if __name__ == '__main__':
    try:
        imports('--sharepoint' in sys.argv)
        if '--imports-only' not in sys.argv:
            asyncio.run(smoke())
        print('PASS: local runtime' + (' and synthetic MCP handshake/dashboard/refusal.' if '--imports-only' not in sys.argv else ' imports.'))
    except Exception:
        print('FAIL: local runtime check. Close Claude and run Repair.cmd from the latest extracted ZIP.')
        sys.exit(1)
