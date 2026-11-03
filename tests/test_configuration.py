from pathlib import Path
import json
import sys
import tempfile
import unittest
from unittest.mock import patch
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import configuration as config

class ConfigurationTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix="safeguard settings ' ")
        self.root = Path(self.temp.name)
        self.patch = patch.object(config, 'ROOT', self.root)
        self.patch.start()
        self.ps = 'C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe'
    def tearDown(self):
        self.patch.stop()
        self.temp.cleanup()
    def test_fresh_json_no_bom_and_blocking_exec_hooks(self):
        config.update(self.ps, fresh=True)
        config.check(self.ps)
        for name in ['.mcp.json', '.claude/settings.json']:
            data = (self.root / name).read_bytes()
            self.assertFalse(data.startswith(b'\xef\xbb\xbf'))
            json.loads(data.decode('utf-8'))
        settings = config.load(self.root / '.claude/settings.json')
        for event in config.EVENTS:
            hook = settings['hooks'][event][0]['hooks'][0]
            self.assertEqual(hook['onFailure'], 'block')
            self.assertEqual(hook['command'], self.ps)
            self.assertEqual(hook['args'][-1], str(self.root / 'guard.ps1'))
    def test_migrate_bom_old_hooks_preserve_unrelated_and_repeat(self):
        (self.root / '.claude').mkdir()
        legacy = 'powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "' + str(self.root / 'guard.ps1') + '"'
        other = {'type': 'command', 'command': 'unrelated-check.exe'}
        settings = {'theme': 'dark', 'env': {'KEEP': 'yes'}, 'permissions': {'deny': ['CustomTool']},
                    'hooks': {'PreToolUse': [{'hooks': [{'type': 'command', 'command': legacy}, other]}]}}
        (self.root / '.claude/settings.json').write_text(json.dumps(settings), encoding='utf-8-sig')
        (self.root / '.mcp.json').write_text(json.dumps({'mcpServers': {'other': {'command': 'other.exe'}}}), encoding='utf-8-sig')
        config.update(self.ps)
        before = (self.root / '.claude/settings.json').read_bytes()
        config.update(self.ps)
        self.assertEqual(before, (self.root / '.claude/settings.json').read_bytes())
        current = config.load(self.root / '.claude/settings.json')
        self.assertEqual(current['theme'], 'dark')
        self.assertEqual(current['env']['KEEP'], 'yes')
        self.assertIn('CustomTool', current['permissions']['deny'])
        self.assertEqual(current['hooks']['PreToolUse'][0]['hooks'], [other])
        self.assertIn('other', config.load(self.root / '.mcp.json')['mcpServers'])
    def test_missing_hook_refuses_check(self):
        config.update(self.ps, fresh=True)
        path = self.root / '.claude/settings.json'
        current = config.load(path)
        current['hooks'].pop('UserPromptSubmit')
        path.write_text(json.dumps(current))
        with self.assertRaises(ValueError): config.check(self.ps)
    def test_disabled_hooks_refuse_migration_without_write(self):
        config.update(self.ps, fresh=True)
        path = self.root / '.claude/settings.json'
        current = config.load(path)
        current['disableAllHooks'] = True
        path.write_text(json.dumps(current))
        before = path.read_bytes()
        with self.assertRaises(ValueError): config.update(self.ps)
        self.assertEqual(before, path.read_bytes())
    def test_duplicate_keys_refused(self):
        path = self.root / 'duplicate.json'
        path.write_text('{"hooks": {}, "hooks": {}}')
        with self.assertRaises(ValueError): config.load(path)

if __name__ == '__main__': unittest.main(verbosity=2)
