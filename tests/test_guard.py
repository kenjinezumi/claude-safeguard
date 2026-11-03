"""Synthetic guard regression checks. No real documents, Claude or network.
SAFEGUARD_TEST_PWSH=/path/to/pwsh python3 tests/test_guard.py
Can also use Windows powershell.exe; Windows execution has not been verified here.
"""
import json
import os
from pathlib import Path
import shutil
import subprocess
import tempfile
import unittest

SOURCE = Path(__file__).resolve().parents[1]
PWSH = os.environ.get('SAFEGUARD_TEST_PWSH') or shutil.which('pwsh') or shutil.which('powershell.exe')

@unittest.skipUnless(PWSH, 'Requires PowerShell')
class GuardTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory(prefix='safeguard guard ')
        self.root = Path(self.tmp.name)
        shutil.copy2(SOURCE / 'guard.ps1', self.root / 'guard.ps1')
        (self.root / 'CLAUDE.md').write_text('Synthetic policy', encoding='utf-8')
        self.approved = self.root / 'Approved-Data'
        self.approved.mkdir()

    def tearDown(self):
        self.tmp.cleanup()

    def file(self, name='policy with spaces.txt', data=b'Meal allowance: GBP 20.00 per day.'):
        path = self.approved / name
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(data)
        return path

    def run_guard(self, payload=None, check_file=None, raw=None):
        cmd = [PWSH, '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', str(self.root / 'guard.ps1')]
        if check_file is not None:
            cmd += ['-CheckFile', str(check_file)]
        return subprocess.run(cmd, input=raw if raw is not None else json.dumps(payload),
                              capture_output=True, encoding='utf-8', timeout=15)

    def read(self, path):
        return self.run_guard({'hook_event_name':'PreToolUse', 'tool_name':'Read',
                               'cwd':str(self.root), 'tool_input':{'file_path':str(path)}})

    def denied(self, result, code):
        self.assertEqual(result.returncode, 0, result.stderr)
        decision = json.loads(result.stdout)['hookSpecificOutput']
        self.assertEqual(decision['permissionDecision'], 'deny')
        self.assertIn('['+code+']', decision['permissionDecisionReason'])
        self.assertEqual(result.stderr, '')

    def allowed(self, result):
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertEqual(result.stdout, '')
        self.assertEqual(result.stderr, '')

    def test_valid_text_with_spaces(self): self.allowed(self.read(self.file()))
    def test_relative_nested_path(self):
        p = self.file('nested/policy.txt')
        self.allowed(self.read(p.relative_to(self.root)))
    def test_utf8_unicode_path_and_content(self):
        self.allowed(self.read(self.file('règles 日本.txt', 'Repas: café, coût 20 €. 日本'.encode('utf-8'))))
    def test_utf8_bom(self): self.allowed(self.read(self.file(data=b'\xef\xbb\xbfApproved test policy')))
    def test_missing_file(self): self.denied(self.read(self.approved / 'missing.txt'), 'SG-NOT-FOUND')
    def test_missing_nested_folder(self): self.denied(self.read(self.approved / 'missing/f.txt'), 'SG-NOT-FOUND')
    def test_wrong_folder(self): self.denied(self.read(self.root / 'elsewhere.txt'), 'SG-FOLDER')
    def test_sibling_prefix(self): self.denied(self.read(self.root / 'Approved-Data-Other/f.txt'), 'SG-FOLDER')
    def test_traversal(self): self.denied(self.read(self.approved / '../elsewhere.txt'), 'SG-FOLDER')
    def test_docx_hidden_extension(self): self.denied(self.read(self.file('policy.txt.docx')), 'SG-TYPE')
    def test_oversized(self): self.denied(self.read(self.file(data=b'A'*1048577)), 'SG-SIZE')
    def test_directory(self):
        p=self.approved/'folder.txt';p.mkdir()
        self.denied(self.read(p),'SG-DIRECTORY')
    def test_invalid_utf8(self): self.denied(self.read(self.file(data=b'Meal: caf\xe9')), 'SG-ENCODING')
    def test_utf16(self): self.denied(self.read(self.file(data='Meal costs'.encode('utf-16'))), 'SG-ENCODING')
    def test_binary_nul(self): self.denied(self.read(self.file(data=b'A\x00B')), 'SG-ENCODING')
    def test_sensitive_data_never_echoed(self):
        result=self.read(self.file(data=b'PRIVATE_TEST_SENTINEL person@example.org'))
        self.denied(result,'SG-SENSITIVE')
        self.assertNotIn('PRIVATE_TEST_SENTINEL',result.stdout+result.stderr)
        self.assertNotIn('person@example.org',result.stdout+result.stderr)
    def test_policy_label_still_blocked(self): self.denied(self.read(self.file(data=b'Confidential policy')), 'SG-SENSITIVE')
    def test_missing_policy(self):
        (self.root/'CLAUDE.md').unlink()
        self.denied(self.read(self.file()),'SG-POLICY')
    @unittest.skipUnless(os.name=='posix','Symlink fixture needs Unix')
    def test_symlink(self):
        target=self.root/'original.txt';target.write_text('test')
        p=self.approved/'link.txt';p.symlink_to(target)
        self.denied(self.read(p),'SG-LINK')
    def test_shell_blocked(self):
        self.denied(self.run_guard({'hook_event_name':'PreToolUse','tool_name':'Bash'}),'SG-TOOL')
    def test_mcp_still_permitted(self):
        self.allowed(self.run_guard({'hook_event_name':'PreToolUse','tool_name':'mcp__claude_safeguard__spreadsheet'}))
    def test_prompt_block_suppresses_original(self):
        result=self.run_guard({'hook_event_name':'UserPromptSubmit','prompt':'person@example.org'})
        self.assertEqual(result.returncode,0)
        data=json.loads(result.stdout)
        self.assertEqual(data['decision'],'block')
        self.assertTrue(data['hookSpecificOutput']['suppressOriginalPrompt'])
        self.assertNotIn('person@example.org',result.stdout+result.stderr)
    def test_local_pass(self):
        result=self.run_guard(check_file=self.file())
        self.assertEqual(result.returncode,0,result.stderr)
        self.assertEqual(json.loads(result.stdout)['status'],'passed')
    def test_local_block_same_check(self):
        result=self.run_guard(check_file=self.approved/'missing.txt')
        self.assertEqual(result.returncode,2)
        self.assertEqual(json.loads(result.stdout)['code'],'SG-NOT-FOUND')
    def test_malformed_payload_has_safe_code(self):
        result=self.run_guard(raw='INVALID_SECRET_SENTINEL')
        self.assertEqual(result.returncode,2)
        self.assertIn('SG-CHECK-INPUT',result.stderr)
        self.assertNotIn('INVALID_SECRET_SENTINEL',result.stdout+result.stderr)

if __name__=='__main__': unittest.main(verbosity=2)
