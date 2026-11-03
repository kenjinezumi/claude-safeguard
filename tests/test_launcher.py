"""Launcher regression tests with PowerShell 7 on Linux, no real Claude/data.
Run: SAFEGUARD_TEST_PWSH=/path/to/pwsh python3 tests/test_launcher.py
Executable fixtures use Unix shebangs; Windows installer/COM/MCP are not tested.
Repair/launch tests mock dependency, configuration and shortcut helpers.
Real configuration and MCP runtime are exercised by separate tests.
"""
import json
import os
from pathlib import Path
import shutil
import subprocess
import tempfile
import unittest

SOURCE = Path(__file__).resolve().parents[1]
PWSH = os.environ.get('SAFEGUARD_TEST_PWSH') or shutil.which('pwsh')

def quote(value):
    return "'" + str(value).replace("'", "''") + "'"

@unittest.skipUnless(PWSH and os.name == 'posix', 'Requires PowerShell 7 on Linux')
class LauncherTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix='safeguard test ')
        self.root = Path(self.temp.name)
        self.profile = self.root / "User Name (test)"
        self.local = self.root / 'Local App Data'
        self.install = self.local / 'claude-safeguard'
        self.package = self.root / 'extracted package'
        shutil.copytree(SOURCE, self.package, ignore=shutil.ignore_patterns('tests', '__pycache__'))
        self.install.mkdir(parents=True)
        self.profile.mkdir()
        self.env = dict(os.environ, USERPROFILE=str(self.profile), LOCALAPPDATA=str(self.local),
                        PATH='/usr/bin:/bin', SG_CAPTURE=str(self.root / 'capture.txt'))

    def tearDown(self):
        self.temp.cleanup()

    def ps(self, command):
        return subprocess.run([PWSH, '-NoProfile', '-NonInteractive', '-Command',
                               "$ErrorActionPreference='Stop'; " + command],
                              env=self.env, capture_output=True, text=True, timeout=30)

    def script(self, path):
        return self.ps('& ' + quote(path))

    def executable(self, path=None, code=0, version="2.1.295"):
        path = path or self.profile / '.local/bin/claude.exe'
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text('#!/bin/sh\n'
                        f'if [ "$1" = "--version" ]; then echo "{version} (Claude Code)"; '
                        f'exit {code}; fi\n'
                        'printf "%s\\n" "$PWD" "$@" > "$SG_CAPTURE"\n'
                        f'exit {code}\n')
        path.chmod(0o755)
        return path

    def installation(self, old=False):
        files = ['guard.ps1', 'CLAUDE.md', '.claude/settings.json', '.mcp.json',
                 'runtime/Scripts/python.exe', 'spreadsheet.py', 'finance.py',
                 'folder_access.py', 'runtime_check.py', 'configuration.py', 'requirements.txt', 'Prepare.ps1', 'SelectFolder.ps1', 'ImportWorkbook.ps1',
                 'Workbooks/keep.xlsx', 'Approved-Data/keep.csv', 'approved-folder.json']
        for name in files:
            p = self.install / name
            p.parent.mkdir(parents=True, exist_ok=True)
            p.write_text('synthetic preserved ' + name)
        if old:
            (self.install / 'Launch.ps1').write_text('# old launcher\n')
        else:
            for name in ['Launch.ps1', 'ClaudeRuntime.ps1']:
                shutil.copy2(SOURCE / name, self.install / name)
            with (self.install / 'ClaudeRuntime.ps1').open('a') as out:
                out.write('\nfunction Assert-SafeguardRuntime { param($Destination) }\nfunction Assert-SafeguardConfiguration { param($Destination) }\n')

    def preserved(self):
        return {str(p.relative_to(self.install)): p.read_bytes()
                for p in self.install.rglob('*') if p.is_file()
                and p.name not in ('Launch.ps1', 'ClaudeRuntime.ps1', 'START-HERE.txt', 'README.txt', 'VALIDATION.txt', 'guard.ps1', 'Prepare.ps1', 'FileCheck.ps1', 'CheckFile.ps1', 'CLAUDE.md', 'spreadsheet.py', 'sharepoint.py', 'ConnectSharePoint.ps1', 'SHAREPOINT-SETUP.txt', 'finance.py', 'folder_access.py', 'convert.py', 'ImportWorkbook.ps1', 'requirements.txt', 'requirements-sharepoint.txt', 'runtime_check.py', 'configuration.py')
                and 'Launcher-Backups' not in p.parts}

    def mock_shortcut(self, failure=False):
        p = self.package / 'ClaudeRuntime.ps1'
        body = "throw 'synthetic shortcut failure'" if failure else "'shortcut updated' | Out-Null"
        with p.open('a') as out:
            out.write('\nfunction Install-SafeguardDependencies { param($Destination) }\nfunction Update-SafeguardConfiguration { param($Destination) }\n')
            out.write('\nfunction Set-SafeguardLaunchShortcut { param([string]$Destination) ' + body + ' }\n')

    def test_all_powershell_parses(self):
        result = self.ps("$errorsFound=0; Get-ChildItem " + quote(SOURCE) + " -Filter *.ps1 | ForEach-Object { "
                         "$tokens=$null; $errors=$null; [void][System.Management.Automation.Language.Parser]::ParseFile($_.FullName,[ref]$tokens,[ref]$errors); "
                         "if ($errors) { $errors | Out-String | Write-Output; $errorsFound++ } }; if ($errorsFound) { exit 1 }")
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)

    def test_native_found_without_path_and_preferred_over_alias(self):
        binary = self.executable()
        result = self.ps('. ' + quote(SOURCE / 'ClaudeRuntime.ps1') +
                         "; function claude { throw 'alias invoked' }; Find-SafeguardClaude")
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(result.stdout.strip(), str(binary))

    def test_path_fallback(self):
        binary = self.executable(self.root / 'Custom Tools/claude.exe')
        self.env['PATH'] = str(binary.parent) + os.pathsep + self.env['PATH']
        result = self.ps('. ' + quote(SOURCE / 'ClaudeRuntime.ps1') + '; Find-SafeguardClaude')
        self.assertEqual(result.stdout.strip(), str(binary))

    def test_alias_not_treated_as_installation(self):
        result = self.ps('. ' + quote(SOURCE / 'ClaudeRuntime.ps1') +
                         "; function claude { 'fake' }; if (Find-SafeguardClaude) { exit 1 }")
        self.assertEqual(result.returncode, 0, result.stderr)

    def test_launch_cwd_arguments_and_spaces(self):
        self.executable()
        self.installation()
        result = self.script(self.install / 'Launch.ps1')
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertEqual(Path(self.env['SG_CAPTURE']).read_text().splitlines(),
                         [str(self.install), '--permission-mode', 'default'])

    def test_missing_binary_has_repair_message(self):
        self.installation()
        result = self.script(self.install / 'Launch.ps1')
        self.assertEqual(result.returncode, 1)
        self.assertIn('run Install.cmd', result.stdout)
        self.assertFalse(Path(self.env['SG_CAPTURE']).exists())

    def test_missing_guard_refuses_launch(self):
        self.executable()
        self.installation()
        (self.install / 'guard.ps1').unlink()
        result = self.script(self.install / 'Launch.ps1')
        self.assertEqual(result.returncode, 1)
        self.assertIn('Required installation file missing: guard.ps1', result.stdout)
        self.assertFalse(Path(self.env['SG_CAPTURE']).exists())

    def test_native_failure_reported(self):
        self.executable(code=9)
        self.installation()
        result = self.script(self.install / 'Launch.ps1')
        self.assertEqual(result.returncode, 1)
        self.assertIn('exit code 9', result.stdout)

    def test_bad_version_check_stops_setup(self):
        self.executable(code=7)
        result = self.ps('. ' + quote(SOURCE / 'ClaudeRuntime.ps1') + '; Install-SafeguardClaude')
        self.assertNotEqual(result.returncode, 0)
        self.assertIn('version check failed', result.stderr)

    def test_old_claude_version_refuses_launch(self):
        self.executable(version='2.1.294')
        self.installation()
        result = self.script(self.install / 'Launch.ps1')
        self.assertEqual(result.returncode, 1)
        self.assertIn('v2.1.295 or later is required', result.stdout)
        self.assertFalse(Path(self.env['SG_CAPTURE']).exists())

    def test_unknown_claude_version_refuses_launch(self):
        self.executable(version='unknown')
        self.installation()
        result = self.script(self.install / 'Launch.ps1')
        self.assertEqual(result.returncode, 1)
        self.assertIn('unrecognised version', result.stdout)
        self.assertFalse(Path(self.env['SG_CAPTURE']).exists())

    def test_repair_preserves_data_and_backs_up_launcher(self):
        self.executable()
        self.installation(old=True)
        self.mock_shortcut()
        before = self.preserved()
        for _ in range(2):
            result = self.script(self.package / 'Repair.ps1')
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            self.assertEqual(self.preserved(), before)
        self.assertEqual((self.install / 'Launch.ps1').read_bytes(), (SOURCE / 'Launch.ps1').read_bytes())
        backups = list((self.install / 'Launcher-Backups').glob('*/Launch.ps1'))
        self.assertEqual(len(backups), 2)
        self.assertTrue(any(p.read_text() == '# old launcher\n' for p in backups))

    def test_repair_rolls_back_when_shortcut_fails(self):
        self.executable()
        self.installation(old=True)
        self.mock_shortcut(failure=True)
        before = self.preserved()
        originals = {str(p.relative_to(self.install)):p.read_bytes() for p in self.install.rglob('*') if p.is_file()}
        result = self.script(self.package / 'Repair.ps1')
        self.assertEqual(result.returncode, 1)
        self.assertEqual(self.preserved(), before)
        for name, data in originals.items():
            self.assertEqual((self.install / name).read_bytes(), data)
        self.assertEqual((self.install / 'Launch.ps1').read_text(), '# old launcher\n')
        self.assertFalse((self.install / 'ClaudeRuntime.ps1').exists())

    def test_incomplete_installation_unchanged(self):
        self.installation(old=True)
        (self.install / '.mcp.json').unlink()
        before = self.preserved()
        result = self.script(self.package / 'Repair.ps1')
        self.assertEqual(result.returncode, 1)
        self.assertIn('incomplete', result.stdout)
        self.assertEqual(self.preserved(), before)
        self.assertFalse((self.install / 'Launcher-Backups').exists())

    def test_install_routes_existing_to_repair(self):
        self.executable()
        self.installation(old=True)
        self.mock_shortcut()
        before = self.preserved()
        result = self.script(self.package / 'Install.ps1')
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertIn('Package repaired', result.stdout)
        self.assertEqual(self.preserved(), before)

if __name__ == '__main__':
    unittest.main(verbosity=2)
