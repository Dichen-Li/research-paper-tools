"""可跨主机执行的平台分支测试，不等于 Windows 真机验证。"""
import base64
import importlib.util
from pathlib import Path, PureWindowsPath
import unittest

ROOT = Path(__file__).resolve().parents[1]


def load(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    result = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(result)
    return result


RENDER = load('windows_render', ROOT / 'scripts/render_office.py')
TEMPLATE = load('windows_template_test', ROOT / 'tests/test_template.py')


class WindowsCompatibilityTests(unittest.TestCase):
    def test_64bit_and_x86_console_and_gui_executables(self):
        candidates = RENDER.office_candidates('win32', {
            'ProgramW6432': r'C:\Program Files', 'ProgramFiles': r'C:\Program Files',
            'ProgramFiles(x86)': r'C:\Program Files (x86)'})
        self.assertEqual(candidates[:2], ['soffice.com', 'soffice.exe'])
        for base in (r'C:\Program Files', r'C:\Program Files (x86)'):
            for executable in ('soffice.com', 'soffice.exe'):
                self.assertIn(str(PureWindowsPath(base)/'LibreOffice'/'program'/executable), candidates)
        self.assertEqual(len(candidates), len(set(candidates)))
        self.assertFalse(any('/Applications/' in item for item in candidates))

    def test_windows_candidates_without_programfiles(self):
        self.assertIn('soffice.exe', RENDER.office_candidates('win32', {}))

    def test_unicode_programfiles_path_not_shell_escaped(self):
        candidates = RENDER.office_candidates('win32', {'ProgramFiles': r'D:\应用 tools & data'})
        self.assertIn(r'D:\应用 tools & data\LibreOffice\program\soffice.exe', candidates)

    def test_junction_arguments_are_literal_not_shell_substitutions(self):
        command = TEMPLATE.junction_command(r"C:\项目's & 100%\node_modules",
                                            r'D:\依赖 空格\node_modules', 'powershell.exe')
        self.assertNotIn('cmd.exe', command)
        self.assertNotIn('-ExecutionPolicy', command)
        script = base64.b64decode(command[-1]).decode('utf-16le')
        self.assertIn("'C:\\项目''s & 100%\\node_modules'", script)
        self.assertIn("'D:\\依赖 空格\\node_modules'", script)
        self.assertIn('-ItemType Junction', script)

    def test_windows_profile_uri_encodes_spaces_and_unicode(self):
        uri = PureWindowsPath(r'C:\用户 space\QA profile').as_uri()
        self.assertTrue(uri.startswith('file:///C:/'))
        self.assertNotIn(' ', uri)
        self.assertNotIn('用户', uri)


if __name__ == '__main__':
    unittest.main()
