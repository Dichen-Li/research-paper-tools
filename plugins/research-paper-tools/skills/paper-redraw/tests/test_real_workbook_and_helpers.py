"""真实工作簿、QA 对照和可选 LibreOffice 集成测试。"""
import importlib.util
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]


def module(name):
    spec = importlib.util.spec_from_file_location(name, ROOT / "scripts" / (name + ".py"))
    result = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(result)
    return result


class WorkbookAndHelpersTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name) / "图表 space"
        self.folder.mkdir()

    def workbook(self, strings=False):
        try:
            from openpyxl import Workbook
            from openpyxl.cell.rich_text import CellRichText, TextBlock
            from openpyxl.cell.text import InlineFont
        except ImportError:
            self.skipTest("openpyxl rich text not available")
        workbook = Workbook()
        sheet = workbook.active
        sheet['A1'] = CellRichText('x', TextBlock(InlineFont(vertAlign='subscript'), '2'))
        sheet['B1'] = CellRichText('R', TextBlock(InlineFont(vertAlign='superscript'), '2'))
        sheet['A2'] = '0.931' if strings else 0.931
        sheet['B2'] = 'T_max' if strings else 94.6
        for index in range(3, 30):
            sheet.cell(2, index, index)
        workbook.create_sheet('Second')['A1'] = 1
        target = self.folder / 'table.xlsx'
        workbook.save(target)
        return target

    def test_real_workbook_types_rich_text_and_warn_exit_code(self):
        checker = module('check_native_office')
        report = checker.inspect(self.workbook())
        self.assertEqual(report['status'], 'PASS')
        self.assertGreaterEqual(report['counts']['script_markers'], 2)
        target = self.workbook(strings=True)
        result = subprocess.run([sys.executable, str(ROOT / 'scripts/check_native_office.py'), str(target)],
                                capture_output=True, text=True, encoding='utf-8', timeout=15)
        self.assertEqual(result.returncode, 0)
        self.assertEqual(json.loads(result.stdout)['status'], 'WARN')

    def test_comparison_preserves_aspect_and_original_bytes(self):
        try:
            from PIL import Image
        except ImportError:
            self.skipTest('Pillow unavailable')
        compare = module('compare_images')
        left, right = self.folder/'left.png', self.folder/'right.png'
        Image.new('RGB', (200, 100), 'blue').save(left)
        Image.new('RGB', (100, 100), 'red').save(right)
        before = left.read_bytes(), right.read_bytes()
        pair, overlay = self.folder/'pair.png', self.folder/'overlay.png'
        compare.compare(left, right, pair, overlay, height=100)
        with Image.open(pair) as result:
            self.assertEqual(result.size, (360, 170))
        with Image.open(overlay) as result:
            self.assertEqual(result.size, (200, 100))
        self.assertEqual(before, (left.read_bytes(), right.read_bytes()))
        with self.assertRaises(ValueError):
            compare.compare(left, right, left)

    def test_unicode_json_report_with_non_utf8_default_stdout(self):
        target = self.workbook(strings=True)
        environment = {**os.environ, 'PYTHONIOENCODING': 'ascii'}
        result = subprocess.run([sys.executable, str(ROOT/'scripts/check_native_office.py'), str(target)],
                                env=environment, capture_output=True, text=True, encoding='utf-8', timeout=15)
        self.assertEqual(result.returncode, 0, result.stderr)
        report = json.loads(result.stdout)
        self.assertEqual(report['file'], str(target))
        self.assertEqual(report['status'], 'WARN')

    def test_failed_renderer_does_not_claim_preview_success(self):
        target = self.workbook()
        result = subprocess.run([sys.executable, str(ROOT/'scripts/render_office.py'), str(target),
                                 '--output-dir', str(self.folder/'failed'), '--soffice',
                                 str(self.folder/'not-installed')], capture_output=True, text=True, timeout=15)
        self.assertEqual(result.returncode, 1)
        self.assertFalse(json.loads(result.stdout)['ready'])

    def test_real_wide_workbook_render_records_engine_and_all_sheets(self):
        soffice, poppler = os.environ.get('PAPER_REDRAW_SOFFICE'), os.environ.get('PAPER_REDRAW_PDFTOPPM')
        if not soffice or not poppler:
            self.skipTest('Set actual LibreOffice and Poppler paths for integration render')
        renderer = module('render_office')
        report = renderer.render(self.workbook(), self.folder/'qa', 100, False, 60,
                                 soffice, poppler, round_number=2)
        self.assertEqual(len(report['pages']), 2)
        self.assertEqual(report['renderer']['name'], 'LibreOffice')
        self.assertTrue(report['renderer']['version'])
        self.assertTrue(report['blind_spots'])
        self.assertEqual(Path(report['outputDir']).name, 'r2')
        with self.assertRaises(RuntimeError):
            renderer.render(self.folder/'table.xlsx', self.folder/'qa', 100, False, 60,
                            soffice, poppler, round_number=2)


if __name__ == '__main__':
    unittest.main()
