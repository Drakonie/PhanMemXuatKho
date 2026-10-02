"""Regression checks for startup without pip or installed Excel packages."""
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import types
import unittest
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import dependencies


class DependencyTests(unittest.TestCase):
    def test_bundled_excel_pipeline_without_site_packages_or_pip(self):
        script = '''
import io, json, sys
from pathlib import Path
sys.path.insert(0, sys.argv[1])
import dependencies
versions = dependencies.ensure_dependencies()
assert versions['source'] == 'bundled'
assert versions['openpyxl'] == '3.1.5'
assert versions['et_xmlfile'] == '2.0.0'
try:
    import pip
except ImportError:
    pass
else:
    raise AssertionError('The isolated test unexpectedly has pip')
from openpyxl import load_workbook
from engine import analyze, export
from image_import import build_image_workbook
root = Path(sys.argv[1])
source = (root / 'samples/data-mau.xlsx').read_bytes()
sheets = analyze(source)
result = export((root / 'templates/Template.xlsx').read_bytes(), sheets)
workbook = load_workbook(io.BytesIO(result))
assert len(workbook.worksheets) == 3
assert sum(sheet['H27'].value for sheet in workbook) == 1569000
assert 'A5:I5' in {str(merged) for merged in workbook.active.merged_cells.ranges}
source, report = build_image_workbook([{'filename': 'photo.png', 'edited_text': True,
    'text': 'Khách hàng: Test\\nNgày xuất: 01/10/2026\\nTên hàng | Số lượng | Đơn giá | Thành tiền\\nGạo | 2 | 10000 | 20000'}])
assert analyze(source)[0]['total'] == 20000
print(json.dumps(versions))
'''
        result = subprocess.run(
            [sys.executable, '-I', '-S', '-c', script, str(ROOT)],
            capture_output=True, text=True, timeout=20,
        )
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(json.loads(result.stdout)['openpyxl'], '3.1.5')

    def test_no_vendor_keeps_installed_package_fallback(self):
        with tempfile.TemporaryDirectory() as directory, patch.object(dependencies, 'ROOT', Path(directory)):
            self.assertFalse(dependencies.configure_vendor())

    def test_partial_extraction_has_actionable_error(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / 'vendor/openpyxl').mkdir(parents=True)
            (root / 'vendor/openpyxl/__init__.py').touch()
            with patch.object(dependencies, 'ROOT', root):
                with self.assertRaisesRegex(RuntimeError, 'et_xmlfile.*giải nén toàn bộ'):
                    dependencies.configure_vendor()

    def test_missing_import_preserves_actual_reason_without_subprocess(self):
        with patch.object(dependencies, 'configure_vendor', return_value=False), patch.object(dependencies.importlib, 'import_module', side_effect=ModuleNotFoundError('No module named openpyxl')), patch.object(subprocess, 'run') as command:
            with self.assertRaisesRegex(RuntimeError, 'No module named openpyxl') as error:
                dependencies.ensure_dependencies()
            self.assertIn('không cần pip', str(error.exception))
            command.assert_not_called()

    def test_incompatible_version_has_specific_reason(self):
        modules = {'openpyxl': types.SimpleNamespace(__version__='3.0.10'), 'et_xmlfile': types.SimpleNamespace(__version__='2.0.0')}
        with patch.object(dependencies, 'configure_vendor', return_value=False), patch.object(dependencies.importlib, 'import_module', side_effect=modules.__getitem__):
            with self.assertRaisesRegex(RuntimeError, 'openpyxl 3.0.10'):
                dependencies.ensure_dependencies()


if __name__ == '__main__':
    unittest.main()
