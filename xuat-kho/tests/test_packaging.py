"""Release ZIP integrity and omission checks using tiny substitute assets."""
import hashlib
import json
import os
from pathlib import Path
import sys
import tempfile
import unittest
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import donggoi


def write_file(root, name, content=b'example\n'):
    path = root / name
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(content)
    return path


def entry(name, content):
    return {'file': name, 'bytes': len(content), 'sha256': hashlib.sha256(content).hexdigest()}


def create_project(root):
    for name in donggoi.ROOT_FILES:
        write_file(root, name)
    write_file(root, 'VERSION.txt', b'4.2.1\n')
    for name in donggoi.SAMPLES:
        write_file(root, f'samples/{name}')
    for name in ('static/index.html', 'templates/Template.xlsx', 'tests/test_example.py', 'api/index.py'):
        write_file(root, name)
    ocr, vendor = b'verified OCR\n', b'verified library\n'
    write_file(root, 'static/ocr/core/core.js', ocr)
    write_file(root, 'vendor/library/__init__.py', vendor)
    write_file(root, 'static/ocr/asset-manifest.json', json.dumps({'assets': [entry('core/core.js', ocr)]}).encode())
    write_file(root, 'vendor/manifest.json', json.dumps({'packages': [{'name': 'library', 'files': [entry('library/__init__.py', vendor)]}]}).encode())


class PackagingTests(unittest.TestCase):
    def test_release_excludes_credentials_previews_and_caches(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory) / 'project'
            create_project(root)
            excluded = (
                '.vercel/anonymous.json', 'Loi-khoi-dong.log', 'unlisted.txt',
                'static/.env', 'static/credentials.json', 'static/hidden.key',
                'static/.private/config.txt', 'static/secrets/token.txt',
                'vendor/library/__pycache__/library.pyc', 'api/server.log',
                'samples/giao-dien-linh-v4.2.png', 'samples/anh-mau-5-khach-v4.2.png',
            )
            for name in excluded:
                write_file(root, name, b'PRIVATE OR GENERATED')
            output = donggoi.build_package(root)
            self.assertEqual(output.name, 'Du-an-xuat-kho-v4.2.1-gon.zip')
            with ZipFile(output) as archive:
                names = {name.removeprefix('Xuat-kho-v4.2.1/') for name in archive.namelist()}
                self.assertTrue(set(donggoi.ROOT_FILES).issubset(names))
                self.assertIn('static/ocr/core/core.js', names)
                self.assertIn('vendor/library/__init__.py', names)
                self.assertIn('samples/anh-nhieu-khach-hang.jpg', names)
                self.assertTrue(names.isdisjoint(excluded))
                self.assertIsNone(archive.testzip())

    def test_identical_sources_produce_identical_archives(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory) / 'project'
            create_project(root)
            first = donggoi.build_package(root, Path(directory) / 'first.zip').read_bytes()
            os.utime(root / 'app.py', (1000000000, 1000000000))
            second = donggoi.build_package(root, Path(directory) / 'second.zip').read_bytes()
            self.assertEqual(first, second)

    def test_missing_or_tampered_dependencies_preserve_existing_output(self):
        for asset in ('static/ocr/core/core.js', 'vendor/library/__init__.py'):
            for damage in ('missing', 'tampered'):
                with self.subTest(asset=asset, damage=damage), tempfile.TemporaryDirectory() as directory:
                    root = Path(directory) / 'project'
                    create_project(root)
                    output = write_file(Path(directory), 'previous.zip', b'KEEP PREVIOUS RELEASE')
                    if damage == 'missing':
                        (root / asset).unlink()
                    else:
                        # Same-length corruption must fail the SHA-256 check, too.
                        data = (root / asset).read_bytes()
                        (root / asset).write_bytes(b'X' + data[1:])
                    with self.assertRaisesRegex(ValueError, 'Thiếu file|bị thay đổi'):
                        donggoi.build_package(root, output)
                    self.assertEqual(output.read_bytes(), b'KEEP PREVIOUS RELEASE')

    def test_manifest_cannot_read_outside_its_asset_directory(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory) / 'project'
            create_project(root)
            manifest = {'assets': [entry('../private.txt', b'PRIVATE')]}
            write_file(root, 'static/private.txt', b'PRIVATE')
            write_file(root, 'static/ocr/asset-manifest.json', json.dumps(manifest).encode())
            with self.assertRaisesRegex(ValueError, 'Đường dẫn manifest'):
                donggoi.build_package(root)

    def test_symlink_cannot_package_an_external_file(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory) / 'project'
            create_project(root)
            external = write_file(Path(directory), 'external.txt', b'PRIVATE')
            try:
                (root / 'static/external.txt').symlink_to(external)
            except (OSError, NotImplementedError):
                self.skipTest('Creating symlinks is unavailable on this platform')
            with self.assertRaisesRegex(ValueError, 'liên kết tượng trưng'):
                donggoi.build_package(root)


if __name__ == '__main__':
    unittest.main()
