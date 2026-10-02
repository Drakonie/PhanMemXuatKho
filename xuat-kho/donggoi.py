"""Build a small, reproducible offline ZIP using only the Python standard library."""
import argparse
import hashlib
import json
import os
from pathlib import Path, PurePosixPath
import re
import tempfile
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo


ROOT = Path(__file__).resolve().parent
ROOT_FILES = (
    'app.py', 'engine.py', 'image_import.py', 'image_tables.py', 'dependencies.py',
    'khoidong.py', 'Chay-Windows.bat', 'requirements.txt', 'README.md',
    'THAY-DOI-v4.2.md', 'THAY-DOI-v4.2.1.md', 'THAY-DOI-v4.2.2.md', 'THAY-DOI-v4.3.md', 'VERSION.txt', 'vercel.json',
    '.python-version', 'pyproject.toml', 'uv.lock', 'donggoi.py',
)
DIRECTORIES = ('static', 'templates', 'tests', 'api', 'vendor')
SAMPLES = (
    'data-mau.xlsx', 'data-da-dang.xlsx', 'Phieu-xuat-kho-mau.xlsx',
    'anh-minh-hoa.png', 'anh-nhieu-khach-hang.jpg',
)
EXCLUDED_NAMES = {
    '__pycache__', 'node_modules', 'venv', 'env', 'secrets', 'credentials',
    'credentials.json', 'secrets.json', 'secret.json', 'anonymous.json',
}
EXCLUDED_SUFFIXES = {'.pyc', '.pyo', '.log', '.tmp', '.key', '.pem'}


def _excluded(relative):
    return (
        any(part.startswith('.') or part.lower() in EXCLUDED_NAMES for part in relative.parts)
        or relative.suffix.lower() in EXCLUDED_SUFFIXES
    )


def _checked_file(root, relative):
    """Reject links and manifest paths that could read outside the project."""
    path = PurePosixPath(relative)
    if not relative or '\\' in relative or path.is_absolute() or any(
        part in ('', '.', '..') or ':' in part for part in path.parts
    ):
        raise ValueError(f'Đường dẫn không hợp lệ: {relative}')
    candidate = root
    for part in path.parts:
        candidate /= part
        if candidate.is_symlink():
            raise ValueError(f'Không đóng gói liên kết tượng trưng: {relative}')
    if not candidate.is_file():
        raise ValueError(f'Thiếu file bắt buộc: {relative}')
    if not candidate.resolve().is_relative_to(root):
        raise ValueError(f'File nằm ngoài dự án: {relative}')
    return candidate


def _read_manifest(root, relative):
    try:
        manifest = json.loads(_checked_file(root, relative).read_text(encoding='utf-8'))
        if not isinstance(manifest, dict):
            raise ValueError('manifest phải là đối tượng JSON')
        return manifest
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise ValueError(f'Không đọc được {relative}: {error}') from error


def _verify_entries(root, directory, entries):
    if not isinstance(entries, list) or not entries:
        raise ValueError(f'Manifest {directory} không có danh sách file hợp lệ')
    verified = set()
    for entry in entries:
        if not isinstance(entry, dict) or not isinstance(entry.get('file'), str):
            raise ValueError(f'Manifest {directory} có mục file không hợp lệ')
        name = entry['file']
        # Check the entry itself before prepending its allowed directory.
        pure = PurePosixPath(name)
        if pure.is_absolute() or '..' in pure.parts or '\\' in name:
            raise ValueError(f'Đường dẫn manifest không hợp lệ: {name}')
        relative = f'{directory}/{name}'
        source = _checked_file(root, relative)
        expected_size, expected_hash = entry.get('bytes'), entry.get('sha256')
        if type(expected_size) is not int or not isinstance(expected_hash, str) or not re.fullmatch(r'[0-9a-f]{64}', expected_hash):
            raise ValueError(f'Manifest thiếu kích thước/hash hợp lệ: {relative}')
        digest = hashlib.sha256()
        with source.open('rb') as stream:
            for chunk in iter(lambda: stream.read(1024 * 1024), b''):
                digest.update(chunk)
        if source.stat().st_size != expected_size or digest.hexdigest() != expected_hash:
            raise ValueError(f'File bị thay đổi hoặc hỏng: {relative}')
        verified.add(relative)
    return verified


def validate_manifests(root):
    """Verify every bundled OCR and Excel dependency before creating a release."""
    root = Path(root).resolve()
    ocr = _read_manifest(root, 'static/ocr/asset-manifest.json')
    verified = _verify_entries(root, 'static/ocr', ocr.get('assets'))
    vendor = _read_manifest(root, 'vendor/manifest.json')
    packages = vendor.get('packages')
    if not isinstance(packages, list) or not packages:
        raise ValueError('Manifest vendor không có danh sách thư viện hợp lệ')
    for package in packages:
        if not isinstance(package, dict):
            raise ValueError('Manifest vendor có thư viện không hợp lệ')
        verified.update(_verify_entries(root, 'vendor', package.get('files')))
    return verified


def package_files(root):
    """Use explicit roots so credentials and generated screenshots stay outside ZIP."""
    root = Path(root).resolve()
    verified = validate_manifests(root)
    names = set(ROOT_FILES)
    names.update(f'samples/{name}' for name in SAMPLES)
    for directory in DIRECTORIES:
        folder = root / directory
        if folder.is_symlink() or not folder.is_dir():
            raise ValueError(f'Thiếu thư mục hoặc thư mục là liên kết: {directory}')
        for path in folder.rglob('*'):
            relative = path.relative_to(root)
            if _excluded(relative):
                continue
            if path.is_symlink():
                raise ValueError(f'Không đóng gói liên kết tượng trưng: {relative}')
            if path.is_file():
                names.add(relative.as_posix())
    if not verified.issubset(names):
        raise ValueError('Có tài nguyên OCR/thư viện bắt buộc bị loại khỏi gói')
    return [(name, _checked_file(root, name)) for name in sorted(names)]


def build_package(root=ROOT, output=None):
    root = Path(root).resolve()
    version = _checked_file(root, 'VERSION.txt').read_text(encoding='utf-8').strip()
    if not re.fullmatch(r'\d+\.\d+(?:\.\d+)?', version):
        raise ValueError('VERSION.txt phải có dạng 4.2.1')
    files = package_files(root)
    output = Path(output).expanduser().resolve() if output else root.parent / f'Du-an-xuat-kho-v{version}-gon.zip'
    if output.suffix.lower() != '.zip':
        raise ValueError('File đầu ra phải có phần mở rộng .zip')
    if output.is_relative_to(root):
        relative = output.relative_to(root)
        if relative.parts[0] in DIRECTORIES or relative.as_posix() in dict(files):
            raise ValueError('Đặt ZIP đầu ra ngoài các thư mục nguồn được đóng gói')
    output.parent.mkdir(parents=True, exist_ok=True)
    temporary = None
    try:
        with tempfile.NamedTemporaryFile(prefix=f'.{output.name}.', suffix='.tmp', dir=output.parent, delete=False) as stream:
            temporary = Path(stream.name)
        with ZipFile(temporary, 'w', compression=ZIP_DEFLATED, compresslevel=9) as archive:
            for relative, source in files:
                info = ZipInfo(f'Xuat-kho-v{version}/{relative}', date_time=(1980, 1, 1, 0, 0, 0))
                info.create_system = 3
                info.external_attr = 0o100644 << 16
                archive.writestr(info, source.read_bytes(), compress_type=ZIP_DEFLATED, compresslevel=9)
        with ZipFile(temporary) as archive:
            damaged = archive.testzip()
            if damaged:
                raise ValueError(f'ZIP không vượt qua kiểm tra CRC: {damaged}')
        os.replace(temporary, output)
        return output
    finally:
        if temporary and temporary.exists():
            temporary.unlink()


def main():
    parser = argparse.ArgumentParser(description='Đóng gói bản gọn; giữ OCR và thư viện Excel chạy offline.')
    parser.add_argument('--output', type=Path, help='Đường dẫn ZIP đầu ra (mặc định ở thư mục cha của dự án)')
    arguments = parser.parse_args()
    try:
        output = build_package(output=arguments.output)
    except (OSError, ValueError) as error:
        parser.exit(1, f'Không thể đóng gói: {error}\n')
    print(f'Đã tạo: {output}\nDung lượng: {output.stat().st_size:,} bytes')


if __name__ == '__main__':
    main()
