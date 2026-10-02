"""Use the Excel libraries shipped with the application, without running pip."""
from __future__ import annotations

import importlib
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parent
PACKAGES = ('openpyxl', 'et_xmlfile')


def configure_vendor():
    """Prefer bundled libraries; a partial extraction must not pass silently."""
    vendor = ROOT / 'vendor'
    if not vendor.exists():
        return False
    missing = [name for name in PACKAGES if not (vendor / name / '__init__.py').is_file()]
    if missing:
        raise RuntimeError(
            'Bộ thư viện Excel đi kèm chưa đầy đủ: ' + ', '.join(missing) + '. '
            'Hãy giải nén toàn bộ bản ZIP mới vào một thư mục mới, giữ thư mục vendor '
            'cạnh khoidong.py rồi mở Chay-Windows.bat. Không cần chạy pip.'
        )
    path = str(vendor)
    if path not in sys.path:
        sys.path.insert(0, path)
    importlib.invalidate_caches()
    return True


def _version_tuple(value):
    parts = str(value).split('.')
    if len(parts) < 2:
        raise ValueError('Không đọc được phiên bản thư viện')
    return tuple(int(part) for part in parts[:3])


def ensure_dependencies():
    """Check imports and supported versions and provide a useful failure reason."""
    bundled = configure_vendor()
    result = {'source': 'bundled' if bundled else 'installed'}
    try:
        for name in PACKAGES:
            module = importlib.import_module(name)
            result[name] = module.__version__
        if not (3, 1, 5) <= _version_tuple(result['openpyxl']) < (4, 0, 0):
            raise ValueError(f"openpyxl {result['openpyxl']} không thuộc phiên bản hỗ trợ 3.1.5–3.x")
        if not (1, 1, 0) <= _version_tuple(result['et_xmlfile']) < (3, 0, 0):
            raise ValueError(f"et_xmlfile {result['et_xmlfile']} không thuộc phiên bản hỗ trợ 1.1–2.x")
    except Exception as exc:
        raise RuntimeError(
            f'Không nạp được thư viện Excel: {exc}\n'
            'Hãy tải bản ZIP mới và giải nén toàn bộ vào một thư mục mới, '
            'bao gồm thư mục vendor. Bản đầy đủ chạy không cần pip hoặc Internet.'
        ) from exc
    return result
