"""Start the local Excel app using one verified Python runtime."""
import errno
from http.client import HTTPException
import json
import ntpath
import os
from pathlib import Path
import platform
import re
import shutil
import subprocess
import sys
import sysconfig
import traceback
from urllib import request
import webbrowser

ROOT = Path(__file__).resolve().parent
APP_VERSION = '4.3.0'
PORTS = range(8080, 8090)
HEALTH_TIMEOUT = 0.35
RUNTIME_GUARD = 'XUAT_KHO_RUNTIME_SELECTED'


def is_free_threaded():
    return bool(sysconfig.get_config_var('Py_GIL_DISABLED'))


def launcher_paths(output):
    """Read executable paths from both Windows launcher list formats."""
    paths = []
    for line in output.splitlines():
        match = re.search(r'(?:[A-Za-z]:[\\/]|\\\\)[^\r\n]*?\.exe"?\s*$', line, re.IGNORECASE)
        if match:
            paths.append(match.group().strip().strip('"'))
    return paths


def windows_runtime_candidates():
    # A standard python.exe is often installed beside python3.13t.exe.
    candidates = [ntpath.join(ntpath.dirname(sys.executable), 'python.exe')]
    for arguments in (['py', '-0p'], ['py', '--list-paths']):
        try:
            result = subprocess.run(arguments, capture_output=True, text=True, errors='replace', timeout=3)
        except (OSError, subprocess.TimeoutExpired):
            continue
        if result.returncode == 0:
            detected = launcher_paths(result.stdout)
            candidates.extend(detected)
            if detected:
                break
    for name in ('python.exe', 'python3.exe', 'python', 'python3'):
        executable = shutil.which(name)
        if executable:
            candidates.append(executable)
    current = ntpath.normcase(ntpath.abspath(sys.executable))
    seen = {current}
    for candidate in candidates:
        key = ntpath.normcase(ntpath.abspath(candidate))
        if key not in seen:
            seen.add(key)
            yield candidate


def find_regular_windows_runtime():
    """Prefer an installed standard CPython; free-threaded remains a fallback."""
    probe = (
        'import json, platform, sys, sysconfig; '
        'print(json.dumps({"version": list(sys.version_info[:3]), '
        '"implementation": platform.python_implementation(), '
        '"free_threaded": bool(sysconfig.get_config_var("Py_GIL_DISABLED")), '
        '"executable": sys.executable}))'
    )
    for executable in windows_runtime_candidates():
        try:
            result = subprocess.run(
                [executable, '-c', probe], capture_output=True, text=True,
                errors='replace', timeout=3,
            )
            if result.returncode != 0:
                continue
            runtime = json.loads(result.stdout.strip())
            version = tuple(runtime['version'])
            resolved = runtime['executable']
            if (
                len(version) == 3 and all(isinstance(part, int) for part in version)
                and version >= (3, 10, 0)
                and runtime.get('implementation') == 'CPython'
                and runtime.get('free_threaded') is False
                and isinstance(resolved, str) and resolved
                and ntpath.normcase(ntpath.abspath(resolved)) != ntpath.normcase(ntpath.abspath(sys.executable))
            ):
                return resolved
        except (OSError, subprocess.TimeoutExpired, ValueError, TypeError, KeyError):
            continue
    return None


def use_preferred_runtime():
    if os.name != 'nt' or not is_free_threaded() or os.environ.get(RUNTIME_GUARD) == '1':
        return None
    executable = find_regular_windows_runtime()
    if not executable:
        print('Dùng Python free-threaded hiện tại với thư viện Excel đi kèm.', flush=True)
        return None
    print(f'Đang chuyển sang Python tiêu chuẩn đã cài: {executable}', flush=True)
    environment = os.environ.copy()
    environment[RUNTIME_GUARD] = '1'
    try:
        return subprocess.run(
            [executable, str(ROOT / 'khoidong.py')], cwd=str(ROOT), env=environment,
        ).returncode
    except OSError:
        print('Không mở được Python tiêu chuẩn; tiếp tục bằng Python hiện tại.', flush=True)
        return None


class _LocalHealthOnly(request.HTTPRedirectHandler):
    """A local probe must not follow a redirect to another service."""

    def redirect_request(self, req, fp, code, msg, headers, newurl):
        return None


def is_running_app(port):
    """Only reuse a server that identifies itself as this app."""
    url = f'http://127.0.0.1:{port}/api/health'
    # Ignore system proxy settings so localhost remains local on office PCs.
    opener = request.build_opener(request.ProxyHandler({}), _LocalHealthOnly())
    try:
        probe = request.Request(url, headers={'Accept': 'application/json'})
        with opener.open(probe, timeout=HEALTH_TIMEOUT) as response:
            if response.status != 200:
                return False
            body = response.read(8193)
        if len(body) > 8192:
            return False
        result = json.loads(body)
        return (
            isinstance(result, dict)
            and result.get('app') == 'xuat-kho'
            and result.get('status') == 'ok'
            and result.get('version') == APP_VERSION
        )
    except (OSError, ValueError, HTTPException):
        return False


def find_running_app():
    return next((port for port in PORTS if is_running_app(port)), None)


def open_running_app(port):
    url = f'http://127.0.0.1:{port}'
    print(f'Ứng dụng đã chạy ở cửa sổ khác: {url}', flush=True)
    print('Giữ cửa sổ đang chạy ứng dụng mở. Có thể đóng cửa sổ này.', flush=True)
    try:
        opened = webbrowser.open(url, new=2)
    except (OSError, webbrowser.Error):
        opened = False
    if not opened:
        print(f'Hãy mở trình duyệt và truy cập {url}', flush=True)


def ensure_dependencies():
    try:
        from dependencies import ensure_dependencies as check_dependencies
    except ImportError as exc:
        raise RuntimeError(
            'Thiếu file dependencies.py trong bộ cài. Hãy giải nén toàn bộ ZIP '
            'vào thư mục mới rồi mở Chay-Windows.bat; không chép riêng file khởi động.'
        ) from exc
    information = check_dependencies()
    origin = 'đi kèm bộ cài' if information['source'] == 'bundled' else 'đã cài trong Python'
    print(f"Thư viện Excel {origin}: openpyxl {information['openpyxl']} · et_xmlfile {information['et_xmlfile']}", flush=True)
    return information


def main():
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, 'reconfigure'):
            stream.reconfigure(encoding='utf-8', errors='replace')
    try:
        os.chdir(ROOT)
        switched = use_preferred_runtime()
        if switched is not None:
            return switched
        if sys.version_info < (3, 10):
            raise RuntimeError('Cần Python 3.10 trở lên. Hãy cài Python mới rồi chạy lại.')
        print('Đang kiểm tra ứng dụng...', flush=True)
        existing_port = find_running_app()
        if existing_port is not None:
            open_running_app(existing_port)
            return 0
        print(f'Python {sys.version.split()[0]} · Đang kiểm tra thư viện Excel...', flush=True)
        ensure_dependencies()
        from app import serve
        for port in PORTS:
            try:
                serve(port=port, open_browser=True)
                return 0
            except OSError as exc:
                occupied = getattr(exc, 'winerror', None) == 10048 or exc.errno in (errno.EADDRINUSE, 10048)
                if not occupied:
                    raise
                # Handle a second launch arriving after the first health scan.
                if is_running_app(port):
                    open_running_app(port)
                    return 0
                print(f'Cổng {port} đang được sử dụng. Đang thử cổng tiếp theo...', flush=True)
        raise RuntimeError('Các cổng 8080–8089 đều đang được sử dụng. Hãy đóng ứng dụng dùng các cổng này và chạy lại.')
    except KeyboardInterrupt:
        return 0
    except Exception:
        details = (
            f'Python: {sys.version.split()[0]}\nRuntime: {sys.executable}\n'
            f'Implementation: {platform.python_implementation()}\n'
            f'Free-threaded: {is_free_threaded()}\n'
            'Thư viện được đọc tại máy; trình khởi động không tự tải bằng pip.\n\n'
            f'{traceback.format_exc()}'
        )
        try:
            (ROOT / 'Loi-khoi-dong.log').write_text(details, encoding='utf-8')
        except OSError:
            pass
        print('\nKhông khởi chạy được ứng dụng:\n' + details, flush=True)
        print('Gửi nội dung trên hoặc file Loi-khoi-dong.log để kiểm tra.', flush=True)
        return 1


if __name__ == '__main__':
    raise SystemExit(main())
