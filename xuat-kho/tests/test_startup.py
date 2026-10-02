import contextlib
from http.client import BadStatusLine, IncompleteRead
import io
import json
import socketserver
import subprocess
import sys
import tempfile
import threading
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
import app
import khoidong

class StartupTests(unittest.TestCase):
    def setUp(self):
        self.context = contextlib.ExitStack()
        self.addCleanup(self.context.close)
        self.find_existing = self.context.enter_context(patch.object(khoidong, 'find_running_app', return_value=None))
        self.probe = self.context.enter_context(patch.object(khoidong, 'is_running_app', return_value=False))
        self.dependencies = self.context.enter_context(patch.object(khoidong, 'ensure_dependencies'))
        self.output = self.context.enter_context(contextlib.redirect_stdout(io.StringIO()))

    def test_port_in_use_tries_next_port(self):
        with patch.object(app,'serve',side_effect=[OSError(98,'in use'),None]) as server:
            self.assertEqual(khoidong.main(),0)
        self.assertEqual(server.call_args_list[0].kwargs,{'port':8080,'open_browser':True})
        self.assertEqual(server.call_args_list[1].kwargs,{'port':8081,'open_browser':True})
    def test_failed_binding_does_not_open_browser(self):
        with patch.object(app,'ThreadingHTTPServer',side_effect=OSError('bind failed')),patch('webbrowser.open') as browser:
            with self.assertRaises(OSError): app.serve(open_browser=True)
            browser.assert_not_called()

    def test_existing_instance_opens_without_dependencies_or_new_server(self):
        self.find_existing.return_value = 8084
        with patch.object(app, 'serve') as server, patch.object(khoidong.webbrowser, 'open', return_value=True) as browser:
            self.assertEqual(khoidong.main(), 0)
        self.dependencies.assert_not_called()
        server.assert_not_called()
        browser.assert_called_once_with('http://127.0.0.1:8084', new=2)

    def test_instance_started_during_scan_is_reused(self):
        self.probe.return_value = True
        with patch.object(app, 'serve', side_effect=OSError(98, 'in use')) as server, patch.object(khoidong.webbrowser, 'open', return_value=True) as browser:
            self.assertEqual(khoidong.main(), 0)
        server.assert_called_once_with(port=8080, open_browser=True)
        browser.assert_called_once_with('http://127.0.0.1:8080', new=2)

    def test_windows_address_in_use_tries_next_port(self):
        occupied = OSError('address in use')
        occupied.winerror = 10048
        with patch.object(app, 'serve', side_effect=[occupied, None]) as server:
            self.assertEqual(khoidong.main(), 0)
        self.assertEqual(server.call_args_list[-1].kwargs['port'], 8081)

    def test_unrelated_binding_failure_records_diagnostic(self):
        with tempfile.TemporaryDirectory() as directory, patch.object(khoidong, 'ROOT', Path(directory)), patch.object(khoidong.os, 'chdir'), patch.object(app, 'serve', side_effect=OSError(13, 'permission denied')) as server:
            self.assertEqual(khoidong.main(), 1)
            contents = (Path(directory) / 'Loi-khoi-dong.log').read_text(encoding='utf-8')
        server.assert_called_once()
        self.assertIn('permission denied', contents)
        self.assertIn(sys.executable, contents)

    def test_all_occupied_ports_report_failure(self):
        with tempfile.TemporaryDirectory() as directory, patch.object(khoidong, 'ROOT', Path(directory)), patch.object(khoidong.os, 'chdir'), patch.object(app, 'serve', side_effect=OSError(98, 'in use')) as server:
            self.assertEqual(khoidong.main(), 1)
            self.assertTrue((Path(directory) / 'Loi-khoi-dong.log').exists())
        self.assertEqual(server.call_count, 10)

    def test_browser_failure_still_shows_existing_url(self):
        self.find_existing.return_value = 8083
        with patch.object(khoidong.webbrowser, 'open', side_effect=OSError('no browser')):
            self.assertEqual(khoidong.main(), 0)
        self.assertIn('http://127.0.0.1:8083', self.output.getvalue())


class HealthProbeTests(unittest.TestCase):
    def response(self, body, status=200):
        response = MagicMock()
        response.__enter__.return_value = response
        response.status = status
        response.read.return_value = body
        return response

    def test_only_correct_health_marker_is_reused(self):
        examples = [
            ({'app': 'xuat-kho', 'status': 'ok', 'version': khoidong.APP_VERSION}, True),
            ({'app': 'xuat-kho', 'status': 'ok', 'version': '2.0.0'}, False),
            ({'app': 'xuat-kho', 'status': 'ok'}, False),
            ({'app': 'other-service', 'status': 'ok'}, False),
            ({'app': 'xuat-kho', 'status': 'failed'}, False),
            ({'status': 'ok'}, False),
            (['xuat-kho'], False),
        ]
        for marker, expected in examples:
            with self.subTest(marker=marker), patch.object(khoidong.request, 'build_opener') as opener:
                opener.return_value.open.return_value = self.response(json.dumps(marker).encode())
                self.assertEqual(khoidong.is_running_app(8080), expected)
                call = opener.return_value.open.call_args
                self.assertEqual(call.args[0].full_url, 'http://127.0.0.1:8080/api/health')
                self.assertEqual(call.kwargs['timeout'], khoidong.HEALTH_TIMEOUT)
                self.assertEqual(opener.call_args.args[0].proxies, {})

    def test_timeout_malformed_and_large_response_are_ignored(self):
        for body, status in [(b'not JSON', 200), (b'x' * 8193, 200), (b'{}', 404)]:
            with self.subTest(status=status, length=len(body)), patch.object(khoidong.request, 'build_opener') as opener:
                opener.return_value.open.return_value = self.response(body, status)
                self.assertFalse(khoidong.is_running_app(8080))
        with patch.object(khoidong.request, 'build_opener') as opener:
            opener.return_value.open.side_effect = TimeoutError('no response')
            self.assertFalse(khoidong.is_running_app(8080))

    def test_non_http_and_truncated_chunked_services_are_ignored(self):
        responses = [
            b'WELCOME TO OTHER SERVICE\r\n',
            b'HTTP/1.1 200 OK\r\nTransfer-Encoding: chunked\r\n\r\n4\r\n{',
        ]
        for response in responses:
            with self.subTest(response=response):
                class OtherService(socketserver.BaseRequestHandler):
                    def handle(self):
                        self.request.settimeout(1)
                        self.request.recv(4096)
                        self.request.sendall(response)

                with socketserver.TCPServer(('127.0.0.1', 0), OtherService) as server:
                    server.timeout = 1
                    worker = threading.Thread(target=server.handle_request, daemon=True)
                    worker.start()
                    try:
                        self.assertFalse(khoidong.is_running_app(server.server_address[1]))
                    finally:
                        worker.join(2)
                    self.assertFalse(worker.is_alive())

    def test_protocol_errors_do_not_prevent_starting_on_next_port(self):
        for failure in (BadStatusLine('other-service'), IncompleteRead(b'')):
            with self.subTest(failure=type(failure).__name__), tempfile.TemporaryDirectory() as directory:
                with patch.object(khoidong, 'ROOT', Path(directory)), patch.object(khoidong.os, 'chdir'), patch.object(khoidong, 'PORTS', [8080, 8081]), patch.object(khoidong.request, 'build_opener') as opener, patch.object(khoidong, 'ensure_dependencies'), patch.object(app, 'serve', side_effect=[OSError(98, 'in use'), None]) as server, contextlib.redirect_stdout(io.StringIO()):
                    opener.return_value.open.side_effect = failure
                    self.assertEqual(khoidong.main(), 0)
                self.assertEqual([call.kwargs['port'] for call in server.call_args_list], [8080, 8081])
                self.assertFalse((Path(directory) / 'Loi-khoi-dong.log').exists())

    def test_probe_does_not_follow_redirects(self):
        self.assertIsNone(khoidong._LocalHealthOnly().redirect_request(None, None, 302, None, None, 'https://example.org'))

    def test_scan_stops_at_first_recognized_instance(self):
        with patch.object(khoidong, 'is_running_app', side_effect=[False, False, True]) as probe:
            self.assertEqual(khoidong.find_running_app(), 8082)
        self.assertEqual([call.args[0] for call in probe.call_args_list], [8080, 8081, 8082])


class OfflineDependencyTests(unittest.TestCase):
    def test_complete_bundle_does_not_execute_pip(self):
        helper = MagicMock()
        helper.ensure_dependencies.return_value = {
            'openpyxl': '3.1.5', 'et_xmlfile': '2.0.0', 'source': 'bundled',
        }
        with patch.dict(sys.modules, {'dependencies': helper}), patch.object(khoidong.subprocess, 'run') as process, contextlib.redirect_stdout(io.StringIO()) as output:
            information = khoidong.ensure_dependencies()
        helper.ensure_dependencies.assert_called_once_with()
        process.assert_not_called()
        self.assertEqual(information['source'], 'bundled')
        self.assertIn('đi kèm bộ cài', output.getvalue())

    def test_missing_bootstrap_explains_full_extraction_without_installer(self):
        with patch.dict(sys.modules, {'dependencies': None}), patch.object(khoidong.subprocess, 'run') as process:
            with self.assertRaisesRegex(RuntimeError, 'giải nén toàn bộ ZIP'):
                khoidong.ensure_dependencies()
        process.assert_not_called()

    def test_dependency_diagnostic_is_preserved_in_error_log(self):
        helper = MagicMock()
        helper.ensure_dependencies.side_effect = RuntimeError('Thiếu vendor/openpyxl; hãy giải nén đầy đủ bộ cài.')
        with tempfile.TemporaryDirectory() as directory:
            with patch.dict(sys.modules, {'dependencies': helper}), patch.object(khoidong, 'ROOT', Path(directory)), patch.object(khoidong.os, 'chdir'), patch.object(khoidong, 'use_preferred_runtime', return_value=None), patch.object(khoidong, 'find_running_app', return_value=None), patch.object(khoidong.subprocess, 'run') as process, patch.object(app, 'serve') as server, contextlib.redirect_stdout(io.StringIO()):
                self.assertEqual(khoidong.main(), 1)
                contents = (Path(directory) / 'Loi-khoi-dong.log').read_text(encoding='utf-8')
            self.assertIn('Thiếu vendor/openpyxl', contents)
            self.assertIn('không tự tải bằng pip', contents)
            self.assertNotIn('CalledProcessError', contents)
            process.assert_not_called()
            server.assert_not_called()


class WindowsRuntimeTests(unittest.TestCase):
    def runtime(self, executable, version=(3, 13, 6), free_threaded=False, implementation='CPython'):
        return subprocess.CompletedProcess([], 0, json.dumps({
            'version': list(version), 'implementation': implementation,
            'free_threaded': free_threaded, 'executable': executable,
        }), '')

    def test_launcher_paths_support_spaces_and_unc_paths(self):
        listing = (
            ' -V:3.13t * D:\\Users\\Linh\\Python313\\python3.13t.exe\n'
            ' -V:3.12 D:\\Program Files\\Python312\\python.exe\n'
            ' -V:3.11 "\\\\server\\python folder\\python.exe"\n'
            'No installed runtimes\n'
        )
        self.assertEqual(khoidong.launcher_paths(listing), [
            r'D:\Users\Linh\Python313\python3.13t.exe',
            r'D:\Program Files\Python312\python.exe',
            r'\\server\python folder\python.exe',
        ])

    def test_runtime_candidates_prefer_sibling_and_skip_current_duplicates(self):
        current = r'D:\Users\Linh\Python313\python3.13t.exe'
        regular = r'D:\Users\Linh\Python313\python.exe'
        older = r'C:\Program Files\Python312\python.exe'
        listing = f' -V:3.13t * {current}\n -V:3.13 {regular}\n -V:3.12 {older}\n'
        with patch.object(khoidong.sys, 'executable', current), patch.object(khoidong.subprocess, 'run', return_value=subprocess.CompletedProcess([], 0, listing, '')), patch.object(khoidong.shutil, 'which', return_value=regular):
            self.assertEqual(list(khoidong.windows_runtime_candidates()), [regular, older])

    def test_alternate_launcher_listing_is_checked_after_no_paths(self):
        responses = [
            subprocess.CompletedProcess([], 0, 'No runtimes found by old command', ''),
            subprocess.CompletedProcess([], 0, r' -V:3.12 C:\Program Files\Python312\python.exe', ''),
        ]
        with patch.object(khoidong.sys, 'executable', r'D:\Python313\python3.13t.exe'), patch.object(khoidong.subprocess, 'run', side_effect=responses) as process, patch.object(khoidong.shutil, 'which', return_value=None):
            candidates = list(khoidong.windows_runtime_candidates())
        self.assertIn(r'C:\Program Files\Python312\python.exe', candidates)
        self.assertEqual(process.call_args_list[-1].args[0], ['py', '--list-paths'])

    def test_rejects_old_other_implementation_and_free_threaded_candidates(self):
        candidates = [r'C:\old.exe', r'C:\pypy.exe', r'C:\python3.13t.exe', r'C:\Python313\python.exe']
        responses = [
            self.runtime(candidates[0], (3, 9, 15)),
            self.runtime(candidates[1], implementation='PyPy'),
            self.runtime(candidates[2], free_threaded=True),
            self.runtime(candidates[3]),
        ]
        with patch.object(khoidong, 'windows_runtime_candidates', return_value=iter(candidates)), patch.object(khoidong.subprocess, 'run', side_effect=responses) as process:
            self.assertEqual(khoidong.find_regular_windows_runtime(), candidates[3])
        self.assertEqual(process.call_count, 4)
        self.assertTrue(all(call.kwargs.get('timeout') == 3 for call in process.call_args_list))

    def test_failed_runtime_probes_do_not_abort_selection(self):
        candidates = [r'C:\missing.exe', r'C:\slow.exe', r'C:\broken.exe', r'C:\regular.exe']
        failures = [OSError('not found'), subprocess.TimeoutExpired('probe', 3), subprocess.CompletedProcess([], 0, 'bad output', ''), self.runtime(candidates[-1])]
        with patch.object(khoidong, 'windows_runtime_candidates', return_value=iter(candidates)), patch.object(khoidong.subprocess, 'run', side_effect=failures):
            self.assertEqual(khoidong.find_regular_windows_runtime(), candidates[-1])

    def test_switch_preserves_environment_and_runs_once(self):
        executable = r'C:\Program Files\Python313\python.exe'
        with patch.object(khoidong.os, 'name', 'nt'), patch.object(khoidong, 'is_free_threaded', return_value=True), patch.dict(khoidong.os.environ, {'XUAT_KHO_TEST_MARKER': 'preserved'}, clear=True), patch.object(khoidong, 'find_regular_windows_runtime', return_value=executable), patch.object(khoidong.subprocess, 'run', return_value=subprocess.CompletedProcess([], 0)) as process, contextlib.redirect_stdout(io.StringIO()):
            self.assertEqual(khoidong.use_preferred_runtime(), 0)
            self.assertNotIn(khoidong.RUNTIME_GUARD, khoidong.os.environ)
        self.assertEqual(process.call_args.args[0], [executable, str(khoidong.ROOT / 'khoidong.py')])
        self.assertEqual(process.call_args.kwargs['env'][khoidong.RUNTIME_GUARD], '1')
        self.assertEqual(process.call_args.kwargs['env']['XUAT_KHO_TEST_MARKER'], 'preserved')

    def test_child_exit_status_is_preserved(self):
        with patch.object(khoidong.os, 'name', 'nt'), patch.object(khoidong, 'is_free_threaded', return_value=True), patch.dict(khoidong.os.environ, {}, clear=True), patch.object(khoidong, 'find_regular_windows_runtime', return_value=r'C:\regular.exe'), patch.object(khoidong.subprocess, 'run', return_value=subprocess.CompletedProcess([], 1)), contextlib.redirect_stdout(io.StringIO()):
            self.assertEqual(khoidong.use_preferred_runtime(), 1)

    def test_free_threaded_without_regular_runtime_still_starts(self):
        with patch.object(khoidong.os, 'name', 'nt'), patch.object(khoidong, 'is_free_threaded', return_value=True), patch.dict(khoidong.os.environ, {}, clear=True), patch.object(khoidong, 'find_regular_windows_runtime', return_value=None), patch.object(khoidong.subprocess, 'run') as process, contextlib.redirect_stdout(io.StringIO()) as output:
            self.assertIsNone(khoidong.use_preferred_runtime())
        process.assert_not_called()
        self.assertIn('free-threaded hiện tại', output.getvalue())

    def test_relaunch_failure_keeps_current_runtime(self):
        with patch.object(khoidong.os, 'name', 'nt'), patch.object(khoidong, 'is_free_threaded', return_value=True), patch.dict(khoidong.os.environ, {}, clear=True), patch.object(khoidong, 'find_regular_windows_runtime', return_value=r'C:\regular.exe'), patch.object(khoidong.subprocess, 'run', side_effect=OSError('access denied')), contextlib.redirect_stdout(io.StringIO()):
            self.assertIsNone(khoidong.use_preferred_runtime())

    def test_standard_runtime_and_guard_skip_probes(self):
        for free_threaded, guard in [(False, None), (True, '1')]:
            environment = {} if guard is None else {khoidong.RUNTIME_GUARD: guard}
            with self.subTest(free_threaded=free_threaded, guard=guard), patch.object(khoidong.os, 'name', 'nt'), patch.object(khoidong, 'is_free_threaded', return_value=free_threaded), patch.dict(khoidong.os.environ, environment, clear=True), patch.object(khoidong, 'find_regular_windows_runtime') as find:
                self.assertIsNone(khoidong.use_preferred_runtime())
            find.assert_not_called()

    def test_non_windows_runtime_is_unchanged(self):
        with patch.object(khoidong.os, 'name', 'posix'), patch.object(khoidong, 'find_regular_windows_runtime') as find:
            self.assertIsNone(khoidong.use_preferred_runtime())
        find.assert_not_called()

if __name__=='__main__':unittest.main()
