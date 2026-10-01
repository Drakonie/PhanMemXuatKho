import contextlib
import io
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
import app
import khoidong

class StartupTests(unittest.TestCase):
    def test_port_in_use_tries_next_port(self):
        with patch.object(app,'serve',side_effect=[OSError(98,'in use'),None]) as server, contextlib.redirect_stdout(io.StringIO()):
            self.assertEqual(khoidong.main(),0)
        self.assertEqual(server.call_args_list[0].kwargs,{'port':8080,'open_browser':True})
        self.assertEqual(server.call_args_list[1].kwargs,{'port':8081,'open_browser':True})
    def test_failed_binding_does_not_open_browser(self):
        with patch.object(app,'ThreadingHTTPServer',side_effect=OSError('bind failed')),patch('webbrowser.open') as browser:
            with self.assertRaises(OSError): app.serve(open_browser=True)
            browser.assert_not_called()

if __name__=='__main__':unittest.main()
