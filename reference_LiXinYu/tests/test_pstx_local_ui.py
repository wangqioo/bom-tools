import socket
import unittest
import urllib.request

from pstx_apps import local_ui as app_local_ui
from pstx_webapp import app_factory as webapp_factory


def _pick_free_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind((app_local_ui.DEFAULT_HOST, 0))
        sock.listen(1)
        return sock.getsockname()[1]


class LocalUiTests(unittest.TestCase):
    def test_apps_local_ui_entrypoint_exports_public_shell_api(self):
        self.assertIs(app_local_ui.create_app, webapp_factory.create_app)

    def test_default_port_tracks_web_entrypoint(self):
        self.assertEqual(44441, app_local_ui.DEFAULT_PORT)

    def test_runtime_url_uses_localhost_port(self):
        self.assertEqual(
            'http://127.0.0.1:44441/',
            app_local_ui._runtime_url('127.0.0.1', 44441),
        )

    def test_local_ui_session_serves_home_page(self):
        session = app_local_ui.LocalUiSession(preferred_port=_pick_free_port())
        try:
            url = session.start()
            with urllib.request.urlopen(url, timeout=5) as response:
                html = response.read().decode('utf-8', errors='replace')
        finally:
            session.stop()

        self.assertIn('PSTX', html)
        self.assertFalse(session.is_running())


if __name__ == '__main__':
    unittest.main()
