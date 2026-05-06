# -*- coding: utf-8 -*-
"""PSTX desktop shell for the localhost Web UI.

The desktop UI intentionally reuses the Web UI backend and frontend.  When
pywebview is available it embeds the localhost page in a native window; when it
is not available it opens the same localhost URL in the system browser.
"""

from __future__ import annotations

import argparse
import os
import socket
import subprocess
import sys
import threading
import time
import webbrowser
from dataclasses import dataclass, field
from typing import List, Optional

from pstx_apps.trae_bridge import DEFAULT_BRIDGE_PORT, BridgeSidecar, bridge_config_from_env, start_background_bridge
from pstx_webapp.app_factory import create_app
from pstx_webapp.server import DEFAULT_HOST, DEFAULT_PORT, resolve_port


WINDOW_TITLE = 'PSTX 本地桌面工作台'
WINDOW_MIN_SIZE = (1120, 760)
WINDOW_SIZE = (1520, 980)


def _runtime_url(host: str, port: int) -> str:
    return f'http://{host}:{port}/'


def _ensure_pywebview():
    try:
        import webview  # type: ignore
        return webview
    except Exception:
        print('未检测到 pywebview，正在尝试安装本地桌面套壳依赖...')
        subprocess.check_call([
            sys.executable,
            '-m',
            'pip',
            'install',
            '--upgrade',
            'pywebview>=5,<6',
        ])
        import webview  # type: ignore
        return webview


@dataclass
class LocalUiSession:
    preferred_port: int = DEFAULT_PORT
    host: str = DEFAULT_HOST
    app: object = None
    port: int = 0
    url: str = ''
    bridge_port: int = DEFAULT_BRIDGE_PORT
    start_trae_bridge: bool = True
    _server: object = field(default=None, init=False, repr=False)
    _thread: Optional[threading.Thread] = field(default=None, init=False, repr=False)
    _bridge: Optional[BridgeSidecar] = field(default=None, init=False, repr=False)

    def start(self) -> str:
        # create_app() lazily verifies Flask/Werkzeug. Import make_server only
        # after that so importing this module has no dependency side effects.
        app = self.app or create_app()
        from werkzeug.serving import make_server  # type: ignore

        self.port = resolve_port(self.preferred_port, self.host)
        self.url = _runtime_url(self.host, self.port)
        self._server = make_server(self.host, self.port, app, threaded=True)
        self._thread = threading.Thread(
            target=self._server.serve_forever,
            name='pstx-local-ui-server',
            daemon=True,
        )
        self._thread.start()
        if self.start_trae_bridge and os.environ.get("PSTX_TRAE_BRIDGE_DISABLED", "").lower() not in {"1", "true", "yes"}:
            bridge_config = bridge_config_from_env()
            bridge_config["port"] = self.bridge_port or bridge_config["port"]
            self._bridge = start_background_bridge(**bridge_config)
        self._wait_until_ready()
        return self.url

    def _wait_until_ready(self, timeout_s: float = 5.0) -> None:
        deadline = time.time() + timeout_s
        last_error: Optional[OSError] = None
        while time.time() < deadline:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
                sock.settimeout(0.2)
                try:
                    sock.connect((self.host, self.port))
                    return
                except OSError as exc:
                    last_error = exc
            time.sleep(0.05)
        raise RuntimeError(
            f'Local UI server did not become ready on {self.host}:{self.port}: {last_error}'
        )

    def stop(self) -> None:
        if self._server is not None:
            try:
                self._server.shutdown()
            except Exception:
                pass
            try:
                self._server.server_close()
            except Exception:
                pass
            self._server = None
        if self._thread is not None:
            self._thread.join(timeout=2.0)
            self._thread = None
        if self._bridge is not None:
            self._bridge.stop()
            self._bridge = None

    def is_running(self) -> bool:
        return bool(self._thread and self._thread.is_alive())


def _wait_forever() -> None:
    while True:
        time.sleep(1.0)


def launch_local_ui(
    *,
    preferred_port: int = DEFAULT_PORT,
    title: str = WINDOW_TITLE,
    force_browser: bool = False,
    install_pywebview: bool = False,
    start_trae_bridge: bool = True,
    trae_bridge_port: int = DEFAULT_BRIDGE_PORT,
) -> int:
    session = LocalUiSession(
        preferred_port=preferred_port,
        start_trae_bridge=start_trae_bridge,
        bridge_port=trae_bridge_port,
    )
    url = session.start()
    print(f'PSTX 本地 UI 已启动：{url}')
    if session._bridge is not None:
        print(f'PSTX Trae Bridge 已随本地 UI 启动：{session._bridge.url}/v1')

    if force_browser:
        webbrowser.open(url)
        try:
            _wait_forever()
        except KeyboardInterrupt:
            return 0
        finally:
            session.stop()

    try:
        if install_pywebview:
            webview = _ensure_pywebview()
        else:
            import webview  # type: ignore
    except Exception as exc:
        print(f'无法启用 pywebview 桌面窗口（{exc}），改为打开系统浏览器。')
        webbrowser.open(url)
        try:
            _wait_forever()
        except KeyboardInterrupt:
            return 0
        finally:
            session.stop()

    try:
        webview.create_window(
            title,
            url,
            width=WINDOW_SIZE[0],
            height=WINDOW_SIZE[1],
            min_size=WINDOW_MIN_SIZE,
            confirm_close=True,
            text_select=True,
        )
        start_kwargs = {'debug': False}
        if sys.platform == 'win32':
            start_kwargs['gui'] = 'edgechromium'
        webview.start(**start_kwargs)
        return 0
    finally:
        session.stop()


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description='Run PSTX desktop local UI shell')
    parser.add_argument('--port', type=int, default=DEFAULT_PORT, help='preferred localhost port, default 44441')
    parser.add_argument('--title', default=WINDOW_TITLE, help='desktop window title')
    parser.add_argument('--browser', action='store_true', help='open the localhost UI in the system browser')
    parser.add_argument('--install-pywebview', action='store_true', help='try to install pywebview when it is missing')
    parser.add_argument('--no-trae-bridge', action='store_true', help='do not start the Trae HTTP bridge sidecar')
    parser.add_argument('--trae-bridge-port', type=int, default=DEFAULT_BRIDGE_PORT, help=f'Trae bridge sidecar port, default {DEFAULT_BRIDGE_PORT}')
    args = parser.parse_args(argv)
    return launch_local_ui(
        preferred_port=args.port,
        title=args.title,
        force_browser=args.browser,
        install_pywebview=args.install_pywebview,
        start_trae_bridge=not args.no_trae_bridge,
        trae_bridge_port=args.trae_bridge_port,
    )


if __name__ == '__main__':
    raise SystemExit(main())
