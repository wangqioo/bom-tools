# -*- coding: utf-8 -*-
"""
PSTX localhost Web UI

Run:
    python pstx_web.py
"""

from __future__ import annotations

import argparse
import os
import threading
import webbrowser
from typing import List, Optional

from pstx_apps.trae_bridge import DEFAULT_BRIDGE_PORT, bridge_config_from_env, start_background_bridge
from pstx_webapp.app_factory import create_app
from pstx_webapp.server import DEFAULT_HOST, DEFAULT_PORT, resolve_port


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description='Run PSTX localhost web UI')
    parser.add_argument('--port', type=int, default=DEFAULT_PORT, help='localhost port, default 44441')
    parser.add_argument('--no-browser', action='store_true', help='do not auto-open the browser')
    parser.add_argument('--no-trae-bridge', action='store_true', help='do not start the Trae HTTP bridge sidecar')
    parser.add_argument('--trae-bridge-port', type=int, default=0, help=f'Trae bridge sidecar port, default {DEFAULT_BRIDGE_PORT}')
    args = parser.parse_args(argv)

    resolved_port = resolve_port(args.port, DEFAULT_HOST)
    app = create_app()
    bridge_sidecar = None
    if not args.no_trae_bridge and os.environ.get("PSTX_TRAE_BRIDGE_DISABLED", "").lower() not in {"1", "true", "yes"}:
        bridge_config = bridge_config_from_env()
        if args.trae_bridge_port:
            bridge_config["port"] = args.trae_bridge_port
        bridge_sidecar = start_background_bridge(**bridge_config)
    url = f'http://{DEFAULT_HOST}:{resolved_port}/'
    if not args.no_browser:
        threading.Timer(1.0, lambda: webbrowser.open(url)).start()

    if resolved_port != args.port:
        print(f'Requested port {args.port} is busy; falling back to localhost port {resolved_port}.')
    print(f'PSTX Web UI is listening on {url}')
    if bridge_sidecar is not None:
        print(f'PSTX Trae Bridge sidecar is listening on {bridge_sidecar.url}/v1')
    print('This service is bound to 127.0.0.1 only and cannot be accessed from other machines.')
    try:
        app.run(host=DEFAULT_HOST, port=resolved_port, debug=False)
    finally:
        if bridge_sidecar is not None:
            bridge_sidecar.stop()
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
