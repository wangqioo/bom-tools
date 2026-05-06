"""Local server helpers shared by Web and desktop launchers."""

from __future__ import annotations

import socket


DEFAULT_HOST = "127.0.0.1"
DEFAULT_PORT = 44441
MAX_TCP_PORT = 65535


def port_is_available(port: int, host: str = DEFAULT_HOST) -> bool:
    if port < 0 or port > MAX_TCP_PORT:
        return False
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        try:
            sock.bind((host, port))
        except OSError:
            return False
    return True


def resolve_port(preferred_port: int, host: str = DEFAULT_HOST, max_attempts: int = 20) -> int:
    for offset in range(max_attempts + 1):
        candidate = preferred_port + offset
        if candidate > MAX_TCP_PORT:
            break
        if port_is_available(candidate, host):
            return candidate
    raise RuntimeError(
        f"Unable to find a free localhost port in range {preferred_port}-{min(preferred_port + max_attempts, MAX_TCP_PORT)}."
    )
