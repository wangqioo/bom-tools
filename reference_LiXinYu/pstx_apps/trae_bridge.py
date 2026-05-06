# -*- coding: utf-8 -*-
"""HTTP bridge for Trae/external agents that cannot run the local Python CLI.

The bridge is intentionally thin: it exposes the existing public CLI command
contract over JSON/HTTP without allowing arbitrary shell execution.  It should
run on the analysis/upper machine where the project files and Python
environment exist; Trae can be on another machine and only needs the bridge URL.
"""

from __future__ import annotations

import argparse
import contextlib
from dataclasses import dataclass
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import io
import json
import os
from pathlib import Path
import socket
import tempfile
import time
from typing import Any, Dict, Iterable, List, Mapping, Optional
from urllib.parse import unquote, urlparse

from pstx_apps import cli as pstx_cli


BRIDGE_VERSION = "1"
BRIDGE_SCHEMA_VERSION = "pstx-trae-bridge.v1"
DEFAULT_BRIDGE_HOST = "127.0.0.1"
DEFAULT_BRIDGE_PORT = 48765
MAX_TCP_PORT = 65535
BRIDGE_ONLY_ARG_KEYS = {
    "run_id",
    "web_run_id",
    "latest_run",
    "use_latest_run",
}
LATEST_RUN_ALIASES = {"", "latest", "last", "__latest__"}


class BridgeArgumentError(ValueError):
    """Invalid bridge request payload."""


@dataclass
class BridgeWebRunContext:
    run_id: str
    project_name: str
    project_root: str
    bundle_source: str
    bundle_cache_path: str


POSITIONAL_ARGS: Dict[str, List[str]] = {
    "schema": ["schema_command"],
    "inspect": ["project_root"],
    "analyze": ["project_root"],
    "query": ["project_root"],
    "batch-query": ["project_root"],
    "module-review": ["project_root"],
    "report-table": ["project_root"],
    "report-aggregate": ["project_root"],
    "evidence-pack": ["project_root"],
    "net-catalog": ["project_root"],
    "topology-netlist": ["project_root"],
    "cadence-page": ["project_root"],
    "cadence-index": ["project_root"],
    "csa-geometry": ["project_root"],
    "schematic-pdf-annotate": ["pdf", "project_root"],
    "harness-skills": ["skill_id"],
    "datasheet-template": ["template_id"],
    "compare": ["left_project_root", "right_project_root"],
    "agent-run-status": ["agent_run_id"],
    "agent-run-artifacts": ["agent_run_id"],
    "agent-run-trace": ["agent_run_id"],
    "offline-migration": ["offline_action", "package_root"],
}


def _json_default(value: Any) -> str:
    if isinstance(value, Path):
        return str(value)
    return str(value)


def _generated_at() -> str:
    return time.strftime("%Y-%m-%d %H:%M:%S")


def _bridge_envelope(command: str, payload: Mapping[str, Any]) -> Dict[str, Any]:
    return {
        "ok": True,
        "interface": "pstx-trae-bridge",
        "interface_version": BRIDGE_VERSION,
        "schema_version": BRIDGE_SCHEMA_VERSION,
        "command": command,
        "generated_at": _generated_at(),
        **dict(payload),
    }


def _error_envelope(command: str, exc: BaseException, *, status: int = 400) -> Dict[str, Any]:
    error_code = "invalid_request" if isinstance(exc, (BridgeArgumentError, ValueError)) else "internal_error"
    return {
        "ok": False,
        "interface": "pstx-trae-bridge",
        "interface_version": BRIDGE_VERSION,
        "schema_version": BRIDGE_SCHEMA_VERSION,
        "command": command,
        "generated_at": _generated_at(),
        "http_status": status,
        "error_code": error_code,
        "error_message": str(exc),
        "error": {
            "code": error_code,
            "message": str(exc),
            "type": exc.__class__.__name__,
        },
    }


def _flag_name(name: str) -> str:
    return "--" + str(name).strip().replace("_", "-")


def _append_option(argv: List[str], name: str, value: Any) -> None:
    if value is None or value == "":
        return
    flag = _flag_name(name)
    if isinstance(value, bool):
        if value:
            argv.append(flag)
        return
    if isinstance(value, (list, tuple)):
        for item in value:
            if item is None or item == "":
                continue
            argv.extend([flag, str(item)])
        return
    argv.extend([flag, str(value)])


def _truthy(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    if value is None:
        return False
    text = str(value).strip().lower()
    return text in {"1", "true", "yes", "y", "on", "latest"}


def _command_accepts_bundle_cache(command: str) -> bool:
    schema = pstx_cli.CLI_COMMAND_SCHEMAS.get(command, {}) or {}
    cache = str(schema.get("cache", "") or "")
    if "--bundle-cache-in" in cache:
        return True
    return any("--bundle-cache-in" in str(item) for item in (schema.get("inputs", []) or []))


def _strip_bridge_only_args(args: Mapping[str, Any]) -> Dict[str, Any]:
    return {
        str(key): value
        for key, value in dict(args or {}).items()
        if str(key) not in BRIDGE_ONLY_ARG_KEYS
    }


def _requested_web_run_id(args: Mapping[str, Any], payload: Mapping[str, Any]) -> str:
    for key in ("run_id", "web_run_id"):
        value = args.get(key, payload.get(key))
        if value is not None and str(value).strip():
            return str(value).strip()
    for key in ("latest_run", "use_latest_run"):
        if _truthy(args.get(key, payload.get(key))):
            return "latest"
    return ""


def _has_explicit_analysis_source(args: Mapping[str, Any]) -> bool:
    return bool(
        str(args.get("project_root", "") or "").strip()
        or str(args.get("bundle_cache_in", "") or "").strip()
    )


def _command_does_not_need_project_source(command: str, args: Mapping[str, Any]) -> bool:
    return command == "csa-geometry" and _truthy(args.get("demo"))


def _web_project_summaries() -> List[Dict[str, Any]]:
    from pstx_webapp.run_store import list_project_summaries

    return list_project_summaries()


def _web_project_summary(run_id: str) -> Optional[Dict[str, Any]]:
    summaries = _web_project_summaries()
    requested = str(run_id or "").strip()
    if requested.lower() in LATEST_RUN_ALIASES:
        return summaries[0] if summaries else None
    return next((item for item in summaries if str(item.get("run_id", "")) == requested), None)


def _resolve_web_run_context(run_id: str, *, bundle_source: str) -> BridgeWebRunContext:
    from pstx_webapp.run_store import get_run

    summary = _web_project_summary(run_id)
    if summary is None:
        requested = str(run_id or "").strip() or "latest"
        raise BridgeArgumentError(
            f"Web 已分析项目中找不到 run_id={requested!r}；"
            "请先在 Web 端完成分析，或调用 /v1/projects 查看可用 run。"
        )
    resolved_run_id = str(summary.get("run_id", "") or "").strip()
    payload = get_run(resolved_run_id)
    if not isinstance(payload, Mapping):
        raise BridgeArgumentError(f"Web run 不存在或已过期：{resolved_run_id}")
    bundle = payload.get("bundle")
    if not isinstance(bundle, Mapping) or "components" not in bundle or "nets" not in bundle:
        raise BridgeArgumentError(f"Web run 缺少可复用的分析 bundle：{resolved_run_id}")

    with tempfile.NamedTemporaryFile(
        "w",
        encoding="utf-8",
        prefix="pstx-bridge-bundle-",
        suffix=".json",
        delete=False,
    ) as tmp:
        json.dump({"bundle": dict(bundle)}, tmp, ensure_ascii=False, default=_json_default)
        tmp_path = tmp.name
    return BridgeWebRunContext(
        run_id=resolved_run_id,
        project_name=str(summary.get("project_name", "") or bundle.get("project_name", "") or ""),
        project_root=str(summary.get("project_root", "") or bundle.get("project_root", "") or ""),
        bundle_source=bundle_source,
        bundle_cache_path=tmp_path,
    )


def _coerce_argv(argv: Iterable[Any]) -> List[str]:
    result = [str(item) for item in argv if item is not None and str(item) != ""]
    if not result:
        raise BridgeArgumentError("argv must contain a command")
    command = result[0]
    if command not in pstx_cli.CLI_COMMAND_SCHEMAS:
        raise BridgeArgumentError(f"unsupported command: {command}")
    return result


def build_cli_argv(command: str,
                   args: Optional[Mapping[str, Any]] = None,
                   argv: Optional[Iterable[Any]] = None) -> List[str]:
    """Build a safe CLI argv list from a bridge request.

    `argv` is accepted for callers that already store the CLI-shaped contract,
    but it still must start with a known public CLI command.  `args` is the
    preferred machine-to-machine shape because it does not require Trae to know
    command-line quoting rules.
    """

    if argv is not None:
        return _coerce_argv(argv)

    command = str(command or "").strip()
    if not command:
        raise BridgeArgumentError("missing command")
    if command not in pstx_cli.CLI_COMMAND_SCHEMAS:
        raise BridgeArgumentError(f"unsupported command: {command}")
    if args is not None and not isinstance(args, Mapping):
        raise BridgeArgumentError("args must be a JSON object")

    remaining: Dict[str, Any] = _strip_bridge_only_args(args or {})
    cli_argv: List[str] = [command]
    for positional_name in POSITIONAL_ARGS.get(command, []):
        value = remaining.pop(positional_name, None)
        if value is None or value == "":
            continue
        cli_argv.append(str(value))

    # `pretty` is useful for humans in CLI mode but just increases bridge
    # payload size.  Keep it accepted, but default bridge calls should omit it.
    for key in sorted(remaining.keys()):
        value = remaining[key]
        # `batch-query --items` is a single comma/newline-separated argument,
        # not a repeatable flag. Accept JSON arrays anyway because remote
        # callers naturally model batch targets as lists.
        if key == "items" and isinstance(value, (list, tuple)):
            value = ",".join(str(item) for item in value if item is not None and str(item) != "")
        _append_option(cli_argv, key, value)
    return cli_argv


def _parse_cli_stdout(text: str) -> Dict[str, Any]:
    stripped = (text or "").strip()
    if not stripped:
        return {}
    try:
        parsed = json.loads(stripped)
    except json.JSONDecodeError:
        # Defensive fallback: if a dependency prints a warning before JSON,
        # parse the last JSON-looking line instead of failing the whole bridge.
        for line in reversed(stripped.splitlines()):
            line = line.strip()
            if not line.startswith("{"):
                continue
            try:
                parsed = json.loads(line)
                break
            except json.JSONDecodeError:
                continue
        else:
            raise BridgeArgumentError("CLI did not return JSON output")
    if not isinstance(parsed, dict):
        raise BridgeArgumentError("CLI output must be a JSON object")
    return parsed


def run_bridge_payload(payload: Mapping[str, Any]) -> Dict[str, Any]:
    """Run one public CLI command and return its JSON envelope with bridge meta."""

    if not isinstance(payload, Mapping):
        raise BridgeArgumentError("request body must be a JSON object")
    raw_argv = payload.get("argv")
    args_payload = payload.get("args") if "args" in payload else None
    if args_payload is not None and not isinstance(args_payload, Mapping):
        raise BridgeArgumentError("args must be a JSON object")
    args_for_context: Mapping[str, Any] = args_payload or {}
    requested_run_id = _requested_web_run_id(args_for_context, payload)
    if raw_argv is not None and requested_run_id:
        raise BridgeArgumentError("run_id/latest_run is only supported with JSON args, not raw argv")
    argv = build_cli_argv(
        str(payload.get("command") or ""),
        args=args_payload,
        argv=raw_argv if raw_argv is not None else None,
    )
    command = argv[0]
    web_context: Optional[BridgeWebRunContext] = None
    explicit_source = _has_explicit_analysis_source(args_for_context)
    supports_web_run = _command_accepts_bundle_cache(command)
    needs_source = supports_web_run and not _command_does_not_need_project_source(command, args_for_context)

    if requested_run_id and explicit_source:
        raise BridgeArgumentError("run_id/latest_run cannot be combined with project_root or bundle_cache_in")
    if requested_run_id and not supports_web_run:
        raise BridgeArgumentError(f"command does not support Web run reuse: {command}")
    if raw_argv is None and needs_source and not explicit_source:
        source = "web_run" if requested_run_id else "web_run_latest"
        web_context = _resolve_web_run_context(requested_run_id or "latest", bundle_source=source)
        argv.extend(["--bundle-cache-in", web_context.bundle_cache_path])

    stream = io.StringIO()
    try:
        with contextlib.redirect_stdout(stream):
            exit_code = pstx_cli.main(argv)
        cli_payload = _parse_cli_stdout(stream.getvalue())
        if not cli_payload:
            cli_payload = {
                "ok": exit_code == 0,
                "schema_version": pstx_cli.SCHEMA_VERSION,
                "command": command,
            }
    finally:
        if web_context:
            try:
                Path(web_context.bundle_cache_path).unlink(missing_ok=True)
            except OSError:
                pass

    bridge_meta: Dict[str, Any] = {
        "interface": "pstx-trae-bridge",
        "schema_version": BRIDGE_SCHEMA_VERSION,
        "cli_exit_code": exit_code,
        "cli_argv": argv,
        "transport": "http-json",
    }
    if web_context:
        bridge_meta.update({
            "bundle_source": web_context.bundle_source,
            "run_id": web_context.run_id,
            "project_name": web_context.project_name,
            "project_root": web_context.project_root,
            "project_discovery_endpoint": "/v1/projects",
            "source_args": {"run_id": web_context.run_id},
            "detail_command_note": (
                "This bridge call used a temporary bundle-cache file; "
                "remote agents should reuse run_id instead of CLI temp paths."
            ),
        })
    elif explicit_source:
        bridge_meta["bundle_source"] = "explicit_args"
    cli_payload["bridge"] = bridge_meta
    return cli_payload


class TraeBridgeServer(ThreadingHTTPServer):
    token: str
    cors_origin: str


@dataclass
class BridgeSidecar:
    server: TraeBridgeServer
    host: str
    port: int
    thread_name: str = "pstx-trae-bridge"

    @property
    def url(self) -> str:
        return f"http://{self.host}:{self.port}"

    def stop(self) -> None:
        try:
            self.server.shutdown()
        except Exception:
            pass
        try:
            self.server.server_close()
        except Exception:
            pass


def port_is_available(port: int, host: str = DEFAULT_BRIDGE_HOST) -> bool:
    if port < 0 or port > MAX_TCP_PORT:
        return False
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        try:
            sock.bind((host, port))
        except OSError:
            return False
    return True


def resolve_bridge_port(preferred_port: int = DEFAULT_BRIDGE_PORT,
                        host: str = DEFAULT_BRIDGE_HOST,
                        max_attempts: int = 50) -> int:
    for offset in range(max_attempts + 1):
        candidate = int(preferred_port) + offset
        if candidate > MAX_TCP_PORT:
            break
        if port_is_available(candidate, host):
            return candidate
    raise RuntimeError(
        f"Unable to find a free Trae bridge port in range "
        f"{preferred_port}-{min(int(preferred_port) + max_attempts, MAX_TCP_PORT)}."
    )


class TraeBridgeHandler(BaseHTTPRequestHandler):
    server_version = "PSTXTraeBridge/1"

    def _cors_origin(self) -> str:
        return getattr(self.server, "cors_origin", "*") or "*"

    def _send_json(self, status: int, payload: Mapping[str, Any]) -> None:
        data = json.dumps(payload, ensure_ascii=False, default=_json_default).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Access-Control-Allow-Origin", self._cors_origin())
        self.send_header("Access-Control-Allow-Headers", "Content-Type, X-PSTX-Bridge-Token")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.end_headers()
        self.wfile.write(data)

    def _authorized(self) -> bool:
        token = getattr(self.server, "token", "") or ""
        if not token:
            return True
        return self.headers.get("X-PSTX-Bridge-Token", "") == token

    def _reject_unauthorized(self) -> None:
        self._send_json(401, _error_envelope("auth", BridgeArgumentError("missing or invalid bridge token"), status=401))

    def do_OPTIONS(self) -> None:  # noqa: N802 - stdlib hook
        self._send_json(200, _bridge_envelope("options", {"allowed": ["GET", "POST", "OPTIONS"]}))

    def do_GET(self) -> None:  # noqa: N802 - stdlib hook
        if not self._authorized():
            self._reject_unauthorized()
            return
        path = urlparse(self.path).path.rstrip("/") or "/"
        if path in {"/health", "/v1/health"}:
            self._send_json(200, _bridge_envelope("health", {
                "status": "ok",
                "capability_count": len(pstx_cli.CLI_COMMAND_SCHEMAS),
                "notes": [
                    "Run this bridge on the analysis machine; Trae may run elsewhere.",
                    "Use POST /v1/run to execute whitelisted PSTX CLI commands.",
                    "Use GET /v1/projects to discover Web-analyzed runs and pass args.run_id instead of local paths.",
                ],
            }))
            return
        if path in {"/v1/capabilities", "/capabilities"}:
            capabilities = [
                {
                    "id": command_id,
                    "description": schema.get("purpose", ""),
                    "outputs": list(schema.get("outputs", []) or []),
                    "cache": schema.get("cache", "none"),
                }
                for command_id, schema in pstx_cli.CLI_COMMAND_SCHEMAS.items()
            ]
            self._send_json(200, _bridge_envelope("capabilities", {
                "capabilities": capabilities,
                "cli_schema_version": pstx_cli.SCHEMA_VERSION,
            }))
            return
        if path in {"/v1/projects", "/projects"}:
            projects = _web_project_summaries()
            self._send_json(200, _bridge_envelope("projects", {
                "projects": projects,
                "count": len(projects),
                "latest_run_id": str(projects[0].get("run_id", "")) if projects else "",
                "notes": [
                    "Use args.run_id with POST /v1/run to reuse a Web-analyzed project.",
                    "Use run_id='latest' only when the latest Web run is the intended project.",
                ],
            }))
            return
        if path.startswith("/v1/projects/") or path.startswith("/projects/"):
            run_id = unquote(path.rsplit("/", 1)[-1])
            summary = _web_project_summary(run_id)
            if summary is None:
                self._send_json(404, _error_envelope("projects", BridgeArgumentError(f"unknown Web run_id: {run_id}"), status=404))
                return
            self._send_json(200, _bridge_envelope("projects", {
                "project": summary,
            }))
            return
        if path in {"/v1/schema", "/schema"}:
            self._send_json(200, _bridge_envelope("schema", {
                "commands": list(pstx_cli.CLI_COMMAND_SCHEMAS.keys()),
                "schema": pstx_cli.CLI_COMMAND_SCHEMAS,
                "cli_schema_version": pstx_cli.SCHEMA_VERSION,
            }))
            return
        if path.startswith("/v1/schema/"):
            command = path.rsplit("/", 1)[-1]
            schema = pstx_cli.CLI_COMMAND_SCHEMAS.get(command)
            if not schema:
                self._send_json(404, _error_envelope(command, BridgeArgumentError(f"unknown command schema: {command}"), status=404))
                return
            self._send_json(200, _bridge_envelope("schema", {
                "commands": [command],
                "schema": {command: schema},
                "cli_schema_version": pstx_cli.SCHEMA_VERSION,
            }))
            return
        self._send_json(404, _error_envelope("", BridgeArgumentError(f"unknown endpoint: {path}"), status=404))

    def do_POST(self) -> None:  # noqa: N802 - stdlib hook
        if not self._authorized():
            self._reject_unauthorized()
            return
        path = urlparse(self.path).path.rstrip("/") or "/"
        if path not in {"/v1/run", "/run"}:
            self._send_json(404, _error_envelope("", BridgeArgumentError(f"unknown endpoint: {path}"), status=404))
            return
        try:
            content_length = int(self.headers.get("Content-Length", "0") or "0")
            raw = self.rfile.read(content_length).decode("utf-8") if content_length else "{}"
            payload = json.loads(raw or "{}")
            result = run_bridge_payload(payload)
            status = 200 if result.get("ok", False) else 400
            self._send_json(status, result)
        except json.JSONDecodeError as exc:
            self._send_json(400, _error_envelope("", BridgeArgumentError(f"invalid JSON body: {exc}"), status=400))
        except Exception as exc:  # pragma: no cover - exercised through HTTP integration.
            status = 400 if isinstance(exc, (BridgeArgumentError, ValueError)) else 500
            command = ""
            if "payload" in locals() and isinstance(payload, Mapping):
                command = str(payload.get("command") or "")
            self._send_json(status, _error_envelope(command, exc, status=status))

    def log_message(self, format: str, *args: Any) -> None:  # noqa: A003 - stdlib hook
        # Keep bridge stdout machine-readable enough for operators.
        print(f"[pstx-trae-bridge] {self.address_string()} - {format % args}")


def run_server(*,
               host: str = DEFAULT_BRIDGE_HOST,
               port: int = DEFAULT_BRIDGE_PORT,
               token: str = "",
               cors_origin: str = "*") -> None:
    resolved_port = resolve_bridge_port(int(port), host)
    server = TraeBridgeServer((host, resolved_port), TraeBridgeHandler)
    server.token = token
    server.cors_origin = cors_origin
    print(json.dumps({
        "ok": True,
        "interface": "pstx-trae-bridge",
        "schema_version": BRIDGE_SCHEMA_VERSION,
        "event": "server.start",
        "host": host,
        "port": resolved_port,
        "requested_port": int(port),
        "auth": "token" if token else "none",
        "endpoints": ["/v1/health", "/v1/capabilities", "/v1/projects", "/v1/schema", "/v1/run"],
    }, ensure_ascii=False))
    try:
        server.serve_forever()
    except KeyboardInterrupt:  # pragma: no cover - operator path.
        print(json.dumps({
            "ok": True,
            "interface": "pstx-trae-bridge",
            "schema_version": BRIDGE_SCHEMA_VERSION,
            "event": "server.stop",
        }, ensure_ascii=False))
    finally:
        server.server_close()


def start_background_bridge(*,
                            host: str = DEFAULT_BRIDGE_HOST,
                            port: int = DEFAULT_BRIDGE_PORT,
                            token: str = "",
                            cors_origin: str = "*",
                            quiet: bool = False) -> BridgeSidecar:
    """Start the Trae bridge as a daemon sidecar for Web/local UI launchers."""

    import threading

    resolved_port = resolve_bridge_port(int(port), host)
    server = TraeBridgeServer((host, resolved_port), TraeBridgeHandler)
    server.token = token
    server.cors_origin = cors_origin
    thread = threading.Thread(
        target=server.serve_forever,
        name="pstx-trae-bridge",
        daemon=True,
    )
    thread.start()
    sidecar = BridgeSidecar(server=server, host=host, port=resolved_port)
    if not quiet:
        print(json.dumps({
            "ok": True,
            "interface": "pstx-trae-bridge",
            "schema_version": BRIDGE_SCHEMA_VERSION,
            "event": "sidecar.start",
            "url": sidecar.url,
            "requested_port": int(port),
            "auth": "token" if token else "none",
        }, ensure_ascii=False))
    return sidecar


def bridge_config_from_env() -> Dict[str, Any]:
    return {
        "host": os.environ.get("PSTX_TRAE_BRIDGE_HOST", DEFAULT_BRIDGE_HOST),
        "port": int(os.environ.get("PSTX_TRAE_BRIDGE_PORT", str(DEFAULT_BRIDGE_PORT)) or DEFAULT_BRIDGE_PORT),
        "token": os.environ.get("PSTX_TRAE_BRIDGE_TOKEN", ""),
        "cors_origin": os.environ.get("PSTX_TRAE_BRIDGE_CORS_ORIGIN", "*"),
    }


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="PSTX Trae HTTP bridge")
    parser.add_argument("--host", default=DEFAULT_BRIDGE_HOST, help=f"bind host, default {DEFAULT_BRIDGE_HOST}")
    parser.add_argument("--port", type=int, default=DEFAULT_BRIDGE_PORT, help=f"bind port, default {DEFAULT_BRIDGE_PORT}")
    parser.add_argument("--allow-remote", action="store_true", help="if host is default, bind 0.0.0.0 for remote Trae clients")
    parser.add_argument("--token", default="", help="optional X-PSTX-Bridge-Token required by clients")
    parser.add_argument("--cors-origin", default="*", help="CORS origin for browser-like clients")
    return parser


def main(argv: Optional[List[str]] = None) -> int:
    args = build_parser().parse_args(argv)
    host = str(args.host or "127.0.0.1")
    if args.allow_remote and host == "127.0.0.1":
        host = "0.0.0.0"
    run_server(host=host, port=int(args.port), token=str(args.token or ""), cors_origin=str(args.cors_origin or "*"))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
