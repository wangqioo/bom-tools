# -*- coding: utf-8 -*-
"""Production Aster client and report-summary adapter.

The default production path uses Aster ChatFlow because it only requires the
server-side API key and employee number. Room/Auth support is implemented for
deployments that need accessToken-based agents.
"""

from __future__ import annotations

import base64
import binascii
import hashlib
import json
import os
import re
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple
from uuid import uuid4

from pstx_integrations.aster.mock import build_aster_mock_summary


class AsterError(RuntimeError):
    """Base class for Aster integration errors."""

    def __init__(self, message: str, *, diagnostics: Optional[dict] = None):
        super().__init__(message)
        self.diagnostics = diagnostics or {}


class AsterConfigError(AsterError):
    """Raised when live Aster mode is missing required configuration."""


class AsterHttpError(AsterError):
    """Raised when Aster returns a non-2xx HTTP response."""


class AsterResponseError(AsterError):
    """Raised when Aster returns an unexpected JSON shape."""


_LOG_LOCK = threading.RLock()
ASTER_FIXED_BASE_URL = 'https://aigc.huaqin.com'
DEFAULT_ASTER_TIMEOUT_SECONDS = 600.0
SENSITIVE_KEY_RE = re.compile(r'(secret|token|apikey|api_key|authorization|ciphertext|password|passwd)', re.I)
EMBEDDED_SECRET_PATTERNS = [
    (
        re.compile(
            r'(?i)'
            r'(apiKey|api_key|appSecret|app_secret|accessToken|access_token|ciphertext|password|passwd)'
            r'(\s*[:=]\s*["\']?)([^"\',&\s}]+)'
        ),
        r'\1\2<redacted>',
    ),
    (
        re.compile(r'(?i)(Authorization\s*[:=]\s*Bearer\s+)([A-Za-z0-9._~+/=-]{8,})'),
        r'\1<redacted>',
    ),
    (
        re.compile(r'(?i)(Bearer\s+)([A-Za-z0-9._~+/=-]{12,})'),
        r'\1<redacted>',
    ),
]


def _short_hash(value: object) -> str:
    text = str(value or '')
    if not text:
        return ''
    return hashlib.sha256(text.encode('utf-8')).hexdigest()[:12]


def get_aster_log_path(log_file: str = '') -> str:
    if log_file:
        return str(Path(log_file).expanduser())
    return str(Path(__file__).resolve().parents[2] / 'logs' / 'aster_debug.log')


def _truncate(value: object, limit: int = 1600) -> str:
    text = str(value or '')
    if len(text) <= limit:
        return text
    return text[:limit] + f'...<truncated {len(text) - limit} chars>'


def _redact_embedded_secrets(text: str) -> str:
    redacted = str(text or '')
    for pattern, replacement in EMBEDDED_SECRET_PATTERNS:
        redacted = pattern.sub(replacement, redacted)
    return redacted


def _safe_text_preview(text: object, limit: int = 1000) -> str:
    return _truncate(_redact_embedded_secrets(str(text or '')), limit)


def _safe_json_preview(value: object, limit: int = 1000) -> str:
    return _truncate(json.dumps(sanitize_for_aster_log(value), ensure_ascii=False), limit)


def _redact_scalar(key: str, value: object) -> object:
    if value is None:
        return value
    if SENSITIVE_KEY_RE.search(str(key)):
        text = str(value)
        return {
            'redacted': True,
            'length': len(text),
            'sha256_12': _short_hash(text),
        }
    if isinstance(value, str):
        return _truncate(_redact_embedded_secrets(value))
    return value


def sanitize_for_aster_log(value: object, *, parent_key: str = '') -> object:
    if isinstance(value, dict):
        return {
            str(key): sanitize_for_aster_log(child, parent_key=str(key))
            for key, child in value.items()
        }
    if isinstance(value, list):
        return [sanitize_for_aster_log(item, parent_key=parent_key) for item in value[:40]]
    return _redact_scalar(parent_key, value)


def _safe_url(url: str) -> str:
    parsed = urllib.parse.urlsplit(str(url or ''))
    query = urllib.parse.parse_qsl(parsed.query, keep_blank_values=True)
    safe_query = urllib.parse.urlencode([
        (key, '<redacted>' if SENSITIVE_KEY_RE.search(key) else value)
        for key, value in query
    ])
    return urllib.parse.urlunsplit((parsed.scheme, parsed.netloc, parsed.path, safe_query, ''))


def _body_summary(body: Optional[dict], *, include_payload: bool) -> dict:
    if not body:
        return {'present': False}
    summary = {
        'present': True,
        'keys': sorted(str(key) for key in body.keys()),
        'json_chars': len(json.dumps(body, ensure_ascii=False)),
    }
    query = body.get('query')
    if isinstance(query, str):
        summary['query_chars'] = len(query)
        summary['query_sha256_12'] = _short_hash(query)
    inputs = body.get('inputs')
    if isinstance(inputs, dict):
        summary['inputs_keys'] = sorted(str(key) for key in inputs.keys())
    if include_payload:
        summary['payload'] = sanitize_for_aster_log(body)
    return summary


def _write_aster_log(event: str, details: dict, *, log_file: str = '') -> None:
    path = Path(get_aster_log_path(log_file))
    record = {
        'ts': datetime.now().isoformat(timespec='seconds'),
        'event': event,
        **sanitize_for_aster_log(details),
    }
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        max_bytes = _env_int(os.environ.get('PSTX_ASTER_LOG_MAX_BYTES'), 2_000_000, minimum=100_000)
        with _LOG_LOCK:
            if path.exists() and path.stat().st_size > max_bytes:
                rotated = path.with_suffix(path.suffix + '.1')
                if rotated.exists():
                    rotated.unlink()
                path.rename(rotated)
            with path.open('a', encoding='utf-8') as handle:
                handle.write(json.dumps(record, ensure_ascii=False, sort_keys=True) + '\n')
    except Exception:
        # Logging must never break the analysis flow.
        return


def _request_id() -> str:
    return uuid4().hex[:12]


def _diagnostics(context: Optional[dict], **extra) -> dict:
    ctx = dict(context or {})
    log_file = ctx.get('log_file') or ''
    data = {
        'request_id': ctx.get('request_id') or _request_id(),
        'operation': ctx.get('operation') or '',
        'backend': ctx.get('backend') or '',
        'url': _safe_url(ctx.get('url') or ''),
        'log_file': get_aster_log_path(str(log_file)),
        **extra,
    }
    return sanitize_for_aster_log(data)


def _env_bool(value: object, default: bool = False) -> bool:
    if value is None:
        return default
    return str(value).strip().lower() in {'1', 'true', 'yes', 'on'}


def _env_int(value: object, default: int, *, minimum: int = 0) -> int:
    try:
        parsed = int(str(value).strip())
    except (TypeError, ValueError):
        return default
    return max(minimum, parsed)


def _env_float(value: object, default: float, *, minimum: float = 0.1) -> float:
    try:
        parsed = float(str(value).strip())
    except (TypeError, ValueError):
        return default
    return max(minimum, parsed)


def _normalize_origin(value: str) -> str:
    text = str(value or '').strip()
    if not text:
        return ''
    parsed = urllib.parse.urlparse(text)
    return parsed.netloc or text.split('/', 1)[0].split('?', 1)[0]


@dataclass(frozen=True)
class AsterConfig:
    mode: str = 'mock'
    backend: str = 'chat-flow'
    base_url: str = ''
    app_id: str = ''
    app_secret: str = ''
    emp_no: str = ''
    api_key: str = ''
    origin: str = ''
    device_id: str = 'DeviceUniqueIdentifier'
    validate_token: bool = True
    validate_auth_header: str = 'encrypted-time'
    timeout_seconds: float = DEFAULT_ASTER_TIMEOUT_SECONDS
    conversation_id: str = ''
    auto_generate_name: bool = True
    max_rows_per_table: int = 16
    max_payload_chars: int = 60000
    retry_count: int = 1
    retry_backoff_seconds: float = 1.0
    redact_paths: bool = True
    log_file: str = ''
    log_payload: bool = False

    @classmethod
    def from_env(cls, environ: Optional[dict] = None) -> 'AsterConfig':
        env = os.environ if environ is None else environ
        return cls(
            mode=str(env.get('PSTX_ASTER_MODE') or 'mock').strip().lower(),
            backend=str(env.get('PSTX_ASTER_BACKEND') or 'chat-flow').strip().lower(),
            base_url=ASTER_FIXED_BASE_URL,
            app_id=str(env.get('ASTER_APP_ID') or '').strip(),
            app_secret=str(env.get('ASTER_APP_SECRET') or '').strip(),
            emp_no=str(env.get('ASTER_EMP_NO') or '').strip(),
            api_key=str(env.get('ASTER_API_KEY') or '').strip(),
            origin=_normalize_origin(str(env.get('ASTER_ORIGIN') or env.get('PSTX_ASTER_ORIGIN') or '')),
            device_id=str(env.get('ASTER_DEVICE_ID') or 'DeviceUniqueIdentifier').strip(),
            validate_token=_env_bool(env.get('PSTX_ASTER_VALIDATE_TOKEN'), True),
            validate_auth_header=str(env.get('PSTX_ASTER_VALIDATE_AUTH_HEADER') or 'encrypted-time').strip().lower(),
            timeout_seconds=_env_float(env.get('PSTX_ASTER_TIMEOUT_SECONDS'), DEFAULT_ASTER_TIMEOUT_SECONDS, minimum=1.0),
            conversation_id=str(env.get('PSTX_ASTER_CONVERSATION_ID') or '').strip(),
            auto_generate_name=_env_bool(env.get('PSTX_ASTER_AUTO_GENERATE_NAME'), True),
            max_rows_per_table=_env_int(env.get('PSTX_ASTER_MAX_ROWS_PER_TABLE'), 16, minimum=0),
            max_payload_chars=_env_int(env.get('PSTX_ASTER_MAX_PAYLOAD_CHARS'), 60000, minimum=4000),
            retry_count=_env_int(env.get('PSTX_ASTER_RETRY_COUNT'), 1, minimum=0),
            retry_backoff_seconds=_env_float(env.get('PSTX_ASTER_RETRY_BACKOFF_SECONDS'), 1.0, minimum=0.0),
            redact_paths=_env_bool(env.get('PSTX_ASTER_REDACT_PATHS'), True),
            log_file=str(env.get('PSTX_ASTER_LOG_FILE') or '').strip(),
            log_payload=_env_bool(env.get('PSTX_ASTER_LOG_PAYLOAD'), False),
        )

    def validate_live(self) -> None:
        if self.mode != 'live':
            return
        missing = []
        if not self.emp_no:
            missing.append('ASTER_EMP_NO')
        if self.backend == 'chat-flow':
            if not self.api_key:
                missing.append('ASTER_API_KEY')
        elif self.backend == 'room':
            if not self.app_id:
                missing.append('ASTER_APP_ID')
            if not self.app_secret:
                missing.append('ASTER_APP_SECRET')
            if self.validate_auth_header not in {'encrypted-time', 'bearer'}:
                raise AsterConfigError('PSTX_ASTER_VALIDATE_AUTH_HEADER 仅支持 encrypted-time 或 bearer')
        else:
            raise AsterConfigError('PSTX_ASTER_BACKEND 仅支持 chat-flow 或 room')
        if missing:
            raise AsterConfigError('Aster live 模式缺少环境变量：' + ', '.join(missing))


def _join_url(base_url: str, path: str) -> str:
    return f"{str(base_url).rstrip('/')}/{str(path).lstrip('/')}"


def _origin_from_base_url(base_url: str) -> str:
    parsed = urllib.parse.urlparse(str(base_url or ''))
    return parsed.netloc or parsed.path.split('/', 1)[0]


def _read_response_text(response) -> str:
    data = response.read()
    content_type = response.headers.get('Content-Type', '')
    encoding = 'utf-8'
    match = re.search(r'charset=([^;\s]+)', content_type, flags=re.I)
    if match:
        encoding = match.group(1)
    return data.decode(encoding, errors='replace')


def _raw_body_summary(body: object) -> dict:
    if body is None:
        return {'present': False}
    text = str(body)
    return {
        'present': True,
        'kind': 'raw',
        'chars': len(text),
        'sha256_12': _short_hash(text),
    }


def _retry_settings(context: dict) -> Tuple[int, float]:
    retry_count = _env_int(context.get('retry_count'), 0, minimum=0)
    retry_backoff = _env_float(context.get('retry_backoff_seconds'), 1.0, minimum=0.0)
    return retry_count, retry_backoff


def _retryable_aster_failure(*, status: int = 0, reason: object = '', response_body: object = '') -> bool:
    text = f'{status} {reason} {response_body}'.lower()
    transient_markers = [
        'chunkedencodingerror',
        'response ended prematurely',
        'connection reset',
        'connection aborted',
        'remote end closed connection',
        'temporarily unavailable',
        'timed out',
        'timeout',
    ]
    if any(marker in text for marker in transient_markers):
        return True
    return status in {408, 429, 500, 502, 503, 504}


def _request_json(url: str,
                  *,
                  method: str = 'POST',
                  headers: Optional[dict] = None,
                  body: Optional[dict] = None,
                  timeout: float = DEFAULT_ASTER_TIMEOUT_SECONDS,
                  context: Optional[dict] = None) -> dict:
    payload = None
    request_headers = dict(headers or {})
    if body is not None:
        payload = json.dumps(body, ensure_ascii=False).encode('utf-8')
        request_headers.setdefault('Content-Type', 'application/json;charset=utf-8')
    context = dict(context or {})
    context.setdefault('request_id', _request_id())
    context['url'] = url
    retry_count, retry_backoff = _retry_settings(context)
    max_attempts = retry_count + 1
    for attempt in range(1, max_attempts + 1):
        attempt_started_at = time.time()
        attempt_context = {
            **context,
            'attempt': attempt,
            'max_attempts': max_attempts,
        }
        _write_aster_log('request.start', {
            **attempt_context,
            'method': method,
            'url': _safe_url(url),
            'timeout_seconds': timeout,
            'headers': sanitize_for_aster_log(request_headers),
            'body': _body_summary(body, include_payload=bool(context.get('log_payload'))),
        }, log_file=str(context.get('log_file') or ''))
        request = urllib.request.Request(url, data=payload, method=method, headers=request_headers)
        try:
            with urllib.request.urlopen(request, timeout=timeout) as response:
                text = _read_response_text(response)
        except urllib.error.HTTPError as exc:
            text = exc.read().decode('utf-8', errors='replace')
            retryable = _retryable_aster_failure(status=exc.code, reason=exc.reason, response_body=text)
            diagnostics = _diagnostics(
                attempt_context,
                status=exc.code,
                reason=exc.reason,
                response_body=_safe_text_preview(text, 1200),
                retryable=retryable,
                elapsed_ms=int((time.time() - attempt_started_at) * 1000),
            )
            _write_aster_log('request.http_error', diagnostics, log_file=str(context.get('log_file') or ''))
            if retryable and attempt < max_attempts:
                delay = retry_backoff * attempt
                _write_aster_log('request.retry', {
                    **attempt_context,
                    'reason': exc.reason,
                    'status': exc.code,
                    'delay_seconds': delay,
                }, log_file=str(context.get('log_file') or ''))
                if delay > 0:
                    time.sleep(delay)
                continue
            raise AsterHttpError(f'Aster HTTP {exc.code}: {_safe_text_preview(text)}', diagnostics=diagnostics) from exc
        except urllib.error.URLError as exc:
            retryable = _retryable_aster_failure(reason=str(exc.reason))
            diagnostics = _diagnostics(
                attempt_context,
                reason=str(exc.reason),
                retryable=retryable,
                elapsed_ms=int((time.time() - attempt_started_at) * 1000),
            )
            _write_aster_log('request.url_error', diagnostics, log_file=str(context.get('log_file') or ''))
            if retryable and attempt < max_attempts:
                delay = retry_backoff * attempt
                _write_aster_log('request.retry', {
                    **attempt_context,
                    'reason': str(exc.reason),
                    'delay_seconds': delay,
                }, log_file=str(context.get('log_file') or ''))
                if delay > 0:
                    time.sleep(delay)
                continue
            raise AsterHttpError(f'Aster 网络请求失败：{exc.reason}', diagnostics=diagnostics) from exc
        try:
            parsed = json.loads(text or '{}')
            _write_aster_log('request.success', {
                **attempt_context,
                'url': _safe_url(url),
                'elapsed_ms': int((time.time() - attempt_started_at) * 1000),
                'response_keys': sorted(str(key) for key in parsed.keys()) if isinstance(parsed, dict) else [],
            }, log_file=str(context.get('log_file') or ''))
            return parsed
        except json.JSONDecodeError as exc:
            diagnostics = _diagnostics(
                attempt_context,
                response_body=_safe_text_preview(text, 1200),
                elapsed_ms=int((time.time() - attempt_started_at) * 1000),
            )
            _write_aster_log('request.json_error', diagnostics, log_file=str(context.get('log_file') or ''))
            raise AsterResponseError(f'Aster 返回不是 JSON：{_safe_text_preview(text)}', diagnostics=diagnostics) from exc
    raise AsterHttpError('Aster 请求重试耗尽且未返回结果', diagnostics=_diagnostics(context))


def _request_raw_text(url: str,
                      *,
                      method: str = 'POST',
                      headers: Optional[dict] = None,
                      body: str = '',
                      timeout: float = DEFAULT_ASTER_TIMEOUT_SECONDS,
                      context: Optional[dict] = None) -> str:
    payload = str(body or '').encode('utf-8')
    request_headers = dict(headers or {})
    context = dict(context or {})
    context.setdefault('request_id', _request_id())
    context['url'] = url
    _write_aster_log('request.start', {
        **context,
        'method': method,
        'url': _safe_url(url),
        'timeout_seconds': timeout,
        'headers': sanitize_for_aster_log(request_headers),
        'body': _raw_body_summary(body),
    }, log_file=str(context.get('log_file') or ''))
    request = urllib.request.Request(url, data=payload, method=method, headers=request_headers)
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            text = _read_response_text(response)
            _write_aster_log('request.success', {
                **context,
                'url': _safe_url(url),
                'response_chars': len(text),
                'response_sha256_12': _short_hash(text),
            }, log_file=str(context.get('log_file') or ''))
            return text
    except urllib.error.HTTPError as exc:
        text = exc.read().decode('utf-8', errors='replace')
        diagnostics = _diagnostics(
            context,
            status=exc.code,
            reason=exc.reason,
            response_body=_safe_text_preview(text, 1200),
        )
        _write_aster_log('request.http_error', diagnostics, log_file=str(context.get('log_file') or ''))
        raise AsterHttpError(f'Aster HTTP {exc.code}: {_safe_text_preview(text)}', diagnostics=diagnostics) from exc
    except urllib.error.URLError as exc:
        diagnostics = _diagnostics(context, reason=str(exc.reason))
        _write_aster_log('request.url_error', diagnostics, log_file=str(context.get('log_file') or ''))
        raise AsterHttpError(f'Aster 网络请求失败：{exc.reason}', diagnostics=diagnostics) from exc


def _request_text(url: str,
                  *,
                  method: str = 'POST',
                  headers: Optional[dict] = None,
                  body: Optional[dict] = None,
                  timeout: float = DEFAULT_ASTER_TIMEOUT_SECONDS,
                  context: Optional[dict] = None) -> str:
    payload = None
    request_headers = dict(headers or {})
    if body is not None:
        payload = json.dumps(body, ensure_ascii=False).encode('utf-8')
        request_headers.setdefault('Content-Type', 'application/json;charset=utf-8')
    context = dict(context or {})
    context.setdefault('request_id', _request_id())
    context['url'] = url
    _write_aster_log('request.start', {
        **context,
        'method': method,
        'url': _safe_url(url),
        'timeout_seconds': timeout,
        'headers': sanitize_for_aster_log(request_headers),
        'body': _body_summary(body, include_payload=bool(context.get('log_payload'))),
    }, log_file=str(context.get('log_file') or ''))
    request = urllib.request.Request(url, data=payload, method=method, headers=request_headers)
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            text = _read_response_text(response)
            _write_aster_log('request.success', {
                **context,
                'url': _safe_url(url),
                'response_chars': len(text),
                'response_sha256_12': _short_hash(text),
            }, log_file=str(context.get('log_file') or ''))
            return text
    except urllib.error.HTTPError as exc:
        text = exc.read().decode('utf-8', errors='replace')
        diagnostics = _diagnostics(
            context,
            status=exc.code,
            reason=exc.reason,
            response_body=_safe_text_preview(text, 1200),
        )
        _write_aster_log('request.http_error', diagnostics, log_file=str(context.get('log_file') or ''))
        raise AsterHttpError(f'Aster HTTP {exc.code}: {_safe_text_preview(text)}', diagnostics=diagnostics) from exc
    except urllib.error.URLError as exc:
        diagnostics = _diagnostics(context, reason=str(exc.reason))
        _write_aster_log('request.url_error', diagnostics, log_file=str(context.get('log_file') or ''))
        raise AsterHttpError(f'Aster 网络请求失败：{exc.reason}', diagnostics=diagnostics) from exc


def _query(params: dict) -> str:
    values = {
        key: value
        for key, value in params.items()
        if value is not None and value != ''
    }
    encoded = urllib.parse.urlencode(values)
    return f'?{encoded}' if encoded else ''


class AsterFlowClient:
    def __init__(self, config: AsterConfig):
        self.config = config

    def _context(self, operation: str) -> dict:
        return {
            'request_id': _request_id(),
            'operation': operation,
            'backend': self.config.backend,
            'base_url': _safe_url(self.config.base_url),
            'emp_no': self.config.emp_no,
            'api_key': self.config.api_key,
            'log_file': self.config.log_file,
            'log_payload': self.config.log_payload,
            'retry_count': self.config.retry_count,
            'retry_backoff_seconds': self.config.retry_backoff_seconds,
        }

    def _agent_context(self, inputs: Optional[dict]) -> dict:
        data = inputs or {}
        context = {}
        for key in ('agent_run_id', 'agent_profile', 'step_index', 'tool_count', 'retry'):
            if key in data:
                context[key] = data.get(key)
        return context

    def chat_flow(self, *, query: str, inputs: Optional[dict] = None) -> dict:
        url = _join_url(
            self.config.base_url,
            '/aster/flow-api/run/chat-flow' + _query({
                'apiKey': self.config.api_key,
                'empNo': self.config.emp_no,
            }),
        )
        body = {
            'query': query,
            'inputs': inputs or {},
            'conversationId': self.config.conversation_id,
            'files': None,
            'autoGenerateName': self.config.auto_generate_name,
        }
        context = {
            **self._context('chat_flow'),
            **self._agent_context(inputs),
        }
        response = _request_json(
            url,
            body=body,
            timeout=self.config.timeout_seconds,
            context=context,
        )
        if response.get('code') != 200 or not response.get('data'):
            preview = _safe_json_preview(response, 1200)
            diagnostics = _diagnostics(
                context,
                response_body=preview,
            )
            _write_aster_log('request.response_error', diagnostics, log_file=self.config.log_file)
            raise AsterResponseError(
                f'Aster ChatFlow 返回异常：{_safe_json_preview(response)}',
                diagnostics=diagnostics,
            )
        return response['data']


def _cipher_key(key: str) -> bytes:
    if not key:
        raise AsterConfigError('Aster cipher key 不能为空')
    return (('*' * 32) + key)[-32:].encode('utf-8')


def _pkcs7_pad(data: bytes, block_size: int = 16) -> bytes:
    pad_len = block_size - (len(data) % block_size)
    return data + bytes([pad_len]) * pad_len


def _pkcs7_unpad(data: bytes, block_size: int = 16) -> bytes:
    if not data:
        raise AsterResponseError('Aster 密文解密结果为空')
    pad_len = data[-1]
    if pad_len < 1 or pad_len > block_size or data[-pad_len:] != bytes([pad_len]) * pad_len:
        raise AsterResponseError('Aster 密文 PKCS7 padding 无效')
    return data[:-pad_len]


def _aes_ecb_encrypt(key: bytes, data: bytes) -> bytes:
    try:
        from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes  # type: ignore
        encryptor = Cipher(algorithms.AES(key), modes.ECB()).encryptor()
        return encryptor.update(data) + encryptor.finalize()
    except ModuleNotFoundError:
        pass
    try:
        from Crypto.Cipher import AES  # type: ignore
        return AES.new(key, AES.MODE_ECB).encrypt(data)
    except ModuleNotFoundError:
        pass
    try:
        from Cryptodome.Cipher import AES  # type: ignore
        return AES.new(key, AES.MODE_ECB).encrypt(data)
    except ModuleNotFoundError as exc:
        raise AsterConfigError('Aster room/auth 需要安装 cryptography、pycryptodome 或 pycryptodomex') from exc


def _aes_ecb_decrypt(key: bytes, data: bytes) -> bytes:
    try:
        from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes  # type: ignore
        decryptor = Cipher(algorithms.AES(key), modes.ECB()).decryptor()
        return decryptor.update(data) + decryptor.finalize()
    except ModuleNotFoundError:
        pass
    try:
        from Crypto.Cipher import AES  # type: ignore
        return AES.new(key, AES.MODE_ECB).decrypt(data)
    except ModuleNotFoundError:
        pass
    try:
        from Cryptodome.Cipher import AES  # type: ignore
        return AES.new(key, AES.MODE_ECB).decrypt(data)
    except ModuleNotFoundError as exc:
        raise AsterConfigError('Aster room/auth 需要安装 cryptography、pycryptodome 或 pycryptodomex') from exc


def aster_encrypt(key: str, content: str) -> str:
    encrypted = _aes_ecb_encrypt(_cipher_key(key), _pkcs7_pad(content.encode('utf-8')))
    first_base64 = base64.b64encode(encrypted).decode('ascii')
    return base64.b64encode(first_base64.encode('utf-8')).decode('ascii')


def aster_decrypt(key: str, ciphertext: str) -> str:
    first_base64 = base64.b64decode(ciphertext).decode('utf-8')
    encrypted = base64.b64decode(first_base64)
    decrypted = _aes_ecb_decrypt(_cipher_key(key), encrypted)
    return _pkcs7_unpad(decrypted).decode('utf-8')


def _aster_datetime() -> str:
    return datetime.now().strftime('%Y-%m-%d %H:%M:%S')


class AsterAuthClient:
    def __init__(self, config: AsterConfig):
        self.config = config
        self.access_token = ''
        self.expires_at = 0.0

    def _context(self, operation: str) -> dict:
        return {
            'request_id': _request_id(),
            'operation': operation,
            'backend': self.config.backend,
            'base_url': _safe_url(self.config.base_url),
            'app_id': self.config.app_id,
            'app_secret': self.config.app_secret,
            'emp_no': self.config.emp_no,
            'origin': self._origin(),
            'log_file': self.config.log_file,
            'log_payload': self.config.log_payload,
        }

    def _origin(self) -> str:
        return self.config.origin or _origin_from_base_url(self.config.base_url)

    def _ciphertext(self) -> str:
        return aster_encrypt(self.config.app_id, json.dumps({
            'accessToken': self.access_token or '',
            'deviceId': self.config.device_id,
            'cipherTimeStr': _aster_datetime(),
        }, ensure_ascii=False, separators=(',', ':')))

    def get_access_token(self, *, force: bool = False) -> str:
        if not force and self.access_token and time.time() + 1800 < self.expires_at:
            return self.access_token
        url = _join_url(self.config.base_url, '/auth/api/v1/generateOrProlongToken')
        context = self._context('auth_generate_or_prolong_token')
        response = _request_json(url, body={
            'appId': self.config.app_id,
            'appSecret': self.config.app_secret,
            'empNo': self.config.emp_no,
            'ciphertext': self._ciphertext(),
        }, timeout=self.config.timeout_seconds, context=context)
        if response.get('code') != 200 or not response.get('data'):
            preview = _safe_json_preview(response, 1200)
            diagnostics = _diagnostics(
                context,
                response_body=preview,
            )
            _write_aster_log('request.response_error', diagnostics, log_file=self.config.log_file)
            raise AsterResponseError(
                f'Aster auth 返回异常：{_safe_json_preview(response)}',
                diagnostics=diagnostics,
            )
        try:
            data = json.loads(aster_decrypt(self.config.app_id, response['data']))
        except (binascii.Error, json.JSONDecodeError, ValueError, TypeError, UnicodeDecodeError) as exc:
            diagnostics = _diagnostics(context, response_body='token decrypt/parse failed')
            _write_aster_log('request.response_error', diagnostics, log_file=self.config.log_file)
            raise AsterResponseError('Aster auth token 解密或解析失败', diagnostics=diagnostics) from exc
        self.access_token = data.get('accessToken') or ''
        if not self.access_token:
            diagnostics = _diagnostics(context, response_body='missing accessToken')
            _write_aster_log('request.response_error', diagnostics, log_file=self.config.log_file)
            raise AsterResponseError('Aster auth 未返回 accessToken', diagnostics=diagnostics)
        self.expires_at = time.time() + float(data.get('validityPeriodSeconds') or 7200)
        _write_aster_log('auth.token_acquired', {
            **self._context('auth_token_acquired'),
            'access_token': self.access_token,
            'validity_seconds': data.get('validityPeriodSeconds') or 7200,
        }, log_file=self.config.log_file)
        return self.access_token

    def _validate_headers(self, token: str, origin: str) -> dict:
        authorization = f'Bearer {token}'
        if self.config.validate_auth_header == 'encrypted-time':
            authorization = aster_encrypt(origin, _aster_datetime())
        return {
            'Authorization': authorization,
            'aigc-origin': origin,
            'Origin': origin,
            'Content-Type': 'text/plain;charset=utf-8',
        }

    def validate_access_token(self, token: str) -> dict:
        origin = self._origin()
        if not origin:
            raise AsterConfigError('Room token 校验需要 ASTER_ORIGIN，或可从固定 Aster 服务地址推导出域名')
        url = _join_url(self.config.base_url, '/auth/js-sdk/validateAccessToken')
        context = self._context('auth_validate_access_token')
        body = aster_encrypt(origin, json.dumps({
            'appId': self.config.app_id,
            'accessTokenRaw': f'Bearer {token}',
            'deviceId': self.config.device_id,
        }, ensure_ascii=False, separators=(',', ':')))
        text = _request_raw_text(
            url,
            method='POST',
            headers=self._validate_headers(token, origin),
            body=body,
            timeout=self.config.timeout_seconds,
            context=context,
        )
        try:
            parsed = json.loads(aster_decrypt(origin, text))
        except (binascii.Error, json.JSONDecodeError, ValueError, TypeError, UnicodeDecodeError) as exc:
            diagnostics = _diagnostics(context, response_body='validate response decrypt/parse failed')
            _write_aster_log('request.response_error', diagnostics, log_file=self.config.log_file)
            raise AsterResponseError('Aster validateAccessToken 响应解密或解析失败', diagnostics=diagnostics) from exc
        _write_aster_log('auth.token_validated', {
            **context,
            'code': parsed.get('code'),
            'is_valid': (parsed.get('data') or {}).get('isValid') if isinstance(parsed.get('data'), dict) else None,
            'status_code': (parsed.get('data') or {}).get('statusCode') if isinstance(parsed.get('data'), dict) else None,
        }, log_file=self.config.log_file)
        return parsed

    def ensure_valid_access_token(self) -> str:
        token = self.get_access_token()
        if not self.config.validate_token:
            return token
        result = self.validate_access_token(token)
        if self._token_validation_passed(result):
            self._update_expiry_from_validation(result)
            return token
        _write_aster_log('auth.token_invalid_retry', {
            **self._context('auth_token_invalid_retry'),
            'validation_result': result,
        }, log_file=self.config.log_file)
        token = self.get_access_token(force=True)
        result = self.validate_access_token(token)
        if self._token_validation_passed(result):
            self._update_expiry_from_validation(result)
            return token
        data = result.get('data') if isinstance(result, dict) else {}
        diagnostics = _diagnostics(
            self._context('auth_validate_access_token'),
            status_code=data.get('statusCode') if isinstance(data, dict) else None,
            response_body=_safe_json_preview(result, 1200),
        )
        _write_aster_log('request.response_error', diagnostics, log_file=self.config.log_file)
        raise AsterResponseError('Aster accessToken 校验失败，请检查员工号、设备 ID、App 权限或 API 调用开关', diagnostics=diagnostics)

    @staticmethod
    def _token_validation_passed(result: dict) -> bool:
        data = result.get('data') if isinstance(result, dict) else None
        if not isinstance(data, dict):
            return False
        return bool(data.get('isValid')) and str(data.get('statusCode')) == '1'

    def _update_expiry_from_validation(self, result: dict) -> None:
        data = result.get('data') if isinstance(result, dict) else {}
        try:
            validity = float(data.get('validityPeriodSeconds') or 0) if isinstance(data, dict) else 0
        except (TypeError, ValueError):
            validity = 0
        if validity > 0:
            self.expires_at = time.time() + validity


class AsterRoomClient:
    def __init__(self, config: AsterConfig):
        self.config = config
        self.auth_client = AsterAuthClient(config)

    def _context(self, operation: str) -> dict:
        return {
            'request_id': _request_id(),
            'operation': operation,
            'backend': self.config.backend,
            'base_url': _safe_url(self.config.base_url),
            'app_id': self.config.app_id,
            'emp_no': self.config.emp_no,
            'log_file': self.config.log_file,
            'log_payload': self.config.log_payload,
        }

    def _headers(self, extra: Optional[dict] = None) -> dict:
        token = self.auth_client.ensure_valid_access_token()
        headers = {'Authorization': f'Bearer {token}'}
        headers.update(extra or {})
        return headers

    def create_room(self, name: str = 'PSTX审查') -> dict:
        url = _join_url(self.config.base_url, '/aster/room/create')
        url = f'{url}?{urllib.parse.urlencode({"name": name[:10]})}'
        context = self._context('room_create')
        response = _request_json(
            url,
            method='POST',
            headers=self._headers(),
            timeout=self.config.timeout_seconds,
            context=context,
        )
        if response.get('code') != 200 or not response.get('data', {}).get('roomId'):
            preview = _safe_json_preview(response, 1200)
            diagnostics = _diagnostics(
                context,
                response_body=preview,
            )
            _write_aster_log('request.response_error', diagnostics, log_file=self.config.log_file)
            raise AsterResponseError(
                f'Aster createRoom 返回异常：{_safe_json_preview(response)}',
                diagnostics=diagnostics,
            )
        return response['data']

    def send_question(self, *, room_id: str, content: str) -> dict:
        text = _request_text(
            _join_url(self.config.base_url, '/aster/room/chat/stream/send'),
            method='POST',
            headers=self._headers({'Content-Type': 'application/json;charset=utf-8'}),
            body={
                'roomId': room_id,
                'isWithContext': False,
                'content': content,
                'isTmpChat': False,
                'sourceIds': [],
                'imageList': [],
                'isStream': True,
            },
            timeout=self.config.timeout_seconds,
            context=self._context('room_chat_stream_send'),
        )
        answer = ''
        refs = None
        raw_chunks = []
        for chunk in _parse_jsonish_stream(text):
            raw_chunks.append(chunk)
            data = chunk.get('data') if isinstance(chunk, dict) else None
            message = data.get('data') if isinstance(data, dict) and isinstance(data.get('data'), dict) else data
            if not isinstance(message, dict):
                continue
            content_text = message.get('content') or ''
            if content_text:
                answer = content_text if str(content_text).startswith(answer) else answer + str(content_text)
            if message.get('refs') is not None:
                refs = message.get('refs')
        return {'answer': answer, 'refs': refs, 'rawChunks': raw_chunks, 'roomId': room_id}


def _parse_jsonish_stream(text: str) -> Iterable[dict]:
    for raw_line in str(text or '').splitlines():
        line = raw_line.strip()
        if not line or line == '[DONE]':
            continue
        if line.startswith('data:'):
            line = line[5:].strip()
        if not line or line == '[DONE]':
            continue
        try:
            parsed = json.loads(line)
        except json.JSONDecodeError:
            parsed = {'text': line}
        if isinstance(parsed, dict):
            yield parsed


PATH_FIELD_NAMES = {'文件', 'filename', 'project_root', 'PRIM_FILE', '路径'}
PATH_PATTERN = re.compile(r'([A-Za-z]:[\\/][^\s,;]+|/(?:Users|home|mnt|Volumes)/[^\s,;]+)')


def _truncate_text(value: object, limit: int = 240) -> object:
    if not isinstance(value, str):
        return value
    if len(value) <= limit:
        return value
    return value[:limit] + f'...<截断 {len(value) - limit} 字符>'


def _redact_paths(text: str) -> str:
    def repl(match: re.Match) -> str:
        path = match.group(0)
        leaf = re.split(r'[\\/]', path.rstrip('/\\'))[-1] or 'path'
        return f'[path:{leaf}]'
    return PATH_PATTERN.sub(repl, text)


def _sanitize_value(key: str, value: object, *, redact_paths: bool) -> object:
    if isinstance(value, dict):
        return {
            str(child_key): _sanitize_value(str(child_key), child_value, redact_paths=redact_paths)
            for child_key, child_value in value.items()
        }
    if isinstance(value, list):
        return [_sanitize_value(key, item, redact_paths=redact_paths) for item in value[:20]]
    if not isinstance(value, str):
        return value
    text = value
    if redact_paths or key in PATH_FIELD_NAMES:
        text = _redact_paths(text)
    return _truncate_text(text)


def _compact_rows(rows: list, *, max_rows: int, redact_paths: bool) -> list:
    compacted = []
    for row in rows[:max_rows]:
        if isinstance(row, dict):
            compacted.append({
                str(key): _sanitize_value(str(key), value, redact_paths=redact_paths)
                for key, value in row.items()
            })
        else:
            compacted.append(_sanitize_value('', row, redact_paths=redact_paths))
    return compacted


def _iter_report_tables(report: dict) -> Iterable[Tuple[dict, dict]]:
    for section in report.get('sections', []) or []:
        for table in section.get('tables', []) or []:
            yield section, table


def _find_report_table(report: dict, table_id: str) -> Optional[dict]:
    for _, table in _iter_report_tables(report):
        if table.get('id') == table_id:
            return table
    return None


def _report_table_count(report: dict, table_id: str) -> int:
    table = _find_report_table(report, table_id)
    try:
        return int((table or {}).get('count') or 0)
    except (TypeError, ValueError):
        return 0


def _metric_lookup(report: dict) -> Dict[str, object]:
    return {
        str(item.get('label') or ''): item.get('value')
        for item in report.get('metrics', []) or []
        if isinstance(item, dict)
    }


def _review_scope_item(name: str,
                       target: str,
                       table_ids: List[str],
                       report: dict,
                       evidence: str,
                       *,
                       high_signal: bool = False) -> dict:
    count = sum(_report_table_count(report, table_id) for table_id in table_ids)
    if count:
        status = 'needs_review' if high_signal else 'covered_with_findings'
    else:
        status = 'covered_no_findings'
    return {
        'item': name,
        'target': target,
        'status': status,
        'count': count,
        'tables': table_ids,
        'evidence': evidence,
    }


def _build_review_scope(report: dict) -> List[dict]:
    metrics = _metric_lookup(report)
    rows = [
        _review_scope_item(
            'BOM 与装配状态',
            'bom',
            ['bom_normal_merged', 'bom_depop_merged', 'bom_option_components', 'bom_option_circle_coverage', 'bom_option_circle_issues'],
            report,
            f"贴装种类 {metrics.get('贴装种类', 0)}，DEPOP 总数 {metrics.get('DEPOP 总数', 0)}，BOM_OPTION 打圈问题 {metrics.get('BOM圈问题', 0)}，include_depop={report.get('include_depop')}",
        ),
        _review_scope_item(
            '网络分类与页码映射',
            'network',
            ['power_net_rows', 'gnd_net_rows', 'diff_pair_rows', 'single_node_rows', 'page_mapping_rows'],
            report,
            f"网络总数 {metrics.get('网络总数', 0)}，重点关注电源/GND/差分对/单节点和逻辑页-真实页映射",
        ),
        _review_scope_item(
            '属性与命名 DRC',
            'drc',
            ['missing_hq_code', 'missing_value', 'missing_package', 'tbd_attrs', 'single_pin_nets', 'unnamed_nets', 'bom_option_typos'],
            report,
            f"DRC 总数 {metrics.get('DRC 总数', 0)}，覆盖缺料号、缺 VALUE、缺封装、TBD、单端网络和未命名网络",
            high_signal=True,
        ),
        _review_scope_item(
            '电阻偏置、串阻与芯片 Pin 状态',
            'resistor',
            ['divider_risks', 'dup_pullups', 'dup_pulldowns', 'od_missing', 'chip_pin_rows'],
            report,
            f"电阻候选 {metrics.get('电阻候选', 0)}，无法判断 {metrics.get('电阻无法判断', 0)}，覆盖串阻分压、重复上下拉、OD/OC 和芯片 Pin",
            high_signal=True,
        ),
        _review_scope_item(
            '电容降额与电压推断边界',
            'derating',
            ['derating'],
            report,
            f"降额不合格 {metrics.get('降额不合格', 0)}，阈值 {report.get('ratio_limit', '')}%，注意候选电压推断不能当作确定结论",
            high_signal=True,
        ),
        _review_scope_item(
            'CSA 几何规范',
            'csa',
            ['csa_summary_rows', 'csa_dot_cross_rows', 'csa_circle_rows'],
            report,
            f"规范候选 {metrics.get('规范候选', 0)}，覆盖 DOT 四向十字交叉、画圈对象和页级汇总",
            high_signal=True,
        ),
    ]
    return rows


def _key_finding_severity(section_id: str, table_id: str) -> str:
    high = {
        'missing_hq_code',
        'missing_value',
        'missing_package',
        'single_pin_nets',
        'divider_risks',
        'derating',
        'csa_dot_cross_rows',
    }
    medium = {
        'bom_option_components',
        'bom_option_circle_coverage',
        'bom_option_circle_issues',
        'tbd_attrs',
        'unnamed_nets',
        'bom_option_typos',
        'dup_pullups',
        'dup_pulldowns',
        'od_missing',
        'chip_pin_rows',
        'csa_circle_rows',
        'page_mapping_rows',
    }
    if table_id in high:
        return 'high'
    if table_id in medium or section_id in {'network', 'resistor', 'csa'}:
        return 'medium'
    return 'low'


def _build_key_findings(report: dict, config: AsterConfig) -> List[dict]:
    findings = []
    for section, table in _iter_report_tables(report):
        count = _report_table_count(report, str(table.get('id') or ''))
        if count <= 0:
            continue
        findings.append({
            'section': section.get('title') or section.get('id') or '',
            'target': section.get('id') or '',
            'table_id': table.get('id') or '',
            'table': table.get('title') or table.get('id') or '',
            'count': count,
            'severity_hint': _key_finding_severity(str(section.get('id') or ''), str(table.get('id') or '')),
            'kind_counts': table.get('kind_counts', {}),
            'sample_rows': _compact_rows(
                table.get('rows', []) or [],
                max_rows=min(max(config.max_rows_per_table, 0), 4),
                redact_paths=config.redact_paths,
            ),
        })
    severity_order = {'high': 0, 'medium': 1, 'low': 2}
    findings.sort(key=lambda item: (severity_order.get(item.get('severity_hint'), 3), -int(item.get('count') or 0)))
    return findings[:18]


def _build_manual_review_boundaries(report: dict) -> List[dict]:
    include_depop = report.get('include_depop')
    return [
        {
            'topic': '电平/电压推断',
            'target': 'derating',
            'boundary': '仅凭网络名 token 不得下确定结论；例如 PG_P1V8 只能作为候选，需结合真实电源、上拉/下拉和器件特性。',
        },
        {
            'topic': 'OD/OC 与上下拉',
            'target': 'resistor',
            'boundary': 'OD/OC 只能基于 pin 名、网络关系和外部偏置提出候选；芯片手册特性未知时必须要求人工确认。',
        },
        {
            'topic': 'AC 耦合与差分对',
            'target': 'derating',
            'boundary': '差分对需成对出现并结合 _P/_N 命名和镜像连接判断，单独 _N 或 _P 不应直接判定为差分 AC 耦合。',
        },
        {
            'topic': 'DEPOP/DNP 参与分析边界',
            'target': 'drc',
            'boundary': f'当前 include_depop={include_depop}；若 DEPOP 被排除，AI 不应把被排除器件当作实际连通器件下确定结论。',
        },
        {
            'topic': 'CSA 几何候选',
            'target': 'csa',
            'boundary': 'DOT 十字交叉和画圈对象是几何规范候选，不等价于网络短接或必然设计错误。',
        },
    ]


def build_report_brief(report: dict, config: AsterConfig) -> dict:
    brief = {
        'project_name': report.get('project_name'),
        'generated_at': report.get('generated_at'),
        'ratio_limit': report.get('ratio_limit'),
        'include_depop': report.get('include_depop'),
        'metrics': report.get('metrics', []),
        'top_insights': report.get('top_insights', [])[:6],
        'section_cards': report.get('section_cards', [])[:12],
        'summary_lines': report.get('summary_lines', []),
        'warnings': [
            _sanitize_value('warning', item, redact_paths=config.redact_paths)
            for item in (report.get('warnings') or [])[:12]
        ],
        'review_scope': _build_review_scope(report),
        'key_findings': _build_key_findings(report, config),
        'manual_review_boundaries': _build_manual_review_boundaries(report),
        'sections': [],
    }
    for section in report.get('sections', []) or []:
        section_brief = {
            'id': section.get('id'),
            'title': section.get('title'),
            'total_rows': section.get('total_rows', 0),
            'tables': [],
        }
        for table in section.get('tables', []) or []:
            section_brief['tables'].append({
                'id': table.get('id'),
                'title': table.get('title'),
                'count': table.get('count', 0),
                'kind_counts': table.get('kind_counts', {}),
                'columns': table.get('columns', []),
                'sample_rows': _compact_rows(
                    table.get('rows', []) or [],
                    max_rows=config.max_rows_per_table,
                    redact_paths=config.redact_paths,
                ),
            })
        brief['sections'].append(section_brief)

    encoded = json.dumps(brief, ensure_ascii=False)
    if len(encoded) <= config.max_payload_chars:
        return brief

    for section in brief['sections']:
        for table in section.get('tables', []):
            table['sample_rows'] = []
    brief['truncation_note'] = (
        f'报告摘要超过 {config.max_payload_chars} 字符，已移除表格样例行，仅保留计数和列信息。'
    )
    return brief


def build_aster_prompt(report_brief: dict) -> str:
    report_json = json.dumps(report_brief, ensure_ascii=False, indent=2)
    return f"""你是硬件原理图审查助手，请基于下面的 PSTX 审查报告摘要生成工程审查建议。

要求：
1. 只输出一个 JSON 对象，不要输出 Markdown，不要输出代码块。
2. 不要编造报告中没有的数据；无法确定时写“需人工确认”。
3. priorities 最多 5 条，severity 只能是 high、medium、low。
4. target 只能从 bom、network、drc、csa、resistor、derating、summary 中选择。
5. section_focus 按建议优先级排序，rows 使用报告中的行数。
6. review_checklist 必须覆盖 BOM/DEPOP、网络/页码映射、DRC、芯片 Pin/电阻、降额、CSA 规范这些审查域。
7. manual_review 用来列出不能自动下结论、必须人工确认的边界，不要把候选判断写成确定结论。

JSON schema:
{{
  "summary": "一句到三句话的中文总览",
  "priorities": [
    {{"title": "建议标题", "body": "建议依据和处理方式", "target": "drc", "severity": "high"}}
  ],
  "section_focus": [
    {{"section": "设计检查", "target": "drc", "rows": 12, "reason": "为什么优先看"}}
  ],
  "review_checklist": [
    {{"item": "BOM 与装配状态", "status": "needs_review", "evidence": "为什么需要看", "target": "bom", "severity": "medium"}}
  ],
  "manual_review": [
    {{"topic": "电平/电压推断", "reason": "为什么必须人工确认", "target": "derating"}}
  ]
}}

报告摘要：
{report_json}
"""


def _extract_balanced_json(text: str) -> Optional[dict]:
    content = str(text or '').strip()
    fence = re.search(r'```(?:json)?\s*(.*?)```', content, flags=re.S | re.I)
    if fence:
        content = fence.group(1).strip()
    start = content.find('{')
    if start < 0:
        return None
    depth = 0
    in_string = False
    escape = False
    for index in range(start, len(content)):
        char = content[index]
        if in_string:
            if escape:
                escape = False
            elif char == '\\':
                escape = True
            elif char == '"':
                in_string = False
            continue
        if char == '"':
            in_string = True
        elif char == '{':
            depth += 1
        elif char == '}':
            depth -= 1
            if depth == 0:
                candidate = content[start:index + 1]
                try:
                    parsed = json.loads(candidate)
                except json.JSONDecodeError:
                    return None
                return parsed if isinstance(parsed, dict) else None
    return None


def _normalize_priorities(value: object, fallback: List[dict]) -> List[dict]:
    if not isinstance(value, list):
        return fallback
    rows = []
    allowed_targets = {'bom', 'network', 'drc', 'csa', 'resistor', 'derating', 'summary'}
    allowed_severity = {'high', 'medium', 'low'}
    for item in value[:5]:
        if not isinstance(item, dict):
            continue
        target = str(item.get('target') or 'summary').strip()
        severity = str(item.get('severity') or 'medium').strip().lower()
        rows.append({
            'title': str(item.get('title') or '建议').strip()[:80],
            'body': str(item.get('body') or '').strip()[:600],
            'target': target if target in allowed_targets else 'summary',
            'severity': severity if severity in allowed_severity else 'medium',
        })
    return rows or fallback


def _normalize_section_focus(value: object, fallback: List[dict]) -> List[dict]:
    if not isinstance(value, list):
        return fallback
    rows = []
    for item in value[:8]:
        if not isinstance(item, dict):
            continue
        rows.append({
            'section': str(item.get('section') or '分区').strip()[:80],
            'target': str(item.get('target') or 'summary').strip()[:40],
            'rows': item.get('rows', 0),
            'reason': str(item.get('reason') or '').strip()[:500],
        })
    return rows or fallback


def _normalize_review_checklist(value: object, fallback: List[dict]) -> List[dict]:
    if not isinstance(value, list):
        return fallback
    rows = []
    allowed_targets = {'bom', 'network', 'drc', 'csa', 'resistor', 'derating', 'summary'}
    allowed_severity = {'high', 'medium', 'low'}
    allowed_status = {'pass', 'covered_no_findings', 'covered_with_findings', 'needs_review', 'manual_only'}
    for item in value[:10]:
        if not isinstance(item, dict):
            continue
        target = str(item.get('target') or 'summary').strip()
        severity = str(item.get('severity') or 'medium').strip().lower()
        status = str(item.get('status') or 'needs_review').strip().lower()
        rows.append({
            'item': str(item.get('item') or item.get('title') or '审查项').strip()[:80],
            'status': status if status in allowed_status else 'needs_review',
            'evidence': str(item.get('evidence') or item.get('body') or '').strip()[:700],
            'target': target if target in allowed_targets else 'summary',
            'severity': severity if severity in allowed_severity else 'medium',
        })
    return rows or fallback


def _normalize_manual_review(value: object, fallback: List[dict]) -> List[dict]:
    if not isinstance(value, list):
        return fallback
    rows = []
    allowed_targets = {'bom', 'network', 'drc', 'csa', 'resistor', 'derating', 'summary'}
    for item in value[:8]:
        if not isinstance(item, dict):
            continue
        target = str(item.get('target') or 'summary').strip()
        rows.append({
            'topic': str(item.get('topic') or item.get('title') or '人工复核项').strip()[:80],
            'reason': str(item.get('reason') or item.get('boundary') or '').strip()[:700],
            'target': target if target in allowed_targets else 'summary',
        })
    return rows or fallback


def normalize_aster_answer(answer: str,
                           *,
                           report: dict,
                           bundle: dict,
                           config: AsterConfig,
                           metadata: Optional[dict] = None,
                           request_payload_chars: int = 0) -> dict:
    fallback = build_aster_mock_summary(report, bundle)
    parsed = _extract_balanced_json(answer)
    summary = ''
    if parsed and isinstance(parsed.get('summary'), str):
        summary = parsed['summary'].strip()
    if not summary:
        summary = str(answer or '').strip()
    if not summary:
        summary = 'Aster 已返回空摘要，请检查对应 ChatFlow/Room 的输出配置。'
    payload = {
        'ok': True,
        'mode': 'live',
        'provider': f'aster-{config.backend}',
        'project_name': report.get('project_name') or bundle.get('project_name') or '未命名项目',
        'summary': summary[:1200],
        'priorities': _normalize_priorities(parsed.get('priorities') if parsed else None, fallback['priorities']),
        'section_focus': _normalize_section_focus(parsed.get('section_focus') if parsed else None, fallback['section_focus']),
        'review_checklist': _normalize_review_checklist(
            parsed.get('review_checklist') if parsed else None,
            fallback.get('review_checklist', []),
        ),
        'manual_review': _normalize_manual_review(
            parsed.get('manual_review') if parsed else None,
            fallback.get('manual_review', []),
        ),
        'safeguards': [
            '当前为 live 模式，由后端请求真实 Aster，前端不接触 appSecret、apiKey、accessToken。',
            '发送内容为后端裁剪后的报告摘要，不上传原始 PSTX 文件。',
            '如果 Aster 不可达，可将 PSTX_ASTER_MODE 切回 mock 保持本地报告可用。',
        ],
        'metadata': {
            'backend': config.backend,
            'conversation_id': metadata.get('conversation_id') if metadata else '',
            'message_id': metadata.get('message_id') if metadata else '',
            'task_id': metadata.get('task_id') if metadata else '',
            'request_payload_chars': request_payload_chars,
            'answer_format': 'json' if parsed else 'text',
        },
    }
    return payload


def ask_aster_live_model(prompt: str,
                         *,
                         inputs: Optional[dict] = None,
                         environ: Optional[dict] = None) -> dict:
    config = AsterConfig.from_env(environ)
    live_config = AsterConfig(**{**config.__dict__, 'mode': 'live'})
    live_config.validate_live()
    request_payload_chars = len(json.dumps({
        'prompt': prompt,
        'inputs': inputs or {},
    }, ensure_ascii=False))

    if live_config.backend == 'chat-flow':
        data = AsterFlowClient(live_config).chat_flow(
            query=prompt,
            inputs=inputs or {},
        )
        answer = str(data.get('answer') or data.get('text') or '')
        return {
            'ok': True,
            'mode': 'live',
            'provider': 'aster-chat-flow',
            'answer': answer,
            'metadata': {
                'backend': live_config.backend,
                'conversation_id': data.get('conversation_id') or data.get('conversationId') or '',
                'message_id': data.get('message_id') or data.get('messageId') or data.get('id') or '',
                'task_id': data.get('task_id') or data.get('taskId') or '',
                'request_payload_chars': request_payload_chars,
            },
            'raw': data,
        }

    room_client = AsterRoomClient(live_config)
    room = room_client.create_room('PSTX审查')
    data = room_client.send_question(room_id=room['roomId'], content=prompt)
    return {
        'ok': True,
        'mode': 'live',
        'provider': 'aster-room',
        'answer': str(data.get('answer') or ''),
        'metadata': {
            'backend': live_config.backend,
            'conversation_id': data.get('roomId') or room.get('roomId') or '',
            'request_payload_chars': request_payload_chars,
        },
        'raw': data,
    }


def build_aster_live_summary(report: dict, bundle: dict, *, environ: Optional[dict] = None) -> dict:
    config = AsterConfig.from_env(environ)
    live_config = AsterConfig(**{**config.__dict__, 'mode': 'live'})
    brief = build_report_brief(report, live_config)
    prompt = build_aster_prompt(brief)
    model_payload = ask_aster_live_model(
        prompt,
        inputs={'project_name': report.get('project_name') or bundle.get('project_name') or ''},
        environ=environ,
    )
    metadata = dict(model_payload.get('metadata') or {})
    return normalize_aster_answer(
        str(model_payload.get('answer') or ''),
        report=report,
        bundle=bundle,
        config=live_config,
        metadata=metadata,
        request_payload_chars=int(metadata.get('request_payload_chars') or 0),
    )
