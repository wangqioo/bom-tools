# -*- coding: utf-8 -*-
"""Aster summary mode switch for Web/API callers."""

from __future__ import annotations

import os
import threading
import urllib.parse
from typing import Dict, Optional, Tuple

from pstx_integrations.aster.client import (
    ASTER_FIXED_BASE_URL,
    AsterConfig,
    AsterConfigError,
    AsterError,
    AsterHttpError,
    AsterResponseError,
    ask_aster_live_model,
    build_aster_live_summary,
    get_aster_log_path,
)
from pstx_integrations.aster.mock import build_aster_mock_summary


SECRET_ENV_NAMES = {'ASTER_API_KEY', 'ASTER_APP_SECRET'}
RUNTIME_ENV_NAMES = {
    'PSTX_ASTER_MODE',
    'PSTX_ASTER_BACKEND',
    'ASTER_EMP_NO',
    'ASTER_API_KEY',
    'ASTER_ORIGIN',
    'ASTER_APP_ID',
    'ASTER_APP_SECRET',
}
RUNTIME_FIELD_MAP = {
    'mode': 'PSTX_ASTER_MODE',
    'backend': 'PSTX_ASTER_BACKEND',
    'emp_no': 'ASTER_EMP_NO',
    'api_key': 'ASTER_API_KEY',
    'origin': 'ASTER_ORIGIN',
    'app_id': 'ASTER_APP_ID',
    'app_secret': 'ASTER_APP_SECRET',
}
_RUNTIME_LOCK = threading.RLock()
_RUNTIME_OVERRIDES: Dict[str, str] = {}


def _merged_env(environ: Optional[dict] = None) -> dict:
    env = dict(os.environ if environ is None else environ)
    with _RUNTIME_LOCK:
        env.update(_RUNTIME_OVERRIDES)
    return env


def _env_present(env: dict, name: str) -> bool:
    return bool(str(env.get(name) or '').strip())


def _safe_origin(value: str) -> str:
    if not value:
        return ''
    parsed = urllib.parse.urlparse(value)
    return parsed.netloc or value.split('/', 1)[0].split('?', 1)[0]


def _status_item(env: dict,
                 name: str,
                 label: str,
                 *,
                 required: bool,
                 secret: bool = False,
                 source: str = 'environment',
                 safe_value: str = '') -> dict:
    item = {
        'name': name,
        'label': label,
        'required': required,
        'configured': _env_present(env, name),
        'secret': secret,
        'source': source,
    }
    if safe_value and not secret:
        item['value'] = safe_value
    return item


def set_aster_runtime_config(values: dict) -> dict:
    if not isinstance(values, dict):
        raise AsterConfigError('Aster runtime 配置必须是 JSON 对象')
    updates = {}
    for field, env_name in RUNTIME_FIELD_MAP.items():
        if field not in values and env_name not in values:
            continue
        raw_value = values.get(field, values.get(env_name))
        value = str(raw_value or '').strip()
        if not value:
            continue
        updates[env_name] = value
    if not updates:
        raise AsterConfigError('未提供可用的 Aster runtime 配置')
    if updates.get('PSTX_ASTER_MODE') and updates['PSTX_ASTER_MODE'] not in {'mock', 'live', 'off'}:
        raise AsterConfigError('PSTX_ASTER_MODE 仅支持 mock、live、off')
    if updates.get('PSTX_ASTER_BACKEND') and updates['PSTX_ASTER_BACKEND'] not in {'chat-flow', 'room'}:
        raise AsterConfigError('PSTX_ASTER_BACKEND 仅支持 chat-flow 或 room')
    with _RUNTIME_LOCK:
        _RUNTIME_OVERRIDES.update(updates)
    return build_aster_status()


def clear_aster_runtime_config() -> dict:
    with _RUNTIME_LOCK:
        _RUNTIME_OVERRIDES.clear()
    return build_aster_status()


def build_aster_status(*, environ: Optional[dict] = None) -> dict:
    env = _merged_env(environ)
    with _RUNTIME_LOCK:
        runtime_names = set(_RUNTIME_OVERRIDES)
    config = AsterConfig.from_env(env)
    backend = config.backend if config.backend in {'chat-flow', 'room'} else config.backend
    required_names = set()
    if config.mode == 'live':
        required_names.update({'ASTER_EMP_NO'})
        if config.backend == 'chat-flow':
            required_names.add('ASTER_API_KEY')
        elif config.backend == 'room':
            required_names.update({'ASTER_APP_ID', 'ASTER_APP_SECRET'})

    items = [
        _status_item(
            env,
            'PSTX_ASTER_MODE',
            '运行模式',
            required=False,
            source='runtime' if 'PSTX_ASTER_MODE' in runtime_names else 'environment',
            safe_value=config.mode or 'mock',
        ),
        _status_item(
            env,
            'PSTX_ASTER_BACKEND',
            'Aster 后端',
            required=False,
            source='runtime' if 'PSTX_ASTER_BACKEND' in runtime_names else 'environment',
            safe_value=backend or 'chat-flow',
        ),
        _status_item(
            {**env, 'ASTER_FIXED_BASE_URL': ASTER_FIXED_BASE_URL},
            'ASTER_FIXED_BASE_URL',
            'Aster 服务地址',
            required=False,
            source='fixed',
            safe_value=ASTER_FIXED_BASE_URL,
        ),
        _status_item(
            env,
            'ASTER_EMP_NO',
            '员工号',
            required='ASTER_EMP_NO' in required_names,
            source='runtime' if 'ASTER_EMP_NO' in runtime_names else 'environment',
            safe_value=config.emp_no,
        ),
        _status_item(
            env,
            'ASTER_API_KEY',
            'ChatFlow API Key',
            required='ASTER_API_KEY' in required_names,
            source='runtime' if 'ASTER_API_KEY' in runtime_names else 'environment',
            secret=True,
        ),
        _status_item(
            env,
            'ASTER_ORIGIN',
            'Room Validate Origin',
            required=False,
            source='runtime' if 'ASTER_ORIGIN' in runtime_names else 'environment',
            safe_value=_safe_origin(config.origin),
        ),
        _status_item(
            env,
            'ASTER_APP_ID',
            'Room App ID',
            required='ASTER_APP_ID' in required_names,
            source='runtime' if 'ASTER_APP_ID' in runtime_names else 'environment',
            safe_value=config.app_id,
        ),
        _status_item(
            env,
            'ASTER_APP_SECRET',
            'Room App Secret',
            required='ASTER_APP_SECRET' in required_names,
            source='runtime' if 'ASTER_APP_SECRET' in runtime_names else 'environment',
            secret=True,
        ),
    ]
    missing = [item['name'] for item in items if item['required'] and not item['configured']]
    status = 'mock'
    live_ready = False
    message = '当前使用本地 mock 摘要，不访问 Aster。'
    if config.mode == 'live':
        try:
            config.validate_live()
        except AsterConfigError as exc:
            status = 'missing'
            message = str(exc)
        else:
            status = 'ready'
            live_ready = True
            message = 'Aster live 配置已满足后端调用要求。'
    elif config.mode in {'off', 'disabled'}:
        status = 'off'
        message = 'Aster 摘要当前已禁用。'
    elif config.mode not in {'', 'mock'}:
        status = 'invalid'
        message = 'PSTX_ASTER_MODE 仅支持 mock、live、off。'

    return {
        'ok': True,
        'mode': config.mode or 'mock',
        'backend': backend or 'chat-flow',
        'status': status,
        'live_ready': live_ready,
        'message': message,
        'missing': missing,
        'runtime_override_active': bool(runtime_names),
        'runtime_override_keys': sorted(runtime_names),
        'log_file': get_aster_log_path(config.log_file),
        'items': items,
        'safeguards': [
            '前端可临时提交 Aster 凭据到后端内存，但状态接口不返回 secret/token/apiKey 原文或片段。',
            'Runtime 覆盖项仅保存在当前 Python 进程内存，重启服务后消失。',
            '当前工具只连接 Aster wrapper，不直连原生 Dify。',
        ],
    }


def build_aster_summary(report: dict, bundle: dict, *, environ: Optional[dict] = None) -> dict:
    env = _merged_env(environ)
    config = AsterConfig.from_env(env)
    if config.mode in {'', 'mock'}:
        return build_aster_mock_summary(report, bundle)
    if config.mode == 'live':
        return build_aster_live_summary(report, bundle, environ=env)
    if config.mode in {'off', 'disabled'}:
        raise AsterConfigError('Aster 摘要已通过 PSTX_ASTER_MODE 禁用')
    raise AsterConfigError('PSTX_ASTER_MODE 仅支持 mock、live、off')


def ask_aster_model(prompt: str,
                    *,
                    inputs: Optional[dict] = None,
                    environ: Optional[dict] = None) -> dict:
    env = _merged_env(environ)
    config = AsterConfig.from_env(env)
    if config.mode in {'', 'mock'}:
        return {
            'ok': True,
            'mode': 'mock',
            'provider': 'local-aster-mock',
            'answer': (
                '{"summary":"当前为本地 mock 模型接口，未访问真实 Aster。",'
                '"priorities":[],"review_checklist":[],"manual_review":[]}'
            ),
            'metadata': {
                'backend': config.backend,
                'prompt_chars': len(str(prompt or '')),
                'inputs_keys': sorted(str(key) for key in (inputs or {}).keys()),
            },
        }
    if config.mode == 'live':
        return ask_aster_live_model(prompt, inputs=inputs or {}, environ=env)
    if config.mode in {'off', 'disabled'}:
        raise AsterConfigError('Aster 模型接口已通过 PSTX_ASTER_MODE 禁用')
    raise AsterConfigError('PSTX_ASTER_MODE 仅支持 mock、live、off')


def aster_error_payload(exc: Exception) -> Tuple[dict, int]:
    if isinstance(exc, AsterConfigError):
        status = 400
        error_type = 'config'
    elif isinstance(exc, (AsterHttpError, AsterResponseError)):
        status = 502
        error_type = 'upstream'
    elif isinstance(exc, AsterError):
        status = 500
        error_type = 'aster'
    else:
        status = 500
        error_type = 'internal'
    diagnostics = getattr(exc, 'diagnostics', {}) or {}
    hints = []
    error_text = str(exc)
    if '401' in error_text or 'unauthorized' in error_text.lower() or 'access token is invalid' in error_text.lower():
        hints.extend([
            '确认 PSTX_ASTER_BACKEND 是否与 Aster 应用类型一致：ChatFlow 用 chat-flow，普通智能体/Room 用 room。',
            '如果 backend=chat-flow，确认 ASTER_API_KEY 是该 ChatFlow/AgentFlow 的 API Key，不是 accessToken。',
            '如果 backend=room，确认 ASTER_APP_ID、ASTER_APP_SECRET、ASTER_EMP_NO 能正常换取 accessToken。',
            'Aster 服务地址已固定为 https://aigc.huaqin.com；如果仍 401，请优先排查 API Key / App Secret / 员工号与应用发布状态。',
        ])
    payload = {
        'ok': False,
        'mode': 'live',
        'provider': 'aster',
        'error_type': error_type,
        'error': str(exc),
        'diagnostics': diagnostics,
        'diagnostic_hints': hints,
        'log_file': diagnostics.get('log_file') or get_aster_log_path(),
        'safeguards': [
            '状态接口不会回显敏感凭据原文。',
            '可将 PSTX_ASTER_MODE 切回 mock 以继续使用本地摘要。',
        ],
    }
    return payload, status
