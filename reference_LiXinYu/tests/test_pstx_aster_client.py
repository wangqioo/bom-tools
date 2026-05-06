import json
import tempfile
import threading
import unittest
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

from pstx_integrations.aster import client as pstx_aster_client
from pstx_integrations.aster import client as integration_aster_client
from pstx_integrations.aster import mock as integration_aster_mock
from pstx_integrations.aster import service as integration_aster_service
from pstx_integrations.aster.client import (
    AsterConfig,
    AsterConfigError,
    AsterHttpError,
    aster_decrypt,
    aster_encrypt,
    build_aster_prompt,
    build_aster_live_summary,
    build_report_brief,
)


def sample_report():
    return {
        'project_name': 'demo_board',
        'ratio_limit': 70,
        'metrics': [
            {'label': 'DRC 总数', 'value': 2},
            {'label': '降额不合格', 'value': 1},
            {'label': '电阻候选', 'value': 3},
            {'label': '规范候选', 'value': 1},
        ],
        'summary_lines': ['DRC 确定结论：1', '电阻候选判断：3'],
        'warnings': ['测试 warning'],
        'sections': [
            {
                'id': 'drc',
                'title': '设计检查',
                'total_rows': 2,
                'tables': [
                    {
                        'id': 'missing_value',
                        'title': '缺少 VALUE',
                        'count': 1,
                        'columns': ['位号', '页面', '文件'],
                        'kind_counts': {'确定结论': 1},
                        'rows': [{'位号': 'R1', '页面': 'PAGE12', '文件': '/Users/rounder/secret/page12.csa'}],
                    }
                ],
            },
            {
                'id': 'resistor',
                'title': '电阻检查',
                'total_rows': 3,
                'tables': [
                    {
                        'id': 'chip_pin_rows',
                        'title': '芯片 Pin 电阻状态',
                        'count': 3,
                        'columns': ['芯片位号', '引脚', '状态'],
                        'kind_counts': {'候选判断': 3},
                        'rows': [{'芯片位号': 'U1', '引脚': 'GPIO1', '状态': '需人工确认'}],
                    }
                ],
            },
        ],
    }


class AsterMockHandler(BaseHTTPRequestHandler):
    calls = []
    answer_mode = 'json'
    auth_count = 0
    validate_count = 0
    validate_invalid_once = False
    flow_count = 0

    def log_message(self, fmt, *args):
        return

    def do_POST(self):
        parsed = urlparse(self.path)
        length = int(self.headers.get('Content-Length') or '0')
        raw_body = self.rfile.read(length).decode('utf-8')
        try:
            body = json.loads(raw_body or '{}')
        except json.JSONDecodeError:
            body = raw_body

        if parsed.path == '/auth/api/v1/generateOrProlongToken':
            self.handle_generate_token(body)
            return
        if parsed.path == '/auth/js-sdk/validateAccessToken':
            self.handle_validate_token(raw_body)
            return
        if parsed.path == '/aster/room/create':
            self.handle_create_room(parsed)
            return
        if parsed.path == '/aster/room/chat/stream/send':
            self.handle_room_chat(body)
            return
        self.__class__.calls.append({
            'path': parsed.path,
            'query': parse_qs(parsed.query),
            'body': body,
        })
        if parsed.path != '/aster/flow-api/run/chat-flow':
            self.send_response(404)
            self.end_headers()
            return
        self.__class__.flow_count += 1
        if self.__class__.answer_mode == 'chunked_once' and self.__class__.flow_count == 1:
            response = {
                'code': 500,
                'data': None,
                'msg': '请求失败,code:invalid_param,message:Run failed: [openai_api_compatible] Error: PluginInvokeError: {"error_type":"ChunkedEncodingError","message":"Response ended prematurely"},status:400',
                'failed': True,
                'success': False,
            }
            data = json.dumps(response, ensure_ascii=False).encode('utf-8')
            self.send_response(401)
            self.send_header('Content-Type', 'application/json;charset=utf-8')
            self.send_header('Content-Length', str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return
        if self.__class__.answer_mode == 'http401':
            api_key = parse_qs(parsed.query).get('apiKey', [''])[0]
            response = {
                'code': 500,
                'data': None,
                'msg': '请求失败,code:unauthorized,message:Access token is invalid,status:401',
                'echo': f'apiKey={api_key}',
                'failed': True,
                'success': False,
            }
            data = json.dumps(response, ensure_ascii=False).encode('utf-8')
            self.send_response(401)
            self.send_header('Content-Type', 'application/json;charset=utf-8')
            self.send_header('Content-Length', str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return
        if self.__class__.answer_mode == 'text':
            answer = '这是 Aster 直接返回的纯文本摘要。'
        else:
            answer = json.dumps({
                'summary': 'Aster 认为需要先处理 DRC 和电阻候选。',
                'priorities': [
                    {
                        'title': '先处理 DRC',
                        'body': '缺少 VALUE 需要优先确认。',
                        'target': 'drc',
                        'severity': 'high',
                    }
                ],
                'section_focus': [
                    {
                        'section': '设计检查',
                        'target': 'drc',
                        'rows': 2,
                        'reason': '存在确定结论。',
                    }
                ],
                'review_checklist': [
                    {
                        'item': '属性与命名 DRC',
                        'status': 'needs_review',
                        'evidence': '缺少 VALUE 需要复核。',
                        'target': 'drc',
                        'severity': 'high',
                    }
                ],
                'manual_review': [
                    {
                        'topic': '电平/电压推断',
                        'reason': 'token 推断不得下确定结论。',
                        'target': 'derating',
                    }
                ],
            }, ensure_ascii=False)
        response = {
            'code': 200,
            'data': {
                'answer': answer,
                'conversation_id': 'conv-live',
                'message_id': 'msg-live',
                'task_id': 'task-live',
            },
        }
        data = json.dumps(response, ensure_ascii=False).encode('utf-8')
        self.send_response(200)
        self.send_header('Content-Type', 'application/json;charset=utf-8')
        self.send_header('Content-Length', str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def write_json(self, payload: dict, status: int = 200):
        data = json.dumps(payload, ensure_ascii=False).encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type', 'application/json;charset=utf-8')
        self.send_header('Content-Length', str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def write_text(self, text: str, status: int = 200):
        data = str(text).encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type', 'text/plain;charset=utf-8')
        self.send_header('Content-Length', str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def handle_generate_token(self, body):
        self.__class__.auth_count += 1
        token = f'room-token-{self.__class__.auth_count}'
        decrypted = json.loads(aster_decrypt(body['appId'], body['ciphertext']))
        self.__class__.calls.append({
            'path': '/auth/api/v1/generateOrProlongToken',
            'body': {**body, 'ciphertext_plain': decrypted},
        })
        response_data = aster_encrypt(body['appId'], json.dumps({
            'accessToken': token,
            'validityPeriodSeconds': 7200,
            'empNo': body['empNo'],
        }, ensure_ascii=False))
        self.write_json({
            'code': 200,
            'data': response_data,
            'message': 'OK',
            'success': True,
            'failed': False,
        })

    def handle_validate_token(self, raw_body: str):
        self.__class__.validate_count += 1
        origin = self.headers.get('aigc-origin') or self.headers.get('Origin') or ''
        decrypted = json.loads(aster_decrypt(origin, raw_body))
        token = str(decrypted.get('accessTokenRaw') or '').replace('Bearer ', '')
        is_valid = not (self.__class__.validate_invalid_once and self.__class__.validate_count == 1)
        self.__class__.calls.append({
            'path': '/auth/js-sdk/validateAccessToken',
            'origin': origin,
            'authorization_header': self.headers.get('Authorization'),
            'body_plain': decrypted,
        })
        payload = {
            'code': 200,
            'data': {
                'isValid': is_valid,
                'accessToken': token,
                'statusCode': 1 if is_valid else -1,
                'appId': decrypted.get('appId'),
                'validityPeriodSeconds': 7200 if is_valid else 0,
                'empNo': '100019100',
            },
            'message': 'OK',
        }
        self.write_text(aster_encrypt(origin, json.dumps(payload, ensure_ascii=False)))

    def handle_create_room(self, parsed):
        self.__class__.calls.append({
            'path': '/aster/room/create',
            'query': parse_qs(parsed.query),
            'authorization_header': self.headers.get('Authorization'),
        })
        self.write_json({
            'code': 200,
            'data': {'roomId': 'room-live', 'name': 'PSTX审查'},
            'message': 'OK',
            'success': True,
            'failed': False,
        })

    def handle_room_chat(self, body):
        self.__class__.calls.append({
            'path': '/aster/room/chat/stream/send',
            'body': body,
            'authorization_header': self.headers.get('Authorization'),
        })
        self.write_text(json.dumps({
            'code': 200,
            'data': {'content': 'Room 模式摘要返回。', 'messageType': 'answer'},
            'success': True,
            'failed': False,
        }, ensure_ascii=False))


class AsterClientTests(unittest.TestCase):
    def test_integration_aster_entrypoints_export_public_api(self):
        self.assertIs(integration_aster_client.AsterConfig, pstx_aster_client.AsterConfig)
        self.assertIs(integration_aster_client.build_aster_live_summary, pstx_aster_client.build_aster_live_summary)
        self.assertIs(integration_aster_mock.build_aster_mock_summary, pstx_aster_client.build_aster_mock_summary)
        self.assertFalse(Path("pstx_aster_client.py").exists())
        self.assertFalse(Path("pstx_aster_service.py").exists())
        self.assertFalse(Path("pstx_aster_mock.py").exists())
        self.assertTrue(callable(integration_aster_service.ask_aster_model))

    def start_server(self):
        AsterMockHandler.calls = []
        AsterMockHandler.answer_mode = 'json'
        AsterMockHandler.auth_count = 0
        AsterMockHandler.validate_count = 0
        AsterMockHandler.validate_invalid_once = False
        AsterMockHandler.flow_count = 0
        server = ThreadingHTTPServer(('127.0.0.1', 0), AsterMockHandler)
        thread = threading.Thread(target=server.serve_forever, daemon=True)
        thread.start()
        self.addCleanup(server.shutdown)
        self.addCleanup(server.server_close)
        return f'http://127.0.0.1:{server.server_address[1]}'

    def use_mock_aster_base_url(self, base_url: str):
        old_base_url = pstx_aster_client.ASTER_FIXED_BASE_URL
        pstx_aster_client.ASTER_FIXED_BASE_URL = base_url
        self.addCleanup(lambda: setattr(pstx_aster_client, 'ASTER_FIXED_BASE_URL', old_base_url))

    def test_build_aster_live_summary_calls_chat_flow_and_normalizes_json(self):
        base_url = self.start_server()
        self.use_mock_aster_base_url(base_url)
        payload = build_aster_live_summary(sample_report(), {}, environ={
            'ASTER_API_KEY': 'flow-key',
            'ASTER_EMP_NO': '100019100',
            'PSTX_ASTER_BACKEND': 'chat-flow',
        })

        self.assertTrue(payload['ok'])
        self.assertEqual('live', payload['mode'])
        self.assertEqual('aster-chat-flow', payload['provider'])
        self.assertEqual('Aster 认为需要先处理 DRC 和电阻候选。', payload['summary'])
        self.assertEqual('先处理 DRC', payload['priorities'][0]['title'])
        self.assertEqual('属性与命名 DRC', payload['review_checklist'][0]['item'])
        self.assertEqual('电平/电压推断', payload['manual_review'][0]['topic'])
        self.assertEqual('conv-live', payload['metadata']['conversation_id'])
        call = AsterMockHandler.calls[0]
        self.assertEqual('/aster/flow-api/run/chat-flow', call['path'])
        self.assertEqual(['flow-key'], call['query']['apiKey'])
        self.assertEqual(['100019100'], call['query']['empNo'])
        self.assertIn('只输出一个 JSON 对象', call['body']['query'])
        self.assertIn('demo_board', call['body']['query'])

    def test_build_aster_live_summary_uses_text_fallback_when_answer_is_not_json(self):
        base_url = self.start_server()
        self.use_mock_aster_base_url(base_url)
        AsterMockHandler.answer_mode = 'text'
        payload = build_aster_live_summary(sample_report(), {}, environ={
            'ASTER_API_KEY': 'flow-key',
            'ASTER_EMP_NO': '100019100',
        })

        self.assertEqual('这是 Aster 直接返回的纯文本摘要。', payload['summary'])
        self.assertEqual('text', payload['metadata']['answer_format'])
        self.assertTrue(payload['priorities'])

    def test_live_config_requires_chat_flow_credentials(self):
        config = AsterConfig.from_env({
            'PSTX_ASTER_MODE': 'live',
            'ASTER_EMP_NO': '100019100',
        })
        with self.assertRaises(AsterConfigError):
            config.validate_live()

    def test_config_normalizes_aster_origin_url_to_host(self):
        config = AsterConfig.from_env({
            'ASTER_ORIGIN': 'https://test-aigc-api.huaqin.com/path?x=1',
        })

        self.assertEqual('test-aigc-api.huaqin.com', config.origin)

    def test_default_model_timeout_is_ten_minutes(self):
        config = AsterConfig.from_env({})

        self.assertEqual(600.0, config.timeout_seconds)

    def test_model_timeout_can_still_be_overridden_by_env(self):
        config = AsterConfig.from_env({'PSTX_ASTER_TIMEOUT_SECONDS': '12.5'})

        self.assertEqual(12.5, config.timeout_seconds)

    def test_report_brief_redacts_paths_and_limits_rows(self):
        config = AsterConfig.from_env({
            'PSTX_ASTER_MAX_ROWS_PER_TABLE': '1',
            'PSTX_ASTER_REDACT_PATHS': '1',
        })
        brief = build_report_brief(sample_report(), config)
        row = brief['sections'][0]['tables'][0]['sample_rows'][0]
        self.assertEqual('[path:page12.csa]', row['文件'])
        self.assertTrue(any(item['item'] == '属性与命名 DRC' for item in brief['review_scope']))
        self.assertTrue(any(item['table_id'] == 'missing_value' for item in brief['key_findings']))
        self.assertTrue(any(item['topic'] == '电平/电压推断' for item in brief['manual_review_boundaries']))
        self.assertEqual(1, len(next(item for item in brief['key_findings'] if item['table_id'] == 'missing_value')['sample_rows']))

    def test_aster_prompt_requests_expanded_review_schema(self):
        config = AsterConfig.from_env({})
        prompt = build_aster_prompt(build_report_brief(sample_report(), config))

        self.assertIn('review_checklist', prompt)
        self.assertIn('manual_review', prompt)
        self.assertIn('BOM/DEPOP、网络/页码映射、DRC、芯片 Pin/电阻、降额、CSA', prompt)

    def test_aster_log_sanitizer_redacts_password_fields(self):
        payload = pstx_aster_client.sanitize_for_aster_log({
            'password': 'super-secret-password',
            'body': 'password=plain-text-secret appSecret=another-secret',
        })

        self.assertTrue(payload['password']['redacted'])
        self.assertNotIn('super-secret-password', json.dumps(payload, ensure_ascii=False))
        self.assertNotIn('plain-text-secret', payload['body'])
        self.assertIn('password=<redacted>', payload['body'])

    def test_http_401_writes_sanitized_diagnostics_log(self):
        base_url = self.start_server()
        self.use_mock_aster_base_url(base_url)
        AsterMockHandler.answer_mode = 'http401'
        with tempfile.TemporaryDirectory() as temp_dir:
            log_file = f'{temp_dir}/aster_debug.log'
            with self.assertRaises(AsterHttpError) as raised:
                build_aster_live_summary(sample_report(), {}, environ={
                    'ASTER_API_KEY': 'flow-secret-key',
                    'ASTER_EMP_NO': '100019100',
                    'PSTX_ASTER_BACKEND': 'chat-flow',
                    'PSTX_ASTER_LOG_FILE': log_file,
                })

            diagnostics = raised.exception.diagnostics
            self.assertEqual(401, diagnostics['status'])
            self.assertEqual('chat_flow', diagnostics['operation'])
            self.assertIn('request_id', diagnostics)
            self.assertNotIn('flow-secret-key', str(raised.exception))
            with open(log_file, encoding='utf-8') as handle:
                text = handle.read()
            self.assertIn('request.start', text)
            self.assertIn('request.http_error', text)
            self.assertIn('Access token is invalid', text)
            self.assertNotIn('flow-secret-key', text)
            self.assertIn('apiKey=<redacted>', text)

    def test_chat_flow_retries_transient_chunked_error_and_logs_attempts(self):
        base_url = self.start_server()
        self.use_mock_aster_base_url(base_url)
        AsterMockHandler.answer_mode = 'chunked_once'
        with tempfile.TemporaryDirectory() as temp_dir:
            log_file = f'{temp_dir}/aster_debug.log'
            payload = build_aster_live_summary(sample_report(), {}, environ={
                'ASTER_API_KEY': 'flow-secret-key',
                'ASTER_EMP_NO': '100019100',
                'PSTX_ASTER_BACKEND': 'chat-flow',
                'PSTX_ASTER_LOG_FILE': log_file,
                'PSTX_ASTER_RETRY_COUNT': '1',
                'PSTX_ASTER_RETRY_BACKOFF_SECONDS': '0',
            })

            self.assertTrue(payload['ok'])
            self.assertEqual(2, AsterMockHandler.flow_count)
            with open(log_file, encoding='utf-8') as handle:
                text = handle.read()
            self.assertIn('request.retry', text)
            self.assertIn('"attempt": 1', text)
            self.assertIn('"attempt": 2', text)
            self.assertIn('"retryable": true', text)
            self.assertIn('"elapsed_ms"', text)
            self.assertNotIn('flow-secret-key', text)

    def test_room_backend_generates_and_validates_token_before_question(self):
        base_url = self.start_server()
        self.use_mock_aster_base_url(base_url)
        payload = build_aster_live_summary(sample_report(), {}, environ={
            'PSTX_ASTER_BACKEND': 'room',
            'ASTER_APP_ID': 'ag_demo',
            'ASTER_APP_SECRET': 'room-secret',
            'ASTER_EMP_NO': '100019100',
            'ASTER_ORIGIN': 'test-aigc-api.huaqin.com',
        })

        self.assertTrue(payload['ok'])
        self.assertEqual('aster-room', payload['provider'])
        self.assertEqual('Room 模式摘要返回。', payload['summary'])
        paths = [call['path'] for call in AsterMockHandler.calls]
        self.assertIn('/auth/api/v1/generateOrProlongToken', paths)
        self.assertGreaterEqual(paths.count('/auth/js-sdk/validateAccessToken'), 2)
        self.assertIn('/aster/room/create', paths)
        self.assertIn('/aster/room/chat/stream/send', paths)
        auth_call = next(call for call in AsterMockHandler.calls if call['path'] == '/auth/api/v1/generateOrProlongToken')
        self.assertEqual('', auth_call['body']['ciphertext_plain']['accessToken'])
        validate_call = next(call for call in AsterMockHandler.calls if call['path'] == '/auth/js-sdk/validateAccessToken')
        self.assertEqual('test-aigc-api.huaqin.com', validate_call['origin'])
        self.assertEqual('Bearer room-token-1', validate_call['body_plain']['accessTokenRaw'])
        self.assertNotEqual('Bearer room-token-1', validate_call['authorization_header'])

    def test_room_backend_force_renews_when_validate_reports_invalid(self):
        base_url = self.start_server()
        self.use_mock_aster_base_url(base_url)
        AsterMockHandler.validate_invalid_once = True

        payload = build_aster_live_summary(sample_report(), {}, environ={
            'PSTX_ASTER_BACKEND': 'room',
            'ASTER_APP_ID': 'ag_demo',
            'ASTER_APP_SECRET': 'room-secret',
            'ASTER_EMP_NO': '100019100',
            'ASTER_ORIGIN': 'test-aigc-api.huaqin.com',
        })

        self.assertTrue(payload['ok'])
        self.assertGreaterEqual(AsterMockHandler.auth_count, 2)
        validate_tokens = [
            call['body_plain']['accessTokenRaw']
            for call in AsterMockHandler.calls
            if call['path'] == '/auth/js-sdk/validateAccessToken'
        ]
        self.assertIn('Bearer room-token-1', validate_tokens)
        self.assertIn('Bearer room-token-2', validate_tokens)


if __name__ == '__main__':
    unittest.main()
