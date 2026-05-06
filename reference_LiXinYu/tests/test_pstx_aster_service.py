import unittest

from pstx_integrations.aster.client import AsterHttpError
from pstx_integrations.aster.service import (
    aster_error_payload,
    build_aster_status,
    build_aster_summary,
    clear_aster_runtime_config,
    set_aster_runtime_config,
)


class AsterServiceTests(unittest.TestCase):
    def tearDown(self):
        clear_aster_runtime_config()

    def test_default_mode_uses_mock_summary(self):
        payload = build_aster_summary({
            'project_name': 'demo',
            'metrics': [],
            'sections': [],
        }, {}, environ={})

        self.assertTrue(payload['ok'])
        self.assertEqual('mock', payload['mode'])
        self.assertEqual('local-aster-mock', payload['provider'])

    def test_live_mode_missing_config_returns_config_error_payload(self):
        with self.assertRaises(Exception) as raised:
            build_aster_summary({}, {}, environ={'PSTX_ASTER_MODE': 'live'})

        payload, status = aster_error_payload(raised.exception)
        self.assertEqual(400, status)
        self.assertFalse(payload['ok'])
        self.assertEqual('config', payload['error_type'])
        self.assertNotIn('ASTER_BASE_URL', payload['error'])
        self.assertIn('ASTER_EMP_NO', payload['error'])

    def test_aster_status_redacts_secret_values(self):
        status = build_aster_status(environ={
            'PSTX_ASTER_MODE': 'live',
            'PSTX_ASTER_BACKEND': 'chat-flow',
            'ASTER_EMP_NO': '100019100',
            'ASTER_API_KEY': 'super-secret-key',
        })

        self.assertEqual('ready', status['status'])
        self.assertTrue(status['live_ready'])
        status_text = str(status)
        self.assertNotIn('super-secret-key', status_text)
        self.assertNotIn('should-hide', status_text)
        item_map = {item['name']: item for item in status['items']}
        self.assertTrue(item_map['ASTER_API_KEY']['configured'])
        self.assertTrue(item_map['ASTER_API_KEY']['secret'])
        self.assertNotIn('value', item_map['ASTER_API_KEY'])
        self.assertNotIn('ASTER_BASE_URL', item_map)
        self.assertEqual('https://aigc.huaqin.com', item_map['ASTER_FIXED_BASE_URL']['value'])
        self.assertEqual('fixed', item_map['ASTER_FIXED_BASE_URL']['source'])

    def test_aster_status_lists_missing_live_credentials(self):
        status = build_aster_status(environ={'PSTX_ASTER_MODE': 'live'})

        self.assertEqual('missing', status['status'])
        self.assertFalse(status['live_ready'])
        self.assertNotIn('ASTER_BASE_URL', status['missing'])
        self.assertIn('ASTER_EMP_NO', status['missing'])
        self.assertIn('ASTER_API_KEY', status['missing'])

    def test_runtime_config_overrides_env_without_echoing_secret(self):
        status = set_aster_runtime_config({
            'mode': 'live',
            'backend': 'chat-flow',
            'emp_no': '100019100',
            'api_key': 'runtime-secret-key',
            'origin': 'runtime-origin.example.local',
        })

        self.assertEqual('ready', status['status'])
        self.assertTrue(status['runtime_override_active'])
        self.assertIn('ASTER_API_KEY', status['runtime_override_keys'])
        self.assertIn('ASTER_ORIGIN', status['runtime_override_keys'])
        status_text = str(status)
        self.assertNotIn('runtime-secret-key', status_text)
        item_map = {item['name']: item for item in status['items']}
        self.assertEqual('runtime', item_map['ASTER_API_KEY']['source'])
        self.assertNotIn('value', item_map['ASTER_API_KEY'])
        self.assertNotIn('ASTER_BASE_URL', item_map)
        self.assertEqual('https://aigc.huaqin.com', item_map['ASTER_FIXED_BASE_URL']['value'])
        self.assertEqual('runtime-origin.example.local', item_map['ASTER_ORIGIN']['value'])

        cleared = clear_aster_runtime_config()
        self.assertFalse(cleared['runtime_override_active'])

    def test_upstream_401_error_payload_includes_diagnostics_and_hints(self):
        exc = AsterHttpError(
            'Aster HTTP 401: {"msg":"Access token is invalid"}',
            diagnostics={
                'request_id': 'req-401',
                'operation': 'chat_flow',
                'backend': 'chat-flow',
                'status': 401,
                'log_file': '/tmp/aster-debug.log',
            },
        )

        payload, status = aster_error_payload(exc)

        self.assertEqual(502, status)
        self.assertFalse(payload['ok'])
        self.assertEqual('upstream', payload['error_type'])
        self.assertEqual('req-401', payload['diagnostics']['request_id'])
        self.assertEqual(401, payload['diagnostics']['status'])
        self.assertEqual('/tmp/aster-debug.log', payload['log_file'])
        self.assertTrue(any('ASTER_API_KEY' in hint for hint in payload['diagnostic_hints']))


if __name__ == '__main__':
    unittest.main()
