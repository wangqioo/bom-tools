import unittest

from pstx_aster_mock import build_aster_mock_summary


class AsterMockSummaryTests(unittest.TestCase):
    def test_build_aster_mock_summary_prioritizes_report_counts(self):
        report = {
            'project_name': 'demo',
            'ratio_limit': 70,
            'metrics': [
                {'label': 'DRC 总数', 'value': 2},
                {'label': '降额不合格', 'value': 1},
                {'label': '电阻候选', 'value': 3},
                {'label': '电阻无法判断', 'value': 1},
                {'label': '规范候选', 'value': 2},
            ],
            'sections': [
                {'id': 'drc', 'total_rows': 2, 'tables': [{'id': 'missing_value', 'count': 1}]},
                {'id': 'resistor', 'total_rows': 4, 'tables': [{'id': 'divider_risks', 'count': 2}]},
                {'id': 'derating', 'total_rows': 5, 'tables': [{'id': 'derating', 'count': 5}]},
                {'id': 'csa', 'total_rows': 2, 'tables': [{'id': 'csa_dot_cross_rows', 'count': 1}, {'id': 'csa_circle_rows', 'count': 1}]},
            ],
        }

        payload = build_aster_mock_summary(report, {})

        self.assertTrue(payload['ok'])
        self.assertEqual('mock', payload['mode'])
        self.assertEqual('local-aster-mock', payload['provider'])
        self.assertIn('没有访问真实 Aster', payload['summary'])
        self.assertEqual('先看设计检查', payload['priorities'][0]['title'])
        self.assertTrue(any(item['target'] == 'csa' for item in payload['priorities']))
        self.assertEqual('电容降额', payload['section_focus'][0]['section'])
        self.assertTrue(any(item['item'] == '芯片 Pin 与电阻网络' for item in payload['review_checklist']))
        self.assertTrue(any(item['topic'] == '电平/电压推断' for item in payload['manual_review']))
        self.assertTrue(any('不访问真实 Aster' in item for item in payload['safeguards']))


if __name__ == '__main__':
    unittest.main()
