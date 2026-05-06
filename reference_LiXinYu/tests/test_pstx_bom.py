# -*- coding: utf-8 -*-
"""Tests for PSTX BOM grouping rules."""

import unittest

from pstx_rules.bom import build_bom, build_total_bom
from tests.pstx_test_fixtures import make_res


class BomTests(unittest.TestCase):

    def test_build_bom_does_not_merge_distinct_items_without_part_number(self):
        components = {
            'R1': {
                'refdes': 'R1', 'hq_code': '', 'part_name': 'RES_0402', 'value': '10k',
                'package': '0402', 'voltage': '', 'power': '', 'tolerance': '1%',
                'material': 'thick', 'comp_type': 'RES', 'page': 'PAGE1', 'room': '',
                'bom_option': '',
            },
            'R2': {
                'refdes': 'R2', 'hq_code': '', 'part_name': 'RES_0402', 'value': '4.7k',
                'package': '0402', 'voltage': '', 'power': '', 'tolerance': '1%',
                'material': 'thick', 'comp_type': 'RES', 'page': 'PAGE1', 'room': '',
                'bom_option': '',
            },
        }
        _, _, merged, _ = build_bom(components)
        self.assertEqual(2, len(merged))
        self.assertEqual({'10k', '4.7k'}, {row['值'] for row in merged})

    def test_build_bom_treats_dnp_as_depop(self):
        components = {
            'C1': {
                'refdes': 'C1', 'hq_code': '', 'part_name': 'CAP_0402', 'value': '1uF',
                'package': '0402', 'voltage': '6.3V', 'power': '', 'tolerance': '',
                'material': '', 'comp_type': 'CAP', 'page': 'PAGE1', 'room': '',
                'bom_option': 'DNP',
            },
        }
        detail_normal, detail_depop, _, merged_depop = build_bom(components)
        self.assertEqual([], detail_normal)
        self.assertEqual(1, len(detail_depop))
        self.assertEqual(1, len(merged_depop))

    def test_build_bom_prefers_real_page_display(self):
        components = {
            'R1': make_res('R1', 'P3V3', 'GPIO1', page='PAGE242', page_real='PAGE518'),
        }
        detail_normal, _, _, _ = build_bom(components)
        self.assertEqual('PAGE518', detail_normal[0]['页面'])

    def test_build_total_bom_merges_mounted_and_depop_counts(self):
        components = {
            'R1': make_res('R1', 'P3V3', 'GPIO1', value='10k', bom_option=''),
            'R2': make_res('R2', 'P3V3', 'GPIO2', value='10k', bom_option='DEPOP'),
        }
        detail_normal, detail_depop, _, _ = build_bom(components)
        total_detail, total_merged = build_total_bom(detail_normal, detail_depop)
        self.assertEqual({'贴装', 'DEPOP'}, {row['BOM状态'] for row in total_detail})
        self.assertEqual(1, len(total_merged))
        row = total_merged[0]
        self.assertEqual(2, row['数量'])
        self.assertEqual(1, row['贴装数量'])
        self.assertEqual(1, row['DEPOP数量'])
        self.assertEqual('贴装 / DEPOP', row['BOM状态'])


if __name__ == '__main__':
    unittest.main()
