# -*- coding: utf-8 -*-

import unittest

from scripts.check_version_bumps import main as check_version_bumps_main


class VersionBumpIntegrationTests(unittest.TestCase):
    def test_current_worktree_has_required_version_bumps(self):
        self.assertEqual(check_version_bumps_main([]), 0)


if __name__ == "__main__":
    unittest.main()
