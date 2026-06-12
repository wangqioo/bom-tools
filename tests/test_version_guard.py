# -*- coding: utf-8 -*-

import unittest
from pathlib import Path

from scripts.check_deploy_bundle_sync import analyze_deploy_bundle_sync
from scripts.check_version_bumps import analyze_version_bumps, parse_versions


ROOT = Path(__file__).resolve().parents[1]


class VersionGuardTests(unittest.TestCase):
    def test_free_bom_change_requires_free_bom_version_bump(self):
        paths = ["web_app2/bom_compare/generic_free.py"]
        old = {"tool:free-bom-compare": "1.1.2"}
        new = {"tool:free-bom-compare": "1.1.2"}

        errors = analyze_version_bumps(paths, old, new)

        self.assertIn("tool:free-bom-compare must be bumped", errors)

    def test_free_bom_change_passes_when_version_bumped(self):
        paths = ["web_app2/bom_compare/generic_free.py"]
        old = {"tool:free-bom-compare": "1.1.2"}
        new = {"tool:free-bom-compare": "1.1.3"}

        errors = analyze_version_bumps(paths, old, new)

        self.assertEqual(errors, [])

    def test_bom_compare_shell_template_requires_bom_compare_version_bump(self):
        paths = ["web_app2/templates/partials/tools/bom-compare.html"]
        old = {"tool:bom-compare": "0.2.0", "tool:free-bom-compare": "1.1.2"}
        new = {"tool:bom-compare": "0.2.0", "tool:free-bom-compare": "1.1.3"}

        errors = analyze_version_bumps(paths, old, new)

        self.assertIn("tool:bom-compare must be bumped", errors)
        self.assertNotIn("tool:free-bom-compare must be bumped", errors)

    def test_customer_hq_compare_change_requires_customer_hq_version_bump(self):
        paths = ["web_app2/bom_compare/customer_hq_export.py"]
        old = {"tool:customer-hq-compare": "1.0.0"}
        new = {"tool:customer-hq-compare": "1.0.0"}

        errors = analyze_version_bumps(paths, old, new)

        self.assertIn("tool:customer-hq-compare must be bumped", errors)

    def test_bom_compare_backend_change_requires_collection_version_bump(self):
        paths = ["web_app2/bom_compare/__init__.py"]
        old = {"tool:bom-compare": "0.2.0"}
        new = {"tool:bom-compare": "0.2.0"}

        errors = analyze_version_bumps(paths, old, new)

        self.assertIn("tool:bom-compare must be bumped", errors)

    def test_shared_frontend_change_requires_any_version_bump(self):
        paths = ["web_app2/static/js/app.js"]
        old = {"platform": "2.1.0"}
        new = {"platform": "2.1.0"}

        errors = analyze_version_bumps(paths, old, new)

        self.assertIn(
            "shared frontend/style changes require at least one platform or tool version bump",
            errors,
        )

    def test_parse_versions_from_shared_module(self):
        versions = parse_versions(
            'PLATFORM_VERSION = "2.1.0"\n'
            'TOOL_VERSIONS = {"free-bom-compare": "1.1.2"}\n'
        )

        self.assertEqual(versions["platform"], "2.1.0")
        self.assertEqual(versions["tool:free-bom-compare"], "1.1.2")

    def test_deploy_bundle_has_main_tool_entries(self):
        errors = analyze_deploy_bundle_sync(ROOT)

        self.assertEqual(errors, [])


if __name__ == "__main__":
    unittest.main()
