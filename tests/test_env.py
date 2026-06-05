import os
import tempfile
from pathlib import Path


def configure_test_environment() -> None:
    base = os.environ.get("BOM_TOOLS_TEST_DATA_DIR")
    if not base:
        base = tempfile.mkdtemp(prefix="bom_tools_tests_")
        os.environ["BOM_TOOLS_TEST_DATA_DIR"] = base
    data_dir = Path(base)
    os.environ.setdefault("BOM_TOOLS_BUG_REPORT_DIR", str(data_dir / "bug_reports"))
    os.environ.setdefault("BOM_TOOLS_FEATURE_REQUEST_DIR", str(data_dir / "feature_requests"))
    os.environ.setdefault("BOM_TOOLS_AUTH_DATA_DIR", str(data_dir / "auth_data"))
