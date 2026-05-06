"""Route modules for the Flask Web UI.

New route groups should register themselves from this package instead of adding
more handlers directly to ``pstx_web.py``.
"""

ROUTE_GROUPS = (
    "diagnostics",
    "pages",
    "system",
    "agent_lab",
    "projects",
    "reports",
    "dfmea",
    "harness",
    "compare",
    "feishu",
)
