"""Small page rendering helpers for the Flask Web app.

Route registration lives in ``pstx_webapp.routes``. This module only owns
template names and small page context helpers.
"""

from __future__ import annotations

from typing import Any, Callable, Mapping


RenderTemplate = Callable[..., str]


PAGE_TEMPLATES: Mapping[str, str] = {
    "home": "index.html",
    "feishu_sync": "feishu_sync.html",
    "feishu_db": "feishu_db.html",
    "ai_settings": "ai_settings.html",
    "guide": "guide.html",
    "dfmea": "dfmea.html",
    "compare": "compare.html",
    "topology": "topology.html",
    "agent_eval": "agent_eval.html",
    "agent_lab": "agent_lab.html",
}


def build_home_context(request_host: str, default_host: str, default_port: int) -> dict[str, str]:
    host_text = request_host or f"{default_host}:{default_port}"
    listen_port = host_text.rsplit(":", 1)[-1] if ":" in host_text else str(default_port)
    return {
        "listen_host": default_host,
        "listen_port": listen_port,
    }


def render_named_page(render_template: RenderTemplate, page: str, **context: Any) -> str:
    template_name = PAGE_TEMPLATES[page]
    return render_template(template_name, **context)


def render_home_page(
    render_template: RenderTemplate,
    *,
    request_host: str,
    default_host: str,
    default_port: int,
) -> str:
    return render_named_page(
        render_template,
        "home",
        **build_home_context(request_host, default_host, default_port),
    )


def render_report_page(
    render_template: RenderTemplate,
    *,
    run_id: str,
    report: dict[str, Any],
    debug_ui: bool = False,
    debug_fixture: bool = False,
) -> str:
    return render_template(
        "report.html",
        run_id=run_id,
        report=report,
        debug_ui=debug_ui,
        debug_fixture=debug_fixture,
    )
