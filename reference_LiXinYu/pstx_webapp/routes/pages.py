# -*- coding: utf-8 -*-
"""Page route registration for the PSTX Web UI."""

from __future__ import annotations

from pstx_webapp.debug_report import build_debug_report_payload


def register_page_routes(
    app,
    *,
    render_template,
    request,
    abort,
    render_home_page,
    render_named_page,
    render_report_page,
    default_host: str,
    default_port: int,
    get_run,
) -> None:
    """Register HTML page endpoints.

    The handlers here intentionally stay presentation-only: they render
    templates or retrieve an existing report payload for the report page.
    API behavior remains in dedicated route modules.
    """

    @app.get('/')
    def home():
        return render_home_page(
            render_template,
            request_host=request.host or '',
            default_host=default_host,
            default_port=default_port,
        )

    @app.get('/feishu-sync')
    def feishu_sync_page():
        return render_named_page(render_template, 'feishu_sync')

    @app.get('/feishu-db')
    def feishu_db_page():
        return render_named_page(render_template, 'feishu_db')

    @app.get('/ai-settings')
    def ai_settings_page():
        return render_named_page(render_template, 'ai_settings')

    @app.get('/guide')
    def guide_page():
        return render_named_page(render_template, 'guide')

    @app.get('/dfmea')
    def dfmea_page():
        run_id = str(request.args.get('run_id') or '').strip()
        if not run_id:
            abort(400, description='请通过 run_id 打开 DFMEA 工作台。')
        try:
            get_run(run_id)
        except KeyError:
            abort(404, description=f'未找到报告 run_id：{run_id}')
        return render_named_page(
            render_template,
            'dfmea',
            run_id=run_id,
            debug_ui=str(request.args.get('debug_ui') or '').strip() in {'1', 'true', 'on'},
            debug_fixture=False,
        )

    @app.get('/topology')
    def topology_page():
        run_id = str(request.args.get('run_id') or '').strip()
        if not run_id:
            abort(400, description='请通过 run_id 打开拓扑视图。')
        try:
            payload = get_run(run_id)
        except KeyError:
            abort(404, description=f'未找到报告 run_id：{run_id}')
        report = payload.get('report', {}) or {}
        return render_named_page(
            render_template,
            'topology',
            run_id=run_id,
            project_name=report.get('project_name', ''),
            debug_fixture=False,
        )

    @app.get('/compare')
    def compare_page():
        return render_named_page(render_template, 'compare')

    @app.get('/debug/dfmea')
    def debug_dfmea():
        return render_named_page(
            render_template,
            'dfmea',
            run_id='debug-dfmea',
            debug_ui=True,
            debug_fixture=True,
        )

    @app.get('/debug/topology')
    def debug_topology():
        return render_named_page(
            render_template,
            'topology',
            run_id='debug-topology',
            project_name='Debug Topology Fixture',
            debug_fixture=True,
        )

    @app.get('/debug/report')
    def debug_report_page():
        report = build_debug_report_payload()
        return render_report_page(
            render_template,
            run_id='debug-report',
            report=report,
            debug_ui=True,
            debug_fixture=True,
        )

    @app.get('/agent-eval')
    def agent_eval_page():
        return render_named_page(render_template, 'agent_eval')

    @app.get('/agent-lab')
    def agent_lab_page():
        return render_named_page(render_template, 'agent_lab')

    @app.get('/report/<run_id>')
    def report_page(run_id: str):
        try:
            payload = get_run(run_id)
        except KeyError:
            abort(404, description=f'未找到报告 run_id：{run_id}')
        report = payload['report']
        return render_report_page(render_template, run_id=run_id, report=report)
