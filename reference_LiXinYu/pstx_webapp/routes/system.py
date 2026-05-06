# -*- coding: utf-8 -*-
"""System, model, harness metadata, datasheet, and eval API routes."""

from __future__ import annotations


def _as_bool(value) -> bool:
    return value if isinstance(value, bool) else str(value).strip().lower() in {'1', 'true', 'yes', 'on'}


def register_system_routes(
    app,
    *,
    request,
    jsonify,
    build_aster_status,
    set_aster_runtime_config,
    clear_aster_runtime_config,
    aster_error_payload,
    build_harness_status,
    build_default_harness_registry,
    list_harness_agent_profiles,
    build_datasheet_status,
    reindex_datasheets,
    build_agent_eval_status,
    run_agent_eval,
    AgentEvalError,
) -> None:
    """Register small system/status APIs that do not own report state."""

    @app.get('/api/aster/status')
    def aster_status():
        return jsonify(build_aster_status())

    @app.post('/api/aster/runtime-config')
    def aster_runtime_config_update():
        data = request.get_json(silent=True) or request.form.to_dict()
        try:
            return jsonify(set_aster_runtime_config(data))
        except Exception as exc:
            error_payload, status = aster_error_payload(exc)
            return jsonify(error_payload), status

    @app.delete('/api/aster/runtime-config')
    def aster_runtime_config_clear():
        return jsonify(clear_aster_runtime_config())

    @app.get('/api/harness/status')
    def harness_status():
        return jsonify(build_harness_status(model_status=build_aster_status()))

    @app.get('/api/harness/tools')
    def harness_tools():
        registry = build_default_harness_registry()
        return jsonify({
            'ok': True,
            'mode': 'local-agent-harness',
            'tools': registry.list_tools(),
            'tool_count': len(registry.list_tools()),
            'safeguards': [
                '所有工具均由本地 harness 白名单和 input_schema 校验后执行。',
                '项目文件读取仅限当前 project_root 的 packaged、sch_1、module_order(.dat)、page.map 范围。',
                '工具第一版只读，不写文件、不修改报告。',
            ],
        })

    @app.get('/api/harness/profiles')
    def harness_profiles():
        return jsonify({
            'ok': True,
            'mode': 'local-agent-harness',
            'profiles': list_harness_agent_profiles(),
            'default_profile': 'quick_scan',
            'safeguards': [
                'Profile 只限制本地 harness 可调用的只读工具集合。',
                'Aster 仍只作为模型 provider，不接收任何本地执行权限。',
            ],
        })

    @app.get('/api/datasheets/status')
    def datasheets_status():
        return jsonify(build_datasheet_status())

    @app.post('/api/datasheets/reindex')
    def datasheets_reindex():
        data = request.get_json(silent=True) or request.form.to_dict()
        raw_force = data.get('force', False) if isinstance(data, dict) else False
        force = _as_bool(raw_force)
        try:
            max_files = int((data.get('max_files') if isinstance(data, dict) else None) or 5000)
        except (TypeError, ValueError):
            return jsonify({'ok': False, 'error': 'max_files 必须是数字。'}), 400
        if max_files < 1 or max_files > 50000:
            return jsonify({'ok': False, 'error': 'max_files 必须在 1-50000 之间。'}), 400
        return jsonify(reindex_datasheets(force=force, max_files=max_files))

    @app.get('/api/agent-eval/status')
    def agent_eval_status():
        return jsonify(build_agent_eval_status())

    @app.post('/api/agent-eval/run')
    def agent_eval_run():
        data = request.get_json(silent=True) or request.form.to_dict()
        raw_case_ids = data.get('case_ids') if isinstance(data, dict) else []
        if isinstance(raw_case_ids, str):
            case_ids = [item.strip() for item in raw_case_ids.split(',') if item.strip()]
        elif isinstance(raw_case_ids, list):
            case_ids = [str(item).strip() for item in raw_case_ids if str(item).strip()]
        else:
            case_ids = []
        try:
            return jsonify(run_agent_eval(case_ids or None))
        except AgentEvalError as exc:
            return jsonify({'ok': False, 'error': str(exc)}), 400
