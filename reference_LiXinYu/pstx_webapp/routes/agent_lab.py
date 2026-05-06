# -*- coding: utf-8 -*-
"""Agent Lab route registration."""

from __future__ import annotations

from pathlib import Path


def _as_bool(value) -> bool:
    return value if isinstance(value, bool) else str(value).strip().lower() in {'1', 'true', 'yes', 'on'}


def register_agent_lab_routes(
    app,
    *,
    request,
    jsonify,
    build_agent_ref_status,
    build_aster_status,
    build_review_checklist_status,
    list_harness_agent_profiles,
    reindex_agent_ref,
    reindex_review_checklists,
    HarnessAgentRequest,
    HarnessError,
    AsterHarnessModelProvider,
    MockHarnessModelProvider,
    new_agent_context,
    run_harness_agent,
) -> None:
    """Register Agent Lab status, reindex, and ask APIs."""

    @app.get('/api/agent-lab/status')
    def agent_lab_status():
        profiles = list_harness_agent_profiles()
        return jsonify({
            'ok': True,
            'mode': 'agent-lab',
            'ref': build_agent_ref_status(),
            'checklist': build_review_checklist_status(),
            'default_profile': 'review_checklist_qa',
            'profiles': [
                profile for profile in profiles
                if profile.get('id') in {'auto', 'agent_ref_qa', 'review_checklist_qa', 'dfmea_prep', 'full_review'}
            ],
            'safeguards': [
                'Agent Lab 默认只读 ref/ PDF 索引，不修改 PDF、不写项目文件。',
                'ref_checklist/ 用于真实 review 问题和 changelist 检索，历史问题只能作为模式参考，不能直接替代当前项目证据。',
                'PDF 内容通过本地页级索引压缩后进入 Agent，上下文只传片段和 evidence id。',
                'Aster 仍只作为模型 provider，不获得本地文件执行权限。',
            ],
        })

    @app.post('/api/agent-lab/ref/reindex')
    def agent_lab_ref_reindex():
        data = request.get_json(silent=True) or request.form.to_dict()
        raw_force = data.get('force', True) if isinstance(data, dict) else True
        force = _as_bool(raw_force)
        try:
            max_files = int((data.get('max_files') if isinstance(data, dict) else None) or 1000)
        except (TypeError, ValueError):
            return jsonify({'ok': False, 'error': 'max_files 必须是数字。'}), 400
        if max_files < 1 or max_files > 50000:
            return jsonify({'ok': False, 'error': 'max_files 必须在 1-50000 之间。'}), 400
        return jsonify(reindex_agent_ref(force=force, max_files=max_files))

    @app.post('/api/agent-lab/checklist/reindex')
    def agent_lab_checklist_reindex():
        data = request.get_json(silent=True) or request.form.to_dict()
        raw_force = data.get('force', True) if isinstance(data, dict) else True
        force = _as_bool(raw_force)
        try:
            max_files = int((data.get('max_files') if isinstance(data, dict) else None) or 1000)
        except (TypeError, ValueError):
            return jsonify({'ok': False, 'error': 'max_files 必须是数字。'}), 400
        if max_files < 1 or max_files > 50000:
            return jsonify({'ok': False, 'error': 'max_files 必须在 1-50000 之间。'}), 400
        return jsonify(reindex_review_checklists(force=force, max_files=max_files))

    @app.post('/api/agent-lab/ask')
    def agent_lab_ask():
        data = request.get_json(silent=True) or request.form.to_dict()
        if isinstance(data, dict):
            data = dict(data)
            data.setdefault('profile', 'review_checklist_qa')
        try:
            agent_request = HarnessAgentRequest.from_mapping(data)
        except HarnessError as exc:
            return jsonify({'ok': False, 'error': str(exc)}), 400
        project_context = new_agent_context()
        synthetic_report = {
            'project_name': 'Agent Capability Lab',
            'summary': 'Agent Lab 使用 ref/ 本地 PDF 和 harness 只读工具测试能力边界。',
            'sections': [],
        }
        synthetic_bundle = {
            'project_name': 'Agent Capability Lab',
            'project_root': str(Path.cwd()),
        }
        try:
            aster_status_payload = build_aster_status()
            provider = (
                MockHarnessModelProvider()
                if aster_status_payload.get('mode') in {'', 'mock'}
                else AsterHarnessModelProvider()
            )
            result = run_harness_agent(
                synthetic_report,
                synthetic_bundle,
                agent_request,
                model_provider=provider,
                project_context=project_context,
            )
        except Exception as exc:
            return jsonify({'ok': False, 'error': f'Agent Lab 执行失败：{exc}'}), 500
        result = dict(result)
        result['lab'] = {
            'ref': build_agent_ref_status(),
            'checklist': build_review_checklist_status(),
            'default_profile': 'review_checklist_qa',
        }
        return jsonify(result), 200 if result.get('ok') else 400
