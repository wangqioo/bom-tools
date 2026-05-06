# -*- coding: utf-8 -*-
"""Project compare API and Compare Agent routes."""

from __future__ import annotations

from pstx_agent_runtime import AgentBackgroundJob, AgentCheckpointReporter, new_agent_run_id


def register_compare_routes(
    app,
    *,
    request,
    jsonify,
    run_cache,
    durable_store,
    background_runner,
    build_compare_payload,
    coerce_compare_detail_limit,
    list_compare_agent_profiles,
    CompareAgentRequest,
    CompareMockModelProvider,
    HarnessError,
    AsterHarnessModelProvider,
    build_aster_status,
    run_compare_agent,
    remember_agent_run,
) -> None:
    """Register project compare and compare-agent routes."""

    def _as_bool(value) -> bool:
        return value if isinstance(value, bool) else str(value or '').strip().lower() in {'1', 'true', 'yes', 'on'}

    def _normalize_agent_run_id(result: dict, agent_run_id: str) -> dict:
        payload = dict(result or {})
        original = str(payload.get('agent_run_id') or '')
        payload['agent_run_id'] = agent_run_id
        if original and original != agent_run_id:
            payload['original_agent_run_id'] = original
        for key in ('trace_summary', 'model_metadata', 'continuation_pack'):
            if isinstance(payload.get(key), dict):
                payload[key] = dict(payload[key])
                payload[key]['agent_run_id'] = agent_run_id
        return payload

    @app.post('/api/compare')
    def compare_projects():
        data = request.get_json(silent=True) or request.form
        left_run_id = str(data.get('left_run_id') or '').strip()
        right_run_id = str(data.get('right_run_id') or '').strip()
        if not left_run_id or not right_run_id:
            return jsonify({'ok': False, 'error': '请选择两个项目后再对比。'}), 400
        if left_run_id == right_run_id:
            return jsonify({'ok': False, 'error': '请选择两个不同项目进行对比。'}), 400
        try:
            detail_limit = coerce_compare_detail_limit(data.get('detail_limit'))
        except ValueError as exc:
            return jsonify({'ok': False, 'error': str(exc)}), 400
        if left_run_id not in run_cache or right_run_id not in run_cache:
            return jsonify({'ok': False, 'error': '未找到用于对比的项目，请重新分析或刷新项目列表。'}), 404
        return jsonify(build_compare_payload(left_run_id, right_run_id, detail_limit=detail_limit))

    @app.get('/api/compare/harness/profiles')
    def compare_harness_profiles():
        return jsonify({
            'ok': True,
            'mode': 'local-compare-agent-harness',
            'profiles': list_compare_agent_profiles(),
            'default_profile': 'compare_quick_scan',
            'safeguards': [
                'Compare profile 只限制本地 compare harness 可调用的只读工具集合。',
                'Aster 仍只作为模型 provider，不接收任何本地执行权限。',
                '项目文件读取仅限 A/B run 对应 project_root 的白名单路径。',
            ],
        })

    @app.post('/api/compare/harness-agent')
    def compare_harness_agent():
        data = request.get_json(silent=True) or request.form.to_dict()
        left_run_id = str(data.get('left_run_id') or '').strip()
        right_run_id = str(data.get('right_run_id') or '').strip()
        if not left_run_id or not right_run_id:
            return jsonify({'ok': False, 'error': '请选择两个项目后再提问。'}), 400
        if left_run_id == right_run_id:
            return jsonify({'ok': False, 'error': '请选择两个不同项目进行 Compare Agent 审查。'}), 400
        if left_run_id not in run_cache or right_run_id not in run_cache:
            return jsonify({'ok': False, 'error': '未找到用于对比的项目，请重新分析或刷新项目列表。'}), 404
        try:
            agent_request = CompareAgentRequest.from_mapping(data)
        except HarnessError as exc:
            return jsonify({'ok': False, 'error': str(exc)}), 400
        request_summary = {
            'profile': agent_request.profile,
            'question': agent_request.question,
            'max_steps': agent_request.max_steps,
            'max_tool_calls': agent_request.max_tool_calls,
            'detail_limit': agent_request.detail_limit,
            'debug': agent_request.debug,
        }
        if _as_bool(data.get('async') if isinstance(data, dict) else False):
            agent_run_id = new_agent_run_id('compare')
            scope_id = f"compare_{left_run_id}_vs_{right_run_id}"
            durable_store.create_run(
                scope_id=scope_id,
                kind='compare',
                request={
                    **request_summary,
                    'left_run_id': left_run_id,
                    'right_run_id': right_run_id,
                },
                agent_run_id=agent_run_id,
            )

            def job(job_run_id: str) -> dict:
                reporter = AgentCheckpointReporter(durable_store, job_run_id, scope_id=scope_id, kind='compare')

                def dispatch_child_tasks(dispatch_request: dict) -> dict:
                    parent_record = durable_store.read_record(job_run_id, scope_id=scope_id)
                    root_agent_run_id = str(parent_record.get('root_agent_run_id') or job_run_id) if parent_record else job_run_id
                    child_records = []
                    for index, task in enumerate(dispatch_request.get('tasks') or [], start=1):
                        if not isinstance(task, dict):
                            continue
                        child_request_payload = {
                            'profile': str(task.get('profile') or 'auto').strip() or 'auto',
                            'question': str(task.get('question') or task.get('title') or '').strip(),
                            'max_steps': int(task.get('max_steps') or max(1, min(agent_request.max_steps, 8))),
                            'max_tool_calls': int(task.get('max_tool_calls') or max(1, min(agent_request.max_tool_calls, 14))),
                            'detail_limit': agent_request.detail_limit,
                            'debug': agent_request.debug,
                            'parent_agent_run_id': job_run_id,
                            'root_agent_run_id': root_agent_run_id,
                            'dispatch_task': dict(task),
                            'dispatch_task_id': str(task.get('task_id') or f'task-{index}'),
                            'left_run_id': left_run_id,
                            'right_run_id': right_run_id,
                        }
                        try:
                            child_agent_request = CompareAgentRequest.from_mapping(child_request_payload)
                        except HarnessError:
                            child_request_payload['profile'] = 'auto'
                            child_agent_request = CompareAgentRequest.from_mapping(child_request_payload)
                        child_run_id = new_agent_run_id('compare')
                        durable_store.create_run(
                            scope_id=scope_id,
                            kind='compare',
                            request=child_request_payload,
                            agent_run_id=child_run_id,
                            parent_agent_run_id=job_run_id,
                            root_agent_run_id=root_agent_run_id,
                            dispatch_task=task,
                            dispatch_group_id=f'{job_run_id}-dispatch',
                        )

                        def child_job_factory(child_request, child_payload, child_task):
                            def child_job(child_job_run_id: str) -> dict:
                                child_reporter = AgentCheckpointReporter(durable_store, child_job_run_id, scope_id=scope_id, kind='compare')
                                try:
                                    child_compare_payload = build_compare_payload(
                                        left_run_id,
                                        right_run_id,
                                        detail_limit=child_request.detail_limit,
                                    )
                                    child_compare_payload['_agent_workspace_scope_id'] = scope_id
                                    child_compare_payload['_agent_workspace_agent_run_id'] = child_job_run_id
                                    aster_status_payload = build_aster_status()
                                    provider = (
                                        CompareMockModelProvider()
                                        if aster_status_payload.get('mode') in {'', 'mock'}
                                        else AsterHarnessModelProvider()
                                    )
                                    result = run_compare_agent(
                                        child_compare_payload,
                                        run_cache[left_run_id],
                                        run_cache[right_run_id],
                                        child_request,
                                        model_provider=provider,
                                        checkpoint_callback=child_reporter.emit,
                                        should_cancel=child_reporter.cancel_requested,
                                    )
                                except Exception as exc:
                                    return {'ok': False, 'status': 'failed', 'answer': '', 'error': f'Compare child agent 执行失败：{exc}', 'agent_run_id': child_job_run_id}
                                result = _normalize_agent_run_id(dict(result), child_job_run_id)
                                result['left_run_id'] = left_run_id
                                result['right_run_id'] = right_run_id
                                result['left'] = child_compare_payload.get('left', {})
                                result['right'] = child_compare_payload.get('right', {})
                                result['diff_totals'] = child_compare_payload.get('diff_totals', {})
                                result['parent_agent_run_id'] = job_run_id
                                result['root_agent_run_id'] = root_agent_run_id
                                result['dispatch_task'] = dict(child_task)
                                result['request'] = child_payload
                                remember_agent_run(child_job_run_id, result)
                                return result
                            return child_job

                        child_record = {
                            'task_id': str(task.get('task_id') or f'task-{index}'),
                            'title': str(task.get('title') or task.get('question') or f'Task {index}')[:180],
                            'profile': child_agent_request.profile,
                            'question': child_agent_request.question,
                            'agent_run_id': child_run_id,
                            'status': 'queued',
                            'status_url': f'/api/harness/agent-runs/{child_run_id}',
                        }
                        try:
                            background_runner.submit(AgentBackgroundJob(
                                agent_run_id=child_run_id,
                                scope_id=scope_id,
                                kind='compare',
                                run=child_job_factory(child_agent_request, child_request_payload, task),
                            ))
                        except RuntimeError as exc:
                            durable_store.fail_record(child_run_id, exc)
                            child_record['status'] = 'failed'
                            child_record['error'] = str(exc)
                        child_records.append(child_record)
                    summary = {
                        'schema_version': dispatch_request.get('schema_version') or 'pstx-agent-task-dispatch.v1',
                        'available': True,
                        'task_count': len(dispatch_request.get('tasks') or []),
                        'dispatched_count': len([item for item in child_records if item.get('status') != 'failed']),
                        'child_count': len(child_records),
                        'source': 'compare',
                    }
                    durable_store.append_child_runs(job_run_id, child_records, scope_id=scope_id, task_dispatch_summary=summary)
                    return {'task_dispatch_summary': summary, 'dispatched_tasks': child_records}

                try:
                    compare_payload = build_compare_payload(
                        left_run_id,
                        right_run_id,
                        detail_limit=agent_request.detail_limit,
                    )
                    compare_payload['_agent_workspace_scope_id'] = scope_id
                    compare_payload['_agent_workspace_agent_run_id'] = job_run_id
                    aster_status_payload = build_aster_status()
                    provider = (
                        CompareMockModelProvider()
                        if aster_status_payload.get('mode') in {'', 'mock'}
                        else AsterHarnessModelProvider()
                    )
                    result = run_compare_agent(
                        compare_payload,
                        run_cache[left_run_id],
                        run_cache[right_run_id],
                        agent_request,
                        model_provider=provider,
                        checkpoint_callback=reporter.emit,
                        should_cancel=reporter.cancel_requested,
                        dispatch_callback=dispatch_child_tasks,
                    )
                except Exception as exc:
                    return {'ok': False, 'status': 'failed', 'answer': '', 'error': f'Compare agent 执行失败：{exc}', 'agent_run_id': job_run_id}
                result = _normalize_agent_run_id(dict(result), job_run_id)
                result['left_run_id'] = left_run_id
                result['right_run_id'] = right_run_id
                result['left'] = compare_payload.get('left', {})
                result['right'] = compare_payload.get('right', {})
                result['diff_totals'] = compare_payload.get('diff_totals', {})
                result['request'] = request_summary
                remember_agent_run(job_run_id, result)
                return result

            try:
                background_runner.submit(AgentBackgroundJob(
                    agent_run_id=agent_run_id,
                    scope_id=scope_id,
                    kind='compare',
                    run=job,
                ))
            except RuntimeError as exc:
                return jsonify({'ok': False, 'error': str(exc)}), 429
            return jsonify({
                'ok': True,
                'async': True,
                'status': 'queued',
                'agent_run_id': agent_run_id,
                'status_url': f'/api/harness/agent-runs/{agent_run_id}',
                'trace_url': f'/api/harness/agent-runs/{agent_run_id}',
            }), 202

        try:
            compare_payload = build_compare_payload(
                left_run_id,
                right_run_id,
                detail_limit=agent_request.detail_limit,
            )
            aster_status_payload = build_aster_status()
            provider = (
                CompareMockModelProvider()
                if aster_status_payload.get('mode') in {'', 'mock'}
                else AsterHarnessModelProvider()
            )
            result = run_compare_agent(
                compare_payload,
                run_cache[left_run_id],
                run_cache[right_run_id],
                agent_request,
                model_provider=provider,
            )
        except Exception as exc:
            return jsonify({'ok': False, 'error': f'Compare agent 执行失败：{exc}'}), 500
        result = dict(result)
        result['left_run_id'] = left_run_id
        result['right_run_id'] = right_run_id
        result['left'] = compare_payload.get('left', {})
        result['right'] = compare_payload.get('right', {})
        result['diff_totals'] = compare_payload.get('diff_totals', {})
        result['request'] = request_summary
        remember_agent_run(str(result.get('agent_run_id') or ''), result)
        status_code = 200 if result.get('ok') else 400
        return jsonify(result), status_code
