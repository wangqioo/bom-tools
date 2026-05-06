# -*- coding: utf-8 -*-
"""Report harness review, agent loop, context, and trace replay routes."""

from __future__ import annotations

import copy

from pstx_agent_runtime import AgentBackgroundJob, AgentCheckpointReporter, new_agent_run_id


def register_harness_routes(
    app,
    *,
    request,
    jsonify,
    run_cache,
    agent_context_cache,
    agent_run_cache,
    durable_store,
    background_runner,
    get_agent_context,
    new_agent_context,
    agent_context_public,
    append_agent_context_answers,
    update_agent_context_after_run,
    build_aster_status,
    remember_agent_run,
    HarnessRunRequest,
    HarnessAgentRequest,
    HarnessError,
    AsterHarnessModelProvider,
    MockHarnessModelProvider,
    run_harness_review,
    run_harness_agent,
    build_compare_payload=None,
    CompareAgentRequest=None,
    CompareMockModelProvider=None,
    run_compare_agent=None,
) -> None:
    """Register harness routes for an existing report run."""

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
        for event in payload.get('execution_journal') or []:
            if isinstance(event, dict) and isinstance(event.get('metadata'), dict):
                if event['metadata'].get('agent_run_id') == original:
                    event['metadata']['agent_run_id'] = agent_run_id
        return payload

    def _durable_response(agent_run_id: str, status_code: int = 200):
        status = durable_store.public_status(agent_run_id)
        if not status.get('ok'):
            return jsonify(status), 404
        envelope = agent_run_cache.get_envelope(agent_run_id)
        if envelope and not status.get('agent_run'):
            status['agent_run'] = envelope.get('payload') or {}
            status['trace'] = {key: value for key, value in envelope.items() if key != 'payload'}
            status['result_available'] = True
        elif envelope and status.get('agent_run'):
            status['trace'] = {key: value for key, value in envelope.items() if key != 'payload'}
        return jsonify(status), status_code

    @app.get('/api/harness/agent-runs/<agent_run_id>')
    def harness_agent_run(agent_run_id: str):
        agent_run_id = str(agent_run_id or '').strip()
        durable = durable_store.public_status(agent_run_id)
        if durable.get('ok'):
            envelope = agent_run_cache.get_envelope(agent_run_id)
            if envelope:
                durable['agent_run'] = envelope.get('payload') or durable.get('agent_run') or {}
                durable['trace'] = {key: value for key, value in envelope.items() if key != 'payload'}
                durable['result_available'] = True
            return jsonify(durable)
        envelope = agent_run_cache.get_envelope(agent_run_id)
        if not envelope:
            return jsonify({'ok': False, 'error': f'未找到 agent_run_id：{agent_run_id}'}), 404
        payload = envelope.get('payload') or {}
        trace = {key: value for key, value in envelope.items() if key != 'payload'}
        return jsonify({
            'ok': True,
            'agent_run_id': agent_run_id,
            'agent_run': payload,
            'trace': trace,
        })

    @app.post('/api/harness/agent-runs/<agent_run_id>/cancel')
    def harness_agent_run_cancel(agent_run_id: str):
        record = background_runner.cancel(agent_run_id)
        if not record:
            return jsonify({'ok': False, 'error': f'未找到 agent_run_id：{agent_run_id}'}), 404
        return _durable_response(agent_run_id)

    @app.get('/api/harness/agent-runs/<agent_run_id>/artifacts')
    def harness_agent_run_artifacts(agent_run_id: str):
        artifacts = durable_store.list_artifacts(agent_run_id)
        if not artifacts:
            return jsonify({'ok': False, 'error': f'未找到 agent_run_id：{agent_run_id}'}), 404
        return jsonify({'ok': True, **artifacts})

    @app.post('/api/harness/agent-runs/<agent_run_id>/continue')
    def harness_agent_run_continue(agent_run_id: str):
        record = durable_store.read_record(agent_run_id)
        if not record:
            return jsonify({'ok': False, 'error': f'未找到 agent_run_id：{agent_run_id}'}), 404
        current_status = str(record.get('status') or '')
        if current_status in {'queued', 'running'}:
            return jsonify({'ok': False, 'error': '该 Agent run 仍在执行中，不能重复 continue。'}), 409
        if current_status == 'cancelled':
            return jsonify({'ok': False, 'error': '已取消的 Agent run 不能 continue，请重新提交任务。'}), 400
        request_payload = dict(record.get('request') or {})
        data = request.get_json(silent=True) or request.form.to_dict()
        if isinstance(data, dict) and data.get('question'):
            request_payload['question'] = str(data.get('question') or request_payload.get('question') or '')
        if isinstance(data, dict) and data.get('context_answers') is not None:
            request_payload['context_answers'] = data.get('context_answers')
        request_payload['continue_agent_run_id'] = agent_run_id

        if record.get('kind') == 'report':
            try:
                agent_request = HarnessAgentRequest.from_mapping(request_payload)
            except HarnessError as exc:
                return jsonify({'ok': False, 'error': str(exc)}), 400
            run_id = str(request_payload.get('run_id') or record.get('scope_id') or '').strip()
            payload = run_cache.get(run_id)
            if not payload:
                return jsonify({'ok': False, 'error': f'未找到报告 run_id：{run_id}'}), 404

            def job(job_run_id: str) -> dict:
                reporter = AgentCheckpointReporter(durable_store, job_run_id, scope_id=run_id, kind='report')
                project_context = get_agent_context(run_id)
                project_context['run_id'] = run_id
                project_context['agent_workspace_scope_id'] = run_id
                project_context['agent_workspace_agent_run_id'] = job_run_id
                continue_envelope = agent_run_cache.get_envelope(agent_run_id)
                project_context['active_continuation_pack'] = dict((continue_envelope or {}).get('continuation_pack') or record.get('continuation_pack') or {})
                append_agent_context_answers(project_context, agent_request.context_answers, source_agent_run_id=agent_run_id)
                aster_status_payload = build_aster_status()
                provider = MockHarnessModelProvider() if aster_status_payload.get('mode') in {'', 'mock'} else AsterHarnessModelProvider()
                result = run_harness_agent(
                    payload['report'],
                    payload['bundle'],
                    agent_request,
                    model_provider=provider,
                    project_context=project_context,
                    checkpoint_callback=reporter.emit,
                    should_cancel=reporter.cancel_requested,
                    resume_context=record,
                )
                result = _normalize_agent_run_id(result, job_run_id)
                result['run_id'] = run_id
                result['project_name'] = payload['report'].get('project_name') or payload['bundle'].get('project_name') or ''
                result['request'] = request_payload
                update_agent_context_after_run(run_id, project_context, result)
                result['project_context_summary'] = agent_context_public(run_id, project_context)
                remember_agent_run(job_run_id, result)
                return result
            scope_id = str(record.get('scope_id') or run_id)
            kind = 'report'
        elif record.get('kind') == 'compare':
            if not (build_compare_payload and CompareAgentRequest and CompareMockModelProvider and run_compare_agent):
                return jsonify({'ok': False, 'error': 'Compare Agent continue handler 未注册。'}), 500
            try:
                compare_request = CompareAgentRequest.from_mapping(request_payload)
            except HarnessError as exc:
                return jsonify({'ok': False, 'error': str(exc)}), 400
            left_run_id = str(request_payload.get('left_run_id') or '').strip()
            right_run_id = str(request_payload.get('right_run_id') or '').strip()
            if left_run_id not in run_cache or right_run_id not in run_cache:
                return jsonify({'ok': False, 'error': '未找到用于继续 Compare Agent 的 A/B 项目。'}), 404

            def job(job_run_id: str) -> dict:
                reporter = AgentCheckpointReporter(durable_store, job_run_id, scope_id=record.get('scope_id') or f'compare_{left_run_id}_vs_{right_run_id}', kind='compare')
                compare_payload = build_compare_payload(left_run_id, right_run_id, detail_limit=compare_request.detail_limit)
                compare_payload['_agent_workspace_scope_id'] = record.get('scope_id') or f'compare_{left_run_id}_vs_{right_run_id}'
                compare_payload['_agent_workspace_agent_run_id'] = job_run_id
                aster_status_payload = build_aster_status()
                provider = CompareMockModelProvider() if aster_status_payload.get('mode') in {'', 'mock'} else AsterHarnessModelProvider()
                result = run_compare_agent(
                    compare_payload,
                    run_cache[left_run_id],
                    run_cache[right_run_id],
                    compare_request,
                    model_provider=provider,
                    checkpoint_callback=reporter.emit,
                    should_cancel=reporter.cancel_requested,
                    resume_context=record,
                )
                result = _normalize_agent_run_id(dict(result), job_run_id)
                result['left_run_id'] = left_run_id
                result['right_run_id'] = right_run_id
                result['left'] = compare_payload.get('left', {})
                result['right'] = compare_payload.get('right', {})
                result['diff_totals'] = compare_payload.get('diff_totals', {})
                result['request'] = request_payload
                remember_agent_run(job_run_id, result)
                return result
            scope_id = str(record.get('scope_id') or f'compare_{left_run_id}_vs_{right_run_id}')
            kind = 'compare'
        else:
            return jsonify({'ok': False, 'error': f"未知 Agent run 类型：{record.get('kind')}"}), 400

        durable_store.update_record(agent_run_id, request=request_payload, status='queued', current_phase='continue_queued', error='', checkpoint={'phase': 'continue_queued'})
        try:
            background_runner.submit(AgentBackgroundJob(
                agent_run_id=agent_run_id,
                scope_id=scope_id,
                kind=kind,
                run=job,
            ))
        except RuntimeError as exc:
            return jsonify({'ok': False, 'error': str(exc)}), 429
        return _durable_response(agent_run_id, status_code=202)

    @app.get('/api/report/<run_id>/harness/context')
    def harness_project_context(run_id: str):
        if run_id not in run_cache:
            return jsonify({'ok': False, 'error': f'未找到报告 run_id：{run_id}'}), 404
        context = get_agent_context(run_id)
        return jsonify({
            'ok': True,
            'context': agent_context_public(run_id, context),
        })

    @app.post('/api/report/<run_id>/harness/context/clear')
    def harness_project_context_clear(run_id: str):
        if run_id not in run_cache:
            return jsonify({'ok': False, 'error': f'未找到报告 run_id：{run_id}'}), 404
        agent_context_cache[run_id] = new_agent_context()
        return jsonify({
            'ok': True,
            'context': agent_context_public(run_id, agent_context_cache[run_id]),
        })

    @app.post('/api/report/<run_id>/harness/review')
    def harness_review(run_id: str):
        payload = run_cache.get(run_id)
        if not payload:
            return jsonify({'ok': False, 'error': f'未找到报告 run_id：{run_id}'}), 404
        data = request.get_json(silent=True) or request.form.to_dict()
        try:
            harness_request = HarnessRunRequest.from_mapping(data)
        except HarnessError as exc:
            return jsonify({'ok': False, 'error': str(exc)}), 400
        try:
            result = run_harness_review(
                payload['report'],
                payload['bundle'],
                harness_request,
                model_provider=AsterHarnessModelProvider(),
            )
        except Exception as exc:
            return jsonify({'ok': False, 'error': f'Harness 审查失败：{exc}'}), 500
        return jsonify(result)

    @app.post('/api/report/<run_id>/harness/agent')
    def harness_agent(run_id: str):
        payload = run_cache.get(run_id)
        if not payload:
            return jsonify({'ok': False, 'error': f'未找到报告 run_id：{run_id}'}), 404
        data = request.get_json(silent=True) or request.form.to_dict()
        try:
            agent_request = HarnessAgentRequest.from_mapping(data)
        except HarnessError as exc:
            return jsonify({'ok': False, 'error': str(exc)}), 400
        continue_envelope = None
        if agent_request.continue_agent_run_id:
            continue_envelope = agent_run_cache.get_envelope(agent_request.continue_agent_run_id)
        if agent_request.continue_agent_run_id and not continue_envelope:
            return jsonify({
                'ok': False,
                'error': f'未找到 continue_agent_run_id：{agent_request.continue_agent_run_id}',
            }), 400
        project_context = get_agent_context(run_id)
        project_context['run_id'] = run_id
        if continue_envelope:
            project_context['active_continuation_pack'] = dict(continue_envelope.get('continuation_pack') or {})
        else:
            project_context['active_continuation_pack'] = {}
        append_agent_context_answers(
            project_context,
            agent_request.context_answers,
            source_agent_run_id=agent_request.continue_agent_run_id,
        )
        request_summary = {
            'profile': agent_request.profile,
            'question': agent_request.question,
            'max_steps': agent_request.max_steps,
            'max_tool_calls': agent_request.max_tool_calls,
            'max_rows_per_table': agent_request.max_rows_per_table,
            'debug': agent_request.debug,
            'enable_subagents': agent_request.enable_subagents,
            'subagent_profiles': list(agent_request.subagent_profiles),
            'max_subagents': agent_request.max_subagents,
            'context_answer_count': len(agent_request.context_answers),
            'continue_agent_run_id': agent_request.continue_agent_run_id,
        }
        raw_async = data.get('async') if isinstance(data, dict) else False
        if _as_bool(raw_async):
            agent_run_id = new_agent_run_id('report')
            durable_store.create_run(
                scope_id=run_id,
                kind='report',
                request={**request_summary, 'run_id': run_id, 'context_answers': list(agent_request.context_answers)},
                agent_run_id=agent_run_id,
            )

            def job(job_run_id: str) -> dict:
                reporter = AgentCheckpointReporter(durable_store, job_run_id, scope_id=run_id, kind='report')
                async_project_context = copy.deepcopy(project_context)
                async_project_context['run_id'] = run_id
                async_project_context['agent_workspace_scope_id'] = run_id
                async_project_context['agent_workspace_agent_run_id'] = job_run_id

                def dispatch_child_tasks(dispatch_request: dict) -> dict:
                    parent_record = durable_store.read_record(job_run_id, scope_id=run_id)
                    root_agent_run_id = str(parent_record.get('root_agent_run_id') or job_run_id) if parent_record else job_run_id
                    child_records = []
                    for index, task in enumerate(dispatch_request.get('tasks') or [], start=1):
                        if not isinstance(task, dict):
                            continue
                        child_profile = str(task.get('profile') or 'auto').strip() or 'auto'
                        child_request_payload = {
                            'profile': child_profile,
                            'question': str(task.get('question') or task.get('title') or '').strip(),
                            'max_steps': int(task.get('max_steps') or max(1, min(agent_request.max_steps, 8))),
                            'max_tool_calls': int(task.get('max_tool_calls') or max(1, min(agent_request.max_tool_calls, 14))),
                            'max_rows_per_table': agent_request.max_rows_per_table,
                            'debug': agent_request.debug,
                            'enable_subagents': False,
                            'parent_agent_run_id': job_run_id,
                            'root_agent_run_id': root_agent_run_id,
                            'dispatch_task': dict(task),
                            'dispatch_task_id': str(task.get('task_id') or f'task-{index}'),
                            'run_id': run_id,
                        }
                        try:
                            child_agent_request = HarnessAgentRequest.from_mapping(child_request_payload)
                        except HarnessError:
                            child_request_payload['profile'] = 'auto'
                            child_agent_request = HarnessAgentRequest.from_mapping(child_request_payload)
                        child_run_id = new_agent_run_id('report')
                        durable_store.create_run(
                            scope_id=run_id,
                            kind='report',
                            request=child_request_payload,
                            agent_run_id=child_run_id,
                            parent_agent_run_id=job_run_id,
                            root_agent_run_id=root_agent_run_id,
                            dispatch_task=task,
                            dispatch_group_id=f'{job_run_id}-dispatch',
                        )

                        def child_job_factory(child_request, child_payload, child_task):
                            def child_job(child_job_run_id: str) -> dict:
                                child_reporter = AgentCheckpointReporter(durable_store, child_job_run_id, scope_id=run_id, kind='report')
                                child_context = copy.deepcopy(async_project_context)
                                child_context['run_id'] = run_id
                                child_context['agent_workspace_scope_id'] = run_id
                                child_context['agent_workspace_agent_run_id'] = child_job_run_id
                                child_context['parent_agent_run_id'] = job_run_id
                                child_context['root_agent_run_id'] = root_agent_run_id
                                child_context['dispatch_task'] = dict(child_task)
                                try:
                                    aster_status_payload = build_aster_status()
                                    provider = (
                                        MockHarnessModelProvider()
                                        if aster_status_payload.get('mode') in {'', 'mock'}
                                        else AsterHarnessModelProvider()
                                    )
                                    result = run_harness_agent(
                                        payload['report'],
                                        payload['bundle'],
                                        child_request,
                                        model_provider=provider,
                                        project_context=child_context,
                                        checkpoint_callback=child_reporter.emit,
                                        should_cancel=child_reporter.cancel_requested,
                                    )
                                except Exception as exc:
                                    return {'ok': False, 'status': 'failed', 'answer': '', 'error': f'Harness child agent 执行失败：{exc}', 'agent_run_id': child_job_run_id}
                                result = _normalize_agent_run_id(dict(result), child_job_run_id)
                                result['run_id'] = run_id
                                result['project_name'] = payload['report'].get('project_name') or payload['bundle'].get('project_name') or ''
                                result['parent_agent_run_id'] = job_run_id
                                result['root_agent_run_id'] = root_agent_run_id
                                result['dispatch_task'] = dict(child_task)
                                result['request'] = child_payload
                                result['project_context_summary'] = agent_context_public(run_id, child_context)
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
                                scope_id=run_id,
                                kind='report',
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
                        'source': 'report',
                    }
                    durable_store.append_child_runs(job_run_id, child_records, scope_id=run_id, task_dispatch_summary=summary)
                    return {'task_dispatch_summary': summary, 'dispatched_tasks': child_records}

                try:
                    aster_status_payload = build_aster_status()
                    provider = (
                        MockHarnessModelProvider()
                        if aster_status_payload.get('mode') in {'', 'mock'}
                        else AsterHarnessModelProvider()
                    )
                    result = run_harness_agent(
                        payload['report'],
                        payload['bundle'],
                        agent_request,
                        model_provider=provider,
                        project_context=async_project_context,
                        checkpoint_callback=reporter.emit,
                        should_cancel=reporter.cancel_requested,
                        dispatch_callback=dispatch_child_tasks,
                    )
                except Exception as exc:
                    return {'ok': False, 'status': 'failed', 'answer': '', 'error': f'Harness agent 执行失败：{exc}', 'agent_run_id': job_run_id}
                result = _normalize_agent_run_id(dict(result), job_run_id)
                result['run_id'] = run_id
                result['project_name'] = payload['report'].get('project_name') or payload['bundle'].get('project_name') or ''
                result['request'] = request_summary
                update_agent_context_after_run(run_id, async_project_context, result)
                agent_context_cache[run_id] = async_project_context
                result['project_context_summary'] = agent_context_public(run_id, async_project_context)
                remember_agent_run(job_run_id, result)
                return result

            try:
                background_runner.submit(AgentBackgroundJob(
                    agent_run_id=agent_run_id,
                    scope_id=run_id,
                    kind='report',
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
            aster_status_payload = build_aster_status()
            provider = (
                MockHarnessModelProvider()
                if aster_status_payload.get('mode') in {'', 'mock'}
                else AsterHarnessModelProvider()
            )
            result = run_harness_agent(
                payload['report'],
                payload['bundle'],
                agent_request,
                model_provider=provider,
                project_context=project_context,
            )
        except Exception as exc:
            return jsonify({'ok': False, 'error': f'Harness agent 执行失败：{exc}'}), 500
        result = dict(result)
        result['run_id'] = run_id
        result['project_name'] = payload['report'].get('project_name') or payload['bundle'].get('project_name') or ''
        result['request'] = request_summary
        update_agent_context_after_run(run_id, project_context, result)
        result['project_context_summary'] = agent_context_public(run_id, project_context)
        remember_agent_run(str(result.get('agent_run_id') or ''), result)
        status_code = 200 if result.get('ok') else 400
        return jsonify(result), status_code
