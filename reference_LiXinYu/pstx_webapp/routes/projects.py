# -*- coding: utf-8 -*-
"""Project analysis and project list API routes."""

from __future__ import annotations

import time
import uuid


def _append_report_payload_timing(bundle: dict, elapsed_seconds: float) -> None:
    timings = bundle.setdefault('analysis_timings', {
        'schema_version': 'pstx-analysis-timings.v1',
        'stages': [],
        'cache': {},
    })
    timings.setdefault('stages', []).append({
        'stage': 'report_payload',
        'elapsed_ms': round(float(elapsed_seconds or 0.0) * 1000.0, 3),
    })
    timings['total_stage_elapsed_ms'] = round(
        sum(float(item.get('elapsed_ms', 0.0) or 0.0) for item in timings.get('stages', [])),
        3,
    )


def register_project_routes(
    app,
    *,
    request,
    jsonify,
    url_for,
    discover_project_files,
    read_local_text_file,
    parse_voltage_map_text,
    parse_checkbox_flag,
    analyze_project_contents,
    remember_run,
    list_project_summaries,
    build_report_payload,
    drc_issue_keys,
) -> None:
    """Register project ingest/list routes."""

    @app.post('/api/analyze')
    def analyze_upload():
        try:
            discovered = discover_project_files(request.form.get('project_root', ''))
            if len(discovered) == 5:
                project_root, prt_path, net_path, ref_path, snapshot_meta = discovered
            else:
                project_root, prt_path, net_path, ref_path = discovered
                snapshot_meta = {}
            prt_text, prt_meta = read_local_text_file(prt_path, 'pstxprt.dat', True)
            net_text, net_meta = read_local_text_file(net_path, 'pstxnet.dat', True)
            ref_text, ref_meta = read_local_text_file(
                ref_path or (project_root / 'packaged' / 'pstxref.dat'),
                'pstxref.dat',
                False,
            )
            project_name = (request.form.get('project_name') or '').strip()
            ratio_limit = float(request.form.get('ratio_limit') or 70)
            custom_volt_map, map_warnings = parse_voltage_map_text(request.form.get('custom_volt_map', ''))
            include_depop = parse_checkbox_flag(request.form.get('include_depop'))
            include_total_bom = parse_checkbox_flag(request.form.get('include_total_bom'))
        except ValueError as exc:
            return jsonify({'ok': False, 'error': str(exc)}), 400
        except Exception as exc:
            return jsonify({'ok': False, 'error': f'参数解析失败：{exc}'}), 400

        run_id = uuid.uuid4().hex[:12]
        bundle = analyze_project_contents(
            prt_text or '',
            net_text or '',
            project_name=project_name or project_root.name,
            project_root=str(project_root),
            ratio_limit=ratio_limit,
            custom_volt_map=custom_volt_map,
            include_depop=include_depop,
            include_total_bom=include_total_bom,
        )
        warnings = list(map_warnings) + list(bundle.get('page_warnings', []))
        warnings.extend(snapshot_meta.get('warnings', []) if isinstance(snapshot_meta, dict) else [])
        if ref_text is not None:
            warnings.append('pstxref.dat 已接收，当前版本仅保留文件记录，暂不参与分析结果。')
        bundle.update({
            'project_name': project_name or '未命名项目',
            'generated_at': time.strftime('%Y-%m-%d %H:%M:%S'),
            'warnings': warnings,
            'input_files': [prt_meta, net_meta, ref_meta],
            'project_input_snapshot': snapshot_meta,
        })
        bundle['project_name'] = project_name or project_root.name
        bundle['project_root'] = str(project_root)
        report_started = time.perf_counter()
        report = build_report_payload(run_id, bundle)
        _append_report_payload_timing(bundle, time.perf_counter() - report_started)
        report['analysis_timings'] = bundle.get('analysis_timings', {})
        payload = {
            'bundle': bundle,
            'report': report,
        }
        remember_run(run_id, payload)
        return jsonify({
            'ok': True,
            'run_id': run_id,
            'redirect_url': url_for('report_page', run_id=run_id),
        })

    @app.get('/api/projects')
    def project_list():
        projects = list_project_summaries(drc_issue_keys=drc_issue_keys)
        return jsonify({
            'ok': True,
            'count': len(projects),
            'projects': projects,
        })
