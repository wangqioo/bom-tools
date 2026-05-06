# -*- coding: utf-8 -*-
"""Report data, export, query, Aster summary, and report-local preview routes."""

from __future__ import annotations

import io
import os
import tempfile
import time
from pathlib import Path


def register_report_routes(
    app,
    *,
    request,
    jsonify,
    send_file,
    get_run,
    build_aster_summary,
    aster_error_payload,
    match_rows_with_feishu_cache,
    query_project_data,
    export_to_excel,
    export_bom_to_excel,
    build_llm_topology_netlist,
    build_cadence_page_payload,
    build_cadence_index_payload,
    build_csa_geometry_payload,
    build_csa_connectivity_overlay,
    feishu_bom_row_sources,
    max_feishu_preview_rows,
) -> None:
    """Register routes bound to an existing report run."""

    def _safe_download_stem(value: str, fallback: str = "pstx") -> str:
        cleaned = "".join(
            char if (char.isalnum() or char in {"-", "_", "."}) else "_"
            for char in str(value or "").strip()
        ).strip("._")
        return cleaned or fallback

    @app.get('/api/report/<run_id>')
    def report_data(run_id: str):
        payload = get_run(run_id)
        return jsonify(payload['report'])

    @app.get('/api/report/<run_id>/aster-summary')
    def aster_summary(run_id: str):
        payload = get_run(run_id)
        try:
            return jsonify(build_aster_summary(payload['report'], payload['bundle']))
        except Exception as exc:
            error_payload, status = aster_error_payload(exc)
            return jsonify(error_payload), status

    @app.post('/api/report/<run_id>/feishu-bom/preview')
    def feishu_bom_preview(run_id: str):
        payload = get_run(run_id)
        data = request.get_json(silent=True) or request.form
        key_field = str(data.get('key_field') or '').strip()
        source = str(data.get('source') or 'bom_normal_detail').strip()
        match_mode = str(data.get('match_mode') or 'auto').strip() or 'auto'
        try:
            limit = int(data.get('limit') or max_feishu_preview_rows)
        except (TypeError, ValueError):
            return jsonify({'ok': False, 'error': 'limit 必须是数字。'}), 400
        limit = max(1, min(limit, max_feishu_preview_rows))
        if source not in feishu_bom_row_sources:
            return jsonify({
                'ok': False,
                'error': f'不支持的 BOM 来源：{source}',
                'allowed_sources': sorted(feishu_bom_row_sources),
            }), 400
        if not key_field:
            return jsonify({'ok': False, 'error': '请提供 key_field。'}), 400
        rows = list(payload['bundle'].get(source, []) or [])
        if rows and key_field not in rows[0]:
            return jsonify({
                'ok': False,
                'error': f'当前 BOM 来源中不存在字段：{key_field}',
                'available_fields': list(rows[0].keys()),
            }), 400
        try:
            result = match_rows_with_feishu_cache(rows, key_field, limit=limit, match_mode=match_mode)
        except Exception as exc:
            return jsonify({'ok': False, 'error': f'飞书 BOM 匹配预览失败：{exc}'}), 500
        result.update({
            'source': source,
            'source_label': feishu_bom_row_sources[source],
        })
        status_code = 200 if result.get('ok') else 400
        return jsonify(result), status_code

    @app.post('/api/report/<run_id>/query')
    def query_report(run_id: str):
        payload = get_run(run_id)
        data = request.get_json(silent=True) or {}
        mode = data.get('mode') or '位号'
        keyword = data.get('keyword') or ''
        result = query_project_data(payload['bundle']['components'], payload['bundle']['nets'], mode, keyword)
        return jsonify(result)

    @app.get('/api/report/<run_id>/topology')
    def report_topology(run_id: str):
        started = time.perf_counter()
        payload = get_run(run_id)

        def bool_arg(name: str, default: bool = False) -> bool:
            raw = str(request.args.get(name, '') or '').strip().lower()
            if not raw:
                return default
            return raw in {'1', 'true', 'yes', 'on'}

        try:
            limit = int(request.args.get('limit') or 120)
        except (TypeError, ValueError):
            return jsonify({'ok': False, 'error': 'limit 必须是数字。'}), 400
        limit = max(1, min(limit, 250))
        try:
            supply_limit = int(request.args.get('supply_limit') or 12)
        except (TypeError, ValueError):
            return jsonify({'ok': False, 'error': 'supply_limit 必须是数字。'}), 400
        supply_limit = max(0, min(supply_limit, 250))
        full_edges = bool_arg('full_edges', False)
        view = str(request.args.get('view') or ('full' if full_edges else 'summary')).strip().lower()
        if view not in {'summary', 'full'}:
            return jsonify({'ok': False, 'error': 'view 必须是 summary 或 full。'}), 400
        supply_mode = str(request.args.get('supply_mode') or ('details' if full_edges else 'grouped')).strip().lower()
        if supply_mode not in {'grouped', 'details', 'hidden'}:
            return jsonify({'ok': False, 'error': 'supply_mode 必须是 grouped、details 或 hidden。'}), 400
        edge_label_mode = str(request.args.get('edge_label_mode') or 'auto').strip().lower()
        if edge_label_mode not in {'auto', 'off', 'all'}:
            return jsonify({'ok': False, 'error': 'edge_label_mode 必须是 auto、off 或 all。'}), 400
        topology = build_llm_topology_netlist(
            payload.get('report', {}) or {},
            payload.get('bundle', {}) or {},
            focus_refdes=str(request.args.get('focus_refdes') or '').strip(),
            role_filter=str(request.args.get('role_filter') or '').strip(),
            include_connectors=bool_arg('include_connectors', False),
            limit=limit,
            return_all_edges=full_edges,
            view=view,
            supply_mode=supply_mode,
            supply_limit=supply_limit,
        )
        elapsed_ms = round((time.perf_counter() - started) * 1000.0, 3)
        return jsonify({
            'ok': True,
            'run_id': run_id,
            'project_name': (payload.get('report', {}) or {}).get('project_name', ''),
            'topology': topology,
            'topology_cache_status': topology.get('topology_cache_status', {}),
            'topology_timing': {'elapsed_ms': elapsed_ms},
            'edge_label_mode': edge_label_mode,
        })

    @app.get('/api/report/<run_id>/cadence-page')
    def report_cadence_page(run_id: str):
        payload = get_run(run_id)
        try:
            page = int(request.args.get('page') or 0)
            limit = int(request.args.get('limit') or 200)
        except (TypeError, ValueError):
            return jsonify({'ok': False, 'error': 'page 和 limit 必须是数字。'}), 400
        try:
            cadence_page = build_cadence_page_payload(
                (payload.get('bundle', {}) or {}).get('project_root', ''),
                page,
                stdout=str(request.args.get('stdout') or 'summary').strip() or 'summary',
                object_id=str(request.args.get('object_id') or '').strip(),
                limit=max(1, min(limit, 5000)),
            )
        except Exception as exc:
            return jsonify({'ok': False, 'error': str(exc)}), 400
        return jsonify({
            'ok': True,
            'run_id': run_id,
            'project_name': (payload.get('report', {}) or {}).get('project_name', ''),
            'cadence_page': cadence_page,
        })

    @app.get('/api/report/<run_id>/csa-geometry')
    def report_csa_geometry(run_id: str):
        payload = get_run(run_id)
        try:
            limit = int(request.args.get('limit') or 200)
        except (TypeError, ValueError):
            return jsonify({'ok': False, 'error': 'limit 必须是数字。'}), 400
        try:
            page = int(request.args.get('page') or 0)
        except (TypeError, ValueError):
            return jsonify({'ok': False, 'error': 'page 必须是数字。'}), 400
        include_connectivity = str(request.args.get('include_connectivity') or '').strip().lower() in {
            '1', 'true', 'yes', 'on'
        }
        bundle = (payload.get('bundle', {}) or {})
        geometry = bundle.get('csa_geometry', {}) or {}
        stdout = str(request.args.get('stdout') or 'summary').strip() or 'summary'
        row_limit = max(1, min(limit, 5000))
        try:
            semantic_overlay = None
            if include_connectivity:
                semantic_overlay = build_csa_connectivity_overlay(
                    geometry,
                    source_root=str(bundle.get('project_root') or ''),
                    page=page,
                    stdout=stdout,
                    limit=row_limit,
                )
            csa_payload = build_csa_geometry_payload(
                geometry,
                stdout=stdout,
                limit=row_limit,
                page=page,
                semantic_overlay=semantic_overlay,
            )
        except Exception as exc:
            return jsonify({'ok': False, 'error': str(exc)}), 400
        return jsonify({
            'ok': True,
            'run_id': run_id,
            'project_name': (payload.get('report', {}) or {}).get('project_name', ''),
            'csa_geometry': csa_payload,
        })

    @app.get('/api/report/<run_id>/cadence-index')
    def report_cadence_index(run_id: str):
        payload = get_run(run_id)
        try:
            limit = int(request.args.get('limit') or 200)
        except (TypeError, ValueError):
            return jsonify({'ok': False, 'error': 'limit 必须是数字。'}), 400
        try:
            page = int(request.args.get('page') or 0)
        except (TypeError, ValueError):
            return jsonify({'ok': False, 'error': 'page 必须是数字。'}), 400
        bundle = (payload.get('bundle', {}) or {})
        try:
            cadence_index = build_cadence_index_payload(
                str(bundle.get('project_root') or ''),
                pstx_nets=bundle.get('nets', {}) or {},
                stdout=str(request.args.get('stdout') or 'summary').strip() or 'summary',
                query=str(request.args.get('query') or ''),
                kind=str(request.args.get('kind') or 'all').strip() or 'all',
                page=page,
                limit=max(1, min(limit, 5000)),
            )
        except Exception as exc:
            return jsonify({'ok': False, 'error': str(exc)}), 400
        return jsonify({
            'ok': True,
            'run_id': run_id,
            'project_name': (payload.get('report', {}) or {}).get('project_name', ''),
            'cadence_index': cadence_index,
        })

    @app.get('/api/report/<run_id>/export')
    def export_report(run_id: str):
        payload = get_run(run_id)
        fd, target = tempfile.mkstemp(suffix='.xlsx')
        os.close(fd)
        os.unlink(target)
        actual = export_to_excel(payload['bundle'], target)
        try:
            with open(actual, 'rb') as handle:
                data = handle.read()
        finally:
            try:
                os.remove(actual)
            except OSError:
                pass
        return send_file(
            io.BytesIO(data),
            as_attachment=True,
            download_name=Path(actual).name,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        )

    @app.get('/api/report/<run_id>/bom/export')
    def export_report_bom(run_id: str):
        payload = get_run(run_id)
        mode = str(request.args.get('mode') or 'all').strip() or 'all'
        fd, target = tempfile.mkstemp(suffix='.xlsx')
        os.close(fd)
        os.unlink(target)
        try:
            actual = export_bom_to_excel(payload['bundle'], target, mode=mode)
        except ValueError as exc:
            try:
                os.remove(target)
            except OSError:
                pass
            return jsonify({'ok': False, 'error': str(exc)}), 400
        try:
            with open(actual, 'rb') as handle:
                data = handle.read()
        finally:
            try:
                os.remove(actual)
            except OSError:
                pass
        project_name = _safe_download_stem((payload.get('report', {}) or {}).get('project_name') or run_id)
        mode_slug = _safe_download_stem(mode, "all")
        return send_file(
            io.BytesIO(data),
            as_attachment=True,
            download_name=f"{project_name}_bom_{mode_slug}.xlsx",
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        )
