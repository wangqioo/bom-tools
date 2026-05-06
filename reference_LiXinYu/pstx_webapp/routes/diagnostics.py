# -*- coding: utf-8 -*-
"""Diagnostics middleware and API route registration."""

from __future__ import annotations

import io
import time


def register_diagnostics_routes(
    app,
    *,
    request,
    jsonify,
    send_file,
    build_diagnostics_status,
    diagnostics_export_bytes,
    format_exception,
    new_diagnostic_request_id,
    summarize_mapping,
    tail_diagnostics,
    write_diagnostic_event,
) -> None:
    """Register request diagnostics hooks and diagnostics API endpoints."""

    def diagnostic_request_id() -> str:
        request_id = request.environ.get('pstx_diagnostic_request_id')
        if not request_id:
            request_id = request.headers.get('X-PSTX-Request-ID') or new_diagnostic_request_id()
            request.environ['pstx_diagnostic_request_id'] = request_id
        return str(request_id)

    def should_log_request() -> bool:
        return not str(request.path or '').startswith('/static/')

    @app.before_request
    def diagnostics_request_start():
        if not should_log_request():
            return None
        request.environ['pstx_diagnostic_start_time'] = time.time()
        request_id = diagnostic_request_id()
        body_summary = {}
        if request.is_json:
            body_summary = summarize_mapping(request.get_json(silent=True) or {})
        elif request.form:
            body_summary = summarize_mapping(request.form.to_dict(flat=False))
        write_diagnostic_event('web.request.start', {
            'method': request.method,
            'path': request.path,
            'endpoint': request.endpoint or '',
            'args': summarize_mapping(request.args.to_dict(flat=False)),
            'body': body_summary,
            'content_type': request.content_type or '',
            'content_length': request.content_length or 0,
            'remote_addr': request.remote_addr or '',
        }, request_id=request_id)
        return None

    @app.after_request
    def diagnostics_request_finish(response):
        if not should_log_request():
            return response
        request_id = diagnostic_request_id()
        started = request.environ.get('pstx_diagnostic_start_time') or time.time()
        details = {
            'method': request.method,
            'path': request.path,
            'endpoint': request.endpoint or '',
            'status': response.status_code,
            'elapsed_ms': round((time.time() - float(started)) * 1000, 2),
            'response_content_length': response.calculate_content_length() or 0,
        }
        if response.status_code >= 400 and str(request.path or '').startswith('/api/'):
            try:
                details['response'] = response.get_json(silent=True) or response.get_data(as_text=True)[:1200]
            except Exception:
                details['response'] = 'response preview unavailable'
        write_diagnostic_event(
            'web.request.finish',
            details,
            level='warning' if response.status_code >= 400 else 'info',
            request_id=request_id,
        )
        response.headers['X-PSTX-Request-ID'] = request_id
        return response

    @app.teardown_request
    def diagnostics_request_exception(exc):
        if exc is None or not should_log_request():
            return None
        write_diagnostic_event('web.request.exception', {
            'method': request.method,
            'path': request.path,
            'endpoint': request.endpoint or '',
            'exception': format_exception(exc),
        }, level='error', request_id=diagnostic_request_id())
        return None

    @app.get('/api/diagnostics/status')
    def diagnostics_status():
        return jsonify(build_diagnostics_status())

    @app.get('/api/diagnostics/tail')
    def diagnostics_tail():
        try:
            limit = int(request.args.get('limit') or 200)
        except ValueError:
            return jsonify({'ok': False, 'error': 'limit 必须是数字。'}), 400
        return jsonify(tail_diagnostics(limit=limit))

    @app.get('/api/diagnostics/export')
    def diagnostics_export():
        data, filename = diagnostics_export_bytes()
        return send_file(
            io.BytesIO(data),
            as_attachment=True,
            download_name=filename,
            mimetype='application/zip',
        )
