# -*- coding: utf-8 -*-
"""DFMEA workbench API routes."""

from __future__ import annotations

import io

from pstx_knowledge.dfmea_workbench import (
    create_dfmea_group,
    delete_dfmea_group,
    export_dfmea_workbook,
    get_dfmea_workbench,
    sync_dfmea_project,
    update_dfmea_group,
)


def _bool_arg(value) -> bool:
    return str(value or "").strip().lower() in {"1", "true", "yes", "on", "y"}


def _error(jsonify, message: str, status: int = 400):
    return jsonify({"ok": False, "error": message}), status


def register_dfmea_routes(
    app,
    *,
    request,
    jsonify,
    send_file,
    get_run,
) -> None:
    """Register report-bound DFMEA workbench routes."""

    def _get_run_payload(run_id: str):
        try:
            payload = get_run(run_id)
        except Exception:
            return None, _error(jsonify, "未找到 DFMEA 对应报告，请重新分析项目或刷新项目列表。", 404)
        if not isinstance(payload, dict):
            return None, _error(jsonify, "未找到 DFMEA 对应报告，请重新分析项目或刷新项目列表。", 404)
        return payload, None

    @app.get("/api/report/<run_id>/dfmea/workbench")
    def dfmea_workbench(run_id: str):
        payload, error_response = _get_run_payload(run_id)
        if error_response:
            return error_response
        try:
            result = get_dfmea_workbench(
                run_id,
                payload.get("report") or {},
                payload.get("bundle") or {},
                include_depop=_bool_arg(request.args.get("include_depop")),
                exclude_rc=_bool_arg(request.args.get("exclude_rc")),
                sort=str(request.args.get("sort") or "page"),
                query=str(request.args.get("q") or ""),
            )
        except Exception as exc:
            return _error(jsonify, f"DFMEA 工作台读取失败：{exc}", 500)
        return jsonify(result)

    @app.post("/api/report/<run_id>/dfmea/groups")
    def dfmea_create_group(run_id: str):
        payload, error_response = _get_run_payload(run_id)
        if error_response:
            return error_response
        data = request.get_json(silent=True) or {}
        try:
            sync_dfmea_project(run_id, payload.get("report") or {}, payload.get("bundle") or {})
            result = create_dfmea_group(run_id, data)
        except ValueError as exc:
            return _error(jsonify, str(exc), 400)
        except Exception as exc:
            return _error(jsonify, f"DFMEA 分组保存失败：{exc}", 500)
        return jsonify(result)

    @app.patch("/api/report/<run_id>/dfmea/groups/<int:group_id>")
    def dfmea_update_group(run_id: str, group_id: int):
        payload, error_response = _get_run_payload(run_id)
        if error_response:
            return error_response
        data = request.get_json(silent=True) or {}
        try:
            sync_dfmea_project(run_id, payload.get("report") or {}, payload.get("bundle") or {})
            result = update_dfmea_group(run_id, group_id, data)
        except ValueError as exc:
            return _error(jsonify, str(exc), 400)
        except Exception as exc:
            return _error(jsonify, f"DFMEA 分组更新失败：{exc}", 500)
        return jsonify(result)

    @app.delete("/api/report/<run_id>/dfmea/groups/<int:group_id>")
    def dfmea_delete_group(run_id: str, group_id: int):
        _payload, error_response = _get_run_payload(run_id)
        if error_response:
            return error_response
        try:
            result = delete_dfmea_group(run_id, group_id)
        except ValueError as exc:
            return _error(jsonify, str(exc), 404)
        except Exception as exc:
            return _error(jsonify, f"DFMEA 分组删除失败：{exc}", 500)
        return jsonify(result)

    @app.get("/api/report/<run_id>/dfmea/export")
    def dfmea_export(run_id: str):
        _payload, error_response = _get_run_payload(run_id)
        if error_response:
            return error_response
        try:
            data = export_dfmea_workbook(run_id)
        except Exception as exc:
            return _error(jsonify, f"DFMEA Excel 导出失败：{exc}", 500)
        return send_file(
            io.BytesIO(data),
            as_attachment=True,
            download_name=f"dfmea_{run_id}.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
