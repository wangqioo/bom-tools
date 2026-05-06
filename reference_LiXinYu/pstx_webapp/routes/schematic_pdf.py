# -*- coding: utf-8 -*-
"""Schematic PDF annotation API routes."""

from __future__ import annotations

import tempfile
from pathlib import Path
from typing import Any, Dict

from pstx_core.schematic_pdf_annotation import load_json_mapping_or_sequence, load_targets_json


BASE_DIR = Path(__file__).resolve().parents[2]
TMP_PDF_DIR = BASE_DIR / "tmp" / "pdfs"


def register_schematic_pdf_routes(
    app,
    *,
    request,
    jsonify,
    get_run,
    build_schematic_pdf_annotation_payload,
) -> None:
    """Register schematic PDF annotation routes."""

    def _json_form_value(data: Dict[str, Any], key: str, default: Any) -> Any:
        value = data.get(key)
        if value is None or value == "":
            return default
        return load_json_mapping_or_sequence(value, default=default)

    def _coerce_calibrations(value: Any) -> list:
        if isinstance(value, dict):
            value = value.get("page_calibrations", value.get("calibrations", []))
        if value in (None, ""):
            return []
        if not isinstance(value, list):
            raise ValueError("page_calibrations/calibrations 必须是数组。")
        return value

    def _truthy(value: Any) -> bool:
        if isinstance(value, bool):
            return value
        if value is None:
            return False
        return str(value).strip().lower() in {"1", "true", "yes", "y", "on"}

    def _save_uploaded_pdf() -> str:
        uploaded = request.files.get("pdf") if getattr(request, "files", None) is not None else None
        if uploaded is None or not getattr(uploaded, "filename", ""):
            return ""
        TMP_PDF_DIR.mkdir(parents=True, exist_ok=True)
        suffix = ".pdf"
        filename = str(getattr(uploaded, "filename", "") or "")
        if filename.lower().endswith(".pdf"):
            suffix = ".pdf"
        with tempfile.NamedTemporaryFile(prefix="schematic-", suffix=suffix, dir=str(TMP_PDF_DIR), delete=False) as handle:
            tmp_path = handle.name
        uploaded.save(tmp_path)
        return tmp_path

    @app.post("/api/report/<run_id>/schematic-pdf/annotations")
    def schematic_pdf_annotations(run_id: str):
        payload = get_run(run_id)
        tmp_pdf_path = ""
        try:
            json_body = request.get_json(silent=True)
            if isinstance(json_body, dict):
                data: Dict[str, Any] = dict(json_body)
            else:
                data = dict(request.form or {})
            tmp_pdf_path = _save_uploaded_pdf()
            pdf_path = tmp_pdf_path or str(data.get("pdf_path") or data.get("pdf") or "").strip()
            if not pdf_path:
                return jsonify({"ok": False, "error": "请上传 pdf 文件，或提供 pdf_path。"}), 400

            if "targets" in data and isinstance(data.get("targets"), list):
                targets = load_targets_json({"targets": data.get("targets")})
            else:
                targets_raw = data.get("targets_json") or data.get("targets") or ""
                if not targets_raw:
                    target = {
                        key: data.get(key)
                        for key in ("kind", "type", "refdes", "net", "page", "project_page", "label", "severity", "message")
                        if data.get(key) not in (None, "")
                    }
                    targets = [target] if target else []
                else:
                    targets = load_targets_json(targets_raw)
            if not targets:
                return jsonify({"ok": False, "error": "请提供 targets 或 targets_json。"}), 400

            pdf_page_map = _json_form_value(data, "pdf_page_map", {})
            if not pdf_page_map:
                pdf_page_map = _json_form_value(data, "pdf_page_map_json", {})
            if not isinstance(pdf_page_map, dict):
                return jsonify({"ok": False, "error": "pdf_page_map 必须是对象。"}), 400

            page_calibrations = _coerce_calibrations(_json_form_value(data, "page_calibrations", []))
            if not page_calibrations:
                page_calibrations = _coerce_calibrations(_json_form_value(data, "calibrations_json", []))

            try:
                limit = int(data.get("limit") or 200)
            except (TypeError, ValueError):
                return jsonify({"ok": False, "error": "limit 必须是数字。"}), 400
            stdout = str(data.get("stdout") or "annotations").strip() or "annotations"
            if stdout not in {"summary", "annotations", "full"}:
                return jsonify({"ok": False, "error": "stdout 必须是 summary、annotations 或 full。"}), 400

            annotation = build_schematic_pdf_annotation_payload(
                pdf_path,
                payload.get("bundle", {}) or {},
                targets,
                pdf_page_map=pdf_page_map,
                page_calibrations=page_calibrations,
                stdout=stdout,
                limit=max(1, min(limit, 5000)),
                allow_page_number_fallback=_truthy(data.get("allow_page_number_fallback")),
            )
            return jsonify({
                "ok": True,
                "run_id": run_id,
                "project_name": (payload.get("report", {}) or {}).get("project_name", ""),
                "schematic_pdf_annotation": annotation,
            })
        except Exception as exc:
            return jsonify({"ok": False, "error": str(exc)}), 400
        finally:
            if tmp_pdf_path:
                try:
                    Path(tmp_pdf_path).unlink(missing_ok=True)
                except OSError:
                    pass
