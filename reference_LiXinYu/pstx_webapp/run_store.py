"""Run cache helpers and project list view models for the Web app."""

from __future__ import annotations

from typing import Iterable

from pstx_webapp.state import AGENT_CONTEXT_CACHE, RUN_CACHE, MAX_RUNS


def remember_run(run_id: str, payload: dict) -> None:
    RUN_CACHE[run_id] = payload
    RUN_CACHE.move_to_end(run_id)
    while len(RUN_CACHE) > MAX_RUNS:
        old_run_id, _ = RUN_CACHE.popitem(last=False)
        AGENT_CONTEXT_CACHE.pop(old_run_id, None)


def get_run(run_id: str) -> dict | None:
    return RUN_CACHE.get(run_id)


def build_project_summary(run_id: str, payload: dict, *, drc_issue_keys: Iterable[str] = ()) -> dict:
    bundle = payload.get("bundle", {})
    report = payload.get("report", {})
    drc = bundle.get("drc", {})
    metrics = report.get("metrics", [])
    metric_map = {str(item.get("label", "")): item.get("value") for item in metrics}
    return {
        "run_id": run_id,
        "project_name": report.get("project_name") or bundle.get("project_name") or "未命名项目",
        "project_root": bundle.get("project_root", ""),
        "project_input_snapshot": dict(bundle.get("project_input_snapshot", {}) or {}),
        "generated_at": report.get("generated_at") or bundle.get("generated_at", ""),
        "ratio_limit": report.get("ratio_limit", bundle.get("ratio_limit", "")),
        "include_depop": bool(report.get("include_depop", bundle.get("include_depop", False))),
        "component_count": len(bundle.get("components", {}) or {}),
        "net_count": len(bundle.get("nets", {}) or {}),
        "drc_count": sum(len(drc.get(key, [])) for key in drc_issue_keys),
        "metrics": metrics,
        "metric_map": metric_map,
    }


def list_project_summaries(*, drc_issue_keys: Iterable[str] = ()) -> list[dict]:
    return [
        build_project_summary(run_id, payload, drc_issue_keys=drc_issue_keys)
        for run_id, payload in reversed(RUN_CACHE.items())
    ]
