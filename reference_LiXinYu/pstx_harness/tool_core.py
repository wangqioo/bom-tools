# -*- coding: utf-8 -*-
"""Core tool protocol primitives for local harness tools."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Callable, Dict, List, Optional


class HarnessToolError(ValueError):
    """Raised when a local harness tool call is invalid or unsafe."""


@dataclass(frozen=True)
class HarnessToolContext:
    report: dict
    bundle: dict
    request: object
    project_context: dict = field(default_factory=dict)


@dataclass(frozen=True)
class HarnessTool:
    name: str
    title: str
    description: str
    target: str
    handler: Callable[[HarnessToolContext, dict], dict]
    input_schema: Dict[str, object] = field(default_factory=lambda: {
        "type": "object",
        "properties": {},
        "additionalProperties": False,
    })
    readonly: bool = True
    file_access: bool = False
    mutating: bool = False
    supports_parallel: bool = False
    approval_scope: str = ""
    evidence_kind: str = ""

    def normalized_approval_scope(self) -> str:
        scope = str(self.approval_scope or "").strip()
        if scope:
            return scope
        return "read_project_file" if self.file_access else "none"

    def normalized_evidence_kind(self) -> str:
        return str(self.evidence_kind or self.target or "general").strip() or "general"


class HarnessToolRegistry:
    def __init__(self):
        self._tools: Dict[str, HarnessTool] = {}

    def register(self, tool: HarnessTool) -> "HarnessToolRegistry":
        if not tool.name or not callable(tool.handler):
            raise ValueError("Harness tool requires name and handler")
        if tool.readonly is not True:
            raise ValueError(f"Harness tool must be readonly: {tool.name}")
        if tool.mutating:
            raise ValueError(f"Harness tool cannot be mutating: {tool.name}")
        if tool.name in self._tools:
            raise ValueError(f"Duplicate harness tool: {tool.name}")
        self._tools[tool.name] = tool
        return self

    def get(self, name: str) -> HarnessTool:
        tool = self._tools.get(name)
        if tool is None:
            raise HarnessToolError(f"Unknown harness tool: {name}")
        return tool

    def run(self, name: str, context: HarnessToolContext, args: Optional[dict] = None) -> dict:
        tool = self.get(name)
        args = _validate_tool_args(tool, args or {})
        result = dict(tool.handler(context, args) or {})
        result.setdefault("id", tool.name)
        result.setdefault("title", tool.title)
        result.setdefault("target", tool.target)
        return result

    def list_tools(self) -> List[dict]:
        return [
            {
                "name": tool.name,
                "title": tool.title,
                "description": tool.description,
                "target": tool.target,
                "input_schema": tool.input_schema,
                "result_contract": {
                    "version": "v2",
                    "fields": [
                        "completeness",
                        "recommended_next_tools",
                        "detail_tool",
                        "aggregation_tool",
                        "scope_summary",
                    ],
                },
                "readonly": tool.readonly,
                "file_access": tool.file_access,
                "mutating": tool.mutating,
                "supports_parallel": tool.supports_parallel,
                "approval_scope": tool.normalized_approval_scope(),
                "evidence_kind": tool.normalized_evidence_kind(),
            }
            for tool in self._tools.values()
        ]


def _validate_tool_args(tool: HarnessTool, args: dict) -> dict:
    schema = tool.input_schema or {}
    if schema.get("type", "object") != "object":
        raise HarnessToolError(f"工具 {tool.name} 的 input_schema 只支持 object。")
    if not isinstance(args, dict):
        raise HarnessToolError(f"工具 {tool.name} 参数必须是 JSON 对象。")
    required = set(schema.get("required") or [])
    properties = dict(schema.get("properties") or {})
    missing = [name for name in required if name not in args]
    if missing:
        raise HarnessToolError(f"工具 {tool.name} 缺少参数：{', '.join(missing)}")
    if not schema.get("additionalProperties", False):
        extra = [name for name in args if name not in properties]
        if extra:
            raise HarnessToolError(f"工具 {tool.name} 不支持参数：{', '.join(extra)}")
    normalized = {}
    for key, value in args.items():
        rule = properties.get(key, {})
        normalized[key] = _validate_arg_value(tool.name, key, value, rule)
    return normalized


def _validate_arg_value(tool_name: str, key: str, value, rule: dict):
    expected = rule.get("type")
    if expected == "string":
        if not isinstance(value, str):
            raise HarnessToolError(f"工具 {tool_name} 参数 {key} 必须是字符串。")
        value = value[: int(rule.get("maxLength") or 10000)]
    elif expected == "integer":
        if isinstance(value, bool):
            raise HarnessToolError(f"工具 {tool_name} 参数 {key} 必须是整数。")
        try:
            value = int(value)
        except (TypeError, ValueError) as exc:
            raise HarnessToolError(f"工具 {tool_name} 参数 {key} 必须是整数。") from exc
        if "minimum" in rule and value < int(rule["minimum"]):
            if rule.get("coerceMinimum"):
                value = int(rule["minimum"])
            else:
                raise HarnessToolError(f"工具 {tool_name} 参数 {key} 不能小于 {rule['minimum']}。")
        if "maximum" in rule and value > int(rule["maximum"]):
            if rule.get("coerceMaximum"):
                value = int(rule["maximum"])
            else:
                raise HarnessToolError(f"工具 {tool_name} 参数 {key} 不能大于 {rule['maximum']}。")
    elif expected == "boolean":
        if not isinstance(value, bool):
            raise HarnessToolError(f"工具 {tool_name} 参数 {key} 必须是布尔值。")
    elif expected == "array":
        if not isinstance(value, list):
            raise HarnessToolError(f"工具 {tool_name} 参数 {key} 必须是数组。")
        if "minItems" in rule and len(value) < int(rule["minItems"]):
            raise HarnessToolError(f"工具 {tool_name} 参数 {key} 至少需要 {rule['minItems']} 项。")
        if "maxItems" in rule and len(value) > int(rule["maxItems"]):
            value = value[: int(rule["maxItems"])]
    elif expected == "object":
        if not isinstance(value, dict):
            raise HarnessToolError(f"工具 {tool_name} 参数 {key} 必须是对象。")
    if "enum" in rule and value not in set(rule["enum"]):
        raise HarnessToolError(f"工具 {tool_name} 参数 {key} 不在允许范围内。")
    return value


def _schema(properties: dict, required: List[str] = None, additional: bool = False) -> dict:
    return {
        "type": "object",
        "properties": properties,
        "required": required or [],
        "additionalProperties": additional,
    }
