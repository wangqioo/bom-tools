# -*- coding: utf-8 -*-
"""Filesystem-backed Harness Skill registry.

The format is intentionally close to Claude Code style SKILL.md files, but this
is for the PSTX Harness Agent itself. It is not Codex's local skill system.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import re
from typing import Mapping, Sequence


SKILL_REGISTRY_VERSION = "pstx-harness-skills/v1"


def _repo_root(start: str | Path | None = None) -> Path:
    path = Path(start or ".").expanduser().resolve()
    if path.is_file():
        path = path.parent
    for current in (path, *path.parents):
        if (current / ".git").exists() or (current / "AGENTS.md").is_file():
            return current
    return path


def _frontmatter_and_body(text: str) -> tuple[dict, str]:
    if not text.startswith("---"):
        return {}, text
    match = re.match(r"^---\s*\n(.*?)\n---\s*\n?(.*)$", text, flags=re.DOTALL)
    if not match:
        return {}, text
    meta: dict[str, object] = {}
    current_key = ""
    for raw in match.group(1).splitlines():
        line = raw.rstrip()
        if not line.strip() or line.lstrip().startswith("#"):
            continue
        if line.startswith("  - ") and current_key:
            meta.setdefault(current_key, [])
            if isinstance(meta[current_key], list):
                meta[current_key].append(line[4:].strip().strip('"\''))
            continue
        if ":" in line:
            key, value = line.split(":", 1)
            current_key = key.strip()
            value = value.strip()
            if not value:
                meta[current_key] = []
            elif value.startswith("[") and value.endswith("]"):
                items = [item.strip().strip('"\'') for item in value[1:-1].split(",") if item.strip()]
                meta[current_key] = items
            else:
                meta[current_key] = value.strip('"\'')
    return meta, match.group(2)


def _list_value(value: object) -> tuple[str, ...]:
    if value is None:
        return ()
    if isinstance(value, (list, tuple)):
        return tuple(str(item).strip() for item in value if str(item).strip())
    text = str(value).strip()
    if not text:
        return ()
    return tuple(item.strip() for item in re.split(r"[,，]", text) if item.strip())


def _preview(text: object, limit: int = 1600) -> str:
    source = "" if text is None else str(text)
    source = source.replace("\r\n", "\n").replace("\r", "\n").strip()
    return source if len(source) <= limit else source[: max(0, limit - 1)] + "…"


@dataclass(frozen=True)
class HarnessSkill:
    id: str
    title: str
    description: str
    triggers: tuple[str, ...]
    capability_profiles: tuple[str, ...]
    playbooks: tuple[str, ...]
    allowed_tools: tuple[str, ...]
    output_rules: tuple[str, ...]
    path: str
    body: str

    def matches(self,
                text: str,
                *,
                capability_profiles: Sequence[object] = (),
                playbook_ids: Sequence[object] = (),
                tool_names: Sequence[object] = ()) -> bool:
        haystack = str(text or "").lower()
        upper = haystack.upper()
        if any(str(item) in self.capability_profiles for item in capability_profiles or []):
            return True
        if any(str(item) in self.playbooks for item in playbook_ids or []):
            return True
        if any(str(item) in self.allowed_tools for item in tool_names or []):
            return True
        return any(token.lower() in haystack or token.upper() in upper for token in self.triggers)

    def card(self, *, include_body: bool = False, max_body_chars: int = 1800) -> dict:
        payload = {
            "id": self.id,
            "title": self.title,
            "description": self.description,
            "triggers": list(self.triggers),
            "capability_profiles": list(self.capability_profiles),
            "playbooks": list(self.playbooks),
            "allowed_tools": list(self.allowed_tools),
            "output_rules": list(self.output_rules),
            "source_path": self.path,
        }
        if include_body:
            payload["body"] = _preview(self.body, max_body_chars)
            payload["body_truncated"] = len(self.body) > max_body_chars
        return payload


def load_harness_skills(root: str | Path | None = None) -> list[HarnessSkill]:
    base = _repo_root(root)
    skills_dir = base / "harness_skills"
    if not skills_dir.is_dir():
        return []
    skills: list[HarnessSkill] = []
    for skill_file in sorted(skills_dir.glob("*/SKILL.md")):
        try:
            raw = skill_file.read_text(encoding="utf-8", errors="replace")
        except OSError:
            continue
        meta, body = _frontmatter_and_body(raw)
        skill_id = str(meta.get("name") or skill_file.parent.name).strip() or skill_file.parent.name
        skills.append(HarnessSkill(
            id=skill_id,
            title=str(meta.get("title") or skill_id),
            description=str(meta.get("description") or "").strip(),
            triggers=_list_value(meta.get("triggers")),
            capability_profiles=_list_value(meta.get("capability_profiles")),
            playbooks=_list_value(meta.get("playbooks")),
            allowed_tools=_list_value(meta.get("allowed_tools")),
            output_rules=_list_value(meta.get("output_rules")),
            path=str(skill_file),
            body=body.strip(),
        ))
    return skills


def _playbook_ids(playbook_plan: Mapping[str, object] | None) -> list[str]:
    result: list[str] = []
    if not isinstance(playbook_plan, Mapping):
        return result
    for item in playbook_plan.get("selected_playbooks") or []:
        if isinstance(item, Mapping) and item.get("id"):
            result.append(str(item.get("id")))
    return result


def select_harness_skills(*,
                          question: object = "",
                          capability_profiles: Sequence[object] = (),
                          playbook_plan: Mapping[str, object] | None = None,
                          tool_result_contracts: Sequence[Mapping[str, object]] = (),
                          root: str | Path | None = None,
                          max_selected: int = 4,
                          include_body: bool = True,
                          max_body_chars: int = 1800) -> dict:
    skills = load_harness_skills(root)
    playbooks = _playbook_ids(playbook_plan)
    tools: list[str] = []
    if isinstance(playbook_plan, Mapping):
        tools.extend(str(item) for item in playbook_plan.get("recommended_first_tools") or [] if str(item))
    for contract in tool_result_contracts or []:
        if not isinstance(contract, Mapping):
            continue
        tools.extend(str(item) for item in contract.get("recommended_next_tools") or [] if str(item))
        for key in ("detail_tool", "aggregation_tool"):
            tool = contract.get(key)
            if isinstance(tool, Mapping) and tool.get("name"):
                tools.append(str(tool.get("name")))
    selected: list[HarnessSkill] = []
    for skill in skills:
        if skill.matches(
            str(question or ""),
            capability_profiles=capability_profiles,
            playbook_ids=playbooks,
            tool_names=tools,
        ):
            selected.append(skill)
        if len(selected) >= max(0, int(max_selected or 0)):
            break
    return {
        "version": SKILL_REGISTRY_VERSION,
        "available_count": len(skills),
        "selected_count": len(selected),
        "selected_skills": [
            skill.card(include_body=include_body, max_body_chars=max_body_chars)
            for skill in selected
        ],
        "skill_cards": [skill.card(include_body=False) for skill in skills[:24]],
    }
