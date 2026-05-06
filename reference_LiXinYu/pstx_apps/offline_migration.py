# -*- coding: utf-8 -*-
"""Prepare and verify offline migration bundles for air-gapped machines."""

from __future__ import annotations

import hashlib
import json
import os
import shutil
import subprocess
import sys
import tarfile
import time
import urllib.parse
import urllib.request
import zipfile
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional


OFFLINE_MIGRATION_SCHEMA_VERSION = "pstx-offline-migration.v1"
MANIFEST_NAME = "offline_manifest.json"
DEFAULT_TARGET_PROFILE = "windows-rtx4060-cuda"
DEFAULT_MINERU_WHEEL_SPEC = "mineru[pipeline]"
DEFAULT_MINERU_MODEL_SOURCE = "huggingface"
DEFAULT_MINERU_MODEL_TYPE = "pipeline"
OFFLINE_ASSET_CACHE_VERSION = "2026-05-06.1"
DEFAULT_ASSET_CACHE_DIRNAME = "_asset_cache"

DEFAULT_EXCLUDE_NAMES = {
    ".git",
    "__pycache__",
    ".pytest_cache",
    ".playwright-cli",
    ".venv",
    ".venv-mineru",
    ".codex-smoke",
    "trash",
    "unused_code",
    "test-results",
    "output",
    "tmp",
    "logs",
    "datasheet_data",
    "agent_ref_data",
    "agent_checklist_data",
    "dfmea_data",
    "agent_memory",
    "agent_workspace",
    "feishu_bom_data",
    "vendor",
}

PYTHON_MIRRORS = {
    "official": "https://www.python.org/ftp/python",
    "tuna": "https://mirrors.tuna.tsinghua.edu.cn/python",
    "npmmirror": "https://npmmirror.com/mirrors/python",
}

REQUIREMENT_IMPORT_ALIASES = {
    "flask": "flask",
    "openpyxl": "openpyxl",
    "pycryptodome": "Crypto",
    "pypdf": "pypdf",
    "pdfplumber": "pdfplumber",
}


def _now() -> str:
    return time.strftime("%Y-%m-%dT%H:%M:%S", time.localtime())


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _safe_name(value: str) -> str:
    text = "".join(ch if ch.isalnum() or ch in {"-", "_", "."} else "-" for ch in str(value or "").strip())
    return text.strip("-") or "pstx-offline"


def _copytree(src: Path, dst: Path) -> None:
    if dst.exists():
        shutil.rmtree(dst)
    shutil.copytree(src, dst, symlinks=False)


def _copytree_merge(src: Path, dst: Path) -> None:
    dst.mkdir(parents=True, exist_ok=True)
    for item in sorted(src.iterdir(), key=lambda entry: entry.name.lower()):
        target = dst / item.name
        if item.is_dir():
            _copytree_merge(item, target)
        elif item.is_file():
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(item, target)


def _dir_total_bytes(root: Path) -> int:
    return sum(item.stat().st_size for item in root.rglob("*") if item.is_file())


def _hash_text(value: str) -> str:
    return hashlib.sha256(value.encode("utf-8")).hexdigest()


def _cache_key(payload: Dict[str, Any]) -> str:
    body = json.dumps(payload, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
    return _hash_text(body)


def _requirement_file_digest(path: Path) -> str:
    return _sha256(path) if path.is_file() else ""


def _asset_cache_root(*, output_root: Path, asset_cache_dir: str | Path = "", reuse_assets: bool = True) -> Optional[Path]:
    if not reuse_assets:
        return None
    raw = str(asset_cache_dir or "").strip()
    root = Path(raw).expanduser() if raw else output_root / DEFAULT_ASSET_CACHE_DIRNAME
    root.mkdir(parents=True, exist_ok=True)
    return root


def _cache_manifest_path(root: Path) -> Path:
    return root / "asset_manifest.json"


def _write_cache_manifest(root: Path, payload: Dict[str, Any]) -> None:
    root.mkdir(parents=True, exist_ok=True)
    data = {
        "schema_version": OFFLINE_MIGRATION_SCHEMA_VERSION,
        "cache_version": OFFLINE_ASSET_CACHE_VERSION,
        "written_at": _now(),
        **payload,
    }
    _cache_manifest_path(root).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def _load_cache_manifest(root: Path) -> Dict[str, Any]:
    path = _cache_manifest_path(root)
    if not path.is_file():
        return {}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    if data.get("schema_version") != OFFLINE_MIGRATION_SCHEMA_VERSION:
        return {}
    if data.get("cache_version") != OFFLINE_ASSET_CACHE_VERSION:
        return {}
    return data


def _copy_project_tree(project_root: Path, target: Path, *, extra_exclude_roots: Iterable[Path] = ()) -> List[str]:
    copied: List[str] = []
    target.mkdir(parents=True, exist_ok=True)
    exclude_roots = []
    for root in extra_exclude_roots:
        try:
            exclude_roots.append(root.resolve())
        except OSError:
            exclude_roots.append(root.absolute())
    for item in sorted(project_root.iterdir(), key=lambda entry: entry.name.lower()):
        if item.name in DEFAULT_EXCLUDE_NAMES or item.name.endswith(".pyc"):
            continue
        try:
            item_resolved = item.resolve()
        except OSError:
            item_resolved = item.absolute()
        if any(item_resolved == root or root in item_resolved.parents for root in exclude_roots):
            continue
        dest = target / item.name
        if item.is_dir():
            shutil.copytree(
                item,
                dest,
                ignore=shutil.ignore_patterns("__pycache__", "*.pyc", ".DS_Store"),
                symlinks=False,
            )
        elif item.is_file():
            shutil.copy2(item, dest)
        copied.append(item.name)
    return copied


def _download_file(url: str, target: Path) -> None:
    target.parent.mkdir(parents=True, exist_ok=True)
    with urllib.request.urlopen(url, timeout=120) as response, target.open("wb") as handle:
        shutil.copyfileobj(response, handle)


def _prepare_python_archive(*,
                            url: str,
                            filename: str,
                            target: Path,
                            asset_cache: Optional[Path]) -> Dict[str, Any]:
    cache_info: Dict[str, Any] = {"enabled": bool(asset_cache), "hit": False, "path": "", "key": ""}
    if not asset_cache:
        _download_file(url, target)
        return cache_info
    key = _cache_key({
        "kind": "python_archive",
        "cache_version": OFFLINE_ASSET_CACHE_VERSION,
        "url": url,
        "filename": filename,
    })
    cache_root = asset_cache / "python_archives" / key
    cached_archive = cache_root / filename
    cache_info.update({"path": str(cached_archive), "key": key})
    manifest = _load_cache_manifest(cache_root)
    if cached_archive.is_file() and manifest.get("kind") == "python_archive":
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(cached_archive, target)
        cache_info["hit"] = True
        return cache_info
    if cache_root.exists():
        shutil.rmtree(cache_root)
    _download_file(url, cached_archive)
    _write_cache_manifest(cache_root, {
        "kind": "python_archive",
        "url": url,
        "filename": filename,
        "size": cached_archive.stat().st_size,
        "sha256": _sha256(cached_archive),
    })
    target.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(cached_archive, target)
    return cache_info


def build_python_download_url(
    *,
    python_version: str,
    python_mirror: str = "official",
    python_mirror_base: str = "",
    python_filename: str = "",
) -> str:
    version = str(python_version or "").strip()
    if not version:
        raise ValueError("--python-version is required when building a Python mirror URL")
    base = str(python_mirror_base or "").strip().rstrip("/")
    if not base:
        base = PYTHON_MIRRORS.get(str(python_mirror or "official").strip().lower())
    if not base:
        raise ValueError(f"unsupported python mirror: {python_mirror}")
    filename = str(python_filename or "").strip() or f"python-{version}-embed-amd64.zip"
    return f"{base}/{version}/{filename}"


def _extract_archive(archive: Path, target: Path) -> None:
    if target.exists():
        shutil.rmtree(target)
    target.mkdir(parents=True, exist_ok=True)
    lower = archive.name.lower()
    if lower.endswith(".zip"):
        with zipfile.ZipFile(archive) as zf:
            zf.extractall(target)
        return
    if lower.endswith((".tar.gz", ".tgz", ".tar")):
        with tarfile.open(archive) as tf:
            tf.extractall(target)
        return
    raise ValueError(f"unsupported Python archive format: {archive.name}")


def _python_candidates(root: Path) -> List[Path]:
    names = [
        root / "python.exe",
        root / "python",
        root / "bin" / "python",
        root / "bin" / "python3",
        root / "Scripts" / "python.exe",
    ]
    return [path for path in names if path.exists()]


def _mineru_candidates(root: Path) -> List[Path]:
    names = [
        root / "bin" / "mineru",
        root / "Scripts" / "mineru.exe",
        root / "mineru",
        root / "mineru.exe",
    ]
    return [path for path in names if path.exists()]


def _portable_python_mineru_candidates(bundle_root: Path, manifest: dict) -> List[Path]:
    python_info = manifest.get("python", {}) or {}
    python_root = bundle_root / str(python_info.get("extracted_path") or python_info.get("path") or "runtime/python")
    candidates: List[Path] = []
    for python_bin in _python_candidates(python_root):
        candidates.extend([
            python_bin.parent / "mineru.exe",
            python_bin.parent / "mineru",
            python_bin.parent / "Scripts" / "mineru.exe",
            python_bin.parent / "Scripts" / "mineru",
            python_bin.parent.parent / "bin" / "mineru",
        ])
    return [path for path in candidates if path.exists()]


def _mineru_model_downloader_candidates(*, mineru_venv: str | Path = "", explicit: str | Path = "") -> List[Path]:
    candidates: List[Path] = []
    raw_explicit = str(explicit or "").strip()
    if raw_explicit:
        explicit_path = Path(raw_explicit).expanduser()
        if explicit_path.is_dir():
            candidates.extend([
                explicit_path / "bin" / "mineru-models-download",
                explicit_path / "Scripts" / "mineru-models-download.exe",
                explicit_path / "Scripts" / "mineru-models-download",
                explicit_path / "mineru-models-download",
                explicit_path / "mineru-models-download.exe",
            ])
        else:
            candidates.append(explicit_path)
    raw_venv = str(mineru_venv or "").strip()
    if raw_venv:
        root = Path(raw_venv).expanduser()
        candidates.extend([
            root / "bin" / "mineru-models-download",
            root / "Scripts" / "mineru-models-download.exe",
            root / "Scripts" / "mineru-models-download",
            root / "mineru-models-download",
            root / "mineru-models-download.exe",
        ])
    path_candidate = shutil.which("mineru-models-download")
    if path_candidate:
        candidates.append(Path(path_candidate))
    seen: set[str] = set()
    result: List[Path] = []
    for path in candidates:
        key = str(path)
        if key in seen:
            continue
        seen.add(key)
        if path.exists():
            result.append(path)
    return result


def _venv_python(venv_root: str | Path) -> Optional[Path]:
    root = Path(venv_root).expanduser()
    candidates = _python_candidates(root)
    return candidates[0] if candidates else None


def _mineru_model_downloader_command_candidates(*,
                                                mineru_venv: str | Path = "",
                                                explicit: str | Path = "") -> List[List[str]]:
    commands: List[List[str]] = []
    for path in _mineru_model_downloader_candidates(mineru_venv=mineru_venv, explicit=explicit):
        commands.append([str(path)])
    raw_explicit = str(explicit or "").strip()
    if raw_explicit:
        explicit_path = Path(raw_explicit).expanduser()
        if explicit_path.exists() and explicit_path.is_file() and explicit_path.name.lower().startswith("python"):
            commands.append([str(explicit_path), "-m", "mineru.cli.models_download"])
        if explicit_path.exists() and explicit_path.is_dir():
            explicit_python = _venv_python(explicit_path)
            if explicit_python:
                commands.append([str(explicit_python), "-m", "mineru.cli.models_download"])
    raw_venv = str(mineru_venv or "").strip()
    if raw_venv:
        python_bin = _venv_python(raw_venv)
        if python_bin:
            commands.append([str(python_bin), "-m", "mineru.cli.models_download"])
    seen: set[tuple[str, ...]] = set()
    result: List[List[str]] = []
    for command in commands:
        key = tuple(command)
        if key in seen:
            continue
        seen.add(key)
        result.append(command)
    return result


def _pip_index_args(*, pip_index_url: str = "", pip_extra_index_url: str = "") -> List[str]:
    args: List[str] = []
    if str(pip_index_url or "").strip():
        args.extend(["--index-url", str(pip_index_url).strip()])
    if str(pip_extra_index_url or "").strip():
        args.extend(["--extra-index-url", str(pip_extra_index_url).strip()])
    return args


def _command_tail(proc: subprocess.CompletedProcess) -> str:
    text = (proc.stderr or proc.stdout or "").strip()
    return text[-4000:]


def _bootstrap_mineru_venv(*,
                           venv_root: str | Path,
                           mineru_spec: str = DEFAULT_MINERU_WHEEL_SPEC,
                           pip_index_url: str = "",
                           pip_extra_index_url: str = "") -> Dict[str, Any]:
    target = Path(venv_root).expanduser()
    spec = str(mineru_spec or "").strip() or DEFAULT_MINERU_WHEEL_SPEC
    info: Dict[str, Any] = {
        "requested": True,
        "ok": False,
        "path": str(target),
        "created": False,
        "installed": False,
        "mineru_spec": spec,
        "commands": [],
        "stdout": "",
        "stderr": "",
    }
    python_bin = _venv_python(target)
    if python_bin is None:
        create_cmd = [sys.executable, "-m", "venv", str(target)]
        info["commands"].append(create_cmd)
        proc = subprocess.run(create_cmd, capture_output=True, text=True, check=False)
        info["stdout"] = (info["stdout"] + "\n" + (proc.stdout or "")).strip()
        info["stderr"] = (info["stderr"] + "\n" + (proc.stderr or "")).strip()
        if proc.returncode != 0:
            raise RuntimeError(f"failed to create MinerU virtualenv at {target}: {_command_tail(proc)}")
        info["created"] = True
        python_bin = _venv_python(target)
    if python_bin is None:
        raise FileNotFoundError(f"MinerU virtualenv Python not found after creation: {target}")

    install_cmd = [str(python_bin), "-m", "pip", "install", spec]
    install_cmd.extend(_pip_index_args(pip_index_url=pip_index_url, pip_extra_index_url=pip_extra_index_url))
    info["commands"].append(install_cmd)
    env = os.environ.copy()
    env["PIP_DISABLE_PIP_VERSION_CHECK"] = "1"
    proc = subprocess.run(install_cmd, capture_output=True, text=True, check=False, env=env)
    info["stdout"] = (info["stdout"] + "\n" + (proc.stdout or "")).strip()[-8000:]
    info["stderr"] = (info["stderr"] + "\n" + (proc.stderr or "")).strip()[-12000:]
    if proc.returncode != 0:
        raise RuntimeError(f"failed to install {spec!r} into MinerU virtualenv {target}: {_command_tail(proc)}")
    info["installed"] = True
    if not (_mineru_candidates(target) or _mineru_model_downloader_command_candidates(mineru_venv=target)):
        raise FileNotFoundError(f"MinerU install finished but no mineru or mineru-models-download entrypoint was found in {target}")
    info["ok"] = True
    return info


def _should_retry_mineru_downloader_without_options(proc: subprocess.CompletedProcess) -> bool:
    text = f"{proc.stderr or ''}\n{proc.stdout or ''}".lower()
    needles = [
        "no such option",
        "unrecognized arguments",
        "unknown option",
        "unexpected argument",
        "got unexpected extra argument",
    ]
    return any(needle in text for needle in needles)


def _read_mineru_model_dirs_from_config(config_path: Path, model_type: str) -> List[Path]:
    if not config_path.is_file():
        raise FileNotFoundError(f"MinerU model download did not create config: {config_path}")
    data = json.loads(config_path.read_text(encoding="utf-8"))
    models = data.get("models-dir") or data.get("models_dir") or {}
    requested = str(model_type or DEFAULT_MINERU_MODEL_TYPE).strip().lower()
    values: List[str] = []
    if isinstance(models, dict):
        if requested == "all":
            values = [str(value) for value in models.values() if str(value or "").strip()]
        else:
            value = models.get(requested)
            if value:
                values = [str(value)]
    elif isinstance(models, str) and models.strip():
        values = [models]
    dirs = [Path(value).expanduser() for value in values]
    existing = [path for path in dirs if path.is_dir()]
    if not existing:
        raise FileNotFoundError(
            f"MinerU model download finished but no model directory for {requested!r} was found in {config_path}"
        )
    return existing


def _model_dir_for_copy(model_dirs: List[Path]) -> Path:
    if len(model_dirs) == 1:
        return model_dirs[0]
    common = Path(os.path.commonpath([str(path) for path in model_dirs]))
    if common.is_dir():
        return common
    raise FileNotFoundError(f"cannot find common MinerU model directory for: {[str(path) for path in model_dirs]}")


def _download_mineru_models(*,
                            mineru_venv: str | Path = "",
                            downloader: str | Path = "",
                            config_path: Path,
                            model_source: str = DEFAULT_MINERU_MODEL_SOURCE,
                            model_type: str = DEFAULT_MINERU_MODEL_TYPE,
                            huggingface_endpoint: str = "") -> Dict[str, Any]:
    source = str(model_source or DEFAULT_MINERU_MODEL_SOURCE).strip().lower()
    if source not in {"huggingface", "modelscope"}:
        raise ValueError(f"unsupported MinerU model source: {model_source}")
    requested_model_type = str(model_type or DEFAULT_MINERU_MODEL_TYPE).strip().lower()
    if requested_model_type not in {"pipeline", "vlm", "all"}:
        raise ValueError(f"unsupported MinerU model type: {model_type}")
    commands = _mineru_model_downloader_command_candidates(mineru_venv=mineru_venv, explicit=downloader)
    if not commands:
        searched = []
        raw_venv = str(mineru_venv or "").strip()
        if raw_venv:
            root = Path(raw_venv).expanduser()
            searched.extend([
                str(root / "bin" / "mineru-models-download"),
                str(root / "Scripts" / "mineru-models-download.exe"),
                str(root / "bin" / "python -m mineru.cli.models_download"),
                str(root / "Scripts" / "python.exe -m mineru.cli.models_download"),
            ])
        if str(downloader or "").strip():
            searched.append(str(Path(downloader).expanduser()))
        searched.append("PATH: mineru-models-download")
        raise FileNotFoundError(
            "mineru-models-download not found after checking known MinerU entrypoints. "
            "offline-migration prepare can auto-create .venv-mineru when --download-mineru-models is used; "
            "otherwise provide --mineru-model-downloader or --mineru-venv with MinerU installed. "
            f"Searched: {searched}"
        )
    config_path.parent.mkdir(parents=True, exist_ok=True)
    env = os.environ.copy()
    env["MINERU_TOOLS_CONFIG_JSON"] = str(config_path)
    env["MINERU_MODEL_SOURCE"] = source
    endpoint = str(huggingface_endpoint or "").strip()
    if endpoint and source == "huggingface":
        env["HF_ENDPOINT"] = endpoint
    attempts: List[Dict[str, Any]] = []
    proc: Optional[subprocess.CompletedProcess] = None
    cmd: List[str] = []
    for base_command in commands:
        option_cmd = [*base_command, "-s", source, "-m", requested_model_type]
        proc = subprocess.run(option_cmd, capture_output=True, text=True, check=False, env=env)
        attempts.append({
            "command": option_cmd,
            "mode": "options",
            "returncode": proc.returncode,
            "stdout": (proc.stdout or "")[-2000:],
            "stderr": (proc.stderr or "")[-4000:],
        })
        if proc.returncode == 0:
            cmd = option_cmd
            break
        if _should_retry_mineru_downloader_without_options(proc):
            interactive_input = f"{source}\n{requested_model_type}\n"
            proc = subprocess.run(
                base_command,
                input=interactive_input,
                capture_output=True,
                text=True,
                check=False,
                env=env,
            )
            attempts.append({
                "command": base_command,
                "mode": "interactive",
                "returncode": proc.returncode,
                "stdout": (proc.stdout or "")[-2000:],
                "stderr": (proc.stderr or "")[-4000:],
            })
            if proc.returncode == 0:
                cmd = base_command
                break
    if proc is None:
        raise RuntimeError("MinerU model download did not run")
    info: Dict[str, Any] = {
        "requested": True,
        "ok": proc.returncode == 0,
        "command": cmd,
        "attempts": attempts,
        "returncode": proc.returncode,
        "source": source,
        "model_type": requested_model_type,
        "config_path": str(config_path),
        "huggingface_endpoint": endpoint if source == "huggingface" else "",
        "stdout": (proc.stdout or "")[-8000:],
        "stderr": (proc.stderr or "")[-12000:],
    }
    if proc.returncode != 0:
        raise RuntimeError(
            "MinerU model download failed after trying available entrypoints: "
            f"{info['stderr'] or info['stdout'] or attempts}"
        )
    model_dirs = _read_mineru_model_dirs_from_config(config_path, requested_model_type)
    model_dir = _model_dir_for_copy(model_dirs)
    info.update({
        "model_dir": str(model_dir),
        "model_dirs": [str(path) for path in model_dirs],
        "file_count": sum(1 for item in model_dir.rglob("*") if item.is_file()),
        "total_bytes": sum(item.stat().st_size for item in model_dir.rglob("*") if item.is_file()),
    })
    return info


def _mineru_model_cache_key(*,
                            model_source: str,
                            model_type: str,
                            huggingface_endpoint: str = "") -> str:
    return _cache_key({
        "kind": "mineru_models",
        "cache_version": OFFLINE_ASSET_CACHE_VERSION,
        "source": str(model_source or DEFAULT_MINERU_MODEL_SOURCE).strip().lower(),
        "model_type": str(model_type or DEFAULT_MINERU_MODEL_TYPE).strip().lower(),
        "huggingface_endpoint": str(huggingface_endpoint or "").strip(),
    })


def _load_cached_mineru_models(*,
                               asset_cache: Optional[Path],
                               model_source: str,
                               model_type: str,
                               huggingface_endpoint: str = "") -> tuple[str, str, Dict[str, Any]]:
    cache_info: Dict[str, Any] = {"enabled": bool(asset_cache), "hit": False, "path": "", "key": ""}
    if not asset_cache:
        return "", "", cache_info
    key = _mineru_model_cache_key(
        model_source=model_source,
        model_type=model_type,
        huggingface_endpoint=huggingface_endpoint,
    )
    cache_root = asset_cache / "mineru_models" / key
    model_dir = cache_root / "models"
    config_path = cache_root / "mineru.json"
    cache_info.update({"path": str(cache_root), "key": key})
    manifest = _load_cache_manifest(cache_root)
    if (
        manifest.get("kind") == "mineru_models"
        and model_dir.is_dir()
        and config_path.is_file()
        and any(path.is_file() for path in model_dir.rglob("*"))
    ):
        cache_info["hit"] = True
        return str(model_dir), str(config_path), cache_info
    return "", "", cache_info


def _store_cached_mineru_models(*,
                                asset_cache: Optional[Path],
                                model_dir: str | Path,
                                config_path: str | Path,
                                model_source: str,
                                model_type: str,
                                huggingface_endpoint: str = "") -> Dict[str, Any]:
    cache_info: Dict[str, Any] = {"enabled": bool(asset_cache), "hit": False, "path": "", "key": ""}
    if not asset_cache:
        return cache_info
    source_dir = Path(model_dir).expanduser()
    source_config = Path(config_path).expanduser()
    if not source_dir.is_dir() or not source_config.is_file():
        return cache_info
    key = _mineru_model_cache_key(
        model_source=model_source,
        model_type=model_type,
        huggingface_endpoint=huggingface_endpoint,
    )
    cache_root = asset_cache / "mineru_models" / key
    cache_info.update({"path": str(cache_root), "key": key})
    model_target = cache_root / "models"
    if cache_root.exists():
        shutil.rmtree(cache_root)
    _copytree(source_dir, model_target)
    cache_root.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source_config, cache_root / "mineru.json")
    _write_cache_manifest(cache_root, {
        "kind": "mineru_models",
        "source": str(model_source or DEFAULT_MINERU_MODEL_SOURCE).strip().lower(),
        "model_type": str(model_type or DEFAULT_MINERU_MODEL_TYPE).strip().lower(),
        "huggingface_endpoint": str(huggingface_endpoint or "").strip(),
        "file_count": sum(1 for item in model_target.rglob("*") if item.is_file()),
        "total_bytes": _dir_total_bytes(model_target),
    })
    return cache_info


def _standalone_verify_script_source() -> str:
    return r'''#!/usr/bin/env python3
"""Standalone offline verifier for PSTX migration bundles.

This file intentionally uses only the Python standard library so it can run
before the project dependency wheelhouse has been installed.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import zipfile


MANIFEST_NAME = "offline_manifest.json"
SCHEMA_VERSION = "pstx-offline-migration.v1"
REQUIREMENT_IMPORT_ALIASES = {
    "flask": "flask",
    "openpyxl": "openpyxl",
    "pycryptodome": "Crypto",
    "pypdf": "pypdf",
    "pdfplumber": "pdfplumber",
}


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def load_manifest(package_root: Path) -> dict:
    manifest_path = package_root / MANIFEST_NAME
    if not manifest_path.is_file():
        raise FileNotFoundError(f"offline manifest not found: {manifest_path}")
    data = json.loads(manifest_path.read_text(encoding="utf-8"))
    if data.get("schema_version") != SCHEMA_VERSION:
        raise ValueError(f"unsupported offline manifest schema: {data.get('schema_version')}")
    return data


def python_candidates(root: Path) -> list[Path]:
    names = [
        root / "python.exe",
        root / "python",
        root / "bin" / "python",
        root / "bin" / "python3",
        root / "Scripts" / "python.exe",
    ]
    return [path for path in names if path.exists()]


def mineru_candidates(root: Path) -> list[Path]:
    names = [
        root / "bin" / "mineru",
        root / "Scripts" / "mineru.exe",
        root / "mineru",
        root / "mineru.exe",
    ]
    return [path for path in names if path.exists()]


def python_mineru_candidates(package_root: Path, manifest: dict) -> list[Path]:
    python_info = manifest.get("python", {}) or {}
    python_root = package_root / str(python_info.get("extracted_path") or python_info.get("path") or "runtime/python")
    candidates = []
    for python_bin in python_candidates(python_root):
        candidates.extend([
            python_bin.parent / "mineru.exe",
            python_bin.parent / "mineru",
            python_bin.parent / "Scripts" / "mineru.exe",
            python_bin.parent / "Scripts" / "mineru",
            python_bin.parent.parent / "bin" / "mineru",
        ])
    return [path for path in candidates if path.exists()]


def verify_hashes(package_root: Path, manifest: dict) -> list[dict]:
    issues = []
    for item in manifest.get("files", []) or []:
        rel = str(item.get("path") or "")
        path = package_root / rel
        if not path.is_file():
            issues.append({"path": rel, "status": "missing", "message": "file missing"})
            continue
        expected_size = int(item.get("size") or 0)
        actual_size = path.stat().st_size
        if actual_size != expected_size:
            issues.append({"path": rel, "status": "size_mismatch", "message": f"expected {expected_size}, got {actual_size}"})
            continue
        if sha256(path) != str(item.get("sha256") or ""):
            issues.append({"path": rel, "status": "hash_mismatch", "message": "sha256 mismatch"})
    return issues


def requirement_names(requirements_path: Path) -> list[str]:
    if not requirements_path.is_file():
        return []
    names = []
    for raw_line in requirements_path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue
        for marker in ("==", ">=", "<=", "~=", "!=", "<", ">"):
            if marker in line:
                line = line.split(marker, 1)[0]
                break
        name = line.split("[", 1)[0].strip().lower().replace("-", "_")
        if name:
            names.append(name)
    return names


def wheelhouse_missing(package_root: Path, requirements_path: Path) -> list[str]:
    wheelhouse = package_root / "wheelhouse"
    if not requirements_path.is_file() or not wheelhouse.is_dir():
        return []
    files = [item.name.lower().replace("-", "_") for item in wheelhouse.iterdir() if item.is_file()]
    return [name for name in requirement_names(requirements_path) if not any(file.startswith(name) for file in files)]


def dependency_import_probe(python_bin: Path, project_root: Path, requirements_path: Path) -> dict:
    modules = []
    for name in requirement_names(requirements_path):
        modules.append(REQUIREMENT_IMPORT_ALIASES.get(name, name))
    modules = sorted(set(modules))
    if not modules:
        return {"ok": True, "checked": [], "missing": [], "error": ""}
    code = (
        "import importlib, json\n"
        f"mods = {modules!r}\n"
        "missing = []\n"
        "for mod in mods:\n"
        "    try:\n"
        "        importlib.import_module(mod)\n"
        "    except Exception as exc:\n"
        "        missing.append({'module': mod, 'error': str(exc)[:300]})\n"
        "print(json.dumps({'missing': missing}, ensure_ascii=False))\n"
        "raise SystemExit(1 if missing else 0)\n"
    )
    env = os.environ.copy()
    env["PYTHONPATH"] = str(project_root)
    try:
        proc = subprocess.run(
            [str(python_bin), "-c", code],
            cwd=str(project_root),
            env=env,
            capture_output=True,
            text=True,
            timeout=25,
            check=False,
        )
    except Exception as exc:
        return {"ok": False, "checked": modules, "missing": [], "error": str(exc)}
    output = (proc.stdout or "").strip().splitlines()
    try:
        parsed = json.loads(output[-1]) if output else {}
    except json.JSONDecodeError:
        parsed = {}
    missing = parsed.get("missing", []) if isinstance(parsed, dict) else []
    return {
        "ok": proc.returncode == 0 and not missing,
        "checked": modules,
        "missing": missing,
        "error": "" if proc.returncode == 0 else (proc.stderr or proc.stdout or "").strip()[-1000:],
    }


def verify(package_root: Path, *, probe_runtime: bool = True) -> dict:
    temp_dir = None
    if package_root.is_file() and package_root.suffix.lower() == ".zip":
        temp_dir = tempfile.TemporaryDirectory(prefix="pstx_offline_verify_")
        with zipfile.ZipFile(package_root) as zf:
            zf.extractall(temp_dir.name)
        children = [path for path in Path(temp_dir.name).iterdir() if path.is_dir()]
        package_root = children[0] if children else Path(temp_dir.name)
    try:
        manifest = load_manifest(package_root)
        issues = verify_hashes(package_root, manifest)
        warnings = []
        project_root = package_root / "project"
        for rel in ("pstx_cli.py", "pstx_web.py", "pstx_apps/cli.py"):
            if not (project_root / rel).is_file():
                issues.append({"path": f"project/{rel}", "status": "missing", "message": "required project entrypoint missing"})

        python_info = manifest.get("python", {}) or {}
        portable_required = bool(python_info.get("required_on_target", True))
        python_root = package_root / str(python_info.get("extracted_path") or python_info.get("path") or "runtime/python")
        pythons = python_candidates(python_root)
        python_archive = package_root / str(python_info.get("path") or "")
        if python_info.get("provided") and not pythons:
            if python_archive.is_file():
                message = "portable Python archive is present but not extracted; computer B without system Python cannot run verification"
                if portable_required:
                    issues.append({"path": str(python_info.get("path") or "runtime/python"), "status": "python_not_extracted", "message": message})
                else:
                    warnings.append(message)
            else:
                issues.append({"path": str(python_info.get("path") or "runtime/python"), "status": "missing", "message": "portable Python missing"})
        if not python_info.get("provided"):
            message = "No portable Python was included; computer B must provide a compatible Python runtime."
            if portable_required:
                issues.append({"path": "runtime/python", "status": "portable_python_required", "message": message})
            else:
                warnings.append(message)

        mineru_info = manifest.get("mineru", {}) or {}
        miners = mineru_candidates(package_root / str(mineru_info.get("path") or "runtime/mineru_venv"))
        miners.extend(python_mineru_candidates(package_root, manifest))
        if mineru_info.get("provided") and not miners:
            issues.append({"path": str(mineru_info.get("path") or "runtime/mineru_venv"), "status": "missing", "message": "MinerU executable missing"})
        if not mineru_info.get("provided"):
            warnings.append("No MinerU venv was included; default datasheet PDF extraction will need PSTX_MINERU_BIN on computer B.")
        mineru_assets = mineru_info.get("assets", {}) or {}
        model_info = mineru_assets.get("models", {}) or {}
        if model_info.get("provided"):
            model_root = package_root / str(model_info.get("path") or "runtime/mineru_models")
            if not model_root.is_dir():
                issues.append({"path": str(model_info.get("path") or "runtime/mineru_models"), "status": "missing", "message": "MinerU model directory missing"})
            elif not any(path.is_file() for path in model_root.rglob("*")):
                issues.append({"path": str(model_info.get("path") or "runtime/mineru_models"), "status": "empty", "message": "MinerU model directory is empty"})
        else:
            warnings.append("No MinerU model directory was included; offline PDF extraction may fail on computer B.")
        config_info = mineru_assets.get("config", {}) or {}
        if config_info.get("provided"):
            template = package_root / str(config_info.get("template_path") or "")
            if not template.is_file():
                issues.append({"path": str(config_info.get("template_path") or "runtime/mineru_config/mineru.template.json"), "status": "missing", "message": "MinerU config template missing"})

        missing_wheels = wheelhouse_missing(package_root, project_root / "requirements.txt")
        if missing_wheels:
            issues.append({"path": "wheelhouse", "status": "missing_wheels", "message": ",".join(missing_wheels)})

        dependency_probe = {"ok": None, "checked": [], "missing": [], "error": ""}
        if probe_runtime and pythons:
            dependency_probe = dependency_import_probe(pythons[0], project_root, project_root / "requirements.txt")
            if dependency_probe.get("ok") is False:
                issues.append({
                    "path": str(pythons[0].relative_to(package_root)),
                    "status": "runtime_import_failed",
                    "message": json.dumps(dependency_probe.get("missing") or dependency_probe.get("error"), ensure_ascii=False)[:1000],
                })

        ok = not issues
        return {
            "ok": ok,
            "schema_version": SCHEMA_VERSION,
            "package_root": str(package_root),
            "package_name": manifest.get("package_name", ""),
            "target_platform": manifest.get("target_platform", ""),
            "target_profile": manifest.get("target_profile", ""),
            "file_count": len(manifest.get("files", []) or []),
            "checked_file_count": len(manifest.get("files", []) or []),
            "issues": issues,
            "warnings": warnings,
            "python": {"provided": bool(python_info.get("provided")), "candidates": [str(path.relative_to(package_root)) for path in pythons]},
            "mineru": {"provided": bool(mineru_info.get("provided")), "candidates": [str(path.relative_to(package_root)) for path in miners]},
            "dependency_probe": dependency_probe,
            "summary": "Offline migration bundle verification passed." if ok else f"Offline migration bundle verification found {len(issues)} issue(s).",
        }
    finally:
        if temp_dir is not None:
            temp_dir.cleanup()


def main(argv=None):
    parser = argparse.ArgumentParser(description="Verify PSTX offline migration bundle integrity.")
    parser.add_argument("package_root", nargs="?", default=str(Path(__file__).resolve().parent), help="bundle root or zip path")
    parser.add_argument("--pretty", action="store_true", help="pretty-print JSON")
    parser.add_argument("--skip-runtime-probe", action="store_true", help="skip Python dependency import probe")
    args = parser.parse_args(argv)
    try:
        payload = verify(Path(args.package_root).expanduser().resolve(), probe_runtime=not args.skip_runtime_probe)
        print(json.dumps(payload, ensure_ascii=False, indent=2 if args.pretty else None))
        return 0 if payload.get("ok") else 1
    except Exception as exc:
        payload = {"ok": False, "schema_version": SCHEMA_VERSION, "error_code": "invalid_request", "error_message": str(exc)}
        print(json.dumps(payload, ensure_ascii=False, indent=2 if args.pretty else None))
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
'''


def _configure_b_script_source() -> str:
    return r'''#!/usr/bin/env python3
"""Configure a PSTX offline bundle on computer B.

This script intentionally uses only the Python standard library. It writes
bundle-local environment scripts, optionally installs the local wheelhouse, and
then runs the standalone verifier.
"""

from __future__ import annotations

import argparse
import json
import os
from pathlib import Path
import subprocess
import sys


MANIFEST_NAME = "offline_manifest.json"


def load_manifest(root: Path) -> dict:
    return json.loads((root / MANIFEST_NAME).read_text(encoding="utf-8"))


def python_candidates(root: Path, manifest: dict) -> list[Path]:
    info = manifest.get("python", {}) or {}
    python_root = root / str(info.get("extracted_path") or info.get("path") or "runtime/python")
    candidates = [
        python_root / "python.exe",
        python_root / "python",
        python_root / "bin" / "python",
        python_root / "bin" / "python3",
        python_root / "Scripts" / "python.exe",
    ]
    return [path for path in candidates if path.exists()]


def mineru_candidates(root: Path, manifest: dict) -> list[Path]:
    info = manifest.get("mineru", {}) or {}
    mineru_root = root / str(info.get("path") or "runtime/mineru_venv")
    pythons = python_candidates(root, manifest)
    python_roots = []
    for python_bin in pythons:
        python_roots.extend([python_bin.parent, python_bin.parent / "Scripts", python_bin.parent.parent / "bin"])
    candidates = [
        mineru_root / "Scripts" / "mineru.exe",
        mineru_root / "bin" / "mineru",
        mineru_root / "mineru.exe",
        mineru_root / "mineru",
    ]
    for python_root in python_roots:
        candidates.extend([
            python_root / "mineru.exe",
            python_root / "mineru",
        ])
    return [path for path in candidates if path.exists()]


def patch_mineru_config(root: Path, manifest: dict) -> str:
    mineru = manifest.get("mineru", {}) or {}
    assets = mineru.get("assets", {}) or {}
    config = assets.get("config", {}) or {}
    models = assets.get("models", {}) or {}
    template_rel = str(config.get("template_path") or "")
    generated_rel = str(config.get("generated_path") or "runtime/mineru_config/mineru.json")
    if not template_rel:
        return ""
    template = root / template_rel
    if not template.is_file():
        return ""
    model_dir = root / str(models.get("path") or "runtime/mineru_models")
    target = root / generated_rel
    target.parent.mkdir(parents=True, exist_ok=True)
    text = template.read_text(encoding="utf-8")
    replacements = {
        "__PSTX_MINERU_MODELS_DIR__": str(model_dir),
        "__PSTX_BUNDLE_ROOT__": str(root),
    }
    source_model = str(models.get("source") or "")
    if source_model:
        replacements[source_model] = str(model_dir)
        replacements[source_model.replace("\\", "/")] = str(model_dir).replace("\\", "/")
    for old, new in replacements.items():
        if old:
            text = text.replace(old, new)
    target.write_text(text, encoding="utf-8")
    return str(target)


def rel_or_abs(path: Path, root: Path) -> str:
    try:
        return str(path.relative_to(root))
    except ValueError:
        return str(path)


def build_env(root: Path, manifest: dict, mineru_config: str) -> dict[str, str]:
    pythons = python_candidates(root, manifest)
    miners = mineru_candidates(root, manifest)
    runtime = manifest.get("runtime", {}) or {}
    mineru_runtime = runtime.get("mineru", {}) or {}
    env = {
        "PYTHONPATH": str(root / "project"),
        "PSTX_PDF_EXTRACTOR": "mineru",
        "PSTX_DATASHEET_DATA_DIR": str(root / "data" / "datasheet_data"),
        "PSTX_MINERU_DEVICE": str(mineru_runtime.get("device") or "cuda"),
        "PSTX_MINERU_MODEL_SOURCE": str(mineru_runtime.get("model_source") or "local"),
        "MINERU_DEVICE_MODE": str(mineru_runtime.get("device") or "cuda"),
        "MINERU_MODEL_SOURCE": str(mineru_runtime.get("model_source") or "local"),
    }
    if pythons:
        env["PSTX_PORTABLE_PYTHON"] = str(pythons[0])
    if miners:
        env["PSTX_MINERU_BIN"] = str(miners[0])
    if mineru_config:
        env["MINERU_TOOLS_CONFIG_JSON"] = mineru_config
        env["PSTX_MINERU_CONFIG"] = mineru_config
    return env


def write_env_scripts(root: Path, env: dict[str, str]) -> None:
    bat_lines = ["@echo off", f"set PSTX_BUNDLE_ROOT={root}"]
    for key, value in env.items():
        bat_lines.append(f"set {key}={value}")
    (root / "RUN_ENV_B.bat").write_text("\r\n".join(bat_lines) + "\r\n", encoding="utf-8")

    ps_lines = [f"$env:PSTX_BUNDLE_ROOT = {str(root)!r}"]
    for key, value in env.items():
        ps_lines.append(f"$env:{key} = {value!r}")
    (root / "RUN_ENV_B.ps1").write_text("\n".join(ps_lines) + "\n", encoding="utf-8")

    py = env.get("PSTX_PORTABLE_PYTHON", "python")
    start_bat = [
        "@echo off",
        "call \"%~dp0RUN_ENV_B.bat\"",
        f"\"{py}\" \"%~dp0project\\pstx_web.py\"",
    ]
    (root / "START_WEB_B.bat").write_text("\r\n".join(start_bat) + "\r\n", encoding="utf-8")

    start_ps = [
        ". \"$PSScriptRoot\\RUN_ENV_B.ps1\"",
        f"& {py!r} \"$PSScriptRoot\\project\\pstx_web.py\"",
    ]
    (root / "START_WEB_B.ps1").write_text("\n".join(start_ps) + "\n", encoding="utf-8")


def install_wheelhouse(root: Path, python_bin: Path) -> dict:
    wheelhouse = root / "wheelhouse"
    requirements = root / "project" / "requirements.txt"
    if not wheelhouse.is_dir():
        return {"ok": True, "skipped": True, "reason": "wheelhouse not included"}
    manifest = load_manifest(root)

    def _run(cmd: list[str]) -> dict:
        proc = subprocess.run(cmd, cwd=str(root / "project"), capture_output=True, text=True, check=False)
        return {
            "ok": proc.returncode == 0,
            "command": cmd,
            "returncode": proc.returncode,
            "stdout": (proc.stdout or "")[-4000:],
            "stderr": (proc.stderr or "")[-4000:],
        }

    project_cmd = [
        str(python_bin),
        "-m",
        "pip",
        "install",
        "--no-index",
        "--find-links",
        str(wheelhouse),
        "-r",
        str(requirements),
    ]
    project = _run(project_cmd)
    mineru = {"ok": None, "skipped": True}
    wheelhouse_info = manifest.get("wheelhouse", {}) or {}
    mineru_info = manifest.get("mineru", {}) or {}
    mineru_wheel_info = wheelhouse_info.get("mineru", {}) or {}
    mineru_spec = str(mineru_wheel_info.get("spec") or "").strip()
    should_install_mineru = (
        not bool(mineru_info.get("provided"))
        and bool(mineru_wheel_info.get("requested"))
        and bool(mineru_wheel_info.get("ok"))
        and bool(mineru_spec)
    )
    if project.get("ok") and should_install_mineru:
        mineru_cmd = [
            str(python_bin),
            "-m",
            "pip",
            "install",
            "--no-index",
            "--find-links",
            str(wheelhouse),
            mineru_spec,
        ]
        mineru = _run(mineru_cmd)
    return {
        "ok": bool(project.get("ok")) and (mineru.get("ok") is not False),
        "skipped": False,
        "command": project_cmd,
        "project": project,
        "mineru": mineru,
        "returncode": mineru.get("returncode") if mineru.get("ok") is False else project.get("returncode"),
        "stdout": ((project.get("stdout") or "") + "\n" + (mineru.get("stdout") or ""))[-4000:],
        "stderr": ((project.get("stderr") or "") + "\n" + (mineru.get("stderr") or ""))[-4000:],
    }


def run_verify(root: Path, python_bin: Path) -> dict:
    cmd = [str(python_bin), str(root / "VERIFY_OFFLINE_B.py"), str(root), "--pretty"]
    proc = subprocess.run(cmd, cwd=str(root), capture_output=True, text=True, check=False)
    try:
        parsed = json.loads(proc.stdout)
    except Exception:
        parsed = {}
    return {
        "ok": proc.returncode == 0,
        "returncode": proc.returncode,
        "payload": parsed,
        "stdout": (proc.stdout or "")[-4000:],
        "stderr": (proc.stderr or "")[-4000:],
    }


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Configure PSTX offline bundle on computer B.")
    parser.add_argument("package_root", nargs="?", default=str(Path(__file__).resolve().parent))
    parser.add_argument("--write-env", action="store_true", help="write RUN_ENV_B and START_WEB_B scripts")
    parser.add_argument("--install-wheels", action="store_true", help="install project dependencies from local wheelhouse")
    parser.add_argument("--verify", action="store_true", help="run VERIFY_OFFLINE_B.py after setup")
    parser.add_argument("--pretty", action="store_true", help="pretty-print JSON")
    args = parser.parse_args(argv)
    root = Path(args.package_root).expanduser().resolve()
    manifest = load_manifest(root)
    pythons = python_candidates(root, manifest)
    python_bin = pythons[0] if pythons else Path(sys.executable)
    install_result = {"ok": None, "skipped": True}
    if args.install_wheels:
        install_result = install_wheelhouse(root, python_bin)
    mineru_config = patch_mineru_config(root, manifest)
    env = build_env(root, manifest, mineru_config)
    if args.write_env:
        write_env_scripts(root, env)
    verify_result = {"ok": None, "skipped": True}
    if args.verify:
        verify_result = run_verify(root, python_bin)
    ok = (install_result.get("ok") is not False) and (verify_result.get("ok") is not False)
    payload = {
        "ok": ok,
        "schema_version": manifest.get("schema_version"),
        "package_root": str(root),
        "target_profile": manifest.get("target_profile", ""),
        "python": [rel_or_abs(path, root) for path in pythons],
        "mineru": [rel_or_abs(path, root) for path in mineru_candidates(root, manifest)],
        "mineru_config": mineru_config,
        "environment": env,
        "install_wheels": install_result,
        "verification": verify_result,
        "next_command": "START_WEB_B.bat",
    }
    print(json.dumps(payload, ensure_ascii=False, indent=2 if args.pretty else None))
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
'''


def _write_verify_scripts(bundle_root: Path) -> None:
    verify_script = bundle_root / "VERIFY_OFFLINE_B.py"
    verify_script.write_text(_standalone_verify_script_source(), encoding="utf-8")
    try:
        os.chmod(verify_script, 0o755)
    except OSError:
        pass
    (bundle_root / "RUN_VERIFY_B.sh").write_text(
        "#!/usr/bin/env sh\n"
        "set -eu\n"
        "ROOT=$(CDPATH= cd -- \"$(dirname -- \"$0\")\" && pwd)\n"
        "PYTHON_BIN=${PYTHON_BIN:-python3}\n"
        "if [ -x \"$ROOT/runtime/python/bin/python\" ]; then PYTHON_BIN=\"$ROOT/runtime/python/bin/python\"; fi\n"
        "if [ -x \"$ROOT/runtime/python/python\" ]; then PYTHON_BIN=\"$ROOT/runtime/python/python\"; fi\n"
        "if [ -x \"$ROOT/runtime/python/python.exe\" ]; then PYTHON_BIN=\"$ROOT/runtime/python/python.exe\"; fi\n"
        "if ! command -v \"$PYTHON_BIN\" >/dev/null 2>&1 && [ ! -x \"$PYTHON_BIN\" ]; then\n"
        "  echo \"No Python runtime found. Rebuild the package with --python-version, --python-url, --python-archive, or --python-dir.\" >&2\n"
        "  exit 2\n"
        "fi\n"
        "exec \"$PYTHON_BIN\" \"$ROOT/VERIFY_OFFLINE_B.py\" \"$ROOT\" --pretty\n",
        encoding="utf-8",
    )
    try:
        os.chmod(bundle_root / "RUN_VERIFY_B.sh", 0o755)
    except OSError:
        pass
    (bundle_root / "RUN_VERIFY_B.bat").write_text(
        "@echo off\r\n"
        "set ROOT=%~dp0\r\n"
        "set PYTHON_BIN=python\r\n"
        "if exist \"%ROOT%runtime\\python\\python.exe\" set PYTHON_BIN=%ROOT%runtime\\python\\python.exe\r\n"
        "where \"%PYTHON_BIN%\" >nul 2>nul\r\n"
        "if errorlevel 1 if not exist \"%PYTHON_BIN%\" (\r\n"
        "  echo No Python runtime found. Rebuild the package with --python-version, --python-url, --python-archive, or --python-dir.\r\n"
        "  exit /b 2\r\n"
        ")\r\n"
        "\"%PYTHON_BIN%\" \"%ROOT%VERIFY_OFFLINE_B.py\" \"%ROOT%\" --pretty\r\n",
        encoding="utf-8",
    )
    (bundle_root / "RUN_VERIFY_B.ps1").write_text(
        "$Root = Split-Path -Parent $MyInvocation.MyCommand.Path\n"
        "$PythonBin = 'python'\n"
        "if (Test-Path (Join-Path $Root 'runtime\\python\\python.exe')) { $PythonBin = Join-Path $Root 'runtime\\python\\python.exe' }\n"
        "& $PythonBin (Join-Path $Root 'VERIFY_OFFLINE_B.py') $Root --pretty\n"
        "exit $LASTEXITCODE\n",
        encoding="utf-8",
    )
    (bundle_root / "RUN_INSTALL_WHEELHOUSE_B.sh").write_text(
        "#!/usr/bin/env sh\n"
        "set -eu\n"
        "ROOT=$(CDPATH= cd -- \"$(dirname -- \"$0\")\" && pwd)\n"
        "PYTHON_BIN=${PYTHON_BIN:-python3}\n"
        "if [ -x \"$ROOT/runtime/python/bin/python\" ]; then PYTHON_BIN=\"$ROOT/runtime/python/bin/python\"; fi\n"
        "if [ -x \"$ROOT/runtime/python/python\" ]; then PYTHON_BIN=\"$ROOT/runtime/python/python\"; fi\n"
        "if [ -x \"$ROOT/runtime/python/python.exe\" ]; then PYTHON_BIN=\"$ROOT/runtime/python/python.exe\"; fi\n"
        "if ! command -v \"$PYTHON_BIN\" >/dev/null 2>&1 && [ ! -x \"$PYTHON_BIN\" ]; then\n"
        "  echo \"No Python runtime found. Rebuild the package with --python-version, --python-url, --python-archive, or --python-dir.\" >&2\n"
        "  exit 2\n"
        "fi\n"
        "exec \"$PYTHON_BIN\" \"$ROOT/CONFIGURE_B.py\" \"$ROOT\" --install-wheels --write-env --pretty\n",
        encoding="utf-8",
    )
    try:
        os.chmod(bundle_root / "RUN_INSTALL_WHEELHOUSE_B.sh", 0o755)
    except OSError:
        pass
    (bundle_root / "RUN_INSTALL_WHEELHOUSE_B.bat").write_text(
        "@echo off\r\n"
        "set ROOT=%~dp0\r\n"
        "set PYTHON_BIN=python\r\n"
        "if exist \"%ROOT%runtime\\python\\python.exe\" set PYTHON_BIN=%ROOT%runtime\\python\\python.exe\r\n"
        "where \"%PYTHON_BIN%\" >nul 2>nul\r\n"
        "if errorlevel 1 if not exist \"%PYTHON_BIN%\" (\r\n"
        "  echo No Python runtime found. Rebuild the package with --python-version, --python-url, --python-archive, or --python-dir.\r\n"
        "  exit /b 2\r\n"
        ")\r\n"
        "\"%PYTHON_BIN%\" \"%ROOT%CONFIGURE_B.py\" \"%ROOT%\" --install-wheels --write-env --pretty\r\n",
        encoding="utf-8",
    )
    (bundle_root / "RUN_INSTALL_WHEELHOUSE_B.ps1").write_text(
        "$Root = Split-Path -Parent $MyInvocation.MyCommand.Path\n"
        "$PythonBin = 'python'\n"
        "if (Test-Path (Join-Path $Root 'runtime\\python\\python.exe')) { $PythonBin = Join-Path $Root 'runtime\\python\\python.exe' }\n"
        "& $PythonBin (Join-Path $Root 'CONFIGURE_B.py') $Root --install-wheels --write-env --pretty\n"
        "exit $LASTEXITCODE\n",
        encoding="utf-8",
    )
    configure_script = bundle_root / "CONFIGURE_B.py"
    configure_script.write_text(_configure_b_script_source(), encoding="utf-8")
    (bundle_root / "RUN_SETUP_B.bat").write_text(
        "@echo off\r\n"
        "set ROOT=%~dp0\r\n"
        "set PYTHON_BIN=python\r\n"
        "if exist \"%ROOT%runtime\\python\\python.exe\" set PYTHON_BIN=%ROOT%runtime\\python\\python.exe\r\n"
        "\"%PYTHON_BIN%\" \"%ROOT%CONFIGURE_B.py\" \"%ROOT%\" --write-env --install-wheels --verify --pretty\r\n"
        "if errorlevel 1 exit /b %errorlevel%\r\n"
        "echo.\r\n"
        "echo Setup complete. Start the app with START_WEB_B.bat\r\n",
        encoding="utf-8",
    )
    (bundle_root / "RUN_SETUP_B.ps1").write_text(
        "$Root = Split-Path -Parent $MyInvocation.MyCommand.Path\n"
        "$PythonBin = 'python'\n"
        "if (Test-Path (Join-Path $Root 'runtime\\python\\python.exe')) { $PythonBin = Join-Path $Root 'runtime\\python\\python.exe' }\n"
        "& $PythonBin (Join-Path $Root 'CONFIGURE_B.py') $Root --write-env --install-wheels --verify --pretty\n"
        "if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }\n"
        "Write-Host 'Setup complete. Start the app with START_WEB_B.ps1 or START_WEB_B.bat'\n",
        encoding="utf-8",
    )


def _write_offline_readme(bundle_root: Path, *, package_name: str) -> None:
    (bundle_root / "OFFLINE_README.md").write_text(
        f"# {package_name} Offline Migration Bundle\n\n"
        "## Computer B verification\n\n"
        "After extracting this folder on the offline machine, run the one-click setup first:\n\n"
        "```bat\n"
        "RUN_SETUP_B.bat\n"
        "```\n\n"
        "or:\n\n"
        "```powershell\n"
        "powershell -ExecutionPolicy Bypass -File RUN_SETUP_B.ps1\n"
        "```\n\n"
        "The setup script writes bundle-local environment scripts, patches the MinerU config to the local bundle path, installs `wheelhouse/` into the portable Python when possible, and runs verification.\n\n"
        "For verification only, run one of:\n\n"
        "```bash\n"
        "./RUN_VERIFY_B.sh\n"
        "```\n\n"
        "```bat\n"
        "RUN_VERIFY_B.bat\n"
        "```\n\n"
        "The verifier is offline-only. It checks the manifest, SHA256 hashes, portable Python, MinerU runtime, MinerU models/config, wheelhouse, and key project entrypoints.\n\n"
        "`RUN_VERIFY_B.*` uses `VERIFY_OFFLINE_B.py`, a standard-library-only verifier, so integrity checks can run before the full project dependencies are installed.\n\n"
        "If verification reports missing Python imports but `wheelhouse/` is present, run `RUN_INSTALL_WHEELHOUSE_B.*` once on computer B and then run verification again. If pip is unavailable, prepare the bundle on computer A with `--python-dir` pointing to a tested portable Python environment that already contains pip/site-packages.\n\n"
        "After setup, use `START_WEB_B.bat` / `START_WEB_B.ps1` to start the app with the generated environment. Recommended environment when starting manually:\n\n"
        "```text\n"
        "PSTX_PDF_EXTRACTOR=mineru\n"
        "PSTX_MINERU_BIN=<bundle>/runtime/mineru_venv/bin/mineru or Scripts\\\\mineru.exe\n"
        "PSTX_MINERU_DEVICE=cuda\n"
        "PSTX_MINERU_MODEL_SOURCE=local\n"
        "MINERU_MODEL_SOURCE=local\n"
        "MINERU_TOOLS_CONFIG_JSON=<bundle>/runtime/mineru_config/mineru.json\n"
        "PSTX_DATASHEET_DATA_DIR=<bundle>/data/datasheet_data\n"
        "```\n",
        encoding="utf-8",
    )


def _collect_manifest_files(bundle_root: Path) -> List[dict]:
    files: List[dict] = []
    for path in sorted(bundle_root.rglob("*"), key=lambda item: item.as_posix().lower()):
        if not path.is_file() or path.name == MANIFEST_NAME:
            continue
        rel = path.relative_to(bundle_root).as_posix()
        files.append({
            "path": rel,
            "size": path.stat().st_size,
            "sha256": _sha256(path),
        })
    return files


def _copy_optional_dir(source: str | Path, target: Path) -> Optional[str]:
    raw = str(source or "").strip()
    if not raw:
        return None
    src = Path(raw).expanduser()
    if not src.exists():
        raise FileNotFoundError(f"offline migration source does not exist: {src}")
    if src.is_dir():
        _copytree(src, target)
    else:
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, target)
    return str(src)


def _copy_mineru_assets(*,
                        model_dir: str | Path = "",
                        mineru_config: str | Path = "",
                        runtime_dir: Path,
                        bundle_root: Path) -> Dict[str, Any]:
    assets: Dict[str, Any] = {
        "models": {"provided": False, "source": "", "path": "", "file_count": 0, "total_bytes": 0},
        "config": {"provided": False, "source": "", "template_path": "", "generated_path": ""},
    }
    raw_model_dir = str(model_dir or "").strip()
    if raw_model_dir:
        source = Path(raw_model_dir).expanduser()
        if not source.is_dir():
            raise FileNotFoundError(f"MinerU model directory not found: {source}")
        target = runtime_dir / "mineru_models"
        _copytree(source, target)
        assets["models"].update({
            "provided": True,
            "source": str(source),
            "path": target.relative_to(bundle_root).as_posix(),
            "file_count": sum(1 for item in target.rglob("*") if item.is_file()),
            "total_bytes": _dir_total_bytes(target),
        })

    raw_config = str(mineru_config or "").strip()
    if raw_config:
        source = Path(raw_config).expanduser()
        if not source.is_file():
            raise FileNotFoundError(f"MinerU config file not found: {source}")
        target_dir = runtime_dir / "mineru_config"
        target_dir.mkdir(parents=True, exist_ok=True)
        template = target_dir / "mineru.template.json"
        text = source.read_text(encoding="utf-8")
        if raw_model_dir:
            source_model = str(Path(raw_model_dir).expanduser())
            text = text.replace(source_model, "__PSTX_MINERU_MODELS_DIR__")
            text = text.replace(source_model.replace("\\", "/"), "__PSTX_MINERU_MODELS_DIR__")
        template.write_text(text, encoding="utf-8")
        assets["config"].update({
            "provided": True,
            "source": str(source),
            "template_path": template.relative_to(bundle_root).as_posix(),
            "generated_path": (target_dir / "mineru.json").relative_to(bundle_root).as_posix(),
        })
    return assets


def _download_wheelhouse(
    *,
    wheelhouse: Path,
    requirements: Path,
    pip_index_url: str = "",
    pip_extra_index_url: str = "",
    include_mineru: bool = False,
    strict_mineru: bool = False,
    mineru_spec: str = DEFAULT_MINERU_WHEEL_SPEC,
) -> dict:
    wheelhouse.mkdir(parents=True, exist_ok=True)
    normalized_mineru_spec = str(mineru_spec or "").strip() or DEFAULT_MINERU_WHEEL_SPEC

    def _with_indexes(cmd: List[str]) -> List[str]:
        result = list(cmd)
        if pip_index_url:
            result.extend(["-i", pip_index_url])
        if pip_extra_index_url:
            result.extend(["--extra-index-url", pip_extra_index_url])
        return result

    def _run(cmd: List[str]) -> dict:
        proc = subprocess.run(cmd, capture_output=True, text=True, check=False)
        return {
            "command": cmd,
            "returncode": proc.returncode,
            "stdout": proc.stdout[-8000:],
            "stderr": proc.stderr[-12000:],
            "ok": proc.returncode == 0,
        }

    def _local_link_args() -> List[str]:
        if any(item.is_file() for item in wheelhouse.iterdir()):
            return ["--find-links", str(wheelhouse)]
        return []

    project_cmd = _with_indexes([
        sys.executable,
        "-m",
        "pip",
        "download",
        "-r",
        str(requirements),
        "-d",
        str(wheelhouse),
        *_local_link_args(),
    ])
    project = _run(project_cmd)
    info = {
        "command": project_cmd,
        "commands": [project_cmd],
        "project": project,
        "mineru": {
            "requested": bool(include_mineru),
            "strict": bool(strict_mineru),
            "spec": normalized_mineru_spec,
            "ok": None,
            "skipped": not bool(include_mineru),
        },
        "returncode": project["returncode"],
        "stdout": project["stdout"],
        "stderr": project["stderr"],
        "warnings": [],
        "ok": bool(project["ok"]),
        "partial": False,
        "failure_phase": "" if project["ok"] else "project_requirements",
    }
    if not project["ok"]:
        return info

    if include_mineru:
        mineru_cmd = _with_indexes([
            sys.executable,
            "-m",
            "pip",
            "download",
            normalized_mineru_spec,
            "-d",
            str(wheelhouse),
            *_local_link_args(),
        ])
        mineru = _run(mineru_cmd)
        info["commands"].append(mineru_cmd)
        info["mineru"].update(mineru)
        info["returncode"] = mineru["returncode"] if strict_mineru and not mineru["ok"] else project["returncode"]
        info["stdout"] = (project["stdout"] + "\n" + mineru["stdout"])[-8000:]
        info["stderr"] = (project["stderr"] + "\n" + mineru["stderr"])[-12000:]
        if not mineru["ok"]:
            warning = (
                f"{normalized_mineru_spec} wheel download failed; continuing because a copied MinerU venv can provide "
                "the runtime on computer B. Use --strict-mineru-wheels to make this fatal."
            )
            info["warnings"].append(warning)
            info["partial"] = True
            info["ok"] = not strict_mineru
            info["failure_phase"] = "mineru_runtime_wheels"
        else:
            info["ok"] = True
            info["failure_phase"] = ""
    return {
        **info,
    }


def _wheelhouse_cache_key(*,
                          requirements: Path,
                          pip_index_url: str = "",
                          pip_extra_index_url: str = "",
                          include_mineru: bool = False,
                          mineru_spec: str = DEFAULT_MINERU_WHEEL_SPEC) -> str:
    return _cache_key({
        "kind": "wheelhouse",
        "cache_version": OFFLINE_ASSET_CACHE_VERSION,
        "requirements_sha256": _requirement_file_digest(requirements),
        "requirements_name": requirements.name,
        "pip_index_url": str(pip_index_url or "").strip(),
        "pip_extra_index_url": str(pip_extra_index_url or "").strip(),
        "include_mineru": bool(include_mineru),
        "mineru_spec": str(mineru_spec or DEFAULT_MINERU_WHEEL_SPEC).strip() or DEFAULT_MINERU_WHEEL_SPEC,
        "python_version": ".".join(str(part) for part in sys.version_info[:3]),
        "platform": sys.platform,
    })


def _wheelhouse_pool_key(*,
                         pip_index_url: str = "",
                         pip_extra_index_url: str = "") -> str:
    return _cache_key({
        "kind": "wheelhouse_pool",
        "cache_version": OFFLINE_ASSET_CACHE_VERSION,
        "pip_index_url": str(pip_index_url or "").strip(),
        "pip_extra_index_url": str(pip_extra_index_url or "").strip(),
        "python_version": ".".join(str(part) for part in sys.version_info[:3]),
        "platform": sys.platform,
    })


def _seed_wheelhouse_from_pool(*,
                               asset_cache: Optional[Path],
                               pool_key: str,
                               wheelhouse: Path) -> tuple[str, int]:
    if not asset_cache or not pool_key:
        return "", 0
    pool_files = asset_cache / "wheelhouse_pool" / pool_key / "files"
    if not pool_files.is_dir():
        return str(pool_files), 0
    _copytree_merge(pool_files, wheelhouse)
    return str(pool_files), sum(1 for item in wheelhouse.iterdir() if item.is_file())


def _store_wheelhouse_pool(*,
                           asset_cache: Optional[Path],
                           pool_key: str,
                           wheelhouse: Path,
                           pip_index_url: str = "",
                           pip_extra_index_url: str = "") -> None:
    if not asset_cache or not pool_key or not wheelhouse.is_dir():
        return
    pool_root = asset_cache / "wheelhouse_pool" / pool_key
    pool_files = pool_root / "files"
    _copytree_merge(wheelhouse, pool_files)
    _write_cache_manifest(pool_root, {
        "kind": "wheelhouse_pool",
        "pip_index_url": str(pip_index_url or "").strip(),
        "pip_extra_index_url": str(pip_extra_index_url or "").strip(),
        "file_count": sum(1 for item in pool_files.iterdir() if item.is_file()),
        "total_bytes": _dir_total_bytes(pool_files),
    })


def _prepare_wheelhouse_cached(*,
                               bundle_wheelhouse: Path,
                               asset_cache: Optional[Path],
                               requirements: Path,
                               pip_index_url: str = "",
                               pip_extra_index_url: str = "",
                               include_mineru: bool = False,
                               strict_mineru: bool = False,
                               mineru_spec: str = DEFAULT_MINERU_WHEEL_SPEC) -> dict:
    if not asset_cache:
        info = _download_wheelhouse(
            wheelhouse=bundle_wheelhouse,
            requirements=requirements,
            pip_index_url=pip_index_url,
            pip_extra_index_url=pip_extra_index_url,
            include_mineru=include_mineru,
            strict_mineru=strict_mineru,
            mineru_spec=mineru_spec,
        )
        info["cache"] = {"enabled": False, "hit": False, "path": "", "key": ""}
        return info
    key = _wheelhouse_cache_key(
        requirements=requirements,
        pip_index_url=pip_index_url,
        pip_extra_index_url=pip_extra_index_url,
        include_mineru=include_mineru,
        mineru_spec=mineru_spec,
    )
    cache_root = asset_cache / "wheelhouse" / key
    cached_files = cache_root / "files"
    pool_key = _wheelhouse_pool_key(
        pip_index_url=pip_index_url,
        pip_extra_index_url=pip_extra_index_url,
    )
    cache_info = {
        "enabled": True,
        "hit": False,
        "path": str(cache_root),
        "key": key,
        "pool_key": pool_key,
        "pool_path": str(asset_cache / "wheelhouse_pool" / pool_key / "files"),
        "seeded_file_count": 0,
    }
    manifest = _load_cache_manifest(cache_root)
    if manifest.get("kind") == "wheelhouse" and cached_files.is_dir() and any(path.is_file() for path in cached_files.iterdir()):
        _copytree(cached_files, bundle_wheelhouse)
        info = manifest.get("wheelhouse", {}) if isinstance(manifest.get("wheelhouse"), dict) else {}
        info = dict(info)
        info.update({"downloaded": True, "ok": True, "cache": {**cache_info, "hit": True}})
        return info
    if cache_root.exists():
        shutil.rmtree(cache_root)
    cached_files.mkdir(parents=True, exist_ok=True)
    _, seeded_count = _seed_wheelhouse_from_pool(
        asset_cache=asset_cache,
        pool_key=pool_key,
        wheelhouse=cached_files,
    )
    cache_info["seeded_file_count"] = seeded_count
    info = _download_wheelhouse(
        wheelhouse=cached_files,
        requirements=requirements,
        pip_index_url=pip_index_url,
        pip_extra_index_url=pip_extra_index_url,
        include_mineru=include_mineru,
        strict_mineru=strict_mineru,
        mineru_spec=mineru_spec,
    )
    info["cache"] = cache_info
    if info.get("ok"):
        _store_wheelhouse_pool(
            asset_cache=asset_cache,
            pool_key=pool_key,
            wheelhouse=cached_files,
            pip_index_url=pip_index_url,
            pip_extra_index_url=pip_extra_index_url,
        )
        _write_cache_manifest(cache_root, {
            "kind": "wheelhouse",
            "requirements": str(requirements),
            "requirements_sha256": _requirement_file_digest(requirements),
            "wheelhouse": info,
            "file_count": sum(1 for item in cached_files.iterdir() if item.is_file()),
            "total_bytes": _dir_total_bytes(cached_files),
        })
        _copytree(cached_files, bundle_wheelhouse)
    return info


def prepare_offline_bundle(
    *,
    project_root: str | Path,
    out_dir: str | Path,
    name: str = "pstx-offline",
    target_platform: str = "windows-amd64",
    target_profile: str = DEFAULT_TARGET_PROFILE,
    python_dir: str = "",
    python_archive: str = "",
    python_url: str = "",
    python_version: str = "",
    python_mirror: str = "official",
    python_mirror_base: str = "",
    python_filename: str = "",
    extract_python: bool = True,
    allow_system_python_on_b: bool = False,
    mineru_venv: str = "",
    mineru_model_dir: str = "",
    mineru_config: str = "",
    download_mineru_models: bool = False,
    mineru_model_source: str = DEFAULT_MINERU_MODEL_SOURCE,
    mineru_model_type: str = DEFAULT_MINERU_MODEL_TYPE,
    huggingface_endpoint: str = "",
    mineru_model_downloader: str = "",
    download_wheels: bool = False,
    pip_index_url: str = "",
    pip_extra_index_url: str = "",
    include_mineru_wheels: bool = False,
    strict_mineru_wheels: bool = False,
    mineru_wheel_spec: str = DEFAULT_MINERU_WHEEL_SPEC,
    asset_cache_dir: str = "",
    reuse_assets: bool = True,
    include_datasheet_data: bool = True,
    include_datasheet_source: bool = False,
    make_zip: bool = True,
) -> dict:
    root = Path(project_root).expanduser().resolve()
    if not root.is_dir():
        raise FileNotFoundError(f"project root does not exist: {root}")
    has_python_source = any(str(value or "").strip() for value in [python_dir, python_archive, python_url, python_version])
    if not has_python_source and not allow_system_python_on_b:
        raise ValueError(
            "computer B may not have Python; provide --python-version, --python-url, "
            "--python-archive, or --python-dir, or pass --allow-system-python-on-b explicitly."
        )
    if not extract_python and not python_dir and not allow_system_python_on_b:
        raise ValueError(
            "--no-extract-python leaves only an archive; computer B without system Python "
            "cannot run verification. Remove --no-extract-python or pass --allow-system-python-on-b."
        )
    effective_mineru_venv = str(mineru_venv or "").strip()
    default_mineru_venv = root / ".venv-mineru"
    mineru_bootstrap_info: Dict[str, Any] = {
        "requested": False,
        "ok": None,
        "path": "",
        "auto_detected": False,
        "created": False,
        "installed": False,
    }
    if not effective_mineru_venv and default_mineru_venv.is_dir():
        effective_mineru_venv = str(default_mineru_venv)
        mineru_bootstrap_info.update({
            "requested": False,
            "ok": True,
            "path": str(default_mineru_venv),
            "auto_detected": True,
        })
    needs_mineru_runtime = bool(
        download_mineru_models
        or str(mineru_model_dir or "").strip()
        or str(mineru_config or "").strip()
    )
    if needs_mineru_runtime:
        bootstrap_target = Path(effective_mineru_venv).expanduser() if effective_mineru_venv else default_mineru_venv
        has_runtime = bool(_mineru_candidates(bootstrap_target))
        has_downloader_command = bool(
            _mineru_model_downloader_command_candidates(mineru_venv=bootstrap_target, explicit=mineru_model_downloader)
        )
        if not effective_mineru_venv or not bootstrap_target.exists() or not has_runtime or (download_mineru_models and not has_downloader_command):
            effective_mineru_venv = str(bootstrap_target)
            mineru_bootstrap_info = _bootstrap_mineru_venv(
                venv_root=bootstrap_target,
                mineru_spec=mineru_wheel_spec,
                pip_index_url=pip_index_url,
                pip_extra_index_url=pip_extra_index_url,
            )
    package_name = _safe_name(name)
    output_root = Path(out_dir).expanduser().resolve()
    asset_cache = _asset_cache_root(output_root=output_root, asset_cache_dir=asset_cache_dir, reuse_assets=reuse_assets)
    asset_cache_info: Dict[str, Any] = {
        "enabled": bool(asset_cache),
        "path": str(asset_cache or ""),
        "cache_version": OFFLINE_ASSET_CACHE_VERSION,
        "python_archive": {"enabled": bool(asset_cache), "hit": False, "path": "", "key": ""},
        "mineru_models": {"enabled": bool(asset_cache), "hit": False, "path": "", "key": ""},
        "wheelhouse": {"enabled": bool(asset_cache), "hit": False, "path": "", "key": ""},
    }
    bundle_root = output_root / package_name
    if bundle_root.exists():
        shutil.rmtree(bundle_root)
    bundle_root.mkdir(parents=True, exist_ok=True)

    project_target = bundle_root / "project"
    exclude_roots = [output_root, bundle_root]
    if asset_cache:
        exclude_roots.append(asset_cache)
    copied_project_items = _copy_project_tree(root, project_target, extra_exclude_roots=exclude_roots)
    runtime_dir = bundle_root / "runtime"
    runtime_dir.mkdir(parents=True, exist_ok=True)
    data_dir = bundle_root / "data"
    data_dir.mkdir(parents=True, exist_ok=True)

    python_info: Dict[str, Any] = {
        "provided": False,
        "required_on_target": not bool(allow_system_python_on_b),
        "kind": "",
        "source": "",
        "path": "",
    }
    archive_path: Optional[Path] = None
    if python_url or python_version:
        url = python_url or build_python_download_url(
            python_version=python_version,
            python_mirror=python_mirror,
            python_mirror_base=python_mirror_base,
            python_filename=python_filename,
        )
        filename = Path(urllib.parse.urlparse(url).path).name or "python-portable.zip"
        archive_path = runtime_dir / "python_archive" / filename
        python_cache = _prepare_python_archive(url=url, filename=filename, target=archive_path, asset_cache=asset_cache)
        asset_cache_info["python_archive"] = python_cache
        python_info.update({
            "provided": True,
            "kind": "downloaded_archive",
            "source": url,
            "path": archive_path.relative_to(bundle_root).as_posix(),
            "cache": python_cache,
        })
    elif python_archive:
        src = Path(python_archive).expanduser()
        if not src.is_file():
            raise FileNotFoundError(f"Python archive not found: {src}")
        archive_path = runtime_dir / "python_archive" / src.name
        archive_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, archive_path)
        python_info.update({"provided": True, "kind": "archive", "source": str(src), "path": archive_path.relative_to(bundle_root).as_posix()})
    elif python_dir:
        src = Path(python_dir).expanduser()
        if not src.is_dir():
            raise FileNotFoundError(f"Python directory not found: {src}")
        target = runtime_dir / "python"
        _copytree(src, target)
        python_info.update({"provided": True, "kind": "directory", "source": str(src), "path": target.relative_to(bundle_root).as_posix()})

    if archive_path and extract_python:
        target = runtime_dir / "python"
        _extract_archive(archive_path, target)
        python_info["extracted_path"] = target.relative_to(bundle_root).as_posix()

    mineru_info: Dict[str, Any] = {"provided": False, "source": "", "path": "", "venv_bootstrap": mineru_bootstrap_info}
    if effective_mineru_venv:
        target = runtime_dir / "mineru_venv"
        source = _copy_optional_dir(effective_mineru_venv, target)
        mineru_info.update({"provided": True, "source": source or "", "path": target.relative_to(bundle_root).as_posix()})
    model_download_info: Dict[str, Any] = {"requested": False, "ok": None}
    effective_mineru_model_dir = str(mineru_model_dir or "").strip()
    effective_mineru_config = str(mineru_config or "").strip()
    staging_config = runtime_dir / "mineru_download" / "mineru.json"
    if download_mineru_models and not effective_mineru_model_dir:
        cached_model_dir, cached_config, model_cache = _load_cached_mineru_models(
            asset_cache=asset_cache,
            model_source=mineru_model_source,
            model_type=mineru_model_type,
            huggingface_endpoint=huggingface_endpoint,
        )
        asset_cache_info["mineru_models"] = model_cache
        if cached_model_dir and cached_config:
            effective_mineru_model_dir = cached_model_dir
            effective_mineru_config = cached_config
            model_download_info = {
                "requested": True,
                "ok": True,
                "source": str(mineru_model_source or DEFAULT_MINERU_MODEL_SOURCE).strip().lower(),
                "model_type": str(mineru_model_type or DEFAULT_MINERU_MODEL_TYPE).strip().lower(),
                "huggingface_endpoint": str(huggingface_endpoint or "").strip(),
                "model_dir": cached_model_dir,
                "model_dirs": [cached_model_dir],
                "config_path": cached_config,
                "file_count": sum(1 for item in Path(cached_model_dir).rglob("*") if item.is_file()),
                "total_bytes": _dir_total_bytes(Path(cached_model_dir)),
                "cache": model_cache,
            }
        else:
            model_download_info = _download_mineru_models(
                mineru_venv=effective_mineru_venv,
                downloader=mineru_model_downloader,
                config_path=staging_config,
                model_source=mineru_model_source,
                model_type=mineru_model_type,
                huggingface_endpoint=huggingface_endpoint,
            )
            stored_cache = _store_cached_mineru_models(
                asset_cache=asset_cache,
                model_dir=str(model_download_info.get("model_dir") or ""),
                config_path=staging_config,
                model_source=mineru_model_source,
                model_type=mineru_model_type,
                huggingface_endpoint=huggingface_endpoint,
            )
            model_download_info["cache"] = stored_cache
            asset_cache_info["mineru_models"] = stored_cache
            effective_mineru_model_dir = str(model_download_info.get("model_dir") or "")
            effective_mineru_config = str(staging_config)
    mineru_assets = _copy_mineru_assets(
        model_dir=effective_mineru_model_dir,
        mineru_config=effective_mineru_config,
        runtime_dir=runtime_dir,
        bundle_root=bundle_root,
    )
    mineru_assets["models"]["download"] = model_download_info
    mineru_info["assets"] = mineru_assets
    if staging_config.exists():
        shutil.rmtree(staging_config.parent, ignore_errors=True)

    datasheet_info: Dict[str, Any] = {"data_included": False, "source_included": False}
    if include_datasheet_data and (root / "datasheet_data").is_dir():
        _copytree(root / "datasheet_data", data_dir / "datasheet_data")
        datasheet_info["data_included"] = True
    if include_datasheet_source:
        raw_dirs = [part.strip().strip('"') for part in os.environ.get("PSTX_DATASHEET_DIR", "").split(os.pathsep) if part.strip()]
        target_root = data_dir / "datasheets"
        copied_sources = []
        for index, raw_dir in enumerate(raw_dirs, start=1):
            src = Path(raw_dir).expanduser()
            if src.is_dir():
                dest = target_root / f"source_{index}_{_safe_name(src.name)}"
                _copytree(src, dest)
                copied_sources.append(str(src))
        datasheet_info["source_included"] = bool(copied_sources)
        datasheet_info["source_dirs"] = copied_sources

    wheel_info: Dict[str, Any] = {"downloaded": False, "ok": False}
    if download_wheels:
        strict_mineru_download = bool(strict_mineru_wheels or (include_mineru_wheels and not mineru_info.get("provided")))
        wheel_info = _prepare_wheelhouse_cached(
            bundle_wheelhouse=bundle_root / "wheelhouse",
            asset_cache=asset_cache,
            requirements=root / "requirements.txt",
            pip_index_url=pip_index_url,
            pip_extra_index_url=pip_extra_index_url,
            include_mineru=include_mineru_wheels,
            strict_mineru=strict_mineru_download,
            mineru_spec=mineru_wheel_spec,
        )
        asset_cache_info["wheelhouse"] = wheel_info.get("cache", asset_cache_info["wheelhouse"])
        if not wheel_info.get("ok"):
            phase = str(wheel_info.get("failure_phase") or "unknown")
            mineru_info_for_error = wheel_info.get("mineru", {}) or {}
            spec = str(mineru_info_for_error.get("spec") or mineru_wheel_spec or DEFAULT_MINERU_WHEEL_SPEC)
            reason = ""
            if phase == "mineru_runtime_wheels" and strict_mineru_download:
                reason = " MinerU wheel download is required because no copied MinerU venv was provided or --strict-mineru-wheels was used."
            raise RuntimeError(
                f"pip download failed during {phase} (mineru_spec={spec!r}).{reason} "
                f"{wheel_info.get('stderr') or wheel_info.get('stdout')}"
            )
        wheel_info["downloaded"] = True

    runtime_device = "cuda" if "cuda" in str(target_profile).lower() or "4060" in str(target_profile).lower() else "auto"
    runtime_model_source = "local" if mineru_assets.get("models", {}).get("provided") else "auto"
    _write_verify_scripts(bundle_root)
    _write_offline_readme(bundle_root, package_name=package_name)

    manifest = {
        "schema_version": OFFLINE_MIGRATION_SCHEMA_VERSION,
        "created_at": _now(),
        "package_name": package_name,
        "target_platform": target_platform,
        "target_profile": target_profile,
        "project": {
            "source_root": str(root),
            "copied_items": copied_project_items,
            "entrypoints": ["pstx_cli.py", "pstx_web.py", "pstx_local_ui.py"],
        },
        "python": python_info,
        "mineru": mineru_info,
        "datasheet": datasheet_info,
        "wheelhouse": wheel_info,
        "asset_cache": asset_cache_info,
        "runtime": {
            "profile": target_profile,
            "mineru": {
                "device": runtime_device,
                "model_source": runtime_model_source,
                "notes": [
                    "windows-rtx4060-cuda expects NVIDIA driver/CUDA-capable PyTorch wheels in the copied MinerU env or wheelhouse.",
                    "Computer B remains offline; RUN_SETUP_B.* only installs from local wheelhouse and patches bundle-local config paths.",
                ],
            },
        },
        "environment": {
            "PSTX_PDF_EXTRACTOR": "mineru",
            "PSTX_MINERU_BIN": "runtime/mineru_venv/bin/mineru or runtime\\mineru_venv\\Scripts\\mineru.exe",
            "PSTX_MINERU_DEVICE": runtime_device,
            "PSTX_MINERU_MODEL_SOURCE": runtime_model_source,
            "MINERU_MODEL_SOURCE": runtime_model_source,
            "MINERU_TOOLS_CONFIG_JSON": "runtime/mineru_config/mineru.json",
            "PSTX_DATASHEET_DATA_DIR": "data/datasheet_data",
        },
        "files": _collect_manifest_files(bundle_root),
    }
    (bundle_root / MANIFEST_NAME).write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")

    zip_path = None
    if make_zip:
        zip_base = output_root / package_name
        zip_path = shutil.make_archive(str(zip_base), "zip", root_dir=output_root, base_dir=package_name)
    return {
        "ok": True,
        "schema_version": OFFLINE_MIGRATION_SCHEMA_VERSION,
        "bundle_root": str(bundle_root),
        "zip_path": zip_path,
        "manifest_path": str(bundle_root / MANIFEST_NAME),
        "file_count": len(manifest["files"]),
        "python": python_info,
        "mineru": mineru_info,
        "target_profile": target_profile,
        "datasheet": datasheet_info,
        "wheelhouse": wheel_info,
        "asset_cache": asset_cache_info,
    }


def _load_manifest(package_root: Path) -> dict:
    manifest_path = package_root / MANIFEST_NAME
    if not manifest_path.is_file():
        raise FileNotFoundError(f"offline manifest not found: {manifest_path}")
    data = json.loads(manifest_path.read_text(encoding="utf-8"))
    if data.get("schema_version") != OFFLINE_MIGRATION_SCHEMA_VERSION:
        raise ValueError(f"unsupported offline manifest schema: {data.get('schema_version')}")
    return data


def _verify_hashes(package_root: Path, manifest: dict) -> List[dict]:
    issues: List[dict] = []
    for item in manifest.get("files", []) or []:
        rel = str(item.get("path") or "")
        path = package_root / rel
        if not path.is_file():
            issues.append({"path": rel, "status": "missing", "message": "file missing"})
            continue
        expected_size = int(item.get("size") or 0)
        if path.stat().st_size != expected_size:
            issues.append({"path": rel, "status": "size_mismatch", "message": f"expected {expected_size}, got {path.stat().st_size}"})
            continue
        actual_hash = _sha256(path)
        if actual_hash != str(item.get("sha256") or ""):
            issues.append({"path": rel, "status": "hash_mismatch", "message": "sha256 mismatch"})
    return issues


def _wheelhouse_missing(package_root: Path, requirements_path: Path) -> List[str]:
    if not requirements_path.is_file():
        return []
    wheelhouse = package_root / "wheelhouse"
    if not wheelhouse.is_dir():
        return []
    files = [item.name.lower().replace("-", "_") for item in wheelhouse.iterdir() if item.is_file()]
    missing: List[str] = []
    for raw_line in requirements_path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue
        name = line.split("==", 1)[0].split(">=", 1)[0].split("<", 1)[0].split("[", 1)[0].strip().lower().replace("-", "_")
        if name and not any(file.startswith(name) for file in files):
            missing.append(name)
    return missing


def _requirement_names(requirements_path: Path) -> List[str]:
    if not requirements_path.is_file():
        return []
    names: List[str] = []
    for raw_line in requirements_path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue
        for marker in ("==", ">=", "<=", "~=", "!=", "<", ">"):
            if marker in line:
                line = line.split(marker, 1)[0]
                break
        name = line.split("[", 1)[0].strip().lower().replace("-", "_")
        if name:
            names.append(name)
    return names


def _dependency_import_probe(python_bin: Path, project_root: Path, requirements_path: Path) -> dict:
    modules = sorted({
        REQUIREMENT_IMPORT_ALIASES.get(name, name)
        for name in _requirement_names(requirements_path)
    })
    if not modules:
        return {"ok": True, "checked": [], "missing": [], "error": ""}
    code = (
        "import importlib, json\n"
        f"mods = {modules!r}\n"
        "missing = []\n"
        "for mod in mods:\n"
        "    try:\n"
        "        importlib.import_module(mod)\n"
        "    except Exception as exc:\n"
        "        missing.append({'module': mod, 'error': str(exc)[:300]})\n"
        "print(json.dumps({'missing': missing}, ensure_ascii=False))\n"
        "raise SystemExit(1 if missing else 0)\n"
    )
    env = os.environ.copy()
    env["PYTHONPATH"] = str(project_root)
    try:
        proc = subprocess.run(
            [str(python_bin), "-c", code],
            cwd=str(project_root),
            env=env,
            capture_output=True,
            text=True,
            timeout=25,
            check=False,
        )
    except Exception as exc:
        return {"ok": False, "checked": modules, "missing": [], "error": str(exc)}
    output = (proc.stdout or "").strip().splitlines()
    try:
        parsed = json.loads(output[-1]) if output else {}
    except json.JSONDecodeError:
        parsed = {}
    missing = parsed.get("missing", []) if isinstance(parsed, dict) else []
    return {
        "ok": proc.returncode == 0 and not missing,
        "checked": modules,
        "missing": missing,
        "error": "" if proc.returncode == 0 else (proc.stderr or proc.stdout or "").strip()[-1000:],
    }


def verify_offline_bundle(package_root: str | Path, *, probe_runtime: bool = True) -> dict:
    root = Path(package_root).expanduser().resolve()
    temp_dir = None
    if root.is_file() and root.suffix.lower() == ".zip":
        import tempfile

        temp_dir = tempfile.TemporaryDirectory(prefix="pstx_offline_verify_")
        with zipfile.ZipFile(root) as zf:
            zf.extractall(temp_dir.name)
        extracted_children = [path for path in Path(temp_dir.name).iterdir() if path.is_dir()]
        root = extracted_children[0] if extracted_children else Path(temp_dir.name)
    try:
        manifest = _load_manifest(root)
        issues = _verify_hashes(root, manifest)
        warnings: List[str] = []

        project_root = root / "project"
        for rel in ("pstx_cli.py", "pstx_web.py", "pstx_apps/cli.py"):
            if not (project_root / rel).is_file():
                issues.append({"path": f"project/{rel}", "status": "missing", "message": "required project entrypoint missing"})

        python_info = manifest.get("python", {}) or {}
        portable_required = bool(python_info.get("required_on_target", True))
        python_candidates = _python_candidates(root / str(python_info.get("extracted_path") or python_info.get("path") or "runtime/python"))
        python_archive = root / str(python_info.get("path") or "")
        if python_info.get("provided") and not python_candidates:
            if python_archive.is_file():
                message = "portable Python archive is present but not extracted; computer B without system Python cannot run verification"
                if portable_required:
                    issues.append({"path": str(python_info.get("path") or "runtime/python"), "status": "python_not_extracted", "message": message})
                else:
                    warnings.append(message)
            else:
                issues.append({"path": str(python_info.get("path") or "runtime/python"), "status": "missing", "message": "portable Python missing"})
        if not python_info.get("provided"):
            message = "No portable Python was included; computer B must provide a compatible Python runtime."
            if portable_required:
                issues.append({"path": "runtime/python", "status": "portable_python_required", "message": message})
            else:
                warnings.append(message)

        mineru_info = manifest.get("mineru", {}) or {}
        mineru_candidates = _mineru_candidates(root / str(mineru_info.get("path") or "runtime/mineru_venv"))
        mineru_candidates.extend(_portable_python_mineru_candidates(root, manifest))
        if mineru_info.get("provided") and not mineru_candidates:
            issues.append({"path": str(mineru_info.get("path") or "runtime/mineru_venv"), "status": "missing", "message": "MinerU executable missing"})
        if not mineru_info.get("provided"):
            warnings.append("No MinerU venv was included; default datasheet PDF extraction will need PSTX_MINERU_BIN on computer B.")
        mineru_assets = mineru_info.get("assets", {}) or {}
        model_info = mineru_assets.get("models", {}) or {}
        if model_info.get("provided"):
            model_root = root / str(model_info.get("path") or "runtime/mineru_models")
            if not model_root.is_dir():
                issues.append({"path": str(model_info.get("path") or "runtime/mineru_models"), "status": "missing", "message": "MinerU model directory missing"})
            elif not any(path.is_file() for path in model_root.rglob("*")):
                issues.append({"path": str(model_info.get("path") or "runtime/mineru_models"), "status": "empty", "message": "MinerU model directory is empty"})
        else:
            warnings.append("No MinerU model directory was included; offline PDF extraction may fail on computer B.")
        config_info = mineru_assets.get("config", {}) or {}
        if config_info.get("provided"):
            template = root / str(config_info.get("template_path") or "")
            if not template.is_file():
                issues.append({"path": str(config_info.get("template_path") or "runtime/mineru_config/mineru.template.json"), "status": "missing", "message": "MinerU config template missing"})

        missing_wheels = _wheelhouse_missing(root, project_root / "requirements.txt")
        if missing_wheels:
            issues.append({"path": "wheelhouse", "status": "missing_wheels", "message": ",".join(missing_wheels)})

        dependency_probe = {"ok": None, "checked": [], "missing": [], "error": ""}
        if probe_runtime and python_candidates:
            dependency_probe = _dependency_import_probe(python_candidates[0], project_root, project_root / "requirements.txt")
            if dependency_probe.get("ok") is False:
                issues.append({
                    "path": str(python_candidates[0].relative_to(root)),
                    "status": "runtime_import_failed",
                    "message": json.dumps(dependency_probe.get("missing") or dependency_probe.get("error"), ensure_ascii=False)[:1000],
                })

        ok = not issues
        return {
            "ok": ok,
            "schema_version": OFFLINE_MIGRATION_SCHEMA_VERSION,
            "package_root": str(root),
            "package_name": manifest.get("package_name", ""),
            "target_platform": manifest.get("target_platform", ""),
            "target_profile": manifest.get("target_profile", ""),
            "file_count": len(manifest.get("files", []) or []),
            "checked_file_count": len(manifest.get("files", []) or []),
            "issues": issues,
            "warnings": warnings,
            "python": {
                "provided": bool(python_info.get("provided")),
                "candidates": [str(path.relative_to(root)) for path in python_candidates if path.exists()],
            },
            "mineru": {
                "provided": bool(mineru_info.get("provided")),
                "candidates": [str(path.relative_to(root)) for path in mineru_candidates if path.exists()],
            },
            "dependency_probe": dependency_probe,
            "summary": "离线迁移包校验通过。" if ok else f"离线迁移包校验发现 {len(issues)} 个问题。",
        }
    finally:
        if temp_dir is not None:
            temp_dir.cleanup()
