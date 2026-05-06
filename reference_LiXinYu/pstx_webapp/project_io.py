"""Project input discovery and text decoding helpers for the Web app."""

from __future__ import annotations

import hashlib
import os
from pathlib import Path
import shutil
import tarfile
import time
from typing import Dict, List, Optional, Tuple
import uuid
import zipfile


TEXT_DECODE_ENCODINGS = (
    "utf-8-sig",
    "utf-8",
    "utf-16",
    "utf-16-le",
    "utf-16-be",
    "gb18030",
    "cp936",
)

TEXT_DECODE_MARKERS = (
    "PART_NAME",
    "NET_NAME",
    "NODE_NAME",
    "SECTION_NUMBER",
    "PAGE_NUMBER",
    "BOM_OPTION",
    "C_PATH",
    "P_PATH",
)

ARCHIVE_SUFFIXES = (
    ".zip",
    ".tar",
    ".tar.gz",
    ".tgz",
    ".tar.bz2",
    ".tbz2",
    ".tar.xz",
    ".txz",
)


def _is_supported_archive(path: Path) -> bool:
    name = path.name.lower()
    return any(name.endswith(suffix) for suffix in ARCHIVE_SUFFIXES)


def _archive_stem(path: Path) -> str:
    name = path.name
    lower = name.lower()
    for suffix in sorted(ARCHIVE_SUFFIXES, key=len, reverse=True):
        if lower.endswith(suffix):
            return name[: -len(suffix)]
    return path.stem


def _snapshot_base_dir() -> Path:
    raw = os.environ.get("PSTX_PROJECT_SNAPSHOT_DIR", "").strip()
    return Path(raw).expanduser() if raw else Path.cwd() / "output" / "project_snapshots"


def _safe_part(value: str) -> str:
    cleaned = "".join(char if char.isalnum() or char in {"-", "_", "."} else "_" for char in str(value or ""))
    return cleaned.strip("._")[:64] or "project"


def _copy_archive_to_snapshot(archive_path: Path) -> Tuple[Path, Path]:
    stat = archive_path.stat()
    digest = hashlib.sha256(
        f"{archive_path.resolve()}|{stat.st_size}|{int(stat.st_mtime)}".encode("utf-8", errors="replace")
    ).hexdigest()[:10]
    snapshot_root = _snapshot_base_dir() / f"{time.strftime('%Y%m%d_%H%M%S')}_{_safe_part(_archive_stem(archive_path))}_{digest}_{uuid.uuid4().hex[:6]}"
    archive_dir = snapshot_root / "archive"
    archive_dir.mkdir(parents=True, exist_ok=False)
    local_copy = archive_dir / archive_path.name
    shutil.copy2(archive_path, local_copy)
    return snapshot_root, local_copy


def _ensure_within_directory(root: Path, target: Path) -> None:
    root_resolved = root.resolve()
    target_resolved = target.resolve()
    if target_resolved == root_resolved:
        return
    if root_resolved not in target_resolved.parents:
        raise ValueError(f"压缩包包含不安全路径，已拒绝解压：{target}")


def _extract_zip_safely(archive_path: Path, extract_dir: Path) -> None:
    with zipfile.ZipFile(archive_path) as archive:
        for member in archive.infolist():
            target = extract_dir / member.filename
            _ensure_within_directory(extract_dir, target)
        archive.extractall(extract_dir)


def _extract_tar_safely(archive_path: Path, extract_dir: Path) -> None:
    with tarfile.open(archive_path) as archive:
        for member in archive.getmembers():
            if member.islnk() or member.issym():
                raise ValueError(f"压缩包包含链接文件，已拒绝解压：{member.name}")
            if not (member.isdir() or member.isfile()):
                raise ValueError(f"压缩包包含非普通文件，已拒绝解压：{member.name}")
            target = extract_dir / member.name
            _ensure_within_directory(extract_dir, target)
        archive.extractall(extract_dir)


def _extract_archive(archive_path: Path, extract_dir: Path) -> None:
    extract_dir.mkdir(parents=True, exist_ok=True)
    name = archive_path.name.lower()
    if name.endswith(".zip"):
        _extract_zip_safely(archive_path, extract_dir)
        return
    if any(name.endswith(suffix) for suffix in ARCHIVE_SUFFIXES if suffix != ".zip"):
        _extract_tar_safely(archive_path, extract_dir)
        return
    raise ValueError(f"暂不支持该压缩包格式：{archive_path.name}")


def _cpm_files(container: Path) -> List[Path]:
    return sorted(
        [path for path in container.glob("*.cpm") if path.is_file()],
        key=lambda item: item.name.lower(),
    )


def _module_name_from_cpm(container: Path) -> Tuple[Optional[str], List[str]]:
    cpm_files = _cpm_files(container)
    if not cpm_files:
        return None, []
    if len(cpm_files) > 1:
        names = ", ".join(path.name for path in cpm_files)
        raise ValueError(f"发现多个 .cpm 文件，无法确定主模块：{names}")
    return cpm_files[0].stem, [str(cpm_files[0])]


def _project_has_packaged_files(project_root: Path) -> bool:
    packaged = project_root / "packaged"
    return (
        packaged.is_dir()
        and (packaged / "pstxprt.dat").is_file()
        and (packaged / "pstxnet.dat").is_file()
    )


def _candidate_project_roots(base: Path, preferred_module: str = "") -> List[Path]:
    candidates: List[Path] = []
    if base.name.lower() == "packaged":
        candidates.append(base.parent)
    candidates.append(base)
    if base.name.lower() == "worklib" and preferred_module:
        candidates.append(base / preferred_module)
    module_name = preferred_module
    if not module_name and base.is_dir():
        cpm_module, _ = _module_name_from_cpm(base)
        module_name = cpm_module or ""
    if module_name:
        candidates.append(base / "worklib" / module_name)
        candidates.append(base / "WORKLIB" / module_name)
    if base.is_dir():
        for packaged in base.rglob("packaged"):
            if len(packaged.relative_to(base).parts) > 7:
                continue
            if (packaged / "pstxprt.dat").is_file() and (packaged / "pstxnet.dat").is_file():
                candidates.append(packaged.parent)
    deduped: List[Path] = []
    seen = set()
    for candidate in candidates:
        key = str(candidate)
        if key in seen:
            continue
        seen.add(key)
        deduped.append(candidate)
    return deduped


def _score_project_root(path: Path, preferred_module: str = "") -> Tuple[int, str]:
    score = 0
    if _project_has_packaged_files(path):
        score += 1000
    if preferred_module and path.name.lower() == preferred_module.lower():
        score += 200
    if path.parent.name.lower() == "worklib":
        score += 120
    if (path / "sch_1").is_dir():
        score += 60
    return score, str(path).lower()


def _locate_project_root(base: Path, preferred_module: str = "") -> Path:
    candidates = [
        path for path in _candidate_project_roots(base, preferred_module)
        if path.exists() and path.is_dir() and _project_has_packaged_files(path)
    ]
    if not candidates:
        raise ValueError(f"未在路径中找到 packaged/pstxprt.dat 与 pstxnet.dat：{base}")
    candidates.sort(key=lambda item: _score_project_root(item, preferred_module), reverse=True)
    return candidates[0]


def _archive_candidates(container: Path, module_name: str = "") -> List[Path]:
    if not container.is_dir():
        return []
    direct = [path for path in container.iterdir() if path.is_file() and _is_supported_archive(path)]
    module_upper = module_name.upper()
    container_upper = container.name.upper()
    sibling_candidates: List[Path] = []
    parent = container.parent
    if parent.is_dir():
        sibling_candidates = [
            path for path in parent.iterdir()
            if path.is_file()
            and _is_supported_archive(path)
            and path.parent != container
            and (
                (module_upper and module_upper in _archive_stem(path).upper())
                or (container_upper and container_upper in _archive_stem(path).upper())
            )
        ]
    all_candidates = direct + sibling_candidates
    if not all_candidates:
        return []
    preferred = [
        path for path in all_candidates
        if (module_upper and module_upper in _archive_stem(path).upper())
        or (container_upper and container_upper in _archive_stem(path).upper())
    ]
    return preferred or all_candidates


def _choose_archive(container: Path, module_name: str = "") -> Tuple[Optional[Path], List[str]]:
    candidates = _archive_candidates(container, module_name)
    if not candidates:
        return None, []
    candidates.sort(key=lambda path: (path.stat().st_mtime, path.name.lower()), reverse=True)
    warnings = []
    if len(candidates) > 1:
        warnings.append(
            f"发现多个项目压缩包，已选择最近修改的 {candidates[0].name}；其余候选："
            f"{', '.join(path.name for path in candidates[1:4])}"
        )
    return candidates[0], warnings


def _resolve_directory_project_root(root: Path,
                                    *,
                                    allow_archive: bool = True,
                                    preferred_module: str = "") -> Tuple[Path, Dict[str, object]]:
    meta: Dict[str, object] = {
        "enabled": False,
        "input": str(root),
        "mode": "directory",
        "warnings": [],
    }
    if root.name.lower() == "packaged":
        root = root.parent
    container = root
    module_name = preferred_module
    cpm_sources: List[str] = []
    if root.name.lower() == "worklib":
        container = root.parent
        module_name, cpm_sources = _module_name_from_cpm(container)
    elif root.parent.name.lower() == "worklib":
        module_name = module_name or root.name
        container = root.parent.parent
    else:
        cpm_module, cpm_sources = _module_name_from_cpm(root)
        if cpm_module:
            module_name = module_name or cpm_module

    if allow_archive and container.is_dir() and module_name:
        archive, archive_warnings = _choose_archive(container, module_name)
        if archive:
            project_root, snapshot_meta = _resolve_archive_project_root(
                archive,
                preferred_module=module_name,
                source_container=container,
                cpm_sources=cpm_sources,
            )
            snapshot_meta.setdefault("warnings", []).extend(archive_warnings)
            return project_root, snapshot_meta

    if _project_has_packaged_files(root):
        meta.update({
            "resolved_project_root": str(root),
            "project_container": str(container),
            "module_name": module_name or "",
            "cpm_files": cpm_sources,
        })
        return root, meta

    if module_name and (container / "worklib" / module_name).is_dir():
        project_root = container / "worklib" / module_name
    elif module_name and (container / "WORKLIB" / module_name).is_dir():
        project_root = container / "WORKLIB" / module_name
    else:
        project_root = root
    project_root = _locate_project_root(project_root, module_name or "")
    meta.update({
        "resolved_project_root": str(project_root),
        "project_container": str(container),
        "module_name": module_name or "",
        "cpm_files": cpm_sources,
    })
    return project_root, meta


def _resolve_archive_project_root(archive_path: Path,
                                  *,
                                  preferred_module: str = "",
                                  source_container: Optional[Path] = None,
                                  cpm_sources: Optional[List[str]] = None) -> Tuple[Path, Dict[str, object]]:
    if not archive_path.is_file():
        raise ValueError(f"项目压缩包不存在：{archive_path}")
    if not _is_supported_archive(archive_path):
        raise ValueError(f"暂不支持该压缩包格式：{archive_path.name}")
    snapshot_root, local_copy = _copy_archive_to_snapshot(archive_path)
    extract_dir = snapshot_root / "extracted"
    _extract_archive(local_copy, extract_dir)
    project_root = _locate_project_root(extract_dir, preferred_module)
    return project_root, {
        "enabled": True,
        "mode": "archive",
        "input": str(archive_path),
        "source_archive": str(archive_path),
        "local_archive": str(local_copy),
        "snapshot_root": str(snapshot_root),
        "extract_dir": str(extract_dir),
        "resolved_project_root": str(project_root),
        "project_container": str(source_container) if source_container else "",
        "module_name": preferred_module,
        "cpm_files": list(cpm_sources or []),
        "warnings": [],
    }


def _score_decoded_text(text: str) -> int:
    upper_text = str(text or "").upper()
    marker_score = sum(upper_text.count(marker) for marker in TEXT_DECODE_MARKERS) * 1000
    control_penalty = sum(
        1
        for char in text
        if ord(char) < 32 and char not in {"\r", "\n", "\t"}
    ) * 50
    ascii_score = sum(1 for char in text if 32 <= ord(char) < 127)
    return marker_score + ascii_score - control_penalty


def _decode_text_bytes(data: bytes) -> Tuple[str, str]:
    candidates = []
    for order, encoding in enumerate(TEXT_DECODE_ENCODINGS):
        try:
            text = data.decode(encoding)
        except UnicodeDecodeError:
            continue
        candidates.append((_score_decoded_text(text), -order, text, encoding))
    if candidates:
        _, _, text, encoding = max(candidates, key=lambda item: (item[0], item[1]))
        return text, encoding
    return data.decode("utf-8", errors="replace"), "utf-8-replace"


def read_local_text_file(path: Path, label: str, required: bool) -> Tuple[Optional[str], Dict[str, str]]:
    if not path.exists():
        if required:
            raise ValueError(f"缺少必需文件：{path}")
        return None, {"label": label, "filename": str(path), "size": "0", "encoding": ""}
    data = path.read_bytes()
    text, encoding = _decode_text_bytes(data)
    return text, {
        "label": label,
        "filename": str(path),
        "size": str(len(data)),
        "encoding": encoding,
    }


def resolve_project_root(root_text: str) -> Path:
    raw = (root_text or "").strip().strip('"')
    if not raw:
        raise ValueError("请输入项目根路径")
    root = Path(raw).expanduser()
    if not root.exists():
        raise ValueError(f"项目根路径不存在：{root}")
    if root.is_file():
        if _is_supported_archive(root):
            project_root, _ = _resolve_archive_project_root(root)
            return project_root
        raise ValueError(f"项目根路径不是文件夹或支持的压缩包：{root}")
    project_root, _ = _resolve_directory_project_root(root)
    return project_root


def discover_project_files_with_snapshot(root_text: str) -> Tuple[Path, Path, Path, Optional[Path], Dict[str, object]]:
    raw = (root_text or "").strip().strip('"')
    if not raw:
        raise ValueError("请输入项目根路径")
    input_path = Path(raw).expanduser()
    if not input_path.exists():
        raise ValueError(f"项目根路径不存在：{input_path}")
    if input_path.is_file():
        if not _is_supported_archive(input_path):
            raise ValueError(f"项目根路径不是文件夹或支持的压缩包：{input_path}")
        project_root, snapshot = _resolve_archive_project_root(input_path)
    else:
        project_root, snapshot = _resolve_directory_project_root(input_path)
    packaged_dir = project_root / "packaged"
    if not packaged_dir.is_dir():
        raise ValueError(f"项目根路径下缺少 packaged 文件夹：{packaged_dir}")

    prt_path = packaged_dir / "pstxprt.dat"
    net_path = packaged_dir / "pstxnet.dat"
    ref_path = packaged_dir / "pstxref.dat"
    if not prt_path.is_file():
        raise ValueError(f"未找到输入文件：{prt_path}")
    if not net_path.is_file():
        raise ValueError(f"未找到输入文件：{net_path}")
    snapshot.setdefault("resolved_project_root", str(project_root))
    return project_root, prt_path, net_path, (ref_path if ref_path.is_file() else None), snapshot


def discover_project_files(root_text: str) -> Tuple[Path, Path, Path, Optional[Path]]:
    project_root, prt_path, net_path, ref_path, _snapshot = discover_project_files_with_snapshot(root_text)
    return project_root, prt_path, net_path, ref_path


def decode_upload(file_storage, label: str, required: bool) -> Tuple[Optional[str], Dict[str, str]]:
    if not file_storage or not getattr(file_storage, "filename", ""):
        if required:
            raise ValueError(f"请上传 {label}")
        return None, {"label": label, "filename": "", "size": "0", "encoding": ""}
    data = file_storage.read()
    text, encoding = _decode_text_bytes(data)
    return text, {
        "label": label,
        "filename": file_storage.filename,
        "size": str(len(data)),
        "encoding": encoding,
    }
