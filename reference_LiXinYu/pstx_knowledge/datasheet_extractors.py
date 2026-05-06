# -*- coding: utf-8 -*-
"""PDF extraction backends for local datasheet indexing."""

from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import Iterable, List, Optional, Tuple


PDF_EXTRACTOR_ENV = "PSTX_PDF_EXTRACTOR"
MINERU_BIN_ENV = "PSTX_MINERU_BIN"
MINERU_BACKEND_ENV = "PSTX_MINERU_BACKEND"
MINERU_DEVICE_ENV = "PSTX_MINERU_DEVICE"
MINERU_METHOD_ENV = "PSTX_MINERU_METHOD"
MINERU_MODEL_SOURCE_ENV = "PSTX_MINERU_MODEL_SOURCE"
MINERU_TIMEOUT_ENV = "PSTX_MINERU_TIMEOUT_SECONDS"
DEFAULT_PDF_EXTRACTOR = "mineru"
DEFAULT_MINERU_BACKEND = "pipeline"
DEFAULT_MINERU_DEVICE = "auto"
DEFAULT_MINERU_METHOD = "auto"
DEFAULT_MINERU_MODEL_SOURCE = "auto"
DEFAULT_MINERU_TIMEOUT_SECONDS = 600
SUPPORTED_PDF_EXTRACTORS = {"auto", "pypdf", "mineru"}
SUPPORTED_MINERU_DEVICES = {"auto", "cpu", "mps", "cuda"}
SUPPORTED_MINERU_METHODS = {"auto", "txt", "ocr"}
SUPPORTED_MINERU_MODEL_SOURCES = {"auto", "huggingface", "modelscope", "local"}


def safe_text(value, limit: int = 1000) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\r", " ").strip()
    return text if len(text) <= limit else text[:limit - 1] + "…"


def normalize_page_text(text: str) -> str:
    content = str(text or "").replace("\r\n", "\n").replace("\r", "\n")
    content = re.sub(r"[ \t]+\n", "\n", content)
    content = re.sub(r"\n{3,}", "\n\n", content)
    return content.strip()


def configured_pdf_extractor() -> str:
    mode = str(os.environ.get(PDF_EXTRACTOR_ENV) or DEFAULT_PDF_EXTRACTOR).strip().lower()
    return mode if mode in SUPPORTED_PDF_EXTRACTORS else DEFAULT_PDF_EXTRACTOR


def mineru_backend() -> str:
    return str(os.environ.get(MINERU_BACKEND_ENV) or DEFAULT_MINERU_BACKEND).strip() or DEFAULT_MINERU_BACKEND


def mineru_device() -> str:
    device = str(os.environ.get(MINERU_DEVICE_ENV) or DEFAULT_MINERU_DEVICE).strip().lower()
    return device if device in SUPPORTED_MINERU_DEVICES else DEFAULT_MINERU_DEVICE


def mineru_method() -> str:
    method = str(os.environ.get(MINERU_METHOD_ENV) or DEFAULT_MINERU_METHOD).strip().lower()
    return method if method in SUPPORTED_MINERU_METHODS else DEFAULT_MINERU_METHOD


def mineru_model_source() -> str:
    source = str(os.environ.get(MINERU_MODEL_SOURCE_ENV) or DEFAULT_MINERU_MODEL_SOURCE).strip().lower()
    return source if source in SUPPORTED_MINERU_MODEL_SOURCES else DEFAULT_MINERU_MODEL_SOURCE


def mineru_timeout_seconds() -> int:
    raw = os.environ.get(MINERU_TIMEOUT_ENV)
    try:
        value = int(raw) if raw is not None else DEFAULT_MINERU_TIMEOUT_SECONDS
    except (TypeError, ValueError):
        value = DEFAULT_MINERU_TIMEOUT_SECONDS
    return max(10, min(value, 3600))


def mineru_bin() -> Optional[str]:
    configured = str(os.environ.get(MINERU_BIN_ENV) or "").strip().strip('"')
    if configured:
        return configured
    return shutil.which("mineru")


def mineru_python_bin(bin_path: str) -> Optional[str]:
    candidate = Path(bin_path).expanduser().parent / "python"
    return str(candidate) if candidate.exists() else None


def probe_mineru_version(bin_path: str) -> Tuple[str, str]:
    python_bin = mineru_python_bin(bin_path)
    if python_bin:
        try:
            proc = subprocess.run(
                [
                    python_bin,
                    "-c",
                    "import importlib.metadata as m; print(m.version('mineru'))",
                ],
                capture_output=True,
                text=True,
                timeout=8,
                check=False,
            )
            output = "\n".join(part for part in [proc.stdout, proc.stderr] if part).strip()
            if proc.returncode == 0 and output:
                return f"mineru {safe_text(output.splitlines()[0], 120)}", ""
        except Exception as exc:
            metadata_error = f"metadata version probe failed: {exc}"
        else:
            metadata_error = safe_text(f"metadata version probe returned {proc.returncode}", 160)
    else:
        metadata_error = "venv python not found"

    try:
        proc = subprocess.run(
            [bin_path, "--version"],
            capture_output=True,
            text=True,
            timeout=8,
            check=False,
        )
        output = "\n".join(part for part in [proc.stdout, proc.stderr] if part).strip()
        version = safe_text(output.splitlines()[0] if output else "", 240)
        if proc.returncode == 0 or version:
            return version, ""
        return "", safe_text(f"mineru --version 返回 {proc.returncode}; {metadata_error}", 300)
    except Exception as exc:
        return "", safe_text(f"mineru 版本探测失败：{exc}; {metadata_error}", 300)


def build_mineru_status(*, include_version: bool = True) -> dict:
    configured_bin = str(os.environ.get(MINERU_BIN_ENV) or "").strip().strip('"')
    bin_path = mineru_bin()
    status = {
        "mode": configured_pdf_extractor(),
        "available": bool(bin_path),
        "bin": bin_path or configured_bin,
        "backend": mineru_backend(),
        "device": mineru_device(),
        "method": mineru_method(),
        "model_source": mineru_model_source(),
        "timeout_seconds": mineru_timeout_seconds(),
        "version": "",
        "error": "" if bin_path else "未找到 mineru CLI；默认 PDF 抽取需要 MinerU。可设置 PSTX_MINERU_BIN 或在 MinerU 专用 venv 中安装 mineru[all]；临时兼容可显式设置 PSTX_PDF_EXTRACTOR=auto 或 pypdf。",
        "env": {
            "extractor": PDF_EXTRACTOR_ENV,
            "bin": MINERU_BIN_ENV,
            "backend": MINERU_BACKEND_ENV,
            "device": MINERU_DEVICE_ENV,
            "method": MINERU_METHOD_ENV,
            "model_source": MINERU_MODEL_SOURCE_ENV,
            "timeout": MINERU_TIMEOUT_ENV,
            "mineru_device_mode": "MINERU_DEVICE_MODE",
            "mineru_model_source": "MINERU_MODEL_SOURCE",
        },
    }
    if not bin_path or not include_version:
        return status
    version, error = probe_mineru_version(bin_path)
    status["version"] = version
    if error:
        status["error"] = error
    return status


def fallback_extract_pdf_text(path: Path) -> List[str]:
    raw = path.read_bytes()
    text = raw.decode("latin-1", errors="ignore")
    chunks = re.findall(r"\(([^()]*)\)\s*Tj", text)
    array_chunks = re.findall(r"\[((?:.|\n)*?)\]\s*TJ", text)
    for array in array_chunks:
        chunks.extend(re.findall(r"\(([^()]*)\)", array))
    if chunks:
        joined = "\n".join(pdf_unescape(chunk) for chunk in chunks)
    else:
        printable = re.sub(r"[^\x09\x0a\x0d\x20-\x7e\u4e00-\u9fff]+", " ", text)
        joined = re.sub(r"\s+", " ", printable).strip()
    return [joined] if joined else []


def pdf_unescape(value: str) -> str:
    return (
        value.replace(r"\(", "(")
        .replace(r"\)", ")")
        .replace(r"\\", "\\")
        .replace(r"\n", "\n")
        .replace(r"\r", "\n")
    )


def extract_pdf_pages_with_pypdf(path: Path) -> Tuple[str, List[str], str, str]:
    try:
        from pypdf import PdfReader  # type: ignore
    except Exception:
        pages = fallback_extract_pdf_text(path)
        if pages and any(page.strip() for page in pages):
            return "indexed", pages, "fallback", ""
        return "needs_manual_review", [], "fallback", "未安装 pypdf，且 fallback 未抽取到文本。"

    try:
        reader = PdfReader(str(path))
        if getattr(reader, "is_encrypted", False):
            return "needs_manual_review", [], "pypdf", "PDF 已加密，第一版不尝试解密。"
        pages = []
        for page in reader.pages:
            try:
                pages.append(str(page.extract_text() or "").strip())
            except Exception:
                pages.append("")
        if not any(page.strip() for page in pages):
            return "needs_manual_review", pages, "pypdf", "PDF 未抽取到文本，可能是扫描版或图片版。"
        return "indexed", pages, "pypdf", ""
    except Exception as exc:
        pages = fallback_extract_pdf_text(path)
        if pages and any(page.strip() for page in pages):
            return "indexed", pages, "fallback", f"pypdf 抽取失败，已使用 fallback：{exc}"
        return "needs_manual_review", [], "pypdf", f"PDF 抽取失败：{exc}"


def json_text_fragments(value) -> List[str]:
    fragments: List[str] = []
    if isinstance(value, str):
        text = value.strip()
        if text:
            fragments.append(text)
    elif isinstance(value, list):
        for item in value:
            fragments.extend(json_text_fragments(item))
    return fragments


def json_item_page(item: dict) -> Optional[int]:
    for key in ("page", "page_no", "page_number", "pageNum", "page_id", "page_idx", "pageIndex"):
        if key not in item:
            continue
        try:
            page = int(item.get(key))
        except (TypeError, ValueError):
            continue
        if key in {"page_idx", "pageIndex"}:
            page += 1
        return max(1, page)
    return None


def json_item_text(item: dict) -> str:
    keys = (
        "section_title",
        "title",
        "text",
        "content",
        "md",
        "markdown",
        "html",
        "table_body",
        "table",
        "latex",
        "formula",
        "caption",
        "img_caption",
    )
    fragments: List[str] = []
    for key in keys:
        if key not in item:
            continue
        fragments.extend(json_text_fragments(item.get(key)))
    return "\n".join(fragment for fragment in fragments if fragment).strip()


def walk_json_items(value) -> Iterable[dict]:
    if isinstance(value, dict):
        yield value
        for child in value.values():
            yield from walk_json_items(child)
    elif isinstance(value, list):
        for child in value:
            yield from walk_json_items(child)


def read_mineru_json_pages(output_dir: Path) -> List[str]:
    page_text: dict[int, List[str]] = {}
    json_files = sorted(
        output_dir.rglob("*.json"),
        key=lambda item: (
            0 if "content" in item.name.lower() else 1 if "middle" in item.name.lower() else 2,
            item.as_posix().lower(),
        ),
    )
    for file in json_files[:24]:
        try:
            payload = json.loads(file.read_text(encoding="utf-8", errors="ignore"))
        except Exception:
            continue
        for item in walk_json_items(payload):
            page = json_item_page(item)
            text = json_item_text(item)
            if page is None or not text:
                continue
            page_text.setdefault(page, []).append(text)
    if not page_text:
        return []
    last_page = max(page_text)
    return [
        normalize_page_text("\n\n".join(page_text.get(page, [])))
        for page in range(1, last_page + 1)
    ]


def split_mineru_markdown_pages(markdown: str) -> List[str]:
    text = normalize_page_text(markdown)
    if not text:
        return []
    if "\f" in text:
        return [normalize_page_text(part) for part in text.split("\f")]
    marker_pattern = re.compile(r"(?im)^\s*(?:<!--\s*)?(?:page|页码|第)\s*[:#：]?\s*(\d+)\s*(?:页)?\s*(?:-->)?\s*$")
    matches = list(marker_pattern.finditer(text))
    if not matches:
        return [text]
    pages: List[str] = []
    for index, match in enumerate(matches):
        start = match.end()
        end = matches[index + 1].start() if index + 1 < len(matches) else len(text)
        pages.append(normalize_page_text(text[start:end]))
    return pages or [text]


def read_mineru_markdown_pages(output_dir: Path) -> List[str]:
    md_files = sorted(
        [*output_dir.rglob("*.md"), *output_dir.rglob("*.markdown")],
        key=lambda item: item.as_posix().lower(),
    )
    parts: List[str] = []
    for file in md_files[:24]:
        try:
            content = file.read_text(encoding="utf-8", errors="ignore").strip()
        except Exception:
            continue
        if content:
            parts.append(content)
    return split_mineru_markdown_pages("\n\n".join(parts))


def read_mineru_output_pages(output_dir: Path) -> List[str]:
    pages = read_mineru_json_pages(output_dir)
    if any(page.strip() for page in pages):
        return pages
    return read_mineru_markdown_pages(output_dir)


def extract_pdf_pages_with_mineru(path: Path) -> Tuple[str, List[str], str, str]:
    bin_path = mineru_bin()
    if not bin_path:
        return "needs_manual_review", [], "mineru", "MinerU CLI 不可用；请设置 PSTX_MINERU_BIN 或安装 mineru[all]。"
    backend = mineru_backend()
    device = mineru_device()
    method = mineru_method()
    model_source = mineru_model_source()
    timeout_seconds = mineru_timeout_seconds()
    with tempfile.TemporaryDirectory(prefix="pstx_mineru_") as tmp:
        output_dir = Path(tmp)
        cmd = [bin_path, "-p", str(path), "-o", str(output_dir), "-b", backend, "-m", method]
        env = os.environ.copy()
        if device != "auto":
            env["MINERU_DEVICE_MODE"] = device
        if model_source != "auto":
            env["MINERU_MODEL_SOURCE"] = model_source
        try:
            proc = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                env=env,
                timeout=timeout_seconds,
                check=False,
            )
        except subprocess.TimeoutExpired:
            return "needs_manual_review", [], "mineru", f"MinerU 抽取超时（>{timeout_seconds}s）。"
        except Exception as exc:
            return "needs_manual_review", [], "mineru", f"MinerU 调用失败：{exc}"
        if proc.returncode != 0:
            detail = "\n".join(part for part in [proc.stdout, proc.stderr] if part).strip()
            return "needs_manual_review", [], "mineru", safe_text(f"MinerU 返回 {proc.returncode}：{detail}", 2000)
        pages = read_mineru_output_pages(output_dir)
        if any(page.strip() for page in pages):
            return "indexed", pages, "mineru", ""
        detail = "\n".join(part for part in [proc.stdout, proc.stderr] if part).strip()
        return "needs_manual_review", pages, "mineru", safe_text(f"MinerU 未生成可用文本输出。{detail}", 2000)
