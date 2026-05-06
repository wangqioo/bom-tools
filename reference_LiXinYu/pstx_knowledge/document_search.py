# -*- coding: utf-8 -*-
"""Lightweight local document search for harness agents.

This module intentionally avoids a persistent index in the first version.  It
scans configured document roots, extracts best-effort text, returns compact
keyword hits, and lets the agent fetch a bounded excerpt around a hit.
"""

from __future__ import annotations

from dataclasses import dataclass
import hashlib
import html
import json
import os
from pathlib import Path
import re
import zipfile
from typing import Iterable, List, Mapping, Optional, Sequence
import xml.etree.ElementTree as ET


DEFAULT_DOC_DIR_NAMES = ("harness_docs", "ref", "ref_checklist", "docs")
SUPPORTED_SUFFIXES = {
    ".txt",
    ".md",
    ".markdown",
    ".rst",
    ".csv",
    ".tsv",
    ".json",
    ".yaml",
    ".yml",
    ".xml",
    ".html",
    ".htm",
    ".log",
    ".docx",
    ".xlsx",
    ".pdf",
}
MAX_FILE_BYTES = 15 * 1024 * 1024
DEFAULT_MAX_FILES = 500


@dataclass(frozen=True)
class DocumentRecord:
    doc_id: str
    root: Path
    path: Path

    @property
    def rel_path(self) -> str:
        try:
            return str(self.path.relative_to(self.root))
        except ValueError:
            return self.path.name


def _safe_text(value: object, limit: int = 500) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\r", " ").strip()
    return text if len(text) <= limit else text[:limit - 1] + "…"


def _doc_id(path: Path) -> str:
    return hashlib.sha1(str(path.resolve()).encode("utf-8", errors="ignore")).hexdigest()[:16]


def configured_document_roots() -> List[Path]:
    raw = os.environ.get("PSTX_HARNESS_DOC_DIR", "")
    roots: List[Path] = []
    if raw.strip():
        roots.extend(Path(item).expanduser() for item in raw.split(os.pathsep) if item.strip())
    cwd = Path.cwd()
    for name in DEFAULT_DOC_DIR_NAMES:
        roots.append(cwd / name)
    result: List[Path] = []
    seen = set()
    for root in roots:
        try:
            resolved = root.resolve()
        except OSError:
            continue
        if resolved in seen or not resolved.exists() or not resolved.is_dir():
            continue
        seen.add(resolved)
        result.append(resolved)
    return result


def _iter_documents(*, max_files: int = DEFAULT_MAX_FILES) -> List[DocumentRecord]:
    records: List[DocumentRecord] = []
    for root in configured_document_roots():
        for path in sorted(root.rglob("*")):
            if len(records) >= max_files:
                return records
            if path.is_symlink() or not path.is_file() or path.suffix.lower() not in SUPPORTED_SUFFIXES:
                continue
            try:
                if path.stat().st_size > MAX_FILE_BYTES:
                    continue
            except OSError:
                continue
            records.append(DocumentRecord(doc_id=_doc_id(path), root=root, path=path))
    return records


def _extract_docx_text(path: Path) -> str:
    try:
        with zipfile.ZipFile(path) as archive:
            xml_text = archive.read("word/document.xml")
    except Exception:
        return ""
    try:
        root = ET.fromstring(xml_text)
    except ET.ParseError:
        return ""
    texts = []
    for elem in root.iter():
        if elem.tag.endswith("}t") and elem.text:
            texts.append(elem.text)
        elif elem.tag.endswith("}tab"):
            texts.append("\t")
        elif elem.tag.endswith("}br"):
            texts.append("\n")
    return " ".join(texts)


def _extract_xlsx_text(path: Path) -> str:
    try:
        with zipfile.ZipFile(path) as archive:
            shared = []
            if "xl/sharedStrings.xml" in archive.namelist():
                root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
                for si in root:
                    parts = [node.text or "" for node in si.iter() if node.tag.endswith("}t")]
                    shared.append("".join(parts))
            chunks = []
            for name in archive.namelist():
                if not re.match(r"xl/worksheets/sheet\d+\.xml$", name):
                    continue
                sheet_root = ET.fromstring(archive.read(name))
                for cell in sheet_root.iter():
                    if not cell.tag.endswith("}c"):
                        continue
                    cell_type = cell.attrib.get("t", "")
                    value = ""
                    for child in cell:
                        if child.tag.endswith("}v") and child.text:
                            value = child.text
                            break
                    if cell_type == "s":
                        try:
                            value = shared[int(value)]
                        except (ValueError, IndexError):
                            pass
                    if value:
                        chunks.append(value)
            return "\n".join(chunks)
    except Exception:
        return ""


def _extract_pdf_text(path: Path) -> str:
    try:
        from pypdf import PdfReader  # type: ignore
    except Exception:
        return ""
    try:
        reader = PdfReader(str(path))
        pages = []
        for page in reader.pages[:200]:
            pages.append(page.extract_text() or "")
        return "\n".join(pages)
    except Exception:
        return ""


def extract_document_text(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix == ".docx":
        return _extract_docx_text(path)
    if suffix == ".xlsx":
        return _extract_xlsx_text(path)
    if suffix == ".pdf":
        return _extract_pdf_text(path)
    try:
        data = path.read_bytes()
    except OSError:
        return ""
    for encoding in ("utf-8-sig", "utf-8", "gb18030", "latin-1"):
        try:
            text = data.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    else:
        text = data.decode("utf-8", errors="ignore")
    if suffix in {".html", ".htm"}:
        text = re.sub(r"<script\b.*?</script>", " ", text, flags=re.IGNORECASE | re.DOTALL)
        text = re.sub(r"<style\b.*?</style>", " ", text, flags=re.IGNORECASE | re.DOTALL)
        text = re.sub(r"<[^>]+>", " ", text)
        text = html.unescape(text)
    return text


def _normalize_text(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip()


def _terms(query: object) -> List[str]:
    raw = str(query or "").strip()
    terms = [item for item in re.findall(r"[0-9A-Za-z_\u4e00-\u9fff.+#-]+", raw) if len(item.strip()) >= 1]
    stopwords = {"的", "和", "与", "请", "搜索", "查找", "文档", "内容", "段落", "keyword", "search"}
    result = []
    for item in terms:
        lowered = item.lower()
        if lowered in stopwords or item in stopwords:
            continue
        if lowered not in [existing.lower() for existing in result]:
            result.append(item)
    return result[:12]


def _line_number(text: str, char_start: int) -> int:
    return text.count("\n", 0, max(0, char_start)) + 1


def _snippet(text: str, start: int, end: int, window: int = 220) -> str:
    left = max(0, start - window)
    right = min(len(text), end + window)
    return _safe_text(_normalize_text(text[left:right]), 520)


def build_document_search_status() -> dict:
    roots = configured_document_roots()
    records = _iter_documents(max_files=DEFAULT_MAX_FILES)
    suffix_counts = {}
    for record in records:
        suffix = record.path.suffix.lower()
        suffix_counts[suffix] = suffix_counts.get(suffix, 0) + 1
    return {
        "ok": True,
        "configured_roots": [str(root) for root in roots],
        "document_count": len(records),
        "suffix_counts": suffix_counts,
        "supported_suffixes": sorted(SUPPORTED_SUFFIXES),
        "summary": f"本地文档搜索根目录 {len(roots)} 个，可搜索文档 {len(records)} 个。",
    }


def search_documents(query: object, *, limit: int = 20, max_files: int = DEFAULT_MAX_FILES) -> dict:
    terms = _terms(query)
    limit = max(1, min(int(limit or 20), 100))
    max_files = max(1, min(int(max_files or DEFAULT_MAX_FILES), 5000))
    if not terms:
        return {
            "ok": True,
            "query": _safe_text(query, 240),
            "terms": [],
            "matches": [],
            "total_matches": 0,
            "summary": "文档搜索缺少有效关键词。",
        }
    matches = []
    scanned_files = 0
    failed_files = 0
    for record in _iter_documents(max_files=max_files):
        scanned_files += 1
        text = extract_document_text(record.path)
        if not text:
            failed_files += 1
            continue
        lowered = text.lower()
        for term in terms:
            start = lowered.find(term.lower())
            if start < 0:
                continue
            end = start + len(term)
            matches.append({
                "doc_id": record.doc_id,
                "title": record.path.name,
                "rel_path": record.rel_path,
                "root": str(record.root),
                "suffix": record.path.suffix.lower(),
                "matched_term": term,
                "line_number": _line_number(text, start),
                "char_start": start,
                "char_end": end,
                "snippet": _snippet(text, start, end),
            })
            break
    selected = matches[:limit]
    return {
        "ok": True,
        "query": _safe_text(query, 240),
        "terms": terms,
        "summary": f"文档搜索 `{_safe_text(query, 80)}` 命中 {len(matches)} 个文件片段，返回 {len(selected)} 个。",
        "scanned_files": scanned_files,
        "failed_extract_files": failed_files,
        "total_matches": len(matches),
        "limit": limit,
        "truncated": len(matches) > len(selected),
        "matches": selected,
    }


def _record_by_doc_id(doc_id: object, *, max_files: int = DEFAULT_MAX_FILES) -> Optional[DocumentRecord]:
    target = str(doc_id or "").strip()
    if not target:
        return None
    for record in _iter_documents(max_files=max_files):
        if record.doc_id == target:
            return record
    return None


def get_document_excerpt(doc_id: object,
                         *,
                         char_start: int = 0,
                         before_chars: int = 800,
                         after_chars: int = 1600,
                         max_chars: int = 5000) -> dict:
    record = _record_by_doc_id(doc_id, max_files=5000)
    if record is None:
        return {"ok": False, "summary": f"未找到文档 doc_id={doc_id}", "doc_id": str(doc_id or "")}
    text = extract_document_text(record.path)
    if not text:
        return {"ok": False, "summary": f"文档无法抽取文本：{record.rel_path}", "doc_id": record.doc_id}
    start = max(0, int(char_start or 0) - max(0, int(before_chars or 0)))
    end = min(len(text), int(char_start or 0) + max(0, int(after_chars or 0)))
    max_chars = max(1, min(int(max_chars or 5000), 20000))
    excerpt = text[start:end]
    truncated = len(excerpt) > max_chars
    excerpt = excerpt[:max_chars]
    return {
        "ok": True,
        "summary": f"读取 {record.rel_path} 第 {_line_number(text, max(0, int(char_start or 0)))} 行附近片段。",
        "doc_id": record.doc_id,
        "title": record.path.name,
        "rel_path": record.rel_path,
        "root": str(record.root),
        "suffix": record.path.suffix.lower(),
        "line_number": _line_number(text, max(0, int(char_start or 0))),
        "char_start": start,
        "char_end": start + len(excerpt),
        "truncated": truncated,
        "excerpt": excerpt,
    }


def batch_search_documents(queries: Sequence[object], *, limit_per_query: int = 8) -> dict:
    normalized = [_safe_text(item, 240) for item in list(queries or [])[:20] if _safe_text(item, 240)]
    limit_per_query = max(1, min(int(limit_per_query or 8), 20))
    items = []
    found = 0
    for query in normalized:
        result = search_documents(query, limit=limit_per_query)
        status = "found" if result.get("total_matches", 0) else "missing"
        if status == "found":
            found += 1
        items.append({
            "query": query,
            "status": status,
            "summary": result.get("summary", ""),
            "total_matches": result.get("total_matches", 0),
            "matches": result.get("matches", []),
            "missing_reason": "" if status == "found" else "本地文档未命中该关键词。",
            "truncated": bool(result.get("truncated")),
        })
    return {
        "ok": True,
        "summary": f"批量文档搜索 {len(normalized)} 项，命中 {found} 项。",
        "query_count": len(normalized),
        "found_count": found,
        "missing_count": len(normalized) - found,
        "items": items,
    }
