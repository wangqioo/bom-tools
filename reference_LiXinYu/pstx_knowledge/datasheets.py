# -*- coding: utf-8 -*-
"""Local datasheet PDF index for DFMEA harness tools."""

from __future__ import annotations

import os
import re
import sqlite3
import time
from pathlib import Path
from typing import Iterable, List, Optional, Tuple

from pstx_knowledge.datasheet_chunks import chunk_page_text as _chunk_page_text
from pstx_knowledge.datasheet_parameters import extract_datasheet_parameters as _extract_datasheet_parameters
from pstx_knowledge.datasheet_extractors import (
    PDF_EXTRACTOR_ENV,
    MINERU_BIN_ENV,
    MINERU_BACKEND_ENV,
    MINERU_TIMEOUT_ENV,
    DEFAULT_PDF_EXTRACTOR,
    DEFAULT_MINERU_BACKEND,
    DEFAULT_MINERU_TIMEOUT_SECONDS,
    SUPPORTED_PDF_EXTRACTORS,
    build_mineru_status,
    configured_pdf_extractor,
    extract_pdf_pages_with_mineru as _extract_pdf_pages_with_mineru,
    extract_pdf_pages_with_pypdf as _extract_pdf_pages_with_pypdf,
)


DATASHEET_DIR_ENV = "PSTX_DATASHEET_DIR"
DATASHEET_DATA_DIR_ENV = "PSTX_DATASHEET_DATA_DIR"
DEFAULT_DATASHEET_DATA_DIR = "datasheet_data"
DB_NAME = "datasheet_index.db"


def _now() -> str:
    return time.strftime("%Y-%m-%dT%H:%M:%S", time.localtime())


def _safe_text(value, limit: int = 1000) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\r", " ").strip()
    return text if len(text) <= limit else text[:limit - 1] + "…"


def _data_dir() -> Path:
    raw = os.environ.get(DATASHEET_DATA_DIR_ENV) or DEFAULT_DATASHEET_DATA_DIR
    path = Path(raw).expanduser()
    if not path.is_absolute():
        path = Path.cwd() / path
    path.mkdir(parents=True, exist_ok=True)
    return path


def datasheet_db_path() -> Path:
    return _data_dir() / DB_NAME


def configured_datasheet_dirs() -> List[dict]:
    raw = os.environ.get(DATASHEET_DIR_ENV, "")
    dirs = []
    for item in [part.strip().strip('"') for part in raw.split(os.pathsep)]:
        if not item:
            continue
        path = Path(item).expanduser()
        dirs.append({
            "path": str(path),
            "exists": path.is_dir(),
        })
    return dirs


def _connect() -> sqlite3.Connection:
    conn = sqlite3.connect(datasheet_db_path())
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    _init_schema(conn)
    return conn


def _init_schema(conn: sqlite3.Connection) -> None:
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS documents (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            path TEXT NOT NULL UNIQUE,
            title TEXT NOT NULL,
            size INTEGER NOT NULL DEFAULT 0,
            mtime REAL NOT NULL DEFAULT 0,
            status TEXT NOT NULL DEFAULT 'pending',
            page_count INTEGER NOT NULL DEFAULT 0,
            extractor TEXT NOT NULL DEFAULT '',
            error TEXT NOT NULL DEFAULT '',
            indexed_at TEXT NOT NULL DEFAULT ''
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS pages (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            doc_id INTEGER NOT NULL,
            page INTEGER NOT NULL,
            text TEXT NOT NULL DEFAULT '',
            FOREIGN KEY(doc_id) REFERENCES documents(id) ON DELETE CASCADE,
            UNIQUE(doc_id, page)
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS chunks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            doc_id INTEGER NOT NULL,
            page INTEGER NOT NULL,
            chunk_id TEXT NOT NULL,
            section_title TEXT NOT NULL DEFAULT '',
            text TEXT NOT NULL DEFAULT '',
            keywords TEXT NOT NULL DEFAULT '',
            char_start INTEGER NOT NULL DEFAULT 0,
            char_end INTEGER NOT NULL DEFAULT 0,
            indexed_at TEXT NOT NULL DEFAULT '',
            FOREIGN KEY(doc_id) REFERENCES documents(id) ON DELETE CASCADE,
            UNIQUE(doc_id, chunk_id)
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS parameters (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            doc_id INTEGER NOT NULL,
            parameter_key TEXT NOT NULL DEFAULT '',
            parameter_name TEXT NOT NULL DEFAULT '',
            value_text TEXT NOT NULL DEFAULT '',
            value_min REAL,
            value_typ REAL,
            value_max REAL,
            unit TEXT NOT NULL DEFAULT '',
            condition TEXT NOT NULL DEFAULT '',
            page INTEGER NOT NULL DEFAULT 1,
            chunk_id TEXT NOT NULL DEFAULT '',
            source_text TEXT NOT NULL DEFAULT '',
            confidence TEXT NOT NULL DEFAULT '',
            extraction_method TEXT NOT NULL DEFAULT '',
            indexed_at TEXT NOT NULL DEFAULT '',
            FOREIGN KEY(doc_id) REFERENCES documents(id) ON DELETE CASCADE
        )
        """
    )
    conn.execute("CREATE INDEX IF NOT EXISTS idx_documents_status ON documents(status)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_pages_doc_page ON pages(doc_id, page)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_chunks_doc_chunk ON chunks(doc_id, chunk_id)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_chunks_doc_page ON chunks(doc_id, page)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_parameters_doc_key ON parameters(doc_id, parameter_key)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_parameters_doc_page ON parameters(doc_id, page)")
    conn.commit()


def _pdf_files(source_dirs: Iterable[dict]) -> List[Path]:
    files: List[Path] = []
    for item in source_dirs:
        path = Path(str(item.get("path") or "")).expanduser()
        if not path.is_dir():
            continue
        files.extend(file for file in path.rglob("*.pdf") if file.is_file() and not file.is_symlink())
        files.extend(file for file in path.rglob("*.PDF") if file.is_file() and not file.is_symlink())
    return sorted(set(files), key=lambda item: item.as_posix().lower())


def _active_source_roots(source_dirs: Optional[Iterable[dict]] = None) -> List[Path]:
    roots: List[Path] = []
    for item in source_dirs if source_dirs is not None else configured_datasheet_dirs():
        path = Path(str(item.get("path") or "")).expanduser()
        if not path.is_dir():
            continue
        try:
            roots.append(path.resolve())
        except OSError:
            roots.append(path.absolute())
    return roots


def _path_in_roots(path: object, roots: Iterable[Path]) -> bool:
    try:
        candidate = Path(str(path or "")).expanduser().resolve(strict=False)
    except OSError:
        candidate = Path(str(path or "")).expanduser().absolute()
    for root in roots:
        try:
            candidate.relative_to(root)
            return True
        except ValueError:
            continue
    return False


def _active_document_rows(conn: sqlite3.Connection, source_dirs: Optional[Iterable[dict]] = None) -> List[sqlite3.Row]:
    roots = _active_source_roots(source_dirs)
    if not roots:
        return []
    rows = conn.execute("SELECT * FROM documents").fetchall()
    return [row for row in rows if _path_in_roots(row["path"], roots)]


def _delete_documents_by_ids(conn: sqlite3.Connection, doc_ids: Iterable[int]) -> int:
    ids = [int(doc_id) for doc_id in doc_ids]
    if not ids:
        return 0
    placeholders = ",".join("?" for _ in ids)
    conn.execute(f"DELETE FROM pages WHERE doc_id IN ({placeholders})", ids)
    conn.execute(f"DELETE FROM chunks WHERE doc_id IN ({placeholders})", ids)
    conn.execute(f"DELETE FROM parameters WHERE doc_id IN ({placeholders})", ids)
    conn.execute(f"DELETE FROM documents WHERE id IN ({placeholders})", ids)
    return len(ids)


def _extract_pdf_pages(path: Path) -> Tuple[str, List[str], str, str]:
    mode = configured_pdf_extractor()
    if mode in {"auto", "mineru"}:
        mineru_status, mineru_pages, mineru_extractor, mineru_error = _extract_pdf_pages_with_mineru(path)
        if mineru_status == "indexed":
            return mineru_status, mineru_pages, mineru_extractor, mineru_error
        if mode == "mineru":
            return mineru_status, mineru_pages, mineru_extractor, mineru_error
        fallback_status, fallback_pages, fallback_extractor, fallback_error = _extract_pdf_pages_with_pypdf(path)
        if fallback_status == "indexed":
            note = f"MinerU 抽取失败，已回退 {fallback_extractor}：{mineru_error}"
            if fallback_error:
                note = f"{note}；{fallback_error}"
            return fallback_status, fallback_pages, fallback_extractor, _safe_text(note, 2000)
        combined_error = f"MinerU 抽取失败：{mineru_error}"
        if fallback_error:
            combined_error = f"{combined_error}；{fallback_extractor} 抽取失败：{fallback_error}"
        return fallback_status, fallback_pages, fallback_extractor, _safe_text(combined_error, 2000)
    return _extract_pdf_pages_with_pypdf(path)


def reindex_datasheets(*, force: bool = False, max_files: int = 5000) -> dict:
    source_dirs = configured_datasheet_dirs()
    configured = bool(source_dirs)
    with _connect() as conn:
        if not configured:
            return {
                "ok": True,
                "configured": False,
                "db_path": str(datasheet_db_path()),
                "source_dirs": [],
                "extractor": {
                    "mode": configured_pdf_extractor(),
                    "mineru": build_mineru_status(include_version=False),
                },
                "indexed_count": 0,
                "skipped_count": 0,
                "failed_count": 0,
                "summary": f"未配置 {DATASHEET_DIR_ENV}，规格书索引不可用。",
            }
        all_files = _pdf_files(source_dirs)
        active_paths = {str(path) for path in all_files}
        stale_ids = [
            int(row["id"])
            for row in conn.execute("SELECT id,path FROM documents").fetchall()
            if str(row["path"] or "") not in active_paths
        ]
        removed_count = _delete_documents_by_ids(conn, stale_ids)
        files = all_files[:max(1, max_files)]
        indexed_count = 0
        skipped_count = 0
        failed_count = 0
        for path in files:
            stat = path.stat()
            existing = conn.execute("SELECT * FROM documents WHERE path=?", (str(path),)).fetchone()
            if existing and not force and int(existing["size"]) == stat.st_size and float(existing["mtime"]) == stat.st_mtime:
                existing_doc_id = int(existing["id"])
                chunk_count = int(conn.execute("SELECT COUNT(*) FROM chunks WHERE doc_id=?", (existing_doc_id,)).fetchone()[0])
                if str(existing["status"] or "") != "indexed" or chunk_count > 0:
                    skipped_count += 1
                    continue
            status, pages, extractor, error = _extract_pdf_pages(path)
            if status != "indexed":
                failed_count += 1
            page_count = len(pages)
            conn.execute(
                """
                INSERT INTO documents(path,title,size,mtime,status,page_count,extractor,error,indexed_at)
                VALUES(?,?,?,?,?,?,?,?,?)
                ON CONFLICT(path) DO UPDATE SET
                    title=excluded.title,
                    size=excluded.size,
                    mtime=excluded.mtime,
                    status=excluded.status,
                    page_count=excluded.page_count,
                    extractor=excluded.extractor,
                    error=excluded.error,
                    indexed_at=excluded.indexed_at
                """,
                (str(path), path.name, stat.st_size, stat.st_mtime, status, page_count, extractor, _safe_text(error, 2000), _now()),
            )
            doc_id = int(conn.execute("SELECT id FROM documents WHERE path=?", (str(path),)).fetchone()["id"])
            conn.execute("DELETE FROM pages WHERE doc_id=?", (doc_id,))
            conn.execute("DELETE FROM chunks WHERE doc_id=?", (doc_id,))
            conn.execute("DELETE FROM parameters WHERE doc_id=?", (doc_id,))
            indexed_at = _now()
            chunks_for_doc = []
            pages_for_parameters = []
            for index, text in enumerate(pages, start=1):
                conn.execute(
                    "INSERT OR REPLACE INTO pages(doc_id,page,text) VALUES(?,?,?)",
                    (doc_id, index, text or ""),
                )
                pages_for_parameters.append({
                    "page": index,
                    "chunk_id": f"p{index}-full",
                    "text": text or "",
                })
                for chunk in _chunk_page_text(text or "", index, indexed_at):
                    chunks_for_doc.append(chunk)
                    conn.execute(
                        """
                        INSERT OR REPLACE INTO chunks(
                            doc_id,page,chunk_id,section_title,text,keywords,char_start,char_end,indexed_at
                        ) VALUES(?,?,?,?,?,?,?,?,?)
                        """,
                        (
                            doc_id,
                            chunk["page"],
                            chunk["chunk_id"],
                            chunk["section_title"],
                            chunk["text"],
                            chunk["keywords"],
                            chunk["char_start"],
                            chunk["char_end"],
                            chunk["indexed_at"],
                        ),
                    )
            # Parameter extraction intentionally uses full page text rather than
            # search chunks.  Complex datasheet tables can be split across chunk
            # boundaries, which would otherwise drop a table row's later cells.
            for parameter in _extract_datasheet_parameters(pages_for_parameters, title=path.name):
                conn.execute(
                    """
                    INSERT INTO parameters(
                        doc_id,parameter_key,parameter_name,value_text,value_min,value_typ,value_max,
                        unit,condition,page,chunk_id,source_text,confidence,extraction_method,indexed_at
                    ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                    """,
                    (
                        doc_id,
                        parameter.get("parameter_key", ""),
                        parameter.get("parameter_name", ""),
                        _safe_text(parameter.get("value_text", ""), 500),
                        parameter.get("value_min"),
                        parameter.get("value_typ"),
                        parameter.get("value_max"),
                        _safe_text(parameter.get("unit", ""), 40),
                        _safe_text(parameter.get("condition", ""), 500),
                        int(parameter.get("page") or 1),
                        _safe_text(parameter.get("chunk_id", ""), 80),
                        _safe_text(parameter.get("source_text", ""), 1200),
                        _safe_text(parameter.get("confidence", ""), 40),
                        _safe_text(parameter.get("extraction_method", ""), 80),
                        indexed_at,
                    ),
                )
            indexed_count += 1
        conn.commit()
        return {
            "ok": True,
            "configured": True,
            "db_path": str(datasheet_db_path()),
            "source_dirs": source_dirs,
            "pdf_count": len(files),
            "indexed_count": indexed_count,
            "skipped_count": skipped_count,
            "failed_count": failed_count,
            "removed_count": removed_count,
            "summary": f"扫描 {len(files)} 个 PDF，更新 {indexed_count} 个，跳过 {skipped_count} 个，失败/需人工 {failed_count} 个，移除过期 {removed_count} 个。",
        }


def build_datasheet_status() -> dict:
    source_dirs = configured_datasheet_dirs()
    with _connect() as conn:
        active_docs = _active_document_rows(conn, source_dirs)
        active_ids = [int(row["id"]) for row in active_docs]
        if active_ids:
            placeholders = ",".join("?" for _ in active_ids)
            page_count = int(conn.execute(f"SELECT COUNT(*) FROM pages WHERE doc_id IN ({placeholders})", active_ids).fetchone()[0])
            chunk_count = int(conn.execute(f"SELECT COUNT(*) FROM chunks WHERE doc_id IN ({placeholders})", active_ids).fetchone()[0])
            parameter_count = int(conn.execute(f"SELECT COUNT(*) FROM parameters WHERE doc_id IN ({placeholders})", active_ids).fetchone()[0])
            section_count = int(conn.execute(f"SELECT COUNT(DISTINCT section_title) FROM chunks WHERE section_title!='' AND doc_id IN ({placeholders})", active_ids).fetchone()[0])
        else:
            page_count = 0
            chunk_count = 0
            parameter_count = 0
            section_count = 0
        doc_count = len(active_docs)
        failed_count = sum(1 for row in active_docs if str(row["status"] or "") != "indexed")
        indexed_count = sum(1 for row in active_docs if str(row["status"] or "") == "indexed")
        last_indexed_at = max((str(row["indexed_at"] or "") for row in active_docs), default="")
        failures = [
            {
                "doc_id": int(row["id"]),
                "title": row["title"],
                "path": row["path"],
                "status": row["status"],
                "error": row["error"],
            }
            for row in sorted(active_docs, key=lambda item: str(item["indexed_at"] or ""), reverse=True)
            if str(row["status"] or "") != "indexed"
        ][:20]
    return {
        "ok": True,
        "configured": bool(source_dirs),
        "source_dirs": source_dirs,
        "db_path": str(datasheet_db_path()),
        "document_count": doc_count,
        "indexed_count": indexed_count,
        "page_count": page_count,
        "chunk_count": chunk_count,
        "parameter_count": parameter_count,
        "section_count": section_count,
        "failed_count": failed_count,
        "last_indexed_at": last_indexed_at,
        "failures": failures,
        "extractor": {
            "mode": configured_pdf_extractor(),
            "mineru": build_mineru_status(),
        },
        "summary": "规格书目录已配置。" if source_dirs else f"未配置 {DATASHEET_DIR_ENV}。",
    }


def _terms(query: str) -> List[str]:
    terms = []
    for term in re.split(r"[\s,;；，/|]+", str(query or "")):
        term = term.strip()
        if len(term) >= 2 and term.lower() not in {"and", "or"} and term not in terms:
            terms.append(term)
    return terms[:12]


def _snippet(text: str, terms: List[str], limit: int = 420) -> str:
    content = re.sub(r"\s+", " ", str(text or "")).strip()
    if not content:
        return ""
    lower = content.lower()
    positions = [lower.find(term.lower()) for term in terms if term and lower.find(term.lower()) >= 0]
    start = max(0, min(positions) - 120) if positions else 0
    snippet = content[start:start + limit]
    return ("…" if start > 0 else "") + (snippet + ("…" if start + limit < len(content) else ""))


def search_datasheets(query: str, *, limit: int = 20, offset: int = 0) -> dict:
    terms = _terms(query)
    if not terms:
        return {"ok": False, "error": "search_datasheets 需要 query。", "matches": []}
    roots = _active_source_roots()
    if not roots:
        return {"ok": True, "query": query, "terms": terms, "total_matches": 0, "limit": limit, "offset": offset, "matches": [], "summary": f"未配置 {DATASHEET_DIR_ENV} 或目录不可用，规格书检索不可用。"}
    with _connect() as conn:
        rows = conn.execute(
            """
            SELECT d.id AS doc_id,d.title,d.path,d.status,d.error,p.page,p.text
            FROM documents d
            LEFT JOIN pages p ON p.doc_id=d.id
            WHERE d.status='indexed'
            ORDER BY d.title,p.page
            """
        ).fetchall()
    matches = []
    for row in rows:
        if not _path_in_roots(row["path"], roots):
            continue
        haystack = f"{row['title']} {row['text'] or ''}".lower()
        matched_terms = [term for term in terms if term.lower() in haystack]
        if not matched_terms:
            continue
        score = len(matched_terms)
        if any(term.lower() in str(row["title"]).lower() for term in terms):
            score += 2
        matches.append({
            "doc_id": int(row["doc_id"]),
            "title": row["title"],
            "path": row["path"],
            "status": row["status"],
            "page": int(row["page"] or 1),
            "score": score,
            "matched_terms": matched_terms,
            "snippet": _snippet(row["text"] or row["title"], matched_terms),
        })
    matches.sort(key=lambda item: (-int(item["score"]), item["title"].lower(), int(item["page"])))
    selected = matches[max(0, offset):max(0, offset) + max(1, limit)]
    return {
        "ok": True,
        "query": query,
        "terms": terms,
        "total_matches": len(matches),
        "limit": limit,
        "offset": offset,
        "matches": selected,
        "summary": f"规格书检索 `{query}` 命中 {len(matches)} 个页级片段。",
    }


def list_datasheet_documents(*, limit: int = 200, offset: int = 0) -> dict:
    limit = max(1, min(int(limit or 200), 1000))
    offset = max(0, int(offset or 0))
    roots = _active_source_roots()
    with _connect() as conn:
        rows = conn.execute(
            """
            SELECT
                d.id,d.title,d.path,d.status,d.error,d.page_count,d.extractor,d.indexed_at,
                (SELECT COUNT(*) FROM chunks c WHERE c.doc_id=d.id) AS chunk_count,
                (SELECT COUNT(*) FROM parameters p WHERE p.doc_id=d.id) AS parameter_count
            FROM documents d
            ORDER BY d.title COLLATE NOCASE
            """
        ).fetchall()
    rows = [row for row in rows if _path_in_roots(row["path"], roots)]
    total = len(rows)
    rows = rows[offset:offset + limit]
    documents = [
        {
            "doc_id": int(row["id"]),
            "title": row["title"],
            "path": row["path"],
            "status": row["status"],
            "error": row["error"],
            "page_count": int(row["page_count"] or 0),
            "chunk_count": int(row["chunk_count"] or 0),
            "parameter_count": int(row["parameter_count"] or 0),
            "extractor": row["extractor"],
            "indexed_at": row["indexed_at"],
        }
        for row in rows
    ]
    return {
        "ok": True,
        "total_documents": total,
        "limit": limit,
        "offset": offset,
        "documents": documents,
        "summary": f"已索引规格书 {total} 个，返回 {len(documents)} 个文档条目。",
    }


def search_datasheet_chunks(query: str, *, limit: int = 20, offset: int = 0) -> dict:
    terms = _terms(query)
    if not terms:
        return {"ok": False, "error": "search_datasheet_chunks 需要 query。", "matches": []}
    limit = max(1, min(int(limit or 20), 100))
    offset = max(0, int(offset or 0))
    roots = _active_source_roots()
    if not roots:
        return {"ok": True, "query": query, "terms": terms, "total_matches": 0, "limit": limit, "offset": offset, "matches": [], "summary": f"未配置 {DATASHEET_DIR_ENV} 或目录不可用，规格书 chunk 检索不可用。"}
    with _connect() as conn:
        rows = conn.execute(
            """
            SELECT
                d.id AS doc_id,d.title,d.path,d.status,d.error,
                c.page,c.chunk_id,c.section_title,c.text,c.keywords,c.char_start,c.char_end
            FROM chunks c
            JOIN documents d ON d.id=c.doc_id
            WHERE d.status='indexed'
            ORDER BY d.title,c.page,c.chunk_id
            """
        ).fetchall()
    matches = []
    for row in rows:
        if not _path_in_roots(row["path"], roots):
            continue
        haystack = f"{row['title']} {row['section_title']} {row['keywords']} {row['text'] or ''}".lower()
        matched_terms = [term for term in terms if term.lower() in haystack]
        if not matched_terms:
            continue
        score = len(matched_terms) * 2
        title_lower = str(row["title"] or "").lower()
        section_lower = str(row["section_title"] or "").lower()
        keyword_lower = str(row["keywords"] or "").lower()
        score += sum(2 for term in terms if term.lower() in title_lower)
        score += sum(1 for term in terms if term.lower() in section_lower or term.lower() in keyword_lower)
        matches.append({
            "doc_id": int(row["doc_id"]),
            "title": row["title"],
            "path": row["path"],
            "status": row["status"],
            "page": int(row["page"] or 1),
            "chunk_id": row["chunk_id"],
            "section_title": row["section_title"],
            "score": score,
            "matched_terms": matched_terms,
            "keywords": row["keywords"],
            "snippet": _snippet(row["text"] or row["title"], matched_terms),
            "char_range": [int(row["char_start"] or 0), int(row["char_end"] or 0)],
        })
    matches.sort(key=lambda item: (-int(item["score"]), item["title"].lower(), int(item["page"]), str(item["chunk_id"])))
    selected = matches[offset:offset + limit]
    return {
        "ok": True,
        "query": query,
        "terms": terms,
        "total_matches": len(matches),
        "limit": limit,
        "offset": offset,
        "matches": selected,
        "summary": f"规格书 chunk 检索 `{query}` 命中 {len(matches)} 个可引用片段。",
    }


def get_datasheet_chunk(doc_id: int, chunk_id: str, *, max_chars: int = 4000) -> dict:
    chunk_id = str(chunk_id or "").strip()
    if not chunk_id:
        return {"ok": False, "error": "get_datasheet_chunk 需要 chunk_id。"}
    with _connect() as conn:
        row = conn.execute(
            """
            SELECT
                d.id AS doc_id,d.title,d.path,d.status,d.error,
                c.page,c.chunk_id,c.section_title,c.text,c.keywords,c.char_start,c.char_end
            FROM chunks c
            JOIN documents d ON d.id=c.doc_id
            WHERE d.id=? AND c.chunk_id=?
            """,
            (int(doc_id), chunk_id),
        ).fetchone()
    if not row:
        return {"ok": False, "error": f"未找到规格书 chunk：doc_id={doc_id} chunk_id={chunk_id}。"}
    if not _path_in_roots(row["path"], _active_source_roots()):
        return {"ok": False, "error": f"规格书 doc_id={doc_id} 不在当前 {DATASHEET_DIR_ENV} 配置目录内。"}
    text = str(row["text"] or "")
    max_chars = max(1, min(int(max_chars or 4000), 12000))
    return {
        "ok": True,
        "doc_id": int(row["doc_id"]),
        "title": row["title"],
        "path": row["path"],
        "status": row["status"],
        "page": int(row["page"] or 1),
        "chunk_id": row["chunk_id"],
        "section_title": row["section_title"],
        "keywords": row["keywords"],
        "char_range": [int(row["char_start"] or 0), int(row["char_end"] or 0)],
        "chars": len(text),
        "truncated": len(text) > max_chars,
        "content": text[:max_chars],
        "summary": f"读取规格书 {row['title']} 第 {int(row['page'] or 1)} 页 chunk {row['chunk_id']}，返回 {min(len(text), max_chars)} 字符。",
    }


def get_datasheet_page_excerpt(doc_id: int, page: int, *, max_chars: int = 2400) -> dict:
    return get_datasheet_excerpt(doc_id, page, max_chars=max_chars)


def batch_search_datasheet_chunks(queries: Iterable[str], *, limit_per_query: int = 8) -> dict:
    query_list = [str(item or "").strip() for item in list(queries or []) if str(item or "").strip()]
    truncated = len(query_list) > 20
    query_list = query_list[:20]
    limit_per_query = max(1, min(int(limit_per_query or 8), 10))
    items = []
    for query in query_list:
        try:
            result = search_datasheet_chunks(query, limit=limit_per_query)
            matches = result.get("matches", []) or []
            items.append({
                "query": query,
                "status": "found" if matches else "missing",
                "total_matches": int(result.get("total_matches") or 0),
                "matches": matches,
                "missing_reason": "" if matches else "未命中本地规格书 chunk 索引。",
            })
        except Exception as exc:
            items.append({
                "query": query,
                "status": "error",
                "total_matches": 0,
                "matches": [],
                "error": str(exc),
            })
    return {
        "ok": True,
        "query_count": len(query_list),
        "limit_per_query": limit_per_query,
        "truncated": truncated,
        "items": items,
        "summary": f"批量规格书 chunk 检索 {len(query_list)} 项，命中 {sum(1 for item in items if item.get('status') == 'found')} 项。",
    }


def _parameter_row_to_dict(row: sqlite3.Row, *, max_source_chars: int = 520) -> dict:
    source_text = str(row["source_text"] or "")
    max_source_chars = max(80, min(int(max_source_chars or 520), 12000))
    parameter_id = int(row["id"])
    doc_id = int(row["doc_id"])
    chunk_id = str(row["chunk_id"] or "")
    return {
        "parameter_id": parameter_id,
        "evidence_id": f"datasheet-param-{parameter_id}",
        "doc_id": doc_id,
        "title": row["title"],
        "path": row["path"],
        "parameter_key": row["parameter_key"],
        "parameter_name": row["parameter_name"],
        "value_text": row["value_text"],
        "value_min": row["value_min"],
        "value_typ": row["value_typ"],
        "value_max": row["value_max"],
        "unit": row["unit"],
        "condition": row["condition"],
        "page": int(row["page"] or 1),
        "chunk_id": chunk_id,
        "source_text": source_text[:max_source_chars],
        "source_truncated": len(source_text) > max_source_chars,
        "confidence": row["confidence"],
        "extraction_method": row["extraction_method"],
        "indexed_at": row["indexed_at"],
        "detail_locator": {
            "doc_id": doc_id,
            "parameter_id": parameter_id,
            "page": int(row["page"] or 1),
            "chunk_id": chunk_id,
        },
    }


def search_datasheet_parameters(query: str = "",
                                *,
                                parameter_key: str = "",
                                doc_id: Optional[int] = None,
                                limit: int = 50,
                                offset: int = 0) -> dict:
    """Search deterministic datasheet parameter cards."""

    limit = max(1, min(int(limit or 50), 200))
    offset = max(0, int(offset or 0))
    query = str(query or "").strip()
    parameter_key = str(parameter_key or "").strip()
    roots = _active_source_roots()
    if not roots:
        return {
            "ok": True,
            "query": query,
            "parameter_key": parameter_key,
            "total_matches": 0,
            "limit": limit,
            "offset": offset,
            "parameters": [],
            "summary": f"未配置 {DATASHEET_DIR_ENV} 或目录不可用，规格书参数检索不可用。",
        }
    where = ["d.status='indexed'"]
    params: List[object] = []
    if doc_id:
        where.append("d.id=?")
        params.append(int(doc_id))
    sql = f"""
        SELECT
            p.*, d.title, d.path, d.status
        FROM parameters p
        JOIN documents d ON d.id=p.doc_id
        WHERE {' AND '.join(where)}
        ORDER BY d.title COLLATE NOCASE,p.page,p.parameter_key,p.id
    """
    with _connect() as conn:
        rows = conn.execute(sql, params).fetchall()

    query_terms = _terms(query)
    key_terms = _terms(parameter_key)
    matches = []
    for row in rows:
        if not _path_in_roots(row["path"], roots):
            continue
        haystack = (
            f"{row['title']} {row['parameter_key']} {row['parameter_name']} "
            f"{row['value_text']} {row['unit']} {row['condition']} {row['source_text']}"
        ).lower()
        matched_terms = [term for term in query_terms if term.lower() in haystack]
        key_matched_terms = [
            term for term in key_terms
            if term.lower() in str(row["parameter_key"] or "").lower()
            or term.lower() in str(row["parameter_name"] or "").lower()
        ]
        if query_terms and not matched_terms:
            continue
        if key_terms and not key_matched_terms:
            continue
        score = len(matched_terms) + len(key_matched_terms) * 3
        if query and query.lower() in str(row["parameter_name"] or "").lower():
            score += 3
        if parameter_key and parameter_key.lower() == str(row["parameter_key"] or "").lower():
            score += 6
        item = _parameter_row_to_dict(row, max_source_chars=520)
        item["score"] = score
        item["matched_terms"] = matched_terms + key_matched_terms
        matches.append(item)

    matches.sort(key=lambda item: (-int(item.get("score") or 0), str(item.get("title") or "").lower(), int(item.get("page") or 1), str(item.get("parameter_key") or "")))
    selected = matches[offset:offset + limit]
    return {
        "ok": True,
        "query": query,
        "parameter_key": parameter_key,
        "doc_id": int(doc_id) if doc_id else None,
        "total_matches": len(matches),
        "limit": limit,
        "offset": offset,
        "parameters": selected,
        "summary": f"规格书参数检索 `{query or parameter_key or '全部'}` 命中 {len(matches)} 张参数卡。",
    }


def get_datasheet_parameter(parameter_id: int, *, max_chars: int = 2400) -> dict:
    """Read one deterministic datasheet parameter card with source evidence."""

    try:
        parameter_id = int(parameter_id)
    except (TypeError, ValueError):
        return {"ok": False, "error": "get_datasheet_parameter 需要正整数 parameter_id。"}
    if parameter_id <= 0:
        return {"ok": False, "error": "get_datasheet_parameter 需要正整数 parameter_id。"}
    with _connect() as conn:
        row = conn.execute(
            """
            SELECT p.*, d.title, d.path, d.status
            FROM parameters p
            JOIN documents d ON d.id=p.doc_id
            WHERE p.id=?
            """,
            (parameter_id,),
        ).fetchone()
    if not row:
        return {"ok": False, "error": f"未找到规格书参数卡：parameter_id={parameter_id}。"}
    if not _path_in_roots(row["path"], _active_source_roots()):
        return {"ok": False, "error": f"规格书参数卡 parameter_id={parameter_id} 不在当前 {DATASHEET_DIR_ENV} 配置目录内。"}
    card = _parameter_row_to_dict(row, max_source_chars=max_chars)
    return {
        "ok": True,
        **card,
        "summary": (
            f"{card['title']} 第 {card['page']} 页参数 {card['parameter_name']}="
            f"{card['value_text']}（{card['condition'] or '无额外条件'}）。"
        ),
    }


def get_datasheet_excerpt(doc_id: int, page: int, *, max_chars: int = 2400) -> dict:
    with _connect() as conn:
        row = conn.execute(
            """
            SELECT d.id AS doc_id,d.title,d.path,d.status,d.error,p.page,p.text
            FROM documents d
            LEFT JOIN pages p ON p.doc_id=d.id AND p.page=?
            WHERE d.id=?
            """,
            (int(page), int(doc_id)),
        ).fetchone()
    if not row:
        return {"ok": False, "error": f"未找到规格书 doc_id={doc_id} page={page}。"}
    if not _path_in_roots(row["path"], _active_source_roots()):
        return {"ok": False, "error": f"规格书 doc_id={doc_id} 不在当前 {DATASHEET_DIR_ENV} 配置目录内。"}
    text = str(row["text"] or "")
    max_chars = max(1, min(int(max_chars or 2400), 12000))
    return {
        "ok": True,
        "doc_id": int(row["doc_id"]),
        "title": row["title"],
        "path": row["path"],
        "status": row["status"],
        "page": int(row["page"] or page),
        "chars": len(text),
        "truncated": len(text) > max_chars,
        "content": text[:max_chars],
        "summary": f"读取规格书 {row['title']} 第 {int(row['page'] or page)} 页，返回 {min(len(text), max_chars)} 字符。",
    }


def match_component_datasheets(card: dict, *, limit: int = 5) -> dict:
    query_parts = [
        card.get("hq_no", ""),
        card.get("spec", ""),
        card.get("candidate_chip_type", ""),
        card.get("refdes", ""),
    ]
    query = " ".join(str(item).strip() for item in query_parts if str(item or "").strip())
    if not query:
        return {
            "ok": True,
            "refdes": card.get("refdes", ""),
            "query": "",
            "matches": [],
            "missing_reason": "identity_card 缺少 HQ 料号、规格和候选芯片类型。",
        }
    result = search_datasheet_chunks(query, limit=limit)
    return {
        "ok": bool(result.get("ok", True)),
        "refdes": card.get("refdes", ""),
        "query": query,
        "matches": result.get("matches", []),
        "missing_reason": "" if result.get("matches") else "未命中本地规格书索引。",
    }


def summarize_datasheet_coverage(cards: Iterable[dict], *, limit: int = 12) -> dict:
    key_cards = [
        card for card in list(cards or [])
        if str(card.get("category") or "") in {"chip", "power_ic", "large_ic", "connector"}
    ]
    matched = []
    gaps = []
    for card in key_cards:
        match = match_component_datasheets(card, limit=3)
        preview = {
            "refdes": card.get("refdes", ""),
            "category": card.get("category", ""),
            "hq_no": card.get("hq_no", ""),
            "spec": card.get("spec", ""),
            "query": match.get("query", ""),
            "matches": match.get("matches", [])[:3],
            "missing_reason": match.get("missing_reason", ""),
        }
        if match.get("matches"):
            matched.append(preview)
        else:
            gaps.append(preview)
    return {
        "total_key_components": len(key_cards),
        "matched_count": len(matched),
        "gap_count": len(gaps),
        "matched_cards": matched[:limit],
        "gap_cards": gaps[:limit],
        "summary": f"{len(key_cards)} 个关键器件中 {len(matched)} 个命中规格书，{len(gaps)} 个缺规格书证据。",
    }
