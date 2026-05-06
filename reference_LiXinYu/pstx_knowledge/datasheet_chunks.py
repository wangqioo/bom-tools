# -*- coding: utf-8 -*-
"""Chunk builders for datasheet evidence indexing."""

from __future__ import annotations

import re
from typing import List


def normalize_page_text(text: str) -> str:
    content = str(text or "").replace("\r\n", "\n").replace("\r", "\n")
    content = re.sub(r"[ \t]+\n", "\n", content)
    content = re.sub(r"\n{3,}", "\n\n", content)
    return content.strip()


def section_title(text: str) -> str:
    for line in str(text or "").splitlines():
        candidate = re.sub(r"\s+", " ", line).strip(" :-\t")
        if not candidate:
            continue
        if len(candidate) <= 96:
            return candidate
    return ""


def chunk_keywords(text: str, limit: int = 18) -> str:
    candidates = re.findall(r"HQ[0-9A-Za-z]+|[A-Za-z][A-Za-z0-9_./+-]{2,}|[\u4e00-\u9fff]{2,}", str(text or ""))
    seen = []
    for item in candidates:
        token = item.strip()
        key = token.lower()
        if key and key not in seen:
            seen.append(key)
    return " ".join(seen[:limit])


def chunk_page_text(page_text: str, page: int, indexed_at: str, *, chunk_chars: int = 1600, overlap: int = 180) -> List[dict]:
    text = normalize_page_text(page_text)
    if not text:
        return []
    chunk_chars = max(400, min(int(chunk_chars or 1600), 4000))
    overlap = max(0, min(int(overlap or 0), chunk_chars // 3))
    chunks = []
    start = 0
    index = 1
    while start < len(text):
        end = min(len(text), start + chunk_chars)
        if end < len(text):
            soft_break = max(text.rfind("\n\n", start, end), text.rfind("\n", start, end), text.rfind("。", start, end))
            if soft_break > start + chunk_chars // 2:
                end = soft_break + 1
        chunk_text = text[start:end].strip()
        if chunk_text:
            chunks.append({
                "page": int(page),
                "chunk_id": f"p{int(page)}-c{index}",
                "section_title": section_title(chunk_text),
                "text": chunk_text,
                "keywords": chunk_keywords(chunk_text),
                "char_start": start,
                "char_end": end,
                "indexed_at": indexed_at,
            })
            index += 1
        if end >= len(text):
            break
        start = max(end - overlap, start + 1)
    return chunks
