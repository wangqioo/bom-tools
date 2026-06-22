# -*- coding: utf-8 -*-
"""Manufacturer naming alias database."""

import os
import re
import sqlite3
import time
import unicodedata
import uuid

from flask import Blueprint

from shared import jsonify, request


manufacturer_alias_bp = Blueprint('manufacturer_alias', __name__)

BASE_DIR = os.path.dirname(os.path.dirname(__file__))
DATA_DIR = os.environ.get('BOM_TOOLS_MANUFACTURER_ALIAS_DIR') or os.path.join(BASE_DIR, 'manufacturer_aliases')
DB_PATH = os.path.join(DATA_DIR, 'aliases.sqlite3')


def normalize_manufacturer_name(value):
    text = unicodedata.normalize('NFKC', str(value or '')).casefold().strip()
    return re.sub(r'[\s\-_.,，。/\\()（）\[\]【】{}+&＆]+', '', text)


def _ensure_dirs():
    os.makedirs(DATA_DIR, exist_ok=True)


def _connect():
    _ensure_dirs()
    conn = sqlite3.connect(DB_PATH, timeout=15)
    conn.row_factory = sqlite3.Row
    conn.execute('PRAGMA journal_mode=WAL')
    conn.execute('PRAGMA busy_timeout=15000')
    conn.execute('''
        CREATE TABLE IF NOT EXISTS manufacturer_aliases (
            id TEXT PRIMARY KEY,
            canonical_name TEXT NOT NULL,
            alias TEXT NOT NULL,
            normalized_alias TEXT NOT NULL UNIQUE,
            source TEXT NOT NULL,
            note TEXT NOT NULL,
            created_at REAL NOT NULL,
            updated_at REAL NOT NULL
        )
    ''')
    conn.commit()
    return conn


def _row_to_alias(row):
    return dict(row)


def lookup_manufacturer(name):
    normalized = normalize_manufacturer_name(name)
    if not normalized:
        return None
    conn = _connect()
    try:
        row = conn.execute(
            'SELECT * FROM manufacturer_aliases WHERE normalized_alias = ?',
            (normalized,),
        ).fetchone()
        return _row_to_alias(row) if row else None
    finally:
        conn.close()


def _list_aliases(query=''):
    q = str(query or '').strip()
    normalized = normalize_manufacturer_name(q)
    conn = _connect()
    try:
        if q:
            rows = conn.execute('''
                SELECT * FROM manufacturer_aliases
                WHERE normalized_alias = ?
                   OR canonical_name LIKE ?
                   OR alias LIKE ?
                   OR source LIKE ?
                   OR note LIKE ?
                ORDER BY canonical_name COLLATE NOCASE, alias COLLATE NOCASE
                LIMIT 300
            ''', (normalized, f'%{q}%', f'%{q}%', f'%{q}%', f'%{q}%')).fetchall()
        else:
            rows = conn.execute('''
                SELECT * FROM manufacturer_aliases
                ORDER BY updated_at DESC
                LIMIT 300
            ''').fetchall()
        return [_row_to_alias(row) for row in rows]
    finally:
        conn.close()


def _clean_form(name, max_len=500):
    return str(request.form.get(name, '') or '').strip()[:max_len]


@manufacturer_alias_bp.route('/api/manufacturer_aliases', methods=['GET'])
def api_list_manufacturer_aliases():
    query = request.args.get('q', '')
    match = lookup_manufacturer(query) if query else None
    return jsonify({
        'success': True,
        'query': query,
        'normalized_query': normalize_manufacturer_name(query),
        'match': match,
        'aliases': _list_aliases(query),
    })


@manufacturer_alias_bp.route('/api/manufacturer_aliases/lookup', methods=['GET'])
def api_lookup_manufacturer_alias():
    name = request.args.get('name', '')
    return jsonify({
        'success': True,
        'query': name,
        'normalized_query': normalize_manufacturer_name(name),
        'match': lookup_manufacturer(name),
    })


@manufacturer_alias_bp.route('/api/manufacturer_aliases', methods=['POST'])
def api_create_manufacturer_alias():
    canonical_name = _clean_form('canonical_name', 200)
    alias = _clean_form('alias', 200)
    source = _clean_form('source', 120)
    note = _clean_form('note', 1000)
    if not canonical_name or not alias:
        return jsonify({'success': False, 'error': '请填写 HQ 规范厂商名和厂商别名'})

    normalized = normalize_manufacturer_name(alias)
    if not normalized:
        return jsonify({'success': False, 'error': '厂商别名无有效字符'})

    item_id = time.strftime('%Y%m%d%H%M%S') + '-' + uuid.uuid4().hex[:6]
    now = time.time()
    conn = _connect()
    try:
        try:
            conn.execute('''
                INSERT INTO manufacturer_aliases (
                    id, canonical_name, alias, normalized_alias, source, note, created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                item_id, canonical_name, alias, normalized,
                source or '手工维护', note, now, now,
            ))
            conn.commit()
        except sqlite3.IntegrityError:
            row = conn.execute(
                'SELECT * FROM manufacturer_aliases WHERE normalized_alias = ?',
                (normalized,),
            ).fetchone()
            return jsonify({
                'success': False,
                'error': '该别名已存在',
                'existing': _row_to_alias(row) if row else None,
            })
        row = conn.execute('SELECT * FROM manufacturer_aliases WHERE id = ?', (item_id,)).fetchone()
        return jsonify({'success': True, 'alias': _row_to_alias(row)})
    finally:
        conn.close()


@manufacturer_alias_bp.route('/api/manufacturer_aliases/<alias_id>', methods=['DELETE'])
def api_delete_manufacturer_alias(alias_id):
    conn = _connect()
    try:
        cur = conn.execute('DELETE FROM manufacturer_aliases WHERE id = ?', (alias_id,))
        conn.commit()
        if cur.rowcount <= 0:
            return jsonify({'success': False, 'error': '记录不存在'})
        return jsonify({'success': True})
    finally:
        conn.close()
