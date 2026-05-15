# -*- coding: utf-8 -*-
"""Feature request work order blueprint."""

import json
import os
import sqlite3
import time
import uuid

from flask import Blueprint, abort, send_file
from werkzeug.utils import safe_join, secure_filename

from shared import jsonify, request


feature_request_bp = Blueprint('feature_request', __name__)

BASE_DIR = os.path.dirname(os.path.dirname(__file__))
REQ_DIR = os.path.join(BASE_DIR, 'feature_requests')
ATTACH_DIR = os.path.join(REQ_DIR, 'attachments')
DB_PATH = os.path.join(REQ_DIR, 'requests.sqlite3')
ALLOWED_ATTACHMENT_EXTS = {
    '.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp',
    '.xlsx', '.xlsm', '.xls', '.csv', '.txt', '.log', '.zip', '.rar', '.7z',
    '.doc', '.docx', '.ppt', '.pptx', '.pdf',
}

SEED_REQUESTS = [
    {
        'id': 'seed-bom-compare-summary',
        'requester': '\u7cfb\u7edf\u793a\u4f8b',
        'employee_id': 'SYSTEM',
        'module': 'BOM\u6bd4\u5bf9\u5de5\u5177\u5408\u96c6',
        'priority': '\u8f83\u9ad8',
        'request_type': '\u529f\u80fd\u4f18\u5316',
        'status': '\u5f85\u8bc4\u4f30',
        'title': 'BOM\u6bd4\u5bf9\u62a5\u544a\u589e\u52a0\u5dee\u5f02\u6982\u89c8\u9875',
        'background': '\u5bf9\u6bd4\u7ed3\u679c\u884c\u6570\u8f83\u591a\u65f6\uff0c\u9700\u8981\u5148\u770b\u603b\u4f53\u5dee\u5f02\u89c4\u6a21\u3002',
        'requirement': '\u5bfc\u51fa Excel \u65f6\u5728\u9996\u9875\u589e\u52a0\u603b\u89c8\uff0c\u7edf\u8ba1\u65b0\u589e\u3001\u5220\u9664\u3001\u53d8\u66f4\u3001\u91cd\u590d\u952e\u6570\u91cf\u3002',
        'value': '\u51cf\u5c11\u4eba\u5de5\u7edf\u8ba1\u65f6\u95f4\uff0c\u4fbf\u4e8e\u5feb\u901f\u5224\u65ad\u53d8\u66f4\u98ce\u9669\u3002',
        'acceptance': '\u5dee\u5f02\u62a5\u544a\u5305\u542b\u603b\u89c8 sheet\uff0c\u5e76\u80fd\u8df3\u8f6c\u5230\u660e\u7ec6 sheet\u3002',
        'likes': 6,
    },
    {
        'id': 'seed-feishu-cache-refresh',
        'requester': '\u7cfb\u7edf\u793a\u4f8b',
        'employee_id': 'SYSTEM',
        'module': '\u98de\u4e66\u4f18\u9009\u5e93+\u5173\u7cfb\u5e93\u5339\u914d',
        'priority': '\u4e00\u822c',
        'request_type': '\u81ea\u52a8\u5316',
        'status': '\u5f85\u8bc4\u4f30',
        'title': '\u652f\u6301\u5b9a\u65f6\u5237\u65b0\u98de\u4e66\u8868\u683c\u7f13\u5b58',
        'background': '\u76ee\u524d\u9700\u8981\u624b\u52a8\u5237\u65b0\u7f13\u5b58\uff0c\u5bb9\u6613\u4f7f\u7528\u5230\u65e7\u6570\u636e\u3002',
        'requirement': '\u53ef\u914d\u7f6e\u6bcf\u5929\u56fa\u5b9a\u65f6\u95f4\u81ea\u52a8\u5237\u65b0\u5df2\u542f\u7528\u7684\u98de\u4e66 sheet \u7f13\u5b58\u3002',
        'value': '\u964d\u4f4e\u5339\u914d\u7ed3\u679c\u8fc7\u671f\u98ce\u9669\uff0c\u51cf\u5c11\u91cd\u590d\u64cd\u4f5c\u3002',
        'acceptance': '\u9875\u9762\u663e\u793a\u6700\u540e\u81ea\u52a8\u5237\u65b0\u65f6\u95f4\uff0c\u5237\u65b0\u5931\u8d25\u6709\u9519\u8bef\u63d0\u793a\u3002',
        'likes': 4,
    },
    {
        'id': 'seed-plm-template-check',
        'requester': '\u7cfb\u7edf\u793a\u4f8b',
        'employee_id': 'SYSTEM',
        'module': '\u8f6c\u6362\u4e3a\u4e0a\u4f20PLM\u7cfb\u7edf\u683c\u5f0f',
        'priority': '\u4e00\u822c',
        'request_type': '\u6d41\u7a0b\u6539\u8fdb',
        'status': '\u5f85\u8bc4\u4f30',
        'title': 'PLM\u5bfc\u5165\u524d\u589e\u52a0\u5fc5\u586b\u9879\u6821\u9a8c',
        'background': '\u90e8\u5206\u5bfc\u51fa\u6587\u4ef6\u4e0a\u4f20 PLM \u540e\u624d\u53d1\u73b0\u5b57\u6bb5\u7f3a\u5931\u3002',
        'requirement': '\u8f6c\u6362\u5b8c\u6210\u524d\u68c0\u67e5\u7269\u6599\u53f7\u3001\u5355\u8017\u3001\u4e3b\u8f85 BOM \u6807\u8bb0\u7b49\u5173\u952e\u5b57\u6bb5\u662f\u5426\u4e3a\u7a7a\u3002',
        'value': '\u63d0\u524d\u66b4\u9732\u6570\u636e\u95ee\u9898\uff0c\u51cf\u5c11 PLM \u5bfc\u5165\u8fd4\u5de5\u3002',
        'acceptance': '\u5b58\u5728\u5fc5\u586b\u7f3a\u5931\u65f6\u8f93\u51fa\u660e\u7ec6\u65e5\u5fd7\uff0c\u660e\u786e\u884c\u53f7\u548c\u5b57\u6bb5\u540d\u3002',
        'likes': 2,
    },
]


def _ensure_dirs():
    os.makedirs(ATTACH_DIR, exist_ok=True)


def _connect():
    _ensure_dirs()
    conn = sqlite3.connect(DB_PATH, timeout=15)
    conn.row_factory = sqlite3.Row
    conn.execute('PRAGMA journal_mode=WAL')
    conn.execute('PRAGMA busy_timeout=15000')
    conn.execute('''
        CREATE TABLE IF NOT EXISTS feature_requests (
            id TEXT PRIMARY KEY,
            submitted_at REAL NOT NULL,
            requester TEXT NOT NULL,
            employee_id TEXT NOT NULL,
            module TEXT NOT NULL,
            priority TEXT NOT NULL,
            request_type TEXT NOT NULL,
            status TEXT NOT NULL,
            title TEXT NOT NULL,
            background TEXT NOT NULL,
            requirement TEXT NOT NULL,
            value TEXT NOT NULL,
            acceptance TEXT NOT NULL,
            attachments TEXT NOT NULL,
            likes INTEGER NOT NULL DEFAULT 0
        )
    ''')
    _ensure_schema(conn)
    _seed_requests(conn)
    conn.commit()
    return conn


def _ensure_schema(conn):
    cols = {row['name'] for row in conn.execute('PRAGMA table_info(feature_requests)').fetchall()}
    if 'likes' not in cols:
        conn.execute('ALTER TABLE feature_requests ADD COLUMN likes INTEGER NOT NULL DEFAULT 0')


def _seed_requests(conn):
    base_time = time.time() - 86400 * len(SEED_REQUESTS)
    for idx, item in enumerate(SEED_REQUESTS):
        conn.execute('''
            INSERT OR IGNORE INTO feature_requests (
                id, submitted_at, requester, employee_id, module, priority, request_type,
                status, title, background, requirement, value, acceptance, attachments, likes
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            item['id'], base_time + idx * 3600, item['requester'], item['employee_id'],
            item['module'], item['priority'], item['request_type'], item['status'],
            item['title'], item['background'], item['requirement'], item['value'],
            item['acceptance'], '[]', item['likes'],
        ))


def _row_to_request(row):
    item = dict(row)
    try:
        item['attachments'] = json.loads(item.get('attachments') or '[]')
    except Exception:
        item['attachments'] = []
    return item


def _read_requests():
    conn = _connect()
    try:
        rows = conn.execute('SELECT * FROM feature_requests ORDER BY submitted_at DESC').fetchall()
        return [_row_to_request(row) for row in rows]
    finally:
        conn.close()


def _insert_request(item):
    conn = _connect()
    try:
        conn.execute('''
            INSERT INTO feature_requests (
                id, submitted_at, requester, employee_id, module, priority, request_type,
                status, title, background, requirement, value, acceptance, attachments, likes
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            item['id'], item['submitted_at'], item['requester'], item['employee_id'],
            item['module'], item['priority'], item['request_type'], item['status'],
            item['title'], item['background'], item['requirement'], item['value'],
            item['acceptance'], json.dumps(item['attachments'], ensure_ascii=False),
            item.get('likes', 0),
        ))
        conn.commit()
    finally:
        conn.close()


def _clean_text(name, max_len=2000):
    return str(request.form.get(name, '') or '').strip()[:max_len]


def _save_attachments(request_id):
    _ensure_dirs()
    saved = []
    for file in request.files.getlist('attachments'):
        if not file or not file.filename:
            continue
        _, ext = os.path.splitext(file.filename.lower())
        if ext not in ALLOWED_ATTACHMENT_EXTS:
            raise ValueError('\u4ec5\u652f\u6301\u4e0a\u4f20\u56fe\u7247\u3001Excel\u3001CSV\u3001Office/PDF\u3001\u65e5\u5fd7\u6587\u672c\u6216\u538b\u7f29\u5305\u9644\u4ef6')
        safe_name = secure_filename(file.filename) or f'attachment{ext}'
        out_name = f'{request_id}_{uuid.uuid4().hex[:8]}_{safe_name}'
        out_path = os.path.join(ATTACH_DIR, out_name)
        file.save(out_path)
        saved.append({
            'name': file.filename,
            'url': f'/feature_attachments/{out_name}',
        })
    return saved


@feature_request_bp.route('/api/feature_requests', methods=['GET'])
def api_feature_requests():
    return jsonify({'success': True, 'requests': _read_requests()})


@feature_request_bp.route('/api/feature_requests', methods=['POST'])
def api_submit_feature_request():
    requester = _clean_text('requester', 80)
    employee_id = _clean_text('employee_id', 40)
    title = _clean_text('title', 120)
    requirement = _clean_text('requirement', 5000)
    if not requester or not employee_id or not title or not requirement:
        return jsonify({'success': False, 'error': '\u8bf7\u586b\u5199\u59d3\u540d\u3001\u5de5\u53f7\u3001\u9700\u6c42\u6807\u9898\u548c\u9700\u6c42\u8bf4\u660e'})

    request_id = time.strftime('%Y%m%d%H%M%S') + '-' + uuid.uuid4().hex[:6]
    try:
        attachments = _save_attachments(request_id)
    except ValueError as e:
        return jsonify({'success': False, 'error': str(e)})

    item = {
        'id': request_id,
        'submitted_at': time.time(),
        'requester': requester,
        'employee_id': employee_id,
        'module': _clean_text('module', 80) or '\u672a\u6307\u5b9a',
        'priority': _clean_text('priority', 40) or '\u4e00\u822c',
        'request_type': _clean_text('request_type', 40) or '\u65b0\u529f\u80fd',
        'status': '\u5f85\u8bc4\u4f30',
        'title': title,
        'background': _clean_text('background', 4000),
        'requirement': requirement,
        'value': _clean_text('value', 3000),
        'acceptance': _clean_text('acceptance', 3000),
        'attachments': attachments,
        'likes': 0,
    }
    _insert_request(item)
    return jsonify({'success': True, 'request': item})


@feature_request_bp.route('/api/feature_requests/<request_id>/like', methods=['POST'])
def api_like_feature_request(request_id):
    conn = _connect()
    try:
        conn.execute('UPDATE feature_requests SET likes = likes + 1 WHERE id = ?', (request_id,))
        row = conn.execute('SELECT * FROM feature_requests WHERE id = ?', (request_id,)).fetchone()
        if not row:
            return jsonify({'success': False, 'error': '\u9700\u6c42\u4e0d\u5b58\u5728'})
        conn.commit()
        return jsonify({'success': True, 'request': _row_to_request(row)})
    finally:
        conn.close()


@feature_request_bp.route('/feature_attachments/<filename>')
def feature_attachment(filename):
    path = safe_join(ATTACH_DIR, filename)
    if not path or not os.path.exists(path):
        abort(404)
    return send_file(path)
