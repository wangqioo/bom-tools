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
from auth import record_activity, require_admin_json


feature_request_bp = Blueprint('feature_request', __name__)

BASE_DIR = os.path.dirname(os.path.dirname(__file__))
REQ_DIR = os.environ.get('BOM_TOOLS_FEATURE_REQUEST_DIR') or os.path.join(BASE_DIR, 'feature_requests')
ATTACH_DIR = os.path.join(REQ_DIR, 'attachments')
DB_PATH = os.path.join(REQ_DIR, 'requests.sqlite3')
ALLOWED_ATTACHMENT_EXTS = {
    '.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp',
    '.xlsx', '.xlsm', '.xls', '.csv', '.txt', '.log', '.zip', '.rar', '.7z',
    '.doc', '.docx', '.ppt', '.pptx', '.pdf',
}
ALLOWED_STATUSES = {'\u5f85\u8bc4\u4f30', '\u5df2\u7eb3\u5165', '\u5f00\u53d1\u4e2d', '\u5df2\u5b8c\u6210', '\u6682\u7f13', '\u5df2\u5173\u95ed'}

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
    conn.execute('''
        CREATE TABLE IF NOT EXISTS feature_request_likes (
            request_id TEXT NOT NULL,
            employee_id TEXT NOT NULL,
            liked_at REAL NOT NULL,
            PRIMARY KEY (request_id, employee_id)
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
    keep_ids = {item['id'] for item in SEED_REQUESTS}
    placeholders = ','.join('?' for _ in keep_ids)
    if keep_ids:
        conn.execute("DELETE FROM feature_requests WHERE id LIKE 'seed-%' AND id NOT IN (" + placeholders + ")", tuple(keep_ids))
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


def _filter_requests(items, status='', module='', query=''):
    status = str(status or '').strip()
    module = str(module or '').strip()
    query = str(query or '').strip().lower()
    filtered = []
    for item in items:
        if status and item.get('status') != status:
            continue
        if module and item.get('module') != module:
            continue
        if query:
            haystack = '\n'.join(str(item.get(key, '') or '') for key in (
                'title', 'background', 'requirement', 'value', 'acceptance',
                'requester', 'employee_id', 'module', 'priority', 'request_type',
            )).lower()
            if query not in haystack:
                continue
        filtered.append(item)
    return filtered


def _sort_requests(items, sort='newest'):
    sort = str(sort or 'newest').strip()
    if sort == 'likes':
        return sorted(items, key=lambda item: (-int(item.get('likes') or 0), -float(item.get('submitted_at') or 0)))
    return sorted(items, key=lambda item: -float(item.get('submitted_at') or 0))


def _read_requests(status='', module='', query='', sort='newest'):
    conn = _connect()
    try:
        rows = conn.execute('SELECT * FROM feature_requests ORDER BY submitted_at DESC').fetchall()
        items = [_row_to_request(row) for row in rows]
        items = _filter_requests(items, status=status, module=module, query=query)
        return _sort_requests(items, sort=sort)
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


def _update_request_status(request_id, status):
    conn = _connect()
    try:
        conn.execute('UPDATE feature_requests SET status = ? WHERE id = ?', (status, request_id))
        row = conn.execute('SELECT * FROM feature_requests WHERE id = ?', (request_id,)).fetchone()
        if not row:
            return None
        conn.commit()
        return _row_to_request(row)
    finally:
        conn.close()


def _like_request(request_id, employee_id=''):
    conn = _connect()
    try:
        row = conn.execute('SELECT * FROM feature_requests WHERE id = ?', (request_id,)).fetchone()
        if not row:
            return None, False

        employee_id = str(employee_id or '').strip()[:40]
        if not employee_id:
            conn.execute('UPDATE feature_requests SET likes = likes + 1 WHERE id = ?', (request_id,))
            row = conn.execute('SELECT * FROM feature_requests WHERE id = ?', (request_id,)).fetchone()
            conn.commit()
            return _row_to_request(row), True

        try:
            conn.execute(
                'INSERT INTO feature_request_likes (request_id, employee_id, liked_at) VALUES (?, ?, ?)',
                (request_id, employee_id, time.time()),
            )
        except sqlite3.IntegrityError:
            return _row_to_request(row), False

        conn.execute('UPDATE feature_requests SET likes = likes + 1 WHERE id = ?', (request_id,))
        row = conn.execute('SELECT * FROM feature_requests WHERE id = ?', (request_id,)).fetchone()
        conn.commit()
        return _row_to_request(row), True
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
    items = _read_requests(
        status=request.args.get('status', ''),
        module=request.args.get('module', ''),
        query=request.args.get('q', ''),
        sort=request.args.get('sort', 'newest'),
    )
    return jsonify({'success': True, 'requests': items, 'total': len(items)})


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
    data = request.get_json(silent=True) or {}
    employee_id = str(data.get('employee_id') or request.form.get('employee_id') or request.args.get('employee_id') or '').strip()
    item, liked = _like_request(request_id, employee_id)
    if not item:
        return jsonify({'success': False, 'error': '\u9700\u6c42\u4e0d\u5b58\u5728'})
    if employee_id and not liked:
        return jsonify({'success': True, 'request': item, 'already_liked': True, 'message': '\u4f60\u5df2\u7ecf\u70b9\u8d5e\u8fc7\u8be5\u9700\u6c42'})
    return jsonify({'success': True, 'request': item, 'liked': liked})


@feature_request_bp.route('/api/feature_requests/<request_id>/status', methods=['POST'])
def api_update_feature_request_status(request_id):
    denied = require_admin_json()
    if denied:
        return denied
    data = request.get_json(silent=True) or {}
    status = str(data.get('status', '') or '').strip()
    if status not in ALLOWED_STATUSES:
        return jsonify({'success': False, 'error': '\u65e0\u6548\u7684\u9700\u6c42\u72b6\u6001'})
    item = _update_request_status(request_id, status)
    if not item:
        return jsonify({'success': False, 'error': '\u9700\u6c42\u4e0d\u5b58\u5728'})
    return jsonify({'success': True, 'request': item})


@feature_request_bp.route('/feature_attachments/<filename>')
def feature_attachment(filename):
    path = safe_join(ATTACH_DIR, filename)
    if not path or not os.path.exists(path):
        abort(404)
    return send_file(path)





