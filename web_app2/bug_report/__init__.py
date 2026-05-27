# -*- coding: utf-8 -*-
"""Bug 閹绘劒姘﹂弽蹇曟窗 閳?Blueprint"""

import json
import os
import sqlite3
import time
import uuid

from flask import Blueprint, abort, send_file
from werkzeug.utils import secure_filename, safe_join

from shared import request, jsonify


bug_report_bp = Blueprint('bug_report', __name__)

BASE_DIR = os.path.dirname(os.path.dirname(__file__))
BUG_DIR = os.path.join(BASE_DIR, 'bug_reports')
ATTACH_DIR = os.path.join(BUG_DIR, 'attachments')
DB_PATH = os.path.join(BUG_DIR, 'reports.sqlite3')
ALLOWED_ATTACHMENT_EXTS = {'.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp', '.xlsx', '.xlsm', '.xls', '.csv', '.txt', '.log', '.zip', '.rar', '.7z'}
ALLOWED_STATUSES = {'\u5f85\u5904\u7406', '\u5904\u7406\u4e2d', '\u5df2\u4fee\u590d', '\u5df2\u5173\u95ed', '\u6682\u7f13', '\u65e0\u6cd5\u590d\u73b0'}

SEED_REPORTS = [
    {
        'id': 'seed-bug-bom-header-detect',
        'reporter': '\u7cfb\u7edf\u793a\u4f8b',
        'employee_id': 'SYSTEM',
        'module': 'BOM \u683c\u5f0f\u8f6c\u6362',
        'severity': '\u4e00\u822c',
        'status': '\u5f85\u5904\u7406',
        'title': 'BOM \u8868\u5934\u884c\u586b\u9519\u540e\u9884\u89c8\u5217\u6620\u5c04\u4e3a\u7a7a',
        'description': '\u4e0a\u4f20\u67d0\u4e9b\u5ba2\u6237 BOM \u65f6\uff0c\u5982\u679c\u8868\u5934\u884c\u586b\u9519\uff0c\u9884\u89c8\u533a\u6ca1\u6709\u660e\u786e\u63d0\u793a\uff0c\u7528\u6237\u4f1a\u8bef\u4ee5\u4e3a\u6587\u4ef6\u4e0d\u652f\u6301\u3002',
        'steps': '1. \u8fdb\u5165 BOM \u683c\u5f0f\u8f6c\u6362\n2. \u4e0a\u4f20\u4e00\u4efd\u8868\u5934\u5728\u7b2c 3 \u884c\u7684 BOM\n3. \u5c06\u8868\u5934\u884c\u8bbe\u4e3a 1 \u540e\u5237\u65b0\u9884\u89c8',
        'expected': '\u7cfb\u7edf\u63d0\u793a\u53ef\u80fd\u8868\u5934\u884c\u4e0d\u6b63\u786e\uff0c\u5e76\u7ed9\u51fa\u68c0\u6d4b\u5230\u7684\u5019\u9009\u8868\u5934\u884c\u3002',
    },
    {
        'id': 'seed-bug-feishu-cache-timeout',
        'reporter': '\u7cfb\u7edf\u793a\u4f8b',
        'employee_id': 'SYSTEM',
        'module': '\u98de\u4e66\u4f18\u9009\u5e93+\u5173\u7cfb\u5e93\u5339\u914d',
        'severity': '\u4e25\u91cd',
        'status': '\u5904\u7406\u4e2d',
        'title': '\u98de\u4e66\u7f13\u5b58\u5237\u65b0\u8d85\u65f6\u65f6\u7f3a\u5c11\u660e\u786e\u9519\u8bef\u4fe1\u606f',
        'description': '\u7f13\u5b58\u591a\u4e2a sheet \u65f6\uff0c\u5076\u53d1\u51fa\u73b0\u8bf7\u6c42\u8d85\u65f6\uff0c\u9875\u9762\u53ea\u663e\u793a\u7f13\u5b58\u5931\u8d25\uff0c\u4e0d\u77e5\u9053\u662f token\u3001\u7f51\u7edc\u8fd8\u662f\u8868\u683c\u6743\u9650\u95ee\u9898\u3002',
        'steps': '1. \u8fde\u63a5\u98de\u4e66 token\n2. \u52fe\u9009\u591a\u4e2a\u4f18\u9009\u5e93 sheet\n3. \u70b9\u51fb\u5237\u65b0\u7f13\u5b58\uff0c\u7b49\u5f85\u8d85\u65f6',
        'expected': '\u5931\u8d25\u65f6\u663e\u793a\u5177\u4f53 sheet \u540d\u3001HTTP \u72b6\u6001\u6216\u8d85\u65f6\u539f\u56e0\uff0c\u65b9\u4fbf\u5b9a\u4f4d\u3002',
    },
    {
        'id': 'seed-bug-plm-multi-config-zip',
        'reporter': '\u7cfb\u7edf\u793a\u4f8b',
        'employee_id': 'SYSTEM',
        'module': '\u8f6c\u6362\u4e3a\u4e0a\u4f20PLM\u7cfb\u7edf\u683c\u5f0f',
        'severity': '\u5efa\u8bae',
        'status': '\u5df2\u4fee\u590d',
        'title': 'PLM \u591a\u914d\u7f6e\u5bfc\u51fa ZIP \u540e\u5e0c\u671b\u663e\u793a\u751f\u6210\u6587\u4ef6\u6e05\u5355',
        'description': '\u591a\u4e2a\u7528\u91cf\u914d\u7f6e\u4f1a\u81ea\u52a8\u6253\u5305 ZIP\uff0c\u4f46\u5b8c\u6210\u540e\u9875\u9762\u6ca1\u6709\u5c55\u793a\u5177\u4f53\u5305\u542b\u54ea\u4e9b\u914d\u7f6e\u6587\u4ef6\u3002',
        'steps': '1. \u4e0a\u4f20\u6574\u673a BOM \u914d\u7f6e\u8868\n2. \u6dfb\u52a0 3 \u4e2a\u7528\u91cf\u914d\u7f6e\n3. \u5f00\u59cb\u8f6c\u6362\u5e76\u4e0b\u8f7d ZIP',
        'expected': '\u8f6c\u6362\u5b8c\u6210\u540e\u663e\u793a ZIP \u5185\u7684\u6587\u4ef6\u540d\u6e05\u5355\uff0c\u4fbf\u4e8e\u6838\u5bf9\u662f\u5426\u6f0f\u751f\u6210\u3002',
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
        CREATE TABLE IF NOT EXISTS bug_reports (
            id TEXT PRIMARY KEY,
            submitted_at REAL NOT NULL,
            reporter TEXT NOT NULL,
            employee_id TEXT NOT NULL,
            module TEXT NOT NULL,
            severity TEXT NOT NULL,
            status TEXT NOT NULL,
            title TEXT NOT NULL,
            description TEXT NOT NULL,
            steps TEXT NOT NULL,
            expected TEXT NOT NULL,
            attachments TEXT NOT NULL
        )
    ''')
    _seed_reports(conn)
    conn.commit()
    return conn


def _seed_reports(conn):
    base_time = time.time() - 86400 * len(SEED_REPORTS)
    for idx, item in enumerate(SEED_REPORTS):
        conn.execute('''
            INSERT OR IGNORE INTO bug_reports (
                id, submitted_at, reporter, employee_id, module, severity, status,
                title, description, steps, expected, attachments
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            item['id'], base_time + idx * 3600, item['reporter'], item['employee_id'],
            item['module'], item['severity'], item['status'], item['title'],
            item['description'], item['steps'], item['expected'], '[]',
        ))


def _row_to_report(row):
    item = dict(row)
    try:
        item['attachments'] = json.loads(item.get('attachments') or '[]')
    except Exception:
        item['attachments'] = []
    return item


def _filter_reports(reports, status='', module='', query=''):
    status = str(status or '').strip()
    module = str(module or '').strip()
    query = str(query or '').strip().lower()
    filtered = []
    for item in reports:
        if status and item.get('status') != status:
            continue
        if module and item.get('module') != module:
            continue
        if query:
            haystack = '\n'.join(str(item.get(key, '') or '') for key in (
                'title', 'description', 'steps', 'expected', 'reporter', 'employee_id', 'module', 'severity',
            )).lower()
            if query not in haystack:
                continue
        filtered.append(item)
    return filtered


def _read_reports(status='', module='', query=''):
    conn = _connect()
    try:
        rows = conn.execute('SELECT * FROM bug_reports ORDER BY submitted_at DESC').fetchall()
        reports = [_row_to_report(row) for row in rows]
        return _filter_reports(reports, status=status, module=module, query=query)
    finally:
        conn.close()

def _insert_report(report):
    conn = _connect()
    try:
        conn.execute('''
            INSERT INTO bug_reports (
                id, submitted_at, reporter, employee_id, module, severity, status,
                title, description, steps, expected, attachments
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            report['id'], report['submitted_at'], report['reporter'], report['employee_id'],
            report['module'], report['severity'], report['status'], report['title'],
            report['description'], report['steps'], report['expected'],
            json.dumps(report['attachments'], ensure_ascii=False),
        ))
        conn.commit()
    finally:
        conn.close()


def _update_report_status(report_id, status):
    conn = _connect()
    try:
        conn.execute('UPDATE bug_reports SET status = ? WHERE id = ?', (status, report_id))
        row = conn.execute('SELECT * FROM bug_reports WHERE id = ?', (report_id,)).fetchone()
        if not row:
            return None
        conn.commit()
        return _row_to_report(row)
    finally:
        conn.close()


def _clean_text(name, max_len=2000):
    return str(request.form.get(name, '') or '').strip()[:max_len]


def _save_attachments(report_id):
    _ensure_dirs()
    saved = []
    for file in request.files.getlist('images'):
        if not file or not file.filename:
            continue
        _, ext = os.path.splitext(file.filename.lower())
        if ext not in ALLOWED_ATTACHMENT_EXTS:
            raise ValueError('仅支持上传图片、Excel、CSV、日志文本或压缩包附件')
        safe_name = secure_filename(file.filename) or f'attachment{ext}'
        out_name = f'{report_id}_{uuid.uuid4().hex[:8]}_{safe_name}'
        out_path = os.path.join(ATTACH_DIR, out_name)
        file.save(out_path)
        saved.append({
            'name': file.filename,
            'url': f'/bug_attachments/{out_name}',
        })
    return saved


@bug_report_bp.route('/api/bug_reports', methods=['GET'])
def api_bug_reports():
    reports = _read_reports(
        status=request.args.get('status', ''),
        module=request.args.get('module', ''),
        query=request.args.get('q', ''),
    )
    return jsonify({'success': True, 'reports': reports, 'total': len(reports)})

@bug_report_bp.route('/api/bug_reports', methods=['POST'])
def api_submit_bug_report():
    reporter = _clean_text('reporter', 80)
    employee_id = _clean_text('employee_id', 40)
    title = _clean_text('title', 120)
    description = _clean_text('description', 4000)
    if not reporter or not employee_id or not title or not description:
        return jsonify({'success': False, 'error': '请填写姓名、工号、问题标题和问题描述'})

    report_id = time.strftime('%Y%m%d%H%M%S') + '-' + uuid.uuid4().hex[:6]
    try:
        attachments = _save_attachments(report_id)
    except ValueError as e:
        return jsonify({'success': False, 'error': str(e)})

    report = {
        'id': report_id,
        'submitted_at': time.time(),
        'reporter': reporter,
        'employee_id': employee_id,
        'module': _clean_text('module', 80) or '未指定',
        'severity': _clean_text('severity', 40) or '一般',
        'status': '待处理',
        'title': title,
        'description': description,
        'steps': _clean_text('steps', 4000),
        'expected': _clean_text('expected', 2000),
        'attachments': attachments,
    }
    _insert_report(report)
    return jsonify({'success': True, 'report': report})


@bug_report_bp.route('/api/bug_reports/<report_id>/status', methods=['POST'])
def api_update_bug_status(report_id):
    data = request.get_json(silent=True) or {}
    status = str(data.get('status', '') or '').strip()
    if status not in ALLOWED_STATUSES:
        return jsonify({'success': False, 'error': '无效的处理状态'})
    report = _update_report_status(report_id, status)
    if not report:
        return jsonify({'success': False, 'error': '问题记录不存在'})
    return jsonify({'success': True, 'report': report})


@bug_report_bp.route('/bug_attachments/<filename>')
def bug_attachment(filename):
    path = safe_join(ATTACH_DIR, filename)
    if not path or not os.path.exists(path):
        abort(404)
    return send_file(path)
