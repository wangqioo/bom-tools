# -*- coding: utf-8 -*-
"""Bug 提交栏目 — Blueprint"""

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
    conn.commit()
    return conn


def _row_to_report(row):
    item = dict(row)
    try:
        item['attachments'] = json.loads(item.get('attachments') or '[]')
    except Exception:
        item['attachments'] = []
    return item


def _read_reports():
    conn = _connect()
    try:
        rows = conn.execute('SELECT * FROM bug_reports ORDER BY submitted_at DESC').fetchall()
        return [_row_to_report(row) for row in rows]
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
    return jsonify({'success': True, 'reports': _read_reports()})


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


@bug_report_bp.route('/bug_attachments/<filename>')
def bug_attachment(filename):
    path = safe_join(ATTACH_DIR, filename)
    if not path or not os.path.exists(path):
        abort(404)
    return send_file(path)
