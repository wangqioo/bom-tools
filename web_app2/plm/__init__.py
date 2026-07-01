# -*- coding: utf-8 -*-
"""PLM 上传工具 — Blueprint"""

import os, uuid, re, json, threading, time, queue
from zipfile import ZipFile, ZIP_DEFLATED
from flask import Blueprint
from activity import track_tool_activity
from shared import (
    openpyxl, Workbook, Font, PatternFill, Alignment, Border, Side,
    get_column_letter,
    request, jsonify,
    UPLOAD_DIR, OUTPUT_DIR, _col_int,
    _cell_str, _open_workbook, _request_int, _save_uploaded_excel, _save_or_reuse_uploaded_excel,
)

plm_bp = Blueprint('plm', __name__)

_PLM_ATTACHMENT_JOBS = {}
_PLM_ATTACHMENT_JOBS_LOCK = threading.Lock()
_PLM_ATTACHMENT_QUEUE = queue.Queue()
_PLM_ATTACHMENT_WORKER_STARTED = False
_PLM_ATTACHMENT_WORKER_LOCK = threading.Lock()
_PLM_ATTACHMENT_BATCHES = {}
_PLM_ATTACHMENT_BATCHES_LOCK = threading.Lock()
_PLM_SPEC_REVERSE_JOBS = {}
_PLM_SPEC_REVERSE_JOBS_LOCK = threading.Lock()
_PLM_SPEC_REVERSE_QUEUE = queue.Queue()
_PLM_SPEC_REVERSE_WORKER_STARTED = False
_PLM_SPEC_REVERSE_WORKER_LOCK = threading.Lock()

def _spec_reverse_progress_from_message(message, current):
    text = str(message or '')
    rules = [
        ('启动浏览器', 8, '启动浏览器'),
        ('打开 EIP', 18, '打开 EIP'),
        ('进入 PLM', 30, '进入 PLM'),
        ('打开功能地图', 42, '打开功能地图'),
        ('打开功能：', 52, '打开规格型号反查物料'),
        ('上传文件', 64, '上传查询文件'),
        ('点击查询', 74, '提交查询'),
        ('等待结果', 82, '等待 PLM 返回结果'),
        ('点击结果导出', 90, '导出结果'),
        ('导出完成', 98, '保存导出文件'),
    ]
    for needle, progress, stage in rules:
        if needle in text:
            return max(current, progress), stage
    return min(max(current, 5) + 1, 90), text[:80] or '处理中'


def _new_spec_reverse_job(source_label):
    job_id = uuid.uuid4().hex
    now = time.time()
    job = {
        'id': job_id,
        'status': 'queued',
        'stage': '任务已创建',
        'progress': 3,
        'source_label': source_label,
        'logs': [],
        'download': '',
        'filename': '',
        'source_path': '',
        'error': '',
        'created_at': now,
        'updated_at': now,
    }
    with _PLM_SPEC_REVERSE_JOBS_LOCK:
        _PLM_SPEC_REVERSE_JOBS[job_id] = job
    return job_id


def _update_spec_reverse_job(job_id, **updates):
    with _PLM_SPEC_REVERSE_JOBS_LOCK:
        job = _PLM_SPEC_REVERSE_JOBS.get(job_id)
        if not job:
            return
        job.update(updates)
        job['updated_at'] = time.time()


def _append_spec_reverse_log(job_id, message):
    message = str(message)
    with _PLM_SPEC_REVERSE_JOBS_LOCK:
        job = _PLM_SPEC_REVERSE_JOBS.get(job_id)
        if not job:
            return
        job['logs'].append(message)
        progress, stage = _spec_reverse_progress_from_message(message, int(job.get('progress') or 0))
        job['progress'] = progress
        job['stage'] = stage
        job['updated_at'] = time.time()


def _snapshot_spec_reverse_job(job_id):
    with _PLM_SPEC_REVERSE_JOBS_LOCK:
        job = _PLM_SPEC_REVERSE_JOBS.get(job_id)
        if not job:
            return None
        return dict(job, logs=list(job.get('logs') or []))


def _cleanup_spec_reverse_jobs():
    cutoff = time.time() - 3600
    with _PLM_SPEC_REVERSE_JOBS_LOCK:
        for job_id, job in list(_PLM_SPEC_REVERSE_JOBS.items()):
            if job.get('updated_at', 0) < cutoff:
                _PLM_SPEC_REVERSE_JOBS.pop(job_id, None)


def _ensure_spec_reverse_worker():
    global _PLM_SPEC_REVERSE_WORKER_STARTED
    with _PLM_SPEC_REVERSE_WORKER_LOCK:
        if _PLM_SPEC_REVERSE_WORKER_STARTED:
            return
        threading.Thread(target=_spec_reverse_worker_loop, daemon=True).start()
        _PLM_SPEC_REVERSE_WORKER_STARTED = True


def _enqueue_spec_reverse_job(job_id, username, password, upload_path):
    _PLM_SPEC_REVERSE_QUEUE.put({
        'job_id': job_id,
        'username': username,
        'password': password,
        'upload_path': upload_path,
    })
    _ensure_spec_reverse_worker()


def _spec_reverse_worker_loop():
    while True:
        task = _PLM_SPEC_REVERSE_QUEUE.get()
        job_id = task['job_id']
        try:
            from pathlib import Path as _Path
            from playwright.sync_api import sync_playwright
            from .automation import require_feature, run_plm_feature

            _update_spec_reverse_job(job_id, status='running', stage='准备启动浏览器', progress=5)
            feature = require_feature('spec_reverse_material')
            with sync_playwright() as playwright:
                output_path = run_plm_feature(
                    playwright,
                    username=task['username'],
                    password=task['password'],
                    feature=feature,
                    upload_file=_Path(task['upload_path']),
                    output_dir=_Path(OUTPUT_DIR),
                    headless=False,
                    log=lambda message: _append_spec_reverse_log(job_id, message),
                )
            output_path = str(output_path)
            if not os.path.exists(output_path):
                raise RuntimeError('自动化完成但未找到导出文件')
            out_name = os.path.basename(output_path)
            _update_spec_reverse_job(
                job_id,
                status='done',
                stage='导出完成',
                progress=100,
                download=f'/download/{out_name}',
                filename=out_name,
                source_path=output_path,
            )
        except ImportError as exc:
            _append_spec_reverse_log(job_id, str(exc))
            _update_spec_reverse_job(
                job_id,
                status='error',
                stage='缺少 Playwright 依赖',
                progress=100,
                error='缺少 Playwright 依赖，请在 BOM 工具环境安装 requirements.txt 并执行 playwright install chromium',
            )
        except Exception as exc:
            _append_spec_reverse_log(job_id, str(exc))
            _update_spec_reverse_job(job_id, status='error', stage='执行失败', progress=100, error=str(exc))
        finally:
            _PLM_SPEC_REVERSE_QUEUE.task_done()


def _split_spec_reverse_single_values(value):
    text = str(value or '').replace('\u3000', ' ').strip()
    values = [part.strip() for part in re.split(r'[,，]', text) if part.strip()]
    if not values:
        raise ValueError('请输入规格型号或 HQ 料号')
    return values


def _create_spec_reverse_single_excel(value, uid):
    values = _split_spec_reverse_single_values(value)
    wb = Workbook()
    ws = wb.active
    ws.title = '规格型号'
    ws.cell(row=1, column=1, value='规格型号').font = Font(bold=True)
    for row_idx, text in enumerate(values, start=2):
        ws.cell(row=row_idx, column=1, value=text)
    ws.column_dimensions['A'].width = 40
    in_path = os.path.join(UPLOAD_DIR, f'plm_auto_single_{uid}.xlsx')
    wb.save(in_path)
    source_label = values[0] if len(values) == 1 else f'{values[0]} 等 {len(values)} 项'
    return in_path, source_label

def _new_attachment_job(hqpn):
    job_id = uuid.uuid4().hex
    now = time.time()
    job = {
        'id': job_id,
        'status': 'queued',
        'stage': '\u4efb\u52a1\u5df2\u521b\u5efa',
        'progress': 3,
        'hqpn': hqpn,
        'logs': [],
        'download': '',
        'filename': '',
        'source_path': '',
        'error': '',
        'batch_id': '',
        'created_at': now,
        'updated_at': now,
    }
    with _PLM_ATTACHMENT_JOBS_LOCK:
        _PLM_ATTACHMENT_JOBS[job_id] = job
    return job_id


def _attachment_progress_from_message(message, current):
    text = str(message or '')
    rules = [
        ('\u542f\u52a8\u6d4f\u89c8\u5668', 8, '\u542f\u52a8\u6d4f\u89c8\u5668'),
        ('\u6253\u5f00 EIP', 15, '\u6253\u5f00 EIP'),
        ('\u8fdb\u5165 PLM', 25, '\u8fdb\u5165 PLM'),
        ('Open PLM search page', 35, '\u6253\u5f00 PLM \u641c\u7d22\u9875'),
        ('\u590d\u7528\u5df2\u767b\u5f55 PLM \u4f1a\u8bdd', 30, '\u590d\u7528\u5df2\u767b\u5f55 PLM \u4f1a\u8bdd'),
        ('\u76f4\u63a5\u8fdb\u5165 PLM \u641c\u7d22\u9875', 35, '\u6253\u5f00 PLM \u641c\u7d22\u9875'),
        ('\u641c\u7d22\u6599\u53f7', 45, '\u641c\u7d22 HQ \u6599\u53f7'),
        ('\u6253\u5f00\u7b2c\u4e00\u6761\u641c\u7d22\u7ed3\u679c', 58, '\u6253\u5f00\u7269\u6599\u8be6\u60c5'),
        ('\u8fdb\u5165\u5185\u5bb9\u9875', 68, '\u8fdb\u5165\u5185\u5bb9\u9875'),
        ('\u52fe\u9009\u9644\u4ef6\u5e76\u4e0b\u8f7d', 76, '\u52fe\u9009\u9644\u4ef6'),
        ('\u8bc6\u522b PDF \u9644\u4ef6', 80, '\u8bc6\u522b\u9644\u4ef6'),
        ('Selected all attachment rows', 84, '\u52fe\u9009\u9644\u4ef6\u884c'),
        ('No immediate download event', 88, '\u7b49\u5f85 PDF \u9884\u89c8\u9875'),
        ('Downloaded PDF response', 94, '\u4fdd\u5b58 PDF \u9644\u4ef6'),
        ('Downloaded PDF viewer', 94, '\u4fdd\u5b58 PDF \u9644\u4ef6'),
        ('Downloaded selected attachments', 94, '\u4fdd\u5b58\u9644\u4ef6\u538b\u7f29\u5305'),
        ('\u4e0b\u8f7d\u5b8c\u6210', 98, '\u6574\u7406\u4e0b\u8f7d\u6587\u4ef6'),
    ]
    for needle, progress, stage in rules:
        if needle in text:
            return max(current, progress), stage
    return min(max(current, 5) + 1, 90), text[:80] or '\u5904\u7406\u4e2d'


def _update_attachment_job(job_id, **updates):
    with _PLM_ATTACHMENT_JOBS_LOCK:
        job = _PLM_ATTACHMENT_JOBS.get(job_id)
        if not job:
            return
        job.update(updates)
        job['updated_at'] = time.time()


def _append_attachment_log(job_id, message):
    message = str(message)
    with _PLM_ATTACHMENT_JOBS_LOCK:
        job = _PLM_ATTACHMENT_JOBS.get(job_id)
        if not job:
            return
        job['logs'].append(message)
        progress, stage = _attachment_progress_from_message(message, int(job.get('progress') or 0))
        job['progress'] = progress
        job['stage'] = stage
        job['updated_at'] = time.time()


def _snapshot_attachment_job(job_id):
    with _PLM_ATTACHMENT_JOBS_LOCK:
        job = _PLM_ATTACHMENT_JOBS.get(job_id)
        if not job:
            return None
        return dict(job, logs=list(job.get('logs') or []))


def _ensure_attachment_worker():
    global _PLM_ATTACHMENT_WORKER_STARTED
    with _PLM_ATTACHMENT_WORKER_LOCK:
        if _PLM_ATTACHMENT_WORKER_STARTED:
            return
        threading.Thread(target=_attachment_worker_loop, daemon=True).start()
        _PLM_ATTACHMENT_WORKER_STARTED = True


def _enqueue_attachment_job(job_id, username, password, hqpn, batch_id=''):
    _PLM_ATTACHMENT_QUEUE.put({
        'job_id': job_id,
        'username': username,
        'password': password,
        'hqpn': hqpn,
        'batch_id': batch_id,
    })
    _ensure_attachment_worker()


def _attachment_worker_loop():
    playwright_ctx = None
    playwright = None
    browser = None
    context = None
    search_page = None
    session_user = ''
    session_password = ''
    while True:
        task = _PLM_ATTACHMENT_QUEUE.get()
        job_id = task['job_id']
        username = task['username']
        password = task['password']
        hqpn = task['hqpn']
        batch_id = task.get('batch_id') or ''
        if batch_id and _batch_cancelled(batch_id):
            _update_attachment_job(job_id, status='cancelled', stage='\\u5df2\\u53d6\\u6d88', progress=100, error='\\u7528\\u6237\\u53d6\\u6d88')
            _PLM_ATTACHMENT_QUEUE.task_done()
            continue
        try:
            from pathlib import Path as _Path
            from playwright.sync_api import sync_playwright
            from .automation import (
                START_URL,
                click_opening_page,
                download_hq_attachment_from_search_page,
                login_if_present,
                wait_for_eip_ready,
                _open_plm_search_page,
                _wait_for_plm_home,
            )

            if playwright is None:
                playwright_ctx = sync_playwright()
                playwright = playwright_ctx.__enter__()

            needs_new_session = (
                browser is None or context is None or search_page is None or
                session_user != username or session_password != password
            )
            if not needs_new_session:
                try:
                    needs_new_session = bool(search_page.is_closed())
                except Exception:
                    needs_new_session = True

            if needs_new_session:
                for resource in (context, browser):
                    try:
                        if resource:
                            resource.close()
                    except Exception:
                        pass
                _update_attachment_job(job_id, status='running', stage='\u51c6\u5907\u542f\u52a8\u6d4f\u89c8\u5668', progress=5)
                _append_attachment_log(job_id, '\u542f\u52a8\u6d4f\u89c8\u5668')
                browser = playwright.chromium.launch(headless=False)
                context = browser.new_context(accept_downloads=True)
                page = context.new_page()

                _append_attachment_log(job_id, '\u6253\u5f00 EIP')
                page.goto(START_URL, wait_until='domcontentloaded', timeout=60000)
                wait_for_eip_ready(page, username, password)

                _append_attachment_log(job_id, '\u8fdb\u5165 PLM')
                plm_page = click_opening_page(
                    page,
                    page.locator('a').filter(has_text=re.compile(r'^PLM$')),
                    timeout=30000,
                )
                login_if_present(plm_page, username, password, timeout=500)
                plm_page = _wait_for_plm_home(context, plm_page, username, password)

                _append_attachment_log(job_id, 'Open PLM search page')
                search_page = _open_plm_search_page(
                    context,
                    plm_page,
                    username,
                    password,
                    log=lambda message: _append_attachment_log(job_id, message),
                )
                session_user = username
                session_password = password
            else:
                _update_attachment_job(job_id, status='running', stage='\u590d\u7528\u5df2\u767b\u5f55 PLM \u4f1a\u8bdd', progress=30)
                _append_attachment_log(job_id, '\u590d\u7528\u5df2\u767b\u5f55 PLM \u4f1a\u8bdd')

            last_error = None
            output_path = None
            for attempt in range(2):
                try:
                    if attempt:
                        _update_attachment_job(job_id, status='running', stage='重试当前 HQ 料号', progress=35)
                        _append_attachment_log(job_id, f'首次下载失败，复用当前 PLM 会话重试：{last_error}')
                        search_page = _open_plm_search_page(
                            context,
                            search_page,
                            username,
                            password,
                            log=lambda message: _append_attachment_log(job_id, message),
                        )
                    output_path, search_page = download_hq_attachment_from_search_page(
                        context,
                        search_page,
                        hqpn=hqpn,
                        output_dir=_Path(OUTPUT_DIR),
                        username=username,
                        password=password,
                        log=lambda message: _append_attachment_log(job_id, message),
                    )
                    break
                except Exception as exc:
                    last_error = exc
                    if attempt:
                        raise
                    try:
                        if context is None or search_page is None or search_page.is_closed():
                            raise
                    except Exception:
                        raise
            output_path = str(output_path)
            if not os.path.exists(output_path):
                raise RuntimeError('\u81ea\u52a8\u5316\u5b8c\u6210\u4f46\u672a\u627e\u5230\u4e0b\u8f7d\u6587\u4ef6')
            out_name = os.path.basename(output_path)
            _update_attachment_job(
                job_id,
                status='done',
                stage='\u4e0b\u8f7d\u5b8c\u6210',
                progress=100,
                download=f'/download/{out_name}',
                filename=out_name,
                source_path=output_path,
            )
        except Exception as e:
            _append_attachment_log(job_id, str(e))
            _update_attachment_job(job_id, status='error', stage='\u6267\u884c\u5931\u8d25', progress=100, error=str(e))
            try:
                if context:
                    context.close()
            except Exception:
                pass
            try:
                if browser:
                    browser.close()
            except Exception:
                pass
            browser = None
            context = None
            search_page = None
            session_user = ''
            session_password = ''
        finally:
            try:
                should_close_session = False
                if not batch_id:
                    should_close_session = True
                elif _batch_jobs_finished(batch_id):
                    _build_attachment_batch_status(batch_id)
                    should_close_session = True
                if should_close_session:
                    try:
                        if context:
                            context.close()
                    except Exception:
                        pass
                    try:
                        if browser:
                            browser.close()
                    except Exception:
                        pass
                    browser = None
                    context = None
                    search_page = None
                    session_user = ''
                    session_password = ''
            finally:
                _PLM_ATTACHMENT_QUEUE.task_done()



def _safe_zip_member_name(filename, used_names):
    base = os.path.basename(str(filename or '')).strip() or 'attachment.zip'
    stem, ext = os.path.splitext(base)
    if not ext:
        ext = '.zip'
    candidate = base
    n = 2
    while candidate in used_names:
        candidate = f"{stem}_{n}{ext}"
        n += 1
    used_names.add(candidate)
    return candidate


def _snapshot_attachment_batch(batch_id):
    with _PLM_ATTACHMENT_BATCHES_LOCK:
        batch = _PLM_ATTACHMENT_BATCHES.get(batch_id)
        if not batch:
            return None
        return dict(batch, job_ids=list(batch.get('job_ids') or []))


def _cleanup_attachment_batches():
    cutoff = time.time() - 3600
    with _PLM_ATTACHMENT_BATCHES_LOCK:
        for batch_id, batch in list(_PLM_ATTACHMENT_BATCHES.items()):
            if batch.get('updated_at', 0) < cutoff:
                _PLM_ATTACHMENT_BATCHES.pop(batch_id, None)


def _build_attachment_batch_status(batch_id):
    batch = _snapshot_attachment_batch(batch_id)
    if not batch:
        return None
    jobs = []
    for job_id in batch.get('job_ids') or []:
        job = _snapshot_attachment_job(job_id)
        if job:
            jobs.append(job)
    total = len(jobs)
    done = sum(1 for job in jobs if job.get('status') == 'done')
    failed = sum(1 for job in jobs if job.get('status') == 'error')
    cancelled = sum(1 for job in jobs if job.get('status') == 'cancelled')
    finished = done + failed + cancelled
    if total:
        progress = int(sum(int(job.get('progress') or 0) for job in jobs) / total)
    else:
        progress = 0
    if total and finished == total:
        progress = 100
        if not batch.get('download') and done:
            uid = batch_id[:8]
            zip_name = f"HQ\u9644\u4ef6\u6279\u91cf_{uid}.zip"
            zip_path = os.path.join(OUTPUT_DIR, zip_name)
            used_names = set()
            with ZipFile(zip_path, 'w', ZIP_DEFLATED) as zf:
                for job in jobs:
                    source_path = job.get('source_path') or ''
                    if job.get('status') == 'done' and source_path and os.path.exists(source_path):
                        member_name = _safe_zip_member_name(job.get('filename') or os.path.basename(source_path), used_names)
                        zf.write(source_path, arcname=member_name)
            with _PLM_ATTACHMENT_BATCHES_LOCK:
                current = _PLM_ATTACHMENT_BATCHES.get(batch_id)
                if current is not None:
                    current.update({
                        'download': f'/download/{zip_name}',
                        'filename': zip_name,
                        'source_path': zip_path,
                        'updated_at': time.time(),
                    })
                    batch = dict(current, job_ids=list(current.get('job_ids') or []))
        stage = '\u6279\u91cf\u4e0b\u8f7d\u5b8c\u6210'
    elif total and any(job.get('status') == 'running' for job in jobs):
        current_job = next((job for job in jobs if job.get('status') == 'running'), None)
        stage = f"\u6b63\u5728\u4e0b\u8f7d\uff1a{current_job.get('hqpn')}" if current_job else '\u6b63\u5728\u4e0b\u8f7d'
    else:
        stage = '\u6392\u961f\u7b49\u5f85'
    return {
        'id': batch_id,
        'status': 'done' if total and finished == total else 'running',
        'stage': stage,
        'progress': progress,
        'total': total,
        'done': done,
        'failed': failed,
        'cancelled': cancelled,
        'finished': finished,
        'download': batch.get('download', ''),
        'filename': batch.get('filename', ''),
        'jobs': [{
            'job_id': job.get('id'),
            'hqpn': job.get('hqpn'),
            'status': job.get('status'),
            'stage': job.get('stage'),
            'progress': job.get('progress'),
            'download': job.get('download'),
            'filename': job.get('filename'),
            'error': job.get('error'),
        } for job in jobs],
    }




def _batch_cancelled(batch_id):
    if not batch_id:
        return False
    with _PLM_ATTACHMENT_BATCHES_LOCK:
        batch = _PLM_ATTACHMENT_BATCHES.get(batch_id)
        return bool(batch and batch.get('cancelled'))


def _cancel_attachment_batch(batch_id):
    with _PLM_ATTACHMENT_BATCHES_LOCK:
        batch = _PLM_ATTACHMENT_BATCHES.get(batch_id)
        if not batch:
            return None
        batch['cancelled'] = True
        batch['updated_at'] = time.time()
        job_ids = list(batch.get('job_ids') or [])
    for job_id in job_ids:
        job = _snapshot_attachment_job(job_id)
        if job and job.get('status') == 'queued':
            _update_attachment_job(job_id, status='cancelled', stage='\u5df2\u53d6\u6d88', progress=100, error='\u7528\u6237\u53d6\u6d88')
    return _build_attachment_batch_status(batch_id)
def _batch_jobs_finished(batch_id):
    if not batch_id:
        return False
    batch = _snapshot_attachment_batch(batch_id)
    if not batch:
        return False
    job_ids = batch.get('job_ids') or []
    if not job_ids:
        return False
    for job_id in job_ids:
        job = _snapshot_attachment_job(job_id)
        if not job or job.get('status') not in ('done', 'error', 'cancelled'):
            return False
    return True
def _cleanup_attachment_jobs():
    cutoff = time.time() - 3600
    with _PLM_ATTACHMENT_JOBS_LOCK:
        for job_id, job in list(_PLM_ATTACHMENT_JOBS.items()):
            if job.get('updated_at', 0) < cutoff:
                _PLM_ATTACHMENT_JOBS.pop(job_id, None)
PLM_HEADERS = [
    "序号", "料号", "型号", "物料描述", "单耗",
    "替代关系\n(A:完全替代/N:独供/X:不完全替代)",
    "位号", "生产厂家", "是否环保", "温敏属性", "备注",
    "主辅BOM标记\n(仅允许填写二供/三供/四供/五供/六供/七供/八供)",
    "MBG优选属性", "CBG优选属性", "DBG优选属性", "首制程", "次制程", "次制程单耗",
    "是否可量产下单", "次制程位号", "ABG优选属性", "IFM_PART", "PCD_PART",
    "是否受EAR管控", "ECCN",
]

PLM_IDX_SEQ  = 0
PLM_IDX_HQPN = 1
PLM_IDX_QTY  = 4
PLM_IDX_MARK = 11  # 主辅BOM标记

def _detect_columns(ws, header_row):
    result = {}
    found_headers = []
    scan_cols = max((ws.max_column or 0) + 5, 30)
    for ci in range(1, scan_cols + 1):
        raw = ws.cell(row=header_row, column=ci).value
        if raw is None:
            continue
        h = str(raw).replace('\n', '').replace('\r', '').strip()
        hl = h.lower().replace(' ', '')
        if h:
            found_headers.append(f"{get_column_letter(ci)}:{h}")
        if '序号' in h:
            result.setdefault('seq', ci)
        if 'hq' in hl and 'pn' in hl:
            result.setdefault('hq_pn', ci)
        if '主二供' in h or '主供' in h:
            result.setdefault('supply_type', ci)
        if '用量' in h or '单耗' in h:
            result.setdefault('qty', ci)
    return result, found_headers


def _safe_qty(v):
    if v is None:
        return None
    s = str(v).strip()
    if s == "":
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _split_col_refs(raw):
    refs = []
    for part in re.split(r'[\s,\uFF0C;\uFF1B]+', str(raw or '').strip()):
        part = part.strip()
        if part:
            refs.append(part)
    return refs


def _safe_filename_part(value):
    text = str(value or '').strip() or '\u672a\u547d\u540d'
    return re.sub(r'[\\/*?:"<>|]', '_', text)


def _do_convert(in_file, sheet_name, header_row,
                col_seq, col_hqpn, col_stype, col_qty, project_name, out_file):
    wb_in = _open_workbook(in_file, data_only=True)
    ws_in = wb_in[sheet_name]
    max_col = ws_in.max_column

    data_rows = []
    for ri in range(header_row + 1, ws_in.max_row + 1):
        rv = {ci: ws_in.cell(row=ri, column=ci).value for ci in range(1, max_col + 1)}
        if any(v is not None and str(v).strip() for v in rv.values()):
            data_rows.append(rv)

    wb_out = Workbook()
    ws_out = wb_out.active
    ws_out.title = "PLM导入"

    bdr = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    meta_font = Font(bold=True, size=10)

    # Rows 1-2: metadata
    ws_out.cell(row=1, column=1, value="料号:").font = meta_font
    ws_out.cell(row=1, column=2, value=project_name or "").font = Font(size=10)
    ws_out.cell(row=1, column=3, value="描述:").font = meta_font
    ws_out.cell(row=1, column=5, value="项目配置名:").font = meta_font
    ws_out.cell(row=1, column=7, value="工程师:").font = meta_font
    ws_out.cell(row=2, column=1, value="版本:").font = meta_font
    ws_out.cell(row=2, column=3, value="替代项").font = meta_font
    ws_out.cell(row=2, column=5, value="BOM名称:").font = meta_font
    ws_out.cell(row=2, column=7, value="归档部门:").font = meta_font

    # Row 3: headers
    for offset, hdr_txt in enumerate(PLM_HEADERS):
        c = ws_out.cell(row=3, column=offset + 1, value=hdr_txt)
        c.font = Font(bold=True, color="FF0000", size=9)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = bdr
        ws_out.column_dimensions[get_column_letter(offset + 1)].width = 14
    ws_out.column_dimensions[get_column_letter(PLM_IDX_HQPN + 1)].width = 22
    ws_out.row_dimensions[3].height = 60

    # Data rows from row 4
    dr = 4
    total = 0
    skipped = 0
    skip_logs = []
    for rv in data_rows:
        source_seq = rv.get(col_seq)
        if not source_seq or str(source_seq).strip() == "":
            skipped += 1
            skip_logs.append("  跳过（序号为空）")
            continue

        qty_raw = rv.get(col_qty)
        if qty_raw is None or str(qty_raw).strip() == "":
            skipped += 1
            skip_logs.append(f"  跳过（用量为空）: 序号 {str(source_seq).strip()}")
            continue

        qty = _safe_qty(qty_raw)
        if qty is None:
            skipped += 1
            skip_logs.append(f"  跳过（用量非数字）: 序号 {str(source_seq).strip()}")
            continue

        hqpn = rv.get(col_hqpn)
        hqpn_str = str(hqpn).strip() if hqpn is not None else ""
        stype_str = str(rv.get(col_stype) or "").strip()


        def wc(idx, val, row=dr):
            cc = ws_out.cell(row=row, column=idx + 1, value=val)
            cc.alignment = Alignment(horizontal="left", vertical="center")
            cc.border = bdr

        wc(PLM_IDX_SEQ, source_seq)
        wc(PLM_IDX_HQPN, hqpn_str)

        if qty != 0:
            wc(PLM_IDX_QTY, qty)

        if stype_str and stype_str != "主供":
            wc(PLM_IDX_MARK, stype_str)

        dr += 1
        total += 1

    wb_out.save(out_file)
    wb_in.close()
    return total, skipped, skip_logs


# ── 路由 ─────────────────────────────────────────────────────

@plm_bp.route('/api/plm/detect', methods=['POST'])
def api_plm_detect():
    file = request.files.get('file')
    try:
        uid, in_path = _save_or_reuse_uploaded_excel(file, "plm_pre", request.form.get('uid', ''))
        wb = _open_workbook(in_path, read_only=True, data_only=True)
    except ValueError as e:
        return jsonify({'success': False, 'error': str(e)})
    sheets = wb.sheetnames
    wb.close()

    sheet_name = request.form.get('sheet_name', '')
    if not sheet_name or sheet_name not in sheets:
        sheet_name = sheets[0] if sheets else ''
    header_row = _request_int('header_row', 4)
    if header_row is None:
        return jsonify({'success': False, 'error': '表头行必须是大于等于 1 的数字'})

    try:
        wb2 = _open_workbook(in_path, data_only=True)
    except ValueError as e:
        return jsonify({'success': False, 'error': str(e)})
    ws = wb2[sheet_name] if sheet_name else wb2[wb2.sheetnames[0]]
    detected, raw_headers = _detect_columns(ws, header_row)

    # Preview
    preview_headers = [ws.cell(row=header_row, column=ci).value for ci in range(1, ws.max_column + 1)]
    preview = []
    for ri in range(header_row + 1, min(header_row + 51, ws.max_row + 1)):
        row = [ws.cell(row=ri, column=ci).value for ci in range(1, ws.max_column + 1)]
        if any(v is not None and str(v).strip() for v in row):
            preview.append([str(v) if v is not None else "" for v in row])
    wb2.close()

    result = {k: get_column_letter(v) for k, v in detected.items() if v}
    return jsonify({
        'success': True,
        'uid': uid,
        'sheets': sheets,
        'current_sheet': sheet_name,
        'headers': raw_headers,
        'preview_headers': [str(h) if h is not None else "" for h in preview_headers],
        'preview': preview,
        'detected': result,
    })


@plm_bp.route('/api/plm/convert', methods=['POST'])
@track_tool_activity('PLM格式转换')
def api_plm_convert():
    file = request.files.get('file')
    if not file:
        return jsonify({'success': False, 'error': '请上传文件'})

    uid = str(uuid.uuid4())[:8]
    try:
        in_path = _save_uploaded_excel(file, "plm_in", uid)
    except ValueError as e:
        return jsonify({'success': False, 'error': str(e)})

    sheet_name = request.form.get('sheet', '')
    header_row = _request_int('header_row', 4)
    if header_row is None:
        return jsonify({'success': False, 'error': '表头行必须是大于等于 1 的数字'})
    col_seq_str = request.form.get('col_seq', '')
    col_hqpn_str = request.form.get('col_hqpn', '')
    col_stype_str = request.form.get('col_stype', '')
    col_qty_str = request.form.get('col_qty', '')
    qty_configs_str = request.form.get('qty_configs', '')
    project_name = request.form.get('project_name', '')

    try:
        wb = _open_workbook(in_path, read_only=True, data_only=True)
    except ValueError as e:
        return jsonify({'success': False, 'error': str(e)})
    sheets = wb.sheetnames
    wb.close()
    if not sheet_name or sheet_name not in sheets:
        sheet_name = sheets[0]

    # Auto-detect if columns not specified
    if not all([col_seq_str, col_hqpn_str, col_stype_str]) or (not col_qty_str and not qty_configs_str):
        wb2 = _open_workbook(in_path, data_only=True)
        ws = wb2[sheet_name]
        detected, raw_headers = _detect_columns(ws, header_row)
        wb2.close()
        if not col_seq_str and 'seq' in detected:
            col_seq_str = get_column_letter(detected['seq'])
        if not col_hqpn_str and 'hq_pn' in detected:
            col_hqpn_str = get_column_letter(detected['hq_pn'])
        if not col_stype_str and 'supply_type' in detected:
            col_stype_str = get_column_letter(detected['supply_type'])
        if not col_qty_str and 'qty' in detected:
            col_qty_str = get_column_letter(detected['qty'])

    col_seq = _col_int(col_seq_str)
    col_hqpn = _col_int(col_hqpn_str)
    col_stype = _col_int(col_stype_str)

    qty_jobs = []
    if qty_configs_str.strip():
        try:
            qty_configs = json.loads(qty_configs_str)
        except Exception:
            return jsonify({'success': False, 'error': '\u7528\u91cf\u914d\u7f6e\u683c\u5f0f\u9519\u8bef'})
        for cfg in qty_configs if isinstance(qty_configs, list) else []:
            col_qty = _col_int((cfg or {}).get('qty_col', ''))
            if not col_qty:
                continue
            qty_project_name = str((cfg or {}).get('name') or '').strip()
            if not qty_project_name:
                qty_project_name = f"\u7528\u91cf{get_column_letter(col_qty)}"
            qty_jobs.append((col_qty, qty_project_name))
    else:
        col_qty_refs = _split_col_refs(col_qty_str)
        if not col_qty_refs and col_qty_str.strip():
            col_qty_refs = [col_qty_str.strip()]
        col_qty_list = [_col_int(ref) for ref in col_qty_refs]
        col_qty_list = [ci for ci in col_qty_list if ci]

        wb_hdr = _open_workbook(in_path, read_only=True, data_only=True)
        ws_hdr = wb_hdr[sheet_name]
        for col_qty in col_qty_list:
            header_val = ws_hdr.cell(row=header_row, column=col_qty).value
            qty_project_name = str(header_val or '').strip() or f"\u7528\u91cf{get_column_letter(col_qty)}"
            if project_name.strip() and len(col_qty_list) == 1:
                qty_project_name = project_name.strip()
            qty_jobs.append((col_qty, qty_project_name))
        wb_hdr.close()

    if not all([col_seq, col_hqpn, col_stype]) or not qty_jobs:
        return jsonify({
            'success': False,
            'error': '\u8bf7\u6307\u5b9a\u6709\u6548\u7684\u5e8f\u53f7\u5217\u3001HQ PN \u5217\u3001\u4e3b\u4e8c\u4f9b\u5217\u3001\u7528\u91cf\u5217\uff08\u53ef\u6dfb\u52a0\u591a\u4e2a\u7528\u91cf\u914d\u7f6e\uff09',
        })

    if len(qty_jobs) == 1:
        col_qty, qty_project_name = qty_jobs[0]
        safe_proj = _safe_filename_part(qty_project_name)
        out_name = f"PLM\u5bfc\u5165_{safe_proj}_{uid}.xlsx"
        out_path = os.path.join(OUTPUT_DIR, out_name)
        total, skipped, skip_logs = _do_convert(
            in_path, sheet_name, header_row,
            col_seq, col_hqpn, col_stype, col_qty, qty_project_name, out_path,
        )
        return jsonify({
            'success': True,
            'download': f'/download/{out_name}',
            'total': total,
            'skipped': skipped,
            'skip_logs': skip_logs,
            'files': [{'name': out_name, 'project_name': qty_project_name,
                       'qty_col': get_column_letter(col_qty),
                       'total': total, 'skipped': skipped}],
        })

    results = []
    all_skip_logs = []
    zip_name = f"PLM\u5bfc\u5165\u6279\u91cf_{uid}.zip"
    zip_path = os.path.join(OUTPUT_DIR, zip_name)
    used_names = set()
    with ZipFile(zip_path, 'w', ZIP_DEFLATED) as zf:
        for col_qty, qty_project_name in qty_jobs:
            safe_proj = _safe_filename_part(qty_project_name)
            out_name = f"PLM\u5bfc\u5165_{safe_proj}_{uid}.xlsx"
            n = 2
            while out_name in used_names:
                out_name = f"PLM\u5bfc\u5165_{safe_proj}_{uid}_{n}.xlsx"
                n += 1
            used_names.add(out_name)

            out_path = os.path.join(OUTPUT_DIR, out_name)
            total, skipped, skip_logs = _do_convert(
                in_path, sheet_name, header_row,
                col_seq, col_hqpn, col_stype, col_qty, qty_project_name, out_path,
            )
            zf.write(out_path, arcname=out_name)
            results.append({
                'name': out_name,
                'project_name': qty_project_name,
                'qty_col': get_column_letter(col_qty),
                'total': total,
                'skipped': skipped,
            })
            all_skip_logs.extend([f"[{qty_project_name}] {msg}" for msg in skip_logs])

    return jsonify({
        'success': True,
        'download': f'/download/{zip_name}',
        'total': sum(r['total'] for r in results),
        'skipped': sum(r['skipped'] for r in results),
        'skip_logs': all_skip_logs,
        'files': results,
        'is_zip': True,
    })

@plm_bp.route('/api/plm/spec_extract', methods=['POST'])
@track_tool_activity('规格型号提取')
def api_spec_extract():
    """提取单列规格型号，去除空格，输出单列 Excel"""
    import json as _json
    f = request.files.get('file')
    if not f:
        return jsonify({'success': False, 'error': '未上传文件'})
    cfg_str = request.form.get('config', '{}')
    try:
        cfg = _json.loads(cfg_str)
    except Exception:
        return jsonify({'success': False, 'error': 'config 格式错误'})

    try:
        header_row = int(cfg.get('header_row', 1))
    except (TypeError, ValueError):
        return jsonify({'success': False, 'error': '表头行必须是大于等于 1 的数字'})
    if header_row < 1:
        return jsonify({'success': False, 'error': '表头行必须是大于等于 1 的数字'})
    sheet_name = cfg.get('sheet_name', '')
    col_name = (cfg.get('col_name') or '').strip()
    exclude_col_name = (cfg.get('exclude_col_name') or cfg.get('hq_col_name') or '').strip()
    if not col_name:
        return jsonify({'success': False, 'error': '未指定提取列'})

    uid = str(uuid.uuid4())[:8]
    try:
        path = _save_uploaded_excel(f, 'se', uid)
    except ValueError as e:
        return jsonify({'success': False, 'error': str(e)})

    try:
        wb = _open_workbook(path, data_only=True)
        sheets = wb.sheetnames
        if not sheet_name or sheet_name not in sheets:
            sheet_name = sheets[0]
        ws = wb[sheet_name]

        headers = [_cell_str(ws.cell(row=header_row, column=ci).value)
                   for ci in range(1, ws.max_column + 1)]
        if col_name not in headers:
            return jsonify({'success': False,
                            'error': f'列 "{col_name}" 不存在，请检查表头行设置'})
        col_idx = headers.index(col_name) + 1  # 1-based
        exclude_col_idx = None
        if exclude_col_name:
            if exclude_col_name not in headers:
                return jsonify({'success': False,
                                'error': f'剔除列 "{exclude_col_name}" 不存在，请检查表头行设置'})
            exclude_col_idx = headers.index(exclude_col_name) + 1

        values = []
        seen_values = set()
        skipped_excluded = 0
        skipped_duplicates = 0
        for ri in range(header_row + 1, ws.max_row + 1):
            if exclude_col_idx is not None and _cell_str(ws.cell(row=ri, column=exclude_col_idx).value):
                skipped_excluded += 1
                continue
            v = _cell_str(ws.cell(row=ri, column=col_idx).value)
            if v:
                cleaned = v.replace(' ', '').replace('\u3000', '')
                if cleaned in seen_values:
                    skipped_duplicates += 1
                    continue
                seen_values.add(cleaned)
                values.append(cleaned)
        wb.close()

        # Write output
        wb_out = Workbook()
        ws_out = wb_out.active
        ws_out.title = '规格型号'
        ws_out.cell(row=1, column=1, value='规格型号').font = Font(bold=True)
        for i, v in enumerate(values, 2):
            ws_out.cell(row=i, column=1, value=v)
        ws_out.column_dimensions['A'].width = 40

        out_name = f'spec_{uid}.xlsx'
        wb_out.save(os.path.join(OUTPUT_DIR, out_name))
        return jsonify({'success': True, 'download': f'/download/{out_name}',
                        'count': len(values), 'skipped_excluded': skipped_excluded,
                        'skipped_duplicates': skipped_duplicates})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})



@plm_bp.route('/api/plm/auto_spec_reverse', methods=['POST'])
@track_tool_activity('PLM规格反查')
def api_auto_spec_reverse():
    """Start a background Playwright PLM spec reverse material job."""
    username = (request.form.get('username') or '').strip()
    password = request.form.get('password') or ''
    f = request.files.get('file')
    single_value = (request.form.get('single_value') or request.form.get('hqpn') or '').strip()
    if not username:
        return jsonify({'success': False, 'error': '请输入账号'})
    if not password:
        return jsonify({'success': False, 'error': '请输入密码'})
    if not f and not single_value:
        return jsonify({'success': False, 'error': '请选择需要上传的 Excel 文件，或输入单个规格型号 / HQ 料号'})

    uid = str(uuid.uuid4())[:8]
    try:
        if f:
            in_path = _save_uploaded_excel(f, 'plm_auto', uid)
            source_label = f.filename or os.path.basename(in_path)
        else:
            in_path, source_label = _create_spec_reverse_single_excel(single_value, uid)
    except ValueError as e:
        return jsonify({'success': False, 'error': str(e)})

    _cleanup_spec_reverse_jobs()
    job_id = _new_spec_reverse_job(source_label)
    _update_spec_reverse_job(job_id, status='queued', stage='已加入查询队列', progress=3)
    _enqueue_spec_reverse_job(job_id, username, password, in_path)
    return jsonify({
        'success': True,
        'job_id': job_id,
        'status_url': f'/api/plm/auto_spec_reverse/status/{job_id}',
        'source_label': source_label,
    })


@plm_bp.route('/api/plm/auto_spec_reverse/status/<job_id>', methods=['GET'])
def api_auto_spec_reverse_status(job_id):
    job = _snapshot_spec_reverse_job(job_id)
    if not job:
        return jsonify({'success': False, 'error': '任务不存在或已过期'}), 404
    return jsonify({
        'success': True,
        'job_id': job['id'],
        'status': job.get('status'),
        'stage': job.get('stage'),
        'progress': job.get('progress'),
        'source_label': job.get('source_label'),
        'download': job.get('download'),
        'filename': job.get('filename'),
        'source_path': job.get('source_path'),
        'error': job.get('error'),
        'log': chr(10).join(job.get('logs') or []),
    })

def _normalize_hq_attachment_value(value):
    text = _cell_str(value).replace('\u3000', ' ')
    if not text:
        return []
    parts = re.split(r'[\s,\uFF0C;\uFF1B/]+', text)
    result = []
    for part in parts:
        hqpn = re.sub(r'\s+', '', str(part or '')).strip().upper()
        if hqpn and re.match(r'^HQ[A-Z0-9_-]{4,}$', hqpn):
            result.append(hqpn)
    return result


def _detect_hq_attachment_column(headers):
    for idx, header in enumerate(headers, 1):
        text = str(header or '').replace(' ', '').upper()
        if 'HQ' in text and ('\u6599\u53f7' in text or 'PN' in text or 'P/N' in text):
            return get_column_letter(idx)
    for idx, header in enumerate(headers, 1):
        text = str(header or '').replace(' ', '').upper()
        if '\u6599\u53f7' in text or text in ('PN', 'P/N'):
            return get_column_letter(idx)
    return ''


def _read_hq_attachment_excel_path(in_path, header_row, sheet_name=''):
    wb = _open_workbook(in_path, data_only=True)
    sheets = wb.sheetnames
    sheet_name = sheet_name if sheet_name in sheets else sheets[0]
    ws = wb[sheet_name]
    headers = [_cell_str(ws.cell(row=header_row, column=ci).value) for ci in range(1, ws.max_column + 1)]
    return wb, ws, sheets, sheet_name, headers


def _read_hq_attachment_excel(file, prefix, uid, header_row, sheet_name=''):
    in_path = _save_uploaded_excel(file, prefix, uid)
    wb = _open_workbook(in_path, data_only=True)
    sheets = wb.sheetnames
    if not sheet_name or sheet_name not in sheets:
        sheet_name = sheets[0] if sheets else ''
    ws = wb[sheet_name]
    headers = [_cell_str(ws.cell(row=header_row, column=ci).value) for ci in range(1, ws.max_column + 1)]
    return wb, ws, sheets, sheet_name, headers


@plm_bp.route('/api/plm/auto_hq_attachments/excel_detect', methods=['POST'])
def api_auto_hq_attachments_excel_detect():
    f = request.files.get('file')
    header_row = _request_int('header_row', 1)
    if header_row is None:
        return jsonify({'success': False, 'error': '\u8868\u5934\u884c\u5fc5\u987b\u662f\u5927\u4e8e\u7b49\u4e8e 1 \u7684\u6570\u5b57'})
    try:
        uid, in_path = _save_or_reuse_uploaded_excel(f, 'plm_att_detect', request.form.get('uid', ''))
        wb, ws, sheets, sheet_name, headers = _read_hq_attachment_excel_path(
            in_path, header_row, request.form.get('sheet_name', '')
        )
        preview = []
        for ri in range(header_row + 1, min(header_row + 21, ws.max_row + 1)):
            row = [_cell_str(ws.cell(row=ri, column=ci).value) for ci in range(1, ws.max_column + 1)]
            if any(row):
                preview.append(row)
        detected_col = _detect_hq_attachment_column(headers)
        wb.close()
        return jsonify({
            'success': True,
            'uid': uid,
            'sheets': sheets,
            'current_sheet': sheet_name,
            'headers': headers,
            'detected': {'hqpn': detected_col} if detected_col else {},
            'preview': preview,
        })
    except ValueError as e:
        return jsonify({'success': False, 'error': str(e)})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})


@plm_bp.route('/api/plm/auto_hq_attachments/batch', methods=['POST'])
@track_tool_activity('PLM\u9644\u4ef6\u6279\u91cf\u4e0b\u8f7d')
def api_auto_hq_attachments_batch():
    username = (request.form.get('username') or '').strip()
    password = request.form.get('password') or ''
    f = request.files.get('file')
    header_row = _request_int('header_row', 1)
    col_hqpn = _col_int(request.form.get('col_hqpn', ''))
    if not username:
        return jsonify({'success': False, 'error': '\u8bf7\u8f93\u5165\u8d26\u53f7'})
    if not password:
        return jsonify({'success': False, 'error': '\u8bf7\u8f93\u5165\u5bc6\u7801'})
    if not f:
        return jsonify({'success': False, 'error': '\u8bf7\u9009\u62e9 Excel \u6587\u4ef6'})
    if header_row is None:
        return jsonify({'success': False, 'error': '\u8868\u5934\u884c\u5fc5\u987b\u662f\u5927\u4e8e\u7b49\u4e8e 1 \u7684\u6570\u5b57'})
    if not col_hqpn:
        return jsonify({'success': False, 'error': '\u8bf7\u9009\u62e9 HQ \u6599\u53f7\u5217'})

    uid = str(uuid.uuid4())[:8]
    try:
        wb, ws, sheets, sheet_name, headers = _read_hq_attachment_excel(
            f, 'plm_att_batch', uid, header_row, request.form.get('sheet_name', '')
        )
        seen = set()
        hqpns = []
        for ri in range(header_row + 1, ws.max_row + 1):
            for hqpn in _normalize_hq_attachment_value(ws.cell(row=ri, column=col_hqpn).value):
                if hqpn not in seen:
                    seen.add(hqpn)
                    hqpns.append(hqpn)
        wb.close()
    except ValueError as e:
        return jsonify({'success': False, 'error': str(e)})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})

    if not hqpns:
        return jsonify({'success': False, 'error': '\u9009\u5b9a\u5217\u6ca1\u6709\u8bc6\u522b\u5230 HQ \u6599\u53f7'})
    if len(hqpns) > 500:
        return jsonify({'success': False, 'error': '\u4e00\u6b21\u6700\u591a\u652f\u6301 500 \u4e2a HQ \u6599\u53f7'})

    _cleanup_attachment_jobs()
    _cleanup_attachment_batches()
    batch_id = uuid.uuid4().hex
    jobs = []
    job_ids = []
    for hqpn in hqpns:
        job_id = _new_attachment_job(hqpn)
        job_ids.append(job_id)
        _update_attachment_job(job_id, status='queued', stage='\u5df2\u52a0\u5165\u4e0b\u8f7d\u961f\u5217', progress=3)
        _enqueue_attachment_job(job_id, username, password, hqpn, batch_id=batch_id)
        jobs.append({
            'hqpn': hqpn,
            'job_id': job_id,
            'status_url': f'/api/plm/auto_hq_attachments/status/{job_id}',
        })
    now = time.time()
    with _PLM_ATTACHMENT_BATCHES_LOCK:
        _PLM_ATTACHMENT_BATCHES[batch_id] = {
            'id': batch_id,
            'job_ids': job_ids,
            'download': '',
            'filename': '',
            'source_path': '',
            'created_at': now,
            'updated_at': now,
        }
    return jsonify({
        'success': True,
        'count': len(jobs),
        'batch_id': batch_id,
        'status_url': f'/api/plm/auto_hq_attachments/batch/status/{batch_id}',
        'jobs': jobs,
    })


@plm_bp.route('/api/plm/auto_hq_attachments/batch/status/<batch_id>', methods=['GET'])
def api_auto_hq_attachments_batch_status(batch_id):
    status = _build_attachment_batch_status(batch_id)
    if not status:
        return jsonify({'success': False, 'error': '\u6279\u91cf\u4efb\u52a1\u4e0d\u5b58\u5728\u6216\u5df2\u8fc7\u671f'}), 404
    status['success'] = True
    return jsonify(status)



@plm_bp.route('/api/plm/auto_hq_attachments/batch/cancel/<batch_id>', methods=['POST'])
def api_auto_hq_attachments_batch_cancel(batch_id):
    status = _cancel_attachment_batch(batch_id)
    if not status:
        return jsonify({'success': False, 'error': '\u6279\u91cf\u4efb\u52a1\u4e0d\u5b58\u5728\u6216\u5df2\u8fc7\u671f'}), 404
    status['success'] = True
    return jsonify(status)

@plm_bp.route('/api/plm/auto_hq_attachments', methods=['POST'])
@track_tool_activity('PLM\u9644\u4ef6\u4e0b\u8f7d')
def api_auto_hq_attachments():
    """Start a background PLM attachment download job."""
    username = (request.form.get('username') or '').strip()
    password = request.form.get('password') or ''
    hqpn = (request.form.get('hqpn') or '').strip()
    if not username:
        return jsonify({'success': False, 'error': '\u8bf7\u8f93\u5165\u8d26\u53f7'})
    if not password:
        return jsonify({'success': False, 'error': '\u8bf7\u8f93\u5165\u5bc6\u7801'})
    if not hqpn:
        return jsonify({'success': False, 'error': '\u8bf7\u8f93\u5165 HQ \u6599\u53f7'})

    _cleanup_attachment_jobs()
    job_id = _new_attachment_job(hqpn)

    _update_attachment_job(job_id, status='queued', stage='\u5df2\u52a0\u5165\u4e0b\u8f7d\u961f\u5217', progress=3)
    _enqueue_attachment_job(job_id, username, password, hqpn)
    return jsonify({'success': True, 'job_id': job_id, 'status_url': f'/api/plm/auto_hq_attachments/status/{job_id}'})


@plm_bp.route('/api/plm/auto_hq_attachments/status/<job_id>', methods=['GET'])
def api_auto_hq_attachments_status(job_id):
    job = _snapshot_attachment_job(job_id)
    if not job:
        return jsonify({'success': False, 'error': '\u4efb\u52a1\u4e0d\u5b58\u5728\u6216\u5df2\u8fc7\u671f'}), 404
    return jsonify({
        'success': True,
        'job_id': job['id'],
        'status': job.get('status'),
        'stage': job.get('stage'),
        'progress': job.get('progress'),
        'hqpn': job.get('hqpn'),
        'download': job.get('download'),
        'filename': job.get('filename'),
        'source_path': job.get('source_path'),
        'error': job.get('error'),
        'log': chr(10).join(job.get('logs') or []),
    })
from . import customer_hq  # noqa: E402,F401
