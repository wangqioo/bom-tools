# -*- coding: utf-8 -*-
"""BOM Tools Web entry."""

import json
import os
import sys
import threading
import time
import warnings

from flask import Flask, abort, render_template, send_file
from werkzeug.utils import safe_join

from shared import (
    UPLOAD_DIR,
    OUTPUT_DIR,
    CACHE_DIR,
    FEISHU_PRESET_TABLES,
    PLATFORM_VERSION,
    TOOL_VERSIONS,
    _cleanup_old_files,
    _file_belongs_to_current_user,
    _register_output_file,
)
from auth import auth_bp, current_user, init_auth_storage, require_login

_DEFAULT_CONFIG_PATH = os.path.join(os.path.dirname(__file__), 'default_config.json')
try:
    with open(_DEFAULT_CONFIG_PATH, 'r', encoding='utf-8') as _f:
        _DEFAULT_CONFIG = json.load(_f)
except Exception:
    _DEFAULT_CONFIG = {}

from bom import bom_bp
from feishu import feishu_bp
from plm import plm_bp
from bom_compare import bom_compare_bp
from bom_compare.generic_free import free_bom_compare_bp
from bug_report import bug_report_bp
from feature_request import feature_request_bp
from manufacturer_alias import manufacturer_alias_bp
from bom_checklist import bom_checklist_bp

app = Flask(__name__)
_SECRET_KEY_PATH = os.path.join(os.path.dirname(__file__), 'auth_data', 'flask_secret_key')


def _load_secret_key():
    configured = os.environ.get('BOM_TOOLS_SECRET_KEY')
    if configured:
        return configured
    os.makedirs(os.path.dirname(_SECRET_KEY_PATH), exist_ok=True)
    try:
        if os.path.exists(_SECRET_KEY_PATH):
            with open(_SECRET_KEY_PATH, 'rb') as f:
                key = f.read().strip()
            if key:
                return key
        key = os.urandom(32).hex()
        with open(_SECRET_KEY_PATH, 'w', encoding='ascii') as f:
            f.write(key)
        return key
    except OSError as exc:
        warnings.warn(f'Failed to persist Flask secret key: {exc}', RuntimeWarning)
        return os.urandom(32)


app.secret_key = _load_secret_key()
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB
app.config['AUTH_REQUIRED'] = (
    os.environ.get('BOM_TOOLS_AUTH_REQUIRED', '').lower() not in ('0', 'false', 'no')
    and 'unittest' not in sys.modules
)

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(CACHE_DIR, exist_ok=True)
init_auth_storage()

app.register_blueprint(auth_bp)
app.register_blueprint(bom_bp)
app.register_blueprint(feishu_bp)
app.register_blueprint(plm_bp)
app.register_blueprint(bom_compare_bp)
app.register_blueprint(free_bom_compare_bp)
app.register_blueprint(bug_report_bp)
app.register_blueprint(feature_request_bp)
app.register_blueprint(manufacturer_alias_bp)
app.register_blueprint(bom_checklist_bp)


@app.before_request
def _require_login():
    return require_login()


@app.route('/')
def index():
    return render_template(
        'index.html',
        preset_tables=FEISHU_PRESET_TABLES,
        default_config=_DEFAULT_CONFIG,
        current_user=current_user(),
        platform_version=PLATFORM_VERSION,
        tool_versions=TOOL_VERSIONS,
    )


@app.route('/download/<filename>')
def download(filename):
    path = safe_join(OUTPUT_DIR, filename)
    if not path or not os.path.exists(path) or not _file_belongs_to_current_user(path):
        abort(404)
    return send_file(path, as_attachment=True, download_name=filename)


@app.after_request
def _record_response_download_owners(response):
    """Assign ownership to synchronous exports returned by existing tool APIs."""
    if not response.is_json:
        return response
    try:
        payload = response.get_json(silent=True) or {}
    except Exception:
        return response

    def visit(value):
        if isinstance(value, dict):
            for child in value.values():
                visit(child)
        elif isinstance(value, list):
            for child in value:
                visit(child)
        elif isinstance(value, str) and value.startswith('/download/'):
            filename = value[len('/download/'):]
            path = safe_join(OUTPUT_DIR, filename)
            if path:
                _register_output_file(path)

    visit(payload)
    return response


def _cleanup_job():
    while True:
        time.sleep(600)
        _cleanup_old_files(UPLOAD_DIR, 30)
        _cleanup_old_files(OUTPUT_DIR, 30)
        _cleanup_old_files(CACHE_DIR, 8 * 60)


threading.Thread(target=_cleanup_job, daemon=True).start()

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False, threaded=True)
