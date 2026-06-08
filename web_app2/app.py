# -*- coding: utf-8 -*-
"""BOM Tools Web v2.0 entry."""

import json
import os
import sys
import threading
import time

from flask import Flask, abort, render_template, send_file
from werkzeug.utils import safe_join

from shared import UPLOAD_DIR, OUTPUT_DIR, CACHE_DIR, FEISHU_PRESET_TABLES, _cleanup_old_files
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
from bug_report import bug_report_bp
from feature_request import feature_request_bp
from manufacturer_alias import manufacturer_alias_bp

app = Flask(__name__)
app.secret_key = os.environ.get('BOM_TOOLS_SECRET_KEY') or os.urandom(24)
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
app.register_blueprint(bug_report_bp)
app.register_blueprint(feature_request_bp)
app.register_blueprint(manufacturer_alias_bp)


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
    )


@app.route('/download/<filename>')
def download(filename):
    path = safe_join(OUTPUT_DIR, filename)
    if not path or not os.path.exists(path):
        abort(404)
    return send_file(path, as_attachment=True, download_name=filename)


def _cleanup_job():
    while True:
        time.sleep(600)
        _cleanup_old_files(UPLOAD_DIR, 30)
        _cleanup_old_files(OUTPUT_DIR, 30)
        # Feishu cache files are refreshed manually by users.


threading.Thread(target=_cleanup_job, daemon=True).start()

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False, threaded=True)
