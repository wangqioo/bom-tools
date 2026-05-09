# -*- coding: utf-8 -*-
"""BOM Tools Web v2.0 — 全新清洁版入口"""

import os, json, threading, time
from flask import Flask, render_template, send_file
from shared import UPLOAD_DIR, OUTPUT_DIR, CACHE_DIR, FEISHU_PRESET_TABLES, _cleanup_old_files

_DEFAULT_CONFIG_PATH = os.path.join(os.path.dirname(__file__), 'default_config.json')
try:
    with open(_DEFAULT_CONFIG_PATH, 'r', encoding='utf-8') as _f:
        _DEFAULT_CONFIG = json.load(_f)
except Exception:
    _DEFAULT_CONFIG = {}

from bom import bom_bp
from feishu import feishu_bp
from plm import plm_bp

app = Flask(__name__)
app.secret_key = os.urandom(24)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(CACHE_DIR,  exist_ok=True)

app.register_blueprint(bom_bp)
app.register_blueprint(feishu_bp)
app.register_blueprint(plm_bp)


@app.route('/')
def index():
    return render_template('index.html', preset_tables=FEISHU_PRESET_TABLES,
                           default_config=_DEFAULT_CONFIG)


@app.route('/download/<filename>')
def download(filename):
    path = os.path.join(OUTPUT_DIR, filename)
    if not os.path.exists(path):
        return "文件不存在或已过期", 404
    return send_file(path, as_attachment=True, download_name=filename)


def _cleanup_job():
    while True:
        time.sleep(600)
        _cleanup_old_files(UPLOAD_DIR, 30)
        _cleanup_old_files(OUTPUT_DIR, 30)
        # 飞书缓存文件不自动清理，由用户手动刷新


threading.Thread(target=_cleanup_job, daemon=True).start()

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
