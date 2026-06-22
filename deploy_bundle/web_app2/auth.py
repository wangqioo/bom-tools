# -*- coding: utf-8 -*-
"""Employee-id based authentication and admin helpers for BOM Tools."""

import json
import os
import re
import sqlite3
import time

from flask import Blueprint, current_app, redirect, render_template, session, url_for

from shared import jsonify, request


auth_bp = Blueprint("auth", __name__)

BASE_DIR = os.path.dirname(__file__)
AUTH_DIR = os.environ.get("BOM_TOOLS_AUTH_DATA_DIR") or os.path.join(BASE_DIR, "auth_data")
DB_PATH = os.path.join(AUTH_DIR, "users.sqlite3")
DEFAULT_ADMIN_EMPLOYEE_IDS = {"ADMIN"}
EMPLOYEE_ID_PATTERN = re.compile(r"^[A-Za-z0-9_-]{2,40}$")


def auth_enabled():
    return bool(current_app.config.get("AUTH_REQUIRED", True))


def admin_employee_ids():
    configured = os.environ.get("BOM_TOOLS_ADMIN_EMPLOYEE_IDS", "")
    values = {item.strip().upper() for item in configured.replace(";", ",").split(",") if item.strip()}
    return DEFAULT_ADMIN_EMPLOYEE_IDS | values


def _ensure_dirs():
    os.makedirs(AUTH_DIR, exist_ok=True)


def _connect():
    _ensure_dirs()
    conn = sqlite3.connect(DB_PATH, timeout=15)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA busy_timeout=15000")
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS users (
            id TEXT PRIMARY KEY,
            username TEXT NOT NULL UNIQUE,
            password_hash TEXT NOT NULL,
            display_name TEXT NOT NULL,
            employee_id TEXT NOT NULL,
            role TEXT NOT NULL,
            is_active INTEGER NOT NULL DEFAULT 1,
            created_at REAL NOT NULL,
            last_login_at REAL
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS user_activity (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id TEXT NOT NULL,
            employee_id TEXT NOT NULL,
            display_name TEXT NOT NULL,
            action TEXT NOT NULL,
            target_type TEXT NOT NULL,
            target_id TEXT NOT NULL,
            detail TEXT NOT NULL,
            created_at REAL NOT NULL
        )
        """
    )
    _ensure_schema(conn)
    _seed_admin(conn)
    conn.commit()
    return conn


def _ensure_schema(conn):
    conn.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_users_employee_id ON users(employee_id)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_user_activity_user_id ON user_activity(user_id)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_user_activity_created_at ON user_activity(created_at)")


def _seed_admin(conn):
    row = conn.execute("SELECT id FROM users WHERE employee_id = ?", ("ADMIN",)).fetchone()
    if row:
        conn.execute("UPDATE users SET role = 'admin', display_name = '系统管理员' WHERE employee_id = 'ADMIN'")
        return
    conn.execute(
        """
        INSERT INTO users (
            id, username, password_hash, display_name, employee_id,
            role, is_active, created_at, last_login_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        ("emp-ADMIN", "ADMIN", "", "系统管理员", "ADMIN", "admin", 1, time.time(), None),
    )


def init_auth_storage():
    conn = _connect()
    conn.close()


def _normalize_employee_id(value):
    employee_id = str(value or "").strip().upper()
    if not employee_id:
        raise ValueError("请输入工号")
    if not EMPLOYEE_ID_PATTERN.match(employee_id):
        raise ValueError("工号只能包含字母、数字、下划线或短横线，长度 2-40 位")
    return employee_id


def _role_for_employee_id(employee_id):
    return "admin" if employee_id.upper() in admin_employee_ids() else "user"


def _row_to_user(row):
    if not row:
        return None
    return {
        "id": row["id"],
        "username": row["username"],
        "display_name": row["display_name"],
        "employee_id": row["employee_id"],
        "role": row["role"],
        "is_active": bool(row["is_active"]),
    }


def get_user_by_id(user_id):
    conn = _connect()
    try:
        row = conn.execute(
            "SELECT * FROM users WHERE id = ? AND is_active = 1",
            (str(user_id or "").strip(),),
        ).fetchone()
        return _row_to_user(row)
    finally:
        conn.close()


def get_or_create_user_by_employee_id(value, display_name=""):
    employee_id = _normalize_employee_id(value)
    role = _role_for_employee_id(employee_id)
    conn = _connect()
    try:
        row = conn.execute(
            "SELECT * FROM users WHERE employee_id = ? AND is_active = 1",
            (employee_id,),
        ).fetchone()
        if row:
            if row["role"] != role:
                conn.execute("UPDATE users SET role = ? WHERE id = ?", (role, row["id"]))
                conn.commit()
                row = conn.execute("SELECT * FROM users WHERE id = ?", (row["id"],)).fetchone()
            return _row_to_user(row), False

        display_name = str(display_name or "").strip()[:80]
        if not display_name:
            raise LookupError("首次注册请填写姓名")
        user_id = f"emp-{employee_id}"
        now = time.time()
        conn.execute(
            """
            INSERT INTO users (
                id, username, password_hash, display_name, employee_id,
                role, is_active, created_at, last_login_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (user_id, employee_id, "", display_name, employee_id, role, 1, now, None),
        )
        conn.commit()
        row = conn.execute("SELECT * FROM users WHERE id = ?", (user_id,)).fetchone()
        return _row_to_user(row), True
    finally:
        conn.close()


def set_last_login(user_id):
    conn = _connect()
    try:
        conn.execute("UPDATE users SET last_login_at = ? WHERE id = ?", (time.time(), user_id))
        conn.commit()
    finally:
        conn.close()


def current_user():
    if not auth_enabled():
        return {
            "id": "test-admin",
            "username": "test-admin",
            "display_name": "测试管理员",
            "employee_id": "TEST",
            "role": "admin",
            "is_active": True,
        }
    cached = session.get("user")
    if cached and cached.get("id"):
        return cached
    user_id = session.get("user_id")
    if not user_id:
        return None
    user = get_user_by_id(user_id)
    if user:
        session["user"] = user
    return user


def is_admin():
    user = current_user()
    return bool(user and user.get("role") == "admin")


def require_admin_json():
    if not auth_enabled():
        return None
    if not current_user():
        return jsonify({"success": False, "error": "请先登录"}), 401
    if not is_admin():
        return jsonify({"success": False, "error": "仅管理员可以执行该操作"}), 403
    return None


def record_activity(action, target_type="", target_id="", detail=None, user=None):
    user = user or current_user()
    if not user:
        return
    conn = _connect()
    try:
        conn.execute(
            """
            INSERT INTO user_activity (
                user_id, employee_id, display_name, action, target_type,
                target_id, detail, created_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                user.get("id", ""),
                user.get("employee_id", ""),
                user.get("display_name", ""),
                str(action or ""),
                str(target_type or ""),
                str(target_id or ""),
                json.dumps(detail or {}, ensure_ascii=False),
                time.time(),
            ),
        )
        conn.commit()
    finally:
        conn.close()


def _activity_summary_by_user(conn):
    rows = conn.execute(
        """
        SELECT user_id,
               SUM(CASE WHEN action IN ('login', 'tool_run', 'tool_export') THEN 1 ELSE 0 END) AS activity_count,
               SUM(CASE WHEN action = 'login' THEN 1 ELSE 0 END) AS login_count,
               SUM(CASE WHEN action = 'tool_run' THEN 1 ELSE 0 END) AS tool_run_count,
               SUM(CASE WHEN action = 'tool_export' THEN 1 ELSE 0 END) AS tool_export_count,
               MAX(CASE WHEN action IN ('login', 'tool_run', 'tool_export') THEN created_at ELSE NULL END) AS last_activity_at
        FROM user_activity
        GROUP BY user_id
        """
    ).fetchall()
    return {row["user_id"]: dict(row) for row in rows}


def _admin_user_payload(row, activity):
    item = _row_to_user(row)
    item.update({
        "created_at": row["created_at"],
        "last_login_at": row["last_login_at"],
    })
    stats = activity.get(row["id"], {})
    item["activity_count"] = int(stats.get("activity_count") or 0)
    item["login_count"] = int(stats.get("login_count") or 0)
    item["tool_run_count"] = int(stats.get("tool_run_count") or 0)
    item["tool_export_count"] = int(stats.get("tool_export_count") or 0)
    item["last_activity_at"] = stats.get("last_activity_at")
    return item


def _wants_json():
    return request.path.startswith("/api/") or "application/json" in (request.headers.get("Accept") or "")


def require_login():
    if not auth_enabled():
        return None
    allowed_paths = {"/login", "/api/login", "/api/logout", "/api/me"}
    if request.path in allowed_paths:
        return None
    if request.path.startswith("/static/"):
        return None
    if current_user():
        return None
    if _wants_json():
        return jsonify({"success": False, "error": "请先登录"}), 401
    return redirect(url_for("auth.login_page", next=request.full_path if request.query_string else request.path))


@auth_bp.route("/api/admin/users", methods=["GET"])
def api_admin_users():
    denied = require_admin_json()
    if denied:
        return denied
    query = str(request.args.get("q") or "").strip().lower()
    conn = _connect()
    try:
        rows = conn.execute("SELECT * FROM users ORDER BY created_at DESC").fetchall()
        activity = _activity_summary_by_user(conn)
        users = [_admin_user_payload(row, activity) for row in rows]
        if query:
            users = [
                user for user in users
                if query in (user.get("employee_id", "") + " " + user.get("display_name", "") + " " + user.get("role", "")).lower()
            ]
        summary = {
            "total": len(users),
            "active": sum(1 for user in users if user.get("is_active")),
            "disabled": sum(1 for user in users if not user.get("is_active")),
            "admins": sum(1 for user in users if user.get("role") == "admin"),
            "normal_users": sum(1 for user in users if user.get("role") != "admin"),
        }
        return jsonify({"success": True, "users": users, "summary": summary})
    finally:
        conn.close()


@auth_bp.route("/api/admin/users/<user_id>/role", methods=["POST"])
def api_admin_user_role(user_id):
    denied = require_admin_json()
    if denied:
        return denied
    data = request.get_json(silent=True) or {}
    role = str(data.get("role") or "").strip()
    if role not in ("admin", "user"):
        return jsonify({"success": False, "error": "无效的用户角色"}), 400
    conn = _connect()
    try:
        row = conn.execute("SELECT * FROM users WHERE id = ?", (user_id,)).fetchone()
        if not row:
            return jsonify({"success": False, "error": "用户不存在"}), 404
        if row["employee_id"] == "ADMIN" and role != "admin":
            return jsonify({"success": False, "error": "内置 ADMIN 不能降级"}), 400
        conn.execute("UPDATE users SET role = ? WHERE id = ?", (role, user_id))
        conn.commit()
        record_activity("admin_update_user_role", "user", user_id, {"role": role})
        updated = conn.execute("SELECT * FROM users WHERE id = ?", (user_id,)).fetchone()
        return jsonify({"success": True, "user": _admin_user_payload(updated, _activity_summary_by_user(conn))})
    finally:
        conn.close()


@auth_bp.route("/api/admin/users/<user_id>/active", methods=["POST"])
def api_admin_user_active(user_id):
    denied = require_admin_json()
    if denied:
        return denied
    data = request.get_json(silent=True) or {}
    is_active = bool(data.get("is_active"))
    conn = _connect()
    try:
        row = conn.execute("SELECT * FROM users WHERE id = ?", (user_id,)).fetchone()
        if not row:
            return jsonify({"success": False, "error": "用户不存在"}), 404
        if row["employee_id"] == "ADMIN" and not is_active:
            return jsonify({"success": False, "error": "内置 ADMIN 不能禁用"}), 400
        conn.execute("UPDATE users SET is_active = ? WHERE id = ?", (1 if is_active else 0, user_id))
        conn.commit()
        record_activity("admin_update_user_active", "user", user_id, {"is_active": is_active})
        updated = conn.execute("SELECT * FROM users WHERE id = ?", (user_id,)).fetchone()
        return jsonify({"success": True, "user": _admin_user_payload(updated, _activity_summary_by_user(conn))})
    finally:
        conn.close()


@auth_bp.route("/api/admin/activity", methods=["GET"])
def api_admin_activity():
    denied = require_admin_json()
    if denied:
        return denied
    raw_limit = str(request.args.get("limit") or "").strip()
    limit = min(max(int(raw_limit), 1), 1000) if raw_limit else None
    conn = _connect()
    try:
        if limit:
            rows = conn.execute("SELECT * FROM user_activity ORDER BY created_at DESC LIMIT ?", (limit,)).fetchall()
        else:
            rows = conn.execute("SELECT * FROM user_activity ORDER BY created_at DESC").fetchall()
        activities = []
        for row in rows:
            item = dict(row)
            try:
                item["detail"] = json.loads(item.get("detail") or "{}")
            except Exception:
                item["detail"] = {}
            activities.append(item)
        return jsonify({"success": True, "activities": activities})
    finally:
        conn.close()


@auth_bp.route("/login", methods=["GET"])
def login_page():
    if current_user():
        return redirect(url_for("index"))
    return render_template("login.html")


@auth_bp.route("/api/login", methods=["POST"])
def api_login():
    data = request.get_json(silent=True) or {}
    employee_id = data.get("employee_id") or data.get("username") or request.form.get("employee_id") or request.form.get("username")
    display_name = data.get("display_name") or data.get("name") or request.form.get("display_name") or request.form.get("name") or ""
    try:
        user, created = get_or_create_user_by_employee_id(employee_id, display_name=display_name)
    except LookupError as exc:
        return jsonify({"success": False, "need_name": True, "error": str(exc)}), 409
    except ValueError as exc:
        return jsonify({"success": False, "error": str(exc)}), 400
    session.clear()
    session["user_id"] = user["id"]
    session["user"] = user
    set_last_login(user["id"])
    record_activity("login", "auth", user["id"], {"created": created}, user=user)
    return jsonify({"success": True, "user": user, "created": created})


@auth_bp.route("/api/logout", methods=["POST"])
def api_logout():
    session.clear()
    return jsonify({"success": True})


@auth_bp.route("/api/me", methods=["GET"])
def api_me():
    user = current_user()
    return jsonify({"success": True, "user": user, "authenticated": bool(user)})

