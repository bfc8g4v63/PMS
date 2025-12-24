#$ db_helper.py
#% SQLite 連線管理，統一由此進入，不再呼叫 fixture_helper

import os
import sqlite3
from urllib.parse import quote
from contextlib import contextmanager

DB_PATH = None
DB_ROLE = "admin"
_PRAGMA_INIT_DONE = set()
_CHECKPOINT_MODES = {"PASSIVE", "FULL", "RESTART", "TRUNCATE"}

def set_db_path(path: str):
    global DB_PATH
    DB_PATH = path

def set_db_role(role: str):
    global DB_ROLE
    r = (role or "").strip().lower()
    DB_ROLE = r or "admin"

def _ensure_pragmas_rw(conn: sqlite3.Connection, target: str):
    key = f"rw::{target}"
    if key not in _PRAGMA_INIT_DONE:
        conn.execute("PRAGMA journal_mode=WAL;")
        conn.execute("PRAGMA synchronous=NORMAL;")
        _PRAGMA_INIT_DONE.add(key)
    conn.execute("PRAGMA foreign_keys=ON;")
    conn.execute("PRAGMA busy_timeout=5000;")

def _ensure_pragmas_ro(conn: sqlite3.Connection, target: str):
    key = f"ro::{target}"
    if key not in _PRAGMA_INIT_DONE:
        _PRAGMA_INIT_DONE.add(key)
    conn.execute("PRAGMA foreign_keys=ON;")
    conn.execute("PRAGMA busy_timeout=5000;")

def _open_ro(target: str):
    if not os.path.exists(target):
        raise FileNotFoundError(f"找不到資料庫檔案：{target}")
    uri = "file:" + quote(target) + "?mode=ro"
    return sqlite3.connect(uri, uri=True, isolation_level=None)

def _open_rw(target: str):
    return sqlite3.connect(target, isolation_level=None)

def get_conn(path: str = None):
    target = path or DB_PATH
    if not target:
        raise ValueError("DB_PATH 尚未設定，請先呼叫 set_db_path()")

    if DB_ROLE == "leader":
        conn = _open_ro(target)
        _ensure_pragmas_ro(conn, target)
        return conn

    conn = _open_rw(target)
    _ensure_pragmas_rw(conn, target)
    return conn

@contextmanager
def conn_ctx(path: str = None):
    conn = get_conn(path)
    try:
        yield conn
    finally:
        conn.close()

@contextmanager
def tx(path: str = None, begin_mode: str = "IMMEDIATE"):
    if DB_ROLE == "leader":
        raise PermissionError("leader 角色禁止寫入交易")

    conn = get_conn(path)
    try:
        bm = (begin_mode or "IMMEDIATE").strip().upper()
        if bm not in ("DEFERRED", "IMMEDIATE", "EXCLUSIVE"):
            bm = "IMMEDIATE"
        conn.execute(f"BEGIN {bm};")
        yield conn
        conn.commit()
    except:
        conn.rollback()
        raise
    finally:
        conn.close()

def db_execute(sql: str, params: tuple = ()):
    if DB_ROLE == "leader":
        raise PermissionError("leader 角色禁止寫入")
    with tx() as conn:
        cur = conn.cursor()
        cur.execute(sql, params)

def db_executemany(sql: str, seq):
    if DB_ROLE == "leader":
        raise PermissionError("leader 角色禁止寫入")
    with tx() as conn:
        cur = conn.cursor()
        cur.executemany(sql, seq)

def db_query_all(sql: str, params: tuple = ()):
    with conn_ctx() as conn:
        cur = conn.cursor()
        cur.execute(sql, params)
        return cur.fetchall()

def db_query_one(sql: str, params: tuple = ()):
    with conn_ctx() as conn:
        cur = conn.cursor()
        cur.execute(sql, params)
        return cur.fetchone()

def safe_checkpoint(mode: str = "PASSIVE"):
    if DB_ROLE == "leader":
        raise PermissionError("leader 角色禁止執行 checkpoint")

    m = (mode or "PASSIVE").strip().upper()
    if m not in _CHECKPOINT_MODES:
        raise ValueError("wal_checkpoint mode 不合法")
    with conn_ctx() as conn:
        cur = conn.cursor()
        cur.execute(f"PRAGMA wal_checkpoint({m});")
        return cur.fetchall()