#$ db_helper.py
#% SQLite 連線管理，統一由此進入，不再呼叫 fixture_helper

import sqlite3
from contextlib import contextmanager

DB_PATH = None
_PRAGMA_INIT_DONE = set()
_CHECKPOINT_MODES = {"PASSIVE", "FULL", "RESTART", "TRUNCATE"}


def set_db_path(path: str):
    global DB_PATH
    DB_PATH = path


def _ensure_pragmas(conn: sqlite3.Connection, target: str):
    if target not in _PRAGMA_INIT_DONE:
        conn.execute("PRAGMA journal_mode=WAL;")
        conn.execute("PRAGMA synchronous=NORMAL;")
        _PRAGMA_INIT_DONE.add(target)
    conn.execute("PRAGMA foreign_keys=ON;")
    conn.execute("PRAGMA busy_timeout=5000;")


def get_conn(path: str = None):
    target = path or DB_PATH
    if not target:
        raise ValueError("DB_PATH 尚未設定，請先呼叫 set_db_path()")
    conn = sqlite3.connect(target, isolation_level=None)
    _ensure_pragmas(conn, target)
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
    with tx() as conn:
        cur = conn.cursor()
        cur.execute(sql, params)


def db_executemany(sql: str, seq):
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
    m = (mode or "PASSIVE").strip().upper()
    if m not in _CHECKPOINT_MODES:
        raise ValueError("wal_checkpoint mode 不合法")
    with conn_ctx() as conn:
        cur = conn.cursor()
        cur.execute(f"PRAGMA wal_checkpoint({m});")
        return cur.fetchall()