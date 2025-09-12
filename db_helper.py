# db_helper.py
#% SQLite 連線管理，統一由此進入，不再呼叫 fixture_helper
import sqlite3

DB_PATH = None

def set_db_path(path: str):
    global DB_PATH
    DB_PATH = path

def get_conn(path: str = None):
    """取得資料庫連線，預設使用 DB_PATH"""
    target = path or DB_PATH
    if not target:
        raise ValueError("DB_PATH 尚未設定，請先呼叫 set_db_path()")
    conn = sqlite3.connect(target, timeout=30, isolation_level=None)
    conn.execute("PRAGMA journal_mode=WAL;")
    conn.execute("PRAGMA foreign_keys = ON;")
    return conn

def tx():
    return get_conn()

def db_execute(sql: str, params: tuple = ()):
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute(sql, params)
        conn.commit()

def db_executemany(sql: str, seq):
    with get_conn() as conn:
        cur = conn.cursor()
        cur.executemany(sql, seq)
        conn.commit()

def db_query_all(sql: str, params: tuple = ()):
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute(sql, params)
        return cur.fetchall()

def db_query_one(sql: str, params: tuple = ()):
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute(sql, params)
        return cur.fetchone()

def safe_checkpoint(mode: str = "PASSIVE"):
    return