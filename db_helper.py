#$ db_helper.py
#% SQLite 連線管理，統一由此進入，不再呼叫 fixture_helper

import sqlite3
from contextlib import contextmanager

DB_PATH = None

def set_db_path(path: str):
    global DB_PATH
    DB_PATH = path

def get_conn(path: str = None):
    target = path or DB_PATH
    if not target:
        raise ValueError("DB_PATH 尚未設定，請先呼叫 set_db_path()")
    conn = sqlite3.connect(target, timeout=5, isolation_level=None)
    conn.execute("PRAGMA journal_mode=WAL;")
    conn.execute("PRAGMA synchronous=NORMAL;")
    conn.execute("PRAGMA foreign_keys = ON;")
    conn.execute("PRAGMA busy_timeout=5000;")
    return conn

@contextmanager
def tx(path: str = None):
    conn = get_conn(path)
    try:
        yield conn
        conn.commit()
    except:
        conn.rollback()
        raise
    finally:
        conn.close()

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
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute(f"PRAGMA wal_checkpoint({mode});")
        return cur.fetchall()