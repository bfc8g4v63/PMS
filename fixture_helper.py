import sqlite3
from datetime import datetime
import os

DB_PATH = None

CORE_WAREHOUSES = [
    "治具室",
    "上齊",
    "睿均",
    "捷暉",
    "立榮",
    "華勤",
    "上貿",
    "麥博",
    "信利",
    "GC",
    "工程",
    "不良品"
]

EXCEPTIONS = set()

def set_db_path(path: str):
    global DB_PATH
    DB_PATH = path

def get_conn():
    if DB_PATH is None:
        raise ValueError("DB_PATH 尚未設定，請先呼叫 set_db_path()")
    conn = sqlite3.connect(DB_PATH, timeout=30, isolation_level=None)
    conn.execute("PRAGMA journal_mode=WAL;")
    conn.execute("PRAGMA foreign_keys = ON;")
    return conn

def add_col_if_missing(conn, table: str, col: str, col_type: str):
    cur = conn.cursor()
    cur.execute(f"PRAGMA table_info({table})")
    cols = [row[1] for row in cur.fetchall()]
    if col not in cols:
        cur.execute(f"ALTER TABLE {table} ADD COLUMN {col} {col_type}")
    cur.close()

def ensure_schemas():
    with get_conn() as conn:
        c = conn.cursor()
        c.execute("""
        CREATE TABLE IF NOT EXISTS fixtures (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            part_no TEXT UNIQUE,
            description TEXT,
            spec TEXT,
            category TEXT,
            unit_price REAL DEFAULT 0,
            safety_stock INTEGER DEFAULT 0,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
        """)
        add_col_if_missing(conn, "fixtures", "description", "TEXT")
        add_col_if_missing(conn, "fixtures", "spec", "TEXT")
        add_col_if_missing(conn, "fixtures", "category", "TEXT")
        add_col_if_missing(conn, "fixtures", "unit_price", "REAL DEFAULT 0")
        add_col_if_missing(conn, "fixtures", "safety_stock", "INTEGER DEFAULT 0")
        add_col_if_missing(conn, "fixtures", "created_at", "TIMESTAMP DEFAULT CURRENT_TIMESTAMP")

        c.execute("""
        CREATE TABLE IF NOT EXISTS warehouse_stock (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            part_no TEXT,
            warehouse TEXT,
            qty INTEGER DEFAULT 0,
            UNIQUE(part_no, warehouse)
        )
        """)
        add_col_if_missing(conn, "warehouse_stock", "part_no", "TEXT")
        add_col_if_missing(conn, "warehouse_stock", "warehouse", "TEXT")
        add_col_if_missing(conn, "warehouse_stock", "qty", "INTEGER DEFAULT 0")

        c.execute("""
        CREATE TABLE IF NOT EXISTS transfer_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            part_no TEXT,
            from_warehouse TEXT,
            to_warehouse TEXT,
            qty INTEGER,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
        """)

        c.execute("""
        CREATE TABLE IF NOT EXISTS fixture_boms (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            parent_part_no TEXT,
            child_part_no TEXT,
            qty INTEGER DEFAULT 1,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
        """)
        add_col_if_missing(conn, "fixture_boms", "parent_part_no", "TEXT")
        add_col_if_missing(conn, "fixture_boms", "child_part_no", "TEXT")
        add_col_if_missing(conn, "fixture_boms", "qty", "INTEGER DEFAULT 1")
        add_col_if_missing(conn, "fixture_boms", "created_at", "TIMESTAMP DEFAULT CURRENT_TIMESTAMP")

        conn.commit()
        c.close()

def is_consumable_part(part_no: str) -> bool:
    return part_no.startswith("939") and part_no not in EXCEPTIONS

def fixture_exists(part_no: str) -> bool:
    with get_conn() as conn:
        c = conn.cursor()
        c.execute("SELECT 1 FROM fixtures WHERE part_no=?", (part_no,))
        ok = c.fetchone() is not None
        c.close()
        return ok


def insert_fixture(part_no: str, name: str = "", spec: str = "", category: str = "",
                   unit_price: float = 0.0, safety_stock: int = 0):
    with get_conn() as conn:
        c = conn.cursor()
        c.execute("SELECT 1 FROM fixtures WHERE part_no=?", (part_no,))
        if c.fetchone():
            raise ValueError(f"治具料號 {part_no} 已存在，不可重複建立")
        c.execute(
            "INSERT INTO fixtures (part_no, description, spec, category, unit_price, safety_stock) VALUES (?, ?, ?, ?, ?, ?)",
            (part_no, name, spec, category, unit_price, safety_stock),
        )
        for wh in CORE_WAREHOUSES:
            c.execute(
                "INSERT OR IGNORE INTO warehouse_stock (part_no, warehouse, qty) VALUES (?, ?, 0)",
                (part_no, wh),
            )
        conn.commit()
        c.close()

def delete_fixture(part_no: str):
    with get_conn() as conn:
        c = conn.cursor()
        c.execute("DELETE FROM fixtures WHERE part_no=?", (part_no,))
        c.execute("DELETE FROM warehouse_stock WHERE part_no=?", (part_no,))
        c.execute("DELETE FROM fixture_boms WHERE parent_part_no=? OR child_part_no=?", (part_no, part_no))
        conn.commit()
        c.close()

def add_stock(part_no: str, qty: int, warehouse: str):
    if qty <= 0:
        raise ValueError("入庫數量必須為正整數")
    with get_conn() as conn:
        c = conn.cursor()
        if not fixture_exists(part_no):
            c.close()
            raise ValueError(f"治具料號 {part_no} 不存在，請先新增治具")

        c.execute(
            "UPDATE warehouse_stock SET qty = qty + ? WHERE part_no = ? AND warehouse = ?",
            (qty, part_no, warehouse),
        )
        if c.rowcount == 0:
            c.execute(
                "INSERT INTO warehouse_stock (part_no, warehouse, qty) VALUES (?, ?, ?)",
                (part_no, warehouse, qty),
            )

        conn.commit()
        c.close()

def transfer_stock(part_no: str, qty: int, from_wh: str, to_wh: str):
    if qty <= 0:
        raise ValueError("調撥數量必須為正整數")
    if from_wh == to_wh:
        raise ValueError("來源與目標倉別相同")

    with get_conn() as conn:
        c = conn.cursor()
        if not fixture_exists(part_no):
            c.close()
            raise ValueError(f"治具料號 {part_no} 不存在，請先新增治具")

        c.execute("SELECT qty FROM warehouse_stock WHERE part_no=? AND warehouse=?", (part_no, from_wh))
        row = c.fetchone()
        if not row or row[0] < qty:
            c.close()
            raise ValueError("來源倉庫數量不足")

        c.execute(
            "INSERT OR IGNORE INTO warehouse_stock (part_no, warehouse, qty) VALUES (?, ?, 0)",
            (part_no, to_wh),
        )

        c.execute("UPDATE warehouse_stock SET qty = qty - ? WHERE part_no = ? AND warehouse = ?", (qty, part_no, from_wh))
        c.execute("UPDATE warehouse_stock SET qty = qty + ? WHERE part_no = ? AND warehouse = ?", (qty, part_no, to_wh))

        c.execute(
            "INSERT INTO transfer_logs (part_no, from_warehouse, to_warehouse, qty) VALUES (?, ?, ?, ?)",
            (part_no, from_wh, to_wh, qty),
        )

        conn.commit()
        c.close()

def update_safety_stock(part_no: str, safety_stock: int):
    with get_conn() as conn:
        c = conn.cursor()
        c.execute("UPDATE fixtures SET safety_stock=? WHERE part_no=?", (safety_stock, part_no))
        if c.rowcount == 0:
            raise ValueError(f"治具料號 {part_no} 不存在")
        conn.commit()
        c.close()


def get_overview_by_warehouse(warehouse: str):
    with get_conn() as conn:
        c = conn.cursor()
        c.execute(
            """
            SELECT f.part_no, f.description, IFNULL(f.spec,''), IFNULL(f.category,''),
                   IFNULL(f.unit_price,0), f.safety_stock, IFNULL(w.qty, 0)
            FROM fixtures f
            LEFT JOIN warehouse_stock w
            ON f.part_no = w.part_no AND w.warehouse = ?
            """,
            (warehouse,),
        )
        rows = c.fetchall()
        c.close()
        return rows

def get_bom_by_part(part_no: str):
    with get_conn() as conn:
        c = conn.cursor()
        c.execute("SELECT id, child_part_no, qty, created_at FROM fixture_boms WHERE parent_part_no=?", (part_no,))
        rows = c.fetchall()
        c.close()
        return rows

def add_bom_item(parent_part_no: str, child_part_no: str, qty: int = 1):
    with get_conn() as conn:
        c = conn.cursor()
        c.execute("INSERT INTO fixture_boms (parent_part_no, child_part_no, qty) VALUES (?, ?, ?)",
                  (parent_part_no, child_part_no, qty))
        conn.commit()
        c.close()

def delete_bom_item(bom_id: int):
    with get_conn() as conn:
        c = conn.cursor()
        c.execute("DELETE FROM fixture_boms WHERE id=?", (bom_id,))
        conn.commit()
        c.close()

def ensure_stock_consistency():
    """確保每個料號在所有倉別都有一筆庫存紀錄 (qty 預設 0)。"""
    with get_conn() as conn:
        c = conn.cursor()
        for wh in CORE_WAREHOUSES:
            c.execute("""
                INSERT OR IGNORE INTO warehouse_stock (part_no, warehouse, qty)
                SELECT part_no, ?, 0 FROM fixtures
            """, (wh,))
        conn.commit()
        c.close()