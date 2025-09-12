#$ fixture_helper.py
#% 治具模組資料庫輔助
from datetime import datetime
import os
from db_helper import get_conn

CORE_WAREHOUSES = [
    "虹堡",
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
        c.execute("PRAGMA foreign_keys = ON;")

        c.execute("""
        CREATE TABLE IF NOT EXISTS fixtures (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            part_no TEXT UNIQUE,
            description TEXT,
            spec TEXT,
            category TEXT,
            unit_price REAL DEFAULT 0,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
        """)
        add_col_if_missing(conn, "fixtures", "description", "TEXT")
        add_col_if_missing(conn, "fixtures", "spec", "TEXT")
        add_col_if_missing(conn, "fixtures", "category", "TEXT")
        add_col_if_missing(conn, "fixtures", "unit_price", "REAL DEFAULT 0")
        add_col_if_missing(conn, "fixtures", "created_at", "TIMESTAMP DEFAULT CURRENT_TIMESTAMP")

        c.execute("""
        CREATE TABLE IF NOT EXISTS warehouse_stock (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            part_no TEXT,
            warehouse TEXT,
            qty INTEGER DEFAULT 0,
            safety_stock INTEGER DEFAULT 0,
            location TEXT DEFAULT '',
            UNIQUE(part_no, warehouse)
        )
        """)
        add_col_if_missing(conn, "warehouse_stock", "part_no", "TEXT")
        add_col_if_missing(conn, "warehouse_stock", "warehouse", "TEXT")
        add_col_if_missing(conn, "warehouse_stock", "qty", "INTEGER DEFAULT 0")
        add_col_if_missing(conn, "warehouse_stock", "safety_stock", "INTEGER DEFAULT 0")
        add_col_if_missing(conn, "warehouse_stock", "location", "TEXT DEFAULT ''")

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

def ensure_inventory_schema():
    ddl = """
    PRAGMA foreign_keys = ON;

    CREATE TABLE IF NOT EXISTS warehouses (
        warehouse_id   INTEGER PRIMARY KEY AUTOINCREMENT,
        code           TEXT UNIQUE NOT NULL,
        name           TEXT NOT NULL,
        is_active      INTEGER NOT NULL DEFAULT 1,
        created_at     TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%S','now','localtime'))
    );

    CREATE VIEW IF NOT EXISTS v_item_stock_summary AS
    SELECT
      f.part_no,
      f.description AS item_name,
      IFNULL(f.spec,'') AS spec,
      (
        SELECT COALESCE(ws.safety_stock, 0)
        FROM warehouse_stock ws
        WHERE ws.part_no = f.part_no AND ws.warehouse = '虹堡'
      ) AS safety_stock,
      w.name AS warehouse_name,
      COALESCE(ws.qty, 0) AS qty,
      CASE
        WHEN (
          SELECT COALESCE(SUM(ws2.qty), 0)
          FROM warehouse_stock ws2
          WHERE ws2.part_no = f.part_no
        ) <
        (
          SELECT COALESCE(s.safety_stock, 0)
          FROM warehouse_stock s
          WHERE s.part_no = f.part_no AND s.warehouse = '虹堡'
        )
        THEN '低於安庫' ELSE 'OK'
      END AS safety_check
    FROM fixtures f
    JOIN warehouses w ON w.is_active = 1
    LEFT JOIN warehouse_stock ws
      ON ws.part_no = f.part_no AND ws.warehouse = w.name;

    CREATE VIEW IF NOT EXISTS v_safety_stock_alerts AS
    SELECT
      f.part_no,
      f.description AS item_name,
      IFNULL(f.spec,'') AS spec,
      (
        SELECT COALESCE(s.safety_stock, 0)
        FROM warehouse_stock s
        WHERE s.part_no = f.part_no AND s.warehouse = '虹堡'
      ) AS safety_stock,
      COALESCE((
        SELECT SUM(ws.qty)
        FROM warehouse_stock ws
        WHERE ws.part_no = f.part_no
      ), 0) AS total_qty,
      (
        (
          SELECT COALESCE(s.safety_stock, 0)
          FROM warehouse_stock s
          WHERE s.part_no = f.part_no AND s.warehouse = '虹堡'
        )
        -
        COALESCE((
          SELECT SUM(ws.qty)
          FROM warehouse_stock ws
          WHERE ws.part_no = f.part_no
        ), 0)
      ) AS shortage
    FROM fixtures f
    WHERE COALESCE((
        SELECT SUM(ws.qty)
        FROM warehouse_stock ws
        WHERE ws.part_no = f.part_no
      ), 0)
      <
      (
        SELECT COALESCE(s.safety_stock, 0)
        FROM warehouse_stock s
        WHERE s.part_no = f.part_no AND s.warehouse = '虹堡'
      );

    INSERT OR IGNORE INTO warehouses(code, name) VALUES
      ('HB', '虹堡'),
      ('SJ', '上齊'),
      ('RJ', '睿均'),
      ('JH', '捷暉'),
      ('LR', '立榮'),
      ('HQ', '華勤'),
      ('SM', '上貿'),
      ('MB', '麥博'),
      ('XL', '信利'),
      ('GC', 'GC'),
      ('EN', '工程'),
      ('NG', '不良品');
    """
    with get_conn() as conn:
        conn.executescript(ddl)
        conn.commit()

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
                   unit_price: float = 0.0, safety_stock: int = 0, location: str = ""):
    with get_conn() as conn:
        c = conn.cursor()
        c.execute("SELECT 1 FROM fixtures WHERE part_no=?", (part_no,))
        if c.fetchone():
            raise ValueError(f"治具料號 {part_no} 已存在，不可重複建立")

        c.execute(
            "INSERT INTO fixtures (part_no, description, spec, category, unit_price) VALUES (?, ?, ?, ?, ?)",
            (part_no, name, spec, category, unit_price),
        )

        for wh in CORE_WAREHOUSES:
            if wh == "虹堡":
                c.execute(
                    "INSERT OR IGNORE INTO warehouse_stock (part_no, warehouse, qty, safety_stock, location) VALUES (?, ?, 0, ?, ?)",
                    (part_no, wh, safety_stock, location),
                )
            else:
                c.execute(
                    "INSERT OR IGNORE INTO warehouse_stock (part_no, warehouse, qty, safety_stock, location) VALUES (?, ?, 0, 0, '')",
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
                "INSERT INTO warehouse_stock (part_no, warehouse, qty, safety_stock, location) VALUES (?, ?, ?, 0, '')",
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
            "INSERT OR IGNORE INTO warehouse_stock (part_no, warehouse, qty, safety_stock, location) VALUES (?, ?, 0, 0, '')",
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

def update_safety_stock(part_no: str, safety_stock: int, location: str = ""):
    with get_conn() as conn:
        c = conn.cursor()
        c.execute(
            "UPDATE warehouse_stock SET safety_stock=?, location=? WHERE part_no=? AND warehouse='虹堡'",
            (safety_stock, location, part_no),
        )
        if c.rowcount == 0:
            raise ValueError(f"治具料號 {part_no} 不存在或不在虹堡")
        conn.commit()
        c.close()

def get_overview_by_warehouse(warehouse: str):
    with get_conn() as conn:
        c = conn.cursor()
        c.execute(
            """
            SELECT f.part_no,
                   f.description,
                   IFNULL(f.spec,''),
                   IFNULL(f.category,''),
                   IFNULL(f.unit_price,0),
                   IFNULL(w.qty, 0),
                   CASE WHEN w.warehouse='虹堡' THEN IFNULL(w.location,'') ELSE '' END AS location_display,
                   CASE WHEN w.warehouse='虹堡' THEN IFNULL(w.safety_stock,0) ELSE 0 END AS safety_stock_display
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
    """確保每個料號在所有倉別都有一筆庫存紀錄 (qty 預設 0)，虹堡允許自訂安庫/儲位。"""
    with get_conn() as conn:
        c = conn.cursor()
        for wh in CORE_WAREHOUSES:
            if wh == "虹堡":
                c.execute("""
                    INSERT OR IGNORE INTO warehouse_stock (part_no, warehouse, qty, safety_stock, location)
                    SELECT part_no, '虹堡', 0, 0, '' FROM fixtures
                """)
            else:
                c.execute("""
                    INSERT OR IGNORE INTO warehouse_stock (part_no, warehouse, qty, safety_stock, location)
                    SELECT part_no, ?, 0, 0, '' FROM fixtures
                """, (wh,))
        conn.commit()
        c.close()