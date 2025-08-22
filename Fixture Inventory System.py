import sqlite3
from datetime import datetime
from config import DB_NAME

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

    CREATE TABLE IF NOT EXISTS fixtures (
        fixture_id     INTEGER PRIMARY KEY AUTOINCREMENT,
        part_no        TEXT UNIQUE NOT NULL,
        name           TEXT NOT NULL,
        spec           TEXT,
        safety_stock   INTEGER NOT NULL DEFAULT 0,
        created_at     TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%S','now','localtime')),
        updated_at     TEXT
    );

    CREATE TABLE IF NOT EXISTS stock_moves (
        move_id        INTEGER PRIMARY KEY AUTOINCREMENT,
        move_time      TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%S','now','localtime')),
        move_type      TEXT NOT NULL,
        fixture_id     INTEGER NOT NULL REFERENCES fixtures(fixture_id) ON DELETE RESTRICT,
        from_wh        INTEGER REFERENCES warehouses(warehouse_id) ON DELETE RESTRICT,
        to_wh          INTEGER REFERENCES warehouses(warehouse_id) ON DELETE RESTRICT,
        qty            INTEGER NOT NULL,
        ref_no         TEXT,
        note           TEXT,
        user_name      TEXT
    );

    CREATE TABLE IF NOT EXISTS item_balances (
        fixture_id     INTEGER NOT NULL REFERENCES fixtures(fixture_id) ON DELETE CASCADE,
        warehouse_id   INTEGER NOT NULL REFERENCES warehouses(warehouse_id) ON DELETE CASCADE,
        qty            INTEGER NOT NULL DEFAULT 0,
        PRIMARY KEY (fixture_id, warehouse_id)
    );

    CREATE VIEW IF NOT EXISTS v_item_stock_summary AS
    WITH wh AS (
        SELECT warehouse_id, name FROM warehouses WHERE is_active = 1
    )
    SELECT
      f.part_no,
      f.name AS item_name,
      f.spec,
      f.safety_stock,
      w.name AS warehouse_name,
      COALESCE(b.qty, 0) AS qty,
      CASE
        WHEN SUM(COALESCE(b2.qty,0)) OVER (PARTITION BY f.fixture_id) < f.safety_stock THEN '低於安庫'
        ELSE 'OK'
      END AS safety_check
    FROM fixtures f
    CROSS JOIN wh w
    LEFT JOIN item_balances b ON b.fixture_id = f.fixture_id AND b.warehouse_id = w.warehouse_id
    LEFT JOIN item_balances b2 ON b2.fixture_id = f.fixture_id;

    CREATE VIEW IF NOT EXISTS v_safety_stock_alerts AS
    SELECT
      f.part_no, f.name AS item_name, f.spec,
      f.safety_stock,
      COALESCE(SUM(b.qty),0) AS total_qty,
      (f.safety_stock - COALESCE(SUM(b.qty),0)) AS shortage
    FROM fixtures f
    LEFT JOIN item_balances b ON b.fixture_id = f.fixture_id
    GROUP BY f.fixture_id
    HAVING COALESCE(SUM(b.qty),0) < f.safety_stock;

    INSERT OR IGNORE INTO warehouses(code, name) VALUES
      ('ZJ', '治具室'),
      ('SJ', '上齊'),
      ('RJ', '睿均'),
      ('NG', '不良品');
    """

    with sqlite3.connect(DB_NAME, timeout=10) as conn:
        conn.executescript(ddl)
        conn.commit()

if __name__ == "__main__":
    ensure_inventory_schema()
    print("治具進存銷 schema 初始化完成")