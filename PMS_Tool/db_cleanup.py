#$ db_cleanup.py
#% 清理多餘欄位，固定路徑 .177 PMS.db

import sqlite3
import os
import shutil
from datetime import datetime

DB_PATH = r"C:\Nelson\Dev\GitHub\PMS_Document\PMS_Test_db\PMS_2509191428_2.db"
BACKUP_DIR = r"C:\Nelson\Dev\GitHub\PMS_Document\PMS_Test_db"

def backup_db():
    if not os.path.exists(BACKUP_DIR):
        os.makedirs(BACKUP_DIR)
    ts = datetime.now().strftime("%y%m%d%H%M%S")
    backup_file = os.path.join(BACKUP_DIR, f"PMS_{ts}.db.bak")
    if not os.path.exists(DB_PATH):
        raise FileNotFoundError(f"[X] 找不到來源 DB：{DB_PATH}")
    shutil.copy(DB_PATH, backup_file)
    print(f"[✔] 已建立備份：{backup_file}")

def recreate_table(conn, table_name, correct_sql, cols_to_copy):
    cur = conn.cursor()
    tmp_table = f"{table_name}_new"

    cur.execute(f"DROP TABLE IF EXISTS {tmp_table}")

    cur.execute(correct_sql.replace(table_name, tmp_table))

    cur.execute(f"INSERT INTO {tmp_table} SELECT {','.join(cols_to_copy)} FROM {table_name}")

    cur.execute(f"DROP TABLE {table_name}")

    cur.execute(f"ALTER TABLE {tmp_table} RENAME TO {table_name}")

    conn.commit()
    print(f"[✔] 已清理 {table_name}")

def main():
    backup_db()
    conn = sqlite3.connect(DB_PATH)

    recreate_table(conn, "activity_logs", """
        CREATE TABLE activity_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT,
            action TEXT,
            filename TEXT,
            timestamp TEXT,
            module TEXT
        )
    """, ["id", "username", "action", "filename", "timestamp", "module"])

    recreate_table(conn, "fixture_boms", """
        CREATE TABLE fixture_boms (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            parent_part_no TEXT,
            child_part_no TEXT,
            bom_qty INTEGER,
            remark TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """, ["bom_id AS id", "parent_part_no", "child_part_no", "qty AS bom_qty", "remark", "created_at"])

    recreate_table(conn, "fixtures", """
        CREATE TABLE fixtures (
            part_no TEXT PRIMARY KEY,
            part_name TEXT,
            part_spec TEXT,
            part_group TEXT,
            unit_price_ntd REAL,
            unit_price_usd REAL,
            safety_stock INTEGER,
            storage_location TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """, ["part_no", "part_name", "part_spec", "part_group",
          "unit_price_ntd", "unit_price_usd", "safety_stock", "storage_location", "created_at"])

    recreate_table(conn, "transfer_logs", """
        CREATE TABLE transfer_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            part_no TEXT,
            from_wh TEXT,
            to_wh TEXT,
            transfer_qty INTEGER,
            user TEXT,
            remark TEXT,
            created_at TEXT
        )
    """, ["id", "part_no", "from_wh", "to_wh", "transfer_qty", "user", "remark", "created_at"])

    recreate_table(conn, "warehouse_stock", """
        CREATE TABLE warehouse_stock (
            part_no TEXT,
            warehouse TEXT,
            usable_qty INTEGER DEFAULT 0,
            safety_stock INTEGER DEFAULT 0,
            PRIMARY KEY (part_no, warehouse)
        )
    """, ["part_no", "warehouse", "usable_qty", "safety_stock"])

    conn.close()

if __name__ == "__main__":
    main()