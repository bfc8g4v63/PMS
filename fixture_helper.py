import sqlite3
from datetime import datetime

def insert_fixture(db_name, part_no, name, spec, safety_stock, user_name):
    """
    新增治具基本資料（若重複則略過）並初始化 item_balances 存量。
    """
    try:
        with sqlite3.connect(db_name) as conn:
            conn.execute("PRAGMA foreign_keys = ON;")
            cursor = conn.cursor()

            cursor.execute("""
                INSERT OR IGNORE INTO fixtures (part_no, name, spec, safety_stock, created_at)
                VALUES (?, ?, ?, ?, ?)
            """, (part_no, name, spec, safety_stock, datetime.now().strftime('%Y-%m-%dT%H:%M:%S')))

            cursor.execute("SELECT fixture_id FROM fixtures WHERE part_no = ?", (part_no,))
            fixture_id = cursor.fetchone()
            if not fixture_id:
                raise Exception(f"無法取得 fixture_id，料號可能插入失敗：{part_no}")
            fixture_id = fixture_id[0]

            cursor.execute("SELECT warehouse_id FROM warehouses WHERE is_active = 1")
            warehouse_ids = [row[0] for row in cursor.fetchall()]

            for wid in warehouse_ids:
                cursor.execute("""
                    INSERT OR IGNORE INTO item_balances (fixture_id, warehouse_id, qty)
                    VALUES (?, ?, 0)
                """, (fixture_id, wid))

            conn.commit()
            return True, "治具新增成功"
    except Exception as e:
        return False, f"錯誤：{str(e)}"