#$ cleanup_schema.py
#% 刪除舊表格與不用的表格，保留必要表格

import sqlite3

DB_PATH = r"C:\Nelson\Dev\GitHub\PMS\PMS.db"

def cleanup():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    drop_tables = ["fixtures_old", "warehouse_stock_old", "issues_old", "dev_logs"]

    for tbl in drop_tables:
        cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name=?", (tbl,))
        if cur.fetchone():
            cur.execute(f"DROP TABLE {tbl}")
            print(f"[✓] 已刪除 {tbl}")
        else:
            print(f"[ ] 無此表格 {tbl}")

    conn.commit()
    conn.close()
    print("[✓] cleanup 完成")

if __name__ == "__main__":
    cleanup()