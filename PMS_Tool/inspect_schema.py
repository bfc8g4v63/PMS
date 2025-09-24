#$ inspect_schema.py
#% 列出所有資料表與其欄位結構

import os
import sys

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from db_helper import get_conn, DB_PATH

def inspect_schema(db_path=DB_PATH):
    if not os.path.exists(db_path):
        print(f"[X] 找不到資料庫: {db_path}")
        return

    with get_conn() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")
        tables = [r[0] for r in cursor.fetchall()]

        for t in tables:
            print(f"\n=== 資料表: {t} ===")
            cursor.execute(f"PRAGMA table_info({t})")
            cols = cursor.fetchall()
            for col in cols:
                print(f"- {col[1]} ({col[2]})")

if __name__ == "__main__":
    inspect_schema()