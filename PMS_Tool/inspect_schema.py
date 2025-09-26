#$ inspect_schema.py
#% 列出所有資料表與其欄位結構

import sys
import os
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parents[1]
sys.path.append(str(ROOT_DIR))

from config import apply_db_path
from db_helper import get_conn, DB_PATH

apply_db_path()

def inspect_schema(db_path=None):
    if db_path is None:
        with get_conn() as conn:
            db_path = conn.execute("PRAGMA database_list").fetchone()[2]

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