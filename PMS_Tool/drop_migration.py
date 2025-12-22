#$ drop_migration.py
#% drop local test

import sys
import shutil
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT_DIR.parent))
sys.path.insert(0, str(ROOT_DIR))

from config import apply_db_path, DB_NAME
from db_helper import get_conn

DROP_TABLES = [
    "fixture_log_audits",
    "fixture_adjustment_logs",
    "dev_logs",
    "consumption_logs",
    "activity_logs_old",
]

def backup_db():
    backup_path = Path(DB_NAME).with_name("PMS_local_backup_before_drop.db")
    shutil.copy(DB_NAME, backup_path)
    print(f"[OK] 已建立備份：{backup_path}")

def table_exists(cur, name):
    cur.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name=?",
        (name,)
    )
    return cur.fetchone() is not None

def migrate():
    with get_conn() as conn:
        cur = conn.cursor()

        for table in DROP_TABLES:
            if table_exists(cur, table):
                cur.execute(f"DROP TABLE {table}")
                print(f"[OK] 已刪除資料表：{table}")
            else:
                print(f"[SKIP] 資料表不存在：{table}")

        conn.commit()

if __name__ == "__main__":
    apply_db_path()
    backup_db()
    migrate()
    print("=== 本地 DB DROP migration 完成 ===")