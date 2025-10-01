import sys
import shutil
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT_DIR.parent))
sys.path.insert(0, str(ROOT_DIR))

from config import apply_db_path, DB_NAME
from db_helper import get_conn

def backup_db():
    backup_path = Path(DB_NAME).with_name("PMS_local_backup.db")
    shutil.copy(DB_NAME, backup_path)
    print(f"[OK] 已建立備份：{backup_path}")

def table_exists(cur, name):
    cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name=?", (name,))
    return cur.fetchone() is not None

def migrate():
    with get_conn() as conn:
        cur = conn.cursor()

        if table_exists(cur, "issues") and not table_exists(cur, "sop_information"):
            cur.execute("ALTER TABLE issues RENAME TO sop_information")
            print("[OK] 已將 issues → sop_information")
        elif table_exists(cur, "SOP") and not table_exists(cur, "sop_information"):
            cur.execute("ALTER TABLE SOP RENAME TO sop_information")
            print("[OK] 已將 SOP → sop_information")

        if table_exists(cur, "activity_logs"):
            cur.execute("PRAGMA table_info(activity_logs)")
            cols = [c[1] for c in cur.fetchall()]
            rename_map = {
                "id": "activity_log_id",
                "username": "activity_log_username",
                "action": "activity_log_action",
                "filename": "activity_log_filename",
                "timestamp": "activity_log_timestamp",
                "module": "activity_log_module"
            }
            needs_migrate = any(col in rename_map for col in cols)

            if needs_migrate:
                cur.execute("""
                    CREATE TABLE IF NOT EXISTS activity_logs_new (
                        activity_log_id INTEGER,
                        activity_log_username TEXT,
                        activity_log_action TEXT,
                        activity_log_filename TEXT,
                        activity_log_timestamp TEXT,
                        activity_log_module TEXT
                    )
                """)
                cur.execute("""
                    INSERT INTO activity_logs_new
                    SELECT id, username, action, filename, timestamp, module
                    FROM activity_logs
                """)
                cur.execute("DROP TABLE activity_logs")
                cur.execute("ALTER TABLE activity_logs_new RENAME TO activity_logs")
                print("[OK] 已重構 activity_logs 欄位名稱")

        conn.commit()

if __name__ == "__main__":
    apply_db_path()
    backup_db()
    migrate()
    print("=== 本地 DB 遷移完成 ===")