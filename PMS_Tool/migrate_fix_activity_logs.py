#$ migrate_fix_activity_logs.py
#% 修復 activity_logs 表格，確保 activity_log_id 為 AUTOINCREMENT

import sqlite3
from pathlib import Path

DB_PATH = Path("PMS.db")

def migrate_activity_logs(db_path=DB_PATH):
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()

    cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='activity_logs'")
    if not cur.fetchone():
        print("[ERROR] activity_logs 不存在，無需修復")
        conn.close()
        return

    cur.execute("""
        SELECT 
            activity_log_username, 
            activity_log_action, 
            activity_log_filename, 
            activity_log_timestamp, 
            activity_log_module
        FROM activity_logs
    """)
    old_data = cur.fetchall()
    print(f"[INFO] 讀出 {len(old_data)} 筆舊資料")

    cur.execute("ALTER TABLE activity_logs RENAME TO activity_logs_old")

    cur.execute("""
        CREATE TABLE activity_logs (
            activity_log_id INTEGER PRIMARY KEY AUTOINCREMENT,
            activity_log_username TEXT,
            activity_log_action TEXT,
            activity_log_filename TEXT,
            activity_log_timestamp TEXT,
            activity_log_module TEXT
        )
    """)

    cur.executemany("""
        INSERT INTO activity_logs (
            activity_log_username,
            activity_log_action,
            activity_log_filename,
            activity_log_timestamp,
            activity_log_module
        ) VALUES (?, ?, ?, ?, ?)
    """, old_data)

    conn.commit()
    conn.close()
    print("[OK] activity_logs 修復完成，所有舊資料已搬移")

if __name__ == "__main__":
    migrate_activity_logs()
