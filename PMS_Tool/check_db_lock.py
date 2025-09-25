#$ check_db_lock.py
#% 檢查 SQLite 資料庫是否被鎖定

import os
import sys
import socket
from datetime import datetime
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parents[1]
sys.path.append(str(ROOT_DIR))

from config import apply_db_path
from db_helper import get_conn, DB_PATH

apply_db_path()

def check_lock(db_path=DB_PATH):
    hostname = socket.gethostname()
    user = os.getenv("USERNAME", "unknown")
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    result = ""
    try:
        with get_conn() as conn:
            cur = conn.cursor()
            try:
                cur.execute("BEGIN IMMEDIATE;")
                conn.rollback()
                result = f"[OK] {hostname} ({user}) {timestamp} → 本機沒有鎖定 DB"
            except Exception as e:
                if "locked" in str(e).lower():
                    result = f"[BUSY] {hostname} ({user}) {timestamp} → DB 已被其他電腦鎖定"
                else:
                    result = f"[ERROR] {hostname} ({user}) {timestamp} → {e}"
    except Exception as e:
        if "locked" in str(e).lower():
            result = f"[LOCKED] {hostname} ({user}) {timestamp} → 本機持有 DB 鎖"
        else:
            result = f"[ERROR] {hostname} ({user}) {timestamp} → {e}"

    print(result)
    with open("db_lock_check.log", "a", encoding="utf-8") as f:
        f.write(result + "\n")

if __name__ == "__main__":
    path = DB_PATH
    if len(sys.argv) > 1:
        path = sys.argv[1]
    check_lock(path)
    input("\n檢查完成，請按 Enter 鍵關閉視窗...")