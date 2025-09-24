#$ force_wal_merge.py
#% 強制將 SQLite 資料庫的 WAL 檔案合併回主資料庫，解除 lock 狀態

import os
import sys

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from db_helper import get_conn

def force_wal_merge():
    try:
        with get_conn() as conn:
            conn.execute("PRAGMA wal_checkpoint(FULL);")
        print("成功強制 WAL 合併至 DB，檔案已解除 lock 狀態。")
    except Exception as e:
        print(f"發生錯誤：{e}")

if __name__ == "__main__":
    force_wal_merge()