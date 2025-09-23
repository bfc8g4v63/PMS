#$ force_wal_merge.py
#% 強制將 SQLite 資料庫的 WAL 檔案合併回主資料庫，解除 lock 狀態

import sqlite3

DB_PATH = r"\\192.120.100.177\工程部\生產管理\生產資訊平台\PMS.db"

try:
    conn = sqlite3.connect(DB_PATH)
    conn.execute("PRAGMA wal_checkpoint(FULL);")
    conn.close()
    print("成功強制 WAL 合併至 DB，檔案已解除 lock 狀態。")
except Exception as e:
    print(f"發生錯誤：{e}")