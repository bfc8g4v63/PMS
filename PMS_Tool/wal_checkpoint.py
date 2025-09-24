#$ wal_checkpoint.py
#% 強制將 SQLite 資料庫的 WAL 檔案合併回主資料庫，解除 lock 狀態

import os
import sys

# 加入父目錄到 sys.path，讓工具能 import db_helper
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from db_helper import get_conn

with get_conn() as conn:
    conn.execute("PRAGMA wal_checkpoint(TRUNCATE);")