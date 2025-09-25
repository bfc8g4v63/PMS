#$ wal_checkpoint.py
#% 強制將 SQLite 資料庫的 WAL 檔案合併回主資料庫，解除 lock 狀態

import sys
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parents[1]
sys.path.append(str(ROOT_DIR))

from config import apply_db_path
from db_helper import get_conn

apply_db_path()

with get_conn() as conn:
    conn.execute("PRAGMA wal_checkpoint(TRUNCATE);")