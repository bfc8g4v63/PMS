#$ PRAGMA_journal_mode_DELETE.py
#% 將 SQLite 資料庫的 journal mode 設定為 DELETE

import sys
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parents[1]
sys.path.append(str(ROOT_DIR))

from config import apply_db_path
from db_helper import get_conn

apply_db_path()

with get_conn() as conn:
    conn.execute("PRAGMA journal_mode=DELETE;")