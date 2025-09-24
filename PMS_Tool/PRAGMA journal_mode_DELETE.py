# PRAGMA_journal_mode_DELETE.py
# 將 SQLite 資料庫的 journal mode 設定為 DELETE

import os
import sys

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from db_helper import get_conn

with get_conn() as conn:
    conn.execute("PRAGMA journal_mode=DELETE;")