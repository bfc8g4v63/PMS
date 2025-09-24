#$ drop_dev_logs.py
#% 移除多餘的 dev_logs 資料表

import os
import sys

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from db_helper import get_conn, DB_PATH

def drop_dev_logs(db_path=DB_PATH):
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("DROP TABLE IF EXISTS dev_logs")
        conn.commit()
    print("[✔] 已刪除 dev_logs 資料表")

if __name__ == "__main__":
    drop_dev_logs()