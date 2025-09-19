# drop_dev_logs.py
# 移除多餘的 dev_logs 資料表

import sqlite3

DB_PATH = r"C:\Nelson\Dev\GitHub\PMS_Document\PMS_Test_db\PMS_2509191428_2.db"

def drop_dev_logs(db_path=DB_PATH):
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.execute("DROP TABLE IF EXISTS dev_logs")
    conn.commit()
    conn.close()
    print("[✔] 已刪除 dev_logs 資料表")

if __name__ == "__main__":
    drop_dev_logs()