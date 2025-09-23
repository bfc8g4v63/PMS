#$ PRAGMA journal_mode_DELETE.py
#% 將 SQLite 資料庫的 journal mode 設定為 DELETE

import sqlite3

conn = sqlite3.connect(r"\\192.120.100.177\工程部\生產管理\生產資訊平台\PMS.db")
cursor = conn.cursor()

cursor.execute("PRAGMA journal_mode=DELETE;")
conn.commit()
conn.close()