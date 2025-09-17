import sqlite3

conn = sqlite3.connect(r"\\192.120.100.177\工程部\生產管理\生產資訊平台\PMS.db")
cursor = conn.cursor()

cursor.execute("PRAGMA journal_mode=DELETE;")
conn.commit()
conn.close()