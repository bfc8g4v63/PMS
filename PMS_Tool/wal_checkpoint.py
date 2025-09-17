import sqlite3

conn = sqlite3.connect(r"\\192.120.100.177\工程部\生產管理\生產資訊平台\PMS.db", timeout=5)
conn.execute("PRAGMA journal_mode=WAL;")
conn.execute("PRAGMA wal_checkpoint(TRUNCATE);")
conn.close()