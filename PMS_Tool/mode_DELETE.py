import sqlite3

conn = sqlite3.connect("C:/Temp/PMS_local.db")
cursor = conn.cursor()
cursor.execute("PRAGMA journal_mode=DELETE;")
conn.commit()
conn.close()