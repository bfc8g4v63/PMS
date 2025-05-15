import sqlite3
from datetime import datetime

def log_activity(db_name, user, action, filename, module=None):
    with sqlite3.connect(db_name) as conn:
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO activity_logs (username, action, filename, timestamp, module)
            VALUES (?, ?, ?, ?, ?)
        """, (user, action, filename, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), module))
        conn.commit()
