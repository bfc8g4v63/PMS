#$ cleanup_schema.py
#%
import sqlite3
import os
import shutil

DB_PATH = r"C:\Nelson\Dev\GitHub\PMS\PMS.db"
BACKUP_PATH = DB_PATH + ".bak"

def backup_db():
    if not os.path.exists(BACKUP_PATH):
        shutil.copy2(DB_PATH, BACKUP_PATH)
        print(f"[✓] 已建立備份 {BACKUP_PATH}")
    else:
        print(f"[!] 備份已存在：{BACKUP_PATH}")

def recreate_users(cur):
    cur.execute("ALTER TABLE users RENAME TO users_old")
    cur.execute("""
    CREATE TABLE users (
        username TEXT PRIMARY KEY,
        password TEXT,
        role TEXT,
        specialty TEXT,
        can_view_logs INTEGER DEFAULT 0,
        can_delete_logs INTEGER DEFAULT 0,
        can_upload_sop INTEGER DEFAULT 0,
        can_view_issues INTEGER DEFAULT 0,
        can_manage_users INTEGER DEFAULT 0
    )
    """)
    cur.execute("""
    INSERT INTO users (username,password,role,specialty,
                       can_view_logs,can_delete_logs,
                       can_upload_sop,can_view_issues,can_manage_users)
    SELECT username,password,role,specialty,
           can_view_logs,can_delete_logs,
           can_upload_sop,can_view_issues,can_manage_users
    FROM users_old
    """)
    cur.execute("DROP TABLE users_old")
    print("[✓] users 已修正")

def recreate_activity_logs(cur):
    cur.execute("ALTER TABLE activity_logs RENAME TO activity_logs_old")
    cur.execute("""
    CREATE TABLE activity_logs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT,
        action TEXT,
        filename TEXT,
        timestamp TEXT,
        module TEXT
    )
    """)
    cur.execute("""
    INSERT INTO activity_logs (username, action, filename, timestamp, module)
    SELECT username, action, filename, timestamp, module
    FROM activity_logs_old
    """)
    cur.execute("DROP TABLE activity_logs_old")
    print("[✓] activity_logs 已修正")

def recreate_transfer_logs(cur):
    cur.execute("ALTER TABLE transfer_logs RENAME TO transfer_logs_old")
    cur.execute("""
    CREATE TABLE transfer_logs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        part_no TEXT,
        from_wh TEXT,
        to_wh TEXT,
        usable_qty INTEGER,
        user TEXT,
        timestamp TEXT,
        remark TEXT
    )
    """)
    cur.execute("""
    INSERT INTO transfer_logs (part_no, from_wh, to_wh, usable_qty, user, remark)
    SELECT part_no,
           COALESCE(from_wh, from_warehouse),
           COALESCE(to_wh, to_warehouse),
           qty,
           user,
           remark
    FROM transfer_logs_old
    """)
    cur.execute("DROP TABLE transfer_logs_old")
    print("[✓] transfer_logs 已修正")

def recreate_consumption_logs(cur):
    cur.execute("ALTER TABLE consumption_logs RENAME TO consumption_logs_old")
    cur.execute("""
    CREATE TABLE consumption_logs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        part_no TEXT,
        warehouse TEXT,
        usable_qty INTEGER,
        user TEXT,
        timestamp TEXT,
        remark TEXT
    )
    """)
    cur.execute("""
    INSERT INTO consumption_logs (part_no, usable_qty, user)
    SELECT part_no, qty, user
    FROM consumption_logs_old
    """)
    cur.execute("DROP TABLE consumption_logs_old")
    print("[✓] consumption_logs 已修正")

def migrate():
    backup_db()
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    # users
    cur.execute("PRAGMA table_info(users)")
    cols = [r[1] for r in cur.fetchall()]
    if "can_add" in cols or "can_delete" in cols or "active" in cols:
        recreate_users(cur)

    # activity_logs
    cur.execute("PRAGMA table_info(activity_logs)")
    cols = [r[1] for r in cur.fetchall()]
    if "product_code" in cols:
        recreate_activity_logs(cur)

    # transfer_logs
    cur.execute("PRAGMA table_info(transfer_logs)")
    cols = [r[1] for r in cur.fetchall()]
    if "qty" in cols or "log_id" in cols:
        recreate_transfer_logs(cur)

    # consumption_logs
    cur.execute("PRAGMA table_info(consumption_logs)")
    cols = [r[1] for r in cur.fetchall()]
    if "qty" in cols or "consume_id" in cols:
        recreate_consumption_logs(cur)

    conn.commit()
    conn.close()
    print("[✓] Schema 清理完成")

if __name__ == "__main__":
    migrate()