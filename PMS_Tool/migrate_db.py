### migrate_db.py
## 資料庫結構修正腳本
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

def recreate_table(cur, table, schema_sql, insert_sql=None, select_sql=None):
    cur.execute(f"ALTER TABLE {table} RENAME TO {table}_old")
    cur.execute(schema_sql)
    if insert_sql and select_sql:
        cur.execute(insert_sql + " " + select_sql)
    cur.execute(f"DROP TABLE {table}_old")
    print(f"[✓] {table} 表格已修正")

def migrate():
    backup_db()
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    # fixtures
    cur.execute("PRAGMA table_info(fixtures)")
    cols = [r[1] for r in cur.fetchall()]
    expected_cols = {"part_no","part_name","part_spec","part_group",
                     "unit_price_ntd","unit_price_usd","safety_stock","storage_location"}
    if set(cols) != expected_cols:
        recreate_table(
            cur,
            "fixtures",
            """
            CREATE TABLE fixtures (
                part_no TEXT PRIMARY KEY,
                part_name TEXT,
                part_spec TEXT,
                part_group TEXT,
                unit_price_ntd REAL,
                unit_price_usd REAL,
                safety_stock INTEGER,
                storage_location TEXT
            )
            """,
            insert_sql="""
            INSERT INTO fixtures (part_no, part_name, part_spec, part_group,
                                  unit_price_ntd, unit_price_usd,
                                  safety_stock, storage_location)
            """,
            select_sql="""
            SELECT part_no, part_name, part_spec, part_group,
                   unit_price_ntd, unit_price_usd,
                   safety_stock, storage_location
            FROM fixtures_old
            """
        )

    # warehouse_stock
    cur.execute("PRAGMA table_info(warehouse_stock)")
    cols = [r[1] for r in cur.fetchall()]
    expected_cols = {"part_no","warehouse","usable_qty","safety_stock"}
    if set(cols) != expected_cols:
        recreate_table(
            cur,
            "warehouse_stock",
            """
            CREATE TABLE warehouse_stock (
                part_no TEXT,
                warehouse TEXT,
                usable_qty INTEGER DEFAULT 0,
                safety_stock INTEGER DEFAULT 0,
                PRIMARY KEY (part_no, warehouse),
                FOREIGN KEY (part_no) REFERENCES fixtures(part_no)
            )
            """,
            insert_sql="""
            INSERT INTO warehouse_stock (part_no, warehouse, usable_qty, safety_stock)
            """,
            select_sql="""
            SELECT part_no, warehouse, qty as usable_qty, safety_stock
            FROM warehouse_stock_old
            """
        )

    # issues
    cur.execute("PRAGMA table_info(issues)")
    cols = [r[1] for r in cur.fetchall()]
    expected_cols = {"product_code","product_name","dip_sop","assembly_sop","test_sop","packaging_sop","oqc_checklist"}
    if set(cols) != expected_cols:
        recreate_table(
            cur,
            "issues",
            """
            CREATE TABLE issues (
                product_code TEXT PRIMARY KEY,
                product_name TEXT,
                dip_sop TEXT,
                assembly_sop TEXT,
                test_sop TEXT,
                packaging_sop TEXT,
                oqc_checklist TEXT
            )
            """,
            insert_sql="""
            INSERT INTO issues (product_code, product_name, dip_sop, assembly_sop, test_sop, packaging_sop, oqc_checklist)
            """,
            select_sql="""
            SELECT product_code, product_name, dip_sop, assembly_sop, test_sop, packaging_sop, oqc_checklist
            FROM issues_old
            """
        )

    # users
    cur.execute("PRAGMA table_info(users)")
    cols = [r[1] for r in cur.fetchall()]
    expected_cols = {"username","password","role","specialty",
                     "can_view_logs","can_delete_logs","can_upload_sop","can_view_issues","can_manage_users"}
    if set(cols) != expected_cols:
        recreate_table(
            cur,
            "users",
            """
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
            """,
            insert_sql="""
            INSERT INTO users (username,password,role,specialty,can_view_logs,can_delete_logs,can_upload_sop,can_view_issues,can_manage_users)
            """,
            select_sql="""
            SELECT username,password,role,specialty,can_view_logs,can_delete_logs,can_upload_sop,can_view_issues,can_manage_users
            FROM users_old
            """
        )

    # activity_logs
    cur.execute("PRAGMA table_info(activity_logs)")
    cols = [r[1] for r in cur.fetchall()]
    expected_cols = {"id","username","action","filename","timestamp","module"}
    if set(cols) != expected_cols:
        recreate_table(
            cur,
            "activity_logs",
            """
            CREATE TABLE activity_logs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT,
                action TEXT,
                filename TEXT,
                timestamp TEXT,
                module TEXT
            )
            """,
            insert_sql="""
            INSERT INTO activity_logs (id, username, action, filename, timestamp, module)
            """,
            select_sql="""
            SELECT id, username, action, filename, timestamp, module
            FROM activity_logs_old
            """
        )

    # transfer_logs
    cur.execute("PRAGMA table_info(transfer_logs)")
    cols = [r[1] for r in cur.fetchall()]
    expected_cols = {"id","part_no","from_wh","to_wh","transfer_qty","user","created_at","remark"}
    if set(cols) != expected_cols:
        recreate_table(
            cur,
            "transfer_logs",
            """
            CREATE TABLE transfer_logs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                part_no TEXT,
                from_wh TEXT,
                to_wh TEXT,
                transfer_qty INTEGER,
                user TEXT,
                remark TEXT,
                created_at TEXT
            )
            """,
            insert_sql="""
            INSERT INTO transfer_logs (id, part_no, from_wh, to_wh, transfer_qty, user, remark, created_at)
            """,
            select_sql="""
            SELECT id, part_no, from_wh, to_wh, qty as transfer_qty, user, remark, timestamp as created_at
            FROM transfer_logs_old
            """
        )

    # consumption_logs
    cur.execute("PRAGMA table_info(consumption_logs)")
    cols = [r[1] for r in cur.fetchall()]
    expected_cols = {"id","part_no","warehouse","consume_qty","user","created_at","remark"}
    if set(cols) != expected_cols:
        recreate_table(
            cur,
            "consumption_logs",
            """
            CREATE TABLE consumption_logs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                part_no TEXT,
                warehouse TEXT,
                consume_qty INTEGER,
                user TEXT,
                remark TEXT,
                created_at TEXT
            )
            """,
            insert_sql="""
            INSERT INTO consumption_logs (id, part_no, warehouse, consume_qty, user, remark, created_at)
            """,
            select_sql="""
            SELECT id, part_no, warehouse, qty as consume_qty, user, remark, timestamp as created_at
            FROM consumption_logs_old
            """
        )

    conn.commit()
    conn.close()
    print("[✓] 資料庫欄位全部檢查並修正完成")

if __name__ == "__main__":
    migrate()