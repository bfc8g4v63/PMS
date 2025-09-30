#$ migrate_db.py
#% SQLite 資料遷移工具

import sys
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parent
sys.path.append(str(ROOT_DIR.parent))

from config import apply_db_path
from db_helper import get_conn

def table_exists(cur, name):
    cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name=?", (name,))
    return cur.fetchone() is not None

def migrate_fixture_logs():
    with get_conn() as conn:
        cur = conn.cursor()
        if not table_exists(cur, "fixture_logs"):
            return
        cur.execute("PRAGMA table_info(fixture_logs)")
        cols = [c[1] for c in cur.fetchall()]
        if "fixture_log_id" in cols:
            print("[SKIP] fixture_logs 已是新版")
            return

        cur.execute("""
            CREATE TABLE fixture_logs_new (
                fixture_log_id TEXT PRIMARY KEY,
                fixture_log_part_no TEXT,
                fixture_log_action TEXT,
                fixture_log_qty INTEGER,
                fixture_log_from_wh TEXT,
                fixture_log_to_wh TEXT,
                fixture_log_user TEXT,
                fixture_log_timestamp TEXT
            )
        """)
        sel_cols = []
        for c in ["part_no","action","change_qty","from_wh","to_wh","user","timestamp"]:
            sel_cols.append(c if c in cols else "NULL")
        cur.execute(f"SELECT {','.join(sel_cols)} FROM fixture_logs")
        rows = cur.fetchall()
        for idx, r in enumerate(rows, 1):
            log_id = f"FLOG-{idx:06d}"
            cur.execute("""
                INSERT INTO fixture_logs_new
                (fixture_log_id, fixture_log_part_no, fixture_log_action, fixture_log_qty,
                 fixture_log_from_wh, fixture_log_to_wh, fixture_log_user, fixture_log_timestamp)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (log_id, r[0], r[1], r[2], r[3], r[4], r[5], r[6]))
        cur.execute("DROP TABLE fixture_logs")
        cur.execute("ALTER TABLE fixture_logs_new RENAME TO fixture_logs")
        conn.commit()
        print("[OK] fixture_logs 已完成遷移")

def migrate_transfer_logs():
    with get_conn() as conn:
        cur = conn.cursor()
        if not table_exists(cur, "transfer_logs"):
            return
        cur.execute("PRAGMA table_info(transfer_logs)")
        cols = [c[1] for c in cur.fetchall()]
        if "transfer_log_id" in cols:
            print("[SKIP] transfer_logs 已是新版")
            return

        cur.execute("""
            CREATE TABLE transfer_logs_new (
                transfer_log_id TEXT PRIMARY KEY,
                transfer_log_part_no TEXT,
                transfer_log_from_wh TEXT,
                transfer_log_to_wh TEXT,
                transfer_log_qty INTEGER,
                transfer_log_user TEXT,
                transfer_log_timestamp TEXT
            )
        """)
        sel_cols = []
        for c in ["part_no","from_wh","to_wh","transfer_qty","user","created_at"]:
            sel_cols.append(c if c in cols else "NULL")
        cur.execute(f"SELECT {','.join(sel_cols)} FROM transfer_logs")
        rows = cur.fetchall()
        for idx, r in enumerate(rows, 1):
            log_id = f"TLOG-{idx:06d}"
            cur.execute("""
                INSERT INTO transfer_logs_new
                (transfer_log_id, transfer_log_part_no, transfer_log_from_wh, transfer_log_to_wh,
                 transfer_log_qty, transfer_log_user, transfer_log_timestamp)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            """, (log_id, r[0], r[1], r[2], r[3], r[4], r[5]))
        cur.execute("DROP TABLE transfer_logs")
        cur.execute("ALTER TABLE transfer_logs_new RENAME TO transfer_logs")
        conn.commit()
        print("[OK] transfer_logs 已完成遷移")

def migrate_consumption_logs():
    with get_conn() as conn:
        cur = conn.cursor()
        if not table_exists(cur, "consumption_logs"):
            return
        cur.execute("PRAGMA table_info(consumption_logs)")
        cols = [c[1] for c in cur.fetchall()]
        if "consumption_log_id" in cols:
            print("[SKIP] consumption_logs 已是新版")
            return

        cur.execute("""
            CREATE TABLE consumption_logs_new (
                consumption_log_id TEXT PRIMARY KEY,
                consumption_log_part_no TEXT,
                consumption_log_warehouse TEXT,
                consumption_log_qty INTEGER,
                consumption_log_user TEXT,
                consumption_log_timestamp TEXT
            )
        """)
        sel_cols = []
        for c in ["part_no","warehouse","consume_qty","user","created_at"]:
            sel_cols.append(c if c in cols else "NULL")
        cur.execute(f"SELECT {','.join(sel_cols)} FROM consumption_logs")
        rows = cur.fetchall()
        for idx, r in enumerate(rows, 1):
            log_id = f"CLOG-{idx:06d}"
            cur.execute("""
                INSERT INTO consumption_logs_new
                (consumption_log_id, consumption_log_part_no, consumption_log_warehouse,
                 consumption_log_qty, consumption_log_user, consumption_log_timestamp)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (log_id, r[0], r[1], r[2], r[3], r[4]))
        cur.execute("DROP TABLE consumption_logs")
        cur.execute("ALTER TABLE consumption_logs_new RENAME TO consumption_logs")
        conn.commit()
        print("[OK] consumption_logs 已完成遷移")

def migrate_fixture_boms():
    with get_conn() as conn:
        cur = conn.cursor()
        if not table_exists(cur, "fixture_boms"):
            return
        cur.execute("PRAGMA table_info(fixture_boms)")
        cols = [c[1] for c in cur.fetchall()]
        if "fixture_bom_id" in cols:
            print("[SKIP] fixture_boms 已是新版")
            return

        cur.execute("""
            CREATE TABLE fixture_boms_new (
                fixture_bom_id TEXT PRIMARY KEY,
                fixture_bom_parent_no TEXT,
                fixture_bom_child_no TEXT,
                fixture_bom_qty INTEGER,
                fixture_bom_timestamp TEXT
            )
        """)
        sel_cols = []
        for c in ["parent_part_no","child_part_no","bom_qty","created_at"]:
            sel_cols.append(c if c in cols else "NULL")
        cur.execute(f"SELECT {','.join(sel_cols)} FROM fixture_boms")
        rows = cur.fetchall()
        for idx, r in enumerate(rows, 1):
            bom_id = f"BOM-{idx:06d}"
            cur.execute("""
                INSERT INTO fixture_boms_new
                (fixture_bom_id, fixture_bom_parent_no, fixture_bom_child_no,
                 fixture_bom_qty, fixture_bom_timestamp)
                VALUES (?, ?, ?, ?, ?)
            """, (bom_id, r[0], r[1], r[2], r[3]))
        cur.execute("DROP TABLE fixture_boms")
        cur.execute("ALTER TABLE fixture_boms_new RENAME TO fixture_boms")
        conn.commit()
        print("[OK] fixture_boms 已完成遷移")

def migrate_issues_to_sop():
    with get_conn() as conn:
        cur = conn.cursor()
        if not table_exists(cur, "issues"):
            return
        if table_exists(cur, "SOP"):
            print("[SKIP] SOP 已存在")
            return
        cur.execute("ALTER TABLE issues RENAME TO SOP")
        print("[OK] issues 已改名為 SOP")

def migrate_changelog_to_change_log():
    with get_conn() as conn:
        cur = conn.cursor()
        if not table_exists(cur, "changelog"):
            return
        if table_exists(cur, "change_log"):
            print("[SKIP] change_log 已存在")
            return
        cur.execute("ALTER TABLE changelog RENAME TO change_log")
        print("[OK] changelog 已改名為 change_log")

def migrate_users_view_issue_to_sop_info():
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("PRAGMA table_info(users)")
        cols = [r[1] for r in cur.fetchall()]
        if "can_view_issues" in cols and "can_view_sop_info" not in cols:
            print("[INFO] users 表存在 can_view_issues，開始轉換為 can_view_sop_info")
            cur.execute("ALTER TABLE users ADD COLUMN can_view_sop_info INTEGER DEFAULT 0")
            cur.execute("UPDATE users SET can_view_sop_info = can_view_issues")
            conn.commit()
            print("[OK] 已完成 users 欄位轉換 (can_view_issues → can_view_sop_info)")
        else:
            print("[SKIP] 不需要轉換 users 欄位")

def migrate_users_drop_old_col():
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("PRAGMA table_info(users)")
        cols = [r[1] for r in cur.fetchall()]
        if "can_view_issues" not in cols:
            print("[SKIP] users 表沒有 can_view_issues，不需要移除")
            return

        print("[INFO] users 表存在 can_view_issues，開始刪除舊欄位")

        cur.execute("""
            CREATE TABLE users_new (
                username TEXT,
                password TEXT,
                role TEXT,
                specialty TEXT,
                can_view_logs INTEGER DEFAULT 0,
                can_delete_logs INTEGER DEFAULT 0,
                can_upload_sop INTEGER DEFAULT 0,
                can_view_sop_info INTEGER DEFAULT 0,
                can_manage_users INTEGER DEFAULT 0,
                can_add INTEGER DEFAULT 0,
                can_delete INTEGER DEFAULT 0,
                active INTEGER DEFAULT 1
            )
        """)

        cur.execute("""
            INSERT INTO users_new (username, password, role, specialty,
                                   can_view_logs, can_delete_logs, can_upload_sop,
                                   can_view_sop_info, can_manage_users, can_add,
                                   can_delete, active)
            SELECT username, password, role, specialty,
                   can_view_logs, can_delete_logs, can_upload_sop,
                   can_view_sop_info, can_manage_users, can_add,
                   can_delete, active
            FROM users
        """)

        cur.execute("DROP TABLE users")
        cur.execute("ALTER TABLE users_new RENAME TO users")
        conn.commit()

        print("[OK] users 表已移除 can_view_issues")

def migrate_users_to_module_management():
    with get_conn() as conn:
        cur = conn.cursor()
        if not table_exists(cur, "users"):
            return
        if table_exists(cur, "module_management"):
            print("[SKIP] module_management 已存在")
            return
        cur.execute("ALTER TABLE users RENAME TO module_management")
        print("[OK] users 已改名為 module_management")

def migrate_sop_to_sop_information():
    with get_conn() as conn:
        cur = conn.cursor()
        if not table_exists(cur, "SOP"):
            return
        if table_exists(cur, "sop_information"):
            print("[SKIP] sop_information 已存在")
            return
        cur.execute("ALTER TABLE SOP RENAME TO sop_information")
        print("[OK] SOP 已改名為 sop_information")

if __name__ == "__main__":
    apply_db_path()
    migrate_fixture_logs()
    migrate_transfer_logs()
    migrate_consumption_logs()
    migrate_fixture_boms()
    migrate_issues_to_sop()
    migrate_changelog_to_change_log()
    migrate_users_view_issue_to_sop_info()
    migrate_users_drop_old_col()
    migrate_users_to_module_management()
    migrate_sop_to_sop_information()

    print("[DONE] 所有資料表已完成一次性遷移")