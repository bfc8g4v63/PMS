#$ changelog_migrator_cl3.py
#% change_log_v2 -> change_log_v3 分類收斂版本（CL3）

import os
import shutil
import sqlite3
from datetime import datetime


DB_PATH = "PMS.db"


CATEGORY_MAP = {
    "修正": "修正",
    "修改": "修正",
    "功能": "功能",
    "設計": "設計",
    "優化": "優化",
    "重構": "重構",
    "開發": "功能",
}


def backup_db(db_path):
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = f"{db_path}.bak_cl3_{stamp}"
    shutil.copy2(db_path, backup_path)
    print(f"已備份：{backup_path}")


def get_conn(db_path):
    conn = sqlite3.connect(db_path)
    conn.execute("PRAGMA busy_timeout = 8000;")
    return conn


def ensure_table(conn):
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS change_log_v3 (
            cl_v3_id INTEGER PRIMARY KEY AUTOINCREMENT,
            version TEXT UNIQUE,
            date TEXT,
            product TEXT,
            change_category TEXT,
            subsystem TEXT,
            module TEXT,
            title TEXT,
            raw_content TEXT,
            created_at TEXT
        )
        """
    )


def normalize_category(v2_type, raw, title):
    if v2_type in CATEGORY_MAP:
        return CATEGORY_MAP[v2_type]

    text = f"{raw or ''} {title or ''}"
    t = text.lower()

    if "修正" in text or "修復" in text or "bug" in t or "fix" in t:
        return "修正"
    if "功能" in text or "新增" in text or "feature" in t:
        return "功能"
    if "重構" in text or "refactor" in t:
        return "重構"
    if "優化" in text or "改善" in text or "optimize" in t:
        return "優化"
    if "設計" in text:
        return "設計"

    return "其他"


def detect_subsystem(v2_type, raw):
    if v2_type and "SafeStock" in v2_type:
        return "SafeStock"
    if raw and "SafeStock" in raw:
        return "SafeStock"
    if v2_type and "系統工具" in v2_type:
        return "系統工具"
    if raw and "系統工具" in raw:
        return "系統工具"
    return "PMS"


def normalize_module(v2_module, raw, title):
    text = " ".join([str(v2_module or ""), str(raw or ""), str(title or "")])
    t = text.lower()

    if "治具" in text:
        return "治具"
    if "sop" in t or "SOP" in text:
        return "SOP"
    if "帳號" in text or "權限" in text or "user" in t or "role" in t:
        return "帳號"
    if "登入" in text or "login" in t or "logout" in t or "閒置登出" in text:
        return "登入"
    if "bom" in t or "BOM" in text:
        return "BOM"
    if "excel" in t or "匯出" in text or "xlsx" in t:
        return "Excel"
    if "scrollbar" in t or "解析度" in text or "容器" in text or "顯示" in text or "ui" in t or "gui" in t or "tkinter" in t or "pyqt" in t or "介面" in text:
        return "GUI"
    if "schema" in t or "欄位" in text or "初始化" in text or "資料庫" in text or "sql" in t or "wal" in t or "pragma" in t or "checkpoint" in t or "shm" in t or " db" in t or "db " in t:
        return "資料庫"
    if "同步" in text or "sync" in t or "unc" in t or "z:" in t or "掛載" in text or "路徑" in text:
        return "同步"
    if "紀錄" in text or "log" in t:
        return "紀錄"

    return "Other"


def migrate():
    if not os.path.isfile(DB_PATH):
        raise FileNotFoundError(f"找不到資料庫檔案：{DB_PATH}")

    backup_db(DB_PATH)

    with get_conn(DB_PATH) as conn:
        ensure_table(conn)

        rows = conn.execute(
            """
            SELECT version, date, product, change_type, module, title, raw_content
            FROM change_log_v2
            ORDER BY date ASC
            """
        ).fetchall()

        for version, date, product, v2_type, v2_module, title, raw in rows:
            change_category = normalize_category(v2_type, raw, title)
            subsystem = detect_subsystem(v2_type, raw)
            module = normalize_module(v2_module, raw, title)

            conn.execute(
                """
                INSERT INTO change_log_v3
                (version, date, product, change_category, subsystem, module, title, raw_content, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(version) DO UPDATE SET
                    date=excluded.date,
                    product=excluded.product,
                    change_category=excluded.change_category,
                    subsystem=excluded.subsystem,
                    module=excluded.module,
                    title=excluded.title,
                    raw_content=excluded.raw_content
                """,
                (
                    version,
                    date,
                    product,
                    change_category,
                    subsystem,
                    module,
                    title,
                    raw,
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                ),
            )

        conn.commit()

        print("CL3 migrate 完成")

        print("")
        print("change_category 統計：")
        for row in conn.execute(
            "SELECT change_category, COUNT(1) FROM change_log_v3 GROUP BY change_category ORDER BY COUNT(1) DESC, change_category"
        ).fetchall():
            print(row)

        print("")
        print("subsystem 統計：")
        for row in conn.execute(
            "SELECT subsystem, COUNT(1) FROM change_log_v3 GROUP BY subsystem ORDER BY COUNT(1) DESC, subsystem"
        ).fetchall():
            print(row)

        print("")
        print("module 統計：")
        for row in conn.execute(
            "SELECT module, COUNT(1) FROM change_log_v3 GROUP BY module ORDER BY COUNT(1) DESC, module"
        ).fetchall():
            print(row)


if __name__ == "__main__":
    migrate()