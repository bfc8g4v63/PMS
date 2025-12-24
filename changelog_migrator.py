#$ changelog_migrator.py
#% change_log -> change_log_v2 改版歷史分類進版工具（地端先測）

import argparse
import os
import shutil
import sqlite3
from datetime import datetime


DEFAULT_MODULE_KEYWORDS = [
    "Changelog", "SOP", "治具", "治具管理", "治具紀錄", "Fixture", "BOM",
    "帳號", "權限", "DB", "資料庫", "Schema", "Logs", "Log", "紀錄",
    "匯出", "Excel", "WAL", "閒置登出", "登入", "同步"
]


def now_stamp():
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def backup_db(db_path):
    if not os.path.isfile(db_path):
        raise FileNotFoundError(f"找不到資料庫檔案：{db_path}")
    backup_path = f"{db_path}.bak_{now_stamp()}"
    shutil.copy2(db_path, backup_path)
    return backup_path


def get_conn(db_path):
    conn = sqlite3.connect(db_path)
    conn.execute("PRAGMA foreign_keys = ON;")
    conn.execute("PRAGMA busy_timeout = 8000;")
    return conn


def ensure_table_v2(conn):
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS change_log_v2 (
            cl_v2_id INTEGER PRIMARY KEY AUTOINCREMENT,
            version TEXT NOT NULL UNIQUE,
            date TEXT,
            product TEXT,
            change_type TEXT,
            module TEXT,
            title TEXT,
            raw_content TEXT NOT NULL,
            parsed_ok INTEGER NOT NULL DEFAULT 1,
            created_at TEXT NOT NULL
        )
        """
    )
    conn.execute("CREATE INDEX IF NOT EXISTS idx_change_log_v2_product ON change_log_v2(product);")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_change_log_v2_type ON change_log_v2(change_type);")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_change_log_v2_module ON change_log_v2(module);")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_change_log_v2_date ON change_log_v2(date);")


def table_exists(conn, table_name):
    row = conn.execute(
        "SELECT 1 FROM sqlite_master WHERE type='table' AND name=?",
        (table_name,)
    ).fetchone()
    return row is not None


def safe_split_content(raw):
    if raw is None:
        return []
    raw = str(raw).strip()
    if not raw:
        return []
    return [p.strip() for p in raw.split("_") if p.strip()]


def normalize_change_type(s):
    if not s:
        return None
    if s in ("修正", "修復", "fix", "Fix", "BUG", "Bug"):
        return "修正"
    if s in ("功能", "新增", "feature", "Feature"):
        return "功能"
    if s in ("重構", "Refactor", "refactor"):
        return "重構"
    if s in ("優化", "改善", "Optimize", "optimize"):
        return "優化"
    return s


def detect_module(tokens, module_keywords):
    if not tokens:
        return "General"
    joined = "_".join(tokens)
    for kw in module_keywords:
        if kw and kw in joined:
            return kw
    first = tokens[0]
    if first:
        return first
    return "General"


def parse_content(raw_content, module_keywords):
    parts = safe_split_content(raw_content)
    product = None
    change_type = None
    module = "General"
    title = None
    parsed_ok = 1
    if len(parts) >= 1:
        product = parts[0]
    if len(parts) >= 2:
        change_type = normalize_change_type(parts[1])
    rest = parts[2:] if len(parts) >= 3 else []
    if not product and not change_type and not rest:
        parsed_ok = 0
        return {
            "product": None,
            "change_type": None,
            "module": "General",
            "title": (raw_content or "").strip(),
            "parsed_ok": 0
        }
    if rest:
        module = detect_module(rest, module_keywords)
        title = "_".join(rest).strip()
    else:
        title = (raw_content or "").strip()
        module = "General"
        parsed_ok = 0
    return {
        "product": product,
        "change_type": change_type,
        "module": module,
        "title": title,
        "parsed_ok": parsed_ok
    }


def rebuild_v2(conn):
    conn.execute("DELETE FROM change_log_v2;")
    conn.execute("DELETE FROM sqlite_sequence WHERE name='change_log_v2';")


def fetch_source_rows(conn):
    return conn.execute(
        "SELECT version, date, content FROM change_log ORDER BY date ASC, version ASC"
    ).fetchall()


def upsert_v2_row(conn, version, date, parsed, raw_content):
    conn.execute(
        """
        INSERT INTO change_log_v2
        (version, date, product, change_type, module, title, raw_content, parsed_ok, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(version) DO UPDATE SET
            date=excluded.date,
            product=excluded.product,
            change_type=excluded.change_type,
            module=excluded.module,
            title=excluded.title,
            raw_content=excluded.raw_content,
            parsed_ok=excluded.parsed_ok
        """,
        (
            version,
            date,
            parsed.get("product"),
            parsed.get("change_type"),
            parsed.get("module"),
            parsed.get("title"),
            raw_content if raw_content is not None else "",
            int(parsed.get("parsed_ok", 1)),
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        )
    )


def print_summary(conn):
    total = conn.execute("SELECT COUNT(1) FROM change_log_v2").fetchone()[0]
    print(f"change_log_v2 總筆數：{total}")
    print("")
    print("依 product 統計：")
    for row in conn.execute(
        "SELECT COALESCE(product,'(NULL)') AS p, COUNT(1) FROM change_log_v2 GROUP BY p ORDER BY COUNT(1) DESC, p ASC"
    ).fetchall():
        print(f"{row[0]}: {row[1]}")
    print("")
    print("依 change_type 統計：")
    for row in conn.execute(
        "SELECT COALESCE(change_type,'(NULL)') AS t, COUNT(1) FROM change_log_v2 GROUP BY t ORDER BY COUNT(1) DESC, t ASC"
    ).fetchall():
        print(f"{row[0]}: {row[1]}")
    print("")
    print("依 module 統計（前 30 名）：")
    for row in conn.execute(
        """
        SELECT COALESCE(module,'(NULL)') AS m, COUNT(1)
        FROM change_log_v2
        GROUP BY m
        ORDER BY COUNT(1) DESC, m ASC
        LIMIT 30
        """
    ).fetchall():
        print(f"{row[0]}: {row[1]}")
    print("")
    bad = conn.execute("SELECT COUNT(1) FROM change_log_v2 WHERE parsed_ok=0").fetchone()[0]
    print(f"解析不完整(parsed_ok=0) 筆數：{bad}")


def run(db_path, rebuild, no_backup, module_keywords):
    if not no_backup:
        backup_path = backup_db(db_path)
        print(f"已備份：{backup_path}")
    with get_conn(db_path) as conn:
        if not table_exists(conn, "change_log"):
            raise RuntimeError("找不到 change_log 表，請先確認 DB 是否為 PMS 使用的資料庫。")
        ensure_table_v2(conn)
        if rebuild:
            rebuild_v2(conn)
        rows = fetch_source_rows(conn)
        for version, date, content in rows:
            raw = content if content is not None else ""
            parsed = parse_content(raw, module_keywords)
            upsert_v2_row(conn, version, date, parsed, raw)
        conn.commit()
        print_summary(conn)


def build_arg_parser():
    p = argparse.ArgumentParser()
    p.add_argument("--db", default="PMS.db")
    p.add_argument("--rebuild", action="store_true")
    p.add_argument("--no-backup", action="store_true")
    p.add_argument("--module-keyword", action="append", default=None)
    return p


def main():
    args = build_arg_parser().parse_args()
    module_keywords = DEFAULT_MODULE_KEYWORDS
    if args.module_keyword:
        module_keywords = args.module_keyword
    run(
        db_path=args.db,
        rebuild=args.rebuild,
        no_backup=args.no_backup,
        module_keywords=module_keywords
    )


if __name__ == "__main__":
    main()