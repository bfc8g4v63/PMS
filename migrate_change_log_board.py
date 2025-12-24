#$ migrate_change_log_board.py
#% change_log 新進板結構一次性 migration

import os
import sys
import shutil
import sqlite3
from datetime import datetime

TABLE_NAME = "change_log"
COL_VERSION = "version"
COL_DATE = "date"
COL_CONTENT = "content"

COL_RELEASE_TYPE = "change_log_release_type"
COL_SCOPE = "change_log_scope"
COL_KIND = "change_log_kind"
COL_MODULE = "change_log_module"
COL_DETAIL = "change_log_detail"

INSERT_IF_MISSING = [
    ("v3.0.9", "PMS_修正_Drop_舊schema"),
    ("v3.0.8", "PMS_修正_DB關閉與交易回滾容錯"),
    ("v3.0.7", "PMS_修正_DB初始化PRAGMA競態"),
    ("v3.0.6", "PMS_調整_DB連線管理統一入口"),
]

EXPLICIT_RELEASE_OVERRIDES = {
    "v2.0.5": "minor",
    "v2.2.1": "minor",
    "v2.2.2": "minor",
}

def backup_db(db_path):
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    bak_path = f"{db_path}.{ts}.bak"
    shutil.copy2(db_path, bak_path)
    return bak_path

def table_exists(conn, table_name):
    row = conn.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name=?",
        (table_name,),
    ).fetchone()
    return row is not None

def get_columns(conn, table_name):
    rows = conn.execute(f"PRAGMA table_info({table_name})").fetchall()
    return [r[1] for r in rows]

def add_column_if_missing(conn, table_name, col_name, col_type):
    cols = get_columns(conn, table_name)
    if col_name in cols:
        return False
    conn.execute(f"ALTER TABLE {table_name} ADD COLUMN {col_name} {col_type}")
    return True

def normalize_text(s):
    if s is None:
        return ""
    return str(s).strip()

def parse_content_board(content_text):
    s = normalize_text(content_text)
    if not s:
        return ("", "", "", "")

    parts = [p.strip() for p in s.split("_") if p.strip()]
    scope = parts[0] if len(parts) >= 1 else ""
    kind = parts[1] if len(parts) >= 2 else ""
    module = parts[2] if len(parts) >= 3 else ""
    detail = "_".join(parts[3:]).strip() if len(parts) >= 4 else ""

    if not scope and s:
        scope = "PMS"

    return (scope, kind, module, detail)

def decide_release_type(version, scope, kind, module, detail, raw_content):
    v = normalize_text(version)
    if v in EXPLICIT_RELEASE_OVERRIDES:
        return EXPLICIT_RELEASE_OVERRIDES[v]

    t = normalize_text(raw_content)
    k = normalize_text(kind)

    if "重大變更" in t:
        return "major"

    if k == "設計":
        return "major"

    if k in ("功能", "開發", "資料管理"):
        return "minor"

    if k == "系統":
        if any(x in t for x in ("資料庫遷移", "重構", "全面改用", "統一資料庫連線模式")):
            return "major"
        return "minor"

    if k in ("修正", "調整", "修改", "系統工具"):
        return "patch"

    return "patch"

def ensure_board_columns(conn):
    changed = []
    changed.append(add_column_if_missing(conn, TABLE_NAME, COL_RELEASE_TYPE, "TEXT"))
    changed.append(add_column_if_missing(conn, TABLE_NAME, COL_SCOPE, "TEXT"))
    changed.append(add_column_if_missing(conn, TABLE_NAME, COL_KIND, "TEXT"))
    changed.append(add_column_if_missing(conn, TABLE_NAME, COL_MODULE, "TEXT"))
    changed.append(add_column_if_missing(conn, TABLE_NAME, COL_DETAIL, "TEXT"))
    return any(changed)

def ensure_insert_versions(conn, today_str):
    existing = set(
        r[0] for r in conn.execute(f"SELECT {COL_VERSION} FROM {TABLE_NAME}").fetchall()
        if r and r[0] is not None
    )
    inserted = 0
    for v, c in INSERT_IF_MISSING:
        if v in existing:
            continue
        conn.execute(
            f"INSERT INTO {TABLE_NAME} ({COL_VERSION}, {COL_DATE}, {COL_CONTENT}) VALUES (?, ?, ?)",
            (v, today_str, c),
        )
        inserted += 1
    return inserted

def update_rows(conn):
    rows = conn.execute(
        f"SELECT {COL_VERSION}, {COL_CONTENT} FROM {TABLE_NAME}"
    ).fetchall()

    updated = 0
    for version, content in rows:
        scope, kind, module, detail = parse_content_board(content)
        release_type = decide_release_type(version, scope, kind, module, detail, content)

        cur = conn.execute(
            f"""
            UPDATE {TABLE_NAME}
            SET
                {COL_RELEASE_TYPE}=?,
                {COL_SCOPE}=?,
                {COL_KIND}=?,
                {COL_MODULE}=?,
                {COL_DETAIL}=?
            WHERE {COL_VERSION}=?
            """,
            (release_type, scope, kind, module, detail, version),
        )
        updated += cur.rowcount
    return updated

def count_by_release_type(conn):
    rows = conn.execute(
        f"""
        SELECT {COL_RELEASE_TYPE}, COUNT(*)
        FROM {TABLE_NAME}
        GROUP BY {COL_RELEASE_TYPE}
        ORDER BY {COL_RELEASE_TYPE}
        """
    ).fetchall()
    return rows

def list_untyped(conn):
    rows = conn.execute(
        f"""
        SELECT {COL_VERSION}, {COL_CONTENT}
        FROM {TABLE_NAME}
        WHERE {COL_RELEASE_TYPE} IS NULL OR {COL_RELEASE_TYPE}=''
        ORDER BY {COL_VERSION}
        """
    ).fetchall()
    return rows

def main():
    db_path = sys.argv[1] if len(sys.argv) >= 2 else "PMS.db"
    db_path = os.path.abspath(db_path)

    if not os.path.exists(db_path):
        print(f"DB not found: {db_path}")
        print("Usage: python migrate_change_log_board.py <path_to_db>")
        sys.exit(1)

    bak = backup_db(db_path)
    print(f"Backup created: {bak}")

    conn = sqlite3.connect(db_path, timeout=30)
    try:
        conn.execute("PRAGMA busy_timeout=5000")
        if not table_exists(conn, TABLE_NAME):
            print(f"Table not found: {TABLE_NAME}")
            sys.exit(2)

        cols = get_columns(conn, TABLE_NAME)
        need_cols = {COL_VERSION, COL_DATE, COL_CONTENT}
        if not need_cols.issubset(set(cols)):
            print(f"Missing required columns in {TABLE_NAME}")
            print("Existing columns:", ", ".join(cols))
            sys.exit(3)

        today_str = datetime.now().strftime("%Y-%m-%d")

        with conn:
            schema_changed = ensure_board_columns(conn)

        with conn:
            inserted = ensure_insert_versions(conn, today_str)

        with conn:
            updated = update_rows(conn)

        stats = count_by_release_type(conn)
        untyped = list_untyped(conn)

        print(f"Schema changed: {schema_changed}")
        print(f"Inserted versions: {inserted}")
        print(f"Updated rows: {updated}")
        print("Release type counts:")
        for rt, cnt in stats:
            print(f"  {rt}: {cnt}")

        print("Rows still untyped:")
        if not untyped:
            print("  (none)")
        else:
            for v, c in untyped:
                print(f"  {v} | {c}")

        print("Done.")
    finally:
        conn.close()

if __name__ == "__main__":
    main()