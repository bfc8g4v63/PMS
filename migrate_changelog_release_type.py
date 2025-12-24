import os
import sys
import shutil
import sqlite3
from datetime import datetime

RELEASE_TYPE_COL = "change_log_release_type"

MAJOR_VERSIONS = {
    "v3.0.9",
    "v3.0.6",
    "v2.8.0",
    "v2.3.6",
    "v1.8.8",
    "v1.8.0",
}

MINOR_VERSIONS = {
    "v3.0.1",
    "v3.0.0",
    "v2.9.4",
    "v2.9.3",
    "v2.9.2",
    "v2.9.1",
    "v2.9.0",
    "v2.8.9",
    "v2.8.8",
    "v2.8.7",
    "v2.8.6",
    "v2.8.5",
    "v2.8.4",
    "v2.8.3",
    "v2.8.2",
    "v2.8.1",
    "v2.7.9",
    "v2.7.5",
    "v2.7.4",
    "v2.7.1",
    "v2.7.0",
    "v2.6.6",
    "v2.6.4",
    "v2.6.3",
    "v2.6.2",
    "v2.6.1",
    "v2.6.0",
    "v2.5.7",
    "v2.5.6",
    "v2.5.4",
    "v2.5.2",
    "v2.4.9",
    "v2.3.5",
    "v2.3.4",
    "v2.3.3",
    "v2.3.0",
    "v2.2.9",
    "v2.2.8",
    "v2.2.7",
    "v2.2.6",
    "v2.1.7",
    "v2.0.9",
    "v2.0.8",
    "v2.0.6",
    "v2.0.4",
    "v1.9.5",
    "v1.9.4",
    "v1.8.7",
    "v1.7.6",
    "v1.7.5",
    "v1.6.2",
    "v1.6.1",
    "v1.6.0",
    "v1.5.9",
    "v1.5.7",
    "v1.5.6",
    "v1.5.3",
    "v1.5.1",
    "v1.4.2",
    "v1.4.1",
    "v1.4.0",
    "v1.3.8",
    "v1.3.7",
    "v1.3.6",
    "v1.3.5",
    "v1.3.2",
    "v1.2.9",
    "v1.2.8",
    "v1.2.7",
    "v1.2.6",
    "v1.2.4",
    "v1.2.1",
    "v1.1.8",
    "v1.1.7",
    "v1.1.6",
    "v1.1.4",
    "v1.1.3",
    "v1.0.8",
    "v1.0.7",
    "v1.0.6",
}

PATCH_VERSIONS = {
    "v3.0.8",
    "v3.0.7",
    "v3.0.5",
    "v3.0.4",
    "v3.0.3",
    "v3.0.2",
    "v2.9.9",
    "v2.9.8",
    "v2.9.7",
    "v2.9.6",
    "v2.9.5",
    "v2.7.8",
    "v2.7.7",
    "v2.7.6",
    "v2.7.3",
    "v2.7.2",
    "v2.6.9",
    "v2.6.8",
    "v2.6.7",
    "v2.6.5",
    "v2.5.9",
    "v2.5.8",
    "v2.5.5",
    "v2.5.3",
    "v2.5.1",
    "v2.5.0",
    "v2.4.8",
    "v2.4.7",
    "v2.4.6",
    "v2.4.5",
    "v2.4.4",
    "v2.4.3",
    "v2.4.2",
    "v2.4.1",
    "v2.4.0",
    "v2.3.9",
    "v2.3.8",
    "v2.3.7",
    "v2.3.2",
    "v2.3.1",
    "v2.2.5",
    "v2.2.4",
    "v2.2.3",
    "v2.2.0",
    "v2.1.9",
    "v2.1.8",
    "v2.1.6",
    "v2.1.5",
    "v2.1.4",
    "v2.1.3",
    "v2.1.2",
    "v2.1.1",
    "v2.1.0",
    "v2.0.7",
    "v2.0.3",
    "v2.0.2",
    "v2.0.1",
    "v2.0.0",
    "v1.9.9",
    "v1.9.8",
    "v1.9.7",
    "v1.9.6",
    "v1.9.3",
    "v1.9.2",
    "v1.9.1",
    "v1.9.0",
    "v1.8.9",
    "v1.8.6",
    "v1.8.5",
    "v1.8.4",
    "v1.8.3",
    "v1.8.2",
    "v1.8.1",
    "v1.7.9",
    "v1.7.8",
    "v1.7.7",
    "v1.7.4",
    "v1.7.3",
    "v1.7.2",
    "v1.7.1",
    "v1.7.0",
    "v1.6.9",
    "v1.6.8",
    "v1.6.7",
    "v1.6.6",
    "v1.6.5",
    "v1.6.4",
    "v1.6.3",
    "v1.5.8",
    "v1.5.5",
    "v1.5.4",
    "v1.5.2",
    "v1.5.0",
    "v1.4.9",
    "v1.4.8",
    "v1.4.7",
    "v1.4.6",
    "v1.4.5",
    "v1.4.4",
    "v1.4.3",
    "v1.3.9",
    "v1.3.4",
    "v1.3.3",
    "v1.3.1",
    "v1.3.0",
    "v1.2.5",
    "v1.2.3",
    "v1.2.2",
    "v1.2.0",
    "v1.1.9",
    "v1.1.5",
    "v1.1.2",
    "v1.1.1",
    "v1.1.0",
    "v1.0.9",
    "v1.0.5",
    "v1.0.4",
    "v1.0.3",
    "v1.0.2",
    "v1.0.1",
    "v1.0.0",
}

def build_release_type_map():
    m = {}
    for v in MAJOR_VERSIONS:
        m[v] = "major"
    for v in MINOR_VERSIONS:
        if v in m and m[v] != "minor":
            raise ValueError(f"Duplicate classification: {v}")
        m[v] = "minor"
    for v in PATCH_VERSIONS:
        if v in m and m[v] != "patch":
            raise ValueError(f"Duplicate classification: {v}")
        m[v] = "patch"
    return m

def backup_db(db_path):
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    bak_path = f"{db_path}.{ts}.bak"
    shutil.copy2(db_path, bak_path)
    return bak_path

def table_exists(conn, table_name):
    cur = conn.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name=?",
        (table_name,),
    )
    return cur.fetchone() is not None

def get_columns(conn, table_name):
    cur = conn.execute(f"PRAGMA table_info({table_name})")
    return [r[1] for r in cur.fetchall()]

def ensure_release_type_column(conn):
    cols = get_columns(conn, "change_log")
    if RELEASE_TYPE_COL in cols:
        return False
    conn.execute(f"ALTER TABLE change_log ADD COLUMN {RELEASE_TYPE_COL} TEXT")
    return True

def fetch_versions(conn):
    cur = conn.execute("SELECT version FROM change_log")
    rows = cur.fetchall()
    return [r[0] for r in rows if r and r[0]]

def apply_updates(conn, release_type_map):
    updated = 0
    for version, rtype in release_type_map.items():
        cur = conn.execute(
            f"UPDATE change_log SET {RELEASE_TYPE_COL}=? WHERE version=?",
            (rtype, version),
        )
        updated += cur.rowcount
    return updated

def count_by_type(conn):
    cur = conn.execute(
        f"SELECT {RELEASE_TYPE_COL}, COUNT(*) FROM change_log GROUP BY {RELEASE_TYPE_COL} ORDER BY {RELEASE_TYPE_COL}"
    )
    return cur.fetchall()

def main():
    db_path = sys.argv[1] if len(sys.argv) >= 2 else "PMS.db"
    db_path = os.path.abspath(db_path)

    if not os.path.exists(db_path):
        print(f"DB not found: {db_path}")
        print("Usage: python migrate_change_log_release_type.py <path_to_db>")
        sys.exit(1)

    bak_path = backup_db(db_path)
    print(f"Backup created: {bak_path}")

    release_type_map = build_release_type_map()

    conn = sqlite3.connect(db_path, timeout=30)
    try:
        conn.execute("PRAGMA busy_timeout=5000")
        conn.execute("PRAGMA foreign_keys=ON")

        if not table_exists(conn, "change_log"):
            print("Table not found: change_log")
            sys.exit(2)

        col_added = False
        with conn:
            col_added = ensure_release_type_column(conn)

        db_versions = set(fetch_versions(conn))
        mapping_versions = set(release_type_map.keys())

        in_db_not_in_mapping = sorted(db_versions - mapping_versions)
        in_mapping_not_in_db = sorted(mapping_versions - db_versions)

        with conn:
            updated_rows = apply_updates(conn, release_type_map)

        stats = count_by_type(conn)

        print(f"Column added: {col_added}")
        print(f"Updated rows: {updated_rows}")
        print("Type counts:")
        for rtype, cnt in stats:
            print(f"  {rtype}: {cnt}")

        print("Versions in DB but not in mapping:")
        if in_db_not_in_mapping:
            for v in in_db_not_in_mapping:
                print(f"  {v}")
        else:
            print("  (none)")

        print("Versions in mapping but not in DB:")
        if in_mapping_not_in_db:
            for v in in_mapping_not_in_db:
                print(f"  {v}")
        else:
            print("  (none)")

        print("Done.")
    finally:
        conn.close()

if __name__ == "__main__":
    main()