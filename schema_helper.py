# schema_helper.py
import sqlite3
import re

def get_required_columns():
    return {
        "activity_logs": {
            "product_code": "TEXT",
            "module": "TEXT"
        },
        "users": {
            "specialty": "TEXT",
            "can_view_logs": "INTEGER DEFAULT 0",
            "can_delete_logs": "INTEGER DEFAULT 0",
            "can_upload_sop": "INTEGER DEFAULT 0",
            "can_view_issues": "INTEGER DEFAULT 0",
            "can_manage_users": "INTEGER DEFAULT 0"
        },
        "issues": {
            "assembly_sop": "TEXT",
            "test_sop": "TEXT",
            "packaging_sop": "TEXT",
            "oqc_checklist": "TEXT"
        },
        "changelog": {
            "version": "TEXT NOT NULL",
            "date": "TEXT NOT NULL",
            "content": "TEXT NOT NULL"
        }
    }

def auto_add_missing_columns(db_path, schema_map):
    with sqlite3.connect(db_path, timeout=10) as conn:
        conn.execute("PRAGMA journal_mode=WAL;")
        cursor = conn.cursor()
        for table, columns in schema_map.items():
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name=?", (table,))
            if not cursor.fetchone():
                print(f"[警告] 資料表不存在：{table}，無法補欄位")
                continue

            cursor.execute(f"PRAGMA table_info({table})")
            existing = [row[1] for row in cursor.fetchall()]
            print(f"資料表 {table} 現有欄位: {existing}")

            for col_name, col_type in columns.items():
                if col_name not in existing:
                    try:
                        cursor.execute(f"ALTER TABLE {table} ADD COLUMN {col_name} {col_type}")
                        print(f"已新增欄位 {col_name} 到資料表 {table}")
                    except sqlite3.OperationalError as e:
                        print(f"欄位新增失敗 {col_name}@{table}: {e}")

def ensure_changelog_schema(db_name):
    with sqlite3.connect(db_name) as conn:
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS changelog (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                version TEXT NOT NULL,
                date TEXT NOT NULL,
                content TEXT NOT NULL
            )
        """)
        cursor.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_changelog_version ON changelog(version);")
        conn.commit()

def get_next_changelog_version(db_path):
    with sqlite3.connect(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT version FROM changelog")
        versions = cursor.fetchall()
        version_set = set(v[0] for v in versions if re.match(r"^v\d+\.\d+\.\d+$", v[0]))

        max_version = (1, 0, -1)
        for v in version_set:
            major, minor, patch = map(int, v[1:].split('.'))
            if (major, minor, patch) > max_version:
                max_version = (major, minor, patch)

        major, minor, patch = max_version

        if patch < 9:
            patch += 1
        else:
            patch = 0
            minor += 1

        next_version = f"v{major}.{minor}.{patch}"
        while next_version in version_set:
            if patch < 9:
                patch += 1
            else:
                patch = 0
                minor += 1
            next_version = f"v{major}.{minor}.{patch}"

        return next_version

def add_col_if_missing(conn, table: str, col: str, col_type: str):
    cur = conn.cursor()
    cur.execute(f"PRAGMA table_info({table})")
    cols = [row[1] for row in cur.fetchall()]
    if col not in cols:
        cur.execute(f"ALTER TABLE {table} ADD COLUMN {col} {col_type}")
    cur.close()

def ensure_fixture_schema(db_path: str):
    """建立/補齊：fixtures, warehouse_stock, transfer_logs, fixture_boms"""
    with sqlite3.connect(db_path, timeout=10) as conn:
        conn.execute("PRAGMA journal_mode=WAL;")
        cur = conn.cursor()

        cur.execute("""
            CREATE TABLE IF NOT EXISTS fixtures (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                part_no TEXT UNIQUE,
                description TEXT,
                spec TEXT,
                category TEXT,
                safety_stock INTEGER DEFAULT 0,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        add_col_if_missing(conn, "fixtures", "description", "TEXT")
        add_col_if_missing(conn, "fixtures", "spec", "TEXT")
        add_col_if_missing(conn, "fixtures", "category", "TEXT")
        add_col_if_missing(conn, "fixtures", "safety_stock", "INTEGER DEFAULT 0")
        add_col_if_missing(conn, "fixtures", "created_at", "TIMESTAMP DEFAULT CURRENT_TIMESTAMP")

        cur.execute("""
            CREATE TABLE IF NOT EXISTS warehouse_stock (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                part_no TEXT,
                warehouse TEXT,
                qty INTEGER DEFAULT 0,
                UNIQUE(part_no, warehouse)
            )
        """)
        add_col_if_missing(conn, "warehouse_stock", "part_no", "TEXT")
        add_col_if_missing(conn, "warehouse_stock", "warehouse", "TEXT")
        add_col_if_missing(conn, "warehouse_stock", "qty", "INTEGER DEFAULT 0")

        cur.execute("""
            CREATE TABLE IF NOT EXISTS transfer_logs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                part_no TEXT,
                from_warehouse TEXT,
                to_warehouse TEXT,
                qty INTEGER,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)

        cur.execute("""
            CREATE TABLE IF NOT EXISTS fixture_boms (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                parent_part_no TEXT,
                child_part_no TEXT,
                qty INTEGER DEFAULT 1,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        add_col_if_missing(conn, "fixture_boms", "parent_part_no", "TEXT")
        add_col_if_missing(conn, "fixture_boms", "child_part_no", "TEXT")
        add_col_if_missing(conn, "fixture_boms", "qty", "INTEGER DEFAULT 1")
        add_col_if_missing(conn, "fixture_boms", "created_at", "TIMESTAMP DEFAULT CURRENT_TIMESTAMP")

        conn.commit()
        cur.close()

def print_tables_info(db_path: str):
    with sqlite3.connect(db_path) as conn:
        cur = conn.cursor()
        cur.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")
        tables = [r[0] for r in cur.fetchall()]
        print(f"使用資料庫：{db_path}")
        print("資料庫初始化完成，實際位置：", db_path)
        for t in tables:
            cur.execute(f"PRAGMA table_info({t})")
            cols = [r[1] for r in cur.fetchall()]
            print(f"資料表 {t} 現有欄位:", cols)
        cur.close()
