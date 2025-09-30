#$ schema_helper.py
#% 資料表結構管理

import re
from db_helper import get_conn, DB_PATH

def get_required_columns():
    return {
        "activity_logs": {
            "activity_log_id": "INTEGER",
            "username": "TEXT",
            "action": "TEXT",
            "filename": "TEXT",
            "timestamp": "TEXT",
            "module": "TEXT"
        },
        "module_management": {
            "username": "TEXT",
            "password": "TEXT",
            "role": "TEXT",
            "specialty": "TEXT",
            "can_add": "INTEGER",
            "can_delete": "INTEGER",
            "can_view_logs": "INTEGER",
            "can_delete_logs": "INTEGER",
            "can_upload_sop": "INTEGER",
            "can_view_sop_info": "INTEGER",
            "can_manage_users": "INTEGER",
            "active": "INTEGER"
        },
        "sop_information": {
            "product_code": "TEXT",
            "product_name": "TEXT",
            "dip_sop": "TEXT",
            "assembly_sop": "TEXT",
            "test_sop": "TEXT",
            "packaging_sop": "TEXT",
            "oqc_checklist": "TEXT",
            "created_by": "TEXT",
            "created_at": "TEXT",
            "dip_sop_bypass": "INTEGER",
            "assembly_sop_bypass": "INTEGER",
            "test_sop_bypass": "INTEGER",
            "packaging_sop_bypass": "INTEGER",
            "oqc_checklist_bypass": "INTEGER"
        },
        "change_log": {
            "version": "TEXT",
            "date": "TEXT",
            "content": "TEXT"
        }
    }

def auto_add_missing_columns(db_path, schema_map, verbose=False):
    with get_conn(db_path) as conn:
        cursor = conn.cursor()
        for table, columns in schema_map.items():
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name=?", (table,))
            if not cursor.fetchone():
                if verbose:
                    print(f"[警告] 資料表不存在：{table}，無法補欄位")
                continue
            cursor.execute(f"PRAGMA table_info({table})")
            existing = [row[1] for row in cursor.fetchall()]
            for col_name, col_type in columns.items():
                if col_name not in existing:
                    try:
                        cursor.execute(f"ALTER TABLE {table} ADD COLUMN {col_name} {col_type}")
                        if verbose:
                            print(f"已新增欄位 {col_name} 到資料表 {table}")
                    except Exception as e:
                        if verbose:
                            print(f"欄位新增失敗 {col_name}@{table}: {e}")

def ensure_changelog_schema(db_path=None, verbose=False):
    with get_conn(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS change_log (
                version TEXT PRIMARY KEY,
                date TEXT NOT NULL,
                content TEXT NOT NULL
            )
        """)
        cursor.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_changelog_version ON change_log(version);")
        conn.commit()

def get_next_changelog_version(db_path=None):
    with get_conn(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT DISTINCT version FROM change_log")
        versions = [row[0] for row in cursor.fetchall()]

    pattern = re.compile(r"^v(\d+)\.(\d+)\.(\d+)$")
    triplets = []
    for v in versions:
        m = pattern.match(v or "")
        if m:
            triplets.append(tuple(map(int, m.groups())))

    if not triplets:
        return "v1.0.0"

    major, minor, patch = max(triplets)

    if patch < 9:
        patch += 1
    else:
        patch = 0
        if minor < 9:
            minor += 1
        else:
            minor = 0
            major += 1

    next_version = f"v{major}.{minor}.{patch}"

    existing = set(v for v in versions if pattern.match(v or ""))
    while next_version in existing:
        if patch < 9:
            patch += 1
        else:
            patch = 0
            if minor < 9:
                minor += 1
            else:
                minor = 0
                major += 1
        next_version = f"v{major}.{minor}.{patch}"

    return next_version

def add_col_if_missing(conn, table: str, col: str, col_type: str):
    cur = conn.cursor()
    cur.execute(f"PRAGMA table_info({table})")
    cols = [row[1] for row in cur.fetchall()]
    if col not in cols:
        cur.execute(f"ALTER TABLE {table} ADD COLUMN {col} {col_type}")
    cur.close()

def ensure_fixture_schema(db_path=None, verbose=False):
    with get_conn(db_path) as conn:
        cur = conn.cursor()

        cur.execute("""
            CREATE TABLE IF NOT EXISTS fixtures (
                part_no TEXT PRIMARY KEY,
                part_name TEXT,
                part_spec TEXT,
                part_group TEXT,
                unit_price_ntd REAL,
                unit_price_usd REAL,
                safety_stock INTEGER,
                storage_location TEXT
            )
        """)
        if verbose:
            cur.execute("PRAGMA table_info(fixtures)")
            print("fixtures:", [r[1] for r in cur.fetchall()])

        cur.execute("""
            CREATE TABLE IF NOT EXISTS warehouse_stock (
                part_no TEXT,
                warehouse TEXT,
                usable_qty INTEGER DEFAULT 0,
                safety_stock INTEGER DEFAULT 0,
                PRIMARY KEY (part_no, warehouse)
            )
        """)
        if verbose:
            cur.execute("PRAGMA table_info(warehouse_stock)")
            print("warehouse_stock:", [r[1] for r in cur.fetchall()])

        cur.execute("""
            CREATE TABLE IF NOT EXISTS transfer_logs (
                transfer_log_id INTEGER PRIMARY KEY AUTOINCREMENT,
                part_no TEXT,
                from_wh TEXT,
                to_wh TEXT,
                transfer_qty INTEGER,
                username TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        if verbose:
            cur.execute("PRAGMA table_info(transfer_logs)")
            print("transfer_logs:", [r[1] for r in cur.fetchall()])

        cur.execute("""
            CREATE TABLE IF NOT EXISTS fixture_boms (
                bom_id INTEGER PRIMARY KEY AUTOINCREMENT,
                parent_part_no TEXT,
                child_part_no TEXT,
                bom_qty INTEGER DEFAULT 1,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        if verbose:
            cur.execute("PRAGMA table_info(fixture_boms)")
            print("fixture_boms:", [r[1] for r in cur.fetchall()])

        conn.commit()
        cur.close()

def print_tables_info(db_path: str):
    with get_conn(db_path) as conn:
        cur = conn.cursor()
        cur.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")
        tables = [r[0] for r in cur.fetchall()]
        print(f"使用資料庫（參數）：{db_path}")
        print("實際使用資料庫（db_helper.DB_PATH）：", DB_PATH)
        for t in tables:
            cur.execute(f"PRAGMA table_info({t})")
            print(f"{t}:", [r[1] for r in cur.fetchall()])
        cur.close()