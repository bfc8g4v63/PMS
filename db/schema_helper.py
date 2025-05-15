import sqlite3

def get_required_columns():
    return {
        "logs_activity": {
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
        }
    }

def auto_add_missing_columns(db_path, schema_map):
    with sqlite3.connect(db_path) as conn:
        cursor = conn.cursor()
        for table, columns in schema_map.items():
            cursor.execute(f"PRAGMA table_info({table})")
            existing = [row[1] for row in cursor.fetchall()]
            print(f"🔍 資料表 {table} 現有欄位: {existing}")

            for col_name, col_type in columns.items():
                if col_name not in existing:
                    try:
                        cursor.execute(f"ALTER TABLE {table} ADD COLUMN {col_name} {col_type}")
                        print(f"✅ 已新增欄位 {col_name} 到資料表 {table}")
                    except sqlite3.OperationalError as e:
                        print(f"⚠️ 欄位新增失敗 {col_name}@{table}: {e}")