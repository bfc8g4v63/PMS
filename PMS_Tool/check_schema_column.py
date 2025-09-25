#$ check_schema_column.py
#% 檢查 SQLite 資料庫中各資料表與欄位結構是否符合預期

import sys
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parents[1]
sys.path.append(str(ROOT_DIR))

from config import apply_db_path
from db_helper import get_conn

apply_db_path()

expected_schema = {
    "fixtures": [
        "part_no","part_name","part_spec","part_group",
        "unit_price_ntd","unit_price_usd",
        "safety_stock","storage_location"
    ],
    "warehouse_stock": [
        "part_no","warehouse","usable_qty","safety_stock"
    ],
    "issues": [
        "product_code","product_name",
        "dip_sop","assembly_sop","test_sop","packaging_sop","oqc_checklist"
    ],
    "users": [
        "username","password","role","specialty",
        "can_view_logs","can_delete_logs","can_upload_sop",
        "can_view_issues","can_manage_users"
    ],
    "activity_logs": [
        "id","username","action","filename","timestamp","module"
    ],
    "transfer_logs": [
        "id","part_no","from_wh","to_wh","transfer_qty","user","created_at"
    ],
    "consumption_logs": [
        "id","part_no","warehouse","consume_qty","user","created_at"
    ]
}

def main():
    with get_conn() as conn:
        cur = conn.cursor()

        cur.execute("SELECT name FROM sqlite_master WHERE type='table'")
        all_tables = [r[0] for r in cur.fetchall()]
        expected_tables = list(expected_schema.keys())

        extra_tables = [t for t in all_tables if t not in expected_tables]
        missing_tables = [t for t in expected_tables if t not in all_tables]

        print("=== 資料表整體檢查 ===")
        print("資料庫中所有表:", all_tables)
        print("預期應存在表:", expected_tables)

        if extra_tables:
            print("多餘的資料表:", extra_tables)
        if missing_tables:
            print("缺少的資料表:", missing_tables)

        for table, exp_cols in expected_schema.items():
            print(f"\n=== {table} ===")
            try:
                cur.execute(f"PRAGMA table_info({table})")
                cols = [r[1] for r in cur.fetchall()]
                print("實際欄位:", cols)
                print("預期欄位:", exp_cols)

                missing = [c for c in exp_cols if c not in cols]
                extra = [c for c in cols if c not in exp_cols]

                if missing:
                    print("缺少欄位:", missing)
                if extra:
                    print("多餘欄位:", extra)

                if not missing and not extra:
                    print("✓ 結構正確")
            except Exception:
                print("× 找不到資料表，預期欄位:", exp_cols)

if __name__ == "__main__":
    main()