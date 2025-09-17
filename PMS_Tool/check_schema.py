### check_schema.py
## 檢查 SQLite 資料庫中各資料表的結構是否符合預期
import sqlite3

DB_PATH = r"C:\Nelson\Dev\GitHub\PMS\PMS.db"

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
        "id","part_no","from_wh","to_wh","transfer_qty","user","remark","created_at"
    ],
    "consumption_logs": [
        "id","part_no","warehouse","consume_qty","user","remark","created_at"
    ]
}

def main():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

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
        except Exception as e:
            print("× 找不到資料表，預期欄位:", exp_cols)

    conn.close()

if __name__ == "__main__":
    main()