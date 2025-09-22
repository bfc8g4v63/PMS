#$ fixture_helper.py
#% 治具管理 DB Helper

from db_helper import get_conn, tx

CORE_WAREHOUSES = [
    "虹堡",
    "上齊",
    "睿均",
    "捷暉",
    "立榮",
    "華勤",
    "上貿",
    "麥博",
    "信利",
    "GC",
    "工程",
    "不良品"
]

def ensure_schemas():
    with get_conn() as conn:
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
            storage_location TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
        """)
        cur.execute("""
        CREATE TABLE IF NOT EXISTS warehouse_stock (
            part_no TEXT,
            warehouse TEXT,
            usable_qty INTEGER DEFAULT 0,
            safety_stock INTEGER DEFAULT 0,
            PRIMARY KEY (part_no, warehouse),
            FOREIGN KEY (part_no) REFERENCES fixtures(part_no) ON DELETE CASCADE
        )
        """)
        cur.execute("""
        CREATE TABLE IF NOT EXISTS transfer_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            part_no TEXT,
            transfer_qty INTEGER,
            from_wh TEXT,
            to_wh TEXT,
            user TEXT,
            remark TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
        """)
        cur.execute("""
        CREATE TABLE IF NOT EXISTS consumption_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            part_no TEXT,
            consume_qty INTEGER,
            warehouse TEXT,
            user TEXT,
            line TEXT,
            purpose TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
        """)
        cur.execute("""
        CREATE TABLE IF NOT EXISTS fixture_boms (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            parent_part_no TEXT,
            child_part_no TEXT,
            bom_qty INTEGER,
            remark TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
        """)

def ensure_stock_consistency():
    with tx() as conn:
        cur = conn.cursor()
        cur.execute("SELECT part_no, safety_stock, storage_location FROM fixtures")
        rows = cur.fetchall()
        for part_no, safety, location in rows:
            for wh in CORE_WAREHOUSES:
                ss = safety if wh == "虹堡" else 0
                loc = location if wh == "虹堡" else ""
                cur.execute("""
                    INSERT OR IGNORE INTO warehouse_stock(part_no, warehouse, usable_qty, safety_stock)
                    VALUES (?,?,0,?)
                """, (part_no, wh, ss))
        conn.commit()

def insert_fixture(part_no, part_name, part_spec, part_group, unit_price_ntd, safety_stock, storage_location,
                   unit_price_usd=0.0):
    if not storage_location or safety_stock is None:
        raise ValueError("建立治具必須包含儲位與安庫")

    unit_price_usd = int(unit_price_usd * 1000) / 1000.0

    with tx() as conn:
        cur = conn.cursor()
        cur.execute("SELECT 1 FROM fixtures WHERE part_no=?", (part_no,))
        if cur.fetchone():
            raise ValueError(f"治具料號 {part_no} 已存在，請勿重複新增")
        cur.execute("SELECT part_no FROM fixtures WHERE storage_location=?", (storage_location,))
        conflict = cur.fetchone()
        if conflict:
            raise ValueError(f"儲位 {storage_location} 已被料號 {conflict[0]} 使用")
        cur.execute("""
        INSERT INTO fixtures(part_no, part_name, part_spec, part_group,
                             unit_price_ntd, unit_price_usd,
                             safety_stock, storage_location)
        VALUES (?,?,?,?,?,?,?,?)
        """, (part_no, part_name, part_spec, part_group,
              unit_price_ntd, unit_price_usd,
              safety_stock, storage_location))
        for wh in CORE_WAREHOUSES:
            ss = safety_stock if wh == "虹堡" else 0
            cur.execute("""
            INSERT OR REPLACE INTO warehouse_stock(part_no, warehouse, usable_qty, safety_stock)
            VALUES (?,?,COALESCE((SELECT usable_qty FROM warehouse_stock WHERE part_no=? AND warehouse=?),0),?)
            """, (part_no, wh, part_no, wh, ss))

def delete_fixture(part_no):
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("DELETE FROM warehouse_stock WHERE part_no = ?", (part_no,))
        cur.execute("DELETE FROM fixture_boms WHERE parent_part_no = ? OR child_part_no = ?", (part_no, part_no))
        cur.execute("DELETE FROM transfer_logs WHERE part_no = ?", (part_no,))
        cur.execute("DELETE FROM consumption_logs WHERE part_no = ?", (part_no,))
        cur.execute("DELETE FROM fixtures WHERE part_no = ?", (part_no,))
        conn.commit()

def update_fixture(part_no, part_name=None, part_spec=None, part_group=None,
                   unit_price_ntd=None, unit_price_usd=None,
                   safety_stock=None, storage_location=None):
    with tx() as conn:
        cur = conn.cursor()
        cur.execute("SELECT 1 FROM fixtures WHERE part_no=?", (part_no,))
        if not cur.fetchone():
            raise ValueError(f"治具料號 {part_no} 不存在，無法修改")
        if not storage_location or storage_location.strip() == "":
            raise ValueError("修改治具時儲位不可為空")
        cur.execute("SELECT part_no FROM fixtures WHERE storage_location=? AND part_no<>?", (storage_location, part_no))
        conflict = cur.fetchone()
        if conflict:
            raise ValueError(f"儲位 {storage_location} 已被料號 {conflict[0]} 使用")
        if unit_price_usd is not None:
            unit_price_usd = int(unit_price_usd * 1000) / 1000.0
        cur.execute("""
        UPDATE fixtures
        SET part_name=?, part_spec=?, part_group=?,
            unit_price_ntd=?, unit_price_usd=?,
            safety_stock=?, storage_location=?
        WHERE part_no=?
        """, (part_name, part_spec, part_group,
              unit_price_ntd, unit_price_usd,
              safety_stock, storage_location, part_no))
        for wh in CORE_WAREHOUSES:
            ss = safety_stock if wh == "虹堡" else 0
            cur.execute("""
            UPDATE warehouse_stock
            SET safety_stock=?
            WHERE part_no=? AND warehouse=?
            """, (ss, part_no, wh))

def add_stock(part_no, qty, warehouse, user="", remark=""):
    if qty <= 0:
        raise ValueError("入庫量必須大於0")
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("UPDATE warehouse_stock SET usable_qty = usable_qty + ? WHERE part_no=? AND warehouse=?",
                    (qty, part_no, warehouse))
        conn.commit()

def transfer_stock(part_no, qty, from_wh, to_wh, user="", remark=""):
    if qty <= 0:
        raise ValueError("調撥量必須大於0")
    with tx() as conn:
        cur = conn.cursor()
        cur.execute("SELECT usable_qty FROM warehouse_stock WHERE part_no=? AND warehouse=?",
                    (part_no, from_wh))
        row = cur.fetchone()
        if not row or row[0] < qty:
            raise ValueError("來源倉庫數量不足")
        cur.execute("UPDATE warehouse_stock SET usable_qty = usable_qty - ? WHERE part_no=? AND warehouse=?",
                    (qty, part_no, from_wh))
        cur.execute("UPDATE warehouse_stock SET usable_qty = usable_qty + ? WHERE part_no=? AND warehouse=?",
                    (qty, part_no, to_wh))
        cur.execute("""
        INSERT INTO transfer_logs(part_no, transfer_qty, from_wh, to_wh, user, remark)
        VALUES (?,?,?,?,?,?)
        """, (part_no, qty, from_wh, to_wh, user, remark))

def consume_stock(part_no, qty, warehouse, user="", line="", purpose=""):
    if qty <= 0:
        raise ValueError("消耗量必須大於0")
    with tx() as conn:
        cur = conn.cursor()
        cur.execute("SELECT usable_qty FROM warehouse_stock WHERE part_no=? AND warehouse=?",
                    (part_no, warehouse))
        row = cur.fetchone()
        if not row or row[0] < qty:
            raise ValueError("倉庫數量不足，無法消耗")
        cur.execute("UPDATE warehouse_stock SET usable_qty = usable_qty - ? WHERE part_no=? AND warehouse=?",
                    (qty, part_no, warehouse))
        cur.execute("""
        INSERT INTO consumption_logs(part_no, consume_qty, warehouse, user, line, purpose)
        VALUES (?,?,?,?,?,?)
        """, (part_no, qty, warehouse, user, line, purpose))

def get_stock(part_no, warehouse):
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("SELECT usable_qty FROM warehouse_stock WHERE part_no=? AND warehouse=?",
                    (part_no, warehouse))
        row = cur.fetchone()
        return row[0] if row else 0

def get_fixture_by_part_no(part_no):
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("""
        SELECT part_no, part_name, part_spec, part_group,
               unit_price_usd, unit_price_ntd,
               safety_stock, storage_location
        FROM fixtures WHERE part_no=?
        """, (part_no,))
        rows = cur.fetchall()
        return [
            {
                "part_no": r[0],
                "part_name": r[1],
                "part_spec": r[2],
                "part_group": r[3],
                "unit_price_usd": r[4],
                "unit_price_ntd": r[5],
                "safety_stock": r[6],
                "storage_location": r[7],
            }
            for r in rows
        ]

def get_overview_by_warehouse(warehouse):
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("""
        SELECT f.part_no, f.part_name, f.part_spec, f.part_group,
               f.unit_price_usd, f.unit_price_ntd,
               CASE WHEN ?='虹堡' THEN f.safety_stock ELSE 0 END as safety_stock,
               CASE WHEN ?='虹堡' THEN f.storage_location ELSE '' END as storage_location,
               s.usable_qty
        FROM fixtures f
        JOIN warehouse_stock s ON f.part_no = s.part_no
        WHERE s.warehouse=?
        ORDER BY f.part_no
        """, (warehouse, warehouse, warehouse))
        return cur.fetchall()

def validate_location(part_no: str, raw: str) -> str:
    s = (raw or "").strip()
    parts = s.split("-")
    if len(parts) != 3:
        raise ValueError("儲位格式必須為 車-層-位置，例如 1-1-1")
    try:
        car = int(parts[0])
        layer = int(parts[1])
        pos = int(parts[2])
    except:
        raise ValueError("儲位格式必須為數字，例如 1-1-1")
    if not (1 <= car <= 9):
        raise ValueError("車號必須介於 1~9")
    if not (1 <= layer <= 4):
        raise ValueError("層號必須介於 1~4")
    if not (1 <= pos <= 50):
        raise ValueError("位置必須介於 1~50")
    formatted = f"{car}-{layer}-{pos}"
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("""
            SELECT part_no FROM fixtures
            WHERE storage_location = ? AND part_no <> ?
        """, (formatted, part_no))
        row = cur.fetchone()
        if row:
            raise ValueError(f"儲位 {formatted} 已被料號 {row[0]} 使用")
    return formatted

def generate_location(prefix: str):
    with tx() as conn:
        cur = conn.cursor()
        cur.execute("SELECT storage_location FROM fixtures WHERE storage_location IS NOT NULL")
        used = {row[0] for row in cur.fetchall()}

    if not prefix:
        for car in range(1, 10):
            for level in range(1, 5):
                for pos in range(1, 51):
                    candidate = f"{car}-{level}-{pos}"
                    if candidate not in used:
                        return candidate
        raise ValueError("所有儲位都已滿")

    if prefix.isdigit():
        car = prefix
        for level in range(1, 5):
            for pos in range(1, 51):
                candidate = f"{car}-{level}-{pos}"
                if candidate not in used:
                    return candidate
        raise ValueError(f"{car} 車已滿，無可用儲位")

    parts = prefix.split("-")
    if len(parts) == 2 and all(p.isdigit() for p in parts):
        car, level = parts
        for pos in range(1, 51):
            candidate = f"{car}-{level}-{pos}"
            if candidate not in used:
                return candidate
        raise ValueError(f"{car}-{level} 已滿，無可用儲位")

    raise ValueError("輸入格式錯誤，請輸入 車號 或 車號-層數")

def get_bom_by_part(parent_part_no: str):
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("""
        SELECT id, parent_part_no, child_part_no, bom_qty, remark
        FROM fixture_boms
        WHERE parent_part_no=?
        ORDER BY id
        """, (parent_part_no,))
        return cur.fetchall()

def add_bom_item(parent_part_no: str, child_part_no: str, qty: int, remark: str = ""):
    if qty <= 0:
        raise ValueError("BOM 數量必須大於0")
    with tx() as conn:
        cur = conn.cursor()
        cur.execute("""
        SELECT id, bom_qty FROM fixture_boms
        WHERE parent_part_no=? AND child_part_no=?
        """, (parent_part_no, child_part_no))
        row = cur.fetchone()
        if row:
            new_qty = row[1] + qty
            cur.execute("UPDATE fixture_boms SET bom_qty=?, remark=? WHERE id=?",
                        (new_qty, remark, row[0]))
        else:
            cur.execute("""
            INSERT INTO fixture_boms(parent_part_no, child_part_no, bom_qty, remark)
            VALUES (?,?,?,?)
            """, (parent_part_no, child_part_no, qty, remark))

def delete_bom_item(bom_id: int):
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("DELETE FROM fixture_boms WHERE id=?", (bom_id,))
        conn.commit()