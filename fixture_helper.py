# fixture_helper.py
from __future__ import annotations
import sqlite3
from contextlib import contextmanager
from typing import Optional, Dict, Any, Tuple

DB_PATH: Optional[str] = None

def set_db_path(path: str) -> None:
    global DB_PATH
    DB_PATH = path

CORE_WAREHOUSES = [
    "治具室", "上齊", "睿均", "捷暉", "立榮",
    "華勤", "上貿", "麥博", "GC", "不良品"
]

CONSUMED_WAREHOUSE = "消耗"

def _validate_part_no(part_no: str) -> None:
    part_no = (part_no or "").strip()
    if len(part_no) not in (8, 12) or not part_no.isdigit():
        raise ValueError("料號必須為 8 或 12 碼數字。")

@contextmanager
def get_conn(db_path: Optional[str] = None):
    db = db_path or DB_PATH
    if not db:
        raise RuntimeError("DB_PATH 尚未設定，請先呼叫 fixture_helper.set_db_path(DB_NAME)。")
    conn = sqlite3.connect(db, timeout=15)
    try:
        conn.execute("PRAGMA journal_mode=WAL;")
        conn.execute("PRAGMA foreign_keys=ON;")
        conn.row_factory = sqlite3.Row
        yield conn
        conn.commit()
        conn.execute("PRAGMA wal_checkpoint(TRUNCATE);")
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

def ensure_schemas() -> None:
    with get_conn() as conn:
        conn.execute("""
        CREATE TABLE IF NOT EXISTS fixtures (
            fixture_id   INTEGER PRIMARY KEY AUTOINCREMENT,
            part_no      TEXT NOT NULL,
            name         TEXT,
            type         TEXT,
            spec         TEXT,
            safe_stock   INTEGER DEFAULT 0,
            created_at   TEXT DEFAULT (datetime('now','localtime')),
            UNIQUE(part_no)
        );""")
        conn.execute("""
        CREATE TABLE IF NOT EXISTS warehouse_stock (
            stock_id     INTEGER PRIMARY KEY AUTOINCREMENT,
            fixture_id   INTEGER NOT NULL,
            warehouse    TEXT NOT NULL,
            qty          INTEGER NOT NULL DEFAULT 0,
            updated_at   TEXT DEFAULT (datetime('now','localtime')),
            FOREIGN KEY(fixture_id) REFERENCES fixtures(fixture_id),
            UNIQUE(fixture_id, warehouse)
        );""")
        conn.execute("""
        CREATE TABLE IF NOT EXISTS transfer_logs (
            log_id       INTEGER PRIMARY KEY AUTOINCREMENT,
            fixture_id   INTEGER NOT NULL,
            action       TEXT NOT NULL,
            from_wh      TEXT,
            to_wh        TEXT,
            qty          INTEGER NOT NULL,
            user         TEXT,
            remark       TEXT,
            created_at   TEXT DEFAULT (datetime('now','localtime')),
            FOREIGN KEY(fixture_id) REFERENCES fixtures(fixture_id)
        );""")
        conn.execute("""
        CREATE TABLE IF NOT EXISTS consumption_logs (
            consume_id   INTEGER PRIMARY KEY AUTOINCREMENT,
            fixture_id   INTEGER NOT NULL,
            qty          INTEGER NOT NULL,
            user         TEXT,
            line         TEXT,
            purpose      TEXT,
            created_at   TEXT DEFAULT (datetime('now','localtime')),
            FOREIGN KEY(fixture_id) REFERENCES fixtures(fixture_id)
        );""")
        conn.execute("""
        CREATE TABLE IF NOT EXISTS fixture_boms (
            bom_id          INTEGER PRIMARY KEY AUTOINCREMENT,
            parent_part_no  TEXT NOT NULL,
            child_part_no   TEXT NOT NULL,
            description     TEXT,
            qty             INTEGER NOT NULL DEFAULT 1,
            remark          TEXT,
            created_at      TEXT DEFAULT (datetime('now','localtime'))
        );
        """)
        def add_col_if_missing(table: str, col: str, ddl: str):
            cols = {r["name"] for r in conn.execute(f"PRAGMA table_info({table})")}
            if col not in cols:
                conn.execute(f"ALTER TABLE {table} ADD COLUMN {ddl};")

        add_col_if_missing("fixtures", "name", "name TEXT")
        add_col_if_missing("fixtures", "type", "type TEXT")
        add_col_if_missing("fixtures", "spec", "spec TEXT")
        add_col_if_missing("fixtures", "safe_stock", "safe_stock INTEGER DEFAULT 0")
        add_col_if_missing("fixtures", "created_at", "created_at TEXT DEFAULT (datetime('now','localtime'))")
        add_col_if_missing("warehouse_stock", "updated_at", "updated_at TEXT DEFAULT (datetime('now','localtime'))")

def get_fixture_by_part_no(part_no: str) -> Optional[sqlite3.Row]:
    _validate_part_no(part_no)
    with get_conn() as conn:
        cur = conn.execute("SELECT * FROM fixtures WHERE part_no = ?;", (part_no,))
        return cur.fetchone()

def _ensure_stock_row(conn: sqlite3.Connection, fixture_id: int, warehouse: str) -> None:
    conn.execute("""
        INSERT OR IGNORE INTO warehouse_stock (fixture_id, warehouse, qty)
        VALUES (?, ?, 0);
    """, (fixture_id, warehouse))

def _get_stock_for_update(conn: sqlite3.Connection, fixture_id: int, warehouse: str) -> int:
    cur = conn.execute("""
        SELECT qty FROM warehouse_stock WHERE fixture_id = ? AND warehouse = ?;
    """, (fixture_id, warehouse))
    row = cur.fetchone()
    if not row:
        _ensure_stock_row(conn, fixture_id, warehouse)
        return 0
    return int(row["qty"])

def get_stock(part_no: str, warehouse: str) -> int:
    _validate_part_no(part_no)
    with get_conn() as conn:
        cur = conn.execute("""
            SELECT ws.qty
            FROM fixtures f
            JOIN warehouse_stock ws ON ws.fixture_id = f.fixture_id
            WHERE f.part_no = ? AND ws.warehouse = ?;
        """, (part_no, warehouse))
        row = cur.fetchone()
        return int(row["qty"]) if row else 0

def insert_fixture(part_no: str, name: str, type_: str, spec: str = "", safe_stock: int = 0, created_by: str = "") -> Dict[str, Any]:
    _validate_part_no(part_no)
    type_ = (type_ or "").strip().lower()
    if type_ not in ("fixture", "consumable"):
        raise ValueError("type 必須為 'fixture' 或 'consumable'。")
    with get_conn() as conn:
        cur = conn.execute("SELECT fixture_id FROM fixtures WHERE part_no = ?;", (part_no,))
        row = cur.fetchone()
        if row:
            fixture_id = int(row["fixture_id"])
            conn.execute("""
                UPDATE fixtures
                   SET name = COALESCE(NULLIF(?,''), name),
                       type = ?,
                       spec = COALESCE(NULLIF(?,''), spec),
                       safe_stock = ?
                 WHERE fixture_id = ?;
            """, (name, type_, spec, max(0, int(safe_stock)), fixture_id))
        else:
            conn.execute("""
                INSERT INTO fixtures (part_no, name, type, spec, safe_stock)
                VALUES (?, ?, ?, ?, ?);
            """, (part_no, name, type_, spec, max(0, int(safe_stock))))
            fixture_id = conn.execute("SELECT last_insert_rowid() AS id;").fetchone()["id"]
        for wh in CORE_WAREHOUSES:
            _ensure_stock_row(conn, fixture_id, wh)
        return {"ok": True, "fixture_id": fixture_id, "part_no": part_no}

def add_stock(part_no: str, qty: int, user: str, remark: str = "") -> Dict[str, Any]:
    _validate_part_no(part_no)
    qty = int(qty)
    if qty <= 0:
        raise ValueError("入庫數量必須為正整數。")
    with get_conn() as conn:
        frow = conn.execute(
            "SELECT fixture_id FROM fixtures WHERE part_no = ?;",
            (part_no,)
        ).fetchone()
        if not frow:
            raise ValueError("料號不存在，請先建立料號。")
        fixture_id = int(frow["fixture_id"])

        _ensure_stock_row(conn, fixture_id, "治具室")
        conn.execute("""
            UPDATE warehouse_stock
               SET qty = qty + ?, updated_at = datetime('now','localtime')
             WHERE fixture_id = ? AND warehouse = '治具室';
        """, (qty, fixture_id))
        conn.execute("""
            INSERT INTO transfer_logs (fixture_id, action, from_wh, to_wh, qty, user, remark)
            VALUES (?, 'in', NULL, '治具室', ?, ?, ?);
        """, (fixture_id, qty, user, remark))
        return {"ok": True, "fixture_id": fixture_id, "to": "治具室", "delta": qty}
    
def transfer_stock(part_no: str, from_wh: str, to_wh: str, qty: int, user: str, remark: str = "") -> Dict[str, Any]:
    _validate_part_no(part_no)
    qty = int(qty)
    from_wh = (from_wh or "").strip()
    to_wh = (to_wh or "").strip()
    if qty <= 0:
        raise ValueError("轉倉數量必須為正整數。")
    if from_wh == to_wh:
        raise ValueError("來源倉與目的倉不可相同。")
    with get_conn() as conn:
        frow = conn.execute("SELECT fixture_id FROM fixtures WHERE part_no = ?;", (part_no,)).fetchone()
        if not frow:
            raise ValueError("料號不存在，請先建立料號。")
        fixture_id = int(frow["fixture_id"])
        src_qty = _get_stock_for_update(conn, fixture_id, from_wh)
        if src_qty < qty:
            raise ValueError(f"來源倉「{from_wh}」庫存不足（現有 {src_qty}，欲轉 {qty}）。")
        _ensure_stock_row(conn, fixture_id, to_wh)
        conn.execute("""
            UPDATE warehouse_stock
               SET qty = qty - ?, updated_at = datetime('now','localtime')
             WHERE fixture_id = ? AND warehouse = ?;
        """, (qty, fixture_id, from_wh))
        conn.execute("""
            UPDATE warehouse_stock
               SET qty = qty + ?, updated_at = datetime('now','localtime')
             WHERE fixture_id = ? AND warehouse = ?;
        """, (qty, fixture_id, to_wh))
        conn.execute("""
            INSERT INTO transfer_logs (fixture_id, action, from_wh, to_wh, qty, user, remark)
            VALUES (?, 'transfer', ?, ?, ?, ?, ?);
        """, (fixture_id, from_wh, to_wh, qty, user, remark))
        return {"ok": True, "fixture_id": fixture_id, "from": from_wh, "to": to_wh, "delta": qty}


def consume_stock(part_no: str, qty: int, user: str, line: str, purpose: str, from_wh: str = "上齊", move_to_consumed: bool = True) -> Dict[str, Any]:
    _validate_part_no(part_no)
    qty = int(qty)
    if qty <= 0:
        raise ValueError("消耗數量必須為正整數。")
    with get_conn() as conn:
        frow = conn.execute("SELECT fixture_id, type FROM fixtures WHERE part_no = ?;", (part_no,)).fetchone()
        if not frow:
            raise ValueError("料號不存在，請先建立料號。")
        fixture_id = int(frow["fixture_id"])
        src_qty = _get_stock_for_update(conn, fixture_id, from_wh)
        if src_qty < qty:
            raise ValueError(f"來源倉「{from_wh}」庫存不足（現有 {src_qty}，欲消耗 {qty}）。")
        conn.execute("""
            UPDATE warehouse_stock
               SET qty = qty - ?, updated_at = datetime('now','localtime')
             WHERE fixture_id = ? AND warehouse = ?;
        """, (qty, fixture_id, from_wh))
        conn.execute("""
            INSERT INTO consumption_logs (fixture_id, qty, user, line, purpose)
            VALUES (?, ?, ?, ?, ?);
        """, (fixture_id, qty, user, line, purpose))
        if move_to_consumed:
            _ensure_stock_row(conn, fixture_id, CONSUMED_WAREHOUSE)
            conn.execute("""
                UPDATE warehouse_stock
                   SET qty = qty + ?, updated_at = datetime('now','localtime')
                 WHERE fixture_id = ? AND warehouse = ?;
            """, (qty, fixture_id, CONSUMED_WAREHOUSE))
            conn.execute("""
                INSERT INTO transfer_logs (fixture_id, action, from_wh, to_wh, qty, user, remark)
                VALUES (?, 'transfer', ?, ?, ?, ?, 'consumption-auto');
            """, (fixture_id, from_wh, CONSUMED_WAREHOUSE, qty, user))
        return {"ok": True, "fixture_id": fixture_id, "from": from_wh, "to": CONSUMED_WAREHOUSE if move_to_consumed else None, "consumed": qty, "user": user, "line": line, "purpose": purpose}

def get_overview_by_warehouse(part_no: Optional[str] = None) -> Tuple[list[Dict[str, Any]], list[str]]:
    with get_conn() as conn:
        whs = CORE_WAREHOUSES + [CONSUMED_WAREHOUSE]
        wh_selects = ", ".join([
            f"COALESCE(MAX(CASE WHEN ws.warehouse='{wh}' THEN ws.qty END),0) AS '{wh}'"
            for wh in whs
        ])
        if part_no:
            _validate_part_no(part_no)
            cur = conn.execute(f"""
                SELECT f.part_no, f.name, f.type, {wh_selects}
                FROM fixtures f
                LEFT JOIN warehouse_stock ws ON ws.fixture_id = f.fixture_id
                WHERE f.part_no = ?
                GROUP BY f.fixture_id
                ORDER BY f.part_no;
            """, (part_no,))
        else:
            cur = conn.execute(f"""
                SELECT f.part_no, f.name, f.type, {wh_selects}
                FROM fixtures f
                LEFT JOIN warehouse_stock ws ON ws.fixture_id = f.fixture_id
                GROUP BY f.fixture_id
                ORDER BY f.part_no;
            """)
        rows = [dict(r) for r in cur.fetchall()]
        return rows, whs