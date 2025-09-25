#$ fixture_logger.py
#% 治具專用紀錄模組

from datetime import datetime
from db_helper import get_conn, tx

FIXTURE_ACTION_MAP = {
    "insert_fixture": ("C", "建立治具"),
    "delete_fixture": ("D", "刪除治具"),
    "update_fixture": ("U", "修改資料"),
    "add_stock": ("I", "入庫"),
    "transfer_stock": ("T", "調撥"),
    "consume_stock": ("X", "消耗")
}

def ensure_fixture_log_schema():
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("""
            CREATE TABLE IF NOT EXISTS fixture_logs (
                sheet_id TEXT PRIMARY KEY,
                part_no TEXT,
                action TEXT,
                change_qty INTEGER,
                from_wh TEXT,
                to_wh TEXT,
                user TEXT,
                timestamp TEXT
            )
        """)
        conn.commit()

def generate_sheet_id(action_code: str) -> str:
    prefix, _ = FIXTURE_ACTION_MAP.get(action_code, ("Z", action_code))
    ym = datetime.now().strftime("%y%m")  # 例如 2509
    base = f"{ym}"
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute(
            "SELECT sheet_id FROM fixture_logs WHERE sheet_id LIKE ? ORDER BY sheet_id DESC LIMIT 1",
            (f"%{base}%",)
        )
        row = cur.fetchone()
        if row:
            last_seq = int(row[0][-5:])
            next_seq = last_seq + 1
        else:
            next_seq = 1
    return f"{prefix}{ym}{next_seq:05d}"

def log_fixture_activity(part_no, action, change_qty=0, from_wh="", to_wh="", user=""):
    prefix, action_name = FIXTURE_ACTION_MAP.get(action, ("Z", action))
    sheet_id = generate_sheet_id(action)
    ts = datetime.now().strftime("%Y%m%dT%H%M%S")
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute(
            """
            INSERT INTO fixture_logs (sheet_id, part_no, action, change_qty, from_wh, to_wh, user, timestamp)
            VALUES (?,?,?,?,?,?,?,?)
            """,
            (sheet_id, part_no, action_name, change_qty, from_wh, to_wh, user, ts),
        )
        conn.commit()

def get_fixture_logs(part_no=None, user=None, action=None, limit=200):
    query = """
        SELECT l.sheet_id,
               l.part_no,
               f.part_name,
               f.part_spec,
               l.action,
               l.change_qty,
               l.from_wh,
               l.to_wh,
               l.user,
               l.timestamp
        FROM fixture_logs l
        LEFT JOIN fixtures f ON l.part_no = f.part_no
        WHERE 1=1
    """
    params = []

    if part_no:
        query += " AND l.part_no LIKE ?"
        params.append(f"%{part_no}%")
    if user:
        query += " AND l.user LIKE ?"
        params.append(f"%{user}%")
    if action:
        query += " AND l.action=?"
        params.append(action)

    query += " ORDER BY l.timestamp DESC LIMIT ?"
    params.append(limit)

    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute(query, params)
        return cur.fetchall()

def delete_fixture_logs(ids=None):
    with tx() as conn:
        cur = conn.cursor()
        if ids:
            qmarks = ",".join("?" for _ in ids)
            cur.execute(f"DELETE FROM fixture_logs WHERE sheet_id IN ({qmarks})", ids)
        else:
            cur.execute("DELETE FROM fixture_logs")