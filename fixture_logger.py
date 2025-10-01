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
                fixture_log_id TEXT PRIMARY KEY,
                fixture_log_part_no TEXT,
                fixture_log_action TEXT,
                fixture_log_qty INTEGER,
                fixture_log_from_wh TEXT,
                fixture_log_to_wh TEXT,
                fixture_log_user TEXT,
                fixture_log_timestamp TEXT
            )
        """)
        conn.commit()

def generate_sheet_id(action_code: str) -> str:
    prefix, _ = FIXTURE_ACTION_MAP.get(action_code, ("Z", action_code))
    ym = datetime.now().strftime("%y%m")
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("SELECT COUNT(*) FROM fixture_logs WHERE fixture_log_id LIKE ?", (f"{prefix}{ym}%",))
        count = cur.fetchone()[0] or 0
    return f"{prefix}{ym}{count+1:04d}"

def log_fixture_activity(part_no, action, change_qty=0, from_wh="", to_wh="", user=""):
    prefix, action_name = FIXTURE_ACTION_MAP.get(action, ("Z", action))
    sheet_id = generate_sheet_id(action)
    ts = datetime.now().strftime("%Y%m%dT%H%M%S")
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute(
            """
            INSERT INTO fixture_logs (fixture_log_id, fixture_log_part_no, fixture_log_action,
                                      fixture_log_qty, fixture_log_from_wh, fixture_log_to_wh,
                                      fixture_log_user, fixture_log_timestamp)
            VALUES (?,?,?,?,?,?,?,?)
            """,
            (sheet_id, part_no, action_name, change_qty, from_wh, to_wh, user, ts),
        )
        conn.commit()

def get_fixture_logs(part_no=None, user=None, action=None, limit=200):
    query = """
        SELECT l.fixture_log_id,
               l.fixture_log_part_no,
               f.part_name,
               f.part_spec,
               l.fixture_log_action,
               l.fixture_log_qty,
               l.fixture_log_from_wh,
               l.fixture_log_to_wh,
               l.fixture_log_user,
               l.fixture_log_timestamp
        FROM fixture_logs l
        LEFT JOIN fixtures f ON l.fixture_log_part_no = f.part_no
        WHERE 1=1
    """
    params = []

    if part_no:
        query += " AND l.fixture_log_part_no LIKE ?"
        params.append(f"%{part_no}%")
    if user:
        query += " AND l.fixture_log_user LIKE ?"
        params.append(f"%{user}%")
    if action:
        query += " AND l.fixture_log_action=?"
        params.append(action)

    query += " ORDER BY l.fixture_log_timestamp DESC LIMIT ?"
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
            cur.execute(f"DELETE FROM fixture_logs WHERE fixture_log_id IN ({qmarks})", ids)
        else:
            cur.execute("DELETE FROM fixture_logs")