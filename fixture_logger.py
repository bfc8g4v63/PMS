#$ fixture_logger.py
#% 治具專用紀錄模組

from datetime import datetime
import sqlite3
import time
from db_helper import conn_ctx, tx

FIXTURE_ACTION_MAP = {
    "insert_fixture": ("C", "建立治具"),
    "delete_fixture": ("D", "刪除治具"),
    "update_fixture": ("U", "修改資料"),
    "add_stock": ("I", "入庫"),
    "transfer_stock": ("T", "調撥"),
    "adjust_stock": ("A", "調帳")
}

def ensure_fixture_log_schema():
    with conn_ctx() as conn:
        cur = conn.cursor()
        cur.execute(
            "CREATE TABLE IF NOT EXISTS fixture_logs ("
            "fixture_log_id TEXT PRIMARY KEY,"
            "fixture_log_part_no TEXT,"
            "fixture_log_action TEXT,"
            "fixture_log_qty INTEGER,"
            "fixture_log_from_wh TEXT,"
            "fixture_log_to_wh TEXT,"
            "fixture_log_user TEXT,"
            "fixture_log_reason TEXT,"
            "fixture_log_timestamp TEXT,"
            "fixture_log_deleted INTEGER DEFAULT 0,"
            "fixture_log_deleted_by TEXT,"
            "fixture_log_deleted_at TEXT"
            ")"
        )
        cur.execute("PRAGMA table_info(fixture_logs)")
        existing_cols = [r[1] for r in cur.fetchall()]

        if "fixture_log_deleted" not in existing_cols:
            cur.execute("ALTER TABLE fixture_logs ADD COLUMN fixture_log_deleted INTEGER DEFAULT 0")
        if "fixture_log_deleted_by" not in existing_cols:
            cur.execute("ALTER TABLE fixture_logs ADD COLUMN fixture_log_deleted_by TEXT")
        if "fixture_log_deleted_at" not in existing_cols:
            cur.execute("ALTER TABLE fixture_logs ADD COLUMN fixture_log_deleted_at TEXT")
        if "fixture_log_reason" not in existing_cols:
            cur.execute("ALTER TABLE fixture_logs ADD COLUMN fixture_log_reason TEXT")

        cur.execute("CREATE INDEX IF NOT EXISTS idx_fixture_logs_ts ON fixture_logs(fixture_log_timestamp)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_fixture_logs_part_no ON fixture_logs(fixture_log_part_no)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_fixture_logs_user ON fixture_logs(fixture_log_user)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_fixture_logs_action ON fixture_logs(fixture_log_action)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_fixture_logs_deleted ON fixture_logs(fixture_log_deleted)")

        conn.commit()

def generate_sheet_id(action_code: str, conn=None) -> str:
    prefix, _ = FIXTURE_ACTION_MAP.get(action_code, ("Z", action_code))
    ym = datetime.now().strftime("%y%m")
    like_prefix = f"{prefix}{ym}"

    if conn is None:
        with conn_ctx() as c:
            return generate_sheet_id(action_code, conn=c)

    cur = conn.cursor()
    cur.execute(
        "SELECT MAX(CAST(SUBSTR(fixture_log_id, 6, 4) AS INTEGER)) "
        "FROM fixture_logs "
        "WHERE fixture_log_id LIKE ?",
        (f"{like_prefix}%",),
    )
    row = cur.fetchone()
    max_seq = row[0] if row else None
    next_seq = (max_seq or 0) + 1
    return f"{like_prefix}{next_seq:04d}"

def log_fixture_activity(part_no, action, change_qty=0, from_wh="", to_wh="", user="", reason="", raise_on_fail: bool = False):
    part_no = "" if part_no is None else str(part_no).strip()
    from_wh = "" if from_wh is None else str(from_wh).strip()
    to_wh = "" if to_wh is None else str(to_wh).strip()
    user = "" if user is None else str(user).strip()
    reason = "" if reason is None else str(reason).strip()

    try:
        qty_val = int(change_qty or 0)
    except Exception:
        qty_val = 0

    _, action_name = FIXTURE_ACTION_MAP.get(action, ("Z", action))
    ts = datetime.now().strftime("%Y%m%dT%H%M%S")

    if action == "adjust_stock":
        if not reason:
            raise ValueError("調帳必須填寫原因")
    else:
        reason = ""

    last_err = None
    for _ in range(6):
        try:
            with tx() as conn:
                sheet_id = generate_sheet_id(action, conn=conn)
                cur = conn.cursor()
                cur.execute(
                    "INSERT INTO fixture_logs ("
                    "fixture_log_id, fixture_log_part_no, fixture_log_action, "
                    "fixture_log_qty, fixture_log_from_wh, fixture_log_to_wh, "
                    "fixture_log_user, fixture_log_reason, fixture_log_timestamp, "
                    "fixture_log_deleted, fixture_log_deleted_by, fixture_log_deleted_at"
                    ") VALUES (?,?,?,?,?,?,?,?,?,0,NULL,NULL)",
                    (sheet_id, part_no, action_name, qty_val, from_wh, to_wh, user, reason, ts),
                )
            return sheet_id
        except sqlite3.IntegrityError as e:
            last_err = e
            time.sleep(0.02)
        except Exception as e:
            last_err = e
            time.sleep(0.05)

    if raise_on_fail and last_err:
        raise last_err
    return None

def _normalize_id_list(ids):
    if not ids:
        return []
    out = []
    for x in ids:
        if x is None:
            continue
        s = str(x).strip()
        if s:
            out.append(s)
    return out

def _count_target_fixture_logs(conn, ids=None, include_deleted=False):
    ids = _normalize_id_list(ids)
    cur = conn.cursor()

    if ids:
        qmarks = ",".join("?" for _ in ids)
        if include_deleted:
            cur.execute(
                f"SELECT COUNT(1) FROM fixture_logs WHERE fixture_log_id IN ({qmarks})",
                ids,
            )
        else:
            cur.execute(
                f"SELECT COUNT(1) FROM fixture_logs WHERE fixture_log_id IN ({qmarks}) AND COALESCE(fixture_log_deleted,0)=0",
                ids,
            )
        row = cur.fetchone()
        return int(row[0] or 0)

    if include_deleted:
        cur.execute("SELECT COUNT(1) FROM fixture_logs")
    else:
        cur.execute("SELECT COUNT(1) FROM fixture_logs WHERE COALESCE(fixture_log_deleted,0)=0")
    row = cur.fetchone()
    return int(row[0] or 0)

def get_fixture_logs(part_no=None, user=None, action=None, start_ts=None, end_ts=None, limit=200, include_deleted=False):
    query = (
        "SELECT l.fixture_log_id, "
        "l.fixture_log_part_no, "
        "f.part_name, "
        "f.part_spec, "
        "l.fixture_log_action, "
        "l.fixture_log_qty, "
        "l.fixture_log_from_wh, "
        "l.fixture_log_to_wh, "
        "l.fixture_log_user, "
        "COALESCE(l.fixture_log_reason, ''), "
        "l.fixture_log_timestamp "
        "FROM fixture_logs l "
        "LEFT JOIN fixtures f ON l.fixture_log_part_no = f.part_no "
        "WHERE 1=1"
    )
    params = []

    if not include_deleted:
        query += " AND COALESCE(l.fixture_log_deleted, 0)=0"

    if part_no:
        query += " AND l.fixture_log_part_no LIKE ?"
        params.append(f"%{part_no}%")
    if user:
        query += " AND l.fixture_log_user LIKE ?"
        params.append(f"%{user}%")
    if action:
        query += " AND l.fixture_log_action=?"
        params.append(action)
    if start_ts:
        query += " AND l.fixture_log_timestamp >= ?"
        params.append(start_ts)
    if end_ts:
        query += " AND l.fixture_log_timestamp <= ?"
        params.append(end_ts)

    query += " ORDER BY l.fixture_log_timestamp DESC LIMIT ?"
    try:
        limit_val = int(limit or 200)
    except Exception:
        limit_val = 200
    params.append(limit_val)

    with conn_ctx() as conn:
        cur = conn.cursor()
        cur.execute(query, params)
        return cur.fetchall()

def delete_fixture_logs(ids=None, deleted_by="", note=""):
    deleted_by = "" if deleted_by is None else str(deleted_by).strip()
    ids = _normalize_id_list(ids)
    deleted_at = datetime.now().strftime("%Y%m%dT%H%M%S")

    with tx() as conn:
        cur = conn.cursor()

        if ids:
            qmarks = ",".join("?" for _ in ids)
            sql = (
                "UPDATE fixture_logs SET "
                "fixture_log_deleted=1, "
                "fixture_log_deleted_by=?, "
                "fixture_log_deleted_at=? "
                f"WHERE fixture_log_id IN ({qmarks}) AND COALESCE(fixture_log_deleted, 0)=0"
            )
            cur.execute(sql, [deleted_by, deleted_at, *ids])
            affected = cur.rowcount
        else:
            cur.execute(
                "UPDATE fixture_logs SET "
                "fixture_log_deleted=1, "
                "fixture_log_deleted_by=?, "
                "fixture_log_deleted_at=? "
                "WHERE COALESCE(fixture_log_deleted, 0)=0",
                (deleted_by, deleted_at),
            )
            affected = cur.rowcount

        conn.commit()
        return int(affected or 0)

def purge_fixture_logs(ids=None, purged_by="", note=""):
    ids = _normalize_id_list(ids)

    with tx() as conn:
        cur = conn.cursor()

        if ids:
            qmarks = ",".join("?" for _ in ids)
            cur.execute(f"DELETE FROM fixture_logs WHERE fixture_log_id IN ({qmarks})", ids)
            affected = cur.rowcount
        else:
            cur.execute("DELETE FROM fixture_logs")
            affected = cur.rowcount

        conn.commit()
        return int(affected or 0)