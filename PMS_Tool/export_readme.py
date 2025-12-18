#$ export_readme.py
#% 匯出 changelog 到 README.md（固定使用正式伺服器 .177 資料庫）

import os
import sys
import sqlite3
import re
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.append(str(ROOT_DIR))

DB_PATH = r"\\192.120.100.177\工程部\生產管理\生產資訊平台\PMS.db"
OUTPUT_FILE = ROOT_DIR / "README.md"

PROJECT_NAME = "PMS"

FEATURES = [
    "SOP資訊 (已完成_產品資料列綁定SOP、使用次數計數、快捷開啟、啟用/停用控管)",
    "SOP生成 (已完成_SOP生成、SOP套用(綁定產品)、資料分類管理)",
    "治具管理 (已完成_建立治具、刪除治具、修改治具資料、入庫、調撥/轉倉、自動生成與檢核儲位、Excel匯出(本幣單價、外幣單價、料件總價、各倉別總價、預估料件請購量、預估請購金額、預估請購總金額)、安全庫存告警、調帳(差額/盤點))",
    "治具紀錄 (已完成_出入庫與異動紀錄_治具料號、治具品名、治具規格、動作、異動數量、來源倉、目的倉、治具操作人、調帳原因、時間)",
    "治具BOM (開發中_父子料號關聯、可替代料件結構、右鍵選單新增父子料號關聯)",
    "治具申請 (待規劃_治具申請單流程、審核機制)",
    "治具損耗 (待規劃_損耗資料彙整、PowerBI報表)",
    "異常平台 (待規劃_Troubleshoot、異常回報與追蹤)",
    "帳號與權限管理 (已完成_帳號啟用/停用、角色管理、新增SOP列、刪除SOP列、上傳SOP、SOP紀錄可見、SOP紀錄刪除、SOP資訊可見、帳號管理、治具可見、治具編輯、治具調帳、治具紀錄可見、治具紀錄刪除)",
    "SOP紀錄 (已完成_SOP建立人、操作動作、檔案名稱、操作時間)",
    "改版歷程 (已完成_版本號、改版日期、改版內容紀錄)"
]


def _pick_changelog_table(conn):
    cur = conn.cursor()
    cur.execute("SELECT name FROM sqlite_master WHERE type='table'")
    names = {r[0] for r in cur.fetchall()}
    if "change_log" in names:
        return "change_log"
    if "changelog" in names:
        return "changelog"
    raise ValueError("找不到 change_log / changelog 表")


def _ensure_changelog_columns(conn, table_name: str):
    cur = conn.cursor()
    cur.execute(f"PRAGMA table_info({table_name})")
    cols = {r[1] for r in cur.fetchall()}
    required = {"version", "date", "content"}
    missing = required - cols
    if missing:
        raise ValueError(f"{table_name} 缺少欄位: {', '.join(sorted(missing))}")


def _parse_version(v: str):
    s = (v or "").strip()
    if s[:1] in ("v", "V"):
        s = s[1:]
    nums = re.findall(r"\d+", s)
    out = [int(x) for x in nums[:3]]
    while len(out) < 3:
        out.append(0)
    return tuple(out)


def _looks_like_md_list_line(s: str) -> bool:
    t = (s or "").lstrip()
    if not t:
        return False
    if t[:2] in ("* ", "- "):
        return True
    if re.match(r"^\d+\.\s+", t):
        return True
    return False


def export_readme():
    print(f"script={Path(__file__).resolve()}")
    print(f"root={ROOT_DIR}")
    print(f"output={OUTPUT_FILE}")

    try:
        with sqlite3.connect(DB_PATH, timeout=10) as conn:
            table = _pick_changelog_table(conn)
            _ensure_changelog_columns(conn, table)
            cur = conn.cursor()
            cur.execute(f"SELECT version, date, content FROM {table}")
            rows = cur.fetchall()
    except Exception as e:
        raise RuntimeError(f"讀取 changelog 失敗: {e}")

    rows = sorted(
        rows,
        key=lambda r: (_parse_version(r[0]), str(r[1] or ""), str(r[2] or "")),
        reverse=True
    )

    lines = []
    lines.append(f"# {PROJECT_NAME}")
    lines.append("")
    lines.append("## 功能簡介")
    lines.append("")
    for f in FEATURES:
        lines.append(f"* {f}")
    lines.append("")
    lines.append("---")
    lines.append("")
    lines.append("## 更新紀錄")
    lines.append("")

    for version, date, content in rows:
        parts = str(content or "").strip().splitlines()
        if not parts:
            continue
        first = parts[0].strip()
        lines.append(f"### {version} — {first}")
        for more in parts[1:]:
            t = more.rstrip()
            if t.strip():
                if _looks_like_md_list_line(t):
                    lines.append(t.strip())
                else:
                    lines.append(f"* {t.strip()}")
        lines.append("")

    OUTPUT_FILE.write_text("\n".join(lines).rstrip() + "\n", encoding="utf-8")
    print(f"已更新 {OUTPUT_FILE}（來源：{DB_PATH}）")


if __name__ == "__main__":
    export_readme()