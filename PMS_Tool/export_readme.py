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
    "SOP紀錄 (已完成_SOP建立人、操作動作、檔案名稱、操作時間)",
    "SOP套用 (已完成_SOP來源選取、批量套用、全選/勾選、套用清單UX優化)",
    "治具管理 (已完成_建立治具、刪除治具、修改治具資料、入庫、調撥/轉倉、自動生成與檢核儲位、各倉別獨立存量顯示、可用數量控管、Excel匯出(本幣單價、外幣單價、料件總價、各倉別總價、預估料件請購量、預估請購金額、預估請購總金額)、安全庫存告警、調帳(差額/盤點))",
    "治具紀錄 (已完成_出入庫、轉倉、調帳異動紀錄_治具料號、治具品名、治具規格、動作、異動數量、來源倉、目的倉、操作人、原因、時間)",
    "治具BOM (開發中_父子料號關聯、可替代料件結構、BOM數量設定、右鍵選單新增/刪除關聯)",
    "資料庫一致性檢查 (已完成_資料表結構自動補齊、欄位缺失檢查、版本差異防呆)",
    "操作防呆機制 (已完成_按鈕重複點擊防止、操作流程鎖定、避免重複寫入)",
    "帳號與權限管理 (已完成_帳號啟用/停用、角色管理、功能權限控管(SOP/治具/調帳/紀錄)、資料可見性限制)",
    "操作紀錄與稽核 (已完成_登入/操作行為紀錄、模組別記錄、時間戳與使用者追蹤)",
    "閒置登出機制 (已完成_閒置時間偵測、自動登出、防止帳號長時間占用)",
    "資料庫連線與穩定性處理 (已完成_SQLite WAL模式、busy_timeout、連線集中管理、避免DB lock)",
    "系統初始化與環境控管 (已完成_DB路徑套用、正式/本地環境切換、啟動時Schema初始化)",
    "改版歷程 (已完成_版本號、改版日期、改版內容紀錄)"
]

UPCOMING_FEATURES = [
    "治具申請與調撥流程 (規劃中_需求單位起單、編輯單身、送出申請、列印申請單、實體簽名、附件上傳)",
    "治具審核決策流程 (規劃中_虹堡集中審核、可於審核階段編輯單身品項與數量)",
    "庫存驗證與核發機制 (規劃中_審核時即時驗證來源倉庫存、避免超發、核可後才執行庫存異動)",
    "單據狀態控管 (規劃中_起單、待審核、核可、否決、逾時否決狀態流轉)",
    "申請單附件與簽名留存 (規劃中_申請單文件綁定、簽名附件保存、審核依據查詢)",
    "逾時防堵機制 (規劃中_申請單48小時未審自動否決、避免單據長期掛單)",
    "跨倉調撥一致流程 (規劃中_B to C 與 C to C 共用流程、統一由虹堡決策核發)"
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
    lines.append("## Upcoming Features")
    lines.append("")
    for f in UPCOMING_FEATURES:
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