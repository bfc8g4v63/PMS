#$ export_readme.py
#% 匯出 changelog 到 README.md

import os
import sys
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parents[1]
sys.path.append(str(ROOT_DIR))

from config import apply_db_path
from db_helper import get_conn

apply_db_path()

OUTPUT_FILE = ROOT_DIR / "README.md"

PROJECT_NAME = "PMS"
FEATURES = [
    "SOP資訊 (已完成_產品資料列綁定SOP、計數、快捷開啟、停用/啟用功能)",
    "SOP生成 (已完成_SOP生成、SOP套用(綁)、資料分類)",
    "治具管理 (已完成_建立治具、刪除治具、修改資料、入庫、執行調撥、生成儲位、匯出Excel；消耗功能待驗證)",
    "治具紀錄 (已完成_出入庫紀錄_治具料號、治具品名、治具規格、動作、異動數量、來源倉、目的倉、治具操作人、時間)",
    "治具BOM (設計中_父子料號關聯)",
    "治具申請 (待規劃_電子簽核、申請單、審核)",
    "治具損耗 (待規劃_POWERBI_報表)",
    "異常平台 (待規劃_Trobleshoot Platform)",
    "帳號管理 (已完成_帳號、角色、新增、刪除、啟用、上傳SOP、可見記錄、刪除記錄、可見SOP、帳號管理)",
    "SOP紀錄 (已完成_SOP建立人、動作、檔案名稱、時間)",
    "改版歷程 (已完成_版本、日期、內容記錄)"
]

def export_readme():
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("SELECT version, date, content FROM change_log ORDER BY date DESC")
        rows = cur.fetchall()

    lines = []
    lines.append(f"# {PROJECT_NAME}\n")
    lines.append("## 功能簡介\n")
    for f in FEATURES:
        lines.append(f"* {f}")
    lines.append("\n---\n")
    lines.append("## 更新紀錄\n")

    for version, date, content in rows:
        parts = content.strip().splitlines()
        if parts:
            first = parts[0]
            lines.append(f"### {version} — {first}")
            for more in parts[1:]:
                if more.strip():
                    lines.append(f"* {more.strip()}")
            lines.append("")

    OUTPUT_FILE.write_text("\n".join(lines), encoding="utf-8")
    print(f"已更新 {OUTPUT_FILE}")

if __name__ == "__main__":
    export_readme()