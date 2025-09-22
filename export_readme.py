#$ export_readme.py
#% 匯出 changelog 到 README.md

# export_readme.py
# 匯出 changelog 到 README.md (符合指定格式)

import sqlite3
from pathlib import Path

DB_PATH = r"\\192.120.100.177\工程部\生產管理\生產資訊平台\PMS.db"
OUTPUT_FILE = Path(__file__).parent / "README.md"

PROJECT_NAME = "PMS"
FEATURES = [
    "SOP資訊 (已完成_產品資料列綁定SOP、計數、快捷開啟、停用/啟用功能)",
    "SOP生成 (已完成_SOP生成、SOP套用(綁)、資料分類)",
    "治具管理 (已完成_建立、刪除、修改、入庫、調撥、生成儲位、匯出Excel；消耗待規劃)",
    "治具紀錄 (設計中_出入庫紀錄)",
    "治具BOM (設計中_父子料號關聯)",
    "治具申請 (待規劃_電子簽核、申請單、審核)",
    "治具消耗 (待規劃_POWERBI 報表)",
    "異常平台 (待規劃_Trobleshoot Platform)",
    "帳號管理 (帳號、角色、新增、刪除、啟用、上傳SOP、可見記錄、刪除記錄、可見SOP、帳號管理)",
    "操作紀錄 (已完成_使用者、動作、檔案名稱、時間)",
    "改版歷程 (已完成_版本、日期、內容記錄)",
]

def export_readme():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("SELECT version, date, content FROM changelog ORDER BY date DESC")
    rows = cur.fetchall()
    conn.close()

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
