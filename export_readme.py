#$ export_readme.py
#% 匯出 changelog 到 README.md

import sqlite3
from pathlib import Path

DB_PATH = r"\\192.120.100.177\工程部\生產管理\生產資訊平台\PMS.db"
OUTPUT_FILE = Path(__file__).parent / "README.md"

def export_readme():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("SELECT version, date, content FROM changelog ORDER BY date DESC")
    rows = cur.fetchall()
    conn.close()

    lines = []
    lines.append("# PMS 系統\n")
    lines.append("## 改版歷程\n")

    for version, date, content in rows:
        lines.append(f"### {version} - {date}")
        lines.append(f"{content}\n")

    OUTPUT_FILE.write_text("\n".join(lines), encoding="utf-8")
    print(f"已更新 {OUTPUT_FILE}")

if __name__ == "__main__":
    export_readme()
