#$ shutil.copy.py
#% 從網路位置複製 SQLite 資料庫到本地端

import shutil
from pathlib import Path
import sys

ROOT_DIR = Path(__file__).resolve().parents[1]
sys.path.append(str(ROOT_DIR))

from config import apply_db_path
from db_helper import DB_PATH

apply_db_path()

shutil.copy(DB_PATH, r"C:\Temp\PMS_local.db")
print(f"[✓] 已複製 {DB_PATH} 到 C:\\Temp\\PMS_local.db")