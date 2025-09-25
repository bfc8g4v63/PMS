#$ wal_shm_check.py
#% 檢查 SQLite 資料庫及其 WAL/SHM 檔案是否被其他程序佔用

import os
import psutil
import sys
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parents[1]
sys.path.append(str(ROOT_DIR))

from config import apply_db_path
from db_helper import DB_PATH

apply_db_path()

LOCK_FILES = [
    DB_PATH,
    f"{DB_PATH}-wal",
    f"{DB_PATH}-shm"
]

def check_file_in_use(path):
    for proc in psutil.process_iter(['pid', 'name', 'open_files']):
        try:
            flist = proc.info['open_files']
            if flist:
                for f in flist:
                    if path.lower() == f.path.lower():
                        print(f"[發現佔用] {f.path} 被 PID {proc.pid} ({proc.name()}) 使用中")
                        return True
        except Exception:
            continue
    return False

for f in LOCK_FILES:
    if not os.path.exists(f):
        print(f"[不存在] {f}")
        continue
    check_file_in_use(f)