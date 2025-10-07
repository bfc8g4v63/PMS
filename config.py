#$ config.py
#% 系統設定檔，包含 DB_PATH（LOCAL/UNC/Z:）、AUTO_LOGOUT、IDLE_TIMEOUT、VERBOSE_SCHEMA_CHECK

import os
from db_helper import set_db_path   # 改為由 db_helper 提供 DB_PATH 控制

USE_LOCAL_DB = False                 # False: 使用本機資料庫；True: 使用 Z: 或 UNC 網路資料庫
ENABLE_AUTO_LOGOUT = False           # True: 啟用自動登出功能；False: 不啟用
VERBOSE_SCHEMA_CHECK = False         # True: 啟用詳細的資料表結構檢查與補齊訊息；False: 關閉
IDLE_TIMEOUT = 180                   # 自動登出閒置時間（秒），預設 3 分鐘


BASE_LOCAL_PATH = r"C:\Nelson\Dev\GitHub\PMS"
BASE_NETWORK_PATH = r"\\192.120.100.177\工程部\生產管理\生產資訊平台"

LOCAL_DB_PATH = os.path.join(BASE_LOCAL_PATH, "PMS.db")
Z_DRIVE_DB = r"Z:\PMS.db"
UNC_DB = os.path.join(BASE_NETWORK_PATH, "PMS.db")

if USE_LOCAL_DB:
    DB_NAME = Z_DRIVE_DB if os.path.exists(Z_DRIVE_DB) else UNC_DB
else:
    DB_NAME = LOCAL_DB_PATH

ORIGINAL_DB = Z_DRIVE_DB if os.path.exists(Z_DRIVE_DB) else UNC_DB

def apply_db_path():
    set_db_path(DB_NAME)