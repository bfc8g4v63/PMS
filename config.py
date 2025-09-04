# config.py
import os

USE_LOCAL_DB = True  # True: 本地測試, False: 寫入 Main Server
ENABLE_AUTO_LOGOUT = False   # True=啟用閒置登出, False=停用
IDLE_TIMEOUT = 180
VERBOSE_SCHEMA_CHECK = False  # True=啟用資料庫欄位檢查, False=停用

BASE_LOCAL_PATH = r"C:\Nelson\Dev\GitHub\PMS"
BASE_NETWORK_PATH = r"\\192.120.100.177\工程部\生產管理\生產資訊平台"

LOCAL_DB_PATH = os.path.join(BASE_LOCAL_PATH, "PMS.db")
Z_DRIVE_DB    = r"Z:\PMS.db"
UNC_DB        = os.path.join(BASE_NETWORK_PATH, "PMS.db")

if USE_LOCAL_DB:
    DB_NAME = LOCAL_DB_PATH
else:
    DB_NAME = Z_DRIVE_DB if os.path.exists(Z_DRIVE_DB) else UNC_DB

ORIGINAL_DB = Z_DRIVE_DB if os.path.exists(Z_DRIVE_DB) else UNC_DB