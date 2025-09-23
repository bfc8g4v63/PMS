#$ shutil.copy.py
#% 從網路位置複製 SQLite 資料庫到本地端

import shutil
shutil.copy(r"\\192.120.100.177\工程部\生產管理\生產資訊平台\PMS.db", r"C:\Temp\PMS_local.db")