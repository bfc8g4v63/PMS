#$ db_backup.py
#% 資料庫備份腳本：專門備份伺服器正式 PMS.db 至本地資料夾（含時間戳）

import shutil
import os
from datetime import datetime

SOURCE = r"\\192.120.100.177\工程部\生產管理\生產資訊平台\PMS.db"
TARGET_ROOT = r"C:\Nelson\Dev\GitHub\PMS_backup\db_backup"

now = datetime.now()
timestamp = now.strftime("%Y%m%d_%H%M")
target_folder = os.path.join(TARGET_ROOT, timestamp)
os.makedirs(target_folder, exist_ok=True)
target_file = os.path.join(target_folder, "PMS.db")

if not os.path.exists(SOURCE):
    print(f"[X] 無法存取伺服器資料庫：{SOURCE}")
    print("[提示] 請確認是否連線至公司內網或 177 共用路徑是否可用。")
else:
    try:
        shutil.copy2(SOURCE, target_file)
        print(f"[✓] 已成功備份伺服器資料庫至：{target_file}")
    except Exception as e:
        print(f"[X] 備份失敗：{e}")
