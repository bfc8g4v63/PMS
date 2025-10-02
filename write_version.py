#$ write_version.py
#% 版本控制腳本：同時輸出在地端與 .177 雲端版本檔案
import os
from datetime import datetime

__version__ = "v2.3.6"
timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

content = f"{__version__} - built at {timestamp}"

LOCAL_FILE = "version.txt"

CLOUD_FILE = r"\\192.120.100.177\工程部\生產管理\生產資訊平台\version.txt"

with open(LOCAL_FILE, "w", encoding="utf-8") as f:
    f.write(content)

try:
    os.makedirs(os.path.dirname(CLOUD_FILE), exist_ok=True)
    with open(CLOUD_FILE, "w", encoding="utf-8") as f:
        f.write(content)
    print(f"[✓] 已更新本地與雲端版本檔：{__version__}")
except Exception as e:
    print(f"[X] 雲端版本檔更新失敗: {e}")