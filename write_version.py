from datetime import datetime

__version__ = "v1.0.2"
timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

with open("version.txt", "w", encoding="utf-8") as f:
    f.write(f"{__version__} - built at {timestamp}")