1.  新增統一資料庫連線函式
讓未來維護更簡潔，例如：

python
複製
編輯
def get_connection():
    conn = sqlite3.connect(DB_NAME, timeout=10)
    conn.execute("PRAGMA journal_mode=WAL;")
    return conn
用法改為：

python
複製
編輯
with get_connection() as conn:
    ...
2.  可選進階最佳化：自動檢查 .wal 檔案大小（預警或提示清理）
例如：

python
複製
編輯
def check_wal_size(threshold_mb=100):
    wal_path = DB_NAME + "-wal"
    if os.path.exists(wal_path):
        size_mb = os.path.getsize(wal_path) / (1024 * 1024)
        if size_mb > threshold_mb:
            print(f"WAL 檔案過大 ({size_mb:.1f} MB)，可考慮執行 checkpoint")
可以綁定定時檢查或管理員提示。

3.
4. 其他可優化建議
項目	評價 / 建議
current_user == username 比較 current_user 應為 dict，請改成 current_user.get("user")

檢查
1. timeout
2. conn.execute("PRAGMA journal_mode=WAL;")
3. conn.execute("PRAGMA wal_checkpoint(TRUNCATE);")
4. 其他BUG
