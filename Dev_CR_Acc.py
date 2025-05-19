import sqlite3
import hashlib
import os

DB_PATH = os.path.join(os.path.dirname(__file__), "PMS.db")

DEV_ADMIN = {
    "username": "Nelson",
    "password": "8463",
    "role": "admin",
    "specialty": "",
    "permissions": {
        "can_add": 1,
        "can_delete": 1,
        "active": 1,
        "can_view_logs": 1,
        "can_delete_logs": 1,
        "can_upload_sop": 1,
        "can_view_issues": 1,
        "can_manage_users": 1
    }
}

def hash_password(pw):
    return hashlib.sha256(pw.encode("utf-8")).hexdigest()

def create_admin():
    if not os.path.exists(DB_PATH):
        print("❌ 資料庫不存在，請先執行主程式建立 DB。")
        return

    with sqlite3.connect(DB_PATH) as conn:
        cursor = conn.cursor()
        cursor.execute("PRAGMA journal_mode=WAL;")

        cursor.execute("SELECT username FROM users WHERE username = ?", (DEV_ADMIN["username"],))
        if cursor.fetchone():
            print(f"⚠️ 帳號 '{DEV_ADMIN['username']}' 已存在，略過新增。")
            return

        fields = ", ".join(["username", "password", "role", "specialty"] + list(DEV_ADMIN["permissions"].keys()))
        placeholders = ", ".join(["?"] * (4 + len(DEV_ADMIN["permissions"])))
        values = (
            DEV_ADMIN["username"],
            hash_password(DEV_ADMIN["password"]),
            DEV_ADMIN["role"],
            DEV_ADMIN["specialty"],
            *DEV_ADMIN["permissions"].values()
        )

        cursor.execute(f"INSERT INTO users ({fields}) VALUES ({placeholders})", values)
        conn.commit()
        print(f"✅ 已建立開發者帳號 '{DEV_ADMIN['username']}'")

if __name__ == "__main__":
    create_admin()
