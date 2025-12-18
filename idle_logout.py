#$ idle_logout.py
#% 閒置自動登出管理

import tkinter as tk
from tkinter import messagebox


def attach_idle_logout(root: tk.Tk, enabled: bool, idle_timeout_seconds: int, logout_callback):
    if not enabled:
        return

    try:
        timeout_sec = int(idle_timeout_seconds or 0)
    except Exception:
        timeout_sec = 0

    if timeout_sec <= 0:
        return

    warning_sec = max(timeout_sec - 60, 1)
    warning_ms = warning_sec * 1000
    idle_ms = timeout_sec * 1000

    if not hasattr(root, "_idle_logout_state"):
        root._idle_logout_state = {
            "warning_after_id": None,
            "idle_after_id": None,
            "warning_shown": False,
            "bound": False,
        }

    state = root._idle_logout_state

    def cancel_timers():
        try:
            if state.get("warning_after_id") is not None:
                root.after_cancel(state["warning_after_id"])
        except Exception:
            pass
        try:
            if state.get("idle_after_id") is not None:
                root.after_cancel(state["idle_after_id"])
        except Exception:
            pass
        state["warning_after_id"] = None
        state["idle_after_id"] = None

    def on_idle_warning():
        if state.get("warning_shown"):
            return
        state["warning_shown"] = True
        try:
            messagebox.showwarning("閒置警告", f"您已閒置 {warning_sec} 秒，若再 60 秒未操作將自動登出。")
        except Exception:
            pass

    def on_idle_timeout():
        try:
            toplevel = tk.Toplevel(root)
            toplevel.title("自動登出")
            tk.Label(toplevel, text=f"您已閒置超過 {timeout_sec // 60} 分鐘，系統將自動登出").pack(padx=20, pady=20)
            toplevel.after(2000, lambda: logout_callback())
        except Exception:
            try:
                logout_callback()
            except Exception:
                pass

    def reset_idle_timer(event=None):
        cancel_timers()
        state["warning_shown"] = False
        try:
            state["warning_after_id"] = root.after(warning_ms, on_idle_warning)
            state["idle_after_id"] = root.after(idle_ms, on_idle_timeout)
        except Exception:
            pass

    if not state.get("bound"):
        for event_type in ("<Motion>", "<Key>", "<Button>"):
            try:
                root.bind_all(event_type, reset_idle_timer)
            except Exception:
                pass
        state["bound"] = True

    reset_idle_timer()