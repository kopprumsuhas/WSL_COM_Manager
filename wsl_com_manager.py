"""
WSL COM Manager v1.0.0
Modern White UI
Author: Suhas KR
Contact: kr.suhas1989@gmail.com
"""

import sys
import os
import ctypes
import subprocess
import json
import re
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime, timedelta

# ---------- Settings ----------
APP_NAME = "WSL COM Manager"
APP_VERSION = "1.0.0"
LOGFILE = "wsl_com_manager.log"
STATE_FILE = "devices_state.json"

# ---------- Admin relaunch (use same interpreter) ----------
def relaunch_as_admin():
    python_exe = sys.executable
    script = os.path.abspath(sys.argv[0])
    params = " ".join(sys.argv[1:])
    # Properly quoted command line so spaces in path are handled
    cmd = f"\"{script}\" {params}"
    ctypes.windll.shell32.ShellExecuteW(None, "runas", python_exe, cmd, os.getcwd(), 1)
    sys.exit(0)

def ensure_admin():
    try:
        is_admin = ctypes.windll.shell32.IsUserAnAdmin()
    except Exception:
        is_admin = False
    if not is_admin:
        relaunch_as_admin()

# Try to elevate immediately
ensure_admin()

# ---------- Utilities: logging, running commands ----------
def log_write(msg):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"{ts} - {msg}"
    try:
        with open(LOGFILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass

def run(cmd):
    log_write(f"CMD: {cmd}")
    try:
        p = subprocess.run(cmd, capture_output=True, text=True, shell=True)
        out = (p.stdout or "").strip()
        err = (p.stderr or "").strip()
        combined = "\n".join(x for x in (out, err) if x) or "(No output)"
        log_write("OUT: " + combined.replace("\n", " | "))
        return combined
    except Exception as e:
        log_write("RUN ERROR: " + str(e))
        return str(e)

# ---------- 7-day cleanup for logs and state ----------
def cleanup_logs(days=7):
    if not os.path.exists(LOGFILE):
        return
    cutoff = datetime.now() - timedelta(days=days)
    kept = []
    try:
        with open(LOGFILE, "r", encoding="utf-8") as f:
            for line in f:
                try:
                    ts_str = line.split(" - ")[0]
                    ts = datetime.strptime(ts_str, "%Y-%m-%d %H:%M:%S")
                    if ts >= cutoff:
                        kept.append(line)
                except Exception:
                    # keep unexpected lines
                    kept.append(line)
        with open(LOGFILE, "w", encoding="utf-8") as f:
            f.writelines(kept)
    except Exception:
        pass

def cleanup_state(days=7):
    if not os.path.exists(STATE_FILE):
        return
    try:
        data = json.load(open(STATE_FILE, "r", encoding="utf-8"))
    except Exception:
        return
    cutoff = datetime.now() - timedelta(days=days)
    new_data = {}
    for k, v in data.items():
        try:
            ts = datetime.fromisoformat(v.get("last_timestamp", "2000-01-01T00:00:00"))
            if ts >= cutoff:
                new_data[k] = v
        except Exception:
            new_data[k] = v
    try:
        json.dump(new_data, open(STATE_FILE, "w", encoding="utf-8"), indent=2)
    except Exception:
        pass

cleanup_logs()
cleanup_state()

# ---------- State handling ----------
def load_state():
    if not os.path.exists(STATE_FILE):
        return {}
    try:
        return json.load(open(STATE_FILE, "r", encoding="utf-8"))
    except Exception:
        return {}

def save_state(s):
    try:
        json.dump(s, open(STATE_FILE, "w", encoding="utf-8"), indent=2)
    except Exception:
        pass

state = load_state()

def make_key(vid, pid, serial):
    return f"{vid}:{pid}:{serial or 'NOSN'}"

# ---------- USB / serial helpers ----------
try:
    import serial.tools.list_ports as serial_list_ports
except Exception:
    serial_list_ports = None

def get_com_ports():
    if serial_list_ports is None:
        return []
    try:
        return list(serial_list_ports.comports())
    except Exception:
        return []

def extract_vid_pid_serial(port):
    vid = getattr(port, "vid", None)
    pid = getattr(port, "pid", None)
    sn = getattr(port, "serial_number", None)
    if vid is not None and pid is not None:
        return f"{vid:04X}", f"{pid:04X}", sn
    hw = getattr(port, "hwid", "") or ""
    m = re.search(r"VID[_:]?([0-9A-Fa-f]{4}).*PID[_:]?([0-9A-Fa-f]{4})", hw)
    if m:
        return m.group(1).upper(), m.group(2).upper(), sn
    return None, None, sn

def find_busids(vid, pid, serial_hint=None):
    out = run("usbipd list")
    pattern = re.compile(r"(\S+)\s+([0-9A-Fa-f]{4}):([0-9A-Fa-f]{4})(.*)")
    found = []
    for line in out.splitlines():
        m = pattern.search(line)
        if not m:
            continue
        busid, v, p, rest = m.group(1), m.group(2).upper(), m.group(3).upper(), m.group(4)
        if v == vid and p == pid:
            found.append((busid, rest))
    if serial_hint:
        for b, rest in found:
            if serial_hint and serial_hint in rest:
                return [b]
    return [b for b, _ in found]

def bind_attach(busid):
    return run(f"usbipd bind --busid {busid}") + "\n" + run(f"usbipd attach --wsl --busid {busid}")

def safe_detach_unbind(busid):
    return run(f"usbipd detach --busid {busid}") + "\n" + run(f"usbipd unbind --busid {busid}")

# ---------- GUI: Modern White theme ----------
root = tk.Tk()
root.title(f"{APP_NAME} v{APP_VERSION}")
root.geometry("640x560")
root.configure(bg="#ffffff")  # white background

# fonts and sizes for two-line centered buttons
BTN_FONT = ("Segoe UI", 12)
BTN_WIDTH = 12        # width in text units
BTN_HEIGHT = 3        # height in text units to accommodate 2 lines comfortably

# ---------- Debug toggles ----------
DEBUG_MODE = False
DEBUG_CLICK_COUNT = 0

def toggle_debug_click(event=None):
    global DEBUG_CLICK_COUNT, DEBUG_MODE
    DEBUG_CLICK_COUNT += 1
    if DEBUG_CLICK_COUNT < 10:
        return
    DEBUG_CLICK_COUNT = 0
    DEBUG_MODE = not DEBUG_MODE
    if DEBUG_MODE:
        debug_btn.pack(pady=8)
        root.title(f"{APP_NAME} v{APP_VERSION} — DEBUG MODE")
    else:
        debug_btn.pack_forget()
        root.title(f"{APP_NAME} v{APP_VERSION}")

def hotkey_debug(event=None):
    toggle_debug_click()

root.bind("<Button-1>", toggle_debug_click)
root.bind_all("<Control-Shift-D>", hotkey_debug)

# ---------- Layout ----------
top_frame = tk.Frame(root, bg="#ffffff")
top_frame.pack(padx=14, pady=10, fill="x")

lbl = tk.Label(top_frame, text="Select USB COM Port:", font=("Segoe UI", 11), bg="#ffffff")
lbl.pack(anchor="w")

combo = ttk.Combobox(top_frame, width=70)
combo.pack(pady=8)

# button area (icons are unicode two-line text)
btn_frame = tk.Frame(root, bg="#ffffff")
btn_frame.pack(pady=8)

# helper to centralize multi-line label config for a button
btn_opts = dict(width=BTN_WIDTH, height=BTN_HEIGHT, font=BTN_FONT, justify="center", anchor="center", bg="#f5f5f5")

attach_text = "🔌IN\nAttach"
detach_text = "🔌OUT\nDetach"
refresh_text = "🔎 COM\nRefresh"
debug_text = "De🪲\nLogs"

attach_btn = tk.Button(btn_frame, text=attach_text, **btn_opts, command=lambda: manual_attach())
attach_btn.grid(row=0, column=0, padx=14)

detach_btn = tk.Button(btn_frame, text=detach_text, **btn_opts, command=lambda: manual_detach())
detach_btn.grid(row=0, column=1, padx=14)

refresh_btn = tk.Button(btn_frame, text=refresh_text, **btn_opts, command=lambda: refresh_ports())
refresh_btn.grid(row=0, column=2, padx=14)

# debug button (hidden initially)
debug_btn = tk.Button(root, text=debug_text, **btn_opts, command=lambda: open_log_file())
# not packed until debug mode activated

# status box
status_box = tk.Text(root, width=82, height=12, wrap="word", state="disabled", bg="#fafafa")
status_box.pack(padx=14, pady=(10,12))

# help and footer
help_frame = tk.Frame(root, bg="#ffffff")
help_frame.pack(fill="x", padx=12, pady=(0,8))
help_btn = tk.Button(help_frame, text="❓ Help", command=lambda: show_help(), bg="#f5f5f5")
help_btn.pack(side="left")
footer = tk.Label(help_frame, text=f"© 2025 Suhas KR — Contact: kopprumsuhas@gmail.com", bg="#ffffff", font=("Segoe UI", 8))
footer.pack(side="right")

# ---------- Status helper ----------
def set_status(msg, success=True):
    status_box.config(state="normal")
    status_box.delete("1.0", tk.END)
    prefix = "✅ SUCCESS\n" if success else "❌ FAILED\n"
    status_box.insert("1.0", prefix + msg)
    status_box.config(state="disabled")

# ---------- Actions: refresh / attach / detach ----------
def refresh_ports():
    try:
        ports = get_com_ports()
        combo["values"] = [f"{p.device} - {p.description}" for p in ports]
        if ports:
            combo.set(combo["values"][0])
        set_status("Ports refreshed.")
    except Exception as e:
        set_status("Failed to list ports: " + str(e), False)

def manual_attach():
    sel = combo.get()
    if not sel:
        set_status("Select a COM port first.", False)
        return
    com = sel.split()[0]
    ports = {p.device: p for p in get_com_ports()}
    port = ports.get(com)
    if not port:
        set_status("COM port not available.", False)
        return
    vid, pid, sn = extract_vid_pid_serial(port)
    if not vid or not pid:
        set_status("Unable to read VID/PID for selected device.", False)
        return
    key = make_key(vid, pid, sn)
    saved = state.get(key)
    # cleanup previous busid if exists
    if saved and saved.get("last_busid"):
        try:
            safe_detach_unbind(saved["last_busid"])
        except Exception:
            pass
    busids = find_busids(vid, pid, sn)
    if not busids:
        set_status("Device not found in usbipd list.", False)
        return
    busid = busids[0]
    out = bind_attach(busid)
    state[key] = {
        "vid": vid,
        "pid": pid,
        "serial": sn,
        "last_busid": busid,
        "last_action": "attached",
        "last_timestamp": datetime.now().isoformat()
    }
    save_state(state)
    set_status(out)

def manual_detach():
    sel = combo.get()
    if not sel:
        set_status("Select a COM port first.", False)
        return
    com = sel.split()[0]
    ports = {p.device: p for p in get_com_ports()}
    port = ports.get(com)
    saved_busid = None
    if port:
        vid, pid, sn = extract_vid_pid_serial(port)
        key = make_key(vid, pid, sn)
        busids = find_busids(vid, pid, sn)
        if busids:
            saved_busid = busids[0]
        elif key in state:
            saved_busid = state[key].get("last_busid")
    else:
        # fallback: find any last saved busid
        for info in state.values():
            if info.get("last_busid"):
                saved_busid = info["last_busid"]
                break
    if not saved_busid:
        set_status("Nothing to detach.", False)
        return
    out = safe_detach_unbind(saved_busid)
    # clear state entries that used this busid
    for k, v in list(state.items()):
        if v.get("last_busid") == saved_busid:
            v["last_busid"] = None
            v["last_action"] = "detached"
            v["last_timestamp"] = datetime.now().isoformat()
    save_state(state)
    set_status(out)

# ---------- Help, logs ----------
def show_help():
    txt = (
        f"{APP_NAME} v{APP_VERSION}\n\n"
        "How to use:\n"
        "1) Plug USB serial device into Windows.\n"
        "2) Select the COM port from the dropdown.\n"
        "3) Click 'Attach' to bind & attach to WSL (usbipd required).\n"
        "4) Click 'Detach' to unbind & detach (admin required).\n"
        "5) Click 'Refresh' to refresh the COM list.\n\n"
        "Debug Mode: click in the window 10 times or press Ctrl+Shift+D.\n"
        "In Debug Mode the 'De🪲\\nLogs' button appears to open the log file.\n\n"
        "Support: kr.suhas1989@gmail.com"
    )
    messagebox.showinfo("Help - How to use", txt)

def open_log_file():
    try:
        if os.path.exists(LOGFILE):
            os.startfile(LOGFILE)
        else:
            set_status("Log file not found.", False)
    except Exception as e:
        set_status("Failed to open log file: " + str(e), False)

# ---------- Initialize UI content ----------
def initialize():
    try:
        refresh_ports()
    except Exception:
        pass
    set_status("Ready. Select a COM port and press Attach or Detach.")

# ---------- Start ----------
initialize()
root.mainloop()
