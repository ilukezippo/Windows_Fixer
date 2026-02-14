import os
import sys
import json
import shutil
import ctypes
import subprocess
import threading
import queue
import tkinter as tk
from tkinter import ttk, messagebox
import webbrowser
from datetime import date

from io import BytesIO
from PIL import Image, ImageDraw  # pillow installed
import winsound

APP_ID = "WindowsFixer"
APP_VERSION = "v1.0.0"
BUILD_DATE = date.today().isoformat()  # shown in About

DONATE_PAGE = "https://buymeacoffee.com/ilukezippo"
GITHUB_PAGE = "https://github.com/ilukezippo/Windows_Fixer"

WIN_W = 1180
WIN_H = 900


# -------------------------
# resource helpers (PyInstaller friendly)
# -------------------------
def resource_path(relative_path: str) -> str:
    try:
        base_path = sys._MEIPASS  # PyInstaller temp folder
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

root = tk.Tk()

# Set window icon (works in Python + EXE)
try:
    root.iconbitmap(resource_path("icon.ico"))
except Exception as e:
    print("Icon load failed:", e)





def set_app_icon(root):
    ico = resource_path("icon.ico")  # your icon file
    if os.path.exists(ico):
        try:
            root.iconbitmap(ico)
            return ico
        except Exception:
            pass
    return None


def apply_icon_to_tlv(tlv, icon):
    if icon:
        try:
            tlv.iconbitmap(icon)
        except Exception:
            pass


def load_flag_image():
    png = resource_path("kuwait.png")  # optional
    if os.path.exists(png):
        try:
            return tk.PhotoImage(file=png)
        except Exception:
            pass
    return None


def make_donate_image(w=160, h=44):
    r = h // 2
    top = (255, 187, 71)
    mid = (247, 162, 28)
    bot = (225, 140, 22)

    im = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    dr = ImageDraw.Draw(im)

    for y in range(h):
        if y < h * 0.6:
            t = y / (h * 0.6)
            c = tuple(int(top[i] * (1 - t) + mid[i] * t) for i in range(3)) + (255,)
        else:
            t = (y - h * 0.6) / (h * 0.4)
            c = tuple(int(mid[i] * (1 - t) + bot[i] * t) for i in range(3)) + (255,)
        dr.line([(0, y), (w, y)], fill=c)

    mask = Image.new("L", (w, h), 0)
    ImageDraw.Draw(mask).rounded_rectangle([0, 0, w - 1, h - 1], radius=r, fill=255)
    im.putalpha(mask)

    hl = Image.new("RGBA", (w, h), (255, 255, 255, 0))
    ImageDraw.Draw(hl).rounded_rectangle([2, 2, w - 3, h // 2], radius=r - 2, fill=(255, 255, 255, 70))
    im = Image.alpha_composite(im, hl)

    ImageDraw.Draw(im).rounded_rectangle(
        [0.5, 0.5, w - 1.5, h - 1.5], radius=r, outline=(200, 120, 20, 255), width=2
    )

    bio = BytesIO()
    im.save(bio, format="PNG")
    bio.seek(0)
    return tk.PhotoImage(data=bio.read())


def play_success_sound():
    wav = resource_path("Success.wav")
    if os.path.exists(wav):
        try:
            winsound.PlaySound(wav, winsound.SND_FILENAME | winsound.SND_ASYNC)
        except Exception:
            pass


# -------------------------
# Settings
# -------------------------
def _settings_path():
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    folder = os.path.join(base, APP_ID)
    os.makedirs(folder, exist_ok=True)
    return os.path.join(folder, "settings.json")


def load_settings():
    path = _settings_path()
    try:
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return {"always_admin": False, "language": "en"}


def save_settings(data: dict):
    path = _settings_path()
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


# -------------------------
# Admin
# -------------------------
def is_admin() -> bool:
    try:
        return ctypes.windll.shell32.IsUserAnAdmin() != 0
    except Exception:
        return False


def relaunch_as_admin():
    params = " ".join([f'"{a}"' for a in sys.argv])
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, params, None, 1)
    sys.exit(0)


# -------------------------
# Drives
# -------------------------
def list_drives():
    drives = []
    bitmask = ctypes.windll.kernel32.GetLogicalDrives()
    for i in range(26):
        if bitmask & (1 << i):
            letter = chr(ord("A") + i)
            path = f"{letter}:\\"
            if os.path.exists(path):
                drives.append(f"{letter}:")
    return drives or ["C:"]


# -------------------------
# Cleanup helpers
# -------------------------
def safe_rmtree(path: str, log_cb):
    try:
        if os.path.isdir(path) and not os.path.islink(path):
            shutil.rmtree(path, ignore_errors=True)
        else:
            try:
                os.remove(path)
            except Exception:
                pass
    except Exception as e:
        log_cb(f"[WARN] Could not delete {path}: {e}")


def delete_temp_folders(delete_prefetch: bool, log_cb, should_abort):
    targets = []
    user_temp = os.environ.get("TEMP") or os.path.join(os.environ.get("USERPROFILE", ""), "AppData", "Local", "Temp")
    if user_temp:
        targets.append(user_temp)

    win_temp = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Temp")
    targets.append(win_temp)

    if delete_prefetch:
        targets.append(os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Prefetch"))

    for folder in targets:
        if should_abort():
            log_cb("[INFO] Aborted cleanup.")
            return

        if not folder or not os.path.exists(folder):
            log_cb(f"[INFO] Skip (not found): {folder}")
            continue

        log_cb(f"[INFO] Cleaning: {folder}")
        try:
            for name in os.listdir(folder):
                if should_abort():
                    log_cb("[INFO] Aborted cleanup.")
                    return
                safe_rmtree(os.path.join(folder, name), log_cb)
            log_cb(f"[OK] Cleaned: {folder}")
        except PermissionError:
            log_cb(f"[WARN] Permission denied: {folder} (try Admin)")
        except Exception as e:
            log_cb(f"[WARN] Error cleaning {folder}: {e}")


def clear_recycle_bin(log_cb):
    try:
        ctypes.windll.shell32.SHEmptyRecycleBinW(None, None, 0x1 | 0x2 | 0x4)
        log_cb("[OK] Recycle Bin cleared (or already empty).")
    except Exception as e:
        log_cb(f"[WARN] Could not clear Recycle Bin: {e}")


# -------------------------
# Command runner
# -------------------------
class CommandRunner:
    def __init__(self, log_cb):
        self.log_cb = log_cb
        self.current_proc = None
        self._cancel_all = False
        self._skip_step = False

    def reset_all(self):
        # IMPORTANT: allow a new run after Cancel
        self._cancel_all = False
        self._skip_step = False
        self.current_proc = None


    def request_cancel_all(self):
        self._cancel_all = True
        self._terminate_current("Cancel requested")

    def request_skip_step(self):
        self._skip_step = True
        self._terminate_current("Skip requested")

    def reset_flags_for_step(self):
        self._skip_step = False

    def cancel_all_requested(self) -> bool:
        return self._cancel_all

    def skip_requested(self) -> bool:
        return self._skip_step

    def _terminate_current(self, reason: str):
        if self.current_proc:
            try:
                self.log_cb(f"[INFO] {reason}. Terminating current command...")
                self.current_proc.terminate()
            except Exception:
                pass

    def run_cmd(self, cmd):
        shown = cmd if isinstance(cmd, str) else " ".join(cmd)
        self.log_cb(f"\n=== RUN: {shown} ===")

        try:
            self.current_proc = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1,
                universal_newlines=True
            )
        except Exception as e:
            self.log_cb(f"[ERROR] Failed to start command: {e}")
            self.current_proc = None
            return "error"

        try:
            for line in self.current_proc.stdout:
                if self._cancel_all:
                    self._terminate_current("Cancel requested")
                    break
                if self._skip_step:
                    self._terminate_current("Skip requested")
                    break
                self.log_cb(line.rstrip("\n"))
        finally:
            try:
                self.current_proc.stdout.close()
            except Exception:
                pass

        try:
            self.current_proc.wait(timeout=10)
        except Exception:
            try:
                self.current_proc.kill()
            except Exception:
                pass

        self.current_proc = None

        if self._cancel_all:
            self.log_cb("=== STOPPED (cancel) ===\n")
            return "cancel"
        if self._skip_step:
            self.log_cb("=== SKIPPED ===\n")
            return "skip"

        self.log_cb("=== DONE ===\n")
        return "ok"


# -------------------------
# UI helper
# -------------------------
def add_option_with_desc(parent, text, desc, variable, wrap=560):
    row = ttk.Frame(parent)
    row.pack(fill="x", anchor="w", pady=(6, 0))

    cb = ttk.Checkbutton(row, text=text, variable=variable)
    cb.pack(anchor="w")

    lbl = ttk.Label(row, text=desc, foreground="#666666", wraplength=wrap)
    lbl.pack(anchor="w", padx=(26, 0))
    return cb, lbl


# -------------------------
# App
# -------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()

        self.settings = load_settings()
        self.lang = self.settings.get("language", "en")
        self.var_always_admin = tk.BooleanVar(value=bool(self.settings.get("always_admin", False)))

        if self.var_always_admin.get() and not is_admin():
            if messagebox.askyesno("Administrator", "Always run as admin is enabled.\n\nRelaunch as Administrator now?"):
                relaunch_as_admin()

        self.icon_path = set_app_icon(self)

        self.log_queue = queue.Queue()
        self.runner = CommandRunner(self.enqueue_log)
        self.worker_thread = None
        self.running = False

        # Select All
        self.var_select_all = tk.BooleanVar(value=False)
        self._select_all_guard = False

        # Repair
        self.var_dism_scan = tk.BooleanVar(value=False)
        self.var_dism_restore = tk.BooleanVar(value=True)
        self.var_sfc = tk.BooleanVar(value=True)
        self.var_chkdsk = tk.BooleanVar(value=False)
        self.var_chkdsk_mode = tk.StringVar(value="scan")
        self.var_drive = tk.StringVar(value="C:")
        self.var_reset_network = tk.BooleanVar(value=False)

        # Cleanup
        self.var_temp = tk.BooleanVar(value=True)
        self.var_prefetch = tk.BooleanVar(value=False)
        self.var_recycle_bin = tk.BooleanVar(value=True)
        self.var_flush_dns = tk.BooleanVar(value=False)
        self.var_dism_component_cleanup = tk.BooleanVar(value=False)
        self.var_wu_cache = tk.BooleanVar(value=False)

        # list of all option vars for select-all
        self._all_option_vars = [
            self.var_dism_scan,
            self.var_dism_restore,
            self.var_sfc,
            self.var_chkdsk,
            self.var_reset_network,
            self.var_temp,
            self.var_prefetch,
            self.var_recycle_bin,
            self.var_flush_dns,
            self.var_dism_component_cleanup,
            self.var_wu_cache,
        ]

        # Progress
        self.var_step_text = tk.StringVar(value="Idle")
        self.total_steps = 0

        self.title(self.title_text())
        self.geometry(f"{WIN_W}x{WIN_H}")
        self.minsize(1040, 760)

        self.create_menu()
        self.create_ui()

        self.refresh_drive_list()
        self.center_window()

        self.var_chkdsk.trace_add("write", lambda *_: self.update_chkdsk_controls())
        self.update_chkdsk_controls()

        # When any option changes, update select-all state
        for v in self._all_option_vars:
            v.trace_add("write", lambda *_: self.update_select_all_state())

        self.var_select_all.trace_add("write", lambda *_: self.on_select_all_toggled())

        self.after(80, self.flush_log_queue)

        self.apply_language()
        self.after(50, self.center_window)
        self.update_select_all_state()

    def title_text(self):
        return f"Windows Fixer {APP_VERSION}"

    # ---------- Center ----------
    def center_window(self):
        self.update_idletasks()
        w = self.winfo_width()
        h = self.winfo_height()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        x = (sw // 2) - (w // 2)
        y = (sh // 2) - (h // 2)
        self.geometry(f"{w}x{h}+{x}+{y}")

    def center_child(self, tlv):
        tlv.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() - tlv.winfo_width()) // 2
        y = self.winfo_y() + (self.winfo_height() - tlv.winfo_height()) // 2
        tlv.geometry(f"+{x}+{y}")

    # ---------- Menu ----------
    def create_menu(self):
        menubar = tk.Menu(self)

        self.file_menu = tk.Menu(menubar, tearoff=0)
        self.file_menu.add_checkbutton(
            label=("Always run as admin" if self.lang == "en" else "تشغيل دائم كمسؤول"),
            variable=self.var_always_admin,
            command=self.on_toggle_always_admin
        )
        self.file_menu.add_separator()

        self.lang_menu = tk.Menu(self.file_menu, tearoff=0)
        self.lang_var = tk.StringVar(value=self.lang)
        self.lang_menu.add_radiobutton(label="English", value="en", variable=self.lang_var, command=self.on_change_language)
        self.lang_menu.add_radiobutton(label="العربية", value="ar", variable=self.lang_var, command=self.on_change_language)
        self.file_menu.add_cascade(label=("Language" if self.lang == "en" else "اللغة"), menu=self.lang_menu)

        self.file_menu.add_separator()
        self.file_menu.add_command(label=("About" if self.lang == "en" else "حول"), command=self.show_about)
        self.file_menu.add_separator()
        self.file_menu.add_command(label=("Exit" if self.lang == "en" else "خروج"), command=self.destroy)

        menubar.add_cascade(label=("File" if self.lang == "en" else "ملف"), menu=self.file_menu)
        self.config(menu=menubar)

    def on_toggle_always_admin(self):
        self.settings["always_admin"] = bool(self.var_always_admin.get())
        save_settings(self.settings)

        if self.var_always_admin.get() and not is_admin():
            if messagebox.askyesno("Administrator", "Enabled.\n\nRelaunch as Administrator now?"):
                relaunch_as_admin()

    def on_change_language(self):
        self.lang = self.lang_var.get()
        self.settings["language"] = self.lang
        save_settings(self.settings)
        self.create_menu()
        self.apply_language()

    # ---------- Language ----------
    def t(self, key: str) -> str:
        en = {
            "admin_yes": "Admin: YES",
            "admin_no": "Admin: NO (recommended)",
            "run_admin": "Run as Admin",

            "choose_fix": "Choose what to fix",
            "select_all": "Select All",
            "repair": "Repair",
            "cleanup": "Cleanup",
            "progress": "Progress",
            "log": "Log",

            "start": "Start",
            "skip": "Skip Step",
            "cancel": "Cancel",
            "clear_log": "Clear Log",

            "drive": "Drive:",
            "refresh": "Refresh",
            "mode": "Mode:",
            "scan_only": "Scan only",
            "fix_f": "Fix errors (/f)",

            "opt_dism_scan": "Check Windows Image Health (DISM ScanHealth)",
            "desc_dism_scan": "Checks for corruption in the Windows image. Useful before RestoreHealth.",
            "opt_dism_restore": "Repair Windows Image (DISM RestoreHealth)",
            "desc_dism_restore": "Repairs corrupted Windows system image using Windows Update sources.",
            "opt_sfc": "Repair System Files (SFC ScanNow)",
            "desc_sfc": "Scans and repairs protected system files. Best after DISM.",
            "opt_chkdsk": "Check Disk for errors (CHKDSK)",
            "desc_chkdsk": "Scans the selected drive for file system errors. Fix mode may require reboot.",
            "opt_reset_net": "Reset Network Stack (Winsock + TCP/IP)",
            "desc_reset_net": "Fixes common network issues. May require reboot or reconnecting VPN/Wi-Fi.",

            "opt_temp": "Clean Temporary Files",
            "desc_temp": "Deletes files from user Temp and Windows Temp. Some locked files may be skipped.",
            "opt_prefetch": "Clean Prefetch Files",
            "desc_prefetch": "Cleans Prefetch cache. Windows will recreate it. Admin recommended.",
            "opt_recycle": "Empty Recycle Bin",
            "desc_recycle": "Clears deleted files from Recycle Bin to free space immediately.",
            "opt_dns": "Flush DNS Cache",
            "desc_dns": "Resets DNS cache (can help with some internet / browsing issues).",
            "opt_comp": "Clean Windows Component Store (DISM StartComponentCleanup)",
            "desc_comp": "Removes superseded Windows component versions. Safe but may take time.",
            "opt_wu": "Fix Windows Update downloads (Clear Update Cache)",
            "desc_wu": "Stops update services and clears old downloaded update files. Requires Admin.",
        }
        ar = {
            "admin_yes": "المسؤول: نعم",
            "admin_no": "المسؤول: لا (مُفضل)",
            "run_admin": "تشغيل كمسؤول",

            "choose_fix": "اختر عمليات الإصلاح",
            "select_all": "تحديد الكل",
            "repair": "إصلاح",
            "cleanup": "تنظيف",
            "progress": "التقدم",
            "log": "السجل",

            "start": "ابدأ",
            "skip": "تخطي الخطوة",
            "cancel": "إلغاء",
            "clear_log": "مسح السجل",

            "drive": "القرص:",
            "refresh": "تحديث",
            "mode": "الوضع:",
            "scan_only": "فحص فقط",
            "fix_f": "إصلاح الأخطاء (/f)",

            "opt_dism_scan": "فحص سلامة صورة ويندوز (DISM ScanHealth)",
            "desc_dism_scan": "يفحص وجود تلف في صورة النظام. مفيد قبل RestoreHealth.",
            "opt_dism_restore": "إصلاح صورة النظام (DISM RestoreHealth)",
            "desc_dism_restore": "يعالج تلف مكونات ويندوز بالاعتماد على مصادر Windows Update.",
            "opt_sfc": "إصلاح ملفات النظام (SFC ScanNow)",
            "desc_sfc": "يفحص ويصلح ملفات النظام المحمية. الأفضل تشغيله بعد DISM.",
            "opt_chkdsk": "فحص القرص للأخطاء (CHKDSK)",
            "desc_chkdsk": "يفحص نظام الملفات في القرص المحدد. وضع الإصلاح قد يتطلب إعادة تشغيل.",
            "opt_reset_net": "إعادة ضبط الشبكة (Winsock + TCP/IP)",
            "desc_reset_net": "يعالج مشاكل الشبكة الشائعة. قد يتطلب إعادة تشغيل أو إعادة الاتصال.",

            "opt_temp": "تنظيف الملفات المؤقتة",
            "desc_temp": "يحذف ملفات Temp للمستخدم و Windows Temp. قد يتم تخطي الملفات المقفلة.",
            "opt_prefetch": "تنظيف ملفات Prefetch",
            "desc_prefetch": "ينظف كاش Prefetch وسيقوم ويندوز بإعادة إنشائه. يفضل تشغيله كمسؤول.",
            "opt_recycle": "تفريغ سلة المحذوفات",
            "desc_recycle": "يحذف الملفات من سلة المحذوفات لتوفير مساحة فورًا.",
            "opt_dns": "مسح كاش DNS",
            "desc_dns": "يعيد ضبط ذاكرة DNS (قد يساعد في بعض مشاكل التصفح/الإنترنت).",
            "opt_comp": "تنظيف مخزن مكونات ويندوز (DISM StartComponentCleanup)",
            "desc_comp": "يزيل إصدارات المكونات القديمة (آمن لكنه قد يأخذ وقت).",
            "opt_wu": "إصلاح تنزيلات تحديثات ويندوز (مسح كاش التحديث)",
            "desc_wu": "يوقف خدمات التحديث ويمسح ملفات التحديث المحملة. يتطلب تشغيل كمسؤول.",
        }
        return (ar if self.lang == "ar" else en).get(key, key)

    def apply_language(self):
        self.title(self.title_text())

        self.lbl_admin.config(text=(self.t("admin_yes") if is_admin() else self.t("admin_no")))
        self.btn_admin.config(text=self.t("run_admin"))

        self.opts_group.config(text=self.t("choose_fix"))
        self.cb_select_all.config(text=self.t("select_all"))
        self.lbl_repair.config(text=self.t("repair"))
        self.lbl_cleanup.config(text=self.t("cleanup"))

        # Repair
        self.cb_dism_scan.config(text=self.t("opt_dism_scan"))
        self.desc_dism_scan.config(text=self.t("desc_dism_scan"))
        self.cb_dism_restore.config(text=self.t("opt_dism_restore"))
        self.desc_dism_restore.config(text=self.t("desc_dism_restore"))
        self.cb_sfc.config(text=self.t("opt_sfc"))
        self.desc_sfc.config(text=self.t("desc_sfc"))
        self.cb_chkdsk.config(text=self.t("opt_chkdsk"))
        self.desc_chkdsk.config(text=self.t("desc_chkdsk"))
        self.cb_reset_net.config(text=self.t("opt_reset_net"))
        self.desc_reset_net.config(text=self.t("desc_reset_net"))

        self.lbl_drive.config(text=self.t("drive"))
        self.btn_drive_refresh.config(text=self.t("refresh"))
        self.lbl_mode.config(text=self.t("mode"))
        self.rb_scan.config(text=self.t("scan_only"))
        self.rb_fix.config(text=self.t("fix_f"))

        # Cleanup
        self.cb_temp.config(text=self.t("opt_temp"))
        self.desc_temp.config(text=self.t("desc_temp"))
        self.cb_prefetch.config(text=self.t("opt_prefetch"))
        self.desc_prefetch.config(text=self.t("desc_prefetch"))
        self.cb_recycle.config(text=self.t("opt_recycle"))
        self.desc_recycle.config(text=self.t("desc_recycle"))
        self.cb_dns.config(text=self.t("opt_dns"))
        self.desc_dns.config(text=self.t("desc_dns"))
        self.cb_comp.config(text=self.t("opt_comp"))
        self.desc_comp.config(text=self.t("desc_comp"))
        self.cb_wu.config(text=self.t("opt_wu"))
        self.desc_wu.config(text=self.t("desc_wu"))

        self.prog_group.config(text=self.t("progress"))
        self.log_group.config(text=self.t("log"))

        self.btn_start.config(text=self.t("start"))
        self.btn_skip.config(text=self.t("skip"))
        self.btn_cancel.config(text=self.t("cancel"))
        self.btn_clear.config(text=self.t("clear_log"))

        self.refresh_admin_ui()

    # ---------- Select All logic ----------
    def on_select_all_toggled(self):
        if self._select_all_guard:
            return
        self._select_all_guard = True
        try:
            val = bool(self.var_select_all.get())
            for v in self._all_option_vars:
                v.set(val)
        finally:
            self._select_all_guard = False

    def update_select_all_state(self):
        if self._select_all_guard:
            return
        # checked only if ALL options are True
        all_on = all(bool(v.get()) for v in self._all_option_vars)
        self._select_all_guard = True
        try:
            self.var_select_all.set(all_on)
        finally:
            self._select_all_guard = False

    # ---------- UI ----------
    def create_ui(self):
        top = ttk.Frame(self, padding=12)
        top.pack(fill="x")

        self.lbl_admin = ttk.Label(top, text="")
        self.lbl_admin.pack(side="left")

        self.btn_admin = ttk.Button(top, text="", command=self.on_run_as_admin)
        self.btn_admin.pack(side="right")

        self.opts_group = ttk.LabelFrame(self, text="", padding=12)
        self.opts_group.pack(fill="x", padx=12, pady=8)

        # Select All row (top)
        sa_row = ttk.Frame(self.opts_group)
        sa_row.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))
        self.cb_select_all = ttk.Checkbutton(sa_row, text="", variable=self.var_select_all)
        self.cb_select_all.pack(anchor="w")

        left = ttk.Frame(self.opts_group)
        left.grid(row=1, column=0, sticky="nsew", padx=(0, 18))

        right = ttk.Frame(self.opts_group)
        right.grid(row=1, column=1, sticky="nsew")

        self.lbl_repair = ttk.Label(left, text="", font=("Segoe UI", 10, "bold"))
        self.lbl_repair.pack(anchor="w")

        # Repair options
        self.cb_dism_scan, self.desc_dism_scan = add_option_with_desc(left, "", "", self.var_dism_scan, wrap=640)
        self.cb_dism_restore, self.desc_dism_restore = add_option_with_desc(left, "", "", self.var_dism_restore, wrap=640)
        self.cb_sfc, self.desc_sfc = add_option_with_desc(left, "", "", self.var_sfc, wrap=640)

        # CHKDSK row
        ch_row = ttk.Frame(left)
        ch_row.pack(fill="x", anchor="w", pady=(6, 0))
        self.cb_chkdsk = ttk.Checkbutton(ch_row, text="", variable=self.var_chkdsk)
        self.cb_chkdsk.pack(anchor="w")
        self.desc_chkdsk = ttk.Label(ch_row, text="", foreground="#666666", wraplength=640)
        self.desc_chkdsk.pack(anchor="w", padx=(26, 0))

        sub = ttk.Frame(left)
        sub.pack(anchor="w", pady=(6, 0), padx=(26, 0))
        self.lbl_drive = ttk.Label(sub, text="")
        self.lbl_drive.pack(side="left")

        self.drive_combo = ttk.Combobox(sub, width=8, textvariable=self.var_drive, state="readonly")
        self.drive_combo.pack(side="left", padx=(6, 10))

        self.btn_drive_refresh = ttk.Button(sub, text="", command=self.refresh_drive_list)
        self.btn_drive_refresh.pack(side="left")

        mode = ttk.Frame(left)
        mode.pack(anchor="w", pady=(6, 0), padx=(26, 0))
        self.lbl_mode = ttk.Label(mode, text="")
        self.lbl_mode.pack(side="left")

        self.rb_scan = ttk.Radiobutton(mode, text="", value="scan", variable=self.var_chkdsk_mode)
        self.rb_scan.pack(side="left", padx=8)

        self.rb_fix = ttk.Radiobutton(mode, text="", value="fix", variable=self.var_chkdsk_mode)
        self.rb_fix.pack(side="left", padx=8)

        self.cb_reset_net, self.desc_reset_net = add_option_with_desc(left, "", "", self.var_reset_network, wrap=640)

        # Cleanup
        self.lbl_cleanup = ttk.Label(right, text="", font=("Segoe UI", 10, "bold"))
        self.lbl_cleanup.pack(anchor="w")

        self.cb_temp, self.desc_temp = add_option_with_desc(right, "", "", self.var_temp, wrap=520)
        self.cb_prefetch, self.desc_prefetch = add_option_with_desc(right, "", "", self.var_prefetch, wrap=520)
        self.cb_recycle, self.desc_recycle = add_option_with_desc(right, "", "", self.var_recycle_bin, wrap=520)
        self.cb_dns, self.desc_dns = add_option_with_desc(right, "", "", self.var_flush_dns, wrap=520)
        self.cb_comp, self.desc_comp = add_option_with_desc(right, "", "", self.var_dism_component_cleanup, wrap=520)
        self.cb_wu, self.desc_wu = add_option_with_desc(right, "", "", self.var_wu_cache, wrap=520)

        self.opts_group.grid_columnconfigure(0, weight=1)
        self.opts_group.grid_columnconfigure(1, weight=1)

        # Progress
        self.prog_group = ttk.LabelFrame(self, text="", padding=10)
        self.prog_group.pack(fill="x", padx=12, pady=(0, 8))
        ttk.Label(self.prog_group, textvariable=self.var_step_text).pack(anchor="w")
        self.progress = ttk.Progressbar(self.prog_group, orient="horizontal", mode="determinate", maximum=100)
        self.progress.pack(fill="x", pady=6)

        # Buttons
        btns = ttk.Frame(self, padding=(12, 0, 12, 0))
        btns.pack(fill="x")

        self.btn_start = ttk.Button(btns, text="", command=self.on_start)
        self.btn_start.pack(side="left")

        self.btn_skip = ttk.Button(btns, text="", command=self.on_skip, state="disabled")
        self.btn_skip.pack(side="left", padx=8)

        self.btn_cancel = ttk.Button(btns, text="", command=self.on_cancel, state="disabled")
        self.btn_cancel.pack(side="left")

        self.btn_clear = ttk.Button(btns, text="", command=self.on_clear)
        self.btn_clear.pack(side="right")

        # Log
        self.log_group = ttk.LabelFrame(self, text="", padding=8)
        self.log_group.pack(fill="both", expand=True, padx=12, pady=10)

        self.txt = tk.Text(self.log_group, wrap="word")
        self.txt.pack(side="left", fill="both", expand=True)

        sb = ttk.Scrollbar(self.log_group, orient="vertical", command=self.txt.yview)
        sb.pack(side="right", fill="y")
        self.txt.config(yscrollcommand=sb.set)

    def refresh_admin_ui(self):
        if is_admin():
            self.btn_admin.config(state="disabled")
        else:
            self.btn_admin.config(state="normal")

    def update_chkdsk_controls(self):
        enabled = bool(self.var_chkdsk.get())
        self.drive_combo.config(state=("readonly" if enabled else "disabled"))
        self.btn_drive_refresh.config(state=("normal" if enabled else "disabled"))
        self.rb_scan.config(state=("normal" if enabled else "disabled"))
        self.rb_fix.config(state=("normal" if enabled else "disabled"))

    def refresh_drive_list(self):
        drives = list_drives()
        self.drive_combo["values"] = drives
        if self.var_drive.get() not in drives:
            self.var_drive.set(drives[0])

    # ---------- Center ----------
    def center_window(self):
        self.update_idletasks()
        w = self.winfo_width()
        h = self.winfo_height()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        x = (sw // 2) - (w // 2)
        y = (sh // 2) - (h // 2)
        self.geometry(f"{w}x{h}+{x}+{y}")

    # ---------- log ----------
    def enqueue_log(self, msg: str):
        self.log_queue.put(msg)

    def flush_log_queue(self):
        try:
            while True:
                msg = self.log_queue.get_nowait()
                self.txt.insert("end", msg + "\n")
                self.txt.see("end")
        except queue.Empty:
            pass
        self.after(80, self.flush_log_queue)

    def on_clear(self):
        self.txt.delete("1.0", "end")

    # ---------- run/cancel/skip ----------
    def set_running(self, running: bool):
        self.running = running
        self.btn_start.config(state="disabled" if running else "normal")
        self.btn_skip.config(state="normal" if running else "disabled")
        self.btn_cancel.config(state="normal" if running else "disabled")
        if not running:
            self.refresh_admin_ui()

    def on_run_as_admin(self):
        if not is_admin():
            relaunch_as_admin()

    def on_skip(self):
        if self.running:
            self.runner.request_skip_step()

    def on_cancel(self):
        if self.running:
            self.runner.request_cancel_all()

    def should_abort_now(self):
        return self.runner.cancel_all_requested() or self.runner.skip_requested()

    def run_command_step(self, cmd):
        self.runner.reset_flags_for_step()
        return self.runner.run_cmd(cmd)

    # ---------- steps ----------
    def build_steps(self):
        steps = []
        if self.var_temp.get() or self.var_prefetch.get():
            steps.append(("Cleanup (Temp/Prefetch)", self.step_temp_prefetch))
        if self.var_recycle_bin.get():
            steps.append(("Empty Recycle Bin", self.step_clear_recycle))
        if self.var_flush_dns.get():
            steps.append(("Flush DNS Cache", self.step_flush_dns))
        if self.var_dism_component_cleanup.get():
            steps.append(("DISM Component Cleanup", self.step_dism_component_cleanup))
        if self.var_wu_cache.get():
            steps.append(("Clear Windows Update Cache", self.step_wu_cache))

        if self.var_dism_scan.get():
            steps.append(("DISM ScanHealth", self.step_dism_scanhealth))
        if self.var_dism_restore.get():
            steps.append(("DISM RestoreHealth", self.step_dism_restorehealth))
        if self.var_sfc.get():
            steps.append(("SFC ScanNow", self.step_sfc))
        if self.var_chkdsk.get():
            drive = self.var_drive.get().strip().upper()
            mode = self.var_chkdsk_mode.get()
            steps.append((f"CHKDSK ({drive}, {mode})", self.step_chkdsk))
        if self.var_reset_network.get():
            steps.append(("Reset Network Stack", self.step_reset_network))

        return steps

    def on_start(self):
        if self.running:
            return

        # ✅ FIX: reset cancel/skip flags so Start works after Cancel
        self.runner.reset_all()

        steps = self.build_steps()
        if not steps:
            messagebox.showwarning("Nothing selected", "Select at least one task.")
            return

        self.total_steps = len(steps)
        self.progress["value"] = 0
        self.var_step_text.set("Starting...")

        self.set_running(True)
        self.enqueue_log(f"--- {self.title_text()} ---")
        self.enqueue_log("Starting...")

        self.worker_thread = threading.Thread(target=self.worker, args=(steps,), daemon=True)
        self.worker_thread.start()

    def set_progress(self, step_index: int, step_name: str):
        pct = 0 if self.total_steps <= 0 else int((step_index / self.total_steps) * 100)

        def _ui():
            self.var_step_text.set(f"Step {step_index}/{self.total_steps}: {step_name}")
            self.progress["value"] = pct
        self.after(0, _ui)

    def finish_progress(self, msg="Done"):
        def _ui():
            self.var_step_text.set(msg)
            self.progress["value"] = 100
        self.after(0, _ui)

    # ----- step implementations -----
    def step_temp_prefetch(self):
        self.runner.reset_flags_for_step()
        delete_temp_folders(self.var_prefetch.get(), self.enqueue_log, self.should_abort_now)
        if self.runner.cancel_all_requested():
            return "cancel"
        if self.runner.skip_requested():
            return "skip"
        return "ok"

    def step_clear_recycle(self):
        self.runner.reset_flags_for_step()
        clear_recycle_bin(self.enqueue_log)
        return "ok"

    def step_flush_dns(self):
        return self.run_command_step(["ipconfig", "/flushdns"])

    def step_dism_component_cleanup(self):
        return self.run_command_step(["DISM", "/Online", "/Cleanup-Image", "/StartComponentCleanup"])

    def step_wu_cache(self):
        self.runner.reset_flags_for_step()
        if not is_admin():
            self.enqueue_log("[WARN] Windows Update cache cleanup needs Admin. Skipping.")
            return "ok"

        r = self.run_command_step(["net", "stop", "wuauserv"])
        if r in ("cancel", "skip"):
            return r
        r = self.run_command_step(["net", "stop", "bits"])
        if r in ("cancel", "skip"):
            return r

        windir = os.environ.get("WINDIR", r"C:\Windows")
        dl = os.path.join(windir, "SoftwareDistribution", "Download")
        self.enqueue_log(f"[INFO] Cleaning: {dl}")

        try:
            if os.path.exists(dl):
                for name in os.listdir(dl):
                    if self.should_abort_now():
                        return "skip" if self.runner.skip_requested() else "cancel"
                    safe_rmtree(os.path.join(dl, name), self.enqueue_log)
                self.enqueue_log("[OK] Windows Update download cache cleaned.")
            else:
                self.enqueue_log("[INFO] Cache folder not found; skip.")
        except Exception as e:
            self.enqueue_log(f"[WARN] Could not clean Windows Update cache: {e}")

        r = self.run_command_step(["net", "start", "bits"])
        if r in ("cancel", "skip"):
            return r
        return self.run_command_step(["net", "start", "wuauserv"])

    def step_dism_scanhealth(self):
        return self.run_command_step(["DISM", "/Online", "/Cleanup-Image", "/ScanHealth"])

    def step_dism_restorehealth(self):
        return self.run_command_step(["DISM", "/Online", "/Cleanup-Image", "/RestoreHealth"])

    def step_sfc(self):
        return self.run_command_step(["sfc", "/scannow"])

    def step_chkdsk(self):
        drive = self.var_drive.get().strip().upper()
        mode = self.var_chkdsk_mode.get()
        if mode == "scan":
            return self.run_command_step(["chkdsk", drive])
        self.enqueue_log("[INFO] Fix mode may require restart (Windows may ask to schedule it).")
        return self.run_command_step(["cmd", "/c", f"chkdsk {drive} /f"])

    def step_reset_network(self):
        r = self.run_command_step(["netsh", "winsock", "reset"])
        if r in ("cancel", "skip"):
            return r
        return self.run_command_step(["netsh", "int", "ip", "reset"])

    def worker(self, steps):
        cancelled = False
        try:
            for idx, (name, fn) in enumerate(steps, start=1):
                if self.runner.cancel_all_requested():
                    cancelled = True
                    self.enqueue_log("[INFO] Cancelled. Stopping all steps.")
                    self.finish_progress("Cancelled")
                    return

                self.set_progress(idx, name)
                result = fn()

                if result == "cancel":
                    cancelled = True
                    self.enqueue_log("[INFO] Cancelled. Stopping all steps.")
                    self.finish_progress("Cancelled")
                    return

                if result == "skip":
                    self.enqueue_log(f"[INFO] Step skipped: {name}")
                    self.runner._skip_step = False
                    continue

            self.finish_progress("Done")
            self.enqueue_log("All selected tasks finished.")
            # SUCCESS SOUND (only when not cancelled)
            play_success_sound()

        except Exception as e:
            self.enqueue_log(f"[ERROR] {e}")
        finally:
            self.after(0, lambda: self.set_running(False))

    # ---------- About ----------
    def show_about(self):
        win = tk.Toplevel(self)
        win.title("About" if self.lang == "en" else "حول")
        win.resizable(False, False)
        apply_icon_to_tlv(win, self.icon_path)

        frame = ttk.Frame(win, padding=16)
        frame.pack(fill="both", expand=True)

        title = "Windows Fixer"
        sub = ("is a freeware Windows repair & cleanup tool.\n"
               "Runs SFC, DISM, CHKDSK and safe cleanup tasks.") if self.lang == "en" else (
              "أداة مجانية لإصلاح وتنظيف ويندوز.\n"
              "تشغل SFC و DISM و CHKDSK مع عمليات تنظيف آمنة.")

        tk.Label(frame, text=title, font=("Segoe UI", 14, "bold")).pack(pady=(0, 4))
        tk.Label(frame, text=sub, wraplength=520, justify="center").pack(pady=(0, 8))
        tk.Label(frame, text=f"Version {APP_VERSION} • {BUILD_DATE}").pack(pady=(0, 10))

        row = ttk.Frame(frame)
        row.pack()
        tk.Label(row, text="Author: ilukezippo (BoYaqoub)").pack(side="left")

        flag = load_flag_image()
        if flag:
            tk.Label(row, image=flag).pack(side="left", padx=(6, 0))
            win._flag = flag

        email_row = ttk.Frame(frame)
        email_row.pack(pady=(6, 0))
        tk.Label(email_row, text=("For any feedback contact: " if self.lang == "en" else "للملاحظات تواصل على: ")).pack(side="left")
        email_lbl = tk.Label(
            email_row,
            text="ilukezippo@gmail.com",
            fg="#1a73e8",
            cursor="hand2",
            font=("Segoe UI", 9, "underline")
        )
        email_lbl.pack(side="left")
        email_lbl.bind("<Button-1>", lambda e: webbrowser.open("mailto:ilukezippo@gmail.com"))

        link_row = ttk.Frame(frame)
        link_row.pack(pady=(8, 0))
        tk.Label(link_row, text=("Info and Latest Updates at " if self.lang == "en" else "المعلومات وآخر التحديثات: ")).pack(side="left")
        link = tk.Label(
            link_row,
            text=GITHUB_PAGE,
            fg="#1a73e8",
            cursor="hand2",
            font=("Segoe UI", 9, "underline")
        )
        link.pack(side="left")
        link.bind("<Button-1>", lambda e: webbrowser.open(GITHUB_PAGE))

        donate_img = make_donate_image(160, 44)
        win._don = donate_img
        tk.Button(
            frame,
            image=donate_img,
            text=("Donate" if self.lang == "en" else "تبرع"),
            compound="center",
            font=("Segoe UI", 11, "bold"),
            fg="#0f3462",
            activeforeground="#0f3462",
            bd=0,
            highlightthickness=0,
            cursor="hand2",
            relief="flat",
            command=lambda: webbrowser.open(DONATE_PAGE)
        ).pack(pady=(12, 0))

        ttk.Button(frame, text=("Close" if self.lang == "en" else "إغلاق"), command=win.destroy).pack(pady=(10, 0))
        self.center_child(win)


# -------------------------
# Main
# -------------------------
if __name__ == "__main__":
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    App().mainloop()
