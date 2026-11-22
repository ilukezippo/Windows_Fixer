"""
Windows Fixer: repair and clean up Windows, in one window.

DISM, SFC, CHKDSK, network reset and safe cleanup tools, with a progress bar,
a live log, light/dark themes and English/Arabic (right-to-left) UI.
"""
import codecs
import ctypes
import json
import os
import queue
import re
import shutil
import stat
import subprocess
import sys
import threading
import time
import urllib.request
import webbrowser
from datetime import date, datetime
from io import BytesIO

import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog, messagebox, ttk

from PIL import Image, ImageDraw, ImageTk

try:
    import winsound
except ImportError:  # not on Windows
    winsound = None

try:
    import winreg
except ImportError:  # not on Windows
    winreg = None

try:
    import sv_ttk  # Sun Valley theme: the Windows 11 look, light and dark
except ImportError:
    sv_ttk = None


APP_ID = "WindowsFixer"
APP_NAME = "Windows Fixer"
APP_VERSION = "v1.2.0"

DONATE_PAGE = "https://buymeacoffee.com/ilukezippo"
GITHUB_PAGE = "https://github.com/ilukezippo/Windows_Fixer"
GITHUB_API_LATEST = "https://api.github.com/repos/ilukezippo/Windows_Fixer/releases/latest"
GITHUB_RELEASES_PAGE = "https://github.com/ilukezippo/Windows_Fixer/releases/latest"
MORE_APPS_PAGE = "https://ilukezippo.github.io/"

IS_WINDOWS = os.name == "nt"
CREATE_NO_WINDOW = 0x08000000 if IS_WINDOWS else 0
MAX_LOG_LINES = 6000


# ============================================================================
# Text (English / Arabic)
# ============================================================================

STRINGS = {
    "en": {
        "tagline": "Repair and clean up Windows, in one window.",
        "admin_yes": "\u25CF  Running as administrator",
        "admin_no": "\u25CF  Not running as administrator",
        "run_admin": "Run as Admin",
        "menu_file": "File",
        "menu_view": "View",
        "menu_help": "Help",
        "always_admin": "Always run as admin",
        "exit": "Exit",
        "theme": "Theme",
        "theme_system": "Same as Windows",
        "theme_light": "Light",
        "theme_dark": "Dark",
        "language": "Language",
        "check_updates": "Check for updates",
        "github": "GitHub page",
        "more_apps": "More apps",
        "about": "About",
        "toggle_theme": "Switch light / dark",
        "toggle_lang": "العربية",
        "select_all": "Select all",
        "selected_count": "{n} of {total} selected",
        "repair": "Repair",
        "repair_sub": "Fix a damaged Windows installation",
        "cleanup": "Cleanup",
        "cleanup_sub": "Free up space and clear caches",
        "needs_admin": "Needs admin",
        "progress": "Progress",
        "idle": "Ready",
        "starting": "Starting\u2026",
        "step_of": "Step {i} of {n}: {name}",
        "done": "All done",
        "done_issues": "Finished, with problems (see the log)",
        "cancelled": "Cancelled",
        "elapsed": "Elapsed {t}",
        "start": "Start",
        "skip": "Skip Step",
        "cancel": "Cancel",
        "clear_log": "Clear Log",
        "copy_log": "Copy Log",
        "save_log": "Save Log\u2026",
        "log": "Log",
        "drive": "Drive:",
        "refresh": "Refresh",
        "mode": "Mode:",
        "scan_only": "Scan only",
        "fix_f": "Fix errors (/f)",
        # options
        "opt_dism_scan": "Check Windows Image Health",
        "desc_dism_scan": "DISM ScanHealth: checks the Windows image for corruption. Useful before RestoreHealth.",
        "name_dism_scan": "DISM ScanHealth",
        "opt_dism_restore": "Repair Windows Image",
        "desc_dism_restore": "DISM RestoreHealth: repairs the Windows image using Windows Update.",
        "name_dism_restore": "DISM RestoreHealth",
        "opt_sfc": "Repair System Files",
        "desc_sfc": "SFC /scannow: scans and repairs protected system files. Best after DISM.",
        "name_sfc": "SFC /scannow",
        "opt_chkdsk": "Check Disk for errors",
        "desc_chkdsk": "CHKDSK: checks the drive you pick for file system errors. Fix mode may need a restart.",
        "name_chkdsk": "CHKDSK {drive}",
        "opt_reset_net": "Reset Network Stack",
        "desc_reset_net": "Winsock + TCP/IP reset: fixes common network problems. Needs a restart.",
        "name_reset_net": "Network reset",
        "opt_temp": "Clean Temporary Files",
        "desc_temp": "Deletes files from your Temp folder and Windows Temp. Files in use are skipped.",
        "name_temp": "Temporary files",
        "opt_prefetch": "Clean Prefetch Files",
        "desc_prefetch": "Clears the Prefetch cache. Windows rebuilds it.",
        "name_prefetch": "Prefetch files",
        "opt_recycle": "Empty Recycle Bin",
        "desc_recycle": "Frees the space used by deleted files right away.",
        "name_recycle": "Recycle Bin",
        "opt_dns": "Flush DNS Cache",
        "desc_dns": "ipconfig /flushdns: can help with some internet and browsing problems.",
        "name_dns": "Flush DNS",
        "opt_comp": "Clean Windows Component Store",
        "desc_comp": "DISM StartComponentCleanup: removes superseded Windows components. Safe, can take a while.",
        "name_comp": "Component cleanup",
        "opt_wu": "Fix Windows Update downloads",
        "desc_wu": "Stops the update services, clears downloaded update files, then starts them again.",
        "name_wu": "Windows Update cache",
        # dialogs
        "nothing_t": "Nothing selected",
        "nothing": "Tick at least one step.",
        "need_admin_t": "Administrator rights",
        "need_admin": "These steps need administrator rights:\n\n{steps}\n\n"
                      "Yes: restart Windows Fixer as administrator (your choices are kept).\n"
                      "No: run anyway (these steps will probably fail).",
        "elevate_failed": "Windows Fixer was not restarted as administrator.",
        "update_t": "Update",
        "update_available": "Version {tag} is available (you have {cur}).\n\nOpen the download page?",
        "update_latest": "You already have the latest version ({cur}).",
        "update_failed": "Could not check for updates. Please try again later.",
        "close_running_t": "A step is running",
        "close_running": "A step is still running. Stop it and close Windows Fixer?",
        "restart_needed": "Restart Windows to finish: {steps}.",
        "log_copied": "Log copied to the clipboard.",
        "log_saved": "Log saved: {path}",
        "save_failed": "Could not save the log: {err}",
        # about
        "about_sub": "A free tool to repair and clean up Windows.\nRuns DISM, SFC, CHKDSK and safe cleanup tasks.",
        "version": "Version {v}  \u2022  {d}",
        "author": "Author: ilukezippo (BoYaqoub)",
        "info_at": "Info and latest updates:",
        "donate": "Donate",
        "close": "Close",
        # log messages
        "log_header": "--- {app} {v}  \u2022  {when} ---",
        "log_admin_warn": "Not running as administrator: some steps may fail.",
        "log_step": "[{i}/{n}] {name}",
        "log_cleaning": "Cleaning: {path}",
        "log_not_found": "Not found, skipped: {path}",
        "log_cleaned": "Freed {size} ({count} items).",
        "log_in_use": "{count} items were in use or protected and were left alone.",
        "log_aborted": "Stopped.",
        "log_rb_empty": "The Recycle Bin is already empty.",
        "log_rb_ok": "Recycle Bin emptied: {size} ({count} items).",
        "log_rb_fail": "Could not empty the Recycle Bin (code {code}).",
        "log_needs_admin_skip": "This step needs administrator rights. Skipped.",
        "log_svc": "Stopping the Windows Update services\u2026",
        "log_svc_start": "Starting the Windows Update services again\u2026",
        "log_exit_code": "Finished with exit code {code}.",
        "log_hint_admin": "This needs administrator rights: click Run as Admin.",
        "log_chkdsk_issues": "CHKDSK reported problems. Run it again with \u201CFix errors\u201D.",
        "log_chkdsk_fix": "If the drive is in use, the check is scheduled for the next restart.",
        "log_start_fail": "Could not start the command: {err}",
        "log_skip_req": "Skipping this step\u2026",
        "log_cancel_req": "Cancelling\u2026",
        "log_summary": "Summary",
        "res_ok": "OK",
        "res_issues": "Finished, check the output (code {code})",
        "res_failed": "Failed (code {code})",
        "res_skipped": "Skipped",
        "res_cancelled": "Cancelled",
        "res_not_run": "Not run",
        "log_all_done": "All selected steps finished.",
        "log_error": "Unexpected error: {err}",
    },
    "ar": {
        "tagline": "إصلاح وتنظيف ويندوز من نافذة واحدة.",
        "admin_yes": "يعمل بصلاحيات المسؤول  \u25CF",
        "admin_no": "لا يعمل بصلاحيات المسؤول  \u25CF",
        "run_admin": "تشغيل كمسؤول",
        "menu_file": "ملف",
        "menu_view": "عرض",
        "menu_help": "مساعدة",
        "always_admin": "التشغيل دائماً كمسؤول",
        "exit": "خروج",
        "theme": "المظهر",
        "theme_system": "مثل ويندوز",
        "theme_light": "فاتح",
        "theme_dark": "داكن",
        "language": "اللغة",
        "check_updates": "التحقق من التحديثات",
        "github": "صفحة GitHub",
        "more_apps": "تطبيقات أخرى",
        "about": "حول",
        "toggle_theme": "التبديل بين الفاتح والداكن",
        "toggle_lang": "English",
        "select_all": "تحديد الكل",
        "selected_count": "تم تحديد {n} من {total}",
        "repair": "إصلاح",
        "repair_sub": "إصلاح نظام ويندوز التالف",
        "cleanup": "تنظيف",
        "cleanup_sub": "توفير المساحة ومسح الذاكرة المؤقتة",
        "needs_admin": "يحتاج مسؤول",
        "progress": "التقدم",
        "idle": "جاهز",
        "starting": "جارٍ البدء\u2026",
        "step_of": "الخطوة {i} من {n}: {name}",
        "done": "اكتمل كل شيء",
        "done_issues": "انتهى مع وجود مشاكل (راجع السجل)",
        "cancelled": "تم الإلغاء",
        "elapsed": "الوقت {t}",
        "start": "ابدأ",
        "skip": "تخطي الخطوة",
        "cancel": "إلغاء",
        "clear_log": "مسح السجل",
        "copy_log": "نسخ السجل",
        "save_log": "حفظ السجل\u2026",
        "log": "السجل",
        "drive": "القرص:",
        "refresh": "تحديث",
        "mode": "الوضع:",
        "scan_only": "فحص فقط",
        "fix_f": "إصلاح الأخطاء (/f)",
        # options
        "opt_dism_scan": "فحص سلامة صورة ويندوز",
        "desc_dism_scan": "DISM ScanHealth: يفحص صورة النظام بحثاً عن التلف. مفيد قبل RestoreHealth.",
        "name_dism_scan": "DISM ScanHealth",
        "opt_dism_restore": "إصلاح صورة ويندوز",
        "desc_dism_restore": "DISM RestoreHealth: يصلح صورة النظام بالاعتماد على Windows Update.",
        "name_dism_restore": "DISM RestoreHealth",
        "opt_sfc": "إصلاح ملفات النظام",
        "desc_sfc": "SFC /scannow: يفحص ويصلح ملفات النظام المحمية. الأفضل بعد DISM.",
        "name_sfc": "SFC /scannow",
        "opt_chkdsk": "فحص القرص بحثاً عن الأخطاء",
        "desc_chkdsk": "CHKDSK: يفحص نظام الملفات في القرص الذي تختاره. وضع الإصلاح قد يحتاج إعادة تشغيل.",
        "name_chkdsk": "CHKDSK {drive}",
        "opt_reset_net": "إعادة ضبط الشبكة",
        "desc_reset_net": "إعادة ضبط Winsock و TCP/IP: يعالج مشاكل الشبكة الشائعة. يحتاج إعادة تشغيل.",
        "name_reset_net": "إعادة ضبط الشبكة",
        "opt_temp": "تنظيف الملفات المؤقتة",
        "desc_temp": "يحذف ملفات Temp الخاصة بك و Windows Temp. الملفات المستخدمة يتم تخطيها.",
        "name_temp": "الملفات المؤقتة",
        "opt_prefetch": "تنظيف ملفات Prefetch",
        "desc_prefetch": "يمسح ذاكرة Prefetch وسيعيد ويندوز إنشاءها.",
        "name_prefetch": "ملفات Prefetch",
        "opt_recycle": "تفريغ سلة المحذوفات",
        "desc_recycle": "يحرر المساحة التي تشغلها الملفات المحذوفة فوراً.",
        "name_recycle": "سلة المحذوفات",
        "opt_dns": "مسح ذاكرة DNS",
        "desc_dns": "ipconfig /flushdns: قد يساعد في بعض مشاكل الإنترنت والتصفح.",
        "name_dns": "مسح DNS",
        "opt_comp": "تنظيف مخزن مكونات ويندوز",
        "desc_comp": "DISM StartComponentCleanup: يزيل المكونات القديمة. آمن لكنه قد يأخذ وقتاً.",
        "name_comp": "تنظيف المكونات",
        "opt_wu": "إصلاح تنزيلات تحديثات ويندوز",
        "desc_wu": "يوقف خدمات التحديث، ويمسح ملفات التحديث المحملة، ثم يشغلها من جديد.",
        "name_wu": "ذاكرة Windows Update",
        # dialogs
        "nothing_t": "لا يوجد اختيار",
        "nothing": "اختر خطوة واحدة على الأقل.",
        "need_admin_t": "صلاحيات المسؤول",
        "need_admin": "هذه الخطوات تحتاج صلاحيات المسؤول:\n\n{steps}\n\n"
                      "نعم: إعادة تشغيل البرنامج كمسؤول (سيتم الاحتفاظ باختياراتك).\n"
                      "لا: التشغيل على أي حال (غالباً ستفشل هذه الخطوات).",
        "elevate_failed": "لم تتم إعادة تشغيل البرنامج كمسؤول.",
        "update_t": "تحديث",
        "update_available": "يوجد إصدار أحدث {tag} (لديك {cur}).\n\nفتح صفحة التنزيل؟",
        "update_latest": "أنت تستخدم أحدث إصدار ({cur}).",
        "update_failed": "تعذر التحقق من التحديثات. حاول لاحقاً.",
        "close_running_t": "توجد خطوة قيد التشغيل",
        "close_running": "ما زالت هناك خطوة قيد التشغيل. إيقافها وإغلاق البرنامج؟",
        "restart_needed": "أعد تشغيل ويندوز لإكمال: {steps}.",
        "log_copied": "تم نسخ السجل.",
        "log_saved": "تم حفظ السجل: {path}",
        "save_failed": "تعذر حفظ السجل: {err}",
        # about
        "about_sub": "أداة مجانية لإصلاح وتنظيف ويندوز.\nتشغل DISM و SFC و CHKDSK مع عمليات تنظيف آمنة.",
        "version": "الإصدار {v}  \u2022  {d}",
        "author": "المطور: ilukezippo (BoYaqoub)",
        "info_at": "المعلومات وآخر التحديثات:",
        "donate": "تبرع",
        "close": "إغلاق",
        # log messages
        "log_header": "--- {app} {v}  \u2022  {when} ---",
        "log_admin_warn": "البرنامج لا يعمل كمسؤول: قد تفشل بعض الخطوات.",
        "log_step": "[{i}/{n}] {name}",
        "log_cleaning": "جارٍ التنظيف: {path}",
        "log_not_found": "غير موجود، تم التخطي: {path}",
        "log_cleaned": "تم توفير {size} ({count} عنصر).",
        "log_in_use": "تم ترك {count} عنصر لأنها مستخدمة أو محمية.",
        "log_aborted": "تم الإيقاف.",
        "log_rb_empty": "سلة المحذوفات فارغة بالفعل.",
        "log_rb_ok": "تم تفريغ سلة المحذوفات: {size} ({count} عنصر).",
        "log_rb_fail": "تعذر تفريغ سلة المحذوفات (الرمز {code}).",
        "log_needs_admin_skip": "هذه الخطوة تحتاج صلاحيات المسؤول. تم التخطي.",
        "log_svc": "جارٍ إيقاف خدمات Windows Update\u2026",
        "log_svc_start": "جارٍ تشغيل خدمات Windows Update من جديد\u2026",
        "log_exit_code": "انتهى بالرمز {code}.",
        "log_hint_admin": "هذا يحتاج صلاحيات المسؤول: اضغط تشغيل كمسؤول.",
        "log_chkdsk_issues": "وجد CHKDSK مشاكل. شغله مرة أخرى مع «إصلاح الأخطاء».",
        "log_chkdsk_fix": "إذا كان القرص مستخدماً، سيتم الفحص عند إعادة التشغيل القادمة.",
        "log_start_fail": "تعذر تشغيل الأمر: {err}",
        "log_skip_req": "جارٍ تخطي هذه الخطوة\u2026",
        "log_cancel_req": "جارٍ الإلغاء\u2026",
        "log_summary": "الملخص",
        "res_ok": "تم",
        "res_issues": "انتهى، راجع المخرجات (الرمز {code})",
        "res_failed": "فشل (الرمز {code})",
        "res_skipped": "تم التخطي",
        "res_cancelled": "تم الإلغاء",
        "res_not_run": "لم تُشغَّل",
        "log_all_done": "انتهت كل الخطوات المحددة.",
        "log_error": "خطأ غير متوقع: {err}",
    },
}


# --- "Find the problem" panel ------------------------------------------------
STRINGS["en"].update({
    "diag_title": "Find the problem",
    "diag_sub": "Describe what's wrong and get the right steps, or let Windows Fixer check your PC.",
    "describe_hint": "e.g. the internet is slow, Windows Update is stuck, blue screen, PC is slow…",
    "suggest": "Suggest fix",
    "auto_check": "Check for problems",
    "auto_check_startup": "Check for problems at startup",
    "checking": "Checking your PC…",
    "check_none": "No problems found.",
    "check_n": "Found {n} thing(s) worth fixing.",
    "use_suggested": "Tick suggested steps",
    "suggested": "Suggested steps: {steps}",
    "no_match": "Couldn't match that to a specific fix, so a general repair is suggested. "
                "“Check for problems” can find more.",
    "empty_desc": "Type a short description of the problem first.",
    "r_network": "Internet or network problem",
    "r_space": "Slow PC or low disk space",
    "r_update": "Windows Update problem",
    "r_system": "Crashes, errors or parts of Windows not working",
    "r_disk": "Drive or file errors",
    "r_general": "General repair",
    "f_space_low": "Drive {drive} is almost full: {free} free of {total}.",
    "f_space_ok": "Drive {drive} has {free} free of {total}.",
    "f_temp": "Temporary files use {size}.",
    "f_recycle": "The Recycle Bin holds {size}.",
    "f_wu": "Old Windows Update downloads use {size}.",
    "f_reboot": "Windows is waiting for a restart. Restart before running repairs.",
    "f_net_ok": "The internet connection works.",
    "f_dns": "The internet is reachable but website names don't resolve (DNS problem).",
    "f_offline": "No internet connection.",
    "f_disk_bad": "Disk “{name}” reports: {status}. Back up your files.",
    "f_disk_ok": "Disks report healthy.",
    "f_image_bad": "The Windows image is damaged and can be repaired.",
    "f_image_ok": "No Windows image corruption detected.",
    "f_admin_more": "Run as admin to also check the Windows image.",
})
STRINGS["ar"].update({
    "diag_title": "تشخيص المشكلة",
    "diag_sub": "صف المشكلة للحصول على الخطوات المناسبة، أو دع البرنامج يفحص جهازك.",
    "describe_hint": "مثال: الإنترنت بطيء، التحديثات عالقة، شاشة زرقاء، الجهاز بطيء…",
    "suggest": "اقترح الحل",
    "auto_check": "افحص المشاكل",
    "auto_check_startup": "فحص المشاكل عند بدء التشغيل",
    "checking": "جارٍ فحص جهازك…",
    "check_none": "لم يتم العثور على مشاكل.",
    "check_n": "تم العثور على {n} أمر يستحق الإصلاح.",
    "use_suggested": "تحديد الخطوات المقترحة",
    "suggested": "الخطوات المقترحة: {steps}",
    "no_match": "لم أتمكن من تحديد حل معين، لذلك يُقترح إصلاح عام. "
                "زر «افحص المشاكل» قد يجد المزيد.",
    "empty_desc": "اكتب وصفاً قصيراً للمشكلة أولاً.",
    "r_network": "مشكلة في الإنترنت أو الشبكة",
    "r_space": "بطء الجهاز أو قلة المساحة",
    "r_update": "مشكلة في تحديثات ويندوز",
    "r_system": "أعطال أو أخطاء أو أجزاء من ويندوز لا تعمل",
    "r_disk": "أخطاء القرص أو الملفات",
    "r_general": "إصلاح عام",
    "f_space_low": "القرص {drive} شبه ممتلئ: {free} متاحة من {total}.",
    "f_space_ok": "القرص {drive} فيه {free} متاحة من {total}.",
    "f_temp": "الملفات المؤقتة تستخدم {size}.",
    "f_recycle": "سلة المحذوفات تحتوي على {size}.",
    "f_wu": "تنزيلات التحديثات القديمة تستخدم {size}.",
    "f_reboot": "ويندوز ينتظر إعادة التشغيل. أعد التشغيل قبل الإصلاح.",
    "f_net_ok": "الاتصال بالإنترنت يعمل.",
    "f_dns": "الإنترنت متاح لكن أسماء المواقع لا تعمل (مشكلة DNS).",
    "f_offline": "لا يوجد اتصال بالإنترنت.",
    "f_disk_bad": "القرص «{name}» يبلغ عن: {status}. احفظ نسخة من ملفاتك.",
    "f_disk_ok": "الأقراص سليمة.",
    "f_image_bad": "صورة ويندوز تالفة ويمكن إصلاحها.",
    "f_image_ok": "لم يتم العثور على تلف في صورة ويندوز.",
    "f_admin_more": "شغّل كمسؤول لفحص صورة ويندوز أيضاً.",
})


STRINGS["en"].update({
    "tab_diag": "Diagnose",
    "tab_steps": "Repair & Cleanup",
    "tab_log": "Log",
    "tab_settings": "Settings",
    "tab_about": "About",
    "set_appearance": "Appearance",
    "set_behavior": "Options",
    "set_updates": "Updates",
    "always_admin_desc": "Ask for administrator rights every time Windows Fixer starts. Most repairs need them.",
    "restore_point": "Create a restore point before repairs",
    "restore_point_desc": "A safety net: you can undo the changes from System Restore. "
                          "Windows makes at most one every 24 hours, and System Protection must be on.",
    "auto_check_startup_desc": "Look for problems as soon as the app opens. The check only reads; it changes nothing.",
    "play_sound": "Play a sound when finished",
    "play_sound_desc": "So you know when a long repair is done.",
    "current_version": "You have Windows Fixer {v}.",
    "about_made_by": "Made by",
    "made_in": "Made in Kuwait",
    "about_free": "Free and open source (MIT license).",
    "about_links": "Links",
    "more_apps_btn": "More apps by ilukezippo",
    "report": "Report a problem",
    "donate_note": "If Windows Fixer saved you time, you can buy me a coffee.",
    "log_rp": "Creating a restore point\u2026",
    "log_rp_ok": "Restore point created.",
    "log_hint_source": "DISM couldn't get the files it needs for the repair. Install all Windows updates, "
                       "restart, and try again. If it still fails: Settings \u203A System \u203A Recovery \u203A "
                       "\u201CFix problems using Windows Update\u201D reinstalls Windows and keeps your files and apps.",
    "log_reboot_first": "Windows is waiting for a restart. This repair may fail until you restart.",
    "log_rp_fail": "Couldn't create a restore point (System Protection may be off, or one was made in the "
                   "last 24 hours). Carrying on.",
})
STRINGS["ar"].update({
    "tab_diag": "التشخيص",
    "tab_steps": "الإصلاح والتنظيف",
    "tab_log": "السجل",
    "tab_settings": "الإعدادات",
    "tab_about": "حول",
    "set_appearance": "المظهر",
    "set_behavior": "الخيارات",
    "set_updates": "التحديثات",
    "always_admin_desc": "طلب صلاحيات المسؤول في كل مرة يبدأ فيها البرنامج. أغلب عمليات الإصلاح تحتاجها.",
    "restore_point": "إنشاء نقطة استعادة قبل الإصلاح",
    "restore_point_desc": "شبكة أمان: يمكنك التراجع عن التغييرات من استعادة النظام. "
                          "ينشئ ويندوز نقطة واحدة كل 24 ساعة كحد أقصى، ويجب أن تكون حماية النظام مفعّلة.",
    "auto_check_startup": "فحص المشاكل عند بدء التشغيل",
    "auto_check_startup_desc": "البحث عن المشاكل فور فتح البرنامج. الفحص للقراءة فقط ولا يغيّر شيئاً.",
    "play_sound": "تشغيل صوت عند الانتهاء",
    "play_sound_desc": "لتعرف متى ينتهي الإصلاح الطويل.",
    "current_version": "لديك Windows Fixer {v}.",
    "about_made_by": "من تطوير",
    "made_in": "صُنع في الكويت",
    "about_free": "مجاني ومفتوح المصدر (رخصة MIT).",
    "about_links": "روابط",
    "more_apps_btn": "تطبيقات أخرى من ilukezippo",
    "report": "الإبلاغ عن مشكلة",
    "donate_note": "إذا وفّر لك البرنامج الوقت، يمكنك دعمي بكوب قهوة.",
    "log_rp": "جارٍ إنشاء نقطة استعادة\u2026",
    "log_rp_ok": "تم إنشاء نقطة الاستعادة.",
    "log_hint_source": "لم يتمكن DISM من الحصول على ملفات الإصلاح. ثبّت كل تحديثات ويندوز، وأعد التشغيل، ثم حاول مرة أخرى. "
                       "إذا استمرت المشكلة: الإعدادات \u203A النظام \u203A الاسترداد \u203A «إصلاح المشاكل باستخدام Windows Update» "
                       "يعيد تثبيت ويندوز مع الاحتفاظ بملفاتك وتطبيقاتك.",
    "log_reboot_first": "ويندوز ينتظر إعادة التشغيل. قد يفشل هذا الإصلاح حتى تعيد التشغيل.",
    "log_rp_fail": "تعذر إنشاء نقطة استعادة (قد تكون حماية النظام معطلة، أو تم إنشاء نقطة خلال آخر 24 ساعة). سيستمر البرنامج.",
})


# Every step: (key, group, needs_admin, ticked by default)
STEPS = [
    ("dism_scan", "repair", True, False),
    ("dism_restore", "repair", True, True),
    ("sfc", "repair", True, True),
    ("chkdsk", "repair", True, False),
    ("reset_net", "repair", True, False),
    ("temp", "cleanup", False, True),
    ("prefetch", "cleanup", True, False),
    ("recycle", "cleanup", False, True),
    ("dns", "cleanup", False, False),
    ("comp", "cleanup", True, False),
    ("wu", "cleanup", True, False),
]
# DISM couldn't get the repair files (source not found / not available / mismatched)
DISM_SOURCE_ERRORS = {0x800F081F, 0x800F0906, 0x800F0907, 0x800F0915, 0x800F0950, 0x800F0954}

STEP_INFO = {k: {"group": g, "admin": a, "default": d} for k, g, a, d in STEPS}

# The order the steps run in: quick cleanups first, then the image repair
# (DISM before SFC, which uses the repaired image), disk and network last.
RUN_ORDER = ["temp", "prefetch", "recycle", "dns", "wu",
             "dism_scan", "dism_restore", "comp", "sfc", "chkdsk", "reset_net"]


# ============================================================================
# "Describe the problem" -> suggested steps (offline keyword matching)
# ============================================================================

PROBLEM_RULES = [
    ("r_network", ["dns", "reset_net"], [
        "internet", "network", "wifi", "wi-fi", "wi fi", "wireless", "ethernet", "dns", "website", "web site",
        "browser", "connect", "connection", "offline", "no internet", "ping", "vpn", "ip address", "lan",
        "page won", "pages don", "can't load", "cant load", "not loading", "router", "online",
        "انترنت", "نت ", "شبكه", "شبكة", "واي فاي", "وايفاي", "وای فای", "اتصال", "متصل", "مواقع", "موقع",
        "المتصفح", "متصفح", "راوتر", "بنق", "البنق", "dns",
    ]),
    ("r_space", ["temp", "recycle", "comp", "prefetch"], [
        "slow", "lag", "laggy", "sluggish", "space", "storage", "disk full", "drive full", "full disk",
        "low disk", "almost full", "out of space", "takes forever", "speed up", "clean", "junk", "cache",
        "بطيء", "بطئ", "بطء", "ثقيل", "تعليق", "يعلق", "مساحه", "مساحة", "ممتلئ", "ممتلي", "امتلا",
        "تنظيف", "ملفات مؤقته", "مؤقت", "تسريع",
    ]),
    ("r_update", ["wu", "dism_restore", "sfc"], [
        "update", "windows update", "updates", "0x80", "0x8007", "patch", "upgrade", "kb",
        "تحديث", "التحديثات", "تحديثات", "ترقيه",
    ]),
    ("r_system", ["dism_restore", "sfc"], [
        "crash", "blue screen", "bsod", "error", "corrupt", "broken", "dll", "missing", "freeze", "frozen",
        "not responding", "start menu", "taskbar", "settings", "won't open", "wont open", "doesn't open",
        "does not open", "black screen", "restart", "reboot", "boot", "startup", "system file", "explorer",
        "search", "app", "program", "stopped working",
        "شاشه زرقاء", "شاشة زرقاء", "شاشه سوداء", "تعطل", "يتعطل", "يتوقف", "توقف", "خطا", "خطأ", "تلف", "تالف",
        "معطوب", "قائمه ابدا", "قائمة ابدأ", "شريط المهام", "الاعدادات", "لا يفتح", "ما يفتح", "مايفتح",
        "يعيد التشغيل", "اعاده تشغيل", "الاقلاع", "برنامج", "برامج",
    ]),
    ("r_disk", ["chkdsk"], [
        "disk", "drive", "hard drive", "hdd", "ssd", "bad sector", "clicking", "file system", "can't read",
        "unreadable", "chkdsk", "usb", "corrupt file", "files disappear",
        "قرص", "هارد", "الهارد", "القرص", "فلاش", "قطاع", "ملفات تالفه", "لا يقرا",
    ]),
]

_AR_DIACRITICS = re.compile("[\u064B-\u0652\u0640]")


def normalize_text(text: str) -> str:
    t = " " + text.lower().replace("\u2019", "'") + " "
    t = _AR_DIACRITICS.sub("", t)
    for a, b in (("أ", "ا"), ("إ", "ا"), ("آ", "ا"),
                 ("ة", "ه"), ("ى", "ي"), ("ؤ", "و"), ("ئ", "ي")):
        t = t.replace(a, b)
    return t


def suggest_from_text(text: str):
    """Returns [(reason_key, [step keys])] for a free-text description of the problem."""
    t = normalize_text(text)
    matches = []
    for reason, steps, words in PROBLEM_RULES:
        score = sum(1 for w in words if normalize_text(w).strip() in t)
        if score:
            matches.append((score, reason, steps))
    matches.sort(key=lambda m: -m[0])
    return [(reason, steps) for _, reason, steps in matches]


# ============================================================================
# Automatic problem check (read-only, safe to run any time)
# ============================================================================

def folder_size(path: str, deadline: float) -> int:
    total = 0
    stack = [path]
    while stack:
        if time.time() > deadline:
            break
        try:
            with os.scandir(stack.pop()) as it:
                for e in it:
                    try:
                        if e.is_dir(follow_symlinks=False):
                            if not _is_link(e):
                                stack.append(e.path)
                        else:
                            total += e.stat(follow_symlinks=False).st_size
                    except OSError:
                        pass
        except OSError:
            pass
    return total


def reboot_pending() -> bool:
    if not winreg:
        return False
    keys = [r"SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending",
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired"]
    for k in keys:
        try:
            winreg.CloseKey(winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, k))
            return True
        except OSError:
            pass
    return False


def check_internet():
    """'ok', 'dns' (reachable by IP but names fail) or 'offline'."""
    import socket

    def can_connect(host):
        try:
            with socket.create_connection((host, 443), timeout=4):
                return True
        except OSError:
            return False

    ip_ok = can_connect("1.1.1.1") or can_connect("8.8.8.8")
    try:
        socket.getaddrinfo("www.microsoft.com", 443)
        dns_ok = True
    except OSError:
        dns_ok = False
    if dns_ok and (ip_ok or can_connect("www.microsoft.com")):
        return "ok"
    if ip_ok and not dns_ok:
        return "dns"
    return "offline"


def disk_health():
    """[(name, status)] for every physical disk whose status is not Healthy; None if unknown."""
    if not IS_WINDOWS:
        return None
    ps = ("Get-PhysicalDisk | Select-Object FriendlyName,"
          "@{n='H';e={[string]$_.HealthStatus}} | ConvertTo-Json -Compress")
    try:
        out = subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", ps],
                             capture_output=True, timeout=25, creationflags=CREATE_NO_WINDOW)
        data = json.loads(out.stdout.decode("utf-8", "replace") or "null")
    except Exception:
        return None
    if data is None:
        return None
    if isinstance(data, dict):
        data = [data]
    bad = []
    for d in data:
        status = str(d.get("H") or "")
        if status and status.lower() not in ("healthy", "0"):
            bad.append((str(d.get("FriendlyName") or "?").strip(), status))
    return bad


def dism_check_health():
    """True if damaged, False if fine, None if unknown (needs admin)."""
    try:
        out = subprocess.run(["DISM", "/Online", "/Cleanup-Image", "/CheckHealth", "/English"],
                             capture_output=True, timeout=180, creationflags=CREATE_NO_WINDOW)
    except Exception:
        return None
    text = out.stdout.decode(oem_encoding(), "replace").lower()
    if "repairable" in text or "not repairable" in text:
        return True
    if "no component store corruption" in text:
        return False
    return None


def run_problem_check(admin: bool):
    """
    Returns [(tag, text_key, kwargs, [suggested steps])].
    tag: 'warn'/'err' for problems, 'ok' for healthy, 'info' for notes.
    """
    findings = []
    GB = 1024 ** 3

    drive = system_drive()
    try:
        usage = shutil.disk_usage(drive + "\\")
        low = usage.free < 10 * GB or usage.free < usage.total * 0.10
        kw = {"drive": drive, "free": human_size(usage.free), "total": human_size(usage.total)}
        if low:
            findings.append(("warn", "f_space_low", kw, ["temp", "recycle", "comp", "wu"]))
        else:
            findings.append(("ok", "f_space_ok", kw, []))
    except OSError:
        pass

    if reboot_pending():
        findings.append(("warn", "f_reboot", {}, []))

    deadline = time.time() + 6
    temps = {os.environ.get("TEMP") or "", os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Temp")}
    tsize = sum(folder_size(p, deadline) for p in temps if p and os.path.isdir(p))
    if tsize > 500 * 1024 ** 2:
        findings.append(("warn", "f_temp", {"size": human_size(tsize)}, ["temp"]))

    if IS_WINDOWS:
        try:
            rb = query_recycle_bin()
            if rb and rb[0] > 1 * GB:
                findings.append(("warn", "f_recycle", {"size": human_size(rb[0])}, ["recycle"]))
        except Exception:
            pass

    dl = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "SoftwareDistribution", "Download")
    if os.path.isdir(dl):
        wsize = folder_size(dl, time.time() + 4)
        if wsize > 1 * GB:
            findings.append(("warn", "f_wu", {"size": human_size(wsize)}, ["wu"]))

    net = check_internet()
    if net == "ok":
        findings.append(("ok", "f_net_ok", {}, []))
    elif net == "dns":
        findings.append(("err", "f_dns", {}, ["dns", "reset_net"]))
    else:
        findings.append(("err", "f_offline", {}, ["dns", "reset_net"]))

    bad = disk_health()
    if bad:
        for name, status in bad:
            findings.append(("err", "f_disk_bad", {"name": name, "status": status}, ["chkdsk"]))
    elif bad == []:
        findings.append(("ok", "f_disk_ok", {}, []))

    if admin:
        damaged = dism_check_health()
        if damaged:
            findings.append(("err", "f_image_bad", {}, ["dism_restore", "sfc"]))
        elif damaged is False:
            findings.append(("ok", "f_image_ok", {}, []))
    else:
        findings.append(("info", "f_admin_more", {}, []))
    return findings


# ============================================================================
# Colours
# ============================================================================

PALETTES = {
    "light": {
        "bg": "#fafafa", "card": "#ffffff", "text": "#1c1c1c", "muted": "#5f6368",
        "log_bg": "#ffffff", "log_fg": "#1f1f1f", "select": "#cce4f7",
        "ok": "#0f7b0f", "warn": "#9d5d00", "err": "#c42b1c", "head": "#005fb8",
        "info": "#5f6368", "link": "#005fb8", "border": "#e0e0e0",
    },
    "dark": {
        "bg": "#1c1c1c", "card": "#2b2b2b", "text": "#ffffff", "muted": "#a8a8a8",
        "log_bg": "#1f1f1f", "log_fg": "#e8e8e8", "select": "#264f78",
        "ok": "#6ccb5f", "warn": "#f5c451", "err": "#ff99a4", "head": "#60cdff",
        "info": "#a8a8a8", "link": "#60cdff", "border": "#3a3a3a",
    },
}


# ============================================================================
# Helpers
# ============================================================================

def resource_path(relative_path: str) -> str:
    """Prefer files next to the exe/script (portable), then PyInstaller's bundle."""
    if getattr(sys, "frozen", False):
        here = os.path.dirname(sys.executable)
    else:
        here = os.path.dirname(os.path.abspath(__file__))
    candidate = os.path.join(here, relative_path)
    if os.path.exists(candidate):
        return candidate
    base = getattr(sys, "_MEIPASS", here)
    return os.path.join(base, relative_path)


def build_date() -> str:
    """The date the exe (or script) was built, not today's date."""
    path = sys.executable if getattr(sys, "frozen", False) else os.path.abspath(__file__)
    try:
        return date.fromtimestamp(os.path.getmtime(path)).isoformat()
    except OSError:
        return date.today().isoformat()


def human_size(n: float) -> str:
    for unit in ("B", "KB", "MB", "GB", "TB"):
        if n < 1024 or unit == "TB":
            return f"{n:.0f} {unit}" if unit == "B" else f"{n:.1f} {unit}"
        n /= 1024
    return f"{n:.1f} TB"


def format_code(code: int) -> str:
    """DISM returns HRESULTs such as 0x800F081F; show those in hex."""
    if code is None:
        return "?"
    if code < 0:
        code &= 0xFFFFFFFF
    return f"0x{code:08X}" if code > 0xFFFF else str(code)


def _settings_path() -> str:
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    folder = os.path.join(base, APP_ID)
    os.makedirs(folder, exist_ok=True)
    return os.path.join(folder, "settings.json")


def windows_ui_is_arabic() -> bool:
    if not IS_WINDOWS:
        return False
    try:
        return (ctypes.windll.kernel32.GetUserDefaultUILanguage() & 0x3FF) == 0x01
    except Exception:
        return False


def default_settings() -> dict:
    return {
        "always_admin": False,
        "language": "ar" if windows_ui_is_arabic() else "en",
        "theme": "system",
        "selected": {k: d for k, _, _, d in STEPS},
        "chkdsk_mode": "scan",
        "drive": "C:",
        "size": None,
        "auto_check": True,
        "restore_point": True,
        "sound": True,
    }


def load_settings() -> dict:
    data = default_settings()
    try:
        with open(_settings_path(), "r", encoding="utf-8") as f:
            saved = json.load(f)
        if isinstance(saved, dict):
            sel = dict(data["selected"])
            if isinstance(saved.get("selected"), dict):
                sel.update({k: bool(v) for k, v in saved["selected"].items() if k in STEP_INFO})
            data.update(saved)
            data["selected"] = sel
    except (OSError, ValueError):
        pass
    if data.get("language") not in STRINGS:
        data["language"] = "en"
    if data.get("theme") not in ("system", "light", "dark"):
        data["theme"] = "system"
    return data


def save_settings(data: dict):
    try:
        path = _settings_path()
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        os.replace(tmp, path)
    except OSError:
        pass


def is_admin() -> bool:
    if not IS_WINDOWS:
        return False
    try:
        return ctypes.windll.shell32.IsUserAnAdmin() != 0
    except Exception:
        return False


def relaunch_as_admin() -> bool:
    """Start an elevated copy of the app. True if it started (UAC accepted)."""
    if not IS_WINDOWS:
        return False
    if getattr(sys, "frozen", False):
        exe = sys.executable
        args = sys.argv[1:]
        workdir = os.path.dirname(exe)
    else:
        exe = sys.executable
        pythonw = os.path.join(os.path.dirname(exe), "pythonw.exe")
        if os.path.exists(pythonw):
            exe = pythonw  # no console window for the elevated copy
        script = os.path.abspath(sys.argv[0])
        args = [script] + sys.argv[1:]
        workdir = os.path.dirname(script)
    try:
        rc = ctypes.windll.shell32.ShellExecuteW(None, "runas", exe, subprocess.list2cmdline(args), workdir, 1)
        return int(rc) > 32
    except Exception:
        return False


def system_prefers_dark() -> bool:
    if not winreg:
        return False
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                             r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize")
        with key:
            return winreg.QueryValueEx(key, "AppsUseLightTheme")[0] == 0
    except OSError:
        return False


def set_title_bar_dark(window, dark: bool):
    """Dark or light window title bar (Windows 10 1809+ / Windows 11)."""
    if not IS_WINDOWS:
        return
    try:
        window.update_idletasks()
        hwnd = ctypes.windll.user32.GetParent(window.winfo_id()) or window.winfo_id()
        value = ctypes.c_int(1 if dark else 0)
        for attr in (20, 19):  # DWMWA_USE_IMMERSIVE_DARK_MODE (new id, then the pre-20H1 id)
            if ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, attr, ctypes.byref(value), ctypes.sizeof(value)) == 0:
                break
        # Windows 10 only repaints the title bar when the window changes
        alpha = window.attributes("-alpha")
        window.attributes("-alpha", 0.99)
        window.attributes("-alpha", alpha)
    except Exception:
        pass


def list_drives():
    if not IS_WINDOWS:
        return ["C:"]
    drives = []
    try:
        bitmask = ctypes.windll.kernel32.GetLogicalDrives()
        for i in range(26):
            if bitmask & (1 << i):
                letter = chr(ord("A") + i)
                # 3 = fixed disk, 2 = removable: the drives CHKDSK can check
                if ctypes.windll.kernel32.GetDriveTypeW(f"{letter}:\\") in (2, 3):
                    drives.append(f"{letter}:")
    except Exception:
        pass
    return drives or ["C:"]


def system_drive() -> str:
    return (os.environ.get("SystemDrive") or "C:").upper()


def oem_encoding() -> str:
    """Console programs write in the OEM code page (cp437, cp720 for Arabic\u2026)."""
    if IS_WINDOWS:
        try:
            enc = f"cp{ctypes.windll.kernel32.GetOEMCP()}"
            codecs.lookup(enc)
            return enc
        except Exception:
            pass
    return "utf-8"


# ---------------------------------------------------------------------------
# Command output decoding
# ---------------------------------------------------------------------------

_PCT_TOTAL = re.compile(r"Total:\s*(\d+)\s*%")
_PCT_ANY = re.compile(r"(\d{1,3}(?:[.,]\d+)?)\s*%")
_LINE_BREAK = re.compile(r"\r*\n|\r+")
# DISM's "[=====   45.3%     ]" and CHKDSK's "Progress: ... Stage: 45%; Total: 30% ..."
_PROGRESS_LINE = re.compile(r"^\[[=\s]*\d{1,3}(?:[.,]\d+)?\s*%[=\s]*\]$|Stage:\s*\d+\s*%|Total:\s*\d+\s*%")


def parse_percent(line: str):
    m = _PCT_TOTAL.search(line)  # CHKDSK: "... Stage: 45%; Total: 30%; ETA ..."
    if m:
        return float(m.group(1))
    found = _PCT_ANY.findall(line)
    if found:
        try:
            v = float(found[-1].replace(",", "."))
        except ValueError:
            return None
        if 0 <= v <= 100:
            return v
    return None


class OutputDecoder:
    """
    Turns raw command output into lines.

    * SFC writes UTF-16; DISM, CHKDSK, netsh write in the OEM code page.
    * Progress (DISM's [====  45%  ], SFC's "Verification 45% complete")
      is redrawn with a bare carriage return: those are reported as live
      progress instead of flooding the log.
    """

    def __init__(self, on_line, on_live):
        self.on_line = on_line
        self.on_live = on_live
        self.decoder = None
        self.buf = ""
        self.prev_blank = True

    def _pick_decoder(self, chunk: bytes):
        nul = chunk.count(b"\x00")
        enc = "utf-16-le" if len(chunk) >= 2 and nul >= len(chunk) // 4 else oem_encoding()
        self.decoder = codecs.getincrementaldecoder(enc)(errors="replace")

    def feed(self, chunk: bytes):
        if not chunk:
            return
        if self.decoder is None:
            self._pick_decoder(chunk)
        self.buf += self.decoder.decode(chunk)
        self._drain(final=False)

    def close(self):
        if self.decoder is not None:
            self.buf += self.decoder.decode(b"", final=True)
        self._drain(final=True)

    def _drain(self, final: bool):
        while True:
            m = _LINE_BREAK.search(self.buf)
            if not m:
                break
            is_cr_only = "\n" not in m.group()
            if is_cr_only and m.end() == len(self.buf) and not final:
                break  # might be the start of "\r\n": wait for more
            line = self.buf[:m.start()]
            self.buf = self.buf[m.end():]
            self._emit(line, live=is_cr_only)
        if final and self.buf:
            self._emit(self.buf, live=False)
            self.buf = ""

    def _emit(self, line: str, live: bool):
        line = line.replace("\x00", "").replace("\ufeff", "").rstrip()
        if not live and _PROGRESS_LINE.search(line.strip()):
            live = True
        if live:
            if line.strip():
                self.on_live(line.strip())
            return
        if not line.strip():
            if self.prev_blank:
                return
            self.prev_blank = True
            self.on_line("")
            return
        self.prev_blank = False
        self.on_line(line)


# ---------------------------------------------------------------------------
# Running commands
# ---------------------------------------------------------------------------

class CommandRunner:
    def __init__(self, on_line, on_live):
        self.on_line = on_line
        self.on_live = on_live
        self._lock = threading.Lock()
        self._proc = None
        self._proc_cancellable = True
        self.cancel_all = False
        self.skip_step = False

    def reset(self):
        self.cancel_all = False
        self.skip_step = False

    def new_step(self):
        self.skip_step = False

    def should_abort(self) -> bool:
        return self.cancel_all or self.skip_step

    def request_cancel(self):
        self.cancel_all = True
        self._kill_current()

    def request_skip(self):
        self.skip_step = True
        self._kill_current()

    def _kill_current(self):
        with self._lock:
            proc, cancellable = self._proc, self._proc_cancellable
        if not proc or not cancellable or proc.poll() is not None:
            return
        try:
            # /T also stops the children (DISM runs DismHost.exe)
            subprocess.run(["taskkill", "/PID", str(proc.pid), "/T", "/F"],
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                           creationflags=CREATE_NO_WINDOW, timeout=15)
        except Exception:
            pass
        try:
            if proc.poll() is None:
                proc.kill()
        except Exception:
            pass

    def run(self, cmd, cancellable=True, stdin_text=None):
        """Run a command, streaming its output. Returns (status, exit_code)."""
        self.on_line("")
        self.on_line("> " + subprocess.list2cmdline(cmd), "head")
        if cancellable and self.should_abort():
            return ("cancel" if self.cancel_all else "skip"), None
        try:
            proc = subprocess.Popen(
                cmd,
                stdin=subprocess.PIPE if stdin_text else subprocess.DEVNULL,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                creationflags=CREATE_NO_WINDOW,
            )
        except OSError as e:
            return "error", str(e)

        with self._lock:
            self._proc, self._proc_cancellable = proc, cancellable

        if stdin_text:
            try:
                proc.stdin.write(stdin_text.encode("ascii"))
                proc.stdin.close()
            except OSError:
                pass

        decoder = OutputDecoder(self.on_line, self.on_live)
        try:
            read = getattr(proc.stdout, "read1", proc.stdout.read)
            while True:
                chunk = read(4096)
                if not chunk:
                    break
                decoder.feed(chunk)
            decoder.close()
        finally:
            try:
                proc.stdout.close()
            except OSError:
                pass
            try:
                code = proc.wait(timeout=30)
            except subprocess.TimeoutExpired:
                proc.kill()
                code = proc.wait()
            with self._lock:
                self._proc = None
                self._proc_cancellable = True

        if cancellable and self.cancel_all:
            return "cancel", code
        if cancellable and self.skip_step:
            return "skip", code
        return ("ok" if code == 0 else "failed"), code


# ---------------------------------------------------------------------------
# Cleanup
# ---------------------------------------------------------------------------

FILE_ATTRIBUTE_REPARSE_POINT = 0x400


class _Abort(Exception):
    pass


def _is_link(entry) -> bool:
    """Symlinks and junctions: delete the link itself, never what it points to."""
    try:
        if entry.is_symlink():
            return True
        st = entry.stat(follow_symlinks=False)
        return bool(getattr(st, "st_file_attributes", 0) & FILE_ATTRIBUTE_REPARSE_POINT)
    except OSError:
        return True


def _unlink(path: str):
    try:
        os.unlink(path)
    except PermissionError:
        os.chmod(path, stat.S_IWRITE)  # read-only files
        os.unlink(path)


def clean_folder_contents(folder: str, should_abort):
    """Delete everything inside folder. Returns (bytes_freed, items_deleted, items_left, aborted)."""
    stats = {"freed": 0, "deleted": 0, "left": 0}

    def remove(entry) -> bool:
        if should_abort():
            raise _Abort
        try:
            if _is_link(entry):
                try:
                    os.rmdir(entry.path)
                except OSError:
                    os.unlink(entry.path)
                stats["deleted"] += 1
                return True
            if entry.is_dir(follow_symlinks=False):
                all_gone = True
                with os.scandir(entry.path) as it:
                    for child in it:
                        all_gone = remove(child) and all_gone
                if all_gone:
                    os.rmdir(entry.path)
                    stats["deleted"] += 1
                return all_gone
            size = entry.stat(follow_symlinks=False).st_size
            _unlink(entry.path)
            stats["freed"] += size
            stats["deleted"] += 1
            return True
        except _Abort:
            raise
        except OSError:
            stats["left"] += 1
            return False

    try:
        with os.scandir(folder) as it:
            entries = list(it)
        for entry in entries:
            remove(entry)
    except _Abort:
        return stats["freed"], stats["deleted"], stats["left"], True
    return stats["freed"], stats["deleted"], stats["left"], False


if ctypes.sizeof(ctypes.c_void_p) == 4:
    class SHQUERYRBINFO(ctypes.Structure):
        _pack_ = 1
        _fields_ = [("cbSize", ctypes.c_ulong), ("i64Size", ctypes.c_longlong), ("i64NumItems", ctypes.c_longlong)]
else:
    class SHQUERYRBINFO(ctypes.Structure):
        _fields_ = [("cbSize", ctypes.c_ulong), ("i64Size", ctypes.c_longlong), ("i64NumItems", ctypes.c_longlong)]


def query_recycle_bin():
    info = SHQUERYRBINFO()
    info.cbSize = ctypes.sizeof(info)
    hr = ctypes.windll.shell32.SHQueryRecycleBinW(None, ctypes.byref(info))
    if hr != 0:
        return None
    return info.i64Size, info.i64NumItems


# ---------------------------------------------------------------------------
# Images
# ---------------------------------------------------------------------------

def make_donate_image(w=160, h=44):
    r = h // 2
    top, mid, bot = (255, 187, 71), (247, 162, 28), (225, 140, 22)
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
    ImageDraw.Draw(im).rounded_rectangle([0.5, 0.5, w - 1.5, h - 1.5], radius=r, outline=(200, 120, 20, 255), width=2)
    bio = BytesIO()
    im.save(bio, format="PNG")
    return tk.PhotoImage(data=bio.getvalue())


def load_icon_image(size: int):
    path = resource_path("icon.ico")
    try:
        with Image.open(path) as im:
            sizes = sorted(im.info.get("sizes") or [im.size], key=lambda s: s[0])
            best = next((s for s in sizes if s[0] >= size), sizes[-1])
            im.size = best
            img = im.convert("RGBA").resize((size, size), Image.LANCZOS)
        return ImageTk.PhotoImage(img)
    except Exception:
        return None


def load_flag_image():
    png = resource_path("kuwait.png")
    if os.path.exists(png):
        try:
            return tk.PhotoImage(file=png)
        except tk.TclError:
            pass
    return None


def play_success_sound():
    if not winsound:
        return
    for name in ("success.wav", "Success.wav"):
        wav = resource_path(name)
        if os.path.exists(wav):
            try:
                winsound.PlaySound(wav, winsound.SND_FILENAME | winsound.SND_ASYNC)
                return
            except RuntimeError:
                break
    try:
        winsound.MessageBeep(winsound.MB_ICONASTERISK)
    except RuntimeError:
        pass


# ============================================================================
# Small UI building blocks
# ============================================================================

class ScrollFrame(ttk.Frame):
    """A frame whose content scrolls vertically when the window is too small."""

    def __init__(self, master, **kw):
        super().__init__(master, **kw)
        self.canvas = tk.Canvas(self, highlightthickness=0, bd=0)
        self.vbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.inner = ttk.Frame(self.canvas)
        self._win = self.canvas.create_window(0, 0, window=self.inner, anchor="nw")
        self.canvas.configure(yscrollcommand=self.vbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self._bar_shown = False
        self.inner.bind("<Configure>", lambda e: self._update())
        self.canvas.bind("<Configure>", self._on_canvas)
        self.canvas.bind("<Enter>", lambda e: self._bind_wheel(True))
        self.canvas.bind("<Leave>", lambda e: self._bind_wheel(False))

    def _on_canvas(self, event):
        self.canvas.itemconfigure(self._win, width=event.width)
        self._update()

    def _update(self):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        need = self.inner.winfo_reqheight() > self.canvas.winfo_height() + 1
        if need and not self._bar_shown:
            self.vbar.pack(side="right", fill="y", before=self.canvas)
            self._bar_shown = True
        elif not need and self._bar_shown:
            self.vbar.pack_forget()
            self._bar_shown = False
            self.canvas.yview_moveto(0)

    def _bind_wheel(self, on: bool):
        if on:
            self.canvas.bind_all("<MouseWheel>", self._on_wheel)
            self.canvas.bind_all("<Button-4>", lambda e: self._scroll(-1))
            self.canvas.bind_all("<Button-5>", lambda e: self._scroll(1))
        else:
            for seq in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
                self.canvas.unbind_all(seq)

    def _on_wheel(self, event):
        self._scroll(-1 if event.delta > 0 else 1)

    def _scroll(self, direction):
        if self._bar_shown:
            self.canvas.yview_scroll(direction * 3, "units")


# ============================================================================
# The app
# ============================================================================

class App(tk.Tk):
    def __init__(self, settings: dict):
        super().__init__()
        self.withdraw()

        self.settings = settings
        self.lang = settings["language"]
        self.admin = is_admin()
        self.dark = False
        self.pal = PALETTES["light"]
        self.sc = max(1.0, self.winfo_fpixels("1i") / 96.0)

        self.ui_queue = queue.Queue()
        self.runner = CommandRunner(self._on_cmd_line, self._on_cmd_live)
        self.worker_thread = None
        self.running = False
        self.closing = False
        self.run_started = 0.0
        self.total_steps = 0
        self.current_index = 0
        self.log_lines = []  # (text, tag): kept so the log survives a language switch
        self.status_key = ("idle", {})
        self.live_text = ""
        self.pct = 0.0
        self._last_system_dark = None

        # variables
        sel = settings["selected"]
        self.vars = {k: tk.BooleanVar(value=bool(sel.get(k, STEP_INFO[k]["default"]))) for k in STEP_INFO}
        self.var_select_all = tk.BooleanVar(value=False)
        self.var_chkdsk_mode = tk.StringVar(value=settings.get("chkdsk_mode", "scan"))
        self.var_drive = tk.StringVar(value=settings.get("drive", "C:"))
        self.var_always_admin = tk.BooleanVar(value=bool(settings.get("always_admin")))
        self.var_lang = tk.StringVar(value=self.lang)
        self.var_theme = tk.StringVar(value=settings["theme"])
        self.var_status = tk.StringVar()
        self.var_elapsed = tk.StringVar()
        self.var_live = tk.StringVar()
        self.var_count = tk.StringVar()
        self.var_problem = tk.StringVar()
        self.var_auto_check = tk.BooleanVar(value=bool(settings.get("auto_check", True)))
        self.var_restore_point = tk.BooleanVar(value=bool(settings.get("restore_point", True)))
        self.var_sound = tk.BooleanVar(value=bool(settings.get("sound", True)))
        self.var_pct = tk.StringVar(value="0%")
        self.current_tab = "diag"
        self.scroll_frames = []
        self.diag_lines = []      # (tag, text_key, kwargs) shown under "Find the problem"
        self.diag_status = None   # (text_key, kwargs)
        self.suggested = []       # step keys the last check / description suggested
        self.diag_running = False
        self._guard = False

        for v in self.vars.values():
            v.trace_add("write", lambda *_: self._on_option_changed())
        self.var_select_all.trace_add("write", lambda *_: self._on_select_all())
        self.vars["chkdsk"].trace_add("write", lambda *_: self.update_chkdsk_controls())

        self._init_fonts()
        self._set_icon()
        self.title(f"{APP_NAME} {APP_VERSION}")
        self._init_geometry()

        self.root_frame = None
        self.apply_theme()
        # sv-ttk finishes setting itself up on the first theme change: build once that's done
        self.update()
        self.retheme()

        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.after(60, self._pump_queue)
        self.after(1000, self._tick)
        self.after(3000, self._watch_system_theme)
        self.after(1200, lambda: self.check_for_update(show_if_latest=False))
        if self.var_auto_check.get():
            self.after(700, self.on_auto_check)

        self.deiconify()
        self.center_window()
        self.lift()
        self.focus_force()

    # ------------------------------------------------------------------ text
    def t(self, key: str, **kw) -> str:
        s = STRINGS[self.lang].get(key) or STRINGS["en"].get(key, key)
        return s.format(**kw) if kw else s

    @property
    def rtl(self) -> bool:
        return self.lang == "ar"

    def side(self, start=True):
        """'left' in English, 'right' in Arabic (start of the line)."""
        return ("right" if self.rtl else "left") if start else ("left" if self.rtl else "right")

    @property
    def anchor(self):
        return "e" if self.rtl else "w"

    def px(self, n) -> int:
        return int(round(n * self.sc))

    # ------------------------------------------------------------ setup bits
    def _init_fonts(self):
        """Segoe UI Bold for everything."""
        families = set(tkfont.families(self))
        ui = next((f for f in ("Segoe UI", "Noto Sans", "DejaVu Sans") if f in families), "TkDefaultFont")
        self.ui_family = ui
        B = "bold"
        self.f_body = tkfont.Font(self, family=ui, size=10, weight=B)
        self.f_small = tkfont.Font(self, family=ui, size=9, weight=B)
        self.f_option = tkfont.Font(self, family=ui, size=10, weight=B)
        self.f_heading = tkfont.Font(self, family=ui, size=13, weight=B)
        self.f_title = tkfont.Font(self, family=ui, size=18, weight=B)
        self.f_mono = tkfont.Font(self, family=ui, size=10, weight=B)
        self.f_mono_small = tkfont.Font(self, family=ui, size=9, weight=B)
        self.f_link = tkfont.Font(self, family=ui, size=10, weight=B, underline=True)
        self.f_pct = tkfont.Font(self, family=ui, size=12, weight=B)

    def _apply_named_fonts(self):
        """Also switch Tk's and sv-ttk's built-in fonts (tabs, entries, dialogs) to Segoe UI Bold."""
        for name in tkfont.names(self):
            if name.startswith(("Tk", "SunValley")):
                try:
                    tkfont.nametofont(name, root=self).configure(family=self.ui_family, weight="bold")
                except tk.TclError:
                    pass

    def _set_icon(self):
        ico = resource_path("icon.ico")
        self.icon_path = ico if os.path.exists(ico) else None
        if self.icon_path and IS_WINDOWS:
            try:
                self.iconbitmap(default=self.icon_path)  # also used by dialogs
            except tk.TclError:
                pass
        self.header_icon = load_icon_image(self.px(44))
        if self.header_icon and not IS_WINDOWS:
            try:
                self.iconphoto(True, self.header_icon)
            except tk.TclError:
                pass

    def _init_geometry(self):
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        w, h = self.px(1120), self.px(900)
        size = self.settings.get("size")
        if isinstance(size, list) and len(size) == 2:
            w, h = int(size[0]), int(size[1])
        w, h = min(w, sw - 40), min(h, sh - 80)
        self.geometry(f"{w}x{h}")
        self.minsize(min(self.px(820), sw - 40), min(self.px(560), sh - 80))

    def center_window(self):
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        x = max(0, (self.winfo_screenwidth() - w) // 2)
        y = max(0, (self.winfo_screenheight() - h) // 2 - self.px(20))
        self.geometry(f"{w}x{h}+{x}+{y}")

    def center_child(self, win):
        win.update_idletasks()
        x = self.winfo_rootx() + (self.winfo_width() - win.winfo_width()) // 2
        y = self.winfo_rooty() + (self.winfo_height() - win.winfo_height()) // 3
        win.geometry(f"+{max(0, x)}+{max(0, y)}")

    # ================================================================== UI
    def check_row(self, parent, var, text, label_style="Option.TLabel", cb_style="Option.TCheckbutton", command=None):
        """A checkbox with its text, mirrored (box on the right) in Arabic."""
        row = ttk.Frame(parent)
        if not self.rtl:
            cb = ttk.Checkbutton(row, text=text, variable=var, style=cb_style, command=command)
            cb.pack(side="left")
            return row, [cb]
        cb = ttk.Checkbutton(row, variable=var, style=cb_style, command=command)
        cb.pack(side="right")
        lbl = ttk.Label(row, text=text, style=label_style, cursor="hand2")
        lbl.pack(side="right")
        lbl.bind("<Button-1>", lambda e: cb.invoke() if cb.instate(["!disabled"]) else None)
        return row, [cb, lbl]

    def radio_row(self, parent, var, value, text, command=None):
        row = ttk.Frame(parent)
        if not self.rtl:
            rb = ttk.Radiobutton(row, text=text, value=value, variable=var, command=command)
            rb.pack(side="left")
            return row, [rb]
        rb = ttk.Radiobutton(row, value=value, variable=var, command=command)
        rb.pack(side="right")
        lbl = ttk.Label(row, text=text, cursor="hand2")
        lbl.pack(side="right")
        lbl.bind("<Button-1>", lambda e: rb.invoke() if rb.instate(["!disabled"]) else None)
        return row, [rb, lbl]

    def card(self, parent, title=None, sub=None, **pack):
        P, A = self.px, self.anchor
        c = ttk.Frame(parent, style="Card.TFrame", padding=(P(18), P(14), P(18), P(16)))
        c.pack(fill="x", **pack)
        if title:
            ttk.Label(c, text=title, style="Heading.TLabel").pack(anchor=A)
        if sub:
            ttk.Label(c, text=sub, style="Muted.TLabel", justify="right" if self.rtl else "left").pack(
                anchor=A, pady=(0, P(6)))
        return c

    def wrap_with(self, container, labels, margin=40):
        container.bind("<Configure>", lambda e: [l.configure(wraplength=max(200, e.width - self.px(margin)))
                                                 for l in labels if l.winfo_exists()], add="+")

    TABS = ("diag", "steps", "log", "settings", "about")

    def build_ui(self):
        """(Re)builds the window. Called again when the language or theme changes."""
        if self.root_frame is not None:
            self.root_frame.destroy()
        self.config(menu="")
        S, E, A, P = self.side(True), self.side(False), self.anchor, self.px
        self.scroll_frames = []

        root = ttk.Frame(self)
        root.pack(fill="both", expand=True)
        self.root_frame = root

        # ---------- header ----------
        header = ttk.Frame(root, padding=(P(18), P(14), P(18), P(6)))
        header.pack(fill="x")
        if self.header_icon:
            ttk.Label(header, image=self.header_icon).pack(side=S, padx=(0, P(12)) if not self.rtl else (P(12), 0))
        titles = ttk.Frame(header)
        titles.pack(side=S)
        ttk.Label(titles, text=APP_NAME, style="Title.TLabel").pack(anchor=A)
        ttk.Label(titles, text=self.t("tagline"), style="Muted.TLabel").pack(anchor=A)

        right = ttk.Frame(header)
        right.pack(side=E)
        self.btn_theme = ttk.Button(right, text="☀" if self.dark else "☾", width=3,
                                    command=self.on_quick_theme)
        self.btn_lang = ttk.Button(right, text=self.t("toggle_lang"), command=self.on_quick_language)
        self.btn_admin = ttk.Button(right, text=self.t("run_admin"), command=self.on_run_as_admin,
                                    style="Accent.TButton")
        self.lbl_admin = ttk.Label(right, text=self.t("admin_yes" if self.admin else "admin_no"),
                                   style="AdminOk.TLabel" if self.admin else "AdminNo.TLabel")
        for w in (self.btn_theme, self.btn_lang, self.btn_admin, self.lbl_admin):
            w.pack(side=E, padx=P(4))
        if self.admin:
            self.btn_admin.pack_forget()

        # ---------- footer: progress + Start/Skip/Cancel, always visible ----------
        footer = ttk.Frame(root, padding=(P(14), 0, P(14), P(14)))
        footer.pack(side="bottom", fill="x")
        prog = ttk.Frame(footer, style="Card.TFrame", padding=(P(16), P(10), P(16), P(12)))
        prog.pack(fill="x", pady=(P(8), P(8)))
        line = ttk.Frame(prog)
        line.pack(fill="x")
        ttk.Label(line, textvariable=self.var_status, style="Option.TLabel").pack(side=S)
        ttk.Label(line, textvariable=self.var_pct, style="Pct.TLabel").pack(side=E)
        ttk.Label(line, textvariable=self.var_elapsed, style="Muted.TLabel").pack(side=E, padx=P(12))
        self.progress = ttk.Progressbar(prog, orient="horizontal", mode="determinate", maximum=100)
        self.progress.pack(fill="x", pady=(P(8), P(4)))
        ttk.Label(prog, textvariable=self.var_live, style="Live.TLabel").pack(anchor="w")

        btns = ttk.Frame(footer)
        btns.pack(fill="x")
        self.btn_start = ttk.Button(btns, text=self.t("start"), style="Accent.TButton", command=self.on_start, width=14)
        self.btn_skip = ttk.Button(btns, text=self.t("skip"), command=self.on_skip)
        self.btn_cancel = ttk.Button(btns, text=self.t("cancel"), command=self.on_cancel)
        for w in (self.btn_start, self.btn_skip, self.btn_cancel):
            w.pack(side=S, padx=(0, P(8)) if not self.rtl else (P(8), 0))
        ttk.Label(btns, textvariable=self.var_count, style="Muted.TLabel").pack(side=E)

        # ---------- tabs ----------
        self.nb = ttk.Notebook(root)
        self.nb.pack(fill="both", expand=True, padx=P(14))
        self.tab_frames = {}
        builders = {"diag": self._build_diag_tab, "steps": self._build_steps_tab, "log": self._build_log_tab,
                    "settings": self._build_settings_tab, "about": self._build_about_tab}
        order = list(reversed(self.TABS)) if self.rtl else list(self.TABS)
        for key in order:
            frame = ttk.Frame(self.nb, padding=(0, P(10), 0, 0))
            self.tab_frames[key] = frame
            self.nb.add(frame, text=f"  {self.t('tab_' + key)}  ")
        for key in self.TABS:
            builders[key](self.tab_frames[key])
        self.nb.select(self.tab_frames[self.current_tab])
        self.nb.bind("<<NotebookTabChanged>>", self._on_tab_changed)

        # restore state into the new widgets
        self._replay_log()
        self.refresh_drive_list()
        self.update_chkdsk_controls()
        self._on_option_changed()
        self._set_status(*self.status_key)
        self.progress["value"] = self.pct
        self.var_pct.set(f"{int(self.pct)}%")
        if self.run_started:
            self._show_elapsed()
        self.set_running(self.running)

    def _on_tab_changed(self, _event=None):
        try:
            sel = self.nb.select()
            for key, frame in self.tab_frames.items():
                if str(frame) == sel:
                    self.current_tab = key
        except tk.TclError:
            pass

    def show_tab(self, key):
        self.current_tab = key
        try:
            self.nb.select(self.tab_frames[key])
        except (AttributeError, tk.TclError):
            pass

    def _scroll(self, parent):
        sf = ScrollFrame(parent)
        sf.pack(fill="both", expand=True)
        self.scroll_frames.append(sf)
        return sf.inner

    # ---------------------------------------------------------- diagnose
    def _build_diag_tab(self, tab):
        self._build_diag_card(self._scroll(tab))

    # ------------------------------------------------------------- steps
    def _build_steps_tab(self, tab):
        S, A, P = self.side(True), self.anchor, self.px
        opts = self._scroll(tab)
        self.badges = []

        bar = ttk.Frame(opts, padding=(P(4), 0, P(4), P(8)))
        bar.pack(fill="x")
        row, self.select_all_widgets = self.check_row(bar, self.var_select_all, self.t("select_all"))
        row.pack(side=S)

        cards = ttk.Frame(opts)
        cards.pack(fill="x")
        cards.columnconfigure(0, weight=1, uniform="c")
        cards.columnconfigure(1, weight=1, uniform="c")
        col_repair, col_cleanup = (1, 0) if self.rtl else (0, 1)

        self.option_widgets = {}
        for group, col in (("repair", col_repair), ("cleanup", col_cleanup)):
            card = ttk.Frame(cards, style="Card.TFrame", padding=(P(16), P(12), P(16), P(14)))
            card.grid(row=0, column=col, sticky="nsew",
                      padx=(0, P(6)) if col == 0 else (P(6), 0), pady=(0, P(4)))
            ttk.Label(card, text=self.t(group), style="Heading.TLabel").pack(anchor=A)
            ttk.Label(card, text=self.t(group + "_sub"), style="Muted.TLabel").pack(anchor=A, pady=(0, P(4)))
            card_descs = []
            for key, g, needs_admin, _ in STEPS:
                if g != group:
                    continue
                item = ttk.Frame(card)
                item.pack(fill="x", anchor=A, pady=(P(8), 0))
                top = ttk.Frame(item)
                top.pack(fill="x")
                row, ws = self.check_row(top, self.vars[key], self.t("opt_" + key))
                row.pack(side=S)
                if needs_admin and not self.admin:
                    badge = ttk.Label(top, text=self.t("needs_admin"), style="Badge.TLabel")
                    badge.pack(side=S, padx=P(8))
                    self.badges.append(badge)
                indent = (P(28), 0) if not self.rtl else (0, P(28))
                desc = ttk.Label(item, text=self.t("desc_" + key), style="Muted.TLabel",
                                 justify="right" if self.rtl else "left", wraplength=P(420))
                desc.pack(anchor=A, padx=indent)
                card_descs.append(desc)
                self.option_widgets[key] = ws
                if key == "chkdsk":
                    self._build_chkdsk_controls(item, indent)
            self.wrap_with(card, card_descs, margin=70)

    # --------------------------------------------------------------- log
    def _build_log_tab(self, tab):
        S, P = self.side(True), self.px
        btns = ttk.Frame(tab)
        btns.pack(fill="x", pady=(0, P(8)))
        self.btn_copy = ttk.Button(btns, text=self.t("copy_log"), command=self.on_copy_log)
        self.btn_save = ttk.Button(btns, text=self.t("save_log"), command=self.on_save_log)
        self.btn_clear = ttk.Button(btns, text=self.t("clear_log"), command=self.on_clear)
        for w in (self.btn_copy, self.btn_save, self.btn_clear):
            w.pack(side=S, padx=(0, P(8)) if not self.rtl else (P(8), 0))

        log_card = ttk.Frame(tab, style="Card.TFrame", padding=P(2))
        log_card.pack(fill="both", expand=True)
        self.txt = tk.Text(log_card, wrap="word", font=self.f_mono, relief="flat", bd=0,
                           padx=P(10), pady=P(8), undo=False, height=8)
        sb = ttk.Scrollbar(log_card, orient="vertical", command=self.txt.yview)
        sb.pack(side="right", fill="y")
        self.txt.pack(side="left", fill="both", expand=True)
        self.txt.configure(yscrollcommand=sb.set, state="disabled")
        self.txt.bind("<Control-a>", lambda e: (self.txt.tag_add("sel", "1.0", "end"), "break")[1])

    # ---------------------------------------------------------- settings
    def _build_settings_tab(self, tab):
        S, A, P = self.side(True), self.anchor, self.px
        page = self._scroll(tab)
        self.settings_lang_widgets = []

        look = self.card(page, self.t("set_appearance"), pady=(0, P(10)))
        ttk.Label(look, text=self.t("theme"), style="Option.TLabel").pack(anchor=A, pady=(P(6), P(2)))
        row = ttk.Frame(look)
        row.pack(anchor=A)
        for value, key in (("system", "theme_system"), ("light", "theme_light"), ("dark", "theme_dark")):
            r, _ = self.radio_row(row, self.var_theme, value, self.t(key), command=self.on_change_theme)
            r.pack(side=S, padx=(0, P(18)) if not self.rtl else (P(18), 0))
        ttk.Label(look, text=self.t("language"), style="Option.TLabel").pack(anchor=A, pady=(P(12), P(2)))
        row = ttk.Frame(look)
        row.pack(anchor=A)
        for value, text in (("en", "English"), ("ar", "العربية")):
            r, ws = self.radio_row(row, self.var_lang, value, text, command=self.on_change_language)
            r.pack(side=S, padx=(0, P(18)) if not self.rtl else (P(18), 0))
            self.settings_lang_widgets += ws

        opts = self.card(page, self.t("set_behavior"), pady=(0, P(10)))
        descs = []
        for var, key, cmd in ((self.var_always_admin, "always_admin", self.on_toggle_always_admin),
                              (self.var_restore_point, "restore_point", self.on_toggle_setting),
                              (self.var_auto_check, "auto_check_startup", self.on_toggle_setting),
                              (self.var_sound, "play_sound", self.on_toggle_setting)):
            item = ttk.Frame(opts)
            item.pack(fill="x", pady=(P(8), 0))
            r, _ = self.check_row(item, var, self.t(key), command=cmd)
            r.pack(anchor=A)
            d = ttk.Label(item, text=self.t(key + "_desc"), style="Muted.TLabel",
                          justify="right" if self.rtl else "left")
            d.pack(anchor=A, padx=(P(28), 0) if not self.rtl else (0, P(28)))
            descs.append(d)
        self.wrap_with(opts, descs, margin=70)

        upd = self.card(page, self.t("set_updates"), pady=(0, P(10)))
        row = ttk.Frame(upd)
        row.pack(fill="x", pady=(P(6), 0))
        ttk.Label(row, text=self.t("current_version", v=APP_VERSION)).pack(side=S)
        ttk.Button(row, text=self.t("check_updates"),
                   command=lambda: self.check_for_update(show_if_latest=True)).pack(side=S, padx=P(14))

    # ------------------------------------------------------------- about
    def _build_about_tab(self, tab):
        S, P = self.side(True), self.px
        page = self._scroll(tab)
        wrap = ttk.Frame(page)
        wrap.pack(pady=(P(10), 0))

        hero = ttk.Frame(wrap)
        hero.pack(pady=(0, P(14)))
        self.about_icon = load_icon_image(P(96))
        if self.about_icon:
            ttk.Label(hero, image=self.about_icon).pack()
        ttk.Label(hero, text=APP_NAME, style="Title.TLabel").pack(pady=(P(8), 0))
        ttk.Label(hero, text=self.t("version", v=APP_VERSION, d=build_date()), style="Badge.TLabel").pack(pady=(P(2), P(6)))
        ttk.Label(hero, text=self.t("about_sub"), justify="center").pack()

        cards = ttk.Frame(wrap)
        cards.pack(fill="x")
        cards.columnconfigure(0, weight=1, uniform="a", minsize=P(330))
        cards.columnconfigure(1, weight=1, uniform="a", minsize=P(330))
        c_made, c_links = (1, 0) if self.rtl else (0, 1)

        made = ttk.Frame(cards, style="Card.TFrame", padding=(P(18), P(14)))
        made.grid(row=0, column=c_made, sticky="nsew", padx=P(6))
        ttk.Label(made, text=self.t("about_made_by"), style="Heading.TLabel").pack(anchor=self.anchor)
        row = ttk.Frame(made)
        row.pack(anchor=self.anchor, pady=(P(8), P(4)))
        ttk.Label(row, text="ilukezippo (BoYaqoub)", style="Option.TLabel").pack(side=S)
        self.about_flag = load_flag_image()
        if self.about_flag:
            ttk.Label(row, image=self.about_flag).pack(side=S, padx=P(8))
        ttk.Label(made, text=self.t("made_in"), style="Muted.TLabel").pack(anchor=self.anchor)
        ttk.Label(made, text=self.t("about_free"), style="Muted.TLabel").pack(anchor=self.anchor, pady=(P(2), 0))

        links = ttk.Frame(cards, style="Card.TFrame", padding=(P(18), P(14)))
        links.grid(row=0, column=c_links, sticky="nsew", padx=P(6))
        ttk.Label(links, text=self.t("about_links"), style="Heading.TLabel").pack(anchor=self.anchor)
        ttk.Button(links, text=self.t("more_apps_btn"), style="Accent.TButton",
                   command=lambda: webbrowser.open(MORE_APPS_PAGE)).pack(fill="x", pady=(P(8), P(2)))
        site = ttk.Label(links, text=MORE_APPS_PAGE.replace("https://", "").rstrip("/"), style="Link.TLabel",
                         cursor="hand2")
        site.pack(pady=(0, P(6)))
        site.bind("<Button-1>", lambda e: webbrowser.open(MORE_APPS_PAGE))
        for key, cmd in (("github", lambda: webbrowser.open(GITHUB_PAGE)),
                         ("report", lambda: webbrowser.open(GITHUB_PAGE + "/issues")),
                         ("check_updates", lambda: self.check_for_update(show_if_latest=True))):
            ttk.Button(links, text=self.t(key), command=cmd).pack(fill="x", pady=P(2))

        donate_box = ttk.Frame(wrap)
        donate_box.pack(pady=(P(16), P(10)))
        ttk.Label(donate_box, text=self.t("donate_note"), style="Muted.TLabel").pack()
        self.donate_img = make_donate_image(P(170), P(44))
        self.btn_donate = tk.Button(donate_box, image=self.donate_img, text=self.t("donate"), compound="center",
                                    font=self.f_option, fg="#0f3462", activeforeground="#0f3462", bd=0,
                                    highlightthickness=0, cursor="hand2", relief="flat",
                                    command=lambda: webbrowser.open(DONATE_PAGE))
        self.btn_donate.pack(pady=(P(8), 0))

    def on_toggle_setting(self):
        self.settings["restore_point"] = bool(self.var_restore_point.get())
        self.settings["auto_check"] = bool(self.var_auto_check.get())
        self.settings["sound"] = bool(self.var_sound.get())
        save_settings(self.settings)

    def _build_diag_card(self, parent):
        S, E, A, P = self.side(True), self.side(False), self.anchor, self.px
        card = ttk.Frame(parent, style="Card.TFrame", padding=(P(16), P(12), P(16), P(12)))
        card.pack(fill="x")
        ttk.Label(card, text=self.t("diag_title"), style="Heading.TLabel").pack(anchor=A)
        ttk.Label(card, text=self.t("diag_sub"), style="Muted.TLabel").pack(anchor=A, pady=(0, P(8)))

        row = ttk.Frame(card)
        row.pack(fill="x")
        self.btn_check = ttk.Button(row, text=self.t("auto_check"), command=self.on_auto_check)
        self.btn_check.pack(side=E, padx=(P(8), 0) if not self.rtl else (0, P(8)))
        self.btn_suggest = ttk.Button(row, text=self.t("suggest"), command=self.on_suggest)
        self.btn_suggest.pack(side=E, padx=(P(8), 0) if not self.rtl else (0, P(8)))
        self.entry_problem = ttk.Entry(row, textvariable=self.var_problem, font=self.f_body,
                                       justify="right" if self.rtl else "left")
        self.entry_problem.pack(side=S, fill="x", expand=True)
        self.entry_problem.bind("<Return>", lambda e: self.on_suggest())
        ttk.Label(card, text=self.t("describe_hint"), style="Muted.TLabel").pack(anchor=A, pady=(P(3), 0))

        self.diag_box = ttk.Frame(card)
        self.diag_box.pack(fill="x", pady=(P(6), 0))
        card.bind("<Configure>", lambda e: self._wrap_diag(e.width - self.px(40)))
        self._render_diag()

    def _wrap_diag(self, width):
        self._diag_wrap = max(200, width)
        for w in self.diag_box.winfo_children():
            if isinstance(w, ttk.Label):
                w.configure(wraplength=self._diag_wrap)

    def _render_diag(self):
        box = getattr(self, "diag_box", None)
        if box is None or not box.winfo_exists():
            return
        for w in box.winfo_children():
            w.destroy()
        A, P = self.anchor, self.px
        wrap = getattr(self, "_diag_wrap", P(800))
        just = "right" if self.rtl else "left"
        if self.diag_status:
            ttk.Label(box, text=self.t(self.diag_status[0], **self.diag_status[1]), style="Option.TLabel",
                      wraplength=wrap, justify=just).pack(anchor=A, pady=(P(2), P(2)))
        styles = {"ok": "DiagOk.TLabel", "warn": "DiagWarn.TLabel", "err": "DiagErr.TLabel", "info": "Muted.TLabel"}
        marks = {"ok": "\u2713", "warn": "!", "err": "\u2715", "info": "i",
                 "tip": "\u2190" if self.rtl else "\u2192"}
        for tag, key, kw in self.diag_lines:
            text = f"{marks.get(tag, '-')}  {self.t(key, **kw)}"
            if self.rtl:
                text = f"{self.t(key, **kw)}  {marks.get(tag, '-')}"
            ttk.Label(box, text=text, style=styles.get(tag, "TLabel"), wraplength=wrap,
                      justify=just).pack(anchor=A, pady=1)
        if self.suggested:
            names = "، ".join(self.t("opt_" + k) for k in RUN_ORDER if k in self.suggested) if self.rtl else \
                ", ".join(self.t("opt_" + k) for k in RUN_ORDER if k in self.suggested)
            ttk.Label(box, text=self.t("suggested", steps=names), style="TLabel", wraplength=wrap,
                      justify=just).pack(anchor=A, pady=(P(6), P(4)))
            self.btn_use = ttk.Button(box, text=self.t("use_suggested"), style="Accent.TButton",
                                      command=self.on_use_suggested)
            self.btn_use.pack(anchor=A)
            if self.running:
                self.btn_use.state(["disabled"])
        if hasattr(self, "btn_check"):
            self.btn_check.state(["disabled"] if self.diag_running else ["!disabled"])

    def on_suggest(self):
        text = self.var_problem.get().strip()
        if not text:
            self.diag_status = ("empty_desc", {})
            self.diag_lines, self.suggested = [], []
            self._render_diag()
            self.entry_problem.focus_set()
            return
        matches = suggest_from_text(text)
        lines, steps = [], []
        if not matches:
            self.diag_status = ("no_match", {})
            matches = [("r_general", ["dism_restore", "sfc", "temp"])]
        else:
            self.diag_status = None
        for reason, keys in matches:
            lines.append(("tip", reason, {}))
            steps += [k for k in keys if k not in steps]
        self.diag_lines = lines
        self.suggested = steps
        self._render_diag()

    def on_auto_check(self):
        if self.diag_running:
            return
        self.diag_running = True
        self.diag_status = ("checking", {})
        self.diag_lines, self.suggested = [], []
        self._render_diag()
        admin = self.admin

        def work():
            try:
                findings = run_problem_check(admin)
            except Exception as e:
                findings = [("err", "log_error", {"err": e}, [])]

            def done():
                self.diag_running = False
                problems = [f for f in findings if f[0] in ("warn", "err")]
                self.diag_status = ("check_n", {"n": len(problems)}) if problems else ("check_none", {})
                order = {"err": 0, "warn": 1, "ok": 2, "info": 3}
                self.diag_lines = [(tag, key, kw) for tag, key, kw, _ in sorted(findings, key=lambda f: order[f[0]])]
                steps = []
                for _, _, _, keys in problems:
                    steps += [k for k in keys if k not in steps]
                self.suggested = steps
                self._render_diag()
            self.call_ui(done)
        threading.Thread(target=work, daemon=True).start()

    def on_use_suggested(self):
        if self.running or not self.suggested:
            return
        self._guard = True
        try:
            for k, v in self.vars.items():
                v.set(k in self.suggested)
        finally:
            self._guard = False
        self._on_option_changed()
        self.show_tab("steps")

    def on_toggle_auto_check(self):
        self.settings["auto_check"] = bool(self.var_auto_check.get())
        save_settings(self.settings)

    def _build_chkdsk_controls(self, item, indent):
        S, P = self.side(True), self.px
        sub = ttk.Frame(item)
        sub.pack(anchor=self.anchor, padx=indent, pady=(P(6), 0))
        self.lbl_drive = ttk.Label(sub, text=self.t("drive"))
        self.lbl_drive.pack(side=S)
        self.drive_combo = ttk.Combobox(sub, width=6, textvariable=self.var_drive, state="readonly")
        self.drive_combo.pack(side=S, padx=P(6))
        self.btn_drive_refresh = ttk.Button(sub, text=self.t("refresh"), command=self.refresh_drive_list)
        self.btn_drive_refresh.pack(side=S, padx=P(4))

        mode = ttk.Frame(item)
        mode.pack(anchor=self.anchor, padx=indent, pady=(P(6), 0))
        self.lbl_mode = ttk.Label(mode, text=self.t("mode"))
        self.lbl_mode.pack(side=S)
        r1, w1 = self.radio_row(mode, self.var_chkdsk_mode, "scan", self.t("scan_only"))
        r1.pack(side=S, padx=P(8))
        r2, w2 = self.radio_row(mode, self.var_chkdsk_mode, "fix", self.t("fix_f"))
        r2.pack(side=S, padx=P(8))
        self.chkdsk_widgets = [self.lbl_drive, self.drive_combo, self.btn_drive_refresh, self.lbl_mode] + w1 + w2

    # ================================================================ theme
    def apply_theme(self):
        mode = self.settings["theme"]
        dark = system_prefers_dark() if mode == "system" else mode == "dark"
        self._last_system_dark = system_prefers_dark()
        self.dark = dark
        pal = dict(PALETTES["dark" if dark else "light"])
        style = ttk.Style(self)

        if sv_ttk:
            sv_ttk.set_theme("dark" if dark else "light", self)
            # sv-ttk recolours on <<ThemeChanged>>; let that run before our styles go on top
            self.update()
        else:
            self._fallback_theme(style, pal)

        self._apply_named_fonts()
        style.configure("TNotebook", tabposition="ne" if self.rtl else "nw")
        style.configure("TNotebook.Tab", font=self.f_body)
        style.configure("Pct.TLabel", font=self.f_pct, foreground=pal["head"])
        bg = style.lookup("TFrame", "background") or pal["bg"]
        pal["bg"] = bg
        self.pal = pal
        self.configure(bg=bg)

        style.configure("TLabel", font=self.f_body)
        style.configure("TCheckbutton", font=self.f_body)
        style.configure("TRadiobutton", font=self.f_body)
        style.configure("TButton", font=self.f_body)
        style.configure("Muted.TLabel", foreground=pal["muted"], font=self.f_small)
        style.configure("Title.TLabel", font=self.f_title)
        style.configure("Heading.TLabel", font=self.f_heading)
        style.configure("Option.TLabel", font=self.f_option)
        style.configure("Option.TCheckbutton", font=self.f_option)
        style.configure("AdminOk.TLabel", foreground=pal["ok"], font=self.f_small)
        style.configure("AdminNo.TLabel", foreground=pal["warn"], font=self.f_small)
        style.configure("Badge.TLabel", foreground=pal["warn"], font=self.f_small)
        style.configure("Live.TLabel", foreground=pal["muted"], font=self.f_mono_small)
        style.configure("Link.TLabel", foreground=pal["link"], font=self.f_link)
        style.configure("DiagOk.TLabel", foreground=pal["ok"], font=self.f_body)
        style.configure("DiagWarn.TLabel", foreground=pal["warn"], font=self.f_body)
        style.configure("DiagErr.TLabel", foreground=pal["err"], font=self.f_body)
        style.map("Muted.TLabel", foreground=[("disabled", pal["border"])])

        self._apply_theme_to_tk_widgets()
        set_title_bar_dark(self, dark)
        for w in self.winfo_children():
            if isinstance(w, tk.Toplevel):
                set_title_bar_dark(w, dark)
                if hasattr(w, "_retheme"):
                    w._retheme()

    def _apply_theme_to_tk_widgets(self):
        pal = self.pal
        try:
            self.btn_theme.configure(text="\u2600" if self.dark else "\u263E")
            for sf in self.scroll_frames:
                if sf.winfo_exists():
                    sf.canvas.configure(bg=pal["bg"])
            if getattr(self, "btn_donate", None) is not None and self.btn_donate.winfo_exists():
                self.btn_donate.configure(bg=pal["bg"], activebackground=pal["bg"])
            self.txt.configure(bg=pal["log_bg"], fg=pal["log_fg"], insertbackground=pal["log_fg"],
                               selectbackground=pal["select"], selectforeground=pal["log_fg"],
                               highlightthickness=0)
            for tag in ("ok", "warn", "err", "head", "info"):
                self.txt.tag_configure(tag, foreground=pal[tag])
            self.txt.tag_configure("head", font=self.f_mono)
            self.txt.tag_configure("title", foreground=pal["head"], font=(self.f_mono.actual("family"), 10, "bold"))
        except (AttributeError, tk.TclError):
            pass

    def _fallback_theme(self, style, pal):
        """Plain 'clam' theme coloured to match, if sv-ttk isn't installed."""
        style.theme_use("clam")
        bg, card, fg, border = pal["bg"], pal["card"], pal["text"], pal["border"]
        accent = pal["head"]
        style.configure(".", background=bg, foreground=fg, fieldbackground=card, bordercolor=border,
                        lightcolor=bg, darkcolor=bg, troughcolor=card, focuscolor=accent)
        style.configure("Card.TFrame", background=bg, relief="solid", borderwidth=1)
        style.configure("TButton", background=card, padding=(10, 4))
        style.map("TButton", background=[("active", border)], foreground=[("disabled", pal["muted"])])
        style.configure("Accent.TButton", background=accent, foreground="#ffffff" if not self.dark else "#000000")
        style.map("Accent.TButton", background=[("disabled", border), ("active", accent)])
        style.configure("TCheckbutton", indicatorbackground=card, indicatorforeground=fg)
        style.configure("TRadiobutton", indicatorbackground=card, indicatorforeground=fg)
        style.map("TCheckbutton", background=[("active", bg)], foreground=[("disabled", pal["muted"])])
        style.map("TRadiobutton", background=[("active", bg)], foreground=[("disabled", pal["muted"])])
        style.map("TLabel", foreground=[("disabled", pal["muted"])])
        style.configure("Horizontal.TProgressbar", background=accent, troughcolor=card)
        style.configure("TCombobox", fieldbackground=card, background=card, foreground=fg)
        self.option_add("*TCombobox*Listbox.background", card)
        self.option_add("*TCombobox*Listbox.foreground", fg)

    def _watch_system_theme(self):
        if not self.closing:
            if self.settings["theme"] == "system":
                now = system_prefers_dark()
                if now != self._last_system_dark:
                    self.retheme()
            self.after(3000, self._watch_system_theme)

    def retheme(self):
        self.apply_theme()
        self.build_ui()  # ttk doesn't always repaint custom styles on existing widgets
        self._apply_theme_to_tk_widgets()

    def on_change_theme(self):
        self.settings["theme"] = self.var_theme.get()
        save_settings(self.settings)
        self.retheme()

    def on_quick_theme(self):
        self.var_theme.set("light" if self.dark else "dark")
        self.on_change_theme()

    # ============================================================= language
    def on_change_language(self):
        new = self.var_lang.get()
        if new == self.lang:
            return
        if self.running:
            self.var_lang.set(self.lang)
            return
        self.lang = new
        self.settings["language"] = new
        save_settings(self.settings)
        self.retheme()

    def on_quick_language(self):
        self.var_lang.set("en" if self.lang == "ar" else "ar")
        self.on_change_language()

    # ============================================================ selection
    def _on_option_changed(self):
        n = sum(1 for v in self.vars.values() if v.get())
        self.var_count.set(self.t("selected_count", n=n, total=len(self.vars)))
        if self._guard:
            return
        self._guard = True
        try:
            self.var_select_all.set(n == len(self.vars))
        finally:
            self._guard = False

    def _on_select_all(self):
        if self._guard:
            return
        self._guard = True
        try:
            val = bool(self.var_select_all.get())
            for v in self.vars.values():
                v.set(val)
        finally:
            self._guard = False
        self._on_option_changed()

    def update_chkdsk_controls(self):
        enabled = bool(self.vars["chkdsk"].get()) and not self.running
        for w in getattr(self, "chkdsk_widgets", []):
            try:
                if w is self.drive_combo:
                    w.configure(state="readonly" if enabled else "disabled")
                else:
                    w.state(["!disabled"] if enabled else ["disabled"])
            except tk.TclError:
                pass

    def refresh_drive_list(self):
        drives = list_drives()
        self.drive_combo["values"] = drives
        if self.var_drive.get() not in drives:
            self.var_drive.set(system_drive() if system_drive() in drives else drives[0])

    def _remember_choices(self):
        self.settings["selected"] = {k: bool(v.get()) for k, v in self.vars.items()}
        self.settings["chkdsk_mode"] = self.var_chkdsk_mode.get()
        self.settings["drive"] = self.var_drive.get()
        try:
            if self.state() == "normal":
                self.settings["size"] = [self.winfo_width(), self.winfo_height()]
        except tk.TclError:
            pass
        save_settings(self.settings)

    # ================================================================ admin
    def on_toggle_always_admin(self):
        self.settings["always_admin"] = bool(self.var_always_admin.get())
        save_settings(self.settings)
        if self.var_always_admin.get() and not self.admin and not self.running:
            self.on_run_as_admin()

    def on_run_as_admin(self):
        if self.admin or self.running:
            return
        self._remember_choices()
        if relaunch_as_admin():
            self.closing = True
            self.destroy()
        else:
            messagebox.showinfo(APP_NAME, self.t("elevate_failed"), parent=self)

    # ============================================================== the log
    @staticmethod
    def _auto_tag(text: str):
        if text.startswith("> "):
            return "head"
        return None

    def log(self, text: str, tag=None):
        """Thread-safe: queue a line for the log."""
        self.ui_queue.put(("log", text, tag))

    def _on_cmd_line(self, text, tag=None):
        self.ui_queue.put(("log", text, tag))

    def _on_cmd_live(self, text):
        self.ui_queue.put(("live", text, parse_percent(text)))

    def _append_lines(self, items):
        self.txt.configure(state="normal")
        for text, tag in items:
            self.txt.insert("end", text + "\n", (tag,) if tag else ())
        lines = int(self.txt.index("end-1c").split(".")[0])
        if lines > MAX_LOG_LINES:
            self.txt.delete("1.0", f"{lines - MAX_LOG_LINES}.0")
        self.txt.configure(state="disabled")
        self.txt.see("end")

    def _replay_log(self):
        if self.log_lines:
            self._append_lines(self.log_lines)

    def on_clear(self):
        self.log_lines.clear()
        self.txt.configure(state="normal")
        self.txt.delete("1.0", "end")
        self.txt.configure(state="disabled")

    def _log_text(self) -> str:
        return "\n".join(t for t, _ in self.log_lines)

    def on_copy_log(self):
        self.clipboard_clear()
        self.clipboard_append(self._log_text())
        self._set_status("log_copied", {}, transient=True)

    def on_save_log(self):
        name = f"WindowsFixer_log_{datetime.now():%Y-%m-%d_%H-%M}.txt"
        path = filedialog.asksaveasfilename(parent=self, defaultextension=".txt", initialfile=name,
                                            filetypes=[("Text", "*.txt"), ("*", "*.*")])
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8-sig") as f:
                f.write(self._log_text() + "\n")
            self._set_status("log_saved", {"path": os.path.basename(path)}, transient=True)
        except OSError as e:
            messagebox.showerror(APP_NAME, self.t("save_failed", err=e), parent=self)

    # ============================================================ UI queue
    def call_ui(self, fn):
        """Run fn on the Tk thread (Tkinter isn't thread-safe)."""
        self.ui_queue.put(("call", fn))

    def _pump_queue(self):
        batch = []
        try:
            for _ in range(2000):
                item = self.ui_queue.get_nowait()
                kind = item[0]
                if kind == "log":
                    text, tag = item[1], item[2] or self._auto_tag(item[1])
                    self.log_lines.append((text, tag))
                    batch.append((text, tag))
                elif kind == "live":
                    self.var_live.set(item[1][:160])
                    if item[2] is not None:
                        self._set_step_fraction(item[2] / 100.0)
                elif kind == "call":
                    if batch:
                        self._append_lines(batch)
                        batch = []
                    item[1]()
        except queue.Empty:
            pass
        if batch:
            self._append_lines(batch)
        if len(self.log_lines) > MAX_LOG_LINES:
            del self.log_lines[: len(self.log_lines) - MAX_LOG_LINES]
        if not self.closing:
            self.after(60, self._pump_queue)

    # ============================================================= progress
    def _set_status(self, key, kw, transient=False):
        if not transient:
            self.status_key = (key, kw)
        self.var_status.set(self.t(key, **kw))
        if transient:
            self.after(2500, lambda: self.var_status.set(self.t(self.status_key[0], **self.status_key[1])))

    def _set_step_fraction(self, frac):
        if self.total_steps <= 0:
            return
        frac = min(max(frac, 0.0), 1.0)
        self.pct = ((self.current_index - 1) + frac) / self.total_steps * 100
        self.progress["value"] = self.pct
        self.var_pct.set(f"{int(self.pct)}%")
        self.title(f"{int(self.pct)}% \u2022 {APP_NAME}")  # shows on the taskbar too

    def _show_elapsed(self):
        secs = int(time.time() - self.run_started)
        self.var_elapsed.set(self.t("elapsed", t=f"{secs // 3600}:{secs // 60 % 60:02d}:{secs % 60:02d}"))

    def _tick(self):
        if self.running:
            self._show_elapsed()
        if not self.closing:
            self.after(1000, self._tick)

    # =========================================================== run state
    def set_running(self, running: bool):
        self.running = running
        if not running:
            self.title(f"{APP_NAME} {APP_VERSION}")
        try:
            self.btn_start.state(["disabled"] if running else ["!disabled"])
            self.btn_skip.state(["!disabled"] if running else ["disabled"])
            self.btn_cancel.state(["!disabled"] if running else ["disabled"])
            self.btn_lang.state(["disabled"] if running else ["!disabled"])
            self.btn_admin.state(["disabled"] if running else ["!disabled"])
            for ws in self.option_widgets.values():
                for w in ws:
                    w.state(["disabled"] if running else ["!disabled"])
            for w in self.select_all_widgets + self.settings_lang_widgets:
                w.state(["disabled"] if running else ["!disabled"])
            if getattr(self, "btn_use", None) is not None and self.btn_use.winfo_exists():
                self.btn_use.state(["disabled"] if running else ["!disabled"])
        except (AttributeError, tk.TclError):
            pass
        self.update_chkdsk_controls()

    def step_name(self, key):
        if key == "chkdsk":
            return self.t("name_chkdsk", drive=self.var_drive.get())
        return self.t("name_" + key)

    def on_start(self):
        if self.running:
            return
        keys = [k for k in RUN_ORDER if self.vars[k].get()]
        if not keys:
            messagebox.showwarning(self.t("nothing_t"), self.t("nothing"), parent=self)
            return
        self._remember_choices()

        if not self.admin:
            admin_keys = [k for k in keys if STEP_INFO[k]["admin"]]
            if admin_keys:
                listing = "\n".join("\u2022 " + self.t("opt_" + k) for k in admin_keys)
                answer = messagebox.askyesnocancel(self.t("need_admin_t"), self.t("need_admin", steps=listing),
                                                   parent=self, icon="warning")
                if answer is None:
                    return
                if answer:
                    self.on_run_as_admin()
                    return

        # everything the worker needs, read on the UI thread
        plan = {
            "keys": keys,
            "names": {k: self.step_name(k) for k in keys},
            "drive": self.var_drive.get().strip().upper(),
            "chkdsk_mode": self.var_chkdsk_mode.get(),
            # a safety net before anything that changes Windows itself
            "restore_point": bool(self.var_restore_point.get()) and self.admin and any(
                k in ("dism_restore", "sfc", "comp", "reset_net", "chkdsk", "wu") for k in keys),
        }
        self.runner.reset()
        self.total_steps = len(keys)
        self.current_index = 0
        self.pct = 0
        self.progress["value"] = 0
        self.var_pct.set("0%")
        self.show_tab("log")
        self.var_live.set("")
        self.run_started = time.time()
        self.var_elapsed.set(self.t("elapsed", t="0:00:00"))
        self._set_status("starting", {})
        self.set_running(True)

        if self.log_lines:
            self.log("")
        self.log(self.t("log_header", app=APP_NAME, v=APP_VERSION, when=f"{datetime.now():%Y-%m-%d %H:%M}"), "title")
        if not self.admin:
            self.log(self.t("log_admin_warn"), "warn")

        self.worker_thread = threading.Thread(target=self.worker, args=(plan,), daemon=True)
        self.worker_thread.start()

    def on_skip(self):
        if self.running and not self.runner.cancel_all:
            self.log(self.t("log_skip_req"), "info")
            threading.Thread(target=self.runner.request_skip, daemon=True).start()

    def on_cancel(self):
        if self.running and not self.runner.cancel_all:
            self.log(self.t("log_cancel_req"), "info")
            threading.Thread(target=self.runner.request_cancel, daemon=True).start()

    # ============================================================== worker
    def worker(self, plan):
        results = {}
        keys = plan["keys"]
        restart_for = []
        try:
            if plan["restore_point"]:
                self.log(self.t("log_rp"), "title")
                self.call_ui(lambda: self._set_status("log_rp", {}))
                status, _ = self._run(["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
                                       "-Command", "Checkpoint-Computer -Description 'Windows Fixer' "
                                                   "-RestorePointType MODIFY_SETTINGS -ErrorAction Stop"])
                if status == "ok":
                    self.log(self.t("log_rp_ok"), "ok")
                elif status != "cancel":
                    self.log(self.t("log_rp_fail"), "warn")
            for idx, key in enumerate(keys, start=1):
                if self.runner.cancel_all:
                    break
                self.runner.new_step()
                name = plan["names"][key]

                def begin(i=idx, n=name):
                    self.current_index = i
                    self.var_live.set("")
                    self._set_step_fraction(0)
                    self._set_status("step_of", {"i": i, "n": self.total_steps, "name": n})
                self.call_ui(begin)
                self.log("")
                self.log(self.t("log_step", i=idx, n=len(keys), name=name), "title")

                status, code = getattr(self, "step_" + key)(plan)
                results[key] = (status, code)
                if status == "error":
                    self.log(self.t("log_start_fail", err=code), "err")
                elif status == "failed":
                    self.log(self.t("log_exit_code", code=format_code(code)), "err")
                    if code in (5, 740) or (key == "reset_net" and code == 1):
                        self.log(self.t("log_hint_admin"), "warn")
                    if isinstance(code, int) and (code & 0xFFFFFFFF) in DISM_SOURCE_ERRORS:
                        self.log(self.t("log_hint_source"), "warn")
                elif status == "issues":
                    self.log(self.t("log_exit_code", code=format_code(code)), "warn")
                if status in ("ok", "issues") and (key == "reset_net" or
                                                   (key == "chkdsk" and plan["chkdsk_mode"] == "fix")):
                    restart_for.append(name)
                self.call_ui(lambda: self._set_step_fraction(1.0))
                if status == "cancel":
                    break

            # summary
            self.log("")
            self.log(self.t("log_summary"), "title")
            had_problems = False
            for key in keys:
                status, code = results.get(key, ("not_run", None))
                if status == "cancel":
                    status = "cancelled"
                text_key = {"ok": "res_ok", "issues": "res_issues", "failed": "res_failed", "error": "res_failed",
                            "skip": "res_skipped", "cancelled": "res_cancelled", "not_run": "res_not_run"}[status]
                tag = {"ok": "ok", "issues": "warn", "failed": "err", "error": "err"}.get(status, "info")
                if status in ("failed", "error", "issues"):
                    had_problems = True
                label = self.t(text_key, code=format_code(code) if isinstance(code, int) else "?")
                self.log(f"  {plan['names'][key]}: {label}", tag)
            if restart_for:
                self.log(self.t("restart_needed", steps=", ".join(restart_for)), "warn")

            cancelled = self.runner.cancel_all

            def finish():
                self.var_live.set("")
                self._show_elapsed()
                if cancelled:
                    self._set_status("cancelled", {})
                else:
                    self.progress["value"] = self.pct = 100
                    self.var_pct.set("100%")
                    self._set_status("done_issues" if had_problems else "done", {})
                    self.log(self.t("log_all_done"), "ok" if not had_problems else "warn")
                    if self.var_sound.get():
                        play_success_sound()
                    self._flash_taskbar()
            self.call_ui(finish)
        except Exception as e:  # never leave the UI stuck in "running"
            self.log(self.t("log_error", err=e), "err")
        finally:
            self.call_ui(lambda: self.set_running(False))

    def _flash_taskbar(self):
        if not IS_WINDOWS:
            return
        try:
            if self.focus_displayof() is None:
                hwnd = ctypes.windll.user32.GetParent(self.winfo_id())
                ctypes.windll.user32.FlashWindow(hwnd, True)
        except Exception:
            pass

    # -------------------------------------------------------------- steps
    def _run(self, cmd, **kw):
        return self.runner.run(cmd, **kw)

    def _clean(self, folders):
        total_freed = total_deleted = total_left = 0
        for folder in folders:
            if self.runner.should_abort():
                break
            if not folder or not os.path.isdir(folder):
                self.log(self.t("log_not_found", path=folder), "info")
                continue
            self.log(self.t("log_cleaning", path=folder))
            self.ui_queue.put(("live", folder, None))
            freed, deleted, left, aborted = clean_folder_contents(folder, self.runner.should_abort)
            total_freed += freed
            total_deleted += deleted
            total_left += left
            if aborted:
                self.log(self.t("log_aborted"), "info")
                break
        self.log(self.t("log_cleaned", size=human_size(total_freed), count=total_deleted), "ok")
        if total_left:
            self.log(self.t("log_in_use", count=total_left), "info")
        if self.runner.cancel_all:
            return "cancel", None
        if self.runner.skip_step:
            return "skip", None
        return "ok", 0

    def step_temp(self, plan):
        folders = []
        user_temp = os.environ.get("TEMP") or os.path.join(os.environ.get("LOCALAPPDATA", ""), "Temp")
        win_temp = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Temp")
        for f in (user_temp, win_temp):
            if f and os.path.normcase(os.path.abspath(f)) not in [os.path.normcase(os.path.abspath(x)) for x in folders]:
                folders.append(f)
        return self._clean(folders)

    def step_prefetch(self, plan):
        if not self.admin:
            self.log(self.t("log_needs_admin_skip"), "warn")
            return "skip", None
        return self._clean([os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Prefetch")])

    def step_recycle(self, plan):
        try:
            before = query_recycle_bin()
            if before and before[1] == 0:
                self.log(self.t("log_rb_empty"), "ok")
                return "ok", 0
            SHERB_NOCONFIRMATION, SHERB_NOPROGRESSUI, SHERB_NOSOUND = 0x1, 0x2, 0x4
            hr = ctypes.windll.shell32.SHEmptyRecycleBinW(None, None,
                                                          SHERB_NOCONFIRMATION | SHERB_NOPROGRESSUI | SHERB_NOSOUND)
            if hr == 0:
                size, count = before if before else (0, 0)
                self.log(self.t("log_rb_ok", size=human_size(size), count=count), "ok")
                return "ok", 0
            if (hr & 0xFFFFFFFF) == 0x8000FFFF:  # E_UNEXPECTED: already empty
                self.log(self.t("log_rb_empty"), "ok")
                return "ok", 0
            self.log(self.t("log_rb_fail", code=format_code(hr)), "err")
            return "failed", hr
        except Exception as e:
            return "error", str(e)

    def step_dns(self, plan):
        return self._run(["ipconfig", "/flushdns"])

    def step_wu(self, plan):
        if not self.admin:
            self.log(self.t("log_needs_admin_skip"), "warn")
            return "skip", None
        result = ("ok", 0)
        try:
            self.log(self.t("log_svc"), "info")
            for svc in ("wuauserv", "bits"):
                status, _ = self._run(["net", "stop", svc])  # "not started" (code 2) is fine
                if status in ("cancel", "skip"):
                    return status, None
            dl = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "SoftwareDistribution", "Download")
            result = self._clean([dl])
        finally:
            # always start the services again, even when skipped or cancelled
            self.log(self.t("log_svc_start"), "info")
            for svc in ("bits", "wuauserv"):
                self._run(["net", "start", svc], cancellable=False)
        return result

    def step_dism_scan(self, plan):
        return self._run(["DISM", "/Online", "/Cleanup-Image", "/ScanHealth"])

    def step_dism_restore(self, plan):
        if reboot_pending():
            self.log(self.t("log_reboot_first"), "warn")
        return self._run(["DISM", "/Online", "/Cleanup-Image", "/RestoreHealth"])

    def step_comp(self, plan):
        return self._run(["DISM", "/Online", "/Cleanup-Image", "/StartComponentCleanup"])

    def step_sfc(self, plan):
        return self._run(["sfc", "/scannow"])

    def step_chkdsk(self, plan):
        drive = plan["drive"]
        if plan["chkdsk_mode"] != "fix":
            status, code = self._run(["chkdsk", drive])
            if status == "failed":  # 1-3: errors were found (not a crash)
                self.log(self.t("log_chkdsk_issues"), "warn")
                return "issues", code
            return status, code
        self.log(self.t("log_chkdsk_fix"), "info")
        # CHKDSK asks before it locks a drive that's in use. The Windows drive: "schedule at
        # next restart?" -> Y. Other drives: "force a dismount?" -> N (open files would be
        # cut off), then "schedule at next restart?" -> Y.
        answers = "Y\r\n" if drive == system_drive() else "N\r\nY\r\n"
        status, code = self._run(["chkdsk", drive, "/f"], stdin_text=answers)
        if status == "failed" and code in (1, 2, 3):
            return "issues", code
        return status, code

    def step_reset_net(self, plan):
        status, code = self._run(["netsh", "winsock", "reset"])
        if status != "ok":
            return status, code
        return self._run(["netsh", "int", "ip", "reset"])

    # ============================================================== update
    @staticmethod
    def _ver(v: str):
        return tuple(int(n) for n in re.findall(r"\d+", v)[:4]) or (0,)

    def check_for_update(self, show_if_latest=False):
        def worker():
            try:
                req = urllib.request.Request(GITHUB_API_LATEST, headers={
                    "User-Agent": f"{APP_ID}/{APP_VERSION}", "Accept": "application/vnd.github+json"})
                with urllib.request.urlopen(req, timeout=10) as r:
                    data = json.loads(r.read().decode("utf-8", "replace"))
                tag = str(data.get("tag_name") or data.get("name") or "").strip()
                if tag and self._ver(tag) > self._ver(APP_VERSION):
                    def ask():
                        if messagebox.askyesno(self.t("update_t"),
                                               self.t("update_available", tag=tag, cur=APP_VERSION), parent=self):
                            webbrowser.open(data.get("html_url") or GITHUB_RELEASES_PAGE)
                    self.call_ui(ask)
                elif show_if_latest:
                    self.call_ui(lambda: messagebox.showinfo(self.t("update_t"),
                                                             self.t("update_latest", cur=APP_VERSION), parent=self))
            except Exception:
                if show_if_latest:
                    self.call_ui(lambda: messagebox.showwarning(self.t("update_t"), self.t("update_failed"),
                                                                parent=self))
        threading.Thread(target=worker, daemon=True).start()

    # =============================================================== close
    def on_close(self):
        if self.closing:
            return
        if self.running:
            if not messagebox.askyesno(self.t("close_running_t"), self.t("close_running"), parent=self, icon="warning"):
                return
            self.closing = True
            threading.Thread(target=self.runner.request_cancel, daemon=True).start()
            self._remember_choices()
            self._wait_then_destroy(time.time())
            return
        self._remember_choices()
        self.closing = True
        self.destroy()

    def _wait_then_destroy(self, since):
        """Let the worker stop (and restart Windows Update services) before quitting."""
        alive = self.worker_thread is not None and self.worker_thread.is_alive()
        if alive and time.time() - since < 30:
            try:
                while True:  # keep draining so the worker never blocks
                    item = self.ui_queue.get_nowait()
                    if item[0] == "call":
                        try:
                            item[1]()
                        except tk.TclError:
                            pass
            except queue.Empty:
                pass
            self.after(150, lambda: self._wait_then_destroy(since))
            return
        self.destroy()


def main():
    if IS_WINDOWS:
        try:
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("ilukezippo.WindowsFixer")
        except Exception:
            pass
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            try:
                ctypes.windll.user32.SetProcessDPIAware()
            except Exception:
                pass

    settings = load_settings()
    if settings.get("always_admin") and not is_admin():
        if relaunch_as_admin():
            return  # the elevated copy takes over
        # UAC was declined: carry on without admin rather than vanishing

    App(settings).mainloop()


if __name__ == "__main__":
    main()
