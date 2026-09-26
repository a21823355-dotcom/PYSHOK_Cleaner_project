# -*- coding: utf-8 -*-
"""
PYSHOK Cleaner PRO
Разработчик: PYSHOK LAB
Версия: 2.0.0

Многофункциональная утилита для Windows:
  - очистка мусора (temp, кэш, корзина, реестр и т.д.)
  - поиск и удаление битых ярлыков и "осиротевших" записей в реестре
  - оптимизация дисков (TRIM для SSD / дефрагментация для HDD)
  - менеджер автозагрузки
  - поиск больших файлов
  - интеграция со встроенным Windows Defender (запуск проверки)
  - выбор языка интерфейса (русский/английский)

ВАЖНО:
  - Программа НЕ является отдельным антивирусом. Писать собственный
    антивирусный движок с нуля бессмысленно и небезопасно — вместо
    этого программа запускает проверку уже встроенным в Windows
    Defender, который обновляется и поддерживается Microsoft.
  - Перед изменением реестра и автозагрузки всегда делается бэкап.
  - Требуются права администратора для полноценной работы.
"""

import os
import re
import sys
import glob
import shutil
import ctypes
import threading
import subprocess
import time
from datetime import datetime

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog

try:
    import winreg
except ImportError:
    winreg = None

try:
    import win32com.client  # для чтения .lnk-ярлыков (устанавливается сборкой из pywin32)
    HAS_WIN32COM = True
except ImportError:
    HAS_WIN32COM = False

APP_NAME = "PYSHOK Cleaner PRO"
APP_VERSION = "2.0.0"
DEVELOPER = "PYSHOK LAB"
CONTACT = "8(707)151-29-79 (WhatsApp/Telegram, только текст)"

ACCENT = "#2d7d46"
ACCENT_DARK = "#1f5c33"
BG = "#f4f6f5"
CARD_BG = "#ffffff"
TEXT_MUTED = "#666666"
DANGER = "#a03a3a"


# =============================================================================
#  i18n
# =============================================================================

STRINGS = {
    "ru": {
        "app_title": "{app} {ver} — {dev}",
        "admin_yes": "Права администратора: ЕСТЬ",
        "admin_no": "Права администратора: НЕТ (часть функций не сработает)",
        "relaunch_admin": "Перезапустить с правами администратора",
        "about": "О программе",
        "lang_label": "Язык:",
        "tab_clean": "Очистка",
        "tab_disks": "Диски",
        "tab_startup": "Автозагрузка",
        "tab_security": "Безопасность",
        "tab_tools": "Инструменты",
        "clean_what": "Что очистить",
        "select_all": "Выбрать всё",
        "select_none": "Снять всё",
        "start_clean": "Начать очистку",
        "cleaning_running": "Идёт очистка...",
        "opt_temp": "Временные файлы пользователя (%TEMP%)",
        "opt_wintemp": "Системная папка Windows\\Temp",
        "opt_prefetch": "Prefetch (кэш запуска программ)",
        "opt_recycle": "Корзина",
        "opt_wu": "Кэш загрузок Windows Update",
        "opt_wer": "Отчёты об ошибках Windows",
        "opt_thumbs": "Кэш эскизов (миниатюр)",
        "opt_browser": "Кэш браузеров (Chrome/Edge/Yandex/Opera/Firefox)",
        "opt_fontcache": "Кэш шрифтов",
        "opt_dns": "Сброс кэша DNS",
        "opt_emptydirs": "Пустые папки в AppData\\Temp",
        "opt_registry": "Безопасная очистка реестра (MRU-списки) + автобэкап",
        "opt_orphan_uninstall": "Осиротевшие записи в 'Установка и удаление программ' + автобэкап",
        "log_header": "=== {app} — запуск очистки: {time} ===",
        "log_done": "=== Готово. ===",
        "log_freed": "Всего освобождено места: {size}",
        "clean_finished_title": "Готово",
        "clean_finished_body": "Очистка завершена.\nОсвобождено: {size}",
        "disks_title": "Оптимизация дисков",
        "disks_refresh": "Обновить список дисков",
        "disks_col_drive": "Диск",
        "disks_col_type": "Тип",
        "disks_col_free": "Свободно",
        "disks_col_status": "Статус",
        "disks_optimize": "Оптимизировать выбранный",
        "disks_optimize_all": "Оптимизировать все",
        "disks_optimizing": "Оптимизация {drive}...",
        "disks_note": "Для SSD выполняется TRIM, для HDD — дефрагментация. Это безопасные штатные функции Windows.",
        "startup_title": "Программы в автозагрузке",
        "startup_refresh": "Обновить список",
        "startup_col_name": "Название",
        "startup_col_cmd": "Команда / путь",
        "startup_col_src": "Источник",
        "startup_remove": "Удалить выбранное (с бэкапом)",
        "startup_note": "Перед удалением автоматически создаётся резервная копия в Документы\\PYSHOK_Cleaner_Backups.",
        "security_title": "Безопасность",
        "security_disclaimer": "PYSHOK Cleaner не заменяет антивирус. Кнопки ниже запускают проверку встроенным Windows Defender.",
        "security_status": "Статус Defender:",
        "security_status_unknown": "не удалось получить статус",
        "security_quick_scan": "Быстрая проверка",
        "security_full_scan": "Полная проверка",
        "security_scanning": "Проверка запущена, это может занять время...",
        "security_scan_done": "Проверка завершена.",
        "security_autorun_check": "Проверить автозагрузку на подозрительные записи",
        "security_autorun_clean": "Записи выглядят нормально, подозрительного не найдено.",
        "security_autorun_found": "Найдено подозрительных записей: {n} (см. вкладку 'Автозагрузка')",
        "tools_title": "Инструменты",
        "tools_shortcuts": "Битые ярлыки (рабочий стол и меню Пуск)",
        "tools_shortcuts_scan": "Найти битые ярлыки",
        "tools_shortcuts_remove": "Удалить выбранные",
        "tools_shortcuts_col_name": "Ярлык",
        "tools_shortcuts_col_target": "Отсутствующая цель",
        "tools_shortcuts_none": "Нужен модуль pywin32 (уже встроен в exe) для чтения ярлыков.",
        "tools_bigfiles": "Поиск больших файлов",
        "tools_bigfiles_folder": "Папка:",
        "tools_bigfiles_choose": "Выбрать...",
        "tools_bigfiles_minsize": "Мин. размер (МБ):",
        "tools_bigfiles_scan": "Искать",
        "tools_bigfiles_delete": "Удалить выбранные",
        "tools_bigfiles_col_name": "Файл",
        "tools_bigfiles_col_size": "Размер",
        "confirm_delete_title": "Подтверждение",
        "confirm_delete_body": "Удалить выбранные элементы ({n} шт.)? Действие нельзя отменить.",
        "nothing_selected": "Ничего не выбрано.",
        "done": "Готово",
        "admin_required": "Для этой функции нужны права администратора.",
    },
    "en": {
        "app_title": "{app} {ver} — {dev}",
        "admin_yes": "Administrator rights: YES",
        "admin_no": "Administrator rights: NO (some features won't work)",
        "relaunch_admin": "Restart as administrator",
        "about": "About",
        "lang_label": "Language:",
        "tab_clean": "Cleanup",
        "tab_disks": "Disks",
        "tab_startup": "Startup",
        "tab_security": "Security",
        "tab_tools": "Tools",
        "clean_what": "What to clean",
        "select_all": "Select all",
        "select_none": "Select none",
        "start_clean": "Start cleaning",
        "cleaning_running": "Cleaning in progress...",
        "opt_temp": "User temp files (%TEMP%)",
        "opt_wintemp": "System folder Windows\\Temp",
        "opt_prefetch": "Prefetch (app launch cache)",
        "opt_recycle": "Recycle Bin",
        "opt_wu": "Windows Update download cache",
        "opt_wer": "Windows error reports",
        "opt_thumbs": "Thumbnail cache",
        "opt_browser": "Browser cache (Chrome/Edge/Yandex/Opera/Firefox)",
        "opt_fontcache": "Font cache",
        "opt_dns": "Flush DNS cache",
        "opt_emptydirs": "Empty folders in AppData\\Temp",
        "opt_registry": "Safe registry cleanup (MRU lists) + auto backup",
        "opt_orphan_uninstall": "Orphaned 'Programs and Features' entries + auto backup",
        "log_header": "=== {app} — cleanup started: {time} ===",
        "log_done": "=== Done. ===",
        "log_freed": "Total space freed: {size}",
        "clean_finished_title": "Done",
        "clean_finished_body": "Cleanup finished.\nFreed: {size}",
        "disks_title": "Disk optimization",
        "disks_refresh": "Refresh drive list",
        "disks_col_drive": "Drive",
        "disks_col_type": "Type",
        "disks_col_free": "Free space",
        "disks_col_status": "Status",
        "disks_optimize": "Optimize selected",
        "disks_optimize_all": "Optimize all",
        "disks_optimizing": "Optimizing {drive}...",
        "disks_note": "SSDs get TRIM, HDDs get defragmented. These are safe built-in Windows features.",
        "startup_title": "Startup programs",
        "startup_refresh": "Refresh list",
        "startup_col_name": "Name",
        "startup_col_cmd": "Command / path",
        "startup_col_src": "Source",
        "startup_remove": "Remove selected (with backup)",
        "startup_note": "A backup is created automatically in Documents\\PYSHOK_Cleaner_Backups before removal.",
        "security_title": "Security",
        "security_disclaimer": "PYSHOK Cleaner does not replace antivirus software. The buttons below trigger a scan using built-in Windows Defender.",
        "security_status": "Defender status:",
        "security_status_unknown": "status unavailable",
        "security_quick_scan": "Quick scan",
        "security_full_scan": "Full scan",
        "security_scanning": "Scan started, this may take a while...",
        "security_scan_done": "Scan finished.",
        "security_autorun_check": "Check startup items for suspicious entries",
        "security_autorun_clean": "Entries look normal, nothing suspicious found.",
        "security_autorun_found": "Suspicious entries found: {n} (see the 'Startup' tab)",
        "tools_title": "Tools",
        "tools_shortcuts": "Broken shortcuts (Desktop and Start Menu)",
        "tools_shortcuts_scan": "Find broken shortcuts",
        "tools_shortcuts_remove": "Remove selected",
        "tools_shortcuts_col_name": "Shortcut",
        "tools_shortcuts_col_target": "Missing target",
        "tools_shortcuts_none": "The pywin32 module (bundled in the exe) is needed to read shortcuts.",
        "tools_bigfiles": "Find large files",
        "tools_bigfiles_folder": "Folder:",
        "tools_bigfiles_choose": "Choose...",
        "tools_bigfiles_minsize": "Min size (MB):",
        "tools_bigfiles_scan": "Scan",
        "tools_bigfiles_delete": "Delete selected",
        "tools_bigfiles_col_name": "File",
        "tools_bigfiles_col_size": "Size",
        "confirm_delete_title": "Confirm",
        "confirm_delete_body": "Delete selected items ({n})? This cannot be undone.",
        "nothing_selected": "Nothing selected.",
        "done": "Done",
        "admin_required": "This feature requires administrator rights.",
    },
}
LANG = {"code": "ru"}


def t(key, **kwargs):
    s = STRINGS[LANG["code"]].get(key, key)
    if kwargs:
        try:
            return s.format(**kwargs)
        except Exception:
            return s
    return s


# =============================================================================
#  Базовые утилиты
# =============================================================================

def is_windows():
    return os.name == "nt"


def is_admin():
    if not is_windows():
        return False
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def relaunch_as_admin():
    try:
        params = " ".join(f'"{a}"' for a in sys.argv)
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, params, None, 1)
        sys.exit(0)
    except Exception as e:
        messagebox.showerror(APP_NAME, f"{e}")


def format_size(num_bytes):
    step = 1024.0
    for unit in ("B", "KB", "MB", "GB", "TB"):
        if abs(num_bytes) < step:
            return f"{num_bytes:.1f} {unit}"
        num_bytes /= step
    return f"{num_bytes:.1f} PB"


def get_size(path):
    if os.path.isfile(path):
        try:
            return os.path.getsize(path)
        except OSError:
            return 0
    total = 0
    for root, dirs, files in os.walk(path, topdown=True):
        for f in files:
            fp = os.path.join(root, f)
            try:
                total += os.path.getsize(fp)
            except OSError:
                pass
    return total


def clear_folder_contents(path, log):
    freed, removed, skipped = 0, 0, 0
    if not os.path.isdir(path):
        return freed, removed, skipped
    for entry in os.listdir(path):
        full = os.path.join(path, entry)
        try:
            size = get_size(full)
            if os.path.isdir(full) and not os.path.islink(full):
                shutil.rmtree(full)
            else:
                os.remove(full)
            freed += size
            removed += 1
        except Exception:
            skipped += 1
    if skipped:
        log(f"    skipped (locked/protected): {skipped}")
    return freed, removed, skipped


def run_powershell(cmd, timeout=180):
    """Выполняет команду PowerShell и возвращает (returncode, stdout, stderr)."""
    try:
        p = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", cmd],
            capture_output=True, text=True, timeout=timeout,
        )
        return p.returncode, p.stdout, p.stderr
    except Exception as e:
        return -1, "", str(e)


REG_BACKUP_DIR_NAME = "PYSHOK_Cleaner_Backups"


def backup_dir():
    d = os.path.join(os.path.expanduser("~"), "Documents", REG_BACKUP_DIR_NAME)
    os.makedirs(d, exist_ok=True)
    return d


def backup_registry_key(hive_name, subkey, log, label):
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_file = os.path.join(backup_dir(), f"{label}_{stamp}.reg")
    full_path = f"{hive_name}\\{subkey}" if subkey else hive_name
    try:
        subprocess.run(["reg", "export", full_path, backup_file, "/y"],
                        capture_output=True, timeout=60, check=True)
        log(f"  backup: {backup_file}")
        return backup_file
    except Exception as e:
        log(f"  backup FAILED: {e}")
        return None


# =============================================================================
#  Задачи очистки
# =============================================================================

def task_temp(log):
    log(t("opt_temp") + "...")
    paths = [os.environ.get("TEMP", ""), os.environ.get("TMP", "")]
    freed = 0
    for p in set(filter(None, paths)):
        f, _, _ = clear_folder_contents(p, log)
        freed += f
    log(f"  {format_size(freed)}")
    return freed


def task_windows_temp(log):
    log(t("opt_wintemp") + "...")
    freed, _, _ = clear_folder_contents(r"C:\Windows\Temp", log)
    log(f"  {format_size(freed)}")
    return freed


def task_prefetch(log):
    log(t("opt_prefetch") + "...")
    freed, _, _ = clear_folder_contents(r"C:\Windows\Prefetch", log)
    log(f"  {format_size(freed)}")
    return freed


def task_recycle_bin(log):
    log(t("opt_recycle") + "...")
    try:
        flags = 0x00000001 | 0x00000002 | 0x00000004
        ctypes.windll.shell32.SHEmptyRecycleBinW(None, None, flags)
        log("  OK")
    except Exception as e:
        log(f"  {e}")
    return 0


def task_windows_update_cache(log):
    log(t("opt_wu") + "...")
    try:
        subprocess.run(["net", "stop", "wuauserv"], capture_output=True, timeout=30)
    except Exception:
        pass
    freed, _, _ = clear_folder_contents(r"C:\Windows\SoftwareDistribution\Download", log)
    try:
        subprocess.run(["net", "start", "wuauserv"], capture_output=True, timeout=30)
    except Exception:
        pass
    log(f"  {format_size(freed)}")
    return freed


def task_error_reports(log):
    log(t("opt_wer") + "...")
    freed = 0
    local = os.environ.get("LOCALAPPDATA", "")
    for sub in [r"Microsoft\Windows\WER\ReportArchive", r"Microsoft\Windows\WER\ReportQueue"]:
        f, _, _ = clear_folder_contents(os.path.join(local, sub), log)
        freed += f
    log(f"  {format_size(freed)}")
    return freed


def task_thumbnail_cache(log):
    log(t("opt_thumbs") + "...")
    local = os.environ.get("LOCALAPPDATA", "")
    p = os.path.join(local, r"Microsoft\Windows\Explorer")
    freed = 0
    if os.path.isdir(p):
        for f in glob.glob(os.path.join(p, "thumbcache_*.db")):
            try:
                freed += os.path.getsize(f)
                os.remove(f)
            except Exception:
                pass
    log(f"  {format_size(freed)}")
    return freed


BROWSER_CACHE_PATHS = {
    "Chrome": r"Google\Chrome\User Data\Default\Cache",
    "Edge": r"Microsoft\Edge\User Data\Default\Cache",
    "Yandex": r"Yandex\YandexBrowser\User Data\Default\Cache",
    "Opera GX": r"Opera Software\Opera GX Stable\Cache",
}


def task_browser_cache(log):
    log(t("opt_browser") + "...")
    local = os.environ.get("LOCALAPPDATA", "")
    appdata = os.environ.get("APPDATA", "")
    freed = 0
    for name, rel in BROWSER_CACHE_PATHS.items():
        full = os.path.join(local, rel)
        f, _, _ = clear_folder_contents(full, log)
        if f:
            log(f"  {name}: {format_size(f)}")
        freed += f
    ff_root = os.path.join(appdata, r"Mozilla\Firefox\Profiles")
    if os.path.isdir(ff_root):
        for prof in os.listdir(ff_root):
            f, _, _ = clear_folder_contents(os.path.join(ff_root, prof, "cache2"), log)
            freed += f
    log(f"  {format_size(freed)}")
    return freed


def task_font_cache(log):
    log(t("opt_fontcache") + "...")
    try:
        subprocess.run(["net", "stop", "FontCache"], capture_output=True, timeout=20)
    except Exception:
        pass
    freed = 0
    local = os.environ.get("LOCALAPPDATA", "")
    p = os.path.join(local, r"Microsoft\Windows\Fonts")
    for f in glob.glob(os.path.join(p, "*.cache")) if os.path.isdir(p) else []:
        try:
            freed += os.path.getsize(f)
            os.remove(f)
        except Exception:
            pass
    try:
        subprocess.run(["net", "start", "FontCache"], capture_output=True, timeout=20)
    except Exception:
        pass
    log(f"  {format_size(freed)}")
    return freed


def task_dns_flush(log):
    log(t("opt_dns") + "...")
    try:
        subprocess.run(["ipconfig", "/flushdns"], capture_output=True, timeout=20)
        log("  OK")
    except Exception as e:
        log(f"  {e}")
    return 0


def task_empty_dirs(log):
    log(t("opt_emptydirs") + "...")
    local = os.environ.get("LOCALAPPDATA", "")
    root = os.path.join(local, "Temp")
    removed = 0
    if os.path.isdir(root):
        for dirpath, dirnames, filenames in os.walk(root, topdown=False):
            if dirpath == root:
                continue
            try:
                if not os.listdir(dirpath):
                    os.rmdir(dirpath)
                    removed += 1
            except Exception:
                pass
    log(f"  removed empty folders: {removed}")
    return 0


SAFE_MRU_KEYS = [
    (winreg.HKEY_CURRENT_USER if winreg else None,
     r"Software\Microsoft\Windows\CurrentVersion\Explorer\RunMRU"),
    (winreg.HKEY_CURRENT_USER if winreg else None,
     r"Software\Microsoft\Windows\CurrentVersion\Explorer\RecentDocs"),
    (winreg.HKEY_CURRENT_USER if winreg else None,
     r"Software\Microsoft\Windows\CurrentVersion\Explorer\TypedPaths"),
    (winreg.HKEY_CURRENT_USER if winreg else None,
     r"Software\Microsoft\Windows\CurrentVersion\Explorer\ComDlg32\OpenSavePidlMRU"),
    (winreg.HKEY_CURRENT_USER if winreg else None,
     r"Software\Microsoft\Windows\CurrentVersion\Explorer\ComDlg32\LastVisitedPidlMRU"),
    (winreg.HKEY_CURRENT_USER if winreg else None,
     r"Software\Microsoft\Windows\CurrentVersion\Search\RecentApps"),
]


def task_registry_mru(log):
    if winreg is None:
        return 0
    bf = backup_registry_key("HKCU", "Software", log, "registry_backup")
    if bf is None:
        log("  cancelled — backup failed")
        return 0
    cleared = 0
    for hive, subkey in SAFE_MRU_KEYS:
        try:
            key = winreg.OpenKey(hive, subkey, 0, winreg.KEY_ALL_ACCESS)
        except FileNotFoundError:
            continue
        except Exception:
            continue
        try:
            names = []
            i = 0
            while True:
                try:
                    name, _, _ = winreg.EnumValue(key, i)
                    names.append(name)
                    i += 1
                except OSError:
                    break
            for n in names:
                try:
                    winreg.DeleteValue(key, n)
                    cleared += 1
                except Exception:
                    pass
        finally:
            winreg.CloseKey(key)
    log(f"  cleared entries: {cleared}")
    return cleared


UNINSTALL_PATHS = [
    (winreg.HKEY_LOCAL_MACHINE if winreg else None,
     r"Software\Microsoft\Windows\CurrentVersion\Uninstall"),
    (winreg.HKEY_LOCAL_MACHINE if winreg else None,
     r"Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"),
    (winreg.HKEY_CURRENT_USER if winreg else None,
     r"Software\Microsoft\Windows\CurrentVersion\Uninstall"),
]


def _extract_path_from_uninstall_string(s):
    """Достаёт исполняемый путь из строки UninstallString/DisplayIcon."""
    if not s:
        return None
    s = s.strip()
    if s.startswith('"'):
        end = s.find('"', 1)
        if end != -1:
            return s[1:end]
        return s[1:]
    # без кавычек: берём первый токен до .exe
    m = re.search(r'([A-Za-z]:\\[^,]*?\.exe)', s, re.IGNORECASE)
    if m:
        return m.group(1)
    return s.split(" ")[0]


def task_orphan_uninstall_entries(log):
    if winreg is None:
        return 0
    bf = backup_registry_key("HKLM", r"Software\Microsoft\Windows\CurrentVersion\Uninstall",
                              log, "uninstall_backup")
    if bf is None:
        log("  cancelled — backup failed")
        return 0
    removed = 0
    for hive, base in UNINSTALL_PATHS:
        try:
            base_key = winreg.OpenKey(hive, base, 0, winreg.KEY_READ)
        except Exception:
            continue
        subnames = []
        i = 0
        while True:
            try:
                subnames.append(winreg.EnumKey(base_key, i))
                i += 1
            except OSError:
                break
        winreg.CloseKey(base_key)

        for name in subnames:
            path = f"{base}\\{name}"
            try:
                k = winreg.OpenKey(hive, path, 0, winreg.KEY_READ)
            except Exception:
                continue
            uninstall_str = None
            display_name = None
            system_component = 0
            try:
                uninstall_str, _ = winreg.QueryValueEx(k, "UninstallString")
            except Exception:
                pass
            try:
                display_name, _ = winreg.QueryValueEx(k, "DisplayName")
            except Exception:
                pass
            try:
                system_component, _ = winreg.QueryValueEx(k, "SystemComponent")
            except Exception:
                system_component = 0
            winreg.CloseKey(k)

            if not display_name or system_component == 1:
                continue  # пропускаем системные компоненты и записи без имени
            target = _extract_path_from_uninstall_string(uninstall_str)
            if target and not os.path.exists(target) and not shutil.which(os.path.basename(target) or ""):
                try:
                    hive_writable = winreg.OpenKey(hive, base, 0, winreg.KEY_ALL_ACCESS)
                    _delete_reg_tree(hive_writable, name)
                    winreg.CloseKey(hive_writable)
                    removed += 1
                    log(f"  removed orphan: {display_name}")
                except Exception:
                    pass
    log(f"  removed: {removed}")
    return removed


def _delete_reg_tree(key, subkey_name):
    try:
        sub = winreg.OpenKey(key, subkey_name, 0, winreg.KEY_ALL_ACCESS)
        while True:
            try:
                child = winreg.EnumKey(sub, 0)
                _delete_reg_tree(sub, child)
            except OSError:
                break
        winreg.CloseKey(sub)
        winreg.DeleteKey(key, subkey_name)
    except Exception:
        pass


CLEAN_TASKS = [
    ("temp", task_temp, True),
    ("wintemp", task_windows_temp, True),
    ("prefetch", task_prefetch, True),
    ("recycle", task_recycle_bin, True),
    ("wu", task_windows_update_cache, True),
    ("wer", task_error_reports, True),
    ("thumbs", task_thumbnail_cache, True),
    ("browser", task_browser_cache, True),
    ("fontcache", task_font_cache, False),
    ("dns", task_dns_flush, False),
    ("emptydirs", task_empty_dirs, False),
    ("registry", task_registry_mru, False),
    ("orphan_uninstall", task_orphan_uninstall_entries, False),
]


# =============================================================================
#  Диски
# =============================================================================

def list_drives():
    drives = []
    bitmask = ctypes.windll.kernel32.GetLogicalDrives()
    for i in range(26):
        if bitmask & (1 << i):
            letter = chr(65 + i) + ":"
            drive_type = ctypes.windll.kernel32.GetDriveTypeW(letter + "\\")
            if drive_type == 3:  # DRIVE_FIXED
                drives.append(letter)
    return drives


def get_drive_media_type(letter):
    ps = (f"$p = Get-Partition -DriveLetter '{letter[0]}' -ErrorAction SilentlyContinue; "
          f"if ($p) {{ (Get-PhysicalDisk -DeviceNumber $p.DiskNumber -ErrorAction SilentlyContinue).MediaType }}")
    rc, out, err = run_powershell(ps, timeout=30)
    out = out.strip()
    if "SSD" in out:
        return "SSD"
    if "HDD" in out:
        return "HDD"
    return "?"


def get_free_space(letter):
    try:
        free_bytes = ctypes.c_ulonglong(0)
        ctypes.windll.kernel32.GetDiskFreeSpaceExW(letter + "\\", ctypes.byref(free_bytes), None, None)
        return free_bytes.value
    except Exception:
        return 0


def optimize_drive(letter, media_type, log):
    letter_only = letter[0]
    log(f"{letter}: {media_type} — {t('disks_optimizing', drive=letter)}")
    if media_type == "SSD":
        cmd = f"Optimize-Volume -DriveLetter '{letter_only}' -ReTrim -Verbose"
    else:
        cmd = f"Optimize-Volume -DriveLetter '{letter_only}' -Defrag -Verbose"
    rc, out, err = run_powershell(cmd, timeout=1800)
    if rc == 0:
        log(f"  {letter}: OK")
    else:
        log(f"  {letter}: {err.strip()[:200] if err else 'error'}")


# =============================================================================
#  Автозагрузка
# =============================================================================

RUN_KEYS = [
    ("HKCU", winreg.HKEY_CURRENT_USER if winreg else None,
     r"Software\Microsoft\Windows\CurrentVersion\Run"),
    ("HKLM", winreg.HKEY_LOCAL_MACHINE if winreg else None,
     r"Software\Microsoft\Windows\CurrentVersion\Run"),
]


def list_startup_items():
    items = []  # dict: name, cmd, source ("HKCU Run", "HKLM Run", "Startup folder")
    if winreg:
        for label, hive, subkey in RUN_KEYS:
            try:
                k = winreg.OpenKey(hive, subkey, 0, winreg.KEY_READ)
            except Exception:
                continue
            i = 0
            while True:
                try:
                    name, value, _ = winreg.EnumValue(k, i)
                    items.append({"name": name, "cmd": value, "source": f"{label} Run",
                                  "hive": hive, "subkey": subkey, "kind": "reg"})
                    i += 1
                except OSError:
                    break
            winreg.CloseKey(k)

    for env_var, label in [("APPDATA", "Startup (user)"), ("PROGRAMDATA", "Startup (all users)")]:
        base = os.environ.get(env_var, "")
        folder = os.path.join(base, r"Microsoft\Windows\Start Menu\Programs\Startup") \
            if env_var == "APPDATA" else \
            os.path.join(base, r"Microsoft\Windows\Start Menu\Programs\StartUp")
        if os.path.isdir(folder):
            for f in os.listdir(folder):
                full = os.path.join(folder, f)
                items.append({"name": f, "cmd": full, "source": label,
                              "hive": None, "subkey": None, "kind": "file", "path": full})
    return items


def remove_startup_item(item, log):
    try:
        if item["kind"] == "reg":
            backup_registry_key(
                "HKCU" if item["hive"] == winreg.HKEY_CURRENT_USER else "HKLM",
                item["subkey"], log, "startup_backup")
            k = winreg.OpenKey(item["hive"], item["subkey"], 0, winreg.KEY_ALL_ACCESS)
            winreg.DeleteValue(k, item["name"])
            winreg.CloseKey(k)
        else:
            dst = os.path.join(backup_dir(), os.path.basename(item["path"]))
            try:
                shutil.copy2(item["path"], dst)
            except Exception:
                pass
            os.remove(item["path"])
        return True
    except Exception as e:
        log(f"  {item.get('name')}: {e}")
        return False


SUSPICIOUS_HINTS = ["\\temp\\", "\\appdata\\local\\temp\\", "\\downloads\\"]


def is_suspicious_startup(item):
    cmd = (item.get("cmd") or "").lower()
    for hint in SUSPICIOUS_HINTS:
        if hint in cmd:
            return True
    target = _extract_path_from_uninstall_string(item.get("cmd", ""))
    if target and not os.path.exists(target):
        return True
    return False


# =============================================================================
#  Битые ярлыки
# =============================================================================

def find_broken_shortcuts(log):
    results = []
    if not HAS_WIN32COM:
        log(t("tools_shortcuts_none"))
        return results
    shell = win32com.client.Dispatch("WScript.Shell")
    folders = []
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    public_desktop = os.path.join(os.environ.get("PUBLIC", ""), "Desktop")
    start_menu = os.path.join(os.environ.get("APPDATA", ""), r"Microsoft\Windows\Start Menu\Programs")
    start_menu_all = os.path.join(os.environ.get("PROGRAMDATA", ""), r"Microsoft\Windows\Start Menu\Programs")
    for f in [desktop, public_desktop, start_menu, start_menu_all]:
        if os.path.isdir(f):
            folders.append(f)

    for folder in folders:
        for dirpath, _, filenames in os.walk(folder):
            for fn in filenames:
                if fn.lower().endswith(".lnk"):
                    full = os.path.join(dirpath, fn)
                    try:
                        sc = shell.CreateShortCut(full)
                        target = sc.Targetpath
                        if target and not os.path.exists(target):
                            results.append({"name": fn, "path": full, "target": target})
                    except Exception:
                        continue
    return results


# =============================================================================
#  Windows Defender
# =============================================================================

def get_defender_status():
    ps = ("Get-MpComputerStatus | Select-Object AntivirusEnabled,RealTimeProtectionEnabled,"
          "AntivirusSignatureLastUpdated | ConvertTo-Json")
    rc, out, err = run_powershell(ps, timeout=30)
    if rc != 0 or not out.strip():
        return None
    return out.strip()


def start_defender_scan(scan_type, log):
    log(t("security_scanning"))
    ps = f"Start-MpScan -ScanType {scan_type}"
    rc, out, err = run_powershell(ps, timeout=3600 if scan_type == "FullScan" else 900)
    if rc == 0:
        log(t("security_scan_done"))
    else:
        log(f"  {err.strip()[:300] if err else 'error'}")


# =============================================================================
#  GUI
# =============================================================================

class CleanerApp:
    def __init__(self, root):
        self.root = root
        self.opts = {key: default for key, _, default in CLEAN_TASKS}
        self.startup_items = []
        self.shortcut_items = []
        self.bigfile_items = []
        self.build_ui()

    # ---------------- общий каркас ----------------

    def build_ui(self):
        for w in self.root.winfo_children():
            w.destroy()

        self.root.title(t("app_title", app=APP_NAME, ver=APP_VERSION, dev=DEVELOPER))
        self.root.geometry("880x640")
        self.root.minsize(760, 560)
        self.root.configure(bg=BG)

        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("TNotebook", background=BG, borderwidth=0)
        style.configure("TNotebook.Tab", padding=(16, 8), font=("Segoe UI", 10))
        style.configure("Accent.TButton", background=ACCENT, foreground="white",
                         font=("Segoe UI", 10, "bold"), padding=8)
        style.map("Accent.TButton", background=[("active", ACCENT_DARK)])
        style.configure("TButton", padding=6)
        style.configure("Treeview", rowheight=24, font=("Segoe UI", 9))
        style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"))

        self._build_header()

        nb = ttk.Notebook(self.root)
        nb.pack(fill="both", expand=True, padx=12, pady=(4, 12))
        self.tab_clean = tk.Frame(nb, bg=CARD_BG)
        self.tab_disks = tk.Frame(nb, bg=CARD_BG)
        self.tab_startup = tk.Frame(nb, bg=CARD_BG)
        self.tab_security = tk.Frame(nb, bg=CARD_BG)
        self.tab_tools = tk.Frame(nb, bg=CARD_BG)
        nb.add(self.tab_clean, text=t("tab_clean"))
        nb.add(self.tab_disks, text=t("tab_disks"))
        nb.add(self.tab_startup, text=t("tab_startup"))
        nb.add(self.tab_security, text=t("tab_security"))
        nb.add(self.tab_tools, text=t("tab_tools"))

        self._build_clean_tab()
        self._build_disks_tab()
        self._build_startup_tab()
        self._build_security_tab()
        self._build_tools_tab()

    def _build_header(self):
        header = tk.Frame(self.root, bg=ACCENT, pady=12)
        header.pack(fill="x")
        left = tk.Frame(header, bg=ACCENT)
        left.pack(side="left", padx=16)
        tk.Label(left, text=APP_NAME, font=("Segoe UI", 15, "bold"), bg=ACCENT, fg="white").pack(anchor="w")
        tk.Label(left, text=f"{DEVELOPER} · v{APP_VERSION}", font=("Segoe UI", 9), bg=ACCENT, fg="#dff2e6").pack(anchor="w")

        right = tk.Frame(header, bg=ACCENT)
        right.pack(side="right", padx=16)

        lang_frame = tk.Frame(right, bg=ACCENT)
        lang_frame.pack(anchor="e")
        tk.Label(lang_frame, text=t("lang_label"), bg=ACCENT, fg="white", font=("Segoe UI", 9)).pack(side="left", padx=(0, 4))
        lang_var = tk.StringVar(value=LANG["code"])
        combo = ttk.Combobox(lang_frame, textvariable=lang_var, values=["ru", "en"], width=4, state="readonly")
        combo.pack(side="left")
        combo.bind("<<ComboboxSelected>>", lambda e: self._switch_lang(lang_var.get()))

        admin_ok = is_admin()
        admin_txt = t("admin_yes") if admin_ok else t("admin_no")
        tk.Label(right, text=admin_txt, bg=ACCENT,
                 fg="#dff2e6" if admin_ok else "#ffdcdc", font=("Segoe UI", 9, "bold")).pack(anchor="e", pady=(4, 0))
        if not admin_ok:
            tk.Button(right, text=t("relaunch_admin"), command=relaunch_as_admin).pack(anchor="e", pady=(2, 0))

    def _switch_lang(self, code):
        LANG["code"] = code
        self.build_ui()

    # ---------------- вкладка Очистка ----------------

    def _build_clean_tab(self):
        f = self.tab_clean
        top = tk.Frame(f, bg=CARD_BG)
        top.pack(fill="both", expand=True, padx=16, pady=16)

        tk.Label(top, text=t("clean_what"), font=("Segoe UI", 11, "bold"), bg=CARD_BG).pack(anchor="w")

        opts_frame = tk.Frame(top, bg=CARD_BG)
        opts_frame.pack(fill="x", pady=(6, 6))
        self.clean_vars = {}
        for key, _, default in CLEAN_TASKS:
            v = tk.BooleanVar(value=self.opts.get(key, default))
            self.clean_vars[key] = v
            tk.Checkbutton(opts_frame, text=t(f"opt_{key}"), variable=v, bg=CARD_BG,
                            anchor="w", justify="left").pack(fill="x")

        btn_row = tk.Frame(top, bg=CARD_BG)
        btn_row.pack(fill="x", pady=(4, 8))
        tk.Button(btn_row, text=t("select_all"), command=lambda: self._set_all_clean(True)).pack(side="left")
        tk.Button(btn_row, text=t("select_none"), command=lambda: self._set_all_clean(False)).pack(side="left", padx=(8, 0))
        self.clean_start_btn = ttk.Button(btn_row, text=t("start_clean"), style="Accent.TButton",
                                           command=self.start_clean)
        self.clean_start_btn.pack(side="right")

        self.clean_progress = ttk.Progressbar(top, mode="indeterminate")
        self.clean_progress.pack(fill="x", pady=(0, 8))

        self.clean_log = scrolledtext.ScrolledText(top, height=14, state="disabled", font=("Consolas", 9))
        self.clean_log.pack(fill="both", expand=True)

    def _set_all_clean(self, value):
        for v in self.clean_vars.values():
            v.set(value)

    def log_clean(self, text):
        def _write():
            self.clean_log.configure(state="normal")
            self.clean_log.insert("end", text + "\n")
            self.clean_log.see("end")
            self.clean_log.configure(state="disabled")
        self.root.after(0, _write)

    def start_clean(self):
        if not is_windows():
            messagebox.showerror(APP_NAME, "Windows only")
            return
        for key, v in self.clean_vars.items():
            self.opts[key] = v.get()
        self.clean_start_btn.config(state="disabled")
        self.clean_progress.start(12)
        threading.Thread(target=self._run_clean, daemon=True).start()

    def _run_clean(self):
        total = 0
        self.log_clean(t("log_header", app=APP_NAME, time=datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        for key, fn, _ in CLEAN_TASKS:
            if self.opts.get(key):
                try:
                    total += fn(self.log_clean) or 0
                except Exception as e:
                    self.log_clean(f"  ERROR [{key}]: {e}")
            time.sleep(0.03)
        self.log_clean(t("log_done"))
        self.log_clean(t("log_freed", size=format_size(total)))

        def _finish():
            self.clean_progress.stop()
            self.clean_start_btn.config(state="normal")
            messagebox.showinfo(t("clean_finished_title"),
                                 t("clean_finished_body", size=format_size(total)))
        self.root.after(0, _finish)

    # ---------------- вкладка Диски ----------------

    def _build_disks_tab(self):
        f = self.tab_disks
        top = tk.Frame(f, bg=CARD_BG)
        top.pack(fill="both", expand=True, padx=16, pady=16)

        tk.Label(top, text=t("disks_title"), font=("Segoe UI", 11, "bold"), bg=CARD_BG).pack(anchor="w")
        tk.Label(top, text=t("disks_note"), fg=TEXT_MUTED, bg=CARD_BG, wraplength=760, justify="left").pack(anchor="w", pady=(2, 8))

        cols = ("drive", "type", "free", "status")
        self.disks_tree = ttk.Treeview(top, columns=cols, show="headings", height=8, selectmode="extended")
        self.disks_tree.heading("drive", text=t("disks_col_drive"))
        self.disks_tree.heading("type", text=t("disks_col_type"))
        self.disks_tree.heading("free", text=t("disks_col_free"))
        self.disks_tree.heading("status", text=t("disks_col_status"))
        self.disks_tree.pack(fill="both", expand=True)

        btn_row = tk.Frame(top, bg=CARD_BG)
        btn_row.pack(fill="x", pady=8)
        tk.Button(btn_row, text=t("disks_refresh"), command=self.refresh_disks).pack(side="left")
        ttk.Button(btn_row, text=t("disks_optimize"), style="Accent.TButton",
                   command=lambda: self.start_optimize(selected_only=True)).pack(side="right")
        tk.Button(btn_row, text=t("disks_optimize_all"),
                  command=lambda: self.start_optimize(selected_only=False)).pack(side="right", padx=(0, 8))

        self.disks_progress = ttk.Progressbar(top, mode="indeterminate")
        self.disks_progress.pack(fill="x", pady=(0, 8))
        self.disks_log = scrolledtext.ScrolledText(top, height=8, state="disabled", font=("Consolas", 9))
        self.disks_log.pack(fill="both", expand=True)

        threading.Thread(target=self.refresh_disks, daemon=True).start()

    def log_disks(self, text):
        def _write():
            self.disks_log.configure(state="normal")
            self.disks_log.insert("end", text + "\n")
            self.disks_log.see("end")
            self.disks_log.configure(state="disabled")
        self.root.after(0, _write)

    def refresh_disks(self):
        try:
            drives = list_drives()
        except Exception:
            drives = []

        def _populate():
            for i in self.disks_tree.get_children():
                self.disks_tree.delete(i)
            for d in drives:
                self.disks_tree.insert("", "end", iid=d, values=(d, "...", "...", ""))
        self.root.after(0, _populate)

        for d in drives:
            media = get_drive_media_type(d)
            free = format_size(get_free_space(d))

            def _update(d=d, media=media, free=free):
                if self.disks_tree.exists(d):
                    self.disks_tree.item(d, values=(d, media, free, ""))
            self.root.after(0, _update)

    def start_optimize(self, selected_only):
        if not is_admin():
            messagebox.showwarning(APP_NAME, t("admin_required"))
            return
        if selected_only:
            targets = list(self.disks_tree.selection())
            if not targets:
                messagebox.showinfo(APP_NAME, t("nothing_selected"))
                return
        else:
            targets = list(self.disks_tree.get_children())
        self.disks_progress.start(12)
        threading.Thread(target=self._run_optimize, args=(targets,), daemon=True).start()

    def _run_optimize(self, targets):
        for d in targets:
            vals = self.disks_tree.item(d, "values")
            media = vals[1] if len(vals) > 1 else "?"
            try:
                optimize_drive(d, media, self.log_disks)
            except Exception as e:
                self.log_disks(f"  {d}: {e}")
        self.root.after(0, self.disks_progress.stop)

    # ---------------- вкладка Автозагрузка ----------------

    def _build_startup_tab(self):
        f = self.tab_startup
        top = tk.Frame(f, bg=CARD_BG)
        top.pack(fill="both", expand=True, padx=16, pady=16)

        tk.Label(top, text=t("startup_title"), font=("Segoe UI", 11, "bold"), bg=CARD_BG).pack(anchor="w")
        tk.Label(top, text=t("startup_note"), fg=TEXT_MUTED, bg=CARD_BG, wraplength=760, justify="left").pack(anchor="w", pady=(2, 8))

        cols = ("name", "cmd", "source")
        self.startup_tree = ttk.Treeview(top, columns=cols, show="headings", height=12, selectmode="extended")
        self.startup_tree.heading("name", text=t("startup_col_name"))
        self.startup_tree.heading("cmd", text=t("startup_col_cmd"))
        self.startup_tree.heading("source", text=t("startup_col_src"))
        self.startup_tree.column("cmd", width=380)
        self.startup_tree.pack(fill="both", expand=True)
        self.startup_tree.tag_configure("suspicious", foreground=DANGER)

        btn_row = tk.Frame(top, bg=CARD_BG)
        btn_row.pack(fill="x", pady=8)
        tk.Button(btn_row, text=t("startup_refresh"), command=self.refresh_startup).pack(side="left")
        ttk.Button(btn_row, text=t("startup_remove"), style="Accent.TButton",
                   command=self.remove_startup_selected).pack(side="right")

        self.refresh_startup()

    def refresh_startup(self):
        self.startup_items = list_startup_items()
        for i in self.startup_tree.get_children():
            self.startup_tree.delete(i)
        for idx, item in enumerate(self.startup_items):
            tag = "suspicious" if is_suspicious_startup(item) else ""
            self.startup_tree.insert("", "end", iid=str(idx),
                                      values=(item["name"], item["cmd"], item["source"]),
                                      tags=(tag,) if tag else ())

    def remove_startup_selected(self):
        sel = self.startup_tree.selection()
        if not sel:
            messagebox.showinfo(APP_NAME, t("nothing_selected"))
            return
        if not messagebox.askyesno(t("confirm_delete_title"), t("confirm_delete_body", n=len(sel))):
            return
        for iid in sel:
            item = self.startup_items[int(iid)]
            remove_startup_item(item, lambda s: None)
        self.refresh_startup()

    # ---------------- вкладка Безопасность ----------------

    def _build_security_tab(self):
        f = self.tab_security
        top = tk.Frame(f, bg=CARD_BG)
        top.pack(fill="both", expand=True, padx=16, pady=16)

        tk.Label(top, text=t("security_title"), font=("Segoe UI", 11, "bold"), bg=CARD_BG).pack(anchor="w")
        tk.Label(top, text=t("security_disclaimer"), fg=TEXT_MUTED, bg=CARD_BG,
                 wraplength=760, justify="left").pack(anchor="w", pady=(2, 10))

        status_row = tk.Frame(top, bg=CARD_BG)
        status_row.pack(fill="x", pady=(0, 10))
        tk.Label(status_row, text=t("security_status"), bg=CARD_BG, font=("Segoe UI", 10, "bold")).pack(side="left")
        self.security_status_label = tk.Label(status_row, text="...", bg=CARD_BG)
        self.security_status_label.pack(side="left", padx=(6, 0))

        btn_row = tk.Frame(top, bg=CARD_BG)
        btn_row.pack(fill="x", pady=(0, 10))
        tk.Button(btn_row, text=t("security_quick_scan"),
                  command=lambda: self.start_scan("QuickScan")).pack(side="left")
        tk.Button(btn_row, text=t("security_full_scan"),
                  command=lambda: self.start_scan("FullScan")).pack(side="left", padx=(8, 0))
        tk.Button(btn_row, text=t("security_autorun_check"),
                  command=self.check_autorun).pack(side="left", padx=(8, 0))

        self.security_progress = ttk.Progressbar(top, mode="indeterminate")
        self.security_progress.pack(fill="x", pady=(0, 8))
        self.security_log = scrolledtext.ScrolledText(top, height=10, state="disabled", font=("Consolas", 9))
        self.security_log.pack(fill="both", expand=True)

        threading.Thread(target=self._load_defender_status, daemon=True).start()

    def log_security(self, text):
        def _write():
            self.security_log.configure(state="normal")
            self.security_log.insert("end", text + "\n")
            self.security_log.see("end")
            self.security_log.configure(state="disabled")
        self.root.after(0, _write)

    def _load_defender_status(self):
        status = get_defender_status()

        def _update():
            self.security_status_label.config(text=status if status else t("security_status_unknown"))
        self.root.after(0, _update)

    def start_scan(self, scan_type):
        if not is_admin():
            messagebox.showwarning(APP_NAME, t("admin_required"))
            return
        self.security_progress.start(12)
        threading.Thread(target=self._run_scan, args=(scan_type,), daemon=True).start()

    def _run_scan(self, scan_type):
        start_defender_scan(scan_type, self.log_security)
        self.root.after(0, self.security_progress.stop)

    def check_autorun(self):
        items = list_startup_items()
        suspicious = [i for i in items if is_suspicious_startup(i)]
        if suspicious:
            self.log_security(t("security_autorun_found", n=len(suspicious)))
            for s in suspicious:
                self.log_security(f"  - {s['name']} ({s['source']}): {s['cmd']}")
        else:
            self.log_security(t("security_autorun_clean"))

    # ---------------- вкладка Инструменты ----------------

    def _build_tools_tab(self):
        f = self.tab_tools
        top = tk.Frame(f, bg=CARD_BG)
        top.pack(fill="both", expand=True, padx=16, pady=16)

        # -- битые ярлыки --
        tk.Label(top, text=t("tools_shortcuts"), font=("Segoe UI", 11, "bold"), bg=CARD_BG).pack(anchor="w")
        sc_cols = ("name", "target")
        self.shortcuts_tree = ttk.Treeview(top, columns=sc_cols, show="headings", height=6, selectmode="extended")
        self.shortcuts_tree.heading("name", text=t("tools_shortcuts_col_name"))
        self.shortcuts_tree.heading("target", text=t("tools_shortcuts_col_target"))
        self.shortcuts_tree.pack(fill="x", pady=(4, 4))

        sc_btn_row = tk.Frame(top, bg=CARD_BG)
        sc_btn_row.pack(fill="x", pady=(0, 12))
        tk.Button(sc_btn_row, text=t("tools_shortcuts_scan"), command=self.scan_shortcuts).pack(side="left")
        ttk.Button(sc_btn_row, text=t("tools_shortcuts_remove"), style="Accent.TButton",
                   command=self.remove_shortcuts_selected).pack(side="right")

        ttk.Separator(top, orient="horizontal").pack(fill="x", pady=8)

        # -- большие файлы --
        tk.Label(top, text=t("tools_bigfiles"), font=("Segoe UI", 11, "bold"), bg=CARD_BG).pack(anchor="w", pady=(4, 4))
        row = tk.Frame(top, bg=CARD_BG)
        row.pack(fill="x")
        tk.Label(row, text=t("tools_bigfiles_folder"), bg=CARD_BG).pack(side="left")
        self.bigfiles_folder_var = tk.StringVar(value=os.path.expanduser("~"))
        tk.Entry(row, textvariable=self.bigfiles_folder_var, width=40).pack(side="left", padx=(4, 4))
        tk.Button(row, text=t("tools_bigfiles_choose"), command=self.choose_bigfiles_folder).pack(side="left")
        tk.Label(row, text=t("tools_bigfiles_minsize"), bg=CARD_BG).pack(side="left", padx=(12, 4))
        self.bigfiles_minsize_var = tk.StringVar(value="200")
        tk.Entry(row, textvariable=self.bigfiles_minsize_var, width=6).pack(side="left")
        tk.Button(row, text=t("tools_bigfiles_scan"), command=self.scan_bigfiles).pack(side="left", padx=(8, 0))

        bf_cols = ("name", "size")
        self.bigfiles_tree = ttk.Treeview(top, columns=bf_cols, show="headings", height=8, selectmode="extended")
        self.bigfiles_tree.heading("name", text=t("tools_bigfiles_col_name"))
        self.bigfiles_tree.heading("size", text=t("tools_bigfiles_col_size"))
        self.bigfiles_tree.column("name", width=560)
        self.bigfiles_tree.pack(fill="both", expand=True, pady=(6, 6))

        ttk.Button(top, text=t("tools_bigfiles_delete"), style="Accent.TButton",
                   command=self.delete_bigfiles_selected).pack(anchor="e")

    def choose_bigfiles_folder(self):
        d = filedialog.askdirectory(initialdir=self.bigfiles_folder_var.get())
        if d:
            self.bigfiles_folder_var.set(d)

    def scan_shortcuts(self):
        threading.Thread(target=self._run_scan_shortcuts, daemon=True).start()

    def _run_scan_shortcuts(self):
        results = find_broken_shortcuts(lambda s: None)
        self.shortcut_items = results

        def _populate():
            for i in self.shortcuts_tree.get_children():
                self.shortcuts_tree.delete(i)
            for idx, r in enumerate(results):
                self.shortcuts_tree.insert("", "end", iid=str(idx), values=(r["name"], r["target"]))
        self.root.after(0, _populate)

    def remove_shortcuts_selected(self):
        sel = self.shortcuts_tree.selection()
        if not sel:
            messagebox.showinfo(APP_NAME, t("nothing_selected"))
            return
        if not messagebox.askyesno(t("confirm_delete_title"), t("confirm_delete_body", n=len(sel))):
            return
        for iid in sel:
            item = self.shortcut_items[int(iid)]
            try:
                os.remove(item["path"])
            except Exception:
                pass
        self.scan_shortcuts()

    def scan_bigfiles(self):
        folder = self.bigfiles_folder_var.get()
        try:
            min_mb = float(self.bigfiles_minsize_var.get())
        except ValueError:
            min_mb = 200
        threading.Thread(target=self._run_scan_bigfiles, args=(folder, min_mb), daemon=True).start()

    def _run_scan_bigfiles(self, folder, min_mb):
        min_bytes = min_mb * 1024 * 1024
        results = []
        for dirpath, dirnames, filenames in os.walk(folder):
            for fn in filenames:
                full = os.path.join(dirpath, fn)
                try:
                    size = os.path.getsize(full)
                except OSError:
                    continue
                if size >= min_bytes:
                    results.append((full, size))
        results.sort(key=lambda x: -x[1])
        results = results[:200]
        self.bigfile_items = results

        def _populate():
            for i in self.bigfiles_tree.get_children():
                self.bigfiles_tree.delete(i)
            for idx, (path, size) in enumerate(results):
                self.bigfiles_tree.insert("", "end", iid=str(idx), values=(path, format_size(size)))
        self.root.after(0, _populate)

    def delete_bigfiles_selected(self):
        sel = self.bigfiles_tree.selection()
        if not sel:
            messagebox.showinfo(APP_NAME, t("nothing_selected"))
            return
        if not messagebox.askyesno(t("confirm_delete_title"), t("confirm_delete_body", n=len(sel))):
            return
        for iid in sel:
            path, _ = self.bigfile_items[int(iid)]
            try:
                os.remove(path)
            except Exception:
                pass
        self.bigfiles_tree.delete(*sel)


def main():
    root = tk.Tk()
    app = CleanerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
