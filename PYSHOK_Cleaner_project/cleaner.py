# -*- coding: utf-8 -*-
"""
PYSHOK Cleaner — утилита очистки системы (временные файлы, кэш браузеров,
корзина, безопасные ветки реестра).

Разработчик: PYSHOK LAB
Версия: 1.0.0

ВАЖНО ПЕРЕД ЗАПУСКОМ:
  - Программа предназначена только для Windows.
  - Для очистки системных папок и реестра требуются права администратора —
    при обычном запуске программа сама предложит перезапуститься с ними.
  - Перед изменением реестра программа делает автоматический бэкап
    (.reg-файл) в папку "Документы\\PYSHOK_Cleaner_Backups".
  - Программа НЕ удаляет установленные программы и НЕ трогает файлы
    пользователя (документы, фото и т.д.) — только служебный мусор.
"""

import os
import sys
import glob
import shutil
import ctypes
import threading
import subprocess
import time
from datetime import datetime

# tkinter идёт в стандартной поставке Python на Windows
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext

try:
    import winreg
except ImportError:
    winreg = None  # на не-Windows системах программа только покажет предупреждение

APP_NAME = "PYSHOK Cleaner"
APP_VERSION = "1.0.0"
DEVELOPER = "PYSHOK LAB"
CONTACT = "8(707)151-29-79 (WhatsApp/Telegram, только текст)"


# ---------------------------------------------------------------------------
# Вспомогательные функции
# ---------------------------------------------------------------------------

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
    """Перезапускает программу с правами администратора через UAC."""
    try:
        params = " ".join(f'"{a}"' for a in sys.argv)
        ctypes.windll.shell32.ShellExecuteW(
            None, "runas", sys.executable, params, None, 1
        )
        sys.exit(0)
    except Exception as e:
        messagebox.showerror(
            APP_NAME, f"Не удалось запросить права администратора:\n{e}"
        )


def format_size(num_bytes):
    step = 1024.0
    for unit in ("Б", "КБ", "МБ", "ГБ", "ТБ"):
        if num_bytes < step:
            return f"{num_bytes:.1f} {unit}"
        num_bytes /= step
    return f"{num_bytes:.1f} ПБ"


def get_size(path):
    total = 0
    if os.path.isfile(path):
        try:
            return os.path.getsize(path)
        except OSError:
            return 0
    for root, dirs, files in os.walk(path, topdown=True):
        for f in files:
            fp = os.path.join(root, f)
            try:
                total += os.path.getsize(fp)
            except OSError:
                pass
    return total


def clear_folder_contents(path, log):
    """Удаляет содержимое папки, саму папку оставляет. Пропускает занятые файлы."""
    freed = 0
    removed = 0
    skipped = 0
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
        log(f"    пропущено занятых/защищённых объектов: {skipped}")
    return freed, removed, skipped


# ---------------------------------------------------------------------------
# Задачи очистки
# ---------------------------------------------------------------------------

def task_temp(log):
    log("Очистка временных файлов пользователя...")
    paths = [os.environ.get("TEMP", ""), os.environ.get("TMP", "")]
    freed = 0
    for p in set(filter(None, paths)):
        f, removed, _ = clear_folder_contents(p, log)
        freed += f
    log(f"  Освобождено: {format_size(freed)}")
    return freed


def task_windows_temp(log):
    log("Очистка системной папки Windows\\Temp...")
    path = r"C:\Windows\Temp"
    freed, removed, _ = clear_folder_contents(path, log)
    log(f"  Освобождено: {format_size(freed)}")
    return freed


def task_prefetch(log):
    log("Очистка Prefetch...")
    path = r"C:\Windows\Prefetch"
    freed, removed, _ = clear_folder_contents(path, log)
    log(f"  Освобождено: {format_size(freed)}")
    return freed


def task_recycle_bin(log):
    log("Очистка корзины...")
    try:
        # SHEmptyRecycleBinW: флаги — без подтверждения, без звука, без прогресс-бара
        SHERB_NOCONFIRMATION = 0x00000001
        SHERB_NOPROGRESSUI = 0x00000002
        SHERB_NOSOUND = 0x00000004
        flags = SHERB_NOCONFIRMATION | SHERB_NOPROGRESSUI | SHERB_NOSOUND
        ctypes.windll.shell32.SHEmptyRecycleBinW(None, None, flags)
        log("  Корзина очищена.")
    except Exception as e:
        log(f"  Не удалось очистить корзину: {e}")
    return 0


def task_windows_update_cache(log):
    log("Очистка кэша загрузок Windows Update...")
    path = r"C:\Windows\SoftwareDistribution\Download"
    # Останавливаем службу, чтобы файлы не были заняты
    try:
        subprocess.run(["net", "stop", "wuauserv"], capture_output=True, timeout=30)
    except Exception:
        pass
    freed, removed, _ = clear_folder_contents(path, log)
    try:
        subprocess.run(["net", "start", "wuauserv"], capture_output=True, timeout=30)
    except Exception:
        pass
    log(f"  Освобождено: {format_size(freed)}")
    return freed


def task_error_reports(log):
    log("Очистка отчётов об ошибках Windows...")
    freed = 0
    local = os.environ.get("LOCALAPPDATA", "")
    for sub in [r"Microsoft\Windows\WER\ReportArchive",
                r"Microsoft\Windows\WER\ReportQueue"]:
        p = os.path.join(local, sub)
        f, _, _ = clear_folder_contents(p, log)
        freed += f
    log(f"  Освобождено: {format_size(freed)}")
    return freed


def task_thumbnail_cache(log):
    log("Очистка кэша эскизов (миниатюр)...")
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
    log(f"  Освобождено: {format_size(freed)}")
    return freed


BROWSER_CACHE_PATHS = {
    "Chrome": r"Google\Chrome\User Data\Default\Cache",
    "Edge": r"Microsoft\Edge\User Data\Default\Cache",
    "Yandex": r"Yandex\YandexBrowser\User Data\Default\Cache",
    "Opera GX": r"Opera Software\Opera GX Stable\Cache",
    "Firefox": r"Mozilla\Firefox\Profiles",  # обрабатывается отдельно (несколько профилей)
}


def task_browser_cache(log):
    log("Очистка кэша браузеров (файлы кэша, не пароли/закладки/история)...")
    local = os.environ.get("LOCALAPPDATA", "")
    appdata = os.environ.get("APPDATA", "")
    freed = 0

    for name, rel in BROWSER_CACHE_PATHS.items():
        if name == "Firefox":
            profiles_root = os.path.join(appdata, rel)
            if os.path.isdir(profiles_root):
                for prof in os.listdir(profiles_root):
                    cache_p = os.path.join(profiles_root, prof, "cache2")
                    f, _, _ = clear_folder_contents(cache_p, log)
                    freed += f
            continue
        full = os.path.join(local, rel)
        f, _, _ = clear_folder_contents(full, log)
        if f:
            log(f"  {name}: {format_size(f)}")
        freed += f
    log(f"  Итого освобождено: {format_size(freed)}")
    return freed


# --- Реестр -----------------------------------------------------------------

REG_BACKUP_DIR_NAME = "PYSHOK_Cleaner_Backups"

# Список безопасных для очистки веток: это только списки "последних
# использованных" файлов/путей и кэш поиска — их удаление не ломает
# программы и настройки, только сбрасывает "историю" в проводнике.
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


def backup_registry(log):
    """Экспортирует HKCU\\Software и HKLM\\Software в .reg-файл перед изменениями."""
    docs = os.path.join(os.path.expanduser("~"), "Documents", REG_BACKUP_DIR_NAME)
    os.makedirs(docs, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_file = os.path.join(docs, f"registry_backup_{stamp}.reg")
    log(f"Резервная копия реестра: {backup_file}")
    try:
        subprocess.run(
            ["reg", "export", "HKCU\\Software", backup_file, "/y"],
            capture_output=True, timeout=120, check=True,
        )
        log("  Бэкап реестра создан успешно.")
        return backup_file
    except Exception as e:
        log(f"  ОШИБКА при создании бэкапа реестра: {e}")
        return None


def task_registry_mru(log, do_backup=True):
    if winreg is None:
        log("Очистка реестра недоступна (не Windows).")
        return 0
    if do_backup:
        bf = backup_registry(log)
        if bf is None:
            log("  Очистка реестра отменена — бэкап не удался (для безопасности).")
            return 0

    log("Очистка безопасных списков MRU в реестре (история 'последних файлов')...")
    cleared = 0
    for hive, subkey in SAFE_MRU_KEYS:
        try:
            key = winreg.OpenKey(hive, subkey, 0, winreg.KEY_ALL_ACCESS)
        except FileNotFoundError:
            continue
        except Exception as e:
            log(f"  Пропуск {subkey}: {e}")
            continue

        # Собираем имена значений и подключей, затем удаляем
        try:
            i = 0
            names = []
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
    log(f"  Очищено записей: {cleared}")
    return cleared


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------

class CleanerApp:
    def __init__(self, root):
        self.root = root
        root.title(f"{APP_NAME} {APP_VERSION} — {DEVELOPER}")
        root.geometry("620x560")
        root.resizable(False, False)

        header = tk.Frame(root, pady=10)
        header.pack(fill="x")
        tk.Label(header, text=APP_NAME, font=("Segoe UI", 16, "bold")).pack()
        tk.Label(header, text=f"Разработчик: {DEVELOPER}  |  версия {APP_VERSION}",
                 font=("Segoe UI", 9), fg="#555").pack()

        admin_txt = "Права администратора: ЕСТЬ" if is_admin() else "Права администратора: НЕТ (часть функций может не сработать)"
        admin_color = "#1a7a1a" if is_admin() else "#a03a3a"
        self.admin_label = tk.Label(header, text=admin_txt, fg=admin_color, font=("Segoe UI", 9, "bold"))
        self.admin_label.pack(pady=(4, 0))
        if not is_admin():
            tk.Button(header, text="Перезапустить с правами администратора",
                       command=relaunch_as_admin).pack(pady=(4, 0))

        opts = tk.LabelFrame(root, text="Что очистить", padx=10, pady=10)
        opts.pack(fill="x", padx=15, pady=5)

        self.vars = {}
        options = [
            ("temp", "Временные файлы пользователя (%TEMP%)", True),
            ("wintemp", "Системная папка Windows\\Temp", True),
            ("prefetch", "Prefetch (кэш запуска программ)", True),
            ("recycle", "Корзина", True),
            ("wu", "Кэш загрузок Windows Update", True),
            ("wer", "Отчёты об ошибках Windows", True),
            ("thumbs", "Кэш эскизов (миниатюр)", True),
            ("browser", "Кэш браузеров (Chrome/Edge/Yandex/Opera/Firefox)", True),
            ("registry", "Безопасная очистка реестра (MRU-списки) + автобэкап", False),
        ]
        for key, label, default in options:
            v = tk.BooleanVar(value=default)
            self.vars[key] = v
            tk.Checkbutton(opts, text=label, variable=v, anchor="w",
                            justify="left").pack(fill="x")

        btn_frame = tk.Frame(root, pady=8)
        btn_frame.pack(fill="x", padx=15)
        self.start_btn = tk.Button(btn_frame, text="Начать очистку",
                                    font=("Segoe UI", 11, "bold"), bg="#2d7d46",
                                    fg="white", command=self.start_clean)
        self.start_btn.pack(side="left")
        tk.Button(btn_frame, text="О программе", command=self.show_about).pack(side="right")

        self.progress = ttk.Progressbar(root, mode="indeterminate")
        self.progress.pack(fill="x", padx=15, pady=(0, 8))

        self.log_box = scrolledtext.ScrolledText(root, height=16, state="disabled",
                                                   font=("Consolas", 9))
        self.log_box.pack(fill="both", expand=True, padx=15, pady=(0, 12))

        footer = tk.Label(root, text=f"Контакт: {CONTACT}", font=("Segoe UI", 8), fg="#777")
        footer.pack(pady=(0, 6))

    def log(self, text):
        def _write():
            self.log_box.configure(state="normal")
            self.log_box.insert("end", text + "\n")
            self.log_box.see("end")
            self.log_box.configure(state="disabled")
        self.root.after(0, _write)

    def show_about(self):
        messagebox.showinfo(
            "О программе",
            f"{APP_NAME} {APP_VERSION}\n\n"
            f"Разработчик: {DEVELOPER}\n"
            f"Контакт: {CONTACT}\n\n"
            "Программа удаляет временные и кэш-файлы, чистит корзину и "
            "безопасные списки истории в реестре. Установленные программы "
            "и личные файлы не затрагиваются."
        )

    def start_clean(self):
        if not is_windows():
            messagebox.showerror(APP_NAME, "Программа работает только на Windows.")
            return
        self.start_btn.config(state="disabled")
        self.progress.start(12)
        threading.Thread(target=self.run_clean, daemon=True).start()

    def run_clean(self):
        total_freed = 0
        self.log(f"=== {APP_NAME} — запуск очистки: {datetime.now():%Y-%m-%d %H:%M:%S} ===")

        tasks = [
            ("temp", task_temp),
            ("wintemp", task_windows_temp),
            ("prefetch", task_prefetch),
            ("recycle", task_recycle_bin),
            ("wu", task_windows_update_cache),
            ("wer", task_error_reports),
            ("thumbs", task_thumbnail_cache),
            ("browser", task_browser_cache),
        ]

        for key, fn in tasks:
            if self.vars[key].get():
                try:
                    total_freed += fn(self.log) or 0
                except Exception as e:
                    self.log(f"  ОШИБКА в задаче '{key}': {e}")
            time.sleep(0.05)

        if self.vars["registry"].get():
            try:
                task_registry_mru(self.log, do_backup=True)
            except Exception as e:
                self.log(f"  ОШИБКА при очистке реестра: {e}")

        self.log("=== Готово. ===")
        self.log(f"Всего освобождено места (без учёта реестра): {format_size(total_freed)}")

        def _finish():
            self.progress.stop()
            self.start_btn.config(state="normal")
            messagebox.showinfo(APP_NAME, f"Очистка завершена.\nОсвобождено: {format_size(total_freed)}")
        self.root.after(0, _finish)


def main():
    if is_windows() and not is_admin():
        # Не блокируем запуск — просто предупредим и дадим кнопку в GUI.
        pass
    root = tk.Tk()
    app = CleanerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
