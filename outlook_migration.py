"""
Outlook Migration Tool v3.0.0
==============================
Vollautomatische Outlook Sicherung & Migration.
Läuft auf altem UND neuem PC – erkennt automatisch den Modus.

Autor: WebNetZone
"""

import os
import sys
import ctypes
import shutil
import json
import time
import threading
import subprocess
import winreg
import pathlib
import datetime
import socket
import uuid
import struct
import hashlib
import platform
from tkinter import *
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText

# ═══════════════════════════════════════════════════════════════
# KONSTANTEN
# ═══════════════════════════════════════════════════════════════

APP_TITLE   = "Outlook Migration Tool"
APP_VERSION = "3.0.0"
CONFIG_FILE = "outlook_migration_config.json"
PROGRESS_FILE = os.path.join(os.environ.get("TEMP", "C:\\Temp"), "outlook_migration_progress.json")
SHARE_NAME  = "OutlookMigration"
MIN_PYTHON  = (3, 10)
MIN_RAM_GB  = 4
MIN_CPU_MHZ = 1500

OUTLOOK_VERSIONS = {
    "16.0": "Outlook 2016 / 2019 / 2021 / 365",
    "15.0": "Outlook 2013",
    "14.0": "Outlook 2010",
}

REGISTRY_PROFILE_PATH = (
    r"Software\Microsoft\Windows NT\CurrentVersion"
    r"\Windows Messaging Subsystem\Profiles"
)

ONEDRIVE_PATHS = [
    os.path.join(os.environ.get("USERPROFILE", ""), "OneDrive"),
    os.path.join(os.environ.get("USERPROFILE", ""), "OneDrive - Personal"),
]

# ═══════════════════════════════════════════════════════════════
# FARBEN & STYLE
# ═══════════════════════════════════════════════════════════════

BG_DARK      = "#0f0f1a"
BG_PANEL     = "#1a1a2e"
BG_CARD      = "#16213e"
BG_CARD2     = "#1f2b47"
ACCENT_BLUE  = "#4361ee"
ACCENT_CYAN  = "#4cc9f0"
ACCENT_GREEN = "#06d6a0"
ACCENT_RED   = "#ef233c"
ACCENT_WARN  = "#f8961e"
ACCENT_PURPLE= "#7209b7"
TEXT_WHITE   = "#edf2f4"
TEXT_GRAY    = "#8d99ae"
BTN_HOVER    = "#3a0ca3"

# ═══════════════════════════════════════════════════════════════
# ADMIN & SYSTEM
# ═══════════════════════════════════════════════════════════════

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception:
        return False

def run_as_admin():
    try:
        script = os.path.abspath(sys.argv[0])
        params = " ".join([f'"{a}"' for a in sys.argv[1:]])
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable,
                                             f'"{script}" {params}', None, 1)
    except Exception as e:
        messagebox.showerror("Fehler", f"Admin-Rechte konnten nicht angefordert werden:\n{e}")
    sys.exit(0)

def get_computer_name():
    return socket.gethostname()

def get_ip_address():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return socket.gethostbyname(socket.gethostname())

def get_mac_address():
    try:
        mac = uuid.getnode()
        return ':'.join(('%012X' % mac)[i:i+2] for i in range(0, 12, 2))
    except Exception:
        return ""

def get_system_info():
    """RAM und CPU Info auslesen."""
    info = {}
    try:
        import psutil
        info["ram_gb"] = round(psutil.virtual_memory().total / (1024**3), 1)
        info["cpu_mhz"] = psutil.cpu_freq().current if psutil.cpu_freq() else 0
        info["cpu_cores"] = psutil.cpu_count()
    except Exception:
        info["ram_gb"] = 0
        info["cpu_mhz"] = 0
        info["cpu_cores"] = 0
    return info

def check_internet():
    try:
        socket.setdefaulttimeout(3)
        socket.socket(socket.AF_INET, socket.SOCK_STREAM).connect(("8.8.8.8", 53))
        return True
    except Exception:
        return False

# ═══════════════════════════════════════════════════════════════
# HILFSFUNKTIONEN
# ═══════════════════════════════════════════════════════════════

def format_size(size_bytes):
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 ** 2:
        return f"{size_bytes / 1024:.1f} KB"
    elif size_bytes < 1024 ** 3:
        return f"{size_bytes / 1024**2:.1f} MB"
    else:
        return f"{size_bytes / 1024**3:.2f} GB"

def get_folder_size(path):
    total = 0
    try:
        for dirpath, _, filenames in os.walk(path):
            for f in filenames:
                fp = os.path.join(dirpath, f)
                try:
                    total += os.path.getsize(fp)
                except Exception:
                    pass
    except Exception:
        pass
    return total

def get_free_space(path):
    try:
        return shutil.disk_usage(path).free
    except Exception:
        return 0

def is_in_onedrive(path):
    path_lower = path.lower()
    for od in ONEDRIVE_PATHS:
        if od and path_lower.startswith(od.lower()):
            return True
    return "onedrive" in path_lower

def is_network_path(path):
    return path.startswith("\\\\") or (len(path) > 1 and path[1] != ":")

def is_pst_valid(path):
    try:
        with open(path, "rb") as f:
            header = f.read(4)
            return header == b'!\xBF\xD5\x00' or os.path.getsize(path) > 0
    except Exception:
        return False

def file_checksum(path):
    """MD5 Checksum einer Datei."""
    h = hashlib.md5()
    try:
        with open(path, "rb") as f:
            while chunk := f.read(8192):
                h.update(chunk)
        return h.hexdigest()
    except Exception:
        return None

def get_windows_users():
    users = []
    users_dir = os.path.join(os.environ.get("SystemDrive", "C:") + "\\", "Users")
    try:
        for entry in os.scandir(users_dir):
            if entry.is_dir() and entry.name not in (
                "Public", "Default", "Default User", "All Users"
            ):
                users.append(entry.path)
    except Exception:
        users.append(os.environ.get("USERPROFILE", ""))
    return users

def find_usb_sticks():
    """Alle USB-Sticks finden (FAT32, exFAT, NTFS - alle removable)."""
    usb_drives = []
    seen = set()
    try:
        import psutil
        for disk in psutil.disk_partitions():
            is_removable = (
                "removable" in disk.opts.lower()
                or disk.fstype in ("FAT32", "exFAT", "FAT", "FAT16")
            )
            if not is_removable:
                # Auch NTFS-Sticks per WinAPI prüfen
                try:
                    import ctypes as ct
                    drive_type = ct.windll.kernel32.GetDriveTypeW(disk.mountpoint)
                    if drive_type == 2:  # DRIVE_REMOVABLE
                        is_removable = True
                except Exception:
                    pass
            if is_removable and disk.mountpoint not in seen:
                seen.add(disk.mountpoint)
                usb_drives.append(disk.mountpoint)
    except Exception:
        pass

    # Fallback falls psutil fehlschlägt
    if not usb_drives:
        try:
            import ctypes as ct
            for letter in "DEFGHIJKLMNOPQRSTUVWXYZ":
                drive = f"{letter}:\\"
                if os.path.exists(drive):
                    drive_type = ct.windll.kernel32.GetDriveTypeW(drive)
                    if drive_type == 2:
                        usb_drives.append(drive)
        except Exception:
            pass
    return usb_drives

def find_config_on_usb():
    """Konfig-Datei auf USB-Sticks suchen (Root + Unterordner bis 3 Ebenen). Laufwerksbuchstabe wird automatisch angepasst."""
    for usb in find_usb_sticks():
        # Kandidaten sammeln: Root zuerst, dann Unterordner (max 3 Ebenen)
        candidates = []
        try:
            for root, dirs, files in os.walk(usb):
                depth = root[len(usb):].count(os.sep)
                if depth >= 3:
                    dirs[:] = []
                    continue
                if CONFIG_FILE in files:
                    candidates.append(os.path.join(root, CONFIG_FILE))
        except Exception:
            pass

        for config_path in candidates:
            try:
                with open(config_path, "r", encoding="utf-8") as f:
                    config = json.load(f)

                # Laufwerksbuchstaben im backup_dir auf aktuellen Stick anpassen
                backup_dir = config.get("backup_dir", "")
                if backup_dir and len(backup_dir) >= 2 and backup_dir[1] == ":":
                    usb_drive = os.path.splitdrive(usb)[0]  # z.B. "F:"
                    config["backup_dir"] = usb_drive + backup_dir[2:]

                return config_path, config, usb
            except Exception:
                pass
    return None, None, None

# ═══════════════════════════════════════════════════════════════
# OUTLOOK FUNKTIONEN
# ═══════════════════════════════════════════════════════════════

def detect_outlook_version():
    # 1. HKCU (Outlook konfiguriert)
    for ver in OUTLOOK_VERSIONS:
        try:
            key_path = rf"Software\Microsoft\Office\{ver}\Outlook"
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path):
                return ver, OUTLOOK_VERSIONS[ver]
        except Exception:
            pass

    # 2. HKLM InstallRoot (installiert, aber noch kein Profil)
    for ver in OUTLOOK_VERSIONS:
        for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_LOCAL_MACHINE):
            for wow in (0, winreg.KEY_WOW64_32KEY):
                try:
                    key_path = rf"SOFTWARE\Microsoft\Office\{ver}\Outlook\InstallRoot"
                    with winreg.OpenKey(hive, key_path, 0, winreg.KEY_READ | wow):
                        return ver, OUTLOOK_VERSIONS[ver]
                except Exception:
                    pass

    # 3. Click-to-Run (Microsoft 365)
    try:
        key_path = r"SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path) as k:
            products, _ = winreg.QueryValueEx(k, "ProductReleaseIds")
            if "Outlook" in products or "O365" in products or "Microsoft365" in products:
                return "16.0", OUTLOOK_VERSIONS["16.0"]
    except Exception:
        pass

    # 4. OUTLOOK.EXE auf Disk suchen (klassisches Outlook)
    exe_paths = [
        r"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE",
        r"C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE",
        r"C:\Program Files\Microsoft Office\Office16\OUTLOOK.EXE",
        r"C:\Program Files (x86)\Microsoft Office\Office16\OUTLOOK.EXE",
        r"C:\Program Files\Microsoft Office\root\Office15\OUTLOOK.EXE",
        r"C:\Program Files (x86)\Microsoft Office\Office15\OUTLOOK.EXE",
    ]
    for path in exe_paths:
        if os.path.exists(path):
            ver = "16.0" if "Office16" in path else "15.0"
            return ver, OUTLOOK_VERSIONS.get(ver, "Outlook")

    # 5. Neues Outlook (Store/App-Version, olk.exe)
    local = os.environ.get("LOCALAPPDATA", "")
    olk_data = os.path.join(local, "Microsoft", "Olk")
    if os.path.isdir(olk_data):
        return "16.0", "Neues Outlook (App-Version)"

    # Auch im WindowsApps-Verzeichnis suchen
    for prog in (r"C:\Program Files\WindowsApps", os.path.join(local, "Microsoft", "WindowsApps")):
        try:
            if os.path.isdir(prog):
                for entry in os.listdir(prog):
                    if entry.lower().startswith("microsoft.outlookforwindows"):
                        return "16.0", "Neues Outlook (App-Version)"
        except Exception:
            pass

    # AppX-Paket per PowerShell prüfen (kein Fenster, kein sichtbarer Fehler)
    try:
        CREATE_NO_WINDOW = 0x08000000
        result = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command",
             "$pkg = Get-AppxPackage -Name Microsoft.OutlookForWindows -ErrorAction SilentlyContinue;"
             "if ($pkg) { $pkg.Version }"],
            capture_output=True, text=True, timeout=8,
            creationflags=CREATE_NO_WINDOW
        )
        ver_str = result.stdout.strip()
        if ver_str:
            return "16.0", f"Neues Outlook (App {ver_str})"
    except Exception:
        pass

    return None, None

def close_outlook():
    try:
        import psutil
        def running():
            return any("OUTLOOK" in p.info["name"].upper()
                       for p in psutil.process_iter(["name"])
                       if p.info["name"])

        if not running():
            return True, "Outlook war nicht geöffnet."

        subprocess.run(["taskkill", "/IM", "OUTLOOK.EXE"], capture_output=True, timeout=5)
        for _ in range(10):
            time.sleep(1)
            if not running():
                return True, "Outlook geschlossen."

        subprocess.run(["taskkill", "/F", "/IM", "OUTLOOK.EXE"], capture_output=True, timeout=5)
        time.sleep(2)

        if not running():
            return True, "Outlook zwangsbeendet."
        return False, "Outlook konnte nicht beendet werden!"
    except Exception as e:
        return False, str(e)

def get_active_pst_from_registry(user_profile):
    active_psts = []
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, REGISTRY_PROFILE_PATH) as profiles_key:
            i = 0
            while True:
                try:
                    profile_name = winreg.EnumKey(profiles_key, i)
                    profile_path = REGISTRY_PROFILE_PATH + "\\" + profile_name
                    try:
                        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, profile_path) as pkey:
                            j = 0
                            while True:
                                try:
                                    subkey_name = winreg.EnumKey(pkey, j)
                                    subkey_path = profile_path + "\\" + subkey_name
                                    try:
                                        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, subkey_path) as skey:
                                            try:
                                                val, _ = winreg.QueryValueEx(skey, "001e6700")
                                                if val and val.lower().endswith(".pst"):
                                                    if val not in active_psts:
                                                        active_psts.append(val)
                                            except Exception:
                                                pass
                                    except Exception:
                                        pass
                                    j += 1
                                except OSError:
                                    break
                    except Exception:
                        pass
                    i += 1
                except OSError:
                    break
    except Exception:
        pass
    return active_psts

def find_all_pst_files(user_profile, log_cb=None):
    found = []
    seen = set()
    priority_paths = [
        os.path.join(user_profile, "AppData", "Local", "Microsoft", "Outlook"),
        os.path.join(user_profile, "AppData", "Roaming", "Microsoft", "Outlook"),
        os.path.join(user_profile, "Documents", "Outlook Files"),
        os.path.join(user_profile, "Documents"),
    ]
    all_paths = priority_paths[:]
    for letter in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
        drive = f"{letter}:\\"
        if os.path.exists(drive):
            all_paths.append(drive)

    for search_root in all_paths:
        try:
            for dirpath, dirs, files in os.walk(search_root):
                dirs[:] = [d for d in dirs if d not in (
                    "Windows", "Program Files", "Program Files (x86)",
                    "$Recycle.Bin", "System Volume Information", "ProgramData"
                )]
                for fname in files:
                    if fname.lower().endswith(".pst"):
                        full_path = os.path.join(dirpath, fname)
                        if full_path not in seen:
                            seen.add(full_path)
                            found.append(full_path)
                            if log_cb:
                                log_cb(f"PST gefunden: {full_path}")
        except Exception:
            pass
    return found

def get_outlook_account_info(outlook_version):
    accounts = []
    try:
        key_path = rf"Software\Microsoft\Office\{outlook_version}\Outlook\Profiles"
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path) as profiles:
            i = 0
            while True:
                try:
                    profile = winreg.EnumKey(profiles, i)
                    sub_path = key_path + "\\" + profile
                    try:
                        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, sub_path) as pkey:
                            j = 0
                            while True:
                                try:
                                    account_key = winreg.EnumKey(pkey, j)
                                    acc_path = sub_path + "\\" + account_key
                                    acc_info = {"Profil": profile}
                                    try:
                                        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, acc_path) as akey:
                                            for val_name, friendly in [
                                                ("001e6601", "E-Mail"),
                                                ("001e6602", "Anzeigename"),
                                                ("001e6603", "SMTP Server"),
                                                ("001e6605", "IMAP Server"),
                                                ("001e0028", "Benutzername"),
                                            ]:
                                                try:
                                                    val, _ = winreg.QueryValueEx(akey, val_name)
                                                    acc_info[friendly] = val
                                                except Exception:
                                                    pass
                                    except Exception:
                                        pass
                                    if len(acc_info) > 1:
                                        accounts.append(acc_info)
                                    j += 1
                                except OSError:
                                    break
                    except Exception:
                        pass
                    i += 1
                except OSError:
                    break
    except Exception:
        pass
    return accounts

# ═══════════════════════════════════════════════════════════════
# NETZWERKFREIGABE
# ═══════════════════════════════════════════════════════════════

def create_network_share(share_path, share_name=SHARE_NAME):
    """Netzwerkfreigabe erstellen."""
    try:
        os.makedirs(share_path, exist_ok=True)
        subprocess.run(
            ["net", "share", f"{share_name}={share_path}", "/GRANT:Everyone,READ"],
            capture_output=True, check=True
        )
        return True, f"\\\\{get_computer_name()}\\{share_name}"
    except subprocess.CalledProcessError:
        # Bereits vorhanden → neu erstellen
        try:
            subprocess.run(["net", "share", share_name, "/DELETE"], capture_output=True)
            subprocess.run(
                ["net", "share", f"{share_name}={share_path}", "/GRANT:Everyone,READ"],
                capture_output=True, check=True
            )
            return True, f"\\\\{get_computer_name()}\\{share_name}"
        except Exception as e:
            return False, str(e)
    except Exception as e:
        return False, str(e)

def remove_network_share(share_name=SHARE_NAME):
    """Netzwerkfreigabe entfernen."""
    try:
        subprocess.run(["net", "share", share_name, "/DELETE"], capture_output=True)
        return True
    except Exception:
        return False

def is_host_reachable(host, timeout=3):
    """Prüft ob ein Host im Netzwerk erreichbar ist."""
    try:
        result = subprocess.run(
            ["ping", "-n", "1", "-w", str(timeout * 1000), host],
            capture_output=True, timeout=timeout + 2
        )
        return result.returncode == 0
    except Exception:
        return False

def find_host_by_mac(target_mac):
    """IP-Adresse anhand MAC-Adresse im Netzwerk suchen."""
    try:
        result = subprocess.run(["arp", "-a"], capture_output=True, text=True)
        for line in result.stdout.splitlines():
            if target_mac.lower().replace(":", "-") in line.lower():
                parts = line.split()
                if parts:
                    return parts[0]
    except Exception:
        pass
    return None

# ═══════════════════════════════════════════════════════════════
# PROGRESS & CONFIG
# ═══════════════════════════════════════════════════════════════

def load_progress():
    try:
        if os.path.exists(PROGRESS_FILE):
            with open(PROGRESS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return {}

def save_progress(data):
    try:
        with open(PROGRESS_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

def clear_progress():
    try:
        if os.path.exists(PROGRESS_FILE):
            os.remove(PROGRESS_FILE)
    except Exception:
        pass

def save_config(usb_root, config_data):
    config_path = os.path.join(usb_root, CONFIG_FILE)
    with open(config_path, "w", encoding="utf-8") as f:
        json.dump(config_data, f, ensure_ascii=False, indent=2)
    return config_path

def delete_config_and_data(usb_root):
    """Konfig und alle Backup-Daten vom Stick löschen. Script bleibt."""
    deleted = []
    keep = {"start.bat", "outlook_migration.py", "requirements.txt", "README.md"}
    try:
        for item in os.listdir(usb_root):
            if item in keep:
                continue
            item_path = os.path.join(usb_root, item)
            try:
                if os.path.isdir(item_path):
                    shutil.rmtree(item_path)
                else:
                    os.remove(item_path)
                deleted.append(item)
            except Exception:
                pass
    except Exception:
        pass
    return deleted

# ═══════════════════════════════════════════════════════════════
# DATEI KOPIEREN
# ═══════════════════════════════════════════════════════════════

def copy_with_progress(src, dst, progress_cb=None, cancel_event=None):
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    size = os.path.getsize(src)
    copied = 0
    chunk = 1024 * 1024

    with open(src, "rb") as fsrc, open(dst, "wb") as fdst:
        while True:
            if cancel_event and cancel_event.is_set():
                return False
            buf = fsrc.read(chunk)
            if not buf:
                break
            fdst.write(buf)
            copied += len(buf)
            if progress_cb:
                progress_cb(copied, size)
    return True

def copy_folder(src, dst, log_cb=None, cancel_event=None):
    if not os.path.exists(src):
        return False, f"Nicht gefunden: {src}"
    try:
        os.makedirs(dst, exist_ok=True)
        for item in pathlib.Path(src).rglob("*"):
            if cancel_event and cancel_event.is_set():
                return False, "Abgebrochen"
            rel = item.relative_to(src)
            target = os.path.join(dst, str(rel))
            if item.is_dir():
                os.makedirs(target, exist_ok=True)
            elif item.is_file():
                os.makedirs(os.path.dirname(target), exist_ok=True)
                shutil.copy2(str(item), target)
                if log_cb:
                    log_cb(f"  → {item.name}")
        return True, "OK"
    except Exception as e:
        return False, str(e)

def verify_file(src, dst):
    """Prüft ob kopierte Datei korrekt ist (Größe + Checksum)."""
    try:
        src_size = os.path.getsize(src)
        dst_size = os.path.getsize(dst)
        if src_size != dst_size:
            return False, f"Größe unterschiedlich: {format_size(src_size)} vs {format_size(dst_size)}"
        # Bei großen Dateien nur Größe prüfen (Checksum dauert zu lange)
        if src_size > 500 * 1024 * 1024:
            return True, "OK (Größe stimmt)"
        src_md5 = file_checksum(src)
        dst_md5 = file_checksum(dst)
        if src_md5 != dst_md5:
            return False, "Checksumme stimmt nicht überein"
        return True, "OK"
    except Exception as e:
        return False, str(e)

# ═══════════════════════════════════════════════════════════════
# OUTLOOK IMPORT (NEUER PC)
# ═══════════════════════════════════════════════════════════════

def import_pst_to_outlook(pst_path):
    """PST in Outlook importieren via PowerShell."""
    try:
        ps_script = f'''
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$namespace.AddStoreEx("{pst_path}", 3)
$outlook.Quit()
'''
        result = subprocess.run(
            ["powershell", "-Command", ps_script],
            capture_output=True, text=True, timeout=60
        )
        if result.returncode == 0:
            return True, "PST importiert"
        return False, result.stderr
    except Exception as e:
        return False, str(e)

def restore_signatures(sig_backup, user_profile):
    sig_dst = os.path.join(user_profile, "AppData", "Roaming", "Microsoft", "Signatures")
    return copy_folder(sig_backup, sig_dst)

def restore_rules(rules_backup, user_profile):
    rules_dst = os.path.join(user_profile, "AppData", "Roaming", "Microsoft", "Outlook")
    os.makedirs(rules_dst, exist_ok=True)
    try:
        for f in os.listdir(rules_backup):
            if f.lower().endswith(".rwz"):
                shutil.copy2(os.path.join(rules_backup, f),
                             os.path.join(rules_dst, f))
        return True, "OK"
    except Exception as e:
        return False, str(e)

def restore_settings(roaming_backup, local_backup, user_profile):
    results = []
    if os.path.exists(roaming_backup):
        roaming_dst = os.path.join(user_profile, "AppData", "Roaming", "Microsoft", "Outlook")
        ok, msg = copy_folder(roaming_backup, roaming_dst)
        results.append(("Roaming Einstellungen", ok, msg))
    if os.path.exists(local_backup):
        local_dst = os.path.join(user_profile, "AppData", "Local", "Microsoft", "Outlook")
        ok, msg = copy_folder(local_backup, local_dst)
        results.append(("Local Einstellungen", ok, msg))
    return results


# ═══════════════════════════════════════════════════════════════
# HAUPT GUI
# ═══════════════════════════════════════════════════════════════

class MigrationApp:

    def __init__(self, root):
        self.root = root
        self.root.title(f"{APP_TITLE} v{APP_VERSION}")
        self.root.geometry("960x720")
        self.root.minsize(800, 600)
        self.root.configure(bg=BG_DARK)
        self.root.resizable(True, True)

        # State
        self.mode = None  # "old_pc" oder "new_pc"
        self.usb_root = StringVar()
        self.usb_config = None
        self.config_path = None
        self.user_profile = StringVar(value=os.environ.get("USERPROFILE", ""))
        self.outlook_version = None
        self.outlook_version_name = ""
        self.all_pst_files = []
        self.active_psts = []
        self.pst_vars = {}
        self.pst_group_vars = {}
        self.cancel_event = threading.Event()
        self.start_time = None
        self.results = {"success": [], "warning": [], "error": []}
        self.share_path = None
        self.pst_dest = StringVar(value="")
        self.opt_pst_manual = BooleanVar(value=False)

        self._build_ui()
        self.root.after(500, self._auto_detect)

    # ── UI AUFBAU ──────────────────────────────────────────────

    def _build_ui(self):
        # Header
        hdr = Frame(self.root, bg=ACCENT_BLUE, height=65)
        hdr.pack(fill=X)
        hdr.pack_propagate(False)
        Label(hdr, text=f"  🔄  {APP_TITLE}",
              font=("Segoe UI", 19, "bold"), bg=ACCENT_BLUE, fg="white").pack(side=LEFT, padx=20)
        self.lbl_mode = Label(hdr, text="", font=("Segoe UI", 11),
                               bg=ACCENT_BLUE, fg="#d0e8ff")
        self.lbl_mode.pack(side=RIGHT, padx=20)

        # Notebook
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TNotebook", background=BG_DARK, borderwidth=0)
        style.configure("TNotebook.Tab", background=BG_PANEL, foreground=TEXT_GRAY,
                        padding=[14, 8], font=("Segoe UI", 10))
        style.map("TNotebook.Tab",
                  background=[("selected", BG_CARD)],
                  foreground=[("selected", TEXT_WHITE)])
        style.configure("Custom.Horizontal.TProgressbar",
                        troughcolor=BG_PANEL, background=ACCENT_CYAN, thickness=22)

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=BOTH, expand=True, padx=10, pady=10)

        self.tab_detect   = Frame(self.notebook, bg=BG_DARK)
        self.tab_settings = Frame(self.notebook, bg=BG_DARK)
        self.tab_pst      = Frame(self.notebook, bg=BG_DARK)
        self.tab_run      = Frame(self.notebook, bg=BG_DARK)
        self.tab_result   = Frame(self.notebook, bg=BG_DARK)

        self.notebook.add(self.tab_detect,   text="🔍  Erkennung")
        self.notebook.add(self.tab_settings, text="⚙️  Einstellungen")
        self.notebook.add(self.tab_pst,      text="📁  PST Dateien")
        self.notebook.add(self.tab_run,      text="🚀  Ausführen")
        self.notebook.add(self.tab_result,   text="✅  Ergebnis")

        self._build_detect_tab()
        self._build_settings_tab()
        self._build_pst_tab()
        self._build_run_tab()
        self._build_result_tab()

    def _card(self, parent, title, color=None):
        c = color or BG_CARD
        frame = Frame(parent, bg=c, bd=0)
        frame.pack(fill=X, padx=15, pady=6)
        if title:
            Label(frame, text=title, font=("Segoe UI", 10, "bold"),
                  bg=c, fg=ACCENT_CYAN).pack(anchor=W, padx=12, pady=(10, 3))
        inner = Frame(frame, bg=c)
        inner.pack(fill=X, padx=12, pady=(0, 10))
        return inner

    def _btn(self, parent, text, cmd, color=ACCENT_BLUE, width=18):
        b = Button(parent, text=text, command=cmd, bg=color, fg="white",
                   font=("Segoe UI", 10, "bold"), relief=FLAT, bd=0,
                   activebackground=BTN_HOVER, activeforeground="white",
                   cursor="hand2", width=width, pady=7)
        b.bind("<Enter>", lambda e: b.config(bg=BTN_HOVER if color != ACCENT_RED else "#c0001a"))
        b.bind("<Leave>", lambda e: b.config(bg=color))
        return b

    def _info_row(self, parent, label, value, val_color=TEXT_WHITE):
        row = Frame(parent, bg=parent["bg"])
        row.pack(fill=X, pady=2)
        Label(row, text=label, width=22, anchor=W, font=("Segoe UI", 10),
              bg=parent["bg"], fg=TEXT_GRAY).pack(side=LEFT)
        Label(row, text=value, font=("Segoe UI", 10, "bold"),
              bg=parent["bg"], fg=val_color).pack(side=LEFT)

    def spacer(self, parent, h=8):
        Frame(parent, bg=BG_DARK, height=h).pack()

    # ── TAB 1: ERKENNUNG ──────────────────────────────────────

    def _build_detect_tab(self):
        Label(self.tab_detect, text="Automatische Erkennung",
              font=("Segoe UI", 14, "bold"), bg=BG_DARK, fg=TEXT_WHITE).pack(pady=(20, 5))
        Label(self.tab_detect, text="Das Script erkennt automatisch ob es auf dem alten oder neuen PC läuft.",
              font=("Segoe UI", 10), bg=BG_DARK, fg=TEXT_GRAY).pack()

        self.spacer(self.tab_detect, 15)

        detect_card = self._card(self.tab_detect, "🔍 Erkennung läuft...")
        self.lbl_detect_status = Label(detect_card, text="Suche USB-Stick mit Konfig...",
                                        font=("Segoe UI", 11), bg=BG_CARD, fg=TEXT_GRAY)
        self.lbl_detect_status.pack(anchor=W, pady=5)

        self.detect_info = Frame(detect_card, bg=BG_CARD)
        self.detect_info.pack(fill=X, pady=5)

        self.spacer(self.tab_detect)

        # Manuelle Auswahl
        manual_card = self._card(self.tab_detect, "✋ Manuell auswählen")
        btn_row = Frame(manual_card, bg=BG_CARD)
        btn_row.pack(anchor=W)
        self._btn(btn_row, "🖥️ Alter PC Modus",
                  lambda: self._set_mode("old_pc"), color=ACCENT_PURPLE, width=20).pack(side=LEFT, padx=(0, 10))
        self._btn(btn_row, "💻 Neuer PC Modus",
                  lambda: self._set_mode("new_pc"), color=ACCENT_GREEN, width=20).pack(side=LEFT)

    # ── TAB 2: EINSTELLUNGEN ──────────────────────────────────

    def _build_settings_tab(self):
        canvas = Canvas(self.tab_settings, bg=BG_DARK, highlightthickness=0)
        sb = ttk.Scrollbar(self.tab_settings, orient=VERTICAL, command=canvas.yview)
        self.sf = Frame(canvas, bg=BG_DARK)
        self.sf.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.sf, anchor=NW)
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side=LEFT, fill=BOTH, expand=True)
        sb.pack(side=RIGHT, fill=Y)

        # Outlook Info
        ol_card = self._card(self.sf, "🔍 Outlook")
        self.lbl_outlook = Label(ol_card, text="Wird erkannt...",
                                  font=("Segoe UI", 10), bg=BG_CARD, fg=TEXT_GRAY)
        self.lbl_outlook.pack(anchor=W)

        # System Info
        sys_card = self._card(self.sf, "🖥️ System Info")
        self.lbl_sysinfo = Label(sys_card, text="Wird geladen...",
                                  font=("Segoe UI", 10), bg=BG_CARD, fg=TEXT_GRAY)
        self.lbl_sysinfo.pack(anchor=W)

        # Benutzer
        user_card = self._card(self.sf, "👤 Benutzerprofil")
        self.user_combo = ttk.Combobox(user_card, textvariable=self.user_profile,
                                        font=("Segoe UI", 10), width=55)
        self.user_combo.pack(anchor=W, pady=4)

        # USB / Ziel
        usb_card = self._card(self.sf, "💾 USB-Stick / Zielordner")
        usb_row = Frame(usb_card, bg=BG_CARD)
        usb_row.pack(fill=X)
        Entry(usb_row, textvariable=self.usb_root, font=("Segoe UI", 10),
              bg=BG_PANEL, fg=TEXT_WHITE, insertbackground=TEXT_WHITE,
              relief=FLAT, bd=5, width=45).pack(side=LEFT, padx=(0, 8))
        self._btn(usb_row, "📂 Wählen", self._choose_usb, width=10).pack(side=LEFT)
        self.lbl_usb_space = Label(usb_card, text="", font=("Segoe UI", 9),
                                    bg=BG_CARD, fg=TEXT_GRAY)
        self.lbl_usb_space.pack(anchor=W, pady=3)

        # PST Zielordner
        pst_card = self._card(self.sf, "📁 PST Zielordner (optional)")
        Label(pst_card, text="Ordner wählen wohin PST-Dateien kopiert werden sollen:",
              font=("Segoe UI", 9), bg=BG_CARD, fg=TEXT_GRAY).pack(anchor=W, pady=(0, 4))
        pst_row = Frame(pst_card, bg=BG_CARD)
        pst_row.pack(fill=X)
        Entry(pst_row, textvariable=self.pst_dest, font=("Segoe UI", 10),
              bg=BG_PANEL, fg=TEXT_WHITE, insertbackground=TEXT_WHITE,
              relief=FLAT, bd=5, width=45).pack(side=LEFT, padx=(0, 8))
        self._btn(pst_row, "📂 Wählen", self._choose_pst_dest, width=10).pack(side=LEFT)
        self._btn(pst_row, "✖ Leer", lambda: self.pst_dest.set(""), width=6).pack(side=LEFT, padx=(4, 0))
        Label(pst_card, text="Leer lassen = PST-Dateien werden nicht extra kopiert",
              font=("Segoe UI", 8), bg=BG_CARD, fg=TEXT_GRAY).pack(anchor=W, pady=(3, 4))
        Checkbutton(pst_card,
                    text="📋  Manuell kopieren (Script legt nur Ordnerstruktur + Pfade-Liste an)",
                    variable=self.opt_pst_manual,
                    bg=BG_CARD, fg=TEXT_WHITE, selectcolor=BG_PANEL,
                    activebackground=BG_CARD, font=("Segoe UI", 10)).pack(anchor=W)

        # Optionen
        opt_card = self._card(self.sf, "⚙️ Optionen")
        self.opt_signatures = BooleanVar(value=True)
        self.opt_settings   = BooleanVar(value=True)
        self.opt_rules      = BooleanVar(value=True)
        self.opt_accounts   = BooleanVar(value=True)

        for text, var in [
            ("✍️  Signaturen sichern", self.opt_signatures),
            ("⚙️  Outlook Einstellungen sichern", self.opt_settings),
            ("📋  Regeln sichern", self.opt_rules),
            ("📧  Kontodaten auslesen", self.opt_accounts),
        ]:
            Checkbutton(opt_card, text=text, variable=var,
                        bg=BG_CARD, fg=TEXT_WHITE, selectcolor=BG_PANEL,
                        activebackground=BG_CARD, font=("Segoe UI", 10)).pack(anchor=W, pady=2)

        Frame(self.sf, bg=BG_DARK, height=10).pack()
        self._btn(self.sf, "Weiter →", self._goto_pst, width=20).pack(pady=8)

    # ── TAB 3: PST ────────────────────────────────────────────

    def _build_pst_tab(self):
        top = Frame(self.tab_pst, bg=BG_DARK)
        top.pack(fill=X, padx=15, pady=10)
        Label(top, text="PST-Dateien auswählen",
              font=("Segoe UI", 13, "bold"), bg=BG_DARK, fg=TEXT_WHITE).pack(side=LEFT)
        btn_r = Frame(top, bg=BG_DARK)
        btn_r.pack(side=RIGHT)
        self._btn(btn_r, "🔄 Neu suchen", self._scan_pst, width=13).pack(side=LEFT, padx=4)
        self._btn(btn_r, "➕ Manuell", self._add_pst_manual, width=11).pack(side=LEFT)

        leg = Frame(self.tab_pst, bg=BG_DARK)
        leg.pack(fill=X, padx=15)
        for col, txt in [(ACCENT_GREEN, "✅ Aktiv"), (ACCENT_WARN, "⚠️ Inaktiv"),
                          (ACCENT_RED, "❌ Problem")]:
            Label(leg, text=txt, font=("Segoe UI", 9), bg=BG_DARK, fg=col).pack(side=LEFT, padx=8)

        lf = Frame(self.tab_pst, bg=BG_PANEL)
        lf.pack(fill=BOTH, expand=True, padx=15, pady=5)
        self.pst_canvas = Canvas(lf, bg=BG_PANEL, highlightthickness=0)
        pst_sb = ttk.Scrollbar(lf, orient=VERTICAL, command=self.pst_canvas.yview)
        self.pst_inner = Frame(self.pst_canvas, bg=BG_PANEL)
        self.pst_inner.bind("<Configure>",
            lambda e: self.pst_canvas.configure(scrollregion=self.pst_canvas.bbox("all")))
        self.pst_canvas.create_window((0, 0), window=self.pst_inner, anchor=NW)
        self.pst_canvas.configure(yscrollcommand=pst_sb.set)
        self.pst_canvas.pack(side=LEFT, fill=BOTH, expand=True)
        pst_sb.pack(side=RIGHT, fill=Y)

        self.lbl_pst_info = Label(self.tab_pst, text="",
                                   font=("Segoe UI", 9), bg=BG_DARK, fg=TEXT_GRAY)
        self.lbl_pst_info.pack(pady=4)
        self._btn(self.tab_pst, "Weiter →", self._goto_run, width=20).pack(pady=8)

    def _render_pst(self):
        for w in self.pst_inner.winfo_children():
            w.destroy()
        if not self.all_pst_files:
            box = Frame(self.pst_inner, bg=BG_PANEL)
            box.pack(pady=30, padx=20, anchor=W)
            Label(box, text="Keine PST-Dateien gefunden.",
                  font=("Segoe UI", 11, "bold"), bg=BG_PANEL, fg=TEXT_GRAY).pack(anchor=W, pady=(0, 6))
            Label(box, text="Wo befinden sich die PST-Dateien?",
                  font=("Segoe UI", 10), bg=BG_PANEL, fg=TEXT_WHITE).pack(anchor=W, pady=(0, 12))
            bf = Frame(box, bg=BG_PANEL)
            bf.pack(anchor=W)
            self._btn(bf, "📂 Ordner durchsuchen",
                      self._browse_folder_for_pst, color=ACCENT_BLUE, width=22).pack(side=LEFT, padx=(0, 8))
            self._btn(bf, "📄 PST-Datei(en) direkt wählen",
                      self._add_pst_manual, color=ACCENT_GREEN, width=26).pack(side=LEFT)
            return

        # Group PSTs by source location
        groups = {}
        for path in self.all_pst_files:
            key = self._pst_subfolder(path)
            groups.setdefault(key, []).append(path)

        for group_name, paths in groups.items():
            # Pre-compute per-path state and init pst_vars
            checkable_paths = []
            for path in paths:
                in_od = is_in_onedrive(path)
                is_valid = is_pst_valid(path)
                if path not in self.pst_vars:
                    self.pst_vars[path] = BooleanVar(value=is_valid and not in_od)
                if not in_od and is_valid:
                    checkable_paths.append(path)

            # Init group var (all checkable files checked → group checked)
            if group_name not in self.pst_group_vars:
                all_on = bool(checkable_paths) and all(self.pst_vars[p].get() for p in checkable_paths)
                self.pst_group_vars[group_name] = BooleanVar(value=all_on)

            # Group header row
            hdr = Frame(self.pst_inner, bg=BG_DARK, pady=4)
            hdr.pack(fill=X, padx=5, pady=(10, 1))
            Checkbutton(hdr, variable=self.pst_group_vars[group_name], bg=BG_DARK,
                        activebackground=BG_DARK, selectcolor=BG_PANEL,
                        state=NORMAL if checkable_paths else DISABLED,
                        command=lambda gn=group_name, cp=checkable_paths: self._toggle_pst_group(gn, cp)
                        ).pack(side=LEFT, padx=(5, 0))
            count_txt = f"{len(paths)} Datei{'en' if len(paths) != 1 else ''}"
            Label(hdr, text=f"📁  {group_name}   ({count_txt})",
                  font=("Segoe UI", 10, "bold"), bg=BG_DARK, fg=ACCENT_CYAN).pack(side=LEFT, padx=6)

            # Individual file rows (indented)
            for path in paths:
                is_active = path in self.active_psts
                in_od = is_in_onedrive(path)
                is_net = is_network_path(path)
                is_valid = is_pst_valid(path)

                if in_od:
                    sc, st, can = ACCENT_RED, "⚠️ OneDrive", False
                elif is_net:
                    sc, st, can = ACCENT_WARN, "🌐 Netzwerk", True
                elif not is_valid:
                    sc, st, can = ACCENT_RED, "❌ Beschädigt", False
                elif is_active:
                    sc, st, can = ACCENT_GREEN, "✅ Aktiv", True
                else:
                    sc, st, can = ACCENT_WARN, "⚠️ Inaktiv", True

                row = Frame(self.pst_inner, bg=BG_CARD2, pady=4)
                row.pack(fill=X, padx=(30, 5), pady=2)

                Checkbutton(row, variable=self.pst_vars[path], bg=BG_CARD2,
                            activebackground=BG_CARD2, selectcolor=BG_PANEL,
                            state=NORMAL if can else DISABLED).pack(side=LEFT, padx=8)

                inf = Frame(row, bg=BG_CARD2)
                inf.pack(side=LEFT, fill=X, expand=True)
                Label(inf, text=os.path.basename(path),
                      font=("Segoe UI", 10, "bold"), bg=BG_CARD2, fg=TEXT_WHITE).pack(anchor=W)
                Label(inf, text=path, font=("Segoe UI", 8),
                      bg=BG_CARD2, fg=TEXT_GRAY).pack(anchor=W)

                rr = Frame(row, bg=BG_CARD2)
                rr.pack(side=RIGHT, padx=10)
                try:
                    sz = format_size(os.path.getsize(path))
                except Exception:
                    sz = "?"
                Label(rr, text=sz, font=("Segoe UI", 9), bg=BG_CARD2, fg=TEXT_GRAY).pack()
                Label(rr, text=st, font=("Segoe UI", 9, "bold"), bg=BG_CARD2, fg=sc).pack()

                if in_od:
                    Button(row, text="Lösung", font=("Segoe UI", 8), bg=ACCENT_WARN, fg="white",
                           relief=FLAT, command=lambda p=path: self._fix_onedrive(p)).pack(side=RIGHT, padx=5)
                if is_net:
                    Button(row, text="Verbinden", font=("Segoe UI", 8), bg=ACCENT_BLUE, fg="white",
                           relief=FLAT, command=lambda p=path: self._connect_net(p)).pack(side=RIGHT, padx=5)

        active_count = sum(1 for p in self.active_psts if p in self.all_pst_files)
        self.lbl_pst_info.config(
            text=f"{len(self.all_pst_files)} PST(s) gefunden | {active_count} aktiv in Outlook"
        )

    def _toggle_pst_group(self, group_name, checkable_paths):
        new_val = self.pst_group_vars[group_name].get()
        for path in checkable_paths:
            if path in self.pst_vars:
                self.pst_vars[path].set(new_val)

    # ── TAB 4: AUSFÜHREN ──────────────────────────────────────

    def _build_run_tab(self):
        # Buttons zuerst mit side=BOTTOM packen damit sie immer sichtbar bleiben
        bf = Frame(self.tab_run, bg=BG_DARK)
        bf.pack(side=BOTTOM, pady=12)
        self.btn_start = self._btn(bf, "🚀 Sicherung starten", self._start, color=ACCENT_GREEN, width=20)
        self.btn_start.pack(side=LEFT, padx=8)
        self.btn_cancel = self._btn(bf, "⛔ Abbrechen", self._cancel, color=ACCENT_RED, width=14)
        self.btn_cancel.pack(side=LEFT, padx=8)
        self.btn_cancel.config(state=DISABLED)

        sum_c = self._card(self.tab_run, "📊 Zusammenfassung")
        self.lbl_summary = Label(sum_c, text="Bitte Einstellungen und PST-Dateien zuerst auswählen.",
                                  font=("Segoe UI", 10), bg=BG_CARD, fg=TEXT_GRAY, justify=LEFT)
        self.lbl_summary.pack(anchor=W)

        prog_c = self._card(self.tab_run, "⏳ Fortschritt")
        self.lbl_current = Label(prog_c, text="", font=("Segoe UI", 10),
                                  bg=BG_CARD, fg=TEXT_WHITE)
        self.lbl_current.pack(anchor=W, pady=2)
        self.progress_var = DoubleVar()
        ttk.Progressbar(prog_c, variable=self.progress_var, maximum=100,
                         style="Custom.Horizontal.TProgressbar").pack(fill=X, pady=5)
        self.lbl_pct = Label(prog_c, text="0%", font=("Segoe UI", 10, "bold"),
                              bg=BG_CARD, fg=ACCENT_CYAN)
        self.lbl_pct.pack()
        self.lbl_speed = Label(prog_c, text="", font=("Segoe UI", 9),
                                bg=BG_CARD, fg=TEXT_GRAY)
        self.lbl_speed.pack(anchor=W)

        log_c = self._card(self.tab_run, "📋 Log")
        self.log_text = ScrolledText(log_c, height=8, bg=BG_PANEL, fg=TEXT_WHITE,
                                      font=("Consolas", 9), relief=FLAT,
                                      insertbackground=TEXT_WHITE)
        self.log_text.pack(fill=X)
        self.log_text.config(state=DISABLED)

    # ── TAB 5: ERGEBNIS ───────────────────────────────────────

    def _build_result_tab(self):
        self.result_inner = Frame(self.tab_result, bg=BG_DARK)
        self.result_inner.pack(fill=BOTH, expand=True, padx=15, pady=15)
        Label(self.result_inner, text="Noch kein Vorgang durchgeführt.",
              font=("Segoe UI", 12), bg=BG_DARK, fg=TEXT_GRAY).pack(pady=40)

    def _update_result(self):
        for w in self.result_inner.winfo_children():
            w.destroy()
        elapsed = int(time.time() - self.start_time) if self.start_time else 0
        mins, secs = divmod(elapsed, 60)

        Label(self.result_inner, text="Vorgang abgeschlossen!",
              font=("Segoe UI", 16, "bold"), bg=BG_DARK, fg=ACCENT_GREEN).pack(pady=(10, 4))
        Label(self.result_inner,
              text=f"Dauer: {mins}m {secs}s  |  Modus: {'Alter PC' if self.mode == 'old_pc' else 'Neuer PC'}",
              font=("Segoe UI", 10), bg=BG_DARK, fg=TEXT_GRAY).pack(pady=(0, 15))

        for section, icon, color, items in [
            ("Erfolgreich", "✅", ACCENT_GREEN, self.results["success"]),
            ("Warnungen",   "⚠️", ACCENT_WARN,  self.results["warning"]),
            ("Fehler",      "❌", ACCENT_RED,   self.results["error"]),
        ]:
            if not items:
                continue
            card = Frame(self.result_inner, bg=BG_CARD)
            card.pack(fill=X, pady=4)
            Label(card, text=f"{icon}  {section}",
                  font=("Segoe UI", 11, "bold"), bg=BG_CARD, fg=color).pack(anchor=W, padx=12, pady=(8, 2))
            for item in items:
                Label(card, text=f"   → {item}", font=("Segoe UI", 10),
                      bg=BG_CARD, fg=TEXT_WHITE).pack(anchor=W, padx=12)
            Frame(card, bg=BG_CARD, height=8).pack()

        # Bericht auf Stick speichern
        if self.usb_root.get():
            self._save_report()

        bf = Frame(self.result_inner, bg=BG_DARK)
        bf.pack(pady=15)
        if self.usb_root.get() and os.path.exists(self.usb_root.get()):
            self._btn(bf, "📂 Stick öffnen",
                      lambda: os.startfile(self.usb_root.get()), width=16).pack(side=LEFT, padx=6)
        self._btn(bf, "🔄 Neu starten", self._reset, color=ACCENT_WARN, width=14).pack(side=LEFT, padx=6)

    def _save_report(self):
        try:
            report_path = os.path.join(self.usb_root.get(),
                f"Bericht_{datetime.datetime.now().strftime('%Y-%m-%d_%H-%M')}.txt")
            with open(report_path, "w", encoding="utf-8") as f:
                f.write(f"Outlook Migration Tool - Bericht\n")
                f.write(f"{'='*50}\n")
                f.write(f"Datum: {datetime.datetime.now().strftime('%d.%m.%Y %H:%M')}\n")
                f.write(f"Modus: {'Alter PC' if self.mode == 'old_pc' else 'Neuer PC'}\n\n")
                for section, items in self.results.items():
                    if items:
                        f.write(f"{section.upper()}:\n")
                        for item in items:
                            f.write(f"  → {item}\n")
                        f.write("\n")
        except Exception:
            pass

    # ── ERKENNUNG ─────────────────────────────────────────────

    def _auto_detect(self):
        self.lbl_detect_status.config(text="🔍 Suche USB-Stick...")

        def detect():
            try:
                config_path, config, usb = find_config_on_usb()
                if config:
                    self.usb_config = config
                    self.config_path = config_path
                    self.root.after(0, lambda: self._found_config(config, usb))
                else:
                    hw_path, usb = find_hardware_info_on_usb()
                    if hw_path:
                        self.root.after(0, lambda: self._hardware_already_scanned(hw_path, usb))
                    else:
                        self.root.after(0, lambda: self._ask_hardware_scan())
            except Exception as e:
                self.root.after(0, lambda: self.lbl_detect_status.config(
                    text=f"❌ Fehler bei Erkennung: {e}", fg=ACCENT_RED))

        threading.Thread(target=detect, daemon=True).start()

    def _found_config(self, config, usb):
        self.lbl_detect_status.config(text="✅ Konfig gefunden!", fg=ACCENT_GREEN)

        for w in self.detect_info.winfo_children():
            w.destroy()

        self._info_row(self.detect_info, "Alter PC:", config.get("computer_name", "?"))
        self._info_row(self.detect_info, "IP-Adresse:", config.get("ip_address", "?"))
        self._info_row(self.detect_info, "Backup Datum:", config.get("backup_date", "?"))
        self._info_row(self.detect_info, "USB-Stick:", usb)

        Frame(self.detect_info, bg=BG_CARD, height=10).pack()

        Label(self.detect_info, text="Was möchtest du tun?",
              font=("Segoe UI", 11, "bold"), bg=BG_CARD, fg=TEXT_WHITE).pack(anchor=W, pady=(5, 8))

        bf = Frame(self.detect_info, bg=BG_CARD)
        bf.pack(anchor=W)
        self._btn(bf, "📥 Kopieren & Importieren",
                  lambda: self._start_new_pc(usb, config), color=ACCENT_GREEN, width=24).pack(side=LEFT, padx=(0, 10))
        self._btn(bf, "🗑️ Stick zurücksetzen",
                  lambda: self._confirm_delete(usb), color=ACCENT_RED, width=20).pack(side=LEFT)

    def _no_config_found(self):
        self.lbl_detect_status.config(text="ℹ️ Kein Konfig gefunden → Alter PC Modus", fg=ACCENT_CYAN)
        self._set_mode("old_pc")

    def _ask_hardware_scan(self):
        """Abfrage: Hardware scannen oder direkt Outlook Backup?"""
        self.lbl_detect_status.config(text="🔍 Kein Konfig & kein Hardware-Scan gefunden.", fg=ACCENT_WARN)

        for w in self.detect_info.winfo_children():
            w.destroy()

        Label(self.detect_info, text="Was möchtest du tun?",
              font=("Segoe UI", 11, "bold"), bg=BG_CARD, fg=TEXT_WHITE).pack(anchor=W, pady=(5, 10))

        bf = Frame(self.detect_info, bg=BG_CARD)
        bf.pack(anchor=W)
        self._btn(bf, "🖥️ Hardware scannen",
                  self._run_hardware_scan, color=ACCENT_PURPLE, width=20).pack(side=LEFT, padx=(0, 10))
        self._btn(bf, "📦 Outlook Backup starten",
                  lambda: self._set_mode("old_pc"), color=ACCENT_BLUE, width=22).pack(side=LEFT)

        bf2 = Frame(self.detect_info, bg=BG_CARD)
        bf2.pack(anchor=W, pady=(8, 0))
        self._btn(bf2, "📂 USB-Stick manuell auswählen",
                  self._manual_usb_select, color=ACCENT_WARN, width=28).pack(side=LEFT)

    def _manual_usb_select(self):
        """USB-Stick oder Backup-Ordner manuell auswählen und nach Konfig suchen."""
        path = filedialog.askdirectory(title="USB-Stick / Backup-Ordner auswählen")
        if not path:
            return

        self.lbl_detect_status.config(text="🔍 Suche Konfig-Datei...", fg=ACCENT_CYAN)

        # Config in ausgewähltem Ordner + Unterordner suchen (max 3 Ebenen)
        found_config_path = None
        try:
            for root, dirs, files in os.walk(path):
                depth = root[len(path):].count(os.sep)
                if depth >= 3:
                    dirs[:] = []
                    continue
                if CONFIG_FILE in files:
                    found_config_path = os.path.join(root, CONFIG_FILE)
                    break
        except Exception:
            pass

        if found_config_path:
            try:
                with open(found_config_path, "r", encoding="utf-8") as f:
                    config = json.load(f)
                # Laufwerksbuchstaben anpassen
                backup_dir = config.get("backup_dir", "")
                if backup_dir and len(backup_dir) >= 2 and backup_dir[1] == ":":
                    usb_drive = os.path.splitdrive(path)[0]
                    config["backup_dir"] = usb_drive + backup_dir[2:]
                self.usb_config = config
                self.config_path = found_config_path
                self._found_config(config, path)
                return
            except Exception as e:
                messagebox.showerror("Fehler", f"Konfig-Datei konnte nicht gelesen werden:\n{e}")
                return

        # Keine Konfig gefunden – trotzdem als USB-Stammordner setzen
        self.lbl_detect_status.config(
            text="⚠️ Keine Konfig gefunden – Ordner als Backup-Ziel gesetzt.", fg=ACCENT_WARN)
        self.usb_root.set(path)
        for w in self.detect_info.winfo_children():
            w.destroy()
        Label(self.detect_info,
              text=f"Ausgewählter Ordner: {path}",
              font=("Segoe UI", 9), bg=BG_CARD, fg=TEXT_GRAY).pack(anchor=W, pady=(0, 8))
        bf = Frame(self.detect_info, bg=BG_CARD)
        bf.pack(anchor=W)
        self._btn(bf, "📦 Outlook Backup starten",
                  lambda: self._set_mode("old_pc"), color=ACCENT_BLUE, width=22).pack(side=LEFT, padx=(0, 10))
        self._btn(bf, "📂 Anderen Ordner wählen",
                  self._manual_usb_select, color=ACCENT_WARN, width=22).pack(side=LEFT)

    def _hardware_already_scanned(self, hw_path, usb):
        """Hardware bereits gescannt – direkt Alter PC Modus + Option erneut scannen."""
        self.lbl_detect_status.config(text="✅ Hardware bereits gescannt!", fg=ACCENT_GREEN)
        self.usb_root.set(usb)

        for w in self.detect_info.winfo_children():
            w.destroy()

        Label(self.detect_info, text=f"Hardware_Info.txt gefunden:",
              font=("Segoe UI", 10), bg=BG_CARD, fg=TEXT_GRAY).pack(anchor=W)
        Label(self.detect_info, text=hw_path,
              font=("Segoe UI", 9), bg=BG_CARD, fg=TEXT_WHITE).pack(anchor=W, pady=(0, 8))

        bf = Frame(self.detect_info, bg=BG_CARD)
        bf.pack(anchor=W)
        self._btn(bf, "📦 Outlook Backup starten",
                  lambda: self._set_mode("old_pc"), color=ACCENT_GREEN, width=22).pack(side=LEFT, padx=(0, 10))
        self._btn(bf, "🔄 Erneut scannen",
                  self._run_hardware_scan, color=ACCENT_WARN, width=16).pack(side=LEFT)

    def _run_hardware_scan(self):
        """Hardware Scan durchführen und speichern."""
        self.lbl_detect_status.config(text="🔍 Scanne Hardware...", fg=ACCENT_CYAN)

        for w in self.detect_info.winfo_children():
            w.destroy()

        self.scan_log = Label(self.detect_info, text="Bitte warten...",
                               font=("Segoe UI", 10), bg=BG_CARD, fg=TEXT_GRAY, justify=LEFT)
        self.scan_log.pack(anchor=W, pady=5)

        self.scan_progress = ttk.Progressbar(self.detect_info, mode="indeterminate",
                                              style="Custom.Horizontal.TProgressbar")
        self.scan_progress.pack(fill=X, pady=5)
        self.scan_progress.start(10)

        def do_scan():
            try:
                hardware, errors = scan_hardware()
                usbs = find_usb_sticks()
                if usbs:
                    save_dir = os.path.join(usbs[0], "Treiber")
                    self.usb_root.set(usbs[0])
                    save_path = os.path.join(save_dir, HARDWARE_INFO_FILE)
                    try:
                        saved = save_hardware_info(hardware, errors, save_path)
                        self.root.after(0, lambda: self._scan_done(hardware, errors, saved))
                    except Exception:
                        self.root.after(0, lambda: self._ask_alt_save_location(hardware, errors))
                else:
                    self.root.after(0, lambda: self._ask_alt_save_location(hardware, errors))
            except Exception as e:
                self.root.after(0, lambda: self._stop_scan_on_error(str(e)))

        threading.Thread(target=do_scan, daemon=True).start()

    def _stop_scan_on_error(self, error_msg):
        try:
            self.scan_progress.stop()
            self.scan_progress.pack_forget()
        except Exception:
            pass
        self.lbl_detect_status.config(text="❌ Fehler beim Hardware-Scan!", fg=ACCENT_RED)
        messagebox.showerror("Fehler", f"Hardware-Scan fehlgeschlagen:\n{error_msg}")

    def _ask_alt_save_location(self, hardware, errors):
        """Alternativen Speicherort vorschlagen wenn USB nicht beschreibbar."""
        try:
            self.scan_progress.stop()
        except Exception:
            pass
        path = filedialog.askdirectory(title="Speicherort für Hardware_Info.txt wählen")
        if path:
            save_path = os.path.join(path, "Treiber", HARDWARE_INFO_FILE)
            try:
                saved = save_hardware_info(hardware, errors, save_path)
                self._scan_done(hardware, errors, saved)
            except Exception as e:
                messagebox.showerror("Fehler", "Konnte nicht speichern:\n" + str(e))
        else:
            self.lbl_detect_status.config(text="⚠️ Scan abgebrochen.", fg=ACCENT_WARN)

    def _scan_done(self, hardware, errors, saved_path):
        """Hardware Scan abgeschlossen."""
        self.scan_progress.stop()
        self.scan_progress.pack_forget()

        self.lbl_detect_status.config(text="✅ Hardware Scan abgeschlossen!", fg=ACCENT_GREEN)

        for w in self.detect_info.winfo_children():
            w.destroy()

        # Gefundene Komponenten anzeigen
        Label(self.detect_info, text="Gefundene Hardware:",
              font=("Segoe UI", 10, "bold"), bg=BG_CARD, fg=TEXT_WHITE).pack(anchor=W, pady=(0, 5))

        for component, value in hardware.items():
            row = Frame(self.detect_info, bg=BG_CARD)
            row.pack(fill=X, pady=1)
            Label(row, text=f"{component}:", width=16, anchor=W,
                  font=("Segoe UI", 9), bg=BG_CARD, fg=TEXT_GRAY).pack(side=LEFT)
            Label(row, text=value, font=("Segoe UI", 9, "bold"),
                  bg=BG_CARD, fg=TEXT_WHITE).pack(side=LEFT)

        if errors:
            Label(self.detect_info,
                  text=f"⚠️ {len(errors)} Komponente(n) nicht auslesbar – trotzdem gespeichert.",
                  font=("Segoe UI", 9), bg=BG_CARD, fg=ACCENT_WARN).pack(anchor=W, pady=4)

        Label(self.detect_info, text=f"💾 Gespeichert: {saved_path}",
              font=("Segoe UI", 9), bg=BG_CARD, fg=TEXT_GRAY).pack(anchor=W, pady=(4, 8))

        Label(self.detect_info,
              text="Nächster Schritt: Stick zum alten PC und Outlook Backup starten.",
              font=("Segoe UI", 10, "bold"), bg=BG_CARD, fg=ACCENT_GREEN).pack(anchor=W)

        bf = Frame(self.detect_info, bg=BG_CARD)
        bf.pack(anchor=W, pady=8)
        self._btn(bf, "📂 Ordner öffnen",
                  lambda: os.startfile(os.path.dirname(saved_path)), width=16).pack(side=LEFT, padx=(0, 10))
        self._btn(bf, "📦 Outlook Backup starten",
                  lambda: self._set_mode("old_pc"), color=ACCENT_GREEN, width=22).pack(side=LEFT)

    def _set_mode(self, mode):
        self.mode = mode
        if mode == "old_pc":
            self.lbl_mode.config(text="🖥️  Modus: Alter PC")
            self.notebook.select(1)
            self._init_old_pc()
        else:
            self.lbl_mode.config(text="💻  Modus: Neuer PC")
            self.notebook.select(3)
            self._goto_run()

    def _init_old_pc(self):
        # Outlook erkennen
        ver, ver_name = detect_outlook_version()
        self.outlook_version = ver
        self.outlook_version_name = ver_name or "Unbekannt"

        if ver:
            self.lbl_outlook.config(text=f"✅  {ver_name} (v{ver})", fg=ACCENT_GREEN)
        else:
            self.lbl_outlook.config(
                text="⚠️  Outlook nicht erkannt (noch kein Profil?)", fg=ACCENT_WARN)

        # Benutzer
        users = get_windows_users()
        self.user_combo["values"] = users
        current = os.environ.get("USERPROFILE", "")
        self.user_profile.set(current if current in users else (users[0] if users else ""))

        # System Info
        info = get_system_info()
        self.lbl_sysinfo.config(
            text=f"RAM: {info['ram_gb']} GB  |  CPU: {info.get('cpu_cores', '?')} Kerne  |  "
                 f"Windows: {platform.version()}",
            fg=TEXT_WHITE
        )

        # USB Sticks
        usbs = find_usb_sticks()
        if usbs:
            self.usb_root.set(usbs[0])
            free = get_free_space(usbs[0])
            self.lbl_usb_space.config(text=f"Freier Speicher: {format_size(free)}", fg=ACCENT_GREEN)

    # ── NEUER PC FLOW ─────────────────────────────────────────

    def _start_new_pc(self, usb, config):
        self.mode = "new_pc"
        self.usb_root.set(usb)
        self.lbl_mode.config(text="💻  Modus: Neuer PC")

        # System prüfen
        info = get_system_info()
        warnings = []

        if info["ram_gb"] > 0 and info["ram_gb"] < MIN_RAM_GB:
            warnings.append(f"RAM: {info['ram_gb']} GB (Minimum: {MIN_RAM_GB} GB)")
        if info["cpu_mhz"] > 0 and info["cpu_mhz"] < MIN_CPU_MHZ:
            warnings.append(f"CPU: {info['cpu_mhz']:.0f} MHz (Minimum: {MIN_CPU_MHZ} MHz)")

        if warnings:
            msg = "Mindestanforderungen nicht erfüllt:\n\n" + "\n".join(warnings)
            msg += "\n\nTrotzdem fortfahren?"
            if not messagebox.askyesno("Systemanforderungen", msg):
                return

        # Outlook prüfen
        ver, ver_name = detect_outlook_version()
        if not ver:
            messagebox.showerror("Outlook fehlt",
                "Outlook wurde nicht gefunden!\n\n"
                "Bitte stelle sicher dass Outlook installiert ist.\n"
                "Hinweis: Outlook muss nicht zwingend eingerichtet sein –\n"
                "das Backup-Tool richtet es mit ein.")
            return
        self.outlook_version = ver
        self.outlook_version_name = ver_name or "Unbekannt"

        # Speicherplatz prüfen
        backup_dir = os.path.join(os.environ.get("USERPROFILE", ""), "Documents", "OutlookRestore")
        free = get_free_space(os.environ.get("USERPROFILE", "C:\\"))
        needed = get_folder_size(usb)
        if needed > free:
            answer = messagebox.askyesno(
                "Speicherplatz",
                f"Möglicherweise nicht genug Speicherplatz!\n"
                f"Benötigt: ~{format_size(needed)}\n"
                f"Verfügbar: {format_size(free)}\n\nTrotzdem fortfahren?"
            )
            if not answer:
                return

        self._update_new_pc_summary(usb, config)

        # PST-Dateien vom Stick laden und PST-Tab zeigen (zur Auswahl)
        backup_dir_val = config.get("backup_dir", os.path.join(usb, "Outlook_Backup"))
        self._scan_pst_from_usb(backup_dir_val)
        self.notebook.select(2)  # PST-Tab

    def _scan_pst_from_usb(self, backup_dir):
        """PST-Dateien aus dem USB-Backup laden (für Neuer-PC-Modus)."""
        self.all_pst_files = []
        self.pst_vars = {}
        self.pst_group_vars = {}

        # Zuerst PST-Unterordner durchsuchen
        pst_root = os.path.join(backup_dir, "PST")
        search_in = pst_root if os.path.isdir(pst_root) else backup_dir

        try:
            for root, _, files in os.walk(search_in):
                for f in files:
                    if f.lower().endswith(".pst"):
                        self.all_pst_files.append(os.path.join(root, f))
        except Exception:
            pass

        self._render_pst()

    def _update_new_pc_summary(self, usb, config):
        old_pc = config.get("computer_name", "?")
        old_ip = config.get("ip_address", "?")
        old_mac = config.get("mac_address", "?")

        # Standardwert = Outlook Standard-Pfad (wird von Outlook automatisch erkannt)
        if not self.pst_dest.get():
            default_pst = os.path.join(
                os.environ.get("USERPROFILE", "C:\\Users"),
                "AppData", "Local", "Microsoft", "Outlook"
            )
            self.pst_dest.set(default_pst)

        # Verbindung prüfen
        net_ok = is_host_reachable(old_pc) or is_host_reachable(old_ip)
        if not net_ok and old_mac:
            found_ip = find_host_by_mac(old_mac)
            if found_ip:
                net_ok = is_host_reachable(found_ip)

        method = "🌐 Netzwerk (alter PC erreichbar)" if net_ok else "💾 USB-Stick (kein Netzwerk)"
        self.lbl_summary.config(
            text=f"Modus:          Neuer PC\n"
                 f"Alter PC:       {old_pc} ({old_ip})\n"
                 f"Kopier-Methode: {method}\n"
                 f"USB-Stick:      {usb}\n"
                 f"PST-Zielordner: {self.pst_dest.get()}",
            fg=TEXT_WHITE
        )

        # PST-Zielordner Auswahl-Button anzeigen
        if not hasattr(self, "_pst_dest_frame") or not self.lbl_summary.winfo_exists():
            return
        try:
            self._pst_dest_frame.destroy()
        except Exception:
            pass
        self._pst_dest_frame = Frame(self.lbl_summary.master, bg=BG_CARD)
        self._pst_dest_frame.pack(anchor=W, pady=(6, 0))
        Label(self._pst_dest_frame, text="PST-Zielordner:",
              font=("Segoe UI", 9), bg=BG_CARD, fg=TEXT_GRAY).pack(side=LEFT, padx=(0, 6))
        Entry(self._pst_dest_frame, textvariable=self.pst_dest, font=("Segoe UI", 9),
              bg=BG_PANEL, fg=TEXT_WHITE, insertbackground=TEXT_WHITE,
              relief=FLAT, bd=4, width=38).pack(side=LEFT, padx=(0, 6))
        self._btn(self._pst_dest_frame, "📂", self._choose_pst_dest, width=3).pack(side=LEFT)

    # ── AKTIONEN ──────────────────────────────────────────────

    def _choose_usb(self):
        path = filedialog.askdirectory(title="USB-Stick / Zielordner auswählen")
        if path:
            self.usb_root.set(path)
            free = get_free_space(path)
            self.lbl_usb_space.config(text=f"Freier Speicher: {format_size(free)}", fg=ACCENT_GREEN)

    def _choose_pst_dest(self):
        path = filedialog.askdirectory(title="Zielordner für PST-Dateien wählen")
        if path:
            self.pst_dest.set(path)

    def _pst_subfolder(self, pst_path):
        """Unterordnername basierend auf Herkunftspfad bestimmen."""
        folder = os.path.dirname(pst_path).lower()
        if "appdata\\local\\microsoft\\outlook" in folder:
            return "Aktiv_Outlook"
        elif "appdata\\roaming\\microsoft\\outlook" in folder:
            return "Aktiv_Outlook_Roaming"
        elif "documents" in folder or "dokumente" in folder:
            return "Dokumente"
        elif "desktop" in folder or "schreibtisch" in folder:
            return "Desktop"
        else:
            # Laufwerksbuchstabe + erster Unterordner
            parts = pst_path.replace("/", "\\").split("\\")
            drive = parts[0].replace(":", "") if parts else "X"
            top = parts[1] if len(parts) > 1 else "Sonstige"
            return f"Sonstige_{drive}_{top}"

    def _goto_pst(self):
        if not self.usb_root.get():
            messagebox.showwarning("USB-Stick", "Bitte zuerst USB-Stick / Zielordner auswählen!")
            return
        self.notebook.select(2)
        if not self.all_pst_files:
            self._scan_pst()

    def _goto_run(self):
        self.notebook.select(3)
        if self.mode == "old_pc":
            self._update_old_pc_summary()

    def _update_old_pc_summary(self):
        sel = [p for p, v in self.pst_vars.items() if v.get()]
        total = sum(os.path.getsize(p) for p in sel if os.path.exists(p))
        free = get_free_space(self.usb_root.get()) if self.usb_root.get() else 0

        self.lbl_summary.config(
            text=f"Modus:            Alter PC\n"
                 f"Outlook:          {self.outlook_version_name}\n"
                 f"PST-Dateien:      {len(sel)} ausgewählt ({format_size(total)})\n"
                 f"PST Methode:      Netzwerkfreigabe (nicht auf Stick kopiert)\n"
                 f"Signaturen:       {'✅' if self.opt_signatures.get() else '❌'}\n"
                 f"Einstellungen:    {'✅' if self.opt_settings.get() else '❌'}\n"
                 f"Regeln:           {'✅' if self.opt_rules.get() else '❌'}\n"
                 f"Kontodaten:       {'✅' if self.opt_accounts.get() else '❌'}\n"
                 f"USB freier Speicher: {format_size(free)}",
            fg=TEXT_WHITE
        )

    def _scan_pst(self):
        self.lbl_pst_info.config(text="🔍 Suche läuft...")
        for w in self.pst_inner.winfo_children():
            w.destroy()

        def scan():
            user = self.user_profile.get()
            self.active_psts = get_active_pst_from_registry(user)
            self.all_pst_files = find_all_pst_files(user, log_cb=self._log)
            for a in self.active_psts:
                if a not in self.all_pst_files and os.path.exists(a):
                    self.all_pst_files.insert(0, a)
            self.root.after(0, self._render_pst)

        threading.Thread(target=scan, daemon=True).start()

    def _add_pst_manual(self):
        paths = filedialog.askopenfilenames(
            title="PST-Datei(en) hinzufügen",
            filetypes=[("PST Dateien", "*.pst"), ("Alle Dateien", "*.*")]
        )
        for p in paths:
            if p not in self.all_pst_files:
                self.all_pst_files.append(p)
        self._render_pst()

    def _browse_folder_for_pst(self):
        folder = filedialog.askdirectory(title="Ordner mit PST-Dateien auswählen")
        if not folder:
            return
        self.lbl_pst_info.config(text=f"🔍 Durchsuche {folder} ...")
        for w in self.pst_inner.winfo_children():
            w.destroy()

        def scan():
            found = []
            try:
                for root, dirs, files in os.walk(folder):
                    for f in files:
                        if f.lower().endswith(".pst"):
                            full = os.path.join(root, f)
                            if full not in self.all_pst_files:
                                found.append(full)
            except Exception:
                pass
            self.all_pst_files.extend(found)
            self.root.after(0, self._render_pst)

        threading.Thread(target=scan, daemon=True).start()

    def _fix_onedrive(self, path):
        if messagebox.askyesno("OneDrive PST",
            f"PST liegt in OneDrive:\n{path}\n\n"
            "OneDrive pausieren und PST in Dokumente verschieben?"):
            try:
                subprocess.Popen(["OneDrive.exe", "/pause"])
                time.sleep(2)
            except Exception:
                pass
            new_dir = os.path.join(self.user_profile.get(), "Documents", "Outlook Files")
            os.makedirs(new_dir, exist_ok=True)
            new_path = os.path.join(new_dir, os.path.basename(path))
            try:
                shutil.move(path, new_path)
                if path in self.all_pst_files:
                    self.all_pst_files.remove(path)
                if path in self.active_psts:
                    self.active_psts.remove(path)
                    self.active_psts.append(new_path)
                if new_path not in self.all_pst_files:
                    self.all_pst_files.insert(0, new_path)
                if path in self.pst_vars:
                    del self.pst_vars[path]
                self.pst_vars[new_path] = BooleanVar(value=True)
                self._render_pst()
                messagebox.showinfo("Erledigt", f"PST verschoben nach:\n{new_path}")
            except Exception as e:
                messagebox.showerror("Fehler", str(e))

    def _connect_net(self, path):
        messagebox.showinfo("Netzlaufwerk",
            f"Bitte Netzlaufwerk manuell verbinden:\n{path}\n\nDann 'Neu suchen' klicken.")

    def _confirm_delete(self, usb):
        if messagebox.askyesno("Stick zurücksetzen",
            "Alle Backup-Daten und Konfig vom Stick löschen?\n\n"
            "Das Script bleibt auf dem Stick.\n"
            "Der Stick ist danach bereit für einen neuen alten PC."):
            deleted = delete_config_and_data(usb)
            messagebox.showinfo("Erledigt",
                f"Stick zurückgesetzt!\n{len(deleted)} Elemente gelöscht.")

    # ── BACKUP STARTEN ────────────────────────────────────────

    def _start(self):
        if not self.usb_root.get():
            messagebox.showwarning("USB-Stick", "Kein Zielordner ausgewählt!")
            return
        if self.mode == "old_pc":
            self._run_old_pc_backup()
        elif self.mode == "new_pc":
            self._run_new_pc_import()

    def _cancel(self):
        if messagebox.askyesno("Abbrechen", "Vorgang wirklich abbrechen?"):
            self.cancel_event.set()
            self._log("⛔ Abbruch angefordert...")

    def _run_old_pc_backup(self):
        if not self.outlook_version:
            messagebox.showerror("Fehler", "Outlook nicht gefunden!")
            return

        sel_psts = [p for p, v in self.pst_vars.items() if v.get()]

        # Speicherplatz prüfen (ohne PSTs)
        free = get_free_space(self.usb_root.get())
        # Signaturen + Einstellungen schätzen
        user = self.user_profile.get()
        sig_size = get_folder_size(os.path.join(user, "AppData", "Roaming", "Microsoft", "Signatures"))
        settings_size = get_folder_size(os.path.join(user, "AppData", "Roaming", "Microsoft", "Outlook"))
        needed = sig_size + settings_size + (10 * 1024 * 1024)  # +10MB Puffer

        if needed > free:
            answer = messagebox.askyesno("Speicherplatz",
                f"Nicht genug Speicherplatz!\nBenötigt: {format_size(needed)}\n"
                f"Verfügbar: {format_size(free)}\n\nAnderen Zielordner wählen?")
            if answer:
                self._choose_usb()
            return

        self.cancel_event.clear()
        self.btn_start.config(state=DISABLED)
        self.btn_cancel.config(state=NORMAL)
        self.start_time = time.time()
        self.results = {"success": [], "warning": [], "error": []}

        threading.Thread(target=self._old_pc_thread, args=(sel_psts,), daemon=True).start()

    def _old_pc_thread(self, sel_psts):
        user = self.user_profile.get()
        usb = self.usb_root.get()
        date_str = datetime.datetime.now().strftime("%Y-%m-%d")
        backup_dir = os.path.join(usb, f"Outlook_Backup_{date_str}")

        progress = load_progress()
        done = progress.get("done", [])
        save_progress({"backup_dir": backup_dir, "done": done})

        # ── Outlook schließen
        self._log("🔄 Schließe Outlook...")
        self._set_cur("Outlook schließen...")
        ok, msg = close_outlook()
        self._log(msg)
        if not ok:
            self.root.after(0, lambda: messagebox.showerror("Fehler", msg))
            self.root.after(0, self._backup_done)
            return

        # ── Signaturen
        if self.opt_signatures.get() and not self.cancel_event.is_set():
            self._log("\n✍️  Kopiere Signaturen...")
            self._set_cur("Kopiere Signaturen...")
            sig_src = os.path.join(user, "AppData", "Roaming", "Microsoft", "Signatures")
            sig_dst = os.path.join(backup_dir, "Signaturen")
            if os.path.exists(sig_src):
                ok, msg = copy_folder(sig_src, sig_dst, log_cb=self._log,
                                       cancel_event=self.cancel_event)
                if ok:
                    self.results["success"].append("Signaturen")
                else:
                    self.results["error"].append(f"Signaturen: {msg}")
            else:
                self.results["warning"].append("Signaturen: Nicht gefunden")

        # ── Einstellungen Roaming
        if self.opt_settings.get() and not self.cancel_event.is_set():
            self._log("\n⚙️  Kopiere Einstellungen (Roaming)...")
            self._set_cur("Kopiere Einstellungen...")
            r_src = os.path.join(user, "AppData", "Roaming", "Microsoft", "Outlook")
            r_dst = os.path.join(backup_dir, "Einstellungen", "Roaming_Outlook")
            if os.path.exists(r_src):
                ok, msg = copy_folder(r_src, r_dst, log_cb=self._log, cancel_event=self.cancel_event)
                if ok:
                    self.results["success"].append("Einstellungen (Roaming)")
                else:
                    self.results["error"].append(f"Einstellungen Roaming: {msg}")
            else:
                self.results["warning"].append("Einstellungen Roaming: Nicht gefunden")

        # ── Einstellungen Local
        if self.opt_settings.get() and not self.cancel_event.is_set():
            self._log("\n⚙️  Kopiere Einstellungen (Local)...")
            l_src = os.path.join(user, "AppData", "Local", "Microsoft", "Outlook")
            l_dst = os.path.join(backup_dir, "Einstellungen", "Local_Outlook")
            if os.path.exists(l_src):
                ok, msg = copy_folder(l_src, l_dst, log_cb=self._log, cancel_event=self.cancel_event)
                if ok:
                    self.results["success"].append("Einstellungen (Local)")
                else:
                    self.results["error"].append(f"Einstellungen Local: {msg}")

        # ── Regeln
        if self.opt_rules.get() and not self.cancel_event.is_set():
            self._log("\n📋 Kopiere Regeln...")
            self._set_cur("Kopiere Regeln...")
            rules_src = os.path.join(user, "AppData", "Roaming", "Microsoft", "Outlook")
            rules_dst = os.path.join(backup_dir, "Regeln")
            os.makedirs(rules_dst, exist_ok=True)
            found_rules = False
            if os.path.exists(rules_src):
                for f in os.listdir(rules_src):
                    if f.lower().endswith(".rwz"):
                        shutil.copy2(os.path.join(rules_src, f), os.path.join(rules_dst, f))
                        found_rules = True
            if found_rules:
                self.results["success"].append("Regeln (.rwz)")
            else:
                self.results["warning"].append("Regeln: Keine .rwz Dateien gefunden")

        # ── Kontodaten
        if self.opt_accounts.get() and not self.cancel_event.is_set():
            self._log("\n📧 Lese Kontodaten...")
            self._set_cur("Lese Kontodaten...")
            accounts = get_outlook_account_info(self.outlook_version)
            acc_file = os.path.join(backup_dir, "Konten_Info.txt")
            os.makedirs(backup_dir, exist_ok=True)
            if accounts:
                with open(acc_file, "w", encoding="utf-8") as f:
                    f.write("Outlook Kontodaten\n" + "="*50 + "\n\n")
                    for acc in accounts:
                        for k, v in acc.items():
                            f.write(f"{k}: {v}\n")
                        f.write("\n")
                self.results["success"].append(f"Kontodaten ({len(accounts)} Konto/Konten)")
            else:
                self.results["warning"].append(
                    "Kontodaten: Nicht auslesbar (bei Outlook 365 normal – Account auf neuem PC einfach neu anmelden)"
                )

        # ── PST-Dateien verarbeiten
        pst_dest = self.pst_dest.get().strip()
        pst_manual = self.opt_pst_manual.get()

        if sel_psts and pst_dest and not self.cancel_event.is_set():
            pst_root = os.path.join(pst_dest, "PST")
            os.makedirs(pst_root, exist_ok=True)

            if pst_manual:
                # Manuell-Modus: nur Ordnerstruktur + Pfade-Liste anlegen
                self._log(f"\n📋 Manuell-Modus: Lege Ordnerstruktur an...")
                self._set_cur("Lege PST-Ordnerstruktur an...")
                liste_path = os.path.join(pst_root, "PST_Pfade.txt")
                with open(liste_path, "w", encoding="utf-8") as f:
                    f.write("PST-Dateien Liste\n")
                    f.write("=" * 60 + "\n\n")
                    for pst_path in sel_psts:
                        subfolder = self._pst_subfolder(pst_path)
                        ziel = os.path.join(pst_root, subfolder)
                        os.makedirs(ziel, exist_ok=True)
                        f.write(f"Quelle:  {pst_path}\n")
                        f.write(f"Ziel:    {os.path.join(ziel, os.path.basename(pst_path))}\n\n")
                self._log(f"  ✅ Ordnerstruktur erstellt. Pfade in: PST_Pfade.txt")
                self.results["success"].append(f"PST Ordnerstruktur angelegt ({len(sel_psts)} Einträge)")
            else:
                # Auto-Modus: PST-Dateien nach Herkunftsordner sortiert kopieren
                self._log(f"\n📁 Kopiere PST-Dateien nach: {pst_root}")
                self._set_cur("Kopiere PST-Dateien...")
                for pst_path in sel_psts:
                    if self.cancel_event.is_set():
                        break
                    pst_name = os.path.basename(pst_path)
                    subfolder = self._pst_subfolder(pst_path)
                    ziel_dir = os.path.join(pst_root, subfolder)
                    os.makedirs(ziel_dir, exist_ok=True)
                    dst = os.path.join(ziel_dir, pst_name)
                    self._log(f"  → [{subfolder}] {pst_name}")
                    try:
                        ok = copy_with_progress(pst_path, dst,
                                                progress_cb=self._file_progress,
                                                cancel_event=self.cancel_event)
                        if ok:
                            self.results["success"].append(f"PST kopiert: {subfolder}\\{pst_name}")
                        else:
                            self.results["error"].append(f"PST abgebrochen: {pst_name}")
                    except Exception as e:
                        self.results["error"].append(f"PST Fehler: {pst_name}: {e}")

        # ── Netzwerkfreigabe erstellen (nur wenn kein Zielordner gewählt)
        if sel_psts and not pst_dest and not self.cancel_event.is_set():
            self._log("\n🌐 Erstelle Netzwerkfreigabe für PST-Dateien...")
            self._set_cur("Erstelle Netzwerkfreigabe...")
            share_folder = os.path.dirname(sel_psts[0])
            ok, share_unc = create_network_share(share_folder)
            if ok:
                self._log(f"  ✅ Freigabe erstellt: {share_unc}")
                self.results["success"].append(f"Netzwerkfreigabe: {share_unc}")
                self.share_path = share_unc
            else:
                self._log(f"  ❌ Freigabe fehlgeschlagen: {share_unc}")
                self.results["error"].append(f"Netzwerkfreigabe fehlgeschlagen: {share_unc}")

        # ── Konfig-Datei speichern
        if not self.cancel_event.is_set():
            self._log("\n💾 Speichere Konfig auf Stick...")
            config = {
                "computer_name": get_computer_name(),
                "ip_address": get_ip_address(),
                "mac_address": get_mac_address(),
                "share_name": SHARE_NAME,
                "share_path": self.share_path or "",
                "pst_files": sel_psts,
                "backup_dir": backup_dir,
                "backup_date": datetime.datetime.now().strftime("%d.%m.%Y %H:%M"),
                "outlook_version": self.outlook_version,
            }
            try:
                save_config(usb, config)
                self.results["success"].append("Konfig-Datei gespeichert")
                self._log("  ✅ Konfig gespeichert.")
            except Exception as e:
                self.results["error"].append(f"Konfig: {e}")

        clear_progress()
        self.root.after(0, self._backup_done)

    # ── NEUER PC IMPORT ───────────────────────────────────────

    def _run_new_pc_import(self):
        if not self.usb_config:
            messagebox.showerror("Fehler", "Keine Konfig gefunden!")
            return

        self.cancel_event.clear()
        self.btn_start.config(state=DISABLED)
        self.btn_cancel.config(state=NORMAL)
        self.start_time = time.time()
        self.results = {"success": [], "warning": [], "error": []}

        threading.Thread(target=self._new_pc_thread, daemon=True).start()

    def _new_pc_thread(self):
        config = self.usb_config
        usb = self.usb_root.get()
        user = os.environ.get("USERPROFILE", "")
        backup_dir = config.get("backup_dir", os.path.join(usb, "Outlook_Backup"))

        # Verbindung zum alten PC prüfen
        old_pc = config.get("computer_name", "")
        old_ip = config.get("ip_address", "")
        old_mac = config.get("mac_address", "")
        share_path = config.get("share_path", "")
        self._log("🌐 Prüfe Verbindung zum alten PC...")
        net_ok = False

        for host in [old_pc, old_ip]:
            if host and is_host_reachable(host):
                net_ok = True
                self._log(f"  ✅ Alter PC erreichbar: {host}")
                break

        if not net_ok and old_mac:
            self._log("  🔍 Suche alter PC via MAC-Adresse...")
            found = find_host_by_mac(old_mac)
            if found and is_host_reachable(found):
                net_ok = True
                self._log(f"  ✅ Alter PC gefunden via MAC: {found}")

        if not net_ok:
            self._log("  ⚠️  Alter PC nicht erreichbar → Kopiere vom USB-Stick")
            self.results["warning"].append("Netzwerk nicht verfügbar – Daten vom Stick")

        # PST Dateien kopieren – direkt die selektierten Pfade vom PST-Tab (echte USB-Pfade)
        sel_psts = [p for p, v in self.pst_vars.items() if v.get()]
        if sel_psts and not self.cancel_event.is_set():
            self._log("\n📁 Kopiere PST-Dateien...")
            pst_dst_dir = self.pst_dest.get().strip() or os.path.join(
                user, "AppData", "Local", "Microsoft", "Outlook")
            os.makedirs(pst_dst_dir, exist_ok=True)
            self._log(f"  Zielordner: {pst_dst_dir}")

            progress = load_progress()
            done = progress.get("done", [])

            for src in sel_psts:
                if self.cancel_event.is_set():
                    break

                pst_name = os.path.basename(src)

                if pst_name in done:
                    self._log(f"  ⏩ Übersprungen (bereits kopiert): {pst_name}")
                    continue

                if not os.path.exists(src):
                    self._log(f"  ❌ Nicht gefunden: {src}")
                    self.results["error"].append(f"PST nicht gefunden: {pst_name}")
                    continue

                dst = os.path.join(pst_dst_dir, pst_name)
                self._log(f"  → {pst_name}")
                self._set_cur(f"Kopiere: {pst_name}")

                # Verbindungsüberwachung während Kopieren
                retry = True
                while retry and not self.cancel_event.is_set():
                    try:
                        ok = copy_with_progress(src, dst,
                                                 progress_cb=self._file_progress,
                                                 cancel_event=self.cancel_event)
                        if ok:
                            # Vollständigkeit prüfen
                            self._log(f"  🔍 Prüfe Vollständigkeit: {pst_name}")
                            ver_ok, ver_msg = verify_file(src, dst)
                            if ver_ok:
                                self.results["success"].append(f"PST: {pst_name}")
                                done.append(pst_name)
                                save_progress({"done": done})
                                retry = False
                            else:
                                self._log(f"  ⚠️  Prüfung fehlgeschlagen: {ver_msg}")
                                action = self._ask_retry(pst_name, ver_msg)
                                if action == "retry":
                                    self._log(f"  🔄 Wiederhole: {pst_name}")
                                elif action == "skip":
                                    self.results["warning"].append(f"PST übersprungen: {pst_name}")
                                    retry = False
                                else:
                                    self.cancel_event.set()
                                    retry = False
                        else:
                            retry = False
                    except Exception as e:
                        self._log(f"  ❌ Fehler: {e}")
                        self.results["error"].append(f"PST Fehler: {pst_name}")
                        retry = False

        # Signaturen wiederherstellen
        sig_backup = os.path.join(backup_dir, "Signaturen")
        if os.path.exists(sig_backup) and not self.cancel_event.is_set():
            self._log("\n✍️  Stelle Signaturen wieder her...")
            self._set_cur("Signaturen wiederherstellen...")
            ok, msg = restore_signatures(sig_backup, user)
            if ok:
                self.results["success"].append("Signaturen wiederhergestellt")
            else:
                self.results["error"].append(f"Signaturen: {msg}")

        # Einstellungen wiederherstellen
        if not self.cancel_event.is_set():
            self._log("\n⚙️  Stelle Einstellungen wieder her...")
            self._set_cur("Einstellungen wiederherstellen...")
            roaming_bk = os.path.join(backup_dir, "Einstellungen", "Roaming_Outlook")
            local_bk   = os.path.join(backup_dir, "Einstellungen", "Local_Outlook")
            res = restore_settings(roaming_bk, local_bk, user)
            for name, ok, msg in res:
                if ok:
                    self.results["success"].append(f"{name} wiederhergestellt")
                else:
                    self.results["error"].append(f"{name}: {msg}")

        # Regeln wiederherstellen
        rules_bk = os.path.join(backup_dir, "Regeln")
        if os.path.exists(rules_bk) and not self.cancel_event.is_set():
            self._log("\n📋 Stelle Regeln wieder her...")
            self._set_cur("Regeln wiederherstellen...")
            ok, msg = restore_rules(rules_bk, user)
            if ok:
                self.results["success"].append("Regeln wiederhergestellt")
            else:
                self.results["error"].append(f"Regeln: {msg}")

        # PST in Outlook importieren – nur bei klassischem Outlook, nicht bei App-Version
        is_new_outlook = "App-Version" in (self.outlook_version_name or "")
        if not is_new_outlook and not self.cancel_event.is_set():
            self._log("\n📧 Importiere PST in Outlook...")
            self._set_cur("PST in Outlook importieren...")
            pst_dst_dir = self.pst_dest.get().strip() or os.path.join(
                user, "AppData", "Local", "Microsoft", "Outlook")
            if os.path.exists(pst_dst_dir):
                for pst_file in os.listdir(pst_dst_dir):
                    if pst_file.lower().endswith(".pst"):
                        pst_path = os.path.join(pst_dst_dir, pst_file)
                        self._log(f"  → Importiere: {pst_file}")
                        ok, msg = import_pst_to_outlook(pst_path)
                        if ok:
                            self.results["success"].append(f"PST importiert: {pst_file}")
                        else:
                            self.results["warning"].append(f"PST Import: {pst_file} – {msg}")

        # Netzwerkfreigabe schließen (Optional)
        if net_ok and share_path and not self.cancel_event.is_set():
            self.root.after(0, self._ask_close_share)

        # Hinweis für neue Outlook-App: Konten müssen manuell eingerichtet werden
        if is_new_outlook and not self.cancel_event.is_set():
            konten_file = os.path.join(backup_dir, "Konten_Info.txt")
            konten_info = ""
            if os.path.exists(konten_file):
                try:
                    with open(konten_file, "r", encoding="utf-8") as f:
                        konten_info = f.read()
                except Exception:
                    pass
            self.root.after(0, lambda ki=konten_info: self._show_new_outlook_hint(ki))

        clear_progress()
        self.root.after(0, self._backup_done)

    def _show_new_outlook_hint(self, konten_info):
        """Hinweis nach Restore: Neue Outlook-App benötigt manuelle Kontoeinrichtung."""
        win = Toplevel(self.root)
        win.title("Konten einrichten – Neue Outlook-App")
        win.configure(bg=BG_DARK)
        win.resizable(False, False)
        win.grab_set()

        Label(win, text="⚠️  Konten manuell einrichten",
              font=("Segoe UI", 13, "bold"), bg=BG_DARK, fg=ACCENT_WARN).pack(pady=(18, 6), padx=24)

        msg = (
            "Du verwendest die neue Outlook-App (Windows 11).\n"
            "Bitte in dieser Reihenfolge vorgehen:\n\n"
            "  1.  Outlook öffnen → Konto hinzufügen\n"
            "       → E-Mail-Adresse & Passwort eingeben\n"
            "       (alle Konten vom alten PC einrichten)\n\n"
            "  2.  PST-Dateien importieren (falls vorhanden)\n"
            "       Outlook → Datei → Öffnen → PST-Datei importieren\n\n"
            "  3.  Signaturen & Einstellungen wurden bereits\n"
            "       wiederhergestellt und sind aktiv sobald\n"
            "       das Konto eingerichtet ist."
        )
        Label(win, text=msg, font=("Segoe UI", 10), bg=BG_DARK, fg=TEXT_WHITE,
              justify=LEFT).pack(padx=24, pady=(0, 10))

        if konten_info:
            Label(win, text="Gesicherte Konten vom alten PC:",
                  font=("Segoe UI", 10, "bold"), bg=BG_DARK, fg=ACCENT_CYAN).pack(anchor=W, padx=24)
            txt = ScrolledText(win, height=10, width=64, bg=BG_PANEL, fg=TEXT_WHITE,
                               font=("Consolas", 9), relief=FLAT)
            txt.pack(padx=24, pady=6)
            txt.insert("end", konten_info)
            txt.config(state=DISABLED)

        self._btn(win, "OK – Verstanden", win.destroy, color=ACCENT_GREEN, width=20).pack(pady=(6, 18))
        win.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - win.winfo_width()) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - win.winfo_height()) // 2
        win.geometry(f"+{x}+{y}")

    def _ask_retry(self, filename, reason):
        """Abfrage: Wiederholen / Überspringen / Abbrechen."""
        result = [None]
        event = threading.Event()

        def ask():
            win = Toplevel(self.root)
            win.title("Datei fehlerhaft")
            win.configure(bg=BG_DARK)
            win.geometry("450x200")
            win.grab_set()

            Label(win, text=f"⚠️  Datei fehlerhaft: {filename}",
                  font=("Segoe UI", 11, "bold"), bg=BG_DARK, fg=ACCENT_WARN).pack(pady=(20, 5))
            Label(win, text=reason, font=("Segoe UI", 10),
                  bg=BG_DARK, fg=TEXT_GRAY).pack()

            bf = Frame(win, bg=BG_DARK)
            bf.pack(pady=20)

            def choose(val):
                result[0] = val
                win.destroy()
                event.set()

            self._btn(bf, "🔄 Nochmal", lambda: choose("retry"), width=12).pack(side=LEFT, padx=5)
            self._btn(bf, "⏭️ Überspringen", lambda: choose("skip"),
                      color=ACCENT_WARN, width=14).pack(side=LEFT, padx=5)
            self._btn(bf, "⛔ Abbrechen", lambda: choose("abort"),
                      color=ACCENT_RED, width=12).pack(side=LEFT, padx=5)

        self.root.after(0, ask)
        event.wait(timeout=300)
        return result[0] or "skip"

    def _ask_close_share(self):
        if messagebox.askyesno("Netzwerkfreigabe",
            "Netzwerkfreigabe auf dem alten PC schließen?"):
            remove_network_share()
            self._log("  ✅ Netzwerkfreigabe entfernt.")

    def _ask_manual_account(self, save_path):
        win = Toplevel(self.root)
        win.title("Kontodaten manuell eingeben")
        win.configure(bg=BG_DARK)
        win.geometry("500x380")

        Label(win, text="Kontodaten manuell eingeben",
              font=("Segoe UI", 13, "bold"), bg=BG_DARK, fg=TEXT_WHITE).pack(pady=15)

        fields = {}
        for label in ["E-Mail Adresse", "IMAP Server", "SMTP Server",
                       "Benutzername", "Port (IMAP)", "Port (SMTP)"]:
            row = Frame(win, bg=BG_DARK)
            row.pack(fill=X, padx=20, pady=4)
            Label(row, text=label, width=18, anchor=W, font=("Segoe UI", 10),
                  bg=BG_DARK, fg=TEXT_GRAY).pack(side=LEFT)
            var = StringVar()
            Entry(row, textvariable=var, bg=BG_PANEL, fg=TEXT_WHITE,
                  insertbackground=TEXT_WHITE, relief=FLAT, bd=5,
                  font=("Segoe UI", 10)).pack(side=LEFT, fill=X, expand=True)
            fields[label] = var

        def save():
            with open(save_path, "w", encoding="utf-8") as f:
                f.write("Outlook Kontodaten (manuell)\n" + "="*50 + "\n\n")
                for k, v in fields.items():
                    f.write(f"{k}: {v.get()}\n")
            self.results["success"].append("Kontodaten (manuell)")
            win.destroy()

        self._btn(win, "💾 Speichern", save, width=16).pack(pady=15)

    # ── ABSCHLUSS ─────────────────────────────────────────────

    def _backup_done(self):
        self.btn_start.config(state=NORMAL)
        self.btn_cancel.config(state=DISABLED)
        self._set_cur("✅ Abgeschlossen!")
        self.progress_var.set(100)
        self.lbl_pct.config(text="100%", fg=ACCENT_GREEN)
        self._update_result()
        self.notebook.select(4)

        messagebox.showinfo("Fertig!",
            "✅ Vorgang abgeschlossen!\n\n" +
            ("Der USB-Stick kann jetzt sicher entfernt werden." if self.mode == "old_pc"
             else "Alle Daten wurden erfolgreich importiert."))

    def _reset(self):
        self.results = {"success": [], "warning": [], "error": []}
        self.progress_var.set(0)
        self.lbl_pct.config(text="0%", fg=ACCENT_CYAN)
        self.lbl_current.config(text="")
        self.log_text.config(state=NORMAL)
        self.log_text.delete(1.0, END)
        self.log_text.config(state=DISABLED)
        self.notebook.select(0)
        self.root.after(300, self._auto_detect)

    # ── HILFSMETHODEN ─────────────────────────────────────────

    def _log(self, msg):
        def _do():
            self.log_text.config(state=NORMAL)
            self.log_text.insert(END, msg + "\n")
            self.log_text.see(END)
            self.log_text.config(state=DISABLED)
        self.root.after(0, _do)

    def _set_cur(self, msg):
        self.root.after(0, lambda: self.lbl_current.config(text=msg))

    def _file_progress(self, copied, total):
        if total > 0:
            pct = (copied / total) * 100
            self.root.after(0, lambda: self.progress_var.set(pct))
            self.root.after(0, lambda: self.lbl_pct.config(text=f"{pct:.1f}%"))
            elapsed = time.time() - self.start_time
            if elapsed > 0 and copied > 0:
                speed = copied / elapsed
                self.root.after(0, lambda: self.lbl_speed.config(
                    text=f"⚡ {format_size(int(speed))}/s"
                ))

    def _manual_account(self, path):
        self._ask_manual_account(path)


# ═══════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════

# ═══════════════════════════════════════════════════════════════
# HARDWARE SCAN FUNKTIONEN
# ═══════════════════════════════════════════════════════════════

HARDWARE_INFO_FILE = "Hardware_Info.txt"

DRIVER_LINKS = {
    "asus":    "https://www.asus.com/support/",
    "gigabyte":"https://www.gigabyte.com/Support",
    "msi":     "https://www.msi.com/support",
    "nvidia":  "https://www.nvidia.com/drivers",
    "amd":     "https://www.amd.com/support",
    "intel":   "https://www.intel.com/content/www/us/en/download-center/home.html",
    "realtek": "https://www.realtek.com/en/downloads",
    "qualcomm":"https://www.qualcomm.com/support",
}

def get_driver_link(component_name):
    name_lower = component_name.lower()
    for brand, link in DRIVER_LINKS.items():
        if brand in name_lower:
            return link
    return "https://www.google.com/search?q=" + component_name.replace(" ", "+") + "+driver+download"

def scan_hardware():
    hardware = {}
    errors = []
    try:
        result = subprocess.run(["wmic", "baseboard", "get", "Manufacturer,Product", "/format:csv"], capture_output=True, text=True, timeout=15)
        for line in result.stdout.splitlines():
            if line.strip() and "Node" not in line and "," in line:
                parts = line.strip().split(",")
                if len(parts) >= 3:
                    manufacturer = parts[1].strip()
                    product = parts[2].strip()
                    if manufacturer and product:
                        hardware["Mainboard"] = f"{manufacturer} {product}"
                        break
    except Exception as e:
        errors.append(f"Mainboard: {e}")
    try:
        result = subprocess.run(["wmic", "cpu", "get", "Name", "/format:csv"], capture_output=True, text=True, timeout=15)
        for line in result.stdout.splitlines():
            if line.strip() and "Node" not in line and "," in line:
                parts = line.strip().split(",")
                if len(parts) >= 2 and parts[1].strip():
                    hardware["CPU"] = parts[1].strip()
                    break
    except Exception as e:
        errors.append(f"CPU: {e}")
    try:
        result = subprocess.run(["wmic", "path", "win32_VideoController", "get", "Name", "/format:csv"], capture_output=True, text=True, timeout=15)
        gpus = []
        for line in result.stdout.splitlines():
            if line.strip() and "Node" not in line and "," in line:
                parts = line.strip().split(",")
                if len(parts) >= 2 and parts[1].strip():
                    gpus.append(parts[1].strip())
        if gpus:
            hardware["GPU"] = " | ".join(gpus)
    except Exception as e:
        errors.append(f"GPU: {e}")
    try:
        result = subprocess.run(["wmic", "computersystem", "get", "TotalPhysicalMemory", "/format:csv"], capture_output=True, text=True, timeout=15)
        for line in result.stdout.splitlines():
            if line.strip() and "Node" not in line and "," in line:
                parts = line.strip().split(",")
                if len(parts) >= 2 and parts[1].strip():
                    ram_bytes = int(parts[1].strip())
                    hardware["RAM"] = f"{round(ram_bytes / (1024**3), 1)} GB"
                    break
    except Exception as e:
        errors.append(f"RAM: {e}")
    try:
        result = subprocess.run(["wmic", "nic", "where", "PhysicalAdapter=TRUE", "get", "Name", "/format:csv"], capture_output=True, text=True, timeout=15)
        nics = []
        for line in result.stdout.splitlines():
            if line.strip() and "Node" not in line and "," in line:
                parts = line.strip().split(",")
                if len(parts) >= 2 and parts[1].strip():
                    nics.append(parts[1].strip())
        if nics:
            hardware["LAN"] = " | ".join(nics)
    except Exception as e:
        errors.append(f"LAN: {e}")
    try:
        result = subprocess.run(["wmic", "sounddev", "get", "Name", "/format:csv"], capture_output=True, text=True, timeout=15)
        audio = []
        for line in result.stdout.splitlines():
            if line.strip() and "Node" not in line and "," in line:
                parts = line.strip().split(",")
                if len(parts) >= 2 and parts[1].strip():
                    audio.append(parts[1].strip())
        if audio:
            hardware["Audio"] = " | ".join(audio)
    except Exception as e:
        errors.append(f"Audio: {e}")
    try:
        hardware["Windows"] = platform.version()
        hardware["Windows Name"] = platform.win32_ver()[0]
    except Exception as e:
        errors.append(f"Windows: {e}")
    return hardware, errors

def save_hardware_info(hardware, errors, save_path):
    os.makedirs(os.path.dirname(save_path), exist_ok=True)
    with open(save_path, "w", encoding="utf-8") as f:
        f.write("=" * 60 + "\n")
        f.write("  Hardware Info\n")
        f.write(f"  Erstellt am: {datetime.datetime.now().strftime('%d.%m.%Y %H:%M')}\n")
        f.write("=" * 60 + "\n\n")
        for component, value in hardware.items():
            f.write(f"{component}: {value}\n")
            if component not in ("RAM", "Windows", "Windows Name"):
                link = get_driver_link(value)
                f.write(f"  Treiber: {link}\n")
            f.write("\n")
        if errors:
            f.write("\n" + "=" * 60 + "\n")
            f.write("WARNUNGEN:\n")
            for err in errors:
                f.write(f"  {err}\n")
    return save_path

def find_hardware_info_on_usb():
    for usb in find_usb_sticks():
        hw_path = os.path.join(usb, "Treiber", HARDWARE_INFO_FILE)
        if os.path.exists(hw_path):
            return hw_path, usb
    return None, None


def main():
    if not is_admin():
        try:
            root = Tk()
            root.withdraw()
            if messagebox.askyesno("Admin-Rechte",
                "Administrator-Rechte erforderlich.\nJetzt neu starten?"):
                root.destroy()
                run_as_admin()
            else:
                root.destroy()
                sys.exit(0)
        except Exception:
            run_as_admin()
        return

    root = Tk()
    app = MigrationApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

