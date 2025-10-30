#!/usr/bin/env python3
import os
import sys
import shutil
import subprocess
import time
import getpass
from datetime import datetime, timedelta

# ---------------- CONFIG ----------------
MAPPED_ROOT = r"Z:\\"  
UNC_ROOT = r"\\server\Unity Exploration\IT Department"
BACKUP_FOLDER_NAME = "BACKUP"
USER_FOLDERS = ["Desktop", "Documents"]
RETRY_INTERVAL_SECONDS = 60
DRIVE_RETRY_MINUTES = 15
MAX_FILE_SIZE_MB = 50000  # Skip any file >1 GB
OUTLOOK_FOLDERS = [
    r"Documents\Outlook", 
    r"Documents\OutlookFiles"
]  # Folder(s) for Outlook
PST_EXTENSION = ".pst"  # Outlook PST file extension
# ----------------------------------------


def is_drive_available(path):
    try:
        return os.path.exists(path)
    except Exception:
        return False


def find_destination(username):
    """Find valid destination path for backup."""
    if UNC_ROOT:
        unc_user = os.path.join(UNC_ROOT, username)
        if is_drive_available(unc_user):
            return os.path.join(unc_user, BACKUP_FOLDER_NAME)

    mapped_user = os.path.join(MAPPED_ROOT, username)
    if is_drive_available(mapped_user):
        return os.path.join(mapped_user, BACKUP_FOLDER_NAME)

    if is_drive_available(MAPPED_ROOT):
        return os.path.join(MAPPED_ROOT, BACKUP_FOLDER_NAME)

    return None


def robocopy_available():
    """Check if robocopy is available in PATH."""
    try:
        subprocess.run(["robocopy"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return True
    except Exception:
        return False


def should_skip(path):
    """Skip system/app folders, Desktop shortcuts, Temp files, and very large files, except PST files."""
    lower = path.lower()

    # Skip system, app, or excluded folders
    if any(skip in lower for skip in ["appdata", "program files", "windows", "microsoft", "temp"]):
        return True

    # Skip Desktop shortcuts (.lnk files)
    if path.endswith(".lnk"):
        return True

    # Skip large files except Outlook .pst files
    if os.path.isfile(path):
        try:
            size_mb = os.path.getsize(path) / (50000 * 50000)
            if size_mb > MAX_FILE_SIZE_MB and not path.endswith(PST_EXTENSION):
                return True
        except Exception:
            return True

    # Do not skip Outlook .pst files
    if any(outlook_folder.lower() in path.lower() for outlook_folder in OUTLOOK_FOLDERS):
        if path.endswith(PST_EXTENSION):
            return False

    return False


def copy_folder_incremental(src, dst):
    """Copy only new or modified files from src to dst."""
    if not os.path.exists(src):
        return

    folder_name = os.path.basename(src.rstrip("\\/"))
    target = os.path.join(dst, folder_name)
    os.makedirs(target, exist_ok=True)

    # --- Use Robocopy for fast and reliable incremental sync ---
    if robocopy_available():
        cmd = [
            "robocopy", src, target,
            "/E", "/XO", "/XN", "/XC",  # skip older/same/unchanged files
            "/NFL", "/NDL", "/NJH", "/NJS", "/NP",
            "/XF", "*.exe", "*.msi", "*.bat", "*.cmd", "*.dll", 
            "/XD", "AppData", "Program Files", "Windows", "Temp", "Desktop"
        ]
        try:
            subprocess.run(
                cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, creationflags=subprocess.CREATE_NO_WINDOW
            )
        except Exception:
            pass
        return

    # --- Fallback: Python incremental copy ---
    for root, dirs, files in os.walk(src):
        rel = os.path.relpath(root, src)
        target_root = os.path.join(target, rel) if rel != "." else target
        os.makedirs(target_root, exist_ok=True)

        for f in files:
            src_f = os.path.join(root, f)
            if should_skip(src_f):
                continue

            dst_f = os.path.join(target_root, f)

            try:
                # Copy only if new or modified
                if not os.path.exists(dst_f) or (
                    os.path.getmtime(src_f) > os.path.getmtime(dst_f)
                ):
                    shutil.copy2(src_f, dst_f)
            except Exception:
                pass


def main():
    username = getpass.getuser()
    userprofile = os.getenv("USERPROFILE") or os.path.expanduser("~")

    src_folders = [
        os.path.join(userprofile, name)
        for name in USER_FOLDERS
        if os.path.exists(os.path.join(userprofile, name))
    ]

    if not src_folders:
        return 0

    # Wait for destination drive
    deadline = datetime.now() + timedelta(minutes=DRIVE_RETRY_MINUTES)
    dest_root = None
    while datetime.now() < deadline:
        dest_root = find_destination(username)
        if dest_root:
            break
        time.sleep(RETRY_INTERVAL_SECONDS)

    if not dest_root:
        return 1

    os.makedirs(dest_root, exist_ok=True)

    for src in src_folders:
        copy_folder_incremental(src, dest_root)

    return 0


if __name__ == "__main__":
    try:
        # Prevent console window reopening
        sys.exit(main())
    except Exception:
        sys.exit(99)
