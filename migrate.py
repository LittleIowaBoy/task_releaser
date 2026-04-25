"""First-run migrator for DocuReader installs.

Older releases installed to ``%ProgramFiles%\\DocuReader``, which is read-only
without UAC and so prevents the in-app auto-updater from overwriting files.
The new install path is ``%LOCALAPPDATA%\\Programs\\DocuReader``.

When a user runs a new ``DocuReader.exe`` that still happens to live under
``Program Files``, this migrator:

1. Copies the entire install tree into the per-user location.
2. Re-points the Start Menu and Desktop shortcuts.
3. Records a sentinel file at ``~/.docureader/.migrated`` so the migrator
   never runs twice.
4. Spawns the new exe and asks the caller to exit.

If anything fails (locked files, permissions, missing source) we bail out
quietly and let the app keep running from its current location - the auto-
updater branch will fall back to a clear error in that scenario.
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Optional


SENTINEL_DIR = Path.home() / ".docureader"
SENTINEL_PATH = SENTINEL_DIR / ".migrated"


def _new_install_dir() -> Path:
    base = os.environ.get("LOCALAPPDATA") or str(Path.home() / "AppData" / "Local")
    return Path(base) / "Programs" / "DocuReader"


def _is_under_program_files(exe_path: Path) -> bool:
    candidates = [
        os.environ.get("ProgramFiles"),
        os.environ.get("ProgramFiles(x86)"),
        os.environ.get("ProgramW6432"),
    ]
    s = str(exe_path).lower()
    return any(p and s.startswith(p.lower()) for p in candidates)


def _refresh_shortcut(lnk_path: Path, target_exe: Path) -> None:
    """Repoint a .lnk to ``target_exe`` via WScript.Shell."""
    if not lnk_path.exists():
        return
    try:
        # Use a tiny VBScript to avoid a PyWin32 dependency.
        vbs = (
            f'Set s = CreateObject("WScript.Shell")\r\n'
            f'Set l = s.CreateShortcut("{lnk_path}")\r\n'
            f'l.TargetPath = "{target_exe}"\r\n'
            f'l.WorkingDirectory = "{target_exe.parent}"\r\n'
            f'l.IconLocation = "{target_exe}"\r\n'
            f'l.Save\r\n'
        )
        tmp = lnk_path.parent / "_repoint.vbs"
        tmp.write_text(vbs, encoding="ascii")
        try:
            subprocess.run(["cscript", "//nologo", str(tmp)], check=False, timeout=15)
        finally:
            try:
                tmp.unlink()
            except OSError:
                pass
    except (OSError, subprocess.SubprocessError):
        pass


def maybe_migrate_install() -> bool:
    """If running from Program Files, copy to per-user dir and relaunch.

    Returns True when a relaunch was scheduled (the caller should exit).
    Returns False otherwise (no action needed, or migration failed).
    """
    if os.name != "nt":
        return False
    if not getattr(sys, "frozen", False):
        return False

    try:
        SENTINEL_DIR.mkdir(parents=True, exist_ok=True)
        if SENTINEL_PATH.exists():
            return False
    except OSError:
        return False

    exe = Path(sys.executable).resolve()
    if not _is_under_program_files(exe):
        # Already in a writable location - mark migrated so we never check again.
        try:
            SENTINEL_PATH.write_text("not-needed\n", encoding="utf-8")
        except OSError:
            pass
        return False

    src_dir = exe.parent
    dst_dir = _new_install_dir()
    if dst_dir.exists() and any(dst_dir.iterdir()):
        # Per-user install already there; just relaunch into it.
        new_exe = dst_dir / exe.name
        if new_exe.exists():
            try:
                SENTINEL_PATH.write_text(f"existing:{dst_dir}\n", encoding="utf-8")
                subprocess.Popen([str(new_exe)], close_fds=True)
                return True
            except OSError:
                return False
        return False

    try:
        dst_dir.parent.mkdir(parents=True, exist_ok=True)
        # copytree refuses an existing dst; only call when it is absent or empty.
        if dst_dir.exists():
            shutil.rmtree(dst_dir, ignore_errors=True)
        shutil.copytree(src_dir, dst_dir)
    except (OSError, shutil.Error) as e:
        try:
            (SENTINEL_DIR / ".migration_error").write_text(repr(e), encoding="utf-8")
        except OSError:
            pass
        return False

    new_exe = dst_dir / exe.name
    if not new_exe.exists():
        return False

    # Best-effort shortcut refresh.
    appdata = os.environ.get("APPDATA")
    if appdata:
        sm = Path(appdata) / "Microsoft" / "Windows" / "Start Menu" / "Programs" / "DocuReader" / "DocuReader.lnk"
        _refresh_shortcut(sm, new_exe)
    desktop = Path.home() / "Desktop" / "DocuReader.lnk"
    _refresh_shortcut(desktop, new_exe)

    try:
        SENTINEL_PATH.write_text(f"migrated:{dst_dir}\n", encoding="utf-8")
    except OSError:
        pass

    try:
        subprocess.Popen([str(new_exe)], close_fds=True)
    except OSError:
        return False
    return True
