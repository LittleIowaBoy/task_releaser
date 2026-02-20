#!/usr/bin/env python3
"""
Quick script to rebuild and repackage DocuReader
"""
import subprocess
import sys
import zipfile
import os
import re
import time
import stat
import shutil
from importlib.util import find_spec
from pathlib import Path

BASE_DIR = Path(__file__).parent
BUILD_DIR = BASE_DIR / "freeze_build" / "cx_freeze"
REQUIRED_BUILD_PACKAGES = ["pandas", "openpyxl", "xlrd", "PyQt6"]
FREEZE_TIMEOUT_SECONDS = 20 * 60


def read_version() -> str:
    try:
        content = (BASE_DIR / "tr_gui.py").read_text(encoding="utf-8")
        match = re.search(r'__version__\s*=\s*["\']([^"\']+)["\']', content)
        return match.group(1) if match else "0.0.0"
    except Exception:
        return "0.0.0"


VERSION = read_version()
ZIP_NAME = f"DocuReader-{VERSION}-portable.zip"


def find_missing_packages(package_names: list[str]) -> list[str]:
    missing = []
    for package_name in package_names:
        if find_spec(package_name) is None:
            missing.append(package_name)
    return missing


def _handle_remove_readonly(func, path, exc_info):
    os.chmod(path, stat.S_IWRITE)
    func(path)


def clean_build_dir(build_dir: Path, retries: int = 3, delay_seconds: float = 1.0) -> bool:
    if not build_dir.exists():
        return True

    for attempt in range(1, retries + 1):
        try:
            shutil.rmtree(build_dir, onerror=_handle_remove_readonly)
            return True
        except OSError as exc:
            if attempt == retries:
                print(f"  ERROR: Could not clean build directory: {build_dir}")
                print(f"  Details: {exc}")
                return False
            time.sleep(delay_seconds)
    return False


def read_log_tail(log_path: Path, max_lines: int = 40) -> str:
    try:
        lines = log_path.read_text(encoding="utf-8", errors="replace").splitlines()
    except OSError:
        return ""
    return "\n".join(lines[-max_lines:])

print("=" * 60)
print("Rebuilding and Repackaging DocuReader")
print("=" * 60)

missing_packages = find_missing_packages(REQUIRED_BUILD_PACKAGES)
if missing_packages:
    print("\nERROR: Missing required build dependencies:")
    for package_name in missing_packages:
        print(f"  - {package_name}")
    print("\nInstall them in this Python environment, for example:")
    print(f"  {sys.executable} -m pip install {' '.join(missing_packages)}")
    sys.exit(2)

print("\nPreparing build directory...")
build_output_dir = BUILD_DIR
if clean_build_dir(BUILD_DIR):
    print("  Build directory ready")
else:
    timestamp = int(time.time())
    build_output_dir = BASE_DIR / "freeze_build" / f"cx_freeze_{timestamp}"
    print(f"  Default directory is locked; using fallback: {build_output_dir.name}")

# Step 1: Build the frozen executable
print("\n[1/2] Building frozen executable...")
freeze_log_path = BASE_DIR / "freeze_build.log"
with freeze_log_path.open("w", encoding="utf-8", errors="replace") as freeze_log:
    freeze_log.write(f"Build output directory: {build_output_dir}\n")
    freeze_log.flush()
    try:
        result = subprocess.run(
            [
                sys.executable,
                "freeze_setup.py",
                "build_exe",
                f"--build-exe={build_output_dir}",
            ],
            cwd=BASE_DIR,
            stdout=freeze_log,
            stderr=subprocess.STDOUT,
            text=True,
            timeout=FREEZE_TIMEOUT_SECONDS,
        )
    except subprocess.TimeoutExpired:
        print(f"  Build timed out after {FREEZE_TIMEOUT_SECONDS} seconds")
        print(f"  Freeze log: {freeze_log_path}")
        log_tail = read_log_tail(freeze_log_path)
        if log_tail:
            print("  Last log lines:")
            print(log_tail)
        sys.exit(1)

if result.returncode == 0:
    print("  Build completed successfully!")
else:
    print(f"  Build failed with exit code {result.returncode}")
    print(f"  Freeze log: {freeze_log_path}")
    log_tail = read_log_tail(freeze_log_path)
    if log_tail:
        print("  Last log lines:")
        print(log_tail)
    sys.exit(1)

# Step 2: Create portable ZIP
print(f"\n[2/2] Creating portable ZIP: {ZIP_NAME}")

if not build_output_dir.exists():
    print(f"  ERROR: Build directory not found: {build_output_dir}")
    sys.exit(1)

zip_path = BASE_DIR / ZIP_NAME

# Remove old ZIP if exists
if zip_path.exists():
    os.remove(zip_path)
    print(f"  Removed old ZIP file")

# Create new ZIP
with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
    for root, dirs, files in os.walk(build_output_dir):
        for file in files:
            file_path = Path(root) / file
            arcname = Path("cx_freeze") / file_path.relative_to(build_output_dir)
            zf.write(file_path, arcname)
            
# Get file size
size_mb = zip_path.stat().st_size / (1024 * 1024)

print(f"  Created: {ZIP_NAME}")
print(f"  Size: {size_mb:.2f} MB")

print("\n" + "=" * 60)
print("Repackaging complete!")
print("=" * 60)
