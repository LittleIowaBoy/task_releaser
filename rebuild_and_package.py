#!/usr/bin/env python3
"""
Quick script to rebuild and repackage DocuReader
"""
import subprocess
import sys
import zipfile
import os
import re
from pathlib import Path

BASE_DIR = Path(__file__).parent
BUILD_DIR = BASE_DIR / "freeze_build" / "cx_freeze"


def read_version() -> str:
    try:
        content = (BASE_DIR / "tr_gui.py").read_text(encoding="utf-8")
        match = re.search(r'__version__\s*=\s*["\']([^"\']+)["\']', content)
        return match.group(1) if match else "0.0.0"
    except Exception:
        return "0.0.0"


VERSION = read_version()
ZIP_NAME = f"DocuReader-{VERSION}-portable.zip"

print("=" * 60)
print("Rebuilding and Repackaging DocuReader")
print("=" * 60)

# Step 1: Build the frozen executable
print("\n[1/2] Building frozen executable...")
result = subprocess.run(
    [sys.executable, "freeze_setup.py", "build"],
    cwd=BASE_DIR,
    capture_output=True,
    text=True
)

if result.returncode == 0:
    print("  Build completed successfully!")
else:
    print(f"  Build failed with exit code {result.returncode}")
    print(result.stderr)
    sys.exit(1)

# Step 2: Create portable ZIP
print(f"\n[2/2] Creating portable ZIP: {ZIP_NAME}")

if not BUILD_DIR.exists():
    print(f"  ERROR: Build directory not found: {BUILD_DIR}")
    sys.exit(1)

zip_path = BASE_DIR / ZIP_NAME

# Remove old ZIP if exists
if zip_path.exists():
    os.remove(zip_path)
    print(f"  Removed old ZIP file")

# Create new ZIP
with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
    for root, dirs, files in os.walk(BUILD_DIR):
        for file in files:
            file_path = Path(root) / file
            arcname = file_path.relative_to(BUILD_DIR.parent)
            zf.write(file_path, arcname)
            
# Get file size
size_mb = zip_path.stat().st_size / (1024 * 1024)

print(f"  Created: {ZIP_NAME}")
print(f"  Size: {size_mb:.2f} MB")

print("\n" + "=" * 60)
print("Repackaging complete!")
print("=" * 60)
