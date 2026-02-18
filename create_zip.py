#!/usr/bin/env python3
"""Quick ZIP packaging script"""
import zipfile
import os
from pathlib import Path

BUILD_DIR = Path("freeze_build/cx_freeze")
ZIP_NAME = "DocuReader-0.2.0-portable.zip"

print(f"Creating {ZIP_NAME} from {BUILD_DIR}...")

if not BUILD_DIR.exists():
    print(f"ERROR: Build directory not found: {BUILD_DIR}")
    exit(1)

# Remove old ZIP if exists
if os.path.exists(ZIP_NAME):
    os.remove(ZIP_NAME)
    print("Removed old ZIP")

# Collect all files first
files_to_add = []
for root, dirs, files in os.walk(BUILD_DIR):
    for file in files:
        file_path = Path(root) / file
        arcname = file_path.relative_to(BUILD_DIR.parent)
        files_to_add.append((file_path, arcname))

print(f"Found {len(files_to_add)} files to package")

# Create new ZIP with NO compression (STORE mode) for speed
with zipfile.ZipFile(ZIP_NAME, 'w', compression=zipfile.ZIP_STORED) as zf:
    for i, (file_path, arcname) in enumerate(files_to_add, 1):
        zf.write(file_path, arcname)
        if i % 100 == 0:
            print(f"  Added {i}/{len(files_to_add)} files...")
            
# Print result
size_mb = Path(ZIP_NAME).stat().st_size / (1024 * 1024)
print(f"[OK] Created {ZIP_NAME}: {size_mb:.2f} MB")
