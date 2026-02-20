import re
import sys
from importlib.util import find_spec
from pathlib import Path

from cx_Freeze import Executable, setup


BASE_DIR = Path(__file__).resolve().parent

if len(sys.argv) <= 1:
    print("Usage: python freeze_setup.py <command>")
    print("")
    print("Common commands:")
    print("  python freeze_setup.py build")
    print("  python freeze_setup.py bdist_msi")
    print("  python freeze_setup.py --help")
    sys.exit(2)

build_exe_options = {
    "packages": [
        "pandas",
        "openpyxl",
        "xlrd", 
        "PyQt6",
    ],
    "includes": ["tr", "tr_gui", "update"],
    "excludes": ["tkinter", "unittest", "test", "build", "dist", "freeze_build"],
    "include_msvcr": False,
    "build_exe": str(BASE_DIR / "freeze_build" / "cx_freeze"),
}


def should_validate_dependencies(argv: list[str]) -> bool:
    build_commands = {"build", "build_exe", "bdist_msi"}
    commands = [arg for arg in argv[1:] if not arg.startswith("-")]
    return any(command in build_commands for command in commands)


def validate_dependencies(package_names: list[str]) -> list[str]:
    missing = []
    for package_name in package_names:
        if find_spec(package_name) is None:
            missing.append(package_name)
    return missing


if should_validate_dependencies(sys.argv):
    missing_packages = validate_dependencies(build_exe_options["packages"])
    if missing_packages:
        print("Missing required build dependencies:")
        for package_name in missing_packages:
            print(f"  - {package_name}")
        print("")
        print(
            "Install them in this Python environment, for example:"
        )
        print(f"  python -m pip install {' '.join(missing_packages)}")
        sys.exit(2)

msi_data = {
    "Shortcut": [
        (
            "DesktopShortcut",
            "DesktopFolder",
            "DocuReader",
            "TARGETDIR",
            "[TARGETDIR]DocuReader.exe",
            None,
            None,
            None,
            None,
            None,
            None,
            "TARGETDIR",
        )
    ]
}

bdist_msi_options = {
    "add_to_path": False,
    "all_users": False,
    "initial_target_dir": r"[ProgramFilesFolder]\\DocuReader",
    "data": msi_data,
}

executables = [
    Executable(
        script="tr_gui.py",
        base="gui",
        target_name="DocuReader.exe",
        shortcut_name="DocuReader",
        shortcut_dir="ProgramMenuFolder",
    ),
    Executable(
        script="update.py",
        base=None,
        target_name="update.exe",
    ),
]

# Read version from tr_gui.py
with open(BASE_DIR / "tr_gui.py", "r") as f:
    content = f.read()
    match = re.search(r'__version__\s*=\s*["\']([^"\']+)["\']', content)
    version = match.group(1) if match else "0.0.0"

setup(
    name="docureader",
    version=version,
    description="Inventory DocuReader GUI and Excel parser utilities.",
    options={
        "build_exe": build_exe_options,
        "bdist_msi": bdist_msi_options,
    },
    executables=executables,
)
