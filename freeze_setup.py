from pathlib import Path

from cx_Freeze import Executable, setup


BASE_DIR = Path(__file__).resolve().parent

build_exe_options = {
    "packages": ["pandas", "openpyxl", "xlrd", "PyQt6"],
    "includes": ["tr", "tr_gui"],
    "excludes": ["tkinter", "unittest", "test", "build", "dist", "freeze_build"],
    "include_msvcr": False,
    "build_exe": str(BASE_DIR / "freeze_build" / "cx_freeze"),
    "include_files": [],
}

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
    )
]

setup(
    name="docureader",
    version="0.1.0",
    description="Inventory DocuReader GUI and Excel parser utilities.",
    options={
        "build_exe": build_exe_options,
        "bdist_msi": bdist_msi_options,
    },
    executables=executables,
)
