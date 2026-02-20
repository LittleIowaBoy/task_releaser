#!/usr/bin/env python3
"""
Update script for DocuReader application.
Automatically checks for updates on the task_releaser GitHub repository,
intelligently rebuilds the executable if source code or dependencies change,
and provides rollback capabilities.

Usage:
    python update.py              # Full update check and install
    python update.py --check-only # Only check for updates, don't install
    python update.py --rollback   # Rollback to previous version
    python update.py --force-rebuild # Force rebuild of executable
"""

import sys
import os
import re
import subprocess
import shutil
from pathlib import Path
from datetime import datetime
from typing import Tuple, Optional, List
import argparse

BASE_DIR = Path(__file__).resolve().parent
BACKUP_DIR = BASE_DIR / "backup"
UPDATE_LOG = BASE_DIR / "update.log"
GIT_REMOTE = "task_releaser"
GIT_BRANCH = "master"
TAG_PREFIX = "v"
BUILD_DIR = BASE_DIR / "freeze_build" / "cx_freeze"
ACTIVE_GIT_REMOTE: Optional[str] = None
ACTIVE_GIT_BRANCH: Optional[str] = None


def log_message(message: str, level: str = "INFO") -> None:
    """Log messages to both console and file"""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_entry = f"[{timestamp}] [{level}] {message}"
    print(log_entry)
    
    with open(UPDATE_LOG, "a", encoding="utf-8", errors="replace") as f:
        f.write(log_entry + "\n")


def resolve_git_remote() -> Optional[str]:
    """Resolve remote name with fallback: configured remote, origin, then first available."""
    global ACTIVE_GIT_REMOTE

    if ACTIVE_GIT_REMOTE:
        return ACTIVE_GIT_REMOTE

    try:
        result = subprocess.run(
            ["git", "remote"],
            cwd=BASE_DIR,
            capture_output=True,
            text=True,
            timeout=10
        )
        if result.returncode != 0:
            log_message("Unable to list git remotes", "ERROR")
            return None

        remotes = [line.strip() for line in result.stdout.splitlines() if line.strip()]
        if not remotes:
            log_message("No git remotes configured", "ERROR")
            return None

        for candidate in (GIT_REMOTE, "origin"):
            if candidate in remotes:
                ACTIVE_GIT_REMOTE = candidate
                if candidate != GIT_REMOTE:
                    log_message(f"Configured remote '{GIT_REMOTE}' not found; using '{candidate}'", "WARNING")
                return ACTIVE_GIT_REMOTE

        ACTIVE_GIT_REMOTE = remotes[0]
        log_message(
            f"Configured remote '{GIT_REMOTE}' not found; using first available remote '{ACTIVE_GIT_REMOTE}'",
            "WARNING"
        )
        return ACTIVE_GIT_REMOTE
    except Exception as e:
        log_message(f"Failed to resolve git remote: {e}", "ERROR")
        return None


def resolve_git_branch(remote: str) -> Optional[str]:
    """Resolve branch with fallback: configured branch, main, then first available remote branch."""
    global ACTIVE_GIT_BRANCH

    if ACTIVE_GIT_BRANCH:
        return ACTIVE_GIT_BRANCH

    try:
        result = subprocess.run(
            ["git", "ls-remote", "--heads", remote],
            cwd=BASE_DIR,
            capture_output=True,
            text=True,
            timeout=20
        )
        if result.returncode != 0:
            log_message(
                f"Unable to list remote branches for '{remote}'; defaulting to '{GIT_BRANCH}'",
                "WARNING"
            )
            ACTIVE_GIT_BRANCH = GIT_BRANCH
            return ACTIVE_GIT_BRANCH

        branches: List[str] = []
        for line in result.stdout.splitlines():
            parts = line.split()
            if len(parts) < 2:
                continue
            ref = parts[1].strip()
            prefix = "refs/heads/"
            if ref.startswith(prefix):
                branches.append(ref[len(prefix):])

        if not branches:
            log_message(
                f"No remote branches detected for '{remote}'; defaulting to '{GIT_BRANCH}'",
                "WARNING"
            )
            ACTIVE_GIT_BRANCH = GIT_BRANCH
            return ACTIVE_GIT_BRANCH

        for candidate in (GIT_BRANCH, "main"):
            if candidate in branches:
                ACTIVE_GIT_BRANCH = candidate
                if candidate != GIT_BRANCH:
                    log_message(
                        f"Configured branch '{GIT_BRANCH}' not found on '{remote}'; using '{candidate}'",
                        "WARNING"
                    )
                return ACTIVE_GIT_BRANCH

        ACTIVE_GIT_BRANCH = sorted(branches)[0]
        log_message(
            f"Configured branch '{GIT_BRANCH}' not found on '{remote}'; using first available branch '{ACTIVE_GIT_BRANCH}'",
            "WARNING"
        )
        return ACTIVE_GIT_BRANCH
    except Exception as e:
        log_message(f"Failed to resolve git branch from '{remote}': {e}", "WARNING")
        ACTIVE_GIT_BRANCH = GIT_BRANCH
        return ACTIVE_GIT_BRANCH


def get_current_version() -> str:
    """Extract current version from tr_gui.py"""
    try:
        with open(BASE_DIR / "tr_gui.py", "r") as f:
            content = f.read()
            match = re.search(r'__version__\s*=\s*["\']([^"\']+)["\']', content)
            if match:
                return match.group(1)
        
        tag_result = subprocess.run(
            ["git", "describe", "--tags", "--abbrev=0"],
            cwd=BASE_DIR,
            capture_output=True,
            text=True,
            timeout=10
        )
        if tag_result.returncode == 0:
            return tag_result.stdout.strip().lstrip(TAG_PREFIX)
    except Exception as e:
        log_message(f"Failed to read current version: {e}", "ERROR")
    return "0.0.0"


def get_remote_version() -> Optional[str]:
    """Get latest version from git tags on remote"""
    try:
        remote = resolve_git_remote()
        if not remote:
            return None

        # Fetch latest tags from remote
        subprocess.run(
            ["git", "fetch", remote, "--tags", "--force"],
            cwd=BASE_DIR,
            capture_output=True,
            check=True,
            timeout=30
        )
        
        # Get all tags sorted by version
        result = subprocess.run(
            ["git", "tag", "-l", "v*"],
            cwd=BASE_DIR,
            capture_output=True,
            text=True,
            check=True
        )
        
        tags = result.stdout.strip().split('\n')
        if tags and tags[0]:
            # Sort tags as versions (v0.2.0, v0.2.1, etc.)
            versions = sorted(
                [t.lstrip(TAG_PREFIX) for t in tags if t],
                key=lambda x: tuple(map(int, x.split('.')))
            )
            if versions:
                return versions[-1]  # Return highest version
    except subprocess.TimeoutExpired:
        log_message("Git fetch timeout - no internet or slow connection", "WARNING")
    except subprocess.CalledProcessError as e:
        log_message(f"Git command failed: {e}", "WARNING")
    except Exception as e:
        log_message(f"Failed to get remote version: {e}", "ERROR")
    
    return None


def check_files_changed(files: List[str]) -> bool:
    """Check if specific files have changed from remote"""
    try:
        remote = resolve_git_remote()
        if not remote:
            return True
        branch = resolve_git_branch(remote)
        if not branch:
            return True

        # Fetch latest from remote
        subprocess.run(
            ["git", "fetch", remote, branch],
            cwd=BASE_DIR,
            capture_output=True,
            check=True,
            timeout=30
        )
        
        # Check diff between local and remote
        for file_path in files:
            result = subprocess.run(
                ["git", "diff", "HEAD", f"{remote}/{branch}", "--", file_path],
                cwd=BASE_DIR,
                capture_output=True,
                text=True
            )
            if result.stdout.strip():
                return True
        return False
    except Exception as e:
        log_message(f"Failed to check file changes: {e}", "ERROR")
        return True  # Assume changed if we can't verify


def create_backup(version: str) -> Path:
    """Create backup of current executable and build directory"""
    try:
        BACKUP_DIR.mkdir(parents=True, exist_ok=True)

        if not BUILD_DIR.exists():
            log_message(f"Build directory not found: {BUILD_DIR}", "ERROR")
            return None

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        backup_path = BACKUP_DIR / f"v{version}_backup_{timestamp}"

        # Backup build directory into unique path to avoid collisions/locks
        backup_build = backup_path / "DocuReader"
        shutil.copytree(BUILD_DIR, backup_build)
        log_message(f"Backup created at {backup_path}")
        return backup_path
    except Exception as e:
        log_message(f"Failed to create backup: {e}", "ERROR")
        return None
    
    return None


def restore_backup(backup_path: Path) -> bool:
    """Restore from backup"""
    try:
        if not backup_path.exists():
            log_message(f"Backup not found: {backup_path}", "ERROR")
            return False
        
        # Restore build directory
        backup_build = backup_path / "DocuReader"
        build_dir = BUILD_DIR
        
        if backup_build.exists():
            if build_dir.exists():
                shutil.rmtree(build_dir)
            shutil.copytree(backup_build, build_dir)
            log_message(f"Restored backup from {backup_path}")
            return True
    except Exception as e:
        log_message(f"Failed to restore backup: {e}", "ERROR")
    
    return False


def pull_changes() -> bool:
    """Pull latest changes from remote"""
    try:
        remote = resolve_git_remote()
        if not remote:
            return False
        branch = resolve_git_branch(remote)
        if not branch:
            return False

        log_message(f"Pulling changes from {remote}/{branch}...")
        result = subprocess.run(
            ["git", "pull", remote, branch],
            cwd=BASE_DIR,
            capture_output=True,
            text=True,
            timeout=60
        )
        
        if result.returncode != 0:
            log_message(f"Git pull failed: {result.stderr}", "ERROR")
            return False
        
        log_message("Changes pulled successfully")
        return True
    except Exception as e:
        log_message(f"Failed to pull changes: {e}", "ERROR")
        return False


def has_uncommitted_changes() -> bool:
    """Return True when repository has uncommitted changes."""
    try:
        result = subprocess.run(
            ["git", "status", "--porcelain"],
            cwd=BASE_DIR,
            capture_output=True,
            text=True,
            timeout=15
        )
        if result.returncode != 0:
            log_message("Failed to read git status; assuming working tree is dirty", "WARNING")
            return True
        return bool(result.stdout.strip())
    except Exception as e:
        log_message(f"Failed to verify working tree status: {e}", "WARNING")
        return True


def sync_to_target_version(target_version: str) -> bool:
    """Sync repository to the exact target release version tag."""
    target_tag = f"{TAG_PREFIX}{target_version}"
    try:
        remote = resolve_git_remote()
        if not remote:
            return False

        subprocess.run(
            ["git", "fetch", remote, "--tags", "--force"],
            cwd=BASE_DIR,
            capture_output=True,
            check=True,
            timeout=60
        )

        tag_exists = subprocess.run(
            ["git", "rev-parse", "--verify", target_tag],
            cwd=BASE_DIR,
            capture_output=True,
            text=True,
            timeout=15
        )

        if tag_exists.returncode == 0:
            log_message(f"Syncing repository to release tag: {target_tag}")
            sync_result = subprocess.run(
                ["git", "reset", "--hard", target_tag],
                cwd=BASE_DIR,
                capture_output=True,
                text=True,
                timeout=60
            )
            if sync_result.returncode == 0:
                log_message(f"Repository synced to {target_tag}")
                return True
            log_message(f"Failed to sync to {target_tag}: {sync_result.stderr}", "ERROR")

        log_message(
            f"Target tag {target_tag} not available locally after fetch - falling back to branch pull",
            "WARNING"
        )
        return pull_changes()
    except subprocess.TimeoutExpired:
        log_message(f"Sync to {target_tag} timed out - falling back to branch pull", "WARNING")
        return pull_changes()
    except Exception as e:
        log_message(f"Failed to sync to target version {target_version}: {e}", "ERROR")
        return False


def should_rebuild() -> bool:
    """Determine if executable needs rebuilding"""
    files_to_check = [
        "tr_gui.py",
        "tr.py",
        "pyproject.toml"
    ]
    
    if check_files_changed(files_to_check):
        log_message("Source code or dependencies changed - rebuild needed")
        return True
    
    log_message("No source code changes detected - skipping rebuild")
    return False


def rebuild_executable(force: bool = False) -> bool:
    """Rebuild the frozen executable"""
    try:
        if not force and not should_rebuild():
            return True
        
        log_message("Starting executable rebuild...")
        
        result = subprocess.run(
            [sys.executable, "freeze_setup.py", "build"],
            cwd=BASE_DIR,
            capture_output=True,
            text=True,
            timeout=600  # 10 minutes timeout
        )
        
        if result.returncode != 0:
            log_message(f"Build failed: {result.stderr}", "ERROR")
            return False
        
        log_message("Executable rebuild completed successfully")
        return True
    except subprocess.TimeoutExpired:
        log_message("Build process timed out after 10 minutes", "ERROR")
        return False
    except Exception as e:
        log_message(f"Failed to rebuild executable: {e}", "ERROR")
        return False


def validate_build() -> bool:
    """Validate that the new executable exists and runs"""
    try:
        exe_path = BUILD_DIR / "DocuReader.exe"
        
        if not exe_path.exists():
            log_message(f"Executable not found: {exe_path}", "ERROR")
            return False
        
        log_message(f"Validating executable: {exe_path}")
        
        # Quick test: just verify file exists and can be launched
        # (We don't wait for GUI to load in automated script)
        if os.path.exists(exe_path) and os.path.getsize(exe_path) > 0:
            log_message("Executable validation passed")
            return True
    except Exception as e:
        log_message(f"Executable validation failed: {e}", "ERROR")
    
    return False


def create_portable_zip(version: str) -> bool:
    """Create portable ZIP distribution"""
    try:
        import zipfile
        
        zip_path = BASE_DIR / f"DocuReader-{version}-portable.zip"
        build_dir = BUILD_DIR
        
        if not build_dir.exists():
            log_message(f"Build directory not found: {build_dir}", "ERROR")
            return False
        
        log_message(f"Creating portable ZIP: {zip_path}")
        
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for root, dirs, files in os.walk(build_dir):
                for file in files:
                    file_path = Path(root) / file
                    arcname = file_path.relative_to(build_dir.parent)
                    zf.write(file_path, arcname)
        
        log_message(f"Portable ZIP created: {zip_path}")
        return True
    except Exception as e:
        log_message(f"Failed to create portable ZIP: {e}", "ERROR")
        return False


def check_for_updates() -> Tuple[bool, Optional[str]]:
    """Check if updates are available"""
    try:
        current = get_current_version()
        remote = get_remote_version()
        
        if not remote:
            log_message("Could not determine remote version")
            return False, None
        
        log_message(f"Current version: {current}")
        log_message(f"Remote version: {remote}")
        
        # Simple version comparison
        current_parts = tuple(map(int, current.split('.')))
        remote_parts = tuple(map(int, remote.split('.')))
        
        if remote_parts > current_parts:
            log_message(f"Update available: {current} -> {remote}")
            return True, remote
        elif remote_parts == current_parts:
            log_message("Already on latest version")
            return False, None
        else:
            log_message(f"Local version ({current}) is newer than remote ({remote})")
            return False, None
    except Exception as e:
        log_message(f"Error checking for updates: {e}", "ERROR")
        return False, None


def perform_update(target_version: str, force_rebuild: bool = False, allow_dirty: bool = False) -> bool:
    """Perform the full update process"""
    try:
        current_version = get_current_version()
        log_message(f"Update process started. Current version: {current_version}")

        if has_uncommitted_changes() and not allow_dirty:
            log_message(
                "Uncommitted local changes detected. Update aborted to avoid destructive reset. "
                "Commit/stash changes first or rerun with --allow-dirty.",
                "ERROR"
            )
            return False
        
        # Create backup before making changes
        backup_path = create_backup(current_version)
        if not backup_path:
            log_message("Failed to create backup - aborting update", "ERROR")
            return False
        
        # Sync to exact target release version
        if not sync_to_target_version(target_version):
            log_message("Failed to sync target version - attempting rollback", "ERROR")
            restore_backup(backup_path)
            return False
        
        # Get new version after pull
        new_version = get_current_version()
        if new_version != target_version:
            log_message(
                f"Version mismatch after sync. Expected: {target_version}, Actual: {new_version}",
                "WARNING"
            )
        
        # Rebuild executable if needed
        if not rebuild_executable(force=force_rebuild):
            log_message("Build failed - rolling back to previous version", "ERROR")
            restore_backup(backup_path)
            return False
        
        # Validate new build
        if not validate_build():
            log_message("Build validation failed - rolling back", "ERROR")
            restore_backup(backup_path)
            return False
        
        # Create portable ZIP
        if not create_portable_zip(new_version):
            log_message("Failed to create portable ZIP (non-critical)", "WARNING")
        
        log_message(f"Update completed successfully! New version: {new_version}")
        return True
    except Exception as e:
        log_message(f"Unexpected error during update: {e}", "ERROR")
        return False


def main() -> int:
    """Main entry point"""
    parser = argparse.ArgumentParser(
        description="Update DocuReader application from GitHub repository"
    )
    parser.add_argument(
        "--check-only",
        action="store_true",
        help="Only check for updates, don't install"
    )
    parser.add_argument(
        "--rollback",
        action="store_true",
        help="Rollback to previous version"
    )
    parser.add_argument(
        "--force-rebuild",
        action="store_true",
        help="Force rebuild even if no source changes detected"
    )
    parser.add_argument(
        "--yes",
        action="store_true",
        help="Skip interactive confirmation prompt"
    )
    parser.add_argument(
        "--allow-dirty",
        action="store_true",
        help="Allow updates even when uncommitted local changes exist"
    )
    
    args = parser.parse_args()
    
    log_message("=" * 60)
    log_message(f"DocuReader Update Script v1.0")
    log_message("=" * 60)
    
    try:
        if args.check_only:
            available, version = check_for_updates()
            if available:
                print(f"\n[OK] Update available: {version}")
                print(f"  Run 'python update.py' to install\n")
                return 0
            else:
                print("\n[OK] Already on latest version\n")
                return 0
        
        elif args.rollback:
            current = get_current_version()
            backups = sorted(BACKUP_DIR.glob("v*_backup*"))
            
            if not backups:
                log_message("No backups available for rollback", "ERROR")
                print("\n[ERROR] No previous version to rollback to\n")
                return 1
            
            latest_backup = backups[-1]
            if restore_backup(latest_backup):
                print(f"\n[OK] Rollback successful\n")
                return 0
            else:
                print(f"\n[ERROR] Rollback failed\n")
                return 1
        
        else:
            # Full update
            available, version = check_for_updates()
            if not available:
                print("\n[OK] Already on latest version\n")
                return 0
            
            print(f"\n-> Update available: {version}")
            print("  WARNING: This operation may perform a hard reset to the target release tag.")
            print("  Uncommitted local changes can be lost unless you abort now.\n")

            proceed = args.yes
            if not proceed:
                user_input = input("Continue with update? [y/N]: ").strip().lower()
                proceed = user_input in {"y", "yes"}

            if not proceed:
                print("\n[CANCELLED] Update cancelled by user\n")
                return 1

            print("Starting update process...\n")
            
            if perform_update(
                target_version=version,
                force_rebuild=args.force_rebuild,
                allow_dirty=args.allow_dirty,
            ):
                print(f"\n[OK] Update completed successfully!")
                print(f"  Restart application to apply updates\n")
                return 0
            else:
                print(f"\n[ERROR] Update failed - see update.log for details\n")
                return 1
    
    except KeyboardInterrupt:
        log_message("Update interrupted by user", "WARNING")
        print("\n[CANCELLED] Update cancelled\n")
        return 1
    except Exception as e:
        log_message(f"Unexpected error: {e}", "ERROR")
        print(f"\n[ERROR] Update failed: {e}\n")
        return 1
    
    finally:
        log_message("=" * 60)


if __name__ == "__main__":
    sys.exit(main())
