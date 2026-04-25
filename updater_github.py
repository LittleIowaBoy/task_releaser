"""GitHub Releases-based auto-updater for DocuReader.

Designed for *frozen* (cx_Freeze) installations on machines that have no
Python toolchain and no git CLI. Workflow:

1. Query ``https://api.github.com/repos/<owner>/<repo>/releases/latest`` (or
   ``/releases`` when pre-releases are allowed) and pick the newest tag.
2. Download the ``DocuReader-<v>-portable.zip`` asset and verify it against
   the published ``SHA256SUMS.txt`` manifest.
3. Extract to ``%LOCALAPPDATA%\\DocuReader\\updates\\<v>\\``.
4. Generate ``_apply_update.cmd`` that waits for the running ``DocuReader.exe``
   to exit, robocopies the staged tree over the install dir, and relaunches.
   Spawning a separate cmd lets us replace the running exe (Windows would
   otherwise refuse).

CLI usage::

    python updater_github.py                  # check + install latest
    python updater_github.py --check-only     # exit 0 if up-to-date, 1 if newer available
    python updater_github.py --include-prereleases
    python updater_github.py --yes            # skip confirmation prompts
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import ssl
import sys
import tempfile
import urllib.error
import urllib.request
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, List, Optional, Tuple

from _release import ASSET_NAME_PATTERN, CHECKSUMS_NAME, GITHUB_OWNER, GITHUB_REPO
from _version import __version__ as CURRENT_VERSION


# ---------------------------------------------------------------------------
# Version comparison (kept dependency-free; we don't pull `packaging` in just
# for this since users may not have it installed in dev environments).
# ---------------------------------------------------------------------------


_TAG_RE = re.compile(r"^v?(\d+)(?:\.(\d+))?(?:\.(\d+))?(?:[-.]?(?P<pre>.+))?$")


def parse_version(tag: str) -> Optional[Tuple[int, int, int, int, str]]:
    """Parse ``v1.2.3`` / ``1.2.3-rc1`` into a tuple suitable for ``max()``.

    Returns ``(major, minor, patch, release_rank, pre)`` where
    ``release_rank`` is 1 for a final release and 0 for a pre-release, so a
    final release sorts higher than its rcs.
    """

    m = _TAG_RE.match(tag.strip())
    if not m:
        return None
    major = int(m.group(1) or 0)
    minor = int(m.group(2) or 0)
    patch = int(m.group(3) or 0)
    pre = m.group("pre") or ""
    rank = 0 if pre else 1
    return (major, minor, patch, rank, pre)


def is_newer(remote_tag: str, current: str) -> bool:
    a = parse_version(remote_tag)
    b = parse_version(current)
    if a is None:
        return False
    if b is None:
        return True
    return a > b


# ---------------------------------------------------------------------------
# GitHub API
# ---------------------------------------------------------------------------


GITHUB_API = "https://api.github.com"
USER_AGENT = f"DocuReader-Updater/{CURRENT_VERSION}"


@dataclass
class ReleaseInfo:
    tag: str
    name: str
    prerelease: bool
    asset_url: Optional[str]
    asset_name: Optional[str]
    checksums_url: Optional[str]


def _http_json(url: str, timeout: int = 30) -> dict:
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT, "Accept": "application/vnd.github+json"})
    ctx = ssl.create_default_context()
    with urllib.request.urlopen(req, timeout=timeout, context=ctx) as resp:
        return json.loads(resp.read().decode("utf-8"))


def _http_download(url: str, dest: Path, timeout: int = 120, progress=None) -> None:
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT, "Accept": "application/octet-stream"})
    ctx = ssl.create_default_context()
    with urllib.request.urlopen(req, timeout=timeout, context=ctx) as resp:
        total = int(resp.headers.get("Content-Length") or 0)
        downloaded = 0
        chunk = 64 * 1024
        with dest.open("wb") as f:
            while True:
                buf = resp.read(chunk)
                if not buf:
                    break
                f.write(buf)
                downloaded += len(buf)
                if progress and total:
                    progress(downloaded, total)


def _select_asset(assets: Iterable[dict], tag: str) -> Tuple[Optional[dict], Optional[dict]]:
    """Return (zip asset, checksums asset) for a release."""
    version = tag.lstrip("v")
    expected_zip = ASSET_NAME_PATTERN.format(version=version).lower()
    zip_asset = None
    sums_asset = None
    for a in assets:
        name = (a.get("name") or "").lower()
        if name == expected_zip or (zip_asset is None and name.endswith(".zip") and version in name):
            zip_asset = a
        elif name == CHECKSUMS_NAME.lower():
            sums_asset = a
    return zip_asset, sums_asset


def fetch_release(include_prereleases: bool = False) -> Optional[ReleaseInfo]:
    """Return metadata for the highest-versioned release, or None on failure."""
    try:
        if include_prereleases:
            data = _http_json(f"{GITHUB_API}/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases")
            if not isinstance(data, list) or not data:
                return None
            ranked = []
            for r in data:
                tag = r.get("tag_name", "")
                parsed = parse_version(tag)
                if parsed is not None:
                    ranked.append((parsed, r))
            if not ranked:
                return None
            ranked.sort(key=lambda x: x[0])
            release = ranked[-1][1]
        else:
            release = _http_json(f"{GITHUB_API}/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest")
    except (urllib.error.URLError, urllib.error.HTTPError, json.JSONDecodeError, OSError) as e:
        print(f"[updater] Could not query GitHub: {e}", file=sys.stderr)
        return None

    tag = release.get("tag_name", "")
    assets = release.get("assets", []) or []
    zip_asset, sums_asset = _select_asset(assets, tag)
    return ReleaseInfo(
        tag=tag,
        name=release.get("name") or tag,
        prerelease=bool(release.get("prerelease")),
        asset_url=(zip_asset or {}).get("browser_download_url"),
        asset_name=(zip_asset or {}).get("name"),
        checksums_url=(sums_asset or {}).get("browser_download_url"),
    )


# ---------------------------------------------------------------------------
# Verify + stage + apply
# ---------------------------------------------------------------------------


def _sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(64 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def verify_checksum(zip_path: Path, checksums_text: str) -> bool:
    """Match ``zip_path`` against an entry in a ``sha256sum``-style manifest.

    Manifest format: ``<hex>  <filename>`` per line. Matching is by basename.
    Returns True if the checksum is present *and* matches; False otherwise.
    """
    expected: Optional[str] = None
    target_name = zip_path.name.lower()
    for line in checksums_text.splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        parts = line.split(None, 1)
        if len(parts) != 2:
            continue
        digest, name = parts
        if name.strip().lstrip("*").lower().endswith(target_name):
            expected = digest.lower()
            break
    if expected is None:
        return False
    return _sha256(zip_path).lower() == expected


def staging_root() -> Path:
    base = os.environ.get("LOCALAPPDATA") or str(Path.home() / "AppData" / "Local")
    return Path(base) / "DocuReader" / "updates"


def stage_release(release: ReleaseInfo, progress=None) -> Optional[Path]:
    """Download + verify + extract. Returns the extracted directory, or None."""
    if not release.asset_url:
        print("[updater] Release has no portable-zip asset attached.", file=sys.stderr)
        return None

    target_root = staging_root() / release.tag.lstrip("v")
    if target_root.exists():
        shutil.rmtree(target_root, ignore_errors=True)
    target_root.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory() as tmp:
        tmp_dir = Path(tmp)
        zip_path = tmp_dir / (release.asset_name or "DocuReader.zip")
        print(f"[updater] Downloading {release.asset_url} ...")
        _http_download(release.asset_url, zip_path, progress=progress)

        if release.checksums_url:
            sums_path = tmp_dir / "SHA256SUMS.txt"
            try:
                _http_download(release.checksums_url, sums_path)
                ok = verify_checksum(zip_path, sums_path.read_text(encoding="utf-8"))
                if not ok:
                    print("[updater] SHA-256 mismatch - aborting.", file=sys.stderr)
                    shutil.rmtree(target_root, ignore_errors=True)
                    return None
                print("[updater] Checksum OK.")
            except (urllib.error.URLError, urllib.error.HTTPError, OSError) as e:
                print(f"[updater] Could not fetch checksums ({e}); refusing to apply unverified update.", file=sys.stderr)
                shutil.rmtree(target_root, ignore_errors=True)
                return None
        else:
            print("[updater] No SHA256SUMS.txt published; refusing to apply unverified update.", file=sys.stderr)
            shutil.rmtree(target_root, ignore_errors=True)
            return None

        with zipfile.ZipFile(zip_path) as zf:
            zf.extractall(target_root)

    # Many portable zips contain a single top-level "DocuReader" folder; flatten it.
    entries = list(target_root.iterdir())
    if len(entries) == 1 and entries[0].is_dir():
        return entries[0]
    return target_root


# ---------------------------------------------------------------------------
# Apply (Windows-only swap script)
# ---------------------------------------------------------------------------


_APPLY_CMD_TEMPLATE = r"""@echo off
setlocal
set STAGED=%~1
set INSTALL=%~2
set EXE=%INSTALL%\DocuReader.exe

echo Waiting for DocuReader.exe to exit...
:waitloop
tasklist /FI "IMAGENAME eq DocuReader.exe" 2>NUL | find /I "DocuReader.exe" >NUL
if not errorlevel 1 (
    timeout /T 1 /NOBREAK >NUL
    goto waitloop
)

echo Copying staged files...
robocopy "%STAGED%" "%INSTALL%" /MIR /NFL /NDL /NJH /NJS /NP /R:3 /W:2
if errorlevel 8 (
    echo Update failed during file copy. >&2
    pause
    exit /b 1
)

echo Update applied. Relaunching DocuReader...
start "" "%EXE%"
endlocal
exit /b 0
"""


def write_apply_script(staged_dir: Path) -> Path:
    """Write the swap-and-relaunch script to a temp file and return its path."""
    fd, name = tempfile.mkstemp(prefix="docureader_apply_", suffix=".cmd")
    os.close(fd)
    Path(name).write_text(_APPLY_CMD_TEMPLATE, encoding="ascii")
    return Path(name)


def apply_update(staged_dir: Path, install_dir: Path) -> int:
    """Spawn the swap script and exit so the running exe can be replaced."""
    if os.name != "nt":
        print("[updater] Auto-apply is only supported on Windows. "
              f"Manually copy {staged_dir} -> {install_dir}.", file=sys.stderr)
        return 1
    script = write_apply_script(staged_dir)
    print(f"[updater] Spawning swap script: {script}")
    # ``start`` detaches the cmd window so we can exit immediately.
    os.system(f'start "" cmd /c "{script}" "{staged_dir}" "{install_dir}"')
    return 0


def install_dir_for_running_exe() -> Path:
    """Return the directory that should be overwritten by the staged update."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def main(argv: Optional[List[str]] = None) -> int:
    p = argparse.ArgumentParser(description="Check for and install DocuReader updates from GitHub Releases.")
    p.add_argument("--check-only", action="store_true", help="Exit 1 if a newer release is available, 0 otherwise.")
    p.add_argument("--include-prereleases", action="store_true", help="Consider pre-release tags too.")
    p.add_argument("--yes", action="store_true", help="Do not prompt before installing.")
    args = p.parse_args(argv)

    print(f"[updater] Current version: {CURRENT_VERSION}")
    release = fetch_release(include_prereleases=args.include_prereleases)
    if release is None:
        print("[updater] Could not determine the latest release.", file=sys.stderr)
        return 3
    print(f"[updater] Latest release: {release.tag} ({'prerelease' if release.prerelease else 'stable'})")

    if not is_newer(release.tag, CURRENT_VERSION):
        print("[updater] Already on the latest version.")
        return 2 if args.check_only else 0
    if args.check_only:
        return 1

    if not args.yes:
        try:
            answer = input(f"Install {release.tag}? [y/N] ").strip().lower()
        except EOFError:
            answer = ""
        if answer not in ("y", "yes"):
            print("[updater] Aborted.")
            return 0

    staged = stage_release(release)
    if staged is None:
        return 4

    install_dir = install_dir_for_running_exe()
    print(f"[updater] Staged at: {staged}")
    print(f"[updater] Install dir: {install_dir}")
    return apply_update(staged, install_dir)


if __name__ == "__main__":
    sys.exit(main())
