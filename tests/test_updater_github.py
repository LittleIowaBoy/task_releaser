import hashlib
import io
import os
import subprocess
import zipfile
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

import updater_github


# ---------------------------------------------------------------------------
# Version parsing
# ---------------------------------------------------------------------------

def test_parse_version_orderings():
    assert updater_github.parse_version("v1.2.3") > updater_github.parse_version("v1.2.2")
    assert updater_github.parse_version("v0.3.0") > updater_github.parse_version("v0.2.4")
    # Final release sorts above its rc.
    assert updater_github.parse_version("v1.0.0") > updater_github.parse_version("v1.0.0-rc1")
    assert updater_github.parse_version("not-a-tag") is None


def test_is_newer():
    assert updater_github.is_newer("v0.3.0", "0.2.4")
    assert not updater_github.is_newer("v0.2.4", "0.2.4")
    assert not updater_github.is_newer("v0.2.0", "0.2.4")


# ---------------------------------------------------------------------------
# Checksum verification
# ---------------------------------------------------------------------------

def test_verify_checksum_match(tmp_path: Path):
    payload = b"hello-docureader"
    f = tmp_path / "DocuReader-9.9.9-portable.zip"
    f.write_bytes(payload)
    digest = hashlib.sha256(payload).hexdigest()
    manifest = f"{digest}  DocuReader-9.9.9-portable.zip\n"
    assert updater_github.verify_checksum(f, manifest) is True


def test_verify_checksum_mismatch(tmp_path: Path):
    f = tmp_path / "z.zip"
    f.write_bytes(b"abc")
    bad = "0" * 64 + "  z.zip\n"
    assert updater_github.verify_checksum(f, bad) is False


def test_verify_checksum_ignores_comment_lines(tmp_path: Path):
    payload = b"data"
    f = tmp_path / "app.zip"
    f.write_bytes(payload)
    digest = hashlib.sha256(payload).hexdigest()
    manifest = f"# this is a comment\n{digest}  app.zip\n"
    assert updater_github.verify_checksum(f, manifest) is True


def test_verify_checksum_missing_entry(tmp_path: Path):
    f = tmp_path / "missing.zip"
    f.write_bytes(b"x")
    assert updater_github.verify_checksum(f, "") is False


# ---------------------------------------------------------------------------
# Asset selection
# ---------------------------------------------------------------------------

def test_select_asset_picks_named_zip_and_sums():
    assets = [
        {"name": "DocuReader-0.3.0-portable.zip", "browser_download_url": "u1"},
        {"name": "SHA256SUMS.txt", "browser_download_url": "u2"},
        {"name": "noise.txt", "browser_download_url": "u3"},
    ]
    z, s = updater_github._select_asset(assets, "v0.3.0")
    assert z["browser_download_url"] == "u1"
    assert s["browser_download_url"] == "u2"


def test_select_asset_no_sums():
    assets = [{"name": "DocuReader-1.0.0-portable.zip", "browser_download_url": "u1"}]
    z, s = updater_github._select_asset(assets, "v1.0.0")
    assert z is not None
    assert s is None


def test_select_asset_empty_list():
    z, s = updater_github._select_asset([], "v1.0.0")
    assert z is None
    assert s is None


# ---------------------------------------------------------------------------
# Staging (download + extract)
# ---------------------------------------------------------------------------

def _make_fake_zip(*files: tuple) -> bytes:
    """Build an in-memory zip; *files* is [(name, content), ...]."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        for name, content in files:
            zf.writestr(name, content)
    return buf.getvalue()


def test_stage_release_with_monkeypatched_http(tmp_path: Path, monkeypatch):
    zip_bytes = _make_fake_zip(("DocuReader/marker.txt", "ok"))
    digest = hashlib.sha256(zip_bytes).hexdigest()

    def fake_download(url, dest, timeout=120, progress=None):
        if url.endswith("zip"):
            Path(dest).write_bytes(zip_bytes)
        else:
            Path(dest).write_text(f"{digest}  fake.zip\n", encoding="utf-8")

    monkeypatch.setattr(updater_github, "_http_download", fake_download)
    monkeypatch.setattr(updater_github, "staging_root", lambda: tmp_path / "staged")

    rel = updater_github.ReleaseInfo(
        tag="v9.9.9",
        name="test",
        prerelease=False,
        asset_url="https://example/fake.zip",
        asset_name="fake.zip",
        checksums_url="https://example/SHA256SUMS.txt",
    )
    staged = updater_github.stage_release(rel)
    assert staged is not None
    # The zip contained a single top-level "DocuReader/" so stage_release flattens.
    assert (staged / "marker.txt").exists()


def test_stage_release_reports_progress(tmp_path: Path, monkeypatch):
    """stage_release must forward the progress callback to _http_download."""
    zip_bytes = _make_fake_zip(("DocuReader/app.exe", "x" * 512))
    digest = hashlib.sha256(zip_bytes).hexdigest()
    progress_calls: list[tuple[int, int]] = []

    def fake_download(url, dest, timeout=120, progress=None):
        if url.endswith("zip"):
            Path(dest).write_bytes(zip_bytes)
            if progress:
                progress(len(zip_bytes) // 2, len(zip_bytes))
                progress(len(zip_bytes), len(zip_bytes))
        else:
            Path(dest).write_text(f"{digest}  fake.zip\n", encoding="utf-8")

    monkeypatch.setattr(updater_github, "_http_download", fake_download)
    monkeypatch.setattr(updater_github, "staging_root", lambda: tmp_path / "staged")

    rel = updater_github.ReleaseInfo(
        tag="v9.9.9",
        name="test",
        prerelease=False,
        asset_url="https://example/fake.zip",
        asset_name="fake.zip",
        checksums_url="https://example/SHA256SUMS.txt",
    )
    staged = updater_github.stage_release(
        rel, progress=lambda d, t: progress_calls.append((d, t))
    )
    assert staged is not None
    assert len(progress_calls) == 2
    assert progress_calls[-1][0] == progress_calls[-1][1]  # final call: downloaded == total


def test_stage_release_aborts_on_checksum_mismatch(tmp_path: Path, monkeypatch):
    zip_bytes = _make_fake_zip(("file.txt", "data"))

    def fake_download(url, dest, timeout=120, progress=None):
        if url.endswith("zip"):
            Path(dest).write_bytes(zip_bytes)
        else:
            # Deliberately wrong digest.
            Path(dest).write_text("0" * 64 + "  fake.zip\n", encoding="utf-8")

    monkeypatch.setattr(updater_github, "_http_download", fake_download)
    monkeypatch.setattr(updater_github, "staging_root", lambda: tmp_path / "staged")

    rel = updater_github.ReleaseInfo(
        tag="v9.9.9",
        name="test",
        prerelease=False,
        asset_url="https://example/fake.zip",
        asset_name="fake.zip",
        checksums_url="https://example/SHA256SUMS.txt",
    )
    staged = updater_github.stage_release(rel)
    assert staged is None


def test_stage_release_refuses_without_checksums(tmp_path: Path, monkeypatch):
    """If no checksums URL is published the update must be rejected."""
    zip_bytes = _make_fake_zip(("file.txt", "data"))

    def fake_download(url, dest, timeout=120, progress=None):
        Path(dest).write_bytes(zip_bytes)

    monkeypatch.setattr(updater_github, "_http_download", fake_download)
    monkeypatch.setattr(updater_github, "staging_root", lambda: tmp_path / "staged")

    rel = updater_github.ReleaseInfo(
        tag="v9.9.9",
        name="test",
        prerelease=False,
        asset_url="https://example/fake.zip",
        asset_name="fake.zip",
        checksums_url=None,  # no checksums published
    )
    staged = updater_github.stage_release(rel)
    assert staged is None


# ---------------------------------------------------------------------------
# Apply script generation
# ---------------------------------------------------------------------------

def test_build_apply_cmd_embeds_paths(tmp_path: Path):
    staged = tmp_path / "staged_release"
    install = tmp_path / "DocuReader"
    script = updater_github._build_apply_cmd(staged, install)

    assert str(staged) in script
    assert str(install) in script


def test_build_apply_cmd_does_not_use_positional_args(tmp_path: Path):
    """The old template used %~1 / %~2; the new one must not."""
    script = updater_github._build_apply_cmd(tmp_path / "s", tmp_path / "i")
    assert "%~1" not in script
    assert "%~2" not in script


def test_build_apply_cmd_contains_robocopy(tmp_path: Path):
    script = updater_github._build_apply_cmd(tmp_path / "s", tmp_path / "i")
    assert "robocopy" in script.lower()


def test_build_apply_cmd_contains_relaunch(tmp_path: Path):
    script = updater_github._build_apply_cmd(tmp_path / "s", tmp_path / "i")
    assert "DocuReader.exe" in script


def test_write_apply_script_creates_file_with_embedded_paths(tmp_path: Path):
    staged = tmp_path / "staged"
    install = tmp_path / "install"
    script_path = updater_github.write_apply_script(staged, install)
    try:
        assert script_path.exists()
        assert script_path.suffix == ".cmd"
        content = script_path.read_text(encoding="ascii")
        assert str(staged) in content
        assert str(install) in content
    finally:
        script_path.unlink(missing_ok=True)


# ---------------------------------------------------------------------------
# apply_update — subprocess spawning
# ---------------------------------------------------------------------------

def test_apply_update_non_windows_returns_1(tmp_path: Path, monkeypatch):
    monkeypatch.setattr(updater_github.os, "name", "posix")
    rc = updater_github.apply_update(tmp_path / "staged", tmp_path / "install")
    assert rc == 1


def test_apply_update_spawns_popen_on_windows(tmp_path: Path, monkeypatch):
    monkeypatch.setattr(updater_github.os, "name", "nt")
    popen_calls: list = []

    class _FakePopen:
        def __init__(self, *args, **kwargs):
            popen_calls.append((args, kwargs))

    monkeypatch.setattr(updater_github.subprocess, "Popen", _FakePopen)

    staged = tmp_path / "staged"
    install = tmp_path / "install"
    rc = updater_github.apply_update(staged, install)

    assert rc == 0
    assert len(popen_calls) == 1
    cmd_args = popen_calls[0][0][0]          # first positional → the argv list
    assert cmd_args[0] == "cmd.exe"
    assert cmd_args[1] == "/c"
    assert cmd_args[2].endswith(".cmd")


def test_apply_update_uses_detached_process_flag(tmp_path: Path, monkeypatch):
    monkeypatch.setattr(updater_github.os, "name", "nt")
    kwargs_seen: list = []

    class _FakePopen:
        def __init__(self, *args, **kwargs):
            kwargs_seen.append(kwargs)

    monkeypatch.setattr(updater_github.subprocess, "Popen", _FakePopen)
    updater_github.apply_update(tmp_path / "s", tmp_path / "i")

    assert kwargs_seen, "Popen was not called"
    flags = kwargs_seen[0].get("creationflags", 0)
    # Script runs silently in the background — no visible CMD window.
    assert flags & subprocess.CREATE_NO_WINDOW
    assert flags & subprocess.CREATE_NEW_PROCESS_GROUP
    assert not (flags & subprocess.CREATE_NEW_CONSOLE), (
        "CREATE_NEW_CONSOLE would pop up a visible CMD window during update"
    )
    assert not (flags & subprocess.DETACHED_PROCESS), (
        "DETACHED_PROCESS and CREATE_NEW_CONSOLE are mutually exclusive Win32 flags"
    )


def test_apply_update_returns_1_on_popen_oserror(tmp_path: Path, monkeypatch):
    monkeypatch.setattr(updater_github.os, "name", "nt")

    def _bad_popen(*args, **kwargs):
        raise OSError("no cmd.exe found")

    monkeypatch.setattr(updater_github.subprocess, "Popen", _bad_popen)
    rc = updater_github.apply_update(tmp_path / "staged", tmp_path / "install")
    assert rc == 1


def test_apply_update_script_contains_correct_paths(tmp_path: Path, monkeypatch):
    """The .cmd file spawned by apply_update must contain the staged+install paths.

    Paths are stored as their 8.3 short-path equivalents for ASCII safety, so
    we compare against the short-path form rather than the original long path.
    """
    monkeypatch.setattr(updater_github.os, "name", "nt")
    script_paths: list[Path] = []

    class _FakePopen:
        def __init__(self, argv, **kwargs):
            script_paths.append(Path(argv[2]))

    monkeypatch.setattr(updater_github.subprocess, "Popen", _FakePopen)

    staged = tmp_path / "my staged dir"
    install = tmp_path / "my install dir"
    staged.mkdir()
    install.mkdir()
    updater_github.apply_update(staged, install)

    assert script_paths, "Popen was not called"
    content = script_paths[0].read_text(encoding="ascii")
    # _get_short_path converts to 8.3 form; check against the same form.
    assert updater_github._get_short_path(staged) in content
    assert updater_github._get_short_path(install) in content
    script_paths[0].unlink(missing_ok=True)


# ---------------------------------------------------------------------------
# install_dir_for_running_exe
# ---------------------------------------------------------------------------

def test_install_dir_for_running_exe_non_frozen(monkeypatch):
    monkeypatch.setattr(updater_github.sys, "frozen", False, raising=False)
    result = updater_github.install_dir_for_running_exe()
    assert result == Path(updater_github.__file__).resolve().parent


def test_install_dir_for_running_exe_frozen(monkeypatch, tmp_path: Path):
    fake_exe = tmp_path / "DocuReader.exe"
    fake_exe.touch()
    monkeypatch.setattr(updater_github.sys, "frozen", True, raising=False)
    monkeypatch.setattr(updater_github.sys, "executable", str(fake_exe))
    result = updater_github.install_dir_for_running_exe()
    assert result == tmp_path


# ===========================================================================
# New tests — update-flow simulation
# ===========================================================================
#
# These tests simulate the full in-app update journey a user triggers via the
# "Check && Install Updates" menu item:
#
#   1. fetch_release  — GitHub API query returns a newer release
#   2. stage_release  — ZIP downloaded, checksum verified, extracted
#   3. apply_update   — .cmd swap script written and spawned correctly
#   4. Restart        — the spawned .cmd script contains all required restart
#                       steps in the right order
#
# All network calls are replaced with in-process fakes so no internet
# connectivity is required.
# ===========================================================================


# ---------------------------------------------------------------------------
# Helpers shared by multiple tests below
# ---------------------------------------------------------------------------

def _make_release_zip(*extra_files: tuple) -> bytes:
    """Return a ZIP bytes object that mimics a real DocuReader portable archive.

    The archive has the canonical top-level ``DocuReader/`` folder so
    ``stage_release`` flattens it correctly.  A minimal ``DocuReader.exe``
    stub is always included; caller may add more ``(name, content)`` pairs.
    """
    files = [("DocuReader/DocuReader.exe", b"MZ-stub")] + list(extra_files)
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        for name, content in files:
            data = content if isinstance(content, bytes) else content.encode()
            zf.writestr(name, data)
    return buf.getvalue()


def _release_info(tag: str = "v9.9.9", *, include_checksums: bool = True) -> updater_github.ReleaseInfo:
    return updater_github.ReleaseInfo(
        tag=tag,
        name=f"DocuReader {tag}",
        prerelease=False,
        asset_url=f"https://example/DocuReader-{tag.lstrip('v')}-portable.zip",
        asset_name=f"DocuReader-{tag.lstrip('v')}-portable.zip",
        checksums_url="https://example/SHA256SUMS.txt" if include_checksums else None,
    )


def _make_fake_downloader(zip_bytes: bytes) -> "callable":
    """Return a drop-in replacement for ``_http_download`` that serves pre-baked bytes."""
    digest = hashlib.sha256(zip_bytes).hexdigest()

    def _download(url: str, dest: Path, timeout: int = 120, progress=None) -> None:
        if url.endswith(".zip"):
            Path(dest).write_bytes(zip_bytes)
            if progress:
                half = len(zip_bytes) // 2
                progress(half, len(zip_bytes))
                progress(len(zip_bytes), len(zip_bytes))
        else:
            # SHA256SUMS — include the correct digest so verification passes.
            name = Path(url.rsplit("/", 1)[-1]).name
            zip_name = url.rsplit("/", 1)[-1].replace("SHA256SUMS.txt", "").rstrip("/")
            # The manifest just needs the zip filename somewhere on the line.
            asset_name = url.replace("SHA256SUMS.txt", "").rstrip("/").rsplit("/", 1)[-1]
            Path(dest).write_text(f"{digest}  some-asset.zip\n", encoding="utf-8")

    return _download


# ---------------------------------------------------------------------------
# 1. fetch_release — GitHub API
# ---------------------------------------------------------------------------

class TestFetchRelease:
    """Simulate the GitHub API response for a newer stable release."""

    def _api_response(self, tag: str = "v0.7.0") -> dict:
        return {
            "tag_name": tag,
            "name": f"DocuReader {tag}",
            "prerelease": False,
            "assets": [
                {
                    "name": f"DocuReader-{tag.lstrip('v')}-portable.zip",
                    "browser_download_url": f"https://example/{tag.lstrip('v')}.zip",
                },
                {
                    "name": "SHA256SUMS.txt",
                    "browser_download_url": "https://example/SHA256SUMS.txt",
                },
            ],
        }

    def test_fetch_release_returns_release_info(self, monkeypatch):
        monkeypatch.setattr(updater_github, "_http_json", lambda url, timeout=30: self._api_response())
        rel = updater_github.fetch_release()
        assert rel is not None
        assert rel.tag == "v0.7.0"
        assert rel.asset_url is not None
        assert rel.checksums_url is not None

    def test_fetch_release_detects_newer_version(self, monkeypatch):
        monkeypatch.setattr(updater_github, "_http_json", lambda url, timeout=30: self._api_response("v99.0.0"))
        rel = updater_github.fetch_release()
        assert rel is not None
        assert updater_github.is_newer(rel.tag, updater_github.CURRENT_VERSION)

    def test_fetch_release_returns_none_on_network_error(self, monkeypatch):
        import urllib.error
        def _fail(url, timeout=30):
            raise urllib.error.URLError("simulated network failure")
        monkeypatch.setattr(updater_github, "_http_json", _fail)
        rel = updater_github.fetch_release()
        assert rel is None

    def test_fetch_release_returns_none_on_bad_json(self, monkeypatch):
        import json
        def _bad(url, timeout=30):
            raise json.JSONDecodeError("bad json", "", 0)
        monkeypatch.setattr(updater_github, "_http_json", _bad)
        assert updater_github.fetch_release() is None

    def test_fetch_release_prerelease_flag(self, monkeypatch):
        """When include_prereleases=True the /releases list endpoint is queried."""
        called_urls: list[str] = []

        def _spy_json(url, timeout=30):
            called_urls.append(url)
            return [self._api_response("v0.8.0-rc1")]

        monkeypatch.setattr(updater_github, "_http_json", _spy_json)
        rel = updater_github.fetch_release(include_prereleases=True)
        assert rel is not None
        assert any("/releases" in u and "latest" not in u for u in called_urls)

    def test_fetch_release_no_matching_asset(self, monkeypatch):
        """A release with no zip asset still returns a ReleaseInfo (asset_url is None)."""
        data = {
            "tag_name": "v0.7.0",
            "name": "DocuReader v0.7.0",
            "prerelease": False,
            "assets": [],
        }
        monkeypatch.setattr(updater_github, "_http_json", lambda url, timeout=30: data)
        rel = updater_github.fetch_release()
        assert rel is not None
        assert rel.asset_url is None


# ---------------------------------------------------------------------------
# 2. stage_release — download, checksum, extract
# ---------------------------------------------------------------------------

class TestStageRelease:
    """Simulate the download + verify + extract pipeline."""

    def test_stage_downloads_and_extracts(self, tmp_path: Path, monkeypatch):
        zip_bytes = _make_release_zip(("DocuReader/data.txt", "payload"))
        digest = hashlib.sha256(zip_bytes).hexdigest()
        rel = _release_info()

        def _dl(url, dest, timeout=120, progress=None):
            if url.endswith(".zip"):
                Path(dest).write_bytes(zip_bytes)
            else:
                # Manifest must match the asset_name so verify_checksum passes.
                Path(dest).write_text(f"{digest}  {rel.asset_name}\n", encoding="utf-8")

        monkeypatch.setattr(updater_github, "_http_download", _dl)
        monkeypatch.setattr(updater_github, "staging_root", lambda: tmp_path / "staged")

        staged = updater_github.stage_release(rel)
        assert staged is not None
        assert (staged / "DocuReader.exe").exists()
        assert (staged / "data.txt").exists()

    def test_stage_flattens_single_top_level_dir(self, tmp_path: Path, monkeypatch):
        """The canonical archive nests files under DocuReader/; stage_release flattens it."""
        zip_bytes = _make_release_zip()
        digest = hashlib.sha256(zip_bytes).hexdigest()
        rel = _release_info()

        def _dl(url, dest, timeout=120, progress=None):
            if url.endswith(".zip"):
                Path(dest).write_bytes(zip_bytes)
            else:
                Path(dest).write_text(f"{digest}  {rel.asset_name}\n", encoding="utf-8")

        monkeypatch.setattr(updater_github, "_http_download", _dl)
        monkeypatch.setattr(updater_github, "staging_root", lambda: tmp_path / "staged")

        staged = updater_github.stage_release(rel)
        # After flattening the returned path should contain the exe directly, not in a sub-folder.
        assert staged is not None
        assert (staged / "DocuReader.exe").exists()

    def test_stage_calls_progress_callback(self, tmp_path: Path, monkeypatch):
        zip_bytes = _make_release_zip()
        digest = hashlib.sha256(zip_bytes).hexdigest()
        calls: list[tuple[int, int]] = []

        def _dl(url, dest, timeout=120, progress=None):
            if url.endswith(".zip"):
                Path(dest).write_bytes(zip_bytes)
                if progress:
                    progress(len(zip_bytes) // 2, len(zip_bytes))
                    progress(len(zip_bytes), len(zip_bytes))
            else:
                Path(dest).write_text(f"{digest}  fake.zip\n", encoding="utf-8")

        monkeypatch.setattr(updater_github, "_http_download", _dl)
        monkeypatch.setattr(updater_github, "staging_root", lambda: tmp_path / "staged")

        updater_github.stage_release(_release_info(), progress=lambda d, t: calls.append((d, t)))
        assert len(calls) == 2
        # Final callback must report downloaded == total (100 %).
        assert calls[-1][0] == calls[-1][1]

    def test_stage_rejects_bad_checksum(self, tmp_path: Path, monkeypatch):
        zip_bytes = _make_release_zip()

        def _dl(url, dest, timeout=120, progress=None):
            if url.endswith(".zip"):
                Path(dest).write_bytes(zip_bytes)
            else:
                Path(dest).write_text("0" * 64 + "  fake.zip\n", encoding="utf-8")

        monkeypatch.setattr(updater_github, "_http_download", _dl)
        monkeypatch.setattr(updater_github, "staging_root", lambda: tmp_path / "staged")

        assert updater_github.stage_release(_release_info()) is None

    def test_stage_rejects_missing_checksums(self, tmp_path: Path, monkeypatch):
        zip_bytes = _make_release_zip()

        def _dl(url, dest, timeout=120, progress=None):
            Path(dest).write_bytes(zip_bytes)

        monkeypatch.setattr(updater_github, "_http_download", _dl)
        monkeypatch.setattr(updater_github, "staging_root", lambda: tmp_path / "staged")

        assert updater_github.stage_release(_release_info(include_checksums=False)) is None

    def test_stage_returns_none_when_no_asset_url(self, tmp_path: Path, monkeypatch):
        rel = _release_info()
        rel = updater_github.ReleaseInfo(
            tag=rel.tag, name=rel.name, prerelease=rel.prerelease,
            asset_url=None, asset_name=None, checksums_url=rel.checksums_url,
        )
        # _http_download must never be called.
        monkeypatch.setattr(updater_github, "_http_download", lambda *a, **k: (_ for _ in ()).throw(AssertionError("should not download")))
        monkeypatch.setattr(updater_github, "staging_root", lambda: tmp_path / "staged")
        assert updater_github.stage_release(rel) is None

    def test_stage_removes_stale_staging_dir(self, tmp_path: Path, monkeypatch):
        """A leftover staging directory from a previous attempt must be cleaned up."""
        zip_bytes = _make_release_zip()
        digest = hashlib.sha256(zip_bytes).hexdigest()
        rel = _release_info()
        stale_file = tmp_path / "staged" / "9.9.9" / "stale.txt"
        stale_file.parent.mkdir(parents=True)
        stale_file.write_text("old")

        def _dl(url, dest, timeout=120, progress=None):
            if url.endswith(".zip"):
                Path(dest).write_bytes(zip_bytes)
            else:
                Path(dest).write_text(f"{digest}  {rel.asset_name}\n", encoding="utf-8")

        monkeypatch.setattr(updater_github, "_http_download", _dl)
        monkeypatch.setattr(updater_github, "staging_root", lambda: tmp_path / "staged")

        staged = updater_github.stage_release(rel)
        assert staged is not None
        assert not stale_file.exists(), "stale file from previous staging should have been removed"


# ---------------------------------------------------------------------------
# 3. Apply script — new content checks (additions since last refactor)
# ---------------------------------------------------------------------------

class TestApplyScriptContent:
    """Verify that the generated .cmd contains all required restart logic."""

    def test_script_uses_short_paths_when_ascii_unsafe(self, tmp_path: Path, monkeypatch):
        """_get_short_path is called on both path args."""
        calls: list[Path] = []
        original = updater_github._get_short_path

        def _spy(p):
            calls.append(p)
            return original(p)

        monkeypatch.setattr(updater_github, "_get_short_path", _spy)
        staged = tmp_path / "staged"
        install = tmp_path / "install"
        updater_github._build_apply_cmd(staged, install)
        assert staged in calls
        assert install in calls

    def test_script_has_log_file_reference(self, tmp_path: Path):
        script = updater_github._build_apply_cmd(tmp_path / "s", tmp_path / "i")
        assert "docureader_update.log" in script

    def test_script_logs_start_and_exit_and_relaunch(self, tmp_path: Path):
        script = updater_github._build_apply_cmd(tmp_path / "s", tmp_path / "i")
        assert "Update script started" in script
        assert "DocuReader.exe exited" in script
        assert "Relaunching DocuReader" in script

    def test_script_has_buffer_timeout_before_robocopy(self, tmp_path: Path):
        """A 1-second timeout must appear between the wait-loop and the robocopy call."""
        script = updater_github._build_apply_cmd(tmp_path / "s", tmp_path / "i")
        wait_loop_end = script.index(")") + 1          # end of if-not-errorlevel block
        robocopy_pos = script.lower().index("robocopy")
        timeout_pos = script.lower().index("timeout /t 1", wait_loop_end)
        assert timeout_pos < robocopy_pos, "1-second buffer must appear before robocopy"

    def test_script_checks_exe_exists_before_start(self, tmp_path: Path):
        """The script must verify DocuReader.exe exists after robocopy before launching."""
        script = updater_github._build_apply_cmd(tmp_path / "s", tmp_path / "i")
        robocopy_pos = script.lower().index("robocopy")
        start_pos = script.lower().index('start ""')
        # "if not exist" check for the exe must appear between robocopy and start.
        assert 'if not exist "%EXE%"' in script
        exe_check_pos = script.index('if not exist "%EXE%"')
        assert robocopy_pos < exe_check_pos < start_pos

    def test_script_endlocal_before_exit(self, tmp_path: Path):
        """endlocal must appear before the final exit /b 0 (was dead code before fix)."""
        script = updater_github._build_apply_cmd(tmp_path / "s", tmp_path / "i")
        endlocal_pos = script.rindex("endlocal")
        exit_pos = script.rindex("exit /b 0")
        assert endlocal_pos < exit_pos

    def test_script_is_ascii_encodable(self, tmp_path: Path):
        """The generated script must be fully ASCII when paths are ASCII."""
        staged = tmp_path / "staged_dir"
        install = tmp_path / "install_dir"
        staged.mkdir()
        install.mkdir()
        script = updater_github._build_apply_cmd(staged, install)
        script.encode("ascii")  # raises UnicodeEncodeError if non-ASCII present

    def test_write_apply_script_file_is_ascii(self, tmp_path: Path):
        staged = tmp_path / "s"
        install = tmp_path / "i"
        staged.mkdir()
        install.mkdir()
        p = updater_github.write_apply_script(staged, install)
        try:
            content = p.read_text(encoding="ascii")
            assert "robocopy" in content.lower()
        finally:
            p.unlink(missing_ok=True)

    def test_write_apply_script_fallback_on_unicode_encode_error(self, tmp_path: Path, monkeypatch):
        """If ascii encoding fails (short-path unavailable), cp1252 fallback is used."""
        monkeypatch.setattr(updater_github, "_get_short_path", lambda p: str(p))
        # Inject a path that contains a non-ASCII character to trigger the fallback.
        staged = tmp_path / "caf\u00e9"   # 'café'
        install = tmp_path / "install"
        staged.mkdir()
        install.mkdir()
        p = updater_github.write_apply_script(staged, install)
        try:
            assert p.exists()
            # The file must be readable under cp1252 regardless.
            content = p.read_text(encoding="cp1252")
            assert "robocopy" in content.lower()
        finally:
            p.unlink(missing_ok=True)


# ---------------------------------------------------------------------------
# 4. apply_update — restart behaviour
# ---------------------------------------------------------------------------

class TestApplyUpdateRestart:
    """Verify the full restart path: script written → spawned → app will relaunch."""

    def test_apply_update_script_has_start_command(self, tmp_path: Path, monkeypatch):
        """The .cmd written to disk must contain a 'start "" "%EXE%"' launch line."""
        monkeypatch.setattr(updater_github.os, "name", "nt")
        script_paths: list[Path] = []

        class _FakePopen:
            def __init__(self, argv, **kw):
                script_paths.append(Path(argv[2]))

        monkeypatch.setattr(updater_github.subprocess, "Popen", _FakePopen)
        staged = tmp_path / "staged"
        install = tmp_path / "install"
        staged.mkdir()
        install.mkdir()
        updater_github.apply_update(staged, install)

        assert script_paths
        content = script_paths[0].read_text(encoding="ascii")
        assert 'start ""' in content
        assert "DocuReader.exe" in content
        script_paths[0].unlink(missing_ok=True)

    def test_apply_update_script_waits_for_process_exit(self, tmp_path: Path, monkeypatch):
        """The .cmd must contain the tasklist wait-loop so it doesn't race the app."""
        monkeypatch.setattr(updater_github.os, "name", "nt")
        script_paths: list[Path] = []

        class _FakePopen:
            def __init__(self, argv, **kw):
                script_paths.append(Path(argv[2]))

        monkeypatch.setattr(updater_github.subprocess, "Popen", _FakePopen)
        updater_github.apply_update(tmp_path / "s", tmp_path / "i")

        content = script_paths[0].read_text(encoding="ascii")
        assert "tasklist" in content.lower()
        assert "waitloop" in content.lower()
        script_paths[0].unlink(missing_ok=True)

    def test_apply_update_returns_1_on_write_script_exception(self, tmp_path: Path, monkeypatch):
        """Any exception from write_apply_script must be caught and return 1."""
        monkeypatch.setattr(updater_github.os, "name", "nt")

        def _bad_write(*args):
            raise RuntimeError("disk full")

        monkeypatch.setattr(updater_github, "write_apply_script", _bad_write)
        rc = updater_github.apply_update(tmp_path / "s", tmp_path / "i")
        assert rc == 1

    def test_apply_update_returns_0_when_popen_succeeds(self, tmp_path: Path, monkeypatch):
        monkeypatch.setattr(updater_github.os, "name", "nt")
        monkeypatch.setattr(updater_github.subprocess, "Popen", lambda *a, **k: None)
        rc = updater_github.apply_update(tmp_path / "staged", tmp_path / "install")
        assert rc == 0

    def test_apply_update_popen_receives_cmd_path(self, tmp_path: Path, monkeypatch):
        """The .cmd file path handed to cmd.exe /c must have a .cmd extension."""
        monkeypatch.setattr(updater_github.os, "name", "nt")
        received: list = []

        class _FakePopen:
            def __init__(self, argv, **kw):
                received.append(argv)

        monkeypatch.setattr(updater_github.subprocess, "Popen", _FakePopen)
        updater_github.apply_update(tmp_path / "s", tmp_path / "i")
        assert received[0][2].endswith(".cmd")


# ---------------------------------------------------------------------------
# 5. _get_short_path — ASCII-safety helper
# ---------------------------------------------------------------------------

class TestGetShortPath:

    def test_returns_string(self, tmp_path: Path):
        result = updater_github._get_short_path(tmp_path)
        assert isinstance(result, str)

    def test_result_is_ascii_for_ascii_path(self, tmp_path: Path):
        result = updater_github._get_short_path(tmp_path)
        result.encode("ascii")  # raises if non-ASCII

    def test_fallback_on_non_windows(self, tmp_path: Path, monkeypatch):
        monkeypatch.setattr(updater_github.os, "name", "posix")
        result = updater_github._get_short_path(tmp_path)
        assert result == str(tmp_path)

    def test_fallback_on_ctypes_failure(self, tmp_path: Path, monkeypatch):
        """If GetShortPathNameW raises, the original path is returned unchanged."""
        import ctypes as _ctypes
        original_windll = getattr(_ctypes, "windll", None)

        class _BadWindll:
            class kernel32:
                @staticmethod
                def GetShortPathNameW(*a):
                    raise OSError("simulated ctypes failure")

        monkeypatch.setattr(_ctypes, "windll", _BadWindll(), raising=False)
        result = updater_github._get_short_path(tmp_path)
        assert result == str(tmp_path)
        if original_windll is not None:
            monkeypatch.setattr(_ctypes, "windll", original_windll)


# ---------------------------------------------------------------------------
# 6. End-to-end update flow simulation
# ---------------------------------------------------------------------------

class TestEndToEndUpdateFlow:
    """Stitch together fetch → stage → apply to simulate the full user journey.

    No network calls are made; _http_json and _http_download are replaced with
    in-process fakes.  The subprocess.Popen call is also intercepted so we
    don't actually run cmd.exe.
    """

    def _setup_http_mocks(self, monkeypatch, zip_bytes: bytes, new_tag: str = "v99.0.0") -> None:
        digest = hashlib.sha256(zip_bytes).hexdigest()
        asset_name = f"DocuReader-{new_tag.lstrip('v')}-portable.zip"

        api_response = {
            "tag_name": new_tag,
            "name": f"DocuReader {new_tag}",
            "prerelease": False,
            "assets": [
                {"name": asset_name, "browser_download_url": f"https://example/{asset_name}"},
                {"name": "SHA256SUMS.txt", "browser_download_url": "https://example/SHA256SUMS.txt"},
            ],
        }
        monkeypatch.setattr(updater_github, "_http_json", lambda url, timeout=30: api_response)

        def _dl(url, dest, timeout=120, progress=None):
            if url.endswith(".zip"):
                Path(dest).write_bytes(zip_bytes)
                if progress:
                    progress(len(zip_bytes), len(zip_bytes))
            else:
                Path(dest).write_text(f"{digest}  {asset_name}\n", encoding="utf-8")

        monkeypatch.setattr(updater_github, "_http_download", _dl)

    def test_full_flow_newer_release_stages_and_applies(self, tmp_path: Path, monkeypatch):
        """Happy path: new release detected → staged → apply script spawned."""
        zip_bytes = _make_release_zip()
        self._setup_http_mocks(monkeypatch, zip_bytes)
        monkeypatch.setattr(updater_github, "staging_root", lambda: tmp_path / "staged")
        monkeypatch.setattr(updater_github.os, "name", "nt")

        popen_calls: list = []

        class _FakePopen:
            def __init__(self, argv, **kw):
                popen_calls.append(argv)

        monkeypatch.setattr(updater_github.subprocess, "Popen", _FakePopen)
        monkeypatch.setattr(
            updater_github.sys, "executable",
            str(tmp_path / "DocuReader.exe"),
            raising=False,
        )
        monkeypatch.setattr(updater_github.sys, "frozen", True, raising=False)

        # Step 1: fetch
        release = updater_github.fetch_release()
        assert release is not None
        assert updater_github.is_newer(release.tag, updater_github.CURRENT_VERSION)

        # Step 2: stage
        staged = updater_github.stage_release(release)
        assert staged is not None
        assert (staged / "DocuReader.exe").exists()

        # Step 3: apply
        install_dir = updater_github.install_dir_for_running_exe()
        rc = updater_github.apply_update(staged, install_dir)
        assert rc == 0

        # Step 4: confirm restart script was spawned
        assert popen_calls, "Popen was not called — restart script was never spawned"
        assert popen_calls[0][0] == "cmd.exe"
        assert popen_calls[0][1] == "/c"
        script_file = Path(popen_calls[0][2])
        assert script_file.suffix == ".cmd"

    def test_full_flow_already_up_to_date_does_not_stage(self, tmp_path: Path, monkeypatch):
        """If the remote tag equals the current version, is_newer returns False."""
        current = updater_github.CURRENT_VERSION
        same_tag = f"v{current}"
        zip_bytes = _make_release_zip()
        self._setup_http_mocks(monkeypatch, zip_bytes, new_tag=same_tag)

        release = updater_github.fetch_release()
        assert release is not None
        assert not updater_github.is_newer(release.tag, updater_github.CURRENT_VERSION)

    def test_full_flow_checksum_mismatch_prevents_apply(self, tmp_path: Path, monkeypatch):
        """If the ZIP's checksum does not match, stage_release returns None and apply is never called."""
        zip_bytes = _make_release_zip()
        # Inject a wrong digest so verification fails.
        monkeypatch.setattr(
            updater_github, "_http_json",
            lambda url, timeout=30: {
                "tag_name": "v99.0.0",
                "name": "DocuReader v99.0.0",
                "prerelease": False,
                "assets": [
                    {"name": "DocuReader-99.0.0-portable.zip", "browser_download_url": "https://example/bad.zip"},
                    {"name": "SHA256SUMS.txt", "browser_download_url": "https://example/SHA256SUMS.txt"},
                ],
            },
        )

        def _dl(url, dest, timeout=120, progress=None):
            if url.endswith(".zip"):
                Path(dest).write_bytes(zip_bytes)
            else:
                Path(dest).write_text("0" * 64 + "  DocuReader-99.0.0-portable.zip\n", encoding="utf-8")

        monkeypatch.setattr(updater_github, "_http_download", _dl)
        monkeypatch.setattr(updater_github, "staging_root", lambda: tmp_path / "staged")
        monkeypatch.setattr(updater_github.os, "name", "nt")

        popen_calls: list = []
        monkeypatch.setattr(
            updater_github.subprocess, "Popen",
            lambda *a, **k: popen_calls.append(a),
        )

        release = updater_github.fetch_release()
        staged = updater_github.stage_release(release)
        assert staged is None, "stage_release must return None on checksum failure"
        assert not popen_calls, "Popen must never be called when staging fails"

    def test_full_flow_network_failure_propagates_gracefully(self, monkeypatch):
        """A network error during fetch_release must not raise an unhandled exception."""
        import urllib.error
        monkeypatch.setattr(updater_github, "_http_json", lambda url, timeout=30: (_ for _ in ()).throw(urllib.error.URLError("down")))
        release = updater_github.fetch_release()
        assert release is None

    def test_full_flow_progress_reaches_100_percent(self, tmp_path: Path, monkeypatch):
        """The progress callback must receive a 100% call during a successful download."""
        zip_bytes = _make_release_zip()
        self._setup_http_mocks(monkeypatch, zip_bytes)
        monkeypatch.setattr(updater_github, "staging_root", lambda: tmp_path / "staged")

        progress_calls: list[tuple[int, int]] = []

        release = updater_github.fetch_release()
        updater_github.stage_release(
            release,
            progress=lambda d, t: progress_calls.append((d, t)),
        )

        assert progress_calls, "progress callback was never called"
        last_downloaded, last_total = progress_calls[-1]
        assert last_downloaded == last_total, "final progress call must report downloaded == total"
