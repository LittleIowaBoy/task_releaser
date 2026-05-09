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
    assert flags & subprocess.DETACHED_PROCESS


def test_apply_update_returns_1_on_popen_oserror(tmp_path: Path, monkeypatch):
    monkeypatch.setattr(updater_github.os, "name", "nt")

    def _bad_popen(*args, **kwargs):
        raise OSError("no cmd.exe found")

    monkeypatch.setattr(updater_github.subprocess, "Popen", _bad_popen)
    rc = updater_github.apply_update(tmp_path / "staged", tmp_path / "install")
    assert rc == 1


def test_apply_update_script_contains_correct_paths(tmp_path: Path, monkeypatch):
    """The .cmd file spawned by apply_update must contain the staged+install paths."""
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
    assert str(staged) in content
    assert str(install) in content
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
