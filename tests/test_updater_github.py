import hashlib
import io
import json
import zipfile
from pathlib import Path

import pytest

import updater_github


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


def test_select_asset_picks_named_zip_and_sums():
    assets = [
        {"name": "DocuReader-0.3.0-portable.zip", "browser_download_url": "u1"},
        {"name": "SHA256SUMS.txt", "browser_download_url": "u2"},
        {"name": "noise.txt", "browser_download_url": "u3"},
    ]
    z, s = updater_github._select_asset(assets, "v0.3.0")
    assert z["browser_download_url"] == "u1"
    assert s["browser_download_url"] == "u2"


def test_stage_release_with_monkeypatched_http(tmp_path: Path, monkeypatch):
    # Build a tiny zip in-memory that we'll "download".
    inner = io.BytesIO()
    with zipfile.ZipFile(inner, "w") as zf:
        zf.writestr("DocuReader/marker.txt", "ok")
    zip_bytes = inner.getvalue()
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
