"""Release-channel constants for the GitHub-based auto-updater.

Committed to the repo so that the frozen executable knows where to look for
updates without depending on a local git checkout. Override at runtime via
the environment variables ``DOCUREADER_GITHUB_OWNER`` and
``DOCUREADER_GITHUB_REPO`` (useful when forking).
"""

from __future__ import annotations

import os

GITHUB_OWNER: str = os.environ.get("DOCUREADER_GITHUB_OWNER", "LittleIowaBoy")
GITHUB_REPO: str = os.environ.get("DOCUREADER_GITHUB_REPO", "task_releaser")

# Asset filename pattern attached to GitHub Releases by the
# .github/workflows/release.yml workflow. ``{version}`` is replaced at runtime.
ASSET_NAME_PATTERN: str = "DocuReader-{version}-portable.zip"

# Optional checksum manifest published alongside the asset.
CHECKSUMS_NAME: str = "SHA256SUMS.txt"
