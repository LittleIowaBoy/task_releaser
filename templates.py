"""Template engine for DocuReader.

A Template describes how to transform a parsed DataFrame for display:
- which columns to drop
- which columns to rename
- the desired output column order
- which columns to sort by (with location-aware natural sort)
- conditional highlight rules

Templates are matched against an incoming file by:
1. Filename glob/regex pattern, then
2. Longest matching `required_columns` subset, then
3. First-defined order, then
4. The built-in passthrough fallback.

Storage:
- Bundled defaults: ``default_templates.json`` next to this module.
- User overrides: ``~/.docureader/templates.json`` (created on first launch
  by copying the defaults; subsequent app upgrades merge by template name
  without clobbering user edits).
"""

from __future__ import annotations

import fnmatch
import json
import re
import shutil
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any, Iterable, List, Optional, Tuple


USER_CONFIG_DIR = Path.home() / ".docureader"
USER_TEMPLATES_PATH = USER_CONFIG_DIR / "templates.json"
DEFAULT_TEMPLATES_PATH = Path(__file__).resolve().parent / "default_templates.json"


# ---------------------------------------------------------------------------
# Highlight rules
# ---------------------------------------------------------------------------


@dataclass
class HighlightRule:
    """A conditional formatting rule applied per-row to one or more cells.

    Attributes:
        name: Human-readable rule label.
        when: Free-form condition expression. Supported forms (kept simple
            on purpose so users can author rules in JSON):
              - ``"col_a > col_b"`` (comparison between two columns; both
                parsed as datetime first, then numeric, then string).
              - ``"col_a < col_b"``
              - ``"col_a == col_b"``
              - ``"col <= <number>"`` / ``"col >= <number>"`` etc.
        target_columns: Columns to colour when the row matches.
        color: One of the named colours below (kept short so JSON stays readable).
        priority: Higher priority overrides lower when multiple rules match.
    """

    name: str = "rule"
    when: str = ""
    target_columns: List[str] = field(default_factory=list)
    color: str = "darkyellow"  # darkgreen | darkyellow | red | blue
    priority: int = 0

    @classmethod
    def from_dict(cls, data: dict) -> "HighlightRule":
        return cls(
            name=data.get("name", "rule"),
            when=data.get("when", ""),
            target_columns=list(data.get("target_columns", [])),
            color=data.get("color", "darkyellow"),
            priority=int(data.get("priority", 0)),
        )

    def to_dict(self) -> dict:
        return asdict(self)


# ---------------------------------------------------------------------------
# Template
# ---------------------------------------------------------------------------


@dataclass
class Template:
    """A user- or system-defined column-organisation template.

    Templates are intentionally declarative so they can round-trip through
    JSON. The runtime application logic lives in ``view.apply_template``.
    """

    name: str
    description: str = ""
    filename_patterns: List[str] = field(default_factory=list)
    required_columns: List[str] = field(default_factory=list)
    drop: List[str] = field(default_factory=list)
    rename: dict = field(default_factory=dict)
    order: List[str] = field(default_factory=list)
    sort_by: List[List[Any]] = field(default_factory=list)  # [[col, "asc"|"desc"], ...]
    location_columns: List[str] = field(default_factory=list)
    highlights: List[HighlightRule] = field(default_factory=list)
    builtin: bool = False  # bundled default; user edits create a copy

    @classmethod
    def from_dict(cls, data: dict) -> "Template":
        return cls(
            name=data["name"],
            description=data.get("description", ""),
            filename_patterns=list(data.get("filename_patterns", [])),
            required_columns=list(data.get("required_columns", [])),
            drop=list(data.get("drop", [])),
            rename=dict(data.get("rename", {})),
            order=list(data.get("order", [])),
            sort_by=[list(s) for s in data.get("sort_by", [])],
            location_columns=list(data.get("location_columns", [])),
            highlights=[HighlightRule.from_dict(h) for h in data.get("highlights", [])],
            builtin=bool(data.get("builtin", False)),
        )

    def to_dict(self) -> dict:
        d = asdict(self)
        d["highlights"] = [h.to_dict() if isinstance(h, HighlightRule) else h for h in self.highlights]
        return d


PASSTHROUGH_TEMPLATE = Template(
    name="(Passthrough)",
    description="No transformation - shown when no other template matches.",
    builtin=True,
)


# ---------------------------------------------------------------------------
# Registry
# ---------------------------------------------------------------------------


@dataclass
class MatchResult:
    template: Template
    reason: str  # human-readable explanation for the "matched by" tooltip


class TemplateRegistry:
    """Loads, persists, and selects templates."""

    def __init__(self, templates: Optional[List[Template]] = None) -> None:
        self.templates: List[Template] = list(templates) if templates else []

    # -- Persistence -------------------------------------------------------

    @classmethod
    def load(cls) -> "TemplateRegistry":
        """Load user templates, copying bundled defaults on first launch.

        Behaviour:
        - If ``~/.docureader/templates.json`` is missing, copy the bundled
          defaults there and load.
        - Otherwise, load user file as-is. Any *new* bundled templates whose
          ``name`` is not present in the user file are appended (so a new
          release that adds a category shows up automatically), but existing
          user-edited entries are never overwritten.
        """
        try:
            USER_CONFIG_DIR.mkdir(parents=True, exist_ok=True)
        except OSError:
            # Fall back to bundled defaults in-memory if home dir is read-only.
            return cls(_load_templates_from(DEFAULT_TEMPLATES_PATH))

        if not USER_TEMPLATES_PATH.exists():
            try:
                shutil.copyfile(DEFAULT_TEMPLATES_PATH, USER_TEMPLATES_PATH)
            except OSError:
                return cls(_load_templates_from(DEFAULT_TEMPLATES_PATH))

        user_templates = _load_templates_from(USER_TEMPLATES_PATH)
        bundled = _load_templates_from(DEFAULT_TEMPLATES_PATH)
        existing_names = {t.name for t in user_templates}
        merged_new = [t for t in bundled if t.name not in existing_names]
        if merged_new:
            user_templates.extend(merged_new)
            cls(user_templates).save()
        return cls(user_templates)

    def save(self, path: Optional[Path] = None) -> None:
        """Persist templates to JSON. Defaults to the user config path."""
        target = path or USER_TEMPLATES_PATH
        target.parent.mkdir(parents=True, exist_ok=True)
        payload = {"version": 1, "templates": [t.to_dict() for t in self.templates]}
        target.write_text(json.dumps(payload, indent=2), encoding="utf-8")

    # -- CRUD --------------------------------------------------------------

    def get(self, name: str) -> Optional[Template]:
        for t in self.templates:
            if t.name == name:
                return t
        return None

    def upsert(self, template: Template) -> None:
        for i, existing in enumerate(self.templates):
            if existing.name == template.name:
                self.templates[i] = template
                return
        self.templates.append(template)

    def remove(self, name: str) -> bool:
        for i, t in enumerate(self.templates):
            if t.name == name:
                if t.builtin:
                    return False
                del self.templates[i]
                return True
        return False

    # -- Selection ---------------------------------------------------------

    def select(self, columns: Iterable[str], filename: str = "") -> MatchResult:
        """Pick the best-matching template for an incoming file.

        Selection order (see plan §"Template conflict resolution"):
            1. Filename pattern match (glob via ``fnmatch``).
            2. Longest ``required_columns`` subset match.
            3. First-defined.
            4. Passthrough fallback.
        """
        cols = list(columns)
        col_set = set(cols)
        name_only = Path(filename).name if filename else ""

        # 1. Filename pattern
        if name_only:
            for t in self.templates:
                for pattern in t.filename_patterns:
                    if fnmatch.fnmatch(name_only.lower(), pattern.lower()):
                        # Still require the columns to be at least minimally present
                        if not t.required_columns or set(t.required_columns).issubset(col_set):
                            return MatchResult(t, f"filename pattern '{pattern}'")

        # 2. Longest required_columns subset match
        best: Optional[Template] = None
        best_len = -1
        for t in self.templates:
            req = set(t.required_columns)
            if not req:
                continue
            if req.issubset(col_set) and len(req) > best_len:
                best = t
                best_len = len(req)
        if best is not None:
            return MatchResult(best, f"matched {best_len} required column(s)")

        return MatchResult(PASSTHROUGH_TEMPLATE, "no template matched - passthrough")


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _load_templates_from(path: Path) -> List[Template]:
    if not path.exists():
        return []
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return []
    raw = data.get("templates", []) if isinstance(data, dict) else data
    return [Template.from_dict(t) for t in raw]


# ---------------------------------------------------------------------------
# Highlight evaluation (kept here so `view.py` stays presentation-only)
# ---------------------------------------------------------------------------


_NUM_RE = re.compile(r"^-?\d+(?:\.\d+)?$")
_OP_RE = re.compile(r"\s*(>=|<=|==|!=|>|<)\s*")


def parse_condition(expr: str) -> Optional[Tuple[str, str, str]]:
    """Return ``(left, op, right)`` for a simple binary expression, or None."""
    if not expr:
        return None
    m = _OP_RE.search(expr)
    if not m:
        return None
    op = m.group(1)
    left = expr[: m.start()].strip()
    right = expr[m.end():].strip()
    if not left or not right:
        return None
    return left, op, right


def is_literal(token: str) -> bool:
    return bool(_NUM_RE.match(token)) or (
        len(token) >= 2 and token[0] == token[-1] and token[0] in ("'", '"')
    )
