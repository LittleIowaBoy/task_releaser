"""Apply a :class:`templates.Template` to a parsed DataFrame.

This module is intentionally GUI-free so it can be unit-tested without PyQt.
It returns the transformed DataFrame and a separate ``ViewMeta`` object that
the GUI consumes to colour cells and group rows.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd

from templates import HighlightRule, Template, parse_condition, is_literal


LOCATION_PATTERN = re.compile(r'^\s*([A-Za-z-]*?)(\d+)?([A-Za-z]*)\s*$')


def parse_location_parts(value: Any) -> Tuple[str, float, str]:
    """Parse a location string into (prefix, numeric body, suffix) for natural sort."""
    text = "" if pd.isna(value) else str(value).strip().upper()
    match = LOCATION_PATTERN.match(text)
    if not match:
        return (text, float("inf"), "")

    prefix, number_part, suffix = match.groups()
    if number_part:
        if suffix and len(number_part) < 7:
            number = int(number_part.ljust(7, "0"))
        else:
            number = int(number_part)
    else:
        number = float("inf")
    return (prefix or "", number, suffix or "")


# A single cell colour assignment. ``color`` matches the names used in
# ``HighlightRule.color`` (darkgreen, darkyellow, red, blue).
HighlightTuple = Tuple[int, str, str]  # (row_index, column_name, color)


@dataclass
class ViewMeta:
    """Side-band data the GUI uses when rendering the transformed table."""

    template_name: str = ""
    match_reason: str = ""
    location_column: Optional[str] = None
    highlights: List[HighlightTuple] = field(default_factory=list)


def apply_template(
    df: pd.DataFrame,
    template: Template,
) -> Tuple[pd.DataFrame, ViewMeta]:
    """Apply ``template`` to ``df`` and return the transformed frame + meta.

    Steps (in order):
    1. Drop columns listed in ``template.drop``.
    2. Rename per ``template.rename``.
    3. Reorder per ``template.order``. Unknown columns are appended.
    4. Sort: location-aware natural sort if the first sort key is a known
       location column; otherwise a normal pandas sort.
    5. Evaluate ``template.highlights`` to produce ``ViewMeta.highlights``.
    """

    meta = ViewMeta(template_name=template.name)

    if df is None or df.empty:
        return df, meta

    out = df.copy()

    # 1. Drop
    if template.drop:
        out = out.drop(columns=[c for c in template.drop if c in out.columns])

    # 2. Rename
    if template.rename:
        out = out.rename(columns={k: v for k, v in template.rename.items() if k in out.columns})

    # 3. Reorder
    if template.order:
        present_in_order = [c for c in template.order if c in out.columns]
        remaining = [c for c in out.columns if c not in present_in_order]
        out = out[present_in_order + remaining]

    # 4. Sort
    location_col = _find_location_column(out.columns, template.location_columns)
    meta.location_column = location_col
    out = _apply_sort(out, template.sort_by, location_col)
    out = out.reset_index(drop=True)

    # 5. Highlights
    meta.highlights = _evaluate_highlights(out, template.highlights)

    return out, meta


def _find_location_column(columns, candidates: List[str]) -> Optional[str]:
    for col in candidates:
        if col in columns:
            return col
    return None


def _apply_sort(
    df: pd.DataFrame,
    sort_by: List[List[Any]],
    location_col: Optional[str],
) -> pd.DataFrame:
    if not sort_by:
        return df

    # Filter to keys that exist in the frame; preserve order/direction.
    valid: List[Tuple[str, bool]] = []
    for entry in sort_by:
        if not entry:
            continue
        col = entry[0]
        direction = entry[1] if len(entry) > 1 else "asc"
        if col in df.columns:
            valid.append((col, str(direction).lower() != "desc"))

    if not valid:
        return df

    cols = [c for c, _ in valid]
    asc = [a for _, a in valid]

    try:
        # Location-aware natural sort applies only when the *primary* sort key
        # is the detected location column.
        if location_col and cols[0] == location_col:
            return df.sort_values(
                by=cols,
                ascending=asc,
                key=lambda s: s.map(parse_location_parts) if s.name == location_col else s,
            )
        return df.sort_values(by=cols, ascending=asc)
    except (KeyError, TypeError, ValueError):
        return df


# ---------------------------------------------------------------------------
# Highlight evaluation
# ---------------------------------------------------------------------------


def _evaluate_highlights(
    df: pd.DataFrame,
    rules: List[HighlightRule],
) -> List[HighlightTuple]:
    """Compute ``(row, column, color)`` tuples by applying each rule."""

    if df.empty or not rules:
        return []

    # Track best (highest priority) colour per (row, col); higher priority wins.
    best: Dict[Tuple[int, str], Tuple[int, str]] = {}

    sorted_rules = sorted(rules, key=lambda r: r.priority)
    for rule in sorted_rules:
        mask = _evaluate_condition(df, rule.when)
        if mask is None or not mask.any():
            continue
        cols = [c for c in rule.target_columns if c in df.columns]
        if not cols:
            continue
        for row_idx in df.index[mask]:
            for col in cols:
                key = (int(row_idx), col)
                prev = best.get(key)
                if prev is None or prev[0] <= rule.priority:
                    best[key] = (rule.priority, rule.color)

    return [(row, col, color) for (row, col), (_, color) in best.items()]


def _evaluate_condition(df: pd.DataFrame, expr: str) -> Optional[pd.Series]:
    """Evaluate a simple ``"left OP right"`` expression against ``df``.

    ``left`` must be a column name. ``right`` may be a column name, a numeric
    literal, or a quoted string literal. Datetime comparison is attempted
    first; on failure falls back to numeric, then string.
    """

    parsed = parse_condition(expr)
    if parsed is None:
        return None
    left, op, right = parsed
    if left not in df.columns:
        return None

    left_series = df[left]
    if right in df.columns:
        right_value: Any = df[right]
        right_is_series = True
    elif is_literal(right):
        right_value = _parse_literal(right)
        right_is_series = False
    else:
        return None

    # Try datetime first (only meaningful when right is also a series), then numeric, then raw.
    casters = []
    if right_is_series:
        casters.append("datetime")
    casters.extend(["numeric", "raw"])

    for kind in casters:
        try:
            l_cast = _cast(left_series, kind)
            r_cast = _cast(right_value, kind) if right_is_series else _cast_scalar(right_value, kind)
        except (TypeError, ValueError):
            continue
        try:
            mask = _compare(l_cast, r_cast, op)
        except (TypeError, ValueError):
            continue
        if mask is None:
            continue
        return mask.fillna(False) if hasattr(mask, "fillna") else mask
    return None


def _cast(series: pd.Series, kind: str) -> pd.Series:
    if kind == "datetime":
        result = pd.to_datetime(series, errors="coerce")
        if result.isna().all():
            raise TypeError("not datetime")
        return result
    if kind == "numeric":
        result = pd.to_numeric(series, errors="coerce")
        if result.isna().all():
            raise TypeError("not numeric")
        return result
    return series


def _cast_scalar(value: Any, kind: str) -> Any:
    if kind == "numeric":
        return float(value)
    return value


def _compare(left, right, op: str):
    if op == ">":
        return left > right
    if op == "<":
        return left < right
    if op == ">=":
        return left >= right
    if op == "<=":
        return left <= right
    if op == "==":
        return left == right
    if op == "!=":
        return left != right
    return None


def _parse_literal(token: str) -> Any:
    if token and token[0] in ("'", '"') and token[-1] == token[0]:
        return token[1:-1]
    try:
        return int(token)
    except ValueError:
        return float(token)
