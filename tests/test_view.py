import pandas as pd
import pytest

from templates import Template, HighlightRule
from view import apply_template


def test_drop_reorder_sort():
    df = pd.DataFrame({
        "Item": ["b", "a", "c"],
        "Drop me": [1, 2, 3],
        "Location": ["L-100", "L-50", "L-200"],
    })
    t = Template(
        name="t",
        drop=["Drop me"],
        order=["Location", "Item"],
        sort_by=[["Location", "asc"]],
        location_columns=["Location"],
    )
    out, meta = apply_template(df, t)
    assert list(out.columns) == ["Location", "Item"]
    assert out["Location"].tolist() == ["L-50", "L-100", "L-200"]
    assert meta.location_column == "Location"


def test_highlight_numeric_threshold():
    df = pd.DataFrame({"OHB": [1, 6, 10]})
    t = Template(
        name="t",
        highlights=[HighlightRule(name="low", when="OHB <= 5", target_columns=["OHB"], color="darkgreen")],
    )
    out, meta = apply_template(df, t)
    rows = sorted(r for r, c, _ in meta.highlights if c == "OHB")
    assert rows == [0]


def test_highlight_two_columns_compare():
    df = pd.DataFrame({
        "Last Replen": ["2026-01-02", "2026-01-01", None],
        "Short Time":  ["2026-01-01", "2026-01-02", "2026-01-01"],
    })
    t = Template(
        name="t",
        highlights=[
            HighlightRule(name="g", when="Last Replen > Short Time", target_columns=["Last Replen"], color="darkgreen"),
            HighlightRule(name="y", when="Last Replen < Short Time", target_columns=["Last Replen"], color="darkyellow"),
        ],
    )
    out, meta = apply_template(df, t)
    by_row = {r: color for (r, _c, color) in meta.highlights}
    assert by_row.get(0) == "darkgreen"
    assert by_row.get(1) == "darkyellow"
    assert 2 not in by_row  # NaN side -> no highlight


def test_passthrough_on_empty_template():
    df = pd.DataFrame({"a": [1, 2], "b": [3, 4]})
    t = Template(name="t")
    out, meta = apply_template(df, t)
    assert list(out.columns) == ["a", "b"]
    assert meta.highlights == []
