import pandas as pd
from templates import (
    Template,
    TemplateRegistry,
    PASSTHROUGH_TEMPLATE,
    HighlightRule,
)


def _registry() -> TemplateRegistry:
    return TemplateRegistry(
        templates=[
            Template(
                name="Replen",
                filename_patterns=["*replen*"],
                required_columns=["Location", "Item", "Current OHB"],
            ),
            Template(
                name="Chase",
                required_columns=["Task ID", "Active OHB", "Allocated", "Item"],
            ),
            Template(
                name="Locked",
                required_columns=["TASK_ID", "Aisle"],
            ),
        ]
    )


def test_select_by_filename_pattern():
    reg = _registry()
    cols = ["Location", "Item", "Current OHB", "Extra"]
    match = reg.select(cols, filename="weekly_replen_audit.csv")
    assert match.template.name == "Replen"
    assert "filename" in match.reason


def test_select_by_required_columns_longest_match():
    reg = _registry()
    # Both Replen and Chase could-not match these; only Chase fits.
    cols = ["Task ID", "Active OHB", "Allocated", "Item", "Aisle", "TASK_ID"]
    match = reg.select(cols, filename="export.xlsx")
    # Chase has 4 required cols, Locked has 2 -> Chase wins.
    assert match.template.name == "Chase"


def test_select_falls_back_to_passthrough():
    reg = _registry()
    match = reg.select(["random", "columns", "nothing matches"], filename="x.csv")
    assert match.template is PASSTHROUGH_TEMPLATE


def test_round_trip_to_dict():
    t = Template(
        name="t1",
        required_columns=["a", "b"],
        drop=["c"],
        order=["a", "b"],
        sort_by=[["a", "asc"]],
        location_columns=["a"],
        highlights=[HighlightRule(name="r", when="a > b", target_columns=["a"], color="red", priority=1)],
    )
    restored = Template.from_dict(t.to_dict())
    assert restored.name == "t1"
    assert restored.drop == ["c"]
    assert restored.highlights[0].when == "a > b"
    assert restored.highlights[0].priority == 1
