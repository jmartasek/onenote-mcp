"""Tests for hierarchy filter helpers and tree-building logic in onenote_mcp."""

import json
import pytest

from onenote_lib.xml_parser import (
    NotebookInfo,
    PageInfo,
    SectionGroupInfo,
    SectionInfo,
)
from onenote_mcp import (
    _build_filter,
    _is_excluded,
    _make_path,
    _node_matches,
    _beyond_depth,
    _ts_in_range,
    _notebook_to_tree,
    _section_group_to_tree,
    _section_to_tree_item,
    _parse_iso,
    _flatten_sections,
    _find_section_pages,
)


# ── Fixtures ─────────────────────────────────────────────────────────

def _make_notebook() -> NotebookInfo:
    """Build a sample notebook hierarchy for testing.

    Structure:
        OneWork/
        ├── General (section) — pages: Intro, Setup
        ├── DataOps (section group)
        │   ├── Team (section) — pages: Standup 2026-04, Sprint Review
        │   └── Archive (section) — pages: Old Notes
        └── Archive-2024 (section group)
            └── Stuff (section) — pages: Legacy
    """
    general = SectionInfo(
        id="sec-general", name="General",
        pages=[
            PageInfo(id="p1", name="Intro", last_modified="2026-04-20T10:00:00Z", level=1),
            PageInfo(id="p2", name="Setup", last_modified="2026-04-15T08:00:00Z", level=1),
        ],
    )
    team = SectionInfo(
        id="sec-team", name="Team",
        pages=[
            PageInfo(id="p3", name="Standup 2026-04", last_modified="2026-04-24T09:00:00Z", level=1),
            PageInfo(id="p4", name="Sprint Review", last_modified="2026-04-10T14:00:00Z", level=1),
        ],
    )
    archive_sec = SectionInfo(
        id="sec-archive", name="Archive",
        pages=[
            PageInfo(id="p5", name="Old Notes", last_modified="2025-01-01T00:00:00Z", level=1),
        ],
    )
    dataops = SectionGroupInfo(
        id="sg-dataops", name="DataOps",
        sections=[team, archive_sec],
    )
    archive_2024 = SectionGroupInfo(
        id="sg-archive", name="Archive-2024",
        sections=[SectionInfo(
            id="sec-stuff", name="Stuff",
            pages=[PageInfo(id="p6", name="Legacy", last_modified="2024-06-01T00:00:00Z", level=1)],
        )],
    )
    return NotebookInfo(
        id="nb1", name="OneWork",
        sections=[general],
        section_groups=[dataops, archive_2024],
    )


# ── _build_filter tests ─────────────────────────────────────────────


class TestBuildFilter:
    def test_no_filters_returns_none(self):
        assert _build_filter() is None

    def test_name_regex_compiles(self):
        f = _build_filter(name_regex="Standup")
        assert f is not None
        assert f.name_re.search("Standup 2026-04")

    def test_invalid_regex_raises(self):
        with pytest.raises(ValueError, match="Invalid regex"):
            _build_filter(name_regex="[invalid")

    def test_invalid_timestamp_raises(self):
        with pytest.raises(ValueError, match="Invalid ISO timestamp"):
            _build_filter(modified_since="not-a-date")

    def test_timestamp_parsing_utc_default(self):
        f = _build_filter(modified_since="2026-04-01")
        assert f.modified_since.tzname() == "UTC"

    def test_timestamp_preserves_timezone(self):
        f = _build_filter(modified_since="2026-04-01T10:00:00+02:00")
        assert f.modified_since.utcoffset().total_seconds() == 7200

    def test_max_depth(self):
        f = _build_filter(max_depth=2)
        assert f.max_depth == 2

    def test_include_pages_false(self):
        f = _build_filter(include_pages=False)
        assert f.include_pages is False


# ── Helper function tests ────────────────────────────────────────────


class TestHelpers:
    def test_make_path_root(self):
        assert _make_path("", "OneWork") == "OneWork"

    def test_make_path_child(self):
        assert _make_path("OneWork", "DataOps") == "OneWork/DataOps"

    def test_is_excluded_no_filter(self):
        assert _is_excluded("any/path", None) is False

    def test_is_excluded_match(self):
        f = _build_filter(path_exclude_regex="Archive")
        assert _is_excluded("OneWork/Archive-2024", f) is True

    def test_is_excluded_no_match(self):
        f = _build_filter(path_exclude_regex="Archive")
        assert _is_excluded("OneWork/DataOps", f) is False

    def test_node_matches_no_filter(self):
        assert _node_matches("any", "any", None) is True

    def test_node_matches_name_hit(self):
        f = _build_filter(name_regex="Team")
        assert _node_matches("OneWork/DataOps/Team", "Team", f) is True

    def test_node_matches_name_miss(self):
        f = _build_filter(name_regex="Team")
        assert _node_matches("OneWork/DataOps/Archive", "Archive", f) is False

    def test_node_matches_path_hit(self):
        f = _build_filter(path_regex="DataOps")
        assert _node_matches("OneWork/DataOps/Team", "Team", f) is True

    def test_node_matches_path_miss(self):
        f = _build_filter(path_regex="DataOps")
        assert _node_matches("OneWork/General", "General", f) is False

    def test_beyond_depth_unlimited(self):
        assert _beyond_depth(100, None) is False
        f = _build_filter(name_regex="x")  # need some filter active
        f.max_depth = -1
        assert _beyond_depth(100, f) is False

    def test_beyond_depth_at_limit(self):
        f = _build_filter(max_depth=2)
        assert _beyond_depth(2, f) is False
        assert _beyond_depth(3, f) is True

    def test_ts_in_range_no_filter(self):
        assert _ts_in_range("2026-01-01T00:00:00Z", None) is True

    def test_ts_in_range_since(self):
        f = _build_filter(modified_since="2026-04-01")
        assert _ts_in_range("2026-04-20T10:00:00Z", f) is True
        assert _ts_in_range("2026-03-31T23:59:59Z", f) is False

    def test_ts_in_range_before(self):
        f = _build_filter(modified_before="2026-04-01")
        assert _ts_in_range("2026-03-15T10:00:00Z", f) is True
        assert _ts_in_range("2026-04-02T00:00:00Z", f) is False

    def test_ts_in_range_no_timestamp(self):
        f = _build_filter(modified_since="2026-04-01")
        assert _ts_in_range(None, f) is False


# ── Tree filter tests ────────────────────────────────────────────────


class TestTreeFiltering:
    def test_no_filter_returns_full_tree(self):
        nb = _make_notebook()
        tree = _notebook_to_tree(nb)
        assert tree is not None
        assert tree["name"] == "OneWork"
        assert len(tree["sections"]) == 1
        assert len(tree["section_groups"]) == 2
        # DataOps has Team + Archive sections
        dops = tree["section_groups"][0]
        assert dops["name"] == "DataOps"
        assert len(dops["sections"]) == 2

    def test_path_exclude_prunes_subtree(self):
        nb = _make_notebook()
        filt = _build_filter(path_exclude_regex="Archive")
        tree = _notebook_to_tree(nb, filt=filt)
        assert tree is not None
        # Archive-2024 group should be gone
        group_names = [g["name"] for g in tree["section_groups"]]
        assert "Archive-2024" not in group_names
        # Archive section inside DataOps should also be gone
        dops = tree["section_groups"][0]
        sec_names = [s["name"] for s in dops["sections"]]
        assert "Archive" not in sec_names
        # Team should survive
        assert "Team" in sec_names

    def test_name_regex_filters_to_matching(self):
        nb = _make_notebook()
        filt = _build_filter(name_regex="Team")
        tree = _notebook_to_tree(nb, filt=filt)
        assert tree is not None
        # Only DataOps group survives (contains Team)
        assert len(tree["section_groups"]) == 1
        assert tree["section_groups"][0]["name"] == "DataOps"
        # Within DataOps, only Team section
        assert len(tree["section_groups"][0]["sections"]) == 1
        assert tree["section_groups"][0]["sections"][0]["name"] == "Team"

    def test_path_regex_keeps_ancestors(self):
        nb = _make_notebook()
        filt = _build_filter(path_regex="DataOps/Team")
        tree = _notebook_to_tree(nb, filt=filt)
        assert tree is not None
        # Notebook kept as ancestor, DataOps kept, Team section present
        assert len(tree["section_groups"]) == 1
        dops = tree["section_groups"][0]
        assert dops["name"] == "DataOps"
        sec_names = [s["name"] for s in dops["sections"]]
        assert "Team" in sec_names

    def test_modified_since_filters_pages(self):
        nb = _make_notebook()
        filt = _build_filter(modified_since="2026-04-18")
        tree = _notebook_to_tree(nb, filt=filt)
        assert tree is not None
        # General section: Intro (Apr 20) passes, Setup (Apr 15) filtered
        gen = tree["sections"][0]
        page_names = [p["name"] for p in gen["pages"]]
        assert "Intro" in page_names
        assert "Setup" not in page_names
        # Team: Standup (Apr 24) passes, Sprint Review (Apr 10) filtered
        dops = tree["section_groups"][0]
        team = [s for s in dops["sections"] if s["name"] == "Team"][0]
        page_names = [p["name"] for p in team["pages"]]
        assert "Standup 2026-04" in page_names
        assert "Sprint Review" not in page_names

    def test_max_depth_0_notebooks_only(self):
        nb = _make_notebook()
        filt = _build_filter(max_depth=0)
        tree = _notebook_to_tree(nb, filt=filt)
        assert tree is not None
        assert tree["name"] == "OneWork"
        assert tree["sections"] == []
        assert tree["section_groups"] == []

    def test_max_depth_1_no_pages(self):
        nb = _make_notebook()
        filt = _build_filter(max_depth=1)
        tree = _notebook_to_tree(nb, filt=filt)
        assert tree is not None
        # Sections and groups present at depth 1
        assert len(tree["sections"]) == 1
        assert len(tree["section_groups"]) == 2
        # But sections have no pages (pages are at depth 2+)
        assert tree["sections"][0]["pages"] == []
        # Section groups have no child sections (those are at depth 2)
        for sg in tree["section_groups"]:
            assert sg["sections"] == []
            assert sg["section_groups"] == []

    def test_include_pages_false(self):
        nb = _make_notebook()
        filt = _build_filter(include_pages=False)
        tree = _notebook_to_tree(nb, filt=filt)
        assert tree is not None
        # Structure present but all page lists empty
        assert tree["sections"][0]["pages"] == []
        dops = tree["section_groups"][0]
        for sec in dops["sections"]:
            assert sec["pages"] == []

    def test_ancestor_match_includes_full_subtree(self):
        """When a section group directly matches name_regex, all its
        children (sections + pages) are included without further filtering."""
        nb = _make_notebook()
        filt = _build_filter(name_regex="DataOps")
        tree = _notebook_to_tree(nb, filt=filt)
        assert tree is not None
        dops = tree["section_groups"][0]
        # Both Team and Archive sections included (ancestor matched)
        sec_names = {s["name"] for s in dops["sections"]}
        assert sec_names == {"Team", "Archive"}
        # All pages in Team included
        team = [s for s in dops["sections"] if s["name"] == "Team"][0]
        assert len(team["pages"]) == 2

    def test_exclude_wins_over_ancestor_match(self):
        """path_exclude_regex prunes even when ancestor matched."""
        nb = _make_notebook()
        filt = _build_filter(name_regex="DataOps", path_exclude_regex="Archive")
        tree = _notebook_to_tree(nb, filt=filt)
        assert tree is not None
        dops = tree["section_groups"][0]
        sec_names = [s["name"] for s in dops["sections"]]
        assert "Archive" not in sec_names
        assert "Team" in sec_names

    def test_combined_exclude_and_date(self):
        nb = _make_notebook()
        filt = _build_filter(
            path_exclude_regex="Archive",
            modified_since="2026-04-20",
        )
        tree = _notebook_to_tree(nb, filt=filt)
        assert tree is not None
        # General: only Intro (Apr 20) passes date filter
        gen = tree["sections"][0]
        assert len(gen["pages"]) == 1
        assert gen["pages"][0]["name"] == "Intro"
        # DataOps/Team: only Standup (Apr 24) passes
        dops = tree["section_groups"][0]
        team = dops["sections"][0]
        assert len(team["pages"]) == 1
        assert team["pages"][0]["name"] == "Standup 2026-04"

    def test_no_matches_returns_none(self):
        nb = _make_notebook()
        filt = _build_filter(name_regex="NonExistent")
        tree = _notebook_to_tree(nb, filt=filt)
        assert tree is None

    def test_notebook_excluded_returns_none(self):
        nb = _make_notebook()
        filt = _build_filter(path_exclude_regex="OneWork")
        tree = _notebook_to_tree(nb, filt=filt)
        assert tree is None


# ── Flat list filter tests ───────────────────────────────────────────


class TestFlatListFiltering:
    def test_find_section_pages_no_filter(self):
        nb = _make_notebook()
        pages = _find_section_pages(nb.section_groups, "sec-team")
        assert pages is not None
        assert len(pages) == 2

    def test_find_section_pages_name_filter(self):
        nb = _make_notebook()
        filt = _build_filter(name_regex="Standup")
        pages = _find_section_pages(nb.section_groups, "sec-team", filt)
        assert len(pages) == 1
        assert pages[0]["name"] == "Standup 2026-04"

    def test_find_section_pages_date_filter(self):
        nb = _make_notebook()
        filt = _build_filter(modified_since="2026-04-15")
        pages = _find_section_pages(nb.section_groups, "sec-team", filt)
        assert len(pages) == 1
        assert pages[0]["name"] == "Standup 2026-04"

    def test_find_section_pages_not_found(self):
        nb = _make_notebook()
        pages = _find_section_pages(nb.section_groups, "nonexistent")
        assert pages is None
