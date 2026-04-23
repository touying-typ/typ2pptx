"""Tests for the paragraph auto-detection toggle.

Paragraph detection is the heuristic that merges consecutive wrapped text
lines (same font, aligned left edge, full-width) into a single word-wrapped
textbox. It is **disabled by default** because the heuristic often mis-merges
tightly-packed content such as tables, which is a frequent source of bugs.

These tests directly exercise :func:`_merge_paragraph_lines` and the
``ConversionConfig.detect_paragraphs`` flag to pin down the new default
behaviour and the opt-in path.
"""
import pytest

from typ2pptx.core.converter import (
    ConversionConfig,
    _group_segments_by_line,
    _merge_paragraph_lines,
)
from typ2pptx.core.typst_svg_parser import TextSegment


def _seg(text: str, x: float, y: float, width: float, font_size: float = 12.0):
    """Helper: create a plain regular-font TextSegment."""
    return TextSegment(
        text=text,
        x=x,
        y=y,
        width=width,
        height=font_size,
        font_size=font_size,
        font_variant="regular",
        fill_color="#000000",
    )


def _make_paragraph_line_groups(
    num_lines: int = 5,
    line_width: float = 700.0,
    left_x: float = 50.0,
    font_size: float = 14.0,
    line_spacing: float = 20.0,
) -> list:
    """Build baseline-grouped line groups that look like a wrapped paragraph.

    Each line is a single full-width segment with identical font and left
    edge, so the paragraph heuristic would happily merge them when enabled.
    """
    line_groups = []
    for i in range(num_lines):
        seg = _seg(
            text=f"wrapped line {i} with enough characters to look full width",
            x=left_x,
            y=100.0 + i * line_spacing,
            width=line_width,
            font_size=font_size,
        )
        line_groups.append([seg])
    return line_groups


class TestConversionConfigDefault:
    """The config default must keep paragraph detection OFF."""

    def test_default_is_disabled(self):
        config = ConversionConfig()
        assert config.detect_paragraphs is False, (
            "detect_paragraphs must default to False so tables and other "
            "tightly-packed layouts are not mis-merged"
        )

    def test_can_be_enabled(self):
        config = ConversionConfig(detect_paragraphs=True)
        assert config.detect_paragraphs is True


class TestMergeParagraphLinesDisabled:
    """With detect_paragraphs=False (default), no merging happens."""

    def test_empty_input_returns_empty(self):
        assert _merge_paragraph_lines([], page_width=800) == []

    def test_paragraph_lines_kept_separate_by_default(self):
        """A clean 5-line paragraph must stay as 5 'line' groups."""
        line_groups = _make_paragraph_line_groups(num_lines=5)

        result = _merge_paragraph_lines(line_groups, page_width=800.0)

        assert len(result) == 5
        assert all(group["type"] == "line" for group in result), (
            "Every group must remain a single 'line' when detection is off"
        )
        # Segments preserved in order, one-per-line
        for i, group in enumerate(result):
            assert len(group["segments"]) == 1
            assert group["segments"][0].text == f"wrapped line {i} with enough characters to look full width"

    def test_explicit_false_matches_default(self):
        line_groups = _make_paragraph_line_groups(num_lines=3)
        default_result = _merge_paragraph_lines(line_groups, page_width=800.0)
        explicit_result = _merge_paragraph_lines(
            line_groups, page_width=800.0, detect_paragraphs=False,
        )
        assert len(default_result) == len(explicit_result) == 3

    def test_table_like_content_not_merged(self):
        """Simulate a simple 3-row "table": several short segments per row,
        vertically stacked. The old heuristic could occasionally pull these
        into one block; with detection off they must remain row-wise."""
        rows = []
        for row_idx in range(3):
            y = 100.0 + row_idx * 18.0
            rows.append([
                _seg("cell-a", x=50.0, y=y, width=40.0),
                _seg("cell-b", x=150.0, y=y, width=40.0),
                _seg("cell-c", x=250.0, y=y, width=40.0),
            ])

        result = _merge_paragraph_lines(rows, page_width=800.0)

        assert len(result) == 3
        for row_idx, group in enumerate(result):
            assert group["type"] == "line"
            assert len(group["segments"]) == 3, (
                f"Row {row_idx} should keep all 3 cells as independent segments"
            )


class TestMergeParagraphLinesEnabled:
    """With detect_paragraphs=True, the heuristic kicks in."""

    def test_paragraph_lines_merged_when_enabled(self):
        line_groups = _make_paragraph_line_groups(num_lines=5)

        result = _merge_paragraph_lines(
            line_groups, page_width=800.0, detect_paragraphs=True,
        )

        # Exactly one merged paragraph group expected
        paragraph_groups = [g for g in result if g["type"] == "paragraph"]
        assert len(paragraph_groups) == 1, (
            f"Expected one merged paragraph group, got {result}"
        )
        assert len(paragraph_groups[0]["lines"]) == 5

    def test_single_line_stays_as_line(self):
        """A single line cannot form a paragraph -> stays as type='line'
        even when detection is enabled."""
        line_groups = _make_paragraph_line_groups(num_lines=1)

        result = _merge_paragraph_lines(
            line_groups, page_width=800.0, detect_paragraphs=True,
        )

        assert len(result) == 1
        assert result[0]["type"] == "line"
