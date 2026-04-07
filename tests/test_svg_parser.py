"""Tests for typst_svg_parser module."""
import pytest
from typst2pptx.core.typst_svg_parser import (
    parse_transform,
    parse_viewbox,
    parse_typst_svg,
    parse_css_styles,
    _compute_accumulated_transform,
)


class TestParseTransform:
    """Test transform string parsing."""

    def test_translate(self):
        dx, dy, sx, sy, rotation = parse_transform("translate(10, 20)")
        assert dx == pytest.approx(10.0)
        assert dy == pytest.approx(20.0)
        assert sx == pytest.approx(1.0)
        assert sy == pytest.approx(1.0)

    def test_scale_two_args(self):
        dx, dy, sx, sy, rotation = parse_transform("scale(2, -3)")
        assert sx == pytest.approx(2.0)
        assert sy == pytest.approx(-3.0)

    def test_scale_one_arg(self):
        dx, dy, sx, sy, rotation = parse_transform("scale(0.5)")
        assert sx == pytest.approx(0.5)
        assert sy == pytest.approx(0.5)

    def test_matrix(self):
        # matrix(a, b, c, d, e, f) = [a c e; b d f; 0 0 1]
        dx, dy, sx, sy, rotation = parse_transform("matrix(2, 0, 0, 3, 10, 20)")
        assert dx == pytest.approx(10.0)
        assert dy == pytest.approx(20.0)
        assert sx == pytest.approx(2.0)
        assert sy == pytest.approx(3.0)

    def test_combined_translate_scale(self):
        dx, dy, sx, sy, rotation = parse_transform(
            "translate(100, 200) scale(0.24, -0.24)"
        )
        assert dx == pytest.approx(100.0)
        assert dy == pytest.approx(200.0)
        assert sx == pytest.approx(0.24)
        assert sy == pytest.approx(-0.24)

    def test_none_input(self):
        dx, dy, sx, sy, rotation = parse_transform(None)
        assert dx == pytest.approx(0.0)
        assert dy == pytest.approx(0.0)
        assert sx == pytest.approx(1.0)
        assert sy == pytest.approx(1.0)

    def test_empty_string(self):
        dx, dy, sx, sy, rotation = parse_transform("")
        assert dx == pytest.approx(0.0)
        assert dy == pytest.approx(0.0)
        assert sx == pytest.approx(1.0)
        assert sy == pytest.approx(1.0)


class TestParseViewbox:
    """Test viewbox string parsing."""

    def test_basic_viewbox(self):
        min_x, min_y, w, h = parse_viewbox("0 0 842 474")
        assert w == pytest.approx(842.0)
        assert h == pytest.approx(474.0)

    def test_viewbox_with_offset(self):
        min_x, min_y, w, h = parse_viewbox("10 20 800 600")
        assert min_x == pytest.approx(10.0)
        assert min_y == pytest.approx(20.0)
        assert w == pytest.approx(800.0)
        assert h == pytest.approx(600.0)


class TestComputeAccumulatedTransform:
    """Test transform chain accumulation."""

    def test_single_translate(self):
        transforms = [(10.0, 20.0, 1.0, 1.0, 0.0)]
        dx, dy, sx, sy = _compute_accumulated_transform(transforms)
        assert dx == pytest.approx(10.0)
        assert dy == pytest.approx(20.0)

    def test_translate_then_scale(self):
        transforms = [
            (100.0, 200.0, 1.0, 1.0, 0.0),  # translate(100, 200)
            (0.0, 0.0, 2.0, 2.0, 0.0),       # scale(2, 2)
        ]
        dx, dy, sx, sy = _compute_accumulated_transform(transforms)
        assert sx == pytest.approx(2.0)
        assert sy == pytest.approx(2.0)

    def test_scale_then_translate(self):
        """When scale comes first, the implementation adds translations directly."""
        transforms = [
            (0.0, 0.0, 2.0, 2.0, 0.0),       # scale(2, 2)
            (50.0, 100.0, 1.0, 1.0, 0.0),     # translate(50, 100)
        ]
        dx, dy, sx, sy = _compute_accumulated_transform(transforms)
        assert sx == pytest.approx(2.0)
        assert sy == pytest.approx(2.0)
        # Implementation sums translations directly
        assert dx == pytest.approx(50.0)
        assert dy == pytest.approx(100.0)

    def test_empty_transforms(self):
        dx, dy, sx, sy = _compute_accumulated_transform([])
        assert dx == pytest.approx(0.0)
        assert dy == pytest.approx(0.0)
        assert sx == pytest.approx(1.0)
        assert sy == pytest.approx(1.0)

    def test_multiple_translates(self):
        transforms = [
            (10.0, 20.0, 1.0, 1.0, 0.0),
            (30.0, 40.0, 1.0, 1.0, 0.0),
        ]
        dx, dy, sx, sy = _compute_accumulated_transform(transforms)
        assert dx == pytest.approx(40.0)  # 10 + 30
        assert dy == pytest.approx(60.0)  # 20 + 40


class TestParseCssStyles:
    """Test CSS style parsing."""

    def test_basic_css(self):
        css = ".tsel { font-size: 62px; }"
        styles = parse_css_styles(css)
        assert ".tsel" in styles
        assert styles[".tsel"]["font-size"] == "62px"

    def test_empty_css(self):
        styles = parse_css_styles("")
        assert styles == {}


class TestParseTypstSVG:
    """Test full SVG parsing with actual typst output."""

    def test_basic_text_page_count(self, basic_text_parsed):
        """basic_text.typ has 3 slides."""
        assert len(basic_text_parsed.pages) == 3

    def test_basic_text_viewbox_width(self, basic_text_parsed):
        """Viewbox width should be 842 (16:9 at typst-ts scale)."""
        assert basic_text_parsed.viewbox_width == pytest.approx(842.0)

    def test_basic_text_viewbox_height_is_total(self, basic_text_parsed):
        """Viewbox height is total height of all pages stacked vertically."""
        # 3 pages at 474px each = 1422px total
        assert basic_text_parsed.viewbox_height == pytest.approx(1422.0, abs=2.0)

    def test_basic_text_page_nums(self, basic_text_parsed):
        """Pages should be numbered 1-3."""
        page_nums = [p.page_num for p in basic_text_parsed.pages]
        assert page_nums == [1, 2, 3]

    def test_basic_text_has_text_segments(self, basic_text_parsed):
        """Each page should have text segments."""
        for page in basic_text_parsed.pages:
            assert len(page.text_segments) > 0, (
                f"Page {page.page_num} has no text segments"
            )

    def test_basic_text_font_variants(self, basic_text_parsed):
        """basic_text.typ uses 5 font variants: regular, bold, italic, bolditalic, mono."""
        variants = basic_text_parsed.font_variants
        assert len(variants) == 5
        # Check that all expected styles are present
        styles = {v.style for v in variants.values()}
        assert styles == {"regular", "bold", "italic", "bolditalic", "mono"}

    def test_basic_text_glyph_defs(self, basic_text_parsed):
        """There should be glyph definitions in glyph_defs."""
        assert len(basic_text_parsed.glyph_defs) > 0

    def test_shapes_test_page_count(self, shapes_test_parsed):
        """shapes_test.typ should produce at least 4 slides."""
        assert len(shapes_test_parsed.pages) >= 4

    def test_shapes_test_has_shapes(self, shapes_test_parsed):
        """shapes_test slides should have path/shape elements."""
        total_shapes = sum(
            len(page.shapes) for page in shapes_test_parsed.pages
        )
        assert total_shapes > 0, "No shapes found in shapes_test"

    def test_text_segment_has_xy_position(self, basic_text_parsed):
        """Text segments should have x, y position after transform chain."""
        seg = basic_text_parsed.pages[0].text_segments[0]
        assert hasattr(seg, 'x')
        assert hasattr(seg, 'y')
        assert isinstance(seg.x, float)
        assert isinstance(seg.y, float)

    def test_text_segment_has_content(self, basic_text_parsed):
        """Text segments should have text content."""
        page1_texts = [
            seg.text for seg in basic_text_parsed.pages[0].text_segments
        ]
        all_text = " ".join(page1_texts)
        assert len(all_text.strip()) > 0, "No text content on page 1"

    def test_text_segment_has_font_info(self, basic_text_parsed):
        """Text segments should have font size and variant."""
        seg = basic_text_parsed.pages[0].text_segments[0]
        assert hasattr(seg, 'font_size')
        assert hasattr(seg, 'font_variant')
        assert seg.font_size > 0

    def test_font_variant_prefixes_are_unique(self, basic_text_parsed):
        """Each font variant should have a unique 5-char prefix."""
        prefixes = list(basic_text_parsed.font_variants.keys())
        assert len(prefixes) == len(set(prefixes))
        for prefix in prefixes:
            assert len(prefix) == 5, f"Prefix '{prefix}' is not 5 chars"

    def test_page_has_dimensions(self, basic_text_parsed):
        """Each page should have width and height."""
        for page in basic_text_parsed.pages:
            assert page.width > 0
            assert page.height > 0
            assert page.width == pytest.approx(842.0)
            assert page.height == pytest.approx(474.0)
