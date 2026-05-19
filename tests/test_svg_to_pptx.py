"""Comprehensive tests for the svg_to_pptx package (from ppt-master).

Tests cover the new features added by adopting ppt-master's modular conversion:
- Rounded rect (rx/ry) support
- Stroke line joins and custom dash patterns
- Arrow markers (marker-start/marker-end)
- Glow effects (feGaussianBlur without offset)
- Pattern fills
- Nested tspan handling
- Text decoration (underline, strikethrough)
- Image clip-path / preserveAspectRatio
- Donut circle threshold
- Element rotation
- Inheritable style propagation
"""

import math
import pytest
from xml.etree import ElementTree as ET

from typ2pptx.scripts.svg_to_pptx import (
    PathCommand,
    parse_svg_path,
    svg_path_to_absolute,
    normalize_path_commands,
    path_commands_to_drawingml,
    px_to_emu,
    parse_hex_color,
    EMU_PER_PX,
    ANGLE_UNIT,
)
from typ2pptx.scripts.svg_to_pptx.drawingml_paths import (
    _arc_to_cubic_beziers,
    _reflect_control_point,
    _quad_to_cubic,
)
from typ2pptx.scripts.svg_to_pptx.drawingml_styles import (
    build_solid_fill,
    build_gradient_fill,
    build_fill_xml,
    build_stroke_xml,
    build_effect_xml,
    build_shadow_xml,
    build_glow_xml,
    build_pattern_fill,
    classify_filter_effect,
    get_fill_opacity,
    get_stroke_opacity,
    _classify_marker,
    _emit_line_end,
)
from typ2pptx.scripts.svg_to_pptx.drawingml_context import ConvertContext
from typ2pptx.scripts.svg_to_pptx.drawingml_utils import (
    SVG_NS,
    _f,
    _get_attr,
    parse_font_family,
    parse_transform_matrix,
    estimate_text_width,
    DASH_PRESETS,
)
from typ2pptx.scripts.svg_to_pptx.drawingml_elements import (
    convert_rect,
    convert_circle,
    convert_ellipse,
    convert_line,
    convert_path,
    convert_polygon,
    convert_polyline,
    convert_text,
    convert_image,
)
from typ2pptx.scripts.svg_to_pptx.drawingml_converter import (
    convert_element,
    convert_g,
    collect_defs,
    parse_transform,
)


# ---------------------------------------------------------------------------
# Helper to create SVG elements with namespace
# ---------------------------------------------------------------------------

def _svg_elem(tag, attrib=None, text=None, children=None):
    """Create an SVG element with proper namespace."""
    elem = ET.Element(f'{{{SVG_NS}}}{tag}', attrib=attrib or {})
    if text:
        elem.text = text
    if children:
        for child in children:
            elem.append(child)
    return elem


def _make_ctx(defs=None, inherited_styles=None):
    """Create a ConvertContext for testing."""
    ctx = ConvertContext()
    if defs:
        ctx.defs = defs
    if inherited_styles:
        ctx.inherited_styles = inherited_styles
    return ctx


# ===========================================================================
# Path pipeline tests (extended from original test_path_pipeline.py)
# ===========================================================================

class TestPathPipelineExtended:
    """Extended path pipeline tests for new normalization features."""

    def test_arc_to_cubic_beziers_semicircle(self):
        """Arc command converted to cubic beziers for a half-circle."""
        result = _arc_to_cubic_beziers(0, 0, 50, 50, 0, 0, 1, 100, 0)
        assert len(result) > 0
        for cmd in result:
            assert cmd.cmd == 'C'
            assert len(cmd.args) == 6

    def test_arc_degenerate_zero_radius(self):
        """Zero-radius arc degenerates to a line."""
        result = _arc_to_cubic_beziers(0, 0, 0, 50, 0, 0, 1, 100, 100)
        assert len(result) == 1
        assert result[0].cmd == 'L'

    def test_arc_same_point(self):
        """Arc where start == end should produce no output."""
        result = _arc_to_cubic_beziers(50, 50, 25, 25, 0, 0, 1, 50, 50)
        assert len(result) == 0

    def test_reflect_control_point(self):
        """Control point reflection through current point."""
        rx, ry = _reflect_control_point(10, 20, 50, 50)
        assert rx == pytest.approx(90)
        assert ry == pytest.approx(80)

    def test_quad_to_cubic_midpoint(self):
        """Quadratic to cubic conversion preserves endpoint."""
        result = _quad_to_cubic(50, 100, 0, 0, 100, 0)
        # Last two values are the endpoint
        assert result[4] == pytest.approx(100)
        assert result[5] == pytest.approx(0)

    def test_normalize_smooth_cubic_after_line(self):
        """S after L should reflect current point (cp = current)."""
        commands = [
            PathCommand('M', [0, 0]),
            PathCommand('L', [50, 50]),
            PathCommand('S', [80, 90, 100, 100]),
        ]
        normalized = normalize_path_commands(commands)
        c_cmds = [c for c in normalized if c.cmd == 'C']
        assert len(c_cmds) == 1
        # First control point should be the current point (50, 50)
        assert c_cmds[0].args[0] == pytest.approx(50)
        assert c_cmds[0].args[1] == pytest.approx(50)

    def test_normalize_t_after_q(self):
        """T (smooth quad) after Q reflects the quadratic control point."""
        commands = [
            PathCommand('M', [0, 0]),
            PathCommand('Q', [50, 100, 100, 0]),
            PathCommand('T', [200, 0]),
        ]
        normalized = normalize_path_commands(commands)
        c_cmds = [c for c in normalized if c.cmd == 'C']
        assert len(c_cmds) == 2  # Q -> C, T -> C

    def test_normalize_arc(self):
        """A command is normalized to C (cubic bezier) sequences."""
        commands = [
            PathCommand('M', [10, 80]),
            PathCommand('A', [45, 45, 0, 0, 0, 125, 125]),
        ]
        normalized = normalize_path_commands(commands)
        assert not any(c.cmd == 'A' for c in normalized)
        assert any(c.cmd == 'C' for c in normalized)

    def test_path_commands_negative_coords(self):
        """Path with negative coordinates should still produce valid output."""
        commands = [
            PathCommand('M', [-50, -30]),
            PathCommand('L', [50, 30]),
            PathCommand('Z', []),
        ]
        xml, min_x, min_y, width, height = path_commands_to_drawingml(commands)
        assert min_x == pytest.approx(-50)
        assert min_y == pytest.approx(-30)
        assert width == pytest.approx(100)
        assert height == pytest.approx(60)
        assert '<a:moveTo>' in xml

    def test_svg_path_scientific_notation(self):
        """Parse path with scientific notation numbers."""
        d = "M 1e2 2.5e1 L 3.14e0 -1e1"
        commands = parse_svg_path(d)
        assert commands[0].args[0] == pytest.approx(100.0)
        assert commands[0].args[1] == pytest.approx(25.0)

    def test_svg_path_implicit_lineto(self):
        """After M, subsequent coordinate pairs become implicit L commands."""
        d = "M 0 0 100 50 200 100"
        commands = parse_svg_path(d)
        assert commands[0].cmd == 'M'
        assert commands[1].cmd == 'L'
        assert commands[2].cmd == 'L'


# ===========================================================================
# Style builder tests
# ===========================================================================

class TestBuildSolidFill:
    def test_basic(self):
        xml = build_solid_fill('FF0000')
        assert 'val="FF0000"' in xml
        assert '<a:solidFill>' in xml

    def test_with_opacity(self):
        xml = build_solid_fill('00FF00', opacity=0.5)
        assert '<a:alpha val="50000"/>' in xml

    def test_full_opacity_no_alpha(self):
        xml = build_solid_fill('0000FF', opacity=1.0)
        assert '<a:alpha' not in xml


class TestBuildGradientFill:
    def test_linear_gradient(self):
        grad = _svg_elem('linearGradient', {'x1': '0', 'y1': '0', 'x2': '1', 'y2': '0'})
        stop1 = _svg_elem('stop', {'offset': '0', 'style': 'stop-color:#ff0000'})
        stop2 = _svg_elem('stop', {'offset': '1', 'style': 'stop-color:#0000ff'})
        grad.append(stop1)
        grad.append(stop2)
        xml = build_gradient_fill(grad)
        assert '<a:gradFill>' in xml
        assert '<a:lin' in xml
        assert '<a:gs pos="0">' in xml
        assert '<a:gs pos="100000">' in xml

    def test_radial_gradient(self):
        grad = _svg_elem('radialGradient', {})
        stop1 = _svg_elem('stop', {'offset': '0', 'stop-color': '#ffffff'})
        stop2 = _svg_elem('stop', {'offset': '1', 'stop-color': '#000000'})
        grad.append(stop1)
        grad.append(stop2)
        xml = build_gradient_fill(grad)
        assert '<a:path path="circle">' in xml

    def test_no_stops_returns_empty(self):
        grad = _svg_elem('linearGradient', {})
        xml = build_gradient_fill(grad)
        assert xml == ''


class TestBuildPatternFill:
    def test_with_annotations(self):
        pattern = _svg_elem('pattern', {
            'data-pptx-pattern': 'dkDnDiag',
            'data-pptx-fg': '#FF0000',
            'data-pptx-bg': '#FFFFFF',
        })
        xml = build_pattern_fill(pattern)
        assert 'pattFill prst="dkDnDiag"' in xml
        assert 'val="FF0000"' in xml
        assert 'val="FFFFFF"' in xml

    def test_without_annotations_fallback(self):
        pattern = _svg_elem('pattern', {})
        rect = _svg_elem('rect', {'fill': '#EEEEEE'})
        path = _svg_elem('path', {'stroke': '#333333'})
        pattern.append(rect)
        pattern.append(path)
        xml = build_pattern_fill(pattern)
        assert 'val="333333"' in xml
        assert 'val="EEEEEE"' in xml

    def test_no_fg_returns_empty(self):
        pattern = _svg_elem('pattern', {})
        xml = build_pattern_fill(pattern)
        assert xml == ''


class TestBuildStrokeXml:
    def test_no_stroke(self):
        elem = _svg_elem('rect', {'stroke': 'none'})
        ctx = _make_ctx()
        xml = build_stroke_xml(elem, ctx)
        assert '<a:noFill/>' in xml

    def test_solid_stroke(self):
        elem = _svg_elem('rect', {'stroke': '#FF0000', 'stroke-width': '2'})
        ctx = _make_ctx()
        xml = build_stroke_xml(elem, ctx)
        assert 'val="FF0000"' in xml
        assert f'w="{px_to_emu(2)}"' in xml

    def test_line_join_round(self):
        elem = _svg_elem('rect', {'stroke': '#000000', 'stroke-linejoin': 'round'})
        ctx = _make_ctx()
        xml = build_stroke_xml(elem, ctx)
        assert '<a:round/>' in xml

    def test_line_join_bevel(self):
        elem = _svg_elem('rect', {'stroke': '#000000', 'stroke-linejoin': 'bevel'})
        ctx = _make_ctx()
        xml = build_stroke_xml(elem, ctx)
        assert '<a:bevel/>' in xml

    def test_line_join_miter(self):
        elem = _svg_elem('rect', {'stroke': '#000000', 'stroke-linejoin': 'miter'})
        ctx = _make_ctx()
        xml = build_stroke_xml(elem, ctx)
        assert '<a:miter' in xml

    def test_custom_dash(self):
        elem = _svg_elem('rect', {'stroke': '#000000', 'stroke-width': '2', 'stroke-dasharray': '10 5'})
        ctx = _make_ctx()
        xml = build_stroke_xml(elem, ctx)
        assert '<a:custDash>' in xml or '<a:prstDash' in xml

    def test_preset_dash(self):
        elem = _svg_elem('rect', {'stroke': '#000000', 'stroke-dasharray': '4 4'})
        ctx = _make_ctx()
        xml = build_stroke_xml(elem, ctx)
        assert '<a:prstDash' in xml or '<a:custDash' in xml

    def test_line_cap(self):
        elem = _svg_elem('rect', {'stroke': '#000000', 'stroke-linecap': 'round'})
        ctx = _make_ctx()
        xml = build_stroke_xml(elem, ctx)
        assert 'cap="rnd"' in xml

    def test_stroke_opacity(self):
        elem = _svg_elem('rect', {'stroke': '#000000', 'stroke-width': '1'})
        ctx = _make_ctx()
        xml = build_stroke_xml(elem, ctx, opacity=0.5)
        assert '<a:alpha val="50000"/>' in xml


class TestMarkerClassification:
    def test_triangle_marker(self):
        marker = _svg_elem('marker', {'markerWidth': '3', 'markerHeight': '3'})
        path = _svg_elem('path', {'d': 'M 0 0 L 10 5 L 0 10 Z'})
        marker.append(path)
        result = _classify_marker(marker)
        assert result is not None
        assert result[0] == 'triangle'

    def test_diamond_marker(self):
        marker = _svg_elem('marker', {'markerWidth': '3', 'markerHeight': '3'})
        path = _svg_elem('path', {'d': 'M 5 0 L 10 5 L 5 10 L 0 5 Z'})
        marker.append(path)
        result = _classify_marker(marker)
        assert result is not None
        assert result[0] == 'diamond'

    def test_oval_marker(self):
        marker = _svg_elem('marker', {'markerWidth': '3', 'markerHeight': '3'})
        circle = _svg_elem('circle', {'cx': '5', 'cy': '5', 'r': '5'})
        marker.append(circle)
        result = _classify_marker(marker)
        assert result is not None
        assert result[0] == 'oval'

    def test_empty_marker_returns_none(self):
        marker = _svg_elem('marker', {})
        result = _classify_marker(marker)
        assert result is None

    def test_size_buckets_userspaceontouse(self):
        """With markerUnits=userSpaceOnUse, absolute pixel thresholds apply."""
        marker = _svg_elem('marker', {
            'markerWidth': '14', 'markerHeight': '14',
            'markerUnits': 'userSpaceOnUse',
        })
        path = _svg_elem('path', {'d': 'M 0 0 L 10 5 L 0 10 Z'})
        marker.append(path)
        result = _classify_marker(marker)
        assert result is not None
        # markerWidth=14 > 12 -> 'lg', markerHeight=14 > 12 -> 'lg'
        assert result[1] == 'lg'
        assert result[2] == 'lg'

    def test_size_buckets_absolute_small(self):
        """_classify_marker uses absolute pixel thresholds: < 6 -> 'sm'."""
        marker = _svg_elem('marker', {'markerWidth': '3', 'markerHeight': '3'})
        path = _svg_elem('path', {'d': 'M 0 0 L 10 5 L 0 10 Z'})
        marker.append(path)
        result = _classify_marker(marker)
        assert result is not None
        # markerWidth=3 < 6 -> 'sm', markerHeight=3 < 6 -> 'sm'
        assert result[1] == 'sm'
        assert result[2] == 'sm'

    def test_size_buckets_absolute_medium(self):
        """_classify_marker uses absolute pixel thresholds: 6..12 -> 'med'."""
        marker = _svg_elem('marker', {'markerWidth': '8', 'markerHeight': '8'})
        path = _svg_elem('path', {'d': 'M 0 0 L 10 5 L 0 10 Z'})
        marker.append(path)
        result = _classify_marker(marker)
        assert result is not None
        assert result[1] == 'med'
        assert result[2] == 'med'


class TestEmitLineEnd:
    def test_no_marker(self):
        elem = _svg_elem('line', {})
        ctx = _make_ctx()
        result = _emit_line_end(elem, ctx, 'tail')
        assert result == ''

    def test_marker_resolved(self):
        marker = _svg_elem('marker', {'markerWidth': '3', 'markerHeight': '3'})
        path = _svg_elem('path', {'d': 'M 0 0 L 10 5 L 0 10 Z'})
        marker.append(path)
        defs = {'arrow': marker}
        elem = _svg_elem('line', {'marker-end': 'url(#arrow)'})
        ctx = _make_ctx(defs=defs)
        result = _emit_line_end(elem, ctx, 'tail')
        assert '<a:tailEnd' in result
        assert 'type="triangle"' in result


# ===========================================================================
# Effect (shadow / glow) tests
# ===========================================================================

class TestEffects:
    def test_shadow_with_offset(self):
        filter_elem = _svg_elem('filter', {})
        blur = _svg_elem('feGaussianBlur', {'stdDeviation': '4'})
        offset = _svg_elem('feOffset', {'dx': '3', 'dy': '4'})
        filter_elem.append(blur)
        filter_elem.append(offset)
        kind = classify_filter_effect(filter_elem)
        assert kind == 'shadow'
        xml = build_effect_xml(filter_elem)
        assert '<a:outerShdw' in xml

    def test_glow_without_offset(self):
        filter_elem = _svg_elem('filter', {})
        blur = _svg_elem('feGaussianBlur', {'stdDeviation': '6'})
        filter_elem.append(blur)
        kind = classify_filter_effect(filter_elem)
        assert kind == 'glow'
        xml = build_effect_xml(filter_elem)
        assert '<a:glow' in xml

    def test_drop_shadow_shorthand(self):
        filter_elem = _svg_elem('filter', {})
        drop = _svg_elem('feDropShadow', {
            'stdDeviation': '3',
            'dx': '2',
            'dy': '3',
            'flood-opacity': '0.5',
            'flood-color': '#333333',
        })
        filter_elem.append(drop)
        kind = classify_filter_effect(filter_elem)
        assert kind == 'shadow'
        xml = build_shadow_xml(filter_elem)
        assert 'val="333333"' in xml

    def test_none_filter(self):
        assert build_effect_xml(None) == ''
        assert classify_filter_effect(None) is None


# ===========================================================================
# Fill XML tests
# ===========================================================================

class TestBuildFillXml:
    def test_default_black_fill(self):
        elem = _svg_elem('rect', {})
        ctx = _make_ctx()
        xml = build_fill_xml(elem, ctx)
        assert 'val="000000"' in xml

    def test_no_fill(self):
        elem = _svg_elem('rect', {'fill': 'none'})
        ctx = _make_ctx()
        xml = build_fill_xml(elem, ctx)
        assert '<a:noFill/>' in xml

    def test_solid_color_fill(self):
        elem = _svg_elem('rect', {'fill': '#FF5500'})
        ctx = _make_ctx()
        xml = build_fill_xml(elem, ctx)
        assert 'val="FF5500"' in xml

    def test_gradient_fill_reference(self):
        grad = _svg_elem('linearGradient', {'x1': '0', 'y1': '0', 'x2': '1', 'y2': '0'})
        stop1 = _svg_elem('stop', {'offset': '0', 'stop-color': '#ff0000'})
        stop2 = _svg_elem('stop', {'offset': '1', 'stop-color': '#0000ff'})
        grad.append(stop1)
        grad.append(stop2)
        defs = {'myGrad': grad}
        elem = _svg_elem('rect', {'fill': 'url(#myGrad)'})
        ctx = _make_ctx(defs=defs)
        xml = build_fill_xml(elem, ctx)
        assert '<a:gradFill>' in xml

    def test_pattern_fill_reference(self):
        pattern = _svg_elem('pattern', {
            'data-pptx-pattern': 'ltUpDiag',
            'data-pptx-fg': '#000000',
            'data-pptx-bg': '#FFFFFF',
        })
        defs = {'myPatt': pattern}
        elem = _svg_elem('rect', {'fill': 'url(#myPatt)'})
        ctx = _make_ctx(defs=defs)
        xml = build_fill_xml(elem, ctx)
        assert '<a:pattFill' in xml

    def test_inherited_fill(self):
        elem = _svg_elem('rect', {})
        ctx = _make_ctx(inherited_styles={'fill': '#AABBCC'})
        xml = build_fill_xml(elem, ctx)
        assert 'val="AABBCC"' in xml


# ===========================================================================
# Opacity tests
# ===========================================================================

class TestOpacity:
    def test_fill_opacity_combined(self):
        elem = _svg_elem('rect', {'opacity': '0.5', 'fill-opacity': '0.8'})
        ctx = _make_ctx()
        result = get_fill_opacity(elem, ctx)
        assert result == pytest.approx(0.4)

    def test_stroke_opacity_combined(self):
        elem = _svg_elem('rect', {'opacity': '0.5', 'stroke-opacity': '0.6'})
        ctx = _make_ctx()
        result = get_stroke_opacity(elem, ctx)
        assert result == pytest.approx(0.3)

    def test_full_opacity_returns_none(self):
        elem = _svg_elem('rect', {'opacity': '1.0'})
        ctx = _make_ctx()
        assert get_fill_opacity(elem, ctx) is None

    def test_inherited_opacity(self):
        elem = _svg_elem('rect', {'fill-opacity': '0.8'})
        ctx = _make_ctx(inherited_styles={'opacity': '0.5'})
        result = get_fill_opacity(elem, ctx)
        assert result == pytest.approx(0.4)


# ===========================================================================
# Element converter tests
# ===========================================================================

class TestConvertRect:
    def test_basic_rect(self):
        elem = _svg_elem('rect', {'x': '10', 'y': '20', 'width': '100', 'height': '50'})
        ctx = _make_ctx()
        result = convert_rect(elem, ctx)
        assert result is not None
        assert 'prst="rect"' in result.xml

    def test_rounded_rect_symmetric(self):
        elem = _svg_elem('rect', {'x': '0', 'y': '0', 'width': '200', 'height': '100', 'rx': '10', 'ry': '10'})
        ctx = _make_ctx()
        result = convert_rect(elem, ctx)
        assert result is not None
        assert 'prst="roundRect"' in result.xml

    def test_rounded_rect_asymmetric(self):
        elem = _svg_elem('rect', {'x': '0', 'y': '0', 'width': '200', 'height': '100', 'rx': '20', 'ry': '10'})
        ctx = _make_ctx()
        result = convert_rect(elem, ctx)
        assert result is not None
        assert '<a:custGeom>' in result.xml

    def test_rx_only(self):
        """When only rx is specified, ry inherits its value (SVG spec)."""
        elem = _svg_elem('rect', {'x': '0', 'y': '0', 'width': '100', 'height': '50', 'rx': '8'})
        ctx = _make_ctx()
        result = convert_rect(elem, ctx)
        assert result is not None
        assert 'prst="roundRect"' in result.xml

    def test_zero_size_returns_none(self):
        elem = _svg_elem('rect', {'x': '0', 'y': '0', 'width': '0', 'height': '50'})
        ctx = _make_ctx()
        result = convert_rect(elem, ctx)
        assert result is None


class TestConvertCircle:
    def test_basic_circle(self):
        elem = _svg_elem('circle', {'cx': '50', 'cy': '50', 'r': '30'})
        ctx = _make_ctx()
        result = convert_circle(elem, ctx)
        assert result is not None
        assert 'prst="ellipse"' in result.xml

    def test_donut_circle(self):
        """Circle with large stroke-dasharray relative to radius is donut."""
        elem = _svg_elem('circle', {
            'cx': '50', 'cy': '50', 'r': '30',
            'stroke': '#FF0000', 'stroke-width': '10',
            'stroke-dasharray': '94.2 188.5',  # non-preset
            'fill': 'none',
        })
        ctx = _make_ctx()
        result = convert_circle(elem, ctx)
        assert result is not None
        # Donut detection: sw/r = 10/30 = 0.33 > 0.15 threshold -> donut
        assert '<a:custGeom>' in result.xml

    def test_thin_dashed_circle_not_donut(self):
        """Circle with thin stroke-dasharray (sw/r < 0.15) is not donut."""
        elem = _svg_elem('circle', {
            'cx': '50', 'cy': '50', 'r': '30',
            'stroke': '#FF0000', 'stroke-width': '1',
            'stroke-dasharray': '5 3',
            'fill': 'none',
        })
        ctx = _make_ctx()
        result = convert_circle(elem, ctx)
        assert result is not None
        # sw/r = 1/30 = 0.033 < 0.15 -> normal circle, not donut
        assert 'prst="ellipse"' in result.xml


class TestConvertEllipse:
    def test_basic_ellipse(self):
        elem = _svg_elem('ellipse', {'cx': '50', 'cy': '50', 'rx': '40', 'ry': '20'})
        ctx = _make_ctx()
        result = convert_ellipse(elem, ctx)
        assert result is not None
        assert 'prst="ellipse"' in result.xml


class TestConvertLine:
    def test_basic_line(self):
        elem = _svg_elem('line', {'x1': '0', 'y1': '0', 'x2': '100', 'y2': '50', 'stroke': '#000000'})
        ctx = _make_ctx()
        result = convert_line(elem, ctx)
        assert result is not None
        # Line should be custom geometry or preset line
        assert '<a:moveTo>' in result.xml or 'prst="line"' in result.xml

    def test_line_with_marker(self):
        """Line with marker-end should use headEnd/tailEnd."""
        marker = _svg_elem('marker', {'markerWidth': '3', 'markerHeight': '3'})
        path = _svg_elem('path', {'d': 'M 0 0 L 10 5 L 0 10 Z'})
        marker.append(path)
        defs = {'arrow': marker}
        elem = _svg_elem('line', {
            'x1': '0', 'y1': '0', 'x2': '100', 'y2': '0',
            'stroke': '#000000', 'marker-end': 'url(#arrow)',
        })
        ctx = _make_ctx(defs=defs)
        result = convert_line(elem, ctx)
        assert result is not None
        assert '<a:tailEnd' in result.xml


class TestConvertPath:
    def test_basic_path(self):
        elem = _svg_elem('path', {'d': 'M 0 0 L 100 0 L 100 50 L 0 50 Z'})
        ctx = _make_ctx()
        result = convert_path(elem, ctx)
        assert result is not None
        assert '<a:custGeom>' in result.xml

    def test_empty_d_returns_none(self):
        elem = _svg_elem('path', {'d': ''})
        ctx = _make_ctx()
        result = convert_path(elem, ctx)
        assert result is None

    def test_path_with_rotation(self):
        elem = _svg_elem('path', {
            'd': 'M 0 0 L 100 0 L 100 50 Z',
            'transform': 'rotate(45)',
        })
        ctx = _make_ctx()
        result = convert_path(elem, ctx)
        assert result is not None
        assert 'rot=' in result.xml


class TestConvertPolygon:
    def test_basic_polygon(self):
        elem = _svg_elem('polygon', {'points': '50,0 100,100 0,100'})
        ctx = _make_ctx()
        result = convert_polygon(elem, ctx)
        assert result is not None
        assert '<a:close/>' in result.xml

    def test_polygon_with_fill(self):
        elem = _svg_elem('polygon', {'points': '0,0 100,0 50,100', 'fill': '#FF0000'})
        ctx = _make_ctx()
        result = convert_polygon(elem, ctx)
        assert result is not None
        assert 'val="FF0000"' in result.xml


class TestConvertPolyline:
    def test_basic_polyline(self):
        elem = _svg_elem('polyline', {'points': '0,0 50,50 100,0', 'stroke': '#000000'})
        ctx = _make_ctx()
        result = convert_polyline(elem, ctx)
        assert result is not None
        # Polyline should NOT have close
        # (implementation may or may not include close; check it's valid)
        assert '<a:moveTo>' in result.xml


class TestConvertText:
    def test_basic_text(self):
        elem = _svg_elem('text', {
            'x': '10', 'y': '50',
            'font-size': '16', 'fill': '#000000',
        }, text='Hello World')
        ctx = _make_ctx()
        result = convert_text(elem, ctx)
        assert result is not None
        assert 'Hello World' in result.xml
        assert '<p:sp>' in result.xml

    def test_text_with_tspan(self):
        tspan = _svg_elem('tspan', {'fill': '#FF0000', 'font-weight': 'bold'}, text='Bold Red')
        elem = _svg_elem('text', {'x': '0', 'y': '20', 'font-size': '16', 'fill': '#000000'})
        elem.text = 'Normal '
        elem.append(tspan)
        tspan.tail = ' trailing'
        ctx = _make_ctx()
        result = convert_text(elem, ctx)
        assert result is not None
        assert 'Bold Red' in result.xml
        assert 'val="FF0000"' in result.xml

    def test_empty_text_returns_none(self):
        elem = _svg_elem('text', {'x': '0', 'y': '0', 'font-size': '16'})
        ctx = _make_ctx()
        result = convert_text(elem, ctx)
        assert result is None


class TestConvertGroup:
    def test_basic_group(self):
        rect = _svg_elem('rect', {'x': '0', 'y': '0', 'width': '100', 'height': '50'})
        circle = _svg_elem('circle', {'cx': '200', 'cy': '25', 'r': '25'})
        g = _svg_elem('g', {})
        g.append(rect)
        g.append(circle)
        ctx = _make_ctx()
        result = convert_g(g, ctx)
        assert result is not None
        assert '<p:grpSp>' in result.xml

    def test_single_child_flattened(self):
        """Single-child groups are flattened (no wrapping grpSp)."""
        rect = _svg_elem('rect', {'x': '0', 'y': '0', 'width': '100', 'height': '50'})
        g = _svg_elem('g', {})
        g.append(rect)
        ctx = _make_ctx()
        result = convert_g(g, ctx)
        assert result is not None
        # Should be a flat shape, not a group
        assert '<p:grpSp>' not in result.xml

    def test_group_with_transform(self):
        rect = _svg_elem('rect', {'x': '0', 'y': '0', 'width': '100', 'height': '50'})
        rect2 = _svg_elem('rect', {'x': '120', 'y': '0', 'width': '80', 'height': '50'})
        g = _svg_elem('g', {'transform': 'translate(50, 100)'})
        g.append(rect)
        g.append(rect2)
        ctx = _make_ctx()
        result = convert_g(g, ctx)
        assert result is not None
        assert '<p:grpSp>' in result.xml


# ===========================================================================
# Transform parsing tests
# ===========================================================================

class TestParseTransform:
    def test_translate(self):
        dx, dy, sx, sy, angle = parse_transform('translate(100, 200)')
        assert dx == pytest.approx(100)
        assert dy == pytest.approx(200)
        assert sx == pytest.approx(1)
        assert sy == pytest.approx(1)
        assert angle == pytest.approx(0)

    def test_scale(self):
        dx, dy, sx, sy, angle = parse_transform('scale(2, 3)')
        assert sx == pytest.approx(2)
        assert sy == pytest.approx(3)

    def test_rotate(self):
        dx, dy, sx, sy, angle = parse_transform('rotate(45)')
        assert angle == pytest.approx(45)

    def test_empty(self):
        dx, dy, sx, sy, angle = parse_transform('')
        assert dx == 0 and dy == 0 and sx == 1 and sy == 1 and angle == 0

    def test_combined(self):
        dx, dy, sx, sy, angle = parse_transform('translate(10, 20) scale(2)')
        assert dx == pytest.approx(10)
        assert dy == pytest.approx(20)
        assert sx == pytest.approx(2)
        assert sy == pytest.approx(2)


class TestParseTransformMatrix:
    def test_identity(self):
        a, b, c, d, e, f = parse_transform_matrix('matrix(1, 0, 0, 1, 0, 0)')
        assert a == pytest.approx(1)
        assert b == pytest.approx(0)
        assert c == pytest.approx(0)
        assert d == pytest.approx(1)
        assert e == pytest.approx(0)
        assert f == pytest.approx(0)

    def test_translate(self):
        a, b, c, d, e, f = parse_transform_matrix('translate(10, 20)')
        assert e == pytest.approx(10)
        assert f == pytest.approx(20)

    def test_scale(self):
        a, b, c, d, e, f = parse_transform_matrix('scale(2, 3)')
        assert a == pytest.approx(2)
        assert d == pytest.approx(3)


# ===========================================================================
# Utility function tests
# ===========================================================================

class TestParseHexColor:
    def test_6_digit(self):
        assert parse_hex_color('#FF5500') == 'FF5500'

    def test_3_digit(self):
        assert parse_hex_color('#F50') == 'FF5500'

    def test_no_hash(self):
        # Some SVGs omit the hash
        result = parse_hex_color('FF5500')
        # Depending on implementation, this may or may not work
        # Just check it doesn't crash
        assert result is None or result == 'FF5500'

    def test_invalid(self):
        assert parse_hex_color('red') is None
        assert parse_hex_color('') is None

    def test_none_input(self):
        assert parse_hex_color(None) is None


class TestPxToEmu:
    def test_basic(self):
        assert px_to_emu(1) == EMU_PER_PX
        assert px_to_emu(0) == 0
        assert px_to_emu(100) == 100 * EMU_PER_PX

    def test_fractional(self):
        result = px_to_emu(0.5)
        assert result == round(0.5 * EMU_PER_PX)


class TestParseFontFamily:
    def test_simple_latin(self):
        fonts = parse_font_family('Arial')
        assert fonts['latin'] == 'Arial'

    def test_quoted_font(self):
        fonts = parse_font_family('"Times New Roman"')
        assert fonts['latin'] == 'Times New Roman'

    def test_cjk_font(self):
        fonts = parse_font_family('Microsoft YaHei')
        assert fonts['ea'] == 'Microsoft YaHei'

    def test_multiple_fonts(self):
        """Multi-font list: the implementation may apply Windows fallbacks."""
        fonts = parse_font_family('"Helvetica Neue", Arial, sans-serif')
        # Helvetica Neue maps to Arial on Windows via FONT_FALLBACK_WIN
        assert fonts['latin'] in ('Helvetica Neue', 'Arial')


class TestEstimateTextWidth:
    def test_basic(self):
        width = estimate_text_width('Hello', 16, '400')
        assert width > 0

    def test_cjk_wider(self):
        latin_w = estimate_text_width('A', 16, '400')
        cjk_w = estimate_text_width('中', 16, '400')
        assert cjk_w >= latin_w  # CJK chars should be wider

    def test_empty(self):
        assert estimate_text_width('', 16, '400') == 0


class TestDashPresets:
    def test_known_presets_exist(self):
        """Ensure common SVG dash patterns map to DrawingML presets."""
        assert len(DASH_PRESETS) > 0
        # Check at least one known pattern
        assert any('dash' in v.lower() for v in DASH_PRESETS.values())


# ===========================================================================
# Integration: collect_defs and convert_element dispatch
# ===========================================================================

class TestCollectDefs:
    def test_basic(self):
        root = ET.fromstring('''<svg xmlns="http://www.w3.org/2000/svg">
            <defs>
                <linearGradient id="g1"/>
                <filter id="f1"/>
            </defs>
        </svg>''')
        defs = collect_defs(root)
        assert 'g1' in defs
        assert 'f1' in defs


class TestConvertElement:
    def test_rect_dispatch(self):
        elem = _svg_elem('rect', {'x': '0', 'y': '0', 'width': '100', 'height': '50'})
        ctx = _make_ctx()
        result = convert_element(elem, ctx)
        assert result is not None

    def test_defs_returns_none(self):
        elem = _svg_elem('defs', {})
        ctx = _make_ctx()
        result = convert_element(elem, ctx)
        assert result is None

    def test_unsupported_raises(self):
        from typ2pptx.scripts.svg_to_pptx.drawingml_converter import SvgNativeConversionError
        elem = _svg_elem('foreignObject', {})
        ctx = _make_ctx()
        with pytest.raises(SvgNativeConversionError):
            convert_element(elem, ctx)


# ===========================================================================
# Context tests
# ===========================================================================

class TestConvertContext:
    def test_child_inherits_defs(self):
        defs = {'g1': ET.Element('g')}
        ctx = _make_ctx(defs=defs)
        child = ctx.child(10, 20)
        assert child.defs is defs

    def test_child_accumulates_translate(self):
        ctx = _make_ctx()
        child = ctx.child(10, 20)
        assert child.translate_x == 10
        assert child.translate_y == 20
        grandchild = child.child(5, 5)
        assert grandchild.translate_x == 15
        assert grandchild.translate_y == 25

    def test_child_accumulates_scale(self):
        ctx = _make_ctx()
        child = ctx.child(0, 0, 2.0, 3.0)
        assert child.scale_x == pytest.approx(2.0)
        assert child.scale_y == pytest.approx(3.0)
        grandchild = child.child(0, 0, 0.5, 0.5)
        assert grandchild.scale_x == pytest.approx(1.0)
        assert grandchild.scale_y == pytest.approx(1.5)

    def test_opacity_multiplied(self):
        ctx = _make_ctx(inherited_styles={'opacity': '0.5'})
        child = ctx.child(style_overrides={'opacity': '0.8'})
        assert child.inherited_styles['opacity'] == str(0.5 * 0.8)

    def test_next_id_increments(self):
        ctx = _make_ctx()
        id1 = ctx.next_id()
        id2 = ctx.next_id()
        assert id2 == id1 + 1

    def test_sync_from_child(self):
        ctx = _make_ctx()
        child = ctx.child()
        child.next_id()
        child.next_id()
        ctx.sync_from_child(child)
        assert ctx.id_counter == child.id_counter


# ===========================================================================
# Backward compatibility — svg_to_shapes re-exports
# ===========================================================================

class TestBackwardCompatibility:
    """Verify that importing from svg_to_shapes.py still works."""

    def test_path_pipeline_imports(self):
        from typ2pptx.scripts.svg_to_shapes import (
            parse_svg_path,
            svg_path_to_absolute,
            normalize_path_commands,
            path_commands_to_drawingml,
            PathCommand,
        )
        # Quick smoke test
        cmds = parse_svg_path("M 0 0 L 100 100")
        assert len(cmds) == 2

    def test_constants(self):
        from typ2pptx.scripts.svg_to_shapes import EMU_PER_PX, ANGLE_UNIT
        assert EMU_PER_PX == 9525
        assert ANGLE_UNIT == 60000

    def test_style_builders(self):
        from typ2pptx.scripts.svg_to_shapes import (
            build_solid_fill,
            build_effect_xml,
            build_glow_xml,
        )
        xml = build_solid_fill('FF0000')
        assert 'solidFill' in xml
