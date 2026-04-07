"""Tests for the SVG path → DrawingML pipeline."""
import pytest

from typ2pptx.scripts.svg_to_shapes import (
    parse_svg_path,
    svg_path_to_absolute,
    normalize_path_commands,
    path_commands_to_drawingml,
    PathCommand,
)


class TestParseSvgPath:
    """Test SVG path d-attribute parsing."""

    def test_simple_rect(self):
        """Parse a simple rectangle path."""
        d = "M 0 0 L 100 0 L 100 50 L 0 50 Z"
        commands = parse_svg_path(d)
        assert len(commands) >= 5
        assert commands[0].cmd == 'M'
        assert commands[-1].cmd == 'Z'

    def test_cubic_bezier(self):
        """Parse cubic bezier commands."""
        d = "M 0 0 C 10 20 30 40 50 60"
        commands = parse_svg_path(d)
        assert any(c.cmd == 'C' for c in commands)

    def test_relative_commands(self):
        """Parse relative commands (lowercase)."""
        d = "M 0 0 l 100 0 l 0 50 l -100 0 z"
        commands = parse_svg_path(d)
        assert any(c.cmd == 'l' for c in commands)

    def test_empty_path(self):
        """Empty path should return empty list."""
        commands = parse_svg_path("")
        assert len(commands) == 0

    def test_arc_command(self):
        """Parse arc commands."""
        d = "M 10 80 A 45 45 0 0 0 125 125"
        commands = parse_svg_path(d)
        assert any(c.cmd == 'A' for c in commands)


class TestSvgPathToAbsolute:
    """Test conversion of relative to absolute coordinates."""

    def test_relative_line(self):
        commands = parse_svg_path("M 10 20 l 30 40")
        abs_commands = svg_path_to_absolute(commands)
        # After conversion, l should become L with absolute coords
        line = [c for c in abs_commands if c.cmd == 'L']
        assert len(line) == 1
        assert line[0].args[0] == pytest.approx(40.0)  # 10 + 30
        assert line[0].args[1] == pytest.approx(60.0)  # 20 + 40

    def test_absolute_unchanged(self):
        """Already absolute commands should stay the same."""
        commands = parse_svg_path("M 10 20 L 30 40")
        abs_commands = svg_path_to_absolute(commands)
        line = [c for c in abs_commands if c.cmd == 'L']
        assert len(line) == 1
        assert line[0].args[0] == pytest.approx(30.0)
        assert line[0].args[1] == pytest.approx(40.0)


class TestNormalizePathCommands:
    """Test normalization of path commands to M, L, C, Z."""

    def test_line_preserved(self):
        commands = [
            PathCommand('M', [0, 0]),
            PathCommand('L', [100, 50]),
            PathCommand('Z', []),
        ]
        normalized = normalize_path_commands(commands)
        cmd_types = [c.cmd for c in normalized]
        assert 'M' in cmd_types
        assert 'L' in cmd_types
        assert 'Z' in cmd_types

    def test_quadratic_to_cubic(self):
        """Quadratic bezier (Q) should be normalized to cubic (C)."""
        commands = [
            PathCommand('M', [0, 0]),
            PathCommand('Q', [50, 100, 100, 0]),
        ]
        normalized = normalize_path_commands(commands)
        # Q should be converted to C
        assert any(c.cmd == 'C' for c in normalized)
        assert not any(c.cmd == 'Q' for c in normalized)

    def test_horizontal_line_preserved(self):
        """H command is kept as-is by normalize (only Q/S/T/A are converted)."""
        commands = [
            PathCommand('M', [0, 0]),
            PathCommand('H', [100]),
        ]
        normalized = normalize_path_commands(commands)
        assert any(c.cmd == 'H' for c in normalized)

    def test_vertical_line_preserved(self):
        """V command is kept as-is by normalize (only Q/S/T/A are converted)."""
        commands = [
            PathCommand('M', [0, 0]),
            PathCommand('V', [50]),
        ]
        normalized = normalize_path_commands(commands)
        assert any(c.cmd == 'V' for c in normalized)

    def test_smooth_cubic_to_cubic(self):
        """S (smooth cubic) should be normalized to C."""
        commands = [
            PathCommand('M', [0, 0]),
            PathCommand('C', [10, 20, 30, 40, 50, 60]),
            PathCommand('S', [80, 90, 100, 100]),
        ]
        normalized = normalize_path_commands(commands)
        # S should be converted to C
        assert not any(c.cmd == 'S' for c in normalized)
        c_cmds = [c for c in normalized if c.cmd == 'C']
        assert len(c_cmds) == 2  # Original C + converted S


class TestPathCommandsToDrawingml:
    """Test conversion from path commands to DrawingML XML."""

    def test_simple_rect(self):
        commands = [
            PathCommand('M', [0, 0]),
            PathCommand('L', [100, 0]),
            PathCommand('L', [100, 50]),
            PathCommand('L', [0, 50]),
            PathCommand('Z', []),
        ]
        xml, min_x, min_y, width, height = path_commands_to_drawingml(commands)
        assert "<a:moveTo>" in xml
        assert "<a:lnTo>" in xml
        assert "<a:close/>" in xml
        assert width > 0
        assert height > 0

    def test_with_offset(self):
        commands = [
            PathCommand('M', [0, 0]),
            PathCommand('L', [100, 0]),
            PathCommand('L', [100, 50]),
            PathCommand('L', [0, 50]),
            PathCommand('Z', []),
        ]
        xml, min_x, min_y, width, height = path_commands_to_drawingml(
            commands, offset_x=200, offset_y=300,
        )
        assert min_x == pytest.approx(200.0, abs=1)
        assert min_y == pytest.approx(300.0, abs=1)

    def test_with_scale(self):
        commands = [
            PathCommand('M', [0, 0]),
            PathCommand('L', [100, 0]),
            PathCommand('L', [100, 50]),
            PathCommand('Z', []),
        ]
        xml1, _, _, w1, h1 = path_commands_to_drawingml(commands, scale_x=1, scale_y=1)
        xml2, _, _, w2, h2 = path_commands_to_drawingml(commands, scale_x=2, scale_y=2)
        assert w2 == pytest.approx(w1 * 2, rel=0.1)
        assert h2 == pytest.approx(h1 * 2, rel=0.1)

    def test_cubic_bezier(self):
        commands = [
            PathCommand('M', [0, 0]),
            PathCommand('C', [10, 20, 30, 40, 50, 60]),
        ]
        xml, min_x, min_y, width, height = path_commands_to_drawingml(commands)
        assert "<a:cubicBezTo>" in xml

    def test_empty_commands(self):
        xml, min_x, min_y, width, height = path_commands_to_drawingml([])
        assert width == 0 or xml == ""
