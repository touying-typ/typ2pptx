"""Tests for the paired geometry/font-size scale (emu_per_px derivation).

Geometry (EMU per SVG unit) and font sizes (points per SVG unit) must share
one physical scale, otherwise text renders 12700 * font_size_scale / emu_per_px
larger than its layout boxes. These tests pin the derivation and guard the
pairing end-to-end.
"""
import re
import zipfile

import pytest

from typ2pptx.core.converter import (
    ConversionConfig, TypstSVGConverter, compile_typst_to_svg, EMU_PER_PX,
)


class TestEmuPerPxDerivation:
    """resolved_emu_per_px() pairs geometry scale with font_size_scale."""

    def test_default_config_matches_legacy_constant(self):
        # Default: font_size_scale=0.75 (CSS px at 96 DPI)
        assert ConversionConfig().resolved_emu_per_px() == EMU_PER_PX == 9525

    def test_pt_true_scale_yields_emu_per_pt(self):
        # 1 SVG unit = 1pt (typst.ts) -> 12700 EMU per unit
        config = ConversionConfig(font_size_scale=1.0)
        assert config.resolved_emu_per_px() == 12700.0

    def test_explicit_override_wins(self):
        config = ConversionConfig(font_size_scale=1.0, emu_per_px=9525.0)
        assert config.resolved_emu_per_px() == 9525.0

    def test_converter_uses_resolved_value(self):
        converter = TypstSVGConverter(ConversionConfig(font_size_scale=1.0))
        assert converter._emu_per_px == 12700.0


def _slide_size_and_runs(pptx_path):
    """Return (slide_cx, slide_cy, [(off_x, ext_cx, sz), ...]) from a PPTX."""
    with zipfile.ZipFile(pptx_path) as zf:
        pres = zf.read('ppt/presentation.xml').decode('utf-8')
        m = re.search(r'<p:sldSz cx="(\d+)" cy="(\d+)"', pres)
        slide_cx, slide_cy = int(m.group(1)), int(m.group(2))
        boxes = []
        for name in zf.namelist():
            if not re.fullmatch(r'ppt/slides/slide\d+\.xml', name):
                continue
            xml = zf.read(name).decode('utf-8')
            for sp in re.findall(r'<p:sp>.*?</p:sp>', xml, re.S):
                if '<p:txBody>' not in sp or 'txBox="1"' not in sp:
                    continue
                off = re.search(r'<a:off x="(-?\d+)" y="(-?\d+)"/>', sp)
                ext = re.search(r'<a:ext cx="(\d+)"', sp)
                sz = re.search(r'sz="(\d+)"', sp)
                if off and ext and sz:
                    boxes.append((
                        int(off.group(1)), int(ext.group(1)),
                        int(sz.group(1)),
                    ))
    return slide_cx, slide_cy, boxes


@pytest.fixture(scope="module")
def scale_pair_pptx(basic_text_svg, output_dir):
    """Convert the same SVG with the default pair (0.75) and pt-true (1.0)."""
    paths = {}
    for label, scale in (("default", 0.75), ("pt_true", 1.0)):
        out = str(output_dir / f"scale_consistency_{label}.pptx")
        converter = TypstSVGConverter(ConversionConfig(font_size_scale=scale))
        converter.convert(basic_text_svg, out)
        paths[label] = out
    return paths


class TestScalePairingEndToEnd:
    """The text/geometry proportion must be identical for any font_size_scale.

    If geometry ever falls back to a constant 9525 while font sizes scale
    independently, the pt-true output's boxes stay put while its font sizes
    grow 4/3 -- caught below.
    """

    def test_geometry_and_font_sizes_scale_together(self, scale_pair_pptx):
        d_cx, d_cy, d_boxes = _slide_size_and_runs(scale_pair_pptx["default"])
        p_cx, p_cy, p_boxes = _slide_size_and_runs(scale_pair_pptx["pt_true"])
        expected = 12700.0 / 9525.0  # = 1/0.75

        assert p_cx / d_cx == pytest.approx(expected, rel=1e-6)
        assert p_cy / d_cy == pytest.approx(expected, rel=1e-6)

        assert d_boxes and len(d_boxes) == len(p_boxes)
        for (dx, dcx, dsz), (px, pcx, psz) in zip(d_boxes, p_boxes):
            # font size and box geometry must grow by the same factor
            assert psz / dsz == pytest.approx(expected, rel=0.02), \
                f"font size ratio {psz}/{dsz} diverged"
            if dcx > 0:
                assert pcx / dcx == pytest.approx(expected, rel=0.02), \
                    f"box width ratio {pcx}/{dcx} diverged from font ratio"
            if dx > 0:
                assert px / dx == pytest.approx(expected, rel=0.02), \
                    f"box offset ratio {px}/{dx} diverged from font ratio"


@pytest.fixture(scope="module")
def pt_canvas_pptx(typ_sources_dir, typst_ts_cli, output_dir):
    """Convert a 2880x1620pt canvas at pt-true scale."""
    typ_path = typ_sources_dir / "pt_canvas_test.typ"
    svg_path = compile_typst_to_svg(str(typ_path), typst_ts_cli=typst_ts_cli)
    out = str(output_dir / "pt_canvas_test.pptx")
    config = ConversionConfig(font_size_scale=1.0)
    TypstSVGConverter(config).convert(svg_path, out)
    return out


class TestPtCanvasGolden:
    """Golden values for a 2880x1620pt canvas converted pt-true."""

    def test_slide_size_is_pt_true(self, pt_canvas_pptx):
        cx, cy, _ = _slide_size_and_runs(pt_canvas_pptx)
        # 2880pt x 12700 EMU/pt, 1620pt x 12700 EMU/pt
        assert cx == 36576000
        assert cy == 20574000

    def test_body_font_size_survives_round_trip(self, pt_canvas_pptx):
        _, _, boxes = _slide_size_and_runs(pt_canvas_pptx)
        sizes = {sz / 100 for _, _, sz in boxes}
        # 40pt body authored; typst.ts emits ~39.68 (its own ~0.8% rounding)
        assert any(sz == pytest.approx(40, abs=0.5) for sz in sizes), \
            f"expected ~40pt body run, got sizes {sorted(sizes)}"
