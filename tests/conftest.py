"""Shared fixtures for typ2pptx tests."""
import os
import shutil
import pytest
from pathlib import Path

# Root directories
TESTS_DIR = Path(__file__).parent
PROJECT_DIR = TESTS_DIR.parent
TYP_SOURCES_DIR = TESTS_DIR / "typ_sources"
OUTPUT_DIR = TESTS_DIR / "output"

# Tool paths
TYPST_TS_CLI = shutil.which("typst-ts-cli") or "/home/admin/bin/typst-ts-cli"


@pytest.fixture(scope="session")
def output_dir():
    """Ensure output directory exists."""
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    return OUTPUT_DIR


@pytest.fixture(scope="session")
def typ_sources_dir():
    """Return path to typ source files."""
    return TYP_SOURCES_DIR


@pytest.fixture(scope="session")
def typst_ts_cli():
    """Return typst-ts-cli path."""
    return TYPST_TS_CLI


@pytest.fixture(scope="session")
def basic_text_svg(typ_sources_dir, typst_ts_cli):
    """Compile basic_text.typ to SVG and return the SVG path."""
    from typ2pptx.core.converter import compile_typst_to_svg
    typ_path = typ_sources_dir / "basic_text.typ"
    svg_path = compile_typst_to_svg(str(typ_path), typst_ts_cli=typst_ts_cli)
    return svg_path


@pytest.fixture(scope="session")
def shapes_test_svg(typ_sources_dir, typst_ts_cli):
    """Compile shapes_test.typ to SVG and return the SVG path."""
    from typ2pptx.core.converter import compile_typst_to_svg
    typ_path = typ_sources_dir / "shapes_test.typ"
    svg_path = compile_typst_to_svg(str(typ_path), typst_ts_cli=typst_ts_cli)
    return svg_path


@pytest.fixture(scope="session")
def speaker_notes_svg(typ_sources_dir, typst_ts_cli):
    """Compile speaker_notes_test.typ to SVG and return the SVG path."""
    from typ2pptx.core.converter import compile_typst_to_svg
    typ_path = typ_sources_dir / "speaker_notes_test.typ"
    svg_path = compile_typst_to_svg(str(typ_path), typst_ts_cli=typst_ts_cli)
    return svg_path


@pytest.fixture(scope="session")
def basic_text_parsed(basic_text_svg):
    """Parse basic_text SVG and return TypstSVGData."""
    from typ2pptx.core.typst_svg_parser import parse_typst_svg
    return parse_typst_svg(basic_text_svg)


@pytest.fixture(scope="session")
def shapes_test_parsed(shapes_test_svg):
    """Parse shapes_test SVG and return TypstSVGData."""
    from typ2pptx.core.typst_svg_parser import parse_typst_svg
    return parse_typst_svg(shapes_test_svg)


@pytest.fixture(scope="session")
def basic_text_pptx(basic_text_svg, output_dir, typ_sources_dir):
    """Convert basic_text to PPTX and return the path."""
    from typ2pptx.core.converter import (
        TypstSVGConverter, ConversionConfig, query_speaker_notes,
    )
    output_path = str(output_dir / "basic_text_test.pptx")
    notes = query_speaker_notes(
        str(typ_sources_dir / "basic_text.typ")    )
    config = ConversionConfig()
    converter = TypstSVGConverter(config)
    converter.convert(basic_text_svg, output_path, speaker_notes=notes)
    return output_path


@pytest.fixture(scope="session")
def shapes_test_pptx(shapes_test_svg, output_dir, typ_sources_dir):
    """Convert shapes_test to PPTX and return the path."""
    from typ2pptx.core.converter import (
        TypstSVGConverter, ConversionConfig, query_speaker_notes,
    )
    output_path = str(output_dir / "shapes_test_test.pptx")
    notes = query_speaker_notes(
        str(typ_sources_dir / "shapes_test.typ")    )
    config = ConversionConfig()
    converter = TypstSVGConverter(config)
    converter.convert(shapes_test_svg, output_path, speaker_notes=notes)
    return output_path


@pytest.fixture(scope="session")
def speaker_notes_pptx(speaker_notes_svg, output_dir, typ_sources_dir):
    """Convert speaker_notes_test to PPTX and return the path."""
    from typ2pptx.core.converter import (
        TypstSVGConverter, ConversionConfig, query_speaker_notes,
    )
    output_path = str(output_dir / "speaker_notes_test_test.pptx")
    notes = query_speaker_notes(
        str(typ_sources_dir / "speaker_notes_test.typ")    )
    config = ConversionConfig()
    converter = TypstSVGConverter(config)
    converter.convert(speaker_notes_svg, output_path, speaker_notes=notes)
    return output_path


@pytest.fixture(scope="session")
def math_test_svg(typ_sources_dir, typst_ts_cli):
    """Compile math_test.typ to SVG and return the SVG path."""
    from typ2pptx.core.converter import compile_typst_to_svg
    typ_path = typ_sources_dir / "math_test.typ"
    svg_path = compile_typst_to_svg(str(typ_path), typst_ts_cli=typst_ts_cli)
    return svg_path


@pytest.fixture(scope="session")
def math_test_parsed(math_test_svg):
    """Parse math_test SVG and return TypstSVGData."""
    from typ2pptx.core.typst_svg_parser import parse_typst_svg
    return parse_typst_svg(math_test_svg)


@pytest.fixture(scope="session")
def math_test_pptx(math_test_svg, output_dir, typ_sources_dir):
    """Convert math_test to PPTX and return the path."""
    from typ2pptx.core.converter import (
        TypstSVGConverter, ConversionConfig, query_speaker_notes,
    )
    output_path = str(output_dir / "math_test_test.pptx")
    notes = query_speaker_notes(
        str(typ_sources_dir / "math_test.typ")    )
    config = ConversionConfig()
    converter = TypstSVGConverter(config)
    converter.convert(math_test_svg, output_path, speaker_notes=notes)
    return output_path


@pytest.fixture(scope="session")
def chinese_test_svg(typ_sources_dir, typst_ts_cli):
    """Compile chinese_test.typ to SVG and return the SVG path."""
    from typ2pptx.core.converter import compile_typst_to_svg
    typ_path = typ_sources_dir / "chinese_test.typ"
    svg_path = compile_typst_to_svg(str(typ_path), typst_ts_cli=typst_ts_cli)
    return svg_path


@pytest.fixture(scope="session")
def chinese_test_parsed(chinese_test_svg):
    """Parse chinese_test SVG and return TypstSVGData."""
    from typ2pptx.core.typst_svg_parser import parse_typst_svg
    return parse_typst_svg(chinese_test_svg)


@pytest.fixture(scope="session")
def chinese_test_pptx(chinese_test_svg, output_dir, typ_sources_dir):
    """Convert chinese_test to PPTX and return the path."""
    from typ2pptx.core.converter import (
        TypstSVGConverter, ConversionConfig, query_speaker_notes,
    )
    output_path = str(output_dir / "chinese_test_test.pptx")
    notes = query_speaker_notes(
        str(typ_sources_dir / "chinese_test.typ")    )
    config = ConversionConfig()
    converter = TypstSVGConverter(config)
    converter.convert(chinese_test_svg, output_path, speaker_notes=notes)
    return output_path


@pytest.fixture(scope="session")
def inline_math_test_svg(typ_sources_dir, typst_ts_cli):
    """Compile inline_math_test.typ to SVG and return the SVG path."""
    from typ2pptx.core.converter import compile_typst_to_svg
    typ_path = typ_sources_dir / "inline_math_test.typ"
    svg_path = compile_typst_to_svg(str(typ_path), typst_ts_cli=typst_ts_cli)
    return svg_path


@pytest.fixture(scope="session")
def inline_math_test_parsed(inline_math_test_svg):
    """Parse inline_math_test SVG and return TypstSVGData."""
    from typ2pptx.core.typst_svg_parser import parse_typst_svg
    return parse_typst_svg(inline_math_test_svg)


@pytest.fixture(scope="session")
def inline_math_test_pptx(inline_math_test_svg, output_dir, typ_sources_dir):
    """Convert inline_math_test to PPTX and return the path."""
    from typ2pptx.core.converter import (
        TypstSVGConverter, ConversionConfig, query_speaker_notes,
    )
    output_path = str(output_dir / "inline_math_test_test.pptx")
    notes = query_speaker_notes(
        str(typ_sources_dir / "inline_math_test.typ")    )
    config = ConversionConfig()
    converter = TypstSVGConverter(config)
    converter.convert(inline_math_test_svg, output_path, speaker_notes=notes)
    return output_path


@pytest.fixture(scope="session")
def table_test_svg(typ_sources_dir, typst_ts_cli):
    """Compile table_test.typ to SVG and return the SVG path."""
    from typ2pptx.core.converter import compile_typst_to_svg
    typ_path = typ_sources_dir / "table_test.typ"
    svg_path = compile_typst_to_svg(str(typ_path), typst_ts_cli=typst_ts_cli)
    return svg_path


@pytest.fixture(scope="session")
def table_test_parsed(table_test_svg):
    """Parse table_test SVG and return TypstSVGData."""
    from typ2pptx.core.typst_svg_parser import parse_typst_svg
    return parse_typst_svg(table_test_svg)


@pytest.fixture(scope="session")
def table_test_pptx(table_test_svg, output_dir, typ_sources_dir):
    """Convert table_test to PPTX and return the path."""
    from typ2pptx.core.converter import (
        TypstSVGConverter, ConversionConfig, query_speaker_notes,
    )
    output_path = str(output_dir / "table_test_test.pptx")
    notes = query_speaker_notes(
        str(typ_sources_dir / "table_test.typ")    )
    config = ConversionConfig()
    converter = TypstSVGConverter(config)
    converter.convert(table_test_svg, output_path, speaker_notes=notes)
    return output_path


@pytest.fixture(scope="session")
def pinit_test_svg(typ_sources_dir, typst_ts_cli):
    """Compile pinit_test.typ to SVG and return the SVG path."""
    from typ2pptx.core.converter import compile_typst_to_svg
    typ_path = typ_sources_dir / "pinit_test.typ"
    svg_path = compile_typst_to_svg(str(typ_path), typst_ts_cli=typst_ts_cli)
    return svg_path


@pytest.fixture(scope="session")
def pinit_test_parsed(pinit_test_svg):
    """Parse pinit_test SVG and return TypstSVGData."""
    from typ2pptx.core.typst_svg_parser import parse_typst_svg
    return parse_typst_svg(pinit_test_svg)


@pytest.fixture(scope="session")
def pinit_test_pptx(pinit_test_svg, output_dir, typ_sources_dir):
    """Convert pinit_test to PPTX and return the path."""
    from typ2pptx.core.converter import (
        TypstSVGConverter, ConversionConfig, query_speaker_notes,
    )
    output_path = str(output_dir / "pinit_test_test.pptx")
    notes = query_speaker_notes(
        str(typ_sources_dir / "pinit_test.typ")    )
    config = ConversionConfig()
    converter = TypstSVGConverter(config)
    converter.convert(pinit_test_svg, output_path, speaker_notes=notes)
    return output_path


@pytest.fixture(scope="session")
def link_test_svg(typ_sources_dir, typst_ts_cli):
    """Compile link_test.typ to SVG and return the SVG path."""
    from typ2pptx.core.converter import compile_typst_to_svg
    typ_path = typ_sources_dir / "link_test.typ"
    svg_path = compile_typst_to_svg(str(typ_path), typst_ts_cli=typst_ts_cli)
    return svg_path


@pytest.fixture(scope="session")
def link_test_parsed(link_test_svg):
    """Parse link_test SVG and return TypstSVGData."""
    from typ2pptx.core.typst_svg_parser import parse_typst_svg
    return parse_typst_svg(link_test_svg)


@pytest.fixture(scope="session")
def link_test_pptx(link_test_svg, output_dir, typ_sources_dir):
    """Convert link_test to PPTX and return the path."""
    from typ2pptx.core.converter import (
        TypstSVGConverter, ConversionConfig, query_speaker_notes,
    )
    output_path = str(output_dir / "link_test_test.pptx")
    notes = query_speaker_notes(
        str(typ_sources_dir / "link_test.typ")    )
    config = ConversionConfig()
    converter = TypstSVGConverter(config)
    converter.convert(link_test_svg, output_path, speaker_notes=notes)
    return output_path


@pytest.fixture(scope="session")
def alignment_test_svg(typ_sources_dir, typst_ts_cli):
    """Compile alignment_test.typ to SVG and return the SVG path."""
    from typ2pptx.core.converter import compile_typst_to_svg
    typ_path = typ_sources_dir / "alignment_test.typ"
    svg_path = compile_typst_to_svg(str(typ_path), typst_ts_cli=typst_ts_cli)
    return svg_path


@pytest.fixture(scope="session")
def alignment_test_parsed(alignment_test_svg):
    """Parse alignment_test SVG and return TypstSVGData."""
    from typ2pptx.core.typst_svg_parser import parse_typst_svg
    return parse_typst_svg(alignment_test_svg)


@pytest.fixture(scope="session")
def alignment_test_pptx(alignment_test_svg, output_dir, typ_sources_dir):
    """Convert alignment_test to PPTX with default config (paragraph detection off)."""
    from typ2pptx.core.converter import (
        TypstSVGConverter, ConversionConfig, query_speaker_notes,
    )
    output_path = str(output_dir / "alignment_test_test.pptx")
    notes = query_speaker_notes(
        str(typ_sources_dir / "alignment_test.typ")    )
    config = ConversionConfig()
    converter = TypstSVGConverter(config)
    converter.convert(alignment_test_svg, output_path, speaker_notes=notes)
    return output_path


@pytest.fixture(scope="session")
def alignment_test_pptx_with_paragraphs(
    alignment_test_svg, output_dir, typ_sources_dir,
):
    """Convert alignment_test to PPTX with paragraph auto-detection enabled.

    Needed for the justified-text assertion: `justify` alignment is only
    emitted on merged paragraph groups, which require the opt-in heuristic.
    """
    from typ2pptx.core.converter import (
        TypstSVGConverter, ConversionConfig, query_speaker_notes,
    )
    output_path = str(output_dir / "alignment_test_with_paragraphs.pptx")
    notes = query_speaker_notes(
        str(typ_sources_dir / "alignment_test.typ")    )
    config = ConversionConfig(detect_paragraphs=True)
    converter = TypstSVGConverter(config)
    converter.convert(alignment_test_svg, output_path, speaker_notes=notes)
    return output_path


@pytest.fixture(scope="session")
def image_test_svg(typ_sources_dir, typst_ts_cli):
    """Compile image_test.typ to SVG and return the SVG path."""
    from typ2pptx.core.converter import compile_typst_to_svg
    typ_path = typ_sources_dir / "image_test.typ"
    svg_path = compile_typst_to_svg(str(typ_path), typst_ts_cli=typst_ts_cli)
    return svg_path


@pytest.fixture(scope="session")
def image_test_parsed(image_test_svg):
    """Parse image_test SVG and return TypstSVGData."""
    from typ2pptx.core.typst_svg_parser import parse_typst_svg
    return parse_typst_svg(image_test_svg)


@pytest.fixture(scope="session")
def image_test_pptx(image_test_svg, output_dir, typ_sources_dir):
    """Convert image_test to PPTX and return the path."""
    from typ2pptx.core.converter import (
        TypstSVGConverter, ConversionConfig, query_speaker_notes,
    )
    output_path = str(output_dir / "image_test_test.pptx")
    notes = query_speaker_notes(
        str(typ_sources_dir / "image_test.typ")    )
    config = ConversionConfig()
    converter = TypstSVGConverter(config)
    converter.convert(image_test_svg, output_path, speaker_notes=notes)
    return output_path


@pytest.fixture(scope="session")
def code_block_test_svg(typ_sources_dir, typst_ts_cli):
    """Compile code_block_test.typ to SVG and return the SVG path."""
    from typ2pptx.core.converter import compile_typst_to_svg
    typ_path = typ_sources_dir / "code_block_test.typ"
    svg_path = compile_typst_to_svg(str(typ_path), typst_ts_cli=typst_ts_cli)
    return svg_path


@pytest.fixture(scope="session")
def code_block_test_parsed(code_block_test_svg):
    """Parse code_block_test SVG and return TypstSVGData."""
    from typ2pptx.core.typst_svg_parser import parse_typst_svg
    return parse_typst_svg(code_block_test_svg)


@pytest.fixture(scope="session")
def code_block_test_pptx(code_block_test_svg, output_dir, typ_sources_dir):
    """Convert code_block_test to PPTX and return the path."""
    from typ2pptx.core.converter import (
        TypstSVGConverter, ConversionConfig, query_speaker_notes,
    )
    output_path = str(output_dir / "code_block_test_test.pptx")
    notes = query_speaker_notes(
        str(typ_sources_dir / "code_block_test.typ")    )
    config = ConversionConfig()
    converter = TypstSVGConverter(config)
    converter.convert(code_block_test_svg, output_path, speaker_notes=notes)
    return output_path

@pytest.fixture(scope="session")
def columns_test_svg(typ_sources_dir, typst_ts_cli):
    """Compile columns_test.typ to SVG and return the SVG path."""
    from typ2pptx.core.converter import compile_typst_to_svg
    typ_path = typ_sources_dir / "columns_test.typ"
    svg_path = compile_typst_to_svg(str(typ_path), typst_ts_cli=typst_ts_cli)
    return svg_path

@pytest.fixture(scope="session")
def columns_test_parsed(columns_test_svg):
    """Parse columns_test SVG and return TypstSVGData."""
    from typ2pptx.core.typst_svg_parser import parse_typst_svg
    return parse_typst_svg(columns_test_svg)

@pytest.fixture(scope="session")
def columns_test_pptx(columns_test_svg, output_dir, typ_sources_dir):
    """Convert columns_test to PPTX with default config (paragraph detection off)."""
    from typ2pptx.core.converter import (
        TypstSVGConverter, ConversionConfig, query_speaker_notes,
    )
    output_path = str(output_dir / "columns_test_test.pptx")
    notes = query_speaker_notes(
        str(typ_sources_dir / "columns_test.typ")    )
    config = ConversionConfig()
    converter = TypstSVGConverter(config)
    converter.convert(columns_test_svg, output_path, speaker_notes=notes)
    return output_path


@pytest.fixture(scope="session")
def columns_test_pptx_with_paragraphs(columns_test_svg, output_dir, typ_sources_dir):
    """Convert columns_test to PPTX with paragraph auto-detection enabled."""
    from typ2pptx.core.converter import (
        TypstSVGConverter, ConversionConfig, query_speaker_notes,
    )
    output_path = str(output_dir / "columns_test_with_paragraphs.pptx")
    notes = query_speaker_notes(
        str(typ_sources_dir / "columns_test.typ")    )
    config = ConversionConfig(detect_paragraphs=True)
    converter = TypstSVGConverter(config)
    converter.convert(columns_test_svg, output_path, speaker_notes=notes)
    return output_path
