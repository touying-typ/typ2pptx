# typ2pptx

> [!WARNING]  
> This project is a vibe coding project and is currently under active development. Many conversion errors and edge cases remain.
> 
> If you encounter conversion issues, please open an issue with a minimal example with source files and images showing the problem — I'll try to fix them. The more issue reports and examples you provide, the faster I can improve the converter.

Convert [Typst](https://typst.app/) presentations (using the [Touying](https://github.com/touying-typ/touying) framework) to editable PowerPoint (.pptx) files.

## Features

- **Editable text**: All text is extracted as editable, selectable, copyable PowerPoint text (not rasterized images)
- **Font variants**: Detects and preserves regular, bold, italic, bold-italic, and monospace text styles
- **Math formulas**: Inline math rendered as Cambria Math text with PPTX-native sub/superscript baseline offsets (auto-sized by PowerPoint); display math rendered as native DrawingML curve shapes (glyph outlines) grouped per formula
- **Heuristic math mode**: Auto-classifies formulas as simple (text) or complex (glyph curves) for best rendering; stacked fractions (e.g. `$1/2$`) are auto-detected and routed to glyph mode
- **Native shapes**: Converts SVG shapes (rectangles, circles, ellipses, lines, paths, polygons) to native DrawingML custom geometry
- **Gradient fills**: Supports SVG linear gradients converted to PowerPoint gradient fills
- **Transparency/opacity**: Full support for alpha channels in colors (#RRGGBBAA, rgba), fill-opacity, and element opacity
- **Hyperlinks**: External links from Typst `#link()` are preserved as clickable hyperlinks in PowerPoint; internal document links are rendered as normal text without hyperlink styling; theme-level hyperlink colors are neutralized to prevent default blue/purple override
- **Speaker notes**: Extracts Touying speaker notes via `typst.query()` (Python package) and attaches them to slides
- **Color preservation**: Supports hex, rgb(), rgba(), named CSS colors, and per-character coloring
- **Multi-run textboxes**: Groups same-line text segments into single textboxes with multiple styled runs
- **Paragraph alignment**: Auto-detects left, center, right, and justify alignment from Typst layout
- **Code blocks**: Syntax-highlighted code blocks rendered with Consolas font, preserving per-token colors
- **Images**: Supports embedded (data URI) and external image references (PNG, SVG, PDF); SVG/PDF images rasterized to PNG with transparent background via the `typst` Python package
- **Tables**: Table content (cells, headers, colored fills) converted to text and shapes

## Prerequisites

- **Python 3.9+**
- **typst** (Python package, v0.14+): For speaker notes extraction and SVG/PDF image rasterization
  - Installed automatically via pip
- **typst-ts-cli** (v0.6.0+): For compiling `.typ` files to SVG with foreignObject text overlays
  - Download from [typst.ts releases](https://github.com/Myriad-Dreamin/typst.ts/releases) automatically

## Installation

### From PyPI (recommended)

```bash
pip install typ2pptx
```

After installation, the `typ2pptx` CLI command is available system-wide.

## Usage

### Command Line

```bash
# Convert a Typst presentation to PPTX
typ2pptx slides.typ -o slides.pptx

# Convert with verbose output
typ2pptx slides.typ -o slides.pptx -v

# Convert from pre-compiled SVG (no typst-ts-cli needed)
typ2pptx slides.artifact.svg -o slides.pptx

# Specify project root directory (for resolving imports/paths)
typ2pptx slides/main.typ -o slides.pptx --root .

# Specify custom tool paths
typ2pptx slides.typ -o slides.pptx \
    --typst-ts-cli /path/to/typst-ts-cli

# Math rendering mode options
typ2pptx slides.typ -o slides.pptx \
    --inline-math-mode auto \      # "text", "glyph", or "auto" (default: auto)
    --display-math-mode glyph      # "text", "glyph", or "auto" (default: glyph)
```

### Math Rendering Modes

| Mode | Inline Math Default | Display Math Default | Description |
|------|--------------------|--------------------|-------------|
| `text` | Cambria Math font | Cambria Math font | Renders math as editable text runs |
| `glyph` | Glyph curves | Glyph curves | Renders math as native DrawingML shapes (pixel-perfect) |
| `auto` | Heuristic | Heuristic | Simple formulas as text, complex ones (integrals, matrices, stacked fractions) as glyph curves |

### Python API

```python
from typ2pptx.core.converter import convert_typst_to_pptx

# Simple conversion
convert_typst_to_pptx("slides.typ", "slides.pptx")

# With options
convert_typst_to_pptx(
    "slides.typ",
    "slides.pptx",
    typst_ts_cli="/path/to/typst-ts-cli",
    root="/path/to/project/root",
    verbose=True,
)

# Convert from SVG directly
convert_typst_to_pptx("slides.artifact.svg", "slides.pptx")
```

### Advanced: Step-by-step conversion

```python
from typ2pptx.core.converter import (
    compile_typst_to_svg,
    query_speaker_notes,
    TypstSVGConverter,
    ConversionConfig,
)

# Step 1: Compile to SVG
svg_path = compile_typst_to_svg("slides.typ")

# Step 2: Extract speaker notes
notes = query_speaker_notes("slides.typ")

# Step 3: Convert to PPTX with custom config
config = ConversionConfig(
    verbose=True,
    inline_math_mode="auto",     # "text", "glyph", or "auto"
    display_math_mode="glyph",   # "text", "glyph", or "auto"
)
converter = TypstSVGConverter(config)
converter.convert(svg_path, "slides.pptx", speaker_notes=notes)
```

## How It Works

### Pipeline

```
Typst (.typ)
    |
    v
typst-ts-cli  -->  SVG (with foreignObject text overlays)
    |
    v
typst.query() -->  Speaker notes (pdfpc JSON format, via Python package)
    |
    v
typ2pptx    -->  PowerPoint (.pptx)
```

### Architecture

1. **SVG Parser** (`typst_svg_parser.py`): Parses the typst.ts SVG structure, which uses:
   - Glyph outlines as `<path>` + `<use>` references (for rendering)
   - `<foreignObject>` overlays with HTML text (for selection/copy)
   - `scale(S, -S)` transforms with Y-axis flipping
   - 5-character hash prefixes per font variant
   - `<a>` hyperlink elements with `<rect>` bounding boxes for link regions

2. **Font Variant Detection**: Identifies font styles by analyzing glyph paths:
   - Quadratic curves (Q commands) indicate monospace fonts
   - Glyph width comparison differentiates bold from italic
   - Unicode math character detection (U+1D400+) identifies math fonts
   - Usage frequency determines the regular (body text) font

3. **Shape Converter**: Uses the ppt-master `svg_to_shapes.py` pipeline:
   - `parse_svg_path()` -> `svg_path_to_absolute()` -> `normalize_path_commands()` -> `path_commands_to_drawingml()`
   - Generates DrawingML `<a:custGeom>` XML for custom geometry
   - Supports solid fills (with alpha transparency), gradient fills, and strokes

4. **Text Converter**: Groups and renders text:
   - Groups same-line segments (within 2px baseline tolerance) into multi-run textboxes
   - Handles gap-based space detection for word boundaries
   - Inline math sub/superscripts merged into adjacent text lines with PPTX-native baseline offsets
   - Math segments clustered spatially into formula regions
   - Sets font properties (bold, italic, name, size, color) per run
   - Auto-detects paragraph alignment (left, center, right, justify) from segment positions

5. **Math Renderer**: Dual-mode math formula handling:
   - **Text mode**: Renders as Cambria Math text runs with sub/superscript baseline offsets
   - **Graphics mode**: Renders glyph outlines as DrawingML `<a:custGeom>` shapes grouped in `<p:grpSp>`
   - **Auto mode**: Heuristic classification - simple formulas (letters, digits, basic operators) as text, complex formulas (integrals, matrices, roots, stacked fractions) as glyph curves

6. **Link Processor**: Detects SVG `<a>` elements and applies hyperlinks to overlapping text runs in PPTX
   - Only external links (http, https, mailto) become PPTX hyperlinks
   - Internal document links are rendered as normal text
   - Explicitly suppresses PowerPoint's default hyperlink styling (blue color + underline) via `u="none"` and explicit `solidFill`
   - Theme-level hyperlink colors (hlink/folHlink) are neutralized to prevent blue/purple override

7. **Speaker Notes**: Extracts via `typst.query()` (Python package):
   - Uses `typst.query(path, "<pdfpc-file>", field="value")` to extract pdfpc JSON
   - Touying framework outputs pdfpc JSON with page indices and note text
   - Notes are attached to the corresponding PowerPoint slides
   - Falls back to CLI `typst query` if the Python package is unavailable

## Supported Typst Content

| Category | Feature | Support |
|----------|---------|---------|
| **Text** | Regular, Bold, Italic, Bold-Italic | Full |
| **Text** | Monospace / inline code | Full |
| **Text** | Colored text | Full |
| **Text** | Font size variants | Full |
| **Text** | Chinese/CJK text | Full |
| **Math** | Inline formulas | Cambria Math text (auto/text mode) or glyph curves (glyph mode) |
| **Math** | Display formulas | Glyph curves (default) or Cambria Math text |
| **Math** | Stacked fractions | Auto-detected and rendered as glyph curves in auto mode |
| **Math** | Subscripts/superscripts | PPTX-native baseline offset (auto-sized by PowerPoint) |
| **Math** | Formula grouping | Grouped as single draggable unit in PPTX |
| **Shapes** | Rectangles, circles, ellipses | Native DrawingML |
| **Shapes** | Lines, paths, polygons | Native DrawingML |
| **Shapes** | Gradient fills | DrawingML gradFill |
| **Shapes** | Transparency/opacity | Alpha channel support (RRGGBBAA, rgba, fill-opacity) |
| **Code** | Inline code | Consolas font |
| **Code** | Code blocks | Consolas font with syntax highlighting colors |
| **Images** | Embedded PNG (data URI) | Full |
| **Images** | Embedded SVG | Rasterized to PNG via typst (transparent background) |
| **Images** | Embedded/external PDF | Rasterized to PNG via typst (transparent background) |
| **Images** | External references | Full |
| **Links** | External hyperlinks (`#link()`) | Clickable in PPTX, preserves original styling |
| **Links** | Internal document links | Rendered as normal text (no hyperlink) |
| **Tables** | Table cells and headers | Text + shapes |
| **Tables** | Colored table backgrounds | Filled shapes |
| **Layout** | Slide dimensions (16:9, 4:3) | Full |
| **Layout** | Text alignment (left, center, right, justify) | Auto-detected |
| **Notes** | Touying speaker notes | Full |
| **Lists** | Bullet points | Text with bullet chars |
| **Plugins** | Pinit highlights | Colored overlay shapes |

## Development

### Install in editable/development mode

```bash
# pip
pip install -e .
pip install -e ".[dev]"    # with dev dependencies (pytest, etc.)

# uv
uv pip install -e .
uv pip install -e ".[dev]"
```

### Updating dependencies

All dependencies are declared in `pyproject.toml`. After modifying them:

```bash
# pip
pip install -e .

# uv
uv pip install -e .
```

## Testing

```bash
# Run all tests
pytest tests/ -v

# Run specific test modules
pytest tests/test_converter.py -v
pytest tests/test_math.py -v
pytest tests/test_inline_math.py -v
pytest tests/test_table.py -v
pytest tests/test_pinit.py -v
pytest tests/test_links.py -v
pytest tests/test_images.py -v
pytest tests/test_chinese.py -v
pytest tests/test_alignment.py -v
pytest tests/test_code_blocks.py -v
```

### Test Coverage

- **254 tests** covering:
  - Converter & configuration (56 tests)
  - SVG parser (30 tests)
  - Inline math & stacked fractions (22 tests)
  - Image embedding & rasterization (21 tests)
  - SVG path pipeline (17 tests)
  - Table rendering (17 tests)
  - Math formulas & display curves (16 tests)
  - Pinit annotations (15 tests)
  - Code blocks & syntax highlighting (12 tests)
  - Text positioning & multi-run merging (12 tests)
  - Chinese text (11 tests)
  - Hyperlinks (11 tests)
  - Columns layout (7 tests)
  - Text alignment (7 tests)

## Project Structure

```
typ2pptx/
  core/
    converter.py          # Main SVG -> PPTX conversion logic
    typst_svg_parser.py   # typst.ts SVG parsing and text extraction
  scripts/
    svg_to_shapes.py      # SVG path -> DrawingML pipeline (from ppt-master)
  data/
    bin/                  # Bundled typst-ts-cli binary (platform-specific)
  __main__.py             # CLI entry point
scripts/
  download_typst_ts_cli.py  # Download typst-ts-cli for bundling
tests/
  conftest.py             # Shared test fixtures
  test_converter.py       # Converter tests
  test_math.py            # Math formula tests
  test_inline_math.py     # Inline math and display math curve tests
  test_chinese.py         # Chinese text tests
  test_table.py           # Table rendering tests
  test_pinit.py           # Pinit annotation tests
  test_links.py           # Hyperlink tests
  test_images.py          # Image embedding tests
  test_alignment.py       # Text alignment tests
  test_code_blocks.py     # Code block & syntax highlighting tests
  test_text_positioning.py # Text positioning tests
  test_svg_parser.py      # SVG parser tests
  test_path_pipeline.py   # Path conversion tests
  typ_sources/            # Test Typst source files
    basic_text.typ
    shapes_test.typ
    speaker_notes_test.typ
    math_test.typ
    chinese_test.typ
    inline_math_test.typ
    table_test.typ
    pinit_test.typ
    link_test.typ
    image_test.typ
    alignment_test.typ
    code_block_test.typ
    columns_test.typ
```

## Acknowledgments

### ppt-master Attribution

The file `typ2pptx/scripts/svg_to_shapes.py` is adapted from the [ppt-master](https://github.com/niccolocorsani/ppt-master) project, which provides the core SVG path to DrawingML conversion pipeline:

- `parse_svg_path()` -- tokenizes SVG path `d` attributes into structured commands
- `svg_path_to_absolute()` -- converts relative path commands to absolute coordinates
- `normalize_path_commands()` -- reduces all curve types (S, Q, T, A) to cubic beziers (C)
- `path_commands_to_drawingml()` -- generates DrawingML `<a:custGeom>` XML for custom geometry

The pipeline supports solid fills, gradient fills (linear and radial), and strokes with configurable dash patterns and line caps.

The original ppt-master code has been modified for typst SVG compatibility, including:

- Added inherited style propagation through nested `<g>` elements via `ConvertContext`
- Added CJK font detection and Windows font fallback mapping for cross-platform PPTX compatibility
- Added donut-chart arc segment detection (SVG `stroke-dasharray` circles converted to filled annular sectors)
- Added support for element opacity, fill-opacity, and stroke-opacity (including multiplicative inheritance)
- Added `<ellipse>`, `<polyline>`, `<image>`, and `<text>` (with multi-run `<tspan>`) element converters
- Added group (`<g>`) to `<p:grpSp>` conversion with automatic bounds calculation
- Added shadow effect support via SVG `<filter>` (feGaussianBlur + feOffset)

## License

MIT
