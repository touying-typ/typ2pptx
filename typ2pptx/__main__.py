"""
typ2pptx CLI - Convert Typst presentations to PowerPoint.

Usage:
    typ2pptx slides.typ -o slides.pptx
    typ2pptx slides.svg -o slides.pptx
    typ2pptx slides.typ --typst-ts-cli /path/to/typst-ts-cli
"""

import argparse
import sys
from pathlib import Path

from .core.converter import convert_typst_to_pptx, ConversionConfig


def main():
    parser = argparse.ArgumentParser(
        prog='typ2pptx',
        description='Convert Typst presentations to PowerPoint (.pptx) files',
    )

    parser.add_argument(
        'input',
        help='Input file (.typ or .svg)',
    )
    parser.add_argument(
        '-o', '--output',
        help='Output PPTX file path (default: same name as input)',
    )
    parser.add_argument(
        '--typst-ts-cli',
        default=None,
        help='Path to typst-ts-cli binary (default: bundled or system)',
    )
    parser.add_argument(
        '--root',
        default=None,
        help='Root directory for the Typst project (for resolving imports/paths)',
    )
    parser.add_argument(
        '-v', '--verbose',
        action='store_true',
        help='Enable verbose output',
    )
    parser.add_argument(
        '--raster-dpi',
        type=int,
        default=300,
        help='DPI for rasterization (default: 300)',
    )
    parser.add_argument(
        '--inline-math-mode',
        choices=['text', 'glyph', 'auto'],
        default='auto',
        help='Inline math rendering: "text" (Cambria Math), "glyph" (glyph curves), or "auto" (heuristic). Default: auto',
    )
    parser.add_argument(
        '--display-math-mode',
        choices=['text', 'glyph', 'auto'],
        default='glyph',
        help='Display/block math rendering: "text" (Cambria Math), "glyph" (glyph curves), or "auto" (heuristic). Default: glyph',
    )
    parser.add_argument(
        '--font-size-scale',
        type=float,
        default=0.75,
        help='Points per SVG font-size unit (default: 0.75 for CSS px at 96 DPI; '
             'use 1.0 when 1 SVG unit = 1pt, e.g. typst.ts SVG)',
    )
    parser.add_argument(
        '--emu-per-px',
        type=float,
        default=None,
        help='EMU per SVG unit for geometry (default: 12700 x font-size-scale, '
             'so geometry and font sizes share one physical scale)',
    )
    parser.add_argument(
        '--detect-paragraphs',
        action='store_true',
        default=False,
        help=(
            'Enable the paragraph auto-detection heuristic that merges '
            'consecutive wrapped lines (same font, aligned left edge, full-width) '
            'into a single word-wrapped textbox. Disabled by default because the '
            'heuristic can mis-merge tightly-packed content such as tables and '
            'list items. Enable it for prose-heavy decks.'
        ),
    )
    parser.add_argument(
        '--latin-font',
        default='Arial',
        help=(
            'Real font family name to use for regular/bold/italic text '
            '(default: Arial). The typst.ts SVG artifact carries no '
            'font-family metadata at all, so the converter cannot recover '
            'the document\'s real display/body font from the SVG itself -- '
            'pass the actual family (e.g. the brand\'s fonts.display/'
            'fonts.body value) here to avoid a silent Arial substitution.'
        ),
    )
    parser.add_argument(
        '--mono-font',
        default='Consolas',
        help=(
            'Real font family name to use for text detected as monospace '
            '(default: Consolas). Same rationale as --latin-font.'
        ),
    )
    parser.add_argument(
        '--display-font',
        default=None,
        help=(
            'Real font family name for text detected as display/headline '
            'size (a non-mono prefix rendering well larger than the '
            'document\'s body text, e.g. 1.4x+). Headline and body copy are '
            'often genuinely different families (a brand\'s fonts.display '
            'vs fonts.body), not just different weights of one family -- '
            'pass this to avoid a too-wide/wrong-family substitution on '
            'title slides. Default: falls back to --latin-font (previous '
            'behavior, unchanged if you don\'t pass this).'
        ),
    )

    args = parser.parse_args()

    # Validate input
    input_path = Path(args.input)
    if not input_path.exists():
        print(f"Error: Input file not found: {input_path}", file=sys.stderr)
        sys.exit(1)

    if input_path.suffix not in ('.typ', '.svg'):
        print(f"Error: Unsupported input format: {input_path.suffix}", file=sys.stderr)
        print("Supported formats: .typ, .svg", file=sys.stderr)
        sys.exit(1)

    # Configuration
    config = ConversionConfig(
        raster_dpi=args.raster_dpi,
        inline_math_mode=args.inline_math_mode,
        display_math_mode=args.display_math_mode,
        detect_paragraphs=args.detect_paragraphs,
        font_size_scale=args.font_size_scale,
        emu_per_px=args.emu_per_px,
        default_latin_font=args.latin_font,
        default_mono_font=args.mono_font,
        default_display_font=args.display_font,
    )

    try:
        output = convert_typst_to_pptx(
            str(input_path),
            output_path=args.output,
            typst_ts_cli=args.typst_ts_cli,
            root=args.root,
            config=config,
            verbose=args.verbose,
        )
        print(f"Successfully created: {output}")
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        if args.verbose:
            import traceback
            traceback.print_exc()
        sys.exit(1)


if __name__ == '__main__':
    main()
