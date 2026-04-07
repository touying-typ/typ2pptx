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
