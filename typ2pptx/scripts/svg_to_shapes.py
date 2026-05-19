"""
SVG to DrawingML Native Shapes Converter — compatibility re-export layer.

This module re-exports all public symbols from the svg_to_pptx package
(adapted from ppt-master) so existing imports continue to work unchanged.

Usage:
    from typ2pptx.scripts.svg_to_shapes import (
        parse_svg_path, svg_path_to_absolute,
        normalize_path_commands, path_commands_to_drawingml,
        PathCommand,
    )
"""

# Re-export the full public API from the modular svg_to_pptx package.
from .svg_to_pptx import (  # noqa: F401
    # Path pipeline (used by converter.py)
    PathCommand,
    parse_svg_path,
    svg_path_to_absolute,
    normalize_path_commands,
    path_commands_to_drawingml,
    # Context
    ConvertContext,
    # Top-level conversion
    convert_svg_to_slide_shapes,
    convert_element,
    collect_defs,
    # Style builders
    build_fill_xml,
    build_stroke_xml,
    build_effect_xml,
    build_shadow_xml,
    build_glow_xml,
    build_solid_fill,
    build_gradient_fill,
    build_pattern_fill,
    get_fill_opacity,
    get_stroke_opacity,
    # Utilities
    px_to_emu,
    parse_hex_color,
    EMU_PER_PX,
    ANGLE_UNIT,
)

__all__ = [
    'PathCommand',
    'parse_svg_path',
    'svg_path_to_absolute',
    'normalize_path_commands',
    'path_commands_to_drawingml',
    'ConvertContext',
    'convert_svg_to_slide_shapes',
    'convert_element',
    'collect_defs',
    'build_fill_xml',
    'build_stroke_xml',
    'build_effect_xml',
    'build_shadow_xml',
    'build_glow_xml',
    'build_solid_fill',
    'build_gradient_fill',
    'build_pattern_fill',
    'get_fill_opacity',
    'get_stroke_opacity',
    'px_to_emu',
    'parse_hex_color',
    'EMU_PER_PX',
    'ANGLE_UNIT',
]
