"""svg_to_pptx — SVG to DrawingML conversion package.

Adapted from ppt-master (https://github.com/hugohe3/ppt-master).
Only the core conversion modules are included here; CLI, builder, and
media-rendering utilities are not needed for typ2pptx.

Public API:
    - convert_svg_to_slide_shapes(): SVG -> DrawingML slide XML
    - parse_svg_path / svg_path_to_absolute / normalize_path_commands /
      path_commands_to_drawingml: SVG path pipeline
    - PathCommand: dataclass for parsed path commands
    - ConvertContext: shared conversion state
"""

from .drawingml_converter import convert_svg_to_slide_shapes, convert_element, collect_defs
from .drawingml_paths import (
    PathCommand,
    parse_svg_path,
    svg_path_to_absolute,
    normalize_path_commands,
    path_commands_to_drawingml,
)
from .drawingml_context import ConvertContext
from .drawingml_styles import (
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
)
from .drawingml_utils import px_to_emu, parse_hex_color, EMU_PER_PX, ANGLE_UNIT

__all__ = [
    'convert_svg_to_slide_shapes',
    'convert_element',
    'collect_defs',
    'PathCommand',
    'parse_svg_path',
    'svg_path_to_absolute',
    'normalize_path_commands',
    'path_commands_to_drawingml',
    'ConvertContext',
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
