"""
Typst SVG Parser - Parses typst.ts generated SVG files.

typst.ts generates SVG with a special structure:
- No <text> elements; all text is rendered via <path> glyph outlines + <use> references
- Text content is in <foreignObject> overlay layers (for selection/copy)
- Each font variant has a unique 5-char ID prefix (hash-based, changes each export)
- Coordinates use scale(S, -S) with Y-axis flipping
- Pages are vertically stacked with translate(0, Y) offsets

This module extracts structured page data from typst.ts SVGs.
"""

import re
import math
import copy
from typing import Optional, Tuple, List, Dict, Any, Set
from xml.etree import ElementTree as ET
from dataclasses import dataclass, field

# Namespaces used in typst.ts SVG
SVG_NS = 'http://www.w3.org/2000/svg'
XLINK_NS = 'http://www.w3.org/1999/xlink'
XHTML_NS = 'http://www.w3.org/1999/xhtml'
H5_NS = 'http://www.w3.org/1999/xhtml'  # typst.ts uses h5: prefix for xhtml

# Register namespaces for clean output
ET.register_namespace('', SVG_NS)
ET.register_namespace('xlink', XLINK_NS)
ET.register_namespace('h5', H5_NS)


@dataclass
class GlyphInfo:
    """Information about a glyph definition in the SVG."""
    glyph_id: str  # Full ID like "gVYuD1jU"
    prefix: str  # 5-char prefix like "gVYuD"
    path_data: str  # The d attribute of the path
    uses_quadratic: bool  # Whether path uses Q commands (mono/math fonts)


@dataclass
class FontVariant:
    """A detected font variant based on glyph prefix analysis."""
    prefix: str  # 5-char prefix
    glyph_count: int  # Number of glyphs with this prefix
    uses_quadratic: bool  # True if glyphs use Q commands
    style: str = "regular"  # regular, bold, italic, bolditalic, mono, math


@dataclass
class GlyphUse:
    """A single glyph usage (a <use> reference to a glyph path in defs)."""
    glyph_id: str  # ID of the referenced glyph path (e.g., "g6wLp1awX")
    x_offset: float  # x offset within the typst-text group (from <use x="...">)
    y_offset: float  # y offset within the typst-text group (from <use y="...">)
    path_data: str  # The SVG path d attribute of the glyph


@dataclass
class TextSegment:
    """A text segment extracted from foreignObject."""
    text: str
    x: float  # Position in SVG coordinates (after all transforms)
    y: float
    width: float
    height: float
    font_size: float  # Computed font size in SVG px
    font_variant: str  # regular, bold, italic, bolditalic, mono, math
    fill_color: str  # Hex color like "#000000"
    # The class name from the div element
    css_class: str = ""
    # Glyph usage info for rendering as curves (especially for math)
    glyph_uses: List['GlyphUse'] = field(default_factory=list)
    # The glyph scale factor (from the typst-text group's scale transform)
    glyph_scale: float = 0.0


@dataclass
class ShapeElement:
    """A shape element (rect, circle, line, path, image, etc.)."""
    tag: str  # rect, circle, ellipse, line, path, polygon, image, etc.
    element: ET.Element  # The original SVG element
    x: float = 0.0
    y: float = 0.0
    width: float = 0.0
    height: float = 0.0
    transform_matrix: List[float] = field(default_factory=lambda: [1, 0, 0, 1, 0, 0])
    is_glyph_path: bool = False  # True if this is a glyph outline (should be skipped)


@dataclass
class LinkRegion:
    """A hyperlink region from SVG <a> elements."""
    href: str  # URL target
    x: float  # Bounding box in SVG coordinates
    y: float
    width: float
    height: float


@dataclass
class PageData:
    """Data for a single slide page."""
    page_num: int
    y_offset: float  # Vertical offset in the combined SVG
    width: float  # Page width in SVG pixels
    height: float  # Page height in SVG pixels
    text_segments: List[TextSegment] = field(default_factory=list)
    shapes: List[ShapeElement] = field(default_factory=list)
    links: List[LinkRegion] = field(default_factory=list)
    # Raw elements for fallback processing
    raw_elements: List[ET.Element] = field(default_factory=list)


@dataclass
class TypstSVGData:
    """Complete parsed data from a typst.ts SVG file."""
    viewbox_width: float
    viewbox_height: float
    pages: List[PageData] = field(default_factory=list)
    glyph_defs: Dict[str, GlyphInfo] = field(default_factory=dict)
    font_variants: Dict[str, FontVariant] = field(default_factory=dict)
    css_styles: str = ""  # Raw CSS from <style> block
    # All defs elements (gradients, clipPaths, etc.)
    defs: Dict[str, ET.Element] = field(default_factory=dict)


def parse_transform(transform_str: Optional[str]) -> Tuple[float, float, float, float, float]:
    """Parse SVG transform string into (dx, dy, sx, sy, rotation_deg).

    Handles:
    - translate(dx, dy) or translate(dx)
    - scale(sx, sy) or scale(s)
    - rotate(deg)
    - matrix(a, b, c, d, e, f)

    Returns (dx, dy, sx, sy, rotation_deg)
    """
    if not transform_str:
        return (0.0, 0.0, 1.0, 1.0, 0.0)

    dx, dy = 0.0, 0.0
    sx, sy = 1.0, 1.0
    rotation = 0.0

    # Parse translate
    for m in re.finditer(r'translate\(\s*([-\d.e+]+)(?:[,\s]+([-\d.e+]+))?\s*\)', transform_str):
        dx += float(m.group(1))
        dy += float(m.group(2)) if m.group(2) else 0.0

    # Parse scale
    for m in re.finditer(r'scale\(\s*([-\d.e+]+)(?:[,\s]+([-\d.e+]+))?\s*\)', transform_str):
        sx_val = float(m.group(1))
        sy_val = float(m.group(2)) if m.group(2) else sx_val
        sx *= sx_val
        sy *= sy_val

    # Parse rotate
    for m in re.finditer(r'rotate\(\s*([-\d.e+]+)\s*\)', transform_str):
        rotation += float(m.group(1))

    # Parse matrix(a, b, c, d, e, f)
    for m in re.finditer(r'matrix\(\s*([-\d.e+]+)[,\s]+([-\d.e+]+)[,\s]+([-\d.e+]+)[,\s]+([-\d.e+]+)[,\s]+([-\d.e+]+)[,\s]+([-\d.e+]+)\s*\)', transform_str):
        a, b, c, d, e, f = [float(m.group(i+1)) for i in range(6)]
        # Extract translation from matrix
        dx += e
        dy += f
        # Extract scale (simplified: assumes no skew for now)
        det_sx = math.sqrt(a*a + b*b)
        det_sy = math.sqrt(c*c + d*d)
        if det_sx > 0:
            sx *= det_sx
        if det_sy > 0:
            sy *= det_sy
        # Extract rotation
        if a != 0 or b != 0:
            rotation += math.degrees(math.atan2(b, a))

    return (dx, dy, sx, sy, rotation)


def parse_viewbox(viewbox_str: str) -> Tuple[float, float, float, float]:
    """Parse SVG viewBox attribute: 'minX minY width height'."""
    parts = viewbox_str.strip().split()
    if len(parts) == 4:
        return tuple(float(p) for p in parts)
    return (0, 0, 0, 0)


def _collect_defs(root: ET.Element) -> Dict[str, ET.Element]:
    """Collect all elements in <defs> blocks, keyed by their 'id' attribute."""
    defs = {}
    for defs_elem in root.iter(f'{{{SVG_NS}}}defs'):
        for child in defs_elem:
            elem_id = child.get('id')
            if elem_id:
                defs[elem_id] = child
    # Also check for elements with id not inside defs (some typst versions)
    for child in root.iter():
        elem_id = child.get('id')
        if elem_id and elem_id not in defs:
            # Only add if it's a reusable element (not a page group)
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if tag in ('path', 'symbol', 'clipPath', 'linearGradient', 'radialGradient', 'pattern', 'mask', 'filter'):
                defs[elem_id] = child
    return defs


def _analyze_glyphs(defs: Dict[str, ET.Element]) -> Tuple[Dict[str, GlyphInfo], Dict[str, FontVariant]]:
    """Analyze glyph definitions to detect font variants.

    typst.ts glyph IDs follow the pattern: g{5-char-prefix}{rest}
    where the 5-char prefix identifies the font variant.

    Key heuristics:
    - Quadratic curves (Q commands) → mono or math font
    - Most glyphs → regular font
    - Needs context from <use> elements to determine bold/italic
    """
    glyph_defs = {}
    prefix_stats: Dict[str, Dict[str, Any]] = {}

    for elem_id, elem in defs.items():
        tag = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag
        if tag != 'path':
            continue

        # Check if this is a glyph outline (class="outline_glyph" or id starts with 'g')
        css_class = elem.get('class', '')
        if 'outline_glyph' not in css_class and not elem_id.startswith('g'):
            continue

        path_data = elem.get('d', '')
        if not path_data:
            continue

        # Extract prefix (first character 'g' + 4 more chars)
        if len(elem_id) >= 6 and elem_id.startswith('g'):
            prefix = elem_id[:5]  # e.g., "gVYuD"
        else:
            prefix = elem_id[:5] if len(elem_id) >= 5 else elem_id

        # Check for quadratic bezier commands
        uses_quadratic = bool(re.search(r'[Qq]', path_data))

        glyph_info = GlyphInfo(
            glyph_id=elem_id,
            prefix=prefix,
            path_data=path_data,
            uses_quadratic=uses_quadratic,
        )
        glyph_defs[elem_id] = glyph_info

        # Accumulate stats per prefix
        if prefix not in prefix_stats:
            prefix_stats[prefix] = {
                'count': 0,
                'quadratic_count': 0,
            }
        prefix_stats[prefix]['count'] += 1
        if uses_quadratic:
            prefix_stats[prefix]['quadratic_count'] += 1

    # Build font variants
    font_variants = {}
    for prefix, stats in prefix_stats.items():
        quadratic_ratio = stats['quadratic_count'] / max(stats['count'], 1)

        # If most glyphs use quadratic curves, it's likely mono or math
        if quadratic_ratio > 0.5:
            style = "mono"  # Will be refined later with context
        else:
            style = "regular"  # Will be refined later

        font_variants[prefix] = FontVariant(
            prefix=prefix,
            glyph_count=stats['count'],
            uses_quadratic=(quadratic_ratio > 0.5),
            style=style,
        )

    return glyph_defs, font_variants


def _get_glyph_width(path_data: str) -> float:
    """Get the x-extent (width) of a glyph from its path data.

    Extracts all numeric values and estimates the x-range.
    """
    nums = re.findall(r'([-\d.]+)', path_data)
    if not nums:
        return 0
    floats = [float(n) for n in nums]
    # In SVG path data, x and y coordinates alternate (roughly)
    xs = floats[0::2]
    if not xs:
        return 0
    return max(xs) - min(xs)


def _prescan_font_variants(
    root: ET.Element,
    glyph_defs: Dict[str, GlyphInfo],
    font_variants: Dict[str, FontVariant],
) -> Dict[str, str]:
    """Pre-scan the SVG to map glyph prefixes to font styles.

    This does a two-pass analysis:
    1. First pass: Map each typst-text group to its text content and glyph prefix
    2. Second pass: Use heuristics to assign bold/italic/mono/math styles

    Heuristics:
    - Most used prefix (non-quadratic) → regular
    - Quadratic curve prefix → mono
    - Prefix used only in groups with larger scale (0.03 vs 0.025) → could be bold (heading)
    - Prefix used together with regular text on same line at same scale → bold or italic
    - We can also check if the glyph paths are thicker (bolder) by path data analysis

    Returns dict mapping prefix → style
    """
    # Collect prefix usage data
    prefix_texts: Dict[str, List[str]] = {}  # prefix -> list of text content
    prefix_scales: Dict[str, Set[float]] = {}  # prefix -> set of scales used
    prefix_usage_count: Dict[str, int] = {}  # prefix -> total use count

    for text_group in root.iter(f'{{{SVG_NS}}}g'):
        if 'typst-text' not in text_group.get('class', ''):
            continue

        # Get text content
        text = ''
        for fo in text_group.iter(f'{{{SVG_NS}}}foreignObject'):
            for div in fo:
                text = (div.text or '').strip()
                break

        # Get scale from transform
        transform = text_group.get('transform', '')
        scale_match = re.search(r'scale\(([-\d.e+]+)', transform)
        scale = abs(float(scale_match.group(1))) if scale_match else 0

        # Get glyph prefixes used
        prefixes = set()
        for use in text_group.findall(f'{{{SVG_NS}}}use'):
            href = use.get(f'{{{XLINK_NS}}}href') or use.get('href', '')
            if href.startswith('#'):
                gid = href[1:]
                if gid in glyph_defs:
                    prefixes.add(glyph_defs[gid].prefix)

        for prefix in prefixes:
            if prefix not in prefix_texts:
                prefix_texts[prefix] = []
                prefix_scales[prefix] = set()
                prefix_usage_count[prefix] = 0
            if text:
                prefix_texts[prefix].append(text)
            prefix_scales[prefix].add(scale)
            prefix_usage_count[prefix] += 1

    # Now assign styles
    prefix_to_style: Dict[str, str] = {}

    # Step 1: Mark quadratic → mono
    for prefix, variant in font_variants.items():
        if variant.uses_quadratic:
            prefix_to_style[prefix] = 'mono'

    # Step 1.5: Detect math font by checking if text content contains
    # Unicode mathematical characters (U+1D400-U+1D7FF Mathematical Alphanumeric Symbols,
    # or common math operators like ∫∑∏√∞±≠≤≥ etc.)
    def _is_math_text(text: str) -> bool:
        """Check if text contains Unicode math symbols."""
        for ch in text:
            cp = ord(ch)
            # Mathematical Alphanumeric Symbols (U+1D400–U+1D7FF)
            if 0x1D400 <= cp <= 0x1D7FF:
                return True
            # Miscellaneous Mathematical Symbols-A (U+27C0–U+27EF)
            if 0x27C0 <= cp <= 0x27EF:
                return True
            # Miscellaneous Mathematical Symbols-B (U+2980–U+29FF)
            if 0x2980 <= cp <= 0x29FF:
                return True
            # Supplemental Mathematical Operators (U+2A00–U+2AFF)
            if 0x2A00 <= cp <= 0x2AFF:
                return True
            # Common math symbols: ∫∑∏√∞±∓≠≤≥≈∝∂∇∅∀∃∈∉⊂⊃∪∩
            if ch in '∫∑∏√∞±∓≠≤≥≈∝∂∇∅∀∃∈∉⊂⊃∪∩∧∨⊕⊗⊥∥⟨⟩':
                return True
        return False

    for prefix in list(font_variants.keys()):
        if prefix in prefix_to_style:
            continue
        texts = prefix_texts.get(prefix, [])
        math_count = sum(1 for t in texts if _is_math_text(t))
        total_count = len(texts) if texts else 0
        # If majority of text segments for this prefix contain math chars, it's math
        if total_count > 0 and math_count / total_count > 0.3:
            prefix_to_style[prefix] = 'math'

    # Step 2: Find the regular prefix among non-mono, non-math prefixes.
    # Key heuristic: the regular (body text) font typically has:
    #  - The most total usage count (headings are fewer but may have longer text)
    #  - Text at the standard body scale (0.025 in typst.ts)
    #  - The most total text content (by character count across all segments)
    remaining = [(p, prefix_usage_count.get(p, font_variants[p].glyph_count))
                 for p in font_variants
                 if p not in prefix_to_style]

    if remaining:
        # Score each prefix: prefer highest usage count, with tie-breaking
        # by total text content length.
        def _total_text_len(prefix):
            texts = prefix_texts.get(prefix, [])
            return sum(len(t) for t in texts)

        # Use a composite score: usage_count * 10 + total_text_len
        # This prioritizes usage count (body text appears on many slides)
        # but uses total text length as tiebreaker
        remaining.sort(key=lambda x: (x[1] * 10 + _total_text_len(x[0])), reverse=True)
        regular_prefix = remaining[0][0]
        prefix_to_style[regular_prefix] = 'regular'
        remaining = remaining[1:]

    # Step 3: For remaining prefixes, use path analysis heuristic
    # Bold fonts tend to have wider strokes → larger bounding boxes relative to em-square
    # We can also check the number of path segments (bold often has more complex outlines)

    # But a simpler approach: compare the same character across prefixes
    # If a glyph with the same suffix has more path data in one prefix, it's likely bolder

    # For now, use a combination of heuristics:
    # - If a prefix is used at the same scale as regular text AND its glyph paths have
    #   more complex/thicker data → bold
    # - If prefix glyphs tend to have slanted paths (check if paths skew right) → italic

    for prefix, _ in remaining:
        if prefix in prefix_to_style:
            continue

        # Get sample glyph paths for this prefix
        sample_glyphs = [g for g in glyph_defs.values() if g.prefix == prefix][:5]

        if not sample_glyphs:
            prefix_to_style[prefix] = 'regular'
            continue

        # Check for italic: italic glyphs tend to have a specific pattern
        # One heuristic: look at the average x-displacement in paths
        # Italic fonts have glyphs that lean to the right

        # Another approach: check path bounding box aspect ratio
        # Bold fonts have wider strokes
        # But the simplest reliable approach for typst:

        # typst uses specific font variants:
        # - Regular uses most glyphs
        # - Bold uses fewer unique glyphs (same letters, different outlines)
        # - Italic uses same set but with different outlines
        # - BoldItalic uses its own set

        # Key insight: if there are exactly 4 non-mono prefixes, they map to:
        # regular, bold, italic, bolditalic
        # We can determine this by analyzing which texts use which prefix

        # Check if this prefix is used only for "bold" marked text
        # In typst simple theme, bold text gets a specific color (like #3f6d80)
        # But this is theme-specific...

        # Better heuristic: Check if this prefix shares characters with the
        # regular prefix (same text appears in both) - if so, it's a style variant

        # For now, assign based on order and glyph characteristics
        # We know: regular is most used. Among remaining:
        # - The one with most uses is likely bold (headings are bold and repeat)
        # - The one with fewest is likely bolditalic (rare combination)
        # - The middle one is italic

        pass

    # Build char-to-glyph mapping for cross-prefix comparison
    char_to_glyph_by_prefix: Dict[str, Dict[str, str]] = {}

    for text_group in root.iter(f'{{{SVG_NS}}}g'):
        if 'typst-text' not in text_group.get('class', ''):
            continue

        # Get text
        text = ''
        for fo in text_group.iter(f'{{{SVG_NS}}}foreignObject'):
            for div in fo:
                text = (div.text or '').strip()
                break

        if not text:
            continue

        # Get use elements (in order)
        uses = []
        for use in text_group.findall(f'{{{SVG_NS}}}use'):
            href = use.get(f'{{{XLINK_NS}}}href') or use.get('href', '')
            if href.startswith('#'):
                uses.append(href[1:])

        # Pair characters with glyphs (approximate - may not work with ligatures)
        if len(uses) == len(text):
            for i, char in enumerate(text):
                glyph_id = uses[i]
                if glyph_id in glyph_defs:
                    prefix = glyph_defs[glyph_id].prefix
                    if char not in char_to_glyph_by_prefix:
                        char_to_glyph_by_prefix[char] = {}
                    char_to_glyph_by_prefix[char][prefix] = glyph_id

    # Compare same characters across prefixes using glyph widths
    unassigned = [p for p, _ in remaining if p not in prefix_to_style]

    if unassigned:
        # Find the regular prefix for comparison
        regular_prefix = None
        for p, s in prefix_to_style.items():
            if s == 'regular':
                regular_prefix = p
                break

        if regular_prefix:
            # For each unassigned prefix, compare glyph widths with regular
            prefix_width_ratios: Dict[str, List[float]] = {p: [] for p in unassigned}

            for char, mappings in char_to_glyph_by_prefix.items():
                if char == ' ':
                    continue  # Skip spaces

                if regular_prefix not in mappings:
                    continue

                regular_glyph = mappings[regular_prefix]
                regular_width = _get_glyph_width(glyph_defs[regular_glyph].path_data)

                if regular_width <= 0:
                    continue

                for prefix in unassigned:
                    if prefix in mappings:
                        this_glyph = mappings[prefix]
                        this_width = _get_glyph_width(glyph_defs[this_glyph].path_data)
                        if this_width > 0:
                            ratio = this_width / regular_width
                            prefix_width_ratios[prefix].append(ratio)

            # Classify based on width ratio relative to regular:
            # Bold: width > regular (ratio > 1.0)
            # Italic: width < regular (ratio < 1.0)
            # BoldItalic: width between regular and bold, or > regular
            prefix_avg_ratios = {}
            for prefix, ratios in prefix_width_ratios.items():
                if ratios:
                    prefix_avg_ratios[prefix] = sum(ratios) / len(ratios)
                else:
                    prefix_avg_ratios[prefix] = 1.0

            if len(unassigned) == 1:
                p = unassigned[0]
                ratio = prefix_avg_ratios.get(p, 1.0)
                if ratio > 1.05:
                    prefix_to_style[p] = 'bold'
                elif ratio < 0.95:
                    prefix_to_style[p] = 'italic'
                else:
                    prefix_to_style[p] = 'bold'  # Default to bold for headings

            elif len(unassigned) == 2:
                p1, p2 = unassigned
                r1 = prefix_avg_ratios.get(p1, 1.0)
                r2 = prefix_avg_ratios.get(p2, 1.0)
                if r1 > r2:
                    prefix_to_style[p1] = 'bold'
                    prefix_to_style[p2] = 'italic'
                else:
                    prefix_to_style[p1] = 'italic'
                    prefix_to_style[p2] = 'bold'

            elif len(unassigned) >= 3:
                # Sort by width ratio (descending)
                sorted_prefixes = sorted(unassigned, key=lambda p: prefix_avg_ratios.get(p, 1.0), reverse=True)
                prefix_to_style[sorted_prefixes[0]] = 'bold'
                prefix_to_style[sorted_prefixes[-1]] = 'italic'
                for p in sorted_prefixes[1:-1]:
                    prefix_to_style[p] = 'bolditalic'
        else:
            # No regular prefix found, assign by count
            for p in unassigned:
                prefix_to_style[p] = 'regular'

    # Update the font_variants dict with the detected styles
    for prefix, style in prefix_to_style.items():
        if prefix in font_variants:
            font_variants[prefix].style = style

    return prefix_to_style


def _detect_font_style_from_context(
    text_group: ET.Element,
    font_variants: Dict[str, FontVariant],
    glyph_defs: Dict[str, GlyphInfo],
) -> str:
    """Detect font style by analyzing the glyphs used in a text group.

    Returns one of: regular, bold, italic, bolditalic, mono, math
    """
    # Collect all glyph prefixes used in this text group
    prefixes_used: Set[str] = set()

    for use_elem in text_group.iter(f'{{{SVG_NS}}}use'):
        href = use_elem.get(f'{{{XLINK_NS}}}href') or use_elem.get('href', '')
        if href.startswith('#'):
            glyph_id = href[1:]
            if glyph_id in glyph_defs:
                prefixes_used.add(glyph_defs[glyph_id].prefix)

    if not prefixes_used:
        return "regular"

    # Use the pre-computed variant styles
    for prefix in prefixes_used:
        if prefix in font_variants:
            variant = font_variants[prefix]
            if variant.style != 'regular':
                return variant.style

    return "regular"


def _extract_text_from_foreign_object(
    fo_elem: ET.Element,
    parent_transform: List[float],
    text_group_transform: List[float],
    fill_color: str,
    font_variant: str,
    computed_font_size: float,
) -> Optional[TextSegment]:
    """Extract text content from a <foreignObject> element.

    typst.ts structure:
    <g transform="scale(16,-16)">
      <foreignObject x="0" y="-55.88" width="200" height="62.5">
        <h5:div class="tsel" style="font-size: 62px">Hello</h5:div>
      </foreignObject>
    </g>
    """
    x = float(fo_elem.get('x', '0'))
    y = float(fo_elem.get('y', '0'))
    width = float(fo_elem.get('width', '0'))
    height = float(fo_elem.get('height', '0'))

    # Find the text content - could be in h5:div or xhtml:div
    text = ""
    fo_font_size = None
    css_class = ""

    for child in fo_elem:
        child_tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if child_tag in ('div', 'span', 'p'):
            text = child.text or ""
            css_class = child.get('class', '')
            # Extract font-size from inline style
            style = child.get('style', '')
            size_match = re.search(r'font-size:\s*([\d.]+)px', style)
            if size_match:
                fo_font_size = float(size_match.group(1))
            # Also check for nested elements
            for sub in child:
                sub_text = sub.text or ""
                if sub_text:
                    text += sub_text
                sub_tail = sub.tail or ""
                if sub_tail:
                    text += sub_tail
            break

    if not text:
        return None

    return TextSegment(
        text=text,
        x=x,
        y=y,
        width=width,
        height=height,
        font_size=computed_font_size,
        font_variant=font_variant,
        fill_color=fill_color,
        css_class=css_class,
    )


def _compute_accumulated_transform(transforms: List[Tuple[float, float, float, float, float]]) -> Tuple[float, float, float, float]:
    """Compute accumulated transform from a stack of (dx, dy, sx, sy, rot) tuples.

    Returns (total_dx, total_dy, total_sx, total_sy).
    Simplified: ignores rotation for position calculation.
    """
    total_dx, total_dy = 0.0, 0.0
    total_sx, total_sy = 1.0, 1.0

    for dx, dy, sx, sy, rot in transforms:
        total_dx = total_dx * 1 + dx * 1  # Simplified accumulation
        total_dy = total_dy * 1 + dy * 1
        total_sx *= sx
        total_sy *= sy

    return (total_dx, total_dy, total_sx, total_sy)


def _process_text_group(
    text_group: ET.Element,
    page_data: PageData,
    glyph_defs: Dict[str, GlyphInfo],
    font_variants: Dict[str, FontVariant],
    parent_transforms: List[Tuple[float, float, float, float, float]],
    page_y_offset: float,
):
    """Process a typst-text group to extract text segments.

    typst.ts text group structure (with full hierarchy):

    <g class="typst-page" transform="translate(0, pageY)">       ← page
      <g class="typst-group">                                     ← no transform
        <g transform="translate(50, 50)">                         ← section offset
          <g class="typst-group">                                 ← no transform
            <g transform="translate(0, 150.882)">                 ← sub-section offset
              <g class="typst-group">                             ← no transform
                <g transform="translate(287.395, 19.350)">        ← text position
                  <g class="typst-text" transform="scale(0.025,-0.025)" fill="#000">
                    <use x="0" href="#gVYuD1jU"/>
                    <g transform="scale(16,-16)">                 ← inner scale (ALWAYS 16,-16)
                      <foreignObject x="0" y="-55.88" width="200" height="62.5">
                        <h5:div class="tsel" style="font-size: 62px">Hello</h5:div>
                      </foreignObject>
                    </g>
                  </g>
                </g>

    Key insight: The text position on the slide is determined by the accumulated
    translate() transforms from parent groups. The foreignObject's x,y coordinates
    are for aligning the HTML text overlay with the glyph outlines and should NOT
    be used as the text's position on the slide.

    Font size formula: glyph_scale × inner_scale × fo_font_size
    e.g., 0.025 × 16 × 62 = 24.8 SVG px
    """
    # Get the text group's own transform
    group_transform = parse_transform(text_group.get('transform'))

    # Get fill color from the text group
    fill_color = text_group.get('fill', '#000000')

    # Detect font style from glyph usage
    font_variant = _detect_font_style_from_context(text_group, font_variants, glyph_defs)

    # Compute the glyph scale from the text group transform
    # typst.ts uses scale(S, -S) where S is the glyph scale
    glyph_scale = abs(group_transform[2])  # sx from scale(S, -S)

    # Compute the text position from accumulated parent translations
    # The parent_transforms list contains (dx, dy, sx, sy, rot) tuples
    # from the page group down to the immediate parent of this typst-text group.
    # For typst.ts SVGs, the intermediate groups are all translate() only (sx=sy=1),
    # so we simply sum the translations.
    text_x = 0.0
    text_y = 0.0
    for ptx, pty, psx, psy, prot in parent_transforms:
        text_x = text_x * psx + ptx
        text_y = text_y * psy + pty

    # Subtract page offset to get position within the slide
    text_y -= page_y_offset

    # Collect glyph <use> references from the typst-text group
    glyph_uses = []
    for child in text_group:
        child_tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if child_tag == 'use':
            href = child.get(f'{{{XLINK_NS}}}href') or child.get('href', '')
            if href.startswith('#'):
                glyph_id = href[1:]
                use_x = float(child.get('x', '0'))
                use_y = float(child.get('y', '0'))
                if glyph_id in glyph_defs:
                    glyph_uses.append(GlyphUse(
                        glyph_id=glyph_id,
                        x_offset=use_x,
                        y_offset=use_y,
                        path_data=glyph_defs[glyph_id].path_data,
                    ))

    # Process nested <g> elements looking for foreignObject
    for child in text_group:
        child_tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag

        if child_tag == 'g':
            # This is the scale(16, -16) inner group
            inner_transform = parse_transform(child.get('transform'))
            inner_scale = abs(inner_transform[2])  # typically 16

            for fo_elem in child:
                fo_tag = fo_elem.tag.split('}')[-1] if '}' in fo_elem.tag else fo_elem.tag
                if fo_tag == 'foreignObject':
                    # Extract fo_font_size from the foreignObject content
                    fo_font_size = None
                    for div in fo_elem:
                        style = div.get('style', '')
                        size_match = re.search(r'font-size:\s*([\d.]+)px', style)
                        if size_match:
                            fo_font_size = float(size_match.group(1))

                    if fo_font_size is None:
                        fo_font_size = 16.0  # default

                    # Compute actual font size in SVG pixels
                    # formula: actual_font_size = glyph_scale * inner_scale * fo_font_size
                    computed_font_size = glyph_scale * inner_scale * fo_font_size

                    segment = _extract_text_from_foreign_object(
                        fo_elem,
                        parent_transform=[],
                        text_group_transform=[],
                        fill_color=fill_color,
                        font_variant=font_variant,
                        computed_font_size=computed_font_size,
                    )

                    if segment:
                        # The text position is from the accumulated parent translations
                        # The foreignObject width/height, scaled through inner and glyph scale,
                        # give us the text box dimensions in SVG coordinates
                        fo_width = float(fo_elem.get('width', '0'))
                        fo_height = float(fo_elem.get('height', '0'))

                        # Scale width/height through inner scale and glyph scale
                        actual_width = fo_width * inner_scale * glyph_scale
                        actual_height = fo_height * inner_scale * glyph_scale

                        # The Y position needs a small adjustment:
                        # The text baseline in typst is at the accumulated translate position.
                        # We need to move up by the ascent to get the top of the text box.
                        # A good approximation: shift up by ~80% of the computed font size
                        # (this accounts for the ascent of typical fonts)
                        baseline_y = text_y
                        top_y = baseline_y - computed_font_size * 0.8

                        segment.x = text_x
                        segment.y = top_y
                        segment.width = actual_width
                        segment.height = actual_height

                        # Attach glyph usage info for curve rendering
                        segment.glyph_uses = glyph_uses
                        segment.glyph_scale = glyph_scale

                        page_data.text_segments.append(segment)


def _process_element_recursive(
    elem: ET.Element,
    page_data: PageData,
    glyph_defs: Dict[str, GlyphInfo],
    font_variants: Dict[str, FontVariant],
    all_defs: Dict[str, ET.Element],
    parent_transforms: List[Tuple[float, float, float, float, float]],
    page_y_offset: float,
    glyph_ids: Set[str],
):
    """Recursively process SVG elements within a page.

    Classifies elements into:
    1. Text groups (class="typst-text") → extract text
    2. Shape elements (rect, circle, etc.) → collect as shapes
    3. Group elements → recurse
    4. Glyph paths → skip (handled by text extraction)
    """
    tag = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag

    if tag == 'defs' or tag == 'style':
        return

    # Check if this is a typst-text group
    css_class = elem.get('class', '')
    if 'typst-text' in css_class:
        _process_text_group(
            elem, page_data, glyph_defs, font_variants,
            parent_transforms, page_y_offset,
        )
        return

    # Handle <a> (hyperlink) elements
    if tag == 'a':
        href = elem.get(f'{{{XLINK_NS}}}href') or elem.get('href', '')
        if href:
            # Look for the bounding rect (pseudo-link) inside
            for child in elem:
                child_tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                if child_tag == 'rect':
                    # Compute position with transforms
                    rx = float(child.get('x', '0'))
                    ry = float(child.get('y', '0'))
                    rw = float(child.get('width', '0'))
                    rh = float(child.get('height', '0'))
                    # Apply parent transforms
                    for ptx, pty, psx, psy, prot in parent_transforms:
                        rx = rx * psx + ptx
                        ry = ry * psy + pty
                        rw = rw * abs(psx)
                        rh = rh * abs(psy)
                    ry -= page_y_offset
                    page_data.links.append(LinkRegion(
                        href=href, x=rx, y=ry, width=rw, height=rh,
                    ))
        return

    if tag == 'g':
        # Regular group - recurse into children
        group_transform = parse_transform(elem.get('transform'))
        new_transforms = parent_transforms + [group_transform]

        for child in elem:
            _process_element_recursive(
                child, page_data, glyph_defs, font_variants,
                all_defs, new_transforms, page_y_offset, glyph_ids,
            )
        return

    if tag == 'use':
        # Handle <use> elements - check if it references a glyph (skip) or something else
        href = elem.get(f'{{{XLINK_NS}}}href') or elem.get('href', '')
        if href.startswith('#'):
            ref_id = href[1:]
            if ref_id in glyph_ids:
                # This is a glyph reference - skip (handled by text extraction)
                return
            # Otherwise, it references another element - need to resolve
            if ref_id in all_defs:
                # Clone and process the referenced element with use's transform
                ref_elem = copy.deepcopy(all_defs[ref_id])
                use_transform = parse_transform(elem.get('transform'))
                use_x = float(elem.get('x', '0'))
                use_y = float(elem.get('y', '0'))
                # Combine use position with transform
                combined_transform = (
                    use_transform[0] + use_x,
                    use_transform[1] + use_y,
                    use_transform[2],
                    use_transform[3],
                    use_transform[4],
                )
                new_transforms = parent_transforms + [combined_transform]
                _process_element_recursive(
                    ref_elem, page_data, glyph_defs, font_variants,
                    all_defs, new_transforms, page_y_offset, glyph_ids,
                )
        return

    # Shape elements
    if tag in ('rect', 'circle', 'ellipse', 'line', 'polygon', 'polyline', 'path', 'image'):
        # Check if this path is a glyph definition
        elem_id = elem.get('id', '')
        if elem_id in glyph_ids:
            return  # Skip glyph definitions

        # Check if it's a glyph path by class
        if 'outline_glyph' in elem.get('class', ''):
            return  # Skip glyph outlines

        shape = ShapeElement(
            tag=tag,
            element=elem,
            is_glyph_path=False,
        )

        # Store the accumulated transforms for later processing
        shape.transform_matrix = _transforms_to_list(parent_transforms)

        page_data.shapes.append(shape)
        return

    # For other elements, try to recurse into children
    for child in elem:
        _process_element_recursive(
            child, page_data, glyph_defs, font_variants,
            all_defs, parent_transforms + [parse_transform(elem.get('transform'))],
            page_y_offset, glyph_ids,
        )


def _transforms_to_list(transforms: List[Tuple[float, float, float, float, float]]) -> List[float]:
    """Convert transform stack to a flat [dx, dy, sx, sy, rot_deg, ...] list."""
    result = []
    for dx, dy, sx, sy, rot in transforms:
        result.extend([dx, dy, sx, sy, rot])
    return result


def parse_css_styles(style_text: str) -> Dict[str, Dict[str, str]]:
    """Parse CSS from <style> block.

    Returns dict mapping selector -> {property: value}
    """
    styles = {}
    if not style_text:
        return styles

    # Remove comments
    style_text = re.sub(r'/\*.*?\*/', '', style_text, flags=re.DOTALL)

    # Parse rules
    for match in re.finditer(r'([^{]+)\{([^}]*)\}', style_text):
        selector = match.group(1).strip()
        properties = {}
        for prop in match.group(2).split(';'):
            prop = prop.strip()
            if ':' in prop:
                key, val = prop.split(':', 1)
                properties[key.strip()] = val.strip()
        styles[selector] = properties

    return styles


def parse_typst_svg(svg_path: str) -> TypstSVGData:
    """Parse a typst.ts generated SVG file.

    Args:
        svg_path: Path to the SVG file generated by typst-ts-cli

    Returns:
        TypstSVGData with pages, text segments, shapes, and metadata
    """
    tree = ET.parse(svg_path)
    root = tree.getroot()

    # Parse viewBox
    viewbox_str = root.get('viewBox', '0 0 0 0')
    vb_min_x, vb_min_y, vb_width, vb_height = parse_viewbox(viewbox_str)

    # Parse CSS styles
    css_text = ""
    for style_elem in root.iter(f'{{{SVG_NS}}}style'):
        if style_elem.text:
            css_text += style_elem.text

    # Collect all defs
    all_defs = _collect_defs(root)

    # Analyze glyphs
    glyph_defs, font_variants = _analyze_glyphs(all_defs)
    glyph_ids = set(glyph_defs.keys())

    # Pre-scan to detect font variant styles (bold/italic/mono)
    _prescan_font_variants(root, glyph_defs, font_variants)

    # Find pages - each page is a top-level <g> with translate(0, Y)
    pages = []
    top_level_groups = []

    for child in root:
        child_tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if child_tag == 'g':
            top_level_groups.append(child)

    if not top_level_groups:
        # Might be a single-page SVG without page groups
        # Treat the entire SVG as one page
        page_data = PageData(
            page_num=1,
            y_offset=0,
            width=vb_width,
            height=vb_height,
        )
        for child in root:
            _process_element_recursive(
                child, page_data, glyph_defs, font_variants,
                all_defs, [], 0, glyph_ids,
            )
        pages.append(page_data)
    else:
        # Determine page height from the y-offsets
        y_offsets = []
        for group in top_level_groups:
            transform = parse_transform(group.get('transform'))
            y_offsets.append(transform[1])

        # Calculate page height
        if len(y_offsets) > 1:
            page_height = y_offsets[1] - y_offsets[0]
        else:
            page_height = vb_height

        for i, group in enumerate(top_level_groups):
            transform = parse_transform(group.get('transform'))
            y_offset = transform[1]

            page_data = PageData(
                page_num=i + 1,
                y_offset=y_offset,
                width=vb_width,
                height=page_height,
            )

            # Process all children of this page group
            for child in group:
                _process_element_recursive(
                    child, page_data, glyph_defs, font_variants,
                    all_defs, [transform], y_offset, glyph_ids,
                )

            pages.append(page_data)

    return TypstSVGData(
        viewbox_width=vb_width,
        viewbox_height=vb_height,
        pages=pages,
        glyph_defs=glyph_defs,
        font_variants=font_variants,
        css_styles=css_text,
        defs=all_defs,
    )
