#!/usr/bin/env -S uv run --script
# /// script
# requires-python = ">=3.11"
# dependencies = ["python-pptx>=0.6.21"]
# ///
"""Fast structural screen: flag PPTX slides where two DIFFERENT text-box
shapes' bounding boxes overlap significantly. This is the exact signature
of every real bug found so far in this fork's font-fidelity work (subtitle
over headline, a detached word over the next line, a numeral over the
title) -- much cheaper than rendering every slide via soffice and
pixel-diffing, since it works directly off the saved XML geometry.

Not a substitute for visual review (a false negative is possible if two
boxes overlap but happen to not contain colliding ink, e.g. one is empty
whitespace padding) -- it's a triage pass: run this across every deck,
then only spend soffice-render time on what it flags.

Usage:
    tools/detect_overlaps.py <pptx_path> [<pptx_path> ...]
    tools/detect_overlaps.py --glob '~/Desktop/talk-structures/**/*.pptx'
"""
from __future__ import annotations

import argparse
import glob
import json
import os
import sys

from pptx import Presentation


def rects_overlap_fraction(a, b):
    """Fraction of the SMALLER rect's area covered by the intersection."""
    ax0, ay0, ax1, ay1 = a
    bx0, by0, bx1, by1 = b
    ix0, iy0 = max(ax0, bx0), max(ay0, by0)
    ix1, iy1 = min(ax1, bx1), min(ay1, by1)
    if ix1 <= ix0 or iy1 <= iy0:
        return 0.0
    inter = (ix1 - ix0) * (iy1 - iy0)
    a_area = (ax1 - ax0) * (ay1 - ay0)
    b_area = (bx1 - bx0) * (by1 - by0)
    smaller = min(a_area, b_area)
    return inter / smaller if smaller > 0 else 0.0


def check_deck(path: str, threshold: float = 0.15):
    findings = []
    try:
        prs = Presentation(path)
    except Exception as e:
        return [{"slide": None, "error": str(e)}]

    for si, slide in enumerate(prs.slides, start=1):
        boxes = []
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            txt = shape.text_frame.text.strip()
            if not txt:
                continue
            if shape.left is None or shape.top is None or shape.width is None or shape.height is None:
                continue
            rect = (shape.left, shape.top, shape.left + shape.width, shape.top + shape.height)
            boxes.append((txt, rect))

        for i in range(len(boxes)):
            for j in range(i + 1, len(boxes)):
                t1, r1 = boxes[i]
                t2, r2 = boxes[j]
                frac = rects_overlap_fraction(r1, r2)
                if frac < threshold:
                    continue
                # A short "chrome" label (page number, kicker, date) is
                # commonly sized larger than its visible text by auto-fit
                # quirks -- overlapping a chrome label's declared box is
                # not the signature we're hunting (subtitle-over-headline,
                # word-over-next-line: two substantial content strings).
                # Only chrome-vs-chrome or a very high-confidence overlap
                # still gets flagged.
                shorter_len = min(len(t1), len(t2))
                if shorter_len < 20 and frac < 0.6:
                    continue
                findings.append({
                    "slide": si,
                    "overlap_fraction": round(frac, 3),
                    "text_a": t1[:60],
                    "text_b": t2[:60],
                })
    return findings


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("paths", nargs="*")
    ap.add_argument("--glob", action="append", default=[])
    ap.add_argument("--threshold", type=float, default=0.15)
    args = ap.parse_args()

    paths = list(args.paths)
    for g in args.glob:
        paths.extend(sorted(glob.glob(os.path.expanduser(g), recursive=True)))

    if not paths:
        print("no paths given", file=sys.stderr)
        return 2

    results = {}
    flagged = 0
    for p in paths:
        findings = check_deck(p, threshold=args.threshold)
        if findings:
            results[p] = findings
            flagged += 1

    print(json.dumps({"checked": len(paths), "flagged": flagged, "results": results}, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
