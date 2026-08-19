#!/usr/bin/env -S uv run --script
# /// script
# requires-python = ">=3.11"
# dependencies = ["Pillow>=10.0.0", "numpy>=1.26.0"]
# ///
"""Pixel-diff a rendered PDF against the PPTX built from the same Typst
source, page by page. Used to drive the font/layout-fidelity fix loop for
this repo's SVG->PPTX converter.

Usage:
    tools/compare_pdf_pptx.py <pdf_path> <pptx_path> <out_dir> [--dpi 150]

Renders both to PNG at matching DPI (PPTX goes via a LibreOffice
--headless PDF conversion first, since soffice is the only reliable
headless PPTX renderer available), diffs each page pair with Pillow,
and writes:
  - out_dir/pdf-page-N.png / out_dir/pptx-page-N.png (the renders)
  - out_dir/diff-page-N.png (a red-highlighted diff heatmap)
  - out_dir/summary.json ({"pages": [{"page", "mean_abs_diff", "pct_pixels_changed"}]})

mean_abs_diff is 0-255 (average per-channel absolute difference).
pct_pixels_changed counts pixels whose per-channel diff exceeds 30/255 --
a proxy for "visibly different", not anti-aliasing noise.
"""
from __future__ import annotations

import argparse
import json
import subprocess
import sys
from pathlib import Path

from PIL import Image, ImageChops
import numpy as np


def render_pdf(pdf_path: Path, out_prefix: Path, dpi: int) -> list[Path]:
    subprocess.run(
        ["pdftoppm", "-png", "-r", str(dpi), str(pdf_path), str(out_prefix)],
        check=True, capture_output=True,
    )
    return sorted(out_prefix.parent.glob(out_prefix.name + "-*.png"))


def pptx_to_pdf(pptx_path: Path, out_dir: Path) -> Path:
    subprocess.run(
        ["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(out_dir), str(pptx_path)],
        check=True, capture_output=True, timeout=120,
    )
    candidates = list(out_dir.glob(pptx_path.stem + ".pdf"))
    if not candidates:
        raise RuntimeError(f"soffice did not produce a PDF for {pptx_path}")
    return candidates[0]


def diff_pair(a_path: Path, b_path: Path, diff_out: Path) -> dict:
    a = Image.open(a_path).convert("RGB")
    b = Image.open(b_path).convert("RGB")
    if a.size != b.size:
        b = b.resize(a.size)
    arr_a = np.asarray(a, dtype=np.int16)
    arr_b = np.asarray(b, dtype=np.int16)
    delta = np.abs(arr_a - arr_b)
    mean_abs_diff = float(delta.mean())
    changed_mask = (delta.max(axis=2) > 30)
    pct_pixels_changed = float(changed_mask.mean() * 100)

    diff_img = ImageChops.difference(a, b)
    diff_arr = np.asarray(diff_img).copy()
    highlight = np.zeros_like(diff_arr)
    highlight[changed_mask] = [255, 0, 0]
    blended = (0.35 * np.asarray(b) + 0.65 * highlight).astype("uint8")
    Image.fromarray(blended).save(diff_out)

    return {"mean_abs_diff": round(mean_abs_diff, 3), "pct_pixels_changed": round(pct_pixels_changed, 3)}


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf_path", type=Path)
    ap.add_argument("pptx_path", type=Path)
    ap.add_argument("out_dir", type=Path)
    ap.add_argument("--dpi", type=int, default=150)
    args = ap.parse_args()

    args.out_dir.mkdir(parents=True, exist_ok=True)

    pdf_pages = render_pdf(args.pdf_path, args.out_dir / "pdf-page", args.dpi)
    pptx_as_pdf = pptx_to_pdf(args.pptx_path, args.out_dir)
    pptx_pages = render_pdf(pptx_as_pdf, args.out_dir / "pptx-page", args.dpi)

    if len(pdf_pages) != len(pptx_pages):
        print(json.dumps({"error": f"page count mismatch: pdf={len(pdf_pages)} pptx={len(pptx_pages)}"}))
        return 2

    results = []
    for i, (pa, pb) in enumerate(zip(pdf_pages, pptx_pages), start=1):
        diff_out = args.out_dir / f"diff-page-{i:02d}.png"
        stats = diff_pair(pa, pb, diff_out)
        stats["page"] = i
        stats["pdf_png"] = str(pa)
        stats["pptx_png"] = str(pb)
        stats["diff_png"] = str(diff_out)
        results.append(stats)

    summary = {
        "pages": results,
        "overall_mean_abs_diff": round(sum(r["mean_abs_diff"] for r in results) / len(results), 3),
        "overall_max_pct_pixels_changed": round(max(r["pct_pixels_changed"] for r in results), 3),
    }
    (args.out_dir / "summary.json").write_text(json.dumps(summary, indent=2))
    print(json.dumps(summary, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
