#!/usr/bin/env python3
"""Download typst-ts-cli binary for the current (or specified) platform.

Usage:
    python scripts/download_typst_ts_cli.py              # auto-detect platform
    python scripts/download_typst_ts_cli.py --target x86_64-unknown-linux-gnu
    python scripts/download_typst_ts_cli.py --version v0.6.0
"""

import argparse
import io
import os
import platform
import stat
import sys
import tarfile
import urllib.request
import zipfile

GITHUB_RELEASE_URL = (
    "https://github.com/Myriad-Dreamin/typst.ts/releases/download"
)

DEFAULT_VERSION = "v0.7.0-rc2"

PLATFORM_MAP = {
    ("Darwin", "arm64"): "aarch64-apple-darwin",
    ("Darwin", "x86_64"): "x86_64-apple-darwin",
    ("Linux", "x86_64"): "x86_64-unknown-linux-gnu",
    ("Linux", "aarch64"): "aarch64-unknown-linux-gnu",
    ("Linux", "armv7l"): "arm-unknown-linux-gnueabihf",
    ("Windows", "AMD64"): "x86_64-pc-windows-msvc",
    ("Windows", "x86"): "i686-pc-windows-msvc",
    ("Windows", "ARM64"): "aarch64-pc-windows-msvc",
}


def detect_target() -> str:
    """Detect the Rust-style target triple for the current platform."""
    system = platform.system()
    machine = platform.machine()

    target = PLATFORM_MAP.get((system, machine))
    if target is None:
        raise RuntimeError(
            f"Unsupported platform: {system}/{machine}. "
            f"Supported: {list(PLATFORM_MAP.keys())}"
        )
    return target


def download_typst_ts_cli(
    version: str = DEFAULT_VERSION,
    target: str | None = None,
    output_dir: str | None = None,
) -> str:
    """Download and extract typst-ts-cli binary.

    Returns the path to the extracted binary.
    """
    if target is None:
        target = detect_target()

    if output_dir is None:
        output_dir = os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            "typ2pptx", "data", "bin",
        )

    is_windows = "windows" in target or "msvc" in target
    archive_ext = "zip" if is_windows else "tar.gz"
    asset_name = f"typst-ts-{target}.{archive_ext}"
    url = f"{GITHUB_RELEASE_URL}/{version}/{asset_name}"

    print(f"Downloading {url} ...")
    response = urllib.request.urlopen(url)
    data = response.read()

    binary_name = "typst-ts-cli.exe" if is_windows else "typst-ts-cli"

    # Try multiple possible paths (different versions have different structures)
    # Windows zip: binary is at root level; Unix tar.gz: may have bin/ subdirectory
    if is_windows:
        possible_paths = [
            f"{binary_name}",                              # v0.7.0-rc2: root level
            f"typst-ts-{target}/{binary_name}",             # older structure
            f"typst-ts-{target}/bin/{binary_name}",
        ]
    else:
        possible_paths = [
            f"typst-ts-{target}/bin/{binary_name}",   # v0.6.0 and earlier
            f"typst-ts-{target}/{binary_name}",         # v0.7.0-rc2 and later
        ]

    print(f"Extracting typst-ts-cli ...")

    if is_windows:
        # Handle ZIP archives for Windows
        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            member_name = None
            for path in possible_paths:
                if path in zf.namelist():
                    member_name = path
                    print(f"  Found at: {path}")
                    break

            if member_name is None:
                available = [n for n in zf.namelist() if "typst-ts-cli" in n]
                raise RuntimeError(
                    f"Could not find typst-ts-cli binary in archive. "
                    f"Tried: {possible_paths}. Available files: {available}"
                )

            os.makedirs(output_dir, exist_ok=True)
            output_path = os.path.join(output_dir, binary_name)

            with zf.open(member_name) as src, open(output_path, "wb") as dst:
                dst.write(src.read())
    else:
        # Handle tar.gz archives for Unix-like systems
        with tarfile.open(fileobj=io.BytesIO(data), mode="r:gz") as tar:
            member = None
            for path in possible_paths:
                try:
                    member = tar.getmember(path)
                    print(f"  Found at: {path}")
                    break
                except KeyError:
                    continue

            if member is None:
                available = [m.name for m in tar.getmembers() if "typst-ts-cli" in m.name]
                raise RuntimeError(
                    f"Could not find typst-ts-cli binary in archive. "
                    f"Tried: {possible_paths}. Available files: {available}"
                )

            fileobj = tar.extractfile(member)
            if fileobj is None:
                raise RuntimeError(f"Could not extract {member.name} from archive")

            os.makedirs(output_dir, exist_ok=True)
            output_path = os.path.join(output_dir, binary_name)

            with open(output_path, "wb") as out:
                out.write(fileobj.read())

        # Make executable on Unix
        current_mode = os.stat(output_path).st_mode
        os.chmod(output_path, current_mode | stat.S_IXUSR | stat.S_IXGRP | stat.S_IXOTH)

    print(f"Installed typst-ts-cli to {output_path}")
    return output_path


def main():
    parser = argparse.ArgumentParser(
        description="Download typst-ts-cli binary for bundling with typ2pptx",
    )
    parser.add_argument(
        "--version",
        default=DEFAULT_VERSION,
        help=f"typst.ts release version (default: {DEFAULT_VERSION})",
    )
    parser.add_argument(
        "--target",
        default=None,
        help="Rust target triple (default: auto-detect)",
    )
    parser.add_argument(
        "--output-dir",
        default=None,
        help="Output directory (default: typ2pptx/data/bin/)",
    )
    args = parser.parse_args()

    try:
        download_typst_ts_cli(
            version=args.version,
            target=args.target,
            output_dir=args.output_dir,
        )
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
