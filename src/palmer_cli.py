"""
palmer_cli.py - Palmer dental notation command-line interface

Usage:
    # Single quadrant to PNG
    python palmer_cli.py --UL 1 -o UL1.png

    # Full dentition
    python palmer_cli.py --UL 12345678 --UR 12345678 --LR 12345678 --LL 12345678 -o full.png

    # Copy to clipboard (Windows)
    python palmer_cli.py --UL 1 --clipboard

    # Batch processing from JSON
    python palmer_cli.py --batch snippets.json --outdir images/

    # Raw TeX input
    python palmer_cli.py --raw "\\Palmer{1}{}{}{}{}{}" -o test.png
"""

from __future__ import annotations

import argparse
import json
import logging
import math
import re
import sys
from pathlib import Path

from palmer_engine import (
    PalmerCompiler, validate_raw_input,
    MAX_FIELD_LEN, MAX_RAW_LEN,
    MIN_DPI, MAX_DPI, DEFAULT_DPI,
    MIN_FONT_SIZE_PT, MAX_FONT_SIZE_PT, DEFAULT_FONT_SIZE_PT,
    DEFAULT_MARGIN_PX,
    tectonic_cache_exists,
)

try:
    from version import __version__
except ImportError:
    __version__ = "dev"


def main():
    parser = argparse.ArgumentParser(
        prog="palmer-type",
        description="Palmer dental notation renderer (tectonic / TeX Live / MiKTeX)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s --UL 1 -o UL1.png
  %(prog)s --UL 12345678 --UR 12345678 --LR 12345678 --LL 12345678 -o full.png
  %(prog)s --UL 321 --UR 123 -o upper_ant.png
  %(prog)s --batch snippets.json --outdir images/
  %(prog)s --raw "\\Palmer[center]{1}{}{}{}{}{}" -o custom.png

Exit codes:
  0  Success
  1  Error (initialization failed, or all batch items failed)
  2  Partial failure (some batch items succeeded, some failed)
        """,
    )

    parser.add_argument(
        "--version", action="version", version=f"%(prog)s {__version__}",
    )

    # Quadrant inputs
    parser.add_argument("--UL", default="", help="upper-left quadrant")
    parser.add_argument("--UR", default="", help="upper-right quadrant")
    parser.add_argument("--LR", default="", help="lower-right quadrant")
    parser.add_argument("--LL", default="", help="lower-left quadrant")
    parser.add_argument("--upper-mid", default="", help="upper midline symbol")
    parser.add_argument("--lower-mid", default="", help="lower midline symbol")
    parser.add_argument("--option", default="base",
                        choices=["base", "center", "bottom"],
                        help="vertical alignment of the cross (default: base)")

    # Font
    parser.add_argument("--font", default="Times New Roman",
                        help="font family name (default: Times New Roman)")
    parser.add_argument("--font-size", type=float, default=DEFAULT_FONT_SIZE_PT,
                        help=f"font size in points, {MIN_FONT_SIZE_PT}-{MAX_FONT_SIZE_PT} (default: {DEFAULT_FONT_SIZE_PT})")

    # Output
    parser.add_argument("-o", "--output", type=Path, help="output PNG file path")
    parser.add_argument("--clipboard", action="store_true",
                        help="copy rendered image to clipboard (Windows)")

    # Batch mode
    parser.add_argument("--batch", type=Path, help="JSON file for batch processing")
    parser.add_argument("--outdir", type=Path, default=Path("output"),
                        help="output directory for batch results (default: output/)")

    # Raw TeX
    parser.add_argument("--raw", type=str, help="raw TeX body to compile directly")

    # Options
    parser.add_argument("--dpi", type=int, default=DEFAULT_DPI, help=f"output resolution in DPI (default: {DEFAULT_DPI})")
    parser.add_argument(
        "--color", default="", metavar="COLOR",
        help="text color: 6-digit hex (#RRGGBB) or named xcolor color (e.g., red, blue, darkgray)",
    )
    parser.add_argument("--sty", type=Path, help="path to palmer.sty")
    parser.add_argument("--transparent", action="store_true",
                        help="use transparent background (PNG and clipboard)")

    # Margins
    parser.add_argument("--margin-top", type=int, default=DEFAULT_MARGIN_PX,
                        help=f"top margin in pixels (default: {DEFAULT_MARGIN_PX})")
    parser.add_argument("--margin-bottom", type=int, default=DEFAULT_MARGIN_PX,
                        help=f"bottom margin in pixels (default: {DEFAULT_MARGIN_PX})")
    parser.add_argument("--margin-left", type=int, default=DEFAULT_MARGIN_PX,
                        help=f"left margin in pixels (default: {DEFAULT_MARGIN_PX})")
    parser.add_argument("--margin-right", type=int, default=DEFAULT_MARGIN_PX,
                        help=f"right margin in pixels (default: {DEFAULT_MARGIN_PX})")

    parser.add_argument(
        "--verbose", "-v", action="store_true",
        help="enable debug logging from internal modules",
    )

    args = parser.parse_args()

    # Configure logging.  By default only WARNING+ is shown so that library
    # modules (palmer_engine, palmer_converter, …) can surface warnings without
    # flooding the terminal.  --verbose lowers the threshold to DEBUG so all
    # logger.debug() calls become visible.
    logging.basicConfig(
        stream=sys.stderr,
        level=logging.DEBUG if args.verbose else logging.WARNING,
        format="%(levelname)s %(name)s: %(message)s",
    )

    # Validate DPI range.
    if not (MIN_DPI <= args.dpi <= MAX_DPI):
        parser.error(f"--dpi must be between {MIN_DPI} and {MAX_DPI}.")

    # Validate font size range.
    if not (MIN_FONT_SIZE_PT <= args.font_size <= MAX_FONT_SIZE_PT):
        parser.error(f"--font-size must be between {MIN_FONT_SIZE_PT} and {MAX_FONT_SIZE_PT}.")

    # Validate incompatible option combinations.
    if args.batch and args.raw:
        parser.error("--batch and --raw cannot be used together.")
    if args.batch and args.output:
        parser.error("--batch and --output cannot be used together. Use --outdir for batch output.")

    try:
        compiler = PalmerCompiler(
            sty_path=args.sty,
            dpi=args.dpi,
            margin_top=args.margin_top,
            margin_bottom=args.margin_bottom,
            margin_left=args.margin_left,
            margin_right=args.margin_right,
        )
        print(f"Engine: {compiler.backend.name} ({compiler.backend.executable})", file=sys.stderr)
        if compiler.backend.name == "tectonic" and not tectonic_cache_exists():
            print(
                "Note: Tectonic TeX support files are not cached.\n"
                "      An internet connection is required for this render.",
                file=sys.stderr,
            )
    except (FileNotFoundError, RuntimeError) as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)

    # --- Batch mode ---
    if args.batch:
        # Reject paths that contain '..' components to prevent directory traversal
        # outside the intended output location.
        if ".." in args.outdir.parts:
            print(
                "Error: --outdir must not contain '..' components.",
                file=sys.stderr,
            )
            sys.exit(1)
        rc = _batch_process(compiler, args.batch, args.outdir, args.dpi)
        if rc:
            sys.exit(rc)
        return

    # --- Single render ---
    if args.raw:
        validate_raw_input(args.raw, "--raw")
        img = compiler.render_raw(args.raw, alpha=args.transparent)
    else:
        if not any((args.UL, args.UR, args.LR, args.LL)):
            parser.error("Specify at least one quadrant (--UL, --UR, --LR, --LL).")
        img = compiler.render(
            UL=args.UL, UR=args.UR, LR=args.LR, LL=args.LL,
            upper_mid=args.upper_mid, lower_mid=args.lower_mid,
            option=args.option,
            font_family=args.font,
            font_size_pt=args.font_size,
            text_color=args.color,
            alpha=args.transparent,
        )

    print(f"Rendered: {img.width}×{img.height}px", file=sys.stderr)

    if args.output:
        suffix = args.output.suffix.lower()
        if args.transparent and suffix in (".jpg", ".jpeg", ".pdf"):
            print(
                f"Warning: --transparent is ignored for {suffix} output "
                f"(only PNG supports transparency).",
                file=sys.stderr,
            )
            img = img.convert("RGB")
        img.save(str(args.output), dpi=(args.dpi, args.dpi))
        print(f"Saved: {args.output}", file=sys.stderr)

    if args.clipboard:
        if sys.platform != "win32":
            print("Error: --clipboard is only supported on Windows.", file=sys.stderr)
            sys.exit(1)
        from palmer_engine import copy_image_to_clipboard_win32
        copy_image_to_clipboard_win32(img, dpi=args.dpi)
        print("Copied to clipboard.", file=sys.stderr)

    if not args.output and not args.clipboard:
        # Write PNG to stdout when no output destination is specified.
        if sys.stdout.isatty():
            print(
                "Error: No output destination specified.\n"
                "Use -o FILE to save, --clipboard to copy, or pipe to another command.",
                file=sys.stderr,
            )
            sys.exit(1)
        import io
        buf = io.BytesIO()
        img.save(buf, format="PNG", dpi=(args.dpi, args.dpi))
        sys.stdout.buffer.write(buf.getvalue())


def _batch_process(compiler: PalmerCompiler, json_path: Path, outdir: Path, dpi: int) -> int:
    """Process a JSON batch file and save each result to outdir.

    Returns an exit code:
        0 -- all items succeeded (or all were skipped with none attempted)
        1 -- all attempted items failed
        2 -- partial success (some succeeded, some failed)
    """
    try:
        with open(json_path, encoding="utf-8") as f:
            snippets = json.load(f)
    except (OSError, json.JSONDecodeError) as e:
        print(f"Error: Failed to read JSON file: {e}", file=sys.stderr)
        return 1

    if not isinstance(snippets, list):
        print("Error: The JSON top level must be an array.", file=sys.stderr)
        return 1

    if not snippets:
        print("[WARN] Batch file is empty.", file=sys.stderr)
        return 0

    outdir.mkdir(parents=True, exist_ok=True)

    ok = 0
    failed = 0
    skipped = 0

    def skip(i: int, reason: str) -> None:
        nonlocal skipped
        print(f"  [SKIP] Index {i}: {reason}", file=sys.stderr)
        skipped += 1

    for i, item in enumerate(snippets):
        if not isinstance(item, dict) or "id" not in item:
            skip(i, f"missing 'id' field: {item!r}")
            continue
        sid = item["id"]
        if not isinstance(sid, str):
            skip(i, f"'id' must be a string, got {type(sid).__name__}: {sid!r}")
            continue
        # Restrict id length and character set (alphanumeric, hyphen, underscore)
        # to prevent directory traversal, filesystem-unsafe filenames, and
        # path-length overflow.
        if len(sid) > MAX_FIELD_LEN:
            skip(i, f"id exceeds maximum length of {MAX_FIELD_LEN} characters: {sid[:40]!r}...")
            continue
        if not re.match(r'^[a-zA-Z0-9_-]+$', sid):
            skip(i, f"id contains unsafe characters: {sid!r}  "
                     f"(only a-z, A-Z, 0-9, hyphen, underscore are allowed)")
            continue
        safe_name = sid
        out_path = outdir / f"{safe_name}.png"
        # Defence-in-depth: ensure the resolved path is within outdir.
        if not out_path.resolve().is_relative_to(outdir.resolve()):
            skip(i, f"resolved path escapes output directory: {sid!r}")
            continue

        if out_path.exists():
            print(
                f"  [WARN] Index {i}: overwriting existing file: {out_path.name}",
                file=sys.stderr,
            )

        try:
            for key in ("UL", "UR", "LR", "LL", "upper_mid", "lower_mid", "raw", "option", "color", "font"):
                if key in item:
                    if not isinstance(item[key], str):
                        raise ValueError(
                            f"Field '{key}' must be a string, "
                            f"got {type(item[key]).__name__}: {item[key]!r}"
                        )
                    limit = MAX_RAW_LEN if key == "raw" else MAX_FIELD_LEN
                    if len(item[key]) > limit:
                        raise ValueError(
                            f"Field '{key}' exceeds the maximum length of {limit} characters"
                        )
            if "option" in item and item["option"] not in ("base", "center", "bottom"):
                raise ValueError(
                    f"Field 'option' must be one of 'base', 'center', 'bottom', "
                    f"got {item['option']!r}"
                )
            item_alpha = bool(item.get("transparent", False))
            if "raw" in item:
                validate_raw_input(item["raw"], f"batch[{i}].raw")
                img = compiler.render_raw(item["raw"], alpha=item_alpha)
            else:
                font_size = item.get("font_size", DEFAULT_FONT_SIZE_PT)
                if not isinstance(font_size, (int, float)):
                    raise ValueError(
                        f"Field 'font_size' must be a number, "
                        f"got {type(font_size).__name__}: {font_size!r}"
                    )
                # Defense-in-depth: JSON cannot produce Infinity/NaN, but
                # guard against unusual decoders or future input sources.
                if not math.isfinite(font_size):
                    raise ValueError(
                        f"Field 'font_size' must be a finite number, "
                        f"got {font_size!r}"
                    )
                if not (MIN_FONT_SIZE_PT <= font_size <= MAX_FONT_SIZE_PT):
                    raise ValueError(
                        f"Field 'font_size' must be between "
                        f"{MIN_FONT_SIZE_PT} and {MAX_FONT_SIZE_PT}, "
                        f"got {font_size}"
                    )
                img = compiler.render(
                    UL=item.get("UL", ""),
                    UR=item.get("UR", ""),
                    LR=item.get("LR", ""),
                    LL=item.get("LL", ""),
                    upper_mid=item.get("upper_mid", ""),
                    lower_mid=item.get("lower_mid", ""),
                    option=item.get("option", "base"),
                    font_family=item.get("font", "Times New Roman"),
                    font_size_pt=float(font_size),
                    text_color=item.get("color", ""),
                    alpha=item_alpha,
                )
            img.save(str(out_path), dpi=(dpi, dpi))
            label = item.get("label", sid)
            # Strip control characters to prevent terminal escape injection.
            if isinstance(label, str):
                label = "".join(c for c in label if c.isprintable() or c == ' ')
                label = label[:200]
            print(f"  [OK] {sid}: {label} ({img.width}x{img.height}px)", file=sys.stderr)
            ok += 1
        except (ValueError, RuntimeError, OSError) as e:
            print(f"  [FAIL] {sid}: {e}", file=sys.stderr)
            failed += 1

    total = len(snippets)
    status = "OK" if failed == 0 else "WARN"
    print(
        f"\n[{status}] {ok} ok, {failed} failed, {skipped} skipped / {total} total",
        file=sys.stderr,
    )
    if failed == 0:
        return 0  # all succeeded or all skipped
    if ok == 0:
        return 1  # all attempted items failed
    return 2  # partial failure


if __name__ == "__main__":
    main()
