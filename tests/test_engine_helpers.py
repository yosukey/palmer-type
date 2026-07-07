"""Tests for numeric and image helpers in palmer_engine.

Covers:
  - clamp_dpi: valid input, clamping, non-finite / garbage fallback
  - reconstruct_alpha_on_white: coverage recovery for black and coloured text
"""

from __future__ import annotations

import sys
from pathlib import Path

from PIL import Image

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from palmer_engine import (
    clamp_dpi,
    reconstruct_alpha_on_white,
    MIN_DPI,
    MAX_DPI,
    DEFAULT_DPI,
)


def _composite_on(rgba: Image.Image, bg: tuple[int, int, int]) -> Image.Image:
    """Composite an RGBA image over a solid background, returning RGB."""
    out = Image.new("RGBA", rgba.size, bg + (255,))
    out.paste(rgba, mask=rgba.split()[3])
    return out.convert("RGB")


# ---------------------------------------------------------------------------
# clamp_dpi
# ---------------------------------------------------------------------------

class TestClampDpi:

    def test_within_range(self):
        assert clamp_dpi(300) == 300
        assert clamp_dpi("300") == 300

    def test_clamped_low_and_high(self):
        assert clamp_dpi(1) == MIN_DPI
        assert clamp_dpi(999999) == MAX_DPI

    def test_float_string(self):
        assert clamp_dpi("300.7") == 300

    def test_garbage_returns_fallback(self):
        assert clamp_dpi("abc") == DEFAULT_DPI
        assert clamp_dpi(None) == DEFAULT_DPI  # type: ignore[arg-type]

    def test_nan_returns_fallback(self):
        assert clamp_dpi("nan") == DEFAULT_DPI

    def test_infinity_returns_fallback(self):
        """Non-finite values must not raise OverflowError (regression)."""
        assert clamp_dpi("inf") == DEFAULT_DPI
        assert clamp_dpi("1e999") == DEFAULT_DPI
        assert clamp_dpi("-inf") == DEFAULT_DPI

    def test_custom_fallback(self):
        assert clamp_dpi("inf", fallback=150) == 150


# ---------------------------------------------------------------------------
# reconstruct_alpha_on_white
# ---------------------------------------------------------------------------

class TestReconstructAlpha:

    def test_black_text_roundtrips_on_white(self):
        """Black text recomposited on white must match the original exactly."""
        img = Image.new("RGB", (3, 1), "white")
        img.putpixel((0, 0), (0, 0, 0))        # solid stroke
        img.putpixel((1, 0), (127, 127, 127))  # antialiased edge
        rgba = reconstruct_alpha_on_white(img, (0, 0, 0))
        comp = _composite_on(rgba, (255, 255, 255))
        assert comp.getpixel((0, 0)) == (0, 0, 0)
        assert comp.getpixel((1, 0)) == (127, 127, 127)
        assert comp.getpixel((2, 0)) == (255, 255, 255)

    def test_black_is_default_text_color(self):
        img = Image.new("RGB", (1, 1), "white")
        img.putpixel((0, 0), (60, 60, 60))
        assert reconstruct_alpha_on_white(img) == reconstruct_alpha_on_white(img, (0, 0, 0))

    def test_background_pixels_fully_transparent(self):
        img = Image.new("RGB", (1, 1), "white")
        rgba = reconstruct_alpha_on_white(img, (255, 255, 0))
        assert rgba.getpixel((0, 0))[3] == 0  # white background -> alpha 0

    def test_light_color_stays_saturated(self):
        """Yellow text must keep full colour and coverage, not wash out.

        A luminance mask would give the solid yellow stroke an alpha near 29;
        the colour-aware reconstruction keeps it fully opaque and saturated.
        """
        img = Image.new("RGB", (2, 1), "white")
        img.putpixel((0, 0), (255, 255, 0))    # solid yellow stroke
        img.putpixel((1, 0), (255, 255, 128))  # ~half-coverage edge
        rgba = reconstruct_alpha_on_white(img, (255, 255, 0))
        # Solid stroke: fully opaque, pure yellow.
        assert rgba.getpixel((0, 0)) == (255, 255, 0, 255)
        # Edge: pure yellow colour, ~half coverage.
        edge = rgba.getpixel((1, 0))
        assert edge[:3] == (255, 255, 0)
        assert 120 <= edge[3] <= 135

    def test_light_color_roundtrips_on_white(self):
        img = Image.new("RGB", (2, 1), "white")
        img.putpixel((0, 0), (255, 255, 0))
        img.putpixel((1, 0), (255, 255, 128))
        rgba = reconstruct_alpha_on_white(img, (255, 255, 0))
        comp = _composite_on(rgba, (255, 255, 255))
        assert comp.getpixel((0, 0)) == (255, 255, 0)
        assert comp.getpixel((1, 0)) == (255, 255, 128)

    def test_white_text_fallback_does_not_crash(self):
        """Near-white text has no channel contrast; fallback must still work."""
        img = Image.new("RGB", (1, 1), "white")
        img.putpixel((0, 0), (200, 200, 200))
        rgba = reconstruct_alpha_on_white(img, (255, 255, 255))
        assert rgba.mode == "RGBA"
        assert rgba.size == (1, 1)
