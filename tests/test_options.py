"""Tests for the palmer.sty v2 optional-argument helpers.

Covers:
  - _build_option_list (palmer_engine): emits the v2 key=value list
    (align=..., gap-ratio=..., no-vert, no-reverse) and validates its inputs.
  - parse_palmer_options (palmer_converter): parses the v2 key=value syntax
    back into structured settings.
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from palmer_engine import _build_option_list, MIN_GAP_RATIO, MAX_GAP_RATIO
from palmer_converter import parse_palmer_options


# ---------------------------------------------------------------------------
# _build_option_list  (engine → TeX)
# ---------------------------------------------------------------------------

class TestBuildOptionList:

    def test_default_base(self):
        assert _build_option_list("base") == "align=base"

    def test_center(self):
        assert _build_option_list("center") == "align=center"

    def test_bottom(self):
        assert _build_option_list("bottom") == "align=bottom"

    def test_no_vert(self):
        assert _build_option_list("base", no_vert=True) == "align=base, no-vert"

    def test_no_reverse(self):
        assert _build_option_list("base", no_reverse=True) == "align=base, no-reverse"

    def test_gap_ratio(self):
        assert _build_option_list("base", gap_ratio=0.5) == "align=base, gap-ratio=0.5"

    def test_gap_ratio_whole_number_is_compact(self):
        # `:g` formatting keeps 1 and 0 without a trailing ".0".
        assert _build_option_list("base", gap_ratio=1) == "align=base, gap-ratio=1"
        assert _build_option_list("base", gap_ratio=0) == "align=base, gap-ratio=0"

    def test_all_keys_order(self):
        # align first, then gap-ratio, then no-vert, then no-reverse.
        assert _build_option_list(
            "center", no_vert=True, no_reverse=True, gap_ratio=0.25,
        ) == "align=center, gap-ratio=0.25, no-vert, no-reverse"

    def test_invalid_option_raises(self):
        with pytest.raises(ValueError, match="option must be one of"):
            _build_option_list("sideways")

    def test_gap_ratio_out_of_range_high(self):
        with pytest.raises(ValueError, match="between"):
            _build_option_list("base", gap_ratio=MAX_GAP_RATIO + 0.5)

    def test_gap_ratio_out_of_range_low(self):
        with pytest.raises(ValueError, match="between"):
            _build_option_list("base", gap_ratio=MIN_GAP_RATIO - 0.5)

    def test_gap_ratio_nan_raises(self):
        with pytest.raises(ValueError, match="finite"):
            _build_option_list("base", gap_ratio=float("nan"))

    def test_gap_ratio_inf_raises(self):
        with pytest.raises(ValueError, match="finite"):
            _build_option_list("base", gap_ratio=float("inf"))

    def test_gap_ratio_bool_raises(self):
        # bool is a subclass of int; reject it so True does not become 1.
        with pytest.raises(ValueError, match="finite"):
            _build_option_list("base", gap_ratio=True)


# ---------------------------------------------------------------------------
# parse_palmer_options  (user-typed command → settings)
# ---------------------------------------------------------------------------

class TestParsePalmerOptions:

    def test_none(self):
        assert parse_palmer_options(None) == {
            "align": "base", "no_vert": False,
            "no_reverse": False, "gap_ratio": None,
        }

    def test_empty_string(self):
        assert parse_palmer_options("")["align"] == "base"

    # --- v2 key=value syntax ---------------------------------------------

    def test_align_center(self):
        assert parse_palmer_options("align=center")["align"] == "center"

    def test_align_centre_british(self):
        assert parse_palmer_options("align=centre")["align"] == "center"

    def test_align_bottom(self):
        assert parse_palmer_options("align=bottom")["align"] == "bottom"

    def test_no_vert_bare(self):
        assert parse_palmer_options("no-vert")["no_vert"] is True

    def test_no_vert_true(self):
        assert parse_palmer_options("no-vert=true")["no_vert"] is True

    def test_no_vert_false(self):
        assert parse_palmer_options("no-vert=false")["no_vert"] is False

    def test_no_reverse_bare(self):
        assert parse_palmer_options("no-reverse")["no_reverse"] is True

    def test_gap_ratio(self):
        assert parse_palmer_options("gap-ratio=0.5")["gap_ratio"] == 0.5

    def test_gap_ratio_clamped_high(self):
        assert parse_palmer_options("gap-ratio=1.5")["gap_ratio"] == MAX_GAP_RATIO

    def test_gap_ratio_clamped_low(self):
        assert parse_palmer_options("gap-ratio=-2")["gap_ratio"] == MIN_GAP_RATIO

    def test_gap_ratio_non_finite_ignored(self):
        assert parse_palmer_options("gap-ratio=inf")["gap_ratio"] is None

    def test_gap_ratio_unparseable_ignored(self):
        assert parse_palmer_options("gap-ratio=abc")["gap_ratio"] is None

    def test_combined(self):
        result = parse_palmer_options("align=center, no-vert, no-reverse, gap-ratio=0.25")
        assert result == {
            "align": "center", "no_vert": True,
            "no_reverse": True, "gap_ratio": 0.25,
        }

    def test_whitespace_tolerated(self):
        result = parse_palmer_options("  align = center ,  no-vert  ")
        assert result["align"] == "center"
        assert result["no_vert"] is True

    def test_unknown_key_ignored(self):
        # palmer.sty is the authority on validity; the converter just ignores
        # keys it does not recognise rather than failing the command.
        assert parse_palmer_options("frobnicate=7")["align"] == "base"

    def test_bare_alignment_keyword_ignored(self):
        # v2 requires align=<value>; a bare `center` token is not a recognised
        # key and leaves the default alignment unchanged.
        assert parse_palmer_options("center")["align"] == "base"
