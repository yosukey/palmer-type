"""Tests for _build_alt_text — alt-text generation for Palmer commands.

Covers:
  - FDI notation mode
  - Universal numbering mode
  - Anatomical naming mode
  - Alphanumeric mode
  - Empty quadrants
  - TeX-decorated fields
  - Range expansion (1-4, A-C)
  - Midline dash expansion
  - Invalid mode rejection
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from palmer_converter import _build_alt_text


def _cmd(UL="", UR="", LR="", LL="", upper_mid="", lower_mid="", option=None):
    """Helper to build a command dict matching find_palmer_commands output."""
    return {
        "UL": UL, "UR": UR, "LR": LR, "LL": LL,
        "upper_mid": upper_mid, "lower_mid": lower_mid,
        "option": option,
    }


# ---------------------------------------------------------------------------
# FDI mode
# ---------------------------------------------------------------------------

class TestFDI:

    def test_single_tooth_ul(self):
        result = _build_alt_text(_cmd(UL="1"), "FDI")
        assert "11" in result  # quadrant 1, tooth 1

    def test_single_tooth_ur(self):
        result = _build_alt_text(_cmd(UR="3"), "FDI")
        assert "23" in result  # quadrant 2, tooth 3

    def test_deciduous(self):
        result = _build_alt_text(_cmd(UL="A"), "FDI")
        assert "51" in result  # deciduous quadrant 5, tooth 1

    def test_multiple_teeth(self):
        result = _build_alt_text(_cmd(UL="12"), "FDI")
        assert "11" in result
        assert "12" in result

    def test_novert_returns_empty(self):
        """FDI numbers encode the quadrant; novert makes side unknown → empty."""
        result = _build_alt_text(
            _cmd(UL="1", upper_mid="novert"), "FDI",
        )
        assert result == ""


# ---------------------------------------------------------------------------
# Universal mode
# ---------------------------------------------------------------------------

class TestUniversal:

    def test_single_tooth(self):
        result = _build_alt_text(_cmd(UL="1"), "Universal")
        assert "8" in result  # UL tooth 1 → Universal #8

    def test_novert_returns_empty(self):
        """Universal numbers are side-specific; novert → empty."""
        result = _build_alt_text(
            _cmd(UL="1", LL="3", upper_mid="novert"), "Universal",
        )
        assert result == ""

    def test_deciduous(self):
        result = _build_alt_text(_cmd(UL="A"), "Universal")
        assert "E" in result  # UL tooth A → Universal E


# ---------------------------------------------------------------------------
# Anatomical mode
# ---------------------------------------------------------------------------

class TestAnatomical:

    def test_central_incisor(self):
        result = _build_alt_text(_cmd(UL="1"), "Anatomical")
        assert "Central incisor" in result
        assert "Right Maxillary" in result

    def test_third_molar(self):
        result = _build_alt_text(_cmd(LR="8"), "Anatomical")
        assert "Third molar" in result
        assert "Left Mandibular" in result

    def test_novert_removes_right_left(self):
        """novert: use 'Maxillary'/'Mandibular' without Right/Left."""
        result = _build_alt_text(
            _cmd(UL="1", LL="8", upper_mid="novert", lower_mid="novert"),
            "Anatomical",
        )
        assert "Maxillary Central incisor" in result
        assert "Mandibular Third molar" in result
        assert "Right" not in result
        assert "Left" not in result


# ---------------------------------------------------------------------------
# Alphanumeric mode
# ---------------------------------------------------------------------------

class TestAlphanumeric:

    def test_basic(self):
        # Engine UL = patient's UR (Quadrant 1)
        result = _build_alt_text(_cmd(UL="1"), "Alphanumeric")
        assert "UR1" in result

    def test_multiple(self):
        # Engine UR = patient's UL (Quadrant 2)
        result = _build_alt_text(_cmd(UR="12"), "Alphanumeric")
        assert "UL1" in result
        assert "UL2" in result

    def test_novert_upper(self):
        """novert: UL field → 'U' prefix (no left/right distinction)."""
        result = _build_alt_text(
            _cmd(UL="1", upper_mid="novert", lower_mid="novert"),
            "Alphanumeric",
        )
        assert "U1" in result
        assert "UL1" not in result

    def test_novert_lower(self):
        """novert: LL field → 'L' prefix (no left/right distinction)."""
        result = _build_alt_text(
            _cmd(LL="3", upper_mid="novert", lower_mid="novert"),
            "Alphanumeric",
        )
        assert "L3" in result
        assert "LL3" not in result

    def test_novert_upper_and_lower(self):
        """novert: both upper and lower produce U/L prefixes."""
        result = _build_alt_text(
            _cmd(UL="12", LL="45", upper_mid="novert", lower_mid="novert"),
            "Alphanumeric",
        )
        assert "U1" in result
        assert "U2" in result
        assert "L4" in result
        assert "L5" in result
        assert "UL" not in result
        assert "LL" not in result

    def test_novert_no_midline_suffix(self):
        """novert: 'novert' should not appear in the midline suffix."""
        result = _build_alt_text(
            _cmd(UL="1", upper_mid="novert", lower_mid="novert"),
            "Alphanumeric",
        )
        assert "novert" not in result

    def test_novert_upper_mid_only(self):
        """novert in upper_mid only still activates novert mode."""
        result = _build_alt_text(
            _cmd(UL="1", upper_mid="novert"),
            "Alphanumeric",
        )
        assert "U1" in result
        assert "UL1" not in result
        assert "novert" not in result

    def test_novert_lower_mid_only(self):
        """novert in lower_mid only still activates novert mode."""
        result = _build_alt_text(
            _cmd(LL="2", lower_mid="novert"),
            "Alphanumeric",
        )
        assert "L2" in result
        assert "LL2" not in result
        assert "novert" not in result


# ---------------------------------------------------------------------------
# Edge cases
# ---------------------------------------------------------------------------

class TestEdgeCases:

    def test_empty_quadrants(self):
        result = _build_alt_text(_cmd(), "FDI")
        assert result == ""

    def test_tex_commands_stripped(self):
        result = _build_alt_text(_cmd(UL=r"\textbf{1}"), "FDI")
        assert "11" in result

    def test_range_expansion(self):
        # Engine UL = patient's UR
        result = _build_alt_text(_cmd(UL="1-3"), "Alphanumeric")
        assert "UR1" in result
        assert "UR2" in result
        assert "UR3" in result

    def test_midline_dash_expansion(self):
        """When upper_mid is a dash, UL='3' means teeth 1, 2, 3."""
        # Engine UL = patient's UR
        result = _build_alt_text(_cmd(UL="3", upper_mid="-"), "Alphanumeric")
        assert "UR1" in result
        assert "UR2" in result
        assert "UR3" in result

    def test_midline_dash_multi_char_perm(self):
        """Multi-char permanent teeth: prepend consecutive teeth from midline."""
        # LR="46" → "12346" (prepend 1,2,3); LL="358" → "12358" (prepend 1,2)
        result = _build_alt_text(
            _cmd(LR="46", LL="358", lower_mid="-"), "Alphanumeric",
        )
        # Engine LL = patient's LR; Engine LR = patient's LL
        # LR (from LL="358"): should have 1,2,3,5,8
        assert "LR1" in result
        assert "LR2" in result
        assert "LR3" in result
        assert "LR5" in result
        assert "LR8" in result
        assert "LR4" not in result  # gap not filled
        # LL (from LR="46"): should have 1,2,3,4,6
        assert "LL1" in result
        assert "LL2" in result
        assert "LL3" in result
        assert "LL4" in result
        assert "LL6" in result
        assert "LL5" not in result  # gap not filled

    def test_midline_dash_multi_char_perm_no_prefix(self):
        """Multi-char permanent teeth already containing 1: no change."""
        result = _build_alt_text(
            _cmd(LR="13", lower_mid="-"), "Alphanumeric",
        )
        assert "LL1" in result
        assert "LL3" in result
        assert "LL2" not in result  # gap not filled

    def test_midline_dash_multi_char_decid(self):
        """Multi-char deciduous teeth: only prepend 'A' if missing."""
        # UL="BC" → "ABC"; UR="BE" → "ABE" (C, D not filled)
        result = _build_alt_text(
            _cmd(UL="BC", UR="BE", upper_mid="\u2014"), "Alphanumeric",
        )
        # Engine UL = patient's UR; Engine UR = patient's UL
        assert "URA" in result
        assert "URB" in result
        assert "URC" in result
        assert "ULA" in result
        assert "ULB" in result
        assert "ULE" in result
        # C and D should NOT appear in UL (gap not filled)
        assert "ULC" not in result
        assert "ULD" not in result

    def test_invalid_mode_raises(self):
        with pytest.raises(ValueError, match="Unknown alt-text mode"):
            _build_alt_text(_cmd(UL="1"), "InvalidMode")

    def test_midline_text_appended(self):
        result = _build_alt_text(_cmd(UL="1", upper_mid="+"), "Alphanumeric")
        assert "upper midline: +" in result
