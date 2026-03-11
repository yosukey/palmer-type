"""Tests for _strip_tex_commands and _expand_ranges.

Covers:
  - TeX command stripping (single, nested, bare commands, braces/spaces)
  - Range expansion (digit ranges, letter ranges, invalid ranges)
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from palmer_converter import _strip_tex_commands, _expand_ranges


# ---------------------------------------------------------------------------
# _strip_tex_commands
# ---------------------------------------------------------------------------

class TestStripTexCommands:

    def test_plain_text(self):
        assert _strip_tex_commands("123") == "123"

    def test_textbf(self):
        assert _strip_tex_commands(r"\textbf{1}") == "1"

    def test_nested_commands(self):
        assert _strip_tex_commands(r"\textbf{\underline{AB}}") == "AB"

    def test_command_with_extra_text(self):
        assert _strip_tex_commands(r"\textbf{1}2") == "12"

    def test_bare_command_removed(self):
        assert _strip_tex_commands(r"\relax") == ""

    def test_braces_and_spaces_stripped(self):
        assert _strip_tex_commands("{ 1 }") == "1"

    def test_empty_input(self):
        assert _strip_tex_commands("") == ""

    def test_multiple_commands(self):
        assert _strip_tex_commands(r"\textbf{A}\textit{B}") == "AB"


# ---------------------------------------------------------------------------
# _expand_ranges
# ---------------------------------------------------------------------------

class TestExpandRanges:

    def test_digit_range(self):
        assert _expand_ranges("1-4") == "1234"

    def test_letter_range(self):
        assert _expand_ranges("A-C") == "ABC"

    def test_single_digit(self):
        assert _expand_ranges("1") == "1"

    def test_no_range(self):
        assert _expand_ranges("123") == "123"

    def test_full_range(self):
        assert _expand_ranges("1-8") == "12345678"

    def test_full_letter_range(self):
        assert _expand_ranges("A-E") == "ABCDE"

    def test_mixed(self):
        assert _expand_ranges("1-3A-C") == "123ABC"

    def test_bare_and_range(self):
        assert _expand_ranges("51-3") == "5123"

    def test_invalid_range_raises(self):
        with pytest.raises(ValueError, match="Invalid range"):
            _expand_ranges("4-1")

    def test_empty(self):
        assert _expand_ranges("") == ""
