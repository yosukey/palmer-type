"""Tests for input validation functions in palmer_engine.

Covers:
  - _check_brace_balance: balanced, unbalanced, escaped braces
  - _check_no_dangerous_cmds: safe inputs, dangerous commands, ^^ escape
  - _validate_color: hex colors, named colors, invalid
  - _get_font_preamble: known fonts, custom fonts, empty, unsafe chars
  - validate_raw_input: length limits
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from palmer_engine import (
    _check_brace_balance,
    _check_no_dangerous_cmds,
    _validate_color,
    _get_font_preamble,
    validate_raw_input,
    MAX_FIELD_LEN,
    MAX_RAW_LEN,
    FONT_PACKAGES,
)


# ---------------------------------------------------------------------------
# _check_brace_balance
# ---------------------------------------------------------------------------

class TestBraceBalance:

    def test_balanced_simple(self):
        _check_brace_balance("{hello}", "test")

    def test_balanced_nested(self):
        _check_brace_balance("{a{b}c}", "test")

    def test_balanced_empty(self):
        _check_brace_balance("", "test")

    def test_no_braces(self):
        _check_brace_balance("plain text", "test")

    def test_unmatched_closing(self):
        with pytest.raises(ValueError, match="unmatched closing brace"):
            _check_brace_balance("a}", "test")

    def test_unmatched_opening(self):
        with pytest.raises(ValueError, match="unmatched opening"):
            _check_brace_balance("{a", "test")

    def test_escaped_braces_ignored(self):
        """Escaped braces should not affect depth count."""
        _check_brace_balance(r"\{literal\}", "test")

    def test_mixed_escaped_and_real(self):
        _check_brace_balance(r"{\{inner\}}", "test")


# ---------------------------------------------------------------------------
# _check_no_dangerous_cmds
# ---------------------------------------------------------------------------

class TestDangerousCmds:

    def test_safe_input(self):
        _check_no_dangerous_cmds(r"\textbf{1}", "test")

    def test_safe_plain(self):
        _check_no_dangerous_cmds("12345", "test")

    def test_dangerous_write(self):
        with pytest.raises(ValueError, match="disallowed"):
            _check_no_dangerous_cmds(r"\write18{whoami}", "test")

    def test_dangerous_def(self):
        with pytest.raises(ValueError, match="disallowed"):
            _check_no_dangerous_cmds(r"\def\x{bad}", "test")

    def test_dangerous_input(self):
        with pytest.raises(ValueError, match="disallowed"):
            _check_no_dangerous_cmds(r"\input{secrets.tex}", "test")

    def test_dangerous_csname(self):
        with pytest.raises(ValueError, match="disallowed"):
            _check_no_dangerous_cmds(r"\csname write\endcsname", "test")

    def test_caret_escape(self):
        """^^ notation is rejected to prevent bypass."""
        with pytest.raises(ValueError, match="\\^\\^"):
            _check_no_dangerous_cmds("^^41", "test")

    def test_word_boundary(self):
        """\\typewriter should NOT match \\write (word boundary aware)."""
        _check_no_dangerous_cmds(r"\typewriter", "test")


# ---------------------------------------------------------------------------
# _validate_color
# ---------------------------------------------------------------------------

class TestValidateColor:

    def test_empty_is_ok(self):
        _validate_color("")

    def test_hex_valid(self):
        _validate_color("#FF0000")
        _validate_color("#00ff00")

    def test_named_valid(self):
        _validate_color("red")
        _validate_color("darkgray")

    def test_invalid_hex(self):
        with pytest.raises(ValueError, match="Invalid color"):
            _validate_color("#GGGGGG")

    def test_invalid_short_hex(self):
        with pytest.raises(ValueError, match="Invalid color"):
            _validate_color("#FFF")

    def test_invalid_chars(self):
        with pytest.raises(ValueError, match="Invalid color"):
            _validate_color("not a color!")


# ---------------------------------------------------------------------------
# _get_font_preamble
# ---------------------------------------------------------------------------

class TestGetFontPreamble:

    def test_known_font(self):
        result = _get_font_preamble("Times New Roman")
        assert result == FONT_PACKAGES["Times New Roman"]

    def test_custom_font(self):
        result = _get_font_preamble("Comic Sans MS")
        assert result == r"\setmainfont{Comic Sans MS}"

    def test_empty_name_raises(self):
        with pytest.raises(ValueError, match="must not be empty"):
            _get_font_preamble("")

    def test_whitespace_only_raises(self):
        with pytest.raises(ValueError, match="must not be empty"):
            _get_font_preamble("   ")

    def test_unsafe_chars_raises(self):
        with pytest.raises(ValueError, match="unsupported characters"):
            _get_font_preamble("Bad{Font}")

    def test_too_long_raises(self):
        with pytest.raises(ValueError, match="maximum length"):
            _get_font_preamble("A" * (MAX_FIELD_LEN + 1))


# ---------------------------------------------------------------------------
# validate_raw_input
# ---------------------------------------------------------------------------

class TestValidateRawInput:

    def test_valid(self):
        validate_raw_input(r"\Palmer{1}{}{}{}{}{}", "test")

    def test_too_long(self):
        with pytest.raises(ValueError, match="maximum length"):
            validate_raw_input("x" * (MAX_RAW_LEN + 1), "test")

    def test_dangerous_rejected(self):
        with pytest.raises(ValueError, match="disallowed"):
            validate_raw_input(r"\input{file}", "test")

    def test_unbalanced_rejected(self):
        with pytest.raises(ValueError, match="unmatched"):
            validate_raw_input("{unclosed", "test")
