"""Tests for find_palmer_commands — Palmer command parser.

Covers:
  - Basic single command detection
  - Multiple commands in one string
  - Optional [option] argument
  - Nested braces in arguments
  - Escaped braces
  - Yen sign (U+00A5) normalisation
  - Malformed input (missing braces, incomplete args)
  - Empty input
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from palmer_converter import find_palmer_commands


# ---------------------------------------------------------------------------
# Basic parsing
# ---------------------------------------------------------------------------

def test_single_command():
    text = r"\Palmer{1}{2}{3}{4}{5}{6}"
    cmds = find_palmer_commands(text)
    assert len(cmds) == 1
    cmd = cmds[0]
    assert cmd["UL"] == "1"
    assert cmd["UR"] == "2"
    assert cmd["LR"] == "3"
    assert cmd["LL"] == "4"
    assert cmd["upper_mid"] == "5"
    assert cmd["lower_mid"] == "6"
    assert cmd["option"] is None
    assert cmd["start"] == 0
    assert cmd["end"] == len(text)


def test_command_with_option():
    text = r"\Palmer[center]{A}{B}{C}{D}{}{}"
    cmds = find_palmer_commands(text)
    assert len(cmds) == 1
    assert cmds[0]["option"] == "center"
    assert cmds[0]["UL"] == "A"


def test_multiple_commands():
    text = r"pre \Palmer{1}{}{}{}{}{} mid \Palmer{}{2}{}{}{}{} post"
    cmds = find_palmer_commands(text)
    assert len(cmds) == 2
    assert cmds[0]["UL"] == "1"
    assert cmds[1]["UR"] == "2"


def test_empty_arguments():
    text = r"\Palmer{}{}{}{}{}{}"
    cmds = find_palmer_commands(text)
    assert len(cmds) == 1
    for key in ("UL", "UR", "LR", "LL", "upper_mid", "lower_mid"):
        assert cmds[0][key] == ""


# ---------------------------------------------------------------------------
# Braces and escaping
# ---------------------------------------------------------------------------

def test_nested_braces():
    text = r"\Palmer{\textbf{1}}{}{}{}{}{}"
    cmds = find_palmer_commands(text)
    assert len(cmds) == 1
    assert cmds[0]["UL"] == r"\textbf{1}"


def test_escaped_braces():
    text = r"\Palmer{\{1\}}{}{}{}{}{}"
    cmds = find_palmer_commands(text)
    assert len(cmds) == 1
    assert cmds[0]["UL"] == r"\{1\}"


# ---------------------------------------------------------------------------
# Yen sign normalisation
# ---------------------------------------------------------------------------

def test_yen_sign_prefix():
    """U+00A5 (yen sign) should be treated as a backslash."""
    text = "\u00a5Palmer{1}{2}{3}{4}{5}{6}"
    cmds = find_palmer_commands(text)
    assert len(cmds) == 1
    assert cmds[0]["UL"] == "1"


# ---------------------------------------------------------------------------
# Malformed input
# ---------------------------------------------------------------------------

def test_no_commands():
    assert find_palmer_commands("no commands here") == []


def test_empty_input():
    assert find_palmer_commands("") == []


def test_incomplete_args():
    """Only 3 braced arguments — should not match."""
    text = r"\Palmer{1}{2}{3}"
    cmds = find_palmer_commands(text)
    assert len(cmds) == 0


def test_unclosed_brace():
    """Unclosed brace in the middle — should not match."""
    text = r"\Palmer{1}{2}{3{}{}{}"
    cmds = find_palmer_commands(text)
    assert len(cmds) == 0


def test_whitespace_between_args():
    """Whitespace between braced groups is tolerated."""
    text = "\\Palmer{1} {2}\n{3}\t{4} {5} {6}"
    cmds = find_palmer_commands(text)
    assert len(cmds) == 1
    assert cmds[0]["UL"] == "1"
    assert cmds[0]["lower_mid"] == "6"


# ---------------------------------------------------------------------------
# Offset tracking
# ---------------------------------------------------------------------------

def test_start_end_offsets():
    """start/end should point to the exact substring in the original text."""
    prefix = "hello "
    command = r"\Palmer{a}{b}{c}{d}{e}{f}"
    text = prefix + command + " world"
    cmds = find_palmer_commands(text)
    assert len(cmds) == 1
    assert text[cmds[0]["start"]: cmds[0]["end"]] == command
