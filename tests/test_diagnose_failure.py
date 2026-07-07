"""Tests for _diagnose_tex_failure in palmer_engine."""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from palmer_engine import _diagnose_tex_failure


def _has_kernel_hint(hints: list[str]) -> bool:
    return any("predates 2022-06-01" in h for h in hints)


def _has_fontconfig_hint(hints: list[str]) -> bool:
    return any("harmless start-up warning" in h for h in hints)


# The exact stderr Tectonic produced for the reported failure.
TASK_FAILURE = (
    "Fontconfig error: Cannot load default config file: No such file: (null)\n"
    "error: palmer.sty:256: Undefined control sequence\n"
    "error: halted on potentially-recoverable error as specified\n"
)


class TestOutdatedKernel:

    def test_task_failure_reports_both_hints(self):
        hints = _diagnose_tex_failure(TASK_FAILURE)
        assert _has_kernel_hint(hints)
        assert _has_fontconfig_hint(hints)

    def test_tectonic_stderr_signature(self):
        """Tectonic reports only file:line (no macro name) on stderr."""
        hints = _diagnose_tex_failure("error: palmer.sty:256: Undefined control sequence")
        assert _has_kernel_hint(hints)

    def test_xelatex_log_signature(self):
        """XeLaTeX's log names the macro rather than the .sty file."""
        log_tail = (
            "! Undefined control sequence.\n"
            "<recently read> \\DeclareKeys \n"
            "l.256 \\DeclareKeys[palmer]{\n"
        )
        hints = _diagnose_tex_failure(log_tail)
        assert _has_kernel_hint(hints)

    def test_processkeyoptions_signature(self):
        hints = _diagnose_tex_failure(
            "! Undefined control sequence.\nl.263 \\ProcessKeyOptions[palmer]"
        )
        assert _has_kernel_hint(hints)

    def test_case_insensitive(self):
        hints = _diagnose_tex_failure("PALMER.STY:256: UNDEFINED CONTROL SEQUENCE")
        assert _has_kernel_hint(hints)


class TestNoFalsePositives:

    def test_user_error_in_main_document_is_not_flagged(self):
        """An undefined macro in the user's document is input error, not old kernel."""
        hints = _diagnose_tex_failure(
            "error: palmer.tex:8: Undefined control sequence\n"
            "l.8 \\PalmerTypoCommand\n"
        )
        assert not _has_kernel_hint(hints)
        assert hints == []

    def test_unrelated_failure_returns_no_hints(self):
        hints = _diagnose_tex_failure(
            "error: palmer.tex:10: File `nosuchfont.otf' not found."
        )
        assert hints == []

    def test_empty_input_returns_no_hints(self):
        assert _diagnose_tex_failure("") == []
        assert _diagnose_tex_failure(None) == []  # type: ignore[arg-type]


class TestFontconfig:

    def test_fontconfig_only_reports_only_that_hint(self):
        hints = _diagnose_tex_failure(
            "Fontconfig error: Cannot load default config file: No such file: (null)\n"
            "error: something unrelated went wrong"
        )
        assert _has_fontconfig_hint(hints)
        assert not _has_kernel_hint(hints)
