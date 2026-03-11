"""
palmer_type.py - Palmer Dental Notation Type (GUI)

Renders Palmer dental notation symbols and copies the result to the
clipboard for pasting into Word.

TeX engine (auto-detected in priority order):
  - Bundled / adjacent tectonic
  - System tectonic on PATH
  - xelatex on PATH (TeX Live / MiKTeX)

Other requirements:
  - palmer.sty  (bundled)
"""

from __future__ import annotations

import datetime
import functools
import json
import logging
import os
import platform
import re
import socket
import subprocess
import sys
import threading
import tkinter as tk
import webbrowser
import tkinter.font as tkfont
from urllib.request import urlopen, Request
from urllib.error import URLError
from tkinter import ttk, filedialog, messagebox, colorchooser
from tkinter.scrolledtext import ScrolledText
from pathlib import Path
from typing import Callable

logger = logging.getLogger(__name__)

from PIL import Image, ImageTk, ImageOps

from palmer_engine import (
    PalmerCompiler, FONT_PACKAGES, tectonic_cache_exists, delete_tectonic_cache, clamp_dpi,
    MIN_DPI, MAX_DPI,
    MIN_FONT_SIZE_PT, MAX_FONT_SIZE_PT, DEFAULT_FONT_SIZE_PT,
    DEFAULT_MARGIN_PX,
    MM_PER_INCH,
)
from config import AppConfig
from version import __version__

# ---------------------------------------------------------------------------
# UI layout constants
# ---------------------------------------------------------------------------
_WHITE: tuple[int, int, int] = (255, 255, 255)          # white background color
_DEFAULT_BG_COLOR: tuple[int, int, int] = (192, 192, 192)  # default gray background
_PROGRESS_BAR_INTERVAL_MS: int = 15                      # tkinter progress bar poll interval (ms)
_MIN_WINDOW_SIZE: int = 600                              # minimum window width/height (pixels)
_MIN_CANVAS_PX: int = 100                                # minimum canvas dimension before first map (pixels)
_FONT_SIZE_DISPLAY: int = 9                              # Consolas font size for log display
_UI_HEADER_FONT_SIZE: int = 14                           # header label font size

_REPO_URL: str = "https://github.com/yosukey/palmer-type"
_RELEASES_API_URL: str = "https://api.github.com/repos/yosukey/palmer-type/releases/latest"
_RELEASES_URL: str = f"{_REPO_URL}/releases/latest"


def _check_online(timeout: float = 5.0) -> bool:
    """Return True if an internet connection appears to be available.

    Tries several well-known hosts on different ports so that environments
    which block outbound UDP/TCP port 53 (DNS) but allow HTTPS still pass.

    All targets use raw IP addresses to avoid DNS resolution hangs — on some
    platforms ``getaddrinfo()`` is **not** bounded by the socket timeout,
    which could block the calling thread indefinitely.
    """
    checks = [
        ("8.8.8.8", 53),        # Google Public DNS
        ("1.1.1.1", 443),       # Cloudflare HTTPS
        ("9.9.9.9", 53),        # Quad9 DNS
    ]
    per_timeout = max(1.0, timeout / len(checks))
    logger.debug("_check_online: starting (timeout=%.1fs, per_timeout=%.1fs)",
                 timeout, per_timeout)
    for host, port in checks:
        try:
            logger.debug("_check_online: trying %s:%d ...", host, port)
            with socket.create_connection((host, port), timeout=per_timeout):
                logger.debug("_check_online: %s:%d connected — online", host, port)
                return True
        except OSError as exc:
            logger.debug("_check_online: %s:%d failed (%s)", host, port, exc)
            continue
    logger.debug("_check_online: all checks failed — offline")
    return False


def _get_platform_str() -> str:
    """Return a human-readable platform string.

    On Windows, correctly distinguishes Windows 11 (build >= 22000) from
    Windows 10, since platform.release() returns "10" for both.
    """
    if platform.system() == "Windows":
        try:
            ver = sys.getwindowsversion()  # type: ignore[attr-defined]
            build = ver.build
            if ver.major == 10 and build >= 22000:
                return f"Windows 11 (build {build})"
            return f"Windows {ver.major}.{ver.minor} (build {build})"
        except Exception:
            pass
    return f"{platform.system()} {platform.release()}"


def _read_fontlink_consolas() -> list[str]:
    """Return raw FontLink\\SystemLink entries for Consolas on Windows.

    Each element is a raw registry string in ``"FILENAME.TTC,Font Name"``
    format (or just a filename when no face name is embedded).
    Returns an empty list on non-Windows systems or when the key is absent.
    """
    if platform.system() != "Windows":
        return []
    try:
        import importlib
        winreg = importlib.import_module("winreg")
        key = winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE,
            r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\FontLink\SystemLink",
        )
        try:
            value, _ = winreg.QueryValueEx(key, "Consolas")
            return list(value) if isinstance(value, list) else [str(value)]
        finally:
            winreg.CloseKey(key)
    except FileNotFoundError:
        return []  # キーまたは値が存在しない — FontLink 未設定の正常な状態
    except Exception as exc:
        logger.debug("Windows FontLink lookup failed: %s", exc)
        return []


def _detect_cjk_fallback_font() -> str:
    """Detect the CJK fallback font the OS would use for Consolas.

    Returns a human-readable string describing the detected fallback font(s),
    or a diagnostic message if detection fails.
    """
    system = platform.system()

    # --- Windows: read FontLink\\SystemLink registry ---
    if system == "Windows":
        entries = _read_fontlink_consolas()
        fonts: list[str] = []
        for entry in entries:
            # Format: "MSGOTHIC.TTC,MS Gothic" or just "MSGOTHIC.TTC"
            parts = entry.split(",")
            if len(parts) >= 2:
                fonts.append(parts[1].strip())
            else:
                fonts.append(parts[0].strip())
        if fonts:
            return "; ".join(fonts) + "  (via FontLink\\SystemLink)"
        return "(DirectWrite render-time fallback — see debug detail)"

    # --- Linux / macOS: use fontconfig ---
    if system in ("Linux", "Darwin"):
        try:
            result = subprocess.run(
                ["fc-match", "-s", "Consolas:lang=ja", "family"],
                capture_output=True, text=True, timeout=5,
            )
            if result.returncode == 0 and result.stdout.strip():
                for line in result.stdout.splitlines():
                    name = line.strip()
                    if name and name.lower() != "consolas":
                        # Take only the primary family name (before commas).
                        return name.split(",")[0].strip()
        except Exception as exc:
            logger.debug("fc-match failed: %s", exc)

        if system == "Darwin":
            return "(macOS Core Text auto-fallback; likely Hiragino Sans)"

    return "(could not detect)"


def _dump_font_fallback_detail() -> str:
    """Return a multi-line diagnostic string showing the full font fallback chain.

    On Windows, dumps all FontLink\\SystemLink entries for Consolas.
    On Linux, shows the top fc-match -s results.
    Intended for the debug log only.
    """
    system = platform.system()
    lines: list[str] = []

    if system == "Windows":
        entries = _read_fontlink_consolas()
        if entries:
            lines.append(f"FontLink\\SystemLink for Consolas ({len(entries)} entries):")
            for i, entry in enumerate(entries, 1):
                lines.append(f"  {i}. {entry}")
        else:
            pass  # No FontLink entry — DirectWrite handles fallback at render time

    elif system in ("Linux", "Darwin"):
        try:
            result = subprocess.run(
                ["fc-match", "-s", "Consolas:lang=ja", "family"],
                capture_output=True, text=True, timeout=5,
            )
            if result.returncode == 0:
                fc_lines = result.stdout.strip().splitlines()[:8]
                lines.append(f"fc-match -s fallback chain (top {len(fc_lines)}):")
                for i, fl in enumerate(fc_lines, 1):
                    lines.append(f"  {i}. {fl.strip()}")
        except Exception as exc:
            lines.append(f"fc-match: {exc}")

    return "\n".join(lines) if lines else "(no detail available)"


# ---------------------------------------------------------------------------
# Global CJK font setup
# ---------------------------------------------------------------------------

# Matches one or more consecutive CJK / full-width characters.
_CJK_RE = re.compile(
    r'[\u1100-\u11FF'   # Hangul Jamo
    r'\u2E80-\u2FFF'    # CJK Radicals Supplement & Kangxi Radicals
    r'\u3000-\u9FFF'    # CJK Unified Ideographs and adjacent CJK blocks
    r'\uA000-\uA4CF'    # Yi Syllables / Yi Radicals
    r'\uAC00-\uD7AF'    # Hangul Syllables
    r'\uF900-\uFAFF'    # CJK Compatibility Ideographs
    r'\uFE30-\uFE4F'    # CJK Compatibility Forms
    r'\uFF00-\uFFEF]+'  # Halfwidth and Fullwidth Forms
)


class _CjkScrolledText(ScrolledText):
    """ScrolledText variant that renders CJK characters with Segoe UI.

    At end-appends (``insert("end", ...)``), CJK runs are tagged with a
    ``"cjk"`` text tag configured to use Segoe UI at the widget's own point
    size.  Non-CJK text and insertions at other indices are handled by the
    parent class unchanged.

    Using a subclass rather than a class-level monkey-patch on ``tk.Text``
    confines the behaviour to widgets that explicitly opt in, avoiding
    side-effects on test suites or embedding scenarios.
    """

    def insert(  # type: ignore[override]
        self, index: str, chars: str, *args: str | list[str] | tuple[str, ...]
    ) -> None:
        # Only intercept simple end-appends of plain strings containing CJK.
        if index != "end" or not _CJK_RE.search(chars):
            super().insert(index, chars, *args)
            return

        # Lazily configure the "cjk" tag on this widget (first CJK insert).
        if "cjk" not in self.tag_names():
            try:
                size = abs(tkfont.Font(font=self.cget("font")).cget("size"))
            except Exception:
                size = 9
            self.tag_configure("cjk", font=("Segoe UI", size))

        # The first element of *args*, if a str or tuple, is a tag list.
        base_tags: tuple = ()
        if args and isinstance(args[0], (str, tuple)):
            t = args[0]
            base_tags = (t,) if isinstance(t, str) else tuple(t)

        # Split text into CJK / non-CJK runs and insert with appropriate tags.
        pos = 0
        for m in _CJK_RE.finditer(chars):
            if m.start() > pos:
                if base_tags:
                    super().insert("end", chars[pos:m.start()], base_tags)
                else:
                    super().insert("end", chars[pos:m.start()])
            super().insert("end", m.group(), base_tags + ("cjk",))
            pos = m.end()
        if pos < len(chars):
            if base_tags:
                super().insert("end", chars[pos:], base_tags)
            else:
                super().insert("end", chars[pos:])


_HAS_DOCX = False
_ALT_TEXT_OPTIONS: list[str] = []
try:
    from palmer_converter import convert_docx, ConversionCancelled, ALT_TEXT_MODES, VALIGN_MODES
    _HAS_DOCX = True
    # Build display labels for the alt-text combobox.
    _ALT_TEXT_OPTIONS = [
        "None" if m is None else m for m in ALT_TEXT_MODES
    ]
    _VALIGN_OPTIONS = list(VALIGN_MODES)
except ImportError as _docx_exc:
    _DOCX_IMPORT_ERROR: str | None = f"{type(_docx_exc).__name__}: {_docx_exc}"
    logger.debug("python-docx not available; Converter tab will be disabled", exc_info=True)
else:
    _DOCX_IMPORT_ERROR = None

logger.debug("_HAS_DOCX=%s", _HAS_DOCX)


class ConverterTab:
    """Encapsulates the Converter tab UI and logic.

    Separated from PalmerTypeApp to keep the main class focused on the
    Palmer notation input/render workflow.
    """

    def __init__(self, root: tk.Tk, status_var: tk.StringVar,
                 get_compiler: Callable[[], PalmerCompiler | None],
                 is_debug: Callable[[], bool] | None = None):
        self.root = root
        self.status_var = status_var
        self._get_compiler = get_compiler
        self._is_debug = is_debug or (lambda: False)

    def build(self, parent: ttk.Frame) -> None:
        """Build the Converter tab for replacing \\Palmer commands in .docx files."""
        if not _HAS_DOCX:
            msg = ttk.Label(
                parent,
                text=(
                    "The Converter tab requires python-docx.\n\n"
                    "Install it with:  pip install python-docx"
                ),
                foreground="#888888",
                wraplength=500,
                justify="center",
            )
            msg.pack(expand=True)
            return

        # ---- EXPERIMENTAL warning ----
        warn_frame = ttk.Frame(parent, relief="solid", borderwidth=1)
        warn_frame.pack(fill="x", padx=10, pady=(10, 5))
        ttk.Label(
            warn_frame,
            text="\u26a0 EXPERIMENTAL",
            font=("TkDefaultFont", 10, "bold"),
            foreground="#b35900",
        ).pack(anchor="w", padx=8, pady=(6, 2))
        ttk.Label(
            warn_frame,
            text=(
                "This converter has been carefully developed and tested, but the "
                "internal structure of .docx files is complex and complete conversion "
                "cannot be guaranteed. When prompted, saving to a new file (the default) "
                "is recommended to keep the original intact."
            ),
            foreground="#555555",
            wraplength=560,
            justify="left",
        ).pack(anchor="w", padx=8, pady=(0, 6))

        # ---- Input file ----
        input_frame = ttk.LabelFrame(parent, text="Input File", padding=10)
        input_frame.pack(fill="x", padx=10, pady=(5, 5))

        self._input_var = tk.StringVar()
        ttk.Entry(
            input_frame, textvariable=self._input_var, state="readonly",
        ).pack(side="left", fill="x", expand=True, padx=(0, 5))
        ttk.Button(
            input_frame, text="Browse...", command=self._on_browse,
        ).pack(side="left")

        # ---- DPI ----
        opt_frame = ttk.LabelFrame(parent, text="Options", padding=10)
        opt_frame.pack(fill="x", padx=10, pady=5)

        ttk.Label(opt_frame, text="DPI:").pack(side="left", padx=(0, 5))
        self._dpi_spin = ttk.Spinbox(
            opt_frame, from_=MIN_DPI, to=MAX_DPI, increment=1, width=6,
        )
        self._dpi_spin.set("600")
        self._dpi_spin.pack(side="left")
        ttk.Label(opt_frame, text="dpi").pack(side="left", padx=(3, 0))

        ttk.Separator(opt_frame, orient="vertical").pack(
            side="left", fill="y", padx=10, pady=2,
        )
        ttk.Label(opt_frame, text="Alt text:").pack(side="left", padx=(0, 5))
        self._alt_text_var = tk.StringVar(value="None")
        self._alt_text_combo = ttk.Combobox(
            opt_frame,
            textvariable=self._alt_text_var,
            values=_ALT_TEXT_OPTIONS,
            state="readonly",
            width=18,
        )
        self._alt_text_combo.pack(side="left")

        ttk.Label(
            opt_frame,
            text="  (Font and size are read from the Word document.)",
            foreground="gray",
        ).pack(side="left", padx=(10, 0))

        # ---- Inline vertical position ----
        valign_frame = ttk.LabelFrame(
            parent, text="Inline vertical position", padding=10,
        )
        valign_frame.pack(fill="x", padx=10, pady=5)

        self._valign_var = tk.StringVar(value="Force center")
        for val in _VALIGN_OPTIONS:
            ttk.Radiobutton(
                valign_frame, text=val,
                variable=self._valign_var, value=val,
            ).pack(side="left", padx=(0, 15))

        # ---- Convert / Stop button ----
        btn_frame = ttk.Frame(parent)
        btn_frame.pack(fill="x", padx=10, pady=5)
        self._convert_btn = ttk.Button(
            btn_frame, text="▶ Convert", command=self._on_convert,
        )
        self._convert_btn.pack(side="left", padx=5)
        self._stop_btn = ttk.Button(
            btn_frame, text="■ Stop", command=self._on_stop,
        )
        # Stop button is hidden until a conversion is running.
        self._stop_event = threading.Event()

        # ---- Progress ----
        self._progress = ttk.Progressbar(parent, mode="indeterminate", length=200)
        self._progress.pack(fill="x", padx=10, pady=(0, 2))
        self._progress.pack_forget()

        # ---- Log ----
        log_frame = ttk.LabelFrame(parent, text="Log", padding=5)
        log_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        self._log = _CjkScrolledText(
            log_frame, wrap="word", state="disabled",
            font=("Consolas", _FONT_SIZE_DISPLAY), height=8,
        )
        self._log.pack(fill="both", expand=True)

    def _on_browse(self) -> None:
        """Open file dialog to select a .docx input file."""
        path = filedialog.askopenfilename(
            filetypes=[("Word Document", "*.docx"), ("All files", "*.*")],
            title="Select a .docx file",
        )
        if path:
            self._input_var.set(path)

    def _log_append(self, text: str) -> None:
        """Append text to the converter log (thread-safe via root.after).

        Each non-empty line is prefixed with a timestamp.
        """
        if text.strip():
            ts = datetime.datetime.now().strftime("%H:%M:%S")
            line = f"[{ts}] {text}\n"
        else:
            line = "\n"
        def _append():
            self._log.configure(state="normal")
            self._log.insert("end", line)
            self._log.see("end")
            self._log.configure(state="disabled")
        self.root.after(0, _append)

    def _log_clear(self) -> None:
        """Clear the converter log."""
        self._log.configure(state="normal")
        self._log.delete("1.0", "end")
        self._log.configure(state="disabled")

    def _get_dpi(self) -> int:
        """Read the Converter DPI spinbox, falling back to 600."""
        return clamp_dpi(self._dpi_spin.get())

    def _on_stop(self) -> None:
        """Request cancellation of the running conversion."""
        self._stop_event.set()
        self._stop_btn.configure(state="disabled")
        self._log_append("Stop requested — finishing current paragraph ...")

    def set_convert_enabled(self, enabled: bool) -> None:
        """Enable or disable the Convert button (used during background downloads)."""
        if not hasattr(self, "_convert_btn"):
            return  # Converter tab not built (python-docx unavailable)
        self._convert_btn.configure(state="normal" if enabled else "disabled")

    def _set_busy(self, busy: bool) -> None:
        """Swap Convert↔Stop button and show/hide the progress bar."""
        if not hasattr(self, "_convert_btn"):
            return  # Converter tab not built (python-docx unavailable)
        if busy:
            self._convert_btn.pack_forget()
            self._stop_btn.configure(state="normal")
            self._stop_btn.pack(side="left", padx=5)
        else:
            self._stop_btn.pack_forget()
            self._convert_btn.pack(side="left", padx=5)
        if busy:
            self._progress.pack(fill="x", padx=10, pady=(0, 2))
            self._progress.start(_PROGRESS_BAR_INTERVAL_MS)
        else:
            self._progress.stop()
            self._progress.pack_forget()

    def _on_convert(self) -> None:
        r"""Validate input, ask for save mode, then run conversion in background."""
        compiler = self._get_compiler()
        if compiler is None:
            messagebox.showerror("Error", "TeX engine not available.")
            return

        input_path = self._input_var.get().strip()
        if not input_path:
            messagebox.showerror("Error", "Please select a .docx file first.")
            return

        src = Path(input_path)
        if not src.exists():
            messagebox.showerror("Error", f"File not found:\n{src}")
            return

        # Ask user for save mode.
        save_mode = messagebox.askyesnocancel(
            "Save Mode",
            f"How would you like to save the result?\n\n"
            f"[Yes] Save as a new file (recommended)\n"
            f"         → {src.stem}_palmered.docx\n\n"
            f"[No]  Overwrite the original file\n\n"
            f"[Cancel] Abort",
        )
        if save_mode is None:
            return

        if save_mode:
            default_name = f"{src.stem}_palmered.docx"
            output_path = filedialog.asksaveasfilename(
                defaultextension=".docx",
                filetypes=[("Word Document", "*.docx"), ("All files", "*.*")],
                title="Save converted document as",
                initialdir=str(src.parent),
                initialfile=default_name,
            )
            if not output_path:
                return
            dst = Path(output_path)
        else:
            confirm = messagebox.askyesno(
                "Confirm Overwrite",
                f"Are you sure you want to overwrite the original file?\n\n{src.name}",
            )
            if not confirm:
                return
            dst = src

        dpi = self._get_dpi()
        alt_sel = self._alt_text_var.get()
        alt_mode: str | None = None if alt_sel == "None" else alt_sel
        valign_mode = self._valign_var.get()

        self._log_clear()
        self._log_append(f"Input:  {src}")
        self._log_append(f"Output: {dst}")
        self._log_append(f"DPI:    {dpi}")
        if alt_mode:
            self._log_append(f"Alt text: {alt_mode}")
        self._log_append(f"Vertical position: {valign_mode}")
        is_debug = self._is_debug()
        if is_debug:
            self._log_append(f"Engine: {compiler.backend.name} ({compiler.backend.executable})")
            self._log_append(
                f"CJK fallback font (log widgets): {_detect_cjk_fallback_font()}")
        self._log_append("")
        self._stop_event.clear()
        self._set_busy(True)

        def _do_convert():
            try:
                replaced, errors = convert_docx(
                    input_path=src,
                    output_path=dst,
                    compiler=compiler,
                    dpi=dpi,
                    on_progress=self._log_append,
                    alt_text_mode=alt_mode,
                    valign_mode=valign_mode,
                    on_debug=self._log_append if is_debug else None,
                    stop_event=self._stop_event,
                )
                summary = f"\nDone — {replaced} command(s) replaced"
                if errors:
                    summary += f", {len(errors)} error(s):"
                    for err in errors:
                        summary += f"\n  • {err}"
                self._log_append(summary)
                status_msg = f"Converter: {replaced} replaced, {len(errors)} error(s)"
                self.root.after(0, functools.partial(self.status_var.set, status_msg))
                if not errors:
                    self.root.after(0, functools.partial(
                        messagebox.showinfo,
                        "Conversion Complete",
                        f"Replaced {replaced} Palmer command(s).\n\nSaved to:\n{dst}",
                    ))
                else:
                    self.root.after(0, functools.partial(
                        messagebox.showwarning,
                        "Conversion Complete (with errors)",
                        f"Replaced {replaced} command(s), {len(errors)} error(s).\n\n"
                        f"See the log for details.\n\nSaved to:\n{dst}",
                    ))
            except ConversionCancelled:
                self._log_append("\nCancelled — conversion stopped by user.")
                self.root.after(0, functools.partial(
                    self.status_var.set, "Converter: cancelled",
                ))
            except Exception as exc:
                self._log_append(f"\nFATAL ERROR: {exc}")
                err_msg = f"Converter error: {exc}"
                self.root.after(0, functools.partial(self.status_var.set, err_msg))
                self.root.after(0, functools.partial(
                    messagebox.showerror, "Conversion Failed", str(exc),
                ))
            finally:
                self.root.after(0, lambda: self._set_busy(False))

        threading.Thread(target=_do_convert, daemon=True).start()


class _StartupLogHandler(logging.Handler):
    """Logging handler that buffers messages during startup.

    Messages are stored in ``app._startup_log`` so they can be replayed into
    the GUI debug-log widget when the user enables debug mode.  While debug
    mode is already active the messages are forwarded directly to
    ``_debug_log_append``.
    """

    def __init__(self, app: "PalmerTypeApp"):
        super().__init__(level=logging.DEBUG)
        self._app = app

    def emit(self, record: logging.LogRecord) -> None:
        try:
            msg = self.format(record)
            ts = datetime.datetime.now().strftime("%H:%M:%S.%f")[:-3]
            line = f"[{ts}] {msg}"
            self._app._startup_log.append(line)
            # If debug mode is already on, also push to the widget directly.
            if self._app._debug_mode:
                self._app._debug_log_append(msg)
        except Exception:
            self.handleError(record)


class PalmerTypeApp:
    """Main window of the Palmer Dental Notation Type."""

    def __init__(self):
        if sys.platform == "win32":
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(  # type: ignore[attr-defined]
                "palmer-tool.palmer-type"
            )
        self.root = tk.Tk()
        self.root.title(f"palmer-type  v{__version__}")
        self.root.geometry("700x680")
        self.root.minsize(_MIN_WINDOW_SIZE, _MIN_WINDOW_SIZE)
        self._set_window_icon()

        self.compiler: PalmerCompiler | None = None
        self.current_image: Image.Image | None = None  # always white-bg RGB
        self.photo_image: ImageTk.PhotoImage | None = None
        self.custom_bg_color: tuple[int, int, int] = _DEFAULT_BG_COLOR
        self.custom_font_color: tuple[int, int, int] = (255, 0, 0)
        # Incremented on every input change (_mark_dirty); read by the render
        # callback (_on_done) to discard stale results.  Protected by a lock
        # so that the counter is safely visible across threads without relying
        # on CPython GIL guarantees.
        self._render_generation: int = 0
        self._gen_lock = threading.Lock()
        self._tectonic_downloading: bool = False  # True while auto-downloading tectonic files
        self._rendering: bool = False  # guard against concurrent render threads
        self._debug_mode: bool = False
        self._show_error: bool = False
        self.config = AppConfig()

        # Buffer startup debug messages so they can be replayed in the debug
        # log widget when the user enables debug mode later.
        self._startup_log: list[str] = []
        self._startup_log_handler = _StartupLogHandler(self)
        logging.getLogger().addHandler(self._startup_log_handler)
        logging.getLogger().setLevel(logging.DEBUG)

        # Write debug logs to a persistent file in the platform config dir so
        # they survive crashes and can be reviewed after the fact.
        _log_dir = AppConfig._default_config_dir() / "logs"
        try:
            _log_dir.mkdir(parents=True, exist_ok=True)
            # Auto-delete log files older than 92 days (≈3 months).
            _cutoff = datetime.datetime.now() - datetime.timedelta(days=92)
            for _old in _log_dir.glob("palmer_debug_*.log"):
                try:
                    if datetime.datetime.fromtimestamp(_old.stat().st_mtime) < _cutoff:
                        _old.unlink()
                except OSError:
                    pass
            _ts_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            self._debug_log_path: Path | None = _log_dir / f"palmer_debug_{_ts_str}.log"
            self._debug_log_dir: Path | None = _log_dir
            _fh = logging.FileHandler(self._debug_log_path, encoding="utf-8")
            _fh.setLevel(logging.DEBUG)
            _fh.setFormatter(logging.Formatter(
                "[%(asctime)s.%(msecs)03d] %(message)s", datefmt="%H:%M:%S"
            ))
            logging.getLogger().addHandler(_fh)
            self._debug_log_file_handler: logging.Handler | None = _fh
        except OSError:
            self._debug_log_path = None
            self._debug_log_dir = None
            self._debug_log_file_handler = None

        self._build_ui()
        self._init_compiler()

    # --- UI construction ---

    def _build_ui(self):
        # ── Stronger LabelFrame borders for better visibility ──
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure(
            "TLabelframe",
            borderwidth=2,
            relief="groove",
            bordercolor="#888888",
            lightcolor="#aaaaaa",
            darkcolor="#666666",
        )
        style.configure("TLabelframe.Label", foreground="#222222")
        # Make disabled Entry widgets visually obvious (darker background).
        style.map(
            "TEntry",
            fieldbackground=[("disabled", "#b8b8b8")],
            foreground=[("disabled", "#606060")],
        )

        # ── Menu bar ──
        self._build_menu_bar()

        # Status bar and progress bar are fixed at the bottom of self.root.
        # Pack status_bar first so it anchors the very bottom; the progress
        # bar is packed after and sits just above it.
        self.status_var = tk.StringVar(value="Initializing...")
        status_bar = ttk.Label(
            self.root, textvariable=self.status_var,
            relief="sunken", anchor="w", padding=(5, 2),
        )
        status_bar.pack(fill="x", side="bottom")

        self.progress_bar = ttk.Progressbar(
            self.root, mode="indeterminate", length=200
        )
        self.progress_bar.pack(fill="x", side="bottom", padx=0, pady=0)
        self.progress_bar.pack_forget()  # shown only while rendering

        # ── Notebook with all tabs — mode switching hides/shows tabs ──
        self._notebook = ttk.Notebook(self.root)
        self._notebook.pack(fill="both", expand=True, padx=10, pady=(10, 5))

        # Palmer-mode tabs
        self._palmer_tab = ttk.Frame(self._notebook)
        self._notebook.add(self._palmer_tab, text="Palmer")

        self._advanced_tab = ttk.Frame(self._notebook)
        self._notebook.add(self._advanced_tab, text="Advanced")

        # Converter-mode tab (hidden initially)
        self._converter_tab_frame = ttk.Frame(self._notebook)
        self._notebook.add(self._converter_tab_frame, text="Docx Converter")
        self._notebook.hide(self._converter_tab_frame)

        # Shared tab — visible in both modes
        self._about_tab = ttk.Frame(self._notebook)
        self._notebook.add(self._about_tab, text="About")

        self._build_palmer_tab(self._palmer_tab)
        self._build_advanced_tab(self._advanced_tab)
        self._converter_tab = ConverterTab(
            self.root, self.status_var, lambda: self.compiler,
            is_debug=lambda: self._debug_mode,
        )
        self._converter_tab.build(self._converter_tab_frame)
        self._build_about_tab(self._about_tab)

        self._current_mode = "palmer"
        self._update_mode_menu()

    def _build_menu_bar(self):
        """Create the application menu bar."""
        menubar = tk.Menu(self.root)
        self.root.configure(menu=menubar)

        file_menu = tk.Menu(menubar, tearoff=False)
        menubar.add_cascade(label="File", menu=file_menu)

        self._debug_menu_var = tk.BooleanVar(value=False)
        file_menu.add_checkbutton(
            label="Debug Mode",
            variable=self._debug_menu_var,
            command=self._on_toggle_debug_mode,
        )
        file_menu.add_separator()
        file_menu.add_command(
            label="Check for Updates\u2026",
            command=self._on_check_for_updates,
        )
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self._on_exit)

        self._file_menu = file_menu

        mode_menu = tk.Menu(menubar, tearoff=False)
        menubar.add_cascade(label="Mode", menu=mode_menu)

        self._input_mode_menu_var = tk.BooleanVar(value=True)
        self._converter_mode_menu_var = tk.BooleanVar(value=False)
        mode_menu.add_checkbutton(
            label="Input Mode",
            variable=self._input_mode_menu_var,
            command=self._switch_to_palmer,
        )
        mode_menu.add_checkbutton(
            label="Converter Mode",
            variable=self._converter_mode_menu_var,
            command=self._switch_to_converter,
        )

        self._mode_menu = mode_menu

    def _update_mode_menu(self):
        """Reflect the current mode in the Mode menu with a checkmark."""
        if self._current_mode == "palmer":
            self._input_mode_menu_var.set(True)
            self._converter_mode_menu_var.set(False)
        else:
            self._input_mode_menu_var.set(False)
            self._converter_mode_menu_var.set(True)

    def _set_window_icon(self) -> None:
        icon_candidates = [
            Path(getattr(sys, "_MEIPASS", "")) / "assets" / "palmer-type.ico",
            Path(__file__).parent / "assets" / "palmer-type.ico",
        ]
        icon_path: Path | None = None
        for candidate in icon_candidates:
            if candidate.exists():
                icon_path = candidate
                break
        if icon_path is None:
            return

        # Baseline: set via Tk (cross-platform fallback).
        try:
            self.root.iconbitmap(str(icon_path))
        except tk.TclError:
            pass

        # On Windows, override with native Win32 API so that the shell
        # picks the exact frame from the multi-size ICO instead of letting
        # Tk's internal handler scale a single frame (which causes blur).
        if sys.platform == "win32":
            try:
                import ctypes

                user32 = ctypes.windll.user32  # type: ignore[attr-defined]

                # Ensure 64-bit safe handle types.
                user32.LoadImageW.restype = ctypes.c_void_p
                user32.SendMessageW.argtypes = [
                    ctypes.c_void_p, ctypes.c_uint,
                    ctypes.c_void_p, ctypes.c_void_p,
                ]
                user32.SendMessageW.restype = ctypes.c_void_p

                SM_CXSMICON = 49
                SM_CXICON = 11
                IMAGE_ICON = 1
                LR_LOADFROMFILE = 0x0010
                WM_SETICON = 0x0080
                ICON_SMALL = 0
                ICON_BIG = 1

                small_cx = user32.GetSystemMetrics(SM_CXSMICON)
                big_cx = user32.GetSystemMetrics(SM_CXICON)

                hicon_small = user32.LoadImageW(
                    None, str(icon_path), IMAGE_ICON,
                    small_cx, small_cx, LR_LOADFROMFILE,
                )
                hicon_big = user32.LoadImageW(
                    None, str(icon_path), IMAGE_ICON,
                    big_cx, big_cx, LR_LOADFROMFILE,
                )
                if not hicon_small or not hicon_big:
                    return

                self.root.update_idletasks()
                frame_hwnd = self.root.winfo_id()
                if not frame_hwnd:
                    return

                hwnd = user32.GetParent(frame_hwnd)
                if not hwnd:
                    GA_ROOT = 2
                    hwnd = user32.GetAncestor(frame_hwnd, GA_ROOT)
                if not hwnd:
                    return

                user32.SendMessageW(hwnd, WM_SETICON, ICON_SMALL, hicon_small)
                user32.SendMessageW(hwnd, WM_SETICON, ICON_BIG, hicon_big)

                self._icon_handles = (hicon_small, hicon_big)
            except Exception:
                logger.debug("Win32 icon override failed", exc_info=True)

    def _update_title(self):
        """Refresh the window title to reflect mode, engine, and debug state."""
        engine = ""
        if self.compiler is not None:
            engine = f" [engine: {self.compiler.backend.name}]"
        debug = " [DEBUG]" if self._debug_mode else ""
        if self._current_mode == "converter":
            self.root.title(
                f"palmer-type  v{__version__}{engine} \u2014 Converter Mode{debug}")
        else:
            self.root.title(f"palmer-type  v{__version__}{engine}{debug}")

    def _switch_to_palmer(self):
        """Switch the notebook to Palmer Notation mode."""
        if self._current_mode == "palmer":
            return
        self._notebook.hide(self._converter_tab_frame)
        self._notebook.add(self._palmer_tab, text="Palmer")
        self._notebook.add(self._advanced_tab, text="Advanced")
        # Move About to the end so tab order is Palmer | Advanced | About.
        self._notebook.hide(self._about_tab)
        self._notebook.add(self._about_tab, text="About")
        self._notebook.select(self._palmer_tab)
        self._current_mode = "palmer"
        self._update_title()
        self._update_mode_menu()

    def _switch_to_converter(self):
        """Switch the notebook to Docx Converter mode."""
        if self._current_mode == "converter":
            return
        self._notebook.hide(self._palmer_tab)
        self._notebook.hide(self._advanced_tab)
        self._notebook.add(self._converter_tab_frame, text="Docx Converter")
        # Move About to the end so tab order is Docx Converter | About.
        self._notebook.hide(self._about_tab)
        self._notebook.add(self._about_tab, text="About")
        self._notebook.select(self._converter_tab_frame)
        self._current_mode = "converter"
        self._update_title()
        self._update_mode_menu()

    def _on_exit(self):
        """Exit the application; confirm only while a render is in progress."""
        if self._rendering:
            if not messagebox.askokcancel(
                "Exit", "A render is still in progress. Exit anyway?"
            ):
                return
        # Close the debug log file handler so the log file is flushed and
        # released.  The file itself is intentionally kept for post-mortem use.
        if self._debug_log_file_handler is not None:
            logging.getLogger().removeHandler(self._debug_log_file_handler)
            self._debug_log_file_handler.close()
        self.root.destroy()

    # ------------------------------------------------------------------
    # Check for Updates
    # ------------------------------------------------------------------
    def _on_check_for_updates(self):
        """Check GitHub for a newer release in a background thread."""
        threading.Thread(target=self._check_for_updates_worker, daemon=True).start()

    def _check_for_updates_worker(self):
        """Fetch the latest release tag from GitHub and compare versions."""
        try:
            req = Request(_RELEASES_API_URL, headers={"Accept": "application/vnd.github+json"})
            with urlopen(req, timeout=10) as resp:
                data = json.loads(resp.read().decode())
            tag = data.get("tag_name", "")
            latest = tag.lstrip("vV")
            current = __version__

            if self._is_newer_version(latest, current):
                self.root.after(0, self._show_update_available, latest)
            else:
                self.root.after(
                    0,
                    lambda: messagebox.showinfo(
                        "Check for Updates",
                        f"You are using the latest version (v{current}).",
                    ),
                )
        except Exception as exc:
            detail = str(exc)

            def _show_error(d: str = detail) -> None:
                messagebox.showwarning(
                    "Check for Updates",
                    f"Could not check for updates.\n\n{d}\n\n"
                    "Please verify your internet connection and try again.",
                )

            self.root.after(0, _show_error)

    @staticmethod
    def _is_newer_version(latest: str, current: str) -> bool:
        """Return True if *latest* is strictly newer than *current* (semver)."""
        def _parse(v: str) -> tuple[int, ...]:
            parts: list[int] = []
            for seg in v.split("."):
                m = re.match(r"(\d+)", seg)
                if m:
                    parts.append(int(m.group(1)))
            return tuple(parts)

        try:
            return _parse(latest) > _parse(current)
        except Exception:
            return False

    def _show_update_available(self, latest: str):
        """Show a dialog offering to open the releases page."""
        answer = messagebox.askyesno(
            "Update Available",
            f"A new version is available: v{latest}\n"
            f"(Current version: v{__version__})\n\n"
            "Would you like to open the download page?",
        )
        if answer:
            webbrowser.open_new_tab(_RELEASES_URL)

    def _on_toggle_debug_mode(self):
        """Toggle debug mode on/off."""
        self._debug_mode = self._debug_menu_var.get()
        self._update_title()
        if not self._debug_mode:
            self._hide_debug_log()
            self._about_tab_debug_frame.pack_forget()
        else:
            self._about_tab_debug_frame.pack(fill="x", in_=self._about_tab)
            self._debug_log_clear()
            self._show_debug_log()
            # ① Show startup log first (chronological order) by reading the
            #   persistent log file.  Fall back to the in-memory buffer when
            #   the file is unavailable.
            self._debug_log_append_raw("--- Startup log ---")
            if self._debug_log_path is not None and self._debug_log_path.exists():
                try:
                    for _line in self._debug_log_path.read_text(
                            encoding="utf-8").splitlines():
                        self._debug_log_append_raw(_line)
                except OSError:
                    for _line in self._startup_log:
                        self._debug_log_append_raw(_line)
            else:
                for _line in self._startup_log:
                    self._debug_log_append_raw(_line)
            # ② Then show current system info with fresh timestamps.
            self._debug_log_append_raw("")
            self._debug_log_append("--- Debug mode enabled ---")
            self._debug_log_append(f"Platform: {_get_platform_str()}")
            self._debug_log_append(f"Python: {sys.version}")
            self._debug_log_append(f"python-docx available: {_HAS_DOCX}")
            if _DOCX_IMPORT_ERROR:
                self._debug_log_append(f"  import error: {_DOCX_IMPORT_ERROR}")
            if self.compiler is not None:
                self._debug_log_append(
                    f"Engine: {self.compiler.backend.name} "
                    f"({self.compiler.backend.executable})")
            else:
                self._debug_log_append("Engine: (not yet initialized)")
            self._debug_log_append(
                f"CJK fallback font (log widgets): "
                f"{_detect_cjk_fallback_font()}")
            self._debug_log_append(_dump_font_fallback_detail())

    def _build_palmer_tab(self, parent: ttk.Frame):
        # ---- Input section ----
        input_frame = ttk.LabelFrame(parent, text="Dental Notation", padding=10)
        input_frame.pack(fill="x", padx=10, pady=(10, 5))

        # Quadrant + cross layout (5-col × 5-row grid).
        # Cols: 0=R-labels  1=R-entries  2=center(sep+mid)  3=L-labels  4=L-entries
        # Rows: 0=upper_mid  1=upper-entries  2=hsep  3=lower-entries  4=lower_mid
        # Patient's Right (R) is on the viewer's left; Patient's Left (L) on the right.
        quad_frame = ttk.Frame(input_frame)
        quad_frame.pack(fill="x")

        self.entries = {}
        self.novert_var = tk.BooleanVar(value=False)

        # Row 0 — Upper Mid: label to the left, entry centered in column 2.
        ttk.Label(quad_frame, text="Upper Mid:").grid(
            row=0, column=1, sticky="e", padx=(0, 4), pady=(0, 2))
        self.upper_mid_entry = ttk.Entry(quad_frame, width=5, font=("Consolas", 12))
        self.upper_mid_entry.grid(row=0, column=2, sticky="s", pady=(0, 2))
        ttk.Checkbutton(
            quad_frame, text="No L/R",
            variable=self.novert_var,
            command=self._on_novert_changed,
        ).grid(row=0, column=3, columnspan=2, sticky="w", padx=(5, 10), pady=(0, 2))

        # Row 1 — Upper quadrant entries.
        # Labels follow patient-view convention (R on left, L on right);
        # dict keys use the engine's UL/UR mapping (UL=patient's left, UR=patient's right).
        self.ur_label = ttk.Label(quad_frame, text="UR:")
        self.ur_label.grid(row=1, column=0, sticky="e", padx=(10, 2), pady=3)
        entry = ttk.Entry(quad_frame, width=18, font=("Consolas", 12))
        entry.grid(row=1, column=1, sticky="ew", padx=(2, 5), pady=3)
        self.entries["UL"] = entry

        ttk.Label(quad_frame, text="UL:").grid(row=1, column=3, sticky="e", padx=(5, 2), pady=3)
        entry = ttk.Entry(quad_frame, width=18, font=("Consolas", 12))
        entry.grid(row=1, column=4, sticky="ew", padx=(2, 10), pady=3)
        self.entries["UR"] = entry

        # Vertical separator — two halves flanking the hsep
        tk.Frame(quad_frame, bg="#999", width=2).grid(row=1, column=2, sticky="ns", padx=5)
        tk.Frame(quad_frame, bg="#999", width=2).grid(row=3, column=2, sticky="ns", padx=5)

        # Row 2 — Horizontal separator (full width)
        tk.Frame(quad_frame, bg="#999", height=2).grid(
            row=2, column=0, columnspan=5, sticky="ew", pady=2
        )

        # Row 3 — Lower quadrant entries (same label/key convention as upper row)
        self.lr_label = ttk.Label(quad_frame, text="LR:")
        self.lr_label.grid(row=3, column=0, sticky="e", padx=(10, 2), pady=3)
        entry = ttk.Entry(quad_frame, width=18, font=("Consolas", 12))
        entry.grid(row=3, column=1, sticky="ew", padx=(2, 5), pady=3)
        self.entries["LL"] = entry

        ttk.Label(quad_frame, text="LL:").grid(row=3, column=3, sticky="e", padx=(5, 2), pady=3)
        entry = ttk.Entry(quad_frame, width=18, font=("Consolas", 12))
        entry.grid(row=3, column=4, sticky="ew", padx=(2, 10), pady=3)
        self.entries["LR"] = entry

        # Row 4 — Lower Mid: label to the left, entry centered in column 2.
        ttk.Label(quad_frame, text="Lower Mid:").grid(
            row=4, column=1, sticky="e", padx=(0, 4), pady=(2, 0))
        self.lower_mid_entry = ttk.Entry(quad_frame, width=5, font=("Consolas", 12))
        self.lower_mid_entry.grid(row=4, column=2, sticky="n", pady=(2, 0))

        quad_frame.columnconfigure(1, weight=1)
        quad_frame.columnconfigure(4, weight=1)

        # Bind all quadrant/mid entries so Render button reflects pending changes.
        for e in self.entries.values():
            e.bind("<KeyRelease>", self._mark_dirty)
        self.upper_mid_entry.bind("<KeyRelease>", self._mark_dirty)
        self.lower_mid_entry.bind("<KeyRelease>", self._mark_dirty)

        # Input-order reminder for right-side quadrants.
        note_frame = ttk.Frame(input_frame)
        note_frame.pack(fill="x", padx=(10, 10), pady=(4, 0))
        ttk.Label(
            note_frame,
            text=(
                "Input order for UR and LR: Enter tooth numbers in natural order "
                "(1, 2, 3, \u2026 8).\nDisplay reversal is applied automatically."
            ),
            foreground="#555555",
            wraplength=480,
            justify="right",
        ).pack(anchor="e")

        # ---- Options section (font, font color, background) ----
        options_frame = ttk.LabelFrame(parent, text="Options", padding=10)
        options_frame.pack(fill="x", padx=10, pady=(5, 5))

        # ---- Font and size ----
        font_frame = ttk.Frame(options_frame)
        font_frame.pack(fill="x", pady=(0, 0))

        ttk.Label(font_frame, text="Font:").pack(side="left", padx=(10, 2))
        self.font_var = tk.StringVar(value="Times New Roman")
        self._system_fonts = sorted(
            f for f in set(tkfont.families()) if not f.startswith("@")
        )
        self._system_fonts_set = set(self._system_fonts)
        self._all_fonts = ["Times New Roman"] + [
            f for f in self._system_fonts if f != "Times New Roman"
        ]
        self._last_valid_font = "Times New Roman"
        self.font_combo = ttk.Combobox(
            font_frame, textvariable=self.font_var,
            values=self._build_font_values(), width=26,
        )
        self.font_combo.pack(side="left", padx=(2, 2))
        self.font_combo.bind("<<ComboboxSelected>>", self._on_font_selected)

        self._fav_var = tk.BooleanVar(
            value=self.config.is_favorite_font("Times New Roman"),
        )
        self._fav_check = ttk.Checkbutton(
            font_frame, text="Fav", variable=self._fav_var,
            command=self._toggle_favorite_font,
        )
        self._fav_check.pack(side="left", padx=(2, 15))

        ttk.Label(font_frame, text="Size:").pack(side="left", padx=(0, 2))
        self.size_spin = ttk.Spinbox(
            font_frame, from_=MIN_FONT_SIZE_PT, to=MAX_FONT_SIZE_PT, increment=0.5, width=6, format="%.1f",
            command=self._mark_dirty,
        )
        self.size_spin.set(str(DEFAULT_FONT_SIZE_PT))
        self.size_spin.pack(side="left", padx=(2, 2))
        self.size_spin.bind("<KeyRelease>", self._mark_dirty)
        ttk.Label(font_frame, text="pt").pack(side="left")

        # ---- Font color ----
        fc_frame = ttk.Frame(options_frame)
        fc_frame.pack(fill="x", pady=(5, 0))

        ttk.Label(fc_frame, text="Font Color:").pack(side="left", padx=(10, 6))

        self.font_color_mode_var = tk.StringVar(value="black")
        for val, label in [("black", "Black"), ("custom", "Custom")]:
            ttk.Radiobutton(
                fc_frame, text=label, variable=self.font_color_mode_var, value=val,
                command=self._on_font_color_mode_changed,
            ).pack(side="left", padx=3)

        # Color swatch button: shows current custom font color and opens picker on click.
        self.custom_font_color_btn = tk.Button(
            fc_frame, width=3, relief="solid", bd=1,
            command=self._on_pick_custom_font_color,
        )
        self._update_custom_font_color_btn()
        self.custom_font_color_btn.pack(side="left", padx=(1, 12))

        # ---- Background color ----
        bg_frame = ttk.Frame(options_frame)
        bg_frame.pack(fill="x", pady=(5, 0))

        ttk.Label(bg_frame, text="Background (live):").pack(side="left", padx=(10, 6))

        self.bg_mode_var = tk.StringVar(value="white")
        for val, label in [("white", "White"), ("custom", "Custom")]:
            ttk.Radiobutton(
                bg_frame, text=label, variable=self.bg_mode_var, value=val,
                command=self._on_bg_mode_changed,
            ).pack(side="left", padx=3)

        # Color swatch button: shows current custom color and opens picker on click.
        self.custom_color_btn = tk.Button(
            bg_frame, width=3, relief="solid", bd=1,
            command=self._on_pick_custom_color,
        )
        self._update_custom_btn_color()
        self.custom_color_btn.pack(side="left", padx=(1, 12))

        ttk.Radiobutton(
            bg_frame, text="Transparent (PNG / clipboard)", variable=self.bg_mode_var,
            value="transparent", command=self._on_bg_mode_changed,
        ).pack(side="left", padx=3)

        # ---- Action buttons ----
        action_frame = ttk.Frame(parent)
        action_frame.pack(fill="x", padx=10, pady=5)

        self.render_btn = ttk.Button(action_frame, text="▶ Render",
                                     command=self._on_render)
        self.render_btn.pack(side="left", padx=5)

        self.clip_btn = ttk.Button(action_frame, text="Copy to Clipboard",
                                   command=self._on_copy, state="disabled")
        self.clip_btn.pack(side="left", padx=5)

        self.save_btn = ttk.Button(action_frame, text="Save...",
                                   command=self._on_save, state="disabled")
        self.save_btn.pack(side="left", padx=5)

        # TeX engine indicator (variable kept for internal tracking; no visible label here)
        self.engine_label_var = tk.StringVar(value="Engine: detecting...")

        # Read-only TeX command display
        self.tex_var = tk.StringVar(value="")
        tex_label = ttk.Entry(action_frame, textvariable=self.tex_var,
                              state="readonly", font=("Consolas", _FONT_SIZE_DISPLAY))
        tex_label.pack(side="right", fill="x", expand=True, padx=5)

        # ---- Preview ----
        self.preview_frame = ttk.LabelFrame(parent, text="Preview", padding=10)
        self.preview_frame.pack(fill="both", expand=True, padx=10, pady=(5, 5))

        self.canvas = tk.Canvas(self.preview_frame, bg="white", relief="sunken", bd=1)
        self.canvas.pack(fill="both", expand=True)

        # ---- Error Log (hidden by default, shown on render failure) ----
        self.error_frame = ttk.LabelFrame(parent, text="Error Log", padding=5)

        _btn_row = ttk.Frame(self.error_frame)
        _btn_row.pack(fill="x", pady=(0, 3))
        ttk.Button(
            _btn_row, text="✕ Dismiss", command=self._hide_error_log,
        ).pack(side="right", padx=(3, 0))
        ttk.Button(
            _btn_row, text="Copy", command=self._copy_error_log,
        ).pack(side="right")

        self.error_log = _CjkScrolledText(
            self.error_frame, wrap="word",
            font=("Consolas", _FONT_SIZE_DISPLAY), height=5,
            bg="#fff3f3", fg="#cc0000",
            state="disabled",
        )
        self.error_log.pack(fill="both", expand=True)
        # Not packed initially — shown only when a render error occurs

        # ---- Debug Log (hidden by default, shown when Debug Mode is on) ----
        self.debug_frame = ttk.LabelFrame(parent, text="Debug Log", padding=5)
        self.debug_log = _CjkScrolledText(
            self.debug_frame, wrap="word",
            font=("Consolas", _FONT_SIZE_DISPLAY), height=5,
            bg="#f3f3ff", fg="#333366",
            state="disabled",
        )
        self.debug_log.pack(fill="both", expand=True)
        # Not packed initially — shown only when debug mode is enabled

    # --- Compiler initialization ---

    def _init_compiler(self):
        """Detect TeX backend in background thread."""
        def _detect():
            try:
                logger.debug("_init_compiler._detect: creating PalmerCompiler ...")
                compiler = PalmerCompiler()
                is_tectonic = compiler.backend.name == "tectonic"
                cache_exists = tectonic_cache_exists()
                first_launch = is_tectonic and not cache_exists
                logger.debug(
                    "_init_compiler._detect: backend=%s, is_tectonic=%s, "
                    "cache_exists=%s, first_launch=%s",
                    compiler.backend.name, is_tectonic, cache_exists, first_launch,
                )

                def _on_ready():
                    logger.debug("_init_compiler._on_ready: setting compiler on main thread")
                    self.compiler = compiler
                    self.engine_label_var.set(f"Engine: {compiler.backend.name}")
                    self._update_title()
                    if first_launch:
                        # Update status immediately so the user never sees
                        # "Initializing..." while we wait for the connectivity
                        # check (up to 5 s).  The download step will overwrite
                        # this message once it knows whether we are online.
                        self.status_var.set(
                            f"Ready — Engine: {compiler.backend.name}"
                            "  (checking connectivity...)")
                        logger.debug(
                            "_init_compiler._on_ready: first_launch=True, "
                            "spawning connectivity check thread")
                        def _check_and_pre():
                            logger.debug("_check_and_pre: calling _check_online() ...")
                            is_online = _check_online()
                            logger.debug("_check_and_pre: _check_online() returned %s", is_online)
                            logger.debug("_check_and_pre: scheduling _pre_download_tectonic on main thread")
                            self.root.after(
                                0,
                                lambda: self._pre_download_tectonic(
                                    compiler, is_online
                                ),
                            )

                        threading.Thread(
                            target=_check_and_pre, daemon=True
                        ).start()
                    else:
                        self.status_var.set(
                            f"Ready — Engine: {compiler.backend.name}")
                        logger.debug("_init_compiler._on_ready: ready (no first-launch)")

                self.root.after(0, _on_ready)
            except FileNotFoundError as e:
                logger.debug("_init_compiler._detect: FileNotFoundError: %s", e)
                def _on_not_found(err=e):
                    self.status_var.set(f"Engine not available: {err}")
                    self.engine_label_var.set("Engine: not available")
                    self.render_btn.configure(state="disabled")
                    messagebox.showerror(
                        "TeX Engine Not Found",
                        f"Could not initialize the TeX engine:\n\n{err}\n\n"
                        "Rendering is disabled.\n"
                        "Place tectonic in the bin/ folder, or install TeX Live / MiKTeX\n"
                        "so that xelatex is available on PATH."
                    )
                self.root.after(0, _on_not_found)
            except Exception as e:
                logger.debug("_init_compiler._detect: unexpected error: %s", e, exc_info=True)
                def _on_error(err=e):
                    self.status_var.set(f"Engine error: {err}")
                    self.engine_label_var.set("Engine: error")
                    self.render_btn.configure(state="disabled")
                    messagebox.showerror(
                        "Initialization Error",
                        f"Could not initialize the TeX engine:\n\n{err}\n\n"
                        "Rendering is disabled."
                    )
                self.root.after(0, _on_error)

        logger.debug("_init_compiler: spawning _detect thread")
        threading.Thread(target=_detect, daemon=True).start()

    def _pre_download_tectonic(self, compiler, is_online: bool) -> None:
        """Automatically download Tectonic support files in the background.

        Called on the main thread immediately after the compiler is detected and
        the cache is found to be empty.  If the machine appears to be offline the
        download is skipped and the buttons are left enabled so the user can still
        attempt a render manually later.

        ``is_online`` must be pre-computed on a background thread (see
        ``_init_compiler``) so this method never blocks the main thread.
        """
        logger.debug("_pre_download_tectonic: is_online=%s", is_online)
        if not is_online:
            logger.debug("_pre_download_tectonic: offline — skipping download")
            self.status_var.set(
                f"Ready — Engine: {compiler.backend.name}"
                "  (offline — TeX support files will be downloaded on first render)")
            return

        try:
            # Disable Render and Convert buttons while downloading.
            logger.debug("_pre_download_tectonic: setting _tectonic_downloading flag")
            self._tectonic_downloading = True
            logger.debug("_pre_download_tectonic: disabling render button")
            self.render_btn.configure(state="disabled")
            logger.debug("_pre_download_tectonic: disabling converter tab")
            self._converter_tab.set_convert_enabled(False)
            logger.debug("_pre_download_tectonic: updating status text")
            self.status_var.set("Downloading TeX support files (~100 MB) — please wait...")
            logger.debug("_pre_download_tectonic: packing progress bar")
            self.progress_bar.pack(fill="x", side="bottom", padx=0, pady=0)
            logger.debug("_pre_download_tectonic: starting progress bar animation")
            self.progress_bar.start(_PROGRESS_BAR_INTERVAL_MS)
            logger.debug("_pre_download_tectonic: spawning _do_download thread")
        except Exception:
            logger.debug(
                "_pre_download_tectonic: EXCEPTION before thread start",
                exc_info=True,
            )
            # Re-enable so the user can at least attempt a manual render.
            self._tectonic_downloading = False
            self.render_btn.configure(state="normal")
            self._converter_tab.set_convert_enabled(True)
            self.status_var.set(
                f"Ready — Engine: {compiler.backend.name}"
                "  (TeX file download skipped — will download on first render)")
            return

        def _do_download():
            logger.debug("_do_download: thread started")
            success = False
            try:
                logger.debug("_do_download: compiling dummy document to populate cache ...")
                # Compile a minimal dummy document to populate the Tectonic cache.
                # Allow up to 10 minutes for the initial ~100 MB bundle download.
                compiler.render_raw("~", validate=False, compile_timeout=600)
                success = True
                logger.debug("_do_download: cache download succeeded")
            except Exception as exc:
                logger.debug("_do_download: cache download failed: %s", exc, exc_info=True)
                pass  # download error is non-fatal; user can retry by clicking Render

            def _on_done():
                logger.debug("_do_download._on_done: re-enabling render button (success=%s)", success)
                self._tectonic_downloading = False
                self.progress_bar.stop()
                self.progress_bar.pack_forget()
                self.render_btn.configure(state="normal")
                self._converter_tab.set_convert_enabled(True)
                if success:
                    self.status_var.set(f"Ready — Engine: {compiler.backend.name}")
                else:
                    self.status_var.set(
                        f"Ready — Engine: {compiler.backend.name}"
                        "  (TeX file download failed — will retry on first render)")

            self.root.after(0, _on_done)

        threading.Thread(target=_do_download, daemon=True).start()

    # --- Advanced tab ---

    def _build_advanced_tab(self, parent: ttk.Frame):
        """Build the Advanced tab (DPI + per-side image margins)."""
        # ---- Output Resolution ----
        dpi_frame = ttk.LabelFrame(parent, text="Output Resolution", padding=10)
        dpi_frame.pack(fill="x", padx=10, pady=(10, 5))

        ttk.Label(dpi_frame, text="DPI:").grid(row=0, column=0, sticky="e", padx=(0, 5), pady=4)
        self.dpi_spin = ttk.Spinbox(
            dpi_frame, from_=MIN_DPI, to=MAX_DPI, increment=1, width=6,
            command=self._on_advanced_change,
        )
        self.dpi_spin.set("600")
        self.dpi_spin.grid(row=0, column=1, sticky="w", pady=4)
        self.dpi_spin.bind("<KeyRelease>", self._on_advanced_change)
        ttk.Label(dpi_frame, text="dpi").grid(row=0, column=2, sticky="w", padx=(3, 0))

        # ---- Image Margins ----
        margin_frame = ttk.LabelFrame(parent, text="Image Margins", padding=10)
        margin_frame.pack(fill="x", padx=10, pady=(5, 5))

        self.margin_spins: dict[str, ttk.Spinbox] = {}
        self.margin_mm_vars: dict[str, tk.StringVar] = {}
        for row, (key, label) in enumerate([
            ("top",    "Top:"),
            ("bottom", "Bottom:"),
            ("left",   "Left:"),
            ("right",  "Right:"),
        ]):
            ttk.Label(margin_frame, text=label).grid(
                row=row, column=0, sticky="e", padx=(0, 5), pady=4
            )
            # Sub-frame keeps spinbox + units tightly packed
            val_frame = ttk.Frame(margin_frame)
            val_frame.grid(row=row, column=1, sticky="w", pady=4)

            spin = ttk.Spinbox(
                val_frame, from_=0, to=200, increment=1, width=6,
                command=self._on_advanced_change,
            )
            spin.set(str(DEFAULT_MARGIN_PX))
            spin.pack(side="left")
            spin.bind("<KeyRelease>", self._on_advanced_change)
            ttk.Label(val_frame, text="px").pack(side="left", padx=(2, 0))
            mm_var = tk.StringVar(value="\u2248 0.34 mm")
            ttk.Label(val_frame, textvariable=mm_var, foreground="#555555").pack(
                side="left", padx=(6, 0)
            )
            self.margin_spins[key] = spin
            self.margin_mm_vars[key] = mm_var

        ttk.Label(
            margin_frame,
            text=(
                "Margins are applied after auto-cropping the rendered image. "
                "Changing any value marks the current preview as stale \u2014 "
                "click Render to update."
            ),
            foreground="#555555",
            wraplength=420,
            justify="left",
        ).grid(row=4, column=0, columnspan=2, sticky="w", pady=(6, 0))

        # Initialise mm labels with the default values.
        self._update_mm_labels()

    def _on_advanced_change(self, *_):
        """Called when any Advanced-tab value changes; refreshes mm labels and dirty flag."""
        self._mark_dirty()
        self._update_mm_labels()

    def _on_delete_font_favorites(self) -> None:
        """Clear all font favorites after user confirmation."""
        if not self.config.get_favorite_fonts():
            messagebox.showinfo("Delete Font Favorites", "No font favorites are saved.")
            return
        if not messagebox.askyesno(
            "Delete Font Favorites",
            "Delete all font favorites?\n\nThis cannot be undone.",
        ):
            return
        self.config.clear_favorite_fonts()
        self._fav_var.set(False)
        self._refresh_font_combo()
        messagebox.showinfo("Delete Font Favorites", "Font favorites have been deleted.")

    def _on_open_log_folder(self) -> None:
        """Open the debug log folder in the system file manager."""
        if self._debug_log_dir is None or not self._debug_log_dir.exists():
            messagebox.showinfo("Log Folder", "Log folder not found.")
            return
        try:
            if sys.platform == "win32":
                os.startfile(str(self._debug_log_dir))
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(self._debug_log_dir)])
            else:
                subprocess.Popen(["xdg-open", str(self._debug_log_dir)])
        except OSError as exc:
            messagebox.showerror("Open Log Folder", f"Failed to open folder:\n{exc}")

    def _on_delete_cache(self):
        """Delete the Tectonic cache after user confirmation."""
        if not tectonic_cache_exists():
            messagebox.showinfo("Delete Cache", "Tectonic cache is already empty.")
            return
        confirm = messagebox.askyesno(
            "Delete Cache",
            "Delete the Tectonic cache?\n\n"
            "This will remove ~100 MB of downloaded TeX support files.\n"
            "They will be re-downloaded on the next render.",
        )
        if not confirm:
            return
        try:
            delete_tectonic_cache()
            messagebox.showinfo("Delete Cache", "Tectonic cache has been deleted.")
        except OSError as e:
            messagebox.showerror("Delete Cache", f"Failed to delete cache:\n{e}")

    def _get_dpi(self) -> int:
        """Read the DPI spinbox value, falling back to 600 on invalid input."""
        return clamp_dpi(self.dpi_spin.get())

    def _update_mm_labels(self):
        """Recalculate and refresh the mm equivalent labels for all margin spinboxes."""
        dpi = self._get_dpi()
        for key, spin in self.margin_spins.items():
            try:
                px = max(0, int(float(spin.get())))
            except ValueError:
                px = 8
            mm = px * MM_PER_INCH / dpi
            self.margin_mm_vars[key].set(f"\u2248 {mm:.2f} mm")

    def _get_margins(self) -> dict[str, int]:
        """Read margin spinbox values, falling back to 8 on invalid input."""
        result = {}
        for key, spin in self.margin_spins.items():
            try:
                result[key] = max(0, int(float(spin.get())))
            except ValueError:
                result[key] = 8
        return result

    def _build_about_tab(self, parent: ttk.Frame):
        """Build the About tab content."""
        # --- Header ---
        header_frame = ttk.Frame(parent, padding=(20, 15, 20, 5))
        header_frame.pack(fill="x")

        ttk.Label(
            header_frame,
            text="palmer-type.exe",
            font=("", _UI_HEADER_FONT_SIZE, "bold"),
        ).pack(anchor="w")
        ttk.Label(
            header_frame,
            text="Render Zsigmondy-Palmer Dental Notation as Images",
            foreground="#555555",
        ).pack(anchor="w")
        ttk.Label(header_frame, text=f"Version: {__version__}").pack(anchor="w")
        ttk.Label(header_frame, text="\u00a9 2026 Yosuke Yamazaki, Dept. Anatomy, Nihon University School of Dentistry").pack(anchor="w")

        repo_link = ttk.Label(
            header_frame,
            text=_REPO_URL,
            foreground="#0066cc",
            cursor="hand2",
            font=("", _FONT_SIZE_DISPLAY, "underline"),
        )
        repo_link.pack(anchor="w")
        repo_link.bind("<Button-1>", lambda e: webbrowser.open_new_tab(_REPO_URL))

        ttk.Separator(parent, orient="horizontal").pack(fill="x", padx=10)

        # --- Scrollable license area ---
        text_frame = ttk.Frame(parent, padding=(10, 5))
        text_frame.pack(fill="both", expand=True)

        st = ScrolledText(
            text_frame,
            wrap="word",
            state="disabled",
            font=("Consolas", _FONT_SIZE_DISPLAY),
            relief="flat",
            padx=10,
            pady=5,
        )
        st.pack(fill="both", expand=True)

        st.tag_configure("h2",   font=("", 10, "bold"), spacing1=10, spacing3=4)
        st.tag_configure("body", font=("Consolas", _FONT_SIZE_DISPLAY))
        st.tag_configure("url",  foreground="#0066cc", underline=True)

        def _open_url(url: str) -> None:
            webbrowser.open_new_tab(url)

        st.tag_bind("url", "<Button-1>", lambda e: _open_url("https://tectonic-typesetting.github.io/"))
        st.tag_bind("url", "<Enter>",    lambda e: st.configure(cursor="hand2"))
        st.tag_bind("url", "<Leave>",    lambda e: st.configure(cursor=""))

        _MIT_BODY = (
            "Permission is hereby granted, free of charge, to any person obtaining a copy\n"
            "of this software and associated documentation files (the \"Software\"), to deal\n"
            "in the Software without restriction, including without limitation the rights\n"
            "to use, copy, modify, merge, publish, distribute, sublicense, and/or sell\n"
            "copies of the Software, and to permit persons to whom the Software is\n"
            "furnished to do so, subject to the following conditions:\n\n"
            "The above copyright notice and this permission notice shall be included in all\n"
            "copies or substantial portions of the Software.\n\n"
            "THE SOFTWARE IS PROVIDED \"AS IS\", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR\n"
            "IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,\n"
            "FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE\n"
            "AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER\n"
            "LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,\n"
            "OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE\n"
            "SOFTWARE."
        )

        _APP_MIT = (
            "MIT License\n\n"
            "Copyright (c) 2026 Yosuke Yamazaki\n\n"
            + _MIT_BODY
        )

        _TECTONIC_MIT = (
            "MIT License\n\n"
            "Copyright 2016-2023 the Tectonic Project\n\n"
            + _MIT_BODY
        )

        def _ins(tag: str, text: str) -> None:
            st.configure(state="normal")
            st.insert("end", text, tag)
            st.configure(state="disabled")

        _ins("h2",   "This Software License\n")
        _ins("body", _APP_MIT + "\n")
        _ins("h2",   "\nAcknowledgments \u2014 Tectonic\n")
        _ins("body", "This application bundles the Tectonic TeX engine.\n")
        _ins("url",  "https://tectonic-typesetting.github.io/\n\n")
        _ins("body", _TECTONIC_MIT + "\n")

        # --- Debug-only section (hidden until debug mode is enabled) ---
        self._about_tab_debug_frame = ttk.Frame(parent, padding=(10, 5, 10, 10))
        ttk.Separator(self._about_tab_debug_frame, orient="horizontal").pack(fill="x", pady=(0, 8))
        btn_frame = ttk.Frame(self._about_tab_debug_frame)
        btn_frame.pack(anchor="w")
        ttk.Button(
            btn_frame,
            text="Delete Tectonic Cache\u2026",
            command=self._on_delete_cache,
        ).grid(row=0, column=0, padx=(0, 6), sticky="w")
        ttk.Button(
            btn_frame,
            text="Delete Font Favorites\u2026",
            command=self._on_delete_font_favorites,
        ).grid(row=0, column=1, padx=(0, 6), sticky="w")
        ttk.Button(
            btn_frame,
            text="Open Log Folder",
            command=self._on_open_log_folder,
        ).grid(row=0, column=2, padx=(0, 6), sticky="w")
        # Initially hidden; shown when debug mode is enabled.

    # --- Dirty state helpers ---

    # --- Font favorites ---

    _FONT_SEPARATOR = "\u2500" * 20

    def _build_font_values(self) -> list[str]:
        """Build font combobox values with favorites at the top."""
        favorites = [
            f for f in self.config.get_favorite_fonts()
            if f in self._system_fonts_set
        ]
        if favorites:
            return favorites + [self._FONT_SEPARATOR] + self._all_fonts
        return list(self._all_fonts)

    def _on_font_selected(self, _event: object = None) -> None:
        """Handle font combobox selection, rejecting the separator."""
        selected = self.font_var.get()
        if selected == self._FONT_SEPARATOR:
            self.font_var.set(self._last_valid_font)
            return
        self._last_valid_font = selected
        self._fav_var.set(self.config.is_favorite_font(selected))
        self._mark_dirty()

    def _toggle_favorite_font(self) -> None:
        """Add or remove the current font from favorites."""
        font = self.font_var.get()
        if not font or font == self._FONT_SEPARATOR:
            return
        if self._fav_var.get():
            self.config.add_favorite_font(font)
        else:
            self.config.remove_favorite_font(font)
        self._refresh_font_combo()

    def _refresh_font_combo(self) -> None:
        """Rebuild font combobox values, preserving current selection."""
        current = self.font_var.get()
        self.font_combo["values"] = self._build_font_values()
        self.font_var.set(current)

    def _mark_dirty(self, *_):
        """Mark inputs as changed; clear the preview until re-rendered."""
        with self._gen_lock:
            self._render_generation += 1
        self.current_image = None
        self.canvas.delete("all")
        self.clip_btn.configure(state="disabled")
        self.save_btn.configure(state="disabled")
        self._update_tex_display(self._get_params())

    # --- Background helpers ---

    @staticmethod
    def _repaint_color_swatch(btn, color: tuple[int, int, int]) -> None:
        """Repaint a color swatch button to match *color*."""
        h = f"#{color[0]:02x}{color[1]:02x}{color[2]:02x}"
        btn.configure(bg=h, activebackground=h)

    @staticmethod
    def _pick_color(
        current: tuple[int, int, int], title: str,
    ) -> tuple[int, int, int] | None:
        """Open a color picker seeded with *current*; return chosen RGB or None."""
        initial = f"#{current[0]:02x}{current[1]:02x}{current[2]:02x}"
        result = colorchooser.askcolor(color=initial, title=title)
        if result[1] is None:
            return None
        h = result[1]
        return (int(h[1:3], 16), int(h[3:5], 16), int(h[5:7], 16))

    def _update_custom_btn_color(self):
        """Repaint the custom color swatch to match self.custom_bg_color."""
        self._repaint_color_swatch(self.custom_color_btn, self.custom_bg_color)

    def _on_pick_custom_color(self):
        """Open a color picker dialog; select 'custom' mode on confirmation."""
        chosen = self._pick_color(self.custom_bg_color, "Choose Background Color")
        if chosen is not None:
            self.custom_bg_color = chosen
            self._update_custom_btn_color()
            self.bg_mode_var.set("custom")
            self._on_bg_mode_changed()

    def _on_bg_mode_changed(self):
        """Update canvas color and refresh preview when background mode changes."""
        self.canvas.configure(bg=self._canvas_bg_hex())
        if self.current_image is not None:
            self._show_preview(self.current_image)

    def _canvas_bg_hex(self) -> str:
        """Return a hex color string suitable for the canvas background."""
        mode = self.bg_mode_var.get()
        if mode == "white":
            return "white"
        if mode == "custom":
            r, g, b = self.custom_bg_color
            return f"#{r:02x}{g:02x}{b:02x}"
        # transparent — represent with a neutral gray checkerboard-like color.
        return "#cccccc"

    # --- Font color helpers ---

    def _update_custom_font_color_btn(self):
        """Repaint the custom font color swatch to match self.custom_font_color."""
        self._repaint_color_swatch(self.custom_font_color_btn, self.custom_font_color)

    def _on_pick_custom_font_color(self):
        """Open a color picker dialog; select 'custom' mode on confirmation."""
        chosen = self._pick_color(self.custom_font_color, "Choose Font Color")
        if chosen is not None:
            self.custom_font_color = chosen
            self._update_custom_font_color_btn()
            self.font_color_mode_var.set("custom")
            self._mark_dirty()

    def _on_font_color_mode_changed(self):
        """Mark dirty when font color mode changes."""
        self._mark_dirty()

    def _get_text_color(self) -> str:
        """Return text_color string for the engine based on current font color mode."""
        if self.font_color_mode_var.get() == "black":
            return ""  # empty string = default black
        r, g, b = self.custom_font_color
        return f"#{r:02x}{g:02x}{b:02x}"

    def _bg_rgb(self) -> tuple[int, int, int] | None:
        """Return the current background as an RGB tuple, or None for transparent."""
        mode = self.bg_mode_var.get()
        if mode == "white":
            return _WHITE
        if mode == "custom":
            return self.custom_bg_color
        return None  # transparent

    def _apply_background(
        self, img: Image.Image, force_opaque: bool = False
    ) -> Image.Image:
        """Composite a white-background image onto the selected background.

        Args:
            img: Source image with a white background (RGB mode from the engine).
            force_opaque: When True, transparent mode falls back to white (for
                          formats that do not support an alpha channel, e.g. JPEG/PDF).

        Returns:
            RGB image for opaque outputs, or RGBA image when transparency is requested.
        """
        bg = self._bg_rgb()

        if bg == _WHITE:
            # Already white — no compositing needed.
            return img.convert("RGB")

        # Build an alpha mask: white → transparent, dark → opaque.
        # ImageOps.invert on a grayscale image maps 255→0 and 0→255.
        mask = ImageOps.invert(img.convert("L"))
        rgba = img.convert("RGBA")
        rgba.putalpha(mask)

        if bg is None:
            # Transparent PNG requested.
            if force_opaque:
                # Fall back to white for opaque formats.
                bg = _WHITE
            else:
                return rgba

        result = Image.new("RGBA", img.size, bg + (255,))
        result.paste(rgba, mask=rgba.split()[3])
        return result.convert("RGB")

    # --- Event handlers ---

    def _on_novert_changed(self):
        """Clear and lock/unlock entries affected by the novert checkbox.

        When novert is enabled the right column (GUI labels UL/LL, engine keys
        UR/LR) and both mid entries are cleared and locked; the left column
        (GUI labels UR/LR, engine keys UL/LL) remains editable.
        'novert' is then passed as both upper_mid and lower_mid to the engine.
        """
        _novert_entries = [
            self.entries["UR"],
            self.entries["LR"],
            self.upper_mid_entry,
            self.lower_mid_entry,
        ]
        if self.novert_var.get():
            for e in _novert_entries:
                e.configure(state="normal")
                e.delete(0, "end")
                e.configure(state="disabled")
            self.ur_label["text"] = "U:"
            self.lr_label["text"] = "L:"
        else:
            for e in _novert_entries:
                e.configure(state="normal")
            self.ur_label["text"] = "UR:"
            self.lr_label["text"] = "LR:"
        self._mark_dirty()

    def _get_params(self) -> dict:
        """Collect parameters from input fields."""
        try:
            font_size_pt = float(self.size_spin.get())
        except ValueError:
            font_size_pt = DEFAULT_FONT_SIZE_PT
        upper_mid = "novert" if self.novert_var.get() else self.upper_mid_entry.get().strip()
        lower_mid = "novert" if self.novert_var.get() else self.lower_mid_entry.get().strip()
        return {
            "UL": self.entries["UL"].get().strip(),
            "UR": self.entries["UR"].get().strip(),
            "LL": self.entries["LL"].get().strip(),
            "LR": self.entries["LR"].get().strip(),
            "upper_mid": upper_mid,
            "lower_mid": lower_mid,
            "font_family": self.font_var.get(),
            "font_size_pt": font_size_pt,
            "text_color": self._get_text_color(),
        }

    def _update_tex_display(self, params: dict):
        """Show the generated TeX command in the read-only entry."""
        tex = (
            f"\\Palmer"
            f"{{{params['UL']}}}{{{params['UR']}}}"
            f"{{{params['LR']}}}{{{params['LL']}}}"
            f"{{{params['upper_mid']}}}{{{params['lower_mid']}}}"
        )
        self.tex_var.set(tex)

    def _set_ui_busy(self, busy: bool):
        """Lock/unlock the UI and show/hide the progress bar during rendering."""
        self._rendering = busy
        self.render_btn.configure(state="disabled" if busy else "normal")
        # Enable Copy/Save only when an image is available.
        can_use = not busy and self.current_image is not None
        self.clip_btn.configure(state="normal" if can_use else "disabled")
        self.save_btn.configure(state="normal" if can_use else "disabled")

        if busy:
            self.progress_bar.pack(fill="x", side="bottom", padx=0, pady=0)
            self.progress_bar.start(_PROGRESS_BAR_INTERVAL_MS)
            self.root.configure(cursor="watch")
        else:
            self.progress_bar.stop()
            self.progress_bar.pack_forget()
            self.root.configure(cursor="")

    def _on_render(self):
        """Run the TeX compilation."""
        if self._rendering:
            return  # prevent concurrent render threads
        if self.compiler is None:
            messagebox.showerror("Error", "TeX engine not found.")
            return

        compiler = self.compiler
        params = self._get_params()
        margins = self._get_margins()
        dpi = self._get_dpi()

        # Validate font size range before starting the background render.
        if not (MIN_FONT_SIZE_PT <= params["font_size_pt"] <= MAX_FONT_SIZE_PT):
            messagebox.showerror(
                "Invalid Font Size",
                f"Font size must be between {MIN_FONT_SIZE_PT} and "
                f"{MAX_FONT_SIZE_PT} pt (got {params['font_size_pt']}).",
            )
            return

        # Warn if all quadrants are empty.
        if all(not params[k] for k in ("UL", "UR", "LR", "LL")):
            proceed = messagebox.askokcancel(
                "Input is empty",
                "All four quadrants (UL, UR, LR, LL) are empty.\n"
                "Rendering now may produce a blank image.\n\n"
                "Do you want to continue?",
            )
            if not proceed:
                return

        self._update_tex_display(params)

        self._set_ui_busy(True)
        self.status_var.set("Compiling...")

        debug = self._debug_mode
        if debug:
            self._debug_log_clear()
            self._show_debug_log()
            self._debug_log_append("Render started")
            self._debug_log_append(
                f"Engine: {compiler.backend.name} ({compiler.backend.executable})")
            self._debug_log_append(
                f"Params: UL={params['UL']!r}  UR={params['UR']!r}  "
                f"LL={params['LL']!r}  LR={params['LR']!r}")
            self._debug_log_append(
                f"Mid: upper={params['upper_mid']!r}  lower={params['lower_mid']!r}")
            self._debug_log_append(
                f"Font: {params['font_family']} {params['font_size_pt']}pt  "
                f"color={params['text_color']}")
            self._debug_log_append(
                f"DPI: {dpi}  Margins: T={margins['top']} B={margins['bottom']} "
                f"L={margins['left']} R={margins['right']}")
            self._debug_log_append(
                f"CJK fallback font (log widgets): {_detect_cjk_fallback_font()}")
        else:
            self._hide_debug_log()

        # Capture the generation counter before launching the background thread.
        # If the user changes inputs while rendering, _mark_dirty() increments
        # the counter and the stale result will be silently discarded.
        with self._gen_lock:
            gen = self._render_generation

        def _do_render():
            try:
                if debug:
                    self._debug_log_append("Compiling TeX ...")
                img = compiler.render(
                    **params,
                    dpi=dpi,
                    margin_top=margins["top"],
                    margin_bottom=margins["bottom"],
                    margin_left=margins["left"],
                    margin_right=margins["right"],
                )
                if debug:
                    self._debug_log_append(
                        f"Render complete: {img.width}×{img.height}px, "
                        f"mode={img.mode}")
                def _on_done(img=img):
                    with self._gen_lock:
                        stale = gen != self._render_generation
                    if stale:
                        self.status_var.set("Inputs changed during render \u2014 click Render to update.")
                        if debug:
                            self._debug_log_append("Result discarded (inputs changed during render)")
                        return
                    self._hide_error_log()
                    self.current_image = img
                    self._show_preview(img)
                    self.status_var.set(f"Done — {img.width}×{img.height}px")
                    if debug:
                        self._debug_log_append("Preview updated")
                self.root.after(0, _on_done)
            except Exception as e:
                err_text = str(e)
                if debug:
                    self._debug_log_append(f"ERROR: {err_text}")
                self.root.after(0, functools.partial(
                    self.status_var.set, "\u26a0 Compile Error \u2014 see Error Log below"))
                self.root.after(0, functools.partial(self._show_error_log, err_text))
            finally:
                self.root.after(0, lambda: self._set_ui_busy(False))

        threading.Thread(target=_do_render, daemon=True).start()

    # --- Error log helpers ---

    def _repack_log_panels(self) -> None:
        """Re-pack preview, error log, and debug log panels in correct order."""
        self.preview_frame.pack_forget()
        self.error_frame.pack_forget()
        self.debug_frame.pack_forget()
        # Pack bottom-up: debug (bottommost), error, then preview (topmost).
        if self._debug_mode:
            self.debug_frame.pack(fill="x", side="bottom", padx=10, pady=(0, 5))
        if self._show_error:
            self.error_frame.pack(fill="x", side="bottom", padx=10, pady=(0, 5))
        self.preview_frame.pack(fill="both", expand=True, padx=10, pady=(5, 5))

    def _show_error_log(self, text: str) -> None:
        """Show the error log panel below the preview."""
        self.error_log.configure(state="normal")
        self.error_log.delete("1.0", "end")
        self.error_log.insert("end", text)
        self.error_log.see("1.0")
        self.error_log.configure(state="disabled")
        self._show_error = True
        self._repack_log_panels()

    def _hide_error_log(self) -> None:
        """Hide the error log panel."""
        self._show_error = False
        self.error_frame.pack_forget()

    def _copy_error_log(self) -> None:
        """Copy error log contents to clipboard."""
        text = self.error_log.get("1.0", "end").rstrip()
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.status_var.set("Error log copied to clipboard.")

    # --- Debug log helpers ---

    def _debug_log_append(self, text: str) -> None:
        """Append a timestamped line to the debug log (thread-safe)."""
        if text.strip():
            ts = datetime.datetime.now().strftime("%H:%M:%S.%f")[:-3]
            line = f"[{ts}] {text}\n"
        else:
            line = "\n"
        def _append():
            self.debug_log.configure(state="normal")
            self.debug_log.insert("end", line)
            self.debug_log.see("end")
            self.debug_log.configure(state="disabled")
        self.root.after(0, _append)

    def _debug_log_append_raw(self, text: str) -> None:
        """Append a line to the debug log without adding a timestamp (thread-safe).

        Used when replaying log lines that already carry their own timestamp
        (e.g. lines read back from the persistent log file).
        """
        line = text + "\n"
        def _append():
            self.debug_log.configure(state="normal")
            self.debug_log.insert("end", line)
            self.debug_log.see("end")
            self.debug_log.configure(state="disabled")
        self.root.after(0, _append)

    def _debug_log_clear(self) -> None:
        """Clear the debug log."""
        self.debug_log.configure(state="normal")
        self.debug_log.delete("1.0", "end")
        self.debug_log.configure(state="disabled")

    def _show_debug_log(self) -> None:
        """Show the debug log panel below the preview."""
        self._repack_log_panels()

    def _hide_debug_log(self) -> None:
        """Hide the debug log panel."""
        self.debug_frame.pack_forget()

    # --- Preview rendering ---

    def _show_preview(self, img: Image.Image):
        """Display image on canvas with the current background applied."""
        # Apply background for display (always opaque so the canvas shows correctly).
        display_src = self._apply_background(img, force_opaque=True)

        # winfo_width/height return 1 before the window is first mapped; use a
        # safe minimum so the preview is not scaled to a near-invisible size.
        cw = max(self.canvas.winfo_width(), _MIN_CANVAS_PX)
        ch = max(self.canvas.winfo_height(), _MIN_CANVAS_PX)

        # Cap at 1.0 — never upscale; downscale only when image exceeds canvas.
        scale = min(cw / max(display_src.width, 1), ch / max(display_src.height, 1), 1.0)
        new_w = max(1, int(display_src.width * scale))
        new_h = max(1, int(display_src.height * scale))
        display_img = display_src.resize((new_w, new_h), Image.Resampling.LANCZOS)

        self.photo_image = ImageTk.PhotoImage(display_img)
        self.canvas.delete("all")
        self.canvas.create_image(cw // 2, ch // 2, anchor="center",
                                 image=self.photo_image)


    def _on_copy(self):
        """Copy current image to clipboard."""
        if self.current_image is None:
            return
        if sys.platform != "win32":
            self.status_var.set("Clipboard copy is only supported on Windows.")
            return
        try:
            from palmer_engine import copy_image_to_clipboard_win32
            out_img = self._apply_background(self.current_image, force_opaque=False)
            copy_image_to_clipboard_win32(out_img, dpi=self._get_dpi())
            self.status_var.set("✓ Copied to clipboard")
        except (OSError, RuntimeError) as e:
            self.status_var.set(f"Copy failed: {e}")
            messagebox.showerror("Copy Error", str(e))

    def _on_save(self):
        """Save current image as PNG, JPEG, or PDF."""
        if self.current_image is None:
            return

        desktop = Path.home() / "Desktop"
        path = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[
                ("PNG Image", "*.png"),
                ("JPEG Image", "*.jpg *.jpeg"),
                ("PDF Document", "*.pdf"),
                ("All files", "*.*"),
            ],
            title="Save Palmer Dental Notation",
            initialdir=str(desktop) if desktop.exists() else str(Path.home()),
        )
        if not path:
            return

        save_path = Path(path)
        suffix = save_path.suffix.lower()
        dpi = self._get_dpi()

        try:
            if suffix == ".pdf":
                # PDF does not support alpha — transparent mode falls back to white.
                out_img = self._apply_background(self.current_image, force_opaque=True)
                out_img.save(str(save_path), "PDF", resolution=dpi)

            elif suffix in (".jpg", ".jpeg"):
                # JPEG does not support alpha — transparent mode falls back to white.
                out_img = self._apply_background(self.current_image, force_opaque=True)
                out_img.save(str(save_path), "JPEG", dpi=(dpi, dpi), quality=95)

            else:
                # PNG — transparent background is fully supported.
                out_img = self._apply_background(self.current_image, force_opaque=False)
                out_img.save(str(save_path), "PNG", dpi=(dpi, dpi))

            self.status_var.set(f"✓ Saved: {save_path}")

        except (OSError, RuntimeError) as e:
            self.status_var.set(f"Save failed: {e}")
            messagebox.showerror("Save Error", str(e))

    def run(self):
        self.root.mainloop()


def main():
    app = PalmerTypeApp()
    app.run()


if __name__ == "__main__":
    main()
