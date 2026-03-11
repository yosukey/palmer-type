"""
Palmer dental notation compiler.

Auto-detects the best available TeX engine in the following order:

  1. Bundled / adjacent Tectonic  (PyInstaller bundle or ``bin/`` folder)
  2. System ``tectonic`` on PATH
  3. System ``xelatex``  on PATH  (TeX Live / MiKTeX)

The bundled variant ships with tectonic.exe inside the exe and always uses
step 1.  The modular variant omits tectonic.exe and walks the full chain,
allowing users with a local TeX Live or MiKTeX to compile without Tectonic.

Note: ``palmer.sty`` loads ``fontspec``, which requires XeLaTeX.
pdflatex is *not* supported.

https://tectonic-typesetting.github.io/
"""

from __future__ import annotations

import io
import logging
import os
import re
import stat
import struct
import subprocess
import sys
import shutil
import tempfile
import threading as _threading
from collections import deque
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import pypdfium2 as pdfium
from PIL import Image, ImageChops

logger = logging.getLogger(__name__)

__all__ = [
    "PalmerCompiler",
    "TeXBackend",
    "FONT_PACKAGES",
    "find_bundled_tectonic",
    "find_local_latex",
    "find_tex_backend",
    "tectonic_cache_exists",
    "delete_tectonic_cache",
    "pdf_to_cropped_png",
    "copy_image_to_clipboard_win32",
    "validate_raw_input",
    "MAX_FIELD_LEN",
    "MAX_RAW_LEN",
    "MIN_DPI",
    "MAX_DPI",
    "DEFAULT_DPI",
    "MIN_FONT_SIZE_PT",
    "MAX_FONT_SIZE_PT",
    "DEFAULT_FONT_SIZE_PT",
    "DEFAULT_MARGIN_PX",
    "DEFAULT_FONT_FAMILY",
    "clamp_dpi",
    "POINTS_PER_INCH",
    "MM_PER_INCH",
]

# --- Shared constants ---

MIN_DPI = 72
MAX_DPI = 2400
DEFAULT_DPI = 600
MIN_FONT_SIZE_PT = 2.0
MAX_FONT_SIZE_PT = 144.0
DEFAULT_FONT_SIZE_PT = 10.0
DEFAULT_MARGIN_PX = 8
DEFAULT_FONT_FAMILY = "Times New Roman"

# Typesetting conversion factors
POINTS_PER_INCH: float = 72.0   # PDF points per inch (PostScript standard)
MM_PER_INCH: float = 25.4       # millimetres per inch (exact, by definition)
_LINE_HEIGHT_RATIO: float = 1.2  # default TeX \baselineskip / font-size ratio
_DPI_TO_PPM_FACTOR: int = 10_000  # numerator factor: pixels/metre = dpi * 10000 / 254
_DPI_TO_PPM_DIVISOR: int = 254    # denominator (254 = 25.4 mm/in × 10)


def clamp_dpi(value: str | int | float, fallback: int = DEFAULT_DPI) -> int:
    """Clamp *value* to the valid DPI range, returning *fallback* on error."""
    try:
        v = int(float(value))
    except (ValueError, TypeError):
        return fallback
    return max(MIN_DPI, min(MAX_DPI, v))


# --- TeX template ---

TEX_TEMPLATE = r"""\documentclass{{article}}
\usepackage{{fontspec}}
\usepackage{{palmer}}
{preamble}
\pagestyle{{empty}}
\begin{{document}}
\noindent {body}
\end{{document}}
"""

# --- Font configuration ---

# Maps display names to \setmainfont preamble commands.
# Fonts not listed here are also accepted by render() and resolved dynamically.
FONT_PACKAGES: dict[str, str] = {
    "Times New Roman": r"\setmainfont{Times New Roman}",
    "Arial":           r"\setmainfont{Arial}",
    "Calibri":         r"\setmainfont{Calibri}",
    "Georgia":         r"\setmainfont{Georgia}",
    "Cambria":         r"\setmainfont{Cambria}",
    "Yu Mincho":       r"\setmainfont{Yu Mincho}",
    "Yu Gothic":       r"\setmainfont{Yu Gothic}",
    "Meiryo":          r"\setmainfont{Meiryo}",
}

# TeX special characters that are unsafe in font names.
_FONT_NAME_UNSAFE: frozenset[str] = frozenset('\\{}$%^&~#')


def _get_font_preamble(font_family: str) -> str:
    """Return the \\setmainfont preamble line for the given font family.

    Returns the \\setmainfont command for "Times New Roman".
    Unknown font names are accepted and resolved dynamically.
    """
    if not font_family or not font_family.strip():
        raise ValueError("Font family name must not be empty")
    if font_family in FONT_PACKAGES:
        return FONT_PACKAGES[font_family]
    if len(font_family) > MAX_FIELD_LEN:
        raise ValueError(
            f"Font name exceeds the maximum length of {MAX_FIELD_LEN} characters"
        )
    bad = set(font_family) & _FONT_NAME_UNSAFE
    if bad:
        raise ValueError(
            f"The font name contains unsupported characters: {''.join(sorted(bad))}\n"
            f"Font name: {font_family!r}"
        )
    return rf"\setmainfont{{{font_family}}}"

# --- Input validation ---

MAX_FIELD_LEN = 256

# Deny-listed TeX commands for dental notation fields.
# File I/O, shell escape, and parser manipulation commands are rejected.
# Decorative commands (\textbf, \underline, etc.) are permitted.
_TEX_DANGEROUS_CMDS: frozenset[str] = frozenset({
    r'\write18',           # shell escape
    r'\write',             # arbitrary file write (\write18 is a subset)
    r'\read',              # file/terminal read
    r'\openout',           # file write handle
    r'\openin',            # file read handle
    r'\newwrite',          # allocate write handle
    r'\newread',           # allocate read handle
    r'\closeout',          # close write handle
    r'\closein',           # close read handle
    r'\immediate',         # force immediate \write / \openout
    r'\input',             # file inclusion
    r'\include',           # file inclusion (high-level \input)
    r'\inputiffileexists', # conditional file inclusion
    r'\catcode',           # category code modification
    r'\lccode',            # lowercase code modification
    r'\uccode',            # uppercase code modification
    r'\mathcode',          # math code modification
    r'\scantokens',        # re-execute an arbitrary token list
    r'\detokenize',        # detokenize a token list
    r'\csname',            # dynamic command construction (can bypass deny-list)
    r'\special',           # arbitrary DVI/PDF backend command
    r'\directlua',         # LuaTeX Lua execution
    r'\luadirect',         # LuaTeX Lua execution (alias)
    r'\luaexec',           # LuaTeX Lua execution (alias)
    r'\latelua',           # LuaTeX deferred Lua execution
    r'\shellescape',       # shell escape (alias)
    r'\string',            # command-to-string conversion (can bypass deny-list)
    r'\expandafter',       # macro expansion order manipulation
    r'\aftergroup',        # insert token after group close
    r'\futurelet',         # inspect/manipulate next token
    r'\everyjob',          # auto-execute hook at job start
    r'\everypar',          # auto-execute hook at paragraph start
    r'\everymath',         # auto-execute hook at math mode start
    r'\toks',              # token register manipulation
    r'\def',               # macro definition (can alias deny-listed commands)
    r'\edef',              # expanded macro definition
    r'\gdef',              # global macro definition
    r'\xdef',              # global expanded macro definition
    r'\let',               # command aliasing (can bypass deny-list)
    r'\global',            # global prefix (used with \def, \let)
    r'\long',              # long macro prefix
    r'\outer',             # outer macro prefix
    r'\unexpanded',        # prevent expansion (e-TeX)
    r'\lowercase',         # case conversion (can construct arbitrary tokens)
    r'\uppercase',         # case conversion (can construct arbitrary tokens)
})

# Midline fields accept plain text only; all TeX special characters are rejected.
_MIDLINE_DANGEROUS: frozenset[str] = frozenset('\\{}$%^&~#')

# --- Color validation ---

_COLOR_HEX_RE = re.compile(r'^#[0-9A-Fa-f]{6}$')
_COLOR_NAME_RE = re.compile(r'^[A-Za-z][A-Za-z0-9\-]*$')


def _validate_color(value: str) -> None:
    """Validate a color value: hex #RRGGBB or a named xcolor color.

    Raises ValueError if the value is non-empty and does not match either form.
    """
    if not value:
        return
    if _COLOR_HEX_RE.match(value) or _COLOR_NAME_RE.match(value):
        return
    raise ValueError(
        f"Invalid color: {value!r}\n"
        "Use a 6-digit hex color (#RRGGBB) or a named xcolor color "
        "(e.g., red, blue, cyan, darkgray)."
    )


def _build_color_tex(text_color: str) -> tuple[str, str]:
    """Return (xcolor_preamble, color_command) for the given color value.

    Returns ('', '') if text_color is empty (no color override).
    For hex colors (#RRGGBB) uses xcolor HTML syntax; for named colors uses
    the standard \\color{name} form.
    """
    if not text_color:
        return "", ""
    preamble = r"\usepackage{xcolor}"
    if _COLOR_HEX_RE.match(text_color):
        hex_val = text_color[1:].upper()
        color_cmd = rf"\color[HTML]{{{hex_val}}}"
    else:
        color_cmd = rf"\color{{{text_color}}}"
    return preamble, color_cmd


def _check_brace_balance(value: str, field: str) -> None:
    """Verify that braces in value are balanced.

    An unmatched closing brace would prematurely close the \\Palmer argument,
    enabling injection into the surrounding TeX source.
    Escaped braces (``\\{`` and ``\\}``) are treated as literal characters and
    do not affect the depth count.
    """
    depth = 0
    i = 0
    while i < len(value):
        if value[i] == '\\' and i + 1 < len(value) and value[i + 1] in '{}':
            i += 2  # skip escaped brace
            continue
        if value[i] == '{':
            depth += 1
        elif value[i] == '}':
            depth -= 1
            if depth < 0:
                raise ValueError(
                    f"'{field}' contains an unmatched closing brace: {value!r}\n"
                    r"Ensure TeX decoration command braces (e.g., \textbf{1}) are balanced."
                )
        i += 1
    if depth != 0:
        raise ValueError(
            f"'{field}' contains unmatched opening braces "
            f"({depth} unclosed '{{' braces): {value!r}"
        )


_DANGEROUS_CMD_RE = re.compile(
    "|".join(
        re.escape(cmd) + r"(?![A-Za-z])"
        for cmd in sorted(_TEX_DANGEROUS_CMDS, key=len, reverse=True)
    ),
)


def _check_no_dangerous_cmds(value: str, field: str) -> None:
    """Raise ValueError if value contains any deny-listed TeX command.

    Matching is case-sensitive (TeX commands are case-sensitive) and uses
    word-boundary-aware regex to avoid false positives (e.g. ``\\typewriter``
    no longer matches ``\\write``).
    The ``^^`` TeX hex escape notation is also rejected to prevent bypass.
    """
    if "^^" in value:
        raise ValueError(
            f"'{field}' contains TeX character escape sequences (^^), "
            "which are not permitted in dental notation fields."
        )
    m = _DANGEROUS_CMD_RE.search(value)
    if m:
        raise ValueError(
            f"'{field}' contains a disallowed TeX command: {m.group()}\n"
            "This command cannot be used to decorate dental notation."
        )


MAX_RAW_LEN = 2048


def validate_raw_input(value: str, field: str) -> None:
    """Validate raw TeX input: length limit, brace balance, and deny-listed commands."""
    if len(value) > MAX_RAW_LEN:
        raise ValueError(
            f"'{field}' exceeds the maximum length of {MAX_RAW_LEN} characters"
        )
    _check_brace_balance(value, field)
    _check_no_dangerous_cmds(value, field)


def _validate_tex_field(value: str, field: str) -> None:
    """Validate a dental notation field (length + brace balance + deny-list check)."""
    if len(value) > MAX_FIELD_LEN:
        raise ValueError(
            f"'{field}' exceeds the maximum length of {MAX_FIELD_LEN} characters"
        )
    _check_brace_balance(value, field)
    _check_no_dangerous_cmds(value, field)


def _validate_midline_field(value: str, field: str) -> None:
    """Validate a midline field (length + special character check).

    Midline fields are plain text only; TeX decorations are not supported.
    The _MIDLINE_DANGEROUS set includes all TeX special characters (\\, {, }, etc.),
    so brace-balance and dangerous-command checks are implicitly covered.
    """
    if len(value) > MAX_FIELD_LEN:
        raise ValueError(
            f"'{field}' exceeds the maximum length of {MAX_FIELD_LEN} characters"
        )
    bad = set(value) & _MIDLINE_DANGEROUS
    if bad:
        raise ValueError(
            f"'{field}' contains TeX special characters: {''.join(sorted(bad))}"
        )


# --- TeX backend ---

@dataclass
class TeXBackend:
    """Represents a TeX compiler executable and its invocation arguments."""
    name: str
    executable: str
    args: list[str] = field(default_factory=list)

    def compile(self, tex_path: Path, cwd: Path, timeout: int = 120) -> Path:
        """Compile .tex → .pdf. Returns path to the generated PDF."""
        cmd = [self.executable] + self.args + [str(tex_path.name)]
        logger.debug(
            "TeXBackend.compile: cmd=%s cwd=%s timeout=%d",
            cmd, cwd, timeout,
        )
        run_kwargs: dict[str, Any] = dict(
            cwd=str(cwd),
            capture_output=True,
            text=True,
            timeout=timeout,
        )
        # Suppress the console window that would briefly appear on Windows.
        if sys.platform == "win32":
            run_kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW
        try:
            result = subprocess.run(cmd, **run_kwargs)
        except subprocess.TimeoutExpired:
            logger.debug("TeXBackend.compile: timed out after %d seconds", timeout)
            raise RuntimeError(
                f"TeX compilation timed out after {timeout} seconds ({self.name}).\n"
                "The document may be too complex, or the engine is unresponsive."
            )
        logger.debug(
            "TeXBackend.compile: returncode=%d stdout_len=%d stderr_len=%d",
            result.returncode,
            len(result.stdout or ""),
            len(result.stderr or ""),
        )
        if result.returncode != 0:
            logger.debug("TeXBackend.compile: stderr=%s", (result.stderr or "")[:500])
        pdf_path = tex_path.with_suffix(".pdf")
        if result.returncode != 0 or not pdf_path.exists():
            log_path = tex_path.with_suffix(".log")
            log_tail = ""
            if log_path.exists():
                with open(log_path, errors="replace") as _lf:
                    log_tail = "".join(deque(_lf, maxlen=20))
            stderr_text = (result.stderr or "").strip()
            raise RuntimeError(
                f"TeX compilation failed ({self.name}):\n"
                f"Exit code: {result.returncode}\n"
                + (f"Stderr:\n{stderr_text}\n" if stderr_text else "")
                + f"Log tail:\n{log_tail}"
            )
        return pdf_path


def find_bundled_tectonic(bundled_dir: Path | None = None) -> TeXBackend:
    """Locate the bundled tectonic binary and return a TeXBackend for it.

    Searches in order: PyInstaller _MEIPASS, the executable's directory, and
    the directory containing this module.  The system PATH is never searched.

    Args:
        bundled_dir: Override all search paths with this directory.
    """
    if bundled_dir is not None:
        search_dirs = [bundled_dir]
    else:
        search_dirs = []

        # PyInstaller onefile/onedir: Tectonic is extracted to _MEIPASS/bin/.
        if hasattr(sys, "_MEIPASS"):
            search_dirs.append(Path(sys._MEIPASS) / "bin")

        # PyInstaller onedir: bin/ beside the executable.
        search_dirs.append(Path(sys.executable).parent / "bin")

        # Plain Python: bin/ beside this module.
        search_dirs.append(Path(__file__).parent / "bin")

    for bd in search_dirs:
        # NOTE: .exe suffix is Windows-specific.  This tool currently targets
        # Windows only; adjust to a platform-aware lookup if cross-platform
        # support is added in the future.
        exe = bd / "tectonic.exe"
        logger.debug("Checking bundled tectonic: %s (exists=%s)", exe, exe.exists())
        if exe.exists():
            logger.info("Found bundled tectonic: %s", exe)
            return TeXBackend("tectonic", str(exe), ["-X", "compile"])

    msg = (
        "Bundled tectonic.exe was not found.\n"
        "Searched paths:\n"
        + "\n".join(f"  {bd / 'tectonic.exe'}" for bd in search_dirs)
    )
    logger.debug(msg)
    raise FileNotFoundError(msg)


def find_local_latex() -> TeXBackend:
    """Search the system PATH for a XeLaTeX executable.

    ``palmer.sty`` loads ``fontspec``, which requires XeLaTeX.

    Raises ``FileNotFoundError`` if xelatex is not found.
    """
    exe = shutil.which("xelatex")
    if exe is not None:
        logger.info("Found system xelatex: %s", exe)
        return TeXBackend(
            "xelatex", exe,
            ["-interaction=nonstopmode", "-halt-on-error"],
        )
    logger.debug("xelatex not found on PATH")
    raise FileNotFoundError(
        "No local TeX installation found on the system PATH.\n"
        "Install TeX Live or MiKTeX so that xelatex is "
        "available, or place tectonic beside the executable."
    )


def find_tex_backend() -> TeXBackend:
    """Auto-detect the best available TeX backend.

    Priority:
      1. Bundled / adjacent ``tectonic`` (PyInstaller bundle or ``bin/``)
      2. System ``tectonic`` on PATH
      3. System ``xelatex``  on PATH  (TeX Live / MiKTeX)

    Raises ``FileNotFoundError`` when no usable engine is found.
    """
    logger.debug("Searching for TeX backend ...")

    # 1. Bundled / adjacent tectonic
    try:
        backend = find_bundled_tectonic()
        logger.info("TeX backend selected: %s (%s)", backend.name, backend.executable)
        return backend
    except FileNotFoundError:
        pass

    # 2. System tectonic on PATH
    tectonic_exe = shutil.which("tectonic")
    if tectonic_exe is not None:
        logger.info("TeX backend selected: system tectonic (%s)", tectonic_exe)
        return TeXBackend("tectonic", tectonic_exe, ["-X", "compile"])

    # 3. Local XeLaTeX
    try:
        backend = find_local_latex()
        logger.info("TeX backend selected: %s (%s)", backend.name, backend.executable)
        return backend
    except FileNotFoundError:
        pass

    raise FileNotFoundError(
        "No TeX engine found.\n"
        "Place tectonic in the bin/ folder beside the executable, or\n"
        "install TeX Live / MiKTeX so that xelatex is on PATH."
    )


# --- Tectonic cache detection ---

def _tectonic_cache_dir() -> Path | None:
    """Return the Tectonic cache directory, or None if undetermined.

    Tectonic v2 (the ``-X`` CLI) uses the ``directories`` Rust crate with
    ``ProjectDirs::from("", "TectonicProject", "Tectonic")``, which resolves
    to platform-specific locations:

    - Windows: ``%LOCALAPPDATA%\\TectonicProject\\Tectonic\\``
    - macOS:   ``~/Library/Caches/Tectonic``
    - Linux:   ``$XDG_CACHE_HOME/Tectonic`` (default ``~/.cache/Tectonic``)
    """
    if sys.platform == "win32":
        local = os.environ.get("LOCALAPPDATA")
        if not local:
            return None
        return Path(local) / "TectonicProject" / "Tectonic"
    elif sys.platform == "darwin":
        return Path.home() / "Library" / "Caches" / "Tectonic"
    else:
        # Linux / other POSIX
        xdg = os.environ.get("XDG_CACHE_HOME")
        base = Path(xdg) if xdg else Path.home() / ".cache"
        return base / "Tectonic"


def tectonic_cache_exists() -> bool:
    """Return True if the Tectonic cache directory exists and is non-empty.

    When False, the first compilation will trigger a ~100 MB download of TeX
    support files.
    """
    cache = _tectonic_cache_dir()
    logger.debug("tectonic_cache_exists: cache_dir=%s", cache)
    if cache is None or not cache.is_dir():
        logger.debug("tectonic_cache_exists: directory missing or None → False")
        return False
    try:
        result = next(cache.iterdir(), None) is not None
        logger.debug("tectonic_cache_exists: non_empty=%s", result)
        return result
    except OSError as exc:
        logger.debug("tectonic_cache_exists: OSError scanning dir: %s", exc)
        return False


def _rmtree_readonly(func, path, _exc_info):
    """Clear the read-only flag and retry removal.

    On Windows, Tectonic cache files are often read-only, causing
    ``shutil.rmtree`` to fail with *WinError 5 (Access Denied)*.
    This handler strips the read-only attribute and retries.
    """
    os.chmod(path, stat.S_IWRITE)
    func(path)


def delete_tectonic_cache() -> bool:
    """Delete the Tectonic cache directory.

    Returns True if the cache was deleted, False if it did not exist.
    Raises OSError on deletion failure.
    """
    cache = _tectonic_cache_dir()
    if cache is None or not cache.is_dir():
        return False
    if sys.version_info >= (3, 12):
        def _onexc(func, path, exc):
            os.chmod(path, stat.S_IWRITE)
            func(path)
        shutil.rmtree(cache, onexc=_onexc)
    else:
        shutil.rmtree(cache, onerror=_rmtree_readonly)
    return True


# --- PDF to PNG conversion ---

def pdf_to_cropped_png(
    pdf_path: Path,
    dpi: int = DEFAULT_DPI,
    padding_top: int = DEFAULT_MARGIN_PX,
    padding_bottom: int = DEFAULT_MARGIN_PX,
    padding_left: int = DEFAULT_MARGIN_PX,
    padding_right: int = DEFAULT_MARGIN_PX,
    bg_color: tuple[int, int, int] = (255, 255, 255),
    alpha: bool = False,
) -> Image.Image:
    """Render the first PDF page and return a cropped PIL Image.

    Uses pypdfium2 (PDFium); does not require poppler or Ghostscript.

    When *alpha* is ``True`` the returned image has a transparent background
    (mode ``RGBA``).  The PDF is still rendered against a white canvas for
    auto-crop detection, and pixels matching the background are then set to
    fully transparent.
    """
    with pdfium.PdfDocument(str(pdf_path)) as doc:
        if len(doc) == 0:
            raise RuntimeError("Failed to render PDF page to image (empty or corrupted page).")
        page = doc[0]
        bitmap = page.render(scale=dpi / POINTS_PER_INCH, rotation=0)
        try:
            img = bitmap.to_pil()
        finally:
            bitmap.close()

    # Auto-crop to content bounds.
    bg = Image.new("RGB", img.size, bg_color)
    diff = ImageChops.difference(img, bg)
    bbox = diff.getbbox()

    if bbox is None:
        if alpha:
            return img.convert("RGBA")  # type: ignore[no-any-return]
        return img  # type: ignore[no-any-return]  # image is entirely background color

    x0 = max(0, bbox[0] - padding_left)
    y0 = max(0, bbox[1] - padding_top)
    x1 = min(img.width, bbox[2] + padding_right)
    y1 = min(img.height, bbox[3] + padding_bottom)

    cropped = img.crop((x0, y0, x1, y1))

    if alpha:
        # Convert background pixels to transparent.
        rgba = cropped.convert("RGBA")
        bg_cropped = Image.new("RGB", cropped.size, bg_color)
        diff_cropped = ImageChops.difference(cropped, bg_cropped)
        # Any pixel that matches the background exactly gets alpha=0.
        mask = diff_cropped.convert("L").point(lambda v: 255 if v > 0 else 0)
        rgba.putalpha(mask)
        return rgba  # type: ignore[no-any-return]

    return cropped  # type: ignore[no-any-return]


# --- Palmer compiler ---

class PalmerCompiler:
    """Compile Palmer dental notation to a PNG image."""

    def __init__(
        self,
        sty_path: Path | None = None,
        backend: TeXBackend | None = None,
        dpi: int = DEFAULT_DPI,
        preamble: str = "",
        margin_top: int = DEFAULT_MARGIN_PX,
        margin_bottom: int = DEFAULT_MARGIN_PX,
        margin_left: int = DEFAULT_MARGIN_PX,
        margin_right: int = DEFAULT_MARGIN_PX,
    ):
        if sty_path is None:
            sty_path = self._find_sty()
        self.sty_path = Path(sty_path)
        if not self.sty_path.exists():
            raise FileNotFoundError(f"palmer.sty not found: {self.sty_path}")

        if backend is None:
            backend = find_tex_backend()
        self.backend = backend
        logger.info(
            "PalmerCompiler initialised: backend=%s exe=%s sty=%s dpi=%d",
            self.backend.name, self.backend.executable, self.sty_path, dpi,
        )

        if not (MIN_DPI <= dpi <= MAX_DPI):
            raise ValueError(f"DPI must be between {MIN_DPI} and {MAX_DPI}, got {dpi}")
        self.dpi = dpi
        self.preamble = preamble
        for name, val in [
            ("margin_top", margin_top), ("margin_bottom", margin_bottom),
            ("margin_left", margin_left), ("margin_right", margin_right),
        ]:
            if val < 0:
                raise ValueError(f"{name} must be non-negative, got {val}")
        self.margin_top = margin_top
        self.margin_bottom = margin_bottom
        self.margin_left = margin_left
        self.margin_right = margin_right

    def _find_sty(self) -> Path:
        """Locate palmer.sty, checking standard paths for all deployment modes.

        Searched in order:
          1. Same directory as this module file (source tree / editable install).
          2. PyInstaller bundle root (``sys._MEIPASS``); falls back to the
             current working directory (``'.'``) when running outside a bundle.
          3. Current working directory (last resort).
        """
        candidates = [
            Path(__file__).parent / "palmer.sty",
            Path(getattr(sys, '_MEIPASS', '.')) / "palmer.sty",
            Path("palmer.sty"),
        ]
        for p in candidates:
            if p.exists():
                return p
        raise FileNotFoundError("palmer.sty was not found.")

    def render(
        self,
        UL: str = "", UR: str = "", LR: str = "", LL: str = "",
        upper_mid: str = "", lower_mid: str = "",
        option: str = "base",
        font_family: str = DEFAULT_FONT_FAMILY,
        font_size_pt: float = 10.0,
        text_color: str = "",
        *,
        dpi: int | None = None,
        margin_top: int | None = None,
        margin_bottom: int | None = None,
        margin_left: int | None = None,
        margin_right: int | None = None,
        alpha: bool = False,
    ) -> Image.Image:
        """Render a Palmer dental notation diagram and return a PIL Image.

        Maps to the \\Palmer[option]{UL}{UR}{LR}{LL}{upper_mid}{lower_mid} command.

        font_family: key from FONT_PACKAGES, or any system font name accepted by fontspec.
        font_size_pt: point size; the cross line dimensions scale proportionally.
        text_color: optional color for the rendered text. Accepts a 6-digit hex
            value (#RRGGBB) or any named color recognised by the xcolor package
            (e.g., ``red``, ``blue``, ``darkgray``). Empty string (default) leaves
            the color unchanged (black).
        dpi: Override instance DPI for this call (thread-safe).
        margin_top/bottom/left/right: Override instance margins for this call.
        alpha: When True, the returned image has a transparent background (RGBA).
        """
        if option not in ("base", "center", "bottom"):
            raise ValueError(
                f"option must be one of 'base', 'center', 'bottom': {option!r}"
            )
        if not (MIN_FONT_SIZE_PT <= font_size_pt <= MAX_FONT_SIZE_PT):
            raise ValueError(
                f"font_size_pt must be in the range {MIN_FONT_SIZE_PT} to {MAX_FONT_SIZE_PT}: {font_size_pt}"
            )
        for _f, _v in [("UL", UL), ("UR", UR), ("LR", LR), ("LL", LL)]:
            _validate_tex_field(_v, _f)
        for _f, _v in [("upper_mid", upper_mid), ("lower_mid", lower_mid)]:
            _validate_midline_field(_v, _f)
        _validate_color(text_color)

        font_pkg = _get_font_preamble(font_family)
        color_preamble, color_cmd = _build_color_tex(text_color)
        leading_pt = font_size_pt * _LINE_HEIGHT_RATIO
        size_cmd = rf"\fontsize{{{font_size_pt}pt}}{{{leading_pt:.2f}pt}}\selectfont "
        palmer_cmd = (
            rf"\Palmer[{option}]"
            + f"{{{UL}}}{{{UR}}}{{{LR}}}{{{LL}}}"
            + f"{{{upper_mid}}}{{{lower_mid}}}"
        )
        if color_cmd:
            tex_body = "{" + color_cmd + " " + size_cmd + palmer_cmd + "}"
        else:
            tex_body = size_cmd + palmer_cmd
        extra_preambles = [p for p in [color_preamble, font_pkg] if p]
        return self.render_raw(
            tex_body,
            extra_preamble="\n".join(extra_preambles),
            validate=False,
            dpi=dpi,
            margin_top=margin_top,
            margin_bottom=margin_bottom,
            margin_left=margin_left,
            margin_right=margin_right,
            alpha=alpha,
        )

    def render_raw(
        self,
        tex_body: str,
        extra_preamble: str = "",
        *,
        validate: bool = True,
        dpi: int | None = None,
        margin_top: int | None = None,
        margin_bottom: int | None = None,
        margin_left: int | None = None,
        margin_right: int | None = None,
        alpha: bool = False,
        compile_timeout: int = 120,
    ) -> Image.Image:
        """Compile arbitrary TeX body in the palmer.sty environment.

        Args:
            validate: When True (default), applies length, brace-balance,
                and dangerous-command checks.  Internal callers (render())
                that have already validated their fields may pass False.
            dpi: Override instance DPI for this call (thread-safe).
            margin_top/bottom/left/right: Override instance margins for this call.
            alpha: When True, return an RGBA image with transparent background.
        """
        if validate:
            validate_raw_input(tex_body, "tex_body")
            if extra_preamble:
                validate_raw_input(extra_preamble, "extra_preamble")

        eff_dpi = dpi if dpi is not None else self.dpi
        eff_margin_top = margin_top if margin_top is not None else self.margin_top
        eff_margin_bottom = margin_bottom if margin_bottom is not None else self.margin_bottom
        eff_margin_left = margin_left if margin_left is not None else self.margin_left
        eff_margin_right = margin_right if margin_right is not None else self.margin_right

        preamble_parts = [p for p in [extra_preamble, self.preamble] if p]
        preamble = "\n".join(preamble_parts)

        logger.debug(
            "render_raw: backend=%s validate=%s compile_timeout=%d",
            self.backend.name, validate, compile_timeout,
        )
        with tempfile.TemporaryDirectory(prefix="palmer_") as tmp_str:
            tmpdir = Path(tmp_str)

            shutil.copy2(self.sty_path, tmpdir / "palmer.sty")

            tex_source = TEX_TEMPLATE.format(
                preamble=preamble,
                body=tex_body,
            )
            tex_path = tmpdir / "palmer.tex"
            tex_path.write_text(tex_source, encoding="utf-8")

            logger.debug("render_raw: calling backend.compile ...")
            pdf_path = self.backend.compile(tex_path, cwd=tmpdir, timeout=compile_timeout)
            logger.debug("render_raw: compile returned pdf=%s", pdf_path)

            return pdf_to_cropped_png(
                pdf_path,
                dpi=eff_dpi,
                padding_top=eff_margin_top,
                padding_bottom=eff_margin_bottom,
                padding_left=eff_margin_left,
                padding_right=eff_margin_right,
                alpha=alpha,
            )

    _SUPPORTED_FORMATS = frozenset({".png", ".jpg", ".jpeg", ".pdf"})

    def render_to_file(self, output: Path, **kwargs: Any) -> Path:
        """Render and save the result to a file (PNG, JPEG, or PDF).

        Remaining keyword arguments are forwarded to :meth:`render`.
        Accepted keys: ``UL``, ``UR``, ``LR``, ``LL``, ``upper_mid``,
        ``lower_mid``, ``option``, ``font_family``, ``font_size_pt``,
        ``text_color``, ``dpi``, ``margin_top``, ``margin_bottom``,
        ``margin_left``, ``margin_right``.

        When *alpha* is not explicitly provided in *kwargs*, it defaults to
        ``True`` for PNG output and ``False`` for JPEG/PDF (which do not
        support transparency).
        """
        eff_dpi = kwargs.get("dpi") if kwargs.get("dpi") is not None else self.dpi
        output = Path(output)
        suffix = output.suffix.lower()
        if suffix not in self._SUPPORTED_FORMATS:
            raise ValueError(
                f"Unsupported output format: {output.suffix}\n"
                f"Supported formats: {', '.join(sorted(self._SUPPORTED_FORMATS))}"
            )
        # Auto-detect alpha based on output format when not explicitly set.
        if "alpha" not in kwargs:
            kwargs["alpha"] = suffix == ".png"
        img = self.render(**kwargs)
        if suffix in (".jpg", ".jpeg") and img.mode == "RGBA":
            img = img.convert("RGB")
        img.save(str(output), dpi=(eff_dpi, eff_dpi))
        return output

    def render_to_clipboard(self, **kwargs: Any) -> Image.Image:
        """Render and copy the result to the Windows clipboard.

        Keyword arguments are forwarded to :meth:`render`.  See
        :meth:`render_to_file` for the list of accepted keys.
        """
        img = self.render(**kwargs)
        _raw_dpi = kwargs.get("dpi")
        eff_dpi = int(_raw_dpi) if _raw_dpi is not None else self.dpi
        copy_image_to_clipboard_win32(img, dpi=eff_dpi)
        return img


# Guard for one-time ctypes initialisation.  Protected by a lock for
# thread safety even though clipboard operations are expected to run on
# the main thread only (Win32 clipboard API requirement).
_clipboard_lock = _threading.Lock()
_clipboard_initialized = False


def _init_clipboard_ctypes() -> None:
    """Configure ctypes argtypes/restype for Win32 clipboard functions.

    Called once on first use.  Setting restype to c_void_p on 64-bit Windows
    is critical — without it, ctypes truncates the 64-bit pointer to 32 bits,
    causing GlobalLock to return a spurious NULL.
    """
    global _clipboard_initialized
    if _clipboard_initialized:
        return
    with _clipboard_lock:
        if _clipboard_initialized:
            return
        import ctypes
        kernel32 = ctypes.windll.kernel32  # type: ignore[attr-defined]
        user32 = ctypes.windll.user32  # type: ignore[attr-defined]
        kernel32.GlobalAlloc.restype = ctypes.c_void_p
        kernel32.GlobalAlloc.argtypes = [ctypes.c_uint, ctypes.c_size_t]
        kernel32.GlobalLock.restype = ctypes.c_void_p
        kernel32.GlobalLock.argtypes = [ctypes.c_void_p]
        kernel32.GlobalUnlock.argtypes = [ctypes.c_void_p]
        kernel32.GlobalFree.restype = ctypes.c_void_p
        kernel32.GlobalFree.argtypes = [ctypes.c_void_p]
        user32.SetClipboardData.restype = ctypes.c_void_p
        user32.SetClipboardData.argtypes = [ctypes.c_uint, ctypes.c_void_p]
        _clipboard_initialized = True


def copy_image_to_clipboard_win32(img: Image.Image, dpi: int = DEFAULT_DPI) -> None:
    """Copy a PIL Image to the Windows clipboard as CF_DIB (DPI-stamped) and CF_PNG.

    Embedding DPI in both formats ensures that Word and other applications
    render the pasted image at the same physical size as a saved PNG file.
    """
    import ctypes

    _init_clipboard_ctypes()

    # CF_DIB requires RGB (BMP has no alpha support).
    # If the source is RGBA, composite onto white first.
    if img.mode == "RGBA":
        white_bg = Image.new("RGBA", img.size, (255, 255, 255, 255))
        white_bg.paste(img, mask=img.split()[3])
        rgb_img = white_bg.convert("RGB")
    else:
        rgb_img = img.convert("RGB")

    # CF_DIB: convert to BMP, strip the 14-byte file header to get a raw DIB,
    # then patch biXPelsPerMeter (offset 24) and biYPelsPerMeter (offset 28)
    # in the BITMAPINFOHEADER so Word scales the image correctly.
    output = io.BytesIO()
    rgb_img.save(output, "BMP")
    dib_data = bytearray(output.getvalue()[14:])
    pixels_per_meter = round(dpi * _DPI_TO_PPM_FACTOR / _DPI_TO_PPM_DIVISOR)  # 1 inch = 0.0254 m
    struct.pack_into("<i", dib_data, 24, pixels_per_meter)
    struct.pack_into("<i", dib_data, 28, pixels_per_meter)

    # CF_PNG: preserve alpha if present; modern Word prefers this format.
    png_output = io.BytesIO()
    img.save(png_output, "PNG", dpi=(dpi, dpi))
    png_data = png_output.getvalue()

    CF_DIB = 8
    GHND = 0x0042

    kernel32 = ctypes.windll.kernel32  # type: ignore[attr-defined]
    user32 = ctypes.windll.user32  # type: ignore[attr-defined]

    CF_PNG = user32.RegisterClipboardFormatA(b"PNG")
    if not CF_PNG:
        logger.warning(
            "RegisterClipboardFormatA('PNG') failed; CF_PNG format will not be set"
        )

    def _alloc_and_write(data):
        hMem = kernel32.GlobalAlloc(GHND, len(data))
        if not hMem:
            raise MemoryError("Failed to allocate global memory for clipboard.")
        pMem = kernel32.GlobalLock(hMem)
        if not pMem:
            kernel32.GlobalFree(hMem)
            raise OSError("Failed to lock global memory.")
        ctypes.memmove(pMem, bytes(data), len(data))
        kernel32.GlobalUnlock(hMem)
        return hMem

    if not user32.OpenClipboard(None):
        raise OSError("Failed to open the clipboard. Another application may be using it.")
    try:
        user32.EmptyClipboard()

        hDib = _alloc_and_write(dib_data)
        if not user32.SetClipboardData(CF_DIB, hDib):
            kernel32.GlobalFree(hDib)
            raise OSError("Failed to set clipboard data.")

        if CF_PNG:
            hPng = _alloc_and_write(png_data)
            if not user32.SetClipboardData(CF_PNG, hPng):
                kernel32.GlobalFree(hPng)  # non-fatal; CF_DIB already committed
    finally:
        user32.CloseClipboard()
