"""
build_exe.py - Build Windows executables with PyInstaller (and optionally an installer).

Produces up to three distribution variants:

  Bundled (default):
    dist/palmer-type-{version}.exe
    Single-file exe with tectonic.exe inside the bundle.

  Modular (--modular):
    dist/palmer-type-modular-{version}.exe
    Single-file exe WITHOUT tectonic.exe; auto-detects tectonic in bin/ or on
    PATH, then falls back to xelatex (TeX Live / MiKTeX) on PATH.

  Installer (--installer):
    dist/palmer-type-{version}-win-x64-setup.exe
    Inno Setup installer wrapping a PyInstaller --onedir build with tectonic
    bundled.  Unlike --onefile, --onedir does NOT self-extract to %TEMP%, so
    antivirus false positives are avoided.

Note on antivirus false positives:
    --onefile self-extracts to %TEMP%\\onefile_XXXXX\\ at runtime, which can
    trigger Windows Defender.  Signing the exe with a code-signing certificate
    resolves this.  The --installer variant avoids this issue entirely.

Usage:
    python build_exe.py               # Bundled variant only
    python build_exe.py --modular     # Modular variant only
    python build_exe.py --all         # Bundled + Modular + Installer
    python build_exe.py --installer   # Installer only

Prerequisites:
    pip install pyinstaller
    Inno Setup 6 (for --installer / --all)

Output:
    dist/palmer-type-{version}.exe                   (bundled)
    dist/palmer-type-modular-{version}.exe           (modular)
    dist/palmer-type-{version}-win-x64-setup.exe     (installer)
"""

import argparse
import ast
import io
import shutil
import subprocess
import sys
import tempfile
import tokenize
from contextlib import contextmanager
from pathlib import Path
from typing import Iterator

BASE_DIR = Path(__file__).parent

sys.path.insert(0, str(BASE_DIR))
from version import __version__  # noqa: E402

if __version__ == "dev":
    print(
        "WARNING: Building with version='dev'. "
        "Set a version tag or edit version.py for release builds.",
        file=sys.stderr,
    )

_tectonic_exe = BASE_DIR / "bin" / "tectonic.exe"
_app_icon = BASE_DIR / "assets" / "palmer-type.ico"


def _require_tectonic() -> None:
    """Exit with error if tectonic.exe is not found (bundled build only)."""
    if not _tectonic_exe.exists():
        print(
            f"ERROR: {_tectonic_exe} was not found.\n"
            "Please retrieve src/bin/tectonic.exe from the repository.",
            file=sys.stderr,
        )
        sys.exit(1)


# --add-data format for PyInstaller 6.x: "source:dest" (colon separator on all platforms).
# Files are extracted to sys._MEIPASS at runtime; find_bundled_tectonic() and
# _find_sty() in palmer_engine.py look for them there.
DATA_FILES_COMMON: list[str] = [
    f"--add-data={BASE_DIR / 'palmer.sty'}:.",                              # → _MEIPASS/palmer.sty
    f"--add-data={BASE_DIR / 'assets' / 'palmer-type.ico'}:assets",        # → _MEIPASS/assets/palmer-type.ico
]
DATA_FILES_TECTONIC: list[str] = [
    f"--add-data={_tectonic_exe}:bin",           # → _MEIPASS/bin/tectonic.exe
]

def _docstring_ranges(source: str) -> list[tuple[int, int, bool]]:
    """Return ``(start_line, end_line, needs_pass)`` 1-based ranges for all docstrings.

    *needs_pass* is ``True`` when the docstring is the only statement in the
    body of a class or function — removing it would leave an empty body that
    causes ``IndentationError``.
    """
    try:
        tree = ast.parse(source)
    except SyntaxError:
        return []

    ranges: list[tuple[int, int, bool]] = []

    def _check(body: list[ast.stmt], is_body: bool = False) -> None:
        if (
            body
            and isinstance(body[0], ast.Expr)
            and isinstance(body[0].value, ast.Constant)
            and isinstance(body[0].value.value, str)
        ):
            node = body[0]
            assert node.end_lineno is not None
            needs_pass = is_body and len(body) == 1
            ranges.append((node.lineno, node.end_lineno, needs_pass))

    _check(tree.body)
    for node in ast.walk(tree):
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
            _check(node.body, is_body=True)

    return ranges


def _strip_comments(source: str) -> str:
    """Remove ``#`` comments and docstrings from Python source code.

    Uses :mod:`tokenize` for comments (so ``#`` inside strings is preserved)
    and :mod:`ast` for docstrings (module/class/function level).
    Line structure is kept intact (removed lines become blank) to avoid
    breaking tracebacks.
    """
    lines = source.splitlines(True)

    # --- Remove docstrings (replace with blank lines) ---
    for start, end, needs_pass in _docstring_ranges(source):
        for i in range(start - 1, end):
            lines[i] = "\n" if lines[i].endswith("\n") else ""
        if needs_pass:
            # Preserve the indentation of the docstring line and insert ``pass``
            # so the class/function body is not left empty.
            orig_line = source.splitlines(True)[start - 1]
            indent = orig_line[: len(orig_line) - len(orig_line.lstrip())]
            lines[start - 1] = f"{indent}pass\n"

    # --- Remove # comments ---
    comments: list[tokenize.TokenInfo] = []
    try:
        for tok in tokenize.generate_tokens(io.StringIO(source).readline):
            if tok.type == tokenize.COMMENT:
                comments.append(tok)
    except tokenize.TokenError:
        return "".join(lines)  # Return with docstrings removed at least

    for tok in reversed(comments):
        lineno = tok.start[0] - 1  # 0-based index
        col = tok.start[1]
        line = lines[lineno]
        before = line[:col].rstrip()
        newline = "\n" if line.endswith("\n") else ""
        if before:
            # Inline comment – keep the code, drop the comment.
            lines[lineno] = before + newline
        else:
            # Full-line comment – replace with blank line.
            lines[lineno] = newline

    return "".join(lines)


@contextmanager
def _minified_source(src_dir: Path) -> Iterator[Path]:
    """Copy ``*.py`` from *src_dir* to a temp dir with comments stripped.

    Yields the temporary directory path.  Non-Python files (assets, binaries,
    etc.) are NOT copied – ``--add-data`` flags still reference the originals.
    """
    tmp = Path(tempfile.mkdtemp(prefix="palmer_build_"))
    try:
        for py_file in src_dir.glob("*.py"):
            original = py_file.read_text(encoding="utf-8")
            stripped = _strip_comments(original)
            (tmp / py_file.name).write_text(stripped, encoding="utf-8")
        print(f"Minified sources prepared in {tmp}")
        yield tmp
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def build_exe(modular: bool = False) -> str:
    """Build a unified GUI/CLI exe.

    Built with --noconsole so no console window appears in GUI mode.
    In CLI mode, palmer_main.py calls AttachConsole(-1) to reconnect
    stdout/stderr to the parent terminal.

    Returns the output exe name (without path).
    """
    if modular:
        name = f"palmer-type-modular-{__version__}"
        data_files = DATA_FILES_COMMON
    else:
        _require_tectonic()
        name = f"palmer-type-{__version__}"
        data_files = DATA_FILES_COMMON + DATA_FILES_TECTONIC

    variant = "modular" if modular else "bundled"
    print(f"=== Building {variant} exe ({name}.exe) with PyInstaller ===")

    icon_args = [f"--icon={_app_icon}"] if _app_icon.exists() else []

    with _minified_source(BASE_DIR) as src:
        cmd = [
            sys.executable, "-m", "PyInstaller",
            "--onefile",
            "--noconsole",
            f"--name={name}",
            "--distpath=dist",
            "--collect-binaries=pypdfium2",
            "--hidden-import=pypdfium2",
            "--collect-submodules=docx",
            "--collect-submodules=lxml",
            "--hidden-import=palmer_converter",
            f"--paths={src}",
            f"--add-data={src / 'palmer_converter.py'}:.",
            *icon_args,
            *data_files,
            str(src / "palmer_main.py"),
        ]
        print(" ".join(str(a) for a in cmd))
        subprocess.run(cmd, check=True)
    print(f"OK: dist/{name}.exe")
    return f"{name}.exe"


def build_onedir() -> str:
    """Build a PyInstaller --onedir bundle with tectonic.

    The output directory is ``dist/palmer-type/``.  Unlike --onefile, --onedir
    does NOT self-extract to %TEMP% at runtime, avoiding antivirus false
    positives.

    Returns the output directory name (e.g. 'palmer-type').
    """
    _require_tectonic()
    name = "palmer-type"
    data_files = DATA_FILES_COMMON + DATA_FILES_TECTONIC

    print(f"=== Building onedir bundle ({name}/) with PyInstaller ===")

    icon_args = [f"--icon={_app_icon}"] if _app_icon.exists() else []

    with _minified_source(BASE_DIR) as src:
        cmd = [
            sys.executable, "-m", "PyInstaller",
            "--onedir",
            "--noconsole",
            f"--name={name}",
            "--distpath=dist",
            "--collect-binaries=pypdfium2",
            "--hidden-import=pypdfium2",
            "--collect-submodules=docx",
            "--collect-submodules=lxml",
            "--hidden-import=palmer_converter",
            f"--paths={src}",
            f"--add-data={src / 'palmer_converter.py'}:.",
            *icon_args,
            *data_files,
            str(src / "palmer_main.py"),
        ]
        print(" ".join(str(a) for a in cmd))
        subprocess.run(cmd, check=True)
    print(f"OK: dist/{name}/")
    return name


def _find_iscc() -> str:
    """Locate the Inno Setup compiler (ISCC.exe).

    Search order:
      1. ISCC on PATH (e.g. added by Chocolatey)
      2. Default Inno Setup 6 install location
    """
    iscc = shutil.which("ISCC")
    if iscc:
        return iscc

    default = Path(r"C:\Program Files (x86)\Inno Setup 6\ISCC.exe")
    if default.exists():
        return str(default)

    print(
        "ERROR: ISCC.exe (Inno Setup compiler) was not found.\n"
        "Install Inno Setup 6 from https://jrsoftware.org/isinfo.php\n"
        "or run: choco install innosetup",
        file=sys.stderr,
    )
    sys.exit(1)


def build_installer() -> str:
    """Build the Inno Setup installer from a --onedir bundle.

    Calls build_onedir() first, then invokes ISCC.exe to compile
    installer.iss into a setup executable.

    Returns the output installer filename.
    """
    build_onedir()

    iscc = _find_iscc()
    iss_path = BASE_DIR / "installer.iss"
    installer_name = f"palmer-type-{__version__}-win-x64-setup.exe"

    print(f"=== Building installer ({installer_name}) with Inno Setup ===")

    cmd = [
        iscc,
        f"/DAppVersion={__version__}",
        str(iss_path),
    ]
    print(" ".join(cmd))
    subprocess.run(cmd, check=True)
    print(f"OK: dist/{installer_name}")
    return installer_name


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Build palmer-type executables with PyInstaller",
    )
    group = parser.add_mutually_exclusive_group()
    group.add_argument(
        "--modular", action="store_true",
        help="Build modular variant only (without bundled tectonic.exe)",
    )
    group.add_argument(
        "--all", action="store_true",
        help="Build all variants (bundled + modular + installer)",
    )
    parser.add_argument(
        "--installer", action="store_true",
        help="Build Inno Setup installer (PyInstaller --onedir + tectonic)",
    )
    return parser.parse_args()


if __name__ == "__main__":
    args = parse_args()

    if args.all:
        bundled = build_exe(modular=False)
        modular = build_exe(modular=True)
        installer = build_installer()
        print("\nBuild complete (all variants)!")
        print(f"   dist/{bundled}")
        print(f"   dist/{modular}")
        print(f"   dist/{installer}")
    elif args.modular:
        modular = build_exe(modular=True)
        print("\nBuild complete!")
        print(f"   dist/{modular}")
    elif args.installer:
        installer = build_installer()
        print("\nBuild complete!")
        print(f"   dist/{installer}")
    else:
        bundled = build_exe(modular=False)
        print("\nBuild complete!")
        print(f"   dist/{bundled}")
