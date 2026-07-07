# Internals

Developer reference for Palmer Dental Notation Type.

## File Layout

```
palmer-tool-dev/
├── src/
│   ├── palmer_main.py        # Unified entry point (GUI / CLI selector)
│   ├── palmer_type.py        # GUI application (tkinter)
│   ├── palmer_cli.py         # CLI interface
│   ├── palmer_engine.py      # Core engine: TeX backend, PDF→PNG, clipboard
│   ├── palmer_converter.py   # Word document converter (replace \Palmer → images)
│   ├── config.py             # Persistent JSON-backed configuration
│   ├── build_exe.py          # PyInstaller build script
│   ├── installer.iss         # Inno Setup script for Windows installer
│   ├── version.py            # Version string
│   ├── requirements.txt      # Runtime Python dependencies
│   ├── requirements-dev.txt  # Development dependencies (PyInstaller)
│   └── assets/               # Icons (ICO, PNG) for app and installer
├── tests/
│   ├── test_validation.py    # Input validation tests
│   ├── test_alt_text.py      # Alt-text generation tests
│   ├── test_extract_font.py  # Font extraction tests
│   ├── test_find_palmer_commands.py  # Command parser tests
│   ├── test_collect_table_paras.py   # Table paragraph collection tests
│   ├── test_strip_and_expand.py      # TeX stripping / range expansion tests
│   ├── test_config.py        # Configuration persistence tests
│   └── test.docx             # Sample document for integration tests
├── .github/
│   └── workflows/
│       ├── build.yml         # CI/CD: build .exe on tag push
│       └── typecheck.yml     # CI: mypy type checking on push
├── docs/
│   ├── index.html            # GitHub Pages landing page (EN/JA bilingual)
│   ├── style.css             # Stylesheet for the landing page
│   ├── script.js             # Language switching logic (i18n, localStorage)
│   └── images/
│       └── screenshot.png    # GUI screenshot embedded in the landing page
├── LICENSE                   # GNU AGPLv3 — palmer-type's own license
├── THIRD-PARTY-NOTICES.txt   # bundled third-party license texts
├── mypy.ini                  # mypy configuration
├── README.md
└── INTERNALS.md
```

## TeX Backend

The application uses the bundled **Tectonic** binary (`bin/tectonic.exe`).
Tectonic is a self-contained TeX engine (~30 MB) that requires no TeX Live or MiKTeX installation.

`palmer_engine.py` locates the bundled binary in this order:

| Search path | Description |
|---|---|
| `sys._MEIPASS/bin/` | PyInstaller onefile/onedir extraction directory |
| `<executable dir>/bin/` | Beside the built `.exe` |
| `<module dir>/bin/` | Beside `palmer_engine.py` (development layout) |

## Building Executables

The build produces up to three distribution variants, all behaving as GUI when launched with no arguments and as CLI when launched with arguments.

### Prerequisites

```powershell
pip install pyinstaller
# tectonic.exe must exist in src/bin/ before building
# Inno Setup 6 required for --installer (https://jrsoftware.org/isinfo.php)
```

### Build

```powershell
python src/build_exe.py               # Bundled portable exe only
python src/build_exe.py --modular     # Modular portable exe only
python src/build_exe.py --installer   # Inno Setup installer only
python src/build_exe.py --all         # All three variants
```

| Variant | Mode | Output |
|---|---|---|
| Bundled | `--onefile` | `dist/palmer-type-{version}.exe` |
| Modular | `--onefile` | `dist/palmer-type-modular-{version}.exe` |
| Installer | `--onedir` → Inno Setup | `dist/palmer-type-{version}-win-x64-setup.exe` |

### Portable exe (--onefile)

PyInstaller bundles `palmer.sty`, `tectonic.exe`, `LICENSE`, and `THIRD-PARTY-NOTICES.txt` into the executable. At runtime they are extracted to `%TEMP%\onefile_XXXXX\` (`sys._MEIPASS`). The installer variant additionally places `LICENSE.txt` and `THIRD-PARTY-NOTICES.txt` at the install root for discoverability. `palmer_engine.py` looks for them there via `find_bundled_tectonic()` and `_find_sty()`.

The exe is built with `--noconsole` to suppress the console window in GUI mode. In CLI mode, `palmer_main.py` calls `AttachConsole(-1)` (Win32) to reconnect stdout/stderr to the parent terminal.

> **Antivirus note:** `--onefile` self-extraction to `%TEMP%` can trigger Windows Defender false positives. Signing the exe with a code-signing certificate resolves this.

### Installer (--onedir + Inno Setup)

The `--installer` flag first builds a PyInstaller `--onedir` bundle to `dist/palmer-type/`, then invokes the Inno Setup compiler (`ISCC.exe`) with `src/installer.iss` to produce a setup executable.

Unlike `--onefile`, the `--onedir` layout keeps all files unpacked in the installation directory — there is no self-extraction to `%TEMP%` at runtime, which avoids Windows Defender false positives entirely.

`sys._MEIPASS` in a `--onedir` build points to the output directory itself (the same directory as the exe). The existing `find_bundled_tectonic()` and `_find_sty()` search paths work without modification.

The installer:
- Installs to `Program Files\PalmerType` (or per-user `AppData` without admin)
- Creates Start Menu shortcuts and an optional desktop shortcut
- Registers an uninstaller accessible from Settings → Apps
- Filename includes version and architecture: `palmer-type-{version}-win-x64-setup.exe`

The Inno Setup script (`src/installer.iss`) receives the version at compile time via `/DAppVersion={version}`.

### CI/CD

`.github/workflows/build.yml` runs on manual dispatch or a `v*` tag push. It installs dependencies and Inno Setup on a `windows-2022` runner, updates `version.py` from the git tag, builds all three variants, and uploads them as release assets.
