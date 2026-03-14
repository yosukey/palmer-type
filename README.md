# palmer-type: Render Zsigmondy-Palmer Dental Notation as Images

Render Zsigmondy-Palmer dental notation symbols as images and paste them directly into other apps such as MS Word or MS PowerPoint.

Zsigmondy-Palmer dental notation is powerful but notoriously hard to type —
special symbols, quadrant brackets, and tooth numbering all make direct input a
chore. **palmer-type.exe** takes that pain away: just fill in the fields and
get crisp, publication-ready images in seconds. Designed for dentists and
dental educators who want professional notation in their documents without any
TeX expertise.

## What is this?

palmer-type.exe compiles dental chart notation written with [palmer.sty](https://github.com/yosukey/palmer-latex) into PNG, JPEG, or PDF images. It bundles **Tectonic** as its TeX engine, so end users need neither TeX Live nor MiKTeX.

The tool ships as a single Windows executable (`palmer-type.exe`) that works as a GUI when launched without arguments and as a CLI when launched with arguments. The GUI includes a **Converter** tab that can open a Word document (`.docx`) containing `\Palmer` commands and replace them with rendered images in one step.

## Download

Three variants are available on the [Releases](../../releases) page:

| File | Tectonic | Notes |
|---|---|---|
| `palmer-type-{ver}-win-x64-setup.exe` | Bundled | **Recommended.** Windows installer — installs to Program Files with Start Menu shortcuts and an uninstaller. |
| `palmer-type-{ver}.exe` | Bundled inside the exe | Portable single-file exe — place anywhere and run |
| `palmer-type-modular-{ver}.exe` | **Not bundled** | Portable single-file exe — Tectonic is auto-detected; falls back to xelatex if not found |

> **`palmer-type-modular.exe` (Tectonic non-bundled):** This variant does not include
> Tectonic inside the exe. At startup it searches for a usable TeX engine in the
> following priority order:
>
> 1. `bin\tectonic.exe` in the same folder as the exe
> 2. `tectonic` on the system PATH
> 3. `xelatex` on the system PATH (TeX Live / MiKTeX)
>
> If none of these are found, the Render button is disabled and an error dialog is shown.
> The same auto-detection order applies to `palmer-type.exe` as well — the bundled
> Tectonic is simply the first candidate found.

The notes below (`Windows Defender`, `First run`) apply to
`palmer-type.exe` (the Tectonic-bundled variant).

> **Windows Defender / antivirus:** The portable exe variants self-extract to `%TEMP%` at startup, which may trigger a false positive. If that happens, allow the file and run it again. The **installer variant** does not self-extract to `%TEMP%`, so this issue does not apply.

### First run

The initial launch requires an **internet connection**. Tectonic automatically downloads the TeX support files it needs (~100 MB) and caches them locally. Subsequent launches work offline.

## GUI Usage

If you used the installer, launch **palmer-type** from the Start Menu. For the portable versions, double-click `palmer-type.exe` directly.

You can also launch it from a terminal:

```powershell
palmer-type.exe
```

### Step-by-step workflow

1. **Enter tooth notation** in the quadrant fields arranged as a cross.

   The layout follows the standard Palmer dental chart convention:

   ```
   patient's right  |  patient's left
   
              UR    |    UL
             -------+-------
              LR    |    LL
   ```

   Each field accepts tooth numbers (e.g. `12345678`) and optional LaTeX decorators.
   The **Upper Mid** and **Lower Mid** fields accept plain-text symbols placed at the midline.

   The **No L/R** checkbox suppresses the vertical midline and disables the right-side quadrant fields (UR/LR) and both Mid fields. Use this when you want to notate teeth without left/right distinction.

   > **Input order for UR and LR:** Enter tooth numbers in their natural numerical order
   > (1, 2, 3, … 8) — do not rearrange them to match the chart's visual layout.
   > The display reversal is applied automatically.

2. **Choose font, size, and color** using the controls at the top of the window.
   The font dropdown lists all fonts installed on your Windows system; Times New Roman is the built-in default.
   Check the **Favorite** checkbox next to a font to mark it as a favorite; favorite fonts appear at the top of the dropdown for quick access.
   Font size can be set in 0.5 pt increments (2–144 pt).
   Font color defaults to black; choose **Custom** to pick an arbitrary color.

3. **Set background color** (optional):
   - White (default)
   - Custom color — opens a color picker dialog
   - Transparent — PNG and clipboard output

4. **Click Render** (or press the ▶ button). The image appears in the preview pane with a scale indicator. Rendering runs in a background thread so the UI stays responsive. The active TeX engine (e.g. `Engine: tectonic`) is displayed next to the Render button.

5. **Copy to Clipboard** — sends the image to the Windows clipboard.

6. **Save** — exports the image to a file. Supported formats: PNG, JPEG, PDF. A DPI selector controls output resolution.

The status bar at the bottom shows rendering status and image dimensions.

### Debug mode

Enable **Debug** from the menu bar to open a log panel at the bottom of the window. It shows timestamped diagnostic information — platform details, TeX engine detection, CJK font fallback, rendering progress, and errors. In the CLI, use `--verbose` / `-v` to enable the equivalent debug output on stderr.

### Converter tab — replace `\Palmer` commands in Word documents

The **Converter** tab lets you open a `.docx` file that contains `\Palmer` commands written as plain text and automatically replace each one with a rendered image — no manual rendering or pasting required.

> **Note:** This feature is available only in the GUI. The CLI does not have a `--docx` option.

#### How it works

1. **Browse** to select a `.docx` file.
2. Set the **DPI** (default: 600), **Alt text** mode, and **Vertical alignment** mode.

   **Alt text** adds an accessibility description to each rendered image:

   | Mode | Description |
   |---|---|
   | None | No alt text (default) |
   | FDI | FDI two-digit numbering (e.g. 11, 21, …) |
   | Universal | Universal numbering (1–32 for permanent teeth) |
   | Anatomical | Anatomical tooth names (e.g. "Right Maxillary Central incisor") |
   | Alphanumeric | Alphanumeric notation (e.g. "UR1") |
   | Palmer command | Uses the original `\Palmer{…}` command text as alt text |

   **Vertical alignment** controls the inline vertical positioning of each image in the document:

   | Mode | Description |
   |---|---|
   | Force center | Always center every image vertically on the text baseline (default) |
   | Follow command option | By default the bottom edge of each image aligns with the bottom of the line. `\Palmer[center]{…}` changes this to vertical center alignment. |

3. Click **Convert**.
4. Choose a save mode:
   - **Save as a new file (recommended)** — the output is written to `<original_name>_palmered.docx`. A "Save As" dialog lets you change the name and location.
   - **Overwrite the original file** — a confirmation dialog is shown before proceeding.
5. The log area shows progress. When finished, a summary dialog reports how many commands were replaced and whether any errors occurred.

#### Writing `\Palmer` commands in Word

Type the commands directly in your Word document using the same syntax as `palmer.sty`:

```
\Palmer{UL}{UR}{LR}{LL}{upper_mid}{lower_mid}
\Palmer[option]{UL}{UR}{LR}{LL}{upper_mid}{lower_mid}
```

Examples:

| Text in Word | Result |
|---|---|
| `\Palmer{1}{}{}{}{}{}` | Upper-left quadrant, tooth 1 |
| `\Palmer{12345678}{12345678}{12345678}{12345678}{}{}` | Full dentition |
| `\Palmer[center]{123}{123}{}{}{}{}` | Upper anterior region (vertically centered) |

#### Font and size detection

The Converter reads the **font name** and **font size** set on the `\Palmer` command text in Word and uses them for rendering. For example, if you type `\Palmer{1}{}{}{}{}{}` in Times New Roman 12 pt, the image will be rendered in Times New Roman at 12 pt.

If the font or size cannot be determined (e.g. inherited from a style), the defaults are used (Times New Roman, 10 pt).

#### Error handling

If an individual `\Palmer` command fails to render (e.g. due to a syntax error), the original text is kept as-is in the output document and the error is recorded in the log. Other commands in the same document are still processed.

## CLI Usage

Pass any argument to `palmer-type.exe` to activate CLI mode:

```powershell
palmer-type.exe [options]
```

### Quick examples

```powershell
# Single tooth (upper-left quadrant of the cross)
palmer-type.exe --UL 1 -o UL1.png

# Full dentition
palmer-type.exe --UL 12345678 --UR 12345678 --LR 12345678 --LL 12345678 -o full.png

# Copy to clipboard (Windows)
palmer-type.exe --UR 6 --clipboard

# Raw LaTeX body
palmer-type.exe --raw "\Palmer{123}{123}{}{}{}{}" -o upper_ant.png

# Text color — named xcolor color
palmer-type.exe --UL 12345678 --UR 12345678 --color red -o full_red.png

# Text color — 6-digit hex
palmer-type.exe --UL 12345678 --UR 12345678 --color "#1A73E8" -o full_blue.png

# Transparent background (PNG and clipboard)
palmer-type.exe --UL 1 --transparent -o UL1_transparent.png
```

If neither `-o` nor `--clipboard` is given, the PNG is written to stdout (useful for piping).

### All options

> **Note on L/R in the CLI:** The quadrant flags follow `palmer.sty`'s drawing convention —
> `UL`/`LL` refer to the **upper/lower-left** of the rendered cross (i.e. the patient's right jaw),
> and `UR`/`LR` refer to the **upper/lower-right** (i.e. the patient's left jaw).
> This is the reverse of everyday anatomical usage, where "left" means the patient's left.
> The GUI displays labels in the anatomical convention (UR/LR on the left side of the screen for patient's right), so be aware of this difference when translating between the two interfaces.

| Flag | Argument | Description |
|---|---|---|
| `--UL` | TEXT | Upper-left quadrant of the cross (patient's **right** upper jaw) |
| `--UR` | TEXT | Upper-right quadrant of the cross (patient's **left** upper jaw) |
| `--LL` | TEXT | Lower-left quadrant of the cross (patient's **right** lower jaw) |
| `--LR` | TEXT | Lower-right quadrant of the cross (patient's **left** lower jaw) |
| `--upper-mid` | TEXT | Upper midline symbol |
| `--lower-mid` | TEXT | Lower midline symbol |
| `-o`, `--output` | PATH | Output file path (PNG / JPEG / PDF) |
| `--clipboard` | — | Copy rendered image to clipboard (Windows) |
| `--batch` | PATH | JSON file for batch processing |
| `--outdir` | PATH | Output directory for batch results (default: `output/`) |
| `--raw` | TEX | Raw TeX body compiled directly (bypasses quadrant inputs) |
| `--dpi` | INT | Output resolution in DPI (default: `600`) |
| `--color` | COLOR | Text color: `#RRGGBB` hex (e.g., `#FF0000`) or named [xcolor](https://ctan.org/pkg/xcolor) color (e.g., `red`, `blue`, `darkgray`). Default: black. |
| `--transparent` | — | Use transparent background (PNG and clipboard). Ignored for JPEG/PDF output. |
| `--font` | NAME | Font family name (default: `Times New Roman`) |
| `--font-size` | FLOAT | Font size in points, 2.0–144.0 (default: `10.0`) |
| `--sty` | PATH | Path to a custom `palmer.sty` |
| `--margin-top` | INT | Top margin in pixels (default: `8`) |
| `--margin-bottom` | INT | Bottom margin in pixels (default: `8`) |
| `--margin-left` | INT | Left margin in pixels (default: `8`) |
| `--margin-right` | INT | Right margin in pixels (default: `8`) |
| `--verbose`, `-v` | — | Enable debug logging on stderr |
| `--version` | — | Show program version and exit |

### Batch processing

Pass a JSON array to `--batch`. Each object may contain:

| Field | Required | Description |
|---|---|---|
| `id` | Yes | Filename stem for the output PNG (e.g. `"UL1"` → `UL1.png`) |
| `label` | No | Human-readable label (logged to stderr) |
| `UL`, `UR`, `LL`, `LR` | — | Quadrant inputs — same spatial convention as the CLI flags above |
| `upper_mid`, `lower_mid` | — | Midline symbols |
| `color` | — | Text color — same format as `--color` above |
| `transparent` | — | `true` for transparent background (PNG and clipboard) |
| `font` | — | Font family name (default: `Times New Roman`) |
| `font_size` | — | Font size in points (default: `10.0`) |
| `raw` | — | Raw TeX body (mutually exclusive with quadrant fields) |

```json
[
  {
    "id": "UL1",
    "label": "Upper-left quadrant, tooth 1 (patient's right central incisor)",
    "UL": "1"
  },
  {
    "id": "full",
    "label": "Full dentition",
    "UL": "12345678",
    "UR": "12345678",
    "LR": "12345678",
    "LL": "12345678"
  },
  {
    "id": "custom",
    "label": "Custom raw command",
    "raw": "\\Palmer{123}{123}{}{}{}{}"
  },
  {
    "id": "colored",
    "label": "Full dentition in red",
    "UL": "12345678",
    "UR": "12345678",
    "LR": "12345678",
    "LL": "12345678",
    "color": "red"
  },
  {
    "id": "transparent",
    "label": "Upper-left tooth 1 with transparent background",
    "UL": "1",
    "transparent": true
  }
]
```

## Cache and data storage

palmer-type stores files in three locations outside the installation folder:

| Data | Location |
|---|---|
| Tectonic binary and TeX cache | `%LOCALAPPDATA%\TectonicProject\Tectonic\` |
| Font favorites | `%APPDATA%\palmer-type\favorites.json` |
| Debug log files | `%APPDATA%\palmer-type\logs\` |

The Tectonic cache and font favorites **persist after uninstalling or deleting `palmer-type.exe`**.
Debug log files are automatically removed after 92 days at the next startup, and are always deleted by the Windows uninstaller.

### Removing cached data

**Via the uninstaller (installer variant):** The Windows uninstaller always removes debug log files, and additionally offers an option to delete the Tectonic cache and font favorites during uninstallation.

**Via the About tab (debug mode):** Enable **Debug** from the menu bar, then open the **About** tab. Buttons are provided there to individually delete the Tectonic cache folder and the font favorites file, and to open the debug log folder in the system file manager.

**Manually:**

```powershell
rmdir /s /q "%LOCALAPPDATA%\TectonicProject\Tectonic"
del "%APPDATA%\palmer-type\favorites.json"
rmdir /s /q "%APPDATA%\palmer-type\logs"
```

## License

This project is available under a dual-licensing model.

- Open-source license: GNU Affero General Public License v3.0 (AGPLv3)
- Alternative license: commercial/proprietary license

You may use this project under AGPLv3 if you comply with its terms.
If you require terms other than AGPLv3 — for example, if you wish to keep modifications or a combined work closed-source — please contact the copyright holder to discuss alternative licensing.

## Acknowledgement

This work was supported by JSPS KAKENHI Grant Number JP25K15395.

## Copyright

Copyright (c) 2026 Yosuke Yamazaki
