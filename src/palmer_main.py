"""palmer_main.py — Palmer Dental Notation Tool unified entry point.

Without arguments: launches the GUI (no console window).
With arguments: runs the CLI, reconnecting stdout/stderr to the parent terminal.

Usage:
    python palmer_main.py               # GUI
    python palmer_main.py --UL 1 -o x.png  # CLI
    pythonw palmer_main.py              # GUI without console
"""

import sys


def _prevent_duplicate_launch_win32() -> None:
    """Prevent duplicate GUI instances on Windows using a named mutex.

    Creates a named mutex scoped to the current user session (``Local\\``
    prefix).  If the mutex already exists (``ERROR_ALREADY_EXISTS``), another
    instance is running; a message box is shown and the process exits.

    The mutex handle is intentionally never closed — it is held for the
    entire process lifetime and released automatically by the OS on exit.
    """
    try:
        import ctypes
        ERROR_ALREADY_EXISTS = 183
        MB_ICONINFORMATION = 0x00000040
        ctypes.windll.kernel32.CreateMutexW(  # type: ignore[attr-defined]
            None, False, "Local\\PalmerToolGUI_SingleInstance"
        )
        if ctypes.windll.kernel32.GetLastError() == ERROR_ALREADY_EXISTS:  # type: ignore[attr-defined]
            ctypes.windll.user32.MessageBoxW(  # type: ignore[attr-defined]
                0,
                "Palmer Tool is already running.\n\nPlease use the existing window.",
                "Palmer Tool",
                MB_ICONINFORMATION,
            )
            sys.exit(0)
    except (AttributeError, OSError):
        # Non-fatal: allow launch if mutex creation fails (e.g. non-Windows).
        pass


def _attach_console_win32() -> None:
    """Reconnect stdout/stderr/stdin to the parent console (Windows only).

    A --noconsole exe has no console subsystem.  AttachConsole(-1) reattaches
    to the parent process's terminal so CLI output appears normally.
    ATTACH_PARENT_PROCESS is represented by the constant -1.
    """
    try:
        import ctypes
        import io
        ATTACH_PARENT_PROCESS = -1
        if ctypes.windll.kernel32.AttachConsole(ATTACH_PARENT_PROCESS):  # type: ignore[attr-defined]
            # The CONOUT$/CONIN$ handles and the replaced sys.stdout/stderr
            # are intentionally never closed — they are needed for the
            # entire process lifetime and will be released on exit.
            conout = open("CONOUT$", "wb", buffering=0)
            conout2 = None
            conin = None
            try:
                conout2 = open("CONOUT$", "wb", buffering=0)
                conin = open("CONIN$", "rb", buffering=0)
                sys.stdout = io.TextIOWrapper(conout, encoding="utf-8")
                sys.stderr = io.TextIOWrapper(conout2, encoding="utf-8")
                sys.stdin  = io.TextIOWrapper(conin,  encoding="utf-8")
            except (AttributeError, OSError):
                conout.close()
                if conout2 is not None:
                    conout2.close()
                if conin is not None:
                    conin.close()
                raise
    except (AttributeError, OSError):
        # Non-fatal: AttachConsole is best-effort for --noconsole exes.
        # Failure means CLI output won't appear, but the program still works.
        pass


def main() -> None:
    if len(sys.argv) > 1:
        if sys.platform == "win32":
            _attach_console_win32()
        from palmer_cli import main as cli_main
        cli_main()
    else:
        if sys.platform == "win32":
            _prevent_duplicate_launch_win32()
        from palmer_type import main as gui_main
        gui_main()


if __name__ == "__main__":
    main()
