"""Offer to install missing dependencies, then restart.

Deliberately an OFFER, not a silent install. Installing software without asking
is a real boundary even on your own machine, it can breach policy on a managed
PC, and it can fail behind a proxy - so the manual path has to exist anyway.
This just removes the typing.

A restart is unavoidable: pywin32 puts its DLL directory on the search path via
a startup hook, so a process that has already failed to import pythoncom cannot
pick it up afterwards.

The pure helpers here (command building, message text) are unit-tested; the Tk
dialog is a thin shell around them.
"""
import logging
import subprocess
import sys
from pathlib import Path

logger = logging.getLogger(__name__)

# Never let an install run forever behind a hanging proxy.
INSTALL_TIMEOUT_SECONDS = 600


def build_pip_command(packages=None, requirements_path=None):
    """The pip command to run, always against the RUNNING interpreter.

    Using sys.executable matters: `py` and `python` can be different Pythons,
    and installing into the wrong one is exactly the confusion this is meant to
    end.
    """
    command = [sys.executable, '-m', 'pip', 'install']
    if requirements_path:
        command += ['-r', str(requirements_path)]
    if packages:
        command += list(packages)
    return command


def install(packages=None, requirements_path=None, log=None):
    """Run pip. Returns (succeeded, combined_output)."""
    command = build_pip_command(packages, requirements_path)
    if log:
        log(f"Running: {' '.join(command)}\n")
    try:
        completed = subprocess.run(
            command, capture_output=True, text=True,
            timeout=INSTALL_TIMEOUT_SECONDS,
        )
    except subprocess.TimeoutExpired:
        return False, (f"pip did not finish within "
                       f"{INSTALL_TIMEOUT_SECONDS} seconds.")
    except Exception as exc:                       # pip missing, blocked, ...
        return False, f"Could not run pip: {exc}"

    output = (completed.stdout or '') + (completed.stderr or '')
    if log:
        log(output)
    return completed.returncode == 0, output


def restart_application():
    """Relaunch this app in a fresh process and end the current one.

    Popen + exit rather than os.execv: more predictable on Windows, and it lets
    the current process shut down its Tk cleanly.
    """
    try:
        subprocess.Popen([sys.executable] + sys.argv)
        return True
    except Exception as exc:
        logger.error("Could not restart automatically: %s", exc)
        return False


def describe_missing(results):
    """Human-readable summary of what is missing and why it matters."""
    lines = []
    for result in results:
        req = result.requirement
        lines.append(f"• {req.package} - {req.purpose}")
        if req.consequence:
            lines.append(f"    {req.consequence}")
    return "\n".join(lines)


def offer_install(results, requirements_path=None, parent=None):
    """Show the offer and, if accepted, install. Returns True if installed.

    Falls back to returning False (so the caller shows its normal message) if Tk
    is unavailable for any reason - the app must never fail *because* of the
    helper meant to fix it.
    """
    try:
        import tkinter as tk
        from tkinter import messagebox, scrolledtext
    except Exception:
        return False

    packages = [r.requirement.package for r in results]
    blocking = any(r.essential for r in results)

    root = None
    try:
        root = tk.Tk() if parent is None else None
        if root is not None:
            root.withdraw()

        prompt = (
            "The app needs these packages:\n\n"
            f"{describe_missing(results)}\n\n"
            f"Install them now for:\n{sys.executable}\n\n"
            "The app will restart afterwards."
        )
        if not messagebox.askyesno("Install missing packages?", prompt,
                                   parent=parent or root):
            return False

        # A visible window while pip runs, so it never looks frozen.
        window = tk.Toplevel(parent or root)
        window.title("Installing…")
        window.geometry("760x380")
        text = scrolledtext.ScrolledText(window, font=('Consolas', 9), wrap='word')
        text.pack(fill='both', expand=True, padx=10, pady=10)

        def log(message):
            text.insert('end', message)
            text.see('end')
            text.update_idletasks()

        log(f"Installing into: {sys.executable}\n\n")
        ok, output = install(packages=None if requirements_path else packages,
                             requirements_path=requirements_path, log=log)

        if ok:
            log("\n\n✅ Installed successfully. Restarting the app…")
            window.update_idletasks()
            return True

        log("\n\n❌ Installation failed.")
        manual = ' '.join(build_pip_command(packages, requirements_path))
        messagebox.showerror(
            "Installation failed",
            "The packages could not be installed automatically.\n\n"
            f"Please run this yourself:\n\n{manual}\n\n"
            f"{'The app cannot start without them.' if blocking else ''}",
            parent=window)
        return False
    except Exception as exc:
        logger.error("Install offer failed: %s", exc, exc_info=True)
        return False
    finally:
        if root is not None:
            try:
                root.destroy()
            except Exception:
                pass
