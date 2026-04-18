"""
build.py — one-command build script for BrightspaceConverter.exe

Usage:
    python build.py           # standard build
    python build.py --fast    # skip cleaning, rebuild only changed files
    python build.py --debug   # build with console window for troubleshooting
"""

import subprocess
import shutil
import sys
import time
from pathlib import Path


# ── Configuration ────────────────────────────────────────────────────────────

APP_NAME    = "BrightspaceConverter"
ENTRY_POINT = "converter.py"
ICON        = None          # Set to e.g. "assets/icon.ico" if you have one

# Extra PyInstaller flags you want baked in permanently
EXTRA_FLAGS = [
    "--collect-all", "tkinterdnd2",   # ensures drag-and-drop works in the exe
]

# ─────────────────────────────────────────────────────────────────────────────


def clean():
    """Remove previous build artefacts so we start fresh."""
    for folder in ("build", "dist"):
        p = Path(folder)
        if p.exists():
            print(f"  Removing {folder}/")
            shutil.rmtree(p)
    for spec in Path(".").glob("*.spec"):
        print(f"  Removing {spec.name}")
        spec.unlink()


def build(fast=False, debug=False):
    if not fast:
        print("\n── Cleaning previous build ──────────────────────────────")
        clean()

    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",
        "--name", APP_NAME,
        *EXTRA_FLAGS,
    ]

    if debug:
        print("\n  [debug mode] Console window will be visible.")
    else:
        cmd.append("--windowed")

    if ICON and Path(ICON).exists():
        cmd += ["--icon", ICON]
    elif ICON:
        print(f"  Warning: icon not found at '{ICON}', skipping.")

    cmd.append(ENTRY_POINT)

    print(f"\n── Building {APP_NAME}.exe ───────────────────────────────")
    print("  " + " ".join(cmd))
    print()

    start = time.time()
    result = subprocess.run(cmd)
    elapsed = time.time() - start

    if result.returncode == 0:
        exe = Path("dist") / f"{APP_NAME}.exe"
        size_mb = exe.stat().st_size / 1_048_576 if exe.exists() else 0
        print(f"\n✅  Build succeeded in {elapsed:.1f}s")
        print(f"    Output : {exe.resolve()}")
        print(f"    Size   : {size_mb:.1f} MB")
    else:
        print(f"\n❌  Build failed (exit code {result.returncode})")
        print("    Tip: run with --debug to see the console output from the exe.")
        sys.exit(result.returncode)


if __name__ == "__main__":
    args = sys.argv[1:]
    build(
        fast="--fast"   in args,
        debug="--debug" in args,
    )
