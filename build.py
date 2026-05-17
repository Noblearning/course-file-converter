"""
build.py — one-command build script for CourseFileConverter.exe

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

# Resolve all paths relative to this script's own directory so the build
# works correctly regardless of which folder you invoke it from.
HERE = Path(__file__).resolve().parent

APP_NAME    = "CourseFileConverter"
ENTRY_POINT = str(HERE / "course_file_converter.py")
ICON        = None          # Set to e.g. "assets/icon.ico" if you have one

# Extra PyInstaller flags you want baked in permanently
EXTRA_FLAGS = [
    "--collect-all", "tkinterdnd2",   # ensures drag-and-drop works in the exe
]

# ─────────────────────────────────────────────────────────────────────────────


def _force_remove(func, path, exc_info):
    """Error handler for shutil.rmtree — clears read-only flag and retries.

    OneDrive and Windows sometimes mark build artefacts read-only or lock
    them briefly during sync.  Stripping the read-only bit and retrying is
    enough to proceed in nearly all cases.
    """
    import stat, os
    try:
        os.chmod(path, stat.S_IWRITE)
        func(path)
    except Exception as e:
        print(f"  Warning: could not remove {path}: {e}")


def clean():
    """Remove previous build artefacts so we start fresh."""
    for folder in ("build",):
        p = HERE / folder
        if p.exists():
            print(f"  Removing {folder}/")
            shutil.rmtree(p, onexc=_force_remove)
    # Remove previous exe if present
    prev_exe = HERE / f"{APP_NAME}.exe"
    if prev_exe.exists():
        print(f"  Removing {prev_exe.name}")
        prev_exe.unlink(missing_ok=True)
    for spec in HERE.glob("*.spec"):
        print(f"  Removing {spec.name}")
        spec.unlink(missing_ok=True)


def build(fast=False, debug=False):
    if not fast:
        print("\n── Cleaning previous build ──────────────────────────────")
        clean()

    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",
        "--name", APP_NAME,
        "--distpath", str(HERE),
        "--workpath", str(HERE / "build"),
        "--specpath", str(HERE),
        *EXTRA_FLAGS,
    ]

    if debug:
        print("\n  [debug mode] Console window will be visible.")
    else:
        cmd.append("--windowed")

    icon_path = HERE / ICON if ICON else None
    if icon_path and icon_path.exists():
        cmd += ["--icon", str(icon_path)]
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
        exe = HERE / f"{APP_NAME}.exe"
        size_mb = exe.stat().st_size / 1_048_576 if exe.exists() else 0
        print(f"\n✅  Build succeeded in {elapsed:.1f}s")
        print(f"    Output : {exe.resolve()}")
        print(f"    Size   : {size_mb:.1f} MB")

        # The .spec file is only needed to drive PyInstaller during the build.
        # It is not required for the compiled .exe to run, so clean it up.
        for spec in HERE.glob("*.spec"):
            spec.unlink()
            print(f"    Removed: {spec.name} (not needed at runtime)")
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
