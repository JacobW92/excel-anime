#!/usr/bin/env python3
"""
Excel Anime Player (macOS)
Uses AppleScript for zoom + scroll_row to show exactly one frame at a time.
"""

import sys
import time
import string
import subprocess
from pathlib import Path


def col_letter(n):
    s = ""
    while n > 0:
        n -= 1
        s = string.ascii_uppercase[n % 26] + s
        n //= 26
    return s


def osa(cmd):
    """Run an AppleScript snippet, return result."""
    return subprocess.run(["osascript", "-e", cmd],
                          capture_output=True, text=True, timeout=30)


def set_zoom(pct):
    """Set Excel zoom percentage."""
    osa(f'''
tell application "Microsoft Excel"
    set zoom of active window to {pct}
end tell
''')


def set_scroll_row(row):
    """Scroll so that `row` appears at the top-left of the window."""
    return osa(f'''
tell application "Microsoft Excel"
    goto reference (range "A{row}" of active sheet of active workbook) scroll true
end tell
''')


def play(excel_path, width, height, num_frames, fps):
    try:
        import xlwings as xw
    except ImportError:
        sys.exit("pip install xlwings")

    print("=" * 52)
    print("   🎬 Excel Anime Player")
    print("=" * 52)
    print(f"   File:     {Path(excel_path).name}")
    print(f"   Frame:    {width}×{height}  |  {num_frames} frames  |  {fps} fps")
    print(f"   Duration: ~{num_frames / fps:.0f}s")
    print("=" * 52)

    # ── Close existing Excel ───────────────────────────
    print("\n[*] Closing existing Excel...")
    osa('quit app "Microsoft Excel"')
    time.sleep(2)

    # ── Open Excel + file ──────────────────────────────
    print("[*] Opening Excel...")
    app = xw.App(visible=True)
    app.activate()

    print("[*] Opening file...")
    wb = app.books.open(str(excel_path))

    print("[*] Loading file (please wait)...")
    time.sleep(15)

    # ── Calculate zoom to fit exactly one frame ────────
    # Zoom based on HEIGHT so only `height` rows are visible.
    # Cell at 100%: ~15px (square, col=2.0/row=16.0)
    # Screen 1512×982, Excel chrome ~180px → usable height ~800px
    # zoom = usable_height / (height * cell_px_at_100) * 100
    # Try a range of heights to find the right zoom
    usable_height = 800  # conservative estimate
    cell_px = 15         # approximate px per cell at 100% zoom (square)
    zoom_for_height = int(usable_height / (height * cell_px) * 100)
    zoom_for_width = int(1360 / (width * cell_px) * 100)
    # Use the SMALLER zoom so the frame fits both ways
    zoom_pct = min(zoom_for_height, zoom_for_width)
    zoom_pct = max(10, zoom_pct)

    print(f"[*] Setting zoom to {zoom_pct}% "
          f"(height-fit={zoom_for_height}%, width-fit={zoom_for_width}%)")
    set_zoom(zoom_pct)
    time.sleep(1)

    # Scroll to frame 0 (top)
    set_scroll_row(1)
    time.sleep(1)

    # ── Ready ──────────────────────────────────────────
    print()
    print("=" * 52)
    print("   🎥 START SCREEN RECORDING NOW!")
    print("   Starting in 3 seconds...")
    print("=" * 52)
    time.sleep(3)

    # ── Playback ───────────────────────────────────────
    delay = 1.0 / fps
    start = time.time()
    errors = 0

    for i in range(num_frames):
        row = i * height + 1
        target_t = start + (i + 1) * delay

        # Set scroll position (frame-aligned)
        r = set_scroll_row(row)
        if r.returncode != 0:
            errors += 1
            if errors > 30:
                print(f"\n    ⚠ Too many errors ({errors}), stopping.")
                break

        # Progress
        if (i + 1) % 40 == 0:
            elapsed = time.time() - start
            pct = (i + 1) * 100 // num_frames
            print(f"    {i+1}/{num_frames} ({pct}%)  "
                  f"{elapsed:.0f}s / ~{num_frames/fps:.0f}s  "
                  f"errors: {errors}")

        # Wait (compensate for processing overhead)
        now = time.time()
        wait = target_t - now
        if wait > 0:
            time.sleep(wait)

    elapsed = time.time() - start
    print(f"\n{'=' * 52}")
    print(f"   ✅ Done!  {elapsed:.0f}s  |  {errors} errors")
    print(f"{'=' * 52}")


if __name__ == "__main__":
    excel = sys.argv[1] if len(sys.argv) > 1 else sys.exit("Usage: python play.py <file.xlsx>")
    w = int(sys.argv[2]) if len(sys.argv) > 2 else 200
    h = int(sys.argv[3]) if len(sys.argv) > 3 else 112
    n = int(sys.argv[4]) if len(sys.argv) > 4 else 720
    fps = int(sys.argv[5]) if len(sys.argv) > 5 else 8

    play(excel, w, h, n, fps)
