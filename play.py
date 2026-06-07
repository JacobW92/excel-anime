#!/usr/bin/env python3
"""
Excel Anime Player (macOS)
Uses AppleScript for zoom + Goto Scroll:=True to show exactly one frame.
Supports multiple files for seamless sequential playback.
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
    return subprocess.run(["osascript", "-e", cmd],
                          capture_output=True, text=True, timeout=30)


def set_zoom(pct):
    osa(f'''
tell application "Microsoft Excel"
    set zoom of active window to {pct}
end tell
''')


def goto_row(row):
    """Scroll so that `row` appears at the top-left (Goto Scroll:=True)."""
    return osa(f'''
tell application "Microsoft Excel"
    goto reference (range "A{row}" of active sheet of active workbook) scroll true
end tell
''')


def play_part(app, excel_path, width, height, num_frames, fps, first_part=False):
    """Play one Excel file. Returns (errors, elapsed)."""
    import xlwings as xw

    print(f"\n{'─' * 52}")
    print(f"  📂 {Path(excel_path).name}  ({num_frames} frames, ~{num_frames/fps:.0f}s)")
    print(f"{'─' * 52}")

    if first_part:
        print("[*] Opening file...")
        wb = app.books.open(str(excel_path))
        print("[*] Loading (please wait)...")
        time.sleep(15)
    else:
        # Close previous workbook, open next
        try:
            for b in app.books:
                b.close()
        except Exception:
            pass
        print("[*] Opening next part...")
        wb = app.books.open(str(excel_path))
        print("[*] Loading...")
        time.sleep(10)

    # Zoom to fit
    cell_px = 15
    zoom_h = int(800 / (height * cell_px) * 100)
    zoom_w = int(1360 / (width * cell_px) * 100)
    zoom_pct = max(10, min(zoom_h, zoom_w))
    set_zoom(zoom_pct)
    time.sleep(1)

    # Go to top
    goto_row(1)
    time.sleep(1)

    # Playback
    delay = 1.0 / fps
    start = time.time()
    errors = 0

    for i in range(num_frames):
        row = i * height + 1
        target_t = start + (i + 1) * delay

        r = goto_row(row)
        if r.returncode != 0:
            errors += 1
            if errors > 30:
                print(f"    ⚠ Too many errors ({errors}), skipping to next part.")
                break

        if (i + 1) % 40 == 0:
            elapsed = time.time() - start
            pct = (i + 1) * 100 // num_frames
            print(f"    {i+1}/{num_frames} ({pct}%)  {elapsed:.0f}s  errors: {errors}")

        now = time.time()
        wait = target_t - now
        if wait > 0:
            time.sleep(wait)

    elapsed = time.time() - start
    print(f"  ✅ Part done: {elapsed:.0f}s, {errors} errors")
    return errors, elapsed


def main():
    # Parse args: play.py file1.xlsx [file2.xlsx ...] [--width W] [--height H] [--fps F]
    files = []
    width, height, fps = 200, 112, 8
    i = 1
    while i < len(sys.argv):
        a = sys.argv[i]
        if a == "--width" and i + 1 < len(sys.argv):
            width = int(sys.argv[i + 1]); i += 2
        elif a == "--height" and i + 1 < len(sys.argv):
            height = int(sys.argv[i + 1]); i += 2
        elif a == "--fps" and i + 1 < len(sys.argv):
            fps = int(sys.argv[i + 1]); i += 2
        elif a.endswith(".xlsx") or a.endswith(".xlsm"):
            files.append(a); i += 1
        else:
            i += 1

    if not files:
        sys.exit("Usage: python play.py file1.xlsx [file2.xlsx ...] [--width 200] [--height 112] [--fps 8]")

    try:
        import xlwings as xw
    except ImportError:
        sys.exit("pip install xlwings")

    # Calculate total frames from each file's row count
    # For auto-detection, we use height to estimate frames
    print("=" * 52)
    print("   🎬 Excel Anime Player")
    print("=" * 52)
    print(f"   Files:    {len(files)}")
    print(f"   Frame:    {width}×{height}  |  {fps} fps")
    print("=" * 52)

    # Close existing Excel
    print("\n[*] Closing existing Excel...")
    osa('quit app "Microsoft Excel"')
    time.sleep(2)

    print("[*] Starting Excel...")
    app = xw.App(visible=True)
    app.activate()

    # Ready
    print()
    print("=" * 52)
    print("   🎥 START SCREEN RECORDING NOW!")
    print("   Starting in 3 seconds...")
    print("=" * 52)
    time.sleep(3)

    total_start = time.time()
    total_errors = 0

    for idx, f in enumerate(files):
        # Auto-detect frame count from file
        wb_temp = app.books.open(str(f))
        last_row = wb_temp.sheets[0].range("A1").end("down").row
        wb_temp.close()
        num_frames = last_row // height
        print(f"  Detected {num_frames} frames in {Path(f).name}")

        errors, elapsed = play_part(
            app, f, width, height, num_frames, fps,
            first_part=(idx == 0)
        )
        total_errors += errors

    total_elapsed = time.time() - total_start
    print(f"\n{'=' * 52}")
    print(f"   ✅ All done!  {total_elapsed:.0f}s  |  {total_errors} total errors")
    print(f"{'=' * 52}")


if __name__ == "__main__":
    main()
