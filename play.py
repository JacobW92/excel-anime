#!/usr/bin/env python3
"""
Excel Anime Player (macOS)
Plays one or more Excel animation files seamlessly.
Uses AppleScript for zoom + Goto Scroll:=True.
"""

import sys
import time
import string
import subprocess
from pathlib import Path


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
    return osa(f'''
tell application "Microsoft Excel"
    goto reference (range "A{row}" of active sheet of active workbook) scroll true
end tell
''')


def calc_zoom(width, height):
    """Calculate zoom to fit one frame in the window."""
    cell_px = 15
    zoom_h = int(800 / (height * cell_px) * 100)
    zoom_w = int(1360 / (width * cell_px) * 100)
    return max(10, min(zoom_h, zoom_w))


def count_frames(app, excel_path, height):
    """Detect number of frames from file."""
    wb = app.books.open(str(excel_path))
    last_row = wb.sheets[0].range("A1").end("down").row
    wb.close()
    return last_row // height


def play_file(app, excel_path, width, height, num_frames, fps,
              is_first=False):
    """Play one Excel file. Returns (errors, elapsed)."""
    print(f"\n{'─' * 52}")
    print(f"  📂 {Path(excel_path).name}  ({num_frames} frames, ~{num_frames/fps:.0f}s)")
    print(f"{'─' * 52}")

    wb = app.books.open(str(excel_path))

    if is_first:
        print("[*] Loading first file (please wait)...")
        time.sleep(15)
    else:
        print("[*] Switching to next part...")
        time.sleep(5)

    zoom_pct = calc_zoom(width, height)
    set_zoom(zoom_pct)
    time.sleep(1)
    goto_row(1)
    time.sleep(1)

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
                print(f"    ⚠ Too many errors, skipping.")
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

    # Close workbook after playing
    try:
        wb.close()
    except Exception:
        pass

    return errors, elapsed


def main():
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

    print("=" * 52)
    print("   🎬 Excel Anime Player")
    print("=" * 52)
    print(f"   Files:    {len(files)}")
    print(f"   Frame:    {width}×{height}  |  {fps} fps")

    # Count frames per file
    print("\n[*] Closing existing Excel...")
    osa('quit app "Microsoft Excel"')
    time.sleep(2)

    print("[*] Starting Excel...")
    app = xw.App(visible=True)
    app.activate()

    frame_counts = []
    for f in files:
        n = count_frames(app, f, height)
        frame_counts.append(n)
        total_est = sum(frame_counts) / fps
    print(f"   Total:    {sum(frame_counts)} frames, ~{total_est:.0f}s")
    print("=" * 52)

    print()
    print("=" * 52)
    print("   🎥 START SCREEN RECORDING NOW!")
    print("   Starting in 3 seconds...")
    print("=" * 52)
    time.sleep(3)

    total_start = time.time()
    total_errors = 0

    for idx, (f, n) in enumerate(zip(files, frame_counts)):
        errors, elapsed = play_file(
            app, f, width, height, n, fps,
            is_first=(idx == 0)
        )
        total_errors += errors

    total_elapsed = time.time() - total_start
    print(f"\n{'=' * 52}")
    print(f"   ✅ All done!  {total_elapsed:.0f}s  |  {total_errors} errors")
    print(f"{'=' * 52}")


if __name__ == "__main__":
    main()
