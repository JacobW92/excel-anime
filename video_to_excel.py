#!/usr/bin/env python3
"""
Excel Anime Player Generator
=============================
Converts a video into an Excel spreadsheet where each cell = one pixel,
and a VBA macro scrolls through frames for animation playback.

Usage:
    python video_to_excel.py input.mp4
    python video_to_excel.py input.mp4 --fps 8 --width 100
    python video_to_excel.py input.mp4 --fps 12 --width 160 --max-frames 60
"""

import argparse
import subprocess
import os
import sys
import tempfile
from pathlib import Path
from time import time

# ── Dependency check ─────────────────────────────────────────────
try:
    from PIL import Image
except ImportError:
    sys.exit("Missing Pillow. Install: pip install Pillow")

try:
    import xlsxwriter
except ImportError:
    sys.exit("Missing XlsxWriter. Install: pip install XlsxWriter")


# ── Frame extraction ─────────────────────────────────────────────
def extract_frames(video_path, output_dir, fps):
    """Use ffmpeg to extract frames at the target framerate."""
    pattern = str(Path(output_dir) / "frame_%05d.png")
    cmd = ["ffmpeg", "-y", "-i", video_path, "-vf", f"fps={fps}", pattern]
    print(f"[*] Extracting frames at {fps} fps ...")
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        print(f"ffmpeg stderr: {result.stderr[:500]}")
        sys.exit(1)

    frames = sorted(Path(output_dir).glob("frame_*.png"))
    print(f"    {len(frames)} frames extracted")
    return [str(f) for f in frames]


# ── Image processing ─────────────────────────────────────────────
def create_stacked_image(frame_paths, width, height):
    """Resize every frame to (width, height) and stack them vertically."""
    n = len(frame_paths)
    stacked = Image.new("RGB", (width, height * n))

    for i, path in enumerate(frame_paths):
        img = Image.open(path).convert("RGB").resize((width, height), Image.LANCZOS)
        stacked.paste(img, (0, i * height))
        if (i + 1) % 50 == 0 or i == n - 1:
            print(f"    Resized {i+1}/{n} frames")

    return stacked


def quantize_image(img, levels_per_channel=32):
    """Reduce colour depth to keep unique colours within Excel's ~64K format limit.

    levels_per_channel: number of quantised levels per R/G/B channel.
        32 levels → 32,768 max colours  (safe, good quality)
        16 levels →  4,096 max colours  (very safe, slight banding)
    """
    if levels_per_channel >= 256:
        return img  # no quantisation needed

    step = 256 // levels_per_channel
    pixels = img.load()
    w, h = img.size

    for y in range(h):
        for x in range(w):
            r, g, b = pixels[x, y]
            # snap each channel to nearest quantised level
            r = round(r / step) * step
            g = round(g / step) * step
            b = round(b / step) * step
            pixels[x, y] = (min(r, 255), min(g, 255), min(b, 255))

    return img


# ── Excel generation ─────────────────────────────────────────────
def generate_excel(stacked_img, width, height, num_frames, fps, output_path,
                   levels_per_channel=32):
    """Write an .xlsx where every cell background = one pixel colour."""
    total_cells = width * height * num_frames
    print(f"[*] Writing Excel  ({total_cells:,} cells, "
          f"colour depth {levels_per_channel} levels/ch) ...")

    t0 = time()
    wb = xlsxwriter.Workbook(output_path)
    ws = wb.add_worksheet("Anime")

    # Cell sizing → approximate squares
    # Excel: 1 char-width ≈ 7 px ;  1 point-height ≈ 1.33 px
    COL_W = 2.0      # square cells (verified on macOS Excel)
    ROW_H = 16.0     # square cells (verified on macOS Excel)

    ws.set_column(0, width - 1, COL_W)

    # Pre-quantise the image so we stay within Excel's format limit
    stacked_img = quantize_image(stacked_img, levels_per_channel)

    fmt_cache = {}
    pixels = stacked_img.load()

    for fi in range(num_frames):
        for y in range(height):
            row = fi * height + y
            ws.set_row(row, ROW_H)
            for x in range(width):
                r, g, b = pixels[x, fi * height + y]
                hx = f"{r:02X}{g:02X}{b:02X}"
                if hx not in fmt_cache:
                    fmt_cache[hx] = wb.add_format({"bg_color": f"#{hx}"})
                ws.write_blank(row, x, None, fmt_cache[hx])

        if (fi + 1) % 5 == 0:
            elapsed = time() - t0
            rate = (fi + 1) / elapsed
            eta = (num_frames - fi - 1) / rate if rate else 0
            pct = (fi + 1) * 100 // num_frames
            print(f"    {fi+1}/{num_frames} frames ({pct}%)  "
                  f"{elapsed:.0f}s elapsed  ~{eta:.0f}s left  "
                  f"{len(fmt_cache):,} colours")

    wb.close()
    elapsed = time() - t0
    size_mb = os.path.getsize(output_path) / 1048576
    print(f"[✓] Saved {output_path}  ({size_mb:.1f} MB, {elapsed:.0f}s)")
    print(f"    Unique colours in file: {len(fmt_cache):,}  "
          f"(Excel limit ≈ 64,000)")


# ── VBA macro ────────────────────────────────────────────────────
def generate_vba(width, height, num_frames, fps, out_dir):
    """Write a .bas file with a PlayAnimation macro."""
    delay_ms = int(1000 / fps)
    path = str(Path(out_dir) / "anime_player_macro.bas")

    delay_s = round(1.0 / fps, 4)
    code = f'''\
Attribute VB_Name = "ModuleAnime"
' ====================================================
'  Excel Anime Player  –  auto-generated macro
'  {width} x {height}  |  {num_frames} frames  |  {fps} fps
' ====================================================

#If Win64 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongLong)
#ElseIf Win32 Then
    Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)
#End If

Sub PlayAnimation()
    Dim frameH  As Long:   frameH = {height}
    Dim nFrames As Long:   nFrames = {num_frames}
    Dim i       As Long
    Dim t0      As Single

    ActiveWindow.ScrollRow = 1
    DoEvents

    For i = 0 To nFrames - 1
        ActiveWindow.ScrollRow = i * frameH + 1
        DoEvents

        #If Mac Then
            t0 = Timer
            Do While (Timer - t0) < {delay_s}
                DoEvents
            Loop
        #Else
            Sleep {delay_ms}
        #End If
    Next i

    MsgBox "Done!", vbInformation
End Sub

Sub ResetView()
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
End Sub
'''
    with open(path, "w", encoding="utf-8") as f:
        f.write(code)
    print(f"[✓] VBA macro  →  {path}")
    return path


# ── CLI ──────────────────────────────────────────────────────────
def main():
    ap = argparse.ArgumentParser(
        description="Convert video to Excel pixel animation",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    ap.add_argument("video", help="Input video (mp4 / mkv / avi …)")
    ap.add_argument("--fps", type=int, default=8, help="Frame rate (default 8)")
    ap.add_argument("--width", type=int, default=100, help="Pixel width (default 100)")
    ap.add_argument("--height", type=int, default=0, help="Pixel height (0 = auto 16:9)")
    ap.add_argument("-o", "--output", default="", help="Output .xlsx path")
    ap.add_argument("--max-frames", type=int, default=0, help="Cap frame count (0 = all)")
    ap.add_argument("--save-stacked", action="store_true", help="Also save stacked PNG")
    ap.add_argument("--colors", type=int, default=32,
                    help="Colour levels per channel (default 32 → max 32,768 colours). "
                         "Use 16 for very safe (4,096 colours). Excel limit ≈ 64,000 unique formats.")
    args = ap.parse_args()

    # ── checks ───────────────────────
    if not Path(args.video).is_file():
        sys.exit(f"File not found: {args.video}")
    try:
        subprocess.run(["ffmpeg", "-version"], capture_output=True, check=True)
    except Exception:
        sys.exit("ffmpeg not found.  macOS: brew install ffmpeg")

    if args.height == 0:
        args.height = round(args.width * 9 / 16)

    if not args.output:
        args.output = f"{Path(args.video).stem}_excel.xlsx"

    out_dir = str(Path(args.output).parent) or "."

    print("=" * 52)
    print("   Excel Anime Player Generator")
    print("=" * 52)
    print(f"   Video      {args.video}")
    print(f"   FPS        {args.fps}")
    print(f"   Resolution {args.width} x {args.height}")
    print(f"   Output     {args.output}")
    print("=" * 52)

    # ── pipeline ─────────────────────
    with tempfile.TemporaryDirectory() as tmp:
        frames = extract_frames(args.video, tmp, args.fps)
        if not frames:
            sys.exit("No frames extracted – check the video file.")
        if 0 < args.max_frames < len(frames):
            frames = frames[: args.max_frames]
            print(f"    Capped at {args.max_frames} frames")

        stacked = create_stacked_image(frames, args.width, args.height)

        if args.save_stacked:
            p = args.output.replace(".xlsx", "_stacked.png")
            stacked.save(p)
            print(f"[✓] Stacked PNG → {p}")

        generate_excel(stacked, args.width, args.height, len(frames), args.fps,
                       args.output, levels_per_channel=args.colors)

    vba_path = generate_vba(args.width, args.height, len(frames), args.fps, out_dir)

    # ── instructions ─────────────────
    print()
    print("=" * 52)
    print("   How to play")
    print("=" * 52)
    print(f"   1. Open  {args.output}  in Excel")
    print(f"   2. 另存为 .xlsm 格式（启用宏）")
    print(f"   3. Alt+F11 → 文件 → 导入文件 → {vba_path}")
    print(f"   4. F5 运行 PlayAnimation")
    print(f"   5. 录屏！")
    print("=" * 52)


if __name__ == "__main__":
    main()
