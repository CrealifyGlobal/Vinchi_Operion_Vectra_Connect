"""
make_assets.py
Generates the minimum installer assets required by WiX:
  - Installer/Assets/banner.bmp   (493x58  — top banner on wizard dialogs)
  - Installer/Assets/dialog.bmp   (493x312 — welcome/finish dialog background)
  - Installer/Assets/publisher.ico
  - Installer/Assets/logo.png     (used by the Burn bootstrapper)

Run this once locally to produce real branded images, or let CI generate
plain-colour placeholders so the build never fails.

To use real branding: replace the output files with your own images and
remove the call to this script from the workflow (the files will be
committed to the repo instead).
"""

import os
import struct
import zlib
import pathlib

ASSETS = pathlib.Path(__file__).parent.parent / "Installer" / "Assets"
ASSETS.mkdir(parents=True, exist_ok=True)


# ── BMP helper ───────────────────────────────────────────────────────────────

def write_bmp(path: pathlib.Path, width: int, height: int,
              bg: tuple = (31, 56, 100),        # default: dark navy
              fg: tuple = (255, 255, 255)):
    """Write a minimal 24-bit BMP with a solid background colour."""
    row_size = (width * 3 + 3) & ~3             # padded to 4-byte boundary
    pixel_data_size = row_size * height
    file_size = 54 + pixel_data_size

    header = struct.pack(
        "<2sIHHI",
        b"BM", file_size, 0, 0, 54              # file header
    )
    dib = struct.pack(
        "<IiiHHIIiiII",
        40, width, -height, 1, 24,              # BITMAPINFOHEADER
        0, pixel_data_size, 2835, 2835, 0, 0
    )

    row = bytes([bg[2], bg[1], bg[0]] * width)  # BGR
    row += b"\x00" * (row_size - width * 3)

    with open(path, "wb") as f:
        f.write(header + dib + row * height)

    print(f"  wrote {path.name}  ({width}x{height})")


# ── ICO helper ───────────────────────────────────────────────────────────────

def write_ico(path: pathlib.Path, colour: tuple = (31, 56, 100)):
    """Write a minimal 32x32 ICO file."""
    size = 32
    bmp_data = bytearray()
    # BITMAPINFOHEADER (40 bytes) for 32x32 icon, 32bpp
    bmp_data += struct.pack("<IiiHHIIiiII",
                            40, size, -(size * 2), 1, 32,
                            0, size * size * 4, 0, 0, 0, 0)
    # XOR mask (BGRA pixels)
    r, g, b = colour
    for _ in range(size * size):
        bmp_data += bytes([b, g, r, 255])
    # AND mask (all transparent = 0)
    bmp_data += bytes(size * size // 8)

    ico_header = struct.pack("<HHH", 0, 1, 1)           # ICONDIR
    ico_entry  = struct.pack("<BBBBHHII",
                             size, size, 0, 0, 1, 32,   # ICONDIRENTRY
                             len(bmp_data), 22)
    with open(path, "wb") as f:
        f.write(ico_header + ico_entry + bytes(bmp_data))
    print(f"  wrote {path.name}  (32x32 ICO)")


# ── PNG helper (no Pillow needed) ────────────────────────────────────────────

def write_png(path: pathlib.Path, width: int, height: int,
              colour: tuple = (31, 56, 100)):
    """Write a minimal solid-colour PNG without external dependencies."""
    r, g, b = colour

    def png_chunk(chunk_type: bytes, data: bytes) -> bytes:
        c = chunk_type + data
        return struct.pack(">I", len(data)) + c + struct.pack(">I", zlib.crc32(c) & 0xFFFFFFFF)

    signature = b"\x89PNG\r\n\x1a\n"
    ihdr_data = struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)
    ihdr = png_chunk(b"IHDR", ihdr_data)

    raw_rows = b""
    row = bytes([0] + [r, g, b] * width)        # filter byte 0 + RGB pixels
    raw_rows = row * height
    idat = png_chunk(b"IDAT", zlib.compress(raw_rows, 9))
    iend = png_chunk(b"IEND", b"")

    with open(path, "wb") as f:
        f.write(signature + ihdr + idat + iend)
    print(f"  wrote {path.name}  ({width}x{height} PNG)")


# ── Main ─────────────────────────────────────────────────────────────────────

def main():
    NAVY   = (31, 56, 100)    # dark navy — matches the Excel header colour
    WHITE  = (255, 255, 255)

    print("Generating installer assets …")

    write_bmp(ASSETS / "banner.bmp",  493,  58, bg=NAVY,  fg=WHITE)
    write_bmp(ASSETS / "dialog.bmp",  493, 312, bg=WHITE, fg=NAVY)
    write_ico(ASSETS / "publisher.ico", colour=NAVY)
    write_png(ASSETS / "logo.png",    120,  40, colour=NAVY)

    print("Done. Replace these placeholders with real branded images before release.")


if __name__ == "__main__":
    main()
