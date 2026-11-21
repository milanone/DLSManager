"""Draw the program icon (DLSManager.ico, next to DLSManager.pyw), in the style of the other programs' icons (white
rounded tile, black axes, blue and orange curves): the decaying correlation function of a DLS measurement with its measured
points (orange) and the narrow size distribution it gives (blue). Drawn large and shrunk for each size, so the small sizes
stay clean.

    python examples/make_icon.py
"""
import os
import struct

import numpy as np
from PIL import Image, ImageDraw

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "..", "DLSManager.ico")
SIZES = [16, 24, 32, 48, 64]      # classic sizes: no 128/256 frames, which only exist as PNG
BIG = 1024
BLUE, ORANGE = "#2d6cb5", "#f28c00"


def disc(d, x, y, r, fill):
    d.ellipse([x - r, y - r, x + r, y + r], fill=fill)


def draw():
    img = Image.new("RGBA", (BIG, BIG), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.rounded_rectangle([40, 40, BIG - 40, BIG - 40], radius=190, fill="#ffffff", outline="#1a1a1a", width=40)
    d.line([(190, 190), (190, 834), (834, 834)], fill="#000000", width=52, joint="curve")
    for p in ((190, 190), (834, 834)):
        disc(d, p[0], p[1], 26, "#000000")
    # the size distribution: a narrow peak on the baseline (behind), stamped as round discs
    for x in np.linspace(540, 820, 900):
        y = 800 - 360 * np.exp(-((x - 690) / 58) ** 2)
        disc(d, x, y, 32, BLUE)
    # the correlation function: a plateau that falls off, with its measured points
    def decay(x):
        return 290 + 400 / (1 + np.exp(-(x - 450) / 55))
    for x in np.linspace(245, 740, 1200):
        disc(d, x, decay(x), 34, ORANGE)
    for x in (270, 380, 470, 560):
        disc(d, x, decay(x), 50, "#ffffff")
        disc(d, x, decay(x), 38, "#000000")
    return img


def write_ico(path, images):
    """A classic .ico: every frame a 32-bit BMP (DIB) with alpha and an empty AND mask, smallest first. Pillow writes
    PNG-compressed frames, which Windows Vista and later read but older or simpler programs (a file manager's
    toolbar button, for one) do not."""
    entries, blobs, offset = [], [], 6 + 16 * len(images)
    for im in images:
        w, h = im.size
        px = np.asarray(im.convert("RGBA"))[::-1][..., [2, 1, 0, 3]]          # BGRA, bottom row first
        pixels = np.ascontiguousarray(px).tobytes()
        mask = bytes(((w + 31) // 32) * 4 * h)                                  # all zero: the alpha channel does the work
        header = struct.pack("<IiiHHIIiiII", 40, w, h * 2, 1, 32, 0, len(pixels) + len(mask), 0, 0, 0, 0)
        blob = header + pixels + mask
        entries.append(struct.pack("<BBBBHHII", w % 256, h % 256, 0, 0, 1, 32, len(blob), offset))
        blobs.append(blob)
        offset += len(blob)
    with open(path, "wb") as f:
        f.write(struct.pack("<HHH", 0, 1, len(images)) + b"".join(entries) + b"".join(blobs))


if __name__ == "__main__":
    big = draw()
    write_ico(OUT, [big.resize((s, s), Image.LANCZOS) for s in SIZES])
    big.resize((256, 256), Image.LANCZOS).save(os.path.join(os.environ.get("TEMP", HERE), "dlsmanager_icon_preview.png"))
    print("written", os.path.abspath(OUT), os.path.getsize(OUT), "bytes")
