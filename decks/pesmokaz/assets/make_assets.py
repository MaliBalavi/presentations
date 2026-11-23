#!/usr/bin/env python3
"""Generate decorative assets for the Pesmokaz deck (spectrogram + waveform).

These are stylised brand visuals, not measurements from the dataset.
"""
import math
import random
from PIL import Image, ImageDraw, ImageFilter

OUT = __file__.rsplit("/", 1)[0]
TEAL = (0, 153, 153)
INK = (14, 17, 24)


def lerp(a, b, t):
    return tuple(int(a[i] + (b[i] - a[i]) * t) for i in range(3))


def spectrogram(path, w=1600, h=520, seed=7, alpha=110):
    random.seed(seed)
    sw, sh = 200, 64
    small = Image.new("L", (sw, sh))
    px = small.load()
    for y in range(sh):
        for x in range(sw):
            # energy decays with frequency, plus random peaks
            base = (1 - y / sh) ** 1.6
            v = base * random.uniform(0.25, 1.0)
            if random.random() < 0.03:
                v = min(1.0, v + random.uniform(0.4, 0.9))
            px[x, y] = int(255 * v)
    small = small.filter(ImageFilter.GaussianBlur(1.1))
    img = small.resize((w, h), Image.BICUBIC)
    img = img.filter(ImageFilter.GaussianBlur(2))
    # colour map: ink -> teal -> amber
    grad = [(0.0, INK), (0.55, TEAL), (0.82, (255, 176, 32)), (1.0, (255, 236, 200))]
    lut = []
    for i in range(256):
        t = i / 255
        for j in range(len(grad) - 1):
            t0, c0 = grad[j]
            t1, c1 = grad[j + 1]
            if t0 <= t <= t1:
                lut.append(lerp(c0, c1, (t - t0) / (t1 - t0)))
                break
        else:
            lut.append(grad[-1][1])
    out = Image.new("RGBA", (w, h))
    src = img.load()
    dst = out.load()
    for y in range(h):
        for x in range(w):
            c = lut[src[x, y]]
            # fade to the right and to the bottom
            a = int(alpha * (1 - 0.55 * x / w) * (1 - 0.25 * y / h))
            dst[x, y] = (c[0], c[1], c[2], a)
    out.save(path)


def waveform(path, w=1800, h=320, seed=3, color=TEAL, alpha=235):
    random.seed(seed)
    img = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    mid = h / 2
    pts = []
    ph = [random.uniform(0, 6.28) for _ in range(7)]
    for x in range(0, w, 3):
        t = x / w
        env = math.sin(math.pi * t) ** 1.3
        v = sum(math.sin(2 * math.pi * (3 + i * 2.7) * t + ph[i]) / (i + 1.4) for i in range(7))
        y = mid + env * v * (h * 0.36)
        pts.append((x, y))
    d.line(pts, fill=(color[0], color[1], color[2], alpha), width=3, joint="curve")
    # faint mirrored envelope
    mirror = [(x, h - y) for x, y in pts]
    d.line(mirror, fill=(color[0], color[1], color[2], 55), width=2, joint="curve")
    img.save(path)


if __name__ == "__main__":
    spectrogram(f"{OUT}/spectrogram.png")
    waveform(f"{OUT}/waveform.png")
    print("assets written")
