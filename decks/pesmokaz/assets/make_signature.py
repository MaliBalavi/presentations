#!/usr/bin/env python3
"""Build the deck's signature visuals from one synthetic 6-second clip.

The clip is not from the dataset: it is a chord bed + a vibrato lead + kick and
hat transients, synthesized here so that the spectrogram has real structure
(harmonic stacks, onsets, a noise floor) instead of being a painted gradient.
The deck labels it as a synthetic example. Outputs:

  spectrogram.png        light treatment, for content slides
  spectrogram-dark.png   dark treatment, for the cover and the close
  waveform.png           teal envelope line on transparent, for the problem slide

Pure Python FFT (no numpy on this machine). Run: python3 assets/make_signature.py
"""
import cmath
import math
import os
import random

from PIL import Image, ImageDraw

OUT = os.path.dirname(os.path.abspath(__file__))
FS = 8000
DUR = 6.0
N = 1024
HOP = 64

PAPER = (244, 245, 242)
INK = (12, 15, 16)
TEAL = (9, 153, 153)
TEAL_WASH = (176, 219, 216)
TEAL_DEEP = (6, 90, 90)
MINT = (206, 243, 240)


# ---------------------------------------------------------------- synthesis
def synth():
    n = int(FS * DUR)
    buf = [0.0] * n
    rnd = random.Random(11)

    chords = [
        (0.00, [261.63, 329.63, 392.00]),  # C
        (1.50, [220.00, 261.63, 329.63]),  # Am
        (3.00, [174.61, 220.00, 261.63]),  # F
        (4.50, [196.00, 246.94, 293.66]),  # G
    ]
    for start, notes in chords:
        for k in range(int(1.5 * FS)):
            t = k / FS
            release = 1.0 if t < 1.30 else max(0.0, (1.5 - t) / 0.20)
            env = min(1.0, t / 0.02) * math.exp(-0.55 * t) * release
            v = 0.0
            for i, f0 in enumerate(notes):
                for h in range(1, 6):
                    v += (0.55 / (h ** 1.25)) * math.sin(2 * math.pi * f0 * h * t + i)
            idx = int(start * FS) + k
            if idx < n:
                buf[idx] += 0.24 * env * v

    # lead line with vibrato
    melody = [(0.0, 523.25), (0.75, 493.88), (1.5, 440.0), (2.25, 392.0),
              (3.0, 349.23), (3.75, 392.0), (4.5, 440.0), (5.25, 523.25)]
    for start, f0 in melody:
        for k in range(int(0.72 * FS)):
            t = k / FS
            release = 1.0 if t < 0.62 else max(0.0, (0.72 - t) / 0.10)
            env = min(1.0, t / 0.03) * math.exp(-1.8 * t) * release
            f = f0 * (1 + 0.003 * math.sin(2 * math.pi * 5.2 * t))
            idx = int(start * FS) + k
            if idx < n:
                buf[idx] += 0.20 * env * (
                    math.sin(2 * math.pi * f * t)
                    + 0.35 * math.sin(4 * math.pi * f * t)
                    + 0.12 * math.sin(6 * math.pi * f * t)
                )

    # kick every 0.5 s
    for beat in range(12):
        for k in range(int(0.18 * FS)):
            t = k / FS
            f = 52 + 90 * math.exp(-28 * t)
            idx = int(beat * 0.5 * FS) + k
            if idx < n:
                buf[idx] += 0.30 * math.exp(-9 * t) * math.sin(2 * math.pi * f * t)

    # hats every 0.25 s: short bright noise
    for beat in range(24):
        for k in range(int(0.035 * FS)):
            t = k / FS
            idx = int(beat * 0.25 * FS) + k
            if idx < n:
                buf[idx] += 0.022 * math.exp(-70 * t) * rnd.uniform(-1, 1)

    # noise floor
    for i in range(n):
        buf[i] += 0.0022 * rnd.uniform(-1, 1)

    peak = max(abs(v) for v in buf) or 1.0
    return [v / peak for v in buf]


def fft(a):
    n = len(a)
    j = 0
    for i in range(1, n):
        bit = n >> 1
        while j & bit:
            j ^= bit
            bit >>= 1
        j |= bit
        if i < j:
            a[i], a[j] = a[j], a[i]
    length = 2
    while length <= n:
        ang = -2 * math.pi / length
        wl = complex(math.cos(ang), math.sin(ang))
        for i in range(0, n, length):
            w = 1 + 0j
            half = length // 2
            for k in range(half):
                u = a[i + k]
                v = a[i + k + half] * w
                a[i + k] = u + v
                a[i + k + half] = u - v
                w *= wl
        length <<= 1
    return a


BANDS = 96
F_LO, F_HI = 50.0, 2400.0
EDGES = [F_LO * (F_HI / F_LO) ** (i / BANDS) for i in range(BANDS + 1)]


def stft(sig):
    """Log-frequency spectrogram: max magnitude per log band, per frame."""
    window = [0.5 - 0.5 * math.cos(2 * math.pi * i / (N - 1)) for i in range(N)]
    frames = (len(sig) - N) // HOP + 1
    cols = []
    for f in range(frames):
        seg = [sig[f * HOP + i] * window[i] for i in range(N)]
        spec = [abs(c) for c in fft(seg)[: N // 2]]
        band = []
        for b in range(BANDS):
            lo, hi = EDGES[b], EDGES[b + 1]
            k0 = max(1, int(lo / (FS / N)))
            k1 = max(k0 + 1, int(hi / (FS / N)) + 1)
            band.append(max(spec[k0:k1]))
        cols.append(band)
    return cols  # frames x BANDS, low frequency first


# ---------------------------------------------------------------- colour
def ramp(stops):
    def f(t):
        t = min(1.0, max(0.0, t))
        for i in range(len(stops) - 1):
            t0, c0 = stops[i]
            t1, c1 = stops[i + 1]
            if t0 <= t <= t1:
                u = (t - t0) / (t1 - t0)
                return tuple(int(round(c0[j] + (c1[j] - c0[j]) * u)) for j in range(3))
        return stops[-1][1]
    return f


LIGHT = ramp([(0.0, PAPER), (0.30, (222, 236, 233)), (0.55, TEAL_WASH), (0.78, TEAL), (0.93, TEAL_DEEP), (1.0, INK)])
DARK = ramp([(0.0, INK), (0.42, (10, 46, 46)), (0.70, TEAL_DEEP), (0.88, TEAL), (1.0, MINT)])


def render(cols, colour, path, w=1600, h=460, dim=0.0):
    frames = len(cols)
    bins = len(cols[0])
    db = [[20 * math.log10(v + 1e-7) for v in col] for col in cols]
    top = max(max(row) for row in db)
    img = Image.new("RGB", (w, h))
    px = img.load()
    for x in range(w):
        col = db[int(x / w * (frames - 1))]
        for y in range(h):
            bin_ = int((1 - y / h) * (bins - 1))
            d = max(0.0, min(1.0, (col[bin_] - (top - 62)) / 62))
            c = colour(d ** 0.78)
            if dim:
                c = tuple(int(round(c[j] * (1 - dim) + INK[j] * dim)) for j in range(3))
            px[x, y] = c
    img.save(path)
    print("wrote", path, img.size)


def waveform(sig, path, w=1800, h=300):
    img = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    mid = h // 2
    step = max(1, len(sig) // w)
    pts = []
    for x in range(w):
        chunk = sig[x * step:(x + 1) * step]
        a = max((abs(v) for v in chunk), default=0.0)
        pts.append((x, int(mid - a * (h * 0.42))))
    d.line(pts, fill=TEAL + (255,), width=3, joint="curve")
    d.line([(x, h - y) for (x, y) in pts], fill=TEAL + (255,), width=3, joint="curve")
    img.save(path)
    print("wrote", path, img.size)


if __name__ == "__main__":
    sig = synth()
    render(stft(sig), LIGHT, os.path.join(OUT, "spectrogram.png"))
    render(stft(sig), DARK, os.path.join(OUT, "spectrogram-dark.png"))
    waveform(sig, os.path.join(OUT, "waveform.png"))
    render(stft(sig), DARK, os.path.join(OUT, "spectrogram-band.png"), w=2000, h=180, dim=0.52)
