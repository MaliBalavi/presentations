#!/usr/bin/env python3
"""Patch the Office theme inside a .pptx so PowerPoint's theme palette and
theme fonts match the deck.

pptxgenjs writes pres.theme.headFontFace / bodyFontFace into
ppt/theme/theme1.xml, but it has no API for theme colors: the file keeps the
default Office blue palette. This script rewrites the theme name, the color
scheme and the font scheme in place and leaves every other archive entry
untouched.

Usage:
    apply_theme.py deck.pptx '{"name": "...", "colors": {...}, "headFontFace": "...", "bodyFontFace": "..."}'
"""

import json
import os
import re
import sys
import tempfile
import zipfile

THEME_PART = "ppt/theme/theme1.xml"

# Order matters only for readability; each key is replaced independently.
COLOR_KEYS = [
    "dk1", "lt1", "dk2", "lt2",
    "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
    "hlink", "folHlink",
]


def _replace_scheme_color(xml, key, value):
    pattern = re.compile(r"(<a:%s>).*?(</a:%s>)" % (key, key), re.DOTALL)
    replacement = r'\g<1><a:srgbClr val="%s"/>\g<2>' % value
    patched, count = pattern.subn(replacement, xml, count=1)
    if count != 1:
        raise SystemExit("theme: color %r not found in %s" % (key, THEME_PART))
    return patched


def patch_theme(xml, theme):
    name = theme["name"]
    colors = theme["colors"]
    head = theme["headFontFace"]
    body = theme["bodyFontFace"]

    xml = re.sub(r'(<a:theme [^>]*name=")[^"]*(")', r"\g<1>%s\g<2>" % name, xml, count=1)
    xml = re.sub(r'(<a:clrScheme name=")[^"]*(")', r"\g<1>%s\g<2>" % name, xml, count=1)
    xml = re.sub(r'(<a:fontScheme name=")[^"]*(")', r"\g<1>%s\g<2>" % name, xml, count=1)

    for key in COLOR_KEYS:
        xml = _replace_scheme_color(xml, key, colors[key])

    xml = re.sub(
        r'(<a:majorFont><a:latin typeface=")[^"]*(")',
        r"\g<1>%s\g<2>" % head,
        xml,
        count=1,
    )
    xml = re.sub(
        r'(<a:minorFont><a:latin typeface=")[^"]*(")',
        r"\g<1>%s\g<2>" % body,
        xml,
        count=1,
    )
    return xml


def rewrite_pptx(path, xml):
    directory = os.path.dirname(os.path.abspath(path))
    fd, tmp = tempfile.mkstemp(suffix=".pptx", dir=directory)
    os.close(fd)
    try:
        with zipfile.ZipFile(path, "r") as src, zipfile.ZipFile(tmp, "w") as dst:
            for item in src.infolist():
                data = src.read(item.filename)
                if item.filename == THEME_PART:
                    data = xml.encode("utf-8")
                dst.writestr(item, data)
        os.replace(tmp, path)
    finally:
        if os.path.exists(tmp):
            os.unlink(tmp)


def main():
    if len(sys.argv) != 3:
        raise SystemExit("usage: apply_theme.py <deck.pptx> <theme-json>")
    path, theme_json = sys.argv[1], sys.argv[2]
    theme = json.loads(theme_json)

    with zipfile.ZipFile(path, "r") as archive:
        xml = archive.read(THEME_PART).decode("utf-8")

    rewrite_pptx(path, patch_theme(xml, theme))


if __name__ == "__main__":
    main()
