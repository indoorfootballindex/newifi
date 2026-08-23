#!/usr/bin/env python3
"""
Indoor Football Index — hero slideshow image list builder.

Scans your images folder for photos and writes hero-images.json, which
index.html's homepage slideshow reads at runtime. Add or remove photos
from the images folder, re-run this (or the full update_site.py), and the
slideshow picks up the change automatically — no HTML editing required.

Site assets that live in the same folder (the logo, the favicon) are
excluded by name so they never accidentally show up in the slideshow.
Add any other non-photo filenames to EXCLUDE below if you add more site
assets to this folder later.

Usage:
    python3 build_hero_images.py path/to/images path/to/hero-images.json
"""

import sys
import os
import json

PHOTO_EXTENSIONS = (".jpg", ".jpeg", ".png", ".webp")
EXCLUDE = {"logo.png", "favicon.ico"}


def main():
    if len(sys.argv) != 3:
        print("Usage: python3 build_hero_images.py <images_folder> <hero-images.json>")
        sys.exit(1)
    images_dir, out_path = sys.argv[1], sys.argv[2]

    if not os.path.isdir(images_dir):
        print(f"NOTE: no images folder found at '{images_dir}' — writing an empty hero-images.json.")
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump([], f)
        return

    photos = []
    for fn in sorted(os.listdir(images_dir)):
        if fn in EXCLUDE:
            continue
        if fn.lower().endswith(PHOTO_EXTENSIONS):
            photos.append(f"images/{fn}")

    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(photos, f, ensure_ascii=False)

    print(f"Wrote {out_path}: {len(photos)} photo(s) found in '{images_dir}'")


if __name__ == "__main__":
    main()
