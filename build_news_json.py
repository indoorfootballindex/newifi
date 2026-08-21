#!/usr/bin/env python3
"""
Indoor Football Index — news article builder (Word doc source).

Reads every .docx in a folder — one file per article — and builds
news.json for news.html and article.html, plus a news-images/ folder
holding any images embedded in the articles.

Filename convention (no need to type metadata inside the document):

    [FEATURED_]Kicker - Title.docx

  - "Season Recap - Strike Force's Historic Season.docx"
        -> kicker "Season Recap", title "Strike Force's Historic Season"
  - "FEATURED_Season Recap - Strike Force's Historic Season.docx"
        -> same, plus featured on the home page
  - No " - " in the filename -> the whole filename becomes the title,
    no kicker

The article body (formatting, images, headings, lists) comes straight from
the Word document via pandoc. The publish date is pulled automatically
from the file's own "created" metadata — nothing to type by hand.

Requires: pandoc

Usage:
    python3 build_news_json.py path/to/articles_folder path/to/news.json
"""

import sys
import os
import re
import glob
import json
import shutil
import subprocess
import datetime


def slug(text):
    return re.sub(r"[^a-z0-9]+", "-", str(text).lower()).strip("-")


def parse_filename(path):
    name = os.path.splitext(os.path.basename(path))[0]
    featured = False
    if name.upper().startswith("FEATURED_"):
        featured = True
        name = name[len("FEATURED_"):]
    if " - " in name:
        kicker, title = name.split(" - ", 1)
        kicker, title = kicker.strip(), title.strip()
    else:
        kicker, title = None, name.strip()
    return kicker, title, featured


def get_date(path):
    """The file's own last-modified time is far more reliable than the
    document's embedded metadata, which can be stale, copied from a
    template, or just wrong depending on what tool last touched the file."""
    return datetime.datetime.fromtimestamp(os.path.getmtime(path)).strftime("%Y-%m-%d")


def convert_body(path, article_slug, images_root):
    """Runs pandoc to get HTML + extracted images, then rewrites image src
    paths to live under news-images/<slug>/ so multiple articles' images
    (which pandoc names generically, e.g. media/image1.png) never collide."""
    media_dir = os.path.join(images_root, article_slug)
    os.makedirs(media_dir, exist_ok=True)

    result = subprocess.run(
        ["pandoc", path, "-t", "html", f"--extract-media={media_dir}"],
        capture_output=True, text=True, encoding="utf-8",
    )
    if result.returncode != 0:
        raise RuntimeError(f"pandoc failed on {path}: {result.stderr}")

    html = result.stdout

    # pandoc extracts into <media_dir>/media/... — flatten that up one level
    nested = os.path.join(media_dir, "media")
    if os.path.isdir(nested):
        for fn in os.listdir(nested):
            shutil.move(os.path.join(nested, fn), os.path.join(media_dir, fn))
        os.rmdir(nested)

    # rewrite src="<whatever prefix>/media/xyz.png" -> src="news-images/<slug>/xyz.png"
    # (pandoc uses whatever path was passed to --extract-media verbatim,
    # absolute or relative, so match on the trailing /media/<filename> part
    # rather than assuming a specific prefix)
    html = re.sub(
        r'src="[^"]*/media/([^"]+)"',
        lambda m: f'src="news-images/{article_slug}/{m.group(1)}"',
        html,
    )
    return html


def main():
    if len(sys.argv) != 3:
        print("Usage: python3 build_news_json.py <folder_of_docx> <news.json>")
        sys.exit(1)
    src_dir, out_path = sys.argv[1], sys.argv[2]

    docx_files = sorted(glob.glob(os.path.join(src_dir, "*.docx")))
    if not docx_files:
        print(f"No .docx files found in {src_dir} — writing an empty news.json.")
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump([], f)
        return

    out_dir = os.path.dirname(os.path.abspath(out_path)) or "."
    images_root = os.path.join(out_dir, "news-images")

    articles = []
    seen_slugs = {}
    for path in docx_files:
        fname = os.path.basename(path)
        # skip Word's temp lock files (~$Article.docx) if the doc is open
        if fname.startswith("~$"):
            continue
        try:
            kicker, title, featured = parse_filename(path)
            base_slug = slug(title)
            s = base_slug
            n = 2
            while s in seen_slugs:
                s = f"{base_slug}-{n}"
                n += 1
            seen_slugs[s] = True

            body_html = convert_body(path, s, images_root)
            date = get_date(path)

            articles.append({
                "slug": s,
                "title": title,
                "kicker": kicker,
                "body": body_html,
                "date": date,
                "featured": featured,
            })
            print(f"  {fname} -> \"{title}\"" + (" [featured]" if featured else ""))
        except Exception as e:
            print(f"  SKIPPED {fname}: {e}")

    articles.sort(key=lambda a: a["date"] or "", reverse=True)

    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(articles, f, ensure_ascii=False)

    print(f"\nWrote {out_path}: {len(articles)} articles ({sum(1 for a in articles if a['featured'])} featured)")
    print(f"Images (if any) written under {images_root}/")


if __name__ == "__main__":
    main()
