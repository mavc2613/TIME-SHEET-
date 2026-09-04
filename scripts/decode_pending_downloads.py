#!/usr/bin/env python3
"""
Scans the Claude tool-results cache for pending Google Drive download_file_content
dumps, decodes each into punch_images/<outdir>/<title>, then marks it processed
(renamed with .done suffix) so re-runs don't redo work.

Usage: python3 scripts/decode_pending_downloads.py <outdir-under-punch_images>
"""
import json
import base64
import glob
import os
import sys

CACHE_DIR = "/root/.claude/projects/-home-user-TIME-SHEET-/114a70c0-46b4-547c-8926-25f7c73af835/tool-results"
OUT_ROOT = "/home/user/TIME-SHEET-/punch_images"


def main():
    outdir_name = sys.argv[1] if len(sys.argv) > 1 else "misc"
    outdir = os.path.join(OUT_ROOT, outdir_name)
    os.makedirs(outdir, exist_ok=True)

    pattern = os.path.join(CACHE_DIR, "mcp-Google_Drive-download_file_content-*.txt")
    saved = []
    for fn in sorted(glob.glob(pattern)):
        try:
            with open(fn) as f:
                d = json.load(f)
        except (json.JSONDecodeError, UnicodeDecodeError):
            continue
        if "content" not in d or "title" not in d:
            continue
        out_path = os.path.join(outdir, d["title"])
        with open(out_path, "wb") as o:
            o.write(base64.b64decode(d["content"]))
        saved.append(d["title"])
        os.rename(fn, fn + ".done")

    for title in saved:
        print(title)
    print(f"-- {len(saved)} file(s) decoded to {outdir}")


if __name__ == "__main__":
    main()
