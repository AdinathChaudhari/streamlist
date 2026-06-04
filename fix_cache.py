#!/usr/bin/env python3
"""
fix_cache.py
------------
One-time script to strip the "01 - " number prefix from M4A filenames
and update streamlist_cache.json to match.

If two tracks would end up with the same filename, the artist is appended
to disambiguate: "Song Title (Artist).m4a"

Usage:
  python fix_cache.py /path/to/playlist/folder
"""

import re
import sys
import json
from pathlib import Path

CACHE_FILE = 'streamlist_cache.json'


def safe_filename(name, max_len=120):
    name = re.sub(r'[\\/:*?"<>|]', '-', name)
    name = name.strip('. ')
    return name[:max_len] if len(name) > max_len else name


def strip_prefix(name):
    return re.sub(r'^\d+\s*-\s*', '', name)


def main():
    if len(sys.argv) < 2:
        folder = Path(input("Path to playlist folder: ").strip().strip('"\''))
    else:
        folder = Path(sys.argv[1])

    if not folder.is_dir():
        print(f"❌ Folder not found: {folder}")
        sys.exit(1)

    cache_path = folder / CACHE_FILE
    if not cache_path.exists():
        print(f"❌ No {CACHE_FILE} found in {folder}")
        sys.exit(1)

    cache = json.loads(cache_path.read_text(encoding='utf-8'))

    # First pass: compute the desired new filename for every entry,
    # tracking which bare titles appear more than once so we can
    # disambiguate with the artist name.
    bare_title_count: dict[str, int] = {}
    for url, entry in cache.items():
        title = entry.get('title', strip_prefix(entry.get('filename', '')).removesuffix('.m4a'))
        bare = safe_filename(title).lower()
        bare_title_count[bare] = bare_title_count.get(bare, 0) + 1

    renamed = 0
    skipped = 0

    for url, entry in cache.items():
        old_name = entry.get('filename', '')
        title    = entry.get('title', strip_prefix(old_name).removesuffix('.m4a'))
        artist   = entry.get('artist', '')

        bare = safe_filename(title).lower()
        if bare_title_count[bare] > 1 and artist:
            new_stem = f"{safe_filename(title)} ({safe_filename(artist)})"
        else:
            new_stem = safe_filename(title)

        new_name = f"{new_stem}.m4a"

        if new_name == old_name:
            skipped += 1
            continue

        old_path = folder / old_name
        new_path = folder / new_name

        # If the target already exists and isn't the source file, try appending artist anyway
        if new_path.exists() and new_path != old_path:
            if artist:
                fallback_stem = f"{safe_filename(title)} ({safe_filename(artist)})"
                new_name = f"{fallback_stem}.m4a"
                new_path = folder / new_name
            if new_path.exists() and new_path != old_path:
                print(f"  ⚠️  Conflict, skipping: {new_name}")
                continue

        if old_path.exists():
            old_path.rename(new_path)
            print(f"  ✅ {old_name}  →  {new_name}")
        else:
            print(f"  ⚠️  File not on disk, updating cache only: {old_name}  →  {new_name}")

        entry['filename'] = new_name
        renamed += 1

    cache_path.write_text(json.dumps(cache, indent=2, ensure_ascii=False), encoding='utf-8')

    print(f"\n{'═'*52}")
    print(f"Done!  {renamed} renamed  |  {skipped} already clean")
    print(f"{'═'*52}")


if __name__ == '__main__':
    main()
