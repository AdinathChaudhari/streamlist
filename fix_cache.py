#!/usr/bin/env python3
"""
fix_cache.py
------------
Fixes streamlist_cache.json and M4A filenames in a playlist folder.

What it does:
  1. Strips "01 - " style number prefixes from filenames (legacy format)
  2. Strips any literal suffix you specify from track titles and filenames
     e.g. ". [indie playlist]" or " (official audio)"
  3. Disambiguates tracks that share the same title by appending the artist name

Usage:
  python fix_cache.py /path/to/folder
  python fix_cache.py /path/to/folder --strip ". [indie playlist]"
  python fix_cache.py /path/to/folder --strip ". [indie playlist]" --strip " (official audio)"
"""

import re
import sys
import json
import argparse
from pathlib import Path

CACHE_FILE = 'streamlist_cache.json'


def safe_filename(name, max_len=120):
    name = re.sub(r'[\\/:*?"<>|]', '-', name)
    name = name.strip('. ')
    return name[:max_len] if len(name) > max_len else name


def strip_number_prefix(name):
    return re.sub(r'^\d+\s*-\s*', '', name)


def apply_strip_patterns(text, patterns):
    for p in patterns:
        # Case-insensitive literal strip from the end of the string
        escaped = re.escape(p)
        text = re.sub(escaped + r'\s*$', '', text, flags=re.IGNORECASE).rstrip()
    return text


def main():
    parser = argparse.ArgumentParser(description='Fix streamlist cache and filenames')
    parser.add_argument('folder', nargs='?', help='Path to playlist folder')
    parser.add_argument('--strip', action='append', default=[], metavar='SUFFIX',
                        help='Literal suffix to strip from titles/filenames (repeatable)')
    args = parser.parse_args()

    folder = Path(args.folder) if args.folder else Path(input("Path to playlist folder: ").strip().strip('"\''))

    if not folder.is_dir():
        print(f"❌ Folder not found: {folder}")
        sys.exit(1)

    cache_path = folder / CACHE_FILE
    if not cache_path.exists():
        print(f"❌ No {CACHE_FILE} found in {folder}")
        sys.exit(1)

    raw_cache = json.loads(cache_path.read_text(encoding='utf-8'))

    # Normalize music.youtube.com → www.youtube.com keys, merging duplicates.
    # When two entries map to the same key, keep the one with the shorter/cleaner title.
    cache = {}
    for url, entry in raw_cache.items():
        normalized = url.replace('music.youtube.com', 'www.youtube.com')
        if normalized not in cache:
            cache[normalized] = entry
        else:
            existing_title = cache[normalized].get('title', '')
            new_title = entry.get('title', '')
            if len(new_title) < len(existing_title):
                print(f"  🔀 Merging duplicate: kept \"{new_title}\" over \"{existing_title}\"")
                cache[normalized] = entry

    strip_patterns = args.strip

    if not strip_patterns:
        print("\nEnter suffix(es) to strip from track titles (e.g. '. [indie playlist]').")
        print("Press Enter with no input when done.\n")
        while True:
            val = input("Strip suffix (or Enter to skip): ").strip()
            if not val:
                break
            strip_patterns.append(val)

    if strip_patterns:
        print(f"   Stripping suffix(es): {strip_patterns}")

    def clean_title(title):
        return apply_strip_patterns(title, strip_patterns).strip()

    # First pass: compute clean titles and count duplicates for disambiguation
    bare_title_count: dict[str, int] = {}
    for url, entry in cache.items():
        raw_title = entry.get('title', strip_number_prefix(entry.get('filename', '')).removesuffix('.m4a'))
        title = clean_title(raw_title)
        bare = safe_filename(title).lower()
        bare_title_count[bare] = bare_title_count.get(bare, 0) + 1

    renamed = 0
    skipped = 0

    for url, entry in cache.items():
        old_name  = entry.get('filename', '')
        raw_title = entry.get('title', strip_number_prefix(old_name).removesuffix('.m4a'))
        title     = clean_title(raw_title)
        artist    = entry.get('artist', '')

        bare = safe_filename(title).lower()
        if bare_title_count[bare] > 1 and artist:
            new_stem = f"{safe_filename(title)} ({safe_filename(artist)})"
        else:
            new_stem = safe_filename(title)

        new_name = f"{new_stem}.m4a"

        # Update the title in cache if it changed
        if title != raw_title:
            entry['title'] = title

        if new_name == old_name:
            skipped += 1
            continue

        old_path = folder / old_name
        new_path = folder / new_name

        # If target already exists and isn't the source, try artist disambiguation
        if new_path.exists() and new_path != old_path:
            if artist:
                new_name = f"{safe_filename(title)} ({safe_filename(artist)}).m4a"
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
    print(f"Done!  {renamed} renamed/updated  |  {skipped} already clean")
    print(f"{'═'*52}")


if __name__ == '__main__':
    main()
