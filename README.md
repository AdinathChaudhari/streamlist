# streamlist

Download any YouTube playlist or custom song list as high-quality M4A tracks — with embedded album art, metadata tags, and a ready-to-play `.m3u` playlist file.

Works with YouTube, YouTube Music, private playlists, and YouTube Premium content.

---

## What it does

- Downloads the best available audio stream for each track
- Re-encodes to M4A (AAC 256k) using the best encoder available on your machine
- Crops thumbnail to square and embeds as album art (600×600 sRGB JPEG)
- Writes tags: title, artist, album, album artist, track number
- Generates a `.m3u` playlist file so any music player can load the full collection
- **Resume & sync** — skips tracks already downloaded; re-running only fetches what's missing
- **Per-playlist cache** — instant re-run skipping with zero network calls for cached tracks
- **Auto-retry** — failed tracks are never cached, so the next run retries them automatically
- **Cache rebuild** — if a folder has existing files but no cache (older download), the cache is rebuilt from filenames before proceeding

**Output structure:**
```
My Playlist/
  01 - Track One.m4a
  02 - Track Two.m4a
  03 - Track Three.m4a
  My Playlist.m3u
  streamlist_cache.json
```

---

## Requirements

**Python 3.8+**

Install dependencies:
```bash
pip install yt-dlp openpyxl tqdm Pillow
```

**FFmpeg** (required for encoding):
```bash
# macOS
brew install ffmpeg

# Ubuntu / Debian
sudo apt install ffmpeg

# Windows
# Download from https://ffmpeg.org/download.html
```

---

## Usage

### Interactive mode (recommended for first use)
```bash
python streamlist.py
```

The script will ask:
1. Which browser to use for cookies (needed for Premium / private playlists)
2. Whether to use a YouTube URL or Excel file
3. The playlist name

### YouTube playlist URL
```bash
python streamlist.py --url "https://music.youtube.com/playlist?list=..."
```

### Excel file
```bash
python streamlist.py --excel my_songs.xlsx --name "My Playlist"
```

### All flags
```bash
python streamlist.py \
  --url "https://youtube.com/playlist?list=..." \
  --name "Summer Mix" \
  --out ~/Music \
  --browser safari \
  --no-notification
```

| Flag | Description |
|------|-------------|
| `--url` | YouTube or YouTube Music playlist / video URL |
| `--excel` | Path to `.xlsx` file with track list |
| `--name` | Override the playlist / output folder name |
| `--out` | Output directory (default: current directory) |
| `--browser` | Browser to read cookies from: `safari`, `chrome`, `firefox`, `edge`, `brave`, `opera` |
| `--no-notification` | Disable the completion sound |

---

## Excel format

Create a `.xlsx` file with these columns. Only `url` is required.

| url | title | artist |
|-----|-------|--------|
| https://youtube.com/watch?v=... | Get Lucky | Daft Punk |
| https://music.youtube.com/watch?v=... | Blinding Lights | The Weeknd |
| https://youtube.com/watch?v=... | | *(fetched from YouTube if blank)* |

A `sample_playlist.xlsx` template is included in this repo.

---

## Resume, sync & caching

streamlist is designed to be run repeatedly against the same playlist without re-downloading anything.

### How it works

After each successful download, the track's metadata is written to `streamlist_cache.json` inside the playlist folder:

```json
{
  "https://music.youtube.com/watch?v=abc123": {
    "title": "Get Lucky",
    "artist": "Daft Punk",
    "filename": "01 - Get Lucky.m4a",
    "duration": 248
  }
}
```

On re-runs:
- **Cache hit** → track is skipped instantly with zero network calls
- **Cache miss** → track is fetched and downloaded as normal
- **Failed tracks** → never written to cache, so the next run retries them automatically

### Adding new songs to a playlist

Just re-run the same command. Existing tracks are skipped; only new ones are downloaded. The `.m3u` is always rewritten to reflect the full current playlist.

### Older downloads (no cache file)

If a playlist folder already has `.m4a` files but no `streamlist_cache.json` (e.g. downloaded with an older version of streamlist), the cache is automatically rebuilt from the existing filenames before proceeding:

```
🔄 No cache found but 28 existing file(s) detected — rebuilding cache from filenames...
   ✅ Rebuilt cache for 28 track(s).
```

### Summary line

At the end of every run:
```
🎉 Done!  3 downloaded  |  28 skipped  |  1 failed
```

---

## Known limitations

### Music Premium exclusive tracks

Some tracks in YouTube Music playlists have a separate video ID that YouTube blocks from all third-party downloaders — even with valid login cookies and a Premium subscription. These are recordings that YouTube Music licenses exclusively and deliberately restricts at the API level. yt-dlp receives a hard "Video unavailable" response with no workaround.

**How to identify them:** the error will be:
```
❌ Failed: ERROR: [youtube] <video_id>: Video unavailable. This video is not available
```
even after the cookie retry attempt.

**Workarounds:**
- Search for the same song on regular YouTube — a different, downloadable upload almost always exists. Add that URL to your Excel file instead.
- Add the file manually (from another source) and add a dummy entry to `streamlist_cache.json` so streamlist stops retrying it:
  ```json
  "https://music.youtube.com/watch?v=<video_id>": {
    "title": "Song Name",
    "artist": "Artist Name",
    "filename": "22 - Song Name.m4a",
    "duration": 0
  }
  ```

### Age-restricted videos

Videos with a YouTube age restriction may fail even with cookies if the account has not verified its age in the browser session.

### Geo-blocked content

Videos restricted to specific countries will fail if your IP address is outside the allowed region. A VPN set to the appropriate country before running the script is the only workaround.

---

## YouTube Premium & private playlists

streamlist supports Premium-only and private/unlisted playlists by reading cookies directly from your browser session.

**Setup for macOS (Safari / Chrome):**

1. Make sure you're logged into YouTube in your browser
2. Go to **System Settings → Privacy & Security → Full Disk Access**
3. Add **Terminal** (or VS Code if running from there)
4. Run the script and choose your browser when prompted

> macOS needs Full Disk Access to let yt-dlp read the browser's cookie database.

**Windows:** Use `chrome` or `firefox` — Safari is not available on Windows.

---

## How it works

1. **Encoder detection** — tests `aac_at` (Apple hardware) → `libfdk_aac` → `aac` (FFmpeg native) and uses the best available
2. **Cache check** — loads `streamlist_cache.json`; rebuilds it from existing files if missing
3. **Input parsing** — fetches playlist metadata via yt-dlp (no download yet) or reads your Excel file
4. **Per-track pipeline** (inside a temp folder that auto-cleans):
   - Skips instantly if URL is in cache and file exists
   - Downloads best audio stream (opus/AAC/webm, whatever YouTube offers)
   - Downloads thumbnail → crops to 600×600 square → saves as sRGB JPEG
   - Re-encodes to M4A with embedded art and tags
   - Writes entry to cache immediately after success
5. **M3U generation** — rewrites the portable playlist file with relative paths

---

## Supported platforms

| Platform | Status |
|----------|--------|
| macOS | ✅ Full support (hardware encoder on Apple Silicon / Intel) |
| Linux | ✅ Full support |
| Windows | ✅ Works (use `chrome` or `firefox` for cookies; Safari not available) |

---

## License

MIT
