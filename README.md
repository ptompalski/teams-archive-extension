# Teams Chat Archiver

This is a Chrome extension that archives Microsoft Teams web chats into a long-term local folder structure.

## What it does

- Adds a popup with `Archive Current Chat` and `Archive All Chats` buttons
- Works with Teams on the web at `https://teams.cloud.microsoft/` and the legacy `https://teams.microsoft.com/` address
- Scrolls upward in the open chat to load older messages before export
- Saves each archive run as an immutable snapshot JSON file
- Rebuilds a merged `latest.json` from all saved snapshots for that chat
- Generates a chat overview `latest.html`
- Generates yearly HTML transcript files under `html/`
- Builds a root `index.html` page that links to all archived chats
- Archives pasted inline images into `assets/images/` and renders them in HTML when image export is enabled
- Preserves a per-chat `manifest.json` with summary metadata such as chat name, archive dates, and message count
- Uses the visible Teams chat name for archive naming where possible
- Sorts the root archive index using the visible Teams chat order when that order is available

## Archive layout

After you choose a root archive folder, the extension creates a structure like:

```text
YourArchiveFolder/
  index.html
  xxx/
    manifest.json
    latest.json
    latest.html
    html/
      2025.html
      2026.html
    snapshots/
      2026-04-09T09-15-22.json
      2026-04-16T10-41-05.json
    assets/
      images/
        2026-04/
          a1b2c3d4.png
```

## How the archive works

- Each run creates a new snapshot in `snapshots/`
- The extension merges snapshots to rebuild `latest.json`
- `latest.html` is a chat overview page with summary information, yearly links, and recent messages
- `html/<year>.html` files contain the readable transcript split by year
- Images are stored as separate files and referenced from the HTML pages
- Image filenames are based on content hashes so the same pasted image does not need to be saved multiple times
- The extension validates that key output files were written before showing a finished status

## Current limitations

- Batch mode only archives chats that are currently visible in the Teams chat list
- Teams changes its HTML structure often, so selectors may need adjustment over time
- Snapshot merging deduplicates by message id when present, otherwise by author + timestamp + text + image source
- Only pasted inline images are currently archived
- Other attachment types, reactions, and some special message layouts may not be captured perfectly
- The root chat order can only match the Teams order for chats that have been archived while visible in the left chat list

## How to download and load it in Chrome

1. Open the GitHub repository in your browser.
2. Download the project files:
   - Click **Code** > **Download ZIP**, then extract the ZIP file somewhere on your computer.
   - Or clone the repository with git if you prefer.
3. Open `chrome://extensions` in Chrome.
4. Turn on **Developer mode**.
5. Click **Load unpacked**.
6. Select the extracted `teams-archive-extension` folder from the repository.
7. Open Teams on the web at `https://teams.cloud.microsoft/`.
8. Open a chat and click the extension icon.
9. Choose whether to include pasted images.
10. Click **Archive Current Chat** or **Archive All Chats**.
11. Choose the root folder where you want the archive stored.
12. After the tool finishes, which may take a while for larger chats, open that archive folder and open `index.html` to browse the archived chats.

While the tool is running, avoid interacting with the Teams tab. You can switch to another tab or another application, but do not click, scroll, change chats, or navigate away in the Teams tab until the status shows that the archive is finished.
