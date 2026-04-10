# Teams Chat Archiver

This is a Chrome extension that archives the currently open Microsoft Teams web chat into a long-term local folder structure.

## What it does today

- Adds a popup with an `Archive Current Chat` button
- Scrolls upward to load older messages from the open Teams web chat
- Saves each run as an immutable snapshot JSON file
- Rebuilds a merged `latest.json`
- Generates a chat overview `latest.html`, yearly HTML transcript files, Markdown files, or both
- Can archive either the current chat or the chats currently visible in the Teams chat list
- Saves pasted inline message images into `assets/images/` and references them from HTML/Markdown

## Archive layout

After you choose a root archive folder, the extension creates a structure like:

```text
YourArchiveFolder/
  index.html
  tompalski-hermosilla/
    manifest.json
    latest.json
    latest.html
    html/
      2025.html
      2026.html
    md/
      2026-04.md
      2026-05.md
    snapshots/
      2026-04-09T09-15-22.json
      2026-04-16T10-41-05.json
```

## Current limitations

- It only works on the currently open chat
- Teams changes its HTML structure often, so selectors may need adjustment
- Snapshot merging currently deduplicates by message id when present, otherwise by author + timestamp + text
- Markdown export is written as monthly files
- `latest.html` is a chat overview page with links to yearly HTML transcript files and recent messages
- Batch mode only walks the chats currently visible in the chat list and may skip entries if Teams changes its list structure
- The archive root gets an `index.html` page that links to each chat archive

## How to load it in Chrome

1. Open `chrome://extensions`
2. Turn on **Developer mode**
3. Click **Load unpacked**
4. Select this folder: `teams-archive-extension`
5. Open Teams on the web: `https://teams.microsoft.com`
6. Open a chat and click the extension icon
7. Click **Archive Current Chat** or **Archive Visible Chats**
8. Choose the root folder where you want archives stored
9. Pick `HTML`, `Markdown`, or `HTML and Markdown` in the popup as needed
10. Use the `Include pasted images` toggle if you want pasted images archived locally

## Next step

Possible next steps are making batch mode scroll the chat list to discover more chats, adding skip/resume behavior, and improving support for attachments and meeting chats.
