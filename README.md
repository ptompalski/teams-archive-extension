# Teams Chat Archiver

Teams Chat Archiver is a Chrome extension for saving a local copy of conversations from Microsoft Teams on the web. It can archive the chat you currently have open or work through all chats visible in your Teams chat list.

Use it when you want to keep an accessible, long-term record of important conversations outside the Teams interface. The extension copies messages and, optionally, pasted images into a folder you choose. It does not remove or change anything in Teams.

In practical terms, the extension automates the copying and pasting you could do yourself to preserve important messages. Everything it collects is stored locally in the folder you select, and no AI is used to read, analyze, or process your messages.

## What you get

The extension creates two useful versions of the archive in the same folder:

- **A simple webpage for browsing:** Open `index.html` in Chrome or another browser to see your archived chats. Select a chat to view the messages.
- **A structured data archive:** JSON snapshot and merged-data files preserve the collected message information for backup, auditing, or future processing.

The browsing webpage uses files stored inside the archive folder. It does not require a web server or an internet connection after the archive has been created. Keep the entire folder together when moving or backing it up so that links and saved images continue to work.

## Main features

- Archives the currently open chat or all chats visible in the Teams chat list
- Works with Teams at `https://teams.cloud.microsoft/`
- Scrolls through each chat to collect older messages that are available in the Teams page
- Keeps a new snapshot from every archive run and combines snapshots into the latest record
- Creates a main `index.html` page linking to every archived chat
- Creates a summary page and readable yearly transcript pages for each chat

## Browse the archive

After an archive finishes:

1. Open the folder you selected when running the extension.
2. Double-click `index.html` to open the archive homepage in your browser.
3. Select a chat to open its overview page.
4. Use the year links to browse the full conversation history.

You can reopen `index.html` at any time. Running the extension again updates the browsing pages while preserving the earlier snapshots.

## Folder contents

The extension creates a structure similar to this:

```text
YourArchiveFolder/
  index.html                 Main page for browsing all archived chats
  xxx/                       Folder for one chat
    manifest.json            Chat summary information
    latest.json              Latest combined message data
    latest.html              Browsable overview for the chat
    html/
      2025.html              Readable transcript for one year
      2026.html
    snapshots/               Unchanged records from individual archive runs
      2026-04-09T09-15-22.json
      2026-04-16T10-41-05.json
    assets/
      images/                Saved pasted images, when enabled
        2026-04/
          a1b2c3d4.png
```

Each run adds a dated file under `snapshots/`. The extension combines those snapshots into `latest.json` and rebuilds the HTML browsing pages. Duplicate messages are removed during this process. Images are stored separately and linked from the transcript pages so the same image does not need to be saved repeatedly.

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
