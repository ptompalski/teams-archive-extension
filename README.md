# Teams Chat Archiver

Teams Chat Archiver is a Chrome extension for saving a local copy of conversations from Microsoft Teams on the web. It can archive the chat or channel you currently have open, conversations in selected sidebar groups, or all conversations visible in your Teams chat list.

Use it when you want to keep an accessible, long-term record of important conversations outside the Teams interface. The extension copies messages and, optionally, pasted images into a folder you choose. It does not remove or change anything in Teams.

In practical terms, the extension automates the copying and pasting you could do yourself to preserve important messages. Everything it collects is stored locally in the folder you select, and no AI is used to read, analyze, or process your messages.

## What you get

The extension creates two useful versions of the archive in the same folder:

- **A simple webpage for browsing:** Open `index.html` in Chrome or another browser to see your archived chats. Select a chat to view the messages.
- **A structured data archive:** JSON snapshot and merged-data files preserve the collected message information for backup, auditing, or future processing.

The browsing webpage uses files stored inside the archive folder. It does not require a web server or an internet connection after the archive has been created. Keep the entire folder together when moving or backing it up so that links and saved images continue to work.

## Main features

- Archives the currently open chat or all chats visible in the Teams chat list
- Archives selected sidebar groups, including Favorites and custom groups, with muted chats and channels included
- Captures channel posts and replies with thread references, post subjects, link URLs, and optional inline images
- Works with Teams at `https://teams.cloud.microsoft/`
- Scrolls through each chat to collect older messages that are available in the Teams page
- Keeps a new snapshot from every archive run and combines snapshots into the latest record
- Creates a main `index.html` page linking to every archived chat
- Creates a summary page and readable yearly transcript pages for each chat

## Archive selected groups

1. Open the Teams chat list and click the extension icon.
2. Click **Archive Selected Groups** to load the sidebar groups.
3. Check one or more groups, such as Favorites, CAT, or other.
4. Click **Archive Groups** and choose your archive folder.

The extension expands selected groups, follows their **See more** controls, and scrolls the sidebar to collect chat and channel rows. Selected groups can remain expanded after the run. Conversations appearing in multiple selected groups are collected once by their Teams ID.

**Archive All Chats** excludes the built-in **Teams and channels** section. Channels placed in Favorites or custom groups, such as other, are still included. **Archive Current Chat** can still archive a channel you explicitly open from the built-in section.

For channels, it expands truncated messages and opens reply summaries when not all replies are shown. It scans channel history in both directions and preserves each message's thread-root ID in JSON and post/reply references in the transcript. If a message or reply thread cannot be fully loaded, or channel history reaches its safety limit, the extension reports an error instead of saving a known partial channel snapshot. A batch run can still contain successfully saved conversations from before that error.

Channel panes that are still loading are allowed to settle before a non-scrolling result is accepted. Threads temporarily removed from the page by Teams virtualization are deferred to the history scan, with a final reply-count check before saving.

Opened reply threads are scanned in both directions. Deleted replies exposed by Teams are included as stable-ID deletion markers and count toward the advertised reply total; the extension cannot recover deleted content that was never archived. Reply-loading errors include the collected and expected counts to help diagnose unsupported layouts.

If Teams stops loading a group or the safety limit is reached, the selected-group run stops before archiving any chats rather than using a known partial list. Changing the active Teams tab after loading the group picker requires reloading the group list. Existing same-day archive skipping still applies, and the archive folder structure does not change.

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

When a channel reply appears first as a preview and later with fuller content, the merged record keeps the longer text, available author and timestamp, and collected images. Later previews do not replace that fuller text or erase its metadata; this also applies when rebuilding from older snapshots.

## Current limitations

- Batch mode only archives chats and channels that are currently visible in the Teams chat list
- Selected-group mode attempts to load all chats and channels in the chosen groups, but depends on Teams exposing its group, pagination, and scrolling controls correctly
- Channel loading depends on Teams exposing reply counts and expansion controls; permission-restricted or unsupported thread views may fail rather than produce a partial snapshot
- Channel support is regression-tested with synthetic markup and extraction from a saved Teams page; live Teams navigation and full-history loading still need validation
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
10. Click **Archive Current Chat**, **Archive All Chats**, or **Archive Selected Groups**. For selected groups, check the groups and click **Archive Groups**.
11. Choose the root folder where you want the archive stored.
12. After the tool finishes, which may take a while for larger chats, open that archive folder and open `index.html` to browse the archived chats.

While the tool is running, avoid interacting with the Teams tab. You can switch to another tab or another application, but do not click, scroll, change chats, or navigate away in the Teams tab until the status shows that the archive is finished.

## Development checks

Serve the repository root with a local static HTTP server, then open `/tests/group-archive.html` in a browser. The dependency-free regression page tests the actual content, popup, and background scripts against synthetic Teams markup and mocked extension APIs. It covers group boundaries, muted chats, channels, deduplication, expansion, pagination, virtualized rows, channel metadata and thread navigation, bidirectional channel history, error handling, transcript output, and popup routing. It does not access Teams or write archives.
