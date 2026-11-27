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
2. In the archiver window, click **Archive Selected Groups** to load the sidebar groups.
3. Check one or more groups, such as Favorites, CAT, or other.
4. Click **Archive Groups** and choose your archive folder.

The extension expands selected groups, follows their **See more** controls, and scrolls the sidebar to collect chat and channel rows. Selected groups can remain expanded after the run. Conversations appearing in multiple selected groups are collected once by their Teams ID.

**Archive All Chats** excludes the built-in **Teams and channels** section. Channels placed in Favorites or custom groups, such as other, are still included. **Archive Current Chat** can still archive a channel you explicitly open from the built-in section.

For channels, it expands truncated messages and opens reply summaries when not all replies are shown. It scans channel history in both directions and preserves each message's thread-root ID in JSON and post/reply references in the transcript. If a message or reply thread cannot be fully loaded, or channel history reaches its safety limit, the extension reports an error instead of saving a known partial channel snapshot. A batch run can still contain successfully saved conversations from before that error.

Channel panes that are still loading are allowed to settle before a non-scrolling result is accepted. Threads temporarily removed from the page by Teams virtualization are deferred to the history scan, with a final reply-count check before saving.

Opened reply threads are scanned in both directions. Deleted replies exposed by Teams are included as stable-ID deletion markers and count toward the advertised reply total; the extension cannot recover deleted content that was never archived. Reply-loading errors include the collected and expected counts to help diagnose unsupported layouts.

If Teams stops loading a group or the safety limit is reached, the selected-group run stops before archiving any chats rather than using a known partial list. The archiver stays connected to its original Teams tab even when you switch tabs or windows. Reconnecting with **Open Teams Website** to a different Teams tab clears the group selection; reload the group list before continuing. Existing same-day archive skipping still applies, and the archive folder structure does not change.

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
9. The extension opens a persistent archiver window connected to that Teams tab. Choose whether to include pasted images.
10. Click **Archive Current Chat**, **Archive All Chats**, or **Archive Selected Groups**. For selected groups, check the groups and click **Archive Groups**.
11. Choose the root folder where you want the archive stored.
12. After the tool finishes, which may take a while for larger chats, open that archive folder and open `index.html` to browse the archived chats.

Keep the archiver window open until the status shows **Finished**. Unlike a toolbar popup, this window does not close when you click elsewhere or switch applications. Clicking the extension icon again brings the existing archiver window forward without changing its connected Teams tab. Closing or reloading the archiver window still interrupts the run; the browser requests confirmation while it is busy.

While the tool is running, avoid interacting with the connected Teams tab. You can switch to another tab or another application, but do not click, scroll, change chats, or navigate away in the Teams tab until the status shows that the archive is finished. Background tabs may load more slowly because of browser throttling; keep Teams visible in its own window if that occurs.

The status follows the entire archive process: history and reply loading, pasted-image extraction, snapshot merging, image processing, transcript writes, archive-wide index/search rebuilding, and final file verification. It includes message, image, snapshot, or transcript counts where available. Batch progress retains the current chat's name and position. History scan counters indicate safety limits, not a completion percentage; the amount of Teams history is not known in advance. Only **Finished** means the full run has completed.

The **Overall archive progress** bar covers the current chat, all visible chats, or all conversations loaded from selected groups. It is indeterminate while the chat list is loading, then counts each saved, skipped, or failed chat as one processed step. Rebuilding and verifying the archive index/search files is one final step, so the bar reaches 100% only after verification. The percentage measures completed steps, not time remaining: larger chats can take much longer than smaller ones. The label reports saved, skipped, and failed counts; 100% can include failed chats and does not mean every chat was archived successfully.

### Troubleshooting archive failures

- **Folder selection was cancelled or access was denied:** The browser uses the same error for cancellation, blocked folders, and denied read/write access. Create a dedicated archive subfolder if needed, navigate into it in the picker, select it, and approve access.
- **Archive silently stopped after clicking elsewhere (older versions):** Older versions ran in Chrome's toolbar popup, which Chrome destroys on loss of focus. Reload the updated extension in `chrome://extensions` to use the persistent window.
- **Teams stopped scrolling / history reached its safety limit:** The history scan could not verify the oldest messages. It retries transient scroll resets and tracks collected message identities rather than just counting visible rows. It reports an error instead of saving a known incomplete chat snapshot. Diagnostics include collected message count, scroll position and dimensions, page visibility, and focus state; copy the full status when reporting a problem.
- **Lost connection to the Teams tab:** Teams was closed, reloaded, navigated away, or possibly discarded by the browser. Keep the Teams tab open and retry. **Open Teams Website** reconnects if the original tab is no longer available.
- **Archive failed while ...:** The status identifies the failing stage and includes the original error details. Folder creation and text-file write errors also identify the affected folder or file. The archiver window's developer console retains the exception for troubleshooting. Files saved before the error may remain; a failed run is not a verified complete archive.
- **Skipped chats were already archived today:** Batch actions skip chats with an archive from today. This is not a configurable cooldown. **Archive Current Chat** bypasses that check and can capture newer messages immediately.

## Extension updates

When the archiver window opens, it checks the latest stable [GitHub release](https://github.com/ptompalski/teams-archive-extension/releases) and shows a **View release** banner if its version is newer than the installed extension. Checks are cached locally for 24 hours, including failed attempts, to avoid repeated requests. The check only requests public release metadata from GitHub; no Teams messages or archive data are sent.

The small version number beside the archiver title shows the installed version. Click it to check for updates immediately, bypassing the daily cache. The notice shows **Checking for updates...**, then either a newer release, confirmation that no newer release is available, or an error. Only the version button is disabled while checking; archiving is unaffected.

The update check does not block archiving. If GitHub is unavailable, rate-limited, or has no published release, a separate notice explains that updates could not be checked. Archiving remains available. There are no background checks or automatic installations.

To update an unpacked installation:

1. Wait for any archive run to finish, then close the archiver window.
2. Follow **View release** and download the release ZIP.
3. Extract it and replace the extension's files in the same folder originally selected with **Load unpacked**. If using git, pull the released version into that folder instead.
4. Open `chrome://extensions` and click **Reload** on Teams Chat Archiver.
5. Refresh your Teams web tab before opening the archiver again.

For maintainers: publish a non-draft, non-prerelease GitHub release with a numeric version tag, such as `v0.1.6` or `0.1.6`, matching `version` in `manifest.json`. Versions can have one to four numeric components, each from 0 to 65535. Merely changing the version on the default branch does not trigger notifications.

## Development checks

Serve the repository root with a local static HTTP server, then open `/tests/group-archive.html` in a browser. The dependency-free regression page tests the actual content, popup, and background scripts against synthetic Teams markup and mocked extension APIs. It covers group boundaries, muted chats, channels, deduplication, expansion, pagination, virtualized rows, channel metadata and thread navigation, bidirectional channel history, error handling, transcript output, persistent-window routing, focus changes, incoming-message scroll resets, incomplete-history diagnostics, detailed progress through saving and verification, and cached and manual update notifications. File-writing tests use in-memory directory handles, and release checks use mocked GitHub responses; the page does not access Teams, contact GitHub, or write archives to disk. Browser APIs are mocked, so live extension-window lifetime and Teams background-loading behavior still require manual validation in Chrome.
