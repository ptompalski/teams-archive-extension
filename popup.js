const archiveButton = document.getElementById("archive-chat-button");
const archiveVisibleButton = document.getElementById("archive-visible-button");
const statusElement = document.getElementById("status");
const includeImagesCheckbox = document.getElementById("include-images");
const runIndicatorElement = document.getElementById("run-indicator");
const runLabelElement = document.getElementById("run-label");
let progressListenerRegistered = false;

initializePopup().catch((error) => {
  setStatus(error.message || "Could not load settings.", true);
  setRunState("error");
});

function setStatus(message, isError = false) {
  statusElement.textContent = message;
  statusElement.style.color = isError ? "#b42318" : "#344054";
}

function setRunState(state) {
  runIndicatorElement.classList.remove("is-idle", "is-running", "is-finished", "is-error");

  const nextState =
    state === "running" || state === "finished" || state === "error" ? state : "idle";

  runIndicatorElement.classList.add(`is-${nextState}`);

  const labels = {
    idle: "Idle",
    running: "Running",
    finished: "Finished",
    error: "Stopped with error"
  };

  runLabelElement.textContent = labels[nextState];
}

async function initializePopup() {
  const settings = await loadSettings();
  includeImagesCheckbox.checked = settings.includeImages;

  includeImagesCheckbox.addEventListener("change", async () => {
    await saveSettings(getSettingsFromForm());
  });
}

function getSettingsFromForm() {
  return {
    includeImages: includeImagesCheckbox.checked
  };
}

async function loadSettings() {
  const defaults = {
    includeImages: true
  };

  if (!chrome.storage?.local) {
    return defaults;
  }

  const stored = await chrome.storage.local.get(defaults);

  return {
    includeImages: stored.includeImages ?? defaults.includeImages
  };
}

async function saveSettings(settings) {
  if (!chrome.storage?.local) {
    return;
  }

  await chrome.storage.local.set(settings);
}

async function getActiveTeamsTab() {
  const tabs = await chrome.tabs.query({
    active: true,
    currentWindow: true
  });

  const activeTab = tabs[0];

  if (!activeTab || !activeTab.id || !activeTab.url) {
    throw new Error("No active browser tab found.");
  }

  if (!activeTab.url.startsWith("https://teams.microsoft.com/")) {
    throw new Error("Open Microsoft Teams on the web, then try again.");
  }

  return activeTab;
}

async function ensureContentScriptLoaded(tabId) {
  try {
    await chrome.tabs.sendMessage(tabId, {
      type: "PING_TEAMS_ARCHIVER"
    });
  } catch (error) {
    await chrome.scripting.executeScript({
      target: { tabId },
      files: ["content.js"]
    });
  }
}

function ensureProgressListener() {
  if (progressListenerRegistered) {
    return;
  }

  chrome.runtime.onMessage.addListener((message) => {
    if (message?.type === "ARCHIVE_PROGRESS" && message.status) {
      setRunState("running");
      setStatus(message.status);
    }
  });

  progressListenerRegistered = true;
}

archiveButton.addEventListener("click", async () => {
  await runArchiveFlow({
    mode: "current"
  });
});

archiveVisibleButton.addEventListener("click", async () => {
  await runArchiveFlow({
    mode: "visible"
  });
});

async function runArchiveFlow({ mode }) {
  setBusy(true);
  ensureProgressListener();
  setRunState("running");
  setStatus("Choose an archive folder...");

  try {
    const settings = getSettingsFromForm();
    await saveSettings(settings);

    const rootDirectory = await window.showDirectoryPicker({
      mode: "readwrite"
    });

    const activeTab = await getActiveTeamsTab();
    await ensureContentScriptLoaded(activeTab.id);

    if (mode === "current") {
      setStatus("Collecting messages from the current Teams chat...");
      const validation = await archiveCurrentChatToFolder(activeTab.id, rootDirectory, settings);
      await rebuildArchiveRootIndex(rootDirectory);
      await validateRootFiles(rootDirectory);
      setRunState("finished");
      setStatus(
        `Finished. Verified ${validation.chatTitle || "chat archive"}, index, and search files.`
      );
      return;
    }

    const batchResult = await archiveVisibleChatsToFolder(activeTab.id, rootDirectory, settings);
    await rebuildArchiveRootIndex(rootDirectory);
    await validateRootFiles(rootDirectory);
    setRunState("finished");
    const skippedSuffix = batchResult.skippedCount
      ? ` Skipped ${batchResult.skippedCount} already archived or unreadable chats.`
      : "";
    setStatus(
      `Finished. Verified ${batchResult.archivedCount} chat archives plus the index and search files.${skippedSuffix}`
    );
  } catch (error) {
    const wasCancelled = error?.name === "AbortError";
    setRunState(wasCancelled ? "idle" : "error");
    setStatus(
      wasCancelled
        ? "Archive cancelled."
        : error.message || "Unexpected error while archiving chats.",
      !wasCancelled
    );
  } finally {
    setBusy(false);
  }
}

function setBusy(isBusy) {
  archiveButton.disabled = isBusy;
  archiveVisibleButton.disabled = isBusy;
  includeImagesCheckbox.disabled = isBusy;
}

async function archiveVisibleChatsToFolder(tabId, rootDirectory, settings) {
  setStatus("Reading visible Teams chats...");

  const listResponse = await chrome.tabs.sendMessage(tabId, {
    type: "LIST_VISIBLE_CHATS"
  });

  if (!listResponse || !listResponse.ok) {
    throw new Error(listResponse?.error || "Could not read the visible Teams chat list.");
  }

  const chats = listResponse.payload || [];

  if (!chats.length) {
    throw new Error("No chats were found in the current Teams view.");
  }

  let archivedCount = 0;
  let skippedCount = 0;
  const validatedChats = [];

  for (let index = 0; index < chats.length; index += 1) {
    const chat = chats[index];
    const skipReason = await getExistingArchiveSkipReason(rootDirectory, chat.label);

    if (skipReason) {
      skippedCount += 1;
      setStatus(`Skipped "${chat.label}": ${skipReason}`);
      await wait(300);
      continue;
    }

    setStatus(`Opening chat ${index + 1} of ${chats.length}: ${chat.label}`);

    const openResponse = await chrome.tabs.sendMessage(tabId, {
      type: "OPEN_CHAT_BY_LABEL",
      chatLabel: chat.label
    });

    if (!openResponse || !openResponse.ok) {
      setStatus(`Skipped "${chat.label}": ${openResponse?.error || "could not open chat"}`, true);
      await wait(1200);
      continue;
    }

    setStatus(`Archiving chat ${index + 1} of ${chats.length}: ${chat.label}`);
    try {
      const validation = await archiveCurrentChatToFolder(tabId, rootDirectory, settings, chat.label);
      archivedCount += 1;
      validatedChats.push(validation);
    } catch (error) {
      if (shouldSkipChatError(error)) {
        skippedCount += 1;
        setStatus(`Skipped "${chat.label}": ${error.message}`, true);
        await wait(1200);
        continue;
      }

      throw error;
    }

    await wait(900);
  }

  if (skippedCount > 0) {
    setStatus(`Archived ${archivedCount} chats and skipped ${skippedCount}.`);
    return {
      archivedCount,
      skippedCount,
      validatedChats
    };
  }

  setStatus(`Archived ${archivedCount} chats.`);
  return {
    archivedCount,
    skippedCount,
    validatedChats
  };
}

async function archiveCurrentChatToFolder(tabId, rootDirectory, settings, chatLabelOverride = "") {
  const response = await chrome.tabs.sendMessage(tabId, {
    type: "ARCHIVE_CURRENT_CHAT",
    chatLabelOverride
  });

  if (!response || !response.ok) {
    throw new Error(response?.error || "The Teams page did not return any data.");
  }

  const snapshot = response.payload;
  const mergedArchive = await rebuildMergedArchive(snapshot);
  const buildResponse = await chrome.runtime.sendMessage({
    type: "BUILD_ARCHIVE_FILES",
    snapshot,
    mergedArchive,
    settings
  });

  if (!buildResponse || !buildResponse.ok) {
    throw new Error(buildResponse?.error || "Could not prepare archive files.");
  }

  return writeArchiveFiles(rootDirectory, buildResponse, snapshot, mergedArchive, settings);
}

function shouldSkipChatError(error) {
  const message = String(error?.message || "");

  return (
    message.includes("No messages were found") ||
    message.includes("The Teams page did not return any data")
  );
}

async function rebuildMergedArchive(snapshot) {
  const mergedMessages = dedupeAndSortMessages(snapshot.messages || []);

  return {
    exportedAt: snapshot.exportedAt,
    pageTitle: snapshot.pageTitle,
    pageUrl: snapshot.pageUrl,
    chatTitle: snapshot.chatTitle,
    chatOrder: Number.isFinite(Number(snapshot.chatOrder)) ? Number(snapshot.chatOrder) : -1,
    messageCount: mergedMessages.length,
    messages: mergedMessages
  };
}

async function writeArchiveFiles(rootDirectory, buildResponse, snapshot, mergedArchive, settings) {
  const chatDirectory = await getOrCreateDirectory(rootDirectory, buildResponse.folderName);
  const snapshotsDirectory = await getOrCreateDirectory(chatDirectory, "snapshots");

  await writeTextFile(
    snapshotsDirectory,
    buildResponse.snapshotFilename,
    JSON.stringify(snapshot, null, 2)
  );

  const mergedFromSnapshots = await rebuildFromSnapshots(snapshotsDirectory, mergedArchive);
  const rebuiltResponse = await chrome.runtime.sendMessage({
    type: "BUILD_ARCHIVE_FILES",
    snapshot,
    mergedArchive: mergedFromSnapshots,
    settings
  });

  if (!rebuiltResponse || !rebuiltResponse.ok) {
    throw new Error(rebuiltResponse?.error || "Could not rebuild latest archive.");
  }

  const archiveWithImages = settings.includeImages
    ? await materializeInlineImages(chatDirectory, mergedFromSnapshots)
    : await stripInlineImages(chatDirectory, mergedFromSnapshots);
  const finalBuildResponse = await chrome.runtime.sendMessage({
    type: "BUILD_ARCHIVE_FILES",
    snapshot,
    mergedArchive: archiveWithImages,
    settings
  });

  if (!finalBuildResponse || !finalBuildResponse.ok) {
    throw new Error(finalBuildResponse?.error || "Could not build archive files with images.");
  }

  await writeTextFile(chatDirectory, "latest.json", finalBuildResponse.latestJsonText);
  await writeTextFile(chatDirectory, "manifest.json", finalBuildResponse.manifestText);

  await writeTextFile(chatDirectory, "latest.html", finalBuildResponse.latestHtmlText);
  await syncYearlyHtmlFiles(chatDirectory, finalBuildResponse.yearlyHtmlFiles || []);
  await deleteDirectoryContentsIfExists(chatDirectory, "md");

  await validateChatArchiveFiles(
    chatDirectory,
    buildResponse.snapshotFilename,
    finalBuildResponse.yearlyHtmlFiles || []
  );

  return {
    folderName: buildResponse.folderName,
    chatTitle: snapshot.chatTitle || mergedArchive.chatTitle || buildResponse.folderName
  };
}

async function validateChatArchiveFiles(chatDirectory, snapshotFilename, yearlyHtmlFiles) {
  await validateTextFile(chatDirectory, "latest.json");
  await validateTextFile(chatDirectory, "manifest.json");
  await validateTextFile(chatDirectory, "latest.html");

  const snapshotsDirectory = await chatDirectory.getDirectoryHandle("snapshots");
  await validateTextFile(snapshotsDirectory, snapshotFilename);

  if (Array.isArray(yearlyHtmlFiles) && yearlyHtmlFiles.length) {
    const htmlDirectory = await chatDirectory.getDirectoryHandle("html");

    for (const file of yearlyHtmlFiles) {
      if (!file?.filename) {
        continue;
      }

      await validateTextFile(htmlDirectory, file.filename);
    }
  }
}

async function validateRootFiles(rootDirectory) {
  await validateTextFile(rootDirectory, "index.html");
  await validateTextFile(rootDirectory, "search.html");
}

async function validateTextFile(directoryHandle, filename) {
  const fileHandle = await directoryHandle.getFileHandle(filename);
  const file = await fileHandle.getFile();

  if (!file || file.size <= 0) {
    throw new Error(`Validation failed: ${filename} was not written correctly.`);
  }

  return file;
}

async function rebuildArchiveRootIndex(rootDirectory) {
  const entries = [];
  const searchRecords = [];

  for await (const entry of rootDirectory.values()) {
    if (entry.kind !== "directory") {
      continue;
    }

    const manifest = await readJsonFileIfExists(entry, "manifest.json");

    if (!manifest) {
      continue;
    }

    const latestArchive = await readJsonFileIfExists(entry, "latest.json");
    const participants = latestArchive?.messages?.length
      ? listArchiveParticipants(latestArchive.messages)
      : Array.isArray(manifest.participants)
        ? manifest.participants.map((name) => cleanDisplayName(name)).filter(Boolean)
        : [];

    entries.push({
      folderName: entry.name,
      chatTitle: String(manifest.chatTitle || manifest.archiveLabel || entry.name),
      archiveLabel: String(manifest.archiveLabel || manifest.chatTitle || entry.name),
      chatOrder: Number.isFinite(Number(manifest.chatOrder)) ? Number(manifest.chatOrder) : -1,
      participants,
      lastArchivedAt: String(manifest.lastArchivedAt || ""),
      latestMessageCount: Number(manifest.latestMessageCount || 0),
      firstMessageAt: String(manifest.firstMessageAt || ""),
      lastMessageAt: String(manifest.lastMessageAt || ""),
      newestSnapshot: String(manifest.newestSnapshot || ""),
      readableFormat: "html"
    });

    if (latestArchive?.messages?.length) {
      searchRecords.push(
        ...buildSearchRecords({
          folderName: entry.name,
          chatTitle: String(manifest.chatTitle || manifest.archiveLabel || entry.name),
          messages: latestArchive.messages
        })
      );
    }
  }

  entries.sort((left, right) => {
    const leftOrder = Number.isFinite(left.chatOrder) ? left.chatOrder : -1;
    const rightOrder = Number.isFinite(right.chatOrder) ? right.chatOrder : -1;
    const leftHasOrder = leftOrder >= 0;
    const rightHasOrder = rightOrder >= 0;

    if (leftHasOrder && rightHasOrder && leftOrder !== rightOrder) {
      return leftOrder - rightOrder;
    }

    if (leftHasOrder !== rightHasOrder) {
      return leftHasOrder ? -1 : 1;
    }

    const leftTime = Date.parse(left.lastArchivedAt || "") || 0;
    const rightTime = Date.parse(right.lastArchivedAt || "") || 0;

    if (leftTime !== rightTime) {
      return rightTime - leftTime;
    }

    return left.chatTitle.localeCompare(right.chatTitle);
  });

  await writeTextFile(rootDirectory, "index.html", buildArchiveRootIndexHtml(entries));
  await writeTextFile(rootDirectory, "search.html", buildArchiveSearchHtml(searchRecords));
}

async function readJsonFileIfExists(directoryHandle, filename) {
  try {
    const fileHandle = await directoryHandle.getFileHandle(filename);
    const file = await fileHandle.getFile();
    return JSON.parse(await file.text());
  } catch (error) {
    return null;
  }
}

async function getExistingArchiveSkipReason(rootDirectory, chatLabel) {
  const folderName = buildArchiveFolderNameFromChatLabel(chatLabel);

  if (!folderName) {
    return "";
  }

  let chatDirectory;

  try {
    chatDirectory = await rootDirectory.getDirectoryHandle(folderName);
  } catch (error) {
    if (error?.name === "NotFoundError") {
      return "";
    }

    throw error;
  }

  const today = new Date();
  const todayPrefix = formatDateForFile(today);

  try {
    const snapshotsDirectory = await chatDirectory.getDirectoryHandle("snapshots");

    for await (const entry of snapshotsDirectory.values()) {
      if (
        entry.kind === "file" &&
        entry.name.startsWith(`${todayPrefix}T`) &&
        entry.name.endsWith(".json")
      ) {
        return "today's snapshot already exists";
      }
    }
  } catch (error) {
    if (error?.name !== "NotFoundError") {
      throw error;
    }
  }

  const manifest = await readJsonFileIfExists(chatDirectory, "manifest.json");

  if (isSameLocalDate(manifest?.lastArchivedAt, today)) {
    return "manifest was already archived today";
  }

  return "";
}

function buildArchiveFolderNameFromChatLabel(chatLabel) {
  return sanitizeFilePart(cleanDisplayName(chatLabel)) || "teams-chat";
}

function isSameLocalDate(value, date) {
  const parsed = new Date(value || "");

  if (Number.isNaN(parsed.getTime())) {
    return false;
  }

  return formatDateForFile(parsed) === formatDateForFile(date);
}

function formatDateForFile(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}
function buildArchiveRootIndexHtml(entries) {
  const rowsHtml = entries
    .map((entry) => {
      const latestLink = `${encodePathPart(entry.folderName)}/latest.html`;
      const chatOrder = Number.isFinite(entry.chatOrder) ? entry.chatOrder : -1;
      const lastArchivedValue = Date.parse(entry.lastArchivedAt || "") || 0;
      const firstMessageValue = Date.parse(entry.firstMessageAt || "") || 0;
      const lastMessageValue = Date.parse(entry.lastMessageAt || "") || 0;
      const participantsLabel = (entry.participants || []).join(", ");

      return `
        <tr
          data-chat-order="${escapeHtml(String(chatOrder))}"
          data-chat-title="${escapeHtml(String(entry.chatTitle || "").toLowerCase())}"
          data-last-archived="${escapeHtml(String(lastArchivedValue))}"
          data-messages="${escapeHtml(String(entry.latestMessageCount || 0))}"
          data-range-start="${escapeHtml(String(firstMessageValue))}"
          data-range-end="${escapeHtml(String(lastMessageValue))}"
        >
          <td class="chat-name"><a href="${latestLink}">${escapeHtml(entry.chatTitle)}</a></td>
          <td>${escapeHtml(participantsLabel || "Unknown")}</td>
          <td>${escapeHtml(formatDisplayDate(entry.lastArchivedAt))}</td>
          <td>${escapeHtml(String(entry.latestMessageCount))}</td>
          <td>${escapeHtml(formatRange(entry.firstMessageAt, entry.lastMessageAt))}</td>
        </tr>
      `;
    })
    .join("\n");

  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Teams Archive Index</title>
  <style>
    :root {
      color-scheme: light;
      --bg: #eef4f9;
      --panel: #ffffff;
      --panel-border: #d7e2ec;
      --text: #172026;
      --muted: #52606d;
      --accent: #005a9c;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: "Segoe UI", Arial, sans-serif;
      background: linear-gradient(180deg, var(--bg) 0%, #f8fbfd 100%);
      color: var(--text);
    }
    main {
      max-width: 1100px;
      margin: 0 auto;
      padding: 32px 20px 48px;
    }
    h1 {
      margin: 0 0 8px;
      font-size: 32px;
      line-height: 1.1;
    }
    .intro {
      margin: 0 0 24px;
      color: var(--muted);
      font-size: 15px;
    }
    .toolbar {
      display: flex;
      align-items: center;
      gap: 12px;
      margin: 0 0 18px;
      flex-wrap: wrap;
    }
    .toolbar label {
      font-size: 13px;
      font-weight: 600;
      color: var(--muted);
    }
    .sort-select {
      min-width: 220px;
      padding: 10px 12px;
      border: 1px solid var(--panel-border);
      border-radius: 10px;
      background: #ffffff;
      color: var(--text);
      font-size: 14px;
    }
    .archive-list {
      background: var(--panel);
      border: 1px solid var(--panel-border);
      border-radius: 14px;
      box-shadow: 0 10px 30px rgba(9, 30, 66, 0.06);
      overflow: hidden;
    }
    table {
      width: 100%;
      border-collapse: collapse;
    }
    th,
    td {
      padding: 14px 16px;
      text-align: left;
      vertical-align: top;
      border-bottom: 1px solid var(--panel-border);
      font-size: 14px;
    }
    th {
      background: #f6f9fc;
      color: var(--muted);
      font-size: 12px;
      font-weight: 700;
      letter-spacing: 0.02em;
      text-transform: uppercase;
    }
    tbody tr:last-child td {
      border-bottom: 0;
    }
    td a {
      color: var(--accent);
      text-decoration: none;
    }
    td a:hover {
      text-decoration: underline;
    }
    .chat-name {
      font-weight: 700;
      min-width: 220px;
    }
    .empty-state {
      padding: 20px;
      color: var(--muted);
      font-size: 14px;
    }
    @media (max-width: 840px) {
      .archive-list {
        overflow-x: auto;
      }
      table {
        min-width: 760px;
      }
    }
  </style>
</head>
<body>
  <main>
    <h1>Teams Archive</h1>
    <p class="intro">Browse archived chats and open each chat overview page. Need message-level search? Open <a href="search.html">archive search</a>.</p>
    <div class="toolbar">
      <label for="sort-chats">Sort by</label>
      <select id="sort-chats" class="sort-select">
        <option value="teams">Teams order</option>
        <option value="name">Chat name</option>
        <option value="archived">Last archived</option>
        <option value="messages">Message count</option>
        <option value="range">Chat range</option>
      </select>
    </div>
    <section class="archive-list">
      ${
        rowsHtml
          ? `<table>
      <thead>
        <tr>
          <th>Chat</th>
          <th>Chat Users</th>
          <th>Last Archived</th>
          <th>Messages</th>
          <th>Chat Range</th>
        </tr>
      </thead>
      <tbody>
        ${rowsHtml}
      </tbody>
    </table>`
          : `<div class="empty-state">No archived chats found yet.</div>`
      }
    </section>
  </main>
  <script>
    (() => {
      const sortSelect = document.getElementById("sort-chats");
      const tbody = document.querySelector("tbody");

      if (!sortSelect || !tbody) {
        return;
      }

      const getNumber = (row, key) => Number(row.getAttribute(key) || 0);
      const getText = (row, key) => String(row.getAttribute(key) || "");

      const compareRows = (left, right, mode) => {
        if (mode === "name") {
          return getText(left, "data-chat-title").localeCompare(getText(right, "data-chat-title"));
        }

        if (mode === "archived") {
          const diff = getNumber(right, "data-last-archived") - getNumber(left, "data-last-archived");
          return diff || getText(left, "data-chat-title").localeCompare(getText(right, "data-chat-title"));
        }

        if (mode === "messages") {
          const diff = getNumber(right, "data-messages") - getNumber(left, "data-messages");
          return diff || getText(left, "data-chat-title").localeCompare(getText(right, "data-chat-title"));
        }

        if (mode === "range") {
          const diff = getNumber(right, "data-range-end") - getNumber(left, "data-range-end");
          return diff || getNumber(right, "data-range-start") - getNumber(left, "data-range-start");
        }

        const leftOrder = getNumber(left, "data-chat-order");
        const rightOrder = getNumber(right, "data-chat-order");
        const leftHasOrder = leftOrder >= 0;
        const rightHasOrder = rightOrder >= 0;

        if (leftHasOrder && rightHasOrder && leftOrder !== rightOrder) {
          return leftOrder - rightOrder;
        }

        if (leftHasOrder !== rightHasOrder) {
          return leftHasOrder ? -1 : 1;
        }

        return getText(left, "data-chat-title").localeCompare(getText(right, "data-chat-title"));
      };

      const applySort = () => {
        const rows = Array.from(tbody.querySelectorAll("tr"));
        const mode = sortSelect.value || "teams";
        rows.sort((left, right) => compareRows(left, right, mode));

        for (const row of rows) {
          tbody.appendChild(row);
        }
      };

      sortSelect.addEventListener("change", applySort);
      applySort();
    })();
  </script>
</body>
</html>`;
}

function buildSearchRecords({ folderName, chatTitle, messages }) {
  return (messages || [])
    .filter((message) => (message?.text || "").trim())
    .map((message) => {
      const timestamp = String(message.timestamp || "");
      const yearLink = buildMessageYearlyLink(folderName, timestamp);

      return {
        chatTitle,
        author: cleanDisplayName(String(message.author || "Unknown")),
        timestamp,
        text: String(message.text || "").replace(/\s+/g, " ").trim(),
        link: yearLink,
        searchText: [chatTitle, cleanDisplayName(message.author || ""), timestamp, message.text || ""]
          .join(" ")
          .toLowerCase()
      };
    })
    .sort((left, right) => {
      const leftTime = Date.parse(left.timestamp || "") || 0;
      const rightTime = Date.parse(right.timestamp || "") || 0;
      return rightTime - leftTime;
    });
}

function buildMessageYearlyLink(folderName, timestamp) {
  const parsed = parseTimestampValue(timestamp);

  if (!parsed) {
    return `${encodePathPart(folderName)}/latest.html`;
  }

  return `${encodePathPart(folderName)}/html/${parsed.getFullYear()}.html`;
}

function buildArchiveSearchHtml(records) {
  const payload = serializeForInlineScript(records);

  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Teams Archive Search</title>
  <style>
    :root {
      color-scheme: light;
      --bg: #eef4f9;
      --panel: #ffffff;
      --panel-border: #d7e2ec;
      --text: #172026;
      --muted: #52606d;
      --accent: #005a9c;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: "Segoe UI", Arial, sans-serif;
      background: linear-gradient(180deg, var(--bg) 0%, #f8fbfd 100%);
      color: var(--text);
    }
    main {
      max-width: 1100px;
      margin: 0 auto;
      padding: 32px 20px 48px;
    }
    h1 {
      margin: 0 0 8px;
      font-size: 32px;
      line-height: 1.1;
    }
    .intro {
      margin: 0 0 18px;
      color: var(--muted);
      font-size: 15px;
    }
    .search-input {
      width: 100%;
      padding: 12px 14px;
      border: 1px solid var(--panel-border);
      border-radius: 10px;
      background: #ffffff;
      color: var(--text);
      font-size: 14px;
      margin-bottom: 18px;
    }
    .results {
      display: grid;
      gap: 12px;
    }
    .result {
      background: var(--panel);
      border: 1px solid var(--panel-border);
      border-radius: 12px;
      padding: 16px 18px;
      box-shadow: 0 10px 30px rgba(9, 30, 66, 0.06);
    }
    .result-head {
      display: flex;
      justify-content: space-between;
      gap: 12px;
      margin-bottom: 8px;
      flex-wrap: wrap;
    }
    .chat-link {
      color: var(--accent);
      text-decoration: none;
      font-weight: 700;
    }
    .chat-link:hover {
      text-decoration: underline;
    }
    .meta {
      color: var(--muted);
      font-size: 13px;
    }
    .snippet {
      margin: 0;
      white-space: pre-wrap;
      word-break: break-word;
      line-height: 1.5;
      font-size: 14px;
    }
    .empty {
      display: none;
      background: var(--panel);
      border: 1px solid var(--panel-border);
      border-radius: 12px;
      padding: 18px;
      color: var(--muted);
      font-size: 14px;
    }
  </style>
</head>
<body>
  <main>
    <p class="intro"><a href="index.html">Back to main index</a></p>
    <h1>Archive Search</h1>
    <p class="intro">Search archived messages across all chats by text, author, chat name, or timestamp.</p>
    <input id="message-search" class="search-input" type="search" placeholder="Search all archived messages..." autocomplete="off">
    <div id="result-count" class="intro"></div>
    <section id="results" class="results"></section>
    <div id="empty" class="empty">No archived messages match your search.</div>
  </main>
  <script>
    (() => {
      const records = ${payload};
      const input = document.getElementById("message-search");
      const results = document.getElementById("results");
      const empty = document.getElementById("empty");
      const resultCount = document.getElementById("result-count");
      const limit = 250;

      const escapeHtml = (value) => String(value)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#39;");

      const render = () => {
        const query = String(input.value || "").trim().toLowerCase();
        const matches = !query
          ? records.slice(0, limit)
          : records.filter((record) => record.searchText.includes(query)).slice(0, limit);

        resultCount.textContent = query
          ? 'Showing ' + matches.length + ' matching messages' + (matches.length === limit ? ' (limited)' : '')
          : 'Showing latest searchable messages';

        results.innerHTML = matches.map((record) => {
          return '<article class="result">' +
            '<div class="result-head">' +
              '<a class="chat-link" href="' + escapeHtml(record.link) + '">' + escapeHtml(record.chatTitle) + '</a>' +
              '<span class="meta">' + escapeHtml(record.timestamp || 'Unknown time') + '</span>' +
            '</div>' +
            '<div class="meta">' + escapeHtml(record.author || 'Unknown') + '</div>' +
            '<p class="snippet">' + escapeHtml(record.text || '') + '</p>' +
          '</article>';
        }).join('');

        empty.style.display = matches.length === 0 ? 'block' : 'none';
      };

      input.addEventListener('input', render);
      render();
    })();
  </script>
</body>
</html>`;
}

function serializeForInlineScript(value) {
  return JSON.stringify(value)
    .replace(/</g, "\\u003c")
    .replace(/>/g, "\\u003e")
    .replace(/&/g, "\\u0026");
}

function listArchiveParticipants(messages) {
  const participants = [];
  const seen = new Set();

  for (const message of messages || []) {
    const cleaned = cleanDisplayName(message?.author || "");

    if (
      !cleaned ||
      cleaned === "Unknown" ||
      seen.has(cleaned) ||
      !looksLikePersonDisplayName(cleaned)
    ) {
      continue;
    }

    seen.add(cleaned);
    participants.push(cleaned);
  }

  return participants;
}

function cleanDisplayName(value) {
  const cleaned = String(value || "")
    .replace(/\s*\([^)]*\)\s*/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  const commaMatch = cleaned.match(/^([^,]+),\s*(.+)$/);

  if (commaMatch) {
    return `${commaMatch[2]} ${commaMatch[1]}`.replace(/\s+/g, " ").trim();
  }

  return cleaned;
}

function looksLikePersonDisplayName(value) {
  const cleaned = String(value || "").trim();

  if (!cleaned || cleaned.length > 80) {
    return false;
  }

  if (/[.!?]/.test(cleaned)) {
    return false;
  }

  const parts = cleaned.split(/\s+/).filter(Boolean);

  if (parts.length < 1 || parts.length > 5) {
    return false;
  }

  return parts.every((part) => /^(?:[A-Z][\p{L}\p{M}'-]*|[A-Z]\.)$/u.test(part));
}

async function syncYearlyHtmlFiles(chatDirectory, yearlyHtmlFiles) {
  const htmlDirectory = await getOrCreateDirectory(chatDirectory, "html");
  const wanted = new Set(yearlyHtmlFiles.map((file) => file.filename));

  for (const file of yearlyHtmlFiles) {
    await writeTextFile(htmlDirectory, file.filename, file.content);
  }

  for await (const entry of htmlDirectory.values()) {
    if (entry.kind !== "file") {
      continue;
    }

    if (!wanted.has(entry.name)) {
      await htmlDirectory.removeEntry(entry.name);
    }
  }
}

async function rebuildFromSnapshots(snapshotsDirectory, fallbackArchive) {
  const snapshots = [];

  for await (const entry of snapshotsDirectory.values()) {
    if (entry.kind !== "file" || !entry.name.endsWith(".json")) {
      continue;
    }

    const file = await entry.getFile();
    const text = await file.text();

    try {
      snapshots.push(JSON.parse(text));
    } catch (error) {
      // Ignore malformed snapshot files so the archive remains usable.
    }
  }

  if (!snapshots.length) {
    return fallbackArchive;
  }

  snapshots.sort((left, right) => {
    const leftTime = Date.parse(left.exportedAt || "") || 0;
    const rightTime = Date.parse(right.exportedAt || "") || 0;
    return leftTime - rightTime;
  });

  const mergedMessages = dedupeAndSortMessages(
    snapshots.flatMap((snapshot) => snapshot.messages || [])
  );
  const lastSnapshot = snapshots[snapshots.length - 1];

  return {
    exportedAt: lastSnapshot.exportedAt || fallbackArchive.exportedAt,
    pageTitle: lastSnapshot.pageTitle || fallbackArchive.pageTitle,
    pageUrl: lastSnapshot.pageUrl || fallbackArchive.pageUrl,
    chatTitle: lastSnapshot.chatTitle || fallbackArchive.chatTitle,
    chatOrder:
      Number.isFinite(Number(lastSnapshot.chatOrder)) && Number(lastSnapshot.chatOrder) >= 0
        ? Number(lastSnapshot.chatOrder)
        : Number.isFinite(Number(fallbackArchive.chatOrder))
          ? Number(fallbackArchive.chatOrder)
          : -1,
    messageCount: mergedMessages.length,
    messages: mergedMessages
  };
}

function dedupeAndSortMessages(messages) {
  const seen = new Set();
  const deduped = [];

  for (const message of messages) {
    const normalized = normalizeMessage(message);
    const key = buildMessageKey(normalized);

    if (seen.has(key)) {
      continue;
    }

    seen.add(key);
    deduped.push(normalized);
  }

  return deduped.sort(compareMessages);
}

function normalizeMessage(message) {
  return {
    id: String(message?.id || "").trim(),
    author: cleanDisplayName(String(message?.author || "").trim()) || "Unknown",
    timestamp: String(message?.timestamp || "").trim(),
    text: String(message?.text || "").replace(/\r/g, "").trim(),
    images: normalizeImages(message?.images)
  };
}

function normalizeImages(images) {
  if (!Array.isArray(images)) {
    return [];
  }

  const seen = new Set();
  const normalized = [];

  for (const image of images) {
    const sourceUrl = String(image?.sourceUrl || "").trim();
    const localPath = String(image?.localPath || "").trim();
    const alt = String(image?.alt || "").trim();
    const embeddedDataUrl = String(image?.embeddedDataUrl || "").trim();
    const key = localPath || sourceUrl;

    if (!key || seen.has(key)) {
      continue;
    }

    seen.add(key);
    normalized.push({
      sourceUrl,
      localPath,
      alt,
      embeddedDataUrl
    });
  }

  return normalized;
}

function buildMessageKey(message) {
  if (message.id) {
    return `id:${message.id}`;
  }

  return [
    message.author,
    normalizeTimestampKey(message.timestamp),
    message.text
  ].join("|");
}

function compareMessages(left, right) {
  const leftTime = parseTimestampValue(left.timestamp);
  const rightTime = parseTimestampValue(right.timestamp);

  if (leftTime && rightTime) {
    const diff = leftTime.getTime() - rightTime.getTime();

    if (diff !== 0) {
      return diff;
    }
  }

  if (left.author !== right.author) {
    return left.author.localeCompare(right.author);
  }

  return left.text.localeCompare(right.text);
}

function normalizeTimestampKey(value) {
  const parsed = parseTimestampValue(value);
  return parsed ? parsed.toISOString() : String(value || "").trim();
}

function parseTimestampValue(value) {
  if (!value || typeof value !== "string") {
    return null;
  }

  const normalized = value.trim();

  if (!normalized) {
    return null;
  }

  const direct = new Date(normalized);

  if (!Number.isNaN(direct.getTime())) {
    return direct;
  }

  const todayMatch = normalized.match(/^today\s+at\s+(.+)$/i);
  if (todayMatch) {
    return parseRelativeTime(todayMatch[1], 0);
  }

  const yesterdayMatch = normalized.match(/^yesterday\s+at\s+(.+)$/i);
  if (yesterdayMatch) {
    return parseRelativeTime(yesterdayMatch[1], -1);
  }

  return null;
}

function parseRelativeTime(timeText, dayOffset) {
  const timeMatch = timeText.trim().match(/^(\d{1,2}):(\d{2})(?:\s*(AM|PM))?$/i);

  if (!timeMatch) {
    return null;
  }

  let hours = Number(timeMatch[1]);
  const minutes = Number(timeMatch[2]);
  const meridiem = (timeMatch[3] || "").toUpperCase();

  if (meridiem === "PM" && hours < 12) {
    hours += 12;
  }

  if (meridiem === "AM" && hours === 12) {
    hours = 0;
  }

  const date = new Date();
  date.setDate(date.getDate() + dayOffset);
  date.setHours(hours, minutes, 0, 0);
  return date;
}

async function getOrCreateDirectory(parentHandle, name) {
  return parentHandle.getDirectoryHandle(name, { create: true });
}

async function writeTextFile(directoryHandle, filename, contents) {
  const fileHandle = await directoryHandle.getFileHandle(filename, { create: true });
  const writable = await fileHandle.createWritable();
  await writable.write(contents);
  await writable.close();
}

async function materializeInlineImages(chatDirectory, archive) {
  const imagesDirectory = await getOrCreateDirectory(chatDirectory, "assets");
  const nestedImagesDirectory = await getOrCreateDirectory(imagesDirectory, "images");
  const wantedPaths = new Set();
  const updatedMessages = [];

  for (let messageIndex = 0; messageIndex < (archive.messages || []).length; messageIndex += 1) {
    const message = archive.messages[messageIndex];
    const updatedImages = [];

    for (let imageIndex = 0; imageIndex < (message.images || []).length; imageIndex += 1) {
      const image = message.images[imageIndex];
      const savedImage = await saveInlineImage(
        nestedImagesDirectory,
        message,
        image,
        imageIndex
      );

      if (!savedImage) {
        continue;
      }

      wantedPaths.add(savedImage.localPath);
      updatedImages.push(savedImage);
    }

    updatedMessages.push({
      ...message,
      images: updatedImages
    });
  }

  await pruneUnusedImageAssets(nestedImagesDirectory, wantedPaths);

  return {
    ...archive,
    messages: updatedMessages,
    messageCount: updatedMessages.length
  };
}

async function stripInlineImages(chatDirectory, archive) {
  const imagesDirectory = await getOrCreateDirectory(chatDirectory, "assets");
  const nestedImagesDirectory = await getOrCreateDirectory(imagesDirectory, "images");
  await pruneUnusedImageAssets(nestedImagesDirectory, new Set());

  return {
    ...archive,
    messages: (archive.messages || []).map((message) => ({
      ...message,
      images: []
    })),
    messageCount: (archive.messages || []).length
  };
}

async function saveInlineImage(imagesDirectory, message, image, imageIndex) {
  const sourceUrl = String(image?.sourceUrl || "").trim();
  const embeddedDataUrl = String(image?.embeddedDataUrl || "").trim();

  if (!sourceUrl && !embeddedDataUrl) {
    return null;
  }

  try {
    const fetched = await fetchImageBlob(sourceUrl, embeddedDataUrl);

    if (!fetched) {
      return image.localPath ? image : null;
    }

    const bucketName = buildImageBucketName(message.timestamp);
    const bucketDirectory = await getOrCreateDirectory(imagesDirectory, bucketName);
    const extension = chooseImageExtension(sourceUrl, fetched.blob.type);
    const contentHash = await computeBlobHash(fetched.blob);
    const filename = `${contentHash}.${extension}`;
    const localPath = `assets/images/${bucketName}/${filename}`;

    await writeBlobFile(bucketDirectory, filename, fetched.blob);

    return {
      sourceUrl,
      alt: String(image?.alt || "").trim(),
      localPath
    };
  } catch (error) {
    return image.localPath ? image : null;
  }
}

async function fetchImageBlob(sourceUrl, embeddedDataUrl = "") {
  if (embeddedDataUrl.startsWith("data:")) {
    const response = await fetch(embeddedDataUrl);
    return {
      blob: await response.blob()
    };
  }

  if (sourceUrl.startsWith("data:")) {
    const response = await fetch(sourceUrl);
    return {
      blob: await response.blob()
    };
  }

  if (!/^https?:/i.test(sourceUrl)) {
    return null;
  }

  const useIncludedCredentials = shouldUseCredentialsForImage(sourceUrl);
  const response = await fetch(sourceUrl, {
    credentials: useIncludedCredentials ? "include" : "omit"
  });

  if (!response.ok) {
    return null;
  }

  return {
    blob: await response.blob()
  };
}

function shouldUseCredentialsForImage(sourceUrl) {
  try {
    const parsed = new URL(sourceUrl);
    return /(^|\.)teams\.microsoft\.com$/i.test(parsed.hostname);
  } catch (error) {
    return false;
  }
}

function buildImageBucketName(timestamp) {
  const parsed = parseTimestampValue(timestamp);

  if (!parsed) {
    return "unknown";
  }

  const year = parsed.getFullYear();
  const month = String(parsed.getMonth() + 1).padStart(2, "0");
  return `${year}-${month}`;
}

function chooseImageExtension(sourceUrl, mimeType) {
  const mimeToExtension = {
    "image/jpeg": "jpg",
    "image/png": "png",
    "image/gif": "gif",
    "image/webp": "webp",
    "image/svg+xml": "svg"
  };

  if (mimeType && mimeToExtension[mimeType]) {
    return mimeToExtension[mimeType];
  }

  const match = sourceUrl.match(/\.([a-z0-9]+)(?:[?#]|$)/i);

  if (match) {
    return match[1].toLowerCase();
  }

  return "png";
}

async function writeBlobFile(directoryHandle, filename, blob) {
  const fileHandle = await directoryHandle.getFileHandle(filename, { create: true });
  const writable = await fileHandle.createWritable();
  await writable.write(blob);
  await writable.close();
}

async function computeBlobHash(blob) {
  const buffer = await blob.arrayBuffer();
  const hashBuffer = await crypto.subtle.digest("SHA-256", buffer);
  const hashBytes = Array.from(new Uint8Array(hashBuffer));
  return hashBytes.map((byte) => byte.toString(16).padStart(2, "0")).join("");
}

async function pruneUnusedImageAssets(imagesDirectory, wantedPaths) {
  const pathByBucket = new Map();

  for (const wantedPath of wantedPaths) {
    const normalized = wantedPath.replace(/\\/g, "/");
    const parts = normalized.split("/");

    if (parts.length < 4) {
      continue;
    }

    const bucketName = parts[2];
    const filename = parts.slice(3).join("/");

    if (!pathByBucket.has(bucketName)) {
      pathByBucket.set(bucketName, new Set());
    }

    pathByBucket.get(bucketName).add(filename);
  }

  for await (const bucketEntry of imagesDirectory.values()) {
    if (bucketEntry.kind !== "directory") {
      continue;
    }

    const wantedFiles = pathByBucket.get(bucketEntry.name) || new Set();

    for await (const entry of bucketEntry.values()) {
      if (entry.kind !== "file") {
        continue;
      }

      if (!wantedFiles.has(entry.name)) {
        await bucketEntry.removeEntry(entry.name);
      }
    }
  }
}

function sanitizeFilePart(value) {
  return String(value || "")
    .trim()
    .replace(/[<>:"/\\|?*]+/g, "-")
    .replace(/\s+/g, "-")
    .replace(/-+/g, "-")
    .replace(/^-|-$/g, "")
    .toLowerCase();
}

async function deleteFileIfExists(directoryHandle, filename) {
  try {
    await directoryHandle.removeEntry(filename);
  } catch (error) {
    if (error?.name !== "NotFoundError") {
      throw error;
    }
  }
}

async function deleteDirectoryContentsIfExists(directoryHandle, directoryName) {
  try {
    const nestedDirectory = await directoryHandle.getDirectoryHandle(directoryName);

    for await (const entry of nestedDirectory.values()) {
      await nestedDirectory.removeEntry(entry.name, { recursive: entry.kind === "directory" });
    }
  } catch (error) {
    if (error?.name !== "NotFoundError") {
      throw error;
    }
  }
}

function wait(ms) {
  return new Promise((resolve) => {
    window.setTimeout(resolve, ms);
  });
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function encodePathPart(value) {
  return String(value || "")
    .split("/")
    .map((part) => encodeURIComponent(part))
    .join("/");
}

function formatDisplayDate(value) {
  const parsed = Date.parse(value || "");

  if (Number.isNaN(parsed)) {
    return value || "Unknown";
  }

  return new Intl.DateTimeFormat(undefined, {
    year: "numeric",
    month: "short",
    day: "numeric",
    hour: "numeric",
    minute: "2-digit"
  }).format(new Date(parsed));
}

function formatRange(firstValue, lastValue) {
  const first = formatShortDate(firstValue);
  const last = formatShortDate(lastValue);

  if (!first && !last) {
    return "Unknown";
  }

  if (first && last && first !== last) {
    return `${first} to ${last}`;
  }

  return first || last;
}

function formatShortDate(value) {
  const parsed = Date.parse(value || "");

  if (Number.isNaN(parsed)) {
    return "";
  }

  return new Intl.DateTimeFormat(undefined, {
    year: "numeric",
    month: "short",
    day: "numeric"
  }).format(new Date(parsed));
}
