const archiveButton = document.getElementById("archive-chat-button");
const archiveVisibleButton = document.getElementById("archive-visible-button");
const archiveGroupsButton = document.getElementById("archive-groups-button");
const archiveGroupSelectionButton = document.getElementById("archive-group-selection-button");
const cancelGroupSelectionButton = document.getElementById("cancel-group-selection-button");
const groupPickerElement = document.getElementById("group-picker");
const groupListElement = document.getElementById("group-list");
const openTeamsButton = document.getElementById("open-teams-button");
const statusElement = document.getElementById("status");
const includeImagesCheckbox = document.getElementById("include-images");
const runIndicatorElement = document.getElementById("run-indicator");
const runLabelElement = document.getElementById("run-label");
const archiveProgressContainer = document.getElementById("archive-progress-container");
const archiveProgressBar = document.getElementById("archive-progress-bar");
const archiveProgressLabel = document.getElementById("archive-progress-label");
const updateNoticeElement = document.getElementById("update-notice");
const updateMessageElement = document.getElementById("update-message");
const updateReleaseLink = document.getElementById("update-release-link");
const versionButton = document.getElementById("extension-version");
let progressListenerRegistered = false;
let groupSelectionTabId = null;
let isBusy = false;
let teamsTabId = Number(new URLSearchParams(window.location?.search || "").get("tabId")) || null;
let contentProgressTabId = null;
let archiveStatusContext = "";
const TEAMS_WEB_URL = "https://teams.cloud.microsoft/";
const TEAMS_URL_PATTERNS = [
  "https://teams.cloud.microsoft/*",
  "https://teams.microsoft.com/*"
];
const RELEASES_URL = "https://github.com/ptompalski/teams-archive-extension/releases";
const RELEASE_API_URL = "https://api.github.com/repos/ptompalski/teams-archive-extension/releases/latest";
const UPDATE_CHECK_INTERVAL_MS = 24 * 60 * 60 * 1000;

initializePopup().catch((error) => {
  setStatus(error.message || "Could not load settings.", true);
  setRunState("error");
});

function setStatus(message, isError = false) {
  statusElement.textContent = archiveStatusContext ? `${archiveStatusContext}\n${message}` : message;
  statusElement.style.color = isError ? "#b42318" : "#344054";
}

function showIndeterminateArchiveProgress(message) {
  archiveProgressContainer.hidden = false;
  archiveProgressContainer.classList.remove("is-error");
  archiveProgressBar.max = 1;
  archiveProgressBar.removeAttribute("value");
  archiveProgressBar.setAttribute("aria-valuetext", message);
  archiveProgressLabel.textContent = message;
}

function updateArchiveProgress({
  totalChats,
  archivedCount = 0,
  skippedCount = 0,
  failedCount = 0,
  phase = "chats"
}) {
  const processedChats = archivedCount + skippedCount + failedCount;
  const completedSteps = processedChats + (phase === "finished" ? 1 : 0);
  const totalSteps = totalChats + 1;
  const percent = Math.floor(completedSteps * 100 / totalSteps);
  const phaseLabel = phase === "finished"
    ? failedCount > 0 ? "Finished with failed chats." : "Finished."
    : phase === "finalizing" ? "Finalizing the archive index and search files..." : "Archive finalization pending.";
  const label = `${percent}% - ${processedChats} of ${totalChats} chats processed (${archivedCount} saved, ${skippedCount} skipped, ${failedCount} failed). ${phaseLabel}`;

  archiveProgressContainer.hidden = false;
  archiveProgressContainer.classList.toggle("is-error", phase === "finished" && failedCount > 0);
  archiveProgressBar.max = totalSteps;
  archiveProgressBar.value = completedSteps;
  archiveProgressBar.setAttribute("aria-valuetext", label);
  archiveProgressLabel.textContent = label;
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
  const installedVersion = chrome.runtime.getManifest().version;
  versionButton.textContent = `v${installedVersion}`;
  versionButton.setAttribute("aria-label", `Version ${installedVersion}. Check for updates.`);
  versionButton.addEventListener("click", () => runUpdateCheck(true));

  window.addEventListener("beforeunload", (event) => {
    if (isBusy) {
      event.preventDefault();
      event.returnValue = "";
    }
  });

  const settings = await loadSettings();
  includeImagesCheckbox.checked = settings.includeImages;

  includeImagesCheckbox.addEventListener("change", async () => {
    await saveSettings(getSettingsFromForm());
  });

  openTeamsButton.addEventListener("click", async () => {
    try {
      await openTeamsWebsite();
    } catch (error) {
      setStatus(error.message || "Could not open the Teams website.", true);
    }
  });

  await runUpdateCheck();
}

async function runUpdateCheck(force = false) {
  if (versionButton.disabled) {
    return;
  }
  versionButton.disabled = true;
  versionButton.setAttribute("aria-busy", "true");
  if (force) {
    showUpdateNotice({ checking: true });
  }
  try {
    await checkForUpdates({ force });
  } catch (error) {
    console.error("Could not check for extension updates.", error);
    showUpdateNotice({ error: getErrorMessage(error, "Unexpected update-check error.") });
  } finally {
    versionButton.disabled = false;
    versionButton.setAttribute("aria-busy", "false");
  }
}

async function checkForUpdates({ force = false } = {}) {
  const now = Date.now();
  let cached;
  if (!force) {
    ({ extensionUpdateCheck: cached } = await chrome.storage.local.get("extensionUpdateCheck"));
  }
  if (cached && Number.isFinite(cached.checkedAt) && now >= cached.checkedAt &&
      now - cached.checkedAt < UPDATE_CHECK_INTERVAL_MS &&
      typeof cached.tag === "string" && typeof cached.error === "string") {
    showUpdateNotice(cached);
    return;
  }

  let result;
  try {
    const response = await fetch(RELEASE_API_URL, {
      headers: { Accept: "application/vnd.github+json" },
      credentials: "omit",
      cache: "no-store",
      signal: AbortSignal.timeout(10000)
    });
    if (!response.ok) {
      throw new Error(response.status === 404
        ? "No published release was found on GitHub (HTTP 404)."
        : `GitHub returned HTTP ${response.status}.`);
    }
    const release = await response.json();
    if (!release || release.draft || release.prerelease || typeof release.tag_name !== "string") {
      throw new Error("GitHub did not return a stable release with a version tag.");
    }
    compareExtensionVersions(release.tag_name.replace(/^v/i, ""), chrome.runtime.getManifest().version);
    result = { checkedAt: now, tag: release.tag_name, error: "" };
  } catch (error) {
    console.error("Could not check for extension updates.", error);
    result = { checkedAt: now, tag: "", error: getErrorMessage(error, "Unexpected update-check error.") };
  }

  await chrome.storage.local.set({ extensionUpdateCheck: result });
  showUpdateNotice(result, force);
}

function compareExtensionVersions(left, right) {
  const parseVersion = (version) => {
    if (typeof version !== "string" || !/^\d{1,5}(?:\.\d{1,5}){0,3}$/.test(version)) {
      throw new Error(`Unsupported extension version: ${String(version)}.`);
    }
    const parts = version.split(".").map(Number);
    if (parts.some((part) => part > 65535)) {
      throw new Error(`Unsupported extension version: ${version}.`);
    }
    return parts;
  };
  const leftParts = parseVersion(left);
  const rightParts = parseVersion(right);
  for (let index = 0; index < Math.max(leftParts.length, rightParts.length); index += 1) {
    const difference = (leftParts[index] || 0) - (rightParts[index] || 0);
    if (difference) {
      return difference;
    }
  }
  return 0;
}

function showUpdateNotice({ tag = "", error = "", checking = false }, showCurrent = false) {
  updateNoticeElement.hidden = true;
  updateNoticeElement.classList.toggle("is-error", Boolean(error));
  updateReleaseLink.href = RELEASES_URL;
  updateReleaseLink.hidden = checking;
  updateReleaseLink.textContent = error ? "View releases" : "View release";

  if (checking) {
    updateMessageElement.textContent = "Checking for updates...";
  } else if (error) {
    updateMessageElement.textContent = `Could not check for updates: ${error} Archiving is still available.`;
  } else {
    const version = tag.replace(/^v/i, "");
    const installedVersion = chrome.runtime.getManifest().version;
    if (compareExtensionVersions(version, installedVersion) <= 0) {
      if (!showCurrent) {
        return;
      }
      updateMessageElement.textContent = `No newer release is available. Installed: ${installedVersion}; latest release: ${version}.`;
      updateReleaseLink.textContent = "View releases";
    } else {
      updateMessageElement.textContent = `Version ${version} is available (installed: ${installedVersion}).`;
      updateReleaseLink.href = `${RELEASES_URL}/tag/${encodeURIComponent(tag)}`;
    }
  }

  updateNoticeElement.hidden = false;
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
  let activeTab;

  if (teamsTabId) {
    try {
      activeTab = await chrome.tabs.get(teamsTabId);
    } catch (error) {
      throw new Error(`The connected Teams tab is no longer available. Open Teams Website to reconnect. Details: ${getErrorDetails(error)}`);
    }
  } else {
    const tabs = await chrome.tabs.query({ active: true, currentWindow: true });
    activeTab = tabs[0];
  }

  if (!activeTab || !activeTab.id || !activeTab.url) {
    throw new Error("No active browser tab found.");
  }

  if (!isTeamsWebUrl(activeTab.url)) {
    throw new Error("Open Microsoft Teams on the web, then try again.");
  }

  return activeTab;
}

async function openTeamsWebsite() {
  const teamsTabs = await chrome.tabs.query({
    url: TEAMS_URL_PATTERNS
  });
  const teamsTab = teamsTabs.find((tab) => tab.id === teamsTabId) || teamsTabs[0];

  if (teamsTab?.id) {
    connectTeamsTab(teamsTab.id);
    await chrome.tabs.update(teamsTab.id, {
      active: true
    });

    if (typeof teamsTab.windowId === "number") {
      await chrome.windows.update(teamsTab.windowId, {
        focused: true
      });
    }

    return;
  }

  const teamsWindow = await chrome.windows.create({
    url: TEAMS_WEB_URL,
    type: "normal"
  });

  if (!teamsWindow.tabs?.[0]?.id) {
    throw new Error("Teams opened, but the browser did not return its tab. Close the archiver window and reopen it from Teams.");
  }

  connectTeamsTab(teamsWindow.tabs[0].id);
}

function connectTeamsTab(tabId) {
  if (teamsTabId !== tabId) {
    groupSelectionTabId = null;
    groupListElement.replaceChildren();
    groupPickerElement.hidden = true;
    archiveGroupsButton.setAttribute("aria-expanded", "false");
    updateGroupArchiveButton();
  }

  teamsTabId = tabId;
  window.history.replaceState(null, "", `?tabId=${tabId}`);
}

function isTeamsWebUrl(value) {
  try {
    const parsed = new URL(value);
    return (
      parsed.protocol === "https:" &&
      (parsed.hostname === "teams.cloud.microsoft" ||
        parsed.hostname === "teams.microsoft.com")
    );
  } catch (error) {
    return false;
  }
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

  chrome.runtime.onMessage.addListener((message, sender) => {
    if (isBusy && sender?.tab?.id === contentProgressTabId &&
        message?.type === "ARCHIVE_PROGRESS" && message.status) {
      setRunState("running");
      setStatus(message.status);
    }
  });

  progressListenerRegistered = true;
}

async function sendTeamsRequest(tabId, message) {
  ensureProgressListener();
  contentProgressTabId = tabId;

  try {
    return await chrome.tabs.sendMessage(tabId, message);
  } finally {
    contentProgressTabId = null;
  }
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

archiveGroupsButton.addEventListener("click", loadChatGroupSelection);

groupListElement.addEventListener("change", updateGroupArchiveButton);

cancelGroupSelectionButton.addEventListener("click", () => {
  groupPickerElement.hidden = true;
  archiveGroupsButton.setAttribute("aria-expanded", "false");
  archiveGroupsButton.focus();
});

archiveGroupSelectionButton.addEventListener("click", async () => {
  const groupIds = Array.from(groupListElement.querySelectorAll("input:checked"))
    .map((input) => input.value);

  if (groupIds.length) {
    await runArchiveFlow({ mode: "groups", groupIds });
  }
});

function updateGroupArchiveButton() {
  archiveGroupSelectionButton.disabled = isBusy || !groupListElement.querySelector("input:checked");
}

async function loadChatGroupSelection() {
  archiveProgressContainer.hidden = true;
  setBusy(true);
  setRunState("running");
  setStatus("Reading Teams chat groups...");

  try {
    const activeTab = await getActiveTeamsTab();
    await ensureContentScriptLoaded(activeTab.id);
    const response = await sendTeamsRequest(activeTab.id, { type: "LIST_CHAT_GROUPS" });

    if (!response?.ok) {
      throw new Error(response?.error || "Could not read Teams chat groups.");
    }

    if (!response.payload?.length) {
      throw new Error("No chat groups were found. Open the Teams chat list and try again.");
    }

    groupSelectionTabId = activeTab.id;
    groupListElement.replaceChildren();

    for (const group of response.payload) {
      const label = document.createElement("label");
      label.className = "group-option";
      const checkbox = document.createElement("input");
      checkbox.type = "checkbox";
      checkbox.value = group.id;
      const title = document.createElement("span");
      title.textContent = group.label;
      label.append(checkbox, title);
      groupListElement.append(label);
    }

    groupPickerElement.hidden = false;
    archiveGroupsButton.setAttribute("aria-expanded", "true");
    setRunState("idle");
    setStatus("");
  } catch (error) {
    groupSelectionTabId = null;
    groupListElement.replaceChildren();
    groupPickerElement.hidden = true;
    archiveGroupsButton.setAttribute("aria-expanded", "false");
    setRunState("error");
    setStatus(error.message || "Could not read Teams chat groups.", true);
  } finally {
    setBusy(false);
  }
}

async function runArchiveFlow({ mode, groupIds = [] }) {
  archiveStatusContext = "";
  setBusy(true);
  ensureProgressListener();
  setRunState("running");
  setStatus("Choose an archive folder...");
  showIndeterminateArchiveProgress("Choosing an archive folder...");
  let stage = "selecting the archive folder";
  let folderSelected = false;

  try {
    const rootDirectory = await window.showDirectoryPicker({
      mode: "readwrite"
    });
    folderSelected = true;

    stage = "saving archive settings";
    setStatus("Saving archive settings...");
    const settings = getSettingsFromForm();
    await saveSettings(settings);
    stage = "connecting to the active Teams tab";
    setStatus("Connecting to the Teams tab...");
    const activeTab = await getActiveTeamsTab();

    if (mode === "groups" && activeTab.id !== groupSelectionTabId) {
      throw new Error("The active Teams tab changed. Reload the group list and try again.");
    }

    teamsTabId = activeTab.id;
    await ensureContentScriptLoaded(activeTab.id);

    if (mode === "current") {
      updateArchiveProgress({ totalChats: 1 });
      stage = "collecting and saving the current chat";
      setStatus("Collecting messages from the current Teams chat...");
      const validation = await archiveCurrentChatToFolder(activeTab.id, rootDirectory, settings);
      updateArchiveProgress({ totalChats: 1, archivedCount: 1, phase: "finalizing" });
      archiveStatusContext = "";
      stage = "rebuilding the archive index and search files";
      setStatus("Chat archive saved and verified. Rebuilding the archive-wide index and search files...");
      await rebuildArchiveRootIndex(rootDirectory);
      stage = "verifying the archive index and search files";
      setStatus("Performing final verification of the archive index and search files...");
      await validateRootFiles(rootDirectory);
      updateArchiveProgress({ totalChats: 1, archivedCount: 1, phase: "finished" });
      setRunState("finished");
      setStatus(
        `Finished. Verified ${validation.chatTitle || "chat archive"}, index, and search files.`
      );
      return;
    }

    stage = "archiving the requested chats";
    const batchResult = await archiveVisibleChatsToFolder(activeTab.id, rootDirectory, settings, {
      groupIds: mode === "groups" ? groupIds : null
    });
    const batchProgress = {
      totalChats: batchResult.archivedCount + batchResult.skippedCount + batchResult.failedChats.length,
      archivedCount: batchResult.archivedCount,
      skippedCount: batchResult.skippedCount,
      failedCount: batchResult.failedChats.length
    };
    updateArchiveProgress({ ...batchProgress, phase: "finalizing" });
    stage = "rebuilding the archive index and search files";
    setStatus("Chat collection complete. Rebuilding the archive-wide index and search files...");
    await rebuildArchiveRootIndex(rootDirectory);
    stage = "verifying the archive index and search files";
    setStatus("Performing final verification of the archive index and search files...");
    await validateRootFiles(rootDirectory);
    updateArchiveProgress({ ...batchProgress, phase: "finished" });
    setRunState("finished");
    setStatus(formatBatchArchiveSummary(batchResult));
  } catch (error) {
    const pickerAborted = !folderSelected && error?.name === "AbortError";
    if (pickerAborted) {
      archiveProgressContainer.hidden = true;
      archiveProgressBar.removeAttribute("value");
      archiveProgressLabel.textContent = "";
    } else {
      archiveProgressContainer.classList.add("is-error");
      if (!archiveProgressBar.hasAttribute("value")) {
        archiveProgressBar.value = 0;
      }
      const label = `Stopped with error. ${archiveProgressLabel.textContent}`;
      archiveProgressLabel.textContent = label;
      archiveProgressBar.setAttribute("aria-valuetext", label);
    }
    console.error(`Archive failed while ${stage}.`, error);
    setRunState(pickerAborted ? "idle" : "error");
    setStatus(
      pickerAborted
        ? `Folder selection was cancelled or access was denied. Create or select an existing writable archive folder and approve read/write access. Details: ${getErrorDetails(error)}`
        : `Archive failed while ${stage}: ${getErrorDetails(error)}`,
      !pickerAborted
    );
  } finally {
    archiveStatusContext = "";
    setBusy(false);
  }
}

function setBusy(busy) {
  isBusy = busy;
  archiveButton.disabled = busy;
  archiveVisibleButton.disabled = busy;
  archiveGroupsButton.disabled = busy;
  cancelGroupSelectionButton.disabled = busy;
  openTeamsButton.disabled = busy;
  includeImagesCheckbox.disabled = busy;
  groupListElement.querySelectorAll("input").forEach((input) => {
    input.disabled = busy;
  });
  updateGroupArchiveButton();
}

async function archiveVisibleChatsToFolder(tabId, rootDirectory, settings, { groupIds = null } = {}) {
  showIndeterminateArchiveProgress(groupIds === null ? "Reading the visible chat list..." : "Loading all conversations in the selected groups...");
  setStatus(groupIds === null ? "Reading visible Teams chats..." : "Reading selected Teams chat groups...");

  const listResponse = await sendTeamsRequest(tabId, {
    type: groupIds === null ? "LIST_VISIBLE_CHATS" : "LIST_SELECTED_GROUP_CHATS",
    ...(groupIds === null ? {} : { groupIds })
  });

  if (!listResponse || !listResponse.ok) {
    throw new Error(listResponse?.error || "Could not read the requested Teams chat list.");
  }

  const chats = listResponse.payload || [];

  if (!chats.length) {
    throw new Error(groupIds === null
      ? "No chats were found in the current Teams view."
      : "No chats were found in the selected groups. Channels are not included.");
  }

  let archivedCount = 0;
  let skippedCount = 0;
  const failedChats = [];
  const validatedChats = [];
  const updateBatchProgress = () => updateArchiveProgress({
    totalChats: chats.length,
    archivedCount,
    skippedCount,
    failedCount: failedChats.length
  });
  updateBatchProgress();

  for (let index = 0; index < chats.length; index += 1) {
    const chat = chats[index];
    archiveStatusContext = `Chat ${index + 1} of ${chats.length}: "${chat.label}"`;
    setStatus(`Checking for an archive saved today (${archivedCount} saved, ${skippedCount} skipped, ${failedChats.length} failed so far)...`);
    const skipReason = await getExistingArchiveSkipReason(rootDirectory, chat.label);

    if (skipReason) {
      skippedCount += 1;
      updateBatchProgress();
      setStatus(`Skipped "${chat.label}": ${skipReason}`);
      await wait(300);
      continue;
    }

    setStatus("Opening the conversation in Teams...");

    const openResponse = await sendTeamsRequest(tabId, {
      type: "OPEN_CHAT_BY_LABEL",
      chatLabel: chat.label,
      chatId: chat.id || "",
      groupId: chat.groupId || ""
    });

    if (!openResponse || !openResponse.ok) {
      const reason = openResponse?.error || "could not open chat";
      failedChats.push({
        label: chat.label,
        reason
      });
      updateBatchProgress();
      setStatus(`Failed "${chat.label}": ${reason}`, true);
      await wait(1200);
      continue;
    }

    setStatus("Collecting messages from Teams...");
    try {
      const validation = await archiveCurrentChatToFolder(tabId, rootDirectory, settings, chat.label);
      archivedCount += 1;
      validatedChats.push(validation);
      updateBatchProgress();
    } catch (error) {
      const reason = getErrorDetails(error);
      console.error(`Archive failed for "${chat.label}".`, error);
      failedChats.push({
        label: chat.label,
        reason
      });
      updateBatchProgress();
      setStatus(`Failed "${chat.label}": ${reason}`, true);
      await wait(1200);
      continue;
    }

    setStatus(`Chat archive saved and verified (${archivedCount} saved, ${skippedCount} skipped, ${failedChats.length} failed so far).${index + 1 < chats.length ? " Preparing the next chat..." : ""}`);
    await wait(900);
  }

  archiveStatusContext = "";
  if (skippedCount > 0) {
    setStatus(`Archived ${archivedCount} chats and skipped ${skippedCount}.`);
    return {
      archivedCount,
      skippedCount,
      failedChats,
      validatedChats
    };
  }

  setStatus(`Archived ${archivedCount} chats.`);
  return {
    archivedCount,
    skippedCount,
    failedChats,
    validatedChats
  };
}

async function archiveCurrentChatToFolder(tabId, rootDirectory, settings, chatLabelOverride = "") {
  let response;

  try {
    response = await sendTeamsRequest(tabId, {
      type: "ARCHIVE_CURRENT_CHAT",
      chatLabelOverride
    });
  } catch (error) {
    throw new Error(`Lost connection to the Teams tab while collecting messages. The tab may have been closed, reloaded, navigated away, or discarded by the browser. Keep it open and retry. Details: ${getErrorDetails(error)}`);
  }

  if (!response || !response.ok) {
    throw new Error(response?.error || "The Teams page did not return any data.");
  }

  const snapshot = response.payload;
  if (!archiveStatusContext) {
    archiveStatusContext = `Chat: "${snapshot.chatTitle || chatLabelOverride || "current conversation"}"`;
  }
  setStatus(`Received ${snapshot.messages?.length || 0} messages from Teams. Preparing the snapshot...`);
  const mergedArchive = await rebuildMergedArchive(snapshot);
  setStatus(`Generating archive files for ${mergedArchive.messages.length} snapshot messages...`);
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

function getErrorMessage(error, fallbackMessage) {
  const message = String(error?.message || "").trim();
  return message || fallbackMessage;
}

function getErrorDetails(error) {
  const message = getErrorMessage(error, "Unexpected error while archiving chats.");
  return error?.name && error.name !== "Error" ? `${error.name}: ${message}` : message;
}

function formatBatchArchiveSummary(result) {
  const failedChats = Array.isArray(result.failedChats) ? result.failedChats : [];
  const lines = [
    `Finished. Archived ${result.archivedCount} chats successfully, skipped ${result.skippedCount}, failed ${failedChats.length}.`
  ];

  if (result.skippedCount > 0) {
    lines.push("Skipped chats were already archived today. Use Archive Current Chat to capture newer messages today without waiting.");
  }

  if (!failedChats.length) {
    lines.push("Verified the index and search files.");
    return lines.join("\n");
  }

  lines.push("Failed chats:");

  for (const failedChat of failedChats) {
    lines.push(`- ${failedChat.label}: ${failedChat.reason}`);
  }

  lines.push("Verified the index and search files for successful archives.");
  return lines.join("\n");
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
  setStatus("Creating or opening the chat archive folders...");
  const chatDirectory = await getOrCreateDirectory(rootDirectory, buildResponse.folderName);
  const snapshotsDirectory = await getOrCreateDirectory(chatDirectory, "snapshots");

  setStatus(`Writing new snapshot: ${buildResponse.snapshotFilename} (${snapshot.messages.length} messages)...`);
  await writeTextFile(
    snapshotsDirectory,
    buildResponse.snapshotFilename,
    JSON.stringify(snapshot, null, 2)
  );

  const mergedFromSnapshots = await rebuildFromSnapshots(snapshotsDirectory, mergedArchive);
  setStatus(`Generating merged archive files (${mergedFromSnapshots.messages.length} messages)...`);
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
  setStatus(`Generating final JSON and readable transcripts (${archiveWithImages.messages.length} messages)...`);
  const finalBuildResponse = await chrome.runtime.sendMessage({
    type: "BUILD_ARCHIVE_FILES",
    snapshot,
    mergedArchive: archiveWithImages,
    settings
  });

  if (!finalBuildResponse || !finalBuildResponse.ok) {
    throw new Error(finalBuildResponse?.error || "Could not build archive files with images.");
  }

  setStatus(`Writing latest.json (${archiveWithImages.messages.length} merged messages)...`);
  await writeTextFile(chatDirectory, "latest.json", finalBuildResponse.latestJsonText);
  setStatus("Writing the chat manifest...");
  await writeTextFile(chatDirectory, "manifest.json", finalBuildResponse.manifestText);

  setStatus("Writing the chat overview webpage...");
  await writeTextFile(chatDirectory, "latest.html", finalBuildResponse.latestHtmlText);
  await syncYearlyHtmlFiles(chatDirectory, finalBuildResponse.yearlyHtmlFiles || []);
  setStatus("Cleaning up legacy transcript files...");
  await deleteDirectoryContentsIfExists(chatDirectory, "md");

  setStatus("Verifying the saved chat snapshot and transcripts...");
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
  for (const filename of ["latest.json", "manifest.json", "latest.html"]) {
    setStatus(`Verifying saved chat file: ${filename}...`);
    await validateTextFile(chatDirectory, filename);
  }

  setStatus(`Verifying saved snapshot: ${snapshotFilename}...`);
  const snapshotsDirectory = await chatDirectory.getDirectoryHandle("snapshots");
  await validateTextFile(snapshotsDirectory, snapshotFilename);

  if (Array.isArray(yearlyHtmlFiles) && yearlyHtmlFiles.length) {
    const htmlDirectory = await chatDirectory.getDirectoryHandle("html");

    for (let index = 0; index < yearlyHtmlFiles.length; index += 1) {
      const file = yearlyHtmlFiles[index];
      if (!file?.filename) {
        continue;
      }

      setStatus(`Verifying yearly transcript ${index + 1} of ${yearlyHtmlFiles.length}: ${file.filename}...`);
      await validateTextFile(htmlDirectory, file.filename);
    }
  }
}

async function validateRootFiles(rootDirectory) {
  setStatus("Final verification 1 of 2: archive index (index.html)...");
  await validateTextFile(rootDirectory, "index.html");
  setStatus("Final verification 2 of 2: search page (search.html)...");
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
  setStatus("Reading saved conversations to rebuild the archive-wide index and search...");
  const entries = [];
  const searchRecords = [];

  for await (const entry of rootDirectory.values()) {
    if (entry.kind !== "directory") {
      continue;
    }

    setStatus(`Reading archive folder "${entry.name}" (${entries.length} chats indexed so far)...`);
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
      setStatus(`Building search entries for "${manifest.chatTitle || entry.name}": ${latestArchive.messages.length} messages (${entries.length} chats indexed)...`);
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

  setStatus(`Generating the archive index for ${entries.length} saved chats...`);
  const indexHtml = buildArchiveRootIndexHtml(entries);
  setStatus(`Writing index.html (${entries.length} saved chats)...`);
  await writeTextFile(rootDirectory, "index.html", indexHtml);
  setStatus(`Generating the search page with ${searchRecords.length} search entries...`);
  const searchHtml = buildArchiveSearchHtml(searchRecords);
  setStatus(`Writing search.html (${searchRecords.length} search entries)...`);
  await writeTextFile(rootDirectory, "search.html", searchHtml);
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
          data-chat-users="${escapeHtml(String(participantsLabel || "").toLowerCase())}"
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
    .search-wrap {
      margin: 0 0 18px;
    }
    .search-input {
      width: 100%;
      padding: 10px 12px;
      border: 1px solid var(--panel-border);
      border-radius: 10px;
      background: #ffffff;
      color: var(--text);
      font-size: 14px;
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
    <div class="search-wrap">
      <input id="chat-user-search" class="search-input" type="search" placeholder="Search chat users..." autocomplete="off">
    </div>
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
      const userSearchInput = document.getElementById("chat-user-search");
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

      const applyFilter = () => {
        const query = String(userSearchInput?.value || "").trim().toLowerCase();
        const rows = Array.from(tbody.querySelectorAll("tr"));

        for (const row of rows) {
          const users = getText(row, "data-chat-users");
          const matches = !query || users.includes(query);
          row.style.display = matches ? "" : "none";
        }
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
      if (userSearchInput) {
        userSearchInput.addEventListener("input", applyFilter);
      }
      applySort();
      applyFilter();
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
  setStatus(`Preparing ${yearlyHtmlFiles.length} yearly transcript files...`);
  const htmlDirectory = await getOrCreateDirectory(chatDirectory, "html");
  const wanted = new Set(yearlyHtmlFiles.map((file) => file.filename));

  for (let index = 0; index < yearlyHtmlFiles.length; index += 1) {
    const file = yearlyHtmlFiles[index];
    setStatus(`Writing yearly transcript ${index + 1} of ${yearlyHtmlFiles.length}: ${file.filename}...`);
    await writeTextFile(htmlDirectory, file.filename, file.content);
  }

  setStatus("Checking for obsolete yearly transcript files...");
  for await (const entry of htmlDirectory.values()) {
    if (entry.kind !== "file") {
      continue;
    }

    if (!wanted.has(entry.name)) {
      setStatus(`Removing obsolete yearly transcript: ${entry.name}...`);
      await htmlDirectory.removeEntry(entry.name);
    }
  }
}

async function rebuildFromSnapshots(snapshotsDirectory, fallbackArchive) {
  setStatus("Reading saved snapshots to merge conversation history...");
  const snapshots = [];
  let filesRead = 0;

  for await (const entry of snapshotsDirectory.values()) {
    if (entry.kind !== "file" || !entry.name.endsWith(".json")) {
      continue;
    }

    filesRead += 1;
    setStatus(`Reading saved snapshot ${filesRead}: ${entry.name}...`);
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

  const snapshotMessages = snapshots.flatMap((snapshot) => snapshot.messages || []);
  setStatus(`Merging ${snapshots.length} snapshots: deduplicating and sorting ${snapshotMessages.length} message records...`);
  const mergedMessages = dedupeAndSortMessages(snapshotMessages);
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
  const byKey = new Map();
  const deduped = [];

  for (const message of messages) {
    const normalized = normalizeMessage(message);
    const key = buildMessageKey(normalized);

    const previous = byKey.get(key);

    if (previous) {
      if (previous.threadId || normalized.threadId) {
        previous.threadId = normalized.threadId || previous.threadId;
        if (normalized.deleted) {
          previous.deleted = true;
        }
        previous.text = normalized.text.length >= previous.text.length ? normalized.text : previous.text;
        previous.author = normalized.author !== "Unknown" ? normalized.author : previous.author;
        previous.timestamp = normalized.timestamp || previous.timestamp;
        previous.images = normalizeImages([...previous.images, ...normalized.images]);
      }
      continue;
    }

    byKey.set(key, normalized);
    deduped.push(normalized);
  }

  return deduped.sort(compareMessages);
}

function normalizeMessage(message) {
  return {
    id: String(message?.id || "").trim(),
    ...(message?.threadId ? { threadId: String(message.threadId).trim() } : {}),
    ...(message?.deleted === true ? { deleted: true } : {}),
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
  try {
    return await parentHandle.getDirectoryHandle(name, { create: true });
  } catch (error) {
    throw new Error(`Could not create or open folder "${name}": ${getErrorDetails(error)}`, { cause: error });
  }
}

async function writeTextFile(directoryHandle, filename, contents) {
  try {
    const fileHandle = await directoryHandle.getFileHandle(filename, { create: true });
    const writable = await fileHandle.createWritable();
    await writable.write(contents);
    await writable.close();
  } catch (error) {
    throw new Error(`Could not write file "${filename}": ${getErrorDetails(error)}`, { cause: error });
  }
}

async function materializeInlineImages(chatDirectory, archive) {
  const totalImages = (archive.messages || []).reduce((count, message) => count + (message.images || []).length, 0);
  setStatus(`Preparing image assets: ${totalImages} inline image references to process...`);
  const imagesDirectory = await getOrCreateDirectory(chatDirectory, "assets");
  const nestedImagesDirectory = await getOrCreateDirectory(imagesDirectory, "images");
  const wantedPaths = new Set();
  const updatedMessages = [];
  let processedImages = 0;

  for (let messageIndex = 0; messageIndex < (archive.messages || []).length; messageIndex += 1) {
    const message = archive.messages[messageIndex];
    const updatedImages = [];

    for (let imageIndex = 0; imageIndex < (message.images || []).length; imageIndex += 1) {
      const image = message.images[imageIndex];
      processedImages += 1;
      setStatus(`Processing pasted image ${processedImages} of ${totalImages} (message ${messageIndex + 1} of ${archive.messages.length}): fetching and writing image data...`);
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

  setStatus(`Cleaning unused image assets (${wantedPaths.size} referenced image paths)...`);
  await pruneUnusedImageAssets(nestedImagesDirectory, wantedPaths);

  return {
    ...archive,
    messages: updatedMessages,
    messageCount: updatedMessages.length
  };
}

async function stripInlineImages(chatDirectory, archive) {
  setStatus("Images are disabled. Removing image references and cleaning saved image assets...");
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
    return /(^|\.)(teams\.microsoft\.com|teams\.cloud\.microsoft)$/i.test(parsed.hostname);
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
