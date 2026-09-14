const contentElement = document.getElementById("content");
const statusElement = document.getElementById("status");
const folderLabelElement = document.getElementById("folder-label");
const chooseFolderButton = document.getElementById("choose-folder-button");
const openTeamsButton = document.getElementById("open-teams-button");

const TEAMS_WEB_URL = "https://teams.cloud.microsoft/";
const TEAMS_URL_PATTERNS = [
  "https://teams.cloud.microsoft/*",
  "https://teams.microsoft.com/*"
];

initializeViewer().catch((error) => {
  setStatus(error.message || "Could not load archive viewer.", true);
});

chooseFolderButton.addEventListener("click", async () => {
  try {
    const handle = await window.showDirectoryPicker({
      mode: "readwrite"
    });
    await window.archiveFolderStore.saveArchiveRootHandle(handle);
    await chrome.storage.local.set({
      archiveRootName: handle.name || ""
    });
    await initializeViewer();
  } catch (error) {
    if (error?.name !== "AbortError") {
      setStatus(error.message || "Could not change the archive folder.", true);
    }
  }
});

openTeamsButton.addEventListener("click", async () => {
  const teamsTabs = await chrome.tabs.query({
    url: TEAMS_URL_PATTERNS
  });
  const teamsTab = teamsTabs[0];

  if (teamsTab?.id) {
    await chrome.tabs.update(teamsTab.id, { active: true });

    if (typeof teamsTab.windowId === "number") {
      await chrome.windows.update(teamsTab.windowId, { focused: true });
    }

    return;
  }

  await chrome.tabs.create({
    url: TEAMS_WEB_URL
  });
});

async function initializeViewer() {
  const storedName = (await chrome.storage.local.get({ archiveRootName: "" })).archiveRootName || "";
  const handle = await window.archiveFolderStore.loadArchiveRootHandle();

  folderLabelElement.textContent = storedName
    ? `Archive folder: ${storedName}`
    : "Archive folder not set yet.";

  if (!handle) {
    contentElement.innerHTML = '<div class="panel"><div class="empty">Choose an archive folder to browse saved chats here.</div></div>';
    setStatus("No archive folder is saved yet.");
    return;
  }

  const readGranted = await window.archiveFolderStore.queryDirectoryPermission(handle, "read");

  if (!readGranted) {
    contentElement.innerHTML = '<div class="panel"><div class="empty">Access to the saved archive folder is not currently granted. Click "Change Archive Folder" and select the same folder again to continue.</div></div>';
    setStatus("Archive folder access is no longer available.", true);
    return;
  }

  const params = new URLSearchParams(window.location.search);
  const chatFolder = params.get("chat") || "";
  const year = params.get("year") || "";

  if (!chatFolder) {
    await renderIndexView(handle);
    return;
  }

  await renderDocumentView(handle, chatFolder, year);
}

function setStatus(message, isError = false) {
  statusElement.textContent = message;
  statusElement.style.color = isError ? "#b42318" : "#52606d";
}

async function renderIndexView(rootHandle) {
  const entries = [];

  for await (const entry of rootHandle.values()) {
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
      : [];

    entries.push({
      folderName: entry.name,
      chatTitle: String(manifest.chatTitle || manifest.archiveLabel || entry.name),
      chatOrder: Number.isFinite(Number(manifest.chatOrder)) ? Number(manifest.chatOrder) : -1,
      participants,
      lastArchivedAt: String(manifest.lastArchivedAt || ""),
      latestMessageCount: Number(manifest.latestMessageCount || 0),
      firstMessageAt: String(manifest.firstMessageAt || ""),
      lastMessageAt: String(manifest.lastMessageAt || "")
    });
  }

  entries.sort((left, right) => {
    const leftOrder = left.chatOrder >= 0 ? left.chatOrder : Number.MAX_SAFE_INTEGER;
    const rightOrder = right.chatOrder >= 0 ? right.chatOrder : Number.MAX_SAFE_INTEGER;

    if (leftOrder !== rightOrder) {
      return leftOrder - rightOrder;
    }

    const leftTime = Date.parse(left.lastArchivedAt || "") || 0;
    const rightTime = Date.parse(right.lastArchivedAt || "") || 0;

    if (leftTime !== rightTime) {
      return rightTime - leftTime;
    }

    return left.chatTitle.localeCompare(right.chatTitle);
  });

  const rowsHtml = entries
    .map((entry) => {
      const link = `archive-viewer.html?chat=${encodeURIComponent(entry.folderName)}`;
      return `
        <tr>
          <td><a href="${link}">${escapeHtml(entry.chatTitle)}</a></td>
          <td>${escapeHtml((entry.participants || []).join(", ") || "Unknown")}</td>
          <td>${escapeHtml(formatDisplayDate(entry.lastArchivedAt))}</td>
          <td>${escapeHtml(String(entry.latestMessageCount))}</td>
          <td>${escapeHtml(formatRange(entry.firstMessageAt, entry.lastMessageAt))}</td>
        </tr>
      `;
    })
    .join("\n");

  contentElement.innerHTML = rowsHtml
    ? `<div class="panel table-wrap"><table>
        <thead>
          <tr>
            <th>Chat</th>
            <th>Chat Users</th>
            <th>Last Archived</th>
            <th>Messages</th>
            <th>Chat Range</th>
          </tr>
        </thead>
        <tbody>${rowsHtml}</tbody>
      </table></div>`
    : '<div class="panel"><div class="empty">No archived chats were found in the saved folder.</div></div>';

  setStatus(`Loaded ${entries.length} archived chats.`);
}

async function renderDocumentView(rootHandle, chatFolder, year) {
  const chatHandle = await rootHandle.getDirectoryHandle(chatFolder);
  const relativePath = year ? `html/${year}.html` : "latest.html";
  const htmlText = await readTextByRelativePath(chatHandle, relativePath);
  const rewrittenHtml = await rewriteArchiveHtml(rootHandle, chatFolder, htmlText);
  const titleMatch = rewrittenHtml.match(/<title>(.*?)<\/title>/i);
  const title = titleMatch ? titleMatch[1] : chatFolder;

  contentElement.innerHTML = `<div class="document"><iframe title="${escapeHtml(title)}" srcdoc="${escapeHtml(rewrittenHtml)}"></iframe></div>`;
  setStatus(`Viewing ${year ? `${chatFolder} (${year})` : chatFolder}.`);
}

async function rewriteArchiveHtml(rootHandle, chatFolder, htmlText) {
  const parser = new DOMParser();
  const documentNode = parser.parseFromString(htmlText, "text/html");

  for (const link of Array.from(documentNode.querySelectorAll("a[href]"))) {
    const href = link.getAttribute("href") || "";

    if (href === "../index.html" || href === "../../index.html" || href === "index.html") {
      link.setAttribute("href", "archive-viewer.html");
      continue;
    }

    if (href === "../latest.html") {
      link.setAttribute("href", `archive-viewer.html?chat=${encodeURIComponent(chatFolder)}`);
      continue;
    }

    const yearlyMatch = href.match(/^html\/([^/]+)\.html$/);
    if (yearlyMatch) {
      link.setAttribute(
        "href",
        `archive-viewer.html?chat=${encodeURIComponent(chatFolder)}&year=${encodeURIComponent(yearlyMatch[1])}`
      );
    }
  }

  for (const image of Array.from(documentNode.querySelectorAll("img[src]"))) {
    const source = image.getAttribute("src") || "";

    if (!source || /^https?:|^data:/i.test(source)) {
      continue;
    }

    const normalized = source.replace(/^\.\.\//, "").replace(/^\.\//, "");
    const dataUrl = await readImageAsDataUrl(rootHandle, chatFolder, normalized);

    if (dataUrl) {
      image.setAttribute("src", dataUrl);
    }
  }

  return documentNode.documentElement.outerHTML;
}

async function readImageAsDataUrl(rootHandle, chatFolder, normalizedRelativePath) {
  const pathParts = [chatFolder, ...normalizedRelativePath.split("/").filter(Boolean)];
  const file = await readFileByPath(rootHandle, pathParts);

  if (!file) {
    return "";
  }

  return fileToDataUrl(file);
}

async function readTextByRelativePath(directoryHandle, relativePath) {
  const file = await readFileByPath(directoryHandle, relativePath.split("/").filter(Boolean));

  if (!file) {
    throw new Error(`Could not read ${relativePath}.`);
  }

  return file.text();
}

async function readFileByPath(startHandle, parts) {
  let current = startHandle;

  for (let index = 0; index < parts.length - 1; index += 1) {
    current = await current.getDirectoryHandle(parts[index]);
  }

  const fileHandle = await current.getFileHandle(parts[parts.length - 1]);
  return fileHandle.getFile();
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

function fileToDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ""));
    reader.onerror = () => reject(reader.error || new Error("Could not read file."));
    reader.readAsDataURL(file);
  });
}

function listArchiveParticipants(messages) {
  const participants = [];
  const seen = new Set();

  for (const message of messages || []) {
    const cleaned = cleanDisplayName(message?.author || "");

    if (!cleaned || cleaned === "Unknown" || seen.has(cleaned) || !looksLikePersonDisplayName(cleaned)) {
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

  if (!cleaned || cleaned.length > 80 || /[.!?]/.test(cleaned)) {
    return false;
  }

  const parts = cleaned.split(/\s+/).filter(Boolean);

  if (parts.length < 1 || parts.length > 5) {
    return false;
  }

  return parts.every((part) => /^(?:[A-Z][\p{L}\p{M}'-]*|[A-Z]\.)$/u.test(part));
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

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}
