chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (message?.type !== "BUILD_ARCHIVE_FILES") {
    return false;
  }

  try {
    const snapshot = message.snapshot || {};
    const mergedArchive = message.mergedArchive || snapshot;
    const settings = normalizeSettings(message.settings || {});
    const archiveLabel = buildArchiveLabel(mergedArchive);
    const folderName =
      sanitizeFilePart(cleanDisplayName(mergedArchive.chatTitle || archiveLabel)) || "teams-chat";
    const snapshotStamp = formatTimestampForFile(snapshot.exportedAt || new Date().toISOString());
    const snapshotFilename = `${snapshotStamp}.json`;
    const latestJsonText = JSON.stringify(mergedArchive, null, 2);
    const yearlyHtmlFiles = buildYearlyHtmlFiles(mergedArchive);
    const latestHtmlText = buildChatOverviewHtml(mergedArchive, yearlyHtmlFiles);
    const manifestText = JSON.stringify(
      buildManifest({
        folderName,
        archiveLabel,
        snapshot,
        mergedArchive,
        snapshotFilename,
        settings,
        yearlyHtmlFiles
      }),
      null,
      2
    );

    sendResponse({
      ok: true,
      folderName,
      snapshotFilename,
      latestJsonText,
      latestHtmlText,
      yearlyHtmlFiles,
      manifestText
    });
  } catch (error) {
    sendResponse({
      ok: false,
      error: error.message || "Could not build archive files."
    });
  }

  return false;
});

function normalizeSettings(settings) {
  const includeImages = settings.includeImages !== false;

  return {
    includeImages
  };
}

function buildManifest({
  folderName,
  archiveLabel,
  snapshot,
  mergedArchive,
  snapshotFilename,
  settings,
  yearlyHtmlFiles
}) {
  return {
    folderName,
    archiveLabel,
    chatTitle: mergedArchive.chatTitle || snapshot.chatTitle || "",
    chatOrder:
      Number.isFinite(Number(mergedArchive.chatOrder)) && Number(mergedArchive.chatOrder) >= 0
        ? Number(mergedArchive.chatOrder)
        : Number.isFinite(Number(snapshot.chatOrder)) && Number(snapshot.chatOrder) >= 0
          ? Number(snapshot.chatOrder)
          : -1,
    participants: listParticipants(mergedArchive.messages || []),
    firstMessageAt: findBoundaryTimestamp(mergedArchive.messages || [], "first"),
    lastMessageAt: findBoundaryTimestamp(mergedArchive.messages || [], "last"),
    lastArchivedAt: snapshot.exportedAt || "",
    latestMessageCount: (mergedArchive.messages || []).length,
    newestSnapshot: snapshotFilename,
    includeImages: settings.includeImages,
    yearlyHtmlFiles: yearlyHtmlFiles.map((file) => file.filename)
  };
}

function buildChatOverviewHtml(archive, yearlyHtmlFiles) {
  const chatTitle = String(archive.chatTitle || "").trim() || "Teams Chat Archive";
  const firstMessageAt = findBoundaryTimestamp(archive.messages || [], "first");
  const lastMessageAt = findBoundaryTimestamp(archive.messages || [], "last");
  const recentGroups = groupMessagesForDisplay(archive.messages || []).slice(-10);
  const recentHtml = buildMessageCardsHtml(recentGroups, "");
  const yearsHtml = yearlyHtmlFiles
    .map((file) => {
      return `<li><a href="html/${escapeHtml(file.filename)}">${escapeHtml(file.label)}</a></li>`;
    })
    .join("\n");

  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${escapeHtml(chatTitle)}</title>
  <style>${sharedHtmlStyles()}</style>
</head>
<body>
  <main>
    <section class="summary">
      <p class="meta"><a href="../index.html">Back to main index</a></p>
      <h1>${escapeHtml(chatTitle)}</h1>
      <p class="meta">Exported: ${escapeHtml(formatTimestampDisplay(archive.exportedAt || ""))}</p>
      <p class="meta">Exported range: ${escapeHtml(formatDateRange(firstMessageAt, lastMessageAt))}</p>
      <p class="meta">Messages captured: ${escapeHtml(String((archive.messages || []).length))}</p>
    </section>

    <section class="summary">
      <h2 class="section-title">Yearly Transcripts</h2>
      <ul class="year-list">
        ${yearsHtml || "<li>No yearly transcript files available.</li>"}
      </ul>
    </section>

    <section class="summary">
      <h2 class="section-title">Recent Messages</h2>
      <div class="messages">
        ${recentHtml || "<p class=\"meta\">No recent messages available.</p>"}
      </div>
    </section>
  </main>
</body>
</html>`;
}

function buildYearlyHtmlFiles(archive) {
  const chatTitle = String(archive.chatTitle || "").trim() || "Teams Chat Archive";
  const groupedMessages = groupMessagesForDisplay(archive.messages || []);
  const buckets = new Map();

  for (const group of groupedMessages) {
    const parsed = parseTimestampValue(group.timestamp);
    const yearKey = parsed ? String(parsed.getFullYear()) : "unknown";

    if (!buckets.has(yearKey)) {
      buckets.set(yearKey, []);
    }

    buckets.get(yearKey).push(group);
  }

  return Array.from(buckets.entries())
    .sort((left, right) => left[0].localeCompare(right[0]))
    .map(([yearKey, groups]) => {
      return {
        filename: `${yearKey}.html`,
        label: yearKey === "unknown" ? "Unknown dates" : yearKey,
        content: buildYearlyTranscriptHtml(chatTitle, yearKey, groups)
      };
    });
}

function buildYearlyTranscriptHtml(chatTitle, yearKey, groups) {
  const messageHtml = buildMessageCardsHtml(groups, "../");
  const yearMessages = groups.flatMap((group) =>
    (group.parts || []).map((part) => ({
      timestamp: part.timestamp || group.timestamp || ""
    }))
  );
  const firstMessageAt = findBoundaryTimestamp(yearMessages, "first");
  const lastMessageAt = findBoundaryTimestamp(yearMessages, "last");

  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${escapeHtml(chatTitle)}</title>
  <style>${sharedHtmlStyles()}</style>
</head>
<body>
  <main>
    <section class="summary">
      <p class="meta"><a href="../../index.html">Back to main index</a></p>
      <h1>${escapeHtml(chatTitle)}</h1>
      <p class="meta">Year: ${escapeHtml(yearKey)}</p>
      <p class="meta">Exported range: ${escapeHtml(formatDateRange(firstMessageAt, lastMessageAt))}</p>
      <p class="meta"><a href="../latest.html">Back to chat overview</a></p>
    </section>
    <section class="messages">
      ${messageHtml}
    </section>
  </main>
</body>
</html>`;
}

function buildMessageCardsHtml(groups, assetPathPrefix = "") {
  return groups
    .map((group) => {
      const partsHtml = group.parts
        .map((part) => {
          const imageHtml = (part.images || [])
            .filter((image) => image.localPath)
            .map((image) => {
              const imagePath = image.localPath.startsWith("assets/")
                ? `${assetPathPrefix}${image.localPath}`
                : image.localPath;

              return `<figure class="message-image-wrap"><img class="message-image" src="${escapeHtml(
                imagePath
              )}" alt="${escapeHtml(image.alt || "")}"></figure>`;
            })
            .join("\n");

          return `<div class="message-part-wrap">
            <pre class="message-part">${escapeHtml(part.text || "")}</pre>
            ${imageHtml}
          </div>`;
        })
        .join("\n");

      return `
        <article class="message">
          <header class="message-header">
            <span class="author">${escapeHtml(cleanDisplayName(group.author || "Unknown"))}</span>
            <time class="timestamp">${escapeHtml(formatTimestampDisplay(group.timestamp || ""))}</time>
          </header>
          <section class="message-body">
            ${partsHtml}
          </section>
        </article>
      `;
    })
    .join("\n");
}

function sharedHtmlStyles() {
  return `
    :root {
      color-scheme: light;
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
      background: linear-gradient(180deg, #eef4f9 0%, #f8fbfd 100%);
      color: var(--text);
    }
    main {
      max-width: 960px;
      margin: 0 auto;
      padding: 32px 20px 48px;
    }
    .summary,
    .message {
      background: var(--panel);
      border: 1px solid var(--panel-border);
      border-radius: 14px;
      box-shadow: 0 10px 30px rgba(9, 30, 66, 0.06);
    }
    .summary {
      padding: 24px;
      margin-bottom: 20px;
    }
    h1 {
      margin: 0 0 8px;
      font-size: 30px;
      line-height: 1.15;
    }
    .section-title {
      margin: 0 0 10px;
      font-size: 20px;
      line-height: 1.2;
    }
    .meta {
      margin: 6px 0;
      color: var(--muted);
      font-size: 14px;
    }
    .meta a {
      color: var(--accent);
      text-decoration: none;
    }
    .meta a:hover {
      text-decoration: underline;
    }
    .year-list {
      margin: 0;
      padding-left: 20px;
    }
    .year-list li {
      margin: 6px 0;
    }
    .year-list a {
      color: var(--accent);
      text-decoration: none;
    }
    .year-list a:hover {
      text-decoration: underline;
    }
    .messages {
      display: grid;
      gap: 14px;
    }
    .message {
      padding: 16px 18px;
    }
    .message-header {
      display: flex;
      justify-content: space-between;
      gap: 12px;
      margin-bottom: 10px;
      font-size: 14px;
    }
    .author {
      font-weight: 700;
      color: var(--accent);
    }
    .timestamp {
      color: var(--muted);
      white-space: nowrap;
    }
    .message-body {
      display: grid;
      gap: 10px;
    }
    .message-part-wrap {
      display: grid;
      gap: 10px;
    }
    .message-part {
      margin: 0;
      white-space: pre-wrap;
      word-break: break-word;
      font-family: inherit;
      line-height: 1.5;
      font-size: 15px;
    }
    .message-image-wrap {
      margin: 0;
    }
    .message-image {
      display: block;
      max-width: 100%;
      height: auto;
      border-radius: 10px;
      border: 1px solid var(--panel-border);
    }
  `;
}

function groupMessagesForDisplay(messages) {
  const groups = [];

  for (const message of messages) {
    const previousGroup = groups[groups.length - 1];

    if (shouldMergeMessages(previousGroup, message)) {
      previousGroup.parts.push({
        id: message.id || "",
        timestamp: message.timestamp || "",
        text: message.text || "",
        images: message.images || []
      });

      if (message.timestamp) {
        previousGroup.timestamp = message.timestamp;
      }

      continue;
    }

    groups.push({
      author: message.author || "Unknown",
      timestamp: message.timestamp || "",
      parts: [
        {
          id: message.id || "",
          timestamp: message.timestamp || "",
          text: message.text || "",
          images: message.images || []
        }
      ]
    });
  }

  return groups;
}

function shouldMergeMessages(previousGroup, message) {
  if (!previousGroup || !message) {
    return false;
  }

  if ((previousGroup.author || "") !== (message.author || "")) {
    return false;
  }

  const previousPart = previousGroup.parts[previousGroup.parts.length - 1];
  const previousTime = parseTimestampValue(previousPart?.timestamp || previousGroup.timestamp || "");
  const currentTime = parseTimestampValue(message.timestamp || "");

  if (previousTime && currentTime) {
    return Math.abs(currentTime.getTime() - previousTime.getTime()) <= 5 * 60 * 1000;
  }

  return true;
}

function buildArchiveLabel(archive) {
  const chatTitle = String(archive.chatTitle || "").trim();
  const participantPart = buildParticipantLabel(archive.messages || []);
  const periodPart = buildTimePeriodLabel(archive.messages || [], archive.exportedAt || "");

  if (chatTitle && periodPart) {
    return `${chatTitle}_${periodPart}`;
  }

  return chatTitle || participantPart || periodPart || "teams-chat";
}

function buildParticipantLabel(messages) {
  const uniqueAuthors = [];
  const seen = new Set();

  for (const message of messages) {
    const author = (message.author || "").trim();

    if (!author || author === "Unknown") {
      continue;
    }

    const compact = normalizeAuthorName(author);

    if (!compact || seen.has(compact)) {
      continue;
    }

    seen.add(compact);
    uniqueAuthors.push(compact);
  }

  return uniqueAuthors.slice(0, 4).join("_");
}

function listParticipants(messages) {
  return buildParticipantLabel(messages)
    .split("_")
    .filter(Boolean);
}

function normalizeAuthorName(author) {
  const cleaned = author
    .replace(/\s*\([^)]*\)\s*/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  if (!cleaned) {
    return "";
  }

  const commaMatch = cleaned.match(/^([^,]+),\s*(.+)$/);

  if (commaMatch) {
    return toTitleToken(commaMatch[1]);
  }

  const parts = cleaned.split(" ").filter(Boolean);
  return parts.length ? toTitleToken(parts[parts.length - 1]) : "";
}

function toTitleToken(value) {
  return value
    .replace(/[^A-Za-z0-9'-]+/g, "")
    .replace(/^[^A-Za-z0-9]+|[^A-Za-z0-9]+$/g, "");
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

function buildTimePeriodLabel(messages, exportedAt) {
  const parsedDates = messages
    .map((message) => parseTimestampValue(message.timestamp || ""))
    .filter(Boolean)
    .sort((a, b) => a.getTime() - b.getTime());

  if (!parsedDates.length) {
    const fallbackDate = parseTimestampValue(exportedAt);
    return fallbackDate ? formatDateForFile(fallbackDate) : "";
  }

  const firstLabel = formatDateForFile(parsedDates[0]);
  const lastLabel = formatDateForFile(parsedDates[parsedDates.length - 1]);
  return firstLabel === lastLabel ? firstLabel : `${firstLabel}_to_${lastLabel}`;
}

function findBoundaryTimestamp(messages, which) {
  const parsedDates = messages
    .map((message) => parseTimestampValue(message.timestamp || ""))
    .filter(Boolean)
    .sort((a, b) => a.getTime() - b.getTime());

  if (!parsedDates.length) {
    return "";
  }

  const date = which === "first" ? parsedDates[0] : parsedDates[parsedDates.length - 1];
  return date.toISOString();
}

function formatDateForFile(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function formatTimestampForFile(value) {
  const parsed = parseTimestampValue(value);
  const date = parsed || new Date(value);
  const safeDate = Number.isNaN(date.getTime()) ? new Date() : date;
  const year = safeDate.getFullYear();
  const month = String(safeDate.getMonth() + 1).padStart(2, "0");
  const day = String(safeDate.getDate()).padStart(2, "0");
  const hours = String(safeDate.getHours()).padStart(2, "0");
  const minutes = String(safeDate.getMinutes()).padStart(2, "0");
  const seconds = String(safeDate.getSeconds()).padStart(2, "0");
  return `${year}-${month}-${day}T${hours}-${minutes}-${seconds}`;
}

function formatTimestampDisplay(value) {
  const parsed = parseTimestampValue(value);

  if (!parsed) {
    return value || "";
  }

  return new Intl.DateTimeFormat(undefined, {
    year: "numeric",
    month: "short",
    day: "numeric",
    hour: "numeric",
    minute: "2-digit"
  }).format(parsed);
}

function formatDateRange(firstValue, lastValue) {
  const first = formatDateDisplay(firstValue);
  const last = formatDateDisplay(lastValue);

  if (!first && !last) {
    return "Unknown";
  }

  if (first && last && first !== last) {
    return `${first} to ${last}`;
  }

  return first || last;
}

function formatDateDisplay(value) {
  const parsed = parseTimestampValue(value);

  if (!parsed) {
    return "";
  }

  return new Intl.DateTimeFormat(undefined, {
    year: "numeric",
    month: "short",
    day: "numeric"
  }).format(parsed);
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

function sanitizeFilePart(value) {
  return String(value || "")
    .trim()
    .replace(/[<>:"/\\|?*]+/g, "-")
    .replace(/\s+/g, "-")
    .replace(/-+/g, "-")
    .replace(/^-|-$/g, "")
    .toLowerCase();
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}
