chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (message?.type === "PING_TEAMS_ARCHIVER") {
    sendResponse({ ok: true });
    return false;
  }

  if (message?.type === "LIST_VISIBLE_CHATS") {
    try {
      sendResponse({
        ok: true,
        payload: listVisibleChats()
      });
    } catch (error) {
      sendResponse({
        ok: false,
        error: error.message || "Could not read the visible Teams chat list."
      });
    }
    return false;
  }

  if (message?.type === "OPEN_CHAT_BY_LABEL") {
    openChatByLabel(message.chatLabel)
      .then((payload) => sendResponse({ ok: true, payload }))
      .catch((error) => {
        sendResponse({
          ok: false,
          error: error.message || "Could not open the requested Teams chat."
        });
      });
    return true;
  }

  if (message?.type !== "ARCHIVE_CURRENT_CHAT") {
    return false;
  }

  archiveCurrentChat(message.chatLabelOverride)
    .then((payload) => sendResponse({ ok: true, payload }))
    .catch((error) => {
      sendResponse({
        ok: false,
        error: error.message || "Unknown error while reading the Teams chat."
      });
    });

  return true;
});

async function archiveCurrentChat(chatLabelOverride = "") {
  const chatTitle = getChatTitle(chatLabelOverride);
  const chatOrder = getCurrentChatListOrder(chatTitle);
  const collectedMessages = collectMessages();
  await loadOlderMessages(collectedMessages);
  const messages = dedupeMessageRecords([
    ...collectedMessages,
    ...collectMessages()
  ]);
  await enrichMessagesWithEmbeddedImages(messages);

  if (!messages.length) {
    throw new Error(
      "No messages were found. Open a Teams chat with visible messages and try again."
    );
  }

  return {
    exportedAt: new Date().toISOString(),
    pageTitle: document.title,
    pageUrl: window.location.href,
    chatTitle,
    chatOrder,
    messageCount: messages.length,
    messages
  };
}

function listVisibleChats() {
  const chatItems = getChatListEntries();
  const seen = new Set();
  const chats = [];

  for (const item of chatItems) {
    if (!looksLikeChatListEntry(item)) {
      continue;
    }

    const chatLabel = getChatListEntryLabel(item);

    if (!chatLabel || seen.has(chatLabel)) {
      continue;
    }

    seen.add(chatLabel);
    chats.push({
      label: chatLabel,
      order: chats.length
    });
  }

  return chats;
}

async function openChatByLabel(chatLabel) {
  if (!chatLabel) {
    throw new Error("No chat label was provided.");
  }

  const chatEntry = findChatListEntry(chatLabel);

  if (!chatEntry) {
    throw new Error(`Could not find the chat "${chatLabel}" in the visible Teams chat list.`);
  }

  chatEntry.scrollIntoView({
    block: "center"
  });
  await wait(250);

  clickElement(chatEntry);
  await waitForChatToOpen(chatLabel);

  return {
    openedChatLabel: chatLabel
  };
}

async function loadOlderMessages(collectedMessages = []) {
  const scrollContainer = getMessageScrollContainer();

  if (!scrollContainer) {
    return;
  }

  let previousScrollTop = scrollContainer.scrollTop;
  let previousScrollHeight = scrollContainer.scrollHeight;
  let previousMessageCount = collectMessages().length;
  let stableRounds = 0;
  const maxRounds = 80;

  for (let round = 0; round < maxRounds; round += 1) {
    notifyProgress(`Loading older messages (${round + 1}/${maxRounds})...`);
    appendUniqueMessages(collectedMessages, collectMessages());

    if (scrollContainer.scrollTop <= 0) {
      scrollContainer.scrollTop = 0;
    } else {
      const nextTop = Math.max(0, scrollContainer.scrollTop - Math.max(600, scrollContainer.clientHeight));
      scrollContainer.scrollTop = nextTop;
    }

    dispatchScrollEvents(scrollContainer);
    await wait(900);

    const currentMessageCount = collectMessages().length;
    const currentScrollTop = scrollContainer.scrollTop;
    const currentScrollHeight = scrollContainer.scrollHeight;
    const noScrollMovement = currentScrollTop === previousScrollTop;
    const noGrowth = currentScrollHeight === previousScrollHeight;
    const noNewMessages = currentMessageCount === previousMessageCount;

    if ((currentScrollTop <= 0 && noNewMessages) || (noScrollMovement && noGrowth && noNewMessages)) {
      stableRounds += 1;
    } else {
      stableRounds = 0;
    }

    previousScrollTop = currentScrollTop;
    previousScrollHeight = currentScrollHeight;
    previousMessageCount = currentMessageCount;

    if (stableRounds >= 3) {
      break;
    }
  }

  appendUniqueMessages(collectedMessages, collectMessages());
  notifyProgress("Finished loading older messages.");
}

function getChatTitle(chatLabelOverride = "") {
  if (chatLabelOverride && chatLabelOverride.trim()) {
    return chatLabelOverride.trim();
  }

  const currentListLabel = getCurrentChatListLabel();

  if (currentListLabel) {
    return currentListLabel;
  }

  const pageTitle = cleanPageTitle(document.title);

  if (pageTitle) {
    return pageTitle;
  }

  const participantTitle = getParticipantBasedTitle();

  if (participantTitle) {
    return participantTitle;
  }

  const titleSelectors = [
    "[data-tid='chat-header-title']",
    "[data-tid='thread-header-chat-title']",
    "[role='heading']"
  ];

  for (const selector of titleSelectors) {
    const element = document.querySelector(selector);

    if (element?.textContent?.trim()) {
      return element.textContent.trim();
    }
  }

  return "teams-chat";
}

function getCurrentChatListLabel() {
  const selectedSelectors = [
    "[role='treeitem'][data-item-type='chat'][aria-selected='true']",
    "[role='treeitem'][data-testid='list-item'][data-item-type='chat'][aria-selected='true']",
    "[role='treeitem'][data-item-type='chat'][aria-current='true']"
  ];

  for (const selector of selectedSelectors) {
    const entry = document.querySelector(selector);

    if (!entry) {
      continue;
    }

    const label = getChatListEntryLabel(entry);

    if (label) {
      return label;
    }
  }

  const pageTitle = cleanPageTitle(document.title);

  if (pageTitle) {
    const matchingEntry = getChatListEntries().find((entry) => {
      const label = getChatListEntryLabel(entry);
      return label && titlesProbablyMatch(label, pageTitle);
    });

    if (matchingEntry) {
      return getChatListEntryLabel(matchingEntry);
    }
  }

  return "";
}

function getCurrentChatListOrder(chatLabel = "") {
  const entries = getChatListEntries().filter((entry) => looksLikeChatListEntry(entry));
  const targetLabel = chatLabel || getCurrentChatListLabel();

  if (!targetLabel) {
    return -1;
  }

  const matchIndex = entries.findIndex((entry) => {
    const label = getChatListEntryLabel(entry);
    return label && titlesProbablyMatch(label, targetLabel);
  });

  return matchIndex >= 0 ? matchIndex : -1;
}

function looksLikeChatListEntry(element) {
  if (!element) {
    return false;
  }

  if (
    element.matches(
      "[role='treeitem'][data-testid='list-item'][data-item-type='chat'], [role='treeitem'][data-item-type='chat']"
    )
  ) {
    return true;
  }

  const text = normalizeText(element.innerText || element.textContent || "");

  if (!text || text.length > 300) {
    return false;
  }

  const positiveSignals = [
    element.getAttribute("data-tid") || "",
    element.getAttribute("data-item-type") || "",
    element.getAttribute("data-testid") || "",
    element.getAttribute("aria-label") || "",
    element.getAttribute("role") || ""
  ].join(" ");

  if (/chat-list|chat-item|conversation|thread|list-item|data-item-type/i.test(positiveSignals)) {
    return true;
  }

  if (element.closest("[data-tid='chat-list-pane'], [aria-label*='Chats'], [aria-label*='Chat list']")) {
    return true;
  }

  return false;
}

function getChatListEntryLabel(element) {
  const titleElement = element.querySelector("[id^='title-chat-list-item_']");

  if (titleElement?.textContent?.trim()) {
    return cleanChatListLabel(titleElement.textContent);
  }

  const labelSources = [
    element.getAttribute("aria-label") || "",
    normalizeText(element.innerText || element.textContent || "")
  ].filter(Boolean);

  for (const source of labelSources) {
    const label = cleanChatListLabel(source);

    if (label) {
      return label;
    }
  }

  return "";
}

function cleanChatListLabel(value) {
  const lines = normalizeText(value)
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean)
    .filter((line) => !shouldIgnoreChatListLine(line));

  return lines[0] || "";
}

function shouldIgnoreChatListLine(value) {
  const ignored = [
    "Unread",
    "Chat",
    "Chats",
    "Meeting chat",
    "Group chat",
    "Has context menu",
    "Last message",
    "Unread message",
    "Draft",
    "Muted",
    "Private",
    "Shared",
    "Meeting in progress",
    "Personal at mention",
    "Everyone at mention",
    "Important",
    "Urgent",
    "See more"
  ];

  if (ignored.includes(value)) {
    return true;
  }

  if (/^\d+\s+new notification$/i.test(value)) {
    return true;
  }

  if (/^\d+$/.test(value)) {
    return true;
  }

  return false;
}

function findChatListEntry(chatLabel) {
  const chatItems = getChatListEntries();

  for (const item of chatItems) {
    if (!looksLikeChatListEntry(item)) {
      continue;
    }

    if (getChatListEntryLabel(item) === chatLabel) {
      return item;
    }
  }

  return null;
}

function getChatListEntries() {
  const selectors = [
    "[role='treeitem'][data-testid='list-item'][data-item-type='chat']",
    "[role='treeitem'][data-item-type='chat']",
    "[data-testid='list-item'][data-item-type='chat']"
  ];

  const entries = selectors.flatMap((selector) => Array.from(document.querySelectorAll(selector)));

  if (entries.length) {
    return dedupeElements(entries);
  }

  return Array.from(document.querySelectorAll("[role='treeitem'], [role='listitem'], [data-tid]"));
}

function dedupeElements(elements) {
  return Array.from(new Set(elements));
}

async function waitForChatToOpen(expectedLabel) {
  const maxAttempts = 30;

  for (let attempt = 0; attempt < maxAttempts; attempt += 1) {
    const currentTitle = getChatTitle();

    if (currentTitle && titlesProbablyMatch(currentTitle, expectedLabel)) {
      await wait(600);
      return;
    }

    await wait(300);
  }

  await wait(800);
}

function getParticipantBasedTitle() {
  const headerRoot = getHeaderRoot();

  if (!headerRoot) {
    return "";
  }

  const rawText = normalizeText(headerRoot.innerText || headerRoot.textContent || "");

  if (!rawText) {
    return "";
  }

  const lines = rawText
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean)
    .filter((line) => !shouldIgnoreHeaderLine(line));

  if (!lines.length) {
    return "";
  }

  const candidate = lines[0];

  if (looksLikePersonList(candidate)) {
    return candidate;
  }

  return lines.slice(0, 3).join(", ");
}

function getHeaderRoot() {
  const selectors = [
    "[data-tid='chat-header']",
    "[data-tid='thread-header']",
    "header",
    "[role='banner']"
  ];

  for (const selector of selectors) {
    const element = document.querySelector(selector);

    if (element) {
      return element;
    }
  }

  return null;
}

function titlesProbablyMatch(left, right) {
  const normalize = (value) =>
    String(value || "")
      .toLowerCase()
      .replace(/\s+/g, " ")
      .replace(/[^\p{L}\p{M}\p{N},+ -]+/gu, "")
      .trim();

  const leftNormalized = normalize(left);
  const rightNormalized = normalize(right);

  return (
    leftNormalized === rightNormalized ||
    leftNormalized.includes(rightNormalized) ||
    rightNormalized.includes(leftNormalized)
  );
}

function shouldIgnoreHeaderLine(value) {
  const ignored = [
    "Chat",
    "Open details",
    "Start a call",
    "Video call",
    "Audio call",
    "More",
    "Search",
    "Pinned messages",
    "People",
    "Manage chat",
    "Add people"
  ];

  if (ignored.includes(value)) {
    return true;
  }

  if (/^\d+\s+(members?|participants?)$/i.test(value)) {
    return true;
  }

  return false;
}

function looksLikePersonList(value) {
  if (!value) {
    return false;
  }

  if (value.includes(",") || value.includes("+")) {
    return true;
  }

  return /\b[A-Z][\p{L}\p{M}'-]+,\s+[A-Z][\p{L}\p{M}' -]+/u.test(value);
}

function cleanPageTitle(value) {
  return normalizeText(value || "")
    .replace(/\|\s*Microsoft Teams.*$/i, "")
    .replace(/\|\s*Teams.*$/i, "")
    .trim();
}

function collectMessages() {
  const chatRoot = getChatRoot();
  const textBodyNodes = Array.from(
    chatRoot.querySelectorAll(
      [
        "[data-tid='messageBodyContent']",
        "[data-tid='chat-pane-message']",
        "[data-tid='threadBodyContent']",
        "[data-tid='richTextEditor'] div[dir='auto']"
      ].join(", ")
    )
  );
  const imageNodes = Array.from(chatRoot.querySelectorAll("img"))
    .filter((imageElement) => !shouldIgnoreInlineImage(imageElement))
    .map((imageElement) => findMessageContainer(imageElement));
  const bodyNodes = dedupeElements([...textBodyNodes, ...imageNodes]);
  const seenKeys = new Set();
  const messages = [];
  let lastKnownAuthor = "";

  for (const bodyNode of bodyNodes) {
    const record = extractMessageRecord(bodyNode, lastKnownAuthor);

    if (!record) {
      continue;
    }

    if (record.author && record.author !== "Unknown") {
      lastKnownAuthor = record.author;
    }

    const recordKey = `${record.author}|${record.timestamp}|${record.text}`;

    if (seenKeys.has(recordKey)) {
      continue;
    }

    seenKeys.add(recordKey);
    messages.push(record);
  }

  return messages;
}

function appendUniqueMessages(targetMessages, sourceMessages) {
  const merged = dedupeMessageRecords([...(targetMessages || []), ...(sourceMessages || [])]);
  targetMessages.length = 0;
  targetMessages.push(...merged);
}

function dedupeMessageRecords(messages) {
  const seen = new Set();
  const deduped = [];

  for (const message of messages) {
    const key = buildMessageRecordKey(message);

    if (!key || seen.has(key)) {
      continue;
    }

    seen.add(key);
    deduped.push(message);
  }

  return deduped;
}

function buildMessageRecordKey(message) {
  if (!message) {
    return "";
  }

  if (message.id) {
    return `id:${message.id}`;
  }

  const imageKey = (message.images || [])
    .map((image) => image.sourceUrl || image.embeddedDataUrl || image.localPath || "")
    .filter(Boolean)
    .join("|");

  return [
    message.author || "",
    message.timestamp || "",
    message.text || "",
    imageKey
  ].join("|");
}

function getChatRoot() {
  const rootSelectors = [
    "[data-tid='message-pane-list-runway']",
    "[data-tid='chat-pane-content']",
    "[data-tid='message-pane-body']",
    "[role='main']"
  ];

  for (const selector of rootSelectors) {
    const element = document.querySelector(selector);

    if (element) {
      return element;
    }
  }

  return document.body;
}

function getMessageScrollContainer() {
  const rootSelectors = [
    "[data-tid='message-pane-list-runway']",
    "[data-tid='chat-pane-content']",
    "[data-tid='message-pane-body']"
  ];

  for (const selector of rootSelectors) {
    const root = document.querySelector(selector);

    if (!root) {
      continue;
    }

    const scrollable = findScrollableAncestor(root) || findScrollableDescendant(root);

    if (scrollable) {
      return scrollable;
    }
  }

  return findScrollableDescendant(document.querySelector("[role='main']")) || null;
}

function extractMessageRecord(bodyNode, lastKnownAuthor) {
  const container = findMessageContainer(bodyNode);
  const metadataRoot = findMetadataRoot(container);
  const text = normalizeText(bodyNode.innerText || bodyNode.textContent || "");
  const images = collectInlineImages(container);

  if ((!text && !images.length) || (text && shouldIgnoreText(text, container))) {
    return null;
  }

  let author = findNestedText(metadataRoot, [
    "[data-tid='message-author-name']",
    "[data-tid='message-author']",
    "[data-tid='author']",
    "[data-tid='threadBodyDisplayName']",
    "[data-tid='senderDisplayName']",
    "[data-tid='messageAuthor']",
    "[data-tid*='author']",
    "[data-tid*='display-name']"
  ]);

  if (!author) {
    author = extractAuthorFromLabels(metadataRoot, text);
  }

  if (!author) {
    author = lastKnownAuthor || "Unknown";
  }

  const timestamp =
    findTimeText(metadataRoot) ||
    findNestedText(metadataRoot, [
    "time",
    "[data-tid='message-timestamp']",
    "[data-tid*='timestamp']",
    "[data-tid*='time']"
    ]) ||
    extractTimestampFromLabels(metadataRoot) ||
    "";

  const messageId =
    container.getAttribute("id") ||
    container.getAttribute("data-item-id") ||
    container.getAttribute("data-message-id") ||
    bodyNode.getAttribute("data-message-id") ||
    "";

  return {
    id: messageId,
    author,
    timestamp,
    text,
    images
  };
}

function findMessageContainer(element) {
  const selectors = [
    "[role='listitem']",
    "[data-item-id]",
    "[data-message-id]",
    "[data-tid*='message']"
  ];

  for (const selector of selectors) {
    const container = element.closest(selector);

    if (container) {
      return container;
    }
  }

  return element.parentElement || element;
}

function findScrollableAncestor(element) {
  let current = element;

  while (current && current !== document.body) {
    if (isScrollable(current)) {
      return current;
    }

    current = current.parentElement;
  }

  return null;
}

function findScrollableDescendant(root) {
  if (!root) {
    return null;
  }

  const elements = [root, ...root.querySelectorAll("*")];

  for (const element of elements) {
    if (isScrollable(element)) {
      return element;
    }
  }

  return null;
}

function isScrollable(element) {
  if (!element) {
    return false;
  }

  const style = window.getComputedStyle(element);
  const overflowY = style.overflowY;
  return (
    (overflowY === "auto" || overflowY === "scroll") &&
    element.scrollHeight > element.clientHeight + 100
  );
}

function findMetadataRoot(container) {
  const ancestorCandidates = [];
  let current = container;

  for (let index = 0; current && index < 5; index += 1) {
    ancestorCandidates.push(current);
    current = current.parentElement;
  }

  for (const candidate of ancestorCandidates) {
    if (candidate.querySelector("time")) {
      return candidate;
    }

    if (candidate.innerText && candidate.innerText.length > container.innerText.length + 10) {
      return candidate;
    }
  }

  return container;
}

function findNestedText(element, selectors) {
  for (const selector of selectors) {
    const nested = element.querySelector(selector);

    if (nested?.textContent?.trim()) {
      return nested.textContent.trim();
    }
  }

  return "";
}

function findTimeText(element) {
  const timeElement = element.querySelector("time");

  if (!timeElement) {
    return "";
  }

  return (
    timeElement.getAttribute("datetime") ||
    timeElement.getAttribute("aria-label") ||
    timeElement.getAttribute("title") ||
    timeElement.textContent ||
    ""
  ).trim();
}

function normalizeText(value) {
  return value.replace(/\r/g, "").replace(/\n{3,}/g, "\n\n").trim();
}

function collectInlineImages(bodyNode) {
  const imageElements = Array.from(bodyNode.querySelectorAll("img"));
  const images = [];
  const seen = new Set();

  for (const imageElement of imageElements) {
    if (shouldIgnoreInlineImage(imageElement)) {
      continue;
    }

    const sourceUrl =
      imageElement.currentSrc ||
      imageElement.getAttribute("src") ||
      imageElement.getAttribute("data-src") ||
      "";

    if (!sourceUrl || seen.has(sourceUrl)) {
      continue;
    }

    seen.add(sourceUrl);
    images.push({
      sourceUrl,
      alt: (imageElement.getAttribute("alt") || "").trim(),
      embeddedDataUrl: sourceUrl.startsWith("data:") ? sourceUrl : ""
    });
  }

  return images;
}

async function enrichMessagesWithEmbeddedImages(messages) {
  for (const message of messages) {
    for (const image of message.images || []) {
      if (image.embeddedDataUrl || !image.sourceUrl || !image.sourceUrl.startsWith("blob:")) {
        continue;
      }

      image.embeddedDataUrl = await resolveBlobImageToDataUrl(image.sourceUrl);
    }
  }
}

async function resolveBlobImageToDataUrl(sourceUrl) {
  try {
    const response = await fetch(sourceUrl);

    if (!response.ok) {
      return "";
    }

    const blob = await response.blob();
    return await blobToDataUrl(blob);
  } catch (error) {
    return "";
  }
}

function blobToDataUrl(blob) {
  return new Promise((resolve) => {
    const reader = new FileReader();
    reader.onloadend = () => resolve(String(reader.result || ""));
    reader.readAsDataURL(blob);
  });
}

function shouldIgnoreInlineImage(imageElement) {
  if (!imageElement) {
    return true;
  }

  if (
    imageElement.matches(
      ".fui-Avatar__image, [role='presentation'][aria-hidden='true'], [data-tid*='avatar'] img"
    )
  ) {
    return true;
  }

  if (imageElement.closest("[data-tid*='avatar'], .fui-Avatar, [aria-label*='Profile picture']")) {
    return true;
  }

  if (imageElement.closest("[data-tid*='reaction'], [aria-label*='reaction']")) {
    return true;
  }

  const sourceUrl =
    imageElement.currentSrc ||
    imageElement.getAttribute("src") ||
    imageElement.getAttribute("data-src") ||
    "";

  if (!sourceUrl) {
    return true;
  }

  if (
    /statics\.teams\.cdn\.office\.net\/evergreen-assets\/personal-expressions/i.test(sourceUrl) ||
    /media\d*\.giphy\.com/i.test(sourceUrl) ||
    /giphy_s\.gif/i.test(sourceUrl)
  ) {
    return true;
  }

  const width = Number(imageElement.getAttribute("width") || 0);
  const height = Number(imageElement.getAttribute("height") || 0);

  if ((width > 0 && width <= 48) || (height > 0 && height <= 48)) {
    return true;
  }

  const altText = (imageElement.getAttribute("alt") || "").toLowerCase();

  if (
    altText.includes("emoji") ||
    altText.includes("sticker") ||
    altText.includes("profile picture")
  ) {
    return true;
  }

  return false;
}

function shouldIgnoreText(text, container) {
  if (!text) {
    return true;
  }

  if (isInsideIgnoredRegion(container)) {
    return true;
  }

  const ignoreExact = new Set([
    "Apply and restart",
    "Has context menu",
    "See more",
    "See all channels",
    "See all your teams",
    "Unread",
    "Chats",
    "Chat",
    "Activity",
    "Calendar",
    "Calls",
    "OneDrive",
    "Apps",
    "Copilot"
  ]);

  if (ignoreExact.has(text)) {
    return true;
  }

  const ignoreContains = [
    "Language changes detected",
    "Please restart Teams",
    "You can't send messages because you are not a member of the chat.",
    "new notification",
    "Unread message",
    "Last message",
    "Meeting in progress",
    "Temporarily shown"
  ];

  if (ignoreContains.some((phrase) => text.includes(phrase))) {
    return true;
  }

  if (looksLikeSidebarText(text)) {
    return true;
  }

  if (looksLikeReactionText(text, container)) {
    return true;
  }

  return false;
}

function isInsideIgnoredRegion(element) {
  const ignoredAncestors = [
    "nav",
    "aside",
    "[role='navigation']",
    "[aria-label='App bar']",
    "[data-tid='app-layout-area--left-rail']",
    "[data-tid='chat-list-pane']",
    "[data-tid='list-header']"
  ];

  return ignoredAncestors.some((selector) => element.closest(selector));
}

function looksLikeSidebarText(text) {
  const sidebarSignals = [
    "Activity",
    "Calendar",
    "Calls",
    "OneDrive",
    "Apps",
    "Chats",
    "Meeting chats",
    "Discover",
    "Mentions",
    "Followed threads",
    "Favourites",
    "Teams and channels"
  ];

  const matches = sidebarSignals.filter((signal) => text.includes(signal)).length;
  return matches >= 3;
}

function looksLikeReactionText(text, container) {
  if (container.closest("[data-tid*='reaction']")) {
    return true;
  }

  const reactionSignals = [
    "Liked",
    "Loved",
    "Laugh",
    "Celebrate",
    "Surprised",
    "Sad",
    "Angry",
    "reacted"
  ];

  return reactionSignals.some((signal) => text === signal || text.startsWith(`${signal} `));
}

function extractAuthorFromLabels(element, messageText) {
  const labelSources = [
    element.getAttribute("aria-label") || "",
    element.getAttribute("data-tid") || "",
    findNearbyLabelText(element)
  ].filter(Boolean);

  for (const source of labelSources) {
    const author = parseAuthorFromSource(source, messageText);

    if (author) {
      return author;
    }
  }

  return "";
}

function extractTimestampFromLabels(element) {
  const labelSources = [
    element.getAttribute("aria-label") || "",
    findNearbyLabelText(element)
  ].filter(Boolean);

  for (const source of labelSources) {
    const timestamp = parseTimestampFromSource(source);

    if (timestamp) {
      return timestamp;
    }
  }

  return "";
}

function findNearbyLabelText(element) {
  const nearby = [
    element,
    element.previousElementSibling,
    element.parentElement,
    element.parentElement?.previousElementSibling
  ].filter(Boolean);

  return nearby
    .map((node) => {
      return [
        node.getAttribute?.("aria-label") || "",
        node.getAttribute?.("title") || "",
        node.innerText || ""
      ]
        .filter(Boolean)
        .join("\n");
    })
    .join("\n");
}

function parseAuthorFromSource(source, messageText) {
  const cleaned = normalizeText(source).replace(messageText, "").trim();

  if (!cleaned) {
    return "";
  }

  const lines = cleaned
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean);

  for (const line of lines) {
    if (looksLikeTimestamp(line) || shouldIgnoreCandidateAuthor(line)) {
      continue;
    }

    if (/^[A-Z][\p{L}\p{M}'.,\- ]{1,80}$/u.test(line)) {
      return line;
    }
  }

  return "";
}

function parseTimestampFromSource(source) {
  const lines = normalizeText(source)
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean);

  for (const line of lines) {
    if (looksLikeTimestamp(line)) {
      return line;
    }
  }

  return "";
}

function looksLikeTimestamp(value) {
  const patterns = [
    /\b\d{1,2}:\d{2}\b/,
    /\b(?:AM|PM)\b/i,
    /\b(?:Mon|Tue|Wed|Thu|Fri|Sat|Sun)\b/i,
    /\b(?:January|February|March|April|May|June|July|August|September|October|November|December)\b/i,
    /\b(?:today|yesterday)\b/i
  ];

  return patterns.some((pattern) => pattern.test(value));
}

function shouldIgnoreCandidateAuthor(value) {
  const badFragments = [
    "Edited",
    "Reply",
    "reacted",
    "Like",
    "Love",
    "Laugh",
    "Celebrate",
    "Unread"
  ];

  return badFragments.some((fragment) => value.includes(fragment));
}

function dispatchScrollEvents(element) {
  element.dispatchEvent(new Event("scroll", { bubbles: true }));
  window.dispatchEvent(new Event("scroll"));
}

function clickElement(element) {
  element.dispatchEvent(new MouseEvent("mousedown", { bubbles: true }));
  element.dispatchEvent(new MouseEvent("mouseup", { bubbles: true }));
  element.click();
}

function notifyProgress(status) {
  try {
    chrome.runtime.sendMessage({
      type: "ARCHIVE_PROGRESS",
      status
    });
  } catch (error) {
    // Ignore progress messaging failures during archive collection.
  }
}

function wait(ms) {
  return new Promise((resolve) => {
    window.setTimeout(resolve, ms);
  });
}
