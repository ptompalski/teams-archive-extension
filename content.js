chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (message?.type === "LIST_SELECTED_GROUP_CHATS") {
    listSelectedGroupChats(message.groupIds)
      .then((payload) => sendResponse({ ok: true, payload }))
      .catch((error) => sendResponse({ ok: false, error: error.message }));
    return true;
  }

  if (message?.type === "LIST_CHAT_GROUPS") {
    sendResponse({ ok: true, payload: listChatGroups() });
    return false;
  }

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
    openChatByLabel(message.chatLabel, message)
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
  const chatTitle = normalizeChatTitle(getChatTitle(chatLabelOverride));
  const chatOrder = getCurrentChatListOrder(chatTitle);
  const collectedMessages = [];
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
    if (!looksLikeChatListEntry(item) || item.closest("[data-item-type='teams-and-channels']")) {
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

function getChatGroupElements() {
  return Array.from(document.querySelectorAll(
    "[data-conversation-folder='true']:is([data-item-type='custom-folder'], [data-item-type='chats'])"
  ));
}

function getChatGroupHeader(group) {
  return group.querySelector("[data-testid='conversation-folder-header']");
}

function getChatGroupId(group) {
  return getChatGroupHeader(group)?.id || group.getAttribute("data-fui-tree-item-value") || "";
}

function listChatGroups() {
  return getChatGroupElements().map((group) => {
    const header = getChatGroupHeader(group);
    const title = header?.querySelector(".fui-TreeItemLayout__main") || header;

    return {
      id: getChatGroupId(group),
      label: normalizeText(title?.textContent || ""),
      expanded: group.getAttribute("aria-expanded") === "true"
    };
  }).filter((group) => group.id && group.label);
}

function findChatGroup(groupId) {
  return getChatGroupElements().find((group) => getChatGroupId(group) === groupId) || null;
}

function getChatListEntryId(entry) {
  const title = entry.querySelector("[id^='title-chat-list-item_'], [id^='title-channel-list-item-']");
  return title?.id.replace(/^title-(?:chat-list-item_|channel-list-item-)/, "") || "";
}

function getChatsInGroup(group) {
  const entries = Array.from(group.querySelectorAll(
    "[role='treeitem']:is([data-item-type='chat'], [data-item-type='muted-chat'], [data-item-type='channel'])"
  )).filter((entry) => entry.closest("[data-conversation-folder='true']") === group);

  return entries.map((entry, order) => ({
    id: getChatListEntryId(entry),
    label: getChatListEntryLabel(entry),
    groupId: getChatGroupId(group),
    order
  })).filter((chat) => chat.id && chat.label);
}

function getChatGroupScrollContainer(group) {
  for (let ancestor = group; ancestor; ancestor = ancestor.parentElement) {
    const overflow = window.getComputedStyle(ancestor).overflowY;

    if (ancestor.clientHeight > 0 && /auto|scroll/.test(overflow)) {
      return ancestor;
    }

    if (ancestor.matches("[data-tid='app-layout-area--mid-nav']")) {
      break;
    }
  }

  return null;
}

async function expandChatGroup(groupId) {
  let group = findChatGroup(groupId);

  if (!group) {
    throw new Error("A selected chat group is no longer available. Reload the group list and try again.");
  }

  if (group.getAttribute("aria-expanded") !== "true") {
    const header = getChatGroupHeader(group);

    if (!header) {
      throw new Error("Could not find the selected group's expansion control.");
    }

    header.scrollIntoView({ block: "nearest" });
    clickElement(header);

    for (let attempt = 0; attempt < 30; attempt += 1) {
      await wait(200);
      group = findChatGroup(groupId);

      if (group?.getAttribute("aria-expanded") === "true") {
        break;
      }
    }
  }

  if (group?.getAttribute("aria-expanded") !== "true") {
    throw new Error("Could not expand the selected chat group.");
  }

  await wait(500);
  return group;
}

async function scanChatGroup(groupId, targetChatId = "") {
  let group = await expandChatGroup(groupId);
  const scrollContainer = getChatGroupScrollContainer(group);
  const originalScrollTop = scrollContainer?.scrollTop || 0;
  const chats = new Map();
  let stableRounds = 0;
  let loadingRounds = 0;
  let requestedPage = "";

  try {
    if (scrollContainer) {
      scrollContainer.scrollTop = 0;
      dispatchScrollEvents(scrollContainer);
      await wait(500);
    }

    for (let round = 0; round < 200; round += 1) {
      group = findChatGroup(groupId);

      if (!group || group.getAttribute("aria-expanded") !== "true") {
        throw new Error("The selected chat group disappeared or collapsed while reading it.");
      }

      const previousCount = chats.size;
      const visibleChats = getChatsInGroup(group);

      for (const chat of visibleChats) {
        if (!chats.has(chat.id)) {
          chats.set(chat.id, chat);
        }
      }

      if (targetChatId) {
        const entry = Array.from(group.querySelectorAll(
          "[role='treeitem']:is([data-item-type='chat'], [data-item-type='muted-chat'], [data-item-type='channel'])"
        )).find((item) => getChatListEntryId(item) === targetChatId);

        if (entry) {
          return { chats: Array.from(chats.values()), entry };
        }
      }

      const more = Array.from(group.querySelectorAll("[data-testid='load-next-page-button']"))
        .find((button) => button.closest("[data-conversation-folder='true']") === group);
      const busy = group.matches("[aria-busy='true']") ||
        group.querySelector("[aria-busy='true'], [role='progressbar']");

      if (more || busy) {
        const pageKey = JSON.stringify(Array.from(group.querySelectorAll("[role='treeitem']"))
          .map((entry) => entry.getAttribute("data-fui-tree-item-value") ||
            entry.querySelector("[id^='title-chat-list-item_']")?.id || entry.textContent));

        if (more && !busy && pageKey !== requestedPage &&
            more.getAttribute("aria-disabled") !== "true" && !more.hasAttribute("disabled")) {
          more.scrollIntoView({ block: "nearest" });
          clickElement(more);
          requestedPage = pageKey;
          loadingRounds = 0;

          if (scrollContainer) {
            scrollContainer.scrollTop = 0;
            dispatchScrollEvents(scrollContainer);
          }
        } else {
          loadingRounds += 1;
        }

        if (loadingRounds >= 20) {
          throw new Error("Teams stopped loading chats in a selected group. No partial group list was archived.");
        }

        stableRounds = 0;
        await wait(500);
        continue;
      }

      loadingRounds = 0;
      const atBottom = !scrollContainer ||
        scrollContainer.scrollTop + scrollContainer.clientHeight >= scrollContainer.scrollHeight - 2;

      if (atBottom && chats.size === previousCount) {
        stableRounds += 1;

        if (stableRounds >= 4) {
          if (targetChatId) {
            throw new Error("The requested chat could not be found in its selected group.");
          }

          return { chats: Array.from(chats.values()), entry: null };
        }
      } else {
        stableRounds = 0;
      }

      if (scrollContainer && !atBottom) {
        scrollContainer.scrollTop += Math.max(100, scrollContainer.clientHeight * 0.8);
        dispatchScrollEvents(scrollContainer);
      }

      await wait(500);
    }

    throw new Error("The selected group could not be fully read within the safety limit. No partial group list was archived.");
  } finally {
    if (scrollContainer && !targetChatId) {
      scrollContainer.scrollTop = originalScrollTop;
      dispatchScrollEvents(scrollContainer);
    }
  }
}

async function listSelectedGroupChats(groupIds) {
  if (!Array.isArray(groupIds) || !groupIds.length ||
      groupIds.some((groupId) => typeof groupId !== "string" || !groupId)) {
    throw new Error("Select at least one chat group.");
  }

  const chats = new Map();

  for (const groupId of new Set(groupIds)) {
    const label = listChatGroups().find((group) => group.id === groupId)?.label;
    notifyProgress(`Reading chat group: ${label || groupId}...`);
    const result = await scanChatGroup(groupId);

    for (const chat of result.chats) {
      if (!chats.has(chat.id)) {
        chats.set(chat.id, { ...chat, order: chats.size });
      }
    }
  }

  return Array.from(chats.values());
}

async function openChatByLabel(chatLabel, { chatId = "", groupId = "" } = {}) {
  if (!chatLabel) {
    throw new Error("No chat label was provided.");
  }

  const chatEntry = groupId && chatId
    ? (await scanChatGroup(groupId, chatId)).entry
    : findChatListEntry(chatLabel);

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
  let isChannel = Boolean(getChatRoot().querySelector("[data-tid='channel-pane-message']"));
  const replyCounts = new Map();
  await expandChannelMessages(collectedMessages, replyCounts);
  let scrollContainer = getMessageScrollContainer();

  if (!scrollContainer && (isChannel || isMessagePaneBusy())) {
    let stableRounds = 0;

    for (let round = 0; round < 200 && !scrollContainer; round += 1) {
      const previousCount = collectedMessages.length;
      await wait(900);
      await expandChannelMessages(collectedMessages, replyCounts);
      scrollContainer = getMessageScrollContainer();
      isChannel = isChannel || Boolean(getChatRoot().querySelector("[data-tid='channel-pane-message']"));
      stableRounds = !isMessagePaneBusy() && collectedMessages.length === previousCount ? stableRounds + 1 : 0;

      if (stableRounds >= 3) {
        break;
      }
    }

    if (!scrollContainer && stableRounds < 3) {
      throw new Error("Channel history reached its safety limit. No partial channel snapshot was saved.");
    }
  }

  if (!scrollContainer) {
    appendUniqueMessages(collectedMessages, collectMessages());
    assertChannelRepliesComplete(collectedMessages, replyCounts);
    return;
  }

  let previousScrollTop = scrollContainer.scrollTop;
  let previousScrollHeight = scrollContainer.scrollHeight;
  let previousMessageCount = collectMessages().length;
  let stableRounds = 0;
  const maxRounds = 80;

  for (let round = 0; round < maxRounds; round += 1) {
    scrollContainer = getMessageScrollContainer() || scrollContainer;
    notifyProgress(`Loading older messages (${round + 1}/${maxRounds})...`);
    appendUniqueMessages(collectedMessages, collectMessages());

    if (scrollContainer.scrollTop <= 0) {
      scrollContainer.scrollTop = 0;
    } else {
      const step = isChannel ? Math.max(100, scrollContainer.clientHeight * 0.8)
        : Math.max(600, scrollContainer.clientHeight);
      const nextTop = Math.max(0, scrollContainer.scrollTop - step);
      scrollContainer.scrollTop = nextTop;
    }

    dispatchScrollEvents(scrollContainer);
    await wait(900);
    await expandChannelMessages(collectedMessages, replyCounts);
    scrollContainer = getMessageScrollContainer() || scrollContainer;

    const currentMessageCount = collectMessages().length;
    const currentScrollTop = scrollContainer.scrollTop;
    const currentScrollHeight = scrollContainer.scrollHeight;
    const noScrollMovement = currentScrollTop === previousScrollTop;
    const noGrowth = currentScrollHeight === previousScrollHeight;
    const noNewMessages = currentMessageCount === previousMessageCount;

    const busy = isChannel && isMessagePaneBusy();

    if (!busy && ((currentScrollTop <= 0 && noNewMessages) || (noScrollMovement && noGrowth && noNewMessages))) {
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

  if (isChannel) {
    if (stableRounds < 3) {
      throw new Error("Channel history reached its safety limit. No partial channel snapshot was saved.");
    }

    await loadChannelHistoryDown(collectedMessages, replyCounts);
  }

  appendUniqueMessages(collectedMessages, collectMessages());
  assertChannelRepliesComplete(collectedMessages, replyCounts);
  notifyProgress("Finished loading older messages.");
}

async function loadChannelHistoryDown(collectedMessages, replyCounts) {
  let stableRounds = 0;

  for (let round = 0; round < 200; round += 1) {
    notifyProgress(`Reading channel history (${round + 1}/200)...`);
    const previousCount = collectedMessages.length;
    let scrollable = getMessageScrollContainer();

    if (scrollable) {
      scrollable.scrollTop += Math.max(100, scrollable.clientHeight * 0.8);
      dispatchScrollEvents(scrollable);
      await wait(900);
    }

    await expandChannelMessages(collectedMessages, replyCounts);
    scrollable = getMessageScrollContainer();
    const atBottom = !scrollable ||
      scrollable.scrollTop + scrollable.clientHeight >= scrollable.scrollHeight - 2;
    const busy = isMessagePaneBusy();

    if (atBottom && !busy && collectedMessages.length === previousCount) {
      stableRounds += 1;

      if (stableRounds >= 3) {
        return;
      }
    } else {
      stableRounds = 0;
    }

    if (!scrollable) {
      await wait(900);
    }
  }

  throw new Error("Channel history reached its safety limit. No partial channel snapshot was saved.");
}

function isMessagePaneBusy() {
  const root = getChatRoot();
  const selector = "[aria-busy='true'], [role='progressbar']";
  return root.matches(selector) || Boolean(root.querySelector(selector));
}

async function expandChannelBodies() {
  for (let round = 0; round < 200; round += 1) {
    const control = Array.from(getChatRoot().querySelectorAll(
      "[aria-controls^='see-more-content-'][aria-expanded='false']"
    )).find((button) => document.getElementById(button.getAttribute("aria-controls"))
      ?.querySelector("[data-tid='message-body']"));

    if (!control) {
      return;
    }

    clickElement(control);

    for (let attempt = 0; attempt < 20 && control.isConnected &&
        control.getAttribute("aria-expanded") === "false"; attempt += 1) {
      await wait(200);
    }

    if (control.isConnected && control.getAttribute("aria-expanded") === "false") {
      throw new Error("Teams did not expand a channel message. No partial channel snapshot was saved.");
    }
  }

  throw new Error("Channel message expansion reached its safety limit. No partial channel snapshot was saved.");
}

async function expandChannelMessages(collectedMessages, replyCounts) {
  await expandChannelBodies();
  appendUniqueMessages(collectedMessages, collectMessages());
  const channelLabel = getChatTitle();
  const threads = Array.from(getChatRoot().querySelectorAll("[data-tid='response-summary-button']"))
    .map((summary) => {
      const elementId = summary.closest("[data-tid='channel-pane-message']")?.id;
      const threadId = elementId?.replace(/^reply-chain-summary-/, "");
      const surface = summary.closest("[data-tid='response-surface']");
      const countText = `${surface?.getAttribute("aria-label") || ""} ${summary.textContent || ""}`;
      const expected = Number(countText.match(/(\d+)\s+repl(?:y|ies)/i)?.[1]);

      if (!threadId || !Number.isFinite(expected)) {
        throw new Error("Could not determine a channel thread's reply count. No partial channel snapshot was saved.");
      }

      replyCounts.set(threadId, expected);
      return { elementId, threadId, expected };
    });

  for (const { elementId, threadId, expected } of threads) {
    const thread = document.getElementById(elementId);
    const summary = thread?.querySelector("[data-tid='response-summary-button']");

    if (!summary || !getChatRoot().contains(thread)) {
      continue;
    }

    const readReplies = () => collectMessages(threadId)
      .filter((message) => message.threadId === threadId && message.id !== threadId);

    if (new Set(collectedMessages.filter((message) => message.threadId === threadId &&
      message.id !== threadId).map((message) => message.id)).size >= expected) {
      continue;
    }

    notifyProgress(`Loading ${expected} replies in a channel thread...`);
    summary.scrollIntoView({ block: "nearest" });
    clickElement(summary);
    let complete = false;
    let scrollDirection = -1;
    let boundaryRounds = 0;

    for (let attempt = 0; attempt < 100; attempt += 1) {
      await wait(300);
      await expandChannelBodies();
      const replies = readReplies();
      appendUniqueMessages(collectedMessages, replies);

      if (new Set(collectedMessages.filter((message) => message.threadId === threadId &&
          message.id !== threadId).map((message) => message.id)).size >= expected) {
        complete = true;
        break;
      }

      const currentThread = document.getElementById(`reply-chain-summary-${threadId}`);
      const replyRoot = document.querySelector("[data-tid='channel-replies-runway']") ||
        currentThread?.querySelector("[data-tid='response-surface']") || getChatRoot();
      const scrollable = findScrollableDescendant(replyRoot) || findScrollableAncestor(replyRoot);

      if (scrollable) {
        const atBoundary = scrollDirection < 0 ? scrollable.scrollTop <= 0
          : scrollable.scrollTop + scrollable.clientHeight >= scrollable.scrollHeight - 2;
        boundaryRounds = atBoundary && !isMessagePaneBusy() ? boundaryRounds + 1 : 0;

        if (boundaryRounds >= 3) {
          scrollDirection *= -1;
          boundaryRounds = 0;
        }

        scrollable.scrollTop = Math.max(0, scrollable.scrollTop +
          scrollDirection * Math.max(100, scrollable.clientHeight * 0.8));
        dispatchScrollEvents(scrollable);
      }
    }

    if (!complete) {
      const collectedCount = new Set(collectedMessages.filter((message) => message.threadId === threadId &&
        message.id !== threadId).map((message) => message.id)).size;
      throw new Error(`Teams stopped loading channel replies (${collectedCount}/${expected} collected for post ${threadId}). No partial channel snapshot was saved.`);
    }

    if (!thread.isConnected || !getChatRoot().contains(thread)) {
      const entry = findChatListEntry(channelLabel);

      if (!entry) {
        throw new Error("Could not return to the channel after reading replies.");
      }

      clickElement(entry);
      await wait(300);

      for (let attempt = 0; attempt < 30 &&
          !getChatRoot().querySelector("[data-tid='channel-pane-message']"); attempt += 1) {
        await wait(300);
      }

      if (!getChatRoot().querySelector("[data-tid='channel-pane-message']")) {
        throw new Error("Teams did not return to the channel after reading replies.");
      }
    }
  }
}

function assertChannelRepliesComplete(collectedMessages, replyCounts) {
  for (const [threadId, expected] of replyCounts) {
    const count = new Set(collectedMessages.filter((message) => message.threadId === threadId &&
      message.id !== threadId).map((message) => message.id)).size;

    if (count < expected) {
      throw new Error("A channel thread could not be rediscovered with all its replies. No partial channel snapshot was saved.");
    }
  }
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

function normalizeChatTitle(value) {
  const cleaned = String(value || "")
    .replace(/\s*\([^)]*\)\s*/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  const commaMatch = cleaned.match(/^([^,]+),\s*([^,+][^,+]*)$/);

  if (commaMatch) {
    return `${commaMatch[2]} ${commaMatch[1]}`.replace(/\s+/g, " ").trim();
  }

  return cleaned;
}

function getCurrentChatListLabel() {
  const selectedSelectors = [
    "[role='treeitem']:is([data-item-type='chat'], [data-item-type='muted-chat'], [data-item-type='channel'])[aria-selected='true']",
    "[role='treeitem']:is([data-item-type='chat'], [data-item-type='muted-chat'], [data-item-type='channel'])[aria-current='true']"
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
      "[role='treeitem']:is([data-item-type='chat'], [data-item-type='muted-chat'], [data-item-type='channel'])"
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
  const titleElement = element.querySelector("[id^='title-chat-list-item_'], [id^='title-channel-list-item-']");

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
    "[role='treeitem']:is([data-item-type='chat'], [data-item-type='muted-chat'], [data-item-type='channel'])",
    "[data-testid='list-item']:is([data-item-type='chat'], [data-item-type='muted-chat'], [data-item-type='channel'])"
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

function collectMessages(channelThreadId = "") {
  const chatRoot = getChatRoot();
  const channelBodies = Array.from(chatRoot.querySelectorAll(
    "[data-tid='message-body'][id^='content-'], :is([data-tid='channel-pane-message'], [data-tid='channel-replies-pane-message']) [data-tid='message-tombstone'][id^='tombstone-']"
  ));

  if (channelBodies.length) {
    return dedupeMessageRecords(channelBodies.map((body) =>
      extractChannelMessageRecord(body, channelThreadId)).filter(Boolean));
  }

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

    if (channelThreadId) {
      record.threadId = channelThreadId;
      record.id = record.id.replace(/^(?:message-body-|message-)/, "");
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

function extractChannelMessageRecord(bodyNode, channelThreadId = "") {
  const container = bodyNode.closest("[id^='message-body-']") || bodyNode;
  const metadata = container.closest("[id^='post-message-renderer-']") || container;
  const thread = bodyNode.closest("[data-tid='channel-pane-message']");
  const replyPost = bodyNode.closest("[data-tid='channel-replies-runway']")
    ?.querySelector("[id^='post-message-renderer-'] [data-tid='message-body']");
  const images = collectInlineImages(bodyNode);
  let text = normalizeText(bodyNode.innerText || bodyNode.textContent || "");
  const subject = findNestedText(metadata, ["[data-tid='subject-line']"]);

  if (subject && !text.startsWith(subject)) {
    text = `${subject}\n\n${text}`;
  }

  const links = Array.from(bodyNode.querySelectorAll("a[href]"))
    .map((link) => link.getAttribute("href") || "")
    .filter((href) => /^https?:\/\//i.test(href) && !text.includes(href));

  if (links.length) {
    text += `${text ? "\n\n" : ""}Links:\n${Array.from(new Set(links)).join("\n")}`;
  }

  if ((!text && !images.length) || isInsideIgnoredRegion(bodyNode)) {
    return null;
  }

  const id = bodyNode.id.replace(/^(?:content-|tombstone-)/, "");
  return {
    id,
    threadId: thread?.id.replace(/^reply-chain-summary-/, "") || channelThreadId ||
      replyPost?.id.replace(/^content-/, "") || id,
    ...(bodyNode.matches("[data-tid='message-tombstone']") ? { deleted: true } : {}),
    author: findNestedText(metadata, ["[id^='author-']"]) || "Unknown",
    timestamp: findTimeText(metadata),
    text,
    images
  };
}

function appendUniqueMessages(targetMessages, sourceMessages) {
  const merged = dedupeMessageRecords([...(targetMessages || []), ...(sourceMessages || [])]);
  targetMessages.length = 0;
  targetMessages.push(...merged);
}

function dedupeMessageRecords(messages) {
  const byKey = new Map();
  const deduped = [];

  for (const message of messages) {
    const key = buildMessageRecordKey(message);

    if (!key) {
      continue;
    }

    const previous = byKey.get(key);

    if (previous) {
      if (previous.threadId || message.threadId) {
        previous.threadId = message.threadId || previous.threadId;
        if (message.deleted) {
          previous.deleted = true;
        }
        previous.text = (message.text || "").length >= (previous.text || "").length
          ? message.text : previous.text;
        previous.author = message.author && message.author !== "Unknown" ? message.author : previous.author;
        previous.timestamp = message.timestamp || previous.timestamp;
        previous.images = Array.from(new Map([...(previous.images || []), ...(message.images || [])]
          .map((image) => [image.sourceUrl || image.localPath || image.embeddedDataUrl, image])).values());
      }
      continue;
    }

    const record = { ...message };
    byKey.set(key, record);
    deduped.push(record);
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
    "[data-tid='channel-replies-runway']",
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
    "[data-tid='channel-replies-runway']",
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
