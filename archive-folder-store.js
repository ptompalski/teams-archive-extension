(function attachArchiveFolderStore() {
  const DB_NAME = "teams-archive-store";
  const STORE_NAME = "handles";
  const ARCHIVE_ROOT_KEY = "archive-root";

  function openDb() {
    return new Promise((resolve, reject) => {
      const request = indexedDB.open(DB_NAME, 1);

      request.onupgradeneeded = () => {
        const db = request.result;
        if (!db.objectStoreNames.contains(STORE_NAME)) {
          db.createObjectStore(STORE_NAME);
        }
      };

      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error || new Error("Could not open archive handle storage."));
    });
  }

  async function withStore(mode, callback) {
    const db = await openDb();

    return new Promise((resolve, reject) => {
      const transaction = db.transaction(STORE_NAME, mode);
      const store = transaction.objectStore(STORE_NAME);
      const request = callback(store);

      transaction.oncomplete = () => resolve(request?.result);
      transaction.onerror = () =>
        reject(transaction.error || new Error("Archive handle storage transaction failed."));
      transaction.onabort = () =>
        reject(transaction.error || new Error("Archive handle storage transaction aborted."));
    }).finally(() => {
      db.close();
    });
  }

  async function saveArchiveRootHandle(handle) {
    await withStore("readwrite", (store) => store.put(handle, ARCHIVE_ROOT_KEY));
  }

  async function loadArchiveRootHandle() {
    return withStore("readonly", (store) => store.get(ARCHIVE_ROOT_KEY));
  }

  async function clearArchiveRootHandle() {
    await withStore("readwrite", (store) => store.delete(ARCHIVE_ROOT_KEY));
  }

  async function queryDirectoryPermission(handle, mode) {
    if (!handle) {
      return false;
    }

    if (typeof handle.queryPermission === "function") {
      const current = await handle.queryPermission({ mode });
      return current === "granted";
    }

    return false;
  }

  async function ensureDirectoryPermission(handle, mode) {
    if (!handle) {
      return false;
    }

    if (await queryDirectoryPermission(handle, mode)) {
      return true;
    }

    if (typeof handle.requestPermission === "function") {
      const requested = await handle.requestPermission({ mode });
      return requested === "granted";
    }

    return false;
  }

  window.archiveFolderStore = {
    saveArchiveRootHandle,
    loadArchiveRootHandle,
    clearArchiveRootHandle,
    queryDirectoryPermission,
    ensureDirectoryPermission
  };
})();
