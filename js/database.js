
const DB = (() => {
  const DB_NAME = "ShivamTutorialsDB";
  const DB_VERSION = 1;
  let database = null;

  function init() {
    return new Promise((resolve, reject) => {
      if (database) {
        if (database.objectStoreNames) {
          resolve(database);
          return;
        }
      }

      const request = indexedDB.open(DB_NAME, DB_VERSION);

      request.onupgradeneeded = event => {
        const db = event.target.result;

        if (!db.objectStoreNames.contains("students")) {
          const store = db.createObjectStore("students", { keyPath: "id" });
          store.createIndex("name", "name", { unique: false });
          store.createIndex("status", "status", { unique: false });
          store.createIndex("className", "className", { unique: false });
        }

        if (!db.objectStoreNames.contains("attendance")) {
          const store = db.createObjectStore("attendance", { keyPath: "id" });
          store.createIndex("date", "date", { unique: false });
          store.createIndex("studentId", "studentId", { unique: false });
        }

        if (!db.objectStoreNames.contains("payments")) {
          const store = db.createObjectStore("payments", { keyPath: "id" });
          store.createIndex("studentId", "studentId", { unique: false });
          store.createIndex("date", "date", { unique: false });
        }

        if (!db.objectStoreNames.contains("settings")) {
          db.createObjectStore("settings", { keyPath: "key" });
        }
      };

      request.onsuccess = () => {
        database = request.result;
        database.onversionchange = () => {
          database.close();
          database = null;
        };
        resolve(database);
      };

      request.onerror = () => reject(request.error);
      request.onblocked = () => reject(
        new Error("Database is blocked. Close other open tabs and reload.")
      );
    });
  }

  async function operation(storeName, mode, callback) {
    const db = await init();

    return new Promise((resolve, reject) => {
      let transaction;
      let request;

      try {
        transaction = db.transaction(storeName, mode);
        request = callback(transaction.objectStore(storeName));
      } catch (error) {
        reject(error);
        return;
      }

      transaction.oncomplete = () => {
        resolve(request && "result" in request ? request.result : undefined);
      };

      transaction.onerror = () => {
        reject(transaction.error || new Error("Database transaction failed."));
      };

      transaction.onabort = () => {
        reject(transaction.error || new Error("Database transaction was aborted."));
      };
    });
  }

  function add(storeName, data) {
    return operation(storeName, "readwrite", store => store.add(data));
  }

  function put(storeName, data) {
    return operation(storeName, "readwrite", store => store.put(data));
  }

  function get(storeName, id) {
    return operation(storeName, "readonly", store => store.get(id));
  }

  function getAll(storeName) {
    return operation(storeName, "readonly", store => store.getAll());
  }

  function remove(storeName, id) {
    return operation(storeName, "readwrite", store => store.delete(id));
  }

  function clear(storeName) {
    return operation(storeName, "readwrite", store => store.clear());
  }

  // Replaces all supplied stores atomically.
  // If any write fails, the transaction is rolled back.
  async function replaceAll(dataByStore) {
    const db = await init();
    const storeNames = Object.keys(dataByStore);

    if (!storeNames.length) return;

    return new Promise((resolve, reject) => {
      let transaction;

      try {
        transaction = db.transaction(storeNames, "readwrite");

        storeNames.forEach(storeName => {
          const store = transaction.objectStore(storeName);
          store.clear();

          (dataByStore[storeName] || []).forEach(record => {
            store.put(record);
          });
        });
      } catch (error) {
        reject(error);
        return;
      }

      transaction.oncomplete = () => resolve(true);
      transaction.onerror = () => reject(
        transaction.error || new Error("Backup restore failed.")
      );
      transaction.onabort = () => reject(
        transaction.error || new Error("Backup restore was cancelled.")
      );
    });
  }

  return {
    init,
    add,
    put,
    get,
    getAll,
    remove,
    clear,
    replaceAll
  };
})();
