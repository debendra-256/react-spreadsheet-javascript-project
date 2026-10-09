const DATABASE_NAME = "react-spreadsheet-contact-manager";
const DATABASE_VERSION = 1;
const STORE_NAME = "contacts";

function openDatabase() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DATABASE_NAME, DATABASE_VERSION);
    request.onupgradeneeded = () => {
      if (!request.result.objectStoreNames.contains(STORE_NAME)) {
        request.result.createObjectStore(STORE_NAME, { keyPath: "id" });
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error || new Error("Could not open the contacts database."));
  });
}

export async function loadContacts() {
  const database = await openDatabase();
  try {
    return await new Promise((resolve, reject) => {
      const transaction = database.transaction(STORE_NAME, "readonly");
      const request = transaction.objectStore(STORE_NAME).getAll();
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error || new Error("Could not read contacts."));
    });
  } finally {
    database.close();
  }
}

export async function saveContacts(contacts) {
  const database = await openDatabase();
  try {
    await new Promise((resolve, reject) => {
      const transaction = database.transaction(STORE_NAME, "readwrite");
      const store = transaction.objectStore(STORE_NAME);
      store.clear();
      contacts.forEach((contact) => store.put(contact));
      transaction.oncomplete = resolve;
      transaction.onerror = () => reject(transaction.error || new Error("Could not save contacts."));
      transaction.onabort = () => reject(transaction.error || new Error("Saving contacts was cancelled."));
    });
  } finally {
    database.close();
  }
}

export async function saveContact(contact) {
  const database = await openDatabase();
  try {
    await new Promise((resolve, reject) => {
      const transaction = database.transaction(STORE_NAME, "readwrite");
      transaction.objectStore(STORE_NAME).put(contact);
      transaction.oncomplete = resolve;
      transaction.onerror = () => reject(transaction.error || new Error("Could not save contact."));
      transaction.onabort = () => reject(transaction.error || new Error("Saving the contact was cancelled."));
    });
  } finally {
    database.close();
  }
}

export async function removeContact(id) {
  const database = await openDatabase();
  try {
    await new Promise((resolve, reject) => {
      const transaction = database.transaction(STORE_NAME, "readwrite");
      transaction.objectStore(STORE_NAME).delete(id);
      transaction.oncomplete = resolve;
      transaction.onerror = () => reject(transaction.error || new Error("Could not delete contact."));
      transaction.onabort = () => reject(transaction.error || new Error("Deleting the contact was cancelled."));
    });
  } finally {
    database.close();
  }
}
