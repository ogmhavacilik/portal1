/**
 * PDF Cache Utility using IndexedDB and in-memory Blob URLs.
 * Bypasses corporate firewall restrictions (which block drive.google.com/.../preview)
 * by fetching raw PDF bytes via Apps Script proxy and rendering direct local browser Blob URLs.
 */

const DB_NAME = 'ha_tech_publications_db';
const STORE_NAME = 'pdf_blobs';
const DB_VERSION = 1;

// In-memory quick lookup for active blob URLs
const blobUrlMemoryCache = new Map<string, string>();
const activeFetchPromises = new Map<string, Promise<string | null>>();

const openDb = (): Promise<IDBDatabase | null> => {
  if (typeof window === 'undefined' || !window.indexedDB) {
    return Promise.resolve(null);
  }

  return new Promise((resolve) => {
    try {
      const request = indexedDB.open(DB_NAME, DB_VERSION);

      request.onupgradeneeded = (e) => {
        const db = (e.target as IDBOpenDBRequest).result;
        if (!db.objectStoreNames.contains(STORE_NAME)) {
          db.createObjectStore(STORE_NAME, { keyPath: 'id' });
        }
      };

      request.onsuccess = () => resolve(request.result);
      request.onerror = () => {
        console.warn('[PdfCache] IndexedDB open error:', request.error);
        resolve(null);
      };
    } catch (err) {
      console.warn('[PdfCache] IndexedDB initialization failed:', err);
      resolve(null);
    }
  });
};

/**
 * Converts a Base64 string to a Blob URL.
 */
export const base64ToBlobUrl = (base64Data: string): string => {
  try {
    const clean = base64Data.replace(/^data:application\/pdf;base64,/, '').trim();
    const binary = atob(clean);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) {
      bytes[i] = binary.charCodeAt(i);
    }
    const blob = new Blob([bytes], { type: 'application/pdf' });
    return URL.createObjectURL(blob);
  } catch (err) {
    console.error('[PdfCache] base64ToBlobUrl conversion failed:', err);
    return `data:application/pdf;base64,${base64Data}`;
  }
};

/**
 * Get cached Blob URL from memory or IndexedDB.
 */
export const getCachedPdfBlobUrl = async (idOrFileId: string): Promise<string | null> => {
  if (!idOrFileId) return null;

  // 1. Check memory cache
  if (blobUrlMemoryCache.has(idOrFileId)) {
    return blobUrlMemoryCache.get(idOrFileId)!;
  }

  // 2. Check IndexedDB
  const db = await openDb();
  if (!db) return null;

  return new Promise((resolve) => {
    try {
      const tx = db.transaction(STORE_NAME, 'readonly');
      const store = tx.objectStore(STORE_NAME);
      const req = store.get(idOrFileId);

      req.onsuccess = () => {
        if (req.result && req.result.base64) {
          const blobUrl = base64ToBlobUrl(req.result.base64);
          blobUrlMemoryCache.set(idOrFileId, blobUrl);
          resolve(blobUrl);
        } else {
          resolve(null);
        }
      };

      req.onerror = () => resolve(null);
    } catch (err) {
      resolve(null);
    }
  });
};

/**
 * Save Base64 PDF to IndexedDB and memory cache.
 */
export const cachePdfBase64 = async (idOrFileId: string, base64: string): Promise<string> => {
  const blobUrl = base64ToBlobUrl(base64);
  blobUrlMemoryCache.set(idOrFileId, blobUrl);

  const db = await openDb();
  if (db) {
    try {
      const tx = db.transaction(STORE_NAME, 'readwrite');
      const store = tx.objectStore(STORE_NAME);
      store.put({ id: idOrFileId, base64, timestamp: Date.now() });
    } catch (err) {
      console.warn('[PdfCache] Failed to write PDF to IndexedDB:', err);
    }
  }

  return blobUrl;
};

/**
 * Fetch PDF via Google Apps Script proxy, cache it, and return its local Blob URL.
 */
export const fetchAndCachePdf = async (
  fileId: string,
  googleScriptUrl: string
): Promise<string | null> => {
  if (!fileId) return null;

  // If already cached
  const existing = await getCachedPdfBlobUrl(fileId);
  if (existing) return existing;

  // Check if a fetch is already in progress for this fileId
  if (activeFetchPromises.has(fileId)) {
    return activeFetchPromises.get(fileId)!;
  }

  const fetchPromise = (async () => {
    try {
      const resp = await fetch(`${googleScriptUrl}?action=getPdfBase64&fileId=${encodeURIComponent(fileId)}`);
      if (!resp.ok) {
        throw new Error(`HTTP error ${resp.status}`);
      }
      const data = await resp.json();
      if (data && data.base64) {
        const blobUrl = await cachePdfBase64(fileId, data.base64);
        return blobUrl;
      }
      return null;
    } catch (err) {
      console.error(`[PdfCache] Failed to fetch PDF for fileId ${fileId}:`, err);
      return null;
    } finally {
      activeFetchPromises.delete(fileId);
    }
  })();

  activeFetchPromises.set(fileId, fetchPromise);
  return fetchPromise;
};

/**
 * Get synchronously available Blob URL if in memory.
 */
export const getSyncBlobUrl = (idOrFileId?: string | null): string | null => {
  if (!idOrFileId) return null;
  return blobUrlMemoryCache.get(idOrFileId) || null;
};
