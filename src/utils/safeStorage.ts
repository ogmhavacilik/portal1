/**
 * Safe Storage Layer with Automatic Quota Management and In-Memory/Session Fallbacks
 * Prevents DOMException: QuotaExceededError and app crashes.
 */

const memoryStore = new Map<string, string>();

/**
 * Remove heavy cached items from localStorage to free up quota.
 */
export const cleanupStorageQuota = (): void => {
  if (typeof window === 'undefined') return;

  try {
    const keysToRemove: string[] = [];
    for (let i = 0; i < localStorage.length; i++) {
      const k = localStorage.key(i);
      if (!k) continue;

      // Remove heavy base64 image caches, temporary preview keys, or old pdf caches
      if (
        k.startsWith('cached_drive_img_') ||
        k.startsWith('pdf_cache_') ||
        k.startsWith('temp_') ||
        k.startsWith('preview_') ||
        k === 'ocr_temp_cache'
      ) {
        keysToRemove.push(k);
      }
    }

    keysToRemove.forEach(k => {
      try {
        localStorage.removeItem(k);
      } catch (e) {}
    });

    // Check if techizat_images in localStorage is overly huge (e.g. > 1MB of raw base64 data)
    const rawImages = localStorage.getItem('techizat_images');
    if (rawImages && rawImages.length > 800000) {
      try {
        const parsed = JSON.parse(rawImages);
        // Only keep keys that are URLs or lightweight thumbnails, or migrate them to memoryStore
        const pruned: Record<string, string> = {};
        Object.entries(parsed).forEach(([key, val]) => {
          if (typeof val === 'string') {
            memoryStore.set(`img_${key}`, val);
            // If it's a short URL or Drive ID, keep it in localStorage
            if (!val.startsWith('data:') || val.length < 50000) {
              pruned[key] = val;
            }
          }
        });
        localStorage.setItem('techizat_images', JSON.stringify(pruned));
      } catch (e) {}
    }
  } catch (e) {
    console.warn("Storage quota cleanup warning:", e);
  }
};

// Run cleanup immediately on module load
if (typeof window !== 'undefined') {
  try {
    cleanupStorageQuota();
  } catch (e) {}
}

/**
 * Safely retrieve an item from Storage (localStorage -> sessionStorage -> memoryStore).
 */
export const safeGetItem = (key: string): string | null => {
  if (typeof window === 'undefined') {
    return memoryStore.get(key) || null;
  }

  try {
    const val = localStorage.getItem(key);
    if (val !== null) return val;
  } catch (e) {}

  try {
    const val = sessionStorage.getItem(key);
    if (val !== null) return val;
  } catch (e) {}

  return memoryStore.get(key) || null;
};

/**
 * Safely save an item to Storage without ever throwing QuotaExceededError.
 */
export const safeSetItem = (key: string, value: string): boolean => {
  // Always update in-memory store for instant zero-throw retrieval
  memoryStore.set(key, value);

  if (typeof window === 'undefined') return true;

  // Try sessionStorage as secondary
  try {
    sessionStorage.setItem(key, value);
  } catch (e) {}

  // Try localStorage with automatic retry and eviction
  try {
    localStorage.setItem(key, value);
    return true;
  } catch (e: any) {
    const isQuotaError = 
      e?.name === 'QuotaExceededError' || 
      e?.name === 'NS_ERROR_DOM_QUOTA_REACHED' ||
      e?.code === 22 ||
      e?.code === 1014 ||
      e?.message?.toLowerCase().includes('quota');

    if (isQuotaError) {
      console.warn(`[SafeStorage] Quota exceeded on key "${key}". Evicting temporary caches...`);
      cleanupStorageQuota();

      try {
        localStorage.setItem(key, value);
        return true;
      } catch (retryError) {
        console.warn(`[SafeStorage] Could not persist "${key}" to localStorage after cleanup, relying on sessionStorage and memory.`);
        return false;
      }
    }
    return false;
  }
};

/**
 * Safely remove an item from all storage layers.
 */
export const safeRemoveItem = (key: string): void => {
  memoryStore.delete(key);
  if (typeof window === 'undefined') return;

  try {
    localStorage.removeItem(key);
  } catch (e) {}

  try {
    sessionStorage.removeItem(key);
  } catch (e) {}
};

/**
 * Safely parse JSON from storage with a reliable fallback.
 */
export const safeGetJSON = <T>(key: string, fallback: T): T => {
  try {
    const raw = safeGetItem(key);
    if (!raw) return fallback;
    const parsed = JSON.parse(raw);
    return parsed !== undefined && parsed !== null ? parsed : fallback;
  } catch (e) {
    console.warn(`[SafeStorage] Failed to parse JSON for key "${key}":`, e);
    return fallback;
  }
};

/**
 * Safely serialize and store JSON data.
 */
export const safeSetJSON = (key: string, data: any): boolean => {
  try {
    const serialized = JSON.stringify(data);
    return safeSetItem(key, serialized);
  } catch (e) {
    console.warn(`[SafeStorage] Failed to serialize JSON for key "${key}":`, e);
    return false;
  }
};
