// IndexedDB & LocalStorage & Google Drive document storage for Hangar and Teçhizat PDF Documents
// Eliminates localStorage 5MB quota errors for large PDF documents while guaranteeing multi-device / reload persistence

export interface HangarPdfDoc {
  id: string;
  itemKey: string;
  fileName: string;
  fileData: string; // Base64 data URL or Blob URL (optional if driveFileId/driveUrl present)
  docType: string;
  firma: string;
  uploadDate: string;
  fileSize: string;
  uploadedAt?: string;
  driveFileId?: string;
  driveUrl?: string;
  mimeType?: string;
}

export function getFileCategory(fileName: string): 'pdf' | 'image' | 'archive' | 'other' {
  const ext = (fileName || '').split('.').pop()?.toLowerCase();
  if (ext === 'pdf') return 'pdf';
  if (['jpg', 'jpeg', 'png', 'webp', 'gif', 'bmp', 'svg'].includes(ext || '')) return 'image';
  if (['zip', 'rar', '7z', 'tar', 'gz'].includes(ext || '')) return 'archive';
  return 'other';
}

export function getDocMimeType(fileName: string, rawData?: string): string {
  if (rawData && rawData.startsWith('data:')) {
    const match = rawData.match(/^data:([^;]+);/);
    if (match && match[1]) return match[1];
  }
  const ext = (fileName || '').split('.').pop()?.toLowerCase();
  if (ext === 'jpg' || ext === 'jpeg') return 'image/jpeg';
  if (ext === 'png') return 'image/png';
  if (ext === 'webp') return 'image/webp';
  if (ext === 'zip') return 'application/zip';
  if (ext === 'rar') return 'application/x-rar-compressed';
  if (ext === 'pdf') return 'application/pdf';
  return 'application/octet-stream';
}

const DB_NAME = 'hangar_techizat_docs_db';
const STORE_NAME = 'pdf_documents';
const DB_VERSION = 1;
const TOMBSTONE_STORAGE_KEY = 'hangar_deleted_docs_tombstone';

export function getDeletedDocTombstone(): Set<string> {
  if (typeof window === 'undefined' || !window.localStorage) return new Set();
  try {
    const raw = localStorage.getItem(TOMBSTONE_STORAGE_KEY);
    if (raw) {
      const arr: string[] = JSON.parse(raw);
      return new Set(arr.filter(Boolean));
    }
  } catch (e) {
    console.warn('Error reading tombstone:', e);
  }
  return new Set();
}

export function addDeletedDocTombstone(...keys: (string | undefined | null)[]): void {
  if (typeof window === 'undefined' || !window.localStorage) return;
  try {
    const current = getDeletedDocTombstone();
    keys.forEach(k => {
      if (k && typeof k === 'string') {
        const clean = k.trim();
        if (clean) {
          current.add(clean);
          if (clean.startsWith('drive_doc_')) {
            current.add(clean.replace('drive_doc_', ''));
          }
        }
      }
    });
    localStorage.setItem(TOMBSTONE_STORAGE_KEY, JSON.stringify(Array.from(current)));
  } catch (e) {
    console.warn('Error saving tombstone:', e);
  }
}

export const DEFAULT_GOOGLE_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbwP1uOo2NrST5a4I8vm1nGBLtI26yY2lWrmu9_e9iymwBkUhJBA9JOPCp7SNKqJbOubOw/exec";

function openDb(): Promise<IDBDatabase | null> {
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
        console.warn('[HangarPdfStorage] IndexedDB open error:', request.error);
        resolve(null);
      };
    } catch (err) {
      console.warn('[HangarPdfStorage] Initialization failed:', err);
      resolve(null);
    }
  });
}

// In-memory cache for ultra-fast access
let memoryCache: HangarPdfDoc[] | null = null;

/**
 * Safely persist document list metadata across all localStorage cache keys
 */
export function persistDocsToLocalStorage(docs: HangarPdfDoc[]): void {
  if (typeof window === 'undefined' || !window.localStorage) return;

  // Lightweight version stripping huge base64 to avoid quota limits
  const lightDocs = docs.map(d => ({
    ...d,
    fileData: d.fileData && d.fileData.length > 500 ? (d.fileData.substring(0, 100) + '...[IDB]') : d.fileData
  }));

  const serialized = JSON.stringify(lightDocs);
  const keys = ['hangar_techizat_pdf_docs', 'hangar_pdf_docs', 'hangar_pdf_docs_meta'];
  keys.forEach(k => {
    try {
      localStorage.setItem(k, serialized);
    } catch (e) {
      // Quota fallback: keep fewer items if localStorage is near full
      try {
        localStorage.setItem(k, JSON.stringify(lightDocs.slice(0, 50)));
      } catch {
        // ignore
      }
    }
  });
}

export async function getAllHangarPdfDocs(): Promise<HangarPdfDoc[]> {
  if (memoryCache && memoryCache.length > 0) {
    return memoryCache;
  }

  // 1. Try to read from IndexedDB
  const db = await openDb();
  let docs: HangarPdfDoc[] = [];

  if (db) {
    docs = await new Promise<HangarPdfDoc[]>((resolve) => {
      try {
        const tx = db.transaction(STORE_NAME, 'readonly');
        const store = tx.objectStore(STORE_NAME);
        const req = store.getAll();

        req.onsuccess = () => resolve(req.result || []);
        req.onerror = () => resolve([]);
      } catch {
        resolve([]);
      }
    });
  }

  // 2. Migration from localStorage if IndexedDB is empty or missing items
  if (typeof window !== 'undefined') {
    try {
      const keys = ['hangar_techizat_pdf_docs', 'hangar_pdf_docs', 'hangar_pdf_docs_meta'];
      for (const k of keys) {
        const saved = localStorage.getItem(k);
        if (saved) {
          const parsed = JSON.parse(saved) as HangarPdfDoc[];
          if (Array.isArray(parsed) && parsed.length > 0) {
            // Merge unique docs
            const existingIds = new Set(docs.map(d => d.id));
            parsed.forEach(p => {
              if (p && p.id && !existingIds.has(p.id)) {
                docs.push(p);
                existingIds.add(p.id);
              }
            });
            break;
          }
        }
      }
    } catch (e) {
      console.warn('LocalStorage retrieval error:', e);
    }
  }

  memoryCache = docs;
  return docs;
}

export async function saveHangarPdfDoc(doc: HangarPdfDoc): Promise<void> {
  // Update memory cache
  if (!memoryCache) {
    memoryCache = [];
  }
  memoryCache = [doc, ...memoryCache.filter(d => d.id !== doc.id)];

  // Save into IndexedDB
  const db = await openDb();
  if (db) {
    await new Promise<void>((resolve, reject) => {
      try {
        const tx = db.transaction(STORE_NAME, 'readwrite');
        const store = tx.objectStore(STORE_NAME);
        const req = store.put(doc);
        req.onsuccess = () => resolve();
        req.onerror = () => reject(req.error);
      } catch (err) {
        reject(err);
      }
    });
  }

  // Save lightweight metadata in localStorage across all keys
  persistDocsToLocalStorage(memoryCache);
}

export async function getHangarPdfDocById(id: string): Promise<HangarPdfDoc | null> {
  // Check memoryCache first if it has full fileData
  if (memoryCache) {
    const found = memoryCache.find(d => d.id === id);
    if (found && found.fileData && !found.fileData.includes('...[IDB]') && found.fileData.length > 500) {
      return found;
    }
  }

  // Fetch full doc from IndexedDB
  const db = await openDb();
  if (db) {
    const docFromDb = await new Promise<HangarPdfDoc | null>((resolve) => {
      try {
        const tx = db.transaction(STORE_NAME, 'readonly');
        const store = tx.objectStore(STORE_NAME);
        const req = store.get(id);
        req.onsuccess = () => resolve(req.result || null);
        req.onerror = () => resolve(null);
      } catch {
        resolve(null);
      }
    });

    if (docFromDb && docFromDb.fileData && !docFromDb.fileData.includes('...[IDB]')) {
      // Update memory cache item with full data
      if (memoryCache) {
        memoryCache = memoryCache.map(d => d.id === id ? docFromDb : d);
      }
      return docFromDb;
    }
  }

  // Fallback to memoryCache if present
  if (memoryCache) {
    const found = memoryCache.find(d => d.id === id);
    if (found) return found;
  }

  return null;
}

export async function deleteHangarPdfDoc(
  id: string,
  driveFileId?: string,
  fileName?: string
): Promise<void> {
  const cleanId = String(id || '').trim();
  const cleanDriveId = String(driveFileId || '').replace(/^drive_doc_/, '').trim() || (cleanId.startsWith('drive_doc_') ? cleanId.replace('drive_doc_', '') : '');
  const cleanFileName = String(fileName || '').trim();

  // 1. Add all identifiers to tombstone blacklist so sync never resurrects it
  addDeletedDocTombstone(cleanId, cleanDriveId, cleanFileName);

  // 2. Remove from in-memory cache
  if (memoryCache) {
    memoryCache = memoryCache.filter(d => {
      if (d.id === cleanId) return false;
      if (cleanDriveId && d.driveFileId === cleanDriveId) return false;
      if (cleanFileName && d.fileName === cleanFileName) return false;
      return true;
    });
    persistDocsToLocalStorage(memoryCache);
  }

  // 3. Remove from IndexedDB
  const db = await openDb();
  if (db) {
    await new Promise<void>((resolve) => {
      try {
        const tx = db.transaction(STORE_NAME, 'readwrite');
        const store = tx.objectStore(STORE_NAME);
        store.delete(cleanId);
        if (cleanDriveId && cleanDriveId !== cleanId) {
          store.delete(cleanDriveId);
          store.delete(`drive_doc_${cleanDriveId}`);
        }
        tx.oncomplete = () => resolve();
        tx.onerror = () => resolve();
      } catch {
        resolve();
      }
    });
  }

  // 4. Send request to backend to trash/delete the file from Google Drive
  const targetDriveId = cleanDriveId || (cleanId.startsWith('drive_doc_') ? cleanId.replace('drive_doc_', '') : '');
  if (targetDriveId || cleanFileName) {
    try {
      fetch('/api/delete-drive-file', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          fileId: targetDriveId,
          fileName: cleanFileName
        })
      }).catch(err => console.warn('[deleteHangarPdfDoc] Backend trash warning:', err));
    } catch (e) {
      console.warn('[deleteHangarPdfDoc] Backend fetch warning:', e);
    }

    // Direct Google Apps Script fallback for trashing file in Drive
    try {
      fetch(DEFAULT_GOOGLE_SCRIPT_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({
          action: 'deleteTechPublication',
          fileId: targetDriveId,
          id: targetDriveId,
          fileName: cleanFileName
        })
      }).catch(err => console.warn('[deleteHangarPdfDoc] GAS trash warning:', err));
    } catch (e) {
      console.warn('[deleteHangarPdfDoc] GAS fallback warning:', e);
    }
  }
}

export function normalizeDocKey(str: string): string {
  if (!str) return '';
  return String(str)
    .replace(/İ/g, 'I')
    .replace(/ı/g, 'i')
    .replace(/Ğ/g, 'G')
    .replace(/ğ/g, 'g')
    .replace(/Ü/g, 'U')
    .replace(/ü/g, 'u')
    .replace(/Ş/g, 'S')
    .replace(/ş/g, 's')
    .replace(/Ö/g, 'O')
    .replace(/ö/g, 'o')
    .replace(/Ç/g, 'C')
    .replace(/ç/g, 'c')
    .replace(/[-_.,/\\()[\]]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toUpperCase();
}

/**
 * Robust matcher that determines whether a document belongs to a table row / equipment item.
 * Matches on normalized itemKey, fileName, part number, serial number, or item title.
 */
export function isDocMatchingRow(doc: HangarPdfDoc, row: any[] | string): boolean {
  if (!doc) return false;

  if (typeof row === 'string') {
    const normSearch = normalizeDocKey(row);
    if (!normSearch || normSearch.length < 2) return false;
    const docItemKey = normalizeDocKey(doc.itemKey || '');
    const docName = normalizeDocKey(doc.fileName || '');
    return docItemKey === normSearch ||
           docItemKey.includes(normSearch) ||
           normSearch.includes(docItemKey) ||
           docName.includes(normSearch);
  }

  if (!Array.isArray(row)) return false;

  const candidateTerms: string[] = [
    row[1], // Tanım / Cihaz / Malzeme Adı
    row[2], // P/N (Parça No) or Tanım
    row[3], // S/N (Seri No) or Model
    row[4], // Model / Konum
  ].filter(val => val && String(val).trim().length >= 2).map(val => normalizeDocKey(String(val)));

  if (candidateTerms.length === 0) return false;

  const docItemKey = normalizeDocKey(doc.itemKey || '');
  const docFileName = normalizeDocKey(doc.fileName || '');

  for (const term of candidateTerms) {
    if (!term || term.length < 2) continue;

    // Direct equality or substring containment on itemKey
    if (docItemKey && (docItemKey === term || docItemKey.includes(term) || term.includes(docItemKey))) {
      return true;
    }

    // Direct containment on fileName
    if (docFileName && (docFileName.includes(term) || term.includes(docFileName))) {
      return true;
    }

    // Split multi-word terms (e.g., "200 LITRE DOKUNTU KITI" -> check "200 LITRE")
    const words = term.split(' ').filter(w => w.length >= 3);
    if (words.length >= 2) {
      const phrase2 = words.slice(0, 2).join(' ');
      if (docItemKey.includes(phrase2) || docFileName.includes(phrase2)) {
        return true;
      }
    }
  }

  return false;
}

/**
 * Filter all docs matching a specific equipment row
 */
export function findMatchingDocs(row: any[] | string | undefined, allDocs: HangarPdfDoc[]): HangarPdfDoc[] {
  if (!row || !Array.isArray(allDocs)) return [];
  return allDocs.filter(d => isDocMatchingRow(d, row));
}

/**
 * Syncs and pulls all PDF documents stored on Google Drive (both Hangar & Forms folders).
 * Ensures that whenever the page reloads or a user opens documents on any computer/browser,
 * all documents are immediately fetched from Google Drive and never disappear.
 */
export async function syncHangarPdfDocsFromDrive(googleScriptUrl?: string): Promise<HangarPdfDoc[]> {
  const scriptUrl = googleScriptUrl || DEFAULT_GOOGLE_SCRIPT_URL;
  const deletedTombstones = getDeletedDocTombstone();

  const isDocDeleted = (id?: string, driveId?: string, name?: string): boolean => {
    if (!id && !driveId && !name) return false;
    if (id && (deletedTombstones.has(id) || deletedTombstones.has(id.replace(/^drive_doc_/, '')))) return true;
    if (driveId && (deletedTombstones.has(driveId) || deletedTombstones.has(`drive_doc_${driveId}`))) return true;
    if (name && (deletedTombstones.has(name) || deletedTombstones.has(name.trim()))) return true;
    return false;
  };

  // 1. Load local docs first (filtering out any deleted tombstones)
  const rawLocalDocs = await getAllHangarPdfDocs();
  const localDocs = rawLocalDocs.filter(d => !isDocDeleted(d.id, d.driveFileId, d.fileName));
  const docsMap = new Map<string, HangarPdfDoc>();

  localDocs.forEach(d => {
    if (d.id) docsMap.set(d.id, d);
    if (d.driveFileId) docsMap.set(d.driveFileId, d);
  });

  const driveFiles: { id: string; name: string; viewUrl?: string; lastUpdated?: string; size?: string; folderId?: string }[] = [];

  // Step 1: Try high-speed internal proxy route /api/drive-techizat-pdfs
  try {
    const res = await fetch('/api/drive-techizat-pdfs?refresh=1', { signal: AbortSignal.timeout(6000) });
    if (res.ok) {
      const json = await res.json();
      if (json && Array.isArray(json.data) && json.data.length > 0) {
        json.data.forEach((item: any) => {
          if (item && item.id && item.name) {
            driveFiles.push(item);
          }
        });
      }
    }
  } catch (apiErr) {
    console.warn('[HangarPdfStorage] Internal /api/drive-techizat-pdfs failed, falling back to direct GAS:', apiErr);
  }

  // Step 2: Fallback to direct Google Apps Script if internal endpoint returned no files
  if (driveFiles.length === 0) {
    const folderIds = [
      '1HQR_NYKhHQGA7_2W3nArI9pCh-LJasTP',
      '1_fIGvuPVpC9N5on1irOfGG8OsD1KSXD0'
    ];

    for (const fId of folderIds) {
      try {
        const res = await fetch(`${scriptUrl}?action=listPdfsFromDrive&folderId=${fId}`, { signal: AbortSignal.timeout(8000) });
        if (res.ok) {
          const json = await res.json();
          if (json && Array.isArray(json.data)) {
            json.data.forEach((item: any) => {
              if (item && item.id && item.name) {
                driveFiles.push({ ...item, folderId: fId });
              }
            });
          }
        }
      } catch (err) {
        console.warn(`[HangarPdfStorage] Direct GAS sync failed for folder ${fId}:`, err);
      }
    }
  }

  // Match and merge Drive files into docs
  for (const file of driveFiles) {
    const rawName = file.name || '';
    if (!rawName.toLowerCase().endsWith('.pdf')) continue;

    // Check if file is tombstoned/deleted
    if (isDocDeleted(file.id, file.id, rawName)) {
      continue;
    }

    // Split "${itemKey}_${fileName}" or "${itemKey}___${fileName}"
    let parsedItemKey = '';
    let parsedFileName = rawName;

    const tripleSplit = rawName.indexOf('___');
    if (tripleSplit > 0) {
      parsedItemKey = rawName.substring(0, tripleSplit).trim();
      parsedFileName = rawName.substring(tripleSplit + 3).trim();
    } else {
      const olayMatch = rawName.match(/^(olay_[a-zA-Z0-9-]+_sheet_[a-zA-Z0-9-_]+_row_\d+)_(.+)$/i);
      const techMatch = rawName.match(/^(tech_[a-zA-Z0-9-_]+_row_\d+)_(.+)$/i);
      if (olayMatch) {
        parsedItemKey = olayMatch[1];
        parsedFileName = olayMatch[2];
      } else if (techMatch) {
        parsedItemKey = techMatch[1];
        parsedFileName = techMatch[2];
      } else {
        const firstUnderscore = rawName.indexOf('_');
        if (firstUnderscore > 0 && firstUnderscore < rawName.length - 1) {
          parsedItemKey = rawName.substring(0, firstUnderscore).trim();
          parsedFileName = rawName.substring(firstUnderscore + 1).trim();
        } else {
          parsedItemKey = rawName.replace(/\.pdf$/i, '').trim();
        }
      }
    }

    if (isDocDeleted(file.id, file.id, parsedFileName)) {
      continue;
    }

    // Look for matching local document
    let matchedDoc: HangarPdfDoc | undefined = docsMap.get(file.id);
    if (!matchedDoc) {
      for (const d of docsMap.values()) {
        if (d.driveFileId === file.id) {
          matchedDoc = d;
          break;
        }
        if (d.fileName && (d.fileName === parsedFileName || rawName.includes(d.fileName))) {
          if (!d.itemKey || d.itemKey.toUpperCase() === parsedItemKey.toUpperCase()) {
            matchedDoc = d;
            break;
          }
        }
      }
    }

    if (!matchedDoc) {
      // Create new document from Google Drive
      const newDoc: HangarPdfDoc = {
        id: `drive_doc_${file.id}`,
        itemKey: parsedItemKey,
        fileName: parsedFileName || rawName,
        fileData: '', // will be loaded on demand via proxy / getPdfBase64
        docType: 'Bakım / Kontrol Evrakı',
        firma: 'OGM / Google Drive',
        uploadDate: file.lastUpdated || new Date().toLocaleDateString('tr-TR'),
        fileSize: file.size || 'Google Drive PDF',
        uploadedAt: file.lastUpdated || '',
        driveFileId: file.id,
        driveUrl: file.viewUrl || `https://drive.google.com/file/d/${file.id}/preview`
      };
      docsMap.set(newDoc.id, newDoc);
      docsMap.set(file.id, newDoc);
    } else {
      // Enrich existing doc with Drive attributes
      matchedDoc.driveFileId = file.id;
      matchedDoc.driveUrl = file.viewUrl || `https://drive.google.com/file/d/${file.id}/preview`;
      if (!matchedDoc.uploadDate && file.lastUpdated) {
        matchedDoc.uploadDate = file.lastUpdated;
      }
      if (!matchedDoc.itemKey && parsedItemKey) {
        matchedDoc.itemKey = parsedItemKey;
      }
      if (!matchedDoc.fileName) {
        matchedDoc.fileName = parsedFileName || rawName;
      }
    }
  }

  // Deduplicate and filter out any tombstoned documents
  const uniqueDocs = Array.from(new Set(docsMap.values())).filter(
    d => !isDocDeleted(d.id, d.driveFileId, d.fileName)
  );
  memoryCache = uniqueDocs;

  // Asynchronously save to IndexedDB & localStorage
  try {
    const db = await openDb();
    if (db) {
      const tx = db.transaction(STORE_NAME, 'readwrite');
      const store = tx.objectStore(STORE_NAME);
      uniqueDocs.forEach(d => store.put(d));
    }
  } catch (e) {
    console.warn('[HangarPdfStorage] IndexedDB batch put warning:', e);
  }

  persistDocsToLocalStorage(uniqueDocs);
  return uniqueDocs;
}

const ROW_DOC_LINKS_KEY = 'olay_takip_row_doc_links';

/**
 * Get map of linked document IDs for table rows
 */
export function getRowDocLinks(): Record<string, string[]> {
  if (typeof window === 'undefined' || !window.localStorage) return {};
  try {
    const saved = localStorage.getItem(ROW_DOC_LINKS_KEY);
    if (saved) return JSON.parse(saved);
  } catch (e) {}
  return {};
}

/**
 * Link or unlink a document ID for a specific table row
 */
export function setRowDocLink(rowKey: string, docId: string, isLinked = true): void {
  if (typeof window === 'undefined' || !window.localStorage) return;
  try {
    const links = getRowDocLinks();
    const current = links[rowKey] || [];
    if (isLinked) {
      if (!current.includes(docId)) {
        links[rowKey] = [...current, docId];
      }
    } else {
      links[rowKey] = current.filter(id => id !== docId);
    }
    localStorage.setItem(ROW_DOC_LINKS_KEY, JSON.stringify(links));
  } catch (e) {}
}

/**
 * Match a document to an Olay Takip row using multiple intelligent heuristics:
 * 1. Saved row link ID
 * 2. Exact itemKey match
 * 3. Row index and unit matching
 * 4. Text/FileName matching against the "Belge / Döküman" cell
 */
export function matchOlayTakipDoc(
  doc: HangarPdfDoc,
  options: {
    unitKey?: string;
    sheetId?: string;
    rowIndex: number;
    cellVal?: string;
    rowCells?: string[];
  }
): boolean {
  if (!doc) return false;

  const rowKey = `olay_${options.unitKey || ''}_${options.sheetId || ''}_row_${options.rowIndex}`;
  const rowLinks = getRowDocLinks();
  const linkedIds = rowLinks[rowKey] || [];

  // 1. Direct saved link by document ID or drive file ID
  if (linkedIds.includes(doc.id) || (doc.driveFileId && linkedIds.includes(doc.driveFileId))) {
    return true;
  }

  // 2. Direct itemKey equality
  const docKey = (doc.itemKey || '').trim();
  if (options.unitKey && options.sheetId) {
    if (docKey === rowKey) return true;
  }

  // 3. Structured row marker match
  const rowMarker = `_row_${options.rowIndex}`;
  if (docKey.includes(rowMarker)) {
    if (options.sheetId && docKey.includes(options.sheetId)) return true;
    if (options.unitKey && docKey.toLowerCase().includes(options.unitKey.toLowerCase())) return true;
  }

  // 4. Name / Text Matching against cellVal (e.g. user typed "kaza_raporu.pdf" or "belge1")
  const normDocName = normalizeDocKey(doc.fileName || '').replace(/\.PDF$/i, '').trim();
  const normRawName = normalizeDocKey(doc.itemKey || '').replace(/\.PDF$/i, '').trim();

  if (options.cellVal && typeof options.cellVal === 'string' && options.cellVal.trim()) {
    const tokens = options.cellVal
      .split(/[,;\n+]+/)
      .map(t => normalizeDocKey(t).replace(/\.PDF$/i, '').trim())
      .filter(t => t.length >= 2);

    for (const token of tokens) {
      if (token === normDocName || normDocName === token) return true;
      if (token.length >= 3 && (normDocName.includes(token) || token.includes(normDocName))) return true;
      if (token.length >= 3 && (normRawName.includes(token) || token.includes(normRawName))) return true;
    }
  }

  return false;
}

