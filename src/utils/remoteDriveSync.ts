/**
 * Remote Google Apps Script & Google Drive Live Sync Engine (Zero-Cache, Netlify Ready)
 */

export const DIRECT_GAS_URLS = [
  "https://script.google.com/macros/s/AKfycbwP1uOo2NrST5a4I8vm1nGBLtI26yY2lWrmu9_e9iymwBkUhJBA9JOPCp7SNKqJbOubOw/exec",
  "https://script.google.com/macros/s/AKfycbw4kruTTc058Y9rLTyO3dKi6KloYsmdDTwV1GSiAk8ZXefyo3Z7_VDSTuurzsS9BHAQyQ/exec"
];

export const GOOGLE_SCRIPT_URL = DIRECT_GAS_URLS[0];
export const DRIVE_FOLDER_ID = "1HQR_NYKhHQGA7_2W3nArI9pCh-LJasTP";

export async function fetchGasGet(action: string, params: Record<string, string> = {}): Promise<any> {
  const queryParams = new URLSearchParams({
    action,
    t: String(Date.now()),
    ...params
  });

  for (const gasUrl of DIRECT_GAS_URLS) {
    const url = `${gasUrl}?${queryParams.toString()}`;
    try {
      const res = await fetch(url, {
        method: 'GET',
        mode: 'cors'
      });

      if (!res.ok) continue;

      const text = await res.text();
      if (!text || text.trim().startsWith('<!DOCTYPE') || text.trim().startsWith('<html')) {
        continue;
      }

      return JSON.parse(text);
    } catch (err: any) {
      console.warn(`[GAS GET] error on ${gasUrl.slice(-25)}:`, err.message);
    }
  }
  return null;
}

/**
 * Robust Drive Excel fetcher working seamlessly on Netlify (direct GAS fallback) and Node.js dev server.
 * Does NOT use hardcoded or stale file IDs so files are dynamically matched by name in the target folder.
 */
export async function fetchExcelFromDrive(
  fileName: string,
  candidateNames: string[] = [],
  folderId: string = DRIVE_FOLDER_ID
): Promise<{ status: string; base64?: string; error?: string; updated?: string; fileId?: string } | null> {
  const allNames = Array.from(new Set([fileName, ...candidateNames].filter(Boolean)));

  // 1. Try local server-side proxy first (if server.ts is active)
  for (const name of allNames) {
    try {
      const res = await fetch('/api/read-excel-from-drive', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ fileName: name, folderId })
      });
      if (res.ok) {
        const contentType = res.headers.get('content-type') || '';
        if (contentType.includes('text/html')) {
          continue;
        }
        const text = await res.text();
        if (text && !text.trim().startsWith('<') && !text.trim().startsWith('<!DOCTYPE')) {
          const json = JSON.parse(text);
          if (json && json.status === 'success' && json.base64) {
            return json;
          }
        }
      }
    } catch {
      // Netlify or proxy unavailable, fallback to direct GAS below
    }
  }

  // 2. Direct Fallback to Google Apps Script (Works on Netlify & All Environments)
  for (const gasUrl of DIRECT_GAS_URLS) {
    for (const name of allNames) {
      try {
        const gasRes = await fetch(gasUrl, {
          method: 'POST',
          headers: { 'Content-Type': 'text/plain;charset=utf-8' },
          body: JSON.stringify({
            action: 'readExcelFromDrive',
            fileName: name,
            folderId
          })
        });

        if (gasRes.ok) {
          const text = await gasRes.text();
          if (text && !text.trim().startsWith('<')) {
            const gasJson = JSON.parse(text);
            if (gasJson && gasJson.status === 'success' && gasJson.base64) {
              return gasJson;
            }
          }
        }
      } catch (gasErr: any) {
        console.warn(`[Direct GAS fetch] error on ${gasUrl.slice(-25)} for ${name}:`, gasErr.message);
      }
    }
  }

  return null;
}

/**
 * Uploads Excel data to Google Drive via server proxy or direct Google Apps Script fallback.
 */
export async function uploadExcelToDrive(
  fileName: string,
  base64Data: string,
  targetKey: string = 'excel_sync',
  folderId: string = DRIVE_FOLDER_ID
): Promise<boolean> {
  let backendSuccess = false;

  // 1. Try local server proxy
  try {
    const res = await fetch('/api/upload-techizat-excel', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        fileName,
        targetKey,
        base64Data,
        folderId
      })
    });
    if (res.ok) {
      backendSuccess = true;
    }
  } catch (apiErr) {
    console.warn("Backend upload failed, continuing to direct GAS upload:", apiErr);
  }

  // 2. Direct Google Apps Script upload (guaranteed for Netlify)
  for (const gasUrl of DIRECT_GAS_URLS) {
    try {
      const res = await fetch(gasUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({
          action: 'uploadTechizatExcel',
          fileName,
          folderId,
          base64Data,
          base64: base64Data,
          targetKey,
          overwrite: true,
          replaceExisting: true,
          deleteDuplicates: true,
          updateIfExists: true,
          cleanDuplicates: true
        })
      });
      if (res.ok) {
        return true;
      }
    } catch (gasErr) {
      console.warn("GAS direct upload warning:", gasErr);
    }
  }

  return backendSuccess;
}

export async function fetchSheetFromSpreadsheet(sheetName: string): Promise<any> {
  return await fetchGasGet('readSheet', { sheetName });
}
