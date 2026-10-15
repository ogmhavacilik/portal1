/**
 * Remote Google Apps Script & Google Drive Live Sync Engine (Zero-Cache, Netlify Ready)
 */

export const GOOGLE_SCRIPT_URL = "https://script.google.com/macros/s/AKfycby2TM-nlURR5iPR4s4I6BpE8hbor93Jin9g014k3XPaQ0rYtS2MWHwtlnlAoph8Y3mZ/exec";
export const DRIVE_FOLDER_ID = "1HQR_NYKhHQGA7_2W3nArI9pCh-LJasTP";

export async function fetchGasGet(action: string, params: Record<string, string> = {}): Promise<any> {
  const queryParams = new URLSearchParams({
    action,
    t: String(Date.now()),
    ...params
  });

  const url = `${GOOGLE_SCRIPT_URL}?${queryParams.toString()}`;

  try {
    const res = await fetch(url, {
      method: 'GET',
      mode: 'cors'
    });

    if (!res.ok) {
      console.warn(`[GAS GET] HTTP error ${res.status} for action: ${action}`);
      return null;
    }

    const text = await res.text();
    if (!text || text.trim().startsWith('<!DOCTYPE') || text.trim().startsWith('<html')) {
      console.warn(`[GAS GET] Received HTML response instead of JSON for action: ${action}`);
      return null;
    }

    return JSON.parse(text);
  } catch (err: any) {
    console.error(`[GAS GET] Failed to fetch action ${action}:`, err.message);
    return null;
  }
}

export async function fetchExcelFromDrive(fileName: string, fileId?: string, folderId: string = DRIVE_FOLDER_ID): Promise<any> {
  // 1. If running on local server (with Node.js backend), try local proxy first
  const isLocalBackend = typeof window !== 'undefined' && (
    window.location.hostname === 'localhost' || 
    window.location.hostname === '127.0.0.1' ||
    window.location.port === '3000'
  );

  if (isLocalBackend) {
    try {
      const res = await fetch('/api/read-excel-from-drive', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ fileName, fileId, folderId })
      });
      if (res.ok) {
        const text = await res.text();
        if (text && !text.trim().startsWith('<')) {
          const json = JSON.parse(text);
          if (json && json.status === 'success' && json.base64) {
            return json;
          }
        }
      }
    } catch (e) {}
  }

  // 2. Netlify / Static / Always active Direct Google Apps Script GET
  const params: Record<string, string> = {
    fileName,
    folderId
  };
  if (fileId) params.fileId = fileId;

  return await fetchGasGet('readExcelFromDrive', params);
}

export async function fetchSheetFromSpreadsheet(sheetName: string): Promise<any> {
  return await fetchGasGet('readSheet', { sheetName });
}
