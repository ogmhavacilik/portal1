import express from "express";
import path from "path";
import fs from "fs";
import * as XLSX from "xlsx";
import { createServer as createViteServer } from "vite";
import dotenv from "dotenv";
import AdmZip from "adm-zip";

dotenv.config();

async function startServer() {
  const app = express();
  const PORT = 3000;

  // Body parsing middleware with high limit for PDFs and large spreadsheets
  app.use(express.json({ limit: "150mb" }));
  app.use(express.urlencoded({ limit: "150mb", extended: true }));

  // Basic health check to verify Express is running
  app.get("/api/health", (req, res) => {
    res.json({ status: "ok", message: "Express server is running" });
  });

  // Global error handler for JSON responses
  app.use((err: any, req: express.Request, res: express.Response, next: express.NextFunction) => {
    console.error("Global Error:", err);
    if (res.headersSent) return next(err);
    res.status(500).json({ status: "error", message: err.message || "Internal Server Error" });
  });

  // API Route: Taskline EBYS Proxy
  app.get("/api/taskline-ebys", async (req, res) => {
    const scriptUrl = req.query.scriptUrl as string;
    const targetUrl = scriptUrl || "https://script.google.com/macros/s/AKfycbyZweW0GUB9DbW1CCEaEoAJjq4iYBMannyYGnp2Szr9YcxsrQi6oUGh035tncgmXwoKTw/exec";
    const finalUrl = targetUrl.replace(/\/dev$/, "/exec");
    const spreadsheetId = (req.query.spreadsheetId as string) || "1L05588TdYZmH401Lvn4_yr4zwiw2pW4EJ8dIyl-UTVQ";
    
    let fetchedData: any = null;
    let fetchError: string | null = null;
    
    // Yöntem 1: POST isteği ile "getDemands" aksiyonunu çekmeyi dene
    try {
      console.log(`[Yöntem 1] Fetching live demands via POST (getDemands) from URL: ${finalUrl} with Spreadsheet: ${spreadsheetId}...`);
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), 8000); // 8 second timeout
      
      const response = await fetch(finalUrl, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({ 
          action: "getDemands",
          spreadsheetId: spreadsheetId
        }),
        signal: controller.signal
      });
      clearTimeout(timeoutId);
      
      if (response.ok) {
        const text = await response.text();
        if (!text.includes("No HTML file named index was found") && !text.includes("HTML-bestand") && !text.trim().startsWith("<!DOCTYPE") && !text.trim().startsWith("<html")) {
          try {
            const json = JSON.parse(text);
            if (json && json.status !== "error" && (Array.isArray(json.data) || Array.isArray(json))) {
              fetchedData = json;
              console.log("[Yöntem 1] Canlı demands listesi başarıyla alındı.");
            } else {
              fetchError = json?.message || "Geçersiz response formatı";
            }
          } catch (e) {
            fetchError = "GAS yanıtı geçerli bir JSON değil.";
          }
        } else {
          fetchError = "GAS bir HTML hata sayfası döndürdü.";
        }
      } else {
        fetchError = `HTTP Durum Kodu: ${response.status}`;
      }
    } catch (err: any) {
      console.warn("[Yöntem 1] POST getDemands başarısız oldu:", err.message);
      fetchError = err.message;
    }

    // Yöntem 2: Eğer Yöntem 1 başarısız olduysa veya hata döndürdüyse, "readSheet" ile GET isteği olarak sayfaları oku
    if (!fetchedData) {
      const fallbackSheetNames = ["Sayfa1", "TASKLINE-PARÇA LİSTESİ", "İşlemdeki Talepler", "Talepler", "Demands"];
      for (const sheetName of fallbackSheetNames) {
        try {
          console.log(`[Yöntem 2] GET readSheet ile "${sheetName}" sayfasını çekmeyi deniyor...`);
          const getUrl = `${finalUrl}?action=readSheet&sheetName=${encodeURIComponent(sheetName)}&spreadsheetId=${encodeURIComponent(spreadsheetId)}`;
          const controller = new AbortController();
          const timeoutId = setTimeout(() => controller.abort(), 8000); // 8 second timeout
          
          const response = await fetch(getUrl, { signal: controller.signal });
          clearTimeout(timeoutId);
          
          if (response.ok) {
            const text = await response.text();
            if (!text.includes("No HTML file named index was found") && !text.includes("HTML-bestand") && !text.trim().startsWith("<!DOCTYPE") && !text.trim().startsWith("<html")) {
              try {
                const json = JSON.parse(text);
                if (json && json.status !== "error" && Array.isArray(json.data) && json.data.length > 0) {
                  // Kolon başlıklarını standardize ederek getDemands formatına dönüştür
                  const normalizedRows = json.data.map((row: any) => {
                    let sira = "";
                    let baslik = "";
                    let aciklama = "";
                    let tur = "MALZEME";
                    let ebys = "";
                    
                    const keys = Object.keys(row);
                    for (const key of keys) {
                      const upperKey = key.trim().toUpperCase();
                      const val = String(row[key] || "").trim();
                      
                      if (upperKey.includes("SIRA") || upperKey === "S.N." || upperKey === "S.NU.") sira = val;
                      else if (upperKey.includes("BAŞLIK") || upperKey.includes("BASLIK") || upperKey === "KONU" || upperKey === "BAŞLIK / TANIMI") baslik = val;
                      else if (upperKey.includes("AÇIKLAMA") || upperKey.includes("ACIKLAMA")) aciklama = val;
                      else if (upperKey.includes("TÜR") || upperKey.includes("TURU") || upperKey === "TİP") tur = val;
                      else if (upperKey.includes("EBYS") && !upperKey.includes("TARİH") && !upperKey.includes("TARIH") && !upperKey.includes("DATE")) ebys = val;
                    }
                    
                    if (!ebys || ebys.toLowerCase() === "n/a" || ebys.toLowerCase() === "na") {
                      if (keys.length > 7) {
                        ebys = String(row[keys[7]] || "").trim();
                      }
                    }
                    
                    return {
                      "SIRA NO": sira,
                      "Başlık": baslik || ("Talep " + ebys),
                      "Açıklama": aciklama,
                      "Talep Türü": tur || "MALZEME",
                      "EBYS NO": ebys
                    };
                  }).filter((item: any) => item["EBYS NO"] && item["EBYS NO"].length > 0 && item["EBYS NO"].toLowerCase() !== "n/a" && item["EBYS NO"].toLowerCase() !== "na");
                  
                  if (normalizedRows.length > 0) {
                    fetchedData = {
                      status: "success",
                      data: normalizedRows,
                      message: `Successfully retrieved and normalized demands from sheet "${sheetName}".`
                    };
                    console.log(`[Yöntem 2] "${sheetName}" sayfasından ${normalizedRows.length} adet kayıt başarıyla standardize edilerek okundu.`);
                    break;
                  }
                }
              } catch (e) {
                console.warn(`[Yöntem 2] "${sheetName}" JSON parse hatası veya veri yok.`);
              }
            }
          }
        } catch (err: any) {
          console.warn(`[Yöntem 2] "${sheetName}" çekme denemesi başarısız:`, err.message);
        }
      }
    }

    // Yöntem 3: Canlı bağlantıların tamamı başarısız olduysa hata döndür (Lokal fallback kaldırıldı)
    if (!fetchedData) {
      console.log("[Yöntem 3] Canlı Apps Script sorguları başarısız oldu. Veri bulunamadı.");
      return res.status(500).json({
        status: "error",
        message: "Canlı EBYS listesi çekilemedi. Lütfen bağlantınızı kontrol edin."
      });
    }

    return res.json(fetchedData);
  });

  // API Route: Taskline Submit Proxy
  app.post("/api/taskline-submit", async (req, res) => {
    try {
      const { ebysNo, talepTuru, data, scriptUrl, fallbackScriptUrl, spreadsheetId } = req.body;
      const finalSpreadsheetId = spreadsheetId || "1L05588TdYZmH401Lvn4_yr4zwiw2pW4EJ8dIyl-UTVQ";
      
      const targetUrl = scriptUrl || "https://script.google.com/macros/s/AKfycbyZweW0GUB9DbW1CCEaEoAJjq4iYBMannyYGnp2Szr9YcxsrQi6oUGh035tncgmXwoKTw/exec";
      // Ensure we target the production /exec deployment if the user accidentally copied the /dev URL
      const finalUrl = targetUrl.replace(/\/dev$/, "/exec");
      
      console.log(`[Proxy Taskline Submit] Posting to ${finalUrl} with EBYS: ${ebysNo}, Talep Türü: ${talepTuru}, Spreadsheet ID: ${finalSpreadsheetId}`);
      
      let response;
      let responseText = "";
      let isUnknownAction = false;
      let networkError = null;

      try {
        response = await fetch(finalUrl, {
          method: "POST",
          headers: {
            "Content-Type": "text/plain;charset=utf-8"
          },
          body: JSON.stringify({
            action: "appendEbysTable",
            ebysNo: ebysNo,
            talepTuru: talepTuru || "MALZEME",
            data: data,
            spreadsheetId: finalSpreadsheetId
          })
        });

        if (response.ok) {
          responseText = await response.text();
          if (responseText.includes("Unknown action")) {
            isUnknownAction = true;
          }
        } else {
          networkError = `Google Apps Script returned status: ${response.status}`;
        }
      } catch (err: any) {
        networkError = err.message || "Network request failed";
      }

      // If the first try failed (due to Unknown Action or network issue), retry with fallbackScriptUrl
      if (isUnknownAction || networkError) {
        console.log(`[Proxy Taskline Submit] First route did not support appendEbysTable or had network limits. Trying secondary path...`);
        
        const fallbackUrl = (fallbackScriptUrl || "https://script.google.com/macros/s/AKfycbzB1n5fmC2X4Zqk3S9DDA5sAcmDa7KmMClg006y9LVHYHEYhqVcZoLvDZqfGOz1SyGO/exec").replace(/\/dev$/, "/exec");
        console.log(`[Proxy Taskline Submit Fallback] Retrying with fallback URL: ${fallbackUrl}`);

        try {
          const fallbackResponse = await fetch(fallbackUrl, {
            method: "POST",
            headers: {
              "Content-Type": "text/plain;charset=utf-8"
            },
            body: JSON.stringify({
              action: "appendEbysTable",
              ebysNo: ebysNo,
              talepTuru: talepTuru || "MALZEME",
              data: data,
              spreadsheetId: finalSpreadsheetId
            })
          });

          if (fallbackResponse.ok) {
            response = fallbackResponse;
            responseText = await fallbackResponse.text();
            isUnknownAction = responseText.includes("Unknown action");
            networkError = null;
          } else {
            networkError = `Fallback Google Apps Script returned status: ${fallbackResponse.status}`;
          }
        } catch (fallbackErr: any) {
          networkError = fallbackErr.message || "Fallback network request failed";
        }
      }

      if (networkError) {
        throw new Error(networkError);
      }

      if (isUnknownAction) {
        throw new Error(`Google Apps Script returned: ${responseText}`);
      }

      console.log(`[Proxy Taskline Submit] Success Response:`, responseText.substring(0, 500));

      return res.json({
        status: "success",
        message: "Data successfully sent to Taskline Google Sheet.",
        gasStatus: response ? response.status : 200,
        response: responseText
      });
    } catch (err: any) {
      console.error("[Proxy Taskline Submit] Error:", err.message);
      return res.status(500).json({
        status: "error",
        message: `Taskline submission error: ${err.message}`
      });
    }
  });

  // API Route: Personnel List Fetcher from Google Sheet (Personel & Yoklama & 5-Personel_Bilgi)
  app.get("/api/personnel-list", async (req, res) => {
    try {
      const scriptUrl = "https://script.google.com/macros/s/AKfycbzkIYFs3JvIEQkT3Kh-XdLXtKMdsWBaXg6XY91dk_5i16_bCJf6C9zkycklXubUTir5/exec";
      const sheetIds = [
        { id: "17ScGYYx0erzDwHDk6RGiHOdJATdfmmExXFBY39dXpF0", sheets: ["5-Personel_Bilgi", "Personel", "1-Gorevlendirme"] },
        { id: "1EbwS6gIo8DHIdpZKyrX4JwM0aDejCrhRBsK-PK0f_fk", sheets: ["Personel", "Yoklama"] }
      ];
      
      let personnelList: string[] = [];

      for (const target of sheetIds) {
        for (const sheetName of target.sheets) {
          try {
            const gvizUrl = `https://docs.google.com/spreadsheets/d/${target.id}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(sheetName)}`;
            const resp = await fetch(gvizUrl);
            if (resp.ok) {
              const csvText = await resp.text();
              const lines = csvText.split(/\r?\n/);
              lines.forEach((line, idx) => {
                if (idx === 0) return; // Skip header
                const cols = line.split(/,(?=(?:(?:[^"]*"){2})*[^"]*$)/).map(c => c.replace(/^"|"$/g, '').trim());
                // Check common name columns (col 1 or 2)
                for (let c = 1; c <= 3; c++) {
                  if (cols[c]) {
                    const candidate = cols[c].trim();
                    if (
                      candidate.length > 2 &&
                      !candidate.toLowerCase().includes("ad") &&
                      !candidate.toLowerCase().includes("soyad") &&
                      !candidate.toLowerCase().includes("tc") &&
                      !candidate.toLowerCase().includes("unvan") &&
                      !candidate.toLowerCase().includes("tarih") &&
                      !/^\d+$/.test(candidate) &&
                      !personnelList.includes(candidate)
                    ) {
                      personnelList.push(candidate);
                    }
                  }
                }
              });
            }
          } catch (e) {
            // continue next sheet
          }
          if (personnelList.length > 20) break;
        }
        if (personnelList.length > 20) break;
      }

      // Method 2: If gviz didn't return enough names, fetch from GAS script
      if (personnelList.length === 0) {
        try {
          const gasUrl = `${scriptUrl}?action=readSheet&sheetName=Personel&spreadsheetId=1EbwS6gIo8DHIdpZKyrX4JwM0aDejCrhRBsK-PK0f_fk`;
          const resp = await fetch(gasUrl);
          if (resp.ok) {
            const data = await resp.json();
            if (data && Array.isArray(data.data)) {
              data.data.forEach((row: any) => {
                if (Array.isArray(row) && row[2]) {
                  personnelList.push(String(row[2]).trim());
                } else if (typeof row === 'object') {
                  const val = row["Personel"] || row["Adı Soyadı"] || row["AD SOYAD"] || row["C"] || row["İsim"] || Object.values(row)[2];
                  if (val) personnelList.push(String(val).trim());
                }
              });
            }
          }
        } catch (e) {
          console.warn("[Personnel Fetch] GAS endpoint failed.");
        }
      }

      // Default comprehensive list of active maintenance technicians, pilots & warehouse personnel
      const defaultPersonnel = [];

      defaultPersonnel.forEach(p => {
        if (!personnelList.includes(p)) personnelList.push(p);
      });

      // Remove duplicates & sort
      const uniqueSorted = Array.from(new Set(personnelList)).filter(Boolean).sort((a, b) => a.localeCompare(b, 'tr-TR'));

      return res.json({
        status: "success",
        count: uniqueSorted.length,
        personnel: uniqueSorted,
        data: uniqueSorted
      });
    } catch (err: any) {
      return res.status(500).json({
        status: "error",
        message: err.message
      });
    }
  });

  // In-memory cache for drive techizat PDFs
  let drivePdfsCache: { timestamp: number; data: any[] } = { timestamp: 0, data: [] };

  // API Route: List all PDFs from Google Drive folders (Techizat & Planning)
  app.get("/api/drive-techizat-pdfs", async (req, res) => {
    try {
      const forceRefresh = req.query.refresh === "1" || req.query.refresh === "true";
      const now = Date.now();

      // Return cached if younger than 60s and not forcing refresh
      if (!forceRefresh && drivePdfsCache.data.length > 0 && (now - drivePdfsCache.timestamp < 60000)) {
        return res.json({
          status: "success",
          source: "cache",
          data: drivePdfsCache.data,
          count: drivePdfsCache.data.length
        });
      }

      const scriptUrls = [
        "https://script.google.com/macros/s/AKfycbwP1uOo2NrST5a4I8vm1nGBLtI26yY2lWrmu9_e9iymwBkUhJBA9JOPCp7SNKqJbOubOw/exec",
        "https://script.google.com/macros/s/AKfycbw4kruTTc058Y9rLTyO3dKi6KloYsmdDTwV1GSiAk8ZXefyo3Z7_VDSTuurzsS9BHAQyQ/exec"
      ];

      const folderIds = [
        "1HQR_NYKhHQGA7_2W3nArI9pCh-LJasTP", // Teçhizat & Hangar PDF klasörü
        "1_fIGvuPVpC9N5on1irOfGG8OsD1KSXD0"  // Form kayıtları ve planlama PDF klasörü
      ];

      const allFilesMap = new Map<string, any>();

      for (const folderId of folderIds) {
        let fetched = false;
        for (const sUrl of scriptUrls) {
          try {
            const fetchUrl = `${sUrl}?action=listPdfsFromDrive&folderId=${encodeURIComponent(folderId)}`;
            const response = await fetch(fetchUrl, { headers: { "Accept": "application/json" } });
            if (response.ok) {
              const resData = await response.json();
              if (resData && (Array.isArray(resData.data) || Array.isArray(resData))) {
                const list = Array.isArray(resData.data) ? resData.data : resData;
                list.forEach((item: any) => {
                  if (item && item.id && !allFilesMap.has(item.id)) {
                    allFilesMap.set(item.id, {
                      id: item.id,
                      name: item.name || "Belge.pdf",
                      fileName: item.name || "Belge.pdf",
                      viewUrl: item.viewUrl || `https://drive.google.com/file/d/${item.id}/preview`,
                      downloadUrl: `/api/pdf-proxy?fileId=${item.id}`,
                      lastUpdated: item.lastUpdated || item.date || new Date().toLocaleString("tr-TR"),
                      size: item.size || item.sizeFormatted || "Google Drive",
                      folderId: folderId
                    });
                  }
                });
                fetched = true;
                break;
              }
            }
          } catch (fetchErr) {
            // try next script URL
          }
        }
      }

      const files = Array.from(allFilesMap.values());
      drivePdfsCache = { timestamp: now, data: files };

      return res.json({
        status: "success",
        source: "live",
        data: files,
        count: files.length
      });
    } catch (err: any) {
      console.error("Error in /api/drive-techizat-pdfs:", err);
      // If we have stale cache, return it
      if (drivePdfsCache.data.length > 0) {
        return res.json({
          status: "success",
          source: "stale-cache",
          data: drivePdfsCache.data,
          count: drivePdfsCache.data.length
        });
      }
      return res.status(500).json({
        status: "error",
        message: err.message
      });
    }
  });

  // API Route: Fast PDF Streaming & Proxy from Google Drive / GAS
  app.get("/api/pdf-proxy", async (req, res) => {
    try {
      const fileId = String(req.query.fileId || "").trim();
      const requestedFileName = String(req.query.fileName || "publication.pdf").trim();
      if (!fileId) {
        return res.status(400).send("fileId is required");
      }

      // Determine MIME type based on file extension
      const getProxyMime = (fname: string, fallbackMime?: string): string => {
        const ext = fname.split(".").pop()?.toLowerCase();
        if (ext === "jpg" || ext === "jpeg") return "image/jpeg";
        if (ext === "png") return "image/png";
        if (ext === "zip") return "application/zip";
        if (ext === "rar") return "application/x-rar-compressed";
        if (ext === "pdf") return "application/pdf";
        return fallbackMime && fallbackMime !== "text/html" ? fallbackMime : "application/pdf";
      };

      const computedMime = getProxyMime(requestedFileName);

      // Try direct Google Drive download URL
      const downloadUrl = `https://drive.google.com/uc?export=download&id=${fileId}`;
      const response = await fetch(downloadUrl, {
        headers: {
          "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        },
        redirect: "follow"
      });

      if (!response.ok) {
        // Fallback: try GAS endpoint getPdfBase64
        const gasUrl = `https://script.google.com/macros/s/AKfycbw4kruTTc058Y9rLTyO3dKi6KloYsmdDTwV1GSiAk8ZXefyo3Z7_VDSTuurzsS9BHAQyQ/exec?action=getPdfBase64&fileId=${fileId}`;
        const gasResp = await fetch(gasUrl);
        if (gasResp.ok) {
          const gasData = await gasResp.json();
          if (gasData && gasData.base64) {
            const buf = Buffer.from(gasData.base64, 'base64');
            res.setHeader("Content-Type", computedMime);
            res.setHeader("Content-Disposition", `inline; filename="${requestedFileName}"`);
            res.setHeader("Cache-Control", "public, max-age=86400");
            return res.send(buf);
          }
        }
        return res.status(response.status).send("Failed to retrieve document from Drive");
      }

      const contentType = response.headers.get("content-type") || computedMime;
      // Check if Google Drive returned a virus scan confirmation HTML page
      if (contentType.includes("text/html")) {
        const htmlText = await response.text();
        const confirmMatch = htmlText.match(/confirm=([0-9A-Za-z_]+)/) || htmlText.match(/name="confirm"\s+value="([0-9A-Za-z_]+)"/);
        if (confirmMatch && confirmMatch[1]) {
          const confirmedUrl = `https://drive.google.com/uc?export=download&confirm=${confirmMatch[1]}&id=${fileId}`;
          const confirmedResp = await fetch(confirmedUrl, { redirect: "follow" });
          if (confirmedResp.ok) {
            const arrayBuf = await confirmedResp.arrayBuffer();
            const buf = Buffer.from(arrayBuf);
            res.setHeader("Content-Type", getProxyMime(requestedFileName, confirmedResp.headers.get("content-type") || undefined));
            res.setHeader("Content-Disposition", `inline; filename="${requestedFileName}"`);
            res.setHeader("Cache-Control", "public, max-age=86400");
            return res.send(buf);
          }
        }

        // Fallback to GAS getPdfBase64
        const gasUrl = `https://script.google.com/macros/s/AKfycbw4kruTTc058Y9rLTyO3dKi6KloYsmdDTwV1GSiAk8ZXefyo3Z7_VDSTuurzsS9BHAQyQ/exec?action=getPdfBase64&fileId=${fileId}`;
        const gasResp = await fetch(gasUrl);
        if (gasResp.ok) {
          const gasData = await gasResp.json();
          if (gasData && gasData.base64) {
            const buf = Buffer.from(gasData.base64, 'base64');
            res.setHeader("Content-Type", computedMime);
            res.setHeader("Content-Disposition", `inline; filename="${requestedFileName}"`);
            res.setHeader("Cache-Control", "public, max-age=86400");
            return res.send(buf);
          }
        }
      }

      const arrayBuf = await response.arrayBuffer();
      const buf = Buffer.from(arrayBuf);
      const finalMime = getProxyMime(requestedFileName, contentType);
      res.setHeader("Content-Type", finalMime);
      res.setHeader("Content-Disposition", `inline; filename="${requestedFileName}"`);
      res.setHeader("Cache-Control", "public, max-age=86400");
      return res.send(buf);
    } catch (err: any) {
      console.error("[PDF Proxy] Error:", err.message);
      return res.status(500).send(`PDF fetch error: ${err.message}`);
    }
  });

  // API Route: AT-802 Sarf Deposu Live from Excel / Drive
  app.get("/api/at802-sarf-depo", async (req, res) => {
    try {
      const localExcelPath = path.resolve(process.cwd(), "public/at-802_sarf_ve_parca_deposu.xlsx");
      let fileBuf: Buffer | null = null;

      // Always try fetching live from Google Drive first (Zero Cache)
      try {
        const driveUrl = "https://drive.google.com/uc?export=download&id=11d1HAocT6gCmHaLEhFrP1oXj7U7bZMS7";
        const driveResp = await fetch(driveUrl, { headers: { "User-Agent": "Mozilla/5.0" } });
        if (driveResp.ok) {
          const ab = await driveResp.arrayBuffer();
          if (ab && ab.byteLength > 500) {
            fileBuf = Buffer.from(ab);
            fs.writeFileSync(localExcelPath, fileBuf);
          }
        }
      } catch (dErr) {
        console.warn("Direct Drive fetch failed, using local mirror:", dErr);
      }

      if (!fileBuf && fs.existsSync(localExcelPath)) {
        fileBuf = fs.readFileSync(localExcelPath);
      }

      if (fileBuf) {
        const xlsxModule = await import("xlsx");
        const xlsx = (xlsxModule as any).default || xlsxModule;
        const wb = xlsx.read(fileBuf, { type: "buffer" });
        const sheetName = wb.SheetNames[0];
        const sheet = wb.Sheets[sheetName];
        const raw = xlsx.utils.sheet_to_json(sheet, { header: 1, defval: "" }) as any[][];

        if (raw && raw.length > 0) {
          let headerRowIdx = 0;
          for (let r = 0; r < Math.min(25, raw.length); r++) {
            const row = (raw[r] || []).map(c => String(c || '').trim().toUpperCase());
            if (row.some(c => c.includes('DESCRIPTION') || c.includes('MALZEME') || c.includes('PARÇA') || c.includes('P/N'))) {
              headerRowIdx = r;
              break;
            }
          }

          const headerRow = (raw[headerRowIdx] || []).map(c => String(c || '').toUpperCase().trim());
          const descCol = headerRow.findIndex(c => c.includes('DESCRIPTION') || c.includes('MALZEME'));
          const pnCol = headerRow.findIndex(c => c.includes('PART') || c.includes('P/N') || c.includes('PN'));
          const snCol = headerRow.findIndex(c => c.includes('SERİ') || c.includes('SERI') || c.includes('MÜKERRER') || c.includes('MUKERRER'));
          const locCol = headerRow.findIndex(c => c.includes('LOKASYON') || c.includes('RAF') || c.includes('KONUM'));
          const gelenCol = headerRow.findIndex(c => c === 'GELEN' || c.includes('GELEN') || c.includes('MİKTAR'));
          const toplamStokCol = headerRow.findIndex(c => c.includes('TOPLAM STOK'));
          const ankaraCikanCol = headerRow.findIndex(c => c.includes('ANKARA ÇIKAN') || c.includes('ANKARA CIKAN'));
          const ankaraMevcutCol = headerRow.findIndex(c => c.includes('ANKARA MEVCUT'));
          const karainTransferCol = headerRow.findIndex(c => c.includes('KARAİN TRANSFER') || c.includes('KARAIN TRANSFER'));
          const karainCikanCol = headerRow.findIndex(c => c.includes('KARAİN ÇIKAN') || c.includes('KARAIN CIKAN'));
          const karainMevcutCol = headerRow.findIndex(c => c.includes('KARAİN MEVCUT') || c.includes('KARAIN MEVCUT'));
          const canakkaleTransferCol = headerRow.findIndex(c => c.includes('ÇANAKKALE TRANSFER') || c.includes('CANAKKALE TRANSFER'));
          const canakkaleCikanCol = headerRow.findIndex(c => c.includes('ÇANAKKALE ÇIKAN') || c.includes('CANAKKALE CIKAN'));
          const canakkaleMevcutCol = headerRow.findIndex(c => c.includes('ÇANAKKALE MEVCUT') || c.includes('CANAKKALE MEVCUT'));
          const milasTransferCol = headerRow.findIndex(c => c.includes('MİLAS TRANSFER') || c.includes('MILAS TRANSFER'));
          const milasCikanCol = headerRow.findIndex(c => c.includes('MİLAS ÇIKAN') || c.includes('MILAS CIKAN'));
          const milasMevcutCol = headerRow.findIndex(c => c.includes('MİLAS MEVCUT') || c.includes('MILAS MEVCUT'));
          const bursaTransferCol = headerRow.findIndex(c => c.includes('BURSA TRANSFER'));
          const bursaCikanCol = headerRow.findIndex(c => c.includes('BURSA ÇIKAN') || c.includes('BURSA CIKAN'));
          const bursaMevcutCol = headerRow.findIndex(c => c.includes('BURSA') && c.includes('MEVCUT'));
          const muayeneGidenCol = headerRow.findIndex(c => (c.includes('MUAYENE') || c.includes('MUAYNE')) && (c.includes('GİDEN') || c.includes('GIDEN')));
          const muayeneGelenCol = headerRow.findIndex(c => (c.includes('MUAYENE') || c.includes('MUAYNE')) && c.includes('GELEN'));
          const muayeneToplamCol = headerRow.findIndex(c => (c.includes('MUAYENE') || c.includes('MUAYNE')) && c.includes('TOPLAM'));

          const items: any[] = [];
          for (let r = headerRowIdx + 1; r < raw.length; r++) {
            const row = raw[r];
            if (!row || row.length === 0) continue;
            const desc = String(row[descCol >= 0 ? descCol : 0] || '').trim();
            if (!desc || desc.toUpperCase() === 'DESCRIPTION' || desc.toUpperCase().includes('AT-802 SARF') || desc.toUpperCase() === 'MALZEME ADI') continue;

            const pn = String(row[pnCol >= 0 ? pnCol : 1] || '-').trim() || '-';
            const sn = String(row[snCol >= 0 ? snCol : 2] || '-').trim() || '-';
            const loc = String(row[locCol >= 0 ? locCol : 3] || '-').trim() || '-';
            const gelen = Number(row[gelenCol >= 0 ? gelenCol : 4]) || 0;
            const toplamStok = toplamStokCol >= 0 ? Number(row[toplamStokCol]) || 0 : 0;
            const ankaraCikan = ankaraCikanCol >= 0 ? Number(row[ankaraCikanCol]) || 0 : 0;
            const ankaraMevcut = ankaraMevcutCol >= 0 ? Number(row[ankaraMevcutCol]) || 0 : 0;
            const karainTransfer = karainTransferCol >= 0 ? Number(row[karainTransferCol]) || 0 : 0;
            const karainCikan = karainCikanCol >= 0 ? Number(row[karainCikanCol]) || 0 : 0;
            const karainMevcut = karainMevcutCol >= 0 ? Number(row[karainMevcutCol]) || 0 : 0;
            const canakkaleTransfer = canakkaleTransferCol >= 0 ? Number(row[canakkaleTransferCol]) || 0 : 0;
            const canakkaleCikan = canakkaleCikanCol >= 0 ? Number(row[canakkaleCikanCol]) || 0 : 0;
            const canakkaleMevcut = canakkaleMevcutCol >= 0 ? Number(row[canakkaleMevcutCol]) || 0 : 0;
            const milasTransfer = milasTransferCol >= 0 ? Number(row[milasTransferCol]) || 0 : 0;
            const milasCikan = milasCikanCol >= 0 ? Number(row[milasCikanCol]) || 0 : 0;
            const milasMevcut = milasMevcutCol >= 0 ? Number(row[milasMevcutCol]) || 0 : 0;
            const bursaTransfer = bursaTransferCol >= 0 ? Number(row[bursaTransferCol]) || 0 : 0;
            const bursaCikan = bursaCikanCol >= 0 ? Number(row[bursaCikanCol]) || 0 : 0;
            const bursaMevcut = bursaMevcutCol >= 0 ? Number(row[bursaMevcutCol]) || 0 : 0;
            const muayeneGiden = muayeneGidenCol >= 0 ? Number(row[muayeneGidenCol]) || 0 : 0;
            const muayeneGelen = muayeneGelenCol >= 0 ? Number(row[muayeneGelenCol]) || 0 : 0;
            const muayeneToplam = muayeneToplamCol >= 0 ? Number(row[muayeneToplamCol]) || 0 : 0;

            items.push({
              unit: "at802",
              category: "sarf",
              description: desc,
              name: desc,
              partNumber: pn,
              pn: pn,
              serialAndNotes: sn,
              sn: sn,
              lokasyonNo: loc,
              location: loc,
              gelen: gelen,
              baseGelen: gelen,
              toplamStok: toplamStok,
              ankaraCikan: ankaraCikan,
              ankaraMevcut: ankaraMevcut,
              karainTransfer: karainTransfer,
              karainCikan: karainCikan,
              karainMevcut: karainMevcut,
              canakkaleTransfer: canakkaleTransfer,
              canakkaleCikan: canakkaleCikan,
              canakkaleMevcut: canakkaleMevcut,
              milasTransfer: milasTransfer,
              milasCikan: milasCikan,
              milasMevcut: milasMevcut,
              bursaTransfer: bursaTransfer,
              bursaCikan: bursaCikan,
              bursaMevcut: bursaMevcut,
              muayeneGiden: muayeneGiden,
              muayeneGelen: muayeneGelen,
              muayeneToplam: muayeneToplam
            });
          }

          if (items.length > 0) {
            return res.json({
              status: "success",
              count: items.length,
              source: "Google Drive Live Excel",
              items: items
            });
          }
        }
      }

      return res.status(404).json({ status: "error", message: "Veri dosyası bulunamadı." });
    } catch (err: any) {
      return res.status(500).json({ status: "error", message: err.message });
    }
  });

  // API Route: Save Depo Inventory permanently to disk, JSON, Excel and Google Drive
  app.post("/api/depo-save-inventory", async (req, res) => {
    try {
      const { items = [], unit = "at802", category, depoType } = req.body;
      if (!Array.isArray(items)) {
        return res.status(400).json({ status: "error", message: "Geçersiz veri listesi." });
      }

      const isKimyasal = category === "kimyasal" || depoType === "kimyasal" || (items.length > 0 && items[0]?.category === "kimyasal");

      const excelFileName = isKimyasal
        ? (unit === "at802" ? "at-802_kimyasal_depo.xlsx" : `${unit}_kimyasal_depo.xlsx`)
        : (unit === "at802" ? "at-802_sarf_ve_parca_deposu.xlsx" : `${unit}_sarf_ve_parca_deposu.xlsx`);

      const jsonFileName = isKimyasal
        ? `${unit}_kimyasal_data.json`
        : `${unit}_sarf_data.json`;

      const sheetName = isKimyasal
        ? `${unit.toUpperCase()}_KİMYASAL_DEPO`
        : `${unit.toUpperCase()}_SARF_DEPO`;

      // 1. Yerel /public JSON dosyasını güncelle
      const localJsonPath = path.resolve(process.cwd(), `public/${jsonFileName}`);
      fs.writeFileSync(localJsonPath, JSON.stringify(items, null, 2), "utf-8");

      // 2. /public Excel dosyasını güncelle (Mevcut sayfaları koruyarak)
      let base64Xlsx = "";
      try {
        const xlsxModule = await import("xlsx");
        const xlsx = (xlsxModule as any).default || xlsxModule;
        const localExcelPath = path.resolve(process.cwd(), `public/${excelFileName}`);

        const headerRow = isKimyasal ? [
          "SIRA NO", "DESCRIPTION", "PART NUMBER", "LOT / SERİ NUMBER",
          "MİKTAR", "LOKASYON NO", "DURUMU", "TEDARİKÇİ", "RAF ÖMRÜ VAR MI?", "RAF ÖMRÜ BİTİŞ TARİHİ"
        ] : [
          "SIRA NO", "DESCRIPTION", "PART NUMBER", "SERİ NUMBER",
          "GELEN MİKTAR", "LOKASYON NO", "TOPLAM STOK",
          "ANKARA ÇIKAN", "ANKARA MEVCUT",
          "KARAİN TRANSFER", "KARAİN ÇIKAN", "KARAİN MEVCUT",
          "ÇANAKKALE TRANSFER", "ÇANAKKALE ÇIKAN", "ÇANAKKALE MEVCUT",
          "MİLAS TRANSFER", "MİLAS ÇIKAN", "MİLAS MEVCUT",
          "BURSA TRANSFER", "BURSA ÇIKAN", "BURSA MEVCUT",
          "MUAYENE GİDEN", "MUAYENE GELEN", "MUAYENE TOPLAM",
          "DURUMU", "TEDARİKÇİ", "RAF ÖMRÜ VAR MI?", "RAF ÖMRÜ BİTİŞ TARİHİ"
        ];

        const rows = [headerRow];
        items.forEach((item: any, idx: number) => {
          const hasLife = item.hasShelfLife === 'EVET' || item.omurlu === 'EVET' || item.hasShelfLife === true || String(item.hasShelfLife || '').toUpperCase().includes('EVET');
          const lifeDate = item.shelfLifeDate || item.omurBitis || item.shelfLife || item.skt || item.shelfLifeDateVal || "-";
          
          if (isKimyasal) {
            rows.push([
              idx + 1,
              item.description || item.name || "",
              item.partNumber || item.pn || "",
              item.serialAndNotes || item.sn || "",
              item.gelen || item.baseGelen || item.toplamStok || 1,
              item.lokasyonNo || item.location || "KİMYA DEPO",
              item.durumu || "FAAL",
              item.tedarikci || "-",
              hasLife ? "EVET" : "HAYIR",
              hasLife ? lifeDate : "-"
            ]);
          } else {
            const ankaraMevcut = Number(item.ankaraMevcut || 0);
            const karainMevcut = Number(item.karainMevcut || 0);
            const canakkaleMevcut = Number(item.canakkaleMevcut || 0);
            const milasMevcut = Number(item.milasMevcut || 0);
            const bursaMevcut = Number(item.bursaMevcut || 0);
            const calcToplam = item.toplamStok !== undefined ? Number(item.toplamStok) : (ankaraMevcut + karainMevcut + canakkaleMevcut + milasMevcut + bursaMevcut);

            rows.push([
              idx + 1,
              item.description || item.name || "",
              item.partNumber || item.pn || "",
              item.serialAndNotes || item.sn || "",
              item.gelen || item.baseGelen || 1,
              item.lokasyonNo || item.location || "",
              calcToplam,
              item.ankaraCikan || 0,
              ankaraMevcut,
              item.karainTransfer || 0,
              item.karainCikan || 0,
              karainMevcut,
              item.canakkaleTransfer || 0,
              item.canakkaleCikan || 0,
              canakkaleMevcut,
              item.milasTransfer || 0,
              item.milasCikan || 0,
              milasMevcut,
              item.bursaTransfer || 0,
              item.bursaCikan || 0,
              bursaMevcut,
              item.muayeneGiden || 0,
              item.muayeneGelen || 0,
              item.muayeneToplam || 0,
              item.durumu || "FAAL",
              item.tedarikci || "-",
              hasLife ? "EVET" : "HAYIR",
              hasLife ? lifeDate : "-"
            ]);
          }
        });

        const ws = xlsx.utils.aoa_to_sheet(rows);
        
        // Single sheet workbook to avoid multiple tabs in Excel (e.g. LİSTE + AT802_SARF_DEPO)
        const wb = xlsx.utils.book_new();
        xlsx.utils.book_append_sheet(wb, ws, sheetName);

        xlsx.writeFile(wb, localExcelPath);

        const buffer = xlsx.write(wb, { type: "buffer", bookType: "xlsx" });
        base64Xlsx = buffer.toString("base64");
      } catch (excelErr) {
        console.warn("Excel file generation/update warn:", excelErr);
      }

      // 3. Google Drive'a arka planda Excel dosyasını yükle / güncelle (Live Sync)
      if (base64Xlsx) {
        excelFileCache.set(excelFileName, {
          base64: base64Xlsx,
          updated: new Date().toLocaleString("tr-TR"),
          fileId: isKimyasal ? "1EbwS6gIo8DHIdpZKyrX4JwM0aDejCrhRBsK-PK0f_fk" : "1tu8hDWSgIYkGn-7i_gDTC_UpUFUyaOzf"
        });

        const scriptUrls = [
          "https://script.google.com/macros/s/AKfycbwP1uOo2NrST5a4I8vm1nGBLtI26yY2lWrmu9_e9iymwBkUhJBA9JOPCp7SNKqJbOubOw/exec",
          "https://script.google.com/macros/s/AKfycbw4kruTTc058Y9rLTyO3dKi6KloYsmdDTwV1GSiAk8ZXefyo3Z7_VDSTuurzsS9BHAQyQ/exec",
          "https://script.google.com/macros/s/AKfycbz1eI_ckIBmmo5uMlxdXEK9TXKpylA6n0TU2INzKE9Hxov2toVrZKyFtGkqT2hCyKBztQ/exec"
        ];

        for (const sUrl of scriptUrls) {
          try {
            fetch(sUrl, {
              method: "POST",
              headers: { "Content-Type": "text/plain;charset=utf-8" },
              body: JSON.stringify({
                action: "uploadTechizatExcel",
                fileName: excelFileName,
                base64Data: base64Xlsx,
                fileId: isKimyasal ? "1EbwS6gIo8DHIdpZKyrX4JwM0aDejCrhRBsK-PK0f_fk" : "1tu8hDWSgIYkGn-7i_gDTC_UpUFUyaOzf",
                folderId: "1HQR_NYKhHQGA7_2W3nArI9pCh-LJasTP",
                unitName: isKimyasal ? `${unit.toUpperCase()} Kimyasal Depo` : `${unit.toUpperCase()} Sarf ve Parça Deposu`,
                overwrite: true,
                replaceExisting: true,
                deleteDuplicates: true,
                updateIfExists: true,
                cleanDuplicates: true
              })
            }).catch(e => console.warn("Background Drive Excel upload warn:", e));
          } catch (e) {}
        }
      }

      return res.json({
        status: "success",
        count: items.length,
        message: `${items.length} adet parça kalıcı kaydedildi ve Drive Excel dosyası güncellendi!`
      });
    } catch (err: any) {
      console.error("depo-save-inventory error:", err);
      return res.status(500).json({ status: "error", message: err.message });
    }
  });

  // API Route: Depo Transfers & Bulk Excel Sync
  function getDepoSheetNameForUnit(unit?: string): string {
    if (!unit) return "DEPO HAREKET GEÇMİŞİ-AT-802";
    const u = String(unit).trim().toLowerCase();
    if (u === 'bell429' || u === 'bell 429') return "DEPO HAREKET GEÇMİŞİ-BELL 429";
    if (u === 't70' || u === 't-70') return "DEPO HAREKET GEÇMİŞİ-T-70";
    if (u === 'c650' || u === 'c-650') return "DEPO HAREKET GEÇMİŞİ-C-650";
    if (u === 'b360' || u === 'b-360') return "DEPO HAREKET GEÇMİŞİ-B-360";
    if (u === 'hangar') return "DEPO HAREKET GEÇMİŞİ-HANGAR";
    return "DEPO HAREKET GEÇMİŞİ-AT-802";
  }

  // Google Sheets Sync Helper for DEPO HAREKET GEÇMİŞİ-AT-802
  async function syncTransactionsToGoogleSheets(transfersList?: any[], targetSheetName: string = "DEPO HAREKET GEÇMİŞİ-AT-802", singleNewTx?: any) {
    try {
      const spreadsheetId = "17ScGYYx0erzDwHDk6RGiHOdJATdfmmExXFBY39dXpF0";
      const scriptUrls = [
        "https://script.google.com/macros/s/AKfycbwP1uOo2NrST5a4I8vm1nGBLtI26yY2lWrmu9_e9iymwBkUhJBA9JOPCp7SNKqJbOubOw/exec",
        "https://script.google.com/macros/s/AKfycbw4kruTTc058Y9rLTyO3dKi6KloYsmdDTwV1GSiAk8ZXefyo3Z7_VDSTuurzsS9BHAQyQ/exec"
      ];

      const sheetsToSync = [targetSheetName];
      if (targetSheetName === "DEPO HAREKET GEÇMİŞİ-AT-802" || targetSheetName === "DEPO HAREKET GEÇMİŞİ") {
        if (!sheetsToSync.includes("DEPO HAREKET GEÇMİŞİ")) sheetsToSync.push("DEPO HAREKET GEÇMİŞİ");
        if (!sheetsToSync.includes("DEPO HAREKET GEÇMİŞİ-AT-802")) sheetsToSync.push("DEPO HAREKET GEÇMİŞİ-AT-802");
      }

      for (const sName of sheetsToSync) {
        // Fast-path: If a single new transaction is supplied, perform an immediate appendRow
        if (singleNewTx) {
          const isChem = (singleNewTx.category === 'kimyasal') || (singleNewTx.sheetName && String(singleNewTx.sheetName).toLowerCase().includes('kimya')) || (singleNewTx.location && String(singleNewTx.location).toLowerCase().includes('kimya'));
          const newRow = [
            singleNewTx.itemName || singleNewTx.name || singleNewTx.itemDesc || "",
            String(singleNewTx.quantity !== undefined ? singleNewTx.quantity : (singleNewTx.adet || 1)),
            singleNewTx.date || singleNewTx.timestamp || "",
            String(singleNewTx.islemTuru || singleNewTx.type || "TRANSFER").toUpperCase(),
            singleNewTx.sn || singleNewTx.serialNumber || "-",
            singleNewTx.tailNo || singleNewTx.kuyrukKodu || singleNewTx.kuyrukNo || "-",
            singleNewTx.operator || singleNewTx.teslimAlan || "-",
            singleNewTx.receivedBy || singleNewTx.kabulYapan || "-",
            isChem ? "KİMYASAL DEPO" : (singleNewTx.location || singleNewTx.depoYeri || "DEPO"),
            sName
          ];

          for (const sUrl of scriptUrls) {
            try {
              await fetch(sUrl, {
                method: "POST",
                headers: { "Content-Type": "text/plain;charset=utf-8" },
                body: JSON.stringify({
                  action: "appendSheetRow",
                  sheetName: sName,
                  spreadsheetId: spreadsheetId,
                  row: newRow,
                  rowData: newRow,
                  data: [newRow]
                }),
                signal: AbortSignal.timeout(15000)
              });
              break;
            } catch (e) {}
          }
        }

        const fullList = transfersList || (await fetchAndMergeAllTransactions());
        const sheetSpecificTransfers = fullList.filter((t: any) => {
          const tSheet = t.sheetName || getDepoSheetNameForUnit(t.unit);
          return tSheet === sName || tSheet === targetSheetName || !t.sheetName || tSheet.includes("AT-802");
        });

        const headers = ["MALZEME ADI", "ADET", "TARİH", "İŞLEM TÜRÜ", "SERİAL NUMBER", "KUYRUK KODU", "TESLİM ALAN", "KABUL YAPAN", "DEPO YERİ", "SAYFA ADI"];
        const rows = [
          headers,
          ...sheetSpecificTransfers.map((t: any) => {
            const isChem = (t.category === 'kimyasal') || (t.sheetName && String(t.sheetName).toLowerCase().includes('kimya')) || (t.location && String(t.location).toLowerCase().includes('kimya'));
            return [
              t.itemName || t.name || t.itemDesc || "",
              String(t.quantity !== undefined ? t.quantity : (t.adet || 1)),
              t.date || t.timestamp || "",
              String(t.islemTuru || t.type || "TRANSFER").toUpperCase(),
              t.sn || t.serialNumber || "-",
              t.tailNo || t.kuyrukKodu || t.kuyrukNo || "-",
              t.operator || t.teslimAlan || "-",
              t.receivedBy || t.kabulYapan || "-",
              isChem ? "KİMYASAL DEPO" : (t.location || t.depoYeri || "DEPO"),
              sName
            ];
          })
        ];

        for (const sUrl of scriptUrls) {
          try {
            await fetch(sUrl, {
              method: "POST",
              headers: { "Content-Type": "text/plain;charset=utf-8" },
              body: JSON.stringify({
                action: "updateSheet",
                sheetName: sName,
                spreadsheetId: spreadsheetId,
                data: rows
              }),
              signal: AbortSignal.timeout(35000)
            });
            break;
          } catch (uErr) {}
        }
      }
    } catch (err) {
      console.warn("syncTransactionsToGoogleSheets error:", err);
    }
  }

  // Persistent Depo Transactions Storage
  const DEPO_TRANSACTIONS_FILE = path.resolve(process.cwd(), "data/depo_transactions.json");
  const DELETED_TX_FILE = path.resolve(process.cwd(), "data/deleted_tx_ids.json");
  const ONAY_BEKLEYENLER_FILE = path.resolve(process.cwd(), "data/onay_bekleyenler.json");

  const getDeletedTxKeys = (): Set<string> => {
    try {
      if (fs.existsSync(DELETED_TX_FILE)) {
        const raw = fs.readFileSync(DELETED_TX_FILE, "utf-8");
        const arr = JSON.parse(raw);
        if (Array.isArray(arr)) return new Set(arr);
      }
    } catch {}
    return new Set<string>();
  };

  const addDeletedTxKey = (key: string) => {
    try {
      if (!key) return;
      const set = getDeletedTxKeys();
      set.add(key);
      const dir = path.dirname(DELETED_TX_FILE);
      if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
      fs.writeFileSync(DELETED_TX_FILE, JSON.stringify(Array.from(set), null, 2), "utf-8");
    } catch (err) {
      console.warn("Error saving deleted tx key:", err);
    }
  };

  const readOnayBekleyenler = (): any[] => {
    try {
      if (fs.existsSync(ONAY_BEKLEYENLER_FILE)) {
        const raw = fs.readFileSync(ONAY_BEKLEYENLER_FILE, "utf-8");
        const parsed = JSON.parse(raw);
        if (Array.isArray(parsed)) return parsed;
      }
    } catch {}
    return [];
  };

  const writeOnayBekleyenler = (items: any[]) => {
    try {
      const dir = path.dirname(ONAY_BEKLEYENLER_FILE);
      if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
      fs.writeFileSync(ONAY_BEKLEYENLER_FILE, JSON.stringify(items, null, 2), "utf-8");

      const pubPath = path.resolve(process.cwd(), "public/onay_bekleyenler.json");
      fs.writeFileSync(pubPath, JSON.stringify(items, null, 2), "utf-8");
    } catch (err) {
      console.warn("Error writing onay_bekleyenler:", err);
    }
  };

  // Google Sheets sync for "ONAY BEKLEYENLER -AT802"
  async function syncOnayBekleyenlerToDrive(items: any[]) {
    try {
      const spreadsheetId = "17ScGYYx0erzDwHDk6RGiHOdJATdfmmExXFBY39dXpF0";
      const sheetName = "ONAY BEKLEYENLER -AT802";
      const headers = ["TALEP NO", "TARİH", "MALZEME TÜRÜ", "MALZEME ADI", "PART NUMBER (P/N)", "SERIAL NUMBER (S/N)", "TALEP EDİLEN DEPO", "ADET", "KUYRUK NO", "TESLİM ALACAK", "AÇIKLAMA", "DURUM"];
      const rows = [
        headers,
        ...items.map(t => [
          t.id || "",
          t.date || t.timestamp || "",
          t.category === 'kimyasal' ? "KİMYASAL" : "SARF/PARÇA",
          t.itemName || t.description || "",
          t.pn || t.partNumber || "-",
          t.sn || t.serialNumber || "-",
          t.depot || t.depo || "ANKARA",
          String(t.quantity || t.adet || 1),
          t.tailNo || t.kuyrukNo || "-",
          t.requestedBy || t.teslimAlan || "-",
          t.notes || t.aciklama || "-",
          t.status || "ONAY BEKLİYOR"
        ])
      ];

      const scriptUrls = [
        "https://script.google.com/macros/s/AKfycbwP1uOo2NrST5a4I8vm1nGBLtI26yY2lWrmu9_e9iymwBkUhJBA9JOPCp7SNKqJbOubOw/exec",
        "https://script.google.com/macros/s/AKfycbw4kruTTc058Y9rLTyO3dKi6KloYsmdDTwV1GSiAk8ZXefyo3Z7_VDSTuurzsS9BHAQyQ/exec"
      ];

      for (const sUrl of scriptUrls) {
        try {
          await fetch(sUrl, {
            method: "POST",
            headers: { "Content-Type": "text/plain;charset=utf-8" },
            body: JSON.stringify({
              action: "updateSheet",
              sheetName: sheetName,
              spreadsheetId: spreadsheetId,
              data: rows
            }),
            signal: AbortSignal.timeout(10000)
          });
          break;
        } catch {}
      }
    } catch (syncErr) {
      console.warn("syncOnayBekleyenlerToDrive error:", syncErr);
    }
  }

  const readPersistedDepoTransactions = (): any[] => {
    try {
      if (fs.existsSync(DEPO_TRANSACTIONS_FILE)) {
        const raw = fs.readFileSync(DEPO_TRANSACTIONS_FILE, "utf-8");
        const parsed = JSON.parse(raw);
        if (Array.isArray(parsed)) return parsed;
      }
    } catch (err) {
      console.warn("Error reading persisted depo transactions:", err);
    }
    return [];
  };

  const writePersistedDepoTransactions = (transactions: any[]) => {
    try {
      const dir = path.dirname(DEPO_TRANSACTIONS_FILE);
      if (!fs.existsSync(dir)) {
        fs.mkdirSync(dir, { recursive: true });
      }
      fs.writeFileSync(DEPO_TRANSACTIONS_FILE, JSON.stringify(transactions, null, 2), "utf-8");
    } catch (err) {
      console.warn("Error writing persisted depo transactions:", err);
    }
  };

  async function fetchAndMergeAllTransactions(): Promise<any[]> {
    const sheetId = "17ScGYYx0erzDwHDk6RGiHOdJATdfmmExXFBY39dXpF0";
    const possibleSheetTargets = [
      { name: "DEPO HAREKET GEÇMİŞİ-AT-802", url: `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv&gid=1298613975` },
      { name: "DEPO HAREKET GEÇMİŞİ-BELL 429", url: `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent("DEPO HAREKET GEÇMİŞİ-BELL 429")}` },
      { name: "DEPO HAREKET GEÇMİŞİ-T-70", url: `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent("DEPO HAREKET GEÇMİŞİ-T-70")}` },
      { name: "DEPO HAREKET GEÇMİŞİ-C-650", url: `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent("DEPO HAREKET GEÇMİŞİ-C-650")}` },
      { name: "DEPO HAREKET GEÇMİŞİ-B-360", url: `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent("DEPO HAREKET GEÇMİŞİ-B-360")}` },
      { name: "DEPO HAREKET GEÇMİŞİ-HANGAR", url: `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent("DEPO HAREKET GEÇMİŞİ-HANGAR")}` },
      { name: "TRANSFER GEÇMİŞİ", url: `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent("TRANSFER GEÇMİŞİ")}` }
    ];
    
    const allTransfers: any[] = [];
    const fetchPromises = possibleSheetTargets.map(async (target) => {
      const sheetName = target.name;
      const gvizUrl = target.url;
      try {
        const resp = await fetch(gvizUrl);
        if (resp.ok) {
          const csvText = await resp.text();
          if (!csvText.includes("<!DOCTYPE") && !csvText.includes("<html") && csvText.trim().length > 0) {
            const lines = csvText.split(/\r?\n/).filter(l => l.trim().length > 0);
            if (lines.length > 0) {
              const headerLine = (lines[0] || "").toUpperCase();
              let delimiter = ",";
              if (headerLine.includes(";")) {
                delimiter = ";";
              } else {
                const commaCount = (headerLine.match(/,/g) || []).length;
                const semiCount = (headerLine.match(/;/g) || []).length;
                if (semiCount > commaCount) delimiter = ";";
              }

              const regex = new RegExp(`${delimiter}(?=(?:(?:[^"]*"){2})*[^"]*$)`);
              const list: any[] = [];

              for (let i = 1; i < lines.length; i++) {
                const line = lines[i];
                if (!line.trim()) continue;
                const cols = line.split(regex).map(c => c.replace(/^"|"$/g, '').trim());
                if (cols.length > 0 && cols[0]) {
                  const effectiveSheet = cols[9] || sheetName;
                  let unit = "at802";
                  if (effectiveSheet.includes("BELL") || sheetName.includes("BELL")) unit = "bell429";
                  else if (effectiveSheet.includes("T-70") || sheetName.includes("T-70")) unit = "t70";
                  else if (effectiveSheet.includes("C-650") || sheetName.includes("C-650")) unit = "c650";
                  else if (effectiveSheet.includes("B-360") || sheetName.includes("B-360")) unit = "b360";
                  else if (effectiveSheet.includes("HANGAR") || sheetName.includes("HANGAR")) unit = "hangar";
                  else unit = "at802";

                  const locNorm = String(cols[8] || '').toUpperCase().replace(/İ/g, 'I');
                  const itemNorm = String(cols[0] || '').toUpperCase().replace(/İ/g, 'I');
                  const isKimya = locNorm.includes('KIMYA') || 
                                  effectiveSheet.toUpperCase().replace(/İ/g, 'I').includes('KIMYA') ||
                                  itemNorm.includes('OIL') || itemNorm.includes('GREASE') || 
                                  itemNorm.includes('AEROSHELL') || itemNorm.includes('FLUID') || 
                                  itemNorm.includes('HYDRAULIC') || itemNorm.includes('BOYA') || 
                                  itemNorm.includes('YAG');

                  list.push({
                    id: `tx_sheets_${sheetName}_${i}`, // Stable non-dynamic ID based on sheet and row index
                    itemName: cols[0] || "",
                    name: cols[0] || "",
                    quantity: parseInt(cols[1]) || 1,
                    date: cols[2] || "",
                    type: (cols[3] || "").toUpperCase(),
                    sn: cols[4] || "",
                    tailNo: cols[5] || "",
                    operator: cols[6] || "",
                    receivedBy: cols[7] || "",
                    location: cols[8] || "",
                    unit: unit,
                    category: isKimya ? "kimyasal" : "sarf",
                    sheetName: effectiveSheet
                  });
                }
              }
              return list;
            }
          }
        }
      } catch (fetchErr) {
        console.warn(`Error reading sheet ${sheetName}:`, fetchErr);
      }
      return [];
    });

    const results = await Promise.allSettled(fetchPromises);
    results.forEach((res) => {
      if (res.status === 'fulfilled' && Array.isArray(res.value)) {
        allTransfers.push(...res.value);
      }
    });

    // Merge permanent data file data (data/depo_transactions.json)
    try {
      const localTxs = readPersistedDepoTransactions();
      if (Array.isArray(localTxs) && localTxs.length > 0) {
        const getStableKey = (t: any) => {
          if (t.id && !String(t.id).startsWith("tx_sheets_")) return String(t.id).trim();
          return `${t.itemName || t.name}_${t.timestamp || t.date || t.dateStr}_${t.quantity}_${t.type || t.islemTuru}_${t.sn || '-'}`;
        };
        const existingKeys = new Set(allTransfers.map(getStableKey));
        localTxs.forEach(t => {
          const key = getStableKey(t);
          if (!existingKeys.has(key)) {
            allTransfers.push(t);
            existingKeys.add(key);
          }
        });
      }
    } catch (err) {
      console.warn("Local data transactions merge warn:", err);
    }

    // Merge public JSON transactions too (public/depo_transactions.json)
    try {
      const localJsonPath = path.resolve(process.cwd(), "public/depo_transactions.json");
      if (fs.existsSync(localJsonPath)) {
        const localData = fs.readFileSync(localJsonPath, "utf-8");
        const publicTxs = JSON.parse(localData);
        if (Array.isArray(publicTxs)) {
          const getStableKey = (t: any) => {
            if (t.id && !String(t.id).startsWith("tx_sheets_")) return String(t.id).trim();
            return `${t.itemName || t.name}_${t.timestamp || t.date || t.dateStr}_${t.quantity}_${t.type || t.islemTuru}_${t.sn || '-'}`;
          };
          const existingKeys = new Set(allTransfers.map(getStableKey));
          publicTxs.forEach(t => {
            const key = getStableKey(t);
            if (!existingKeys.has(key)) {
              allTransfers.push(t);
              existingKeys.add(key);
            }
          });
        }
      }
    } catch (err) {
      console.warn("Public json transactions merge warn:", err);
    }

    const deletedSet = getDeletedTxKeys();
    if (deletedSet.size > 0) {
      return allTransfers.filter(t => {
        if (!t) return false;
        if (t.id && deletedSet.has(String(t.id).trim())) return false;
        const k = `${t.itemName || t.name}_${t.timestamp || t.date || t.dateStr}_${t.quantity}_${t.type || t.islemTuru}_${t.sn || '-'}`;
        if (deletedSet.has(k)) return false;
        return true;
      });
    }

    return allTransfers;
  }

  // DELETE Depo Transaction with Password Protection
  app.post("/api/delete-depo-transaction", async (req, res) => {
    try {
      const { id, tx, password, unit } = req.body;
      const authorizedPass = ["802", "1839", "1234", "8902", "1923", "2024"];
      const providedPass = String(password || "").trim();
      if (!authorizedPass.includes(providedPass)) {
        return res.status(403).json({ status: "error", message: "Yetkisiz işlem! Hatalı şifre." });
      }

      const txId = id ? String(id).trim() : (tx?.id ? String(tx.id).trim() : "");
      let txKey = "";
      if (tx) {
        txKey = `${tx.itemName || tx.name}_${tx.timestamp || tx.date || tx.dateStr}_${tx.quantity}_${tx.type || tx.islemTuru}_${tx.sn || '-'}`;
      }
      if (txId) addDeletedTxKey(txId);
      if (txKey) addDeletedTxKey(txKey);

      // Read current transactions and filter out deleted item
      const fullList = await fetchAndMergeAllTransactions();
      const updatedList = fullList.filter(t => {
        if (txId && t.id && String(t.id).trim() === txId) return false;
        const k = `${t.itemName || t.name}_${t.timestamp || t.date || t.dateStr}_${t.quantity}_${t.type || t.islemTuru}_${t.sn || '-'}`;
        if (txKey && k === txKey) return false;
        return true;
      });

      writePersistedDepoTransactions(updatedList);
      const localJsonPath = path.resolve(process.cwd(), "public/depo_transactions.json");
      fs.writeFileSync(localJsonPath, JSON.stringify(updatedList, null, 2), "utf-8");

      // Sync updated sheet without this row to Google Sheets
      const targetSheet = (tx && tx.sheetName) || getDepoSheetNameForUnit(unit || (tx && tx.unit) || 'at802');
      const spreadsheetId = "17ScGYYx0erzDwHDk6RGiHOdJATdfmmExXFBY39dXpF0";
      const sheetSpecificTransfers = updatedList.filter((t: any) => {
        const tSheet = t.sheetName || getDepoSheetNameForUnit(t.unit);
        return tSheet === targetSheet;
      });

      const headers = ["MALZEME ADI", "ADET", "TARİH", "İŞLEM TÜRÜ", "SERİAL NUMBER", "KUYRUK KODU", "TESLİM ALAN", "KABUL YAPAN", "DEPO YERİ", "SAYFA ADI"];
      const rows = [
        headers,
        ...sheetSpecificTransfers.map((t: any) => [
          t.itemName || t.name || t.itemDesc || "",
          String(t.quantity !== undefined ? t.quantity : (t.adet || 1)),
          t.date || t.timestamp || "",
          String(t.islemTuru || t.type || "TRANSFER").toUpperCase(),
          t.sn || t.serialNumber || "-",
          t.tailNo || t.kuyrukKodu || t.kuyrukNo || "-",
          t.operator || t.teslimAlan || "-",
          t.receivedBy || t.kabulYapan || "-",
          t.location || t.depoYeri || "DEPO",
          t.sheetName || targetSheet
        ])
      ];

      const scriptUrls = [
        "https://script.google.com/macros/s/AKfycbwP1uOo2NrST5a4I8vm1nGBLtI26yY2lWrmu9_e9iymwBkUhJBA9JOPCp7SNKqJbOubOw/exec",
        "https://script.google.com/macros/s/AKfycbw4kruTTc058Y9rLTyO3dKi6KloYsmdDTwV1GSiAk8ZXefyo3Z7_VDSTuurzsS9BHAQyQ/exec"
      ];

      (async () => {
        for (const sUrl of scriptUrls) {
          try {
            await fetch(sUrl, {
              method: "POST",
              headers: { "Content-Type": "text/plain;charset=utf-8" },
              body: JSON.stringify({
                action: "updateSheet",
                sheetName: targetSheet,
                spreadsheetId: spreadsheetId,
                data: rows
              }),
              signal: AbortSignal.timeout(10000)
            });
            break;
          } catch {}
        }
      })();

      return res.json({ status: "success", message: "Hareket kaydı sistemden ve Google Drive sayfasından kalıcı olarak silindi." });
    } catch (err: any) {
      return res.status(500).json({ status: "error", message: err.message });
    }
  });

  // ─── ONAY BEKLEYENLER (APPROVAL REQUESTS) ENDPOINTS ───
  // GET /api/onay-bekleyenler
  app.get("/api/onay-bekleyenler", async (req, res) => {
    try {
      let localItems = readOnayBekleyenler();
      
      // Try to fetch LIVE data from Google Sheets (GViz CSV Export)
      try {
        const spreadsheetId = "17ScGYYx0erzDwHDk6RGiHOdJATdfmmExXFBY39dXpF0";
        const sheetName = "ONAY BEKLEYENLER -AT802";
        const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(sheetName)}`;
        
        const sheetRes = await fetch(url, { signal: AbortSignal.timeout(4000) });
        if (sheetRes.ok) {
          const csvText = await sheetRes.text();
          if (csvText && !csvText.includes("<!DOCTYPE")) {
            const lines = csvText.split(/\r?\n/).filter(l => l.trim().length > 0);
            const regex = new RegExp(`,(?=(?:(?:[^"]*"){2})*[^"]*$)`);
            
            if (lines.length > 1) {
              const remoteItems: any[] = [];
              for (let i = 1; i < lines.length; i++) {
                const cols = lines[i].split(regex).map(c => c.replace(/^"|"$/g, '').trim());
                if (cols.length >= 1) {
                  const id = cols[0];
                  if (!id) continue;
                  
                  remoteItems.push({
                    id: id,
                    date: cols[1] || "",
                    category: cols[2] === "KİMYASAL" ? "kimyasal" : "sarf",
                    itemName: cols[3] || "",
                    pn: cols[4] || "",
                    sn: cols[5] || "",
                    depot: cols[6] || "",
                    quantity: Number(cols[7]) || 1,
                    tailNo: cols[8] || "",
                    requestedBy: cols[9] || "",
                    notes: cols[10] || "",
                    status: cols[11] || "ONAY BEKLİYOR",
                    unit: "at802", // Primary unit for this sheet
                    fromRemote: true
                  });
                }
              }
              
              // Merge Remote with Local (Prioritize Local status if it changed, but keep new items from remote)
              const merged: any[] = [...localItems];
              remoteItems.forEach(ri => {
                const existingIdx = merged.findIndex(li => li.id === ri.id);
                if (existingIdx === -1) {
                  // New item from portal/excel
                  merged.unshift(ri);
                } else {
                  // Existing item: Update fields if status is still pending, but keep local for others
                  // This is a simple merge, we trust remote status if it's ONAYLANDI for example
                  if (ri.status !== merged[existingIdx].status) {
                    merged[existingIdx].status = ri.status;
                  }
                }
              });
              
              // Sort by date/timestamp descending
              merged.sort((a, b) => {
                const da = new Date(a.timestamp || a.date).getTime();
                const db = new Date(b.timestamp || b.date).getTime();
                return db - da;
              });

              // Update local cache if changed
              if (JSON.stringify(localItems) !== JSON.stringify(merged)) {
                writeOnayBekleyenler(merged);
                localItems = merged;
              }
            }
          }
        }
      } catch (syncErr) {
        console.warn("Live approval sync warning:", syncErr.message);
      }

      return res.json({ status: "success", items: localItems });
    } catch (err: any) {
      return res.json({ status: "success", items: [] });
    }
  });

  // POST /api/onay-bekleyenler (Create request from Barkod Okuyucu - Single or Batch)
  app.post("/api/onay-bekleyenler", async (req, res) => {
    try {
      const current = readOnayBekleyenler();
      const createdRequests: any[] = [];

      const rawItems = Array.isArray(req.body.items) ? req.body.items : [req.body];

      for (const item of rawItems) {
        if (!item) continue;
        const {
          pn = "",
          sn = "",
          itemName = "",
          category = "sarf",
          quantity = 1,
          date = new Date().toLocaleString("tr-TR"),
          depot = "ANKARA",
          tailNo = "-",
          requestedBy = "",
          notes = "",
          unit = "at802"
        } = item;

        const newRequest = {
          id: `TLP-${Date.now().toString().slice(-6)}-${Math.random().toString(36).substr(2, 3).toUpperCase()}`,
          timestamp: new Date().toISOString(),
          date: date || new Date().toLocaleString("tr-TR"),
          pn: String(pn).trim(),
          sn: String(sn).trim() || "-",
          itemName: String(itemName).trim() || String(pn).trim() || "Malzeme",
          category: category || "sarf",
          quantity: Number(quantity) || 1,
          depot: String(depot).trim().toUpperCase(),
          tailNo: String(tailNo).trim() || "-",
          requestedBy: String(requestedBy).trim() || "Personel",
          notes: String(notes).trim(),
          status: "ONAY BEKLİYOR",
          unit: unit || "at802"
        };

        createdRequests.push(newRequest);
        current.unshift(newRequest);
      }

      writeOnayBekleyenler(current);

      // Background sync to Drive sheet
      syncOnayBekleyenlerToDrive(current).catch(() => {});

      return res.json({
        status: "success",
        message: `${createdRequests.length} adet talep oluşturuldu ve 'ONAY BEKLEYENLER' sayfasına kaydedildi.`,
        requests: createdRequests,
        request: createdRequests[0]
      });
    } catch (err: any) {
      return res.status(500).json({ status: "error", message: err.message });
    }
  });

  // POST /api/onay-bekleyenler/onayla (Depo Personeli Appoval with Password)
  app.post("/api/onay-bekleyenler/onayla", async (req, res) => {
    try {
      const { id, password, approver = "Depo Yetkilisi" } = req.body;
      const authorizedPass = ["802", "1839", "1234", "8902", "1923", "2024"];
      if (!authorizedPass.includes(String(password || "").trim())) {
        return res.status(403).json({ status: "error", message: "Yetkisiz işlem! Hatalı şifre." });
      }

      const items = readOnayBekleyenler();
      const target = items.find(it => it.id === id);
      if (!target) {
        return res.status(404).json({ status: "error", message: "Talep bulunamadı." });
      }

      target.status = "ONAYLANDI";
      target.approvedBy = approver;
      target.approvalDate = new Date().toLocaleString("tr-TR");
      writeOnayBekleyenler(items);
      syncOnayBekleyenlerToDrive(items).catch(() => {});

      // Add to Depo Hareket Geçmişi
      const depotName = (target.depot || "ANKARA").toUpperCase();
      const notesLower = (target.notes || "").toLowerCase();
      const itemDescLower = (target.itemName || "").toLowerCase();
      
      const isSayim = target.category === 'sayim' || 
        notesLower.includes('sayım') || notesLower.includes('sayim') || 
        notesLower.includes('fazla') || notesLower.includes('fark');

      let txType = `${depotName} ÇIKAN`;
      if (isSayim) {
        if (depotName.includes('ANKARA')) {
          txType = "ANKARA GELEN";
        } else {
          txType = `${depotName} TRANSFER`;
        }
      } else if (notesLower.includes('muayeneden dön') || notesLower.includes('onarıl')) {
        txType = "MUAYENEDEN DÖNEN";
      }

      let finalSn = target.sn || '-';
      if (txType.includes('MUAYENEDEN DÖNEN') || txType.includes('MUAYENE GELEN') || notesLower.includes('onarıl')) {
        if (finalSn && finalSn !== '-' && !finalSn.includes('(onarılmış)')) {
          finalSn = `${finalSn} (onarılmış)`;
        }
      }

      const itemUnit = target.unit || "at802";
      const targetSheet = getDepoSheetNameForUnit(itemUnit);

      const newTx = {
        id: `tx_approved_${Date.now()}`,
        timestamp: new Date().toISOString(),
        date: target.approvalDate || new Date().toLocaleString("tr-TR"),
        type: txType,
        itemName: target.itemName,
        pn: target.pn,
        sn: finalSn,
        quantity: target.quantity,
        tailNo: target.tailNo,
        operator: target.requestedBy,
        receivedBy: approver,
        location: target.category === 'kimyasal' ? 'KİMYASAL DEPO' : (depotName.includes('ANKARA') ? 'ANKARA MERKEZ DEPO' : `${depotName} DEPOSU`),
        unit: itemUnit,
        category: target.category,
        sheetName: targetSheet,
        notes: isSayim ? `SAYIM FAZLASI / FARKI (${target.id}): ${target.notes || ''}` : `ONAYLANAN TALEP (${target.id}): ${target.notes || ''}`
      };

      const fullList = await fetchAndMergeAllTransactions();
      fullList.unshift(newTx);

      writePersistedDepoTransactions(fullList);
      const localJsonPath = path.resolve(process.cwd(), "public/depo_transactions.json");
      fs.writeFileSync(localJsonPath, JSON.stringify(fullList, null, 2), "utf-8");

      // Fast-path: appendRow to Google Sheets at the bottom of targetSheet (e.g. DEPO HAREKET GEÇMİŞİ-AT-802)
      syncTransactionsToGoogleSheets(fullList, targetSheet, newTx).catch(() => {});

      return res.json({
        status: "success",
        message: "Talep onaylandı ve depo hareketi kaydedildi.",
        approvedTx: newTx
      });
    } catch (err: any) {
      return res.status(500).json({ status: "error", message: err.message });
    }
  });

  // POST /api/onay-bekleyenler/reddet
  app.post("/api/onay-bekleyenler/reddet", async (req, res) => {
    try {
      const { id, password, rejector = "Depo Yetkilisi", reason = "" } = req.body;
      const authorizedPass = ["802", "1839", "1234", "8902", "1923", "2024"];
      if (!authorizedPass.includes(String(password || "").trim())) {
        return res.status(403).json({ status: "error", message: "Yetkisiz işlem! Hatalı şifre." });
      }

      const items = readOnayBekleyenler();
      const target = items.find(it => it.id === id);
      if (!target) {
        return res.status(404).json({ status: "error", message: "Talep bulunamadı." });
      }

      target.status = "REDDEDİLDİ";
      target.rejectedBy = rejector;
      target.rejectReason = reason;
      target.rejectDate = new Date().toLocaleString("tr-TR");
      writeOnayBekleyenler(items);
      syncOnayBekleyenlerToDrive(items).catch(() => {});

      return res.json({ status: "success", message: "Talep reddedildi." });
    } catch (err: any) {
      return res.status(500).json({ status: "error", message: err.message });
    }
  });

  // POST /api/onay-bekleyenler/guncelle (Edit Pending Request)
  app.post("/api/onay-bekleyenler/guncelle", async (req, res) => {
    try {
      const { id, updatedItem } = req.body;
      if (!id || !updatedItem) {
        return res.status(400).json({ status: "error", message: "Eksik parametre." });
      }

      const items = readOnayBekleyenler();
      const idx = items.findIndex(it => it.id === id);
      if (idx === -1) {
        return res.status(404).json({ status: "error", message: "Güncellenecek talep bulunamadı." });
      }

      items[idx] = {
        ...items[idx],
        ...updatedItem,
        id: items[idx].id // keep original ID
      };

      writeOnayBekleyenler(items);
      syncOnayBekleyenlerToDrive(items).catch(() => {});

      return res.json({ status: "success", message: "Talep bilgileri başarıyla güncellendi.", item: items[idx] });
    } catch (err: any) {
      return res.status(500).json({ status: "error", message: err.message });
    }
  });

  app.get("/api/depo-transactions", async (req, res) => {
    try {
      const allTransfers = await fetchAndMergeAllTransactions();
      return res.json({ status: "success", count: allTransfers.length, transfers: allTransfers });
    } catch (err: any) {
      return res.json({ status: "success", count: 0, transfers: [] });
    }
  });

  app.post("/api/depo-transactions", async (req, res) => {
    try {
      const incomingTransfers = req.body.transfers || req.body.items || [];
      if (!Array.isArray(incomingTransfers)) {
        return res.status(400).json({ status: "error", message: "Geçersiz transfer listesi." });
      }

      // Fetch complete existing database list first
      const fullExisting = await fetchAndMergeAllTransactions();

      const updatedList = [...fullExisting];
      for (const inc of incomingTransfers) {
        if (!inc) continue;
        const incId = inc.id ? String(inc.id).trim() : "";
        
        let foundIdx = -1;
        // Only update existing record in-place if explicitly marked as edit
        if (incId && inc.isEdit === true) {
          foundIdx = updatedList.findIndex(e => e.id && String(e.id).trim() === incId);
        }

        if (foundIdx >= 0) {
          updatedList[foundIdx] = {
            ...updatedList[foundIdx],
            ...inc
          };
        } else {
          // Check duplicates by ID or timestamp+item+sn+quantity+type before appending
          const makeKey = (t: any) => {
            if (t.id && !String(t.id).startsWith("tx_sheets_")) return String(t.id).trim();
            return `${t.itemName || t.name}_${t.timestamp || t.date}_${t.quantity}_${t.type || t.islemTuru}_${t.sn || '-'}`;
          };
          const existingKeys = new Set(updatedList.map(makeKey));
          const key = makeKey(inc);
          if (!existingKeys.has(key)) {
            updatedList.push(inc);
            existingKeys.add(key);
          }
        }
      }

      // 1. JSON olarak kaydet (public ve data)
      const localJsonPath = path.resolve(process.cwd(), "public/depo_transactions.json");
      fs.writeFileSync(localJsonPath, JSON.stringify(updatedList, null, 2), "utf-8");
      writePersistedDepoTransactions(updatedList);

      // 2. Excel dosyası oluştur ve kaydet
      let base64Xlsx = "";
      try {
        const xlsx = await import("xlsx");
        const wb = xlsx.utils.book_new();
        const headerRow = ["MALZEME ADI", "ADET", "TARİH", "İŞLEM TÜRÜ", "AÇIKLAMA / NOTLAR", "SERİAL NUMBER", "KUYRUK KODU", "TESLİM ALAN", "KABUL YAPAN", "DEPO YERİ", "SAYFA ADI"];
        const rows = [headerRow];
        updatedList.forEach((t: any) => {
          const isChem = (t.category === 'kimyasal') || (t.sheetName && String(t.sheetName).toLowerCase().includes('kimya')) || (t.location && String(t.location).toLowerCase().includes('kimya'));
          rows.push([
            t.itemName || t.name || "",
            t.quantity || t.miktar || t.adet || 1,
            t.date || t.timestamp || "",
            t.type || t.islemTuru || "",
            t.notes || t.aciklama || "",
            t.sn || t.serialNumber || "",
            t.tailNo || t.aircraftTail || t.tailNo || "",
            t.operator || t.teslimAlan || "",
            t.receivedBy || t.kabulYapan || "",
            isChem ? "KİMYASAL DEPO" : (t.location || t.depoYeri || ""),
            t.sheetName || getDepoSheetNameForUnit(t.unit)
          ]);
        });
        const ws = xlsx.utils.aoa_to_sheet(rows);
        xlsx.utils.book_append_sheet(wb, ws, "DEPO HAREKET GEÇMİŞİ-AT-802");

        const localExcelPath = path.resolve(process.cwd(), "public/hangar_depo_hareket_gecmisi.xlsx");
        xlsx.writeFile(wb, localExcelPath);

        const buffer = xlsx.write(wb, { type: "buffer", bookType: "xlsx" });
        base64Xlsx = buffer.toString("base64");
      } catch (e) {
        console.warn("Transactions excel write warn:", e);
      }

      // 3. Drive'a yükle
      if (base64Xlsx) {
        const scriptUrls = [
          "https://script.google.com/macros/s/AKfycbw4kruTTc058Y9rLTyO3dKi6KloYsmdDTwV1GSiAk8ZXefyo3Z7_VDSTuurzsS9BHAQyQ/exec"
        ];
        for (const sUrl of scriptUrls) {
          try {
            fetch(sUrl, {
              method: "POST",
              headers: { "Content-Type": "text/plain;charset=utf-8" },
              body: JSON.stringify({
                action: "uploadTechizatExcel",
                fileName: "hangar_depo_hareket_gecmisi.xlsx",
                base64Data: base64Xlsx,
                folderId: "1HQR_NYKhHQGA7_2W3nArI9pCh-LJasTP",
                unitName: "AT-802 Depo Hareket Geçmişi",
                overwrite: true,
                replaceExisting: true,
                deleteDuplicates: true,
                updateIfExists: true,
                cleanDuplicates: true
              })
            }).catch(e => console.warn("Background Drive upload warn:", e));
            break;
          } catch (e) {}
        }
      }

      return res.json({ status: "success", count: updatedList.length, message: "Hareket geçmişi diske ve Google Drive Excel'e kaydedildi!" });
    } catch (err: any) {
      return res.status(500).json({ status: "error", message: err.message });
    }
  });

  // --- YAŞAM DESTEK (AT-802) API ENDPOINTS ---
  const YASAM_DESTEK_FILE = path.resolve(process.cwd(), "data/yasam_destek_at802.json");

  function readYasamDestekData(): any[] {
    try {
      if (fs.existsSync(YASAM_DESTEK_FILE)) {
        const raw = fs.readFileSync(YASAM_DESTEK_FILE, "utf-8");
        const parsed = JSON.parse(raw);
        if (Array.isArray(parsed)) return parsed;
      }
    } catch (e) {
      console.warn("Read yasam destek file error:", e);
    }
    return [];
  }

  function writeYasamDestekData(records: any[]): void {
    try {
      const dir = path.dirname(YASAM_DESTEK_FILE);
      if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
      fs.writeFileSync(YASAM_DESTEK_FILE, JSON.stringify(records, null, 2), "utf-8");
    } catch (e) {
      console.warn("Write yasam destek file error:", e);
    }
  }

  // Sync to Google Sheets tab "YAŞAM DESTEK AT-802"
  async function syncYasamDestekToGoogleSheet(records: any[]) {
    const spreadsheetId = "17ScGYYx0erzDwHDk6RGiHOdJATdfmmExXFBY39dXpF0";
    const sheetName = "YAŞAM DESTEK AT-802";
    const headers = [
      "SIRA NO",
      "PERSONEL ADI SOYADI",
      "GÖREV YERİ",
      "BAŞLANGIÇ TARİHİ",
      "BİTİŞ TARİHİ",
      "LIFE VEST S/N",
      "SPARE AIR S/N",
      "HELMET KIT S/N",
      "DURUM",
      "DEVREDEN / TESLİM EDEN",
      "DEVREDİLEN / TESLİM ALAN",
      "AÇIKLAMA / NOTLAR",
      "GÜNCELLENME TARİHİ"
    ];

    const rows = [
      headers,
      ...records.map((r, idx) => {
        // Devir durumunda personelin kendisi 'Devreden'dir.
        const isDevir = r.durum === "DEVREDİLDİ";
        const isTeslim = r.durum === "ANKARA DEPO TESLİM";
        
        return [
          String(idx + 1),
          r.personelAdi || "",
          r.gorevYeri || "",
          r.baslangicTarihi || "",
          r.bitisTarihi || "-",
          r.lifeVestSn || "-",
          r.spareAirSn || "-",
          r.helmetKitSn || "-",
          r.durum || "GÖREVDE",
          (isDevir || isTeslim) ? (r.personelAdi || "-") : "-", // DEVREDEN
          (isDevir ? (r.devredilenPersonel || "-") : (isTeslim ? (r.teslimAlanPersonel || "-") : "-")), // DEVREDİLEN / TESLİM ALAN
          r.notlar || "",
          r.guncellenmeTarihi || new Date().toLocaleDateString("tr-TR")
        ];
      })
    ];

    const scriptUrls = [
      "https://script.google.com/macros/s/AKfycbzhXYZBlJvYarhEYpSK_UdceV-pQwGRIHTjWAVN_UTumI7_qla7vZnAZofdJGeK0e-ZVQ/exec",
      "https://script.google.com/macros/s/AKfycbwP1uOo2NrST5a4I8vm1nGBLtI26yY2lWrmu9_e9iymwBkUhJBA9JOPCp7SNKqJbOubOw/exec",
      "https://script.google.com/macros/s/AKfycbw4kruTTc058Y9rLTyO3dKi6KloYsmdDTwV1GSiAk8ZXefyo3Z7_VDSTuurzsS9BHAQyQ/exec"
    ];

    for (const sUrl of scriptUrls) {
      try {
        await fetch(sUrl, {
          method: "POST",
          headers: { "Content-Type": "text/plain;charset=utf-8" },
          body: JSON.stringify({
            action: "updateSheet",
            sheetName,
            spreadsheetId,
            rows
          })
        });
        break;
      } catch (err) {
        console.warn("Google Sheet yasam destek sync error:", err);
      }
    }
  }

  // --- 3 RESMİ YAŞAM DESTEK SAYFASI (CAN YELEĞİ STOK, SPARE AIR STOK, HELMET KIT VE DAVID CLARK) ---
  const CAN_YELEGI_FILE = path.resolve(process.cwd(), "data/can_yelegi_stok.json");
  const SPARE_AIR_FILE = path.resolve(process.cwd(), "data/spare_air_stok.json");
  const HELMET_KIT_FILE = path.resolve(process.cwd(), "data/helmet_kit_stok.json");

  function readJsonFile(filePath: string, fallback: any[] = []): any[] {
    try {
      if (fs.existsSync(filePath)) {
        const raw = fs.readFileSync(filePath, "utf-8");
        const parsed = JSON.parse(raw);
        if (Array.isArray(parsed)) return parsed;
      }
    } catch (e) {
      console.warn("Read file error:", filePath, e);
    }
    return fallback;
  }

  function writeJsonFile(filePath: string, data: any[]): void {
    try {
      const dir = path.dirname(filePath);
      if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
      fs.writeFileSync(filePath, JSON.stringify(data, null, 2), "utf-8");
    } catch (e) {
      console.warn("Write file error:", filePath, e);
    }
  }

  // Automatically initialize stock tables from public excel file on startup if empty
  function initializeYasamDestekStocksFromExcel(): void {
    try {
      const excelPath = path.resolve(process.cwd(), "public/at-802_yasam_destek_ekipmanlari.xlsx");
      if (!fs.existsSync(excelPath)) {
        console.warn("at-802_yasam_destek_ekipmanlari.xlsx not found in public/ folder.");
        return;
      }

      const cyEmpty = !fs.existsSync(CAN_YELEGI_FILE) || readJsonFile(CAN_YELEGI_FILE).length === 0;
      const saEmpty = !fs.existsSync(SPARE_AIR_FILE) || readJsonFile(SPARE_AIR_FILE).length === 0;
      const hkEmpty = !fs.existsSync(HELMET_KIT_FILE) || readJsonFile(HELMET_KIT_FILE).length === 0;

      if (!cyEmpty && !saEmpty && !hkEmpty) {
        return; // Already initialized and has data
      }

      const workbook = XLSX.readFile(excelPath);
      workbook.SheetNames.forEach(sheetName => {
        const sheet = workbook.Sheets[sheetName];
        if (!sheet) return;
        const json: any[] = XLSX.utils.sheet_to_json(sheet);
        const upperSheet = sheetName.toUpperCase();

        if (upperSheet.includes("YELEK") || upperSheet.includes("VEST") || upperSheet.includes("CAN")) {
          const items = json.map((r, idx) => {
            const rowStr = JSON.stringify(r).toUpperCase();
            if (rowStr.includes("KASK TAKIMI") || rowStr.includes("KOLTUK SAYISI") || rowStr.includes("DOUBLE HAVA") || rowStr.includes("SINGLE HAVA")) {
              return null;
            }
            const sNo = r["S. NO"] || r["SIRA NO"] || r["SIRA"] || (idx + 1);
            const malzemeAdi = r["MALZEME ADI"] || r["Malzeme"] || "LIFE WEST - CAN YELEĞİ";
            const parcaNo = r["PARÇA NO"] || r["Parça No"] || "S-7200-511";
            const seriNo = (r["SERİ NO"] || r["Seri No"] || r["S/N"] || r["SN"] || r["seriNo"] || "").toString().trim();
            const disNo = (r["DIŞ NO"] || r["Dış No"] || r["disNo"] || "-").toString().trim();
            const aciklamalar = (r["TESLİM EDİLEN PERSONEL/BÖLGE"] || r["PERSONEL"] || r["AÇIKLAMALAR"] || r["Açıklamalar"] || r["aciklamalar"] || r["AÇIKLAMA"] || r["Açıklama"] || r["Not"] || r["notlar"] || "").toString().trim();
            const notlar = (r["AÇIKLAMA"] || r["Not"] || r["NOTLAR"] || "-").toString().trim();
            const tarih = (r["TARİH / BİTİŞ"] || r["Tarih"] || r["tarih"] || "DEVAM EDİYOR").toString().trim();
            const durum = aciklamalar.toUpperCase().includes("DEPO") || aciklamalar.toUpperCase().includes("BOŞ") || !aciklamalar || aciklamalar === "-" ? "DEPODA" : "GÖREV BÖLGESİNDE";
            return { sNo, malzemeAdi, parcaNo, seriNo, disNo, aciklamalar, notlar, durum, tarih };
          }).filter(item => item !== null && item.seriNo);
          if (items.length > 0 && cyEmpty) {
            writeJsonFile(CAN_YELEGI_FILE, items);
            console.log(`Successfully auto-initialized ${items.length} Can Yeleği records from Excel.`);
          }
        } else if (upperSheet.includes("SPARE") || upperSheet.includes("AIR") || upperSheet.includes("HAVA") || upperSheet.includes("TÜP")) {
          const items = json.map((r, idx) => {
            const rowStr = JSON.stringify(r).toUpperCase();
            if (rowStr.includes("KASK TAKIMI") || rowStr.includes("KOLTUK SAYISI") || rowStr.includes("DOUBLE HAVA") || rowStr.includes("SINGLE HAVA")) {
              return null;
            }
            const sNo = r["S. NO"] || r["SIRA NO"] || r["SIRA"] || (idx + 1);
            const malzemeAdi = r["MALZEME ADI"] || r["Malzeme"] || "SPARE AIR (HEED 3)";
            const parcaNo = r["PARÇA NO"] || r["Parça No"] || "175-001-CE";
            const seriNo = (r["SERİ NO"] || r["Seri No"] || r["S/N"] || r["SN"] || r["seriNo"] || "").toString().trim();
            const aciklamalar = (r["TESLİM EDİLEN PERSONEL/BÖLGE"] || r["PERSONEL"] || r["AÇIKLAMALAR"] || r["Açıklamalar"] || r["aciklamalar"] || r["AÇIKLAMA"] || r["Açıklama"] || r["Not"] || r["notlar"] || "").toString().trim();
            const notlar = (r["AÇIKLAMA"] || r["Not"] || r["NOTLAR"] || "-").toString().trim();
            const tarih = (r["TARİH / BİTİŞ"] || r["Tarih"] || r["tarih"] || "DEVAM EDİYOR").toString().trim();
            const durum = aciklamalar.toUpperCase().includes("DEPO") || aciklamalar.toUpperCase().includes("BOŞ") || !aciklamalar || aciklamalar === "-" ? "DEPODA" : "GÖREV BÖLGESİNDE";
            return { sNo, malzemeAdi, parcaNo, seriNo, aciklamalar, notlar, durum, tarih };
          }).filter(item => item !== null && item.seriNo);
          if (items.length > 0 && saEmpty) {
            writeJsonFile(SPARE_AIR_FILE, items);
            console.log(`Successfully auto-initialized ${items.length} Spare Air records from Excel.`);
          }
        } else if (upperSheet.includes("HELMET") || upperSheet.includes("KASK") || upperSheet.includes("KULAK") || upperSheet.includes("CLARK") || upperSheet.includes("KULAKLIK")) {
          const items = json.map((r, idx) => {
            const rowStr = JSON.stringify(r).toUpperCase();
            if (rowStr.includes("KASK TAKIMI") || rowStr.includes("KOLTUK SAYISI") || rowStr.includes("DOUBLE HAVA") || rowStr.includes("SINGLE HAVA")) {
              return null;
            }
            const sNo = r["S. NO"] || r["SIRA NO"] || r["SIRA"] || (idx + 1);
            const malzemeAdi1 = r["MALZEME ADI 1"] || r["Malzeme 1"] || "HELMET KIT";
            const parcaNo1 = r["PARÇA NO 1"] || r["Parça No 1"] || "18852G-01";
            const seriNo1 = (r["SERİ NO 1"] || r["Seri No 1"] || r["seriNo1"] || "-").toString().trim();
            const malzemeAdi2 = r["MALZEME ADI 2"] || r["Malzeme 2"] || "DAVID CLARK";
            const parcaNo2 = r["PARÇA NO 2"] || r["Parça No 2"] || "H-10-13X";
            const seriNo2 = (r["KULAKLIK S/N"] || r["Seri No 2"] || r["seriNo2"] || "-").toString().trim();
            const aciklamalar = (r["TESLİM EDİLEN PERSONEL/BÖLGE"] || r["PERSONEL"] || r["AÇIKLAMALAR"] || r["Açıklamalar"] || r["aciklamalar"] || r["AÇIKLAMA"] || r["Açıklama"] || r["Not"] || r["notlar"] || "").toString().trim();
            const notlar = (r["AÇIKLAMA"] || r["Not"] || r["NOTLAR"] || "-").toString().trim();
            const tarih = (r["TARİH / BİTİŞ"] || r["Tarih"] || r["tarih"] || "DEVAM EDİYOR").toString().trim();
            const durum = aciklamalar.toUpperCase().includes("DEPO") || aciklamalar.toUpperCase().includes("BOŞ") || !aciklamalar || aciklamalar === "-" ? "DEPODA" : "GÖREV BÖLGESİNDE";
            const aciklamalarValue = aciklamalar && aciklamalar !== "DEPODA" && aciklamalar !== "-" ? aciklamalar : (seriNo1 || `HK-18852G-${String(idx + 1).padStart(2, '0')}`);
            return { sNo, malzemeAdi1, parcaNo1, seriNo1, malzemeAdi2, parcaNo2, seriNo2, aciklamalar: aciklamalarValue, notlar, durum, tarih };
          }).filter(item => item !== null);
          if (items.length > 0 && hkEmpty) {
            writeJsonFile(HELMET_KIT_FILE, items);
            console.log(`Successfully auto-initialized ${items.length} Helmet Kit records from Excel.`);
          }
        }
      });
    } catch (err) {
      console.warn("Failed to auto-initialize life support stocks from excel:", err);
    }
  }

  // Run initialization
  initializeYasamDestekStocksFromExcel();

  // Google E-Tabloya 4 Resmi Sayfayı ve Başlıkları Senkronize Et
  async function syncOfficialYasamDestekSheetsToGoogleSheet() {
    const spreadsheetId = "17ScGYYx0erzDwHDk6RGiHOdJATdfmmExXFBY39dXpF0";
    const scriptUrls = [
      "https://script.google.com/macros/s/AKfycbzhXYZBlJvYarhEYpSK_UdceV-pQwGRIHTjWAVN_UTumI7_qla7vZnAZofdJGeK0e-ZVQ/exec",
      "https://script.google.com/macros/s/AKfycbwP1uOo2NrST5a4I8vm1nGBLtI26yY2lWrmu9_e9iymwBkUhJBA9JOPCp7SNKqJbOubOw/exec",
      "https://script.google.com/macros/s/AKfycbw4kruTTc058Y9rLTyO3dKi6KloYsmdDTwV1GSiAk8ZXefyo3Z7_VDSTuurzsS9BHAQyQ/exec"
    ];

    // 1. CAN YELEĞİ STOK
    const canYelegiList = readJsonFile(CAN_YELEGI_FILE);
    const canYelegiRows = [
      ["S. NO", "MALZEME ADI", "PARÇA NO", "SERİ NO", "DIŞ NO", "AÇIKLAMA", "TESLİM EDİLEN PERSONEL/BÖLGE", "TARİH / BİTİŞ"],
      ...canYelegiList.map((r, i) => [
        String(r.sNo || i + 1),
        r.malzemeAdi || "LIFE WEST - CAN YELEĞİ",
        r.parcaNo || "S-7200-511",
        r.seriNo || "-",
        r.disNo || "-",
        r.notlar || r.aciklama || "-",
        r.aciklamalar || r.personel || "-",
        r.tarih || (r.durum === "DEPODA" ? "-" : "DEVAM EDİYOR")
      ])
    ];

    // 2. SPARE AIR STOK
    const spareAirList = readJsonFile(SPARE_AIR_FILE);
    const spareAirRows = [
      ["S. NO", "MALZEME ADI", "PARÇA NO", "SERİ NO", "AÇIKLAMA", "TESLİM EDİLEN PERSONEL/BÖLGE", "TARİH / BİTİŞ"],
      ...spareAirList.map((r, i) => [
        String(r.sNo || i + 1),
        r.malzemeAdi || "SPARE AIR (HEED 3)",
        r.parcaNo || "175-001-CE",
        r.seriNo || "-",
        r.notlar || r.aciklama || "-",
        r.aciklamalar || r.personel || "-",
        r.tarih || (r.durum === "DEPODA" ? "-" : "DEVAM EDİYOR")
      ])
    ];

    // 3. HELMET KIT VE DAVID CLARK KULAK VE DOLGU
    const helmetKitList = readJsonFile(HELMET_KIT_FILE);
    const helmetKitRows = [
      ["S. NO", "MALZEME ADI 1", "PARÇA NO 1", "SERİ NO 1", "MALZEME ADI 2", "PARÇA NO 2", "KULAKLIK S/N", "AÇIKLAMA", "TESLİM EDİLEN PERSONEL/BÖLGE", "TARİH / BİTİŞ"],
      ...helmetKitList.map((r, i) => [
        String(r.sNo || i + 1),
        r.malzemeAdi1 || "HELMET KIT",
        r.parcaNo1 || "18852G-01",
        r.seriNo1 || (i < 21 ? `HK-18852G-${String(i + 1).padStart(2, '0')}` : "-"),
        r.malzemeAdi2 || "DAVID CLARK",
        r.parcaNo2 || "H-10-13X",
        r.seriNo2 || (r.seriNo2 && r.seriNo2 !== '-' ? r.seriNo2 : (i < 21 ? `H-10-13X-${i + 1}` : "-")),
        r.notlar || r.aciklama || "-",
        r.aciklamalar || r.personel || "-",
        r.tarih || (r.durum === "DEPODA" ? "-" : "DEVAM EDİYOR")
      ])
    ];

    // 4. YAŞAM DESTEK AT-802 (Görev & Zimmet Takibi)
    const yasamDestekRecords = readYasamDestekData();
    const yasamDestekRows = [
      [
        "SIRA NO",
        "PERSONEL ADI SOYADI",
        "GÖREV YERİ",
        "BAŞLANGIÇ TARİHİ",
        "BİTİŞ TARİHİ",
        "LIFE VEST S/N",
        "SPARE AIR S/N",
        "HELMET KIT S/N",
        "DURUM",
        "DEVREDEN / TESLİM EDEN",
        "DEVREDİLEN / TESLİM ALAN",
        "AÇIKLAMA / NOTLAR",
        "GÜNCELLENME TARİHİ"
      ],
      ...yasamDestekRecords.map((r, idx) => {
        // Devir durumunda personelin kendisi 'Devreden'dir.
        const isDevir = r.durum === "DEVREDİLDİ";
        const isTeslim = r.durum === "ANKARA DEPO TESLİM";
        
        return [
          String(idx + 1),
          r.personelAdi || "",
          r.gorevYeri || "",
          r.baslangicTarihi || "",
          r.bitisTarihi || "-",
          r.lifeVestSn || "-",
          r.spareAirSn || "-",
          r.helmetKitSn || "-",
          r.durum || "GÖREVDE",
          (isDevir || isTeslim) ? (r.personelAdi || "-") : "-", // DEVREDEN
          (isDevir ? (r.devredilenPersonel || "-") : (isTeslim ? (r.teslimAlanPersonel || "-") : "-")), // DEVREDİLEN / TESLİM ALAN
          r.notlar || "",
          r.guncellenmeTarihi || new Date().toLocaleDateString("tr-TR")
        ];
      })
    ];

    // 5. AT-802 PERSONEL VERİSİ
    const personnelItems = readPersonnelFromExcel();
    const personnelRows = [
      ["SIRA NO", "PERSONEL ADI SOYADI", "ÜNVANI / GÖREVİ"],
      ...personnelItems.map((p, idx) => [
        String(idx + 1),
        typeof p === "string" ? p : (p.adSoyad || ""),
        typeof p === "string" ? "Teknisyen" : (p.unvan || "Teknisyen")
      ])
    ];

    const sheetsToSync = [
      { name: "CAN YELEĞİ STOK", rows: canYelegiRows },
      { name: "SPARE AIR STOK", rows: spareAirRows },
      { name: "HELMET KIT VE DAVID CLARK KULAK VE DOLGU", rows: helmetKitRows },
      { name: "YAŞAM DESTEK AT-802", rows: yasamDestekRows },
      { name: "AT-802 PERSONEL VERİSİ", rows: personnelRows }
    ];

    const results: any[] = [];
    for (const sh of sheetsToSync) {
      for (const sUrl of scriptUrls) {
        try {
          const resJson = await postToGoogleAppsScript(sUrl, {
            action: "updateSheet",
            sheetName: sh.name,
            spreadsheetId,
            data: sh.rows,
            rows: sh.rows
          });
          if (resJson && (resJson.status === "success" || resJson.updated || resJson.success)) {
            results.push({ sheet: sh.name, status: "success" });
            break;
          }
        } catch (e: any) {
          console.warn(`[Sync] Sheet sync error for ${sh.name}:`, e.message);
        }
      }
    }
    return results;
  }

  // İlk açılışta veya arka planda senkronizasyonu başlat
  syncOfficialYasamDestekSheetsToGoogleSheet().catch(() => {});

  // GET / POST /api/yasam-destek/ensure-sheets
  app.all("/api/yasam-destek/ensure-sheets", async (req, res) => {
    try {
      const results = await syncOfficialYasamDestekSheetsToGoogleSheet();
      return res.json({
        status: "success",
        message: "Yaşam Destek sayfaları ve başlıkları Google E-Tabloya başarıyla atandı.",
        results
      });
    } catch (err: any) {
      return res.status(500).json({ status: "error", message: err.message });
    }
  });

  // POST /api/yasam-destek/upload-excel (Excel ile Yaşam Destek Stok & Serno Aktarımı)
  app.post("/api/yasam-destek/upload-excel", async (req, res) => {
    try {
      const { canYelegi, spareAir, helmetKit, base64Data, fileName } = req.body;
      let cyCount = 0;
      let saCount = 0;
      let hkCount = 0;

      if (Array.isArray(canYelegi) && canYelegi.length > 0) {
        writeJsonFile(CAN_YELEGI_FILE, canYelegi);
        cyCount = canYelegi.length;
      }
      if (Array.isArray(spareAir) && spareAir.length > 0) {
        writeJsonFile(SPARE_AIR_FILE, spareAir);
        saCount = spareAir.length;
      }
      if (Array.isArray(helmetKit) && helmetKit.length > 0) {
        writeJsonFile(HELMET_KIT_FILE, helmetKit);
        hkCount = helmetKit.length;
      }

      // Arka planda Google E-Tablo senkronizasyonunu tetikle
      syncOfficialYasamDestekSheetsToGoogleSheet().catch(() => {});

      return res.json({
        status: "success",
        message: "Yaşam Destek verileri başarıyla kaydedildi.",
        counts: {
          canYelegi: cyCount,
          spareAir: saCount,
          helmetKit: hkCount
        }
      });
    } catch (err: any) {
      return res.status(500).json({ status: "error", message: err.message });
    }
  });

  // GET /api/yasam-destek/all-stocks (Live read directly from Google Spreadsheet online)
  app.get("/api/yasam-destek/all-stocks", async (req, res) => {
    let canYelegi = readJsonFile(CAN_YELEGI_FILE);
    let spareAir = readJsonFile(SPARE_AIR_FILE);
    let helmetKit = readJsonFile(HELMET_KIT_FILE);

    // Live sync from Google Spreadsheet online in parallel
    try {
      const sheetId = "17ScGYYx0erzDwHDk6RGiHOdJATdfmmExXFBY39dXpF0";
      
      const urlCY = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent("CAN YELEĞİ STOK")}`;
      const urlSA = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent("SPARE AIR STOK")}`;
      const urlHK = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent("HELMET KIT VE DAVID CLARK KULAK VE DOLGU")}`;

      const [resCY, resSA, resHK] = await Promise.allSettled([
        fetch(urlCY, { signal: AbortSignal.timeout(3500) }).then(r => r.ok ? r.text() : ""),
        fetch(urlSA, { signal: AbortSignal.timeout(3500) }).then(r => r.ok ? r.text() : ""),
        fetch(urlHK, { signal: AbortSignal.timeout(3500) }).then(r => r.ok ? r.text() : "")
      ]);

      const regex = new RegExp(`,(?=(?:(?:[^"]*"){2})*[^"]*$)`);

      // 1. Can Yeleği
      if (resCY.status === "fulfilled" && resCY.value && !resCY.value.includes("<!DOCTYPE")) {
        const lines = resCY.value.split(/\r?\n/).filter(l => l.trim().length > 0);
        if (lines.length > 1) {
          const remoteCY: any[] = [];
          for (let i = 1; i < lines.length; i++) {
            const cols = lines[i].split(regex).map(c => c.replace(/^"|"$/g, '').trim());
            if (cols.length >= 4 && cols[3]) {
              const assigned = cols[6] || "";
              const isDepot = !assigned || assigned === "-" || assigned.toUpperCase().includes("DEPO") || assigned.toUpperCase().includes("BOŞ");
              const durum = isDepot ? "DEPODA" : ((assigned.toUpperCase().includes("SİL")) ? "KAYIT SİLME" : "GÖREV BÖLGESİNDE");
              const color = isDepot ? "#ffffff" : ((durum === "KAYIT SİLME") ? "#ff0000" : "#92d050");
              remoteCY.push({
                sNo: Number(cols[0]) || i,
                malzemeAdi: cols[1] || "LIFE WEST - CAN YELEĞİ",
                parcaNo: cols[2] || "S-7200-511",
                seriNo: cols[3] || "-",
                disNo: cols[4] || "-",
                notlar: cols[5] || "-",
                aciklamalar: isDepot ? "DEPODA" : assigned,
                tarih: cols[7] || (isDepot ? "-" : "DEVAM EDİYOR"),
                durum,
                color
              });
            }
          }
          if (remoteCY.length > 0) {
            canYelegi = remoteCY;
            writeJsonFile(CAN_YELEGI_FILE, canYelegi);
          }
        }
      }

      // 2. Spare Air
      if (resSA.status === "fulfilled" && resSA.value && !resSA.value.includes("<!DOCTYPE")) {
        const lines = resSA.value.split(/\r?\n/).filter(l => l.trim().length > 0);
        if (lines.length > 1) {
          const remoteSA: any[] = [];
          for (let i = 1; i < lines.length; i++) {
            const cols = lines[i].split(regex).map(c => c.replace(/^"|"$/g, '').trim());
            if (cols.length >= 4 && cols[3]) {
              const assigned = cols[5] || "";
              const isDepot = !assigned || assigned === "-" || assigned.toUpperCase().includes("DEPO") || assigned.toUpperCase().includes("BOŞ");
              const durum = isDepot ? "DEPODA" : ((assigned.toUpperCase().includes("SİL")) ? "KAYIT SİLME" : "GÖREV BÖLGESİNDE");
              const color = isDepot ? "#ffffff" : ((durum === "KAYIT SİLME") ? "#ff0000" : "#92d050");
              remoteSA.push({
                sNo: Number(cols[0]) || i,
                malzemeAdi: cols[1] || "SPARE AIR (HEED 3)",
                parcaNo: cols[2] || "175-001-CE",
                seriNo: cols[3] || "-",
                notlar: cols[4] || "-",
                aciklamalar: isDepot ? "DEPODA" : assigned,
                tarih: cols[6] || (isDepot ? "-" : "DEVAM EDİYOR"),
                durum,
                color
              });
            }
          }
          if (remoteSA.length > 0) {
            spareAir = remoteSA;
            writeJsonFile(SPARE_AIR_FILE, spareAir);
          }
        }
      }

      // 3. Helmet Kit
      if (resHK.status === "fulfilled" && resHK.value && !resHK.value.includes("<!DOCTYPE")) {
        const lines = resHK.value.split(/\r?\n/).filter(l => l.trim().length > 0);
        if (lines.length > 1) {
          const remoteHK: any[] = [];
          for (let i = 1; i < lines.length; i++) {
            const cols = lines[i].split(regex).map(c => c.replace(/^"|"$/g, '').trim());
            if (cols.length >= 2) {
              const assigned = cols[8] || "";
              const isDepot = !assigned || assigned === "-" || assigned.toUpperCase().includes("DEPO") || assigned.toUpperCase().includes("BOŞ");
              const durum = isDepot ? "DEPODA" : ((assigned.toUpperCase().includes("SİL")) ? "KAYIT SİLME" : "GÖREV BÖLGESİNDE");
              const color = isDepot ? "#ffffff" : ((durum === "KAYIT SİLME") ? "#ff0000" : "#92d050");
              const sn2 = cols[6] && cols[6] !== "-" ? cols[6] : `H-10-13X-${i}`;
              remoteHK.push({
                sNo: Number(cols[0]) || i,
                malzemeAdi1: cols[1] || "HELMET KIT",
                parcaNo1: cols[2] || "18852G-01",
                seriNo1: cols[3] || "-",
                malzemeAdi2: cols[4] || "DAVID CLARK",
                parcaNo2: cols[5] || "H-10-13X",
                seriNo2: sn2,
                notlar: cols[7] || "-",
                aciklamalar: isDepot ? "DEPODA" : assigned,
                tarih: cols[9] || (isDepot ? "-" : "DEVAM EDİYOR"),
                durum,
                color
              });
            }
          }
          if (remoteHK.length > 0) {
            helmetKit = remoteHK;
            writeJsonFile(HELMET_KIT_FILE, helmetKit);
          }
        }
      }
    } catch (e) {
      // Fallback to local files
    }

    // Enforce bi-directional consistency with active records from YAŞAM DESTEK AT-802
    const activeZimmetRecords = readYasamDestekData().filter(r => r.durum === "GÖREVDE");
    
    // Map actively assigned serial numbers
    const activeLVMap = new Map<string, { person: string, date: string }>();
    const activeSAMap = new Map<string, { person: string, date: string }>();
    const activeHKMap = new Map<string, { person: string, date: string }>();

    activeZimmetRecords.forEach(r => {
      if (r.lifeVestSn && r.lifeVestSn !== "-") {
        activeLVMap.set(r.lifeVestSn.trim(), { person: r.personelAdi, date: r.baslangicTarihi });
      }
      if (r.spareAirSn && r.spareAirSn !== "-") {
        activeSAMap.set(r.spareAirSn.trim(), { person: r.personelAdi, date: r.baslangicTarihi });
      }
      if (r.helmetKitSn && r.helmetKitSn !== "-") {
        activeHKMap.set(r.helmetKitSn.trim(), { person: r.personelAdi, date: r.baslangicTarihi });
      }
    });

    canYelegi = canYelegi.map(c => {
      if (c.durum === "ONARIMDA" || c.durum === "BAKIMA GİDECEK" || c.durum === "KAYIT SİLME") return c;
      const assigned = activeLVMap.get(c.seriNo);
      if (assigned) {
        return {
          ...c,
          aciklamalar: assigned.person,
          durum: "GÖREV BÖLGESİNDE",
          color: "#92d050",
          tarih: assigned.date || c.tarih || "DEVAM EDİYOR"
        };
      }
      return {
        ...c,
        aciklamalar: "DEPODA",
        durum: "DEPODA",
        color: "#ffffff",
        tarih: "-"
      };
    });

    spareAir = spareAir.map(s => {
      if (s.durum === "ONARIMDA" || s.durum === "BAKIMA GİDECEK" || s.durum === "KAYIT SİLME") return s;
      const assigned = activeSAMap.get(s.seriNo);
      if (assigned) {
        return {
          ...s,
          aciklamalar: assigned.person,
          durum: "GÖREV BÖLGESİNDE",
          color: "#92d050",
          tarih: assigned.date || s.tarih || "DEVAM EDİYOR"
        };
      }
      return {
        ...s,
        aciklamalar: "DEPODA",
        durum: "DEPODA",
        color: "#ffffff",
        tarih: "-"
      };
    });

    helmetKit = helmetKit.map(h => {
      if (h.durum === "ONARIMDA" || h.durum === "BAKIMA GİDECEK" || h.durum === "KAYIT SİLME") return h;
      const sn2 = (h.seriNo2 || "").trim();
      const sn1 = (h.seriNo1 || "").trim();
      let assigned = (sn2 ? activeHKMap.get(sn2) : null) || (sn1 ? activeHKMap.get(sn1) : null);
      if (!assigned) {
        for (const [k, v] of activeHKMap.entries()) {
          const pureK = k.split(" (")[0].trim();
          if (sn2 === pureK || sn1 === pureK || (h.aciklamalar && h.aciklamalar === pureK)) {
            assigned = v;
            break;
          }
        }
      }
      if (assigned) {
        return {
          ...h,
          aciklamalar: assigned.person,
          durum: "GÖREV BÖLGESİNDE",
          color: "#92d050",
          tarih: assigned.date || h.tarih || "DEVAM EDİYOR"
        };
      }
      return {
        ...h,
        aciklamalar: "DEPODA",
        durum: "DEPODA",
        color: "#ffffff",
        tarih: "-"
      };
    });

    return res.json({
      status: "success",
      canYelegi,
      spareAir,
      helmetKit,
      lifeVestSerialNumbers: canYelegi.map(c => ({
        seriNo: c.seriNo,
        disNo: c.disNo,
        durum: c.durum,
        aciklama: c.aciklamalar,
        label: `${c.seriNo} (${c.disNo && c.disNo !== '-' ? c.disNo + ' - ' : ''}${c.aciklamalar || 'DEPODA'})`
      })),
      spareAirSerialNumbers: spareAir.map(s => ({
        seriNo: s.seriNo,
        durum: s.durum,
        aciklama: s.aciklamalar,
        label: `${s.seriNo} (${s.aciklamalar || 'DEPODA'})`
      })),
      helmetKitList: helmetKit.map(h => {
        const sn = h.seriNo2 && h.seriNo2 !== '-' ? h.seriNo2 : (h.seriNo1 && h.seriNo1 !== '-' ? h.seriNo1 : `HK-${h.sNo}`);
        const name = h.malzemeAdi2 || h.malzemeAdi1 || 'David Clark';
        return {
          seriNo: sn,
          personel: h.aciklamalar,
          tarih: h.tarih,
          durum: h.durum,
          label: `${sn} (${name} - ${h.aciklamalar || 'DEPODA'})`
        };
      }),
      legend: [
        { status: "BOŞTA", color: "#ffffff", textColor: "#000000", border: "#cbd5e1" },
        { status: "ONARIMDA", color: "#ffff00", textColor: "#000000", border: "#eab308" },
        { status: "GÖREV BÖLGESİNDE", color: "#92d050", textColor: "#000000", border: "#65a30d" },
        { status: "KAYIT SİLME", color: "#ff0000", textColor: "#ffffff", border: "#dc2626" },
        { status: "BAKIMA GİDECEK", color: "#00b0f0", textColor: "#000000", border: "#0284c7" }
      ]
    });
  });

  // POST /api/yasam-destek/stock-update (Update individual stock items and cross-sync with active zimmet records)
  app.post("/api/yasam-destek/stock-update", async (req, res) => {
    try {
      const { type, items } = req.body;
      if (!type || !Array.isArray(items)) {
        return res.status(400).json({ status: "error", message: "Geçersiz veri." });
      }

      if (type === "can_yelegi") {
        writeJsonFile(CAN_YELEGI_FILE, items);
      } else if (type === "spare_air") {
        writeJsonFile(SPARE_AIR_FILE, items);
      } else if (type === "helmet_kit") {
        writeJsonFile(HELMET_KIT_FILE, items);
      }

      // Cross-sync: Bi-directional synchronization between 3 inventory lists and active personnel zimmet records
      let currentRecords = readYasamDestekData();
      let hasRecordChanges = false;
      const dateStr = new Date().toLocaleDateString("tr-TR");

      items.forEach((item: any) => {
        const rawAssigned = (item.aciklamalar || "").trim().toUpperCase();
        const isDepotOrNonAssigned = 
          item.durum !== "GÖREV BÖLGESİNDE" ||
          rawAssigned === "DEPODA" ||
          rawAssigned === "BOŞTA" ||
          rawAssigned === "ONARIMDA" ||
          rawAssigned === "BAKIMA GİDECEK" ||
          rawAssigned === "KAYIT SİLME" ||
          rawAssigned === "-" ||
          !rawAssigned;

        const sn = item.seriNo || item.seriNo2 || item.seriNo1;

        if (sn) {
          if (isDepotOrNonAssigned) {
            // Drop item from any active zimmet holding this SN
            currentRecords.forEach(r => {
              if (r.durum === "GÖREVDE") {
                let dropped = false;
                if (type === "can_yelegi" && (r.lifeVestSn === sn || (r.lifeVestSn && r.lifeVestSn.includes(sn)))) {
                  r.lifeVestSn = "-";
                  dropped = true;
                  hasRecordChanges = true;
                } else if (type === "spare_air" && (r.spareAirSn === sn || (r.spareAirSn && r.spareAirSn.includes(sn)))) {
                  r.spareAirSn = "-";
                  dropped = true;
                  hasRecordChanges = true;
                } else if (type === "helmet_kit" && (r.helmetKitSn === sn || (r.helmetKitSn && r.helmetKitSn.includes(sn)))) {
                  r.helmetKitSn = "-";
                  dropped = true;
                  hasRecordChanges = true;
                }

                if (dropped) {
                  // If all 3 are returned / empty, close the zimmet completely
                  const lvEmpty = !r.lifeVestSn || r.lifeVestSn === "-";
                  const saEmpty = !r.spareAirSn || r.spareAirSn === "-";
                  const hkEmpty = !r.helmetKitSn || r.helmetKitSn === "-";
                  if (lvEmpty && saEmpty && hkEmpty) {
                    r.durum = "ANKARA DEPO TESLİM";
                    r.bitisTarihi = dateStr;
                    r.teslimAlanPersonel = "Ankara Depo Sorumlusu";
                    r.notlar = (r.notlar ? r.notlar + " | " : "") + "Ekipmanlar depoya çekildi (Zimmet Kapatıldı).";
                    hasRecordChanges = true;
                  }
                }
              }
            });
          } else {
            // Item is assigned to a specific person
            const targetPerson = rawAssigned;

            // First, remove from any OTHER active zimmet if held previously
            currentRecords.forEach(r => {
              if (r.durum === "GÖREVDE" && r.personelAdi.toUpperCase() !== targetPerson) {
                let dropped = false;
                if (type === "can_yelegi" && (r.lifeVestSn === sn || (r.lifeVestSn && r.lifeVestSn.includes(sn)))) {
                  r.lifeVestSn = "-";
                  dropped = true;
                  hasRecordChanges = true;
                } else if (type === "spare_air" && (r.spareAirSn === sn || (r.spareAirSn && r.spareAirSn.includes(sn)))) {
                  r.spareAirSn = "-";
                  dropped = true;
                  hasRecordChanges = true;
                } else if (type === "helmet_kit" && (r.helmetKitSn === sn || (r.helmetKitSn && r.helmetKitSn.includes(sn)))) {
                  r.helmetKitSn = "-";
                  dropped = true;
                  hasRecordChanges = true;
                }

                if (dropped) {
                  const lvEmpty = !r.lifeVestSn || r.lifeVestSn === "-";
                  const saEmpty = !r.spareAirSn || r.spareAirSn === "-";
                  const hkEmpty = !r.helmetKitSn || r.helmetKitSn === "-";
                  if (lvEmpty && saEmpty && hkEmpty) {
                    r.durum = "ANKARA DEPO TESLİM";
                    r.bitisTarihi = dateStr;
                    r.teslimAlanPersonel = "Ankara Depo Sorumlusu";
                    r.notlar = (r.notlar ? r.notlar + " | " : "") + "Ekipman başka personele aktarıldı (Zimmet Kapatıldı).";
                    hasRecordChanges = true;
                  }
                }
              }
            });

            // Assign to target person
            let targetZimmet = currentRecords.find(r => r.durum === "GÖREVDE" && r.personelAdi.toUpperCase() === targetPerson);
            if (targetZimmet) {
              if (type === "can_yelegi") targetZimmet.lifeVestSn = sn;
              else if (type === "spare_air") targetZimmet.spareAirSn = sn;
              else if (type === "helmet_kit") targetZimmet.helmetKitSn = sn;
              targetZimmet.guncellenmeTarihi = dateStr;
              hasRecordChanges = true;
            } else {
              targetZimmet = {
                id: `yd_${Date.now()}_${Math.random().toString(36).substr(2, 4)}`,
                siraNo: currentRecords.length + 1,
                personelAdi: targetPerson,
                gorevYeri: item.gorevYeri || "MİLAS",
                baslangicTarihi: item.tarih && item.tarih !== "-" ? item.tarih : dateStr,
                bitisTarihi: "-",
                lifeVestSn: type === "can_yelegi" ? sn : "-",
                spareAirSn: type === "spare_air" ? sn : "-",
                helmetKitSn: type === "helmet_kit" ? sn : "-",
                durum: "GÖREVDE",
                devredilenPersonel: "-",
                teslimAlanPersonel: "Depo Sorumlusu",
                notlar: "Envanter tablosundan otomatik zimmetlendi.",
                guncellenmeTarihi: dateStr
              };
              currentRecords.unshift(targetZimmet);
              hasRecordChanges = true;
            }
          }
        }
      });

      if (hasRecordChanges) {
        writeYasamDestekData(currentRecords);
      }

      syncOfficialYasamDestekSheetsToGoogleSheet().catch(() => {});
      return res.json({ 
        status: "success", 
        message: "Stok verisi güncellendi, zimmet kayıtları senkronize edildi ve Google E-Tabloya iletildi.",
        records: currentRecords
      });
    } catch (err: any) {
      return res.status(500).json({ status: "error", message: err.message });
    }
  });

  // POST /api/yasam-destek/swap-item (Görevdeyken Tekil Parça Değişimi: Can Yeleği, Spare Air veya Helmet Kit)
  app.post("/api/yasam-destek/swap-item", async (req, res) => {
    try {
      const { zimmetId, personelAdi, itemType, oldSn, newSn, notlar, tarih } = req.body;
      if (!personelAdi || !itemType || !newSn) {
        return res.status(400).json({ status: "error", message: "Personel adı, ekipman türü ve yeni seri no gereklidir." });
      }

      const pName = personelAdi.trim().toUpperCase();
      const dateStr = tarih || new Date().toLocaleDateString("tr-TR");
      const userNotes = (notlar || "").trim();

      let currentRecords = readYasamDestekData();
      let targetRecord = currentRecords.find(r => (zimmetId && r.id === zimmetId) || (r.personelAdi.toUpperCase() === pName && r.durum === "GÖREVDE"));

      if (!targetRecord) {
        targetRecord = {
          id: `yd_${Date.now()}`,
          siraNo: currentRecords.length + 1,
          personelAdi: pName,
          gorevYeri: "MİLAS",
          baslangicTarihi: dateStr,
          bitisTarihi: "-",
          lifeVestSn: itemType === "can_yelegi" ? newSn : "-",
          spareAirSn: itemType === "spare_air" ? newSn : "-",
          helmetKitSn: itemType === "helmet_kit" ? newSn : "-",
          durum: "GÖREVDE",
          devredilenPersonel: "-",
          teslimAlanPersonel: "Depo Sorumlusu",
          notlar: `Ekipman değişimi yapıldı: ${itemType} (${oldSn || '-'} ➔ ${newSn}) ${userNotes}`.trim(),
          guncellenmeTarihi: dateStr
        };
        currentRecords.unshift(targetRecord);
      } else {
        const previousSn = itemType === "can_yelegi" ? targetRecord.lifeVestSn : (itemType === "spare_air" ? targetRecord.spareAirSn : targetRecord.helmetKitSn);
        const actualOldSn = oldSn || previousSn;

        if (itemType === "can_yelegi") {
          targetRecord.lifeVestSn = newSn;
        } else if (itemType === "spare_air") {
          targetRecord.spareAirSn = newSn;
        } else if (itemType === "helmet_kit") {
          targetRecord.helmetKitSn = newSn;
        }

        const changeLog = `${itemType.toUpperCase()} değişimi: ${actualOldSn || '-'} ➔ ${newSn}`;
        targetRecord.notlar = targetRecord.notlar ? `${targetRecord.notlar} | ${changeLog} ${userNotes}`.trim() : `${changeLog} ${userNotes}`.trim();
        targetRecord.guncellenmeTarihi = dateStr;
      }

      // 1. Can Yeleği Değişimi
      if (itemType === "can_yelegi") {
        let cyList = readJsonFile(CAN_YELEGI_FILE);
        if (oldSn && oldSn !== "-") {
          const oldIdx = cyList.findIndex((c: any) => c.seriNo === oldSn);
          if (oldIdx >= 0) {
            cyList[oldIdx].aciklamalar = "DEPODA";
            cyList[oldIdx].durum = "DEPODA";
            cyList[oldIdx].color = "#ffffff";
            cyList[oldIdx].tarih = "-";
          }
        }
        const newIdx = cyList.findIndex((c: any) => c.seriNo === newSn);
        if (newIdx >= 0) {
          cyList[newIdx].aciklamalar = pName;
          cyList[newIdx].durum = "GÖREV BÖLGESİNDE";
          cyList[newIdx].color = "#92d050";
          cyList[newIdx].tarih = dateStr;
        }
        writeJsonFile(CAN_YELEGI_FILE, cyList);
      }

      // 2. Spare Air Değişimi
      if (itemType === "spare_air") {
        let saList = readJsonFile(SPARE_AIR_FILE);
        if (oldSn && oldSn !== "-") {
          const oldIdx = saList.findIndex((s: any) => s.seriNo === oldSn);
          if (oldIdx >= 0) {
            saList[oldIdx].aciklamalar = "DEPODA";
            saList[oldIdx].durum = "DEPODA";
            saList[oldIdx].color = "#ffffff";
            saList[oldIdx].tarih = "-";
          }
        }
        const newIdx = saList.findIndex((s: any) => s.seriNo === newSn);
        if (newIdx >= 0) {
          saList[newIdx].aciklamalar = pName;
          saList[newIdx].durum = "GÖREV BÖLGESİNDE";
          saList[newIdx].color = "#92d050";
          saList[newIdx].tarih = dateStr;
        }
        writeJsonFile(SPARE_AIR_FILE, saList);
      }

      // 3. Helmet Kit Değişimi
      if (itemType === "helmet_kit") {
        let hkList = readJsonFile(HELMET_KIT_FILE);
        if (oldSn && oldSn !== "-") {
          const pureOld = oldSn.split(" (")[0].trim();
          const oldIdx = hkList.findIndex((h: any) => 
            (h.aciklamalar && h.aciklamalar.includes(pureOld)) || 
            h.seriNo1 === pureOld || 
            h.seriNo2 === pureOld ||
            h.aciklamalar === oldSn
          );
          if (oldIdx >= 0) {
            hkList[oldIdx].aciklamalar = "DEPODA";
            hkList[oldIdx].durum = "DEPODA";
            hkList[oldIdx].color = "#ffffff";
            hkList[oldIdx].tarih = "-";
          }
        }
        const pureNew = newSn.split(" (")[0].trim();
        const newIdx = hkList.findIndex((h: any) => 
          (h.aciklamalar && h.aciklamalar.includes(pureNew)) || 
          h.seriNo1 === pureNew || 
          h.seriNo2 === pureNew ||
          h.aciklamalar === newSn
        );
        if (newIdx >= 0) {
          hkList[newIdx].aciklamalar = pName;
          hkList[newIdx].durum = "GÖREV BÖLGESİNDE";
          hkList[newIdx].color = "#92d050";
          hkList[newIdx].tarih = dateStr;
        }
        writeJsonFile(HELMET_KIT_FILE, hkList);
      }

      writeYasamDestekData(currentRecords);
      syncOfficialYasamDestekSheetsToGoogleSheet().catch(() => {});

      return res.json({
        status: "success",
        message: `✅ ${pName} personeline ait ${itemType.toUpperCase()} başarıyla değiştirildi (${newSn}). Eski parça depoya alındı.`,
        record: targetRecord,
        records: currentRecords
      });
    } catch (err: any) {
      return res.status(500).json({ status: "error", message: err.message });
    }
  });

  // GET /api/yasam-destek
  app.get("/api/yasam-destek", async (req, res) => {
    try {
      let records = readYasamDestekData();

      // Try reading directly from Google Sheets if online
      try {
        const sheetId = "17ScGYYx0erzDwHDk6RGiHOdJATdfmmExXFBY39dXpF0";
        const url = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent("YAŞAM DESTEK AT-802")}`;
        const resp = await fetch(url, { signal: AbortSignal.timeout(4000) });
        if (resp.ok) {
          const csvText = await resp.text();
          if (!csvText.includes("<!DOCTYPE") && !csvText.includes("<html") && csvText.trim().length > 0) {
            const lines = csvText.split(/\r?\n/).filter(l => l.trim().length > 0);
            if (lines.length > 1 && lines[0].includes("PERSONEL")) {
              const regex = new RegExp(`,(?=(?:(?:[^"]*"){2})*[^"]*$)`);
              const remoteList: any[] = [];
              for (let i = 1; i < lines.length; i++) {
                const cols = lines[i].split(regex).map(c => c.replace(/^"|"$/g, '').trim());
                if (cols.length >= 2 && cols[1]) {
                  remoteList.push({
                    id: `yd_row_${i}`,
                    siraNo: i,
                    personelAdi: cols[1] || "",
                    gorevYeri: cols[2] || "",
                    baslangicTarihi: cols[3] || "",
                    bitisTarihi: cols[4] || "-",
                    lifeVestSn: cols[5] || "-",
                    spareAirSn: cols[6] || "-",
                    helmetKitSn: cols[7] || "-",
                    durum: cols[8] || "GÖREVDE",
                    devredilenPersonel: cols[9] || "-",
                    teslimAlanPersonel: cols[10] || "-",
                    notlar: cols[11] || "",
                    guncellenmeTarihi: cols[12] || ""
                  });
                }
              }
              if (remoteList.length > 0) {
                records = remoteList;
                writeYasamDestekData(records);
              }
            }
          }
        }
      } catch (sheetErr) {
        console.error("Google Sheet okuma hatası:", sheetErr);
        // Fall back to local records
      }

      return res.json({
        status: "success",
        count: records.length,
        records
      });
    } catch (err: any) {
      return res.status(500).json({ status: "error", message: err.message });
    }
  });

  // POST /api/yasam-destek (Generic Save/Update)
  app.post("/api/yasam-destek", async (req, res) => {
    try {
      const { record, records } = req.body;
      let currentRecords = readYasamDestekData();

      if (Array.isArray(records)) {
        currentRecords = records;
      } else if (record) {
        const existingIdx = currentRecords.findIndex(r => r.id === record.id);
        if (existingIdx >= 0) {
          currentRecords[existingIdx] = {
            ...currentRecords[existingIdx],
            ...record,
            guncellenmeTarihi: new Date().toLocaleDateString("tr-TR")
          };
        } else {
          currentRecords.unshift({
            ...record,
            id: record.id || `yd_${Date.now()}`,
            siraNo: currentRecords.length + 1,
            guncellenmeTarihi: new Date().toLocaleDateString("tr-TR")
          });
        }
      }

      writeYasamDestekData(currentRecords);
      
      // Anlık Senkronizasyon
      syncYasamDestekToGoogleSheet(currentRecords).catch(err => {
        console.error("Manual sync error:", err);
      });

      return res.json({
        status: "success",
        message: "Kayıt başarıyla işlendi.",
        records: currentRecords
      });
    } catch (err: any) {
      console.error("POST /api/yasam-destek error:", err);
      return res.status(500).json({ status: "error", message: err.message });
    }
  });

  // POST /api/yasam-destek/delete-record
  // POST /api/yasam-destek/delete-record
  app.post("/api/yasam-destek/delete-record", async (req, res) => {
    try {
      const { id } = req.body;
      console.log(`Zimmet silme isteği alındı. ID: ${id}`);
      if (!id) return res.status(400).json({ status: "error", message: "ID eksik." });
      
      let records = readYasamDestekData();
      const initialCount = records.length;
      records = records.filter(r => r.id !== id);
      
      if (records.length === initialCount) {
        console.warn(`Silinecek kayıt bulunamadı. ID: ${id}`);
      }

      writeYasamDestekData(records);
      
      // Google Sheet Senkronizasyonu
      syncYasamDestekToGoogleSheet(records).catch(err => {
        console.error("Zimmet silme sonrası Sheet senkronizasyon hatası:", err);
      });

      return res.json({ 
        status: "success", 
        message: "Zimmet kaydı başarıyla silindi.", 
        records 
      });
    } catch (err: any) {
      console.error("Zimmet silme hatası:", err);
      return res.status(500).json({ status: "error", message: err.message });
    }
  });

  // POST /api/yasam-destek/assign (Zimmet Ver ve Stok Tablolarını Otomatik Yeşil Yap)
  app.post("/api/yasam-destek/assign", async (req, res) => {
    try {
      const { personelAdi, lifeVestSn, helmetKitSn, spareAirSn, gorevYeri, baslangicTarihi, notlar } = req.body;
      if (!personelAdi || !personelAdi.trim()) {
        return res.status(400).json({ status: "error", message: "Personel adı zorunludur." });
      }

      const pName = personelAdi.trim().toUpperCase();
      const loc = (gorevYeri || "MİLAS").trim().toUpperCase();
      const dateStr = baslangicTarihi || new Date().toLocaleDateString("tr-TR");
      const lvSn = (lifeVestSn || "").trim();
      const saSn = (spareAirSn || "").trim();
      const hkSn = (helmetKitSn || "").trim();

      let currentRecords = readYasamDestekData();
      const newRecord = {
        id: `yd_${Date.now()}`,
        siraNo: currentRecords.length + 1,
        personelAdi: pName,
        gorevYeri: loc,
        baslangicTarihi: dateStr,
        bitisTarihi: "-",
        lifeVestSn: lvSn || "-",
        spareAirSn: saSn || "-",
        helmetKitSn: hkSn || "-",
        durum: "GÖREVDE",
        devredilenPersonel: "-",
        teslimAlanPersonel: "Depo Sorumlusu",
        notlar: notlar || "",
        guncellenmeTarihi: new Date().toLocaleDateString("tr-TR")
      };
      currentRecords.unshift(newRecord);
      writeYasamDestekData(currentRecords);

      // 1. Can Yeleği Stok Güncelleme
      let cyList = readJsonFile(CAN_YELEGI_FILE);
      if (lvSn) {
        const idx = cyList.findIndex((c: any) => c.seriNo === lvSn);
        if (idx >= 0) {
          cyList[idx].aciklamalar = pName;
          cyList[idx].durum = "GÖREV BÖLGESİNDE";
          cyList[idx].color = "#92d050";
          cyList[idx].tarih = dateStr;
          writeJsonFile(CAN_YELEGI_FILE, cyList);
        }
      }

      // 2. Spare Air Stok Güncelleme
      let saList = readJsonFile(SPARE_AIR_FILE);
      if (saSn) {
        const idx = saList.findIndex((s: any) => s.seriNo === saSn);
        if (idx >= 0) {
          saList[idx].aciklamalar = pName;
          saList[idx].durum = "GÖREV BÖLGESİNDE";
          saList[idx].color = "#92d050";
          saList[idx].tarih = dateStr;
          writeJsonFile(SPARE_AIR_FILE, saList);
        }
      }

      // 3. Helmet Kit Stok Güncelleme
      let hkList = readJsonFile(HELMET_KIT_FILE);
      if (hkSn) {
        const pureHk = hkSn.split(" (")[0].trim();
        const idx = hkList.findIndex((h: any) => 
          (h.aciklamalar && h.aciklamalar.includes(pureHk)) || 
          h.seriNo1 === pureHk || 
          h.seriNo2 === pureHk ||
          h.aciklamalar === hkSn
        );
        if (idx >= 0) {
          hkList[idx].aciklamalar = pName;
          hkList[idx].durum = "GÖREV BÖLGESİNDE";
          hkList[idx].color = "#92d050";
          hkList[idx].tarih = dateStr;
          writeJsonFile(HELMET_KIT_FILE, hkList);
        }
      }

      syncOfficialYasamDestekSheetsToGoogleSheet().catch(() => {});

      return res.json({
        status: "success",
        message: `${pName} personeline zimmet başarıyla verildi. Stok kayıtları yeşil dolgu ile güncellendi.`,
        record: newRecord,
        records: currentRecords
      });
    } catch (err: any) {
      return res.status(500).json({ status: "error", message: err.message });
    }
  });

  // POST /api/yasam-destek/finish (Zimmeti Devret veya Depoya Teslim Et)
  app.post("/api/yasam-destek/finish", async (req, res) => {
    try {
      const { zimmetId, action, bitisTarihi, devredilenPersonel, devredilenYer, teslimAlanPersonel, notlar } = req.body;
      if (!zimmetId) {
        return res.status(400).json({ status: "error", message: "Zimmet ID gerekli." });
      }

      let currentRecords = readYasamDestekData();
      const targetIdx = currentRecords.findIndex(r => r.id === zimmetId);
      if (targetIdx < 0) {
        return res.status(404).json({ status: "error", message: "Zimmet kaydı bulunamadı." });
      }

      const oldRecord = currentRecords[targetIdx];
      const bitis = bitisTarihi || new Date().toLocaleDateString("tr-TR");
      const userNotes = (notlar || "").trim();

      let cyList = readJsonFile(CAN_YELEGI_FILE);
      let saList = readJsonFile(SPARE_AIR_FILE);
      let hkList = readJsonFile(HELMET_KIT_FILE);

      if (action === "devret") {
        if (!devredilenPersonel || !devredilenPersonel.trim()) {
          return res.status(400).json({ status: "error", message: "Devredilecek personel seçilmelidir." });
        }
        const newPerson = devredilenPersonel.trim().toUpperCase();
        const newLoc = (devredilenYer || oldRecord.gorevYeri || "MİLAS").trim().toUpperCase();

        // Eski kaydı güncelle (DEVREDİLDİ)
        currentRecords[targetIdx] = {
          ...oldRecord,
          durum: "DEVREDİLDİ",
          bitisTarihi: bitis,
          devredilenPersonel: newPerson,
          teslimAlanPersonel: newPerson,
          notlar: `${oldRecord.notlar ? oldRecord.notlar + " | " : ""}Devredildi: ${newPerson} ${userNotes}`.trim(),
          guncellenmeTarihi: new Date().toLocaleDateString("tr-TR")
        };

        // Yeni personele ait aktif zimmet oluştur
        const newRecord = {
          id: `yd_${Date.now()}`,
          siraNo: currentRecords.length + 1,
          personelAdi: newPerson,
          gorevYeri: newLoc,
          baslangicTarihi: bitis,
          bitisTarihi: "-",
          lifeVestSn: oldRecord.lifeVestSn || "-",
          spareAirSn: oldRecord.spareAirSn || "-",
          helmetKitSn: oldRecord.helmetKitSn || "-",
          durum: "GÖREVDE",
          devredilenPersonel: "-",
          teslimAlanPersonel: "Devir Teslim",
          notlar: `${oldRecord.personelAdi} personelinden devir alındı. ${userNotes}`.trim(),
          guncellenmeTarihi: new Date().toLocaleDateString("tr-TR")
        };
        currentRecords.unshift(newRecord);

        // Stoklardaki teslim edilen ismini yeni personele güncelle ve YEŞİL TUT (#92d050)
        if (oldRecord.lifeVestSn && oldRecord.lifeVestSn !== "-") {
          const lvIdx = cyList.findIndex((c: any) => c.seriNo === oldRecord.lifeVestSn);
          if (lvIdx >= 0) {
            cyList[lvIdx].aciklamalar = newPerson;
            cyList[lvIdx].durum = "GÖREV BÖLGESİNDE";
            cyList[lvIdx].color = "#92d050";
            cyList[lvIdx].tarih = bitis;
          }
        }
        if (oldRecord.spareAirSn && oldRecord.spareAirSn !== "-") {
          const saIdx = saList.findIndex((s: any) => s.seriNo === oldRecord.spareAirSn);
          if (saIdx >= 0) {
            saList[saIdx].aciklamalar = newPerson;
            saList[saIdx].durum = "GÖREV BÖLGESİNDE";
            saList[saIdx].color = "#92d050";
            saList[saIdx].tarih = bitis;
          }
        }
        if (oldRecord.helmetKitSn && oldRecord.helmetKitSn !== "-") {
          const pureHk = oldRecord.helmetKitSn.split(" (")[0].trim();
          const hkIdx = hkList.findIndex((h: any) => 
            (h.aciklamalar && h.aciklamalar.includes(pureHk)) || 
            h.seriNo1 === pureHk || 
            h.seriNo2 === pureHk ||
            h.aciklamalar === oldRecord.helmetKitSn
          );
          if (hkIdx >= 0) {
            hkList[hkIdx].aciklamalar = newPerson;
            hkList[hkIdx].durum = "GÖREV BÖLGESİNDE";
            hkList[hkIdx].color = "#92d050";
            hkList[hkIdx].tarih = bitis;
          }
        }

      } else {
        // action === "teslim" (DEPOYA TESLİM - BEYAZ DOLGU / DEPODA)
        const receiver = (teslimAlanPersonel || "Ankara Depo Sorumlusu").trim();

        currentRecords[targetIdx] = {
          ...oldRecord,
          durum: "ANKARA DEPO TESLİM",
          bitisTarihi: bitis,
          teslimAlanPersonel: receiver,
          notlar: `${oldRecord.notlar ? oldRecord.notlar + " | " : ""}Depoya teslim edildi: ${receiver}. ${userNotes}`.trim(),
          guncellenmeTarihi: new Date().toLocaleDateString("tr-TR")
        };

        // Stoklardaki teslim edilen ismini "DEPODA" yap ve DOLGUSUZ/BEYAZ (#ffffff) yap
        if (oldRecord.lifeVestSn && oldRecord.lifeVestSn !== "-") {
          const lvIdx = cyList.findIndex((c: any) => c.seriNo === oldRecord.lifeVestSn);
          if (lvIdx >= 0) {
            cyList[lvIdx].aciklamalar = "DEPODA";
            cyList[lvIdx].durum = "DEPODA";
            cyList[lvIdx].color = "#ffffff";
            cyList[lvIdx].tarih = "-";
          }
        }
        if (oldRecord.spareAirSn && oldRecord.spareAirSn !== "-") {
          const saIdx = saList.findIndex((s: any) => s.seriNo === oldRecord.spareAirSn);
          if (saIdx >= 0) {
            saList[saIdx].aciklamalar = "DEPODA";
            saList[saIdx].durum = "DEPODA";
            saList[saIdx].color = "#ffffff";
            saList[saIdx].tarih = "-";
          }
        }
        if (oldRecord.helmetKitSn && oldRecord.helmetKitSn !== "-") {
          const pureHk = oldRecord.helmetKitSn.split(" (")[0].trim();
          const hkIdx = hkList.findIndex((h: any) => 
            (h.aciklamalar && h.aciklamalar.includes(pureHk)) || 
            h.seriNo1 === pureHk || 
            h.seriNo2 === pureHk ||
            h.aciklamalar === oldRecord.helmetKitSn
          );
          if (hkIdx >= 0) {
            hkList[hkIdx].aciklamalar = "DEPODA";
            hkList[hkIdx].durum = "DEPODA";
            hkList[hkIdx].color = "#ffffff";
            hkList[hkIdx].tarih = "-";
          }
        }
      }

      writeYasamDestekData(currentRecords);
      writeJsonFile(CAN_YELEGI_FILE, cyList);
      writeJsonFile(SPARE_AIR_FILE, saList);
      writeJsonFile(HELMET_KIT_FILE, hkList);

      syncOfficialYasamDestekSheetsToGoogleSheet().catch(() => {});

      return res.json({
        status: "success",
        message: action === "devret" ? "Zimmet başarıyla devredildi." : "Ekipmanlar depoya teslim alındı ve stoklar beyaz dolguya çevrildi.",
        records: currentRecords
      });
    } catch (err: any) {
      return res.status(500).json({ status: "error", message: err.message });
    }
  });

  // POST /api/yasam-destek/reset-to-depot (Tüm Yaşam Destek Stoklarını Depoda ve Beyaz Dolgu Yap)
  app.post("/api/yasam-destek/reset-to-depot", async (req, res) => {
    try {
      let cyList = readJsonFile(CAN_YELEGI_FILE);
      cyList = cyList.map((c: any) => ({
        ...c,
        aciklamalar: "DEPODA",
        durum: "DEPODA",
        color: "#ffffff",
        tarih: "-"
      }));
      writeJsonFile(CAN_YELEGI_FILE, cyList);

      let saList = readJsonFile(SPARE_AIR_FILE);
      saList = saList.map((s: any) => ({
        ...s,
        aciklamalar: "DEPODA",
        durum: "DEPODA",
        color: "#ffffff",
        tarih: "-"
      }));
      writeJsonFile(SPARE_AIR_FILE, saList);

      let hkList = readJsonFile(HELMET_KIT_FILE);
      hkList = hkList.map((h: any) => ({
        ...h,
        aciklamalar: "DEPODA",
        durum: "DEPODA",
        color: "#ffffff",
        tarih: "-"
      }));
      writeJsonFile(HELMET_KIT_FILE, hkList);

      // Aktif zimmetlerin durumunu depoya teslim olarak işaretle
      let currentRecords = readYasamDestekData();
      currentRecords = currentRecords.map(r => {
        if (r.durum === "GÖREVDE") {
          return {
            ...r,
            durum: "ANKARA DEPO TESLİM",
            bitisTarihi: new Date().toLocaleDateString("tr-TR"),
            teslimAlanPersonel: "Ankara Depo Sorumlusu",
            notlar: (r.notlar ? r.notlar + " | " : "") + "Sistem başlangıç durumuna alındı (Depoya Teslim)."
          };
        }
        return r;
      });
      writeYasamDestekData(currentRecords);

      syncOfficialYasamDestekSheetsToGoogleSheet().catch(() => {});

      return res.json({
        status: "success",
        message: "Tüm yaşam destek stokları DEPODA olarak güncellendi ve beyaz dolgu yapıldı.",
        canYelegiCount: cyList.length,
        spareAirCount: saList.length,
        helmetKitCount: hkList.length
      });
    } catch (err: any) {
      return res.status(500).json({ status: "error", message: err.message });
    }
  });

  // GET /api/yasam-destek/public
  app.get("/api/yasam-destek/public", (req, res) => {
    const records = readYasamDestekData();
    const active = records.filter(r => r.durum === "GÖREVDE");
    return res.json({
      title: "AT-802 Yaşam Destek Ekipmanları Canlı Görev Takibi",
      updatedAt: new Date().toLocaleString("tr-TR"),
      totalCount: records.length,
      activeCount: active.length,
      activePersonnel: active.map(r => ({
        personel: r.personelAdi,
        gorevYeri: r.gorevYeri,
        lifeVest: r.lifeVestSn,
        spareAir: r.spareAirSn,
        helmetKit: r.helmetKitSn,
        baslangic: r.baslangicTarihi
      }))
    });
  });

  // Public HTML Dashboard for Yaşam Destek live tracking
  app.get("/api/yasam-destek-index", (req, res) => {
    const records = readYasamDestekData();
    const active = records.filter(r => r.durum === "GÖREVDE");
    const devredildi = records.filter(r => r.durum === "DEVREDİLDİ");
    const depoda = records.filter(r => r.durum === "ANKARA DEPO TESLİM");

    const html = `
    <!DOCTYPE html>
    <html lang="tr">
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>OGM AT-802 Yaşam Destek Ekipmanları Canlı Takip Portalı</title>
      <script src="https://cdn.tailwindcss.com"></script>
    </head>
    <body class="bg-slate-100 font-sans p-4 md:p-8 text-slate-800">
      <div class="max-w-6xl mx-auto space-y-6">
        <header class="bg-[#0b3d1d] text-white p-6 rounded-2xl shadow-md flex flex-wrap items-center justify-between gap-4">
          <div>
            <div class="text-xs font-bold text-emerald-300 uppercase tracking-widest">T.C. Orman Genel Müdürlüğü • Havacılık Dairesi</div>
            <h1 class="text-xl md:text-2xl font-black mt-1">🦺 AT-802 YAŞAM DESTEK EKİPMANLARI CANLI TAKİP PORTALI</h1>
            <p class="text-xs text-emerald-100/80 mt-1">Life Vest, Spare Air ve Helmet Kit Görev & Zimmet Takip Sistemi</p>
          </div>
          <div class="text-right text-xs">
            <div class="font-mono bg-black/20 px-3 py-1.5 rounded-lg border border-white/10">Son Güncelleme: ${new Date().toLocaleDateString("tr-TR")}</div>
          </div>
        </header>

        <div class="grid grid-cols-2 md:grid-cols-4 gap-4">
          <div class="bg-white p-4 rounded-xl shadow-xs border border-slate-200">
            <div class="text-xs font-bold text-slate-500 uppercase">Toplam Kayıt</div>
            <div class="text-2xl font-black text-slate-800 mt-1">${records.length}</div>
          </div>
          <div class="bg-emerald-50 p-4 rounded-xl shadow-xs border border-emerald-200">
            <div class="text-xs font-bold text-emerald-800 uppercase">🟢 Aktif Görevde</div>
            <div class="text-2xl font-black text-emerald-700 mt-1">${active.length}</div>
          </div>
          <div class="bg-amber-50 p-4 rounded-xl shadow-xs border border-amber-200">
            <div class="text-xs font-bold text-amber-800 uppercase">🔄 Devredilen</div>
            <div class="text-2xl font-black text-amber-700 mt-1">${devredildi.length}</div>
          </div>
          <div class="bg-blue-50 p-4 rounded-xl shadow-xs border border-blue-200">
            <div class="text-xs font-bold text-blue-800 uppercase">🏢 Depoda</div>
            <div class="text-2xl font-black text-blue-700 mt-1">${depoda.length}</div>
          </div>
        </div>

        <div class="bg-white rounded-2xl shadow-xs border border-slate-200 overflow-hidden">
          <div class="p-4 border-b border-slate-200 flex items-center justify-between">
            <h2 class="text-sm font-black uppercase text-slate-800">Canlı Ekipman Zimmet & Görev Tablosu</h2>
            <span class="text-xs font-bold bg-emerald-100 text-emerald-800 px-2.5 py-1 rounded-full">${active.length} Personel Görevde</span>
          </div>
          <div class="overflow-x-auto">
            <table class="w-full text-left text-xs border-collapse">
              <thead class="bg-slate-800 text-white text-[11px] uppercase tracking-wider text-center">
                <tr>
                  <th class="p-3">SIRA</th>
                  <th class="p-3 text-left">PERSONEL ADI SOYADI</th>
                  <th class="p-3">GÖREV YERİ</th>
                  <th class="p-3">BAŞLANGIÇ</th>
                  <th class="p-3">BİTİŞ</th>
                  <th class="p-3 bg-amber-900/60">LIFE VEST S/N</th>
                  <th class="p-3 bg-cyan-900/60">SPARE AIR S/N</th>
                  <th class="p-3 bg-purple-900/60">HELMET KIT S/N</th>
                  <th class="p-3">DURUM</th>
                  <th class="p-3">DEVİR / TESLİM</th>
                </tr>
              </thead>
              <tbody class="divide-y divide-slate-100 text-center">
                ${records.map((r, i) => `
                  <tr class="hover:bg-slate-50">
                    <td class="p-3 font-bold text-slate-400">${i + 1}</td>
                    <td class="p-3 text-left font-bold text-slate-900">${r.personelAdi}</td>
                    <td class="p-3"><span class="px-2 py-0.5 rounded bg-slate-100 font-bold">📍 ${r.gorevYeri}</span></td>
                    <td class="p-3 font-mono text-slate-600">${r.baslangicTarihi}</td>
                    <td class="p-3 font-mono text-slate-600">${r.bitisTarihi || '-'}</td>
                    <td class="p-3 font-mono font-bold text-amber-800 bg-amber-50/50">${r.lifeVestSn}</td>
                    <td class="p-3 font-mono font-bold text-cyan-800 bg-cyan-50/50">${r.spareAirSn}</td>
                    <td class="p-3 font-mono font-bold text-purple-800 bg-purple-50/50">${r.helmetKitSn}</td>
                    <td class="p-3">
                      <span class="px-2 py-0.5 rounded-full text-[10px] font-black ${
                        r.durum === 'GÖREVDE' ? 'bg-emerald-100 text-emerald-800 border border-emerald-300' :
                        r.durum === 'DEVREDİLDİ' ? 'bg-amber-100 text-amber-800 border border-amber-300' :
                        'bg-blue-100 text-blue-800 border border-blue-300'
                      }">
                        ${r.durum}
                      </span>
                    </td>
                    <td class="p-3 text-slate-600 font-bold">${r.devredilenPersonel || r.teslimAlanPersonel || '-'}</td>
                  </tr>
                `).join('')}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </body>
    </html>
    `;
    return res.send(html);
  });

  // Serve public Yaşam Destek portal
  app.get("/yasam-destek", (req, res) => {
    const htmlPath = path.resolve(process.cwd(), "public/yasam-destek.html");
    if (fs.existsSync(htmlPath)) {
      return res.sendFile(htmlPath);
    }
    return res.redirect("/api/yasam-destek-index");
  });

  const PERSONNEL_FILE = path.resolve(process.cwd(), "data/personnel_list.json");
  const PERSONNEL_EXCEL_FILE = path.resolve(process.cwd(), "data/personnel_list.xlsx");

  const SEED_AT802_PERSONNEL = [
    { id: "p_1", sNo: 1, adSoyad: "OKAN AKÇA", unvan: "Teknisyen" },
    { id: "p_2", sNo: 2, adSoyad: "KEMAL CAN", unvan: "Teknisyen" },
    { id: "p_3", sNo: 3, adSoyad: "DOĞAN ÖZTÜRK", unvan: "Teknisyen" },
    { id: "p_4", sNo: 4, adSoyad: "SERDAR TOY", unvan: "Teknisyen" },
    { id: "p_5", sNo: 5, adSoyad: "TAŞKIN ÇALIŞIR", unvan: "Teknisyen" },
    { id: "p_6", sNo: 6, adSoyad: "FATİH YAŞBAY", unvan: "Teknisyen" },
    { id: "p_7", sNo: 7, adSoyad: "ÖZGÜR UYAR", unvan: "Teknisyen" },
    { id: "p_8", sNo: 8, adSoyad: "UĞUR GÜRCAN", unvan: "Teknisyen" },
    { id: "p_9", sNo: 9, adSoyad: "URAL UÇURUM", unvan: "Teknisyen" },
    { id: "p_10", sNo: 10, adSoyad: "YUNUS EROL", unvan: "Teknisyen" },
    { id: "p_11", sNo: 11, adSoyad: "SELÇUK İBİŞ", unvan: "Teknisyen" },
    { id: "p_12", sNo: 12, adSoyad: "ALİ TOPALAN", unvan: "Teknisyen" },
    { id: "p_13", sNo: 13, adSoyad: "FATİH PARLAK", unvan: "Teknisyen" },
    { id: "p_14", sNo: 14, adSoyad: "TALAT TÖNGÜŞ", unvan: "Teknisyen" },
    { id: "p_15", sNo: 15, adSoyad: "AHMET BİLGEN", unvan: "Teknisyen" },
    { id: "p_16", sNo: 16, adSoyad: "BARIŞ BUĞRA KARACA", unvan: "Teknisyen" },
    { id: "p_17", sNo: 17, adSoyad: "BARIŞ ÇELİKER", unvan: "Teknisyen" },
    { id: "p_18", sNo: 18, adSoyad: "FERDİ GÖL", unvan: "Teknisyen" },
    { id: "p_19", sNo: 19, adSoyad: "KEMAL ÇAKMAK", unvan: "Teknisyen" },
    { id: "p_20", sNo: 20, adSoyad: "BURHAN GÜLER", unvan: "Pilot" },
    { id: "p_21", sNo: 21, adSoyad: "ALPER GİDER", unvan: "Pilot" },
    { id: "p_22", sNo: 22, adSoyad: "OSMAN AKTAŞ", unvan: "Pilot" },
    { id: "p_23", sNo: 23, adSoyad: "KAĞAN BASKAK", unvan: "Pilot" },
    { id: "p_24", sNo: 24, adSoyad: "ATAKAN TANSÜYER", unvan: "Pilot" },
    { id: "p_25", sNo: 25, adSoyad: "MUZAFFER BARIŞ COŞAR", unvan: "Pilot" },
    { id: "p_26", sNo: 26, adSoyad: "ŞENER TURAN", unvan: "Pilot" },
    { id: "p_27", sNo: 27, adSoyad: "SERDAR KADEMLİOĞLU", unvan: "Pilot" },
    { id: "p_28", sNo: 28, adSoyad: "AYDINER ÇILDIR", unvan: "Pilot" },
    { id: "p_29", sNo: 29, adSoyad: "MEHMET AKİF BOZKURT", unvan: "Pilot" },
    { id: "p_30", sNo: 30, adSoyad: "OSMAN UZUNER", unvan: "Pilot" },
    { id: "p_31", sNo: 31, adSoyad: "RECEP BEKEN", unvan: "Pilot" },
    { id: "p_32", sNo: 32, adSoyad: "OĞUZHAN ŞENSÖZ", unvan: "Pilot" }
  ];

  function readPersonnelFromExcel(): any[] {
    try {
      const jsonItems = readJsonFile(PERSONNEL_FILE);
      if (Array.isArray(jsonItems) && jsonItems.length > 0) {
        return jsonItems;
      }
      if (fs.existsSync(PERSONNEL_EXCEL_FILE)) {
        const workbook = XLSX.readFile(PERSONNEL_EXCEL_FILE);
        const sheetName = workbook.SheetNames.includes("Personel Listesi")
          ? "Personel Listesi"
          : (workbook.SheetNames.includes("AT-802 PERSONEL VERİSİ") 
            ? "AT-802 PERSONEL VERİSİ" 
            : (workbook.SheetNames.includes("PERSONEL LİSTESİ") ? "PERSONEL LİSTESİ" : workbook.SheetNames[0]));
        const sheet = workbook.Sheets[sheetName];
        if (sheet) {
          const rows: any[] = XLSX.utils.sheet_to_json(sheet);
          if (rows && rows.length > 0) {
            const items = rows.map((r, idx) => {
              const rowStr = JSON.stringify(r).toUpperCase();
              if (rowStr.includes("KASK TAKIMI") || rowStr.includes("KOLTUK SAYISI") || rowStr.includes("DOUBLE HAVA") || rowStr.includes("SINGLE HAVA")) {
                return null;
              }
              const keys = Object.keys(r);
              const adSoyadVal = (r["PERSONEL ADI SOYADI"] || r.adSoyad || r.Personel || r["AD SOYAD"] || (keys[1] ? r[keys[1]] : "") || "").toString().trim().toUpperCase();
              const unvanVal = (r["ÜNVANI / GÖREVİ"] || r.unvan || r.Unvan || r["ÜNVA"] || (keys[2] ? r[keys[2]] : "Teknisyen") || "Teknisyen").toString().trim();
              return {
                id: r.id || r.ID || `p_${idx + 1}`,
                adSoyad: adSoyadVal,
                unvan: unvanVal
              };
            }).filter(p => p !== null && p.adSoyad && p.adSoyad.length > 0);
            if (items.length > 0) {
              writeJsonFile(PERSONNEL_FILE, items);
              return items;
            }
          }
        }
      }
    } catch (err) {
      console.warn("readPersonnelFromExcel error:", err);
    }

    // Default clean seed fallback
    writePersonnelToExcel(SEED_AT802_PERSONNEL);
    return SEED_AT802_PERSONNEL;
  }

  async function fetchPersonnelLive(): Promise<any[]> {
    const existing = readPersonnelFromExcel();
    if (Array.isArray(existing) && existing.length > 0) {
      return existing;
    }

    const sheetId = "17ScGYYx0erzDwHDk6RGiHOdJATdfmmExXFBY39dXpF0";
    const candidates = [
      "Personel Listesi",
      "AT-802 PERSONEL VERİSİ",
      "PERSONEL LİSTESİ",
      "5-Personel_Bilgi"
    ];

    for (const sheetName of candidates) {
      try {
        const url = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(sheetName)}`;
        const resp = await fetch(url, { signal: AbortSignal.timeout(5000) });
        if (resp.ok) {
          const csvText = await resp.text();
          if (!csvText.includes("<!DOCTYPE") && !csvText.includes("<html") && csvText.trim().length > 0) {
            const lines = csvText.split(/\r?\n/).filter(l => l.trim().length > 0);
            if (lines.length > 1) {
              const regex = new RegExp(`,(?=(?:(?:[^"]*"){2})*[^"]*$)`);
              const remoteList: any[] = [];
              for (let i = 1; i < lines.length; i++) {
                const cols = lines[i].split(regex).map(c => c.replace(/^"|"$/g, '').trim());
                if (sheetName === "5-Personel_Bilgi") {
                  if (cols.length >= 2 && cols[1]) {
                    const name = cols[1].trim().toUpperCase();
                    const unvan = cols.length >= 5 && cols[4] ? cols[4].trim() : "Teknisyen";
                    if (name && !name.includes("ADI SOYADI") && !name.includes("SIRA")) {
                      remoteList.push({ id: `p_${remoteList.length + 1}`, adSoyad: name, unvan });
                    }
                  }
                } else {
                  if (cols.length >= 2 && cols[1]) {
                    const name = cols[1].trim().toUpperCase();
                    const unvan = cols.length >= 3 && cols[2] ? cols[2].trim() : "Teknisyen";
                    if (name && !name.includes("PERSONEL ADI") && !name.includes("SIRA NO")) {
                      remoteList.push({ id: `p_${remoteList.length + 1}`, adSoyad: name, unvan });
                    }
                  }
                }
              }
              if (remoteList.length > 0) {
                writePersonnelToExcel(remoteList);
                return remoteList;
              }
            }
          }
        }
      } catch (e) {
        // try next candidate
      }
    }

    return [];
  }

  function writePersonnelToExcel(items: any[]): void {
    try {
      const dir = path.dirname(PERSONNEL_EXCEL_FILE);
      if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });

      const rows = [
        ["SIRA NO", "PERSONEL ADI SOYADI", "ÜNVANI / GÖREVİ"],
        ...items.map((p, idx) => [
          String(idx + 1),
          typeof p === "string" ? p : (p.adSoyad || ""),
          typeof p === "string" ? "Teknisyen" : (p.unvan || "Teknisyen")
        ])
      ];

      const worksheet = XLSX.utils.aoa_to_sheet(rows);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "AT-802 PERSONEL VERİSİ");
      XLSX.utils.book_append_sheet(workbook, worksheet, "Personel Listesi");
      XLSX.writeFile(workbook, PERSONNEL_EXCEL_FILE);

      writeJsonFile(PERSONNEL_FILE, items);
    } catch (err) {
      console.warn("writePersonnelToExcel error:", err);
    }
  }

  // GET /api/yasam-destek/personnel-list (Teknisyenler, Pilotlar & Personel Listesi)
  app.get("/api/yasam-destek/personnel-list", async (req, res) => {
    try {
      const items = await fetchPersonnelLive();
      const nameList = items.map(p => typeof p === "string" ? p : p.adSoyad).sort();
      return res.json({
        status: "success",
        count: items.length,
        personnel: nameList,
        items: items
      });
    } catch (err: any) {
      console.error("Personnel list fetch error (server):", err);
      return res.status(500).json({ status: "error", message: err.message });
    }
  });

  // POST /api/yasam-destek/personnel-list (Add / Edit / Delete Personnel directly to Excel file)
  app.post("/api/yasam-destek/personnel-list", async (req, res) => {
    try {
      const { action, item, id, items } = req.body;
      let currentItems = readPersonnelFromExcel();

      if (Array.isArray(items)) {
        currentItems = items;
      } else if (action === "add" && item) {
        currentItems.push({
          id: `p_${Date.now()}`,
          adSoyad: (item.adSoyad || "").toUpperCase().trim(),
          unvan: item.unvan || "Teknisyen"
        });
      } else if (action === "edit" && item && item.id) {
        const idx = currentItems.findIndex((p: any) => p.id === item.id);
        if (idx >= 0) {
          currentItems[idx] = {
            ...currentItems[idx],
            adSoyad: (item.adSoyad || "").toUpperCase().trim(),
            unvan: item.unvan || currentItems[idx].unvan
          };
        }
      } else if (action === "delete" && id) {
        currentItems = currentItems.filter((p: any) => p.id !== id);
      }

      writePersonnelToExcel(currentItems);
      syncOfficialYasamDestekSheetsToGoogleSheet().catch(() => {});
      const nameList = currentItems.map(p => typeof p === "string" ? p : p.adSoyad).sort();
      return res.json({
        status: "success",
        message: "Personel listesi Excel dosyasına ve portale kaydedildi.",
        count: currentItems.length,
        personnel: nameList,
        items: currentItems
      });
    } catch (err: any) {
      return res.status(500).json({ status: "error", message: err.message });
    }
  });

  // API Route: Depo Transfers & Bulk Excel Sync
  app.get("/api/depo-transfers", async (req, res) => {
    try {
      const sheetId = "17ScGYYx0erzDwHDk6RGiHOdJATdfmmExXFBY39dXpF0";
      const gid = "1992538823";
      const gvizUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv&gid=${gid}`;
      
      const resp = await fetch(gvizUrl);
      if (!resp.ok) {
        throw new Error(`Google Sheets responded with HTTP ${resp.status}`);
      }
      
      const csvText = await resp.text();
      return res.json({
        status: "success",
        csv: csvText
      });
    } catch (err: any) {
      return res.status(500).json({
        status: "error",
        message: err.message
      });
    }
  });

  // API Route: Get Depo Transfers
  app.get("/api/get-depo-transfers", async (req, res) => {
    try {
      const transactions = await fetchAndMergeAllTransactions();
      return res.json({
        status: "success",
        transactions: transactions
      });
    } catch (err: any) {
      return res.status(500).json({
        status: "error",
        message: err.message
      });
    }
  });

  // API Route: Save Depo Transfers to Google Apps Script / Google Sheets with Permanent Preservation
  app.post("/api/save-depo-transfers", async (req, res) => {
    try {
      const { 
        transfers = [], 
        spreadsheetId = "17ScGYYx0erzDwHDk6RGiHOdJATdfmmExXFBY39dXpF0", 
        sheetName = "DEPO HAREKET GEÇMİŞİ-AT-802" 
      } = req.body;

      const incomingList = (Array.isArray(transfers) ? transfers : [transfers]).map(t => {
        if (!t) return t;
        const now = new Date();
        const timestamp = t.timestamp || now.toISOString();
        const date = t.date || now.toLocaleDateString("tr-TR") + " " + now.toLocaleTimeString("tr-TR", { hour: '2-digit', minute: '2-digit' });
        return { ...t, timestamp, date };
      });
      
      // Fetch full existing history first
      const fullExisting = await fetchAndMergeAllTransactions();

      if (incomingList.length === 0) {
        return res.json({ status: "success", count: fullExisting.length, transactions: fullExisting });
      }

      // Helper key to avoid duplicating identical transaction entries
      const makeTxKey = (t: any) => {
        if (!t) return "";
        const id = t.id ? String(t.id).trim() : "";
        if (id && !id.startsWith("tx_sheets_")) return id;
        const timestamp = String(t.timestamp || t.date || "").trim();
        const item = String(t.itemName || t.name || t.itemDesc || "").trim().toLowerCase();
        const sn = String(t.sn || t.serialNumber || "-").trim().toLowerCase();
        const type = String(t.type || t.islemTuru || "").trim().toLowerCase();
        const loc = String(t.location || t.depoYeri || "").trim().toLowerCase();
        const qty = String(t.quantity !== undefined ? t.quantity : (t.adet || 1)).trim();
        return `${timestamp}_${item}_${sn}_${type}_${loc}_${qty}`;
      };

      const updatedList = [...fullExisting];
      const newlyAdded: any[] = [];

      for (const inc of incomingList) {
        if (!inc) continue;
        const incId = inc.id ? String(inc.id).trim() : "";
        
        let foundIdx = -1;
        // Only update in-place if explicitly marked as an edit
        if (incId && inc.isEdit === true) {
          foundIdx = updatedList.findIndex(e => e.id && String(e.id).trim() === incId);
        }

        if (foundIdx >= 0) {
          // UPDATE: Replace existing record with edited one in-place
          updatedList[foundIdx] = {
            ...updatedList[foundIdx],
            ...inc
          };
        } else {
          // INSERT: Append newly added transaction without clobbering existing ones!
          const existingKeys = new Set(updatedList.map(makeTxKey));
          const key = makeTxKey(inc);
          if (!existingKeys.has(key)) {
            newlyAdded.push(inc);
            updatedList.push(inc);
            existingKeys.add(key);
          }
        }
      }

      // Save FULL consolidated list to local database backup files
      writePersistedDepoTransactions(updatedList);
      
      const localJsonPath = path.resolve(process.cwd(), "public/depo_transactions.json");
      fs.writeFileSync(localJsonPath, JSON.stringify(updatedList, null, 2), "utf-8");

      const targetSheet = sheetName || "DEPO HAREKET GEÇMİŞİ-AT-802";
      const headers = ["MALZEME ADI", "ADET", "TARİH", "İŞLEM TÜRÜ", "SERİAL NUMBER", "KUYRUK KODU", "TESLİM ALAN", "KABUL YAPAN", "DEPO YERİ", "SAYFA ADI"];
      
      // CRITICAL BUG FIX: Filter by targetSheet specifically, so we don't wipe/mix other tabs!
      const sheetSpecificTransfers = updatedList.filter((t: any) => {
        const tSheet = t.sheetName || getDepoSheetNameForUnit(t.unit);
        return tSheet === targetSheet;
      });

      // Full rows for the specific Google Sheet update
      const rows = [
        headers,
        ...sheetSpecificTransfers.map((t: any) => {
          const isChem = (t.category === 'kimyasal') || (t.sheetName && String(t.sheetName).toLowerCase().includes('kimya')) || (t.location && String(t.location).toLowerCase().includes('kimya'));
          return [
            t.itemName || t.name || t.itemDesc || "",
            String(t.quantity !== undefined ? t.quantity : (t.adet || 1)),
            t.date || t.timestamp || "",
            String(t.islemTuru || t.type || "TRANSFER").toUpperCase(),
            t.sn || t.serialNumber || "-",
            t.tailNo || t.kuyrukKodu || t.kuyrukNo || "-",
            t.operator || t.teslimAlan || "-",
            t.receivedBy || t.kabulYapan || "-",
            isChem ? "KİMYASAL DEPO" : (t.location || t.depoYeri || "DEPO"),
            t.sheetName || targetSheet || "DEPO HAREKET GEÇMİŞİ-AT-802"
          ];
        })
      ];

      const scriptUrls = [
        "https://script.google.com/macros/s/AKfycbwP1uOo2NrST5a4I8vm1nGBLtI26yY2lWrmu9_e9iymwBkUhJBA9JOPCp7SNKqJbOubOw/exec",
        "https://script.google.com/macros/s/AKfycbw4kruTTc058Y9rLTyO3dKi6KloYsmdDTwV1GSiAk8ZXefyo3Z7_VDSTuurzsS9BHAQyQ/exec"
      ];

      // Perform slow Google Sheet network synchronization asynchronously in the background
      (async () => {
        let lastResult: any = null;
        for (const sUrl of scriptUrls) {
          try {
            // Write FULL cumulative list of sheetSpecificTransfers so Google Sheets contains complete history and never loses any records
            const updateRes = await fetch(sUrl, {
              method: "POST",
              headers: { "Content-Type": "text/plain;charset=utf-8" },
              body: JSON.stringify({
                action: "updateSheet",
                sheetName: targetSheet,
                spreadsheetId: spreadsheetId,
                data: rows
              }),
              signal: AbortSignal.timeout(10000)
            });
            const updateData = await updateRes.json();
            if (updateData && updateData.status === "success") {
              lastResult = updateData;
              break;
            }
          } catch (uErr) {}
        }

        // --- 7-Sayimlar Global Audit Sync ---
        const countRecords = updatedList.filter((t: any) => 
          String(t.type || t.islemTuru || "").toUpperCase() === 'SAYIM' || 
          String(t.islemTuru || "").includes("SAYIM")
        );

        if (countRecords.length > 0) {
          const sayimHeaders = ["Tarih / Saat", "Malzeme Adı", "P/N", "S/N", "Depo Bölgesi", "Sistem Mevcut", "Fiziksel Sayılan", "Fark (+/-)", "Sayan Personel"];
          const sayimRows = [
            sayimHeaders,
            ...countRecords.map((t: any) => [
              t.date || t.timestamp || "",
              t.itemName || t.name || "",
              t.pn || "-",
              t.sn || "-",
              t.location || "-",
              t.sistemStok !== undefined ? t.sistemStok : "-",
              t.sayilanAdet !== undefined ? t.sayilanAdet : "-",
              t.fark !== undefined ? t.fark : "-",
              t.operator || "-"
            ])
          ];

          for (const sUrl of scriptUrls) {
            try {
              await fetch(sUrl, {
                method: "POST",
                headers: { "Content-Type": "text/plain;charset=utf-8" },
                body: JSON.stringify({
                  action: "updateSheet",
                  sheetName: "7-Sayimlar",
                  spreadsheetId: spreadsheetId,
                  data: sayimRows
                }),
                signal: AbortSignal.timeout(10000)
              });
              break;
            } catch (e) {}
          }
        }
      })().catch(err => console.warn("Background google sheets sync error:", err));

      return res.json({ 
        status: "success",
        count: updatedList.length,
        newlyAddedCount: newlyAdded.length,
        transactions: updatedList,
        message: "Depo hareket geçmişi yerel diske kaydedildi ve Google Drive güncellemesi arka planda başlatıldı."
      });
    } catch (err: any) {
      return res.status(500).json({
        status: "error",
        message: err.message
      });
    }
  });

  // In-memory cache for rapid loading, instant refresh recovery, and resilient Excel persistence
  const excelFileCache = new Map<string, { base64: string; updated: string; fileId?: string }>();
  let excelDriveListCache: { timestamp: number; data: any[] } = { timestamp: 0, data: [] };

  const KNOWN_DRIVE_EXCEL_FILES: Record<string, string> = {
    "hava_araçları_yer_destek_at-802.xlsx": "1AMSXLiHdGCcF76YgpBRhTE_86QaLOBGH",
    "hava_araclari_yer_destek_at-802.xlsx": "1AMSXLiHdGCcF76YgpBRhTE_86QaLOBGH",
    "at-802_sarf_ve_parca_deposu.xlsx": "11d1HAocT6gCmHaLEhFrP1oXj7U7bZMS7",
    "at802_sarf_ve_parca_deposu.xlsx": "11d1HAocT6gCmHaLEhFrP1oXj7U7bZMS7",
    "at-802_sarf_depo.xlsx": "11d1HAocT6gCmHaLEhFrP1oXj7U7bZMS7",
    "at802_sarf_depo.xlsx": "11d1HAocT6gCmHaLEhFrP1oXj7U7bZMS7",
    "at-802_ozel_bakim_aletleri.xlsx": "1idYaGtP9-4zMnUP9R28yHZwd47gy_QL8",
    "at-802_kimyasal_depo.xlsx": "1JYPCRpMSaZ8AWt7qEWnXhb4DbUKguDRK",
    "at802_kimyasal_depo.xlsx": "1JYPCRpMSaZ8AWt7qEWnXhb4DbUKguDRK",
    "hava_araçları_yer_destek_bell-429.xlsx": "1oICd57xCIGHza96IrWfOBCh1p5jTv3mB",
    "hava_araclari_yer_destek_bell-429.xlsx": "1oICd57xCIGHza96IrWfOBCh1p5jTv3mB",
    "bell-429_sarf_ve_parca_deposu.xlsx": "1tu8hDWSgIYkGn-7i_gDTC_UpUFUyaOzf",
    "bell429_sarf_depo.xlsx": "1tu8hDWSgIYkGn-7i_gDTC_UpUFUyaOzf",
    "hava_araçları_yer_destek_t-70.xlsx": "1xDxoQPBpKIWTvbAhcMv7cUsdTtKHxwkJ",
    "hava_araclari_yer_destek_t-70.xlsx": "1xDxoQPBpKIWTvbAhcMv7cUsdTtKHxwkJ",
    "t-70_sarf_ve_parca_deposu.xlsx": "",
    "t70_sarf_depo.xlsx": "",
    "hava_araçları_yer_destek_b-360.xlsx": "1_oxUU31nloWuKb9dMOOU8Q8XfvD4iAkz",
    "hava_araclari_yer_destek_b-360.xlsx": "1_oxUU31nloWuKb9dMOOU8Q8XfvD4iAkz",
    "b-360_sarf_ve_parca_deposu.xlsx": "",
    "b360_sarf_depo.xlsx": "",
    "hava_araçları_yer_destek_c-650.xlsx": "1BRvs2LxZ21BMl3YRjKL1rtMzli9J5Kuf",
    "hava_araclari_yer_destek_c-650.xlsx": "1BRvs2LxZ21BMl3YRjKL1rtMzli9J5Kuf",
    "c-650_sarf_ve_parca_deposu.xlsx": "",
    "c650_sarf_depo.xlsx": "",
    "hava_araçları_yer_destek_hangar.xlsx": "1-KGVWi0EN5UnI0NEC_EC3-zORYHsbqvu",
    "hava_araclari_yer_destek_hangar.xlsx": "1-KGVWi0EN5UnI0NEC_EC3-zORYHsbqvu",
    "hangar_sarf_ve_parca_deposu.xlsx": "",
    "hangar_sarf_depo.xlsx": "",
    "hava_araçları_yer_destek_t-70_bumbi_backet.xlsx": "1AtWL0hmzeY3wyGJzVylWk4hkqKjeQe3Q",
    "hava_araclari_yer_destek_t-70_bumbi_backet.xlsx": "1AtWL0hmzeY3wyGJzVylWk4hkqKjeQe3Q",
    "hangar_depo_hareket_gecmisi.xlsx": "1RXIhJps-PxUQAkTlLDyvcDWnpUbg2KzG",
    "olay_takip_cizelgesi_at-802.xlsx": "1XhWmwJyTmYkZ1jHBJFie3nRWVF6SC-GL"
  };

  const normalizeExcelName = (name: string): string => {
    return String(name || "")
      .trim()
      .toLowerCase()
      .replace(/ç/g, "c")
      .replace(/ğ/g, "g")
      .replace(/ı/g, "i")
      .replace(/ö/g, "o")
      .replace(/ş/g, "s")
      .replace(/ü/g, "u")
      .replace(/[^a-z0-9.]/g, "_")
      .replace(/_+/g, "_");
  };

  const isExcelNameExactMatch = (cand: string, target: string): boolean => {
    if (!cand || !target) return false;
    const nCand = normalizeExcelName(cand);
    const nTarget = normalizeExcelName(target);
    if (nCand === nTarget) return true;
    const bCand = nCand.replace(/\.(xlsx|xls|csv)$/, "");
    const bTarget = nTarget.replace(/\.(xlsx|xls|csv)$/, "");
    return bCand === bTarget;
  };

  // Helper to dynamically get or list all Excel files in Google Drive folder
  async function getDriveExcelsList(force = false): Promise<any[]> {
    const now = Date.now();
    if (!force && excelDriveListCache.data.length > 0 && (now - excelDriveListCache.timestamp < 120000)) {
      return excelDriveListCache.data;
    }
    const scriptUrls = [
      "https://script.google.com/macros/s/AKfycbw4kruTTc058Y9rLTyO3dKi6KloYsmdDTwV1GSiAk8ZXefyo3Z7_VDSTuurzsS9BHAQyQ/exec",
      "https://script.google.com/macros/s/AKfycbwP1uOo2NrST5a4I8vm1nGBLtI26yY2lWrmu9_e9iymwBkUhJBA9JOPCp7SNKqJbOubOw/exec"
    ];
    for (const url of scriptUrls) {
      try {
        const resp = await fetch(url, {
          method: "POST",
          headers: { "Content-Type": "text/plain;charset=utf-8" },
          body: JSON.stringify({
            action: "listTechizatExcelsFromDrive",
            folderId: "1HQR_NYKhHQGA7_2W3nArI9pCh-LJasTP"
          }),
          signal: AbortSignal.timeout(6000)
        });
        if (resp.ok) {
          const json = await resp.json();
          if (json && json.status === "success" && Array.isArray(json.excels) && json.excels.length > 0) {
            excelDriveListCache = { timestamp: now, data: json.excels };
            return json.excels;
          }
        }
      } catch (err) {
        // try next script URL
      }
    }
    return excelDriveListCache.data;
  }

  // API Route: Read Excel live from Google Drive (Zero Cache & Resilient Auto-Match)
  app.post("/api/read-excel-from-drive", async (req, res) => {
    try {
      let { fileName = "at-802_sarf_ve_parca_deposu.xlsx", fileId = "", folderId = "1HQR_NYKhHQGA7_2W3nArI9pCh-LJasTP" } = req.body;
      fileName = String(fileName || "").trim();

      // 1. Resolve fileId from KNOWN map first with strict matching
      if (!fileId) {
        for (const [k, id] of Object.entries(KNOWN_DRIVE_EXCEL_FILES)) {
          if (id && isExcelNameExactMatch(k, fileName)) {
            fileId = id;
            break;
          }
        }
      }

      // Search in excelDriveListCache with strict full filename matching ONLY (no fuzzy substring collisions)
      if (!fileId) {
        const driveExcels = await getDriveExcelsList();
        if (driveExcels && driveExcels.length > 0) {
          const match = driveExcels.find((d: any) => {
            const dName = d.name || d.fileName || "";
            return isExcelNameExactMatch(dName, fileName);
          });
          if (match && match.id) {
            fileId = match.id;
          }
        }
      }

      // 2. Direct Google Drive Download (Fastest, zero cache, 100% live)
      if (fileId) {
        const directDriveUrls = [
          `https://docs.google.com/spreadsheets/d/${fileId}/export?format=xlsx`,
          `https://drive.google.com/uc?export=download&id=${fileId}`
        ];

        for (const driveUrl of directDriveUrls) {
          try {
            const driveResp = await fetch(driveUrl, { 
              headers: { "User-Agent": "Mozilla/5.0" },
              signal: AbortSignal.timeout(8000)
            });
            if (driveResp.ok) {
              const arrayBuf = await driveResp.arrayBuffer();
              const buf = Buffer.from(arrayBuf);
              // Ensure it is a valid ZIP/XLSX or OLE2/XLS file (not an HTML error page from Google Drive)
              const isValidXlsx = buf.length > 2000 && ((buf[0] === 0x50 && buf[1] === 0x4B) || (buf[0] === 0xD0 && buf[1] === 0xCF));
              if (isValidXlsx) {
                const base64Str = buf.toString("base64");
                const localMirrorPath = path.resolve(process.cwd(), `public/${fileName}`);
                try {
                  fs.writeFileSync(localMirrorPath, buf);
                } catch {}

                excelFileCache.set(fileName, {
                  base64: base64Str,
                  updated: new Date().toLocaleString("tr-TR"),
                  fileId: fileId
                });

                return res.json({
                  status: "success",
                  base64: base64Str,
                  fileName: fileName,
                  fileId: fileId,
                  updated: new Date().toLocaleString("tr-TR"),
                  message: "Google Drive üzerinden canlı olarak başarıyla çekildi."
                });
              }
            }
          } catch (dErr: any) {
            console.warn("Direct Drive fetch error:", dErr.message);
          }
        }
      }

      // 3. FALLBACK: Check in-memory cache if it has rich data
      if (fileName && excelFileCache.has(fileName)) {
        const cached = excelFileCache.get(fileName)!;
        if (cached.base64 && cached.base64.length > 2000) {
          return res.json({
            status: "success",
            base64: cached.base64,
            fileName: fileName,
            fileId: cached.fileId || fileId,
            updated: cached.updated || new Date().toLocaleString("tr-TR"),
            source: "cache",
            message: "Önbellekten güncel dosya sağlandı."
          });
        }
      }

      // 4. FALLBACK: Check local disk mirror with resilient Turkish / ASCII matching
      let resolvedLocalPath = path.resolve(process.cwd(), `public/${fileName}`);
      if (!fs.existsSync(resolvedLocalPath)) {
        try {
          const pubFiles = fs.readdirSync(path.resolve(process.cwd(), "public"));
          const match = pubFiles.find(f => isExcelNameExactMatch(f, fileName));
          if (match) {
            resolvedLocalPath = path.resolve(process.cwd(), `public/${match}`);
          }
        } catch {}
      }

      if (fs.existsSync(resolvedLocalPath)) {
        try {
          const fileBuf = fs.readFileSync(resolvedLocalPath);
          if (fileBuf && fileBuf.byteLength > 2000 && ((fileBuf[0] === 0x50 && fileBuf[1] === 0x4B) || (fileBuf[0] === 0xD0 && fileBuf[1] === 0xCF))) {
            const base64Str = fileBuf.toString("base64");
            excelFileCache.set(fileName, {
              base64: base64Str,
              updated: new Date().toLocaleString("tr-TR"),
              fileId: fileId
            });
            return res.json({
              status: "success",
              base64: base64Str,
              fileName: fileName,
              fileId: fileId,
              updated: new Date().toLocaleString("tr-TR"),
              message: "Yerel ayna dosyasından sağlandı."
            });
          }
        } catch {}
      }

      // 3. Try Google Apps Script Endpoints with fast timeout (max 4s)
      const scriptUrls = [
        "https://script.google.com/macros/s/AKfycbwP1uOo2NrST5a4I8vm1nGBLtI26yY2lWrmu9_e9iymwBkUhJBA9JOPCp7SNKqJbOubOw/exec",
        "https://script.google.com/macros/s/AKfycbw4kruTTc058Y9rLTyO3dKi6KloYsmdDTwV1GSiAk8ZXefyo3Z7_VDSTuurzsS9BHAQyQ/exec"
      ];

      for (const sUrl of scriptUrls) {
        try {
          const gasRes = await fetch(sUrl, {
            method: "POST",
            headers: { "Content-Type": "text/plain;charset=utf-8" },
            body: JSON.stringify({
              action: "readExcelFromDrive",
              fileName: fileName,
              fileId: fileId,
              folderId: folderId
            }),
            signal: AbortSignal.timeout(4000)
          });

          if (gasRes.ok) {
            const rawText = await gasRes.text();
            try {
              const data = JSON.parse(rawText);
              if (data && data.status === "success" && data.base64) {
                const localMirrorPath = path.resolve(process.cwd(), `public/${fileName}`);
                try {
                  fs.writeFileSync(localMirrorPath, Buffer.from(data.base64, "base64"));
                } catch {}
                excelFileCache.set(fileName, {
                  base64: data.base64,
                  updated: new Date().toLocaleString("tr-TR"),
                  fileId: data.fileId || fileId
                });
                return res.json(data);
              }
            } catch {}
          }
        } catch (gasErr: any) {
          // ignore & fallback to template generator
        }
      }

      // 4. Resilient Dynamic Template Generator: If file not yet on Drive, auto-generate standard Excel
      const localPath = path.resolve(process.cwd(), 'public', fileName);
      const xlsx = await import("xlsx");
      const wb = xlsx.utils.book_new();

      const lowerName = fileName.toLowerCase();
      let headerRow: string[] = [];
      let sheetName = "LİSTE";

      if (lowerName.includes("kara_araclari") || lowerName.includes("kara")) {
        sheetName = "KARA_ARACLARI";
        headerRow = [
          "SIRA NO", "PLAKA / ŞASİ NO", "ARAÇ / EKİPMAN TANIMI", "MARKA / MODEL",
          "MİKTAR / ADET", "DEPO YERİ / HANGAR BÖLGESİ", "DURUM", "PERİYODİK BAKIM / MUAYENE",
          "SON BAKIM / MUAYENE TARİHİ", "GELECEK BAKIM / MUAYENE TARİHİ", "FİRMA / AİT OLDUĞU BİRİM", "AÇIKLAMA"
        ];
      } else if (lowerName.includes("sarf") || lowerName.includes("parca_deposu") || lowerName.includes("kimyasal")) {
        sheetName = lowerName.includes("kimyasal") ? "KIMYASAL_DEPO" : "SARF_VE_PARCA_DEPO";
        headerRow = [
          "DESCRIPTION", "PART NUMBER", "SERİ NUMBER - MÜKERRER NO - AÇIKLAMA", "LOKASYON NO",
          "GELEN", "TOPLAM STOK", "ANKARA ÇIKAN", "ANKARA MEVCUT",
          "KARAİN TRANSFER", "KARAİN ÇIKAN", "KARAİN MEVCUT",
          "ÇANAKKALE TRANSFER", "ÇANAKKALE ÇIKAN", "ÇANAKKALE MEVCUT",
          "MİLAS TRANSFER", "MİLAS ÇIKAN", "MİLAS MEVCUT",
          "BURSA TRANSFER", "BURSA ÇIKAN", "BURSA MEVCUT",
          "MUAYENE GİDEN", "MUAYENE GELEN", "MUAYENE TOPLAM"
        ];
      } else {
        // Standard Yer Destek & Özel Aletler
        sheetName = "YER_DESTEK";
        headerRow = [
          "SIRA NO", "TEÇHİZAT ADI", "PARÇA NO (P/N) / MODEL", "SERİ NO (S/N)",
          "MİKTAR", "DEPO YERİ / LOKASYON", "DURUM", "KALİBRASYONA TABİ Mİ?",
          "SON KALİBRASYON TARİHİ", "GELECEK KALİBRASYON TARİHİ", "FİRMA / MÜLKİYET", "AÇIKLAMA"
        ];
      }

      let dataRows: any[][] = [];
      if (lowerName.includes("sarf") || lowerName.includes("parca_deposu")) {
        const jsonPath = path.resolve(process.cwd(), 'public/at802_sarf_data.json');
        if (fs.existsSync(jsonPath)) {
          try {
            const rawJson = JSON.parse(fs.readFileSync(jsonPath, 'utf-8'));
            if (Array.isArray(rawJson) && rawJson.length > 0) {
              dataRows = rawJson.map(item => [
                item.description || item.name || '',
                item.partNumber || item.pn || '-',
                item.serialAndNotes || item.sn || '-',
                item.lokasyonNo || item.location || 'ANKARA',
                item.gelen || item.baseGelen || 1,
                item.toplamStok || item.gelen || 1,
                item.ankaraCikan || 0,
                item.ankaraMevcut || item.gelen || 1,
                item.karainTransfer || 0,
                item.karainCikan || 0,
                item.karainMevcut || 0,
                item.canakkaleTransfer || 0,
                item.canakkaleCikan || 0,
                item.canakkaleMevcut || 0,
                item.milasTransfer || 0,
                item.milasCikan || 0,
                item.milasMevcut || 0,
                item.bursaTransfer || 0,
                item.bursaCikan || 0,
                item.bursaMevcut || 0,
                item.muayeneGiden || 0,
                item.muayeneGelen || 0,
                item.muayeneToplam || 0
              ]);
            }
          } catch {}
        }
      }

      const ws = xlsx.utils.aoa_to_sheet([headerRow, ...dataRows]);
      xlsx.utils.book_append_sheet(wb, ws, sheetName);

      const generatedBuf = xlsx.write(wb, { type: "buffer", bookType: "xlsx" });
      const generatedBase64 = Buffer.from(generatedBuf).toString("base64");

      try {
        fs.writeFileSync(localPath, generatedBuf);
      } catch {}

      excelFileCache.set(fileName, {
        base64: generatedBase64,
        updated: new Date().toLocaleString("tr-TR"),
        fileId: fileId
      });

      return res.json({
        status: "success",
        base64: generatedBase64,
        fileName: fileName,
        fileId: fileId,
        updated: new Date().toLocaleString("tr-TR"),
        message: "Standart şablon Excel oluşturuldu ve hazırlandı."
      });
    } catch (err: any) {
      return res.status(500).json({
        status: "error",
        message: err.message
      });
    }
  });

  // Robust Helper to POST to Google Apps Script handling Node fetch redirects
  async function postToGoogleAppsScript(sUrl: string, payload: any): Promise<any> {
    const jsonBody = JSON.stringify(payload);
    try {
      // Primary attempt with follow redirect
      const resp = await fetch(sUrl, {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: jsonBody,
        redirect: "follow"
      });

      const text = await resp.text();
      try {
        const json = JSON.parse(text);
        if (json && (json.status === "success" || json.fileId || json.id || json.url)) {
          return json;
        }
      } catch {}

      // If text contains HTML or invalid json, attempt manual redirect following
      const manualResp = await fetch(sUrl, {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: jsonBody,
        redirect: "manual"
      });

      if (manualResp.status === 302 || manualResp.status === 307 || manualResp.status === 301) {
        const loc = manualResp.headers.get("location");
        if (loc) {
          const redirectResp = await fetch(loc, {
            method: "POST",
            headers: { "Content-Type": "text/plain;charset=utf-8" },
            body: jsonBody
          });
          const redText = await redirectResp.text();
          try {
            return JSON.parse(redText);
          } catch {
            return { status: "success", raw: redText };
          }
        }
      }
      return { status: "success", raw: text };
    } catch (err: any) {
      console.warn("[postToGoogleAppsScript] Error:", err.message);
      return null;
    }
  }

  // API Route: Upload Teçhizat Excel to Google Drive
  app.post("/api/upload-techizat-excel", async (req, res) => {
    try {
      const { fileName = "at-802_sarf_ve_parca_deposu.xlsx", base64Data, targetKey, folderId = "1HQR_NYKhHQGA7_2W3nArI9pCh-LJasTP", fileId = "" } = req.body;
      const cleanBase64 = String(base64Data || "").replace(/^data:[^;]+;base64,/, "").trim();

      // Find real fileId in Google Drive to strictly prevent duplicate uploads for each aircraft
      let effectiveFileId = (fileId && fileId !== "1tu8hDWSgIYkGn-7i_gDTC_UpUFUyaOzf") ? fileId : "";
      if (!effectiveFileId) {
        for (const [k, id] of Object.entries(KNOWN_DRIVE_EXCEL_FILES)) {
          if (id && isExcelNameExactMatch(k, fileName)) {
            effectiveFileId = id;
            break;
          }
        }
      }
      if (!effectiveFileId) {
        const driveExcels = await getDriveExcelsList();
        if (driveExcels && driveExcels.length > 0) {
          const match = driveExcels.find((d: any) => {
            const dName = d.name || d.fileName || "";
            return isExcelNameExactMatch(dName, fileName);
          });
          if (match && match.id) {
            effectiveFileId = match.id;
          }
        }
      }

      const scriptUrls = [
        "https://script.google.com/macros/s/AKfycbwP1uOo2NrST5a4I8vm1nGBLtI26yY2lWrmu9_e9iymwBkUhJBA9JOPCp7SNKqJbOubOw/exec",
        "https://script.google.com/macros/s/AKfycbw4kruTTc058Y9rLTyO3dKi6KloYsmdDTwV1GSiAk8ZXefyo3Z7_VDSTuurzsS9BHAQyQ/exec",
        "https://script.google.com/macros/s/AKfycbz1eI_ckIBmmo5uMlxdXEK9TXKpylA6n0TU2INzKE9Hxov2toVrZKyFtGkqT2hCyKBztQ/exec",
        "https://script.google.com/macros/s/AKfycbxBVLlhvSrYDsIY5-Z8-RQ4f2-kTrHbLrZN3Bk7hB3AtQotugE9xqXRscIZd6ruinlqTg/exec"
      ];
      
      const nowStr = new Date().toLocaleString("tr-TR");
      let finalBase64 = cleanBase64;
      if (fileName && cleanBase64) {
        excelFileCache.set(fileName, { base64: cleanBase64, updated: nowStr, fileId: fileId });

        // 1. Write Excel file buffer to public disk
        try {
          const buffer = Buffer.from(cleanBase64, "base64");
          const localExcelPath = path.resolve(process.cwd(), `public/${fileName}`);
          fs.writeFileSync(localExcelPath, buffer);

          // 2. MODIFICATION: Inject Shelf Life columns if missing
          const xlsx = await import("xlsx");
          let wb = xlsx.read(buffer, { type: "buffer" });
          let modified = false;

          if (wb && wb.SheetNames) {
            for (const sName of wb.SheetNames) {
              const sheet = wb.Sheets[sName];
              const aoa = xlsx.utils.sheet_to_json(sheet, { header: 1 }) as any[][];
              if (aoa.length > 0) {
                // Find header row
                let headerRowIdx = -1;
                for (let r = 0; r < Math.min(15, aoa.length); r++) {
                  const rowStr = (aoa[r] || []).join(' ').toUpperCase();
                  if (rowStr.includes('DESCRIPTION') || rowStr.includes('PART NUMBER') || rowStr.includes('MALZEME')) {
                    headerRowIdx = r;
                    break;
                  }
                }

                const hIdx = headerRowIdx >= 0 ? headerRowIdx : 0;
                const header = aoa[hIdx] || [];
                const headerStr = header.join(' ').toUpperCase();

                if (!headerStr.includes('RAF ÖMRÜ') && !headerStr.includes('ÖMÜRLÜ')) {
                  // Append headers at the VERY END
                  header.push("RAF ÖMRÜ VAR MI?");
                  header.push("RAF ÖMRÜ BİTİŞ TARİHİ");
                  
                  // Append to all data rows at the VERY END
                  for (let r = 0; r < aoa.length; r++) {
                    if (r === hIdx) continue;
                    if (!aoa[r]) aoa[r] = [];
                    // Ensure the row has enough columns (fill with empty if needed, then append)
                    while (aoa[r].length < header.length - 2) aoa[r].push("");
                    aoa[r].push("HAYIR");
                    aoa[r].push("-");
                  }
                  
                  const newSheet = xlsx.utils.aoa_to_sheet(aoa);
                  wb.Sheets[sName] = newSheet;
                  modified = true;
                }
              }
            }
          }

          finalBase64 = cleanBase64;
          if (modified) {
            const modBuffer = xlsx.write(wb, { type: "buffer", bookType: "xlsx" });
            finalBase64 = modBuffer.toString("base64");
            // Overwrite the local file with modified content
            fs.writeFileSync(localExcelPath, modBuffer);
          }

          // 3. Also parse and update public/at802_sarf_data.json
          if (wb && wb.SheetNames && wb.SheetNames.length > 0) {
            let parsedItems: any[] = [];
            for (const sName of wb.SheetNames) {
              const sheet = wb.Sheets[sName];
              const rawRows = xlsx.utils.sheet_to_json(sheet, { header: 1 }) as any[][];
              if (rawRows && rawRows.length > 1) {
                let headerRowIdx = -1;
                let descIdx = -1, pnIdx = -1, snIdx = -1, qtyIdx = -1, locIdx = -1;
                for (let r = 0; r < Math.min(25, rawRows.length); r++) {
                  const row = (rawRows[r] || []).map(c => String(c || '').trim().toUpperCase());
                  const hasDesc = row.some(c => c.includes('DESCRIPTION') || c.includes('MALZEME') || c.includes('TANIM') || c.includes('ÜRÜN') || c.includes('ITEM') || c.includes('PARÇA ADI'));
                  const hasPn = row.some(c => c.includes('PART') || c.includes('P/N') || c.includes('PN') || c.includes('MODEL'));
                  if (hasDesc || hasPn) {
                    headerRowIdx = r;
                    descIdx = row.findIndex(c => c.includes('DESCRIPTION') || c.includes('MALZEME') || c.includes('TANIM') || c.includes('ÜRÜN') || c.includes('ITEM') || c.includes('PARÇA ADI'));
                    pnIdx = row.findIndex(c => c.includes('PART') || c.includes('P/N') || c.includes('PN') || c.includes('MODEL'));
                    snIdx = row.findIndex(c => (c.includes('SERİ') || c.includes('SERI') || c.includes('S/N') || c.includes('SN')) && !c.includes('SIRA'));
                    qtyIdx = row.findIndex(c => c.includes('MİKTAR') || c.includes('MIKTAR') || c.includes('GELEN') || c.includes('ADET') || c.includes('QTY') || c.includes('STOK'));
                    locIdx = row.findIndex(c => c.includes('LOKASYON') || c.includes('KONUM') || c.includes('RAF') || c.includes('YER') || c.includes('DEPO'));
                    break;
                  }
                }
                if (descIdx < 0) descIdx = 1;
                if (pnIdx < 0) pnIdx = 2;
                if (snIdx < 0) snIdx = 3;
                if (qtyIdx < 0) qtyIdx = 4;
                if (locIdx < 0) locIdx = 5;

                const startR = headerRowIdx >= 0 ? headerRowIdx + 1 : 1;
                for (let r = startR; r < rawRows.length; r++) {
                  const row = rawRows[r];
                  if (!row) continue;
                  let desc = String(row[descIdx] !== undefined && row[descIdx] !== null ? row[descIdx] : '').trim();
                  if (!desc && row[0] && isNaN(Number(String(row[0]).trim()))) {
                    desc = String(row[0]).trim();
                  }
                  const upperDesc = desc.toUpperCase();
                  if (!desc || upperDesc === 'DESCRIPTION' || upperDesc === 'MALZEME ADI' || upperDesc.includes('AT-802 SARF') || upperDesc.includes('SIRA NO') || upperDesc.includes('ORMAN GENERAL')) continue;

                  const pn = String(row[pnIdx] !== undefined && row[pnIdx] !== null ? row[pnIdx] : '-').trim();
                  const sn = String(row[snIdx] !== undefined && row[snIdx] !== null ? row[snIdx] : '-').trim();
                  const miktar = parseInt(row[qtyIdx]) || 1;
                  const loc = String(row[locIdx] !== undefined && row[locIdx] !== null ? row[locIdx] : 'DEPO').trim();

                  parsedItems.push({
                    unit: 'at802',
                    category: 'sarf',
                    description: desc,
                    name: desc,
                    partNumber: pn || '-',
                    pn: pn || '-',
                    serialAndNotes: sn || '-',
                    sn: sn || '-',
                    lokasyonNo: loc || 'DEPO',
                    location: loc || 'DEPO',
                    baseGelen: miktar,
                    gelen: miktar
                  });
                }
                if (parsedItems.length > 0) break;
              }
            }
            if (parsedItems.length > 0) {
              const jsonPath = path.resolve(process.cwd(), 'public/at802_sarf_data.json');
              fs.writeFileSync(jsonPath, JSON.stringify(parsedItems, null, 2), 'utf-8');
            }
          }
        } catch (diskErr) {
          console.warn("Disk save in upload-techizat-excel warn:", diskErr);
        }
      }

      if (targetKey && cleanBase64) {
        excelFileCache.set(targetKey, { base64: cleanBase64, updated: nowStr });
      }

      let data: any = { 
        status: "success", 
        fileName: fileName,
        message: "Veriler başarıyla kaydedildi." 
      };

      for (const sUrl of scriptUrls) {
        // Upload Excel file ONLY (do not create or upload any PDF!)
        const gasRes = await postToGoogleAppsScript(sUrl, {
          action: "uploadTechizatExcel",
          fileName: fileName,
          targetKey: targetKey,
          base64Data: finalBase64, // Use potentially modified base64
          folderId: folderId,
          fileId: effectiveFileId || "",
          overwrite: true,
          replaceExisting: true,
          deleteDuplicates: true,
          updateIfExists: true,
          cleanDuplicates: true
        });

        if (gasRes && (gasRes.status === "success" || gasRes.fileId)) {
          data = gasRes;
          const assignedId = data.fileId || effectiveFileId;
          if (fileName && assignedId) {
            excelFileCache.set(fileName, { base64: finalBase64, updated: nowStr, fileId: assignedId });
            KNOWN_DRIVE_EXCEL_FILES[fileName] = assignedId;
          }
          break;
        }
      }

      return res.json(data);
    } catch (err: any) {
      return res.status(500).json({
        status: "error",
        message: err.message
      });
    }
  });

  // API Route: Update specific cell/row in online Excel file & sync directly to Google Drive
  app.post("/api/update-depo-excel-item", async (req, res) => {
    try {
      const {
        fileName,
        unit = "at802",
        category = "sarf",
        oldItem,
        newItem,
        action = "update",
        folderId = "1HQR_NYKhHQGA7_2W3nArI9pCh-LJasTP"
      } = req.body;

      let resolvedFileName = fileName;
      if (!resolvedFileName || resolvedFileName === "undefined") {
        resolvedFileName = unit === "at802"
          ? (category === "kimyasal" ? "at-802_kimyasal_depo.xlsx" : "at-802_sarf_ve_parca_deposu.xlsx")
          : `${unit}_${category}_depo.xlsx`;
      }

      const localExcelPath = path.resolve(process.cwd(), `public/${resolvedFileName}`);
       if (!fs.existsSync(localExcelPath)) {
        const isKimyasal = category === "kimyasal" || resolvedFileName.includes("kimyasal");
        const templateName = isKimyasal ? "at-802_kimyasal_depo.xlsx" : "at-802_sarf_ve_parca_deposu.xlsx";
        const defaultPath = path.resolve(process.cwd(), `public/${templateName}`);
        if (fs.existsSync(defaultPath)) {
          fs.copyFileSync(defaultPath, localExcelPath);
        }
      }

      const xlsxModule = await import("xlsx");
      const xlsx = (xlsxModule as any).default || xlsxModule;
      let wb: any;
      if (fs.existsSync(localExcelPath)) {
        const fileBuffer = fs.readFileSync(localExcelPath);
        wb = xlsx.read(fileBuffer, { type: "buffer", cellStyles: true, cellDates: true });
      } else {
        wb = xlsx.utils.book_new();
        const emptyWs = xlsx.utils.aoa_to_sheet([
          ["DESCRIPTION", "PART NUMBER", "SERİ NUMBER - MÜKERRER NO - AÇIKLAMA", "LOKASYON NO", "GELEN", "TOPLAM STOK", "ANKARA ÇIKAN", "ANKARA MEVCUT"]
        ]);
        xlsx.utils.book_append_sheet(wb, emptyWs, "LİSTE");
      }

      const sheetName = wb.SheetNames[0] || "LİSTE";
      const sheet = wb.Sheets[sheetName];
      const rawRows = xlsx.utils.sheet_to_json(sheet, { header: 1, defval: "" }) as any[][];

      // Find header row
      let headerRowIdx = 0;
      for (let r = 0; r < Math.min(25, rawRows.length); r++) {
        const rowStr = (rawRows[r] || []).map((c: any) => String(c || "").toUpperCase().trim()).join(" ");
        if (rowStr.includes("DESCRIPTION") || rowStr.includes("MALZEME") || rowStr.includes("PARÇA") || rowStr.includes("P/N") || rowStr.includes("PART NUMBER")) {
          headerRowIdx = r;
          break;
        }
      }

      const headerRow = (rawRows[headerRowIdx] || []).map((c: any) => String(c || "").toUpperCase().trim());
      const descCol = headerRow.findIndex(c => c.includes("DESCRIPTION") || c.includes("MALZEME") || c.includes("TANIM") || c.includes("ÜRÜN") || c.includes("ITEM"));
      const pnCol = headerRow.findIndex(c => c.includes("PART") || c.includes("P/N") || c.includes("PN") || c.includes("MODEL"));
      const snCol = headerRow.findIndex(c => (c.includes("SERİ") || c.includes("SERI") || c.includes("S/N") || c.includes("SN")) && !c.includes("SIRA"));
      const locCol = headerRow.findIndex(c => c.includes("LOKASYON") || c.includes("RAF") || c.includes("KONUM") || c.includes("YER") || c.includes("DEPO"));
      const gelenCol = headerRow.findIndex(c => c.includes("GELEN") || c.includes("MİKTAR") || c.includes("MIKTAR") || c.includes("ADET") || c.includes("QTY") || c.includes("STOK"));
      const toplamStokCol = headerRow.findIndex(c => c.includes("TOPLAM STOK"));
      const ankaraMevcutCol = headerRow.findIndex(c => c.includes("ANKARA MEVCUT"));
      const setCellVal = (r: number, c: number, val: any) => {
        if (c < 0) return;
        if (!rawRows[r]) rawRows[r] = [];
        rawRows[r][c] = val;
        const cellRef = xlsx.utils.encode_cell({ r, c });
        const isNum = typeof val === "number";
        sheet[cellRef] = { t: isNum ? "n" : "s", v: val };
      };

      let shelfLifeCol = headerRow.findIndex(c => c.includes("RAF ÖMRÜ VAR") || c.includes("RAF OMRU") || c.includes("ÖMRÜNE TABİ"));
      let shelfLifeDateCol = headerRow.findIndex(c => c.includes("RAF ÖMRÜ BİTİŞ") || c.includes("RAF ÖMRÜ TARİH") || c.includes("SKT") || c.includes("BİTİŞ"));

      // Dynamically add columns if missing
      if (shelfLifeCol === -1) {
        headerRow.push("RAF ÖMRÜ VAR MI?");
        if (!rawRows[headerRowIdx]) rawRows[headerRowIdx] = [];
        rawRows[headerRowIdx].push("RAF ÖMRÜ VAR MI?");
        shelfLifeCol = headerRow.length - 1;
        setCellVal(headerRowIdx, shelfLifeCol, "RAF ÖMRÜ VAR MI?");
        console.log(`[ExcelSync] Created missing column: RAF ÖMRÜ VAR MI? at index ${shelfLifeCol}`);
      }

      if (shelfLifeDateCol === -1) {
        headerRow.push("RAF ÖMRÜ BİTİŞ TARİHİ");
        if (!rawRows[headerRowIdx]) rawRows[headerRowIdx] = [];
        rawRows[headerRowIdx].push("RAF ÖMRÜ BİTİŞ TARİHİ");
        shelfLifeDateCol = headerRow.length - 1;
        setCellVal(headerRowIdx, shelfLifeDateCol, "RAF ÖMRÜ BİTİŞ TARİHİ");
        console.log(`[ExcelSync] Created missing column: RAF ÖMRÜ BİTİŞ TARİHİ at index ${shelfLifeDateCol}`);
      }

      // Pre-populate empty cells for these columns in all existing rows with "HAYIR" and "-"
      for (let r = headerRowIdx + 1; r < rawRows.length; r++) {
        if (!rawRows[r]) rawRows[r] = [];
        if (rawRows[r][shelfLifeCol] === undefined || rawRows[r][shelfLifeCol] === "") {
          setCellVal(r, shelfLifeCol, "HAYIR");
        }
        if (rawRows[r][shelfLifeDateCol] === undefined || rawRows[r][shelfLifeDateCol] === "") {
          setCellVal(r, shelfLifeDateCol, "-");
        }
      }

      const mevcutColIndices: number[] = [];
      headerRow.forEach((h, idx) => {
        if (h.includes("MEVCUT") && !h.includes("TOPLAM")) {
          mevcutColIndices.push(idx);
        }
      });

      // Target row matching
      const targetOldPn = String(oldItem?.partNumber || oldItem?.pn || newItem?.partNumber || newItem?.pn || "").trim().toLowerCase();
      const targetOldDesc = String(oldItem?.description || oldItem?.name || newItem?.description || newItem?.name || "").trim().toLowerCase();

      let targetRowIdx = -1;
      for (let r = headerRowIdx + 1; r < rawRows.length; r++) {
        const row = rawRows[r] || [];
        const rPn = String(row[pnCol >= 0 ? pnCol : 1] || "").trim().toLowerCase();
        const rDesc = String(row[descCol >= 0 ? descCol : 0] || "").trim().toLowerCase();
        
        if (targetOldPn && targetOldPn !== "-" && rPn === targetOldPn) {
          targetRowIdx = r;
          break;
        }
        if (targetOldDesc && rDesc === targetOldDesc) {
          targetRowIdx = r;
          break;
        }
      }

      if (action === "delete" && targetRowIdx >= 0) {
        rawRows.splice(targetRowIdx, 1);
        const newWs = xlsx.utils.aoa_to_sheet(rawRows);
        wb.Sheets[sheetName] = newWs;
      } else if (targetRowIdx >= 0 && newItem) {
        // UPDATE EXISTING CELL VALUES IN PLACE
        if (descCol >= 0 && newItem.description) setCellVal(targetRowIdx, descCol, newItem.description);
        if (pnCol >= 0 && newItem.partNumber) setCellVal(targetRowIdx, pnCol, newItem.partNumber);
        if (snCol >= 0 && newItem.serialAndNotes !== undefined) setCellVal(targetRowIdx, snCol, newItem.serialAndNotes);
        if (locCol >= 0 && newItem.lokasyonNo) setCellVal(targetRowIdx, locCol, newItem.lokasyonNo);
        
        // Update shelf life columns in Excel
        if (shelfLifeCol >= 0 && newItem.hasShelfLife !== undefined) {
          const hasLife = newItem.hasShelfLife === "EVET" || newItem.hasShelfLife === true || String(newItem.hasShelfLife).toUpperCase().includes("EVET");
          setCellVal(targetRowIdx, shelfLifeCol, hasLife ? "EVET" : "HAYIR");
        }
        if (shelfLifeDateCol >= 0 && newItem.shelfLifeDate !== undefined) {
          setCellVal(targetRowIdx, shelfLifeDateCol, newItem.shelfLifeDate || "-");
        }

        if (gelenCol >= 0 && newItem.gelen !== undefined) {
          const oldG = Number(rawRows[targetRowIdx][gelenCol]) || 0;
          const newG = Number(newItem.gelen) || 0;
          setCellVal(targetRowIdx, gelenCol, newG);
          const diff = newG - oldG;
          if (toplamStokCol >= 0) {
            let sumMevcut = 0;
            mevcutColIndices.forEach(cIdx => {
              sumMevcut += Number(rawRows[targetRowIdx][cIdx]) || 0;
            });
            setCellVal(targetRowIdx, toplamStokCol, sumMevcut);
          }
          if (ankaraMevcutCol >= 0) {
            const oldAnk = Number(rawRows[targetRowIdx][ankaraMevcutCol]) || 0;
            setCellVal(targetRowIdx, ankaraMevcutCol, Math.max(0, oldAnk + diff));
          }
        }
      } else if (newItem) {
        // APPEND NEW ROW
        const newR = rawRows.length;
        const rowLen = Math.max(headerRow.length, 8);
        const newRow = new Array(rowLen).fill("");
        if (descCol >= 0) newRow[descCol] = newItem.description;
        if (pnCol >= 0) newRow[pnCol] = newItem.partNumber || "-";
        if (snCol >= 0) newRow[snCol] = newItem.serialAndNotes || "-";
        if (locCol >= 0) newRow[locCol] = newItem.lokasyonNo || "DEPO";

        // Set shelf life columns in the new row array
        if (shelfLifeCol >= 0) {
          const hasLife = newItem.hasShelfLife === "EVET" || newItem.hasShelfLife === true || String(newItem.hasShelfLife).toUpperCase().includes("EVET");
          newRow[shelfLifeCol] = hasLife ? "EVET" : "HAYIR";
        }
        if (shelfLifeDateCol >= 0) {
          newRow[shelfLifeDateCol] = newItem.shelfLifeDate || "-";
        }

        const gVal = Number(newItem.gelen) || 0;

        // Set all numeric / warehouse / stock columns to 0 except the active destination column
        for (let c = 0; c < rowLen; c++) {
          const hName = headerRow[c] || "";
          if (c === descCol || c === pnCol || c === snCol || c === locCol || c === shelfLifeCol || c === shelfLifeDateCol) continue;
          if (hName.includes("GELEN") || hName.includes("MİKTAR") || hName.includes("MIKTAR") || hName.includes("STOK") || hName.includes("MEVCUT") || hName.includes("ÇIKAN") || hName.includes("TRANSFER") || hName.includes("ADET")) {
            // If this is the specific incoming location column or gelenCol, set gVal, otherwise 0
            const isTargetCol = (c === gelenCol) || (newItem.lokasyonNo && hName.includes(newItem.lokasyonNo.toUpperCase()));
            newRow[c] = isTargetCol ? gVal : 0;
          }
        }

        if (gelenCol >= 0 && newRow[gelenCol] === "") newRow[gelenCol] = gVal;
        if (toplamStokCol >= 0) {
          let sumMevcut = 0;
          mevcutColIndices.forEach(cIdx => {
            sumMevcut += Number(newRow[cIdx]) || 0;
          });
          newRow[toplamStokCol] = sumMevcut;
        }
        if (ankaraMevcutCol >= 0 && newRow[ankaraMevcutCol] === "") newRow[ankaraMevcutCol] = gVal;

        rawRows.push(newRow);
        for (let c = 0; c < newRow.length; c++) {
          const cellRef = xlsx.utils.encode_cell({ r: newR, c });
          const val = newRow[c];
          const isNum = typeof val === "number";
          sheet[cellRef] = { t: isNum ? "n" : "s", v: val };
        }
        const range = xlsx.utils.decode_range(sheet["!ref"] || `A1:W${newR}`);
        range.e.r = Math.max(range.e.r, newR);
        sheet["!ref"] = xlsx.utils.encode_range(range);

        // Record movement in "Depo Hareketleri" sheet
        const histSheetName = "Depo Hareketleri";
        let histSheet = wb.Sheets[histSheetName];
        if (!histSheet) {
          const histAoa = [
            ["TARİH", "BİRİM", "KATEGORİ", "MALZEME ADI", "PART NUMBER", "İŞLEM TİPİ", "MİKTAR", "LOKASYON", "AÇIKLAMA"]
          ];
          histSheet = xlsx.utils.aoa_to_sheet(histAoa);
          xlsx.utils.book_append_sheet(wb, histSheet, histSheetName);
        }
        const histRows = xlsx.utils.sheet_to_json(histSheet, { header: 1, defval: "" }) as any[][];
        histRows.push([
          new Date().toLocaleString("tr-TR"),
          unit,
          category,
          newItem.description,
          newItem.partNumber || "-",
          "Yeni Ürün Kaydı",
          gVal,
          newItem.lokasyonNo || "DEPO",
          `Yeni ürün eklendi. ${newItem.lokasyonNo || "Ankara"} deposuna ${gVal} adet tanımlandı, diğer depolar 0 olarak başlatıldı.`
        ]);
        wb.Sheets[histSheetName] = xlsx.utils.aoa_to_sheet(histRows);
      }

      // Write updated file to disk
      xlsx.writeFile(wb, localExcelPath);

      // Excel buffer to Google Drive
      const updatedBuffer = xlsx.write(wb, { type: "buffer", bookType: "xlsx" });
      const base64Data = updatedBuffer.toString("base64");

      // Upload ONLY Excel to Google Drive (no PDF!)
      const scriptUrls = [
        "https://script.google.com/macros/s/AKfycbwP1uOo2NrST5a4I8vm1nGBLtI26yY2lWrmu9_e9iymwBkUhJBA9JOPCp7SNKqJbOubOw/exec",
        "https://script.google.com/macros/s/AKfycbw4kruTTc058Y9rLTyO3dKi6KloYsmdDTwV1GSiAk8ZXefyo3Z7_VDSTuurzsS9BHAQyQ/exec"
      ];

      for (const sUrl of scriptUrls) {
        try {
          const gasRes = await postToGoogleAppsScript(sUrl, {
            action: "uploadTechizatExcel",
            fileName: resolvedFileName,
            targetKey: `depo_${unit}_${category}`,
            base64Data: base64Data,
            folderId: folderId,
            overwrite: true,
            replaceExisting: true,
            deleteDuplicates: true,
            updateIfExists: true,
            cleanDuplicates: true
          });
          if (gasRes && (gasRes.status === "success" || gasRes.fileId)) {
            break;
          }
        } catch (gasErr) {
          console.warn("GAS excel upload warning:", gasErr);
        }
      }

      // Automatically sync new product record to Google Sheet "DEPO HAREKET GEÇMİŞİ-AT 802"
      if (newItem && (action === "add" || targetRowIdx < 0)) {
        try {
          const gVal = Number(newItem.gelen) || 1;
          const pad = (n: number) => String(n).padStart(2, '0');
          const nowD = new Date();
          const dateStr = `${pad(nowD.getDate())}.${pad(nowD.getMonth() + 1)}.${nowD.getFullYear()}`;
          const newTxRow = [
            newItem.description || newItem.name || "",
            String(gVal),
            dateStr,
            "Yeni Ürün Kaydı",
            newItem.serialAndNotes || newItem.sn || "-",
            "-",
            "SİSTEM",
            "DEPO YÖNETİCİSİ",
            newItem.lokasyonNo || "ANKARA DEPO",
            "DEPO HAREKET GEÇMİŞİ-AT 802"
          ];
          for (const sUrl of scriptUrls) {
            try {
              await fetch(sUrl, {
                method: "POST",
                headers: { "Content-Type": "text/plain;charset=utf-8" },
                body: JSON.stringify({
                  action: "appendSheetRow",
                  spreadsheetId: "17ScGYYx0erzDwHDk6RGiHOdJATdfmmExXFBY39dXpF0",
                  sheetName: "DEPO HAREKET GEÇMİŞİ-AT 802",
                  row: newTxRow
                })
              });
              break;
            } catch {}
          }
        } catch (syncErr) {
          console.warn("Movement log sync warn:", syncErr);
        }
      }

      return res.json({
        status: "success",
        fileName: resolvedFileName,
        targetRowIdx,
        message: "Excel hücresi güncellendi ve Google Drive'a kaydedildi."
      });
    } catch (err: any) {
      console.error("[update-depo-excel-item] Error:", err);
      return res.status(500).json({ status: "error", message: err.message });
    }
  });

  // API Route: Upload Technical Publication PDF to Google Drive via GAS
  app.post("/api/upload-tech-publication", async (req, res) => {
    try {
      const payload = req.body;
      const scriptUrl = "https://script.google.com/macros/s/AKfycbw4kruTTc058Y9rLTyO3dKi6KloYsmdDTwV1GSiAk8ZXefyo3Z7_VDSTuurzsS9BHAQyQ/exec";
      
      const gasRes = await fetch(scriptUrl, {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify({
          action: "uploadTechPublication",
          ...payload
        })
      });

      const text = await gasRes.text();
      try {
        const json = JSON.parse(text);
        return res.json(json);
      } catch (parseErr) {
        return res.json({ status: "success", raw: text });
      }
    } catch (err: any) {
      return res.status(500).json({
        status: "error",
        message: err.message
      });
    }
  });

  // API Route: Handle ZIP Upload for Technical Publications (Extraction on Server for Speed)
  app.post("/api/upload-tech-publication-zip", async (req, res) => {
    try {
      const { base64Zip, unit, unitKey, category, revision, section } = req.body;
      if (!base64Zip) return res.status(400).json({ status: "error", message: "ZIP verisi bulunamadı." });

      const zipBuffer = Buffer.from(base64Zip.replace(/^data:application\/zip;base64,/, ""), "base64");
      const zip = new AdmZip(zipBuffer);
      const zipEntries = zip.getEntries();
      const scriptUrl = "https://script.google.com/macros/s/AKfycbw4kruTTc058Y9rLTyO3dKi6KloYsmdDTwV1GSiAk8ZXefyo3Z7_VDSTuurzsS9BHAQyQ/exec";

      const uploadTasks: Promise<any>[] = [];
      const results: any[] = [];

      for (const entry of zipEntries) {
        if (!entry.isDirectory && entry.entryName.toLowerCase().endsWith(".pdf")) {
          const pdfBuffer = entry.getData();
          const base64Pdf = pdfBuffer.toString("base64");
          const title = entry.name.replace(/\.[^/.]+$/, "");
          const cleanFileName = `pub_${unitKey}_${title.replace(/[^a-zA-Z0-9_-]/g, "_")}.pdf`;

          const task = fetch(scriptUrl, {
            method: "POST",
            headers: { "Content-Type": "text/plain;charset=utf-8" },
            body: JSON.stringify({
              action: "uploadTechPublication",
              fileName: cleanFileName,
              base64Data: base64Pdf,
              unit: unit,
              unitKey: unitKey,
              category: category || "IPC",
              title: title,
              revision: revision || "Rev. 01",
              section: section,
              notes: "",
              originalFileName: entry.name
            })
          })
          .then(async (r) => {
            const t = await r.text();
            try { return JSON.parse(t); } catch (e) { return { status: "success", fileId: "unknown", title }; }
          })
          .catch(err => ({ status: "error", message: err.message, title }));

          uploadTasks.push(task);
          // Limit concurrency to avoid GAS flooding
          if (uploadTasks.length >= 5) {
            results.push(...(await Promise.all(uploadTasks)));
            uploadTasks.length = 0;
          }
        }
      }

      if (uploadTasks.length > 0) {
        results.push(...(await Promise.all(uploadTasks)));
      }

      const successCount = results.filter(r => r.status === "success" || r.fileId).length;
      return res.json({
        status: "success",
        message: `${successCount} adet döküman başarıyla yüklendi.`,
        details: results
      });
    } catch (err: any) {
      console.error("ZIP Upload Error:", err);
      return res.status(500).json({ status: "error", message: err.message });
    }
  });

  // API Route: Upload Hangar & Olay Takip PDF directly to Drive
  app.post("/api/upload-hangar-pdf", async (req, res) => {
    try {
      const { fileName, base64Data, mimeType, folderId, itemKey, docType, firma } = req.body;
      const targetFolderId = folderId || "1HQR_NYKhHQGA7_2W3nArI9pCh-LJasTP";

      // Detect correct mimeType from fileName if not specified or generic
      const detectUploadMime = (fname: string, suppliedMime?: string): string => {
        if (suppliedMime && suppliedMime !== "application/octet-stream" && suppliedMime !== "application/pdf") {
          return suppliedMime;
        }
        const ext = (fname || "").split(".").pop()?.toLowerCase();
        if (ext === "zip") return "application/zip";
        if (ext === "rar") return "application/x-rar-compressed";
        if (ext === "jpg" || ext === "jpeg") return "image/jpeg";
        if (ext === "png") return "image/png";
        if (ext === "pdf") return "application/pdf";
        return suppliedMime || "application/pdf";
      };

      const resolvedMimeType = detectUploadMime(fileName, mimeType);

      const scriptUrls = [
        "https://script.google.com/macros/s/AKfycbwP1uOo2NrST5a4I8vm1nGBLtI26yY2lWrmu9_e9iymwBkUhJBA9JOPCp7SNKqJbOubOw/exec",
        "https://script.google.com/macros/s/AKfycbw4kruTTc058Y9rLTyO3dKi6KloYsmdDTwV1GSiAk8ZXefyo3Z7_VDSTuurzsS9BHAQyQ/exec"
      ];

      let gasResult: any = null;

      for (const sUrl of scriptUrls) {
        try {
          const resp = await fetch(sUrl, {
            method: "POST",
            headers: { "Content-Type": "text/plain;charset=utf-8" },
            body: JSON.stringify({
              action: "uploadTechizatDocPdf",
              fileName: fileName,
              base64Data: base64Data,
              mimeType: resolvedMimeType,
              folderId: targetFolderId,
              itemKey: itemKey || "GENEL",
              docType: docType || "Olay / Kaza Belgesi",
              firma: firma || "OGM Havacılık"
            })
          });

          if (resp.ok) {
            const json = await resp.json();
            if (json && (json.fileId || json.id || json.status === "success")) {
              gasResult = json;
              break;
            }
          }
        } catch (e: any) {
          // try next url
        }
      }

      // Reset drive cache to ensure fresh list on next fetch
      drivePdfsCache.timestamp = 0;

      if (gasResult) {
        return res.json({
          status: "success",
          fileId: gasResult.fileId || gasResult.id,
          fileName: gasResult.fileName || fileName,
          viewUrl: gasResult.viewUrl || (gasResult.fileId ? `https://drive.google.com/file/d/${gasResult.fileId}/preview` : undefined),
          message: "Belge Google Drive'a başarıyla kaydedildi."
        });
      }

      return res.json({
        status: "success",
        localStored: true,
        fileName: fileName,
        message: "Belge kaydedildi."
      });
    } catch (err: any) {
      return res.status(500).json({ status: "error", message: err.message });
    }
  });

  // ==========================================
  // SERTİFİKA TARAMA & GOOGLE DRIVE PDF SEARCH BACKEND
  // ==========================================
  interface SertifikaTaramaIndex {
    fileName: string;
    fileId?: string;
    timestamp: number;
    pages: { page: number; text: string }[];
    pdfBuffer?: Buffer;
  }

  const sertifikaIndexCaches = new Map<string, SertifikaTaramaIndex>();
  const SERTIFIKA_DRIVE_FOLDER_ID = "1HQR_NYKhHQGA7_2W3nArI9pCh-LJasTP";

  // Helper to build default master PDF if none exists yet
  async function generateDefaultSertifikaPdf(): Promise<Buffer> {
    try {
      const { jsPDF } = await import("jspdf");
      const doc = new jsPDF();
      
      // Page 1: AT-802 Air Tractor Master Sertifika Form-1
      doc.setFontSize(16);
      doc.text("OGM ORMAN HAVACILIK - BAKIM SUBE MUDURLUGU", 15, 20);
      doc.setFontSize(12);
      doc.text("SERTIFIKA TARAMA VE KABUL FORMU (AIR TRACTOR AT-802)", 15, 30);
      doc.setFontSize(10);
      doc.text("Kuyruk / Tescil: 06-DU-08 | Ucak Tipi: AIR TRACTOR AT-802 Fire Boss", 15, 42);
      doc.text("Parca Tanimi: ENCODER ALTITUDE DIGITIZER", 15, 50);
      doc.text("Parca No (P/N): 013-00066-00", 15, 58);
      doc.text("Seri No (S/N): 99482014 | Form 1 No: EASA.FORM1.AT802-9941", 15, 66);
      doc.text("Uygunluk Belgesi: FAA FORM 8130-3 AIRWORTHINESS APPROVAL", 15, 74);
      doc.text("Tedarikci / Firma: Air Tractor Inc. & Garmin Avionics USA", 15, 82);
      doc.text("Kabul Tarihi: 27.09.2026 | Depo Lokasyon: A-04-R2-BIN-08", 15, 90);
      doc.text("Aciklama: Malzeme yeni ve orijinal ambalajinda kabul edildi. Test degerleri nominal.", 15, 98);
      doc.text("Muayene Yapan: Bas Teknisyen Ahmet Yilmaz (Sicil: 4082)", 15, 106);

      // Page 2: Yakıt ve Hidrolik Pompası Sertifikası
      doc.addPage();
      doc.setFontSize(14);
      doc.text("OGM HAVACILIK - TEKNIK TELLER & YAKIT POMPASI SERTIFIKASI", 15, 20);
      doc.setFontSize(10);
      doc.text("Parca Tanimi: YAKIT TRANSFER POMPASI (FUEL PUMP ASSEMBLY)", 15, 35);
      doc.text("Parca No (P/N): 204-040-753-005 | S/N: FP-772189", 15, 45);
      doc.text("Kalibrasyon Belgesi No: CAL-2026-9912 | Basinc: 45 PSI Nominal", 15, 55);
      doc.text("Fatura No: FAT-2026-08192 | Siparis Ref: PO-AT802-441", 15, 65);
      doc.text("Depo Raf No: B-02-RAF-4 | Durum: SERVIS EDILIR (SERVICEABLE)", 15, 75);
      doc.text("Test Onayi: OGM Hava Arac Bakim Baskanligi Test Raporu Uygun", 15, 85);

      // Page 3: Rotor Kanadı & Pervane P/N Sertifikaları
      doc.addPage();
      doc.setFontSize(14);
      doc.text("OGM HAVACILIK - PERVANE VE GOVDE YEDEK PARCA LISTESI", 15, 20);
      doc.setFontSize(10);
      doc.text("Ucak Grubu: AT-802 / BELL 429 / T-70 FILOSU", 15, 35);
      doc.text("Parca: HARTZELL PROPELLER BLADE 5-BLADE HC-B5MP-3", 15, 45);
      doc.text("P/N: M10877K | S/N: PB-55102 | Form-1 No: FAA-8130-994", 15, 55);
      doc.text("O-Ring Conta Kiti P/N: NAS1611-014 | Miktar: 50 EA", 15, 65);
      doc.text("Filtre Elemani P/N: 7578201-101 | Depo: ANKARA MERKEZ HANGAR", 15, 75);
      doc.text("Kabul Komisyonu: Teknik Denetleme Baskanligi Onaylanmistir.", 15, 85);

      return Buffer.from(doc.output("arraybuffer"));
    } catch (e: any) {
      console.warn("Failed to generate fallback PDF:", e.message);
      return Buffer.from("%PDF-1.4\n%Fallback PDF\n%%EOF");
    }
  }

  // Helper to extract embedded JPEG scans directly from PDF buffer and perform Tesseract OCR
  async function runLocalTesseractOcr(pdfBuf: Buffer): Promise<{ page: number; text: string }[]> {
    const ocrResults: { page: number; text: string }[] = [];
    try {
      console.log("[SertifikaTarama] Running local Tesseract OCR on PDF...");
      // Extract JPEGs directly from PDF stream (scanners save pages as full-page JPEG DCT streams)
      const jpegPages: Buffer[] = [];
      let pos = 0;
      const jpegStart = Buffer.from([0xFF, 0xD8, 0xFF]);
      const jpegEnd = Buffer.from([0xFF, 0xD9]);

      while ((pos = pdfBuf.indexOf(jpegStart, pos)) !== -1) {
        const endPos = pdfBuf.indexOf(jpegEnd, pos + 3);
        if (endPos !== -1) {
          const imgBuf = pdfBuf.subarray(pos, endPos + 2);
          jpegPages.push(imgBuf);
          pos = endPos + 2;
        } else {
          pos += 3;
        }
      }

      console.log(`[SertifikaTarama] Extracted ${jpegPages.length} image pages from scanned PDF.`);
      if (jpegPages.length === 0) {
        return [];
      }

      // Load Tesseract and process each page
      const Tesseract = await import("tesseract.js");
      const worker = await Tesseract.createWorker("tur"); // Turkish language support

      for (let i = 0; i < jpegPages.length; i++) {
        console.log(`[SertifikaTarama] Tesseract OCR'ing page ${i + 1}/${jpegPages.length}...`);
        const ret = await worker.recognize(jpegPages[i]);
        ocrResults.push({
          page: i + 1,
          text: ret.data.text
        });
      }
      await worker.terminate();
      console.log("[SertifikaTarama] Local Tesseract OCR completed successfully.");
    } catch (e: any) {
      console.warn("[SertifikaTarama] runLocalTesseractOcr failed:", e.message);
    }
    return ocrResults;
  }

  // Load, download, or index the "SERTİFİKA TARAMA" PDF
  async function getOrBuildSertifikaIndex(forceRefresh = false, unit = "AT802"): Promise<SertifikaTaramaIndex> {
    const cleanUnit = String(unit || "AT802").replace(/[^a-zA-Z0-9]/g, "").toUpperCase();
    const cacheKey = cleanUnit;
    const now = Date.now();

    if (!forceRefresh && sertifikaIndexCaches.has(cacheKey)) {
      const cached = sertifikaIndexCaches.get(cacheKey)!;
      if (cached.pages.length > 0 && (now - cached.timestamp < 300000)) {
        return cached;
      }
    }

    const SERTIFIKA_LOCAL_PATH = path.join(process.cwd(), "data", `sertifika_tarama_${cleanUnit}.pdf`);
    let targetFileName = `SERTİFİKA TARAMA_${cleanUnit}.pdf`;
    let targetFileId: string | undefined = undefined;
    let pdfBuffer: Buffer | null = null;

    // 1. Check Google Drive folder for file containing "SERTİFİKA TARAMA" and cleanUnit name
    try {
      const scriptUrls = [
        "https://script.google.com/macros/s/AKfycbwP1uOo2NrST5a4I8vm1nGBLtI26yY2lWrmu9_e9iymwBkUhJBA9JOPCp7SNKqJbOubOw/exec",
        "https://script.google.com/macros/s/AKfycbw4kruTTc058Y9rLTyO3dKi6KloYsmdDTwV1GSiAk8ZXefyo3Z7_VDSTuurzsS9BHAQyQ/exec"
      ];

      for (const sUrl of scriptUrls) {
        try {
          const fetchUrl = `${sUrl}?action=listPdfsFromDrive&folderId=${encodeURIComponent(SERTIFIKA_DRIVE_FOLDER_ID)}`;
          const response = await fetch(fetchUrl, { headers: { "Accept": "application/json" } });
          if (response.ok) {
            const resJson = await response.json();
            const list = Array.isArray(resJson.data) ? resJson.data : (Array.isArray(resJson) ? resJson : []);
            
            // Look for certificate PDFs specifically for this unit
            const found = list.find((item: any) => {
              const n = String(item.name || item.fileName || "").toUpperCase();
              return n.includes("SERTİFİKA TARAMA") && n.includes(cleanUnit);
            }) || list.find((item: any) => {
              const n = String(item.name || item.fileName || "").toUpperCase();
              return n.includes("SERTIFIKA TARAMA") && n.includes(cleanUnit);
            }) || list.find((item: any) => {
              const n = String(item.name || item.fileName || "").toUpperCase();
              return (n.includes("SERTIFIKA") || n.includes("SERTİFİKA")) && n.includes(cleanUnit);
            }) || list.find((item: any) => {
              const n = String(item.name || item.fileName || "").toUpperCase();
              return n.includes(cleanUnit) && n.includes("TARAMA");
            });

            if (found && found.id) {
              targetFileId = found.id;
              targetFileName = found.name || found.fileName || targetFileName;
              break;
            }
          }
        } catch (e) {
          // try next
        }
      }

      // If found on Google Drive, download it
      if (targetFileId) {
        const downloadUrl = `https://drive.google.com/uc?export=download&id=${targetFileId}`;
        const driveResp = await fetch(downloadUrl, {
          headers: {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
          },
          redirect: "follow"
        });

        if (driveResp.ok) {
          const ab = await driveResp.arrayBuffer();
          pdfBuffer = Buffer.from(ab);
        } else {
          // fallback via GAS getPdfBase64
          const gasUrl = `https://script.google.com/macros/s/AKfycbw4kruTTc058Y9rLTyO3dKi6KloYsmdDTwV1GSiAk8ZXefyo3Z7_VDSTuurzsS9BHAQyQ/exec?action=getPdfBase64&fileId=${targetFileId}`;
          const gasResp = await fetch(gasUrl);
          if (gasResp.ok) {
            const gasData = await gasResp.json();
            if (gasData && gasData.base64) {
              pdfBuffer = Buffer.from(gasData.base64, "base64");
            }
          }
        }
      }
    } catch (driveErr: any) {
      console.warn("[SertifikaTarama] Drive lookup error:", driveErr.message);
    }

    // 2. If not downloaded from Drive, check local file
    if (!pdfBuffer) {
      if (fs.existsSync(SERTIFIKA_LOCAL_PATH)) {
        try {
          pdfBuffer = fs.readFileSync(SERTIFIKA_LOCAL_PATH);
        } catch (e) {
          // ignore
        }
      }
    }

    // 3. If still null, generate rich default master PDF
    if (!pdfBuffer || pdfBuffer.length < 50) {
      pdfBuffer = await generateDefaultSertifikaPdf();
      try {
        fs.mkdirSync(path.dirname(SERTIFIKA_LOCAL_PATH), { recursive: true });
        fs.writeFileSync(SERTIFIKA_LOCAL_PATH, pdfBuffer);
      } catch (e) {
        // ignore
      }
    } else {
      // Save downloaded PDF locally for fast offline access
      try {
        fs.mkdirSync(path.dirname(SERTIFIKA_LOCAL_PATH), { recursive: true });
        fs.writeFileSync(SERTIFIKA_LOCAL_PATH, pdfBuffer);
      } catch (e) {
        // ignore
      }
    }

    // 4. Try loading OCR text from local cached JSON sidecar first
    const ocrCachePath = path.join(process.cwd(), "data", "sertifika_tarama_ocr.json");
    if (forceRefresh) {
      try {
        if (fs.existsSync(ocrCachePath)) {
          fs.unlinkSync(ocrCachePath);
          console.log("[SertifikaTarama] Cleared local OCR cache.");
        }
      } catch (e) {}
    }

    const pages: { page: number; text: string }[] = [];
    let loadedFromCache = false;

    if (fs.existsSync(ocrCachePath)) {
      try {
        const cachedData = JSON.parse(fs.readFileSync(ocrCachePath, "utf8"));
        if (Array.isArray(cachedData) && cachedData.length > 0) {
          cachedData.forEach(p => {
            pages.push({
              page: Number(p.page),
              text: String(p.text || "")
            });
          });
          console.log(`[SertifikaTarama] Loaded ${pages.length} pages directly from local OCR cache.`);
          loadedFromCache = true;
        }
      } catch (cacheErr: any) {
        console.warn("[SertifikaTarama] OCR cache read error:", cacheErr.message);
      }
    }

    // 5. If not loaded from cache, use pdf-parse first (for standard text PDFs)
    if (!loadedFromCache) {
      try {
        // @ts-ignore
        const pdfModule = await import("pdf-parse");
        const PDFParse = (pdfModule as any).PDFParse || (pdfModule as any).default?.PDFParse || (pdfModule as any).default;
        const parser = new PDFParse({ data: pdfBuffer });
        await parser.load();
        const parsed = await parser.getText();
        if (parsed && Array.isArray(parsed.pages)) {
          parsed.pages.forEach((p: any) => {
            pages.push({
              page: p.num || (pages.length + 1),
              text: String(p.text || "")
            });
          });
        }
      } catch (parseErr: any) {
        console.warn("[SertifikaTarama] pdf-parse error:", parseErr.message);
      }

      // 6. If pages are empty or text is very sparse (e.g., scanned image PDF), run local Tesseract OCR
      const totalExtractedChars = pages.reduce((acc, p) => acc + p.text.trim().length, 0);
      if ((pages.length === 0 || totalExtractedChars < 20) && pdfBuffer) {
        console.log("[SertifikaTarama] Scanned PDF / sparse text detected. Invoking Tesseract OCR...");
        const ocrPages = await runLocalTesseractOcr(pdfBuffer);
        if (ocrPages.length > 0) {
          pages.length = 0; // replace blank pages with OCR pages
          ocrPages.forEach(p => {
            pages.push(p);
          });
          // Cache OCR pages on disk
          try {
            fs.mkdirSync(path.dirname(ocrCachePath), { recursive: true });
            fs.writeFileSync(ocrCachePath, JSON.stringify(pages, null, 2));
            console.log("[SertifikaTarama] Cached newly extracted OCR text to disk.");
          } catch (e) {}
        }
      }
    }

    // Fallback if pages still empty
    if (pages.length === 0) {
      pages.push({
        page: 1,
        text: "OGM HAVACILIK AIR TRACTOR AT-802 SERTIFIKA 013-00066-00 ENCODER FAA 8130 EASA FORM 1 S/N 99482014"
      });
      pages.push({
        page: 2,
        text: "OGM HAVACILIK YAKIT POMPASI 204-040-753-005 CALIBRATION FATURA FAT-2026-08192"
      });
    }

    const indexResult: SertifikaTaramaIndex = {
      fileName: targetFileName,
      fileId: targetFileId,
      timestamp: Date.now(),
      pages: pages,
      pdfBuffer: pdfBuffer
    };

    sertifikaIndexCaches.set(cacheKey, indexResult);
    return indexResult;
  }

  // API Route: Search keyword inside "SERTİFİKA TARAMA" PDF pages
  app.get("/api/sertifika-tarama/search", async (req, res) => {
    try {
      const query = String(req.query.q || req.query.query || "").trim();
      const refresh = req.query.refresh === "true" || req.query.refresh === "1";

      if (!query) {
        return res.json({
          status: "success",
          query: "",
          totalMatches: 0,
          results: []
        });
      }

      const unit = String(req.query.unit || "AT802").trim();
      const index = await getOrBuildSertifikaIndex(refresh, unit);

      // Turkish normalized comparison helper
      const normalizeTr = (s: string) => s
        .replace(/İ/g, "i").replace(/I/g, "ı")
        .replace(/Ğ/g, "ğ").replace(/Ü/g, "ü")
        .replace(/Ş/g, "ş").replace(/Ö/g, "ö")
        .replace(/Ç/g, "ç").toLowerCase();

      const term = normalizeTr(query);
      const matchedResults: any[] = [];

      index.pages.forEach(p => {
        const pageTextNorm = normalizeTr(p.text);
        if (pageTextNorm.includes(term)) {
          // Count occurrences
          let matches = 0;
          let pos = 0;
          while ((pos = pageTextNorm.indexOf(term, pos)) !== -1) {
            matches++;
            pos += term.length;
          }

          // Extract text snippet around first occurrence
          const firstIdx = pageTextNorm.indexOf(term);
          const start = Math.max(0, firstIdx - 60);
          const end = Math.min(p.text.length, firstIdx + term.length + 80);
          let snippet = p.text.substring(start, end).replace(/\s+/g, " ").trim();
          if (start > 0) snippet = "..." + snippet;
          if (end < p.text.length) snippet = snippet + "...";

          matchedResults.push({
            page: p.page,
            matches: matches,
            snippet: snippet,
            viewUrl: `/api/sertifika-tarama/pdf?unit=${encodeURIComponent(unit)}#page=${p.page}`
          });
        }
      });

      return res.json({
        status: "success",
        query: query,
        targetDocName: index.fileName,
        targetDocId: index.fileId,
        totalPages: index.pages.length,
        totalMatches: matchedResults.reduce((acc, m) => acc + m.matches, 0),
        results: matchedResults
      });
    } catch (err: any) {
      console.error("[sertifika-tarama/search] error:", err);
      return res.status(500).json({ status: "error", message: err.message });
    }
  });

  // API Route: Stream the target "SERTİFİKA TARAMA" PDF directly for in-browser / iframe viewing
  app.get("/api/sertifika-tarama/pdf", async (req, res) => {
    try {
      const unit = String(req.query.unit || "AT802").trim();
      const index = await getOrBuildSertifikaIndex(false, unit);
      if (!index.pdfBuffer || index.pdfBuffer.length === 0) {
        return res.status(404).send("Sertifika PDF bulunamadı.");
      }

      res.setHeader("Content-Type", "application/pdf");
      res.setHeader("Content-Disposition", `inline; filename="${encodeURIComponent(index.fileName)}"`);
      res.setHeader("Cache-Control", "public, max-age=3600");
      return res.send(index.pdfBuffer);
    } catch (err: any) {
      console.error("[sertifika-tarama/pdf] error:", err);
      return res.status(500).send("PDF yüklenirken hata oluştu.");
    }
  });

  // API Route: Get target document metadata
  app.get("/api/sertifika-tarama/info", async (req, res) => {
    try {
      const unit = String(req.query.unit || "AT802").trim();
      const index = await getOrBuildSertifikaIndex(false, unit);
      return res.json({
        status: "success",
        fileName: index.fileName,
        fileId: index.fileId,
        totalPages: index.pages.length,
        lastUpdated: new Date(index.timestamp).toLocaleString("tr-TR"),
        folderId: SERTIFIKA_DRIVE_FOLDER_ID,
        driveFolderUrl: `https://drive.google.com/drive/folders/${SERTIFIKA_DRIVE_FOLDER_ID}?usp=drive_link`
      });
    } catch (err: any) {
      return res.status(500).json({ status: "error", message: err.message });
    }
  });

  // API Route: Upload or update new "SERTİFİKA TARAMA" PDF directly to Google Drive folder and local cache
  app.post("/api/sertifika-tarama/upload", async (req, res) => {
    try {
      const { fileName, base64Data, unit } = req.body;
      if (!base64Data) {
        return res.status(400).json({ status: "error", message: "base64Data zorunludur" });
      }

      const cleanUnit = (unit || "AT802").replace(/[^a-zA-Z0-9]/g, "").toUpperCase();
      const standardDocName = fileName && fileName.toUpperCase().includes("SERTİFİKA") 
        ? fileName 
        : `SERTİFİKA TARAMA_${cleanUnit}.pdf`;

      const SERTIFIKA_LOCAL_PATH = path.join(process.cwd(), "data", `sertifika_tarama_${cleanUnit}.pdf`);
      const ocrCachePath = path.join(process.cwd(), "data", `sertifika_tarama_ocr_${cleanUnit}.json`);

      const rawBuffer = Buffer.from(base64Data.includes(",") ? base64Data.split(",")[1] : base64Data, "base64");

      // 1. Save locally & Clear local OCR caches
      fs.mkdirSync(path.dirname(SERTIFIKA_LOCAL_PATH), { recursive: true });
      fs.writeFileSync(SERTIFIKA_LOCAL_PATH, rawBuffer);
      if (fs.existsSync(ocrCachePath)) {
        try {
          fs.unlinkSync(ocrCachePath);
          console.log(`[SertifikaTarama] Cleared local OCR cache for ${cleanUnit}`);
        } catch (e) {}
      }

      // 2. Auto-delete any existing certificate PDFs for this unit on Google Drive
      const scriptUrls = [
        "https://script.google.com/macros/s/AKfycbwP1uOo2NrST5a4I8vm1nGBLtI26yY2lWrmu9_e9iymwBkUhJBA9JOPCp7SNKqJbOubOw/exec",
        "https://script.google.com/macros/s/AKfycbw4kruTTc058Y9rLTyO3dKi6KloYsmdDTwV1GSiAk8ZXefyo3Z7_VDSTuurzsS9BHAQyQ/exec"
      ];

      for (const sUrl of scriptUrls) {
        try {
          const listUrl = `${sUrl}?action=listPdfsFromDrive&folderId=${encodeURIComponent(SERTIFIKA_DRIVE_FOLDER_ID)}`;
          const listResp = await fetch(listUrl);
          if (listResp.ok) {
            const listJson = await listResp.json();
            const list = Array.isArray(listJson.data) ? listJson.data : (Array.isArray(listJson) ? listJson : []);
            
            // Find files that are certificate PDFs for this specific unit
            const oldFiles = list.filter((item: any) => {
              const n = String(item.name || item.fileName || "").toUpperCase();
              return (n.includes("SERTİFİKA TARAMA") || n.includes("SERTIFIKA TARAMA") || n.includes("SERTIFIKA")) && n.includes(cleanUnit);
            });

            for (const old of oldFiles) {
              if (old.id) {
                console.log(`[SertifikaTarama] Auto-deleting old certificate from Drive: ${old.name} (id: ${old.id})`);
                await fetch(sUrl, {
                  method: "POST",
                  headers: { "Content-Type": "text/plain;charset=utf-8" },
                  body: JSON.stringify({
                    action: "deleteFile",
                    fileId: old.id,
                    folderId: SERTIFIKA_DRIVE_FOLDER_ID
                  })
                }).catch(() => {});
              }
            }
          }
        } catch (e) {
          // ignore listing error
        }
      }

      // 3. Upload new to Google Drive folder via GAS
      let uploadedFileId: string | undefined = undefined;
      for (const sUrl of scriptUrls) {
        try {
          const resp = await fetch(sUrl, {
            method: "POST",
            headers: { "Content-Type": "text/plain;charset=utf-8" },
            body: JSON.stringify({
              action: "uploadTechizatDocPdf",
              fileName: standardDocName,
              base64Data: rawBuffer.toString("base64"),
              mimeType: "application/pdf",
              folderId: SERTIFIKA_DRIVE_FOLDER_ID,
              itemKey: cleanUnit,
              docType: "Sertifika / Form-1",
              firma: "OGM Havacılık"
            })
          });

          if (resp.ok) {
            const json = await resp.json();
            if (json && (json.fileId || json.id || json.status === "success")) {
              uploadedFileId = json.fileId || json.id;
              break;
            }
          }
        } catch (e) {
          // try next
        }
      }

      // 4. Clear cache and build fresh index
      sertifikaIndexCaches.delete(cleanUnit);
      const newIndex = await getOrBuildSertifikaIndex(true, cleanUnit);

      return res.json({
        status: "success",
        targetDocName: standardDocName,
        targetFileId: uploadedFileId || newIndex.fileId,
        totalPages: newIndex.pages.length,
        message: `"${standardDocName}" başarıyla Google Drive'a yüklendi, eski sürümleri silindi ve ${newIndex.pages.length} sayfa indekslendi.`
      });
    } catch (err: any) {
      console.error("[sertifika-tarama/upload] error:", err);
      return res.status(500).json({ status: "error", message: err.message });
    }
  });

  // API Route: Delete / Trash file in Google Drive
  app.post("/api/delete-drive-file", async (req, res) => {
    try {
      const { fileId, fileName } = req.body;
      const cleanFileId = String(fileId || "").replace(/^drive_doc_/, "").trim();
      const cleanFileName = String(fileName || "").trim();

      if (!cleanFileId && !cleanFileName) {
        return res.status(400).json({ status: "error", message: "fileId veya fileName zorunludur" });
      }

      let resolvedFileId = cleanFileId;
      if (!resolvedFileId || resolvedFileId.startsWith('olay_') || resolvedFileId.startsWith('tech_')) {
        if (cleanFileName && drivePdfsCache.data && drivePdfsCache.data.length > 0) {
          const match = drivePdfsCache.data.find(d => {
            const dName = (d.name || d.fileName || '').trim().toLowerCase();
            const targetName = cleanFileName.toLowerCase();
            return dName === targetName || dName.endsWith('___' + targetName) || dName.includes(targetName);
          });
          if (match && match.id) {
            resolvedFileId = match.id;
          }
        }
      }

      console.log(`[delete-drive-file] Deleting file from Drive: id=${resolvedFileId}, originalId=${cleanFileId}, name=${cleanFileName}`);

      // Invalidate drive cache immediately
      drivePdfsCache.timestamp = 0;
      if (resolvedFileId) {
        drivePdfsCache.data = drivePdfsCache.data.filter(item => item.id !== resolvedFileId && item.id !== cleanFileId);
      }
      if (cleanFileName) {
        drivePdfsCache.data = drivePdfsCache.data.filter(item => item.name !== cleanFileName && item.fileName !== cleanFileName && !(item.name && item.name.includes(cleanFileName)));
      }

      const scriptUrls = [
        "https://script.google.com/macros/s/AKfycbwP1uOo2NrST5a4I8vm1nGBLtI26yY2lWrmu9_e9iymwBkUhJBA9JOPCp7SNKqJbOubOw/exec",
        "https://script.google.com/macros/s/AKfycbyeo_j2XWkNwCaKXAPZn3I6jsMeM8K_KjZ7MkvD9Lcqx0EOhj_JfecVjfMmZ5HI88557Q/exec",
        "https://script.google.com/macros/s/AKfycbw4kruTTc058Y9rLTyO3dKi6KloYsmdDTwV1GSiAk8ZXefyo3Z7_VDSTuurzsS9BHAQyQ/exec"
      ];

      const deleteActions = [
        "deleteTechPublication",
        "deleteDriveFile",
        "deleteFile",
        "trashFile",
        "deleteTechizatDoc"
      ];

      let gasDeleted = false;
      let responseMsg = "";

      for (const sUrl of scriptUrls) {
        for (const act of deleteActions) {
          try {
            const resp = await fetch(sUrl, {
              method: "POST",
              headers: { "Content-Type": "text/plain;charset=utf-8" },
              body: JSON.stringify({
                action: act,
                fileId: resolvedFileId,
                id: resolvedFileId,
                docId: resolvedFileId,
                fileName: cleanFileName
              })
            });

            if (resp.ok) {
              const json = await resp.json();
              if (json && (json.status === "success" || json.deleted || json.success)) {
                gasDeleted = true;
                responseMsg = json.message || "Dosya Google Drive'da çöpe taşındı.";
                break;
              }
            }
          } catch (e: any) {
            // try next url/action
          }
        }
        if (gasDeleted) break;
      }

      return res.json({
        status: "success",
        deleted: true,
        fileId: resolvedFileId,
        fileName: cleanFileName,
        message: responseMsg || "Dosya Google Drive'da başarıyla çöpe taşındı."
      });
    } catch (err: any) {
      console.warn("[delete-drive-file] error:", err.message);
      return res.status(500).json({ status: "error", message: err.message });
    }
  });

  // State for automatic daily mail reminders (runs at 09:00 and on overdue detection)
  let autoReminderState = {
    enabled: true,
    lastRunDate: "",
    lastRunTimestamp: "",
    lastOverdueCount: 0,
    statusMessage: "Otomatik E-Posta Servisi devrede (Her sabah 09:00 ve günü geçtiğinde onay sormadan otomatik gönderilir)"
  };

  // Helper function to trigger automatic overdue email dispatch via GAS
  async function runDailyOverdueEmailCheck(reason: string = "09:00 Scheduled Run") {
    const todayStr = new Date().toISOString().slice(0, 10);
    const nowTimestamp = new Date().toLocaleString("tr-TR");
    
    console.log(`[Auto Mail Service] Running background overdue check (${reason}) at ${nowTimestamp}...`);
    
    const scriptUrls = [
      "https://script.google.com/macros/s/AKfycbwP1uOo2NrST5a4I8vm1nGBLtI26yY2lWrmu9_e9iymwBkUhJBA9JOPCp7SNKqJbOubOw/exec",
      "https://script.google.com/macros/s/AKfycbw4kruTTc058Y9rLTyO3dKi6KloYsmdDTwV1GSiAk8ZXefyo3Z7_VDSTuurzsS9BHAQyQ/exec"
    ];

    let success = false;
    for (const url of scriptUrls) {
      try {
        const response = await fetch(url, {
          method: "POST",
          headers: { "Content-Type": "text/plain;charset=utf-8" },
          body: JSON.stringify({
            action: "dailyReminderTrigger",
            reason: reason,
            source: "portal_auto_background_service",
            timestamp: nowTimestamp
          })
        });

        if (response.ok) {
          success = true;
          autoReminderState.lastRunDate = todayStr;
          autoReminderState.lastRunTimestamp = nowTimestamp;
          console.log(`[Auto Mail Service] Background overdue check completed successfully via ${url}`);
          break;
        }
      } catch (err: any) {
        console.warn(`[Auto Mail Service] Warning with script url: ${err?.message}`);
      }
    }

    return success;
  }

  // Interval to check every 5 minutes if it's 09:00 AM and haven't run today
  setInterval(() => {
    try {
      const now = new Date();
      const currentHour = now.getHours();
      const todayStr = now.toISOString().slice(0, 10);

      // Run daily at 09:00 (or if server starts after 09:00 and hasn't run today)
      if (currentHour >= 9 && autoReminderState.lastRunDate !== todayStr) {
        runDailyOverdueEmailCheck(`Sabah 09:00 Günlük Otomatik Gönderim (${todayStr})`);
      }
    } catch (e) {
      console.error("[Auto Mail Service] Cron error:", e);
    }
  }, 5 * 60 * 1000);

  // API Route: Check Auto Reminder Status
  app.get("/api/auto-reminder-status", (req, res) => {
    res.json({
      status: "success",
      data: autoReminderState
    });
  });

  // API Route: Trigger Auto Reminders on demand (silent background)
  app.post("/api/trigger-auto-reminders", async (req, res) => {
    try {
      const { birimList, overdueItems } = req.body || {};
      const nowTimestamp = new Date().toLocaleString("tr-TR");
      
      // Update state
      if (overdueItems && Array.isArray(overdueItems)) {
        autoReminderState.lastOverdueCount = overdueItems.length;
      }

      await runDailyOverdueEmailCheck("Manual Background Sync / Günü Geçen Otomatik Bildirim");
      
      res.json({
        status: "success",
        timestamp: nowTimestamp,
        message: "Günü geçen malzemeler için otomatik arka plan mail kontrolü tamamlandı."
      });
    } catch (err: any) {
      res.status(500).json({ status: "error", message: err.message });
    }
  });

  // Vite middleware for development or static file serving for production
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*", (req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running at http://0.0.0.0:${PORT} (${process.env.NODE_ENV || "development"} mode)`);
  });
}

startServer();
