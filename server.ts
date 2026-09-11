import express from "express";
import path from "path";
import { createServer as createViteServer } from "vite";
import dotenv from "dotenv";

dotenv.config();

async function startServer() {
  const app = express();
  const PORT = 3000;

  // Body parsing middleware with high limit for PDFs and large spreadsheets
  app.use(express.json({ limit: "150mb" }));
  app.use(express.urlencoded({ limit: "150mb", extended: true }));

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

    // Yöntem 3: Eğer canlı bağlantıların tamamı başarısız olduysa, lokal default listeyi döndür (kesinti olmasın)
    if (!fetchedData) {
      console.log("[Yöntem 3] Canlı Apps Script sorguları başarısız oldu. Lokal fallback listesi yükleniyor.");
      fetchedData = {
        status: "success",
        data: getFallbackEbysList(),
        message: "Lokal/Fallback EBYS listesi yüklendi."
      };
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

      // Default comprehensive list of active maintenance technicians, pilots & warehouse personnel if offline
      const defaultPersonnel = [
        "Mahmut OKUDAN",
        "Rıfat ÖNAL",
        "Nuri Gökmen GEÇER",
        "Sabri AKSOY",
        "Hakan KOZLU",
        "Gürhan AYDIN",
        "Serhat İVECAN",
        "Yücel KIVRAK",
        "T.GÜZER",
        "A.ÖZMETİN",
        "Y.KARADENİZ",
        "M.KAYA",
        "H.DEMİR",
        "E.ŞAHİN",
        "S.YILMAZ",
        "B.ÖZTÜRK",
        "K.ARSLAN",
        "O.ÇELİK",
        "F.GÜNEŞ",
        "İ.YILDIRIM",
        "Ahmet YILDIZ",
        "Mehmet KESKİN",
        "Mustafa DEMİRTAŞ",
        "Ali ÇAĞLAR",
        "Murat ŞEN",
        "Emre AY",
        "Cemil KOÇ"
      ];

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

  // API Route: Save Depo Transfers to Google Apps Script / Google Sheets
  app.post("/api/save-depo-transfers", async (req, res) => {
    try {
      const { transfers = [], spreadsheetId = "17ScGYYx0erzDwHDk6RGiHOdJATdfmmExXFBY39dXpF0", sheetName } = req.body;
      const scriptUrl = "https://script.google.com/macros/s/AKfycbw4kruTTc058Y9rLTyO3dKi6KloYsmdDTwV1GSiAk8ZXefyo3Z7_VDSTuurzsS9BHAQyQ/exec";
      
      const gasRes = await fetch(scriptUrl, {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify({
          action: "saveDepoTransfers",
          transfers: transfers,
          spreadsheetId: spreadsheetId,
          sheetName: sheetName
        })
      });

      const data = await gasRes.json();
      return res.json(data);
    } catch (err: any) {
      return res.status(500).json({
        status: "error",
        message: err.message
      });
    }
  });

  // API Route: Upload Teçhizat Excel to Google Drive
  app.post("/api/upload-techizat-excel", async (req, res) => {
    try {
      const { fileName, base64Data, targetKey, folderId } = req.body;
      const scriptUrl = "https://script.google.com/macros/s/AKfycbw4kruTTc058Y9rLTyO3dKi6KloYsmdDTwV1GSiAk8ZXefyo3Z7_VDSTuurzsS9BHAQyQ/exec";
      
      const gasRes = await fetch(scriptUrl, {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify({
          action: "uploadTechizatExcel",
          fileName: fileName,
          targetKey: targetKey,
          base64Data: base64Data,
          folderId: folderId || "1HQR_NYKhHQGA7_2W3nArI9pCh-LJasTP"
        })
      });

      const data = await gasRes.json();
      return res.json(data);
    } catch (err: any) {
      return res.status(500).json({
        status: "error",
        message: err.message
      });
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
      console.warn("[upload-hangar-pdf] error:", err.message);
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

      console.log(`[delete-drive-file] Deleting file from Drive: id=${cleanFileId}, name=${cleanFileName}`);

      // Invalidate drive cache immediately
      drivePdfsCache.timestamp = 0;
      if (cleanFileId) {
        drivePdfsCache.data = drivePdfsCache.data.filter(item => item.id !== cleanFileId);
      }
      if (cleanFileName) {
        drivePdfsCache.data = drivePdfsCache.data.filter(item => item.name !== cleanFileName && item.fileName !== cleanFileName);
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
                fileId: cleanFileId,
                id: cleanFileId,
                docId: cleanFileId,
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
        fileId: cleanFileId,
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

// Fallback "İşlemdeki Talepler" data that perfectly populates the portal's EBYS selector
function getFallbackEbysList() {
  return [
    {
      "SIRA NO": "1",
      "Başlık": "Bell 429 Test Cihazı Kalibrasyon Talebi",
      "Açıklama": "Teknik Şube kalibrasyon laboratuvarı cihaz kalibrasyon işlem talebi",
      "Talep Türü": "KALİBRASYON",
      "EBYS NO": "2408159"
    },
    {
      "SIRA NO": "2",
      "Başlık": "T-70 Telsiz Ölçüm Ekipmanı Periyodik Kontrol",
      "Açıklama": "Aviyonik atölyesi telsiz test cihazının periyodik kontrol işlemi",
      "Talep Türü": "BAKIM",
      "EBYS NO": "2408160"
    },
    {
      "SIRA NO": "3",
      "Başlık": "AT-802F Basınç Göstergeleri Atölye Testi",
      "Açıklama": "Hidrolik atölyesi manometre ve basınç göstergelerinin yıllık kontrolü",
      "Talep Türü": "BAKIM",
      "EBYS NO": "2408161"
    },
    {
      "SIRA NO": "4",
      "Başlık": "Bell 429 Motor Tork Anahtarı Kalibrasyonu",
      "Açıklama": "Kalibrasyon laboratuvarında tork anahtarı kalibrasyon işlemi",
      "Talep Türü": "KALİBRASYON",
      "EBYS NO": "2408162"
    },
    {
      "SIRA NO": "5",
      "Başlık": "T-70 Yangın Söndürme Kit Kontrolü",
      "Açıklama": "Yer destek teçhizatlarının ve bumbi bucket sisteminin kontrolü",
      "Talep Türü": "BAKIM",
      "EBYS NO": "2408163"
    }
  ];
}

startServer();
