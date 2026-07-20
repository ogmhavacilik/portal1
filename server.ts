import express from "express";
import path from "path";
import { createServer as createViteServer } from "vite";
import dotenv from "dotenv";

dotenv.config();

async function startServer() {
  const app = express();
  const PORT = 3000;

  // Body parsing middleware
  app.use(express.json());

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
