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
    try {
      const targetUrl = "https://script.google.com/macros/s/AKfycbxzwmR6Bd2QOMuTzN984Z_jrpbzPeN88OWnomxhmI8e0TFL1iqyS0ZZmHh1Ln2Cg-iN/exec";
      
      console.log("Fetching live data from GAS Web App via POST (action: getDemands)...");
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), 12000); // 12 second timeout
      
      const response = await fetch(targetUrl, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({ action: "getDemands" }),
        signal: controller.signal
      });
      clearTimeout(timeoutId);
      
      if (!response.ok) {
        throw new Error(`GAS Web App sunucu durum kodu döndürdü: ${response.status}`);
      }
      
      const text = await response.text();
      
      // If it's an HTML error page (like 'index not found' or general Google Web App error)
      if (text.includes("No HTML file named index was found") || text.includes("HTML-bestand") || text.trim().startsWith("<!DOCTYPE") || text.trim().startsWith("<html")) {
        console.warn("GAS returned an HTML error page.");
        return res.json({ 
          status: "error", 
          message: "Google Apps Script hatası: E-Tablo dosyası açılamadı veya yetki sorunu var. Lütfen Google Apps Script projesindeki e-tablo kimliğinin (SPREADSHEET_ID) doğru olduğunu ve Web Uygulaması ayarlarında yetkinin 'Anyone' (Herkes) olarak seçildiğini kontrol edin." 
        });
      }
      
      try {
        const json = JSON.parse(text);
        console.log("Successfully fetched live data from GAS Web App.");
        return res.json(json);
      } catch (e) {
        console.warn("GAS response was not valid JSON.");
        return res.json({ 
          status: "error", 
          message: "Google Apps Script JSON yerine geçersiz veri döndürdü." 
        });
      }
    } catch (err: any) {
      console.error("Taskline EBYS proxy error:", err.message);
      return res.json({ 
        status: "error", 
        message: `Taskline EBYS bağlantı hatası: ${err.message}` 
      });
    }
  });

  // API Route: Taskline Submit Proxy
  app.post("/api/taskline-submit", async (req, res) => {
    try {
      const { ebysNo, data, scriptUrl } = req.body;
      
      const targetUrl = scriptUrl || "https://script.google.com/macros/s/AKfycbsqDITjd3ZddAvRKQhpNF3ymQncNfzgC0IHCvm-rUi/exec";
      // Ensure we target the production /exec deployment if the user accidentally copied the /dev URL
      const finalUrl = targetUrl.replace(/\/dev$/, "/exec");
      
      console.log(`[Proxy Taskline Submit] Posting to ${finalUrl} with EBYS: ${ebysNo}`);
      
      const response = await fetch(finalUrl, {
        method: "POST",
        headers: {
          "Content-Type": "text/plain;charset=utf-8"
        },
        body: JSON.stringify({
          action: "appendEbysTable",
          ebysNo: ebysNo,
          data: data
        })
      });

      if (!response.ok) {
        throw new Error(`Google Apps Script returned status: ${response.status}`);
      }

      const responseText = await response.text();
      console.log(`[Proxy Taskline Submit] Response from GAS:`, responseText.substring(0, 500));

      return res.json({
        status: "success",
        message: "Data successfully sent to Taskline Google Sheet.",
        gasStatus: response.status,
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
