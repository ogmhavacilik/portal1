import serverless from 'serverless-http';
import express from 'express';
import path from 'path';
import fs from 'fs';
import * as XLSX from 'xlsx';
import dotenv from 'dotenv';
import AdmZip from 'adm-zip';

dotenv.config();

const app = express();

app.use(express.json({ limit: '150mb' }));
app.use(express.urlencoded({ limit: '150mb', extended: true }));

app.get('/api/health', (req, res) => {
  res.json({ status: 'ok', message: 'Netlify serverless function is running' });
});

// Resilient Excel from Drive Route for Netlify
app.post('/api/read-excel-from-drive', async (req, res) => {
  try {
    const { fileName = 'at-802_sarf_ve_parca_deposu.xlsx' } = req.body || {};
    const localPath = path.resolve(process.cwd(), `public/${fileName}`);
    if (fs.existsSync(localPath)) {
      const buf = fs.readFileSync(localPath);
      return res.json({
        status: 'success',
        source: 'local-mirror',
        base64: buf.toString('base64'),
        updated: new Date().toLocaleString('tr-TR')
      });
    }
    return res.status(404).json({ status: 'error', message: 'Dosya bulunamadı' });
  } catch (err: any) {
    return res.status(500).json({ status: 'error', message: err.message });
  }
});

// Gün Takip Live Scan Fallback
app.get('/api/gun-takip-live-scan', async (_req, res) => {
  res.json({
    status: 'success',
    timestamp: new Date().toISOString(),
    data: {
      at802: [],
      at802_yer: [],
      at802_ozel_alet: [],
      bell429: [],
      t70: [],
      t70_bumbi_backet: [],
      t70_helitak: [],
      hangar: [],
      kara_araclari: [],
      b360: [],
      c650: []
    }
  });
});

const handler = serverless(app);
export { handler };
