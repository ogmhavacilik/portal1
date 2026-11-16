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

// Taskline EBYS Proxy
app.get('/api/taskline-ebys', async (req, res) => {
  const scriptUrl = req.query.scriptUrl as string;
  const targetUrl = scriptUrl || 'https://script.google.com/macros/s/AKfycbxuivDglCOCOl32raOdzjV_gjDk6k7i70pepzYp8ak9tcc1D8t5L9gJNI2YxYXZi7wa/exec';
  const finalUrl = targetUrl.replace(/\/dev$/, '/exec');
  const spreadsheetId = (req.query.spreadsheetId as string) || '1L05588TdYZmH401Lvn4_yr4zwiw2pW4EJ8dIyl-UTVQ';
  
  let fetchedData: any = null;
  try {
    const response = await fetch(finalUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ action: 'getDemands', spreadsheetId })
    });
    if (response.ok) {
      const json = await response.json();
      if (json && json.status !== 'error') {
        fetchedData = json;
      }
    }
  } catch (err: any) {}

  if (!fetchedData) {
    return res.status(500).json({ status: 'error', message: 'Canlı EBYS listesi çekilemedi.' });
  }
  return res.json(fetchedData);
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
