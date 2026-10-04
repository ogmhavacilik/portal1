import express from 'express';
import fs from 'fs';
import path from 'path';

const app = express();
app.use(express.json());

const DATA_DIR = path.join(__dirname, 'data');
if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR);

const FILES = {
  RECORDS: path.join(DATA_DIR, 'zimmetler.json'),
  PERSONNEL: path.join(DATA_DIR, 'personnel.json'),
  CAN_YELEGI: path.join(DATA_DIR, 'can_yelegi_stok.json'),
  SPARE_AIR: path.join(DATA_DIR, 'spare_air_stok.json'),
  HELMET_KIT: path.join(DATA_DIR, 'helmet_kit_stok.json'),
};

function readJson(file: string) {
  if (!fs.existsSync(file)) return [];
  try {
    return JSON.parse(fs.readFileSync(file, 'utf-8'));
  } catch { return []; }
}

function writeJson(file: string, data: any[]) {
  fs.writeFileSync(file, JSON.stringify(data, null, 2));
}

// ENDPOINTS
app.get('/api/yasam-destek/data', (req, res) => res.json(readJson(FILES.RECORDS)));
app.get('/api/yasam-destek/personnel', (req, res) => res.json(readJson(FILES.PERSONNEL)));
app.get('/api/yasam-destek/stocks', (req, res) => {
  res.json({
    canYelegi: readJson(FILES.CAN_YELEGI),
    spareAir: readJson(FILES.SPARE_AIR),
    helmetKit: readJson(FILES.HELMET_KIT),
  });
});

app.post('/api/yasam-destek/assign', (req, res) => {
  const records = readJson(FILES.RECORDS);
  const newRecord = { ...req.body, id: `yd_${Date.now()}`, bitisTarihi: '-', guncellenmeTarihi: new Date().toLocaleDateString('tr-TR') };
  records.unshift(newRecord);
  writeJson(FILES.RECORDS, records);

  // Update stocks
  const updateStock = (file: string, sn: string, pName: string) => {
    const stocks = readJson(file);
    const item = stocks.find((s: any) => s.seriNo === sn || s.seriNo1 === sn || s.seriNo2 === sn);
    if (item) {
      item.durum = "GÖREV BÖLGESİNDE";
      item.aciklamalar = pName;
      item.tarih = new Date().toLocaleDateString('tr-TR');
      writeJson(file, stocks);
    }
  };
  if(newRecord.lifeVestSn !== '-') updateStock(FILES.CAN_YELEGI, newRecord.lifeVestSn, newRecord.personelAdi);
  if(newRecord.spareAirSn !== '-') updateStock(FILES.SPARE_AIR, newRecord.spareAirSn, newRecord.personelAdi);
  if(newRecord.helmetKitSn !== '-') updateStock(FILES.HELMET_KIT, newRecord.helmetKitSn, newRecord.personelAdi);

  res.json({ status: 'success' });
});

app.post('/api/yasam-destek/finish', (req, res) => {
  const { recordId, bitisTarihi, action, notlar } = req.body;
  const records = readJson(FILES.RECORDS);
  const idx = records.findIndex((r: any) => r.id === recordId);
  if (idx !== -1) {
    records[idx].bitisTarihi = bitisTarihi;
    records[idx].durum = action === 'devret' ? 'DEVREDİLDİ' : 'DEPODA';
    records[idx].notlar = notlar;
    writeJson(FILES.RECORDS, records);

    // Update stocks to DEPODA
    const updateStock = (file: string, sn: string) => {
      const stocks = readJson(file);
      const item = stocks.find((s: any) => s.seriNo === sn || s.seriNo1 === sn || s.seriNo2 === sn);
      if (item) {
        item.durum = "DEPODA";
        item.aciklamalar = "-";
        writeJson(file, stocks);
      }
    };
    if(records[idx].lifeVestSn !== '-') updateStock(FILES.CAN_YELEGI, records[idx].lifeVestSn);
    if(records[idx].spareAirSn !== '-') updateStock(FILES.SPARE_AIR, records[idx].spareAirSn);
    if(records[idx].helmetKitSn !== '-') updateStock(FILES.HELMET_KIT, records[idx].helmetKitSn);
  }
  res.json({ status: 'success' });
});

app.listen(3000, () => console.log('Sunucu 3000 portunda çalışıyor...'));
