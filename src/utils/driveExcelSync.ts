import * as XLSX from 'xlsx';

export const GOOGLE_SCRIPT_URL = "https://script.google.com/macros/s/AKfycby2TM-nlURR5iPR4s4I6BpE8hbor93Jin9g014k3XPaQ0rYtS2MWHwtlnlAoph8Y3mZ/exec";
export const DRIVE_TECHIZAT_FOLDER_ID = "1HQR_NYKhHQGA7_2W3nArI9pCh-LJasTP";

export interface DriveExcelFile {
  id: string;
  name: string;
  size?: number;
  lastUpdated?: string;
  downloadUrl?: string;
}

// Tarih verilerini GG.AA.YYYY formatına kesin olarak temizleyen ve Excel seri numaralarını (46234 vb) dönüştüren fonksiyon
export const cleanAndFormatDateString = (val: any): string => {
  if (val === undefined || val === null) return '';
  
  // Date object direct handling
  if (val instanceof Date) {
    if (!isNaN(val.getTime())) {
      const d = String(val.getDate()).padStart(2, '0');
      const m = String(val.getMonth() + 1).padStart(2, '0');
      const y = String(val.getFullYear());
      return `${d}.${m}.${y}`;
    }
  }

  let str = String(val).trim();
  if (!str || str === '-' || str === '--' || str.toLowerCase() === 'null' || str.toLowerCase() === 'undefined') return str || '';

  // 1. Excel seri numarası kontrolü (Örn: 45234, 46234, 46965 vb.)
  // Kullanıcı kuralı: "TARİH VERİLERİNİ RAKAM GİBİ GÖRÜYOR HATADIR."
  if (/^\d{5}(\.\d+)?$/.test(str)) {
    const serial = parseFloat(str);
    if (serial >= 20000 && serial <= 90000) {
      const utcDays = Math.floor(serial - 25569);
      const dateObj = new Date(utcDays * 86400 * 1000);
      if (!isNaN(dateObj.getTime())) {
        const d = String(dateObj.getUTCDate()).padStart(2, '0');
        const m = String(dateObj.getUTCMonth() + 1).padStart(2, '0');
        const y = String(dateObj.getUTCFullYear());
        return `${d}.${m}.${y}`;
      }
    }
  }

  // 2. Yapışık gelen tarihleri ayıkla
  str = str.replace(/(\d{1,2}[\.\/-]\d{1,2}[\.\/-]\d{2,4})(\d{1,2}[\.\/-]\d{1,2}[\.\/-]\d{2,4})/g, '$1\n$2');

  // 3. Çok satırlı ise her bir satırı formatla
  const lines = str.split(/[\r\n;]+/).map(s => s.trim()).filter(Boolean);
  if (lines.length > 1) {
    return lines.map(l => cleanAndFormatDateString(l)).join('\n');
  }

  // 4. ISO Date string veya datetime (Örn: 2026-11-24T00:00:00.000Z veya 2026-11-24 14:30)
  if (str.includes('T') || /^\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}/.test(str)) {
    const parsed = new Date(str);
    if (!isNaN(parsed.getTime())) {
      const d = String(parsed.getDate()).padStart(2, '0');
      const m = String(parsed.getMonth() + 1).padStart(2, '0');
      const y = String(parsed.getFullYear());
      return `${d}.${m}.${y}`;
    }
  }

  // 5. Türkçe ay isimleri (Örn: 31 Temmuz 2026, 15 OCAK 2025)
  const turkishMonths: { [key: string]: string } = {
    'OCAK': '01', 'OCA': '01', 'JAN': '01', 'JANUARY': '01',
    'ŞUBAT': '02', 'SUBAT': '02', 'ŞUB': '02', 'SUB': '02', 'FEB': '02', 'FEBRUARY': '02',
    'MART': '03', 'MAR': '03', 'MARCH': '03',
    'NİSAN': '04', 'NISAN': '04', 'NİS': '04', 'NIS': '04', 'APR': '04', 'APRIL': '04',
    'MAYIS': '05', 'MAY': '05',
    'HAZİRAN': '06', 'HAZIRAN': '06', 'HAZ': '06', 'JUN': '06', 'JUNE': '06',
    'TEMMUZ': '07', 'TEM': '07', 'JUL': '07', 'JULY': '07',
    'AĞUSTOS': '08', 'AGUSTOS': '08', 'AĞU': '08', 'AGU': '08', 'AUG': '08', 'AUGUST': '08',
    'EYLÜL': '09', 'EYLUL': '09', 'EYL': '09', 'SEP': '09', 'SEPTEMBER': '09',
    'EKİM': '10', 'EKIM': '10', 'EKİ': '10', 'EKI': '10', 'OCT': '10', 'OCTOBER': '10',
    'KASIM': '11', 'KAS': '11', 'NOV': '11', 'NOVEMBER': '11',
    'ARALIK': '12', 'ARA': '12', 'DEC': '12', 'DECEMBER': '12'
  };
  const trMonthMatch = str.match(/^(\d{1,2})\s+([a-zA-ZçÇğĞıİöÖşŞüÜ]+)\s+(\d{2,4})$/);
  if (trMonthMatch) {
    const d = trMonthMatch[1].padStart(2, '0');
    const mStr = trMonthMatch[2].toUpperCase();
    const yStr = trMonthMatch[3].length === 2 ? `20${trMonthMatch[3]}` : trMonthMatch[3];
    if (turkishMonths[mStr]) {
      return `${d}.${turkishMonths[mStr]}.${yStr}`;
    }
  }

  // 6. YYYY-MM-DD veya YYYY/MM/DD veya YYYY.MM.DD -> GG.AA.YYYY
  const ymd = str.match(/^(\d{4})[\.\/-](\d{1,2})[\.\/-](\d{1,2})$/);
  if (ymd) {
    const d = ymd[3].padStart(2, '0');
    const m = ymd[2].padStart(2, '0');
    const y = ymd[1];
    return `${d}.${m}.${y}`;
  }

  // 7. DD.MM.YYYY veya DD.MM.YY (Noktalı format) -> GG.AA.YYYY
  const dotMatch = str.match(/^(\d{1,2})\.(\d{1,2})\.(\d{2,4})$/);
  if (dotMatch) {
    const d = dotMatch[1].padStart(2, '0');
    const m = dotMatch[2].padStart(2, '0');
    let y = dotMatch[3];
    if (y.length === 2) y = parseInt(y, 10) > 50 ? `19${y}` : `20${y}`;
    return `${d}.${m}.${y}`;
  }

  // 8. M/D/YY veya M/D/YYYY veya DD/MM/YYYY (Slaşlı format) -> GG.AA.YYYY
  // Kullanıcı kuralı: "TARİHLER AY YIL GÜN ŞEKLİNDE OLACAKTIR KESİNLİKLE HATADIR."
  const slashDate = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (slashDate) {
    const p1 = parseInt(slashDate[1], 10);
    const p2 = parseInt(slashDate[2], 10);
    const yearPart = slashDate[3];
    const y = yearPart.length === 2 ? (parseInt(yearPart, 10) > 50 ? `19${yearPart}` : `20${yearPart}`) : yearPart;

    let day = p1;
    let month = p2;
    // Eğer 1. kısım <= 12 ve 2. kısım > 12 ise: Kesinlikle Ay/Gün/Yıl formatıdır (Örn: 7/31/2026 -> 31.07.2026)
    if (p1 <= 12 && p2 > 12) {
      day = p2;
      month = p1;
    } else if (p1 > 12 && p2 <= 12) {
      // Eğer 1. kısım > 12 ve 2. kısım <= 12 ise: Kesinlikle Gün/Ay/Yıl formatıdır (Örn: 31/7/2026 -> 31.07.2026)
      day = p1;
      month = p2;
    } else {
      // Her ikisi de <= 12 ise: Varsayılan olarak Gün/Ay kabul et
      day = p1;
      month = p2;
    }
    return `${String(day).padStart(2, '0')}.${String(month).padStart(2, '0')}.${y}`;
  }

  // 9. DD-MM-YYYY veya MM-DD-YYYY (Tireli format) -> GG.AA.YYYY
  const dashDate = str.match(/^(\d{1,2})-(\d{1,2})-(\d{2,4})$/);
  if (dashDate) {
    const p1 = parseInt(dashDate[1], 10);
    const p2 = parseInt(dashDate[2], 10);
    const yearPart = dashDate[3];
    const y = yearPart.length === 2 ? (parseInt(yearPart, 10) > 50 ? `19${yearPart}` : `20${yearPart}`) : yearPart;

    let day = p1;
    let month = p2;
    if (p1 <= 12 && p2 > 12) {
      day = p2;
      month = p1;
    }
    return `${String(day).padStart(2, '0')}.${String(month).padStart(2, '0')}.${y}`;
  }

  return str;
};

// Satırın başlık olup olmadığını kesin olarak tespit eden fonksiyon
// Kullanıcı kuralı: "VERİ GÜNCELLEMEDE BAŞLIK 1. SATIR ÜRÜN DİYE ATMIŞ HATADIR."
export const isHeaderLikeRow = (r: any[]): boolean => {
  if (!Array.isArray(r) || r.length === 0) return true;

  const nonEmpty = r.filter(c => c !== undefined && c !== null && String(c).trim() !== "");
  if (nonEmpty.length === 0) return true;

  // Başlık banner'ı (tüm satırda aynı metin tekrarlanıyorsa örn: "HANGAR YER DESTEK TEÇHİZATLARI")
  const uniqueVals = new Set(nonEmpty.map(c => String(c).trim().toUpperCase()));
  if (uniqueVals.size <= 1 && nonEmpty.length >= 2) return true;

  const c0 = String(r[0] || "").trim().toUpperCase();
  const c1 = String(r[1] || "").trim().toUpperCase();
  const c2 = String(r[2] || "").trim().toUpperCase();
  const c3 = String(r[3] || "").trim().toUpperCase();
  const c4 = String(r[4] || "").trim().toUpperCase();
  const c5 = String(r[5] || "").trim().toUpperCase();
  const c6 = String(r[6] || "").trim().toUpperCase();

  if (
    c0 === "SIRA NO" || c0 === "SIRA" || c0 === "NO" || c0 === "NO." || c0 === "S.N." || c0 === "S.NO" || c0 === "S.NU" ||
    c1 === "TEÇHİZAT ADI" || c1 === "TECHIZAT ADI" ||
    c1 === "MALZEME ADI" || c1 === "MALZEME / PARÇA ADI" ||
    c1 === "ÜRÜN ADI" || c1 === "URUN ADI" || c1 === "ÜRÜN" || c1 === "URUN" ||
    c1 === "ARAÇ PLAKASI" || c1 === "ARAÇ PLAKASI / TANIMI" ||
    c1 === "TEÇHİZAT" || c1 === "TECHIZAT" || c1 === "MALZEME" ||
    c1 === "PRODUCT NAME" || c1 === "ITEM NAME"
  ) {
    return true;
  }
  if (c2 === "PARÇA NO (P/N) / MODEL" || c2 === "PARÇA NO (P/N)" || c2 === "PARÇA NO" || c2 === "P/N") return true;
  if (c3 === "SERİ NO (S/N)" || c3 === "SERİ NO" || c3 === "S/N") return true;
  if (c4 === "MİKTAR / KAPASİTE" || c4 === "MİKTAR") return true;
  if (c5 === "BULUNDUĞU YER" || c6 === "DURUMU") return true;

  let headerMatches = 0;
  for (const cell of r) {
    const s = String(cell || "").trim().toUpperCase();
    if (
      s === "SIRA NO" || s === "NO." || s === "SIRA" || s === "S.N." || s === "S.NO" ||
      s === "TEÇHİZAT ADI" || s === "TECHIZAT ADI" ||
      s === "PARÇA NO (P/N) / MODEL" || s === "PARÇA NO (P/N)" || s === "PARÇA NO" || s === "P/N" ||
      s === "SERİ NO (S/N)" || s === "SERİ NO" || s === "S/N" ||
      s === "MİKTAR / KAPASİTE" || s === "MİKTAR" ||
      s === "BULUNDUĞU YER" || s === "DURUMU" ||
      s === "KALİBRASYONA TABİ" || s === "BAKIMA TABİ" || s.includes("SON KONTROL") ||
      s.includes("GELECEK KONTROL") || s.includes("YAPAN FİRMA") ||
      s === "AÇIKLAMA" || s === "ÖMÜR BİTİŞ TARİHİ" || s.includes("MAİL GÖNDERİM")
    ) {
      headerMatches++;
    }
  }
  return headerMatches >= 2;
};

// Excel Sayfasındaki Gerçek Başlık Satırı İndeksini Kategori Çeşitliliğine Göre Tespit Eden Fonksiyon
export const detectHeaderRowIndex = (rawRows: any[][]): number => {
  if (!rawRows || rawRows.length === 0) return 0;

  let bestRowIdx = 0;
  let bestScore = -1;

  for (let r = 0; r < Math.min(20, rawRows.length); r++) {
    const rCells = (rawRows[r] || []).map(c => String(c || '').trim());
    const nonEmptyCells = rCells.filter(c => c !== '');
    if (nonEmptyCells.length < 2) continue;

    // Başlık banner satırlarını (tekrar eden tekil metinleri) atla
    const uniqueValues = new Set(nonEmptyCells.map(c => c.toUpperCase()));
    if (uniqueValues.size < 3 && nonEmptyCells.length >= 3) {
      continue;
    }

    let score = 0;
    const upperCells = rCells.map(c => c.toUpperCase());

    // Farklı sütun kategorilerinin varlığını puanla
    const hasSira = upperCells.some(c => (c.includes("SIRA") || c === "NO" || c === "NO." || c === "S.N." || c === "S.NO" || c === "S.NU") && !c.includes("SERİ") && !c.includes("SERI") && !c.includes("PARÇA") && !c.includes("P/N"));
    const hasName = upperCells.some(c => (c.includes("TEÇHİZAT") || c.includes("TECHIZAT") || c.includes("MALZEME") || c.includes("ÜRÜN") || c.includes("URUN") || c.includes("ARAÇ") || c.includes("ARAC") || c.includes("PLAKA") || c.includes("TANIM") || c.includes("EKİPMAN") || c.includes("ALET") || c.includes("ITEM")) && !c.includes("FİRMA") && !c.includes("YAPAN"));
    const hasPn = upperCells.some(c => c.includes("P/N") || c.includes("PN") || c.includes("PARÇA NO") || c.includes("PARCA NO") || c.includes("MODEL") || c.includes("PART"));
    const hasSn = upperCells.some(c => (c.includes("S/N") || c.includes("SN") || c.includes("SERİ NO") || c.includes("SERI NO") || c.includes("SERİ") || c.includes("SERI") || c.includes("SERIAL")) && !c.includes("SIRA"));
    const hasQty = upperCells.some(c => c.includes("MİKTAR") || c.includes("MIKTAR") || c.includes("KAPASİTE") || c.includes("KAPASITE") || c.includes("ADET") || c.includes("QTY"));
    const hasLoc = upperCells.some(c => c.includes("BULUNDUĞU") || c.includes("BULUNDUGU") || c.includes("LOKASYON") || c.includes("KONUM") || c.includes("YER") || c.includes("RAF") || c.includes("DEPO"));
    const hasStatus = upperCells.some(c => c.includes("DURUM") || c.includes("DURUMU") || c.includes("STATUS") || c.includes("FAALİYET"));
    const hasTabi = upperCells.some(c => c.includes("TABİ") || c.includes("TABI") || c.includes("KALİBRASYON") || c.includes("KALIBRASYON") || c.includes("BAKIM") || c.includes("ÖMÜRLÜ") || c.includes("OMURLU"));
    const hasDate = upperCells.some(c => c.includes("SON KONTROL") || c.includes("GELECEK KONTROL") || c.includes("SON BAKIM") || c.includes("GELECEK BAKIM") || c.includes("TARİH") || c.includes("TARIH") || c.includes("ÖMÜR BİTİŞ") || c.includes("OMUR BITIS") || c.includes("MUAYENE") || c.includes("SON KM"));
    const hasFirma = upperCells.some(c => c.includes("FİRMA") || c.includes("FIRMA") || c.includes("TEDARİK") || c.includes("TEDARIK") || c.includes("YAPAN"));
    const hasAciklama = upperCells.some(c => c.includes("AÇIKLAMA") || c.includes("ACIKLAMA") || c.includes("NOT") || c.includes("REMARKS") || c.includes("DESCRIPTION"));

    if (hasSira) score += 15;
    if (hasName) score += 20;
    if (hasPn) score += 15;
    if (hasSn) score += 10;
    if (hasQty) score += 10;
    if (hasLoc) score += 10;
    if (hasStatus) score += 10;
    if (hasTabi) score += 10;
    if (hasDate) score += 15;
    if (hasFirma) score += 10;
    if (hasAciklama) score += 10;

    // Tam metin eşleşmeleri için ekstra bonus
    for (const c of upperCells) {
      if (
        c === "SIRA NO" || c === "TEÇHİZAT ADI" || c === "TECHİZAT ADI" || 
        c.includes("PARÇA NO (P/N)") || c.includes("SERİ NO (S/N)") || 
        c.includes("MİKTAR / KAPASİTE") || c === "BULUNDUĞU YER" || 
        c === "DURUMU" || c === "BAKIMA TABİ" || 
        c.includes("SON KONTROL") || c.includes("GELECEK KONTROL") || 
        c.includes("YAPAN FİRMA") || c === "AÇIKLAMA"
      ) {
        score += 5;
      }
    }

    if (score > bestScore && score >= 25) {
      bestScore = score;
      bestRowIdx = r;
    }
  }

  return bestRowIdx;
};

/**
 * Excel çalışma sayfasındaki birleştirilmiş hücreleri (merged cells - worksheet['!merges'])
 * unroll yaparak tüm birleştirilmiş hücre aralığına ana hücrenin değerini aktarır.
 * Böylece sheet_to_json okuduğunda birleştirilmiş satırlar boş ('') gelmez!
 */
export const unmergeAndFillWorksheet = (worksheet: XLSX.WorkSheet): void => {
  if (!worksheet || !worksheet['!merges'] || !Array.isArray(worksheet['!merges'])) return;
  
  try {
    worksheet['!merges'].forEach(range => {
      const startCellRef = XLSX.utils.encode_cell(range.s);
      const startCell = worksheet[startCellRef];
      if (!startCell) return;
      
      for (let R = range.s.r; R <= range.e.r; ++R) {
        for (let C = range.s.c; C <= range.e.c; ++C) {
          if (R === range.s.r && C === range.s.c) continue;
          const cellRef = XLSX.utils.encode_cell({ r: R, c: C });
          if (!worksheet[cellRef] || worksheet[cellRef].v === undefined || worksheet[cellRef].v === '') {
            worksheet[cellRef] = { ...startCell };
          }
        }
      }
    });
  } catch (err) {
    console.warn("unmergeAndFillWorksheet warning:", err);
  }
};

/**
 * Aynı ürünün birden fazla lokasyonda bulunması durumunda satırları tek ürün altında toplayan
 * ve birleştirilmiş hücreli/boş ad satırlarını ana ürüne katan yardımcı fonksiyon.
 */
export const groupMultiLocationRowsHelper = (
  rawRows: string[][],
  nameColIdx: number = 1,
  locColIdx: number = 5,
  miktarColIdx: number = 4,
  siraColIdx: number = 0,
  pnColIdx: number = 2,
  seriNoColIdx: number = 3
): string[][] => {
  const grouped: string[][] = [];

  for (let i = 0; i < rawRows.length; i++) {
    const row = [...rawRows[i]];
    const rawSira = (row[siraColIdx] || "").trim();
    const rawName = (row[nameColIdx] || "").trim();
    const rawPn = pnColIdx >= 0 ? (row[pnColIdx] || "").trim() : "";
    const rawSn = seriNoColIdx >= 0 ? (row[seriNoColIdx] || "").trim() : "";
    const rawLoc = locColIdx >= 0 ? (row[locColIdx] || "").trim() : "";
    const rawMiktar = miktarColIdx >= 0 ? (row[miktarColIdx] || "").trim() : "";

    const lastRow = grouped.length > 0 ? grouped[grouped.length - 1] : null;
    const lastName = lastRow ? (lastRow[nameColIdx] || "").trim() : "";
    const lastPn = lastRow && pnColIdx >= 0 ? (lastRow[pnColIdx] || "").trim() : "";
    const lastSira = lastRow ? (lastRow[siraColIdx] || "").trim() : "";

    const isNameEmptyOrDash = !rawName || rawName === "-" || rawName === "--";
    const isSiraEmptyOrDash = !rawSira || rawSira === "-" || rawSira === "--";
    const isSamePn = rawPn !== "" && lastPn !== "" && rawPn.toUpperCase() === lastPn.toUpperCase();
    const isSameName = rawName !== "" && lastName !== "" && rawName.toUpperCase() === lastName.toUpperCase();
    const isSameSira = rawSira !== "" && lastSira !== "" && rawSira === lastSira;

    // Alt lokasyon/devam satırı tespiti:
    // 1) Eğer ürün adı boşsa veya "-" ise ve bir önceki ürün varsa, BU KESİNLİKLE BİR ÖNCEKİ ÜRÜNÜN ALT LOKASYONUDUR!
    // 2) Eğer aynı isim ve aynı P/N varsa, yine aynı ürünün farklı bir lokasyonudur!
    // 3) Eğer aynı sıra numarası ve aynı isim/PN varsa, yine aynı ürünün satırıdır!
    const isSubLocation =
      !!lastRow &&
      (
        isNameEmptyOrDash ||
        (isSameName && (isSamePn || !rawPn || !lastPn)) ||
        (isSameSira && (isSamePn || isSameName))
      );

    if (isSubLocation && lastRow) {
      const existingLocs = locColIdx >= 0 && lastRow[locColIdx] ? lastRow[locColIdx].split('\n') : [""];
      const existingMiktars = miktarColIdx >= 0 && lastRow[miktarColIdx] ? lastRow[miktarColIdx].split('\n') : ["1"];
      const existingSns = seriNoColIdx >= 0 && lastRow[seriNoColIdx] ? lastRow[seriNoColIdx].split('\n') : [""];

      const targetCount = Math.max(1, existingLocs.length, existingMiktars.length, existingSns.length);
      while (existingLocs.length < targetCount) existingLocs.push("");
      while (existingMiktars.length < targetCount) existingMiktars.push("1");
      while (existingSns.length < targetCount) existingSns.push("");

      const matchingLocIdx = rawLoc
        ? existingLocs.findIndex(l => l.trim().toUpperCase() === rawLoc.trim().toUpperCase())
        : -1;

      if (matchingLocIdx !== -1 && (!rawSn || rawSn === "-" || !existingSns[matchingLocIdx] || existingSns[matchingLocIdx] === "-")) {
        // Aynı lokasyon: miktarları topla
        const eVal = Number(existingMiktars[matchingLocIdx]) || 0;
        const nVal = Number(rawMiktar) || 1;
        if (eVal > 0) {
          existingMiktars[matchingLocIdx] = String(eVal + nVal);
        } else {
          existingMiktars[matchingLocIdx] = rawMiktar || "1";
        }
        if (rawSn && rawSn !== "-") {
          existingSns[matchingLocIdx] = existingSns[matchingLocIdx] && existingSns[matchingLocIdx] !== "-" 
            ? `${existingSns[matchingLocIdx]}, ${rawSn}` 
            : rawSn;
        }
      } else {
        // Yeni lokasyon veya alt bölge
        existingLocs.push(rawLoc);
        existingMiktars.push(rawMiktar || "1");
        existingSns.push(rawSn);
      }

      if (locColIdx >= 0) lastRow[locColIdx] = existingLocs.join('\n');
      if (miktarColIdx >= 0) lastRow[miktarColIdx] = existingMiktars.join('\n');
      if (seriNoColIdx >= 0) lastRow[seriNoColIdx] = existingSns.join('\n');

      // Ana satırda eksik kalmış diğer hücreleri alt satırdan doldur
      for (let c = 0; c < row.length; c++) {
        if (c !== siraColIdx && c !== nameColIdx && c !== locColIdx && c !== miktarColIdx && c !== seriNoColIdx) {
          const cellVal = (row[c] || "").trim();
          if (cellVal && cellVal !== "-" && (!lastRow[c] || lastRow[c] === "-")) {
            lastRow[c] = cellVal;
          }
        }
      }
    } else {
      // Bağımsız ürün satırı (ancak adı ve PN'si tamamen boş olan tekil çöp satırları alma)
      if (rawName || rawPn || rawSn !== "-") {
        grouped.push(row);
      }
    }
  }

  // Sıra numaralarını 1'den itibaren ardışık diz
  return grouped.map((row, idx) => {
    const r = [...row];
    if (siraColIdx >= 0) {
      r[siraColIdx] = String(idx + 1);
    }
    return r;
  });
};

// Dosya ve Sayfa Adından Hedef Hava/Kara Birimi ve Alt Bölümünü Otomatik Tespit Eden Fonksiyon
export const detectUnitAndSectionFromNames = (fileName: string, sheetName: string): { techType: string, subSection: string } => {
  const combined = `${fileName} ${sheetName}`
    .toLowerCase()
    .replace(/ı/g, 'i')
    .replace(/ğ/g, 'g')
    .replace(/ü/g, 'u')
    .replace(/ş/g, 's')
    .replace(/ö/g, 'o')
    .replace(/ç/g, 'c');
  
  let techType = '';
  let subSection = 'all';

  if (combined.includes("at-802") || combined.includes("at802") || combined.includes("at 802") || combined.includes("air tractor")) {
    techType = "at802";
  } else if (combined.includes("bell-429") || combined.includes("bell429") || combined.includes("bell 429") || combined.includes("bell")) {
    techType = "bell429";
  } else if (combined.includes("bumbi") || combined.includes("bambi") || combined.includes("t-70_bumbi") || combined.includes("t70_bumbi") || combined.includes("t-70 bumbi")) {
    techType = "t70_bumbi_backet";
  } else if (combined.includes("helitak") || combined.includes("t-70_helitak") || combined.includes("t70_helitak") || combined.includes("t-70 helitak")) {
    techType = "t70_helitak";
  } else if (combined.includes("t-70") || combined.includes("t70") || combined.includes("t 70") || combined.includes("sikorsky")) {
    techType = "t70";
  } else if (combined.includes("b-360") || combined.includes("b360") || combined.includes("b 360") || combined.includes("king air") || combined.includes("kingair")) {
    techType = "b360";
  } else if (combined.includes("c-650") || combined.includes("c650") || combined.includes("c 650") || combined.includes("citation")) {
    techType = "c650";
  } else if (combined.includes("hangar")) {
    techType = "hangar";
  } else if (combined.includes("kara") || combined.includes("arac") || combined.includes("plaka") || combined.includes("forklift") || combined.includes("traktor")) {
    techType = "kara_araclari";
  }

  if (combined.includes("ozel") || combined.includes("özel") || combined.includes("alet")) {
    subSection = "ozel_alet";
  } else if (combined.includes("sarf") || combined.includes("yedek parca") || combined.includes("parca depo")) {
    subSection = "depo_sarf";
  } else if (combined.includes("kimyasal") || combined.includes("yag") || combined.includes("boya") || combined.includes("tiner")) {
    subSection = "depo_kimyasal";
  } else if (combined.includes("yer") || combined.includes("destek") || combined.includes("techizat")) {
    subSection = "yer_destek";
  }

  return { techType, subSection };
};

// Akıllı Başlık ve Kolon Tespiti ile Excel Sayfasını Standart Matris Formatına Çeviren Fonksiyon
export const parseExcelWorksheetToStandardRows = (
  worksheet: XLSX.WorkSheet, 
  techType: string, 
  subSection: string = 'all',
  groupMultiLocationRowsFn?: (rows: string[][], nameIdx: number, locIdx: number, miktarIdx: number, siraIdx: number, pnIdx: number, snIdx: number) => string[][]
): string[][] => {
  try {
    // 1. Birleştirilmiş hücreleri (merged cells) unroll yap: böylece alt satırlar boş kalmaz
    unmergeAndFillWorksheet(worksheet);

    const rawRows = XLSX.utils.sheet_to_json<string[]>(worksheet, { header: 1, defval: "", raw: false });
    if (!rawRows || rawRows.length === 0) return [];

    const headerRowIdx = detectHeaderRowIndex(rawRows);

    const rawHeaders = (rawRows[headerRowIdx] || []).map(h => String(h || '').trim().toUpperCase());
    const finalHeaders = rawHeaders.map((h, hIdx) => h || `KOLON ${hIdx + 1}`);

    const findColIdx = (keywords: string[], excludeKeywords: string[] = []): number => {
      return finalHeaders.findIndex(h => {
        const upper = h.toUpperCase().trim();
        const hasKey = keywords.some(k => upper.includes(k.toUpperCase()));
        const hasExclude = excludeKeywords.some(ex => upper.includes(ex.toUpperCase()));
        return hasKey && !hasExclude;
      });
    };

    const siraColIdx = findColIdx(["SIRA", "NO.", "S.N.", "S.NO", "S.NU"], ["SERİ", "SERI", "PARÇA", "PARCA", "P/N", "MODEL"]);
    const nameColIdx = findColIdx(
      ["TEÇHİZAT", "TECHİZAT", "MALZEME", "ARAÇ", "ARAC", "PLAKA", "ÜRÜN", "URUN", "EKİPMAN", "TANIM", "NAME", "ALET", "ITEM"],
      ["FİRMA", "FIRMA", "KONTROL", "BAKIM", "YAPAN", "SIRA"]
    );
    const pnColIdx = findColIdx(["P/N", "PN", "PARÇA NO", "PARCA NO", "MODEL", "PART NUMBER", "PART NO"]);
    const snColIdx = findColIdx(["S/N", "SN", "SERİ NO", "SERI NO", "SERİ", "SERI", "SERIAL"], ["SIRA"]);
    const miktarColIdx = findColIdx(["MİKTAR", "MIKTAR", "KAPASİTE", "KAPASITE", "ADET", "QTY", "QUANTITY"]);
    const locColIdx = findColIdx(["BULUNDUĞU", "BULUNDUGU", "LOKASYON", "KONUM", "YER", "RAF", "DEPO", "LOCATION"]);
    const durumColIdx = findColIdx(["DURUM", "DURUMU", "STATUS", "FAALİYET"]);
    const kalibTabiColIdx = findColIdx(["KALİBRASYONA TABİ", "KALIBRASYONA TABI", "BAKIMA TABİ", "BAKIMA TABI", "TABİ Mİ", "TABI MI", "TABİ", "TABI", "ÖMÜRLÜ", "OMURLU"]);
    const sonBakimColIdx = findColIdx(["SON KONTROL", "SON BAKIM", "SON KALİBRASYON", "SON TEST", "SON MUAYENE", "SON KM", "YAPILAN KONTROL", "SON TARİH"]);
    const gelecekBakimColIdx = findColIdx(["GELECEK KONTROL", "GELECEK BAKIM", "GELECEK KALİBRASYON", "BİR SONRAKİ", "SONRAKİ BAKIM", "SONRAKİ KONTROL", "ÖMÜR BİTİŞ", "OMUR BITIS", "SON KULLANMA", "EXPIRY"]);
    const firmaColIdx = findColIdx(["KONTROLÜ YAPAN", "KONTROLU YAPAN", "YAPAN FİRMA", "YAPAN FIRMA", "FİRMA", "FIRMA", "TEDARİK", "TEDARIK", "SERVİS", "VENDOR", "SUPPLIER"]);
    const aciklamaColIdx = findColIdx(["AÇIKLAMA", "ACIKLAMA", "AÇIKLAMALAR", "ACIKLAMALAR", "NOT", "NOTLAR", "DESCRIPTION", "REMARKS", "DETAY", "ÖZEL NOT", "LOT", "PARTİ"]);
    const mailColIdx = findColIdx(["MAİL GÖNDERİM", "MAIL GONDERIM", "90 GÜN", "90 GUN", "E-POSTA", "MAIL", "MAİL"]);

    const isKara = techType === 'kara_araclari';
    const isDepoType = subSection === 'depo_sarf' || subSection === 'depo_kimyasal';

    const rawParsedRows: string[][] = [];
    for (let r = headerRowIdx + 1; r < rawRows.length; r++) {
      const rawRow = rawRows[r] || [];
      const isRowEmpty = rawRow.every(cell => String(cell || '').trim() === '');
      if (isRowEmpty) continue;

      // Kullanıcı kuralı: "VERİ GÜNCELLEMEDE BAŞLIK 1. SATIR ÜRÜN DİYE ATMIŞ HATADIR."
      if (isHeaderLikeRow(rawRow)) continue;

      const getVal = (idx: number, fallback = "") => {
        if (idx >= 0 && rawRow[idx] !== undefined && rawRow[idx] !== null) {
          const valStr = String(rawRow[idx]).trim();
          if (valStr !== "") return valStr;
        }
        return fallback;
      };

      // Standart kolon fall back pozisyonları (Kullanıcı kuralı: Sıra no A sütunu [0], Teçhizat adı B sütunu [1], P/N C [2], S/N D [3], Miktar E [4], Konum F [5], Durum G [6], Bakım H [7], Son Bakım I [8], Gelecek Bakım J [9], Firma K [10], Açıklama L [11])
      const rawSira = getVal(siraColIdx, siraColIdx < 0 && rawRow[0] !== undefined ? String(rawRow[0]).trim() : "");
      const rawName = getVal(nameColIdx, nameColIdx < 0 && rawRow[1] !== undefined ? String(rawRow[1]).trim() : (rawRow[0] && !/^\d+$/.test(String(rawRow[0]).trim()) ? String(rawRow[0]).trim() : ""));
      const rawPn = getVal(pnColIdx, pnColIdx < 0 && rawRow[2] !== undefined ? String(rawRow[2]).trim() : "");
      const rawSn = getVal(snColIdx, snColIdx < 0 && rawRow[3] !== undefined ? String(rawRow[3]).trim() : "-");
      const rawMiktar = getVal(miktarColIdx, miktarColIdx < 0 && rawRow[4] !== undefined ? String(rawRow[4]).trim() : "1");
      const rawLoc = getVal(locColIdx, locColIdx < 0 && rawRow[5] !== undefined ? String(rawRow[5]).trim() : "");
      const rawDurum = getVal(durumColIdx, durumColIdx < 0 && rawRow[6] !== undefined ? String(rawRow[6]).trim() : "FAAL");
      
      let rawKalib = getVal(kalibTabiColIdx, kalibTabiColIdx < 0 && rawRow[7] !== undefined ? String(rawRow[7]).trim() : "");
      if (!rawKalib) {
        rawKalib = "EVET";
      }

      // Tarihleri temizle ve GG.AA.YYYY formatına çevir
      const rawSonBakim = cleanAndFormatDateString(getVal(sonBakimColIdx, sonBakimColIdx < 0 && rawRow[8] !== undefined ? String(rawRow[8]).trim() : ""));
      const rawGelecekBakim = cleanAndFormatDateString(getVal(gelecekBakimColIdx, gelecekBakimColIdx < 0 && rawRow[9] !== undefined ? String(rawRow[9]).trim() : ""));
      const rawFirma = getVal(firmaColIdx, firmaColIdx < 0 && rawRow[10] !== undefined ? String(rawRow[10]).trim() : "");
      
      let rawAciklama = "";
      if (aciklamaColIdx >= 0 && rawRow[aciklamaColIdx] !== undefined) {
        rawAciklama = String(rawRow[aciklamaColIdx]).trim();
      } else {
        const fallbackIdx = finalHeaders.length >= 12 ? 11 : 10;
        if (rawRow[fallbackIdx] !== undefined) {
          rawAciklama = String(rawRow[fallbackIdx]).trim();
        }
      }

      const rawMail = getVal(mailColIdx, "");

      let targetRow: string[] = [];
      if (isKara) {
        targetRow = [
          rawSira, // Sıra no boşsa boş bırakılır, gruplamada doldurulur
          rawName,
          rawPn,
          rawLoc,
          rawMiktar || rawSonBakim || "",
          rawDurum || "FAAL",
          rawSonBakim,
          rawGelecekBakim,
          rawFirma,
          rawAciklama,
          rawMail,
          subSection !== 'all' ? subSection : 'kara_araclari'
        ];
      } else if (isDepoType) {
        targetRow = [
          rawSira,
          rawName,
          rawPn,
          rawSn || "-",
          rawMiktar || "1",
          rawLoc,
          rawDurum || "FAAL",
          rawKalib || "HAYIR",
          rawGelecekBakim || rawSonBakim || "",
          rawFirma,
          rawAciklama,
          rawMail,
          subSection
        ];
      } else {
        targetRow = [
          rawSira,
          rawName,
          rawPn,
          rawSn || "-",
          rawMiktar || "1",
          rawLoc,
          rawDurum || "FAAL",
          rawKalib || "EVET",
          rawSonBakim,
          rawGelecekBakim,
          rawFirma,
          rawAciklama,
          rawMail,
          subSection !== 'all' ? subSection : 'yer_destek'
        ];
      }

      // Satırda en az isim, PN, SN veya Lokasyon varsa al
      if (targetRow[1] || targetRow[2] || (targetRow[3] && targetRow[3] !== "-") || targetRow[5]) {
        rawParsedRows.push(targetRow);
      }
    }

    const groupFn = groupMultiLocationRowsFn || groupMultiLocationRowsHelper;
    return groupFn(
      rawParsedRows,
      1,
      isKara ? 3 : 5,
      isKara ? -1 : 4,
      0,
      2,
      isKara ? -1 : 3
    );
  } catch (e) {
    console.error("Excel parse error for " + techType, e);
    return [];
  }
};
