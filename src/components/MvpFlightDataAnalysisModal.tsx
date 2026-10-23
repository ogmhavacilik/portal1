import React, { useState, useRef, useMemo, useEffect, useCallback } from 'react';
import {
  X,
  UploadCloud,
  FileSpreadsheet,
  AlertTriangle,
  Flame,
  CheckCircle2,
  Gauge,
  ArrowRight,
  Filter,
  Eye,
  Info,
  ChevronDown,
  ChevronUp,
  FileText,
  RotateCcw,
  Plane,
  Save,
  Check,
  AlertCircle,
  Clock,
  ChevronLeft,
  ChevronRight,
  ChevronsLeft,
  ChevronsRight,
  SlidersHorizontal,
  Activity,
  Search,
  CheckCircle,
  RefreshCw,
  FileCheck,
  Download,
  Trash2,
  TrendingUp,
  ExternalLink,
  Database,
  Camera,
  KeyRound,
  CheckSquare
} from 'lucide-react';
import html2canvas from 'html2canvas-pro';
import { AIRCRAFT_TAILS } from './DepoManagementModal';
import { FlightEngineChartModal } from './FlightEngineChartModal';

/**
 * Resmi AT-802 Yangın Söndürme Filosu (Depo Envanteri Birebir Eşleşmesi)
 */
export interface At802AircraftInfo {
  tail: string;
  callSign: string;
  code: string;
  fullLabel: string;
  motorSn: string;
  model: string;
  base: string;
}

// Depo kuyruk kodlarına göre motor seri numaraları ve üs bilgileri
const AT802_ENGINE_MAP: Record<string, { motorSn: string; base: string; civilReg?: string }> = {
  'OR-2021': { motorSn: 'PCE-RZ0354', base: 'Antalya Orman Havacılık', civilReg: 'TC-ZAA' },
  'OR-2022': { motorSn: 'PCE-RZ0306', base: 'Antalya Orman Havacılık', civilReg: 'TC-ZAB' },
  'OR-2023': { motorSn: 'PCE-RZ0362', base: 'Muğla Milas', civilReg: 'TC-ZAC' },
  'OR-2024': { motorSn: 'PCE-RZ0333', base: 'İzmir Adnan Menderes', civilReg: 'TC-ZAD' },
  'OR-2025': { motorSn: 'PCE-RZ0332', base: 'Çanakkale', civilReg: 'TC-ZAE' },
  'OR-2026': { motorSn: 'PCE-RZ0357', base: 'Bursa Yenişehir', civilReg: 'TC-ZAF' },
  'OR-2027': { motorSn: 'PCE-RZ0389', base: 'Balıkesir', civilReg: 'TC-ZAG' },
  'OR-2028': { motorSn: 'PCE-RZ0402', base: 'Adana', civilReg: 'TC-ZAH' },
  'OR-2029': { motorSn: 'PCE-RZ0406', base: 'Mersin', civilReg: 'TC-ZAI' },
  'OR-2030': { motorSn: 'PCE-RZ0412', base: 'Antalya Karain', civilReg: 'TC-ZAJ' },
  'OR-2031': { motorSn: 'PCE-RZ0437', base: 'Ankara Güvercinlik', civilReg: 'TC-ZAK' },
  'OR-2036': { motorSn: 'PCE-RZ0317', base: 'Muğla Dalaman', civilReg: 'TC-ZAP' },
  'OR-2037': { motorSn: 'PCE-RZ0413', base: 'Aydın Çıldır', civilReg: 'TC-ZAR' },
  'OR-2038': { motorSn: 'PCE-RZ0372', base: 'Denizli Çardak', civilReg: 'TC-ZAS' },
  'OR-2039': { motorSn: 'PCE-RZ0441', base: 'Isparta Süleyman Demirel', civilReg: 'TC-ZAT' },
  'OR-2040': { motorSn: 'PCE-RZ0455', base: 'Antalya Gazipaşa', civilReg: 'TC-ZAU' },
};

// Depodan gelen 16 adet AT-802 uçağının birebir listesi
export const AT802_DEPO_TAIL_STRINGS: string[] = AIRCRAFT_TAILS?.at802 || [
  'ORMAN 21 (OR-2021) - AT-802',
  'ORMAN 22 (OR-2022) - AT-802',
  'ORMAN 23 (OR-2023) - AT-802',
  'ORMAN 24 (OR-2024) - AT-802',
  'ORMAN 25 (OR-2025) - AT-802',
  'ORMAN 26 (OR-2026) - AT-802',
  'ORMAN 27 (OR-2027) - AT-802',
  'ORMAN 28 (OR-2028) - AT-802',
  'ORMAN 29 (OR-2029) - AT-802',
  'ORMAN 30 (OR-2030) - AT-802',
  'ORMAN 31 (OR-2031) - AT-802',
  'ORMAN 36 (OR-2036) - AT-802',
  'ORMAN 37 (OR-2037) - AT-802',
  'ORMAN 38 (OR-2038) - AT-802',
  'ORMAN 39 (OR-2039) - AT-802',
  'ORMAN 40 (OR-2040) - AT-802'
];

export const AT802_FLEET: At802AircraftInfo[] = AT802_DEPO_TAIL_STRINGS.map(fullStr => {
  const orMatch = fullStr.match(/OR-?(\d+)/i);
  const code = orMatch ? `OR-${orMatch[1]}` : fullStr;
  const callSignMatch = fullStr.match(/ORMAN\s*\d+/i);
  const callSign = callSignMatch ? callSignMatch[0] : fullStr;
  const eng = AT802_ENGINE_MAP[code] || {
    motorSn: `PCE-AG08${orMatch ? orMatch[1].slice(-2) : '00'}`,
    base: 'Orman Havacılık Üssü'
  };
  return {
    tail: eng.civilReg || code,
    code,
    callSign,
    fullLabel: fullStr, // Birebir depodaki kuyruk verisi
    motorSn: eng.motorSn,
    model: 'AT-802F Fire Boss',
    base: eng.base
  };
});

/**
 * Resmi AT-802 MVP Uçuş Parametreleri Başlık Listesi (Görseldeki 44 Sütun)
 */
export const OFFICIAL_ALL_MVP_HEADERS = [
  'TIME',          // Col 1 (A)
  'MSTR_WRN',      // Col 2 (B)
  'NG %',          // Col 3 (C) - SARI ODAK
  'NP RPM',        // Col 4 (D) - SARI ODAK
  'F.FLOWL/H',     // Col 5 (E)
  'FUEL L LTR',    // Col 6 (F)
  'FUEL R LTR',    // Col 7 (G)
  'VOLTS V',       // Col 8 (H)
  'AMPS A',        // Col 9 (I)
  'TORQ FT-LB',    // Col 10 (J) - SARI ODAK
  'FUEL P PSI',    // Col 11 (K) - SARI ODAK
  'OIL P PSI',     // Col 12 (L) - SARI ODAK
  'TRQ LO',        // Col 13 (M)
  'VAC INHG',      // Col 14 (N)
  'OAT *C',        // Col 15 (O)
  'ITT *C',        // Col 16 (P) - SARI ODAK
  'OIL T *C',      // Col 17 (Q) - SARI ODAK
  'AMPS 2 A',      // Col 18 (R)
  'BUS V V',       // Col 19 (S)
  'BARO INHG',     // Col 20 (T)
  'PRESS ALT',     // Col 21 (U)
  'DALT FT',       // Col 22 (V)
  'IAS KTS',       // Col 23 (W)
  'TAS KTS',       // Col 24 (X)
  'P-ALT FT',      // Col 25 (Y)
  'M-ALT FT',      // Col 26 (Z)
  'P-MAG DEG',     // Col 27 (AA)
  'TA-MAG DEG',    // Col 28 (AB)
  'M-MAG DEG',     // Col 29 (AC)
  'M-TRK DEG',     // Col 30 (AD)
  'W-DIR DEG',     // Col 31 (AE)
  'W-SPD KTS',     // Col 32 (AF)
  'X-WND KTS',     // Col 33 (AG)
  'H-WND KTS',     // Col 34 (AH)
  'G-SPD KTS',     // Col 35 (AI)
  'LATITUDE',      // Col 36 (AJ)
  'LONGITUDE',     // Col 37 (AK)
  'GPS TIME',      // Col 38 (AL)
  'FLAP DEG',      // Col 39 (AM)
  'RUD-T DEG',     // Col 40 (AN)
  'ELE-T DEG',     // Col 41 (AO)
  'AIL-T DEG',     // Col 42 (AP)
  'BATT T *C',     // Col 43 (AQ)
  'P-RPM RPM'      // Col 44 (AR)
];

/**
 * 0-tabanlı sütun indeksini Excel harfine dönüştürür: 0 -> A, 1 -> B, 25 -> Z, 26 -> AA...
 */
export const getExcelColumnLetter = (colIndex: number): string => {
  let temp = colIndex;
  let letter = '';
  while (temp >= 0) {
    letter = String.fromCharCode((temp % 26) + 65) + letter;
    temp = Math.floor(temp / 26) - 1;
  }
  return letter;
};

/**
 * GÖRSELLERDEKİ LİMİT TABLOSUNDA YALNIZCA SARI DOLGULU 7 ODAK SÜTUN VE KURALLARI:
 * 1. NG (%) [C - Sütun 3]
 * 2. ITT (*C) [P - Sütun 16]
 * 3. NP (RPM) [D - Sütun 4]
 * 4. TORQ (FT-LB) [J - Sütun 10]
 * 5. FUEL P ( PSI) [K - Sütun 11]
 * 6. OIL P (PSI) [L - Sütun 12]
 * 7. OIL T (*C) [Q - Sütun 17]
 */
export interface YellowFocusColumnConfig {
  key: string;
  name: string;
  excelLetter: string;
  colNumber: number; // 1-tabanlı Excel sütun numarası
  expectedIndex: number; // 0-tabanlı dizi indeksi
  rulesText: string;
  matchPattern: RegExp;
}

export const YELLOW_FOCUS_COLUMNS: YellowFocusColumnConfig[] = [
  {
    key: 'ng',
    name: 'NG (%)',
    excelLetter: 'C',
    colNumber: 3,
    expectedIndex: 2,
    rulesText: 'Yeşil: < %104 | Sarı: %104-105,5 (0-20 sn) | Kırmızı: > %105,5 veya %104-105,5 (>20 sn)',
    matchPattern: /^ng\b(?!.*\b(lo|hi|warn|status|flg)\b)/i
  },
  {
    key: 'itt',
    name: 'ITT (°C)',
    excelLetter: 'P',
    colNumber: 16,
    expectedIndex: 15,
    rulesText: 'Yeşil: < 870°C | Sarı: 870-910°C (20 saniye), 910-1000°C (5 saniye) | Kırmızı: > 1000°C',
    matchPattern: /^itt\b(?!.*\b(lo|hi|warn|status|flg)\b)/i
  },
  {
    key: 'np',
    name: 'NP (RPM)',
    excelLetter: 'D',
    colNumber: 4,
    expectedIndex: 3,
    rulesText: 'Yeşil: < 1700 RPM | Sarı: 1700-1870 RPM | Kırmızı: > 1870 RPM',
    matchPattern: /^np\b(?!.*\b(lo|hi|warn|status|flg)\b)/i
  },
  {
    key: 'torq',
    name: 'TQ (LB/FT)',
    excelLetter: 'J',
    colNumber: 10,
    expectedIndex: 9,
    rulesText: 'Yeşil: < 5252 LB/FT | Sarı: 5252-6092 LB/FT (20 saniye) | Kırmızı: > 6092 LB/FT',
    matchPattern: /^(tq|torq|torque)\b(?!.*\b(lo|hi|warn|status|flg)\b)/i
  },
  {
    key: 'fuel_p',
    name: 'FUEL P (PSI)',
    excelLetter: 'K',
    colNumber: 11,
    expectedIndex: 10,
    rulesText: 'Yeşil: 10-25 PSI | Kırmızı: < 10 PSI veya > 25 PSI',
    matchPattern: /^fuel\s*p\b(?!.*\b(lo|hi|warn|status|flg)\b)/i
  },
  {
    key: 'oil_p',
    name: 'OIL P (PSI)',
    excelLetter: 'L',
    colNumber: 12,
    expectedIndex: 11,
    rulesText: 'Yeşil: 90-135 PSI | Sarı: 60-90 PSI | Kırmızı: < 60 PSI veya > 135 PSI',
    matchPattern: /^oil\s*p\b(?!.*\b(lo|hi|warn|status|flg)\b)/i
  },
  {
    key: 'oil_t',
    name: 'OIL T (°C)',
    excelLetter: 'Q',
    colNumber: 17,
    expectedIndex: 16,
    rulesText: 'Yeşil: -40-105°C | Sarı: 105-110°C (10 dakika) | Kırmızı: > 110°C',
    matchPattern: /^oil\s*t\b(?!.*\b(lo|hi|warn|status|flg)\b)/i
  }
];

export const isYellowFocusHeader = (headerName: string, colIndex: number): boolean => {
  const clean = headerName.trim();
  return YELLOW_FOCUS_COLUMNS.some(c => c.matchPattern.test(clean));
};

export const parseCleanNumber = (val: string | number | undefined | null): number => {
  if (val === undefined || val === null) return NaN;
  if (typeof val === 'number') return val;
  let str = String(val).trim();
  if (!str) return NaN;

  str = str.replace(/"/g, '').replace(/'/g, '').replace(/%/g, '').replace(/rpm/gi, '').trim();

  if (str.includes(',') && str.includes('.')) {
    const lastComma = str.lastIndexOf(',');
    const lastDot = str.lastIndexOf('.');
    if (lastComma > lastDot) {
      str = str.replace(/\./g, '').replace(',', '.');
    } else {
      str = str.replace(/,/g, '');
    }
  } else if (str.includes(',')) {
    str = str.replace(',', '.');
  }

  const num = parseFloat(str);
  return isNaN(num) ? NaN : num;
};

export const parseCsvLineWithQuotes = (line: string, delimiter: string): string[] => {
  const result: string[] = [];
  let current = '';
  let inQuotes = false;

  for (let i = 0; i < line.length; i++) {
    const char = line[i];

    if (char === '"') {
      if (inQuotes && line[i + 1] === '"') {
        current += '"';
        i++;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (char === delimiter && !inQuotes) {
      result.push(current.trim());
      current = '';
    } else {
      current += char;
    }
  }
  result.push(current.trim());
  return result;
};

export const detectDelimiterAndHeader = (rawLines: string[]): { delimiter: string; headerLineIndex: number } => {
  const candidateDelimiters = [',', ';', '\t'];

  let bestScore = -1;
  let bestHeaderIndex = 0;
  let bestDelimiter = ',';

  // İlk 60 satırı tarayarak havacılık/motor başlıklarını ve en uygun ayırıcıyı tespit et
  for (let i = 0; i < Math.min(rawLines.length, 60); i++) {
    const line = rawLines[i].trim();
    if (!line) continue;
    const upper = line.toUpperCase();

    // MVP ve uçuş kaydı başlık anahtar kelimeleri
    const keywords = [
      'TIME', 'NG', 'NP', 'TORQ', 'TRQ', 'ITT', 'MSTR', 'FUEL', 'VOLT', 'AMPS', 'RPM',
      'OIL', 'ALT', 'BARO', 'TAS', 'IAS', 'FLAP', 'LATITUDE', 'LONGITUDE', 'F.FLOW', 'F_FLOW'
    ];
    let kwCount = 0;
    for (const kw of keywords) {
      if (upper.includes(kw)) kwCount++;
    }

    for (const delim of candidateDelimiters) {
      const parts = parseCsvLineWithQuotes(line, delim);
      const colCount = parts.length;

      // Başlık satırı genelde 5'ten fazla sütun içerir (MVP-50T için ~44 sütun)
      if (colCount >= 5) {
        let score = kwCount * 30 + colCount;

        // İlk sütun TIME içeriyorsa kesin başlık göstergesidir
        const firstCol = parts[0]?.toUpperCase().trim();
        if (firstCol === 'TIME' || firstCol === '"TIME"' || firstCol?.startsWith('TIME')) {
          score += 150;
        }

        // Elemanların çoğu metin ise (sayı değilse) başlık satırıdır
        const nonNumericCount = parts.filter(p => isNaN(Number(p.replace(/,/g, '.')))).length;
        if (nonNumericCount / colCount > 0.6) {
          score += 60;
        }

        if (score > bestScore) {
          bestScore = score;
          bestHeaderIndex = i;
          bestDelimiter = delim;
        }
      }
    }
  }

  // Eğer anahtar kelimelerden bulunamadıysa, ilk çok sütunlu satırı al
  if (bestScore <= 0) {
    for (let i = 0; i < Math.min(rawLines.length, 60); i++) {
      const line = rawLines[i].trim();
      if (!line) continue;
      for (const delim of candidateDelimiters) {
        const parts = parseCsvLineWithQuotes(line, delim);
        if (parts.length >= 10 && parts.length * 10 > bestScore) {
          bestScore = parts.length * 10;
          bestHeaderIndex = i;
          bestDelimiter = delim;
        }
      }
    }
  }

  // Veri satırlarını inceleyerek ayırıcıyı doğrula (virgül vs noktalı virgül)
  const dataLines = rawLines.slice(bestHeaderIndex + 1, bestHeaderIndex + 11).filter(l => l.trim().length > 0);
  if (dataLines.length > 0) {
    let commaCols = 0;
    let semiCols = 0;
    for (const dl of dataLines) {
      commaCols += parseCsvLineWithQuotes(dl, ',').length;
      semiCols += parseCsvLineWithQuotes(dl, ';').length;
    }
    const avgComma = commaCols / dataLines.length;
    const avgSemi = semiCols / dataLines.length;

    // Eğer veri satırları virgüle göre ayrıldığında daha çok sütun veriyorsa virgülü kesin seç
    if (avgComma > avgSemi && avgComma >= 5) {
      bestDelimiter = ',';
    } else if (avgSemi > avgComma && avgSemi >= 5) {
      bestDelimiter = ';';
    }
  }

  return { delimiter: bestDelimiter, headerLineIndex: bestHeaderIndex };
};

export const repairUnquotedDecimalCommas = (rawParts: string[], expectedColCount: number): string[] => {
  if (rawParts.length <= expectedColCount) {
    return rawParts;
  }

  const repaired: string[] = [];
  let i = 0;
  while (i < rawParts.length) {
    const curr = rawParts[i];
    const next = rawParts[i + 1];

    if (
      next !== undefined &&
      repaired.length + (rawParts.length - i - 1) >= expectedColCount &&
      /^\d+$/.test(curr.trim()) &&
      /^\d{1,3}$/.test(next.trim())
    ) {
      repaired.push(`${curr.trim()}.${next.trim()}`);
      i += 2;
    } else {
      repaired.push(curr);
      i++;
    }
  }
  return repaired;
};

export interface ExceedanceEpisode {
  id: string;
  colKey: string;
  colName: string;
  excelLetter: string;
  colNumber: number;
  colIndex: number;
  level: 'RED' | 'YELLOW';
  startTimeStr: string;
  endTimeStr: string;
  startSec: number;
  endSec: number;
  durationSeconds: number;
  durationFormatted: string; // e.g. "3 saniye" veya "3,2 saniye"
  startRow: number;
  endRow: number;
  rowCount: number;
  peakValue: number;
  peakValueFormatted: string; // e.g. "105 Tork", "%104,5 Ng", "860°C ITT"
  summaryText: string; // e.g. "3 saniye 105 Tork"
  rule: string;
  durationExceeded?: boolean;
}

export interface ViolationDetail {
  colName: string;
  colIndex: number;
  excelLetter: string;
  colNumber: number;
  value: number;
  rawVal: string;
  level: 'RED' | 'YELLOW';
  rule: string;
  episodeId?: string;
  timeRangeStr?: string; // e.g. "14:23:02.10 - 14:23:05.30"
  durationSeconds?: number;
  durationFormatted?: string;
  summaryText?: string; // e.g. "3 saniye 105 Tork"
  peakValueFormatted?: string;
  maxAllowedSeconds?: number;
  durationExceeded?: boolean;
}

export interface ParsedMvpRow {
  originalIndex: number; // 1-tabanlı orijinal veri satırı indeksi
  timeStr: string;
  timeSec: number | null;
  cells: string[];
  violations: ViolationDetail[];
  severity: 'RED' | 'YELLOW' | 'NORMAL';
  primaryEpisode?: ExceedanceEpisode; // TIME hücresinde kırmızı çerçeve içinde süre aralığı göstermek için
}

export interface MvpAnalysisResult {
  fileName: string;
  tailNo: string;
  motorSn: string;
  delimiter: string;
  metadataLines: string[];
  parsedMetadata: Record<string, string>;
  headers: string[];
  rows: ParsedMvpRow[];
  episodes: ExceedanceEpisode[]; // Süre aralıkları hesaplanmış limit aşım olayları
  redCount: number;
  yellowCount: number;
  durationExceededCount: number;
  peakNg: number | null;
  peakItt: number | null;
  peakTq: number | null;
  peakNp: number | null;
  minFuelP: number | null;
  maxFuelP: number | null;
  minOilP: number | null;
  maxOilP: number | null;
  minOilT: number | null;
  maxOilT: number | null;
  maxFflow?: number | null;
  minFuelLevel?: number | null;
  maxFuelImbalance?: number | null;
  processingTimeMs: number;
}

export interface MultiFlightItem {
  id: string;
  file: File;
  fileName: string;
  content: string;
  result: MvpAnalysisResult;
  isVisited: boolean;
  yellowCount: number;
  redCount: number;
  durationExceededCount: number;
  hasTransferredYellow?: boolean;
}

export interface MvpFlightDataAnalysisModalProps {
  isOpen: boolean;
  onClose: () => void;
  onAddToOlayTakip: (data: {
    kuyrukNo: string;
    tarih: string;
    motorSeriNo: string;
    olayTuru: string;
    tepeNg: string;
    daaRefNo: string;
    verilenSart: string;
    pilot: string;
    teknikDegerlendirme: string;
    parcaDurumu: string;
  }) => void;
  showNotification: (msg: string, type: 'success' | 'error' | 'info') => void;
  defaultTailNo?: string;
}

/**
 * Saliseleri (ondalık saniye ve kesirleri) eksiksiz ayrıştıran zaman çözücü
 * Desteklenen formatlar: HH:MM:SS.sss, HH:MM:SS,sss, HH:MM:SS:cs, MM:SS.sss
 */
export const parseTimeToSeconds = (tStr: string): number | null => {
  if (!tStr) return null;
  const clean = tStr.trim().replace(/^["']|["']$/g, '');
  if (!clean) return null;

  // Örn: "15:28:15.40" veya "15:28:15,40" veya "15:28:15:40"
  // Eğer formatta nokta, virgül veya son kısımda ondalık varsa saniyeye saliseleri de katarak ekle
  const normalized = clean.replace(/[,:](\d+)$/, '.$1');

  const colonParts = normalized.split(':');
  if (colonParts.length === 3) {
    const h = parseFloat(colonParts[0]) || 0;
    const m = parseFloat(colonParts[1]) || 0;
    // 3. kısım saniye ve ondalık/salise içerebilir (örn: "15.40" veya "15:40")
    const secPart = colonParts[2].includes('.') ? parseFloat(colonParts[2]) : parseFloat(colonParts[2].replace(',', '.')) || 0;
    return h * 3600 + m * 60 + secPart;
  } else if (colonParts.length === 4) {
    const h = parseFloat(colonParts[0]) || 0;
    const m = parseFloat(colonParts[1]) || 0;
    const s = parseFloat(colonParts[2]) || 0;
    const cs = parseFloat(colonParts[3]) || 0;
    const frac = cs >= 100 ? cs / 1000 : cs / 100;
    return h * 3600 + m * 60 + s + frac;
  } else if (colonParts.length === 2) {
    const m = parseFloat(colonParts[0]) || 0;
    const s = parseFloat(colonParts[1].replace(',', '.')) || 0;
    return m * 60 + s;
  } else if (colonParts.length === 1) {
    const num = parseFloat(colonParts[0].replace(',', '.'));
    return isNaN(num) ? null : num;
  }
  return null;
};

/**
 * Süre farkını havacılık raporlama standardında formatlar (örn: "3 saniye" veya "3,2 saniye")
 */
export const formatDurationSeconds = (sec: number): string => {
  if (sec <= 0) return '0 saniye';
  const rounded = Math.round(sec);
  if (Math.abs(sec - rounded) < 0.08) {
    return `${rounded} saniye`;
  }
  return `${sec.toFixed(1).replace('.', ',')} saniye`;
};

/**
 * Tepe değerini parametre birimine göre formatlar (örn: "105 Tork" veya "%104,5 Ng")
 */
export const formatPeakValue = (colKey: string, val: number): string => {
  if (colKey === 'torq') {
    if (val <= 150) {
      return `${Math.round(val)} Tork`;
    }
    return `${Math.round(val)} FT-LB Tork`;
  }
  if (colKey === 'ng') {
    return `%${val.toFixed(1).replace('.', ',')} Ng`;
  }
  if (colKey === 'itt') {
    return `${Math.round(val)}°C ITT`;
  }
  if (colKey === 'np') {
    return `${Math.round(val)} RPM Np`;
  }
  if (colKey === 'fuel_p') {
    return `${val.toFixed(1).replace('.', ',')} PSI Yakıt Bas.`;
  }
  if (colKey === 'oil_p') {
    return `${val.toFixed(1).replace('.', ',')} PSI Yağ Bas.`;
  }
  if (colKey === 'oil_t') {
    return `${val.toFixed(1).replace('.', ',')}°C Yağ Sıcak.`;
  }
  if (colKey === 'fflow') {
    return `${Math.round(val)} L/h Akış`;
  }
  if (colKey === 'fuel_l') {
    return `${Math.round(val)} L Sol Yakıt`;
  }
  if (colKey === 'fuel_r') {
    return `${Math.round(val)} L Sağ Yakıt`;
  }
  if (colKey === 'fuel_imbalance') {
    return `${Math.round(val)} L Dengesizlik`;
  }
  return `${val}`;
};

export const MvpFlightDataAnalysisModal: React.FC<MvpFlightDataAnalysisModalProps> = ({
  isOpen,
  onClose,
  onAddToOlayTakip,
  showNotification,
  defaultTailNo
}) => {
  // AT-802 Uçak Listesi (Depo envanterinin birebir aynısı)
  const [selectedTail, setSelectedTail] = useState<At802AircraftInfo>(() => {
    if (defaultTailNo) {
      const match = AT802_FLEET.find(
        f => f.tail.toLowerCase() === defaultTailNo.toLowerCase() ||
             f.code.toLowerCase() === defaultTailNo.toLowerCase() ||
             f.fullLabel.toLowerCase().includes(defaultTailNo.toLowerCase())
      );
      if (match) return match;
    }
    return AT802_FLEET[1] || AT802_FLEET[0]; // Varsayılan ORMAN 22 (OR-2022) - AT-802
  });

  // Dosya ve Analiz State'leri
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [isDragging, setIsDragging] = useState(false);
  const [parsingProgress, setParsingProgress] = useState(0); // 0 to 100%
  const [parsingStatusText, setParsingStatusText] = useState('');
  const [analysisResult, setAnalysisResult] = useState<MvpAnalysisResult | null>(null);

  // Çoklu Uçuş Analiz Modu (Kullanıcı İsteği: Varsayılan olarak tikli)
  const [isMultiFlightMode, setIsMultiFlightMode] = useState<boolean>(true);
  const [multiFlights, setMultiFlights] = useState<MultiFlightItem[]>([]);
  const [activeFlightIndex, setActiveFlightIndex] = useState<number>(0);
  const [isFlightDropdownOpen, setIsFlightDropdownOpen] = useState<boolean>(false);

  // Şifre Doğrulama Modal State'leri ("Sarı İkazları Saydır" Butonu İçin Şifre: 802)
  const [isPasswordModalOpen, setIsPasswordModalOpen] = useState<boolean>(false);
  const [passwordInput, setPasswordInput] = useState<string>('');
  const [passwordError, setPasswordError] = useState<string | null>(null);

  // Mükerrer Dosya Adı Uyarısı Modal State'leri
  const [isDuplicateWarningOpen, setIsDuplicateWarningOpen] = useState<boolean>(false);
  const [duplicateWarningFileName, setDuplicateWarningFileName] = useState<string>('');

  // 3 Adımlı Analiz Akışı State'leri:
  // 1. Adım: 7 Parametreden seçilen parametre anahtarı (İlk açılışta null -> SADECE 1. Kısım görünür)
  const [selectedParamKey, setSelectedParamKey] = useState<string | null>(null);

  // 2. Adım: Tespit edilen limit aşım aralıklarından seçilen olay ID'si (Seçildiğinde 3. Kısım açılır)
  const [selectedEpisodeId, setSelectedEpisodeId] = useState<string | null>(null);

  // Tablo Filtreleri & Sayfalama
  const [filterMode, setFilterMode] = useState<'ALL' | 'VIOLATIONS_ONLY' | 'RED_ONLY' | 'YELLOW_ONLY'>('VIOLATIONS_ONLY');
  const [currentPage, setCurrentPage] = useState(1);
  const [pageSize, setPageSize] = useState(50);
  const [showOnlyYellowColumns, setShowOnlyYellowColumns] = useState(false);
  const [sortMode, setSortMode] = useState<'CHRONO' | 'SEVERITY'>('CHRONO');

  // 3. Kısım: Sütun Filtreleme Modu (SIRA, İKAZ, TIME ve Seçilen Parametre)
  const [tableColumnFilterMode, setTableColumnFilterMode] = useState<'SELECTED_ONLY' | 'ALL_YELLOW' | 'ALL_COLUMNS'>('SELECTED_ONLY');

  // AT-802F Motor Uçuş Grafiği & Oynatıcı Modal State'leri
  const [isFlightChartModalOpen, setIsFlightChartModalOpen] = useState(false);
  const [chartInitialParamKey, setChartInitialParamKey] = useState<string>('ng');

  // Analiz Ekranı Görüntüsü İndirme State'leri
  const analysisModalContentRef = useRef<HTMLDivElement>(null);
  const [isCapturingAnalysis, setIsCapturingAnalysis] = useState(false);

  const handleDownloadAnalysisScreenshot = async () => {
    if (!analysisModalContentRef.current || isCapturingAnalysis) return;
    setIsCapturingAnalysis(true);
    try {
      const canvas = await html2canvas(analysisModalContentRef.current, {
        backgroundColor: '#ffffff',
        scale: 1.5,
        useCORS: true,
        logging: false
      });
      const dataUrl = canvas.toDataURL('image/png');
      const link = document.createElement('a');
      link.href = dataUrl;
      const cleanTail = selectedTail.code || 'AT802';
      link.download = `AT802_${cleanTail}_Analiz_Raporu_${new Date().toISOString().slice(0, 10)}.png`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    } catch (err) {
      console.error('Screenshot error:', err);
      alert('Ekran görüntüsü alınamadı: ' + (err instanceof Error ? err.message : String(err)));
    } finally {
      setIsCapturingAnalysis(false);
    }
  };

  // Akordiyon UI toggle'ları
  const [isMetadataExpanded, setIsMetadataExpanded] = useState(false);
  const [isRefTableExpanded, setIsRefTableExpanded] = useState(false);

  // Sarı İkaz Takip Adedi & Kayıtları State'i (Her hava aracı için)
  const [yellowNgCounts, setYellowNgCounts] = useState<Record<string, number>>({});
  const [yellowNgRecords, setYellowNgRecords] = useState<any[]>([]);
  const [isYellowNgHistoryModalOpen, setIsYellowNgHistoryModalOpen] = useState(false);
  const [isSyncingYellowNg, setIsSyncingYellowNg] = useState(false);
  const [isLoadingYellowNg, setIsLoadingYellowNg] = useState(false);
  const [lastYellowNgFetchTime, setLastYellowNgFetchTime] = useState<string | null>(null);

  // Google E-Tablodan canlı sarı ikaz satırlarını ve adetlerini çek (JSON dosyası değil, doğrudan e-tablo)
  const fetchLiveYellowNg = useCallback(async (notify = false) => {
    setIsLoadingYellowNg(true);
    let success = false;
    try {
      // 1. Doğrudan API proxy üzerinden Google E-Tablodan canlı oku (cache-busting ile)
      const res = await fetch(`/api/yellow-ng?_t=${Date.now()}`, {
        headers: {
          'Cache-Control': 'no-cache, no-store, must-revalidate',
          'Pragma': 'no-cache'
        }
      });
      if (res.ok) {
        const data = await res.json();
        if (data && data.status === 'success') {
          if (data.data?.counts) setYellowNgCounts(data.data.counts);
          if (data.data?.records) setYellowNgRecords(data.data.records);
          const now = new Date();
          const timeStr = `${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}:${String(now.getSeconds()).padStart(2, '0')}`;
          setLastYellowNgFetchTime(timeStr);
          success = true;
          if (notify) {
            showNotification(`Google E-Tablodan ("Sarı İkaz Ng Takip Adedi") ${data.data?.records?.length || 0} kayıt başarıyla güncellendi.`, 'success');
          }
        }
      }
    } catch (err) {
      console.warn('API /api/yellow-ng warn, falling back to direct Google Apps Script:', err);
    }

    // 2. Yedek: Sunucu proxy erişilemezse doğrudan Google Apps Script ile canlı E-Tablo oku
    if (!success) {
      const fallbackUrls = [
        'https://script.google.com/macros/s/AKfycbwP1uOo2NrST5a4I8vm1nGBLtI26yY2lWrmu9_e9iymwBkUhJBA9JOPCp7SNKqJbOubOw/exec',
        'https://script.google.com/macros/s/AKfycbw4kruTTc058Y9rLTyO3dKi6KloYsmdDTwV1GSiAk8ZXefyo3Z7_VDSTuurzsS9BHAQyQ/exec'
      ];
      for (const sUrl of fallbackUrls) {
        try {
          const directUrl = `${sUrl}?action=readSheet&sheetName=${encodeURIComponent('Sarı İkaz Ng Takip Adedi')}&spreadsheetId=17ScGYYx0erzDwHDk6RGiHOdJATdfmmExXFBY39dXpF0&_t=${Date.now()}`;
          const gRes = await fetch(directUrl);
          if (gRes.ok) {
            const gData = await gRes.json();
            if (gData && gData.status === 'success' && Array.isArray(gData.data)) {
              const mapped = gData.data.map((r: any, idx: number) => ({
                id: String(r['#'] || r.id || (idx + 1)).trim(),
                '#': String(r['#'] || r.id || (idx + 1)).trim(),
                tailCode: String(r['Kuyruk No'] || r.tailCode || '').trim(),
                'Kuyruk No': String(r['Kuyruk No'] || r.tailCode || '').trim(),
                motorSn: String(r['Motor S/N'] || r.motorSn || '').trim(),
                'Motor S/N': String(r['Motor S/N'] || r.motorSn || '').trim(),
                olayNo: String(r['Olay Numarası'] || r.olayNo || '').trim(),
                'Olay Numarası': String(r['Olay Numarası'] || r.olayNo || '').trim(),
                zamanAraligi: String(r['Zaman Aralığı (TIME)'] || r.zamanAraligi || '').trim(),
                'Zaman Aralığı (TIME)': String(r['Zaman Aralığı (TIME)'] || r.zamanAraligi || '').trim(),
                olayTepesi: String(r['Olay Tepesi'] || r.olayTepesi || '').trim(),
                'Olay Tepesi': String(r['Olay Tepesi'] || r.olayTepesi || '').trim(),
                satirAraligi: String(r['Satır Aralığı'] || r.satirAraligi || '').trim(),
                'Satır Aralığı': String(r['Satır Aralığı'] || r.satirAraligi || '').trim(),
                sure: String(r['Süre'] || r.sure || '').trim(),
                'Süre': String(r['Süre'] || r.sure || '').trim(),
                sistemKayitTarihi: String(r['Sisteme Kayıt Tarihi'] || r.sistemKayitTarihi || '').trim(),
                'Sisteme Kayıt Tarihi': String(r['Sisteme Kayıt Tarihi'] || r.sistemKayitTarihi || '').trim(),
                fileName: String(r['Dosya Adı'] || r.fileName || '').trim(),
                'Dosya Adı': String(r['Dosya Adı'] || r.fileName || '').trim()
              }));
              const counts: Record<string, number> = {};
              mapped.forEach((r: any) => {
                if (r.tailCode) counts[r.tailCode] = (counts[r.tailCode] || 0) + 1;
              });
              setYellowNgRecords(mapped);
              setYellowNgCounts(counts);
              const now = new Date();
              setLastYellowNgFetchTime(`${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}:${String(now.getSeconds()).padStart(2, '0')}`);
              success = true;
              if (notify) {
                showNotification(`Doğrudan Google Apps Script ile E-Tablodan ${mapped.length} satır çekildi.`, 'success');
              }
              break;
            }
          }
        } catch (e) {
          console.warn('Direct GAS read error:', e);
        }
      }
    }
    setIsLoadingYellowNg(false);
  }, [showNotification]);

  // Modal açıldığında veya hava aracı değiştiğinde E-Tablodan canlı veri çek
  useEffect(() => {
    if (isOpen) {
      fetchLiveYellowNg();
    }
  }, [isOpen, selectedTail.code, fetchLiveYellowNg]);

  // Kayıtlı Ng Sarı İkaz Satırları geçmiş penceresi açıldığında daima en taze E-Tablo verisini çek
  useEffect(() => {
    if (isYellowNgHistoryModalOpen) {
      fetchLiveYellowNg();
    }
  }, [isYellowNgHistoryModalOpen, fetchLiveYellowNg]);

  // Ng parametresindeki sarı ikaz epizotları (analiz edilmiş mevcut dosya için)
  const currentNgYellowEpisodes = useMemo(() => {
    if (!analysisResult) return [];
    return analysisResult.episodes.filter(ep => {
      const isNg = ep.colKey === 'ng' || ep.colName.toUpperCase().includes('NG');
      return isNg && ep.level === 'YELLOW';
    });
  }, [analysisResult]);

  // Sarı İkazları Saydır Butonuna basıldığında şifre sorma modalını aç
  const handleInitiateCountYellowNg = () => {
    if (isSyncingYellowNg) return;

    if (!analysisResult) {
      showNotification('Lütfen önce bir MVP/CSV uçuş kayıt dosyası seçip yükleyiniz.', 'error');
      return;
    }

    if (currentNgYellowEpisodes.length === 0) {
      showNotification(`${selectedTail.code} için bu uçuş kaydında Ng (%) parametresinde sarı ikaz tespit edilmemiştir.`, 'info');
      return;
    }

    setPasswordInput('');
    setPasswordError(null);
    setIsPasswordModalOpen(true);
  };

  // Şifre Doğrulandıktan sonra Mükerrer Kontrolü yap ve Aktarımı Gerçekleştir
  const handleVerifyPasswordAndTransfer = async () => {
    if (passwordInput.trim() !== '802') {
      setPasswordError('Hatalı yetkili şifresi! Doğru şifre: 802');
      return;
    }

    // Şifre doğru -> modalı kapat
    setIsPasswordModalOpen(false);
    setPasswordInput('');
    setPasswordError(null);

    // MÜKERRER ARAMASI (Aynı Hava Aracı ve Aynı Dosya Adı kontrolü)
    const currentFileName = (analysisResult?.fileName || '').trim().toLowerCase();
    const currentTail = (selectedTail?.code || '').trim().toUpperCase();
    const isDuplicate = yellowNgRecords.some(r => {
      const fn = String(r.fileName || r["Dosya Adı"] || "").trim().toLowerCase();
      const tail = String(r.tailCode || r["Kuyruk No"] || "").trim().toUpperCase();
      // Farklı hava aracı ise mükerrer sayılmaz!
      return fn && fn === currentFileName && tail === currentTail;
    });

    if (isDuplicate) {
      setDuplicateWarningFileName(analysisResult?.fileName || 'Uçuş Dosyası');
      setIsDuplicateWarningOpen(true);
      return;
    }

    // Mükerrer yok -> Aktarımı Başlat
    await executeActualCountYellowNg();
  };

  // Gerçek Google E-Tabloya Sarı İkaz Aktarımı
  const executeActualCountYellowNg = async () => {
    if (!analysisResult) return;
    setIsSyncingYellowNg(true);

    const now = new Date();
    const dateFormatted = `${String(now.getDate()).padStart(2, '0')}.${String(now.getMonth() + 1).padStart(2, '0')}.${now.getFullYear()} ${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}:${String(now.getSeconds()).padStart(2, '0')}`;

    // Her bir sarı ikaz için bir satır
    const newRows = currentNgYellowEpisodes.map((ep, idx) => ({
      olayNo: `Olay #${idx + 1} (${ep.id})`,
      tailCode: selectedTail.code,
      tailFull: selectedTail.fullLabel,
      motorSn: selectedTail.motorSn,
      zamanAraligi: `${ep.startTimeStr} ➔ ${ep.endTimeStr}`,
      olayTepesi: ep.peakValueFormatted || `%${ep.peakValue.toFixed(1).replace('.', ',')} Ng`,
      satirAraligi: `Satır #${ep.startRow} - #${ep.endRow} (${ep.rowCount} Satır)`,
      sure: ep.durationFormatted || `${ep.durationSeconds.toFixed(1)} saniye`,
      sistemKayitTarihi: dateFormatted,
      fileName: analysisResult.fileName
    }));

    let success = false;
    const fallbackScriptUrls = [
      'https://script.google.com/macros/s/AKfycbwP1uOo2NrST5a4I8vm1nGBLtI26yY2lWrmu9_e9iymwBkUhJBA9JOPCp7SNKqJbOubOw/exec',
      'https://script.google.com/macros/s/AKfycbw4kruTTc058Y9rLTyO3dKi6KloYsmdDTwV1GSiAk8ZXefyo3Z7_VDSTuurzsS9BHAQyQ/exec'
    ];
    const spreadsheetId = '17ScGYYx0erzDwHDk6RGiHOdJATdfmmExXFBY39dXpF0';
    const sheetName = 'Sarı İkaz Ng Takip Adedi';

    try {
      const res = await fetch('/api/yellow-ng', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          tailCode: selectedTail.code,
          records: newRows
        })
      });

      if (res.ok) {
        const data = await res.json();
        if (res.status === 409 || data?.status === 'duplicate') {
          if (data?.data?.counts) setYellowNgCounts(data.data.counts);
          if (data?.data?.records) setYellowNgRecords(data.data.records);
          setDuplicateWarningFileName(analysisResult.fileName);
          setIsDuplicateWarningOpen(true);
          setIsSyncingYellowNg(false);
          return;
        }
        if (data?.status === 'success') {
          if (data.data?.counts) setYellowNgCounts(data.data.counts);
          if (data.data?.records) setYellowNgRecords(data.data.records);
          const updatedTotal = data.data?.counts?.[selectedTail.code] || 0;
          showNotification(`${selectedTail.code} için ${newRows.length} adet Ng Sarı İkazı Google E-Tabloya ("${sheetName}") başarıyla aktarıldı! (Toplam: ${updatedTotal})`, 'success');
          
          // Çoklu uçuş listesinde aktarıldı olarak işaretle
          setMultiFlights(prev => prev.map(f => f.fileName === analysisResult.fileName ? { ...f, hasTransferredYellow: true } : f));
          success = true;
          setIsSyncingYellowNg(false);
          return;
        }
      } else if (res.status === 409) {
        setDuplicateWarningFileName(analysisResult.fileName);
        setIsDuplicateWarningOpen(true);
        setIsSyncingYellowNg(false);
        return;
      }
    } catch (apiErr) {
      console.warn('API proxy error, falling back to direct Google Apps Script:', apiErr);
    }

    if (!success) {
      for (const sUrl of fallbackScriptUrls) {
        try {
          const readUrl = `${sUrl}?action=readSheet&sheetName=${encodeURIComponent(sheetName)}&spreadsheetId=${spreadsheetId}&_t=${Date.now()}`;
          const rRes = await fetch(readUrl);
          if (rRes.ok) {
            const rJson = await rRes.json();
            const existing = Array.isArray(rJson.data) ? rJson.data : [];

            // Direct check for duplicate filename on the same aircraft
            const currentTail = (selectedTail?.code || '').trim().toUpperCase();
            const hasDuplicateInSheet = existing.some((r: any) => {
              const fn = String(r.fileName || r["Dosya Adı"] || r[9] || "").trim().toLowerCase();
              const tail = String(r.tailCode || r["Kuyruk No"] || r[1] || "").trim().toUpperCase();
              return fn && fn === analysisResult.fileName.trim().toLowerCase() && tail === currentTail;
            });

            if (hasDuplicateInSheet) {
              setDuplicateWarningFileName(analysisResult.fileName);
              setIsDuplicateWarningOpen(true);
              setIsSyncingYellowNg(false);
              return;
            }

            const combined = [...existing, ...newRows];
            const headers = ["#", "Kuyruk No", "Motor S/N", "Olay Numarası", "Zaman Aralığı (TIME)", "Olay Tepesi", "Satır Aralığı", "Süre", "Sisteme Kayıt Tarihi", "Dosya Adı"];
            const rows = [
              headers,
              ...combined.map((r: any, idx: number) => [
                String(idx + 1),
                String(r.tailCode || r["Kuyruk No"] || "").trim(),
                String(r.motorSn || r["Motor S/N"] || "").trim(),
                String(r.olayNo || r["Olay Numarası"] || "").trim(),
                String(r.zamanAraligi || r["Zaman Aralığı (TIME)"] || "").trim(),
                String(r.olayTepesi || r["Olay Tepesi"] || "").trim(),
                String(r.satirAraligi || r["Satır Aralığı"] || "").trim(),
                String(r.sure || r["Süre"] || "").trim(),
                String(r.sistemKayitTarihi || r["Sisteme Kayıt Tarihi"] || "").trim(),
                String(r.fileName || r["Dosya Adı"] || "").trim()
              ])
            ];

            const pRes = await fetch(sUrl, {
              method: 'POST',
              headers: { 'Content-Type': 'text/plain;charset=utf-8' },
              body: JSON.stringify({
                action: 'updateSheet',
                sheetName,
                spreadsheetId,
                data: rows
              })
            });
            if (pRes.ok) {
              const pJson = await pRes.json().catch(() => null);
              if (pJson && pJson.status === 'success') {
                await fetchLiveYellowNg(false);
                showNotification(`${selectedTail.code} için ${newRows.length} adet Ng Sarı İkazı Google E-Tabloya doğrudan aktarıldı!`, 'success');
                setMultiFlights(prev => prev.map(f => f.fileName === analysisResult.fileName ? { ...f, hasTransferredYellow: true } : f));
                success = true;
                break;
              }
            }
          }
        } catch (scriptErr) {
          console.warn('Direct GAS write error:', scriptErr);
        }
      }
    }

    if (!success) {
      showNotification('Google E-Tabloya veri aktarılamadı. Lütfen internet bağlantınızı kontrol edin.', 'error');
    }
    setIsSyncingYellowNg(false);
  };

  const handleDeleteYellowNgRecord = async (recordId: string | number) => {
    let success = false;
    const fallbackScriptUrls = [
      'https://script.google.com/macros/s/AKfycbwP1uOo2NrST5a4I8vm1nGBLtI26yY2lWrmu9_e9iymwBkUhJBA9JOPCp7SNKqJbOubOw/exec',
      'https://script.google.com/macros/s/AKfycbw4kruTTc058Y9rLTyO3dKi6KloYsmdDTwV1GSiAk8ZXefyo3Z7_VDSTuurzsS9BHAQyQ/exec'
    ];
    const spreadsheetId = '17ScGYYx0erzDwHDk6RGiHOdJATdfmmExXFBY39dXpF0';
    const sheetName = 'Sarı İkaz Ng Takip Adedi';

    try {
      const res = await fetch(`/api/yellow-ng/${recordId}`, { method: 'DELETE' });
      if (res.ok) {
        const data = await res.json();
        if (data && data.status === 'success') {
          if (data.data?.counts) setYellowNgCounts(data.data.counts);
          if (data.data?.records) setYellowNgRecords(data.data.records);
          showNotification('Kayıtlı sarı ikaz satırı Google E-Tablodan silindi.', 'success');
          success = true;
          return;
        }
      }
    } catch (e) {
      console.warn('API delete error, falling back to direct GAS:', e);
    }

    if (!success) {
      for (const sUrl of fallbackScriptUrls) {
        try {
          const readUrl = `${sUrl}?action=readSheet&sheetName=${encodeURIComponent(sheetName)}&spreadsheetId=${spreadsheetId}&_t=${Date.now()}`;
          const rRes = await fetch(readUrl);
          if (rRes.ok) {
            const rJson = await rRes.json();
            const existing = Array.isArray(rJson.data) ? rJson.data : [];
            const filtered = existing.filter((r: any, idx: number) => {
              const rowId = String(r.id || r["#"] || (idx + 1)).trim();
              return rowId !== String(recordId).trim();
            });
            const headers = ["#", "Kuyruk No", "Motor S/N", "Olay Numarası", "Zaman Aralığı (TIME)", "Olay Tepesi", "Satır Aralığı", "Süre", "Sisteme Kayıt Tarihi", "Dosya Adı"];
            const rows = [
              headers,
              ...filtered.map((r: any, idx: number) => [
                String(idx + 1),
                String(r.tailCode || r["Kuyruk No"] || "").trim(),
                String(r.motorSn || r["Motor S/N"] || "").trim(),
                String(r.olayNo || r["Olay Numarası"] || "").trim(),
                String(r.zamanAraligi || r["Zaman Aralığı (TIME)"] || "").trim(),
                String(r.olayTepesi || r["Olay Tepesi"] || "").trim(),
                String(r.satirAraligi || r["Satır Aralığı"] || "").trim(),
                String(r.sure || r["Süre"] || "").trim(),
                String(r.sistemKayitTarihi || r["Sisteme Kayıt Tarihi"] || "").trim(),
                String(r.fileName || r["Dosya Adı"] || "").trim()
              ])
            ];
            const pRes = await fetch(sUrl, {
              method: 'POST',
              headers: { 'Content-Type': 'text/plain;charset=utf-8' },
              body: JSON.stringify({
                action: 'updateSheet',
                sheetName,
                spreadsheetId,
                data: rows
              })
            });
            if (pRes.ok) {
              await fetchLiveYellowNg(false);
              showNotification('Kayıtlı sarı ikaz satırı Google E-Tablodan silindi.', 'success');
              success = true;
              break;
            }
          }
        } catch (err) {
          console.warn('Direct GAS delete error:', err);
        }
      }
    }

    if (!success) {
      showNotification('Kayıt silinirken bağlantı hatası oluştu.', 'error');
    }
  };

  // Olay Takip Çizelgesine Aktarım Formu
  const [isIntegrateModalOpen, setIsIntegrateModalOpen] = useState(false);
  const [integrateForm, setIntegrateForm] = useState({
    kuyrukNo: selectedTail.code,
    tarih: new Date().toLocaleDateString('tr-TR'),
    motorSeriNo: selectedTail.motorSn,
    olayTuru: '',
    tepeNg: '',
    daaRefNo: '',
    verilenSart: '',
    pilot: '',
    teknikDegerlendirme: '',
    parcaDurumu: 'Onarıma Alındı'
  });

  const fileInputRef = useRef<HTMLInputElement | null>(null);

  // Anlık (süre koşulunu sağlamayan) yüksek değer uyarı penceresi state'i
  const [transientWarningPopup, setTransientWarningPopup] = useState<{
    isOpen: boolean;
    colKey: string;
    colName: string;
    peakVal: number;
    peakValFormatted: string;
    row: ParsedMvpRow | null;
  } | null>(null);
  const [dontShowTransientAgain, setDontShowTransientAgain] = useState<Record<string, boolean>>({});

  // 1. Kısım 7 Parametre En Yüksek Değerler ve Aşım İstatistikleri Yardımcısı
  const getParamData = (key: string) => {
    if (!analysisResult) return { peak: '-', limit: '', total: 0, red: 0, yellow: 0 };
    let peak = '-';
    let limit = '';

    if (key === 'ng') {
      peak = analysisResult.peakNg !== null ? `%${analysisResult.peakNg.toFixed(1).replace('.', ',')}` : '-';
      limit = analysisResult.peakNg !== null ? `Dosya Geneli Tepe: %${analysisResult.peakNg.toFixed(1).replace('.', ',')}` : 'Limit: %104 (20s)';
    } else if (key === 'itt') {
      peak = analysisResult.peakItt !== null ? `${analysisResult.peakItt.toFixed(1).replace('.', ',')}°C` : '-';
      limit = analysisResult.peakItt !== null ? `Dosya Geneli Tepe: ${analysisResult.peakItt.toFixed(1).replace('.', ',')}°C` : 'Limit: 870°C (20s)';
    } else if (key === 'np') {
      peak = analysisResult.peakNp !== null ? `${analysisResult.peakNp.toFixed(0)} RPM` : '-';
      limit = analysisResult.peakNp !== null ? `Dosya Geneli Tepe: ${analysisResult.peakNp.toFixed(0)} RPM` : 'Limit: 1700 RPM (20s)';
    } else if (key === 'torq') {
      peak = analysisResult.peakTq !== null ? `${analysisResult.peakTq.toFixed(0)} FT-LB` : '-';
      limit = analysisResult.peakTq !== null ? `Dosya Geneli Tepe: ${analysisResult.peakTq.toFixed(0)} LB/FT` : 'Limit: 5252 LB/FT';
    } else if (key === 'fuel_p') {
      peak = analysisResult.minFuelP !== null ? `${analysisResult.minFuelP.toFixed(1)} PSI` : (analysisResult.maxFuelP !== null ? `${analysisResult.maxFuelP.toFixed(1)} PSI` : '-');
      limit = 'Normal: 10 - 25 PSI';
    } else if (key === 'oil_p') {
      peak = analysisResult.minOilP !== null ? `${analysisResult.minOilP.toFixed(1)} PSI` : (analysisResult.maxOilP !== null ? `${analysisResult.maxOilP.toFixed(1)} PSI` : '-');
      limit = 'Normal: 90 - 135 PSI';
    } else if (key === 'oil_t') {
      peak = analysisResult.maxOilT !== null ? `${analysisResult.maxOilT.toFixed(1).replace('.', ',')}°C` : '-';
      limit = analysisResult.maxOilT !== null ? `Dosya Geneli Tepe: ${analysisResult.maxOilT.toFixed(1).replace('.', ',')}°C` : 'Normal: -40-105°C';
    }

    const eps = analysisResult.episodes.filter(ep => {
      if (ep.colKey === key) return true;
      const up = ep.colName.toUpperCase();
      if (key === 'ng') return up.includes('NG');
      if (key === 'itt') return up.includes('ITT');
      if (key === 'np') return up.includes('NP') || up.includes('RPM');
      if (key === 'torq') return up.includes('TORQ') || up.includes('TQ');
      if (key === 'fuel_p') return up.includes('FUEL P');
      if (key === 'oil_p') return up.includes('OIL P');
      if (key === 'oil_t') return up.includes('OIL T');
      return false;
    });

    const red = eps.filter(e => e.level === 'RED').length;
    const yellow = eps.filter(e => e.level === 'YELLOW').length;
    
    let peakValNumeric = 0;
    if (key === 'ng') peakValNumeric = analysisResult.peakNg || 0;
    else if (key === 'itt') peakValNumeric = analysisResult.peakItt || 0;
    else if (key === 'np') peakValNumeric = analysisResult.peakNp || 0;
    else if (key === 'torq') peakValNumeric = analysisResult.peakTq || 0;
    else if (key === 'oil_t') peakValNumeric = analysisResult.maxOilT || 0;

    return { peak, limit, total: eps.length, red, yellow, peakValNumeric };
  };

  // 2. Kısım: Seçili parametreye ait tespit edilen limit aşım aralıkları (episodes)
  const currentParamEpisodes = useMemo(() => {
    if (!analysisResult || !selectedParamKey) return [];
    return analysisResult.episodes.filter(ep => {
      if (ep.colKey === selectedParamKey) return true;
      const up = ep.colName.toUpperCase();
      if (selectedParamKey === 'ng') return up.includes('NG');
      if (selectedParamKey === 'itt') return up.includes('ITT');
      if (selectedParamKey === 'np') return up.includes('NP') || up.includes('RPM');
      if (selectedParamKey === 'torq') return up.includes('TORQ') || up.includes('TQ');
      if (selectedParamKey === 'fuel_p') return up.includes('FUEL P');
      if (selectedParamKey === 'oil_p') return up.includes('OIL P');
      if (selectedParamKey === 'oil_t') return up.includes('OIL T');
      return false;
    });
  }, [analysisResult, selectedParamKey]);

  // Seçili parametre konfigürasyonu
  const selectedColConfig = useMemo(() => {
    if (!selectedParamKey) return null;
    return YELLOW_FOCUS_COLUMNS.find(c => c.key === selectedParamKey) || null;
  }, [selectedParamKey]);

  // 3. Kısım: Seçili olay (aralık) detayları
  const selectedEpisode = useMemo(() => {
    if (!analysisResult || !selectedEpisodeId || selectedEpisodeId === 'ALL_PARAM_ROWS') return null;
    return analysisResult.episodes.find(e => e.id === selectedEpisodeId) || null;
  }, [analysisResult, selectedEpisodeId]);

  // Sayfalama sıfırlama
  useEffect(() => {
    setCurrentPage(1);
  }, [filterMode, selectedParamKey, pageSize, analysisResult, sortMode, selectedEpisodeId]);

  if (!isOpen) return null;

  // Kuyruk numarası seçimi değiştiğinde
  const handleTailChange = (newFullLabel: string) => {
    const found = AT802_FLEET.find(f => f.fullLabel === newFullLabel);
    if (!found) return;
    setSelectedTail(found);
    setIntegrateForm(prev => ({
      ...prev,
      kuyrukNo: found.code,
      motorSeriNo: found.motorSn
    }));
    if (analysisResult) {
      setAnalysisResult(prev => prev ? {
        ...prev,
        tailNo: found.fullLabel,
        motorSn: found.motorSn
      } : null);
    }
  };

  // Dosya seçildiğinde doğrudan ve otomatik analiz başlat (Tekli veya Çoklu Uçuş)
  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      processIncomingFiles(e.target.files);
    }
  };

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      processIncomingFiles(e.dataTransfer.files);
    }
  };

  /**
   * Ultra-Hızlı ve Virgül Ayırma Hatalarını Tamamen Çözen Asenkron Analiz Motoru
   */
  const executeFastAnalysis = async (
    content: string,
    fileName: string,
    currentTailInfo: At802AircraftInfo,
    isBatch = false
  ): Promise<MvpAnalysisResult | null> => {
    const startTime = performance.now();
    setIsAnalyzing(true);
    setParsingProgress(5);
    setParsingStatusText('Dosya okunuyor ve satırlara ayrıştırılıyor...');

    await new Promise(r => setTimeout(r, 20));

    const rawLines = content.split(/\r?\n/).map(l => l.trimEnd()).filter(l => l.trim().length > 0);
    const totalLines = rawLines.length;

    if (totalLines === 0) {
      showNotification('Seçilen uçuş dosyası boş!', 'error');
      setIsAnalyzing(false);
      return;
    }

    setParsingProgress(15);
    setParsingStatusText('Başlık satırı ve ayırıcı (; veya , veya tab) tespit ediliyor...');
    await new Promise(r => setTimeout(r, 20));

    let { delimiter: detectedDelimiter, headerLineIndex } = detectDelimiterAndHeader(rawLines);
    const metadataLines = rawLines.slice(0, headerLineIndex);

    const parsedMetadata: Record<string, string> = {};
    metadataLines.forEach(line => {
      const parts = line.split(/[:=]/);
      if (parts.length >= 2) {
        const k = parts[0].trim().toUpperCase().replace(/[^A-Z0-9_]/g, '_');
        const v = parts.slice(1).join(':').trim();
        parsedMetadata[k] = v;
      }
    });

    const headerLineRaw = rawLines[headerLineIndex];
    let headers = parseCsvLineWithQuotes(headerLineRaw, detectedDelimiter);

    // Eğer tespit edilen ayırıcı 5'ten az sütun verdiyse, diğer adayları dene ve detectedDelimiter'ı MUTLAKA güncelle!
    if (headers.length < 5) {
      for (const altDelim of [',', ';', '\t']) {
        if (altDelim === detectedDelimiter) continue;
        const altHeaders = parseCsvLineWithQuotes(headerLineRaw, altDelim);
        if (altHeaders.length > headers.length) {
          headers = altHeaders;
          detectedDelimiter = altDelim;
        }
      }
    }

    // İlk gerçek veri satırını kontrol ederek ayırıcının veri satırlarını da ayırdığından emin ol
    const firstDataRow = rawLines.slice(headerLineIndex + 1).find(l => l.trim().length > 0);
    if (firstDataRow) {
      const dataColsCurrent = parseCsvLineWithQuotes(firstDataRow, detectedDelimiter).length;
      if (dataColsCurrent < 5) {
        for (const altDelim of [',', ';', '\t']) {
          const testCols = parseCsvLineWithQuotes(firstDataRow, altDelim).length;
          if (testCols > dataColsCurrent && testCols >= 5) {
            detectedDelimiter = altDelim;
            headers = parseCsvLineWithQuotes(headerLineRaw, detectedDelimiter);
            break;
          }
        }
      }
    }

    // Başlıkları tırnaklardan arındır ve temizle
    headers = headers.map((h, idx) => {
      const clean = h.trim().replace(/^["']|["']$/g, '');
      if (clean) return clean;
      return OFFICIAL_ALL_MVP_HEADERS[idx] || getExcelColumnLetter(idx);
    });

    // Veri satırındaki sütun sayısı başlık sayısından fazlaysa resmi başlıklarla tamamla
    if (firstDataRow) {
      const sampleColCount = parseCsvLineWithQuotes(firstDataRow, detectedDelimiter).length;
      if (sampleColCount > headers.length && sampleColCount >= 40) {
        const expandedHeaders = [...headers];
        for (let c = headers.length; c < sampleColCount; c++) {
          expandedHeaders.push(OFFICIAL_ALL_MVP_HEADERS[c] || getExcelColumnLetter(c));
        }
        headers = expandedHeaders;
      }
    }

    if (headers.length < 5) {
      headers = OFFICIAL_ALL_MVP_HEADERS;
    }

    const expectedColCount = headers.length;

    const findColIndex = (regex: RegExp, fallbackIdx: number): number => {
      const idx = headers.findIndex(h => regex.test(h.trim()));
      return idx >= 0 ? idx : (fallbackIdx < expectedColCount ? fallbackIdx : -1);
    };

    const timeColIdx = findColIndex(/^time/i, 0);
    const ngColIdx = findColIndex(/^ng/i, 2);
    const ittColIdx = findColIndex(/^itt/i, 15);
    const npColIdx = findColIndex(/^np/i, 3);
    const torqColIdx = findColIndex(/torq|trq/i, 9);
    const fuelPColIdx = findColIndex(/fuel\s*p/i, 10);
    const oilPColIdx = findColIndex(/oil\s*p/i, 11);
    const oilTColIdx = findColIndex(/oil\s*t/i, 16);

    const dataRows = rawLines.slice(headerLineIndex + 1);
    const totalDataRows = dataRows.length;

    setParsingProgress(30);
    setParsingStatusText(`${totalDataRows.toLocaleString('tr-TR')} satır için 7 sarı odak sütunu taranıyor...`);
    await new Promise(r => setTimeout(r, 20));

    let redCount = 0;
    let yellowCount = 0;
    let durationExceededCount = 0;

    let peakNg: number | null = null;
    let peakItt: number | null = null;
    let peakTq: number | null = null;
    let peakNp: number | null = null;
    let minFuelP: number | null = null;
    let maxFuelP: number | null = null;
    let minOilP: number | null = null;
    let maxOilP: number | null = null;
    let minOilT: number | null = null;
    let maxOilT: number | null = null;

    interface ActiveStreak {
      id: string;
      colKey: string;
      colName: string;
      excelLetter: string;
      colNumber: number;
      colIndex: number;
      worstLevel: 'RED' | 'YELLOW';
      startTimeStr: string;
      endTimeStr: string;
      startSec: number;
      endSec: number;
      startRow: number;
      endRow: number;
      peakValue: number;
      rule: string;
      rowIndices: number[];
    }

    const allEpisodes: ExceedanceEpisode[] = [];
    const activeStreaks: Record<string, ActiveStreak | null> = {
      ng: null,
      itt: null,
      np: null,
      torq: null,
      fuel_p: null,
      oil_p: null,
      oil_t: null
    };
    let episodeCounter = 0;

    const parsedRowList: ParsedMvpRow[] = [];
    const chunkSize = 2000;

    const closeStreak = (colKey: string) => {
      const streak = activeStreaks[colKey];
      if (!streak) return;
      activeStreaks[colKey] = null;

      let dur = Math.max(0, streak.endSec - streak.startSec);
      if (dur === 0 && streak.rowIndices.length > 1) {
        dur = (streak.rowIndices.length - 1) * 0.1;
      }
      if (dur === 0) {
        dur = 0.1;
      }

      // Sarı ikazlar için gerekli süre kontrolü (Ng, Np, Torq: 20s; ITT 870-910: 20s, 910-1000: 5s; Oil T: 10 dk)
      // ÖNEMLİ: Sarı değerler süresi ne olursa olsun K KIRMIZI olamaz, sarı kalır! Kırmızı yalnızca değer aşımıyla olur.
      let requiredDur = 0;
      let durationExceeded = false;
      if (streak.worstLevel === 'YELLOW') {
        if (colKey === 'np' || colKey === 'torq') {
          requiredDur = 20;
        } else if (colKey === 'ng') {
          requiredDur = 0;
        } else if (colKey === 'itt') {
          requiredDur = streak.peakValue >= 910 ? 5 : 20;
        } else if (colKey === 'oil_t') {
          requiredDur = 600; // 10 dakika
        }
        durationExceeded = colKey === 'ng' ? true : (dur >= requiredDur);
      } else {
        durationExceeded = dur > 20;
      }

      const durFormatted = formatDurationSeconds(dur);
      const peakFormatted = formatPeakValue(colKey, streak.peakValue);
      const summary = `${durFormatted} ${peakFormatted}`;

      // Süre koşulunu sağlamayan sarı ikazları (durationExceeded === false) epizot olarak kaydetme ve atla!
      const shouldKeep = streak.worstLevel === 'RED' || durationExceeded;

      if (shouldKeep) {
        const episode: ExceedanceEpisode = {
          id: streak.id,
          colKey: streak.colKey,
          colName: streak.colName,
          excelLetter: streak.excelLetter,
          colNumber: streak.colNumber,
          colIndex: streak.colIndex,
          level: streak.worstLevel,
          startTimeStr: streak.startTimeStr,
          endTimeStr: streak.endTimeStr,
          startSec: streak.startSec,
          endSec: streak.endSec,
          durationSeconds: dur,
          durationFormatted: durFormatted,
          startRow: streak.startRow,
          endRow: streak.endRow,
          rowCount: streak.rowIndices.length,
          peakValue: streak.peakValue,
          peakValueFormatted: peakFormatted,
          summaryText: summary,
          rule: `${streak.rule} (Süre: ${durFormatted})`,
          durationExceeded
        };

        allEpisodes.push(episode);

        if (episode.level === 'RED') {
          redCount++;
        } else {
          yellowCount++;
        }

        // Bu olaya ait tüm satırlara süre aralığı ve çerçeve etiketini bağla
        for (const rowIdx of streak.rowIndices) {
          const row = parsedRowList[rowIdx];
          if (row) {
            const v = row.violations.find(viol => viol.colIndex === streak.colIndex || viol.colName === streak.colName);
            if (v) {
              v.episodeId = episode.id;
              v.timeRangeStr = `${episode.startTimeStr} - ${episode.endTimeStr}`;
              v.durationSeconds = dur;
              v.durationFormatted = durFormatted;
              v.peakValueFormatted = peakFormatted;
              v.summaryText = summary;
              v.durationExceeded = durationExceeded;
              if (episode.level === 'RED') v.level = 'RED';
            }
            if (episode.level === 'RED') {
              row.severity = 'RED';
            }
            if (!row.primaryEpisode || episode.level === 'RED') {
              row.primaryEpisode = episode;
            }
          }
        }
      } else {
        // Süre koşulunu sağlamayan sarı ikaz satırlarının ihlal kayıtlarını temizle (normal kabul et)
        for (const rowIdx of streak.rowIndices) {
          const row = parsedRowList[rowIdx];
          if (row) {
            row.violations = row.violations.filter(viol => viol.colIndex !== streak.colIndex && viol.colName !== streak.colName);
            if (row.violations.length === 0) {
              row.severity = 'NORMAL';
            }
          }
        }
      }
    };

    for (let i = 0; i < totalDataRows; i += chunkSize) {
      const end = Math.min(i + chunkSize, totalDataRows);

      for (let r = i; r < end; r++) {
        const rawLine = dataRows[r];
        if (!rawLine.trim()) continue;

        let cells = parseCsvLineWithQuotes(rawLine, detectedDelimiter);

        if (cells.length > expectedColCount && detectedDelimiter === ',') {
          cells = repairUnquotedDecimalCommas(cells, expectedColCount);
        }

        const timeStr = timeColIdx >= 0 && cells[timeColIdx] ? cells[timeColIdx] : `Satır ${r + 1}`;
        const currentSec = parseTimeToSeconds(timeStr) || (r * 0.1);

        const violations: ViolationDetail[] = [];
        let rowSeverity: 'RED' | 'YELLOW' | 'NORMAL' = 'NORMAL';

        const registerColViolation = (
          colKey: string,
          vInfo: {
            colName: string;
            excelLetter: string;
            colNumber: number;
            colIndex: number;
            val: number;
            rawVal: string;
            level: 'RED' | 'YELLOW';
            rule: string;
          }
        ) => {
          violations.push({
            colName: vInfo.colName,
            colIndex: vInfo.colIndex,
            excelLetter: vInfo.excelLetter,
            colNumber: vInfo.colNumber,
            value: vInfo.val,
            rawVal: vInfo.rawVal,
            level: vInfo.level,
            rule: vInfo.rule
          });

          if (vInfo.level === 'RED') {
            rowSeverity = 'RED';
          } else if (rowSeverity !== 'RED') {
            rowSeverity = 'YELLOW';
          }

          let streak = activeStreaks[colKey];
          if (streak && streak.worstLevel !== vInfo.level) {
            closeStreak(colKey);
            streak = null;
          }
          if (!streak) {
            episodeCounter++;
            streak = {
              id: `ep_${colKey}_${episodeCounter}`,
              colKey,
              colName: vInfo.colName,
              excelLetter: vInfo.excelLetter,
              colNumber: vInfo.colNumber,
              colIndex: vInfo.colIndex,
              worstLevel: vInfo.level,
              startTimeStr: timeStr,
              endTimeStr: timeStr,
              startSec: currentSec,
              endSec: currentSec,
              startRow: r + 1,
              endRow: r + 1,
              peakValue: vInfo.val,
              rule: vInfo.rule,
              rowIndices: [r]
            };
            activeStreaks[colKey] = streak;
          } else {
            streak.endTimeStr = timeStr;
            streak.endSec = currentSec;
            streak.endRow = r + 1;
            if (vInfo.level === 'RED') streak.worstLevel = 'RED';
            if (colKey === 'fuel_p' && vInfo.val < 15) {
              if (vInfo.val < streak.peakValue) streak.peakValue = vInfo.val;
            } else if (colKey === 'oil_t' && vInfo.val < 10) {
              if (vInfo.val < streak.peakValue) streak.peakValue = vInfo.val;
            } else {
              if (vInfo.val > streak.peakValue) streak.peakValue = vInfo.val;
            }
            streak.rowIndices.push(r);
          }
        };

        // 1. NG (%) SÜTUNU: %104 - %105.5 sarı, > %105.5 kırmızı
        if (ngColIdx >= 0 && ngColIdx < cells.length) {
          const val = parseCleanNumber(cells[ngColIdx]);
          const rawVal = cells[ngColIdx]?.trim() || '';
          const intDigits = rawVal.split(/[.,]/)[0].replace(/[^\d]/g, '');
          const isFaultyNg = isNaN(val) || val <= 0 || val >= 1000 || intDigits.length >= 4;

          if (!isFaultyNg) {
            if (peakNg === null || val > peakNg) peakNg = val;

            if (val > 105.5) {
              registerColViolation('ng', {
                colName: 'NG (%)',
                colIndex: ngColIdx,
                excelLetter: 'C',
                colNumber: 3,
                val,
                rawVal: cells[ngColIdx],
                level: 'RED',
                rule: `Ng > %105.5 kritik limit (%${val.toFixed(1).replace('.', ',')})!`
              });
            } else if (val >= 104.0) {
              let activeNgStreak = activeStreaks['ng'];
              let currentDur = 0;
              if (activeNgStreak) {
                currentDur = Math.max(0, currentSec - activeNgStreak.startSec);
              }
              const level = currentDur > 20 ? 'RED' : 'YELLOW';
              const ruleText = level === 'RED'
                ? `Ng %104 - %105,5 aralığı 20 saniyeyi aştı (21. saniye ve sonrası kritik limit: %${val.toFixed(1).replace('.', ',')}).`
                : `Ng %104 - %105,5 sarı ikaz (%${val.toFixed(1).replace('.', ',')}, 0-20s).`;

              registerColViolation('ng', {
                colName: 'NG (%)',
                colIndex: ngColIdx,
                excelLetter: 'C',
                colNumber: 3,
                val,
                rawVal: cells[ngColIdx],
                level,
                rule: ruleText
              });
            } else {
              closeStreak('ng');
            }
          } else {
            closeStreak('ng');
          }
        } else {
          closeStreak('ng');
        }

        // 2. ITT (*C) SÜTUNU: 870-910 (20s) veya 910-1000 (5s) sarı, > 1000 kırmızı
        if (ittColIdx >= 0 && ittColIdx < cells.length) {
          const val = parseCleanNumber(cells[ittColIdx]);
          if (!isNaN(val)) {
            if (peakItt === null || val > peakItt) peakItt = val;

            if (val > 1000) {
              registerColViolation('itt', {
                colName: 'ITT (*C)',
                colIndex: ittColIdx,
                excelLetter: 'P',
                colNumber: 16,
                val,
                rawVal: cells[ittColIdx],
                level: 'RED',
                rule: `ITT > 1000 °C kritik limit (${val.toFixed(1).replace('.', ',')} °C)!`
              });
            } else if (val >= 870) {
              registerColViolation('itt', {
                colName: 'ITT (*C)',
                colIndex: ittColIdx,
                excelLetter: 'P',
                colNumber: 16,
                val,
                rawVal: cells[ittColIdx],
                level: 'YELLOW',
                rule: `ITT 870-1000 °C sarı ikaz (${val.toFixed(1).replace('.', ',')} °C).`
              });
            } else {
              closeStreak('itt');
            }
          } else {
            closeStreak('itt');
          }
        } else {
          closeStreak('itt');
        }

        // 3. NP (RPM) SÜTUNU: 1700 - 1870 sarı ikaz, > 1870 kırmızı aşım
        if (npColIdx >= 0 && npColIdx < cells.length) {
          const val = parseCleanNumber(cells[npColIdx]);
          if (!isNaN(val)) {
            if (peakNp === null || val > peakNp) peakNp = val;

            if (val > 1870) {
              registerColViolation('np', {
                colName: 'NP (RPM)',
                colIndex: npColIdx,
                excelLetter: 'D',
                colNumber: 4,
                val,
                rawVal: cells[npColIdx],
                level: 'RED',
                rule: `Np > 1870 RPM kritik limit (${val.toFixed(0)} RPM)!`
              });
            } else if (val >= 1700) {
              registerColViolation('np', {
                colName: 'NP (RPM)',
                colIndex: npColIdx,
                excelLetter: 'D',
                colNumber: 4,
                val,
                rawVal: cells[npColIdx],
                level: 'YELLOW',
                rule: `Np 1700-1870 RPM sarı ikaz (${val.toFixed(0)} RPM).`
              });
            } else {
              closeStreak('np');
            }
          } else {
            closeStreak('np');
          }
        } else {
          closeStreak('np');
        }

        // 4. TORQ (FT-LB) SÜTUNU: Yeşil: < 5252 LB/FT | Sarı: 5252 - 6092 LB/FT (20 saniye) | Kırmızı: > 6092 LB/FT
        if (torqColIdx >= 0 && torqColIdx < cells.length) {
          const val = parseCleanNumber(cells[torqColIdx]);
          if (!isNaN(val)) {
            if (peakTq === null || val > peakTq) peakTq = val;

            if (val > 6092) {
              registerColViolation('torq', {
                colName: 'TORQ (FT-LB)',
                colIndex: torqColIdx,
                excelLetter: 'J',
                colNumber: 10,
                val,
                rawVal: cells[torqColIdx],
                level: 'RED',
                rule: `Tork > 6092 LB/FT kritik limit (${val.toFixed(0)} LB/FT)!`
              });
            } else if (val >= 5252) {
              registerColViolation('torq', {
                colName: 'TORQ (FT-LB)',
                colIndex: torqColIdx,
                excelLetter: 'J',
                colNumber: 10,
                val,
                rawVal: cells[torqColIdx],
                level: 'YELLOW',
                rule: `Tork 5252 - 6092 LB/FT sarı ikaz (${val.toFixed(0)} LB/FT, 20s).`
              });
            } else {
              closeStreak('torq');
            }
          } else {
            closeStreak('torq');
          }
        } else {
          closeStreak('torq');
        }

        // 5. FUEL P ( PSI) SÜTUNU: Sarı ikaz yok. < 10 veya > 25 direkt kırmızı.
        if (fuelPColIdx >= 0 && fuelPColIdx < cells.length) {
          const val = parseCleanNumber(cells[fuelPColIdx]);
          if (!isNaN(val)) {
            if (minFuelP === null || val < minFuelP) minFuelP = val;
            if (maxFuelP === null || val > maxFuelP) maxFuelP = val;

            if (val < 10.0 || val > 25.0) {
              registerColViolation('fuel_p', {
                colName: 'FUEL P ( PSI)',
                colIndex: fuelPColIdx,
                excelLetter: 'K',
                colNumber: 11,
                val,
                rawVal: cells[fuelPColIdx],
                level: 'RED',
                rule: `Yakıt basıncı < 10 veya > 25 PSI kritik limit (${val.toFixed(1)} PSI)!`
              });
            } else {
              closeStreak('fuel_p');
            }
          } else {
            closeStreak('fuel_p');
          }
        } else {
          closeStreak('fuel_p');
        }

        // 6. OIL P (PSI) SÜTUNU: 60 - 90 sarı, < 60 veya > 135 kırmızı
        if (oilPColIdx >= 0 && oilPColIdx < cells.length) {
          const val = parseCleanNumber(cells[oilPColIdx]);
          if (!isNaN(val)) {
            if (minOilP === null || val < minOilP) minOilP = val;
            if (maxOilP === null || val > maxOilP) maxOilP = val;

            if (val < 60.0 || val > 135.0) {
              registerColViolation('oil_p', {
                colName: 'OIL P (PSI)',
                colIndex: oilPColIdx,
                excelLetter: 'L',
                colNumber: 12,
                val,
                rawVal: cells[oilPColIdx],
                level: 'RED',
                rule: val < 60.0 ? `Yağ basıncı < 60 PSI kritik limit (${val.toFixed(1)} PSI)!` : `Yağ basıncı > 135 PSI kritik limit (${val.toFixed(1)} PSI)!`
              });
            } else if (val >= 60.0 && val < 90.0) {
              registerColViolation('oil_p', {
                colName: 'OIL P (PSI)',
                colIndex: oilPColIdx,
                excelLetter: 'L',
                colNumber: 12,
                val,
                rawVal: cells[oilPColIdx],
                level: 'YELLOW',
                rule: `Yağ basıncı 60 - 90 PSI sarı ikaz (${val.toFixed(1)} PSI).`
              });
            } else {
              closeStreak('oil_p');
            }
          } else {
            closeStreak('oil_p');
          }
        } else {
          closeStreak('oil_p');
        }

        // 7. OIL T (*C) SÜTUNU: 105 - 110 °C (10 dk) sarı, > 110 °C kırmızı
        if (oilTColIdx >= 0 && oilTColIdx < cells.length) {
          const val = parseCleanNumber(cells[oilTColIdx]);
          if (!isNaN(val)) {
            if (minOilT === null || val < minOilT) minOilT = val;
            if (maxOilT === null || val > maxOilT) maxOilT = val;

            if (val > 110.0) {
              registerColViolation('oil_t', {
                colName: 'OIL T (*C)',
                colIndex: oilTColIdx,
                excelLetter: 'Q',
                colNumber: 17,
                val,
                rawVal: cells[oilTColIdx],
                level: 'RED',
                rule: `Yağ sıcaklığı > 110 °C kritik limit (${val.toFixed(1)} °C)!`
              });
            } else if (val >= 105.0 && val <= 110.0) {
              registerColViolation('oil_t', {
                colName: 'OIL T (*C)',
                colIndex: oilTColIdx,
                excelLetter: 'Q',
                colNumber: 17,
                val,
                rawVal: cells[oilTColIdx],
                level: 'YELLOW',
                rule: `Yağ sıcaklığı 105 - 110 °C sarı ikaz (${val.toFixed(1)} °C, 10 dk).`
              });
            } else {
              closeStreak('oil_t');
            }
          } else {
            closeStreak('oil_t');
          }
        } else {
          closeStreak('oil_t');
        }

        parsedRowList.push({
          originalIndex: r + 1,
          timeStr,
          timeSec: currentSec,
          cells,
          violations,
          severity: rowSeverity
        });
      }

      const progress = Math.min(95, Math.round(30 + ((end / totalDataRows) * 65)));
      setParsingProgress(progress);
      setParsingStatusText(`${end.toLocaleString('tr-TR')} / ${totalDataRows.toLocaleString('tr-TR')} satır tarandı...`);
      

      await new Promise(r => setTimeout(r, 0));
    }

    // Dosya sonundaki açık kalan tüm aktif serileri kapat
    Object.keys(activeStreaks).forEach(key => closeStreak(key));

    const elapsed = Math.round(performance.now() - startTime);

    const result: MvpAnalysisResult = {
      fileName,
      tailNo: currentTailInfo.fullLabel,
      motorSn: currentTailInfo.motorSn,
      delimiter: detectedDelimiter,
      metadataLines,
      parsedMetadata,
      headers,
      rows: parsedRowList,
      episodes: allEpisodes,
      redCount,
      yellowCount,
      durationExceededCount,
      peakNg,
      peakItt,
      peakTq,
      peakNp,
      minFuelP,
      maxFuelP,
      minOilP,
      maxOilP,
      minOilT,
      maxOilT,
      processingTimeMs: elapsed
    };

    setAnalysisResult(result);
    setParsingProgress(100);

    const detectedPilot = parsedMetadata['PILOT_NAME'] || 'Sorumlu Pilot';
    const detectedDate = parsedMetadata['DATE_UTC'] || new Date().toLocaleDateString('tr-TR');

    let defaultOlayTuru = 'Normal Uçuş Analizi';
    let defaultTeknik = 'Tüm sarı odak uçuş parametreleri referans limitleri dahilinde.';
    let defaultParca = 'Faal';

    if (redCount > 0) {
      defaultOlayTuru = `KIRMIZI LİMİT AŞIMI (${peakNg ? `Ng %${peakNg.toFixed(1).replace('.', ',')}` : 'Limit Aşımı'})`;
      defaultTeknik = `MVP uçuş kaydında kritik limit aşımı tespit edildi: Tepe Ng: %${peakNg ? peakNg.toFixed(1).replace('.', ',') : '-'}, Tepe ITT: ${peakItt ? peakItt.toFixed(1).replace('.', ',') : '-'}°C, TQ: ${peakTq ? peakTq.toFixed(0) : '-'} LB/FT. ${durationExceededCount > 0 ? `(${durationExceededCount} süre sınırı aşımı tespit edildi).` : ''} P&WC motor kılavuzuna göre DAA ve boroskop muayenesi başlatılmalıdır.`;
      defaultParca = 'Bakım / Kontrol Gerekli';
    } else if (yellowCount > 0) {
      defaultOlayTuru = `SARI UYARI SINIRI AŞIMI (${peakNg ? `Ng %${peakNg.toFixed(1).replace('.', ',')}` : 'Uyarı'})`;
      defaultTeknik = `MVP uçuş kaydı analizi: Sarı uyarı sınırlarında çalışma tespit edildi. Tepe Ng: %${peakNg ? peakNg.toFixed(1).replace('.', ',') : '-'}, Tepe ITT: ${peakItt ? peakItt.toFixed(1).replace('.', ',') : '-'}°C. 20 sn kuralı ve motor parametre trend takibi gerekmektedir.`;
      defaultParca = 'Gözlem Altında Faal';
    }

    if (!isBatch) {
      setAnalysisResult(result);
      setParsingProgress(100);

      setIntegrateForm({
        kuyrukNo: currentTailInfo.code,
        tarih: detectedDate,
        motorSeriNo: currentTailInfo.motorSn,
        olayTuru: defaultOlayTuru,
        tepeNg: peakNg ? `%${peakNg.toFixed(1).replace('.', ',')}` : '%104,0',
        daaRefNo: '',
        verilenSart: '',
        pilot: '',
        teknikDegerlendirme: defaultTeknik,
        parcaDurumu: defaultParca
      });

      if (redCount > 0 || yellowCount > 0) {
        setFilterMode('VIOLATIONS_ONLY');
      } else {
        setFilterMode('ALL');
      }

      // 3 Adımlı Akış: İlk açılışta yalnızca 1. kısım (en yüksek değerler) görünür
      setSelectedParamKey(null);
      setSelectedEpisodeId(null);

      setIsAnalyzing(false);
      showNotification(
        `Yüksek Hızlı Tarama Tamamlandı (${elapsed} ms): ${currentTailInfo.fullLabel} için ${totalLines.toLocaleString('tr-TR')} satır incelendi. (${redCount} Kırmızı Aşım, ${yellowCount} Sarı Uyarı${durationExceededCount > 0 ? `, ${durationExceededCount} Süre Aşımı` : ''})`,
        redCount > 0 ? 'info' : 'success'
      );
    }

    return result;
  };

  // Doğrudan dosya analizi çalıştırma
  const handleRunAnalysis = async (fileToProcess?: File, customTail?: At802AircraftInfo) => {
    try {
      const tailToUse = customTail || selectedTail;
      let content = '';
      let fName = 'flight_log.mvp';

      if (fileToProcess || selectedFile) {
        const file = fileToProcess || selectedFile!;
        fName = file.name;
        content = await file.text();
      } else {
        fileInputRef.current?.click();
        return;
      }

      await executeFastAnalysis(content, fName, tailToUse);
    } catch (err: any) {
      console.error('Fast Analysis Error:', err);
      showNotification('Dosya analiz edilirken hata oluştu: ' + (err?.message || err), 'error');
      setIsAnalyzing(false);
    }
  };

  // Çoklu Dosya Toplu Tarama & Analiz Motoru
  const handleBatchFlightFiles = async (files: File[]) => {
    if (files.length === 0) return;
    setIsAnalyzing(true);
    setParsingProgress(5);
    setParsingStatusText(`Toplam ${files.length} uçuş dosyası taranıyor...`);

    const items: MultiFlightItem[] = [];

    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      setParsingProgress(Math.round(((i + 1) / files.length) * 90));
      setParsingStatusText(`(${i + 1}/${files.length}) "${file.name}" taranıyor...`);
      await new Promise(r => setTimeout(r, 10));

      try {
        const text = await file.text();
        const res = await executeFastAnalysis(text, file.name, selectedTail, true);
        if (res) {
          items.push({
            id: `flight_${i}_${Date.now()}_${Math.random().toString(36).substr(2, 4)}`,
            file,
            fileName: file.name,
            content: text,
            result: res,
            isVisited: i === 0, // İlk uçuş açılışta tıklandı/incelendi olarak yeşil işaretlenir
            yellowCount: res.yellowCount,
            redCount: res.redCount,
            durationExceededCount: res.durationExceededCount
          });
        }
      } catch (err) {
        console.warn(`File scan error for ${file.name}:`, err);
      }
    }

    setIsAnalyzing(false);
    setParsingProgress(100);

    if (items.length > 0) {
      setMultiFlights(items);
      setActiveFlightIndex(0);
      setSelectedFile(items[0].file);
      setAnalysisResult(items[0].result);
      applyAnalysisToForm(items[0].result, selectedTail);
      showNotification(`✅ ${items.length} adet uçuş kaydı başarıyla tarandı. 1. uçuş seçildi.`, 'success');
    } else {
      showNotification('Seçilen uçuş dosyaları taranırken hata oluştu.', 'error');
    }
  };

  // Açılır listeden veya adımdan başka uçuşa geçiş (Her tıklanan yeşil dolgu ve tikli kalır)
  const handleSelectMultiFlight = (index: number) => {
    if (index < 0 || index >= multiFlights.length) return;
    const target = multiFlights[index];
    if (!target || !target.result) return;

    // Tıklanan kaydı yeşil dolgu ve tikli (isVisited: true) olarak güncelle
    setMultiFlights(prev => prev.map((f, i) => i === index ? { ...f, isVisited: true } : f));
    setActiveFlightIndex(index);
    setSelectedFile(target.file);
    setAnalysisResult(target.result);
    applyAnalysisToForm(target.result, selectedTail);
    setIsFlightDropdownOpen(false);
    showNotification(`✈️ ${index + 1}. Uçuş yüklendi: ${target.fileName}`, 'info');
  };

  const applyAnalysisToForm = (res: MvpAnalysisResult, currentTailInfo: At802AircraftInfo) => {
    const detectedPilot = res.parsedMetadata['PILOT_NAME'] || 'Sorumlu Pilot';
    const detectedDate = res.parsedMetadata['DATE_UTC'] || new Date().toLocaleDateString('tr-TR');

    let defaultOlayTuru = 'Normal Uçuş Analizi';
    let defaultTeknik = 'Tüm sarı odak uçuş parametreleri referans limitleri dahilinde.';
    let defaultParca = 'Faal';

    if (res.redCount > 0) {
      defaultOlayTuru = `KIRMIZI LİMİT AŞIMI (${res.peakNg ? `Ng %${res.peakNg.toFixed(1).replace('.', ',')}` : 'Limit Aşımı'})`;
      defaultTeknik = `MVP uçuş kaydında kritik limit aşımı tespit edildi: Tepe Ng: %${res.peakNg ? res.peakNg.toFixed(1).replace('.', ',') : '-'}, Tepe ITT: ${res.peakItt ? res.peakItt.toFixed(1).replace('.', ',') : '-'}°C, TQ: ${res.peakTq ? res.peakTq.toFixed(0) : '-'} LB/FT. ${res.durationExceededCount > 0 ? `(${res.durationExceededCount} süre sınırı aşımı tespit edildi).` : ''} P&WC motor kılavuzuna göre DAA ve boroskop muayenesi başlatılmalıdır.`;
      defaultParca = 'Bakım / Kontrol Gerekli';
    } else if (res.yellowCount > 0) {
      defaultOlayTuru = `SARI UYARI SINIRI AŞIMI (${res.peakNg ? `Ng %${res.peakNg.toFixed(1).replace('.', ',')}` : 'Uyarı'})`;
      defaultTeknik = `MVP uçuş kaydı analizi: Sarı uyarı sınırlarında çalışma tespit edildi. Tepe Ng: %${res.peakNg ? res.peakNg.toFixed(1).replace('.', ',') : '-'}, Tepe ITT: ${res.peakItt ? res.peakItt.toFixed(1).replace('.', ',') : '-'}°C. 20 sn kuralı ve motor parametre trend takibi gerekmektedir.`;
      defaultParca = 'Gözlem Altında Faal';
    }

    setIntegrateForm({
      kuyrukNo: currentTailInfo.code,
      tarih: detectedDate,
      motorSeriNo: currentTailInfo.motorSn,
      olayTuru: defaultOlayTuru,
      tepeNg: res.peakNg ? `%${res.peakNg.toFixed(1).replace('.', ',')}` : '%104,0',
      daaRefNo: '',
      verilenSart: '',
      pilot: '',
      teknikDegerlendirme: defaultTeknik,
      parcaDurumu: defaultParca
    });

    if (res.redCount > 0 || res.yellowCount > 0) {
      setFilterMode('VIOLATIONS_ONLY');
    } else {
      setFilterMode('ALL');
    }
    setSelectedParamKey(null);
    setSelectedEpisodeId(null);
  };

  // Çoklu veya Tekli Dosya Yükleme Karar Motoru
  const processIncomingFiles = (fileList: FileList | File[]) => {
    const files = Array.from(fileList);
    if (files.length === 0) return;

    if (isMultiFlightMode && files.length > 1) {
      handleBatchFlightFiles(files);
    } else {
      const file = files[0];
      setSelectedFile(file);
      setMultiFlights([{
        id: `flight_single_${Date.now()}`,
        file,
        fileName: file.name,
        content: '',
        result: null as any,
        isVisited: true,
        yellowCount: 0,
        redCount: 0,
        durationExceededCount: 0
      }]);
      setActiveFlightIndex(0);
      handleRunAnalysis(file);
    }
  };

  // Filtreleme
  // Filtreleme ve Sıralama
  const filteredRows = useMemo(() => {
    if (!analysisResult) return [];

    // 3. Adım: Belirli bir limit aşım aralığına basılmışsa:
    // O aralığa denk gelen orijinal Excel satır aralığını getirir (startRow - endRow)
    if (selectedEpisodeId && selectedEpisodeId !== 'ALL_PARAM_ROWS') {
      const ep = analysisResult.episodes.find(e => e.id === selectedEpisodeId);
      if (ep) {
        const subRows = analysisResult.rows.filter(
          r => r.originalIndex >= ep.startRow && r.originalIndex <= ep.endRow
        );
        // Olayın en kritik (RED veya YELLOW) satırlarını en üste al, gerisi kronolojik
        return subRows.slice().sort((a, b) => {
          const aWeight = a.severity === 'RED' ? 2 : a.severity === 'YELLOW' ? 1 : 0;
          const bWeight = b.severity === 'RED' ? 2 : b.severity === 'YELLOW' ? 1 : 0;
          if (aWeight !== bWeight) return bWeight - aWeight;

          return a.originalIndex - b.originalIndex;
        });
      }
    }

    let list = analysisResult.rows;

    if (selectedParamKey) {
      list = list.filter(r => r.violations.some(v => {
        const up = v.colName.toUpperCase();
        if (selectedParamKey === 'ng') return up.includes('NG');
        if (selectedParamKey === 'itt') return up.includes('ITT');
        if (selectedParamKey === 'np') return up.includes('NP') || up.includes('RPM');
        if (selectedParamKey === 'torq') return up.includes('TORQ') || up.includes('TQ');
        if (selectedParamKey === 'fuel_p' || selectedParamKey === 'fuelP') return up.includes('FUEL P');
        if (selectedParamKey === 'oil_p' || selectedParamKey === 'oilP') return up.includes('OIL P');
        if (selectedParamKey === 'oil_t' || selectedParamKey === 'oilT') return up.includes('OIL T');
        return true;
      }));
    } else {
      if (filterMode === 'VIOLATIONS_ONLY') {
        list = list.filter(r => r.severity === 'RED' || r.severity === 'YELLOW');
      } else if (filterMode === 'RED_ONLY') {
        list = list.filter(r => r.severity === 'RED');
      } else if (filterMode === 'YELLOW_ONLY') {
        list = list.filter(r => r.severity === 'YELLOW');
      }
    }

    // Sıralama modu: Kronolojik (Excel orijinal satır akışı #1, #2, #3...) veya Öncelikli Kritik Aşımlar
    if (sortMode === 'SEVERITY') {
      const weight = (s: string) => (s === 'RED' ? 3 : s === 'YELLOW' ? 2 : 1);
      return list.slice().sort((a, b) => {
        const diff = weight(b.severity) - weight(a.severity);
        if (diff !== 0) return diff;
        return a.originalIndex - b.originalIndex;
      });
    } else {
      // Varsayılan: Excel orijinal kronolojik akışı
      return list.slice().sort((a, b) => a.originalIndex - b.originalIndex);
    }
  }, [analysisResult, filterMode, selectedParamKey, selectedEpisodeId, sortMode]);

  // Sayfalama
  const totalPages = Math.max(1, Math.ceil(filteredRows.length / pageSize));
  const validCurrentPage = Math.min(Math.max(currentPage, 1), totalPages);
  const paginatedRows = useMemo(() => {
    const start = (validCurrentPage - 1) * pageSize;
    return filteredRows.slice(start, start + pageSize);
  }, [filteredRows, validCurrentPage, pageSize]);

  // 3. Kısım: Tabloda gösterilecek başlıklar ve orijinal indeksleri
  const { activeHeaders, activeHeaderIndices } = useMemo(() => {
    if (!analysisResult) return { activeHeaders: [] as string[], activeHeaderIndices: [] as number[] };

    // Mod 1: Sadece seçilen parametre sütunu (Kullanıcı Talebi: SIRA, İKAZ, TIME ve SEÇİLEN PARAMETRE)
    if (tableColumnFilterMode === 'SELECTED_ONLY' && selectedColConfig) {
      const hdrs: string[] = [];
      const indices: number[] = [];

      // 1. TIME Sütunu (varsa)
      const timeIdx = analysisResult.headers.findIndex((h, idx) => h.toUpperCase().includes('TIME') || idx === 0);
      if (timeIdx >= 0) {
        hdrs.push(analysisResult.headers[timeIdx]);
        indices.push(timeIdx);
      }

      // 2. Seçilen Parametrenin Sütunu (Hemen ardından)
      const paramIdx = analysisResult.headers.findIndex((h, idx) =>
        selectedColConfig.matchPattern.test(h) || idx === selectedColConfig.expectedIndex
      );
      if (paramIdx >= 0 && paramIdx !== timeIdx) {
        hdrs.push(analysisResult.headers[paramIdx]);
        indices.push(paramIdx);
      } else if (paramIdx < 0 && selectedColConfig.expectedIndex < analysisResult.headers.length) {
        hdrs.push(analysisResult.headers[selectedColConfig.expectedIndex]);
        indices.push(selectedColConfig.expectedIndex);
      }

      return { activeHeaders: hdrs, activeHeaderIndices: indices };
    }

    // Mod 2: 7 Odak Sarı Parametre Sütunları
    if (tableColumnFilterMode === 'ALL_YELLOW') {
      const hdrs: string[] = [];
      const indices: number[] = [];
      analysisResult.headers.forEach((h, idx) => {
        if (h.toUpperCase().includes('TIME') || isYellowFocusHeader(h, idx)) {
          hdrs.push(h);
          indices.push(idx);
        }
      });
      return { activeHeaders: hdrs, activeHeaderIndices: indices };
    }

    // Mod 3: Orijinal Excel'deki Tüm Sütunlar
    return {
      activeHeaders: analysisResult.headers,
      activeHeaderIndices: analysisResult.headers.map((_, idx) => idx)
    };
  }, [analysisResult, tableColumnFilterMode, selectedColConfig]);

  // Olay Takip Çizelgesine Kaydet
  const handleCommitToOlayTakip = () => {
    if (!integrateForm.pilot || !integrateForm.pilot.trim()) {
      showNotification('Sorumlu Pilot alanı boş bırakılamaz!', 'error');
      return;
    }
    onAddToOlayTakip(integrateForm);
    setIsIntegrateModalOpen(false);
    showNotification('Limit aşımı kaydı AT-802 "Ng Limit Aşımları Detayı" sayfasına başarıyla eklendi!', 'success');
  };

  return (
    <div className="fixed inset-0 z-[10000] flex items-center justify-center p-2 sm:p-4 bg-slate-950/80 backdrop-blur-md animate-fade-in">
      <div ref={analysisModalContentRef} className="bg-white text-slate-800 w-full max-w-7xl h-[94vh] rounded-3xl shadow-2xl border border-slate-200 overflow-hidden flex flex-col">
        
        {/* ÜST BAŞLIK ÇUBUĞU */}
        <div className="px-6 py-4 bg-gradient-to-r from-[#0b3d1d] via-[#14532d] to-[#0b3d1d] text-white flex items-center justify-between shrink-0 shadow-md">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-2xl bg-orange-500/20 border border-orange-400/40 flex items-center justify-center text-orange-400 shadow-inner">
              <Gauge className="w-6 h-6 text-orange-400" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <span className="text-xs font-black tracking-widest text-orange-300 uppercase px-2 py-0.5 rounded-md bg-orange-950/60 border border-orange-500/30">
                  AT-802 MOTOR & UÇUŞ VERİLERİ
                </span>
                <span className="text-xs text-emerald-200/80 font-mono">MVP-50T ANALİZ MOTORU</span>
              </div>
              <h2 className="text-base sm:text-lg font-black tracking-tight text-white mt-0.5 flex items-center gap-2">
                <span>Ng Limit Aşımları ve Uçuş Kaydı Analiz Motoru</span>
              </h2>
            </div>
          </div>

          <div className="flex items-center gap-2">
            {/* 8 Sarı Sütun Limit Bilgi Çekmecesi */}
            <button
              type="button"
              onClick={() => setIsRefTableExpanded(!isRefTableExpanded)}
              className="px-3 py-1.5 rounded-xl bg-white/10 hover:bg-white/20 text-white text-xs font-bold transition-all flex items-center gap-1.5 cursor-pointer border border-white/10"
              title="7 Sarı Sütun ve Eşik Değerleri Tablosu"
            >
              <Info className="w-4 h-4 text-emerald-300" />
              <span className="hidden sm:inline">7 Sarı Odak Sütunu & Limitler</span>
            </button>

            {/* Kapat */}
            <button
              type="button"
              onClick={onClose}
              className="w-9 h-9 rounded-xl bg-white/10 hover:bg-red-600/80 hover:text-white text-white/80 transition-all flex items-center justify-center cursor-pointer"
            >
              <X className="w-5 h-5" />
            </button>
          </div>
        </div>

        {/* GÖRSELLERDEKİ 7 SARI SÜTUN REFERANS LİMİTLERİ ÇEKMECESİ */}
        {isRefTableExpanded && (
          <div className="bg-slate-900 text-slate-200 px-6 py-4 border-b border-slate-700 text-xs shrink-0 animate-slide-down">
            <div className="flex items-center justify-between mb-3">
              <div className="flex items-center gap-2">
                <span className="w-3 h-3 rounded-full bg-amber-400 animate-pulse" />
                <h4 className="font-black text-amber-400 uppercase tracking-wider text-xs">
                  Limit Değer Tablosu
                </h4>
              </div>
              <button
                type="button"
                onClick={() => setIsRefTableExpanded(false)}
                className="text-slate-400 hover:text-white text-xs cursor-pointer"
              >
                ✕ Kapat
              </button>
            </div>
            <div className="grid grid-cols-2 md:grid-cols-4 lg:grid-cols-7 gap-2.5 font-mono text-[11px]">
              {YELLOW_FOCUS_COLUMNS.map((col, idx) => (
                <div key={idx} className="bg-slate-800/90 p-2.5 rounded-xl border border-amber-500/40 shadow-inner">
                  <div className="font-black text-amber-300 mb-1 flex items-center justify-between">
                    <span>[{col.excelLetter}] {col.name}</span>
                  </div>
                  <div className="text-[10px] text-slate-300 leading-tight">
                    {col.rulesText}
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}

        {/* KONTROL VE HIZLI DOSYA YÜKLEME ÇUBUĞU (Yalnızca Analiz Yapıldığında Üstte Gösterilir) */}
        {analysisResult && (
          <div className="p-3 bg-slate-50 border-b border-slate-200 shrink-0">
            <div className="flex flex-wrap items-center justify-between gap-3">
              <div className="flex flex-wrap items-center gap-2.5">
                <input
                  ref={fileInputRef}
                  type="file"
                  multiple={isMultiFlightMode}
                  accept=".csv,.mvp,.txt"
                  onChange={handleFileChange}
                  className="hidden"
                />

                {/* Kuyruk Numarası Seçici (Depo AT-802 Listesi) */}
                <div className="flex items-center gap-1.5 bg-white border border-slate-300 rounded-xl px-2.5 py-1.5 shadow-2xs">
                  <Plane className="w-4 h-4 text-emerald-700 shrink-0" />
                  <span className="text-[11px] font-bold text-slate-500">Uçak (Depo):</span>
                  <select
                    value={selectedTail.fullLabel}
                    onChange={e => handleTailChange(e.target.value)}
                    className="text-xs font-black text-slate-900 bg-transparent focus:outline-none cursor-pointer font-mono"
                  >
                    {AT802_FLEET.map(ac => (
                      <option key={ac.fullLabel} value={ac.fullLabel}>
                        {ac.fullLabel} | {ac.motorSn}
                      </option>
                    ))}
                  </select>
                </div>

                {/* Çoklu Uçuş Seçici (Farklı Uçuş Seç Dropdown) */}
                {multiFlights.length > 1 && (
                  <div className="relative">
                    <div className="flex items-center gap-1">
                      <button
                        type="button"
                        onClick={() => setIsFlightDropdownOpen(!isFlightDropdownOpen)}
                        className={`px-3 py-1.5 rounded-xl font-black text-xs border-2 shadow-xs flex items-center gap-2 cursor-pointer transition-all active:scale-95 ${
                          multiFlights[activeFlightIndex]?.isVisited
                            ? 'bg-emerald-50 border-emerald-500 text-emerald-950'
                            : 'bg-white border-slate-300 text-slate-800 hover:bg-slate-50'
                        }`}
                        title="Taranan çoklu uçuşlar arasından farklı uçuş seç"
                      >
                        <div className="w-2.5 h-2.5 rounded-full bg-emerald-600 animate-pulse" />
                        <span className="font-extrabold uppercase tracking-wide">Farklı Uçuş Seç:</span>
                        <span className="max-w-[130px] truncate font-mono text-[11px]">
                          {multiFlights[activeFlightIndex]?.fileName || analysisResult.fileName}
                        </span>
                        <span className="text-[10px] font-mono px-1.5 py-0.5 bg-emerald-200 text-emerald-900 rounded font-black flex items-center gap-1">
                          <Check className="w-3 h-3 stroke-[3]" />
                          <span>{activeFlightIndex + 1}/{multiFlights.length}</span>
                        </span>
                        <ChevronDown className={`w-3.5 h-3.5 text-slate-600 transition-transform ${isFlightDropdownOpen ? 'rotate-180' : ''}`} />
                      </button>

                      {/* Hızlı Uçuş Değiştirme Butonları (Önceki / Sonraki) */}
                      <button
                        type="button"
                        disabled={activeFlightIndex === 0}
                        onClick={() => handleSelectMultiFlight(activeFlightIndex - 1)}
                        className="p-1.5 rounded-xl bg-white hover:bg-slate-100 disabled:opacity-40 disabled:cursor-not-allowed border border-slate-300 text-slate-700 cursor-pointer shadow-2xs"
                        title="Önceki Uçuş"
                      >
                        <ChevronLeft className="w-4 h-4" />
                      </button>
                      <button
                        type="button"
                        disabled={activeFlightIndex === multiFlights.length - 1}
                        onClick={() => handleSelectMultiFlight(activeFlightIndex + 1)}
                        className="p-1.5 rounded-xl bg-white hover:bg-slate-100 disabled:opacity-40 disabled:cursor-not-allowed border border-slate-300 text-slate-700 cursor-pointer shadow-2xs"
                        title="Sonraki Uçuş"
                      >
                        <ChevronRight className="w-4 h-4" />
                      </button>
                    </div>

                    {/* Alta Doğru Açılır Uçuş Listesi (Dropdown Menü) */}
                    {isFlightDropdownOpen && (
                      <div className="absolute left-0 top-full mt-2 w-96 max-h-[380px] overflow-y-auto bg-white rounded-2xl shadow-2xl border-2 border-emerald-500/40 z-50 p-2.5 animate-in fade-in slide-in-from-top-2">
                        <div className="px-3 py-2 border-b border-slate-100 flex items-center justify-between text-[11px] font-black text-slate-600 uppercase tracking-wider">
                          <span className="flex items-center gap-1.5 text-emerald-900">
                            <Plane className="w-3.5 h-3.5 text-emerald-700" />
                            <span>Taranan Uçuşlar ({multiFlights.length})</span>
                          </span>
                          <span className="text-emerald-700 font-mono text-[10px] bg-emerald-100 px-2 py-0.5 rounded-full">
                            {multiFlights.filter(f => f.isVisited).length} / {multiFlights.length} İncelendi
                          </span>
                        </div>
                        <div className="py-1.5 space-y-1.5">
                          {multiFlights.map((item, idx) => {
                            const isCurrent = activeFlightIndex === idx;
                            const isVisited = item.isVisited;
                            return (
                              <button
                                key={item.id}
                                type="button"
                                onClick={() => handleSelectMultiFlight(idx)}
                                className={`w-full text-left p-3 rounded-xl text-xs transition-all flex items-center justify-between gap-2.5 cursor-pointer border ${
                                  isCurrent
                                    ? 'bg-emerald-100/90 border-emerald-500 ring-2 ring-emerald-500 text-emerald-950 font-black shadow-xs'
                                    : isVisited
                                      ? 'bg-emerald-50/90 hover:bg-emerald-100/70 border-emerald-300 text-emerald-900 font-bold'
                                      : 'bg-white hover:bg-slate-50 border-slate-200 text-slate-700'
                                }`}
                              >
                                <div className="flex items-center gap-2.5 min-w-0">
                                  <span className={`w-6 h-6 rounded-lg flex items-center justify-center text-[11px] font-mono font-black shrink-0 ${
                                    isVisited ? 'bg-emerald-200 text-emerald-900 border border-emerald-300' : 'bg-slate-100 text-slate-600 border border-slate-200'
                                  }`}>
                                    #{idx + 1}
                                  </span>
                                  <div className="min-w-0">
                                    <p className="truncate font-mono font-bold text-xs">{item.fileName}</p>
                                    <div className="flex items-center gap-2 text-[10px] font-mono mt-0.5">
                                      {item.yellowCount > 0 ? (
                                        <span className="text-amber-800 font-bold bg-amber-100 px-1.5 py-0.2 rounded">
                                          {item.yellowCount} Sarı İkaz
                                        </span>
                                      ) : (
                                        <span className="text-emerald-700">0 İkaz</span>
                                      )}
                                      {item.redCount > 0 && (
                                        <span className="text-rose-700 font-bold bg-rose-100 px-1.5 py-0.2 rounded">
                                          {item.redCount} Kırmızı
                                        </span>
                                      )}
                                      {item.hasTransferredYellow && (
                                        <span className="text-emerald-800 font-black bg-emerald-200/90 px-1.5 py-0.2 rounded">
                                          Aktarıldı
                                        </span>
                                      )}
                                    </div>
                                  </div>
                                </div>

                                {/* Yeşil Dolgu ve Tikli Rozet */}
                                {isVisited ? (
                                  <span className="shrink-0 flex items-center gap-1 px-2.5 py-1 rounded-full text-[10px] font-black bg-emerald-600 text-white shadow-2xs">
                                    <Check className="w-3.5 h-3.5 stroke-[3]" />
                                    <span>İncelendi</span>
                                  </span>
                                ) : (
                                  <span className="shrink-0 text-[10px] text-slate-500 font-semibold px-2 py-0.5 bg-slate-100 rounded-full border border-slate-200">
                                    Bekliyor
                                  </span>
                                )}
                              </button>
                            );
                          })}
                        </div>
                      </div>
                    )}
                  </div>
                )}

                {/* Dosya Değiştir Butonu */}
                <button
                  type="button"
                  onClick={() => fileInputRef.current?.click()}
                  className="px-3.5 py-1.5 rounded-xl bg-white hover:bg-slate-100 text-slate-800 font-bold text-xs border border-slate-300 shadow-2xs flex items-center gap-1.5 cursor-pointer transition-all active:scale-95"
                  title="Farklı bir MVP/CSV uçuş dosyası seç"
                >
                  <UploadCloud className="w-4 h-4 text-emerald-700" />
                  <span className="max-w-[200px] truncate">{analysisResult.fileName}</span>
                  <span className="text-[10px] bg-slate-100 text-slate-600 px-1.5 py-0.2 rounded font-mono font-bold">
                    Değiştir
                  </span>
                </button>

                {/* Tespit Edilen Ayırıcı ve Sütun Sayısı Göstergesi */}
                <div className="flex items-center gap-1.5 px-2.5 py-1.5 rounded-xl bg-white border border-slate-300 text-[11px] font-mono text-slate-700 shadow-2xs">
                  <span className="w-2 h-2 rounded-full bg-emerald-500" />
                  <span>Ayırıcı: <strong className="text-emerald-800">{analysisResult.delimiter === ',' ? 'Virgül (,)' : analysisResult.delimiter === ';' ? 'Noktalı Virgül (;)' : 'Tab'}</strong></span>
                  <span className="text-slate-300">|</span>
                  <span><strong className="text-slate-900">{analysisResult.headers.length} Sütun</strong> (Excel A - {getExcelColumnLetter(analysisResult.headers.length - 1)})</span>
                </div>

                {/* 7 Sarı Sütun / Tüm Sütunlar Görünüm Toggle */}
                <button
                  type="button"
                  onClick={() => setShowOnlyYellowColumns(!showOnlyYellowColumns)}
                  className={`px-3 py-1.5 rounded-xl font-bold text-xs border transition-all flex items-center gap-1.5 cursor-pointer ${
                    showOnlyYellowColumns
                      ? 'bg-amber-300 text-slate-900 border-amber-400 font-black shadow-2xs'
                      : 'bg-white text-slate-700 border-slate-300 hover:bg-slate-50'
                  }`}
                  title="Görseldeki 7 sarı dolgulu sütuna odaklan veya tüm sütunları gör"
                >
                  <SlidersHorizontal className="w-3.5 h-3.5 text-amber-800" />
                  <span>
                    {showOnlyYellowColumns
                      ? 'Sadece 7 Sarı Sütun'
                      : 'Tüm Sütunlar (A - AR)'}
                  </span>
                </button>

                {/* Sıralama Modu Toggle: Kronolojik (Excel Orijinal Akışı) vs Öncelikli Aşımlar */}
                <button
                  type="button"
                  onClick={() => setSortMode(sortMode === 'CHRONO' ? 'SEVERITY' : 'CHRONO')}
                  className={`px-3 py-1.5 rounded-xl font-bold text-xs border transition-all flex items-center gap-1.5 cursor-pointer ${
                    sortMode === 'CHRONO'
                      ? 'bg-emerald-800 text-white border-emerald-900 font-black shadow-2xs'
                      : 'bg-white text-slate-700 border-slate-300 hover:bg-slate-50'
                  }`}
                  title="Excel orijinal kronolojik sıra (#1, #2, #3...) veya aşımları önceleme"
                >
                  <Clock className="w-3.5 h-3.5 text-emerald-300" />
                  <span>
                    {sortMode === 'CHRONO'
                      ? 'Kronolojik Akış (#1, #2...)'
                      : 'Önce Aşımlar'}
                  </span>
                </button>

                {/* Aktif Olay Filtresi Temizleme Butonu */}
                {selectedEpisodeId && (
                  <button
                    type="button"
                    onClick={() => setSelectedEpisodeId(null)}
                    className="px-2.5 py-1.5 rounded-xl bg-red-100 hover:bg-red-200 text-red-900 font-black text-xs border border-red-300 flex items-center gap-1 cursor-pointer transition-colors shadow-2xs animate-fade-in"
                    title="Seçili olay filtresini kaldırıp tüm satırları göster"
                  >
                    <span>Filtreyi Temizle</span>
                    <X className="w-3.5 h-3.5 text-red-700" />
                  </button>
                )}
              </div>

              {/* Olay Takip Çizelgesine Ekle Butonu */}
              {(analysisResult.redCount > 0 || analysisResult.yellowCount > 0) && (
                <div className="flex items-center gap-2">
                  <button
                    type="button"
                    onClick={() => setIsIntegrateModalOpen(true)}
                    className="px-4 py-1.5 rounded-xl bg-gradient-to-r from-red-600 via-orange-600 to-amber-600 hover:from-red-700 hover:to-amber-700 text-white font-black text-xs uppercase tracking-wider flex items-center gap-2 shadow-md transition-all hover:scale-[1.02] active:scale-95 cursor-pointer animate-pulse"
                  >
                    <Plane className="w-4 h-4 text-white" />
                    <span>Olay Takip Çizelgesine Ekle</span>
                    <span className="bg-white/20 text-white px-1.5 py-0.5 rounded text-[10px] font-mono font-bold">
                      {analysisResult.redCount + analysisResult.yellowCount} Aşım
                    </span>
                  </button>
                </div>
              )}
            </div>
          </div>
        )}

        {/* ASENKRON TARAMA YÜZDE ÇUBUĞU */}
        {isAnalyzing && (
          <div className="p-3 bg-white border-b border-orange-300 shadow-sm animate-fade-in shrink-0">
            <div className="flex items-center justify-between text-xs font-mono font-bold text-slate-800 mb-2">
              <span className="flex items-center gap-2">
                <Activity className="w-4 h-4 text-orange-600 animate-spin" />
                <span className="text-orange-950 font-black">
                  {selectedTail.fullLabel} Uçuş Veri Kaydı Taranıyor:
                </span>
                <span className="text-slate-600">{parsingStatusText}</span>
              </span>
              <span className="text-orange-600 font-black text-sm bg-orange-50 px-2 py-0.5 rounded-md border border-orange-200">
                %{parsingProgress}
              </span>
            </div>
            <div className="w-full bg-slate-100 rounded-full h-2.5 overflow-hidden border border-slate-200">
              <div
                className="bg-gradient-to-r from-orange-500 via-amber-500 to-emerald-500 h-2.5 rounded-full transition-all duration-150 ease-out"
                style={{ width: `${parsingProgress}%` }}
              />
            </div>
          </div>
        )}



        {/* METADATA AKORDİYONU */}
        {analysisResult && (
          <div className="bg-emerald-50/50 border-b border-emerald-200/80 px-6 py-2 text-xs shrink-0">
            <div className="flex items-center justify-between">
              <button
                type="button"
                onClick={() => setIsMetadataExpanded(!isMetadataExpanded)}
                className="flex items-center gap-2 font-black text-emerald-900 cursor-pointer hover:text-emerald-950 select-none"
              >
                <FileText className="w-4 h-4 text-emerald-700" />
                <span>Uçak & Kayıt Tanıtım Bilgileri (Metadata Satırları)</span>
                {isMetadataExpanded ? (
                  <ChevronUp className="w-4 h-4 text-emerald-700" />
                ) : (
                  <ChevronDown className="w-4 h-4 text-emerald-700" />
                )}
              </button>

              <div className="text-[11px] font-mono text-emerald-900 flex items-center gap-2 flex-wrap">
                <span>
                  Uçak Depo: <strong>{selectedTail.fullLabel}</strong>
                </span>
                <span>|</span>
                <span>
                  Motor: <strong>{analysisResult.motorSn || selectedTail.motorSn}</strong>
                </span>
              </div>
            </div>

            {isMetadataExpanded && (
              <div className="mt-2.5 grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-2 bg-white p-3 rounded-xl border border-emerald-200 text-[11px] font-mono shadow-2xs animate-slide-down max-h-40 overflow-y-auto">
                {analysisResult.metadataLines.map((line, idx) => (
                  <div key={idx} className="bg-slate-50 p-1.5 rounded border border-slate-200/60 truncate">
                    <span className="text-slate-400 font-bold mr-1.5">{idx + 1}.</span>
                    <span className="text-slate-800">{line}</span>
                  </div>
                ))}
              </div>
            )}
          </div>
        )}

        {/* ANA GÖVDE: DOSYA YÜKLEME KARTI YA DA SONUÇ VERİ TABLOSU */}
        <div className="flex-1 overflow-auto bg-slate-50 relative">
          {!analysisResult ? (
            /* KUYRUK SEÇ VE DOSYA YÜKLE TEK ALANI (Kullanıcı İsteğine Göre Sadeleştirilmiş) */
            <div className="h-full flex items-center justify-center p-6">
              <input
                ref={fileInputRef}
                type="file"
                multiple={isMultiFlightMode}
                accept=".csv,.mvp,.txt"
                onChange={handleFileChange}
                className="hidden"
              />

              <div className="w-full max-w-xl bg-white rounded-3xl border border-slate-200 shadow-xl overflow-hidden animate-fade-in">
                {/* Kart Başlığı */}
                <div className="px-6 py-5 bg-gradient-to-r from-[#0b3d1d] to-[#14532d] text-white flex items-center gap-3">
                  <div className="w-10 h-10 rounded-2xl bg-amber-400/20 border border-amber-400/30 flex items-center justify-center text-amber-300">
                    <Plane className="w-5 h-5 text-amber-300" />
                  </div>
                  <div>
                    <h3 className="text-base font-black text-white">AT-802 Uçuş Kaydı Analizi (MVP-50T)</h3>
                    <p className="text-xs text-emerald-200/90 font-medium">Hava aracı kuyruğunu seçip MVP/CSV uçuş dosyasını yükleyiniz</p>
                  </div>
                </div>

                <div className="p-6 space-y-5">
                  {/* 1. ADIM: KUYRUK NUMARASI SEÇİMİ (DEPO ENVANTERİ BİREBİR LİSTESİ) */}
                  <div>
                    <label className="block text-xs font-black text-slate-700 uppercase tracking-wider mb-2 flex items-center justify-between">
                      <span className="flex items-center gap-1.5">
                        <span className="w-5 h-5 rounded-full bg-emerald-100 text-emerald-800 text-[11px] font-black flex items-center justify-center">1</span>
                        <span>Hava Aracı Kuyruk Numarası (Depo Envanteri)</span>
                      </span>
                      <span className="text-[10px] text-emerald-700 font-bold font-mono">16 AT-802 Uçağı</span>
                    </label>
                    <div className="relative">
                      <select
                        value={selectedTail.fullLabel}
                        onChange={e => handleTailChange(e.target.value)}
                        className="w-full px-3.5 py-3 rounded-2xl bg-slate-50 border-2 border-slate-200 text-xs font-black text-slate-800 hover:border-emerald-500 focus:border-emerald-600 focus:bg-white focus:outline-none transition-all cursor-pointer font-mono"
                      >
                        {AT802_FLEET.map(ac => (
                          <option key={ac.fullLabel} value={ac.fullLabel}>
                            {ac.fullLabel} • Motor: {ac.motorSn}
                          </option>
                        ))}
                      </select>
                    </div>
                    <div className="mt-2 flex items-center justify-between text-[11px] text-slate-500 bg-slate-50 px-3 py-1.5 rounded-xl border border-slate-200 font-mono">
                      <span>Çağrı: <strong className="text-slate-800">{selectedTail.callSign}</strong></span>
                      <span>Kod: <strong className="text-slate-800">{selectedTail.code}</strong></span>
                      <span>Motor S/N: <strong className="text-emerald-800">{selectedTail.motorSn}</strong></span>
                    </div>
                  </div>

                  {/* 2. ADIM: DOSYA YÜKLEME ALANI (SÜRÜKLE & BIRAK YA DA TIKLA) */}
                  <div>
                    <div className="flex items-center justify-between mb-2">
                      <label className="text-xs font-black text-slate-700 uppercase tracking-wider flex items-center gap-1.5">
                        <span className="w-5 h-5 rounded-full bg-emerald-100 text-emerald-800 text-[11px] font-black flex items-center justify-center">2</span>
                        <span>Uçuş Kayıt Dosyası (.csv / .mvp / .txt)</span>
                      </label>
                      <label className="flex items-center gap-1.5 px-3 py-1 rounded-xl bg-emerald-50 border border-emerald-300 text-emerald-900 font-bold text-xs cursor-pointer select-none hover:bg-emerald-100 transition shadow-2xs">
                        <input
                          type="checkbox"
                          checked={isMultiFlightMode}
                          onChange={(e) => setIsMultiFlightMode(e.target.checked)}
                          className="w-4 h-4 text-emerald-700 rounded focus:ring-emerald-500 cursor-pointer accent-emerald-600"
                        />
                        <span>Birden fazla uçuş analiz et</span>
                        <span className="text-[10px] bg-emerald-200/90 text-emerald-950 font-black px-1.5 py-0.2 rounded font-mono">ÇOKLU</span>
                      </label>
                    </div>

                    <div
                      onDragOver={handleDragOver}
                      onDragLeave={handleDragLeave}
                      onDrop={handleDrop}
                      onClick={() => fileInputRef.current?.click()}
                      className={`border-2 border-dashed rounded-3xl p-8 text-center cursor-pointer transition-all flex flex-col items-center justify-center group ${
                        isDragging
                          ? 'border-emerald-500 bg-emerald-50/80 scale-[1.01]'
                          : 'border-slate-300 hover:border-emerald-500 hover:bg-slate-50/60 bg-white'
                      }`}
                    >
                      <div className="w-14 h-14 rounded-2xl bg-emerald-50 border border-emerald-200 flex items-center justify-center text-emerald-700 mb-3 group-hover:scale-105 transition-transform">
                        <UploadCloud className="w-7 h-7 text-emerald-700" />
                      </div>
                      <h4 className="text-sm font-black text-slate-800 mb-1">
                        {isMultiFlightMode
                          ? 'Birden Fazla MVP / CSV Dosyası Seçin veya Sürükleyin'
                          : 'MVP / CSV Uçuş Dosyası Seçin veya Buraya Sürükleyin'}
                      </h4>
                      <p className="text-xs text-slate-500 max-w-sm mb-3">
                        {isMultiFlightMode
                          ? 'Çoklu dosya seçimi aktif: Seçilen tüm uçuşlar taranacak, ardından açılır listeden yeşil tikli olarak sırayla incelenebilecektir.'
                          : 'Dosya seçildiği anda otomatik olarak 7 sarı odak sütunu taranacak ve limit aşımları listelenecektir.'}
                      </p>
                      <span className="px-5 py-2.5 rounded-xl bg-emerald-700 text-white font-black text-xs uppercase tracking-wider shadow-sm group-hover:bg-emerald-800 transition-colors flex items-center gap-2">
                        <UploadCloud className="w-4 h-4" />
                        <span>{isMultiFlightMode ? 'Çoklu Uçuş Dosyaları Seç & Yükle' : 'Dosya Seç & Yükle'}</span>
                      </span>
                    </div>
                  </div>

                  {/* Odak Limit Bilgilendirme Kutusu */}
                  <div className="p-3 bg-amber-50 rounded-2xl border border-amber-200 text-[11px] text-amber-900 leading-relaxed flex items-start gap-2 font-mono">
                    <Info className="w-4 h-4 text-amber-700 shrink-0 mt-0.5" />
                    <div>
                      <strong>Odak Analiz Motoru:</strong> MVP-50T kayıtlarındaki 44 sütundan yalnızca belirlenen <strong>7 Sarı Odak Sütunu</strong> (C: NG %, P: ITT *C, D: NP RPM, J: TORQ FT-LB, K: FUEL P PSI, L: OIL P PSI, Q: OIL T *C) incelenir. 0 ve 4 basamaklı hatalı NG değerleri yok sayılır.
                    </div>
                  </div>
                </div>
              </div>
            </div>
          ) : (
            <div className="flex flex-col min-h-full">
              {/* ========================================================================= */}
              {/* 1. KISIM: KAYITLARDA OLUŞAN EN YÜKSEK DEĞERLER (7 SARI SÜTUN) */}
              {/* ========================================================================= */}
              <div className="p-4 bg-white border-b border-slate-200 shrink-0 shadow-xs">
                <div className="flex flex-wrap items-center justify-between gap-2 mb-3">
                  <div className="flex items-center gap-2.5">
                    <span className="w-6 h-6 rounded-lg bg-emerald-800 text-white text-xs font-black flex items-center justify-center shadow-xs">
                      1
                    </span>
                    <div>
                      <h3 className="text-xs sm:text-sm font-black text-slate-900 uppercase tracking-tight flex items-center gap-2 flex-wrap">
                        <span>1. KISIM: Kayıtlarda Oluşan En Yüksek Değerler</span>
                        <span className="text-[10px] text-emerald-800 bg-emerald-50 px-2 py-0.5 rounded-full font-mono font-bold border border-emerald-200">
                          7 Sarı Sütun Takibi
                        </span>
                        {(() => {
                          const flightStartRow = analysisResult.rows.find(r => r.originalIndex >= 17) || analysisResult.rows[0];
                          const flightEndRow = analysisResult.rows[analysisResult.rows.length - 1];
                          const startTime = flightStartRow ? flightStartRow.timeStr : '-';
                          const endTime = flightEndRow ? flightEndRow.timeStr : '-';
                          let durationStr = '-';
                          if (flightStartRow?.timeSec != null && flightEndRow?.timeSec != null) {
                            const diffSec = Math.max(0, flightEndRow.timeSec - flightStartRow.timeSec);
                            const hrs = Math.floor(diffSec / 3600);
                            const mins = Math.floor((diffSec % 3600) / 60);
                            const secs = Math.round(diffSec % 60);
                            if (hrs > 0) {
                              durationStr = `${hrs} sa ${mins} dk ${secs} sn`;
                            } else if (mins > 0) {
                              durationStr = `${mins} dk ${secs} sn`;
                            } else {
                              durationStr = `${secs} sn`;
                            }
                          }
                          return (
                            <span className="flex items-center gap-2 ml-1">
                              <span className="text-[10px] text-slate-700 bg-slate-100 px-2 py-0.5 rounded-full font-mono font-bold border border-slate-200 flex items-center gap-1">
                                <Clock className="w-3 h-3 text-slate-500" />
                                <span>Başlangıç - Bitiş: {startTime} — {endTime}</span>
                              </span>
                              <span className="text-[10px] text-emerald-800 bg-emerald-50 px-2 py-0.5 rounded-full font-mono font-bold border border-emerald-200 flex items-center gap-1">
                                <Activity className="w-3 h-3 text-emerald-600" />
                                <span>Süre: {durationStr}</span>
                              </span>
                            </span>
                          );
                        })()}
                      </h3>
                      <p className="text-[11px] text-slate-500">
                        Limit aşım süre ve zaman aralıklarını görmek için incelemek istediğiniz parametrenin üzerine tıklayınız.
                      </p>
                    </div>
                  </div>
                  <div className="flex items-center gap-2 flex-wrap">
                    <div className="flex items-center gap-2 bg-amber-50 px-3 py-1.5 rounded-xl border border-amber-300 shadow-2xs">
                      <span className="font-black text-amber-950 text-xs">Sarı İkaz Adedi:</span>
                      <button
                        type="button"
                        onClick={() => setIsYellowNgHistoryModalOpen(true)}
                        className="px-2.5 py-0.5 bg-amber-500 hover:bg-amber-600 text-white font-black rounded-lg text-xs cursor-pointer shadow-xs transition-transform active:scale-95 flex items-center gap-1"
                        title="Bu hava aracına ait kayıtlı sarı ikaz satırlarını incele (Google E-Tablo)"
                      >
                        <span>{yellowNgCounts[selectedTail.code] || 0} Adet</span>
                        <Eye className="w-3.5 h-3.5 opacity-80" />
                      </button>
                      <button
                        type="button"
                        disabled={isSyncingYellowNg}
                        onClick={handleInitiateCountYellowNg}
                        className={`px-3.5 py-1.5 ${
                          isSyncingYellowNg 
                            ? 'bg-amber-400 text-white cursor-not-allowed opacity-90' 
                            : 'bg-gradient-to-r from-amber-600 via-amber-700 to-amber-800 hover:from-amber-700 hover:via-amber-800 hover:to-amber-900 active:scale-95 text-white shadow-xs cursor-pointer'
                        } font-black text-xs uppercase tracking-wider rounded-lg transition-all flex items-center gap-1.5 border border-amber-500/40`}
                        title={
                          isSyncingYellowNg
                            ? 'Google E-Tabloya aktarılıyor, lütfen bekleyiniz...'
                            : analysisResult
                              ? `Yüklü uçuş kaydındaki ${currentNgYellowEpisodes.length} adet Ng Sarı İkazını Google E-Tabloya canlı aktar ve saydır`
                              : 'Önce MVP/CSV uçuş kayıt dosyası yükleyiniz'
                        }
                      >
                        {isSyncingYellowNg ? (
                          <>
                            <RefreshCw className="w-3.5 h-3.5 animate-spin text-amber-100" />
                            <span>E-Tabloya Aktarılıyor...</span>
                          </>
                        ) : (
                          <>
                            <Flame className="w-3.5 h-3.5 text-amber-200" />
                            <span>Sarı İkazları Saydır{currentNgYellowEpisodes.length > 0 ? ` (${currentNgYellowEpisodes.length})` : ''}</span>
                          </>
                        )}
                      </button>
                    </div>

                    {/* AT-802F Motor Uçuş Grafiği & Oynatıcı Butonu */}
                    <button
                      type="button"
                      onClick={() => {
                        setChartInitialParamKey(selectedParamKey || 'ng');
                        setIsFlightChartModalOpen(true);
                      }}
                      className="px-3.5 py-1.5 rounded-xl bg-gradient-to-r from-blue-700 via-indigo-700 to-slate-900 hover:from-blue-800 text-white font-black text-xs transition-all shadow-xs flex items-center gap-1.5 cursor-pointer active:scale-95 border border-blue-500/40"
                      title="AT-802F Uçak İmleçli Motor Uçuş Grafiğini ve Oynatıcıyı Aç (X: Zaman, Y: Limitler)"
                    >
                      <TrendingUp className="w-3.5 h-3.5 text-amber-300" />
                      <span>📈 Motor Uçuş Grafiği Aç</span>
                    </button>

                    {/* Analiz Ekran Görüntüsü İndir Butonu */}
                    <button
                      type="button"
                      onClick={handleDownloadAnalysisScreenshot}
                      disabled={isCapturingAnalysis}
                      className="px-3.5 py-1.5 rounded-xl bg-gradient-to-r from-emerald-600 via-teal-600 to-slate-900 hover:from-emerald-700 text-white font-black text-xs transition-all shadow-xs flex items-center gap-1.5 cursor-pointer active:scale-95 border border-emerald-500/40 disabled:opacity-50"
                      title="Mevcut Uçuş Analiz Görünümünün Ekran Görüntüsünü (PNG) İndir"
                    >
                      {isCapturingAnalysis ? (
                        <>
                          <Activity className="w-3.5 h-3.5 animate-spin text-emerald-200" />
                          <span>Görüntü Alınıyor...</span>
                        </>
                      ) : (
                        <>
                          <Camera className="w-3.5 h-3.5 text-emerald-300" />
                          <span>📸 Ekran Görüntüsü İndir</span>
                        </>
                      )}
                    </button>

                    {selectedParamKey && (
                      <button
                        type="button"
                        onClick={() => {
                          setSelectedParamKey(null);
                          setSelectedEpisodeId(null);
                        }}
                        className="px-2.5 py-1 rounded-lg bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold transition-all cursor-pointer flex items-center gap-1"
                      >
                        <RotateCcw className="w-3.5 h-3.5 text-slate-500" />
                        <span>Seçimi Temizle (Yalnızca 1. Kısım)</span>
                      </button>
                    )}
                  </div>
                </div>

                {/* 7 PARAMETRE EN YÜKSEK DEĞERLER KARTLARI (Sütun 1 2 3 yok, Odak aşımları yok) */}
                <div className="grid grid-cols-2 sm:grid-cols-4 lg:grid-cols-7 gap-2.5">
                  {YELLOW_FOCUS_COLUMNS.map((col) => {
                    const isSelected = selectedParamKey === col.key;
                    const info = getParamData(col.key);
                    return (
                      <div
                        key={col.key}
                        onClick={() => {
                          if (!analysisResult) return;
                          const eps = analysisResult.episodes.filter(ep => {
                            if (ep.colKey === col.key) return true;
                            const up = ep.colName.toUpperCase();
                            if (col.key === 'ng') return up.includes('NG');
                            if (col.key === 'itt') return up.includes('ITT');
                            if (col.key === 'np') return up.includes('NP') || up.includes('RPM');
                            if (col.key === 'torq') return up.includes('TORQ') || up.includes('TQ');
                            if (col.key === 'fuel_p') return up.includes('FUEL P');
                            if (col.key === 'oil_p') return up.includes('OIL P');
                            if (col.key === 'oil_t') return up.includes('OIL T');
                            return false;
                          });

                          // Bu parametre için 2. kısımda resmi aşım epizodu var mı ve bu epizotlar tepe değerini kapsıyor mu?
                          const peakValue = (info as any).peakValNumeric || 0; // info'dan nümerik değeri almalıyız
                          // Bu parametre için 2. kısımda resmi aşım epizodu var mı?
                          // Epizotların tepe değerleri ile mutlak tepe değerini karşılaştır.
                          // Eğer fark çok küçükse bu bir epizot parçasıdır, değilse anlıktır.
                          const isPeakInEpisodes = eps.some(ep => Math.abs(peakValue - ep.peakValue) < 2); // 2 RPM tolerans

                          // Eğer 2. kısımda epizot yoksa VEYA tepe değerini içeren bir epizot yoksa,
                          // bu değer süre koşulunu sağlamayan anlık bir tepe değeridir.
                          if ((eps.length === 0 || !isPeakInEpisodes) && !dontShowTransientAgain[col.key]) {
                            // Bu tepe değerine sahip veya en yakın satırı bul
                            let peakRow: ParsedMvpRow | null = null;
                            let targetPeakVal: number | null = null;
                            if (col.key === 'np') targetPeakVal = analysisResult.peakNp;
                            else if (col.key === 'ng') targetPeakVal = analysisResult.peakNg;
                            else if (col.key === 'itt') targetPeakVal = analysisResult.peakItt;
                            else if (col.key === 'torq') targetPeakVal = analysisResult.peakTq;
                            else if (col.key === 'oil_t') targetPeakVal = analysisResult.maxOilT;

                            if (targetPeakVal !== null) {
                              const matchRow = analysisResult.rows.find(r => {
                                const activeIndex = col.key === 'np' ? 3 : col.key === 'ng' ? 2 : col.key === 'itt' ? 15 : col.key === 'torq' ? 9 : 16;
                                const val = parseCleanNumber(r.cells[activeIndex]);
                                return !isNaN(val) && Math.abs(val - targetPeakVal!) < 0.001;
                              });
                              if (matchRow) peakRow = matchRow;
                            }

                            if (peakRow) {
                              setTransientWarningPopup({
                                isOpen: true,
                                colKey: col.key,
                                colName: col.name,
                                peakVal: targetPeakVal || 0,
                                peakValFormatted: info.peak,
                                row: peakRow
                              });

                            }
                          }

                          if (selectedParamKey === col.key) {
                            setSelectedParamKey(null);
                            setSelectedEpisodeId(null);
                          } else {
                            setSelectedParamKey(col.key);
                            setSelectedEpisodeId(null);
                          }
                        }}
                        className={`p-3 rounded-2xl border-2 transition-all cursor-pointer text-left relative group ${
                          isSelected
                            ? 'ring-3 ring-emerald-500 bg-amber-100/90 border-emerald-600 shadow-md scale-[1.02]'
                            : info.red > 0
                            ? 'border-red-300 bg-red-50/40 hover:border-red-500 hover:bg-red-50'
                            : info.yellow > 0
                            ? 'border-amber-300 bg-amber-50/40 hover:border-amber-500 hover:bg-amber-50'
                            : 'border-slate-200 bg-white hover:border-amber-400 hover:bg-amber-50/30'
                        }`}
                      >
                        <div className="flex items-center justify-between gap-1 mb-1">
                          <span className="text-[11px] font-black text-slate-900 uppercase truncate">
                            {col.name}
                          </span>
                          <span className="text-[9px] font-mono font-bold px-1 rounded bg-slate-100 text-slate-600">
                            {col.excelLetter}
                          </span>
                        </div>

                        <div className="text-base sm:text-lg font-black text-slate-900 my-1 font-mono tracking-tight">
                          {info.peak}
                        </div>

                        <div className="text-[10px] text-slate-500 truncate font-mono">
                          {info.limit}
                        </div>

                        <div className="mt-2 pt-1.5 border-t border-slate-200/70 flex items-center justify-between gap-1">
                          {info.total > 0 ? (
                            <div className="flex items-center gap-1 font-mono text-[10px]">
                              {info.red > 0 && (
                                <span className="bg-red-600 text-white font-black px-1 py-0.2 rounded">
                                  {info.red} Kırmızı
                                </span>
                              )}
                              {info.yellow > 0 && (
                                <span className="bg-amber-400 text-slate-950 font-black px-1 py-0.2 rounded">
                                  {info.yellow} Sarı
                                </span>
                              )}
                            </div>
                          ) : (
                            <span className="text-[10px] font-bold text-emerald-700 bg-emerald-50 px-1.5 py-0.2 rounded font-mono">
                              Aşım Yok
                            </span>
                          )}

                          <div className="flex items-center gap-1.5">
                            <button
                              type="button"
                              onClick={(e) => {
                                e.stopPropagation();
                                setChartInitialParamKey(col.key);
                                setIsFlightChartModalOpen(true);
                              }}
                              className="px-2 py-0.5 rounded-md bg-amber-500 hover:bg-amber-600 active:scale-95 text-slate-950 font-black text-[9px] shadow-2xs transition-all flex items-center gap-1 cursor-pointer border border-amber-600/30"
                              title={`${col.name} Uçuş Grafiğini Aç`}
                            >
                              <TrendingUp className="w-2.5 h-2.5 text-slate-950" />
                              <span>Grafik Aç</span>
                            </button>
                            <span className={`text-[10px] font-black ${isSelected ? 'text-emerald-700' : 'text-slate-400 group-hover:text-amber-800'}`}>
                              {isSelected ? '✓ Açık' : '➔'}
                            </span>
                          </div>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>

              {/* ========================================================================= */}
              {/* İLK AÇILIŞTA (Henüz parametre seçilmediyse): SADECE 1. KISIM VE REHBER */}
              {/* ========================================================================= */}
              {!selectedParamKey && (
                <div className="flex-1 flex items-center justify-center p-12 text-center animate-fade-in">
                  <div className="max-w-md mx-auto">
                    <div className="w-14 h-14 mx-auto rounded-3xl bg-amber-100 text-amber-800 flex items-center justify-center font-black text-xl mb-3 shadow-inner">
                      1
                    </div>
                    <h3 className="text-base font-black text-slate-800 mb-1">
                      1. Kısım: İncelenecek Parametreyi Seçin
                    </h3>
                    <p className="text-xs text-slate-500 leading-relaxed mb-4">
                      İlk açılışta yalnızca en yüksek değerler gösterilmektedir. Tespit edilen limit aşım aralık ve sürelerini (2. Kısım) ve orijinal Excel satırlarını (3. Kısım) adım adım görmek için yukarıdaki parametre kartlarından birine tıklayınız.
                    </p>
                    <div className="flex flex-wrap items-center justify-center gap-2">
                      {YELLOW_FOCUS_COLUMNS.map(c => (
                        <button
                          key={c.key}
                          type="button"
                          onClick={() => setSelectedParamKey(c.key)}
                          className="px-3 py-1.5 rounded-xl bg-white hover:bg-amber-50 text-slate-800 border border-slate-300 hover:border-amber-400 font-bold text-xs transition-all shadow-2xs cursor-pointer flex items-center gap-1.5"
                        >
                          <span className="font-mono text-amber-700 font-black">[{c.excelLetter}]</span>
                          <span>{c.name}</span>
                        </button>
                      ))}
                    </div>
                  </div>
                </div>
              )}

              {/* ========================================================================= */}
              {/* 2. KISIM: TESPİT EDİLEN LİMİT AŞIM ARALIK VE SÜRELERİ (Seçilen Tür) */}
              {/* ========================================================================= */}
              {selectedParamKey && (
                <div className="p-4 bg-slate-100/70 border-b border-slate-200 shrink-0 animate-fade-in">
                  <div className="flex items-center justify-between mb-3">
                    <div className="flex items-center gap-2">
                      <span className="w-6 h-6 rounded-lg bg-amber-700 text-white text-xs font-black flex items-center justify-center shadow-xs">
                        2
                      </span>
                      <div>
                        <h3 className="text-xs sm:text-sm font-black text-slate-900 uppercase tracking-tight flex items-center gap-2">
                          <span>2. KISIM: {selectedColConfig?.name} Tespit Edilen Limit Aşım Aralık ve Süreleri</span>
                          <span className="text-[11px] font-mono font-bold text-slate-500">
                            ({currentParamEpisodes.length} Olay Tespit Edildi)
                          </span>
                        </h3>
                        <p className="text-[11px] text-slate-500">
                          Aşağıdaki aralıklardan birine basarak o aralığa denk gelen orijinal Excel satır kayıtlarını 3. Kısımda açabilirsiniz.
                        </p>
                      </div>
                    </div>

                    {selectedEpisodeId && (
                      <button
                        type="button"
                        onClick={() => setSelectedEpisodeId(null)}
                        className="text-xs font-bold text-red-700 hover:text-red-900 underline cursor-pointer"
                      >
                        ✕ 3. Kısım Aralığını Kapat
                      </button>
                    )}
                  </div>

                  {currentParamEpisodes.length === 0 ? (
                    <div className="p-4 bg-emerald-50 rounded-2xl border border-emerald-200 text-emerald-950 flex flex-wrap items-center justify-between gap-3 text-xs">
                      <div className="flex items-center gap-2.5">
                        <CheckCircle className="w-5 h-5 text-emerald-600 shrink-0" />
                        <div>
                          <strong>Aşım Yok:</strong> Bu uçuş kaydında <strong>{selectedColConfig?.name}</strong> parametresinde herhangi bir limit aşımı tespit edilmemiştir. (Kayıtlarda oluşan en tepe değer: <strong>{getParamData(selectedParamKey).peak}</strong>).
                        </div>
                      </div>
                      <button
                        type="button"
                        onClick={() => setSelectedEpisodeId('ALL_PARAM_ROWS')}
                        className="px-3.5 py-1.5 rounded-xl bg-white border border-emerald-300 text-emerald-800 text-xs font-bold hover:bg-emerald-100 cursor-pointer shadow-2xs"
                      >
                        Tüm Satırları Excel Tablosunda Göster ➔
                      </button>
                    </div>
                  ) : (
                    <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-3">
                      {currentParamEpisodes.map((ep, eIdx) => {
                        const isSelected = selectedEpisodeId === ep.id;
                        const isRed = ep.level === 'RED';
                        return (
                          <div
                            key={ep.id}
                            onClick={() => setSelectedEpisodeId(isSelected ? null : ep.id)}
                            className={`p-3.5 rounded-2xl border-2 transition-all cursor-pointer text-left relative ${
                              isSelected
                                ? isRed
                                  ? 'border-red-600 bg-red-50 ring-2 ring-red-500/40 shadow-lg scale-[1.01]'
                                  : 'border-amber-600 bg-amber-50 ring-2 ring-amber-500/40 shadow-lg scale-[1.01]'
                                : isRed
                                ? 'border-red-300 bg-white hover:border-red-500 hover:bg-red-50/50 shadow-xs'
                                : 'border-amber-400 bg-amber-50/70 hover:border-amber-600 hover:bg-amber-100/60 shadow-xs'
                            }`}
                          >
                            <div className="flex items-center justify-between gap-1 mb-1.5">
                              <span className={`text-[10px] font-black px-2 py-0.5 rounded-md uppercase ${
                                isRed ? 'bg-red-600 text-white' : 'bg-amber-500 text-slate-950'
                              }`}>
                                {isRed ? '🔴 KIRMIZI AŞIM' : '🟡 SARI İKAZ'}
                              </span>
                              <span className="text-xs font-mono font-black text-slate-600">
                                Olay #{eIdx + 1}
                              </span>
                            </div>

                            <div className="text-xs font-black text-slate-900 flex items-center justify-between my-1">
                              <span>{ep.colName}</span>
                              <span className={`font-mono px-1.5 py-0.5 rounded text-xs font-black ${
                                isRed ? 'text-red-700 bg-red-100' : 'text-amber-950 bg-amber-200'
                              }`}>
                                Olay Tepesi: {ep.peakValueFormatted}
                              </span>
                            </div>

                            {/* Time Aralığı Çerçeveli Kutu */}
                            <div className={`mt-2 p-2 rounded-xl border-2 font-mono text-[11px] ${
                              isRed ? 'border-red-500 bg-red-50/60 text-red-950' : 'border-amber-500 bg-amber-100/70 text-amber-950'
                            }`}>
                              <div className={`flex items-center justify-between text-[10px] font-black mb-0.5 ${
                                isRed ? 'text-red-700' : 'text-amber-900'
                              }`}>
                                <span className="flex items-center gap-1">
                                  <Clock className={`w-3 h-3 ${isRed ? 'text-red-600' : 'text-amber-700'}`} />
                                  <span>Zaman Aralığı (TIME):</span>
                                </span>
                                <span>{ep.rowCount} Satır</span>
                              </div>
                              <div className={`text-xs font-black tracking-tight ${
                                isRed ? 'text-red-900' : 'text-amber-950'
                              }`}>
                                {ep.startTimeStr} ➔ {ep.endTimeStr}
                              </div>
                            </div>

                            <div className="mt-2 flex items-center justify-between text-[11px] font-black">
                              <span className="text-slate-600 text-[10px]">Devam Eden Süre:</span>
                              <span className={`px-2 py-0.5 rounded-md font-mono text-[11px] shadow-2xs flex items-center gap-1 ${
                                isRed
                                  ? 'bg-red-600 text-white'
                                  : ep.durationExceeded
                                  ? 'bg-amber-500 text-slate-950 font-black'
                                  : 'bg-amber-200 text-amber-950 animate-pulse border border-amber-400'
                              }`}>
                                <span>⏱ {ep.durationFormatted}</span>
                                <span className="text-[9px] font-bold opacity-90">
                                  {isRed ? '' : ep.durationExceeded ? '(Süre Aşıldı)' : '(Anlık / 20sn Altı)'}
                                </span>
                              </span>
                            </div>
                            <div className="mt-1 flex items-center justify-between text-[10px] text-slate-500 font-mono">
                              <span>Excel Satır No:</span>
                              <strong className="text-slate-800">Satır #{ep.startRow} - #{ep.endRow}</strong>
                            </div>

                            <div className="mt-2.5 pt-2 border-t border-slate-200/80 flex items-center justify-between text-xs font-bold">
                              <span className={isRed ? 'text-red-700' : 'text-amber-900'}>
                                {isSelected ? '✓ 3. Kısımda Açık' : 'Orijinal Excel Satırlarını Getir'}
                              </span>
                              <span className={isRed ? 'text-red-700' : 'text-amber-900'}>➔</span>
                            </div>

                            {/* Olay Takip Çizelgesine Ekle Butonu (Seçilen Olay Üzerinde, 2. Kısım) */}
                            {isSelected && (
                              <div className="mt-3 pt-2.5 border-t border-dashed border-slate-300">
                                <button
                                  type="button"
                                  onClick={(e) => {
                                    e.stopPropagation();
                                    setIsIntegrateModalOpen(true);
                                  }}
                                  className="w-full py-1.5 px-3 rounded-xl bg-gradient-to-r from-red-600 via-orange-600 to-amber-600 hover:from-red-700 text-white text-xs font-black uppercase tracking-wider transition-all cursor-pointer shadow-xs flex items-center justify-center gap-1.5"
                                  title="Seçilen bu olayı Olay Takip Çizelgesine ekle"
                                >
                                  <Plane className="w-3.5 h-3.5 text-white" />
                                  <span>Olay Takip Çizelgesine Ekle</span>
                                </button>
                              </div>
                            )}
                          </div>
                        );
                      })}
                    </div>
                  )}
                </div>
              )}

              {/* ========================================================================= */}
              {/* 3. KISIM: ORİJİNAL EXCELDEKİ SATIR ARALIĞI (2. Kısımda Aralığa Basıldığında) */}
              {/* ========================================================================= */}
              {selectedParamKey && !selectedEpisodeId && (
                <div className="flex-1 flex items-center justify-center p-8 text-center bg-white border-t border-slate-200 animate-fade-in">
                  <div className="max-w-md mx-auto">
                    <div className="w-10 h-10 mx-auto rounded-2xl bg-red-100 text-red-800 flex items-center justify-center font-black text-sm mb-2 shadow-inner" style={{display: 'none'}}>
                      3
                    </div>
                    <h4 className="text-xs font-black text-slate-800 uppercase tracking-wider mb-1" style={{display: 'none'}}>
                      3. Kısım: Orijinal Excel Satır Kayıtları
                    </h4>
                    <p className="text-xs text-slate-500 mb-3" style={{display: 'none'}}>
                      2. Kısımdaki limit aşım aralıklarından birine basarak o aralığa denk gelen orijinal Excel satır kayıtlarını ve satır numaralarını inceleyiniz.
                    </p>
                    <button
                      type="button"
                      onClick={() => setSelectedEpisodeId('ALL_PARAM_ROWS')}
                      className="px-3.5 py-1.5 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-800 text-xs font-bold transition-all cursor-pointer shadow-2xs"
                    >
                      Veya Bu Parametrenin Tüm Satırlarını Tabloda Göster ➔
                    </button>
                  </div>
                </div>
              )}

              {selectedParamKey && selectedEpisodeId && (
                <div className="flex-1 flex flex-col min-h-0 bg-white">
                  {/* 3. KISIM BAŞLIK ŞERİDİ */}
                  <div className="p-3 bg-red-50 border-b border-red-200 flex flex-wrap items-center justify-between gap-2 shrink-0">
                    <div className="flex items-center gap-2">
                      <span className="w-6 h-6 rounded-lg bg-red-700 text-white text-xs font-black flex items-center justify-center shadow-xs">
                        3
                      </span>
                      <div>
                        <h4 className="text-xs sm:text-sm font-black text-red-950 uppercase tracking-tight flex flex-wrap items-center gap-2">
                          <span>
                            3. KISIM: Orijinal Excel Satır Kayıtları — {selectedEpisode ? `Satır #${selectedEpisode.startRow} - #${selectedEpisode.endRow}` : 'Tüm Satırlar'}
                          </span>
                          {selectedEpisode && (
                            <span className="text-[11px] font-mono font-bold text-red-800 bg-red-100 px-2 py-0.5 rounded-md border border-red-200">
                              {selectedEpisode.colName} | {selectedEpisode.startTimeStr} ➔ {selectedEpisode.endTimeStr} (⏱ {selectedEpisode.durationFormatted})
                            </span>
                          )}
                        </h4>
                        <p className="text-[10px] text-slate-600 font-mono">
                          {selectedEpisode ? `Bu aralıkta ${selectedEpisode.rowCount} adet orijinal Excel satırı listelenmektedir. Satır numaraları orijinal dosya satırları ile birebirdir.` : 'Parametreye ait tüm satırlar listelenmektedir.'}
                        </p>
                      </div>
                    </div>
                    <div className="flex items-center gap-2 flex-wrap">

                      {/* Sütun Görünüm Filtresi Butonları (Kullanıcı Talebi: SIRA, İKAZ, TIME ve Seçilen Parametre) */}
                      <div className="flex items-center gap-1 bg-white p-1 rounded-xl border border-red-200 shadow-2xs">
                        <Filter className="w-3.5 h-3.5 text-slate-500 ml-1" />
                        <span className="text-[10px] font-bold text-slate-600 mr-0.5">Sütunlar:</span>
                        <button
                          type="button"
                          onClick={() => setTableColumnFilterMode('SELECTED_ONLY')}
                          className={`px-2 py-0.5 rounded-lg text-xs font-bold transition-all cursor-pointer ${
                            tableColumnFilterMode === 'SELECTED_ONLY'
                              ? 'bg-amber-400 text-slate-950 font-black shadow-xs'
                              : 'text-slate-600 hover:bg-slate-100'
                          }`}
                          title="Sadece SIRA, İKAZ, TIME ve seçilen parametre sütununu göster"
                        >
                          🎯 Sadece {selectedColConfig?.name || 'Seçilen'}
                        </button>
                        <button
                          type="button"
                          onClick={() => setTableColumnFilterMode('ALL_YELLOW')}
                          className={`px-2 py-0.5 rounded-lg text-xs font-bold transition-all cursor-pointer ${
                            tableColumnFilterMode === 'ALL_YELLOW'
                              ? 'bg-amber-400 text-slate-950 font-black shadow-xs'
                              : 'text-slate-600 hover:bg-slate-100'
                          }`}
                          title="7 Sarı Odak parametresini göster"
                        >
                          🟡 7 Parametre
                        </button>
                        <button
                          type="button"
                          onClick={() => setTableColumnFilterMode('ALL_COLUMNS')}
                          className={`px-2 py-0.5 rounded-lg text-xs font-bold transition-all cursor-pointer ${
                            tableColumnFilterMode === 'ALL_COLUMNS'
                              ? 'bg-amber-400 text-slate-950 font-black shadow-xs'
                              : 'text-slate-600 hover:bg-slate-100'
                          }`}
                          title="Orijinal dosyadaki tüm sütunları aç"
                        >
                          📋 Tüm Sütunlar ({analysisResult?.headers.length})
                        </button>
                      </div>

                      <button
                        type="button"
                        onClick={() => {
                          if (!analysisResult) return;
                          const rowsToExport = filteredRows;
                          const headers = activeHeaders;
                          const indices = activeHeaderIndices;

                          let html = `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">`;
                          html += `<head><meta http-equiv="content-type" content="text/html; charset=UTF-8"><style>`;
                          html += `table { border-collapse: collapse; width: 100%; font-family: Arial, sans-serif; font-size: 10pt; }`;
                          html += `th { background-color: #cbd5e1; color: #000000; font-weight: bold; vertical-align: middle; padding: 8px 10px; border: 1px solid #94a3b8; text-align: center; }`;
                          html += `td { padding: 6px 8px; border: 1px solid #cbd5e1; text-align: left; }`;
                          html += `.row-red { background-color: #fee2e2; color: #991b1b; font-weight: bold; }`;
                          html += `.row-yellow { background-color: #fef9c3; color: #713f12; font-weight: bold; }`;
                          html += `.row-normal { background-color: #ffffff; color: #0f172a; }`;
                          html += `</style></head><body>`;

                          // Uçak & Kayıt Tanıtım Bilgileri (Ham CSV ilk 15 Metadata Satırı)
                          html += `<table style="border-collapse: collapse; border: 2px solid #0f172a; margin-bottom: 20px; width: 550px;">`;
                          html += `<tr><td colspan="2" style="background-color: #1e293b; color: #ffffff; font-weight: bold; text-align: center; padding: 8px; border: 1px solid #0f172a; font-size: 11pt;">Uçak & Kayıt Tanıtım Bilgileri (Metadata Satırları)</td></tr>`;
                          
                          const rawMetaLines = (analysisResult.metadataLines || []).slice(0, 15);
                          if (rawMetaLines.length > 0) {
                            rawMetaLines.forEach((mLine, idx) => {
                              const cleanedLine = String(mLine).replace(/[;]+$/, '').trim();
                              let label = `Metadata Satır ${idx + 1}`;
                              let val = cleanedLine;
                              if (cleanedLine.includes(':')) {
                                const parts = cleanedLine.split(':');
                                label = parts[0].trim();
                                val = parts.slice(1).join(':').replace(/[;]+$/, '').trim();
                              } else if (idx === 0) {
                                label = 'Description';
                              } else if (idx === 1) {
                                label = 'Model';
                              }
                              html += `<tr><td style="font-weight: bold; background-color: #f8fafc; padding: 6px 10px; border: 1px solid #cbd5e1; width: 220px;">${String(label).replace(/</g, '&lt;').replace(/>/g, '&gt;')}</td><td style="padding: 6px 10px; border: 1px solid #cbd5e1; mso-number-format:'\\@';">${String(val).replace(/</g, '&lt;').replace(/>/g, '&gt;')}</td></tr>`;
                            });
                          } else {
                            html += `<tr><td colspan="2" style="padding: 6px 10px; border: 1px solid #cbd5e1; text-align: center;">Metadata satırı bulunamadı</td></tr>`;
                          }
                          html += `</table><br/>`;

                          html += `<table><thead><tr>`;
                          html += `<th>SATIR NO</th><th>DURUM</th>`;
                          headers.forEach(h => {
                            html += `<th>${String(h || '').replace(/</g, '&lt;').replace(/>/g, '&gt;')}</th>`;
                          });
                          html += `</tr></thead><tbody>`;

                          rowsToExport.forEach(r => {
                            const isRed = r.severity === 'RED';
                            const isYellow = r.severity === 'YELLOW';
                            const statText = isRed ? 'KIRMIZI AŞIM' : isYellow ? 'SARI İKAZ' : 'Normal';
                            const statStyle = isRed ? 'background-color: #dc2626; color: #ffffff; font-weight: bold; text-align: center;' : isYellow ? 'background-color: #fef08a; color: #000000; font-weight: bold; text-align: center;' : 'text-align: center;';

                            html += `<tr>`;
                            html += `<td style="text-align: center; font-weight: bold;">#${r.originalIndex}</td>`;
                            html += `<td style="${statStyle}">${statText}</td>`;
                            indices.forEach(origColIdx => {
                              const cellVal = r.cells[origColIdx] !== undefined ? r.cells[origColIdx] : '';
                              const safeVal = String(cellVal || '').replace(/</g, '&lt;').replace(/>/g, '&gt;');
                              const violation = r.violations.find(v => v.colIndex === origColIdx);
                              const cellStyle = violation?.level === 'RED'
                                ? 'background-color: #dc2626; color: #ffffff; font-weight: bold;'
                                : violation?.level === 'YELLOW'
                                ? 'background-color: #fef08a; color: #000000; font-weight: bold;'
                                : '';
                              html += `<td style="${cellStyle}">${safeVal}</td>`;
                            });
                            html += `</tr>`;
                          });

                          html += `</tbody></table></body></html>`;

                          const blob = new Blob([html], { type: 'application/vnd.ms-excel;charset=utf-8;' });
                          const url = URL.createObjectURL(blob);
                          const link = document.createElement('a');
                          link.setAttribute('href', url);
                          link.setAttribute('download', `${selectedTail.fullLabel}_${selectedColConfig?.key || 'analiz'}_filtrelenmis_excel.xls`);
                          document.body.appendChild(link);
                          link.click();
                          document.body.removeChild(link);
                          showNotification('Seçilen filtrelenmiş sütunlara sahip Excel raporu (.xls) başarıyla indirildi!', 'success');
                        }}
                        className="px-3.5 py-1.5 rounded-xl bg-emerald-700 hover:bg-emerald-800 text-white text-xs font-black transition-all cursor-pointer shadow-xs flex items-center gap-1.5"
                        title="Tabloyu yalnızca seçilen ve filtrelenmiş sütunlarla Excel (.xls) formatında indir"
                      >
                        <Download className="w-3.5 h-3.5 text-white" />
                        <span>Filtrelenmiş Excel İndir (.xls)</span>
                      </button>



                      <button
                        type="button"
                        onClick={() => setSelectedEpisodeId(null)}
                        className="px-3 py-1 rounded-xl bg-white border border-red-300 text-red-800 text-xs font-bold hover:bg-red-100 transition-all cursor-pointer shadow-2xs"
                      >
                        ✕ Aralığı Kapat
                      </button>
                    </div>
                  </div>

                  {/* TABLO GÖRÜNÜMÜ: ORİJİNAL EXCEL TASARIMI */}
                  <div className="w-full overflow-x-auto flex-1">
                    <table className="min-w-max w-full text-left border-collapse bg-white border border-slate-300">
                      <thead className="sticky top-0 z-20 shadow-xs font-mono">
                        {/* EXCEL ÜST SATIR: SÜTUN HARFLERİ (A, B, C, D... AR) */}
                        <tr className="border-b border-slate-300 text-[10px]">
                          <th className="px-2 py-1 text-center font-black text-slate-700 uppercase tracking-wider border-r border-b border-slate-300 bg-slate-200 sticky left-0 z-30 w-14">
                            #
                          </th>
                          <th className={`px-2 py-1 text-center font-black uppercase tracking-wider border-r border-b border-slate-300 w-28 ${
                            analysisResult?.redCount > 0 ? 'bg-red-600 text-white' : analysisResult?.yellowCount > 0 ? 'bg-amber-500 text-slate-950' : 'bg-slate-200 text-slate-700'
                          }`}>
                            DURUM
                          </th>
                          {activeHeaders.map((hdr, hIdx) => {
                            const origColIdx = activeHeaderIndices[hIdx];
                            const isSelectedParamCol = selectedColConfig
                              ? (origColIdx === selectedColConfig.expectedIndex || selectedColConfig.matchPattern.test(hdr) || hdr.toUpperCase().includes(selectedColConfig.name.toUpperCase()))
                              : false;
                            const excelLetter = getExcelColumnLetter(origColIdx);

                            return (
                              <th
                                key={`letter-${hIdx}`}
                                className={`px-3 py-1 text-center border-r border-b transition-colors ${
                                  isSelectedParamCol
                                    ? 'bg-[#facc15] text-slate-950 border-[#ca8a04] font-black ring-2 ring-amber-400 shadow-sm'
                                    : 'bg-slate-200 text-slate-700 border-slate-300 font-bold'
                                }`}
                              >
                                <span className="font-mono text-xs font-black">
                                  {excelLetter}
                                </span>
                              </th>
                            );
                          })}
                        </tr>

                        {/* EXCEL BAŞLIK SATIRI: PARAMETRE ADLARI (YALNIZCA SEÇİLEN PARAMETREDE SARI DOLGU) */}
                        <tr className="border-b border-slate-300">
                          <th className="px-2 py-2 text-[10px] font-black text-slate-700 uppercase tracking-wider border-r border-b border-slate-300 text-center bg-slate-100 sticky left-0 z-30">
                            SIRA
                          </th>
                          <th className={`px-2 py-2 text-[10px] font-black uppercase tracking-wider border-r border-b border-slate-300 text-center ${
                            analysisResult?.redCount > 0 ? 'bg-red-700 text-white' : analysisResult?.yellowCount > 0 ? 'bg-amber-600 text-slate-950' : 'bg-slate-100 text-slate-700'
                          }`}>
                            İKAZ
                          </th>

                          {activeHeaders.map((hdr, hIdx) => {
                            const origColIdx = activeHeaderIndices[hIdx];
                            const isSelectedParamCol = selectedColConfig
                              ? (origColIdx === selectedColConfig.expectedIndex || selectedColConfig.matchPattern.test(hdr) || hdr.toUpperCase().includes(selectedColConfig.name.toUpperCase()))
                              : false;

                            return (
                              <th
                                key={`header-${hIdx}`}
                                style={{ minWidth: isSelectedParamCol ? '150px' : hdr.toUpperCase().includes('TIME') ? '230px' : '95px' }}
                                className={`px-3 py-3 text-[11px] uppercase tracking-wider border-r border-b whitespace-nowrap text-center align-middle transition-all ${
                                  isSelectedParamCol
                                    ? 'bg-[#fef08a] text-black border-[#ca8a04] font-bold shadow-md ring-2 ring-amber-400'
                                    : 'bg-slate-300 text-black border-slate-400 font-bold'
                                }`}
                              >
                                <div className="flex flex-col items-center justify-center py-0.5">
                                  <span className="font-bold text-[11px] text-black tracking-tight">{hdr}</span>
                                  {isSelectedParamCol && (
                                    <span className="text-[8px] bg-amber-950 text-amber-200 px-2 py-0.5 rounded font-bold mt-0.5 tracking-wider shadow-2xs">
                                      ★ SEÇİLEN ODAK
                                    </span>
                                  )}
                                </div>
                              </th>
                            );
                          })}
                        </tr>
                      </thead>

                      <tbody className="divide-y divide-slate-100 font-mono text-[11px]">
                        {paginatedRows.length === 0 ? (
                          <tr>
                            <td colSpan={activeHeaders.length + 2} className="px-6 py-12 text-center text-slate-400">
                              Seçilen filtreye uygun kayıt bulunamadı.
                            </td>
                          </tr>
                        ) : (
                          paginatedRows.map((row, rIdx) => {
                            const isRed = row.severity === 'RED' || row.violations.some(v => v.level === 'RED');
                            const isYellow = !isRed && (row.severity === 'YELLOW' || row.violations.some(v => v.level === 'YELLOW'));
                            const isViolatedRow = isRed || isYellow;
                            const ep = row.primaryEpisode;
                            const primaryViolation = row.violations[0];

                            return (
                              <tr
                                key={rIdx}
                                className="hover:bg-slate-50 border-b border-slate-100 text-slate-700 transition-colors bg-white"
                              >
                                {/* Satır sırası (Excel orijinal satır no birebir aynı!) */}
                                <td className="px-3 py-2 text-center font-bold text-slate-700 border-r border-b border-slate-300 bg-slate-100 sticky left-0 z-10 font-mono text-xs">
                                  #{row.originalIndex}
                                </td>

                                {/* İkaz Durumu Rozeti ve Hücresi */}
                                <td className={`px-3 py-2 text-center border-r border-b border-slate-200 whitespace-nowrap ${
                                  isRed ? 'bg-red-600 text-white font-black' : isYellow ? 'bg-amber-400 text-slate-950 font-black' : 'bg-white text-slate-700'
                                }`}>
                                  {isRed ? (
                                    <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-black bg-red-700 text-white shadow-2xs">
                                      <Flame className="w-3 h-3 text-red-200 animate-pulse" />
                                      <span>KIRMIZI AŞIM</span>
                                    </span>
                                  ) : isYellow ? (
                                    <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-black bg-amber-500 text-slate-950 shadow-2xs">
                                      <AlertTriangle className="w-3 h-3 text-amber-950" />
                                      <span>SARI İKAZ</span>
                                    </span>
                                  ) : (
                                    <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-bold bg-slate-100 text-slate-600">
                                      <span>Normal</span>
                                    </span>
                                  )}
                                </td>

                                {/* Sütun Hücreleri (Excel hücreleri) */}
                                {activeHeaders.map((hdr, hIdx) => {
                                  const origColIdx = activeHeaderIndices[hIdx];
                                  const cellVal = row.cells[origColIdx] !== undefined ? row.cells[origColIdx] : '-';
                                  const isSelectedParamCol = selectedColConfig
                                    ? (origColIdx === selectedColConfig.expectedIndex || selectedColConfig.matchPattern.test(hdr) || hdr.toUpperCase().includes(selectedColConfig.name.toUpperCase()))
                                    : false;

                                  // TIME SÜTUNU: SARI İSE SARI ÇERÇEVELİ "SARI İKAZ ZAMAN ARALIĞI", KIRMIZI İSE KIRMIZI ÇERÇEVELİ "KIRMIZI AŞIM"
                                  if (hdr.toUpperCase().includes('TIME') || origColIdx === 0) {
                                    return (
                                      <td
                                        key={hIdx}
                                        style={{ minWidth: isViolatedRow ? '240px' : '110px' }}
                                        className={`px-2 py-1.5 border-r border-b border-slate-200 whitespace-nowrap font-mono text-center transition-colors ${
                                          isRed ? 'bg-red-50/80' : isYellow ? 'bg-amber-50/80' : 'bg-slate-50 text-slate-800'
                                        }`}
                                      >
                                        {isRed ? (
                                          <div className="border-2 border-red-600 bg-red-50/90 rounded-xl p-2 shadow-xs ring-2 ring-red-400/40 text-left">
                                            <div className="flex items-center justify-between gap-1.5 pb-1 border-b border-red-200">
                                              <span className="font-mono font-black text-xs text-red-950 flex items-center gap-1">
                                                <Clock className="w-3.5 h-3.5 text-red-600 shrink-0" />
                                                <span>{cellVal}</span>
                                              </span>
                                              <span className="px-1.5 py-0.2 rounded text-[9px] font-black bg-red-600 text-white uppercase tracking-wider animate-pulse">
                                                🔴 KIRMIZI AŞIM
                                              </span>
                                            </div>

                                            {(ep?.startTimeStr || primaryViolation?.timeRangeStr) && (
                                              <div className="mt-1 text-[10px] font-bold text-red-900 bg-red-100/70 px-1.5 py-0.5 rounded border border-red-300 flex items-center gap-1">
                                                <span className="text-red-800 font-black">Kırmızı Aşım Zaman Aralığı:</span>
                                                <span className="font-mono font-black tracking-tight">
                                                  {ep ? `${ep.startTimeStr} ➔ ${ep.endTimeStr}` : primaryViolation?.timeRangeStr}
                                                </span>
                                              </div>
                                            )}

                                            <div className="mt-1 flex items-center gap-1.5 text-[10px] font-black">
                                              <span className="px-1.5 py-0.2 rounded bg-red-700 text-white font-mono shrink-0">
                                                ⏱ {ep?.durationFormatted || primaryViolation?.durationFormatted || '0 sn'}
                                              </span>
                                              <span
                                                className="text-red-950 truncate font-mono font-black"
                                                title={ep?.summaryText || primaryViolation?.summaryText}
                                              >
                                                {ep?.summaryText || primaryViolation?.summaryText || primaryViolation?.colName}
                                              </span>
                                            </div>
                                          </div>
                                        ) : isYellow ? (
                                          <div className="border-2 border-amber-500 bg-amber-50/95 rounded-xl p-2 shadow-xs ring-2 ring-amber-300/60 text-left">
                                            <div className="flex items-center justify-between gap-1.5 pb-1 border-b border-amber-200">
                                              <span className="font-mono font-black text-xs text-amber-950 flex items-center gap-1">
                                                <Clock className="w-3.5 h-3.5 text-amber-700 shrink-0" />
                                                <span>{cellVal}</span>
                                              </span>
                                              <span className="px-1.5 py-0.2 rounded text-[9px] font-black bg-amber-500 text-slate-950 uppercase tracking-wider">
                                                🟡 SARI İKAZ
                                              </span>
                                            </div>

                                            {(ep?.startTimeStr || primaryViolation?.timeRangeStr) && (
                                              <div className="mt-1 text-[10px] font-bold text-amber-950 bg-amber-100/90 px-1.5 py-0.5 rounded border border-amber-300 flex items-center gap-1">
                                                <span className="text-amber-900 font-black">Sarı İkaz Zaman Aralığı:</span>
                                                <span className="font-mono font-black tracking-tight">
                                                  {ep ? `${ep.startTimeStr} ➔ ${ep.endTimeStr}` : primaryViolation?.timeRangeStr}
                                                </span>
                                              </div>
                                            )}

                                            <div className="mt-1 flex items-center gap-1.5 text-[10px] font-black">
                                              <span className="px-1.5 py-0.2 rounded bg-amber-500 text-slate-950 font-mono font-black shrink-0">
                                                ⏱ {ep?.durationFormatted || primaryViolation?.durationFormatted || '0 sn'}
                                              </span>
                                              <span
                                                className="text-amber-950 truncate font-mono font-black"
                                                title={ep?.summaryText || primaryViolation?.summaryText}
                                              >
                                                {ep?.summaryText || primaryViolation?.summaryText || primaryViolation?.colName}
                                              </span>
                                            </div>
                                          </div>
                                        ) : (
                                          <span className="font-mono font-semibold text-slate-700">{cellVal}</span>
                                        )}
                                      </td>
                                    );
                                  }

                                  // DİĞER SÜTUNLAR: YALNIZCA SEÇİLEN PARAMETRE SÜTUNUNDA DİKKAT ÇEKİCİ SARI DOLGU
                                  const violation = row.violations.find(
                                    v => v.colIndex === origColIdx || (selectedColConfig && isYellowFocusHeader(v.colName, origColIdx))
                                  );

                                  return (
                                    <td
                                      key={hIdx}
                                      style={{ minWidth: isSelectedParamCol ? '150px' : '95px' }}
                                      className={`px-3 py-1.5 text-center border-r border-b border-slate-200 whitespace-nowrap transition-colors font-mono text-[11px] ${
                                        violation || isSelectedParamCol
                                          ? violation?.level === 'RED'
                                            ? 'bg-red-600 text-white font-black shadow-inner ring-1 ring-red-400'
                                            : violation?.level === 'YELLOW'
                                            ? 'bg-amber-300 text-slate-950 font-black border-2 border-amber-500 shadow-inner ring-1 ring-amber-400'
                                            : isSelectedParamCol
                                            ? 'bg-amber-100/80 text-slate-950 font-black border-x-2 border-amber-400 shadow-2xs'
                                            : 'text-slate-800 bg-white hover:bg-slate-50'
                                          : 'text-slate-800 bg-white hover:bg-slate-50'
                                      }`}
                                      title={violation ? violation.rule : undefined}
                                    >
                                      <div className="flex flex-col items-center justify-center">
                                        <span className={`font-bold ${isSelectedParamCol ? 'text-xs' : ''}`}>{cellVal}</span>
                                        {violation && (
                                          <span
                                            className={`text-[9px] px-1 rounded font-black mt-0.5 ${
                                              violation.level === 'RED'
                                                ? 'bg-red-950 text-red-200'
                                                : 'bg-amber-950 text-amber-200'
                                            }`}
                                          >
                                            {violation.summaryText || `${violation.durationSeconds?.toFixed(1) || '0'}s`}
                                          </span>
                                        )}
                                      </div>
                                    </td>
                                  );
                                })}
                              </tr>
                            );
                          })
                        )}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}
            </div>
          )}
        </div>

        {/* SAYFALAMA VE TABLO BİLGİ ALTI */}
        {analysisResult && (
          <div className="p-3 bg-white border-t border-slate-200 flex flex-wrap items-center justify-between gap-3 text-xs shrink-0 font-mono">
            <div className="flex items-center gap-3 text-slate-600">
              <span>
                Toplam Satır: <strong>{analysisResult.rows.length.toLocaleString('tr-TR')}</strong>
              </span>
              <span>|</span>
              <span>
                Filtrelenen: <strong>{filteredRows.length.toLocaleString('tr-TR')}</strong>
              </span>
              <span>|</span>
              <span className="text-red-700 font-bold">
                Kırmızı: {analysisResult.redCount}
              </span>
              <span>|</span>
              <span className="text-amber-700 font-bold">
                Sarı: {analysisResult.yellowCount}
              </span>
              {analysisResult.durationExceededCount > 0 && (
                <>
                  <span>|</span>
                  <span className="text-purple-700 font-bold">
                    Süre Aşımı: {analysisResult.durationExceededCount}
                  </span>
                </>
              )}
            </div>

            <div className="flex items-center gap-2">
              <span className="text-slate-500">Sayfa Başına:</span>
              <select
                value={pageSize}
                onChange={e => setPageSize(Number(e.target.value))}
                className="px-2 py-1 rounded-lg border border-slate-200 text-xs bg-slate-50 font-bold cursor-pointer"
              >
                <option value={25}>25 Satır</option>
                <option value={50}>50 Satır</option>
                <option value={100}>100 Satır</option>
                <option value={250}>250 Satır</option>
              </select>

              <span className="text-slate-400 mx-1">|</span>

              <span className="text-slate-600 font-bold">
                Sayfa {validCurrentPage} / {totalPages}
              </span>

              <button
                type="button"
                disabled={validCurrentPage <= 1}
                onClick={() => setCurrentPage(1)}
                className="p-1.5 rounded-lg border border-slate-200 hover:bg-slate-100 disabled:opacity-40 disabled:cursor-not-allowed cursor-pointer"
                title="İlk Sayfa"
              >
                <ChevronsLeft className="w-4 h-4 text-slate-700" />
              </button>
              <button
                type="button"
                disabled={validCurrentPage <= 1}
                onClick={() => setCurrentPage(p => Math.max(p - 1, 1))}
                className="p-1.5 rounded-lg border border-slate-200 hover:bg-slate-100 disabled:opacity-40 disabled:cursor-not-allowed cursor-pointer"
                title="Önceki Sayfa"
              >
                <ChevronLeft className="w-4 h-4 text-slate-700" />
              </button>

              <span className="px-3 py-1 bg-slate-100 rounded-lg font-mono font-bold text-slate-800">
                {validCurrentPage}
              </span>

              <button
                type="button"
                disabled={validCurrentPage >= totalPages}
                onClick={() => setCurrentPage(p => Math.min(p + 1, totalPages))}
                className="p-1.5 rounded-lg border border-slate-200 hover:bg-slate-100 disabled:opacity-40 disabled:cursor-not-allowed cursor-pointer"
                title="Sonraki Sayfa"
              >
                <ChevronRight className="w-4 h-4 text-slate-700" />
              </button>
              <button
                type="button"
                disabled={validCurrentPage >= totalPages}
                onClick={() => setCurrentPage(totalPages)}
                className="p-1.5 rounded-lg border border-slate-200 hover:bg-slate-100 disabled:opacity-40 disabled:cursor-not-allowed cursor-pointer"
                title="Son Sayfa"
              >
                <ChevronsRight className="w-4 h-4 text-slate-700" />
              </button>
            </div>
          </div>
        )}
      </div>

      {/* ========================================================================= */}
      {/* OLAY TAKİP ÇİZELGESİNE AKTARIM FORMU MODALI */}
      {/* ========================================================================= */}
      {isIntegrateModalOpen && (
        <div className="fixed inset-0 z-[10015] flex items-center justify-center p-3 sm:p-4 bg-slate-950/80 backdrop-blur-sm animate-fade-in">
          <div className="bg-white rounded-3xl shadow-2xl border border-slate-200 w-full max-w-xl overflow-hidden flex flex-col">
            <div className="p-5 bg-gradient-to-r from-red-700 via-orange-600 to-amber-600 text-white flex items-center justify-between">
              <div className="flex items-center gap-3">
                <div className="w-9 h-9 rounded-xl bg-white/20 flex items-center justify-center text-white">
                  <Plane className="w-5 h-5 text-white" />
                </div>
                <div>
                  <div className="text-[10px] uppercase font-black tracking-widest text-amber-200">
                    SİSTEME KAYIT
                  </div>
                  <h3 className="text-base font-black text-white">
                    Olay Takip Çizelgesine Kaydet
                  </h3>
                </div>
              </div>
              <button
                type="button"
                onClick={() => setIsIntegrateModalOpen(false)}
                className="text-white/80 hover:text-white cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="p-5 space-y-4 max-h-[75vh] overflow-y-auto text-xs">
              <div className="p-3 bg-red-50 rounded-2xl border border-red-200 text-red-950 leading-relaxed">
                Bu analizdeki kritik motor limit aşımı ve uyarı kayıtları, AT-802 uçağınızın{' '}
                <strong>"Ng Limit Aşımları Detayı"</strong> çizelgesine yeni bir satır olarak kaydedilecektir.
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-slate-600 font-bold mb-1">Kuyruk Kodu:</label>
                  <input
                    type="text"
                    value={integrateForm.kuyrukNo}
                    onChange={e => setIntegrateForm({ ...integrateForm, kuyrukNo: e.target.value })}
                    className="w-full px-3 py-2 rounded-xl border border-slate-200 font-mono font-bold bg-slate-50"
                  />
                </div>
                <div>
                  <label className="block text-slate-600 font-bold mb-1">Tarih:</label>
                  <input
                    type="text"
                    value={integrateForm.tarih}
                    onChange={e => setIntegrateForm({ ...integrateForm, tarih: e.target.value })}
                    className="w-full px-3 py-2 rounded-xl border border-slate-200 font-mono"
                  />
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-slate-600 font-bold mb-1">Motor Seri No:</label>
                  <input
                    type="text"
                    value={integrateForm.motorSeriNo}
                    onChange={e => setIntegrateForm({ ...integrateForm, motorSeriNo: e.target.value })}
                    className="w-full px-3 py-2 rounded-xl border border-slate-200 font-mono font-bold text-emerald-900 bg-slate-50"
                  />
                </div>
                <div>
                  <label className="block text-slate-600 font-bold mb-1">Tepe Ng Değeri:</label>
                  <input
                    type="text"
                    value={integrateForm.tepeNg}
                    onChange={e => setIntegrateForm({ ...integrateForm, tepeNg: e.target.value })}
                    className="w-full px-3 py-2 rounded-xl border border-slate-200 font-mono font-bold text-red-700 bg-red-50/50"
                  />
                </div>
              </div>

              <div>
                <label className="block text-slate-600 font-bold mb-1">Olay / İşlem Türü:</label>
                <input
                  type="text"
                  value={integrateForm.olayTuru}
                  onChange={e => setIntegrateForm({ ...integrateForm, olayTuru: e.target.value })}
                  className="w-full px-3 py-2 rounded-xl border border-slate-200 font-bold"
                />
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-slate-600 font-bold mb-1">DAA Referans No:</label>
                  <input
                    type="text"
                    value={integrateForm.daaRefNo}
                    onChange={e => setIntegrateForm({ ...integrateForm, daaRefNo: e.target.value })}
                    className="w-full px-3 py-2 rounded-xl border border-slate-200 font-mono"
                  />
                </div>
                <div>
                  <label className="block text-slate-600 font-bold mb-1">Verilen Şart / Kontrol:</label>
                  <input
                    type="text"
                    value={integrateForm.verilenSart}
                    onChange={e => setIntegrateForm({ ...integrateForm, verilenSart: e.target.value })}
                    className="w-full px-3 py-2 rounded-xl border border-slate-200"
                  />
                </div>
              </div>

              <div>
                <label className="block text-slate-600 font-bold mb-1">Sorumlu Pilot:</label>
                <input
                  type="text"
                  value={integrateForm.pilot}
                  onChange={e => setIntegrateForm({ ...integrateForm, pilot: e.target.value })}
                  className="w-full px-3 py-2 rounded-xl border border-slate-200"
                />
              </div>

              <div>
                <label className="block text-slate-600 font-bold mb-1">Teknik Değerlendirme & Açıklama:</label>
                <textarea
                  rows={3}
                  value={integrateForm.teknikDegerlendirme}
                  onChange={e => setIntegrateForm({ ...integrateForm, teknikDegerlendirme: e.target.value })}
                  className="w-full px-3 py-2 rounded-xl border border-slate-200 leading-relaxed font-sans"
                />
              </div>

              <div>
                <label className="block text-slate-600 font-bold mb-1">Parça / Uçak Durumu:</label>
                <select
                  value={integrateForm.parcaDurumu}
                  onChange={e => setIntegrateForm({ ...integrateForm, parcaDurumu: e.target.value })}
                  className="w-full px-3 py-2 rounded-xl border border-slate-200 bg-white font-bold"
                >
                  <option value="Faal">Faal</option>
                  <option value="Gözlem Altında Faal">Gözlem Altında Faal</option>
                  <option value="Bakım / Kontrol Gerekli">Bakım / Kontrol Gerekli</option>
                  <option value="Onarıma Alındı">Onarıma Alındı</option>
                  <option value="Gayrifaal">Gayrifaal</option>
                </select>
              </div>
            </div>

            <div className="p-4 bg-slate-50 border-t border-slate-200 flex items-center justify-end gap-2">
              <button
                type="button"
                onClick={() => setIsIntegrateModalOpen(false)}
                className="px-4 py-2 rounded-xl border border-slate-300 bg-white hover:bg-slate-100 text-slate-700 text-xs font-bold cursor-pointer"
              >
                Vazgeç
              </button>
              <button
                type="button"
                onClick={handleCommitToOlayTakip}
                className="px-5 py-2 rounded-xl bg-red-600 hover:bg-red-700 text-white font-black text-xs uppercase tracking-wider flex items-center gap-2 shadow-md cursor-pointer transition-all active:scale-95"
              >
                <Save className="w-4 h-4" />
                <span>Çizelgeye Ekle</span>
              </button>
            </div>
          </div>
        </div>
      )}


      {/* Anlık (Süre Koşulunu Sağlamayan) Tepe Değer Uyarı Penceresi */}
      {transientWarningPopup && transientWarningPopup.isOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/60 backdrop-blur-xs p-4 animate-fade-in">
          <div className="bg-white rounded-3xl shadow-2xl border-2 border-amber-500 w-full max-w-2xl overflow-hidden flex flex-col">
            <div className="p-4 bg-amber-500 text-slate-950 flex items-center justify-between">
              <div className="flex items-center gap-2">
                <AlertTriangle className="w-5 h-5 text-slate-950 animate-bounce" />
                <h3 className="text-sm font-black uppercase tracking-wider">
                  Anlık Tepe Değer Bilgilendirmesi (Süre Koşulu Sağlanmadı)
                </h3>
              </div>
              <button
                type="button"
                onClick={() => setTransientWarningPopup(null)}
                className="w-7 h-7 rounded-full bg-slate-950/10 hover:bg-slate-950/20 flex items-center justify-center text-slate-950 font-black cursor-pointer transition-colors"
              >
                ✕
              </button>
            </div>

            <div className="p-5 space-y-4">
              <div className="p-3 bg-amber-50 rounded-2xl border border-amber-200 text-xs text-amber-950 leading-relaxed font-mono">
                İlgili parametre (<strong>{transientWarningPopup.colName}</strong>) en yüksek <strong>{transientWarningPopup.peakValFormatted}</strong> değerini görmüştür, fakat <strong>süre koşulunu (gerekli eşik süreyi) sağlamadığı</strong> için resmi limit aşım epizodlarına dahil edilmemiştir. İlgili anlık satır ve değer aşağıda sarı dolguyla gösterilmektedir.
              </div>

              {/* İlgili Satırın Tablo Görünümü */}
              {transientWarningPopup.row && analysisResult && (
                <div className="w-full overflow-x-auto border border-slate-300 rounded-xl">
                  <table className="min-w-max w-full text-left border-collapse bg-white text-xs font-mono">
                    <thead>
                      <tr className="bg-slate-200 text-slate-700 text-[10px] border-b border-slate-300">
                        <th className="px-2 py-1 text-center border-r border-slate-300 w-14">#</th>
                        <th className="px-2 py-1 text-center border-r border-slate-300 w-28">ZAMAN / DURUM</th>
                        {analysisResult.headers.map((h, hIdx) => {
                          const origColIndex = analysisResult.headers.findIndex(item => item === h);
                          const isTargetCol = h.toUpperCase().includes(transientWarningPopup.colKey.toUpperCase()) ||
                            (transientWarningPopup.colKey === 'np' && h.toUpperCase().includes('NP')) ||
                            (transientWarningPopup.colKey === 'ng' && h.toUpperCase().includes('NG')) ||
                            (transientWarningPopup.colKey === 'itt' && h.toUpperCase().includes('ITT')) ||
                            (transientWarningPopup.colKey === 'torq' && (h.toUpperCase().includes('TORQ') || h.toUpperCase().includes('TQ')));

                          return (
                            <th
                              key={hIdx}
                              className={`px-3 py-1 text-center border-r border-slate-300 whitespace-nowrap ${
                                isTargetCol ? 'bg-[#facc15] text-slate-950 font-black' : 'bg-slate-200 text-slate-700 font-bold'
                              }`}
                            >
                              {h}
                            </th>
                          );
                        })}
                      </tr>
                    </thead>
                    <tbody>
                      <tr className="bg-amber-100/90 text-slate-950 font-bold animate-pulse">
                        <td className="px-2 py-1.5 text-center border-r border-slate-300 font-black">
                          #{transientWarningPopup.row.originalIndex}
                        </td>
                        <td className="px-2 py-1.5 text-center border-r border-slate-300 font-black text-amber-900">
                          {transientWarningPopup.row.timeStr}
                        </td>
                        {transientWarningPopup.row.cells.map((c, cIdx) => {
                          const h = analysisResult.headers[cIdx] || '';
                          const isTargetCol = h.toUpperCase().includes(transientWarningPopup.colKey.toUpperCase()) ||
                            (transientWarningPopup.colKey === 'np' && h.toUpperCase().includes('NP')) ||
                            (transientWarningPopup.colKey === 'ng' && h.toUpperCase().includes('NG')) ||
                            (transientWarningPopup.colKey === 'itt' && h.toUpperCase().includes('ITT')) ||
                            (transientWarningPopup.colKey === 'torq' && (h.toUpperCase().includes('TORQ') || h.toUpperCase().includes('TQ')));

                          return (
                            <td
                              key={cIdx}
                              className={`px-3 py-1.5 text-center border-r border-slate-300 ${
                                isTargetCol ? 'bg-[#fef08a] text-slate-950 font-black ring-2 ring-amber-500' : ''
                              }`}
                            >
                              {c}
                            </td>
                          );
                        })}
                      </tr>
                    </tbody>
                  </table>
                </div>
              )}

              {/* Bir Daha Gösterme Checkbox */}
              <div className="flex items-center gap-2 pt-2">
                <input
                  type="checkbox"
                  id="dontShowTransient"
                  checked={!!dontShowTransientAgain[transientWarningPopup.colKey]}
                  onChange={e => {
                    const checked = e.target.checked;
                    setDontShowTransientAgain(prev => ({
                      ...prev,
                      [transientWarningPopup.colKey]: checked
                    }));
                  }}
                  className="w-4 h-4 rounded text-amber-600 focus:ring-amber-500 cursor-pointer"
                />
                <label htmlFor="dontShowTransient" className="text-xs font-bold text-slate-700 cursor-pointer select-none">
                  Bu parametre için anlık tepe uyarı penceresini bir daha gösterme
                </label>
              </div>
            </div>

            <div className="p-4 bg-slate-50 border-t border-slate-200 flex items-center justify-end">
              <button
                type="button"
                onClick={() => setTransientWarningPopup(null)}
                className="px-6 py-2 rounded-xl bg-amber-500 hover:bg-amber-600 text-slate-950 font-black text-xs uppercase tracking-wider shadow-md cursor-pointer transition-all active:scale-95"
              >
                Tamam
              </button>
            </div>
          </div>
        </div>
      )}

      {/* SARI İKAZ KAYITLARI & GEÇMİŞ DETAYI MODALI */}
      {isYellowNgHistoryModalOpen && (
        <div className="fixed inset-0 z-[11000] flex items-center justify-center p-3 sm:p-6 bg-slate-950/80 backdrop-blur-md animate-fade-in">
          <div className="bg-white rounded-3xl shadow-2xl border border-slate-200 w-full max-w-5xl max-h-[88vh] flex flex-col overflow-hidden animate-scale-up">
            <div className="px-6 py-4 bg-gradient-to-r from-amber-600 via-amber-700 to-amber-900 text-white flex items-center justify-between shrink-0 shadow-sm">
              <div className="flex items-center gap-3">
                <div className="w-9 h-9 rounded-xl bg-white/20 flex items-center justify-center text-white">
                  <Flame className="w-5 h-5 text-amber-200" />
                </div>
                <div>
                  <h3 className="text-base font-black tracking-tight text-white flex items-center gap-2">
                    <span>{selectedTail.fullLabel}</span>
                    <span className="text-xs bg-amber-500/80 px-2 py-0.5 rounded font-mono font-bold">
                      Motor S/N: {selectedTail.motorSn}
                    </span>
                  </h3>
                  <div className="flex items-center gap-2 mt-0.5 flex-wrap">
                    <p className="text-xs text-amber-100 font-mono font-bold">
                      Kayıtlı Ng Sarı İkaz Satırları (Mevcut Adet: {(yellowNgRecords.filter(r => r.tailCode === selectedTail.code)).length} Adet)
                    </p>
                    <span className="inline-flex items-center gap-1.5 text-[10px] bg-emerald-950/60 text-emerald-300 px-2.5 py-0.5 rounded-full border border-emerald-400/40 font-mono font-semibold">
                      <span className="w-1.5 h-1.5 rounded-full bg-emerald-400 animate-pulse" />
                      Canlı Google E-Tablo ("Sarı İkaz Ng Takip Adedi")
                    </span>
                  </div>
                </div>
              </div>
              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={() => setIsYellowNgHistoryModalOpen(false)}
                  className="w-8 h-8 rounded-xl bg-white/10 hover:bg-white/20 text-white flex items-center justify-center cursor-pointer transition-colors"
                >
                  <X className="w-5 h-5" />
                </button>
              </div>
            </div>

            <div className="p-4 overflow-y-auto flex-1 font-mono text-xs">
              {isLoadingYellowNg && (
                <div className="p-3 mb-3 bg-amber-50 text-amber-900 border border-amber-300 rounded-xl flex items-center gap-2 shadow-xs animate-pulse">
                  <RefreshCw className="w-4 h-4 animate-spin text-amber-700 shrink-0" />
                  <span className="font-bold">Google E-Tablodan canlı sarı ikaz satırları güncelleniyor, lütfen bekleyiniz...</span>
                </div>
              )}

              {(() => {
                const tailRecords = yellowNgRecords.filter(r => r.tailCode === selectedTail.code);
                if (tailRecords.length === 0) {
                  return (
                    <div className="text-center py-12 text-slate-500">
                      <AlertTriangle className="w-10 h-10 text-amber-400 mx-auto mb-2 opacity-60" />
                      <p className="font-bold text-sm text-slate-700">Bu hava aracı ({selectedTail.code}) için henüz kayıtlı sarı ikaz satırı bulunmuyor.</p>
                      <p className="text-xs text-slate-400 mt-1 max-w-md mx-auto">
                        Uçuş kaydı yüklendiğinde "Sarı İkazları Saydır" butonuna basarak dosyadaki sarı ikaz adedi kadar satır açabilir ve adedi otomatik işletebilirsiniz.
                      </p>
                    </div>
                  );
                }

                return (
                  <div className="space-y-3">
                    <div className="flex items-center justify-between text-xs text-slate-600 bg-amber-50/70 p-2.5 rounded-xl border border-amber-200 flex-wrap gap-2">
                      <span className="font-bold text-amber-950 flex items-center gap-2">
                        <span>{selectedTail.code} Kuyruk Nolu Uçağın Kayıtlı Satır Adedi: <strong>{tailRecords.length} Adet</strong></span>
                        <span className="text-[10px] bg-emerald-100 text-emerald-800 px-2 py-0.5 rounded font-mono font-bold border border-emerald-200">
                          Google E-Tablo ile Canlı Eşit
                        </span>
                      </span>
                      <span className="text-[11px] text-slate-500">
                        E-tabloda kaç satır varsa o kadar adet sayılır.
                      </span>
                    </div>

                    <div className="overflow-x-auto border border-slate-200 rounded-2xl shadow-2xs">
                      <table className="w-full text-left border-collapse">
                        <thead>
                          <tr className="bg-amber-100/70 text-slate-900 border-b border-amber-200 text-[11px] font-black uppercase">
                            <th className="px-3 py-2 text-center w-12">#</th>
                            <th className="px-3 py-2">Kuyruk No</th>
                            <th className="px-3 py-2">Motor S/N</th>
                            <th className="px-3 py-2">Olay Numarası</th>
                            <th className="px-3 py-2">Zaman Aralığı (TIME)</th>
                            <th className="px-3 py-2 text-center">Olay Tepesi</th>
                            <th className="px-3 py-2">Satır Aralığı</th>
                            <th className="px-3 py-2 text-center">Süre</th>
                            <th className="px-3 py-2">Sisteme Kayıt Tarihi</th>
                            <th className="px-3 py-2 text-center">İşlem</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-100 text-[11px]">
                          {tailRecords.map((r, idx) => (
                            <tr key={r.id || idx} className="hover:bg-amber-50/50 transition-colors">
                              <td className="px-3 py-2 text-center font-bold text-slate-400">
                                {tailRecords.length - idx}
                              </td>
                              <td className="px-3 py-2 font-black text-slate-800">
                                {r.tailCode || selectedTail.code}
                              </td>
                              <td className="px-3 py-2 text-emerald-800 font-bold">
                                {r.motorSn || selectedTail.motorSn}
                              </td>
                              <td className="px-3 py-2 font-bold text-slate-700">
                                {r.olayNo || `#${idx + 1}`}
                              </td>
                              <td className="px-3 py-2 text-amber-900 font-bold">
                                {r.zamanAraligi || '-'}
                              </td>
                              <td className="px-3 py-2 text-center font-black text-amber-600 bg-amber-50/50 rounded">
                                {r.olayTepesi || '-'}
                              </td>
                              <td className="px-3 py-2 text-slate-600">
                                {r.satirAraligi || '-'}
                              </td>
                              <td className="px-3 py-2 text-center font-bold text-slate-700">
                                {r.sure || '-'}
                              </td>
                              <td className="px-3 py-2 text-slate-500 whitespace-nowrap">
                                {r.sistemKayitTarihi || new Date(r.createdAt || Date.now()).toLocaleString('tr-TR')}
                              </td>
                              <td className="px-3 py-2 text-center">
                                <button
                                  type="button"
                                  onClick={() => handleDeleteYellowNgRecord(r.id)}
                                  className="p-1.5 rounded-lg text-slate-400 hover:text-red-600 hover:bg-red-50 cursor-pointer transition-colors"
                                  title="Bu sarı ikaz satırını sil (Google E-Tablodan silinir ve adet otomatik 1 düşer)"
                                >
                                  <Trash2 className="w-3.5 h-3.5" />
                                </button>
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                );
              })()}
            </div>

            <div className="p-4 bg-slate-50 border-t border-slate-200 flex flex-wrap items-center justify-between gap-3">
              <div className="flex items-center gap-2 text-xs text-slate-600 font-mono">
                <Database className="w-4 h-4 text-emerald-600 shrink-0" />
                <span>
                  Veri Kaynağı: <strong>Google E-Tablo</strong> ("Sarı İkaz Ng Takip Adedi") • JSON veya statik depolama kullanılmaz.
                  {lastYellowNgFetchTime && ` • Son E-Tablo Senkronizasyonu: ${lastYellowNgFetchTime}`}
                </span>
              </div>
              <button
                type="button"
                onClick={() => setIsYellowNgHistoryModalOpen(false)}
                className="px-5 py-2 rounded-xl bg-slate-800 hover:bg-slate-900 text-white font-bold text-xs cursor-pointer transition-all active:scale-95"
              >
                Kapat
              </button>
            </div>
          </div>
        </div>
      )}

      {/* AT-802F Motor Uçuş Grafiği & Oynatıcı Modalı */}
      {isFlightChartModalOpen && (
        <FlightEngineChartModal
          isOpen={isFlightChartModalOpen}
          onClose={() => setIsFlightChartModalOpen(false)}
          rows={analysisResult?.rows || []}
          headers={analysisResult?.headers || []}
          episodes={analysisResult?.episodes || []}
          initialParamKey={chartInitialParamKey}
          tailLabel={selectedTail.fullLabel}
          motorSn={analysisResult?.motorSn || selectedTail.motorSn || 'PT6A-67AG'}
        />
      )}

      {/* ─── ŞİFRE DOĞRULAMA MODALI (ŞİFRE: 802) ─── */}
      {isPasswordModalOpen && (
        <div className="fixed inset-0 z-[10020] bg-black/60 backdrop-blur-xs flex items-center justify-center p-4 animate-in fade-in">
          <div className="w-full max-w-md bg-white rounded-3xl shadow-2xl border border-slate-200 overflow-hidden animate-in zoom-in-95">
            <div className="px-6 py-4 bg-gradient-to-r from-amber-600 to-amber-700 text-white flex items-center justify-between">
              <div className="flex items-center gap-2.5">
                <KeyRound className="w-5 h-5 text-amber-200" />
                <h4 className="text-sm font-black uppercase tracking-wider">Yetkili Şifresi Doğrulama</h4>
              </div>
              <button
                type="button"
                onClick={() => {
                  setIsPasswordModalOpen(false);
                  setPasswordInput('');
                  setPasswordError(null);
                }}
                className="text-white/80 hover:text-white p-1 rounded-lg hover:bg-white/10 cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="p-6 space-y-4">
              <p className="text-xs text-slate-600 leading-relaxed font-sans">
                <strong>Sarı İkaz Ng Takip Adedi</strong> Google E-Tablo aktarımı yetkili şifresi gerektirmektedir. Lütfen devam etmek için yetkili şifresini giriniz:
              </p>

              <form
                onSubmit={(e) => {
                  e.preventDefault();
                  handleVerifyPasswordAndTransfer();
                }}
                className="space-y-3"
              >
                <div>
                  <label className="block text-[11px] font-black text-slate-700 uppercase mb-1">
                    Yetki Şifresi
                  </label>
                  <input
                    type="password"
                    autoFocus
                    value={passwordInput}
                    onChange={(e) => {
                      setPasswordInput(e.target.value);
                      setPasswordError(null);
                    }}
                    placeholder="Şifre giriniz"
                    className="w-full px-4 py-2.5 rounded-xl border-2 border-slate-300 text-sm font-mono tracking-widest focus:border-amber-600 focus:outline-none transition-colors"
                  />
                  {passwordError && (
                    <p className="text-xs font-bold text-rose-600 mt-1.5 flex items-center gap-1">
                      <AlertTriangle className="w-3.5 h-3.5 shrink-0" />
                      <span>{passwordError}</span>
                    </p>
                  )}
                </div>

                <div className="flex items-center justify-end gap-2 pt-2">
                  <button
                    type="button"
                    onClick={() => {
                      setIsPasswordModalOpen(false);
                      setPasswordInput('');
                      setPasswordError(null);
                    }}
                    className="px-4 py-2 rounded-xl text-xs font-bold text-slate-600 hover:bg-slate-100 cursor-pointer"
                  >
                    İptal
                  </button>
                  <button
                    type="submit"
                    className="px-5 py-2 rounded-xl bg-amber-600 hover:bg-amber-700 text-white font-black text-xs uppercase tracking-wider shadow-sm cursor-pointer active:scale-95 transition-all"
                  >
                    Doğrula ve Aktar
                  </button>
                </div>
              </form>
            </div>
          </div>
        </div>
      )}

      {/* ─── MÜKERRER KAYIT UYARISI MODALI ─── */}
      {isDuplicateWarningOpen && (
        <div className="fixed inset-0 z-[10030] bg-black/70 backdrop-blur-xs flex items-center justify-center p-4 animate-in fade-in">
          <div className="w-full max-w-lg bg-white rounded-3xl shadow-2xl border-2 border-rose-400 overflow-hidden animate-in zoom-in-95">
            <div className="px-6 py-4 bg-gradient-to-r from-rose-600 to-rose-700 text-white flex items-center justify-between">
              <div className="flex items-center gap-2.5">
                <AlertTriangle className="w-5 h-5 text-amber-300" />
                <h4 className="text-sm font-black uppercase tracking-wider">Mükerrer Kayıt Uyarısı</h4>
              </div>
              <button
                type="button"
                onClick={() => setIsDuplicateWarningOpen(false)}
                className="text-white/80 hover:text-white p-1 rounded-lg hover:bg-white/10 cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="p-6 space-y-4">
              <div className="p-4 bg-rose-50 rounded-2xl border border-rose-200 flex items-start gap-3">
                <div className="w-9 h-9 rounded-xl bg-rose-100 text-rose-700 flex items-center justify-center shrink-0 mt-0.5">
                  <AlertTriangle className="w-5 h-5" />
                </div>
                <div className="text-xs text-rose-950 space-y-1">
                  <p className="font-black text-sm">{selectedTail.code} için bu uçuş dosyası daha önce aktarılmıştır!</p>
                  <p className="leading-relaxed">
                    <strong>{selectedTail.fullLabel} ({selectedTail.code})</strong> hava aracı için{' '}
                    <strong className="font-mono text-rose-900 bg-rose-100 px-1.5 py-0.5 rounded">
                      "{duplicateWarningFileName}"
                    </strong>{' '}
                    adlı uçuş dosyası daha önce sisteme ve <strong>"Sarı İkaz Ng Takip Adedi"</strong> Google E-Tablo sayfasına aktarılmıştır.
                  </p>
                </div>
              </div>

              <p className="text-xs text-slate-600 leading-relaxed font-sans">
                Sistemde aynı hava aracı için mükerrer kayıt oluşmasını engellemek amacıyla aktarım durdurulmuştur. (Farklı hava araçları için aynı dosya adı aktarılabilir; ancak aynı hava aracı için mükerrer aktarım engellenir.)
              </p>

              <div className="flex items-center justify-end pt-2">
                <button
                  type="button"
                  onClick={() => setIsDuplicateWarningOpen(false)}
                  className="px-6 py-2.5 rounded-xl bg-rose-600 hover:bg-rose-700 text-white font-black text-xs uppercase tracking-wider shadow-md cursor-pointer active:scale-95 transition-all"
                >
                  Anladım, Kapat
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
