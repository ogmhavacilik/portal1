import React, { useState, useEffect, useMemo, useRef } from 'react';
import {
  X,
  FileSpreadsheet,
  Plus,
  Trash2,
  Edit3,
  Search,
  CheckCircle2,
  AlertTriangle,
  RotateCw,
  Lock,
  FileText,
  Database,
  ArrowLeft,
  UploadCloud,
  Check,
  Paperclip,
  Download,
  Eye,
  ArrowUpDown,
  ArrowUp,
  ArrowDown,
  RefreshCw,
  Image as ImageIcon,
  Archive,
  Link as LinkIcon,
  ExternalLink
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { unmergeAndFillWorksheet, cleanAndFormatDateString } from '../utils/driveExcelSync';
import { safeGetJSON, safeSetJSON } from '../utils/safeStorage';
import {
  getAllHangarPdfDocs,
  saveHangarPdfDoc,
  deleteHangarPdfDoc,
  syncHangarPdfDocsFromDrive,
  HangarPdfDoc,
  matchOlayTakipDoc,
  setRowDocLink,
  getFileCategory,
  getDocMimeType
} from '../utils/hangarPdfStorage';
import { PdfPreviewModal } from './PdfPreviewModal';

export const GOOGLE_SCRIPT_URL =
  "https://script.google.com/macros/s/AKfycbwP1uOo2NrST5a4I8vm1nGBLtI26yY2lWrmu9_e9iymwBkUhJBA9JOPCp7SNKqJbOubOw/exec";
export const DRIVE_FOLDER_ID = "1HQR_NYKhHQGA7_2W3nArI9pCh-LJasTP";

export interface OlayTakipSheet {
  id: string;
  sheetName: string;
  title: string;
  columns: string[];
  rows: string[][];
}

export const EK_DOSYA_LINK_COL_NAME = 'EK DOSYA BAĞLANTI LİNKİ';

/**
 * Checks if a column name represents the Ek Dosya Bağlantı Linki column
 */
export const isLinkColumnName = (colName: string): boolean => {
  if (!colName) return false;
  return /ek\s*dosya\s*ba[gğ]lant[ıi]\s*l[ıi]nk[ıi]/i.test(colName) || 
         /ba[gğ]lant[ıi]\s*l[ıi]nk/i.test(colName);
};

/**
 * Ensures a sheet contains the 'EK DOSYA BAĞLANTI LİNKİ' column, inserting it right after
 * 'Belge / Döküman' or at the end if missing.
 */
export const ensureEkDosyaLinkColumn = (sheet: OlayTakipSheet): { sheet: OlayTakipSheet; colIndex: number } => {
  if (!sheet || !sheet.columns) return { sheet, colIndex: -1 };
  
  const existingIdx = sheet.columns.findIndex(c => isLinkColumnName(c));
  if (existingIdx >= 0) {
    return { sheet, colIndex: existingIdx };
  }

  // Find doc column index to place link column directly adjacent
  const docColIdx = sheet.columns.findIndex(c => /belge|d[oö]k[uü]man/i.test(c));
  const insertIdx = docColIdx >= 0 ? docColIdx + 1 : sheet.columns.length;

  const newCols = [...sheet.columns];
  newCols.splice(insertIdx, 0, EK_DOSYA_LINK_COL_NAME);

  const newRows = (sheet.rows || []).map(r => {
    const rowCopy = [...r];
    while (rowCopy.length < insertIdx) rowCopy.push('');
    rowCopy.splice(insertIdx, 0, '');
    return rowCopy;
  });

  return {
    sheet: {
      ...sheet,
      columns: newCols,
      rows: newRows
    },
    colIndex: insertIdx
  };
};

/**
 * Ensures all sheets across all units have the EK DOSYA BAĞLANTI LİNKİ column
 */
export const ensureLinkColumnInAllUnitsData = (data: Record<string, OlayTakipSheet[]>): Record<string, OlayTakipSheet[]> => {
  if (!data || typeof data !== 'object') return data;
  const result: Record<string, OlayTakipSheet[]> = {};
  for (const [unit, sheets] of Object.entries(data)) {
    if (!Array.isArray(sheets)) {
      result[unit] = sheets;
      continue;
    }
    result[unit] = sheets.map(sheet => ensureEkDosyaLinkColumn(sheet).sheet);
  }
  return result;
};

export interface UnitOlayTakipData {
  unitKey: string;
  unitLabel: string;
  sheets: OlayTakipSheet[];
}

export const getOlayTakipDriveFileName = (unitKey: string): string => {
  const map: Record<string, string> = {
    'at802': 'olay_takip_cizelgesi_at-802.xlsx',
    'bell429': 'olay_takip_cizelgesi_bell-429.xlsx',
    't70': 'olay_takip_cizelgesi_t-70.xlsx',
    't70_bumbi': 'olay_takip_cizelgesi_t-70_bumbi_backet.xlsx',
    't70_helitak': 'olay_takip_cizelgesi_t-70_helitak.xlsx',
    'c650': 'olay_takip_cizelgesi_c-650.xlsx',
    'b360': 'olay_takip_cizelgesi_b-360.xlsx',
    'hangar': 'olay_takip_cizelgesi_hangar.xlsx',
  };
  return map[unitKey] || `olay_takip_cizelgesi_${unitKey}.xlsx`;
};

// Default empty templates with Document and Link columns
const DEFAULT_AT802_SHEETS: OlayTakipSheet[] = [
  {
    id: 'sheet_at802_kronoloji',
    sheetName: 'Olay ve Kaza-Kırım Kronolojisi',
    title: 'AT-802 HAVA ARAÇLARI ÖNEMLİ OLAY VE KAZA/KIRIM KRONOLOJİSİ',
    columns: [
      'Sıra',
      'Tarih',
      'Kuyruk No',
      'Olay / Kaza-Kırım Tipi',
      'Konum / Bölge',
      'Sorumlu Pilot',
      'Olay Oluş Şekli ve Hasar Açıklaması',
      'Tahliye / Lojistik / Bakım Aksiyonu',
      'Hava Aracı / Parça Durumu',
      'Belge / Döküman',
      'EK DOSYA BAĞLANTI LİNKİ'
    ],
    rows: []
  },
  {
    id: 'sheet_at802_ng_limit',
    sheetName: 'Ng Limit Aşımları Detayı',
    title: 'MOTOR NG LİMİT AŞIMLARI VE DAA SÜRECİ',
    columns: [
      'Sıra',
      'Tarih / Dönem',
      'Kuyruk No',
      'Motor Seri No',
      'Olay Türü / İşlem',
      'Ölçülen Tepe Ng (%)',
      'P&WC DAA Referans No',
      'Verilen Şart / Kalan Saati',
      'Sorumlu Pilot',
      'Teknik Değerlendirme ve Sonuç',
      'Hava Aracı / Parça Durumu',
      'Belge / Döküman',
      'EK DOSYA BAĞLANTI LİNKİ'
    ],
    rows: []
  },
  {
    id: 'sheet_at802_kaza_kirim',
    sheetName: 'Kaza-Kırım Yapan Uçaklar',
    title: 'AT-802 HAVA ARAÇLARI KAZA / KIRIM ÖZET TABLOSU',
    columns: [
      'Sıra',
      'Tarih',
      'Kuyruk No',
      'Uçak Tipi / Model',
      'Kaza / Kırım Konumu',
      'Görev / Faaliyet Tipi',
      'Kaza / Kırım Oluş Şekli ve Hasar Kapsamı',
      'Tahliye / Lojistik Aksiyonu',
      'Hava Aracı / Parça Durumu',
      'Belge / Döküman',
      'EK DOSYA BAĞLANTI LİNKİ'
    ],
    rows: []
  },
  {
    id: 'sheet_at802_motor_ozet',
    sheetName: 'Motor Durum Özet',
    title: 'AT-802 FİLOSU MOTOR DURUM VE NG LİMİT AŞIMI ÖZETİ',
    columns: [
      'Sıra',
      'Kuyruk No',
      'Motor Seri No (PCE)',
      'Tespit Edilen En Yüksek Ng (%)',
      'Ng Aşımı Sayısı / Tarihi',
      'Hava Aracı / Parça Durumu',
      'Kalan Uçuş Saati (DAA)',
      'Planlanan Bakım / Onarım Aksiyonu',
      'Belge / Döküman',
      'EK DOSYA BAĞLANTI LİNKİ'
    ],
    rows: []
  }
];

export const UNITS_CONFIG = [
  { key: 'at802', label: 'AT-802', subtitle: 'Air Tractor Yangın Söndürme', password: '802' },
  { key: 'bell429', label: 'BELL 429', subtitle: 'Bell 429 Keşif & Yangın', password: '429' },
  { key: 't70', label: 'T-70', subtitle: 'T-70 Genel Maksat Yangın', password: '70' },
  { key: 't70_bumbi', label: 'T-70 BUMBİ BUCKET', subtitle: 'Bumbi Bucket Donanım Takip', password: '70' },
  { key: 't70_helitak', label: 'T-70 HELİTAK', subtitle: 'Helitak Yangın Donanım', password: '70' },
  { key: 'c650', label: 'C-650', subtitle: 'Cessna Citation Sovereign', password: '650' },
  { key: 'b360', label: 'B-360', subtitle: 'Beechcraft King Air 360', password: '360' },
  { key: 'hangar', label: 'HANGAR & DİĞER', subtitle: 'Hangar Genel Olay Çizelgesi', password: '1839' }
];

const STORAGE_KEY = 'olay_takip_cizelgesi_store_v3';

/**
 * Akıllı Başlık Satırı Tespiti:
 * Birleştirilmiş hücrelerin açılması sonucu oluşan başlık banner'ları (tekil tekrar eden değerler)
 * veya veri satırları (tarih, uzun kaza cümleleri) elenerek gerçek kolon başlık satırı tespit edilir.
 */
export const detectOlayTakipHeaderRow = (rawRows: any[][]): number => {
  let bestRowIdx = 0;
  let bestScore = -1;

  const headerKeywords = [
    'SIRA', 'NO', 'S.NO', 'S.NU', 'TARİH', 'TARIH', 'KUYRUK', 'UÇAK', 'UCAK', 'MODEL', 'TİP', 'TIP',
    'MOTOR', 'SERİ', 'SERI', 'S/N', 'NG', 'LİMİT', 'LIMIT', 'TEPE', 'DAA', 'AŞIMI', 'ASIMI',
    'OLAY', 'KAZA', 'KIRIM', 'KONUM', 'BÖLGE', 'BOLGE', 'PİLOT', 'PILOT', 'SORUMLU',
    'HASAR', 'AÇIKLAMA', 'ACIKLAMA', 'TAHLİYE', 'TAHLIYE', 'LOJİSTİK', 'LOJISTIK', 'AKSİYON', 'AKSIYON',
    'DURUM', 'DURUMU', 'SONUÇ', 'SONUC', 'KARAR', 'SAAT', 'SAATİ', 'DÖNEM', 'DONEM', 'İŞLEM', 'ISLEM',
    'ŞART', 'SART', 'GÖREV', 'GOREV', 'FAALİYET', 'FAALIYET', 'BELGE', 'DÖKÜMAN', 'DOKUMAN'
  ];

  for (let r = 0; r < Math.min(rawRows.length, 12); r++) {
    const row = rawRows[r];
    if (!row || !Array.isArray(row)) continue;

    const cells = row.map(c => String(c || '').trim());
    const nonEmpties = cells.filter(c => c !== '');
    if (nonEmpties.length === 0) continue;

    // Check distinct non-empty count: eğer tüm hücreler aynı metinse (örneğin unmerge olmuş başlık banner'ı) başlık değildir
    const distinctSet = new Set(nonEmpties.map(c => c.toUpperCase()));
    if (distinctSet.size <= 1 && nonEmpties.length > 1) {
      continue;
    }

    let score = 0;
    for (const cell of nonEmpties) {
      const upper = cell.toUpperCase();

      // Çok uzun açıklama cümleleri veri satırıdır, başlık değil
      if (cell.length > 75) {
        score -= 15;
      } else {
        score += 3;
      }

      // Tarih veya seri no görünümü varsa ceza puanı (veri satırıdır)
      if (/^\d{1,2}[\.\/-]\d{1,2}[\.\/-]\d{2,4}$/.test(cell) || /^\d{5}$/.test(cell)) {
        score -= 25;
      }

      for (const kw of headerKeywords) {
        if (upper.includes(kw)) {
          score += 15;
          if (upper === kw || upper === `${kw} NO` || upper === `SIRA ${kw}` || upper === `${kw} ADI`) {
            score += 10;
          }
        }
      }
    }

    if (score > bestScore && score >= 20) {
      bestScore = score;
      bestRowIdx = r;
    }
  }

  return bestRowIdx;
};

/**
 * Dynamically parse any Excel workbook into OlayTakipSheet[] array.
 * Adapts to any sheet count, sheet names, and column headers.
 */
export const parseOlayTakipWorkbook = (wb: XLSX.WorkBook): OlayTakipSheet[] => {
  const result: OlayTakipSheet[] = [];
  for (let sIdx = 0; sIdx < wb.SheetNames.length; sIdx++) {
    const sheetName = wb.SheetNames[sIdx];
    const ws = wb.Sheets[sheetName];
    if (!ws) continue;
    unmergeAndFillWorksheet(ws);
    const rawRows: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    if (!rawRows || rawRows.length === 0) continue;

    const headerRowIdx = detectOlayTakipHeaderRow(rawRows);

    const headerRow = rawRows[headerRowIdx] || [];
    let lastFilled = headerRow.length - 1;
    while (lastFilled >= 0 && !String(headerRow[lastFilled] || '').trim()) {
      lastFilled--;
    }
    const columns: string[] = [];
    for (let c = 0; c <= Math.max(lastFilled, 0); c++) {
      const colName = String(headerRow[c] || '').trim();
      columns.push(colName || `Sütun ${c + 1}`);
    }

    // Her sayfaya kesinlikle 'Belge / Döküman' ve 'EK DOSYA BAĞLANTI LİNKİ' sütunlarını ekle
    const hasDocCol = columns.some(c => /belge|d[oö]k[uü]man/i.test(c));
    if (!hasDocCol) {
      columns.push('Belge / Döküman');
    }

    const hasLinkCol = columns.some(c => isLinkColumnName(c));
    if (!hasLinkCol) {
      columns.push(EK_DOSYA_LINK_COL_NAME);
    }

    // Extract data rows
    const rows: string[][] = [];
    for (let r = headerRowIdx + 1; r < rawRows.length; r++) {
      const row = rawRows[r];
      if (!row) continue;
      const hasAny = row.some((c: any) => String(c || '').trim() !== '');
      if (!hasAny) continue;

      // Skip row if it repeats header columns
      const isHeaderRepeat = row
        .filter((c: any) => String(c || '').trim() !== '')
        .every((c: any) => columns.some(col => col.toLowerCase() === String(c).trim().toLowerCase()));
      if (isHeaderRepeat) continue;

      const cleanRow: string[] = [];
      for (let c = 0; c < columns.length; c++) {
        const rawCell = row[c];
        const colName = columns[c] || '';
        let cellStr = rawCell !== undefined && rawCell !== null ? String(rawCell).trim() : '';
        const isDateCol = /tar[ıi]h|date/i.test(colName);
        if (isDateCol || (/^\d{5}(\.\d+)?$/.test(cellStr) && Number(cellStr) >= 20000 && Number(cellStr) <= 90000 && isDateCol)) {
          cellStr = cleanAndFormatDateString(rawCell);
        }
        cleanRow.push(cellStr);
      }
      rows.push(cleanRow);
    }

    result.push({
      id: `sheet_${sheetName.toLowerCase().replace(/[^a-z0-9]/g, '_')}_${sIdx}`,
      sheetName: sheetName,
      title: sheetName.toUpperCase(),
      columns: columns.length > 0 ? columns : ['Sıra', 'Tarih', 'Kuyruk No', 'Olay Detayı', 'Hava Aracı / Parça Durumu', 'Belge / Döküman'],
      rows: rows
    });
  }
  return result;
};

/**
 * Ng Limit sayısal değerini metinden çıkaran yardımcı fonksiyon.
 * Örn: "%107,1" -> 107.1, "%106,0 (Son: %105,7)" -> 106.0
 */
export const extractNgValue = (str: string): number | null => {
  if (!str) return null;
  const cleaned = str.replace('%', '').trim();
  const match = cleaned.match(/(\d{2,3}(?:[.,]\d+)?)/);
  if (!match) return null;
  const num = parseFloat(match[1].replace(',', '.'));
  if (isNaN(num) || num < 80 || num > 160) return null;
  return num;
};

/**
 * Ng Limit Hücresi Renk Skalası:
 * "en yüksek koyu en azı da az kırmızı şekilde sıralı olsun ama neyi 105 çok geçen daha koyu az geçen de koyu ama arada fark olsun"
 */
export const getNgLimitBadgeStyle = (cellValue: string): { bgClass: string; isNg: boolean } => {
  const ngVal = extractNgValue(cellValue);
  if (ngVal === null) return { bgClass: '', isNg: false };

  if (ngVal >= 107.0) {
    // 107.0%+ Çok yüksek aşım -> En koyu bordo/kırmızı
    return {
      bgClass: 'bg-red-950 text-rose-100 border border-red-900 font-black shadow-xs',
      isNg: true
    };
  } else if (ngVal >= 106.5) {
    // 106.5 - 106.9% -> Koyu kırmızı
    return {
      bgClass: 'bg-red-800 text-white border border-red-900 font-extrabold shadow-xs',
      isNg: true
    };
  } else if (ngVal >= 106.0) {
    // 106.0 - 106.4% -> Belirgin orta-koyu kırmızı
    return {
      bgClass: 'bg-red-700 text-white border border-red-800 font-bold shadow-xs',
      isNg: true
    };
  } else if (ngVal >= 105.5) {
    // 105.5 - 105.9% -> Orta kırmızı
    return {
      bgClass: 'bg-rose-600 text-white border border-rose-700 font-bold shadow-xs',
      isNg: true
    };
  } else if (ngVal > 105.0) {
    // 105.1 - 105.4% -> Az geçen kırmızı (ama net kırmızı, fark edilebilir)
    return {
      bgClass: 'bg-rose-500 text-white border border-rose-600 font-semibold shadow-xs',
      isNg: true
    };
  }

  // 105 ve altı: Normal / Güvenli
  return {
    bgClass: 'bg-emerald-50 text-emerald-800 border border-emerald-200 font-mono font-medium',
    isNg: true
  };
};

/**
 * Hava Aracı / Parça Durumu Dolgu Renkleri:
 * "faaal ise yeşil onarıma alında dolgusu sarı Bakım/Overhaul Yapıldı turuncu onarılda faal yeşil ağoır hasarlı kırmızı dolgular olsun"
 */
export const getStatusBadgeStyle = (statusStr: string): { bgClass: string; isStatus: boolean } => {
  const s = (statusStr || '').trim().toLowerCase();
  if (!s || s === '-') return { bgClass: '', isStatus: false };

  // Ağır Hasarlı / Kaza-Kırım / Gayrifaal -> Kırmızı dolgu
  if (
    s.includes('ağır') ||
    s.includes('agir') ||
    s.includes('hasarlı') ||
    s.includes('hasarli') ||
    s.includes('gayrifaal') ||
    s.includes('kaza-kırım') ||
    s.includes('kaza / kırım') ||
    s.includes('kırım')
  ) {
    return {
      bgClass: 'bg-rose-600 text-white border border-rose-700 font-black shadow-xs',
      isStatus: true
    };
  }

  // Bakım/Overhaul Yapıldı / Overhaul / Bakımda -> Turuncu dolgu
  if (
    s.includes('overhaul') ||
    s.includes('bakım yapıldı') ||
    s.includes('bakim yapildi') ||
    s.includes('bakım/overhaul') ||
    s.includes('bakim/overhaul') ||
    s.includes('bakımda') ||
    s.includes('bakimda')
  ) {
    return {
      bgClass: 'bg-orange-500 text-white border border-orange-600 font-bold shadow-xs',
      isStatus: true
    };
  }

  // Onarıma Alındı / Onarımda -> Sarı dolgu
  if (
    s.includes('onarıma alındı') ||
    s.includes('onarima alindi') ||
    s.includes('onarımda') ||
    s.includes('onarimda') ||
    s.includes('onarıma') ||
    s.includes('onarima')
  ) {
    return {
      bgClass: 'bg-yellow-400 text-slate-950 border border-yellow-500 font-black shadow-xs',
      isStatus: true
    };
  }

  // Faal / Onarıldı / Onarıldı Faal -> Yeşil dolgu
  if (s.includes('faal') || s.includes('onarıldı') || s.includes('onarildi')) {
    return {
      bgClass: 'bg-emerald-600 text-white border border-emerald-700 font-black shadow-xs',
      isStatus: true
    };
  }

  return {
    bgClass: 'bg-slate-100 text-slate-700 border border-slate-300 font-semibold',
    isStatus: true
  };
};

interface OlayTakipCizelgesiModalProps {
  isOpen: boolean;
  onClose: () => void;
  showNotification: (msg: string, type?: 'success' | 'error' | 'info') => void;
  initialUnit?: string;
}

export const OlayTakipCizelgesiModal: React.FC<OlayTakipCizelgesiModalProps> = ({
  isOpen,
  onClose,
  showNotification,
  initialUnit = 'at802'
}) => {
  const [selectedUnit, setSelectedUnit] = useState<string>(initialUnit);
  const [activeSheetIndex, setActiveSheetIndex] = useState<number>(0);
  const [allUnitsData, setAllUnitsData] = useState<Record<string, OlayTakipSheet[]>>({});
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [isSyncing, setIsSyncing] = useState<boolean>(false);
  const [lastSyncTime, setLastSyncTime] = useState<string>('-');

  // Sorting
  const [sortConfig, setSortConfig] = useState<{ colIndex: number; direction: 'asc' | 'desc' } | null>(null);

  // Row Edit / Add State
  const [editingRow, setEditingRow] = useState<{ rowIndex: number; row: string[] } | null>(null);
  const [isAddingRow, setIsAddingRow] = useState<boolean>(false);
  const [newRowValues, setNewRowValues] = useState<string[]>([]);

  // New Sheet Modal State
  const [isAddingSheet, setIsAddingSheet] = useState<boolean>(false);
  const [newSheetName, setNewSheetName] = useState<string>('');
  const [newSheetTitle, setNewSheetTitle] = useState<string>('');
  const [newSheetColumnsText, setNewSheetColumnsText] = useState<string>(
    'Sıra, Tarih, Kuyruk No, Olay Açıklaması, Sorumlu Personel, Hava Aracı / Parça Durumu, Belge / Döküman'
  );

  // Editable Sheet Name, Sheet Title, and Column Header states (Double-Click to edit)
  const [editingSheetIndex, setEditingSheetIndex] = useState<number | null>(null);
  const [editingSheetNameValue, setEditingSheetNameValue] = useState<string>('');
  
  const [isEditingSheetTitle, setIsEditingSheetTitle] = useState<boolean>(false);
  const [editingSheetTitleValue, setEditingSheetTitleValue] = useState<string>('');

  const [editingColIndex, setEditingColIndex] = useState<number | null>(null);
  const [editingColNameValue, setEditingColNameValue] = useState<string>('');

  // Search filter for Drive PDFs in document modal
  const [driveDocSearchTerm, setDriveDocSearchTerm] = useState<string>('');

  // Document Upload & Preview State
  const [hangarPdfDocs, setHangarPdfDocs] = useState<HangarPdfDoc[]>([]);
  const [previewDoc, setPreviewDoc] = useState<HangarPdfDoc | null>(null);
  const [activeDocRow, setActiveDocRow] = useState<{ rowIndex: number; row: string[] } | null>(null);
  const [docModalTab, setDocModalTab] = useState<'upload' | 'drive' | 'link'>('upload');
  const [customLinkUrl, setCustomLinkUrl] = useState<string>('');
  const [customLinkTitle, setCustomLinkTitle] = useState<string>('');
  const [docUploadFile, setDocUploadFile] = useState<File | null>(null);
  const [docUploadName, setDocUploadName] = useState<string>('');
  const [docUploadCategory, setDocUploadCategory] = useState<string>('Kaza/Kırım & Olay Belgesi');
  const [isDocUploading, setIsDocUploading] = useState<boolean>(false);
  const docFileInputRef = useRef<HTMLInputElement | null>(null);

  // Veri Güncelleme Modal State
  const [isDataSyncModalOpen, setIsDataSyncModalOpen] = useState<boolean>(false);
  const [dataSyncStep, setDataSyncStep] = useState<1 | 2>(1);
  const [syncPasswordInput, setSyncPasswordInput] = useState<string>('');
  const [syncPasswordError, setSyncPasswordError] = useState<string>('');
  const [syncProgressPercent, setSyncProgressPercent] = useState<number>(0);
  const [syncProgressStage, setSyncProgressStage] = useState<string>('');
  const syncFileInputRef = useRef<HTMLInputElement | null>(null);

  // General Password Modal for Row deletion
  const [passwordModalOpen, setPasswordModalOpen] = useState<boolean>(false);
  const [passwordInput, setPasswordInput] = useState<string>('');
  const [passwordError, setPasswordError] = useState<string>('');
  const [pendingAction, setPendingAction] = useState<(() => void) | null>(null);

  // State for syncing docs from Drive
  const [isSyncingDocs, setIsSyncingDocs] = useState<boolean>(false);

  // Load Hangar PDF documents from IndexedDB and sync from Google Drive
  useEffect(() => {
    getAllHangarPdfDocs()
      .then(docs => {
        if (Array.isArray(docs) && docs.length > 0) {
          setHangarPdfDocs(docs);
        }
      })
      .catch(err => console.warn('Olay takip docs load warning:', err));

    syncHangarPdfDocsFromDrive(GOOGLE_SCRIPT_URL)
      .then(driveDocs => {
        if (Array.isArray(driveDocs) && driveDocs.length > 0) {
          setHangarPdfDocs(driveDocs);
        }
      })
      .catch(err => console.warn('Olay takip docs drive sync warning:', err));
  }, [isOpen]);

  // Sync initial unit prop if changed
  useEffect(() => {
    if (initialUnit && UNITS_CONFIG.some(u => u.key === initialUnit)) {
      setSelectedUnit(initialUnit);
    }
  }, [initialUnit]);

  // Load from Storage or Defaults
  useEffect(() => {
    if (!isOpen) return;
    const saved = safeGetJSON<Record<string, OlayTakipSheet[]> | null>(STORAGE_KEY, null);
    if (saved && typeof saved === 'object' && Object.keys(saved).length > 0) {
      const migrated = ensureLinkColumnInAllUnitsData(saved);
      setAllUnitsData(migrated);
      safeSetJSON(STORAGE_KEY, migrated);
    } else {
      const initialData: Record<string, OlayTakipSheet[]> = {
        at802: DEFAULT_AT802_SHEETS
      };
      UNITS_CONFIG.forEach(u => {
        if (u.key !== 'at802') {
          initialData[u.key] = [
            {
              id: `sheet_${u.key}_genel`,
              sheetName: 'Olay ve Kaza-Kırım Çizelgesi',
              title: `${u.label} ÖNEMLİ OLAY VE KAZA/KIRIM ÇİZELGESİ`,
              columns: [
                'Sıra',
                'Tarih',
                'Kuyruk No',
                'Olay / Kaza-Kırım Tipi',
                'Konum / Bölge',
                'Sorumlu Personel',
                'Olay Detayı',
                'Yapılan İşlemler',
                'Hava Aracı / Parça Durumu',
                'Belge / Döküman',
                'EK DOSYA BAĞLANTI LİNKİ'
              ],
              rows: []
            }
          ];
        }
      });
      setAllUnitsData(initialData);
      safeSetJSON(STORAGE_KEY, initialData);
    }

    fetchExcelFromDrive(selectedUnit, false);
  }, [isOpen]);

  // When selected unit changes, reset sheet index and fetch from Drive if empty
  useEffect(() => {
    setActiveSheetIndex(0);
    setSortConfig(null);
    const unitSheets = allUnitsData[selectedUnit];
    const totalRows = (unitSheets || []).reduce((acc, s) => acc + (s.rows ? s.rows.length : 0), 0);
    if (totalRows === 0) {
      fetchExcelFromDrive(selectedUnit, false);
    }
  }, [selectedUnit]);

  /**
   * Reads the Excel file for unitKey directly from Google Drive.
   */
  const fetchExcelFromDrive = async (unitKey: string, notifyUser = true): Promise<boolean> => {
    const fileName = getOlayTakipDriveFileName(unitKey);
    setIsSyncing(true);
    try {
      const res = await fetch(GOOGLE_SCRIPT_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({
          action: 'readExcelFromDrive',
          fileName: fileName,
          folderId: DRIVE_FOLDER_ID
        })
      });

      if (!res.ok) {
        if (notifyUser) showNotification(`Drive bağlantısı kurulamadı (${res.status}).`, 'error');
        return false;
      }

      const data = await res.json();
      if (data.status === 'success' && data.base64) {
        const binaryString = atob(data.base64);
        const len = binaryString.length;
        const bytes = new Uint8Array(len);
        for (let i = 0; i < len; i++) {
          bytes[i] = binaryString.charCodeAt(i);
        }
        const wb = XLSX.read(bytes, { type: 'array' });
        const parsedSheets = parseOlayTakipWorkbook(wb);
        if (parsedSheets.length > 0) {
          setAllUnitsData(prev => {
            const updated = { ...prev, [unitKey]: parsedSheets };
            safeSetJSON(STORAGE_KEY, updated);
            return updated;
          });
          if (data.updated) {
            setLastSyncTime(data.updated);
          } else {
            setLastSyncTime(
              new Date().toLocaleString('tr-TR', {
                day: '2-digit',
                month: '2-digit',
                year: 'numeric',
                hour: '2-digit',
                minute: '2-digit'
              })
            );
          }
          if (notifyUser) {
            showNotification(`"${fileName}" Google Drive'dan başarıyla okundu (${parsedSheets.length} sayfa).`, 'success');
          }
          return true;
        }
      } else {
        if (notifyUser) {
          showNotification(`"${fileName}" Drive'da henüz bulunmuyor. 'VERİ GÜNCELLE' ile Excel yükleyebilirsiniz.`, 'info');
        }
      }
      return false;
    } catch (err: any) {
      console.warn(`Drive Excel okuma hatası (${unitKey}):`, err);
      if (notifyUser) {
        showNotification(`Drive okuma hatası: ${err?.message || err}`, 'error');
      }
      return false;
    } finally {
      setIsSyncing(false);
    }
  };

  /**
   * Upload an Excel file to Drive under standard canonical filename, then read & adapt sheets.
   */
  const handleUploadExcelToDrive = async (file: File) => {
    setIsSyncing(true);
    setSyncProgressPercent(15);
    setSyncProgressStage('1/4: Excel dosyası yerel olarak okunuyor...');

    try {
      const arrayBuffer = await file.arrayBuffer();
      setSyncProgressPercent(45);
      setSyncProgressStage('2/4: Matris, sayfa yapısı ve sütun başlıkları taranıyor...');

      const wb = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
      const parsedSheets = parseOlayTakipWorkbook(wb);
      if (parsedSheets.length === 0) {
        showNotification('Seçilen Excel dosyasında okunabilir sayfa veya sütun bulunamadı.', 'error');
        setIsSyncing(false);
        return;
      }

      setSyncProgressPercent(70);
      setSyncProgressStage(`3/4: Drive'a '${getOlayTakipDriveFileName(selectedUnit)}' olarak kaydediliyor...`);

      let binary = '';
      const bytes = new Uint8Array(arrayBuffer);
      for (let i = 0; i < bytes.byteLength; i++) {
        binary += String.fromCharCode(bytes[i]);
      }
      const base64Data = btoa(binary);
      const standardFileName = getOlayTakipDriveFileName(selectedUnit);

      try {
        const apiRes = await fetch('/api/upload-techizat-excel', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            fileName: standardFileName,
            targetKey: `olay_takip_${selectedUnit}`,
            base64Data: base64Data,
            folderId: DRIVE_FOLDER_ID
          })
        });
        if (!apiRes.ok) throw new Error('API route proxy failed');
      } catch (localErr) {
        try {
          await fetch(GOOGLE_SCRIPT_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
            body: JSON.stringify({
              action: 'uploadTechizatExcel',
              fileName: standardFileName,
              targetKey: `olay_takip_${selectedUnit}`,
              base64Data: base64Data,
              folderId: DRIVE_FOLDER_ID
            })
          });
        } catch (gasErr) {
          console.warn('Drive upload fallback error:', gasErr);
        }
      }

      setSyncProgressPercent(95);
      setSyncProgressStage('4/4: Sistem çizelgeleri ve dinamik sayfalar güncelleniyor...');

      setAllUnitsData(prev => {
        const updated = { ...prev, [selectedUnit]: parsedSheets };
        safeSetJSON(STORAGE_KEY, updated);
        return updated;
      });

      const nowStr = new Date().toLocaleString('tr-TR', {
        day: '2-digit',
        month: '2-digit',
        year: 'numeric',
        hour: '2-digit',
        minute: '2-digit'
      });
      setLastSyncTime(nowStr);
      setActiveSheetIndex(0);

      setTimeout(() => {
        setSyncProgressPercent(100);
        setSyncProgressStage('Tamamlandı!');
        setTimeout(() => {
          setIsSyncing(false);
          setIsDataSyncModalOpen(false);
          showNotification(
            `"${standardFileName}" Google Drive'a başarıyla yüklendi! (${parsedSheets.length} sayfa ve ${parsedSheets.reduce((a, s) => a + s.rows.length, 0)} kayıt güncellendi)`,
            'success'
          );
        }, 400);
      }, 300);
    } catch (err: any) {
      console.error('Olay takip Excel yükleme hatası:', err);
      setIsSyncing(false);
      showNotification('Excel dosyası işlenirken hata oluştu: ' + (err?.message || err), 'error');
    } finally {
      if (syncFileInputRef.current) syncFileInputRef.current.value = '';
    }
  };

  /**
   * Save updated rows in background to Google Drive as an updated XLSX.
   */
  const persistAndSyncToDrive = (updated: Record<string, OlayTakipSheet[]>) => {
    // Ensure all sheets have EK DOSYA BAĞLANTI LİNKİ column
    const standardized = ensureLinkColumnInAllUnitsData(updated);
    setAllUnitsData(standardized);
    safeSetJSON(STORAGE_KEY, standardized);

    try {
      const sheets = standardized[selectedUnit];
      if (!sheets || sheets.length === 0) return;

      const wb = XLSX.utils.book_new();
      sheets.forEach(sheet => {
        const { sheet: guaranteedSheet, colIndex: linkIdx } = ensureEkDosyaLinkColumn(sheet);
        const aoa: any[][] = [];
        aoa.push(guaranteedSheet.columns);
        
        guaranteedSheet.rows.forEach((r, rIdx) => {
          const rowCopy = [...r];
          // Auto fill link if empty and row has docs
          if (linkIdx >= 0 && (!rowCopy[linkIdx] || !rowCopy[linkIdx].trim())) {
            const docs = getRowDocs(rIdx, rowCopy[belgeColIndex] || '');
            if (docs.length > 0) {
              const link = docs[0].driveUrl || (docs[0].driveFileId ? `https://drive.google.com/file/d/${docs[0].driveFileId}/view` : '');
              if (link) rowCopy[linkIdx] = link;
            }
          }
          aoa.push(rowCopy);
        });

        const ws = XLSX.utils.aoa_to_sheet(aoa);
        XLSX.utils.book_append_sheet(wb, ws, sheet.sheetName.slice(0, 31));
      });
      const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      let binary = '';
      const bytes = new Uint8Array(wbout);
      for (let i = 0; i < bytes.byteLength; i++) {
        binary += String.fromCharCode(bytes[i]);
      }
      const base64Data = btoa(binary);
      const standardFileName = getOlayTakipDriveFileName(selectedUnit);

      fetch('/api/upload-techizat-excel', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          fileName: standardFileName,
          targetKey: `olay_takip_${selectedUnit}`,
          base64Data: base64Data,
          folderId: DRIVE_FOLDER_ID
        })
      }).catch(err => console.warn('Background sync warning:', err));
    } catch (e) {
      console.warn('Background XLSX generation warning:', e);
    }
  };

  const currentUnitSheets = useMemo(() => {
    return allUnitsData[selectedUnit] || [];
  }, [allUnitsData, selectedUnit]);

  const currentSheet: OlayTakipSheet | undefined = useMemo(() => {
    if (!currentUnitSheets || currentUnitSheets.length === 0) return undefined;
    return currentUnitSheets[activeSheetIndex] || currentUnitSheets[0];
  }, [currentUnitSheets, activeSheetIndex]);

  // İlk kolon zaten "Sıra" veya "Sıra No" ise çift sıra no olmasını engelle
  const isCol0Sira = useMemo(() => {
    if (!currentSheet || !currentSheet.columns || currentSheet.columns.length === 0) return false;
    const col0 = currentSheet.columns[0].trim().toLowerCase();
    return /^(s[ıi]ra|s\.?\s*no\.?|s\.?\s*nu\.?|no\.?|s[ıi]ra\s*no)$/i.test(col0);
  }, [currentSheet]);

  // Belge / Döküman sütunu indexi
  const belgeColIndex = useMemo(() => {
    if (!currentSheet || !currentSheet.columns) return -1;
    return currentSheet.columns.findIndex(c => /belge|d[oö]k[uü]man/i.test(c));
  }, [currentSheet]);

  // EK DOSYA BAĞLANTI LİNKİ sütunu indexi
  const ekDosyaLinkColIndex = useMemo(() => {
    if (!currentSheet || !currentSheet.columns) return -1;
    return currentSheet.columns.findIndex(c => isLinkColumnName(c));
  }, [currentSheet]);

  // Row filtering & sorting
  const processedRows = useMemo(() => {
    if (!currentSheet) return [];
    let rows = [...currentSheet.rows];

    // Filter
    if (searchQuery.trim()) {
      const q = searchQuery.toLowerCase().trim();
      rows = rows.filter(row =>
        row.some(cell => String(cell || '').toLowerCase().includes(q))
      );
    }

    // Sort
    if (sortConfig) {
      const { colIndex, direction } = sortConfig;
      rows.sort((a, b) => {
        const valA = String(a[colIndex] || '').trim();
        const valB = String(b[colIndex] || '').trim();

        // Ng Limit number check
        const numA = extractNgValue(valA);
        const numB = extractNgValue(valB);
        if (numA !== null && numB !== null) {
          return direction === 'asc' ? numA - numB : numB - numA;
        }

        // Date check for date columns
        const colName = currentSheet?.columns[colIndex] || '';
        const isDateCol = /tar[ıi]h|date/i.test(colName);
        if (isDateCol) {
          const toTime = (v: string): number => {
            if (!v) return 0;
            if (/^\d{5}(\.\d+)?$/.test(v)) return parseFloat(v) * 86400000;
            const cleaned = cleanAndFormatDateString(v);
            const m = cleaned.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{4})/);
            if (m) {
              return new Date(parseInt(m[3], 10), parseInt(m[2], 10) - 1, parseInt(m[1], 10)).getTime();
            }
            return 0;
          };
          const tA = toTime(valA);
          const tB = toTime(valB);
          if (tA > 0 || tB > 0) {
            return direction === 'asc' ? tA - tB : tB - tA;
          }
        }

        // Numeric check
        const parsedA = parseFloat(valA.replace(',', '.'));
        const parsedB = parseFloat(valB.replace(',', '.'));
        if (!isNaN(parsedA) && !isNaN(parsedB)) {
          return direction === 'asc' ? parsedA - parsedB : parsedB - parsedA;
        }

        return direction === 'asc'
          ? valA.localeCompare(valB, 'tr')
          : valB.localeCompare(valA, 'tr');
      });
    }

    return rows;
  }, [currentSheet, searchQuery, sortConfig]);

  const handleSortColumn = (cIdx: number) => {
    if (sortConfig && sortConfig.colIndex === cIdx) {
      if (sortConfig.direction === 'asc') {
        setSortConfig({ colIndex: cIdx, direction: 'desc' });
      } else {
        setSortConfig(null);
      }
    } else {
      const colName = currentSheet?.columns[cIdx] || '';
      const isDescDefault = /ng|limit|tepe|sayı|tarih/i.test(colName);
      setSortConfig({ colIndex: cIdx, direction: isDescDefault ? 'desc' : 'asc' });
    }
  };

  // Require Password for protected actions
  const requirePassword = (action: () => void) => {
    setPendingAction(() => action);
    setPasswordInput('');
    setPasswordError('');
    setPasswordModalOpen(true);
  };

  const handleVerifyPassword = (e: React.FormEvent) => {
    e.preventDefault();
    const currentUnitCfg = UNITS_CONFIG.find(u => u.key === selectedUnit) || UNITS_CONFIG[0];
    const expectedPassword = currentUnitCfg.password || '1839';
    if (passwordInput !== expectedPassword && passwordInput !== '1839') {
      setPasswordError('Hatalı yetki şifresi girdiniz!');
      return;
    }
    setPasswordModalOpen(false);
    if (pendingAction) {
      pendingAction();
      setPendingAction(null);
    }
  };

  const handleOpenDataSyncModal = () => {
    setDataSyncStep(1);
    setSyncPasswordInput('');
    setSyncPasswordError('');
    setIsDataSyncModalOpen(true);
  };

  const handleVerifySyncPassword = (e: React.FormEvent) => {
    e.preventDefault();
    const currentUnitCfg = UNITS_CONFIG.find(u => u.key === selectedUnit) || UNITS_CONFIG[0];
    const expectedPassword = currentUnitCfg.password || '1839';
    if (syncPasswordInput !== expectedPassword && syncPasswordInput !== '1839') {
      setSyncPasswordError('Hatalı yetki şifresi girdiniz!');
      return;
    }
    setDataSyncStep(2);
  };

  // Save edited row
  const handleSaveRow = (rowIndex: number, updatedRow: string[]) => {
    if (!currentSheet) return;
    const updatedSheets = currentUnitSheets.map((sheet, idx) => {
      if (idx !== activeSheetIndex) return sheet;
      const newRows = [...sheet.rows];
      newRows[rowIndex] = updatedRow;
      return { ...sheet, rows: newRows };
    });
    const updatedAll = { ...allUnitsData, [selectedUnit]: updatedSheets };
    persistAndSyncToDrive(updatedAll);
    setEditingRow(null);
    showNotification('Kayıt başarıyla güncellendi ve Drive ile senkronize edildi.', 'success');
  };

  // Delete row
  const handleDeleteRow = (rowIndex: number) => {
    requirePassword(() => {
      if (!currentSheet) return;
      const updatedSheets = currentUnitSheets.map((sheet, idx) => {
        if (idx !== activeSheetIndex) return sheet;
        const newRows = sheet.rows.filter((_, rIdx) => rIdx !== rowIndex);
        return { ...sheet, rows: newRows };
      });
      const updatedAll = { ...allUnitsData, [selectedUnit]: updatedSheets };
      persistAndSyncToDrive(updatedAll);
      showNotification('Kayıt silindi ve Drive güncellendi.', 'info');
    });
  };

  // Add new row directly (no password blockage for standard data entry)
  const handleOpenAddRow = () => {
    if (!currentSheet) return;
    const initialVals = new Array(currentSheet.columns.length).fill('');
    if (isCol0Sira) {
      initialVals[0] = String(currentSheet.rows.length + 1);
    }
    setNewRowValues(initialVals);
    setIsAddingRow(true);
  };

  const handleSaveNewRow = () => {
    if (!currentSheet) return;
    const updatedSheets = currentUnitSheets.map((sheet, idx) => {
      if (idx !== activeSheetIndex) return sheet;
      return { ...sheet, rows: [...sheet.rows, newRowValues] };
    });
    const updatedAll = { ...allUnitsData, [selectedUnit]: updatedSheets };
    persistAndSyncToDrive(updatedAll);
    setIsAddingRow(false);
    showNotification('Yeni satır başarıyla eklendi ve Drive ile senkronize edildi.', 'success');
  };

  // Create New Sheet (Excel mantığı ile yeni sayfa oluşturma)
  const handleCreateNewSheet = () => {
    if (!newSheetName.trim()) {
      showNotification('Lütfen geçerli bir sayfa adı giriniz.', 'error');
      return;
    }

    const rawCols = newSheetColumnsText
      .split(/[,;\n]+/)
      .map(c => c.trim())
      .filter(Boolean);

    const columns = rawCols.length > 0
      ? rawCols
      : ['Sıra', 'Tarih', 'Kuyruk No', 'Olay Açıklaması', 'Hava Aracı / Parça Durumu', 'Belge / Döküman'];

    if (!columns.some(c => /belge|d[oö]k[uü]man/i.test(c))) {
      columns.push('Belge / Döküman');
    }

    const newSheet: OlayTakipSheet = {
      id: `sheet_${Date.now()}_${Math.random().toString(36).substring(2, 6)}`,
      sheetName: newSheetName.trim(),
      title: newSheetTitle.trim() || newSheetName.trim().toUpperCase(),
      columns: columns,
      rows: []
    };

    const updatedSheets = [...currentUnitSheets, newSheet];
    const updatedAll = { ...allUnitsData, [selectedUnit]: updatedSheets };
    persistAndSyncToDrive(updatedAll);

    setIsAddingSheet(false);
    setNewSheetName('');
    setNewSheetTitle('');
    setActiveSheetIndex(updatedSheets.length - 1);
    showNotification(`Yeni Excel sayfası '${newSheet.sheetName}' başarıyla oluşturuldu!`, 'success');
  };

  // Double-Click Handlers for Sheet Names, Sheet Title, and Column Headers
  const handleSaveSheetName = (sIdx: number) => {
    if (editingSheetIndex === null) return;
    const newName = editingSheetNameValue.trim();
    if (!newName) {
      setEditingSheetIndex(null);
      return;
    }
    const updatedSheets = currentUnitSheets.map((s, idx) =>
      idx === sIdx ? { ...s, sheetName: newName } : s
    );
    const updatedAll = { ...allUnitsData, [selectedUnit]: updatedSheets };
    persistAndSyncToDrive(updatedAll);
    setEditingSheetIndex(null);
    showNotification(`Sayfa adı "${newName}" olarak güncellendi.`, 'success');
  };

  const handleSaveSheetTitle = () => {
    if (!isEditingSheetTitle || !currentSheet) return;
    const newTitle = editingSheetTitleValue.trim();
    if (!newTitle) {
      setIsEditingSheetTitle(false);
      return;
    }
    const updatedSheets = currentUnitSheets.map((s, idx) =>
      idx === activeSheetIndex ? { ...s, title: newTitle } : s
    );
    const updatedAll = { ...allUnitsData, [selectedUnit]: updatedSheets };
    persistAndSyncToDrive(updatedAll);
    setIsEditingSheetTitle(false);
    showNotification(`Sayfa başlığı "${newTitle}" olarak güncellendi.`, 'success');
  };

  const handleSaveColumnName = (colIdx: number) => {
    if (editingColIndex === null || !currentSheet) return;
    const newColName = editingColNameValue.trim();
    if (!newColName) {
      setEditingColIndex(null);
      return;
    }
    const updatedCols = [...currentSheet.columns];
    updatedCols[colIdx] = newColName;
    const updatedSheets = currentUnitSheets.map((s, idx) =>
      idx === activeSheetIndex ? { ...s, columns: updatedCols } : s
    );
    const updatedAll = { ...allUnitsData, [selectedUnit]: updatedSheets };
    persistAndSyncToDrive(updatedAll);
    setEditingColIndex(null);
    showNotification(`Sütun başlığı "${newColName}" olarak güncellendi.`, 'success');
  };

  // Document Management for a specific row
  const getRowDocKey = (rowIndex: number): string => {
    return `olay_${selectedUnit}_${currentSheet?.id || 'sheet'}_row_${rowIndex}`;
  };

  const getRowDocs = (rowIndex: number, cellVal: string): HangarPdfDoc[] => {
    const row = currentSheet?.rows ? currentSheet.rows[rowIndex] : undefined;
    return hangarPdfDocs.filter(d => {
      return matchOlayTakipDoc(d, {
        unitKey: selectedUnit,
        sheetId: currentSheet?.id,
        rowIndex,
        cellVal: cellVal || '',
        rowCells: row
      });
    });
  };

  const handleOpenDocModal = (rowIndex: number, row: string[], defaultTab: 'upload' | 'drive' | 'link' = 'upload') => {
    setActiveDocRow({ rowIndex, row: [...row] });
    setDocUploadFile(null);
    setDocUploadName('');
    setDocUploadCategory('Kaza/Kırım & Olay Belgesi');
    setDriveDocSearchTerm('');
    setDocModalTab(defaultTab);
    const existingLink = ekDosyaLinkColIndex >= 0 ? (row[ekDosyaLinkColIndex] || '') : '';
    setCustomLinkUrl(existingLink);
    setCustomLinkTitle('');
  };

  /**
   * Opens a document or link in a new browser tab
   */
  const handleOpenDocInNewTab = (doc?: HangarPdfDoc | null, directUrl?: string) => {
    const rawUrl = (directUrl || doc?.driveUrl || (doc?.driveFileId ? `https://drive.google.com/file/d/${doc.driveFileId}/view` : ''))?.trim();
    if (rawUrl && (rawUrl.startsWith('http://') || rawUrl.startsWith('https://'))) {
      window.open(rawUrl, '_blank', 'noopener,noreferrer');
      return;
    }
    if (doc?.driveFileId) {
      window.open(`/api/pdf-proxy?fileId=${doc.driveFileId}&fileName=${encodeURIComponent(doc.fileName || 'belge.pdf')}`, '_blank');
      return;
    }
    if (doc?.fileData && doc.fileData.startsWith('data:')) {
      try {
        const arr = doc.fileData.split(',');
        const mime = arr[0].match(/:(.*?);/)?.[1] || 'application/pdf';
        const bstr = atob(arr[1]);
        let n = bstr.length;
        const u8arr = new Uint8Array(n);
        while (n--) {
          u8arr[n] = bstr.charCodeAt(n);
        }
        const blob = new Blob([u8arr], { type: mime });
        const blobUrl = URL.createObjectURL(blob);
        window.open(blobUrl, '_blank');
        return;
      } catch {
        window.open(doc.fileData, '_blank');
        return;
      }
    }
    showNotification('Açılacak geçerli bir belge veya bağlantı linki bulunamadı.', 'info');
  };

  const handleSaveDocToRow = async () => {
    if (!activeDocRow || !docUploadFile || !currentSheet) return;

    setIsDocUploading(true);
    try {
      // Convert file to base64
      const reader = new FileReader();
      const base64Promise = new Promise<string>((resolve, reject) => {
        reader.onload = () => resolve(reader.result as string);
        reader.onerror = reject;
      });
      reader.readAsDataURL(docUploadFile);
      const fileDataUrl = await base64Promise;

      const finalDocName = docUploadName.trim() || docUploadFile.name;
      const docKey = getRowDocKey(activeDocRow.rowIndex);
      const cleanBase64 = fileDataUrl.includes(',') ? fileDataUrl.split(',')[1] : fileDataUrl;
      const fileMimeType = getDocMimeType(finalDocName, fileDataUrl) || docUploadFile.type || 'application/pdf';

      let driveFileId = '';
      let driveUrl = '';

      // 1. Upload via backend proxy (/api/upload-hangar-pdf)
      try {
        const proxyRes = await fetch('/api/upload-hangar-pdf', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            fileName: `${docKey}___${finalDocName}`,
            base64Data: cleanBase64,
            mimeType: fileMimeType,
            folderId: DRIVE_FOLDER_ID,
            itemKey: docKey,
            docType: docUploadCategory,
            firma: 'OGM Havacılık'
          })
        });
        if (proxyRes.ok) {
          const proxyData = await proxyRes.json();
          if (proxyData && proxyData.fileId) {
            driveFileId = proxyData.fileId;
            driveUrl = proxyData.viewUrl || `https://drive.google.com/file/d/${driveFileId}/preview`;
          }
        }
      } catch (proxyErr) {
        console.warn('Proxy upload failed, attempting direct GAS fallback:', proxyErr);
      }

      // 2. Direct fallback to Google Apps Script if needed
      if (!driveFileId) {
        try {
          const uploadRes = await fetch(GOOGLE_SCRIPT_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
            body: JSON.stringify({
              action: 'uploadPdfToDrive',
              fileName: `${docKey}___${finalDocName}`,
              mimeType: fileMimeType,
              base64Data: cleanBase64,
              folderId: DRIVE_FOLDER_ID,
              itemKey: docKey,
              docType: docUploadCategory,
              firma: 'OGM Havacılık'
            })
          });
          if (uploadRes.ok) {
            const driveData = await uploadRes.json();
            if (driveData && (driveData.fileId || driveData.id)) {
              driveFileId = driveData.fileId || driveData.id;
              driveUrl = driveData.viewUrl || `https://drive.google.com/file/d/${driveFileId}/preview`;
            }
          }
        } catch (driveErr) {
          console.warn('Google Drive direct upload warning:', driveErr);
        }
      }

      const effectiveLink = driveUrl || (driveFileId ? `https://drive.google.com/file/d/${driveFileId}/view` : '');

      const newDoc: HangarPdfDoc = {
        id: driveFileId ? `drive_doc_${driveFileId}` : `olay_doc_${Date.now()}_${Math.random().toString(36).substring(2, 7)}`,
        itemKey: docKey,
        fileName: finalDocName,
        fileData: fileDataUrl,
        docType: docUploadCategory,
        firma: 'OGM Havacılık',
        mimeType: fileMimeType,
        uploadDate: new Date().toLocaleDateString('tr-TR', {
          day: '2-digit',
          month: '2-digit',
          year: 'numeric',
          hour: '2-digit',
          minute: '2-digit'
        }),
        fileSize: `${(docUploadFile.size / (1024 * 1024) > 1
          ? (docUploadFile.size / (1024 * 1024)).toFixed(2) + ' MB'
          : (docUploadFile.size / 1024).toFixed(1) + ' KB')}`,
        uploadedAt: new Date().toLocaleDateString('tr-TR', {
          day: '2-digit',
          month: '2-digit',
          year: 'numeric',
          hour: '2-digit',
          minute: '2-digit'
        }),
        driveFileId: driveFileId || undefined,
        driveUrl: effectiveLink || undefined
      };

      await saveHangarPdfDoc(newDoc);
      setRowDocLink(docKey, newDoc.id, true);
      if (driveFileId) {
        setRowDocLink(docKey, driveFileId, true);
      }
      setHangarPdfDocs(prev => [newDoc, ...prev.filter(d => d.id !== newDoc.id)]);

      // Ensure EK DOSYA BAĞLANTI LİNKİ column exists and update row
      const { sheet: guaranteedSheet, colIndex: linkIdx } = ensureEkDosyaLinkColumn(currentSheet);
      const updatedRow = [...(guaranteedSheet.rows[activeDocRow.rowIndex] || activeDocRow.row)];
      while (updatedRow.length <= Math.max(belgeColIndex, linkIdx)) updatedRow.push('');

      if (linkIdx >= 0 && effectiveLink) {
        updatedRow[linkIdx] = effectiveLink;
      }

      // Update row cell text with doc name
      const docColIdx = guaranteedSheet.columns.findIndex(c => /belge|d[oö]k[uü]man/i.test(c));
      if (docColIdx >= 0) {
        const existingVal = updatedRow[docColIdx] || '';
        if (!existingVal.includes(finalDocName)) {
          updatedRow[docColIdx] = existingVal ? `${existingVal}, ${finalDocName}` : finalDocName;
        }
      }

      const updatedRows = [...guaranteedSheet.rows];
      updatedRows[activeDocRow.rowIndex] = updatedRow;
      const updatedSheet = { ...guaranteedSheet, rows: updatedRows };
      const updatedSheets = currentUnitSheets.map((s, idx) =>
        idx === activeSheetIndex ? updatedSheet : s
      );
      const updatedAll = { ...allUnitsData, [selectedUnit]: updatedSheets };
      persistAndSyncToDrive(updatedAll);
      setActiveDocRow({ rowIndex: activeDocRow.rowIndex, row: updatedRow });

      setDocUploadFile(null);
      setDocUploadName('');
      showNotification(`"${finalDocName}" belgesi yüklendi ve bağlantı linki Excel sütununa kaydedildi!`, 'success');
    } catch (err: any) {
      console.error('Doc upload error:', err);
      showNotification('Belge yüklenirken hata oluştu: ' + (err?.message || err), 'error');
    } finally {
      setIsDocUploading(false);
    }
  };

  const handleAttachExistingDriveDoc = async (driveDoc: HangarPdfDoc) => {
    if (!activeDocRow || !currentSheet) return;
    const docKey = getRowDocKey(activeDocRow.rowIndex);
    const cleanDocName = driveDoc.fileName || 'Belge.pdf';

    setRowDocLink(docKey, driveDoc.id, true);
    if (driveDoc.driveFileId) {
      setRowDocLink(docKey, driveDoc.driveFileId, true);
    }

    const effectiveLink = driveDoc.driveUrl || (driveDoc.driveFileId ? `https://drive.google.com/file/d/${driveDoc.driveFileId}/view` : '');

    const updatedDoc: HangarPdfDoc = {
      ...driveDoc,
      itemKey: docKey,
      driveUrl: effectiveLink || driveDoc.driveUrl
    };
    await saveHangarPdfDoc(updatedDoc);
    setHangarPdfDocs(prev => [updatedDoc, ...prev.filter(d => d.id !== driveDoc.id)]);

    const { sheet: guaranteedSheet, colIndex: linkIdx } = ensureEkDosyaLinkColumn(currentSheet);
    const updatedRow = [...(guaranteedSheet.rows[activeDocRow.rowIndex] || activeDocRow.row)];
    while (updatedRow.length <= Math.max(belgeColIndex, linkIdx)) updatedRow.push('');

    if (linkIdx >= 0 && effectiveLink) {
      updatedRow[linkIdx] = effectiveLink;
    }

    const docColIdx = guaranteedSheet.columns.findIndex(c => /belge|d[oö]k[uü]man/i.test(c));
    if (docColIdx >= 0) {
      const existingVal = updatedRow[docColIdx] || '';
      if (!existingVal.includes(cleanDocName)) {
        updatedRow[docColIdx] = existingVal ? `${existingVal}, ${cleanDocName}` : cleanDocName;
      }
    }

    const updatedRows = [...guaranteedSheet.rows];
    updatedRows[activeDocRow.rowIndex] = updatedRow;
    const updatedSheet = { ...guaranteedSheet, rows: updatedRows };
    const updatedSheets = currentUnitSheets.map((s, idx) =>
      idx === activeSheetIndex ? updatedSheet : s
    );
    const updatedAll = { ...allUnitsData, [selectedUnit]: updatedSheets };
    persistAndSyncToDrive(updatedAll);
    setActiveDocRow({ rowIndex: activeDocRow.rowIndex, row: updatedRow });

    showNotification(`"${cleanDocName}" bu satıra başarıyla bağlandı ve link sütununa eklendi!`, 'success');
  };

  /**
   * Saves or updates a custom external URL / Google Drive link for this row
   */
  const handleSaveCustomLink = async () => {
    if (!activeDocRow || !currentSheet) return;
    const rawUrl = customLinkUrl.trim();
    if (!rawUrl) {
      showNotification('Lütfen geçerli bir bağlantı linki (URL) giriniz.', 'error');
      return;
    }
    let finalUrl = rawUrl;
    if (!/^https?:\/\//i.test(finalUrl)) {
      finalUrl = 'https://' + finalUrl;
    }
    const finalTitle = customLinkTitle.trim() || 'Ek Dosya Bağlantısı';
    const docKey = getRowDocKey(activeDocRow.rowIndex);

    // Ensure EK DOSYA BAĞLANTI LİNKİ column exists
    const { sheet: guaranteedSheet, colIndex: linkIdx } = ensureEkDosyaLinkColumn(currentSheet);
    const updatedRow = [...(guaranteedSheet.rows[activeDocRow.rowIndex] || activeDocRow.row)];
    while (updatedRow.length <= Math.max(belgeColIndex, linkIdx)) updatedRow.push('');
    
    if (linkIdx >= 0) {
      updatedRow[linkIdx] = finalUrl;
    }

    // Also update Belge / Döküman column if exists
    const docColIdx = guaranteedSheet.columns.findIndex(c => /belge|d[oö]k[uü]man/i.test(c));
    if (docColIdx >= 0) {
      const curDocVal = updatedRow[docColIdx] || '';
      if (!curDocVal.includes(finalTitle)) {
        updatedRow[docColIdx] = curDocVal ? `${curDocVal}, ${finalTitle}` : finalTitle;
      }
    }

    // Create a virtual HangarPdfDoc for this link so it is listed with actions (open, preview, delete)
    const linkDoc: HangarPdfDoc = {
      id: `link_${Date.now()}_${Math.random().toString(36).substring(2, 6)}`,
      itemKey: docKey,
      fileName: finalTitle,
      fileData: '',
      docType: 'Dış Bağlantı / Drive Linki',
      firma: 'OGM / Dış Bağlantı',
      uploadDate: new Date().toLocaleDateString('tr-TR', { day: '2-digit', month: '2-digit', year: 'numeric' }),
      fileSize: 'Web Bağlantısı',
      uploadedAt: new Date().toLocaleDateString('tr-TR'),
      driveUrl: finalUrl
    };

    await saveHangarPdfDoc(linkDoc);
    setRowDocLink(docKey, linkDoc.id, true);
    setHangarPdfDocs(prev => [linkDoc, ...prev.filter(d => d.id !== linkDoc.id)]);

    // Update sheet and save
    const updatedRows = [...guaranteedSheet.rows];
    updatedRows[activeDocRow.rowIndex] = updatedRow;
    const updatedSheet = { ...guaranteedSheet, rows: updatedRows };
    const updatedSheets = currentUnitSheets.map((s, idx) =>
      idx === activeSheetIndex ? updatedSheet : s
    );
    const updatedAll = { ...allUnitsData, [selectedUnit]: updatedSheets };
    persistAndSyncToDrive(updatedAll);

    setActiveDocRow({ rowIndex: activeDocRow.rowIndex, row: updatedRow });
    setCustomLinkTitle('');
    showNotification(`"${finalTitle}" bağlantı linki satıra eklendi ve Excel'deki "${EK_DOSYA_LINK_COL_NAME}" sütununa kaydedildi.`, 'success');
  };

  /**
   * Removes custom external URL link from row
   */
  const handleRemoveCustomLink = async () => {
    if (!activeDocRow || !currentSheet) return;
    const { sheet: guaranteedSheet, colIndex: linkIdx } = ensureEkDosyaLinkColumn(currentSheet);
    const updatedRow = [...(guaranteedSheet.rows[activeDocRow.rowIndex] || activeDocRow.row)];
    if (linkIdx >= 0 && updatedRow.length > linkIdx) {
      updatedRow[linkIdx] = '';
    }

    const updatedRows = [...guaranteedSheet.rows];
    updatedRows[activeDocRow.rowIndex] = updatedRow;
    const updatedSheet = { ...guaranteedSheet, rows: updatedRows };
    const updatedSheets = currentUnitSheets.map((s, idx) =>
      idx === activeSheetIndex ? updatedSheet : s
    );
    const updatedAll = { ...allUnitsData, [selectedUnit]: updatedSheets };
    persistAndSyncToDrive(updatedAll);

    setActiveDocRow({ rowIndex: activeDocRow.rowIndex, row: updatedRow });
    setCustomLinkUrl('');
    showNotification('Bağlantı linki satırdan kaldırıldı.', 'info');
  };

  const handleDeleteDoc = async (docId: string, docName: string) => {
    try {
      const targetDoc = hangarPdfDocs.find(d => d.id === docId || d.driveFileId === docId);
      const driveFileId = targetDoc?.driveFileId || (docId.startsWith('drive_doc_') ? docId.replace('drive_doc_', '') : '');
      const fileName = targetDoc?.fileName || docName;

      // 1. Delete from IndexedDB, localStorage, and send trash command to Google Drive
      await deleteHangarPdfDoc(docId, driveFileId, fileName);

      // 2. Clear row links
      if (activeDocRow) {
        const docKey = getRowDocKey(activeDocRow.rowIndex);
        setRowDocLink(docKey, docId, false);
        if (driveFileId) setRowDocLink(docKey, driveFileId, false);
      }

      // 3. Update state
      setHangarPdfDocs(prev => prev.filter(d => 
        d.id !== docId && 
        (!driveFileId || d.driveFileId !== driveFileId) &&
        (!fileName || d.fileName !== fileName)
      ));

      // 4. Update row cells: clean Belge / Döküman AND EK DOSYA BAĞLANTI LİNKİ
      if (activeDocRow && currentSheet) {
        const updatedRow = [...activeDocRow.row];
        
        // Clean Belge / Döküman
        if (belgeColIndex >= 0) {
          const currentCell = updatedRow[belgeColIndex] || '';
          const cleaned = currentCell
            .split(',')
            .map(s => s.trim())
            .filter(s => s && s !== docName && s !== fileName)
            .join(', ');
          updatedRow[belgeColIndex] = cleaned;
        }

        // Clean EK DOSYA BAĞLANTI LİNKİ if it matches this doc
        if (ekDosyaLinkColIndex >= 0) {
          const currentLink = updatedRow[ekDosyaLinkColIndex] || '';
          if ((driveFileId && currentLink.includes(driveFileId)) || (targetDoc?.driveUrl && currentLink === targetDoc.driveUrl)) {
            updatedRow[ekDosyaLinkColIndex] = '';
          }
        }

        handleSaveRow(activeDocRow.rowIndex, updatedRow);
        setActiveDocRow({ rowIndex: activeDocRow.rowIndex, row: updatedRow });
      }

      showNotification(`"${fileName}" belgesi Google Drive'da çöpe taşındı ve satırdan silindi.`, 'success');
    } catch (err: any) {
      console.warn('Doc delete error:', err);
      showNotification('Belge silinirken hata: ' + (err?.message || err), 'error');
    }
  };

  /**
   * Downloads the complete workbook / current table in styled Excel format (.xlsx)
   * Ensures the 'EK DOSYA BAĞLANTI LİNKİ' column is automatically created and populated!
   */
  const handleDownloadExcel = () => {
    try {
      const sheetsToExport = allUnitsData[selectedUnit] || [];
      if (sheetsToExport.length === 0) {
        showNotification('İndirilecek çizelge verisi bulunamadı.', 'info');
        return;
      }

      const wb = XLSX.utils.book_new();

      sheetsToExport.forEach((sheet, sIdx) => {
        // Ensure 'EK DOSYA BAĞLANTI LİNKİ' column exists
        const { sheet: guaranteedSheet, colIndex: linkIdx } = ensureEkDosyaLinkColumn(sheet);
        const docColIdx = guaranteedSheet.columns.findIndex(c => /belge|d[oö]k[uü]man/i.test(c));

        // Format data rows and guarantee link column is filled if row has attached docs
        const formattedRows = guaranteedSheet.rows.map((r, rIdx) => {
          const rowCopy = [...r];
          while (rowCopy.length < guaranteedSheet.columns.length) rowCopy.push('');

          // Auto-fill link if missing but row has attached doc
          if (linkIdx >= 0 && (!rowCopy[linkIdx] || !String(rowCopy[linkIdx]).trim())) {
            const docs = getRowDocs(rIdx, docColIdx >= 0 ? (rowCopy[docColIdx] || '') : '');
            if (docs.length > 0) {
              const link = docs[0].driveUrl || (docs[0].driveFileId ? `https://drive.google.com/file/d/${docs[0].driveFileId}/view` : '');
              if (link) rowCopy[linkIdx] = link;
            }
          }

          return guaranteedSheet.columns.map((colName, cIdx) => {
            const cell = rowCopy[cIdx] || '';
            const isDateCol = /tar[ıi]h|date/i.test(colName);
            if (isDateCol || (/^\d{5}(\.\d+)?$/.test(String(cell)) && Number(cell) >= 20000 && Number(cell) <= 90000 && isDateCol)) {
              return cleanAndFormatDateString(cell);
            }
            return cell;
          });
        });

        const tableData = [guaranteedSheet.columns, ...formattedRows];
        const ws = XLSX.utils.aoa_to_sheet(tableData);

        // Auto-calculate column widths matching content
        const colWidths = guaranteedSheet.columns.map((colName, cIdx) => {
          let maxLen = (colName || '').length;
          guaranteedSheet.rows.forEach(r => {
            const cellVal = String(r[cIdx] || '');
            if (cellVal.length > maxLen) maxLen = Math.min(cellVal.length, 60);
          });
          return { wch: Math.max(maxLen + 4, 14) };
        });
        ws['!cols'] = colWidths;

        // Clean sheet name (max 31 chars, no invalid chars : \ / ? * [ ])
        let safeSheetName = (guaranteedSheet.sheetName || `Sayfa ${sIdx + 1}`)
          .replace(/[\\/*?:\[\]]/g, '')
          .slice(0, 31);
        if (!safeSheetName) safeSheetName = `Sayfa ${sIdx + 1}`;

        let finalSheetName = safeSheetName;
        let counter = 1;
        while (wb.SheetNames.includes(finalSheetName)) {
          finalSheetName = `${safeSheetName.slice(0, 27)}_${counter++}`;
        }

        XLSX.utils.book_append_sheet(wb, ws, finalSheetName);
      });

      const fileName = `Olay_Takip_Cizelgesi_${selectedUnit.toUpperCase()}_${new Date().toLocaleDateString('tr-TR').replace(/\./g, '_')}.xlsx`;
      XLSX.writeFile(wb, fileName);
      showNotification(`✅ "${fileName}" çizelgesi ("EK DOSYA BAĞLANTI LİNKİ" sütunuyla birlikte) Excel olarak başarıyla indirildi!`, 'success');
    } catch (err) {
      console.error('Excel export error:', err);
      showNotification('Excel dosyası oluşturulurken hata: ' + (err as Error).message, 'error');
    }
  };

  if (!isOpen) return null;

  const activeUnitConfig = UNITS_CONFIG.find(u => u.key === selectedUnit) || UNITS_CONFIG[0];

  return (
    <div
      id="olay-takip-cizelgesi-modal"
      className="fixed inset-0 z-[9999] flex items-center justify-center p-2 sm:p-4 bg-slate-900/70 backdrop-blur-sm animate-fade-in"
    >
      {/* MASTER MODAL CONTAINER */}
      <div className="bg-white text-slate-800 w-full max-w-7xl h-[94vh] rounded-2xl shadow-2xl border border-slate-200 flex flex-col overflow-hidden">
        {/* HEADER TOOLBAR */}
        <div className="px-5 py-3.5 bg-white border-b border-slate-200 flex flex-wrap items-center justify-between gap-3 shrink-0">
          <div className="flex items-center gap-3">
            {/* Prominent Back Button */}
            <button
              type="button"
              onClick={onClose}
              className="flex items-center gap-1.5 px-3.5 py-2 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold text-xs border border-slate-200 transition-all cursor-pointer shadow-xs active:scale-95"
              title="Geri Dön"
            >
              <ArrowLeft className="w-4 h-4 text-[#0b3d1d]" />
              <span>Geri</span>
            </button>

            {/* Icon and Clean Title */}
            <div className="w-10 h-10 rounded-xl bg-emerald-50 border border-emerald-200 flex items-center justify-center shadow-xs text-[#0b3d1d]">
              <FileSpreadsheet className="w-5 h-5 text-[#0b3d1d]" />
            </div>
            <div>
              <h2 className="text-base sm:text-lg font-black tracking-tight text-[#0b3d1d]">
                OLAY TAKİP ÇİZELGESİ
              </h2>
              <p className="text-[11px] text-slate-500 font-medium">
                Hava Araçları Önemli Olay, Kaza-Kırım ve Motor Limit Aşımları Canlı Takip Sistemi
              </p>
            </div>
          </div>

          {/* Action Tools */}
          <div className="flex items-center gap-2.5">
            {/* Search Input */}
            <div className="relative w-48 sm:w-64">
              <Search className="w-3.5 h-3.5 absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
              <input
                type="text"
                placeholder="Çizelgede Ara (Kuyruk, Olay, Pilot...)"
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                className="w-full pl-8 pr-3 py-1.5 bg-slate-50 hover:bg-slate-100/80 focus:bg-white text-xs text-slate-800 placeholder:text-slate-400 border border-slate-200 rounded-xl focus:outline-none focus:border-[#0b3d1d] transition-all font-medium"
              />
            </div>

            {/* Drive Refresh Button */}
            <button
              type="button"
              onClick={() => fetchExcelFromDrive(selectedUnit, true)}
              disabled={isSyncing}
              className="p-2 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-700 border border-slate-200 transition-all cursor-pointer disabled:opacity-50"
              title="Drive'dan Yenile"
            >
              <RotateCw className={`w-4 h-4 text-emerald-800 ${isSyncing ? 'animate-spin' : ''}`} />
            </button>

            {/* UNIFIED "VERİ GÜNCELLE" BUTTON */}
            <button
              type="button"
              onClick={handleOpenDataSyncModal}
              disabled={isSyncing}
              className="px-4 py-2 rounded-xl bg-[#0b3d1d] hover:bg-[#072612] text-white font-black text-xs uppercase tracking-wider transition-all flex items-center gap-2 shadow-sm cursor-pointer active:scale-95 disabled:opacity-50"
            >
              <Database className="w-3.5 h-3.5 text-emerald-300" />
              <span>VERİ GÜNCELLE</span>
            </button>

            {/* EXCEL İNDİR BUTTON */}
            <button
              type="button"
              onClick={handleDownloadExcel}
              className="px-4 py-2 rounded-xl bg-emerald-700 hover:bg-emerald-600 text-white font-black text-xs uppercase tracking-wider transition-all flex items-center gap-2 shadow-sm cursor-pointer active:scale-95 border border-emerald-600/50"
              title="Görüntülenen Çizelgeyi Excel Formatında İndir"
            >
              <Download className="w-3.5 h-3.5 text-emerald-200" />
              <span>EXCEL İNDİR</span>
            </button>

            {/* Close Button */}
            <button
              type="button"
              onClick={onClose}
              className="p-2 rounded-xl bg-slate-100 hover:bg-rose-50 text-slate-600 hover:text-rose-700 border border-slate-200 transition-all cursor-pointer active:scale-95"
              title="Kapat"
            >
              <X className="w-4 h-4" />
            </button>
          </div>
        </div>

        {/* SUB-HEADER: AIR/LAND UNIT SELECTION BAR */}
        <div className="px-5 py-2.5 bg-slate-50 border-b border-slate-200 flex items-center gap-2 overflow-x-auto shrink-0">
          <span className="text-[11px] font-black text-slate-500 uppercase tracking-wider shrink-0 mr-1 font-mono">
            FİLO / BİRİM:
          </span>
          {UNITS_CONFIG.map(unit => {
            const isSelected = selectedUnit === unit.key;
            return (
              <button
                key={unit.key}
                type="button"
                onClick={() => setSelectedUnit(unit.key)}
                className={`px-3 py-1.5 rounded-xl text-xs font-black tracking-wide transition-all shrink-0 cursor-pointer border ${
                  isSelected
                    ? 'bg-[#0b3d1d] text-white border-[#0b3d1d] shadow-sm'
                    : 'bg-white text-slate-700 hover:bg-slate-100 border-slate-200'
                }`}
              >
                {unit.label}
              </button>
            );
          })}
        </div>

        {/* WORKBOOK SHEET TABS & ACTION BUTTONS */}
        <div className="px-5 py-2 bg-white border-b border-slate-200 flex items-center justify-between gap-3 overflow-x-auto shrink-0">
          <div className="flex items-center gap-1.5 overflow-x-auto">
            {currentUnitSheets.map((sheet, sIdx) => {
              const isActive = activeSheetIndex === sIdx;
              const isEditingThisSheet = editingSheetIndex === sIdx;
              return (
                <div key={sheet.id || sIdx} className="relative shrink-0">
                  {isEditingThisSheet ? (
                    <div className="flex items-center bg-white border-2 border-emerald-500 rounded-lg px-2 py-1 shadow-md">
                      <input
                        type="text"
                        autoFocus
                        value={editingSheetNameValue}
                        onChange={(e) => setEditingSheetNameValue(e.target.value)}
                        onKeyDown={(e) => {
                          if (e.key === 'Enter') handleSaveSheetName(sIdx);
                          if (e.key === 'Escape') setEditingSheetIndex(null);
                        }}
                        onBlur={() => handleSaveSheetName(sIdx)}
                        className="text-xs font-bold text-slate-900 bg-transparent focus:outline-none w-28"
                        placeholder="Sayfa adı..."
                      />
                      <button
                        type="button"
                        onMouseDown={() => handleSaveSheetName(sIdx)}
                        className="ml-1 text-emerald-700 hover:text-emerald-900 text-xs font-black cursor-pointer"
                        title="Kaydet"
                      >
                        ✓
                      </button>
                    </div>
                  ) : (
                    <button
                      type="button"
                      onClick={() => {
                        setActiveSheetIndex(sIdx);
                        setSortConfig(null);
                      }}
                      onDoubleClick={(e) => {
                        e.stopPropagation();
                        setEditingSheetIndex(sIdx);
                        setEditingSheetNameValue(sheet.sheetName);
                      }}
                      className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all flex items-center gap-1.5 shrink-0 cursor-pointer border select-none group ${
                        isActive
                          ? 'bg-emerald-600 text-white border-emerald-600 shadow-xs'
                          : 'bg-slate-50 text-slate-600 hover:bg-slate-100 border-slate-200'
                      }`}
                      title="Seçmek için tıkla, adı düzenlemek için çift tıkla"
                    >
                      <FileText className="w-3 h-3 opacity-80" />
                      <span>{sheet.sheetName}</span>
                      <Edit3 className="w-2.5 h-2.5 opacity-0 group-hover:opacity-70 transition-opacity" />
                      <span
                        className={`text-[9px] font-mono px-1 rounded ${
                          isActive ? 'bg-emerald-700 text-white' : 'bg-slate-200/70 text-slate-600'
                        }`}
                      >
                        {sheet.rows ? sheet.rows.length : 0}
                      </span>
                    </button>
                  )}
                </div>
              );
            })}
          </div>

          {/* Action Buttons: Yeni Sayfa Ekle & Yeni Satır Ekle */}
          <div className="flex items-center gap-2 shrink-0">
            {/* Yeni Sayfa Ekle Button */}
            <button
              type="button"
              onClick={() => setIsAddingSheet(true)}
              className="px-3 py-1.5 rounded-lg bg-emerald-700 hover:bg-emerald-800 text-white font-black text-xs transition-all flex items-center gap-1.5 shrink-0 shadow-xs cursor-pointer active:scale-95"
              title="Excel Mantığı ile Yeni Sayfa / Çizelge Ekle"
            >
              <Plus className="w-3.5 h-3.5" />
              <span>Yeni Sayfa Ekle</span>
            </button>

            {/* Yeni Satır Ekle Button */}
            {currentSheet && (
              <button
                type="button"
                onClick={handleOpenAddRow}
                className="px-3 py-1.5 rounded-lg bg-[#0b3d1d] hover:bg-[#072612] text-white font-black text-xs transition-all flex items-center gap-1.5 shrink-0 shadow-xs cursor-pointer active:scale-95"
                title="Mevcut Sayfaya Yeni Satır Ekle"
              >
                <Plus className="w-3.5 h-3.5" />
                <span>Yeni Satır Ekle</span>
              </button>
            )}
          </div>
        </div>

        {/* MAIN CONTENT AREA - TABLE CONTAINER */}
        <div className="flex-1 overflow-hidden flex flex-col bg-slate-50/50">
          {currentSheet ? (
            <div className="flex-1 flex flex-col overflow-hidden">
              {/* Sheet Title Bar */}
              <div className="px-5 py-2.5 bg-white border-b border-slate-200 flex items-center justify-between shrink-0">
                <div className="flex items-center gap-2">
                  <span className="w-2 h-2 rounded-full bg-[#0b3d1d]" />
                  {isEditingSheetTitle ? (
                    <div className="flex items-center bg-white border border-emerald-500 rounded-lg px-2 py-0.5 shadow-sm">
                      <input
                        type="text"
                        autoFocus
                        value={editingSheetTitleValue}
                        onChange={(e) => setEditingSheetTitleValue(e.target.value)}
                        onKeyDown={(e) => {
                          if (e.key === 'Enter') handleSaveSheetTitle();
                          if (e.key === 'Escape') setIsEditingSheetTitle(false);
                        }}
                        onBlur={handleSaveSheetTitle}
                        className="text-xs font-black tracking-wide text-slate-800 uppercase font-mono bg-transparent focus:outline-none w-64"
                        placeholder="Sayfa başlığı..."
                      />
                      <button
                        type="button"
                        onMouseDown={handleSaveSheetTitle}
                        className="ml-1 text-emerald-700 hover:text-emerald-900 text-xs font-black cursor-pointer"
                        title="Kaydet"
                      >
                        ✓
                      </button>
                    </div>
                  ) : (
                    <h3
                      onDoubleClick={() => {
                        setIsEditingSheetTitle(true);
                        setEditingSheetTitleValue(currentSheet.title || currentSheet.sheetName);
                      }}
                      className="text-xs font-black tracking-wide text-slate-800 uppercase font-mono cursor-pointer hover:text-emerald-700 hover:underline transition-colors flex items-center gap-1.5 group select-none"
                      title="Başlığı düzenlemek için çift tıklayınız"
                    >
                      <span>{currentSheet.title || currentSheet.sheetName}</span>
                      <Edit3 className="w-3 h-3 opacity-30 group-hover:opacity-100 text-slate-500 transition-opacity" />
                    </h3>
                  )}
                  <span className="text-[10px] text-slate-400 font-normal ml-2">
                    (Sayfa adı, başlık veya sütunları çift tıklayarak düzenleyebilirsiniz)
                  </span>
                </div>
                <div className="text-[11px] text-slate-500 font-mono">
                  Toplam <strong className="text-slate-800 font-black">{currentSheet.rows.length}</strong> kayıt
                  {searchQuery.trim() && (
                    <span> (Filtrelenen: <strong className="text-emerald-700 font-black">{processedRows.length}</strong>)</span>
                  )}
                </div>
              </div>

              {/* Responsive Scrollable Table */}
              <div className="flex-1 overflow-auto">
                <table className="w-full text-left border-collapse bg-white">
                  <thead className="sticky top-0 z-10 bg-slate-100/95 backdrop-blur-xs border-b border-slate-200 shadow-xs">
                    <tr>
                      {/* Tek Sıra No Kuralı: Eğer kolon 0 Sıra No ise fazladan # kolonu basma */}
                      {!isCol0Sira && (
                        <th className="px-3 py-2.5 text-[11px] font-black text-slate-600 uppercase tracking-wider border-r border-slate-200/80 w-12 text-center">
                          #
                        </th>
                      )}
                      {currentSheet.columns.map((col, idx) => {
                        const isSorted = sortConfig && sortConfig.colIndex === idx;
                        const isEditingThisCol = editingColIndex === idx;
                        return (
                          <th
                            key={idx}
                            onClick={() => {
                              if (!isEditingThisCol) handleSortColumn(idx);
                            }}
                            onDoubleClick={(e) => {
                              e.stopPropagation();
                              setEditingColIndex(idx);
                              setEditingColNameValue(col);
                            }}
                            className="px-4 py-2.5 text-[11px] font-black text-slate-700 uppercase tracking-wider border-r border-slate-200/80 whitespace-nowrap cursor-pointer hover:bg-slate-200/70 select-none transition-colors"
                            title="Sıralamak için tek tık, başlığı düzenlemek için çift tıklayınız"
                          >
                            {isEditingThisCol ? (
                              <div
                                onClick={(e) => e.stopPropagation()}
                                className="flex items-center bg-white border-2 border-emerald-500 rounded px-1.5 py-0.5 shadow-sm"
                              >
                                <input
                                  type="text"
                                  autoFocus
                                  value={editingColNameValue}
                                  onChange={(e) => setEditingColNameValue(e.target.value)}
                                  onKeyDown={(e) => {
                                    if (e.key === 'Enter') handleSaveColumnName(idx);
                                    if (e.key === 'Escape') setEditingColIndex(null);
                                  }}
                                  onBlur={() => handleSaveColumnName(idx)}
                                  className="text-[11px] font-bold text-slate-900 bg-transparent focus:outline-none w-32"
                                />
                                <button
                                  type="button"
                                  onMouseDown={() => handleSaveColumnName(idx)}
                                  className="ml-1 text-emerald-700 hover:text-emerald-900 text-xs font-black cursor-pointer"
                                >
                                  ✓
                                </button>
                              </div>
                            ) : (
                              <div className="flex items-center gap-1.5 group">
                                <span>{col}</span>
                                <Edit3 className="w-2.5 h-2.5 opacity-0 group-hover:opacity-70 text-slate-400 transition-opacity" />
                                {isSorted ? (
                                  sortConfig.direction === 'asc' ? (
                                    <ArrowUp className="w-3.5 h-3.5 text-emerald-800" />
                                  ) : (
                                    <ArrowDown className="w-3.5 h-3.5 text-emerald-800" />
                                  )
                                ) : (
                                  <ArrowUpDown className="w-3 h-3 text-slate-400 opacity-40 group-hover:opacity-100" />
                                )}
                              </div>
                            )}
                          </th>
                        );
                      })}
                      <th className="px-3 py-2.5 text-[11px] font-black text-slate-600 uppercase tracking-wider text-center w-24">
                        İŞLEM
                      </th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {processedRows.length === 0 ? (
                      <tr>
                        <td
                          colSpan={currentSheet.columns.length + (isCol0Sira ? 1 : 2)}
                          className="px-6 py-16 text-center"
                        >
                          <div className="max-w-md mx-auto flex flex-col items-center justify-center text-slate-400">
                            <div className="w-14 h-14 rounded-2xl bg-slate-100 flex items-center justify-center mb-3 text-slate-400 border border-slate-200">
                              <FileSpreadsheet className="w-7 h-7 text-slate-400" />
                            </div>
                            <h4 className="text-sm font-black text-slate-700 mb-1">
                              {searchQuery ? 'Aramanıza Uygun Kayıt Bulunamadı' : 'Bu Çizelgede Henüz Kayıt Bulunmuyor'}
                            </h4>
                            <p className="text-xs text-slate-500 leading-relaxed text-center mb-4">
                              {searchQuery
                                ? 'Farklı bir arama terimi deneyebilir veya filtreyi temizleyebilirsiniz.'
                                : "Veri aktarımı için yukarıdaki 'VERİ GÜNCELLE' butonunu kullanarak güncel Excel dosyasını Drive'a yükleyebilir veya 'Yeni Satır Ekle' butonuyla manuel kayıt girebilirsiniz."}
                            </p>
                            {!searchQuery && (
                              <button
                                type="button"
                                onClick={handleOpenDataSyncModal}
                                className="px-4 py-2 rounded-xl bg-[#0b3d1d] hover:bg-[#072612] text-white font-black text-xs uppercase tracking-wider transition-all flex items-center gap-2 shadow-sm cursor-pointer"
                              >
                                <Database className="w-3.5 h-3.5 text-emerald-300" />
                                <span>VERİ GÜNCELLE</span>
                              </button>
                            )}
                          </div>
                        </td>
                      </tr>
                    ) : (
                      processedRows.map((row, rIdx) => {
                        const originalIndex = currentSheet.rows.indexOf(row);
                        const rowDocs = getRowDocs(originalIndex, row[belgeColIndex] || '');

                        return (
                          <tr
                            key={rIdx}
                            onDoubleClick={() => {
                              const formattedRow = row.map((cell, cIdx) => {
                                const col = currentSheet.columns[cIdx] || '';
                                if (/tar[ıi]h|date/i.test(col)) {
                                  return cleanAndFormatDateString(cell);
                                }
                                return cell;
                              });
                              setEditingRow({ rowIndex: originalIndex, row: formattedRow });
                            }}
                            title="Kayıt düzenlemek için çift tıklayınız"
                            className="hover:bg-emerald-50/50 transition-colors border-b border-slate-100 text-xs text-slate-700 cursor-pointer group"
                          >
                            {/* Sequence number if not column 0 */}
                            {!isCol0Sira && (
                              <td className="px-3 py-3 text-center text-slate-400 font-mono text-[11px] border-r border-slate-100">
                                {originalIndex + 1}
                              </td>
                            )}

                            {currentSheet.columns.map((colName, cIdx) => {
                              const cellValue = row[cIdx] || '';
                              const lowerCol = colName.toLowerCase();
                              const isKuyruk = lowerCol.includes('kuyruk');
                              const isDateCol = /tar[ıi]h|date/i.test(colName);
                              const isDurum =
                                lowerCol.includes('durum') ||
                                lowerCol.includes('durumu') ||
                                lowerCol.includes('sonuç') ||
                                lowerCol.includes('karar');
                              const isNgCol = lowerCol.includes('ng') || lowerCol.includes('limit') || lowerCol.includes('tepe');
                              const isDocCol = lowerCol.includes('belge') || lowerCol.includes('döküman') || lowerCol.includes('dokuman');
                              const isLinkCol = isLinkColumnName(colName);

                              // Check Ng Limit Styling
                              const ngStyle = isNgCol ? getNgLimitBadgeStyle(cellValue) : { bgClass: '', isNg: false };

                              // Check Status Styling
                              const statusStyle = isDurum ? getStatusBadgeStyle(cellValue) : { bgClass: '', isStatus: false };

                              // First column as Sıra No
                              if (cIdx === 0 && isCol0Sira) {
                                return (
                                  <td
                                    key={cIdx}
                                    className="px-3 py-3 text-center text-slate-500 font-mono text-[11px] font-bold border-r border-slate-100 w-12"
                                  >
                                    {cellValue || (originalIndex + 1)}
                                  </td>
                                );
                              }

                              // Document Attachment Column
                              if (isDocCol) {
                                const associatedLink = ekDosyaLinkColIndex >= 0 ? (row[ekDosyaLinkColIndex] || '') : '';
                                return (
                                  <td
                                    key={cIdx}
                                    className="px-3 py-2 border-r border-slate-100 align-middle whitespace-nowrap min-w-[140px]"
                                    onClick={(e) => e.stopPropagation()}
                                  >
                                    <div className="flex items-center gap-1.5 flex-wrap">
                                      {rowDocs.length > 0 ? (
                                        <>
                                          {rowDocs.map(doc => (
                                            <div key={doc.id} className="inline-flex items-center gap-0.5 rounded-md bg-emerald-50 border border-emerald-200 shadow-2xs overflow-hidden">
                                              <button
                                                type="button"
                                                onClick={() => setPreviewDoc(doc)}
                                                className="inline-flex items-center gap-1 px-2 py-1 hover:bg-emerald-100 text-emerald-800 text-[11px] font-bold transition-all cursor-pointer max-w-[130px] truncate"
                                                title={`${doc.fileName} (${doc.fileSize || 'PDF'}) - Önizle`}
                                              >
                                                <FileText className="w-3 h-3 text-emerald-700 shrink-0" />
                                                <span className="truncate">{doc.fileName}</span>
                                              </button>
                                              <button
                                                type="button"
                                                onClick={() => handleOpenDocInNewTab(doc)}
                                                className="p-1 text-emerald-700 hover:text-emerald-950 hover:bg-emerald-200 border-l border-emerald-200 transition-colors cursor-pointer"
                                                title="Yeni Sekmede Aç"
                                              >
                                                <ExternalLink className="w-3 h-3" />
                                              </button>
                                            </div>
                                          ))}
                                          <button
                                            type="button"
                                            onClick={() => handleOpenDocModal(originalIndex, row)}
                                            className="p-1 rounded-md bg-slate-100 hover:bg-emerald-100 text-slate-600 hover:text-emerald-800 border border-slate-200 transition-all cursor-pointer"
                                            title="Yeni Belge Ekle / Link Bağla"
                                          >
                                            <Plus className="w-3 h-3" />
                                          </button>
                                        </>
                                      ) : associatedLink ? (
                                        <div className="inline-flex items-center gap-1">
                                          <button
                                            type="button"
                                            onClick={() => handleOpenDocInNewTab(null, associatedLink)}
                                            className="inline-flex items-center gap-1 px-2.5 py-1 rounded-md bg-blue-50 hover:bg-blue-100 text-blue-800 border border-blue-200 text-[11px] font-bold transition-all cursor-pointer"
                                            title="Bağlantıyı Yeni Sekmede Aç"
                                          >
                                            <ExternalLink className="w-3 h-3 text-blue-600" />
                                            <span>Yeni Sekmede Aç</span>
                                          </button>
                                          <button
                                            type="button"
                                            onClick={() => handleOpenDocModal(originalIndex, row, 'link')}
                                            className="p-1 rounded-md bg-slate-100 hover:bg-blue-100 text-slate-600 hover:text-blue-800 border border-slate-200 transition-all cursor-pointer"
                                            title="Link / Belgeyi Düzenle"
                                          >
                                            <Edit3 className="w-3 h-3" />
                                          </button>
                                        </div>
                                      ) : (
                                        <div className="inline-flex items-center gap-1">
                                          <button
                                            type="button"
                                            onClick={() => handleOpenDocModal(originalIndex, row, 'upload')}
                                            className="inline-flex items-center gap-1 px-2 py-1 rounded-md bg-slate-100 hover:bg-emerald-50 text-slate-600 hover:text-emerald-800 border border-slate-200 text-[11px] font-medium transition-all cursor-pointer"
                                            title="PDF Belge / Döküman Ekle"
                                          >
                                            <Paperclip className="w-3 h-3 text-slate-400" />
                                            <span>+ Belge</span>
                                          </button>
                                          <button
                                            type="button"
                                            onClick={() => handleOpenDocModal(originalIndex, row, 'link')}
                                            className="inline-flex items-center gap-1 px-2 py-1 rounded-md bg-blue-50/70 hover:bg-blue-100 text-blue-700 border border-blue-200 text-[11px] font-medium transition-all cursor-pointer"
                                            title="Bağlantı Linki Bağla"
                                          >
                                            <LinkIcon className="w-3 h-3 text-blue-500" />
                                            <span>+ Link</span>
                                          </button>
                                        </div>
                                      )}
                                    </div>
                                  </td>
                                );
                              }

                              // EK DOSYA BAĞLANTI LİNKİ Column
                              if (isLinkCol) {
                                const effectiveLink = (cellValue || (rowDocs.length > 0 ? (rowDocs[0].driveUrl || (rowDocs[0].driveFileId ? `https://drive.google.com/file/d/${rowDocs[0].driveFileId}/view` : '')) : ''))?.trim();

                                return (
                                  <td
                                    key={cIdx}
                                    className="px-3 py-2 border-r border-slate-100 align-middle whitespace-nowrap min-w-[150px] max-w-xs truncate"
                                    onClick={(e) => e.stopPropagation()}
                                  >
                                    {effectiveLink ? (
                                      <div className="flex items-center gap-1.5">
                                        <a
                                          href={effectiveLink}
                                          target="_blank"
                                          rel="noopener noreferrer"
                                          className="inline-flex items-center gap-1 px-2.5 py-1 rounded-md bg-blue-50 hover:bg-blue-100 text-blue-800 hover:text-blue-950 border border-blue-200 text-[11px] font-semibold transition-all shadow-2xs max-w-[180px] truncate"
                                          title={`${effectiveLink} - Yeni sekmede açmak için tıklayın`}
                                        >
                                          <ExternalLink className="w-3 h-3 text-blue-600 shrink-0" />
                                          <span className="truncate">{effectiveLink}</span>
                                        </a>
                                        <button
                                          type="button"
                                          onClick={() => handleOpenDocModal(originalIndex, row, 'link')}
                                          className="p-1 rounded bg-slate-100 hover:bg-slate-200 text-slate-600 hover:text-slate-800 text-[10px] cursor-pointer"
                                          title="Linki Düzenle"
                                        >
                                          <Edit3 className="w-3 h-3" />
                                        </button>
                                      </div>
                                    ) : (
                                      <button
                                        type="button"
                                        onClick={() => handleOpenDocModal(originalIndex, row, 'link')}
                                        className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-[11px] text-slate-400 hover:text-blue-700 hover:bg-blue-50 border border-dashed border-slate-200 hover:border-blue-300 transition-colors cursor-pointer"
                                        title="Bağlantı Linki Ekle"
                                      >
                                        <LinkIcon className="w-3 h-3" />
                                        <span>+ Link Ekle</span>
                                      </button>
                                    )}
                                  </td>
                                );
                              }

                              return (
                                <td
                                  key={cIdx}
                                  className="px-4 py-3 border-r border-slate-100 align-middle max-w-md"
                                >
                                  {isKuyruk && cellValue ? (
                                    <span className="font-mono font-bold text-amber-900 bg-amber-50 border border-amber-200 px-2.5 py-0.5 rounded-md inline-block">
                                      {cellValue}
                                    </span>
                                  ) : ngStyle.isNg && cellValue ? (
                                    <span
                                      className={`inline-block px-2.5 py-0.5 rounded-md text-[11px] ${ngStyle.bgClass}`}
                                    >
                                      {cellValue}
                                    </span>
                                  ) : statusStyle.isStatus && cellValue ? (
                                    <span
                                      className={`inline-block px-2.5 py-0.5 rounded-md text-[11px] ${statusStyle.bgClass}`}
                                    >
                                      {cellValue}
                                    </span>
                                  ) : isDateCol || (/^\d{5}(\.\d+)?$/.test(cellValue) && Number(cellValue) >= 20000 && Number(cellValue) <= 90000) ? (
                                    <span className="text-slate-800 font-mono font-medium leading-relaxed">
                                      {cleanAndFormatDateString(cellValue) || '-'}
                                    </span>
                                  ) : (
                                    <span className="text-slate-700 font-normal leading-relaxed">
                                      {cellValue || '-'}
                                    </span>
                                  )}
                                </td>
                              );
                            })}

                            <td
                              className="px-3 py-3 text-center align-middle whitespace-nowrap"
                              onClick={(e) => e.stopPropagation()}
                            >
                              <div className="flex items-center justify-center gap-1.5">
                                <button
                                  type="button"
                                  onClick={() => {
                                    const formattedRow = row.map((cell, cIdx) => {
                                      const col = currentSheet.columns[cIdx] || '';
                                      if (/tar[ıi]h|date/i.test(col)) {
                                        return cleanAndFormatDateString(cell);
                                      }
                                      return cell;
                                    });
                                    setEditingRow({ rowIndex: originalIndex, row: formattedRow });
                                  }}
                                  className="p-1.5 rounded-lg bg-slate-100 hover:bg-amber-100 text-slate-600 hover:text-amber-800 transition-all cursor-pointer"
                                  title="Düzenle"
                                >
                                  <Edit3 className="w-3.5 h-3.5" />
                                </button>
                                <button
                                  type="button"
                                  onClick={() => handleDeleteRow(originalIndex)}
                                  className="p-1.5 rounded-lg bg-slate-100 hover:bg-rose-100 text-slate-600 hover:text-rose-700 transition-all cursor-pointer"
                                  title="Sil"
                                >
                                  <Trash2 className="w-3.5 h-3.5" />
                                </button>
                              </div>
                            </td>
                          </tr>
                        );
                      })
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          ) : (
            <div className="flex-1 flex items-center justify-center text-slate-400">
              <p>Sayfa seçiniz.</p>
            </div>
          )}
        </div>

        {/* BOTTOM STATUS FOOTER BAR */}
        <div className="bg-slate-50 px-5 py-2.5 border-t border-slate-200 flex items-center justify-between text-xs text-slate-500 shrink-0">
          <div className="flex items-center gap-2 font-mono">
            <span className="w-2 h-2 rounded-full bg-emerald-500 animate-pulse" />
            <span>
              Son Drive Senkronizasyonu:{' '}
              <strong className="text-slate-700 font-bold">{lastSyncTime}</strong>
            </span>
            <span className="text-slate-300">|</span>
            <span>
              Dosya: <strong className="text-emerald-800 font-bold">{getOlayTakipDriveFileName(selectedUnit)}</strong>
            </span>
          </div>
          <div className="font-semibold text-slate-600">
            OGM Havacılık Dairesi Başkanlığı • Olay Takip Çizelgesi Modülü
          </div>
        </div>
      </div>

      {/* DEDICATED "VERİ GÜNCELLE" MODAL */}
      {isDataSyncModalOpen && (
        <div className="fixed inset-0 z-[10001] flex items-center justify-center p-4 bg-slate-900/70 backdrop-blur-sm animate-fade-in">
          <div className="bg-white text-slate-800 w-full max-w-lg rounded-2xl shadow-2xl border border-slate-200 overflow-hidden flex flex-col">
            {/* Header */}
            <div className="px-6 py-4 bg-white border-b border-slate-200 flex items-center justify-between">
              <div className="flex items-center gap-2.5">
                <div className="w-8 h-8 rounded-xl bg-emerald-50 border border-emerald-200 flex items-center justify-center text-[#0b3d1d]">
                  <Database className="w-4 h-4 text-[#0b3d1d]" />
                </div>
                <div>
                  <h3 className="text-sm font-black text-[#0b3d1d] uppercase tracking-wider">
                    VERİ GÜNCELLE • {activeUnitConfig.label}
                  </h3>
                  <p className="text-[11px] text-slate-500 font-medium">
                    Google Drive Olay Takip Excel Senkronizasyonu
                  </p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => setIsDataSyncModalOpen(false)}
                className="text-slate-400 hover:text-slate-700 text-sm p-1 rounded-lg hover:bg-slate-100 cursor-pointer"
              >
                ✕
              </button>
            </div>

            {/* Content Body */}
            <div className="p-6">
              {dataSyncStep === 1 ? (
                <form onSubmit={handleVerifySyncPassword} className="space-y-4">
                  <div className="text-center mb-4">
                    <div className="w-12 h-12 rounded-full bg-emerald-50 text-[#0b3d1d] mx-auto flex items-center justify-center mb-2 border border-emerald-200">
                      <Lock className="w-6 h-6 text-[#0b3d1d]" />
                    </div>
                    <h4 className="text-sm font-black text-slate-800">Yetkili Doğrulaması</h4>
                    <p className="text-xs text-slate-500 mt-1">
                      {activeUnitConfig.label} olay takip çizelgesini güncellemek için yetkili şifresini giriniz.
                    </p>
                  </div>

                  <div className="space-y-1.5">
                    <label className="text-xs font-bold text-slate-700 uppercase font-mono">
                      Şifre (Varsayılan: {activeUnitConfig.password || '1839'})
                    </label>
                    <input
                      type="password"
                      autoFocus
                      value={syncPasswordInput}
                      onChange={(e) => {
                        setSyncPasswordInput(e.target.value);
                        setSyncPasswordError('');
                      }}
                      placeholder="Yetkili şifresi..."
                      className="w-full px-4 py-2.5 bg-slate-50 border border-slate-300 rounded-xl text-sm text-slate-900 font-mono tracking-widest text-center focus:outline-none focus:border-[#0b3d1d] focus:bg-white"
                    />
                    {syncPasswordError && (
                      <p className="text-xs text-rose-600 font-bold text-center mt-1">{syncPasswordError}</p>
                    )}
                  </div>

                  <div className="flex gap-2 pt-2">
                    <button
                      type="button"
                      onClick={() => setIsDataSyncModalOpen(false)}
                      className="flex-1 py-2.5 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold text-xs cursor-pointer"
                    >
                      İptal
                    </button>
                    <button
                      type="submit"
                      className="flex-1 py-2.5 rounded-xl bg-[#0b3d1d] hover:bg-[#072612] text-white font-black text-xs uppercase tracking-wider cursor-pointer shadow-sm"
                    >
                      Onayla & Devam Et
                    </button>
                  </div>
                </form>
              ) : (
                <div className="space-y-4">
                  <div className="bg-emerald-50/70 border border-emerald-200 rounded-xl p-3 text-xs text-emerald-900 flex items-start gap-2.5">
                    <CheckCircle2 className="w-4 h-4 text-emerald-700 shrink-0 mt-0.5" />
                    <div>
                      <span className="font-bold">Hedef Google Drive Dosyası:</span>
                      <div className="font-mono text-[11px] text-emerald-800 font-bold mt-0.5">
                        {getOlayTakipDriveFileName(selectedUnit)}
                      </div>
                      <p className="text-[10px] text-emerald-700 mt-0.5 leading-tight">
                        Yükleyeceğiniz Excel dosyası otomatik olarak Drive klasörüne yüklenecek ve sistem sayfalarını canlı olarak adapte edecektir.
                      </p>
                    </div>
                  </div>

                  {isSyncing ? (
                    <div className="py-6 space-y-3 text-center">
                      <div className="w-10 h-10 rounded-full bg-emerald-100 text-[#0b3d1d] mx-auto flex items-center justify-center animate-spin">
                        <RotateCw className="w-5 h-5 text-[#0b3d1d]" />
                      </div>
                      <p className="text-xs font-bold text-slate-800">{syncProgressStage}</p>
                      <div className="w-full bg-slate-200 rounded-full h-2.5 overflow-hidden">
                        <div
                          className="bg-[#0b3d1d] h-2.5 rounded-full transition-all duration-300"
                          style={{ width: `${syncProgressPercent}%` }}
                        />
                      </div>
                      <span className="text-[11px] font-mono text-slate-500">%{syncProgressPercent}</span>
                    </div>
                  ) : (
                    <div className="space-y-3">
                      <input
                        ref={syncFileInputRef}
                        type="file"
                        accept=".xlsx, .xls"
                        className="hidden"
                        onChange={(e) => {
                          const file = e.target.files?.[0];
                          if (file) handleUploadExcelToDrive(file);
                        }}
                      />
                      <div
                        onClick={() => syncFileInputRef.current?.click()}
                        className="border-2 border-dashed border-emerald-300 hover:border-emerald-600 bg-emerald-50/30 hover:bg-emerald-50/60 rounded-2xl p-8 flex flex-col items-center justify-center text-center cursor-pointer transition-all group"
                      >
                        <div className="w-14 h-14 rounded-2xl bg-emerald-100 flex items-center justify-center mb-3 group-hover:scale-105 transition-transform text-[#0b3d1d]">
                          <UploadCloud className="w-7 h-7 text-[#0b3d1d]" />
                        </div>
                        <h4 className="text-sm font-black text-slate-800 mb-1">
                          Excel Dosyasını Seçiniz (.xlsx, .xls)
                        </h4>
                        <p className="text-xs text-slate-500 max-w-xs">
                          Bilgisayarınızdaki güncel olay takip Excel çizelgesini seçin veya buraya tıklayın.
                        </p>
                      </div>

                      <div className="pt-2 border-t border-slate-200 flex items-center justify-between">
                        <button
                          type="button"
                          onClick={() => {
                            fetchExcelFromDrive(selectedUnit, true);
                            setIsDataSyncModalOpen(false);
                          }}
                          className="text-xs font-bold text-emerald-800 hover:text-emerald-950 flex items-center gap-1.5 cursor-pointer py-1"
                        >
                          <RotateCw className="w-3.5 h-3.5 text-emerald-700" />
                          <span>Mevcut Drive Dosyasından Yeniden Oku</span>
                        </button>
                        <button
                          type="button"
                          onClick={() => setIsDataSyncModalOpen(false)}
                          className="px-4 py-2 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold text-xs cursor-pointer"
                        >
                          Kapat
                        </button>
                      </div>
                    </div>
                  )}
                </div>
              )}
            </div>
          </div>
        </div>
      )}

      {/* YENİ SAYFA EKLE MODAL (EXCEL MANTIĞI) */}
      {isAddingSheet && (
        <div className="fixed inset-0 z-[10002] flex items-center justify-center p-4 bg-slate-900/70 backdrop-blur-sm animate-fade-in">
          <div className="bg-white text-slate-800 w-full max-w-xl rounded-2xl shadow-2xl border border-slate-200 overflow-hidden flex flex-col">
            <div className="px-6 py-4 bg-white border-b border-slate-200 flex items-center justify-between">
              <h4 className="text-sm font-black text-[#0b3d1d] uppercase tracking-wider flex items-center gap-2">
                <Plus className="w-4 h-4 text-[#0b3d1d]" />
                <span>YENİ SAYFA OLUŞTUR (Excel Mantığı)</span>
              </h4>
              <button
                type="button"
                onClick={() => setIsAddingSheet(false)}
                className="text-slate-400 hover:text-slate-700 text-sm p-1 rounded-lg"
              >
                ✕
              </button>
            </div>
            <div className="p-6 space-y-4">
              <div className="flex flex-col gap-1.5">
                <label className="text-xs font-bold text-slate-700 uppercase font-mono">
                  Sayfa Adı (Excel Sekme Adı) *
                </label>
                <input
                  type="text"
                  autoFocus
                  placeholder="Örn: 2026 Arıza ve İncelemeler, Özel Bakım Takip..."
                  value={newSheetName}
                  onChange={(e) => setNewSheetName(e.target.value)}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-xl text-xs text-slate-900 focus:outline-none focus:border-[#0b3d1d] focus:bg-white font-medium"
                />
              </div>

              <div className="flex flex-col gap-1.5">
                <label className="text-xs font-bold text-slate-700 uppercase font-mono">
                  Sayfa Başlığı (Çizelge Üst Başlığı - İsteğe Bağlı)
                </label>
                <input
                  type="text"
                  placeholder="Örn: AT-802 YILLIK ARIZA VE İNCELEME ÇİZELGESİ"
                  value={newSheetTitle}
                  onChange={(e) => setNewSheetTitle(e.target.value)}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-xl text-xs text-slate-900 focus:outline-none focus:border-[#0b3d1d] focus:bg-white font-medium"
                />
              </div>

              <div className="flex flex-col gap-1.5">
                <div className="flex items-center justify-between">
                  <label className="text-xs font-bold text-slate-700 uppercase font-mono">
                    Sütun Başlıkları (Virgülle veya Satırla Ayırınız) *
                  </label>
                  <span className="text-[10px] text-emerald-800 font-semibold font-mono">
                    (Belge / Döküman otomatik eklenir)
                  </span>
                </div>
                <textarea
                  rows={4}
                  value={newSheetColumnsText}
                  onChange={(e) => setNewSheetColumnsText(e.target.value)}
                  placeholder="Sıra, Tarih, Kuyruk No, Olay Açıklaması, Sorumlu Personel, Hava Aracı / Parça Durumu, Belge / Döküman"
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-xl text-xs text-slate-900 focus:outline-none focus:border-[#0b3d1d] focus:bg-white"
                />
              </div>

              {/* Hızlı Şablonlar */}
              <div className="space-y-1.5">
                <span className="text-[11px] font-bold text-slate-500 uppercase tracking-wider font-mono">
                  Hızlı Şablon Seç:
                </span>
                <div className="flex flex-wrap gap-1.5">
                  <button
                    type="button"
                    onClick={() =>
                      setNewSheetColumnsText(
                        'Sıra, Tarih, Kuyruk No, Olay / Kaza-Kırım Tipi, Konum, Sorumlu Pilot, Hasar Açıklaması, Hava Aracı / Parça Durumu, Belge / Döküman'
                      )
                    }
                    className="px-2.5 py-1 rounded-md bg-slate-100 hover:bg-slate-200 text-slate-700 text-[10px] font-bold transition-all"
                  >
                    Olay & Kaza-Kırım Şablonu
                  </button>
                  <button
                    type="button"
                    onClick={() =>
                      setNewSheetColumnsText(
                        'Sıra, Tarih, Kuyruk No, Motor Seri No, Tespit Edilen En Yüksek Ng (%), DAA Referans No, Hava Aracı / Parça Durumu, Belge / Döküman'
                      )
                    }
                    className="px-2.5 py-1 rounded-md bg-slate-100 hover:bg-slate-200 text-slate-700 text-[10px] font-bold transition-all"
                  >
                    Motor & Ng Limit Şablonu
                  </button>
                  <button
                    type="button"
                    onClick={() =>
                      setNewSheetColumnsText(
                        'Sıra, Tarih, Kuyruk No, Arıza / İşlem Tanımı, İşlemi Yapan Ekip, Hava Aracı / Parça Durumu, Belge / Döküman'
                      )
                    }
                    className="px-2.5 py-1 rounded-md bg-slate-100 hover:bg-slate-200 text-slate-700 text-[10px] font-bold transition-all"
                  >
                    Genel Bakım & İşlem Şablonu
                  </button>
                </div>
              </div>
            </div>

            <div className="px-6 py-3 bg-slate-50 border-t border-slate-200 flex items-center justify-end gap-2">
              <button
                type="button"
                onClick={() => setIsAddingSheet(false)}
                className="px-4 py-2 rounded-xl bg-slate-200 hover:bg-slate-300 text-slate-700 font-bold text-xs cursor-pointer"
              >
                İptal
              </button>
              <button
                type="button"
                onClick={handleCreateNewSheet}
                className="px-5 py-2 rounded-xl bg-[#0b3d1d] hover:bg-[#072612] text-white font-black text-xs uppercase tracking-wider cursor-pointer shadow-sm"
              >
                Sayfayı Oluştur & Kaydet
              </button>
            </div>
          </div>
        </div>
      )}

      {/* BELGE / DÖKÜMAN EKLEME & YÖNETİM MODAL (HANGAR YER DESTEK MANTIĞI) */}
      {activeDocRow && currentSheet && (
        <div className="fixed inset-0 z-[10003] flex items-center justify-center p-4 bg-slate-900/70 backdrop-blur-sm animate-fade-in">
          <div className="bg-white text-slate-800 w-full max-w-xl rounded-2xl shadow-2xl border border-slate-200 overflow-hidden flex flex-col max-h-[90vh]">
            <div className="px-6 py-4 bg-white border-b border-slate-200 flex items-center justify-between">
              <div className="flex items-center gap-3">
                <div className="flex items-center gap-2">
                  <Paperclip className="w-4 h-4 text-[#0b3d1d]" />
                  <h4 className="text-sm font-black text-[#0b3d1d] uppercase tracking-wider">
                    BELGE & DÖKÜMAN YÖNETİMİ (Satır #{activeDocRow.rowIndex + 1})
                  </h4>
                </div>
                <button
                  type="button"
                  onClick={async () => {
                    setIsSyncingDocs(true);
                    try {
                      const docs = await syncHangarPdfDocsFromDrive(GOOGLE_SCRIPT_URL);
                      if (Array.isArray(docs)) {
                        setHangarPdfDocs(docs);
                        showNotification(`Google Drive'dan ${docs.length} adet belge senkronize edildi.`, 'success');
                      }
                    } catch (e: any) {
                      showNotification('Drive senkronizasyon hatası: ' + (e?.message || e), 'error');
                    } finally {
                      setIsSyncingDocs(false);
                    }
                  }}
                  disabled={isSyncingDocs}
                  className="px-2.5 py-1 rounded-lg bg-emerald-100 hover:bg-emerald-200 text-emerald-900 font-bold text-[11px] transition-all flex items-center gap-1.5 cursor-pointer disabled:opacity-50"
                  title="Google Drive'daki tüm belgeleri tara ve listeyi yenile"
                >
                  <RefreshCw className={`w-3 h-3 ${isSyncingDocs ? 'animate-spin' : ''}`} />
                  <span>{isSyncingDocs ? 'Çekiliyor...' : "Drive'dan Yenile"}</span>
                </button>
              </div>
              <button
                type="button"
                onClick={() => setActiveDocRow(null)}
                className="text-slate-400 hover:text-slate-700 text-sm p-1 rounded-lg"
              >
                ✕
              </button>
            </div>

            <div className="p-6 overflow-y-auto space-y-5 flex-1">
              {/* Row Summary */}
              <div className="bg-slate-50 p-3 rounded-xl border border-slate-200 text-xs">
                <div className="font-bold text-slate-800 mb-1">Kayıt Özeti:</div>
                <div className="text-slate-600 font-mono text-[11px] leading-relaxed">
                  {activeDocRow.row.slice(0, 4).filter(Boolean).join(' • ')}
                </div>
              </div>

              {/* Existing Documents List */}
              <div className="space-y-2">
                <label className="text-xs font-bold text-slate-700 uppercase font-mono">
                  Ekli Belgeler ({getRowDocs(activeDocRow.rowIndex, activeDocRow.row[belgeColIndex] || '').length})
                </label>
                {getRowDocs(activeDocRow.rowIndex, activeDocRow.row[belgeColIndex] || '').length === 0 ? (
                  <div className="p-4 border border-dashed border-slate-200 rounded-xl text-center text-xs text-slate-400">
                    Bu satıra ait henüz eklenmiş bir PDF belge bulunmuyor.
                  </div>
                ) : (
                  <div className="space-y-2">
                    {getRowDocs(activeDocRow.rowIndex, activeDocRow.row[belgeColIndex] || '').map(doc => {
                      const cat = getFileCategory(doc.fileName);
                      const isImg = cat === 'image';
                      const isArch = cat === 'archive';
                      const ext = doc.fileName?.split('.').pop()?.toUpperCase() || 'BELGE';

                      return (
                        <div
                          key={doc.id}
                          className="flex items-center justify-between p-3 rounded-xl bg-slate-50 hover:bg-emerald-50/40 border border-slate-200 transition-colors"
                        >
                          <div className="flex items-center gap-2.5 min-w-0">
                            <div className={`w-8 h-8 rounded-lg flex items-center justify-center shrink-0 ${
                              isImg ? 'bg-blue-100 text-blue-800' :
                              isArch ? 'bg-indigo-100 text-indigo-800' :
                              'bg-emerald-100 text-emerald-800'
                            }`}>
                              {isImg ? <ImageIcon className="w-4 h-4" /> : isArch ? <Archive className="w-4 h-4" /> : <FileText className="w-4 h-4" />}
                            </div>
                            <div className="min-w-0">
                              <div className="flex items-center gap-1.5">
                                <div className="font-bold text-xs text-slate-800 truncate" title={doc.fileName}>
                                  {doc.fileName}
                                </div>
                                <span className={`text-[9px] font-mono px-1 py-0.2 rounded font-bold uppercase shrink-0 ${
                                  isImg ? 'bg-blue-100 text-blue-800 border border-blue-200' :
                                  isArch ? 'bg-indigo-100 text-indigo-800 border border-indigo-200' :
                                  'bg-emerald-100 text-emerald-800 border border-emerald-200'
                                }`}>
                                  {ext}
                                </span>
                              </div>
                              <div className="text-[10px] text-slate-500 font-mono">
                                {doc.docType} • {doc.fileSize || ext} • {doc.uploadDate || doc.uploadedAt}
                              </div>
                            </div>
                          </div>

                          <div className="flex items-center gap-1.5 shrink-0 ml-2">
                            <button
                              type="button"
                              onClick={() => handleOpenDocInNewTab(doc)}
                              className="px-2 py-1.5 rounded-lg bg-blue-100 hover:bg-blue-200 text-blue-900 transition-all cursor-pointer flex items-center gap-1 text-[10px] font-bold"
                              title="Yeni Sekmede Aç"
                            >
                              <ExternalLink className="w-3.5 h-3.5 text-blue-700" />
                              <span>Yeni Sekmede Aç</span>
                            </button>
                            <button
                              type="button"
                              onClick={() => setPreviewDoc(doc)}
                              className="p-1.5 rounded-lg bg-emerald-100 hover:bg-emerald-200 text-emerald-900 transition-all cursor-pointer"
                              title="Görüntüle / İncele"
                            >
                              <Eye className="w-3.5 h-3.5" />
                            </button>
                            <a
                              href={doc.fileData && doc.fileData.length > 500 && !doc.fileData.includes('...[IDB]') ? doc.fileData : (doc.driveUrl || `https://drive.google.com/uc?export=download&id=${doc.driveFileId}`)}
                              download={doc.fileName}
                              target="_blank"
                              rel="noreferrer"
                              className="p-1.5 rounded-lg bg-slate-100 hover:bg-slate-200 text-slate-700 transition-all cursor-pointer"
                              title="İndir"
                            >
                              <Download className="w-3.5 h-3.5" />
                            </a>
                            <button
                              type="button"
                              onClick={() => handleDeleteDoc(doc.id, doc.fileName)}
                              className="p-1.5 rounded-lg bg-slate-100 hover:bg-rose-100 text-slate-600 hover:text-rose-700 transition-all cursor-pointer"
                              title="Sil (Drive'dan Çöpe Taşır)"
                            >
                              <Trash2 className="w-3.5 h-3.5" />
                            </button>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>

              {/* Document Management Tabs: Upload / Drive / Link */}
              <div className="pt-3 border-t border-slate-200 space-y-3">
                <div className="flex items-center justify-between">
                  <div className="flex items-center gap-1 p-1 bg-slate-100 rounded-xl border border-slate-200 w-full sm:w-auto">
                    <button
                      type="button"
                      onClick={() => setDocModalTab('upload')}
                      className={`flex-1 sm:flex-initial px-3 py-1.5 rounded-lg text-xs font-bold transition-all flex items-center justify-center gap-1.5 cursor-pointer ${
                        docModalTab === 'upload'
                          ? 'bg-white text-emerald-900 shadow-xs border border-slate-200/80 font-black'
                          : 'text-slate-600 hover:text-slate-900'
                      }`}
                    >
                      <UploadCloud className="w-3.5 h-3.5 text-emerald-700" />
                      <span>Belge / Dosya Yükle</span>
                    </button>
                    <button
                      type="button"
                      onClick={() => setDocModalTab('drive')}
                      className={`flex-1 sm:flex-initial px-3 py-1.5 rounded-lg text-xs font-bold transition-all flex items-center justify-center gap-1.5 cursor-pointer ${
                        docModalTab === 'drive'
                          ? 'bg-white text-emerald-900 shadow-xs border border-slate-200/80 font-black'
                          : 'text-slate-600 hover:text-slate-900'
                      }`}
                    >
                      <Search className="w-3.5 h-3.5 text-blue-600" />
                      <span>Drive'dan Seç & Bağla</span>
                    </button>
                    <button
                      type="button"
                      onClick={() => setDocModalTab('link')}
                      className={`flex-1 sm:flex-initial px-3 py-1.5 rounded-lg text-xs font-bold transition-all flex items-center justify-center gap-1.5 cursor-pointer ${
                        docModalTab === 'link'
                          ? 'bg-white text-blue-900 shadow-xs border border-slate-200/80 font-black'
                          : 'text-slate-600 hover:text-slate-900'
                      }`}
                    >
                      <LinkIcon className="w-3.5 h-3.5 text-blue-600" />
                      <span>Bağlantı Linki Ekle</span>
                    </button>
                  </div>
                  <span className="hidden sm:inline-block text-[10px] text-slate-400 font-mono">
                    Drive: {hangarPdfDocs.length} belge
                  </span>
                </div>

                {/* TAB 1: UPLOAD LOCAL FILE */}
                {docModalTab === 'upload' && (
                  <div className="space-y-3">
                    <input
                      ref={docFileInputRef}
                      type="file"
                      accept=".pdf,application/pdf,image/jpeg,image/png,image/jpg,.jpg,.jpeg,.png,.zip,application/zip,application/x-zip-compressed,.rar,application/x-rar-compressed,application/octet-stream"
                      className="hidden"
                      onChange={(e) => {
                        const file = e.target.files?.[0];
                        if (file) {
                          setDocUploadFile(file);
                          if (!docUploadName.trim()) {
                            const cleanBase = file.name.replace(/\.[^/.]+$/, '');
                            setDocUploadName(cleanBase);
                          }
                        }
                      }}
                    />

                    {!docUploadFile ? (
                      <div
                        onClick={() => docFileInputRef.current?.click()}
                        className="border-2 border-dashed border-emerald-300 hover:border-emerald-600 bg-emerald-50/20 hover:bg-emerald-50/50 rounded-xl p-6 flex flex-col items-center justify-center text-center cursor-pointer transition-all"
                      >
                        <UploadCloud className="w-9 h-9 text-emerald-800 mb-2" />
                        <span className="text-xs font-bold text-slate-800">
                          Yeni Belge / Döküman Seç (PDF, JPEG, ZIP, RAR)
                        </span>
                        <span className="text-[10px] text-slate-500 mt-1 max-w-sm">
                          Kaza-kırım raporu, fotoğraflar, teknik evraklar vb. Seçtiğinizde dosya otomatik Google Drive'a aktarılır ve satırın "{EK_DOSYA_LINK_COL_NAME}" sütununa kaydedilir.
                        </span>
                      </div>
                    ) : (
                      <div className="space-y-3 bg-slate-50 p-4 rounded-xl border border-slate-200">
                        <div className="flex items-center justify-between text-xs">
                          <div className="flex items-center gap-2">
                            {(() => {
                              const c = getFileCategory(docUploadFile.name);
                              if (c === 'image') return <ImageIcon className="w-4 h-4 text-blue-600" />;
                              if (c === 'archive') return <Archive className="w-4 h-4 text-indigo-600" />;
                              return <FileText className="w-4 h-4 text-emerald-800" />;
                            })()}
                            <span className="font-bold text-slate-800">{docUploadFile.name}</span>
                            <span className="text-[10px] text-slate-400 font-mono">
                              ({(docUploadFile.size / 1024).toFixed(1)} KB)
                            </span>
                          </div>
                          <button
                            type="button"
                            onClick={() => {
                              setDocUploadFile(null);
                              setDocUploadName('');
                            }}
                            className="text-xs text-rose-600 hover:underline cursor-pointer"
                          >
                            Değiştir
                          </button>
                        </div>

                        <div className="flex flex-col gap-1">
                          <label className="text-xs font-bold text-slate-700 uppercase font-mono">
                            Döküman Adı (Kullanıcıya Görünecek İsim) *
                          </label>
                          <input
                            type="text"
                            value={docUploadName}
                            onChange={(e) => setDocUploadName(e.target.value)}
                            placeholder="Örn: Kaza Raporu - OGM-01"
                            className="w-full px-3 py-1.5 bg-white border border-slate-300 rounded-lg text-xs text-slate-900 focus:outline-none focus:border-[#0b3d1d] font-medium"
                          />
                        </div>

                        <div className="flex flex-col gap-1">
                          <label className="text-xs font-bold text-slate-700 uppercase font-mono">
                            Belge Kategorisi / Tipi
                          </label>
                          <select
                            value={docUploadCategory}
                            onChange={(e) => setDocUploadCategory(e.target.value)}
                            className="w-full px-3 py-1.5 bg-white border border-slate-300 rounded-lg text-xs text-slate-900 focus:outline-none focus:border-[#0b3d1d] font-medium"
                          >
                            <option value="Kaza/Kırım & Olay Belgesi">Kaza/Kırım & Olay Belgesi</option>
                            <option value="Motor Ng Limit Aşım Raporu">Motor Ng Limit Aşım Raporu</option>
                            <option value="DAA İzin & Değerlendirme Belgesi">DAA İzin & Değerlendirme Belgesi</option>
                            <option value="Servis Bülteni & Teknik Yayın">Servis Bülteni & Teknik Yayın</option>
                            <option value="Tutanak & Fotoğraf">Tutanak & Fotoğraf</option>
                            <option value="Diğer Resmi Yazı & Ekler">Diğer Resmi Yazı & Ekler</option>
                          </select>
                        </div>

                        <button
                          type="button"
                          disabled={isDocUploading}
                          onClick={handleSaveDocToRow}
                          className="w-full py-2.5 rounded-xl bg-[#0b3d1d] hover:bg-[#072612] text-white font-black text-xs uppercase tracking-wider cursor-pointer shadow-sm flex items-center justify-center gap-2 disabled:opacity-50"
                        >
                          {isDocUploading ? (
                            <>
                              <RotateCw className="w-3.5 h-3.5 animate-spin" />
                              <span>Drive'a Yükleniyor & Satıra Bağlanıyor...</span>
                            </>
                          ) : (
                            <>
                              <Check className="w-3.5 h-3.5" />
                              <span>Drive'a Yükle, Kaydet & Satıra Bağla</span>
                            </>
                          )}
                        </button>
                      </div>
                    )}
                  </div>
                )}

                {/* TAB 2: SEARCH & LINK FROM GOOGLE DRIVE */}
                {docModalTab === 'drive' && (
                  <div className="bg-slate-50 p-4 rounded-xl border border-slate-200 space-y-3">
                    <div className="flex items-center justify-between">
                      <span className="text-[11px] font-bold text-slate-700 uppercase font-mono flex items-center gap-1">
                        <Search className="w-3 h-3 text-emerald-800" />
                        <span>Google Drive'da Yüklü Belgeleri Ara & Satıra Bağla</span>
                      </span>
                      <button
                        type="button"
                        onClick={async () => {
                          setIsSyncingDocs(true);
                          try {
                            const docs = await syncHangarPdfDocsFromDrive(GOOGLE_SCRIPT_URL);
                            if (Array.isArray(docs)) {
                              setHangarPdfDocs(docs);
                              showNotification(`Drive'dan ${docs.length} adet güncel belge getirildi.`, 'success');
                            }
                          } catch (err: any) {
                            showNotification('Hata: ' + (err?.message || err), 'error');
                          } finally {
                            setIsSyncingDocs(false);
                          }
                        }}
                        disabled={isSyncingDocs}
                        className="text-[10px] text-emerald-800 hover:text-emerald-950 font-bold underline flex items-center gap-1 cursor-pointer disabled:opacity-50"
                      >
                        <RefreshCw className={`w-2.5 h-2.5 ${isSyncingDocs ? 'animate-spin' : ''}`} />
                        <span>{isSyncingDocs ? 'Taranıyor...' : 'Drive Listesini Yenile'}</span>
                      </button>
                    </div>

                    <div className="relative">
                      <Search className="w-3.5 h-3.5 absolute left-2.5 top-1/2 -translate-y-1/2 text-slate-400" />
                      <input
                        type="text"
                        value={driveDocSearchTerm}
                        onChange={(e) => setDriveDocSearchTerm(e.target.value)}
                        placeholder="Belge adı, kuyruk no veya dosya adıyla Drive'da ara..."
                        className="w-full pl-8 pr-3 py-2 bg-white border border-slate-300 rounded-lg text-xs text-slate-800 placeholder:text-slate-400 focus:outline-none focus:border-[#0b3d1d] font-medium"
                      />
                    </div>

                    <div className="max-h-56 overflow-y-auto space-y-1.5 pt-1">
                      {hangarPdfDocs
                        .filter(d => {
                          if (!driveDocSearchTerm.trim()) return true;
                          const term = driveDocSearchTerm.toLowerCase();
                          return (
                            (d.fileName && d.fileName.toLowerCase().includes(term)) ||
                            (d.docType && d.docType.toLowerCase().includes(term)) ||
                            (d.itemKey && d.itemKey.toLowerCase().includes(term))
                          );
                        })
                        .slice(0, 20)
                        .map(driveDoc => (
                          <div
                            key={driveDoc.id}
                            className="flex items-center justify-between p-2.5 rounded-lg bg-white border border-slate-200 hover:border-emerald-500 transition-colors text-xs"
                          >
                            <div className="flex items-center gap-2 min-w-0">
                              {(() => {
                                const c = getFileCategory(driveDoc.fileName);
                                if (c === 'image') return <ImageIcon className="w-3.5 h-3.5 text-blue-600 shrink-0" />;
                                if (c === 'archive') return <Archive className="w-3.5 h-3.5 text-indigo-600 shrink-0" />;
                                return <FileText className="w-3.5 h-3.5 text-emerald-700 shrink-0" />;
                              })()}
                              <div className="min-w-0">
                                <div className="font-bold text-slate-800 truncate text-[11px]" title={driveDoc.fileName}>
                                  {driveDoc.fileName}
                                </div>
                                <div className="text-[9px] text-slate-400 font-mono">
                                  {driveDoc.docType || 'Belge'} • {driveDoc.fileSize || ''} • {driveDoc.uploadDate || ''}
                                </div>
                              </div>
                            </div>
                            <div className="flex items-center gap-1.5 shrink-0 ml-2">
                              <button
                                type="button"
                                onClick={() => handleOpenDocInNewTab(driveDoc)}
                                className="px-2 py-1 rounded bg-blue-50 hover:bg-blue-100 text-blue-700 text-[10px] font-bold cursor-pointer flex items-center gap-1 border border-blue-200"
                                title="Yeni Sekmede Aç"
                              >
                                <ExternalLink className="w-3 h-3" />
                                <span>Aç</span>
                              </button>
                              <button
                                type="button"
                                onClick={() => setPreviewDoc(driveDoc)}
                                className="px-2 py-1 rounded bg-slate-100 hover:bg-slate-200 text-slate-700 text-[10px] font-bold cursor-pointer"
                                title="Önizle"
                              >
                                Gör
                              </button>
                              <button
                                type="button"
                                onClick={() => handleAttachExistingDriveDoc(driveDoc)}
                                className="px-2.5 py-1 rounded bg-emerald-700 hover:bg-emerald-800 text-white text-[10px] font-black tracking-wide flex items-center gap-1 cursor-pointer active:scale-95 shadow-2xs"
                              >
                                <Plus className="w-3 h-3" />
                                <span>Satıra Bağla</span>
                              </button>
                            </div>
                          </div>
                        ))}
                      {hangarPdfDocs.length === 0 && (
                        <div className="text-center py-4 text-xs text-slate-400">
                          Drive'da kayıtlı belge listesi boş veya taranmadı. Yukarıdaki "Drive Listesini Yenile" butonuna basarak belgeleri çekebilirsiniz.
                        </div>
                      )}
                    </div>
                  </div>
                )}

                {/* TAB 3: CUSTOM LINK (URL / DRIVE LINK) */}
                {docModalTab === 'link' && (
                  <div className="space-y-3 bg-blue-50/50 p-4 rounded-xl border border-blue-200">
                    <div className="flex items-start gap-2 bg-blue-100/60 p-2.5 rounded-lg text-blue-950 text-xs">
                      <LinkIcon className="w-4 h-4 text-blue-700 shrink-0 mt-0.5" />
                      <div>
                        <span className="font-bold">Otomatik Excel & Drive Sütun Entegrasyonu:</span>
                        <p className="text-[11px] text-blue-900 mt-0.5 leading-relaxed">
                          Buraya girdiğiniz bağlantı linki, Excel tablosunda <strong>"{EK_DOSYA_LINK_COL_NAME}"</strong> sütununa otomatik kaydedilir ve Google Drive'a senkronize edilir. Tabloda veya belgede <em>"Yeni Sekmede Aç"</em> butonuna tıklandığında bu bağlantı doğrudan yeni sekmede açılır.
                        </p>
                      </div>
                    </div>

                    {/* Current Link If Exists */}
                    {ekDosyaLinkColIndex >= 0 && activeDocRow.row[ekDosyaLinkColIndex] && (
                      <div className="p-3 bg-white rounded-lg border border-blue-200 space-y-1.5">
                        <div className="text-[10px] font-bold uppercase text-slate-500 font-mono">
                          Mevcut Kayıtlı Bağlantı Linki:
                        </div>
                        <div className="flex items-center justify-between gap-2">
                          <a
                            href={activeDocRow.row[ekDosyaLinkColIndex]}
                            target="_blank"
                            rel="noopener noreferrer"
                            className="text-xs font-semibold text-blue-700 hover:text-blue-900 underline truncate flex items-center gap-1"
                          >
                            <ExternalLink className="w-3 h-3 shrink-0" />
                            <span className="truncate">{activeDocRow.row[ekDosyaLinkColIndex]}</span>
                          </a>
                          <div className="flex items-center gap-1 shrink-0">
                            <button
                              type="button"
                              onClick={() => handleOpenDocInNewTab(null, activeDocRow.row[ekDosyaLinkColIndex])}
                              className="px-2 py-1 rounded bg-blue-100 hover:bg-blue-200 text-blue-900 text-[10px] font-bold cursor-pointer flex items-center gap-1"
                            >
                              <ExternalLink className="w-3 h-3" />
                              <span>Yeni Sekmede Aç</span>
                            </button>
                            <button
                              type="button"
                              onClick={handleRemoveCustomLink}
                              className="px-2 py-1 rounded bg-rose-50 hover:bg-rose-100 text-rose-700 text-[10px] font-bold cursor-pointer"
                            >
                              Kaldır
                            </button>
                          </div>
                        </div>
                      </div>
                    )}

                    <div className="space-y-1">
                      <label className="text-xs font-bold text-slate-700 uppercase font-mono flex items-center gap-1">
                        <LinkIcon className="w-3 h-3 text-blue-600" />
                        <span>Bağlantı Linki (URL) *</span>
                      </label>
                      <input
                        type="url"
                        value={customLinkUrl}
                        onChange={(e) => setCustomLinkUrl(e.target.value)}
                        placeholder="https://drive.google.com/... veya https://..."
                        className="w-full px-3 py-2 bg-white border border-slate-300 rounded-lg text-xs text-slate-900 focus:outline-none focus:border-blue-600 font-mono"
                      />
                    </div>

                    <div className="space-y-1">
                      <label className="text-xs font-bold text-slate-700 uppercase font-mono">
                        Bağlantı Başlığı / Döküman Adı (İsteğe Bağlı)
                      </label>
                      <input
                        type="text"
                        value={customLinkTitle}
                        onChange={(e) => setCustomLinkTitle(e.target.value)}
                        placeholder="Örn: OGM-01 Olay Dosyası / Teknik İnceleme Linki"
                        className="w-full px-3 py-1.5 bg-white border border-slate-300 rounded-lg text-xs text-slate-900 focus:outline-none focus:border-blue-600"
                      />
                    </div>

                    <button
                      type="button"
                      onClick={handleSaveCustomLink}
                      disabled={!customLinkUrl.trim()}
                      className="w-full py-2.5 rounded-xl bg-blue-700 hover:bg-blue-800 text-white font-black text-xs uppercase tracking-wider cursor-pointer shadow-sm flex items-center justify-center gap-2 disabled:opacity-50 transition-all"
                    >
                      <Check className="w-3.5 h-3.5" />
                      <span>Bağlantı Linkini Satıra Ekle & Excel'e Kaydet</span>
                    </button>
                  </div>
                )}
              </div>
            </div>

            <div className="px-6 py-3 bg-slate-50 border-t border-slate-200 flex items-center justify-end">
              <button
                type="button"
                onClick={() => setActiveDocRow(null)}
                className="px-4 py-2 rounded-xl bg-slate-200 hover:bg-slate-300 text-slate-700 font-bold text-xs cursor-pointer"
              >
                Kapat
              </button>
            </div>
          </div>
        </div>
      )}

      {/* EDIT ROW MODAL */}
      {editingRow && currentSheet && (
        <div className="fixed inset-0 z-[10002] flex items-center justify-center p-4 bg-slate-900/70 backdrop-blur-sm animate-fade-in">
          <div className="bg-white text-slate-800 w-full max-w-2xl rounded-2xl shadow-2xl border border-slate-200 overflow-hidden flex flex-col max-h-[90vh]">
            <div className="px-6 py-4 bg-white border-b border-slate-200 flex items-center justify-between">
              <h4 className="text-sm font-black text-amber-700 uppercase tracking-wider flex items-center gap-2">
                <Edit3 className="w-4 h-4 text-amber-700" />
                <span>KAYIT DÜZENLE (Satır #{editingRow.rowIndex + 1})</span>
              </h4>
              <button
                type="button"
                onClick={() => setEditingRow(null)}
                className="text-slate-400 hover:text-slate-700 text-sm p-1 rounded-lg"
              >
                ✕
              </button>
            </div>
            <div className="p-6 overflow-y-auto space-y-4 flex-1">
              {currentSheet.columns.map((col, idx) => {
                const lowerCol = col.toLowerCase();
                const isDurum =
                  lowerCol.includes('durum') ||
                  lowerCol.includes('durumu') ||
                  lowerCol.includes('sonuç') ||
                  lowerCol.includes('karar');

                return (
                  <div key={idx} className="flex flex-col gap-1.5">
                    <div className="flex items-center justify-between">
                      <label className="text-xs font-bold text-slate-700 uppercase font-mono">
                        {col}
                      </label>
                      {isDurum && (
                        <span className="text-[10px] text-slate-400 font-mono">
                          Hızlı Seçim Yapabilirsiniz:
                        </span>
                      )}
                    </div>

                    {isDurum && (
                      <div className="flex flex-wrap gap-1.5 mb-1">
                        {[
                          { label: 'Faal', cls: 'bg-emerald-600 text-white' },
                          { label: 'Onarıma Alındı', cls: 'bg-yellow-400 text-slate-950' },
                          { label: 'Bakım/Overhaul Yapıldı', cls: 'bg-orange-500 text-white' },
                          { label: 'Ağır Hasarlı', cls: 'bg-rose-600 text-white' }
                        ].map(st => (
                          <button
                            key={st.label}
                            type="button"
                            onClick={() => {
                              const newR = [...editingRow.row];
                              newR[idx] = st.label;
                              setEditingRow({ ...editingRow, row: newR });
                            }}
                            className={`px-2.5 py-1 rounded-md text-[11px] font-bold cursor-pointer ${st.cls}`}
                          >
                            {st.label}
                          </button>
                        ))}
                      </div>
                    )}

                    {col.toLowerCase().includes('detay') ||
                    col.toLowerCase().includes('sonuç') ||
                    col.toLowerCase().includes('aksiyon') ||
                    col.toLowerCase().includes('açıklama') ? (
                      <textarea
                        rows={3}
                        value={editingRow.row[idx] || ''}
                        onChange={(e) => {
                          const newR = [...editingRow.row];
                          newR[idx] = e.target.value;
                          setEditingRow({ ...editingRow, row: newR });
                        }}
                        className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-xl text-xs text-slate-900 focus:outline-none focus:border-[#0b3d1d] focus:bg-white"
                      />
                    ) : (
                      <input
                        type="text"
                        value={
                          /tar[ıi]h|date/i.test(col) && /^\d{5}(\.\d+)?$/.test(editingRow.row[idx] || '')
                            ? cleanAndFormatDateString(editingRow.row[idx])
                            : (editingRow.row[idx] || '')
                        }
                        placeholder={/tar[ıi]h|date/i.test(col) ? 'GG.AA.YYYY (Örn: 15.07.2026)' : ''}
                        onChange={(e) => {
                          const newR = [...editingRow.row];
                          newR[idx] = e.target.value;
                          setEditingRow({ ...editingRow, row: newR });
                        }}
                        className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-xl text-xs text-slate-900 focus:outline-none focus:border-[#0b3d1d] focus:bg-white font-medium"
                      />
                    )}
                  </div>
                );
              })}
            </div>
            <div className="px-6 py-3 bg-slate-50 border-t border-slate-200 flex items-center justify-end gap-2">
              <button
                type="button"
                onClick={() => setEditingRow(null)}
                className="px-4 py-2 rounded-xl bg-slate-200 hover:bg-slate-300 text-slate-700 font-bold text-xs cursor-pointer"
              >
                İptal
              </button>
              <button
                type="button"
                onClick={() => handleSaveRow(editingRow.rowIndex, editingRow.row)}
                className="px-5 py-2 rounded-xl bg-[#0b3d1d] hover:bg-[#072612] text-white font-black text-xs uppercase tracking-wider cursor-pointer shadow-sm"
              >
                Kaydet & Senkronize Et
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ADD NEW ROW MODAL */}
      {isAddingRow && currentSheet && (
        <div className="fixed inset-0 z-[10002] flex items-center justify-center p-4 bg-slate-900/70 backdrop-blur-sm animate-fade-in">
          <div className="bg-white text-slate-800 w-full max-w-2xl rounded-2xl shadow-2xl border border-slate-200 overflow-hidden flex flex-col max-h-[90vh]">
            <div className="px-6 py-4 bg-white border-b border-slate-200 flex items-center justify-between">
              <h4 className="text-sm font-black text-[#0b3d1d] uppercase tracking-wider flex items-center gap-2">
                <Plus className="w-4 h-4 text-[#0b3d1d]" />
                <span>YENİ SATIR EKLE ({currentSheet.sheetName})</span>
              </h4>
              <button
                type="button"
                onClick={() => setIsAddingRow(false)}
                className="text-slate-400 hover:text-slate-700 text-sm p-1 rounded-lg"
              >
                ✕
              </button>
            </div>
            <div className="p-6 overflow-y-auto space-y-4 flex-1">
              {currentSheet.columns.map((col, idx) => {
                const lowerCol = col.toLowerCase();
                const isDurum =
                  lowerCol.includes('durum') ||
                  lowerCol.includes('durumu') ||
                  lowerCol.includes('sonuç') ||
                  lowerCol.includes('karar');

                return (
                  <div key={idx} className="flex flex-col gap-1.5">
                    <div className="flex items-center justify-between">
                      <label className="text-xs font-bold text-slate-700 uppercase font-mono">
                        {col}
                      </label>
                      {isDurum && (
                        <span className="text-[10px] text-slate-400 font-mono">
                          Hızlı Seçim:
                        </span>
                      )}
                    </div>

                    {isDurum && (
                      <div className="flex flex-wrap gap-1.5 mb-1">
                        {[
                          { label: 'Faal', cls: 'bg-emerald-600 text-white' },
                          { label: 'Onarıma Alındı', cls: 'bg-yellow-400 text-slate-950' },
                          { label: 'Bakım/Overhaul Yapıldı', cls: 'bg-orange-500 text-white' },
                          { label: 'Ağır Hasarlı', cls: 'bg-rose-600 text-white' }
                        ].map(st => (
                          <button
                            key={st.label}
                            type="button"
                            onClick={() => {
                              const newR = [...newRowValues];
                              newR[idx] = st.label;
                              setNewRowValues(newR);
                            }}
                            className={`px-2.5 py-1 rounded-md text-[11px] font-bold cursor-pointer ${st.cls}`}
                          >
                            {st.label}
                          </button>
                        ))}
                      </div>
                    )}

                    {col.toLowerCase().includes('detay') ||
                    col.toLowerCase().includes('sonuç') ||
                    col.toLowerCase().includes('aksiyon') ||
                    col.toLowerCase().includes('açıklama') ? (
                      <textarea
                        rows={3}
                        value={newRowValues[idx] || ''}
                        onChange={(e) => {
                          const newR = [...newRowValues];
                          newR[idx] = e.target.value;
                          setNewRowValues(newR);
                        }}
                        placeholder={`${col} giriniz...`}
                        className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-xl text-xs text-slate-900 focus:outline-none focus:border-[#0b3d1d] focus:bg-white"
                      />
                    ) : (
                      <input
                        type="text"
                        value={newRowValues[idx] || ''}
                        onChange={(e) => {
                          const newR = [...newRowValues];
                          newR[idx] = e.target.value;
                          setNewRowValues(newR);
                        }}
                        placeholder={/tar[ıi]h|date/i.test(col) ? 'GG.AA.YYYY (Örn: 15.07.2026)' : `${col} giriniz...`}
                        className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-xl text-xs text-slate-900 focus:outline-none focus:border-[#0b3d1d] focus:bg-white font-medium"
                      />
                    )}
                  </div>
                );
              })}
            </div>
            <div className="px-6 py-3 bg-slate-50 border-t border-slate-200 flex items-center justify-end gap-2">
              <button
                type="button"
                onClick={() => setIsAddingRow(false)}
                className="px-4 py-2 rounded-xl bg-slate-200 hover:bg-slate-300 text-slate-700 font-bold text-xs cursor-pointer"
              >
                İptal
              </button>
              <button
                type="button"
                onClick={handleSaveNewRow}
                className="px-5 py-2 rounded-xl bg-[#0b3d1d] hover:bg-[#072612] text-white font-black text-xs uppercase tracking-wider cursor-pointer shadow-sm"
              >
                Satırı Ekle & Drive'a Kaydet
              </button>
            </div>
          </div>
        </div>
      )}

      {/* PASSWORD PROTECTION MODAL (FOR DELETING) */}
      {passwordModalOpen && (
        <div className="fixed inset-0 z-[10004] flex items-center justify-center p-4 bg-slate-900/70 backdrop-blur-sm animate-fade-in">
          <div className="bg-white text-slate-800 w-full max-w-sm rounded-2xl shadow-2xl border border-slate-200 overflow-hidden">
            <div className="px-6 py-4 bg-white border-b border-slate-200 flex items-center justify-between">
              <h4 className="text-sm font-black text-slate-800 uppercase tracking-wider flex items-center gap-2">
                <Lock className="w-4 h-4 text-[#0b3d1d]" />
                <span>YETKİ ONAYI</span>
              </h4>
              <button
                type="button"
                onClick={() => setPasswordModalOpen(false)}
                className="text-slate-400 hover:text-slate-700 text-sm"
              >
                ✕
              </button>
            </div>
            <form onSubmit={handleVerifyPassword} className="p-6 space-y-4">
              <p className="text-xs text-slate-600 leading-relaxed">
                Bu işlemi gerçekleştirmek için yetkili şifresini giriniz.
              </p>
              <div className="space-y-1.5">
                <input
                  type="password"
                  autoFocus
                  placeholder="Yetkili şifresi..."
                  value={passwordInput}
                  onChange={(e) => {
                    setPasswordInput(e.target.value);
                    setPasswordError('');
                  }}
                  className="w-full px-4 py-2 bg-slate-50 border border-slate-300 rounded-xl text-xs text-slate-900 font-mono tracking-widest text-center focus:outline-none focus:border-[#0b3d1d] focus:bg-white"
                />
                {passwordError && (
                  <p className="text-xs text-rose-600 font-bold text-center">{passwordError}</p>
                )}
              </div>
              <div className="flex gap-2 pt-2">
                <button
                  type="button"
                  onClick={() => setPasswordModalOpen(false)}
                  className="flex-1 py-2 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold text-xs cursor-pointer"
                >
                  İptal
                </button>
                <button
                  type="submit"
                  className="flex-1 py-2 rounded-xl bg-[#0b3d1d] hover:bg-[#072612] text-white font-black text-xs uppercase tracking-wider cursor-pointer shadow-sm"
                >
                  Onayla
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* PDF PREVIEW MODAL */}
      {previewDoc && (
        <div className="fixed inset-0 z-[10010]">
          <PdfPreviewModal preview={previewDoc} onClose={() => setPreviewDoc(null)} />
        </div>
      )}
    </div>
  );
};
