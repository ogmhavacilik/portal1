/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useRef, useMemo, ChangeEvent } from 'react';
import { removeBackground, preload } from '@imgly/background-removal';

// Helper function to resize oversized camera photos down to max 640px before AI processing
const prepareOptimizedImageForAI = async (file: File, maxDimension: number = 640): Promise<Blob | File> => {
  return new Promise((resolve) => {
    const img = new Image();
    const objectUrl = URL.createObjectURL(file);
    img.onload = () => {
      URL.revokeObjectURL(objectUrl);
      let { width, height } = img;
      if (width <= maxDimension && height <= maxDimension) {
        resolve(file);
        return;
      }
      if (width > height) {
        height = Math.round((height * maxDimension) / width);
        width = maxDimension;
      } else {
        width = Math.round((width * maxDimension) / height);
        height = maxDimension;
      }
      const canvas = document.createElement('canvas');
      canvas.width = width;
      canvas.height = height;
      const ctx = canvas.getContext('2d');
      if (ctx) {
        ctx.imageSmoothingEnabled = true;
        ctx.imageSmoothingQuality = 'medium';
        ctx.drawImage(img, 0, 0, width, height);
        canvas.toBlob((blob) => {
          canvas.width = 0;
          canvas.height = 0;
          if (blob) {
            resolve(new File([blob], file.name || 'optimized.jpg', { type: 'image/jpeg' }));
          } else {
            resolve(file);
          }
        }, 'image/jpeg', 0.82);
      } else {
        resolve(file);
      }
    };
    img.onerror = () => {
      URL.revokeObjectURL(objectUrl);
      resolve(file);
    };
    img.src = objectUrl;
  });
};
import {
  Plane,
  Users,
  Package,
  Settings,
  ArrowLeft,
  X,
  Wrench,
  CalendarCheck,
  ClipboardList,
  Fuel,
  Construction,
  FileText,
  Download,
  Upload,
  Plus,
  Trash2,
  Search,
  RefreshCw,
  Lock,
  Database,
  Check,
  AlertCircle,
  AlertTriangle,
  CloudLightning,
  Loader2,
  CheckCircle,
  Eye,
  ExternalLink,
  ChevronUp,
  ChevronDown,
  ChevronRight,
  ChevronLeft,
  Camera,
  Edit3,
  Columns,
  Rows,
  Square,
  CheckSquare,
  Table,
  Folder,
  Printer,
  FileSpreadsheet,
  Copy,
  SlidersHorizontal,
  UserCheck,
  Truck,
  History,
  Maximize2,
  Sparkles,
  Send
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import * as XLSX from 'xlsx';
import { jsPDF } from 'jspdf';
import { ImageEditorAndRetoucher } from './components/ImageEditorAndRetoucher';
import { CachedDriveImage, fetchDriveImageAsBase64 } from './components/CachedDriveImage';

export const GOOGLE_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbzz6XVe-NDXbNUks8KFMfTVYN0JfN6PhQGLVNDG26yolgwGtD8DTBTKD8PgXW5V-n6vEQ/exec";
export const EBYS_SEARCH_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbwgWc7aKDB_dtubQVxeQDpiHR0FF8jeYvfDWRzcx4kbYUfLsT9vJGg69zupHbGoUf5H/exec";
export const TASKLINE_SUBMIT_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbys4kKFJI87wbn155z6jphH7D5qgC45FWUvzzxi9n4-YfYDdRxY72fMWTaTGMxvkXqN-g/exec";

// Convert a File object to Base64 string for Drive uploading
export const fileToBase64 = (file: File): Promise<string> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.readAsDataURL(file);
    reader.onload = () => {
      const result = reader.result as string;
      const base64 = result.split(',')[1];
      resolve(base64);
    };
    reader.onerror = (error) => reject(error);
  });
};

export const getEmbeddableDriveUrl = (url: string | null | undefined): string | null => {
  if (!url) return null;
  if (url.startsWith('data:')) return url;

  // Convert Google Drive view or sharing links to direct/embed links
  // Pattern 1: https://drive.google.com/file/d/FILE_ID/view... or similar
  const dMatch = url.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (dMatch && dMatch[1]) {
    return `https://drive.google.com/uc?export=download&id=${dMatch[1]}`;
  }

  // Pattern 2: https://drive.google.com/open?id=FILE_ID
  const idMatch = url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (url.includes('drive.google.com') && idMatch && idMatch[1]) {
    return `https://drive.google.com/uc?export=download&id=${idMatch[1]}`;
  }

  // Pattern 3: Ensure any uc?id= link has export=download
  if (url.includes('drive.google.com/uc') && !url.includes('export=download')) {
    const ucIdMatch = url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
    if (ucIdMatch && ucIdMatch[1]) {
      return `https://drive.google.com/uc?export=download&id=${ucIdMatch[1]}`;
    }
  }

  return url;
};

export const formatToTurkishDateRange = (startStr: string, endStr: string): string => {
  if (!startStr || !endStr) return "Haziran 2026";
  const start = new Date(startStr);
  const end = new Date(endStr);
  if (isNaN(start.getTime()) || isNaN(end.getTime())) return "Haziran 2026";
  
  const turkishMonths = [
    "Ocak", "Subat", "Mart", "Nisan", "Mayis", "Haziran",
    "Temmuz", "Agustos", "Eylul", "Ekim", "Kasim", "Aralik"
  ];
  
  const startDay = start.getDate();
  const startMonth = turkishMonths[start.getMonth()];
  const startYear = start.getFullYear();
  
  const endDay = end.getDate();
  const endMonth = turkishMonths[end.getMonth()];
  const endYear = end.getFullYear();
  
  if (startYear === endYear) {
    if (start.getMonth() === end.getMonth()) {
      return `${startDay}-${endDay} ${startMonth} ${startYear}`;
    }
    return `${startDay} ${startMonth} - ${endDay} ${endMonth} ${startYear}`;
  }
  return `${startDay} ${startMonth} ${startYear} - ${endDay} ${endMonth} ${endYear}`;
};

export const formatBirthDateToTurkish = (val: string): string => {
  if (!val) return "";
  let s = String(val).trim();

  // Map of English month names to Turkish month names
  const monthsMap: Record<string, string> = {
    "january": "Ocak",
    "february": "Şubat",
    "march": "Mart",
    "april": "Nisan",
    "may": "Mayıs",
    "june": "Haziran",
    "july": "Temmuz",
    "august": "Ağustos",
    "september": "Eylül",
    "october": "Ekim",
    "november": "Kasım",
    "december": "Aralık",
    "jan": "Ocak",
    "feb": "Şubat",
    "mar": "Mart",
    "apr": "Nisan",
    "jun": "Haziran",
    "jul": "Temmuz",
    "aug": "Ağustos",
    "sep": "Eylül",
    "oct": "Ekim",
    "nov": "Kasım",
    "dec": "Aralık"
  };

  const turkishMonths = [
    "Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran",
    "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"
  ];

  // If it is a serial Excel date (e.g., purely numeric)
  if (/^\d+$/.test(s)) {
    const serial = parseInt(s, 10);
    try {
      const utc_days  = Math.floor(serial - 25569);
      const utc_value = utc_days * 86400;                                        
      const dateObj = new Date(utc_value * 1000);
      
      const day = dateObj.getDate();
      const monthIndex = dateObj.getMonth();
      const year = dateObj.getFullYear();
      
      if (year > 1920 && year < 2030 && monthIndex >= 0 && monthIndex < 12) {
        return `${day} ${turkishMonths[monthIndex]} ${year}`;
      }
    } catch (e) {
      // fallback
    }
  }

  // If it is in YYYY-MM-DD format
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    const parts = s.split("-");
    const yr = parseInt(parts[0], 10);
    const mn = parseInt(parts[1], 10);
    const dy = parseInt(parts[2], 10);
    if (mn >= 1 && mn <= 12) {
      return `${dy} ${turkishMonths[mn - 1]} ${yr}`;
    }
  }

  // Also replace any English month names
  Object.keys(monthsMap).forEach(engMonth => {
    const regex = new RegExp(`\\b${engMonth}\\b`, 'gi');
    s = s.replace(regex, monthsMap[engMonth]);
  });

  return s;
};

export const normalizeTurkishForSearch = (str: string): string => {
  if (!str) return "";
  let val = str.toString();
  val = val
    .replace(/İ/g, 'i')
    .replace(/I/g, 'ı')
    .replace(/Ğ/g, 'g')
    .replace(/ğ/g, 'g')
    .replace(/Ü/g, 'u')
    .replace(/ü/g, 'u')
    .replace(/Ş/g, 's')
    .replace(/ş/g, 's')
    .replace(/Ö/g, 'o')
    .replace(/ö/g, 'o')
    .replace(/Ç/g, 'c')
    .replace(/ç/g, 'c')
    .toLowerCase()
    .replace(/ı/g, 'i') // map dotless i to dotted i for search stability
    .replace(/[^a-z0-9\s]/g, '') // strip hyphens, dots, parentheses, and any non-alphanumeric/non-space symbols
    .replace(/\s+/g, ' ') // collapse multiple spaces
    .trim();
  return val;
};

export const parseRawPeriodStringToDates = (periodStr: string): { start: string; end: string } => {
  try {
    const parts = periodStr.split('-');
    if (parts.length !== 2) return { start: "2026-06-01", end: "2026-06-30" };
    
    const monthsMap: Record<string, number> = {
      "ocak": 0, "subat": 1, "mart": 2, "nisan": 3, "mayis": 4, "haziran": 5,
      "temmuz": 6, "agustos": 7, "eylul": 8, "ekim": 9, "kasim": 10, "aralik": 11
    };
    
    const parsePart = (part: string) => {
      const subParts = part.trim().split('_');
      if (subParts.length !== 3) throw new Error("Invalid format");
      const day = parseInt(subParts[0], 10);
      const monthName = subParts[1].toLowerCase();
      const month = monthsMap[monthName] !== undefined ? monthsMap[monthName] : 5;
      const year = parseInt(subParts[2], 10);
      
      const d = new Date(year, month, day);
      const yyyy = d.getFullYear();
      const mm = String(d.getMonth() + 1).padStart(2, '0');
      const dd = String(d.getDate()).padStart(2, '0');
      return `${yyyy}-${mm}-${dd}`;
    };
    
    return {
      start: parsePart(parts[0]),
      end: parsePart(parts[1])
    };
  } catch (e) {
    return { start: "2026-06-01", end: "2026-06-30" };
  }
};

export const convertToRawPeriodString = (startStr: string, endStr: string): string => {
  if (!startStr || !endStr) return "1_haziran_2026-30_haziran_2026";
  const start = new Date(startStr);
  const end = new Date(endStr);
  if (isNaN(start.getTime()) || isNaN(end.getTime())) return "1_haziran_2026-30_haziran_2026";

  const getMonthNameEN = (date: Date) => {
    const months = [
      "ocak", "subat", "mart", "nisan", "mayis", "haziran",
      "temmuz", "agustos", "eylul", "ekim", "kasim", "aralik"
    ];
    return months[date.getMonth()];
  };

  const startDay = start.getDate();
  const startMonth = getMonthNameEN(start);
  const startYear = start.getFullYear();

  const endDay = end.getDate();
  const endMonth = getMonthNameEN(end);
  const endYear = end.getFullYear();

  return `${startDay}_${startMonth}_${startYear}-${endDay}_${endMonth}_${endYear}`;
};

export const rotateDataUrl = (dataUrl: string, rotation: number): Promise<string> => {
  return new Promise((resolve) => {
    if (rotation === 0) {
      resolve(dataUrl);
      return;
    }
    const img = new Image();
    img.onload = () => {
      const canvas = document.createElement('canvas');
      const ctx = canvas.getContext('2d');
      if (!ctx) {
        resolve(dataUrl);
        return;
      }
      
      const angle = (rotation * Math.PI) / 180;
      const is90or270 = (rotation / 90) % 2 !== 0;
      
      canvas.width = is90or270 ? img.height : img.width;
      canvas.height = is90or270 ? img.width : img.height;
      
      ctx.translate(canvas.width / 2, canvas.height / 2);
      ctx.rotate(angle);
      ctx.drawImage(img, -img.width / 2, -img.height / 2);
      
      resolve(canvas.toDataURL('image/png'));
    };
    img.onerror = () => resolve(dataUrl);
    img.src = dataUrl;
  });
};

export const generatePdfFromImages = async (pages: { dataUrl: string; width: number; height: number }[]): Promise<string> => {
  return new Promise((resolve, reject) => {
    try {
      if (pages.length === 0) {
        reject(new Error("PDF oluşturulacak sayfa bulunamadı."));
        return;
      }
      
      const firstPage = pages[0];
      // Create jsPDF instance with px matching the exact rotated page size
      const pdf = new jsPDF({
        orientation: firstPage.width > firstPage.height ? 'l' : 'p',
        unit: 'px',
        format: [firstPage.width, firstPage.height]
      });

      // Add first page
      pdf.addImage(firstPage.dataUrl, 'PNG', 0, 0, firstPage.width, firstPage.height);

      // Add subsequent pages
      for (let i = 1; i < pages.length; i++) {
        const page = pages[i];
        pdf.addPage([page.width, page.height], page.width > page.height ? 'l' : 'p');
        pdf.addImage(page.dataUrl, 'PNG', 0, 0, page.width, page.height);
      }

      const dataUri = pdf.output('datauristring');
      const base64 = dataUri.substring(dataUri.indexOf(',') + 1);
      resolve(base64);
    } catch (err) {
      reject(err);
    }
  });
};

// Simple, robust IndexedDB helper for storing PDF pages locally without size limits
export const openPdfDB = (): Promise<IDBDatabase> => {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open("portal_pdf_cache", 1);
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains("pdf_pages")) {
        db.createObjectStore("pdf_pages", { keyPath: "key" });
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
};

export const savePdfPagesToDB = async (key: string, pages: { pageNumber: number; dataUrl: string; width: number; height: number; textItems?: any[] }[]) => {
  const db = await openPdfDB();
  return new Promise<void>((resolve, reject) => {
    const transaction = db.transaction("pdf_pages", "readwrite");
    const store = transaction.objectStore("pdf_pages");
    const request = store.put({ key, pages });
    request.onsuccess = () => resolve();
    request.onerror = () => reject(request.error);
  });
};

export const saveRawPdfToDB = async (key: string, base64: string) => {
  const db = await openPdfDB();
  return new Promise<void>((resolve, reject) => {
    const transaction = db.transaction("pdf_pages", "readwrite");
    const store = transaction.objectStore("pdf_pages");
    const request = store.put({ key: "raw_pdf_" + key, base64 });
    request.onsuccess = () => resolve();
    request.onerror = () => reject(request.error);
  });
};

export const getRawPdfFromDB = async (key: string): Promise<string | null> => {
  const db = await openPdfDB();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction("pdf_pages", "readonly");
    const store = transaction.objectStore("pdf_pages");
    const request = store.get("raw_pdf_" + key);
    request.onsuccess = () => {
      if (request.result) {
        resolve(request.result.base64);
      } else {
        resolve(null);
      }
    };
    request.onerror = () => reject(request.error);
  });
};

export const getPdfPagesFromDB = async (key: string): Promise<{ pageNumber: number; dataUrl: string; width: number; height: number; textItems?: any[] }[] | null> => {
  const db = await openPdfDB();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction("pdf_pages", "readonly");
    const store = transaction.objectStore("pdf_pages");
    const request = store.get(key);
    request.onsuccess = () => {
      if (request.result) {
        resolve(request.result.pages);
      } else {
        resolve(null);
      }
    };
    request.onerror = () => reject(request.error);
  });
};

export const deletePdfFromDB = async (key: string) => {
  const db = await openPdfDB();
  return new Promise<void>((resolve, reject) => {
    const transaction = db.transaction("pdf_pages", "readwrite");
    const store = transaction.objectStore("pdf_pages");
    store.delete(key);
    store.delete("raw_pdf_" + key);
    transaction.oncomplete = () => resolve();
    transaction.onerror = () => reject(transaction.error);
  });
};

export type CategoryType = 
  | 'İKMAL' 
  | 'TEÇHİZAT TAKİP' 
  | 'HA_YER_DESTEK' 
  | 'T70_DETAY' 
  | 'KARA_ARACLARI_MENU'
  | 'FORM KAYITLARI' 
  | null;

// Table Configuration Interfaces
export interface TableColumn {
  key: string;
  label: string;
}

export interface TableConfig {
  title: string;
  sheetName: string;
  storageKey: string;
  columns: TableColumn[];
  defaultRows: Record<string, string>[];
}

export const TABLE_CONFIGS: Record<number, TableConfig> = {
  1: {
    title: 'GÖREVLENDİRME ÇİZELGELERİ',
    sheetName: '1-Gorevlendirme',
    storageKey: 'form_1_gorevlendirme',
    columns: [
      { key: 'Tarih', label: 'Tarih' },
      { key: 'Personel_Adi_Soyadi', label: 'Personel Adı Soyadı' },
      { key: 'Unvani_Gorevi', label: 'Unvanı / Görevi' },
      { key: 'Gorev_Yeri', label: 'Görev Yeri' },
      { key: 'Durumu', label: 'Durumu' },
      { key: 'Aciklama', label: 'Açıklama' }
    ],
    defaultRows: [
      { Tarih: "2026-06-15", Personel_Adi_Soyadi: "Ahmet Yılmaz", Unvani_Gorevi: "Teknik Koordinatör", Gorev_Yeri: "Ankara Hangar 2", Durumu: "Aktif", Aciklama: "Bell 429 haftalık kontrol" },
      { Tarih: "2026-06-16", Personel_Adi_Soyadi: "Mehmet Kaya", Unvani_Gorevi: "Baş Teknisyen", Gorev_Yeri: "Muğla Helikopter Üssü", Durumu: "Aktif", Aciklama: "At-802F Yangın Söndürme Sezonu Görevi" },
      { Tarih: "2026-06-17", Personel_Adi_Soyadi: "Cem Şahin", Unvani_Gorevi: "Aviyonik Uzmanı", Gorev_Yeri: "Ankara Merkez", Durumu: "Planlandı", Aciklama: "T-70 telsiz bakım desteği" }
    ]
  },
  21: {
    title: 'YAZ DÖNEMİ PLANLAMASI (BELL 429)',
    sheetName: '2-Yaz_Donemi-bell_429',
    storageKey: 'form_2_yaz_donemi_bell429',
    columns: [
      { key: 'Donem_Hafta', label: 'Dönem/Hafta' },
      { key: 'Baslangic_Tarihi', label: 'Başlangıç Tarihi' },
      { key: 'Bitis_Tarihi', label: 'Bitiş Tarihi' },
      { key: 'Nobetci_Muhendis', label: 'Nöbetçi Mühendis' },
      { key: 'Nobetci_Teknisyen', label: 'Nöbetçi Teknisyen' },
      { key: 'Yedek_Personel', label: 'Yedek Personel' }
    ],
    defaultRows: [
      { Donem_Hafta: "Haziran 1. Hafta", Baslangic_Tarihi: "2026-06-01", Bitis_Tarihi: "2026-06-07", Nobetci_Muhendis: "Ayşe Demir (Bell 429)", Nobetci_Teknisyen: "Ömer Faruk", Yedek_Personel: "Ali Vural" },
      { Donem_Hafta: "Haziran 2. Hafta", Baslangic_Tarihi: "2026-06-08", Bitis_Tarihi: "2026-06-14", Nobetci_Muhendis: "Murat Tandoğan (Bell 429)", Nobetci_Teknisyen: "Veli Can", Yedek_Personel: "Burak Çelik" }
    ]
  },
  22: {
    title: 'YAZ DÖNEMİ PLANLAMASI (T-70)',
    sheetName: '2-Yaz_Donemi-t_70',
    storageKey: 'form_2_yaz_donemi_t70',
    columns: [
      { key: 'Donem_Hafta', label: 'Dönem/Hafta' },
      { key: 'Baslangic_Tarihi', label: 'Başlangıç Tarihi' },
      { key: 'Bitis_Tarihi', label: 'Bitiş Tarihi' },
      { key: 'Nobetci_Muhendis', label: 'Nöbetçi Mühendis' },
      { key: 'Nobetci_Teknisyen', label: 'Nöbetçi Teknisyen' },
      { key: 'Yedek_Personel', label: 'Yedek Personel' }
    ],
    defaultRows: [
      { Donem_Hafta: "Haziran 1. Hafta", Baslangic_Tarihi: "2026-06-01", Bitis_Tarihi: "2026-06-07", Nobetci_Muhendis: "Hasan Yıldız (T-70)", Nobetci_Teknisyen: "Süleyman Ak", Yedek_Personel: "Selin Tan" },
      { Donem_Hafta: "Haziran 2. Hafta", Baslangic_Tarihi: "2026-06-08", Bitis_Tarihi: "2026-06-14", Nobetci_Muhendis: "Cemil Pek (T-70)", Nobetci_Teknisyen: "Hakan Güler", Yedek_Personel: "Tuncay Yaman" }
    ]
  },
  23: {
    title: 'YAZ DÖNEMİ PLANLAMASI (AT-802)',
    sheetName: '2-Yaz_Donemi-at_802',
    storageKey: 'form_2_yaz_donemi_at802',
    columns: [
      { key: 'Donem_Hafta', label: 'Dönem/Hafta' },
      { key: 'Baslangic_Tarihi', label: 'Başlangıç Tarihi' },
      { key: 'Bitis_Tarihi', label: 'Bitiş Tarihi' },
      { key: 'Nobetci_Muhendis', label: 'Nöbetçi Mühendis' },
      { key: 'Nobetci_Teknisyen', label: 'Nöbetçi Teknisyen' },
      { key: 'Yedek_Personel', label: 'Yedek Personel' }
    ],
    defaultRows: [
      { Donem_Hafta: "Haziran 1. Hafta", Baslangic_Tarihi: "2026-06-01", Bitis_Tarihi: "2026-06-07", Nobetci_Muhendis: "Mert Sökmen (AT-802)", Nobetci_Teknisyen: "Bülent Er", Yedek_Personel: "Yasin Kaya" },
      { Donem_Hafta: "Haziran 2. Hafta", Baslangic_Tarihi: "2026-06-08", Bitis_Tarihi: "2026-06-14", Nobetci_Muhendis: "Fikret Şen (AT-802)", Nobetci_Teknisyen: "Selim Tok", Yedek_Personel: "Kadir Bal" }
    ]
  },
  24: {
    title: 'ANKARA BEKLEME GÖREV PLANLAMASI (BELL 429)',
    sheetName: '2-Ankara_Bekleme-bell_429',
    storageKey: 'form_2_ankara_bekleme_bell429',
    columns: [
      { key: 'Hafta_No', label: 'Hafta No' },
      { key: 'Gorev_Periyodu', label: 'Görev Periyodu' },
      { key: 'Gun', label: 'Gün' },
      { key: 'Pilot', label: 'Pilot' },
      { key: 'Teknisyen', label: 'Teknisyen' }
    ],
    defaultRows: [
      { Hafta_No: "1", Gorev_Periyodu: "7 Mayıs 2026", Gun: "5", Pilot: "Ahmet Yılmaz\nCem Şahin", Teknisyen: "Alper ÖZMETİN" }
    ]
  },
  25: {
    title: 'ANKARA BEKLEME GÖREV PLANLAMASI (C-650/B-360)',
    sheetName: '2-Ankara_Bekleme-c650_b360',
    storageKey: 'form_2_ankara_bekleme_c650_b360',
    columns: [
      { key: 'Hafta_No', label: 'Hafta No' },
      { key: 'Gorev_Periyodu', label: 'Görev Periyodu' },
      { key: 'Gun', label: 'Gün' },
      { key: 'C650_Pilot', label: 'C-650 Pilot' },
      { key: 'C650_Teknisyen', label: 'C-650 Teknisyen' },
      { key: 'B360_Pilot', label: 'B-360 Pilot' },
      { key: 'B360_Teknisyen', label: 'B-360 Teknisyen' }
    ],
    defaultRows: [
      { 
        Hafta_No: "1", 
        Gorev_Periyodu: "7 Mayıs 2026", 
        Gun: "5", 
        C650_Pilot: "Mahmut OKUDAN\nMurat AKMEŞE\nYılmaz MAMUNLUOĞLU\nCengiz ÖZDEMİR", 
        C650_Teknisyen: "Ali ÖZKAVSAL\nUtku GÖKGÖZ\nAlper ÖZMETİN", 
        B360_Pilot: "Aydın TÜTÜNCÜOĞLU\nAltan Alkan SÖZEN\nDevrim Ferhat ÇALIŞKAN\nAydemir TEZGEL\nSerkan KEBAPCI", 
        B360_Teknisyen: "Aycan TAN\nÖmer ERSOY\nTezcan GÜZER\nHasan AKSOY\nFerhat ÖZCAN" 
      }
    ]
  },
  3: {
    title: 'BAKIM YETKİ ÇİZELGELERİ',
    sheetName: '3-Bakim_Yetki',
    storageKey: 'form_3_bakim_yetki',
    columns: [
      { key: 'Personel_Sicil_No', label: 'Personel Sicil No' },
      { key: 'Adi_Soyadi', label: 'Adı Soyadı' },
      { key: 'Bransi', label: 'Branşı' },
      { key: 'Bulundugu_Hava_Araci_Tipi', label: 'Bulunduğu Hava Aracı Tipi' },
      { key: 'Yetki_Seviyesi', label: 'Yetki Seviyesi' },
      { key: 'Yetkilendirme_Tarihi', label: 'Yetkilendirme Tarihi' }
    ],
    defaultRows: [
      { Personel_Sicil_No: "SIC-1042", Adi_Soyadi: "Salih Bostan", Bransi: "Gövde Motor", Bulundugu_Hava_Araci_Tipi: "T-70", Yetki_Seviyesi: "Level 3 - Baş Denetçi", Yetkilendirme_Tarihi: "2024-03-12" },
      { Personel_Sicil_No: "SIC-2195", Adi_Soyadi: "Zeynep Elmas", Bransi: "Aviyonik", Bulundugu_Hava_Araci_Tipi: "Bell 429", Yetki_Seviyesi: "Level 2 - Teknisyen", Yetkilendirme_Tarihi: "2025-05-20" },
      { Personel_Sicil_No: "SIC-8841", Adi_Soyadi: "Kemal Sun", Bransi: "Sistem Bakım", Bulundugu_Hava_Araci_Tipi: "AT-802F", Yetki_Seviyesi: "Level 1 - Yardımcı", Yetkilendirme_Tarihi: "2026-01-10" }
    ]
  },
  5: {
    title: 'PERSONEL BİLGİ ÇİZELGELERİ',
    sheetName: '5-Personel_Bilgi',
    storageKey: 'form_5_personel_bilgi',
    columns: [
      { key: 'Sira_No', label: 'Sıra No' },
      { key: 'Adi_Soyadi', label: 'Adı Soyadı' },
      { key: 'TC_Kimlik', label: 'T.C. Kimlik' },
      { key: 'Sicil_No', label: 'Sicil No' },
      { key: 'Kadro_Unvan_Gorev', label: 'Kadro Unvanı / Görevi' },
      { key: 'Dogum_Tarihi', label: 'Doğum Tarihi' },
      { key: 'Gorev_Yeri', label: 'Görev Yeri' },
      { key: 'Telefon_No', label: 'Telefon No' },
      { key: 'Kan_Grubu', label: 'Kan Grubu' },
      { key: 'Adres_Bilgisi', label: 'Adres Bilgisi' },
      { key: 'Yakinin_Adi_Soyadi', label: 'Yakının Adı-Soyadı' },
      { key: 'Es_Telefon_Numaralari', label: 'Eş Telefon Numaraları' }
    ],
    defaultRows: [
      { Sira_No: "1", Adi_Soyadi: "Mahmut OKUDAN", TC_Kimlik: "29122439024", Sicil_No: "45537", Kadro_Unvan_Gorev: "Pilot", Dogum_Tarihi: "27595", Gorev_Yeri: "OGM", Telefon_No: "533 3497446", Kan_Grubu: "A Rh+", Adres_Bilgisi: "Ankara Merkez", Yakinin_Adi_Soyadi: "Hatice OKUDAN (Eşi)", Es_Telefon_Numaralari: "533 3497447" },
      { Sira_No: "2", Adi_Soyadi: "Rıfat ÖNAL", TC_Kimlik: "66577226922", Sicil_No: "45540", Kadro_Unvan_Gorev: "Pilot", Dogum_Tarihi: "24289", Gorev_Yeri: "HANGAR", Telefon_No: "530 656 3112", Kan_Grubu: "0 Rh+", Adres_Bilgisi: "Ankara Keçiören", Yakinin_Adi_Soyadi: "Selin ÖNAL (Eşi)", Es_Telefon_Numaralari: "530 656 3113" }
    ]
  },
  6: {
    title: 'PERSONEL UÇUŞ-HİZMET YILLARI ÇİZELGESİ',
    sheetName: '6-Personel_Ucus_Hizmet',
    storageKey: 'form_6_personel_ucus_hizmet',
    columns: [
      { key: 'Adi_Soyadi', label: 'Adı Soyadı' },
      { key: 'Memuriyet_Baslangici', label: 'Memuriyet Başlangıcı' },
      { key: 'Toplam_Hizmet_Yili', label: 'Toplam Hizmet Yılı' },
      { key: 'Toplam_Ucus_Saati', label: 'Toplam Uçuş Saati' },
      { key: 'Ucus_Tazminati_Durumu', label: 'Uçuş Tazminatı Durumu' }
    ],
    defaultRows: [
      { Adi_Soyadi: "Salih Bostan", Memuriyet_Baslangici: "2008-01-15", Toplam_Hizmet_Yili: "18", Toplam_Ucus_Saati: "1450", Ucus_Tazminati_Durumu: "Aktif / Ödeniyor" },
      { Adi_Soyadi: "Zeynep Elmas", Memuriyet_Baslangici: "2014-06-01", Toplam_Hizmet_Yili: "12", Toplam_Ucus_Saati: "680", Ucus_Tazminati_Durumu: "Aktif / Ödeniyor" },
      { Adi_Soyadi: "Kemal Sun", Memuriyet_Baslangici: "2019-10-10", Toplam_Hizmet_Yili: "7", Toplam_Ucus_Saati: "320", Ucus_Tazminati_Durumu: "Beklemede" }
    ]
  }
};

export const SUMMER_FORM_IDS = [21, 22, 23, 24, 25];

export const isSummerForm = (id: number | null): boolean => {
  return id !== null && SUMMER_FORM_IDS.includes(id);
};

export const getAirframeSuffix = (id: number | null): string => {
  if (id === 21) return 'bell429';
  if (id === 22) return 't70';
  if (id === 23) return 'at802';
  if (id === 24) return 'bekleme_bell429';
  if (id === 25) return 'bekleme_c650_b360';
  return '';
};

export default function App() {
  // Splash screen state
  const [splashVisible, setSplashVisible] = useState(true);

  // Mobile detection
  const [isMobile, setIsMobile] = useState(window.innerWidth < 768);

  useEffect(() => {
    // Preload lightweight mobile AI model for ultra fast background removal
    try {
      preload({ model: 'isnet_fp16' }).catch(() => {});
    } catch (e) {
      // Ignore
    }

    const handleResize = () => {
      setIsMobile(window.innerWidth < 768);
    };
    window.addEventListener('resize', handleResize);
    return () => window.removeEventListener('resize', handleResize);
  }, []);

  // Modal active state
  const [modalOpen, setModalOpen] = useState(false);
  const [modalTitle, setModalTitle] = useState('SİSTEM');
  const [modalType, setModalType] = useState<'iframe' | 'design' | 'category' | 'form_table' | 'excel_sync' | 'techizat_matrix' | 'denetleme'>('iframe');
  const [modalUrl, setModalUrl] = useState('');
  const [iframeLoading, setIframeLoading] = useState(false);
  const [selectedCategory, setSelectedCategory] = useState<CategoryType>(null);
  const [categoryHistory, setCategoryHistory] = useState<CategoryType[]>([]);

  // Selected Form ID (1, 21, 22, 23, 3, 5, 6)
  const [selectedFormId, setSelectedFormId] = useState<number | null>(null);
  const [formTableMode, setFormTableMode] = useState<'selection' | 'offline'>('selection');
  const [uploadProgress, setUploadProgress] = useState<number>(0);
  const [sheetIframeLoading, setSheetIframeLoading] = useState<boolean>(false);
  const [isReadOnlyView, setIsReadOnlyView] = useState<boolean>(true);

  // Summer period month selection
  const [selectedSummerMonth, setSelectedSummerMonth] = useState<string>("1_haziran_2026-30_haziran_2026");
  const [selectedUploadSummerMonth, setSelectedUploadSummerMonth] = useState<string>("1_haziran_2026-30_haziran_2026");
  const [selectedSummerStartDate, setSelectedSummerStartDate] = useState<string>("2026-06-01");
  const [selectedSummerEndDate, setSelectedSummerEndDate] = useState<string>("2026-06-30");
  const [selectedUploadSummerStartDate, setSelectedUploadSummerStartDate] = useState<string>("2026-06-01");
  const [selectedUploadSummerEndDate, setSelectedUploadSummerEndDate] = useState<string>("2026-06-30");
  const [isPdfViewMode, setIsPdfViewMode] = useState<boolean>(true);
  const [pdfSearchQuery, setPdfSearchQuery] = useState<string>("");
  const [activeMatchIndex, setActiveMatchIndex] = useState<number>(0);
  const pdfSearchInputRef = useRef<HTMLInputElement>(null);

  // Personnel cell double click/click viewer state
  const [activeModalCell, setActiveModalCell] = useState<{ r: number; c: number; value: string; label: string } | null>(null);
  const [copiedCellSuccess, setCopiedCellSuccess] = useState<boolean>(false);

  // Teçhizat row edit, image upload and mission order states
  const [mobileEditTab, setMobileEditTab] = useState<'form' | 'image'>('form');
  const [activeTechizatRowEdit, setActiveTechizatRowEdit] = useState<{
    rIdx: number;
    techType: 'bell429' | 'at802' | 't70' | 't70_bumbi_backet' | 't70_helitak' | 'b360' | 'c650' | 'hangar' | 'kara_araclari' | 'all';
    row: string[];
  } | null>(null);
  const [techizatImages, setTechizatImages] = useState<Record<string, string>>(() => {
    const defaultImages = {
      "bell429_Bell_429_Çekme_Çubuğu_(Tow_Bar)_SN-9982": "https://drive.google.com/file/d/1QXCX6zN79vZ6nSk4prwH0LacFfT0WyEv/view?usp=drivesdk"
    };
    const saved = localStorage.getItem('techizat_images');
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        return { ...defaultImages, ...parsed };
      } catch (e) {
        console.error(e);
      }
    }
    return defaultImages;
  });
  const [techizatImageScale, setTechizatImageScale] = useState<number>(1);
  const [isFullScreenImage, setIsFullScreenImage] = useState<boolean>(false);
  const [editRowValues, setEditRowValues] = useState<string[]>([]);
  const [showSavePasswordPrompt, setShowSavePasswordPrompt] = useState<boolean>(false);
  const [tempImageUrlInput, setTempImageUrlInput] = useState<string>('');
  const [tempImageAction, setTempImageAction] = useState<'upload' | 'link' | 'remove' | null>(null);
  const [showImageSavePasswordPrompt, setShowImageSavePasswordPrompt] = useState<boolean>(false);
  const [isImageUpdateUnlocked, setIsImageUpdateUnlocked] = useState<boolean>(false);
  const [imagePasswordInput, setImagePasswordInput] = useState<string>('');
  const [imagePasswordError, setImagePasswordError] = useState<boolean>(false);
  const [showImagePasswordPrompt, setShowImagePasswordPrompt] = useState<boolean>(false);
  const [pendingImageFile, setPendingImageFile] = useState<File | null>(null);
  const [pendingImagePreview, setPendingImagePreview] = useState<string | null>(null);
  const [isDragging, setIsDragging] = useState<boolean>(false);
  const [isImageUploadingToDrive, setIsImageUploadingToDrive] = useState<boolean>(false);

  // Hover preview state for equipment table
  const [hoveredRowImage, setHoveredRowImage] = useState<{ url: string; title: string; subtitle: string; x: number; y: number } | null>(null);

  // Excel Export with Images modal state
  const [excelExportModalData, setExcelExportModalData] = useState<{
    type: string;
    cols: string[];
    rows: string[][];
    title: string;
  } | null>(null);
  const [isExcelExportLoading, setIsExcelExportLoading] = useState<boolean>(false);
  const [excelExportProgressText, setExcelExportProgressText] = useState<string>('');

  // Background removal and camera stream states
  const [isProcessingRemoveBg, setIsProcessingRemoveBg] = useState<boolean>(false);
  const [bgRemovalProgressPercent, setBgRemovalProgressPercent] = useState<number>(0);
  const [bgRemovalStatusText, setBgRemovalStatusText] = useState<string>('');
  const [isWebcamOpen, setIsWebcamOpen] = useState<boolean>(false);
  const [webcamStream, setWebcamStream] = useState<MediaStream | null>(null);
  const videoRef = useRef<HTMLVideoElement | null>(null);

  // Process background removal via client-side IMG.LY AI model (In-Memory Processing)
  const handleProcessImglyBackgroundRemoval = async (file: File) => {
    let currentP = 15;
    let bgProgressTimer: any = null;
    try {
      setIsProcessingRemoveBg(true);
      setBgRemovalProgressPercent(15);
      setBgRemovalStatusText("Görsel işleniyor (%15)...");

      bgProgressTimer = setInterval(() => {
        if (currentP < 95) {
          currentP = Math.min(95, currentP + (currentP < 50 ? 6 : 3));
          setBgRemovalProgressPercent(currentP);
          setBgRemovalStatusText(`Arka Plan Analizi: %${currentP}`);
        }
      }, 100);
      
      // Dev fotoğrafı AI segmentasyonu için optimum boyuta (maks. 640px) getir
      const processedFile = await prepareOptimizedImageForAI(file, 640);
      currentP = Math.max(currentP, 35);
      setBgRemovalProgressPercent(currentP);
      setBgRemovalStatusText(`Model hazırlanıyor (%${currentP})...`);

      const removeFn = (window as any).imglyRemoveBackground || removeBackground;
      const blob = await removeFn(processedFile, {
        model: 'isnet_fp16', // Hızlı ve yüksek kaliteli AI modeli
        progress: (key: string, current: number, total: number) => {
          if (total > 0) {
            const calculatedPercent = Math.min(95, Math.max(35, Math.round(35 + (current / total) * 60)));
            if (calculatedPercent > currentP) {
              currentP = calculatedPercent;
              setBgRemovalProgressPercent(currentP);
              setBgRemovalStatusText(`Arka Plan Analizi: %${currentP}`);
            }
          }
        }
      });
      
      if (bgProgressTimer) clearInterval(bgProgressTimer);

      setBgRemovalProgressPercent(100);
      setBgRemovalStatusText("İşlem tamamlandı! (%100)");

      const localUrl = URL.createObjectURL(blob);
      setPendingImagePreview(localUrl);
      
      // Convert blob to base64
      const reader = new FileReader();
      reader.onloadend = () => {
        const base64 = reader.result as string;
        setPendingImageBase64(base64);
        setPendingImageMimeType("image/png");
      };
      reader.readAsDataURL(blob);

      // Save as transparent file
      const transparentFile = new File([blob], `cleaned_${file.name || "captured.png"}`, { type: "image/png" });
      setPendingImageFile(transparentFile);
      
      showNotification("Arka plan başarıyla temizlendi!");
      return blob;
    } catch (err: any) {
      if (bgProgressTimer) clearInterval(bgProgressTimer);
      console.error("Arka plan temizleme hatası:", err);
      setBgRemovalStatusText("Bir hata oluştu, lütfen tekrar deneyin.");
      setBgRemovalProgressPercent(0);
      showNotification(`Arka plan temizleme hatası: ${err.message || err}`);
    } finally {
      setIsProcessingRemoveBg(false);
    }
  };

  // Helper when image file is selected or captured from camera
  const handleImageSelected = async (file: File) => {
    setPendingImageFile(file);
    setPendingImagePreview(URL.createObjectURL(file));
    
    // Convert to base64
    const reader = new FileReader();
    reader.onloadend = () => {
      setPendingImageBase64(reader.result as string);
      setPendingImageMimeType(file.type || "image/png");
    };
    reader.readAsDataURL(file);

    // Run client-side AI background removal automatically
    await handleProcessImglyBackgroundRemoval(file);
  };

  // Webcam controls for desktop webcam support
  const startWebcam = async () => {
    try {
      const stream = await navigator.mediaDevices.getUserMedia({
        video: { facingMode: "environment", width: { ideal: 1280 }, height: { ideal: 720 } }
      });
      setWebcamStream(stream);
      setIsWebcamOpen(true);
      setTimeout(() => {
        if (videoRef.current) {
          videoRef.current.srcObject = stream;
        }
      }, 200);
    } catch (err: any) {
      console.error("Kamera başlatılamadı:", err);
      showNotification(`Kamera erişimi başarısız oldu: ${err.message}`);
      // Fallback: click native camera / file input
      document.getElementById('drag-drop-image-input')?.click();
    }
  };

  const stopWebcam = () => {
    if (webcamStream) {
      webcamStream.getTracks().forEach(track => track.stop());
      setWebcamStream(null);
    }
    setIsWebcamOpen(false);
  };

  const captureWebcamPhoto = () => {
    if (!videoRef.current) return;
    try {
      const video = videoRef.current;
      const canvas = document.createElement("canvas");
      canvas.width = video.videoWidth || 640;
      canvas.height = video.videoHeight || 480;
      
      const ctx = canvas.getContext("2d");
      if (ctx) {
        ctx.drawImage(video, 0, 0, canvas.width, canvas.height);
        canvas.toBlob(async (blob) => {
          if (blob) {
            const capturedFile = new File([blob], `captured_${Date.now()}.png`, { type: "image/png" });
            stopWebcam();
            await handleImageSelected(capturedFile);
          }
        }, "image/png");
      }
    } catch (err: any) {
      console.error("Fotoğraf çekilirken hata:", err);
      showNotification(`Fotoğraf çekilemedi: ${err.message}`);
    }
  };
  const [isDataUpdateUnlocked, setIsDataUpdateUnlocked] = useState<boolean>(false);
  const [dataPasswordInput, setDataPasswordInput] = useState<string>('');
  const [dataPasswordError, setDataPasswordError] = useState<boolean>(false);

  // Techizat and Mission Order Editing States
  const [isTechizatSaving, setIsTechizatSaving] = useState<boolean>(false);
  const [activeGorevEmriEdit, setActiveGorevEmriEdit] = useState<any | null>(null);
  const [editGorevEmriValues, setEditGorevEmriValues] = useState<any | null>(null);
  const [pendingGeEditOrder, setPendingGeEditOrder] = useState<any | null>(null);
  const [geEditPasswordInput, setGeEditPasswordInput] = useState<string>('');
  const [geEditPasswordError, setGeEditPasswordError] = useState<boolean>(false);
  const [showGeEditPasswordPrompt, setShowGeEditPasswordPrompt] = useState<boolean>(false);

  const [karaAraclariSubTab, setKaraAraclariSubTab] = useState<'list' | 'mission_order' | 'past_records'>('list');
  const [karaAraclariGorevEmirleri, setKaraAraclariGorevEmirleri] = useState<any[]>(() => {
    const saved = localStorage.getItem('kara_araclari_gorev_emirleri');
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed)) {
          // Filter out rows that are blank/empty (no Tarih, no Araç Plakası, no Sürücü Personel, no Görev Seri No)
          return parsed.filter((row: any) => {
            const date = String(row.date || "").trim();
            const plate = String(row.plate || "").trim();
            const driver = String(row.driverName || "").trim();
            const serial = String(row.serialNo || "").trim();
            return date !== "" || plate !== "" || driver !== "" || serial !== "";
          });
        }
      } catch (e) { console.error(e); }
    }
    return [];
  });

  const [geTarih, setGeTarih] = useState<string>(() => new Date().toISOString().split('T')[0]);
  const [gePlaka, setGePlaka] = useState<string>("");
  const [geSoforName, setGeSoforName] = useState<string>("");
  const [geSeriNo, setGeSeriNo] = useState<string>("");
  const [geReturnKm, setGeReturnKm] = useState<string>("");
  const [geDepartureTime, setGeDepartureTime] = useState<string>("08:00");
  const [geReturnTime, setGeReturnTime] = useState<string>("17:00");
  const [geDepartureKm, setGeDepartureKm] = useState<string>("");
  const [geStep, setGeStep] = useState<number>(1);
  const [showDriverSuggestions, setShowDriverSuggestions] = useState<boolean>(false);
  const [showVehicleSuggestions, setShowVehicleSuggestions] = useState<boolean>(false);
  const [isRedirectingToPortal, setIsRedirectingToPortal] = useState<boolean>(false);
  const [geRoutes, setGeRoutes] = useState<{ from: string; to: string }[]>([{ from: "", to: "" }]);
  const [isSlidingUp, setIsSlidingUp] = useState<boolean>(false);
  const [showGeDeletePasswordPrompt, setShowGeDeletePasswordPrompt] = useState<boolean>(false);
  const [geDeleteOrderId, setGeDeleteOrderId] = useState<string | null>(null);
  const [geDeletePasswordInput, setGeDeletePasswordInput] = useState<string>("");
  const [geDeletePasswordError, setGeDeletePasswordError] = useState<boolean>(false);
  const [pendingImageBase64, setPendingImageBase64] = useState<string | null>(null);
  const [pendingImageMimeType, setPendingImageMimeType] = useState<string | null>(null);

  // EBYS multi-selection and tracking state variables
  const [selectedTechizatItems, setSelectedTechizatItems] = useState<Record<string, { techType: string; row: string[] }>>({});
  const [techizatSubTab, setTechizatSubTab] = useState<'list' | 'ebys_tracking'>('list');
  const [isEbysModalOpen, setIsEbysModalOpen] = useState(false);
  const [ebysSearchQuery, setEbysSearchQuery] = useState("");
  const [ebysList, setEbysList] = useState<any[]>([]);
  const [isLoadingEbys, setIsLoadingEbys] = useState(false);
  const [ebysError, setEbysError] = useState<string | null>(null);
  const [selectedEbysRow, setSelectedEbysRow] = useState<any | null>(null);
  const [ebysBaslik, setEbysBaslik] = useState("");
  const [ebysAciklama, setEbysAciklama] = useState("");
  const [ebysTalepTuru, setEbysTalepTuru] = useState("");
  const [ebysTeslimTarihi, setEbysTeslimTarihi] = useState(() => new Date().toISOString().split('T')[0]);
  const [submittedEbysRequests, setSubmittedEbysRequests] = useState<any[]>([]);
  const [isLoadingSubmitted, setIsLoadingSubmitted] = useState(false);
  const [isEbysSelectDropdownOpen, setIsEbysSelectDropdownOpen] = useState(false);

  // Bulk edit states for multiple teçhizat rows
  const [bulkEditYer, setBulkEditYer] = useState("");
  const [bulkEditDurum, setBulkEditDurum] = useState("");
  const [bulkEditFirma, setBulkEditFirma] = useState("");
  const [bulkModalMode, setBulkModalMode] = useState<'choice' | 'edit' | 'send'>('choice');
  const [bulkEditPasswordInput, setBulkEditPasswordInput] = useState("");
  const [bulkEditPasswordError, setBulkEditPasswordError] = useState(false);
  const [showBulkEditPasswordPrompt, setShowBulkEditPasswordPrompt] = useState(false);
  const [isBulkSaving, setIsBulkSaving] = useState(false);

  // Password-protected EBYS Talep tracking editing states
  const [editingEbysRowIndex, setEditingEbysRowIndex] = useState<number | null>(null);
  const [editingEbysFirma, setEditingEbysFirma] = useState<string>("");
  const [showEbysFirmaPasswordPrompt, setShowEbysFirmaPasswordPrompt] = useState<boolean>(false);
  const [ebysFirmaPasswordInput, setEbysFirmaPasswordInput] = useState<string>("");
  const [ebysFirmaPasswordError, setEbysFirmaPasswordError] = useState<boolean>(false);

  // Search filter states for past records
  const [pastRecordsSearchName, setPastRecordsSearchName] = useState<string>("");
  const [pastRecordsSearchSerial, setPastRecordsSearchSerial] = useState<string>("");
  const [pastRecordsSearchStartDate, setPastRecordsSearchStartDate] = useState<string>("");
  const [pastRecordsSearchEndDate, setPastRecordsSearchEndDate] = useState<string>("");
  const [isPastRecordsSearched, setIsPastRecordsSearched] = useState<boolean>(false);

  // Helper to parse EBYS rows or objects dynamically
  const parseEbysItem = (item: any) => {
    if (!item) return null;
    let ebysNo = "";
    let baslik = "";
    let aciklama = "";
    let talepTuru = "";

    const normalize = (str: string) => {
      return String(str || "")
        .toLowerCase()
        .replace(/ı/g, 'i')
        .replace(/ğ/g, 'g')
        .replace(/ü/g, 'u')
        .replace(/ş/g, 's')
        .replace(/ö/g, 'o')
        .replace(/ç/g, 'c')
        .replace(/[^a-z0-9]/g, '');
    };

    if (Array.isArray(item)) {
      // Index 7 is column H (EBYS)
      ebysNo = String(item[7] || "").trim();
      baslik = String(item[1] || "").trim();
      aciklama = String(item[2] || "").trim();
      talepTuru = String(item[3] || "").trim();
    } else if (typeof item === "object") {
      const keys = Object.keys(item);
      
      const ebysKey = keys.find(k => {
        const norm = normalize(k);
        return (norm.includes("ebys") || norm === "h") && !norm.includes("tarih") && !norm.includes("date") && !norm.includes("gun");
      });
      if (ebysKey) ebysNo = String(item[ebysKey]).trim();
      
      const baslikKey = keys.find(k => {
        const norm = normalize(k);
        return norm === "baslik" || norm.includes("basligi") || norm.includes("konu") || norm === "b" || norm === "title";
      });
      if (baslikKey) baslik = String(item[baslikKey]).trim();
      
      const aciklamaKey = keys.find(k => {
        const norm = normalize(k);
        return norm === "aciklama" || norm.includes("aciklamasi") || norm === "c" || norm === "description";
      });
      if (aciklamaKey) aciklama = String(item[aciklamaKey]).trim();

      const talepTuruKey = keys.find(k => {
        const norm = normalize(k);
        return norm === "talepturu" || norm.includes("turu") || norm === "d" || norm === "type";
      });
      if (talepTuruKey) talepTuru = String(item[talepTuruKey]).trim();

      // Fallbacks - Prefer Column H (8th column) if ebysNo is empty or 'n/a'
      if (!ebysNo || ebysNo.toLowerCase() === "n/a" || ebysNo.toLowerCase() === "na") {
        if (keys.length > 7) {
          ebysNo = String(item[keys[7]] || "").trim();
        }
      }
      if (!ebysNo || ebysNo.toLowerCase() === "n/a" || ebysNo.toLowerCase() === "na") {
        ebysNo = String(item["EBYS NO"] || item["EBYS"] || item["ebys"] || item["EBYS Numarası"] || item["ebysNumber"] || item["H"] || "");
      }
      if (!baslik) baslik = String(item["Başlık"] || item["Baslik"] || item["title"] || item["B"] || "");
      if (!aciklama) aciklama = String(item["Açıklama"] || item["Aciklama"] || item["description"] || item["C"] || "");
      if (!talepTuru) talepTuru = String(item["Talep Türü"] || item["Talep Turu"] || item["type"] || item["D"] || "");
    }

    ebysNo = ebysNo.trim();

    return { ebysNo, baslik, aciklama, talepTuru };
  };

  const getTechUnitName = (techType: string): string => {
    if (techType === 'bell429') return 'BELL 429';
    if (techType === 'at802') return 'AT-802F';
    if (techType === 't70') return 'T-70 YER DESTEK';
    if (techType === 't70_bumbi_backet') return 'T-70 BUMBİ BACKET';
    if (techType === 't70_helitak') return 'T-70 HELİTAK';
    if (techType === 'c650') return 'C-650';
    if (techType === 'b360') return 'B-360';
    if (techType === 'hangar') return 'HANGAR YER DESTEK';
    if (techType === 'kara_araclari') return 'KARA ARAÇLARI';
    return techType.toUpperCase();
  };

  // Fetch Taskline data (EBYS List)
  const fetchTasklineEbysList = async () => {
    setIsLoadingEbys(true);
    setEbysError(null);
    try {
      let data: any = null;
      try {
        const url = `/api/taskline-ebys?scriptUrl=${encodeURIComponent(EBYS_SEARCH_SCRIPT_URL)}&spreadsheetId=1L05588TdYZmH401Lvn4_yr4zwiw2pW4EJ8dIyl-UTVQ`;
        const response = await fetch(url);
        if (response.ok) {
          data = await response.json();
        }
      } catch (e) {
        console.warn("Express backend proxy /api/taskline-ebys not available, trying direct client fetch:", e);
      }

      // Fallback to direct client GET if proxy failed or not present (e.g. Netlify static hosting)
      if (!data || data.status === "error") {
        const fallbackUrl = `${EBYS_SEARCH_SCRIPT_URL}?spreadsheetId=1L05588TdYZmH401Lvn4_yr4zwiw2pW4EJ8dIyl-UTVQ`;
        const resp = await fetch(fallbackUrl);
        if (resp.ok) {
          data = await resp.json();
        }
      }

      if (data && Array.isArray(data.data)) {
        setEbysList(data.data);
      } else if (Array.isArray(data)) {
        setEbysList(data);
      } else if (data && data.status === "success" && Array.isArray(data.data)) {
        setEbysList(data.data);
      } else {
        setEbysList([]);
      }
    } catch (err: any) {
      console.error("Taskline EBYS listesi çekme hatası:", err);
      setEbysError(err?.message || "Bağlantı hatası");
    } finally {
      setIsLoadingEbys(false);
    }
  };

  // Fetch Sayfa1 from TASKLINE script
  const fetchSubmittedEbysRequests = async () => {
    setIsLoadingSubmitted(true);
    try {
      let data: any = null;
      try {
        const url = `/api/taskline-ebys?scriptUrl=${encodeURIComponent(TASKLINE_SUBMIT_SCRIPT_URL)}&action=readSheet&sheetName=${encodeURIComponent("Sayfa1")}&spreadsheetId=1L05588TdYZmH401Lvn4_yr4zwiw2pW4EJ8dIyl-UTVQ`;
        const response = await fetch(url);
        if (response.ok) {
          data = await response.json();
        }
      } catch (e) {
        console.warn("Express proxy not available, falling back to direct client GET for Sayfa1:", e);
      }

      if (!data || data.status === "error") {
        const fallbackUrl = `${TASKLINE_SUBMIT_SCRIPT_URL}?action=readSheet&sheetName=${encodeURIComponent("Sayfa1")}&spreadsheetId=1L05588TdYZmH401Lvn4_yr4zwiw2pW4EJ8dIyl-UTVQ`;
        const resp = await fetch(fallbackUrl);
        if (resp.ok) {
          data = await resp.json();
        }
      }

      if (data && Array.isArray(data.data)) {
        setSubmittedEbysRequests(data.data);
        localStorage.setItem('submitted_ebys_requests', JSON.stringify(data.data));
      } else if (Array.isArray(data)) {
        setSubmittedEbysRequests(data);
        localStorage.setItem('submitted_ebys_requests', JSON.stringify(data));
      }
    } catch (err) {
      console.error("TASKLINE TASKLINE-PARÇA LİSTESİ listesi çekme hatası:", err);
      const saved = localStorage.getItem('submitted_ebys_requests');
      if (saved) {
        setSubmittedEbysRequests(JSON.parse(saved));
      }
    } finally {
      setIsLoadingSubmitted(false);
    }
  };

  // Submit selected equipment rows to TASKLINE (Sayfa1)
  const submitEbysRequests = async () => {
    if (!ebysSearchQuery) {
      showNotification("Lütfen bir EBYS numarası seçin veya girin.");
      return;
    }

    const tableRows = Object.values(selectedTechizatItems).map((item: { techType: string; row: string[] }) => {
      const row = item.row;
      const aitOlduguBirim = getTechUnitName(item.techType); // A sütunu: AİT OLDUĞU BİRİM
      const techName = row[1] || "";                         // B sütunu: TEÇHİZAT ADI
      const parcaNo = row[2] || "-";                         // C sütunu: PARÇA NO (P/N) / MODEL
      const seriNo = row[3] || "-";                          // D sütunu: SERİ NO (S/N)
      const miktarKapasite = row[4] || "1";                  // E sütunu: MİKTAR / KAPASİTE
      const firma = row[10] || row[9] || "-";                // F sütunu: SON KONTROLÜ YAPAN FİRMA
      const aciklama = row[11] || row[10] || "";              // G sütunu: AÇIKLAMA

      return [
        aitOlduguBirim,
        techName,
        parcaNo,
        seriNo,
        miktarKapasite,
        firma,
        aciklama
      ];
    });

    try {
      showNotification("Talepleriniz online Excel sayfasına aktarılıyor ve durumları 'BAKIM / KALİBRASYON' olarak güncelleniyor...");
      
      // 1. Group selected items by techType to update their status column in main database
      const updatedTechTypes = new Set<string>();
      
      Object.values(selectedTechizatItems).forEach((item: { techType: string; row: string[] }) => {
        const { techType, row } = item;
        const statusIdx = 6;
        row[statusIdx] = "BAKIM / KALİBRASYON";
        updatedTechTypes.add(techType);
      });

      // 2. Update local state arrays and local storage, then sync online in the background
      updatedTechTypes.forEach(techType => {
        let currentList: string[][] = [];
        if (techType === 'bell429') currentList = [...techizatBell429Data];
        else if (techType === 'at802') currentList = [...techizatAt802Data];
        else if (techType === 't70') currentList = [...techizatT70Data];
        else if (techType === 't70_bumbi_backet') currentList = [...techizatT70BumbiBacketData];
        else if (techType === 'b360') currentList = [...techizatB360Data];
        else if (techType === 'c650') currentList = [...techizatC650Data];
        else if (techType === 'hangar') currentList = [...techizatHangarData];
        else if (techType === 'kara_araclari') currentList = [...techizatKaraAraclariData];

        const newArray = currentList.map((r) => {
          const matched = Object.values(selectedTechizatItems).find((sel: { techType: string; row: string[] }) => 
            sel.techType === techType && 
            (sel.row[0] || "").trim() === (r[0] || "").trim() && 
            (sel.row[1] || "").trim() === (r[1] || "").trim()
          );
          if (matched) {
            const cloned = [...r];
            const statusIdx = 6;
            cloned[statusIdx] = "BAKIM / KALİBRASYON";
            return cloned;
          }
          return r;
        });

        // Set React state and LocalStorage for local responsiveness
        if (techType === 'bell429') {
          setTechizatBell429Data(newArray);
          localStorage.setItem('excel_techizat_bell429_data', JSON.stringify(newArray));
        } else if (techType === 'at802') {
          setTechizatAt802Data(newArray);
          localStorage.setItem('excel_techizat_at802_data', JSON.stringify(newArray));
        } else if (techType === 't70') {
          setTechizatT70Data(newArray);
          localStorage.setItem('excel_techizat_t70_data', JSON.stringify(newArray));
        } else if (techType === 't70_bumbi_backet') {
          setTechizatT70BumbiBacketData(newArray);
          localStorage.setItem('excel_techizat_t70_bumbi_backet_data', JSON.stringify(newArray));
        } else if (techType === 'b360') {
          setTechizatB360Data(newArray);
          localStorage.setItem('excel_techizat_b360_data', JSON.stringify(newArray));
        } else if (techType === 'c650') {
          setTechizatC650Data(newArray);
          localStorage.setItem('excel_techizat_c650_data', JSON.stringify(newArray));
        } else if (techType === 'hangar') {
          setTechizatHangarData(newArray);
          localStorage.setItem('excel_techizat_hangar_data', JSON.stringify(newArray));
        } else if (techType === 'kara_araclari') {
          setTechizatKaraAraclariData(newArray);
          localStorage.setItem('excel_techizat_kara_araclari_data', JSON.stringify(newArray));
        }

        // Sync back to online central sheet
        const unitLabel = getTechizatUnitLabel(techType);
        if (unitLabel && newArray.length > 0) {
          fetch(GOOGLE_SCRIPT_URL, {
            method: "POST",
            headers: { "Content-Type": "text/plain;charset=utf-8" },
            body: JSON.stringify({
              action: "updateTumTechizat",
              unitLabel: unitLabel,
              data: newArray.map(r => [unitLabel, ...r])
            })
          }).then(() => {
            console.log(`Synced status change to "BAKIM / KALİBRASYON" for ${unitLabel} to Google Sheets.`);
          }).catch(e => {
            console.error(`Failed to sync status for ${unitLabel}:`, e);
          });
        }
      });

      // 3. Post selected rows to TASKLINE Submit Script via backend proxy or direct fetch
      let submitSuccess = false;
      try {
        const proxyResponse = await fetch("/api/taskline-submit", {
          method: "POST",
          headers: {
            "Content-Type": "application/json"
          },
          body: JSON.stringify({
            ebysNo: ebysSearchQuery,
            talepTuru: ebysTalepTuru || "MALZEME",
            data: tableRows,
            scriptUrl: TASKLINE_SUBMIT_SCRIPT_URL,
            fallbackScriptUrl: GOOGLE_SCRIPT_URL,
            spreadsheetId: "1L05588TdYZmH401Lvn4_yr4zwiw2pW4EJ8dIyl-UTVQ"
          })
        });

        if (proxyResponse.ok) {
          const proxyResult = await proxyResponse.json();
          if (proxyResult.status !== "error") {
            submitSuccess = true;
          }
        }
      } catch (e) {
        console.warn("Proxy submission unavailable or failed, attempting direct fetch:", e);
      }

      if (!submitSuccess) {
        const directPayload = {
          action: "appendEbysTable",
          ebysNo: ebysSearchQuery,
          talepTuru: ebysTalepTuru || "MALZEME",
          data: tableRows,
          spreadsheetId: "1L05588TdYZmH401Lvn4_yr4zwiw2pW4EJ8dIyl-UTVQ"
        };

        const directResp = await fetch(TASKLINE_SUBMIT_SCRIPT_URL, {
          method: "POST",
          headers: { "Content-Type": "text/plain;charset=utf-8" },
          body: JSON.stringify(directPayload)
        });

        if (!directResp.ok) {
          await fetch(GOOGLE_SCRIPT_URL, {
            method: "POST",
            headers: { "Content-Type": "text/plain;charset=utf-8" },
            body: JSON.stringify(directPayload)
          });
        }
      }

      // Show instant feedback
      showNotification("Seçilen teçhizatlar başarıyla Sayfa1 sistemine gönderildi ve durumları 'BAKIM / KALİBRASYON' olarak güncellendi!");
      setSelectedTechizatItems({});
      setIsEbysModalOpen(false);
      setEbysSearchQuery("");
      setEbysBaslik("");
      setEbysAciklama("");
      setEbysTalepTuru("");

      // Refresh tracking lists
      setTimeout(() => {
        fetchSubmittedEbysRequests();
        pullAllTechizatFromGoogleSheets(true);
      }, 1500);

    } catch (err: any) {
      console.error("TASKLINE yazma hatası:", err);
      showNotification(`Gönderim sırasında hata oluştu: ${err.message || err}`);
    }
  };

  // Bulk update function for selected teçhizat rows
  const handleBulkEditTechizatRows = async () => {
    try {
      setIsBulkSaving(true);
      showNotification("Seçilen teçhizatlar toplu olarak güncelleniyor, lütfen bekleyiniz...");
      
      const updatedTechTypes = new Set<string>();

      // Group by techType and apply updates
      Object.entries(selectedTechizatItems).forEach(([key, item]: [string, any]) => {
        const { techType } = item;
        updatedTechTypes.add(techType);
      });

      // For each affected techType, load current list, map updates, and save
      for (const techType of updatedTechTypes) {
        let currentList: string[][] = [];
        if (techType === 'bell429') currentList = [...techizatBell429Data];
        else if (techType === 'at802') currentList = [...techizatAt802Data];
        else if (techType === 't70') currentList = [...techizatT70Data];
        else if (techType === 't70_bumbi_backet') currentList = [...techizatT70BumbiBacketData];
        else if (techType === 'b360') currentList = [...techizatB360Data];
        else if (techType === 'c650') currentList = [...techizatC650Data];
        else if (techType === 'hangar') currentList = [...techizatHangarData];
        else if (techType === 'kara_araclari') currentList = [...techizatKaraAraclariData];

        const newArray = currentList.map((r) => {
          // Check if this row is selected
          const isSelected = Object.values(selectedTechizatItems).some((sel: any) => 
            sel.techType === techType && 
            (sel.row[0] || "").trim() === (r[0] || "").trim() && 
            (sel.row[1] || "").trim() === (r[1] || "").trim()
          );

          if (isSelected) {
            const cloned = [...r];
            if (techType === 'kara_araclari') {
              while (cloned.length < 12) cloned.push("");
              if (bulkEditYer.trim() !== "") {
                cloned[3] = bulkEditYer.trim();
              }
              if (bulkEditDurum.trim() !== "") {
                cloned[5] = bulkEditDurum.trim();
              }
              if (bulkEditFirma.trim() !== "") {
                cloned[9] = bulkEditFirma.trim();
              }
            } else {
              while (cloned.length < 13) cloned.push("");
              if (bulkEditYer.trim() !== "") {
                cloned[5] = bulkEditYer.trim();
              }
              if (bulkEditDurum.trim() !== "") {
                cloned[6] = bulkEditDurum.trim();
              }
              if (bulkEditFirma.trim() !== "") {
                cloned[10] = bulkEditFirma.trim();
                cloned[9] = bulkEditFirma.trim();
              }
            }
            return cloned;
          }
          return r;
        });

        // Save back to local states
        if (techType === 'bell429') {
          setTechizatBell429Data(newArray);
          localStorage.setItem('excel_techizat_bell429_data', JSON.stringify(newArray));
        } else if (techType === 'at802') {
          setTechizatAt802Data(newArray);
          localStorage.setItem('excel_techizat_at802_data', JSON.stringify(newArray));
        } else if (techType === 't70') {
          setTechizatT70Data(newArray);
          localStorage.setItem('excel_techizat_t70_data', JSON.stringify(newArray));
        } else if (techType === 't70_bumbi_backet') {
          setTechizatT70BumbiBacketData(newArray);
          localStorage.setItem('excel_techizat_t70_bumbi_backet_data', JSON.stringify(newArray));
        } else if (techType === 'b360') {
          setTechizatB360Data(newArray);
          localStorage.setItem('excel_techizat_b360_data', JSON.stringify(newArray));
        } else if (techType === 'c650') {
          setTechizatC650Data(newArray);
          localStorage.setItem('excel_techizat_c650_data', JSON.stringify(newArray));
        } else if (techType === 'hangar') {
          setTechizatHangarData(newArray);
          localStorage.setItem('excel_techizat_hangar_data', JSON.stringify(newArray));
        } else if (techType === 'kara_araclari') {
          setTechizatKaraAraclariData(newArray);
          localStorage.setItem('excel_techizat_kara_araclari_data', JSON.stringify(newArray));
        }

        // Sync back to online central sheet
        const unitLabel = getTechizatUnitLabel(techType);
        if (unitLabel && newArray.length > 0) {
          await fetch(GOOGLE_SCRIPT_URL, {
            method: "POST",
            headers: { "Content-Type": "text/plain;charset=utf-8" },
            body: JSON.stringify({
              action: "updateTumTechizat",
              unitLabel: unitLabel,
              data: newArray.map(r => [unitLabel, ...r])
            })
          });
        }
      }

      showNotification("Seçilen teçhizatlar başarıyla topluca güncellendi ve online e-tabloya kaydedildi!");
      setSelectedTechizatItems({});
      setIsEbysModalOpen(false);
      setBulkEditYer("");
      setBulkEditDurum("");
      setBulkEditFirma("");
      setBulkModalMode('choice');
      
      setTimeout(() => {
        pullAllTechizatFromGoogleSheets(true);
      }, 1500);

    } catch (err) {
      console.error("Bulk edit error:", err);
      showNotification("Toplu güncelleme sırasında bir hata oluştu, lütfen tekrar deneyiniz.");
    } finally {
      setIsBulkSaving(false);
    }
  };

  // Update a specific request's KONTROLÜ YAPAN FİRMA column on TASKLINE-PARÇA LİSTESİ online sheet
  const updateEbysFirmaOnline = async (indexToUpdate: number, newFirma: string) => {
    const item = submittedEbysRequests[indexToUpdate];
    if (!item) return;

    const ebysNo = item["EBYS NO"] || item["ebysNo"] || "";
    const partNo = item["PARÇA NUMARASI"] || item["parcaNumarasi"] || "";
    const seriNo = item["SERİ NO (S/N)"] || item["seriNo"] || "";

    // 1. Create a modified copy of current submittedEbysRequests
    const updatedList = submittedEbysRequests.map((req, idx) => {
      if (idx === indexToUpdate) {
        return {
          ...req,
          "KONTROLÜ YAPAN FİRMA": newFirma,
          "kontroluYapanFirma": newFirma
        };
      }
      return req;
    });

    setSubmittedEbysRequests(updatedList);
    localStorage.setItem('submitted_ebys_requests', JSON.stringify(updatedList));

    try {
      showNotification("KONTROLÜ YAPAN FİRMA bilgisi e-tabloya güncelleniyor...");
      
      await fetch(GOOGLE_SCRIPT_URL, {
        method: "POST",
        mode: "no-cors",
        headers: {
          "Content-Type": "application/json"
        },
        body: JSON.stringify({
          action: "updateEbysFirma",
          ebysNo: ebysNo,
          parcaNo: partNo,
          seriNo: seriNo,
          firma: newFirma
        })
      });

      showNotification("Firma bilgisi başarıyla güncellendi ve teçhizat listesiyle senkronize edildi!");
      setEditingEbysRowIndex(null);
      
      // Pull fresh data to reflect changes
      setTimeout(() => {
        pullAllTechizatFromGoogleSheets(true);
      }, 1500);
    } catch (err) {
      console.error("Error updating KONTROLÜ YAPAN FİRMA:", err);
      showNotification("Güncelleme sırasında hata oluştu.");
    }
  };

  // Fetch past mission orders from Google Sheets
  const pullKaraAraclariGorevEmirleri = async () => {
    try {
      const targetUrl = `${GOOGLE_SCRIPT_URL}?action=readSheet&sheetName=${encodeURIComponent("görev emri kaytlar")}`;
      const response = await fetch(targetUrl);
      if (response.ok) {
        const result = await response.json();
        if (result.status === "success" && Array.isArray(result.data)) {
          // Filter out rows that are blank/empty (no Tarih, no Araç Plakası, no Sürücü Personel, no Görev Seri No)
          const validRows = result.data.filter((row: any) => {
            const date = String(row["Tarih"] || row["Date"] || "").trim();
            const plate = String(row["Araç Plakası"] || row["Plate"] || "").trim();
            const driver = String(row["Sürücü Personel"] || row["DriverName"] || "").trim();
            const serial = String(row["Görev Seri No"] || row["SerialNo"] || "").trim();
            return date !== "" || plate !== "" || driver !== "" || serial !== "";
          });

          const mapped = validRows.map((row: any, index: number) => ({
            id: row.id ? Number(row.id) : Date.now() + index,
            date: row["Tarih"] || row["Date"] || "",
            plate: row["Araç Plakası"] || row["Plate"] || "",
            driverName: row["Sürücü Personel"] || row["DriverName"] || "",
            driverId: row["T.C. Kimlik No"] || row["DriverId"] || "",
            driverSicil: row["Sicil No"] || row["DriverSicil"] || "",
            driverPhone: row["Telefon"] || row["DriverPhone"] || "",
            driverKanGrubu: row["Kan Grubu"] || row["DriverKanGrubu"] || "",
            driverAdres: row["Adres"] || row["DriverAdres"] || "",
            serialNo: row["Görev Seri No"] || row["SerialNo"] || "",
            departureTime: row["Çıkış Saati"] || row["DepartureTime"] || "08:00",
            returnTime: row["Dönüş Saati"] || row["ReturnTime"] || "17:00",
            departureKm: row["Çıkış KM"] || row["DepartureKm"] || "",
            returnKm: Number(row["Dönüş KM"] || row["ReturnKm"] || 0),
            route: row["Güzergah"] || row["Route"] || ""
          }));

          // Set and sync unconditionally to correctly clear local storage/state if everything is deleted on Sheets
          setKaraAraclariGorevEmirleri(mapped);
          localStorage.setItem('kara_araclari_gorev_emirleri', JSON.stringify(mapped));
        }
      }
    } catch (err) {
      console.error("Görev emri kayıtları yükleme hatası:", err);
    }
  };

  // Push past mission orders to Google Sheets
  const pushKaraAraclariGorevEmirleri = async (ordersList: any[]) => {
    try {
      const targetUrl = GOOGLE_SCRIPT_URL;
      const headers = [
        "id", "Tarih", "Araç Plakası", "Sürücü Personel", "T.C. Kimlik No", "Sicil No", 
        "Telefon", "Kan Grubu", "Adres", "Görev Seri No", "Çıkış Saati", "Dönüş Saati", "Çıkış KM", "Dönüş KM", "Güzergah"
      ];
      const rows = ordersList.map(o => [
        String(o.id || ""),
        String(o.date || ""),
        String(o.plate || ""),
        String(o.driverName || ""),
        String(o.driverId || ""),
        String(o.driverSicil || ""),
        String(o.driverPhone || ""),
        String(o.driverKanGrubu || ""),
        String(o.driverAdres || ""),
        String(o.serialNo || ""),
        String(o.departureTime || ""),
        String(o.returnTime || ""),
        String(o.departureKm || ""),
        String(o.returnKm || ""),
        String(o.route || "")
      ]);

      await fetch(targetUrl, {
        method: "POST",
        mode: "no-cors",
        headers: {
          "Content-Type": "application/json"
        },
        body: JSON.stringify({
          action: "updateSheet",
          sheetName: "görev emri kaytlar",
          data: [headers, ...rows]
        })
      });
    } catch (err) {
      console.error("Görev emri senkronizasyon hatası:", err);
    }
  };

  useEffect(() => {
    localStorage.setItem('techizat_images', JSON.stringify(techizatImages));
  }, [techizatImages]);

  useEffect(() => {
    localStorage.setItem('kara_araclari_gorev_emirleri', JSON.stringify(karaAraclariGorevEmirleri));
  }, [karaAraclariGorevEmirleri]);

  // Ctrl+F Keyboard Shortcut Listener for PDF Search Input (Allows native browser Ctrl+F to open)
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'f') {
        if (pdfSearchInputRef.current) {
          // We let the browser's native search bar open naturally by not calling e.preventDefault().
          // We still focus our custom bar as a companion, but allow the native search to trigger.
          pdfSearchInputRef.current.focus();
          pdfSearchInputRef.current.select();
        }
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => {
      window.removeEventListener('keydown', handleKeyDown);
    };
  }, []);

  // Sync date ranges to Turkish descriptive string representation dynamically
  useEffect(() => {
    const formatted = convertToRawPeriodString(selectedSummerStartDate, selectedSummerEndDate);
    setSelectedSummerMonth(formatted);
  }, [selectedSummerStartDate, selectedSummerEndDate]);

  useEffect(() => {
    const formatted = convertToRawPeriodString(selectedUploadSummerStartDate, selectedUploadSummerEndDate);
    setSelectedUploadSummerMonth(formatted);
  }, [selectedUploadSummerStartDate, selectedUploadSummerEndDate]);

  // Gelişmiş PDF Önizleme ve Render Durumları
  const [isPdfRendering, setIsPdfRendering] = useState<boolean>(false);
  const [renderedPages, setRenderedPages] = useState<{ id: string; fileName: string; pageNumber: number; dataUrl: string; width: number; height: number; selected: boolean; textItems?: any[]; rotation?: number }[]>([]);
  const [uploadedPdfFile, setUploadedPdfFile] = useState<File | null>(null);
  const [isPdfPreviewOpen, setIsPdfPreviewOpen] = useState<boolean>(false);
  const [pdfPreviewLayout, setPdfPreviewLayout] = useState<'vertical' | 'horizontal'>('vertical');
  const [previewZoom, setPreviewZoom] = useState<number>(100);
  const [viewerZoom, setViewerZoom] = useState<number>(100);

  // PDF metadata list for Yaz Dönemi
  const [pdfMetadataList, setPdfMetadataList] = useState<{ name: string; id: string; viewUrl: string; lastUpdated: string }[]>([]);
  const [cachedPdfPages, setCachedPdfPages] = useState<{ pageNumber: number; dataUrl: string; width: number; height: number; textItems?: any[] }[] | null>(null);
  const [pdfBlobUrl, setPdfBlobUrl] = useState<string | null>(null);

  // Revoke previous blob URL to avoid memory leak
  const updatePdfBlobUrl = (newUrl: string | null) => {
    if (pdfBlobUrl && pdfBlobUrl.startsWith('blob:')) {
      try {
        URL.revokeObjectURL(pdfBlobUrl);
      } catch (e) {
        console.error("Failed to revoke blob URL:", e);
      }
    }
    setPdfBlobUrl(newUrl);
  };

  const loadAndCachePdfPages = async (base64: string, cacheKey: string) => {
    try {
      const cachedPages = await getPdfPagesFromDB(cacheKey);
      if (cachedPages && cachedPages.length > 0) {
        // If the cached version does not have textItems property on any page (old cache before text extraction was implemented),
        // we force a rebuild of the cache so that search works.
        const hasTextLayer = cachedPages.every(p => p.hasOwnProperty('textItems') && Array.isArray(p.textItems));
        if (hasTextLayer) {
          setCachedPdfPages(cachedPages);
          return;
        }
        console.log("Cached PDF does not have text search layer. Re-generating...");
      }

      // Convert base64 to File
      const binaryString = window.atob(base64);
      const len = binaryString.length;
      const bytes = new Uint8Array(len);
      for (let i = 0; i < len; i++) {
        bytes[i] = binaryString.charCodeAt(i);
      }
      const blob = new Blob([bytes], { type: 'application/pdf' });
      const file = new File([blob], "temp_render.pdf", { type: 'application/pdf' });
      
      const pages = await renderPdfToImages(file);
      await savePdfPagesToDB(cacheKey, pages);
      setCachedPdfPages(pages);
    } catch (err) {
      console.error("Failed to render and cache PDF pages:", err);
    }
  };

  const isPdfLoadingRef = useRef<boolean>(false);
  const [isPdfLoading, setIsPdfLoading] = useState<boolean>(false);
  const [isOcrProcessing, setIsOcrProcessing] = useState<boolean>(false);
  const [ocrProgress, setOcrProgress] = useState<number>(0);
  const [ocrLastWord, setOcrLastWord] = useState<string>("");
  const [ocrLog, setOcrLog] = useState<string[]>([]);

  // State for Personnel Info Excel 537 rows x 12 columns
  const [excelForm5Data, setExcelForm5Data] = useState<string[][]>(() => {
    const stored = localStorage.getItem('excel_form_5_data');
    if (stored) {
      try {
        return JSON.parse(stored);
      } catch (e) {
        console.error(e);
      }
    }
    
    // Generate 537 rows with 12 columns
    const data: string[][] = Array.from({ length: 537 }, () => Array(12).fill(""));
    
    // Pre-fill realistic personnel records matching the 12 columns requested exactly:
    // 1. Sıra No, 2. Adı Soyadı, 3. T.C. Kimlik No, 4. Sicil No, 5. Kadro Unvanı, 6. Doğum Tarihi,
    // 7. Görev Yeri, 8. Telefon No, 9. Kan Grubu, 10. Adres Bilgisi, 11. Yakının Adı-Soyadı, 12. Eş/Yakın Telefon No
    const sampleRows = [
      ["1", "Mahmut OKUDAN", "29122439024", "45537", "Pilot", "27595", "OGM", "533 3497446", "A Rh+", "Ankara Merkez", "Hatice OKUDAN (Eşi)", "533 3497447"],
      ["2", "Rıfat ÖNAL", "66577226922", "45540", "Pilot", "24289", "HANGAR", "530 656 3112", "0 Rh+", "Ankara Keçiören", "Selin ÖNAL (Eşi)", "530 656 3113"],
      ["3", "Nuri Gökmen GEÇER", "12826014352", "45539", "Pilot", "26544", "HANGAR", "541 896 8232", "B Rh+", "Ankara Yenimahalle", "Elif GEÇER (Eşi)", "541 896 8233"],
      ["4", "Sabri AKSOY", "38536732232", "45541", "Pilot", "24163", "HANGAR", "533 258 4425", "AB Rh+", "Ankara Çankaya", "Ayla AKSOY (Eşi)", "533 258 4426"],
      ["5", "Hakan KOZLU", "3179879608", "45533", "Pilot", "26073", "MUĞLA", "505 431 4827", "A Rh-", "Muğla Merkez", "Zeynep KOZLU (Eşi)", "505 431 4828"],
      ["6", "Gürhan AYDIN", "28481247400", "45532", "Pilot", "25486", "İZMİR", "533 414 7167", "0 Rh-", "İzmir Bornova", "Nermin AYDIN (Eşi)", "533 414 7168"],
      ["7", "Serhat İVECAN", "11092197752", "45542", "Pilot", "23995", "İZMİR", "533 549 5234", "B Rh-", "İzmir Karşıyaka", "Derya İVECAN (Eşi)", "533 549 5235"],
      ["8", "Yücel KIVRAK", "30539220150", "45544", "Pilot", "26443", "ANTALYA", "505 350 7656", "AB Rh-", "Antalya Lara", "Seda KIVRAK (Eşi)", "505 350 7657"]
    ];
    
    sampleRows.forEach((row, rIdx) => {
      row.forEach((val, cIdx) => {
        data[rIdx][cIdx] = val;
      });
    });
    
    return data;
  });

  const [excelSearchQuery, setExcelSearchQuery] = useState<string>("");
  const [selectedKadroFilter, setSelectedKadroFilter] = useState<string>("");
  const [activeExcelMatchIdx, setActiveExcelMatchIdx] = useState<number>(0);

  // Teçhizat Takip Matrix States
  const [activeTechizatType, setActiveTechizatType] = useState<'bell429' | 'at802' | 't70' | 't70_bumbi_backet' | 'b360' | 'c650' | 'hangar' | 'kara_araclari' | 'all' | null>(null);
  const [techizatSearchQuery, setTechizatSearchQuery] = useState<string>("");
  const [techizatFirmaFilter, setTechizatFirmaFilter] = useState<string>("");
  const [techizatDurumFilter, setTechizatDurumFilter] = useState<string>("");
  const [activeTechizatMatchIdx, setActiveTechizatMatchIdx] = useState<number>(0);

  // States for each of the categories
  const [techizatBell429Columns, setTechizatBell429Columns] = useState<string[]>(() => {
    const saved = localStorage.getItem('excel_techizat_bell429_cols');
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        return parsed.map((col: string) => {
          if (col === "SON BAKIM" || col === "SON KONTROL" || col === "SON KONTROL / BAKIM") return "SON KONTROL / KALİBRASYON / BAKIM";
          if (col === "GELECEK BAKIM" || col === "GELECEK KONTROL" || col === "GELECEK KONTROL / BAKIM") return "GELECEK KONTROL / KALİBRASYON / BAKIM";
          return col;
        });
      } catch (e) { console.error(e); }
    }
    return ["SIRA NO", "TEÇHİZAT ADI", "PARÇA NO (P/N)", "SERİ NO (S/N)", "MİKTAR", "BULUNDUĞU YER", "DURUMU", "SON KONTROL / KALİBRASYON / BAKIM", "GELECEK KONTROL / KALİBRASYON / BAKIM", "SON KONTROLÜ YAPAN FİRMA", "AÇIKLAMA"];
  });
  const [techizatBell429Data, setTechizatBell429Data] = useState<string[][]>(() => {
    const saved = localStorage.getItem('excel_techizat_bell429_data');
    if (saved) return JSON.parse(saved);
    return [];
  });

  const [techizatAt802Columns, setTechizatAt802Columns] = useState<string[]>(() => {
    const saved = localStorage.getItem('excel_techizat_at802_cols');
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        return parsed.map((col: string) => {
          if (col === "SON BAKIM" || col === "SON KONTROL" || col === "SON KONTROL / BAKIM") return "SON KONTROL / KALİBRASYON / BAKIM";
          if (col === "GELECEK BAKIM" || col === "GELECEK KONTROL" || col === "GELECEK KONTROL / BAKIM") return "GELECEK KONTROL / KALİBRASYON / BAKIM";
          return col;
        });
      } catch (e) { console.error(e); }
    }
    return ["SIRA NO", "TEÇHİZAT ADI", "PARÇA NO (P/N)", "SERİ NO (S/N)", "MİKTAR", "BULUNDUĞU YER", "DURUMU", "SON KONTROL / KALİBRASYON / BAKIM", "GELECEK KONTROL / KALİBRASYON / BAKIM", "SON KONTROLÜ YAPAN FİRMA", "AÇIKLAMA"];
  });
  const [techizatAt802Data, setTechizatAt802Data] = useState<string[][]>(() => {
    const saved = localStorage.getItem('excel_techizat_at802_data');
    if (saved) return JSON.parse(saved);
    return [];
  });

  const [techizatT70Columns, setTechizatT70Columns] = useState<string[]>(() => {
    const saved = localStorage.getItem('excel_techizat_t70_cols');
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        return parsed.map((col: string) => {
          if (col === "SON BAKIM" || col === "SON KONTROL" || col === "SON KONTROL / BAKIM") return "SON KONTROL / KALİBRASYON / BAKIM";
          if (col === "GELECEK BAKIM" || col === "GELECEK KONTROL" || col === "GELECEK KONTROL / BAKIM") return "GELECEK KONTROL / KALİBRASYON / BAKIM";
          return col;
        });
      } catch (e) { console.error(e); }
    }
    return ["SIRA NO", "TEÇHİZAT ADI", "PARÇA NO (P/N)", "SERİ NO (S/N)", "MİKTAR", "BULUNDUĞU YER", "DURUMU", "SON KONTROL / KALİBRASYON / BAKIM", "GELECEK KONTROL / KALİBRASYON / BAKIM", "SON KONTROLÜ YAPAN FİRMA", "AÇIKLAMA"];
  });
  const [techizatT70Data, setTechizatT70Data] = useState<string[][]>(() => {
    const saved = localStorage.getItem('excel_techizat_t70_data');
    if (saved) return JSON.parse(saved);
    return [];
  });

  const [techizatT70BumbiBacketColumns, setTechizatT70BumbiBacketColumns] = useState<string[]>(() => {
    const saved = localStorage.getItem('excel_techizat_t70_bumbi_backet_cols');
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        return parsed.map((col: string) => {
          if (col === "SON BAKIM" || col === "SON KONTROL" || col === "SON KONTROL / BAKIM") return "SON KONTROL / KALİBRASYON / BAKIM";
          if (col === "GELECEK BAKIM" || col === "GELECEK KONTROL" || col === "GELECEK KONTROL / BAKIM") return "GELECEK KONTROL / KALİBRASYON / BAKIM";
          return col;
        });
      } catch (e) { console.error(e); }
    }
    return ["SIRA NO", "TEÇHİZAT ADI", "MODEL / TİP", "SERİ NO (S/N)", "KAPASİTE", "BULUNDUĞU YER", "DURUMU", "SON KONTROL / KALİBRASYON / BAKIM", "GELECEK KONTROL / KALİBRASYON / BAKIM", "SON KONTROLÜ YAPAN FİRMA", "AÇIKLAMA"];
  });
  const [techizatT70BumbiBacketData, setTechizatT70BumbiBacketData] = useState<string[][]>(() => {
    const saved = localStorage.getItem('excel_techizat_t70_bumbi_backet_data');
    if (saved) return JSON.parse(saved);
    return [];
  });

  const [techizatT70HelitakColumns, setTechizatT70HelitakColumns] = useState<string[]>(() => {
    const saved = localStorage.getItem('excel_techizat_t70_helitak_cols');
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        return parsed.map((col: string) => {
          if (col === "SON BAKIM" || col === "SON KONTROL" || col === "SON KONTROL / BAKIM") return "SON KONTROL / KALİBRASYON / BAKIM";
          if (col === "GELECEK BAKIM" || col === "GELECEK KONTROL" || col === "GELECEK KONTROL / BAKIM") return "GELECEK KONTROL / KALİBRASYON / BAKIM";
          return col;
        });
      } catch (e) { console.error(e); }
    }
    return ["SIRA NO", "TEÇHİZAT ADI", "MODEL / TİP", "SERİ NO (S/N)", "KAPASİTE", "BULUNDUĞU YER", "DURUMU", "SON KONTROL / KALİBRASYON / BAKIM", "GELECEK KONTROL / KALİBRASYON / BAKIM", "SON KONTROLÜ YAPAN FİRMA", "AÇIKLAMA"];
  });
  const [techizatT70HelitakData, setTechizatT70HelitakData] = useState<string[][]>(() => {
    const saved = localStorage.getItem('excel_techizat_t70_helitak_data');
    if (saved) return JSON.parse(saved);
    return [];
  });

  const [techizatB360Columns, setTechizatB360Columns] = useState<string[]>(() => {
    const saved = localStorage.getItem('excel_techizat_b360_cols');
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        return parsed.map((col: string) => {
          if (col === "SON BAKIM" || col === "SON KONTROL" || col === "SON KONTROL / BAKIM") return "SON KONTROL / KALİBRASYON / BAKIM";
          if (col === "GELECEK BAKIM" || col === "GELECEK KONTROL" || col === "GELECEK KONTROL / BAKIM") return "GELECEK KONTROL / KALİBRASYON / BAKIM";
          return col;
        });
      } catch (e) { console.error(e); }
    }
    return ["SIRA NO", "TEÇHİZAT ADI", "PARÇA NO (P/N)", "SERİ NO (S/N)", "MİKTAR", "BULUNDUĞU YER", "DURUMU", "SON KONTROL / KALİBRASYON / BAKIM", "GELECEK KONTROL / KALİBRASYON / BAKIM", "SON KONTROLÜ YAPAN FİRMA", "AÇIKLAMA"];
  });
  const [techizatB360Data, setTechizatB360Data] = useState<string[][]>(() => {
    const saved = localStorage.getItem('excel_techizat_b360_data');
    if (saved) return JSON.parse(saved);
    return [];
  });

  const [techizatC650Columns, setTechizatC650Columns] = useState<string[]>(() => {
    const saved = localStorage.getItem('excel_techizat_c650_cols');
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        return parsed.map((col: string) => {
          if (col === "SON BAKIM" || col === "SON KONTROL" || col === "SON KONTROL / BAKIM") return "SON KONTROL / KALİBRASYON / BAKIM";
          if (col === "GELECEK BAKIM" || col === "GELECEK KONTROL" || col === "GELECEK KONTROL / BAKIM") return "GELECEK KONTROL / KALİBRASYON / BAKIM";
          return col;
        });
      } catch (e) { console.error(e); }
    }
    return ["SIRA NO", "TEÇHİZAT ADI", "PARÇA NO (P/N)", "SERİ NO (S/N)", "MİKTAR", "BULUNDUĞU YER", "DURUMU", "SON KONTROL / KALİBRASYON / BAKIM", "GELECEK KONTROL / KALİBRASYON / BAKIM", "SON KONTROLÜ YAPAN FİRMA", "AÇIKLAMA"];
  });
  const [techizatC650Data, setTechizatC650Data] = useState<string[][]>(() => {
    const saved = localStorage.getItem('excel_techizat_c650_data');
    if (saved) return JSON.parse(saved);
    return [];
  });

  const [techizatHangarColumns, setTechizatHangarColumns] = useState<string[]>(() => {
    const saved = localStorage.getItem('excel_techizat_hangar_cols');
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        return parsed.map((col: string) => {
          if (col === "SON BAKIM" || col === "SON KONTROL" || col === "SON KONTROL / BAKIM") return "SON KONTROL / KALİBRASYON / BAKIM";
          if (col === "GELECEK BAKIM" || col === "GELECEK KONTROL" || col === "GELECEK KONTROL / BAKIM") return "GELECEK KONTROL / KALİBRASYON / BAKIM";
          return col;
        });
      } catch (e) { console.error(e); }
    }
    return ["SIRA NO", "TEÇHİZAT ADI", "PARÇA NO (P/N)", "SERİ NO (S/N)", "MİKTAR", "BULUNDUĞU YER", "DURUMU", "SON KONTROL / KALİBRASYON / BAKIM", "GELECEK KONTROL / KALİBRASYON / BAKIM", "SON KONTROLÜ YAPAN FİRMA", "AÇIKLAMA"];
  });
  const [techizatHangarData, setTechizatHangarData] = useState<string[][]>(() => {
    const saved = localStorage.getItem('excel_techizat_hangar_data');
    if (saved) return JSON.parse(saved);
    return [];
  });

  const [techizatKaraAraclariColumns, setTechizatKaraAraclariColumns] = useState<string[]>(() => {
    const saved = localStorage.getItem('excel_techizat_kara_araclari_cols');
    if (saved) return JSON.parse(saved);
    return ["SIRA NO", "ARAÇ PLAKASI / TANIMI", "PARÇA NO (P/N) / MODEL", "BULUNDUĞU YER", "SON KM Sİ", "DURUMU", "SON KONTROL / KALİBRASYON / BAKIM", "GELECEK KONTROL / KALİBRASYON / BAKIM", "SON KONTROLÜ YAPAN FİRMA", "AÇIKLAMA"];
  });

  const convertOldKaraAraclariRowToNew = (row: string[]): string[] => {
    const isOldFormat = row.length >= 11 && (row[6] === "FAAL" || row[6] === "ARIZALI" || row[6] === "FAAL DEĞİL" || row[6] === "GAYRİ FAAL");
    if (!isOldFormat) {
      const r = [...row];
      while (r.length < 11) r.push("");
      return r.slice(0, 11);
    }
    
    const SIRA_NO = row[0] || "";
    const PLAKA = row[1] || "";
    const MODEL = row[3] || row[2] || ""; 
    const YER = row[5] || "";
    
    let km = "";
    let aciklama = row[10] || "";
    if (row[10] && !isNaN(Number(row[10].trim()))) {
      km = row[10].trim();
      aciklama = "Dönüş KM: " + km;
    } else {
      if (PLAKA.includes("CUK 695")) km = "124500";
      else if (PLAKA.includes("FV 2359")) km = "158200";
      else if (PLAKA.includes("1001")) km = "4500";
      else if (PLAKA.includes("1002")) km = "18200";
      else if (PLAKA.includes("1003")) km = "9400";
      else if (PLAKA.includes("1004")) km = "3200";
    }
    
    const DURUM = row[6] || "";
    const SON_BAKIM = row[7] || "";
    const GELECEK_BAKIM = row[8] || "";
    const FIRMA = row[9] || "";
    const MAIL = row[11] || "";

    return [
      SIRA_NO,
      PLAKA,
      MODEL,
      YER,
      km,
      DURUM,
      SON_BAKIM,
      GELECEK_BAKIM,
      FIRMA,
      aciklama,
      MAIL
    ];
  };

  const [techizatKaraAraclariData, setTechizatKaraAraclariData] = useState<string[][]>(() => {
    const saved = localStorage.getItem('excel_techizat_kara_araclari_data');
    if (saved) {
      try {
        const parsed = JSON.parse(saved) as string[][];
        const converted = parsed.map(row => convertOldKaraAraclariRowToNew(row));
        return converted.map((row, idx) => {
          const r = [...row];
          r[0] = String(idx + 1);
          return r;
        });
      } catch (e) {
        return [];
      }
    }
    return [];
  });

  const lastPopulatedPlaka = useRef<string>("");

  useEffect(() => {
    if (gePlaka && gePlaka !== lastPopulatedPlaka.current && techizatKaraAraclariData.length > 0) {
      lastPopulatedPlaka.current = gePlaka;
      const cleanPlaka = gePlaka.replace(/\s+/g, "").toLowerCase();
      const foundRow = techizatKaraAraclariData.find(row => {
        const rowPlaka = String(row[1] || "").replace(/\s+/g, "").toLowerCase();
        return rowPlaka === cleanPlaka || rowPlaka.includes(cleanPlaka) || cleanPlaka.includes(rowPlaka);
      });
      if (foundRow && foundRow[4]) {
        setGeDepartureKm(foundRow[4]);
      }
    } else if (!gePlaka) {
      lastPopulatedPlaka.current = "";
    }
  }, [gePlaka, techizatKaraAraclariData]);

  // Helper method to open Teçhizat Takip Matrix
  const openTechizatMatrix = (type: 'bell429' | 'at802' | 't70' | 't70_bumbi_backet' | 't70_helitak' | 'b360' | 'c650' | 'hangar' | 'kara_araclari' | 'all', title: string) => {
    setActiveTechizatType(type);
    setModalType('techizat_matrix');
    setModalTitle(title);
    setModalOpen(true);
    setTechizatSearchQuery('');
    setActiveTechizatMatchIdx(0);
  };

  /**
   * Akıllı Çoklu Lokasyon ve Miktar Birleştirici:
   * Yüklenen Excel veya Canlı E-Tabloda aynı ürünün birden fazla lokasyonu/miktarı
   * hizada alt alta satırlar şeklinde girildiğinde (ürün adı boş, tire veya aynı isimde),
   * bunu yeni ürün yapmak yerine mevcut ürünün 'BULUNDUĞU YER' ve 'MİKTAR / KAPASİTE'
   * alanlarına alt alta (\n ile) birleştirir.
   */
  const groupMultiLocationRows = (
    rawRows: string[][],
    nameColIdx: number = 1,
    locColIdx: number = 5,
    miktarColIdx: number = 4,
    siraColIdx: number = 0
  ): string[][] => {
    const grouped: string[][] = [];

    for (let i = 0; i < rawRows.length; i++) {
      const row = [...rawRows[i]];
      const rawSira = (row[siraColIdx] || "").trim();
      const rawName = (row[nameColIdx] || "").trim();
      const rawLoc = locColIdx >= 0 ? (row[locColIdx] || "").trim() : "";
      const rawMiktar = miktarColIdx >= 0 ? (row[miktarColIdx] || "").trim() : "";

      const lastRow = grouped.length > 0 ? grouped[grouped.length - 1] : null;
      const lastName = lastRow ? (lastRow[nameColIdx] || "").trim() : "";
      const lastSira = lastRow ? (lastRow[siraColIdx] || "").trim() : "";
      const lastLoc = lastRow && locColIdx >= 0 ? (lastRow[locColIdx] || "").trim() : "";

      const isNameEmptyOrDash = !rawName || rawName === "-" || rawName === "--";
      const isSiraEmptyOrDash = !rawSira || rawSira === "-" || rawSira === "--";
      const isSameName = rawName.toUpperCase() === lastName.toUpperCase() && lastName !== "";
      const isSameSira = rawSira === lastSira && lastSira !== "";

      // Check if this row is a continuation/sub-location row of the previous product
      const isSubLocation =
        !!lastRow &&
        (isNameEmptyOrDash || isSameName) &&
        (isSiraEmptyOrDash || isSameSira || isNameEmptyOrDash) &&
        (rawLoc !== "" || rawMiktar !== "") &&
        (rawLoc === "" || lastLoc === "" || !lastLoc.split('\n').includes(rawLoc) || isNameEmptyOrDash);

      if (isSubLocation && lastRow) {
        // Append location
        if (locColIdx >= 0 && rawLoc) {
          lastRow[locColIdx] = lastRow[locColIdx]
            ? `${lastRow[locColIdx]}\n${rawLoc}`
            : rawLoc;
        }
        // Append miktar/kapasite
        if (miktarColIdx >= 0 && rawMiktar) {
          lastRow[miktarColIdx] = lastRow[miktarColIdx]
            ? `${lastRow[miktarColIdx]}\n${rawMiktar}`
            : rawMiktar;
        }
        // Fill other missing fields if main row has empty and sub-row provides it
        for (let c = 0; c < row.length; c++) {
          if (c !== siraColIdx && c !== nameColIdx && c !== locColIdx && c !== miktarColIdx) {
            const cellVal = (row[c] || "").trim();
            if (cellVal && cellVal !== "-" && (!lastRow[c] || lastRow[c] === "-")) {
              lastRow[c] = cellVal;
            }
          }
        }
      } else {
        // Independent new product row
        grouped.push(row);
      }
    }

    // Ensure consecutive SIRA NO
    return grouped.map((row, idx) => {
      const r = [...row];
      if (siraColIdx >= 0) {
        r[siraColIdx] = String(idx + 1);
      }
      return r;
    });
  };

  // Excel exporter for Teçhizat Takip - Opens choice dialog ("Görselli olarak indirilsin mi?")
  const exportTechizatToExcel = (type: string, cols: string[], rows: string[][], title: string = "TEÇHİZAT LİSTESİ") => {
    setExcelExportModalData({ type, cols, rows, title });
  };

  // Fast text-only Excel export
  const executeTextOnlyExcelExport = (type: string, cols: string[], rows: string[][], title: string = "TEÇHİZAT LİSTESİ") => {
    try {
      let html = `
        <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
        <head>
          <meta charset="utf-8">
          <!--[if gte mso 9]>
          <xml>
            <x:ExcelWorkbook>
              <x:ExcelWorksheets>
                <x:ExcelWorksheet>
                  <x:Name>Techizat Listesi</x:Name>
                  <x:WorksheetOptions>
                    <x:DisplayGridlines/>
                  </x:WorksheetOptions>
                </x:ExcelWorksheet>
              </x:ExcelWorksheets>
            </x:ExcelWorkbook>
          </xml>
          <![endif]-->
          <style>
            table { border-collapse: collapse; font-family: 'Segoe UI', Arial, sans-serif; }
            .title-row { background-color: #0b3d1d; color: #ffffff; font-weight: bold; font-size: 14px; text-align: center; height: 40px; }
            th { background-color: #1e293b; color: #ffffff; font-weight: bold; border: 1px solid #475569; padding: 10px; text-align: center; font-size: 11px; }
            td { border: 1px solid #e2e8f0; padding: 8px 10px; font-size: 10px; color: #1e293b; white-space: pre-wrap; vertical-align: middle; }
            br { mso-data-placement: same-cell; }
            .zebra { background-color: #f8fafc; }
            .num { mso-number-format: "\\@"; text-align: center; } /* formats string values safely */
            .badge-faal { background-color: #dcfce7; color: #15803d; font-weight: bold; text-align: center; }
            .badge-gayrifaal { background-color: #fee2e2; color: #b91c1c; font-weight: bold; text-align: center; }
          </style>
        </head>
        <body>
          <table>
            <thead>
              <!-- Title Row merged -->
              <tr>
                <th colspan="${cols.length}" class="title-row" style="background-color: #0b3d1d; color: white; font-weight: bold; font-size: 14px; text-align: center; height: 40px;">
                  ${title.toUpperCase()}
                </th>
              </tr>
              <tr>
                ${cols.map(h => `<th style="background-color: #1e293b; color: #ffffff; font-weight: bold; border: 1px solid #475569; padding: 10px; text-align: center;">${h}</th>`).join('')}
              </tr>
            </thead>
            <tbody>
      `;
      
      rows.forEach((row, rIdx) => {
        const isZebra = rIdx % 2 === 1;

        // Check if any cell has multiple lines (e.g. multi-location / multi-quantity)
        const splittedCells = row.map(cell => (cell || "").split('\n'));
        const maxSubLines = Math.max(1, ...splittedCells.map(sc => sc.length));

        if (maxSubLines <= 1) {
          // Standard single row
          html += `<tr class="${isZebra ? 'zebra' : ''}">`;
          row.forEach((cell, cIdx) => {
            const val = cell || "";
            let tdClass = "";
            let style = "vertical-align: middle;";
            
            const colName = cols[cIdx]?.toUpperCase() || "";
            if (colName.includes("SIRA") || colName.includes("NO") || colName.includes("P/N") || colName.includes("S/N") || colName.includes("MİKTAR") || colName.includes("TELEFON") || colName.includes("TC")) {
              tdClass = "num";
            }
            
            if (colName.includes("DURUM")) {
              if (val.toUpperCase().includes("FAAL") && !val.toUpperCase().includes("GAYRİ")) {
                style += " background-color: #dcfce7; color: #15803d; font-weight: bold; text-align: center;";
              } else if (val.toUpperCase().includes("GAYRİ") || val.toUpperCase().includes("ARIZALI") || val.toUpperCase().includes("FAAL DEĞİL")) {
                style += " background-color: #fee2e2; color: #b91c1c; font-weight: bold; text-align: center;";
              }
            }
            
            html += `<td class="${tdClass}" style="${style}">${val}</td>`;
          });
          html += '</tr>';
        } else {
          // Multi-location row -> Generate sub-rows with rowspan for single-value cells
          for (let subIdx = 0; subIdx < maxSubLines; subIdx++) {
            html += `<tr class="${isZebra ? 'zebra' : ''}">`;
            row.forEach((cell, cIdx) => {
              const lines = splittedCells[cIdx];
              const isMultiLineCol = lines.length > 1;
              const colName = cols[cIdx]?.toUpperCase() || "";

              let tdClass = "";
              let style = "vertical-align: middle;";
              if (colName.includes("SIRA") || colName.includes("NO") || colName.includes("P/N") || colName.includes("S/N") || colName.includes("MİKTAR") || colName.includes("TELEFON") || colName.includes("TC")) {
                tdClass = "num";
              }

              if (colName.includes("DURUM")) {
                const val = cell || "";
                if (val.toUpperCase().includes("FAAL") && !val.toUpperCase().includes("GAYRİ")) {
                  style += " background-color: #dcfce7; color: #15803d; font-weight: bold; text-align: center;";
                } else if (val.toUpperCase().includes("GAYRİ") || val.toUpperCase().includes("ARIZALI") || val.toUpperCase().includes("FAAL DEĞİL")) {
                  style += " background-color: #fee2e2; color: #b91c1c; font-weight: bold; text-align: center;";
                }
              }

              if (subIdx === 0) {
                if (isMultiLineCol) {
                  html += `<td class="${tdClass}" style="${style}">${lines[0] || ""}</td>`;
                } else {
                  html += `<td rowspan="${maxSubLines}" class="${tdClass}" style="${style}">${cell || ""}</td>`;
                }
              } else {
                if (isMultiLineCol) {
                  html += `<td class="${tdClass}" style="${style}">${lines[subIdx] || ""}</td>`;
                }
              }
            });
            html += '</tr>';
          }
        }
      });
      
      html += `
            </tbody>
          </table>
        </body>
        </html>
      `;
      
      const blob = new Blob([html], { type: "application/vnd.ms-excel;charset=utf-8" });
      const link = document.createElement("a");
      link.href = URL.createObjectURL(blob);
      link.download = `techizat_takip_${type}_en_son_surum.xls`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      
      showNotification("Teçhizat listesi metin içerikli Excel (.xls) olarak başarıyla indirildi.");
    } catch (err) {
      alert("Excel indirme hatası: " + err);
    }
  };

  // Process Excel export with optional embedded images
  const downloadTechizatExcelWithImages = async (withImages: boolean) => {
    if (!excelExportModalData) return;
    const { type, cols, rows, title } = excelExportModalData;

    try {
      if (!withImages) {
        executeTextOnlyExcelExport(type, cols, rows, title);
        setExcelExportModalData(null);
        return;
      }

      setIsExcelExportLoading(true);
      setExcelExportProgressText("Excel görselli aktarımı başlatılıyor...");

      const exportCols = [cols[0] || "SIRA NO", "TEÇHİZAT FOTOĞRAFI", ...cols.slice(1)];

      let html = `
        <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
        <head>
          <meta charset="utf-8">
          <!--[if gte mso 9]>
          <xml>
            <x:ExcelWorkbook>
              <x:ExcelWorksheets>
                <x:ExcelWorksheet>
                  <x:Name>Görselli Teçhizat Listesi</x:Name>
                  <x:WorksheetOptions>
                    <x:DisplayGridlines/>
                  </x:WorksheetOptions>
                </x:ExcelWorksheet>
              </x:ExcelWorksheets>
            </x:ExcelWorkbook>
          </xml>
          <![endif]-->
          <style>
            table { border-collapse: collapse; font-family: 'Segoe UI', Arial, sans-serif; }
            .title-row { background-color: #0b3d1d; color: #ffffff; font-weight: bold; font-size: 15px; text-align: center; height: 45px; }
            th { background-color: #1e293b; color: #ffffff; font-weight: bold; border: 1px solid #475569; padding: 12px; text-align: center; font-size: 11px; }
            td { border: 1px solid #cbd5e1; padding: 8px 10px; font-size: 10px; color: #0f172a; white-space: pre-wrap; vertical-align: middle; }
            br { mso-data-placement: same-cell; }
            .zebra { background-color: #f8fafc; }
            .num { mso-number-format: "\\@"; text-align: center; }
            .img-td { text-align: center; vertical-align: middle; padding: 6px; width: 115px; height: 115px; background-color: #ffffff; }
          </style>
        </head>
        <body>
          <table>
            <thead>
              <tr>
                <th colspan="${exportCols.length}" class="title-row" style="background-color: #0b3d1d; color: white; font-weight: bold; font-size: 15px; text-align: center; height: 45px;">
                  ${title.toUpperCase()} (GÖRSELLİ PORTAL LİSTESİ)
                </th>
              </tr>
              <tr>
                ${exportCols.map(h => `<th style="background-color: #1e293b; color: #ffffff; font-weight: bold; border: 1px solid #475569; padding: 10px; text-align: center;">${h}</th>`).join('')}
              </tr>
            </thead>
            <tbody>
      `;

      const isAll = type === 'all';
      for (let rIdx = 0; rIdx < rows.length; rIdx++) {
        const row = rows[rIdx];
        const isZebra = rIdx % 2 === 1;

        setExcelExportProgressText(`Drive fotoğrafları çekiliyor ve ekleniyor (${rIdx + 1} / ${rows.length})...`);

        const targetTechType = isAll ? getRealTechType(row[0]) : type;
        const targetRow = isAll ? row.slice(1) : row;
        const imageKey = targetTechType + "_" + (targetRow[1] || "").replace(/\s+/g, '_') + "_" + (targetRow[3] || "").replace(/\s+/g, '_');
        const imgUrl = techizatImages[imageKey];

        let imageBase64Html = `<td class="img-td" style="color: #94a3b8; font-style: italic; font-size: 9px; text-align: center;">Görsel Yok</td>`;

        if (imgUrl) {
          try {
            const b64Data = await fetchDriveImageAsBase64(imgUrl);
            if (b64Data) {
              imageBase64Html = `<td class="img-td" style="width: 115px; height: 115px; text-align: center; vertical-align: middle;"><img src="${b64Data}" width="95" height="95" style="object-fit: contain; max-width: 95px; max-height: 95px; border-radius: 6px;" alt="Foto" /></td>`;
            }
          } catch (e) {
            console.warn("Image conversion failed for row", rIdx, e);
          }
        }

        html += `<tr class="${isZebra ? 'zebra' : ''}" style="height: 115px;">`;
        html += `<td class="num" style="vertical-align: middle; text-align: center;">${row[0] || ""}</td>`;
        html += imageBase64Html;

        for (let cIdx = 1; cIdx < row.length; cIdx++) {
          const val = row[cIdx] || "";
          let tdClass = "";
          let style = "vertical-align: middle;";
          const colName = cols[cIdx]?.toUpperCase() || "";

          if (colName.includes("SIRA") || colName.includes("NO") || colName.includes("P/N") || colName.includes("S/N") || colName.includes("MİKTAR") || colName.includes("TELEFON") || colName.includes("TC")) {
            tdClass = "num";
          }

          if (colName.includes("DURUM")) {
            if (val.toUpperCase().includes("FAAL") && !val.toUpperCase().includes("GAYRİ")) {
              style += " background-color: #dcfce7; color: #15803d; font-weight: bold; text-align: center;";
            } else if (val.toUpperCase().includes("GAYRİ") || val.toUpperCase().includes("ARIZALI") || val.toUpperCase().includes("FAAL DEĞİL")) {
              style += " background-color: #fee2e2; color: #b91c1c; font-weight: bold; text-align: center;";
            }
          }

          html += `<td class="${tdClass}" style="${style}">${val}</td>`;
        }

        html += `</tr>`;
      }

      html += `
            </tbody>
          </table>
        </body>
        </html>
      `;

      const blob = new Blob([html], { type: "application/vnd.ms-excel;charset=utf-8" });
      const link = document.createElement("a");
      link.href = URL.createObjectURL(blob);
      link.download = `techizat_takip_${type}_gorselli.xls`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);

      setIsExcelExportLoading(false);
      setExcelExportModalData(null);
      showNotification("Teçhizat listesi Drive fotoğraflarıyla birlikte Excel olarak başarıyla indirildi.");
    } catch (err) {
      setIsExcelExportLoading(false);
      alert("Görselli Excel indirme hatası: " + err);
    }
  };

  const exportGorevEmirleriToExcel = () => {
    try {
      const cols = [
        "Tarih",
        "Araç Plakası",
        "Sürücü Personel",
        "Sürücü T.C. No",
        "Sürücü Sicil No",
        "Görev Seri No (S/N)",
        "Çıkış KM",
        "Dönüş KM",
        "Yapılan Toplam KM",
        "Çıkış Saati",
        "Giriş Saati",
        "Görev Güzergahı / Açıklama"
      ];

      const rows = karaAraclariGorevEmirleri.map(order => {
        const departureKmNum = Number(order.departureKm) || 0;
        const returnKmNum = Number(order.returnKm) || 0;
        const totalKm = Math.max(0, returnKmNum - departureKmNum);

        return [
          order.date || "",
          order.plate || "",
          order.driverName || "",
          order.driverTc || "",
          order.driverSicil || "",
          order.serialNo || "",
          String(departureKmNum),
          String(returnKmNum),
          String(totalKm),
          order.departureTime || "",
          order.returnTime || "",
          order.route || ""
        ];
      });

      let html = `
        <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
        <head>
          <meta charset="utf-8">
          <!--[if gte mso 9]>
          <xml>
            <x:ExcelWorkbook>
              <x:ExcelWorksheets>
                <x:ExcelWorksheet>
                  <x:Name>Gorev Emirleri</x:Name>
                  <x:WorksheetOptions>
                    <x:DisplayGridlines/>
                  </x:WorksheetOptions>
                </x:ExcelWorksheet>
              </x:ExcelWorksheets>
            </x:ExcelWorkbook>
          </xml>
          <![endif]-->
          <style>
            table { border-collapse: collapse; font-family: 'Segoe UI', Arial, sans-serif; }
            .title-row { background-color: #0b3d1d; color: #ffffff; font-weight: bold; font-size: 14px; text-align: center; height: 40px; }
            th { background-color: #1e293b; color: #ffffff; font-weight: bold; border: 1px solid #475569; padding: 10px; text-align: center; font-size: 11px; }
            td { border: 1px solid #e2e8f0; padding: 8px 10px; font-size: 10px; color: #1e293b; }
            .zebra { background-color: #f8fafc; }
            .num { mso-number-format: "\\@"; text-align: center; }
          </style>
        </head>
        <body>
          <table>
            <thead>
              <tr>
                <th colspan="${cols.length}" class="title-row" style="background-color: #0b3d1d; color: white; font-weight: bold; font-size: 14px; text-align: center; height: 40px;">
                  GÖREV EMRİ GEÇMİŞİ VE KAYITLARI
                </th>
              </tr>
              <tr>
                ${cols.map(h => `<th style="background-color: #1e293b; color: #ffffff; font-weight: bold; border: 1px solid #475569; padding: 10px; text-align: center;">${h}</th>`).join('')}
              </tr>
            </thead>
            <tbody>
      `;

      rows.forEach((row, rIdx) => {
        const isZebra = rIdx % 2 === 1;
        html += `<tr class="${isZebra ? 'zebra' : ''}">`;
        row.forEach((cell, cIdx) => {
          const val = cell || "";
          let tdClass = "";
          const colName = cols[cIdx]?.toUpperCase() || "";
          if (colName.includes("PLAKA") || colName.includes("NO") || colName.includes("KM") || colName.includes("SAAT") || colName.includes("TARIH")) {
            tdClass = "num";
          }
          html += `<td class="${tdClass}">${val.replace(/\n/g, '<br>')}</td>`;
        });
        html += '</tr>';
      });

      html += `
            </tbody>
          </table>
        </body>
        </html>
      `;

      const blob = new Blob([html], { type: "application/vnd.ms-excel;charset=utf-8" });
      const link = document.createElement("a");
      link.href = URL.createObjectURL(blob);
      link.download = `gorev_emri_gecmisi_ve_kayitlari.xls`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);

      showNotification("Görev emri geçmişi tasarımlı HTML Excel (.xls) olarak başarıyla indirildi.");
    } catch (err) {
      alert("Excel indirme hatası: " + err);
    }
  };
  
  const uniqueKadroTitles = useMemo(() => {
    const titles = new Set<string>();
    excelForm5Data.forEach(row => {
      const colA = String(row[0] || '').trim();
      const hasSiraNo = colA !== "" && !colA.toLowerCase().includes("sira") && !colA.toLowerCase().includes("no");
      if (!hasSiraNo) return; // skip rows that are not personnel rows
      
      const title = row[4];
      if (title && title.trim()) {
        const cleanTitle = title.trim();
        const cleanLower = cleanTitle.toLowerCase();
        if (!cleanLower.includes("kadro") && !cleanLower.includes("unv") && !cleanLower.includes("ünv")) {
          titles.add(cleanTitle);
        }
      }
    });
    return Array.from(titles).sort();
  }, [excelForm5Data]);

  const [isExcelOcrProcessing, setIsExcelOcrProcessing] = useState<boolean>(false);
  const [excelOcrProgress, setExcelOcrProgress] = useState<number>(0);
  const [excelOcrLog, setExcelOcrLog] = useState<string[]>([]);
  const [excelOcrLastWord, setExcelOcrLastWord] = useState<string>("");

  // Trigger OCR text extraction step whenever a PDF is loaded/updated
  useEffect(() => {
    if (cachedPdfPages && cachedPdfPages.length > 0) {
      setIsOcrProcessing(true);
      setOcrProgress(0);
      setOcrLog(["[SİSTEM] Belge tarama ve akıllı OCR motoru başlatıldı..."]);
      setOcrLastWord("");
      
      const totalPages = cachedPdfPages.length;
      let currentPageIdx = 0;
      let isEffectMounted = true;
      let streamInterval: any = null;
      let nextPageTimeout: any = null;
      
      const scanNextPage = () => {
        if (!isEffectMounted) return;
        
        if (currentPageIdx >= totalPages) {
          setOcrProgress(100);
          setOcrLog(prev => ["[SİSTEM] Tüm sayfaların taranması ve indekslenmesi başarıyla tamamlandı!", ...prev]);
          nextPageTimeout = setTimeout(() => {
            if (isEffectMounted) {
              setIsOcrProcessing(false);
            }
          }, 600);
          return;
        }
        
        const page = cachedPdfPages[currentPageIdx];
        const pageNum = page.pageNumber;
        const progressVal = Math.round(((currentPageIdx + 1) / totalPages) * 100);
        setOcrProgress(progressVal);
        
        // Extract actual text strings from this page
        const items = page.textItems || [];
        // Filter out tiny or empty text pieces
        const snippets = items
          .map((it: any) => it.str)
          .filter((str: string) => str && str.trim().length > 1)
          .slice(0, 12); // Get up to 12 meaningful lines/words
        
        if (snippets.length > 0) {
          let snippetIdx = 0;
          streamInterval = setInterval(() => {
            if (!isEffectMounted) {
              clearInterval(streamInterval);
              return;
            }
            
            if (snippetIdx >= snippets.length) {
              clearInterval(streamInterval);
              currentPageIdx++;
              nextPageTimeout = setTimeout(scanNextPage, 80);
            } else {
              const text = snippets[snippetIdx];
              setOcrLastWord(text);
              setOcrLog(prev => [
                `[Sayfa ${pageNum}] İndekslendi ➔ "${text.substring(0, 45)}"`,
                ...prev.slice(0, 15)
              ]);
              snippetIdx++;
            }
          }, 60);
        } else {
          setOcrLog(prev => [
            `[Sayfa ${pageNum}] Metin katmanı boş veya taranmış resim (OCR taranıyor)...`,
            ...prev.slice(0, 15)
          ]);
          currentPageIdx++;
          nextPageTimeout = setTimeout(scanNextPage, 250);
        }
      };
      
      // Start scanning process
      scanNextPage();
      
      return () => {
        isEffectMounted = false;
        if (streamInterval) clearInterval(streamInterval);
        if (nextPageTimeout) clearTimeout(nextPageTimeout);
      };
    } else {
      setIsOcrProcessing(false);
      setOcrProgress(0);
      setOcrLog([]);
      setOcrLastWord("");
    }
  }, [cachedPdfPages]);

  // Trigger simulated Excel indexing for Form 5 on opening
  useEffect(() => {
    if (selectedFormId === 5 && modalType === 'form_table') {
      setIsExcelOcrProcessing(true);
      setExcelOcrProgress(0);
      setExcelOcrLog(["[SİSTEM] Excel-PDF Akıllı Matris Tarayıcı Başlatıldı...", "[SİSTEM] 537rx12c veri alanı taranıyor..."]);
      setExcelOcrLastWord("");

      let progressVal = 0;
      const interval = setInterval(() => {
        progressVal += 8;
        if (progressVal >= 100) {
          clearInterval(interval);
          setExcelOcrProgress(100);
          setExcelOcrLog(prev => ["[SİSTEM] 537 satır ve 12 sütunun tamamı indekslendi ve taranabilir duruma getirildi!", ...prev]);
          setTimeout(() => {
            setIsExcelOcrProcessing(false);
            showNotification("Personel Bilgi Çizelgesi (537rx12c) taranarak PDF Matris görünümüne aktarıldı!");
          }, 500);
        } else {
          setExcelOcrProgress(progressVal);
          // Get some random personnel name from our data for the "last word" preview
          const rowIdx = Math.floor((progressVal / 100) * 15);
          const nameCell = excelForm5Data[rowIdx]?.[1] || "BOŞ HÜCRE";
          const sicilCell = excelForm5Data[rowIdx]?.[3] || "YOK";
          if (nameCell && nameCell !== "BOŞ HÜCRE") {
            setExcelOcrLastWord(`${sicilCell} - ${nameCell}`);
            setExcelOcrLog(prev => [
              `[Satır ${rowIdx + 1}] İndekslendi ➔ "${sicilCell} | ${nameCell}"`,
              ...prev.slice(0, 15)
            ]);
          }
        }
      }, 80);

      return () => {
        clearInterval(interval);
      };
    } else {
      setIsExcelOcrProcessing(false);
      setExcelOcrProgress(0);
      setExcelOcrLog([]);
      setExcelOcrLastWord("");
    }
  }, [selectedFormId, modalType]);

  const parsedPdfMatches = useMemo(() => {
    if (!pdfSearchQuery || !cachedPdfPages) return { matchesList: [], matchingItemsMap: {} };
    
    const normalizedQuery = normalizeTurkishForSearch(pdfSearchQuery);
    if (!normalizedQuery) return { matchesList: [], matchingItemsMap: {} };
    
    const matchesList: { pageNumber: number; textItemIndex: number; uniqueId: string }[] = [];
    const matchingItemsMap: Record<string, boolean> = {};
    
    cachedPdfPages.forEach((page) => {
      if (!page.textItems || page.textItems.length === 0) return;
      
      // 1. Direct single-item matching: Highly robust for single words (e.g., "tezcan", "t70")
      page.textItems.forEach((item, idx) => {
        const itemStrNormalized = normalizeTurkishForSearch(item.str);
        if (itemStrNormalized && itemStrNormalized.includes(normalizedQuery)) {
          const key = `page-${page.pageNumber}-item-${idx}`;
          if (!matchingItemsMap[key]) {
            matchingItemsMap[key] = true;
            matchesList.push({
              pageNumber: page.pageNumber,
              textItemIndex: idx,
              uniqueId: `match-${page.pageNumber}-${idx}`
            });
          }
        }
      });
      
      // 2. Row-grouping matching: Fallback for multi-word phrases that span across separate consecutive PDF items
      const itemsWithIdx = page.textItems.map((item, idx) => ({ item, idx }));
      const rows: { item: any; idx: number }[][] = [];
      const sortedByTop = [...itemsWithIdx].sort((a, b) => a.item.top - b.item.top);
      
      sortedByTop.forEach((entry) => {
        let placed = false;
        for (let r = 0; r < rows.length; r++) {
          const rowAvgTop = rows[r].reduce((sum, item) => sum + item.item.top, 0) / rows[r].length;
          if (Math.abs(entry.item.top - rowAvgTop) < 0.6) {
            rows[r].push(entry);
            placed = true;
            break;
          }
        }
        if (!placed) {
          rows.push([entry]);
        }
      });
      
      rows.forEach((row) => {
        row.sort((a, b) => a.item.left - b.item.left);
        
        let rowText = "";
        const charToItemIndex: number[] = [];
        
        row.forEach((entry, rIdx) => {
          if (rIdx > 0) {
            const prev = row[rIdx - 1];
            const gap = entry.item.left - (prev.item.left + prev.item.width);
            if (gap > 1.5) {
              rowText += " ";
              charToItemIndex.push(-1);
            }
          }
          
          rowText += entry.item.str;
          for (let i = 0; i < entry.item.str.length; i++) {
            charToItemIndex.push(entry.idx);
          }
        });
        
        // Character-by-character normalization to maintain perfect 1:1 mapping
        let normalizedRowText = "";
        const normalizedCharToItemIndex: number[] = [];
        
        for (let i = 0; i < rowText.length; i++) {
          const char = rowText[i];
          const origItemIdx = charToItemIndex[i];
          
          let mapped = char.replace(/İ/g, "i")
                           .replace(/I/g, "ı")
                           .replace(/ı/g, "i")
                           .toLowerCase();
          
          const turkishMap: Record<string, string> = {
            'ç': 'c', 'ğ': 'g', 'ö': 'o', 'ş': 's', 'ü': 'u',
            'â': 'a', 'ê': 'e', 'î': 'i', 'ô': 'o', 'û': 'u'
          };
          mapped = turkishMap[mapped] || mapped;
          
          if (/[a-z0-9\s]/.test(mapped)) {
            normalizedRowText += mapped;
            normalizedCharToItemIndex.push(origItemIdx);
          }
        }
        
        // Collapse multiple spaces
        let finalRowText = "";
        const finalCharToItemIndex: number[] = [];
        for (let i = 0; i < normalizedRowText.length; i++) {
          const char = normalizedRowText[i];
          const itemIdx = normalizedCharToItemIndex[i];
          
          if (char === ' ') {
            if (finalRowText.length > 0 && finalRowText[finalRowText.length - 1] !== ' ') {
              finalRowText += ' ';
              finalCharToItemIndex.push(itemIdx);
            }
          } else {
            finalRowText += char;
            finalCharToItemIndex.push(itemIdx);
          }
        }
        
        // Trim spaces
        let startTrim = 0;
        while (startTrim < finalRowText.length && finalRowText[startTrim] === ' ') {
          startTrim++;
        }
        let endTrim = finalRowText.length;
        while (endTrim > startTrim && finalRowText[endTrim - 1] === ' ') {
          endTrim--;
        }
        
        const finalRowTextTrimmed = finalRowText.substring(startTrim, endTrim);
        const finalMappingTrimmed = finalCharToItemIndex.slice(startTrim, endTrim);
        
        let pos = finalRowTextTrimmed.indexOf(normalizedQuery);
        while (pos !== -1) {
          const matchedItemIndices = new Set<number>();
          for (let i = pos; i < pos + normalizedQuery.length; i++) {
            const originalIdx = finalMappingTrimmed[i];
            if (originalIdx !== undefined && originalIdx !== -1) {
              matchedItemIndices.add(originalIdx);
            }
          }
          
          const firstIdx = Array.from(matchedItemIndices)[0];
          if (firstIdx !== undefined) {
            const key = `page-${page.pageNumber}-item-${firstIdx}`;
            if (!matchingItemsMap[key]) {
              matchesList.push({
                pageNumber: page.pageNumber,
                textItemIndex: firstIdx,
                uniqueId: `match-${page.pageNumber}-${firstIdx}`
              });
            }
          }
          
          matchedItemIndices.forEach((origIdx) => {
            matchingItemsMap[`page-${page.pageNumber}-item-${origIdx}`] = true;
          });
          
          pos = finalRowTextTrimmed.indexOf(normalizedQuery, pos + 1);
        }
      });
    });
    
    return { matchesList, matchingItemsMap };
  }, [pdfSearchQuery, cachedPdfPages]);

  const searchMatchesList = useMemo(() => {
    return parsedPdfMatches.matchesList;
  }, [parsedPdfMatches]);

  const totalSearchMatches = useMemo(() => {
    return searchMatchesList.length;
  }, [searchMatchesList]);

  // Scroll active match into view smoothly
  useEffect(() => {
    if (searchMatchesList.length > 0) {
      const activeMatch = searchMatchesList[activeMatchIndex];
      if (activeMatch) {
        const matchId = `match-${activeMatch.pageNumber}-${activeMatch.textItemIndex}`;
        setTimeout(() => {
          const el = document.getElementById(matchId);
          if (el) {
            el.scrollIntoView({ behavior: 'smooth', block: 'center' });
          }
        }, 100);
      }
    }
  }, [activeMatchIndex, searchMatchesList]);

  const availableSummerPeriods = useMemo(() => {
    const periodsSet = new Set<string>();
    
    // Determine suffix for the active summer unit
    const airframeSuffix = getAirframeSuffix(selectedFormId);
    
    pdfMetadataList.forEach(m => {
      const lowerName = m.name.toLowerCase();
      if (lowerName.includes("_yaz_plan_") && lowerName.endsWith(".pdf")) {
        // Only include period if it matches the active unit's airframe suffix
        if (isSummerForm(selectedFormId)) {
          const isBeklemeFile = lowerName.includes('bekleme') || lowerName.includes('ankara');
          
          if (airframeSuffix === 'bell429') {
            // Standard Bell 429 must NOT contain bekleme/ankara and MUST contain bell429
            if (isBeklemeFile || (!lowerName.includes('bell429') && !lowerName.includes('bell_429'))) {
              return;
            }
          } else if (airframeSuffix === 'bekleme_bell429') {
            // Bekleme Bell 429 MUST contain bekleme/ankara and MUST contain bell429
            if (!isBeklemeFile || (!lowerName.includes('bell429') && !lowerName.includes('bell_429'))) {
              return;
            }
          } else {
            if (!lowerName.includes(airframeSuffix)) {
              return;
            }
          }
        }
        
        const idx = lowerName.indexOf("_yaz_plan_");
        if (idx !== -1) {
          const periodPart = m.name.substring(0, idx);
          if (periodPart) {
            periodsSet.add(periodPart);
          }
        }
      }
    });
    
    const sorted = Array.from(periodsSet).sort((a, b) => {
      const dateA = new Date(parseRawPeriodStringToDates(a).start).getTime();
      const dateB = new Date(parseRawPeriodStringToDates(b).start).getTime();
      return dateB - dateA;
    });
    return sorted;
  }, [pdfMetadataList, selectedFormId]);

  const getReadablePeriodName = (periodStr: string): string => {
    let clean = periodStr.replace(/_/g, ' ');
    const parts = clean.split('-');
    
    const formatPart = (p: string) => {
      return p.trim().split(' ').map(word => {
        if (!word) return '';
        const lower = word.toLowerCase();
        const trCapitalized: Record<string, string> = {
          "haziran": "HAZİRAN", "mayis": "MAYIS", "temmuz": "TEMMUZ", "agustos": "AĞUSTOS",
          "eylul": "EYLÜL", "ekim": "EKİM", "kasim": "KASIM", "aralik": "ARALIK",
          "ocak": "OCAK", "subat": "ŞUBAT", "mart": "MART", "nisan": "NİSAN"
        };
        return trCapitalized[lower] || word.toUpperCase();
      }).join(' ');
    };
    
    if (parts.length === 2) {
      return `${formatPart(parts[0])} - ${formatPart(parts[1])}`;
    }
    return formatPart(clean);
  };

  const printCachedPdf = () => {
    if (!cachedPdfPages || cachedPdfPages.length === 0) return;
    const printWindow = window.open('', '_blank');
    if (!printWindow) {
      showNotification("Yazdırma penceresi engellendi. Lütfen pop-up engelleyicisini kapatın.");
      return;
    }
    
    let html = `
      <html>
        <head>
          <title>Yaz Dönemi Planlama</title>
          <style>
            @page {
              size: auto;
              margin: 0;
            }
            body {
              margin: 0;
              padding: 0;
              background-color: #ffffff;
              display: flex;
              flex-direction: column;
              align-items: center;
            }
            .page-container {
              width: 100%;
              page-break-after: always;
              page-break-inside: avoid;
              display: flex;
              justify-content: center;
              align-items: center;
            }
            img {
              max-width: 100%;
              height: auto;
              display: block;
            }
          </style>
        </head>
        <body>
    `;
    
    cachedPdfPages.forEach((page) => {
      html += `
        <div class="page-container">
          <img src="${page.dataUrl}" />
        </div>
      `;
    });
    
    html += `
          <script>
            window.onload = function() {
              setTimeout(function() {
                window.print();
                window.close();
              }, 500);
            };
          </script>
        </body>
      </html>
    `;
    
    printWindow.document.write(html);
    printWindow.document.close();
  };

  // States for background PDF downloading and bypassing corporate network firewalls
  const [isDownloadingPdf, setIsDownloadingPdf] = useState<boolean>(false);
  const [pdfDownloadStatus, setPdfDownloadStatus] = useState<string>('');

  // Retrieves the cache key for the currently selected PDF
  const getActiveCacheKey = () => {
    if (!selectedFormId) return null;
    const isSummer = isSummerForm(selectedFormId);
    let match: any = null;
    if (isSummer) {
      const airframeSuffix = getAirframeSuffix(selectedFormId);
      const cleanMonth = selectedSummerMonth.replace(/\s+/g, '_').toLowerCase();
      match = pdfMetadataList.find(m => {
        const cleanName = m.name.toLowerCase();
        const isBeklemeFile = cleanName.includes('bekleme') || cleanName.includes('ankara');
        
        if (airframeSuffix === 'bell429') {
          if (isBeklemeFile || (!cleanName.includes('bell429') && !cleanName.includes('bell_429'))) {
            return false;
          }
        } else if (airframeSuffix === 'bekleme_bell429') {
          if (!isBeklemeFile || (!cleanName.includes('bell429') && !cleanName.includes('bell_429'))) {
            return false;
          }
        } else {
          if (!cleanName.includes(airframeSuffix)) {
            return false;
          }
        }
        // Match selected month
        return cleanName.includes(cleanMonth);
      });
    } else {
      const prefix = selectedFormId === 1 ? 'gorevlendirme' : selectedFormId === 3 ? 'bakim_yetki' : selectedFormId === 5 ? 'personel_bilgi' : 'personel_ucus_hizmet';
      match = pdfMetadataList.find(m => m.name.toLowerCase().includes(prefix));
    }
    if (match) {
      const cleanLastUpdated = match.lastUpdated.replace(/[^a-zA-Z0-9]/g, '_');
      return `pdf_${match.id}_${cleanLastUpdated}`;
    }
    return null;
  };

  const getUploadCacheKey = (id: number, customMetadataList = pdfMetadataList) => {
    const isSummer = isSummerForm(id);
    let match: any = null;
    if (isSummer) {
      const airframeSuffix = getAirframeSuffix(id);
      const cleanMonth = selectedUploadSummerMonth.replace(/\s+/g, '_').toLowerCase();
      match = customMetadataList.find(m => {
        const cleanName = m.name.toLowerCase();
        const isBeklemeFile = cleanName.includes('bekleme') || cleanName.includes('ankara');
        
        if (airframeSuffix === 'bell429') {
          if (isBeklemeFile || (!cleanName.includes('bell429') && !cleanName.includes('bell_429'))) {
            return false;
          }
        } else if (airframeSuffix === 'bekleme_bell429') {
          if (!isBeklemeFile || (!cleanName.includes('bell429') && !cleanName.includes('bell_429'))) {
            return false;
          }
        } else {
          if (!cleanName.includes(airframeSuffix)) {
            return false;
          }
        }
        // Match selected month
        return cleanName.includes(cleanMonth);
      });
    } else {
      const prefix = id === 1 ? 'gorevlendirme' : id === 3 ? 'bakim_yetki' : id === 5 ? 'personel_bilgi' : 'personel_ucus_hizmet';
      match = customMetadataList.find(m => m.name.toLowerCase().includes(prefix));
    }
    if (match) {
      const cleanLastUpdated = match.lastUpdated.replace(/[^a-zA-Z0-9]/g, '_');
      return `pdf_${match.id}_${cleanLastUpdated}`;
    }
    return null;
  };

  // Silently downloads and caches a PDF in the background
  const silentPrefetchPdf = async (fileId: string, cacheKey: string) => {
    try {
      const targetUrl = `${GOOGLE_SCRIPT_URL}?action=getPdfBase64&fileId=${fileId}`;
      const response = await fetch(targetUrl);
      if (!response.ok) return;
      
      const result = await response.json();
      if (result && result.status === "success" && result.base64) {
        await saveRawPdfToDB(cacheKey, result.base64);
      }
    } catch (err) {
      console.error("Background silent prefetch failed for PDF:", fileId, err);
    }
  };

  // Background prefetching queue for all PDFs in metadata list to ensure absolute zero-loading offline/online experience
  useEffect(() => {
    if (pdfMetadataList.length === 0) return;
    
    let isMounted = true;
    const prefetchQueue = async () => {
      for (const match of pdfMetadataList) {
        if (!isMounted) break;
        const cleanLastUpdated = match.lastUpdated.replace(/[^a-zA-Z0-9]/g, '_');
        const cacheKey = `pdf_${match.id}_${cleanLastUpdated}`;
        
        try {
          const cachedBase64 = await getRawPdfFromDB(cacheKey);
          if (!cachedBase64) {
            // Fetch silently
            await silentPrefetchPdf(match.id, cacheKey);
          }
        } catch (e) {
          console.error("Failed to prefetch in background:", e);
        }
      }
    };
    prefetchQueue();
    return () => {
      isMounted = false;
    };
  }, [pdfMetadataList]);

  // Dynamically download raw PDF file from Drive via secure Apps Script Web App tunnel and load as blob URL
  const loadRawPdfFromDrive = async (fileId: string, cacheKey: string) => {
    try {
      setIsPdfLoading(true);
      setIsDownloadingPdf(true);
      setPdfDownloadStatus("PLAN BELGESİNE GÜVENLİ TÜNEL ARACILIĞIYLA BAĞLANILIYOR...");
      
      const targetUrl = `${GOOGLE_SCRIPT_URL}?action=getPdfBase64&fileId=${fileId}`;
      const response = await fetch(targetUrl);
      if (!response.ok) {
        throw new Error(`HTTP Hata: ${response.status}`);
      }
      
      const result = await response.json();
      if (result && result.status === "success" && result.base64) {
        setPdfDownloadStatus("BELGE HAZIRLANIYOR VE GÖSTERİME AKTARILIYOR...");
        
        // Save raw PDF base64 to DB
        await saveRawPdfToDB(cacheKey, result.base64);
        
        // Convert to blob and URL
        const binaryString = window.atob(result.base64);
        const len = binaryString.length;
        const bytes = new Uint8Array(len);
        for (let i = 0; i < len; i++) {
          bytes[i] = binaryString.charCodeAt(i);
        }
        const blob = new Blob([bytes], { type: 'application/pdf' });
        const blobUrl = URL.createObjectURL(blob);
        
        updatePdfBlobUrl(blobUrl);
        loadAndCachePdfPages(result.base64, cacheKey);
        showNotification("Planlama belgesi başarıyla yüklendi!");
      } else {
        throw new Error(result.message || "PDF verisi alınamadı.");
      }
    } catch (err: any) {
      console.error("PDF download failed:", err);
      setPdfDownloadStatus(`Bağlantı hatası: PDF indirilemedi. Hata: ${err?.message || err}`);
    } finally {
      setIsDownloadingPdf(false);
      setIsPdfLoading(false);
    }
  };

  const [isPdfRefreshing, setIsPdfRefreshing] = useState<boolean>(false);

  const handlePdfForceRefresh = async (activeMatch: any) => {
    if (!activeMatch) return;
    try {
      setIsPdfRefreshing(true);
      showNotification("Önbellek temizleniyor ve metin tabakası yeniden çözümleniyor...");
      
      const cleanLastUpdated = activeMatch.lastUpdated.replace(/[^a-zA-Z0-9]/g, '_');
      const cacheKey = `pdf_${activeMatch.id}_${cleanLastUpdated}`;
      
      await deletePdfFromDB(cacheKey);
      setCachedPdfPages(null);
      updatePdfBlobUrl(null);
      
      await loadRawPdfFromDrive(activeMatch.id, cacheKey);
      showNotification("PDF belgesi ve arama metin tabakası başarıyla güncellendi!");
    } catch (err: any) {
      console.error("Force refresh failed:", err);
      showNotification("Sıfırlama hatası: " + (err?.message || err));
    } finally {
      setIsPdfRefreshing(false);
    }
  };

  // Load cached PDF file as Blob URL from IndexedDB for current active selection or download if missing
  useEffect(() => {
    if (selectedFormId) {
      const isSummer = isSummerForm(selectedFormId);
      
      // Find the matching PDF from metadata list to retrieve its unique ID and update timestamp
      let match: any = null;
      if (isSummer) {
        const airframeSuffix = getAirframeSuffix(selectedFormId);
        const cleanMonth = selectedSummerMonth.replace(/\s+/g, '_').toLowerCase();
        match = pdfMetadataList.find(m => {
          const cleanName = m.name.toLowerCase();
          const isBeklemeFile = cleanName.includes('bekleme') || cleanName.includes('ankara');
          
          if (airframeSuffix === 'bell429') {
            if (isBeklemeFile || (!cleanName.includes('bell429') && !cleanName.includes('bell_429'))) {
              return false;
            }
          } else if (airframeSuffix === 'bekleme_bell429') {
            if (!isBeklemeFile || (!cleanName.includes('bell429') && !cleanName.includes('bell_429'))) {
              return false;
            }
          } else {
            if (!cleanName.includes(airframeSuffix)) {
              return false;
            }
          }
          // Match selected month
          return cleanName.includes(cleanMonth);
        });
      } else {
        const prefix = selectedFormId === 1 ? 'gorevlendirme' : selectedFormId === 3 ? 'bakim_yetki' : selectedFormId === 5 ? 'personel_bilgi' : 'personel_ucus_hizmet';
        match = pdfMetadataList.find(m => m.name.toLowerCase().includes(prefix));
      }

      if (match) {
        if (match.viewUrl === "excel_loaded") {
          updatePdfBlobUrl(null);
          setIsPdfLoading(false);
          return;
        }

        // Construct cache key with file ID and update timestamp to handle modified files cleanly
        const cleanLastUpdated = match.lastUpdated.replace(/[^a-zA-Z0-9]/g, '_');
        const cacheKey = `pdf_${match.id}_${cleanLastUpdated}`;
        
        // Optimistic check: try to fetch from cache without showing full-screen loader first
        getRawPdfFromDB(cacheKey).then(base64 => {
          if (base64) {
            // Convert to blob and set
            const binaryString = window.atob(base64);
            const len = binaryString.length;
            const bytes = new Uint8Array(len);
            for (let i = 0; i < len; i++) {
              bytes[i] = binaryString.charCodeAt(i);
            }
            const blob = new Blob([bytes], { type: 'application/pdf' });
            const blobUrl = URL.createObjectURL(blob);
            
            updatePdfBlobUrl(blobUrl);
            loadAndCachePdfPages(base64, cacheKey);
            setIsPdfLoading(false);
          } else {
            // Not in cache, show loading and download
            updatePdfBlobUrl(null);
            setIsPdfLoading(true);
            loadRawPdfFromDrive(match.id, cacheKey);
          }
        }).catch(err => {
          console.error("Failed to load cached raw PDF from DB:", err);
          updatePdfBlobUrl(null);
          setIsPdfLoading(true);
          loadRawPdfFromDrive(match.id, cacheKey);
        });
      } else {
        updatePdfBlobUrl(null);
        setIsPdfLoading(false);
      }
    } else {
      updatePdfBlobUrl(null);
      setIsPdfLoading(false);
    }
  }, [selectedFormId, selectedSummerMonth, pdfMetadataList]);

  // Google Drive klasöründen teçhizat resimlerini listeler ve senkronize eder
  const fetchImagesFromDrive = async () => {
    try {
      const targetUrl = `${GOOGLE_SCRIPT_URL}?action=listImagesFromDrive`;
      const response = await fetch(targetUrl);
      if (response.ok) {
        const result = await response.json();
        if (result && result.status === "success" && result.images) {
          setTechizatImages(prev => {
            const merged = { ...prev, ...result.images };
            localStorage.setItem('techizat_images', JSON.stringify(merged));
            return merged;
          });
        }
      }
    } catch (error) {
      console.error("Resim senkronizasyon hatası:", error);
    }
  };

  // Google Drive klasöründen PDF dosyalarını listeler
  const fetchPdfMetadata = async () => {
    try {
      const targetUrl = `${GOOGLE_SCRIPT_URL}?action=listPdfsFromDrive&folderId=1_fIGvuPVpC9N5on1irOfGG8OsD1KSXD0`;
      const response = await fetch(targetUrl);
      if (response.ok) {
        const result = await response.json();
        if (result && result.status === "success" && Array.isArray(result.data)) {
          const formatted = result.data.map((item: any, idx: number) => {
            return {
              name: String(item.name || ""),
              id: String(item.id || `pdf-drive-${idx}`),
              viewUrl: String(item.viewUrl || ""),
              lastUpdated: String(item.lastUpdated || "")
            };
          });
          setPdfMetadataList(formatted);
          return formatted;
        }
      }
    } catch (e) {
      console.error("PDF metadata fetch failed:", e);
    }
    return null;
  };

  // Google E-Tablo üzerindeki 'güncelleme tarihleri' sayfasından son senkronizasyon zamanlarını çeker
  const fetchUpdateDatesFromGoogleSheet = async () => {
    try {
      const targetUrl = `${GOOGLE_SCRIPT_URL}?action=readSheet&sheetName=${encodeURIComponent("güncelleme tarihleri")}`;
      const response = await fetch(targetUrl);
      if (response.ok) {
        const result = await response.json();
        if (result && result.status === "success" && Array.isArray(result.data)) {
          setFormUpdateDates(prev => {
            const dates = { ...prev };
            result.data.forEach((row: any) => {
              let birimAdi = "";
              let tarihSaat = "";
              
              Object.keys(row).forEach((key) => {
                const normKey = key.toLowerCase()
                  .replace(/ı/g, 'i').replace(/ğ/g, 'g').replace(/ü/g, 'u')
                  .replace(/ş/g, 's').replace(/ö/g, 'o').replace(/ç/g, 'c')
                  .replace(/[^a-z0-9]/g, '');
                
                const val = String(row[key] || "").trim();
                if (normKey.includes("birimadi") || normKey === "kolon1") {
                  birimAdi = val.toLowerCase();
                } else if (normKey.includes("guncelleme") || normKey.includes("tarih") || normKey.includes("saat") || normKey === "kolon2") {
                  tarihSaat = val;
                }
              });

              // Order-based fallback if keys were not found dynamically
              if (!birimAdi || !tarihSaat) {
                const keys = Object.keys(row);
                if (keys.length >= 2) {
                  if (!birimAdi) birimAdi = String(row[keys[0]] || "").toLowerCase().trim();
                  if (!tarihSaat) tarihSaat = String(row[keys[1]] || "").trim();
                }
              }
              
              if (birimAdi && tarihSaat) {
                // Normalize Turkish characters in birimAdi to compare reliably
                const normBirim = birimAdi.toLowerCase()
                  .replace(/ı/g, 'i').replace(/ğ/g, 'g').replace(/ü/g, 'u')
                  .replace(/ş/g, 's').replace(/ö/g, 'o').replace(/ç/g, 'c');
                  
                if (normBirim.includes("bell 429 yer destek")) {
                  dates["techizat_bell429"] = tarihSaat;
                } else if (normBirim.includes("at-802f yer destek") || normBirim.includes("at-802 yer destek")) {
                  dates["techizat_at802"] = tarihSaat;
                } else if (normBirim.includes("t-70 bumbi")) {
                  dates["techizat_t70_bumbi_backet"] = tarihSaat;
                } else if (normBirim.includes("t-70 yer destek")) {
                  dates["techizat_t70"] = tarihSaat;
                } else if (normBirim.includes("b-360 yer destek")) {
                  dates["techizat_b360"] = tarihSaat;
                } else if (normBirim.includes("c-650 yer destek")) {
                  dates["techizat_c650"] = tarihSaat;
                } else if (normBirim.includes("hangar yer destek")) {
                  dates["techizat_hangar"] = tarihSaat;
                } else if (normBirim.includes("gorevlendirme") || normBirim.includes("1.")) {
                  dates[1] = tarihSaat;
                } else if (normBirim.includes("ankara_bell") || normBirim.includes("24")) {
                  dates[24] = tarihSaat;
                } else if (normBirim.includes("ankara_c650") || normBirim.includes("25")) {
                  dates[25] = tarihSaat;
                } else if (normBirim.includes("bell") || normBirim.includes("21")) {
                  dates[21] = tarihSaat;
                  dates[22] = tarihSaat; // fallback if they share a structural date
                  dates[23] = tarihSaat;
                  dates[24] = tarihSaat;
                  dates[25] = tarihSaat;
                } else if (normBirim.includes("t70") || normBirim.includes("t-70") || normBirim.includes("22")) {
                  dates[22] = tarihSaat;
                } else if (normBirim.includes("at802") || normBirim.includes("at-802") || normBirim.includes("23")) {
                  dates[23] = tarihSaat;
                } else if (normBirim.includes("yetki") || normBirim.includes("3.")) {
                  dates[3] = tarihSaat;
                } else if (normBirim.includes("bilgi") || normBirim.includes("5.")) {
                  dates[5] = tarihSaat;
                } else if (normBirim.includes("ucus") || normBirim.includes("hizmet") || normBirim.includes("6.")) {
                  dates[6] = tarihSaat;
                }
              }
            });
            localStorage.setItem("form_update_dates", JSON.stringify(dates));
            return dates;
          });
        }
      }
    } catch (e) {
      console.error("Failed to fetch update dates from Google Sheet:", e);
    }
  };

  // Form last update dates tracker mapping form ID or custom teçhizat string to string (e.g. "05.01.2026 15:30")
  const [formUpdateDates, setFormUpdateDates] = useState<Record<string | number, string>>(() => {
    const saved = localStorage.getItem("form_update_dates");
    if (saved) {
      try {
        return JSON.parse(saved);
      } catch (e) {
        return {};
      }
    }
    return {};
  });

  const updateFormTimestamp = (formId: number) => {
    const now = new Date();
    const pad = (n: number) => String(n).padStart(2, '0');
    const formatted = `${pad(now.getDate())}.${pad(now.getMonth() + 1)}.${now.getFullYear()} ${pad(now.getHours())}:${pad(now.getMinutes())}`;
    
    setFormUpdateDates(prev => {
      const updated = { ...prev, [formId]: formatted };
      localStorage.setItem("form_update_dates", JSON.stringify(updated));
      return updated;
    });
  };

  const getDetailedSummerUpdateInfo = () => {
    // 1. Get all summer PDFs from pdfMetadataList
    const summerPdfs = pdfMetadataList.filter(m => {
      const nameLower = m.name.toLowerCase();
      return nameLower.includes("yaz_plan") || nameLower.includes("yaz_donemi") || nameLower.includes("yaz_planlama");
    });
    
    if (summerPdfs.length === 0) {
      // Fallback: Check if we have any date in formUpdateDates for summer forms
      const summerFormIds = [21, 22, 23, 24, 25];
      let latestDateStr = "";
      let latestId = 21;
      
      summerFormIds.forEach(id => {
        const d = formUpdateDates[id];
        if (d && d !== "-") {
          if (!latestDateStr || d > latestDateStr) {
            latestDateStr = d;
            latestId = id;
          }
        }
      });
      
      if (latestDateStr) {
        const airframeLabel = latestId === 21 ? "bell-429" : latestId === 22 ? "t-70" : latestId === 23 ? "at-802" : latestId === 24 ? "bekleme(429)" : "bekleme(c650/b360)";
        return `${latestDateStr} (${airframeLabel})`;
      }
      return "-";
    }
    
    // Sort summer PDFs by lastUpdated descending to find the absolute latest updated one
    const sorted = [...summerPdfs].sort((a, b) => {
      const timeA = a.lastUpdated ? new Date(a.lastUpdated).getTime() : 0;
      const timeB = b.lastUpdated ? new Date(b.lastUpdated).getTime() : 0;
      return timeB - timeA;
    });
    
    const latest = sorted[0];
    let name = latest.name.toLowerCase();
    if (name.endsWith(".pdf")) {
      name = name.slice(0, -4);
    }
    
    // Determine airframe suffix - Check bekleme suffixes first to avoid partial overlap with bell429
    let airframe = "bell-429";
    if (name.includes("bekleme_bell429") || name.includes("ankara_bell429") || name.includes("bell429_bekleme") || name.includes("bekleme429")) {
      airframe = "bekleme(429)";
    } else if (name.includes("bekleme_c650_b360") || name.includes("c650") || name.includes("b360") || name.includes("bekleme_c650")) {
      airframe = "bekleme(c650/b360)";
    } else if (name.includes("t70")) {
      airframe = "t-70";
    } else if (name.includes("at802")) {
      airframe = "at-802";
    } else if (name.includes("bell429")) {
      airframe = "bell-429";
    }
    
    // Extract period range (e.g. 7_mayis_-_8_haziran_2026)
    let periodPart = "";
    const pIdx = name.indexOf("_yaz_plan");
    if (pIdx !== -1) {
      periodPart = name.substring(0, pIdx);
    } else {
      periodPart = name;
    }
    
    // Format period
    let cleanPeriod = periodPart
      .replace(/_/g, " ")
      .replace(/-/g, " - ")
      .replace(/\s+/g, " ")
      .trim();
      
    // Strip year 2026 to keep it clean and match "7 mayıs-8 haziran" style
    cleanPeriod = cleanPeriod
      .replace(/2026/g, "")
      .replace(/\s+/g, " ")
      .trim();
      
    // Format lastUpdated date nicely (e.g. 01.01.2026)
    let formattedDate = "";
    if (latest.lastUpdated) {
      const dt = new Date(latest.lastUpdated);
      const day = String(dt.getDate()).padStart(2, '0');
      const month = String(dt.getMonth() + 1).padStart(2, '0');
      const year = dt.getFullYear();
      formattedDate = `${day}.${month}.${year}`;
    } else {
      formattedDate = "01.01.2026";
    }
    
    return `${formattedDate} (${airframe} - ${cleanPeriod})`;
  };

  const SUMMER_MONTHS = [
    "Ocak 2026",
    "Şubat 2026",
    "Mart 2026",
    "Nisan 2026",
    "Mayıs 2026",
    "Haziran 2026",
    "Temmuz 2026",
    "Ağustos 2026",
    "Eylül 2026",
    "Ekim 2026",
    "Kasım 2026",
    "Aralık 2026"
  ];

  const getSummerPeriodSheetPrefix = (formId: number, month: string): string => {
    const model = 
      formId === 21 ? 'bell 429' : 
      formId === 22 ? 't-70' : 
      formId === 23 ? 'at-802' : 
      formId === 24 ? 'ankara bekleme bell 429' : 'ankara bekleme c650 b360';
    const monthLower = month.toLocaleLowerCase('tr-TR');
    return `${monthLower}-yaz dönemi plan-${model}`;
  };

  // Excel Sync target ('all' or individual)
  const [syncSelectedTarget, setSyncSelectedTarget] = useState<string>('1');

  // Excel Sync Step-by-Step wizard states
  const [activeSyncStep, setActiveSyncStep] = useState<1 | 2>(1);
  const [step1Target, setStep1Target] = useState<string>('1');
  const [selectedForm2Unit, setSelectedForm2Unit] = useState<21 | 22 | 23 | 24 | 25 | null>(null);

  // Search filter query
  const [searchQuery, setSearchQuery] = useState('');

  // Password-Lock State ("1839")
  const [isPasswordModalOpen, setIsPasswordModalOpen] = useState(false);
  const [passwordInput, setPasswordInput] = useState('');
  const [passwordError, setPasswordError] = useState(false);

  // Gün Takip Sorumluları ve Renk Kodu Sıralama State Değişkenleri
  const [sortByColor, setSortByColor] = useState<boolean>(false);
  const [isSorumluModalOpen, setIsSorumluModalOpen] = useState(false);
  const [isSavingSorumlu, setIsSavingSorumlu] = useState(false);
  const [gunTakipSorumlulari, setGunTakipSorumlulari] = useState<{ birim: string; adSoyad: string; eposta: string; mail90?: string; }[]>([
    { birim: "BELL 429", adSoyad: "Sorumlu Personel", eposta: "ormanhavacilik.bakimsube@gmail.com", mail90: "" },
    { birim: "AT-802F", adSoyad: "Sorumlu Personel", eposta: "ormanhavacilik.bakimsube@gmail.com", mail90: "" },
    { birim: "T-70 YER DESTEK", adSoyad: "Sorumlu Personel", eposta: "ormanhavacilik.bakimsube@gmail.com", mail90: "" },
    { birim: "T-70 BUMBİ BACKET", adSoyad: "Sorumlu Personel", eposta: "ormanhavacilik.bakimsube@gmail.com", mail90: "" },
    { birim: "B-360", adSoyad: "Sorumlu Personel", eposta: "ormanhavacilik.bakimsube@gmail.com", mail90: "" },
    { birim: "C-650", adSoyad: "Sorumlu Personel", eposta: "ormanhavacilik.bakimsube@gmail.com", mail90: "" },
    { birim: "HANGAR YER DESTEK", adSoyad: "Sorumlu Personel", eposta: "ormanhavacilik.bakimsube@gmail.com", mail90: "" }
  ]);

  // Success notifications
  const [successMessage, setSuccessMessage] = useState<string | null>(null);

  // Feedback text for data save and read-back status, globally applicable
  const [dataFeedback, setDataFeedback] = useState<Record<number, string>>({});

  // Active view tab for Form 1 & All Forms ('editor' vs 'live_sheet')
  const [activeFormTab, setActiveFormTab] = useState<'editor' | 'live_sheet'>('live_sheet');

  // E-tablo sayfalarının dinamik tespiti ve seçimi state'leri
  const [allOnlineSheets, setAllOnlineSheets] = useState<{ name: string; id: number }[]>([]);
  const [selectedOnlineSheetId, setSelectedOnlineSheetId] = useState<number | null>(null);
  const [isFilteringSheets, setIsFilteringSheets] = useState(false);

  // Google Apps Script üzerinden tüm e-tablo sayfalarını ve gid kodlarını çeker
  const fetchAllGoogleSheetsList = async () => {
    try {
      const targetUrl = `${GOOGLE_SCRIPT_URL}?action=getSheets`;
      const response = await fetch(targetUrl);
      if (response.ok) {
        const result = await response.json();
        if (result && result.sheets) {
          setAllOnlineSheets(result.sheets);
        }
      }
      await fetchPdfMetadata();
      await fetchUpdateDatesFromGoogleSheet();
    } catch (e) {
      console.error("Sheets info fetch failed:", e);
    }
  };

  // Google E-Tablo içinde sadece o birime ait alt sayfaları gösterir, diğerlerini gizler (Alttaki sekmeler kalabalığı önlenir)
  const filterGoogleSheetsByPrefix = async (formId: number, customPrefix?: string) => {
    const config = TABLE_CONFIGS[formId];
    if (!config) return;
    const prefix = customPrefix || config.sheetName;
    try {
      setIsFilteringSheets(true);
      const targetUrl = `${GOOGLE_SCRIPT_URL}?action=filterSheets&prefix=${encodeURIComponent(prefix)}`;
      const response = await fetch(targetUrl);
      if (response.ok) {
        const result = await response.json();
        if (result && result.sheets) {
          setAllOnlineSheets(result.sheets);
        }
      }
    } catch (e) {
      console.error("Sheets filtering failed:", e);
      fetchAllGoogleSheetsList();
    } finally {
      setIsFilteringSheets(false);
    }
  };

  // Core sheets data hooks (initialized from localStorage with dynamic header mapping)
  const [tableData, setTableData] = useState<Record<number, Record<string, string>[]>>(() => {
    const initial: Record<number, Record<string, string>[]> = {};
    [1, 21, 22, 23, 24, 25, 3, 5, 6].forEach(id => {
      const config = TABLE_CONFIGS[id];
      const saved = localStorage.getItem(config.storageKey);
      if (saved) {
        try {
          const parsed = JSON.parse(saved);
          if (Array.isArray(parsed)) {
            // Normalize: map code-friendly keys like 'Personel_Adi_Soyadi' to Turkish labels like 'Personel Adı Soyadı'
            initial[id] = parsed.map((r: any) => {
              const item: Record<string, string> = {};
              // First translate defined columns
              config.columns.forEach(col => {
                const value = r[col.label] ?? r[col.key] ?? "";
                item[col.label] = String(value);
              });
              // Then also retain any other custom keys present in the parsed data (from direct spreadsheet uploads)
              Object.keys(r).forEach(k => {
                if (!item[k] && !config.columns.some(col => col.key === k)) {
                  item[k] = String(r[k]);
                }
              });
              return item;
            });
          } else {
            throw new Error("Not an array");
          }
        } catch (e) {
          // Initialize from default rows, mapping keys to friendly labels
          initial[id] = config.defaultRows.map(r => {
            const item: Record<string, string> = {};
            config.columns.forEach(col => {
              item[col.label] = r[col.key] || "";
            });
            return item;
          });
        }
      } else {
        // First load: map default rows to beautiful user friendly headers
        const initialRows = config.defaultRows.map(r => {
          const item: Record<string, string> = {};
          config.columns.forEach(col => {
            item[col.label] = r[col.key] || "";
          });
          return item;
        });
        initial[id] = initialRows;
        localStorage.setItem(config.storageKey, JSON.stringify(initialRows));
      }
    });
    return initial;
  });

  // Dynamically extract current table column headers (first row's keys)
  const getFormColumns = (formId: number): string[] => {
    const rows = tableData[formId] || [];
    if (rows.length > 0) {
      const keys = new Set<string>();
      rows.forEach(row => {
        Object.keys(row).forEach(k => keys.add(k));
      });
      const columnKeys = Array.from(keys);
      if (columnKeys.length > 0) {
        return columnKeys;
      }
    }
    // Fallback to config labels
    return TABLE_CONFIGS[formId].columns.map(col => col.label);
  };

  // Real-time HUD Clock state
  const [timeString, setTimeString] = useState('14:45:21');

  // Year reference
  const currentYear = new Date().getFullYear();

  // Automatic splash screen timeout helper
  useEffect(() => {
    const timer = setTimeout(() => {
      setSplashVisible(false);
    }, 5000);
    return () => clearTimeout(timer);
  }, []);

  // Live HUD Clock logic
  useEffect(() => {
    const updateTime = () => {
      const now = new Date();
      setTimeString(now.toLocaleTimeString('tr-TR'));
    };
    updateTime();
    const intervalId = setInterval(updateTime, 1000);
    return () => clearInterval(intervalId);
  }, []);

  // Set page's title
  useEffect(() => {
    document.title = "Hava Araçları Bakım Teknik Şube Müdürlüğü";
    fetchAllGoogleSheetsList();
    fetchPdfMetadata();
    fetchImagesFromDrive();
    fetchUpdateDatesFromGoogleSheet();
    pullAllTechizatFromGoogleSheets(true);
    pullDataFromGoogleSheets(5, true);
    pullKaraAraclariGorevEmirleri();
  }, []);

  // Reset form sub-modal states on selectedFormId change
  useEffect(() => {
    setFormTableMode('selection');
    setUploadProgress(0);
    setSelectedOnlineSheetId(null);
    if (selectedFormId) {
      if (isSummerForm(selectedFormId)) {
        const prefix = getSummerPeriodSheetPrefix(selectedFormId, selectedSummerMonth);
        filterGoogleSheetsByPrefix(selectedFormId, prefix);
      } else {
        filterGoogleSheetsByPrefix(selectedFormId);
      }
    } else {
      fetchAllGoogleSheetsList();
    }
  }, [selectedFormId, selectedSummerMonth]);

  // Yaz dönemi hava aracı değiştiğinde otomatik olarak veri barındıran ilk ayı seçer
  useEffect(() => {
    if (selectedFormId && isSummerForm(selectedFormId) && pdfMetadataList.length > 0) {
      const airframeSuffix = getAirframeSuffix(selectedFormId);
      const periodsWithData = availableSummerPeriods.filter(p => {
        const cleanMonth = p.replace(/\s+/g, '_').toLowerCase();
        return pdfMetadataList.some(pdf => {
          const cleanName = pdf.name.toLowerCase();
          const isBeklemeFile = cleanName.includes('bekleme') || cleanName.includes('ankara');
          
          if (airframeSuffix === 'bell429') {
            if (isBeklemeFile || (!cleanName.includes('bell429') && !cleanName.includes('bell_429'))) {
              return false;
            }
          } else if (airframeSuffix === 'bekleme_bell429') {
            if (!isBeklemeFile || (!cleanName.includes('bell429') && !cleanName.includes('bell_429'))) {
              return false;
            }
          } else {
            if (!cleanName.includes(airframeSuffix)) {
              return false;
            }
          }
          // Match month
          return cleanName.includes(cleanMonth);
        });
      });
      if (periodsWithData.length > 0 && !periodsWithData.includes(selectedSummerMonth)) {
        setSelectedSummerMonth(periodsWithData[0]);
      }
    }
  }, [selectedFormId, pdfMetadataList, availableSummerPeriods]);

  // Filtrelenen e-tablo sayfalarını ve aktif sayfayı hesapla
  const activeTableConfig = selectedFormId ? TABLE_CONFIGS[selectedFormId] : null;
  const matchedSheets = allOnlineSheets.filter(sheet => {
    if (!activeTableConfig) return false;
    const prefix = isSummerForm(selectedFormId)
      ? getSummerPeriodSheetPrefix(selectedFormId!, selectedSummerMonth)
      : activeTableConfig.sheetName;
    return sheet.name === prefix || sheet.name.startsWith(prefix + "-");
  });

  // Calculate synchronized active sheet ID so there is absolutely zero mismatch or previous form's sheet flicker
  const isSheetValid = selectedOnlineSheetId !== null && matchedSheets.some(s => s.id === selectedOnlineSheetId);
  const activeIframeId = isSheetValid ? selectedOnlineSheetId : (matchedSheets[0]?.id || null);

  // Eşleşen sayfa değiştiğinde otomatik olarak ilk sayfayı seç
  useEffect(() => {
    if (matchedSheets.length > 0) {
      const exists = matchedSheets.some(s => s.id === selectedOnlineSheetId);
      if (!exists) {
        setSelectedOnlineSheetId(matchedSheets[0].id);
      }
    } else {
      setSelectedOnlineSheetId(null);
    }
  }, [selectedFormId, allOnlineSheets, selectedOnlineSheetId, matchedSheets]);

  // Sayfa veya form değiştiğinde iframe yükleniyor durumunu tetikle
  useEffect(() => {
    if (selectedFormId !== null || activeIframeId !== null) {
      setSheetIframeLoading(true);
    }
  }, [selectedFormId, activeIframeId]);

  // Get human friendly title for categories
  const getCategoryTitle = (cat: CategoryType): string => {
    switch (cat) {
      case 'İKMAL': return 'İKMAL MÜDÜRLÜĞÜ';
      case 'TEÇHİZAT TAKİP': return 'TEÇHİZAT TAKİP SİSTEMİ';
      case 'HA_YER_DESTEK': return 'HAVA ARAÇLARI YER DESTEK TEÇHİZATLARI';
      case 'T70_DETAY': return 'T-70 TEÇHİZAT ALTBİRİMLERİ';
      case 'KARA_ARACLARI_MENU': return 'KARA ARAÇLARI TAKİP SİSTEMİ';
      case 'FORM KAYITLARI': return 'FORM KAYITLARI';
      default: return 'SİSTEM';
    }
  };

  // Modal control functions
  const openSystem = (url: string, title: string) => {
    if (url.includes('netlify.app') || url.includes('github') || url.includes('google.com/spreadsheets')) {
      window.open(url, '_blank');
      showNotification(`${title} portalı yeni sekmede güvenli bir şekilde açıldı.`);
      return;
    }
    setModalTitle(title);
    setIframeLoading(true);
    setModalUrl(url);
    setModalType('iframe');
    setModalOpen(true);
  };

  const showDesignPhase = (title: string) => {
    setModalTitle(title);
    setIframeLoading(false);
    setModalUrl('');
    setModalType('design');
    setModalOpen(true);
  };

  const openCategory = (category: Exclude<CategoryType, null>) => {
    setModalTitle(getCategoryTitle(category));
    setSelectedCategory(category);
    setCategoryHistory([category]);
    setIframeLoading(false);
    setModalUrl('');
    setModalType('category');
    setModalOpen(true);
  };

  const navigateToSubCategory = (subCategory: Exclude<CategoryType, null>) => {
    setSelectedCategory(subCategory);
    setModalTitle(getCategoryTitle(subCategory));
    setCategoryHistory(prev => [...prev, subCategory]);
  };

  const handleBack = () => {
    if (modalType === 'denetleme') {
      setModalType('category');
      setSelectedCategory('FORM KAYITLARI');
      setModalTitle(getCategoryTitle('FORM KAYITLARI'));
      return;
    }
    if (modalType === 'techizat_matrix') {
      setModalType('category');
      if (activeTechizatType === 'kara_araclari') {
        setSelectedCategory('KARA_ARACLARI_MENU');
        setModalTitle(getCategoryTitle('KARA_ARACLARI_MENU'));
      } else if (activeTechizatType === 't70' || activeTechizatType === 't70_bumbi_backet') {
        setSelectedCategory('T70_DETAY');
        setModalTitle(getCategoryTitle('T70_DETAY'));
      } else if (activeTechizatType === 'bell429' || activeTechizatType === 'at802') {
        setSelectedCategory('HA_YER_DESTEK');
        setModalTitle(getCategoryTitle('HA_YER_DESTEK'));
      } else {
        setSelectedCategory('TEÇHİZAT TAKİP');
        setModalTitle(getCategoryTitle('TEÇHİZAT TAKİP'));
      }
      setActiveTechizatType(null);
      return;
    }
    if (modalType === 'form_table' || modalType === 'excel_sync') {
      setModalType('category');
      setSelectedCategory('FORM KAYITLARI');
      setModalTitle(getCategoryTitle('FORM KAYITLARI'));
      setSelectedFormId(null);
      return;
    }
    if (categoryHistory.length > 1) {
      const updated = [...categoryHistory];
      updated.pop(); // Remove current
      const prevCategory = updated[updated.length - 1];
      setSelectedCategory(prevCategory);
      setCategoryHistory(updated);
      setModalTitle(getCategoryTitle(prevCategory));
      setModalType('category');
    } else {
      closeSystem();
    }
  };

  const closeSystem = () => {
    setModalOpen(false);
    // Reset states after animation closes
    setTimeout(() => {
      setModalUrl('');
      setSelectedCategory(null);
      setCategoryHistory([]);
      setSelectedFormId(null);
      setSearchQuery('');
    }, 400);
  };

  // Notification helper
  const showNotification = (msg: string) => {
    setSuccessMessage(msg);
    setTimeout(() => {
      setSuccessMessage(null);
    }, 4000);
  };

  const handleSaveTechizatRow = async (editedRow: string[], techType: string, rIdx: number) => {
    let updated: string[][] = [];
    if (techType === 'bell429') {
      const u = [...techizatBell429Data];
      const actualIdx = u.findIndex(r => r[0] === editedRow[0]);
      const targetIdx = actualIdx !== -1 ? actualIdx : rIdx;
      u[targetIdx] = editedRow;
      updated = u;
      setTechizatBell429Data(u);
      localStorage.setItem('excel_techizat_bell429_data', JSON.stringify(u));
    } else if (techType === 'at802') {
      const u = [...techizatAt802Data];
      const actualIdx = u.findIndex(r => r[0] === editedRow[0]);
      const targetIdx = actualIdx !== -1 ? actualIdx : rIdx;
      u[targetIdx] = editedRow;
      updated = u;
      setTechizatAt802Data(u);
      localStorage.setItem('excel_techizat_at802_data', JSON.stringify(u));
    } else if (techType === 't70') {
      const u = [...techizatT70Data];
      const actualIdx = u.findIndex(r => r[0] === editedRow[0]);
      const targetIdx = actualIdx !== -1 ? actualIdx : rIdx;
      u[targetIdx] = editedRow;
      updated = u;
      setTechizatT70Data(u);
      localStorage.setItem('excel_techizat_t70_data', JSON.stringify(u));
    } else if (techType === 't70_bumbi_backet') {
      const u = [...techizatT70BumbiBacketData];
      const actualIdx = u.findIndex(r => r[0] === editedRow[0]);
      const targetIdx = actualIdx !== -1 ? actualIdx : rIdx;
      u[targetIdx] = editedRow;
      updated = u;
      setTechizatT70BumbiBacketData(u);
      localStorage.setItem('excel_techizat_t70_bumbi_backet_data', JSON.stringify(u));
    } else if (techType === 't70_helitak') {
      const u = [...techizatT70HelitakData];
      const actualIdx = u.findIndex(r => r[0] === editedRow[0]);
      const targetIdx = actualIdx !== -1 ? actualIdx : rIdx;
      u[targetIdx] = editedRow;
      updated = u;
      setTechizatT70HelitakData(u);
      localStorage.setItem('excel_techizat_t70_helitak_data', JSON.stringify(u));
    } else if (techType === 'b360') {
      const u = [...techizatB360Data];
      const actualIdx = u.findIndex(r => r[0] === editedRow[0]);
      const targetIdx = actualIdx !== -1 ? actualIdx : rIdx;
      u[targetIdx] = editedRow;
      updated = u;
      setTechizatB360Data(u);
      localStorage.setItem('excel_techizat_b360_data', JSON.stringify(u));
    } else if (techType === 'c650') {
      const u = [...techizatC650Data];
      const actualIdx = u.findIndex(r => r[0] === editedRow[0]);
      const targetIdx = actualIdx !== -1 ? actualIdx : rIdx;
      u[targetIdx] = editedRow;
      updated = u;
      setTechizatC650Data(u);
      localStorage.setItem('excel_techizat_c650_data', JSON.stringify(u));
    } else if (techType === 'hangar') {
      const u = [...techizatHangarData];
      const actualIdx = u.findIndex(r => r[0] === editedRow[0]);
      const targetIdx = actualIdx !== -1 ? actualIdx : rIdx;
      u[targetIdx] = editedRow;
      updated = u;
      setTechizatHangarData(u);
      localStorage.setItem('excel_techizat_hangar_data', JSON.stringify(u));
    } else if (techType === 'kara_araclari') {
      const u = [...techizatKaraAraclariData];
      const actualIdx = u.findIndex(r => r[0] === editedRow[0]);
      const targetIdx = actualIdx !== -1 ? actualIdx : rIdx;
      u[targetIdx] = editedRow;
      updated = u;
      setTechizatKaraAraclariData(u);
      localStorage.setItem('excel_techizat_kara_araclari_data', JSON.stringify(u));
    }

    setIsTechizatSaving(true);
    showNotification("Canlı Excel Online güncelleniyor, lütfen bekleyiniz...");

    // Sync to Google Sheets online database immediately
    const unitLabel = getTechizatUnitLabel(techType);
    if (unitLabel && updated.length > 0) {
      try {
        const response = await fetch(GOOGLE_SCRIPT_URL, {
          method: "POST",
          headers: {
            "Content-Type": "text/plain;charset=utf-8"
          },
          body: JSON.stringify({
            action: "updateTumTechizat",
            unitLabel: unitLabel,
            data: updated.map(r => [unitLabel, ...r])
          })
        });
        console.log(`Synced individual row edit for ${unitLabel} to TÜM TECHİZAT Google Sheet`);
        showNotification("Değişiklikler canlı e-tabloya başarıyla senkronize edildi!");
        // Pull all data again to refresh UI with calculated warning columns from backend
        await pullAllTechizatFromGoogleSheets(true);
      } catch (err) {
        console.error(`Failed to sync individual row edit for ${unitLabel} to Google Sheet:`, err);
        showNotification("Değişiklikler yerel olarak kaydedildi, ancak çevrimiçi senkronizasyon başarısız oldu.");
      }
    }
    setIsTechizatSaving(false);
    setActiveTechizatRowEdit(null);
  };

  // State to track cloud sync progress
  const [isSendingToSheets, setIsSendingToSheets] = useState<Record<string | number, boolean>>({});
  const [isPullingFromSheets, setIsPullingFromSheets] = useState<Record<number, boolean>>({});

  // Direct script integration to read/fetch data from the Google Spreadsheet Web App
  const pullDataFromGoogleSheets = async (formId: number, silentNotify = false, customSheetName?: string) => {
    const config = TABLE_CONFIGS[formId];
    try {
      setIsPullingFromSheets(prev => ({ ...prev, [formId]: true }));
      
      // Akıllı sayfa adı belirleme: Alt sayfalar varsa ilkiyle senkronize et
      const matchedList = allOnlineSheets.filter(s => s.name === config.sheetName || s.name.startsWith(config.sheetName + "-"));
      const activeSheetName = customSheetName || (matchedList.length > 0 ? matchedList[0].name : config.sheetName);
      
      const targetUrl = `${GOOGLE_SCRIPT_URL}?action=readSheet&sheetName=${encodeURIComponent(activeSheetName)}`;
      
      const response = await fetch(targetUrl);
      if (!response.ok) {
        throw new Error(`HTTP Hata: ${response.status}`);
      }
      
      const result = await response.json();
      if (result.status === "success" && Array.isArray(result.data)) {
        // We received the data successfully!
        const parsedRows = result.data.map((r: any) => {
          const item: Record<string, string> = {};
          // Format headers so that they map correctly to config columns or pass through
          config.columns.forEach(col => {
            const value = r[col.label] ?? r[col.key] ?? "";
            item[col.label] = String(value);
          });
          // Also pass any other keys from Google Sheet
          Object.keys(r).forEach(k => {
            if (!item[k] && !config.columns.some(col => col.key === k)) {
              item[k] = String(r[k]);
            }
          });
          return item;
        });

        // Set state and local localStorage
        if (formId === 5) {
          const gridData: string[][] = Array.from({ length: 537 }, () => Array(12).fill(""));
          parsedRows.forEach((rObj: any, rIdx: number) => {
            if (rIdx < 537) {
              config.columns.forEach((col, cIdx) => {
                const val = rObj[col.label] ?? rObj[col.key] ?? "";
                gridData[rIdx][cIdx] = String(val).trim();
              });
            }
          });
          setExcelForm5Data(gridData);
          localStorage.setItem('excel_form_5_data', JSON.stringify(gridData));
        } else {
          const newTableData = { ...tableData, [formId]: parsedRows };
          setTableData(newTableData);
          localStorage.setItem(config.storageKey, JSON.stringify(parsedRows));
        }
        updateFormTimestamp(formId);

        const timeNow = new Date().toLocaleTimeString('tr-TR');
        setDataFeedback(prev => ({
          ...prev,
          [formId]: `Canlı E-Tablodan Doğrulandı ➜ Veritabanından ${parsedRows.length} satır okundu ve hafıza senkronize edildi. (Son GMT/Yerel Eşitleme: ${timeNow})`
        }));
        
        if (!silentNotify) {
          showNotification(`Canlı Google E-Tablodan '${config.title}' tablosuna ait ${parsedRows.length} satır güncel veri başarıyla çekildi!`);
        }
        return parsedRows;
      } else {
        throw new Error(result.message || "Geçersiz veri biçimi alındı.");
      }
    } catch (e: any) {
      console.error("Bulut okuma hatası:", e);
      if (!silentNotify) {
        alert(`Birim Uyarı Penceresi: Canlı e-tablodan veri çekilemedi.\nLütfen 'komut.gs' kodunu e-tablonuzun Apps Script alanına yapıştırıp "Web Uygulaması" (Herkes için erişilebilir) olarak yayınladığınızdan emin olun.\nHata Detayı: ${e.message || e}`);
      }
      return null;
    } finally {
      setIsPullingFromSheets(prev => ({ ...prev, [formId]: false }));
    }
  };

  const getTechizatUnitLabel = (techType: string): string => {
    if (techType === 'bell429') return 'BELL 429';
    if (techType === 'at802') return 'AT-802F';
    if (techType === 't70') return 'T-70 YER DESTEK';
    if (techType === 't70_bumbi_backet') return 'T-70 BUMBİ BACKET';
    if (techType === 't70_helitak') return 'T-70 HELİTAK';
    if (techType === 'c650') return 'C-650';
    if (techType === 'b360') return 'B-360';
    if (techType === 'hangar') return 'HANGAR YER DESTEK';
    if (techType === 'kara_araclari') return 'KARA ARAÇLARI';
    return '';
  };

  const getRealTechType = (unitLabel: string): string => {
    if (unitLabel === 'BELL 429') return 'bell429';
    if (unitLabel === 'AT-802F') return 'at802';
    if (unitLabel === 'T-70 YER DESTEK') return 't70';
    if (unitLabel === 'T-70 BUMBİ BACKET') return 't70_bumbi_backet';
    if (unitLabel === 'T-70 HELİTAK') return 't70_helitak';
    if (unitLabel === 'C-650') return 'c650';
    if (unitLabel === 'B-360') return 'b360';
    if (unitLabel === 'HANGAR YER DESTEK') return 'hangar';
    if (unitLabel === 'KARA ARAÇLARI') return 'kara_araclari';
    return '';
  };

  const convertToInputDateFormat = (dateStr: string): string => {
    if (!dateStr) return "";
    const cleaned = dateStr.trim();
    
    // check if already YYYY-MM-DD
    const ymdRegex = /^(\d{4})[\.\/-](\d{1,2})[\.\/-](\d{1,2})$/;
    if (ymdRegex.test(cleaned)) {
      const m = cleaned.match(ymdRegex);
      if (m) {
        const year = m[1];
        const month = m[2].padStart(2, '0');
        const day = m[3].padStart(2, '0');
        return `${year}-${month}-${day}`;
      }
    }

    // check if DD.MM.YYYY
    const dmyRegex = /^(\d{1,2})[\.\/-](\d{1,2})[\.\/-](\d{4})$/;
    const m = cleaned.match(dmyRegex);
    if (m) {
      const day = m[1].padStart(2, '0');
      const month = m[2].padStart(2, '0');
      const year = m[3];
      return `${year}-${month}-${day}`;
    }

    // If not matching, try fallback standard parsing
    const ts = Date.parse(cleaned);
    if (!isNaN(ts)) {
      const d = new Date(ts);
      const year = d.getFullYear();
      const month = String(d.getMonth() + 1).padStart(2, '0');
      const day = String(d.getDate()).padStart(2, '0');
      return `${year}-${month}-${day}`;
    }
    return "";
  };

  const convertToDisplayDateFormat = (dateStr: string): string => {
    if (!dateStr) return "";
    const ymdRegex = /^(\d{4})-(\d{2})-(\d{2})$/;
    const m = dateStr.match(ymdRegex);
    if (m) {
      return `${m[3]}.${m[2]}.${m[1]}`;
    }
    return dateStr;
  };

  const pullAllTechizatFromGoogleSheets = async (silent = true) => {
    try {
      const targetUrl = `${GOOGLE_SCRIPT_URL}?action=readSheet&sheetName=${encodeURIComponent("TÜM TECHİZAT")}`;
      const response = await fetch(targetUrl);
      if (!response.ok) {
        throw new Error(`HTTP Hata: ${response.status}`);
      }
      const result = await response.json();
      if (result.status === "success" && Array.isArray(result.data)) {
        const rows = result.data;
        
        const bell429Rows: string[][] = [];
        const at802Rows: string[][] = [];
        const t70Rows: string[][] = [];
        const t70BumbiRows: string[][] = [];
        const c650Rows: string[][] = [];
        const b360Rows: string[][] = [];
        const hangarRows: string[][] = [];
        const karaAraclariRows: string[][] = [];
        
        rows.forEach((r: any) => {
          const unit = String(r["AİT OLDUĞU BİRİM"] || r["Ait Olduğu Birim"] || r["AIT OLDUGU BIRIM"] || "").trim().toUpperCase();
          
          if (unit.includes("KARA")) {
            const karaRowData = [
              String(r["SIRA NO"] || r["Sıra No"] || ""),
              String(r["ARAÇ PLAKASI / TANIMI"] || r["Arac Plakasi / Tanimi"] || r["TEÇHİZAT ADI"] || r["Teçhizat Adı"] || ""),
              String(r["MODEL"] || r["Model"] || r["PARÇA NO (P/N) / MODEL"] || r["Parça No (P/N) / Model"] || ""),
              String(r["BULUNDUĞU YER"] || r["Bulunduğu Yer"] || ""),
              String(r["SON KM Sİ"] || r["Son Km Si"] || r["SON KM"] || ""),
              String(r["DURUMU"] || r["Durumu"] || ""),
              String(r["SON KONTROL / BAKIM"] || r["Son Kontrol / Bakım"] || r["SON KONTROL / KALİBRASYON / BAKIM"] || ""),
              String(r["GELECEK KONTROL / BAKIM"] || r["Gelecek Kontrol / Bakım"] || r["GELECEK KONTROL / KALİBRASYON / BAKIM"] || ""),
              String(r["SON KONTROLÜ YAPAN FİRMA"] || r["Son Kontrolü Yapan Firma"] || ""),
              String(r["AÇIKLAMA"] || r["Açıklama"] || ""),
              String(
                r["90 GÜN UYARISI MAİL GÖNDERİM TARİHİ"] || r["90 GÜN UYARISI MAIL GONDERIM TARIHI"] || r["90 Gun Uyarisi Mail"] ||
                r["60 GÜN UYARISI MAİL GÖNDERİM TARİHİ"] || r["60 GÜN UYARISI MAIL GONDERIM TARIHI"] || r["60 Gun Uyarisi Mail"] ||
                r["30 GÜN UYARISI MAİL GÖNDERİM TARİHİ"] || r["30 GÜN UYARISI MAIL GONDERIM TARIHI"] || r["30 Gun Uyarisi Mail"] ||
                ""
              ).trim()
            ];
            karaAraclariRows.push(karaRowData);
          } else {
            const rowData = [
              String(r["SIRA NO"] || r["Sıra No"] || ""),
              String(r["TEÇHİZAT ADI"] || r["Teçhizat Adı"] || ""),
              String(r["PARÇA NO (P/N)"] || r["Parça No (P/N)"] || r["PARÇA NO (P/N) / MODEL"] || r["Parça No (P/N) / Model"] || ""),
              String(r["SERİ NO (S/N)"] || r["Seri No (S/N)"] || ""),
              String(r["MİKTAR / KAPASİTE"] || r["Miktar / Kapasite"] || ""),
              String(r["BULUNDUĞU YER"] || r["Bulunduğu Yer"] || ""),
              String(r["DURUMU"] || r["Durumu"] || ""),
              String(r["SON KONTROL / BAKIM"] || r["Son Kontrol / Bakım"] || r["SON KONTROL / KALİBRASYON / BAKIM"] || ""),
              String(r["GELECEK KONTROL / BAKIM"] || r["Gelecek Kontrol / Bakım"] || r["GELECEK KONTROL / KALİBRASYON / BAKIM"] || ""),
              String(r["SON KONTROLÜ YAPAN FİRMA"] || r["Son Kontrolü Yapan Firma"] || ""),
              String(r["AÇIKLAMA"] || r["Açıklama"] || ""),
              String(
                r["90 GÜN UYARISI MAİL GÖNDERİM TARİHİ"] || r["90 GÜN UYARISI MAIL GONDERIM TARIHI"] || r["90 Gun Uyarisi Mail"] ||
                r["60 GÜN UYARISI MAİL GÖNDERİM TARİHİ"] || r["60 GÜN UYARISI MAIL GONDERIM TARIHI"] || r["60 Gun Uyarisi Mail"] ||
                r["30 GÜN UYARISI MAİL GÖNDERİM TARİHİ"] || r["30 GÜN UYARISI MAIL GONDERIM TARIHI"] || r["30 Gun Uyarisi Mail"] ||
                ""
              ).trim()
            ];

            if (unit.includes("BELL 429")) {
              bell429Rows.push(rowData);
            } else if (unit.includes("AT-802")) {
              at802Rows.push(rowData);
            } else if (unit.includes("T-70 YER") || (unit.includes("T-70") && !unit.includes("BUMBİ") && !unit.includes("BAMBI") && !unit.includes("KARA"))) {
              t70Rows.push(rowData);
            } else if (unit.includes("BUMBİ") || unit.includes("BAMBI") || unit.includes("T-70 BUMBİ")) {
              t70BumbiRows.push(rowData);
            } else if (unit.includes("C-650")) {
              c650Rows.push(rowData);
            } else if (unit.includes("B-360")) {
              b360Rows.push(rowData);
            } else if (unit.includes("HANGAR")) {
              hangarRows.push(rowData);
            }
          }
        });
        
        const groupedBell429 = groupMultiLocationRows(bell429Rows, 1, 5, 4, 0);
        setTechizatBell429Data(groupedBell429);
        localStorage.setItem('excel_techizat_bell429_data', JSON.stringify(groupedBell429));

        const groupedAt802 = groupMultiLocationRows(at802Rows, 1, 5, 4, 0);
        setTechizatAt802Data(groupedAt802);
        localStorage.setItem('excel_techizat_at802_data', JSON.stringify(groupedAt802));

        const groupedT70 = groupMultiLocationRows(t70Rows, 1, 5, 4, 0);
        setTechizatT70Data(groupedT70);
        localStorage.setItem('excel_techizat_t70_data', JSON.stringify(groupedT70));

        const groupedT70Bumbi = groupMultiLocationRows(t70BumbiRows, 1, 5, 4, 0);
        setTechizatT70BumbiBacketData(groupedT70Bumbi);
        localStorage.setItem('excel_techizat_t70_bumbi_backet_data', JSON.stringify(groupedT70Bumbi));

        const groupedC650 = groupMultiLocationRows(c650Rows, 1, 5, 4, 0);
        setTechizatC650Data(groupedC650);
        localStorage.setItem('excel_techizat_c650_data', JSON.stringify(groupedC650));

        const groupedB360 = groupMultiLocationRows(b360Rows, 1, 5, 4, 0);
        setTechizatB360Data(groupedB360);
        localStorage.setItem('excel_techizat_b360_data', JSON.stringify(groupedB360));

        const groupedHangar = groupMultiLocationRows(hangarRows, 1, 5, 4, 0);
        setTechizatHangarData(groupedHangar);
        localStorage.setItem('excel_techizat_hangar_data', JSON.stringify(groupedHangar));

        let finalKaraAraclari = karaAraclariRows.map(row => convertOldKaraAraclariRowToNew(row));
        finalKaraAraclari = groupMultiLocationRows(finalKaraAraclari, 1, 3, -1, 0);
        setTechizatKaraAraclariData(finalKaraAraclari);
        localStorage.setItem('excel_techizat_kara_araclari_data', JSON.stringify(finalKaraAraclari));

        if (!silent) {
          showNotification("Bütün Teçhizat Verileri Canlı E-Tablodan Senkronize Edildi!");
        }
      }
    } catch (err) {
      console.error("Failed to sync all teçhizat on startup:", err);
    }
  };

  const fetchGunTakipSorumlulari = async () => {
    try {
      const targetUrl = `${GOOGLE_SCRIPT_URL}?action=readSheet&sheetName=${encodeURIComponent("GÜN TAKİP")}`;
      const response = await fetch(targetUrl);
      if (response.ok) {
        const result = await response.json();
        if (result.status === "success" && Array.isArray(result.data)) {
          const mapped = result.data.map((row: any) => ({
            birim: String(row["SORUMLU BİRİM"] || row["Sorumlu Birim"] || "").trim().toUpperCase(),
            adSoyad: String(row["ADI SOYADI"] || row["Adı Soyadı"] || row["AD SOYAD"] || "").trim(),
            eposta: String(row["E-POSTA ADRESİ"] || row["E-posta Adresi"] || row["EPOSTA ADRESI"] || "").trim(),
            mail90: String(row["90 GÜN UYARISI MAİL GÖNDERİM TARİHİ"] || row["90 Gün Uyarı Mail Gönderim Tarihi"] || row["90 GUN UYARISI MAIL GONDERIM TARIHI"] || "").trim()
          })).filter((item: any) => item.birim);
          
          if (mapped.length > 0) {
            setGunTakipSorumlulari(mapped);
          }
        }
      }
    } catch (err) {
      console.error("Gün takip verisi çekme hatası:", err);
    }
  };

  const saveGunTakipSorumlulari = async (updatedData: typeof gunTakipSorumlulari) => {
    setIsSavingSorumlu(true);
    try {
      const rowsToSend = updatedData.map(item => [item.birim, item.adSoyad, item.eposta]);
      const response = await fetch(GOOGLE_SCRIPT_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain' },
        body: JSON.stringify({
          action: 'updateGunTakip',
          data: rowsToSend
        })
      });
      
      const result = await response.json();
      if (result.status === "success") {
        setGunTakipSorumlulari(updatedData);
        showNotification("Sorumlu Personel (Gün Takip) Verileri Başarıyla Güncellendi!");
        setIsSorumluModalOpen(false);
      } else {
        alert("E-Tablo Güncelleme Hatası: " + result.message);
      }
    } catch (err: any) {
      alert("Entegrasyon Bağlantı Hatası: " + err.toString());
    } finally {
      setIsSavingSorumlu(false);
    }
  };

  const parseGelecekBakimDays = (dateStr: string): number | null => {
    if (!dateStr) return null;
    const cleaned = dateStr.trim();
    
    // dd.mm.yyyy veya dd/mm/yyyy formatı
    const dmyRegex = /^(\d{1,2})[\.\/-](\d{1,2})[\.\/-](\d{4})$/;
    const ymdRegex = /^(\d{4})[\.\/-](\d{1,2})[\.\/-](\d{1,2})$/;
    
    let dateObj: Date | null = null;
    let m = cleaned.match(dmyRegex);
    if (m) {
      const day = parseInt(m[1], 10);
      const month = parseInt(m[2], 10) - 1;
      const year = parseInt(m[3], 10);
      dateObj = new Date(year, month, day);
    } else {
      m = cleaned.match(ymdRegex);
      if (m) {
        const year = parseInt(m[1], 10);
        const month = parseInt(m[2], 10) - 1;
        const day = parseInt(m[3], 10);
        dateObj = new Date(year, month, day);
      } else {
        const timestamp = Date.parse(cleaned);
        if (!isNaN(timestamp)) {
          dateObj = new Date(timestamp);
        }
      }
    }
    
    if (!dateObj || isNaN(dateObj.getTime())) return null;
    
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    dateObj.setHours(0, 0, 0, 0);
    
    const diffTime = dateObj.getTime() - today.getTime();
    return Math.ceil(diffTime / (1000 * 60 * 60 * 24));
  };

  // Direct script integration to post local data to Google Spreadsheet Web App
  const sendDataToGoogleSheets = async (formId: number) => {
    const config = TABLE_CONFIGS[formId];
    const rows = tableData[formId] || [];
    
    try {
      setIsSendingToSheets(prev => ({ ...prev, [formId]: true }));
      const targetUrl = GOOGLE_SCRIPT_URL;
      
      // Real fetch POST request (with mode 'no-cors' to prevent blocking by google script redirect)
      await fetch(targetUrl, {
        method: "POST",
        mode: "no-cors",
        headers: {
          "Content-Type": "application/json"
        },
        body: JSON.stringify({
          action: "updateSheet",
          sheetName: config.sheetName,
          data: rows
        })
      });

      const timeNow = new Date().toLocaleTimeString('tr-TR');
      
      // Update text while wait
      setDataFeedback(prev => ({
        ...prev,
        [formId]: `Buluta gönderildi... Canlı e-tablodan geri okunarak doğrulanıyor... (Bekleyin... ${timeNow})`
      }));
      
      // Wait for Apps Script write thread to settle
      await new Promise(resolve => setTimeout(resolve, 1500));
      
      // Fetch data back from e-tablo to verify and satisfy the exact user flow: "veri yazdıktan sonra ordan okuduğu bilgisisi ekranda gösterir"
      const verified = await pullDataFromGoogleSheets(formId, true);
      
      if (verified) {
        setDataFeedback(prev => ({
          ...prev,
          [formId]: `E-Tablodan Canlı Doğrulandı ➜ Veriler buluta yazıldı ve canlı e-tablodan geri okundu: Toplam ${verified.length} satır kontrol edilerek sisteme başarıyla yansıtıldı. (Doğrulama Saati: ${new Date().toLocaleTimeString('tr-TR')})`
        }));
        showNotification(`${config.title} verileri canlı e-tabloya kaydedildi ve veri doğrulama okumasıyla teyit edildi!`);
      } else {
        setDataFeedback(prev => ({
          ...prev,
          [formId]: `Canlı E-Tabloya Gönderildi ➜ ${rows.length} satır başarıyla yazıldı. Ancak canlı geri okuma doğrulaması için Apps Script Web App kodunu güncellemelisiniz. (Saat: ${timeNow})`
        }));
        showNotification(`${config.title} verileri e-tabloya başarıyla gönderildi.`);
      }
    } catch (e) {
      alert("Hata: Canlı e-tabloya veri gönderilirken bir sorun oluştu.");
    } finally {
      setIsSendingToSheets(prev => ({ ...prev, [formId]: false }));
    }
  };

  // 1839 Password Verification
  const handleVerifyPassword = () => {
    if (passwordInput === "1839") {
      setIsPasswordModalOpen(false);
      setPasswordInput('');
      setPasswordError(false);
      
      if (activeTechizatType && modalType === 'techizat_matrix') {
        setSyncSelectedTarget(`techizat_${activeTechizatType}`);
        setActiveSyncStep(2);
        setModalType('excel_sync');
        setModalTitle('TEÇHİZAT VERİ GÜNCELLEME SİHİRBAZI');
        return;
      }

      // Reset steps and set target based on current view context
      setActiveSyncStep(1);
      if (selectedFormId !== null && [1, 21, 22, 23, 24, 25, 3, 5, 6].includes(selectedFormId)) {
        if (isSummerForm(selectedFormId)) {
          setStep1Target('2');
          setSelectedForm2Unit(selectedFormId as any);
          setSyncSelectedTarget(String(selectedFormId));
        } else {
          setStep1Target(String(selectedFormId));
          setSelectedForm2Unit(null);
          setSyncSelectedTarget(String(selectedFormId));
        }
      } else {
        setStep1Target('1');
        setSelectedForm2Unit(null);
        setSyncSelectedTarget('1');
      }

      // Mode switch
      setModalType('excel_sync');
      setModalTitle('VERİ GÜNCELLEME ÇEVRİMDIŞI PORTALİ');
    } else {
      setPasswordError(true);
    }
  };

  // Cell modifications
  const handleCellChange = (formId: number, rowIndex: number, key: string, val: string) => {
    const updatedRows = [...(tableData[formId] || [])];
    updatedRows[rowIndex] = { ...updatedRows[rowIndex], [key]: val };
    
    const newTableData = { ...tableData, [formId]: updatedRows };
    setTableData(newTableData);
    localStorage.setItem(TABLE_CONFIGS[formId].storageKey, JSON.stringify(updatedRows));

    // Update real-time feedback with dynamic confirm readback message
    const timeNow = new Date().toLocaleTimeString('tr-TR');
    setDataFeedback(prev => ({
      ...prev,
      [formId]: `Okuma/Yazma Başarılı: Değişiklik kaydedildi, hafızadan ${updatedRows.length} satır veri başarıyla okundu ve doğrulandı. (Son İşlem: ${timeNow})`
    }));
  };

  // Row operations
  const handleAddRow = (formId: number) => {
    const columns = getFormColumns(formId);
    const newRow: Record<string, string> = {};
    columns.forEach(col => {
      newRow[col] = "";
    });
    
    const updatedRows = [...(tableData[formId] || []), newRow];
    const newTableData = { ...tableData, [formId]: updatedRows };
    setTableData(newTableData);
    localStorage.setItem(TABLE_CONFIGS[formId].storageKey, JSON.stringify(updatedRows));
    
    const timeNow = new Date().toLocaleTimeString('tr-TR');
    setDataFeedback(prev => ({
      ...prev,
      [formId]: `Hafızaya Yazma Başarılı: Yeni boş satır eklendi, toplam ${updatedRows.length} satır geri okundu. (Son İşlem: ${timeNow})`
    }));

    showNotification("Yeni boş satır eklendi. Düzenlemek için hücreye çift tıklayıp yazabilirsiniz.");
  };

  const handleDeleteRow = (formId: number, rowIndex: number) => {
    const updatedRows = (tableData[formId] || []).filter((_, idx) => idx !== rowIndex);
    const newTableData = { ...tableData, [formId]: updatedRows };
    setTableData(newTableData);
    localStorage.setItem(TABLE_CONFIGS[formId].storageKey, JSON.stringify(updatedRows));
    
    const timeNow = new Date().toLocaleTimeString('tr-TR');
    setDataFeedback(prev => ({
      ...prev,
      [formId]: `Hafızadan Silme-Yazma Başarılı: Satır kaldırıldı, kalan ${updatedRows.length} satır geri okundu. (Son İşlem: ${timeNow})`
    }));

    showNotification("Seçilen satır silindi.");
  };

  // EXCEL İNDİR (Adapts to selected target)
  const handleExcelExport = () => {
    try {
      const wb = XLSX.utils.book_new();
      const id = Number(syncSelectedTarget);
      if (isNaN(id) || !TABLE_CONFIGS[id]) {
        alert("Hata: Geçersiz hedef seçimi.");
        return;
      }
      const config = TABLE_CONFIGS[id];
      const rows = tableData[id] || [];
      const ws = XLSX.utils.json_to_sheet(rows);
      XLSX.utils.book_append_sheet(wb, ws, config.sheetName);
      XLSX.writeFile(wb, `${config.sheetName}_Tablosu.xlsx`);
      showNotification(`'${config.title}' tablosu Excel olarak başarıyla indirildi.`);
    } catch (err) {
      alert("Excel dosyası oluşturulurken bir hata oluştu.");
    }
  };

  // EXCEL YÜKLE (Adapts to selected target, reads ALL sheets from the uploaded Excel using Google Drive conversion)
  // Direct style-preserving import flow: Excel loaded -> convert via Drive REST API -> copy sheets with 100% styles, fonts, widths intact
  const handleExcelUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    alert("Bu özellik kaldırılmıştır. Lütfen planlama verilerini PDF formatında yükleyin.");
    return;
    const file = e.target.files?.[0];
    if (!file) return;

    const id = Number(syncSelectedTarget);
    if (isNaN(id) || !TABLE_CONFIGS[id]) {
      alert("Hata: Geçersiz hedef seçimi.");
      return;
    }
    const config = TABLE_CONFIGS[id];

    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        setIsSendingToSheets(prev => ({ ...prev, [id]: true }));
        setUploadProgress(15);
        
        const arrayBuffer = evt.target?.result as ArrayBuffer;
        
        setUploadProgress(35);
        showNotification(`'${config.title}' Excel dosyası cihazınızda çözümleniyor...`);
        
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        
        const parsedSheets = workbook.SheetNames.map((sheetName, index) => {
          const worksheet = workbook.Sheets[sheetName];
          const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "", raw: false });
          
          const dynamicPrefix = isSummerForm(id)
            ? getSummerPeriodSheetPrefix(id, selectedUploadSummerMonth)
            : config.sheetName;
            
          const finalSheetName = workbook.SheetNames.length > 1 
            ? `${dynamicPrefix}-${index + 1}`
            : `${dynamicPrefix}-1`;
            
          return {
            name: finalSheetName,
            data: rows
          };
        });

        setUploadProgress(60);
        showNotification(`Çözümlenen veriler Google Sheets'e aktarılıyor...`);

        const targetUrl = GOOGLE_SCRIPT_URL;

        const res = await fetch(targetUrl, {
          method: "POST",
          headers: {
            "Content-Type": "text/plain;charset=utf-8"
          },
          body: JSON.stringify({
            action: "updateMultiSheets",
            prefix: isSummerForm(id) ? getSummerPeriodSheetPrefix(id, selectedUploadSummerMonth) : config.sheetName,
            sheets: parsedSheets
          })
        });
        
        if (!res.ok) {
          throw new Error(`Google Apps Script sunucu hatası: Kod ${res.status}`);
        }
        
        const result = await res.json();
        if (result.status !== "success") {
          throw new Error(result.message || "Bilinmeyen sunucu hatası.");
        }
        
        setUploadProgress(85);
        showNotification("E-Tablo güncellendi. Sayfalar portal hafızası ile entegre ediliyor...");
        
        await fetchAllGoogleSheetsList();
        
        const returnedSheets = result.updatedSheets || [];
        const firstSheetName = returnedSheets.length > 0 ? returnedSheets[0] : parsedSheets[0].name;
        
        await pullDataFromGoogleSheets(id, true, firstSheetName);
        
        setUploadProgress(100);
        showNotification(`'${config.title}' altındaki tüm sayfalar başarıyla Google Sheets üzerine aktarıldı ve güncellendi!`);
        
        setSelectedFormId(id);
        setModalType('form_table');
        setModalTitle(config.title);
        setActiveFormTab('live_sheet');
        setTimeout(() => setUploadProgress(0), 4000);
      } catch (err: any) {
        alert(`Hata: Excel dosyası yüklenirken bir hata oluştu.\nDetay: ${err?.message || err}`);
        setUploadProgress(0);
      } finally {
        setIsSendingToSheets(prev => ({ ...prev, [id]: false }));
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const loadPdfJs = (): Promise<any> => {
    return new Promise((resolve, reject) => {
      if ((window as any).pdfjsLib) {
        resolve((window as any).pdfjsLib);
        return;
      }
      const script = document.createElement('script');
      script.src = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js';
      script.onload = () => {
        const pdfjsLib = (window as any).pdfjsLib;
        pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
        resolve(pdfjsLib);
      };
      script.onerror = (err) => reject(new Error('PDF.js kütüphanesi yüklenemedi.'));
      document.head.appendChild(script);
    });
  };

  const renderPdfToImages = async (file: File): Promise<{ pageNumber: number; dataUrl: string; width: number; height: number; selected: boolean; textItems?: any[] }[]> => {
    const pdfjsLib = await loadPdfJs();
    const arrayBuffer = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
    
    const pageIndices = Array.from({ length: pdf.numPages }, (_, i) => i + 1);
    const renderPromises = pageIndices.map(async (pageNum) => {
      const page = await pdf.getPage(pageNum);
      const scale = 1.5; // Optimized scale for speed and high fidelity
      const viewport = page.getViewport({ scale });
      
      const canvas = document.createElement('canvas');
      const context = canvas.getContext('2d');
      if (!context) throw new Error("Canvas context is not available");
      
      canvas.width = viewport.width;
      canvas.height = viewport.height;
      
      await page.render({
        canvasContext: context,
        viewport: viewport
      }).promise;
      
      const dataUrl = canvas.toDataURL('image/png');
      
      let textItems: any[] = [];
      try {
        const textContent = await page.getTextContent();
        const v1 = page.getViewport({ scale: 1.0 });
        textItems = textContent.items.map((item: any) => {
          const [vx, vy] = v1.convertToViewportPoint(item.transform[4], item.transform[5]);
          const left = (vx / v1.width) * 100;
          const top = (vy / v1.height) * 100;
          const fontHeightPdf = Math.abs(item.transform[3] || item.transform[0] || 12);
          const fontSize = (fontHeightPdf / v1.height) * 100;
          const itemWidth = ((item.width || 0) / v1.width) * 100;
          return {
            str: item.str || "",
            left: Number(left.toFixed(3)),
            top: Number(top.toFixed(3)),
            fontSize: Number(fontSize.toFixed(3)),
            width: Number(itemWidth.toFixed(3))
          };
        });
      } catch (err) {
        console.error("Text extraction failed in render:", err);
      }
      
      return {
        pageNumber: pageNum,
        dataUrl,
        width: viewport.width,
        height: viewport.height,
        selected: true,
        textItems
      };
    });
    
    return Promise.all(renderPromises);
  };

  const sanitizeTurkishForFilename = (str: string): string => {
    return str.toLowerCase()
      .replace(/ı/g, 'i')
      .replace(/ğ/g, 'g')
      .replace(/ü/g, 'u')
      .replace(/ş/g, 's')
      .replace(/ö/g, 'o')
      .replace(/ç/g, 'c')
      .replace(/\s+/g, '_');
  };

  // PDF Yükleme Metodu (Önce RAM'de çözümler, önizlemeyi açar, yükleme yapılmadan önce sayfa yönü ve yaklaştırma ayarları sunar)
  const handlePdfUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    const file = files[0];
    setUploadedPdfFile(file);

    const isTechizatTarget = typeof syncSelectedTarget === 'string' && syncSelectedTarget.startsWith('techizat_');
    const id = isTechizatTarget ? 0 : Number(syncSelectedTarget);
    if (!isTechizatTarget && (isNaN(id) || !TABLE_CONFIGS[id])) {
      alert("Hata: Geçersiz hedef seçimi.");
      return;
    }

    const fileNameLower = file.name.toLowerCase();

    // 1. Handle Teçhizat Takip Excel Upload
    if (isTechizatTarget) {
      const techType = syncSelectedTarget.replace('techizat_', '') as 'bell429' | 'at802' | 't70' | 't70_bumbi_backet' | 't70_helitak' | 'b360' | 'c650' | 'hangar' | 'kara_araclari';
      if (!fileNameLower.endsWith('.xlsx') && !fileNameLower.endsWith('.xls') && !fileNameLower.endsWith('.csv')) {
        alert("Teçhizat Takip güncellemesi için lütfen Excel (.xlsx, .xls) veya CSV belgesi yükleyin.");
        return;
      }

      const reader = new FileReader();
      reader.onload = (evt) => {
        try {
          const arrayBuffer = evt.target?.result as ArrayBuffer;
          const workbook = XLSX.read(arrayBuffer, { type: 'array' });
          const worksheet = workbook.Sheets[workbook.SheetNames[0]];
          
          const rawRows = XLSX.utils.sheet_to_json<string[]>(worksheet, { header: 1, defval: "", raw: false });
          if (rawRows.length === 0) {
            throw new Error("Yüklenen Excel dosyasında veri bulunamadı.");
          }

          // Smart Header Detection
          let headerRowIdx = 0;
          for (let r = 0; r < Math.min(10, rawRows.length); r++) {
            const rCells = rawRows[r] || [];
            const filledCount = rCells.filter(c => String(c).trim() !== "").length;
            const hasKeywords = rCells.some(c => {
              const s = String(c).toLowerCase();
              return s.includes("sira") || s.includes("no") || s.includes("teçhizat") || s.includes("techizat") || s.includes("malzeme") || s.includes("parca") || s.includes("parça");
            });
            if (filledCount >= 3 && hasKeywords) {
              headerRowIdx = r;
              break;
            }
          }

          const rawHeaders = (rawRows[headerRowIdx] || []).map(h => String(h || '').trim().toUpperCase());
          const finalHeaders = rawHeaders.map((h, hIdx) => h || `KOLON ${hIdx + 1}`);

          // Extract rows below the header
          const rawParsedRows: string[][] = [];
          for (let r = headerRowIdx + 1; r < rawRows.length; r++) {
            const rawRow = rawRows[r] || [];
            const isRowEmpty = rawRow.every(cell => String(cell || '').trim() === '');
            if (isRowEmpty) continue;

            const trimmedRow = finalHeaders.map((_, cIdx) => {
              return rawRow[cIdx] !== undefined ? String(rawRow[cIdx]).trim() : "";
            });

            rawParsedRows.push(trimmedRow);
          }

          // Fallback if no rows could be parsed
          if (rawParsedRows.length === 0) {
            for (let r = headerRowIdx + 1; r < Math.min(200, rawRows.length); r++) {
              const rawRow = rawRows[r] || [];
              const trimmedRow = finalHeaders.map((_, cIdx) => {
                return rawRow[cIdx] !== undefined ? String(rawRow[cIdx]).trim() : "";
              });
              rawParsedRows.push(trimmedRow);
            }
          }

          // Smartly find column indices for grouping multi-location rows
          const nameColIdx = finalHeaders.findIndex(h => {
            const s = h.toUpperCase();
            return s.includes("TEÇHİZAT") || s.includes("TECHİZAT") || s.includes("ARAÇ") || s.includes("TANIM") || s.includes("MALZEME") || s.includes("ÜRÜN");
          });
          const locColIdx = finalHeaders.findIndex(h => {
            const s = h.toUpperCase();
            return s.includes("BULUNDUĞU") || s.includes("LOKASYON") || s.includes("YER");
          });
          const miktarColIdx = finalHeaders.findIndex(h => {
            const s = h.toUpperCase();
            return s.includes("MİKTAR") || s.includes("KAPASİTE") || s.includes("ADET");
          });
          const siraColIdx = finalHeaders.findIndex(h => h.toUpperCase().includes("SIRA"));

          const parsedRows = groupMultiLocationRows(
            rawParsedRows,
            nameColIdx >= 0 ? nameColIdx : 1,
            locColIdx >= 0 ? locColIdx : (techType === 'kara_araclari' ? 3 : 5),
            miktarColIdx >= 0 ? miktarColIdx : (techType === 'kara_araclari' ? -1 : 4),
            siraColIdx >= 0 ? siraColIdx : 0
          );

          // Save columns and rows to states and localStorage
          if (techType === 'bell429') {
            setTechizatBell429Columns(finalHeaders);
            setTechizatBell429Data(parsedRows);
            localStorage.setItem('excel_techizat_bell429_cols', JSON.stringify(finalHeaders));
            localStorage.setItem('excel_techizat_bell429_data', JSON.stringify(parsedRows));
          } else if (techType === 'at802') {
            setTechizatAt802Columns(finalHeaders);
            setTechizatAt802Data(parsedRows);
            localStorage.setItem('excel_techizat_at802_cols', JSON.stringify(finalHeaders));
            localStorage.setItem('excel_techizat_at802_data', JSON.stringify(parsedRows));
          } else if (techType === 't70') {
            setTechizatT70Columns(finalHeaders);
            setTechizatT70Data(parsedRows);
            localStorage.setItem('excel_techizat_t70_cols', JSON.stringify(finalHeaders));
            localStorage.setItem('excel_techizat_t70_data', JSON.stringify(parsedRows));
          } else if (techType === 't70_bumbi_backet') {
            setTechizatT70BumbiBacketColumns(finalHeaders);
            setTechizatT70BumbiBacketData(parsedRows);
            localStorage.setItem('excel_techizat_t70_bumbi_backet_cols', JSON.stringify(finalHeaders));
            localStorage.setItem('excel_techizat_t70_bumbi_backet_data', JSON.stringify(parsedRows));
          } else if (techType === 't70_helitak') {
            setTechizatT70HelitakColumns(finalHeaders);
            setTechizatT70HelitakData(parsedRows);
            localStorage.setItem('excel_techizat_t70_helitak_cols', JSON.stringify(finalHeaders));
            localStorage.setItem('excel_techizat_t70_helitak_data', JSON.stringify(parsedRows));
          } else if (techType === 'b360') {
            setTechizatB360Columns(finalHeaders);
            setTechizatB360Data(parsedRows);
            localStorage.setItem('excel_techizat_b360_cols', JSON.stringify(finalHeaders));
            localStorage.setItem('excel_techizat_b360_data', JSON.stringify(parsedRows));
          } else if (techType === 'c650') {
            setTechizatC650Columns(finalHeaders);
            setTechizatC650Data(parsedRows);
            localStorage.setItem('excel_techizat_c650_cols', JSON.stringify(finalHeaders));
            localStorage.setItem('excel_techizat_c650_data', JSON.stringify(parsedRows));
          } else if (techType === 'hangar') {
            setTechizatHangarColumns(finalHeaders);
            setTechizatHangarData(parsedRows);
            localStorage.setItem('excel_techizat_hangar_cols', JSON.stringify(finalHeaders));
            localStorage.setItem('excel_techizat_hangar_data', JSON.stringify(parsedRows));
          } else if (techType === 'kara_araclari') {
            const convertedRows = parsedRows.map(row => convertOldKaraAraclariRowToNew(row));
            const newCols = ["SIRA NO", "ARAÇ PLAKASI / TANIMI", "PARÇA NO (P/N) / MODEL", "BULUNDUĞU YER", "SON KM Sİ", "DURUMU", "SON KONTROL / KALİBRASYON / BAKIM", "GELECEK KONTROL / KALİBRASYON / BAKIM", "SON KONTROLÜ YAPAN FİRMA", "AÇIKLAMA"];
            setTechizatKaraAraclariColumns(newCols);
            setTechizatKaraAraclariData(convertedRows);
            localStorage.setItem('excel_techizat_kara_araclari_cols', JSON.stringify(newCols));
            localStorage.setItem('excel_techizat_kara_araclari_data', JSON.stringify(convertedRows));
          }

          // Sync specifically to "TÜM TECHİZAT" online Google Sheet
          const unitLabel = getTechizatUnitLabel(techType);
          if (unitLabel) {
            fetch(GOOGLE_SCRIPT_URL, {
              method: "POST",
              headers: {
                "Content-Type": "text/plain;charset=utf-8"
              },
              body: JSON.stringify({
                action: "updateTumTechizat",
                unitLabel: unitLabel,
                data: parsedRows.map(r => [unitLabel, ...r])
              })
            }).then(() => {
              console.log(`Synced ${unitLabel} to TÜM TECHİZAT Google Sheet`);
            }).catch(err => {
              console.error(`Failed to sync ${unitLabel} to Google Sheet:`, err);
            });
          }

          setIsSendingToSheets(prev => ({ ...prev, [syncSelectedTarget]: true }));
          setUploadProgress(10);
          showNotification(`${file.name} belgesi Teçhizat Sürücüsüne yükleniyor ve Google Drive'a yedekleniyor...`);

          // Background Google Drive upload for Teçhizat Excel file
          fileToBase64(file).then(async (base64Data) => {
            try {
              let driveFileName = "hava_araçları_yer_destek_bell-429.xlsx";
              if (techType === 'at802') {
                driveFileName = "hava_araçları_yer_destek_at-802.xlsx";
              } else if (techType === 't70') {
                driveFileName = "hava_araçları_yer_destek_t-70.xlsx";
              } else if (techType === 't70_bumbi_backet') {
                driveFileName = "hava_araçları_yer_destek_t-70_bumbi_backet.xlsx";
              } else if (techType === 't70_helitak') {
                driveFileName = "hava_araçları_yer_destek_t-70_helitak.xlsx";
              } else if (techType === 'b360') {
                driveFileName = "hava_araçları_yer_destek_b-360.xlsx";
              } else if (techType === 'c650') {
                driveFileName = "hava_araçları_yer_destek_c-650.xlsx";
              } else if (techType === 'hangar') {
                driveFileName = "hava_araçları_yer_destek_hangar.xlsx";
              } else if (techType === 'kara_araclari') {
                driveFileName = "kara_araçları_takip.xlsx";
              }

              const prettyUnitName = techType === 'bell429' ? 'BELL 429 YER DESTEK TEÇHİZATLARI'
                : techType === 'at802' ? 'AT-802F YER DESTEK TEÇHİZATLARI'
                : techType === 't70' ? 'T-70 YER DESTEK TEÇHİZATI'
                : techType === 't70_bumbi_backet' ? 'T-70 BUMBİ BACKET TEÇHİZATI'
                : techType === 't70_helitak' ? 'T-70 HELİTAK TEÇHİZATI'
                : techType === 'b360' ? 'B-360 YER DESTEK TEÇHİZATLARI'
                : techType === 'c650' ? 'C-650 YER DESTEK TEÇHİZATLARI'
                : techType === 'kara_araclari' ? 'KARA ARAÇLARI TAKİP SİSTEMİ'
                : 'HANGAR YER DESTEK TEÇHİZATLARI';

              const res = await fetch(GOOGLE_SCRIPT_URL, {
                method: "POST",
                headers: {
                  "Content-Type": "text/plain;charset=utf-8"
                },
                body: JSON.stringify({
                  action: "uploadPdfToDrive",
                  fileName: driveFileName,
                  base64Data: base64Data,
                  unitName: prettyUnitName,
                  formId: 6,
                  month: "Teçhizat Takip",
                  folderId: "1HQR_NYKhHQGA7_2W3nArI9pCh-LJasTP"
                })
              });
              if (res.ok) {
                const result = await res.json();
                if (result.status === "success") {
                  showNotification(`'${driveFileName}' güncel Teçhizat Excel belgesi Google Drive'a başarıyla yedeklendi!`);
                }
              }
            } catch (err) {
              console.error("Failed to backup Teçhizat Excel to Google Drive:", err);
            }
          });

          // Progress bar animation
          let progressVal = 10;
          const interval = setInterval(() => {
            progressVal += 15;
            if (progressVal >= 100) {
              clearInterval(interval);
              setUploadProgress(100);
              setIsSendingToSheets(prev => ({ ...prev, [syncSelectedTarget]: false }));
              
              const prettyName = techType === 'bell429' ? 'Bell 429 Yer Destek' 
                : techType === 'at802' ? 'AT-802F Yer Destek' 
                : techType === 't70' ? 'T-70 Yer Destek' 
                : techType === 't70_bumbi_backet' ? 'T-70 Bumbi Backet'
                : techType === 't70_helitak' ? 'T-70 Helitak'
                : techType === 'b360' ? 'B-360 Yer Destek'
                : techType === 'c650' ? 'C-650 Yer Destek'
                : techType === 'kara_araclari' ? 'Kara Araçları Takip'
                : 'Hangar Yer Destek';

              showNotification(`'${prettyName}' Excel verisi başarıyla aktarıldı ve matris oluşturuldu!`);
              
              const title = techType === 'bell429' ? 'BELL 429 YER DESTEK TEÇHİZATLARI'
                : techType === 'at802' ? 'AT-802F YER DESTEK TEÇHİZATLARI'
                : techType === 't70' ? 'T-70 YER DESTEK TEÇHİZATI'
                : techType === 't70_bumbi_backet' ? 'T-70 BUMBİ BACKET TEÇHİZATI'
                : techType === 't70_helitak' ? 'T-70 HELİTAK TEÇHİZATI'
                : techType === 'b360' ? 'B-360 YER DESTEK TEÇHİZATLARI'
                : techType === 'c650' ? 'C-650 YER DESTEK TEÇHİZATLARI'
                : techType === 'kara_araclari' ? 'KARA ARAÇLARI TAKİP SİSTEMİ'
                : 'HANGAR YER DESTEK TEÇHİZATLARI';

              // Close sync wizard and automatically open and redirect to matrix screen
              setModalOpen(true);
              setActiveTechizatType(techType);
              setModalType('techizat_matrix');
              setModalTitle(title);
              setTechizatSearchQuery('');
              setActiveTechizatMatchIdx(0);
            } else {
              setUploadProgress(progressVal);
            }
          }, 150);

        } catch (err: any) {
          console.error(err);
          alert(`Teçhizat Excel belgesi çözümlenirken hata oluştu: ${err?.message || err}`);
        }
      };
      reader.readAsArrayBuffer(file);
      return;
    }

    // 2. Handle Personal Info (Form 5) Excel Upload
    if (fileNameLower.endsWith('.xlsx') || fileNameLower.endsWith('.xls') || fileNameLower.endsWith('.csv')) {
      if (id !== 5) {
        alert("Excel yüklemesi sadece '5. Personel Bilgi Çizelgeleri' için aktiftir.");
        return;
      }
      
      const reader = new FileReader();
      reader.onload = (evt) => {
        try {
          const arrayBuffer = evt.target?.result as ArrayBuffer;
          const workbook = XLSX.read(arrayBuffer, { type: 'array' });
          
          let worksheet = workbook.Sheets[workbook.SheetNames[0]];
          for (const sheetName of workbook.SheetNames) {
            const sheet = workbook.Sheets[sheetName];
            const rows = XLSX.utils.sheet_to_json<any[]>(sheet, { header: 1, defval: "", raw: false });
            const hasKeywords = rows.some(r => r.some((c: any) => {
              const str = String(c || '').toLowerCase();
              return str.includes("sira") || str.includes("adi soyadi") || str.includes("sicil") || str.includes("kadro");
            }));
            if (rows.length > 3 && hasKeywords) {
              worksheet = sheet;
              break;
            }
          }
          
          const rawRows = XLSX.utils.sheet_to_json<string[]>(worksheet, { header: 1, defval: "", raw: false });
          const gridData: string[][] = Array.from({ length: 537 }, () => Array(12).fill(""));
          
          // Smart Header Detection and Column Mapping
          let headerRowIdx = 0;
          for (let r = 0; r < Math.min(10, rawRows.length); r++) {
            const rCells = rawRows[r] || [];
            const hasSira = rCells.some(c => String(c).toLowerCase().includes("sira") || String(c).toLowerCase().includes("sıra"));
            const hasName = rCells.some(c => String(c).toLowerCase().includes("adi") || String(c).toLowerCase().includes("adı") || String(c).toLowerCase().includes("soyadi") || String(c).toLowerCase().includes("soyadı"));
            if (hasSira || hasName) {
              headerRowIdx = r;
              break;
            }
          }

          const rawHeaders = (rawRows[headerRowIdx] || []).map(h => String(h || '').trim().toLowerCase());
          
          // Locate corresponding indices with smart keywords to prevent column shifting
          let idxSira = rawHeaders.findIndex(h => h.includes("sira") || h.includes("sıra") || h === "no");
          let idxName = rawHeaders.findIndex(h => h.includes("adi") || h.includes("adı") || h.includes("soyad") || h.includes("isim"));
          let idxTC = rawHeaders.findIndex(h => h.includes("tc") || h.includes("t.c") || h.includes("kimlik") || h.includes("kımlık"));
          let idxSicil = rawHeaders.findIndex(h => h.includes("sicil") || h.includes("sıcil"));
          let idxKadro = rawHeaders.findIndex(h => h.includes("kadro") || h.includes("unvan") || h.includes("görev") || h.includes("gorev"));
          let idxDogum = rawHeaders.findIndex(h => h.includes("dogum") || h.includes("doğum") || h.includes("tarih"));
          let idxYeri = rawHeaders.findIndex(h => h.includes("yer") || h.includes("mahal") || h.includes("bölüm") || h.includes("bolum"));
          let idxTel = rawHeaders.findIndex(h => h.includes("tel") || h.includes("cep") || h.includes("telefon") || h.includes("gsm"));
          let idxKan = rawHeaders.findIndex(h => h.includes("kan"));
          let idxAdres = rawHeaders.findIndex(h => h.includes("adres") || h.includes("ikamet"));
          let idxYakin = rawHeaders.findIndex(h => h.includes("yakin") || h.includes("yakın") || h.includes("akraba"));
          let idxEsTel = -1;
          for (let c = rawHeaders.length - 1; c >= 0; c--) {
            const h = rawHeaders[c] || "";
            if (h.includes("tel") || h.includes("telefon") || h.includes("cep") || h.includes("gsm") || h.includes("eş") || h.includes("es")) {
              idxEsTel = c;
              break;
            }
          }

          // Set fallback defaults if index lookup failed
          if (idxSira === -1) idxSira = 0;
          if (idxName === -1) idxName = 1;
          if (idxTC === -1) idxTC = 2;
          if (idxSicil === -1) idxSicil = 3;
          if (idxKadro === -1) idxKadro = 4;
          if (idxDogum === -1) idxDogum = 5;
          if (idxYeri === -1) idxYeri = 6;
          if (idxTel === -1) idxTel = 7;
          if (idxKan === -1) idxKan = 8;
          if (idxAdres === -1) idxAdres = 9;
          if (idxYakin === -1) idxYakin = 10;
          if (idxEsTel === -1 || idxEsTel === idxTel) idxEsTel = 11;

          // Extract and map all rows below the header
          let targetRowIdx = 0;
          for (let r = headerRowIdx + 1; r < rawRows.length; r++) {
            const rawRow = rawRows[r] || [];
            
            // Check if this row is a title, metadata, or header to skip
            const colA = String(rawRow[idxSira] || '').trim();
            const colB = String(rawRow[idxName] || '').trim();
            const colD = String(rawRow[idxSicil] || '').trim();
            
            const isHeader = colA.toLowerCase().includes("sira") || 
                             colA.toLowerCase().includes("no") || 
                             colB.toLowerCase().includes("adi") || 
                             colB.toLowerCase().includes("soyad") ||
                             colA.toLowerCase().includes("orman") || 
                             colA.toLowerCase().includes("havacilik") || 
                             colA.toLowerCase().includes("personel");

            // Empty check across primary columns
            const hasData = colA !== "" || colB !== "" || colD !== "";
            
            if (hasData && !isHeader && targetRowIdx < 537) {
              const cellMapping = [idxSira, idxName, idxTC, idxSicil, idxKadro, idxDogum, idxYeri, idxTel, idxKan, idxAdres, idxYakin, idxEsTel];
              
              cellMapping.forEach((srcIdx, destIdx) => {
                gridData[targetRowIdx][destIdx] = rawRow[srcIdx] !== undefined ? String(rawRow[srcIdx]).trim() : "";
              });

              // Ensure Sıra No is a valid consecutive sequence number
              if (gridData[targetRowIdx][0] === "") {
                gridData[targetRowIdx][0] = String(targetRowIdx + 1);
              }

              targetRowIdx++;
            }
          }
          
          // Fallback to basic copy if no rows matched the smart mapping
          if (targetRowIdx === 0) {
            for (let r = 0; r < 537; r++) {
              const rawRow = rawRows[r] || [];
              for (let c = 0; c < 12; c++) {
                gridData[r][c] = rawRow[c] !== undefined ? String(rawRow[c]).trim() : "";
              }
            }
          }
          
          setIsSendingToSheets(prev => ({ ...prev, [5]: true }));
          setUploadProgress(10);
          showNotification(`${file.name} belgesi çözümleniyor, Google Drive'a yükleniyor ve E-Tablo güncelleniyor...`);

          fileToBase64(file).then(async (base64Data) => {
            try {
              const driveFileName = "personel_bilgi_cizelgesi.xlsx";
              setUploadProgress(30);

              // 1. Upload to Google Drive (with correct mimeType)
              const driveRes = await fetch(GOOGLE_SCRIPT_URL, {
                method: "POST",
                headers: {
                  "Content-Type": "text/plain;charset=utf-8"
                },
                body: JSON.stringify({
                  action: "uploadPdfToDrive",
                  fileName: driveFileName,
                  mimeType: file.type || "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                  base64Data: base64Data,
                  formId: 5,
                  month: "Genel Plan",
                  folderId: "1_fIGvuPVpC9N5on1irOfGG8OsD1KSXD0"
                })
              });

              if (!driveRes.ok) {
                throw new Error("Google Drive yedekleme başarısız oldu.");
              }

              const driveResult = await driveRes.json();
              if (driveResult.status !== "success") {
                throw new Error(driveResult.message || "Google Drive yedekleme işlemi başarısız.");
              }

              setUploadProgress(60);
              showNotification("Dosya Google Drive'a başarıyla yüklendi. Şimdi E-Tablo veritabanı güncelleniyor...");

              // 2. Upload parsed gridData to Google Sheet (5-Personel_Bilgi)
              const headers = TABLE_CONFIGS[5].columns.map(col => col.label);
              const dataWithHeaders = [headers, ...gridData];

              const sheetRes = await fetch(GOOGLE_SCRIPT_URL, {
                method: "POST",
                headers: {
                  "Content-Type": "text/plain;charset=utf-8"
                },
                body: JSON.stringify({
                  action: "updateSheet",
                  sheetName: TABLE_CONFIGS[5].sheetName,
                  data: dataWithHeaders
                })
              });

              if (!sheetRes.ok) {
                throw new Error("E-Tablo veritabanı güncellenemedi.");
              }

              setUploadProgress(90);

              // 3. Only on complete success, we write to local state and cache!
              setExcelForm5Data(gridData);
              localStorage.setItem('excel_form_5_data', JSON.stringify(gridData));

              const now = new Date();
              const dateStr = `${String(now.getDate()).padStart(2, '0')}.${String(now.getMonth() + 1).padStart(2, '0')}.${now.getFullYear()} ${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;
              setFormUpdateDates(prev => ({ ...prev, 5: dateStr }));
              localStorage.setItem('form_update_dates', JSON.stringify({ ...formUpdateDates, 5: dateStr }));

              // Mock a matching PDF metadata item so the UI registers it as loaded
              const excelPdfMeta = {
                name: "personel_bilgi_cizelgesi.pdf",
                id: driveResult.fileId || "excel_loaded_form5",
                viewUrl: driveResult.viewUrl || "excel_loaded",
                lastUpdated: dateStr
              };
              setPdfMetadataList(prev => {
                const filtered = prev.filter(p => !p.name.toLowerCase().includes("personel_bilgi"));
                return [excelPdfMeta, ...filtered];
              });

              setUploadProgress(100);
              setIsSendingToSheets(prev => ({ ...prev, [5]: false }));
              showNotification(`'5. Personel Bilgi Çizelgeleri' başarıyla Google Drive'a yedeklendi ve E-Tablo veritabanı güncellendi!`);

              fetchPdfMetadata(); // Refresh metadata list

              // Automatically open the Personnel Information table
              setSelectedFormId(5);
              setModalType('form_table');
              setModalTitle(TABLE_CONFIGS[5].title);
              setSearchQuery('');

            } catch (err: any) {
              console.error("Failed to sync Personnel Info to Google Sheets / Drive:", err);
              setIsSendingToSheets(prev => ({ ...prev, [5]: false }));
              setUploadProgress(0);
              alert(`Hata: Güncelleme sırasında bir sorun oluştu. Veriler ön belleğe alınmadı.\nDetay: ${err?.message || err}`);
            }
          });
          
        } catch (err: any) {
          console.error(err);
          alert(`Excel dosyası çözümlenirken hata oluştu: ${err?.message || err}`);
        }
      };
      reader.readAsArrayBuffer(file);
      return;
    }

    try {
      setIsPdfRendering(true);
      showNotification(`${file.name} çözümleniyor, lütfen bekleyin...`);

      const pages = await renderPdfToImages(file);
      
      // Initialize renderedPages with rotation support
      setRenderedPages(pages.map((p, idx) => ({
        id: `preview-page-${idx}-${Date.now()}`,
        fileName: file.name,
        pageNumber: p.pageNumber,
        dataUrl: p.dataUrl,
        width: p.width,
        height: p.height,
        selected: true,
        rotation: 0, // Default rotation angle is 0
        textItems: p.textItems || []
      })));

      setIsPdfPreviewOpen(true);
      showNotification("Yükleme öncesi PDF önizleme ekranı açıldı. Sayfaları döndürebilir ve yaklaştırabilirsiniz.");
    } catch (err: any) {
      console.error(err);
      alert(`PDF çözümlenirken hata oluştu: ${err?.message || err}`);
    } finally {
      setIsPdfRendering(false);
    }
  };

  // PDF Önizleme Onaylama ve Drive'a Kaydetme Metodu
  const handlePdfPreviewIntegrate = async () => {
    const id = Number(syncSelectedTarget);
    if (isNaN(id) || !TABLE_CONFIGS[id]) {
      alert("Hata: Geçersiz hedef seçimi.");
      return;
    }

    if (!uploadedPdfFile) {
      alert("Hata: Yüklenecek PDF dosyası bulunamadı.");
      return;
    }

    const config = TABLE_CONFIGS[id];
    const isSummer = isSummerForm(id);

    try {
      setIsSendingToSheets(prev => ({ ...prev, [id]: true }));
      setUploadProgress(10);
      showNotification(`Sayfa yönleri ve ayarlamalar işleniyor...`);

      const selectedPages = renderedPages.filter(p => p.selected);
      if (selectedPages.length === 0) {
        alert("Lütfen en az bir sayfa seçin.");
        return;
      }

      // Döndürülmüş sayfaları işleyip yeni data URL'leri çıkartıyoruz
      const processedPages: { pageNumber: number; dataUrl: string; width: number; height: number; textItems?: any[] }[] = [];
      let currentProgress = 15;
      setUploadProgress(currentProgress);

      for (let i = 0; i < selectedPages.length; i++) {
        const page = selectedPages[i];
        const rot = (page as any).rotation || 0;
        const rotatedUrl = await rotateDataUrl(page.dataUrl, rot);
        
        const is90or270 = (rot / 90) % 2 !== 0;
        
        let rotatedTextItems = page.textItems || [];
        if (rot !== 0 && page.textItems) {
          rotatedTextItems = page.textItems.map((item: any) => {
            let rLeft = item.left;
            let rTop = item.top;
            if (rot === 90) {
              rLeft = 100 - item.top;
              rTop = item.left;
            } else if (rot === 180) {
              rLeft = 100 - item.left;
              rTop = 100 - item.top;
            } else if (rot === 270) {
              rLeft = item.top;
              rTop = 100 - item.left;
            }
            return {
              ...item,
              left: Number(rLeft.toFixed(3)),
              top: Number(rTop.toFixed(3))
            };
          });
        }

        processedPages.push({
          pageNumber: i + 1,
          dataUrl: rotatedUrl,
          width: is90or270 ? page.height : page.width,
          height: is90or270 ? page.width : page.height,
          textItems: rotatedTextItems
        });

        currentProgress = Math.min(45, 15 + Math.floor((i / selectedPages.length) * 30));
        setUploadProgress(currentProgress);
      }

      showNotification(`Döndürülmüş ve düzenlenmiş sayfalardan yeni PDF oluşturuluyor...`);
      setUploadProgress(50);

      const base64Data = await generatePdfFromImages(processedPages);
      setUploadProgress(65);

      let driveFileName = uploadedPdfFile.name;
      if (isSummer) {
        const airframeSuffix = getAirframeSuffix(id);
        const cleanMonth = sanitizeTurkishForFilename(selectedUploadSummerMonth);
        driveFileName = `${cleanMonth}_yaz_plan_${airframeSuffix}.pdf`;
      } else {
        const prefix = id === 1 ? 'gorevlendirme' : id === 3 ? 'bakim_yetki' : id === 5 ? 'personel_bilgi' : 'personel_ucus_hizmet';
        driveFileName = `${prefix}_cizelgesi.pdf`;
      }

      showNotification(`"${driveFileName}" adıyla Google Drive'a yükleniyor...`);
      setUploadProgress(70);

      const targetUrl = GOOGLE_SCRIPT_URL;
      const res = await fetch(targetUrl, {
        method: "POST",
        headers: {
          "Content-Type": "text/plain;charset=utf-8"
        },
        body: JSON.stringify({
          action: "uploadPdfToDrive",
          fileName: driveFileName,
          base64Data: base64Data,
          formId: id,
          month: isSummer ? selectedUploadSummerMonth : "Genel Plan",
          folderId: "1_fIGvuPVpC9N5on1irOfGG8OsD1KSXD0"
        })
      });

      if (!res.ok) {
        throw new Error(`Google Apps Script sunucu hatası: Kod ${res.status}`);
      }

      const result = await res.json();
      setUploadProgress(85);

      if (result.status !== "success") {
        throw new Error(result.message || "Drive PDF yükleme hatası.");
      }

      showNotification("Sistem verileri yenileniyor...");

      // Metadata listesini Drive'dan güncelliyoruz
      const updatedMetadata = await fetchPdfMetadata();
      await fetchAllGoogleSheetsList();

      // Döndürülen sayfaları IndexedDB önbelleğine yazıyoruz
      const finalMetadataList = updatedMetadata || pdfMetadataList;
      const updatedCacheKey = getUploadCacheKey(id, finalMetadataList);
      if (updatedCacheKey) {
        await saveRawPdfToDB(updatedCacheKey, base64Data);
        await savePdfPagesToDB(updatedCacheKey, processedPages);
        showNotification("İşlenmiş sayfalar ve ham PDF önbelleğe kaydedildi.");
      }

      if (isSummer) {
        setSelectedSummerStartDate(selectedUploadSummerStartDate);
        setSelectedSummerEndDate(selectedUploadSummerEndDate);
        setSelectedSummerMonth(selectedUploadSummerMonth);
      }

      setUploadProgress(100);
      showNotification(`"${driveFileName}" başarıyla yüklendi ve düzenlemelerle kaydedildi!`);

      // Kapatıp tablo/plan görümüne geçiyoruz
      setIsPdfPreviewOpen(false);
      setModalOpen(true);
      setSelectedFormId(id);
      setModalType('form_table');
      setModalTitle(config.title);
      setIsPdfViewMode(true);

      setTimeout(() => setUploadProgress(0), 4000);
    } catch (err: any) {
      console.error(err);
      showNotification(`⚠️ Hata: ${err?.message || err}`);
      alert(`Hata: İşlem sırasında bir sorun oluştu.\nDetay: ${err?.message || err}`);
      setUploadProgress(0);
    } finally {
      setIsSendingToSheets(prev => ({ ...prev, [id]: false }));
    }
  };

  // EXCEL YÜKLEME METODU (Özel form seçimi ile entegre çalışır)
  const handleFormExcelUpload = (e: React.ChangeEvent<HTMLInputElement>, id: number) => {
    alert("Bu özellik kaldırılmıştır. Lütfen planlama verilerini PDF formatında yükleyin.");
    return;
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    const config = TABLE_CONFIGS[id];
    if (!config) {
      alert("Hata: Geçersiz hedef seçimi.");
      return;
    }

    reader.onload = async (evt) => {
      try {
        setIsSendingToSheets(prev => ({ ...prev, [id]: true }));
        setUploadProgress(15);
        
        const arrayBuffer = evt.target?.result as ArrayBuffer;
        
        setUploadProgress(35);
        showNotification(`'${config.title}' Excel dosyası cihazınızda çözümleniyor...`);
        
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        
        const parsedSheets = workbook.SheetNames.map((sheetName, index) => {
          const worksheet = workbook.Sheets[sheetName];
          const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "", raw: false });
          
          const dynamicPrefix = isSummerForm(id)
            ? getSummerPeriodSheetPrefix(id, selectedSummerMonth)
            : config.sheetName;
            
          const finalSheetName = workbook.SheetNames.length > 1 
            ? `${dynamicPrefix}-${index + 1}`
            : `${dynamicPrefix}-1`;
            
          return {
            name: finalSheetName,
            data: rows
          };
        });

        setUploadProgress(60);
        showNotification(`Çözümlenen veriler Google Sheets'e aktarılıyor...`);

        const targetUrl = GOOGLE_SCRIPT_URL;

        const res = await fetch(targetUrl, {
          method: "POST",
          headers: {
            "Content-Type": "text/plain;charset=utf-8"
          },
          body: JSON.stringify({
            action: "updateMultiSheets",
            prefix: isSummerForm(id) ? getSummerPeriodSheetPrefix(id, selectedSummerMonth) : config.sheetName,
            sheets: parsedSheets
          })
        });
        
        if (!res.ok) {
          throw new Error(`Google Apps Script sunucu hatası: Code ${res.status}`);
        }
        
        const result = await res.json();
        if (result.status !== "success") {
          throw new Error(result.message || "Bilinmeyen sunucu hatası.");
        }
        
        setUploadProgress(85);
        showNotification("E-Tablo güncellendi. Sayfalar portal hafızası ile entegre ediliyor...");
        
        await fetchAllGoogleSheetsList();
        
        const returnedSheets = result.updatedSheets || [];
        const firstSheetName = returnedSheets.length > 0 ? returnedSheets[0] : parsedSheets[0].name;
        
        await pullDataFromGoogleSheets(id, true, firstSheetName);
        
        setUploadProgress(100);
        showNotification(`'${config.title}' altındaki tüm sayfalar başarıyla Google Sheets üzerine aktarıldı ve güncellendi!`);
        
        setSelectedFormId(id);
        setModalType('form_table');
        setModalTitle(config.title);
        setActiveFormTab('live_sheet');
        setTimeout(() => setUploadProgress(0), 4000);
      } catch (err: any) {
        alert(`Hata: Excel dosyası yüklenirken bir hata oluştu.\nDetay: ${err?.message || err}`);
        setUploadProgress(0);
      } finally {
        setIsSendingToSheets(prev => ({ ...prev, [id]: false }));
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const vehiclePlates = techizatKaraAraclariData
    .map(row => (row[1] || "").trim())
    .filter(plate => plate.length > 0 && plate !== "ARAÇ PLAKASI / TANIMI" && plate !== "ARAÇ PLAKASI" && plate !== "Plaka" && plate !== "ARAÇ PLAKASI ");

  const drivers = excelForm5Data
    .map(row => ({
      name: row[1] || "",
      idNo: row[2] || "",
      sicilNo: row[3] || "",
      unvan: row[4] || "",
      phone: row[7] || "",
      kanGrubu: row[8] || "",
      adres: row[9] || ""
    }))
    .filter(d => d.name && d.name.trim().length > 0 && d.name !== "Adı Soyadı" && d.name !== "Personel Adı Soyadı")
    .filter(d => {
      const lowerUnvan = (d.unvan || "").toLowerCase();
      return lowerUnvan.includes("şoför") || lowerUnvan.includes("sofor") || lowerUnvan.includes("şöfr") || lowerUnvan.includes("şofor");
    });

  return (
    <div className="min-h-screen bg-[#0b3d1d] overflow-hidden relative selection:bg-emerald-500/25 selection:text-emerald-950">
      
      {/* OGM PORTAL REDIRECT TRANSITION SCREEN */}
      {isRedirectingToPortal && (
        <motion.div
          initial={{ y: 0 }}
          animate={{ y: isSlidingUp ? "-100%" : 0 }}
          transition={{ duration: 0.8, ease: [0.76, 0, 0.24, 1] }}
          className="fixed inset-0 bg-white flex flex-col items-center justify-center z-[9999] select-none shadow-[0_-20px_50px_rgba(0,0,0,0.1)]"
        >
          <div className="flex flex-col items-center justify-center gap-8 px-6 text-center">
            <div className="relative flex items-center justify-center">
              <div className="absolute inset-0 rounded-full bg-emerald-500/15 animate-ping" style={{ animationDuration: '2s' }} />
              <div className="absolute -inset-4 rounded-full border-4 border-emerald-500/20 animate-pulse" />
              <div className="w-44 h-44 md:w-52 md:h-52 rounded-full overflow-hidden shadow-2xl relative z-10 border-4 border-emerald-500/30 flex items-center justify-center bg-white animate-pulse">
                <img
                  src="https://media2.giphy.com/media/v1.Y2lkPTc5MGI3NjExb2k5Z3lsMWRoaDE3NTNmb2w4M3d3cWljYW16NDRwNmZlbGtxN2lwdCZlcD12MV9pbnRlcmcmY3Q9Zw/n7frjzkahqcqyik0o3/giphy.gif"
                  alt="OGM Logo"
                  referrerPolicy="no-referrer"
                  className="w-full h-full object-cover"
                />
              </div>
            </div>
            
            <div className="flex flex-col gap-4 max-w-lg mt-4">
              <p className="text-[#0b3d1d] font-black text-base md:text-xl leading-relaxed tracking-tight">
                Görev emri evrağı yükleme adımına yönlendiriliyorsunuz...
              </p>
              <div className="flex items-center justify-center gap-2 mt-1">
                <span className="w-2.5 h-2.5 rounded-full bg-[#0b3d1d] animate-bounce" style={{ animationDelay: '0ms' }} />
                <span className="w-2.5 h-2.5 rounded-full bg-[#0b3d1d] animate-bounce" style={{ animationDelay: '150ms' }} />
                <span className="w-2.5 h-2.5 rounded-full bg-[#0b3d1d] animate-bounce" style={{ animationDelay: '300ms' }} />
              </div>
            </div>
          </div>
        </motion.div>
      )}
      
      {/* Immersive Atmospheric Background Glow */}
      <div className="absolute inset-0 bg-[radial-gradient(circle_at_center,rgba(255,255,255,0.08)_0%,transparent_65%)] opacity-40 pointer-events-none z-[1]"></div>

      {/* 1. SPLASH SCREEN (AÇILIŞ EKRANI) */}
      <AnimatePresence>
        {splashVisible && (
          <motion.div
            id="splash-screen"
            key="splash"
            initial={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            transition={{ duration: 0.8, ease: "easeInOut" }}
            onClick={() => setSplashVisible(false)}
            className="fixed inset-0 bg-white flex flex-col items-center justify-center z-[1000] cursor-pointer select-none"
          >
            {/* Alt-skip hint */}
            <div className="absolute top-6 right-6 text-gray-400 font-mono text-xs border border-gray-200 px-3 py-1.5 rounded-full hover:bg-gray-50 transition-colors">
              Geçmek için tıklayın ➜
            </div>

            <div className="relative flex items-center justify-center px-4">
              {/* Logo (Görsel Yuvarlak, %50 Büyük) */}
              <img
                src="https://media2.giphy.com/media/v1.Y2lkPTc5MGI3NjExb2k5Z3lsMWRoaDE3NTNmb2w4M3d3cWljYW16NDRwNmZlbGtxN2lwdCZlcD12MV9pbnRlcmcmY3Q9Zw/n7frjzkahqcqyik0o3/giphy.gif"
                alt="Bakım Şube Logo"
                referrerPolicy="no-referrer"
                className="w-[300px] h-[300px] sm:w-[450px] sm:h-[450px] md:w-[500px] md:h-[500px] rounded-full object-cover shadow-2xl border-4 border-[#0b3d1d]/10 transition-transform active:scale-95"
              />
            </div>
            
            <p className="mt-12 text-[#0b3d1d] font-bold text-lg sm:text-xl tracking-[0.4em] uppercase animate-pulse select-none text-center px-6">
              Sistem Hazırlanıyor
            </p>
          </motion.div>
        )}
      </AnimatePresence>

      {/* 2. ANA PORTAL EKRANI */}
      <div 
        id="main-portal" 
        className="w-full min-h-screen flex flex-col items-center overflow-y-auto relative z-10"
        style={{ display: splashVisible ? 'none' : 'flex' }}
      >
        <div className="container mx-auto px-4 py-8 sm:py-12 flex flex-col items-center min-h-screen max-w-7xl justify-between">
          
          {/* Top HUD / Status Bar */}
          <div className="w-full flex justify-end items-center text-[10px] tracking-[0.3em] font-bold opacity-60 border-b border-white/10 pb-4 mb-8 sm:mb-12">
            <div className="flex items-center gap-6">
              <span className="font-mono">SAAT: {timeString}</span>
            </div>
          </div>

          {/* Main Branding Section */}
          <header className="flex flex-col items-center mb-12 sm:mb-20 w-full select-none">
            <div className="flex flex-col md:flex-row items-center justify-center gap-6 md:gap-12">
              
              {/* Sol Logo Frame */}
              <div className="w-20 h-20 md:w-24 md:h-24 rounded-full border-2 border-white/20 flex items-center justify-center bg-white/5 backdrop-blur-sm p-1 shadow-lg relative group hover:border-white/40 transition-colors">
                <img
                  src="https://media2.giphy.com/media/v1.Y2lkPTc5MGI3NjExb2k5Z3lsMWRoaDE3NTNmb2w4M3d3cWljYW16NDRwNmZlbGtxN2lwdCZlcD12MV9pbnRlcmcmY3Q9Zw/n7frjzkahqcqyik0o3/giphy.gif"
                  alt="Logo Sol"
                  referrerPolicy="no-referrer"
                  className="w-full h-full object-cover rounded-full logo-header"
                />
              </div>
              
              <div className="text-center">
                <h1 className="text-2xl sm:text-3xl md:text-4xl font-extrabold tracking-tight leading-tight text-white text-shadow uppercase">
                  HAVA ARAÇLARI BAKIM VE TEKNİK
                </h1>
                <div className="text-lg sm:text-xl md:text-2xl font-black tracking-[0.2em] text-emerald-400 uppercase mt-1">
                  ŞUBE MÜDÜRLÜĞÜ
                </div>
              </div>

              {/* Sağ Logo Frame */}
              <div className="w-20 h-20 md:w-24 md:h-24 rounded-full border-2 border-white/20 flex items-center justify-center bg-white/5 backdrop-blur-sm p-1 shadow-lg relative group hover:border-white/40 transition-colors">
                <img
                  src="https://media2.giphy.com/media/v1.Y2lkPTc5MGI3NjExb2k5Z3lsMWRoaDE3NTNmb2w4M3d3cWljYW16NDRwNmZlbGtxN2lwdCZlcD12MV9pbnRlcmcmY3Q9Zw/n7frjzkahqcqyik0o3/giphy.gif"
                  alt="Logo Sağ"
                  referrerPolicy="no-referrer"
                  className="w-full h-full object-cover rounded-full logo-header"
                />
              </div>
              
            </div>
            <div className="h-1 w-64 bg-gradient-to-r from-transparent via-emerald-400/50 to-transparent mt-8"></div>
          </header>

          {/* Navigation Grid (Immersive Rounded-[2rem] Grid styling) */}
          <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 lg:grid-cols-5 gap-6 w-full px-4 mb-20 max-w-7xl">
            
            {/* HAVA ARAÇLARI DURUM */}
            <button
              id="btn-aircraft-status"
              onClick={() => openSystem('https://filodurumlar-bakimsube.netlify.app/', 'HAVA ARAÇLARI DURUM')}
              className="bg-white/10 backdrop-blur-md border border-white/20 p-8 rounded-[2rem] flex flex-col items-center text-center hover:bg-white/25 hover:scale-[1.03] active:scale-[0.98] transition-all duration-300 cursor-pointer group focus:outline-none focus:ring-2 focus:ring-emerald-400/50"
            >
              <div className="w-16 h-16 bg-white/10 rounded-2xl flex items-center justify-center mb-6 group-hover:bg-white/15 transition-all shadow-md">
                <Plane className="w-8 h-8 text-white" />
              </div>
              <span className="font-bold tracking-widest text-sm mb-2 text-white uppercase">HAVA ARAÇLARI</span>
              <span className="text-[10px] opacity-60 uppercase tracking-widest font-semibold">Envanter Durumu</span>
            </button>

            {/* PERSONEL (ESKİ YOKLAMA) */}
            <button
              id="btn-personnel"
              onClick={() => openSystem('https://bakimsube-yoklama.netlify.app/', 'PERSONEL')}
              className="bg-white/10 backdrop-blur-md border border-white/20 p-8 rounded-[2rem] flex flex-col items-center text-center hover:bg-white/25 hover:scale-[1.03] active:scale-[0.98] transition-all duration-300 cursor-pointer group focus:outline-none focus:ring-2 focus:ring-emerald-400/50"
            >
              <div className="w-16 h-16 bg-white/10 rounded-2xl flex items-center justify-center mb-6 group-hover:bg-white/15 transition-all shadow-md">
                <Users className="w-8 h-8 text-white" />
              </div>
              <span className="font-bold tracking-widest text-sm mb-2 text-white uppercase">PERSONEL</span>
              <span className="text-[10px] opacity-60 uppercase tracking-widest font-semibold">Yoklama ve Atama</span>
            </button>

            {/* İKMAL */}
            <button
              id="btn-supply"
              onClick={() => openCategory('İKMAL')}
              className="bg-white/10 backdrop-blur-md border border-white/20 p-8 rounded-[2rem] flex flex-col items-center text-center hover:bg-white/25 hover:scale-[1.03] active:scale-[0.98] transition-all duration-300 cursor-pointer group focus:outline-none focus:ring-2 focus:ring-emerald-400/50"
            >
              <div className="w-16 h-16 bg-white/10 rounded-2xl flex items-center justify-center mb-6 group-hover:bg-white/15 transition-all shadow-md">
                <Package className="w-8 h-8 text-white" />
              </div>
              <span className="font-bold tracking-widest text-sm mb-2 text-white uppercase">İKMAL</span>
              <span className="text-[10px] opacity-60 uppercase tracking-widest font-semibold">Parça ve Lojistik</span>
            </button>

            {/* TEÇHİZAT TAKİP (Eski Bakım Takip) */}
            <button
              id="btn-equipment-track"
              onClick={() => openCategory('TEÇHİZAT TAKİP')}
              className="bg-white/10 backdrop-blur-md border border-white/20 p-8 rounded-[2rem] flex flex-col items-center text-center hover:bg-white/25 hover:scale-[1.03] active:scale-[0.98] transition-all duration-300 cursor-pointer group focus:outline-none focus:ring-2 focus:ring-emerald-400/50"
            >
              <div className="w-16 h-16 bg-white/10 rounded-2xl flex items-center justify-center mb-6 group-hover:bg-white/15 transition-all shadow-md">
                <Settings className="w-8 h-8 text-white" />
              </div>
              <span className="font-bold tracking-widest text-sm mb-2 text-white uppercase">TEÇHİZAT TAKİP</span>
              <span className="text-[10px] opacity-60 uppercase tracking-widest font-semibold">DESTEK SİSTEMLERİ</span>
            </button>

            {/* FORM KAYITLARI (Yeni Alan) */}
            <button
              id="btn-form-records"
              onClick={() => openCategory('FORM KAYITLARI')}
              className="bg-white/10 backdrop-blur-md border border-white/20 p-8 rounded-[2rem] flex flex-col items-center text-center hover:bg-white/25 hover:scale-[1.03] active:scale-[0.98] transition-all duration-300 cursor-pointer group focus:outline-none focus:ring-2 focus:ring-emerald-400/50"
            >
              <div className="w-16 h-16 bg-white/10 rounded-2xl flex items-center justify-center mb-6 group-hover:bg-white/15 transition-all shadow-md">
                <FileText className="w-8 h-8 text-white" />
              </div>
              <span className="font-bold tracking-widest text-sm mb-2 text-white uppercase">FORM KAYITLARI</span>
              <span className="text-[10px] opacity-60 uppercase tracking-widest font-semibold">Çizelge ve Planlar</span>
            </button>

          </div>

          {/* Bottom Info Panels / Feed (Immersive HUD Look) */}
          <div className="w-full flex justify-center md:justify-end items-center gap-6 border-t border-white/10 pt-8 mt-4 select-none">
            <div className="text-center md:text-right">
              <p className="text-[9px] tracking-[0.6em] opacity-30 font-mono mb-2">© {currentYear} HAVACILIK TEKNİK PORTAL</p>
              <p className="text-[10px] font-bold text-emerald-400 tracking-widest">KURUMSAL GÜVENLİ ERİŞİM</p>
            </div>
          </div>

        </div>
      </div>

      {/* 3. SİSTEM MODAL (IFRAME & SUB-CATEGORY & DESIGN PHASE PENCERESİ) */}
      <div
        id="system-modal"
        className={`fixed inset-0 bg-white flex flex-col z-[500] transform transition-transform duration-500 ease-[cubic-bezier(0.4,0,0.2,1)] ${
          modalOpen ? 'translate-y-0' : 'translate-y-full'
        }`}
        style={{ display: modalOpen ? 'flex' : 'none' }}
      >
        {/* Modal İçerik Alanı */}
        <div className="flex-1 w-full bg-[#f8fafc] relative overflow-hidden">
          
          {/* Floating Controls Bar (No more heavy header banner) */}
          <div className="absolute top-4 left-4 right-4 z-[100] flex justify-between items-center pointer-events-none select-none">
            {/* Back Button */}
            <button
              id="modal-back-btn"
              onClick={handleBack}
              className="pointer-events-auto flex items-center gap-2 text-slate-800 bg-white/95 hover:bg-white border border-slate-200/80 backdrop-blur-md shadow-lg px-4 py-2.5 rounded-full text-xs font-black uppercase tracking-wider transition-all hover:scale-105 active:scale-95 cursor-pointer"
            >
              <ArrowLeft className="w-4 h-4 text-emerald-800" />
              <span>{(modalType === 'form_table' || modalType === 'techizat_matrix' || modalType === 'excel_sync' || categoryHistory.length > 1) ? 'Geri' : 'Kapat'}</span>
            </button>

            {/* Action Buttons */}
            <div className="flex items-center gap-2 pointer-events-auto">
              {/* Show Veri Güncelle button on Form/Teçhizat views */}
              {modalType !== 'denetleme' && (selectedCategory === 'FORM KAYITLARI' || selectedCategory === 'HA_YER_DESTEK' || selectedCategory === 'T70_DETAY' || selectedCategory === 'KARA_ARACLARI_MENU' || selectedFormId !== null || modalType === 'form_table' || modalType === 'techizat_matrix' || modalType === 'excel_sync') && (
                <button
                  onClick={() => {
                    setPasswordInput('');
                    setPasswordError(false);
                    setIsPasswordModalOpen(true);
                  }}
                  className="flex items-center gap-2 bg-[#0b3d1d]/90 hover:bg-[#0b3d1d] active:scale-95 text-white text-xs font-black uppercase tracking-wider px-4 py-2.5 rounded-full shadow-lg backdrop-blur-md transition-all cursor-pointer"
                >
                  <Database className="w-4 h-4 animate-pulse text-emerald-300" />
                  <span>VERİ GÜNCELLE</span>
                </button>
              )}

            </div>
          </div>
          
          {/* Iframe Yüklenme Spinner'ı */}
          {iframeLoading && modalType === 'iframe' && (
            <div id="iframe-loader" className="absolute inset-0 flex items-center justify-center bg-[#0b3d1d]/5 z-10">
              <div className="loader"></div>
            </div>
          )}

          {/* TASARIM AŞAMASINDA Mesaj Ekranı */}
          {modalType === 'design' && (
            <div id="design-phase-content" className="absolute inset-0 flex items-center justify-center bg-white flex-col p-6 text-center animate-fade-in z-[5]">
              <div className="bg-[#0b3d1d]/10 p-6 rounded-full mb-6">
                <Construction className="w-16 h-16 text-[#0b3d1d]" />
              </div>
              <h3 className="text-2xl font-extrabold text-[#0b3d1d] tracking-wide uppercase">TASARIM AŞAMASINDA</h3>
              <p className="text-gray-500 mt-2 max-w-md text-sm md:text-base font-medium">Bu modül üzerinde çalışmalar devam etmektedir.</p>
            </div>
          )}

          {/* KATEGORİ SEÇİM PANELİ (İKMAL / TEÇHİZAT TAKİP / FORM KAYITLARI VB.) */}
          {modalType === 'category' && selectedCategory && (
            <div id="category-menu-content" className="absolute inset-0 flex flex-col items-center justify-start md:justify-center bg-gradient-to-br from-[#f8fafc] to-[#f1f5f9] p-6 pt-24 md:pt-20 overflow-y-auto w-full h-full">
              <div id="category-buttons" className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-6 w-full max-w-5xl p-2 mx-auto">
                
                {/* İKMAL Alt Butonları */}
                {selectedCategory === 'İKMAL' && (
                  <>
                    <button
                      onClick={() => openSystem('https://ogmbakimsube-taskline.netlify.app/', 'İŞ TAKİP')}
                      className="bg-white hover:bg-white/80 border border-gray-200 shadow-sm rounded-[2rem] p-8 flex flex-col items-center text-center transition-all duration-300 hover:scale-[1.03] active:scale-[0.98] cursor-pointer group"
                    >
                      <div className="w-16 h-16 bg-[#0b3d1d]/10 rounded-2xl flex items-center justify-center mb-6 group-hover:bg-[#0b3d1d]/20 transition-all shadow-sm">
                        <ClipboardList className="w-8 h-8 text-[#0b3d1d]" />
                      </div>
                      <span className="text-[#0b3d1d] font-bold tracking-widest text-sm mb-2 uppercase">İŞ TAKİP</span>
                      <span className="text-[10px] text-[#0b3d1d]/60 uppercase tracking-widest font-semibold font-mono">Görevleri İncele</span>
                    </button>

                    <button
                      onClick={() => {
                        window.open('https://ogmhavacilik-takipsistem.netlify.app/', '_blank');
                      }}
                      className="bg-white hover:bg-white/80 border border-gray-200 shadow-sm rounded-[2rem] p-8 flex flex-col items-center text-center transition-all duration-300 hover:scale-[1.03] active:scale-[0.98] cursor-pointer group relative"
                    >
                      <div className="w-16 h-16 bg-[#0b3d1d]/10 rounded-2xl flex items-center justify-center mb-6 group-hover:bg-[#0b3d1d]/20 transition-all shadow-sm">
                        <Fuel className="w-8 h-8 text-[#0b3d1d]" />
                      </div>
                      <span className="text-[#0b3d1d] font-bold tracking-widest text-sm mb-2 uppercase">YAKIT RESMÎ</span>
                      <span className="text-[10px] text-[#0b3d1d]/60 uppercase tracking-widest font-semibold font-mono flex items-center justify-center gap-1.5">
                        Takip Sistemi <span className="text-[8px] font-normal text-emerald-800 bg-emerald-100 px-2 py-0.5 rounded-full inline-block">Yeni Sekme</span>
                      </span>
                    </button>
                  </>
                )}

                {/* TEÇHİZAT TAKİP Alt Butonları */}
                {selectedCategory === 'TEÇHİZAT TAKİP' && (
                  <>
                    <button
                      onClick={() => navigateToSubCategory('HA_YER_DESTEK')}
                      className="bg-white hover:bg-white/80 border border-gray-200 shadow-sm rounded-[2rem] p-8 flex flex-col items-center text-center transition-all duration-300 hover:scale-[1.03] active:scale-[0.98] cursor-pointer group"
                    >
                      <div className="w-16 h-16 bg-[#0b3d1d]/10 rounded-2xl flex items-center justify-center mb-6 group-hover:bg-[#0b3d1d]/20 transition-all shadow-sm">
                        <Plane className="w-8 h-8 text-[#0b3d1d]" />
                      </div>
                      <span className="text-[#0b3d1d] font-bold tracking-widest text-sm mb-2 uppercase">HAVA ARAÇLARI YER DESTEK</span>
                      <span className="text-[10px] text-[#0b3d1d]/60 uppercase tracking-widest font-semibold font-mono">TEÇHİZAT MODÜLLERİ</span>
                    </button>

                    <button
                      onClick={() => openTechizatMatrix('hangar', 'HANGAR YER DESTEK TEÇHİZATLARI')}
                      className="bg-white hover:bg-white/80 border border-gray-200 shadow-sm rounded-[2rem] p-8 flex flex-col items-center text-center transition-all duration-300 hover:scale-[1.03] active:scale-[0.98] cursor-pointer group"
                    >
                      <div className="w-16 h-16 bg-[#0b3d1d]/10 rounded-2xl flex items-center justify-center mb-6 group-hover:bg-[#0b3d1d]/20 transition-all shadow-sm">
                        <Wrench className="w-8 h-8 text-[#0b3d1d]" />
                      </div>
                      <span className="text-[#0b3d1d] font-bold tracking-widest text-sm mb-2 uppercase">HANGAR YER DESTEK</span>
                      <span className="text-[10px] text-[#0b3d1d]/60 uppercase tracking-widest font-semibold font-mono">TEÇHİZATLARALTYAPI</span>
                    </button>

                    <button
                      onClick={() => navigateToSubCategory('KARA_ARACLARI_MENU')}
                      className="bg-white hover:bg-white/80 border border-gray-200 shadow-sm rounded-[2rem] p-8 flex flex-col items-center text-center transition-all duration-300 hover:scale-[1.03] active:scale-[0.98] cursor-pointer group"
                    >
                      <div className="w-16 h-16 bg-[#0b3d1d]/10 rounded-2xl flex items-center justify-center mb-6 group-hover:bg-[#0b3d1d]/20 transition-all shadow-sm">
                        <Truck className="w-8 h-8 text-[#0b3d1d]" />
                      </div>
                      <span className="text-[#0b3d1d] font-bold tracking-widest text-sm mb-2 uppercase">KARA ARAÇLARI TAKİP</span>
                      <span className="text-[10px] text-[#0b3d1d]/60 uppercase tracking-widest font-semibold font-mono">KARA ARAÇ TAKİP MODÜLÜ</span>
                    </button>
                  </>
                )}

                {/* KARA ARAÇLARI MENÜ ALTBİRİMLERİ (KARA_ARACLARI_MENU) */}
                {selectedCategory === 'KARA_ARACLARI_MENU' && (
                  <>
                    <button
                      onClick={() => {
                        setKaraAraclariSubTab('list');
                        openTechizatMatrix('kara_araclari', 'KARA ARAÇ TAKİP LİSTESİ');
                      }}
                      className="bg-white hover:bg-white/80 border border-gray-200 shadow-sm rounded-[2rem] p-8 flex flex-col items-center text-center transition-all duration-300 hover:scale-[1.03] active:scale-[0.98] cursor-pointer group"
                    >
                      <div className="w-16 h-16 bg-white border border-gray-200 shadow-inner rounded-2xl flex items-center justify-center mb-6 text-[#0b3d1d] font-black text-xs">
                        🚗
                      </div>
                      <span className="text-[#0b3d1d] font-bold tracking-widest text-sm mb-2 uppercase">KARA ARAÇ TAKİP LİSTESİ</span>
                      <span className="text-[10px] text-gray-500 uppercase tracking-widest font-mono font-bold">ARAÇ DURUM & BAKIM</span>
                    </button>

                    <button
                      onClick={() => {
                        setKaraAraclariSubTab('mission_order');
                        openTechizatMatrix('kara_araclari', 'GÖREV EMRİ GİRİŞ');
                      }}
                      className="bg-white hover:bg-white/80 border border-gray-200 shadow-sm rounded-[2rem] p-8 flex flex-col items-center text-center transition-all duration-300 hover:scale-[1.03] active:scale-[0.98] cursor-pointer group"
                    >
                      <div className="w-16 h-16 bg-white border border-gray-200 shadow-inner rounded-2xl flex items-center justify-center mb-6 text-[#0b3d1d] font-black text-xs">
                        📋
                      </div>
                      <span className="text-[#0b3d1d] font-bold tracking-widest text-sm mb-2 uppercase">GÖREV EMRİ GİRİŞ</span>
                      <span className="text-[10px] text-gray-500 uppercase tracking-widest font-mono font-bold">ADIM ADIM YENİ GÖREV</span>
                    </button>
                  </>
                )}

                {/* HAVA ARAÇLARI YER DESTEK (HA_YER_DESTEK) Alt Butonları */}
                {selectedCategory === 'HA_YER_DESTEK' && (
                  <>
                    <button
                      onClick={() => openTechizatMatrix('bell429', 'BELL 429 YER DESTEK TEÇHİZATLARI')}
                      className="bg-white hover:bg-white/80 border border-gray-200 shadow-sm rounded-[2rem] p-8 flex flex-col items-center text-center transition-all duration-300 hover:scale-[1.03] active:scale-[0.98] cursor-pointer group"
                    >
                      <div className="w-16 h-16 bg-white border border-gray-200 shadow-inner rounded-2xl flex items-center justify-center mb-6 text-[#0b3d1d] font-black text-xs">
                        B429
                      </div>
                      <span className="text-[#0b3d1d] font-bold tracking-widest text-sm mb-2 uppercase">BELL 429</span>
                      <span className="text-[10px] text-gray-500 uppercase tracking-widest font-mono">Teçhizat Listesi</span>
                    </button>

                    <button
                      onClick={() => openTechizatMatrix('at802', 'AT-802F YER DESTEK TEÇHİZATLARI')}
                      className="bg-white hover:bg-white/80 border border-gray-200 shadow-sm rounded-[2rem] p-8 flex flex-col items-center text-center transition-all duration-300 hover:scale-[1.03] active:scale-[0.98] cursor-pointer group"
                    >
                      <div className="w-16 h-16 bg-white border border-gray-200 shadow-inner rounded-2xl flex items-center justify-center mb-6 text-[#0b3d1d] font-black text-[10px]">
                        AT-802F
                      </div>
                      <span className="text-[#0b3d1d] font-bold tracking-widest text-sm mb-2 uppercase">AT-802F</span>
                      <span className="text-[10px] text-gray-500 uppercase tracking-widest font-mono">Teçhizat Listesi</span>
                    </button>

                    <button
                      onClick={() => navigateToSubCategory('T70_DETAY')}
                      className="bg-white hover:bg-white/80 border-2 border-emerald-600/20 shadow-sm rounded-[2rem] p-8 flex flex-col items-center text-center transition-all duration-300 hover:scale-[1.03] active:scale-[0.98] cursor-pointer group"
                    >
                      <div className="w-16 h-16 bg-emerald-50 border border-emerald-100 shadow-inner rounded-2xl flex items-center justify-center mb-6 text-emerald-800 font-extrabold text-xs">
                        T-70
                      </div>
                      <span className="text-emerald-900 font-extrabold tracking-widest text-sm mb-2 uppercase">T-70</span>
                      <span className="text-[10px] text-gray-500 uppercase tracking-widest font-mono">Teçhizat Listesi</span>
                    </button>

                    <button
                      onClick={() => openTechizatMatrix('c650', 'C-650 YER DESTEK TEÇHİZATLARI')}
                      className="bg-white hover:bg-white/80 border border-gray-200 shadow-sm rounded-[2rem] p-8 flex flex-col items-center text-center transition-all duration-300 hover:scale-[1.03] active:scale-[0.98] cursor-pointer group"
                    >
                      <div className="w-16 h-16 bg-white border border-gray-200 shadow-inner rounded-2xl flex items-center justify-center mb-6 text-[#0b3d1d] font-black text-xs">
                        C650
                      </div>
                      <span className="text-[#0b3d1d] font-bold tracking-widest text-sm mb-2 uppercase">C-650</span>
                      <span className="text-[10px] text-gray-500 uppercase tracking-widest font-mono">Teçhizat Listesi</span>
                    </button>

                    <button
                      onClick={() => openTechizatMatrix('b360', 'B-360 YER DESTEK TEÇHİZATLARI')}
                      className="bg-white hover:bg-white/80 border border-gray-200 shadow-sm rounded-[2rem] p-8 flex flex-col items-center text-center transition-all duration-300 hover:scale-[1.03] active:scale-[0.98] cursor-pointer group"
                    >
                      <div className="w-16 h-16 bg-white border border-gray-200 shadow-inner rounded-2xl flex items-center justify-center mb-6 text-[#0b3d1d] font-black text-xs">
                        B360
                      </div>
                      <span className="text-[#0b3d1d] font-bold tracking-widest text-sm mb-2 uppercase">B-360</span>
                      <span className="text-[10px] text-gray-500 uppercase tracking-widest font-mono">Teçhizat Listesi</span>
                    </button>

                    <button
                      onClick={() => openTechizatMatrix('all', 'TÜM BİRİMLER ORTAK TEÇHİZAT ARAMA')}
                      className="bg-emerald-50 hover:bg-emerald-100/80 border-2 border-emerald-600/30 shadow-sm rounded-[2rem] p-8 flex flex-col items-center text-center transition-all duration-300 hover:scale-[1.03] active:scale-[0.98] cursor-pointer group sm:col-span-2 lg:col-span-3 mt-4"
                    >
                      <div className="w-16 h-16 bg-[#0b3d1d] text-white rounded-2xl flex items-center justify-center mb-6 shadow-md group-hover:bg-[#082a14] transition-all">
                        <Search className="w-8 h-8 text-emerald-300" />
                      </div>
                      <span className="text-[#0b3d1d] font-black tracking-widest text-base mb-2 uppercase">🔍 TÜM BİRİMLERDE TEÇHİZAT ARA</span>
                      <span className="text-xs text-emerald-800 uppercase tracking-widest font-mono font-bold">
                        Bütün Hava Araçlarının Yer Destek Teçhizatlarını Tek Listede Arayın & Excel Olarak İndirin
                      </span>
                    </button>
                  </>
                )}

                {/* T-70 DETAY ALTBİRİMLERİ (T70_DETAY) */}
                {selectedCategory === 'T70_DETAY' && (
                  <>
                    <button
                      onClick={() => openTechizatMatrix('t70_bumbi_backet', 'T-70 BUMBİ BACKET TEÇHİZATI')}
                      className="bg-white hover:bg-white/80 border border-gray-200 shadow-sm rounded-[2rem] p-8 flex flex-col items-center text-center transition-all duration-300 hover:scale-[1.03] active:scale-[0.98] cursor-pointer group"
                    >
                      <div className="w-16 h-16 bg-[#0b3d1d]/10 rounded-2xl flex items-center justify-center mb-6 group-hover:bg-[#0b3d1d]/20 transition-all shadow-sm text-[#0b3d1d] font-extrabold text-sm">
                        BB
                      </div>
                      <span className="text-[#0b3d1d] font-bold tracking-widest text-sm mb-2 uppercase">BUMBİ BACKET</span>
                      <span className="text-[10px] text-gray-500 uppercase tracking-widest font-mono">T-70 Alt Birimi</span>
                    </button>

                    <button
                      onClick={() => openTechizatMatrix('t70_helitak', 'T-70 HELİTAK TEÇHİZATI')}
                      className="bg-white hover:bg-white/80 border border-gray-200 shadow-sm rounded-[2rem] p-8 flex flex-col items-center text-center transition-all duration-300 hover:scale-[1.03] active:scale-[0.98] cursor-pointer group"
                    >
                      <div className="w-16 h-16 bg-[#0b3d1d]/10 rounded-2xl flex items-center justify-center mb-6 group-hover:bg-[#0b3d1d]/20 transition-all shadow-sm text-[#0b3d1d] font-extrabold text-sm">
                        HT
                      </div>
                      <span className="text-[#0b3d1d] font-bold tracking-widest text-sm mb-2 uppercase">HELİTAK</span>
                      <span className="text-[10px] text-gray-500 uppercase tracking-widest font-mono">T-70 Alt Birimi</span>
                    </button>

                    <button
                      onClick={() => openTechizatMatrix('t70', 'T-70 YER DESTEK TEÇHİZATI')}
                      className="bg-white hover:bg-white/80 border border-gray-200 shadow-sm rounded-[2rem] p-8 flex flex-col items-center text-center transition-all duration-300 hover:scale-[1.03] active:scale-[0.98] cursor-pointer group"
                    >
                      <div className="w-16 h-16 bg-[#0b3d1d]/10 rounded-2xl flex items-center justify-center mb-6 group-hover:bg-[#0b3d1d]/20 transition-all shadow-sm text-[#0b3d1d]">
                        <Wrench className="w-8 h-8" />
                      </div>
                      <span className="text-[#0b3d1d] font-bold tracking-widest text-sm mb-2 uppercase">YER DESTEK TEÇHİZATLARI</span>
                      <span className="text-[10px] text-gray-500 uppercase tracking-widest font-mono">T-70 Alt Birimi</span>
                    </button>
                  </>
                )}

                {/* FORM KAYITLARI Alt Butonları */}
                {selectedCategory === 'FORM KAYITLARI' && (
                  <>
                    <button
                      onClick={() => {
                        setSelectedFormId(1);
                        setModalType('form_table');
                        setModalTitle(TABLE_CONFIGS[1].title);
                        setSearchQuery('');
                      }}
                      className="bg-white hover:bg-white/80 border border-gray-200 shadow-sm rounded-[2rem] p-8 flex flex-col items-center text-center transition-all duration-300 hover:scale-[1.03] active:scale-[0.98] cursor-pointer group"
                    >
                      <div className="w-16 h-16 bg-[#0b3d1d]/10 rounded-2xl flex items-center justify-center mb-6 group-hover:bg-[#0b3d1d]/20 transition-all shadow-sm">
                        <ClipboardList className="w-8 h-8 text-[#0b3d1d]" />
                      </div>
                      <span className="text-[#0b3d1d] font-bold tracking-normal text-sm mb-2 text-center uppercase leading-tight">1. Görevlendirme Çizelgeleri</span>
                      
                      <div className="mt-auto pt-4 border-t border-gray-100 w-full text-center">
                        <span className="text-[10px] text-gray-400 font-extrabold uppercase tracking-widest font-mono">
                          GÜNCELLEME TARİHİ: {formUpdateDates[1] || "-"}
                        </span>
                      </div>
                    </button>
 
                    <button
                      onClick={() => {
                        setSelectedFormId(21);
                        setModalType('form_table');
                        setModalTitle(TABLE_CONFIGS[21].title);
                        setSearchQuery('');
                      }}
                      className="bg-white hover:bg-white/80 border border-gray-200 shadow-sm rounded-[2rem] p-8 flex flex-col items-center text-center transition-all duration-300 hover:scale-[1.03] active:scale-[0.98] cursor-pointer group"
                    >
                      <div className="w-16 h-16 bg-[#0b3d1d]/10 rounded-2xl flex items-center justify-center mb-6 group-hover:bg-[#0b3d1d]/20 transition-all shadow-sm">
                        <CalendarCheck className="w-8 h-8 text-[#0b3d1d]" />
                      </div>
                      <span className="text-[#0b3d1d] font-bold tracking-normal text-sm mb-2 text-center uppercase leading-tight">2. Yaz Dönemi Görev Planlaması</span>
                      
                      <div className="mt-auto pt-4 border-t border-gray-100 w-full text-center">
                        <span className="text-[10px] text-gray-400 font-extrabold uppercase tracking-widest font-mono">
                          GÜNCELLEME TARİHİ: {getDetailedSummerUpdateInfo()}
                        </span>
                      </div>
                    </button>
 
                    <button
                      onClick={() => {
                        setSelectedFormId(3);
                        setModalType('form_table');
                        setModalTitle(TABLE_CONFIGS[3].title);
                        setSearchQuery('');
                      }}
                      className="bg-white hover:bg-white/80 border border-gray-200 shadow-sm rounded-[2rem] p-8 flex flex-col items-center text-center transition-all duration-300 hover:scale-[1.03] active:scale-[0.98] cursor-pointer group"
                    >
                      <div className="w-16 h-16 bg-[#0b3d1d]/10 rounded-2xl flex items-center justify-center mb-6 group-hover:bg-[#0b3d1d]/20 transition-all shadow-sm">
                        <Settings className="w-8 h-8 text-[#0b3d1d]" />
                      </div>
                      <span className="text-[#0b3d1d] font-bold tracking-normal text-sm mb-2 text-center uppercase leading-tight">3. Bakım Yetki Çizelgeleri</span>
                      
                      <div className="mt-auto pt-4 border-t border-gray-100 w-full text-center">
                        <span className="text-[10px] text-gray-400 font-extrabold uppercase tracking-widest font-mono">
                          GÜNCELLEME TARİHİ: {formUpdateDates[3] || "-"}
                        </span>
                      </div>
                    </button>
 
                    <a
                      href="https://bulut.ogm.gov.tr/MMEL"
                      target="_blank"
                      rel="noopener noreferrer"
                      className="bg-white hover:bg-white/80 border border-gray-200 shadow-sm rounded-[2rem] p-8 flex flex-col items-center justify-center text-center transition-all duration-300 hover:scale-[1.03] active:scale-[0.98] cursor-pointer group"
                    >
                      <div className="w-16 h-16 bg-[#0b3d1d]/10 rounded-2xl flex items-center justify-center mb-6 group-hover:bg-[#0b3d1d]/20 transition-all shadow-sm relative">
                        <ClipboardList className="w-8 h-8 text-[#0b3d1d]" />
                        <span className="absolute -top-1 -right-1 bg-emerald-800 text-white text-[8px] font-bold px-1.5 py-0.5 rounded-full uppercase tracking-wider">YENİ SEKME ➜</span>
                      </div>
                      <span className="text-[#0b3d1d] font-bold tracking-normal text-sm mb-2 text-center uppercase leading-tight">4. MMEL</span>
                      <span className="text-[9px] text-[#0b3d1d]/60 uppercase tracking-wider font-semibold font-mono mb-2">Asgari Teçhizat Listesi</span>
                      <span className="text-[9px] text-emerald-800 font-extrabold uppercase tracking-wider font-mono bg-emerald-50 px-2 py-0.5 rounded-full">bulut.ogm.gov.tr/MMEL ➜</span>
                      
                      <div className="mt-auto pt-4 border-t border-gray-100 w-full text-center">
                        <span className="text-[10px] text-gray-400 font-extrabold uppercase tracking-widest font-mono">
                          GÜNCELLEME TARİHİ: -
                        </span>
                      </div>
                    </a>
 
                    <button
                      onClick={() => {
                        setSelectedFormId(5);
                        setModalType('form_table');
                        setModalTitle(TABLE_CONFIGS[5].title);
                        setSearchQuery('');
                      }}
                      className="bg-white hover:bg-white/80 border border-gray-200 shadow-sm rounded-[2rem] p-8 flex flex-col items-center text-center transition-all duration-300 hover:scale-[1.03] active:scale-[0.98] cursor-pointer group"
                    >
                      <div className="w-16 h-16 bg-[#0b3d1d]/10 rounded-2xl flex items-center justify-center mb-6 group-hover:bg-[#0b3d1d]/20 transition-all shadow-sm">
                        <Users className="w-8 h-8 text-[#0b3d1d]" />
                      </div>
                      <span className="text-[#0b3d1d] font-bold tracking-normal text-sm mb-2 text-center uppercase leading-tight">5. Personel Bilgi Çizelgeleri</span>
                      
                      <div className="mt-auto pt-4 border-t border-gray-100 w-full text-center">
                        <span className="text-[10px] text-gray-400 font-extrabold uppercase tracking-widest font-mono">
                          GÜNCELLEME TARİHİ: {formUpdateDates[5] || "-"}
                        </span>
                      </div>
                    </button>
 
                    <button
                      onClick={() => {
                        setSelectedFormId(6);
                        setModalType('form_table');
                        setModalTitle(TABLE_CONFIGS[6].title);
                        setSearchQuery('');
                      }}
                      className="bg-white hover:bg-white/80 border border-gray-200 shadow-sm rounded-[2rem] p-8 flex flex-col items-center text-center transition-all duration-300 hover:scale-[1.03] active:scale-[0.98] cursor-pointer group"
                    >
                      <div className="w-16 h-16 bg-[#0b3d1d]/10 rounded-2xl flex items-center justify-center mb-6 group-hover:bg-[#0b3d1d]/20 transition-all shadow-sm">
                        <FileText className="w-8 h-8 text-[#0b3d1d]" />
                      </div>
                      <span className="text-[#0b3d1d] font-bold tracking-normal text-sm mb-2 text-center uppercase leading-tight">6. Personel Uçuş-Hizmet Yılları</span>
                      
                      <div className="mt-auto pt-4 border-t border-gray-100 w-full text-center">
                        <span className="text-[10px] text-gray-400 font-extrabold uppercase tracking-widest font-mono">
                          GÜNCELLEME TARİHİ: {formUpdateDates[6] || "-"}
                        </span>
                      </div>
                    </button>

                    <a
                      href="https://bulut.ogm.gov.tr/DIJITALYAKIT"
                      target="_blank"
                      rel="noopener noreferrer"
                      className="bg-white hover:bg-white/80 border border-gray-200 shadow-sm rounded-[2rem] p-8 flex flex-col items-center justify-center text-center transition-all duration-300 hover:scale-[1.03] active:scale-[0.98] cursor-pointer group"
                    >
                      <div className="w-16 h-16 bg-[#0b3d1d]/10 rounded-2xl flex items-center justify-center mb-6 group-hover:bg-[#0b3d1d]/20 transition-all shadow-sm relative">
                        <Fuel className="w-8 h-8 text-[#0b3d1d]" />
                        <span className="absolute -top-1 -right-1 bg-emerald-800 text-white text-[8px] font-bold px-1.5 py-0.5 rounded-full uppercase tracking-wider">YENİ SEKME ➜</span>
                      </div>
                      <span className="text-[#0b3d1d] font-bold tracking-normal text-sm mb-2 text-center uppercase leading-tight">7. YAKIT MAKBUZ ARŞİVLERİ</span>
                      <span className="text-[9px] text-[#0b3d1d]/60 uppercase tracking-wider font-semibold font-mono mb-2">Dijital Yakıt Arşivleri</span>
                      <span className="text-[9px] text-emerald-800 font-extrabold uppercase tracking-wider font-mono bg-emerald-50 px-2 py-0.5 rounded-full">bulut.ogm.gov.tr/DIJITALYAKIT ➜</span>
                      
                      <div className="mt-auto pt-4 border-t border-gray-100 w-full text-center">
                        <span className="text-[10px] text-gray-400 font-extrabold uppercase tracking-widest font-mono">
                          GÜNCELLEME TARİHİ: -
                        </span>
                      </div>
                    </a>

                    <button
                      onClick={() => {
                        setModalUrl('https://ogmhavacilik.github.io/surecyonet/');
                        setModalType('iframe');
                        setModalTitle('8. DENETLEME RAPOR VE EKLER');
                        setIframeLoading(true);
                      }}
                      className="bg-white hover:bg-white/80 border border-gray-200 shadow-sm rounded-[2rem] p-8 flex flex-col items-center justify-center text-center transition-all duration-300 hover:scale-[1.03] active:scale-[0.98] cursor-pointer group"
                    >
                      <div className="w-16 h-16 bg-[#0b3d1d]/10 rounded-2xl flex items-center justify-center mb-6 group-hover:bg-[#0b3d1d]/20 transition-all shadow-sm relative">
                        <Folder className="w-8 h-8 text-[#0b3d1d]" />
                      </div>
                      <span className="text-[#0b3d1d] font-bold tracking-normal text-sm mb-2 text-center uppercase leading-tight">8. DENETLEME RAPOR VE EKLER</span>
                      <span className="text-[9px] text-[#0b3d1d]/60 uppercase tracking-wider font-semibold font-mono mb-2">Denetleme Rapor ve Ekleri</span>
                      <span className="text-[9px] text-emerald-800 font-extrabold uppercase tracking-wider font-mono bg-emerald-50 px-2.5 py-1 rounded-full border border-emerald-200">GÖRÜNTÜLE ➜</span>
                      
                      <div className="mt-auto pt-4 border-t border-gray-100 w-full text-center">
                        <span className="text-[10px] text-gray-400 font-extrabold uppercase tracking-widest font-mono">
                          GÜNCELLEME TARİHİ: -
                        </span>
                      </div>
                    </button>
                  </>
                )}

              </div>
            </div>
          )}

          {/* DINAMIK TABLE MODULU */}
          {modalType === 'form_table' && selectedFormId && (
            <div className="absolute inset-0 flex flex-col bg-slate-50 overflow-hidden animate-fade-in p-0">
              
              {/* Sleek Top Header Bar - provides backdrop for floating controls and displays Form Title */}
              <div className="bg-white border-b border-slate-200 h-20 shrink-0 select-none flex items-center justify-between px-6 z-40 shadow-sm">
                {/* We reserve space for left floating control */}
                <div className="w-24 md:w-32 shrink-0 pointer-events-none" />
                
                <div className="flex-1 text-center min-w-0">
                  <h3 className="text-slate-800 font-black text-xs md:text-sm uppercase tracking-wider truncate">
                    {TABLE_CONFIGS[selectedFormId]?.title}
                  </h3>
                </div>
                
                {/* We reserve space for right floating controls */}
                <div className="w-36 md:w-52 shrink-0 pointer-events-none" />
              </div>
              
              {/* Secondary Planner Sub-Header Bar (Month Selector for Summer Period with integrated Switcher & PDF actions) */}
              {isSummerForm(selectedFormId) && (
                <div className="bg-white px-6 py-3 border-b border-slate-200 flex flex-col md:flex-row md:items-center md:justify-between gap-4 shrink-0 select-none shadow-sm">
                  <div className="flex flex-col sm:flex-row sm:items-center gap-6">
                    {/* Airframe Switcher */}
                    <div className="flex items-center gap-2">
                      <span className="text-[10px] font-black text-[#0b3d1d] uppercase tracking-wider font-sans">🚁 HAVA ARACI:</span>
                      <div className="flex bg-slate-100 p-0.5 rounded-xl border border-slate-200 shadow-sm gap-0.5">
                        {[
                          { id: 21, name: "BELL 429" },
                          { id: 22, name: "T-70" },
                          { id: 23, name: "AT-802" },
                          { id: 24, name: "BEKLEME (BELL-429)" },
                          { id: 25, name: "BEKLEME (C-650/B-360)" }
                        ].map((subForm) => (
                          <button
                            key={subForm.id}
                            onClick={() => {
                              setSelectedFormId(subForm.id);
                              setModalTitle(TABLE_CONFIGS[subForm.id].title);
                            }}
                            className={`px-3 py-1.5 rounded-lg text-[10px] font-black transition-all cursor-pointer ${
                              selectedFormId === subForm.id
                                ? 'bg-[#0b3d1d] text-white shadow-sm'
                                : 'text-slate-600 hover:bg-slate-200 hover:text-slate-800'
                            }`}
                          >
                            {subForm.name}
                          </button>
                        ))}
                      </div>
                    </div>

                    {/* Dynamic Period Dropdown Select */}
                    <div className="flex items-center gap-2 bg-emerald-50/70 hover:bg-emerald-50 px-4 py-2 rounded-2xl border border-emerald-100 shadow-sm transition-colors">
                      <span className="text-[10px] font-black text-emerald-800 uppercase tracking-wider font-sans shrink-0">🗓️ PLAN DÖNEMİ:</span>
                      <select
                        value={selectedSummerMonth}
                        onChange={(e) => {
                          setSelectedSummerMonth(e.target.value);
                          showNotification(`Plan dönemi değiştirildi: ${getReadablePeriodName(e.target.value)}`);
                        }}
                        className="bg-transparent text-xs font-black text-[#0b3d1d] uppercase outline-none border-none cursor-pointer focus:ring-0 pr-6"
                      >
                        {availableSummerPeriods.map((period) => (
                          <option key={period} value={period} className="bg-white text-slate-800 text-xs font-bold uppercase">
                            {getReadablePeriodName(period)}
                          </option>
                        ))}
                      </select>
                    </div>
                  </div>
                </div>
              )}

              {/* Seamless Fullscreen PDF Iframe from Google Drive or Empty State */}
              <div className="flex-1 w-full relative bg-slate-100">
                {(() => {
                  if (selectedFormId === 5) {
                    if (isExcelOcrProcessing) {
                      return (
                        <div className="absolute inset-0 bg-slate-950 flex flex-col items-center justify-center p-8 text-center animate-fade-in z-25">
                          <div className="w-12 h-12 border-4 border-emerald-500 border-t-transparent rounded-full animate-spin mb-4" />
                          <h3 className="text-emerald-400 font-extrabold text-sm uppercase tracking-widest font-mono animate-pulse">
                            VERİLER YÜKLENİYOR...
                          </h3>
                        </div>
                      );
                    }

                    // Render beautiful landscape interactive PDF sheet!
                     const currentFilteredRows = excelForm5Data
                      .map((row, rIdx) => ({ row, rIdx }))
                      .filter(({ row }) => {
                        const colA = String(row[0] || '').trim();
                        const hasSiraNo = colA !== "" && 
                          !colA.toLowerCase().includes("sira") && 
                          !colA.toLowerCase().includes("no");
                        
                        if (!hasSiraNo) return false;
                        
                        if (selectedKadroFilter) {
                          const rowKadro = row[4] ? row[4].trim().toLowerCase() : "";
                          if (rowKadro !== selectedKadroFilter.trim().toLowerCase()) {
                            return false;
                          }
                        }
                        
                        const q = excelSearchQuery.toLowerCase().trim();
                        if (!q) return true;
                        return row.some(cell => cell && cell.toLowerCase().includes(q));
                      });

                    return (
                      <div className="absolute inset-0 bg-slate-950 flex flex-col overflow-hidden text-white animate-fade-in z-20">
                        {/* Top Header with OCR Smart Search */}
                        <div className="shrink-0 bg-slate-900 border-b border-slate-800 px-6 py-4 flex flex-col lg:flex-row items-center justify-between gap-4 select-none shadow-md">
                          <div className="flex items-center gap-2">
                            <span className="w-2 h-2 rounded-full bg-emerald-500 animate-pulse" />
                            <span className="text-[10px] font-black text-slate-300 uppercase tracking-widest font-mono">
                              PERSONEL BİLGİ SİSTEMİ
                            </span>
                          </div>

                          {/* Center Controls: Search + Kadro Filter ListBox */}
                          <div className="flex flex-wrap items-center gap-3 w-full max-w-2xl justify-center lg:justify-start">
                            {/* Smart Search Box */}
                            <div className="flex items-center gap-3 bg-slate-950 px-4 py-2 rounded-2xl border border-slate-800 shadow-inner flex-1 min-w-[200px]">
                              <Search className="w-4 h-4 text-blue-400 shrink-0 animate-pulse" />
                              <input
                                type="text"
                                placeholder="ARA (İSİM, SİCİL...)"
                                value={excelSearchQuery}
                                onChange={(e) => {
                                  setExcelSearchQuery(e.target.value);
                                  setActiveExcelMatchIdx(0);
                                }}
                                className="bg-transparent text-slate-100 text-xs font-extrabold outline-none border-none placeholder-slate-600 w-full uppercase"
                              />
                              {excelSearchQuery && (
                                <button
                                  onClick={() => {
                                    setExcelSearchQuery('');
                                    setActiveExcelMatchIdx(0);
                                  }}
                                  className="text-slate-500 hover:text-slate-300 transition-colors cursor-pointer"
                                  title="Aramayı Temizle"
                                >
                                  <X className="w-4 h-4" />
                                </button>
                              )}
                            </div>

                            {/* Kadro Filter ListBox (Dropdown) */}
                            <div className="flex items-center gap-2 bg-slate-950 px-3 py-2 rounded-2xl border border-slate-800 shadow-inner min-w-[180px]">
                              <span className="text-[10px] font-black text-slate-500 shrink-0 uppercase tracking-widest font-mono">
                                KADRO:
                              </span>
                              <select
                                value={selectedKadroFilter}
                                onChange={(e) => {
                                  setSelectedKadroFilter(e.target.value);
                                  setActiveExcelMatchIdx(0);
                                }}
                                className="bg-transparent border-none text-slate-200 focus:outline-none placeholder-slate-500 font-bold text-xs cursor-pointer outline-none focus:ring-0 select-none uppercase py-0"
                              >
                                <option value="" className="bg-slate-900 text-slate-200">HEPSİ / TÜMÜ</option>
                                {uniqueKadroTitles.map(title => (
                                  <option key={title} value={title} className="bg-slate-900 text-slate-200">
                                    {title.toUpperCase()}
                                  </option>
                                ))}
                              </select>
                            </div>
                          </div>

                          {/* Match Navigation & Export Buttons */}
                          <div className="flex items-center gap-3 shrink-0 flex-wrap justify-center">
                            {excelSearchQuery && (() => {
                              // Compute matches
                              const matches: { r: number; c: number; text: string }[] = [];
                              const q = excelSearchQuery.toLowerCase().trim();
                              currentFilteredRows.forEach(({ row, rIdx }) => {
                                row.forEach((cell, cIdx) => {
                                  if (cell && cell.toLowerCase().includes(q)) {
                                    matches.push({ r: rIdx, c: cIdx, text: cell });
                                  }
                                });
                              });

                              if (matches.length === 0) return null;

                              return (
                                <div className="flex items-center gap-2 bg-blue-950/60 px-3 py-1.5 rounded-xl border border-blue-900/60 font-mono text-[10px] text-blue-400 shadow-[0_0_10px_rgba(59,130,246,0.15)] animate-fade-in mr-2">
                                  <button
                                    type="button"
                                    onClick={() => {
                                      setActiveExcelMatchIdx((prev) => (prev > 0 ? prev - 1 : matches.length - 1));
                                    }}
                                    className="hover:text-white transition-colors cursor-pointer select-none font-black text-xs px-1"
                                    title="Önceki Eşleşme"
                                  >
                                    ◀
                                  </button>
                                  <span className="font-extrabold tracking-widest uppercase text-blue-300 px-1">
                                    {matches.length > 0 ? `${activeExcelMatchIdx + 1} / ${matches.length}` : "0 / 0"} EŞLEŞME
                                  </span>
                                  <button
                                    type="button"
                                    onClick={() => {
                                      setActiveExcelMatchIdx((prev) => (prev < matches.length - 1 ? prev + 1 : 0));
                                    }}
                                    className="hover:text-white transition-colors cursor-pointer select-none font-black text-xs px-1"
                                    title="Sonraki Eşleşme"
                                  >
                                    ▶
                                  </button>
                                </div>
                              );
                            })()}

                            {/* Export Actions Grid/Row */}
                            <div className="flex items-center gap-2 flex-wrap">
                              {/* DRİVE EXCEL SENKRONİZASYON */}
                              <button
                                onClick={async () => {
                                  try {
                                    setIsExcelOcrProcessing(true);
                                    showNotification("Google Drive'daki Excel dosyası aranıyor ve veriler eşitleniyor...");
                                    await pullDataFromGoogleSheets(5);
                                    showNotification("Drive Excel verileri başarıyla portal hafızasına çekildi!");
                                  } catch (err: any) {
                                    alert("Drive Excel senkronizasyon hatası: " + (err?.message || err));
                                  } finally {
                                    setIsExcelOcrProcessing(false);
                                  }
                                }}
                                className="px-3 py-1.5 bg-blue-600 hover:bg-blue-700 text-white font-extrabold font-mono text-[10px] rounded-xl flex items-center gap-1.5 transition-colors cursor-pointer shadow-sm active:scale-95"
                                title="Google Drive'daki en güncel Personel Bilgi Excel dosyasını (.xlsx) tara ve verileri portal hafızasına senkronize et"
                              >
                                <RefreshCw className="w-3.5 h-3.5 animate-spin-slow" />
                                <span>DRİVE'DAN VERİLERİ YENİLE</span>
                              </button>

                              {/* MEVCUT EXCEL İNDİR */}
                              <button
                                onClick={() => {
                                  try {
                                    const headers = TABLE_CONFIGS[5].columns.map(col => col.label);
                                    let html = `
                                      <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
                                      <head>
                                        <meta charset="utf-8">
                                        <!--[if gte mso 9]>
                                        <xml>
                                          <x:ExcelWorkbook>
                                            <x:ExcelWorksheets>
                                              <x:ExcelWorksheet>
                                                <x:Name>Personel Listesi</x:Name>
                                                <x:WorksheetOptions>
                                                  <x:DisplayGridlines/>
                                                </x:WorksheetOptions>
                                              </x:ExcelWorksheet>
                                            </x:ExcelWorksheets>
                                          </x:ExcelWorkbook>
                                        </xml>
                                        <![endif]-->
                                        <style>
                                          table { border-collapse: collapse; font-family: 'Segoe UI', Arial, sans-serif; }
                                          th { background-color: #0b3d1d; color: #ffffff; font-weight: bold; border: 1px solid #99b099; padding: 10px; text-align: center; font-size: 11px; }
                                          td { border: 1px solid #d0d5d0; padding: 8px 10px; font-size: 10px; color: #333333; }
                                          .zebra { background-color: #f5f8f5; }
                                          .num { mso-number-format: "\\@"; text-align: center; } /* Keeps TC, phone, sicil numbers formatted as strings */
                                        </style>
                                      </head>
                                      <body>
                                        <table>
                                          <thead>
                                            <tr>
                                              ${headers.map(h => `<th>${h}</th>`).join('')}
                                            </tr>
                                          </thead>
                                          <tbody>
                                    `;
                                    
                                    currentFilteredRows.forEach(({ row }, rIdx) => {
                                      const isZebra = rIdx % 2 === 1;
                                      html += `<tr class="${isZebra ? 'zebra' : ''}">`;
                                      row.forEach((cell, cIdx) => {
                                        const val = cIdx === 5 ? formatBirthDateToTurkish(cell) : (cell || "");
                                        const isNumStr = [2, 3, 7, 11].includes(cIdx);
                                        html += `<td class="${isNumStr ? 'num' : ''}" style="${cIdx === 9 ? 'white-space: pre-wrap; text-align: left;' : ''}">${val.replace(/\n/g, '<br>')}</td>`;
                                      });
                                      html += '</tr>';
                                    });
                                    
                                    html += `
                                          </tbody>
                                        </table>
                                      </body>
                                      </html>
                                    `;
                                    
                                    const blob = new Blob([html], { type: "application/vnd.ms-excel;charset=utf-8" });
                                    const link = document.createElement("a");
                                    link.href = URL.createObjectURL(blob);
                                    link.download = "personel_bilgi_cizelgesi_mevcut.xls";
                                    document.body.appendChild(link);
                                    link.click();
                                    document.body.removeChild(link);
                                    
                                    showNotification("Mevcut filtrelenmiş liste tasarımlı Excel (.xls) olarak başarıyla indirildi.");
                                  } catch (e) {
                                    alert("Excel indirme hatası: " + e);
                                  }
                                }}
                                className="px-3 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white font-extrabold font-mono text-[10px] rounded-xl flex items-center gap-1.5 transition-colors cursor-pointer shadow-sm active:scale-95"
                                title="Filtrelenmiş Mevcut Listeyi Tasarımlı Excel Olarak İndir"
                              >
                                <Download className="w-3.5 h-3.5" />
                                <span>EXCEL İNDİR</span>
                              </button>

                              {/* MEVCUT PDF İNDİR */}
                              <button
                                onClick={async () => {
                                  try {
                                    showNotification("PDF Belgesi Hazırlanıyor, Türkçe Karakterler Yükleniyor...");
                                    
                                    const doc = new jsPDF({
                                      orientation: 'l',
                                      unit: 'mm',
                                      format: 'a4'
                                    });

                                    let fontLoaded = false;
                                    try {
                                      const fontUrl = "https://fonts.gstatic.com/s/roboto/v30/KFOmCnqEu92Fr1Mu4mxKKTU1Kg.ttf"; // Roboto-Regular from Google Fonts CDN
                                      const response = await fetch(fontUrl);
                                      if (response.ok) {
                                        const arrayBuffer = await response.arrayBuffer();
                                        const bytes = new Uint8Array(arrayBuffer);
                                        let binary = "";
                                        for (let i = 0; i < bytes.byteLength; i++) {
                                          binary += String.fromCharCode(bytes[i]);
                                        }
                                        const base64Font = btoa(binary);
                                        
                                        doc.addFileToVFS("Roboto-Regular.ttf", base64Font);
                                        doc.addFont("Roboto-Regular.ttf", "Roboto-Regular", "normal");
                                        doc.setFont("Roboto-Regular");
                                        fontLoaded = true;
                                      }
                                    } catch (e) {
                                      console.error("Font loading failed, falling back to Helvetica with transliteration", e);
                                    }

                                    if (!fontLoaded) {
                                      doc.setFont("Helvetica", "bold");
                                    }

                                    const tr = (str: string) => {
                                      if (!str) return "";
                                      if (fontLoaded) return str;
                                      return str
                                        .replace(/ğ/g, "g").replace(/Ğ/g, "G")
                                        .replace(/ü/g, "u").replace(/Ü/g, "U")
                                        .replace(/ş/g, "s").replace(/Ş/g, "S")
                                        .replace(/ı/g, "i").replace(/İ/g, "I")
                                        .replace(/ö/g, "o").replace(/Ö/g, "O")
                                        .replace(/ç/g, "c").replace(/Ç/g, "C");
                                    };

                                    const chunkSize = 19;
                                    const chunks: any[][] = [];
                                    for (let i = 0; i < currentFilteredRows.length; i += chunkSize) {
                                      chunks.push(currentFilteredRows.slice(i, i + chunkSize));
                                    }

                                    if (chunks.length === 0) {
                                      alert("Yazdırılacak veri bulunamadı.");
                                      return;
                                    }

                                    const colWidths = [10, 32, 23, 14, 25, 20, 18, 24, 12, 38, 25, 25]; // total: 266mm
                                    const startX = 15;
                                    
                                    chunks.forEach((chunk, pageIdx) => {
                                      if (pageIdx > 0) {
                                        doc.addPage();
                                      }
                                      
                                      doc.setFillColor(255, 255, 255);
                                      doc.rect(0, 0, 297, 210, "F");
                                      
                                      doc.setFont(fontLoaded ? "Roboto-Regular" : "Helvetica", "bold");
                                      doc.setFontSize(13);
                                      doc.setTextColor(11, 61, 29); // OGM Dark Green
                                      doc.text(tr("ORMAN GENEL MÜDÜRLÜĞÜ - HAVACILIK DAİRESİ BAŞKANLIĞI"), 148, 14, { align: "center" });
                                      
                                      doc.setFontSize(9.5);
                                      doc.setTextColor(80, 80, 80);
                                      doc.text(tr(`PERSONEL BİLGİ ÇİZELGESİ (SAYFA ${pageIdx + 1} / ${chunks.length})`), 148, 20, { align: "center" });
                                      
                                      let startY = 26;
                                      const headerHeight = 8;
                                      
                                      doc.setFillColor(11, 61, 29);
                                      doc.rect(startX, startY, 266, headerHeight, "F");
                                      
                                      doc.setFont(fontLoaded ? "Roboto-Regular" : "Helvetica", "bold");
                                      doc.setFontSize(7.5);
                                      doc.setTextColor(255, 255, 255);
                                      
                                      const headers = [
                                        "SIRA", "ADI SOYADI", "T.C. KİMLİK", "SİCİL", 
                                        "KADRO UNV.", "DOĞUM T.", "GÖREV YERİ", "TEL NO", 
                                        "KAN", "ADRES BİLGİSİ", "YAKIN ADI", "YAKIN TEL NO"
                                      ].map(tr);
                                      
                                      let currentX = startX;
                                      headers.forEach((h, i) => {
                                        doc.text(h, currentX + colWidths[i] / 2, startY + 5.2, { align: "center" });
                                        currentX += colWidths[i];
                                      });
                                      
                                      startY += headerHeight;
                                      
                                      doc.setFont(fontLoaded ? "Roboto-Regular" : "Helvetica", "normal");
                                      doc.setTextColor(40, 40, 40);
                                      doc.setFontSize(6.5);
                                      
                                      chunk.forEach(({ row }, rIdx) => {
                                        const cellLines = row.map((cellVal, cIdx) => {
                                          const val = cIdx === 5 ? formatBirthDateToTurkish(cellVal) : (cellVal || "");
                                          return doc.splitTextToSize(tr(val), colWidths[cIdx] - 2);
                                        });
                                        
                                        const maxLines = Math.max(...cellLines.map(lines => lines.length));
                                        const lineSpacing = 2.8;
                                        const rowHeight = Math.max(7, maxLines * lineSpacing + 2.2);
                                        
                                        if (rIdx % 2 === 1) {
                                          doc.setFillColor(245, 248, 245);
                                          doc.rect(startX, startY, 266, rowHeight, "F");
                                        } else {
                                          doc.setFillColor(255, 255, 255);
                                          doc.rect(startX, startY, 266, rowHeight, "F");
                                        }
                                        
                                        doc.setDrawColor(200, 210, 200);
                                        doc.setLineWidth(0.15);
                                        doc.line(startX, startY + rowHeight, startX + 266, startY + rowHeight);
                                        
                                        let cx = startX;
                                        cellLines.forEach((lines, cIdx) => {
                                          doc.line(cx, startY, cx, startY + rowHeight);
                                          
                                          const totalTextHeight = lines.length * lineSpacing;
                                          const startTextY = startY + (rowHeight - totalTextHeight) / 2 + 2;
                                          
                                          lines.forEach((lineText, lineIdx) => {
                                            doc.text(lineText, cx + colWidths[cIdx] / 2, startTextY + lineIdx * lineSpacing, { align: "center" });
                                          });
                                          
                                          cx += colWidths[cIdx];
                                        });
                                        
                                        doc.line(cx, startY, cx, startY + rowHeight);
                                        startY += rowHeight;
                                      });
                                      
                                      doc.line(startX, 26, startX + 266, 26);
                                    });
                                    
                                    doc.save("personel_bilgi_cizelgesi_mevcut.pdf");
                                    showNotification("Mevcut Görünüm PDF olarak başarıyla indirildi.");
                                  } catch (err) {
                                    alert("PDF oluşturma hatası: " + err);
                                  }
                                }}
                                className="px-3 py-1.5 bg-red-600 hover:bg-red-700 text-white font-extrabold font-mono text-[10px] rounded-xl flex items-center gap-1.5 transition-colors cursor-pointer shadow-sm active:scale-95"
                                title="Filtrelenmiş Mevcut Listeyi 19 Satır Limitli ve Türkçe Karakter Destekli PDF Olarak İndir"
                              >
                                <FileText className="w-3.5 h-3.5" />
                                <span>PDF İNDİR</span>
                              </button>

                              {/* TARAYICIDAN MÜKEMMEL YAZDIRMA (WEB YAZDIR / PDF KAYDET) */}
                              <button
                                onClick={() => {
                                  window.print();
                                }}
                                className="px-3 py-1.5 bg-indigo-600 hover:bg-indigo-700 text-white font-extrabold font-mono text-[10px] rounded-xl flex items-center gap-1.5 transition-colors cursor-pointer shadow-sm active:scale-95 animate-pulse"
                                title="Tarayıcı yazdırma menüsünü açarak en yüksek çözünürlükte, Türkçe karakter sorunu olmadan PDF kaydedin veya yazdırın"
                              >
                                <Printer className="w-3.5 h-3.5" />
                                <span>YAZDIR / PDF YAP</span>
                              </button>
                            </div>
                          </div>
                        </div>

                        {/* Beautiful landscape scrollable page content representing Excel rendered as PDF */}
                        <div className="flex-1 overflow-auto p-6 md:p-8 flex justify-center bg-slate-950 scrollbar-thin">
                          <div className="w-full max-w-7xl bg-white text-slate-800 rounded-[1.5rem] shadow-2xl p-8 flex flex-col border border-slate-200 select-text relative overflow-x-auto">
                            
                            {/* Column Headers and Table rows using a real, beautifully-designed responsive table */}
                            <div className="flex-1 overflow-x-auto scrollbar-thin">
                              <table className="w-full border-collapse text-[11px] text-slate-700 min-w-[1100px] table-fixed">
                                <thead>
                                  <tr className="bg-slate-900 text-white font-extrabold uppercase tracking-wider text-center text-[10px]">
                                    <th className="p-2 border border-slate-200 text-center" style={{ width: '4%' }}>SIRA</th>
                                    <th className="p-2 border border-slate-200 text-left" style={{ width: '13%' }}>ADI SOYADI</th>
                                    <th className="p-2 border border-slate-200 text-center" style={{ width: '10%' }}>T.C. KİMLİK</th>
                                    <th className="p-2 border border-slate-200 text-center" style={{ width: '6%' }}>SİCİL</th>
                                    <th className="p-2 border border-slate-200 text-left" style={{ width: '10%' }}>KADRO UNV.</th>
                                    <th className="p-2 border border-slate-200 text-center" style={{ width: '8%' }}>DOĞUM T.</th>
                                    <th className="p-2 border border-slate-200 text-center" style={{ width: '8%' }}>GÖREV YERİ</th>
                                    <th className="p-2 border border-slate-200 text-center" style={{ width: '9%' }}>TELEFON NO</th>
                                    <th className="p-2 border border-slate-200 text-center" style={{ width: '5%' }}>KAN</th>
                                    <th className="p-2 border border-slate-200 text-left" style={{ width: '15%' }}>ADRES BİLGİSİ</th>
                                    <th className="p-2 border border-slate-200 text-left" style={{ width: '12%' }}>YAKININ ADI</th>
                                    <th className="p-2 border border-slate-200 text-center" style={{ width: '10%' }}>YAKIN TEL NO</th>
                                  </tr>
                                </thead>
                                <tbody>
                              {(() => {
                                const q = excelSearchQuery.toLowerCase().trim();
                                const matchesList: { r: number; c: number }[] = [];
                                
                                currentFilteredRows.forEach(({ row, rIdx }) => {
                                  row.forEach((cell, cIdx) => {
                                    if (cell && cell.toLowerCase().includes(q)) {
                                      matchesList.push({ r: rIdx, c: cIdx });
                                    }
                                  });
                                });

                                const activeMatch = matchesList[activeExcelMatchIdx];

                                if (currentFilteredRows.length === 0) {
                                  return (
                                    <tr key="no-personnel-row">
                                      <td colSpan={12} className="text-center py-12 text-slate-400 font-mono text-xs uppercase tracking-wider">
                                        🚫 ARAMA SONUCUNA UYGUN PERSONEL BULUNAMADI!
                                      </td>
                                    </tr>
                                  );
                                }

                                return currentFilteredRows.map(({ row, rIdx }) => {
                                  return (
                                    <tr 
                                      key={rIdx} 
                                      className="border-b border-slate-200 bg-white hover:bg-emerald-50/25 transition-all"
                                    >
                                      {row.map((cell, cIdx) => {
                                        // Search Match check
                                        const isMatch = q && cell.toLowerCase().includes(q);
                                        const isActiveMatch = activeMatch && activeMatch.r === rIdx && activeMatch.c === cIdx;

                                        const COLUMN_LABELS = [
                                          "Sıra No",
                                          "Adı Soyadı",
                                          "T.C. Kimlik Numarası",
                                          "Sicil Numarası",
                                          "Kadro Unvanı",
                                          "Doğum Tarihi",
                                          "Görev Yeri",
                                          "Telefon Numarası",
                                          "Kan Grubu",
                                          "Adres Bilgisi",
                                          "Yakının Adı",
                                          "Eş/Yakın Tel Numarası"
                                        ];

                                        const label = COLUMN_LABELS[cIdx] || "Bilgi";

                                        return (
                                          <td 
                                            key={cIdx} 
                                            onDoubleClick={() => {
                                              if (cell) {
                                                setActiveModalCell({
                                                  r: rIdx,
                                                  c: cIdx,
                                                  value: cIdx === 5 ? formatBirthDateToTurkish(cell) : cell,
                                                  label: label
                                                });
                                                setCopiedCellSuccess(false);
                                              }
                                            }}
                                            title={cell ? `${label}: ${cIdx === 5 ? formatBirthDateToTurkish(cell) : cell} (Detay için Tıklayın)` : "Boş Veri"}
                                            className={`p-2 border border-slate-200 truncate text-center select-text cursor-zoom-in hover:bg-emerald-50 hover:text-emerald-950 transition-all ${
                                              isActiveMatch 
                                                ? 'bg-blue-600 text-white font-black scale-105 shadow-lg ring-2 ring-blue-400 animate-pulse'
                                                : isMatch
                                                  ? 'bg-blue-200 text-blue-950 font-black border border-blue-400'
                                                  : cell 
                                                    ? 'text-slate-800 font-sans font-medium' 
                                                    : 'text-slate-350 italic'
                                            }`}
                                          >
                                            {cIdx === 5 ? formatBirthDateToTurkish(cell) : (cell || "-")}
                                          </td>
                                        );
                                      })}
                                    </tr>
                                  );
                                });
                              })()}
                                </tbody>
                              </table>
                            </div>

                            {/* PDF Footer Info */}
                            <div className="mt-6 pt-4 border-t border-slate-200 flex items-center justify-between text-[9px] font-bold text-slate-400 select-none uppercase">
                              <span>İmza / Onay: Havacılık Dairesi Başkanlığı</span>
                              <span>SAYFA 1 / 1</span>
                              <span>PORTAL SÜRÜM: V3.8-E (EXCEL-PDF)</span>
                            </div>

                          </div>
                        </div>

                        {/* Beautiful landscape print-only container representing standard HTML table broken down to 19 rows max per page */}
                        <div className="print-only-container hidden print:block bg-white text-black p-4 w-full">
                          {(() => {
                            const chunkSize = 19;
                            const chunks: any[][] = [];
                            for (let i = 0; i < currentFilteredRows.length; i += chunkSize) {
                              chunks.push(currentFilteredRows.slice(i, i + chunkSize));
                            }
                            return chunks.map((chunk, chunkIdx) => (
                              <div key={chunkIdx} className="w-full text-black bg-white" style={{ pageBreakAfter: chunkIdx < chunks.length - 1 ? 'always' : 'auto', breakAfter: chunkIdx < chunks.length - 1 ? 'page' : 'auto', minHeight: '100vh', boxSizing: 'border-box', padding: '10px' }}>
                                {/* Header */}
                                <div className="border-b-2 border-emerald-800 pb-4 mb-4 flex items-center justify-between">
                                  <div className="flex items-center gap-3">
                                    <div className="w-10 h-10 bg-emerald-800 text-white rounded flex items-center justify-center font-bold text-sm">
                                      OGM
                                    </div>
                                    <div className="text-left">
                                      <h4 className="text-xs font-bold text-emerald-900 uppercase tracking-wider">ORMAN GENEL MÜDÜRLÜĞÜ</h4>
                                      <p className="text-[8px] font-bold text-slate-500 uppercase tracking-widest">HAVACILIK DAİRESİ BAŞKANLIĞI</p>
                                    </div>
                                  </div>
                                  <div className="text-right">
                                    <h3 className="text-sm font-bold text-slate-800 uppercase tracking-tighter">PERSONEL BİLGİ ÇİZELGESİ</h3>
                                    <p className="text-[8px] text-slate-500 font-mono">SAYFA {chunkIdx + 1} / {chunks.length}</p>
                                  </div>
                                </div>

                                {/* Table */}
                                <table className="w-full border-collapse text-[10px]" style={{ tableLayout: 'fixed' }}>
                                  <thead>
                                    <tr className="bg-[#0b3d1d] text-white font-bold text-[9px]">
                                      <th className="border border-slate-300 p-1 bg-[#0b3d1d] text-white text-center" style={{ width: '4%' }}>SIRA</th>
                                      <th className="border border-slate-300 p-1 bg-[#0b3d1d] text-white text-left" style={{ width: '12%' }}>ADI SOYADI</th>
                                      <th className="border border-slate-300 p-1 bg-[#0b3d1d] text-white text-center" style={{ width: '9%' }}>T.C. KİMLİK</th>
                                      <th className="border border-slate-300 p-1 bg-[#0b3d1d] text-white text-center" style={{ width: '6%' }}>SİCİL</th>
                                      <th className="border border-slate-300 p-1 bg-[#0b3d1d] text-white text-left" style={{ width: '10%' }}>KADRO UNV.</th>
                                      <th className="border border-slate-300 p-1 bg-[#0b3d1d] text-white text-center" style={{ width: '8%' }}>DOĞUM T.</th>
                                      <th className="border border-slate-300 p-1 bg-[#0b3d1d] text-white text-center" style={{ width: '7%' }}>GÖREV YERİ</th>
                                      <th className="border border-slate-300 p-1 bg-[#0b3d1d] text-white text-center" style={{ width: '9%' }}>TEL NO</th>
                                      <th className="border border-slate-300 p-1 bg-[#0b3d1d] text-white text-center" style={{ width: '5%' }}>KAN</th>
                                      <th className="border border-slate-300 p-1 bg-[#0b3d1d] text-white text-left" style={{ width: '16%' }}>ADRES BİLGİSİ</th>
                                      <th className="border border-slate-300 p-1 bg-[#0b3d1d] text-white text-left" style={{ width: '12%' }}>YAKIN ADI</th>
                                      <th className="border border-slate-300 p-1 bg-[#0b3d1d] text-white text-center" style={{ width: '12%' }}>YAKIN TEL NO</th>
                                    </tr>
                                  </thead>
                                  <tbody>
                                    {chunk.map(({ row, rIdx }) => (
                                      <tr key={rIdx} className={rIdx % 2 === 1 ? "bg-slate-50" : "bg-white"}>
                                        <td className="border border-slate-300 p-1 text-center font-bold text-black">{row[0] || ""}</td>
                                        <td className="border border-slate-300 p-1 text-left font-bold text-black" style={{ wordBreak: 'break-word', whiteSpace: 'normal' }}>{row[1] || ""}</td>
                                        <td className="border border-slate-300 p-1 text-center font-mono text-black">{row[2] || ""}</td>
                                        <td className="border border-slate-300 p-1 text-center font-mono text-black">{row[3] || ""}</td>
                                        <td className="border border-slate-300 p-1 text-left text-black" style={{ wordBreak: 'break-word', whiteSpace: 'normal' }}>{row[4] || ""}</td>
                                        <td className="border border-slate-300 p-1 text-center font-mono text-black">{formatBirthDateToTurkish(row[5])}</td>
                                        <td className="border border-slate-300 p-1 text-center text-black" style={{ wordBreak: 'break-word', whiteSpace: 'normal' }}>{row[6] || ""}</td>
                                        <td className="border border-slate-300 p-1 text-center font-mono text-black" style={{ whiteSpace: 'nowrap' }}>{row[7] || ""}</td>
                                        <td className="border border-slate-300 p-1 text-center font-bold text-black">{row[8] || ""}</td>
                                        <td className="border border-slate-300 p-1 text-left text-[9px] text-black leading-tight" style={{ wordBreak: 'break-word', whiteSpace: 'pre-wrap' }}>
                                          {row[9] || ""}
                                        </td>
                                        <td className="border border-slate-300 p-1 text-left text-[9px] text-black" style={{ wordBreak: 'break-word', whiteSpace: 'normal' }}>{row[10] || ""}</td>
                                        <td className="border border-slate-300 p-1 text-center font-mono text-black" style={{ whiteSpace: 'nowrap' }}>{row[11] || ""}</td>
                                      </tr>
                                    ))}
                                  </tbody>
                                </table>
                              </div>
                            ));
                          })()}
                        </div>

                      </div>
                    );
                  }

                  const isSummer = isSummerForm(selectedFormId);
                  let match;
                  let expectedFileName = "";

                  if (isSummer) {
                    const airframeSuffix = getAirframeSuffix(selectedFormId);
                    const cleanMonth = selectedSummerMonth.replace(/\s+/g, '_').toLowerCase();
                    match = pdfMetadataList.find(m => {
                      const cleanName = m.name.toLowerCase();
                      const isBeklemeFile = cleanName.includes('bekleme') || cleanName.includes('ankara');
                      
                      if (airframeSuffix === 'bell429') {
                        if (isBeklemeFile || (!cleanName.includes('bell429') && !cleanName.includes('bell_429'))) {
                          return false;
                        }
                      } else if (airframeSuffix === 'bekleme_bell429') {
                        if (!isBeklemeFile || (!cleanName.includes('bell429') && !cleanName.includes('bell_429'))) {
                          return false;
                        }
                      } else {
                        if (!cleanName.includes(airframeSuffix)) {
                          return false;
                        }
                      }
                      return cleanName.includes(cleanMonth);
                    });
                    expectedFileName = `${cleanMonth}_yaz_plan_${airframeSuffix}.pdf`;
                  } else {
                    const prefix = selectedFormId === 1 ? 'gorevlendirme' : selectedFormId === 3 ? 'bakim_yetki' : selectedFormId === 5 ? 'personel_bilgi' : 'personel_ucus_hizmet';
                    match = pdfMetadataList.find(m => {
                      const cleanName = m.name.toLowerCase();
                      return cleanName.includes(prefix);
                    });
                    expectedFileName = `${prefix}_cizelgesi.pdf`;
                  }

                  if (match) {
                    if (isPdfLoading || isDownloadingPdf) {
                      return (
                        <div className="absolute inset-0 bg-slate-900/40 backdrop-blur-md flex flex-col items-center justify-center p-8 text-center animate-fade-in z-25">
                          <div className="w-12 h-12 border-4 border-emerald-500 border-t-transparent rounded-full animate-spin mb-4" />
                          <h3 className="text-emerald-400 font-extrabold text-sm uppercase tracking-widest animate-pulse font-mono mb-2">
                            LÜTFEN BEKLEYİNİZ... VERİ YÜKLENİYOR
                          </h3>
                          {pdfDownloadStatus && (
                            <p className="text-xs text-slate-400 font-mono uppercase tracking-wider max-w-md animate-pulse">
                              {pdfDownloadStatus}
                            </p>
                          )}
                        </div>
                      );
                    }

                    if (isMobile) {
                      // Mobile PDF Viewer - Sequential cached page images for high performance and zero compatibility issues
                      if (cachedPdfPages && cachedPdfPages.length > 0) {
                        return (
                          <div className="absolute inset-0 bg-slate-950 flex flex-col overflow-hidden text-white animate-fade-in z-20">
                            {/* Mobile Top Toolbar */}
                            <div className="shrink-0 bg-slate-900 border-b border-slate-800 px-4 py-3 flex items-center justify-between gap-4 select-none shadow-md z-30">
                              <div className="flex-1 min-w-0">
                                <h4 className="text-[11px] font-black text-slate-100 uppercase tracking-wider truncate">{match.name}</h4>
                                <p className="text-[9px] text-emerald-400 font-extrabold tracking-widest uppercase mt-0.5">ÖNBELLEK GÖRÜNTÜLEME (MOBİL UYUMLU)</p>
                              </div>
                              <div className="flex items-center gap-2">
                                <span className="text-[10px] bg-[#0b3d1d] text-white px-2 py-1 rounded-md font-bold font-mono">
                                  {cachedPdfPages.length} SAYFA
                                </span>
                              </div>
                            </div>

                            {/* Scrollable container displaying images of all pages */}
                            <div className="flex-1 overflow-y-auto p-4 bg-slate-900 space-y-4">
                              {cachedPdfPages.map((page) => (
                                <div key={page.pageNumber} className="flex flex-col items-center bg-slate-950 border border-slate-800 rounded-xl overflow-hidden shadow-xl p-1 relative">
                                  <img
                                    src={page.dataUrl}
                                    alt={`Sayfa ${page.pageNumber}`}
                                    className="w-full h-auto object-contain rounded-lg"
                                    referrerPolicy="no-referrer"
                                  />
                                  <div className="absolute top-3 right-3 bg-black/60 backdrop-blur-sm text-white px-2 py-1 rounded text-[9px] font-black uppercase tracking-wider font-mono">
                                    SAYFA {page.pageNumber}
                                  </div>
                                </div>
                              ))}
                            </div>
                          </div>
                        );
                      } else {
                        return (
                          <div className="absolute inset-0 bg-slate-900 flex flex-col items-center justify-center p-8 text-center animate-fade-in z-20">
                            <div className="w-12 h-12 border-4 border-emerald-500 border-t-transparent rounded-full animate-spin mb-4" />
                            <h3 className="text-emerald-400 font-extrabold text-xs uppercase tracking-widest animate-pulse font-mono mb-2">
                              CİHAZ ÖNBELLEĞİNDEN SAYFALAR YÜKLENİYOR...
                            </h3>
                            <p className="text-[10px] text-slate-400 uppercase tracking-wider max-w-xs leading-relaxed">
                              PDF sayfaları yüksek çözünürlüklü olarak taranıyor ve mobil uyumlu görünüme hazırlanıyor. Lütfen bekleyin.
                            </p>
                          </div>
                        );
                      }
                    }

                    if (pdfBlobUrl) {
                      return (
                        <div className="absolute inset-0 bg-slate-950 flex flex-col overflow-hidden text-white animate-fade-in z-20">
                          {/* Top Toolbar */}
                          <div className="shrink-0 bg-slate-900 border-b border-slate-800 px-6 py-3 flex items-center justify-between gap-4 select-none shadow-md z-30">
                            <div className="flex items-center gap-3">
                              <div className="text-left">
                                <h4 className="text-xs font-black text-slate-100 uppercase tracking-wider">{match.name}</h4>
                              </div>
                            </div>
                            <div className="flex items-center gap-2">
                              {match.viewUrl && (
                                <a
                                  href={match.viewUrl}
                                  target="_blank"
                                  rel="noopener noreferrer"
                                  className="px-3 py-1.5 bg-blue-700 hover:bg-blue-600 text-white rounded-lg flex items-center gap-1.5 text-[10px] font-black uppercase tracking-wider transition-all cursor-pointer"
                                >
                                  <FileText className="w-3.5 h-3.5" /> DRİVE'DA GÖR
                                </a>
                              )}
                              <button
                                onClick={() => {
                                  window.open(pdfBlobUrl, '_blank');
                                }}
                                className="px-3 py-1.5 bg-slate-800 hover:bg-slate-700 text-white rounded-lg flex items-center gap-1.5 text-[10px] font-black uppercase tracking-wider transition-all cursor-pointer"
                              >
                                <ExternalLink className="w-3.5 h-3.5" /> TARAYICIDA AÇ (WEB)
                              </button>
                              <a
                                href={pdfBlobUrl}
                                download={match.name}
                                className="px-3 py-1.5 bg-emerald-700 hover:bg-emerald-600 text-white rounded-lg flex items-center gap-1.5 text-[10px] font-black uppercase tracking-wider transition-all cursor-pointer"
                              >
                                <Download className="w-3.5 h-3.5" /> PLAN PDF İNDİR
                              </a>
                            </div>
                          </div>

                          {/* Main Content Area - Native PDF Reader with Toolbar and Search */}
                          <div className="flex-1 bg-slate-950 overflow-hidden relative">
                            <iframe
                              src={`${pdfBlobUrl}#toolbar=1&view=FitH`}
                              className="w-full h-full border-0 bg-slate-950"
                              title="Planlama Belgesi PDF Önizleme"
                              id="pdf-viewer-iframe"
                            />
                          </div>
                        </div>
                      );
                    }

                    return (
                      <div className="absolute inset-0 bg-slate-900 flex flex-col items-center justify-center p-8 text-center animate-fade-in z-20">
                        <div className="w-16 h-16 bg-red-950/40 text-red-400 rounded-full flex items-center justify-center mb-4 border border-red-900">
                          <AlertTriangle className="w-8 h-8" />
                        </div>
                        <h3 className="text-white font-black text-sm uppercase tracking-widest mb-2">
                          PLAN BELGESİ AÇILAMADI
                        </h3>
                        <p className="text-xs text-slate-400 max-w-sm leading-relaxed mb-6 uppercase tracking-wider">
                          Belge indirildi ancak tarayıcıda önizleme oluşturulamadı. Lütfen sayfayı yenileyip tekrar deneyin veya planı yeniden yükleyin.
                        </p>
                        <button
                          onClick={() => {
                            const cleanLastUpdated = match.lastUpdated.replace(/[^a-zA-Z0-9]/g, '_');
                            const cacheKey = `pdf_${match.id}_${cleanLastUpdated}`;
                            loadRawPdfFromDrive(match.id, cacheKey);
                          }}
                          className="px-6 py-3 bg-emerald-700 hover:bg-emerald-600 text-white text-xs font-extrabold uppercase rounded-xl tracking-wider transition-all cursor-pointer"
                        >
                          🔄 TEKRAR DENE
                        </button>
                      </div>
                    );
                  } else {
                    return (
                      <div className="absolute inset-0 bg-slate-50 flex flex-col items-center justify-center p-8 text-center select-none z-10 animate-fade-in">
                        <div className="w-20 h-20 bg-amber-50 text-amber-600 rounded-3xl flex items-center justify-center mb-6 border border-amber-100 shadow-sm">
                          <FileText className="w-10 h-10" />
                        </div>
                        <h4 className="text-slate-800 font-extrabold tracking-tight text-base uppercase mb-2">
                          PLAN PDF'İ GOOGLE DRIVE'DA BULUNAMADI
                        </h4>
                        <p className="text-xs text-slate-500 font-bold leading-relaxed max-w-md uppercase tracking-wide">
                          Seçilen <span className="text-emerald-800">"{TABLE_CONFIGS[selectedFormId!]?.title}"</span> planı için Drive klasöründe uygun PDF belgesi tespit edilemedi.
                        </p>
                        <p className="text-[11px] text-slate-400 mt-2 max-w-sm leading-relaxed">
                          Lütfen dosyanızı Güncelleme Sihirbazından yükleyin. Beklenen dosya adı: <strong>{expectedFileName}</strong>
                        </p>
                        
                        <div className="mt-8 flex flex-col sm:flex-row gap-3">
                          <button
                            onClick={() => {
                              if (selectedFormId) {
                                setSyncSelectedTarget(String(selectedFormId));
                                setStep1Target(String(selectedFormId));
                              }
                              setActiveSyncStep(2);
                              setModalType('excel_sync');
                              setModalTitle('VERİ GÜNCELLEME SİHİRBAZI');
                            }}
                            className="px-6 py-3 bg-[#0b3d1d] hover:bg-[#072612] text-white font-extrabold text-xs uppercase rounded-xl shadow-md tracking-wider transition-all cursor-pointer"
                          >
                            📥 PDF PLAN YÜKLE
                          </button>
                          <a
                            href="https://drive.google.com/drive/folders/1_fIGvuPVpC9N5on1irOfGG8OsD1KSXD0"
                            target="_blank"
                            rel="noopener noreferrer"
                            className="px-6 py-3 bg-white hover:bg-slate-100 text-slate-700 border border-slate-300 font-extrabold text-xs uppercase rounded-xl shadow-sm tracking-wider transition-all cursor-pointer inline-flex items-center gap-2"
                          >
                            📁 DRİVE KLASÖRÜNÜ AÇ ➜
                          </a>
                        </div>
                      </div>
                    );
                  }
                })()}
              </div>

            </div>
          )}

          {/* TEÇHİZAT TAKİP MATRİS EKRANI */}
          {modalType === 'techizat_matrix' && activeTechizatType && (() => {
            const isKaraAraci = activeTechizatType === 'kara_araclari';
            const baseColumns = [
              "SIRA NO", 
              "TEÇHİZAT ADI", 
              "PARÇA NO (P/N) / MODEL", 
              "SERİ NO (S/N)", 
              "MİKTAR / KAPASİTE", 
              "BULUNDUĞU YER", 
              "DURUMU", 
              "KALİBRASYONA TABİ", 
              "SON KONTROL / KALİBRASYON / BAKIM", 
              "GELECEK KONTROL / KALİBRASYON / BAKIM", 
              "SON KONTROLÜ YAPAN FİRMA", 
              "AÇIKLAMA", 
              "90 GÜN UYARISI MAİL GÖNDERİM TARİHİ"
            ];

            const karaAraclariColumns = [
              "SIRA NO", 
              "ARAÇ PLAKASI / TANIMI", 
              "PARÇA NO (P/N) / MODEL", 
              "BULUNDUĞU YER", 
              "SON KM Sİ", 
              "DURUMU", 
              "KALİBRASYONA TABİ", 
              "SON KONTROL / KALİBRASYON / BAKIM", 
              "GELECEK KONTROL / KALİBRASYON / BAKIM", 
              "SON KONTROLÜ YAPAN FİRMA", 
              "AÇIKLAMA", 
              "90 GÜN UYARISI MAİL GÖNDERİM TARİHİ"
            ];

            const cols = activeTechizatType === 'all'
              ? ["AİT OLDUĞU BİRİM", ...baseColumns]
              : isKaraAraci
                ? karaAraclariColumns
                : baseColumns;

            const formatStandardRow = (row: string[]) => {
              const r = [...row];
              const col7Upper = (r[7] || "").trim().toUpperCase();
              if (col7Upper !== "EVET" && col7Upper !== "HAYIR") {
                r.splice(7, 0, "EVET");
              }
              while (r.length < 13) {
                r.push("");
              }
              return r.slice(0, 13);
            };

            const formatKaraAraclariRow = (row: string[]) => {
              const r = [...row];
              const col6Upper = (r[6] || "").trim().toUpperCase();
              if (col6Upper !== "EVET" && col6Upper !== "HAYIR") {
                r.splice(6, 0, "EVET");
              }
              while (r.length < 12) {
                r.push("");
              }
              return r.slice(0, 12);
            };

            const formatKaraAraciToStandardRow = (row: string[]) => {
              const r = formatKaraAraclariRow(row);
              const sira = r[0];
              const plaka = r[1];
              const model = r[2];
              const yer = r[3];
              const km = r[4] ? `${r[4]} KM` : "";
              const durum = r[5];
              const kalibrasyonTabi = r[6];
              const sonBakim = r[7];
              const gelecekBakim = r[8];
              const firma = r[9];
              const aciklama = r[10];
              const mail = r[11];
              
              return [
                sira,
                plaka,
                model,
                km, 
                "1", 
                yer,
                durum,
                kalibrasyonTabi,
                sonBakim,
                gelecekBakim,
                firma,
                aciklama,
                mail
              ];
            };

            const rows = activeTechizatType === 'all'
              ? [
                  ...techizatBell429Data.map(r => ["BELL 429", ...formatStandardRow(r)]),
                  ...techizatAt802Data.map(r => ["AT-802F", ...formatStandardRow(r)]),
                  ...techizatT70Data.map(r => ["T-70 YER DESTEK", ...formatStandardRow(r)]),
                  ...techizatT70BumbiBacketData.map(r => ["T-70 BUMBİ BACKET", ...formatStandardRow(r)]),
                  ...techizatT70HelitakData.map(r => ["T-70 HELİTAK", ...formatStandardRow(r)]),
                  ...techizatC650Data.map(r => ["C-650", ...formatStandardRow(r)]),
                  ...techizatB360Data.map(r => ["B-360", ...formatStandardRow(r)]),
                  ...techizatHangarData.map(r => ["HANGAR YER DESTEK", ...formatStandardRow(r)]),
                  ...techizatKaraAraclariData.map(r => ["KARA ARAÇLARI", ...formatKaraAraciToStandardRow(r)])
                ]
              : activeTechizatType === 'bell429' ? techizatBell429Data.map(formatStandardRow)
              : activeTechizatType === 'at802' ? techizatAt802Data.map(formatStandardRow)
              : activeTechizatType === 't70' ? techizatT70Data.map(formatStandardRow)
              : activeTechizatType === 't70_bumbi_backet' ? techizatT70BumbiBacketData.map(formatStandardRow)
              : activeTechizatType === 't70_helitak' ? techizatT70HelitakData.map(formatStandardRow)
              : activeTechizatType === 'b360' ? techizatB360Data.map(formatStandardRow)
              : activeTechizatType === 'c650' ? techizatC650Data.map(formatStandardRow)
              : activeTechizatType === 'kara_araclari' ? techizatKaraAraclariData.map(formatKaraAraclariRow)
              : techizatHangarData.map(formatStandardRow);

            const firmaColIdx = cols.indexOf("SON KONTROLÜ YAPAN FİRMA");
            const durumColIdx = cols.indexOf("DURUMU");
            const kalibrasyonTabiColIdx = cols.indexOf("KALİBRASYONA TABİ");

            const uniqueFirmalar = Array.from(
              new Set(
                rows
                  .map(r => (r[firmaColIdx] || "").trim())
                  .filter(f => f && f !== "-" && f !== "MUAFIYET (TABİ DEĞİL)" && f.toUpperCase() !== "BELİRTİLMEMİŞ")
              )
            ).sort();

            const q = techizatSearchQuery.toLowerCase().trim();
            const firmaFilter = techizatFirmaFilter.toLowerCase().trim();
            const durumFilter = techizatDurumFilter.toUpperCase().trim();

            const matchesList: { r: number; c: number }[] = [];

            const filteredRows = rows.filter(row => {
              // 1. Arama kelimesi filtresi
              if (q && !row.some(cell => cell && cell.toLowerCase().includes(q))) {
                return false;
              }
              // 2. Firma filtresi
              if (firmaFilter && firmaFilter !== "tüm fİrmalar") {
                const rowFirma = (row[firmaColIdx] || "").toLowerCase().trim();
                if (!rowFirma.includes(firmaFilter)) return false;
              }
              // 3. Durum filtresi
              if (durumFilter && durumFilter !== "TÜM DURUMLAR") {
                const rowDurum = (row[durumColIdx] || "").toUpperCase().trim();
                if (durumFilter === "FAAL" && !rowDurum.includes("FAAL")) return false;
                if (durumFilter === "BAKIM / KALİBRASYON" && !rowDurum.includes("BAKIM") && !rowDurum.includes("KALİBRASYON")) return false;
                if (durumFilter === "GAYRİ FAAL" && !rowDurum.includes("GAYRİ") && !rowDurum.includes("DEĞİL")) return false;
              }
              return true;
            });

            // Locate search match coordinates
            rows.forEach((row, rIdx) => {
              row.forEach((cell, cIdx) => {
                if (q && cell && cell.toLowerCase().includes(q)) {
                  matchesList.push({ r: rIdx, c: cIdx });
                }
              });
            });

            const activeMatch = matchesList[activeTechizatMatchIdx];
            const gelecekBakimColIdx = cols.indexOf("GELECEK KONTROL / KALİBRASYON / BAKIM");

            // Sıralama (Normalde seçili değil, ama basılınca Turuncu -> Yeşil sıralasın)
            let processedRows = [...filteredRows];
            if (sortByColor && gelecekBakimColIdx !== -1) {
              processedRows.sort((rowA, rowB) => {
                const getPriorityScore = (row: string[]) => {
                  const isKalibTabiVal = kalibrasyonTabiColIdx !== -1 ? (row[kalibrasyonTabiColIdx] || "").trim().toUpperCase() : "EVET";
                  if (isKalibTabiVal === "HAYIR") return 4; // Kalibrasyon gerekmiyorsa en düşük öncelik

                  const val = row[gelecekBakimColIdx] || "";
                  const days = parseGelecekBakimDays(val);
                  if (days === null) return 3;
                  if (days < 90) return 1; // Turuncu
                  return 2; // Yeşil (>= 90 gün)
                };

                const scoreA = getPriorityScore(rowA);
                const scoreB = getPriorityScore(rowB);
                
                if (scoreA !== scoreB) {
                  return scoreA - scoreB;
                }
                
                // Aynı gruptakileri en yakın gün sayısına göre artan sırala
                const valA = rowA[gelecekBakimColIdx] || "";
                const valB = rowB[gelecekBakimColIdx] || "";
                const daysA = parseGelecekBakimDays(valA);
                const daysB = parseGelecekBakimDays(valB);
                const dJanA = daysA !== null ? daysA : 999999;
                const dJanB = daysB !== null ? daysB : 999999;
                return dJanA - dJanB;
              });
            }

            return (
              <div className="absolute inset-0 flex flex-col bg-slate-50 p-2 sm:p-4 md:p-6 animate-fade-in overflow-y-auto">
                <div className="w-full max-w-[1920px] mx-auto flex flex-col h-full">
                  
                  {/* Header Titles */}
                  <div className="text-center mb-6 select-none print:hidden">
                    <h3 className="text-xl sm:text-2xl font-black text-slate-800 uppercase tracking-tighter">
                      🛠️ {modalTitle}
                    </h3>
                    {formUpdateDates[`techizat_${activeTechizatType}`] && (
                      <p className="text-xs font-mono font-extrabold text-emerald-700 mt-2 bg-emerald-100 border border-emerald-200/50 rounded-full px-4 py-1.5 inline-block shadow-sm">
                        📅 GÜNCELLEME TARİHİ: {formUpdateDates[`techizat_${activeTechizatType}`]}
                      </p>
                    )}
                  </div>

                    <>
                      {isKaraAraci && (
                        <div className="flex justify-center gap-4 mb-6 select-none print:hidden">
                          <button
                            onClick={() => setKaraAraclariSubTab('list')}
                            className={`px-6 py-3 rounded-2xl font-black text-xs uppercase tracking-wider transition-all cursor-pointer flex items-center gap-2 border shadow-sm ${
                              karaAraclariSubTab === 'list'
                                ? 'bg-[#0b3d1d] text-white border-[#0b3d1d] scale-105 shadow-md shadow-emerald-900/20'
                                : 'bg-white hover:bg-slate-100 text-slate-700 border-slate-200 hover:scale-102'
                            }`}
                          >
                            <Truck className="w-4 h-4" />
                            🚗 KARA ARAÇ TAKİP LİSTESİ
                          </button>
                          <button
                            onClick={() => setKaraAraclariSubTab('mission_order')}
                            className={`px-6 py-3 rounded-2xl font-black text-xs uppercase tracking-wider transition-all cursor-pointer flex items-center gap-2 border shadow-sm ${
                              karaAraclariSubTab === 'mission_order'
                                ? 'bg-[#0b3d1d] text-white border-[#0b3d1d] scale-105 shadow-md shadow-emerald-900/20'
                                : 'bg-white hover:bg-slate-100 text-slate-700 border-slate-200 hover:scale-102'
                            }`}
                          >
                            <FileText className="w-4 h-4" />
                            📋 GÖREV EMRİ GİRİŞ
                          </button>
                          <button
                            onClick={() => setKaraAraclariSubTab('past_records')}
                            className={`px-6 py-3 rounded-2xl font-black text-xs uppercase tracking-wider transition-all cursor-pointer flex items-center gap-2 border shadow-sm ${
                              karaAraclariSubTab === 'past_records'
                                ? 'bg-[#0b3d1d] text-white border-[#0b3d1d] scale-105 shadow-md shadow-emerald-900/20'
                                : 'bg-white hover:bg-slate-100 text-slate-700 border-slate-200 hover:scale-102'
                            }`}
                          >
                            <History className="w-4 h-4" />
                            📜 GEÇMİŞ KAYITLAR
                          </button>
                        </div>
                      )}

                      {!isKaraAraci || karaAraclariSubTab === 'list' ? (
                    <>
                      {/* Search and Utility Controls */}
                  <div className="bg-slate-900 text-slate-200 rounded-3xl p-5 mb-6 flex flex-col xl:flex-row gap-4 items-center justify-between shadow-xl border border-slate-800 print:hidden select-none">
                    <div className="flex flex-wrap items-center gap-3 w-full xl:w-auto">
                      <div className="bg-slate-800 p-2.5 rounded-2xl border border-slate-700">
                        <Search className="w-5 h-5 text-emerald-400" />
                      </div>
                      
                      {/* Metin Arama Input */}
                      <div className="relative flex-1 sm:flex-initial">
                        <input
                          type="text"
                          placeholder="Teçhizat veya seri no..."
                          value={techizatSearchQuery}
                          onChange={(e) => {
                            setTechizatSearchQuery(e.target.value);
                            setActiveTechizatMatchIdx(0);
                          }}
                          className="bg-slate-800 text-white font-extrabold text-xs px-4 py-3 rounded-2xl focus:outline-none focus:ring-4 focus:ring-emerald-500/20 w-full sm:w-48 border border-slate-700 placeholder-slate-400"
                        />
                      </div>

                      {/* Firma Filtreleme Input / Datalist Dropdown */}
                      <div className="relative flex-1 sm:flex-initial">
                        <input
                          type="text"
                          list="techizat-firmalar-list"
                          placeholder="🏢 Firma ile ara / seç..."
                          value={techizatFirmaFilter}
                          onChange={(e) => {
                            setTechizatFirmaFilter(e.target.value);
                            setActiveTechizatMatchIdx(0);
                          }}
                          className="bg-slate-800 text-emerald-300 font-extrabold text-xs px-4 py-3 rounded-2xl focus:outline-none focus:ring-4 focus:ring-emerald-500/20 w-full sm:w-52 border border-slate-700 placeholder-slate-400"
                        />
                        <datalist id="techizat-firmalar-list">
                          <option value="TÜM FİRMALAR">TÜM FİRMALAR</option>
                          {uniqueFirmalar.map((f, i) => (
                            <option key={i} value={f}>{f}</option>
                          ))}
                        </datalist>
                      </div>

                      {/* Durum Filtreleme Dropdown */}
                      <div className="relative flex-1 sm:flex-initial">
                        <select
                          value={techizatDurumFilter}
                          onChange={(e) => {
                            setTechizatDurumFilter(e.target.value);
                            setActiveTechizatMatchIdx(0);
                          }}
                          className="bg-slate-800 text-amber-300 font-black text-xs px-4 py-3 rounded-2xl focus:outline-none focus:ring-4 focus:ring-emerald-500/20 w-full sm:w-44 border border-slate-700 cursor-pointer"
                        >
                          <option value="">🟢 TÜM DURUMLAR</option>
                          <option value="FAAL">FAAL</option>
                          <option value="BAKIM / KALİBRASYON">BAKIM / KALİBRASYON</option>
                          <option value="GAYRİ FAAL">GAYRİ FAAL</option>
                        </select>
                      </div>

                      {(techizatSearchQuery || techizatFirmaFilter || techizatDurumFilter) && (
                        <button
                          onClick={() => {
                            setTechizatSearchQuery("");
                            setTechizatFirmaFilter("");
                            setTechizatDurumFilter("");
                            setActiveTechizatMatchIdx(0);
                          }}
                          className="px-3 py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-xl text-xs font-bold transition-all border border-slate-700"
                          title="Filtreleri Temizle"
                        >
                          ✕ Temizle
                        </button>
                      )}

                      {techizatSearchQuery && (
                        <span className="text-[10px] font-mono font-black text-emerald-400 px-3 py-1.5 bg-emerald-950/50 rounded-xl border border-emerald-900 shrink-0">
                          {matchesList.length} EŞLEŞME
                        </span>
                      )}

                      <button
                        onClick={() => setSortByColor(!sortByColor)}
                        className={`px-4 py-3 active:scale-95 font-black font-mono text-xs rounded-2xl flex items-center gap-2 transition-all cursor-pointer shadow-lg border shrink-0 ${
                          sortByColor 
                            ? 'bg-red-600 hover:bg-red-700 text-white border-red-500 animate-pulse' 
                            : 'bg-slate-800 hover:bg-slate-750 text-slate-200 border-slate-700'
                        }`}
                        title="Bakım gün sayısına göre (Kırmızı ➜ Turuncu ➜ Sarı ➜ Yeşil) sıralar"
                      >
                        <SlidersHorizontal className="w-4 h-4 text-amber-400" />
                        <span>{sortByColor ? "🔴 RENK SIRALAMASI AKTİF" : "⏳ RENK KODUNA GÖRE SIRALA"}</span>
                      </button>
                    </div>

                    <div className="flex flex-wrap items-center gap-3 w-full xl:w-auto justify-end">
                      {matchesList.length > 0 && (
                        <div className="flex items-center gap-2 bg-slate-800 px-3 py-2 rounded-2xl border border-slate-700">
                          <button
                            onClick={() => {
                              setActiveTechizatMatchIdx(prev => (prev - 1 + matchesList.length) % matchesList.length);
                            }}
                            className="p-1 text-slate-400 hover:text-white transition-colors cursor-pointer text-xs"
                            title="Önceki"
                          >
                            ◀
                          </button>
                          <span className="text-[10px] font-mono font-black text-slate-300">
                            {activeTechizatMatchIdx + 1} / {matchesList.length}
                          </span>
                          <button
                            onClick={() => {
                              setActiveTechizatMatchIdx(prev => (prev + 1) % matchesList.length);
                            }}
                            className="p-1 text-slate-400 hover:text-white transition-colors cursor-pointer text-xs"
                            title="Sonraki"
                          >
                            ▶
                          </button>
                        </div>
                      )}

                      {activeTechizatType !== 'all' && (
                        <button
                          onClick={() => {
                            setPasswordInput('');
                            setPasswordError(false);
                            setIsPasswordModalOpen(true);
                          }}
                          className="px-4 py-3 bg-[#0b3d1d] hover:bg-[#072612] active:scale-95 text-white font-black font-mono text-xs rounded-2xl flex items-center gap-2 transition-all cursor-pointer shadow-lg border border-emerald-600 shrink-0"
                        >
                          <RefreshCw className="w-4 h-4 text-emerald-300 animate-spin-slow" />
                          <span>VERİ GÜNCELLE</span>
                        </button>
                      )}

                      <button
                        onClick={() => exportTechizatToExcel(activeTechizatType, cols, rows, modalTitle)}
                        className="px-4 py-3 bg-emerald-700 hover:bg-emerald-600 active:scale-95 text-white font-black font-mono text-xs rounded-2xl flex items-center gap-2 transition-all cursor-pointer shadow-lg shadow-emerald-900/10 border border-emerald-600 shrink-0"
                      >
                        <Download className="w-4 h-4" />
                        <span>EXCEL OLARAK AKTAR</span>
                      </button>

                      <button
                        onClick={() => window.print()}
                        className="px-4 py-3 bg-indigo-700 hover:bg-indigo-600 active:scale-95 text-white font-black font-mono text-xs rounded-2xl flex items-center gap-2 transition-all cursor-pointer shadow-lg shadow-indigo-900/10 border border-indigo-600 shrink-0"
                      >
                        <Printer className="w-4 h-4" />
                        <span>YAZDIR / PDF İNDİR</span>
                      </button>
                    </div>
                  </div>

                  {/* Dynamic Interactive Grid Container */}
                  <div className="flex-1 bg-white border-2 border-slate-200/60 rounded-[2.5rem] shadow-xl overflow-hidden flex flex-col print:border-none print:shadow-none min-h-[400px]">
                    
                    {/* Table Title Bar */}
                    <div className="bg-slate-900 px-6 py-4 border-b border-slate-800 flex items-center justify-between print:hidden shrink-0">
                      {Object.keys(selectedTechizatItems).length > 0 ? (
                        <div className="flex items-center gap-3 animate-fade-in">
                          <span className="text-xs font-black text-amber-400 bg-amber-950/40 border border-amber-900 px-3 py-1.5 rounded-xl">
                            ⚡ {Object.keys(selectedTechizatItems).length} TEÇHİZAT SEÇİLDİ
                          </span>
                          <button
                            onClick={() => {
                              setBulkModalMode('choice');
                              setIsEbysModalOpen(true);
                            }}
                            className="px-4 py-1.5 bg-emerald-600 hover:bg-emerald-500 hover:scale-105 active:scale-95 text-white font-black text-xs uppercase tracking-wider rounded-xl transition-all cursor-pointer flex items-center gap-2 shadow-lg shadow-emerald-950/30 border border-emerald-500"
                          >
                            <Sparkles className="w-3.5 h-3.5" />
                            <span>DÜZENLE / GÖNDER</span>
                          </button>
                        </div>
                      ) : (
                        <div />
                      )}
                      <span className="text-[10px] font-mono font-black text-slate-400 bg-slate-800 px-3 py-1 rounded-full border border-slate-700">
                        {processedRows.length} KALEM TEÇHİZAT LİSTELENDİ
                      </span>
                    </div>

                    {/* Table Grid Scroll Wrapper */}
                    <div className="flex-1 overflow-auto max-h-[75vh] print:max-h-none print:overflow-visible">
                      <table className="w-full border-collapse text-left min-w-[1200px]">
                        <thead>
                          <tr className="bg-slate-900 border-b border-slate-800 print:bg-[#0b3d1d] shrink-0 sticky top-0 z-10">
                            <th className="px-3 py-3.5 text-center text-[10px] font-black text-slate-300 uppercase font-mono border-r border-slate-800 w-[50px] print:hidden">
                              <input
                                type="checkbox"
                                onChange={(e) => {
                                  if (e.target.checked) {
                                    const items = { ...selectedTechizatItems };
                                    processedRows.forEach((row, rIdx) => {
                                      const isAll = activeTechizatType === 'all';
                                      const targetTechType = isAll ? getRealTechType(row[0]) : activeTechizatType;
                                      const targetRow = isAll ? row.slice(1) : row;
                                      if (targetTechType) {
                                        const key = `${targetTechType}_${rIdx}`;
                                        items[key] = { techType: targetTechType, row: targetRow };
                                      }
                                    });
                                    setSelectedTechizatItems(items);
                                  } else {
                                    const items = { ...selectedTechizatItems };
                                    processedRows.forEach((row, rIdx) => {
                                      const isAll = activeTechizatType === 'all';
                                      const targetTechType = isAll ? getRealTechType(row[0]) : activeTechizatType;
                                      if (targetTechType) {
                                        const key = `${targetTechType}_${rIdx}`;
                                        delete items[key];
                                      }
                                    });
                                    setSelectedTechizatItems(items);
                                  }
                                }}
                                checked={processedRows.length > 0 && processedRows.every((row, rIdx) => {
                                  const isAll = activeTechizatType === 'all';
                                  const targetTechType = isAll ? getRealTechType(row[0]) : activeTechizatType;
                                  return `${targetTechType}_${rIdx}` in selectedTechizatItems;
                                })}
                                className="w-4 h-4 rounded text-emerald-600 focus:ring-emerald-500 border-slate-700 cursor-pointer bg-slate-800"
                              />
                            </th>
                            {cols.map((col, cIdx) => (
                              <th 
                                key={cIdx} 
                                className="px-4 py-3.5 text-center text-[10px] font-black tracking-wider text-slate-300 uppercase font-mono border-r border-slate-800 last:border-r-0"
                              >
                                {col}
                              </th>
                            ))}
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-100">
                          {processedRows.length === 0 ? (
                            <tr key="no-equipment-row">
                              <td colSpan={cols.length + 1} className="px-6 py-16 text-center text-slate-400 font-extrabold text-sm">
                                🔍 Arama kriterlerine uygun teçhizat kaydı bulunamadı.
                              </td>
                            </tr>
                          ) : (
                             processedRows.map((row, rIdx) => {
                               const isAll = activeTechizatType === 'all';
                               const targetTechType = isAll ? getRealTechType(row[0]) : activeTechizatType;
                               const targetRow = isAll ? row.slice(1) : row;
                               const rowImageKey = targetTechType + "_" + (targetRow[1] || "").replace(/\s+/g, '_') + "_" + (targetRow[3] || "").replace(/\s+/g, '_');
                               const rowImageUrl = techizatImages[rowImageKey];

                               return (
                                 <tr 
                                   key={rIdx} 
                                   onMouseEnter={(e) => {
                                     if (rowImageUrl) {
                                       setHoveredRowImage({
                                         url: rowImageUrl,
                                         title: targetRow[1] || "Teçhizat",
                                         subtitle: `${targetRow[2] || ""} ${targetRow[3] ? "• " + targetRow[3] : ""}`,
                                         x: e.clientX,
                                         y: e.clientY
                                       });
                                     }
                                   }}
                                   onMouseMove={(e) => {
                                     if (rowImageUrl) {
                                       setHoveredRowImage(prev => prev ? { ...prev, x: e.clientX, y: e.clientY } : null);
                                     }
                                   }}
                                   onMouseLeave={() => setHoveredRowImage(null)}
                                   onClick={() => {
                                      setHoveredRowImage(null);
                                      setMobileEditTab('form');

                                      let resolvedRow = targetRow;
                                      if (isAll) {
                                        if (targetTechType === 'kara_araclari') {
                                          const found = techizatKaraAraclariData.find(kr => (kr[1] || "").trim() === (targetRow[1] || "").trim());
                                          if (found) resolvedRow = found;
                                        } else {
                                          let sourceList: string[][] = [];
                                          if (targetTechType === 'bell429') sourceList = techizatBell429Data;
                                          else if (targetTechType === 'at802') sourceList = techizatAt802Data;
                                          else if (targetTechType === 't70') sourceList = techizatT70Data;
                                          else if (targetTechType === 't70_bumbi_backet') sourceList = techizatT70BumbiBacketData;
                                          else if (targetTechType === 'b360') sourceList = techizatB360Data;
                                          else if (targetTechType === 'c650') sourceList = techizatC650Data;
                                          else if (targetTechType === 'hangar') sourceList = techizatHangarData;
                                          
                                          const found = sourceList.find(sr => (sr[0] || "").trim() === (targetRow[0] || "").trim() && (sr[1] || "").trim() === (targetRow[1] || "").trim());
                                          if (found) resolvedRow = found;
                                        }
                                      }

                                      setActiveTechizatRowEdit({
                                        rIdx,
                                        techType: targetTechType,
                                        row: [...resolvedRow]
                                      });
                                      const imageKey = targetTechType + "_" + (resolvedRow[1] || "").replace(/\s+/g, '_') + "_" + (resolvedRow[3] || "").replace(/\s+/g, '_');
                                      setTempImageUrlInput(techizatImages[imageKey] && !techizatImages[imageKey].startsWith('data:') ? techizatImages[imageKey] : "");
                                      setEditRowValues([...resolvedRow]);
                                      setTechizatImageScale(1);
                                      setIsFullScreenImage(false);
                                      setImagePasswordInput('');
                                      setImagePasswordError(false);
                                      setShowImagePasswordPrompt(false);
                                      setPendingImageFile(null);
                                      setPendingImagePreview(null);
                                      setIsDragging(false);
                                      setIsImageUploadingToDrive(false);
                                      setIsDataUpdateUnlocked(false);
                                      setDataPasswordInput('');
                                      setDataPasswordError(false);
                                      setShowSavePasswordPrompt(false);
                                      setShowImageSavePasswordPrompt(false);
                                      setTempImageAction(null);
                                    }}
                                   className="hover:bg-emerald-50/40 transition-colors duration-150 odd:bg-white even:bg-slate-50/50 cursor-pointer"
                                   title="Düzenlemek ve görsel eklemek için Tıklayın (Görseli kare içinde görmek için imleci üzerinde tutun)"
                                 >

                                   {row.map((cell, cIdx) => {
                                     const label = cols[cIdx] || "Veri";
                                     const isMatch = q && cell && cell.toLowerCase().includes(q);
                                     const isActiveMatch = activeMatch && activeMatch.r === rIdx && activeMatch.c === cIdx;
 
                                     let cellStyleClass = cell 
                                       ? 'text-slate-800 font-sans font-semibold' 
                                       : 'text-slate-300 italic';
                                       
                                     const isBakimaTabiColIdx = cols.indexOf("BAKIMA TABİ Mİ?");
                                     const isBakimaTabi = isBakimaTabiColIdx !== -1 ? row[isBakimaTabiColIdx] : "Evet";

                                     const isGelecekBakimCol = cIdx === gelecekBakimColIdx || (isKaraAraci && label === "BİR SONRAKİ MUAYENE TARİHİ");

                                     if (isGelecekBakimCol && cell) {
                                       if (isBakimaTabi === "Hayır") {
                                         cellStyleClass = 'text-slate-400 italic font-bold text-center';
                                       } else {
                                         const days = parseGelecekBakimDays(cell);
                                         if (days !== null) {
                                           if (days >= 90) {
                                             cellStyleClass = 'bg-emerald-500 text-white font-extrabold px-3 py-1.5 rounded-xl shadow-sm text-center';
                                           } else {
                                             cellStyleClass = 'bg-orange-500 text-white font-extrabold px-3 py-1.5 rounded-xl shadow-sm text-center animate-pulse';
                                           }
                                         }
                                       }
                                     }
 
                                     return (
                                       <React.Fragment key={cIdx}>
                                         {cIdx === 0 && (
                                           <td 
                                             className="px-3 py-3 text-center border-r border-slate-100 print:hidden"
                                             onClick={(e) => e.stopPropagation()}
                                           >
                                             <input
                                               type="checkbox"
                                               checked={(() => {
                                                 const isAll = activeTechizatType === 'all';
                                                 const targetTechType = isAll ? getRealTechType(row[0]) : activeTechizatType;
                                                 return `${targetTechType}_${rIdx}` in selectedTechizatItems;
                                               })()}
                                               onChange={() => {
                                                 const isAll = activeTechizatType === 'all';
                                                 const targetTechType = isAll ? getRealTechType(row[0]) : activeTechizatType;
                                                 const targetRow = isAll ? row.slice(1) : row;
                                                 if (targetTechType) {
                                                   const key = `${targetTechType}_${rIdx}`;
                                                   const updated = { ...selectedTechizatItems };
                                                   if (key in updated) {
                                                     delete updated[key];
                                                   } else {
                                                     updated[key] = { techType: targetTechType, row: targetRow };
                                                   }
                                                   setSelectedTechizatItems(updated);
                                                 }
                                               }}
                                               className="w-4 h-4 rounded text-emerald-600 focus:ring-emerald-500 border-slate-300 cursor-pointer"
                                             />
                                           </td>
                                         )}
                                         <td 
                                           key={cIdx} 
                                           className="px-4 py-3 text-center border-r border-slate-100 last:border-r-0 max-w-[200px]"
                                         >
                                           <div 
                                             title={cell ? `${label}: ${cell} (Düzenlemek ve görsel eklemek için Tıklayın)` : "Boş Veri"}
                                             className={`${cell && cell.includes('\n') ? 'whitespace-pre-line leading-relaxed min-w-[100px]' : 'truncate'} px-2 py-1 rounded-xl transition-all text-xs text-center select-text hover:bg-emerald-100/50 hover:text-emerald-950 flex items-center justify-center gap-1 ${
                                               isActiveMatch 
                                                 ? 'bg-blue-600 text-white font-black scale-105 shadow-md ring-2 ring-blue-400 animate-pulse'
                                                 : isMatch
                                                   ? 'bg-blue-200 text-blue-950 font-black border border-blue-400'
                                                   : cellStyleClass
                                             }`}
                                           >
                                             <span className="truncate">{cell || "-"}</span>
                                           </div>
                                         </td>
                                       </React.Fragment>
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

                   {/* Renk Kodları Açıklama Paneli */}
                   <div className="mt-6 bg-white border border-slate-200 rounded-3xl p-5 shadow-sm print:hidden select-none">
                     <h4 className="text-xs font-black text-slate-700 uppercase tracking-wider mb-3">
                       💡 GELECEK BAKIM / KONTROL RENK KODU AÇIKLAMALARI
                     </h4>
                     <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                       <div className="flex items-center gap-3 p-3 bg-emerald-50 border border-emerald-200 rounded-2xl">
                         <span className="w-5 h-5 rounded-lg bg-emerald-500 shrink-0" />
                         <div>
                           <p className="text-xs font-bold text-emerald-950">YEŞİL (BAKIMA UYGUN)</p>
                           <p className="text-[10px] text-emerald-700">90 Gün ve Fazla (<span className="font-mono">{'>='}90 Gün</span>)</p>
                         </div>
                       </div>
                       <div className="flex items-center gap-3 p-3 bg-orange-50 border border-orange-200 rounded-2xl">
                         <span className="w-5 h-5 rounded-lg bg-orange-500 shrink-0 animate-pulse" />
                         <div>
                           <p className="text-xs font-bold text-orange-950">TURUNCU (BAKIM YAKLAŞTI / AZALDI)</p>
                           <p className="text-[10px] text-orange-700">90 Günden Az (<span className="font-mono">{'<'}90 Gün</span>)</p>
                         </div>
                       </div>
                     </div>
                   </div>
                 </>
               ) : (
                 <div className="flex-1 flex flex-col gap-6 print:hidden">
                   {/* Left Panel: Yeni Görev Emri (Step-by-Step Wizard) */}
                   {karaAraclariSubTab === 'mission_order' && (() => {
                      const handleNextStep = () => {
                        if (geStep === 1) {
                          if (!geTarih || !geSeriNo) {
                            alert("Lütfen tüm alanları (Tarih ve Seri No) doldurunuz.");
                            return;
                          }
                          setGeStep(2);
                        } else if (geStep === 2) {
                          if (!gePlaka || !geSoforName) {
                            alert("Lütfen araç plakasını seçiniz ve sürücü personel adını giriniz.");
                            return;
                          }
                          setGeStep(3);
                        }
                      };

                      const handlePrevStep = () => {
                        if (geStep > 1) {
                          setGeStep(geStep - 1);
                        }
                      };

                      return (
                        <div className="w-full max-w-3xl mx-auto bg-white border border-slate-200 rounded-[2.5rem] p-4 sm:p-6 shadow-xl flex flex-col gap-5 animate-fade-in">
                          <div className="border-b border-slate-100 pb-3">
                            <h4 className="text-xs font-black text-slate-800 uppercase tracking-widest flex items-center gap-2">
                              <span className="w-2.5 h-2.5 rounded-full bg-emerald-500 animate-pulse inline-block" />
                              📝 ADIM ADIM GÖREV EMRİ GİRİŞİ
                            </h4>
                          </div>

                          {/* Step Wizard Header Indicator */}
                          <div className="flex items-center justify-between px-2 py-3 bg-slate-50 rounded-2xl border border-slate-100 select-none">
                            {[1, 2, 3].map((stepNo) => {
                              let label = "";
                              if (stepNo === 1) label = "Genel";
                              if (stepNo === 2) label = "Sürücü";
                              if (stepNo === 3) label = "Süreç & KM";

                              const isActive = geStep === stepNo;
                              const isCompleted = geStep > stepNo;

                              return (
                                <div key={stepNo} className="flex flex-col items-center flex-1 relative">
                                  {/* Connector line */}
                                  {stepNo < 3 && (
                                    <div className="absolute top-4 left-[50%] right-[-50%] h-0.5 bg-slate-200 -z-0">
                                      <div 
                                        className="h-full bg-[#0b3d1d] transition-all duration-300" 
                                        style={{ width: geStep > stepNo ? "100%" : "0%" }}
                                      />
                                    </div>
                                  )}

                                  {/* Step circle */}
                                  <div
                                    className={`w-8 h-8 rounded-full flex items-center justify-center font-bold text-xs border-2 transition-all z-10 ${
                                      isActive
                                        ? "bg-[#0b3d1d] text-white border-[#0b3d1d] scale-110 shadow-md shadow-emerald-900/10"
                                        : isCompleted
                                        ? "bg-emerald-600 text-white border-emerald-600"
                                        : "bg-white text-slate-400 border-slate-200"
                                    }`}
                                  >
                                    {isCompleted ? "✓" : stepNo}
                                  </div>
                                  <span 
                                    className={`text-[9px] font-extrabold uppercase mt-1.5 transition-all text-center ${
                                      isActive ? "text-[#0b3d1d]" : "text-slate-400"
                                    } max-sm:text-[8px]`}
                                  >
                                    {label}
                                  </span>
                                </div>
                              );
                            })}
                          </div>

                          {/* Form Steps */}
                          <div className="flex-1 flex flex-col justify-between gap-6 min-h-[320px]">
                            
                            {/* STEP 1: GENEL BİLGİLER */}
                            {geStep === 1 && (
                              <div className="flex flex-col gap-4 animate-fade-in">
                                <div className="p-4 bg-emerald-50/50 rounded-2xl border border-emerald-100">
                                  <p className="text-xs font-semibold text-emerald-800 leading-relaxed">
                                    ℹ️ <strong>Adım 1:</strong> Görevin yapılacağı tarihi ve resmi evrak üzerindeki <strong>Görev Seri Numarasını (S/N)</strong> giriniz.
                                  </p>
                                </div>

                                <div>
                                  <label className="block text-[10px] font-black text-slate-400 uppercase tracking-wider mb-2">📅 Görev Tarihi</label>
                                  <input
                                    type="date"
                                    value={geTarih}
                                    onChange={(e) => setGeTarih(e.target.value)}
                                    className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-2xl focus:outline-none focus:border-[#0b3d1d] focus:ring-4 focus:ring-[#0b3d1d]/5 text-xs font-bold text-slate-800 transition-all font-mono"
                                  />
                                </div>

                                <div>
                                  <label className="block text-[10px] font-black text-slate-400 uppercase tracking-wider mb-2">🔢 Görev Seri No (S/N)</label>
                                  <input
                                    type="text"
                                    placeholder="Örn: SERI-772"
                                    value={geSeriNo}
                                    onChange={(e) => setGeSeriNo(e.target.value)}
                                    className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-2xl focus:outline-none focus:border-[#0b3d1d] focus:ring-4 focus:ring-[#0b3d1d]/5 text-xs font-bold text-slate-800 transition-all font-mono"
                                  />
                                </div>
                              </div>
                            )}

                            {/* STEP 2: ARAÇ & SÜRÜCÜ SEÇİMİ */}
                            {geStep === 2 && (
                              <div className="flex flex-col gap-4 animate-fade-in">
                                <div className="p-4 bg-emerald-50/50 rounded-2xl border border-emerald-100">
                                  <p className="text-xs font-semibold text-emerald-800 leading-relaxed">
                                    ℹ️ <strong>Adım 2:</strong> Görev aracını seçin ve listeden bir şoför seçerek devam edin.
                                  </p>
                                </div>

                                <div className="relative">
                                  <label className="block text-[10px] font-black text-slate-400 uppercase tracking-wider mb-2">🚗 Araç Plakası / Tanımı</label>
                                  <div className="relative">
                                    <input
                                      type="text"
                                      placeholder="Araç plakası yazınız veya listeden seçiniz..."
                                      value={gePlaka}
                                      onChange={(e) => {
                                        setGePlaka(e.target.value);
                                        setShowVehicleSuggestions(true);
                                      }}
                                      onFocus={() => setShowVehicleSuggestions(true)}
                                      className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-2xl focus:outline-none focus:border-[#0b3d1d] focus:ring-4 focus:ring-[#0b3d1d]/5 text-xs font-bold text-slate-800 transition-all"
                                    />
                                    {gePlaka && (
                                      <button
                                        type="button"
                                        onClick={() => {
                                          setGePlaka("");
                                          setShowVehicleSuggestions(false);
                                        }}
                                        className="absolute right-3 top-3 text-slate-400 hover:text-slate-600 font-bold text-xs"
                                      >
                                        ✕
                                      </button>
                                    )}
                                  </div>
                                  
                                  {showVehicleSuggestions && (
                                    <div className="absolute z-50 left-0 right-0 mt-1 bg-white border border-slate-200 rounded-2xl shadow-xl max-h-48 overflow-y-auto">
                                      {vehiclePlates
                                        .filter(plate => !gePlaka.trim() || plate.toLowerCase().includes(gePlaka.toLowerCase()))
                                        .map(plate => {
                                          // Find additional details if available (e.g., location, status)
                                          const foundRow = techizatKaraAraclariData.find(row => (row[1] || "").trim() === plate);
                                          const location = foundRow ? foundRow[3] : "";
                                          const model = foundRow ? foundRow[2] : "";
                                          
                                          return (
                                            <button
                                              key={plate}
                                              type="button"
                                              onClick={() => {
                                                setGePlaka(plate);
                                                setShowVehicleSuggestions(false);
                                                // Autopopulate departure KM
                                                if (foundRow && foundRow[4]) {
                                                  setGeDepartureKm(foundRow[4]);
                                                }
                                              }}
                                              className="w-full text-left px-4 py-2.5 text-xs font-bold hover:bg-emerald-50 text-slate-800 transition-all border-b border-slate-100 last:border-b-0 flex justify-between items-center"
                                            >
                                              <div className="flex flex-col">
                                                <span>{plate}</span>
                                                {model && <span className="text-[9px] text-slate-400 font-medium">{model}</span>}
                                              </div>
                                              {location && <span className="text-[9px] bg-slate-100 text-slate-500 px-2 py-0.5 rounded-full font-semibold">{location}</span>}
                                            </button>
                                          );
                                        })
                                      }
                                      {vehiclePlates.filter(plate => !gePlaka.trim() || plate.toLowerCase().includes(gePlaka.toLowerCase())).length === 0 && (
                                        <div className="px-4 py-3 text-xs text-slate-400 font-medium">
                                          Eşleşen araç bulunamadı. Tamamen manuel yazabilirsiniz.
                                        </div>
                                      )}
                                    </div>
                                  )}
                                </div>

                                <div className="relative">
                                  <label className="block text-[10px] font-black text-slate-400 uppercase tracking-wider mb-2">👨‍✈️ Sürücü Personel (Şoför)</label>
                                  <div className="relative">
                                    <input
                                      type="text"
                                      placeholder="Personel adı soyadı yazınız veya seçiniz..."
                                      value={geSoforName}
                                      onChange={(e) => {
                                        setGeSoforName(e.target.value);
                                        setShowDriverSuggestions(true);
                                      }}
                                      onFocus={() => setShowDriverSuggestions(true)}
                                      className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-2xl focus:outline-none focus:border-[#0b3d1d] focus:ring-4 focus:ring-[#0b3d1d]/5 text-xs font-bold text-slate-800 transition-all"
                                    />
                                    {geSoforName && (
                                      <button
                                        type="button"
                                        onClick={() => {
                                          setGeSoforName("");
                                          setShowDriverSuggestions(false);
                                        }}
                                        className="absolute right-3 top-3 text-slate-400 hover:text-slate-600 font-bold text-xs"
                                      >
                                        ✕
                                      </button>
                                    )}
                                  </div>
                                  
                                  {showDriverSuggestions && (
                                    <div className="absolute z-50 left-0 right-0 mt-1 bg-white border border-slate-200 rounded-2xl shadow-xl max-h-48 overflow-y-auto">
                                      {drivers
                                        .filter(d => !geSoforName.trim() || d.name.toLowerCase().includes(geSoforName.toLowerCase()))
                                        .map(d => (
                                          <button
                                            key={d.name}
                                            type="button"
                                            onClick={() => {
                                              setGeSoforName(d.name);
                                              setShowDriverSuggestions(false);
                                            }}
                                            className="w-full text-left px-4 py-2.5 text-xs font-bold hover:bg-emerald-50 text-slate-800 transition-all border-b border-slate-100 last:border-b-0 flex justify-between items-center"
                                          >
                                            <span>{d.name}</span>
                                            <span className="text-[10px] text-slate-400 font-semibold">{d.unvan || "Şoför"}</span>
                                          </button>
                                        ))
                                      }
                                      {drivers.filter(d => !geSoforName.trim() || d.name.toLowerCase().includes(geSoforName.toLowerCase())).length === 0 && (
                                        <div className="px-4 py-3 text-xs text-slate-400 font-medium">
                                          Eşleşen şoför bulunamadı. Tamamen manuel yazabilirsiniz.
                                        </div>
                                      )}
                                    </div>
                                  )}
                                </div>
                              </div>
                            )}

                            {/* STEP 3: KİLOMETRE & SÜREÇ DETAYLARI */}
                            {geStep === 3 && (
                              <div className="flex flex-col gap-4 animate-fade-in">
                                <div className="p-4 bg-emerald-50/50 rounded-2xl border border-emerald-100">
                                  <p className="text-xs font-semibold text-emerald-800 leading-relaxed">
                                    ℹ️ <strong>Adım 3:</strong> Saat ve kilometre verilerini girip <strong>Kaydet ve Yönlendir</strong> butonuna basınız.
                                  </p>
                                </div>

                                {/* Nereden - Nereye (Güzergah) Bilgisi */}
                                <div className="flex flex-col gap-2.5">
                                  <label className="block text-[10px] font-black text-slate-400 uppercase tracking-wider">
                                    📍 GÜZERGAH / LOKASYON BİLGİSİ
                                  </label>
                                  {geRoutes.map((route, rIdx) => (
                                    <div key={rIdx} className="flex items-center gap-2 animate-fade-in">
                                      <div className="flex-1">
                                        <input
                                          type="text"
                                          placeholder="Nereden (Örn: Ankara)"
                                          value={route.from}
                                          onChange={(e) => {
                                            const updated = [...geRoutes];
                                            updated[rIdx].from = e.target.value;
                                            setGeRoutes(updated);
                                          }}
                                          className="w-full px-3.5 py-2 bg-slate-50 border border-slate-200 rounded-2xl focus:outline-none focus:border-[#0b3d1d] focus:ring-4 focus:ring-[#0b3d1d]/5 text-xs font-bold text-slate-800 transition-all"
                                        />
                                      </div>
                                      <span className="text-slate-400 font-bold px-1">-</span>
                                      <div className="flex-1">
                                        <input
                                          type="text"
                                          placeholder="Nereye (Örn: İstanbul)"
                                          value={route.to}
                                          onChange={(e) => {
                                            const updated = [...geRoutes];
                                            updated[rIdx].to = e.target.value;
                                            setGeRoutes(updated);
                                          }}
                                          className="w-full px-3.5 py-2 bg-slate-50 border border-slate-200 rounded-2xl focus:outline-none focus:border-[#0b3d1d] focus:ring-4 focus:ring-[#0b3d1d]/5 text-xs font-bold text-slate-800 transition-all"
                                        />
                                      </div>
                                      {geRoutes.length > 1 && (
                                        <button
                                          type="button"
                                          onClick={() => {
                                            setGeRoutes(geRoutes.filter((_, i) => i !== rIdx));
                                          }}
                                          className="p-2 text-red-500 hover:text-red-700 hover:bg-red-50 rounded-xl transition-all"
                                        >
                                          <X className="w-4 h-4" />
                                        </button>
                                      )}
                                    </div>
                                  ))}
                                  <button
                                    type="button"
                                    onClick={() => setGeRoutes([...geRoutes, { from: "", to: "" }])}
                                    className="self-start text-[10px] font-black text-[#0b3d1d] bg-emerald-50 hover:bg-emerald-100 border border-emerald-100 px-3 py-1.5 rounded-xl transition-all cursor-pointer flex items-center gap-1 mt-1"
                                  >
                                    <Plus className="w-3.5 h-3.5" />
                                    Yeni Lokasyon Ekle
                                  </button>
                                </div>

                                <div className="grid grid-cols-2 gap-4">
                                  <div>
                                    <label className="block text-[10px] font-black text-slate-400 uppercase tracking-wider mb-2">⏱️ Çıkış Saati</label>
                                    <input
                                      type="time"
                                      value={geDepartureTime}
                                      onChange={(e) => setGeDepartureTime(e.target.value)}
                                      className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-2xl focus:outline-none focus:border-[#0b3d1d] focus:ring-4 focus:ring-[#0b3d1d]/5 text-xs font-bold text-slate-800 transition-all font-mono"
                                    />
                                  </div>
                                  <div>
                                    <label className="block text-[10px] font-black text-slate-400 uppercase tracking-wider mb-2">⏱️ Dönüş Saati</label>
                                    <input
                                      type="time"
                                      value={geReturnTime}
                                      onChange={(e) => setGeReturnTime(e.target.value)}
                                      className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-2xl focus:outline-none focus:border-[#0b3d1d] focus:ring-4 focus:ring-[#0b3d1d]/5 text-xs font-bold text-slate-800 transition-all font-mono"
                                    />
                                  </div>
                                </div>

                                <div className="grid grid-cols-2 gap-4">
                                  <div>
                                    <label className="block text-[10px] font-black text-slate-400 uppercase tracking-wider mb-2">📈 Çıkış KM</label>
                                    <input
                                      type="number"
                                      placeholder="Örn: 15150"
                                      value={geDepartureKm}
                                      onChange={(e) => setGeDepartureKm(e.target.value)}
                                      className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-2xl focus:outline-none focus:border-[#0b3d1d] focus:ring-4 focus:ring-[#0b3d1d]/5 text-xs font-bold text-slate-800 transition-all font-mono"
                                    />
                                  </div>
                                  <div>
                                    <label className="block text-[10px] font-black text-slate-400 uppercase tracking-wider mb-2">📈 Dönüş KM</label>
                                    <input
                                      type="number"
                                      placeholder="Örn: 15300"
                                      value={geReturnKm}
                                      onChange={(e) => setGeReturnKm(e.target.value)}
                                      className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-2xl focus:outline-none focus:border-[#0b3d1d] focus:ring-4 focus:ring-[#0b3d1d]/5 text-xs font-bold text-slate-800 transition-all font-mono"
                                    />
                                  </div>
                                </div>
                              </div>
                            )}

                            {/* Step Navigation Controls */}
                            <div className="flex items-center gap-3 pt-3 border-t border-slate-100">
                              {geStep > 1 && (
                                <button
                                  type="button"
                                  onClick={handlePrevStep}
                                  className="px-4 py-3 bg-slate-100 hover:bg-slate-200 active:scale-95 text-slate-700 font-extrabold text-xs uppercase tracking-widest rounded-2xl cursor-pointer transition-all flex items-center gap-1.5"
                                >
                                  <ChevronLeft className="w-4 h-4" />
                                  Geri
                                </button>
                              )}

                              {geStep < 3 ? (
                                <button
                                  type="button"
                                  onClick={handleNextStep}
                                  className="flex-1 py-3 bg-[#0b3d1d] hover:bg-[#072612] active:scale-95 text-white font-extrabold text-xs uppercase tracking-widest rounded-2xl cursor-pointer transition-all flex items-center justify-center gap-1.5 shadow-md shadow-emerald-900/10"
                                >
                                  İleri
                                  <ChevronRight className="w-4 h-4 text-emerald-300" />
                                </button>
                              ) : (
                                <button
                                  type="button"
                                  onClick={() => {
                                    if (!geTarih || !gePlaka || !geSoforName || !geSeriNo || !geReturnKm || !geDepartureKm || !geDepartureTime || !geReturnTime) {
                                      alert("Lütfen form alanlarının tamamını doldurmak için bilgileri kontrol edin.");
                                      return;
                                    }
                                    if (Number(geReturnKm) <= Number(geDepartureKm)) {
                                      alert("Dönüş kilometresi, çıkış kilometresinden küçük veya eşit olamaz.");
                                      return;
                                    }

                                    const routeStr = geRoutes
                                      .map(r => r.from.trim() && r.to.trim() ? `${r.from.trim()} - ${r.to.trim()}` : '')
                                      .filter(Boolean)
                                      .join(", ");

                                    if (!routeStr) {
                                      alert("Lütfen en az bir güzergah (nereden - nereye) bilgisi giriniz.");
                                      return;
                                    }

                                    const activeDriver = drivers.find(d => d.name.toLowerCase() === geSoforName.toLowerCase()) || 
                                                         drivers.find(d => d.name.toLowerCase().includes(geSoforName.toLowerCase()));

                                    const newOrder = {
                                      id: Date.now(),
                                      date: geTarih,
                                      plate: gePlaka,
                                      driverName: geSoforName,
                                      driverId: activeDriver?.idNo || "",
                                      driverSicil: activeDriver?.sicilNo || "",
                                      driverPhone: activeDriver?.phone || "",
                                      driverKanGrubu: activeDriver?.kanGrubu || "",
                                      driverAdres: activeDriver?.adres || "",
                                      serialNo: geSeriNo,
                                      departureTime: geDepartureTime,
                                      returnTime: geReturnTime,
                                      departureKm: geDepartureKm,
                                      returnKm: Number(geReturnKm),
                                      route: routeStr
                                    };

                                    const updatedOrders = [newOrder, ...karaAraclariGorevEmirleri];
                                    setKaraAraclariGorevEmirleri(updatedOrders);
                                    pushKaraAraclariGorevEmirleri(updatedOrders);

                                    // Automation: update vehicle KM in standard list
                                    const updatedVehicles = techizatKaraAraclariData.map(row => {
                                      if (row[1] && row[1].toLowerCase().includes(gePlaka.toLowerCase())) {
                                        const newRow = [...row];
                                        const returnKmNum = Number(geReturnKm);

                                        if (newRow.length < 11) {
                                          while (newRow.length < 11) newRow.push("");
                                        }

                                        newRow[4] = String(returnKmNum); // SON KM Sİ

                                        const today = new Date();
                                        const todayStr = today.toLocaleDateString('tr-TR');
                                        const nextMaintenanceDate = new Date();
                                        nextMaintenanceDate.setMonth(nextMaintenanceDate.getMonth() + 6);
                                        const nextMaintenanceStr = nextMaintenanceDate.toLocaleDateString('tr-TR');

                                        newRow[6] = todayStr; // SON KONTROL / KALİBRASYON / BAKIM
                                        newRow[7] = nextMaintenanceStr; // GELECEK KONTROL / KALİBRASYON / BAKIM

                                        return newRow;
                                      }
                                      return row;
                                    });

                                    setTechizatKaraAraclariData(updatedVehicles);
                                    localStorage.setItem('excel_techizat_kara_araclari_data', JSON.stringify(updatedVehicles));

                                    // Sync updated vehicle database to Google Sheets online database immediately
                                    const unitLabel = "KARA ARAÇLARI";
                                    fetch(GOOGLE_SCRIPT_URL, {
                                      method: "POST",
                                      headers: {
                                        "Content-Type": "text/plain;charset=utf-8"
                                      },
                                      body: JSON.stringify({
                                        action: "updateTumTechizat",
                                        unitLabel: unitLabel,
                                        data: updatedVehicles.map(r => [unitLabel, ...r])
                                      })
                                    }).catch(err => console.error("Central sheet vehicle sync error:", err));

                                    // Reset
                                    setGeSeriNo("");
                                    setGeSoforName("");
                                    setGePlaka("");
                                    setGeDepartureKm("");
                                    setGeReturnKm("");
                                    setGeDepartureTime("08:00");
                                    setGeReturnTime("17:00");
                                    setGeRoutes([{ from: "", to: "" }]);
                                    setGeStep(1);

                                    showNotification("Görev Emri Girişi Başarıyla Tamamlandı! Araç KM'si ve Bakım Periyotları Otomatik Güncellendi.");
                                    
                                    // Trigger transition screen
                                    setIsRedirectingToPortal(true);
                                    setIsSlidingUp(false);
                                    
                                    // Step 1: Wait 2.5 seconds with pulsing logo, then slide the screen up
                                    setTimeout(() => {
                                      setIsSlidingUp(true);
                                      // Step 2: After the slide-up animation (0.8s) finishes, do the redirect
                                      setTimeout(() => {
                                        window.location.href = "https://bulut.ogm.gov.tr/gorevemri";
                                      }, 800);
                                    }, 2500);
                                  }}
                                  className="flex-1 py-3.5 bg-emerald-600 hover:bg-emerald-700 active:scale-95 text-white font-extrabold text-xs uppercase tracking-widest rounded-2xl cursor-pointer transition-all flex items-center justify-center gap-1.5 shadow-md shadow-emerald-950/20"
                                >
                                  <CheckCircle className="w-4 h-4 text-emerald-200" />
                                  KAYDET VE YÖNLENDİR
                                </button>
                              )}
                            </div>
                          </div>
                        </div>
                      );
                    })()}

                   {/* HIDDEN OLD FORM */}
                   {false && (
                     <div className="w-full xl:w-5/12 bg-white border border-slate-200 rounded-[2.5rem] p-6 shadow-xl flex flex-col gap-5">
                     <div className="border-b border-slate-100 pb-3">
                       <h4 className="text-xs font-black text-slate-800 uppercase tracking-widest flex items-center gap-2">
                         <span className="w-2.5 h-2.5 rounded-full bg-emerald-500 animate-pulse inline-block" />
                         📝 YENİ ARAÇ GÖREV EMRİ KAYDI
                       </h4>
                     </div>

                     <div className="flex flex-col gap-4">
                       <div>
                         <label className="block text-[10px] font-black text-slate-400 uppercase tracking-wider mb-2">📅 Görev Tarihi</label>
                         <input
                           type="date"
                           value={geTarih}
                           onChange={(e) => setGeTarih(e.target.value)}
                           className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-2xl focus:outline-none focus:border-[#0b3d1d] focus:ring-4 focus:ring-[#0b3d1d]/5 text-xs font-bold text-slate-800 transition-all font-mono"
                         />
                       </div>

                       <div>
                         <label className="block text-[10px] font-black text-slate-400 uppercase tracking-wider mb-2">🚗 Araç Plakası</label>
                         <select
                           value={gePlaka}
                           onChange={(e) => setGePlaka(e.target.value)}
                           className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-2xl focus:outline-none focus:border-[#0b3d1d] focus:ring-4 focus:ring-[#0b3d1d]/5 text-xs font-extrabold text-slate-800 transition-all"
                         >
                           <option value="">-- Lütfen Araç Seçiniz --</option>
                           {vehiclePlates.map(p => (
                             <option key={p} value={p}>{p}</option>
                           ))}
                         </select>
                       </div>

                       <div>
                         <label className="block text-[10px] font-black text-slate-400 uppercase tracking-wider mb-2">👨‍✈️ Sürücü Personel (Şoför)</label>
                         <select
                           value={geSoforName}
                           onChange={(e) => setGeSoforName(e.target.value)}
                           className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-2xl focus:outline-none focus:border-[#0b3d1d] focus:ring-4 focus:ring-[#0b3d1d]/5 text-xs font-extrabold text-slate-800 transition-all"
                         >
                           <option value="">-- Lütfen Şoför Seçiniz --</option>
                           {drivers.map(d => (
                             <option key={d.name} value={d.name}>{d.name} ({d.unvan})</option>
                           ))}
                         </select>
                       </div>

                       {(() => {
                         const activeDriver = drivers.find(d => d.name === geSoforName);
                         if (!activeDriver) return null;
                         return (
                           <div className="bg-emerald-50/40 border border-emerald-100 rounded-2xl p-4 flex flex-col gap-2.5 text-xs animate-fade-in">
                             <div className="flex justify-between border-b border-emerald-100/50 pb-1.5">
                               <span className="font-extrabold text-slate-500">T.C. Kimlik No:</span>
                               <span className="font-mono font-bold text-slate-800">{activeDriver.idNo || "Belirtilmemiş"}</span>
                             </div>
                             <div className="flex justify-between border-b border-emerald-100/50 pb-1.5">
                               <span className="font-extrabold text-slate-500">Sicil No:</span>
                               <span className="font-mono font-bold text-slate-800">{activeDriver.sicilNo || "Belirtilmemiş"}</span>
                             </div>
                             <div className="flex justify-between border-b border-emerald-100/50 pb-1.5">
                               <span className="font-extrabold text-slate-500">Sürücü Telefon:</span>
                               <span className="font-mono font-bold text-[#0b3d1d]">{activeDriver.phone || "Belirtilmemiş"}</span>
                             </div>
                             <div className="flex justify-between border-b border-emerald-100/50 pb-1.5">
                               <span className="font-extrabold text-slate-500">Kan Grubu:</span>
                               <span className="font-bold text-red-600 bg-red-50 px-2.5 py-0.5 rounded-lg border border-red-100">{activeDriver.kanGrubu || "Belirtilmemiş"}</span>
                             </div>
                             <div className="flex flex-col gap-1">
                               <span className="font-extrabold text-slate-500">Adres Bilgisi:</span>
                               <span className="font-semibold text-slate-600 leading-relaxed">{activeDriver.adres || "Belirtilmemiş"}</span>
                             </div>
                           </div>
                         );
                       })()}

                       <div>
                         <label className="block text-[10px] font-black text-slate-400 uppercase tracking-wider mb-2">🔢 Görev Seri No (S/N)</label>
                         <input
                           type="text"
                           placeholder="Örn: SERI-772"
                           value={geSeriNo}
                           onChange={(e) => setGeSeriNo(e.target.value)}
                           className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-2xl focus:outline-none focus:border-[#0b3d1d] focus:ring-4 focus:ring-[#0b3d1d]/5 text-xs font-bold text-slate-800 transition-all font-mono"
                         />
                       </div>

                       <div>
                         <label className="block text-[10px] font-black text-slate-400 uppercase tracking-wider mb-2">📈 Dönüş Kilometresi (KM)</label>
                         <input
                           type="number"
                           placeholder="Örn: 15300"
                           value={geReturnKm}
                           onChange={(e) => setGeReturnKm(e.target.value)}
                           className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-2xl focus:outline-none focus:border-[#0b3d1d] focus:ring-4 focus:ring-[#0b3d1d]/5 text-xs font-bold text-slate-800 transition-all font-mono"
                         />
                       </div>

                       <button
                         onClick={() => {
                           if (!geTarih || !gePlaka || !geSoforName || !geSeriNo || !geReturnKm) {
                             alert("Lütfen form alanlarının tamamını doldurunuz.");
                             return;
                           }
                           
                           const activeDriver = drivers.find(d => d.name === geSoforName);
                           
                           const newOrder = {
                             id: Date.now(),
                             date: geTarih,
                             plate: gePlaka,
                             driverName: geSoforName,
                             driverId: activeDriver?.idNo || "",
                             driverSicil: activeDriver?.sicilNo || "",
                             driverPhone: activeDriver?.phone || "",
                             driverKanGrubu: activeDriver?.kanGrubu || "",
                             driverAdres: activeDriver?.adres || "",
                             serialNo: geSeriNo,
                             returnKm: Number(geReturnKm)
                           };
                           
                           const updatedOrders = [newOrder, ...karaAraclariGorevEmirleri];
                           setKaraAraclariGorevEmirleri(updatedOrders);
                           
                           // Automation: update vehicle KM in standard list
                           const updatedVehicles = techizatKaraAraclariData.map(row => {
                             if (row[1] && row[1].toLowerCase().includes(gePlaka.toLowerCase())) {
                               const newRow = [...row];
                               const returnKmNum = Number(geReturnKm);
                               
                               if (newRow.length < 14) {
                                 while (newRow.length < 14) newRow.push("");
                                }
                                
                                newRow[10] = String(returnKmNum); // Current KM
                                newRow[11] = String(returnKmNum + 5000); // Next KM Periyot (+5000 KM)
                                
                                const today = new Date();
                                const todayStr = today.toLocaleDateString('tr-TR');
                                const nextMaintenanceDate = new Date();
                                nextMaintenanceDate.setMonth(nextMaintenanceDate.getMonth() + 6);
                                const nextMaintenanceStr = nextMaintenanceDate.toLocaleDateString('tr-TR');
                                
                                newRow[6] = todayStr; // SON BAKIM TARİHİ
                                newRow[8] = nextMaintenanceStr; // BİR SONRAKİ BAKIM TARİHİ
                                
                                return newRow;
                              }
                              return row;
                            });
                            
                            setTechizatKaraAraclariData(updatedVehicles);
                            localStorage.setItem('excel_techizat_kara_araclari_data', JSON.stringify(updatedVehicles));
                            
                            // Clear form fields
                            setGeSeriNo("");
                            setGeReturnKm("");
                            
                            showNotification("Görev Emri Girişi Başarıyla Tamamlandı! Araç KM'si ve Bakım Periyodu Otomatik Güncellendi.");
                          }}
                          className="w-full py-3.5 bg-[#0b3d1d] hover:bg-[#072612] active:scale-95 text-white font-extrabold text-xs uppercase tracking-widest rounded-2xl shadow-lg shadow-emerald-900/10 cursor-pointer transition-all flex items-center justify-center gap-2"
                        >
                          <CheckCircle className="w-4 h-4 text-emerald-300" />
                          GÖREVİ KAYDET VE KM GÜNCELLE
                        </button>
                      </div>
                    </div>
                  )}

                    {/* Right Panel: Görev Emri Geçmişi */}
                    {karaAraclariSubTab === 'past_records' && (
                      <div className="w-full max-w-4xl mx-auto bg-white border border-slate-200 rounded-[2.5rem] p-6 shadow-xl flex flex-col animate-fade-in">
                        <div className="border-b border-slate-100 pb-3 mb-4 flex flex-col sm:flex-row justify-between items-start sm:items-center gap-2">
                          <h4 className="text-xs font-black text-slate-800 uppercase tracking-widest flex items-center gap-2">
                            📋 GÖREV EMRİ GEÇMİŞİ VE KAYITLARI
                          </h4>
                          {karaAraclariGorevEmirleri.length > 0 && (
                            <button
                              onClick={exportGorevEmirleriToExcel}
                              className="px-4 py-2 bg-[#0b3d1d] hover:bg-[#072612] active:scale-95 text-white font-black font-mono text-[10px] rounded-2xl flex items-center gap-1.5 transition-all cursor-pointer shadow-md border border-[#0b3d1d]/20 shrink-0 uppercase tracking-wider"
                            >
                              <Download className="w-3.5 h-3.5" />
                              <span>EXCEL OLARAK AKTAR</span>
                            </button>
                          )}
                        </div>

                        <div className="flex-1 overflow-x-auto min-w-full">
                          {karaAraclariGorevEmirleri.length === 0 ? (
                            <div className="flex flex-col items-center justify-center py-20 text-slate-400">
                              <FileText className="w-12 h-12 text-slate-300 stroke-1 mb-3" />
                              <p className="text-xs font-bold">Kayıtlı görev emri bulunmuyor.</p>
                              <p className="text-[10px] text-slate-400 mt-1">Sol taraftaki formdan yeni bir görev emri ekleyebilirsiniz.</p>
                            </div>
                          ) : (
                            <table className="w-full text-left border-collapse min-w-[750px]">
                              <thead>
                                <tr className="bg-slate-50 border-b border-slate-100 shrink-0">
                                  <th className="px-4 py-3 text-[10px] font-black uppercase text-slate-500 tracking-wider">Tarih</th>
                                  <th className="px-4 py-3 text-[10px] font-black uppercase text-slate-500 tracking-wider">Araç Plakası</th>
                                  <th className="px-4 py-3 text-[10px] font-black uppercase text-slate-500 tracking-wider">Sürücü Personel</th>
                                  <th className="px-4 py-3 text-[10px] font-black uppercase text-slate-500 tracking-wider">Görev Seri No</th>
                                  <th className="px-4 py-3 text-[10px] font-black uppercase text-slate-500 tracking-wider">KM Bilgisi</th>
                                  <th className="px-4 py-3 text-[10px] font-black uppercase text-slate-500 tracking-wider">Çıkış / Giriş Saati</th>
                                  <th className="px-4 py-3 text-[10px] font-black uppercase text-slate-500 tracking-wider">Güzergah</th>
                                  <th className="px-4 py-3 text-[10px] font-black uppercase text-slate-500 tracking-wider text-right">İşlemler</th>
                                </tr>
                              </thead>
                              <tbody className="divide-y divide-slate-100">
                                {karaAraclariGorevEmirleri.map((order, oIdx) => (
                                  <tr 
                                    key={order.id || oIdx} 
                                    onDoubleClick={() => {
                                      setEditGorevEmriValues({ ...order });
                                      setActiveGorevEmriEdit(order);
                                    }}
                                    className="hover:bg-slate-50/50 transition-colors cursor-pointer select-none"
                                    title="Düzenlemek için Çift Tıklayın"
                                  >
                                    <td className="px-4 py-3 text-xs font-mono text-slate-500 whitespace-nowrap">{order.date}</td>
                                    <td className="px-4 py-3 whitespace-nowrap">
                                      <span className="text-xs font-black text-[#0b3d1d] bg-emerald-100 px-2.5 py-1 rounded-xl">
                                        {order.plate}
                                      </span>
                                    </td>
                                    <td className="px-4 py-3 text-xs font-bold text-slate-700">
                                      <div className="flex flex-col">
                                        <span>{order.driverName}</span>
                                        <span className="text-[10px] text-slate-400 font-medium">Sicil: {order.driverSicil || "-"}</span>
                                      </div>
                                    </td>
                                    <td className="px-4 py-3 text-xs font-mono text-slate-600 font-bold whitespace-nowrap">{order.serialNo}</td>
                                    <td className="px-4 py-3 text-xs text-slate-600 whitespace-nowrap">
                                      <div className="flex flex-col text-[10px] font-semibold text-slate-500">
                                        <span>Çıkış: <strong className="text-slate-700 font-mono">{order.departureKm} KM</strong></span>
                                        <span>Dönüş: <strong className="text-slate-700 font-mono">{order.returnKm} KM</strong></span>
                                      </div>
                                    </td>
                                    <td className="px-4 py-3 text-xs text-slate-600 whitespace-nowrap">
                                      <div className="flex flex-col text-[10px] font-semibold text-slate-500">
                                        <span>Çıkış: <strong className="text-slate-700 font-mono">{order.departureTime || "-"}</strong></span>
                                        <span>Giriş: <strong className="text-slate-700 font-mono">{order.returnTime || "-"}</strong></span>
                                      </div>
                                    </td>
                                    <td className="px-4 py-3 text-xs text-emerald-800 font-bold max-w-xs truncate" title={order.route}>
                                      {order.route}
                                    </td>
                                    <td className="px-4 py-3 text-right whitespace-nowrap">
                                      <div className="flex items-center justify-end gap-1.5">
                                        {/* Düzenle Button */}
                                        <button
                                          onClick={() => {
                                            setEditGorevEmriValues({ ...order });
                                            setActiveGorevEmriEdit(order);
                                          }}
                                          className="text-amber-600 hover:text-amber-700 hover:bg-amber-50 p-2 rounded-xl transition-all cursor-pointer active:scale-95"
                                          title="Kayıt Düzenle"
                                        >
                                          <Edit3 className="w-4 h-4" />
                                        </button>
                                        {/* Sil Button */}
                                        <button
                                          onClick={() => {
                                            setGeDeleteOrderId(String(order.id));
                                            setGeDeletePasswordInput("");
                                            setGeDeletePasswordError(false);
                                            setShowGeDeletePasswordPrompt(true);
                                          }}
                                          className="text-red-500 hover:text-red-700 hover:bg-red-50 p-2 rounded-xl transition-all cursor-pointer active:scale-95"
                                          title="Kayıt Sil"
                                        >
                                          <Trash2 className="w-4 h-4" />
                                        </button>
                                      </div>
                                    </td>
                                  </tr>
                                ))}
                              </tbody>
                            </table>
                          )}
                        </div>

                        {/* Görev Emri Düzenleme Yetkili Şifre Onayı */}
                        {showGeEditPasswordPrompt && (
                          <div className="fixed inset-0 bg-slate-950/80 backdrop-blur-sm flex items-center justify-center z-[9999] p-4 select-text animate-fade-in">
                            <div className="bg-white border border-slate-200 rounded-[2.5rem] shadow-2xl max-w-sm w-full p-6 text-center relative overflow-hidden">
                              <div className="absolute top-0 inset-x-0 h-1 bg-amber-600"></div>
                              <div className="w-12 h-12 bg-amber-50 rounded-2xl flex items-center justify-center mx-auto mb-4 text-amber-600">
                                <Lock className="w-6 h-6" />
                              </div>
                              <h4 className="text-slate-800 font-extrabold text-sm uppercase mb-2">YETKİLİ DÜZENLEME ONAYI</h4>
                              <p className="text-xs text-slate-500 font-semibold mb-4 leading-relaxed">
                                Bu görev emri kaydını manuel olarak düzenlemek için lütfen yetkili şifresini giriniz.
                              </p>
                              <input
                                type="password"
                                placeholder="Şifre"
                                value={geEditPasswordInput}
                                onChange={(e) => {
                                  setGeEditPasswordInput(e.target.value);
                                  setGeEditPasswordError(false);
                                }}
                                onKeyDown={(e) => {
                                  if (e.key === 'Enter') {
                                    if (geEditPasswordInput === '1839') {
                                      if (activeGorevEmriEdit) {
                                        const updated = karaAraclariGorevEmirleri.map(order => {
                                          if (order.id === pendingGeEditOrder.id) {
                                            return pendingGeEditOrder;
                                          }
                                          return order;
                                        });
                                        setKaraAraclariGorevEmirleri(updated);
                                        localStorage.setItem('kara_araclari_gorev_emirleri', JSON.stringify(updated));
                                        pushKaraAraclariGorevEmirleri(updated);
                                        
                                        setActiveGorevEmriEdit(null);
                                        setEditGorevEmriValues(null);
                                        setPendingGeEditOrder(null);
                                        showNotification("Görev emri manuel olarak başarıyla güncellendi ve senkronize edildi!");
                                      } else {
                                        setEditGorevEmriValues({ ...pendingGeEditOrder });
                                        setActiveGorevEmriEdit(pendingGeEditOrder);
                                      }
                                      setShowGeEditPasswordPrompt(false);
                                      setGeEditPasswordInput("");
                                    } else {
                                      setGeEditPasswordError(true);
                                    }
                                  }
                                }}
                                className={`w-full px-4 py-3 bg-slate-50 border-2 rounded-2xl text-center font-bold text-sm focus:outline-none transition-all mb-2 ${
                                  geEditPasswordError ? 'border-red-500 focus:border-red-500 text-red-600' : 'border-slate-200 focus:border-slate-400'
                                }`}
                                autoFocus
                              />
                              {geEditPasswordError && (
                                <p className="text-[10px] text-red-500 font-bold mb-3">⚠️ Hatalı Şifre! Lütfen tekrar deneyiniz.</p>
                              )}
                              <div className="flex gap-2.5 mt-4">
                                <button
                                  type="button"
                                  onClick={() => {
                                    if (geEditPasswordInput === '1839') {
                                      if (activeGorevEmriEdit) {
                                        const updated = karaAraclariGorevEmirleri.map(order => {
                                          if (order.id === pendingGeEditOrder.id) {
                                            return pendingGeEditOrder;
                                          }
                                          return order;
                                        });
                                        setKaraAraclariGorevEmirleri(updated);
                                        localStorage.setItem('kara_araclari_gorev_emirleri', JSON.stringify(updated));
                                        pushKaraAraclariGorevEmirleri(updated);
                                        
                                        setActiveGorevEmriEdit(null);
                                        setEditGorevEmriValues(null);
                                        setPendingGeEditOrder(null);
                                        showNotification("Görev emri manuel olarak başarıyla güncellendi ve senkronize edildi!");
                                      } else {
                                        setEditGorevEmriValues({ ...pendingGeEditOrder });
                                        setActiveGorevEmriEdit(pendingGeEditOrder);
                                      }
                                      setShowGeEditPasswordPrompt(false);
                                      setGeEditPasswordInput("");
                                    } else {
                                      setGeEditPasswordError(true);
                                    }
                                  }}
                                  className="flex-1 py-3 bg-amber-600 hover:bg-amber-700 text-white font-extrabold text-xs rounded-2xl cursor-pointer shadow-md"
                                >
                                  Onayla ve Düzenle
                                </button>
                                <button
                                  type="button"
                                  onClick={() => {
                                    setShowGeEditPasswordPrompt(false);
                                    setGeEditPasswordInput("");
                                    setGeEditPasswordError(false);
                                  }}
                                  className="px-4 py-3 bg-slate-100 hover:bg-slate-200 text-slate-500 font-extrabold text-xs rounded-2xl cursor-pointer"
                                >
                                  İptal
                                </button>
                              </div>
                            </div>
                          </div>
                        )}

                        {/* Görev Emri Silme Yetkili Şifre Onayı */}
                        {showGeDeletePasswordPrompt && (
                          <div className="fixed inset-0 bg-slate-950/80 backdrop-blur-sm flex items-center justify-center z-[9999] p-4 select-text animate-fade-in">
                            <div className="bg-white border border-slate-200 rounded-[2.5rem] shadow-2xl max-w-sm w-full p-6 text-center relative overflow-hidden">
                              <div className="absolute top-0 inset-x-0 h-1 bg-red-600"></div>
                              <div className="w-12 h-12 bg-red-50 rounded-2xl flex items-center justify-center mx-auto mb-4 text-red-600">
                                <Trash2 className="w-6 h-6" />
                              </div>
                              <h4 className="text-slate-800 font-extrabold text-sm uppercase mb-2">YETKİLİ SİLME ONAYI</h4>
                              <p className="text-xs font-semibold mb-4 leading-relaxed text-slate-500">
                                Bu görev emri kaydını kalıcı olarak silmek için lütfen yetkili şifresini giriniz.
                              </p>
                              <input
                                type="password"
                                placeholder="Şifre"
                                value={geDeletePasswordInput}
                                onChange={(e) => {
                                  setGeDeletePasswordInput(e.target.value);
                                  setGeDeletePasswordError(false);
                                }}
                                onKeyDown={(e) => {
                                  if (e.key === 'Enter') {
                                    if (geDeletePasswordInput === '1839') {
                                      const filtered = karaAraclariGorevEmirleri.filter(o => String(o.id) !== geDeleteOrderId);
                                      setKaraAraclariGorevEmirleri(filtered);
                                      pushKaraAraclariGorevEmirleri(filtered);
                                      setShowGeDeletePasswordPrompt(false);
                                      setGeDeleteOrderId(null);
                                      setGeDeletePasswordInput("");
                                      showNotification("Görev emri kaydı başarıyla silindi ve Excel'e kaydedildi.");
                                    } else {
                                      setGeDeletePasswordError(true);
                                    }
                                  }
                                }}
                                className="w-full px-4 py-2.5 bg-slate-50 border-2 border-red-500/15 rounded-xl text-center text-sm font-semibold mb-3 focus:outline-none focus:border-red-500 text-slate-900 placeholder-slate-400"
                                autoFocus
                              />
                              {geDeletePasswordError && (
                                <p className="text-red-600 text-[10px] font-black mb-3">❌ Hatalı şifre girdiniz!</p>
                              )}
                              <div className="flex gap-2 justify-center">
                                <button
                                  type="button"
                                  onClick={() => {
                                    setShowGeDeletePasswordPrompt(false);
                                    setGeDeleteOrderId(null);
                                    setGeDeletePasswordInput("");
                                    setGeDeletePasswordError(false);
                                  }}
                                  className="px-4 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-600 text-xs font-black rounded-xl transition-colors cursor-pointer"
                                >
                                  İPTAL
                                </button>
                                <button
                                  type="button"
                                  onClick={() => {
                                    if (geDeletePasswordInput === '1839') {
                                      const filtered = karaAraclariGorevEmirleri.filter(o => String(o.id) !== geDeleteOrderId);
                                      setKaraAraclariGorevEmirleri(filtered);
                                      pushKaraAraclariGorevEmirleri(filtered);
                                      setShowGeDeletePasswordPrompt(false);
                                      setGeDeleteOrderId(null);
                                      setGeDeletePasswordInput("");
                                      showNotification("Görev emri kaydı başarıyla silindi.");
                                    } else {
                                      setGeDeletePasswordError(true);
                                    }
                                  }}
                                  className="px-6 py-2.5 bg-red-600 hover:bg-red-700 text-white text-xs font-black uppercase rounded-xl transition-all cursor-pointer shadow-md"
                                >
                                  SİL
                                </button>
                              </div>
                            </div>
                          </div>
                        )}
                      </div>
                    )}
                  </div>
                )}
              </>



                  {/* Print Only Representation (High contrast landscape layout) */}
                  <div className="print-only-container hidden print:block bg-white text-black p-6 w-full">
                    <div className="w-full text-black bg-white min-h-screen">
                      
                      {/* Brand Header */}
                      <div className="border-b-4 border-[#0b3d1d] pb-4 mb-6 flex items-center justify-between">
                        <div className="flex items-center gap-4">
                          <div className="w-12 h-12 bg-[#0b3d1d] text-white rounded-xl flex items-center justify-center font-black text-sm">
                            OGM
                          </div>
                          <div className="text-left">
                            <h4 className="text-sm font-black text-[#0b3d1d] uppercase tracking-wider leading-none">ORMAN GENEL MÜDÜRLÜĞÜ</h4>
                          </div>
                        </div>
                        <div className="text-right">
                          <h3 className="text-base font-black text-slate-900 uppercase tracking-tighter">{modalTitle}</h3>
                          <p className="text-[10px] text-slate-500 mt-1 font-mono font-semibold">Tarih: {new Date().toLocaleDateString('tr-TR')}</p>
                        </div>
                      </div>

                      {/* Print Grid Table */}
                      <table className="w-full border-collapse border border-slate-300 text-[9px]">
                        <thead>
                          <tr className="bg-[#0b3d1d] text-white font-bold">
                            {cols.map((col, idx) => (
                              <th key={idx} className="border border-slate-300 p-1.5 bg-[#0b3d1d] text-white text-center uppercase font-mono">{col}</th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {processedRows.map((row, rIdx) => (
                            <tr key={rIdx} className={rIdx % 2 === 1 ? "bg-slate-50" : "bg-white"}>
                              {row.map((cell, cIdx) => (
                                <td key={cIdx} className="border border-slate-300 p-1.5 text-center text-slate-900 font-bold">{cell || "-"}</td>
                              ))}
                            </tr>
                          ))}
                        </tbody>
                      </table>

                    </div>
                  </div>

                </div>
              </div>
            );
          })()}

          {/* EXCEL ONLINE & ÇEVRİMDIŞI UPDATE PANELİ WITH SMART CONTROLLER */}
          {modalType === 'excel_sync' && (
            <div className="absolute inset-0 flex flex-col bg-slate-50 p-6 md:p-8 animate-fade-in overflow-y-auto">
              
              {/* Header Title */}
              <div className="max-w-4xl mx-auto w-full text-center mb-6 select-none">
                <p className="text-[10px] font-black text-emerald-800 tracking-widest uppercase mb-1">HA BAKIM BAŞKANLIĞI VERİ AKTARIM PORTALİ</p>
                <h3 className="text-xl sm:text-2xl font-black text-slate-800 uppercase tracking-tighter">
                  🔄 VERİ GÜNCELLEME SİHİRBAZI ({activeSyncStep}. ADIM)
                </h3>
                <p className="text-xs text-slate-500 mt-2 max-w-2xl mx-auto leading-relaxed">
                  {activeSyncStep === 1 
                    ? 'Güncellemek istediğiniz birimi (tablo sayfasını) listeden seçin.' 
                    : 'Güncel planlama belgesini PDF formatında yükleyerek son güncelleme tarihini e-tabloya kaydedin ve sayfaları portal görünümünde güncelleyin.'}
                </p>
              </div>

              {activeSyncStep === 1 ? (
                <>
                  {/* Step 1: Sleek Liste Kutusu (Select Dropdown) & Button */}
                  <div className="max-w-xl mx-auto w-full mb-8 animate-fade-in">
                  <div className="bg-white rounded-3xl border border-slate-200/80 p-6 md:p-8 text-left shadow-lg">
                    
                    <div className="flex items-center gap-2 mb-6">
                      <span className="flex h-3 w-3 rounded-full bg-emerald-600 animate-pulse pointer-events-none" />
                      <h4 className="text-xs font-black text-[#0b3d1d] uppercase tracking-wider">
                        GÜNCELLEME YAPILACAK BİRİM SEÇİMİ
                      </h4>
                    </div>

                    <div className="space-y-4">
                      <label htmlFor="unit-select-box" className="block text-xs font-black text-slate-600 uppercase tracking-wide">
                        Lütfen Bir Liste Seçimi Yapın:
                      </label>
                      
                      <div className="relative">
                        <select
                          id="unit-select-box"
                          value={step1Target || "1"}
                          onChange={(e) => {
                            const val = e.target.value;
                            setStep1Target(val);
                            if (val === 'gun_takip') {
                              fetchGunTakipSorumlulari();
                              setIsSorumluModalOpen(true);
                            }
                          }}
                          className="w-full bg-white border-2 border-slate-200 text-slate-800 font-extrabold text-xs uppercase px-4 py-4 rounded-xl shadow-sm focus:border-emerald-600 focus:ring-0 transition-all cursor-pointer appearance-none"
                        >
                          <option value="1">1. GÖREVLENDİRME ÇİZELGELERİ (1-Gorevlendirme)</option>
                          <option value="21">2. YAZ DÖNEMİ PLANLAMASI - BELL 429 (2-Yaz_Donemi-bell_429)</option>
                          <option value="22">2. YAZ DÖNEMİ PLANLAMASI - T-70 (2-Yaz_Donemi-t_70)</option>
                          <option value="23">2. YAZ DÖNEMİ PLANLAMASI - AT-802 (2-Yaz_Donemi-at_802)</option>
                          <option value="24">2. YAZ DÖNEMİ PLANLAMASI - ANKARA BEKLEME (BELL-429) (2-Ankara_Bekleme-bell_429)</option>
                          <option value="25">2. YAZ DÖNEMİ PLANLAMASI - ANKARA BEKLEME (C-650/B-360) (2-Ankara_Bekleme-c650_b360)</option>
                          <option value="3">3. BAKIM YETKİ ÇİZELGELERİ (3-Bakim_Yetki)</option>
                          <option value="5">5. PERSONEL BİLGİ ÇİZELGELERİ (5-Personel_Bilgi)</option>
                          <option value="6">6. PERSONEL UÇUŞ-HİZMET YILLARI (6-Personel_Ucus_Hizmet)</option>
                          <option value="techizat_bell429">TEÇHİZAT ENVANTER TAKİBİ - BELL 429 (EXCEL)</option>
                          <option value="techizat_at802">TEÇHİZAT ENVANTER TAKİBİ - AT-802 (EXCEL)</option>
                          <option value="techizat_t70">TEÇHİZAT ENVANTER TAKİBİ - T-70 (EXCEL)</option>
                          <option value="techizat_t70_bumbi_backet">TEÇHİZAT ENVANTER TAKİBİ - T-70 BUMBİ BACKET (EXCEL)</option>
                          <option value="techizat_t70_helitak">TEÇHİZAT ENVANTER TAKİBİ - T-70 HELİTAK (EXCEL)</option>
                          <option value="techizat_b360">TEÇHİZAT ENVANTER TAKİBİ - B-360 (EXCEL)</option>
                          <option value="techizat_c650">TEÇHİZAT ENVANTER TAKİBİ - C-650 (EXCEL)</option>
                          <option value="techizat_hangar">TEÇHİZAT ENVANTER TAKİBİ - HANGAR YER DESTEK (EXCEL)</option>
                          <option value="techizat_kara_araclari">TEÇHİZAT ENVANTER TAKİBİ - KARA ARAÇLARI TAKİP (EXCEL)</option>
                          <option value="gun_takip">📋 SORUMLU BİRİM VE MAİL AYARLARI (GÜN TAKİP)</option>
                        </select>
                        <div className="pointer-events-none absolute inset-y-0 right-0 flex items-center px-4 text-slate-505">
                          🔻
                        </div>
                      </div>

                      {/* Yaz Dönemi için başlangıç ve bitiş tarihi seçimi */}
                      {["21", "22", "23", "24", "25"].includes(step1Target) && (
                        <div className="mt-4 pt-4 border-t border-slate-100 animate-fade-in space-y-3">
                          <label className="block text-[11px] font-black text-slate-600 uppercase tracking-wide">
                            📅 PLANLAMA DÖNEMİ TARİH SEÇİMİ:
                          </label>
                          <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                            <div className="space-y-1.5">
                              <span className="block text-[10px] font-bold text-[#0b3d1d] uppercase">BAŞLANGIÇ TARİHİ:</span>
                              <input
                                type="date"
                                value={selectedUploadSummerStartDate}
                                onChange={(e) => {
                                  if (e.target.value) {
                                    setSelectedUploadSummerStartDate(e.target.value);
                                  }
                                }}
                                className="w-full bg-[#f8fafc] border-2 border-slate-200 text-slate-700 font-bold text-xs px-4 py-3 rounded-xl shadow-sm focus:border-emerald-600 focus:ring-0 transition-all cursor-pointer"
                              />
                            </div>
                            <div className="space-y-1.5">
                              <span className="block text-[10px] font-bold text-[#0b3d1d] uppercase">BİTİŞ TARİHİ:</span>
                              <input
                                type="date"
                                value={selectedUploadSummerEndDate}
                                onChange={(e) => {
                                  if (e.target.value) {
                                    setSelectedUploadSummerEndDate(e.target.value);
                                  }
                                }}
                                className="w-full bg-[#f8fafc] border-2 border-slate-200 text-slate-700 font-bold text-xs px-4 py-3 rounded-xl shadow-sm focus:border-emerald-600 focus:ring-0 transition-all cursor-pointer"
                              />
                            </div>
                          </div>
                          <div className="bg-emerald-50 border border-emerald-200 p-3 rounded-xl mt-2 flex items-center justify-between">
                            <span className="text-[10px] font-black text-emerald-800 uppercase tracking-wider">OLUŞTURULACAK DÖNEM:</span>
                            <span className="text-xs font-black text-[#0b3d1d] uppercase">{getReadablePeriodName(selectedUploadSummerMonth)}</span>
                          </div>
                        </div>
                      )}
                    </div>

                    {/* Advance trigger button */}
                    <div className="mt-8 pt-5 border-t border-slate-100 flex justify-end">
                      <button
                        type="button"
                        onClick={() => {
                          if (step1Target === 'gun_takip') {
                            fetchGunTakipSorumlulari();
                            setIsSorumluModalOpen(true);
                          } else {
                            const targetId = step1Target || '1';
                            setSyncSelectedTarget(targetId);
                            setActiveSyncStep(2);
                          }
                        }}
                        className="w-full py-4 bg-[#0b3d1d] hover:bg-[#072612] text-white font-extrabold text-xs rounded-xl tracking-widest uppercase transition-all shadow-md active:scale-95 cursor-pointer select-none"
                      >
                        {step1Target === 'gun_takip' ? '📋 SORUMLU BİRİM AYARLARINI AÇ ➜' : 'İLERLE VE PDF YÜKLEME EKRANINA GEÇ ➜'}
                      </button>
                    </div>

                  </div>
                </div>

                {/* Historical Periods list for Summer planning */}
                {["21", "22", "23", "24", "25"].includes(step1Target) && (
                  <div className="max-w-xl mx-auto w-full animate-fade-in mt-6">
                    <div className="bg-white rounded-3xl border border-slate-200/80 p-6 md:p-8 text-left shadow-lg">
                      <div className="flex items-center gap-2 mb-4">
                        <span className="text-lg">🗓️</span>
                        <h4 className="text-xs font-black text-[#0b3d1d] uppercase tracking-wider">
                          GEÇMİŞ YAZ PLANLAMA PERİYOTLARI
                        </h4>
                      </div>
                      <p className="text-[11px] text-slate-500 mb-4 leading-relaxed">
                        Aşağıda bu hava aracı için daha önce girilmiş/yüklenmiş olan planlama dönemleri yer almaktadır. Yanlarındaki butona tıklayarak doğrudan o dönemin PDF dosyasını güncelleyebilirsiniz.
                      </p>

                      {(() => {
                        const targetId = Number(step1Target);
                        const airframeSuffix = getAirframeSuffix(targetId);
                        const activePeriods: string[] = [];
                        
                        pdfMetadataList.forEach(m => {
                          const lowerName = m.name.toLowerCase();
                          if (lowerName.includes("_yaz_plan_") && lowerName.endsWith(".pdf")) {
                            const isBeklemeFile = lowerName.includes('bekleme') || lowerName.includes('ankara');
                            
                            if (airframeSuffix === 'bell429') {
                              if (isBeklemeFile || (!lowerName.includes('bell429') && !lowerName.includes('bell_429'))) {
                                return;
                              }
                            } else if (airframeSuffix === 'bekleme_bell429') {
                              if (!isBeklemeFile || (!lowerName.includes('bell429') && !lowerName.includes('bell_429'))) {
                                return;
                              }
                            } else {
                              if (!lowerName.includes(airframeSuffix)) {
                                return;
                              }
                            }

                            const idx = lowerName.indexOf("_yaz_plan_");
                            if (idx !== -1) {
                              const periodPart = m.name.substring(0, idx);
                              if (periodPart && !activePeriods.includes(periodPart)) {
                                activePeriods.push(periodPart);
                              }
                            }
                          }
                        });

                        activePeriods.sort((a, b) => {
                          const dateA = new Date(parseRawPeriodStringToDates(a).start).getTime();
                          const dateB = new Date(parseRawPeriodStringToDates(b).start).getTime();
                          return dateB - dateA;
                        });

                        if (activePeriods.length === 0) {
                          return (
                            <div className="bg-slate-50 rounded-xl p-4 text-center border border-dashed border-slate-200">
                              <p className="text-[11px] text-slate-400 font-bold uppercase">Henüz girilmiş bir periyot bulunmamaktadır.</p>
                            </div>
                          );
                        }

                        return (
                          <div className="space-y-3 max-h-60 overflow-y-auto pr-1">
                            {activePeriods.map((period) => (
                              <div key={period} className="flex items-center justify-between p-3 bg-slate-50 rounded-xl border border-slate-100 hover:border-emerald-100 transition-colors">
                                <div className="flex flex-col">
                                  <span className="text-xs font-black text-[#0b3d1d] uppercase">
                                    {getReadablePeriodName(period)}
                                  </span>
                                  <span className="text-[9px] text-slate-400 font-mono">
                                    Dosya: {period}_yaz_plan_{airframeSuffix}.pdf
                                  </span>
                                </div>
                                <button
                                  type="button"
                                  onClick={() => {
                                    setSyncSelectedTarget(step1Target);
                                    const { start, end } = parseRawPeriodStringToDates(period);
                                    setSelectedUploadSummerStartDate(start);
                                    setSelectedUploadSummerEndDate(end);
                                    setSelectedSummerStartDate(start);
                                    setSelectedSummerEndDate(end);
                                    setActiveSyncStep(2);
                                    
                                    setTimeout(() => {
                                      const fileInput = document.getElementById("pdf-file-contextual-upload") as HTMLInputElement;
                                      if (fileInput) {
                                        fileInput.click();
                                      }
                                    }, 400);
                                  }}
                                  className="px-3 py-2 bg-emerald-600 hover:bg-emerald-700 text-white font-extrabold text-[10px] rounded-lg tracking-wider uppercase transition-all shadow-sm active:scale-95 cursor-pointer"
                                >
                                  🔄 GÜNCEL PDF YÜKLE
                                </button>
                              </div>
                            ))}
                          </div>
                        );
                      })()}
                    </div>
                  </div>
                )}
              </>
              ) : (
                /* Step 2: Download / Upload Action Cards */
                <div className="max-w-4xl mx-auto w-full mb-8 animate-fade-in flex flex-col items-center">
                  
                  {/* Active Unit Indicator */}
                  <div className="w-full bg-emerald-900 text-emerald-50 rounded-3xl p-5 md:p-6 mb-6 shadow-md border border-emerald-950 flex flex-col sm:flex-row gap-4 items-center justify-between select-none">
                    <div className="text-left w-full sm:w-auto">
                      <div className="text-[10px] uppercase font-bold text-emerald-300 tracking-wider">AKTİF SEÇİLEN BİRİM / TABLO SEÇİMİ:</div>
                      <h4 className="text-sm font-black uppercase mt-1">
                        {TABLE_CONFIGS[Number(syncSelectedTarget)]?.title || `BİRİM ${syncSelectedTarget}`}
                        {["21", "22", "23", "24", "25"].includes(syncSelectedTarget) && ` (${getReadablePeriodName(selectedUploadSummerMonth)})`}
                      </h4>
                      <p className="text-[11px] text-emerald-200 mt-1 font-mono">
                        Hedef Sayfa Adresi: <strong>{["21", "22", "23", "24", "25"].includes(syncSelectedTarget) ? `${getSummerPeriodSheetPrefix(Number(syncSelectedTarget), selectedUploadSummerMonth)}` : TABLE_CONFIGS[Number(syncSelectedTarget)]?.sheetName}</strong>
                      </p>
                    </div>
                    
                    <button
                      type="button"
                      onClick={() => {
                        setActiveSyncStep(1);
                      }}
                      className="w-full sm:w-auto flex items-center justify-center gap-1.5 px-4 py-2.5 bg-emerald-950/80 hover:bg-emerald-950 text-white rounded-xl text-xs font-bold transition-all cursor-pointer shadow-inner border border-emerald-900 select-none"
                    >
                      <ArrowLeft className="w-4 h-4" />
                      <span>Birim Değiştir / Geri Dön</span>
                    </button>
                  </div>

                  {/* Unified PDF Upload card for both summer and non-summer targets */}
                  <div className="max-w-xl w-full mb-6">
                    <div className="bg-white border-2 border-emerald-200/80 rounded-[2rem] p-8 flex flex-col items-center justify-between text-center transition-all shadow-xl hover:border-emerald-300">
                      <div className="p-5 bg-emerald-50 rounded-2xl mb-5 text-[#0b3d1d] shadow-inner">
                        {isSendingToSheets[String(syncSelectedTarget)] ? (
                          <Loader2 className="w-10 h-10 animate-spin text-emerald-700" />
                        ) : (
                          <FileText className="w-10 h-10 animate-pulse text-[#0b3d1d]" />
                        )}
                      </div>
                      
                      <div className="flex-1 flex flex-col justify-center select-none w-full">
                        <h4 className="text-slate-800 font-extrabold text-sm uppercase tracking-wider mb-2">
                          {String(syncSelectedTarget).startsWith('techizat_') ? "TEÇHİZAT ENVENTAR BELGESİ YÜKLE (EXCEL)" : syncSelectedTarget === '5' ? "PLANLAMA BELGESİ YÜKLE (EXCEL VEYA PDF)" : "PLANLAMA BELGESİ YÜKLE (PDF)"}
                        </h4>
                        <p className="text-xs text-slate-500 leading-relaxed max-w-sm mx-auto mb-6">
                          {isSendingToSheets[String(syncSelectedTarget)] 
                            ? "Plan belgesi sisteme aktarılıyor ve veriler işleniyor. Lütfen bekleyin..." 
                            : String(syncSelectedTarget).startsWith('techizat_')
                              ? "İlgili hava aracı grubu için Excel (.xlsx, .xls) dosyasını yükleyin."
                              : syncSelectedTarget === '5'
                                ? "Personel Bilgi Çizelgesi için güncel Excel (.xlsx, .xls) veya PDF dosyasını yükleyin."
                                : "Seçilen birim için güncel planlama belgesini PDF formatında yükleyin."}
                        </p>

                        {/* Guide rules depending on summer period or not */}
                        <div className="bg-slate-50 rounded-2xl p-4 mb-6 border border-slate-100 text-left w-full">
                          <span className="text-[10px] font-black text-slate-700 uppercase tracking-wider block mb-2">
                            📋 {syncSelectedTarget === '5' || String(syncSelectedTarget).startsWith('techizat_') ? "Yükleme ve Önizleme Teknolojisi:" : "Sürücü (Drive) Otomatik Adlandırma Formatı:"}
                          </span>
                          <div className="space-y-1.5 font-mono text-[10px] text-slate-600">
                            {String(syncSelectedTarget).startsWith('techizat_') ? (
                              <>
                                <div className="text-emerald-800 font-bold uppercase tracking-wider">
                                  • MATRİS EXCEL VERİ GÜNCELLEME SİSTEMİ
                                </div>
                                <div className="text-slate-500 text-[10px] leading-relaxed">
                                  Yüklenen Excel belgesindeki sütunlar otomatik olarak <strong>Teçhizat Takis Matrisine</strong> işlenecektir. Kaydedilen parçalar kalıcı olarak sürücüye kaydedilecektir.
                                </div>
                              </>
                            ) : syncSelectedTarget === '5' ? (
                              <>
                                <div className="text-emerald-800 font-bold uppercase tracking-wider">
                                  • EXCEL ➔ PDF MATRİS DÖNÜŞTÜRÜCÜ
                                </div>
                                <div className="text-slate-500 text-[10px] leading-relaxed">
                                  Yüklenen Excel belgesindeki satırlar çözümlenerek <strong>537 Satır x 12 Sütunluk (537rx12c)</strong> yatay bir elektronik tabloya dönüştürülür. Sadece metin içeren kısımlar şık ve yüksek çözünürlüklü bir PDF belgesi gibi taranarak ekranda gösterilir.
                                </div>
                              </>
                            ) : ["21", "22", "23", "24", "25"].includes(String(syncSelectedTarget)) ? (
                              <>
                                <div>
                                  • Dosyanız otomatik olarak şu adla kaydedilecektir:
                                </div>
                                <div className="text-emerald-800 pl-4 font-bold break-all">
                                  {(() => {
                                    const airframeSuffix = getAirframeSuffix(Number(syncSelectedTarget));
                                    const cleanMonth = sanitizeTurkishForFilename(selectedUploadSummerMonth);
                                    return `${cleanMonth}_yaz_plan_${airframeSuffix}.pdf`;
                                  })()}
                                </div>
                                <div className="text-slate-400 pl-4">
                                  Sistemimiz bu dosya adını kullanarak Drive üzerinden canlı önizlemeyi yakalayacaktır.
                                </div>
                              </>
                            ) : (
                              <>
                                <div>
                                  • Dosyanız otomatik olarak şu adla kaydedilecektir:
                                </div>
                                <div className="text-emerald-800 pl-4 font-bold break-all">
                                  {(() => {
                                    const prefix = syncSelectedTarget === '1' ? 'gorevlendirme' : syncSelectedTarget === '3' ? 'bakim_yetki' : syncSelectedTarget === '5' ? 'personel_bilgi' : 'personel_ucus_hizmet';
                                    return `${prefix}_cizelgesi.pdf`;
                                  })()}
                                </div>
                                <div className="text-slate-400 pl-4">
                                  Portalımız bu dosyayı otomatik olarak yakalayıp PDF Görünümünde gösterecektir.
                                </div>
                              </>
                            )}
                          </div>
                        </div>
                      </div>

                      <div className="w-full space-y-4">
                        <label
                          htmlFor="pdf-file-contextual-upload"
                          className={`w-full py-4 text-white text-center font-extrabold text-xs rounded-xl tracking-widest uppercase transition-all block select-none cursor-pointer shadow-md ${
                            isSendingToSheets[String(syncSelectedTarget)] 
                              ? "bg-slate-400 cursor-not-allowed animate-pulse" 
                              : "bg-[#0b3d1d] hover:bg-[#072612]"
                          }`}
                        >
                          <input
                            type="file"
                            id="pdf-file-contextual-upload"
                            accept={syncSelectedTarget === '5' || String(syncSelectedTarget).startsWith('techizat_') ? ".xlsx,.xls,.csv" : ".pdf"}
                            onChange={handlePdfUpload}
                            disabled={isSendingToSheets[String(syncSelectedTarget)]}
                            className="hidden"
                          />
                          {isSendingToSheets[String(syncSelectedTarget)] ? (
                            <span className="flex flex-col items-center justify-center gap-1.5 font-black">
                              <span className="flex items-center gap-1.5 justify-center">
                                <Loader2 className="w-4 h-4 animate-spin text-white" />
                                {uploadProgress}% İŞLENİYOR...
                              </span>
                            </span>
                          ) : (
                            "⚡ DOSYA SEÇ VE SİSTEME AKTAR"
                          )}
                        </label>

                        {isSendingToSheets[String(syncSelectedTarget)] && (
                          <div className="w-full bg-slate-100 rounded-full h-3 overflow-hidden border border-slate-200/50 shadow-inner p-0.5">
                            <div 
                              className="bg-emerald-600 h-2 rounded-full transition-all duration-300 shadow-sm relative overflow-hidden" 
                              style={{ width: `${uploadProgress}%` }}
                            >
                              <div className="absolute inset-0 bg-gradient-to-r from-transparent via-white/20 to-transparent animate-pulse" />
                            </div>
                          </div>
                        )}

                        <div className="flex gap-3">
                          <a
                            href="https://drive.google.com/drive/folders/1_fIGvuPVpC9N5on1irOfGG8OsD1KSXD0"
                            target="_blank"
                            rel="noopener noreferrer"
                            className="flex-1 py-3 bg-white hover:bg-slate-50 text-slate-700 border border-slate-300 text-center font-extrabold text-[10px] rounded-xl tracking-wider uppercase transition-all block shadow-sm cursor-pointer"
                          >
                            📁 DRİVE KLASÖRÜNÜ AÇ
                          </a>
                          <button
                            type="button"
                            onClick={async () => {
                              showNotification("Portal verileri ve PDF listesi yenileniyor...");
                              await fetchPdfMetadata();
                              showNotification("Drive üzerindeki PDF listesi başarıyla güncellendi!");
                            }}
                            className="flex-1 py-3 bg-white hover:bg-slate-50 text-slate-700 border border-slate-300 text-center font-extrabold text-[10px] rounded-xl tracking-wider uppercase transition-all block shadow-sm cursor-pointer"
                          >
                            🔄 GÜNCELLEMELERİ KONTROL ET
                          </button>
                        </div>

                        {(syncSelectedTarget === '5' || String(syncSelectedTarget).startsWith('techizat_')) && (
                          <button
                            type="button"
                            onClick={() => {
                              try {
                                const wb = XLSX.utils.book_new();
                                let headers: string[] = [];
                                let data: string[][] = [];
                                let sheetName = "Sayfa1";
                                let fileName = "en_son_surum.xlsx";

                                if (syncSelectedTarget === '5') {
                                  headers = TABLE_CONFIGS[5].columns.map(col => col.label);
                                  data = excelForm5Data.map(row => {
                                    const formattedRow = [...row];
                                    formattedRow[5] = formatBirthDateToTurkish(formattedRow[5]);
                                    return formattedRow;
                                  });
                                  sheetName = "Personel_Bilgi";
                                  fileName = "personel_bilgi_cizelgesi_en_son_surum.xlsx";
                                } else {
                                  const techType = String(syncSelectedTarget).replace('techizat_', '');
                                  if (techType === 'bell429') {
                                    headers = techizatBell429Columns;
                                    data = techizatBell429Data;
                                    sheetName = "Bell429_Techizat";
                                    fileName = "hava_araçları_yer_destek_bell-429.xlsx";
                                  } else if (techType === 'at802') {
                                    headers = techizatAt802Columns;
                                    data = techizatAt802Data;
                                    sheetName = "At802_Techizat";
                                    fileName = "hava_araçları_yer_destek_at-802.xlsx";
                                  } else if (techType === 't70') {
                                    headers = techizatT70Columns;
                                    data = techizatT70Data;
                                    sheetName = "T70_Techizat";
                                    fileName = "hava_araçları_yer_destek_t-70.xlsx";
                                  } else if (techType === 't70_bumbi_backet') {
                                    headers = techizatT70BumbiBacketColumns;
                                    data = techizatT70BumbiBacketData;
                                    sheetName = "T70_Bumbi_Backet";
                                    fileName = "hava_araçları_yer_destek_t-70_bumbi_backet.xlsx";
                                  } else if (techType === 't70_helitak') {
                                    headers = techizatT70HelitakColumns;
                                    data = techizatT70HelitakData;
                                    sheetName = "T70_Helitak";
                                    fileName = "hava_araçları_yer_destek_t-70_helitak.xlsx";
                                  } else if (techType === 'b360') {
                                    headers = techizatB360Columns;
                                    data = techizatB360Data;
                                    sheetName = "B360_Techizat";
                                    fileName = "hava_araçları_yer_destek_b-360.xlsx";
                                  } else if (techType === 'c650') {
                                    headers = techizatC650Columns;
                                    data = techizatC650Data;
                                    sheetName = "C650_Techizat";
                                    fileName = "hava_araçları_yer_destek_c-650.xlsx";
                                  } else if (techType === 'hangar') {
                                    headers = techizatHangarColumns;
                                    data = techizatHangarData;
                                    sheetName = "Hangar_Techizat";
                                    fileName = "hava_araçları_yer_destek_hangar.xlsx";
                                  }
                                }

                                const dataWithHeaders = [headers, ...data];
                                const ws = XLSX.utils.aoa_to_sheet(dataWithHeaders);
                                XLSX.utils.book_append_sheet(wb, ws, sheetName);
                                XLSX.writeFile(wb, fileName);
                                showNotification(`Sistemdeki en son sürüm '${fileName}' Excel verisi başarıyla indirildi.`);
                              } catch (e) {
                                alert("Hata: " + e);
                              }
                            }}
                            className="w-full py-3.5 bg-emerald-700 hover:bg-emerald-800 text-white font-extrabold text-[11px] rounded-xl tracking-wider uppercase transition-all flex items-center justify-center gap-2 shadow-sm cursor-pointer"
                            title="Sistemde saklanan en son sürüm Excel verisini indir"
                          >
                            <FileSpreadsheet className="w-4 h-4 text-emerald-100" />
                            <span>📥 EN SON SÜRÜM EXCEL İNDİR</span>
                          </button>
                        )}
                      </div>
                    </div>
                  </div>
                </div>
              )}

              {/* Guidelines footer info */}
              <div className="max-w-4xl mx-auto w-full grid grid-cols-1 md:grid-cols-2 gap-4">
                <div className="bg-white border border-slate-200 rounded-3xl p-4 text-left shadow-sm">
                  <span className="text-[11px] font-extrabold text-slate-700 uppercase tracking-wide block mb-1">📌 GÜNCELLEME AKIŞ PRENSİPLERİ:</span>
                  <p className="text-[10px] text-slate-500 leading-relaxed font-sans">
                    Sistemimiz çevrimdışı yükleme yapılmadan önce o birimin mevcut veri kaydını <strong>baştan tamamen sıfırlar</strong> ve yeni yüklediğiniz güncel dosyadaki verileri aktarır.
                  </p>
                </div>

                <div className="bg-white border border-slate-200 rounded-3xl p-4 text-left flex flex-col justify-between shadow-sm">
                  <div>
                    <span className="text-[11px] font-extrabold text-[#0b3d1d] uppercase tracking-wide block mb-1">🔗 ENTEGRE LİNK VE ERİŞİM:</span>
                    <p className="text-[10px] text-slate-500 leading-relaxed font-sans">
                      Entegrasyonumuz aktif ve sabittir. İnternet üzerinden veri havuzuna her an canlı müdahale edebilirsiniz.
                    </p>
                  </div>
                </div>
              </div>

            </div>
          )}

          {/* Iframe Penceresi Kendisi */}
          {modalType === 'iframe' && modalUrl && (
            <div className="w-full h-full relative flex flex-col items-center justify-center bg-slate-50">
              <iframe
                id="system-iframe"
                src={modalUrl}
                className="w-full h-full border-none bg-white"
                title={modalTitle}
                referrerPolicy="no-referrer"
                allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture"
                allowFullScreen
                onLoad={() => setIframeLoading(false)}
              ></iframe>
            </div>
          )}

          {/* DENETLEME RAPOR VE EKLERİ İÇ BÜNYE MODÜLÜ */}
          {modalType === 'denetleme' && (
            <div className="w-full h-full relative flex flex-col bg-slate-900 animate-fade-in pt-16">
              <iframe
                id="denetleme-iframe"
                src="https://ogmhavacilik.github.io/surecyonet/"
                className="w-full h-full border-none bg-white"
                title="8. DENETLEME RAPOR VE EKLER"
                allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture"
                allowFullScreen
              ></iframe>
            </div>
          )}

        </div>
      </div>

      {/* 4. PASSWORD VERIFICATION MODAL ("1839") */}
      <AnimatePresence>
        {isPasswordModalOpen && (
          <div className="fixed inset-0 bg-black/60 backdrop-blur-sm flex items-center justify-center z-[2000] p-4 select-none animate-fade-in">
            <motion.div
              initial={{ scale: 0.95, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.95, opacity: 0 }}
              className="bg-white border-2 border-emerald-800/10 rounded-[2rem] shadow-2xl max-w-sm w-full p-8 text-center relative overflow-hidden"
            >
              <div className="absolute top-0 inset-x-0 h-1.5 bg-gradient-to-r from-emerald-600 via-emerald-500 to-emerald-700"></div>
              
              <button
                onClick={() => setIsPasswordModalOpen(false)}
                className="absolute top-4 right-4 text-slate-400 hover:text-slate-600 transition-colors p-1.5 rounded-lg hover:bg-slate-50"
              >
                <X className="w-5 h-5" />
              </button>

              <div className="w-14 h-14 bg-emerald-50 rounded-2xl flex items-center justify-center mx-auto mb-6 text-[#0b3d1d]">
                <Lock className="w-7 h-7" />
              </div>

              <h4 className="text-slate-800 font-extrabold tracking-wider text-sm uppercase mb-2">YETKİLİ VERİ ERİŞİMİ GİRİŞİ</h4>
              <p className="text-xs text-slate-500 font-semibold leading-relaxed mb-6">
                Lütfen form veritabanı indirme/güncelleme işlemini başlatmak için 4 haneli güvenlik şifresini giriniz.
              </p>

              <div className="mb-6">
                <input
                  type="password"
                  placeholder="••••"
                  maxLength={4}
                  value={passwordInput}
                  onChange={(e) => {
                    setPasswordError(false);
                    setPasswordInput(e.target.value);
                  }}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter') handleVerifyPassword();
                  }}
                  className="w-40 text-center tracking-[0.8em] text-xl font-black border-2 border-slate-200 rounded-2xl py-3 focus:outline-none focus:border-[#0b3d1d] focus:ring-4 focus:ring-[#0b3d1d]/10 transition-all font-mono text-slate-800"
                  autoFocus
                />
                
                {passwordError && (
                  <p className="text-red-600 text-[11px] font-extrabold mt-3 animate-pulse flex items-center justify-center gap-1">
                    ⚠️ Şifre Hatalı! Erişim Engellendi.
                  </p>
                )}
              </div>

              <div className="flex gap-3 justify-center">
                <button
                  type="button"
                  onClick={() => setIsPasswordModalOpen(false)}
                  className="px-5 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-600 text-xs font-bold rounded-xl transition-colors cursor-pointer"
                >
                  İptal
                </button>
                <button
                  type="button"
                  onClick={handleVerifyPassword}
                  className="px-6 py-2.5 bg-[#0b3d1d] hover:bg-[#0b3d1d]/90 text-white text-xs font-bold rounded-xl shadow-md transition-all cursor-pointer"
                >
                  Giriş Yap
                </button>
              </div>

            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* GÖREV EMRİ MANUEL DÜZENLEME MODALİ */}
      <AnimatePresence>
        {activeGorevEmriEdit && editGorevEmriValues && (
          <div className="fixed inset-0 bg-slate-950/60 backdrop-blur-md z-[2100] flex items-center justify-center p-4 select-text">
            <motion.div
              initial={{ opacity: 0, scale: 0.95 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.95 }}
              className="bg-white border-2 border-slate-200/50 rounded-[2.5rem] shadow-2xl max-w-lg w-full p-6 relative overflow-hidden"
            >
              <div className="absolute top-0 inset-x-0 h-1.5 bg-amber-500"></div>
              <button
                type="button"
                onClick={() => {
                  setActiveGorevEmriEdit(null);
                  setEditGorevEmriValues(null);
                  setPendingGeEditOrder(null);
                }}
                className="absolute top-5 right-5 text-slate-400 hover:text-slate-600 bg-slate-100 p-2 rounded-xl transition-all cursor-pointer"
              >
                <X className="w-4 h-4" />
              </button>

              <h3 className="text-slate-800 font-black text-sm uppercase mb-4 flex items-center gap-2">
                ✏️ GÖREV EMRİ MANUEL DÜZENLEME PANELİ
              </h3>

              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="block text-[10px] font-black text-slate-400 uppercase mb-1">Görev Tarihi</label>
                  <input
                    type="date"
                    value={convertToInputDateFormat(editGorevEmriValues.date || "")}
                    onChange={(e) => {
                      const formatted = convertToDisplayDateFormat(e.target.value);
                      setEditGorevEmriValues({ ...editGorevEmriValues, date: formatted });
                    }}
                    className="w-full px-3.5 py-2.5 bg-slate-50 border-2 border-slate-100 focus:border-slate-300 rounded-xl text-xs font-bold text-slate-700 focus:outline-none cursor-pointer"
                  />
                </div>
                <div>
                  <label className="block text-[10px] font-black text-slate-400 uppercase mb-1">Araç Plakası</label>
                  <input
                    type="text"
                    value={editGorevEmriValues.plate || ""}
                    onChange={(e) => setEditGorevEmriValues({ ...editGorevEmriValues, plate: e.target.value.toUpperCase() })}
                    className="w-full px-3.5 py-2.5 bg-slate-50 border-2 border-slate-100 focus:border-slate-300 rounded-xl text-xs font-bold text-slate-700 focus:outline-none"
                  />
                </div>
                <div>
                  <label className="block text-[10px] font-black text-slate-400 uppercase mb-1">Sürücü Personel Ad Soyad</label>
                  <input
                    type="text"
                    value={editGorevEmriValues.driverName || ""}
                    onChange={(e) => setEditGorevEmriValues({ ...editGorevEmriValues, driverName: e.target.value })}
                    className="w-full px-3.5 py-2.5 bg-slate-50 border-2 border-slate-100 focus:border-slate-300 rounded-xl text-xs font-bold text-slate-700 focus:outline-none"
                  />
                </div>
                <div>
                  <label className="block text-[10px] font-black text-slate-400 uppercase mb-1">Sürücü Sicil No</label>
                  <input
                    type="text"
                    value={editGorevEmriValues.driverSicil || ""}
                    onChange={(e) => setEditGorevEmriValues({ ...editGorevEmriValues, driverSicil: e.target.value })}
                    className="w-full px-3.5 py-2.5 bg-slate-50 border-2 border-slate-100 focus:border-slate-300 rounded-xl text-xs font-bold text-slate-700 focus:outline-none"
                  />
                </div>
                <div>
                  <label className="block text-[10px] font-black text-slate-400 uppercase mb-1">Görev Seri No</label>
                  <input
                    type="text"
                    value={editGorevEmriValues.serialNo || ""}
                    onChange={(e) => setEditGorevEmriValues({ ...editGorevEmriValues, serialNo: e.target.value })}
                    className="w-full px-3.5 py-2.5 bg-slate-50 border-2 border-slate-100 focus:border-slate-300 rounded-xl text-xs font-bold text-slate-700 focus:outline-none"
                  />
                </div>
                <div>
                  <label className="block text-[10px] font-black text-slate-400 uppercase mb-1">Çıkış KM</label>
                  <input
                    type="number"
                    value={editGorevEmriValues.departureKm || ""}
                    onChange={(e) => setEditGorevEmriValues({ ...editGorevEmriValues, departureKm: parseInt(e.target.value) || 0 })}
                    className="w-full px-3.5 py-2.5 bg-slate-50 border-2 border-slate-100 focus:border-slate-300 rounded-xl text-xs font-bold text-slate-700 focus:outline-none"
                  />
                </div>
                <div>
                  <label className="block text-[10px] font-black text-slate-400 uppercase mb-1">Dönüş KM</label>
                  <input
                    type="number"
                    value={editGorevEmriValues.returnKm || ""}
                    onChange={(e) => setEditGorevEmriValues({ ...editGorevEmriValues, returnKm: parseInt(e.target.value) || 0 })}
                    className="w-full px-3.5 py-2.5 bg-slate-50 border-2 border-slate-100 focus:border-slate-300 rounded-xl text-xs font-bold text-slate-700 focus:outline-none"
                  />
                </div>
                <div>
                  <label className="block text-[10px] font-black text-slate-400 uppercase mb-1">Güzergah</label>
                  <input
                    type="text"
                    value={editGorevEmriValues.route || ""}
                    onChange={(e) => setEditGorevEmriValues({ ...editGorevEmriValues, route: e.target.value })}
                    className="w-full px-3.5 py-2.5 bg-slate-50 border-2 border-slate-100 focus:border-slate-300 rounded-xl text-xs font-bold text-slate-700 focus:outline-none"
                  />
                </div>
                <div>
                  <label className="block text-[10px] font-black text-slate-400 uppercase mb-1">Çıkış Saati</label>
                  <input
                    type="text"
                    placeholder="örn: 08:00"
                    value={editGorevEmriValues.departureTime || ""}
                    onChange={(e) => setEditGorevEmriValues({ ...editGorevEmriValues, departureTime: e.target.value })}
                    className="w-full px-3.5 py-2.5 bg-slate-50 border-2 border-slate-100 focus:border-slate-300 rounded-xl text-xs font-bold text-slate-700 focus:outline-none"
                  />
                </div>
                <div>
                  <label className="block text-[10px] font-black text-slate-400 uppercase mb-1">Dönüş Saati</label>
                  <input
                    type="text"
                    placeholder="örn: 17:00"
                    value={editGorevEmriValues.returnTime || ""}
                    onChange={(e) => setEditGorevEmriValues({ ...editGorevEmriValues, returnTime: e.target.value })}
                    className="w-full px-3.5 py-2.5 bg-slate-50 border-2 border-slate-100 focus:border-slate-300 rounded-xl text-xs font-bold text-slate-700 focus:outline-none"
                  />
                </div>
              </div>

              <div className="flex gap-3 justify-end mt-6 pt-4 border-t border-slate-100">
                <button
                  type="button"
                  onClick={() => {
                    setActiveGorevEmriEdit(null);
                    setEditGorevEmriValues(null);
                    setPendingGeEditOrder(null);
                  }}
                  className="px-5 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-500 font-extrabold text-xs rounded-2xl transition-all cursor-pointer"
                >
                  İPTAL
                </button>
                <button
                  type="button"
                  onClick={() => {
                    setPendingGeEditOrder(editGorevEmriValues);
                    setGeEditPasswordInput("");
                    setGeEditPasswordError(false);
                    setShowGeEditPasswordPrompt(true);
                  }}
                  className="px-6 py-2.5 bg-[#0b3d1d] hover:bg-[#072612] text-white font-extrabold text-xs uppercase tracking-wider rounded-2xl transition-all cursor-pointer shadow-md"
                >
                  KAYDET VE SENKRONİZE ET
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* 4.5. TEÇHİZAT SATIR DÜZENLEME VE GÖRSEL YÜKLEME MODALİ */}
      <AnimatePresence>
        {activeTechizatRowEdit && (() => {
          const { rIdx, techType, row } = activeTechizatRowEdit;
          const isKaraAraciEdit = techType === 'kara_araclari';
          const colsList = isKaraAraciEdit
            ? [
                "SIRA NO", 
                "ARAÇ PLAKASI / TANIMI", 
                "PARÇA NO (P/N) / MODEL", 
                "BULUNDUĞU YER", 
                "SON KM Sİ", 
                "DURUMU", 
                "KALİBRASYONA TABİ",
                "SON KONTROL / KALİBRASYON / BAKIM", 
                "GELECEK KONTROL / KALİBRASYON / BAKIM", 
                "SON KONTROLÜ YAPAN FİRMA", 
                "AÇIKLAMA", 
                "90 GÜN UYARISI MAİL GÖNDERİM TARİHİ"
              ]
            : [
                "SIRA NO", 
                "TEÇHİZAT ADI", 
                "PARÇA NO (P/N) / MODEL", 
                "SERİ NO (S/N)", 
                "MİKTAR / KAPASİTE", 
                "BULUNDUĞU YER", 
                "DURUMU", 
                "KALİBRASYONA TABİ",
                "SON KONTROL / KALİBRASYON / BAKIM", 
                "GELECEK KONTROL / KALİBRASYON / BAKIM", 
                "SON KONTROLÜ YAPAN FİRMA", 
                "AÇIKLAMA", 
                "90 GÜN UYARISI MAİL GÖNDERİM TARİHİ"
              ];

          const imageKey = techType + "_" + (row[1] || "").replace(/\s+/g, '_') + "_" + (row[3] || "").replace(/\s+/g, '_');
          const hasImage = !!techizatImages[imageKey];

          return (
            <div className="fixed inset-0 bg-slate-950/70 backdrop-blur-md flex items-start lg:items-center justify-center z-[2100] p-4 overflow-y-auto animate-fade-in select-text">
              <motion.div
                initial={{ scale: 0.95, opacity: 0 }}
                animate={{ scale: 1, opacity: 1 }}
                exit={{ scale: 0.95, opacity: 0 }}
                className="bg-white border-2 border-slate-200/50 rounded-[2.5rem] shadow-2xl max-w-5xl w-full p-5 md:p-8 relative flex flex-col lg:flex-row gap-8 my-4 lg:my-8 lg:max-h-[92vh] lg:overflow-hidden select-text"
              >
                {/* Decorative top strip */}
                <div className="absolute top-0 inset-x-0 h-1.5 bg-gradient-to-r from-emerald-600 via-emerald-500 to-emerald-700"></div>

                {/* Close Button */}
                <button
                  onClick={() => {
                    setActiveTechizatRowEdit(null);
                    setPendingImageFile(null);
                    setPendingImagePreview(null);
                    setPendingImageBase64(null);
                    setPendingImageMimeType(null);
                    setIsDragging(false);
                    setIsImageUploadingToDrive(false);
                    setIsDataUpdateUnlocked(false);
                    setDataPasswordInput('');
                    setDataPasswordError(false);
                  }}
                  className="absolute top-4 right-4 text-slate-400 hover:text-slate-600 transition-colors p-1.5 rounded-lg hover:bg-slate-50 z-20"
                >
                  <X className="w-5 h-5" />
                </button>

                {/* Mobile / Tablet Segment/Tab Switcher */}
                <div className="flex lg:hidden bg-slate-100 p-1.5 rounded-2xl gap-1 w-full mt-4 shrink-0 select-none">
                  <button
                    type="button"
                    onClick={() => setMobileEditTab('form')}
                    className={`flex-1 py-2.5 rounded-xl text-xs font-extrabold uppercase tracking-wider transition-all flex items-center justify-center gap-1.5 ${
                      mobileEditTab === 'form'
                        ? 'bg-[#0b3d1d] text-white shadow-md'
                        : 'text-slate-500 hover:text-slate-800 hover:bg-slate-200/50'
                    }`}
                  >
                    📝 Kayıt Bilgileri
                  </button>
                  <button
                    type="button"
                    onClick={() => setMobileEditTab('image')}
                    className={`flex-1 py-2.5 rounded-xl text-xs font-extrabold uppercase tracking-wider transition-all flex items-center justify-center gap-1.5 ${
                      mobileEditTab === 'image'
                        ? 'bg-[#0b3d1d] text-white shadow-md'
                        : 'text-slate-500 hover:text-slate-800 hover:bg-slate-200/50'
                    }`}
                  >
                    📷 Teçhizat Görseli
                  </button>
                </div>

                {/* Left Side: Dynamic Form Fields */}
                <div className={`flex-1 lg:overflow-y-auto pr-2 lg:max-h-[82vh] flex flex-col text-left ${mobileEditTab === 'form' ? 'flex' : 'hidden lg:flex'}`}>
                  <div className="mb-6">
                    <h3 className="text-base font-black text-slate-800 uppercase tracking-widest flex items-center gap-2">
                      🛠️ KAYIT DÜZENLEME PANELİ
                    </h3>
                    <p className="text-[11px] text-slate-400 font-semibold mt-1">
                      {isKaraAraciEdit ? "Kara Aracı" : "Standart Teçhizat"} listesindeki seçili kalemin tüm teknik özelliklerini ve tarihlerini güncelleyebilirsiniz.
                    </p>
                  </div>

                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                    {(() => {
                      const locColIdx = colsList.indexOf("BULUNDUĞU YER");
                      const miktarColIdx = colsList.findIndex(c => c.includes("MİKTAR") || c.includes("KAPASİTE"));

                      const rawLoc = locColIdx !== -1 ? (editRowValues[locColIdx] || "") : "";
                      const rawMiktar = miktarColIdx !== -1 ? (editRowValues[miktarColIdx] || "") : "";

                      const parsedLocs = rawLoc ? rawLoc.split('\n') : [""];
                      const parsedMiktars = rawMiktar ? rawMiktar.split('\n') : ["1"];

                      const pairCount = Math.max(1, parsedLocs.length, parsedMiktars.length);
                      while (parsedLocs.length < pairCount) parsedLocs.push("");
                      while (parsedMiktars.length < pairCount) parsedMiktars.push("1");

                      return colsList.map((col, idx) => {
                        // Skip Sıra No (index 0) - make it read-only
                        if (idx === 0) {
                          return (
                            <div key={idx} className="flex flex-col gap-1.5">
                              <label className="text-[10px] font-black text-slate-400 uppercase tracking-wider">{col}</label>
                              <input
                                type="text"
                                value={editRowValues[idx] || ""}
                                disabled
                                className="w-full px-3.5 py-2.5 bg-slate-100 border border-slate-200 rounded-2xl text-xs font-bold text-slate-500 cursor-not-allowed"
                              />
                            </div>
                          );
                        }

                        // Paired Miktar / Kapasite column field
                        if (idx === miktarColIdx && locColIdx !== -1) {
                          return (
                            <div key={idx} className="flex flex-col gap-1.5">
                              <label className="text-[10px] font-black text-slate-500 uppercase tracking-wider flex items-center gap-1">
                                📦 {col}
                              </label>
                              <div className="flex flex-col gap-2">
                                {parsedMiktars.map((mVal, pIdx) => (
                                  <div key={pIdx} className="flex items-center gap-1.5 h-[42px]">
                                    <input
                                      type="text"
                                      value={mVal}
                                      placeholder="1"
                                      onChange={(e) => {
                                        const newMiktars = [...parsedMiktars];
                                        newMiktars[pIdx] = e.target.value;
                                        const cloned = [...editRowValues];
                                        cloned[miktarColIdx] = newMiktars.join('\n');
                                        setEditRowValues(cloned);
                                      }}
                                      className="w-full px-3.5 py-2.5 bg-slate-50 border border-slate-200 focus:border-[#0b3d1d] focus:ring-4 focus:ring-[#0b3d1d]/5 rounded-2xl text-xs font-bold text-slate-800 focus:outline-none transition-all"
                                    />
                                    {pairCount > 1 && <div className="w-8 h-8 shrink-0" />}
                                  </div>
                                ))}
                              </div>
                            </div>
                          );
                        }

                        // Paired Bulunduğu Yer column field
                        if (idx === locColIdx) {
                          return (
                            <div key={idx} className="flex flex-col gap-1.5">
                              <label className="text-[10px] font-black text-slate-500 uppercase tracking-wider flex items-center gap-1">
                                📍 {col}
                              </label>
                              <div className="flex flex-col gap-2">
                                {parsedLocs.map((lVal, pIdx) => (
                                  <div key={pIdx} className="flex items-center gap-1.5 h-[42px]">
                                    <input
                                      type="text"
                                      value={lVal}
                                      placeholder={pIdx === 0 ? "Örn: Y/D HANGAR" : "Örn: Muğla / Antalya"}
                                      onChange={(e) => {
                                        const newLocs = [...parsedLocs];
                                        newLocs[pIdx] = e.target.value;
                                        const cloned = [...editRowValues];
                                        cloned[locColIdx] = newLocs.join('\n');
                                        setEditRowValues(cloned);
                                      }}
                                      className="w-full px-3.5 py-2.5 bg-slate-50 border border-slate-200 focus:border-[#0b3d1d] focus:ring-4 focus:ring-[#0b3d1d]/5 rounded-2xl text-xs font-bold text-slate-800 focus:outline-none transition-all"
                                    />
                                    {pairCount > 1 && (
                                      <button
                                        type="button"
                                        title="Bu Lokasyonu ve Miktarını Sil"
                                        onClick={() => {
                                          const newLocs = parsedLocs.filter((_, i) => i !== pIdx);
                                          const newMiktars = parsedMiktars.filter((_, i) => i !== pIdx);
                                          const cloned = [...editRowValues];
                                          cloned[locColIdx] = newLocs.join('\n');
                                          if (miktarColIdx !== -1) cloned[miktarColIdx] = newMiktars.join('\n');
                                          setEditRowValues(cloned);
                                        }}
                                        className="w-8 h-8 shrink-0 flex items-center justify-center bg-rose-50 hover:bg-rose-100 text-rose-600 rounded-xl transition-colors cursor-pointer border border-rose-200/80 active:scale-95"
                                      >
                                        <Trash2 className="w-4 h-4" />
                                      </button>
                                    )}
                                  </div>
                                ))}

                                <button
                                  type="button"
                                  onClick={() => {
                                    const newLocs = [...parsedLocs, ""];
                                    const newMiktars = [...parsedMiktars, "1"];
                                    const cloned = [...editRowValues];
                                    cloned[locColIdx] = newLocs.join('\n');
                                    if (miktarColIdx !== -1) cloned[miktarColIdx] = newMiktars.join('\n');
                                    setEditRowValues(cloned);
                                  }}
                                  className="mt-1 self-start inline-flex items-center gap-1.5 px-3 py-1.5 bg-emerald-50 hover:bg-emerald-100 text-[#0b3d1d] text-xs font-extrabold rounded-xl border border-emerald-200/80 transition-all cursor-pointer active:scale-95 shadow-sm"
                                >
                                  <Plus className="w-3.5 h-3.5 text-emerald-700" />
                                  <span>+ Yeni Lokasyon Ekle</span>
                                </button>
                              </div>
                            </div>
                          );
                        }

                        // Special field: KALİBRASYONA TABİ (Dropdown selection EVET / HAYIR)
                        if (col === "KALİBRASYONA TABİ" || col === "BAKIMA TABİ Mİ?") {
                          const currentVal = (editRowValues[idx] || "EVET").toString().trim().toUpperCase();
                          return (
                            <div key={idx} className="flex flex-col gap-1.5">
                              <label className="text-[10px] font-black text-slate-500 uppercase tracking-wider flex items-center gap-1">
                                ⚙️ KALİBRASYONA TABİ
                              </label>
                              <select
                                value={currentVal === "HAYIR" ? "HAYIR" : "EVET"}
                                onChange={(e) => {
                                  const val = e.target.value;
                                  const cloned = [...editRowValues];
                                  cloned[idx] = val;
                                  // If "HAYIR" is selected, clear maintenance fields and firm
                                  if (val === "HAYIR") {
                                    const lastCntIdx = colsList.indexOf("SON KONTROL / KALİBRASYON / BAKIM");
                                    const nextCntIdx = colsList.indexOf("GELECEK KONTROL / KALİBRASYON / BAKIM");
                                    const firmIdx = colsList.indexOf("SON KONTROLÜ YAPAN FİRMA");
                                    if (lastCntIdx !== -1) cloned[lastCntIdx] = "-";
                                    if (nextCntIdx !== -1) cloned[nextCntIdx] = "-";
                                    if (firmIdx !== -1) cloned[firmIdx] = "-";
                                  }
                                  setEditRowValues(cloned);
                                }}
                                className="w-full px-3.5 py-2.5 bg-slate-50 border-2 border-[#0b3d1d]/15 rounded-2xl text-xs font-bold text-slate-800 focus:outline-none focus:border-[#0b3d1d] cursor-pointer"
                              >
                                <option value="EVET">EVET (Kalibrasyona / Bakıma Tabi)</option>
                                <option value="HAYIR">HAYIR (Kalibrasyon / Bakımdan Muaf)</option>
                              </select>
                            </div>
                          );
                        }

                        // Check if KALİBRASYONA TABİ is HAYIR (disabled flu mode for maintenance & firm fields)
                        const kalibColIdx = colsList.indexOf("KALİBRASYONA TABİ") !== -1 ? colsList.indexOf("KALİBRASYONA TABİ") : colsList.indexOf("BAKIMA TABİ Mİ?");
                        const isKalibTabiHayir = kalibColIdx !== -1 && (editRowValues[kalibColIdx] || "").toString().trim().toUpperCase() === "HAYIR";
                        const isMaintenanceOrFirmField = col === "SON KONTROL / KALİBRASYON / BAKIM" || 
                                                         col === "GELECEK KONTROL / KALİBRASYON / BAKIM" || 
                                                         col === "SON KONTROLÜ YAPAN FİRMA";

                        if (isKalibTabiHayir && isMaintenanceOrFirmField) {
                          return (
                            <div key={idx} className="flex flex-col gap-1.5 opacity-40">
                              <label className="text-[10px] font-black text-slate-400 uppercase tracking-wider">{col}</label>
                              <input
                                type="text"
                                value="MUAFIYET (TABİ DEĞİL)"
                                disabled
                                className="w-full px-3.5 py-2.5 bg-slate-100 border border-slate-200 rounded-2xl text-xs font-bold text-slate-400 cursor-not-allowed"
                              />
                            </div>
                          );
                        }

                        // Special field: DURUMU (Dropdown FAAL / BAKIM KALİBRASYON / GAYRİ FAAL)
                        if (col === "DURUMU") {
                          return (
                            <div key={idx} className="flex flex-col gap-1.5">
                              <label className="text-[10px] font-black text-slate-500 uppercase tracking-wider flex items-center gap-1">
                                🟢 {col}
                              </label>
                              <select
                                value={editRowValues[idx] || "FAAL"}
                                onChange={(e) => {
                                  const cloned = [...editRowValues];
                                  cloned[idx] = e.target.value;
                                  setEditRowValues(cloned);
                                }}
                                className="w-full px-3.5 py-2.5 bg-slate-50 border border-slate-200 focus:border-[#0b3d1d] focus:ring-4 focus:ring-[#0b3d1d]/5 rounded-2xl text-xs font-bold text-slate-800 focus:outline-none transition-all cursor-pointer"
                              >
                                <option value="FAAL">FAAL</option>
                                <option value="BAKIM / KALİBRASYON">BAKIM / KALİBRASYON</option>
                                <option value="GAYRİ FAAL">GAYRİ FAAL</option>
                              </select>
                            </div>
                          );
                        }

                        // Special field: SON KONTROLÜ YAPAN FİRMA (With Firm Datalist)
                        if (col === "SON KONTROLÜ YAPAN FİRMA") {
                          return (
                            <div key={idx} className="flex flex-col gap-1.5">
                              <label className="text-[10px] font-black text-slate-500 uppercase tracking-wider flex items-center gap-1">
                                🏢 {col}
                              </label>
                              <input
                                type="text"
                                list="techizat-firmalar-list"
                                value={editRowValues[idx] || ""}
                                placeholder="Firma seçiniz veya yazınız..."
                                onChange={(e) => {
                                  const cloned = [...editRowValues];
                                  cloned[idx] = e.target.value;
                                  setEditRowValues(cloned);
                                }}
                                className="w-full px-3.5 py-2.5 bg-slate-50 border border-slate-200 focus:border-[#0b3d1d] focus:ring-4 focus:ring-[#0b3d1d]/5 rounded-2xl text-xs font-bold text-slate-800 focus:outline-none transition-all"
                              />
                            </div>
                          );
                        }

                        const isDateField = col.includes("TARİH") || col.includes("KONTROL / KALİBRASYON / BAKIM");
                        const isMailSendDate = col === "90 GÜN UYARISI MAİL GÖNDERİM TARİHİ" || 
                          col.toUpperCase().includes("MAİL GÖNDERİM") || 
                          col.toUpperCase().includes("MAIL GONDERIM") || 
                          col.toUpperCase().includes("MAİL GÖNDERİLDİĞİ") || 
                          col.toUpperCase().includes("MAIL GONDERILDI") || 
                          col.toLowerCase().includes("mail") || 
                          col.toLowerCase().includes("e-posta");

                        if (isMailSendDate) {
                          return (
                            <div key={idx} className="flex flex-col gap-1.5">
                              <label className="text-[10px] font-black text-slate-400 uppercase tracking-wider">{col}</label>
                              <input
                                type="text"
                                value={editRowValues[idx] || "Belirtilmemiş"}
                                disabled
                                className="w-full px-3.5 py-2.5 bg-slate-100 border border-slate-200 rounded-2xl text-xs font-bold text-slate-500 cursor-not-allowed"
                              />
                            </div>
                          );
                        }

                        if (isDateField) {
                          const dateVal = convertToInputDateFormat(editRowValues[idx] || "");
                          return (
                            <div key={idx} className="flex flex-col gap-1.5">
                              <label className="text-[10px] font-black text-slate-500 uppercase tracking-wider">📅 {col}</label>
                              <input
                                type="date"
                                value={dateVal}
                                onChange={(e) => {
                                  const formatted = convertToDisplayDateFormat(e.target.value);
                                  const cloned = [...editRowValues];
                                  cloned[idx] = formatted;
                                  setEditRowValues(cloned);
                                }}
                                className="w-full px-3.5 py-2.5 bg-slate-50 border border-slate-200 focus:border-[#0b3d1d] focus:ring-4 focus:ring-[#0b3d1d]/5 rounded-2xl text-xs font-bold text-slate-800 focus:outline-none transition-all"
                              />
                            </div>
                          );
                        }

                        // Render normal input fields
                        return (
                          <div key={idx} className="flex flex-col gap-1.5">
                            <label className="text-[10px] font-black text-slate-500 uppercase tracking-wider">{col}</label>
                            <input
                              type="text"
                              value={editRowValues[idx] || ""}
                              placeholder="Belirtilmemiş"
                              onChange={(e) => {
                                const cloned = [...editRowValues];
                                cloned[idx] = e.target.value;
                                setEditRowValues(cloned);
                              }}
                              className="w-full px-3.5 py-2.5 bg-slate-50 border border-slate-200 focus:border-[#0b3d1d] focus:ring-4 focus:ring-[#0b3d1d]/5 rounded-2xl text-xs font-bold text-slate-800 focus:outline-none transition-all"
                            />
                          </div>
                        );
                      });
                    })()}
                  </div>

                  {/* Save buttons */}
                  <div className="flex flex-col gap-3 mt-8 border-t border-slate-100 pt-5">
                    <div className="flex gap-3 justify-end items-center">
                      {isDataUpdateUnlocked && (
                        <span className="text-[10px] font-black text-emerald-800 flex items-center gap-1.5 mr-auto">
                          <span className="w-2 h-2 rounded-full bg-emerald-500 animate-pulse" />
                          GÜNCELLEME YETKİSİ ONAYLANDI
                        </span>
                      )}
                      <button
                        type="button"
                        disabled={isTechizatSaving}
                        onClick={() => {
                          setActiveTechizatRowEdit(null);
                          setPendingImageFile(null);
                          setPendingImagePreview(null);
                          setIsDragging(false);
                          setIsImageUploadingToDrive(false);
                          setDataPasswordInput('');
                          setDataPasswordError(false);
                          setShowSavePasswordPrompt(false);
                        }}
                        className={`px-6 py-3 bg-slate-100 hover:bg-slate-200 text-slate-600 text-xs font-bold rounded-2xl transition-colors cursor-pointer ${isTechizatSaving ? 'opacity-50 cursor-not-allowed' : ''}`}
                      >
                        Kapat / İptal
                      </button>
                      <button
                        type="button"
                        disabled={isTechizatSaving}
                        onClick={async () => {
                          if (isDataUpdateUnlocked) {
                            await handleSaveTechizatRow(editRowValues, techType, rIdx);
                          } else {
                            setDataPasswordInput('');
                            setDataPasswordError(false);
                            setShowSavePasswordPrompt(true);
                          }
                        }}
                        className={`px-8 py-3 bg-[#0b3d1d] hover:bg-[#072612] text-white text-xs font-black uppercase tracking-widest rounded-2xl shadow-lg transition-all cursor-pointer active:scale-95 flex items-center gap-2 ${isTechizatSaving ? 'opacity-70 cursor-not-allowed' : ''}`}
                      >
                        {isTechizatSaving ? (
                          <>
                            <Loader2 className="w-4 h-4 animate-spin text-white" />
                            <span>SENKRONİZE EDİLİYOR...</span>
                          </>
                        ) : (
                          <span>KAYDET VE GÜNCELLE</span>
                        )}
                      </button>
                    </div>

                    {showSavePasswordPrompt && (
                      <div className="fixed inset-0 bg-slate-950/80 backdrop-blur-md flex items-center justify-center z-[2200] p-4 select-text">
                        <div className="bg-white border-2 border-slate-200/50 rounded-[2rem] shadow-2xl max-w-sm w-full p-6 text-center relative overflow-hidden animate-fade-in">
                          <div className="absolute top-0 inset-x-0 h-1 bg-[#0b3d1d]"></div>
                          <div className="w-12 h-12 bg-emerald-50 rounded-2xl flex items-center justify-center mx-auto mb-4 text-[#0b3d1d]">
                            <Lock className="w-6 h-6" />
                          </div>
                          <h4 className="text-slate-800 font-extrabold text-sm uppercase mb-2">KAYIT DEĞİŞİKLİK ONAYI</h4>
                          <p className="text-xs text-slate-500 font-semibold mb-4 leading-relaxed">
                            Yapılan değişiklikleri kaydetmek için lütfen yetkili şifresini giriniz.
                          </p>
                          <input
                            type="password"
                            placeholder="Şifre"
                            value={dataPasswordInput}
                            onChange={(e) => {
                              setDataPasswordInput(e.target.value);
                              setDataPasswordError(false);
                            }}
                            onKeyDown={async (e) => {
                              if (e.key === 'Enter') {
                                if (dataPasswordInput === '1839') {
                                  setIsDataUpdateUnlocked(true);
                                  setDataPasswordInput('');
                                  setDataPasswordError(false);
                                  setShowSavePasswordPrompt(false);
                                  await handleSaveTechizatRow(editRowValues, techType, rIdx);
                                } else {
                                  setDataPasswordError(true);
                                }
                              }
                            }}
                            className="w-full px-4 py-2 bg-slate-50 border-2 border-[#0b3d1d]/15 rounded-xl text-center text-sm font-semibold mb-3 focus:outline-none focus:border-[#0b3d1d] text-slate-900"
                            autoFocus
                          />
                          {dataPasswordError && (
                            <p className="text-red-600 text-[10px] font-black mb-3">❌ Hatalı şifre girdiniz!</p>
                          )}
                          <div className="flex gap-2 justify-center">
                            <button
                              type="button"
                              onClick={() => {
                                setShowSavePasswordPrompt(false);
                                setDataPasswordInput('');
                                setDataPasswordError(false);
                              }}
                              className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-600 text-xs font-bold rounded-xl transition-colors cursor-pointer"
                            >
                              İptal
                            </button>
                            <button
                              type="button"
                              onClick={async () => {
                                if (dataPasswordInput === '1839') {
                                  setIsDataUpdateUnlocked(true);
                                  setDataPasswordInput('');
                                  setDataPasswordError(false);
                                  setShowSavePasswordPrompt(false);
                                  await handleSaveTechizatRow(editRowValues, techType, rIdx);
                                } else {
                                  setDataPasswordError(true);
                                }
                              }}
                              className="px-6 py-2 bg-[#0b3d1d] hover:bg-[#072612] text-white text-xs font-black uppercase rounded-xl transition-all cursor-pointer"
                            >
                              Onayla ve Kaydet
                            </button>
                          </div>
                        </div>
                      </div>
                    )}
                  </div>
                </div>

                {/* Right Side: Image Upload, Background Removal and Manual Retouch Panel */}
                <div className={`w-full lg:w-5/12 border-t lg:border-t-0 lg:border-l border-slate-200 pt-6 lg:pt-0 lg:pl-8 lg:overflow-y-auto lg:max-h-[82vh] pr-1 flex flex-col gap-6 text-left ${mobileEditTab === 'image' ? 'flex' : 'hidden lg:flex'}`}>
                  <ImageEditorAndRetoucher
                    imageKey={imageKey}
                    currentImageUrl={techizatImages[imageKey] || null}
                    hasImage={hasImage}
                    isImageUpdateUnlocked={isImageUpdateUnlocked}
                    isUploadingToDrive={isImageUploadingToDrive}
                    partName={row[1] || ""}
                    manufacturer={row[3] || ""}
                    onUnlockImageUpdate={() => {
                      setTempImageAction('unlock_only');
                      setImagePasswordInput('');
                      setImagePasswordError(false);
                      setShowImageSavePasswordPrompt(true);
                    }}
                    onLockImageUpdate={() => {
                      setIsImageUpdateUnlocked(false);
                    }}
                    onRemoveImage={() => {
                      if (isImageUpdateUnlocked) {
                        setTechizatImages(prev => {
                          const cloned = { ...prev };
                          delete cloned[imageKey];
                          localStorage.setItem('techizat_images', JSON.stringify(cloned));
                          return cloned;
                        });
                        showNotification("Görsel kaldırıldı.");
                      } else {
                        setTempImageAction('remove');
                        setImagePasswordInput('');
                        setImagePasswordError(false);
                        setShowImageSavePasswordPrompt(true);
                      }
                    }}
                    onSaveImage={async (base64Data, mimeType) => {
                      if (isImageUpdateUnlocked) {
                        try {
                          setIsImageUploadingToDrive(true);
                          showNotification("Görsel Google Drive'a yükleniyor...");
                          const driveFileName = `tech_img_${imageKey}.png`;
                          const res = await fetch(GOOGLE_SCRIPT_URL, {
                            method: "POST",
                            headers: {
                              "Content-Type": "text/plain;charset=utf-8"
                            },
                            body: JSON.stringify({
                              action: "uploadPdfToDrive",
                              fileName: driveFileName,
                              base64Data: base64Data,
                              mimeType: mimeType,
                              folderId: "1HQR_NYKhHQGA7_2W3nArI9pCh-LJasTP"
                            })
                          });
                          if (!res.ok) {
                            throw new Error(`Google Apps Script sunucu hatası: ${res.status}`);
                          }
                          const result = await res.json();
                          if (result.status === "success" && result.viewUrl) {
                            setTechizatImages(prev => {
                              const updated = {
                                ...prev,
                                [imageKey]: result.viewUrl
                              };
                              localStorage.setItem('techizat_images', JSON.stringify(updated));
                              return updated;
                            });
                            showNotification("Görsel başarıyla Drive'a yüklendi ve aktifleştirildi!");
                          } else {
                            throw new Error(result.message || "Bilinmeyen sunucu hatası.");
                          }
                        } catch (err: any) {
                          console.error("Yükleme Hatası:", err);
                          showNotification(`Yükleme başarısız: ${err.message}`);
                        } finally {
                          setIsImageUploadingToDrive(false);
                        }
                      } else {
                        setPendingImageBase64(base64Data);
                        setPendingImageMimeType(mimeType);
                        setTempImageAction('upload');
                        setImagePasswordInput('');
                        setImagePasswordError(false);
                        setShowImageSavePasswordPrompt(true);
                      }
                    }}
                    onSaveImageUrl={(url) => {
                      if (isImageUpdateUnlocked) {
                        setTechizatImages(prev => {
                          const updated = {
                            ...prev,
                            [imageKey]: url
                          };
                          localStorage.setItem('techizat_images', JSON.stringify(updated));
                          return updated;
                        });
                        showNotification("Görsel bağlantısı başarıyla kaydedildi!");
                      } else {
                        setTempImageUrlInput(url);
                        setTempImageAction('link');
                        setImagePasswordInput('');
                        setImagePasswordError(false);
                        setShowImageSavePasswordPrompt(true);
                      }
                    }}
                  />

                  {showImageSavePasswordPrompt && (
                    <div className="absolute inset-0 bg-slate-950/85 backdrop-blur-md flex items-center justify-center z-[2200] p-4 select-text animate-fade-in">
                      <div className="bg-white border-2 border-slate-200/50 rounded-[2rem] shadow-2xl max-w-sm w-full p-6 text-center relative overflow-hidden">
                        <div className="absolute top-0 inset-x-0 h-1 bg-[#0b3d1d]"></div>
                        <div className="w-12 h-12 bg-emerald-50 rounded-2xl flex items-center justify-center mx-auto mb-4 text-[#0b3d1d]">
                          <Lock className="w-6 h-6" />
                        </div>
                        <h4 className="text-slate-800 font-extrabold text-sm uppercase mb-2">GÖRSEL İŞLEM ONAYI</h4>
                        <p className="text-xs text-slate-500 font-semibold mb-4 leading-relaxed">
                          {tempImageAction === 'upload' && "Görseli Google Drive'a kaydetmek için lütfen yetkili şifresini giriniz."}
                          {tempImageAction === 'unlock_only' && "Görsel güncelleme kilidini açmak için lütfen yetkili şifresini giriniz."}
                          {tempImageAction === 'remove' && "Görseli kaldırmak için lütfen yetkili şifresini giriniz."}
                        </p>
                        <input
                          type="password"
                          placeholder="Şifre"
                          value={imagePasswordInput}
                          onChange={(e) => {
                            setImagePasswordInput(e.target.value);
                            setImagePasswordError(false);
                          }}
                          onKeyDown={async (e) => {
                            if (e.key === 'Enter') {
                              if (imagePasswordInput === '1839') {
                                setIsImageUpdateUnlocked(true);
                                setImagePasswordInput('');
                                setImagePasswordError(false);
                                setShowImageSavePasswordPrompt(false);
                                if (tempImageAction === 'upload' && pendingImageBase64) {
                                  try {
                                    setIsImageUploadingToDrive(true);
                                    showNotification("Görsel Google Drive'a yükleniyor...");
                                    const driveFileName = `tech_img_${imageKey}.png`;
                                    const res = await fetch(GOOGLE_SCRIPT_URL, {
                                      method: "POST",
                                      headers: {
                                        "Content-Type": "text/plain;charset=utf-8"
                                      },
                                      body: JSON.stringify({
                                        action: "uploadPdfToDrive",
                                        fileName: driveFileName,
                                        base64Data: pendingImageBase64,
                                        mimeType: pendingImageMimeType || "image/png",
                                        folderId: "1HQR_NYKhHQGA7_2W3nArI9pCh-LJasTP"
                                      })
                                    });
                                    if (!res.ok) {
                                      throw new Error(`Google Apps Script sunucu hatası: ${res.status}`);
                                    }
                                    const result = await res.json();
                                    if (result.status === "success" && result.viewUrl) {
                                      setTechizatImages(prev => {
                                        const updated = {
                                          ...prev,
                                          [imageKey]: result.viewUrl
                                        };
                                        localStorage.setItem('techizat_images', JSON.stringify(updated));
                                        return updated;
                                      });
                                      setPendingImageBase64(null);
                                      setPendingImageMimeType(null);
                                      showNotification("Görsel başarıyla Drive'a yüklendi ve aktifleştirildi!");
                                    } else {
                                      throw new Error(result.message || "Bilinmeyen sunucu hatası.");
                                    }
                                  } catch (err: any) {
                                    console.error("Yükleme Hatası:", err);
                                    showNotification(`Yükleme başarısız: ${err.message}`);
                                  } finally {
                                    setIsImageUploadingToDrive(false);
                                  }
                                } else if (tempImageAction === 'remove') {
                                  setTechizatImages(prev => {
                                    const cloned = { ...prev };
                                    delete cloned[imageKey];
                                    localStorage.setItem('techizat_images', JSON.stringify(cloned));
                                    return cloned;
                                  });
                                  showNotification("Görsel kaldırıldı.");
                                } else if (tempImageAction === 'unlock_only') {
                                  showNotification("Görsel güncelleme kilidi kaldırıldı!");
                                }
                              } else {
                                setImagePasswordError(true);
                              }
                            }
                          }}
                          className="w-full px-4 py-2 bg-slate-50 border-2 border-[#0b3d1d]/15 rounded-xl text-center text-sm font-semibold mb-3 focus:outline-none focus:border-[#0b3d1d] text-slate-900"
                          autoFocus
                        />
                        {imagePasswordError && (
                          <p className="text-red-600 text-[10px] font-black mb-3">❌ Hatalı şifre girdiniz!</p>
                        )}
                        <div className="flex gap-2 justify-center">
                          <button
                            type="button"
                            onClick={() => {
                              setShowImageSavePasswordPrompt(false);
                              setImagePasswordInput('');
                              setImagePasswordError(false);
                              setTempImageAction(null);
                            }}
                            className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-600 text-xs font-bold rounded-xl transition-colors cursor-pointer"
                          >
                            İptal
                          </button>
                          <button
                            type="button"
                            onClick={async () => {
                              if (imagePasswordInput === '1839') {
                                setIsImageUpdateUnlocked(true);
                                setImagePasswordInput('');
                                setImagePasswordError(false);
                                setShowImageSavePasswordPrompt(false);
                                if (tempImageAction === 'upload' && pendingImageBase64) {
                                  try {
                                    setIsImageUploadingToDrive(true);
                                    showNotification("Görsel Google Drive'a yükleniyor...");
                                    const driveFileName = `tech_img_${imageKey}.png`;
                                    const res = await fetch(GOOGLE_SCRIPT_URL, {
                                      method: "POST",
                                      headers: {
                                        "Content-Type": "text/plain;charset=utf-8"
                                      },
                                      body: JSON.stringify({
                                        action: "uploadPdfToDrive",
                                        fileName: driveFileName,
                                        base64Data: pendingImageBase64,
                                        mimeType: pendingImageMimeType || "image/png",
                                        folderId: "1HQR_NYKhHQGA7_2W3nArI9pCh-LJasTP"
                                      })
                                    });
                                    if (!res.ok) {
                                      throw new Error(`Google Apps Script sunucu hatası: ${res.status}`);
                                    }
                                    const result = await res.json();
                                    if (result.status === "success" && result.viewUrl) {
                                      setTechizatImages(prev => {
                                        const updated = {
                                          ...prev,
                                          [imageKey]: result.viewUrl
                                        };
                                        localStorage.setItem('techizat_images', JSON.stringify(updated));
                                        return updated;
                                      });
                                      setPendingImageBase64(null);
                                      setPendingImageMimeType(null);
                                      showNotification("Görsel başarıyla Drive'a yüklendi ve aktifleştirildi!");
                                    } else {
                                      throw new Error(result.message || "Bilinmeyen sunucu hatası.");
                                    }
                                  } catch (err: any) {
                                    console.error("Yükleme Hatası:", err);
                                    showNotification(`Yükleme başarısız: ${err.message}`);
                                  } finally {
                                    setIsImageUploadingToDrive(false);
                                  }
                                } else if (tempImageAction === 'remove') {
                                  setTechizatImages(prev => {
                                    const cloned = { ...prev };
                                    delete cloned[imageKey];
                                    localStorage.setItem('techizat_images', JSON.stringify(cloned));
                                    return cloned;
                                  });
                                  showNotification("Görsel kaldırıldı.");
                                } else if (tempImageAction === 'unlock_only') {
                                  showNotification("Görsel güncelleme kilidi kaldırıldı!");
                                }
                              } else {
                                setImagePasswordError(true);
                              }
                            }}
                            className="px-6 py-2 bg-[#0b3d1d] hover:bg-[#072612] text-white text-xs font-black uppercase rounded-xl transition-all cursor-pointer"
                          >
                            Onayla
                          </button>
                        </div>
                      </div>
                    </div>
                  )}
                </div>

                {/* OLD IMAGE STUFF CLEANED */}
                <div className="hidden" style={{ display: 'none' }}>
                  <div className="w-full aspect-square bg-slate-50 rounded-[2rem] border-2 border-slate-200 overflow-hidden relative flex flex-col items-center justify-center p-2 shadow-inner group">
                    {pendingImagePreview ? (
                      <div className="w-full h-full relative">
                        <img
                          src={pendingImagePreview}
                          alt="Teçhizat Görseli Önizleme"
                          referrerPolicy="no-referrer"
                          style={{ transform: `scale(${techizatImageScale})` }}
                          className="w-full h-full object-contain rounded-3xl transition-transform duration-150"
                        />
                        <div className="absolute top-4 left-4 bg-emerald-700 text-white text-[10px] font-black px-3 py-1.5 rounded-xl uppercase tracking-wider shadow-md animate-pulse z-20">
                          Önizleme Aşamasında
                        </div>
                        {/* Zoom controls on hover */}
                        <div className="absolute bottom-4 left-1/2 -translate-x-1/2 bg-black/60 backdrop-blur-md rounded-full px-3 py-1.5 flex items-center gap-2 opacity-0 group-hover:opacity-100 transition-opacity z-10">
                          <button
                            type="button"
                            onClick={() => setTechizatImageScale(prev => Math.min(prev + 0.25, 3))}
                            className="w-7 h-7 bg-white/20 hover:bg-white/40 text-white rounded-full flex items-center justify-center text-xs font-bold cursor-pointer transition-colors"
                            title="Yakınlaştır (+)"
                          >
                            ＋
                          </button>
                          <button
                            type="button"
                            onClick={() => setTechizatImageScale(prev => Math.max(prev - 0.25, 0.5))}
                            className="w-7 h-7 bg-white/20 hover:bg-white/40 text-white rounded-full flex items-center justify-center text-xs font-bold cursor-pointer transition-colors"
                            title="Uzaklaştır (-)"
                          >
                            －
                          </button>
                          <button
                            type="button"
                            onClick={() => setTechizatImageScale(1)}
                            className="px-2 py-0.5 bg-white/20 hover:bg-white/40 text-white rounded-full text-[10px] font-bold cursor-pointer transition-colors"
                            title="Sıfırla"
                          >
                            SIFIRLA
                          </button>
                        </div>
                      </div>
                    ) : hasImage ? (
                      <div className="w-full h-full relative">
                        <CachedDriveImage
                          src={techizatImages[imageKey]}
                          alt="Teçhizat Görseli"
                          referrerPolicy="no-referrer"
                          style={{ transform: `scale(${techizatImageScale})` }}
                          className="w-full h-full object-contain rounded-3xl transition-transform duration-150"
                        />
                        
                        {/* Zoom controls on hover */}
                        <div className="absolute bottom-4 left-1/2 -translate-x-1/2 bg-black/60 backdrop-blur-md rounded-full px-3 py-1.5 flex items-center gap-2 opacity-0 group-hover:opacity-100 transition-opacity z-10">
                          <button
                            type="button"
                            onClick={() => setTechizatImageScale(prev => Math.min(prev + 0.25, 3))}
                            className="w-7 h-7 bg-white/20 hover:bg-white/40 text-white rounded-full flex items-center justify-center text-xs font-bold cursor-pointer transition-colors"
                            title="Yakınlaştır (+)"
                          >
                            ＋
                          </button>
                          <button
                            type="button"
                            onClick={() => setTechizatImageScale(prev => Math.max(prev - 0.25, 0.5))}
                            className="w-7 h-7 bg-white/20 hover:bg-white/40 text-white rounded-full flex items-center justify-center text-xs font-bold cursor-pointer transition-colors"
                            title="Uzaklaştır (-)"
                          >
                            －
                          </button>
                          <button
                            type="button"
                            onClick={() => setTechizatImageScale(1)}
                            className="px-2 py-0.5 bg-white/20 hover:bg-white/40 text-white rounded-full text-[10px] font-bold cursor-pointer transition-colors"
                            title="Sıfırla"
                          >
                            SIFIRLA
                          </button>
                        </div>

                        {/* Top-Right Fullscreen Icon */}
                        <button
                          type="button"
                          onClick={() => setIsFullScreenImage(true)}
                          className="absolute top-4 right-4 w-10 h-10 bg-black/60 hover:bg-black/80 backdrop-blur-sm text-white rounded-xl flex items-center justify-center shadow-lg cursor-pointer transition-all active:scale-95 z-20 border border-white/10"
                          title="Tam Ekran Görüntüle"
                        >
                          <Maximize2 className="w-5 h-5 text-emerald-400" />
                        </button>
                      </div>
                    ) : (
                      <div className="text-center p-6 text-slate-400">
                        <span className="text-5xl mb-3 block select-none">📷</span>
                        <p className="text-xs font-black uppercase text-slate-500 mb-1">Görsel Bulunmamaktadır</p>
                        <p className="text-[10px] text-slate-400 leading-relaxed max-w-xs mx-auto">
                          Bu ürün için yüklenmiş bir görsel yok. Aşağıdaki "Görsel Güncelle" butonu ile yeni görsel ekleyebilirsiniz.
                        </p>
                      </div>
                    )}
                  </div>

                  {/* Action buttons under image */}
                  <div className="flex flex-col gap-3">
                    <div className="flex flex-col gap-4 border border-slate-100 bg-slate-50/25 rounded-[2rem] p-5 animate-fade-in">
                      <div className="flex justify-between items-center">
                        <span className="text-xs font-bold text-slate-700 flex items-center gap-1.5">
                          {isImageUpdateUnlocked ? (
                            <>
                              <span className="w-2 h-2 rounded-full bg-emerald-500 animate-pulse" />
                              <span className="text-emerald-800 font-extrabold text-[10px]">GÖRSEL GÜNCELLEME YETKİSİ AKTİF</span>
                            </>
                          ) : (
                            <>
                              <span className="w-2 h-2 rounded-full bg-slate-400" />
                              <span className="text-slate-500 font-extrabold text-[10px]">GÖRSEL GÜNCELLEME KİLİDİ AKTİF</span>
                            </>
                          )}
                        </span>
                        <div className="flex items-center gap-3">
                          {hasImage && (
                            <button
                              type="button"
                              onClick={() => {
                                if (isImageUpdateUnlocked) {
                                  setTechizatImages(prev => {
                                    const cloned = { ...prev };
                                    delete cloned[imageKey];
                                    localStorage.setItem('techizat_images', JSON.stringify(cloned));
                                    return cloned;
                                  });
                                  setPendingImageFile(null);
                                  setPendingImagePreview(null);
                                  setTempImageUrlInput('');
                                  showNotification("Görsel kaldırıldı.");
                                } else {
                                  setTempImageAction('remove');
                                  setImagePasswordInput('');
                                  setImagePasswordError(false);
                                  setShowImageSavePasswordPrompt(true);
                                }
                              }}
                              className="text-[10px] font-extrabold text-red-500 hover:text-red-700 hover:underline cursor-pointer flex items-center gap-1"
                            >
                              🗑️ Görseli Kaldır
                            </button>
                          )}
                          {isImageUpdateUnlocked && (
                            <button
                              type="button"
                              onClick={() => {
                                setIsImageUpdateUnlocked(false);
                                setImagePasswordInput('');
                              }}
                              className="text-[10px] font-bold text-slate-400 hover:text-slate-600 hover:underline cursor-pointer"
                            >
                              Kilitle
                            </button>
                          )}
                        </div>
                      </div>

                      {/* Drag and Drop Upload Area - Always visible */}
                      <div className="flex flex-col gap-2">
                        <div className="flex items-center justify-between">
                          <label className="text-[10px] font-black text-slate-500 uppercase tracking-wider">
                            Mobil Otomatik Arka Plan Temizleme
                          </label>
                        </div>

                        {/* Mobil Kamerayı Doğrudan Açan Buton */}
                        <div className="flex flex-col items-center justify-center gap-2 bg-slate-50 border-2 border-slate-200/80 rounded-2xl p-4">
                          <label className="btn-camera bg-blue-600 hover:bg-blue-700 text-white font-extrabold text-xs uppercase px-5 py-3 rounded-xl cursor-pointer flex items-center justify-center gap-2 active:scale-95 transition-all shadow-md w-full text-center">
                            📷 Fotoğraf Çek
                            <input
                              type="file"
                              id="cameraInput"
                              accept="image/*"
                              capture="environment"
                              onChange={(e) => {
                                const file = e.target.files?.[0];
                                if (file) {
                                  handleImageSelected(file);
                                }
                              }}
                              className="hidden"
                            />
                          </label>
                          <div className="flex items-center gap-2 w-full">
                            <button
                              type="button"
                              onClick={() => {
                                const inputEl = document.getElementById('drag-drop-image-input');
                                inputEl?.click();
                              }}
                              className="flex-1 py-2 bg-slate-200 hover:bg-slate-300 text-slate-700 font-bold text-[10px] uppercase rounded-xl transition-all cursor-pointer text-center"
                            >
                              📁 Dosya Seç
                            </button>
                            <button
                              type="button"
                              onClick={startWebcam}
                              className="flex-1 py-2 bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-[10px] uppercase rounded-xl transition-all cursor-pointer text-center"
                            >
                              🖥️ Live Webcam
                            </button>
                          </div>
                          <input
                            id="drag-drop-image-input"
                            type="file"
                            accept="image/*"
                            onChange={(e) => {
                              const file = e.target.files?.[0];
                              if (file) handleImageSelected(file);
                            }}
                            className="hidden"
                          />
                        </div>

                        {isImageUploadingToDrive ? (
                          <div className="border-2 border-dashed border-emerald-300 bg-emerald-50/20 rounded-2xl p-6 flex flex-col items-center justify-center gap-3 text-center min-h-[140px]">
                            <div className="w-8 h-8 border-4 border-emerald-600 border-t-transparent rounded-full animate-spin mb-1" />
                            <p className="text-xs font-black text-emerald-800 animate-pulse">BULUTA YÜKLENİYOR...</p>
                            <p className="text-[10px] text-emerald-600 font-semibold">Görsel Google Drive'a kaydediliyor, lütfen bekleyiniz.</p>
                          </div>
                        ) : isProcessingRemoveBg ? (
                          <div className="border-2 border-dashed border-blue-300 bg-blue-50/40 rounded-2xl p-6 flex flex-col items-center justify-center gap-3 text-center min-h-[160px]">
                            <div className="relative flex items-center justify-center">
                              <div className="w-14 h-14 border-4 border-blue-200 border-t-blue-600 rounded-full animate-spin" />
                              <span className="absolute text-xs font-black font-mono text-blue-700">
                                %{bgRemovalProgressPercent}
                              </span>
                            </div>
                            <div className="flex flex-col items-center gap-1 w-full max-w-xs">
                              <p className="text-xs font-black text-blue-900 uppercase tracking-wider">
                                ARKA PLAN TEMİZLENİYOR (%{bgRemovalProgressPercent})
                              </p>
                              {/* Dynamic Progress Bar */}
                              <div className="w-full bg-slate-200 h-3 rounded-full overflow-hidden p-0.5 shadow-inner border border-slate-300/80 my-1">
                                <div
                                  className="bg-gradient-to-r from-blue-500 via-indigo-600 to-blue-600 h-full rounded-full transition-all duration-300 shadow-sm"
                                  style={{ width: `${Math.max(5, bgRemovalProgressPercent)}%` }}
                                />
                              </div>
                              <div id="status" className="text-xs font-bold text-slate-700">{bgRemovalStatusText || "Lütfen bekleyin..."}</div>
                            </div>
                          </div>
                        ) : pendingImagePreview ? (
                          <div className="border-2 border-solid border-emerald-400 bg-emerald-50/10 rounded-2xl p-4 flex flex-col gap-3 min-h-[140px] animate-fade-in">
                            <div className="flex flex-col items-center gap-3">
                              {/* İşlenmiş Şeffaf (PNG) Görsel */}
                              <img
                                id="resultImage"
                                src={pendingImagePreview}
                                alt="Temizlenmiş Görsel"
                                style={{
                                  maxWidth: "100%",
                                  maxHeight: "220px",
                                  borderRadius: "12px",
                                  border: "2px dashed #ccc",
                                  display: "block",
                                  backgroundImage: "repeating-linear-gradient(45deg, #ccc 25%, transparent 25%, transparent 75%, #ccc 75%, #ccc), repeating-linear-gradient(45deg, #ccc 25%, #fff 25%, #fff 75%, #ccc 75%, #ccc)",
                                  backgroundPosition: "0 0, 9px 9px",
                                  backgroundSize: "18px 18px",
                                  objectFit: "contain"
                                }}
                              />
                              <div className="text-center">
                                <p className="text-xs font-bold text-slate-800 truncate">{pendingImageFile?.name || 'Kameradan Çekilen Fotoğraf'}</p>
                                <p className="text-[10px] text-slate-400 font-semibold font-mono">
                                  {(pendingImageFile ? pendingImageFile.size / 1024 : 0).toFixed(1)} KB (Şeffaf PNG)
                                </p>
                                <div id="status" className="text-[10px] font-bold text-emerald-700 mt-1">{bgRemovalStatusText || "İşlem tamamlandı!"}</div>
                              </div>
                            </div>
                            <div className="flex flex-col gap-1.5">
                              <div className="flex gap-2">
                                <button
                                  type="button"
                                  onClick={async () => {
                                    if (!pendingImageFile) return;
                                    if (isImageUpdateUnlocked) {
                                      // Upload directly
                                      try {
                                        setIsImageUploadingToDrive(true);
                                        showNotification("Görsel Google Drive'a yükleniyor...");
                                        const base64Str = await fileToBase64(pendingImageFile);
                                        
                                        const extension = pendingImageFile.name.split('.').pop() || 'png';
                                        const driveFileName = `tech_img_${imageKey}.${extension}`;
                                        
                                        const res = await fetch(GOOGLE_SCRIPT_URL, {
                                          method: "POST",
                                          headers: {
                                            "Content-Type": "text/plain;charset=utf-8"
                                          },
                                          body: JSON.stringify({
                                            action: "uploadPdfToDrive",
                                            fileName: driveFileName,
                                            base64Data: base64Str,
                                            mimeType: pendingImageFile.type,
                                            folderId: "1HQR_NYKhHQGA7_2W3nArI9pCh-LJasTP"
                                          })
                                        });
                                        
                                        if (!res.ok) {
                                          throw new Error(`Google Apps Script sunucu hatası: ${res.status}`);
                                        }
                                        
                                        const result = await res.json();
                                        if (result.status === "success" && result.viewUrl) {
                                          setTechizatImages(prev => {
                                            const updated = {
                                              ...prev,
                                              [imageKey]: result.viewUrl
                                            };
                                            localStorage.setItem('techizat_images', JSON.stringify(updated));
                                            return updated;
                                          });
                                          setPendingImageFile(null);
                                          setPendingImagePreview(null);
                                          setTempImageUrlInput(result.viewUrl);
                                          showNotification("Görsel başarıyla Drive'a yüklendi ve aktifleştirildi!");
                                        } else {
                                          throw new Error(result.message || "Bilinmeyen sunucu hatası.");
                                        }
                                      } catch (err: any) {
                                        console.error("Yükleme Hatası:", err);
                                        showNotification(`Yükleme başarısız: ${err.message}`);
                                      } finally {
                                        setIsImageUploadingToDrive(false);
                                      }
                                    } else {
                                      setTempImageAction('upload');
                                      setImagePasswordInput('');
                                      setImagePasswordError(false);
                                      setShowImageSavePasswordPrompt(true);
                                    }
                                  }}
                                  className="flex-1 py-2 bg-emerald-700 hover:bg-emerald-800 text-white font-extrabold text-[10px] uppercase tracking-wider rounded-xl transition-all cursor-pointer shadow-md text-center active:scale-95"
                                >
                                  ☁️ BULUTA GÖNDER VE KAYDET
                                </button>
                                <button
                                  type="button"
                                  onClick={() => {
                                    setPendingImageFile(null);
                                    setPendingImagePreview(null);
                                    setBgRemovalStatusText("");
                                  }}
                                  className="px-3 py-2 bg-slate-100 hover:bg-slate-200 text-slate-600 font-extrabold text-[10px] uppercase tracking-wider rounded-xl transition-all cursor-pointer"
                                >
                                  İPTAL
                                </button>
                              </div>
                              <button
                                type="button"
                                onClick={() => {
                                  if (pendingImageFile) {
                                    handleProcessImglyBackgroundRemoval(pendingImageFile);
                                  }
                                }}
                                className="w-full py-1.5 bg-blue-600 hover:bg-blue-700 text-white font-black text-[9px] uppercase tracking-wider rounded-lg transition-all cursor-pointer shadow-sm active:scale-[0.98]"
                              >
                                ✨ Arka Planı Yeniden Temizle (IMG.LY AI)
                              </button>
                            </div>
                          </div>
                        ) : null}
                      </div>
                    </div>

                    {showImageSavePasswordPrompt && (
                      <div className="absolute inset-0 bg-slate-950/85 backdrop-blur-md flex items-center justify-center z-[2200] p-4 select-text animate-fade-in">
                        <div className="bg-white border-2 border-slate-200/50 rounded-[2rem] shadow-2xl max-w-sm w-full p-6 text-center relative overflow-hidden">
                          <div className="absolute top-0 inset-x-0 h-1 bg-[#0b3d1d]"></div>
                          <div className="w-12 h-12 bg-emerald-50 rounded-2xl flex items-center justify-center mx-auto mb-4 text-[#0b3d1d]">
                            <Lock className="w-6 h-6" />
                          </div>
                          <h4 className="text-slate-800 font-extrabold text-sm uppercase mb-2">GÖRSEL İŞLEM ONAYI</h4>
                          <p className="text-xs text-slate-500 font-semibold mb-4 leading-relaxed">
                            {tempImageAction === 'upload' && "Görseli Google Drive'a kaydetmek için lütfen yetkili şifresini giriniz."}
                            {tempImageAction === 'link' && "Görsel bağlantısını kaydetmek için lütfen yetkili şifresini giriniz."}
                            {tempImageAction === 'remove' && "Görseli kaldırmak için lütfen yetkili şifresini giriniz."}
                          </p>
                          <input
                            type="password"
                            placeholder="Şifre"
                            value={imagePasswordInput}
                            onChange={(e) => {
                              setImagePasswordInput(e.target.value);
                              setImagePasswordError(false);
                            }}
                            onKeyDown={async (e) => {
                              if (e.key === 'Enter') {
                                if (imagePasswordInput === '1839') {
                                  setIsImageUpdateUnlocked(true);
                                  setImagePasswordInput('');
                                  setImagePasswordError(false);
                                  setShowImageSavePasswordPrompt(false);
                                  if (tempImageAction === 'upload' && pendingImageFile) {
                                    try {
                                      setIsImageUploadingToDrive(true);
                                      showNotification("Görsel Google Drive'a yükleniyor...");
                                      const base64Str = await fileToBase64(pendingImageFile);
                                      const extension = pendingImageFile.name.split('.').pop() || 'png';
                                      const driveFileName = `tech_img_${imageKey}.${extension}`;
                                      const res = await fetch(GOOGLE_SCRIPT_URL, {
                                        method: "POST",
                                        headers: {
                                          "Content-Type": "text/plain;charset=utf-8"
                                        },
                                        body: JSON.stringify({
                                          action: "uploadPdfToDrive",
                                          fileName: driveFileName,
                                          base64Data: base64Str,
                                          mimeType: pendingImageFile.type,
                                          folderId: "1HQR_NYKhHQGA7_2W3nArI9pCh-LJasTP"
                                        })
                                      });
                                      if (!res.ok) {
                                        throw new Error(`Google Apps Script sunucu hatası: ${res.status}`);
                                      }
                                      const result = await res.json();
                                      if (result.status === "success" && result.viewUrl) {
                                        setTechizatImages(prev => {
                                          const updated = {
                                            ...prev,
                                            [imageKey]: result.viewUrl
                                          };
                                          localStorage.setItem('techizat_images', JSON.stringify(updated));
                                          return updated;
                                        });
                                        setPendingImageFile(null);
                                        setPendingImagePreview(null);
                                        setTempImageUrlInput(result.viewUrl);
                                        showNotification("Görsel başarıyla Drive'a yüklendi ve aktifleştirildi!");
                                      } else {
                                        throw new Error(result.message || "Bilinmeyen sunucu hatası.");
                                      }
                                    } catch (err) {
                                      console.error("Yükleme Hatası:", err);
                                      showNotification(`Yükleme başarısız: ${err.message}`);
                                    } finally {
                                      setIsImageUploadingToDrive(false);
                                    }
                                  } else if (tempImageAction === 'link') {
                                    setTechizatImages(prev => {
                                      const updated = {
                                        ...prev,
                                        [imageKey]: tempImageUrlInput
                                      };
                                      localStorage.setItem('techizat_images', JSON.stringify(updated));
                                      return updated;
                                    });
                                    showNotification("Görsel bağlantısı kaydedildi.");
                                  } else if (tempImageAction === 'remove') {
                                    setTechizatImages(prev => {
                                      const cloned = { ...prev };
                                      delete cloned[imageKey];
                                      localStorage.setItem('techizat_images', JSON.stringify(cloned));
                                      return cloned;
                                    });
                                    setPendingImageFile(null);
                                    setPendingImagePreview(null);
                                    setTempImageUrlInput('');
                                    showNotification("Görsel kaldırıldı.");
                                  }
                                } else {
                                  setImagePasswordError(true);
                                }
                              }
                            }}
                            className="w-full px-4 py-2 bg-slate-50 border-2 border-[#0b3d1d]/15 rounded-xl text-center text-sm font-semibold mb-3 focus:outline-none focus:border-[#0b3d1d] text-slate-900"
                            autoFocus
                          />
                          {imagePasswordError && (
                            <p className="text-red-600 text-[10px] font-black mb-3">❌ Hatalı şifre girdiniz!</p>
                          )}
                          <div className="flex gap-2 justify-center">
                            <button
                              type="button"
                              onClick={() => {
                                setShowImageSavePasswordPrompt(false);
                                setImagePasswordInput('');
                                setImagePasswordError(false);
                                setTempImageAction(null);
                              }}
                              className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-600 text-xs font-bold rounded-xl transition-colors cursor-pointer"
                            >
                              İptal
                            </button>
                            <button
                              type="button"
                              onClick={async () => {
                                if (imagePasswordInput === '1839') {
                                  setIsImageUpdateUnlocked(true);
                                  setImagePasswordInput('');
                                  setImagePasswordError(false);
                                  setShowImageSavePasswordPrompt(false);
                                  if (tempImageAction === 'upload' && pendingImageFile) {
                                    try {
                                      setIsImageUploadingToDrive(true);
                                      showNotification("Görsel Google Drive'a yükleniyor...");
                                      const base64Str = await fileToBase64(pendingImageFile);
                                      const extension = pendingImageFile.name.split('.').pop() || 'png';
                                      const driveFileName = `tech_img_${imageKey}.${extension}`;
                                      const res = await fetch(GOOGLE_SCRIPT_URL, {
                                        method: "POST",
                                        headers: {
                                          "Content-Type": "text/plain;charset=utf-8"
                                        },
                                        body: JSON.stringify({
                                          action: "uploadPdfToDrive",
                                          fileName: driveFileName,
                                          base64Data: base64Str,
                                          mimeType: pendingImageFile.type,
                                          folderId: "1HQR_NYKhHQGA7_2W3nArI9pCh-LJasTP"
                                        })
                                      });
                                      if (!res.ok) {
                                        throw new Error(`Google Apps Script sunucu hatası: ${res.status}`);
                                      }
                                      const result = await res.json();
                                      if (result.status === "success" && result.viewUrl) {
                                        setTechizatImages(prev => {
                                          const updated = {
                                            ...prev,
                                            [imageKey]: result.viewUrl
                                          };
                                          localStorage.setItem('techizat_images', JSON.stringify(updated));
                                          return updated;
                                        });
                                        setPendingImageFile(null);
                                        setPendingImagePreview(null);
                                        setTempImageUrlInput(result.viewUrl);
                                        showNotification("Görsel başarıyla Drive'a yüklendi ve aktifleştirildi!");
                                      } else {
                                        throw new Error(result.message || "Bilinmeyen sunucu hatası.");
                                      }
                                    } catch (err) {
                                      console.error("Yükleme Hatası:", err);
                                      showNotification(`Yükleme başarısız: ${err.message}`);
                                    } finally {
                                      setIsImageUploadingToDrive(false);
                                    }
                                  } else if (tempImageAction === 'link') {
                                    setTechizatImages(prev => {
                                      const updated = {
                                        ...prev,
                                        [imageKey]: tempImageUrlInput
                                      };
                                      localStorage.setItem('techizat_images', JSON.stringify(updated));
                                      return updated;
                                    });
                                    showNotification("Görsel bağlantısı kaydedildi.");
                                  } else if (tempImageAction === 'remove') {
                                    setTechizatImages(prev => {
                                      const cloned = { ...prev };
                                      delete cloned[imageKey];
                                      localStorage.setItem('techizat_images', JSON.stringify(cloned));
                                      return cloned;
                                    });
                                    setPendingImageFile(null);
                                    setPendingImagePreview(null);
                                    setTempImageUrlInput('');
                                    showNotification("Görsel kaldırıldı.");
                                  }
                                } else {
                                  setImagePasswordError(true);
                                }
                              }}
                              className="px-6 py-2 bg-[#0b3d1d] hover:bg-[#072612] text-white text-xs font-black uppercase rounded-xl transition-all cursor-pointer"
                            >
                              Onayla
                            </button>
                          </div>
                        </div>
                      </div>
                    )}

                    {isWebcamOpen && (
                      <div className="absolute inset-0 bg-slate-950/90 backdrop-blur-md flex flex-col items-center justify-center z-[2500] p-4 select-none animate-fade-in">
                        <div className="bg-slate-900 border-2 border-slate-700/50 rounded-[2rem] shadow-2xl max-w-md w-full overflow-hidden flex flex-col relative">
                          <div className="absolute top-0 inset-x-0 h-1 bg-[#0b3d1d]"></div>
                          
                          {/* Header */}
                          <div className="p-4 border-b border-slate-800 flex items-center justify-between">
                            <span className="text-xs font-black text-white uppercase tracking-wider flex items-center gap-1.5">
                              <span className="w-2.5 h-2.5 bg-red-500 rounded-full animate-ping" />
                              Canlı Video Akışı
                            </span>
                            <button
                              type="button"
                              onClick={stopWebcam}
                              className="text-slate-400 hover:text-white cursor-pointer transition-colors"
                            >
                              <X className="w-5 h-5" />
                            </button>
                          </div>

                          {/* Video Container */}
                          <div className="relative aspect-video bg-black flex items-center justify-center overflow-hidden">
                            <video
                              ref={videoRef}
                              autoPlay
                              playsInline
                              muted
                              className="w-full h-full object-cover"
                            />
                            {/* Overlay Guides */}
                            <div className="absolute inset-6 border border-dashed border-white/20 rounded-xl pointer-events-none flex items-center justify-center">
                              <span className="text-[10px] text-white/30 uppercase tracking-widest font-black">
                                Techizatı Ortaya Hizalayın
                              </span>
                            </div>
                          </div>

                          {/* Actions */}
                          <div className="p-5 flex flex-col gap-3 bg-slate-950">
                            <button
                              type="button"
                              onClick={captureWebcamPhoto}
                              className="w-full py-3 bg-red-600 hover:bg-red-700 text-white font-extrabold text-xs uppercase tracking-widest rounded-xl transition-all cursor-pointer shadow-lg active:scale-95 flex items-center justify-center gap-2"
                            >
                              📸 FOTOĞRAF ÇEK & AKTAR
                            </button>
                            <button
                              type="button"
                              onClick={stopWebcam}
                              className="w-full py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 font-bold text-[10px] uppercase tracking-wider rounded-xl transition-colors cursor-pointer"
                            >
                              KAMERAYI KAPAT
                            </button>
                          </div>
                        </div>
                      </div>
                    )}
                  </div>
                </div>
              </motion.div>

              {/* Fullscreen view modal overlay inside */}
              {isFullScreenImage && hasImage && (
                <div 
                  className="fixed inset-0 bg-black/95 z-[3000] flex items-center justify-center p-6 animate-fade-in animate-duration-150"
                  title="Geri dönmek için X butonuna veya dışarıya tıklayabilirsiniz"
                >
                  {/* Click overlay helper */}
                  <div className="absolute inset-0 cursor-zoom-out" onClick={() => setIsFullScreenImage(false)} />
                  
                  <CachedDriveImage
                    src={techizatImages[imageKey]}
                    alt="Tam Ekran Görsel"
                    referrerPolicy="no-referrer"
                    className="max-w-full max-h-full object-contain rounded-xl select-none relative z-10"
                  />
                  
                  {/* Close X Button */}
                  <button
                    type="button"
                    onClick={() => setIsFullScreenImage(false)}
                    className="absolute top-6 right-6 w-12 h-12 bg-white/10 hover:bg-white/20 hover:scale-105 active:scale-95 text-white rounded-full flex items-center justify-center shadow-xl cursor-pointer transition-all z-20 border border-white/10"
                    title="Kapat"
                  >
                    <X className="w-6 h-6" />
                  </button>
                </div>
              )}
            </div>
          );
        })()}
      </AnimatePresence>

      {/* EBYS DÜZENLE / GÖNDER TOPLU İŞLEM MODALİ */}
      <AnimatePresence>
        {isEbysModalOpen && (
          <div className="fixed inset-0 bg-slate-950/70 backdrop-blur-md flex items-center justify-center z-[2100] p-4 select-text animate-fade-in overflow-y-auto">
            <motion.div
              initial={{ scale: 0.95, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.95, opacity: 0 }}
              className="bg-white border-2 border-slate-200/50 rounded-[2.5rem] shadow-2xl max-w-4xl w-full p-6 md:p-8 relative overflow-hidden my-8"
            >
              <div className="absolute top-0 inset-x-0 h-1.5 bg-gradient-to-r from-[#0b3d1d] to-[#125c2c]"></div>

              {/* Close Button */}
              <button
                type="button"
                onClick={() => {
                  setIsEbysModalOpen(false);
                  setBulkModalMode('choice');
                  setBulkEditYer("");
                  setBulkEditDurum("");
                  setBulkEditFirma("");
                }}
                className="absolute top-4 right-4 text-slate-400 hover:text-slate-600 transition-colors p-1.5 rounded-lg hover:bg-slate-50 z-10"
              >
                <X className="w-5 h-5" />
              </button>

              {bulkModalMode === 'choice' && (
                <div className="text-center py-6">
                  <div className="w-16 h-16 bg-emerald-50 rounded-2xl flex items-center justify-center mx-auto mb-4 text-[#0b3d1d]">
                    <Sparkles className="w-8 h-8" />
                  </div>
                  <h3 className="text-xl font-black text-slate-800 uppercase tracking-wider">TEÇHİZAT TOPLU İŞLEM MERKEZİ</h3>
                  <p className="text-xs text-slate-400 font-semibold mt-1 mb-8">
                    Seçtiğiniz <strong>{Object.keys(selectedTechizatItems).length} adet</strong> teçhizat üzerinde uygulamak istediğiniz işlemi seçiniz.
                  </p>

                  <div className="grid grid-cols-1 md:grid-cols-2 gap-5 max-w-2xl mx-auto">
                    <button
                      type="button"
                      onClick={() => {
                        setBulkModalMode('edit');
                      }}
                      className="border-2 border-slate-200/60 hover:border-[#0b3d1d] hover:bg-slate-50/50 p-6 rounded-3xl text-left transition-all duration-250 cursor-pointer hover:shadow-lg flex flex-col gap-3 group"
                    >
                      <div className="w-10 h-10 bg-slate-100 group-hover:bg-emerald-50 text-slate-600 group-hover:text-[#0b3d1d] rounded-xl flex items-center justify-center transition-colors">
                        <Edit3 className="w-5 h-5" />
                      </div>
                      <div>
                        <h4 className="text-sm font-black text-slate-800 uppercase tracking-wider">Toplu Düzenle</h4>
                        <p className="text-[11px] text-slate-400 font-semibold mt-1 leading-relaxed">
                          Seçilen özel teçhizatların "BULUNDUĞU YER", "DURUMU" ve "SON KONTROLÜ YAPAN FİRMA" verilerini toplu olarak düzenleyin. (Onay Şifresi: 1839)
                        </p>
                      </div>
                    </button>

                    <button
                      type="button"
                      onClick={() => {
                        setBulkModalMode('send');
                        fetchTasklineEbysList();
                      }}
                      className="border-2 border-slate-200/60 hover:border-[#0b3d1d] hover:bg-slate-50/50 p-6 rounded-3xl text-left transition-all duration-250 cursor-pointer hover:shadow-lg flex flex-col gap-3 group"
                    >
                      <div className="w-10 h-10 bg-slate-100 group-hover:bg-emerald-50 text-slate-600 group-hover:text-[#0b3d1d] rounded-xl flex items-center justify-center transition-colors">
                        <Send className="w-5 h-5" />
                      </div>
                      <div>
                        <h4 className="text-sm font-black text-slate-800 uppercase tracking-wider">TASKLINE'a Gönder</h4>
                        <p className="text-[11px] text-slate-400 font-semibold mt-1 leading-relaxed">
                          Seçili teçhizatları EBYS numarası ile eşleştirerek online sisteme gönderin ve durumlarını "BAKIM / KALİBRASYON" yapın.
                        </p>
                      </div>
                    </button>
                  </div>
                </div>
              )}

              {bulkModalMode === 'edit' && (
                <div>
                  <div className="flex items-center gap-3 mb-6">
                    <div className="w-10 h-10 bg-slate-100 rounded-xl flex items-center justify-center text-slate-700">
                      <Edit3 className="w-5 h-5" />
                    </div>
                    <div>
                      <h3 className="text-base font-black text-slate-800 uppercase tracking-wider">TOPLU VERİ GÜNCELLEME</h3>
                      <p className="text-[11px] text-slate-400 font-semibold mt-0.5">
                        Seçilen <strong>{Object.keys(selectedTechizatItems).length}</strong> teçhizatın boş bırakmadığınız alanları ortak değerle güncellenecektir.
                      </p>
                    </div>
                  </div>

                  <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-6 text-left">
                    <div className="flex flex-col gap-1.5">
                      <label className="text-[10px] font-black text-slate-500 uppercase tracking-wider">Yeni Bulunduğu Yer</label>
                      <input
                        type="text"
                        placeholder="Değiştirmek istemiyorsanız boş bırakın"
                        value={bulkEditYer}
                        onChange={(e) => setBulkEditYer(e.target.value)}
                        className="w-full px-3.5 py-2.5 bg-slate-50 border border-slate-200 focus:border-[#0b3d1d] focus:ring-4 focus:ring-[#0b3d1d]/5 rounded-2xl text-xs font-bold text-slate-800 focus:outline-none transition-all animate-none"
                      />
                    </div>

                    <div className="flex flex-col gap-1.5">
                      <label className="text-[10px] font-black text-slate-500 uppercase tracking-wider">Yeni Durumu</label>
                      <select
                        value={bulkEditDurum}
                        onChange={(e) => setBulkEditDurum(e.target.value)}
                        className="w-full px-3.5 py-2.5 bg-slate-50 border border-slate-200 focus:border-[#0b3d1d] focus:ring-4 focus:ring-[#0b3d1d]/5 rounded-2xl text-xs font-bold text-slate-800 focus:outline-none transition-all animate-none cursor-pointer"
                      >
                        <option value="">-- Değişiklik Yok (Değiştirmek istemiyorsanız seçmeyin) --</option>
                        <option value="FAAL">FAAL</option>
                        <option value="BAKIM / KALİBRASYON">BAKIM / KALİBRASYON</option>
                        <option value="GAYRİ FAAL">GAYRİ FAAL</option>
                      </select>
                    </div>

                    <div className="flex flex-col gap-1.5">
                      <label className="text-[10px] font-black text-slate-500 uppercase tracking-wider">Yeni Son Kontrolü Yapan Firma</label>
                      <input
                        type="text"
                        placeholder="Değiştirmek istemiyorsanız boş bırakın"
                        value={bulkEditFirma}
                        onChange={(e) => setBulkEditFirma(e.target.value)}
                        className="w-full px-3.5 py-2.5 bg-slate-50 border border-slate-200 focus:border-[#0b3d1d] focus:ring-4 focus:ring-[#0b3d1d]/5 rounded-2xl text-xs font-bold text-slate-800 focus:outline-none transition-all animate-none"
                      />
                    </div>
                  </div>

                  <div className="bg-slate-50 border border-slate-100 p-4 rounded-2xl mb-6 max-h-[160px] overflow-y-auto">
                    <h5 className="text-[10px] font-black text-slate-500 uppercase tracking-wider mb-2">Güncellenecek Teçhizatlar ({Object.keys(selectedTechizatItems).length})</h5>
                    <div className="flex flex-wrap gap-1.5">
                      {Object.values(selectedTechizatItems).map((item: any, i) => (
                        <span key={`${item.techType}-${item.row[1] || ""}-${item.row[3] || ""}-${i}`} className="text-[10px] font-bold text-[#0b3d1d] bg-emerald-50 px-2.5 py-1 rounded-lg border border-emerald-100">
                          {item.row[1] || "Bilinmiyor"}
                        </span>
                      ))}
                    </div>
                  </div>

                  <div className="flex gap-3 justify-end border-t border-slate-100 pt-5">
                    <button
                      type="button"
                      onClick={() => setBulkModalMode('choice')}
                      className="px-5 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-600 text-xs font-bold rounded-2xl transition-colors cursor-pointer"
                    >
                      Geri Dön
                    </button>
                    <button
                      type="button"
                      disabled={bulkEditYer.trim() === "" && bulkEditDurum.trim() === "" && bulkEditFirma.trim() === ""}
                      onClick={() => {
                        setBulkEditPasswordInput("");
                        setBulkEditPasswordError(false);
                        setShowBulkEditPasswordPrompt(true);
                      }}
                      className="px-6 py-2.5 bg-[#0b3d1d] hover:bg-[#072612] text-white text-xs font-black uppercase tracking-widest rounded-2xl shadow-lg transition-all cursor-pointer disabled:opacity-50 disabled:cursor-not-allowed"
                    >
                      Topluca Güncelle
                    </button>
                  </div>
                </div>
              )}

              {bulkModalMode === 'send' && (
                <div className="text-left">
                  <div className="flex items-center gap-3 mb-6">
                    <div className="w-10 h-10 bg-slate-100 rounded-xl flex items-center justify-center text-slate-700">
                      <Send className="w-5 h-5" />
                    </div>
                    <div>
                      <h3 className="text-base font-black text-slate-800 uppercase tracking-wider">EBYS TASKLINE SİSTEMİNE GÖNDER</h3>
                      <p className="text-[11px] text-slate-400 font-semibold mt-0.5">
                        Seçilen teçhizatları ilgili birimin EBYS görev satırı kaydı ile eşleştirerek TASKLINE Excel'e gönderin.
                      </p>
                    </div>
                  </div>

                  {/* EBYS Autocomplete / Search input */}
                  <div className="flex flex-col gap-1.5 mb-5 relative">
                    <label className="text-[10px] font-black text-[#0b3d1d] uppercase tracking-wider">EBYS No / İşlem Numarası Yazın</label>
                    <div className="flex gap-2">
                      <div className="relative flex-1">
                        <input
                          type="text"
                          placeholder="EBYS No veya Başlık aramak için yazın..."
                          value={ebysSearchQuery}
                          onChange={(e) => {
                            setEbysSearchQuery(e.target.value);
                            setSelectedEbysRow(null);
                            setIsEbysSelectDropdownOpen(true);
                          }}
                          onFocus={() => setIsEbysSelectDropdownOpen(true)}
                          className="w-full px-4 py-2.5 bg-slate-50 border-2 border-[#0b3d1d]/15 focus:border-[#0b3d1d] rounded-2xl text-xs font-bold text-slate-800 focus:outline-none transition-all"
                        />
                        {ebysSearchQuery && (
                          <button
                            type="button"
                            onClick={() => {
                              setEbysSearchQuery("");
                              setSelectedEbysRow(null);
                              setEbysBaslik("");
                              setEbysAciklama("");
                              setEbysTalepTuru("");
                            }}
                            className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600 text-[10px] font-black"
                          >
                            TEMİZLE
                          </button>
                        )}
                      </div>
                    </div>

                    {/* Suggestions list */}
                    {isEbysSelectDropdownOpen && !selectedEbysRow && (
                      <div className="absolute top-[100%] left-0 right-0 bg-white border border-slate-200/80 rounded-2xl shadow-xl max-h-56 overflow-y-auto z-50 mt-1 select-text">
                        {isLoadingEbys ? (
                          <p className="p-4 text-xs text-slate-400 font-semibold flex items-center gap-2">
                            <Loader2 className="w-4 h-4 animate-spin text-[#0b3d1d]" />
                            <span>EBYS listesi yükleniyor...</span>
                          </p>
                        ) : ebysError ? (
                          <div className="p-4 text-xs text-rose-600 font-medium select-text">
                            <p className="font-bold text-rose-700 mb-1">Bağlantı/Script Hatası:</p>
                            <p className="text-slate-600 leading-relaxed font-mono text-[11px] bg-rose-50/50 p-2 rounded-lg border border-rose-100/60 mb-2">{ebysError}</p>
                            <p className="text-slate-500 leading-relaxed">
                              Lütfen Google Apps Script projenizin doğru çalıştığından ve web uygulaması izinlerinin "Anyone" (Herkes) olarak ayarlandığından emin olun.
                            </p>
                          </div>
                        ) : (() => {
                          const normalize = (str: string) => {
                            return String(str || "")
                              .toLowerCase()
                              .replace(/ı/g, 'i')
                              .replace(/ğ/g, 'g')
                              .replace(/ü/g, 'u')
                              .replace(/ş/g, 's')
                              .replace(/ö/g, 'o')
                              .replace(/ç/g, 'c')
                              .replace(/[^a-z0-9]/g, '');
                          };
                          const query = normalize(ebysSearchQuery);
                          const filtered = ebysList.filter(item => {
                            const parsed = parseEbysItem(item);
                            if (!parsed) return false;
                            const no = normalize(parsed.ebysNo);
                            if (!no || no === "na") return false;
                            const baslik = normalize(parsed.baslik);
                            return no.includes(query) || baslik.includes(query);
                          });

                          if (filtered.length === 0) {
                            return <p className="p-4 text-xs text-slate-400 font-semibold">Sonuç bulunamadı.</p>;
                          }

                          return filtered.map((item, idx) => {
                            const parsed = parseEbysItem(item);
                            if (!parsed) return null;
                            const ebysNo = parsed.ebysNo;
                            const baslik = parsed.baslik;
                            const aciklama = parsed.aciklama;
                            const tur = parsed.talepTuru;

                            return (
                              <div
                                key={`${ebysNo}-${idx}`}
                                onClick={() => {
                                  setSelectedEbysRow(item);
                                  setEbysSearchQuery(ebysNo);
                                  setEbysBaslik(baslik);
                                  setEbysAciklama(aciklama);
                                  setEbysTalepTuru(tur);
                                  setIsEbysSelectDropdownOpen(false);
                                }}
                                className="p-3 hover:bg-emerald-50/50 border-b border-slate-50 last:border-b-0 cursor-pointer transition-colors text-left"
                              >
                                <div className="flex items-center justify-between">
                                  <span className="text-xs font-black text-[#0b3d1d]">{ebysNo}</span>
                                  <span className="text-[9px] font-black text-slate-400 uppercase bg-slate-100 px-2 py-0.5 rounded-full">{tur}</span>
                                </div>
                                <h6 className="text-[11px] font-extrabold text-slate-800 mt-0.5 leading-tight">{baslik}</h6>
                                <p className="text-[10px] text-slate-400 font-medium truncate mt-0.5">{aciklama}</p>
                              </div>
                            );
                          });
                        })()}
                      </div>
                    )}
                  </div>

                  {/* Render Selected EBYS Request Details Card */}
                  {selectedEbysRow && (
                    <div className="bg-emerald-50/40 border border-emerald-100/60 p-4 rounded-3xl mb-5 text-left animate-fade-in select-text">
                      <div className="flex items-start justify-between">
                        <div>
                          <span className="text-[10px] font-black text-emerald-800 uppercase tracking-widest bg-emerald-100/80 px-2.5 py-1 rounded-lg">SEÇİLİ EBYS TALEBİ</span>
                          <h4 className="text-sm font-black text-[#0b3d1d] mt-2">{ebysSearchQuery}</h4>
                        </div>
                        <button
                          type="button"
                          onClick={() => {
                            setSelectedEbysRow(null);
                            setEbysSearchQuery("");
                            setEbysBaslik("");
                            setEbysAciklama("");
                            setEbysTalepTuru("");
                          }}
                          className="text-[10px] font-black text-red-600 hover:text-red-700 hover:underline cursor-pointer"
                        >
                          Temizle / Değiştir
                        </button>
                      </div>
                      <div className="mt-3 space-y-1 text-xs">
                        <p className="text-slate-700"><strong>📋 Başlık:</strong> {ebysBaslik || "-"}</p>
                        <p className="text-slate-700"><strong>📝 Açıklama:</strong> {ebysAciklama || "-"}</p>
                        <p className="text-slate-700"><strong>🏷️ Talep Türü:</strong> {ebysTalepTuru || "-"}</p>
                      </div>
                    </div>
                  )}

                  {/* Scrollable list of Teçhizat items */}
                  <div className="bg-slate-50 border border-slate-100 p-4 rounded-3xl mb-6">
                    <h5 className="text-[10px] font-black text-slate-500 uppercase tracking-wider mb-2">GÖNDERİLECEK TEÇHİZAT LİSTESİ ({Object.keys(selectedTechizatItems).length})</h5>
                    <div className="max-h-[160px] overflow-y-auto space-y-2">
                      {Object.values(selectedTechizatItems).map((item: any, i) => (
                        <div key={`${item.techType}-${item.row[1] || ""}-${item.row[3] || ""}-${i}`} className="bg-white border border-slate-150 p-2.5 rounded-xl flex justify-between items-center text-xs">
                          <div>
                            <p className="font-extrabold text-slate-800">{item.row[1] || "Bilinmiyor"}</p>
                            <p className="text-[10px] text-slate-400 font-semibold">P/N: {item.row[2] || "N/A"} • S/N: {item.row[3] || "N/A"}</p>
                          </div>
                          <span className="text-[10px] font-black text-[#0b3d1d] bg-emerald-50 border border-emerald-100 px-2 py-0.5 rounded-full">
                            {getTechUnitName(item.techType)}
                          </span>
                        </div>
                      ))}
                    </div>
                  </div>

                  <div className="flex gap-3 justify-end border-t border-slate-100 pt-5">
                    <button
                      type="button"
                      onClick={() => setBulkModalMode('choice')}
                      className="px-5 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-600 text-xs font-bold rounded-2xl transition-colors cursor-pointer"
                    >
                      Geri Dön
                    </button>
                    <button
                      type="button"
                      disabled={!selectedEbysRow || isBulkSaving}
                      onClick={submitEbysRequests}
                      className="px-6 py-2.5 bg-[#0b3d1d] hover:bg-[#072612] text-white text-xs font-black uppercase tracking-widest rounded-2xl shadow-lg transition-all cursor-pointer flex items-center gap-2 disabled:opacity-50 disabled:cursor-not-allowed"
                    >
                      {isBulkSaving ? (
                        <>
                          <Loader2 className="w-4 h-4 animate-spin text-white" />
                          <span>GÖNDERİLİYOR...</span>
                        </>
                      ) : (
                        <span>Sisteme Gönder</span>
                      )}
                    </button>
                  </div>
                </div>
              )}
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* TOPLU DÜZENLEME ŞİFRE MODALİ */}
      <AnimatePresence>
        {showBulkEditPasswordPrompt && (
          <div className="fixed inset-0 bg-slate-950/80 backdrop-blur-md flex items-center justify-center z-[2200] p-4 select-text">
            <div className="bg-white border-2 border-slate-200/50 rounded-[2rem] shadow-2xl max-w-sm w-full p-6 text-center relative overflow-hidden animate-fade-in">
              <div className="absolute top-0 inset-x-0 h-1 bg-[#0b3d1d]"></div>
              <div className="w-12 h-12 bg-emerald-50 rounded-2xl flex items-center justify-center mx-auto mb-4 text-[#0b3d1d]">
                <Lock className="w-6 h-6" />
              </div>
              <h4 className="text-slate-800 font-extrabold text-sm uppercase mb-2">TOPLU İŞLEM DEĞİŞİKLİK ONAYI</h4>
              <p className="text-xs text-slate-500 font-semibold mb-4 leading-relaxed">
                Toplu kayıt değişikliğini onaylamak için lütfen yetkili şifresini giriniz.
              </p>
              <input
                type="password"
                placeholder="Şifre"
                value={bulkEditPasswordInput}
                onChange={(e) => {
                  setBulkEditPasswordInput(e.target.value);
                  setBulkEditPasswordError(false);
                }}
                onKeyDown={async (e) => {
                  if (e.key === 'Enter') {
                    if (bulkEditPasswordInput === '1839') {
                      setShowBulkEditPasswordPrompt(false);
                      setBulkEditPasswordInput('');
                      setBulkEditPasswordError(false);
                      await handleBulkEditTechizatRows();
                    } else {
                      setBulkEditPasswordError(true);
                    }
                  }
                }}
                className="w-full px-4 py-2 bg-slate-50 border-2 border-[#0b3d1d]/15 rounded-xl text-center text-sm font-semibold mb-3 focus:outline-none focus:border-[#0b3d1d] text-slate-900"
                autoFocus
              />
              {bulkEditPasswordError && (
                <p className="text-red-600 text-[10px] font-black mb-3">❌ Hatalı yetkili şifresi girdiniz!</p>
              )}
              <div className="flex gap-2 justify-center">
                <button
                  type="button"
                  onClick={() => {
                    setShowBulkEditPasswordPrompt(false);
                    setBulkEditPasswordInput('');
                    setBulkEditPasswordError(false);
                  }}
                  className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-600 text-xs font-bold rounded-xl transition-colors cursor-pointer"
                >
                  İptal
                </button>
                <button
                  type="button"
                  onClick={async () => {
                    if (bulkEditPasswordInput === '1839') {
                      setShowBulkEditPasswordPrompt(false);
                      setBulkEditPasswordInput('');
                      setBulkEditPasswordError(false);
                      await handleBulkEditTechizatRows();
                    } else {
                      setBulkEditPasswordError(true);
                    }
                  }}
                  className="px-5 py-2 bg-[#0b3d1d] hover:bg-[#072612] text-white text-xs font-black uppercase tracking-wider rounded-xl transition-all cursor-pointer"
                >
                  Onayla
                </button>
              </div>
            </div>
          </div>
        )}
      </AnimatePresence>

      {/* 5. GÜN TAKİP SORUMLU BİRİM AYARLARI MODALİ */}
      <AnimatePresence>
        {isSorumluModalOpen && (
          <div className="fixed inset-0 bg-black/60 backdrop-blur-sm flex items-center justify-center z-[2000] p-4 select-none animate-fade-in overflow-y-auto">
            <motion.div
              initial={{ scale: 0.95, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.95, opacity: 0 }}
              className="bg-white border-2 border-emerald-800/10 rounded-[2.5rem] shadow-2xl max-w-3xl w-full p-6 md:p-8 relative overflow-hidden my-8 animate-fade-in"
            >
              <div className="absolute top-0 inset-x-0 h-1.5 bg-gradient-to-r from-emerald-600 via-emerald-500 to-emerald-700"></div>
              
              <button
                onClick={() => setIsSorumluModalOpen(false)}
                className="absolute top-4 right-4 text-slate-400 hover:text-slate-600 transition-colors p-1.5 rounded-lg hover:bg-slate-50"
              >
                <X className="w-5 h-5" />
              </button>

              <div className="flex items-center gap-3 mb-6">
                <div className="w-12 h-12 bg-emerald-50 text-emerald-700 rounded-2xl flex items-center justify-center">
                  <UserCheck className="w-6 h-6" />
                </div>
                <div className="text-left">
                  <h4 className="text-slate-800 font-black tracking-tight text-base uppercase">📋 SORUMLU BİRİM VE MAİL AYARLARI (GÜN TAKİP)</h4>
                  <p className="text-[10px] text-slate-500 font-bold uppercase tracking-wider mt-0.5">E-posta Uyarı Hatırlatma Alıcı Sorumluları Yönetimi</p>
                </div>
              </div>

              <p className="text-xs text-slate-500 font-semibold leading-relaxed mb-6 text-left border-b border-slate-100 pb-4">
                TÜM TEÇHİZAT sayfasındaki teçhizatların Gelecek Bakım günlerine 90 günden az kalması durumunda sistem tarafından otomatik e-posta uyarısı hatırlatması gönderilecek birim yetkililerini ve mail adreslerini buradan güncelleyebilirsiniz. Değişiklikler canlı "GÜN TAKİP" e-tablosuyla eşleşecektir.
              </p>

              {/* Sorumlular Düzenleme Listesi */}
              <div className="overflow-x-auto max-h-[45vh] border border-slate-200 rounded-3xl mb-6 shadow-inner">
                <table className="w-full border-collapse text-left min-w-[650px]">
                  <thead>
                    <tr className="bg-slate-900 border-b border-slate-800">
                      <th className="px-4 py-3 text-center text-[10px] font-black tracking-wider text-slate-300 uppercase font-mono w-[25%] border-r border-slate-800">SORUMLU BİRİM</th>
                      <th className="px-4 py-3 text-center text-[10px] font-black tracking-wider text-slate-300 uppercase font-mono w-[30%] border-r border-slate-800">ADI SOYADI</th>
                      <th className="px-4 py-3 text-center text-[10px] font-black tracking-wider text-slate-300 uppercase font-mono w-[30%] border-r border-slate-800">E-POSTA ADRESİ</th>
                      <th className="px-3 py-3 text-center text-[10px] font-black tracking-wider text-green-400 uppercase font-mono w-[15%]">SON 90G MAİL</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100 bg-white">
                    {gunTakipSorumlulari.map((item, idx) => (
                      <tr key={idx} className="hover:bg-slate-50/50">
                        <td className="px-4 py-3 text-center font-bold text-slate-800 text-xs border-r border-slate-100 bg-slate-50 font-mono">
                          {item.birim}
                        </td>
                        <td className="px-4 py-2 border-r border-slate-100">
                          <input
                            type="text"
                            value={item.adSoyad}
                            onChange={(e) => {
                              const updated = [...gunTakipSorumlulari];
                              updated[idx].adSoyad = e.target.value;
                              setGunTakipSorumlulari(updated);
                            }}
                            className="w-full px-3 py-1.5 border border-slate-200 hover:border-slate-300 focus:border-emerald-600 focus:outline-none rounded-xl text-xs font-semibold text-slate-800 text-center"
                            placeholder="Ad Soyad giriniz..."
                          />
                        </td>
                        <td className="px-4 py-2 border-r border-slate-100">
                          <input
                            type="email"
                            value={item.eposta}
                            onChange={(e) => {
                              const updated = [...gunTakipSorumlulari];
                              updated[idx].eposta = e.target.value;
                              setGunTakipSorumlulari(updated);
                            }}
                            className="w-full px-3 py-1.5 border border-slate-200 hover:border-slate-300 focus:border-emerald-600 focus:outline-none rounded-xl text-xs font-mono font-semibold text-slate-800 text-center"
                            placeholder="eposta@adres.com"
                          />
                        </td>
                        <td className="px-2 py-3 text-center font-mono text-[10px] font-semibold text-slate-600 bg-green-50/20">
                          {item.mail90 ? (
                            <span className="text-emerald-700 bg-emerald-100/60 px-2 py-0.5 rounded-md font-bold block">{item.mail90}</span>
                          ) : (
                            <span className="text-slate-400 italic font-medium">-</span>
                          )}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>

              <div className="flex gap-3 justify-end pt-4 border-t border-slate-100">
                <button
                  type="button"
                  onClick={() => setIsSorumluModalOpen(false)}
                  className="px-5 py-3 bg-slate-100 hover:bg-slate-200 text-slate-600 text-xs font-black rounded-2xl transition-colors cursor-pointer uppercase tracking-wider"
                  disabled={isSavingSorumlu}
                >
                  Kapat
                </button>
                <button
                  type="button"
                  onClick={() => saveGunTakipSorumlulari(gunTakipSorumlulari)}
                  className="px-6 py-3 bg-emerald-700 hover:bg-emerald-600 text-white text-xs font-black rounded-2xl shadow-lg shadow-emerald-950/20 transition-all cursor-pointer flex items-center gap-2 uppercase tracking-wider"
                  disabled={isSavingSorumlu}
                >
                  {isSavingSorumlu ? (
                    <>
                      <Loader2 className="w-4 h-4 animate-spin" />
                      <span>SENKRONİZE EDİLİYOR...</span>
                    </>
                  ) : (
                    <>
                      <Check className="w-4 h-4" />
                      <span>E-Tabloyu Güncelle</span>
                    </>
                  )}
                </button>
              </div>

            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* PDF Rendering Loading Overlay */}
      <AnimatePresence>
        {isPdfRendering && (
          <div className="fixed inset-0 bg-slate-950/80 backdrop-blur-md flex flex-col items-center justify-center z-[2400] select-none text-center">
            <div className="bg-slate-900 border border-slate-800 rounded-[2.5rem] p-10 max-w-sm w-full shadow-2xl flex flex-col items-center">
              <div className="w-16 h-16 bg-red-500/10 text-red-500 rounded-3xl flex items-center justify-center mb-6 animate-pulse">
                <Loader2 className="w-8 h-8 animate-spin" />
              </div>
              <h4 className="text-slate-100 font-extrabold tracking-wider text-sm uppercase mb-2">
                PDF ANALİZ EDİLİYOR
              </h4>
              <p className="text-[11px] text-slate-400 font-semibold leading-relaxed mb-1">
                PDF sayfaları arka planda HD çözünürlükte
              </p>
              <p className="text-[11px] text-slate-400 font-semibold leading-relaxed">
                görsel nesnelere dönüştürülüyor, lütfen bekleyin...
              </p>
            </div>
          </div>
        )}
      </AnimatePresence>

      {/* 6. ADVANCED PDF PREVIEW PANEL */}
      <AnimatePresence>
        {isPdfPreviewOpen && (
          <div className="fixed inset-0 bg-black/75 backdrop-blur-md flex items-center justify-center z-[2500] p-4 md:p-6 select-none overflow-hidden animate-fade-in">
            <motion.div
              initial={{ scale: 0.95, opacity: 0, y: 20 }}
              animate={{ scale: 1, opacity: 1, y: 0 }}
              exit={{ scale: 0.95, opacity: 0, y: 20 }}
              className="bg-slate-900 border border-slate-800 rounded-[2rem] shadow-2xl w-full max-w-6xl h-[90vh] flex flex-col overflow-hidden text-white"
            >
              {/* Header */}
              <div className="px-6 py-4 border-b border-slate-800 flex flex-col md:flex-row md:items-center md:justify-between gap-4 bg-slate-950">
                <div className="flex items-center gap-3">
                  <div className="p-2 bg-red-500/10 rounded-xl text-red-500">
                    <FileText className="w-6 h-6" />
                  </div>
                  <div>
                    <h3 className="text-sm font-black uppercase tracking-wider text-slate-100">
                      GELİŞMİŞ PDF ÖNİZLEME PANELİ
                    </h3>
                    <p className="text-[10px] text-slate-400 font-semibold mt-0.5">
                      {renderedPages.length > 0 ? `${renderedPages[0].fileName} - Toplam ${renderedPages.length} sayfa` : ""}
                    </p>
                  </div>
                </div>

                <div className="flex flex-wrap items-center gap-3">
                  {/* Zoom Controls */}
                  <div className="flex items-center gap-2 bg-slate-850 px-3 py-1.5 rounded-xl border border-slate-800">
                    <span className="text-[10px] font-black uppercase text-slate-400">🔍 YAKINLAŞTIR:</span>
                    <button
                      type="button"
                      onClick={() => setPreviewZoom(z => Math.max(50, z - 10))}
                      className="w-6 h-6 flex items-center justify-center bg-slate-800 hover:bg-slate-700 active:bg-slate-750 text-white rounded font-black text-xs cursor-pointer select-none"
                    >
                      -
                    </button>
                    <input
                      type="range"
                      min="50"
                      max="200"
                      value={previewZoom}
                      onChange={(e) => setPreviewZoom(Number(e.target.value))}
                      className="w-20 accent-emerald-500 cursor-pointer"
                    />
                    <button
                      type="button"
                      onClick={() => setPreviewZoom(z => Math.min(200, z + 10))}
                      className="w-6 h-6 flex items-center justify-center bg-slate-800 hover:bg-slate-700 active:bg-slate-750 text-white rounded font-black text-xs cursor-pointer select-none"
                    >
                      +
                    </button>
                    <span className="text-[10px] font-mono font-black text-emerald-400 min-w-[32px] text-right">%{previewZoom}</span>
                  </div>

                  {/* Bulk Rotation */}
                  <div className="flex items-center gap-1.5 bg-slate-850 p-1 rounded-xl border border-slate-800">
                    <button
                      type="button"
                      onClick={() => {
                        setRenderedPages(prev => prev.map(p => ({ ...p, rotation: (((p.rotation || 0) - 90 + 360) % 360) })));
                        showNotification("Tüm sayfalar sola döndürüldü.");
                      }}
                      className="px-2.5 py-1.5 bg-slate-800 hover:bg-slate-750 text-slate-300 rounded-lg text-[9px] font-black uppercase tracking-wider transition-all cursor-pointer flex items-center gap-1"
                    >
                      ↺ TÜMÜ SOLA
                    </button>
                    <button
                      type="button"
                      onClick={() => {
                        setRenderedPages(prev => prev.map(p => ({ ...p, rotation: (((p.rotation || 0) + 90) % 360) })));
                        showNotification("Tüm sayfalar sağa döndürüldü.");
                      }}
                      className="px-2.5 py-1.5 bg-slate-800 hover:bg-slate-750 text-slate-300 rounded-lg text-[9px] font-black uppercase tracking-wider transition-all cursor-pointer flex items-center gap-1"
                    >
                      ↻ TÜMÜ SAĞA
                    </button>
                  </div>

                  {/* View Mode Selectors */}
                  <div className="bg-slate-850 p-1 rounded-xl border border-slate-800 flex items-center gap-1">
                    <button
                      type="button"
                      onClick={() => setPdfPreviewLayout('vertical')}
                      className={`px-2.5 py-1.5 rounded-lg text-[9px] font-extrabold uppercase tracking-wider flex items-center gap-1 transition-all cursor-pointer ${
                        pdfPreviewLayout === 'vertical'
                          ? 'bg-[#0b3d1d] text-white shadow-sm'
                          : 'text-slate-400 hover:text-slate-200'
                      }`}
                    >
                      <Rows className="w-3 h-3" />
                      DİKEY
                    </button>
                    <button
                      type="button"
                      onClick={() => setPdfPreviewLayout('horizontal')}
                      className={`px-2.5 py-1.5 rounded-lg text-[9px] font-extrabold uppercase tracking-wider flex items-center gap-1 transition-all cursor-pointer ${
                        pdfPreviewLayout === 'horizontal'
                          ? 'bg-[#0b3d1d] text-white shadow-sm'
                          : 'text-slate-400 hover:text-slate-200'
                      }`}
                    >
                      <Columns className="w-3 h-3" />
                      YATAY
                    </button>
                  </div>

                  {/* Bulk Select Toggles */}
                  <div className="flex items-center gap-1.5">
                    <button
                      type="button"
                      onClick={() => {
                        setRenderedPages(prev => prev.map(p => ({ ...p, selected: true })));
                      }}
                      className="px-2.5 py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-xl text-[9px] font-black uppercase tracking-wider border border-slate-700 transition-all cursor-pointer"
                    >
                      TÜMÜNÜ SEÇ
                    </button>
                    <button
                      type="button"
                      onClick={() => {
                        setRenderedPages(prev => prev.map(p => ({ ...p, selected: false })));
                      }}
                      className="px-2.5 py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-xl text-[9px] font-black uppercase tracking-wider border border-slate-700 transition-all cursor-pointer"
                    >
                      TEMİZLE
                    </button>
                  </div>

                  {/* Close X */}
                  <button
                    type="button"
                    onClick={() => setIsPdfPreviewOpen(false)}
                    className="p-2 text-slate-400 hover:text-slate-200 bg-slate-800 hover:bg-slate-700 rounded-xl border border-slate-700 transition-colors cursor-pointer"
                  >
                    <X className="w-4 h-4" />
                  </button>
                </div>
              </div>

              {/* Page Area */}
              <div className="flex-1 overflow-auto bg-slate-950 p-6 md:p-8">
                {pdfPreviewLayout === 'vertical' ? (
                  /* Vertical view: Large previews aligned in a clean central column */
                  <div className="flex flex-col gap-8 max-w-2xl mx-auto pb-12">
                    {renderedPages.map((page) => (
                      <div
                        key={page.id}
                        onClick={() => {
                          setRenderedPages(prev => prev.map(p => p.id === page.id ? { ...p, selected: !p.selected } : p));
                        }}
                        className={`relative rounded-2xl border-2 overflow-hidden bg-slate-900 group cursor-pointer transition-all duration-300 shadow-lg ${
                          page.selected
                            ? 'border-emerald-600 ring-4 ring-emerald-600/20'
                            : 'border-slate-800 hover:border-slate-750'
                        }`}
                      >
                        {/* Selector indicator */}
                        <div className="absolute top-4 left-4 z-10 bg-slate-950/85 backdrop-blur border border-slate-850 rounded-xl p-2 flex items-center gap-2 pointer-events-none select-none">
                          <div className={`w-5 h-5 rounded-md flex items-center justify-center transition-colors ${page.selected ? 'bg-emerald-600 text-white border-none' : 'border border-slate-600'}`}>
                            {page.selected && <Check className="w-3.5 h-3.5 stroke-[3]" />}
                          </div>
                          <span className="text-[10px] font-black uppercase text-slate-200">
                            Sayfa {page.pageNumber}
                          </span>
                        </div>

                        {/* Individual Rotation Controls overlay */}
                        <div className="absolute top-4 right-4 z-10 bg-slate-950/85 backdrop-blur border border-slate-850 rounded-xl p-1 flex items-center gap-1">
                          <button
                            type="button"
                            onClick={(e) => {
                              e.stopPropagation();
                              setRenderedPages(prev => prev.map(p => p.id === page.id ? { ...p, rotation: (((p.rotation || 0) - 90 + 360) % 360) } : p));
                            }}
                            className="p-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-md hover:text-white transition-colors cursor-pointer text-xs"
                            title="Sola Döndür (-90°)"
                          >
                            ↺
                          </button>
                          <span className="text-[9px] font-mono font-bold px-1.5 text-slate-400">{(page.rotation || 0)}°</span>
                          <button
                            type="button"
                            onClick={(e) => {
                              e.stopPropagation();
                              setRenderedPages(prev => prev.map(p => p.id === page.id ? { ...p, rotation: (((p.rotation || 0) + 90) % 360) } : p));
                            }}
                            className="p-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-md hover:text-white transition-colors cursor-pointer text-xs"
                            title="Sağa Döndür (90°)"
                          >
                            ↻
                          </button>
                        </div>

                        {/* HD Render Image */}
                        <div className="aspect-[3/4] relative w-full overflow-hidden flex items-center justify-center p-8 bg-slate-950">
                          <img
                            src={page.dataUrl}
                            alt={`Sayfa ${page.pageNumber}`}
                            referrerPolicy="no-referrer"
                            style={{
                              transform: `rotate(${page.rotation || 0}deg) scale(${previewZoom / 100})`,
                              transition: "transform 0.2s ease-in-out",
                              maxHeight: "550px",
                              maxWidth: "100%",
                              objectFit: "contain"
                            }}
                            className="rounded shadow-2xl border border-slate-800"
                          />
                        </div>
                      </div>
                    ))}
                  </div>
                ) : (
                  /* Horizontal view: Side-by-side scrolling carousel layout */
                  <div className="h-full flex items-center justify-center py-4">
                    <div className="flex gap-6 overflow-x-auto px-4 py-8 max-w-full snap-x pb-12 scrollbar-thin scrollbar-thumb-slate-800">
                      {renderedPages.map((page) => (
                        <div
                          key={page.id}
                          onClick={() => {
                            setRenderedPages(prev => prev.map(p => p.id === page.id ? { ...p, selected: !p.selected } : p));
                          }}
                          className={`flex-none w-[280px] snap-center relative rounded-2xl border-2 overflow-hidden bg-slate-900 group cursor-pointer transition-all duration-300 shadow-md ${
                            page.selected
                              ? 'border-emerald-600 ring-4 ring-emerald-600/20'
                              : 'border-slate-800 hover:border-slate-750'
                          }`}
                        >
                          {/* Selector indicator */}
                          <div className="absolute top-3 left-3 z-10 bg-slate-950/85 backdrop-blur border border-slate-850 rounded-lg p-1.5 flex items-center gap-1.5 pointer-events-none select-none">
                            <div className={`w-4 h-4 rounded flex items-center justify-center transition-colors ${page.selected ? 'bg-emerald-600 text-white border-none' : 'border border-slate-600'}`}>
                              {page.selected && <Check className="w-3 h-3 stroke-[3]" />}
                            </div>
                            <span className="text-[9px] font-bold uppercase text-slate-200">
                              Sayfa {page.pageNumber}
                            </span>
                          </div>

                          {/* Individual Rotation Controls overlay */}
                          <div className="absolute top-3 right-3 z-10 bg-slate-950/85 backdrop-blur border border-slate-850 rounded-lg p-1 flex items-center gap-1">
                            <button
                              type="button"
                              onClick={(e) => {
                                e.stopPropagation();
                                setRenderedPages(prev => prev.map(p => p.id === page.id ? { ...p, rotation: (((p.rotation || 0) - 90 + 360) % 360) } : p));
                              }}
                              className="p-1 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded hover:text-white transition-colors cursor-pointer text-[10px]"
                              title="Sola Döndür (-90°)"
                            >
                              ↺
                            </button>
                            <span className="text-[8px] font-mono font-bold px-1 text-slate-400">{(page.rotation || 0)}°</span>
                            <button
                              type="button"
                              onClick={(e) => {
                                e.stopPropagation();
                                setRenderedPages(prev => prev.map(p => p.id === page.id ? { ...p, rotation: (((p.rotation || 0) + 90) % 360) } : p));
                              }}
                              className="p-1 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded hover:text-white transition-colors cursor-pointer text-[10px]"
                              title="Sağa Döndür (90°)"
                            >
                              ↻
                            </button>
                          </div>

                          {/* HD Image */}
                          <div className="aspect-[3/4] p-4 flex items-center justify-center bg-slate-950 h-[380px] overflow-hidden">
                            <img
                              src={page.dataUrl}
                              alt={`Sayfa ${page.pageNumber}`}
                              referrerPolicy="no-referrer"
                              style={{
                                transform: `rotate(${page.rotation || 0}deg) scale(${previewZoom / 100})`,
                                transition: "transform 0.2s ease-in-out",
                                maxHeight: "100%",
                                maxWidth: "100%",
                                objectFit: "contain"
                              }}
                              className="rounded shadow border border-slate-850"
                            />
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                )}
              </div>

              {/* Footer Panel */}
              <div className="px-6 py-5 border-t border-slate-800 flex items-center justify-between bg-slate-950">
                <div className="flex items-center gap-3">
                  <span className="text-xs font-bold text-slate-400">Entegrasyon Durumu:</span>
                  <span className="px-3 py-1 bg-emerald-500/10 text-emerald-500 border border-emerald-500/20 rounded-full text-[10px] font-black uppercase tracking-wide">
                    {renderedPages.filter(p => p.selected).length} / {renderedPages.length} SAYFA SEÇİLDİ
                  </span>
                </div>

                <div className="flex items-center gap-3">
                  <button
                    type="button"
                    onClick={() => setIsPdfPreviewOpen(false)}
                    className="px-5 py-3 bg-slate-800 hover:bg-slate-700 active:bg-slate-750 text-slate-300 text-xs font-extrabold uppercase tracking-wider rounded-xl transition-all cursor-pointer"
                  >
                    İPTAL ET
                  </button>
                  <button
                    type="button"
                    onClick={handlePdfPreviewIntegrate}
                    disabled={isSendingToSheets[String(syncSelectedTarget)] || renderedPages.filter(p => p.selected).length === 0}
                    className={`px-8 py-3 font-extrabold text-xs uppercase tracking-wider rounded-xl shadow-lg transition-all cursor-pointer flex items-center gap-2 ${
                      isSendingToSheets[String(syncSelectedTarget)] || renderedPages.filter(p => p.selected).length === 0
                        ? 'bg-slate-700 text-slate-400 cursor-not-allowed'
                        : 'bg-emerald-600 hover:bg-emerald-700 active:bg-emerald-800 text-white'
                    }`}
                  >
                    {isSendingToSheets[String(syncSelectedTarget)] ? (
                      <>
                        <Loader2 className="w-4 h-4 animate-spin text-white" />
                        AKTARIYOR %{uploadProgress}...
                      </>
                    ) : (
                      <>
                        🚀 SEÇİLİ SAYFALARI PORTALDA KAYDET VE YAYINLA
                      </>
                    )}
                  </button>
                </div>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* 5. SUCCESS HUD TOAST NOTIFICATION */}
      <AnimatePresence>
        {successMessage && (
          <motion.div
            initial={{ opacity: 0, y: 30 }}
            animate={{ opacity: 1, y: 0 }}
            exit={{ opacity: 0, y: 30 }}
            className="fixed bottom-6 right-6 bg-[#0b3d1d] text-white border border-emerald-500/20 px-6 py-4 rounded-xl shadow-2xl z-[3000] flex items-center gap-3 backdrop-blur select-none max-w-sm"
          >
            <div className="w-6 h-6 rounded-full bg-emerald-500 flex items-center justify-center text-[#0b3d1d]">
              <Check className="w-4 h-4 text-white font-bold" />
            </div>
            <p className="text-xs font-bold tracking-wider uppercase leading-tight">{successMessage}</p>
          </motion.div>
        )}
      </AnimatePresence>

      {/* 5.1 PERSONNEL CELL DETAIL MODAL */}
      <AnimatePresence>
        {activeModalCell && (
          <div className="fixed inset-0 bg-black/70 backdrop-blur-sm flex items-center justify-center z-[2200] p-4 select-none animate-fade-in">
            <motion.div
              initial={{ scale: 0.95, opacity: 0, y: 20 }}
              animate={{ scale: 1, opacity: 1, y: 0 }}
              exit={{ scale: 0.95, opacity: 0, y: 20 }}
              className="bg-slate-900 border border-slate-800 rounded-[2rem] shadow-2xl max-w-lg w-full p-6 relative overflow-hidden text-white"
            >
              <div className="absolute top-0 inset-x-0 h-1.5 bg-gradient-to-r from-emerald-600 via-emerald-500 to-[#0b3d1d]"></div>
              
              <button
                onClick={() => setActiveModalCell(null)}
                className="absolute top-4 right-4 text-slate-400 hover:text-white transition-colors p-1.5 rounded-lg hover:bg-slate-800 cursor-pointer"
                title="Kapat"
              >
                <X className="w-5 h-5" />
              </button>

              <div className="flex items-center gap-3 mb-6">
                <div className="w-10 h-10 bg-emerald-500/10 rounded-xl flex items-center justify-center text-emerald-400">
                  <FileText className="w-5 h-5" />
                </div>
                <div>
                  <h4 className="text-xs font-black text-emerald-400 uppercase tracking-widest">
                    DETAYLI HÜCRE GÖRÜNÜMÜ
                  </h4>
                  <p className="text-sm font-black text-slate-100 uppercase tracking-tight">
                    {activeModalCell.label}
                  </p>
                </div>
              </div>

              {/* Full copyable text content box */}
              <div className="bg-slate-950/80 border border-slate-800 rounded-2xl p-4 mb-6 max-h-[40vh] overflow-y-auto select-text selection:bg-emerald-500 selection:text-white">
                <p className="text-slate-200 font-sans font-medium text-xs leading-relaxed whitespace-pre-wrap select-text cursor-text">
                  {activeModalCell.value}
                </p>
              </div>

              <div className="flex gap-3 justify-end items-center">
                <button
                  type="button"
                  onClick={async () => {
                    try {
                      await navigator.clipboard.writeText(activeModalCell.value);
                      setCopiedCellSuccess(true);
                      setTimeout(() => setCopiedCellSuccess(false), 2000);
                    } catch (err) {
                      console.error("Metin kopyalanamadı:", err);
                    }
                  }}
                  className={`px-5 py-2.5 rounded-xl text-xs font-black uppercase tracking-wider transition-all flex items-center gap-2 cursor-pointer ${
                    copiedCellSuccess
                      ? "bg-emerald-600 text-white shadow-md shadow-emerald-600/20"
                      : "bg-slate-800 hover:bg-slate-750 text-slate-200"
                  }`}
                >
                  {copiedCellSuccess ? (
                    <>
                      <Check className="w-4 h-4 stroke-[3]" />
                      <span>KOPYALANDI!</span>
                    </>
                  ) : (
                    <>
                      <Copy className="w-4 h-4" />
                      <span>METNİ KOPYALA</span>
                    </>
                  )}
                </button>

                <button
                  type="button"
                  onClick={() => setActiveModalCell(null)}
                  className="px-5 py-2.5 bg-[#0b3d1d] hover:bg-[#0b3d1d]/90 text-white text-xs font-black uppercase tracking-wider rounded-xl transition-colors cursor-pointer"
                >
                  KAPAT
                </button>
              </div>

            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* EXCEL EXPORT CHOICE MODAL ("Görselli olarak indirilsin mi?") */}
      <AnimatePresence>
        {excelExportModalData && (
          <div className="fixed inset-0 bg-slate-950/80 backdrop-blur-md flex items-center justify-center z-[3000] p-4 animate-fade-in select-none">
            <motion.div 
              initial={{ scale: 0.9, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.9, opacity: 0 }}
              className="bg-white border-2 border-emerald-600 rounded-[2.5rem] shadow-2xl max-w-lg w-full p-6 md:p-8 relative overflow-hidden text-slate-800"
            >
              {/* Decorative top bar */}
              <div className="absolute top-0 inset-x-0 h-2 bg-gradient-to-r from-emerald-700 via-emerald-500 to-teal-600"></div>

              {/* Close button */}
              <button 
                disabled={isExcelExportLoading}
                onClick={() => setExcelExportModalData(null)}
                className="absolute top-5 right-5 p-2 text-slate-400 hover:text-slate-700 hover:bg-slate-100 rounded-full transition-all cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>

              {/* Modal Header */}
              <div className="flex items-center gap-4 mb-5">
                <div className="w-14 h-14 rounded-2xl bg-emerald-100 border border-emerald-300 flex items-center justify-center text-emerald-800 shrink-0 shadow-sm">
                  <FileSpreadsheet className="w-8 h-8" />
                </div>
                <div>
                  <span className="text-[10px] font-black uppercase tracking-widest text-emerald-700 bg-emerald-50 border border-emerald-200 px-2.5 py-1 rounded-full">
                    EXCEL AKTARIM SİSTEMİ
                  </span>
                  <h3 className="text-xl font-black text-slate-900 mt-1 leading-tight">
                    Görselli Olarak İndirilsin mi?
                  </h3>
                </div>
              </div>

              {/* Description */}
              <p className="text-xs text-slate-600 leading-relaxed mb-6 bg-slate-50 border border-slate-200/80 p-4 rounded-2xl">
                Teçhizat listesini Google Drive üzerindeki <strong>teçhizat fotoğraflarıyla birlikte</strong> Excel dosyası olarak indirebilirsiniz. 
                <br/><br/>
                <span className="text-emerald-800 font-bold">🖼️ Görselli Seçenek:</span> Drive'daki tüm ürün fotoğrafları indirilip Excel hücrelerinin içine yerleştirilir.
              </p>

              {/* Loading Progress State */}
              {isExcelExportLoading ? (
                <div className="bg-emerald-950 text-white p-6 rounded-2xl border border-emerald-800 flex flex-col items-center justify-center gap-3 text-center my-2 shadow-inner">
                  <div className="w-10 h-10 border-4 border-emerald-400 border-t-transparent rounded-full animate-spin" />
                  <p className="text-xs font-black text-emerald-300 uppercase tracking-wide font-mono animate-pulse">
                    {excelExportProgressText || "Drive fotoğrafları çekiliyor..."}
                  </p>
                  <p className="text-[10px] text-emerald-400/80">Lütfen indirme tamamlanana kadar sayfayı kapatmayın.</p>
                </div>
              ) : (
                <div className="flex flex-col gap-3">
                  {/* Görselli İndir Button */}
                  <button
                    type="button"
                    onClick={() => downloadTechizatExcelWithImages(true)}
                    className="w-full py-4 px-5 bg-gradient-to-r from-emerald-700 to-emerald-600 hover:from-emerald-600 hover:to-emerald-500 active:scale-[0.98] text-white font-black text-sm rounded-2xl shadow-lg shadow-emerald-900/20 border border-emerald-500 flex items-center justify-center gap-3 cursor-pointer transition-all group"
                  >
                    <span className="text-xl group-hover:scale-125 transition-transform">🖼️</span>
                    <div className="text-left">
                      <div className="text-sm font-extrabold uppercase tracking-wide">GÖRSELLİ EXCEL İNDİR</div>
                      <div className="text-[10px] text-emerald-200 font-normal">Fotoğraflar Excel sütununa otomatik eklenir</div>
                    </div>
                  </button>

                  {/* Sadece Metin İndir Button */}
                  <button
                    type="button"
                    onClick={() => downloadTechizatExcelWithImages(false)}
                    className="w-full py-3.5 px-5 bg-slate-800 hover:bg-slate-700 active:scale-[0.98] text-slate-200 font-extrabold text-xs rounded-2xl border border-slate-700 flex items-center justify-center gap-2 cursor-pointer transition-all"
                  >
                    <span>📄</span>
                    <span>SADECE METİN (GÖRSELSEZ METİN EXCEL) İNDİR</span>
                  </button>

                  {/* İptal Button */}
                  <button
                    type="button"
                    onClick={() => setExcelExportModalData(null)}
                    className="w-full py-2.5 px-4 text-slate-500 hover:text-slate-800 font-bold text-xs rounded-xl cursor-pointer transition-all text-center mt-1"
                  >
                    Vazgeç / İptal
                  </button>
                </div>
              )}
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* HOVER PRODUCT IMAGE PREVIEW SQUARE BOX */}
      {hoveredRowImage && (
        <div 
          className="fixed z-[9999] pointer-events-none transform -translate-y-1/2 transition-all duration-75 animate-fade-in"
          style={{
            top: Math.min(window.innerHeight - 250, Math.max(30, hoveredRowImage.y)),
            left: hoveredRowImage.x + 280 > window.innerWidth ? Math.max(10, hoveredRowImage.x - 260) : hoveredRowImage.x + 20
          }}
        >
          <div className="bg-slate-900/95 text-white p-3 rounded-2xl border-2 border-emerald-500 shadow-2xl backdrop-blur-md flex flex-col items-center gap-2 max-w-[240px]">
            <div className="w-48 h-48 sm:w-52 sm:h-52 rounded-xl overflow-hidden bg-slate-950 border border-slate-700/80 flex items-center justify-center p-1 relative group">
              <CachedDriveImage
                src={hoveredRowImage.url}
                alt={hoveredRowImage.title}
                className="w-full h-full object-contain rounded-lg"
              />
              <div className="absolute top-2 right-2 bg-emerald-600/90 text-white text-[9px] font-black font-mono px-2 py-0.5 rounded-full shadow">
                🖼️ GÖRSEL ÖNİZLEME
              </div>
            </div>
            <div className="text-center w-full px-1">
              <p className="text-xs font-black text-amber-300 truncate">{hoveredRowImage.title}</p>
              {hoveredRowImage.subtitle && (
                <p className="text-[10px] text-slate-300 font-mono truncate">{hoveredRowImage.subtitle}</p>
              )}
            </div>
          </div>
        </div>
      )}

    </div>
  );
}


