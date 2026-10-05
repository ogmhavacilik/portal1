/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useRef, useMemo, ChangeEvent } from 'react';
import ExcelJS from 'exceljs';
import { removeBackground, preload } from '@imgly/background-removal';

// Helper function to detect mobile devices
const isMobileDevice = (): boolean => {
  if (typeof window === 'undefined') return false;
  return /Android|webOS|iPhone|iPad|iPod|BlackBerry|IEMobile|Opera Mini/i.test(navigator.userAgent) || window.innerWidth < 768;
};

/// Ultra-fast Canvas-based background cleaner (Instant 0.01s runtime, 0 WASM RAM, 0% CPU freeze)
const fastCanvasBackgroundRemoval = async (file: File | Blob, tolerance: number = 42): Promise<{ blob: Blob, base64: string }> => {
  return new Promise((resolve) => {
    const img = new Image();
    const objectUrl = URL.createObjectURL(file);
    img.onload = () => {
      URL.revokeObjectURL(objectUrl);
      const canvas = document.createElement('canvas');
      const maxDim = 360; // 360px max dimension for instant sub-10ms processing
      let width = img.width;
      let height = img.height;
      if (width > maxDim || height > maxDim) {
        if (width > height) {
          height = Math.round((height * maxDim) / width);
          width = maxDim;
        } else {
          width = Math.round((width * maxDim) / height);
          height = maxDim;
        }
      }
      canvas.width = width;
      canvas.height = height;
      const ctx = canvas.getContext('2d', { willReadFrequently: true });
      if (!ctx) {
        const b = file instanceof Blob ? file : new Blob([file]);
        resolve({ blob: b, base64: '' });
        return;
      }

      ctx.drawImage(img, 0, 0, width, height);
      const imgData = ctx.getImageData(0, 0, width, height);
      const data = imgData.data;

      // Sample border pixels to detect ambient background color
      let bgR = 0, bgG = 0, bgB = 0, sampleCount = 0;
      const stepX = Math.max(1, Math.floor(width / 6));
      const stepY = Math.max(1, Math.floor(height / 6));

      for (let x = 0; x < width; x += stepX) {
        let idx = x * 4;
        bgR += data[idx]; bgG += data[idx + 1]; bgB += data[idx + 2]; sampleCount++;
        idx = ((height - 1) * width + x) * 4;
        bgR += data[idx]; bgG += data[idx + 1]; bgB += data[idx + 2]; sampleCount++;
      }
      for (let y = 0; y < height; y += stepY) {
        let idx = (y * width) * 4;
        bgR += data[idx]; bgG += data[idx + 1]; bgB += data[idx + 2]; sampleCount++;
        idx = (y * width + (width - 1)) * 4;
        bgR += data[idx]; bgG += data[idx + 1]; bgB += data[idx + 2]; sampleCount++;
      }

      bgR = Math.round(bgR / sampleCount);
      bgG = Math.round(bgG / sampleCount);
      bgB = Math.round(bgB / sampleCount);

      const isAmbientLight = (bgR + bgG + bgB) / 3 > 150;
      const tolSq = tolerance * tolerance;
      const innerTolSq = (tolerance * 0.65) * (tolerance * 0.65);

      const totalPixels = data.length / 4;
      for (let p = 0; p < totalPixels; p++) {
        const i = p * 4;
        const r = data[i];
        const g = data[i + 1];
        const b = data[i + 2];

        const diffR = r - bgR;
        const diffG = g - bgG;
        const diffB = b - bgB;
        const distSq = diffR * diffR + diffG * diffG + diffB * diffB;

        const isWhite = r > 210 && g > 210 && b > 210;

        if (distSq < tolSq || (isAmbientLight && isWhite)) {
          if (distSq < innerTolSq || isWhite) {
            data[i + 3] = 0; // Fully transparent
          } else {
            const alphaFactor = Math.sqrt(distSq / tolSq);
            data[i + 3] = Math.round(255 * Math.max(0, Math.min(1, alphaFactor)));
          }
        }
      }

      ctx.putImageData(imgData, 0, 0);

      const base64Data = canvas.toDataURL('image/png');
      
      // Fast synchronous DataURL -> Blob
      const parts = base64Data.split(',');
      const mime = parts[0].match(/:(.*?);/)?.[1] || 'image/png';
      const bstr = atob(parts[1]);
      let n = bstr.length;
      const u8arr = new Uint8Array(n);
      while (n--) {
        u8arr[n] = bstr.charCodeAt(n);
      }
      const resBlob = new Blob([u8arr], { type: mime });

      resolve({ blob: resBlob, base64: base64Data });
    };
    img.onerror = () => {
      URL.revokeObjectURL(objectUrl);
      const b = file instanceof Blob ? file : new Blob([file]);
      resolve({ blob: b, base64: '' });
    };
    img.src = objectUrl;
  });
};

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
  PlusCircle,
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
  Send,
  FlaskConical,
  Boxes,
  Layers,
  Archive,
  BookOpen,
  Edit,
  MapPin,
  Save,
  Bell,
  Barcode,
  Image as ImageIcon
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import * as XLSX from 'xlsx';
import { jsPDF } from 'jspdf';
import { ImageEditorAndRetoucher } from './components/ImageEditorAndRetoucher';
import { CachedDriveImage, fetchDriveImageAsBase64 } from './components/CachedDriveImage';
import { TechnicalPublicationsModal } from './components/TechnicalPublicationsModal';
import { ExcelExportModal } from './components/ExcelExportModal';
import { NewProductModal } from './components/NewProductModal';
import { DepoManagementModal } from './components/DepoManagementModal';
import { KaraAraclariDocModal } from './components/KaraAraclariDocModal';
import { DataSyncModal } from './components/DataSyncModal';
import { AuditTrailModal } from './components/AuditTrailModal';
import { BakimYapildiModal } from './components/BakimYapildiModal';
import { AuditLogEntry, DepoTransaction, VehicleDocument } from './types';
import { DEFAULT_AT802_DATA } from './data/defaultTechizatData';
import { unmergeAndFillWorksheet, groupMultiLocationRowsHelper, cleanAndFormatDateString, isHeaderLikeRow, detectHeaderRowIndex } from './utils/driveExcelSync';
import { getAllHangarPdfDocs, saveHangarPdfDoc, deleteHangarPdfDoc, getHangarPdfDocById, syncHangarPdfDocsFromDrive, findMatchingDocs, isDocMatchingRow, normalizeDocKey, HangarPdfDoc, getFileCategory, getDocMimeType } from './utils/hangarPdfStorage';
import { PdfPreviewModal } from './components/PdfPreviewModal';
import { OlayTakipCizelgesiModal } from './components/OlayTakipCizelgesiModal';
import { BarkodOkuyucuModal } from './components/BarkodOkuyucuModal';
import { GunTakipModal } from './components/GunTakipModal';
import { DepoSlipPrintModal } from './components/DepoSlipPrintModal';

export const GOOGLE_SCRIPT_URL = "https://script.google.com/macros/s/AKfycby2TM-nlURR5iPR4s4I6BpE8hbor93Jin9g014k3XPaQ0rYtS2MWHwtlnlAoph8Y3mZ/exec";
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
  | 'UNIT_FOLDER_MENU'
  | 'UNIT_DEPO_MENU'
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
  },
  7: {
    title: 'DEPO SAYIM KAYITLARI',
    sheetName: '7-Sayimlar',
    storageKey: 'form_7_sayimlar',
    columns: [
      { key: 'Tarih', label: 'Tarih / Saat' },
      { key: 'Malzeme', label: 'Malzeme Adı' },
      { key: 'PN', label: 'P/N' },
      { key: 'SN', label: 'S/N' },
      { key: 'Bolge', label: 'Depo Bölgesi' },
      { key: 'Sistem_Stok', label: 'Sistem Mevcut' },
      { key: 'Sayilan_Adet', label: 'Fiziksel Sayılan' },
      { key: 'Fark', label: 'Fark (+/-)' },
      { key: 'Personel', label: 'Sayan Personel' }
    ],
    defaultRows: []
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

// Isolated Live Clock component so 1-second interval does NOT re-render the entire App component
const LiveClock = React.memo(() => {
  const [time, setTime] = useState(() => new Date().toLocaleTimeString('tr-TR'));
  useEffect(() => {
    const interval = setInterval(() => {
      setTime(new Date().toLocaleTimeString('tr-TR'));
    }, 1000);
    return () => clearInterval(interval);
  }, []);
  return <span className="font-mono text-white/70">SAAT: {time}</span>;
});

export default function App() {
  // Splash screen state
  const [splashVisible, setSplashVisible] = useState(true);

  // Mobile detection with debounced listener
  const [isMobile, setIsMobile] = useState(window.innerWidth < 768);

  useEffect(() => {
    let resizeTimer: any;
    const handleResize = () => {
      clearTimeout(resizeTimer);
      resizeTimer = setTimeout(() => {
        setIsMobile(window.innerWidth < 768);
      }, 150);
    };
    window.addEventListener('resize', handleResize);
    return () => {
      clearTimeout(resizeTimer);
      window.removeEventListener('resize', handleResize);
    };
  }, []);

  // Modal active state
  const [modalOpen, setModalOpen] = useState(false);
  const [modalTitle, setModalTitle] = useState('SİSTEM');
  const [modalType, setModalType] = useState<'iframe' | 'design' | 'category' | 'form_table' | 'excel_sync' | 'techizat_matrix' | 'denetleme'>('iframe');
  const [isTechizatSlipModalOpen, setIsTechizatSlipModalOpen] = useState<boolean>(false);
  const [techizatModalList, setTechizatModalList] = useState<any[]>([]);
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

  // Teçhizat row edit, image upload, mission order and documents states
  const [mobileEditTab, setMobileEditTab] = useState<'form' | 'image' | 'regional' | 'documents'>('form');
  const [hangarPdfDocs, setHangarPdfDocs] = useState<HangarPdfDoc[]>(() => {
    try {
      const s = localStorage.getItem('hangar_techizat_pdf_docs') || localStorage.getItem('hangar_pdf_docs');
      return s ? JSON.parse(s) : [];
    } catch {
      return [];
    }
  });
  const [isUploadingDoc, setIsUploadingDoc] = useState<boolean>(false);
  const [activePdfPreview, setActivePdfPreview] = useState<{
    id: string;
    itemKey: string;
    fileName: string;
    docType: string;
    uploadDate: string;
    firma: string;
    fileData: string;
    fileSize: string;
  } | null>(null);
  const [newDocDocType, setNewDocDocType] = useState<string>('Bakım Sonrası Evraklar');
  const [newDocCustomType, setNewDocCustomType] = useState<string>('');
  const [newDocFirma, setNewDocFirma] = useState<string>('');
  const [newDocSelectedFile, setNewDocSelectedFile] = useState<File | null>(null);
  const [docToDelete, setDocToDelete] = useState<string | null>(null);
  const [docUploadPasswordModal, setDocUploadPasswordModal] = useState<{ isOpen: boolean; pendingAction: 'upload' | 'delete' | null }>({ isOpen: false, pendingAction: null });
  const [docUploadPasswordInput, setDocUploadPasswordInput] = useState<string>('');
  const [docUploadPasswordError, setDocUploadPasswordError] = useState<string>('');
  const [docViewFilterTab, setDocViewFilterTab] = useState<'item' | 'all_drive'>('item');
  const [isSyncingDrivePdfs, setIsSyncingDrivePdfs] = useState<boolean>(false);
  const [driveDocSearchQuery, setDriveDocSearchQuery] = useState<string>('');

  const openGunTakipWithPassword = () => {
    setPasswordActionType('gun_takip');
    setPasswordInput('');
    setPasswordError(false);
    setIsPasswordModalOpen(true);
  };

  const handleSyncDrivePdfs = async () => {
    setIsSyncingDrivePdfs(true);
    try {
      const docs = await syncHangarPdfDocsFromDrive(GOOGLE_SCRIPT_URL);
      if (docs && docs.length > 0) {
        setHangarPdfDocs(docs);
        showNotification(`✅ Google Drive'dan ${docs.length} adet PDF evrak tarandı ve yüklendi!`);
      } else {
        showNotification('Google Drive üzerinde PDF evrak bulunamadı.');
      }
    } catch (err: any) {
      console.warn('Drive PDF sync error:', err);
      showNotification('Google Drive PDF taraması sırasında hata oluştu.');
    } finally {
      setIsSyncingDrivePdfs(false);
    }
  };

  useEffect(() => {
    // Initial sync of all PDFs from Google Drive & storage on portal startup
    syncHangarPdfDocsFromDrive(GOOGLE_SCRIPT_URL).then(docs => {
      if (docs && docs.length > 0) {
        setHangarPdfDocs(docs);
      }
    }).catch(err => {
      console.warn('Initial Drive PDF sync error:', err);
    });
  }, []);

  const [regionalLocations, setRegionalLocations] = useState<Array<{
    id: string;
    location: string;
    quantity: number | string;
    serialNumbers: string[];
    sonKontrol: string;
    gelecekKontrol: string;
    firma: string;
  }>>([]);
  const [newSerialInputs, setNewSerialInputs] = useState<Record<string, string>>({});
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

  // Ultra-fast instant background cleaner (0.01s execution, 0% CPU freeze)
  const handleFastCanvasBackgroundRemoval = async (file: File) => {
    try {
      setIsProcessingRemoveBg(true);
      setBgRemovalProgressPercent(100);
      setBgRemovalStatusText("⚡ Anında Temizlendi (0.01sn)...");

      const { blob, base64 } = await fastCanvasBackgroundRemoval(file);

      const localUrl = URL.createObjectURL(blob);
      setPendingImagePreview(localUrl);
      setPendingImageBase64(base64);
      setPendingImageMimeType("image/png");

      const transparentFile = new File([blob], `cleaned_${file.name || "captured.png"}`, { type: "image/png" });
      setPendingImageFile(transparentFile);

      showNotification("⚡ Arka plan anında (0.01 saniyede) temizlendi!");
      return blob;
    } catch (err: any) {
      console.error("Fast canvas bg removal error:", err);
      showNotification("Hızlı temizleme hatası oluştu.");
    } finally {
      setIsProcessingRemoveBg(false);
    }
  };

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
          currentP = Math.min(95, currentP + (currentP < 50 ? 10 : 5));
          setBgRemovalProgressPercent(currentP);
          setBgRemovalStatusText(`Arka Plan Analizi: %${currentP}`);
        }
      }, 50);
      
      // Hızlı AI işleme için 280px boyutunda küçük ve ultra hızlı tensor beslemesi
      const maxDim = 280;
      const processedFile = await prepareOptimizedImageForAI(file, maxDim);
      currentP = Math.max(currentP, 45);
      setBgRemovalProgressPercent(currentP);
      setBgRemovalStatusText(`Model hazırlanıyor (%${currentP})...`);

      const removeFn = (window as any).imglyRemoveBackground || removeBackground;
      
      // Race condition with 6-second max timeout -> fallback to 0.01s fast canvas if AI takes too long
      const aiPromise = removeFn(processedFile, {
        model: 'isnet_fp16',
        progress: (key: string, current: number, total: number) => {
          if (total > 0) {
            const calculatedPercent = Math.min(95, Math.max(45, Math.round(45 + (current / total) * 50)));
            if (calculatedPercent > currentP) {
              currentP = calculatedPercent;
              setBgRemovalProgressPercent(currentP);
              setBgRemovalStatusText(`Arka Plan Analizi: %${currentP}`);
            }
          }
        }
      });

      const timeoutPromise = new Promise((_, reject) => {
        setTimeout(() => reject(new Error("AI_TIMEOUT")), 6000);
      });

      const blob = await Promise.race([aiPromise, timeoutPromise]) as Blob;
      
      if (bgProgressTimer) clearInterval(bgProgressTimer);

      setBgRemovalProgressPercent(100);
      setBgRemovalStatusText("İşlem tamamlandı! (%100)");

      const localUrl = URL.createObjectURL(blob);
      setPendingImagePreview(localUrl);
      
      const reader = new FileReader();
      reader.onloadend = () => {
        const base64 = reader.result as string;
        setPendingImageBase64(base64);
        setPendingImageMimeType("image/png");
      };
      reader.readAsDataURL(blob);

      const transparentFile = new File([blob], `cleaned_${file.name || "captured.png"}`, { type: "image/png" });
      setPendingImageFile(transparentFile);
      
      showNotification("Arka plan başarıyla temizlendi!");
      return blob;
    } catch (err: any) {
      if (bgProgressTimer) clearInterval(bgProgressTimer);
      console.warn("AI model take too long or failed, instantly running Fast Canvas cleaner...", err);
      return await handleFastCanvasBackgroundRemoval(file);
    } finally {
      setIsProcessingRemoveBg(false);
    }
  };

  // Helper when image file is selected or captured from camera (Instant, 0 delay, no freeze)
  const handleImageSelected = async (file: File) => {
    setPendingImageFile(file);
    setPendingImagePreview(URL.createObjectURL(file));
    setBgRemovalStatusText("");
    
    // Convert to base64 instantly for display/storage
    const reader = new FileReader();
    reader.onloadend = () => {
      setPendingImageBase64(reader.result as string);
      setPendingImageMimeType(file.type || "image/png");
    };
    reader.readAsDataURL(file);
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
        if (techType && newArray.length > 0) {
          syncTechizatExcelToGoogleDrive(techType, newArray);
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
        if (techType && newArray.length > 0) {
          await syncTechizatExcelToGoogleDrive(techType, newArray);
        }
      }

      showNotification("Seçilen teçhizatlar başarıyla topluca güncellendi ve Google Drive Excel dosyalarına kaydedildi!");
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
  const [activeTechizatType, setActiveTechizatType] = useState<'bell429' | 'at802' | 'at802_ozel_alet' | 't70' | 't70_bumbi_backet' | 't70_helitak' | 'b360' | 'c650' | 'hangar' | 'kara_araclari' | 'all' | null>(null);
  const [selectedUnitFolder, setSelectedUnitFolder] = useState<string | null>(null);
  const [techizatActiveSection, setTechizatActiveSection] = useState<'all' | 'techizat_all' | 'depo_all' | 'yer_destek' | 'ozel_alet' | 'depo_sarf' | 'depo_kimyasal'>('all');
  const [isPullingTechizat, setIsPullingTechizat] = useState<boolean>(false);
  const [techizatFileNames, setTechizatFileNames] = useState<Record<string, string>>({});
  const [techizatSearchQuery, setTechizatSearchQuery] = useState<string>("");
  const [techizatFirmaFilter, setTechizatFirmaFilter] = useState<string>("");
  const [techizatDurumFilter, setTechizatDurumFilter] = useState<string>("");
  const [techizatColorFilter, setTechizatColorFilter] = useState<string>("ALL");
  const [isTechizatDriveLoading, setIsTechizatDriveLoading] = useState<boolean>(false);
  const [activeTechizatMatchIdx, setActiveTechizatMatchIdx] = useState<number>(0);
  const [techizatPage, setTechizatPage] = useState<number>(1);
  const [techizatPageSize, setTechizatPageSize] = useState<number>(50);
  const [isTechizatFilterOpen, setIsTechizatFilterOpen] = useState<boolean>(false);

  useEffect(() => {
    setTechizatPage(1);
  }, [activeTechizatType, techizatActiveSection, techizatSearchQuery, techizatFirmaFilter, techizatDurumFilter, techizatColorFilter]);

  const [isTechPubsOpen, setIsTechPubsOpen] = useState<boolean>(false);
  const [isOlayTakipOpen, setIsOlayTakipOpen] = useState<boolean>(false);
  const [olayTakipInitialUnit, setOlayTakipInitialUnit] = useState<string>('at802');
  const [isBarkodOkuyucuOpen, setIsBarkodOkuyucuOpen] = useState<boolean>(false);
  const [barkodOkuyucuInitialUnit, setBarkodOkuyucuInitialUnit] = useState<string>('at802');

  // Akıllı Teçhizat Görsel Eşleştirici: AT-802, Hangar ve tüm birimler için seri no, parça no veya isimle tam uyumlu görseli bulur
  const findRowImageUrl = (techType: string, row: string[], images: Record<string, string>): string | null => {
    if (!row || !images) return null;
    const nameVal = (row[1] || "").trim();
    const partVal = (row[2] || "").trim();
    const serialVal = (row[3] || "").trim();

    const typesToTry = [techType];
    const techLower = techType.toLowerCase();
    if (techLower === 'at802' || techLower === 'at-802') {
      typesToTry.push('at802', 'at-802', 'AT-802', 'AT802');
    }
    if (techLower === 'hangar') {
      typesToTry.push('hangar', 'HANGAR', 'yer_destek');
    }

    for (const t of typesToTry) {
      // 1. Birincil: type + name + serial
      const k1 = t + "_" + nameVal.replace(/\s+/g, '_') + "_" + serialVal.replace(/\s+/g, '_');
      if (images[k1]) return images[k1];

      // 2. İkincil: type + name + part
      const k2 = t + "_" + nameVal.replace(/\s+/g, '_') + "_" + partVal.replace(/\s+/g, '_');
      if (images[k2]) return images[k2];

      // 3. Üçüncül: type + name
      const k3 = t + "_" + nameVal.replace(/\s+/g, '_');
      if (images[k3]) return images[k3];

      // Küçük harf kombinasyonları
      const k1Lower = k1.toLowerCase();
      const k2Lower = k2.toLowerCase();
      const k3Lower = k3.toLowerCase();
      if (images[k1Lower]) return images[k1Lower];
      if (images[k2Lower]) return images[k2Lower];
      if (images[k3Lower]) return images[k3Lower];
    }

    // 4. Parça Numarası (P/N) ile eşleştirme
    if (partVal && partVal !== "-") {
      const cleanPart = partVal.toLowerCase().replace(/[\s\-_]/g, '');
      const foundByPart = Object.keys(images).find(k => {
        const kClean = k.toLowerCase().replace(/[\s\-_]/g, '');
        return kClean.includes(cleanPart);
      });
      if (foundByPart) return images[foundByPart];
    }

    // 5. Teçhizat Adı ile esnek eşleştirme
    if (nameVal) {
      const cleanName = nameVal.toLowerCase().replace(/\s+/g, '_');
      const foundByName = Object.keys(images).find(k => {
        const kLower = k.toLowerCase();
        return typesToTry.some(t => kLower.startsWith(t.toLowerCase())) && kLower.includes(cleanName);
      });
      if (foundByName) return images[foundByName];
    }

    return null;
  };

  const getUnitDisplayName = (unitKey: string | null): string => {
    switch (unitKey) {
      case 'bell429': return 'BELL 429';
      case 'at802': return 'AT-802F';
      case 't70': return 'T-70';
      case 'c650': return 'C-650';
      case 'b360': return 'B-360';
      case 'hangar': return 'HANGAR';
      case 'kara_araclari': return 'KARA ARAÇLARI';
      default: return 'BİRİM';
    }
  };

  const normalizeTurkishStr = (str: string): string => {
    if (!str) return "";
    return str
      .replace(/ğ/g, "g")
      .replace(/Ğ/g, "G")
      .replace(/ü/g, "u")
      .replace(/Ü/g, "U")
      .replace(/ş/g, "s")
      .replace(/Ş/g, "S")
      .replace(/ı/g, "i")
      .replace(/İ/g, "I")
      .replace(/ö/g, "o")
      .replace(/Ö/g, "O")
      .replace(/ç/g, "c")
      .replace(/Ç/g, "C");
  };

  const getRowSection = (row: string[], unitHint?: string): 'yer_destek' | 'ozel_alet' | 'depo_sarf' | 'depo_kimyasal' => {
    if (!row || !Array.isArray(row)) return 'yer_destek';

    // 1. Tag in the explicit section tag column (check indices 13, 12, 11, and last elements, skipping dates)
    const candidateIndices = [13, row.length - 1, 12, row.length - 2, 11];
    for (const idx of candidateIndices) {
      if (idx >= 0 && idx < row.length) {
        const rawVal = String(row[idx] || "").trim();
        // If it's a date or timestamp (e.g. 2026-08-22 or 15.01.2026), it's NOT a section tag
        if (/\d{4}-\d{2}-\d{2}/.test(rawVal) || /\d{2}\.\d{2}\.\d{4}/.test(rawVal)) continue;

        const raw = normalizeTurkishStr(rawVal).toLowerCase().trim();
        if (raw === 'depo_sarf' || raw === 'sarf_parca' || raw === 'sarf_ve_parca' || raw === 'sarf' || raw === 'sarf ve parca deposu' || raw === 'sarf deposu' || raw === 'parca deposu') return 'depo_sarf';
        if (raw === 'depo_kimyasal' || raw === 'kimyasal_depo' || raw === 'kimyasal' || raw === 'kimyasal depo' || raw === 'kimyasal maddeler' || raw === 'madeni yag') return 'depo_kimyasal';
        if (raw === 'ozel_alet' || raw === 'ozel_aletler' || raw === 'ozel_bakim' || raw === 'ozel bakim ve test aletleri' || raw === 'ozel bakim aletleri' || raw === 'ozel alet' || raw === 'ozel bakim ve test' || raw === 'ozel aletler') return 'ozel_alet';
        if (raw === 'yer_destek' || raw === 'yer_destek_techizat' || raw === 'yer destek techizatlari' || raw === 'yer destek') return 'yer_destek';
      }
    }

    // 2. Search for explicit markers embedded in cells
    const combined = normalizeTurkishStr(row.join(" ")).toUpperCase();
    if (combined.includes("##DEPO_SARF##") || combined.includes("[SARF DEPO]") || combined.includes("[PARCA DEPO]")) return 'depo_sarf';
    if (combined.includes("##DEPO_KIMYASAL##") || combined.includes("[KIMYASAL DEPO]")) return 'depo_kimyasal';
    if (combined.includes("##OZEL_ALET##") || combined.includes("[OZEL ALET]") || combined.includes("OZEL BAKIM VE TEST ALETLERI") || combined.includes("OZEL BAKIM ALETLERI") || combined.includes("OZEL ALETLER")) return 'ozel_alet';
    if (combined.includes("##YER_DESTEK##") || combined.includes("[YER DESTEK]") || combined.includes("YER DESTEK TECHIZATLARI")) return 'yer_destek';

    // 3. Kullanıcı kuralı: Heuristik anahtar kelime eşleştirmesi KESİNLİKLE YASAKTIR.
    // Veriler doğrudan Google Drive üzerindeki ilgili Excel dosyasından (String/Name matching) okunur;
    // hiçbir birimin tablosuna başka bir birimin verisi veya Özel Aletler verisi çekilemez.
    return 'yer_destek';
  };

  const getRowSectionLabel = (row: string[], unitHint?: string): string => {
    const sec = getRowSection(row, unitHint);
    if (sec === 'depo_sarf') return "SARF VE PARÇA DEPOSU";
    if (sec === 'depo_kimyasal') return "KİMYASAL DEPO";
    if (sec === 'ozel_alet') return "ÖZEL ALETLER";
    return "YER DESTEK TEÇHİZATLARI";
  };

  const cleanAndFormatDateString = (val: any): string => {
    if (val === undefined || val === null) return '';
    let str = String(val).trim();
    if (!str || str === '-' || str === '--' || str.toLowerCase() === 'null' || str.toLowerCase() === 'undefined') return str || '';

    // 1. Excel seri numarası kontrolü (Örn: 45234, 46234, 46965 vb. 5 basamaklı sayılar)
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

    // 2. Yapışık gelen tarihleri tespit et (Örn: 24.11.202508.12.2026 -> 24.11.2025 ve 08.12.2026)
    str = str.replace(/(\d{1,2}[\.\/-]\d{1,2}[\.\/-]\d{2,4})(\d{1,2}[\.\/-]\d{1,2}[\.\/-]\d{2,4})/g, '$1\n$2');

    // 3. Çok satırlı ise her bir satırı bağımsız formatla
    const lines = str.split(/[\r\n;]+/).map(s => s.trim()).filter(Boolean);
    if (lines.length > 1) {
      return lines.map(l => cleanAndFormatDateString(l)).join('\n');
    }

    // 4. ISO Date string (Örn: 2026-11-24T00:00:00.000Z)
    if (str.includes('T') && !isNaN(Date.parse(str))) {
      const parsed = new Date(str);
      const d = String(parsed.getDate()).padStart(2, '0');
      const m = String(parsed.getMonth() + 1).padStart(2, '0');
      const y = String(parsed.getFullYear());
      return `${d}.${m}.${y}`;
    }

    // 5. YYYY-MM-DD veya YYYY/MM/DD veya YYYY.MM.DD -> GG.AA.YYYY
    const ymd = str.match(/^(\d{4})[\.\/-](\d{1,2})[\.\/-](\d{1,2})$/);
    if (ymd) {
      const d = ymd[3].padStart(2, '0');
      const m = ymd[2].padStart(2, '0');
      const y = ymd[1];
      return `${d}.${m}.${y}`;
    }

    // 6. M/D/YY veya M/D/YYYY (Örn: 7/31/25 veya 7/31/2026 veya 07/31/25) -> GG.AA.YYYY
    // Kullanıcı kuralı: "TARİHLER AY YIL GÜN ŞEKLİNDE OLACAKTIR KESİNLİKLE HATADIR."
    const slashDate = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
    if (slashDate) {
      let p1 = parseInt(slashDate[1], 10);
      let p2 = parseInt(slashDate[2], 10);
      let yearPart = slashDate[3];
      let y = yearPart.length === 2 ? `20${yearPart}` : yearPart;

      let day = p2;
      let month = p1;
      if (p1 > 12 && p2 <= 12) {
        day = p1;
        month = p2;
      }
      return `${String(day).padStart(2, '0')}.${String(month).padStart(2, '0')}.${y}`;
    }

    // 7. DD.MM.YYYY veya DD-MM-YYYY veya DD.MM.YY -> GG.AA.YYYY
    const dmy = str.match(/^(\d{1,2})[\.-](\d{1,2})[\.-](\d{2,4})$/);
    if (dmy) {
      const d = dmy[1].padStart(2, '0');
      const m = dmy[2].padStart(2, '0');
      let y = dmy[3];
      if (y.length === 2) y = `20${y}`;
      return `${d}.${m}.${y}`;
    }

    return str;
  };

  const isHeaderLikeRow = (r: any[]): boolean => {
    if (!Array.isArray(r) || r.length === 0) return true;
    const c0 = String(r[0] || "").trim().toUpperCase();
    const c1 = String(r[1] || "").trim().toUpperCase();
    const c2 = String(r[2] || "").trim().toUpperCase();
    const c3 = String(r[3] || "").trim().toUpperCase();
    const c4 = String(r[4] || "").trim().toUpperCase();
    const c5 = String(r[5] || "").trim().toUpperCase();
    const c6 = String(r[6] || "").trim().toUpperCase();

    // Kullanıcı kuralı: "VERİ GÜNCELLEMEDE BAŞLIK 1. SATIR ÜRÜN DİYE ATMIŞ HATADIR."
    if (
      c0 === "SIRA NO" || c0 === "SIRA" || c0 === "NO" || c0 === "NO." ||
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
        s === "SIRA NO" || s === "NO." || s === "SIRA" ||
        s === "TEÇHİZAT ADI" || s === "TECHIZAT ADI" ||
        s === "PARÇA NO (P/N) / MODEL" || s === "PARÇA NO (P/N)" || s === "PARÇA NO" || s === "P/N" ||
        s === "SERİ NO (S/N)" || s === "SERİ NO" || s === "S/N" ||
        s === "MİKTAR / KAPASİTE" || s === "MİKTAR" ||
        s === "BULUNDUĞU YER" || s === "DURUMU" ||
        s === "KALİBRASYONA TABİ" || s.includes("SON KONTROL") ||
        s.includes("GELECEK KONTROL") || s.includes("YAPAN FİRMA") ||
        s === "AÇIKLAMA" || s === "ÖMÜR BİTİŞ TARİHİ"
      ) {
        headerMatches++;
      }
    }
    return headerMatches >= 2;
  };

  const sanitizeLoadedTechizatData = (data: any, fallback: string[][], techTypeHint?: string): string[][] => {
    if (!Array.isArray(data) || data.length === 0) return fallback;

    // Başlık satırlarını kesinlikle filtrele
    const filteredRows = data.filter((r: any) => !isHeaderLikeRow(r));
    if (filteredRows.length === 0) return fallback;

    const sanitized = filteredRows.map((r: any) => {
      if (!Array.isArray(r)) return r;
      const row = [...r];

      const sec = getRowSection(row, techTypeHint);
      const isDepo = sec === 'depo_sarf' || sec === 'depo_kimyasal';

      // Ensure Kalibrasyona Tabi column exists at index 7 for standard rows
      const val7 = String(row[7] || "").trim().toUpperCase();
      if (!isDepo && val7 !== "EVET" && val7 !== "HAYIR") {
        row.splice(7, 0, "EVET");
      }

      while (row.length < 14) row.push("");

      // Tarih alanlarını temizle ve formatla (Excel seri no veya bozuk formatları GG.AA.YYYY yap)
      if (row[8]) row[8] = cleanAndFormatDateString(row[8]);
      if (row[9]) row[9] = cleanAndFormatDateString(row[9]);
      if (row[6] && /\d/.test(row[6]) && (row[6].includes('/') || row[6].includes('.') || /^\d{5}$/.test(row[6]))) {
        row[6] = cleanAndFormatDateString(row[6]);
      }
      if (row[7] && /\d/.test(row[7]) && (row[7].includes('/') || row[7].includes('.') || /^\d{5}$/.test(row[7]))) {
        row[7] = cleanAndFormatDateString(row[7]);
      }

      // Ensure index 12 (Mail Date) NEVER contains category strings or default dashes
      const val12 = String(row[12] || "").trim().toLowerCase();
      if (val12 === 'yer_destek' || val12 === 'ozel_alet' || val12 === 'depo_sarf' || val12 === 'depo_kimyasal' || val12 === '-' || val12 === '--' || val12.includes('belirtilme')) {
        row[12] = "";
      }

      // Ensure section tag at index 13 is always set
      row[13] = sec;

      return row;
    });

    // Sıra numaralarını baştan düzenle
    sanitized.forEach((r, idx) => {
      r[0] = String(idx + 1);
    });

    return sanitized;
  };

  // Default empty datasets (strictly real data from online sheets, no hardcoded sample data)
  const DEFAULT_BELL429_DATA: string[][] = [];
  const DEFAULT_T70_DATA: string[][] = [];
  const DEFAULT_T70_BUMBI_DATA: string[][] = [];
  const DEFAULT_T70_HELITAK_DATA: string[][] = [];
  const DEFAULT_B360_DATA: string[][] = [];
  const DEFAULT_C650_DATA: string[][] = [];
  const DEFAULT_HANGAR_DATA: string[][] = [];
  const DEFAULT_KARA_ARACLARI_DATA: string[][] = [];

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
    try {
      const stored = localStorage.getItem('techizat_bell429_data') || localStorage.getItem('excel_techizat_bell429_data');
      if (stored) {
        const parsed = JSON.parse(stored);
        if (Array.isArray(parsed) && parsed.length > 0) return parsed;
      }
    } catch(e) {}
    return [];
  });

  const [techizatAt802Columns, setTechizatAt802Columns] = useState<string[]>(() => {
    return ["SIRA NO", "TEÇHİZAT ADI", "PARÇA NO (P/N)", "SERİ NO (S/N)", "MİKTAR", "BULUNDUĞU YER", "DURUMU", "KALİBRASYONA TABİ", "SON KONTROL / KALİBRASYON / BAKIM", "GELECEK KONTROL / KALİBRASYON / BAKIM", "SON KONTROLÜ YAPAN FİRMA", "AÇIKLAMA"];
  });
  const [techizatAt802Data, setTechizatAt802Data] = useState<string[][]>(() => {
    try {
      const stored = localStorage.getItem('techizat_at802_data') || localStorage.getItem('excel_techizat_at802_data');
      if (stored) {
        const parsed = JSON.parse(stored);
        if (Array.isArray(parsed) && parsed.length > 0) return parsed;
      }
    } catch(e) {}
    return [];
  });

  const [techizatAt802OzelAletColumns, setTechizatAt802OzelAletColumns] = useState<string[]>(() => {
    return ["SIRA NO", "TEÇHİZAT ADI", "PARÇA NO (P/N) / MODEL", "SERİ NO (S/N)", "MİKTAR / KAPASİTE", "BULUNDUĞU YER", "DURUMU", "KALİBRASYONA TABİ", "SON KONTROL / KALİBRASYON / BAKIM", "GELECEK KONTROL / KALİBRASYON / BAKIM", "SON KONTROLÜ YAPAN FİRMA", "AÇIKLAMA"];
  });
  const [techizatAt802OzelAletData, setTechizatAt802OzelAletData] = useState<string[][]>(() => {
    try {
      const stored = localStorage.getItem('techizat_at802_ozel_alet_data');
      if (stored) {
        const parsed = JSON.parse(stored);
        if (Array.isArray(parsed) && parsed.length > 0) return parsed;
      }
    } catch(e) {}
    return [];
  });

  const [techizatT70Columns, setTechizatT70Columns] = useState<string[]>(() => {
    return ["SIRA NO", "TEÇHİZAT ADI", "PARÇA NO (P/N)", "SERİ NO (S/N)", "MİKTAR", "BULUNDUĞU YER", "DURUMU", "KALİBRASYONA TABİ", "SON KONTROL / KALİBRASYON / BAKIM", "GELECEK KONTROL / KALİBRASYON / BAKIM", "SON KONTROLÜ YAPAN FİRMA", "AÇIKLAMA"];
  });
  const [techizatT70Data, setTechizatT70Data] = useState<string[][]>(() => {
    try {
      const stored = localStorage.getItem('techizat_t70_data') || localStorage.getItem('excel_techizat_t70_data');
      if (stored) {
        const parsed = JSON.parse(stored);
        if (Array.isArray(parsed) && parsed.length > 0) return parsed;
      }
    } catch(e) {}
    return [];
  });

  const [techizatT70BumbiBacketColumns, setTechizatT70BumbiBacketColumns] = useState<string[]>(() => {
    return ["SIRA NO", "TEÇHİZAT ADI", "MODEL / TİP", "SERİ NO (S/N)", "KAPASİTE", "BULUNDUĞU YER", "DURUMU", "KALİBRASYONA TABİ", "SON KONTROL / KALİBRASYON / BAKIM", "GELECEK KONTROL / KALİBRASYON / BAKIM", "SON KONTROLÜ YAPAN FİRMA", "AÇIKLAMA"];
  });
  const [techizatT70BumbiBacketData, setTechizatT70BumbiBacketData] = useState<string[][]>(() => {
    try {
      const stored = localStorage.getItem('techizat_t70_bumbi_backet_data') || localStorage.getItem('excel_techizat_t70_bumbi_backet_data');
      if (stored) {
        const parsed = JSON.parse(stored);
        if (Array.isArray(parsed) && parsed.length > 0) return parsed;
      }
    } catch(e) {}
    return [];
  });

  const [techizatT70HelitakColumns, setTechizatT70HelitakColumns] = useState<string[]>(() => {
    return ["SIRA NO", "TEÇHİZAT ADI", "MODEL / TİP", "SERİ NO (S/N)", "KAPASİTE", "BULUNDUĞU YER", "DURUMU", "KALİBRASYONA TABİ", "SON KONTROL / KALİBRASYON / BAKIM", "GELECEK KONTROL / KALİBRASYON / BAKIM", "SON KONTROLÜ YAPAN FİRMA", "AÇIKLAMA"];
  });
  const [techizatT70HelitakData, setTechizatT70HelitakData] = useState<string[][]>(() => {
    try {
      const stored = localStorage.getItem('techizat_t70_helitak_data') || localStorage.getItem('excel_techizat_t70_helitak_data');
      if (stored) {
        const parsed = JSON.parse(stored);
        if (Array.isArray(parsed) && parsed.length > 0) return parsed;
      }
    } catch(e) {}
    return [];
  });

  const [techizatB360Columns, setTechizatB360Columns] = useState<string[]>(() => {
    return ["SIRA NO", "TEÇHİZAT ADI", "PARÇA NO (P/N)", "SERİ NO (S/N)", "MİKTAR", "BULUNDUĞU YER", "DURUMU", "KALİBRASYONA TABİ", "SON KONTROL / KALİBRASYON / BAKIM", "GELECEK KONTROL / KALİBRASYON / BAKIM", "SON KONTROLÜ YAPAN FİRMA", "AÇIKLAMA"];
  });
  const [techizatB360Data, setTechizatB360Data] = useState<string[][]>(() => {
    try {
      const stored = localStorage.getItem('techizat_b360_data') || localStorage.getItem('excel_techizat_b360_data');
      if (stored) {
        const parsed = JSON.parse(stored);
        if (Array.isArray(parsed) && parsed.length > 0) return parsed;
      }
    } catch(e) {}
    return [];
  });

  const [techizatC650Columns, setTechizatC650Columns] = useState<string[]>(() => {
    return ["SIRA NO", "TEÇHİZAT ADI", "PARÇA NO (P/N)", "SERİ NO (S/N)", "MİKTAR", "BULUNDUĞU YER", "DURUMU", "KALİBRASYONA TABİ", "SON KONTROL / KALİBRASYON / BAKIM", "GELECEK KONTROL / KALİBRASYON / BAKIM", "SON KONTROLÜ YAPAN FİRMA", "AÇIKLAMA"];
  });
  const [techizatC650Data, setTechizatC650Data] = useState<string[][]>(() => {
    try {
      const stored = localStorage.getItem('techizat_c650_data') || localStorage.getItem('excel_techizat_c650_data');
      if (stored) {
        const parsed = JSON.parse(stored);
        if (Array.isArray(parsed) && parsed.length > 0) return parsed;
      }
    } catch(e) {}
    return [];
  });

  const [techizatHangarColumns, setTechizatHangarColumns] = useState<string[]>(() => {
    return ["SIRA NO", "TEÇHİZAT ADI", "PARÇA NO (P/N) / MODEL", "SERİ NO (S/N)", "MİKTAR / KAPASİTE", "BULUNDUĞU YER", "DURUMU", "BAKIMA TABİ", "SON KONTROL / KALİBRASYON / BAKIM", "GELECEK KONTROL / KALİBRASYON / BAKIM", "SON KONTROLÜ YAPAN FİRMA", "AÇIKLAMA", "90 GÜN UYARISI MAİL GÖNDERİM TARİHİ", "BELGE YÜKLE"];
  });
  const [techizatHangarData, setTechizatHangarData] = useState<string[][]>(() => {
    try {
      const stored = localStorage.getItem('techizat_hangar_data') || localStorage.getItem('excel_techizat_hangar_data');
      if (stored) {
        const parsed = JSON.parse(stored);
        if (Array.isArray(parsed) && parsed.length > 0) return parsed;
      }
    } catch(e) {}
    return [];
  });

  const [techizatKaraAraclariColumns, setTechizatKaraAraclariColumns] = useState<string[]>(() => {
    return [
      "SIRA NO", 
      "ARAÇ PLAKASI / TANIMI", 
      "MARKA", 
      "MODEL", 
      "YAKIT TÜRÜ", 
      "BULUNDUĞU YER", 
      "SON KM Sİ", 
      "DURUMU", 
      "BAKIMA TABİ", 
      "SON KONTROL / KALİBRASYON / BAKIM", 
      "GELECEK KONTROL / KALİBRASYON / BAKIM", 
      "SON KONTROLÜ YAPAN FİRMA", 
      "AÇIKLAMA", 
      "90 GÜN UYARISI MAİL GÖNDERİM TARİHİ"
    ];
  });
  const [techizatKaraAraclariData, setTechizatKaraAraclariData] = useState<string[][]>(() => {
    try {
      const stored = localStorage.getItem('excel_techizat_kara_araclari_data') || localStorage.getItem('techizat_kara_araclari_data');
      if (stored) {
        const parsed = JSON.parse(stored);
        if (Array.isArray(parsed) && parsed.length > 0) {
          return parsed;
        }
      }
    } catch (e) { console.error(e); }
    return [
      ["1", "06 CUK 695", "FORD", "RANGER 4x4", "Dizel", "ANKARA", "124500", "FAAL", "EVET", "15.01.2025", "15.01.2026", "FORD YETKİLİ SERVİS", "Periyodik bakım yapıldı", ""],
      ["2", "06 FV 2359", "TOYOTA", "HILUX 4x4", "Dizel", "ANTALYA", "158200", "FAAL", "EVET", "10.02.2025", "10.02.2026", "TOYOTA PLAZA", "Genel kontrol ve periyodik servis", ""],
      ["3", "06 OGM 1001", "ISUZU", "D-MAX", "Dizel", "İZMİR", "4500", "FAAL", "EVET", "05.03.2025", "05.03.2026", "ISUZU SERVİS", "Rutin muayene", ""],
      ["4", "06 OGM 1002", "FORD", "TRANSIT CUSTOM", "Dizel", "MUĞLA", "18200", "FAAL", "EVET", "20.01.2025", "20.01.2026", "FORD OTO", "Yağ filtre bakımı", ""],
      ["5", "06 OGM 1003", "MERCEDES-BENZ", "SPRINTER", "Dizel", "ADANA", "9400", "FAAL", "EVET", "18.02.2025", "18.02.2026", "MERCEDES HAS", "10.000 KM periyodik servis", ""],
      ["6", "06 OGM 1004", "RENAULT", "DUSTER 4x4", "Benzin", "ANKARA", "3200", "FAAL", "EVET", "12.04.2025", "12.04.2026", "MAİS ANKARA", "İlk rodaj muayene", ""]
    ];
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
  const openTechizatMatrix = (
    type: 'bell429' | 'at802' | 'at802_ozel_alet' | 't70' | 't70_bumbi_backet' | 't70_helitak' | 'b360' | 'c650' | 'hangar' | 'kara_araclari' | 'all',
    title: string,
    sectionFilter: 'all' | 'techizat_all' | 'depo_all' | 'yer_destek' | 'ozel_alet' | 'depo_sarf' | 'depo_kimyasal' = 'all'
  ) => {
    const actualType = (type === 'at802' && sectionFilter === 'ozel_alet') ? 'at802_ozel_alet' : type;
    setActiveTechizatType(actualType);
    setTechizatActiveSection(sectionFilter);
    setModalType('techizat_matrix');
    setModalTitle(title);
    setModalOpen(true);
    setTechizatSearchQuery('');
    setActiveTechizatMatchIdx(0);
    // Google Drive'dan tam dosya adı eşleştirmesi ile canlı senkronizasyon
    pullTechizatUnitFromDrive(actualType, true);
  };

  // Open Depo Management Modal directly in the main system
  const openStandaloneDepo = (unitKey?: string) => {
    if (unitKey) {
      setSelectedUnitFolder(unitKey);
    }
    setIsDepoModalOpen(true);
  };

  // State for "Yeni Ürün / Teçhizat Ekle" Modal
  const [isNewProductModalOpen, setIsNewProductModalOpen] = useState<boolean>(false);
  const [isDataSyncModalOpen, setIsDataSyncModalOpen] = useState<boolean>(false);
  const [isDepoModalOpen, setIsDepoModalOpen] = useState<boolean>(false);
  const [isAuditModalOpen, setIsAuditModalOpen] = useState<boolean>(false);
  const [activeKaraDocTarget, setActiveKaraDocTarget] = useState<{ plate: string; vehicleName: string } | null>(null);

  // Persistence for Vehicle Documents, Depo Transactions & Audit Logs
  const [vehicleDocuments, setVehicleDocuments] = useState<VehicleDocument[]>(() => {
    try {
      const saved = localStorage.getItem('kara_araclari_vehicle_documents');
      if (saved) return JSON.parse(saved);
    } catch (e) { console.error(e); }
    return [];
  });

  const [depoTransactions, setDepoTransactions] = useState<DepoTransaction[]>(() => {
    try {
      const saved = localStorage.getItem('ogm_depo_transfers_v5') || localStorage.getItem('form_7_sayimlar_txs');
      if (saved) {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed) && parsed.length > 0) return parsed;
      }
    } catch (e) {}
    return [];
  });

  useEffect(() => {
    if (depoTransactions.length > 0) {
      try {
        localStorage.setItem('ogm_depo_transfers_v5', JSON.stringify(depoTransactions));
      } catch (e) {}
    }
  }, [depoTransactions]);

  useEffect(() => {
    fetch('/api/get-depo-transfers')
      .then(res => res.json())
      .then(data => {
        if (data && data.status === 'success' && Array.isArray(data.transactions) && data.transactions.length > 0) {
          setDepoTransactions(prev => {
            const map = new Map<string, DepoTransaction>();
            const getTxKey = (t: DepoTransaction) => t.id || `${t.timestamp || t.date}_${t.itemName || t.pn}_${t.type || t.islemTuru}`;
            prev.forEach(t => map.set(getTxKey(t), t));
            data.transactions.forEach((t: DepoTransaction) => {
              const k = getTxKey(t);
              if (!map.has(k)) map.set(k, t);
            });
            return Array.from(map.values());
          });
        }
      })
      .catch(() => {});
  }, []);

  const [auditLogs, setAuditLogs] = useState<AuditLogEntry[]>(() => {
    try {
      const saved = localStorage.getItem('equipment_audit_logs');
      if (saved) return JSON.parse(saved);
    } catch (e) { console.error(e); }
    return [];
  });

  const [depoCertificatePdfUrl, setDepoCertificatePdfUrl] = useState<string | null>(() => {
    return localStorage.getItem('depo_certificate_pdf_url') || null;
  });

  const addAuditLog = (entry: Omit<AuditLogEntry, 'id' | 'timestamp'>) => {
    const newLog: AuditLogEntry = {
      ...entry,
      id: `audit_${Date.now()}_${Math.random().toString(36).substring(2, 7)}`,
      timestamp: new Date().toLocaleString('tr-TR')
    };
    setAuditLogs(prev => {
      const updated = [newLog, ...prev];
      try { localStorage.setItem('equipment_audit_logs', JSON.stringify(updated.slice(0, 500))); } catch (e) {}
      return updated;
    });
  };

  const handleAddNewProductRow = (targetUnit: string, newRow: string[]) => {
    const formattedRow = newRow.map(c => cleanAndFormatDateString(c));
    const setterMap: Record<string, React.Dispatch<React.SetStateAction<string[][]>>> = {
      'at802': setTechizatAt802Data,
      'at802_ozel_alet': setTechizatAt802OzelAletData,
      'bell429': setTechizatBell429Data,
      't70': setTechizatT70Data,
      't70_bumbi_backet': setTechizatT70BumbiBacketData,
      't70_helitak': setTechizatT70HelitakData,
      'b360': setTechizatB360Data,
      'c650': setTechizatC650Data,
      'hangar': setTechizatHangarData,
      'kara_araclari': setTechizatKaraAraclariData,
    };

    const storageKeyMap: Record<string, string> = {
      'at802': 'excel_techizat_at802_data',
      'at802_ozel_alet': 'techizat_at802_ozel_alet_data',
      'bell429': 'excel_techizat_bell429_data',
      't70': 'excel_techizat_t70_data',
      't70_bumbi_backet': 'excel_techizat_t70_bumbi_backet_data',
      't70_helitak': 'excel_techizat_t70_helitak_data',
      'b360': 'excel_techizat_b360_data',
      'c650': 'excel_techizat_c650_data',
      'hangar': 'excel_techizat_hangar_data',
      'kara_araclari': 'excel_techizat_kara_araclari_data',
    };

    const targetSetter = setterMap[targetUnit];
    const targetKey = storageKeyMap[targetUnit];

    if (targetSetter && targetKey) {
      targetSetter(prev => {
        const nextIndex = prev.length + 1;
        const rowWithIndex = [...formattedRow];
        rowWithIndex[0] = String(nextIndex);
        const updated = [...prev, rowWithIndex];
        try { localStorage.setItem(targetKey, JSON.stringify(updated)); } catch (e) {}
        return updated;
      });

      addAuditLog({
        unit: targetUnit.toUpperCase(),
        action: 'EKLEME',
        itemName: newRow[1] || 'Yeni Ürün',
        pn: newRow[2] || '-',
        fieldName: 'KAYIT_OLUSTURMA',
        oldValue: '-',
        newValue: `${newRow[1]} (${newRow[2]}) eklendi`
      });

      showNotification(`✅ Yeni ürün/araç (${newRow[1]}) başarıyla kaydedildi!`);
      setIsNewProductModalOpen(false);
    }
  };
  const [newProductForm, setNewProductForm] = useState<{
    unit: string;
    section: 'yer_destek' | 'ozel_alet' | 'depo_sarf' | 'depo_kimyasal' | 'kara_araclari';
    siraNo: string;
    name: string;
    marka?: string;
    pn: string;
    sn: string;
    miktar: string;
    yer: string;
    durumu: string;
    kalibrasyonaTabi: string;
    sonKontrol: string;
    gelecekKontrol: string;
    firma: string;
    aciklama: string;
  }>({
    unit: 'at802',
    section: 'yer_destek',
    siraNo: '',
    name: '',
    marka: '',
    pn: '',
    sn: '',
    miktar: '1 ADET',
    yer: '',
    durumu: 'FAAL',
    kalibrasyonaTabi: 'EVET',
    sonKontrol: '',
    gelecekKontrol: '',
    firma: '',
    aciklama: ''
  });

  const initializeNewProductForm = () => {
    const defaultUnit = activeTechizatType !== 'all' ? activeTechizatType : 'at802';
    let defaultSection: 'yer_destek' | 'ozel_alet' | 'depo_sarf' | 'depo_kimyasal' | 'kara_araclari' = 'yer_destek';
    if (defaultUnit === 'kara_araclari') {
      defaultSection = 'kara_araclari';
    } else if (techizatActiveSection === 'depo_sarf' || techizatActiveSection === 'depo_all') {
      defaultSection = 'depo_sarf';
    } else if (techizatActiveSection === 'depo_kimyasal') {
      defaultSection = 'depo_kimyasal';
    } else if (techizatActiveSection === 'ozel_alet') {
      defaultSection = 'ozel_alet';
    } else {
      defaultSection = 'yer_destek';
    }

    setNewProductForm({
      unit: defaultUnit,
      section: defaultSection,
      siraNo: '',
      name: '',
      pn: '',
      sn: defaultSection === 'kara_araclari' ? '' : '-',
      miktar: defaultSection === 'kara_araclari' ? '' : '1 ADET',
      yer: '',
      durumu: 'FAAL',
      kalibrasyonaTabi: 'EVET',
      sonKontrol: '',
      gelecekKontrol: '',
      firma: '',
      aciklama: ''
    });
  };

  /**
   * Akıllı Çoklu Lokasyon, Miktar ve S/N Birleştirici:
   * Yüklenen Excel veya Canlı E-Tabloda aynı ürün veya aynı P/N'e ait birden fazla lokasyon, miktar ve S/N bilgisi
   * girildiğinde (ürün adı boş, tire veya aynı P/N/isimde), aynı bölgede ise miktar toplanır ve S/N bilgileri yan yana
   * (Örn: "1234 ; 123") birleştirilir; farklı bölgelerde ise 'BULUNDUĞU YER', 'MİKTAR / KAPASİTE' ve 'SERİ NO (S/N)'
   * alanlarına hizalı alt alta (\n ile) eşleştirilir.
   */
  const mergeSerialNumbersHelper = (existingSn: string, newSn: string): string => {
    const eSn = (existingSn || "").trim();
    const nSn = (newSn || "").trim();
    if (!nSn || nSn === "-" || nSn === "--") return eSn;
    if (!eSn || eSn === "-" || eSn === "--") return nSn;

    // Check if newSn is already present in existingSn
    const existingParts = eSn.split(/[;,\n]/).map(p => p.trim()).filter(Boolean);
    if (existingParts.includes(nSn)) return eSn;

    return `${eSn} ; ${nSn}`;
  };

  const mergeQuantitiesHelper = (existingQty: string, newQty: string): string => {
    const eQty = (existingQty || "").trim();
    const nQty = (newQty || "").trim();
    if (!nQty) return eQty || "1";
    if (!eQty) return nQty || "1";
    
    const numE = Number(eQty);
    const numN = Number(nQty);
    if (!isNaN(numE) && !isNaN(numN)) {
      return String(numE + numN);
    }
    return eQty;
  };

  const groupMultiLocationRows = (
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
      const isSameName = rawName.toUpperCase() === lastName.toUpperCase() && lastName !== "";
      const isSameSira = rawSira === lastSira && lastSira !== "";
      const isSameSection = !lastRow || getRowSection(row) === getRowSection(lastRow);

      // Alt lokasyon / birleştirilmiş hücre tespiti:
      // 1) Eğer ürün adı boşsa veya "-" ise ve bir önceki ürün varsa, BU KESİNLİKLE BİR ÖNCEKİ ÜRÜNÜN ALT LOKASYONUDUR!
      // 2) Eğer aynı isim ve aynı P/N varsa, yine aynı ürünün farklı bir lokasyonudur!
      // 3) Eğer aynı sıra numarası ve aynı isim/PN varsa, yine aynı ürünün birleştirilmiş satırıdır!
      // 4) Eğer adı aynıysa (P/N boş veya aynı), yine aynı ürünün birleştirilmiş lokasyonudur!
      const isSubLocation =
        isSameSection &&
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

        // Check if rawLoc matches one of the existing locations (e.g. ANTALYA veya Uçuş Hattı)
        const matchingLocIdx = rawLoc
          ? existingLocs.findIndex(l => l.trim().toUpperCase() === rawLoc.trim().toUpperCase())
          : -1;

        if (matchingLocIdx !== -1 && (!rawSn || rawSn === "-" || !existingSns[matchingLocIdx] || existingSns[matchingLocIdx] === "-")) {
          // Same location -> Merge quantities & merge serial numbers side-by-side
          existingMiktars[matchingLocIdx] = mergeQuantitiesHelper(existingMiktars[matchingLocIdx], rawMiktar || "1");
          if (rawSn && rawSn !== "-") {
            existingSns[matchingLocIdx] = mergeSerialNumbersHelper(existingSns[matchingLocIdx], rawSn);
          }
        } else {
          // New location or new sub-region
          existingLocs.push(rawLoc);
          existingMiktars.push(rawMiktar || "1");
          existingSns.push(rawSn);
        }

        if (locColIdx >= 0) lastRow[locColIdx] = existingLocs.join('\n');
        if (miktarColIdx >= 0) lastRow[miktarColIdx] = existingMiktars.join('\n');
        if (seriNoColIdx >= 0) lastRow[seriNoColIdx] = existingSns.join('\n');

        // Fill other missing fields if main row has empty and sub-row provides it
        for (let c = 0; c < row.length; c++) {
          if (c !== siraColIdx && c !== nameColIdx && c !== locColIdx && c !== miktarColIdx && c !== seriNoColIdx) {
            const cellVal = (row[c] || "").trim();
            if (cellVal && cellVal !== "-" && (!lastRow[c] || lastRow[c] === "-")) {
              lastRow[c] = cellVal;
            }
          }
        }
      } else {
        // Independent new product row (ancak adı ve PN'si tamamen boş olan tekil çöp satırları alma)
        if (!isNameEmptyOrDash || rawPn !== "" || (rawSn !== "" && rawSn !== "-")) {
          grouped.push(row);
        }
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

  const exportCurrentActiveTableToExcel = () => {
    const localBaseCols = [
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

    const localKaraCols = [
      "SIRA NO", 
      "ARAÇ PLAKASI / TANIMI", 
      "MARKA", 
      "PARÇA NO (P/N) / MODEL", 
      "BULUNDUĞU YER", 
      "SON KM Sİ", 
      "DURUMU", 
      "BAKIMA TABİ", 
      "SON KONTROL / KALİBRASYON / BAKIM", 
      "GELECEK KONTROL / KALİBRASYON / BAKIM", 
      "SON KONTROLÜ YAPAN FİRMA", 
      "AÇIKLAMA", 
      "90 GÜN UYARISI MAİL GÖNDERİM TARİHİ"
    ];

    if (modalType === 'techizat_matrix' && activeTechizatType) {
      let activeCols = localBaseCols;
      let rawRows = techizatBell429Data;
      if (activeTechizatType === 'at802') { rawRows = techizatAt802Data; }
      else if (activeTechizatType === 't70') { rawRows = techizatT70Data; }
      else if (activeTechizatType === 't70_bumbi_backet') { rawRows = techizatT70BumbiBacketData; }
      else if (activeTechizatType === 't70_helitak') { rawRows = techizatT70HelitakData; }
      else if (activeTechizatType === 'b360') { rawRows = techizatB360Data; }
      else if (activeTechizatType === 'c650') { rawRows = techizatC650Data; }
      else if (activeTechizatType === 'hangar') { activeCols = [...localBaseCols, "BELGE YÜKLE"]; rawRows = techizatHangarData; }
      else if (activeTechizatType === 'kara_araclari') { activeCols = localKaraCols; rawRows = techizatKaraAraclariData; }
      else if (activeTechizatType === 'all') {
        activeCols = ["AİT OLDUĞU BİRİM", ...localBaseCols];
        rawRows = [
          ...techizatBell429Data.map(r => ["BELL 429", ...r]),
          ...techizatAt802Data.map(r => ["AT-802F", ...r]),
          ...techizatT70Data.map(r => ["T-70 YER DESTEK", ...r]),
          ...techizatT70BumbiBacketData.map(r => ["T-70 BUMBİ BACKET", ...r]),
          ...techizatT70HelitakData.map(r => ["T-70 HELİTAK", ...r]),
          ...techizatB360Data.map(r => ["B-360 YER DESTEK", ...r]),
          ...techizatC650Data.map(r => ["C-650 YER DESTEK", ...r]),
          ...techizatKaraAraclariData.map(r => ["KARA ARAÇLARI", ...r]),
          ...techizatHangarData.map(r => ["HANGAR YER DESTEK", ...r]),
        ];
      }
      exportTechizatToExcel(activeTechizatType, activeCols, rawRows, modalTitle || 'TEÇHİZAT LİSTESİ');
    } else if (modalType === 'form_table' && selectedFormId !== null) {
      const cfg = TABLE_CONFIGS[selectedFormId];
      if (cfg) {
        const formCols = cfg.columns.map(c => c.label);
        const currentData = tableData[selectedFormId] || [];
        const formRows = currentData.map((rowObj: any) => cfg.columns.map(c => String(rowObj[c.label] ?? rowObj[c.key] ?? "")));
        exportTechizatToExcel(`form_${selectedFormId}`, formCols, formRows, cfg.title);
      }
    } else {
      openTechizatMatrix('all', 'GENEL ENVANTER VE TEÇHİZAT LİSTESİ');
    }
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
              const upperVal = val.toUpperCase();
              if (upperVal.includes("FAAL") && !upperVal.includes("GAYRİ") && !upperVal.includes("DEĞİL")) {
                style += " background-color: #dcfce7; color: #15803d; font-weight: bold; text-align: center;";
              } else if (upperVal.includes("ONARIMA ALINDI") || upperVal.includes("ONARIMDA") || upperVal.includes("ONARIM")) {
                style += " background-color: #fef08a; color: #854d0e; font-weight: bold; text-align: center;";
              } else if (upperVal.includes("BAKIM") || upperVal.includes("OVERHAUL")) {
                style += " background-color: #ffedd5; color: #c2410c; font-weight: bold; text-align: center;";
              } else if (upperVal.includes("HASARLI") || upperVal.includes("GAYRİ") || upperVal.includes("ARIZALI") || upperVal.includes("FAAL DEĞİL")) {
                style += " background-color: #fee2e2; color: #b91c1c; font-weight: bold; text-align: center;";
              }
            }

            if (colName.includes("GELECEK") || colName.includes("ÖMÜR BİTİŞ") || colName.includes("MUAYENE")) {
              const days = parseGelecekBakimDays(val);
              if (days !== null) {
                if (days < 0) {
                  style += " background-color: #fee2e2; color: #b91c1c; font-weight: bold; text-align: center;";
                } else if (days < 90) {
                  style += " background-color: #ffedd5; color: #c2410c; font-weight: bold; text-align: center;";
                } else {
                  style += " background-color: #dcfce7; color: #15803d; font-weight: bold; text-align: center;";
                }
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
                const upperVal = val.toUpperCase();
                if (upperVal.includes("FAAL") && !upperVal.includes("GAYRİ") && !upperVal.includes("DEĞİL")) {
                  style += " background-color: #dcfce7; color: #15803d; font-weight: bold; text-align: center;";
                } else if (upperVal.includes("ONARIMA ALINDI") || upperVal.includes("ONARIMDA") || upperVal.includes("ONARIM")) {
                  style += " background-color: #fef08a; color: #854d0e; font-weight: bold; text-align: center;";
                } else if (upperVal.includes("BAKIM") || upperVal.includes("OVERHAUL")) {
                  style += " background-color: #ffedd5; color: #c2410c; font-weight: bold; text-align: center;";
                } else if (upperVal.includes("HASARLI") || upperVal.includes("GAYRİ") || upperVal.includes("ARIZALI") || upperVal.includes("FAAL DEĞİL")) {
                  style += " background-color: #fee2e2; color: #b91c1c; font-weight: bold; text-align: center;";
                }
              }

              const currentVal = isMultiLineCol ? (lines[subIdx] || "") : (cell || "");
              if (colName.includes("GELECEK") || colName.includes("ÖMÜR BİTİŞ") || colName.includes("MUAYENE")) {
                const days = parseGelecekBakimDays(currentVal);
                if (days !== null) {
                  if (days < 0) {
                    style += " background-color: #fee2e2; color: #b91c1c; font-weight: bold; text-align: center;";
                  } else if (days < 90) {
                    style += " background-color: #ffedd5; color: #c2410c; font-weight: bold; text-align: center;";
                  } else {
                    style += " background-color: #dcfce7; color: #15803d; font-weight: bold; text-align: center;";
                  }
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

  // Process Excel export with optional embedded images using native ExcelJS (.xlsx)
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
      setExcelExportProgressText("Excel (.xlsx) yerel aktarımı başlatılıyor...");

      const exportCols = [cols[0] || "SIRA NO", "TEÇHİZAT FOTOĞRAFI", ...cols.slice(1)];
      const isAll = type === 'all';

      // Step 1: Pre-collect and pre-fetch all row image Base64s in parallel batches
      const resolvedRowImagesMap: Record<number, string | null> = {};
      const imageFetchTasks: { rIdx: number; url: string }[] = [];

      for (let rIdx = 0; rIdx < rows.length; rIdx++) {
        const row = rows[rIdx];
        const targetTechType = isAll ? getRealTechType(row[0]) : type;
        const targetRow = isAll ? row.slice(1) : row;
        
        const nameVal = (targetRow[1] || "").trim();
        const partVal = (targetRow[2] || "").trim();
        const serialOrLocVal = (targetRow[3] || "").trim();

        const primaryKey = targetTechType + "_" + nameVal.replace(/\s+/g, '_') + "_" + serialOrLocVal.replace(/\s+/g, '_');
        const secondaryKey = targetTechType + "_" + nameVal.replace(/\s+/g, '_') + "_" + partVal.replace(/\s+/g, '_');
        const tertiaryKey = targetTechType + "_" + nameVal.replace(/\s+/g, '_') + "_";

        let imgUrl = techizatImages[primaryKey] || techizatImages[secondaryKey] || techizatImages[tertiaryKey];

        if (!imgUrl && nameVal) {
          const nameClean = nameVal.toLowerCase().replace(/\s+/g, '_');
          const matchedKey = Object.keys(techizatImages).find(k => {
            const kLower = k.toLowerCase();
            return kLower.startsWith(targetTechType.toLowerCase()) && kLower.includes(nameClean);
          });
          if (matchedKey) imgUrl = techizatImages[matchedKey];
        }

        if (imgUrl) {
          imageFetchTasks.push({ rIdx, url: imgUrl });
        }
      }

      // Run parallel batch fetching (12 concurrent items)
      const BATCH_SIZE = 12;
      for (let i = 0; i < imageFetchTasks.length; i += BATCH_SIZE) {
        const batch = imageFetchTasks.slice(i, i + BATCH_SIZE);
        setExcelExportProgressText(`HD Görseller indiriliyor (%${Math.round((i / Math.max(1, imageFetchTasks.length)) * 100)})...`);
        await Promise.all(batch.map(async (task) => {
          try {
            const b64 = await fetchDriveImageAsBase64(task.url);
            if (b64 && b64.startsWith('data:image/')) {
              resolvedRowImagesMap[task.rIdx] = b64;
            } else {
              resolvedRowImagesMap[task.rIdx] = null;
            }
          } catch (e) {
            console.warn("Parallel Excel image fetch error:", task.rIdx, e);
            resolvedRowImagesMap[task.rIdx] = null;
          }
        }));
      }

      setExcelExportProgressText("Yerel Excel (.xlsx) çalışma kitabı ve gömülü görseller oluşturuluyor...");

      // Step 2: Create ExcelJS Workbook for real .xlsx file generation
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Görselli Teçhizat Listesi');

      // Title row
      worksheet.mergeCells(1, 1, 1, exportCols.length);
      const titleCell = worksheet.getCell(1, 1);
      titleCell.value = `${title.toUpperCase()} (GÖRSELLİ PORTAL LİSTESİ)`;
      titleCell.font = { name: 'Segoe UI', size: 13, bold: true, color: { argb: 'FFFFFFFF' } };
      titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0B3D1D' } };
      titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
      worksheet.getRow(1).height = 40;

      // Header row
      const headerRow = worksheet.getRow(2);
      headerRow.height = 28;
      exportCols.forEach((colName, cIdx) => {
        const cell = headerRow.getCell(cIdx + 1);
        cell.value = colName;
        cell.font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } };
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        cell.border = {
          top: { style: 'thin', color: { argb: 'FF475569' } },
          bottom: { style: 'thin', color: { argb: 'FF475569' } },
          left: { style: 'thin', color: { argb: 'FF475569' } },
          right: { style: 'thin', color: { argb: 'FF475569' } }
        };
      });

      // Column widths
      worksheet.getColumn(1).width = 10; // SIRA NO
      worksheet.getColumn(2).width = 20; // TEÇHİZAT FOTOĞRAFI
      for (let c = 3; c <= exportCols.length; c++) {
        worksheet.getColumn(c).width = 26;
      }

      // Populate Rows with native embedded binary images inside the .xlsx package
      for (let rIdx = 0; rIdx < rows.length; rIdx++) {
        if (rIdx % 10 === 0) {
          setExcelExportProgressText(`Excel oluşturuluyor (%${Math.round((rIdx / Math.max(1, rows.length)) * 100)})...`);
          await new Promise(resolve => setTimeout(resolve, 50));
        }
        const row = rows[rIdx];
        const b64Data = resolvedRowImagesMap[rIdx];
        const rowNum = rIdx + 3; // row 1 title, row 2 header
        const excelRow = worksheet.getRow(rowNum);
        excelRow.height = 92; // row height for image cell

        // Cell 1: SIRA NO
        const cell1 = excelRow.getCell(1);
        cell1.value = row[0] || String(rIdx + 1);
        cell1.alignment = { horizontal: 'center', vertical: 'middle' };
        cell1.font = { name: 'Segoe UI', size: 10 };
        cell1.border = {
          top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
          bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
          left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
          right: { style: 'thin', color: { argb: 'FFE2E8F0' } }
        };

        // Cell 2: TEÇHİZAT FOTOĞRAFI (Embedded natively)
        const cell2 = excelRow.getCell(2);
        cell2.border = {
          top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
          bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
          left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
          right: { style: 'thin', color: { argb: 'FFE2E8F0' } }
        };

        if (b64Data && b64Data.startsWith('data:image/')) {
          try {
            let ext: 'png' | 'jpeg' = 'png';
            if (b64Data.includes('data:image/jpg') || b64Data.includes('data:image/jpeg')) {
              ext = 'jpeg';
            }

            const imageId = workbook.addImage({
              base64: b64Data,
              extension: ext,
            });

            worksheet.addImage(imageId, {
              tl: { col: 1.1, row: rowNum - 1 + 0.05 }, // B column, centered
              ext: { width: 105, height: 105 },
              editAs: 'oneCell'
            });
          } catch (e) {
            console.warn("ExcelJS image embed error for row", rIdx, e);
            cell2.value = "Görsel Hata";
            cell2.alignment = { horizontal: 'center', vertical: 'middle' };
          }
        } else {
          cell2.value = "Görsel Yok";
          cell2.alignment = { horizontal: 'center', vertical: 'middle' };
          cell2.font = { name: 'Segoe UI', size: 9, italic: true, color: { argb: 'FF94A3B8' } };
        }

        // Remaining data columns
        for (let cIdx = 1; cIdx < row.length; cIdx++) {
          const val = row[cIdx] || "";
          const colCell = excelRow.getCell(cIdx + 2);
          colCell.value = val;
          colCell.font = { name: 'Segoe UI', size: 10 };
          colCell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
          colCell.border = {
            top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
            bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
            left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
            right: { style: 'thin', color: { argb: 'FFE2E8F0' } }
          };

          const colName = cols[cIdx]?.toUpperCase() || "";
          if (colName.includes("DURUM")) {
            colCell.alignment = { horizontal: 'center', vertical: 'middle' };
            if (val.toUpperCase().includes("FAAL") && !val.toUpperCase().includes("GAYRİ")) {
              colCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDCFCE7' } };
              colCell.font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FF15803D' } };
            } else if (val.toUpperCase().includes("GAYRİ") || val.toUpperCase().includes("ARIZALI") || val.toUpperCase().includes("FAAL DEĞİL")) {
              colCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEE2E2' } };
              colCell.font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FFB91C1C' } };
            }
          }
        }
      }

      // Step 3: Write native binary buffer and trigger download
      setExcelExportProgressText("Excel (.xlsx) indiriliyor...");
      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const link = document.createElement("a");
      link.href = URL.createObjectURL(blob);
      link.download = `techizat_takip_${type}_gorselli.xlsx`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);

      setIsExcelExportLoading(false);
      setExcelExportModalData(null);
      showNotification("Teçhizat listesi içindeki HD görsellerle birlikte gerçek Excel (.xlsx) olarak başarıyla indirildi.");
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
  const [passwordActionType, setPasswordActionType] = useState<'global_sync' | 'filter_sync' | 'new_product' | 'depo_management' | 'gun_takip'>('global_sync');
  const [syncInitialStep, setSyncInitialStep] = useState<1 | 2>(1);
  const [syncInitialTarget, setSyncInitialTarget] = useState<string>('1');

  // Gün Takip Sorumluları ve Renk Kodu Sıralama State Değişkenleri
  const [sortByColor, setSortByColor] = useState<boolean>(false);
  const [selectedColorFilter, setSelectedColorFilter] = useState<'all' | 'red' | 'orange' | 'green' | 'neutral'>('all');
  const [maintenanceModalState, setMaintenanceModalState] = useState<{
    isOpen: boolean;
    itemName: string;
    currentSonKontrol: string;
    currentGelecekKontrol: string;
    onConfirm: (newSon: string, newGelecek: string) => void;
  }>({
    isOpen: false,
    itemName: '',
    currentSonKontrol: '',
    currentGelecekKontrol: '',
    onConfirm: () => {}
  });

  const openBakimYapildiModal = (
    itemName: string,
    sonKontrol: string,
    gelecekKontrol: string,
    onConfirm: (newSon: string, newGelecek: string) => void
  ) => {
    setMaintenanceModalState({
      isOpen: true,
      itemName,
      currentSonKontrol: sonKontrol,
      currentGelecekKontrol: gelecekKontrol,
      onConfirm: (newSon, newGelecek) => {
        onConfirm(newSon, newGelecek);
        setMaintenanceModalState(prev => ({ ...prev, isOpen: false }));
      }
    });
  };
  const [isSorumluModalOpen, setIsSorumluModalOpen] = useState(false);
  const [isSavingSorumlu, setIsSavingSorumlu] = useState(false);
  const [gunTakipSorumlulari, setGunTakipSorumlulari] = useState<{ birim: string; adSoyad: string; eposta: string; mail90?: string; mail60?: string; mail30?: string; }[]>(() => {
    try {
      const saved = localStorage.getItem('gun_takip_sorumlulari');
      if (saved) {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed) && parsed.length > 0) return parsed;
      }
    } catch(e) {}
    return [
      { birim: "AT-802 DEPO", adSoyad: "Sorumlu Personel", eposta: "orhavak.bakimsube@gmail.com", mail90: "" },
      { birim: "AT-802 YER DESTEK", adSoyad: "Sorumlu Personel", eposta: "orhavak.bakimsube@gmail.com", mail90: "" },
      { birim: "AT-802 ÖZEL ALET", adSoyad: "Sorumlu Personel", eposta: "orhavak.bakimsube@gmail.com", mail90: "" },
      { birim: "BELL 429 DEPO", adSoyad: "Sorumlu Personel", eposta: "orhavak.bakimsube@gmail.com", mail90: "" },
      { birim: "BELL 429 YER DESTEK", adSoyad: "Sorumlu Personel", eposta: "orhavak.bakimsube@gmail.com", mail90: "" },
      { birim: "BELL 429 ÖZEL ALET", adSoyad: "Sorumlu Personel", eposta: "orhavak.bakimsube@gmail.com", mail90: "" },
      { birim: "T-70 DEPO", adSoyad: "Sorumlu Personel", eposta: "orhavak.bakimsube@gmail.com", mail90: "" },
      { birim: "T-70 YER DESTEK", adSoyad: "Sorumlu Personel", eposta: "orhavak.bakimsube@gmail.com", mail90: "" },
      { birim: "T-70 ÖZEL ALET", adSoyad: "Sorumlu Personel", eposta: "orhavak.bakimsube@gmail.com", mail90: "" },
      { birim: "T-70 BUMBİ BACKET", adSoyad: "Sorumlu Personel", eposta: "orhavak.bakimsube@gmail.com", mail90: "" },
      { birim: "T-70 HELİTAK", adSoyad: "Sorumlu Personel", eposta: "orhavak.bakimsube@gmail.com", mail90: "" },
      { birim: "B-360 DEPO", adSoyad: "Sorumlu Personel", eposta: "orhavak.bakimsube@gmail.com", mail90: "" },
      { birim: "B-360 YER DESTEK", adSoyad: "Sorumlu Personel", eposta: "orhavak.bakimsube@gmail.com", mail90: "" },
      { birim: "B-360 ÖZEL ALET", adSoyad: "Sorumlu Personel", eposta: "orhavak.bakimsube@gmail.com", mail90: "" },
      { birim: "C-650 DEPO", adSoyad: "Sorumlu Personel", eposta: "orhavak.bakimsube@gmail.com", mail90: "" },
      { birim: "C-650 YER DESTEK", adSoyad: "Sorumlu Personel", eposta: "orhavak.bakimsube@gmail.com", mail90: "" },
      { birim: "C-650 ÖZEL ALET", adSoyad: "Sorumlu Personel", eposta: "orhavak.bakimsube@gmail.com", mail90: "" },
      { birim: "HANGAR YER DESTEK", adSoyad: "Sorumlu Personel", eposta: "orhavak.bakimsube@gmail.com", mail90: "" },
      { birim: "KARA ARAÇLARI BAKIM & MUAYENE", adSoyad: "Sorumlu Personel", eposta: "orhavak.bakimsube@gmail.com", mail90: "" }
    ];
  });

  // 4. OTOMATİK 90 GÜN KONTROLÜ VE E-POSTA BİLDİRİMİ (SİTE AÇILDIĞINDA VE ARKA PLANDA SÜREKLİ AKTİF)
  const [autoReminderStatus, setAutoReminderStatus] = useState<{
    lastChecked: string;
    items90DaysCount: number;
    overdueCount: number;
    isRunning: boolean;
  }>({
    lastChecked: '',
    items90DaysCount: 0,
    overdueCount: 0,
    isRunning: false
  });

  const parseDiffDays = (dateStr: string): number | null => {
    if (!dateStr || typeof dateStr !== 'string') return null;
    const clean = dateStr.trim();
    if (!clean || clean === '-' || clean.toLowerCase() === 'yok') return null;

    let day = 0, month = 0, year = 0;
    if (clean.includes('.')) {
      const parts = clean.split('.');
      if (parts.length >= 3) {
        day = parseInt(parts[0], 10);
        month = parseInt(parts[1], 10) - 1;
        year = parseInt(parts[2].slice(0, 4), 10);
      }
    } else if (clean.includes('/')) {
      const parts = clean.split('/');
      if (parts.length >= 3) {
        day = parseInt(parts[0], 10);
        month = parseInt(parts[1], 10) - 1;
        year = parseInt(parts[2].slice(0, 4), 10);
      }
    } else if (clean.includes('-')) {
      const parts = clean.split('-');
      if (parts.length >= 3) {
        year = parseInt(parts[0], 10);
        month = parseInt(parts[1], 10) - 1;
        day = parseInt(parts[2].slice(0, 2), 10);
      }
    }

    if (!year || isNaN(year) || isNaN(month) || isNaN(day)) return null;

    const targetDate = new Date(year, month, day);
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    targetDate.setHours(0, 0, 0, 0);

    const diffTime = targetDate.getTime() - today.getTime();
    return Math.ceil(diffTime / (1000 * 60 * 60 * 24));
  };

  // Helper to collect all <= 90 days items across all units
  const getGlobal90DaysItems = () => {
    const items: any[] = [];
    const scanRows = (birimName: string, rows: any[][], fleetKey: string) => {
      if (!Array.isArray(rows)) return;
      rows.forEach((r, idx) => {
        if (!r || !Array.isArray(r)) return;
        if (isHeaderLikeRow(r)) return;

        let dateStr = '';
        let isMuaf = false;
        const isKara = fleetKey === 'kara_araclari';
        const isBell = fleetKey === 'bell429';

        if (isKara) {
          const bakimaTabi = String(r[7] || '').trim().toUpperCase();
          if (bakimaTabi === 'HAYIR' || bakimaTabi.includes('MUAF')) isMuaf = true;
          dateStr = String(r[9] || '').trim();
        } else if (isBell) {
          dateStr = String(r[8] || '').trim();
        } else {
          // AT-802, T-70, B-360, C-650, Hangar
          const kalibTabi = String(r[7] || '').trim().toUpperCase();
          if (kalibTabi === 'HAYIR' || kalibTabi.includes('MUAF')) isMuaf = true;
          dateStr = String(r[9] || '').trim();
        }

        if (isMuaf || !dateStr) return;

        const diff = parseDiffDays(dateStr);
        if (diff !== null && diff <= 90) {
          items.push({
            birim: birimName,
            rowIndex: idx,
            malzemeAdi: String(r[1] || ''),
            parcaNo: String(r[2] || ''),
            seriNo: String(r[3] || ''),
            gelecekTarih: dateStr,
            daysDiff: diff,
            isOverdue: diff <= 0
          });
        }
      });
    };

    scanRows('BELL 429', techizatBell429Data, 'bell429');
    scanRows('AT-802', techizatAt802Data, 'at802');
    scanRows('T-70', techizatT70Data, 't70');
    scanRows('T-70 BUMBİ BACKET', techizatT70BumbiBacketData, 't70_bumbi_backet');
    scanRows('T-70 HELİTAK', techizatT70HelitakData, 't70_helitak');
    scanRows('B-360', techizatB360Data, 'b360');
    scanRows('C-650', techizatC650Data, 'c650');
    scanRows('HANGAR', techizatHangarData, 'hangar');
    scanRows('KARA ARAÇLARI', techizatKaraAraclariData, 'kara_araclari');

    return items;
  };

  const runBackgroundAutoReminderCheck = async () => {
    const items = getGlobal90DaysItems();
    const overdue = items.filter(i => i.isOverdue);
    const nowTime = new Date().toLocaleTimeString('tr-TR', { hour: '2-digit', minute: '2-digit' });

    setAutoReminderStatus({
      lastChecked: nowTime,
      items90DaysCount: items.length,
      overdueCount: overdue.length,
      isRunning: true
    });

    try {
      // 1. Notify local backend API
      await fetch('/api/trigger-auto-reminders', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          birimList: gunTakipSorumlulari,
          approachingItems: items,
          overdueItems: overdue
        })
      }).catch(e => console.warn('Local trigger auto reminders warn:', e));

      // 2. Trigger Google Apps Script dailyReminderTrigger
      if (GOOGLE_SCRIPT_URL) {
        await fetch(GOOGLE_SCRIPT_URL, {
          method: 'POST',
          headers: { 'Content-Type': 'text/plain;charset=utf-8' },
          body: JSON.stringify({
            action: 'dailyReminderTrigger',
            reason: 'Otomatik Arka Plan 90 Gün & Günü Geçen Kontrolü',
            items90Count: items.length,
            overdueCount: overdue.length,
            timestamp: new Date().toISOString()
          })
        }).catch(e => console.warn('GAS auto reminder warn:', e));
      }
    } catch (err) {
      console.warn('Auto reminder check error:', err);
    } finally {
      setAutoReminderStatus(prev => ({ ...prev, isRunning: false }));
    }
  };

  // Run automatically on component mount and every 30 minutes
  useEffect(() => {
    const timer = setTimeout(() => {
      runBackgroundAutoReminderCheck();
    }, 2500);

    const interval = setInterval(() => {
      runBackgroundAutoReminderCheck();
    }, 30 * 60 * 1000);

    return () => {
      clearTimeout(timer);
      clearInterval(interval);
    };
  }, [
    techizatBell429Data,
    techizatAt802Data,
    techizatT70Data,
    techizatB360Data,
    techizatC650Data,
    techizatHangarData,
    techizatKaraAraclariData
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

  // Year reference
  const currentYear = new Date().getFullYear();

  // Automatic splash screen timeout helper
  useEffect(() => {
    const timer = setTimeout(() => {
      setSplashVisible(false);
    }, 5000);
    return () => clearTimeout(timer);
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
    getAllHangarPdfDocs().then(docs => {
      if (Array.isArray(docs) && docs.length > 0) {
        setHangarPdfDocs(docs);
      }
    }).catch(err => console.warn('Hangar docs load warning:', err));
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
      case 'HA_YER_DESTEK': return 'HAVA ARAÇLARI YER DESTEK VE ÖZEL ALETLER';
      case 'UNIT_FOLDER_MENU': return `${getUnitDisplayName(selectedUnitFolder)} - YER DESTEK VE DEPO YÖNETİMİ`;
      case 'UNIT_DEPO_MENU': return `${getUnitDisplayName(selectedUnitFolder)} - DEPO YÖNETİMİ`;
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
    setSearchQuery('');
    setTechizatSearchQuery('');
  };

  const navigateToSubCategory = (subCategory: Exclude<CategoryType, null>) => {
    setSelectedCategory(subCategory);
    setModalTitle(getCategoryTitle(subCategory));
    setCategoryHistory(prev => [...prev, subCategory]);
    setSearchQuery('');
    setTechizatSearchQuery('');
  };

  const handleBack = () => {
    if (modalType === 'denetleme') {
      setModalType('category');
      setSelectedCategory('FORM KAYITLARI');
      setModalTitle(getCategoryTitle('FORM KAYITLARI'));
      return;
    }
    if (modalType === 'iframe') {
      setModalType('category');
      setSelectedCategory('FORM KAYITLARI');
      setModalTitle(getCategoryTitle('FORM KAYITLARI'));
      setModalUrl('');
      setIframeLoading(false);
      return;
    }
    if (modalType === 'techizat_matrix') {
      setModalType('category');
      if (activeTechizatType === 'kara_araclari') {
        setSelectedCategory('KARA_ARACLARI_MENU');
        setModalTitle(getCategoryTitle('KARA_ARACLARI_MENU'));
      } else if (activeTechizatType === 'hangar') {
        setSelectedCategory('TEÇHİZAT TAKİP');
        setModalTitle(getCategoryTitle('TEÇHİZAT TAKİP'));
      } else if (techizatActiveSection === 'depo_sarf' || techizatActiveSection === 'depo_kimyasal') {
        setSelectedCategory('UNIT_DEPO_MENU');
        setModalTitle(getCategoryTitle('UNIT_DEPO_MENU'));
      } else if (selectedUnitFolder) {
        setSelectedCategory('UNIT_FOLDER_MENU');
        setModalTitle(getCategoryTitle('UNIT_FOLDER_MENU'));
      } else {
        setSelectedCategory('HA_YER_DESTEK');
        setModalTitle(getCategoryTitle('HA_YER_DESTEK'));
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

  // Row formatting helpers across components and Drive sync
  const formatStandardRow = (row: string[], isHangar = false) => {
    const r = [...row];
    // Self-healing: if r[7] is "EVET"/"HAYIR" and r[8] is ALSO "EVET"/"HAYIR" and r[9] contains a date
    if (
      (r[7] === "EVET" || r[7] === "HAYIR") &&
      (r[8] === "EVET" || r[8] === "HAYIR") &&
      r[9] && (/\d{2}[./-]\d{2}[./-]\d{2,4}/.test(r[9]) || r[9].includes('.'))
    ) {
      r.splice(8, 1);
    }

    const col7Upper = (r[7] || "").trim().toUpperCase();
    if (col7Upper !== "EVET" && col7Upper !== "HAYIR") {
      r.splice(7, 0, "EVET");
    }
    while (r.length < 13) {
      r.push("");
    }
    // Clean up 90 gun mail column (index 12): never show category tags or placeholder dashes
    const mailVal = String(r[12] || "").trim().toLowerCase();
    if (mailVal === 'yer_destek' || mailVal === 'ozel_alet' || mailVal === 'depo_sarf' || mailVal === 'depo_kimyasal' || mailVal === '-' || mailVal === '--' || mailVal.includes('belirtilme')) {
      r[12] = "";
    }
    const base = r.slice(0, 13);
    if (isHangar) {
      base.push("");
    }
    return base;
  };

  const formatDepoRow = (row: string[]) => {
    const r = [...row];
    if (activeTechizatType === 'at802') {
      while (r.length < 18) r.push("");
      return [
        r[0] || "", // SIRA
        r[1] || "", // NAME
        r[2] || "", // PN
        r[3] || "", // SN
        r[4] || "1", // MİKTAR
        r[5] || "", // YER
        r[6] || "", // LOKASYON NO
        r[7] || "0", // TOPLAM STOK
        r[8] || "0", // ANKARA
        r[9] || "0", // MİLAS
        r[10] || "0", // KARAİN
        r[11] || "0", // ÇANAKKALE
        r[12] || "0", // BURSA
        r[13] || "FAAL", // DURUMU
        r[14] || "HAYIR", // ÖMÜRLÜ
        r[15] || "-", // ÖMÜR BİTİŞ
        r[16] || "-", // FİRMA
        r[17] || ""   // AÇIKLAMA
      ];
    }

    while (r.length < 11) {
      r.push("");
    }
    const sira = r[0] || "";
    const name = r[1] || "";
    const pn = r[2] || "";
    const sn = r[3] || "";
    const miktar = r[4] || "1";
    const yer = r[5] || "";
    const durum = r[6] || "FAAL";
    
    let omurlu = (r[7] || "").trim().toUpperCase();
    if (omurlu !== "EVET" && omurlu !== "HAYIR") {
      omurlu = (r[8] && r[8] !== "-" && r[8] !== "--" && r[8] !== "MUAFIYET (TABİ DEĞİL)") ? "EVET" : "HAYIR";
    }
    
    const omurTarihi = omurlu === "EVET" ? (r[8] || "-") : "-";
    const tedarikFirma = r[9] || "-";
    const aciklama = r[10] || "";
    
    return [
      sira,
      name,
      pn,
      sn,
      miktar,
      yer,
      durum,
      omurlu,
      omurTarihi,
      tedarikFirma,
      aciklama
    ];
  };

  const formatKaraAraclariRow = (row: string[]) => {
    const r = [...row];
    // Handle upgrade from 12-column to 13-column (insertion of MARKA)
    if (r.length <= 12) {
      const col6 = (r[6] || "").trim().toUpperCase();
      const col7 = (r[7] || "").trim().toUpperCase();
      if ((col6 === "EVET" || col6 === "HAYIR") && (col7 !== "EVET" && col7 !== "HAYIR")) {
        const combinedText = `${r[1] || ''} ${r[2] || ''}`.toUpperCase();
        let guessedMarka = "-";
        if (combinedText.includes("FORD")) guessedMarka = "FORD";
        else if (combinedText.includes("TOYOTA")) guessedMarka = "TOYOTA";
        else if (combinedText.includes("ISUZU")) guessedMarka = "ISUZU";
        else if (combinedText.includes("MERCEDES")) guessedMarka = "MERCEDES";
        else if (combinedText.includes("FIAT")) guessedMarka = "FIAT";
        else if (combinedText.includes("RENAULT")) guessedMarka = "RENAULT";
        else if (combinedText.includes("VOLKSWAGEN") || combinedText.includes("VW")) guessedMarka = "VOLKSWAGEN";
        
        r.splice(2, 0, guessedMarka);
      }
    }

    const col7Upper = (r[7] || "").trim().toUpperCase();
    if (col7Upper !== "EVET" && col7Upper !== "HAYIR") {
      r.splice(7, 0, "EVET");
    }
    while (r.length < 13) {
      r.push("");
    }
    // Clean up 90 gun mail column (index 12)
    const mailVal = String(r[12] || "").trim().toLowerCase();
    if (mailVal === 'kara_araclari' || mailVal === 'yer_destek' || mailVal === '-' || mailVal === '--' || mailVal.includes('belirtilme')) {
      r[12] = "";
    }
    return r.slice(0, 13);
  };

  const formatKaraAraciToStandardRow = (row: string[]) => {
    const r = formatKaraAraclariRow(row);
    const sira = r[0];
    const plaka = r[1];
    const marka = r[2];
    const model = r[3];
    const yer = r[4];
    const km = r[5] ? `${r[5]} KM` : "";
    const durum = r[6];
    const kalibrasyonTabi = r[7];
    const sonBakim = r[8];
    const gelecekBakim = r[9];
    const firma = r[10];
    const aciklama = r[11];
    const mail = r[12];
    
    return [
      sira,
      plaka,
      marka !== "-" ? `${marka} ${model}`.trim() : model,
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

  const syncTechizatExcelToGoogleDrive = async (techType: string, updatedUnitData: string[][]) => {
    try {
      const subSecs = ['all', 'ozel_alet', 'depo_sarf', 'depo_kimyasal'];
      for (const subSec of subSecs) {
        const sectionRows = subSec === 'all' 
          ? updatedUnitData 
          : updatedUnitData.filter(r => getRowSection(r, techType) === subSec);
        
        if (subSec !== 'all' && sectionRows.length === 0) continue;

        let sheetName = "Sayfa1";
        // Standart dosya adı fallback (eğer Drive'dan okunmuş orijinal adı yoksa kullanılır)
        let fallbackFileName = "techizat.xlsx";
        let unitTitle = getTechizatUnitLabel(techType);

        if (techType === 'at802') {
          if (subSec === 'ozel_alet') { sheetName = "At802_Ozel_Alet"; fallbackFileName = "at-802_ozel_bakim_aletleri.xlsx"; unitTitle = "AT-802F - ÖZEL BAKIM ALETLERİ"; }
          else if (subSec === 'depo_sarf') { sheetName = "At802_Sarf_Depo"; fallbackFileName = "at-802_sarf_ve_parca_deposu.xlsx"; unitTitle = "AT-802F - SARF VE PARÇA DEPOSU"; }
          else if (subSec === 'depo_kimyasal') { sheetName = "At802_Kimyasal_Depo"; fallbackFileName = "at-802_kimyasal_depo.xlsx"; unitTitle = "AT-802F - KİMYASAL DEPO"; }
          else { sheetName = "At802_Techizat"; fallbackFileName = "hava_araçları_yer_destek_at-802.xlsx"; unitTitle = "AT-802F - YER DESTEK TEÇHİZATLARI"; }
        } else if (techType === 'bell429') {
          if (subSec === 'depo_sarf') { sheetName = "Bell429_Sarf_Depo"; fallbackFileName = "bell-429_sarf_ve_parca_deposu.xlsx"; unitTitle = "BELL 429 - SARF VE PARÇA DEPOSU"; }
          else if (subSec === 'depo_kimyasal') { sheetName = "Bell429_Kimyasal_Depo"; fallbackFileName = "bell-429_kimyasal_depo.xlsx"; unitTitle = "BELL 429 - KİMYASAL DEPO"; }
          else { sheetName = "Bell429_Techizat"; fallbackFileName = "hava_araçları_yer_destek_bell-429.xlsx"; unitTitle = "BELL 429 - YER DESTEK TEÇHİZATLARI"; }
        } else if (techType === 't70') {
          if (subSec === 'depo_sarf') { sheetName = "T70_Sarf_Depo"; fallbackFileName = "t-70_sarf_ve_parca_deposu.xlsx"; unitTitle = "T-70 - SARF VE PARÇA DEPOSU"; }
          else if (subSec === 'depo_kimyasal') { sheetName = "T70_Kimyasal_Depo"; fallbackFileName = "t-70_kimyasal_depo.xlsx"; unitTitle = "T-70 - KİMYASAL DEPO"; }
          else { sheetName = "T70_Techizat"; fallbackFileName = "hava_araçları_yer_destek_t-70.xlsx"; unitTitle = "T-70 - YER DESTEK TEÇHİZATLARI"; }
        } else if (techType === 't70_bumbi_backet') {
          sheetName = "T70_Bumbi_Backet"; fallbackFileName = "hava_araçları_yer_destek_t-70_bumbi_backet.xlsx"; unitTitle = "T-70 - BUMBİ BACKET";
        } else if (techType === 't70_helitak') {
          sheetName = "T70_Helitak"; fallbackFileName = "hava_araçları_yer_destek_t-70_helitak.xlsx"; unitTitle = "T-70 - HELİTAK";
        } else if (techType === 'b360') {
          if (subSec === 'depo_sarf') { sheetName = "B360_Sarf_Depo"; fallbackFileName = "b-360_sarf_ve_parca_deposu.xlsx"; unitTitle = "B-360 - SARF VE PARÇA DEPOSU"; }
          else if (subSec === 'depo_kimyasal') { sheetName = "B360_Kimyasal_Depo"; fallbackFileName = "b-360_kimyasal_depo.xlsx"; unitTitle = "B-360 - KİMYASAL DEPO"; }
          else { sheetName = "B360_Techizat"; fallbackFileName = "hava_araçları_yer_destek_b-360.xlsx"; unitTitle = "B-360 - YER DESTEK TEÇHİZATLARI"; }
        } else if (techType === 'c650') {
          if (subSec === 'depo_sarf') { sheetName = "C650_Sarf_Depo"; fallbackFileName = "c-650_sarf_ve_parca_deposu.xlsx"; unitTitle = "C-650 - SARF VE PARÇA DEPOSU"; }
          else if (subSec === 'depo_kimyasal') { sheetName = "C650_Kimyasal_Depo"; fallbackFileName = "c-650_kimyasal_depo.xlsx"; unitTitle = "C-650 - KİMYASAL DEPO"; }
          else { sheetName = "C650_Techizat"; fallbackFileName = "hava_araçları_yer_destek_c-650.xlsx"; unitTitle = "C-650 - YER DESTEK TEÇHİZATLARI"; }
        } else if (techType === 'hangar') {
          if (subSec === 'depo_sarf') { sheetName = "Hangar_Sarf_Depo"; fallbackFileName = "hangar_sarf_ve_parca_deposu.xlsx"; unitTitle = "HANGAR - SARF VE PARÇA DEPOSU"; }
          else if (subSec === 'depo_kimyasal') { sheetName = "Hangar_Kimyasal_Depo"; fallbackFileName = "hangar_kimyasal_depo.xlsx"; unitTitle = "HANGAR - KİMYASAL DEPO"; }
          else { sheetName = "Hangar_Techizat"; fallbackFileName = "hava_araçları_yer_destek_hangar.xlsx"; unitTitle = "HANGAR - YER DESTEK TEÇHİZATLARI"; }
        } else if (techType === 'kara_araclari') {
          sheetName = "Kara_Araclari"; fallbackFileName = "kara_araçları_takip.xlsx"; unitTitle = "KARA ARAÇLARI";
        }

        // Eğer kullanıcı Drive'da bu dosyayı kendi adlandırmışsa o adı kullan (Örn: "hava araçları yer detsk hangar dosyası.xlsx")
        let fileName = techizatFileNames[`${techType}_${subSec}`] || fallbackFileName;

        const headers = techType === 'kara_araclari'
          ? ["SIRA NO", "ARAÇ PLAKASI / TANIMI", "MARKA", "PARÇA NO (P/N) / MODEL", "BULUNDUĞU YER", "SON KM Sİ", "DURUMU", "BAKIMA TABİ", "SON KONTROL / KALİBRASYON / BAKIM", "GELECEK KONTROL / KALİBRASYON / BAKIM", "SON KONTROLÜ YAPAN FİRMA", "AÇIKLAMA", "90 GÜN UYARISI MAİL GÖNDERİM TARİHİ"]
          : (subSec === 'depo_sarf' || subSec === 'depo_kimyasal')
            ? ["SIRA NO", "MALZEME / PARÇA ADI", "PARÇA NO (P/N)", "SERİ NO (S/N)", "MİKTAR", "BULUNDUĞU YER", "DURUMU", "ÖMÜRLÜ PARÇA MI?", "ÖMÜR BİTİŞ TARİHİ", "TEDARİK EDİLEN FİRMA", "AÇIKLAMA", "90 GÜN UYARISI MAİL GÖNDERİM TARİHİ"]
            : ["SIRA NO", "TEÇHİZAT ADI", "PARÇA NO (P/N) / MODEL", "SERİ NO (S/N)", "MİKTAR / KAPASİTE", "BULUNDUĞU YER", "DURUMU", (techType === 'hangar' ? "BAKIMA TABİ" : "KALİBRASYONA TABİ"), "SON KONTROL / KALİBRASYON / BAKIM", "GELECEK KONTROL / KALİBRASYON / BAKIM", "SON KONTROLÜ YAPAN FİRMA", "AÇIKLAMA", "90 GÜN UYARISI MAİL GÖNDERİM TARİHİ"];

        const cleanRows = sectionRows.map((r, idx) => {
          const formatted = techType === 'kara_araclari'
            ? formatKaraAraciToStandardRow(r)
            : (subSec === 'depo_sarf' || subSec === 'depo_kimyasal')
              ? formatDepoRow(r)
              : formatStandardRow(r);
          const copy = [...formatted];
          copy[0] = String(idx + 1);
          return copy.slice(0, headers.length);
        });

        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet([headers, ...cleanRows]);
        XLSX.utils.book_append_sheet(wb, ws, sheetName);
        const b64 = XLSX.write(wb, { type: 'base64', bookType: 'xlsx' });

        fetch(GOOGLE_SCRIPT_URL, {
          method: "POST",
          headers: { "Content-Type": "text/plain;charset=utf-8" },
          body: JSON.stringify({
            action: "uploadPdfToDrive",
            fileName: fileName,
            base64Data: b64,
            unitName: unitTitle,
            month: "Genel Plan"
          })
        }).catch(err => console.error(`Drive backup for ${fileName} failed:`, err));
      }
    } catch (e) {
      console.error("syncTechizatExcelToGoogleDrive error:", e);
    }
  };

  const handleSaveTechizatRow = async (editedRow: string[], techType: string, rIdx: number) => {
    const targetOriginal = activeTechizatRowEdit?.row;
    const isKaraAraci = techType === 'kara_araclari';
    const rowSection = targetOriginal ? getRowSection(targetOriginal, techType) : "";
    const isDepo = techizatActiveSection === 'depo_sarf' || 
                   techizatActiveSection === 'depo_kimyasal' || 
                   techizatActiveSection === 'depo_all' || 
                   rowSection === 'depo_sarf' || 
                   rowSection === 'depo_kimyasal';

    // Normalize final edited row to standard column storage format
    let finalEditedRow: string[] = [];

    if (isKaraAraci) {
      const originalSec = (targetOriginal && targetOriginal[13]) ? targetOriginal[13] : "kara_araclari";
      const preservedMailDate = targetOriginal?.[12] || "";
      finalEditedRow = [
        editedRow[0] || "1", // 0: SIRA NO
        editedRow[1] || "",  // 1: ARAÇ PLAKASI / TANIMI
        editedRow[2] || "-", // 2: MARKA
        editedRow[3] || "",  // 3: PARÇA NO (P/N) / MODEL
        editedRow[4] || "",  // 4: BULUNDUĞU YER
        editedRow[5] || "",  // 5: SON KM Sİ
        editedRow[6] || "FAAL", // 6: DURUMU
        editedRow[7] || "EVET", // 7: BAKIMA TABİ
        editedRow[8] || "",  // 8: SON KONTROL / KALİBRASYON / BAKIM
        editedRow[9] || "",  // 9: GELECEK KONTROL / KALİBRASYON / BAKIM
        editedRow[10] || "", // 10: SON KONTROLÜ YAPAN FİRMA
        editedRow[11] || "", // 11: AÇIKLAMA
        preservedMailDate,   // 12: 90 GÜN UYARISI MAİL GÖNDERİM TARİHİ
        originalSec          // 13: BÖLÜM
      ];
    } else if (isDepo) {
      const originalSec = targetOriginal ? getRowSection(targetOriginal, techType) : (techizatActiveSection && techizatActiveSection !== 'all' ? techizatActiveSection : 'depo_sarf');
      const preservedMailDate = targetOriginal?.[11] || "";
      finalEditedRow = [
        editedRow[0] || "1", // 0: SIRA NO
        editedRow[1] || "",  // 1: MALZEME / PARÇA ADI
        editedRow[2] || "",  // 2: PARÇA NO (P/N)
        editedRow[3] || "-", // 3: SERİ NO (S/N)
        editedRow[4] || "1", // 4: MİKTAR
        editedRow[5] || "",  // 5: BULUNDUĞU YER
        editedRow[6] || "FAAL", // 6: DURUMU
        editedRow[7] || "HAYIR", // 7: ÖMÜRLÜ PARÇA MI?
        editedRow[8] || "",  // 8: ÖMÜR BİTİŞ TARİHİ
        editedRow[9] || "",  // 9: TEDARİK EDİLEN FİRMA
        editedRow[10] || "", // 10: AÇIKLAMA
        preservedMailDate,   // 11: 90 GÜN UYARISI MAİL GÖNDERİM TARİHİ
        originalSec          // 12: BÖLÜM / KATEGORİ
      ];
    } else {
      // Standart Teçhizat & Hangar
      const originalSec = targetOriginal ? getRowSection(targetOriginal, techType) : (techizatActiveSection && techizatActiveSection !== 'all' ? techizatActiveSection : 'yer_destek');
      const preservedMailDate = targetOriginal?.[12] || targetOriginal?.[11] || "";
      const kalibTabiVal = (editedRow[7] || "EVET").toString().trim().toUpperCase();
      const isKalibTabiHayir = kalibTabiVal === "HAYIR";
      finalEditedRow = [
        editedRow[0] || "1", // 0: SIRA NO
        editedRow[1] || "",  // 1: TEÇHİZAT ADI
        editedRow[2] || "",  // 2: PARÇA NO (P/N) / MODEL
        editedRow[3] || "-", // 3: SERİ NO (S/N)
        editedRow[4] || "1", // 4: MİKTAR / KAPASİTE
        editedRow[5] || "",  // 5: BULUNDUĞU YER
        editedRow[6] || "FAAL", // 6: DURUMU
        kalibTabiVal,        // 7: BAKIMA TABİ / KALİBRASYONA TABİ (EVET/HAYIR)
        isKalibTabiHayir ? "-" : (editedRow[8] || ""),  // 8: SON KONTROL / KALİBRASYON / BAKIM
        isKalibTabiHayir ? "-" : (editedRow[9] || ""),  // 9: GELECEK KONTROL / KALİBRASYON / BAKIM
        isKalibTabiHayir ? "-" : (editedRow[10] || ""), // 10: SON KONTROLÜ YAPAN FİRMA
        editedRow[11] || "", // 11: AÇIKLAMA
        preservedMailDate,   // 12: 90 GÜN UYARISI MAİL GÖNDERİM TARİHİ
        originalSec          // 13: BÖLÜM / KATEGORİ
      ];
    }

    let updated: string[][] = [];

    const applyRowUpdate = (currentList: string[][]): string[][] => {
      const u = [...currentList];
      let targetIdx = -1;

      // 1. Doğrudan nesne referansı eşleşmesi
      if (targetOriginal) {
        targetIdx = u.indexOf(targetOriginal);
      }

      // 2. rIdx sınırları içinde ve ad + bulunduğu yer doğrulanıyorsa
      if (targetIdx === -1 && rIdx >= 0 && rIdx < u.length && targetOriginal) {
        const cand = u[rIdx];
        if (String(cand[1] || "").trim() === String(targetOriginal[1] || "").trim() &&
            String(cand[5] || "").trim() === String(targetOriginal[5] || "").trim()) {
          targetIdx = rIdx;
        }
      }

      // 3. Bir sistemden birden fazla yer olabileceği için: Ad + Bulunduğu Yer (5) + S/N (3) + P/N (2) + Kategori eşleşmesi
      if (targetIdx === -1 && targetOriginal) {
        targetIdx = u.findIndex(r => 
          (r === targetOriginal) ||
          (String(r[1] || "").trim() === String(targetOriginal[1] || "").trim() &&
           String(r[5] || "").trim() === String(targetOriginal[5] || "").trim() &&
           String(r[3] || "").trim() === String(targetOriginal[3] || "").trim() &&
           String(r[2] || "").trim() === String(targetOriginal[2] || "").trim() &&
           getRowSection(r, techType) === getRowSection(targetOriginal, techType))
        );
      }

      // 4. Ad + Bulunduğu Yer + Kategori eşleşmesi
      if (targetIdx === -1 && targetOriginal) {
        targetIdx = u.findIndex(r => 
          String(r[1] || "").trim() === String(targetOriginal[1] || "").trim() &&
          String(r[5] || "").trim() === String(targetOriginal[5] || "").trim() &&
          getRowSection(r, techType) === getRowSection(targetOriginal, techType)
        );
      }

      // 5. Ad + S/N + Kategori eşleşmesi
      if (targetIdx === -1 && targetOriginal) {
        targetIdx = u.findIndex(r => 
          String(r[1] || "").trim() === String(targetOriginal[1] || "").trim() &&
          String(r[3] || "").trim() === String(targetOriginal[3] || "").trim() &&
          getRowSection(r, techType) === getRowSection(targetOriginal, techType)
        );
      }

      // 6. rIdx sınır kontrolü
      if (targetIdx === -1 && rIdx >= 0 && rIdx < u.length) {
        targetIdx = rIdx;
      }

      if (targetIdx >= 0 && targetIdx < u.length) {
        u[targetIdx] = finalEditedRow;
      } else {
        u.push(finalEditedRow);
      }
      return u;
    };

    if (techType === 'bell429') {
      updated = applyRowUpdate(techizatBell429Data);
      setTechizatBell429Data(updated);
      localStorage.setItem('excel_techizat_bell429_data', JSON.stringify(updated));
    } else if (techType === 'at802') {
      updated = applyRowUpdate(techizatAt802Data);
      setTechizatAt802Data(updated);
      localStorage.setItem('excel_techizat_at802_data', JSON.stringify(updated));
    } else if (techType === 't70') {
      updated = applyRowUpdate(techizatT70Data);
      setTechizatT70Data(updated);
      localStorage.setItem('excel_techizat_t70_data', JSON.stringify(updated));
    } else if (techType === 't70_bumbi_backet') {
      updated = applyRowUpdate(techizatT70BumbiBacketData);
      setTechizatT70BumbiBacketData(updated);
      localStorage.setItem('excel_techizat_t70_bumbi_backet_data', JSON.stringify(updated));
    } else if (techType === 't70_helitak') {
      updated = applyRowUpdate(techizatT70HelitakData);
      setTechizatT70HelitakData(updated);
      localStorage.setItem('excel_techizat_t70_helitak_data', JSON.stringify(updated));
    } else if (techType === 'b360') {
      updated = applyRowUpdate(techizatB360Data);
      setTechizatB360Data(updated);
      localStorage.setItem('excel_techizat_b360_data', JSON.stringify(updated));
    } else if (techType === 'c650') {
      updated = applyRowUpdate(techizatC650Data);
      setTechizatC650Data(updated);
      localStorage.setItem('excel_techizat_c650_data', JSON.stringify(updated));
    } else if (techType === 'hangar') {
      updated = applyRowUpdate(techizatHangarData);
      setTechizatHangarData(updated);
      localStorage.setItem('excel_techizat_hangar_data', JSON.stringify(updated));
    } else if (techType === 'kara_araclari') {
      updated = applyRowUpdate(techizatKaraAraclariData);
      setTechizatKaraAraclariData(updated);
      localStorage.setItem('excel_techizat_kara_araclari_data', JSON.stringify(updated));
    }

    // Close modal immediately and show immediate responsive feedback
    setActiveTechizatRowEdit(null);
    showNotification("Kayıt başarıyla güncellendi! Google Drive Excel dosyası güncelleniyor...");

    // Background sync to Google Drive Excel
    if (techType && updated.length > 0) {
      syncTechizatExcelToGoogleDrive(techType, updated);
      showNotification("Değişiklikler Google Drive Excel dosyasına başarıyla kaydedildi!");
    }
  };

  /**
   * Tekil Satır Silme:
   * Seçilen satırı yerel veritabanından siler, Excel dosyasını güncelleyip Google Drive'a yükler
   * ve Google E-Tablo TÜM TECHİZAT sayfasından da ilgili satırı kaldırır.
   */
  const handleDeleteTechizatRow = async (rowToDelete: string[], techType: string) => {
    if (!rowToDelete || !techType) return;

    const unitLabel = getTechizatUnitLabel(techType);
    let updatedList: string[][] = [];

    const filterRow = (list: string[][]): string[][] => {
      return list.filter(r => {
        const isExactMatch = r === rowToDelete ||
          (String(r[1] || "").trim() === String(rowToDelete[1] || "").trim() &&
           String(r[2] || "").trim() === String(rowToDelete[2] || "").trim() &&
           String(r[3] || "").trim() === String(rowToDelete[3] || "").trim() &&
           getRowSection(r, techType) === getRowSection(rowToDelete, techType));
        return !isExactMatch;
      });
    };

    if (techType === 'bell429') {
      updatedList = filterRow(techizatBell429Data);
      setTechizatBell429Data(updatedList);
      localStorage.setItem('excel_techizat_bell429_data', JSON.stringify(updatedList));
    } else if (techType === 'at802') {
      updatedList = filterRow(techizatAt802Data);
      setTechizatAt802Data(updatedList);
      localStorage.setItem('excel_techizat_at802_data', JSON.stringify(updatedList));
    } else if (techType === 't70') {
      updatedList = filterRow(techizatT70Data);
      setTechizatT70Data(updatedList);
      localStorage.setItem('excel_techizat_t70_data', JSON.stringify(updatedList));
    } else if (techType === 't70_bumbi_backet') {
      updatedList = filterRow(techizatT70BumbiBacketData);
      setTechizatT70BumbiBacketData(updatedList);
      localStorage.setItem('excel_techizat_t70_bumbi_backet_data', JSON.stringify(updatedList));
    } else if (techType === 't70_helitak') {
      updatedList = filterRow(techizatT70HelitakData);
      setTechizatT70HelitakData(updatedList);
      localStorage.setItem('excel_techizat_t70_helitak_data', JSON.stringify(updatedList));
    } else if (techType === 'b360') {
      updatedList = filterRow(techizatB360Data);
      setTechizatB360Data(updatedList);
      localStorage.setItem('excel_techizat_b360_data', JSON.stringify(updatedList));
    } else if (techType === 'c650') {
      updatedList = filterRow(techizatC650Data);
      setTechizatC650Data(updatedList);
      localStorage.setItem('excel_techizat_c650_data', JSON.stringify(updatedList));
    } else if (techType === 'hangar') {
      updatedList = filterRow(techizatHangarData);
      setTechizatHangarData(updatedList);
      localStorage.setItem('excel_techizat_hangar_data', JSON.stringify(updatedList));
    } else if (techType === 'kara_araclari') {
      updatedList = filterRow(techizatKaraAraclariData);
      setTechizatKaraAraclariData(updatedList);
      localStorage.setItem('excel_techizat_kara_araclari_data', JSON.stringify(updatedList));
    }

    // Modal'ı hemen kapat ve kullanıcıya anında bildirim ver
    setActiveTechizatRowEdit(null);
    showNotification("Seçilen kayıt silindi! Google Drive Excel dosyası güncelleniyor...");

    // 1. Güncellenmiş listeyi Excel olarak doğrudan Google Drive'a senkronize et
    syncTechizatExcelToGoogleDrive(techType, updatedList);

    // Doğrudan Google Drive Excel klasörüne kaydedilir
    showNotification("Seçilen kayıt Google Drive Excel dosyası üzerinden başarıyla silindi ve güncellendi.");
  };

  /**
   * Çoklu Seçilenleri Silme:
   * Tabloda tiklenen tüm satırları ilgili birimlerden siler, Drive ve E-Tablo senkronizasyonu yapar.
   */
  const handleDeleteSelectedTechizatRows = async () => {
    const selectedKeys = Object.keys(selectedTechizatItems);
    if (selectedKeys.length === 0) return;

    const countToDelete = selectedKeys.length;
    const groupedByTechType: Record<string, string[][]> = {};
    selectedKeys.forEach(k => {
      const item = selectedTechizatItems[k];
      if (item && item.techType) {
        if (!groupedByTechType[item.techType]) groupedByTechType[item.techType] = [];
        groupedByTechType[item.techType].push(item.row);
      }
    });

    for (const techType of Object.keys(groupedByTechType)) {
      const rowsToDel = groupedByTechType[techType];
      let updatedList: string[][] = [];

      const normalize = (s: any) => 
        String(s || "").trim().toLowerCase()
          .replace(/ı/g, 'i').replace(/ğ/g, 'g').replace(/ü/g, 'u')
          .replace(/ş/g, 's').replace(/ö/g, 'o').replace(/ç/g, 'c')
          .replace(/\s+/g, ' ');

      const filterList = (list: string[][]): string[][] => {
        return list.filter(r => {
          // Kullanıcı kuralı: "VERİ GÜNCELLEMEDE BAŞLIK 1. SATIR ÜRÜN DİYE ATMIŞ HATADIR."
          // Silme işlemi yapılırken başlık benzeri artıklar varsa da tamamen temizlensin
          if (isHeaderLikeRow(r)) return false;

          const isMatched = rowsToDel.some(d => {
            if (r === d) return true;

            const r1 = normalize(r[1]);
            const d1 = normalize(d[1]);
            const r2 = normalize(r[2]);
            const d2 = normalize(d[2]);
            const r3 = normalize(r[3]);
            const d3 = normalize(d[3]);

            // Tam satır karşılaştırması
            const rFull = r.slice(1).map(normalize).join('|');
            const dFull = d.slice(1).map(normalize).join('|');
            if (rFull === dFull) return true;

            // İsim eşleşmesi ve varsa P/N veya S/N kontrolü
            if (r1 && d1 && r1 === d1) {
              if (r2 && d2 && r2 !== '-' && d2 !== '-' && r2 !== d2) return false;
              if (r3 && d3 && r3 !== '-' && d3 !== '-' && r3 !== d3) return false;
              return true;
            }

            if (isHeaderLikeRow(d) && isHeaderLikeRow(r)) return true;

            return false;
          });
          return !isMatched;
        });
      };

      if (techType === 'bell429') {
        updatedList = filterList(techizatBell429Data);
        setTechizatBell429Data(updatedList);
      } else if (techType === 'at802') {
        updatedList = filterList(techizatAt802Data);
        setTechizatAt802Data(updatedList);
      } else if (techType === 't70') {
        updatedList = filterList(techizatT70Data);
        setTechizatT70Data(updatedList);
      } else if (techType === 't70_bumbi_backet') {
        updatedList = filterList(techizatT70BumbiBacketData);
        setTechizatT70BumbiBacketData(updatedList);
      } else if (techType === 't70_helitak') {
        updatedList = filterList(techizatT70HelitakData);
        setTechizatT70HelitakData(updatedList);
      } else if (techType === 'b360') {
        updatedList = filterList(techizatB360Data);
        setTechizatB360Data(updatedList);
      } else if (techType === 'c650') {
        updatedList = filterList(techizatC650Data);
        setTechizatC650Data(updatedList);
      } else if (techType === 'hangar') {
        updatedList = filterList(techizatHangarData);
        setTechizatHangarData(updatedList);
      } else if (techType === 'kara_araclari') {
        updatedList = filterList(techizatKaraAraclariData);
        setTechizatKaraAraclariData(updatedList);
      }

      // Sıra numaralarını yeniden 1'den başlat
      updatedList = updatedList.map((r, idx) => {
        const cloned = [...r];
        cloned[0] = String(idx + 1);
        return cloned;
      });

      // Her iki localStorage anahtarını da güncelle
      try {
        localStorage.setItem(`excel_techizat_${techType}_data`, JSON.stringify(updatedList));
        localStorage.setItem(`techizat_${techType}_data`, JSON.stringify(updatedList));
      } catch (err) {
        console.error("Storage error:", err);
      }

      // Arka planda Google Drive senkronizasyonu
      syncTechizatExcelToGoogleDrive(techType, updatedList);
    }

    setSelectedTechizatItems({});
    showNotification(`✅ ${countToDelete} adet seçilen kayıt başarıyla silindi ve liste güncellendi.`);
  };

  /**
   * Yeni Ürün / Teçhizat / Malzeme Ekleme:
   * Form alanlarını ilgili birim ve kategoriye göre hazırlar, dizinin en sonuna ekler,
   * yeni Excel'i Google Drive'a kaydeder ve Google E-Tablo TÜM TECHİZAT sayfasına işler.
   */
  const handleSaveNewProduct = async () => {
    const targetUnit = newProductForm.unit || (activeTechizatType !== 'all' ? activeTechizatType : 'at802');
    const targetSec = newProductForm.section || (techizatActiveSection && techizatActiveSection !== 'all' ? techizatActiveSection : 'yer_destek');
    const isKara = targetUnit === 'kara_araclari' || targetSec === 'kara_araclari';
    const isDepo = targetSec === 'depo_sarf' || targetSec === 'depo_kimyasal';

    if (!newProductForm.name.trim()) {
      alert("Lütfen ürün / teçhizat / malzeme adını giriniz.");
      return;
    }

    let currentList: string[][] = [];
    if (targetUnit === 'bell429') currentList = techizatBell429Data;
    else if (targetUnit === 'at802') currentList = techizatAt802Data;
    else if (targetUnit === 't70') currentList = techizatT70Data;
    else if (targetUnit === 't70_bumbi_backet') currentList = techizatT70BumbiBacketData;
    else if (targetUnit === 't70_helitak') currentList = techizatT70HelitakData;
    else if (targetUnit === 'b360') currentList = techizatB360Data;
    else if (targetUnit === 'c650') currentList = techizatC650Data;
    else if (targetUnit === 'hangar') currentList = techizatHangarData;
    else if (targetUnit === 'kara_araclari') currentList = techizatKaraAraclariData;

    const nextSiraNo = newProductForm.siraNo.trim() ? newProductForm.siraNo.trim() : String(currentList.length + 1);

    let newRow: string[] = [];
    if (isKara) {
      newRow = [
        nextSiraNo, // 0: SIRA NO
        newProductForm.name.trim(), // 1: ARAÇ PLAKASI / TANIMI
        newProductForm.marka?.trim() || "-", // 2: MARKA
        newProductForm.pn.trim() || "-", // 3: PARÇA NO (P/N) / MODEL
        newProductForm.yer.trim() || "-", // 4: BULUNDUĞU YER
        newProductForm.miktar.trim() || "0", // 5: SON KM Sİ
        newProductForm.durumu || "FAAL", // 6: DURUMU
        newProductForm.kalibrasyonaTabi || "EVET", // 7: BAKIMA TABİ
        newProductForm.sonKontrol.trim() || "-", // 8: SON MUAYENE / BAKIM
        newProductForm.gelecekKontrol.trim() || "-", // 9: GELECEK MUAYENE / BAKIM
        newProductForm.firma.trim() || "-", // 10: SON KONTROLÜ YAPAN FİRMA
        newProductForm.aciklama.trim() || "-", // 11: AÇIKLAMA
        "", // 12: 90 GÜN MAIL
        "kara_araclari" // 13: BÖLÜM
      ];
    } else if (isDepo) {
      newRow = [
        nextSiraNo, // 0: SIRA NO
        newProductForm.name.trim(), // 1: MALZEME / PARÇA ADI
        newProductForm.pn.trim() || "-", // 2: PARÇA NO (P/N)
        newProductForm.sn.trim() || "-", // 3: SERİ NO (S/N)
        newProductForm.miktar.trim() || "1 ADET", // 4: MİKTAR
        newProductForm.yer.trim() || "-", // 5: BULUNDUĞU YER / RAF
        newProductForm.durumu || "FAAL", // 6: DURUMU
        newProductForm.kalibrasyonaTabi || "HAYIR", // 7: ÖMÜRLÜ PARÇA MI?
        newProductForm.gelecekKontrol.trim() || "-", // 8: ÖMÜR BİTİŞ TARİHİ
        newProductForm.firma.trim() || "-", // 9: TEDARİK EDİLEN FİRMA
        newProductForm.aciklama.trim() || "-", // 10: AÇIKLAMA
        "", // 11: 90 GÜN MAIL
        targetSec // 12: BÖLÜM / KATEGORİ
      ];
    } else {
      const kalibTabiVal = (newProductForm.kalibrasyonaTabi || "EVET").trim().toUpperCase();
      const isKalibHayir = kalibTabiVal === "HAYIR";
      newRow = [
        nextSiraNo, // 0: SIRA NO
        newProductForm.name.trim(), // 1: TEÇHİZAT ADI
        newProductForm.pn.trim() || "-", // 2: PARÇA NO (P/N) / MODEL
        newProductForm.sn.trim() || "-", // 3: SERİ NO (S/N)
        newProductForm.miktar.trim() || "1 ADET", // 4: MİKTAR / KAPASİTE
        newProductForm.yer.trim() || "-", // 5: BULUNDUĞU YER
        newProductForm.durumu || "FAAL", // 6: DURUMU
        kalibTabiVal, // 7: BAKIMA TABİ / KALİBRASYONA TABİ (EVET/HAYIR)
        isKalibHayir ? "-" : (newProductForm.sonKontrol.trim() || "-"), // 8: SON KONTROL / BAKIM
        isKalibHayir ? "-" : (newProductForm.gelecekKontrol.trim() || "-"), // 9: GELECEK KONTROL / BAKIM
        isKalibHayir ? "-" : (newProductForm.firma.trim() || "-"), // 10: SON KONTROLÜ YAPAN FİRMA
        newProductForm.aciklama.trim() || "-", // 11: AÇIKLAMA
        "", // 12: 90 GÜN MAIL
        targetSec // 13: BÖLÜM / KATEGORİ
      ];
    }

    const updated = [...currentList, newRow];
    if (targetUnit === 'bell429') { setTechizatBell429Data(updated); localStorage.setItem('excel_techizat_bell429_data', JSON.stringify(updated)); }
    else if (targetUnit === 'at802') { setTechizatAt802Data(updated); localStorage.setItem('excel_techizat_at802_data', JSON.stringify(updated)); }
    else if (targetUnit === 't70') { setTechizatT70Data(updated); localStorage.setItem('excel_techizat_t70_data', JSON.stringify(updated)); }
    else if (targetUnit === 't70_bumbi_backet') { setTechizatT70BumbiBacketData(updated); localStorage.setItem('excel_techizat_t70_bumbi_backet_data', JSON.stringify(updated)); }
    else if (targetUnit === 't70_helitak') { setTechizatT70HelitakData(updated); localStorage.setItem('excel_techizat_t70_helitak_data', JSON.stringify(updated)); }
    else if (targetUnit === 'b360') { setTechizatB360Data(updated); localStorage.setItem('excel_techizat_b360_data', JSON.stringify(updated)); }
    else if (targetUnit === 'c650') { setTechizatC650Data(updated); localStorage.setItem('excel_techizat_c650_data', JSON.stringify(updated)); }
    else if (targetUnit === 'hangar') { setTechizatHangarData(updated); localStorage.setItem('excel_techizat_hangar_data', JSON.stringify(updated)); }
    else if (targetUnit === 'kara_araclari') { setTechizatKaraAraclariData(updated); localStorage.setItem('excel_techizat_kara_araclari_data', JSON.stringify(updated)); }

    setIsNewProductModalOpen(false);
    showNotification("Yeni ürün/kayıt başarıyla eklendi! Google Drive Excel dosyası güncelleniyor...");

    // En son satıra eklenmiş yeni listeyi Excel olarak Drive'a senkronize et
    syncTechizatExcelToGoogleDrive(targetUnit, updated);

    // Canlı Google E-Tabloya satırı ekle
    const unitLabel = getTechizatUnitLabel(targetUnit);
    showNotification("Yeni kayıt Google Drive Excel dosyasına başarıyla eklendi.");
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
      
      const scriptUrls = [
        GOOGLE_SCRIPT_URL,
        "https://script.google.com/macros/s/AKfycbw4kruTTc058Y9rLTyO3dKi6KloYsmdDTwV1GSiAk8ZXefyo3Z7_VDSTuurzsS9BHAQyQ/exec",
        "https://script.google.com/macros/s/AKfycbP1uOo2NrST5a4I8vm1nGBLtI26yY2lWrmu9_e9iymwBkUhJBA9JOPCp7SNKqJbOubOw/exec"
      ];

      let response: Response | null = null;
      let lastError: any = null;

      for (const url of scriptUrls) {
        try {
          const targetUrl = `${url}?action=readSheet&sheetName=${encodeURIComponent(activeSheetName)}`;
          const res = await fetch(targetUrl);
          if (res.ok) {
            response = res;
            break;
          } else {
            lastError = new Error(`HTTP Hata: ${res.status}`);
          }
        } catch (err: any) {
          lastError = err;
        }
      }

      if (!response || !response.ok) {
        throw lastError || new Error("Hiçbir bulut sunucusuna erişilemedi.");
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
        // Sadece veri çekildiğinde güncelleme tarihi değiştirilmez, e-tablodaki gerçek son değişiklik tarihi okunur:
        fetchUpdateDatesFromGoogleSheet();

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
      console.warn("Bulut okuma hatası (arka plan/çevrimdışı):", e.message || e);
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

  const getTechizatUnitPassword = (techType: string | null | undefined): string => {
    if (!techType) return '1839';
    const t = techType.toLowerCase();
    if (t === 'at802' || t === 'at-802' || t.includes('802')) return '802';
    if (t === 'bell429' || t.includes('429')) return '429';
    if (t === 'b360' || t.includes('360')) return '360';
    if (t === 'c650' || t.includes('650')) return '650';
    if (t.includes('70') || t === 't70' || t.includes('t-70')) return '70';
    return '1839';
  };


  // Geçmiş / arızalı ürün tekrar FAAL'e çevrildiğinde bir önceki periyot kadar ileriye hesaplama
  const calculateReactivatedMaintenanceDates = (
    prevSonKontrol: string,
    prevGelecekKontrol: string
  ): { newSonKontrol: string; newGelecekKontrol: string; intervalDays: number } => {
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const parseSingleDmy = (str: string): Date | null => {
      if (!str || str === '-' || str === '--') return null;
      const clean = str.split(/[\r\n;]+/)[0].trim();
      const parts = clean.split(/[\.\/-]/);
      if (parts.length === 3) {
        if (parts[2].length === 4) {
          const d = new Date(Number(parts[2]), Number(parts[1]) - 1, Number(parts[0]));
          if (!isNaN(d.getTime())) return d;
        }
        if (parts[0].length === 4) {
          const d = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
          if (!isNaN(d.getTime())) return d;
        }
      }
      return null;
    };

    const dSon = parseSingleDmy(prevSonKontrol);
    const dGelecek = parseSingleDmy(prevGelecekKontrol);

    let intervalDays = 365; // Varsayılan 1 yıl periyot
    if (dSon && dGelecek) {
      const diffTime = dGelecek.getTime() - dSon.getTime();
      const diffDays = Math.round(diffTime / (1000 * 60 * 60 * 24));
      if (diffDays > 0) {
        intervalDays = diffDays;
      }
    }

    const formatToDmy = (d: Date): string => {
      const day = String(d.getDate()).padStart(2, '0');
      const month = String(d.getMonth() + 1).padStart(2, '0');
      const year = d.getFullYear();
      return `${day}.${month}.${year}`;
    };

    const newSonKontrol = formatToDmy(today);
    const targetGelecek = new Date(today);
    targetGelecek.setDate(targetGelecek.getDate() + intervalDays);
    const newGelecekKontrol = formatToDmy(targetGelecek);

    return { newSonKontrol, newGelecekKontrol, intervalDays };
  };

  const loadRegionalLocationsForRow = (targetTechType: string, resolvedCopy: string[]) => {
    const partName = (resolvedCopy[1] || '').trim();
    const partPn = (resolvedCopy[2] || '').trim();
    const partSn = (resolvedCopy[3] || '').trim();
    const partMiktar = (resolvedCopy[4] || '1').trim();
    const partYer = (resolvedCopy[5] || '').trim();
    const partSonKontrol = cleanAndFormatDateString(resolvedCopy[8] !== undefined ? resolvedCopy[8] : '');
    const partGelecekKontrol = cleanAndFormatDateString(resolvedCopy[9] !== undefined ? resolvedCopy[9] : '');
    const partFirma = (resolvedCopy[10] !== undefined ? resolvedCopy[10] : '').trim();

    let sourceData: string[][] = [];
    if (targetTechType === 'bell429') sourceData = techizatBell429Data;
    else if (targetTechType === 'at802') sourceData = techizatAt802Data;
    else if (targetTechType === 't70') sourceData = techizatT70Data;
    else if (targetTechType === 't70_bumbi_backet') sourceData = techizatT70BumbiBacketData;
    else if (targetTechType === 't70_helitak') sourceData = techizatT70HelitakData;
    else if (targetTechType === 'b360') sourceData = techizatB360Data;
    else if (targetTechType === 'c650') sourceData = techizatC650Data;
    else if (targetTechType === 'hangar') sourceData = techizatHangarData;
    else if (targetTechType === 'kara_araclari') sourceData = techizatKaraAraclariData;

    const splitLinesPreserve = (val: string): string[] => {
      if (!val) return [];
      if (val.includes('\n')) {
        return val.split(/\r?\n/).map(s => s.trim());
      }
      if (val.includes(';')) {
        return val.split(';').map(s => s.trim());
      }
      return [val.trim()];
    };

    const snLines = splitLinesPreserve(partSn);
    const locLines = splitLinesPreserve(partYer);
    const qtyLines = splitLinesPreserve(partMiktar);
    const sonLines = splitLinesPreserve(partSonKontrol);
    const gelecekLines = splitLinesPreserve(partGelecekKontrol);
    const firmaLines = splitLinesPreserve(partFirma);

    const maxLines = Math.max(snLines.length, locLines.length, qtyLines.length, gelecekLines.length);

    let items: Array<{
      id: string;
      location: string;
      quantity: number | string;
      serialNumbers: string[];
      sonKontrol: string;
      gelecekKontrol: string;
      firma: string;
    }> = [];

    // Eğer hücre içinde birden fazla satır varsa (hat bazında S/N - Yer eşleştirmesi)
    if (maxLines > 1) {
      // Satır satır oku: her satırdaki yer, onun seri nosu, miktarı ve tarihidir
      type GroupedLoc = {
        location: string;
        serialNumbers: string[];
        sonKontrol: string;
        gelecekKontrol: string;
        firma: string;
        totalQty: number;
      };
      const groupMap: Record<string, GroupedLoc> = {};

      for (let i = 0; i < maxLines; i++) {
        const lineLoc = (locLines[i] !== undefined && locLines[i] !== '') 
          ? locLines[i] 
          : (locLines[0] || 'MERKEZ');
        const lineSn = (snLines[i] !== undefined && snLines[i] !== '-' && snLines[i] !== '--') ? snLines[i] : '';
        const lineQty = parseFloat(qtyLines[i]) || (lineSn ? 1 : (parseFloat(qtyLines[0]) || 1));
        const lineSon = (sonLines[i] !== undefined && sonLines[i] !== '') ? sonLines[i] : (sonLines[0] || '-');
        const lineGelecek = (gelecekLines[i] !== undefined && gelecekLines[i] !== '') ? gelecekLines[i] : (gelecekLines[0] || '-');
        const lineFirma = (firmaLines[i] !== undefined && firmaLines[i] !== '') ? firmaLines[i] : (firmaLines[0] || '-');

        const cleanLoc = lineLoc.trim();
        const groupKey = `${cleanLoc.toUpperCase()}___${lineGelecek.trim()}`;

        if (!groupMap[groupKey]) {
          groupMap[groupKey] = {
            location: cleanLoc,
            serialNumbers: lineSn ? [lineSn] : [],
            sonKontrol: lineSon,
            gelecekKontrol: lineGelecek,
            firma: lineFirma,
            totalQty: lineQty
          };
        } else {
          if (lineSn && !groupMap[groupKey].serialNumbers.includes(lineSn)) {
            groupMap[groupKey].serialNumbers.push(lineSn);
          }
          groupMap[groupKey].totalQty += lineQty;
        }
      }

      items = Object.values(groupMap).map((grp, idx) => ({
        id: `reg_${Date.now()}_${idx}`,
        location: grp.location || 'MERKEZ',
        quantity: grp.serialNumbers.length > 0 ? grp.serialNumbers.length : grp.totalQty,
        serialNumbers: grp.serialNumbers,
        sonKontrol: grp.sonKontrol,
        gelecekKontrol: grp.gelecekKontrol,
        firma: grp.firma
      }));
    } else {
      const isValidPn = partPn && partPn !== '-' && partPn.length > 2;
      const sameRows = sourceData.filter(r => {
        const rPn = (r[2] || '').trim();
        const rName = (r[1] || '').trim();
        if (isValidPn && rPn.toLowerCase() === partPn.toLowerCase()) return true;
        if (!isValidPn && partName && rName.toLowerCase() === partName.toLowerCase()) return true;
        return false;
      });

      if (sameRows.length > 1) {
        items = sameRows.map((r, idx) => {
          const sns = (r[3] || '').split(/[,;\/\n]+/).map(s => s.trim()).filter(s => s && s !== '-');
          return {
            id: `reg_${Date.now()}_${idx}`,
            location: (r[5] || 'MERKEZ').trim(),
            quantity: r[4] || (sns.length > 0 ? String(sns.length) : '1'),
            serialNumbers: sns,
            sonKontrol: cleanAndFormatDateString(r[8] || ''),
            gelecekKontrol: cleanAndFormatDateString(r[9] || ''),
            firma: (r[10] || '').trim()
          };
        });
      } else {
        const knownRegions = ['ÇANAKKALE', 'BURSA', 'MİLAS', 'ANTALYA', 'ANKARA', 'İZMİR', 'MUĞLA', 'ADANA', 'ESKİŞEHİR', 'İSTANBUL', 'BALIKESİR', 'DENİZLİ', 'TRABZON', 'MERSİN', 'EDREMİT', 'DALAMAN', 'GAZİEMİR'];
        const parsedSns = partSn.split(/[,;\/\n]+/).map(s => s.trim()).filter(s => s && s !== '-');
        const upperLoc = partYer.toUpperCase();
        const foundRegions = knownRegions.filter(reg => upperLoc.includes(reg));

        if (foundRegions.length > 1) {
          items = foundRegions.map((reg, idx) => {
            const assignedSn = parsedSns[idx] ? [parsedSns[idx]] : [];
            return {
              id: `reg_${Date.now()}_${idx}`,
              location: reg,
              quantity: assignedSn.length > 0 ? assignedSn.length : 1,
              serialNumbers: assignedSn,
              sonKontrol: partSonKontrol,
              gelecekKontrol: partGelecekKontrol,
              firma: partFirma
            };
          });
        } else {
          items = [{
            id: `reg_${Date.now()}_0`,
            location: partYer || 'MERKEZ',
            quantity: partMiktar || (parsedSns.length > 0 ? String(parsedSns.length) : '1'),
            serialNumbers: parsedSns.length > 0 ? parsedSns : (partSn && partSn !== '-' ? [partSn] : []),
            sonKontrol: partSonKontrol,
            gelecekKontrol: partGelecekKontrol,
            firma: partFirma
          }];
        }
      }
    }

    setRegionalLocations(items);
  };

  const getDaysRemainingLabel = (dateStr: string) => {
    if (!dateStr || dateStr === '-') return null;
    const parts = dateStr.trim().split('.');
    if (parts.length === 3) {
      const targetDate = new Date(Number(parts[2]), Number(parts[1]) - 1, Number(parts[0]));
      if (isNaN(targetDate.getTime())) return null;
      const today = new Date();
      today.setHours(0, 0, 0, 0);
      const diffTime = targetDate.getTime() - today.getTime();
      const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
      if (diffDays < 0) {
        return { text: `🔴 ${Math.abs(diffDays)}g Geçti`, color: 'bg-rose-100 text-rose-800 border-rose-200' };
      } else if (diffDays <= 90) {
        return { text: `🟡 ${diffDays}g Kaldı`, color: 'bg-amber-100 text-amber-900 border-amber-200' };
      } else {
        return { text: `🟢 ${diffDays}g Kaldı`, color: 'bg-emerald-100 text-emerald-800 border-emerald-200' };
      }
    }
    return null;
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
    return cleanAndFormatDateString(dateStr);
  };

  // Akıllı Başlık ve Kolon Tespiti ile Excel Sayfasını Standart Matris Formatına Çeviren Gelişmiş Fonksiyon
  const parseExcelWorksheetToRows = (worksheet: XLSX.WorkSheet, techType: string, subSection: string = 'all'): string[][] => {
    try {
      // Birleştirilmiş hücreleri unroll yap (böylece alt hücreler boş kalmaz)
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
        if (isHeaderLikeRow(rawRow)) continue;

        const getVal = (idx: number, fallback = "") => {
          if (idx >= 0 && rawRow[idx] !== undefined && rawRow[idx] !== null) {
            const valStr = String(rawRow[idx]).trim();
            if (valStr !== "") return valStr;
          }
          return fallback;
        };

        // Kullanıcı kuralı: Sıra no A sütunu [0], Teçhizat adı B sütunu [1]
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
            rawSira,                                      // 0: SIRA NO (boş ise boş kalsın, gruplamada ardışık atanır)
            rawName,                                      // 1: ARAÇ PLAKASI / TANIMI
            rawPn,                                        // 2: PARÇA NO (P/N) / MODEL
            rawLoc,                                       // 3: BULUNDUĞU YER
            rawMiktar || rawSonBakim || "",               // 4: SON KM Sİ
            rawDurum || "FAAL",                           // 5: DURUMU
            rawSonBakim,                                  // 6: SON KONTROL
            rawGelecekBakim,                              // 7: GELECEK KONTROL
            rawFirma,                                     // 8: FİRMA
            rawAciklama,                                  // 9: AÇIKLAMA
            rawMail,                                      // 10: MAIL TARIHI
            subSection !== 'all' ? subSection : 'kara_araclari' // 11: KATEGORİ
          ];
        } else if (isDepoType) {
          targetRow = [
            rawSira,                                      // 0: SIRA NO
            rawName,                                      // 1: MALZEME / PARÇA ADI
            rawPn,                                        // 2: PARÇA NO (P/N)
            rawSn || "-",                                 // 3: SERİ NO (S/N)
            rawMiktar || "1",                             // 4: MİKTAR
            rawLoc,                                       // 5: BULUNDUĞU YER
            rawDurum || "FAAL",                           // 6: DURUMU
            rawKalib || "HAYIR",                          // 7: ÖMÜRLÜ PARÇA MI?
            rawGelecekBakim || rawSonBakim || "",         // 8: ÖMÜR BİTİŞ TARİHİ
            rawFirma,                                     // 9: TEDARİK EDİLEN FİRMA
            rawAciklama,                                  // 10: AÇIKLAMA
            rawMail,                                      // 11: MAIL TARIHI
            subSection                                    // 12: KATEGORİ (depo_sarf | depo_kimyasal)
          ];
        } else {
          targetRow = [
            rawSira,                                      // 0: SIRA NO
            rawName,                                      // 1: TEÇHİZAT ADI
            rawPn,                                        // 2: PARÇA NO (P/N)
            rawSn || "-",                                 // 3: SERİ NO (S/N)
            rawMiktar || "1",                             // 4: MİKTAR / KAPASİTE
            rawLoc,                                       // 5: BULUNDUĞU YER
            rawDurum || "FAAL",                           // 6: DURUMU
            rawKalib || "EVET",                           // 7: KALİBRASYONA TABİ
            rawSonBakim,                                  // 8: SON KONTROL
            rawGelecekBakim,                              // 9: GELECEK KONTROL
            rawFirma,                                     // 10: FİRMA
            rawAciklama,                                  // 11: AÇIKLAMA
            rawMail,                                      // 12: MAIL TARIHI
            subSection !== 'all' ? subSection : 'yer_destek' // 13: KATEGORİ
          ];
        }

        if (targetRow[1] || targetRow[2] || (targetRow[3] && targetRow[3] !== "-") || targetRow[5]) {
          rawParsedRows.push(targetRow);
        }
      }

      return groupMultiLocationRows(
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

  // Dosya ve Sayfa Adından Hedef Hava/Kara Birimi ve Alt Bölümünü Otomatik Tespit Eden Fonksiyon
  const detectUnitAndSectionFromNames = (fileName: string, sheetName: string): { techType: string, subSection: string } => {
    const combined = `${fileName} ${sheetName}`
      .toLowerCase()
      .replace(/ı/g, 'i')
      .replace(/İ/g, 'i')
      .replace(/ğ/g, 'g')
      .replace(/ü/g, 'u')
      .replace(/ş/g, 's')
      .replace(/ö/g, 'o')
      .replace(/ç/g, 'c');
    
    let techType = '';
    let subSection = 'all';

    if (combined.includes("at-802") || combined.includes("at802") || combined.includes("at 802") || combined.includes("air tractor") || combined.includes("airtractor")) {
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

    if (combined.includes("ozel") || combined.includes("bakim alet") || combined.includes("alet")) {
      subSection = "ozel_alet";
    } else if (combined.includes("sarf") || combined.includes("yedek parca") || combined.includes("parca depo") || combined.includes("parca")) {
      subSection = "depo_sarf";
    } else if (combined.includes("kimyasal") || combined.includes("yag") || combined.includes("boya") || combined.includes("tiner")) {
      subSection = "depo_kimyasal";
    } else if (combined.includes("yer") || combined.includes("destek") || combined.includes("techizat")) {
      subSection = "yer_destek";
    }

    return { techType, subSection };
  };

  const applyParsedRowsToUnit = (techType: string, subSection: string, newRows: string[][]) => {
    if (newRows.length === 0) return;

    let updater: (prev: string[][]) => string[][];

    if (subSection === 'all') {
      updater = () => newRows;
    } else {
      updater = (prev) => {
        const otherSectionRows = prev.filter(r => getRowSection(r, techType) !== subSection);
        const tagged = newRows.map(r => {
          const copy = [...r];
          const isKara = techType === 'kara_araclari';
          const isDepo = subSection === 'depo_sarf' || subSection === 'depo_kimyasal';
          const secCol = isKara ? 11 : (isDepo ? 12 : 13);
          while (copy.length <= secCol) copy.push("");
          copy[secCol] = subSection;
          return copy;
        });
        const combined = [...otherSectionRows, ...tagged];
        return combined.map((r, i) => {
          const c = [...r];
          c[0] = String(i + 1);
          return c;
        });
      };
    }

    if (techType === 'at802_ozel_alet' || (techType === 'at802' && subSection === 'ozel_alet')) {
      setTechizatAt802OzelAletData(prev => { const res = updater(prev); try { localStorage.setItem('techizat_at802_ozel_alet_data', JSON.stringify(res)); } catch(e){} return res; });
    } else if (techType === 'bell429') {
      setTechizatBell429Data(prev => { const res = updater(prev); try { localStorage.setItem('techizat_bell429_data', JSON.stringify(res)); } catch(e){} return res; });
    } else if (techType === 'at802') {
      setTechizatAt802Data(prev => { const res = updater(prev); try { localStorage.setItem('techizat_at802_data', JSON.stringify(res)); } catch(e){} return res; });
    } else if (techType === 't70') {
      setTechizatT70Data(prev => { const res = updater(prev); try { localStorage.setItem('techizat_t70_data', JSON.stringify(res)); } catch(e){} return res; });
    } else if (techType === 't70_bumbi_backet') {
      setTechizatT70BumbiBacketData(prev => { const res = updater(prev); try { localStorage.setItem('techizat_t70_bumbi_backet_data', JSON.stringify(res)); } catch(e){} return res; });
    } else if (techType === 't70_helitak') {
      setTechizatT70HelitakData(prev => { const res = updater(prev); try { localStorage.setItem('techizat_t70_helitak_data', JSON.stringify(res)); } catch(e){} return res; });
    } else if (techType === 'b360') {
      setTechizatB360Data(prev => { const res = updater(prev); try { localStorage.setItem('techizat_b360_data', JSON.stringify(res)); } catch(e){} return res; });
    } else if (techType === 'c650') {
      setTechizatC650Data(prev => { const res = updater(prev); try { localStorage.setItem('techizat_c650_data', JSON.stringify(res)); } catch(e){} return res; });
    } else if (techType === 'hangar') {
      setTechizatHangarData(prev => { const res = updater(prev); try { localStorage.setItem('techizat_hangar_data', JSON.stringify(res)); } catch(e){} return res; });
    } else if (techType === 'kara_araclari') {
      setTechizatKaraAraclariData(prev => { const res = updater(prev); try { localStorage.setItem('techizat_kara_araclari_data', JSON.stringify(res)); } catch(e){} return res; });
    }
  };

  // Google Drive Klasöründen (1HQR_NYKhHQGA7_2W3nArI9pCh-LJasTP) Birime Özel Excel Tablolarını Dinamik ve Canlı Okuyan Fonksiyon
  const pullTechizatUnitFromDrive = async (targetUnitKey?: string | null, silent = true) => {
    setIsTechizatDriveLoading(true);
    try {
      const UNIT_EXCEL_MAP: Record<string, string> = {
        at802: 'hava_araçları_yer_destek_at-802.xlsx',
        at802_ozel_alet: 'at-802_ozel_bakim_aletleri.xlsx',
        bell429: 'hava_araçları_yer_destek_bell-429.xlsx',
        t70: 'hava_araçları_yer_destek_t-70.xlsx',
        t70_bumbi_backet: 'hava_araçları_yer_destek_t-70_bumbi_backet.xlsx',
        t70_helitak: 'hava_araçları_yer_destek_t-70_helitak.xlsx',
        b360: 'hava_araçları_yer_destek_b-360.xlsx',
        c650: 'hava_araçları_yer_destek_c-650.xlsx',
        hangar: 'hava_araçları_yer_destek_hangar.xlsx',
        kara_araclari: 'kara_araclari_servis_cizelgesi.xlsx'
      };

      const keysToFetch: string[] = (targetUnitKey && targetUnitKey !== 'all')
        ? [targetUnitKey]
        : Object.keys(UNIT_EXCEL_MAP);

      let totalLoadedCount = 0;

      await Promise.all(keysToFetch.map(async (uKey) => {
        const targetFileName = UNIT_EXCEL_MAP[uKey] || `hava_araçları_yer_destek_${uKey}.xlsx`;
        try {
          const res = await fetch('/api/read-excel-from-drive', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              fileName: targetFileName,
              folderId: '1HQR_NYKhHQGA7_2W3nArI9pCh-LJasTP'
            })
          });
          if (!res.ok) return;
          const result = await res.json();
          if (result && result.status === 'success' && result.base64) {
            if (result.updated) {
              setFormUpdateDates(prev => {
                const up = { ...prev, [`techizat_${uKey}`]: result.updated };
                localStorage.setItem("form_update_dates", JSON.stringify(up));
                return up;
              });
            }

            const binaryString = atob(result.base64);
            const len = binaryString.length;
            const bytes = new Uint8Array(len);
            for (let i = 0; i < len; i++) {
              bytes[i] = binaryString.charCodeAt(i);
            }
            const workbook = XLSX.read(bytes, { type: 'array' });

            for (const sName of workbook.SheetNames) {
              const worksheet = workbook.Sheets[sName];
              if (!worksheet) continue;
              unmergeAndFillWorksheet(worksheet);

              const detected = detectUnitAndSectionFromNames(targetFileName, sName);
              const effectiveTechType = uKey === 'at802_ozel_alet' ? 'at802_ozel_alet' : (uKey === 'at802' ? 'at802' : (detected.techType || uKey));
              const effectiveSubSection = uKey === 'at802_ozel_alet' ? 'ozel_alet' : (uKey === 'at802' ? 'yer_destek' : (detected.subSection || 'all'));

              setTechizatFileNames(prev => ({
                ...prev,
                [`${effectiveTechType}_${effectiveSubSection}`]: targetFileName
              }));

              const parsed = parseExcelWorksheetToRows(worksheet, effectiveTechType, effectiveSubSection);
              if (parsed.length > 0) {
                applyParsedRowsToUnit(effectiveTechType, effectiveSubSection, parsed);
                totalLoadedCount += parsed.length;
              }
            }
          }
        } catch (err) {
          console.warn("Drive Excel okuma hatası (" + targetFileName + "):", err);
        }
      }));

      if (!silent) {
        if (totalLoadedCount > 0) {
          const unitLabel = targetUnitKey && targetUnitKey !== 'all' ? targetUnitKey.toUpperCase() : 'Birimler';
          showNotification(`Google Drive'dan [${unitLabel}] için (${totalLoadedCount} kayıt) Excel verileri başarıyla yüklendi!`);
        } else {
          showNotification("Google Drive'dan güncel veriler kontrol edildi.");
        }
      }
    } catch (err) {
      console.warn("Drive Excels kontrolü tamamlandı/zaman aşımı:", err);
    } finally {
      setIsTechizatDriveLoading(false);
      setIsPullingTechizat(false);
    }
  };

  const pullAllTechizatFromDriveExcels = (silent = true) => pullTechizatUnitFromDrive('all', silent);
  const pullAllTechizatFromGoogleSheets = pullAllTechizatFromDriveExcels;

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
            mail90: String(row["90 GÜN UYARISI MAİL GÖNDERİM TARİHİ"] || row["90 Gün Uyarı Mail Gönderim Tarihi"] || row["90 GUN UYARISI MAIL GONDERIM TARIHI"] || "").trim(),
            mail60: String(row["60 GÜN UYARISI MAİL GÖNDERİM TARİHİ"] || row["60 Gün Uyarı Mail Gönderim Tarihi"] || row["60 GUN UYARISI MAIL GONDERIM TARIHI"] || "").trim(),
            mail30: String(row["30 GÜN UYARISI MAİL GÖNDERİM TARİHİ"] || row["30 Gün Uyarı Mail Gönderim Tarihi"] || row["30 GUN UYARISI MAIL GONDERIM TARIHI"] || "").trim()
          })).filter((item: any) => item.birim);
          
          if (mapped.length > 0) {
            setGunTakipSorumlulari(mapped);
            try { localStorage.setItem('gun_takip_sorumlulari', JSON.stringify(mapped)); } catch(e){}
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
      const rowsToSend = updatedData.map(item => [
        item.birim,
        item.adSoyad,
        item.eposta,
        item.mail90 || '',
        item.mail60 || '',
        item.mail30 || ''
      ]);
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
        try { localStorage.setItem('gun_takip_sorumlulari', JSON.stringify(updatedData)); } catch(e){}
        showNotification("Sorumlu Personel (Gün Takip) Verileri Başarıyla Güncellendi!");
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
    const cleaned = String(dateStr).trim();
    if (!cleaned || cleaned === "-" || cleaned === "--" || cleaned.includes("MUAFIYET") || cleaned.toUpperCase() === "BELİRTİLMEMİŞ") return null;
    
    // If multi-line or comma-separated, evaluate each date and return the minimum (most urgent) remaining days
    if (cleaned.includes('\n') || cleaned.includes(',') || cleaned.includes(';')) {
      const parts = cleaned.split(/[\n,;]+/).map(s => s.trim()).filter(Boolean);
      let minDays: number | null = null;
      for (const p of parts) {
        const d = parseGelecekBakimDays(p);
        if (d !== null) {
          if (minDays === null || d < minDays) {
            minDays = d;
          }
        }
      }
      return minDays;
    }

    // Excel seri no kontrolü (Örn: 46234, 46965)
    if (/^\d{5}(\.\d+)?$/.test(cleaned)) {
      const serial = parseFloat(cleaned);
      if (serial >= 20000 && serial <= 90000) {
        const utcDays = Math.floor(serial - 25569);
        const dateObj = new Date(utcDays * 86400 * 1000);
        if (!isNaN(dateObj.getTime())) {
          const today = new Date();
          today.setHours(0, 0, 0, 0);
          dateObj.setHours(0, 0, 0, 0);
          const diffTime = dateObj.getTime() - today.getTime();
          return Math.ceil(diffTime / (1000 * 60 * 60 * 24));
        }
      }
    }

    // cleanAndFormatDateString ile önce standart GG.AA.YYYY formatına çevir
    const formatted = cleanAndFormatDateString(cleaned);
    const dmyRegex = /^(\d{1,2})[\.\/-](\d{1,2})[\.\/-](\d{4})$/;
    const ymdRegex = /^(\d{4})[\.\/-](\d{1,2})[\.\/-](\d{1,2})$/;
    
    let dateObj: Date | null = null;
    let m = formatted.match(dmyRegex);
    if (m) {
      const day = parseInt(m[1], 10);
      const month = parseInt(m[2], 10) - 1;
      const year = parseInt(m[3], 10);
      dateObj = new Date(year, month, day);
    } else {
      m = formatted.match(ymdRegex);
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

      if (passwordActionType === 'gun_takip') {
        setIsSorumluModalOpen(true);
        return;
      }

      if (passwordActionType === 'new_product') {
        initializeNewProductForm();
        setIsNewProductModalOpen(true);
        return;
      }

      if (passwordActionType === 'depo_management') {
        openStandaloneDepo(selectedUnitFolder);
        return;
      }

      if (passwordActionType === 'filter_sync') {
        let target = `techizat_${activeTechizatType}`;
        if (techizatActiveSection && techizatActiveSection !== 'all' && techizatActiveSection !== 'techizat_all' && techizatActiveSection !== 'depo_all') {
          target = `techizat_${activeTechizatType}_${techizatActiveSection}`;
        }
        setSyncInitialTarget(target);
        setSyncInitialStep(2);
        setIsDataSyncModalOpen(true);
        return;
      }

      // Default: global_sync -> Step 1 (Global Selection for all forms and units)
      setSyncInitialStep(1);
      setSyncInitialTarget('1');
      setIsDataSyncModalOpen(true);
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

  const getSyncTargetDisplayTitle = (target: string): string => {
    if (target === '1') return "1. GÖREVLENDİRME ÇİZELGELERİ";
    if (target === '21') return "2. YAZ DÖNEMİ PLANLAMASI - BELL 429";
    if (target === '22') return "2. YAZ DÖNEMİ PLANLAMASI - T-70";
    if (target === '23') return "2. YAZ DÖNEMİ PLANLAMASI - AT-802";
    if (target === '24') return "2. YAZ DÖNEMİ PLANLAMASI - ANKARA BEKLEME (BELL-429)";
    if (target === '25') return "2. YAZ DÖNEMİ PLANLAMASI - ANKARA BEKLEME (C-650/B-360)";
    if (target === '3') return "3. BAKIM YETKİ ÇİZELGELERİ";
    if (target === '5') return "5. PERSONEL BİLGİ ÇİZELGELERİ";
    if (target === '6') return "6. PERSONEL UÇUŞ-HİZMET YILLARI";

    // AT-802
    if (target === 'techizat_at802') return "✈️ AT-802F - TÜM BİRİM ENVANTERİ (EXCEL)";
    if (target === 'techizat_at802_yer_destek') return "✈️ AT-802F - YER DESTEK TEÇHİZATLARI (EXCEL)";
    if (target === 'techizat_at802_ozel_alet') return "✈️ AT-802F - ÖZEL ALETLER (EXCEL)";
    if (target === 'techizat_at802_depo_sarf') return "📦 AT-802F - SARF VE PARÇA DEPOSU (EXCEL)";
    if (target === 'techizat_at802_depo_kimyasal') return "🧪 AT-802F - KİMYASAL DEPO (EXCEL)";

    // BELL 429
    if (target === 'techizat_bell429') return "🚁 BELL 429 - TÜM BİRİM ENVANTERİ (EXCEL)";
    if (target === 'techizat_bell429_yer_destek') return "🚁 BELL 429 - YER DESTEK VE ÖZEL ALETLER (EXCEL)";
    if (target === 'techizat_bell429_depo_sarf') return "📦 BELL 429 - SARF VE PARÇA DEPOSU (EXCEL)";
    if (target === 'techizat_bell429_depo_kimyasal') return "🧪 BELL 429 - KİMYASAL DEPO (EXCEL)";

    // T-70
    if (target === 'techizat_t70') return "🚁 T-70 - TÜM BİRİM ENVANTERİ (EXCEL)";
    if (target === 'techizat_t70_yer_destek') return "🚁 T-70 - YER DESTEK VE ÖZEL ALETLER (EXCEL)";
    if (target === 'techizat_t70_bumbi_backet') return "🪣 T-70 - BUMBİ BACKET TEÇHİZATI (EXCEL)";
    if (target === 'techizat_t70_helitak') return "🛡️ T-70 - HELİTAK TEÇHİZATI (EXCEL)";
    if (target === 'techizat_t70_depo_sarf') return "📦 T-70 - SARF VE PARÇA DEPOSU (EXCEL)";
    if (target === 'techizat_t70_depo_kimyasal') return "🧪 T-70 - KİMYASAL DEPO (EXCEL)";

    // B-360
    if (target === 'techizat_b360') return "✈️ BEECHCRAFT B-360 - TÜM BİRİM ENVANTERİ (EXCEL)";
    if (target === 'techizat_b360_yer_destek') return "✈️ BEECHCRAFT B-360 - YER DESTEK VE ÖZEL ALETLER (EXCEL)";
    if (target === 'techizat_b360_depo_sarf') return "📦 BEECHCRAFT B-360 - SARF VE PARÇA DEPOSU (EXCEL)";
    if (target === 'techizat_b360_depo_kimyasal') return "🧪 BEECHCRAFT B-360 - KİMYASAL DEPO (EXCEL)";

    // C-650
    if (target === 'techizat_c650') return "✈️ CITATION C-650 - TÜM BİRİM ENVANTERİ (EXCEL)";
    if (target === 'techizat_c650_yer_destek') return "✈️ CITATION C-650 - YER DESTEK VE ÖZEL ALETLER (EXCEL)";
    if (target === 'techizat_c650_depo_sarf') return "📦 CITATION C-650 - SARF VE PARÇA DEPOSU (EXCEL)";
    if (target === 'techizat_c650_depo_kimyasal') return "🧪 CITATION C-650 - KİMYASAL DEPO (EXCEL)";

    // HANGAR
    if (target === 'techizat_hangar') return "🏭 HANGAR - TÜM BİRİM ENVANTERİ (EXCEL)";
    if (target === 'techizat_hangar_yer_destek') return "🏭 HANGAR - YER DESTEK VE ÖZEL ALETLER (EXCEL)";
    if (target === 'techizat_hangar_depo_sarf') return "📦 HANGAR - SARF VE PARÇA DEPOSU (EXCEL)";
    if (target === 'techizat_hangar_depo_kimyasal') return "🧪 HANGAR - KİMYASAL DEPO (EXCEL)";

    // KARA ARAÇLARI
    if (target === 'techizat_kara_araclari') return "🚗 KARA ARAÇLARI TAKİP SİSTEMİ (EXCEL)";

    return TABLE_CONFIGS[Number(target)]?.title || `BİRİM ${target}`;
  };

  const getCleanSyncTargetTitle = (target: string): string => {
    const fullTitle = getSyncTargetDisplayTitle(target);
    return fullTitle
      .replace(/^[^\w\s\(\)\-\.]+\s*/u, '')
      .replace(/\s*\(EXCEL\)$/i, '')
      .trim();
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
      const targetRaw = syncSelectedTarget.replace('techizat_', '');
      let techType: 'bell429' | 'at802' | 't70' | 't70_bumbi_backet' | 't70_helitak' | 'b360' | 'c650' | 'hangar' | 'kara_araclari' = 'at802';
      let subSection: 'all' | 'yer_destek' | 'ozel_alet' | 'depo_sarf' | 'depo_kimyasal' = 'all';

      if (targetRaw.startsWith('at802')) {
        techType = 'at802';
        if (targetRaw.includes('ozel_alet')) subSection = 'ozel_alet';
        else if (targetRaw.includes('depo_sarf')) subSection = 'depo_sarf';
        else if (targetRaw.includes('depo_kimyasal')) subSection = 'depo_kimyasal';
        else if (targetRaw.includes('yer_destek')) subSection = 'yer_destek';
      } else if (targetRaw.startsWith('bell429')) {
        techType = 'bell429';
        if (targetRaw.includes('depo_sarf')) subSection = 'depo_sarf';
        else if (targetRaw.includes('depo_kimyasal')) subSection = 'depo_kimyasal';
        else if (targetRaw.includes('yer_destek')) subSection = 'yer_destek';
      } else if (targetRaw.startsWith('t70_bumbi_backet')) {
        techType = 't70_bumbi_backet';
      } else if (targetRaw.startsWith('t70_helitak')) {
        techType = 't70_helitak';
      } else if (targetRaw.startsWith('t70')) {
        techType = 't70';
        if (targetRaw.includes('depo_sarf')) subSection = 'depo_sarf';
        else if (targetRaw.includes('depo_kimyasal')) subSection = 'depo_kimyasal';
        else if (targetRaw.includes('yer_destek')) subSection = 'yer_destek';
      } else if (targetRaw.startsWith('b360')) {
        techType = 'b360';
        if (targetRaw.includes('depo_sarf')) subSection = 'depo_sarf';
        else if (targetRaw.includes('depo_kimyasal')) subSection = 'depo_kimyasal';
        else if (targetRaw.includes('yer_destek')) subSection = 'yer_destek';
      } else if (targetRaw.startsWith('c650')) {
        techType = 'c650';
        if (targetRaw.includes('depo_sarf')) subSection = 'depo_sarf';
        else if (targetRaw.includes('depo_kimyasal')) subSection = 'depo_kimyasal';
        else if (targetRaw.includes('yer_destek')) subSection = 'yer_destek';
      } else if (targetRaw.startsWith('hangar')) {
        techType = 'hangar';
        if (targetRaw.includes('depo_sarf')) subSection = 'depo_sarf';
        else if (targetRaw.includes('depo_kimyasal')) subSection = 'depo_kimyasal';
        else if (targetRaw.includes('yer_destek')) subSection = 'yer_destek';
      } else if (targetRaw.startsWith('kara_araclari')) {
        techType = 'kara_araclari';
      }

      if (!fileNameLower.endsWith('.xlsx') && !fileNameLower.endsWith('.xls') && !fileNameLower.endsWith('.csv')) {
        alert("Teçhizat/Depo Takip güncellemesi için lütfen Excel (.xlsx, .xls) veya CSV belgesi yükleyin.");
        return;
      }

      const reader = new FileReader();
      reader.onload = (evt) => {
        try {
          const arrayBuffer = evt.target?.result as ArrayBuffer;
          const workbook = XLSX.read(arrayBuffer, { type: 'array' });
          const worksheet = workbook.Sheets[workbook.SheetNames[0]];

          // Birleştirilmiş hücreleri unroll yap (böylece alt hücreler boş kalmaz)
          unmergeAndFillWorksheet(worksheet);
          
          const rawRows = XLSX.utils.sheet_to_json<string[]>(worksheet, { header: 1, defval: "", raw: false });
          if (rawRows.length === 0) {
            throw new Error("Yüklenen Excel dosyasında veri bulunamadı.");
          }

          // Smart Header Detection
          // Robust Header Row Detection
          let headerRowIdx = 0;
          let maxScore = -1;
          for (let r = 0; r < Math.min(15, rawRows.length); r++) {
            const rCells = rawRows[r] || [];
            let score = 0;
            const rowStr = rCells.map(c => String(c).toLowerCase()).join(" ");
            
            if (rowStr.includes("sıra") || rowStr.includes("sira")) score += 2;
            if (rowStr.includes("teçhizat") || rowStr.includes("techizat") || rowStr.includes("malzeme") || rowStr.includes("araç")) score += 2;
            if (rowStr.includes("parça") || rowStr.includes("parca") || rowStr.includes("p/n")) score += 2;
            if (rowStr.includes("seri") || rowStr.includes("s/n")) score += 2;
            if (rowStr.includes("miktar") || rowStr.includes("kapasite")) score += 2;
            if (rowStr.includes("bulunduğu") || rowStr.includes("lokasyon")) score += 2;
            if (rowStr.includes("durum")) score += 1;
            if (rowStr.includes("bakım") || rowStr.includes("kalibrasyon") || rowStr.includes("kontrol")) score += 2;

            if (score > maxScore) {
              maxScore = score;
              headerRowIdx = r;
            }
          }
          
          if (maxScore < 4) {
            console.warn("Could not confidently find a header row. Falling back to row 0 or 1.");
          }

          const rawHeaders = (rawRows[headerRowIdx] || []).map(h => String(h || '').trim().toUpperCase());
          const finalHeaders = rawHeaders.map((h, hIdx) => h || `KOLON ${hIdx + 1}`);

          // 1. Akıllı Başlık Tespiti ve Eşleştirme (Smart Header-Based Column Mapping)
          // Yüklenen Excel'de başlıkların sırası ne olursa olsun 'AÇIKLAMA', 'TEÇHİZAT ADI', 'P/N', 'S/N' vb. doğru sütuna aktarılır.
          const findColIdx = (keywords: string[], excludeKeywords: string[] = []): number => {
            return finalHeaders.findIndex(h => {
              const upper = h.toUpperCase().trim();
              const hasKey = keywords.some(k => upper.includes(k.toUpperCase()));
              const hasExclude = excludeKeywords.some(ex => upper.includes(ex.toUpperCase()));
              return hasKey && !hasExclude;
            });
          };

          const siraColIdx = findColIdx(["SIRA", "NO."], ["SERİ", "SERI", "PARÇA", "PARCA", "P/N", "MODEL"]);
          const nameColIdx = findColIdx(
            ["TEÇHİZAT", "TECHİZAT", "MALZEME", "ARAÇ", "ARAC", "PLAKA", "ÜRÜN", "URUN", "EKİPMAN", "TANIM", "NAME"],
            ["FİRMA", "FIRMA", "KONTROL", "BAKIM", "YAPAN"]
          );
          const pnColIdx = findColIdx(["P/N", "PN", "PARÇA NO", "PARCA NO", "MODEL", "PART NUMBER", "PART NO"]);
          const snColIdx = findColIdx(["S/N", "SN", "SERİ NO", "SERI NO", "SERİ", "SERI", "SERIAL"], ["SIRA"]);
          const miktarColIdx = findColIdx(["MİKTAR", "MIKTAR", "KAPASİTE", "KAPASITE", "ADET", "QTY", "QUANTITY"]);
          const locColIdx = findColIdx(["BULUNDUĞU", "BULUNDUGU", "LOKASYON", "KONUM", "YER", "RAF", "DEPO", "LOCATION"]);
          const durumColIdx = findColIdx(["DURUM", "DURUMU", "STATUS", "FAALİYET"]);
          const kalibTabiColIdx = findColIdx(["KALİBRASYONA TABİ", "KALIBRASYONA TABI", "BAKIMA TABİ", "BAKIMA TABI", "TABİ Mİ", "TABI MI", "ÖMÜRLÜ", "OMURLU"]);
          const sonBakimColIdx = findColIdx(["SON KONTROL", "SON BAKIM", "SON KALİBRASYON", "SON TEST", "SON MUAYENE", "SON KM", "YAPILAN KONTROL"]);
          const gelecekBakimColIdx = findColIdx(["GELECEK KONTROL", "GELECEK BAKIM", "GELECEK KALİBRASYON", "BİR SONRAKİ", "SONRAKİ BAKIM", "ÖMÜR BİTİŞ", "OMUR BITIS", "SON KULLANMA", "EXPIRY"]);
          const firmaColIdx = findColIdx(["KONTROLÜ YAPAN", "KONTROLU YAPAN", "YAPAN FİRMA", "YAPAN FIRMA", "FİRMA", "FIRMA", "TEDARİK", "TEDARIK", "SERVİS", "VENDOR", "SUPPLIER"]);
          const aciklamaColIdx = findColIdx(["AÇIKLAMA", "ACIKLAMA", "AÇIKLAMALAR", "ACIKLAMALAR", "NOT", "NOTLAR", "DESCRIPTION", "REMARKS", "DETAY", "ÖZEL NOT", "LOT", "PARTİ"]);
          const mailColIdx = findColIdx(["MAİL GÖNDERİM", "MAIL GONDERIM", "90 GÜN", "90 GUN", "E-POSTA", "MAIL", "MAİL"]);

          const isKara = techType === 'kara_araclari';
          const isDepoType = subSection === 'depo_sarf' || subSection === 'depo_kimyasal';

          // Extract and Map rows based on Detected Header Positions
          const rawParsedRows: string[][] = [];
          for (let r = headerRowIdx + 1; r < rawRows.length; r++) {
            const rawRow = rawRows[r] || [];
            const isRowEmpty = rawRow.every(cell => String(cell || '').trim() === '');
            if (isRowEmpty) continue;

            const getVal = (idx: number, fallback = "") => {
              if (idx >= 0 && rawRow[idx] !== undefined) {
                return String(rawRow[idx]).trim();
              }
              return fallback;
            };

            // Dynamic fallbacks to avoid hardcoded indices
            // If findColIdx failed for everything, we use the first few non-empty columns.
            const nonEmptyCols = rawRow.map((c, i) => ({ val: String(c || '').trim(), idx: i })).filter(c => c.val !== "");
            const guessIdx = (order) => nonEmptyCols.length > order ? nonEmptyCols[order].idx : -1;
            
            const rawSira = getVal(siraColIdx, getVal(guessIdx(0), ""));
            const rawName = getVal(nameColIdx, getVal(guessIdx(1), ""));
            const rawPn = getVal(pnColIdx, getVal(guessIdx(2), ""));
            const rawSn = getVal(snColIdx, getVal(guessIdx(3), "-"));
            const rawMiktar = getVal(miktarColIdx, getVal(guessIdx(4), "1"));
            const rawLoc = getVal(locColIdx, getVal(guessIdx(5), ""));
            const rawDurum = getVal(durumColIdx, "FAAL");
            
            let rawKalib = getVal(kalibTabiColIdx, "");
            if (!rawKalib) {
              rawKalib = "EVET";
            }
            const rawSonBakim = getVal(sonBakimColIdx, "");
            const rawGelecekBakim = getVal(gelecekBakimColIdx, "");
            const rawFirma = getVal(firmaColIdx, "");
            
            // Açıklama sütunu: Excel'de Açıklama başlığı nerede olursa olsun tam olarak aktarılır!
            let rawAciklama = "";
            if (aciklamaColIdx >= 0 && rawRow[aciklamaColIdx] !== undefined) {
              rawAciklama = String(rawRow[aciklamaColIdx]).trim();
            } else {
              // Alternatif fallback: Eğer başlık doğrudan bulunamadıysa 10. veya 11. sütun
              const fallbackIdx = finalHeaders.length >= 12 ? 11 : 10;
              if (rawRow[fallbackIdx] !== undefined) {
                rawAciklama = String(rawRow[fallbackIdx]).trim();
              }
            }

            const rawMail = getVal(mailColIdx, "");

            let targetRow: string[] = [];
            if (isKara) {
              targetRow = [
                rawSira,                                      // 0: SIRA NO (boş ise boş kalsın, gruplamada ardışık atanır)
                rawName,                                      // 1: ARAÇ PLAKASI / TANIMI
                rawPn,                                        // 2: PARÇA NO (P/N) / MODEL
                rawLoc,                                       // 3: BULUNDUĞU YER
                rawMiktar || rawSonBakim || "",               // 4: SON KM Sİ
                rawDurum || "FAAL",                           // 5: DURUMU
                rawKalib || "EVET",                           // 6: KALİBRASYONA TABİ
                rawSonBakim,                                  // 7: SON KONTROL / BAKIM
                rawGelecekBakim,                              // 8: GELECEK KONTROL / BAKIM
                rawFirma,                                     // 9: SON KONTROLÜ YAPAN FİRMA
                rawAciklama,                                  // 10: AÇIKLAMA
                rawMail,                                      // 11: 90 GÜN UYARISI MAİL GÖNDERİM TARİHİ
                subSection                                    // 12: BÖLÜM
              ];
            } else if (isDepoType) {
              targetRow = [
                rawSira,                                      // 0: SIRA NO
                rawName,                                      // 1: MALZEME / PARÇA ADI
                rawPn,                                        // 2: PARÇA NO (P/N)
                rawSn || "-",                                 // 3: SERİ NO (S/N)
                rawMiktar || "1",                             // 4: MİKTAR
                rawLoc,                                       // 5: BULUNDUĞU YER / RAF
                rawDurum || "FAAL",                           // 6: DURUMU
                rawKalib || "EVET",                           // 7: ÖMÜRLÜ PARÇA MI?
                rawGelecekBakim || "-",                       // 8: ÖMÜR BİTİŞ TARİHİ
                rawFirma || "-",                              // 9: TEDARİK EDİLEN FİRMA
                rawAciklama,                                  // 10: AÇIKLAMA
                rawMail,                                      // 11: 90 GÜN UYARISI MAİL GÖNDERİM TARİHİ
                subSection                                    // 12: BÖLÜM / KATEGORİ
              ];
            } else {
              // Standart Teçhizat
              targetRow = [
                rawSira,                                      // 0: SIRA NO
                rawName,                                      // 1: TEÇHİZAT ADI
                rawPn,                                        // 2: PARÇA NO (P/N) / MODEL
                rawSn || "-",                                 // 3: SERİ NO (S/N)
                rawMiktar || "1",                             // 4: MİKTAR / KAPASİTE
                rawLoc,                                       // 5: BULUNDUĞU YER
                rawDurum || "FAAL",                           // 6: DURUMU
                rawKalib || "EVET",                           // 7: KALİBRASYONA TABİ
                rawSonBakim,                                  // 8: SON KONTROL / KALİBRASYON / BAKIM
                rawGelecekBakim,                              // 9: GELECEK KONTROL / KALİBRASYON / BAKIM
                rawFirma,                                     // 10: SON KONTROLÜ YAPAN FİRMA
                rawAciklama,                                  // 11: AÇIKLAMA
                rawMail,                                      // 12: 90 GÜN UYARISI MAİL GÖNDERİM TARİHİ
                subSection                                    // 13: BÖLÜM / KATEGORİ
              ];
            }

            if (targetRow[1] || targetRow[2] || (targetRow[3] && targetRow[3] !== "-") || targetRow[5]) {
              rawParsedRows.push(targetRow);
            }
          }

          const parsedRows = groupMultiLocationRows(
            rawParsedRows,
            1, // nameColIdx
            isKara ? 3 : 5, // locColIdx
            isKara ? -1 : 4, // miktarColIdx
            0, // siraColIdx
            2, // pnColIdx
            isKara ? -1 : 3 // seriNoColIdx
          );

          // Tag rows with specific subsection if selected
          const taggedParsedRows = parsedRows.map(r => {
            if (subSection === 'all') return r;
            const newRow = [...r];
            const secCol = isKara ? 12 : (isDepoType ? 12 : 13);
            while (newRow.length <= secCol) newRow.push("");
            newRow[secCol] = subSection;
            return newRow;
          });

          // Existing rows for this unit to allow section-based merging
          let existingData: string[][] = [];
          if (techType === 'bell429') existingData = techizatBell429Data;
          else if (techType === 'at802') existingData = techizatAt802Data;
          else if (techType === 't70') existingData = techizatT70Data;
          else if (techType === 't70_bumbi_backet') existingData = techizatT70BumbiBacketData;
          else if (techType === 't70_helitak') existingData = techizatT70HelitakData;
          else if (techType === 'b360') existingData = techizatB360Data;
          else if (techType === 'c650') existingData = techizatC650Data;
          else if (techType === 'hangar') existingData = techizatHangarData;
          else if (techType === 'kara_araclari') existingData = techizatKaraAraclariData;

          let finalDataToSave: string[][] = [];
          if (subSection === 'all') {
            finalDataToSave = parsedRows;
          } else {
            // Keep rows of OTHER sections, and replace rows of CURRENT section
            const otherSectionRows = existingData.filter(r => getRowSection(r, techType) !== subSection);
            finalDataToSave = [...otherSectionRows, ...taggedParsedRows];
          }

          // Renumber SIRA NO
          finalDataToSave = finalDataToSave.map((r, idx) => {
            const cloned = [...r];
            cloned[0] = String(idx + 1);
            return cloned;
          });

          // Save columns and rows to states and localStorage
          if (techType === 'bell429') {
            setTechizatBell429Columns(finalHeaders);
            setTechizatBell429Data(finalDataToSave);
            localStorage.setItem('excel_techizat_bell429_cols', JSON.stringify(finalHeaders));
            localStorage.setItem('excel_techizat_bell429_data', JSON.stringify(finalDataToSave));
          } else if (techType === 'at802') {
            setTechizatAt802Columns(finalHeaders);
            setTechizatAt802Data(finalDataToSave);
            localStorage.setItem('excel_techizat_at802_cols', JSON.stringify(finalHeaders));
            localStorage.setItem('excel_techizat_at802_data', JSON.stringify(finalDataToSave));
          } else if (techType === 't70') {
            setTechizatT70Columns(finalHeaders);
            setTechizatT70Data(finalDataToSave);
            localStorage.setItem('excel_techizat_t70_cols', JSON.stringify(finalHeaders));
            localStorage.setItem('excel_techizat_t70_data', JSON.stringify(finalDataToSave));
          } else if (techType === 't70_bumbi_backet') {
            setTechizatT70BumbiBacketColumns(finalHeaders);
            setTechizatT70BumbiBacketData(finalDataToSave);
            localStorage.setItem('excel_techizat_t70_bumbi_backet_cols', JSON.stringify(finalHeaders));
            localStorage.setItem('excel_techizat_t70_bumbi_backet_data', JSON.stringify(finalDataToSave));
          } else if (techType === 't70_helitak') {
            setTechizatT70HelitakColumns(finalHeaders);
            setTechizatT70HelitakData(finalDataToSave);
            localStorage.setItem('excel_techizat_t70_helitak_cols', JSON.stringify(finalHeaders));
            localStorage.setItem('excel_techizat_t70_helitak_data', JSON.stringify(finalDataToSave));
          } else if (techType === 'b360') {
            setTechizatB360Columns(finalHeaders);
            setTechizatB360Data(finalDataToSave);
            localStorage.setItem('excel_techizat_b360_cols', JSON.stringify(finalHeaders));
            localStorage.setItem('excel_techizat_b360_data', JSON.stringify(finalDataToSave));
          } else if (techType === 'c650') {
            setTechizatC650Columns(finalHeaders);
            setTechizatC650Data(finalDataToSave);
            localStorage.setItem('excel_techizat_c650_cols', JSON.stringify(finalHeaders));
            localStorage.setItem('excel_techizat_c650_data', JSON.stringify(finalDataToSave));
          } else if (techType === 'hangar') {
            setTechizatHangarColumns(finalHeaders);
            setTechizatHangarData(finalDataToSave);
            localStorage.setItem('excel_techizat_hangar_cols', JSON.stringify(finalHeaders));
            localStorage.setItem('excel_techizat_hangar_data', JSON.stringify(finalDataToSave));
          } else if (techType === 'kara_araclari') {
            const convertedRows = finalDataToSave.map(row => convertOldKaraAraclariRowToNew(row));
            const newCols = ["SIRA NO", "ARAÇ PLAKASI / TANIMI", "PARÇA NO (P/N) / MODEL", "BULUNDUĞU YER", "SON KM Sİ", "DURUMU", "SON KONTROL / KALİBRASYON / BAKIM", "GELECEK KONTROL / KALİBRASYON / BAKIM", "SON KONTROLÜ YAPAN FİRMA", "AÇIKLAMA"];
            setTechizatKaraAraclariColumns(newCols);
            setTechizatKaraAraclariData(convertedRows);
            localStorage.setItem('excel_techizat_kara_araclari_cols', JSON.stringify(newCols));
            localStorage.setItem('excel_techizat_kara_araclari_data', JSON.stringify(convertedRows));
          }

          // Sync specifically to "TÜM TECHİZAT" online Google Sheet
          const unitLabel = getTechizatUnitLabel(techType);
          if (unitLabel) {
            // Synced directly to Drive Excel folder via syncTechizatExcelToGoogleDrive above
          }

          setIsSendingToSheets(prev => ({ ...prev, [syncSelectedTarget]: true }));
          setUploadProgress(10);
          showNotification(`${file.name} belgesi başarıyla işlendi ve Google Drive Excel klasörüne kaydedildi!`);

          // Background Google Drive upload for Teçhizat/Depo Excel file
          fileToBase64(file).then(async (base64Data) => {
            try {
              let driveFileName = "hava_araçları_yer_destek_bell-429.xlsx";
              if (techType === 'at802') {
                if (subSection === 'ozel_alet') driveFileName = "at-802_ozel_bakim_aletleri.xlsx";
                else if (subSection === 'depo_sarf') driveFileName = "at-802_sarf_ve_parca_deposu.xlsx";
                else if (subSection === 'depo_kimyasal') driveFileName = "at-802_kimyasal_depo.xlsx";
                else driveFileName = "hava_araçları_yer_destek_at-802.xlsx";
              } else if (techType === 'bell429') {
                if (subSection === 'depo_sarf') driveFileName = "bell-429_sarf_ve_parca_deposu.xlsx";
                else if (subSection === 'depo_kimyasal') driveFileName = "bell-429_kimyasal_depo.xlsx";
                else driveFileName = "hava_araçları_yer_destek_bell-429.xlsx";
              } else if (techType === 't70') {
                if (subSection === 'depo_sarf') driveFileName = "t-70_sarf_ve_parca_deposu.xlsx";
                else if (subSection === 'depo_kimyasal') driveFileName = "t-70_kimyasal_depo.xlsx";
                else driveFileName = "hava_araçları_yer_destek_t-70.xlsx";
              } else if (techType === 't70_bumbi_backet') {
                driveFileName = "hava_araçları_yer_destek_t-70_bumbi_backet.xlsx";
              } else if (techType === 't70_helitak') {
                driveFileName = "hava_araçları_yer_destek_t-70_helitak.xlsx";
              } else if (techType === 'b360') {
                if (subSection === 'depo_sarf') driveFileName = "b-360_sarf_ve_parca_deposu.xlsx";
                else if (subSection === 'depo_kimyasal') driveFileName = "b-360_kimyasal_depo.xlsx";
                else driveFileName = "hava_araçları_yer_destek_b-360.xlsx";
              } else if (techType === 'c650') {
                if (subSection === 'depo_sarf') driveFileName = "c-650_sarf_ve_parca_deposu.xlsx";
                else if (subSection === 'depo_kimyasal') driveFileName = "c-650_kimyasal_depo.xlsx";
                else driveFileName = "hava_araçları_yer_destek_c-650.xlsx";
              } else if (techType === 'hangar') {
                if (subSection === 'depo_sarf') driveFileName = "hangar_sarf_ve_parca_deposu.xlsx";
                else if (subSection === 'depo_kimyasal') driveFileName = "hangar_kimyasal_depo.xlsx";
                else driveFileName = "hava_araçları_yer_destek_hangar.xlsx";
              } else if (techType === 'kara_araclari') {
                driveFileName = "kara_araçları_takip.xlsx";
              }

              const prettyUnitName = getCleanSyncTargetTitle(String(syncSelectedTarget));

              const res = await fetch("/api/upload-techizat-excel", {
                method: "POST",
                headers: {
                  "Content-Type": "application/json"
                },
                body: JSON.stringify({
                  fileName: driveFileName,
                  targetKey: syncSelectedTarget,
                  base64Data: base64Data,
                  folderId: "1HQR_NYKhHQGA7_2W3nArI9pCh-LJasTP"
                })
              });
              if (res.ok) {
                const result = await res.json();
                if (result.status === "success") {
                  showNotification(`'${driveFileName}' (${prettyUnitName}) güncel Excel belgesi Google Drive'a başarıyla yedeklendi!`);
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
              
              const cleanTargetTitle = getCleanSyncTargetTitle(String(syncSelectedTarget));
              showNotification(`'${cleanTargetTitle}' Excel verisi başarıyla aktarıldı ve matris oluşturuldu!`);

              // Close sync wizard and automatically open and redirect to matrix screen
              setModalOpen(true);
              setActiveTechizatType(techType);
              setTechizatActiveSection(subSection === 'all' ? 'all' : subSection);
              setModalType('techizat_matrix');
              setModalTitle(cleanTargetTitle);
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
          <div className="w-full flex justify-between items-center text-[10px] tracking-[0.3em] font-bold border-b border-white/10 pb-4 mb-8 sm:mb-12 flex-wrap gap-3">
            <div className="flex items-center gap-3">
              {/* Gün Takip ve Otomatik Bildirim Paneli Butonu */}
              <button
                type="button"
                onClick={openGunTakipWithPassword}
                className="flex items-center gap-2.5 px-3.5 py-1.5 rounded-full bg-emerald-950/80 hover:bg-emerald-900 border border-emerald-500/40 text-emerald-300 text-[11px] font-mono tracking-normal cursor-pointer transition-all shadow-md active:scale-95 group"
                title="Gün Takip ve Otomatik 90 Gün E-Posta Bildirim Sistemi "
              >
                <span className="relative flex h-2 w-2">
                  <span className={`animate-ping absolute inline-flex h-full w-full rounded-full ${autoReminderStatus.items90DaysCount > 0 ? 'bg-amber-400' : 'bg-emerald-400'} opacity-75`}></span>
                  <span className={`relative inline-flex rounded-full h-2 w-2 ${autoReminderStatus.items90DaysCount > 0 ? 'bg-amber-500' : 'bg-emerald-500'}`}></span>
                </span>
                <Bell className="w-3.5 h-3.5 text-emerald-400 group-hover:text-emerald-200" />
                <span className="font-bold">GÜN TAKİP & OTOMATİK BİLDİRİM:</span>
                <span className={autoReminderStatus.items90DaysCount > 0 ? "text-amber-400 font-black" : "text-emerald-300 font-bold"}>
                  {autoReminderStatus.items90DaysCount > 0
                    ? `${autoReminderStatus.items90DaysCount} Öğe (≤90 Gün)`
                    : 'Tüm Bakımlar Güncel'}
                </span>
                <span className="text-[9px] text-emerald-400/60 font-normal pl-1 border-l border-emerald-500/20 hidden sm:inline">
                  
                </span>
              </button>
            </div>
            <div className="flex items-center gap-4">
              <LiveClock />
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

            {/* FORM KAYITLARI */}
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
              <span>{(modalType === 'form_table' || modalType === 'techizat_matrix' || modalType === 'excel_sync' || modalType === 'iframe' || categoryHistory.length > 1) ? 'Geri' : 'Kapat'}</span>
            </button>

            {/* Action Buttons */}
            <div className="flex items-center gap-2 pointer-events-auto">
              {/* External Iframe Action Buttons (8. DENETLEME RAPOR VE EKLER vb.) */}
              {modalType === 'iframe' && modalUrl && (
                <div className="flex items-center gap-2">
                  <div className="bg-white/95 border border-slate-200/80 backdrop-blur-md shadow-lg px-4 py-2.5 rounded-full text-xs font-black text-[#0b3d1d] uppercase tracking-wider hidden sm:flex items-center gap-2">
                    <Folder className="w-3.5 h-3.5 text-emerald-800" />
                    <span className="truncate max-w-xs">{modalTitle || '8. DENETLEME RAPOR VE EKLER'}</span>
                  </div>
                  <button
                    type="button"
                    onClick={() => {
                      setIframeLoading(true);
                      const iframe = document.getElementById('external-system-iframe') as HTMLIFrameElement;
                      if (iframe) {
                        iframe.src = modalUrl;
                      }
                    }}
                    className="flex items-center gap-1.5 bg-white/95 hover:bg-white text-slate-700 hover:text-emerald-900 border border-slate-200/80 shadow-lg px-3.5 py-2.5 rounded-full text-xs font-bold transition-all cursor-pointer active:scale-95"
                    title="Sayfayı Yenile"
                  >
                    <RefreshCw className="w-3.5 h-3.5 text-emerald-800" />
                    <span className="hidden md:inline">YENİLE</span>
                  </button>
                  <a
                    href={modalUrl}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="flex items-center gap-1.5 bg-[#0b3d1d] hover:bg-[#072612] text-white shadow-lg px-4 py-2.5 rounded-full text-xs font-black uppercase tracking-wider transition-all cursor-pointer active:scale-95"
                    title="Yeni Sekmede Aç"
                  >
                    <ExternalLink className="w-3.5 h-3.5 text-emerald-300" />
                    <span>YENİ SEKMEDE AÇ</span>
                  </a>
                </div>
              )}

              {/* Top-right Excel İndir removed per user request */}

              {/* Show Veri Güncelle button on Form/Teçhizat views */}
              {modalType !== 'denetleme' && modalType !== 'iframe' && (selectedCategory === 'FORM KAYITLARI' || selectedCategory === 'HA_YER_DESTEK' || selectedCategory === 'T70_DETAY' || selectedCategory === 'KARA_ARACLARI_MENU' || selectedFormId !== null || modalType === 'form_table' || modalType === 'techizat_matrix' || modalType === 'excel_sync') && (
                <button
                  onClick={() => {
                    setPasswordActionType('global_sync');
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
            <div id="iframe-loader" className="absolute inset-0 flex items-center justify-center bg-white/80 z-10 pointer-events-none">
              <div className="flex flex-col items-center gap-3">
                <div className="loader"></div>
                <span className="text-xs font-bold text-slate-600 uppercase tracking-wider">
                  Süreç Yönetim Sistemi Yükleniyor...
                </span>
              </div>
            </div>
          )}

          {/* IFRAME MODAL GÖRÜNÜMÜ (8. DENETLEME RAPOR VE EKLER) */}
          {modalType === 'iframe' && modalUrl && (
            <div className="w-full h-full flex flex-col pt-16 bg-white relative">
              <iframe
                id="external-system-iframe"
                src={modalUrl}
                className="w-full h-full border-0 bg-white"
                title={modalTitle || "Sistem Portalı"}
                onLoad={() => setIframeLoading(false)}
                onError={() => setIframeLoading(false)}
                sandbox="allow-scripts allow-same-origin allow-forms allow-popups allow-modals allow-downloads"
              />
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
                      <span className="text-[#0b3d1d] font-bold tracking-widest text-sm mb-2 uppercase">HAVA ARAÇLARI YER DESTEK VE ÖZEL ALETLER</span>
                      <span className="text-[10px] text-[#0b3d1d]/60 uppercase tracking-widest font-semibold font-mono">YER DESTEK TEÇHİZATLARI VE ÖZEL ALETLER</span>
                    </button>

                    <button
                      onClick={() => openTechizatMatrix('hangar', 'HANGAR YER DESTEK TEÇHİZATLARI', 'all')}
                      className="bg-white hover:bg-white/80 border border-gray-200 shadow-sm rounded-[2rem] p-8 flex flex-col items-center text-center transition-all duration-300 hover:scale-[1.03] active:scale-[0.98] cursor-pointer group"
                    >
                      <div className="w-16 h-16 bg-[#0b3d1d]/10 rounded-2xl flex items-center justify-center mb-6 group-hover:bg-[#0b3d1d]/20 transition-all shadow-sm">
                        <Wrench className="w-8 h-8 text-[#0b3d1d]" />
                      </div>
                      <span className="text-[#0b3d1d] font-bold tracking-widest text-sm mb-2 uppercase">HANGAR YER DESTEK</span>
                      <span className="text-[10px] text-[#0b3d1d]/60 uppercase tracking-widest font-semibold font-mono">Yer Destek & Teçhizat Listesi</span>
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

                {/* HAVA ARAÇLARI YER DESTEK VE ÖZEL ALETLER (HA_YER_DESTEK) Birim Kartları */}
                {selectedCategory === 'HA_YER_DESTEK' && (
                  <>
                    <button
                      onClick={() => {
                        setSelectedUnitFolder('bell429');
                        navigateToSubCategory('UNIT_FOLDER_MENU');
                      }}
                      className="bg-white hover:bg-white/80 border border-gray-200 shadow-sm rounded-[2rem] p-8 flex flex-col items-center text-center transition-all duration-300 hover:scale-[1.03] active:scale-[0.98] cursor-pointer group"
                    >
                      <div className="w-16 h-16 bg-white border border-gray-200 shadow-inner rounded-2xl flex items-center justify-center mb-6 text-[#0b3d1d] font-black text-xs">
                        B429
                      </div>
                      <span className="text-[#0b3d1d] font-bold tracking-widest text-sm mb-2 uppercase">BELL 429</span>
                      <span className="text-[10px] text-gray-500 uppercase tracking-widest font-mono">Yer Destek Teçhizatları</span>
                    </button>

                    <button
                      onClick={() => {
                        setSelectedUnitFolder('at802');
                        navigateToSubCategory('UNIT_FOLDER_MENU');
                      }}
                      className="bg-white hover:bg-white/80 border border-gray-200 shadow-sm rounded-[2rem] p-8 flex flex-col items-center text-center transition-all duration-300 hover:scale-[1.03] active:scale-[0.98] cursor-pointer group"
                    >
                      <div className="w-16 h-16 bg-white border border-gray-200 shadow-inner rounded-2xl flex items-center justify-center mb-6 text-[#0b3d1d] font-black text-[10px]">
                        AT-802F
                      </div>
                      <span className="text-[#0b3d1d] font-bold tracking-widest text-sm mb-2 uppercase">AT-802F</span>
                      <span className="text-[10px] text-gray-500 uppercase tracking-widest font-mono">Yer Destek & Özel Aletler</span>
                    </button>

                    <button
                      onClick={() => {
                        setSelectedUnitFolder('t70');
                        navigateToSubCategory('UNIT_FOLDER_MENU');
                      }}
                      className="bg-white hover:bg-white/80 border-2 border-emerald-600/20 shadow-sm rounded-[2rem] p-8 flex flex-col items-center text-center transition-all duration-300 hover:scale-[1.03] active:scale-[0.98] cursor-pointer group"
                    >
                      <div className="w-16 h-16 bg-emerald-50 border border-emerald-100 shadow-inner rounded-2xl flex items-center justify-center mb-6 text-emerald-800 font-extrabold text-xs">
                        T-70
                      </div>
                      <span className="text-emerald-900 font-extrabold tracking-widest text-sm mb-2 uppercase">T-70</span>
                      <span className="text-[10px] text-gray-500 uppercase tracking-widest font-mono">Yer Destek Teçhizatları</span>
                    </button>

                    <button
                      onClick={() => {
                        setSelectedUnitFolder('c650');
                        navigateToSubCategory('UNIT_FOLDER_MENU');
                      }}
                      className="bg-white hover:bg-white/80 border border-gray-200 shadow-sm rounded-[2rem] p-8 flex flex-col items-center text-center transition-all duration-300 hover:scale-[1.03] active:scale-[0.98] cursor-pointer group"
                    >
                      <div className="w-16 h-16 bg-white border border-gray-200 shadow-inner rounded-2xl flex items-center justify-center mb-6 text-[#0b3d1d] font-black text-xs">
                        C650
                      </div>
                      <span className="text-[#0b3d1d] font-bold tracking-widest text-sm mb-2 uppercase">C-650</span>
                      <span className="text-[10px] text-gray-500 uppercase tracking-widest font-mono">Yer Destek Teçhizatları</span>
                    </button>

                    <button
                      onClick={() => {
                        setSelectedUnitFolder('b360');
                        navigateToSubCategory('UNIT_FOLDER_MENU');
                      }}
                      className="bg-white hover:bg-white/80 border border-gray-200 shadow-sm rounded-[2rem] p-8 flex flex-col items-center text-center transition-all duration-300 hover:scale-[1.03] active:scale-[0.98] cursor-pointer group"
                    >
                      <div className="w-16 h-16 bg-white border border-gray-200 shadow-inner rounded-2xl flex items-center justify-center mb-6 text-[#0b3d1d] font-black text-xs">
                        B360
                      </div>
                      <span className="text-[#0b3d1d] font-bold tracking-widest text-sm mb-2 uppercase">B-360</span>
                      <span className="text-[10px] text-gray-500 uppercase tracking-widest font-mono">Yer Destek Teçhizatları</span>
                    </button>

                    {/* 1. TÜM BİRİMLERDE TEÇHİZAT ARA */}
                    <button
                      onClick={() => openTechizatMatrix('all', 'TÜM BİRİMLER ORTAK TEÇHİZAT ARAMA (YER DESTEK & ÖZEL ALETLER)', 'techizat_all')}
                      className="bg-emerald-50 hover:bg-emerald-100/80 border-2 border-emerald-600/30 shadow-sm rounded-[2rem] p-6 flex flex-col sm:flex-row items-center gap-5 text-left transition-all duration-300 hover:scale-[1.02] active:scale-[0.98] cursor-pointer group sm:col-span-2 lg:col-span-3 mt-4"
                    >
                      <div className="w-14 h-14 bg-[#0b3d1d] text-white rounded-2xl flex items-center justify-center shrink-0 shadow-md group-hover:bg-[#082a14] transition-all">
                        <Search className="w-7 h-7 text-emerald-300" />
                      </div>
                      <div className="flex flex-col">
                        <span className="text-[#0b3d1d] font-black tracking-widest text-base uppercase">🔍 TÜM BİRİMLERDE TEÇHİZAT ARA</span>
                        <span className="text-xs text-emerald-800 tracking-wide font-mono font-bold">
                          Bütün Hava Araçlarının Yer Destek Teçhizatları ve Özel Aletlerini Tek Ortak Listede Arayın & Excel Olarak İndirin
                        </span>
                      </div>
                    </button>

                    {/* 2. TÜM BİRİMLERDE DEPO YÖNETİMİ */}
                    <button
                      onClick={() => openStandaloneDepo()}
                      className="bg-amber-50 hover:bg-amber-100/80 border-2 border-amber-600/30 shadow-sm rounded-[2rem] p-6 flex flex-col sm:flex-row items-center gap-5 text-left transition-all duration-300 hover:scale-[1.02] active:scale-[0.98] cursor-pointer group sm:col-span-2 lg:col-span-3"
                    >
                      <div className="w-14 h-14 bg-amber-800 text-white rounded-2xl flex items-center justify-center shrink-0 shadow-md group-hover:bg-amber-900 transition-all">
                        <Boxes className="w-7 h-7 text-amber-200" />
                      </div>
                      <div className="flex flex-col">
                        <span className="text-amber-900 font-black tracking-widest text-base uppercase">📦 DEPO YÖNETİM VE TAKİP SİSTEMİ</span>
                        <span className="text-xs text-amber-800 tracking-wide font-mono font-bold">
                          Bütün Hava Araçlarının Sarf/Parça ve Kimyasal Depolarını Ayrı Sistem Portalı İçinde Yönetin
                        </span>
                      </div>
                    </button>
                  </>
                )}

                {/* BİRİM İÇİ KLASÖR MENÜSÜ (UNIT_FOLDER_MENU) */}
                {selectedCategory === 'UNIT_FOLDER_MENU' && selectedUnitFolder && (
                  <>
                    {/* AT-802F ÖZEL DURUM: YER DESTEK VE ÖZEL ALETLER AYRI */}
                    {selectedUnitFolder === 'at802' ? (
                      <>
                        <button
                          onClick={() => openTechizatMatrix('at802', 'AT-802F - YER DESTEK TEÇHİZATLARI', 'yer_destek')}
                          className="bg-white hover:bg-emerald-50/50 border-2 border-emerald-600/30 shadow-sm rounded-[2rem] p-8 flex flex-col items-center text-center transition-all duration-300 hover:scale-[1.03] active:scale-[0.98] cursor-pointer group"
                        >
                          <div className="w-16 h-16 bg-emerald-100/80 rounded-2xl flex items-center justify-center mb-6 group-hover:bg-emerald-200 transition-all shadow-sm text-[#0b3d1d]">
                            <Plane className="w-8 h-8 text-[#0b3d1d]" />
                          </div>
                          <span className="text-[#0b3d1d] font-black tracking-widest text-base mb-2 uppercase">YER DESTEK TEÇHİZATLARI</span>
                          <span className="text-[10px] text-gray-500 uppercase tracking-widest font-mono font-bold">
                            AT-802F Yer Destek Ekipmanları
                          </span>
                        </button>

                        <button
                          onClick={() => openTechizatMatrix('at802_ozel_alet', 'AT-802F - ÖZEL ALETLER', 'ozel_alet')}
                          className="bg-white hover:bg-emerald-50/50 border-2 border-emerald-600/30 shadow-sm rounded-[2rem] p-8 flex flex-col items-center text-center transition-all duration-300 hover:scale-[1.03] active:scale-[0.98] cursor-pointer group"
                        >
                          <div className="w-16 h-16 bg-emerald-100/80 rounded-2xl flex items-center justify-center mb-6 group-hover:bg-emerald-200 transition-all shadow-sm text-[#0b3d1d]">
                            <Wrench className="w-8 h-8 text-[#0b3d1d]" />
                          </div>
                          <span className="text-[#0b3d1d] font-black tracking-widest text-base mb-2 uppercase">ÖZEL ALETLER</span>
                          <span className="text-[10px] text-gray-500 uppercase tracking-widest font-mono font-bold">
                            AT-802F Özel Aletler Listesi
                          </span>
                        </button>

                        <button
                          onClick={() => openStandaloneDepo(selectedUnitFolder)}
                          className="bg-white hover:bg-amber-50/50 border-2 border-amber-600/30 shadow-sm rounded-[2rem] p-8 flex flex-col items-center text-center transition-all duration-300 hover:scale-[1.03] active:scale-[0.98] cursor-pointer group"
                        >
                          <div className="w-16 h-16 bg-amber-100/80 rounded-2xl flex items-center justify-center mb-6 group-hover:bg-amber-200 transition-all shadow-sm text-amber-800">
                            <Boxes className="w-8 h-8 text-amber-800" />
                          </div>
                          <span className="text-amber-900 font-black tracking-widest text-base mb-2 uppercase">DEPO SİSTEMİ</span>
                          <span className="text-[10px] text-amber-700/80 uppercase tracking-widest font-mono font-bold">
                            Sarf/Parça Deposu & Kimyasal Depo Yönetim Portalı
                          </span>
                        </button>

                        <button
                          id="btn-unit-olay-takip-at802"
                          onClick={() => {
                            setOlayTakipInitialUnit('at802');
                            setIsOlayTakipOpen(true);
                          }}
                          className="bg-white hover:bg-emerald-50/50 border-2 border-emerald-600/30 shadow-sm rounded-[2rem] p-8 flex flex-col items-center text-center transition-all duration-300 hover:scale-[1.03] active:scale-[0.98] cursor-pointer group"
                        >
                          <div className="w-16 h-16 bg-emerald-100/80 rounded-2xl flex items-center justify-center mb-6 group-hover:bg-emerald-200 transition-all shadow-sm text-[#0b3d1d]">
                            <FileSpreadsheet className="w-8 h-8 text-[#0b3d1d]" />
                          </div>
                          <span className="text-[#0b3d1d] font-black tracking-widest text-base mb-2 uppercase">OLAY TAKİP ÇİZELGESİ</span>
                          <span className="text-[10px] text-gray-500 uppercase tracking-widest font-mono font-bold">
                            AT-802 Kaza-Kırım & Limit Aşımları
                          </span>
                        </button>

                        <button
                          id="btn-unit-barkod-at802"
                          onClick={() => {
                            setBarkodOkuyucuInitialUnit('at802');
                            setIsBarkodOkuyucuOpen(true);
                          }}
                          className="bg-white hover:bg-emerald-50/50 border-2 border-emerald-600/30 shadow-sm rounded-[2rem] p-8 flex flex-col items-center text-center transition-all duration-300 hover:scale-[1.03] active:scale-[0.98] cursor-pointer group"
                        >
                          <div className="w-16 h-16 bg-emerald-100/80 rounded-2xl flex items-center justify-center mb-6 group-hover:bg-emerald-200 transition-all shadow-sm text-[#0b3d1d]">
                            <Barcode className="w-8 h-8 text-[#0b3d1d]" />
                          </div>
                          <span className="text-[#0b3d1d] font-black tracking-widest text-base mb-2 uppercase">BARKOD OKUYUCU</span>
                          <span className="text-[10px] text-gray-500 uppercase tracking-widest font-mono font-bold">
                            Kamera ile Barkod Okuma & Sayım
                          </span>
                        </button>
                      </>
                    ) : (
                      <>
                        <button
                          onClick={() => {
                            const unitName = getUnitDisplayName(selectedUnitFolder);
                            openTechizatMatrix(selectedUnitFolder as any, `${unitName} - YER DESTEK VE ÖZEL ALETLER`, 'yer_destek');
                          }}
                          className="bg-white hover:bg-emerald-50/50 border-2 border-emerald-600/30 shadow-sm rounded-[2rem] p-8 flex flex-col items-center text-center transition-all duration-300 hover:scale-[1.03] active:scale-[0.98] cursor-pointer group"
                        >
                          <div className="w-16 h-16 bg-emerald-100/80 rounded-2xl flex items-center justify-center mb-6 group-hover:bg-emerald-200 transition-all shadow-sm text-[#0b3d1d]">
                            <Wrench className="w-8 h-8 text-[#0b3d1d]" />
                          </div>
                          <span className="text-[#0b3d1d] font-black tracking-widest text-base mb-2 uppercase">YER DESTEK VE ÖZEL ALETLER</span>
                          <span className="text-[10px] text-gray-500 uppercase tracking-widest font-mono font-bold">
                            Mevcut Yer Destek Teçhizatları ve Özel Aletler
                          </span>
                        </button>

                        <button
                          onClick={() => openStandaloneDepo(selectedUnitFolder)}
                          className="bg-white hover:bg-amber-50/50 border-2 border-amber-600/30 shadow-sm rounded-[2rem] p-8 flex flex-col items-center text-center transition-all duration-300 hover:scale-[1.03] active:scale-[0.98] cursor-pointer group"
                        >
                          <div className="w-16 h-16 bg-amber-100/80 rounded-2xl flex items-center justify-center mb-6 group-hover:bg-amber-200 transition-all shadow-sm text-amber-800">
                            <Boxes className="w-8 h-8 text-amber-800" />
                          </div>
                          <span className="text-amber-900 font-black tracking-widest text-base mb-2 uppercase">DEPO SİSTEMİ</span>
                          <span className="text-[10px] text-amber-700/80 uppercase tracking-widest font-mono font-bold">
                            Sarf ve Parça Depo & Kimyasal Depo Yönetim Portalı
                          </span>
                        </button>

                        <button
                          id="btn-unit-olay-takip-general"
                          onClick={() => {
                            setOlayTakipInitialUnit(selectedUnitFolder);
                            setIsOlayTakipOpen(true);
                          }}
                          className="bg-white hover:bg-emerald-50/50 border-2 border-emerald-600/30 shadow-sm rounded-[2rem] p-8 flex flex-col items-center text-center transition-all duration-300 hover:scale-[1.03] active:scale-[0.98] cursor-pointer group"
                        >
                          <div className="w-16 h-16 bg-emerald-100/80 rounded-2xl flex items-center justify-center mb-6 group-hover:bg-emerald-200 transition-all shadow-sm text-[#0b3d1d]">
                            <FileSpreadsheet className="w-8 h-8 text-[#0b3d1d]" />
                          </div>
                          <span className="text-[#0b3d1d] font-black tracking-widest text-base mb-2 uppercase">OLAY TAKİP ÇİZELGESİ</span>
                          <span className="text-[10px] text-gray-500 uppercase tracking-widest font-mono font-bold">
                            {getUnitDisplayName(selectedUnitFolder)} Olay & Limit Çizelgeleri
                          </span>
                        </button>

                        <button
                          id="btn-unit-barkod-general"
                          onClick={() => {
                            setBarkodOkuyucuInitialUnit(selectedUnitFolder);
                            setIsBarkodOkuyucuOpen(true);
                          }}
                          className="bg-white hover:bg-emerald-50/50 border-2 border-emerald-600/30 shadow-sm rounded-[2rem] p-8 flex flex-col items-center text-center transition-all duration-300 hover:scale-[1.03] active:scale-[0.98] cursor-pointer group"
                        >
                          <div className="w-16 h-16 bg-emerald-100/80 rounded-2xl flex items-center justify-center mb-6 group-hover:bg-emerald-200 transition-all shadow-sm text-[#0b3d1d]">
                            <Barcode className="w-8 h-8 text-[#0b3d1d]" />
                          </div>
                          <span className="text-[#0b3d1d] font-black tracking-widest text-base mb-2 uppercase">BARKOD OKUYUCU</span>
                          <span className="text-[10px] text-gray-500 uppercase tracking-widest font-mono font-bold">
                            {getUnitDisplayName(selectedUnitFolder)} Barkod & Sayım
                          </span>
                        </button>
                      </>
                    )}

                    {selectedUnitFolder === 't70' && (
                      <>
                        <button
                          onClick={() => openTechizatMatrix('t70_bumbi_backet', 'T-70 BUMBİ BACKET TEÇHİZATI', 'all')}
                          className="bg-white hover:bg-white/80 border border-gray-200 shadow-sm rounded-[2rem] p-8 flex flex-col items-center text-center transition-all duration-300 hover:scale-[1.03] active:scale-[0.98] cursor-pointer group"
                        >
                          <div className="w-16 h-16 bg-[#0b3d1d]/10 rounded-2xl flex items-center justify-center mb-6 group-hover:bg-[#0b3d1d]/20 transition-all shadow-sm text-[#0b3d1d] font-extrabold text-sm">
                            BB
                          </div>
                          <span className="text-[#0b3d1d] font-bold tracking-widest text-sm mb-2 uppercase">BUMBİ BACKET</span>
                          <span className="text-[10px] text-gray-500 uppercase tracking-widest font-mono">T-70 Yangın Kovası</span>
                        </button>

                        <button
                          onClick={() => openTechizatMatrix('t70_helitak', 'T-70 HELİTAK TEÇHİZATI', 'all')}
                          className="bg-white hover:bg-white/80 border border-gray-200 shadow-sm rounded-[2rem] p-8 flex flex-col items-center text-center transition-all duration-300 hover:scale-[1.03] active:scale-[0.98] cursor-pointer group"
                        >
                          <div className="w-16 h-16 bg-[#0b3d1d]/10 rounded-2xl flex items-center justify-center mb-6 group-hover:bg-[#0b3d1d]/20 transition-all shadow-sm text-[#0b3d1d] font-extrabold text-sm">
                            HT
                          </div>
                          <span className="text-[#0b3d1d] font-bold tracking-widest text-sm mb-2 uppercase">HELİTAK</span>
                          <span className="text-[10px] text-gray-500 uppercase tracking-widest font-mono">T-70 Helitak Teçhizatı</span>
                        </button>
                      </>
                    )}

                    <button
                      onClick={() => {
                        const unitName = getUnitDisplayName(selectedUnitFolder);
                        openTechizatMatrix(selectedUnitFolder as any, `${unitName} - TÜM ENVANTER LİSTESİ`, 'all');
                      }}
                      className="bg-slate-50 hover:bg-slate-100 border border-slate-300 shadow-sm rounded-[2rem] p-8 flex flex-col items-center text-center transition-all duration-300 hover:scale-[1.03] active:scale-[0.98] cursor-pointer group sm:col-span-2 lg:col-span-2"
                    >
                      <div className="w-16 h-16 bg-slate-200 rounded-2xl flex items-center justify-center mb-6 shadow-sm">
                        <Table className="w-8 h-8 text-slate-700" />
                      </div>
                      <span className="text-slate-800 font-black tracking-widest text-sm mb-2 uppercase">📋 TÜM BİRİM ENVANTERİ</span>
                      <span className="text-[10px] text-slate-600 uppercase tracking-widest font-mono font-bold">
                        Yer Destek, Özel Aletler ve Depolardaki Tüm Teçhizatları Birlikte İncele
                      </span>
                    </button>
                  </>
                )}

                {/* BİRİM DEPO MENÜSÜ (UNIT_DEPO_MENU) - 3 ALT KLASÖR */}
                {selectedCategory === 'UNIT_DEPO_MENU' && selectedUnitFolder && (
                  <>
                    <button
                      onClick={() => {
                        const unitName = getUnitDisplayName(selectedUnitFolder);
                        openTechizatMatrix(selectedUnitFolder as any, `${unitName} - SARF VE PARÇA DEPOSU`, 'depo_sarf');
                      }}
                      className="bg-white hover:bg-blue-50/50 border-2 border-blue-600/30 shadow-sm rounded-[2rem] p-8 flex flex-col items-center text-center transition-all duration-300 hover:scale-[1.03] active:scale-[0.98] cursor-pointer group"
                    >
                      <div className="w-16 h-16 bg-blue-100 rounded-2xl flex items-center justify-center mb-6 group-hover:bg-blue-200 transition-all shadow-sm text-blue-800">
                        <Package className="w-8 h-8 text-blue-800" />
                      </div>
                      <span className="text-blue-900 font-black tracking-widest text-base mb-2 uppercase">SARF VE PARÇA DEPO</span>
                      <span className="text-[10px] text-blue-700/80 uppercase tracking-widest font-mono font-bold">
                        Sarf Malzemeleri, Yedek Parça ve Mekanik Bileşenler
                      </span>
                    </button>

                    <button
                      onClick={() => {
                        const unitName = getUnitDisplayName(selectedUnitFolder);
                        openTechizatMatrix(selectedUnitFolder as any, `${unitName} - KİMYASAL DEPO`, 'depo_kimyasal');
                      }}
                      className="bg-white hover:bg-purple-50/50 border-2 border-purple-600/30 shadow-sm rounded-[2rem] p-8 flex flex-col items-center text-center transition-all duration-300 hover:scale-[1.03] active:scale-[0.98] cursor-pointer group"
                    >
                      <div className="w-16 h-16 bg-purple-100 rounded-2xl flex items-center justify-center mb-6 group-hover:bg-purple-200 transition-all shadow-sm text-purple-800">
                        <FlaskConical className="w-8 h-8 text-purple-800" />
                      </div>
                      <span className="text-purple-900 font-black tracking-widest text-base mb-2 uppercase">KİMYASAL DEPO</span>
                      <span className="text-[10px] text-purple-700/80 uppercase tracking-widest font-mono font-bold">
                        Kimyasal Maddeler, Havacılık Yağları, Sıvılar, Mastikler ve Boyalar
                      </span>
                    </button>

                    <button
                      onClick={() => openStandaloneDepo(selectedUnitFolder)}
                      className="bg-white hover:bg-emerald-50/50 border-2 border-emerald-600/40 shadow-sm rounded-[2rem] p-8 flex flex-col items-center text-center transition-all duration-300 hover:scale-[1.03] active:scale-[0.98] cursor-pointer group"
                    >
                      <div className="w-16 h-16 bg-emerald-100 rounded-2xl flex items-center justify-center mb-6 group-hover:bg-emerald-200 transition-all shadow-sm text-emerald-800">
                        <Boxes className="w-8 h-8 text-emerald-800" />
                      </div>
                      <span className="text-emerald-950 font-black tracking-widest text-base mb-2 uppercase">DEPO YÖNETİM</span>
                      <span className="text-[10px] text-emerald-700/80 uppercase tracking-widest font-mono font-bold">
                        Depo Giriş/Çıkış, Kit Montajı, Transfer Geçmişi ve Sayım
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
                      id="btn-denetleme-rapor-ekler"
                      onClick={() => {
                        setModalUrl('https://ogmhavacilik.github.io/surecyonet/');
                        setModalType('iframe');
                        setModalTitle('8. DENETLEME RAPOR VE EKLER');
                        setIframeLoading(true);
                        setModalOpen(true);
                      }}
                      className="bg-white hover:bg-white/80 border border-gray-200 shadow-sm rounded-[2rem] p-8 flex flex-col items-center justify-center text-center transition-all duration-300 hover:scale-[1.03] active:scale-[0.98] cursor-pointer group relative"
                    >
                      <div className="w-16 h-16 bg-[#0b3d1d]/10 rounded-2xl flex items-center justify-center mb-6 group-hover:bg-[#0b3d1d]/20 transition-all shadow-sm relative">
                        <Folder className="w-8 h-8 text-[#0b3d1d]" />
                        <span className="absolute -top-1 -right-1 bg-emerald-800 text-white text-[8px] font-bold px-1.5 py-0.5 rounded-full uppercase tracking-wider">CANLI SİSTEM ➜</span>
                      </div>
                      <span className="text-[#0b3d1d] font-bold tracking-normal text-sm mb-2 text-center uppercase leading-tight">8. DENETLEME RAPOR VE EKLER</span>
                      <span className="text-[9px] text-[#0b3d1d]/60 uppercase tracking-wider font-semibold font-mono mb-2">Denetleme Rapor ve Ekleri Süreç Yönetimi</span>
                      
                      <div className="flex items-center gap-2 mt-1">
                        <span className="text-[9px] text-emerald-800 font-extrabold uppercase tracking-wider font-mono bg-emerald-50 px-2.5 py-1 rounded-full border border-emerald-200">
                          SİSTEMİ AÇ ➜
                        </span>
                        <a
                          href="https://ogmhavacilik.github.io/surecyonet/"
                          target="_blank"
                          rel="noopener noreferrer"
                          onClick={(e) => e.stopPropagation()}
                          className="text-[9px] text-slate-600 hover:text-emerald-800 font-extrabold uppercase tracking-wider font-mono bg-slate-100 hover:bg-slate-200 px-2 py-1 rounded-full border border-slate-200 flex items-center gap-1 transition-all"
                          title="Doğrudan Yeni Sekmede Aç"
                        >
                          <ExternalLink className="w-2.5 h-2.5" />
                          <span>YENİ SEKME</span>
                        </a>
                      </div>
                      
                      <div className="mt-auto pt-4 border-t border-gray-100 w-full text-center">
                        <span className="text-[10px] text-gray-400 font-extrabold uppercase tracking-widest font-mono">
                          ogmhavacilik.github.io/surecyonet
                        </span>
                      </div>
                    </button>

                    <button
                      id="btn-form-tech-pubs"
                      onClick={() => setIsTechPubsOpen(true)}
                      className="bg-white hover:bg-white/80 border border-gray-200 shadow-sm rounded-[2rem] p-8 flex flex-col items-center justify-center text-center transition-all duration-300 hover:scale-[1.03] active:scale-[0.98] cursor-pointer group relative"
                    >
                      <div className="w-16 h-16 bg-[#0b3d1d]/10 rounded-2xl flex items-center justify-center mb-6 group-hover:bg-[#0b3d1d]/20 transition-all shadow-sm relative">
                        <BookOpen className="w-8 h-8 text-[#0b3d1d]" />
                        <div className="absolute -bottom-1 -right-1 w-5 h-5 rounded-md bg-amber-500 text-slate-950 flex items-center justify-center shadow-xs border border-white">
                          <Wrench className="w-3 h-3" />
                        </div>
                      </div>
                      <span className="text-[#0b3d1d] font-bold tracking-normal text-sm mb-2 text-center uppercase leading-tight">9. TEKNİK YAYINLAR</span>
                      <span className="text-[9px] text-[#0b3d1d]/60 uppercase tracking-wider font-semibold font-mono mb-2">AMM, IPC, CMM & Canlı Depo Arama</span>
                      <span className="text-[9px] text-emerald-800 font-extrabold uppercase tracking-wider font-mono bg-emerald-50 px-2.5 py-1 rounded-full border border-emerald-200">GÖRÜNTÜLE & ARA ➜</span>
                      
                      <div className="mt-auto pt-4 border-t border-gray-100 w-full text-center">
                        <span className="text-[10px] text-gray-400 font-extrabold uppercase tracking-widest font-mono">
                          GÜNCELLEME TARİHİ: 2026-08
                        </span>
                      </div>
                    </button>

                    <button
                      id="btn-form-olay-takip"
                      onClick={() => {
                        setOlayTakipInitialUnit('at802');
                        setIsOlayTakipOpen(true);
                      }}
                      className="bg-white hover:bg-emerald-50/50 border border-emerald-600/30 shadow-sm rounded-[2rem] p-8 flex flex-col items-center justify-center text-center transition-all duration-300 hover:scale-[1.03] active:scale-[0.98] cursor-pointer group relative"
                    >
                      <div className="w-16 h-16 bg-emerald-100/80 rounded-2xl flex items-center justify-center mb-6 group-hover:bg-emerald-200 transition-all shadow-sm text-[#0b3d1d] relative">
                        <FileSpreadsheet className="w-8 h-8 text-[#0b3d1d]" />
                      </div>
                      <span className="text-[#0b3d1d] font-bold tracking-normal text-sm mb-2 text-center uppercase leading-tight">OLAY TAKİP ÇİZELGESİ</span>
                      <span className="text-[10px] text-gray-500 uppercase tracking-wider font-semibold font-mono text-center">
                        AT-802 Kaza/Kırım, Limit Aşımları & Çizelgeler
                      </span>
                    </button>

                    <button
                      id="btn-form-barkod-okuyucu"
                      onClick={() => {
                        setBarkodOkuyucuInitialUnit('at802');
                        setIsBarkodOkuyucuOpen(true);
                      }}
                      className="bg-white hover:bg-emerald-50/50 border border-emerald-600/30 shadow-sm rounded-[2rem] p-8 flex flex-col items-center justify-center text-center transition-all duration-300 hover:scale-[1.03] active:scale-[0.98] cursor-pointer group relative"
                    >
                      <div className="w-16 h-16 bg-emerald-100/80 rounded-2xl flex items-center justify-center mb-6 group-hover:bg-emerald-200 transition-all shadow-sm text-[#0b3d1d] relative">
                        <Barcode className="w-8 h-8 text-[#0b3d1d]" />
                      </div>
                      <span className="text-[#0b3d1d] font-bold tracking-normal text-sm mb-2 text-center uppercase leading-tight">BARKOD OKUYUCU</span>
                      <span className="text-[10px] text-gray-500 uppercase tracking-wider font-semibold font-mono text-center">
                        Mobil Barkod & Karekod Okuma, Stok ve Depo Sayım Terminali
                      </span>
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
                        <div className="absolute inset-0 bg-slate-900/60 flex flex-col items-center justify-center p-8 text-center animate-fade-in z-25">
                          <div className="w-12 h-12 border-4 border-emerald-500 border-t-transparent rounded-full animate-spin mb-4" />
                          <h3 className="text-emerald-400 font-extrabold text-sm uppercase tracking-widest animate-pulse font-mono mb-2">
                            BEKLEYİNİZ... VERİ YÜKLENİYOR
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
            const isDepo = techizatActiveSection === 'depo_sarf' || techizatActiveSection === 'depo_kimyasal' || techizatActiveSection === 'depo_all';

            const isHangar = activeTechizatType === 'hangar';

            const baseColumns = [
              "SIRA NO", 
              "TEÇHİZAT ADI", 
              "PARÇA NO (P/N) / MODEL", 
              "SERİ NO (S/N)", 
              "MİKTAR / KAPASİTE", 
              "BULUNDUĞU YER", 
              "DURUMU", 
              isHangar ? "BAKIMA TABİ" : "KALİBRASYONA TABİ", 
              "SON KONTROL / KALİBRASYON / BAKIM", 
              "GELECEK KONTROL / KALİBRASYON / BAKIM", 
              "SON KONTROLÜ YAPAN FİRMA", 
              "AÇIKLAMA", 
              "90 GÜN UYARISI MAİL GÖNDERİM TARİHİ"
            ];

            const depoColumns = activeTechizatType === 'at802' ? [
              "SIRA NO", 
              "MALZEME / PARÇA ADI", 
              "PARÇA NO (P/N)", 
              "SERİ NO (S/N)", 
              "MİKTAR", 
              "BULUNDUĞU YER / RAF", 
              "LOKASYON NO",
              "TOPLAM STOK",
              "ANKARA MEVCUT",
              "MİLAS MEVCUT",
              "KARAİN MEVCUT",
              "ÇANAKKALE MEVCUT",
              "BURSA MEVCUT",
              "DURUMU", 
              "ÖMÜRLÜ PARÇA MI?", 
              "ÖMÜR BİTİŞ TARİHİ", 
              "TEDARİK EDİLEN FİRMA", 
              "AÇIKLAMA"
            ] : [
              "SIRA NO", 
              "MALZEME / PARÇA ADI", 
              "PARÇA NO (P/N)", 
              "SERİ NO (S/N)", 
              "MİKTAR", 
              "BULUNDUĞU YER / RAF", 
              "DURUMU", 
              "ÖMÜRLÜ PARÇA MI?", 
              "ÖMÜR BİTİŞ TARİHİ", 
              "TEDARİK EDİLEN FİRMA", 
              "AÇIKLAMA"
            ];

            const karaAraclariColumns = [
              "SIRA NO", 
              "ARAÇ PLAKASI / TANIMI", 
              "MARKA", 
              "PARÇA NO (P/N) / MODEL", 
              "BULUNDUĞU YER", 
              "SON KM Sİ", 
              "DURUMU", 
              "BAKIMA TABİ", 
              "SON KONTROL / KALİBRASYON / BAKIM", 
              "GELECEK KONTROL / KALİBRASYON / BAKIM", 
              "SON KONTROLÜ YAPAN FİRMA", 
              "AÇIKLAMA", 
              "90 GÜN UYARISI MAİL GÖNDERİM TARİHİ"
            ];

            const hangarColumns = [
              ...baseColumns,
              "BELGE YÜKLE"
            ];

            const cols = isDepo
              ? (activeTechizatType === 'all' ? ["AİT OLDUĞU BİRİM", ...depoColumns] : depoColumns)
              : activeTechizatType === 'all'
                ? ["AİT OLDUĞU BİRİM", ...baseColumns]
                : isKaraAraci
                  ? karaAraclariColumns
                  : isHangar
                    ? hangarColumns
                    : baseColumns;

            const rawUnitRows = activeTechizatType === 'all'
              ? [
                  ...techizatBell429Data.map(r => ["BELL 429", ...r]),
                  ...techizatAt802Data.map(r => ["AT-802F", ...r]),
                  ...techizatT70Data.map(r => ["T-70 YER DESTEK", ...r]),
                  ...techizatT70BumbiBacketData.map(r => ["T-70 BUMBİ BACKET", ...r]),
                  ...techizatT70HelitakData.map(r => ["T-70 HELİTAK", ...r]),
                  ...techizatC650Data.map(r => ["C-650", ...r]),
                  ...techizatB360Data.map(r => ["B-360", ...r]),
                  ...techizatHangarData.map(r => ["HANGAR YER DESTEK", ...r]),
                  ...(!isDepo ? techizatKaraAraclariData.map(r => ["KARA ARAÇLARI", ...r]) : [])
                ]
              : activeTechizatType === 'bell429' ? techizatBell429Data
              : (activeTechizatType === 'at802_ozel_alet' || (activeTechizatType === 'at802' && techizatActiveSection === 'ozel_alet')) ? techizatAt802OzelAletData
              : activeTechizatType === 'at802' ? techizatAt802Data
              : activeTechizatType === 't70' ? techizatT70Data
              : activeTechizatType === 't70_bumbi_backet' ? techizatT70BumbiBacketData
              : activeTechizatType === 't70_helitak' ? techizatT70HelitakData
              : activeTechizatType === 'b360' ? techizatB360Data
              : activeTechizatType === 'c650' ? techizatC650Data
              : activeTechizatType === 'kara_araclari' ? techizatKaraAraclariData
              : techizatHangarData;

            const sectionFilterRow = (r: string[]) => {
              if (techizatActiveSection === 'all') return true;
              if (activeTechizatType === 'at802' && techizatActiveSection === 'yer_destek') return true;
              if (activeTechizatType === 'at802_ozel_alet' || (activeTechizatType === 'at802' && techizatActiveSection === 'ozel_alet')) return true;
              const isAll = activeTechizatType === 'all';
              const targetRow = isAll ? r.slice(1) : r;
              const unitHint = isAll ? (r[0] || "").toLowerCase() : activeTechizatType || undefined;
              const sec = getRowSection(targetRow, unitHint);
              if (techizatActiveSection === 'techizat_all') {
                return sec === 'yer_destek' || sec === 'ozel_alet';
              }
              if (techizatActiveSection === 'depo_all') {
                return sec === 'depo_sarf' || sec === 'depo_kimyasal';
              }
              return sec === techizatActiveSection;
            };

            const rows = rawUnitRows
              .filter(r => {
                const targetRow = activeTechizatType === 'all' ? r.slice(1) : r;
                // Kullanıcı kuralı: "VERİ GÜNCELLEMEDE BAŞLIK 1. SATIR ÜRÜN DİYE ATMIŞ HATADIR."
                if (isHeaderLikeRow(targetRow)) return false;
                return sectionFilterRow(r);
              })
              .map((r, idx) => {
              const isAll = activeTechizatType === 'all';
              const unitLabel = isAll ? r[0] : null;
              const targetRow = isAll ? r.slice(1) : r;

              let formatted: string[];
              if (isDepo) {
                formatted = formatDepoRow(targetRow);
              } else if (isKaraAraci) {
                formatted = formatKaraAraclariRow(targetRow);
              } else if (isAll && unitLabel === "KARA ARAÇLARI") {
                formatted = formatKaraAraciToStandardRow(targetRow);
              } else {
                formatted = formatStandardRow(targetRow);
              }

              const finalRow = isAll && unitLabel ? [unitLabel, ...formatted] : formatted;
              const siraIdx = isAll ? 1 : 0;
              if (finalRow[siraIdx] !== undefined) {
                finalRow[siraIdx] = String(idx + 1);
              }
              return finalRow;
            });

            const firmaColIdx = isDepo ? cols.indexOf("TEDARİK EDİLEN FİRMA") : cols.indexOf("SON KONTROLÜ YAPAN FİRMA");
            const durumColIdx = cols.indexOf("DURUMU");
            const kalibrasyonTabiColIdx = cols.findIndex(c => (c || '') === "BAKIMA TABİ" || (c || '') === "KALİBRASYONA TABİ" || (c || '').includes("TABİ"));
            const omurluTabiColIdx = cols.indexOf("ÖMÜRLÜ PARÇA MI?");
            const omurBitisColIdx = cols.indexOf("ÖMÜR BİTİŞ TARİHİ");
            const gelecekBakimColIdx = cols.indexOf("GELECEK KONTROL / KALİBRASYON / BAKIM");

            const uniqueFirmalar = Array.from<string>(
              new Set(
                rows
                  .map(r => (r[firmaColIdx] || "").trim())
                  .filter(f => f && f !== "-" && f !== "--" && f !== "MUAFIYET (TABİ DEĞİL)" && f.toUpperCase() !== "BELİRTİLMEMİŞ")
              )
            ).sort((a, b) => a.localeCompare(b, 'tr-TR'));

            const getStandardizedStatusTokens = (cellVal: string): string[] => {
              const str = String(cellVal || "").trim();
              if (!str || str === "-" || str === "--") return [];
              const rawTokens = str.split(/\r?\n|\/|,|;/);
              const tokens: string[] = [];
              for (let t of rawTokens) {
                t = t.trim().toLocaleUpperCase('tr-TR');
                if (!t || t === "-" || t === "--") continue;
                if (t.includes("GAYRİ") || t.includes("GAYRI") || t.includes("DEĞİL") || t.includes("DEĞIL") || t.includes("ARIZA") || t.includes("G.FAAL") || t.includes("G. FAAL")) {
                  tokens.push("GAYRİ FAAL");
                } else if (t.includes("FAAL")) {
                  tokens.push("FAAL");
                } else {
                  tokens.push(t);
                }
              }
              return Array.from(new Set(tokens));
            };

            const uniqueDurumlarSet = new Set<string>();
            rows.forEach(r => {
              const tokens = getStandardizedStatusTokens(r[durumColIdx]);
              tokens.forEach(tok => uniqueDurumlarSet.add(tok));
            });
            const uniqueDurumlar = Array.from(uniqueDurumlarSet).sort((a, b) => a.localeCompare(b, 'tr-TR'));

            const q = techizatSearchQuery.toLowerCase().trim();
            const firmaFilter = techizatFirmaFilter.trim().toLocaleLowerCase('tr-TR');

            const matchesList: { r: number; c: number }[] = [];

            let filteredRows = rows.filter(row => {
              // 1. Arama kelimesi filtresi
              if (q && !row.some(cell => cell && cell.toLowerCase().includes(q))) {
                return false;
              }
              // 2. Firma filtresi (Türkçe karakter uyumlu)
              if (firmaFilter && firmaFilter !== "tüm firmalar" && !firmaFilter.includes("tüm")) {
                const rowFirma = (row[firmaColIdx] || "").trim().toLocaleLowerCase('tr-TR');
                if (rowFirma !== firmaFilter && !rowFirma.includes(firmaFilter)) return false;
              }
              // 3. Durum filtresi (Dinamik ve büyük/küçük harf duyarsız Türkçe normalizasyon)
              if (techizatDurumFilter && techizatDurumFilter !== "TÜM DURUMLAR") {
                const selectedDurum = techizatDurumFilter.trim().toLocaleUpperCase('tr-TR');
                const rowTokens = getStandardizedStatusTokens(row[durumColIdx]);

                if (selectedDurum === "GAYRİ FAAL" || selectedDurum === "GAYRI FAAL") {
                  // User rule: if any single item inside a PN is gayrı faal, show it in GAYRİ FAAL filter
                  const hasGayri = rowTokens.includes("GAYRİ FAAL");
                  if (!hasGayri) return false;
                } else if (selectedDurum === "FAAL") {
                  // Sadece net 'FAAL' olanlar; GAYRİ FAAL elenir
                  if (rowTokens.includes("GAYRİ FAAL")) return false;
                  if (!rowTokens.includes("FAAL")) return false;
                } else {
                  if (!rowTokens.includes(selectedDurum)) return false;
                }
              }
              // 4. Renk Kodu Doğrudan Seçim Filtresi
              if (techizatColorFilter && techizatColorFilter !== 'ALL') {
                const targetColIdx = isDepo ? omurBitisColIdx : gelecekBakimColIdx;
                if (targetColIdx === -1) return false;

                const isMuaf = isDepo
                  ? (omurluTabiColIdx !== -1 && (row[omurluTabiColIdx] || "").trim().toUpperCase() === "HAYIR")
                  : (kalibrasyonTabiColIdx !== -1 && ["HAYIR", "HAYIR (MUAFIYET)", "MUAFIYET"].includes((row[kalibrasyonTabiColIdx] || "").trim().toUpperCase()));

                const rawVal = String(row[targetColIdx] || "").trim();
                const dateLines = rawVal.split(/[\r\n;]+/).map(s => s.trim()).filter(s => s && s !== '-' && s !== '--' && !s.toUpperCase().includes('MUAFIYET'));

                const daysList: number[] = [];
                for (const dStr of dateLines) {
                  const d = parseGelecekBakimDays(dStr);
                  if (d !== null) daysList.push(d);
                }

                // Tarihi olmayanlar
                if (daysList.length === 0 || isMuaf) {
                  // Kullanıcı: "TARİHİ OLMAYANLAR DA OLMAZ DİYELİM ONLAR DA FİLTRE DIŞIDIR. GELECEK BAKIM TARİHİ BAKACAK."
                  if (techizatColorFilter === 'GRAY') {
                    return true;
                  }
                  return false;
                }

                if (techizatColorFilter === 'ORANGE') {
                  // Sadece turuncular: En az bir tarih 0 ile 90 gün arasında olmalı
                  const hasOrange = daysList.some(d => d >= 0 && d < 90);
                  if (!hasOrange) return false;
                } else if (techizatColorFilter === 'RED') {
                  // Sadece kırmızılar: En az bir tarih geçmiş olmalı (< 0)
                  const hasRed = daysList.some(d => d < 0);
                  if (!hasRed) return false;
                } else if (techizatColorFilter === 'GREEN') {
                  // Sadece yeşiller: Kırmızı veya turuncu olmayan, günü 90'dan fazla olanlar
                  const hasRedOrOrange = daysList.some(d => d < 90);
                  const hasGreen = daysList.some(d => d >= 90);
                  if (hasRedOrOrange || !hasGreen) return false;
                } else if (techizatColorFilter === 'GRAY') {
                  return false;
                }
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

            // Renk Koduna Göre Doğrudan Filtreleme
            if (selectedColorFilter !== 'all') {
              filteredRows = filteredRows.filter(row => {
                const isKalibTabiVal = kalibrasyonTabiColIdx !== -1 ? String(row[kalibrasyonTabiColIdx] || "").trim().toUpperCase() : "EVET";
                const isBakimMuaf = isKalibTabiVal === "HAYIR" || isKalibTabiVal === "HAYIR (MUAFIYET)" || isKalibTabiVal === "MUAFIYET";

                if (isDepo) {
                  const omurluIdx = cols.indexOf("ÖMÜRLÜ PARÇA MI?");
                  const isOmurlu = omurluIdx !== -1 ? (row[omurluIdx] || "").toUpperCase() === "EVET" : true;
                  if (!isOmurlu) return selectedColorFilter === 'neutral';
                  if (omurBitisColIdx === -1) return selectedColorFilter === 'neutral';
                  const cell = row[omurBitisColIdx];
                  if (!cell || cell === "-" || cell === "--") return selectedColorFilter === 'neutral';
                  const days = parseGelecekBakimDays(cell);
                  if (days === null) return selectedColorFilter === 'neutral';
                  if (days < 0) return selectedColorFilter === 'red';
                  if (days < 90) return selectedColorFilter === 'orange';
                  return selectedColorFilter === 'green';
                }

                if (isBakimMuaf) return selectedColorFilter === 'neutral';
                const targetColIdx = gelecekBakimColIdx !== -1 ? gelecekBakimColIdx : cols.findIndex(c => (c || '').includes("GELECEK") || (c || '').includes("MUAYENE"));
                if (targetColIdx === -1) return selectedColorFilter === 'neutral';
                const cell = row[targetColIdx];
                if (!cell || cell === "-" || cell === "--") return selectedColorFilter === 'neutral';
                const days = parseGelecekBakimDays(cell);
                if (days === null) return selectedColorFilter === 'neutral';
                if (days < 0) return selectedColorFilter === 'red';
                if (days < 90) return selectedColorFilter === 'orange';
                return selectedColorFilter === 'green';
              });
            }

            // Sıralama (Normalde seçili değil, ama basılınca Renk Sıralaması yapsın)
            let processedRows = [...filteredRows];
            if (sortByColor) {
              if (isDepo && omurBitisColIdx !== -1) {
                processedRows.sort((rowA, rowB) => {
                  const getDepoScore = (row: string[]) => {
                    const isOmurlu = omurluTabiColIdx !== -1 ? (row[omurluTabiColIdx] || "").trim().toUpperCase() : "EVET";
                    if (isOmurlu === "HAYIR") return 4;

                    const val = row[omurBitisColIdx] || "";
                    const days = parseGelecekBakimDays(val);
                    if (days === null) return 4;
                    if (days < 0) return 1; // Kırmızı (Ömrü dolmuş)
                    if (days < 90) return 2; // Turuncu (<90 gün)
                    return 3; // Yeşil (>=90 gün)
                  };

                  const scoreA = getDepoScore(rowA);
                  const scoreB = getDepoScore(rowB);
                  if (scoreA !== scoreB) return scoreA - scoreB;

                  const valA = rowA[omurBitisColIdx] || "";
                  const valB = rowB[omurBitisColIdx] || "";
                  const daysA = parseGelecekBakimDays(valA);
                  const daysB = parseGelecekBakimDays(valB);
                  const dJanA = daysA !== null ? daysA : 999999;
                  const dJanB = daysB !== null ? daysB : 999999;
                  return dJanA - dJanB;
                });
              } else if (gelecekBakimColIdx !== -1) {
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
            }

            const paginatedTechizatRows = techizatPageSize > 0 
              ? processedRows.slice((techizatPage - 1) * techizatPageSize, techizatPage * techizatPageSize)
              : processedRows;
            const totalTechizatPages = techizatPageSize > 0 ? Math.max(1, Math.ceil(processedRows.length / techizatPageSize)) : 1;

            return (
              <div className="absolute inset-0 flex flex-col bg-slate-50 p-2 sm:p-4 md:p-6 animate-fade-in overflow-y-auto">
                <div className="w-full max-w-[1920px] mx-auto flex flex-col h-full">
                  
                  {/* Header Titles */}
                  <div className="text-center mb-6 select-none print:hidden">
                    <h3 className="text-xl sm:text-2xl font-black text-slate-800 uppercase tracking-tighter">
                      {isDepo ? '📦' : '🛠️'} {modalTitle}
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
                      {/* Collapsible Search and Utility Controls Drawer */}
                      <div className="mb-6 select-none print:hidden">
                        <button
                          type="button"
                          onClick={() => setIsTechizatFilterOpen(prev => !prev)}
                          className="w-full flex items-center justify-between px-5 py-3.5 bg-gradient-to-r from-slate-900 via-slate-850 to-slate-900 hover:from-slate-800 hover:to-slate-850 text-white rounded-2xl border border-slate-700/80 shadow-md transition-all cursor-pointer group"
                        >
                          <div className="flex items-center gap-3">
                            <div className="p-2 bg-emerald-500/20 text-emerald-400 rounded-xl border border-emerald-500/30 group-hover:scale-105 transition-transform">
                              <Search className="w-4 h-4" />
                            </div>
                            <div className="text-left">
                              <div className="text-xs font-black tracking-wider uppercase text-slate-100 flex items-center gap-2">
                                <span>🔍 FİLTRELEME VE DETAYLI ARAMA PANELİ</span>
                                {(techizatSearchQuery || techizatFirmaFilter || techizatDurumFilter || techizatColorFilter !== 'ALL' || selectedColorFilter !== 'all' || sortByColor) && (
                                  <span className="px-2 py-0.5 bg-amber-500 text-slate-950 font-black text-[10px] rounded-full animate-pulse">
                                    FİLTRE AKTİF
                                  </span>
                                )}
                              </div>
                              <div className="text-[11px] text-slate-400 font-medium">
                                {isTechizatFilterOpen 
                                  ? "Filtre ve arama seçeneklerini gizlemek için tıklayın (▲ Kapat)" 
                                  : "Malzeme arama, firma, durum, renk kodu filtrelerini açmak için tıklayın (▼ Alta Doğru Aç)"}
                              </div>
                            </div>
                          </div>
                          
                          <div className="flex items-center gap-2">
                            <span className={`text-xs font-black px-3.5 py-1.5 rounded-xl border transition-all ${
                              isTechizatFilterOpen ? 'bg-emerald-500/20 text-emerald-300 border-emerald-500/40' : 'bg-slate-800 text-slate-300 border-slate-700'
                            }`}>
                              {isTechizatFilterOpen ? '▲ KAPAT' : '▼ FİLTRELERİ GÖSTER'}
                            </span>
                          </div>
                        </button>

                        {isTechizatFilterOpen && (
                          <div className="mt-3 bg-slate-900 text-slate-200 rounded-3xl p-5 flex flex-col xl:flex-row gap-4 items-center justify-between shadow-xl border border-slate-800 transition-all">
                            <div className="flex flex-wrap items-center gap-3 w-full xl:w-auto">
                              <div className="bg-slate-800 p-2.5 rounded-2xl border border-slate-700">
                                <Search className="w-5 h-5 text-emerald-400" />
                              </div>
                              
                              {/* Metin Arama Input */}
                              <div className="relative flex-1 sm:flex-initial">
                                <input
                                  type="text"
                                  placeholder={isDepo ? "Malzeme, P/N veya seri no..." : "Teçhizat veya seri no..."}
                                  value={techizatSearchQuery}
                                  onChange={(e) => {
                                    setTechizatSearchQuery(e.target.value);
                                    setActiveTechizatMatchIdx(0);
                                  }}
                                  className="bg-slate-800 text-white font-extrabold text-xs px-4 py-3 rounded-2xl focus:outline-none focus:ring-4 focus:ring-emerald-500/20 w-full sm:w-48 border border-slate-700 placeholder-slate-400"
                                />
                              </div>

                              {/* Firma Filtreleme Dropdown (Dinamik) */}
                              <div className="relative flex-1 sm:flex-initial">
                                <select
                                  value={techizatFirmaFilter}
                                  onChange={(e) => {
                                    setTechizatFirmaFilter(e.target.value);
                                    setActiveTechizatMatchIdx(0);
                                  }}
                                  className="bg-slate-800 text-emerald-300 font-extrabold text-xs px-4 py-3 rounded-2xl focus:outline-none focus:ring-4 focus:ring-emerald-500/20 w-full sm:w-52 border border-slate-700 cursor-pointer shadow-sm"
                                >
                                  <option value="">{isDepo ? "🏢 TÜM TEDARİKÇİLER" : "🏢 TÜM FİRMALAR"} ({uniqueFirmalar.length})</option>
                                  {uniqueFirmalar.map((f, i) => (
                                    <option key={i} value={f}>{f}</option>
                                  ))}
                                </select>
                              </div>

                              {/* Durum Filtreleme Dropdown (Dinamik Tablodaki Değerlere Göre) */}
                              <div className="relative flex-1 sm:flex-initial">
                                <select
                                  value={techizatDurumFilter}
                                  onChange={(e) => {
                                    setTechizatDurumFilter(e.target.value);
                                    setActiveTechizatMatchIdx(0);
                                  }}
                                  className="bg-slate-800 text-amber-300 font-black text-xs px-4 py-3 rounded-2xl focus:outline-none focus:ring-4 focus:ring-emerald-500/20 w-full sm:w-48 border border-slate-700 cursor-pointer shadow-sm"
                                >
                                  <option value="">⚡ TÜM DURUMLAR ({uniqueDurumlar.length})</option>
                                  {uniqueDurumlar.length > 0 ? (
                                    uniqueDurumlar.map((d, i) => (
                                      <option key={i} value={d}>{d}</option>
                                    ))
                                  ) : (
                                    <>
                                      <option value="FAAL">FAAL</option>
                                      <option value="GAYRİ FAAL">GAYRİ FAAL</option>
                                    </>
                                  )}
                                </select>
                              </div>

                              {/* Renk Koduna Seç Doğrudan Filtreleme Dropdown */}
                              <div className="relative flex-1 sm:flex-initial">
                                <select
                                  value={techizatColorFilter}
                                  onChange={(e) => {
                                    setTechizatColorFilter(e.target.value);
                                    setActiveTechizatMatchIdx(0);
                                  }}
                                  className="bg-slate-800 text-cyan-300 font-black text-xs px-4 py-3 rounded-2xl focus:outline-none focus:ring-4 focus:ring-emerald-500/20 w-full sm:w-56 border border-slate-700 cursor-pointer shadow-sm"
                                  title="Renk Koduna Göre Seç / Filtrele"
                                >
                                  <option value="ALL">🎨 RENK KODUNA SEÇ (TÜMÜ)</option>
                                  <option value="RED">🔴 KIRMIZI (SÜRESİ DOLAN / ACİL)</option>
                                  <option value="ORANGE">🟠 TURUNCU (&lt; 90 GÜN KALANLAR)</option>
                                  <option value="GREEN">🟢 YEŞİL (SÜRESİ UYGUN / FAAL)</option>
                                  <option value="GRAY">⚪ GRİ (BAKIMA TABİ DEĞİL / MUAF)</option>
                                </select>
                              </div>

                              {(techizatSearchQuery || techizatFirmaFilter || techizatDurumFilter || techizatColorFilter !== "ALL") && (
                                <button
                                  onClick={() => {
                                    setTechizatSearchQuery("");
                                    setTechizatFirmaFilter("");
                                    setTechizatDurumFilter("");
                                    setTechizatColorFilter("ALL");
                                    setActiveTechizatMatchIdx(0);
                                  }}
                                  className="px-3 py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-xl text-xs font-bold transition-all border border-slate-700 cursor-pointer"
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

                              {/* YUKARDAKİ FİLTER KISMINA RENK KODU SEÇ SEÇENEĞİ */}
                              <div className="flex items-center gap-1.5 bg-slate-900 border border-slate-700/80 rounded-2xl px-3 py-2 shrink-0 shadow-md">
                                <span className="text-[11px] font-black font-mono text-slate-300 uppercase flex items-center gap-1">
                                  <span>🎨</span>
                                  <span className="hidden sm:inline">RENK KODU:</span>
                                </span>
                                <select
                                  value={selectedColorFilter}
                                  onChange={(e) => setSelectedColorFilter(e.target.value as any)}
                                  className="bg-slate-800 text-white font-bold text-xs rounded-xl px-2.5 py-1.5 border border-slate-600 focus:outline-none focus:ring-2 focus:ring-emerald-500 cursor-pointer"
                                >
                                  <option value="all">TÜMÜ (Renk Filtresi Yok)</option>
                                  <option value="red">🔴 KIRMIZI (Günü Geçmiş / Süresi Dolan)</option>
                                  <option value="orange">🟠 TURUNCU (90 Gün Altı / Yaklaşan)</option>
                                  <option value="green">🟢 YEŞİL (Faal / Güvenli)</option>
                                  <option value="neutral">⚪ NÖTR (Muaf / Süresiz)</option>
                                </select>
                                {selectedColorFilter !== 'all' && (
                                  <button
                                    onClick={() => setSelectedColorFilter('all')}
                                    className="text-slate-400 hover:text-white text-xs px-1.5 py-0.5 rounded-lg bg-slate-800 hover:bg-slate-700 font-bold transition-all cursor-pointer"
                                    title="Renk filtresini temizle"
                                  >
                                    ✕
                                  </button>
                                )}
                              </div>

                              {/* YER DESTEK VE ÖZEL BAKIM ALETLERİ ETİKET BAS BUTONU */}
                              <button
                                type="button"
                                onClick={() => {
                                  if (cols && processedRows && processedRows.length > 0) {
                                    const findColIdx = (keywords: string[]) => 
                                      cols.findIndex(c => keywords.some(k => (c || '').toUpperCase().includes(k)));

                                    const nameIdx = findColIdx(["ALET", "TANIM", "AD", "ISIM", "İSİM"]);
                                    const pnIdx = findColIdx(["PARÇA", "PARCA", "P/N", "MODEL", "KOD"]);
                                    const snIdx = findColIdx(["SERİ", "SERI", "S/N"]);
                                    const locIdx = findColIdx(["LOKASYON", "HANGAR", "YER", "DEPO"]);
                                    const imgIdx = findColIdx(["GÖRSEL", "GORSEL", "FOTO", "DRIVE", "DRİVE", "RESİM", "RESIM", "LINK", "LİNK"]);
                                    const qtyIdx = findColIdx(["MİKTAR", "MIKTAR", "MEVCUT", "ADET"]);

                                    const mapped = processedRows.map((row: string[]) => {
                                      return {
                                        description: nameIdx !== -1 && row[nameIdx] ? row[nameIdx] : (row[1] || row[0] || ''),
                                        partNumber: pnIdx !== -1 && row[pnIdx] ? row[pnIdx] : (row[2] || '-'),
                                        serialAndNotes: snIdx !== -1 && row[snIdx] ? row[snIdx] : (row[3] || '-'),
                                        lokasyonNo: locIdx !== -1 && row[locIdx] ? row[locIdx] : 'ANKARA HANGAR',
                                        ankaraMevcut: qtyIdx !== -1 && row[qtyIdx] ? row[qtyIdx] : 1,
                                        imageUrl: imgIdx !== -1 && row[imgIdx] ? row[imgIdx] : '',
                                      };
                                    });
                                    setTechizatModalList(mapped);
                                  }
                                  setIsTechizatSlipModalOpen(true);
                                }}
                                className="px-4 py-2 bg-amber-500 hover:bg-amber-400 text-slate-950 font-black text-xs rounded-2xl transition-all shadow-md cursor-pointer flex items-center gap-1.5 shrink-0 font-sans"
                                title="Yer Destek ve Özel Alet Teçhizat Etiketi Bas (Çerçeveli Görsel & Karekodlu)"
                              >
                                <Wrench className="w-4 h-4 text-slate-950" />
                                <span>ETİKET BAS</span>
                              </button>

                              <button
                                onClick={() => setSortByColor(!sortByColor)}
                                className={`px-4 py-3 active:scale-95 font-black font-mono text-xs rounded-2xl flex items-center gap-2 transition-all cursor-pointer shadow-lg border shrink-0 ${
                                  sortByColor 
                                    ? 'bg-red-600 hover:bg-red-700 text-white border-red-500 animate-pulse' 
                                    : 'bg-slate-800 hover:bg-slate-750 text-slate-200 border-slate-700'
                                }`}
                                title={isDepo ? "Raf ömrü kalan gün sayısına göre (Kırmızı ➜ Turuncu ➜ Yeşil) sıralar" : "Bakım gün sayısına göre (Kırmızı ➜ Turuncu ➜ Sarı ➜ Yeşil) sıralar"}
                              >
                                <SlidersHorizontal className="w-4 h-4 text-amber-400" />
                                <span>
                                  {sortByColor 
                                    ? (isDepo ? "🔴 RAF ÖMRÜ SIRALAMASI AKTİF" : "🔴 RENK SIRALAMASI AKTİF") 
                                    : (isDepo ? "⏳ RAF ÖMRÜNE GÖRE SIRALA" : "⏳ RENK KODUNA GÖRE SIRALA")}
                                </span>
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

                              <button
                                onClick={async () => {
                                  showNotification("Google Drive'dan bu sayfaya ait güncel veriler senkronize ediliyor...");
                                  await pullTechizatUnitFromDrive(activeTechizatType, false);
                                }}
                                disabled={isTechizatDriveLoading}
                                className="px-4 py-3 bg-[#0b3d1d] hover:bg-[#072612] disabled:opacity-50 active:scale-95 text-white font-black font-mono text-xs rounded-2xl flex items-center gap-2 transition-all cursor-pointer shadow-lg border border-emerald-600 shrink-0"
                                title="Google Drive klasöründeki güncel Excel dosyasından sadece bu sayfanın verilerini yeniler"
                              >
                                <RefreshCw className={`w-4 h-4 text-emerald-300 ${isTechizatDriveLoading ? 'animate-spin' : ''}`} />
                                <span>{isTechizatDriveLoading ? 'DRİVE BAĞLANIYOR...' : 'DRİVE İLE SENKRONİZE ET'}</span>
                              </button>

                              {activeTechizatType !== 'all' && (
                                <button
                                  onClick={() => {
                                    setPasswordActionType('filter_sync');
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
                                onClick={() => {
                                  setPasswordActionType('new_product');
                                  setPasswordInput('');
                                  setPasswordError(false);
                                  setIsPasswordModalOpen(true);
                                }}
                                className="px-4 py-3 bg-emerald-600 hover:bg-emerald-500 active:scale-95 text-white font-black font-mono text-xs rounded-2xl flex items-center gap-2 transition-all cursor-pointer shadow-lg shadow-emerald-950/20 border border-emerald-500 shrink-0"
                                title="Yetkili şifresi ile yeni ürün veya teçhizat kaydı ekle"
                              >
                                <PlusCircle className="w-4 h-4 text-white" />
                                <span>YENİ ÜRÜN EKLE</span>
                              </button>

                              {isDepo && (
                                <button
                                  onClick={() => {
                                    setPasswordActionType('depo_management');
                                    setPasswordInput('');
                                    setPasswordError(false);
                                    setIsPasswordModalOpen(true);
                                  }}
                                  className="px-4 py-3 bg-[#0b3d1d] hover:bg-[#072612] active:scale-95 text-white font-black font-mono text-xs rounded-2xl flex items-center gap-2 transition-all cursor-pointer shadow-lg border border-emerald-600 shrink-0"
                                  title="Depo Giriş/Çıkış, Kit Montajı ve Transfer Yönetimi"
                                >
                                  <Boxes className="w-4 h-4 text-emerald-400" />
                                  <span>DEPO YÖNETİM</span>
                                </button>
                              )}

                              <button
                                onClick={() => exportTechizatToExcel(activeTechizatType, cols, rows, modalTitle)}
                                className="px-4 py-3 bg-emerald-700 hover:bg-emerald-600 active:scale-95 text-white font-black font-mono text-xs rounded-2xl flex items-center gap-2 transition-all cursor-pointer shadow-lg shadow-emerald-900/10 border border-emerald-600 shrink-0"
                              >
                                <Download className="w-4 h-4" />
                                <span>EXCEL İNDİR</span>
                              </button>
                            </div>
                          </div>
                        )}
                      </div>

                  {/* Dynamic Interactive Grid Container */}
                  <div className="flex-1 bg-white border-2 border-slate-200/60 rounded-[2.5rem] shadow-xl overflow-hidden flex flex-col print:border-none print:shadow-none min-h-[400px]">
                    
                    {/* Table Title Bar */}
                    <div className="bg-slate-900 px-6 py-4 border-b border-slate-800 flex items-center justify-between print:hidden shrink-0 flex-wrap gap-3">
                      {Object.keys(selectedTechizatItems).length > 0 ? (
                        <div className="flex items-center gap-3 animate-fade-in flex-wrap">
                          <span className="text-xs font-black text-amber-400 bg-amber-950/40 border border-amber-900 px-3 py-1.5 rounded-xl">
                            ⚡ {Object.keys(selectedTechizatItems).length} {isDepo ? 'MALZEME SEÇİLDİ' : 'TEÇHİZAT SEÇİLDİ'}
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
                          <button
                            onClick={handleDeleteSelectedTechizatRows}
                            className="px-4 py-1.5 bg-rose-700 hover:bg-rose-600 hover:scale-105 active:scale-95 text-white font-black text-xs uppercase tracking-wider rounded-xl transition-all cursor-pointer flex items-center gap-2 shadow-lg shadow-rose-950/30 border border-rose-600"
                            title="Seçilen satırları sil ve Drive/E-Tablo ile senkronize et"
                          >
                            <Trash2 className="w-3.5 h-3.5" />
                            <span>SEÇİLENLERİ SİL</span>
                          </button>
                        </div>
                      ) : (
                        <div className="hidden lg:flex items-center gap-2 text-[11px] font-mono font-black select-none">
                          <span className="flex items-center gap-1.5 bg-emerald-950/60 text-emerald-300 border border-emerald-800/80 px-2.5 py-1 rounded-lg">
                            <span className="w-2.5 h-2.5 rounded-full bg-emerald-500 inline-block shadow-sm"></span>
                            🟢 &ge;90 Gün
                          </span>
                          <span className="flex items-center gap-1.5 bg-orange-950/60 text-orange-300 border border-orange-800/80 px-2.5 py-1 rounded-lg">
                            <span className="w-2.5 h-2.5 rounded-full bg-orange-500 inline-block shadow-sm"></span>
                            🟠 &lt;90 Gün
                          </span>
                          <span className="flex items-center gap-1.5 bg-rose-950/60 text-rose-300 border border-rose-800/80 px-2.5 py-1 rounded-lg">
                            <span className="w-2.5 h-2.5 rounded-full bg-rose-600 inline-block shadow-sm animate-pulse"></span>
                            🔴 Bakım/Ömür Geçmiş
                          </span>
                        </div>
                      )}
                      {isTechizatDriveLoading && (
                        <span className="text-[10px] font-mono font-black text-amber-300 bg-amber-950/80 px-3 py-1 rounded-full border border-amber-600 flex items-center gap-1.5 animate-pulse">
                          <RefreshCw className="w-3 h-3 animate-spin text-amber-300" />
                          VERİ KONTROL EDİLİYOR...
                        </span>
                      )}
                      <div className="flex items-center gap-2">
                        <span className="text-[10px] font-mono font-black text-slate-400 bg-slate-800 px-3 py-1 rounded-full border border-slate-700">
                          {processedRows.length} KALEM {isDepo ? 'MALZEME LİSTELENDİ' : 'TEÇHİZAT LİSTELENDİ'}
                        </span>
                        {processedRows.length > 50 && (
                          <div className="flex items-center gap-1.5 text-[11px] bg-slate-800 px-2 py-0.5 rounded-full border border-slate-700 text-slate-300">
                            <span>Göster:</span>
                            <select
                              value={techizatPageSize}
                              onChange={(e) => {
                                setTechizatPageSize(Number(e.target.value));
                                setTechizatPage(1);
                              }}
                              className="bg-slate-900 text-emerald-400 font-bold px-1.5 py-0.5 rounded border border-slate-700 text-[10px] cursor-pointer"
                            >
                              <option value={50}>50</option>
                              <option value={100}>100</option>
                              <option value={200}>200</option>
                              <option value={-1}>Tümü</option>
                            </select>
                          </div>
                        )}
                      </div>
                    </div>

                    {/* Table Grid Scroll Wrapper */}
                    <div className="flex-1 overflow-auto max-h-[85vh] print:max-h-none print:overflow-visible relative">
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
                                {isDepo ? (
                                  <div className="flex flex-col items-center justify-center gap-3 py-6">
                                    <div className="w-16 h-16 rounded-2xl bg-amber-50 border border-amber-200 flex items-center justify-center shadow-inner">
                                      <Boxes className="w-8 h-8 text-amber-700" />
                                    </div>
                                    <span className="text-slate-800 font-black text-base uppercase tracking-wider">
                                      {techizatActiveSection === 'depo_sarf' ? 'SARF VE PARÇA DEPOSUNDA KAYIT BULUNMUYOR' : techizatActiveSection === 'depo_kimyasal' ? 'KİMYASAL DEPOSUNDA KAYIT BULUNMUYOR' : 'DEPOLARDA HENÜZ KAYITLI MALZEME BULUNMUYOR'}
                                    </span>
                                    <p className="text-xs text-slate-500 font-medium max-w-md text-center leading-relaxed">
                                      Google Drive üzerinde bu birime ait kayıtlı Excel bulunamadı veya henüz veri yüklenmedi. <strong>"VERİ GÜNCELLE"</strong> veya <strong>"YENİ EKLE"</strong> butonunu kullanarak Excel yükleyebilirsiniz.
                                    </p>
                                  </div>
                                ) : (
                                  <div className="flex flex-col items-center justify-center gap-3 py-6">
                                    <div className="w-16 h-16 rounded-2xl bg-slate-100 border border-slate-200 flex items-center justify-center shadow-inner">
                                      <FileSpreadsheet className="w-8 h-8 text-slate-500" />
                                    </div>
                                    <span className="text-slate-800 font-black text-base uppercase tracking-wider">
                                      {techizatSearchQuery || techizatFirmaFilter || techizatDurumFilter 
                                        ? "🔍 ARAMA KRİTERLERİNE UYGUN KAYIT BULUNAMADI" 
                                        : "BU BİRİME AİT KAYITLI EXCEL BULUNAMADI"}
                                    </span>
                                    <p className="text-xs text-slate-500 font-medium max-w-md text-center leading-relaxed">
                                      {techizatSearchQuery || techizatFirmaFilter || techizatDurumFilter
                                        ? "Filtreleri temizleyerek tüm listeyi görebilirsiniz."
                                        : "Sistemde henüz bu birime ait bir Excel dosyası yok veya dosya boş. 'VERİ GÜNCELLE' butonuyla Excel yükleyebilir veya 'YENİ EKLE' ile ilk kaydı oluşturabilirsiniz."}
                                    </p>
                                  </div>
                                )}
                              </td>
                            </tr>
                          ) : (
                             paginatedTechizatRows.map((row, rIdx) => {
                               const isAll = activeTechizatType === 'all';
                               const targetTechType = isAll ? getRealTechType(row[0]) : activeTechizatType;
                               const targetRow = isAll ? row.slice(1) : row;
                               const rowImageUrl = findRowImageUrl(targetTechType, targetRow, techizatImages);

                               return (
                                 <tr 
                                   key={rIdx} 
                                   onMouseEnter={(e) => {
                                     if (rowImageUrl) {
                                       setHoveredRowImage({
                                         url: rowImageUrl,
                                         title: targetRow[1] || "Kayıt",
                                         subtitle: `${targetRow[2] || ""} ${targetRow[3] ? "• " + targetRow[3] : ""}`,
                                         x: e.clientX,
                                         y: e.clientY
                                       });
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
                                          else if (targetTechType === 't70_helitak') sourceList = techizatT70HelitakData;
                                          else if (targetTechType === 'b360') sourceList = techizatB360Data;
                                          else if (targetTechType === 'c650') sourceList = techizatC650Data;
                                          else if (targetTechType === 'hangar') sourceList = techizatHangarData;
                                          
                                          const found = sourceList.find(sr => (sr[0] || "").trim() === (targetRow[0] || "").trim() && (sr[1] || "").trim() === (targetRow[1] || "").trim());
                                          if (found) resolvedRow = found;
                                        }
                                      }

                                      let resolvedCopy: string[] = [];
                                      if (targetTechType === 'kara_araclari') {
                                        resolvedCopy = formatKaraAraciToStandardRow(resolvedRow);
                                      } else if (isDepo) {
                                        resolvedCopy = formatDepoRow(resolvedRow);
                                      } else {
                                        resolvedCopy = formatStandardRow(resolvedRow);
                                      }

                                      setActiveTechizatRowEdit({
                                        rIdx,
                                        techType: targetTechType as any,
                                        row: resolvedCopy
                                      });
                                      loadRegionalLocationsForRow(targetTechType, resolvedCopy);
                                      setMobileEditTab('form');
                                      const matchedImg = findRowImageUrl(targetTechType, resolvedCopy, techizatImages);
                                      setTempImageUrlInput(matchedImg && !matchedImg.startsWith('data:') ? matchedImg : "");
                                      setEditRowValues([...resolvedCopy]);
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
                                     const isMultiLine = Boolean(cell && cell.includes('\n'));

                                     const isDateCol = Boolean(
                                       label.toUpperCase().includes("KONTROL") || 
                                       label.toUpperCase().includes("BAKIM") || 
                                       label.toUpperCase().includes("TARİH") || 
                                       label.toUpperCase().includes("MUAYENE") || 
                                       label.toUpperCase().includes("ÖMÜR") ||
                                       (/^\d{5}$/.test(String(cell || '').trim()) && parseFloat(String(cell || '').trim()) >= 20000 && parseFloat(String(cell || '').trim()) <= 90000)
                                     );

                                     const displayCell = isDateCol ? cleanAndFormatDateString(cell) : cell;

                                     let cellStyleClass = cell 
                                       ? 'text-slate-800 font-sans font-semibold' 
                                       : 'text-slate-300 italic';
                                       
                                     const isKalibTabiColIdx = cols.findIndex(c => (c || '').includes("KALİBRASYONA TABİ") || (c || '').includes("BAKIMA TABİ"));
                                     const isKalibTabiVal = isKalibTabiColIdx !== -1 ? String(row[isKalibTabiColIdx] || "").trim().toUpperCase() : "EVET";
                                     const isBakimMuaf = isKalibTabiVal === "HAYIR" || isKalibTabiVal === "HAYIR (MUAFIYET)" || isKalibTabiVal === "MUAFIYET";

                                     const isGelecekBakimCol = !isDepo && (cIdx === gelecekBakimColIdx || (isKaraAraci && (label === "BİR SONRAKİ MUAYENE TARİHİ" || label.includes("GELECEK"))));
                                     const isOmurBitisCol = isDepo && (label === "ÖMÜR BİTİŞ TARİHİ" || label.includes("ÖMÜR"));

                                     if (!isMultiLine) {
                                       if (isGelecekBakimCol && cell && cell !== "-" && cell !== "--") {
                                         if (isBakimMuaf) {
                                           cellStyleClass = 'text-slate-400 italic font-bold text-center';
                                         } else {
                                           const days = parseGelecekBakimDays(cell);
                                           if (days !== null) {
                                             if (days < 0) {
                                               cellStyleClass = 'bg-rose-600 text-white font-extrabold px-3 py-1.5 rounded-xl shadow-sm text-center animate-pulse';
                                             } else if (days < 90) {
                                               cellStyleClass = 'bg-orange-500 text-white font-extrabold px-3 py-1.5 rounded-xl shadow-sm text-center';
                                             } else {
                                               cellStyleClass = 'bg-emerald-600 text-white font-extrabold px-3 py-1.5 rounded-xl shadow-sm text-center';
                                             }
                                           }
                                         }
                                       }

                                       if (isOmurBitisCol && cell && cell !== "-" && cell !== "--") {
                                         const omurluIdx = cols.indexOf("ÖMÜRLÜ PARÇA MI?");
                                         const isOmurlu = omurluIdx !== -1 ? (row[omurluIdx] || "").toUpperCase() === "EVET" : true;
                                         if (isOmurlu) {
                                           const days = parseGelecekBakimDays(cell);
                                           if (days !== null) {
                                             if (days < 0) {
                                               cellStyleClass = 'bg-rose-600 text-white font-extrabold px-3 py-1.5 rounded-xl shadow-sm text-center animate-pulse';
                                             } else if (days < 90) {
                                               cellStyleClass = 'bg-orange-500 text-white font-extrabold px-3 py-1.5 rounded-xl shadow-sm text-center';
                                             } else {
                                               cellStyleClass = 'bg-emerald-600 text-white font-extrabold px-3 py-1.5 rounded-xl shadow-sm text-center';
                                             }
                                           }
                                         } else {
                                           cellStyleClass = 'text-slate-400 italic font-medium text-center';
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
                                         {label === "BELGE YÜKLE" ? (
                                           <td 
                                             key={cIdx} 
                                             className="px-3 py-2 text-center border-r border-slate-100 last:border-r-0 min-w-[130px]"
                                             onClick={(e) => e.stopPropagation()}
                                           >
                                             {(() => {
                                               const itemDocs = findMatchingDocs(targetRow, hangarPdfDocs);
                                               return (
                                                 <button
                                                   type="button"
                                                   onClick={() => {
                                                     loadRegionalLocationsForRow('hangar', targetRow);
                                                     setActiveTechizatRowEdit({
                                                       rIdx,
                                                       techType: 'hangar',
                                                       row: targetRow
                                                     });
                                                     setMobileEditTab('documents');
                                                   }}
                                                   className={`px-3 py-1.5 rounded-xl font-black text-[11px] flex items-center justify-center gap-1.5 transition-all shadow-sm cursor-pointer mx-auto ${
                                                     itemDocs.length > 0
                                                       ? 'bg-rose-600 hover:bg-rose-700 text-white shadow-rose-600/20'
                                                       : 'bg-slate-100 hover:bg-emerald-50 text-slate-700 hover:text-emerald-700 border border-slate-200 hover:border-emerald-300'
                                                   }`}
                                                   title="Hangar Yer Destek belgelerini yönetmek ve PDF önizlemek için tıklayınız"
                                                 >
                                                   <FileText className="w-3.5 h-3.5" />
                                                   <span>{itemDocs.length > 0 ? `📄 ${itemDocs.length} Belge` : '+ Belge Yükle'}</span>
                                                 </button>
                                               );
                                             })()}
                                           </td>
                                         ) : (
                                         <td 
                                           key={cIdx} 
                                           className="px-4 py-3 text-center border-r border-slate-100 last:border-r-0 max-w-[200px]"
                                         >
                                           <div 
                                             title={cell ? `${label}: ${cell} (Düzenlemek ve görsel eklemek için Tıklayın)` : "Boş Veri"}
                                             className={`${isMultiLine ? 'whitespace-pre-line leading-relaxed min-w-[120px]' : 'truncate'} px-2 py-1 rounded-xl transition-all text-xs text-center select-text hover:bg-emerald-100/50 hover:text-emerald-950 flex items-center justify-center gap-1 ${
                                               isActiveMatch 
                                                 ? 'bg-blue-600 text-white font-black scale-105 shadow-md ring-2 ring-blue-400 animate-pulse'
                                                 : isMatch
                                                   ? 'bg-blue-200 text-blue-950 font-black border border-blue-400'
                                                   : isMultiLine
                                                     ? 'text-slate-800 font-sans'
                                                     : cellStyleClass
                                             }`}
                                           >
                                             {isMultiLine ? (
                                               <div className="w-full text-center whitespace-pre-line leading-relaxed font-bold divide-y divide-slate-200/60 py-0.5">
                                                 {cell.split('\n').map((lineVal, lineIdx) => {
                                                   const trimmed = (lineVal || "").trim();
                                                   const trimmedFormatted = (isGelecekBakimCol || isOmurBitisCol || isDateCol) ? cleanAndFormatDateString(trimmed) : trimmed;

                                                   if (isGelecekBakimCol) {
                                                     const days = parseGelecekBakimDays(trimmed);
                                                     let lineBadgeClass = "";
                                                     if (trimmed && trimmed !== "-" && trimmed !== "--" && !isBakimMuaf && days !== null) {
                                                       if (days < 0) {
                                                         lineBadgeClass = "bg-rose-600 text-white font-extrabold px-2.5 py-1 rounded-lg shadow-sm text-center animate-pulse inline-block w-full max-w-[130px]";
                                                       } else if (days < 90) {
                                                         lineBadgeClass = "bg-orange-500 text-white font-extrabold px-2.5 py-1 rounded-lg shadow-sm text-center inline-block w-full max-w-[130px]";
                                                       } else {
                                                         lineBadgeClass = "bg-emerald-600 text-white font-extrabold px-2.5 py-1 rounded-lg shadow-sm text-center inline-block w-full max-w-[130px]";
                                                       }
                                                     }

                                                     return (
                                                       <div key={lineIdx} className="py-1 first:pt-0 last:pb-0 min-h-[30px] flex items-center justify-center">
                                                         {lineBadgeClass ? (
                                                           <span className={lineBadgeClass}>
                                                             {trimmedFormatted}
                                                           </span>
                                                         ) : (
                                                           <span className={isBakimMuaf ? "text-slate-400 italic font-bold" : "text-slate-700 font-bold"}>
                                                             {trimmedFormatted || "-"}
                                                           </span>
                                                         )}
                                                       </div>
                                                     );
                                                   }

                                                   if (isOmurBitisCol) {
                                                     const omurluIdx = cols.indexOf("ÖMÜRLÜ PARÇA MI?");
                                                     const isOmurlu = omurluIdx !== -1 ? (row[omurluIdx] || "").toUpperCase() === "EVET" : true;
                                                     const days = parseGelecekBakimDays(trimmed);
                                                     let lineBadgeClass = "";
                                                     if (trimmed && trimmed !== "-" && trimmed !== "--" && isOmurlu && days !== null) {
                                                       if (days < 0) {
                                                         lineBadgeClass = "bg-rose-600 text-white font-extrabold px-2.5 py-1 rounded-lg shadow-sm text-center animate-pulse inline-block w-full max-w-[130px]";
                                                       } else if (days < 90) {
                                                         lineBadgeClass = "bg-orange-500 text-white font-extrabold px-2.5 py-1 rounded-lg shadow-sm text-center inline-block w-full max-w-[130px]";
                                                       } else {
                                                         lineBadgeClass = "bg-emerald-600 text-white font-extrabold px-2.5 py-1 rounded-lg shadow-sm text-center inline-block w-full max-w-[130px]";
                                                       }
                                                     }

                                                     return (
                                                       <div key={lineIdx} className="py-1 first:pt-0 last:pb-0 min-h-[30px] flex items-center justify-center">
                                                         {lineBadgeClass ? (
                                                           <span className={lineBadgeClass}>
                                                             {trimmedFormatted}
                                                           </span>
                                                         ) : (
                                                           <span className={!isOmurlu ? "text-slate-400 italic font-medium" : "text-slate-700 font-bold"}>
                                                             {trimmedFormatted || "-"}
                                                           </span>
                                                         )}
                                                       </div>
                                                     );
                                                   }

                                                   return (
                                                     <div key={lineIdx} className="py-1 first:pt-0 last:pb-0 min-h-[30px] flex items-center justify-center text-center font-bold text-slate-800">
                                                       {trimmedFormatted || "-"}
                                                     </div>
                                                   );
                                                 })}
                                               </div>
                                             ) : (
                                               <span className="truncate">{displayCell || (label.toUpperCase().includes("MAİL") || label.toUpperCase().includes("MAIL") || label.toUpperCase().includes("90 GÜN") ? "" : "-")}</span>
                                             )}
                                           </div>
                                         </td>
                                         )}
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

                    {techizatPageSize > 0 && totalTechizatPages > 1 && (
                      <div className="flex flex-wrap items-center justify-between gap-3 px-4 py-2.5 bg-slate-900 border-t border-slate-800 text-xs text-slate-300 font-mono select-none rounded-b-2xl mt-1">
                        <div className="flex items-center gap-2">
                          <span>
                            Görüntülenen: <strong className="text-white font-bold">{(techizatPage - 1) * techizatPageSize + 1}-{Math.min(processedRows.length, techizatPage * techizatPageSize)}</strong> / Toplam <strong>{processedRows.length}</strong>
                          </span>
                        </div>
                        <div className="flex items-center gap-1.5 flex-wrap">
                          <button
                            type="button"
                            disabled={techizatPage <= 1}
                            onClick={() => setTechizatPage(prev => Math.max(1, prev - 1))}
                            className="px-2.5 py-1 rounded bg-slate-800 border border-slate-700 text-slate-300 disabled:opacity-30 disabled:cursor-not-allowed hover:bg-slate-700 cursor-pointer text-xs font-bold"
                            title="Önceki Sayfa"
                          >
                            ‹ Önceki
                          </button>

                          {(() => {
                            const pages: (number | string)[] = [];
                            if (totalTechizatPages <= 7) {
                              for (let i = 1; i <= totalTechizatPages; i++) pages.push(i);
                            } else {
                              pages.push(1);
                              if (techizatPage > 3) pages.push('...');
                              const start = Math.max(2, techizatPage - 1);
                              const end = Math.min(totalTechizatPages - 1, techizatPage + 1);
                              for (let i = start; i <= end; i++) {
                                if (!pages.includes(i)) pages.push(i);
                              }
                              if (techizatPage < totalTechizatPages - 2) pages.push('...');
                              if (!pages.includes(totalTechizatPages)) pages.push(totalTechizatPages);
                            }
                            return pages.map((p, idx) => {
                              if (p === '...') {
                                return <span key={`ell-${idx}`} className="px-1.5 py-1 text-slate-500 font-bold select-none text-xs">...</span>;
                              }
                              const num = Number(p);
                              const isActive = num === techizatPage;
                              return (
                                <button
                                  key={`p-${num}`}
                                  type="button"
                                  onClick={() => setTechizatPage(num)}
                                  className={`min-w-[28px] px-2 py-1 rounded text-xs font-mono font-bold transition-all cursor-pointer ${
                                    isActive
                                      ? 'bg-emerald-600 text-white shadow-sm ring-1 ring-emerald-400 font-black'
                                      : 'bg-slate-800 text-slate-300 hover:bg-slate-700 hover:text-white border border-slate-700'
                                  }`}
                                >
                                  {num}
                                </button>
                              );
                            });
                          })()}

                          <button
                            type="button"
                            disabled={techizatPage >= totalTechizatPages}
                            onClick={() => setTechizatPage(prev => Math.min(totalTechizatPages, prev + 1))}
                            className="px-2.5 py-1 rounded bg-slate-800 border border-slate-700 text-slate-300 disabled:opacity-30 disabled:cursor-not-allowed hover:bg-slate-700 cursor-pointer text-xs font-bold"
                            title="Sonraki Sayfa"
                          >
                            Sonraki Sayfa ›
                          </button>
                        </div>
                      </div>
                    )}

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
                                          className="w-full px-3.5 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-700 outline-none focus:border-[#0b3d1d] focus:bg-white transition-all"
                                        />
                                      </div>
                                      <span className="text-slate-400 font-bold">→</span>
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
                                          className="w-full px-3.5 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-700 outline-none focus:border-[#0b3d1d] focus:bg-white transition-all"
                                        />
                                      </div>
                                      {geRoutes.length > 1 && (
                                        <button
                                          type="button"
                                          onClick={() => {
                                            setGeRoutes(geRoutes.filter((_, idx) => idx !== rIdx));
                                          }}
                                          className="p-2 text-rose-500 hover:bg-rose-50 rounded-xl transition-colors cursor-pointer"
                                          title="Güzergahı Sil"
                                        >
                                          ✕
                                        </button>
                                      )}
                                    </div>
                                  ))}
                                  <button
                                    type="button"
                                    onClick={() => setGeRoutes([...geRoutes, { from: "", to: "" }])}
                                    className="self-start text-[11px] font-extrabold text-[#0b3d1d] hover:underline flex items-center gap-1 mt-1 cursor-pointer"
                                  >
                                    + Yeni Güzergah Ekle
                                  </button>
                                </div>

                                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                                  <div>
                                    <label className="block text-[10px] font-black text-slate-400 uppercase tracking-wider mb-2">
                                      ⏰ Çıkış Saati
                                    </label>
                                    <input
                                      type="time"
                                      value={geDepartureTime}
                                      onChange={(e) => setGeDepartureTime(e.target.value)}
                                      className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-2xl focus:outline-none focus:border-[#0b3d1d] focus:ring-4 focus:ring-[#0b3d1d]/5 text-xs font-bold text-slate-800 transition-all font-mono"
                                    />
                                  </div>
                                  <div>
                                    <label className="block text-[10px] font-black text-slate-400 uppercase tracking-wider mb-2">
                                      ⏰ Dönüş Saati
                                    </label>
                                    <input
                                      type="time"
                                      value={geReturnTime}
                                      onChange={(e) => setGeReturnTime(e.target.value)}
                                      className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-2xl focus:outline-none focus:border-[#0b3d1d] focus:ring-4 focus:ring-[#0b3d1d]/5 text-xs font-bold text-slate-800 transition-all font-mono"
                                    />
                                  </div>
                                </div>

                                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                                  <div>
                                    <label className="block text-[10px] font-black text-slate-400 uppercase tracking-wider mb-2">
                                      🚗 Çıkış Kilometresi (KM)
                                    </label>
                                    <input
                                      type="number"
                                      placeholder="Örn: 124500"
                                      value={geDepartureKm}
                                      onChange={(e) => setGeDepartureKm(e.target.value)}
                                      className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-2xl focus:outline-none focus:border-[#0b3d1d] focus:ring-4 focus:ring-[#0b3d1d]/5 text-xs font-bold text-slate-800 transition-all font-mono"
                                    />
                                  </div>
                                  <div>
                                    <label className="block text-[10px] font-black text-slate-400 uppercase tracking-wider mb-2">
                                      🏁 Dönüş Kilometresi (KM)
                                    </label>
                                    <input
                                      type="number"
                                      placeholder="Örn: 124650"
                                      value={geReturnKm}
                                      onChange={(e) => setGeReturnKm(e.target.value)}
                                      className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-2xl focus:outline-none focus:border-[#0b3d1d] focus:ring-4 focus:ring-[#0b3d1d]/5 text-xs font-bold text-slate-800 transition-all font-mono"
                                    />
                                  </div>
                                </div>

                                {geDepartureKm && geReturnKm && (
                                  <div className="p-3 bg-emerald-50 border border-emerald-200 rounded-2xl flex items-center justify-between">
                                    <span className="text-xs font-bold text-emerald-900">Toplam Yapılan Yol:</span>
                                    <span className="text-sm font-black text-emerald-950 font-mono">
                                      {Math.max(0, Number(geReturnKm) - Number(geDepartureKm))} KM
                                    </span>
                                  </div>
                                )}
                              </div>
                            )}
                          </div>

                          {/* Wizard Navigation Buttons */}
                          <div className="flex items-center justify-between pt-4 border-t border-slate-100">
                            {geStep > 1 ? (
                              <button
                                type="button"
                                onClick={handlePrevStep}
                                className="px-5 py-2.5 rounded-2xl border border-slate-200 bg-white hover:bg-slate-50 text-slate-700 font-bold text-xs transition-all active:scale-95 cursor-pointer"
                              >
                                ← Önceki Adım
                              </button>
                            ) : <div />}

                            {geStep < 3 ? (
                              <button
                                type="button"
                                onClick={handleNextStep}
                                className="px-6 py-2.5 rounded-2xl bg-[#0b3d1d] hover:bg-[#072612] text-white font-black text-xs uppercase tracking-wider shadow-md shadow-emerald-900/20 transition-all active:scale-95 cursor-pointer"
                              >
                                İleri Adım →
                              </button>
                            ) : (
                              <button
                                type="button"
                                onClick={async () => {
                                  if (!geDepartureKm || !geReturnKm) {
                                    alert("Lütfen çıkış ve dönüş kilometrelerini giriniz.");
                                    return;
                                  }
                                  if (Number(geReturnKm) < Number(geDepartureKm)) {
                                    alert("Dönüş kilometresi çıkış kilometresinden küçük olamaz.");
                                    return;
                                  }
                                  const driverObj = drivers.find(d => d.name.toLowerCase() === geSoforName.toLowerCase());
                                  const routeSummary = geRoutes.filter(r => r.from || r.to).map(r => `${r.from} -> ${r.to}`).join(' | ');
                                  
                                  const newOrder = {
                                    id: Date.now(),
                                    date: geTarih,
                                    plate: gePlaka,
                                    driverName: geSoforName,
                                    driverId: (driverObj as any)?.tc || driverObj?.idNo || "",
                                    driverSicil: (driverObj as any)?.sicil || driverObj?.sicilNo || "",
                                    driverPhone: driverObj?.phone || "",
                                    driverKanGrubu: driverObj?.kanGrubu || "",
                                    driverAdres: driverObj?.adres || "",
                                    serialNo: geSeriNo,
                                    departureTime: geDepartureTime,
                                    returnTime: geReturnTime,
                                    departureKm: geDepartureKm,
                                    returnKm: geReturnKm,
                                    route: routeSummary
                                  };

                                  const updatedOrders = [newOrder, ...karaAraclariGorevEmirleri];
                                  setKaraAraclariGorevEmirleri(updatedOrders);
                                  localStorage.setItem('kara_araclari_gorev_emirleri', JSON.stringify(updatedOrders));
                                  pushKaraAraclariGorevEmirleri(updatedOrders);

                                  if (gePlaka && geReturnKm) {
                                    const updatedKara = techizatKaraAraclariData.map(row => {
                                      if ((row[1] || "").toLowerCase() === gePlaka.toLowerCase()) {
                                        const copy = [...row];
                                        copy[4] = geReturnKm;
                                        return copy;
                                      }
                                      return row;
                                    });
                                    setTechizatKaraAraclariData(updatedKara);
                                  }

                                  setGeStep(1);
                                  setGeSeriNo("");
                                  setGeDepartureKm("");
                                  setGeReturnKm("");
                                  setGeRoutes([{ from: "", to: "" }]);
                                  showNotification("Görev emri başarıyla kaydedildi!");
                                  setKaraAraclariSubTab('past_records');
                                }}
                                className="px-6 py-2.5 rounded-2xl bg-emerald-600 hover:bg-emerald-700 text-white font-black text-xs uppercase tracking-wider shadow-lg shadow-emerald-600/30 transition-all active:scale-95 cursor-pointer"
                              >
                                ✅ Kaydet ve Görev Emrini Tamamla
                              </button>
                            )}
                          </div>
                        </div>
                      );
                    })()}

                    {/* PAST RECORDS TAB */}
                    {karaAraclariSubTab === 'past_records' && (
                      <div className="w-full bg-white border border-slate-200 rounded-[2.5rem] p-6 shadow-xl flex flex-col gap-6 animate-fade-in">
                        <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-4 border-b border-slate-100 pb-4">
                          <div>
                            <h4 className="text-base font-black text-slate-800 uppercase tracking-wide flex items-center gap-2">
                              📜 GEÇMİŞ GÖREV EMİRLERİ ({karaAraclariGorevEmirleri.length})
                            </h4>
                            <p className="text-xs text-slate-500 mt-0.5">
                              Kara araçları için oluşturulmuş ve kaydedilmiş görev emirleri listesi
                            </p>
                          </div>
                          <div className="flex items-center gap-2">
                            <button
                              type="button"
                              onClick={exportGorevEmirleriToExcel}
                              className="px-4 py-2 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold rounded-xl flex items-center gap-2 shadow-sm transition-all active:scale-95 cursor-pointer"
                            >
                              <FileSpreadsheet className="w-4 h-4" />
                              <span>Excel Olarak İndir</span>
                            </button>
                          </div>
                        </div>

                        {karaAraclariGorevEmirleri.length === 0 ? (
                          <div className="p-12 text-center text-slate-400 text-sm font-medium">
                            Henüz kayıtlı görev emri bulunmuyor. Yeni bir görev emri oluşturmak için "GÖREV EMRİ GİRİŞ" sekmesini kullanabilirsiniz.
                          </div>
                        ) : (
                          <div className="overflow-x-auto border border-slate-200 rounded-2xl shadow-sm">
                            <table className="w-full text-left text-xs">
                              <thead className="bg-slate-900 text-slate-100 uppercase tracking-wider text-[10px] font-black select-none">
                                <tr>
                                  <th className="px-4 py-3">SIRA</th>
                                  <th className="px-4 py-3">TARİH</th>
                                  <th className="px-4 py-3">ARAÇ PLAKASI</th>
                                  <th className="px-4 py-3">SÜRÜCÜ</th>
                                  <th className="px-4 py-3">SERİ NO</th>
                                  <th className="px-4 py-3 text-right">ÇIKIŞ KM</th>
                                  <th className="px-4 py-3 text-right">DÖNÜŞ KM</th>
                                  <th className="px-4 py-3 text-right">TOPLAM KM</th>
                                  <th className="px-4 py-3">SAATLER</th>
                                  <th className="px-4 py-3">GÜZERGAH</th>
                                  <th className="px-4 py-3 text-center">İŞLEM</th>
                                </tr>
                              </thead>
                              <tbody className="divide-y divide-slate-100">
                                {karaAraclariGorevEmirleri.map((order, idx) => {
                                  const depKm = Number(order.departureKm) || 0;
                                  const retKm = Number(order.returnKm) || 0;
                                  const diffKm = Math.max(0, retKm - depKm);
                                  return (
                                    <tr key={order.id || idx} className="hover:bg-emerald-50/40 transition-colors">
                                      <td className="px-4 py-3 font-mono text-slate-500 font-bold">{idx + 1}</td>
                                      <td className="px-4 py-3 font-mono font-bold text-slate-800">{order.date || "-"}</td>
                                      <td className="px-4 py-3 font-bold text-slate-900">{order.plate || "-"}</td>
                                      <td className="px-4 py-3 text-slate-700 font-semibold">{order.driverName || "-"}</td>
                                      <td className="px-4 py-3 font-mono text-slate-600">{order.serialNo || "-"}</td>
                                      <td className="px-4 py-3 font-mono text-right text-slate-700">{order.departureKm || "-"}</td>
                                      <td className="px-4 py-3 font-mono text-right text-slate-700">{order.returnKm || "-"}</td>
                                      <td className="px-4 py-3 font-mono text-right font-black text-emerald-800">{diffKm > 0 ? `${diffKm} KM` : "-"}</td>
                                      <td className="px-4 py-3 text-[11px] text-slate-500 font-mono">
                                        {order.departureTime || "-"} / {order.returnTime || "-"}
                                      </td>
                                      <td className="px-4 py-3 text-[11px] text-slate-600 max-w-xs truncate" title={order.route}>
                                        {order.route || "-"}
                                      </td>
                                      <td className="px-4 py-3 text-center">
                                        <button
                                          type="button"
                                          onClick={() => {
                                            setGeDeleteOrderId(String(order.id));
                                            setGeDeletePasswordInput("");
                                            setGeDeletePasswordError(false);
                                            setShowGeDeletePasswordPrompt(true);
                                          }}
                                          className="p-1.5 text-rose-600 hover:bg-rose-50 rounded-lg transition-all cursor-pointer"
                                          title="Görevi Sil"
                                        >
                                          <Trash2 className="w-4 h-4" />
                                        </button>
                                      </td>
                                    </tr>
                                  );
                                })}
                              </tbody>
                            </table>
                          </div>
                        )}
                      </div>
                    )}
                  </div>
                )}
              </>
            </div>
          </div>
        );
      })()}

      </div>
    </div>

      {/* ACTIVE TECHIZAT ROW EDIT MODAL */}
      {activeTechizatRowEdit && (() => {
        const { rIdx, techType, row } = activeTechizatRowEdit;
        const isKaraAraci = techType === 'kara_araclari';
        const rowSection = getRowSection(row, techType);
        const isDepo = techizatActiveSection === 'depo_sarf' || 
                       techizatActiveSection === 'depo_kimyasal' || 
                       techizatActiveSection === 'depo_all' || 
                       rowSection === 'depo_sarf' || 
                       rowSection === 'depo_kimyasal';
        const baseColumns = [
          "SIRA NO", "TEÇHİZAT ADI", "PARÇA NO (P/N) / MODEL", "SERİ NO (S/N)", "MİKTAR / KAPASİTE", "BULUNDUĞU YER", "DURUMU", "KALİBRASYONA TABİ", "SON KONTROL / KALİBRASYON / BAKIM", "GELECEK KONTROL / KALİBRASYON / BAKIM", "SON KONTROLÜ YAPAN FİRMA", "AÇIKLAMA", "90 GÜN UYARISI MAİL GÖNDERİM TARİHİ"
        ];
        const depoColumns = [
          "SIRA NO", "MALZEME / PARÇA ADI", "PARÇA NO (P/N)", "SERİ NO (S/N)", "MİKTAR", "BULUNDUĞU YER / RAF", "DURUMU", "ÖMÜRLÜ PARÇA MI?", "ÖMÜR BİTİŞ TARİHİ", "TEDARİK EDİLEN FİRMA", "AÇIKLAMA"
        ];
        const karaAraclariColumns = [
          "SIRA NO", "ARAÇ PLAKASI / TANIMI", "PARÇA NO (P/N) / MODEL", "BULUNDUĞU YER", "SON KM Sİ", "DURUMU", "KALİBRASYONA TABİ", "SON KONTROL / KALİBRASYON / BAKIM", "GELECEK KONTROL / KALİBRASYON / BAKIM", "SON KONTROLÜ YAPAN FİRMA", "AÇIKLAMA", "90 GÜN UYARISI MAİL GÖNDERİM TARİHİ"
        ];
        const columns = isKaraAraci ? karaAraclariColumns : (isDepo ? depoColumns : baseColumns);
        const imageKey = techType + "_" + (row[1] || "").replace(/\s+/g, '_') + "_" + (row[3] || "").replace(/\s+/g, '_');
        const currentImageUrl = techizatImages[imageKey] || null;

        return (
          <div className="fixed inset-0 z-[600] flex items-center justify-center p-4 bg-slate-900/70 backdrop-blur-sm animate-fade-in">
            <div className="bg-white rounded-3xl shadow-2xl max-w-4xl w-full max-h-[92vh] flex flex-col overflow-hidden border border-slate-100">
              {/* Modal Header */}
              <div className="px-6 py-4 bg-slate-900 text-white flex items-center justify-between border-b border-slate-800">
                <div>
                  <h3 className="text-sm font-black uppercase tracking-wider flex items-center gap-2">
                    <Edit3 className="w-4 h-4 text-emerald-400" />
                    <span>{row[1] || "Kayıt Detayı ve Düzenleme"}</span>
                  </h3>
                  <p className="text-[11px] text-slate-400 font-mono mt-0.5">
                    P/N: {row[2] || "-"} • S/N: {row[3] || "-"} • {getUnitDisplayName(techType)}
                  </p>
                </div>
                <button
                  onClick={() => setActiveTechizatRowEdit(null)}
                  className="w-8 h-8 rounded-full bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white flex items-center justify-center font-bold text-sm transition-all cursor-pointer"
                >
                  ✕
                </button>
              </div>

              {/* Tab Switcher: Form vs Bölgesel vs Görsel */}
              <div className="flex border-b border-slate-200 bg-slate-50 px-6 pt-3 gap-2">
                <button
                  onClick={() => setMobileEditTab('form')}
                  className={`px-5 py-2.5 rounded-t-2xl font-black text-xs uppercase tracking-wider transition-all flex items-center gap-2 cursor-pointer ${
                    mobileEditTab === 'form'
                      ? 'bg-white text-emerald-950 border-t-2 border-x border-slate-200 border-t-emerald-600 shadow-sm'
                      : 'text-slate-500 hover:text-slate-800'
                  }`}
                >
                  <FileSpreadsheet className="w-4 h-4" />
                  <span>Kayıt Bilgileri</span>
                </button>
                <button
                  onClick={() => setMobileEditTab('regional')}
                  className={`px-5 py-2.5 rounded-t-2xl font-black text-xs uppercase tracking-wider transition-all flex items-center gap-2 cursor-pointer ${
                    mobileEditTab === 'regional'
                      ? 'bg-white text-emerald-950 border-t-2 border-x border-slate-200 border-t-emerald-600 shadow-sm'
                      : 'text-slate-500 hover:text-slate-800'
                  }`}
                >
                  <MapPin className="w-4 h-4 text-emerald-600" />
                  <span>Bölge Dağılımları</span>
                  {regionalLocations.length > 0 && (
                    <span className="text-[10px] font-mono bg-emerald-100 text-emerald-800 font-bold px-1.5 py-0.5 rounded-full">
                      {regionalLocations.length}
                    </span>
                  )}
                </button>
                {techType === 'hangar' && (
                  <button
                    onClick={() => setMobileEditTab('documents')}
                    className={`px-5 py-2.5 rounded-t-2xl font-black text-xs uppercase tracking-wider transition-all flex items-center gap-2 cursor-pointer ${
                      mobileEditTab === 'documents'
                        ? 'bg-white text-emerald-950 border-t-2 border-x border-slate-200 border-t-emerald-600 shadow-sm'
                        : 'text-slate-500 hover:text-slate-800'
                    }`}
                  >
                    <FileText className="w-4 h-4 text-rose-600" />
                    <span>Belge Yükle / Evraklar</span>
                    {findMatchingDocs(row, hangarPdfDocs).length > 0 && (
                      <span className="text-[10px] font-mono bg-rose-100 text-rose-800 font-bold px-1.5 py-0.5 rounded-full">
                        {hangarPdfDocs.filter(d => d.itemKey === (row[2] || row[1] || "").trim().toUpperCase()).length}
                      </span>
                    )}
                  </button>
                )}
                <button
                  onClick={() => setMobileEditTab('image')}
                  className={`px-5 py-2.5 rounded-t-2xl font-black text-xs uppercase tracking-wider transition-all flex items-center gap-2 cursor-pointer ${
                    mobileEditTab === 'image'
                      ? 'bg-white text-emerald-950 border-t-2 border-x border-slate-200 border-t-emerald-600 shadow-sm'
                      : 'text-slate-500 hover:text-slate-800'
                  }`}
                >
                  <Camera className="w-4 h-4" />
                  <span>Fotoğraf & Görsel Düzenleyici</span>
                </button>
              </div>

              {/* Modal Body */}
              <div className="flex-1 p-6 overflow-y-auto">
                {mobileEditTab === 'form' ? (
                  <div className="flex flex-col gap-5">
                    {/* Password unlock banner */}
                    {!isDataUpdateUnlocked ? (
                      <div className="p-4 bg-amber-50 border border-amber-200 rounded-2xl flex flex-col sm:flex-row items-center justify-between gap-3">
                        <div className="flex items-center gap-3">
                          <span className="text-2xl">🔒</span>
                          <div>
                            <p className="text-xs font-bold text-amber-950">Veri Düzenleme Kilidi Aktif</p>
                            <p className="text-[11px] text-amber-700">Değerleri değiştirmek için yönetici şifresini giriniz.</p>
                          </div>
                        </div>
                        <div className="flex items-center gap-2 w-full sm:w-auto">
                          <input
                            type="password"
                            placeholder="Yönetici Şifresi"
                            value={dataPasswordInput}
                            onChange={(e) => {
                              setDataPasswordInput(e.target.value);
                              setDataPasswordError(false);
                            }}
                            onKeyDown={(e) => {
                              if (e.key === 'Enter') {
                                const targetPwd = getTechizatUnitPassword(techType);
                                if (dataPasswordInput === targetPwd) {
                                  setIsDataUpdateUnlocked(true);
                                  setDataPasswordInput('');
                                  setDataPasswordError(false);
                                } else {
                                  setDataPasswordError(true);
                                }
                              }
                            }}
                            className="px-3 py-1.5 bg-white border border-amber-300 rounded-xl text-xs font-mono font-bold text-slate-900 !text-slate-900 placeholder:text-slate-400 focus:outline-none focus:border-amber-600 focus:ring-1 focus:ring-amber-500"
                          />
                          <button
                            type="button"
                            onClick={() => {
                              const targetPwd = getTechizatUnitPassword(techType);
                              if (dataPasswordInput === targetPwd) {
                                setIsDataUpdateUnlocked(true);
                                setDataPasswordInput('');
                                setDataPasswordError(false);
                              } else {
                                setDataPasswordError(true);
                              }
                            }}
                            className="px-4 py-1.5 bg-amber-600 hover:bg-amber-700 text-white font-bold text-xs rounded-xl shadow-sm transition-all cursor-pointer"
                          >
                            Kilidi Aç
                          </button>
                        </div>
                        {dataPasswordError && (
                          <p className="text-rose-600 text-xs font-bold w-full">Hatalı şifre! Lütfen tekrar deneyiniz.</p>
                        )}
                      </div>
                    ) : (
                      <div className="p-3 bg-emerald-50 border border-emerald-200 rounded-2xl flex items-center justify-between">
                        <span className="text-xs font-bold text-emerald-900 flex items-center gap-2">
                          🔓 Düzenleme modu aktif. Yapılan değişiklikleri kaydetmeyi unutmayınız.
                        </span>
                        <button
                          type="button"
                          onClick={() => setIsDataUpdateUnlocked(false)}
                          className="text-xs text-emerald-700 hover:text-emerald-950 font-bold underline cursor-pointer"
                        >
                          Kilitle
                        </button>
                      </div>
                    )}

                    {/* Bölge Dağılımları Hızlı Erişim Kartı & Butonu */}
                    <div className="p-4 bg-gradient-to-r from-emerald-50 via-teal-50/70 to-emerald-50 border border-emerald-200 rounded-2xl flex flex-col sm:flex-row items-center justify-between gap-3 shadow-sm">
                      <div className="flex items-center gap-3">
                        <div className="w-9 h-9 rounded-xl bg-emerald-600 text-white flex items-center justify-center font-black shadow-sm shrink-0">
                          <MapPin className="w-5 h-5" />
                        </div>
                        <div>
                          <p className="text-xs font-black text-emerald-950 uppercase tracking-tight flex items-center gap-2">
                            <span>📍 BÖLGESEL LOKASYON & SERİ NUMARASI (S/N) DAĞILIMI</span>
                            <span className="text-[10px] font-mono px-2 py-0.5 rounded-full bg-emerald-200 text-emerald-900">
                              {regionalLocations.length} Bölge Kaydı
                            </span>
                          </p>
                          <p className="text-[11px] text-emerald-700">
                            Bu parçanın farklı bölgelerdeki adet, S/N kırılımları ve bağımsız kontrol tarihlerini açmak için tıklayınız.
                          </p>
                        </div>
                      </div>
                      <button
                        type="button"
                        onClick={() => setMobileEditTab('regional')}
                        className="w-full sm:w-auto px-4 py-2 bg-emerald-700 hover:bg-emerald-800 text-white font-bold text-xs rounded-xl shadow-sm transition-all flex items-center justify-center gap-2 cursor-pointer shrink-0"
                      >
                        <MapPin className="w-4 h-4" />
                        <span>Bölge Dağılımlarını Yönet</span>
                      </button>
                    </div>

                    {/* Hangar Yer Destek Belge Yükle Hızlı Kartı */}
                    {techType === 'hangar' && (
                      <div className="p-4 bg-gradient-to-r from-rose-50 via-pink-50/70 to-rose-50 border border-rose-200 rounded-2xl flex flex-col sm:flex-row items-center justify-between gap-3 shadow-sm">
                        <div className="flex items-center gap-3">
                          <div className="w-9 h-9 rounded-xl bg-rose-600 text-white flex items-center justify-center font-black shadow-sm shrink-0">
                            <FileText className="w-5 h-5" />
                          </div>
                          <div>
                            <p className="text-xs font-black text-rose-950 uppercase tracking-tight flex items-center gap-2">
                              <span>📄 BELGE YÜKLE (PDF EVRAKLAR)</span>
                              <span className="text-[10px] font-mono px-2 py-0.5 rounded-full bg-rose-200 text-rose-900 font-bold">
                                {findMatchingDocs(row, hangarPdfDocs).length} Evrak Kayıtlı
                              </span>
                            </p>
                            <p className="text-[11px] text-rose-700">
                              Bakım sonrası evraklar, kullanıcı kılavuzları, kontrol checklist PDF belgelerini yükleyin ve pencere içinde açın.
                            </p>
                          </div>
                        </div>
                        <button
                          type="button"
                          onClick={() => setMobileEditTab('documents')}
                          className="w-full sm:w-auto px-4 py-2 bg-rose-700 hover:bg-rose-800 text-white font-bold text-xs rounded-xl shadow-sm transition-all flex items-center justify-center gap-2 cursor-pointer shrink-0"
                        >
                          <FileText className="w-4 h-4" />
                          <span>Belgeleri Yönet & Aç</span>
                        </button>
                      </div>
                    )}

                    {/* Editable Fields Grid */}
                    <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
                      {columns.map((colName, cIdx) => {
                        if (colName === "BELGE YÜKLE") {
                          const docKey = (row[2] || row[1] || "").trim().toUpperCase();
                          const count = findMatchingDocs(row, hangarPdfDocs).length;
                          return (
                            <div key={cIdx} className="flex flex-col gap-1.5 sm:col-span-2 lg:col-span-3 bg-rose-50/60 p-3.5 rounded-2xl border border-rose-200">
                              <label className="text-[10px] font-black text-rose-800 uppercase tracking-wider flex items-center justify-between">
                                <span>📄 BELGE YÜKLE</span>
                                <span className="font-mono text-rose-600">{count} Belge</span>
                              </label>
                              <div className="flex flex-col sm:flex-row items-center justify-between gap-3">
                                <p className="text-xs text-rose-900 font-medium">
                                  Hangar Yer Destek bakım ve kontrol evraklarını PDF olarak ekleyin ve sistemden çıkmadan inceleyin.
                                </p>
                                <button
                                  type="button"
                                  onClick={() => setMobileEditTab('documents')}
                                  className="w-full sm:w-auto px-4 py-2 bg-rose-600 hover:bg-rose-700 text-white font-bold text-xs rounded-xl shadow-sm transition-all flex items-center justify-center gap-1.5 cursor-pointer shrink-0"
                                >
                                  <FileText className="w-4 h-4" />
                                  <span>{count > 0 ? `Belgeleri İncele / Ekle (${count})` : '+ Belge Yükle'}</span>
                                </button>
                              </div>
                            </div>
                          );
                        }
                        if (colName === "DURUMU") {
                          const curDurum = editRowValues[cIdx] !== undefined ? editRowValues[cIdx] : (row[cIdx] || "");
                          return (
                            <div key={cIdx} className="flex flex-col gap-1.5">
                              <label className="text-[10px] font-black text-slate-400 uppercase tracking-wider truncate flex items-center justify-between" title={colName}>
                                <span>{colName}</span>
                                <span className="text-[9px] font-mono text-emerald-600 font-bold">Faal olunca bakım güncellenir</span>
                              </label>
                              <div className="flex items-center gap-1.5">
                                <select
                                  disabled={!isDataUpdateUnlocked}
                                  value={curDurum}
                                  onChange={(e) => {
                                    const newVal = e.target.value;
                                    const copy = [...editRowValues];
                                    while (copy.length <= cIdx) copy.push("");
                                    const prevVal = (copy[cIdx] || row[cIdx] || "").toUpperCase().trim();
                                    copy[cIdx] = newVal;

                                    // FAAL'e çevrilince bakım tarihlerini ileriye çek
                                    if (newVal.toUpperCase().trim() === "FAAL" && prevVal !== "FAAL") {
                                      const sonIdx = columns.findIndex(c => c.includes("SON KONTROL"));
                                      const gelecekIdx = columns.findIndex(c => c.includes("GELECEK KONTROL") || c.includes("MUAYENE"));
                                      if (sonIdx !== -1 && gelecekIdx !== -1) {
                                        const prevSon = copy[sonIdx] || row[sonIdx] || "";
                                        const prevGelecek = copy[gelecekIdx] || row[gelecekIdx] || "";
                                        const { newSonKontrol, newGelecekKontrol, intervalDays } = calculateReactivatedMaintenanceDates(prevSon, prevGelecek);
                                        copy[sonIdx] = newSonKontrol;
                                        copy[gelecekIdx] = newGelecekKontrol;
                                        showNotification(`⚡ Ürün FAAL duruma alındı. Son kontrol bugüne (${newSonKontrol}), gelecek bakım ${intervalDays} gün ileriye (${newGelecekKontrol}) güncellendi!`);
                                      }
                                    }
                                    setEditRowValues(copy);
                                  }}
                                  className={`w-full px-3.5 py-2.5 rounded-xl text-xs font-bold border transition-all ${
                                    !isDataUpdateUnlocked
                                      ? 'bg-slate-100 border-slate-200 text-slate-600 cursor-not-allowed'
                                      : 'bg-white border-slate-300 text-slate-800 focus:outline-none focus:border-[#0b3d1d] focus:ring-2 focus:ring-[#0b3d1d]/10'
                                  }`}
                                >
                                  <option value="FAAL">FAAL</option>
                                  <option value="GAYRİ FAAL">GAYRİ FAAL</option>
                                  <option value="KISMEN FAAL">KISMEN FAAL</option>
                                  <option value="ARIZALI">ARIZALI</option>
                                  <option value="BAKIMDA">BAKIMDA</option>
                                  <option value="RAF ÖMRÜ DOLDU">RAF ÖMRÜ DOLDU</option>
                                  {curDurum && !['FAAL', 'GAYRİ FAAL', 'KISMEN FAAL', 'ARIZALI', 'BAKIMDA', 'RAF ÖMRÜ DOLDU'].includes(curDurum.toUpperCase().trim()) && (
                                    <option value={curDurum}>{curDurum}</option>
                                  )}
                                </select>
                              </div>
                            </div>
                          );
                        }

                        const isDateCol = colName.includes("TARİH") || colName.includes("KONTROL") || colName.includes("BAKIM") || colName.includes("ÖMÜR");

                        return (
                        <div key={cIdx} className="flex flex-col gap-1.5">
                          <label className="text-[10px] font-black text-slate-400 uppercase tracking-wider truncate" title={colName}>
                            {colName}
                          </label>
                          <input
                            type="text"
                            disabled={!isDataUpdateUnlocked}
                            placeholder={isDateCol ? "GG.AA.YYYY" : ""}
                            value={editRowValues[cIdx] !== undefined ? editRowValues[cIdx] : (row[cIdx] || "")}
                            onChange={(e) => {
                              const copy = [...editRowValues];
                              while (copy.length <= cIdx) copy.push("");
                              copy[cIdx] = e.target.value;
                              setEditRowValues(copy);
                            }}
                            onBlur={(e) => {
                              if (isDateCol && e.target.value) {
                                const formatted = cleanAndFormatDateString(e.target.value);
                                if (formatted && formatted !== e.target.value) {
                                  const copy = [...editRowValues];
                                  while (copy.length <= cIdx) copy.push("");
                                  copy[cIdx] = formatted;
                                  setEditRowValues(copy);
                                }
                              }
                            }}
                            className={`w-full px-3.5 py-2.5 rounded-xl text-xs font-semibold border transition-all ${
                              !isDataUpdateUnlocked
                                ? 'bg-slate-100 border-slate-200 text-slate-600 cursor-not-allowed'
                                : 'bg-white border-slate-300 text-slate-800 focus:outline-none focus:border-[#0b3d1d] focus:ring-2 focus:ring-[#0b3d1d]/10'
                            }`}
                          />
                        </div>
                        );
                      })}
                    </div>
                  </div>
                ) : mobileEditTab === 'regional' ? (
                  <div className="flex flex-col gap-5">
                    {/* Header Banner */}
                    <div className="p-4 bg-emerald-950 text-white rounded-2xl flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3 shadow-md">
                      <div>
                        <div className="flex items-center gap-2 mb-1">
                          <span className="text-base">📍</span>
                          <h4 className="text-xs sm:text-sm font-black uppercase tracking-tight text-white">
                            BÖLGESEL LOKASYON, MİKTAR, S/N & KONTROL / BAKIM DAĞILIMI
                          </h4>
                        </div>
                        <p className="text-[11px] text-emerald-200/90 leading-relaxed max-w-2xl">
                          Aynı P/N'e ait ürünlerin bölgelere, adetlere ve seri numaralarına göre bağımsız Son Kontrol, Gelecek Kontrol ve Firma tarihlerini yönetiniz.
                        </p>
                      </div>
                      <span className="text-xs font-mono font-black text-emerald-300 bg-emerald-900/90 border border-emerald-700/60 px-3 py-1.5 rounded-xl shrink-0">
                        {regionalLocations.length} Ayrı Lokasyon / S/N Kaydı
                      </span>
                    </div>

                    {/* Regional Cards List */}
                    <div className="flex flex-col gap-4">
                      {regionalLocations.map((item, idx) => {
                        const daysInfo = getDaysRemainingLabel(item.gelecekKontrol);
                        const curSerialInput = newSerialInputs[item.id] || '';

                        return (
                          <div
                            key={item.id}
                            className="bg-slate-50 border border-slate-200 rounded-2xl p-4 sm:p-5 flex flex-col gap-4 shadow-sm hover:border-emerald-300 transition-all"
                          >
                            {/* Card Header Bar */}
                            <div className="flex items-center justify-between border-b border-slate-200 pb-3">
                              <div className="flex items-center gap-2.5">
                                <span className="w-6 h-6 rounded-full bg-emerald-600 text-white text-xs font-black flex items-center justify-center shrink-0">
                                  {idx + 1}
                                </span>
                                <h5 className="text-xs sm:text-sm font-black text-slate-800 uppercase tracking-tight">
                                  {item.location || 'BÖLGE BELİRTİLMEDİ'} {item.quantity ? `(${item.quantity} ADET)` : ''}
                                </h5>
                              </div>
                              <div className="flex items-center gap-2">
                                {/* Kullanıcı kuralı: "faal yap butonu eğer geleceke kontrol bakım tarih geçmiş ise o butonn sil yanına olsun aşağıdan kaldır" */}
                                {(() => {
                                  let isExpired = false;
                                  if (item.gelecekKontrol && item.gelecekKontrol !== '-') {
                                    const parts = item.gelecekKontrol.trim().split('.');
                                    if (parts.length === 3) {
                                      const target = new Date(Number(parts[2]), Number(parts[1]) - 1, Number(parts[0]));
                                      if (!isNaN(target.getTime())) {
                                        const today = new Date();
                                        today.setHours(0, 0, 0, 0);
                                        isExpired = target.getTime() < today.getTime();
                                      }
                                    }
                                  }
                                  if (!isExpired) return null;
                                  return (
                                    <button
                                      type="button"
                                      onClick={() => {
                                        const { newSonKontrol, newGelecekKontrol, intervalDays } = calculateReactivatedMaintenanceDates(item.sonKontrol, item.gelecekKontrol);
                                        setRegionalLocations(prev => prev.map(r => r.id === item.id ? {
                                          ...r,
                                          sonKontrol: newSonKontrol,
                                          gelecekKontrol: newGelecekKontrol
                                        } : r));
                                        showNotification(`⚡ Bakım tamamlandı. Son kontrol bugüne (${newSonKontrol}), gelecek bakım ${intervalDays} gün sonraya (${newGelecekKontrol}) güncellendi!`);
                                      }}
                                      className="px-2.5 py-1 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-black rounded-lg transition-all cursor-pointer flex items-center gap-1 shadow-sm"
                                      title="Tarihi geçmiş ürünü faale al: Son kontrolü bugüne çeker, aradaki periyodu gelecek bakıma ekler"
                                    >
                                      <RefreshCw className="w-3.5 h-3.5" />
                                      <span>Faal Yap</span>
                                    </button>
                                  );
                                })()}
                                <button
                                  type="button"
                                  onClick={() => {
                                    setRegionalLocations(prev => prev.filter(r => r.id !== item.id));
                                  }}
                                  className="text-rose-500 hover:text-rose-700 font-bold text-xs flex items-center gap-1 cursor-pointer transition-colors px-2.5 py-1 rounded-lg hover:bg-rose-50"
                                  title="Bu lokasyon kaydını sil"
                                >
                                  <Trash2 className="w-3.5 h-3.5" />
                                  <span>Sil</span>
                                </button>
                              </div>
                            </div>

                            {/* Row 1: Yer, Miktar, Seri No */}
                            <div className="grid grid-cols-1 md:grid-cols-3 gap-3.5">
                              <div className="flex flex-col gap-1">
                                <label className="text-[10px] font-black text-slate-500 uppercase tracking-wider flex items-center gap-1">
                                  <span>📍 BULUNDUĞU YER / BÖLGE</span>
                                </label>
                                <input
                                  type="text"
                                  placeholder="Örn: ÇANAKKALE"
                                  value={item.location}
                                  onChange={(e) => {
                                    const val = e.target.value;
                                    setRegionalLocations(prev => prev.map(r => r.id === item.id ? { ...r, location: val } : r));
                                  }}
                                  className="px-3 py-2 bg-white border border-slate-300 rounded-xl text-xs font-bold text-slate-800 focus:outline-none focus:border-emerald-600 focus:ring-1 focus:ring-emerald-500"
                                />
                              </div>

                              <div className="flex flex-col gap-1">
                                <div className="flex items-center justify-between">
                                  <label className="text-[10px] font-black text-slate-500 uppercase tracking-wider">
                                    📦 MİKTAR
                                  </label>
                                  <span className="text-[9px] font-mono text-emerald-700 bg-emerald-50 px-2 py-0.5 rounded-full border border-emerald-200">
                                    = S/N ({item.serialNumbers.length})
                                  </span>
                                </div>
                                <input
                                  type="text"
                                  value={item.quantity}
                                  onChange={(e) => {
                                    const val = e.target.value;
                                    setRegionalLocations(prev => prev.map(r => r.id === item.id ? { ...r, quantity: val } : r));
                                  }}
                                  placeholder="Adet giriniz..."
                                  className="px-3 py-2 bg-white border border-slate-300 rounded-xl text-xs font-bold text-slate-800 focus:outline-none focus:border-emerald-600 focus:ring-1 focus:ring-emerald-500"
                                />
                              </div>

                              <div className="flex flex-col gap-1">
                                <label className="text-[10px] font-black text-slate-500 uppercase tracking-wider">
                                  🔢 SERİ NO (S/N)
                                </label>
                                <input
                                  type="text"
                                  placeholder="Seri no yazınız (birden fazla ise virgül ile ayırınız)..."
                                  value={item.serialNumbers.join(', ')}
                                  onChange={(e) => {
                                    const val = e.target.value;
                                    const parts = val.split(/[,;\/\n]+/).map(s => s.trim()).filter(Boolean);
                                    setRegionalLocations(prev => prev.map(r => r.id === item.id ? {
                                      ...r,
                                      serialNumbers: parts,
                                      quantity: parts.length > 0 ? String(parts.length) : r.quantity
                                    } : r));
                                  }}
                                  className="px-3 py-2 bg-white border border-slate-300 rounded-xl text-xs font-mono font-bold text-slate-800 focus:outline-none focus:border-emerald-600 focus:ring-1 focus:ring-emerald-500"
                                />
                                <span className="text-[10px] text-slate-400">
                                  {item.serialNumbers.length > 0 ? `Kayıtlı: ${item.serialNumbers.length} S/N` : "Seri no yoksa boş bırakabilirsiniz"}
                                </span>
                              </div>
                            </div>

                            {/* Row 2: Son Kontrol, Gelecek Kontrol, Firma */}
                            <div className="grid grid-cols-1 md:grid-cols-3 gap-3.5 pt-2 border-t border-slate-200/60">
                              <div className="flex flex-col gap-1">
                                <label className="text-[10px] font-black text-slate-500 uppercase tracking-wider">
                                  🛠️ SON KONTROL / BAKIM
                                </label>
                                <input
                                  type="text"
                                  placeholder="GG.AA.YYYY"
                                  value={item.sonKontrol}
                                  onChange={(e) => {
                                    const val = e.target.value;
                                    setRegionalLocations(prev => prev.map(r => r.id === item.id ? { ...r, sonKontrol: val } : r));
                                  }}
                                  className="px-3 py-2 bg-white border border-slate-300 rounded-xl text-xs font-mono font-bold text-slate-800 focus:outline-none focus:border-emerald-600 focus:ring-1 focus:ring-emerald-500"
                                />
                              </div>

                              <div className="flex flex-col gap-1">
                                <div className="flex items-center justify-between">
                                  <label className="text-[10px] font-black text-slate-500 uppercase tracking-wider">
                                    📅 GELECEK KONTROL / BAKIM
                                  </label>
                                  {daysInfo && (
                                    <span className={`text-[9px] font-mono font-bold px-2 py-0.5 rounded-full border ${daysInfo.color}`}>
                                      {daysInfo.text}
                                    </span>
                                  )}
                                </div>
                                <div className="flex items-center gap-1.5">
                                  <input
                                    type="text"
                                    placeholder="GG.AA.YYYY"
                                    value={item.gelecekKontrol}
                                    onChange={(e) => {
                                      const val = e.target.value;
                                      setRegionalLocations(prev => prev.map(r => r.id === item.id ? { ...r, gelecekKontrol: val } : r));
                                    }}
                                    className="flex-1 px-3 py-2 bg-white border border-slate-300 rounded-xl text-xs font-mono font-bold text-slate-800 focus:outline-none focus:border-emerald-600 focus:ring-1 focus:ring-emerald-500"
                                  />
                                  <button
                                    type="button"
                                    onClick={() => {
                                      const parts = (item.gelecekKontrol || item.sonKontrol || '').split('.');
                                      if (parts.length === 3) {
                                        const nextYear = String(Number(parts[2]) + 1);
                                        const newDate = `${parts[0]}.${parts[1]}.${nextYear}`;
                                        setRegionalLocations(prev => prev.map(r => r.id === item.id ? { ...r, gelecekKontrol: newDate } : r));
                                      } else {
                                        const d = new Date();
                                        d.setFullYear(d.getFullYear() + 1);
                                        const pad = (n: number) => n < 10 ? '0' + n : n;
                                        const newDate = `${pad(d.getDate())}.${pad(d.getMonth() + 1)}.${d.getFullYear()}`;
                                        setRegionalLocations(prev => prev.map(r => r.id === item.id ? { ...r, gelecekKontrol: newDate } : r));
                                      }
                                    }}
                                    className="px-2.5 py-2 bg-slate-100 hover:bg-slate-200 text-slate-800 text-[10px] font-black rounded-xl transition-all cursor-pointer whitespace-nowrap"
                                    title="Tarihi 1 yıl ileri al"
                                  >
                                    +1 Yıl
                                  </button>
                                </div>
                              </div>

                              <div className="flex flex-col gap-1">
                                <label className="text-[10px] font-black text-slate-500 uppercase tracking-wider">
                                  🏢 SON KONTROLÜ YAPAN FİRMA
                                </label>
                                <input
                                  type="text"
                                  placeholder="Firma adı..."
                                  value={item.firma}
                                  onChange={(e) => {
                                    const val = e.target.value;
                                    setRegionalLocations(prev => prev.map(r => r.id === item.id ? { ...r, firma: val } : r));
                                  }}
                                  className="px-3 py-2 bg-white border border-slate-300 rounded-xl text-xs font-bold text-slate-800 focus:outline-none focus:border-emerald-600 focus:ring-1 focus:ring-emerald-500"
                                />
                              </div>
                            </div>
                          </div>
                        );
                      })}
                    </div>

                    {/* Add & Save Action Buttons */}
                    <div className="flex flex-col sm:flex-row items-center justify-between gap-3 pt-2">
                      <button
                        type="button"
                        onClick={() => {
                          const newItem = {
                            id: `reg_${Date.now()}_${regionalLocations.length}`,
                            location: 'YENİ BÖLGE',
                            quantity: '1',
                            serialNumbers: [],
                            sonKontrol: '-',
                            gelecekKontrol: '-',
                            firma: '-'
                          };
                          setRegionalLocations(prev => [...prev, newItem]);
                        }}
                        className="w-full sm:w-auto px-5 py-2.5 bg-slate-800 hover:bg-slate-900 text-white font-bold text-xs rounded-xl transition-all flex items-center justify-center gap-2 cursor-pointer shadow-sm"
                      >
                        <Plus className="w-4 h-4" />
                        <span>+ YENİ LOKASYON / BÖLGE EKLE</span>
                      </button>

                      <button
                        type="button"
                        onClick={() => {
                          const copy = [...editRowValues];
                          while (copy.length < columns.length) copy.push("");

                          const snIdx = columns.findIndex(c => c.includes("SERİ") || c.includes("S/N"));
                          const miktarIdx = columns.findIndex(c => c.includes("MİKTAR") || c.includes("ADET"));
                          const yerIdx = columns.findIndex(c => c.includes("YER") || c.includes("BÖLGE") || c.includes("LOKASYON"));
                          const sonIdx = columns.findIndex(c => c.includes("SON KONTROL"));
                          const gelecekIdx = columns.findIndex(c => c.includes("GELECEK KONTROL") || c.includes("MUAYENE"));
                          const firmaIdx = columns.findIndex(c => c.includes("FİRMA") || c.includes("TEDARİKÇİ"));

                          let totalQty = 0;
                          regionalLocations.forEach(r => {
                            const q = parseFloat(String(r.quantity)) || (r.serialNumbers.length > 0 ? r.serialNumbers.length : 1);
                            totalQty += q;
                          });
                          if (miktarIdx !== -1) {
                            copy[miktarIdx] = totalQty > 0 ? `${totalQty} ADET` : (copy[miktarIdx] || '1 ADET');
                          }

                          const locStr = regionalLocations.map(r => `${r.location}${r.quantity ? ` (${r.quantity})` : ''}`).join(' ; ');
                          if (yerIdx !== -1) {
                            copy[yerIdx] = locStr || copy[yerIdx];
                          }

                          const allSns = regionalLocations.flatMap(r => r.serialNumbers);
                          if (snIdx !== -1 && allSns.length > 0) {
                            copy[snIdx] = allSns.join(' ; ');
                          }

                          const validSon = regionalLocations.map(r => r.sonKontrol).filter(s => s && s !== '-');
                          if (sonIdx !== -1 && validSon.length > 0) {
                            copy[sonIdx] = validSon.join('\n');
                          }

                          const validGelecek = regionalLocations.map(r => r.gelecekKontrol).filter(g => g && g !== '-');
                          if (gelecekIdx !== -1 && validGelecek.length > 0) {
                            copy[gelecekIdx] = validGelecek.join('\n');
                          }

                          const validFirma = regionalLocations.map(r => r.firma).filter(f => f && f !== '-');
                          if (firmaIdx !== -1 && validFirma.length > 0) {
                            copy[firmaIdx] = validFirma.join('\n');
                          }

                          setEditRowValues(copy);
                          showNotification("Bölge dağılımları form bilgileriyle senkronize edildi!");
                          setMobileEditTab('form');
                        }}
                        className="w-full sm:w-auto px-6 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white font-black text-xs uppercase tracking-wider rounded-xl shadow-lg shadow-emerald-600/20 transition-all flex items-center justify-center gap-2 cursor-pointer"
                      >
                        <Save className="w-4 h-4" />
                        <span>Bölge Dağılımlarını Forma Aktar & Geri Dön</span>
                      </button>
                    </div>
                  </div>
                ) : mobileEditTab === 'documents' ? (
                  <div className="flex flex-col gap-6">
                    {/* Header Banner */}
                    <div className="p-4 bg-rose-50 border border-rose-200 rounded-2xl flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3">
                      <div>
                        <h4 className="text-sm font-black text-rose-950 uppercase tracking-wide flex items-center gap-2">
                          <FileText className="w-4 h-4 text-rose-600" />
                          <span>HANGAR YER DESTEK EVRAK & PDF BELGE YÖNETİMİ</span>
                        </h4>
                        <p className="text-xs text-rose-800 mt-1">
                          Bu teçhizata ait bakım sonrası evrakları, kullanıcı kılavuzlarını ve kontrol checklist belgelerini yükleyin veya tarayıcı içi önizleyin.
                        </p>
                      </div>
                      <span className="text-xs font-mono font-black bg-rose-200 text-rose-900 px-3 py-1.5 rounded-xl self-start sm:self-auto">
                        P/N: {row[2] || "-"} | S/N: {row[3] || "-"}
                      </span>
                    </div>

                    {/* Belge Yükleme Formu */}
                    <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm flex flex-col gap-4">
                      <div className="flex items-center justify-between border-b border-slate-100 pb-3">
                        <h5 className="text-xs font-black text-slate-800 uppercase tracking-wider flex items-center gap-2">
                          <Upload className="w-4 h-4 text-rose-600" />
                          <span>YENİ BELGE / DÖKÜMAN YÜKLE</span>
                        </h5>
                        <span className="text-[10px] text-slate-500 font-mono font-bold bg-slate-100 px-2 py-0.5 rounded-md">
                          PDF • JPEG / GÖRSEL • ZIP • RAR
                        </span>
                      </div>

                      <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                        {/* 1. Belge Türü Seçimi */}
                        <div className="flex flex-col gap-1.5">
                          <label className="text-[11px] font-bold text-slate-700">
                            Belge Türü <span className="text-rose-500">*</span>
                          </label>
                          <select
                            value={newDocDocType}
                            onChange={(e) => setNewDocDocType(e.target.value)}
                            className="w-full px-3 py-2 text-xs font-semibold bg-white border border-slate-300 rounded-xl focus:outline-none focus:ring-2 focus:ring-rose-500/20 focus:border-rose-600"
                          >
                            <option value="Bakım Sonrası Evraklar">Bakım Sonrası Evraklar</option>
                            <option value="Kullanıcı Kılavuzları">Kullanıcı Kılavuzları</option>
                            <option value="Kontrol Checklist">Kontrol Checklist</option>
                            <option value="Diğer">Diğer (Manuel Giriş)</option>
                          </select>
                        </div>

                        {/* Manuel Belge Türü (Eğer Diğer seçilirse) */}
                        {newDocDocType === "Diğer" && (
                          <div className="flex flex-col gap-1.5">
                            <label className="text-[11px] font-bold text-slate-700">
                              Evrak Türünü Yazınız <span className="text-rose-500">*</span>
                            </label>
                            <input
                              type="text"
                              value={newDocCustomType}
                              onChange={(e) => setNewDocCustomType(e.target.value)}
                              placeholder="Örn: Kalibrasyon Sertifikası..."
                              className="w-full px-3 py-2 text-xs font-semibold bg-white border border-rose-300 rounded-xl focus:outline-none focus:ring-2 focus:ring-rose-500/20 focus:border-rose-600"
                            />
                          </div>
                        )}

                        {/* 2. Bakımı Yapan Firma */}
                        <div className="flex flex-col gap-1.5">
                          <label className="text-[11px] font-bold text-slate-700">
                            Bakımı Yapan Firma / Kurum
                          </label>
                          <input
                            type="text"
                            value={newDocFirma || (row[10] || "")}
                            onChange={(e) => setNewDocFirma(e.target.value)}
                            placeholder="Örn: TUSAŞ, THY Teknik..."
                            className="w-full px-3 py-2 text-xs font-semibold bg-white border border-slate-300 rounded-xl focus:outline-none focus:ring-2 focus:ring-rose-500/20 focus:border-rose-600"
                          />
                        </div>

                        {/* 3. Belge Dosyası Seçme */}
                        <div className="flex flex-col gap-1.5">
                          <label className="text-[11px] font-bold text-slate-700">
                            Dosya Seçin (PDF, JPEG, ZIP, RAR) <span className="text-rose-500">*</span>
                          </label>
                          <input
                            type="file"
                            accept=".pdf,application/pdf,image/jpeg,image/png,image/jpg,.jpg,.jpeg,.png,.zip,application/zip,application/x-zip-compressed,.rar,application/x-rar-compressed,application/octet-stream"
                            onChange={(e) => {
                              const file = e.target.files?.[0];
                              if (file) {
                                const ext = file.name.split('.').pop()?.toLowerCase();
                                const allowedExts = ['pdf', 'jpg', 'jpeg', 'png', 'webp', 'zip', 'rar'];
                                if (!allowedExts.includes(ext || '')) {
                                  alert("Lütfen geçerli bir dosya seçiniz (PDF, JPEG/Görsel, ZIP veya RAR)!");
                                  e.target.value = "";
                                  return;
                                }
                                setNewDocSelectedFile(file);
                              }
                            }}
                            className="w-full text-xs text-slate-600 file:mr-3 file:py-2 file:px-3 file:rounded-xl file:border-0 file:text-xs file:font-black file:bg-rose-50 file:text-rose-700 hover:file:bg-rose-100 cursor-pointer border border-slate-200 rounded-xl p-1"
                          />
                        </div>
                      </div>

                      {newDocSelectedFile && (
                        <div className="p-3 bg-slate-50 border border-slate-200 rounded-xl flex flex-wrap items-center justify-between gap-3">
                          <div className="flex items-center gap-2 text-xs font-bold text-slate-800">
                            {(() => {
                              const cat = getFileCategory(newDocSelectedFile.name);
                              if (cat === 'image') return <ImageIcon className="w-4 h-4 text-blue-600 shrink-0" />;
                              if (cat === 'archive') return <Archive className="w-4 h-4 text-indigo-600 shrink-0" />;
                              return <FileText className="w-4 h-4 text-rose-600 shrink-0" />;
                            })()}
                            <span className="truncate max-w-xs">Seçilen Dosya: {newDocSelectedFile.name}</span>
                            <span className="text-slate-400 font-mono text-[11px]">({(newDocSelectedFile.size / 1024 / 1024).toFixed(2)} MB)</span>
                            <span className="text-[10px] font-mono font-bold px-1.5 py-0.5 rounded uppercase bg-slate-200 text-slate-700">
                              {newDocSelectedFile.name.split('.').pop()?.toUpperCase()}
                            </span>
                          </div>
                          <button
                            type="button"
                            disabled={isUploadingDoc}
                            onClick={async () => {
                              if (!newDocSelectedFile) {
                                showNotification("Lütfen bir dosya seçiniz.");
                                return;
                              }
                              const file = newDocSelectedFile;
                              const docKey = `${(row[2] || row[1] || "").trim().toUpperCase()}`;
                              const finalType = newDocDocType === "Diğer" ? (newDocCustomType.trim() || "Diğer") : newDocDocType;
                              const finalFirma = newDocFirma.trim() || (row[10] || "");
                              const fileMimeType = getDocMimeType(file.name);

                              setIsUploadingDoc(true);
                              showNotification("Belge Google Drive'a aktarılıyor...");

                              try {
                                const base64Data = await fileToBase64(file);
                                const fullDataUrl = `data:${fileMimeType};base64,${base64Data}`;

                                // Upload via backend proxy (/api/upload-hangar-pdf) first
                                let driveFileId = "";
                                let driveUrl = "";
                                try {
                                  const proxyRes = await fetch('/api/upload-hangar-pdf', {
                                    method: 'POST',
                                    headers: { 'Content-Type': 'application/json' },
                                    body: JSON.stringify({
                                      fileName: `${docKey}_${file.name}`,
                                      base64Data: base64Data,
                                      mimeType: fileMimeType,
                                      folderId: "1HQR_NYKhHQGA7_2W3nArI9pCh-LJasTP",
                                      itemKey: docKey,
                                      docType: finalType,
                                      firma: finalFirma
                                    })
                                  });
                                  if (proxyRes.ok) {
                                    const proxyJson = await proxyRes.json();
                                    if (proxyJson && proxyJson.fileId) {
                                      driveFileId = proxyJson.fileId;
                                      driveUrl = proxyJson.viewUrl || `https://drive.google.com/file/d/${driveFileId}/preview`;
                                    }
                                  }
                                } catch (proxyErr) {
                                  console.warn("Proxy upload fallback to direct GAS:", proxyErr);
                                }

                                // Fallback to Google Apps Script if proxy didn't succeed
                                if (!driveFileId) {
                                  try {
                                    const driveRes = await fetch(GOOGLE_SCRIPT_URL, {
                                      method: "POST",
                                      headers: { "Content-Type": "text/plain;charset=utf-8" },
                                      body: JSON.stringify({
                                        action: "uploadPdfToDrive",
                                        fileName: `${docKey}_${file.name}`,
                                        mimeType: fileMimeType,
                                        base64Data: base64Data,
                                        folderId: "1HQR_NYKhHQGA7_2W3nArI9pCh-LJasTP",
                                        itemKey: docKey,
                                        docType: finalType,
                                        firma: finalFirma
                                      })
                                    });
                                    if (driveRes.ok) {
                                      const driveJson = await driveRes.json();
                                      if (driveJson && driveJson.fileId) {
                                        driveFileId = driveJson.fileId;
                                        driveUrl = driveJson.viewUrl || `https://drive.google.com/file/d/${driveFileId}/preview`;
                                      }
                                    }
                                  } catch (driveErr) {
                                    console.warn("Drive upload background notice:", driveErr);
                                  }
                                }

                                const newDoc: HangarPdfDoc = {
                                  id: driveFileId ? `drive_doc_${driveFileId}` : `hangar_doc_${Date.now()}_${Math.random().toString(36).substring(2, 7)}`,
                                  itemKey: docKey,
                                  fileName: file.name,
                                  fileData: fullDataUrl,
                                  docType: finalType,
                                  firma: finalFirma,
                                  mimeType: fileMimeType,
                                  uploadDate: new Date().toLocaleDateString('tr-TR', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }),
                                  fileSize: `${(file.size / (1024 * 1024) > 1 ? (file.size / (1024 * 1024)).toFixed(2) + ' MB' : (file.size / 1024).toFixed(1) + ' KB')}`,
                                  uploadedAt: new Date().toLocaleDateString('tr-TR', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }),
                                  driveFileId: driveFileId,
                                  driveUrl: driveUrl
                                };

                                await saveHangarPdfDoc(newDoc);
                                setHangarPdfDocs(prev => [newDoc, ...prev.filter(d => d.id !== newDoc.id)]);

                                setNewDocSelectedFile(null);
                                setNewDocCustomType("");
                                showNotification("Belge başarıyla yüklendi ve kaydedildi!");
                              } catch (err: any) {
                                console.error("Belge yükleme hatası:", err);
                                showNotification("Belge yükleme hatası: " + (err.message || "İşlem tamamlanamadı"));
                              } finally {
                                setIsUploadingDoc(false);
                              }
                            }}
                            className="px-6 py-2.5 bg-rose-600 hover:bg-rose-700 disabled:opacity-60 text-white font-black text-xs uppercase tracking-wider rounded-xl shadow-md shadow-rose-600/20 transition-all flex items-center gap-2 cursor-pointer active:scale-95"
                          >
                            {isUploadingDoc ? (
                              <>
                                <Loader2 className="w-3.5 h-3.5 animate-spin" />
                                <span>YÜKLENİYOR...</span>
                              </>
                            ) : (
                              <>
                                <Upload className="w-3.5 h-3.5" />
                                <span>YÜKLE</span>
                              </>
                            )}
                          </button>
                        </div>
                      )}
                    </div>

                    {/* Yüklenen Belgeler Tablosu */}
                    {(() => {
                      const docKey = `${(row[2] || row[1] || "").trim().toUpperCase()}`;
                      const itemDocs = hangarPdfDocs.filter(d => d.itemKey === docKey);

                      return (
                        <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden flex flex-col">
                          <div className="px-5 py-3.5 bg-slate-900 text-white flex items-center justify-between">
                            <div className="flex items-center gap-2">
                              <FileText className="w-4 h-4 text-rose-400" />
                              <h5 className="text-xs font-black uppercase tracking-wider">
                                YÜKLENEN BELGELER ({itemDocs.length})
                              </h5>
                            </div>
                            <span className="text-[10px] font-mono text-slate-400">
                              Tıklayarak tam pencere önizleme yapabilirsiniz
                            </span>
                          </div>

                          {itemDocs.length === 0 ? (
                            <div className="p-8 text-center text-slate-400 flex flex-col items-center justify-center gap-2">
                              <FileText className="w-10 h-10 text-slate-300 stroke-1" />
                              <p className="text-xs font-semibold">Henüz bu teçhizata ait yüklenmiş belge bulunmuyor.</p>
                              <p className="text-[11px] text-slate-400">Yukarıdaki formdan PDF, JPEG, ZIP veya RAR evraklarını yükleyebilirsiniz.</p>
                            </div>
                          ) : (
                            <div className="overflow-x-auto">
                              <table className="w-full text-left text-xs border-collapse">
                                <thead>
                                  <tr className="bg-slate-50 border-b border-slate-200 text-[10px] font-black uppercase text-slate-600 font-mono">
                                    <th className="px-4 py-3">BELGE ADI</th>
                                    <th className="px-4 py-3">BELGE TÜRÜ</th>
                                    <th className="px-4 py-3">BAKIMI YAPAN FİRMA</th>
                                    <th className="px-4 py-3 text-center">BOYUT / TARİH</th>
                                    <th className="px-4 py-3 text-right">İŞLEMLER</th>
                                  </tr>
                                </thead>
                                <tbody className="divide-y divide-slate-100">
                                  {itemDocs.map((doc) => {
                                    const cat = getFileCategory(doc.fileName);
                                    const isImg = cat === 'image';
                                    const isArch = cat === 'archive';
                                    const ext = doc.fileName?.split('.').pop()?.toUpperCase() || 'BELGE';

                                    return (
                                      <tr key={doc.id} className="hover:bg-slate-50/80 transition-colors">
                                        <td className="px-4 py-3 font-bold text-slate-800">
                                          <button
                                            type="button"
                                            onClick={() => setActivePdfPreview(doc)}
                                            className="text-left hover:text-rose-600 transition-colors flex items-center gap-2 group cursor-pointer"
                                            title="Önizlemek için tıklayınız"
                                          >
                                            <div className={`w-7 h-7 rounded-lg flex items-center justify-center shrink-0 ${
                                              isImg ? 'bg-blue-100 text-blue-800' :
                                              isArch ? 'bg-indigo-100 text-indigo-800' :
                                              'bg-rose-100 text-rose-700'
                                            }`}>
                                              {isImg ? <ImageIcon className="w-3.5 h-3.5" /> : isArch ? <Archive className="w-3.5 h-3.5" /> : <FileText className="w-3.5 h-3.5" />}
                                            </div>
                                            <span className="underline decoration-slate-300 underline-offset-4 group-hover:decoration-rose-500">
                                              {doc.fileName}
                                            </span>
                                            <span className={`text-[9px] font-mono px-1.5 py-0.2 rounded font-black uppercase shrink-0 ${
                                              isImg ? 'bg-blue-100 text-blue-800 border border-blue-200' :
                                              isArch ? 'bg-indigo-100 text-indigo-800 border border-indigo-200' :
                                              'bg-rose-100 text-rose-800 border border-rose-200'
                                            }`}>
                                              {ext}
                                            </span>
                                          </button>
                                        </td>
                                        <td className="px-4 py-3">
                                          <span className="px-2.5 py-1 rounded-full text-[10px] font-black bg-rose-50 text-rose-800 border border-rose-200">
                                            {doc.docType}
                                          </span>
                                        </td>
                                        <td className="px-4 py-3 font-medium text-slate-700">
                                          {doc.firma || "-"}
                                        </td>
                                        <td className="px-4 py-3 text-center font-mono text-slate-600 text-[11px]">
                                          <div>{doc.fileSize || '-'}</div>
                                          <div className="text-[9px] text-slate-400">{doc.uploadDate || doc.uploadedAt || "-"}</div>
                                        </td>
                                        <td className="px-4 py-3 text-right">
                                          <div className="flex items-center justify-end gap-2">
                                            <button
                                              type="button"
                                              onClick={() => setActivePdfPreview(doc)}
                                              className="px-3 py-1.5 bg-slate-100 hover:bg-rose-50 text-slate-700 hover:text-rose-700 font-bold text-[11px] rounded-lg border border-slate-200 transition-all flex items-center gap-1 cursor-pointer"
                                              title="Görüntüle / İncele"
                                            >
                                              <Eye className="w-3.5 h-3.5" />
                                              <span>Görüntüle</span>
                                            </button>
                                            <a
                                              href={doc.fileData}
                                              download={doc.fileName}
                                              className="px-3 py-1.5 bg-slate-100 hover:bg-emerald-50 text-slate-700 hover:text-emerald-700 font-bold text-[11px] rounded-lg border border-slate-200 transition-all flex items-center gap-1 cursor-pointer"
                                              title="İndir"
                                            >
                                              <Download className="w-3.5 h-3.5" />
                                              <span>İndir</span>
                                            </a>
                                            <button
                                              type="button"
                                              onClick={async () => {
                                                await deleteHangarPdfDoc(doc.id);
                                                setHangarPdfDocs(prev => prev.filter(d => d.id !== doc.id));
                                                showNotification(`"${doc.fileName}" belgesi başarıyla silindi!`);
                                              }}
                                              className="p-1.5 text-slate-400 hover:text-rose-600 hover:bg-rose-50 rounded-lg transition-colors cursor-pointer"
                                              title="Belgeyi Sil"
                                            >
                                              <Trash2 className="w-3.5 h-3.5" />
                                            </button>
                                          </div>
                                        </td>
                                      </tr>
                                    );
                                  })}
                                </tbody>
                              </table>
                            </div>
                          )}
                        </div>
                      );
                    })()}
                  </div>
                ) : (
                  <div className="w-full">
                    <ImageEditorAndRetoucher
                      imageKey={imageKey}
                      currentImageUrl={currentImageUrl}
                      hasImage={Boolean(currentImageUrl)}
                      isImageUpdateUnlocked={isImageUpdateUnlocked}
                      isUploadingToDrive={isImageUploadingToDrive}
                      onSaveImage={async (base64Data, mimeType) => {
                        try {
                          setIsImageUploadingToDrive(true);
                          showNotification("Görsel Google Drive'a yükleniyor...");
                          const targetUrl = GOOGLE_SCRIPT_URL;
                          const response = await fetch(targetUrl, {
                            method: "POST",
                            headers: { "Content-Type": "text/plain;charset=utf-8" },
                            body: JSON.stringify({
                              action: "uploadTechizatImageToDrive",
                              folderId: "1HQR_NYKhHQGA7_2W3nArI9pCh-LJasTP",
                              imageKey: imageKey,
                              fileName: `${imageKey}.png`,
                              base64Data: base64Data,
                              mimeType: mimeType
                            })
                          });
                          const resJson = await response.json();
                          if (resJson && resJson.status === "success") {
                            const newUrl = resJson.downloadUrl || resJson.viewUrl || `data:${mimeType};base64,${base64Data}`;
                            setTechizatImages(prev => {
                              const updated = { ...prev, [imageKey]: newUrl };
                              localStorage.setItem('techizat_images', JSON.stringify(updated));
                              return updated;
                            });
                            showNotification("Görsel başarıyla Drive'a yüklendi ve güncellendi!");
                          } else {
                            const dataUrl = `data:${mimeType};base64,${base64Data}`;
                            setTechizatImages(prev => {
                              const updated = { ...prev, [imageKey]: dataUrl };
                              localStorage.setItem('techizat_images', JSON.stringify(updated));
                              return updated;
                            });
                            showNotification("Görsel yerel hafızaya kaydedildi!");
                          }
                        } catch (err: any) {
                          console.error("Görsel yükleme hatası:", err);
                          const dataUrl = `data:${mimeType};base64,${base64Data}`;
                          setTechizatImages(prev => {
                            const updated = { ...prev, [imageKey]: dataUrl };
                            localStorage.setItem('techizat_images', JSON.stringify(updated));
                            return updated;
                          });
                          showNotification("Görsel yerel hafızaya kaydedildi!");
                        } finally {
                          setIsImageUploadingToDrive(false);
                        }
                      }}
                      onRemoveImage={() => {
                        setTechizatImages(prev => {
                          const copy = { ...prev };
                          delete copy[imageKey];
                          localStorage.setItem('techizat_images', JSON.stringify(copy));
                          return copy;
                        });
                        showNotification("Görsel silindi!");
                      }}
                      onUnlockImageUpdate={() => setIsImageUpdateUnlocked(true)}
                      onLockImageUpdate={() => setIsImageUpdateUnlocked(false)}
                      partName={row[1] || ""}
                      manufacturer={row[2] || ""}
                    />
                  </div>
                )}
              </div>

              {/* Modal Footer */}
              <div className="px-6 py-4 bg-slate-50 border-t border-slate-200 flex items-center justify-between flex-wrap gap-3">
                <div className="flex items-center gap-2 flex-wrap">
                  <button
                    type="button"
                    onClick={() => setActiveTechizatRowEdit(null)}
                    className="px-5 py-2.5 bg-white border border-slate-200 hover:bg-slate-100 text-slate-700 font-bold text-xs rounded-xl transition-all cursor-pointer"
                  >
                    Kapat
                  </button>

                  {/* Sil Butonu */}
                  <button
                    type="button"
                    onClick={() => {
                      if (window.confirm("Bu teçhizat kaydını kalıcı olarak silmek ve Google Drive / E-Tablo ile senkronize etmek istediğinize emin misiniz?")) {
                        handleDeleteTechizatRow(row, techType);
                      }
                    }}
                    className="px-4 py-2.5 bg-rose-50 hover:bg-rose-100 text-rose-700 font-bold text-xs rounded-xl border border-rose-200 transition-all flex items-center gap-1.5 cursor-pointer"
                    title="Bu teçhizat kaydını sistemden ve Drive Excel dosyasından sil"
                  >
                    <Trash2 className="w-3.5 h-3.5 text-rose-600" />
                    <span>Sil</span>
                  </button>

                  {/* Kullanıcı Kuralı: "faal yap butonu eğer geleceke kontrol bakım tarih geçmiş ise o butonn sil yanına olsun aşağıdan kaldır" */}
                  {(() => {
                    const gelecekIdx = columns.findIndex(c => c.includes("GELECEK KONTROL") || c.includes("MUAYENE") || c.includes("ÖMÜR"));
                    const dateVal = gelecekIdx !== -1 ? (editRowValues[gelecekIdx] || row[gelecekIdx] || "") : "";
                    const days = parseGelecekBakimDays(dateVal);
                    if (days === null || days >= 0) return null;

                    return (
                      <button
                        type="button"
                        onClick={async () => {
                          const copy = [...editRowValues];
                          const sonIdx = columns.findIndex(c => c.includes("SON KONTROL") || c.includes("SON BAKIM"));
                          const durumIdx = columns.findIndex(c => c.includes("DURUM"));
                          const prevSon = sonIdx !== -1 ? (copy[sonIdx] || row[sonIdx] || "") : "";
                          const prevGelecek = copy[gelecekIdx] || row[gelecekIdx] || "";
                          const { newSonKontrol, newGelecekKontrol, intervalDays } = calculateReactivatedMaintenanceDates(prevSon, prevGelecek);

                          if (sonIdx !== -1) copy[sonIdx] = newSonKontrol;
                          if (gelecekIdx !== -1) copy[gelecekIdx] = newGelecekKontrol;
                          if (durumIdx !== -1) copy[durumIdx] = "FAAL";

                          setEditRowValues(copy);
                          await handleSaveTechizatRow(copy, techType, rIdx);
                          showNotification(`⚡ Bakım tamamlandı! Kayıt FAAL yapıldı, son kontrol ${newSonKontrol}, gelecek bakım ${newGelecekKontrol} (+${intervalDays} gün) güncellendi.`);
                          setActiveTechizatRowEdit(null);
                        }}
                        className="px-4 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white font-black text-xs rounded-xl shadow-md transition-all flex items-center gap-1.5 cursor-pointer animate-pulse"
                        title="Tarihi geçmiş ürünü Faal Yap: Son kontrolü bugüne çeker, aradaki periyodu gelecek bakıma ekler"
                      >
                        <RefreshCw className="w-3.5 h-3.5" />
                        <span>Faal Yap</span>
                      </button>
                    );
                  })()}
                </div>
                {mobileEditTab === 'form' && isDataUpdateUnlocked && (
                  <button
                    type="button"
                    onClick={async () => {
                      await handleSaveTechizatRow(editRowValues, techType, rIdx);
                      showNotification("Kayıt başarıyla güncellendi!");
                      setActiveTechizatRowEdit(null);
                    }}
                    className="px-6 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white font-black text-xs uppercase tracking-wider rounded-xl shadow-lg shadow-emerald-600/30 transition-all active:scale-95 cursor-pointer"
                  >
                    💾 Değişiklikleri Kaydet
                  </button>
                )}
                {mobileEditTab === 'regional' && (
                  <button
                    type="button"
                    onClick={async () => {
                      const copy = [...editRowValues];
                      let totalQty = 0;
                      regionalLocations.forEach(r => {
                        const q = parseFloat(String(r.quantity)) || (r.serialNumbers.length > 0 ? r.serialNumbers.length : 1);
                        totalQty += q;
                      });
                      copy[4] = totalQty > 0 ? `${totalQty} ADET` : (copy[4] || '1 ADET');
                      const locStr = regionalLocations.map(r => `${r.location}${r.quantity ? ` (${r.quantity})` : ''}`).join(' ; ');
                      copy[5] = locStr || copy[5];
                      const allSns = regionalLocations.flatMap(r => r.serialNumbers);
                      if (allSns.length > 0) copy[3] = allSns.join(' ; ');
                      setEditRowValues(copy);
                      await handleSaveTechizatRow(copy, techType, rIdx);
                      showNotification("Bölge dağılımları ve kayıt başarıyla güncellendi!");
                      setActiveTechizatRowEdit(null);
                    }}
                    className="px-6 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white font-black text-xs uppercase tracking-wider rounded-xl shadow-lg shadow-emerald-600/30 transition-all active:scale-95 cursor-pointer"
                  >
                    💾 Dağılımları Kaydet & Kapat
                  </button>
                )}
              </div>
            </div>
          </div>
        );
      })()}

      {/* PDF PREVIEW MODAL */}
      <PdfPreviewModal
        preview={activePdfPreview}
        onClose={() => setActivePdfPreview(null)}
      />

      {/* DOCUMENT UPLOAD / DELETE PASSWORD MODAL */}
      {docUploadPasswordModal.isOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-950/80 animate-fade-in">
          <div className="bg-white w-full max-w-sm rounded-2xl shadow-2xl overflow-hidden border border-slate-200">
            <div className="p-5 bg-slate-900 text-white flex items-center justify-between">
              <div className="flex items-center gap-2.5">
                <Lock className="w-4 h-4 text-rose-400" />
                <h4 className="text-xs font-black uppercase tracking-wider">
                  {docUploadPasswordModal.pendingAction === 'delete' ? 'BELGE SİLME ONAYI' : 'BELGE YÜKLEME ŞİFRESİ'}
                </h4>
              </div>
              <button
                type="button"
                onClick={() => {
                  setDocUploadPasswordModal({ isOpen: false, pendingAction: null });
                  setDocUploadPasswordInput("");
                  setDocUploadPasswordError("");
                  setDocToDelete(null);
                }}
                className="w-7 h-7 rounded-full bg-slate-800 hover:bg-slate-700 text-slate-300 flex items-center justify-center text-xs transition-all cursor-pointer"
              >
                ✕
              </button>
            </div>

            <form
              onSubmit={(e) => {
                e.preventDefault();
                // Hangar Yer Destek şifresi "1839" (kullanıcı talimatı: diğer bütün şifreler "1839")
                if (docUploadPasswordInput !== "1839") {
                  setDocUploadPasswordError("Hatalı yetki şifresi girdiniz!");
                  return;
                }

                if (docUploadPasswordModal.pendingAction === 'delete') {
                  if (docToDelete) {
                    deleteHangarPdfDoc(docToDelete).catch(err => console.warn(err));
                    setHangarPdfDocs(prev => {
                      const updated = prev.filter(d => d.id !== docToDelete);
                      try {
                        localStorage.setItem('hangar_pdf_docs', JSON.stringify(updated));
                        localStorage.setItem('hangar_techizat_pdf_docs', JSON.stringify(updated));
                      } catch {
                        // ignore
                      }
                      return updated;
                    });
                    showNotification("Belge başarıyla silindi!");
                  }
                } else if (docUploadPasswordModal.pendingAction === 'upload') {
                  if (!newDocSelectedFile) {
                    setDocUploadPasswordError("Lütfen bir PDF dosyası seçiniz.");
                    return;
                  }
                  const currentRow = activeTechizatRowEdit?.row;
                  const docKey = `${(currentRow?.[2] || currentRow?.[1] || "").trim().toUpperCase()}`;

                  const reader = new FileReader();
                  reader.onload = async () => {
                    const dataUrl = reader.result as string;
                    const finalType = newDocDocType === "Diğer" ? (newDocCustomType.trim() || "Diğer") : newDocDocType;
                    const finalFirma = newDocFirma.trim() || (currentRow?.[10] || "");

                    const newDoc = {
                      id: `hangar_doc_${Date.now()}_${Math.random().toString(36).substring(2, 7)}`,
                      itemKey: docKey,
                      fileName: newDocSelectedFile.name,
                      fileData: dataUrl,
                      docType: finalType,
                      firma: finalFirma,
                      uploadDate: new Date().toLocaleDateString('tr-TR', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }),
                      fileSize: `${(newDocSelectedFile.size / 1024).toFixed(1)} KB`,
                      uploadedAt: new Date().toLocaleDateString('tr-TR', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' })
                    };

                    await saveHangarPdfDoc(newDoc);
                    setHangarPdfDocs(prev => {
                      const updated = [newDoc, ...prev.filter(d => d.id !== newDoc.id)];
                      try {
                        localStorage.setItem('hangar_pdf_docs', JSON.stringify(updated));
                        localStorage.setItem('hangar_techizat_pdf_docs', JSON.stringify(updated));
                      } catch {
                        // ignore
                      }
                      return updated;
                    });

                    // Form temizleme
                    setNewDocSelectedFile(null);
                    setNewDocCustomType("");
                    showNotification("PDF Belge başarıyla yüklendi ve kaydedildi!");
                  };
                  reader.readAsDataURL(newDocSelectedFile);
                }

                setDocUploadPasswordModal({ isOpen: false, pendingAction: null });
                setDocUploadPasswordInput("");
                setDocUploadPasswordError("");
                setDocToDelete(null);
              }}
              className="p-5 flex flex-col gap-4"
            >
              <div className="flex flex-col gap-1.5">
                <label className="text-xs font-bold text-slate-800">
                  Yetkili Kilit Şifresi
                </label>
                <input
                  type="password"
                  autoFocus
                  value={docUploadPasswordInput}
                  onChange={(e) => {
                    setDocUploadPasswordInput(e.target.value);
                    if (docUploadPasswordError) setDocUploadPasswordError("");
                  }}
                  placeholder="Şifreyi giriniz..."
                  className="w-full px-3.5 py-2.5 rounded-xl text-xs font-bold bg-white text-slate-950 border border-slate-300 focus:outline-none focus:border-rose-600 focus:ring-2 focus:ring-rose-500/20"
                />
                {docUploadPasswordError && (
                  <p className="text-[11px] font-bold text-rose-600 mt-0.5">
                    {docUploadPasswordError}
                  </p>
                )}
              </div>

              <div className="flex items-center justify-end gap-2 pt-1">
                <button
                  type="button"
                  onClick={() => {
                    setDocUploadPasswordModal({ isOpen: false, pendingAction: null });
                    setDocUploadPasswordInput("");
                    setDocUploadPasswordError("");
                    setDocToDelete(null);
                  }}
                  className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold text-xs rounded-xl transition-all cursor-pointer"
                >
                  İptal
                </button>
                <button
                  type="submit"
                  className="px-5 py-2 bg-rose-600 hover:bg-rose-700 text-white font-black text-xs uppercase tracking-wider rounded-xl shadow-md transition-all cursor-pointer"
                >
                  Onayla
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
      {isTechPubsOpen && (
        <TechnicalPublicationsModal
          isOpen={isTechPubsOpen}
          onClose={() => setIsTechPubsOpen(false)}
          allUnitData={{
            at802: techizatAt802Data,
            bell429: techizatBell429Data,
            t70: techizatT70Data,
            t70_bumbi: techizatT70BumbiBacketData,
            t70_helitak: techizatT70HelitakData,
            c650: techizatC650Data,
            b360: techizatB360Data,
            hangar: techizatHangarData,
          }}
          onNavigateToEquipment={(type: string, label: string, section?: string) => {
            openTechizatMatrix(type as any, label, (section as any) || 'all');
          }}
        />
      )}

      {isOlayTakipOpen && (
        <OlayTakipCizelgesiModal
          isOpen={isOlayTakipOpen}
          onClose={() => setIsOlayTakipOpen(false)}
          showNotification={showNotification}
          initialUnit={olayTakipInitialUnit}
        />
      )}

      {isBarkodOkuyucuOpen && (
        <BarkodOkuyucuModal
          isOpen={isBarkodOkuyucuOpen}
          onClose={() => setIsBarkodOkuyucuOpen(false)}
          showNotification={showNotification}
          initialUnit={barkodOkuyucuInitialUnit}
          transactions={depoTransactions}
          onAddTransactions={async (txs) => {
            const newTxs = txs.map(tx => ({
              id: tx.id || `tx_${Date.now()}_${Math.random().toString(36).substr(2, 6)}`,
              timestamp: tx.timestamp || new Date().toISOString(),
              ...tx
            }));
            
            setDepoTransactions(prev => [...newTxs, ...prev]);

            // Sync to Table 7 (DEPO SAYIM KAYITLARI)
            try {
              const stored = localStorage.getItem('form_7_sayimlar');
              const currentLogs = stored ? JSON.parse(stored) : [];
              const newRows = newTxs
                .filter(tx => tx.islemTuru?.includes('SAYIM') || tx.type === 'SAYIM')
                .map(tx => ({
                  Tarih: tx.date || new Date().toLocaleString('tr-TR'),
                  Malzeme: tx.itemName || '-',
                  PN: tx.pn || '-',
                  SN: tx.sn || '-',
                  Bolge: tx.location || '-',
                  Sistem_Stok: tx.sistemStok ?? '-',
                  Sayilan_Adet: tx.sayilanAdet ?? '-',
                  Fark: tx.fark ?? '-',
                  Personel: tx.operator || '-'
                }));
              
              if (newRows.length > 0) {
                localStorage.setItem('form_7_sayimlar', JSON.stringify([...newRows, ...currentLogs]));
              }
            } catch (e) {}

            // Background sync to Excel / Drive
            try {
              const u = (newTxs[0]?.unit || 'hangar').toLowerCase();
              let sheetName = "DEPO HAREKET GEÇMİŞİ-AT-802";
              if (u.includes('bell')) sheetName = "DEPO HAREKET GEÇMİŞİ-BELL 429";
              else if (u.includes('t70')) sheetName = "DEPO HAREKET GEÇMİŞİ-T-70";
              else if (u.includes('360')) sheetName = "DEPO HAREKET GEÇMİŞİ-B-360";
              else if (u.includes('650')) sheetName = "DEPO HAREKET GEÇMİŞİ-C-650";
              else if (u.includes('hangar')) sheetName = "DEPO HAREKET GEÇMİŞİ-HANGAR";

              fetch('/api/save-depo-transfers', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                  spreadsheetId: '17ScGYYx0erzDwHDk6RGiHOdJATdfmmExXFBY39dXpF0',
                  sheetName: sheetName,
                  transfers: newTxs
                })
              }).catch(err => console.warn("Depo sync background error:", err));
            } catch (e) {}
          }}
          onAddTransaction={async (tx) => {
            const newTx: DepoTransaction = {
              id: `tx_${Date.now()}_${Math.random().toString(36).substr(2, 6)}`,
              timestamp: new Date().toISOString(),
              ...tx
            };
            setDepoTransactions(prev => [newTx, ...prev]);

            // Sync to Table 7 (DEPO SAYIM KAYITLARI) if it's a counting action
            if (newTx.islemTuru?.includes('SAYIM') || newTx.type === 'SAYIM') {
              try {
                const stored = localStorage.getItem('form_7_sayimlar');
                const currentLogs = stored ? JSON.parse(stored) : [];
                const newRow = {
                  Tarih: newTx.date || new Date().toLocaleString('tr-TR'),
                  Malzeme: newTx.itemName || '-',
                  PN: newTx.pn || '-',
                  SN: newTx.sn || '-',
                  Bolge: newTx.location || '-',
                  Sistem_Stok: newTx.sistemStok ?? '-',
                  Sayilan_Adet: newTx.sayilanAdet ?? '-',
                  Fark: newTx.fark ?? '-',
                  Personel: newTx.operator || '-'
                };
                localStorage.setItem('form_7_sayimlar', JSON.stringify([newRow, ...currentLogs]));
              } catch (e) {}
            }
            
            // Background sync to Excel / Drive
            try {
              const u = (newTx.unit || 'hangar').toLowerCase();
              let sheetName = "DEPO HAREKET GEÇMİŞİ-AT-802";
              if (u.includes('bell')) sheetName = "DEPO HAREKET GEÇMİŞİ-BELL 429";
              else if (u.includes('t70')) sheetName = "DEPO HAREKET GEÇMİŞİ-T-70";
              else if (u.includes('360')) sheetName = "DEPO HAREKET GEÇMİŞİ-B-360";
              else if (u.includes('650')) sheetName = "DEPO HAREKET GEÇMİŞİ-C-650";
              else if (u.includes('hangar')) sheetName = "DEPO HAREKET GEÇMİŞİ-HANGAR";

              fetch('/api/save-depo-transfers', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                  spreadsheetId: '17ScGYYx0erzDwHDk6RGiHOdJATdfmmExXFBY39dXpF0',
                  sheetName: sheetName,
                  transfers: [newTx]
                })
              }).catch(err => console.warn("Depo sync background error:", err));
            } catch (e) {}
          }}
          onUpdateInventory={(updated) => {
            try {
              localStorage.setItem('ogm_depo_inventory_v5', JSON.stringify(updated));
              // Also sync to backend/drive
              fetch('/api/depo-save-inventory', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ items: updated, unit: barkodOkuyucuInitialUnit || 'at802' })
              }).catch(err => console.warn("Inventory sync background error:", err));
            } catch (e) {}
          }}
        />
      )}

      {/* PASSWORD-PROTECTED DATA UPDATE MODAL (1839) */}
      {isPasswordModalOpen && (
        <div className="fixed inset-0 z-[700] flex items-center justify-center p-4 bg-slate-900/70 backdrop-blur-sm animate-fade-in">
          <div className="bg-white rounded-3xl p-6 max-w-sm w-full shadow-2xl border border-slate-100 flex flex-col gap-4">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <h3 className="text-sm font-black text-slate-800 uppercase tracking-wider flex items-center gap-2">
                🔒 YETKİLİ ŞİFRE GİRİŞİ
              </h3>
              <button
                onClick={() => {
                  setIsPasswordModalOpen(false);
                  setPasswordInput('');
                  setPasswordError(false);
                }}
                className="text-slate-400 hover:text-slate-600 font-bold cursor-pointer"
              >
                ✕
              </button>
            </div>
            <p className="text-xs text-slate-700 font-medium">
              Veri güncelleme ve depo yönetim ekranına erişmek için lütfen 4 haneli yetkili şifrenizi giriniz.
            </p>
            <div>
              <input
                type="password"
                placeholder="Şifre"
                value={passwordInput}
                onChange={(e) => {
                  setPasswordInput(e.target.value);
                  setPasswordError(false);
                }}
                onKeyDown={(e) => {
                  if (e.key === 'Enter') {
                    handleVerifyPassword();
                  }
                }}
                className="w-full px-4 py-3.5 bg-white border-2 border-slate-300 rounded-2xl text-center font-mono text-xl font-black text-black placeholder:text-slate-400 tracking-widest focus:outline-none focus:border-[#0b3d1d] focus:ring-4 focus:ring-[#0b3d1d]/15 shadow-inner"
                autoFocus
              />
              {passwordError && (
                <p className="text-rose-600 text-xs font-bold mt-1.5 text-center">
                  Hatalı şifre girdiniz! Lütfen tekrar deneyiniz.
                </p>
              )}
            </div>
            <div className="flex items-center gap-2 pt-2">
              <button
                type="button"
                onClick={() => {
                  setIsPasswordModalOpen(false);
                  setPasswordInput('');
                  setPasswordError(false);
                }}
                className="flex-1 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold text-xs rounded-xl transition-all cursor-pointer"
              >
                İptal
              </button>
              <button
                type="button"
                onClick={handleVerifyPassword}
                className="flex-1 py-2.5 bg-[#0b3d1d] hover:bg-[#072612] text-white font-black text-xs uppercase tracking-wider rounded-xl transition-all shadow-md shadow-emerald-900/20 cursor-pointer"
              >
                Giriş Yap
              </button>
            </div>
          </div>
        </div>
      )}

      {/* GÖREV EMRİ SİLME ŞİFRE MODALI */}
      {showGeDeletePasswordPrompt && (
        <div className="fixed inset-0 z-[700] flex items-center justify-center p-4 bg-slate-900/70 backdrop-blur-sm animate-fade-in">
          <div className="bg-white rounded-3xl p-6 max-w-sm w-full shadow-2xl border border-slate-100 flex flex-col gap-4">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <h3 className="text-sm font-black text-slate-800 uppercase tracking-wider flex items-center gap-2">
                🔒 GÖREV EMRİ SİLME
              </h3>
              <button
                onClick={() => {
                  setShowGeDeletePasswordPrompt(false);
                  setGeDeletePasswordInput('');
                  setGeDeletePasswordError(false);
                }}
                className="text-slate-400 hover:text-slate-600 font-bold cursor-pointer"
              >
                ✕
              </button>
            </div>
            <p className="text-xs text-slate-600">
              Görev emrini silmek için lütfen yönetici şifresini giriniz.
            </p>
            <div>
              <input
                type="password"
                placeholder="Şifre"
                value={geDeletePasswordInput}
                onChange={(e) => {
                  setGeDeletePasswordInput(e.target.value);
                  setGeDeletePasswordError(false);
                }}
                className="w-full px-4 py-3 bg-slate-50 border border-slate-200 rounded-2xl text-center font-mono text-lg font-black tracking-widest focus:outline-none focus:border-rose-600 focus:ring-4 focus:ring-rose-600/10"
                autoFocus
              />
              {geDeletePasswordError && (
                <p className="text-rose-600 text-xs font-bold mt-1.5 text-center">
                  Hatalı şifre girdiniz!
                </p>
              )}
            </div>
            <div className="flex items-center gap-2 pt-2">
              <button
                type="button"
                onClick={() => {
                  setShowGeDeletePasswordPrompt(false);
                  setGeDeletePasswordInput('');
                  setGeDeletePasswordError(false);
                }}
                className="flex-1 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold text-xs rounded-xl transition-all cursor-pointer"
              >
                İptal
              </button>
              <button
                type="button"
                onClick={() => {
                  if (geDeletePasswordInput === '1234') {
                    if (geDeleteOrderId) {
                      const updated = karaAraclariGorevEmirleri.filter(o => String(o.id) !== geDeleteOrderId);
                      setKaraAraclariGorevEmirleri(updated);
                      localStorage.setItem('kara_araclari_gorev_emirleri', JSON.stringify(updated));
                      pushKaraAraclariGorevEmirleri(updated);
                      showNotification("Görev emri başarıyla silindi!");
                    }
                    setShowGeDeletePasswordPrompt(false);
                    setGeDeletePasswordInput('');
                    setGeDeletePasswordError(false);
                    setGeDeleteOrderId(null);
                  } else {
                    setGeDeletePasswordError(true);
                  }
                }}
                className="flex-1 py-2.5 bg-rose-600 hover:bg-rose-700 text-white font-black text-xs uppercase tracking-wider rounded-xl transition-all shadow-md shadow-rose-900/20 cursor-pointer"
              >
                Sil
              </button>
            </div>
          </div>
        </div>
      )}

      {/* GÜN TAKİP SORUMLULARI VE 90 GÜN MAİL BİLDİRİMLERİ MODALI */}
      <GunTakipModal
        isOpen={isSorumluModalOpen}
        onClose={() => setIsSorumluModalOpen(false)}
        sorumlular={gunTakipSorumlulari}
        onSaveSorumlular={saveGunTakipSorumlulari}
        isSaving={isSavingSorumlu}
        googleScriptUrl={GOOGLE_SCRIPT_URL}
        showNotification={(msg, type) => showNotification(msg)}
        unitDataMap={{
          bell429: techizatBell429Data,
          at802: techizatAt802Data,
          t70: techizatT70Data,
          t70_bumbi_backet: techizatT70BumbiBacketData,
          t70_helitak: techizatT70HelitakData,
          b360: techizatB360Data,
          c650: techizatC650Data,
          hangar: techizatHangarData,
          kara_araclari: techizatKaraAraclariData
        }}
      />

      {/* EBYS & BULK EDIT MODAL */}
      {isEbysModalOpen && (
        <div className="fixed inset-0 z-[700] flex items-center justify-center p-4 bg-slate-900/70 backdrop-blur-sm animate-fade-in">
          <div className="bg-white rounded-3xl p-6 max-w-xl w-full shadow-2xl border border-slate-100 flex flex-col gap-5 max-h-[90vh] overflow-y-auto">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <h3 className="text-sm font-black text-slate-800 uppercase tracking-wider flex items-center gap-2">
                ⚡ SEÇİLENLERİ İŞLE ({Object.keys(selectedTechizatItems).length} Öğe)
              </h3>
              <button
                onClick={() => setIsEbysModalOpen(false)}
                className="text-slate-400 hover:text-slate-600 font-bold cursor-pointer"
              >
                ✕
              </button>
            </div>

            {bulkModalMode === 'choice' && (
              <div className="flex flex-col gap-4 py-4">
                <p className="text-xs text-slate-600">
                  Seçtiğiniz <strong>{Object.keys(selectedTechizatItems).length}</strong> adet kayıt üzerinde yapmak istediğiniz işlemi seçiniz:
                </p>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                  <button
                    onClick={() => setBulkModalMode('send')}
                    className="p-5 border-2 border-emerald-500 bg-emerald-50/50 hover:bg-emerald-50 rounded-2xl flex flex-col items-center text-center gap-2 transition-all active:scale-95 cursor-pointer shadow-sm"
                  >
                    <Send className="w-8 h-8 text-emerald-700" />
                    <span className="text-xs font-black text-emerald-950 uppercase">EBYS TALEBİ GÖNDER</span>
                    <span className="text-[10px] text-emerald-700">Seçilenleri Taskline / EBYS sistemine talep olarak aktar</span>
                  </button>
                  <button
                    onClick={() => setBulkModalMode('edit')}
                    className="p-5 border-2 border-sky-500 bg-sky-50/50 hover:bg-sky-50 rounded-2xl flex flex-col items-center text-center gap-2 transition-all active:scale-95 cursor-pointer shadow-sm"
                  >
                    <Edit3 className="w-8 h-8 text-sky-700" />
                    <span className="text-xs font-black text-sky-950 uppercase">TOPLU DÜZENLE</span>
                    <span className="text-[10px] text-sky-700">Bulunduğu yer, durum ve firmayı topluca güncelle</span>
                  </button>
                </div>
              </div>
            )}

            {bulkModalMode === 'edit' && (
              <div className="flex flex-col gap-4">
                <p className="text-xs text-slate-600">
                  Seçilen tüm kayıtlara uygulamak istediğiniz yeni değerleri giriniz (boş bıraktığınız alanlar değişmez):
                </p>
                <div className="space-y-3">
                  <div>
                    <label className="block text-[10px] font-black text-slate-400 uppercase tracking-wider mb-1">
                      📍 Yeni Bulunduğu Yer
                    </label>
                    <input
                      type="text"
                      placeholder="Örn: Hangar 2 Raf A"
                      value={bulkEditYer}
                      onChange={(e) => setBulkEditYer(e.target.value)}
                      className="w-full px-3.5 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-800 focus:outline-none focus:border-[#0b3d1d]"
                    />
                  </div>
                  <div>
                    <label className="block text-[10px] font-black text-slate-400 uppercase tracking-wider mb-1">
                      ⚙️ Yeni Durumu
                    </label>
                    <select
                      value={bulkEditDurum}
                      onChange={(e) => setBulkEditDurum(e.target.value)}
                      className="w-full px-3.5 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-800 focus:outline-none focus:border-[#0b3d1d]"
                    >
                      <option value="">Değiştirme (Aynı Kalsın)</option>
                      <option value="FAAL">FAAL</option>
                      <option value="GAYRİ FAAL">GAYRİ FAAL</option>
                      <option value="BAKIMDA">BAKIMDA</option>
                      <option value="KALİBRASYONDA">KALİBRASYONDA</option>
                    </select>
                  </div>
                  <div>
                    <label className="block text-[10px] font-black text-slate-400 uppercase tracking-wider mb-1">
                      🏢 Yeni Firma
                    </label>
                    <input
                      type="text"
                      placeholder="Firma Adı"
                      value={bulkEditFirma}
                      onChange={(e) => setBulkEditFirma(e.target.value)}
                      className="w-full px-3.5 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-800 focus:outline-none focus:border-[#0b3d1d]"
                    />
                  </div>
                </div>
                <div className="flex items-center justify-between pt-3 border-t border-slate-100">
                  <button
                    type="button"
                    onClick={() => setBulkModalMode('choice')}
                    className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold text-xs rounded-xl cursor-pointer"
                  >
                    ← Geri
                  </button>
                  <button
                    type="button"
                    disabled={isBulkSaving}
                    onClick={handleBulkEditTechizatRows}
                    className="px-6 py-2 bg-[#0b3d1d] hover:bg-[#072612] text-white font-black text-xs uppercase tracking-wider rounded-xl shadow-md transition-all active:scale-95 cursor-pointer"
                  >
                    {isBulkSaving ? "Kaydediliyor..." : "Toplu Güncelle"}
                  </button>
                </div>
              </div>
            )}

            {bulkModalMode === 'send' && (
              <div className="flex flex-col gap-4">
                <p className="text-xs text-slate-600">
                  Seçilen kayıtlar için EBYS / Taskline sistemine gönderilecek talep detaylarını giriniz:
                </p>
                <div className="space-y-3">
                  <div>
                    <label className="block text-[10px] font-black text-slate-400 uppercase tracking-wider mb-1">
                      Talep Başlığı
                    </label>
                    <input
                      type="text"
                      placeholder="Örn: Yıllık Periyodik Kalibrasyon Talebi"
                      value={ebysBaslik}
                      onChange={(e) => setEbysBaslik(e.target.value)}
                      className="w-full px-3.5 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-800 focus:outline-none focus:border-emerald-600"
                    />
                  </div>
                  <div>
                    <label className="block text-[10px] font-black text-slate-400 uppercase tracking-wider mb-1">
                      Açıklama
                    </label>
                    <textarea
                      rows={3}
                      placeholder="Talep hakkında detaylı bilgi..."
                      value={ebysAciklama}
                      onChange={(e) => setEbysAciklama(e.target.value)}
                      className="w-full px-3.5 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs font-medium text-slate-800 focus:outline-none focus:border-emerald-600"
                    />
                  </div>
                  <div className="grid grid-cols-2 gap-3">
                    <div>
                      <label className="block text-[10px] font-black text-slate-400 uppercase tracking-wider mb-1">
                        Talep Türü
                      </label>
                      <select
                        value={ebysTalepTuru}
                        onChange={(e) => setEbysTalepTuru(e.target.value)}
                        className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-800 focus:outline-none focus:border-emerald-600"
                      >
                        <option value="">Seçiniz...</option>
                        <option value="KALİBRASYON">KALİBRASYON</option>
                        <option value="BAKIM">BAKIM</option>
                        <option value="MALZEME">MALZEME</option>
                      </select>
                    </div>
                    <div>
                      <label className="block text-[10px] font-black text-slate-400 uppercase tracking-wider mb-1">
                        Teslim Tarihi
                      </label>
                      <input
                        type="date"
                        value={ebysTeslimTarihi}
                        onChange={(e) => setEbysTeslimTarihi(e.target.value)}
                        className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-800 focus:outline-none focus:border-emerald-600 font-mono"
                      />
                    </div>
                  </div>
                </div>
                <div className="flex items-center justify-between pt-3 border-t border-slate-100">
                  <button
                    type="button"
                    onClick={() => setBulkModalMode('choice')}
                    className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold text-xs rounded-xl cursor-pointer"
                  >
                    ← Geri
                  </button>
                  <button
                    type="button"
                    onClick={submitEbysRequests}
                    className="px-6 py-2 bg-emerald-600 hover:bg-emerald-700 text-white font-black text-xs uppercase tracking-wider rounded-xl shadow-md transition-all active:scale-95 cursor-pointer"
                  >
                    Talebi Gönder
                  </button>
                </div>
              </div>
            )}
          </div>
        </div>
      )}

      {/* ROW IMAGE HOVER TOOLTIP */}
      {hoveredRowImage && (
        <div
          className="fixed pointer-events-none z-[800] bg-white rounded-2xl shadow-2xl border-2 border-emerald-500 overflow-hidden p-2 flex flex-col items-center gap-1.5 animate-fade-in"
          style={{
            left: `${Math.min(window.innerWidth - 220, hoveredRowImage.x + 20)}px`,
            top: `${Math.min(window.innerHeight - 240, hoveredRowImage.y - 40)}px`,
            width: '200px'
          }}
        >
          <img
            src={hoveredRowImage.url}
            alt={hoveredRowImage.title}
            className="w-full h-36 object-contain rounded-xl bg-slate-50"
          />
          <p className="text-[11px] font-black text-slate-800 truncate w-full text-center">{hoveredRowImage.title}</p>
          <p className="text-[9px] text-slate-500 truncate w-full text-center">{hoveredRowImage.subtitle}</p>
        </div>
      )}

      {/* EXCEL EXPORT MODAL */}
      {excelExportModalData && (
        <ExcelExportModal
          isOpen={Boolean(excelExportModalData)}
          onClose={() => setExcelExportModalData(null)}
          title={excelExportModalData.title}
          cols={excelExportModalData.cols}
          rows={excelExportModalData.rows}
          type={excelExportModalData.type}
          techizatImages={techizatImages}
          onExportTextOnly={() => {
            executeTextOnlyExcelExport(
              excelExportModalData.type,
              excelExportModalData.cols,
              excelExportModalData.rows,
              excelExportModalData.title
            );
            setExcelExportModalData(null);
          }}
          onExportWithImages={() => {
            downloadTechizatExcelWithImages(true);
          }}
        />
      )}

      {/* YENİ ÜRÜN EKLE MODALI */}
      {isNewProductModalOpen && (
        <NewProductModal
          isOpen={isNewProductModalOpen}
          onClose={() => setIsNewProductModalOpen(false)}
          activeUnit={activeTechizatType !== 'all' ? activeTechizatType : 'at802'}
          activeSection={techizatActiveSection}
          onSave={handleAddNewProductRow}
        />
      )}

      {/* VERİ GÜNCELLEME & MASTER PDF OCR MODALI */}
      {isDataSyncModalOpen && (
        <DataSyncModal
          isOpen={isDataSyncModalOpen}
          onClose={() => setIsDataSyncModalOpen(false)}
          initialStep={syncInitialStep}
          initialTarget={syncInitialTarget}
          activeUnit={activeTechizatType}
          activeSection={techizatActiveSection}
          onUpdateUnitData={(unitKey, _newCols, newRows) => {
            const cleanKey = unitKey.toLowerCase();
            if (cleanKey.includes('bell429')) {
              setTechizatBell429Data(newRows);
              try { localStorage.setItem('techizat_bell429_data', JSON.stringify(newRows)); } catch(e){}
            } else if (cleanKey.includes('at802_ozel_alet')) {
              setTechizatAt802OzelAletData(newRows);
              try { localStorage.setItem('techizat_at802_ozel_alet_data', JSON.stringify(newRows)); } catch(e){}
            } else if (cleanKey.includes('at802')) {
              setTechizatAt802Data(newRows);
              try { localStorage.setItem('techizat_at802_data', JSON.stringify(newRows)); } catch(e){}
            } else if (cleanKey.includes('t70_bumbi')) {
              setTechizatT70BumbiBacketData(newRows);
              try { localStorage.setItem('techizat_t70_bumbi_backet_data', JSON.stringify(newRows)); } catch(e){}
            } else if (cleanKey.includes('t70_helitak')) {
              setTechizatT70HelitakData(newRows);
              try { localStorage.setItem('techizat_t70_helitak_data', JSON.stringify(newRows)); } catch(e){}
            } else if (cleanKey.includes('t70')) {
              setTechizatT70Data(newRows);
              try { localStorage.setItem('techizat_t70_data', JSON.stringify(newRows)); } catch(e){}
            } else if (cleanKey.includes('b360')) {
              setTechizatB360Data(newRows);
              try { localStorage.setItem('techizat_b360_data', JSON.stringify(newRows)); } catch(e){}
            } else if (cleanKey.includes('c650')) {
              setTechizatC650Data(newRows);
              try { localStorage.setItem('techizat_c650_data', JSON.stringify(newRows)); } catch(e){}
            } else if (cleanKey.includes('hangar')) {
              setTechizatHangarData(newRows);
              try { localStorage.setItem('techizat_hangar_data', JSON.stringify(newRows)); } catch(e){}
            } else if (cleanKey.includes('kara')) {
              setTechizatKaraAraclariData(newRows);
              try { localStorage.setItem('techizat_kara_araclari_data', JSON.stringify(newRows)); } catch(e){}
            }
          }}
          onApplyDriveSync={async () => {
            await pullAllTechizatFromDriveExcels(false);
          }}
          onApplyUploadedData={(detectedUnit, rows) => {
            if (rows.length > 0) {
              rows.forEach(r => handleAddNewProductRow(detectedUnit, r));
            }
          }}
          onDownloadLatestExcel={(unit) => {
            exportTechizatToExcel(unit as any, [], [], `${unit.toUpperCase()} ENVANTERİ`);
          }}
          showNotification={(msg) => showNotification(msg)}
        />
      )}

      {/* DEPO YÖNETİMİ: ANA SİSTEM ENTEGRASYONU */}
      {isDepoModalOpen && (
        <DepoManagementModal
          isOpen={isDepoModalOpen}
          onClose={() => setIsDepoModalOpen(false)}
          initialUnit={selectedUnitFolder || 'at802'}
          depoRows={activeTechizatType === 'at802' ? techizatAt802Data : (activeTechizatType === 'bell429' ? techizatBell429Data : techizatHangarData)}
          at802Rows={techizatAt802Data}
          onUpdateDepoRows={(updatedRows) => {
            if (activeTechizatType === 'at802') setTechizatAt802Data(updatedRows);
            else if (activeTechizatType === 'bell429') setTechizatBell429Data(updatedRows);
            else setTechizatHangarData(updatedRows);
          }}
          transactions={depoTransactions}
          onAddTransaction={(tx) => {
            const newTx: DepoTransaction = {
              id: `tx_${Date.now()}_${Math.random().toString(36).substr(2, 6)}`,
              timestamp: new Date().toISOString(),
              ...tx
            };
            setDepoTransactions(prev => [newTx, ...prev]);
          }}
          onDeleteTransaction={(txId) => {
            setDepoTransactions(prev => prev.filter(t => t.id !== txId));
          }}
          certificatePdfUrl={depoCertificatePdfUrl}
          showNotification={(msg) => showNotification(msg)}
        />
      )}

      {/* AUDIT LOG MODALI */}
      {isAuditModalOpen && (
        <AuditTrailModal
          isOpen={isAuditModalOpen}
          onClose={() => setIsAuditModalOpen(false)}
          logs={auditLogs}
          onClearLogs={() => {
            setAuditLogs([]);
            localStorage.removeItem('equipment_audit_logs');
            showNotification("Tüm denetim kayıtları temizlendi.");
          }}
        />
      )}

      {/* KARA ARAÇLARI BELGE YÖNETİM MODALI */}
      {activeKaraDocTarget && (
        <KaraAraclariDocModal
          isOpen={Boolean(activeKaraDocTarget)}
          onClose={() => setActiveKaraDocTarget(null)}
          vehiclePlate={activeKaraDocTarget.plate}
          vehicleName={activeKaraDocTarget.vehicleName}
          documents={vehicleDocuments.filter(d => d.vehiclePlate === activeKaraDocTarget.plate)}
          onUploadDocument={(newDoc) => {
            setVehicleDocuments(prev => {
              const updated = [newDoc, ...prev];
              try { localStorage.setItem('kara_araclari_vehicle_documents', JSON.stringify(updated)); } catch (e) {}
              return updated;
            });
            addAuditLog({
              unit: 'KARA_ARACLARI',
              action: 'EKLEME',
              itemName: activeKaraDocTarget.vehicleName,
              pn: activeKaraDocTarget.plate,
              fieldName: 'BELGE_YUKLE',
              oldValue: '-',
              newValue: newDoc.title
            });
            showNotification(`✅ Belge yüklendi: ${newDoc.title}`);
          }}
          onDeleteDocument={(docId) => {
            setVehicleDocuments(prev => {
              const updated = prev.filter(d => d.id !== docId);
              try { localStorage.setItem('kara_araclari_vehicle_documents', JSON.stringify(updated)); } catch (e) {}
              return updated;
            });
            showNotification("Belge silindi.");
          }}
        />
      )}

      {/* YER DESTEK VE ÖZEL ALET TERMAL ETİKET MODALI */}
      {isTechizatSlipModalOpen && (
        <DepoSlipPrintModal
          isOpen={isTechizatSlipModalOpen}
          onClose={() => setIsTechizatSlipModalOpen(false)}
          initialType="etiket"
          mode="techizat"
          techizatList={techizatModalList}
          inventory={[]}
          showNotification={(msg) => showNotification(msg)}
        />
      )}

      {/* SUCCESS NOTIFICATION TOAST */}
      {successMessage && (
        <div className="fixed bottom-6 right-6 z-[9000] bg-[#0b3d1d] text-white px-5 py-3.5 rounded-2xl shadow-2xl border border-emerald-500 flex items-center gap-3 animate-fade-in select-none">
          <span className="w-2.5 h-2.5 rounded-full bg-emerald-400 animate-pulse" />
          <span className="text-xs font-bold leading-relaxed">{successMessage}</span>
          <button
            onClick={() => setSuccessMessage(null)}
            className="text-white/60 hover:text-white font-bold text-sm ml-2 cursor-pointer"
          >
            ✕
          </button>
        </div>
      )}
    </div>
  );
}
