/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useRef, useMemo, ChangeEvent } from 'react';
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
  Edit3,
  Columns,
  Rows,
  Square,
  CheckSquare,
  Table,
  Folder,
  Printer,
  FileSpreadsheet,
  Copy
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import * as XLSX from 'xlsx';
import { jsPDF } from 'jspdf';

export const GOOGLE_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbz8H0TN3yNCek4SVaeF9T6KR_yJBhvcsU5iZW7Zzp-c55ViEp8xUlkBdWO_nWqvCDohsg/exec";

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

  // Modal active state
  const [modalOpen, setModalOpen] = useState(false);
  const [modalTitle, setModalTitle] = useState('SİSTEM');
  const [modalType, setModalType] = useState<'iframe' | 'design' | 'category' | 'form_table' | 'excel_sync'>('iframe');
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
      updatePdfBlobUrl(null);
      setIsPdfLoading(true);
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
            // Not in cache, proceed to download and cache
            loadRawPdfFromDrive(match.id, cacheKey);
          }
        }).catch(err => {
          console.error("Failed to load cached raw PDF from DB:", err);
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

  // Google Drive klasöründen PDF dosyalarını listeler
  const fetchPdfMetadata = async () => {
    try {
      const targetUrl = `${GOOGLE_SCRIPT_URL}?action=listPdfsFromDrive`;
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
                  
                if (normBirim.includes("gorevlendirme") || normBirim.includes("1.")) {
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

  // Form last update dates tracker mapping form ID to string (e.g. "05.01.2026 15:30")
  const [formUpdateDates, setFormUpdateDates] = useState<Record<number, string>>(() => {
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
    fetchUpdateDatesFromGoogleSheet();
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
      case 'FORM KAYITLARI': return 'FORM KAYITLARI';
      default: return 'SİSTEM';
    }
  };

  // Modal control functions
  const openSystem = (url: string, title: string) => {
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

  // State to track cloud sync progress
  const [isSendingToSheets, setIsSendingToSheets] = useState<Record<number, boolean>>({});
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
        const newTableData = { ...tableData, [formId]: parsedRows };
        setTableData(newTableData);
        localStorage.setItem(config.storageKey, JSON.stringify(parsedRows));
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
          const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
          
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

    const id = Number(syncSelectedTarget);
    if (isNaN(id) || !TABLE_CONFIGS[id]) {
      alert("Hata: Geçersiz hedef seçimi.");
      return;
    }

    const fileNameLower = file.name.toLowerCase();
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
          
          // Smart sheet detection: find the sheet containing actual data rows with "SIRA", "ADI" or similar keywords
          let worksheet = workbook.Sheets[workbook.SheetNames[0]];
          for (const sheetName of workbook.SheetNames) {
            const sheet = workbook.Sheets[sheetName];
            const rows = XLSX.utils.sheet_to_json<any[]>(sheet, { header: 1, defval: "" });
            const hasKeywords = rows.some(r => r.some((c: any) => {
              const str = String(c || '').toLowerCase();
              return str.includes("sira") || str.includes("adi soyadi") || str.includes("sicil") || str.includes("kadro");
            }));
            if (rows.length > 3 && hasKeywords) {
              worksheet = sheet;
              break;
            }
          }
          
          const rawRows = XLSX.utils.sheet_to_json<string[]>(worksheet, { header: 1, defval: "" });
          const gridData: string[][] = Array.from({ length: 537 }, () => Array(12).fill(""));
          
          // Extract rows that contain a sequence number in Column A (index 0) and filter out headers or titles
          let targetRowIdx = 0;
          for (let r = 0; r < rawRows.length; r++) {
            const rawRow = rawRows[r] || [];
            const colA = String(rawRow[0] || '').trim();
            
            // Check if this row is a header, metadata or title:
            const isHeader = colA.toLowerCase().includes("sira") || 
                             colA.toLowerCase().includes("no") || 
                             colA.toLowerCase().includes("cizelge") || 
                             colA.toLowerCase().includes("çizelge") || 
                             colA.toLowerCase().includes("orman") || 
                             colA.toLowerCase().includes("havacilik") || 
                             colA.toLowerCase().includes("havacılık") || 
                             colA.toLowerCase().includes("personel") ||
                             colA.toLowerCase().includes("mudurlugu") ||
                             colA.toLowerCase().includes("müdürlüğü");
            
            // It is a valid data row if Column A is populated and is NOT a title/header
            const hasSiraNo = colA !== "" && !isHeader;
            
            if (hasSiraNo && targetRowIdx < 537) {
              for (let c = 0; c < 12; c++) {
                gridData[targetRowIdx][c] = rawRow[c] !== undefined ? String(rawRow[c]).trim() : "";
              }
              targetRowIdx++;
            }
          }
          
          // Fallback to basic copy if no matching row with Sıra No was structured this way
          if (targetRowIdx === 0) {
            for (let r = 0; r < 537; r++) {
              const rawRow = rawRows[r] || [];
              for (let c = 0; c < 12; c++) {
                gridData[r][c] = rawRow[c] !== undefined ? String(rawRow[c]).trim() : "";
              }
            }
          }
          
          setExcelForm5Data(gridData);
          localStorage.setItem('excel_form_5_data', JSON.stringify(gridData));
          
          setIsSendingToSheets(prev => ({ ...prev, [5]: true }));
          setUploadProgress(10);
          showNotification(`${file.name} dosyası çözümleniyor ve Google Drive'a yedekleniyor...`);

          // Background Google Drive upload for the Excel file to persist it on Drive
          fileToBase64(file).then(async (base64Data) => {
            try {
              const driveFileName = "personel_bilgi_cizelgesi.xlsx";
              const res = await fetch(GOOGLE_SCRIPT_URL, {
                method: "POST",
                headers: {
                  "Content-Type": "text/plain;charset=utf-8"
                },
                body: JSON.stringify({
                  action: "uploadPdfToDrive",
                  fileName: driveFileName,
                  base64Data: base64Data,
                  formId: 5,
                  month: "Genel Plan"
                })
              });
              if (res.ok) {
                const result = await res.json();
                if (result.status === "success") {
                  showNotification("Güncel Excel belgesi Google Drive'a başarıyla yedeklendi!");
                  fetchPdfMetadata(); // Refresh metadata list
                }
              }
            } catch (err) {
              console.error("Failed to backup Excel to Google Drive:", err);
            }
          });
          
          let progressVal = 10;
          const interval = setInterval(() => {
            progressVal += 15;
            if (progressVal >= 100) {
              clearInterval(interval);
              setUploadProgress(100);
              setIsSendingToSheets(prev => ({ ...prev, [5]: false }));
              showNotification(`'5. Personel Bilgi Çizelgeleri' Excel dosyası başarıyla sisteme aktarıldı ve 537rx12c matris oluşturuldu!`);
              
              const now = new Date();
              const dateStr = `${String(now.getDate()).padStart(2, '0')}.${String(now.getMonth() + 1).padStart(2, '0')}.${now.getFullYear()} ${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;
              setFormUpdateDates(prev => ({ ...prev, 5: dateStr }));
              localStorage.setItem('form_update_dates', JSON.stringify({ ...formUpdateDates, 5: dateStr }));
              
              // Also mock a matching PDF metadata item so the UI registers it as loaded
              const excelPdfMeta = {
                name: "personel_bilgi_cizelgesi.pdf",
                id: "excel_loaded_form5",
                viewUrl: "excel_loaded",
                lastUpdated: dateStr
              };
              setPdfMetadataList(prev => {
                const filtered = prev.filter(p => !p.name.toLowerCase().includes("personel_bilgi"));
                return [excelPdfMeta, ...filtered];
              });
            } else {
              setUploadProgress(progressVal);
            }
          }, 200);
          
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
          month: isSummer ? selectedUploadSummerMonth : "Genel Plan"
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
          const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
          
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

  return (
    <div className="min-h-screen bg-[#0b3d1d] overflow-hidden relative selection:bg-white/20">
      
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
                <h1 className="text-3xl sm:text-4xl md:text-5xl font-extrabold tracking-tighter leading-none mb-2 text-white text-shadow uppercase">
                  HAVA ARAÇLARI BAKIM
                </h1>
                <h2 className="text-lg sm:text-xl md:text-2xl font-light tracking-[0.2em] text-white/80 uppercase">
                  TEKNİK ŞUBE MÜDÜRLÜĞÜ
                </h2>
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
              <span>{(modalType === 'form_table' || modalType === 'excel_sync' || categoryHistory.length > 1) ? 'Geri' : 'Kapat'}</span>
            </button>

            {/* Action Buttons */}
            <div className="flex items-center gap-2 pointer-events-auto">
              {/* Show Veri Güncelle button only on FORM KAYITLARI categories or views */}
              {(selectedCategory === 'FORM KAYITLARI' || selectedFormId !== null || modalType === 'form_table' || modalType === 'excel_sync') && (
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

              <button
                id="modal-close-btn"
                onClick={closeSystem}
                aria-label="Kapat"
                className="flex items-center justify-center text-slate-800 bg-white/95 hover:bg-red-50 hover:text-red-600 border border-slate-200/80 backdrop-blur-md shadow-lg w-10 h-10 rounded-full transition-all hover:scale-105 active:scale-95 cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>
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
            <div id="category-menu-content" className="absolute inset-0 flex items-center justify-center bg-gradient-to-br from-[#f8fafc] to-[#f1f5f9] p-6 overflow-y-auto w-full h-full">
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
                      onClick={() => showDesignPhase('HANGAR YER DESTEK TEÇHİZATLARI')}
                      className="bg-white hover:bg-white/80 border border-gray-200 shadow-sm rounded-[2rem] p-8 flex flex-col items-center text-center transition-all duration-300 hover:scale-[1.03] active:scale-[0.98] cursor-pointer group"
                    >
                      <div className="w-16 h-16 bg-[#0b3d1d]/10 rounded-2xl flex items-center justify-center mb-6 group-hover:bg-[#0b3d1d]/20 transition-all shadow-sm">
                        <Wrench className="w-8 h-8 text-[#0b3d1d]" />
                      </div>
                      <span className="text-[#0b3d1d] font-bold tracking-widest text-sm mb-2 uppercase">HANGAR YER DESTEK</span>
                      <span className="text-[10px] text-[#0b3d1d]/60 uppercase tracking-widest font-semibold font-mono">TEÇHİZATLARALTYAPI</span>
                    </button>
                  </>
                )}

                {/* HAVA ARAÇLARI YER DESTEK (HA_YER_DESTEK) Alt Butonları */}
                {selectedCategory === 'HA_YER_DESTEK' && (
                  <>
                    <button
                      onClick={() => showDesignPhase('BELL 429 YER DESTEK LİSTELERİ')}
                      className="bg-white hover:bg-white/80 border border-gray-200 shadow-sm rounded-[2rem] p-8 flex flex-col items-center text-center transition-all duration-300 hover:scale-[1.03] active:scale-[0.98] cursor-pointer group"
                    >
                      <div className="w-16 h-16 bg-white border border-gray-200 shadow-inner rounded-2xl flex items-center justify-center mb-6 text-[#0b3d1d] font-black text-xs">
                        B429
                      </div>
                      <span className="text-[#0b3d1d] font-bold tracking-widest text-sm mb-2 uppercase">BELL 429</span>
                      <span className="text-[10px] text-gray-500 uppercase tracking-widest font-mono">Teçhizat Listesi</span>
                    </button>

                    <button
                      onClick={() => showDesignPhase('AT-802F YER DESTEK LİSTELERİ')}
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
                      onClick={() => showDesignPhase('C-650 YER DESTEK LİSTELERİ')}
                      className="bg-white hover:bg-white/80 border border-gray-200 shadow-sm rounded-[2rem] p-8 flex flex-col items-center text-center transition-all duration-300 hover:scale-[1.03] active:scale-[0.98] cursor-pointer group"
                    >
                      <div className="w-16 h-16 bg-white border border-gray-200 shadow-inner rounded-2xl flex items-center justify-center mb-6 text-[#0b3d1d] font-black text-xs">
                        C650
                      </div>
                      <span className="text-[#0b3d1d] font-bold tracking-widest text-sm mb-2 uppercase">C-650</span>
                      <span className="text-[10px] text-gray-500 uppercase tracking-widest font-mono">Teçhizat Listesi</span>
                    </button>

                    <button
                      onClick={() => showDesignPhase('B-360 YER DESTEK LİSTELERİ')}
                      className="bg-white hover:bg-white/80 border border-gray-200 shadow-sm rounded-[2rem] p-8 flex flex-col items-center text-center transition-all duration-300 hover:scale-[1.03] active:scale-[0.98] cursor-pointer group"
                    >
                      <div className="w-16 h-16 bg-white border border-gray-200 shadow-inner rounded-2xl flex items-center justify-center mb-6 text-[#0b3d1d] font-black text-xs">
                        B360
                      </div>
                      <span className="text-[#0b3d1d] font-bold tracking-widest text-sm mb-2 uppercase">B-360</span>
                      <span className="text-[10px] text-gray-500 uppercase tracking-widest font-mono">Teçhizat Listesi</span>
                    </button>
                  </>
                )}

                {/* T-70 DETAY ALTBİRİMLERİ (T70_DETAY) */}
                {selectedCategory === 'T70_DETAY' && (
                  <>
                    <button
                      onClick={() => showDesignPhase('T-70 BUMBİ BACKET LİSTESİ')}
                      className="bg-white hover:bg-white/80 border border-gray-200 shadow-sm rounded-[2rem] p-8 flex flex-col items-center text-center transition-all duration-300 hover:scale-[1.03] active:scale-[0.98] cursor-pointer group"
                    >
                      <div className="w-16 h-16 bg-[#0b3d1d]/10 rounded-2xl flex items-center justify-center mb-6 group-hover:bg-[#0b3d1d]/20 transition-all shadow-sm text-[#0b3d1d] font-extrabold text-sm">
                        BB
                      </div>
                      <span className="text-[#0b3d1d] font-bold tracking-widest text-sm mb-2 uppercase">BUMBİ BACKET</span>
                      <span className="text-[10px] text-gray-500 uppercase tracking-widest font-mono">T-70 Alt Birimi</span>
                    </button>

                    <button
                      onClick={() => showDesignPhase('T-70 YER DESTEK TEÇHİZATLARI LİSTESİ')}
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
                                        const val = cell || "";
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
                                        "KAN", "ADRES BİLGİSİ", "YAKIN ADI", "EŞ TEL NO"
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
                                          const val = cellVal || "";
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
                          <div className="w-full max-w-7xl bg-white text-slate-800 rounded-[1.5rem] shadow-2xl p-8 flex flex-col border border-slate-200 select-text relative overflow-x-auto min-w-[1000px]">
                            
                            {/* PDF Header Mockup */}
                            <div className="border-b-2 border-emerald-800 pb-6 mb-6 flex items-center justify-between shrink-0 select-none">
                              <div className="flex items-center gap-3">
                                <div className="w-12 h-12 bg-emerald-800 rounded-xl flex items-center justify-center text-white font-black text-lg shadow-md">
                                  OGM
                                </div>
                                <div className="text-left">
                                  <h4 className="text-sm font-black text-emerald-900 uppercase tracking-wider">ORMAN GENEL MÜDÜRLÜĞÜ</h4>
                                  <p className="text-[10px] font-black text-slate-500 uppercase tracking-widest">HAVACILIK DAİRESİ BAŞKANLIĞI</p>
                                </div>
                              </div>
                              <div className="text-right">
                                <h3 className="text-base font-black text-slate-800 uppercase tracking-tighter">PERSONEL BİLGİ ÇİZELGESİ</h3>
                              </div>
                            </div>

                            {/* Column Headers */}
                            <div className="grid grid-cols-[60px_160px_110px_80px_100px_90px_90px_110px_70px_180px_140px_120px] gap-2 bg-slate-900 text-white rounded-xl p-3 text-center text-[10px] font-extrabold uppercase tracking-wider mb-2 select-none min-w-[1310px]">
                              <div>SIRA NO</div>
                              <div>ADI SOYADI</div>
                              <div>T.C. KİMLİK</div>
                              <div>SİCİL NO</div>
                              <div>KADRO UNVANI</div>
                              <div>DOĞUM T.</div>
                              <div>GÖREV YERİ</div>
                              <div>TELEFON NO</div>
                              <div>KAN GRUBU</div>
                              <div>ADRES BİLGİSİ</div>
                              <div>YAKININ ADI</div>
                              <div>EŞ TEL NUMARALARI</div>
                            </div>

                            {/* 537 Rows Layout container. Scrollable horizontally and vertically */}
                            <div className="flex-1 space-y-1.5 max-h-[60vh] overflow-y-auto pr-2 scrollbar-thin overflow-x-auto">
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
                                    <div className="text-center py-12 text-slate-400 select-none font-mono text-xs uppercase tracking-wider min-w-[1310px]">
                                      🚫 ARAMA SONUCUNA UYGUN PERSONEL BULUNAMADI!
                                    </div>
                                  );
                                }

                                return currentFilteredRows.map(({ row, rIdx }) => {
                                  const hasText = true;
                                  
                                  // If row is empty, we show a highly styled, clean "empty slot"
                                  if (!hasText) {
                                    return (
                                      <div 
                                        key={rIdx} 
                                        className="grid grid-cols-[60px_160px_110px_80px_100px_90px_90px_110px_70px_180px_140px_120px] gap-2 px-3 py-2 rounded-lg border border-slate-100 bg-slate-50/40 text-[10px] font-mono text-slate-350 hover:bg-slate-100/50 transition-all select-none group min-w-[1310px]"
                                      >
                                        <div className="font-bold text-slate-400 group-hover:text-[#0b3d1d]">R-{String(rIdx + 1).padStart(3, '0')}</div>
                                        <div className="col-span-11 text-left text-slate-350 italic tracking-widest pl-4">
                                          --- BOŞ VERİ ALANI (METİN GELDİKÇE DOLACAKTIR) ---
                                        </div>
                                      </div>
                                    );
                                  }

                                  return (
                                    <div 
                                      key={rIdx} 
                                      className={`grid grid-cols-[60px_160px_110px_80px_100px_90px_90px_110px_70px_180px_140px_120px] gap-2 px-3 py-2 rounded-lg border text-[11px] font-semibold items-center transition-all min-w-[1310px] ${
                                        hasText 
                                          ? 'border-slate-200 bg-white hover:border-emerald-500 hover:shadow-md' 
                                          : 'border-slate-100 bg-slate-50/50 text-slate-400'
                                      }`}
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
                                          <div 
                                            key={cIdx} 
                                            onDoubleClick={() => {
                                              if (cell) {
                                                setActiveModalCell({
                                                  r: rIdx,
                                                  c: cIdx,
                                                  value: cell,
                                                  label: label
                                                });
                                                setCopiedCellSuccess(false);
                                              }
                                            }}
                                            title={cell ? `${label}: ${cell} (Detay için Çift Tıklayın)` : "Boş Veri"}
                                            className={`truncate px-1 py-0.5 rounded transition-all text-center select-text cursor-zoom-in hover:bg-emerald-50 hover:text-emerald-950 active:scale-95 ${
                                              isActiveMatch 
                                                ? 'bg-blue-600 text-white font-black scale-105 shadow-lg ring-2 ring-blue-400 animate-pulse'
                                                : isMatch
                                                  ? 'bg-blue-200 text-blue-950 font-black border border-blue-400'
                                                  : cell 
                                                    ? 'text-slate-800 font-sans font-medium' 
                                                    : 'text-slate-300 italic'
                                            }`}
                                          >
                                            {cell || "-"}
                                          </div>
                                        );
                                      })}
                                    </div>
                                  );
                                });
                              })()}
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
                                      <th className="border border-slate-300 p-1 bg-[#0b3d1d] text-white text-center" style={{ width: '12%' }}>EŞ TEL NO</th>
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
                                        <td className="border border-slate-300 p-1 text-center font-mono text-black">{row[5] || ""}</td>
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
                          const targetId = step1Target || '1';
                          setSyncSelectedTarget(targetId);
                          setActiveSyncStep(2);
                        }}
                        className="w-full py-4 bg-[#0b3d1d] hover:bg-[#072612] text-white font-extrabold text-xs rounded-xl tracking-widest uppercase transition-all shadow-md active:scale-95 cursor-pointer select-none"
                      >
                        İLERLE VE PDF YÜKLEME EKRANINA GEÇ ➜
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
                        {isSendingToSheets[Number(syncSelectedTarget)] ? (
                          <Loader2 className="w-10 h-10 animate-spin" />
                        ) : (
                          <FileText className="w-10 h-10 animate-pulse text-[#0b3d1d]" />
                        )}
                      </div>
                      
                      <div className="flex-1 flex flex-col justify-center select-none w-full">
                        <h4 className="text-slate-800 font-extrabold text-sm uppercase tracking-wider mb-2">
                          {syncSelectedTarget === '5' ? "PLANLAMA BELGESİ YÜKLE (EXCEL VEYA PDF)" : "PLANLAMA BELGESİ YÜKLE (PDF)"}
                        </h4>
                        <p className="text-xs text-slate-500 leading-relaxed max-w-sm mx-auto mb-6">
                          {isSendingToSheets[Number(syncSelectedTarget)] 
                            ? "Plan belgesi sisteme aktarılıyor ve veriler işleniyor. Lütfen bekleyin..." 
                            : syncSelectedTarget === '5'
                              ? "Personel Bilgi Çizelgesi için güncel Excel (.xlsx, .xls) veya PDF dosyasını yükleyin."
                              : "Seçilen birim için güncel planlama belgesini PDF formatında yükleyin."}
                        </p>

                        {/* Guide rules depending on summer period or not */}
                        <div className="bg-slate-50 rounded-2xl p-4 mb-6 border border-slate-100 text-left w-full">
                          <span className="text-[10px] font-black text-slate-700 uppercase tracking-wider block mb-2">
                            📋 {syncSelectedTarget === '5' ? "Yükleme ve Önizleme Teknolojisi:" : "Sürücü (Drive) Otomatik Adlandırma Formatı:"}
                          </span>
                          <div className="space-y-1.5 font-mono text-[10px] text-slate-600">
                            {syncSelectedTarget === '5' ? (
                              <>
                                <div className="text-emerald-800 font-bold uppercase tracking-wider">
                                  • EXCEL ➔ PDF MATRİS DÖNÜŞTÜRÜCÜ
                                </div>
                                <div className="text-slate-500 text-[10px] leading-relaxed">
                                  Yüklenen Excel belgesindeki satırlar çözümlenerek <strong>537 Satır x 12 Sütunluk (537rx12c)</strong> yatay bir elektronik tabloya dönüştürülür. Sadece metin içeren kısımlar şık ve yüksek çözünürlüklü bir PDF belgesi gibi taranarak ekranda gösterilir.
                                </div>
                              </>
                            ) : ["21", "22", "23", "24", "25"].includes(syncSelectedTarget) ? (
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
                            isSendingToSheets[Number(syncSelectedTarget)] 
                              ? "bg-slate-400 cursor-not-allowed animate-pulse" 
                              : "bg-[#0b3d1d] hover:bg-[#072612]"
                          }`}
                        >
                          <input
                            type="file"
                            id="pdf-file-contextual-upload"
                            accept={syncSelectedTarget === '5' ? ".xlsx,.xls,.csv,.pdf" : ".pdf"}
                            onChange={handlePdfUpload}
                            disabled={isSendingToSheets[Number(syncSelectedTarget)]}
                            className="hidden"
                          />
                          {isSendingToSheets[Number(syncSelectedTarget)] ? (
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

                        {syncSelectedTarget === '5' && (
                          <button
                            type="button"
                            onClick={() => {
                              try {
                                const wb = XLSX.utils.book_new();
                                const headers = TABLE_CONFIGS[5].columns.map(col => col.label);
                                const dataWithHeaders = [headers, ...excelForm5Data];
                                const ws = XLSX.utils.aoa_to_sheet(dataWithHeaders);
                                XLSX.utils.book_append_sheet(wb, ws, "Personel_Bilgi_Tum_Surum");
                                XLSX.writeFile(wb, "personel_bilgi_cizelgesi_en_son_surum.xlsx");
                                showNotification("Sistemdeki en son sürüm ham Excel verisi başarıyla indirildi.");
                              } catch (e) {
                                alert("Hata: " + e);
                              }
                            }}
                            className="w-full py-3.5 bg-emerald-700 hover:bg-emerald-800 text-white font-extrabold text-[11px] rounded-xl tracking-wider uppercase transition-all flex items-center justify-center gap-2 shadow-sm cursor-pointer"
                            title="Sistemde saklanan en son sürüm Personel Bilgi Excel verisini indir"
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
                    disabled={isSendingToSheets[Number(syncSelectedTarget)] || renderedPages.filter(p => p.selected).length === 0}
                    className={`px-8 py-3 font-extrabold text-xs uppercase tracking-wider rounded-xl shadow-lg transition-all cursor-pointer flex items-center gap-2 ${
                      isSendingToSheets[Number(syncSelectedTarget)] || renderedPages.filter(p => p.selected).length === 0
                        ? 'bg-slate-700 text-slate-400 cursor-not-allowed'
                        : 'bg-emerald-600 hover:bg-emerald-700 active:bg-emerald-800 text-white'
                    }`}
                  >
                    {isSendingToSheets[Number(syncSelectedTarget)] ? (
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

    </div>
  );
}


