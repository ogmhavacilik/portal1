import React, { useState, useEffect, useMemo, useRef } from 'react';
import {
  GraduationCap,
  Search,
  X,
  Download,
  Plus,
  Upload,
  ExternalLink,
  CheckCircle2,
  FileCheck,
  FilePlus,
  RefreshCw,
  Building2,
  Calendar,
  User,
  BookOpen,
  Trash2,
  Edit2,
  Check,
  FileSpreadsheet,
  ChevronDown,
  ChevronUp,
  ArrowUpDown,
  ArrowUp,
  ArrowDown,
  Mail,
  AlertTriangle,
  Lock,
  ShieldCheck,
  Clock
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { DEFAULT_TEKNISYEN_PERSONEL_NAMES } from '../utils/personnelData';
import { cleanAndFormatDateString } from '../utils/driveExcelSync';

export interface TrainingItem {
  id: string;
  egitimAdi: string;
  egitimTarihi: string;
  egitimVeren: string;
}

export interface PersonelEgitimPerson {
  id: string;
  siraNo: number;
  adiSoyadi: string;
  trainings: TrainingItem[];
  isgEgitimTarihi?: string;
  cbikoTarihi?: string;
  kadroUnvani?: string;
  hasPdf: boolean;
  linkTitle?: string;
}

const STORAGE_KEY = 'ogm_personel_egitim_listesi_v2';
const DRIVE_FILE_NAME = 'personel_egitim_listesi.xlsx';
const REQUIRED_PIN = '1839';

// Date parser helper to JS Date
function parseTrDate(dateStr?: string): Date | null {
  if (!dateStr || dateStr === '-' || dateStr.trim() === '') return null;
  const parts = dateStr.trim().split(/[\.\/-]/);
  if (parts.length === 3) {
    const d = parseInt(parts[0], 10);
    const m = parseInt(parts[1], 10) - 1;
    let y = parseInt(parts[2], 10);
    if (y < 100) y += 2000;
    if (!isNaN(d) && !isNaN(m) && !isNaN(y)) {
      return new Date(y, m, d);
    }
  }
  return null;
}

const TR_MONTH_MAP: Record<string, number> = {
  OCAK: 0,
  SUBAT: 1,
  MART: 2,
  NISAN: 3,
  MAYIS: 4,
  HAZIRAN: 5,
  TEMMUZ: 6,
  AGUSTOS: 7,
  EYLUL: 8,
  EKIM: 9,
  KASIM: 10,
  ARALIK: 11
};

// Extract the latest valid timestamp from free-text Turkish training date strings
// e.g. "06 -24 MART 2023 (TEORİK) 08-19 MAYIS 2023 (PRATİK)", "31 AĞUSTOS-10 EYLÜL 2021", "15.04.2024"
export function getTrainingSortTimestamp(dateStr?: string): number {
  if (!dateStr || dateStr === '-' || !dateStr.trim()) return 0;
  let maxTs = 0;

  // 1) Check numeric dates dd.mm.yyyy / dd/mm/yyyy / dd-mm-yyyy
  const numMatches = dateStr.matchAll(/(\d{1,2})[\.\/-](\d{1,2})[\.\/-](\d{2,4})/g);
  for (const m of numMatches) {
    const d = parseInt(m[1], 10);
    const mo = parseInt(m[2], 10) - 1;
    let y = parseInt(m[3], 10);
    if (y < 100) y += 2000;
    if (d >= 1 && d <= 31 && mo >= 0 && mo <= 11 && y >= 1950 && y <= 2100) {
      const ts = new Date(y, mo, d).getTime();
      if (ts > maxTs) maxTs = ts;
    }
  }

  // 2) Normalize Turkish characters & check "<days> <MONTH> <YEAR>"
  const norm = dateStr
    .toLocaleUpperCase('tr-TR')
    .replace(/Ş/g, 'S')
    .replace(/Ğ/g, 'G')
    .replace(/Ü/g, 'U')
    .replace(/Ö/g, 'O')
    .replace(/Ç/g, 'C')
    .replace(/İ/g, 'I');

  const trMonthRegex = /(?:(\d{1,2})(?:\s*[-–]\s*(\d{1,2}))?\s+)?(OCAK|SUBAT|MART|NISAN|MAYIS|HAZIRAN|TEMMUZ|AGUSTOS|EYLUL|EKIM|KASIM|ARALIK)\s+(19\d{2}|20\d{2})/g;
  for (const m of norm.matchAll(trMonthRegex)) {
    const day = parseInt(m[2] || m[1] || '1', 10);
    const month = TR_MONTH_MAP[m[3]];
    const year = parseInt(m[4], 10);
    if (month !== undefined && !isNaN(year)) {
      const ts = new Date(year, month, Math.max(1, Math.min(31, day || 1))).getTime();
      if (ts > maxTs) maxTs = ts;
    }
  }

  // 3) Fallback: 4-digit year anywhere in text
  if (maxTs === 0) {
    for (const m of norm.matchAll(/\b(19\d{2}|20\d{2})\b/g)) {
      const year = parseInt(m[1], 10);
      const ts = new Date(year, 0, 1).getTime();
      if (ts > maxTs) maxTs = ts;
    }
  }

  return maxTs;
}

function sortTrainingsNewestFirst(list: TrainingItem[]): TrainingItem[] {
  return [...list].sort((a, b) => getTrainingSortTimestamp(b.egitimTarihi) - getTrainingSortTimestamp(a.egitimTarihi));
}

// Check expiration status for ISG (3 yrs) or Muayene (5 yrs)
export function getExpiryStatus(dateStr?: string, yearsValid: number = 3) {
  if (!dateStr || dateStr === '-' || dateStr.trim() === '') {
    return { hasDate: false, isExpired: false, isWarning: false, daysLeft: 9999, expiryDateStr: '-' };
  }
  const dateObj = parseTrDate(dateStr);
  if (!dateObj) {
    return { hasDate: false, isExpired: false, isWarning: false, daysLeft: 9999, expiryDateStr: '-' };
  }

  const expiryDate = new Date(dateObj);
  expiryDate.setFullYear(expiryDate.getFullYear() + yearsValid);

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const diffTime = expiryDate.getTime() - today.getTime();
  const daysLeft = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

  const expiryDateStr = `${String(expiryDate.getDate()).padStart(2, '0')}.${String(expiryDate.getMonth() + 1).padStart(2, '0')}.${expiryDate.getFullYear()}`;

  const isExpired = daysLeft <= 0;
  const isWarning = daysLeft > 0 && daysLeft <= 30; // Son 1 ay

  return {
    hasDate: true,
    isExpired,
    isWarning,
    daysLeft,
    expiryDateStr
  };
}

// Calculate 1st day of next month for email scheduling
export function getNextMonthFirstDayStr(): string {
  const today = new Date();
  let nextMonth = today.getMonth() + 1;
  let year = today.getFullYear();
  if (nextMonth > 11) {
    nextMonth = 0;
    year += 1;
  }
  const firstDay = new Date(year, nextMonth, 1);
  return `${String(firstDay.getDate()).padStart(2, '0')}.${String(firstDay.getMonth() + 1).padStart(2, '0')}.${firstDay.getFullYear()}`;
}

interface PersonelEgitimModalProps {
  isOpen: boolean;
  onClose: () => void;
  showNotification?: (message: string) => void;
}

export const PersonelEgitimModal: React.FC<PersonelEgitimModalProps> = ({
  isOpen,
  onClose,
  showNotification
}) => {
  const [people, setPeople] = useState<PersonelEgitimPerson[]>([]);
  const [isLoadingDrive, setIsLoadingDrive] = useState(false);
  const [syncStatusMessage, setSyncStatusMessage] = useState('');
  const [searchQuery, setSearchQuery] = useState('');
  const [activeMatchIdx, setActiveMatchIdx] = useState(0);

  // Sorting
  type SortKey = 'siraNo' | 'adiSoyadi' | 'egitimTarihi' | 'isgEgitimTarihi' | 'cbikoTarihi';
  const [sortKey, setSortKey] = useState<SortKey>('siraNo');
  const [sortDir, setSortDir] = useState<'asc' | 'desc'>('asc');

  // Test Mail / Password Modal State (1839)
  const [isTestMailModalOpen, setIsTestMailModalOpen] = useState(false);
  const [testMailPinInput, setTestMailPinInput] = useState('');
  const [testMailPinError, setTestMailPinError] = useState(false);
  const [testMailAuthorized, setTestMailAuthorized] = useState(false);
  const [isSimulatingMail, setIsSimulatingMail] = useState(false);

  // New Person Modal
  const [isAddPersonModalOpen, setIsAddPersonModalOpen] = useState(false);
  const [newPersonName, setNewPersonName] = useState('');
  const [newPersonIsgDate, setNewPersonIsgDate] = useState('');
  const [newPersonCbikoDate, setNewPersonCbikoDate] = useState('');
  const [newInitialTrainingName, setNewInitialTrainingName] = useState('');
  const [newInitialTrainingDate, setNewInitialTrainingDate] = useState(() => new Date().toLocaleDateString('tr-TR'));
  const [newInitialTrainingProvider, setNewInitialTrainingProvider] = useState('');
  const [activeProviderField, setActiveProviderField] = useState<'inline' | 'new' | 'edit' | null>(null);

  // Mevcut tüm personellerin eğitim verilerinden benzersiz "Eğitim Veren Kurum" listesi çıkar
  const uniqueProviders = useMemo(() => {
    const map = new Map<string, string>();
    people.forEach(p => {
      (p.trainings || []).forEach(t => {
        const raw = String(t.egitimVeren || '').trim();
        if (!raw || raw === '-' || raw === '--') return;
        const key = raw.toLocaleUpperCase('tr-TR');
        if (!map.has(key)) {
          map.set(key, raw);
        }
      });
    });
    return Array.from(map.values()).sort((a, b) => a.localeCompare(b, 'tr'));
  }, [people]);

  const getFilteredProviders = (query: string) => {
    const q = query.trim().toLocaleLowerCase('tr-TR');
    if (!q) return uniqueProviders.slice(0, 12);
    return uniqueProviders
      .filter(p => p.toLocaleLowerCase('tr-TR').includes(q))
      .slice(0, 12);
  };

  // Inline "Eğitim Ekle" State for a specific person
  const [inlineAddingPersonId, setInlineAddingPersonId] = useState<string | null>(null);
  const [inlineTrainingName, setInlineTrainingName] = useState('');
  const [inlineTrainingDate, setInlineTrainingDate] = useState(() => new Date().toLocaleDateString('tr-TR'));
  const [inlineTrainingProvider, setInlineTrainingProvider] = useState('');

  // Edit Training Modal State
  const [editingTrainingInfo, setEditingTrainingInfo] = useState<{ personId: string; training: TrainingItem } | null>(null);
  const [editTrainingName, setEditTrainingName] = useState('');
  const [editTrainingDate, setEditTrainingDate] = useState('');
  const [editTrainingProvider, setEditTrainingProvider] = useState('');

  // Edit Person Modal State
  const [editingPersonInfo, setEditingPersonInfo] = useState<{
    personId: string;
    siraNo: number;
    adiSoyadi: string;
    isgEgitimTarihi?: string;
    cbikoTarihi?: string;
  } | null>(null);
  const [editPersonSiraNo, setEditPersonSiraNo] = useState<number>(1);
  const [editPersonName, setEditPersonName] = useState('');
  const [editPersonIsgDate, setEditPersonIsgDate] = useState('');
  const [editPersonCbikoDate, setEditPersonCbikoDate] = useState('');

  // Prompt state for "PDF Yüklediniz mi?"
  const [promptPersonId, setPromptPersonId] = useState<string | null>(null);
  const [isPdfConfirmedChecked, setIsPdfConfirmedChecked] = useState<boolean>(true);

  // Accordion state
  const [expandedPersonIds, setExpandedPersonIds] = useState<Set<string>>(new Set());

  const toggleExpandPerson = (personId: string) => {
    setExpandedPersonIds(prev => {
      const next = new Set(prev);
      if (next.has(personId)) next.delete(personId);
      else next.add(personId);
      return next;
    });
  };

  const expandAll = () => setExpandedPersonIds(new Set(people.map(p => p.id)));
  const collapseAll = () => setExpandedPersonIds(new Set());

  const fileInputRef = useRef<HTMLInputElement>(null);

  // Fetch live Excel on opening
  const fetchLiveExcelFromDrive = async (isManual = false, customMsg = '') => {
    if (isManual || people.length === 0) {
      setIsLoadingDrive(true);
      setSyncStatusMessage(customMsg || (isManual ? "Bulut Arşivi Exceli Aranıyor..." : "Veriler Yükleniyor..."));
      if (showNotification && isManual) {
        showNotification("🔄 Bulut Arşivindeki Personel Eğitim Excel dosyası taranıyor...");
      }
    }

    try {
      // Direct GAS fetch
      const directGasUrls = [
        "https://script.google.com/macros/s/AKfycbwP1uOo2NrST5a4I8vm1nGBLtI26yY2lWrmu9_e9iymwBkUhJBA9JOPCp7SNKqJbOubOw/exec",
        "https://script.google.com/macros/s/AKfycbw4kruTTc058Y9rLTyO3dKi6KloYsmdDTwV1GSiAk8ZXefyo3Z7_VDSTuurzsS9BHAQyQ/exec"
      ];

      const targets = [
        DRIVE_FILE_NAME,
        "personel_egitim_listesi_at-802.xlsx"
      ];

      for (const gasUrl of directGasUrls) {
        for (const targetName of targets) {
          try {
            const gasRes = await fetch(gasUrl, {
              method: 'POST',
              headers: { 'Content-Type': 'text/plain;charset=utf-8' },
              body: JSON.stringify({
                action: 'readExcelFromDrive',
                fileName: targetName,
                folderId: '1HQR_NYKhHQGA7_2W3nArI9pCh-LJasTP'
              })
            });

            if (gasRes.ok) {
              const gasJson = await gasRes.json();
              if (gasJson && gasJson.status === 'success' && gasJson.base64) {
                const parsed = parseExcelAndSetPeople(gasJson.base64, false);
                if (parsed) {
                  if (showNotification) {
                    showNotification("✅ Bulut Arşivinden Personel Eğitim Listesi canlı olarak yüklendi.");
                  }
                  return;
                }
              }
            }
          } catch (gasErr) {
            console.warn("Direct GAS fetch warning:", gasErr);
          }
        }
      }
    } catch (err) {
      console.warn("Fetch live Excel warning:", err);
    } finally {
      setIsLoadingDrive(false);
      setSyncStatusMessage('');
    }
  };

  useEffect(() => {
    if (isOpen) {
      setTimeout(() => {
        fetchLiveExcelFromDrive(false);
      }, 50);
    }
  }, [isOpen]);

  // Parse Excel Base64
  const parseExcelAndSetPeople = (base64: string, syncToDrive = false): boolean => {
    try {
      const binary = atob(base64);
      const len = binary.length;
      const bytes = new Uint8Array(len);
      for (let i = 0; i < len; i++) {
        bytes[i] = binary.charCodeAt(i);
      }
      const wb = XLSX.read(bytes, { type: 'array' });
      const firstSheetName = wb.SheetNames[0];
      const ws = wb.Sheets[firstSheetName];
      if (!ws) return false;

      const rawRows: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
      if (!rawRows || rawRows.length === 0) return false;

      let headerRowIdx = -1;
      for (let r = 0; r < Math.min(15, rawRows.length); r++) {
        const rowStr = (rawRows[r] || []).map(c => String(c || '').toUpperCase().trim()).join(' ');
        if (rowStr.includes('ADI') || rowStr.includes('PERSONEL') || rowStr.includes('EĞİTİM') || rowStr.includes('MESLEKİ')) {
          headerRowIdx = r;
          break;
        }
      }
      if (headerRowIdx < 0) headerRowIdx = 0;

      const normHeader = (rawRows[headerRowIdx] || []).map(c =>
        String(c || '').toUpperCase().trim()
          .replace(/Ç/g, 'C').replace(/Ğ/g, 'G').replace(/İ/g, 'I').replace(/I/g, 'I').replace(/Ö/g, 'O').replace(/Ş/g, 'S').replace(/Ü/g, 'U')
      );

      let siraCol = normHeader.findIndex(c => c.includes('SIRA') || c === 'NO');
      let adCol = normHeader.findIndex(c => c.includes('ADI') || c.includes('SOYADI') || c.includes('PERSONEL'));
      let egtCol = normHeader.findIndex(c => c.includes('MESLEKI') || c.includes('EGITIM') || c.includes('KURS'));
      let tarCol = normHeader.findIndex(c => c.includes('EGITIM TARIHI') || (c.includes('TARIH') && !c.includes('ISG') && !c.includes('CBIKO') && !c.includes('MUAYENE')));
      let isgCol = normHeader.findIndex(c => c.includes('ISG'));
      let cbikoCol = normHeader.findIndex(c => c.includes('CBIKO') || c.includes('GIRIS'));
      let muayeneCol = normHeader.findIndex(c => c.includes('MUAYENE') || c.includes('PERIYODIK'));
      let kurumCol = normHeader.findIndex(c => c.includes('KURUM') || c.includes('VEREN') || c.includes('EGITMEN'));
      let linkCol = normHeader.findIndex(c => c.includes('LINK') || c.includes('BELGE') || c.includes('DURUM'));

      if (siraCol < 0) siraCol = 0;
      if (adCol < 0) adCol = 1;
      if (egtCol < 0) egtCol = 2;
      if (tarCol < 0) tarCol = 3;

      const parsedPeople: PersonelEgitimPerson[] = [];
      let lastPerson: PersonelEgitimPerson | null = null;

      for (let r = headerRowIdx + 1; r < rawRows.length; r++) {
        const row = rawRows[r];
        if (!row || row.length === 0) continue;

        const rawSira = String(siraCol >= 0 && row[siraCol] !== undefined ? row[siraCol] : row[0] || '').trim();
        const rawAd = String(adCol >= 0 && row[adCol] !== undefined ? row[adCol] : row[1] || '').trim();
        const rawEgt = String(egtCol >= 0 && row[egtCol] !== undefined ? row[egtCol] : row[2] || '').trim();
        const rawTar = String(tarCol >= 0 && row[tarCol] !== undefined ? row[tarCol] : row[3] || '').trim();
        const rawIsg = String(isgCol >= 0 && row[isgCol] !== undefined ? row[isgCol] : '').trim();
        const rawCbiko = String(cbikoCol >= 0 && row[cbikoCol] !== undefined ? row[cbikoCol] : '').trim();
        const rawKurum = String(kurumCol >= 0 && row[kurumCol] !== undefined ? row[kurumCol] : '').trim();
        const rawLink = String(linkCol >= 0 && row[linkCol] !== undefined ? row[linkCol] : '').trim();

        const adUpper = rawAd.toUpperCase();
        let person: PersonelEgitimPerson | null = null;

        if (rawAd && adUpper !== 'ADI SOYADI' && adUpper !== 'PERSONEL' && adUpper !== 'ADI' && !adUpper.includes('ORMAN GENEL')) {
          person = parsedPeople.find(p => p.adiSoyadi.toUpperCase() === adUpper) || null;
          if (!person) {
            const numSira = parseInt(rawSira) || (parsedPeople.length + 1);
            const hasPdf = rawLink.toUpperCase().includes('GÖR') || rawLink.toUpperCase().includes('GOR') || rawLink.toUpperCase().includes('EVET');
            person = {
              id: `person-${Date.now()}-${parsedPeople.length + 1}`,
              siraNo: numSira,
              adiSoyadi: adUpper,
              isgEgitimTarihi: cleanAndFormatDateString(rawIsg) || rawIsg || '-',
              cbikoTarihi: cleanAndFormatDateString(rawCbiko) || rawCbiko || '-',
              trainings: [],
              hasPdf,
              linkTitle: rawLink || `${adUpper} EĞİTİM SERTİFİKALARI`
            };
            parsedPeople.push(person);
          }
          lastPerson = person;
        } else {
          person = lastPerson;
        }

        if (person && rawEgt && rawEgt !== '-' && !rawEgt.toUpperCase().includes('MESLEKİ EĞİTİM') && !rawEgt.toUpperCase().includes('EGITIMIN ADI')) {
          const egtUpper = rawEgt.toUpperCase();
          const egtTarVal = cleanAndFormatDateString(rawTar) || rawTar || '-';
          const egtKurumVal = (rawKurum || 'OGM HAVACILIK').toUpperCase();

          if (!person.trainings.some(t => t.egitimAdi === egtUpper && t.egitimTarihi === egtTarVal)) {
            person.trainings.push({
              id: `tr-${Date.now()}-${r}`,
              egitimAdi: egtUpper,
              egitimTarihi: egtTarVal,
              egitimVeren: egtKurumVal
            });
          }
        }
      }

      if (parsedPeople.length > 0) {
        const sortedPeople = parsedPeople.map(p => ({
          ...p,
          trainings: sortTrainingsNewestFirst(p.trainings)
        }));
        setPeople(sortedPeople);
        if (syncToDrive) {
          uploadExcelToGoogleDrive(buildExcelBase64FromPeople(sortedPeople));
        }
        return true;
      }
      return false;
    } catch (err) {
      console.error("Excel parse error:", err);
      return false;
    }
  };

  // Upload an Excel file via "VERİ EKLE (EXCEL YÜKLE)"
  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    if (showNotification) {
      showNotification(`📤 "${file.name}" taranıyor ve Bulut Arşivine kaydediliyor...`);
    }

    try {
      const arrayBuffer = await file.arrayBuffer();
      const bytes = new Uint8Array(arrayBuffer);
      let binary = '';
      for (let i = 0; i < bytes.byteLength; i++) {
        binary += String.fromCharCode(bytes[i]);
      }
      const base64 = btoa(binary);

      parseExcelAndSetPeople(base64, true);

      if (showNotification) {
        showNotification(`✅ Excel yükleme başarılı, kalıcı veritabanı ve Bulut Arşivi güncellendi.`);
      }
    } catch (err: any) {
      alert("Excel yükleme hatası: " + err.message);
    } finally {
      if (fileInputRef.current) fileInputRef.current.value = '';
    }
  };

  // Upload generated or updated Excel to Google Drive
  const uploadExcelToGoogleDrive = async (base64Data?: string, actionMsg: string = "Excel Senkronize Ediliyor...") => {
    setIsLoadingDrive(true);
    setSyncStatusMessage(actionMsg);
    try {
      let b64 = base64Data;
      if (!b64) {
        b64 = buildExcelBase64FromPeople(people);
      }

      const directGasUrls = [
        "https://script.google.com/macros/s/AKfycbwP1uOo2NrST5a4I8vm1nGBLtI26yY2lWrmu9_e9iymwBkUhJBA9JOPCp7SNKqJbOubOw/exec",
        "https://script.google.com/macros/s/AKfycbw4kruTTc058Y9rLTyO3dKi6KloYsmdDTwV1GSiAk8ZXefyo3Z7_VDSTuurzsS9BHAQyQ/exec"
      ];

      for (const gasUrl of directGasUrls) {
        try {
          await fetch(gasUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
            body: JSON.stringify({
              action: 'uploadTechizatExcel',
              fileName: DRIVE_FILE_NAME,
              folderId: '1HQR_NYKhHQGA7_2W3nArI9pCh-LJasTP',
              base64Data: b64,
              base64: b64,
              targetKey: 'personel_egitim_listesi',
              overwrite: true,
              replaceExisting: true,
              deleteDuplicates: true,
              updateIfExists: true,
              cleanDuplicates: true
            })
          });
        } catch (gasErr) {
          console.warn("GAS direct upload warning:", gasErr);
        }
      }
    } catch (e) {
      console.warn("Drive sync failed:", e);
    } finally {
      setIsLoadingDrive(false);
      setSyncStatusMessage('');
    }
  };

  // Build Excel base64 from people with new columns
  const buildExcelBase64FromPeople = (peopleList: PersonelEgitimPerson[]): string => {
    const aoa: any[][] = [
      [
        'SIRA NO',
        'ADI SOYADI',
        'MESLEKİ EĞİTİMLER',
        'EĞİTİM TARİHİ',
        'İSG EĞİTİM TARİHİ',
        'CBİKO (İŞ YERİ GİRİŞ EĞİTİM) TARİHİ',
        'EĞİTİM VEREN KURUM/PERSONEL',
        'LİNK'
      ]
    ];
    const merges: XLSX.Range[] = [];
    let currentRow = 1;

    peopleList.forEach(p => {
      const docText = p.hasPdf ? `BELGE GÖR (https://bulut.ogm.gov.tr/BELGEGOR)` : `BELGE YÜKLE (https://bulut.ogm.gov.tr/BELGEYUKLE)`;
      const isgVal = p.isgEgitimTarihi || '-';
      const cbikoVal = p.cbikoTarihi || '-';

      if (p.trainings.length === 0) {
        aoa.push([p.siraNo, p.adiSoyadi, '-', '-', isgVal, cbikoVal, '-', docText]);
        currentRow++;
      } else {
        const startRow = currentRow;
        const count = p.trainings.length;
        p.trainings.forEach((t, tIdx) => {
          aoa.push([
            tIdx === 0 ? p.siraNo : '',
            tIdx === 0 ? p.adiSoyadi : '',
            t.egitimAdi,
            t.egitimTarihi,
            tIdx === 0 ? isgVal : '',
            tIdx === 0 ? cbikoVal : '',
            t.egitimVeren,
            tIdx === 0 ? docText : ''
          ]);
          currentRow++;
        });
        if (count > 1) {
          merges.push({ s: { r: startRow, c: 0 }, e: { r: startRow + count - 1, c: 0 } });
          merges.push({ s: { r: startRow, c: 1 }, e: { r: startRow + count - 1, c: 1 } });
          merges.push({ s: { r: startRow, c: 4 }, e: { r: startRow + count - 1, c: 4 } });
          merges.push({ s: { r: startRow, c: 5 }, e: { r: startRow + count - 1, c: 5 } });
          merges.push({ s: { r: startRow, c: 7 }, e: { r: startRow + count - 1, c: 7 } });
        }
      }
    });

    const ws = XLSX.utils.aoa_to_sheet(aoa);
    if (merges.length > 0) {
      ws['!merges'] = merges;
    }
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Personel_Egitim_Listesi');
    return XLSX.write(wb, { type: 'base64', bookType: 'xlsx' });
  };

  // Add a new Person
  const handleSaveNewPerson = (e: React.FormEvent) => {
    e.preventDefault();
    if (!newPersonName.trim()) {
      alert("Lütfen personelin adını ve soyadını giriniz.");
      return;
    }

    const nextSiraNo = people.length > 0 ? Math.max(...people.map(p => p.siraNo)) + 1 : 1;
    const initialTrainings: TrainingItem[] = [];

    if (newInitialTrainingName.trim()) {
      initialTrainings.push({
        id: `tr-${Date.now()}-1`,
        egitimAdi: newInitialTrainingName.trim().toUpperCase(),
        egitimTarihi: newInitialTrainingDate.trim() || new Date().toLocaleDateString('tr-TR'),
        egitimVeren: (newInitialTrainingProvider.trim() || 'OGM HAVACILIK').toUpperCase()
      });
    }

    const newPerson: PersonelEgitimPerson = {
      id: `person-${Date.now()}`,
      siraNo: nextSiraNo,
      adiSoyadi: newPersonName.trim().toUpperCase(),
      isgEgitimTarihi: cleanAndFormatDateString(newPersonIsgDate) || newPersonIsgDate || '-',
      cbikoTarihi: cleanAndFormatDateString(newPersonCbikoDate) || newPersonCbikoDate || '-',
      trainings: initialTrainings,
      hasPdf: false,
      linkTitle: `${newPersonName.trim().toUpperCase()} EĞİTİM SERTİFİKALARI`
    };

    const updated = [...people, newPerson];
    setPeople(updated);
    setIsAddPersonModalOpen(false);
    setNewPersonName('');
    setNewPersonIsgDate('');
    setNewPersonCbikoDate('');
    setNewInitialTrainingName('');
    setNewInitialTrainingProvider('');

    uploadExcelToGoogleDrive(buildExcelBase64FromPeople(updated), "Yeni Personel Ekleniyor...");

    if (showNotification) {
      showNotification(`✅ "${newPerson.adiSoyadi}" listeye eklendi.`);
    }
  };

  // Inline "Eğitim Ekle" for a person
  const handleAddInlineTraining = (personId: string) => {
    if (!inlineTrainingName.trim()) {
      alert("Lütfen eğitim adını giriniz.");
      return;
    }

    const newTraining: TrainingItem = {
      id: `tr-${Date.now()}`,
      egitimAdi: inlineTrainingName.trim().toUpperCase(),
      egitimTarihi: inlineTrainingDate.trim() || new Date().toLocaleDateString('tr-TR'),
      egitimVeren: (inlineTrainingProvider.trim() || 'OGM HAVACILIK').toUpperCase()
    };

    const updated = people.map(p => {
      if (p.id === personId) {
        return {
          ...p,
          trainings: sortTrainingsNewestFirst([...p.trainings, newTraining])
        };
      }
      return p;
    });

    setPeople(updated);
    setInlineAddingPersonId(null);
    setInlineTrainingName('');
    setInlineTrainingProvider('');

    uploadExcelToGoogleDrive(buildExcelBase64FromPeople(updated), "Yeni Eğitim Ekleniyor...");

    if (showNotification) {
      showNotification(`✅ Eğitim başarıyla eklendi.`);
    }
  };

  // Delete training
  const handleDeleteTraining = (personId: string, trainingId: string) => {
    if (!confirm("Bu eğitimi listeden silmek istediğinize emin misiniz?")) return;
    const updated = people.map(p => {
      if (p.id === personId) {
        return {
          ...p,
          trainings: p.trainings.filter(t => t.id !== trainingId)
        };
      }
      return p;
    });
    setPeople(updated);
    uploadExcelToGoogleDrive(buildExcelBase64FromPeople(updated), "Eğitim Siliniyor...");
  };

  // Delete person
  const handleDeletePerson = (personId: string) => {
    if (!confirm("Bu personeli ve tüm eğitimlerini silmek istediğinize emin misiniz?")) return;
    const updated = people.filter(p => p.id !== personId);
    setPeople(updated);
    uploadExcelToGoogleDrive(buildExcelBase64FromPeople(updated), "Personel Kaydı Siliniyor...");
  };

  // Edit Training
  const handleOpenEditTrainingModal = (personId: string, training: TrainingItem) => {
    setEditingTrainingInfo({ personId, training });
    setEditTrainingName(training.egitimAdi);
    setEditTrainingDate(training.egitimTarihi);
    setEditTrainingProvider(training.egitimVeren);
  };

  const handleSaveEditedTraining = (e: React.FormEvent) => {
    e.preventDefault();
    if (!editingTrainingInfo) return;
    if (!editTrainingName.trim()) {
      alert("Lütfen eğitim adını giriniz.");
      return;
    }

    const updatedPeople = people.map(p => {
      if (p.id === editingTrainingInfo.personId) {
        return {
          ...p,
          trainings: sortTrainingsNewestFirst(
            p.trainings.map(t => {
              if (t.id === editingTrainingInfo.training.id) {
                return {
                  ...t,
                  egitimAdi: editTrainingName.trim().toUpperCase(),
                  egitimTarihi: editTrainingDate.trim() || new Date().toLocaleDateString('tr-TR'),
                  egitimVeren: (editTrainingProvider.trim() || 'OGM HAVACILIK').toUpperCase()
                };
              }
              return t;
            })
          )
        };
      }
      return p;
    });

    setPeople(updatedPeople);
    setEditingTrainingInfo(null);
    uploadExcelToGoogleDrive(buildExcelBase64FromPeople(updatedPeople), "Eğitim Bilgisi Güncelleniyor...");

    if (showNotification) {
      showNotification("✅ Eğitim bilgisi güncellendi ve Bulut Arşivine kaydedildi.");
    }
  };

  // Edit Person
  const handleOpenEditPersonModal = (person: PersonelEgitimPerson) => {
    setEditingPersonInfo({
      personId: person.id,
      siraNo: person.siraNo,
      adiSoyadi: person.adiSoyadi,
      isgEgitimTarihi: person.isgEgitimTarihi,
      cbikoTarihi: person.cbikoTarihi
    });
    setEditPersonSiraNo(person.siraNo);
    setEditPersonName(person.adiSoyadi);
    setEditPersonIsgDate(person.isgEgitimTarihi || '');
    setEditPersonCbikoDate(person.cbikoTarihi || '');
  };

  const handleSaveEditedPerson = (e: React.FormEvent) => {
    e.preventDefault();
    if (!editingPersonInfo) return;
    if (!editPersonName.trim()) {
      alert("Lütfen personel adını giriniz.");
      return;
    }

    const updatedPeople = people.map(p => {
      if (p.id === editingPersonInfo.personId) {
        return {
          ...p,
          siraNo: editPersonSiraNo || p.siraNo,
          adiSoyadi: editPersonName.trim().toUpperCase(),
          isgEgitimTarihi: cleanAndFormatDateString(editPersonIsgDate) || editPersonIsgDate || '-',
          cbikoTarihi: cleanAndFormatDateString(editPersonCbikoDate) || editPersonCbikoDate || '-'
        };
      }
      return p;
    });

    setPeople(updatedPeople);
    setEditingPersonInfo(null);
    uploadExcelToGoogleDrive(buildExcelBase64FromPeople(updatedPeople), "Personel Bilgileri Güncelleniyor...");

    if (showNotification) {
      showNotification("✅ Personel bilgisi güncellendi.");
    }
  };

  // Belge Yükle / Gör
  const handleBelgeEkleClick = (person: PersonelEgitimPerson) => {
    window.open('https://bulut.ogm.gov.tr/BELGEYUKLE', '_blank');
    setPromptPersonId(person.id);
    setIsPdfConfirmedChecked(true);

    if (showNotification) {
      showNotification(`📄 Belge yükleme sayfası yeni sekmede açıldı.`);
    }
  };

  const handleConfirmPdfUpload = (personId: string) => {
    const updated = people.map(p => (p.id === personId ? { ...p, hasPdf: true } : p));
    setPeople(updated);
    setPromptPersonId(null);
    uploadExcelToGoogleDrive(buildExcelBase64FromPeople(updated), "Belge Onayı Excel'e İşleniyor...");

    if (showNotification) {
      showNotification('✅ Belge onaylandı. Artık "Belge Gör" ile açılabilir.');
    }
  };

  const handleBelgeGorClick = () => {
    window.open('https://bulut.ogm.gov.tr/BELGEGOR', '_blank');
  };

  // Expiration Warning Calculation
  const expiringRecords = useMemo(() => {
    const list: { person: PersonelEgitimPerson; type: 'ISG'; daysLeft: number; expiryDateStr: string }[] = [];
    people.forEach(p => {
      // ISG 3 Yil
      const isgStatus = getExpiryStatus(p.isgEgitimTarihi, 3);
      if (isgStatus.hasDate && (isgStatus.isWarning || isgStatus.isExpired)) {
        list.push({
          person: p,
          type: 'ISG',
          daysLeft: isgStatus.daysLeft,
          expiryDateStr: isgStatus.expiryDateStr
        });
      }
    });
    return list;
  }, [people]);

  // Test Mail Password Handler (1839)
  const handleVerifyTestMailPin = (e: React.FormEvent) => {
    e.preventDefault();
    if (testMailPinInput.trim() === REQUIRED_PIN) {
      setTestMailAuthorized(true);
      setTestMailPinError(false);
      setTestMailPinInput('');
    } else {
      setTestMailPinError(true);
    }
  };

  const handleSimulateSendMail = () => {
    setIsSimulatingMail(true);
    setTimeout(() => {
      setIsSimulatingMail(false);
      setIsTestMailModalOpen(false);
      setTestMailAuthorized(false);
      if (showNotification) {
        showNotification(`📧 Test maili simüle edildi! Gelecek bildirim tarihi: ${getNextMonthFirstDayStr()} (Alıcı: ormanhavacilik.bakimsube@gmail.com)`);
      } else {
        alert(`✅ Test maili başarıyla simüle edildi ve zamanlandı!\nOtomatik Gönderim Tarihi: ${getNextMonthFirstDayStr()}\nAlıcı: ormanhavacilik.bakimsube@gmail.com\nToplam Uyarı Sayısı: ${expiringRecords.length}`);
      }
    }, 1200);
  };

  // Sort Handler
  const toggleSort = (key: SortKey) => {
    if (sortKey === key) {
      setSortDir(prev => (prev === 'asc' ? 'desc' : 'asc'));
    } else {
      setSortKey(key);
      setSortDir('asc');
    }
  };

  // Filtered & Sorted People
  const filteredAndSortedPeople = useMemo(() => {
    let result = people;
    if (searchQuery.trim()) {
      const q = searchQuery.toLowerCase().trim();
      result = result.filter(p =>
        String(p.siraNo).includes(q) ||
        p.adiSoyadi.toLowerCase().includes(q) ||
        (p.isgEgitimTarihi && p.isgEgitimTarihi.toLowerCase().includes(q)) ||
        (p.cbikoTarihi && p.cbikoTarihi.toLowerCase().includes(q)) ||
        p.trainings.some(t =>
          t.egitimAdi.toLowerCase().includes(q) ||
          t.egitimTarihi.toLowerCase().includes(q) ||
          t.egitimVeren.toLowerCase().includes(q)
        )
      );
    }

    return result
      .map(p => ({
        ...p,
        trainings: sortTrainingsNewestFirst(p.trainings)
      }))
      .sort((a, b) => {
        let aVal: any = '';
        let bVal: any = '';

        if (sortKey === 'siraNo') {
          aVal = a.siraNo;
          bVal = b.siraNo;
        } else if (sortKey === 'adiSoyadi') {
          aVal = a.adiSoyadi;
          bVal = b.adiSoyadi;
        } else if (sortKey === 'egitimTarihi') {
          const aPrimary = a.trainings[0]?.egitimTarihi || '';
          const bPrimary = b.trainings[0]?.egitimTarihi || '';
          aVal = getTrainingSortTimestamp(aPrimary);
          bVal = getTrainingSortTimestamp(bPrimary);
        } else if (sortKey === 'isgEgitimTarihi') {
          aVal = parseTrDate(a.isgEgitimTarihi)?.getTime() || 0;
          bVal = parseTrDate(b.isgEgitimTarihi)?.getTime() || 0;
        } else if (sortKey === 'cbikoTarihi') {
          aVal = parseTrDate(a.cbikoTarihi)?.getTime() || 0;
          bVal = parseTrDate(b.cbikoTarihi)?.getTime() || 0;
        }

        if (aVal < bVal) return sortDir === 'asc' ? -1 : 1;
        if (aVal > bVal) return sortDir === 'asc' ? 1 : -1;
        return 0;
      });
  }, [people, searchQuery, sortKey, sortDir]);

  // Excel Download
  const handleDownloadExcel = () => {
    try {
      if (filteredAndSortedPeople.length === 0) {
        alert("İndirilecek personel eğitimi verisi bulunamadı.");
        return;
      }

      let tableRowsHtml = '';

      filteredAndSortedPeople.forEach(p => {
        const rowCount = Math.max(1, p.trainings.length);
        const linkHref = p.hasPdf ? 'https://bulut.ogm.gov.tr/BELGEGOR' : 'https://bulut.ogm.gov.tr/BELGEYUKLE';
        const linkText = p.hasPdf ? `${p.adiSoyadi} EĞİTİM SERTİFİKALARI (BELGE GÖR)` : 'BELGE YÜKLE';
        const linkColor = p.hasPdf ? '#1d4ed8' : '#b45309';

        const isgStr = p.isgEgitimTarihi || '-';
        const cbikoStr = p.cbikoTarihi || '-';

        if (p.trainings.length === 0) {
          tableRowsHtml += `
            <tr>
              <td style="text-align: center; border: 1px solid #cbd5e1; font-weight: bold; background-color: #f1f5f9;">${p.siraNo}</td>
              <td style="border: 1px solid #cbd5e1; font-weight: bold; background-color: #dbeafe; padding: 6px;">${p.adiSoyadi}</td>
              <td style="border: 1px solid #cbd5e1; padding: 6px;">-</td>
              <td style="border: 1px solid #cbd5e1; text-align: center; padding: 6px;">-</td>
              <td style="border: 1px solid #cbd5e1; text-align: center; padding: 6px;">${isgStr}</td>
              <td style="border: 1px solid #cbd5e1; text-align: center; padding: 6px;">${cbikoStr}</td>
              <td style="border: 1px solid #cbd5e1; padding: 6px;">-</td>
              <td style="border: 1px solid #cbd5e1; text-align: center; padding: 6px;">
                <a href="${linkHref}" style="color: ${linkColor}; text-decoration: underline; font-weight: bold;">${linkText}</a>
              </td>
            </tr>
          `;
        } else {
          p.trainings.forEach((t, idx) => {
            if (idx === 0) {
              tableRowsHtml += `
                <tr>
                  <td rowspan="${rowCount}" style="text-align: center; vertical-align: middle; border: 1px solid #cbd5e1; font-weight: bold; background-color: #f1f5f9;">${p.siraNo}</td>
                  <td rowspan="${rowCount}" style="vertical-align: middle; border: 1px solid #cbd5e1; font-weight: bold; background-color: #dbeafe; padding: 8px;">${p.adiSoyadi}</td>
                  <td style="border: 1px solid #cbd5e1; padding: 6px; font-weight: 500;">${t.egitimAdi}</td>
                  <td style="border: 1px solid #cbd5e1; text-align: center; padding: 6px;">${t.egitimTarihi}</td>
                  <td rowspan="${rowCount}" style="vertical-align: middle; text-align: center; border: 1px solid #cbd5e1; padding: 6px;">${isgStr}</td>
                  <td rowspan="${rowCount}" style="vertical-align: middle; text-align: center; border: 1px solid #cbd5e1; padding: 6px;">${cbikoStr}</td>
                  <td style="border: 1px solid #cbd5e1; padding: 6px;">${t.egitimVeren}</td>
                  <td rowspan="${rowCount}" style="vertical-align: middle; text-align: center; border: 1px solid #cbd5e1; padding: 8px;">
                    <a href="${linkHref}" style="color: ${linkColor}; text-decoration: underline; font-weight: bold;">${linkText}</a>
                  </td>
                </tr>
              `;
            } else {
              tableRowsHtml += `
                <tr>
                  <td style="border: 1px solid #cbd5e1; padding: 6px; font-weight: 500;">${t.egitimAdi}</td>
                  <td style="border: 1px solid #cbd5e1; text-align: center; padding: 6px;">${t.egitimTarihi}</td>
                  <td style="border: 1px solid #cbd5e1; padding: 6px;">${t.egitimVeren}</td>
                </tr>
              `;
            }
          });
        }
      });

      const html = `
        <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
        <head>
          <meta charset="utf-8">
          <style>
            table { border-collapse: collapse; width: 100%; font-family: Calibri, Arial, sans-serif; font-size: 11pt; }
            th { background-color: #0b3d1d; color: #ffffff; font-weight: bold; padding: 10px; border: 1px solid #000000; text-align: left; }
            td { padding: 6px; border: 1px solid #cbd5e1; vertical-align: middle; }
          </style>
        </head>
        <body>
          <h2 style="color: #0b3d1d; font-family: Calibri, Arial, sans-serif; margin-bottom: 8px;">
            ORMAN GENEL MÜDÜRLÜĞÜ - PERSONEL EĞİTİM LİSTESİ
          </h2>
          <table border="1">
            <thead>
              <tr style="background-color: #0b3d1d; color: #ffffff;">
                <th style="text-align: center; width: 60px;">SIRA NO</th>
                <th style="width: 220px;">ADI SOYADI</th>
                <th style="width: 380px;">MESLEKİ EĞİTİMLER</th>
                <th style="text-align: center; width: 140px;">EĞİTİM TARİHİ</th>
                <th style="text-align: center; width: 150px;">İSG EĞİTİM TARİHİ</th>
                <th style="text-align: center; width: 220px;">CBİKO (İŞ YERİ GİRİŞ EĞİTİM) TARİHİ</th>
                <th style="width: 260px;">EĞİTİM VEREN KURUM/PERSONEL</th>
                <th style="text-align: center; width: 260px;">LİNK</th>
              </tr>
            </thead>
            <tbody>
              ${tableRowsHtml}
            </tbody>
          </table>
        </body>
        </html>
      `;

      const blob = new Blob([html], { type: 'application/vnd.ms-excel;charset=utf-8;' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `personel_egitim_listesi_${new Date().toISOString().slice(0, 10)}.xls`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);

      if (showNotification) {
        showNotification('✅ Personel Eğitim Listesi Excel formatında indirildi.');
      }
    } catch (e: any) {
      alert("Excel indirme hatası: " + e.message);
    }
  };

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-[600] bg-slate-950 flex flex-col overflow-hidden text-white animate-fade-in select-none">
      {/* 1. TOP HEADER */}
      <div className="shrink-0 bg-slate-900 border-b border-slate-800 px-6 py-4 flex flex-col lg:flex-row items-center justify-between gap-4 shadow-md select-none">
        
        {/* Brand & Badge */}
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 rounded-2xl bg-emerald-500/10 border border-emerald-500/30 flex items-center justify-center text-emerald-400 shadow-inner">
            <GraduationCap className="w-5 h-5" />
          </div>
          <div>
            <div className="flex items-center gap-2">
              <span className="w-2 h-2 rounded-full bg-emerald-400 animate-pulse" />
              <h2 className="text-sm font-black text-white tracking-widest uppercase font-mono">
                PERSONEL EĞİTİM LİSTESİ
              </h2>
            </div>
            {expiringRecords.length > 0 && (
              <span className="text-[10px] bg-amber-500/10 text-amber-400 border border-amber-500/30 font-bold px-2 py-0.5 rounded-full flex items-center gap-1 mt-0.5">
                <AlertTriangle className="w-3 h-3 text-amber-400" />
                {expiringRecords.length} Personelde Tazeleme Yaklaştı / Doldu
              </span>
            )}
          </div>
        </div>

        {/* Center: Search Box */}
        <div className="flex items-center gap-3 w-full max-w-xl">
          <div className="flex items-center gap-3 bg-slate-950 px-4 py-2.5 rounded-2xl border border-slate-800 shadow-inner w-full">
            <Search className="w-4 h-4 text-emerald-400 shrink-0" />
            <input
              type="text"
              placeholder="PERSONEL ADI, MESLEKİ EĞİTİM VEYA TARİH ARA..."
              value={searchQuery}
              onChange={(e) => {
                setSearchQuery(e.target.value);
                setActiveMatchIdx(0);
              }}
              className="bg-transparent text-slate-100 text-xs font-extrabold outline-none border-none placeholder-slate-600 w-full uppercase"
            />
            {searchQuery && (
              <button
                onClick={() => {
                  setSearchQuery('');
                  setActiveMatchIdx(0);
                }}
                className="text-slate-500 hover:text-slate-300 transition-colors cursor-pointer"
                title="Aramayı Temizle"
              >
                <X className="w-4 h-4" />
              </button>
            )}
          </div>
        </div>

        {/* Right Actions & Filters */}
        <div className="flex items-center gap-2.5 shrink-0 flex-wrap">
          
          <input
            type="file"
            ref={fileInputRef}
            onChange={handleFileUpload}
            accept=".xlsx,.xls"
            className="hidden"
          />

          {/* TEST MAİLİ GÖNDER BUTTON (PIN: 1839) */}
          <button
            onClick={() => {
              setIsTestMailModalOpen(true);
              setTestMailAuthorized(false);
              setTestMailPinError(false);
            }}
            className="px-3.5 py-2 bg-amber-600 hover:bg-amber-500 text-white font-black font-mono text-xs rounded-xl flex items-center gap-2 transition-all cursor-pointer shadow-md hover:scale-[1.02] active:scale-95 border border-amber-400/40"
            title="Süresi yaklaşan İSG ve Muayeneler için test bildirim maili zamanla"
          >
            <Mail className="w-4 h-4 text-amber-200 animate-bounce" />
            <span>TEST MAİLİ GÖNDER</span>
          </button>

          {/* VERİ EKLE (EXCEL YÜKLE) */}
          <button
            onClick={() => fileInputRef.current?.click()}
            className="px-3.5 py-2 bg-blue-600 hover:bg-blue-500 text-white font-black font-mono text-xs rounded-xl flex items-center gap-2 transition-all cursor-pointer shadow-md hover:scale-[1.02] active:scale-95"
            title="Bulut Arşivine ve sisteme Excel dosyası yükle"
          >
            <Upload className="w-4 h-4 text-blue-200" />
            <span>VERİ EKLE (EXCEL YÜKLE)</span>
          </button>

          {/* YENİ PERSONEL EKLE */}
          <button
            onClick={() => setIsAddPersonModalOpen(true)}
            className="px-3.5 py-2 bg-emerald-600 hover:bg-emerald-500 text-white font-black font-mono text-xs rounded-xl flex items-center gap-2 transition-all cursor-pointer shadow-md hover:scale-[1.02] active:scale-95"
            title="Sıfırdan Yeni Personel ve Eğitimi Ekle"
          >
            <Plus className="w-4 h-4" />
            <span>YENİ PERSONEL EKLE</span>
          </button>

          {/* DRIVE'DAN YENİLE */}
          <button
            onClick={() => fetchLiveExcelFromDrive(true)}
            disabled={isLoadingDrive}
            className="p-2 bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white rounded-xl transition-all border border-slate-700 cursor-pointer"
            title="Bulut Arşivinden verileri canlı tazele"
          >
            <RefreshCw className={`w-4 h-4 text-emerald-400 ${isLoadingDrive ? 'animate-spin' : ''}`} />
          </button>

          {/* EXCEL İNDİR */}
          <button
            onClick={handleDownloadExcel}
            className="px-3.5 py-2 bg-slate-800 hover:bg-slate-700 text-slate-200 hover:text-white font-extrabold font-mono text-xs rounded-xl flex items-center gap-2 transition-all cursor-pointer border border-slate-700 shadow-sm active:scale-95"
            title="Mevcut Tabloyu Excel Olarak İndir"
          >
            <Download className="w-4 h-4 text-emerald-400" />
            <span>EXCEL İNDİR</span>
          </button>

          {/* Kapat Butonu */}
          <button
            onClick={onClose}
            className="p-2 bg-slate-800 hover:bg-red-950 hover:text-red-400 text-slate-400 rounded-xl transition-all border border-slate-700 cursor-pointer ml-1"
            title="Kapat"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

      </div>

      {/* 2. TABLE BODY CONTAINER */}
      <div className="flex-1 overflow-auto p-4 md:p-6 bg-slate-950 flex justify-center scrollbar-thin">
        <div className="w-full bg-white text-slate-800 rounded-2xl shadow-2xl p-6 flex flex-col border border-slate-200 overflow-x-auto">
          
          {/* Info Banner */}
          <div className="mb-4 flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3 bg-emerald-50/80 border border-emerald-200/80 px-4 py-2.5 rounded-xl text-xs text-emerald-950">
            <div className="flex items-center gap-2">
              <BookOpen className="w-4 h-4 text-emerald-700 shrink-0" />
              <span className="font-semibold">
                Her personelin mesleki eğitimleri, İSG tarihi (3 yıl), CBİKO tarihi ve periyodik muayene (5 yıl) tarihleri listelenmektedir. Tarih başlıklarına tıklayarak Eskiden Yeniye veya Yeniden Eskiye sıralayabilirsiniz.
              </span>
            </div>
            <div className="flex items-center gap-3 font-mono text-[11px] text-emerald-800 shrink-0">
              <div className="flex items-center gap-1.5">
                <span className="font-bold">Toplam Personel:</span>
                <span className="bg-emerald-200/80 px-2 py-0.5 rounded-full font-black text-emerald-950">{people.length}</span>
              </div>
              <div className="flex items-center gap-1 bg-white/80 p-0.5 rounded-lg border border-emerald-200 shadow-xs">
                <button
                  type="button"
                  onClick={expandAll}
                  className="px-2 py-1 text-[10px] font-bold text-emerald-900 hover:bg-emerald-100 rounded cursor-pointer transition-colors"
                >
                  Tümünü Aç
                </button>
                <span className="text-emerald-300">|</span>
                <button
                  type="button"
                  onClick={collapseAll}
                  className="px-2 py-1 text-[10px] font-bold text-slate-600 hover:bg-slate-100 rounded cursor-pointer transition-colors"
                >
                  Tümünü Kapat
                </button>
              </div>
            </div>
          </div>

          {/* TABLE CONTAINER */}
          <div className="relative flex-1 overflow-x-auto w-full rounded-xl border border-slate-300 min-h-[420px]">
            {isLoadingDrive && (
              <div className="absolute inset-0 z-40 bg-white/80 backdrop-blur-md flex flex-col items-center justify-center gap-3">
                <div className="w-12 h-12 rounded-full border-4 border-emerald-200 border-t-emerald-600 animate-spin" />
                <div className="text-sm font-black text-slate-900 font-mono">
                  {syncStatusMessage || "Veriler Yükleniyor..."}
                </div>
              </div>
            )}

            <table className={`table-auto w-full border-collapse text-left text-xs text-slate-800 transition duration-300 ${isLoadingDrive ? 'filter blur-xs opacity-50' : ''}`}>
              <thead>
                <tr className="bg-slate-900 text-white font-extrabold uppercase tracking-wider text-[11px] select-none">
                  {/* SIRA NO */}
                  <th
                    onClick={() => toggleSort('siraNo')}
                    className="p-3 border-b-4 border-r border-slate-600 text-center whitespace-nowrap w-16 cursor-pointer hover:bg-slate-800 transition-colors"
                  >
                    <div className="flex items-center justify-center gap-1">
                      <span>SIRA NO</span>
                      {sortKey === 'siraNo' ? (sortDir === 'asc' ? <ArrowUp className="w-3 h-3 text-emerald-400" /> : <ArrowDown className="w-3 h-3 text-emerald-400" />) : <ArrowUpDown className="w-3 h-3 opacity-40" />}
                    </div>
                  </th>

                  {/* ADI SOYADI */}
                  <th
                    onClick={() => toggleSort('adiSoyadi')}
                    className="p-3 border-b-4 border-r border-slate-600 text-left whitespace-nowrap min-w-[200px] cursor-pointer hover:bg-slate-800 transition-colors"
                  >
                    <div className="flex items-center gap-1">
                      <span>ADI SOYADI</span>
                      {sortKey === 'adiSoyadi' ? (sortDir === 'asc' ? <ArrowUp className="w-3 h-3 text-emerald-400" /> : <ArrowDown className="w-3 h-3 text-emerald-400" />) : <ArrowUpDown className="w-3 h-3 opacity-40" />}
                    </div>
                  </th>

                  {/* MESLEKİ EĞİTİMLER */}
                  <th className="p-3 border-b-4 border-r border-slate-600 text-left whitespace-nowrap min-w-[300px]">
                    MESLEKİ EĞİTİMLER
                  </th>

                  {/* EĞİTİM TARİHİ */}
                  <th
                    onClick={() => toggleSort('egitimTarihi')}
                    className="p-3 border-b-4 border-r border-slate-600 text-center whitespace-nowrap min-w-[130px] cursor-pointer hover:bg-slate-800 transition-colors"
                  >
                    <div className="flex items-center justify-center gap-1">
                      <span>EĞİTİM TARİHİ</span>
                      {sortKey === 'egitimTarihi' ? (sortDir === 'asc' ? <ArrowUp className="w-3 h-3 text-emerald-400" /> : <ArrowDown className="w-3 h-3 text-emerald-400" />) : <ArrowUpDown className="w-3 h-3 opacity-40" />}
                    </div>
                  </th>

                  {/* EĞİTİM VEREN KURUM */}
                  <th className="p-3 border-b-4 border-r border-slate-600 text-left whitespace-nowrap min-w-[220px]">
                    EĞİTİM VEREN KURUM/PERSONEL
                  </th>

                  {/* İSG EĞİTİM TARİHİ */}
                  <th
                    onClick={() => toggleSort('isgEgitimTarihi')}
                    className="p-3 border-b-4 border-r border-slate-600 text-center whitespace-nowrap min-w-[150px] cursor-pointer hover:bg-slate-800 transition-colors"
                    title="İSG Eğitimi 3 Yılda Bir Yenilenmelidir. Sıralamak için tıklayınız."
                  >
                    <div className="flex items-center justify-center gap-1 text-amber-300">
                      <span>İSG EĞİTİM TARİHİ (3 YIL)</span>
                      {sortKey === 'isgEgitimTarihi' ? (sortDir === 'asc' ? <ArrowUp className="w-3 h-3 text-amber-400" /> : <ArrowDown className="w-3 h-3 text-amber-400" />) : <ArrowUpDown className="w-3 h-3 opacity-40" />}
                    </div>
                  </th>

                  {/* CBİKO TARİHİ */}
                  <th
                    onClick={() => toggleSort('cbikoTarihi')}
                    className="p-3 border-b-4 border-r border-slate-600 text-center whitespace-nowrap min-w-[180px] cursor-pointer hover:bg-slate-800 transition-colors"
                  >
                    <div className="flex items-center justify-center gap-1">
                      <span>CBİKO (İŞ YERİ GİRİŞ EĞİTİM) TARİHİ</span>
                      {sortKey === 'cbikoTarihi' ? (sortDir === 'asc' ? <ArrowUp className="w-3 h-3 text-emerald-400" /> : <ArrowDown className="w-3 h-3 text-emerald-400" />) : <ArrowUpDown className="w-3 h-3 opacity-40" />}
                    </div>
                  </th>

                  {/* LİNK */}
                  <th className="p-3 border-b-4 border-slate-600 text-center whitespace-nowrap min-w-[200px]">
                    LİNK
                  </th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-300">
                {!isLoadingDrive && filteredAndSortedPeople.length === 0 ? (
                  <tr>
                    <td colSpan={8} className="text-center py-20 text-slate-400 font-mono text-xs uppercase tracking-wider">
                      <div className="flex flex-col items-center justify-center gap-3">
                        <FileSpreadsheet className="w-12 h-12 text-slate-300 animate-pulse" />
                        <span className="font-bold text-slate-600">Henüz personel eğitimi verisi bulunmamaktadır.</span>
                      </div>
                    </td>
                  </tr>
                ) : (
                  filteredAndSortedPeople.map((person) => {
                    const isPromptActive = promptPersonId === person.id;
                    const isInlineAdding = inlineAddingPersonId === person.id;
                    const trainingsCount = person.trainings.length;
                    const isExpanded = expandedPersonIds.has(person.id);
                    const primaryTraining = trainingsCount > 0 ? person.trainings[0] : null;

                    const isgStatus = getExpiryStatus(person.isgEgitimTarihi, 3);

                    return (
                      <React.Fragment key={person.id}>
                        {/* MAIN PERSON ROW */}
                        <tr className="border-b-2 border-slate-300 hover:bg-slate-50/90 transition-colors">
                          {/* 1. SIRA NO */}
                          <td className="p-3 border-r border-slate-300 font-mono font-bold text-center text-slate-700 bg-slate-100/70 align-middle w-16">
                            <div className="flex flex-col items-center justify-center gap-1.5">
                              <span className="text-slate-900 font-extrabold text-sm">{person.siraNo}</span>
                              <button
                                type="button"
                                onClick={() => {
                                  setInlineAddingPersonId(person.id);
                                  setInlineTrainingName('');
                                  setInlineTrainingDate(new Date().toLocaleDateString('tr-TR'));
                                  setInlineTrainingProvider('');
                                }}
                                className="w-6 h-6 rounded-full bg-emerald-600 hover:bg-emerald-500 text-white flex items-center justify-center cursor-pointer shadow-xs transition-all active:scale-90"
                                title="Bu personele yeni mesleki eğitim ekle (+)"
                              >
                                <Plus className="w-3.5 h-3.5 stroke-[3]" />
                              </button>
                            </div>
                          </td>

                          {/* 2. ADI SOYADI */}
                          <td className="p-3.5 border-r border-slate-300 font-bold text-slate-900 bg-[#dbeafe] align-middle min-w-[200px]">
                            <div className="flex flex-col gap-1">
                              <span className="font-extrabold text-sm tracking-wide text-slate-900">
                                {person.adiSoyadi}
                              </span>
                              <div className="flex items-center justify-between gap-1 mt-1 text-[10px]">
                                <span className="font-mono text-slate-600 bg-blue-100/80 px-2 py-0.5 rounded-md font-semibold">
                                  {trainingsCount} Eğitim
                                </span>
                                <div className="flex items-center gap-1.5 ml-auto">
                                  <button
                                    type="button"
                                    onClick={() => handleOpenEditPersonModal(person)}
                                    className="text-blue-700 hover:text-blue-900 hover:underline flex items-center gap-0.5 cursor-pointer font-bold"
                                    title="Personel Bilgileri & Tarihleri Düzenle"
                                  >
                                    <Edit2 className="w-3 h-3" />
                                    <span>Düzenle</span>
                                  </button>
                                  <span className="text-slate-300">|</span>
                                  <button
                                    type="button"
                                    onClick={() => handleDeletePerson(person.id)}
                                    className="text-red-600 hover:text-red-800 hover:underline flex items-center gap-0.5 cursor-pointer font-bold"
                                    title="Personeli Sil"
                                  >
                                    <Trash2 className="w-3 h-3" />
                                    <span>Sil</span>
                                  </button>
                                </div>
                              </div>
                            </div>
                          </td>

                          {/* 3. MESLEKİ EĞİTİMLER */}
                          <td className="p-3 border-r border-slate-200 font-semibold text-slate-800 align-middle">
                            {primaryTraining ? (
                              <div className="flex flex-col gap-1.5">
                                <div className="flex items-center justify-between gap-2">
                                  <span className="font-bold text-slate-900 text-xs">{primaryTraining.egitimAdi}</span>
                                  <div className="flex items-center gap-1 shrink-0">
                                    <button
                                      type="button"
                                      onClick={() => handleOpenEditTrainingModal(person.id, primaryTraining)}
                                      className="text-slate-400 hover:text-blue-600 transition-colors p-1 cursor-pointer rounded hover:bg-blue-50"
                                      title="Bu Eğitimi Düzenle"
                                    >
                                      <Edit2 className="w-3.5 h-3.5" />
                                    </button>
                                    <button
                                      type="button"
                                      onClick={() => handleDeleteTraining(person.id, primaryTraining.id)}
                                      className="text-slate-400 hover:text-red-600 transition-colors p-1 cursor-pointer rounded hover:bg-red-50"
                                      title="Bu Eğitimi Sil"
                                    >
                                      <Trash2 className="w-3.5 h-3.5" />
                                    </button>
                                  </div>
                                </div>

                                {trainingsCount > 1 && (
                                  <button
                                    type="button"
                                    onClick={() => toggleExpandPerson(person.id)}
                                    className="self-start inline-flex items-center gap-1.5 px-2.5 py-1 rounded-lg bg-emerald-50 hover:bg-emerald-100 text-emerald-800 border border-emerald-300 font-mono text-[11px] font-bold transition-all cursor-pointer shadow-xs active:scale-95"
                                  >
                                    {isExpanded ? (
                                      <>
                                        <ChevronUp className="w-3.5 h-3.5 text-emerald-600" />
                                        <span>Diğer Eğitimleri Gizle</span>
                                      </>
                                    ) : (
                                      <>
                                        <ChevronDown className="w-3.5 h-3.5 text-emerald-600" />
                                        <span>Diğer Eğitimleri Gör ({trainingsCount - 1} Eğitim Daha)</span>
                                      </>
                                    )}
                                  </button>
                                )}
                              </div>
                            ) : (
                              <span className="text-slate-400 italic text-xs">Henüz mesleki eğitim kaydı bulunmuyor</span>
                            )}
                          </td>

                          {/* 4. EĞİTİM TARİHİ */}
                          <td className="p-3 border-r border-slate-200 text-center font-mono font-bold text-slate-700 whitespace-nowrap align-middle min-w-[130px]">
                            {primaryTraining ? (
                              <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-full bg-slate-100 text-slate-800 border border-slate-200 text-[11px]">
                                <Calendar className="w-3 h-3 text-slate-400" />
                                {primaryTraining.egitimTarihi}
                              </span>
                            ) : (
                              <span>-</span>
                            )}
                          </td>

                          {/* 7. EĞİTİM VEREN KURUM/PERSONEL */}
                          <td className="p-3 border-r border-slate-200 font-medium text-slate-700 whitespace-nowrap align-middle min-w-[220px]">
                            {primaryTraining ? (
                              <div className="flex items-center gap-1.5 text-xs">
                                <Building2 className="w-3.5 h-3.5 text-slate-400 shrink-0" />
                                <span>{primaryTraining.egitimVeren}</span>
                              </div>
                            ) : (
                              <span>-</span>
                            )}
                          </td>

                          {/* 5. İSG EĞİTİM TARİHİ (3 YIL KALANLAR ALERT) */}
                          <td className="p-3 border-r border-slate-200 text-center font-mono font-bold whitespace-nowrap align-middle min-w-[150px]">
                            {person.isgEgitimTarihi && person.isgEgitimTarihi !== '-' ? (
                              <div className="flex flex-col items-center gap-1">
                                <span className={`inline-flex items-center gap-1 px-2.5 py-1 rounded-full border text-[11px] font-bold ${
                                  isgStatus.isExpired
                                    ? 'bg-red-100 text-red-800 border-red-300 animate-pulse'
                                    : isgStatus.isWarning
                                    ? 'bg-amber-100 text-amber-900 border-amber-300 animate-pulse'
                                    : 'bg-slate-100 text-slate-800 border-slate-200'
                                }`}>
                                  <Calendar className="w-3 h-3" />
                                  {person.isgEgitimTarihi}
                                </span>
                                {isgStatus.isWarning && (
                                  <span className="text-[9px] font-extrabold text-amber-700 bg-amber-50 px-1.5 py-0.5 rounded border border-amber-200">
                                    ⚠️ {isgStatus.daysLeft} Gün Kaldı (Mail Zamanlandı)
                                  </span>
                                )}
                                {isgStatus.isExpired && (
                                  <span className="text-[9px] font-extrabold text-red-700 bg-red-50 px-1.5 py-0.5 rounded border border-red-200">
                                    🚨 TAZELENMELİ (SÜRE DOLDU)
                                  </span>
                                )}
                              </div>
                            ) : (
                              <span className="text-slate-400 italic text-[11px]">-</span>
                            )}
                          </td>

                          {/* 6. CBİKO TARİHİ */}
                          <td className="p-3 border-r border-slate-200 text-center font-mono font-bold text-slate-700 whitespace-nowrap align-middle min-w-[180px]">
                            {person.cbikoTarihi && person.cbikoTarihi !== '-' ? (
                              <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-full bg-slate-100 text-slate-800 border border-slate-200 text-[11px]">
                                <Calendar className="w-3 h-3 text-slate-400" />
                                {person.cbikoTarihi}
                              </span>
                            ) : (
                              <span className="text-slate-400 italic text-[11px]">-</span>
                            )}
                          </td>

                          {/* 8. LİNK */}
                          <td className="p-3 border-r border-slate-200 text-center align-middle whitespace-nowrap bg-white min-w-[200px]">
                            {isPromptActive ? (
                              <div className="inline-flex flex-col items-center gap-2 bg-amber-50 border-2 border-amber-400 p-2.5 rounded-2xl shadow-lg text-amber-950">
                                <label className="flex items-center gap-1.5 cursor-pointer text-[11px] font-bold">
                                  <input
                                    type="checkbox"
                                    checked={isPdfConfirmedChecked}
                                    onChange={(e) => setIsPdfConfirmedChecked(e.target.checked)}
                                    className="w-4 h-4 text-emerald-600 rounded border-amber-300"
                                  />
                                  <span className="font-black text-amber-900">PDF YÜKLENDİ Mİ?</span>
                                </label>
                                <div className="flex items-center gap-2">
                                  <button
                                    type="button"
                                    disabled={!isPdfConfirmedChecked}
                                    onClick={() => handleConfirmPdfUpload(person.id)}
                                    className={`px-3 py-1 text-white font-black text-[11px] rounded-lg transition-all flex items-center gap-1 shadow cursor-pointer ${
                                      isPdfConfirmedChecked
                                        ? 'bg-emerald-600 hover:bg-emerald-500'
                                        : 'bg-slate-300 cursor-not-allowed'
                                    }`}
                                  >
                                    <CheckCircle2 className="w-3.5 h-3.5" />
                                    <span>EVET</span>
                                  </button>
                                  <button
                                    type="button"
                                    onClick={() => setPromptPersonId(null)}
                                    className="p-1 text-slate-400 hover:text-slate-600"
                                  >
                                    <X className="w-3.5 h-3.5" />
                                  </button>
                                </div>
                              </div>
                            ) : person.hasPdf ? (
                              <div className="inline-flex flex-col items-center gap-1">
                                <button
                                  type="button"
                                  onClick={handleBelgeGorClick}
                                  className="px-3 py-1.5 bg-emerald-600 hover:bg-emerald-500 text-white font-black text-[11px] rounded-xl flex items-center gap-1.5 transition-all cursor-pointer shadow-sm active:scale-95 group"
                                >
                                  <FileCheck className="w-3.5 h-3.5 text-emerald-200" />
                                  <span>BELGE GÖR</span>
                                  <ExternalLink className="w-3 h-3 text-emerald-300" />
                                </button>
                                <button
                                  type="button"
                                  onClick={() => handleBelgeEkleClick(person)}
                                  className="text-[10px] text-slate-400 hover:text-slate-700 flex items-center gap-1 cursor-pointer"
                                >
                                  <RefreshCw className="w-2.5 h-2.5" />
                                  <span>Değiştir</span>
                                </button>
                              </div>
                            ) : (
                              <button
                                type="button"
                                onClick={() => handleBelgeEkleClick(person)}
                                className="px-3 py-1.5 bg-amber-500 hover:bg-amber-600 text-white font-black text-[11px] rounded-xl flex items-center gap-1.5 transition-all shadow-sm cursor-pointer active:scale-95 group"
                              >
                                <FilePlus className="w-3.5 h-3.5 text-amber-200" />
                                <span>BELGE EKLE</span>
                                <ExternalLink className="w-3 h-3 text-amber-200" />
                              </button>
                            )}
                          </td>
                        </tr>

                        {/* ACCORDION ROW */}
                        {isExpanded && trainingsCount > 1 && (
                          <tr className="bg-slate-50/95 border-b-4 border-slate-800">
                            <td colSpan={8} className="p-4 pl-12 pr-6">
                              <div className="bg-white rounded-2xl border-2 border-emerald-500/40 shadow-lg p-4">
                                <div className="flex items-center justify-between pb-3 mb-3 border-b border-slate-200">
                                  <div className="flex items-center gap-2.5">
                                    <div className="w-8 h-8 rounded-xl bg-emerald-50 border border-emerald-300 flex items-center justify-center text-emerald-700">
                                      <GraduationCap className="w-4 h-4" />
                                    </div>
                                    <div>
                                      <h4 className="text-xs font-black uppercase text-slate-900 tracking-wider font-mono">
                                        {person.adiSoyadi} • TÜM EĞİTİM GEÇMİŞİ
                                      </h4>
                                    </div>
                                  </div>
                                </div>

                                <div className="overflow-x-auto rounded-xl border border-slate-200">
                                  <table className="w-full text-xs text-left">
                                    <thead>
                                      <tr className="bg-slate-100 text-slate-700 font-extrabold uppercase text-[10px] tracking-wider border-b border-slate-200">
                                        <th className="p-2.5 text-center w-12 font-mono">#</th>
                                        <th className="p-2.5 text-left">MESLEKİ EĞİTİM ADI</th>
                                        <th className="p-2.5 text-center w-40">EĞİTİM TARİHİ</th>
                                        <th className="p-2.5 text-left w-72">EĞİTİM VEREN KURUM / PERSONEL</th>
                                        <th className="p-2.5 text-center w-16">İŞLEM</th>
                                      </tr>
                                    </thead>
                                    <tbody className="divide-y divide-slate-100">
                                      {person.trainings.map((t, idx) => (
                                        <tr key={t.id} className="hover:bg-emerald-50/40 transition-colors">
                                          <td className="p-2.5 text-center font-mono font-bold text-slate-500 bg-slate-50/50">
                                            {idx + 1}
                                          </td>
                                          <td className="p-2.5 font-bold text-slate-900">
                                            {t.egitimAdi}
                                          </td>
                                          <td className="p-2.5 text-center font-mono font-bold text-slate-700">
                                            {t.egitimTarihi}
                                          </td>
                                          <td className="p-2.5 text-slate-700 font-medium">
                                            {t.egitimVeren}
                                          </td>
                                          <td className="p-2.5 text-center">
                                            <div className="flex items-center justify-center gap-1">
                                              <button
                                                type="button"
                                                onClick={() => handleOpenEditTrainingModal(person.id, t)}
                                                className="p-1 text-slate-400 hover:text-blue-600 rounded-md hover:bg-blue-50 transition cursor-pointer"
                                              >
                                                <Edit2 className="w-3.5 h-3.5" />
                                              </button>
                                              <button
                                                type="button"
                                                onClick={() => handleDeleteTraining(person.id, t.id)}
                                                className="p-1 text-slate-400 hover:text-red-600 rounded-md hover:bg-red-50 transition cursor-pointer"
                                              >
                                                <Trash2 className="w-3.5 h-3.5" />
                                              </button>
                                            </div>
                                          </td>
                                        </tr>
                                      ))}
                                    </tbody>
                                  </table>
                                </div>
                              </div>
                            </td>
                          </tr>
                        )}

                        {/* INLINE TRAINING INPUT ROW */}
                        {isInlineAdding && (
                          <tr className="border-b-2 border-emerald-400 bg-emerald-50/80">
                            <td className="p-3 text-center font-mono font-bold text-emerald-800 bg-emerald-100/50">
                              +
                            </td>
                            <td className="p-3 font-bold text-emerald-950 font-mono text-xs">
                              {person.adiSoyadi} (Yeni Mesleki Eğitim)
                            </td>
                            <td className="p-3 border-r border-slate-200">
                              <input
                                type="text"
                                autoFocus
                                placeholder="Mesleki Eğitim Adı..."
                                value={inlineTrainingName}
                                onChange={(e) => setInlineTrainingName(e.target.value)}
                                className="w-full px-3 py-2 bg-white border-2 border-emerald-500 rounded-xl text-xs font-bold uppercase text-slate-900 outline-none"
                              />
                            </td>
                            <td className="p-3 border-r border-slate-200 text-center">
                              <input
                                type="text"
                                placeholder="GG.AA.YYYY"
                                value={inlineTrainingDate}
                                onChange={(e) => setInlineTrainingDate(e.target.value)}
                                className="w-32 px-3 py-2 bg-white border-2 border-emerald-500 rounded-xl text-xs font-mono font-bold text-slate-900 text-center outline-none"
                              />
                            </td>
                            <td colSpan={2} className="p-3 border-r border-slate-200" />
                            <td className="p-3 border-r border-slate-200 overflow-visible">
                              <div className="flex items-center gap-2">
                                <div className="relative flex-1">
                                  <input
                                    type="text"
                                    placeholder="Veren Kurum / Personel"
                                    value={inlineTrainingProvider}
                                    onFocus={() => setActiveProviderField('inline')}
                                    onBlur={() => setTimeout(() => setActiveProviderField(prev => (prev === 'inline' ? null : prev)), 150)}
                                    onChange={(e) => {
                                      setInlineTrainingProvider(e.target.value);
                                      setActiveProviderField('inline');
                                    }}
                                    className="w-full px-3 py-2 bg-white border-2 border-emerald-500 rounded-xl text-xs font-bold uppercase text-slate-900 outline-none"
                                  />
                                  {activeProviderField === 'inline' && getFilteredProviders(inlineTrainingProvider).length > 0 && (
                                    <div className="absolute top-full left-0 right-0 mt-1 z-50 bg-white border border-emerald-300 rounded-xl shadow-xl max-h-44 overflow-y-auto divide-y divide-slate-100">
                                      {getFilteredProviders(inlineTrainingProvider).map((prov, pIdx) => (
                                        <button
                                          key={pIdx}
                                          type="button"
                                          onMouseDown={(e) => {
                                            e.preventDefault();
                                            setInlineTrainingProvider(prov);
                                            setActiveProviderField(null);
                                          }}
                                          className="w-full text-left px-3 py-2 text-xs font-bold uppercase text-slate-800 hover:bg-emerald-50 hover:text-emerald-900 flex items-center gap-1.5 cursor-pointer transition"
                                        >
                                          <Building2 className="w-3.5 h-3.5 text-emerald-600 shrink-0" />
                                          <span className="truncate">{prov}</span>
                                        </button>
                                      ))}
                                    </div>
                                  )}
                                </div>
                                <button
                                  type="button"
                                  onClick={() => handleAddInlineTraining(person.id)}
                                  className="px-4 py-2 bg-emerald-600 hover:bg-emerald-500 text-white font-black text-xs rounded-xl flex items-center gap-1 shadow cursor-pointer shrink-0"
                                >
                                  <Check className="w-3.5 h-3.5 stroke-[3]" />
                                  <span>Ekle</span>
                                </button>
                                <button
                                  type="button"
                                  onClick={() => setInlineAddingPersonId(null)}
                                  className="p-2 bg-slate-200 hover:bg-slate-300 text-slate-700 rounded-xl cursor-pointer shrink-0"
                                >
                                  <X className="w-4 h-4" />
                                </button>
                              </div>
                            </td>
                            <td colSpan={2} className="p-3 text-center bg-white" />
                          </tr>
                        )}
                      </React.Fragment>
                    );
                  })
                )}
              </tbody>
            </table>
          </div>

          {/* Table Footer */}
          <div className="mt-4 pt-3 border-t border-slate-200 flex flex-col sm:flex-row items-center justify-between text-[11px] text-slate-500 font-mono">
            <span>
              Resmi Havacılık Personel Eğitim Portalı • Son Senkronizasyon: {new Date().toLocaleDateString('tr-TR')}
            </span>
            <span className="font-bold text-emerald-800">
              OGM Yangınla Mücadele Havacılık Teknik Şube
            </span>
          </div>

        </div>
      </div>

      {/* 3. TEST MAİLİ GÖNDER MODALI (PIN: 1839) */}
      {isTestMailModalOpen && (
        <div className="fixed inset-0 z-[750] bg-black/80 backdrop-blur-md flex items-center justify-center p-4">
          <div className="bg-slate-900 border border-slate-800 text-white rounded-3xl shadow-2xl max-w-lg w-full overflow-hidden animate-scale-up">
            
            {/* Header */}
            <div className="bg-gradient-to-r from-amber-950 via-slate-900 to-amber-950 p-5 border-b border-slate-800 flex items-center justify-between">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 rounded-2xl bg-amber-500/20 border border-amber-500/40 flex items-center justify-center text-amber-400">
                  <Mail className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="font-extrabold text-sm uppercase tracking-wider font-mono text-amber-400">
                    SÜRE DOLUM TEST BİLDİRİM MAİLİ
                  </h3>
                  <p className="text-[11px] text-slate-400 font-mono">
                    İSG (3 Yıl) & Periyodik Muayene (5 Yıl) Otomatik Takip Servisi
                  </p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => setIsTestMailModalOpen(false)}
                className="p-1.5 text-slate-400 hover:text-white rounded-xl hover:bg-slate-800"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* Content */}
            {!testMailAuthorized ? (
              <form onSubmit={handleVerifyTestMailPin} className="p-6 space-y-4">
                <div className="flex items-center gap-2 bg-amber-500/10 border border-amber-500/20 p-3 rounded-2xl text-amber-300 text-xs">
                  <Lock className="w-4 h-4 shrink-0 text-amber-400" />
                  <span>Test maili yetkilendirme doğrulama adımı. Lütfen yetkili şifresini giriniz.</span>
                </div>

                <div>
                  <label className="block text-xs font-bold uppercase font-mono text-slate-300 mb-2">
                    YETKİLİ ŞİFRESİ
                  </label>
                  <input
                    type="password"
                    autoFocus
                    placeholder="****"
                    value={testMailPinInput}
                    onChange={(e) => setTestMailPinInput(e.target.value)}
                    className="w-full px-4 py-3 bg-slate-950 border border-slate-800 rounded-2xl font-mono text-lg font-bold text-center tracking-widest outline-none focus:border-amber-500"
                  />
                  {testMailPinError && (
                    <span className="text-red-400 text-xs font-bold mt-1 block">Hatalı Şifre! Lütfen tekrar deneyiniz.</span>
                  )}
                </div>

                <div className="flex items-center justify-end gap-2 pt-2">
                  <button
                    type="button"
                    onClick={() => setIsTestMailModalOpen(false)}
                    className="px-4 py-2.5 bg-slate-800 hover:bg-slate-700 text-slate-300 text-xs font-bold rounded-xl"
                  >
                    Vazgeç
                  </button>
                  <button
                    type="submit"
                    className="px-5 py-2.5 bg-amber-600 hover:bg-amber-500 text-white font-black text-xs rounded-xl shadow cursor-pointer"
                  >
                    ŞİFREYİ DOĞRULA
                  </button>
                </div>
              </form>
            ) : (
              <div className="p-6 space-y-4 text-xs">
                <div className="bg-slate-950 border border-slate-800 p-4 rounded-2xl space-y-2">
                  <div className="flex items-center justify-between text-[11px] font-mono border-b border-slate-800 pb-2">
                    <span className="text-slate-400">HEDEF E-POSTA:</span>
                    <span className="font-bold text-emerald-400">ormanhavacilik.bakimsube@gmail.com</span>
                  </div>
                  <div className="flex items-center justify-between text-[11px] font-mono border-b border-slate-800 pb-2">
                    <span className="text-slate-400">ZAMANLANMIŞ MAİL TARİHİ:</span>
                    <span className="font-bold text-amber-400 font-mono flex items-center gap-1">
                      <Clock className="w-3.5 h-3.5 text-amber-400" />
                      {getNextMonthFirstDayStr()} (BİR SONRAKİ AYIN İLK GÜNÜ)
                    </span>
                  </div>
                  <div className="flex items-center justify-between text-[11px] font-mono">
                    <span className="text-slate-400">UYARISI YAKLAŞAN / DOLAN PERSONEL:</span>
                    <span className="font-extrabold text-white bg-amber-500/20 px-2 py-0.5 rounded border border-amber-500/30 font-mono">
                      {expiringRecords.length} Personel
                    </span>
                  </div>
                </div>

                {expiringRecords.length > 0 ? (
                  <div className="max-h-48 overflow-y-auto space-y-2 pr-1">
                    {expiringRecords.map((item, idx) => (
                      <div key={idx} className="bg-slate-950 border border-slate-800 p-2.5 rounded-xl flex items-center justify-between text-[11px]">
                        <div>
                          <span className="font-bold text-white block">{item.person.adiSoyadi}</span>
                          <span className="text-[10px] text-amber-400 font-mono">
                            {item.type === 'ISG' ? 'İSG Eğitimi (3 Yıl)' : 'Periyodik Muayene (5 Yıl)'} • Bitiş: {item.expiryDateStr}
                          </span>
                        </div>
                        <span className="font-bold font-mono text-amber-300 bg-amber-950/80 px-2 py-1 rounded border border-amber-700/50">
                          {item.daysLeft <= 0 ? 'Süre Doldu' : `${item.daysLeft} Gün Kaldı`}
                        </span>
                      </div>
                    ))}
                  </div>
                ) : (
                  <div className="bg-emerald-950/40 border border-emerald-800/40 p-3 rounded-2xl text-emerald-300 text-center font-bold">
                    ✅ Süresi 1 ay kalan veya dolan herhangi bir İSG/Muayene kaydı bulunmamaktadır.
                  </div>
                )}

                <div className="pt-2 flex items-center justify-end gap-2">
                  <button
                    type="button"
                    onClick={() => setIsTestMailModalOpen(false)}
                    className="px-4 py-2.5 bg-slate-800 hover:bg-slate-700 text-slate-300 font-bold rounded-xl"
                  >
                    Kapat
                  </button>
                  <button
                    type="button"
                    disabled={isSimulatingMail}
                    onClick={handleSimulateSendMail}
                    className="px-5 py-2.5 bg-emerald-600 hover:bg-emerald-500 text-white font-black rounded-xl shadow cursor-pointer flex items-center gap-2 active:scale-95"
                  >
                    {isSimulatingMail ? (
                      <>
                        <RefreshCw className="w-4 h-4 animate-spin" />
                        <span>SİMÜLE EDİLİYOR...</span>
                      </>
                    ) : (
                      <>
                        <Mail className="w-4 h-4" />
                        <span>ŞİMDİ SİMÜLE ET & GÖNDER</span>
                      </>
                    )}
                  </button>
                </div>
              </div>
            )}

          </div>
        </div>
      )}

      {/* 4. YENİ PERSONEL EKLEME MODALI */}
      {isAddPersonModalOpen && (
        <div className="fixed inset-0 z-[700] bg-black/70 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-white text-slate-900 rounded-2xl shadow-2xl border border-slate-200 w-full max-w-lg overflow-visible animate-scale-up">
            
            <div className="bg-slate-900 text-white px-6 py-4 rounded-t-2xl flex items-center justify-between">
              <div className="flex items-center gap-2.5">
                <GraduationCap className="w-5 h-5 text-emerald-400" />
                <h3 className="font-black text-sm uppercase tracking-wider font-mono">
                  Yeni Personel ve Eğitim Kaydı
                </h3>
              </div>
              <button
                type="button"
                onClick={() => setIsAddPersonModalOpen(false)}
                className="text-slate-400 hover:text-white p-1 rounded-lg"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <form onSubmit={handleSaveNewPerson} className="p-6 space-y-3.5 text-xs">
              <div>
                <label className="block font-bold text-slate-700 uppercase font-mono mb-1">
                  Personel Adı Soyadi *
                </label>
                <input
                  type="text"
                  required
                  placeholder="Örn: AHMED SELİM KANAY"
                  value={newPersonName}
                  onChange={(e) => setNewPersonName(e.target.value)}
                  className="w-full px-3.5 py-2.5 bg-slate-50 border border-slate-200 rounded-xl font-bold uppercase text-slate-800 outline-none"
                />
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block font-bold text-slate-700 uppercase font-mono mb-1 text-[10px]">
                    İSG Tarihi
                  </label>
                  <input
                    type="text"
                    placeholder="GG.AA.YYYY"
                    value={newPersonIsgDate}
                    onChange={(e) => setNewPersonIsgDate(e.target.value)}
                    className="w-full px-2.5 py-2 bg-slate-50 border border-slate-200 rounded-xl font-bold text-slate-800 outline-none font-mono text-center"
                  />
                </div>
                <div>
                  <label className="block font-bold text-slate-700 uppercase font-mono mb-1 text-[10px]">
                    CBİKO Tarihi
                  </label>
                  <input
                    type="text"
                    placeholder="GG.AA.YYYY"
                    value={newPersonCbikoDate}
                    onChange={(e) => setNewPersonCbikoDate(e.target.value)}
                    className="w-full px-2.5 py-2 bg-slate-50 border border-slate-200 rounded-xl font-bold text-slate-800 outline-none font-mono text-center"
                  />
                </div>
              </div>

              <div>
                <label className="block font-bold text-slate-700 uppercase font-mono mb-1">
                  İlk Mesleki Eğitimin Adı (Opsiyonel)
                </label>
                <input
                  type="text"
                  placeholder="Örn: BELL 429 (PWC PW207D1/D2) EASA B1.3"
                  value={newInitialTrainingName}
                  onChange={(e) => setNewInitialTrainingName(e.target.value)}
                  className="w-full px-3.5 py-2.5 bg-slate-50 border border-slate-200 rounded-xl font-bold uppercase text-slate-800 outline-none"
                />
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block font-bold text-slate-700 uppercase font-mono mb-1">
                    Eğitim Tarihi
                  </label>
                  <input
                    type="text"
                    placeholder="GG.AA.YYYY"
                    value={newInitialTrainingDate}
                    onChange={(e) => setNewInitialTrainingDate(e.target.value)}
                    className="w-full px-3.5 py-2.5 bg-slate-50 border border-slate-200 rounded-xl font-bold text-slate-800 outline-none font-mono text-center"
                  />
                </div>

                <div className="relative">
                  <label className="block font-bold text-slate-700 uppercase font-mono mb-1">
                    Eğitim Veren Kurum
                  </label>
                  <input
                    type="text"
                    placeholder="Örn: BELL TRAINING ACADEMY"
                    value={newInitialTrainingProvider}
                    onFocus={() => setActiveProviderField('new')}
                    onBlur={() => setTimeout(() => setActiveProviderField(prev => (prev === 'new' ? null : prev)), 150)}
                    onChange={(e) => {
                      setNewInitialTrainingProvider(e.target.value);
                      setActiveProviderField('new');
                    }}
                    className="w-full px-3.5 py-2.5 bg-slate-50 border border-slate-200 rounded-xl font-bold uppercase text-slate-800 outline-none focus:border-emerald-500"
                  />
                  {activeProviderField === 'new' && getFilteredProviders(newInitialTrainingProvider).length > 0 && (
                    <div className="absolute top-full left-0 right-0 mt-1 z-50 bg-white border border-emerald-300 rounded-xl shadow-xl max-h-40 overflow-y-auto divide-y divide-slate-100">
                      {getFilteredProviders(newInitialTrainingProvider).map((prov, pIdx) => (
                        <button
                          key={pIdx}
                          type="button"
                          onMouseDown={(e) => {
                            e.preventDefault();
                            setNewInitialTrainingProvider(prov);
                            setActiveProviderField(null);
                          }}
                          className="w-full text-left px-3 py-2 text-xs font-bold uppercase text-slate-800 hover:bg-emerald-50 hover:text-emerald-900 flex items-center gap-1.5 cursor-pointer transition"
                        >
                          <Building2 className="w-3.5 h-3.5 text-emerald-600 shrink-0" />
                          <span className="truncate">{prov}</span>
                        </button>
                      ))}
                    </div>
                  )}
                </div>
              </div>

              <div className="pt-3 border-t border-slate-100 flex items-center justify-end gap-2.5">
                <button
                  type="button"
                  onClick={() => setIsAddPersonModalOpen(false)}
                  className="px-4 py-2 bg-slate-100 text-slate-700 font-bold rounded-xl"
                >
                  Vazgeç
                </button>
                <button
                  type="submit"
                  className="px-5 py-2 bg-emerald-600 hover:bg-emerald-500 text-white font-black rounded-xl shadow cursor-pointer"
                >
                  Kaydet
                </button>
              </div>
            </form>

          </div>
        </div>
      )}

      {/* 5. EĞİTİM DÜZENLEME MODALI */}
      {editingTrainingInfo && (
        <div className="fixed inset-0 z-[750] bg-black/70 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-white text-slate-900 rounded-2xl shadow-2xl border border-slate-200 w-full max-w-lg overflow-visible animate-scale-up">
            <div className="bg-slate-900 text-white px-6 py-4 rounded-t-2xl flex items-center justify-between">
              <div className="flex items-center gap-2.5">
                <Edit2 className="w-5 h-5 text-blue-400" />
                <h3 className="font-black text-sm uppercase tracking-wider font-mono">
                  Mesleki Eğitim Kaydını Düzenle
                </h3>
              </div>
              <button
                type="button"
                onClick={() => setEditingTrainingInfo(null)}
                className="text-slate-400 hover:text-white p-1 rounded-lg"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <form onSubmit={handleSaveEditedTraining} className="p-6 space-y-4 text-xs">
              <div>
                <label className="block font-bold text-slate-700 uppercase font-mono mb-1">
                  Mesleki Eğitimin Adı *
                </label>
                <input
                  type="text"
                  required
                  placeholder="Eğitimin Tam Adını Yazınız..."
                  value={editTrainingName}
                  onChange={(e) => setEditTrainingName(e.target.value)}
                  className="w-full px-3.5 py-2.5 bg-slate-50 border border-slate-200 rounded-xl font-bold uppercase text-slate-800 outline-none"
                />
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block font-bold text-slate-700 uppercase font-mono mb-1">
                    Eğitim Tarihi
                  </label>
                  <input
                    type="text"
                    placeholder="GG.AA.YYYY"
                    value={editTrainingDate}
                    onChange={(e) => setEditTrainingDate(e.target.value)}
                    className="w-full px-3.5 py-2.5 bg-slate-50 border border-slate-200 rounded-xl font-bold text-slate-800 outline-none font-mono text-center"
                  />
                </div>

                <div className="relative">
                  <label className="block font-bold text-slate-700 uppercase font-mono mb-1">
                    Eğitim Veren Kurum
                  </label>
                  <input
                    type="text"
                    placeholder="Veren Kurum / Personel"
                    value={editTrainingProvider}
                    onFocus={() => setActiveProviderField('edit')}
                    onBlur={() => setTimeout(() => setActiveProviderField(prev => (prev === 'edit' ? null : prev)), 150)}
                    onChange={(e) => {
                      setEditTrainingProvider(e.target.value);
                      setActiveProviderField('edit');
                    }}
                    className="w-full px-3.5 py-2.5 bg-slate-50 border border-slate-200 rounded-xl font-bold uppercase text-slate-800 outline-none focus:border-blue-500"
                  />
                  {activeProviderField === 'edit' && getFilteredProviders(editTrainingProvider).length > 0 && (
                    <div className="absolute top-full left-0 right-0 mt-1 z-50 bg-white border border-blue-300 rounded-xl shadow-xl max-h-40 overflow-y-auto divide-y divide-slate-100">
                      {getFilteredProviders(editTrainingProvider).map((prov, pIdx) => (
                        <button
                          key={pIdx}
                          type="button"
                          onMouseDown={(e) => {
                            e.preventDefault();
                            setEditTrainingProvider(prov);
                            setActiveProviderField(null);
                          }}
                          className="w-full text-left px-3 py-2 text-xs font-bold uppercase text-slate-800 hover:bg-blue-50 hover:text-blue-900 flex items-center gap-1.5 cursor-pointer transition"
                        >
                          <Building2 className="w-3.5 h-3.5 text-blue-600 shrink-0" />
                          <span className="truncate">{prov}</span>
                        </button>
                      ))}
                    </div>
                  )}
                </div>
              </div>

              <div className="pt-4 border-t border-slate-100 flex items-center justify-end gap-2.5">
                <button
                  type="button"
                  onClick={() => setEditingTrainingInfo(null)}
                  className="px-4 py-2 bg-slate-100 text-slate-700 font-bold rounded-xl"
                >
                  Vazgeç
                </button>
                <button
                  type="submit"
                  className="px-5 py-2 bg-blue-600 hover:bg-blue-500 text-white font-black rounded-xl shadow cursor-pointer"
                >
                  Güncelle & Drive&apos;a Kaydet
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* 6. PERSONEL BİLGİLERİ VE TARİHLERİ DÜZENLEME MODALI */}
      {editingPersonInfo && (
        <div className="fixed inset-0 z-[750] bg-black/70 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-white text-slate-900 rounded-2xl shadow-2xl border border-slate-200 w-full max-w-md overflow-hidden animate-scale-up">
            <div className="bg-slate-900 text-white px-6 py-4 flex items-center justify-between">
              <div className="flex items-center gap-2.5">
                <User className="w-5 h-5 text-emerald-400" />
                <h3 className="font-black text-sm uppercase tracking-wider font-mono">
                  Personel & Tarih Bilgilerini Düzenle
                </h3>
              </div>
              <button
                type="button"
                onClick={() => setEditingPersonInfo(null)}
                className="text-slate-400 hover:text-white p-1 rounded-lg"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <form onSubmit={handleSaveEditedPerson} className="p-6 space-y-3.5 text-xs">
              <div>
                <label className="block font-bold text-slate-700 uppercase font-mono mb-1">
                  Sıra No
                </label>
                <input
                  type="number"
                  value={editPersonSiraNo}
                  onChange={(e) => setEditPersonSiraNo(parseInt(e.target.value) || 1)}
                  className="w-full px-3.5 py-2 bg-slate-50 border border-slate-200 rounded-xl font-bold font-mono text-slate-800 outline-none"
                />
              </div>

              <div>
                <label className="block font-bold text-slate-700 uppercase font-mono mb-1">
                  Personel Adı Soyadı *
                </label>
                <input
                  type="text"
                  required
                  placeholder="Personel Adı Soyadı..."
                  value={editPersonName}
                  onChange={(e) => setEditPersonName(e.target.value)}
                  className="w-full px-3.5 py-2 bg-slate-50 border border-slate-200 rounded-xl font-bold uppercase text-slate-800 outline-none"
                />
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block font-bold text-slate-700 uppercase font-mono mb-1 text-[10px]">
                    İSG Tarihi
                  </label>
                  <input
                    type="text"
                    placeholder="GG.AA.YYYY"
                    value={editPersonIsgDate}
                    onChange={(e) => setEditPersonIsgDate(e.target.value)}
                    className="w-full px-2 py-2 bg-slate-50 border border-slate-200 rounded-xl font-bold font-mono text-slate-800 text-center outline-none"
                  />
                </div>

                <div>
                  <label className="block font-bold text-slate-700 uppercase font-mono mb-1 text-[10px]">
                    CBİKO Tarihi
                  </label>
                  <input
                    type="text"
                    placeholder="GG.AA.YYYY"
                    value={editPersonCbikoDate}
                    onChange={(e) => setEditPersonCbikoDate(e.target.value)}
                    className="w-full px-2 py-2 bg-slate-50 border border-slate-200 rounded-xl font-bold font-mono text-slate-800 text-center outline-none"
                  />
                </div>
              </div>

              <div className="pt-4 border-t border-slate-100 flex items-center justify-end gap-2.5">
                <button
                  type="button"
                  onClick={() => setEditingPersonInfo(null)}
                  className="px-4 py-2 bg-slate-100 text-slate-700 font-bold rounded-xl"
                >
                  Vazgeç
                </button>
                <button
                  type="submit"
                  className="px-5 py-2 bg-emerald-600 hover:bg-emerald-500 text-white font-black rounded-xl shadow cursor-pointer"
                >
                  Güncelle & Drive&apos;a Kaydet
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

    </div>
  );
};
