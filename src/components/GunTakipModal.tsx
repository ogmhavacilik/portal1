import React, { useState, useMemo, useEffect } from 'react';
import {
  X,
  Clock,
  Plus,
  Trash2,
  Download,
  Save,
  CheckCircle2,
  AlertTriangle,
  Search,
  RefreshCw,
  Eye,
  Bell,
  ShieldCheck,
  ShieldAlert,
  Boxes,
  FileSpreadsheet,
  Filter,
  Check,
  Edit2,
  Info,
  ChevronDown
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { TEKNISYEN_PERSONEL_SCRIPT_URL } from '../utils/personnelData';

export interface GunTakipBirimRow {
  birim: string;
  adSoyad: string;
  eposta: string;
  norm?: string;
  mail90?: string;
  mail60?: string;
  mail30?: string;
}

export interface ApproachingItem {
  unitKey: string;
  unitName: string;
  subSection: string;
  techName: string;
  partNo: string;
  seriNo: string;
  miktar: string;
  location: string;
  status: string;
  sonTarih: string;
  gelecekTarih: string;
  daysDiff: number;
  lastMailSentDate?: string;
  isNew: boolean;
  rowIndex: number;
}

export interface BirimStats {
  totalCount: number;
  approachingItems: ApproachingItem[];
  overdueCount: number;
  upcomingCount: number;
}

interface GunTakipModalProps {
  isOpen: boolean;
  onClose: () => void;
  sorumlular: GunTakipBirimRow[];
  onSaveSorumlular: (updated: GunTakipBirimRow[]) => Promise<void>;
  isSaving: boolean;
  showNotification: (msg: string, type?: 'success' | 'error' | 'info' | 'warning') => void;
  unitDataMap: {
    bell429?: any[][];
    at802?: any[][];
    t70?: any[][];
    t70_bumbi_backet?: any[][];
    t70_helitak?: any[][];
    b360?: any[][];
    c650?: any[][];
    hangar?: any[][];
    kara_araclari?: any[][];
    yasam_destek?: any[];
    personel_bilgi?: any[];
    personel_egitim?: any[];
  };
  onUpdateEquipmentMailDate?: (unitKey: string, rowIndex: number, mailColIndex: number, dateStr: string) => void;
  googleScriptUrl: string;
}

// Standart Norm Seçenekleri Listesi
export const NORM_OPTIONS = [
  {
    id: 'depo_omur_ve_kritik_stok',
    label: 'Depo: Ömür Takip (Tarih Bazlı - 3 Ay Kala) & Kritik Stok (Kritik +1 Adet Kala)',
    desc: 'Sarf/kimyasal SKT/ömrü 3 ay kalanlar veya stok miktarı kritik stok +1 adet kalanlar'
  },
  {
    id: 'yer_destek_kalibrasyon',
    label: 'Yer Destek & Özel Alet: Kalibrasyon / Bakım Tarih Takip (3 Ay Kala)',
    desc: 'Hava araçları yer destek ekipmanlarının gelecek kalibrasyon ve periyodik bakım tarihine 3 ay kalanlar'
  },
  {
    id: 'ozel_alet_kalibrasyon',
    label: 'Özel Aletler: Kalibrasyon & Test Aletleri Tarih Takibi (3 Ay Kala)',
    desc: 'Özel aletlerin kalibrasyon ve test sertifika yenilemesine 3 ay kalanlar'
  },
  {
    id: 'hangar_yer_destek',
    label: 'Hangar Yer Destek: Son Kontrol / Kalibrasyon / Bakım Takip (3 Ay Kala)',
    desc: 'Hangar genel ekipmanlarının son kontrol, kalibrasyon ve bakım süresine 3 ay kalanlar'
  },
  {
    id: 'kara_araclari_bakim_km',
    label: 'Kara Araçları: Bakım Tarihi & Sonraki Bakım / KM Takip (3 Ay Kala)',
    desc: 'Kara araçlarının periyodik bakım tarihi veya KM yaklaşım limitine 3 ay kalanlar'
  },
  {
    id: 'at802_yasam_destek',
    label: 'AT-802 Yaşam Destek: Can Yeleği, Spare Air & Kask Kit SKT Takip (3 Ay Kala)',
    desc: 'Zimmetli can yeleği, spare air ve kask kitlerinin periyodik muayene ve SKT süresine 3 ay kalanlar'
  },
  {
    id: 'personel_bilgi_muayene',
    label: 'Personel Bilgi Sistemi: OGM Periyodik Muayene Tarih Takip (1 Ay Kala)',
    desc: 'Personel periyodik sağlık muayenesi geçerlilik süresine 1 ay kalanlar'
  },
  {
    id: 'personel_egitim_isg',
    label: 'Personel Eğitim Bilgi Sistemi: 3 Yıllık İSG Eğitim Tarihi Takip (1 Ay Kala)',
    desc: '3 yıllık periyodik İSG eğitiminin yenilenmesine 1 ay kalan personel'
  },
  {
    id: 'yangin_ve_helitak',
    label: 'Yangın Ekipmanları / Helitak / Bumbi: Kalibrasyon & Bakım Takip (3 Ay Kala)',
    desc: 'Bumbi backet ve helitak yangın ekipmanlarının periyodik muayenesine 3 ay kalanlar'
  }
];

// Türkçe karakterleri normalize et
export const normalizeTurkishStr = (str: string): string => {
  if (!str) return "";
  return str
    .replace(/İ/g, "I")
    .replace(/ı/g, "i")
    .replace(/Ğ/g, "G")
    .replace(/ğ/g, "g")
    .replace(/Ü/g, "U")
    .replace(/ü/g, "u")
    .replace(/Ş/g, "S")
    .replace(/ş/g, "s")
    .replace(/Ö/g, "O")
    .replace(/ö/g, "o")
    .replace(/Ç/g, "C")
    .replace(/ç/g, "c")
    .toUpperCase();
};

// Birim adından varsayılan takip normu metni üretici
export const getDefaultNormForBirim = (birimName: string): string => {
  const norm = normalizeTurkishStr(birimName);
  if (norm.includes('YASAM DESTEK') || norm.includes('YAŞAM DESTEK')) {
    return 'AT-802 Yaşam Destek: Can Yeleği, Spare Air & Kask Kit SKT Takip (3 Ay Kala)';
  }
  if (norm.includes('PERSONEL BILGI') || norm.includes('PERSONEL BİLGİ')) {
    return 'Personel Bilgi Sistemi: OGM Periyodik Muayene Tarih Takip (1 Ay Kala)';
  }
  if (norm.includes('PERSONEL EGITIM') || norm.includes('PERSONEL EĞİTİM')) {
    return 'Personel Eğitim Bilgi Sistemi: 3 Yıllık İSG Eğitim Tarihi Takip (1 Ay Kala)';
  }
  if (norm.includes('DEPO')) {
    return 'Depo: Ömür Takip (Tarih Bazlı - 3 Ay Kala) & Kritik Stok (Kritik +1 Adet Kala)';
  }
  if (norm.includes('KARA')) {
    return 'Kara Araçları: Bakım Tarihi & Sonraki Bakım / KM Takip (3 Ay Kala)';
  }
  if (norm.includes('HANGAR')) {
    return 'Hangar Yer Destek: Son Kontrol / Kalibrasyon / Bakım Takip (3 Ay Kala)';
  }
  if (norm.includes('OZEL') || norm.includes('ÖZEL')) {
    return 'Özel Aletler: Kalibrasyon & Test Aletleri Tarih Takibi (3 Ay Kala)';
  }
  if (norm.includes('BUMBI') || norm.includes('BUMBİ') || norm.includes('HELITAK') || norm.includes('HELİTAK')) {
    return 'Yangın Ekipmanları / Helitak / Bumbi: Kalibrasyon & Bakım Takip (3 Ay Kala)';
  }
  return 'Yer Destek & Özel Alet: Kalibrasyon / Bakım Tarih Takip (3 Ay Kala)';
};

// Başlık satırı kontrolü
export const isHeaderLikeRow = (r: any[]): boolean => {
  if (!Array.isArray(r) || r.length === 0) return true;
  const c0 = String(r[0] || "").trim().toUpperCase();
  const c1 = String(r[1] || "").trim().toUpperCase();
  const c2 = String(r[2] || "").trim().toUpperCase();
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
  return false;
};

// Satırın bölümünü belirle
export const getRowSection = (row: string[]): 'yer_destek' | 'ozel_alet' | 'depo_sarf' | 'depo_kimyasal' => {
  if (!row || !Array.isArray(row)) return 'yer_destek';

  const candidateIndices = [13, row.length - 1, 12, row.length - 2, 11];
  for (const idx of candidateIndices) {
    if (idx >= 0 && idx < row.length) {
      const rawVal = String(row[idx] || "").trim();
      if (/\d{4}-\d{2}-\d{2}/.test(rawVal) || /\d{2}\.\d{2}\.\d{4}/.test(rawVal)) continue;

      const raw = normalizeTurkishStr(rawVal).toLowerCase().trim();
      if (raw === 'depo_sarf' || raw === 'sarf_parca' || raw === 'sarf_ve_parca' || raw === 'sarf' || raw === 'sarf ve parca deposu' || raw === 'sarf deposu' || raw === 'parca deposu') return 'depo_sarf';
      if (raw === 'depo_kimyasal' || raw === 'kimyasal_depo' || raw === 'kimyasal' || raw === 'kimyasal depo' || raw === 'kimyasal maddeler' || raw === 'madeni yag') return 'depo_kimyasal';
      if (raw === 'ozel_alet' || raw === 'ozel_aletler' || raw === 'ozel_bakim' || raw === 'ozel bakim ve test aletleri' || raw === 'ozel bakim aletleri' || raw === 'ozel alet' || raw === 'ozel bakim ve test') return 'ozel_alet';
      if (raw === 'yer_destek' || raw === 'yer_destek_techizat' || raw === 'yer destek techizatlari' || raw === 'yer destek') return 'yer_destek';
    }
  }

  const combined = normalizeTurkishStr(row.join(" ")).toUpperCase();
  if (combined.includes("##DEPO_SARF##") || combined.includes("[SARF DEPO]") || combined.includes("[PARCA DEPO]")) return 'depo_sarf';
  if (combined.includes("##DEPO_KIMYASAL##") || combined.includes("[KIMYASAL DEPO]")) return 'depo_kimyasal';
  if (combined.includes("##OZEL_ALET##") || combined.includes("[OZEL ALET]") || combined.includes("OZEL BAKIM VE TEST ALETLERI") || combined.includes("OZEL BAKIM ALETLERI") || combined.includes("OZEL ALETLER")) return 'ozel_alet';
  if (combined.includes("##YER_DESTEK##") || combined.includes("[YER DESTEK]") || combined.includes("YER DESTEK TECHIZATLARI")) return 'yer_destek';

  const itemName = normalizeTurkishStr(String(row[1] || "") + " " + String(row[2] || "") + " " + String(row[11] || "")).toUpperCase();

  if (itemName.includes("FILTRE") || itemName.includes("FILTER") || itemName.includes("CONTA") || itemName.includes("O-RING") || itemName.includes("BALATA") || itemName.includes("KECE") || itemName.includes("SEAL") || itemName.includes("SARF") || itemName.includes("CIVATA") || itemName.includes("SOMUN") || itemName.includes("PUL") || itemName.includes("RIVET") || itemName.includes("PERCIN")) {
    return 'depo_sarf';
  }

  if (itemName.includes("OZEL ALET") || itemName.includes("TEST ADAPTOR") || itemName.includes("TORK ADAPTOR") || itemName.includes("KOMPRESOR YIKAMA") || itemName.includes("AYAR MASTAR") || itemName.includes("BOSALTMA KAPAK TEST") || itemName.includes("CALIBRATION") || itemName.includes("TEST RIG") || itemName.includes("FIREGATE") || itemName.includes("OZEL BAKIM") || itemName.includes("ACHIOLCER") || itemName.includes("ACIOLCER") || itemName.includes("KALIBRASYON KIT") || itemName.includes("HAT BAKIM VE KALIBRASYON") || itemName.includes("TORQUE") || itemName.includes("MASTAR") || itemName.includes("ADAPTOR") || itemName.includes("TEST KIT") || itemName.includes("TORK ANAHTAR") || itemName.includes("MOTOR TEST") || itemName.includes("TORQ") || itemName.includes("WRENCH") || itemName.includes("PULLER") || itemName.includes("SOCKET") || itemName.includes("VOLTMETRE") || itemName.includes("MULTIMETRE") || itemName.includes("MANOMETRE")) {
    return 'ozel_alet';
  }

  return 'yer_destek';
};

// Türkçe tarih stringini JS Date nesnesine çevir
function parseTrDateObj(dateStr?: string): Date | null {
  if (!dateStr || dateStr === '-' || dateStr.trim() === '') return null;
  const clean = dateStr.trim();
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
  if (year < 100) year += 2000;
  return new Date(year, month, day);
}

function getFutureDateStr(dateStr: string, yearsToAdd: number): string {
  const dObj = parseTrDateObj(dateStr);
  if (!dObj) return dateStr;
  dObj.setFullYear(dObj.getFullYear() + yearsToAdd);
  const pad = (n: number) => String(n).padStart(2, '0');
  return `${pad(dObj.getDate())}.${pad(dObj.getMonth() + 1)}.${dObj.getFullYear()}`;
}

export const calculateDaysDiff = (dateStr: string): number | null => {
  if (!dateStr || typeof dateStr !== 'string') return null;
  const clean = dateStr.trim();
  if (!clean || clean === '-' || clean === '--' || clean.toLowerCase() === 'yok' || clean.toLowerCase().includes('belirtilme')) return null;

  const targetDate = parseTrDateObj(clean);
  if (!targetDate) return null;

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  targetDate.setHours(0, 0, 0, 0);

  const diffTime = targetDate.getTime() - today.getTime();
  return Math.ceil(diffTime / (1000 * 60 * 60 * 24));
};

export const calculateDaysDiffForPersonelMuayene = (dateStr: string): number | null => {
  const dObj = parseTrDateObj(dateStr);
  if (!dObj) return null;
  const expiryDate = new Date(dObj);
  expiryDate.setFullYear(expiryDate.getFullYear() + 5);
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  expiryDate.setHours(0, 0, 0, 0);
  const diffTime = expiryDate.getTime() - today.getTime();
  return Math.ceil(diffTime / (1000 * 60 * 60 * 24));
};

export const calculateDaysDiffForIsgEgitim = (dateStr: string): number | null => {
  const dObj = parseTrDateObj(dateStr);
  if (!dObj) return null;
  const expiryDate = new Date(dObj);
  expiryDate.setFullYear(expiryDate.getFullYear() + 3);
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  expiryDate.setHours(0, 0, 0, 0);
  const diffTime = expiryDate.getTime() - today.getTime();
  return Math.ceil(diffTime / (1000 * 60 * 60 * 24));
};

export const GunTakipModal: React.FC<GunTakipModalProps> = ({
  isOpen,
  onClose,
  sorumlular,
  onSaveSorumlular,
  isSaving,
  showNotification,
  unitDataMap,
  googleScriptUrl
}) => {
  const [list, setList] = useState<GunTakipBirimRow[]>(sorumlular);
  const [activeFleetFilter, setActiveFleetFilter] = useState<string>('all');
  const [searchQuery, setSearchQuery] = useState<string>('');

  // Live Drive / Server Fetching State
  const [isLiveFetching, setIsLiveFetching] = useState<boolean>(true);
  const [liveServerData, setLiveServerData] = useState<Record<string, any[][]>>({});

  // Inline Cell Edit State (Double-click activation)
  const [editingCell, setEditingCell] = useState<{ index: number; field: 'birim' | 'adSoyad' | 'eposta' | 'mail90' } | null>(null);

  // Yaşam Destek & Personel Canlı Veri State'leri
  const [yasamDestekData, setYasamDestekData] = useState<any[]>([]);
  const [personelData, setPersonelData] = useState<any[]>([]);

  // New Sub-Unit Dialog
  const [isAddDialogOpen, setIsAddDialogOpen] = useState<boolean>(false);
  const [newBirimName, setNewBirimName] = useState<string>('');
  const [newAdSoyad, setNewAdSoyad] = useState<string>('');
  const [newEposta, setNewEposta] = useState<string>('orhavak.bakimsube@gmail.com');
  const [newNorm, setNewNorm] = useState<string>(NORM_OPTIONS[0].label);

  // Detail Modal State
  const [selectedBirimDetails, setSelectedBirimDetails] = useState<{
    birim: GunTakipBirimRow;
    stats: BirimStats;
  } | null>(null);

  const [detailFilter, setDetailFilter] = useState<'all' | 'overdue' | 'upcoming'>('all');
  const [detailSearch, setDetailSearch] = useState<string>('');

  // Background auto check state
  const [isAutoChecking, setIsAutoChecking] = useState<boolean>(false);

  // Sync internal list if props change
  useEffect(() => {
    if (sorumlular && sorumlular.length > 0) {
      setList(sorumlular);
    }
  }, [sorumlular]);

  // CANLI VERİLERİ DRIVEDEN VE SERVER'DAN ÇEKME FONKSİYONU
  const fetchAllLiveDriveData = async () => {
    setIsLiveFetching(true);
    try {
      // 1. Server-side live scan API
      const serverRes = await fetch('/api/gun-takip-live-scan').catch(() => null);
      if (serverRes && serverRes.ok) {
        const json = await serverRes.json();
        if (json && json.status === 'success' && json.data) {
          setLiveServerData(json.data);
        }
      }

      // 2. Yaşam Destek Verileri
      const ydRes = await fetch('/api/yasam-destek').catch(() => null);
      if (ydRes && ydRes.ok) {
        const json = await ydRes.json();
        if (json && Array.isArray(json.records)) {
          setYasamDestekData(json.records);
        }
      }

      // 3. Personel Bilgi & Eğitim Verileri
      const persRes = await fetch(`${TEKNISYEN_PERSONEL_SCRIPT_URL}?sheet=Personel&action=getPersonel`).catch(() => null);
      if (persRes && persRes.ok) {
        const json = await persRes.json();
        if (json && json.status === 'success' && Array.isArray(json.data)) {
          setPersonelData(json.data);
        }
      } else {
        const localPers = await fetch('/api/yasam-destek/personnel-list').catch(() => null);
        if (localPers && localPers.ok) {
          const json = await localPers.json();
          if (json && Array.isArray(json.data)) setPersonelData(json.data);
        }
      }

      // 4. Canlı Gün Takip E-Tablosu Okuma
      if (googleScriptUrl) {
        try {
          const targetUrl = `${googleScriptUrl}?action=readSheet&sheetName=${encodeURIComponent("GÜN TAKİP")}&t=${Date.now()}`;
          const gtRes = await fetch(targetUrl);
          if (gtRes.ok) {
            const result = await gtRes.json();
            if (result && result.status === 'success' && Array.isArray(result.data)) {
              const mapped = result.data.map((row: any) => ({
                birim: String(row["SORUMLU BİRİM"] || row["Sorumlu Birim"] || row[0] || "").trim().toUpperCase(),
                adSoyad: String(row["ADI SOYADI"] || row["Adı Soyadı"] || row["AD SOYAD"] || row[1] || "").trim(),
                eposta: String(row["E-POSTA ADRESİ"] || row["E-posta Adresi"] || row["EPOSTA ADRESI"] || row[2] || "").trim(),
                norm: String(row["TAKİP EDİLECEK NORM"] || row["Norm"] || row[3] || "").trim(),
                mail90: String(row["SON BİLDİRİM TARİHİ"] || row["90 GÜN UYARISI MAİL GÖNDERİM TARİHİ"] || row[4] || "").trim(),
                mail60: String(row["60 GÜN UYARISI MAİL GÖNDERİM TARİHİ"] || row[5] || "").trim(),
                mail30: String(row["30 GÜN UYARISI MAİL GÖNDERİM TARİHİ"] || row[6] || "").trim()
              })).filter((item: any) => item.birim);

              if (mapped.length > 0) {
                setList(mapped);
              }
            }
          }
        } catch (e) {
          console.warn('GunTakip live sheet read error:', e);
        }
      }
    } catch (err) {
      console.warn('Live Drive Scan error:', err);
    } finally {
      setIsLiveFetching(false);
    }
  };

  // Modal her açıldığında canlı verileri otomatik çekip inceler
  useEffect(() => {
    if (isOpen) {
      fetchAllLiveDriveData();
    }
  }, [isOpen]);

  // Format date helper
  const getNowFormattedWithSeconds = (): string => {
    const now = new Date();
    const pad = (n: number) => String(n).padStart(2, '0');
    return `${pad(now.getDate())}.${pad(now.getMonth() + 1)}.${now.getFullYear()} ${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}`;
  };

  /**
   * GERÇEK VERİ EŞLEME VE DİNAMİK NORM KONTROLÜ:
   * Sorumlu birim adına ve seçili takip normuna göre canlı veriyi inceler, norm kurallarına uymayanları tespit eder.
   */
  const getBirimStats = (birimName: string, assignedNorm?: string): BirimStats => {
    const norm = normalizeTurkishStr(birimName);
    const activeNormStr = normalizeTurkishStr(assignedNorm || getDefaultNormForBirim(birimName));

    // ----------------------------------------------------
    // 1. PERSONEL BİLGİ SİSTEMİ (OGM Sağlık Muayenesi Takibi - 1 Ay / 30 Gün Kala)
    // ----------------------------------------------------
    if (activeNormStr.includes('PERSONEL BILGI') || activeNormStr.includes('MUAYENE') || norm.includes('PERSONEL BILGI') || (norm.includes('PERSONEL') && !norm.includes('EGITIM') && !norm.includes('EĞİTİM'))) {
      const approachingItems: ApproachingItem[] = [];
      const totalCount = personelData.length;

      personelData.forEach((p, rIdx) => {
        const name = p.adiSoyadi || p['ADI SOYADI'] || p.adSoyad || p[1] || 'Personel';
        const muayeneDateStr = p.periyodikMuayeneTarihi || p['PERİYODİK SAĞLIK MUAYENE TARİHİ'] || p['MUAYENE TARİHİ'] || p.muayeneTarihi || p[13] || '';

        if (!muayeneDateStr || muayeneDateStr === '-' || muayeneDateStr.trim() === '') return;

        const daysDiff = calculateDaysDiffForPersonelMuayene(muayeneDateStr);
        if (daysDiff !== null && daysDiff <= 30) {
          approachingItems.push({
            unitKey: 'personel_bilgi',
            unitName: 'PERSONEL BİLGİ SİSTEMİ',
            subSection: 'OGM SAĞLIK MUAYENESİ',
            techName: `${name} - Periyodik Sağlık Muayenesi`,
            partNo: p.sicilNo || p.tcKimlik || '-',
            seriNo: p.kadroUnvani || 'Teknisyen',
            miktar: '1',
            location: p.gorevYeri || 'Ankara',
            status: daysDiff <= 0 ? 'MUAYENE GÜNÜ GEÇTİ (KRİTİK)' : 'MUAYENE YAKLAŞIYOR (≤30 GÜN)',
            sonTarih: muayeneDateStr,
            gelecekTarih: getFutureDateStr(muayeneDateStr, 5),
            daysDiff,
            isNew: true,
            rowIndex: rIdx
          });
        }
      });

      approachingItems.sort((a, b) => a.daysDiff - b.daysDiff);
      const overdueCount = approachingItems.filter(i => i.daysDiff <= 0).length;
      const upcomingCount = approachingItems.filter(i => i.daysDiff > 0).length;
      return { totalCount, approachingItems, overdueCount, upcomingCount };
    }

    // ----------------------------------------------------
    // 2. PERSONEL EĞİTİM BİLİGİ SİSTEMİ (3 Yıllık İSG Eğitimi - 1 Ay / 30 Gün Kala)
    // ----------------------------------------------------
    if (activeNormStr.includes('EGITIM') || activeNormStr.includes('ISG') || norm.includes('EGITIM') || norm.includes('EĞİTİM')) {
      const approachingItems: ApproachingItem[] = [];
      const totalCount = personelData.length;

      personelData.forEach((p, rIdx) => {
        const name = p.adiSoyadi || p['ADI SOYADI'] || p.adSoyad || p[1] || 'Personel';
        const isgDateStr = p.isgEgitimTarihi || p['İSG EĞİTİM TARİHİ'] || p['ISG EGITIM TARIHI'] || p.egitimTarihi || p[14] || '';

        if (!isgDateStr || isgDateStr === '-' || isgDateStr.trim() === '') return;

        const daysDiff = calculateDaysDiffForIsgEgitim(isgDateStr);
        if (daysDiff !== null && daysDiff <= 30) {
          approachingItems.push({
            unitKey: 'personel_egitim',
            unitName: 'PERSONEL EĞİTİM BİLİGİ SİSTEMİ',
            subSection: 'İSG EĞİTİMİ (3 YILLIK)',
            techName: `${name} - İSG Periyodik Eğitimi`,
            partNo: p.sicilNo || p.tcKimlik || '-',
            seriNo: p.kadroUnvani || 'Teknisyen',
            miktar: '1',
            location: p.gorevYeri || 'Ankara',
            status: daysDiff <= 0 ? 'İSG EĞİTİM SÜRESİ DOLDU (KRİTİK)' : 'İSG EĞİTİMİ YAKLAŞIYOR (≤30 GÜN)',
            sonTarih: isgDateStr,
            gelecekTarih: getFutureDateStr(isgDateStr, 3),
            daysDiff,
            isNew: true,
            rowIndex: rIdx
          });
        }
      });

      approachingItems.sort((a, b) => a.daysDiff - b.daysDiff);
      const overdueCount = approachingItems.filter(i => i.daysDiff <= 0).length;
      const upcomingCount = approachingItems.filter(i => i.daysDiff > 0).length;
      return { totalCount, approachingItems, overdueCount, upcomingCount };
    }

    // ----------------------------------------------------
    // 3. AT-802 YAŞAM DESTEK (Can Yeleği, Spare Air, Kask Kit SKT & Kontrol - 3 Ay Kala)
    // ----------------------------------------------------
    if (activeNormStr.includes('YASAM DESTEK') || activeNormStr.includes('YAŞAM DESTEK') || norm.includes('YASAM DESTEK') || norm.includes('YAŞAM DESTEK')) {
      const approachingItems: ApproachingItem[] = [];
      const totalCount = yasamDestekData.length;

      yasamDestekData.forEach((item, rIdx) => {
        const person = item.personelAdi || item.personel || 'Teknisyen';
        const vestSn = item.lifeVestSn || '-';
        const airSn = item.spareAirSn || '-';
        const helmetSn = item.helmetKitSn || '-';
        const checkDate = item.bitisTarihi || item.baslangicTarihi || '';

        if (checkDate && checkDate !== 'DEVAM EDİYOR') {
          const daysDiff = calculateDaysDiff(checkDate);
          if (daysDiff !== null && daysDiff <= 90) {
            approachingItems.push({
              unitKey: 'at802_yasam_destek',
              unitName: 'AT-802 YAŞAM DESTEK',
              subSection: 'ZİMMET EKİPMAN KONTROLÜ',
              techName: `${person} (Vest: ${vestSn}, Air: ${airSn}, Helmet: ${helmetSn})`,
              partNo: vestSn,
              seriNo: airSn,
              miktar: '1',
              location: item.gorevYeri || 'Ankara',
              status: daysDiff <= 0 ? 'ZİMMET KONTROL GÜNÜ GEÇTİ' : 'ZİMMET KONTROLÜ YAKLAŞIYOR (≤90 GÜN)',
              sonTarih: item.baslangicTarihi || '-',
              gelecekTarih: checkDate,
              daysDiff,
              isNew: true,
              rowIndex: rIdx
            });
          }
        }
      });

      approachingItems.sort((a, b) => a.daysDiff - b.daysDiff);
      const overdueCount = approachingItems.filter(i => i.daysDiff <= 0).length;
      const upcomingCount = approachingItems.filter(i => i.daysDiff > 0).length;
      return { totalCount, approachingItems, overdueCount, upcomingCount };
    }

    // ----------------------------------------------------
    // 4. HAVA ARAÇLARI, DEPOLAR, TEÇHİZAT VE KARA ARAÇLARI
    // ----------------------------------------------------
    let targetFleetKey = '';
    if (norm.includes('BUMBI')) targetFleetKey = 't70_bumbi_backet';
    else if (norm.includes('HELITAK')) targetFleetKey = 't70_helitak';
    else if (norm.includes('T-70') || norm.includes('T70')) targetFleetKey = 't70';
    else if (norm.includes('BELL') || norm.includes('429')) targetFleetKey = 'bell429';
    else if (norm.includes('AT-802') || norm.includes('AT802')) targetFleetKey = 'at802';
    else if (norm.includes('B-360') || norm.includes('B360')) targetFleetKey = 'b360';
    else if (norm.includes('C-650') || norm.includes('C650')) targetFleetKey = 'c650';
    else if (norm.includes('HANGAR')) targetFleetKey = 'hangar';
    else if (norm.includes('KARA')) targetFleetKey = 'kara_araclari';

    let targetSection: 'all' | 'depo' | 'ozel_alet' | 'yer_destek' = 'all';
    if (activeNormStr.includes('DEPO') || norm.includes('DEPO')) {
      targetSection = 'depo';
    } else if (activeNormStr.includes('OZEL') || activeNormStr.includes('ÖZEL') || norm.includes('OZEL') || norm.includes('ÖZEL')) {
      targetSection = 'ozel_alet';
    } else if (activeNormStr.includes('YER DESTEK') || norm.includes('YER DESTEK')) {
      targetSection = 'yer_destek';
    }

    // Canlı sunucu verisi öncelikli, yoksa unitDataMap fallback
    const rawRows = (liveServerData as any)[targetFleetKey] || (unitDataMap as any)[targetFleetKey] || [];
    if (!Array.isArray(rawRows) || rawRows.length === 0) {
      return { totalCount: 0, approachingItems: [], overdueCount: 0, upcomingCount: 0 };
    }

    let totalCount = 0;
    const approachingItems: ApproachingItem[] = [];

    rawRows.forEach((row, rIdx) => {
      if (!row || !Array.isArray(row)) return;
      if (isHeaderLikeRow(row)) return;

      const sec = getRowSection(row);
      const isKara = targetFleetKey === 'kara_araclari';
      const isDepoRow = sec === 'depo_sarf' || sec === 'depo_kimyasal';

      if (targetSection === 'depo') {
        const loc = String(row[5] || '').toLowerCase();
        if (!isDepoRow && !loc.includes('depo')) return;
      } else if (targetSection === 'ozel_alet') {
        if (sec !== 'ozel_alet') return;
      } else if (targetSection === 'yer_destek') {
        if (sec !== 'yer_destek') return;
      }

      totalCount++;

      let techName = '';
      let partNo = '';
      let seriNo = '';
      let miktar = '';
      let location = '';
      let status = '';
      let sonTarih = '';
      let gelecekTarih = '';
      let isMuaf = false;
      let lastMailSentDate = '';
      let kritikStok = 0;
      let mevcutStok = 0;

      if (isKara) {
        techName = String(row[1] || 'Araç / Ekipman').trim();
        seriNo = String(row[2] || '-').trim();
        partNo = String(row[3] || '-').trim();
        location = String(row[4] || '-').trim();
        status = String(row[6] || 'Faal').trim();
        const bakimaTabi = String(row[7] || '').trim().toUpperCase();
        if (bakimaTabi === 'HAYIR' || bakimaTabi.includes('MUAF')) isMuaf = true;
        sonTarih = String(row[8] || '').trim();
        gelecekTarih = String(row[9] || '').trim();
        lastMailSentDate = String(row[12] || '').trim();
      } else if (isDepoRow) {
        techName = String(row[1] || 'Malzeme').trim();
        partNo = String(row[2] || '-').trim();
        seriNo = String(row[3] || '-').trim();
        miktar = String(row[4] || '1').trim();
        location = String(row[5] || '-').trim();
        status = String(row[6] || 'Faal').trim();
        const omurlu = String(row[7] || '').trim().toUpperCase();
        if (omurlu !== 'EVET') isMuaf = true;
        gelecekTarih = String(row[8] || '').trim();
        sonTarih = String(row[9] || '').trim();
        lastMailSentDate = String(row[11] || '').trim();

        mevcutStok = Number(row[4]) || 1;
        kritikStok = Number(row[15] || row[16] || 0);
      } else if (targetFleetKey === 'bell429') {
        techName = String(row[1] || 'Teçhizat').trim();
        partNo = String(row[2] || '-').trim();
        seriNo = String(row[3] || '-').trim();
        miktar = String(row[4] || '1').trim();
        location = String(row[5] || '-').trim();
        status = String(row[6] || 'Faal').trim();
        sonTarih = String(row[7] || '').trim();
        gelecekTarih = String(row[8] || '').trim();
        lastMailSentDate = String(row[11] || '').trim();
      } else {
        techName = String(row[1] || 'Teçhizat').trim();
        partNo = String(row[2] || '-').trim();
        seriNo = String(row[3] || '-').trim();
        miktar = String(row[4] || '1').trim();
        location = String(row[5] || '-').trim();
        status = String(row[6] || 'Faal').trim();
        const kalibTabi = String(row[7] || '').trim().toUpperCase();
        if (kalibTabi === 'HAYIR' || kalibTabi.includes('MUAF')) isMuaf = true;
        sonTarih = String(row[8] || '').trim();
        gelecekTarih = String(row[9] || '').trim();
        lastMailSentDate = String(row[12] || row[11] || '').trim();
      }

      // Depo Kritik Stok Takibi (1 Adet Kala veya Altına İnince Uyarır)
      if (isDepoRow && kritikStok > 0 && mevcutStok <= (kritikStok + 1)) {
        approachingItems.push({
          unitKey: targetFleetKey,
          unitName: targetFleetKey.toUpperCase(),
          subSection: 'KRİTİK STOK SEVİYESİ',
          techName: `${techName} (Kritik Stok Uyarısı)`,
          partNo: partNo || '-',
          seriNo: seriNo || '-',
          miktar: String(mevcutStok),
          location: location || '-',
          status: mevcutStok <= kritikStok ? `KRİTİK STOK SEVİYESİNE İNDİ (Mevcut: ${mevcutStok}, Limit: ${kritikStok})` : `KRİTİK STOĞA 1 ADET KALDI (Mevcut: ${mevcutStok}, Limit: ${kritikStok})`,
          sonTarih: '-',
          gelecekTarih: 'STOK LİMİTİ',
          daysDiff: mevcutStok <= kritikStok ? 0 : 1,
          lastMailSentDate,
          isNew: true,
          rowIndex: rIdx
        });
      }

      // Tarih Bazlı Takip (3 Ay / 90 Gün Kala Uyarı)
      if (isMuaf || !gelecekTarih) return;

      const daysDiff = calculateDaysDiff(gelecekTarih);
      if (daysDiff === null || daysDiff > 90) return;

      const isNew = !lastMailSentDate || lastMailSentDate.trim() === '' || lastMailSentDate.trim() === '-';

      approachingItems.push({
        unitKey: targetFleetKey,
        unitName: targetFleetKey.toUpperCase(),
        subSection: isDepoRow ? 'DEPO ÖMÜR' : (sec === 'ozel_alet' ? 'ÖZEL ALET' : 'YER DESTEK'),
        techName: techName || 'Teçhizat / Malzeme',
        partNo: partNo || '-',
        seriNo: seriNo || '-',
        miktar: miktar || '1',
        location: location || '-',
        status: status || 'Faal',
        sonTarih,
        gelecekTarih,
        daysDiff,
        lastMailSentDate,
        isNew,
        rowIndex: rIdx
      });
    });

    approachingItems.sort((a, b) => a.daysDiff - b.daysDiff);
    const overdueCount = approachingItems.filter(i => i.daysDiff <= 0).length;
    const upcomingCount = approachingItems.filter(i => i.daysDiff > 0).length;

    return {
      totalCount,
      approachingItems,
      overdueCount,
      upcomingCount
    };
  };

  // Filtered Birim List
  const filteredList = useMemo(() => {
    return list.filter(item => {
      if (activeFleetFilter !== 'all') {
        const norm = normalizeTurkishStr(item.birim);
        if (!norm.includes(normalizeTurkishStr(activeFleetFilter))) return false;
      }
      if (searchQuery.trim()) {
        const q = searchQuery.toLowerCase();
        const b = item.birim.toLowerCase();
        const a = item.adSoyad.toLowerCase();
        const e = item.eposta.toLowerCase();
        const n = (item.norm || '').toLowerCase();
        return b.includes(q) || a.includes(q) || e.includes(q) || n.includes(q);
      }
      return true;
    });
  }, [list, activeFleetFilter, searchQuery]);

  // Tüm filo genelindeki toplam günü geçenler ve yaklaşanlar
  const allOverdueItems = useMemo(() => {
    const map = new Map<string, ApproachingItem>();
    list.forEach(b => {
      const stats = getBirimStats(b.birim, b.norm);
      stats.approachingItems.filter(i => i.daysDiff <= 0).forEach(i => {
        const key = `${i.unitKey}_${i.partNo}_${i.seriNo}_${i.gelecekTarih}`;
        map.set(key, i);
      });
    });
    return Array.from(map.values());
  }, [list, unitDataMap, liveServerData, yasamDestekData, personelData]);

  const allApproachingItems = useMemo(() => {
    const map = new Map<string, ApproachingItem>();
    list.forEach(b => {
      const stats = getBirimStats(b.birim, b.norm);
      stats.approachingItems.filter(i => i.daysDiff > 0 && i.daysDiff <= 90).forEach(i => {
        const key = `${i.unitKey}_${i.partNo}_${i.seriNo}_${i.gelecekTarih}`;
        map.set(key, i);
      });
    });
    return Array.from(map.values());
  }, [list, unitDataMap, liveServerData, yasamDestekData, personelData]);

  // Arka plan otomatik kontrol
  const runSilentOverdueCheck = async (isManualTrigger = false) => {
    setIsAutoChecking(true);
    const nowTimestamp = getNowFormattedWithSeconds();

    try {
      await fetch('/api/trigger-auto-reminders', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          birimList: list,
          overdueItems: allOverdueItems,
          approachingItems: allApproachingItems
        })
      }).catch(e => console.warn('Local trigger reminder warn:', e));

      if (allOverdueItems.length > 0 && googleScriptUrl) {
        await fetch(googleScriptUrl, {
          method: 'POST',
          headers: { 'Content-Type': 'text/plain;charset=utf-8' },
          body: JSON.stringify({
            action: 'dailyReminderTrigger',
            reason: 'Otomatik Arka Plan Günü Geçen Kontrolü',
            timestamp: nowTimestamp
          })
        }).catch(e => console.warn('GAS auto trigger warn:', e));
      }

      if (isManualTrigger) {
        showNotification(`✅ Arka plan kontrolü tamamlandı. (${allOverdueItems.length} günü geçen, ${allApproachingItems.length} yaklaşan)`, 'success');
      }
    } catch (err: any) {
      console.warn('Auto check error:', err);
    } finally {
      setIsAutoChecking(false);
    }
  };

  useEffect(() => {
    if (isOpen) {
      runSilentOverdueCheck(false);
    }
  }, [isOpen]);

  // Yeni Birim Ekleme (E-Tabloya Anında Kaydeder)
  const handleAddNewBirim = async () => {
    if (!newBirimName.trim()) {
      showNotification('Lütfen geçerli bir birim veya alt birim adı giriniz.', 'warning');
      return;
    }
    const cleanName = newBirimName.trim().toUpperCase();
    if (list.some(item => item.birim.toUpperCase() === cleanName)) {
      showNotification('Bu birim veya alt birim listede zaten mevcut.', 'warning');
      return;
    }

    const normText = newNorm.trim() || getDefaultNormForBirim(cleanName);

    const newRow: GunTakipBirimRow = {
      birim: cleanName,
      adSoyad: newAdSoyad.trim() || 'Sorumlu Personel',
      eposta: newEposta.trim() || 'orhavak.bakimsube@gmail.com',
      norm: normText,
      mail90: ''
    };

    const updated = [...list, newRow];
    setList(updated);
    setNewBirimName('');
    setNewAdSoyad('');
    setNewNorm(NORM_OPTIONS[0].label);
    setIsAddDialogOpen(false);
    showNotification(`✅ "${cleanName}" başarıyla eklendi!`, 'success');

    await onSaveSorumlular(updated);
  };

  // Birim Silme (E-Tabloya Anında Kaydeder)
  const handleDeleteBirim = async (index: number) => {
    const itemToDelete = filteredList[index];
    if (!itemToDelete) return;
    if (confirm(`"${itemToDelete.birim}" birimini ve takibini silmek istediğinize emin misiniz?`)) {
      const updated = list.filter(item => item.birim !== itemToDelete.birim);
      setList(updated);
      showNotification(`"${itemToDelete.birim}" listeden silindi.`, 'info');
      await onSaveSorumlular(updated);
    }
  };

  // Inline hücre güncelleme ve E-Tabloya kaydetme
  const handleUpdateItemField = (itemBirim: string, field: 'birim' | 'adSoyad' | 'eposta' | 'mail90', value: string) => {
    setList(prev => {
      const updated = prev.map(p => p.birim === itemBirim ? { ...p, [field]: value } : p);
      return updated;
    });
  };

  const handleCellBlurOrEnter = async () => {
    setEditingCell(null);
    await onSaveSorumlular(list);
  };

  // Norm Seçim Listesinden Değiştirme (Anında Kaydeder & Canlı Analizi Yeniler)
  const handleSelectNorm = async (itemBirim: string, newNormLabel: string) => {
    const updatedList = list.map(p => p.birim === itemBirim ? { ...p, norm: newNormLabel } : p);
    setList(updatedList);
    await onSaveSorumlular(updatedList);
    showNotification(`✅ "${itemBirim}" için takip normu "${newNormLabel}" olarak güncellendi ve E-Tabloya kaydedildi.`, 'success');
  };

  // Tüm sorumluları Excel'e aktar
  const handleExportExcel = () => {
    try {
      const headers = [
        'SORUMLU BİRİM / ALT BİRİM',
        'ADI SOYADI',
        'E-POSTA ADRESİ',
        'TAKİP EDİLECEK NORM',
        'TOPLAM KAYITLI',
        'GÜNÜ GEÇEN',
        'YAKLAŞAN (≤90 GÜN / ≤30 GÜN)',
        'DURUM',
        'SON BİLDİRİM TARİHİ'
      ];

      const dataRows = list.map(item => {
        const stats = getBirimStats(item.birim, item.norm);
        let durum = 'GÜVENLİ (SORUN YOK)';
        if (stats.overdueCount > 0) durum = `${stats.overdueCount} KALEM GÜNÜ GEÇTİ`;
        else if (stats.upcomingCount > 0) durum = `${stats.upcomingCount} KALEM YAKLAŞIYOR`;
        else if (stats.totalCount === 0) durum = 'KAYITLI MALZEME YOK (SORUN YOK)';

        return [
          item.birim,
          item.adSoyad,
          item.eposta,
          item.norm || getDefaultNormForBirim(item.birim),
          stats.totalCount,
          stats.overdueCount,
          stats.upcomingCount,
          durum,
          item.mail90 || '-'
        ];
      });

      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.aoa_to_sheet([headers, ...dataRows]);
      ws['!cols'] = [
        { wch: 30 },
        { wch: 22 },
        { wch: 32 },
        { wch: 50 },
        { wch: 16 },
        { wch: 16 },
        { wch: 20 },
        { wch: 26 },
        { wch: 24 }
      ];

      XLSX.utils.book_append_sheet(wb, ws, 'GÜN TAKİP');
      const filename = `GUN_TAKIP_DURUMU_${new Date().toISOString().slice(0, 10)}.xlsx`;
      XLSX.writeFile(wb, filename);
      showNotification(`✅ "${filename}" Excel dosyası olarak indirildi!`, 'success');
    } catch (e: any) {
      showNotification('Excel indirme hatası: ' + e.message, 'error');
    }
  };

  // Seçili birimin bildirimlerini Excel'e aktar
  const handleExportSingleBirimExcel = (birimName: string, items: ApproachingItem[]) => {
    try {
      const headers = [
        'SIRA NO',
        'TEÇHİZAT / MALZEME / PERSONEL ADI',
        'PARÇA NO / SİCİL NO',
        'SERİ NO / ÜNVAN',
        'BÖLÜM / NORM',
        'BULUNDUĞU YER',
        'SON KONTROL / BAKIM / EĞİTİM',
        'GELECEK KONTROL / YENİLEME TARİHİ',
        'KALAN GÜN',
        'DURUM'
      ];

      const dataRows = items.map((item, idx) => [
        idx + 1,
        item.techName,
        item.partNo,
        item.seriNo,
        item.subSection,
        item.location,
        item.sonTarih || '-',
        item.gelecekTarih,
        item.daysDiff <= 0 ? `${Math.abs(item.daysDiff)} Gün Geçti` : `${item.daysDiff} Gün Kaldı`,
        item.daysDiff <= 0 ? 'GÜNÜ GEÇTİ (KRİTİK)' : 'YAKLAŞIYOR'
      ]);

      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.aoa_to_sheet([headers, ...dataRows]);
      ws['!cols'] = [
        { wch: 8 },
        { wch: 38 },
        { wch: 22 },
        { wch: 20 },
        { wch: 24 },
        { wch: 22 },
        { wch: 20 },
        { wch: 24 },
        { wch: 18 },
        { wch: 24 }
      ];

      XLSX.utils.book_append_sheet(wb, ws, 'BİLDİRİM LİSTESİ');
      const safeName = birimName.replace(/[^a-zA-Z0-9]/g, '_');
      const filename = `BILDIRIM_${safeName}_${new Date().toISOString().slice(0, 10)}.xlsx`;
      XLSX.writeFile(wb, filename);
      showNotification(`✅ "${filename}" Excel dosyası olarak indirildi!`, 'success');
    } catch (e: any) {
      showNotification('Excel indirme hatası: ' + e.message, 'error');
    }
  };

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-[800] flex items-center justify-center p-2 sm:p-4 bg-slate-900/85 backdrop-blur-md animate-fade-in">
      <div className="bg-white rounded-3xl w-full max-w-7xl h-[94vh] shadow-2xl border border-slate-200 flex flex-col overflow-hidden">
        
        {/* HEADER */}
        <div className="px-6 py-4 bg-[#0b3d1d] text-white flex items-center justify-between gap-4 shrink-0 shadow-md">
          <div className="flex items-center gap-3.5">
            <div className="w-11 h-11 rounded-2xl bg-emerald-700/60 border border-emerald-500/40 flex items-center justify-center shadow-inner">
              <Bell className="w-5 h-5 text-emerald-300 animate-pulse" />
            </div>
            <div>
              <h2 className="text-base sm:text-lg font-black tracking-tight text-white flex items-center gap-2.5">
                <span>GÜN TAKİP VE OTOMATİK BİLDİRİM PANELİ</span>
                <span className="text-[10px] bg-emerald-500/30 text-emerald-200 font-mono font-black px-2.5 py-0.5 rounded-full border border-emerald-400/40">
                  {isLiveFetching ? '⏳ CANLI VERİLER ALINIYOR...' : '🟢 GERÇEK VERİ & NORM KONTROLÜ'}
                </span>
              </h2>
              <p className="text-xs text-emerald-200/80 font-medium mt-0.5">
                Pencere açılır açılmaz veriler Google Drive ve sunucudan canlı çekilir; seçilen norma göre anında analiz edilip uymayanlar listelenir.
              </p>
            </div>
          </div>

          <div className="flex items-center gap-2">
            <button
              onClick={() => fetchAllLiveDriveData()}
              disabled={isLiveFetching}
              className="px-3 py-1.5 rounded-xl bg-emerald-800 hover:bg-emerald-700 text-white font-bold text-xs flex items-center gap-1.5 transition-all cursor-pointer border border-emerald-600/50 active:scale-95 disabled:opacity-50"
              title="Canlı Drive Verilerini Yeniden Tara"
            >
              <RefreshCw className={`w-3.5 h-3.5 ${isLiveFetching ? 'animate-spin' : ''}`} />
              <span>{isLiveFetching ? 'Taranıyor...' : 'Yeniden Tara'}</span>
            </button>
            <button
              onClick={onClose}
              className="p-2 rounded-xl bg-emerald-900/50 hover:bg-rose-600/80 text-emerald-200 hover:text-white transition-all cursor-pointer border border-emerald-700/50"
              title="Kapat"
            >
              <X className="w-5 h-5" />
            </button>
          </div>
        </div>

        {/* STATUS OVERVIEW CARDS */}
        <div className="px-6 py-3 bg-slate-900 text-white grid grid-cols-1 sm:grid-cols-3 gap-3 border-b border-slate-800 shrink-0">
          {/* Card 1: Günü Geçenler */}
          <div className="bg-rose-950/40 border border-rose-800/50 rounded-2xl p-3 flex items-center gap-3.5">
            <div className="w-10 h-10 rounded-xl bg-rose-600/20 border border-rose-500/40 flex items-center justify-center shrink-0">
              <ShieldAlert className="w-5 h-5 text-rose-400" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <span className="text-xl font-black text-rose-200 font-mono">
                  {isLiveFetching ? '...' : allOverdueItems.length}
                </span>
                <span className="text-[10px] bg-rose-700 text-white font-bold px-1.5 py-0.5 rounded-md uppercase">Günü Geçti</span>
              </div>
              <p className="text-[11px] text-rose-300/80 leading-tight mt-0.5">
                Kayıtlı ve günü geçen teçhizatlar/personel e-postaya otomatik bildirilir.
              </p>
            </div>
          </div>

          {/* Card 2: Yaklaşanlar */}
          <div className="bg-amber-950/40 border border-amber-800/50 rounded-2xl p-3 flex items-center gap-3.5">
            <div className="w-10 h-10 rounded-xl bg-amber-600/20 border border-amber-500/40 flex items-center justify-center shrink-0">
              <Clock className="w-5 h-5 text-amber-400" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <span className="text-xl font-black text-amber-200 font-mono">
                  {isLiveFetching ? '...' : allApproachingItems.length}
                </span>
                <span className="text-[10px] bg-amber-700 text-white font-bold px-1.5 py-0.5 rounded-md uppercase">Yaklaşan Normlar</span>
              </div>
              <p className="text-[11px] text-amber-300/80 leading-tight mt-0.5">
                Depolarda 3 ay, personelde 1 ay kala uyarı otomatik listelenir.
              </p>
            </div>
          </div>

          {/* Card 3: Otomatik Servis Durumu */}
          <div className="bg-emerald-950/40 border border-emerald-800/50 rounded-2xl p-3 flex items-center justify-between gap-3">
            <div className="flex items-center gap-3.5">
              <div className="w-10 h-10 rounded-xl bg-emerald-600/20 border border-emerald-500/40 flex items-center justify-center shrink-0">
                <ShieldCheck className="w-5 h-5 text-emerald-400" />
              </div>
              <div>
                <div className="flex items-center gap-2">
                  <span className="w-2 h-2 rounded-full bg-emerald-400 animate-ping"></span>
                  <span className="text-xs font-black text-emerald-300 uppercase tracking-wide">HER GÜN SABAHLARI 09:00 OTOMATİK TARAMA</span>
                </div>
                <p className="text-[11px] text-emerald-200/80 leading-tight mt-0.5">
                  Her gün 09:00'da tüm parametreler Drive'dan çekilir ve onay sormadan bildirilir.
                </p>
              </div>
            </div>
            <button
              onClick={() => runSilentOverdueCheck(true)}
              disabled={isAutoChecking}
              className="p-2 rounded-xl bg-emerald-800/70 hover:bg-emerald-700 text-emerald-200 hover:text-white transition-all border border-emerald-600/50 cursor-pointer shrink-0 disabled:opacity-50"
              title="Şimdi Arka Plan Kontrolünü Tetikle"
            >
              <RefreshCw className={`w-4 h-4 ${isAutoChecking ? 'animate-spin' : ''}`} />
            </button>
          </div>
        </div>

        {/* TOOLBAR & CONTROLS */}
        <div className="px-6 py-3 bg-slate-50 border-b border-slate-200 flex flex-wrap items-center justify-between gap-3 shrink-0">
          
          {/* Fleet Filter Tabs */}
          <div className="flex items-center gap-1.5 overflow-x-auto pb-1 sm:pb-0 scrollbar-none">
            {[
              { key: 'all', label: 'TÜMÜ' },
              { key: 'AT-802', label: 'AT-802' },
              { key: 'BELL 429', label: 'BELL 429' },
              { key: 'T-70', label: 'T-70' },
              { key: 'B-360', label: 'B-360' },
              { key: 'C-650', label: 'C-650' },
              { key: 'HANGAR', label: 'HANGAR' },
              { key: 'KARA', label: 'KARA ARAÇLARI' },
              { key: 'PERSONEL', label: 'PERSONEL SİSTEMLERİ' }
            ].map(tab => (
              <button
                key={tab.key}
                onClick={() => setActiveFleetFilter(tab.key)}
                className={`px-3 py-1.5 rounded-xl font-black text-xs transition-all cursor-pointer whitespace-nowrap ${
                  activeFleetFilter === tab.key
                    ? 'bg-[#0b3d1d] text-white shadow-sm'
                    : 'bg-white hover:bg-slate-200/80 text-slate-700 border border-slate-200'
                }`}
              >
                {tab.label}
              </button>
            ))}
          </div>

          {/* Action Buttons */}
          <div className="flex items-center gap-2 flex-wrap">
            {/* Search Input */}
            <div className="relative w-44 sm:w-56">
              <Search className="w-3.5 h-3.5 absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
              <input
                type="text"
                placeholder="Birim, Sorumlu, Mail, Norm..."
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                className="w-full pl-8 pr-3 py-1.5 bg-white text-xs text-slate-800 placeholder:text-slate-400 border border-slate-200 rounded-xl focus:outline-none focus:border-[#0b3d1d] font-medium"
              />
            </div>

            {/* Add Sub-Unit Button */}
            <button
              onClick={() => {
                setNewBirimName('');
                setNewAdSoyad('');
                setNewEposta('orhavak.bakimsube@gmail.com');
                setNewNorm(NORM_OPTIONS[0].label);
                setIsAddDialogOpen(true);
              }}
              className="px-3.5 py-1.5 bg-emerald-700 hover:bg-emerald-600 text-white font-bold text-xs rounded-xl transition-all flex items-center gap-1.5 shadow-sm cursor-pointer active:scale-95"
            >
              <Plus className="w-4 h-4 text-emerald-200" />
              <span>BİRİM EKLE</span>
            </button>

            {/* Export Excel Button */}
            <button
              onClick={handleExportExcel}
              className="px-3.5 py-1.5 bg-slate-800 hover:bg-slate-700 text-white font-bold text-xs rounded-xl transition-all flex items-center gap-1.5 shadow-sm cursor-pointer active:scale-95"
              title="GÜN TAKİP Durumunu Excel Formatında İndir"
            >
              <Download className="w-4 h-4 text-slate-300" />
              <span>EXCEL İNDİR</span>
            </button>

            {/* Save All to Sheets */}
            <button
              onClick={() => onSaveSorumlular(list)}
              disabled={isSaving}
              className="px-4 py-1.5 bg-[#0b3d1d] hover:bg-[#072612] text-white font-black text-xs rounded-xl transition-all flex items-center gap-1.5 shadow-sm cursor-pointer active:scale-95 disabled:opacity-50"
            >
              <Save className="w-4 h-4 text-emerald-300" />
              <span>{isSaving ? 'KAYDEDİLİYOR...' : 'KAYDET'}</span>
            </button>
          </div>
        </div>

        {/* DATA TABLE (Çift Tıkla Düzenleme & Listeden Norm Seçimi) */}
        <div className="flex-1 overflow-auto bg-slate-100/50 p-4">
          <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden min-w-[1100px]">
            <table className="w-full border-collapse text-left text-xs">
              <thead>
                <tr className="bg-[#0b3d1d] text-white font-black uppercase text-[11px] tracking-wider border-b border-[#072612]">
                  <th className="py-3 px-3 text-center w-12 border-r border-emerald-900/60">#</th>
                  <th className="py-3 px-4 border-r border-emerald-900/60 w-52">
                    <div className="flex items-center justify-between">
                      <span>SORUMLU BİRİM / ALT BİRİM</span>
                      <span className="text-[9px] font-normal text-emerald-200/70 lowercase font-mono">çift tıkla</span>
                    </div>
                  </th>
                  <th className="py-3 px-4 border-r border-emerald-900/60 w-40">
                    <div className="flex items-center justify-between">
                      <span>ADI SOYADI</span>
                      <span className="text-[9px] font-normal text-emerald-200/70 lowercase font-mono">çift tıkla</span>
                    </div>
                  </th>
                  <th className="py-3 px-4 border-r border-emerald-900/60 w-56">
                    <div className="flex items-center justify-between">
                      <span>E-POSTA ADRESİ</span>
                      <span className="text-[9px] font-normal text-emerald-200/70 lowercase font-mono">çift tıkla</span>
                    </div>
                  </th>
                  <th className="py-3 px-4 border-r border-emerald-900/60 w-80 bg-emerald-950 text-emerald-200">
                    <div className="flex items-center justify-between">
                      <span>TAKİP EDİLECEK NORM</span>
                      <span className="text-[9px] font-bold text-emerald-300 lowercase font-mono">listeden seç</span>
                    </div>
                  </th>
                  <th className="py-3 px-3 text-center border-r border-emerald-900/60 bg-slate-900 text-emerald-300 w-44">
                    BİLDİRİM VE TAKİP DURUMU
                  </th>
                  <th className="py-3 px-3 text-center border-r border-emerald-900/60 w-32">SON BİLDİRİM</th>
                  <th className="py-3 px-3 text-center w-36">İŞLEMLER</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200 font-medium text-slate-700">
                {filteredList.map((item, idx) => {
                  const stats = getBirimStats(item.birim, item.norm);
                  const displayNorm = item.norm || getDefaultNormForBirim(item.birim);

                  const isEditingBirim = editingCell?.index === idx && editingCell.field === 'birim';
                  const isEditingAd = editingCell?.index === idx && editingCell.field === 'adSoyad';
                  const isEditingMail = editingCell?.index === idx && editingCell.field === 'eposta';
                  const isEditingMailDate = editingCell?.index === idx && editingCell.field === 'mail90';

                  return (
                    <tr
                      key={item.birim + idx}
                      className={`hover:bg-slate-50 transition-colors ${
                        idx % 2 === 0 ? 'bg-white' : 'bg-slate-50/50'
                      }`}
                    >
                      {/* Sıra */}
                      <td className="py-2.5 px-3 text-center font-mono text-slate-400 font-bold border-r border-slate-200">
                        {idx + 1}
                      </td>

                      {/* Sorumlu Birim / Alt Birim (Çift Tıkla Düzenlenebilir) */}
                      <td
                        onDoubleClick={() => setEditingCell({ index: idx, field: 'birim' })}
                        className="py-2 px-3 font-black text-slate-900 border-r border-slate-200 cursor-pointer group hover:bg-emerald-50/40 transition-all"
                        title="Çift tıklayarak birim adını düzenleyebilirsiniz"
                      >
                        {isEditingBirim ? (
                          <input
                            type="text"
                            autoFocus
                            value={item.birim}
                            onChange={(e) => handleUpdateItemField(item.birim, 'birim', e.target.value.toUpperCase())}
                            onBlur={handleCellBlurOrEnter}
                            onKeyDown={(e) => e.key === 'Enter' && handleCellBlurOrEnter()}
                            className="w-full px-2 py-1 bg-white border border-emerald-600 rounded-lg text-xs font-mono font-bold text-slate-900 shadow-sm focus:outline-none"
                          />
                        ) : (
                          <div className="flex items-center justify-between gap-1.5">
                            <span className="px-2 py-0.5 rounded-lg bg-emerald-50 text-emerald-900 text-[11px] font-mono font-bold border border-emerald-200">
                              {item.birim}
                            </span>
                            <span className="text-[10px] text-slate-400 font-mono font-semibold">
                              ({isLiveFetching ? '...' : stats.totalCount})
                            </span>
                          </div>
                        )}
                      </td>

                      {/* Adı Soyadı (Çift Tıkla Düzenlenebilir) */}
                      <td
                        onDoubleClick={() => setEditingCell({ index: idx, field: 'adSoyad' })}
                        className="py-2 px-3 border-r border-slate-200 cursor-pointer group hover:bg-emerald-50/40 transition-all"
                        title="Çift tıklayarak adı soyadı düzenleyebilirsiniz"
                      >
                        {isEditingAd ? (
                          <input
                            type="text"
                            autoFocus
                            value={item.adSoyad}
                            onChange={(e) => handleUpdateItemField(item.birim, 'adSoyad', e.target.value)}
                            onBlur={handleCellBlurOrEnter}
                            onKeyDown={(e) => e.key === 'Enter' && handleCellBlurOrEnter()}
                            className="w-full px-2 py-1 bg-white border border-emerald-600 rounded-lg text-xs font-semibold text-slate-900 shadow-sm focus:outline-none"
                          />
                        ) : (
                          <span className="text-xs font-semibold text-slate-800">
                            {item.adSoyad || 'Sorumlu Personel'}
                          </span>
                        )}
                      </td>

                      {/* E-posta Adresi (Çift Tıkla Düzenlenebilir) */}
                      <td
                        onDoubleClick={() => setEditingCell({ index: idx, field: 'eposta' })}
                        className="py-2 px-3 border-r border-slate-200 cursor-pointer group hover:bg-emerald-50/40 transition-all"
                        title="Çift tıklayarak e-posta adresini düzenleyebilirsiniz"
                      >
                        {isEditingMail ? (
                          <input
                            type="email"
                            autoFocus
                            value={item.eposta}
                            onChange={(e) => handleUpdateItemField(item.birim, 'eposta', e.target.value)}
                            onBlur={handleCellBlurOrEnter}
                            onKeyDown={(e) => e.key === 'Enter' && handleCellBlurOrEnter()}
                            className="w-full px-2 py-1 bg-white border border-emerald-600 rounded-lg text-xs font-mono font-medium text-slate-900 shadow-sm focus:outline-none"
                          />
                        ) : (
                          <span className="text-xs font-mono text-slate-700">
                            {item.eposta || 'orhavak.bakimsube@gmail.com'}
                          </span>
                        )}
                      </td>

                      {/* TAKİP EDİLECEK NORM (LİSTEDEN SEÇMELİ DROPDOWN - CANLI KAYDEDER) */}
                      <td className="py-2 px-3 border-r border-slate-200 bg-emerald-50/20">
                        <div className="relative flex items-center">
                          <select
                            value={displayNorm}
                            onChange={(e) => handleSelectNorm(item.birim, e.target.value)}
                            className="w-full py-1.5 pl-2 pr-6 bg-white hover:bg-emerald-50/40 border border-emerald-300 focus:border-emerald-600 rounded-xl text-[11px] font-bold text-slate-800 focus:outline-none transition-all cursor-pointer shadow-2xs appearance-none leading-tight"
                            title="Takip edilecek norm kuralını listeden seçiniz"
                          >
                            {NORM_OPTIONS.map(opt => (
                              <option key={opt.id} value={opt.label}>
                                {opt.label}
                              </option>
                            ))}
                          </select>
                          <ChevronDown className="w-3.5 h-3.5 text-emerald-600 absolute right-2 pointer-events-none" />
                        </div>
                      </td>

                      {/* BİLDİRİM VE TAKİP DURUMU (Gerçek Veri & Norm Analizine Göre Hesaplanır) */}
                      <td className="py-2.5 px-3 text-center border-r border-slate-200">
                        {isLiveFetching ? (
                          <span className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-[10.5px] font-bold bg-emerald-50 text-emerald-800 border border-emerald-200 animate-pulse">
                            <span className="w-2 h-2 rounded-full bg-emerald-500 animate-ping"></span>
                            <span>İnceleniyor...</span>
                          </span>
                        ) : stats.totalCount === 0 ? (
                          <span className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-[10.5px] font-bold bg-slate-100 text-slate-600 border border-slate-300">
                            <Check className="w-3 h-3 text-slate-500" />
                            <span>Kayıt Yok (Sorun Yok)</span>
                          </span>
                        ) : stats.overdueCount > 0 ? (
                          <span className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-[10.5px] font-black bg-rose-100 text-rose-800 border border-rose-300 shadow-xs">
                            <span className="w-2 h-2 rounded-full bg-rose-600 animate-pulse"></span>
                            <span>{stats.overdueCount} Kalem Günü Geçti</span>
                          </span>
                        ) : stats.upcomingCount > 0 ? (
                          <span className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-[10.5px] font-bold bg-amber-50 text-amber-800 border border-amber-300">
                            <Clock className="w-3 h-3 text-amber-600" />
                            <span>{stats.upcomingCount} Kalem Yaklaşıyor</span>
                          </span>
                        ) : (
                          <span className="inline-flex items-center gap-1 text-[10.5px] font-bold text-emerald-700 bg-emerald-50 px-2.5 py-1 rounded-full border border-emerald-200">
                            <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600" />
                            <span>Tümü Güncel (Sorun Yok)</span>
                          </span>
                        )}
                      </td>

                      {/* Son Bildirim Tarihi (Çift Tıkla Düzenlenebilir) */}
                      <td
                        onDoubleClick={() => setEditingCell({ index: idx, field: 'mail90' })}
                        className="py-2.5 px-3 text-center border-r border-slate-200 cursor-pointer group hover:bg-emerald-50/40"
                        title="Çift tıklayarak son bildirim tarihini düzenleyebilirsiniz"
                      >
                        {isEditingMailDate ? (
                          <input
                            type="text"
                            autoFocus
                            value={item.mail90 || ''}
                            onChange={(e) => handleUpdateItemField(item.birim, 'mail90', e.target.value)}
                            onBlur={handleCellBlurOrEnter}
                            onKeyDown={(e) => e.key === 'Enter' && handleCellBlurOrEnter()}
                            className="w-full px-2 py-1 bg-white border border-emerald-600 rounded-lg text-xs font-mono font-medium text-slate-900 shadow-sm focus:outline-none"
                            placeholder="GG.AA.YYYY"
                          />
                        ) : item.mail90 ? (
                          <div className="inline-flex items-center gap-1 px-2 py-0.5 rounded-lg bg-emerald-50 text-emerald-900 border border-emerald-200 text-[10.5px] font-mono font-bold">
                            <span>📅 {item.mail90}</span>
                          </div>
                        ) : (
                          <span className="text-slate-400 text-[10.5px] font-mono italic">Beklemede</span>
                        )}
                      </td>

                      {/* İşlemler (Canlı İnceleme & Sil) */}
                      <td className="py-2.5 px-3 text-center">
                        <div className="flex items-center justify-center gap-1.5">
                          {isLiveFetching ? (
                            <div className="flex items-center justify-center gap-1.5 py-1 px-3 rounded-xl bg-slate-100 text-slate-500 font-bold text-[10.5px] animate-pulse">
                              <RefreshCw className="w-3.5 h-3.5 animate-spin text-emerald-600" />
                              <span>Veri Alınıyor...</span>
                            </div>
                          ) : (
                            <button
                              type="button"
                              onClick={() => {
                                setSelectedBirimDetails({ birim: item, stats });
                                setDetailFilter('all');
                                setDetailSearch('');
                              }}
                              className={`px-2.5 py-1 rounded-xl transition-all cursor-pointer font-bold text-[10.5px] flex items-center gap-1 active:scale-95 shadow-xs ${
                                stats.overdueCount > 0
                                  ? 'bg-rose-700 hover:bg-rose-800 text-white'
                                  : stats.upcomingCount > 0
                                  ? 'bg-amber-600 hover:bg-amber-700 text-white'
                                  : 'bg-slate-800 hover:bg-slate-700 text-white'
                              }`}
                              title="Norm kurallarına uymayan ve yaklaşan kayıtları detaylı incele"
                            >
                              <Eye className="w-3.5 h-3.5 text-emerald-300" />
                              <span>
                                {stats.approachingItems.length > 0
                                  ? `İncele (${stats.approachingItems.length})`
                                  : `İncele (0)`}
                              </span>
                            </button>
                          )}
                          <button
                            type="button"
                            onClick={() => handleDeleteBirim(idx)}
                            className="p-1.5 rounded-xl bg-rose-100 hover:bg-rose-200 text-rose-700 transition-all cursor-pointer active:scale-95"
                            title="Birimi Listeden Sil"
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
        </div>

        {/* FOOTER INFO BAR */}
        <div className="px-6 py-3 bg-slate-50 border-t border-slate-200 flex items-center justify-between text-xs text-slate-500 font-medium shrink-0">
          <div className="flex items-center gap-2">
            <span className="font-black text-slate-800">Toplam {list.length} Birim Takip Ediliyor.</span>
            <span>(Normu listeden anında değiştirebilir, hücrelere çift tıklayarak düzenleyebilirsiniz)</span>
          </div>
          <div className="flex items-center gap-2 text-[11px] text-slate-600">
            <span className="w-2 h-2 rounded-full bg-emerald-500"></span>
            <span>Otomatik Takip: Her gün 09:00'da tüm Drive verileri taranır, günü geçen veya yaklaşımlar E-Tabloya ve maile otomatik yansıtılır.</span>
          </div>
        </div>
      </div>

      {/* ========================================================================= */}
      {/* ŞIK VE GELİŞMİŞ BİLDİRİM AYRINTILARI TABLO MODALI (UYMAYANLARI LİSTELEME) */}
      {/* ========================================================================= */}
      {selectedBirimDetails && (() => {
        const { birim, stats } = selectedBirimDetails;

        const filteredItems = stats.approachingItems.filter(item => {
          if (detailFilter === 'overdue' && item.daysDiff > 0) return false;
          if (detailFilter === 'upcoming' && item.daysDiff <= 0) return false;

          if (detailSearch.trim()) {
            const q = detailSearch.toLowerCase();
            const n = item.techName.toLowerCase();
            const p = item.partNo.toLowerCase();
            const s = item.seriNo.toLowerCase();
            const l = item.location.toLowerCase();
            return n.includes(q) || p.includes(q) || s.includes(q) || l.includes(q);
          }
          return true;
        });

        return (
          <div className="fixed inset-0 z-[860] flex items-center justify-center p-2 sm:p-4 bg-black/80 backdrop-blur-sm animate-fade-in">
            <div className="bg-white rounded-3xl w-full max-w-6xl max-h-[92vh] shadow-2xl border border-slate-200 flex flex-col overflow-hidden">
              
              {/* MODAL HEADER */}
              <div className="px-6 py-4 bg-slate-900 text-white flex items-center justify-between border-b border-slate-800 shrink-0">
                <div className="flex items-center gap-3.5">
                  <div className="w-10 h-10 rounded-2xl bg-emerald-700/50 border border-emerald-500/40 flex items-center justify-center">
                    <Bell className="w-5 h-5 text-emerald-300" />
                  </div>
                  <div>
                    <h3 className="text-base font-black text-white flex items-center gap-2.5 flex-wrap">
                      <span>{birim.birim}</span>
                      <span className="text-[11px] bg-slate-800 text-emerald-300 font-mono px-2.5 py-0.5 rounded-full border border-slate-700 font-bold">
                        Toplam Taranan: {stats.totalCount} Malzeme/Kayıt
                      </span>
                      {stats.approachingItems.length > 0 ? (
                        <span className="text-[11px] bg-rose-950 text-rose-300 font-mono px-2.5 py-0.5 rounded-full border border-rose-800 font-bold">
                          {stats.overdueCount} Günü Geçti &bull; {stats.upcomingCount} Yaklaşan Norm
                        </span>
                      ) : (
                        <span className="text-[11px] bg-emerald-950 text-emerald-300 font-mono px-2.5 py-0.5 rounded-full border border-emerald-800 font-bold">
                          Norma Uygun (Tümü Güncel ve Güvenli)
                        </span>
                      )}
                    </h3>
                    <p className="text-xs text-slate-300 font-medium mt-0.5">
                      Sorumlu: <strong className="text-white">{birim.adSoyad}</strong> &bull; E-Posta: <strong className="text-emerald-300 font-mono">{birim.eposta}</strong>
                    </p>
                  </div>
                </div>

                <div className="flex items-center gap-2">
                  {stats.approachingItems.length > 0 && (
                    <button
                      type="button"
                      onClick={() => handleExportSingleBirimExcel(birim.birim, stats.approachingItems)}
                      className="px-3 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-200 hover:text-white text-xs font-bold rounded-xl border border-slate-700 flex items-center gap-1.5 transition-all cursor-pointer"
                      title="Bu tablodaki uymayan bildirim listesini Excel olarak indir"
                    >
                      <Download className="w-4 h-4 text-emerald-400" />
                      <span className="hidden sm:inline">Excel İndir</span>
                    </button>
                  )}
                  <button
                    onClick={() => setSelectedBirimDetails(null)}
                    className="p-2 rounded-xl text-slate-400 hover:text-white hover:bg-slate-800 transition-all cursor-pointer"
                    title="Kapat"
                  >
                    <X className="w-5 h-5" />
                  </button>
                </div>
              </div>

              {/* TAKİP NORMU BİLGİLENDİRME BANTI */}
              <div className="px-6 py-2.5 bg-emerald-950 text-emerald-200 text-xs flex items-center justify-between border-b border-emerald-900 shrink-0 font-medium">
                <div className="flex items-center gap-2">
                  <Info className="w-4 h-4 text-emerald-400 shrink-0" />
                  <span><strong>Uygulanan Takip Normu:</strong> {birim.norm || getDefaultNormForBirim(birim.birim)}</span>
                </div>
                <span className="text-[10.5px] font-mono text-emerald-300 bg-emerald-900/60 px-2 py-0.5 rounded border border-emerald-800">
                  Canlı Drive & Server Taraması
                </span>
              </div>

              {/* TOOLBAR FOR DETAIL TABLE */}
              {stats.approachingItems.length > 0 && (
                <div className="px-6 py-3 bg-slate-50 border-b border-slate-200 flex flex-wrap items-center justify-between gap-3 shrink-0">
                  <div className="flex items-center gap-2">
                    <button
                      type="button"
                      onClick={() => setDetailFilter('all')}
                      className={`px-3 py-1 rounded-xl text-xs font-black transition-all cursor-pointer ${
                        detailFilter === 'all'
                          ? 'bg-slate-900 text-white shadow-xs'
                          : 'bg-white text-slate-600 hover:bg-slate-200 border border-slate-200'
                      }`}
                    >
                      TÜMÜ ({stats.approachingItems.length})
                    </button>
                    <button
                      type="button"
                      onClick={() => setDetailFilter('overdue')}
                      className={`px-3 py-1 rounded-xl text-xs font-black transition-all cursor-pointer ${
                        detailFilter === 'overdue'
                          ? 'bg-rose-700 text-white shadow-xs'
                          : 'bg-white text-rose-700 hover:bg-rose-50 border border-rose-200'
                      }`}
                    >
                      🔴 GÜNÜ GEÇENLER ({stats.overdueCount})
                    </button>
                    <button
                      type="button"
                      onClick={() => setDetailFilter('upcoming')}
                      className={`px-3 py-1 rounded-xl text-xs font-black transition-all cursor-pointer ${
                        detailFilter === 'upcoming'
                          ? 'bg-amber-600 text-white shadow-xs'
                          : 'bg-white text-amber-700 hover:bg-amber-50 border border-amber-200'
                      }`}
                    >
                      🟡 YAKLAŞANLAR ({stats.upcomingCount})
                    </button>
                  </div>

                  <div className="relative w-60">
                    <Search className="w-3.5 h-3.5 absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
                    <input
                      type="text"
                      placeholder="Malzeme, P/N, S/N veya Yer Ara..."
                      value={detailSearch}
                      onChange={(e) => setDetailSearch(e.target.value)}
                      className="w-full pl-8 pr-3 py-1 bg-white text-xs border border-slate-200 rounded-xl focus:outline-none focus:border-[#0b3d1d]"
                    />
                  </div>
                </div>
              )}

              {/* DETAIL CONTENT AREA */}
              <div className="flex-1 overflow-auto p-6 bg-slate-100/50">
                {stats.approachingItems.length === 0 ? (
                  <div className="py-16 text-center flex flex-col items-center justify-center bg-white border border-slate-200 rounded-2xl shadow-xs">
                    <div className="w-16 h-16 rounded-3xl bg-emerald-50 text-emerald-600 flex items-center justify-center mb-3 text-2xl font-black">
                      <CheckCircle2 className="w-8 h-8 text-emerald-600" />
                    </div>
                    <h4 className="text-sm font-black text-slate-800 uppercase tracking-wide">
                      TÜM KAYITLAR NORM KURALLARINA UYGUNDUR
                    </h4>
                    <p className="text-xs text-slate-500 max-w-md mt-1 font-medium">
                      Bu birimde kayıtlı {stats.totalCount} malzeme/teçhizat/personel taranmış olup, norm kurallarına uymayan (günü geçen veya yaklaşan) hiçbir kayıt bulunmamaktadır. Tüm kayıtlar güncel ve güvenlidir.
                    </p>
                  </div>
                ) : (
                  <div className="bg-white rounded-2xl border border-slate-200 shadow-xs overflow-hidden">
                    <table className="w-full border-collapse text-left text-xs">
                      <thead>
                        <tr className="bg-slate-900 text-white font-black uppercase text-[10.5px] tracking-wider border-b border-slate-800">
                          <th className="py-3 px-3 text-center w-12 border-r border-slate-800">#</th>
                          <th className="py-3 px-4 border-r border-slate-800">TEÇHİZAT / MALZEME / PERSONEL ADI</th>
                          <th className="py-3 px-3 border-r border-slate-800">P/N VEYA SİCİL NO</th>
                          <th className="py-3 px-3 border-r border-slate-800">SERİ NO / ÜNVAN</th>
                          <th className="py-3 px-3 border-r border-slate-800">NORM KAPSAMI</th>
                          <th className="py-3 px-3 border-r border-slate-800">BULUNDUĞU YER</th>
                          <th className="py-3 px-3 text-center border-r border-slate-800">SON İŞLEM</th>
                          <th className="py-3 px-3 text-center border-r border-slate-800">GELECEK TARİH</th>
                          <th className="py-3 px-3 text-center border-r border-slate-800">KALAN GÜN</th>
                          <th className="py-3 px-3 text-center">BİLDİRİM DURUMU</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-200 font-medium text-slate-700">
                        {filteredItems.map((item, idx) => {
                          const isOverdue = item.daysDiff <= 0;

                          return (
                            <tr
                              key={item.techName + idx}
                              className={`hover:bg-slate-50 transition-colors ${
                                isOverdue ? 'bg-rose-50/40' : idx % 2 === 0 ? 'bg-white' : 'bg-slate-50/50'
                              }`}
                            >
                              <td className="py-2.5 px-3 text-center font-mono text-slate-400 font-bold border-r border-slate-200">
                                {idx + 1}
                              </td>

                              <td className="py-2.5 px-4 font-bold text-slate-900 border-r border-slate-200">
                                {item.techName}
                              </td>

                              <td className="py-2.5 px-3 font-mono text-slate-700 border-r border-slate-200">
                                {item.partNo}
                              </td>

                              <td className="py-2.5 px-3 font-mono text-slate-700 border-r border-slate-200">
                                {item.seriNo}
                              </td>

                              <td className="py-2.5 px-3 border-r border-slate-200">
                                <span className="px-2 py-0.5 rounded-md bg-slate-100 text-slate-800 text-[10.5px] font-mono font-bold border border-slate-200">
                                  {item.subSection}
                                </span>
                              </td>

                              <td className="py-2.5 px-3 text-slate-700 border-r border-slate-200">
                                {item.location}
                              </td>

                              <td className="py-2.5 px-3 text-center font-mono text-slate-600 border-r border-slate-200">
                                {item.sonTarih || '-'}
                              </td>

                              <td className="py-2.5 px-3 text-center font-mono font-black text-slate-900 border-r border-slate-200">
                                {item.gelecekTarih}
                              </td>

                              <td className="py-2.5 px-3 text-center font-mono font-black border-r border-slate-200">
                                {isOverdue ? (
                                  <span className="text-rose-700 font-black">
                                    {Math.abs(item.daysDiff)} GÜN GEÇTİ
                                  </span>
                                ) : (
                                  <span className="text-amber-700 font-bold">
                                    {item.daysDiff} GÜN KALDI
                                  </span>
                                )}
                              </td>

                              <td className="py-2.5 px-3 text-center">
                                {isOverdue ? (
                                  <span className="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-[10.5px] font-black bg-rose-700 text-white shadow-xs">
                                    🔴 OTOMATİK BİLDİRİMDE
                                  </span>
                                ) : (
                                  <span className="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-[10.5px] font-bold bg-amber-100 text-amber-900 border border-amber-300">
                                    🟡 TAKİPTE (YAKLAŞAN)
                                  </span>
                                )}
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>

              <div className="px-6 py-3 bg-slate-50 border-t border-slate-200 flex items-center justify-between text-xs text-slate-500 font-medium shrink-0">
                <span>Görüntülenen {filteredItems.length} Kalem Uymayan Kayıt.</span>
                <button
                  type="button"
                  onClick={() => setSelectedBirimDetails(null)}
                  className="px-4 py-1.5 bg-slate-900 hover:bg-slate-800 text-white font-bold rounded-xl transition-all cursor-pointer"
                >
                  Kapat
                </button>
              </div>
            </div>
          </div>
        );
      })()}

      {/* ========================================================================= */}
      {/* YENİ BİRİM / ALT BİRİM EKLEME DİYALOĞU */}
      {/* ========================================================================= */}
      {isAddDialogOpen && (
        <div className="fixed inset-0 z-[880] flex items-center justify-center p-4 bg-black/75 backdrop-blur-xs animate-fade-in">
          <div className="bg-white rounded-3xl w-full max-w-lg shadow-2xl border border-slate-200 overflow-hidden font-sans">
            <div className="px-6 py-4 bg-[#0b3d1d] text-white flex items-center justify-between">
              <h3 className="text-sm font-black flex items-center gap-2">
                <Plus className="w-4 h-4 text-emerald-300" />
                <span>YENİ BİRİM / ALT BİRİM EKLE</span>
              </h3>
              <button
                onClick={() => setIsAddDialogOpen(false)}
                className="text-emerald-200 hover:text-white"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="p-6 space-y-4">
              <div>
                <label className="block text-xs font-black text-slate-700 uppercase mb-1">
                  BİRİM / ALT BİRİM ADI:
                </label>
                <input
                  type="text"
                  value={newBirimName}
                  onChange={(e) => {
                    const name = e.target.value;
                    setNewBirimName(name);
                    setNewNorm(getDefaultNormForBirim(name));
                  }}
                  placeholder="Örn: AT-802 YAŞAM DESTEK, PERSONEL BİLGİ SİSTEMİ..."
                  className="w-full px-3 py-2 border border-slate-300 rounded-xl text-xs font-bold focus:outline-none focus:border-[#0b3d1d]"
                />
              </div>

              <div>
                <label className="block text-xs font-black text-slate-700 uppercase mb-1">
                  SORUMLU ADI SOYADI:
                </label>
                <input
                  type="text"
                  value={newAdSoyad}
                  onChange={(e) => setNewAdSoyad(e.target.value)}
                  placeholder="Örn: Sorumlu Teknisyen"
                  className="w-full px-3 py-2 border border-slate-300 rounded-xl text-xs font-semibold focus:outline-none focus:border-[#0b3d1d]"
                />
              </div>

              <div>
                <label className="block text-xs font-black text-slate-700 uppercase mb-1">
                  E-POSTA ADRESİ:
                </label>
                <input
                  type="email"
                  value={newEposta}
                  onChange={(e) => setNewEposta(e.target.value)}
                  placeholder="orhavak.bakimsube@gmail.com"
                  className="w-full px-3 py-2 border border-slate-300 rounded-xl text-xs font-mono focus:outline-none focus:border-[#0b3d1d]"
                />
              </div>

              <div>
                <label className="block text-xs font-black text-slate-700 uppercase mb-1">
                  TAKİP EDİLECEK NORM:
                </label>
                <div className="relative">
                  <select
                    value={newNorm}
                    onChange={(e) => setNewNorm(e.target.value)}
                    className="w-full py-2 pl-3 pr-8 border border-slate-300 rounded-xl text-xs font-bold focus:outline-none focus:border-[#0b3d1d] bg-white appearance-none"
                  >
                    {NORM_OPTIONS.map(opt => (
                      <option key={opt.id} value={opt.label}>
                        {opt.label}
                      </option>
                    ))}
                  </select>
                  <ChevronDown className="w-4 h-4 text-slate-500 absolute right-2.5 top-1/2 -translate-y-1/2 pointer-events-none" />
                </div>
              </div>
            </div>

            <div className="px-6 py-3.5 bg-slate-50 border-t border-slate-200 flex items-center justify-end gap-2">
              <button
                type="button"
                onClick={() => setIsAddDialogOpen(false)}
                className="px-4 py-2 rounded-xl border border-slate-300 text-xs font-bold text-slate-700 hover:bg-slate-100"
              >
                İptal
              </button>
              <button
                type="button"
                onClick={handleAddNewBirim}
                className="px-5 py-2 rounded-xl bg-[#0b3d1d] hover:bg-[#072612] text-white text-xs font-black uppercase tracking-wider cursor-pointer"
              >
                BİRİMİ KAYDET
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
