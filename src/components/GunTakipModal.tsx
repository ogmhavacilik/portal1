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
  Check
} from 'lucide-react';
import * as XLSX from 'xlsx';

export interface GunTakipBirimRow {
  birim: string;
  adSoyad: string;
  eposta: string;
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
  };
  onUpdateEquipmentMailDate?: (unitKey: string, rowIndex: number, mailColIndex: number, dateStr: string) => void;
  googleScriptUrl: string;
}

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

  // 1. Tag in explicit section tag column
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

  // 2. Search for explicit markers embedded in cells
  const combined = normalizeTurkishStr(row.join(" ")).toUpperCase();
  if (combined.includes("##DEPO_SARF##") || combined.includes("[SARF DEPO]") || combined.includes("[PARCA DEPO]")) return 'depo_sarf';
  if (combined.includes("##DEPO_KIMYASAL##") || combined.includes("[KIMYASAL DEPO]")) return 'depo_kimyasal';
  if (combined.includes("##OZEL_ALET##") || combined.includes("[OZEL ALET]") || combined.includes("OZEL BAKIM VE TEST ALETLERI") || combined.includes("OZEL BAKIM ALETLERI") || combined.includes("OZEL ALETLER")) return 'ozel_alet';
  if (combined.includes("##YER_DESTEK##") || combined.includes("[YER DESTEK]") || combined.includes("YER DESTEK TECHIZATLARI")) return 'yer_destek';

  // 3. Heuristics based on item name (row[1]) or part number (row[2]) or description (row[11])
  const itemName = normalizeTurkishStr(String(row[1] || "") + " " + String(row[2] || "") + " " + String(row[11] || "")).toUpperCase();

  // Sarf depo kontrolü
  if (itemName.includes("FILTRE") || itemName.includes("FILTER") || itemName.includes("CONTA") || itemName.includes("O-RING") || itemName.includes("BALATA") || itemName.includes("KECE") || itemName.includes("SEAL") || itemName.includes("SARF") || itemName.includes("CIVATA") || itemName.includes("SOMUN") || itemName.includes("PUL") || itemName.includes("RIVET") || itemName.includes("PERCIN")) {
    return 'depo_sarf';
  }

  // Özel Alet kontrolü
  if (itemName.includes("OZEL ALET") || itemName.includes("TEST ADAPTOR") || itemName.includes("TORK ADAPTOR") || itemName.includes("KOMPRESOR YIKAMA") || itemName.includes("AYAR MASTAR") || itemName.includes("BOSALTMA KAPAK TEST") || itemName.includes("CALIBRATION") || itemName.includes("TEST RIG") || itemName.includes("FIREGATE") || itemName.includes("OZEL BAKIM") || itemName.includes("ACHIOLCER") || itemName.includes("ACIOLCER") || itemName.includes("KALIBRASYON KIT") || itemName.includes("HAT BAKIM VE KALIBRASYON") || itemName.includes("TORQUE") || itemName.includes("MASTAR") || itemName.includes("ADAPTOR") || itemName.includes("TEST KIT") || itemName.includes("TORK ANAHTAR") || itemName.includes("MOTOR TEST") || itemName.includes("TORQ") || itemName.includes("WRENCH") || itemName.includes("PULLER") || itemName.includes("SOCKET") || itemName.includes("VOLTMETRE") || itemName.includes("MULTIMETRE") || itemName.includes("MANOMETRE")) {
    return 'ozel_alet';
  }

  return 'yer_destek';
};

// Tarih ayrıştırma ve fark hesaplama
export const calculateDaysDiff = (dateStr: string): number | null => {
  if (!dateStr || typeof dateStr !== 'string') return null;
  const clean = dateStr.trim();
  if (!clean || clean === '-' || clean === '--' || clean.toLowerCase() === 'yok' || clean.toLowerCase().includes('belirtilme')) return null;

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

  // New Sub-Unit Dialog
  const [isAddDialogOpen, setIsAddDialogOpen] = useState<boolean>(false);
  const [newBirimName, setNewBirimName] = useState<string>('');
  const [newAdSoyad, setNewAdSoyad] = useState<string>('');
  const [newEposta, setNewEposta] = useState<string>('orhavak.bakimsube@gmail.com');

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

  // Format date helper
  const getNowFormattedWithSeconds = (): string => {
    const now = new Date();
    const pad = (n: number) => String(n).padStart(2, '0');
    return `${pad(now.getDate())}.${pad(now.getMonth() + 1)}.${now.getFullYear()} ${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}`;
  };

  /**
   * GERÇEK VERİ EŞLEME:
   * Sorumlu birim adına göre ilgili veri kümesini filtreler ve hem toplam kayıt sayısını hem de yaklaşan/günü geçen öğeleri tespit eder.
   */
  const getBirimStats = (birimName: string): BirimStats => {
    const norm = normalizeTurkishStr(birimName);

    // 1. Hedef Filo Tespiti
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

    // 2. Hedef Bölüm Tespiti (Depo mu, Yer Destek mi, Özel Alet mi?)
    let targetSection: 'all' | 'depo' | 'ozel_alet' | 'yer_destek' = 'all';
    if (norm.includes('DEPO')) {
      targetSection = 'depo';
    } else if (norm.includes('OZEL') || norm.includes('ÖZEL')) {
      targetSection = 'ozel_alet';
    } else if (norm.includes('YER DESTEK') || norm.includes('YERDESTEK')) {
      targetSection = 'yer_destek';
    }

    // İlgili veri dizisini al
    const rawRows = (unitDataMap as any)[targetFleetKey] || [];
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

      // Bölüm Filtresi Eşleşmesi
      if (targetSection === 'depo') {
        // Yalnızca DEPO satırları (Yer destek veya özel aletler ASLA depoya giremez!)
        const loc = String(row[5] || '').toLowerCase();
        if (!isDepoRow && !loc.includes('depo')) return;
      } else if (targetSection === 'ozel_alet') {
        if (sec !== 'ozel_alet') return;
      } else if (targetSection === 'yer_destek') {
        if (sec !== 'yer_destek') return;
      }

      // Bu satır bu birime aittir -> Toplam sayacı artır
      totalCount++;

      // Alanları ve tarihleri güvenli şekilde oku
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
        if (omurlu !== 'EVET') isMuaf = true; // Ömürlü parça değilse gün takibi yapılmaz
        gelecekTarih = String(row[8] || '').trim();
        sonTarih = String(row[9] || '').trim();
        lastMailSentDate = String(row[11] || '').trim();
      } else if (targetFleetKey === 'bell429') {
        // Bell 429 özel indeksleri
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
        // Standart Uçak ve Hangar tabloları (T-70, AT-802, B-360, C-650, Hangar)
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

      // Eğer bakımdan muafsa veya gelecek tarih boşsa bildirim oluşturulmaz
      if (isMuaf || !gelecekTarih) return;

      const daysDiff = calculateDaysDiff(gelecekTarih);
      // Sadece günü gelen veya 90 günden az kalanlar bildirime girer
      if (daysDiff === null || daysDiff > 90) return;

      const isNew = !lastMailSentDate || lastMailSentDate.trim() === '' || lastMailSentDate.trim() === '-';

      approachingItems.push({
        unitKey: targetFleetKey,
        unitName: targetFleetKey.toUpperCase(),
        subSection: isDepoRow ? 'DEPO' : (sec === 'ozel_alet' ? 'ÖZEL ALET' : 'YER DESTEK'),
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
        return b.includes(q) || a.includes(q) || e.includes(q);
      }
      return true;
    });
  }, [list, activeFleetFilter, searchQuery]);

  // Tüm filo genelindeki toplam günü geçenler ve yaklaşanlar
  const allOverdueItems = useMemo(() => {
    const map = new Map<string, ApproachingItem>();
    list.forEach(b => {
      const stats = getBirimStats(b.birim);
      stats.approachingItems.filter(i => i.daysDiff <= 0).forEach(i => {
        const key = `${i.unitKey}_${i.partNo}_${i.seriNo}_${i.gelecekTarih}`;
        map.set(key, i);
      });
    });
    return Array.from(map.values());
  }, [list, unitDataMap]);

  const allApproachingItems = useMemo(() => {
    const map = new Map<string, ApproachingItem>();
    list.forEach(b => {
      const stats = getBirimStats(b.birim);
      stats.approachingItems.filter(i => i.daysDiff > 0 && i.daysDiff <= 90).forEach(i => {
        const key = `${i.unitKey}_${i.partNo}_${i.seriNo}_${i.gelecekTarih}`;
        map.set(key, i);
      });
    });
    return Array.from(map.values());
  }, [list, unitDataMap]);

  // Arka plan otomatik kontrol
  const runSilentOverdueCheck = async (isManualTrigger = false) => {
    setIsAutoChecking(true);
    const nowTimestamp = getNowFormattedWithSeconds();

    try {
      // 1. Local backend API
      await fetch('/api/trigger-auto-reminders', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          birimList: list,
          overdueItems: allOverdueItems,
          approachingItems: allApproachingItems
        })
      }).catch(e => console.warn('Local trigger reminder warn:', e));

      // 2. Google Apps Script
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

  // Yeni Birim Ekleme
  const handleAddNewBirim = () => {
    if (!newBirimName.trim()) {
      showNotification('Lütfen geçerli bir birim veya alt birim adı giriniz.', 'warning');
      return;
    }
    const cleanName = newBirimName.trim().toUpperCase();
    if (list.some(item => item.birim.toUpperCase() === cleanName)) {
      showNotification('Bu birim veya alt birim listede zaten mevcut.', 'warning');
      return;
    }

    const newRow: GunTakipBirimRow = {
      birim: cleanName,
      adSoyad: newAdSoyad.trim() || 'Sorumlu Personel',
      eposta: newEposta.trim() || 'orhavak.bakimsube@gmail.com',
      mail90: ''
    };

    const updated = [...list, newRow];
    setList(updated);
    setNewBirimName('');
    setNewAdSoyad('');
    setIsAddDialogOpen(false);
    showNotification(`✅ "${cleanName}" başarıyla listeye eklendi!`, 'success');
  };

  // Birim Silme
  const handleDeleteBirim = (index: number) => {
    const itemToDelete = filteredList[index];
    if (!itemToDelete) return;
    if (confirm(`"${itemToDelete.birim}" birimini silmek istediğinize emin misiniz?`)) {
      const updated = list.filter(item => item.birim !== itemToDelete.birim);
      setList(updated);
      showNotification(`"${itemToDelete.birim}" listeden silindi.`, 'info');
    }
  };

  // Tüm sorumluları Excel'e aktar
  const handleExportExcel = () => {
    try {
      const headers = [
        'SORUMLU BİRİM / ALT BİRİM',
        'ADI SOYADI',
        'E-POSTA ADRESİ',
        'TOPLAM KAYITLI',
        'GÜNÜ GEÇEN',
        'YAKLAŞAN (≤90 GÜN)',
        'DURUM',
        'SON BİLDİRİM TARİHİ'
      ];

      const dataRows = list.map(item => {
        const stats = getBirimStats(item.birim);
        let durum = 'GÜVENLİ (SORUN YOK)';
        if (stats.overdueCount > 0) durum = `${stats.overdueCount} KALEM GÜNÜ GEÇTİ`;
        else if (stats.upcomingCount > 0) durum = `${stats.upcomingCount} KALEM YAKLAŞIYOR`;
        else if (stats.totalCount === 0) durum = 'KAYITLI MALZEME YOK (SORUN YOK)';

        return [
          item.birim,
          item.adSoyad,
          item.eposta,
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
        'TEÇHİZAT / MALZEME ADI',
        'PARÇA NO (P/N)',
        'SERİ NO (S/N)',
        'BÖLÜM',
        'BULUNDUĞU YER',
        'SON KONTROL / BAKIM',
        'GELECEK KONTROL / BAKIM',
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
        item.daysDiff <= 0 ? 'GÜNÜ GEÇTİ (KRİTİK)' : 'YAKLAŞIYOR (≤90 GÜN)'
      ]);

      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.aoa_to_sheet([headers, ...dataRows]);
      ws['!cols'] = [
        { wch: 8 },
        { wch: 35 },
        { wch: 22 },
        { wch: 20 },
        { wch: 16 },
        { wch: 22 },
        { wch: 20 },
        { wch: 22 },
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
                  GERÇEK VERİ KONTROLÜ
                </span>
              </h2>
              <p className="text-xs text-emerald-200/80 font-medium mt-0.5">
                Sistemdeki gerçek envanter kayıtları taranır; kayıt yoksa "SORUN YOK" kabul edilir, ≤90 gün kalan teçhizatlar otomatik takip edilir.
              </p>
            </div>
          </div>

          <div className="flex items-center gap-2">
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
                <span className="text-xl font-black text-rose-200 font-mono">{allOverdueItems.length}</span>
                <span className="text-[10px] bg-rose-700 text-white font-bold px-1.5 py-0.5 rounded-md uppercase">Günü Geçti</span>
              </div>
              <p className="text-[11px] text-rose-300/80 leading-tight mt-0.5">
                Kayıtlı ve günü geçen teçhizatlar arka planda onay sormadan bildirilir.
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
                <span className="text-xl font-black text-amber-200 font-mono">{allApproachingItems.length}</span>
                <span className="text-[10px] bg-amber-700 text-white font-bold px-1.5 py-0.5 rounded-md uppercase">Yaklaşan (≤ 90 Gün)</span>
              </div>
              <p className="text-[11px] text-amber-300/80 leading-tight mt-0.5">
                Sistem takibinde. Günü geldiğinde e-posta otomatik postalanır.
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
                  <span className="text-xs font-black text-emerald-300 uppercase tracking-wide">OTOMATİK SERVİS AKTİF</span>
                </div>
                <p className="text-[11px] text-emerald-200/80 leading-tight mt-0.5">
                  Her gün 09:00 ve günü geçenlerde arka planda otomatik çalışır.
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
              { key: 'KARA', label: 'KARA ARAÇLARI' }
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
                placeholder="Birim, Sorumlu, Mail..."
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                className="w-full pl-8 pr-3 py-1.5 bg-white text-xs text-slate-800 placeholder:text-slate-400 border border-slate-200 rounded-xl focus:outline-none focus:border-[#0b3d1d] font-medium"
              />
            </div>

            {/* Add Sub-Unit Button */}
            <button
              onClick={() => setIsAddDialogOpen(true)}
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

        {/* DATA TABLE */}
        <div className="flex-1 overflow-auto bg-slate-100/50 p-4">
          <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden min-w-[1000px]">
            <table className="w-full border-collapse text-left text-xs">
              <thead>
                <tr className="bg-[#0b3d1d] text-white font-black uppercase text-[11px] tracking-wider border-b border-[#072612]">
                  <th className="py-3 px-3 text-center w-12 border-r border-emerald-900/60">#</th>
                  <th className="py-3 px-4 border-r border-emerald-900/60">SORUMLU BİRİM / ALT BİRİM</th>
                  <th className="py-3 px-4 border-r border-emerald-900/60">ADI SOYADI</th>
                  <th className="py-3 px-4 border-r border-emerald-900/60">E-POSTA ADRESİ</th>
                  <th className="py-3 px-3 text-center border-r border-emerald-900/60 bg-slate-900 text-emerald-300">
                    BİLDİRİM VE TAKİP DURUMU
                  </th>
                  <th className="py-3 px-3 text-center border-r border-emerald-900/60">SON BİLDİRİM</th>
                  <th className="py-3 px-3 text-center w-40">İŞLEMLER</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200 font-medium text-slate-700">
                {filteredList.map((item, idx) => {
                  const stats = getBirimStats(item.birim);

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

                      {/* Birim / Alt Birim */}
                      <td className="py-2.5 px-4 font-black text-slate-900 border-r border-slate-200">
                        <div className="flex items-center gap-2">
                          <span className="px-2.5 py-1 rounded-lg bg-emerald-50 text-emerald-900 text-[11px] font-mono font-bold border border-emerald-200">
                            {item.birim}
                          </span>
                          <span className="text-[10px] text-slate-400 font-mono font-semibold">
                            ({stats.totalCount} Ürün)
                          </span>
                        </div>
                      </td>

                      {/* Adı Soyadı */}
                      <td className="py-2.5 px-4 border-r border-slate-200">
                        <input
                          type="text"
                          value={item.adSoyad}
                          onChange={(e) => {
                            const val = e.target.value;
                            setList(prev => prev.map(p => p.birim === item.birim ? { ...p, adSoyad: val } : p));
                          }}
                          className="w-full px-2.5 py-1 bg-transparent hover:bg-slate-100 focus:bg-white border border-transparent focus:border-[#0b3d1d] rounded-lg text-xs font-semibold text-slate-800 transition-all"
                          placeholder="Ad Soyad giriniz..."
                        />
                      </td>

                      {/* E-posta */}
                      <td className="py-2.5 px-4 border-r border-slate-200">
                        <input
                          type="email"
                          value={item.eposta}
                          onChange={(e) => {
                            const val = e.target.value;
                            setList(prev => prev.map(p => p.birim === item.birim ? { ...p, eposta: val } : p));
                          }}
                          className="w-full px-2.5 py-1 bg-transparent hover:bg-slate-100 focus:bg-white border border-transparent focus:border-[#0b3d1d] rounded-lg text-xs font-mono font-medium text-slate-800 transition-all"
                          placeholder="orhavak@ogm.gov.tr"
                        />
                      </td>

                      {/* Bildirim Durumu: GERÇEK VERİYE GÖRE HESAPLANIR */}
                      <td className="py-2.5 px-3 text-center border-r border-slate-200">
                        {stats.totalCount === 0 ? (
                          <span className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full text-[11px] font-bold bg-slate-100 text-slate-600 border border-slate-300">
                            <Check className="w-3 h-3 text-slate-500" />
                            <span>Kayıt Yok (Sorun Yok)</span>
                          </span>
                        ) : stats.overdueCount > 0 ? (
                          <span className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full text-[11px] font-black bg-rose-100 text-rose-800 border border-rose-300 shadow-xs">
                            <span className="w-2 h-2 rounded-full bg-rose-600 animate-pulse"></span>
                            <span>{stats.overdueCount} Kalem Günü Geçti</span>
                          </span>
                        ) : stats.upcomingCount > 0 ? (
                          <span className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full text-[11px] font-bold bg-amber-50 text-amber-800 border border-amber-300">
                            <Clock className="w-3 h-3 text-amber-600" />
                            <span>{stats.upcomingCount} Kalem Yaklaşıyor</span>
                          </span>
                        ) : (
                          <span className="inline-flex items-center gap-1 text-[11px] font-bold text-emerald-700 bg-emerald-50 px-3 py-1 rounded-full border border-emerald-200">
                            <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600" />
                            <span>Tümü Güncel (Sorun Yok)</span>
                          </span>
                        )}
                      </td>

                      {/* Son Bildirim Tarihi */}
                      <td className="py-2.5 px-3 text-center border-r border-slate-200">
                        {item.mail90 ? (
                          <div className="inline-flex items-center gap-1 px-2.5 py-1 rounded-lg bg-emerald-50 text-emerald-900 border border-emerald-200 text-[11px] font-mono font-bold">
                            <span>📅 {item.mail90}</span>
                          </div>
                        ) : (
                          <span className="text-slate-400 text-[11px] font-mono italic">Beklemede</span>
                        )}
                      </td>

                      {/* İşlemler */}
                      <td className="py-2.5 px-3 text-center">
                        <div className="flex items-center justify-center gap-1.5">
                          <button
                            type="button"
                            onClick={() => {
                              setSelectedBirimDetails({ birim: item, stats });
                              setDetailFilter('all');
                              setDetailSearch('');
                            }}
                            className={`px-3 py-1.5 rounded-xl transition-all cursor-pointer font-bold text-[11px] flex items-center gap-1.5 active:scale-95 shadow-xs ${
                              stats.overdueCount > 0
                                ? 'bg-rose-700 hover:bg-rose-800 text-white'
                                : stats.upcomingCount > 0
                                ? 'bg-amber-600 hover:bg-amber-700 text-white'
                                : 'bg-slate-800 hover:bg-slate-700 text-white'
                            }`}
                            title="Bu birimin teçhizat ve bildirim detaylarını incele"
                          >
                            <Eye className="w-3.5 h-3.5 text-emerald-300" />
                            <span>
                              {stats.totalCount === 0
                                ? 'İncele (0)'
                                : stats.approachingItems.length > 0
                                ? `Bildirim (${stats.approachingItems.length})`
                                : `İncele (${stats.totalCount})`}
                            </span>
                          </button>
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
            <span className="font-black text-slate-800">Toplam {list.length} Birim Kayıtlı.</span>
            <span>(Filtrelenen: {filteredList.length})</span>
          </div>
          <div className="flex items-center gap-2 text-[11px] text-slate-600">
            <span className="w-2 h-2 rounded-full bg-emerald-500"></span>
            <span>Otomatik Takip: Bakım/kontrol günü geçen teçhizatlar her sabah 09:00'da onay sormadan arka planda bildirilir.</span>
          </div>
        </div>
      </div>

      {/* ========================================================================= */}
      {/* ŞIK VE GELİŞMİŞ BİLDİRİM AYRINTILARI TABLO MODALI (ANA PANEL GÖRÜNÜMÜNDE) */}
      {/* ========================================================================= */}
      {selectedBirimDetails && (() => {
        const { birim, stats } = selectedBirimDetails;

        // Detay içi filtreleme
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
                        Toplam Kayıtlı: {stats.totalCount} Malzeme
                      </span>
                      {stats.approachingItems.length > 0 ? (
                        <span className="text-[11px] bg-rose-950 text-rose-300 font-mono px-2.5 py-0.5 rounded-full border border-rose-800 font-bold">
                          {stats.overdueCount} Günü Geçti &bull; {stats.upcomingCount} Yaklaşan
                        </span>
                      ) : (
                        <span className="text-[11px] bg-emerald-950 text-emerald-300 font-mono px-2.5 py-0.5 rounded-full border border-emerald-800 font-bold">
                          Sorun Yok (Güvenli)
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
                      title="Bu tablodaki bildirim listesini Excel olarak indir"
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

              {/* TOOLBAR FOR DETAIL TABLE */}
              {stats.approachingItems.length > 0 && (
                <div className="px-6 py-3 bg-slate-50 border-b border-slate-200 flex flex-wrap items-center justify-between gap-3 shrink-0">
                  {/* Filter Tabs */}
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

                  {/* Search Input */}
                  <div className="relative w-56">
                    <Search className="w-3.5 h-3.5 absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
                    <input
                      type="text"
                      placeholder="Tabloda Ara (Malzeme, P/N, S/N)..."
                      value={detailSearch}
                      onChange={(e) => setDetailSearch(e.target.value)}
                      className="w-full pl-8 pr-3 py-1.5 bg-white text-xs text-slate-800 placeholder:text-slate-400 border border-slate-200 rounded-xl focus:outline-none focus:border-[#0b3d1d] font-medium"
                    />
                  </div>
                </div>
              )}

              {/* MODAL BODY */}
              <div className="flex-1 overflow-auto p-4 bg-slate-50">
                {stats.totalCount === 0 ? (
                  /* DURUM 1: Bu birimde kayıtlı 0 ürün var (Örn: T-70 DEPO) */
                  <div className="py-16 px-6 text-center flex flex-col items-center justify-center gap-4 bg-white rounded-2xl border border-slate-200 shadow-xs">
                    <div className="w-16 h-16 rounded-3xl bg-emerald-50 border-2 border-emerald-200 flex items-center justify-center shadow-inner">
                      <ShieldCheck className="w-8 h-8 text-emerald-600" />
                    </div>
                    <div>
                      <h4 className="text-base font-black text-slate-800 uppercase tracking-wider">
                        KAYITLI MALZEME BULUNMUYOR &bull; SORUN YOK
                      </h4>
                      <p className="text-xs text-slate-500 font-medium max-w-lg mt-1.5 leading-relaxed">
                        <strong>"{birim.birim}"</strong> birimine ait sistemde henüz kayıtlı bir envanter / teçhizat bulunmamaktadır. Dolayısıyla herhangi bir bakım gecikmesi veya bildirim riski yoktur (Güvenilirlik Tamdır).
                      </p>
                    </div>
                    <div className="inline-flex items-center gap-2 px-4 py-2 rounded-xl bg-emerald-100/60 text-emerald-900 border border-emerald-300 text-xs font-bold">
                      <CheckCircle2 className="w-4 h-4 text-emerald-700" />
                      <span>Sistem Durumu: GÜVENLİ (SORUN YOK)</span>
                    </div>
                  </div>
                ) : stats.approachingItems.length === 0 ? (
                  /* DURUM 2: Kayıtlı ürünler var ama hiçbirinin günü geçmemiş ve <=90 gün kalmamış */
                  <div className="py-16 px-6 text-center flex flex-col items-center justify-center gap-4 bg-white rounded-2xl border border-slate-200 shadow-xs">
                    <div className="w-16 h-16 rounded-3xl bg-emerald-50 border-2 border-emerald-200 flex items-center justify-center shadow-inner">
                      <CheckCircle2 className="w-8 h-8 text-emerald-600" />
                    </div>
                    <div>
                      <h4 className="text-base font-black text-slate-800 uppercase tracking-wider">
                        TÜM BAKIM VE KONTROLLER GÜNCEL &bull; SORUN YOK
                      </h4>
                      <p className="text-xs text-slate-500 font-medium max-w-lg mt-1.5 leading-relaxed">
                        <strong>"{birim.birim}"</strong> birimindeki kayıtlı <strong>{stats.totalCount}</strong> malzemenin hiçbirinde günü geçen veya 90 günden az kalan bakım bulunmamaktadır. Tüm kalibrasyon ve kontroller güvenli periyottadır.
                      </p>
                    </div>
                    <div className="inline-flex items-center gap-2 px-4 py-2 rounded-xl bg-emerald-100/60 text-emerald-900 border border-emerald-300 text-xs font-bold">
                      <CheckCircle2 className="w-4 h-4 text-emerald-700" />
                      <span>{stats.totalCount} Malzeme Güvenli Aralıkta (Sorun Yok)</span>
                    </div>
                  </div>
                ) : filteredItems.length === 0 ? (
                  /* DURUM 3: Filtreye uyan kayıt yok */
                  <div className="py-12 text-center text-slate-400 font-bold text-xs">
                    Arama kriterlerine uygun bildirim kaydı bulunamadı.
                  </div>
                ) : (
                  /* DURUM 4: ANA PANEL GÖRÜNÜMÜNDE ŞIK, NET, OKUNAKLI TEÇHİZAT BİLDİRİM TABLOSU */
                  <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
                    <table className="w-full text-left text-xs border-collapse">
                      <thead>
                        <tr className="bg-slate-900 text-slate-200 font-black text-[10px] uppercase font-mono tracking-wider border-b border-slate-800">
                          <th className="py-3 px-3 text-center w-12 border-r border-slate-800">#</th>
                          <th className="py-3 px-4 border-r border-slate-800">TEÇHİZAT / MALZEME ADI</th>
                          <th className="py-3 px-3 border-r border-slate-800">P/N & S/N</th>
                          <th className="py-3 px-3 border-r border-slate-800 text-center">BÖLÜM</th>
                          <th className="py-3 px-3 border-r border-slate-800">BULUNDUĞU YER</th>
                          <th className="py-3 px-3 border-r border-slate-800 text-center">SON KONTROL</th>
                          <th className="py-3 px-3 border-r border-slate-800 text-center">GELECEK KONTROL</th>
                          <th className="py-3 px-3 text-center border-r border-slate-800">KALAN / GEÇEN SÜRE</th>
                          <th className="py-3 px-3 text-center w-28">DURUM</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100">
                        {filteredItems.map((item, i) => {
                          const isOverdue = item.daysDiff <= 0;
                          return (
                            <tr
                              key={i}
                              className={`transition-colors ${
                                isOverdue ? 'bg-rose-50/40 hover:bg-rose-100/50' : 'hover:bg-amber-50/40 odd:bg-white even:bg-slate-50/40'
                              }`}
                            >
                              {/* Sıra */}
                              <td className="py-2.5 px-3 text-center font-mono text-slate-400 font-bold border-r border-slate-100">
                                {i + 1}
                              </td>

                              {/* Malzeme Adı */}
                              <td className="py-2.5 px-4 font-bold text-slate-900 border-r border-slate-100">
                                <div className="leading-tight">
                                  <span className="text-slate-900 font-black">{item.techName}</span>
                                  {item.miktar && item.miktar !== '1' && (
                                    <span className="ml-1.5 text-[10px] text-slate-500 font-mono">
                                      (Miktar: {item.miktar})
                                    </span>
                                  )}
                                </div>
                              </td>

                              {/* P/N & S/N */}
                              <td className="py-2.5 px-3 font-mono text-[11px] text-slate-700 border-r border-slate-100 whitespace-nowrap">
                                <div><strong className="text-slate-500 text-[10px]">P/N:</strong> {item.partNo}</div>
                                <div><strong className="text-slate-500 text-[10px]">S/N:</strong> {item.seriNo}</div>
                              </td>

                              {/* Bölüm */}
                              <td className="py-2.5 px-3 text-center border-r border-slate-100">
                                <span className="px-2 py-0.5 rounded-md bg-slate-100 text-slate-700 font-mono text-[10px] font-bold border border-slate-200">
                                  {item.subSection}
                                </span>
                              </td>

                              {/* Bulunduğu Yer */}
                              <td className="py-2.5 px-3 text-slate-700 font-medium border-r border-slate-100 text-[11px]">
                                {item.location}
                              </td>

                              {/* Son Kontrol */}
                              <td className="py-2.5 px-3 text-center font-mono text-slate-600 text-[11px] border-r border-slate-100">
                                {item.sonTarih || '-'}
                              </td>

                              {/* Gelecek Kontrol */}
                              <td className="py-2.5 px-3 text-center font-mono font-bold text-slate-900 border-r border-slate-100">
                                {item.gelecekTarih}
                              </td>

                              {/* Kalan / Geçen Süre */}
                              <td className="py-2.5 px-3 text-center border-r border-slate-100">
                                {isOverdue ? (
                                  <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-full bg-rose-600 text-white text-[10px] font-black shadow-xs">
                                    <span className="w-1.5 h-1.5 rounded-full bg-white animate-ping"></span>
                                    <span>{Math.abs(item.daysDiff)} GÜN GEÇTİ</span>
                                  </span>
                                ) : (
                                  <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-full bg-amber-500 text-white text-[10px] font-black shadow-xs">
                                    <Clock className="w-3 h-3 text-white" />
                                    <span>{item.daysDiff} GÜN KALDI</span>
                                  </span>
                                )}
                              </td>

                              {/* Durumu */}
                              <td className="py-2.5 px-3 text-center">
                                <span className="inline-flex items-center px-2 py-0.5 rounded-md bg-emerald-50 text-emerald-800 text-[10px] font-bold border border-emerald-200">
                                  {item.status || 'Faal'}
                                </span>
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>

              {/* MODAL FOOTER */}
              <div className="px-6 py-3 bg-slate-100 border-t border-slate-200 flex items-center justify-between shrink-0">
                <div className="text-xs text-slate-600 font-medium">
                  {stats.totalCount > 0 ? (
                    <span>
                      Bu birimde kayıtlı <strong>{stats.totalCount}</strong> teçhizattan <strong>{stats.approachingItems.length}</strong> tanesi bildirim kapsamındadır.
                    </span>
                  ) : (
                    <span>Bu birimde henüz kayıtlı teçhizat yoktur. Durum güvenlidir.</span>
                  )}
                </div>
                <button
                  type="button"
                  onClick={() => setSelectedBirimDetails(null)}
                  className="px-5 py-2 bg-slate-800 hover:bg-slate-700 text-white font-bold text-xs rounded-xl cursor-pointer transition-all active:scale-95"
                >
                  Kapat
                </button>
              </div>
            </div>
          </div>
        );
      })()}

      {/* MODAL: ADD NEW SUB-UNIT */}
      {isAddDialogOpen && (
        <div className="fixed inset-0 z-[850] flex items-center justify-center p-4 bg-black/60 backdrop-blur-xs animate-fade-in">
          <div className="bg-white rounded-3xl p-6 max-w-md w-full shadow-2xl border border-slate-200 flex flex-col gap-4">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <h3 className="text-sm font-black text-slate-900 uppercase tracking-wider flex items-center gap-2">
                <Plus className="w-4 h-4 text-emerald-700" />
                <span>YENİ BİRİM / ALT BİRİM EKLE</span>
              </h3>
              <button
                onClick={() => setIsAddDialogOpen(false)}
                className="text-slate-400 hover:text-slate-600 font-bold cursor-pointer"
              >
                ✕
              </button>
            </div>

            <div className="flex flex-col gap-3">
              <div>
                <label className="text-[11px] font-bold text-slate-600 block mb-1">
                  Birim / Alt Birim Adı (Örn: AT-802 DEPO, T-70 ÖZEL ALET):
                </label>
                <input
                  type="text"
                  value={newBirimName}
                  onChange={(e) => setNewBirimName(e.target.value)}
                  placeholder="Örn: AT-802 ÖZEL ALET"
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-800 focus:outline-none focus:border-[#0b3d1d]"
                />
              </div>

              <div>
                <label className="text-[11px] font-bold text-slate-600 block mb-1">Sorumlu Adı Soyadı:</label>
                <input
                  type="text"
                  value={newAdSoyad}
                  onChange={(e) => setNewAdSoyad(e.target.value)}
                  placeholder="Örn: Bakım Sorumlusu"
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-800 focus:outline-none focus:border-[#0b3d1d]"
                />
              </div>

              <div>
                <label className="text-[11px] font-bold text-slate-600 block mb-1">E-Posta Adresi:</label>
                <input
                  type="email"
                  value={newEposta}
                  onChange={(e) => setNewEposta(e.target.value)}
                  placeholder="orhavak.bakimsube@gmail.com"
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-mono text-slate-800 focus:outline-none focus:border-[#0b3d1d]"
                />
              </div>
            </div>

            <div className="flex items-center justify-end gap-2 pt-2 border-t border-slate-100">
              <button
                onClick={() => setIsAddDialogOpen(false)}
                className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold rounded-xl cursor-pointer"
              >
                Vazgeç
              </button>
              <button
                onClick={handleAddNewBirim}
                className="px-4 py-2 bg-[#0b3d1d] hover:bg-[#072612] text-white text-xs font-black rounded-xl cursor-pointer"
              >
                Ekle
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
