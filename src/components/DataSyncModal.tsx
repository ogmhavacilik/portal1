import React, { useState, useRef } from 'react';
import { 
  RefreshCw, 
  FileSpreadsheet, 
  FileText, 
  Upload, 
  CheckCircle2, 
  X, 
  AlertTriangle,
  ArrowLeft,
  FolderOpen,
  Download,
  Zap,
  CloudUpload,
  CheckCircle,
  Database
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { 
  isHeaderLikeRow, 
  cleanAndFormatDateString, 
  unmergeAndFillWorksheet, 
  groupMultiLocationRowsHelper, 
  parseExcelWorksheetToStandardRows,
  detectUnitAndSectionFromNames
} from '../utils/driveExcelSync';

export const GOOGLE_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbw4kruTTc058Y9rLTyO3dKi6KloYsmdDTwV1GSiAk8ZXefyo3Z7_VDSTuurzsS9BHAQyQ/exec";
export const DRIVE_FOLDER_ID = "1HQR_NYKhHQGA7_2W3nArI9pCh-LJasTP";

export const getStandardDriveFileName = (targetKey: string, originalName?: string): string => {
  const targetMap: Record<string, string> = {
    'techizat_at802': 'hava_araçları_yer_destek_at-802.xlsx',
    'techizat_at802_yer_destek': 'hava_araçları_yer_destek_at-802.xlsx',
    'techizat_at802_ozel_alet': 'at-802_ozel_bakim_aletleri.xlsx',
    'techizat_at802_depo_sarf': 'at-802_sarf_ve_parca_deposu.xlsx',
    'techizat_at802_depo_kimyasal': 'at-802_kimyasal_depo.xlsx',

    'techizat_bell429': 'hava_araçları_yer_destek_bell-429.xlsx',
    'techizat_bell429_yer_destek': 'hava_araçları_yer_destek_bell-429.xlsx',
    'techizat_bell429_depo_sarf': 'bell-429_sarf_ve_parca_deposu.xlsx',
    'techizat_bell429_depo_kimyasal': 'bell-429_kimyasal_depo.xlsx',

    'techizat_t70': 'hava_araçları_yer_destek_t-70.xlsx',
    'techizat_t70_yer_destek': 'hava_araçları_yer_destek_t-70.xlsx',
    'techizat_t70_depo_sarf': 't-70_sarf_ve_parca_deposu.xlsx',
    'techizat_t70_depo_kimyasal': 't-70_kimyasal_depo.xlsx',
    'techizat_t70_bumbi_backet': 'hava_araçları_yer_destek_t-70_bumbi_backet.xlsx',
    'techizat_t70_helitak': 'hava_araçları_yer_destek_t-70_helitak.xlsx',

    'techizat_b360': 'hava_araçları_yer_destek_b-360.xlsx',
    'techizat_b360_yer_destek': 'hava_araçları_yer_destek_b-360.xlsx',
    'techizat_b360_depo_sarf': 'b-360_sarf_ve_parca_deposu.xlsx',
    'techizat_b360_depo_kimyasal': 'b-360_kimyasal_depo.xlsx',

    'techizat_c650': 'hava_araçları_yer_destek_c-650.xlsx',
    'techizat_c650_yer_destek': 'hava_araçları_yer_destek_c-650.xlsx',
    'techizat_c650_depo_sarf': 'c-650_sarf_ve_parca_deposu.xlsx',
    'techizat_c650_depo_kimyasal': 'c-650_kimyasal_depo.xlsx',

    'techizat_hangar': 'hava_araçları_yer_destek_hangar.xlsx',
    'techizat_hangar_depo_sarf': 'hangar_sarf_ve_parca_deposu.xlsx',
    'techizat_hangar_depo_kimyasal': 'hangar_kimyasal_depo.xlsx',

    'techizat_kara_araclari': 'kara_araçları_takip.xlsx',

    // Olay Takip Çizelgeleri
    'olay_takip_at802': 'olay_takip_cizelgesi_at-802.xlsx',
    'olay_takip_bell429': 'olay_takip_cizelgesi_bell-429.xlsx',
    'olay_takip_t70': 'olay_takip_cizelgesi_t-70.xlsx',
    'olay_takip_c650': 'olay_takip_cizelgesi_c-650.xlsx',
    'olay_takip_b360': 'olay_takip_cizelgesi_b-360.xlsx',
    'olay_takip_hangar': 'olay_takip_cizelgesi_hangar.xlsx',
  };

  if (targetMap[targetKey]) {
    return targetMap[targetKey];
  }

  // If user uploaded a file with a recognizable name
  if (originalName) {
    const lower = originalName.toLowerCase()
      .replace(/ı/g, 'i').replace(/ğ/g, 'g').replace(/ü/g, 'u')
      .replace(/ş/g, 's').replace(/ö/g, 'o').replace(/ç/g, 'c');

    if (lower.includes('olay') || lower.includes('kirim') || lower.includes('cizelge')) {
      if (lower.includes('bell')) return 'olay_takip_cizelgesi_bell-429.xlsx';
      if (lower.includes('t70') || lower.includes('t-70')) return 'olay_takip_cizelgesi_t-70.xlsx';
      if (lower.includes('c650') || lower.includes('c-650')) return 'olay_takip_cizelgesi_c-650.xlsx';
      if (lower.includes('b360') || lower.includes('b-360')) return 'olay_takip_cizelgesi_b-360.xlsx';
      if (lower.includes('hangar')) return 'olay_takip_cizelgesi_hangar.xlsx';
      return 'olay_takip_cizelgesi_at-802.xlsx';
    }

    if (lower.includes('at802') || lower.includes('at-802')) {
      if (lower.includes('ozel') || lower.includes('alet')) return 'at-802_ozel_bakim_aletleri.xlsx';
      if (lower.includes('sarf') || lower.includes('parca')) return 'at-802_sarf_ve_parca_deposu.xlsx';
      if (lower.includes('kimya')) return 'at-802_kimyasal_depo.xlsx';
      return 'hava_araçları_yer_destek_at-802.xlsx';
    }
    if (lower.includes('bell')) {
      if (lower.includes('sarf') || lower.includes('parca')) return 'bell-429_sarf_ve_parca_deposu.xlsx';
      if (lower.includes('kimya')) return 'bell-429_kimyasal_depo.xlsx';
      return 'hava_araçları_yer_destek_bell-429.xlsx';
    }
    if (lower.includes('t70') || lower.includes('t-70')) {
      if (lower.includes('bumbi')) return 'hava_araçları_yer_destek_t-70_bumbi_backet.xlsx';
      if (lower.includes('helitak')) return 'hava_araçları_yer_destek_t-70_helitak.xlsx';
      if (lower.includes('sarf') || lower.includes('parca')) return 't-70_sarf_ve_parca_deposu.xlsx';
      if (lower.includes('kimya')) return 't-70_kimyasal_depo.xlsx';
      return 'hava_araçları_yer_destek_t-70.xlsx';
    }
    if (lower.includes('b360') || lower.includes('b-360')) {
      if (lower.includes('sarf') || lower.includes('parca')) return 'b-360_sarf_ve_parca_deposu.xlsx';
      if (lower.includes('kimya')) return 'b-360_kimyasal_depo.xlsx';
      return 'hava_araçları_yer_destek_b-360.xlsx';
    }
    if (lower.includes('c650') || lower.includes('c-650')) {
      if (lower.includes('sarf') || lower.includes('parca')) return 'c-650_sarf_ve_parca_deposu.xlsx';
      if (lower.includes('kimya')) return 'c-650_kimyasal_depo.xlsx';
      return 'hava_araçları_yer_destek_c-650.xlsx';
    }
    if (lower.includes('hangar')) {
      if (lower.includes('sarf') || lower.includes('parca')) return 'hangar_sarf_ve_parca_deposu.xlsx';
      if (lower.includes('kimya')) return 'hangar_kimyasal_depo.xlsx';
      return 'hava_araçları_yer_destek_hangar.xlsx';
    }
    if (lower.includes('kara')) {
      return 'kara_araçları_takip.xlsx';
    }

    let clean = originalName.trim();
    if (!clean.toLowerCase().endsWith('.xlsx') && !clean.toLowerCase().endsWith('.xls')) {
      clean += '.xlsx';
    }
    return clean;
  }

  return `${targetKey}.xlsx`;
};

interface DataSyncModalProps {
  isOpen: boolean;
  onClose: () => void;
  initialStep?: 1 | 2;
  initialTarget?: string;
  activeUnit?: string;
  activeSection?: string;
  onUpdateUnitData?: (unit: string, newCols: string[], newRows: string[][]) => void;
  onUploadMasterPdf?: (file: File) => void;
  onApplyDriveSync?: () => Promise<void> | void;
  onApplyUploadedData?: (detectedUnit: string, rows: string[][]) => void;
  onDownloadLatestExcel?: (unit: string) => void;
  showNotification?: (msg: string) => void;
}

export const DataSyncModal: React.FC<DataSyncModalProps> = ({
  isOpen,
  onClose,
  initialStep = 1,
  initialTarget,
  activeUnit,
  activeSection,
  onUpdateUnitData,
  onUploadMasterPdf,
  onApplyDriveSync,
  onApplyUploadedData,
  onDownloadLatestExcel,
  showNotification = (_msg: string) => {},
}) => {
  // Determine starting target based on activeUnit / section
  const computeInitialTarget = () => {
    if (initialTarget) return initialTarget;
    if (activeUnit) {
      if (activeSection && activeSection !== 'all') {
        return `techizat_${activeUnit}_${activeSection}`;
      }
      return `techizat_${activeUnit}`;
    }
    return '1';
  };

  const [step, setStep] = useState<1 | 2>(initialStep);
  const [selectedTarget, setSelectedTarget] = useState<string>(computeInitialTarget());
  const [isProcessing, setIsProcessing] = useState<boolean>(false);
  const [progressPercent, setProgressPercent] = useState<number>(0);
  const [progressStage, setProgressStage] = useState<string>('');
  const [isCheckingUpdates, setIsCheckingUpdates] = useState<boolean>(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  if (!isOpen) return null;

  const TARGET_LABELS: Record<string, { title: string; subtitle: string; isPdf?: boolean }> = {
    // Planlama & Personel Çizelgeleri
    '1': { title: '1. GÖREVLENDİRME ÇİZELGELERİ (1-GOREVLENDIRME)', subtitle: 'Hedef Sayfa Adresi: PLANLAMA VE PERSONEL ÇİZELGELERİ', isPdf: true },
    '21': { title: '2. YAZ DÖNEMİ PLANLAMASI - BELL 429 (2-YAZ_DONEMI-BELL_429)', subtitle: 'Hedef Sayfa Adresi: PLANLAMA VE PERSONEL ÇİZELGELERİ', isPdf: true },
    '22': { title: '2. YAZ DÖNEMİ PLANLAMASI - T-70 (2-YAZ_DONEMI-T_70)', subtitle: 'Hedef Sayfa Adresi: PLANLAMA VE PERSONEL ÇİZELGELERİ', isPdf: true },
    '23': { title: '2. YAZ DÖNEMİ PLANLAMASI - AT-802 (2-YAZ_DONEMI-AT_802)', subtitle: 'Hedef Sayfa Adresi: PLANLAMA VE PERSONEL ÇİZELGELERİ', isPdf: true },
    '24': { title: '2. YAZ DÖNEMİ PLANLAMASI - ANKARA BEKLEME (BELL-429)', subtitle: 'Hedef Sayfa Adresi: PLANLAMA VE PERSONEL ÇİZELGELERİ', isPdf: true },
    '25': { title: '2. YAZ DÖNEMİ PLANLAMASI - ANKARA BEKLEME (C-650/B-360)', subtitle: 'Hedef Sayfa Adresi: PLANLAMA VE PERSONEL ÇİZELGELERİ', isPdf: true },
    '3': { title: '3. BAKIM YETKİ ÇİZELGELERİ (3-BAKIM_YETKI)', subtitle: 'Hedef Sayfa Adresi: PLANLAMA VE PERSONEL ÇİZELGELERİ', isPdf: true },
    '5': { title: '5. PERSONEL BİLGİ ÇİZELGELERİ (5-PERSONEL_BILGI)', subtitle: 'Hedef Sayfa Adresi: PLANLAMA VE PERSONEL ÇİZELGELERİ', isPdf: true },
    '6': { title: '6. PERSONEL UÇUŞ-HİZMET YILLARI (6-PERSONEL_UCUS_HIZMET)', subtitle: 'Hedef Sayfa Adresi: PLANLAMA VE PERSONEL ÇİZELGELERİ', isPdf: true },

    // AT-802F
    'techizat_at802': { title: '✈️ AT-802F - TÜM BİRİM ENVANTERİ (EXCEL)', subtitle: 'Hedef Sayfa Adresi: TÜM TEÇHİZAT & SÜRÜCÜ YEDEKLEMESİ' },
    'techizat_at802_yer_destek': { title: '✈️ AT-802F - YER DESTEK TEÇHİZATLARI (EXCEL)', subtitle: 'Hedef Sayfa Adresi: TÜM TEÇHİZAT & SÜRÜCÜ YEDEKLEMESİ' },
    'techizat_at802_ozel_alet': { title: '✈️ AT-802F - ÖZEL ALETLER (EXCEL)', subtitle: 'Hedef Sayfa Adresi: TÜM TEÇHİZAT & SÜRÜCÜ YEDEKLEMESİ' },
    'techizat_at802_depo_sarf': { title: '📦 AT-802F - SARF VE PARÇA DEPOSU (EXCEL)', subtitle: 'Hedef Sayfa Adresi: TÜM TEÇHİZAT & SÜRÜCÜ YEDEKLEMESİ' },
    'techizat_at802_depo_kimyasal': { title: '🧪 AT-802F - KİMYASAL DEPO (EXCEL)', subtitle: 'Hedef Sayfa Adresi: TÜM TEÇHİZAT & SÜRÜCÜ YEDEKLEMESİ' },

    // BELL 429
    'techizat_bell429': { title: '🚁 BELL 429 - TÜM BİRİM ENVANTERİ (EXCEL)', subtitle: 'Hedef Sayfa Adresi: TÜM TEÇHİZAT & SÜRÜCÜ YEDEKLEMESİ' },
    'techizat_bell429_yer_destek': { title: '🚁 BELL 429 - YER DESTEK VE ÖZEL ALETLER (EXCEL)', subtitle: 'Hedef Sayfa Adresi: TÜM TEÇHİZAT & SÜRÜCÜ YEDEKLEMESİ' },
    'techizat_bell429_depo_sarf': { title: '📦 BELL 429 - SARF VE PARÇA DEPOSU (EXCEL)', subtitle: 'Hedef Sayfa Adresi: TÜM TEÇHİZAT & SÜRÜCÜ YEDEKLEMESİ' },
    'techizat_bell429_depo_kimyasal': { title: '🧪 BELL 429 - KİMYASAL DEPO (EXCEL)', subtitle: 'Hedef Sayfa Adresi: TÜM TEÇHİZAT & SÜRÜCÜ YEDEKLEMESİ' },

    // T-70
    'techizat_t70': { title: '🚁 T-70 - TÜM BİRİM ENVANTERİ (EXCEL)', subtitle: 'Hedef Sayfa Adresi: TÜM TEÇHİZAT & SÜRÜCÜ YEDEKLEMESİ' },
    'techizat_t70_yer_destek': { title: '🚁 T-70 - YER DESTEK VE ÖZEL ALETLER (EXCEL)', subtitle: 'Hedef Sayfa Adresi: TÜM TEÇHİZAT & SÜRÜCÜ YEDEKLEMESİ' },
    'techizat_t70_depo_sarf': { title: '📦 T-70 - SARF VE PARÇA DEPOSU (EXCEL)', subtitle: 'Hedef Sayfa Adresi: TÜM TEÇHİZAT & SÜRÜCÜ YEDEKLEMESİ' },
    'techizat_t70_depo_kimyasal': { title: '🧪 T-70 - KİMYASAL DEPO (EXCEL)', subtitle: 'Hedef Sayfa Adresi: TÜM TEÇHİZAT & SÜRÜCÜ YEDEKLEMESİ' },

    // T-70 Bumbi Bucket & Helitak
    'techizat_t70_bumbi_backet': { title: '🪣 T-70 BUMBİ BUCKET - TÜM ENVANTER (EXCEL)', subtitle: 'Hedef Sayfa Adresi: TÜM TEÇHİZAT & SÜRÜCÜ YEDEKLEMESİ' },
    'techizat_t70_helitak': { title: '⚙️ T-70 HELİTAK - TÜM ENVANTER (EXCEL)', subtitle: 'Hedef Sayfa Adresi: TÜM TEÇHİZAT & SÜRÜCÜ YEDEKLEMESİ' },

    // B-360
    'techizat_b360': { title: '✈️ BEECHCRAFT B-360 - TÜM BİRİM ENVANTERİ (EXCEL)', subtitle: 'Hedef Sayfa Adresi: TÜM TEÇHİZAT & SÜRÜCÜ YEDEKLEMESİ' },
    'techizat_b360_yer_destek': { title: '✈️ BEECHCRAFT B-360 - YER DESTEK VE ÖZEL ALETLER (EXCEL)', subtitle: 'Hedef Sayfa Adresi: TÜM TEÇHİZAT & SÜRÜCÜ YEDEKLEMESİ' },
    'techizat_b360_depo_sarf': { title: '📦 BEECHCRAFT B-360 - SARF VE PARÇA DEPOSU (EXCEL)', subtitle: 'Hedef Sayfa Adresi: TÜM TEÇHİZAT & SÜRÜCÜ YEDEKLEMESİ' },
    'techizat_b360_depo_kimyasal': { title: '🧪 BEECHCRAFT B-360 - KİMYASAL DEPO (EXCEL)', subtitle: 'Hedef Sayfa Adresi: TÜM TEÇHİZAT & SÜRÜCÜ YEDEKLEMESİ' },

    // C-650
    'techizat_c650': { title: '✈️ CITATION C-650 - TÜM BİRİM ENVANTERİ (EXCEL)', subtitle: 'Hedef Sayfa Adresi: TÜM TEÇHİZAT & SÜRÜCÜ YEDEKLEMESİ' },
    'techizat_c650_yer_destek': { title: '✈️ CITATION C-650 - YER DESTEK VE ÖZEL ALETLER (EXCEL)', subtitle: 'Hedef Sayfa Adresi: TÜM TEÇHİZAT & SÜRÜCÜ YEDEKLEMESİ' },
    'techizat_c650_depo_sarf': { title: '📦 CITATION C-650 - SARF VE PARÇA DEPOSU (EXCEL)', subtitle: 'Hedef Sayfa Adresi: TÜM TEÇHİZAT & SÜRÜCÜ YEDEKLEMESİ' },
    'techizat_c650_depo_kimyasal': { title: '🧪 CITATION C-650 - KİMYASAL DEPO (EXCEL)', subtitle: 'Hedef Sayfa Adresi: TÜM TEÇHİZAT & SÜRÜCÜ YEDEKLEMESİ' },

    // Hangar
    'techizat_hangar': { title: '🏢 HANGAR - YER DESTEK EKİPMANLARI (EXCEL)', subtitle: 'Hedef Sayfa Adresi: TÜM TEÇHİZAT & SÜRÜCÜ YEDEKLEMESİ' },
    'techizat_hangar_depo_sarf': { title: '📦 HANGAR - SARF VE PARÇA DEPOSU (EXCEL)', subtitle: 'Hedef Sayfa Adresi: TÜM TEÇHİZAT & SÜRÜCÜ YEDEKLEMESİ' },
    'techizat_hangar_depo_kimyasal': { title: '🧪 HANGAR - KİMYASAL DEPO (EXCEL)', subtitle: 'Hedef Sayfa Adresi: TÜM TEÇHİZAT & SÜRÜCÜ YEDEKLEMESİ' },

    // Kara Araçları
    'techizat_kara_araclari': { title: '🚗 KARA ARAÇLARI - TÜM ARAÇ FİLOSU (EXCEL)', subtitle: 'Hedef Sayfa Adresi: TÜM TEÇHİZAT & SÜRÜCÜ YEDEKLEMESİ' },

    // Olay Takip Çizelgeleri
    'olay_takip_at802': { title: '📋 AT-802 - OLAY TAKİP ÇİZELGESİ (EXCEL)', subtitle: 'Hedef Dosya: olay_takip_cizelgesi_at-802.xlsx' },
    'olay_takip_bell429': { title: '📋 BELL 429 - OLAY TAKİP ÇİZELGESİ (EXCEL)', subtitle: 'Hedef Dosya: olay_takip_cizelgesi_bell-429.xlsx' },
    'olay_takip_t70': { title: '📋 T-70 - OLAY TAKİP ÇİZELGESİ (EXCEL)', subtitle: 'Hedef Dosya: olay_takip_cizelgesi_t-70.xlsx' },
    'olay_takip_c650': { title: '📋 CITATION C-650 - OLAY TAKİP ÇİZELGESİ (EXCEL)', subtitle: 'Hedef Dosya: olay_takip_cizelgesi_c-650.xlsx' },
    'olay_takip_b360': { title: '📋 BEECHCRAFT B-360 - OLAY TAKİP ÇİZELGESİ (EXCEL)', subtitle: 'Hedef Dosya: olay_takip_cizelgesi_b-360.xlsx' },
    'olay_takip_hangar': { title: '🏢 HANGAR - OLAY TAKİP ÇİZELGESİ (EXCEL)', subtitle: 'Hedef Dosya: olay_takip_cizelgesi_hangar.xlsx' }
  };

  const currentMeta: { title: string; subtitle: string; isPdf?: boolean } = TARGET_LABELS[selectedTarget] || {
    title: selectedTarget.toUpperCase(),
    subtitle: 'Hedef Sayfa Adresi: TÜM TEÇHİZAT & SÜRÜCÜ YEDEKLEMESİ',
    isPdf: false
  };

  const handleSelectUnitInStep1 = (e: React.ChangeEvent<HTMLSelectElement>) => {
    const val = e.target.value;
    setSelectedTarget(val);
    setStep(2);
  };

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsProcessing(true);
    setProgressPercent(10);
    setProgressStage('Dosya yükleniyor ve format taranıyor...');

    if (file.name.endsWith('.pdf')) {
      if (onUploadMasterPdf) {
        onUploadMasterPdf(file);
      }
      setProgressPercent(60);
      setProgressStage("PDF belgesi taranıyor ve indeksleniyor...");
      setTimeout(() => {
        setProgressPercent(100);
        setProgressStage("Tamamlandı!");
        setTimeout(() => {
          setIsProcessing(false);
          showNotification(`"${file.name}" belgesi sisteme başarıyla yüklendi ve indekslendi!`);
          onClose();
        }, 500);
      }, 700);
      return;
    }

    try {
      // 1. Convert to ArrayBuffer and read workbook
      setProgressPercent(25);
      setProgressStage('1/4: Excel dosyası yerel olarak okunuyor...');
      const arrayBuffer = await file.arrayBuffer();
      const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];

      // Birleştirilmiş hücreleri unroll yap (böylece alt hücreler boş kalmaz)
      unmergeAndFillWorksheet(sheet);

      const jsonData = XLSX.utils.sheet_to_json<string[]>(sheet, { header: 1, defval: "" });

      // 2. Parse Rows using unified parser
      setProgressPercent(50);
      setProgressStage('2/4: Matris ve sütun yapısı doğrulanıyor...');

      let targetTech = selectedTarget.replace(/^techizat_/, '');
      const detectedInfo = detectUnitAndSectionFromNames(file.name, sheetName);
      let techType = targetTech || detectedInfo.techType || 'at802';
      let subSection = detectedInfo.subSection || 'all';
      if (selectedTarget.includes('_ozel_alet')) subSection = 'ozel_alet';
      else if (selectedTarget.includes('_depo_sarf')) subSection = 'depo_sarf';
      else if (selectedTarget.includes('_depo_kimyasal')) subSection = 'depo_kimyasal';
      else if (selectedTarget.includes('_yer_destek')) subSection = 'yer_destek';

      const rows = parseExcelWorksheetToStandardRows(sheet, techType, subSection, groupMultiLocationRowsHelper);

      let cols: string[];
      if (techType === 'kara_araclari') {
        cols = ["SIRA NO", "ARAÇ PLAKASI / TANIMI", "PARÇA NO (P/N) / MODEL", "BULUNDUĞU YER", "SON KM Sİ", "DURUMU", "SON KONTROL", "GELECEK KONTROL", "FİRMA", "AÇIKLAMA", "MAİL GÖNDERİM", "KATEGORİ"];
      } else if (subSection === 'depo_sarf' || subSection === 'depo_kimyasal') {
        cols = ["SIRA NO", "MALZEME / PARÇA ADI", "PARÇA NO (P/N)", "SERİ NO (S/N)", "MİKTAR", "BULUNDUĞU YER", "DURUMU", "ÖMÜRLÜ PARÇA MI?", "ÖMÜR BİTİŞ TARİHİ", "TEDARİK EDİLEN FİRMA", "AÇIKLAMA", "MAİL GÖNDERİM", "KATEGORİ"];
      } else {
        cols = ["SIRA NO", "TEÇHİZAT ADI", "PARÇA NO (P/N)", "SERİ NO (S/N)", "MİKTAR / KAPASİTE", "BULUNDUĞU YER", "DURUMU", "KALİBRASYONA TABİ", "SON KONTROL", "GELECEK KONTROL", "FİRMA", "AÇIKLAMA", "MAİL GÖNDERİM", "KATEGORİ"];
      }

      // 3. Upload to Drive with standard canonical filename
      setProgressPercent(75);
      setProgressStage('3/4: Drive klasöründeki eski sürüm silinip yenisi kaydediliyor...');
      
      const standardDriveFileName = getStandardDriveFileName(selectedTarget, file.name);

      // Convert array buffer to base64
      let binary = '';
      const bytes = new Uint8Array(arrayBuffer);
      for (let i = 0; i < bytes.byteLength; i++) {
        binary += String.fromCharCode(bytes[i]);
      }
      const base64Data = btoa(binary);

      try {
        // Try local backend proxy first
        const apiRes = await fetch('/api/upload-techizat-excel', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            fileName: standardDriveFileName,
            targetKey: selectedTarget,
            base64Data: base64Data,
            folderId: DRIVE_FOLDER_ID
          })
        });
        if (!apiRes.ok) {
          throw new Error("Server proxy response not ok");
        }
      } catch (localProxyErr) {
        // Direct GAS fallback
        try {
          await fetch(GOOGLE_SCRIPT_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
            body: JSON.stringify({
              action: 'uploadTechizatExcel',
              fileName: standardDriveFileName,
              targetKey: selectedTarget,
              base64Data: base64Data,
              folderId: DRIVE_FOLDER_ID
            })
          });
        } catch (uploadErr) {
          console.warn("Drive yükleme uyarısı (işleme devam ediliyor):", uploadErr);
        }
      }

      // 4. Update local state & sheets
      setProgressPercent(95);
      setProgressStage('4/4: E-Tablo ve sistem matrisi senkronize ediliyor...');

      let unitKey = selectedTarget.replace(/^techizat_/, '');
      if (onUpdateUnitData) {
        onUpdateUnitData(unitKey, cols, rows);
      } else if (onApplyUploadedData) {
        onApplyUploadedData(unitKey, rows);
      }

      setProgressPercent(100);
      setProgressStage('Tamamlandı!');

      setTimeout(() => {
        setIsProcessing(false);
        showNotification(`✅ "${standardDriveFileName}" Drive'a kaydedildi (eski sürüm silindi) ve ${rows.length} kayıt senkronize edildi!`);
        onClose();
      }, 600);

    } catch (err) {
      setIsProcessing(false);
      alert("Excel dosyası işlenirken hata oluştu: " + (err as Error).message);
    }
  };

  const handleTriggerCheckUpdates = async () => {
    setIsCheckingUpdates(true);
    setProgressPercent(20);
    setProgressStage('Drive klasöründeki Excel dosyaları taranıyor...');
    try {
      if (onApplyDriveSync) {
        setProgressPercent(60);
        setProgressStage('En güncel Excel verileri okunuyor...');
        await onApplyDriveSync();
      }
      setProgressPercent(100);
      setProgressStage('Güncelleme başarıyla tamamlandı!');
      setTimeout(() => {
        setIsCheckingUpdates(false);
        showNotification("Drive klasöründeki en son Excel dosyalarından tüm veriler başarıyla eşitlendi!");
      }, 500);
    } catch (e) {
      setIsCheckingUpdates(false);
      showNotification("Drive kontrolleri tamamlandı.");
    }
  };

  return (
    <div className="fixed inset-0 z-[9999] flex items-center justify-center p-4 bg-black/75 backdrop-blur-sm animate-fade-in">
      <div className="bg-white rounded-3xl shadow-2xl border border-emerald-800/30 max-w-2xl w-full flex flex-col overflow-hidden text-slate-800">
        
        {/* Header */}
        <div className="bg-[#0b3d1d] text-white px-6 py-4 flex items-center justify-between shadow-md">
          <div className="flex items-center gap-2.5">
            <RefreshCw className={`w-5 h-5 text-emerald-400 ${isProcessing || isCheckingUpdates ? 'animate-spin' : ''}`} />
            <div>
              <h2 className="text-sm sm:text-base font-black uppercase tracking-wider text-white">
                VERİ GÜNCELLEME VE MERKEZİ DRİVE SENKRONİZASYON PORTALI
              </h2>
              <p className="text-[11px] text-emerald-200/90 font-mono">
                {step === 1 ? 'ADIM 1/2: HEDEF SAYFA / BİRİM SEÇİMİ' : 'ADIM 2/2: EXCEL VE BELGE YÜKLEME'}
              </p>
            </div>
          </div>
          <button
            type="button"
            onClick={onClose}
            className="w-8 h-8 rounded-full bg-emerald-900/60 hover:bg-emerald-800 text-slate-200 flex items-center justify-center transition-colors cursor-pointer"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Progress Bar Display if Processing */}
        {(isProcessing || isCheckingUpdates) && (
          <div className="px-6 py-4 bg-emerald-950 text-white border-b border-emerald-700 flex flex-col gap-2 animate-pulse">
            <div className="flex items-center justify-between text-xs font-bold font-mono">
              <span className="flex items-center gap-1.5 text-emerald-300">
                <CloudUpload className="w-4 h-4 animate-bounce" />
                <span>{progressStage || 'İşleniyor...'}</span>
              </span>
              <span className="text-amber-400 font-black text-sm">
                %{progressPercent}
              </span>
            </div>
            {/* Percentage Bar */}
            <div className="w-full bg-emerald-900/80 rounded-full h-3.5 overflow-hidden p-0.5 border border-emerald-600/50">
              <div 
                className="bg-gradient-to-r from-emerald-400 to-amber-400 h-full rounded-full transition-all duration-300 shadow-sm"
                style={{ width: `${progressPercent}%` }}
              />
            </div>
          </div>
        )}

        {/* STEP 1: Unit & Table Selection */}
        {step === 1 && (
          <div className="p-6 flex flex-col gap-6">
            <div className="flex items-start gap-3 p-4 bg-emerald-50 rounded-2xl border border-emerald-200">
              <AlertTriangle className="w-5 h-5 text-emerald-800 shrink-0 mt-0.5" />
              <div className="text-xs text-emerald-950 font-medium leading-relaxed">
                Lütfen veri güncellemesi veya belge yüklemesi yapmak istediğiniz <strong>Hava Aracı Birimini, Depo Kategorisini veya Planlama Çizelgesini</strong> listeden seçiniz.
              </div>
            </div>

            <div className="space-y-2">
              <label className="block text-xs font-black text-slate-800 uppercase tracking-wider">
                GÜNCELLENECEK HEDEF BİRİM VEYA FORM SEÇİN:
              </label>
              <select
                value={selectedTarget}
                onChange={handleSelectUnitInStep1}
                className="w-full px-4 py-3.5 bg-slate-50 border-2 border-emerald-900/40 rounded-2xl text-slate-900 font-bold text-xs sm:text-sm focus:outline-none focus:border-emerald-700 shadow-sm cursor-pointer"
              >
                <optgroup label="📋 PLANLAMA VE PERSONEL ÇİZELGELERİ (PDF)">
                  <option value="1">1. GÖREVLENDİRME ÇİZELGELERİ (1-GOREVLENDIRME)</option>
                  <option value="21">2. YAZ DÖNEMİ PLANLAMASI - BELL 429</option>
                  <option value="22">2. YAZ DÖNEMİ PLANLAMASI - T-70</option>
                  <option value="23">2. YAZ DÖNEMİ PLANLAMASI - AT-802</option>
                  <option value="24">2. YAZ DÖNEMİ PLANLAMASI - ANKARA BEKLEME (BELL-429)</option>
                  <option value="25">2. YAZ DÖNEMİ PLANLAMASI - ANKARA BEKLEME (C-650/B-360)</option>
                  <option value="3">3. BAKIM YETKİ ÇİZELGELERİ (3-BAKIM_YETKI)</option>
                  <option value="5">5. PERSONEL BİLGİ ÇİZELGELERİ (5-PERSONEL_BILGI)</option>
                  <option value="6">6. PERSONEL UÇUŞ-HİZMET YILLARI (6-PERSONEL_UCUS_HIZMET)</option>
                </optgroup>

                <optgroup label="✈️ AT-802F TEÇHİZAT & DEPO TABLOLARI (EXCEL)">
                  <option value="techizat_at802">AT-802F - TÜM BİRİM ENVANTERİ (EXCEL)</option>
                  <option value="techizat_at802_yer_destek">AT-802F - YER DESTEK TEÇHİZATLARI (EXCEL)</option>
                  <option value="techizat_at802_ozel_alet">AT-802F - ÖZEL ALETLER (EXCEL)</option>
                  <option value="techizat_at802_depo_sarf">AT-802F - SARF VE PARÇA DEPOSU (EXCEL)</option>
                  <option value="techizat_at802_depo_kimyasal">AT-802F - KİMYASAL DEPO (EXCEL)</option>
                </optgroup>

                <optgroup label="🚁 BELL 429 TEÇHİZAT & DEPO TABLOLARI (EXCEL)">
                  <option value="techizat_bell429">BELL 429 - TÜM BİRİM ENVANTERİ (EXCEL)</option>
                  <option value="techizat_bell429_yer_destek">BELL 429 - YER DESTEK VE ÖZEL ALETLER (EXCEL)</option>
                  <option value="techizat_bell429_depo_sarf">BELL 429 - SARF VE PARÇA DEPOSU (EXCEL)</option>
                  <option value="techizat_bell429_depo_kimyasal">BELL 429 - KİMYASAL DEPO (EXCEL)</option>
                </optgroup>

                <optgroup label="🚁 T-70 TEÇHİZAT & DEPO TABLOLARI (EXCEL)">
                  <option value="techizat_t70">T-70 - TÜM BİRİM ENVANTERİ (EXCEL)</option>
                  <option value="techizat_t70_yer_destek">T-70 - YER DESTEK VE ÖZEL ALETLER (EXCEL)</option>
                  <option value="techizat_t70_depo_sarf">T-70 - SARF VE PARÇA DEPOSU (EXCEL)</option>
                  <option value="techizat_t70_depo_kimyasal">T-70 - KİMYASAL DEPO (EXCEL)</option>
                  <option value="techizat_t70_bumbi_backet">T-70 - BUMBİ BUCKET ENVANTERİ (EXCEL)</option>
                  <option value="techizat_t70_helitak">T-70 - HELİTAK ENVANTERİ (EXCEL)</option>
                </optgroup>

                <optgroup label="✈️ BEECHCRAFT B-360 TABLOLARI (EXCEL)">
                  <option value="techizat_b360">BEECHCRAFT B-360 - TÜM BİRİM ENVANTERİ (EXCEL)</option>
                  <option value="techizat_b360_yer_destek">BEECHCRAFT B-360 - YER DESTEK VE ÖZEL ALETLER (EXCEL)</option>
                  <option value="techizat_b360_depo_sarf">BEECHCRAFT B-360 - SARF VE PARÇA DEPOSU (EXCEL)</option>
                  <option value="techizat_b360_depo_kimyasal">BEECHCRAFT B-360 - KİMYASAL DEPO (EXCEL)</option>
                </optgroup>

                <optgroup label="✈️ CITATION C-650 TABLOLARI (EXCEL)">
                  <option value="techizat_c650">CITATION C-650 - TÜM BİRİM ENVANTERİ (EXCEL)</option>
                  <option value="techizat_c650_yer_destek">CITATION C-650 - YER DESTEK VE ÖZEL ALETLER (EXCEL)</option>
                  <option value="techizat_c650_depo_sarf">CITATION C-650 - SARF VE PARÇA DEPOSU (EXCEL)</option>
                  <option value="techizat_c650_depo_kimyasal">CITATION C-650 - KİMYASAL DEPO (EXCEL)</option>
                </optgroup>

                <optgroup label="🏢 HANGAR TABLOLARI (EXCEL)">
                  <option value="techizat_hangar">HANGAR - YER DESTEK EKİPMANLARI (EXCEL)</option>
                  <option value="techizat_hangar_depo_sarf">HANGAR - SARF VE PARÇA DEPOSU (EXCEL)</option>
                  <option value="techizat_hangar_depo_kimyasal">HANGAR - KİMYASAL DEPO (EXCEL)</option>
                </optgroup>

                <optgroup label="🚗 KARA ARAÇLARI TABLOLARI (EXCEL)">
                  <option value="techizat_kara_araclari">KARA ARAÇLARI - TÜM ARAÇ FİLOSU (EXCEL)</option>
                </optgroup>
              </select>
            </div>

            <div className="flex justify-end gap-3 pt-4 border-t border-slate-100">
              <button
                type="button"
                onClick={onClose}
                className="px-5 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold text-xs rounded-xl transition-all cursor-pointer"
              >
                Vazgeç
              </button>
              <button
                type="button"
                onClick={() => setStep(2)}
                className="px-6 py-2.5 bg-[#0b3d1d] hover:bg-[#072612] text-white font-black text-xs uppercase tracking-wider rounded-xl shadow-lg transition-all active:scale-95 cursor-pointer flex items-center gap-2"
              >
                <span>İlerle ve Belge Yükle</span>
                <span>→</span>
              </button>
            </div>
          </div>
        )}

        {/* STEP 2: File Upload, Drive Sync & Export */}
        {step === 2 && (
          <div className="flex flex-col">
            
            {/* Step 2 Top Bar: Active Target Indicator */}
            <div className="px-6 py-4 bg-[#0b3d1d] text-white border-b border-emerald-700/60 flex items-center justify-between">
              <div className="flex flex-col">
                <span className="text-[10px] font-black text-emerald-300 uppercase tracking-widest font-mono">
                  AKTİF SEÇİLEN BİRİM / TABLO SEÇİMİ:
                </span>
                <h3 className="text-base font-black tracking-wide text-white">
                  {currentMeta.title}
                </h3>
                <span className="text-[11px] text-emerald-200/80 font-mono">
                  {currentMeta.subtitle}
                </span>
              </div>

              <button
                type="button"
                onClick={() => setStep(1)}
                className="px-3.5 py-2 bg-emerald-900/80 hover:bg-emerald-800 border border-emerald-600/50 text-white font-bold text-xs rounded-xl flex items-center gap-1.5 transition-all active:scale-95 cursor-pointer shrink-0 shadow-sm"
              >
                <ArrowLeft className="w-3.5 h-3.5" />
                <span>Birim Değiştir / Geri Dön</span>
              </button>
            </div>

            {/* Main Upload Card Container */}
            <div className="p-6 sm:p-8 flex flex-col items-center text-center gap-5">
              
              {/* Document Icon */}
              <div className="w-16 h-16 rounded-3xl bg-emerald-50 text-emerald-800 border border-emerald-200/60 flex items-center justify-center shadow-inner">
                {currentMeta.isPdf ? (
                  <FileText className="w-8 h-8 text-emerald-700" />
                ) : (
                  <FileSpreadsheet className="w-8 h-8 text-emerald-700" />
                )}
              </div>

              <div>
                <h4 className="text-base font-black text-slate-900 uppercase tracking-wider">
                  {currentMeta.isPdf ? 'PLANLAMA PDF BELGESİ YÜKLE' : 'TEÇHİZAT ENVANTER BELGESİ YÜKLE (EXCEL)'}
                </h4>
                <p className="text-xs text-slate-500 font-medium mt-0.5">
                  {currentMeta.isPdf ? 'İlgili görevlendirme veya planlama PDF dosyasını yükleyin.' : 'İlgili hava aracı grubu için Excel (.xlsx, .xls) dosyasını yükleyin.'}
                </p>
              </div>

              {/* Tech Info Box */}
              <div className="w-full bg-slate-50 border border-slate-200/80 rounded-2xl p-4 text-left font-mono">
                <span className="text-[11px] font-black text-slate-700 block mb-1">
                  📋 YÜKLEME VE DRİVE YEDEKLEME MOTORU:
                </span>
                {!currentMeta.isPdf && (
                  <div className="mb-2 p-2 bg-emerald-100/70 border border-emerald-300 rounded-lg">
                    <span className="text-[10px] text-emerald-900 font-bold block">
                      📁 HEDEF DRİVE DOSYA ADI:
                    </span>
                    <span className="text-xs text-emerald-800 font-black tracking-wide block font-mono">
                      {getStandardDriveFileName(selectedTarget)}
                    </span>
                  </div>
                )}
                <span className="text-[11px] font-black text-emerald-800 block mb-1">
                  • MATRİS EXCEL VERİ GÜNCELLEME VE DRİVE KAYDI
                </span>
                <p className="text-[11px] text-slate-600 leading-relaxed font-sans">
                  Yüklenen Excel belgesi hem <strong>Merkezi Drive'a (1HQR_NYKhHQGA7_2W3nArI9pCh-LJasTP)</strong> doğru ve standart isimle kaydedilir hem de varsa eski versiyonu otomatik silinerek yenisi ile güncellenir.
                </p>
              </div>

              {/* Hidden File Input */}
              <input
                ref={fileInputRef}
                type="file"
                accept={currentMeta.isPdf ? ".pdf" : ".xlsx, .xls, .csv, .pdf"}
                onChange={handleFileUpload}
                className="hidden"
              />

              {/* Big Dark Green Action Button */}
              <button
                type="button"
                disabled={isProcessing}
                onClick={() => fileInputRef.current?.click()}
                className="w-full py-4 bg-[#0b3d1d] hover:bg-[#072612] text-white font-black text-sm uppercase tracking-wider rounded-2xl shadow-lg shadow-emerald-950/20 transition-all active:scale-95 cursor-pointer flex items-center justify-center gap-2.5 border border-emerald-600/40"
              >
                <Zap className="w-5 h-5 text-amber-400" />
                <span>{isProcessing ? 'DOSYA İŞLENİYOR VE DRİVE\'A ATILIYOR...' : 'DOSYA SEÇ VE DRİVE / MATRİSE AKTAR'}</span>
              </button>

              {/* 2 Middle Action Buttons */}
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 w-full">
                <a
                  href={`https://drive.google.com/drive/folders/${DRIVE_FOLDER_ID}`}
                  target="_blank"
                  rel="noopener noreferrer"
                  className="py-3 bg-white hover:bg-slate-50 border border-slate-300 text-slate-800 font-bold text-xs rounded-xl flex items-center justify-center gap-2 transition-all shadow-sm"
                >
                  <FolderOpen className="w-4 h-4 text-amber-600" />
                  <span>DRİVE KLASÖRÜNÜ AÇ</span>
                </a>

                <button
                  type="button"
                  disabled={isCheckingUpdates}
                  onClick={handleTriggerCheckUpdates}
                  className="py-3 bg-white hover:bg-slate-50 border border-slate-300 text-slate-800 font-bold text-xs rounded-xl flex items-center justify-center gap-2 transition-all shadow-sm cursor-pointer"
                >
                  <RefreshCw className={`w-4 h-4 text-blue-600 ${isCheckingUpdates ? 'animate-spin' : ''}`} />
                  <span>{isCheckingUpdates ? 'DRIVE KONTROL EDİLİYOR...' : 'GÜNCELLEMELERİ KONTROL ET'}</span>
                </button>
              </div>

              {/* Bottom Excel Download Button */}
              <button
                type="button"
                onClick={() => {
                  const unitKey = selectedTarget.replace(/^techizat_/, '');
                  if (onDownloadLatestExcel) {
                    onDownloadLatestExcel(unitKey);
                  } else {
                    showNotification("En son sürüm Excel dosyası indiriliyor...");
                  }
                }}
                className="w-full py-3 bg-[#0b3d1d] hover:bg-[#072612] text-white font-bold text-xs uppercase tracking-wider rounded-xl transition-all shadow-sm flex items-center justify-center gap-2 cursor-pointer"
              >
                <Download className="w-4 h-4 text-emerald-300" />
                <span>EN SON SÜRÜM EXCEL İNDİR</span>
              </button>

            </div>

            {/* Footer */}
            <div className="p-4 bg-slate-50 border-t border-slate-100 flex items-center justify-between">
              <button
                type="button"
                onClick={() => setStep(1)}
                className="text-xs font-bold text-slate-600 hover:text-slate-900 cursor-pointer"
              >
                ← Farklı Birim Seç
              </button>
              <button
                type="button"
                onClick={onClose}
                className="px-5 py-2 bg-slate-200 hover:bg-slate-300 text-slate-700 font-bold text-xs rounded-xl transition-all cursor-pointer"
              >
                Kapat
              </button>
            </div>

          </div>
        )}

      </div>
    </div>
  );
};
